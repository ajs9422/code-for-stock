"""
DART API 기반 공시 Catalyst 탐지
- 최근 60일 주요 공시 조회 및 필터링
- 계약 금액 파싱 + 전년 매출 대비 비중 계산
- 일일 2만건 API 제한 → SQLite 캐싱
"""

import re
import sqlite3
from contextlib import contextmanager
from dataclasses import dataclass
from datetime import date, timedelta
from typing import Optional

import requests

from config import DART_API_KEY, DB_PATH, TIMEOUT

DART_BASE = "https://opendart.fss.or.kr/api"

# 주요 공시 유형 키워드 (report_nm 매칭)
_KEY_DISCLOSURES = [
    "단일판매·공급계약체결",
    "단일판매공급계약",
    "타법인주식및출자증권취득결정",
    "타법인주식및출자증권처분결정",
    "유형자산양수결정",
    "유형자산양도결정",
    "주요경영사항",
    "영업(잠정)실적",
    "잠정실적",
    "주요사항보고서",
]


@dataclass
class DartDisclosure:
    rcept_no:    str
    report_nm:   str
    rcept_dt:    str          # YYYYMMDD
    corp_name:   str
    amount:      Optional[float] = None   # 계약/거래 금액 (억원)
    revenue_pct: Optional[float] = None   # 전년 매출 대비 비중 (%)
    is_major:    bool = False             # 매출의 10% 이상


# ── DB 캐싱 ──────────────────────────────────────────────────────────────────

@contextmanager
def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    finally:
        con.close()


def _init_table():
    with _conn() as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS dart_disclosure (
            rcept_no   TEXT PRIMARY KEY,
            code       TEXT NOT NULL,
            report_nm  TEXT,
            rcept_dt   TEXT,
            corp_name  TEXT,
            fetched_at TEXT DEFAULT (date('now'))
        )
        """)


def _save_disclosures(code: str, items: list[dict]):
    if not items:
        return
    with _conn() as con:
        con.executemany("""
        INSERT OR IGNORE INTO dart_disclosure
            (rcept_no, code, report_nm, rcept_dt, corp_name)
        VALUES
            (:rcept_no, :code, :report_nm, :rcept_dt, :corp_name)
        """, [{"code": code, **item} for item in items])


def _load_cached(code: str, since: str) -> list[dict]:
    _init_table()
    with _conn() as con:
        rows = con.execute("""
        SELECT rcept_no, report_nm, rcept_dt, corp_name
        FROM   dart_disclosure
        WHERE  code = ? AND rcept_dt >= ? AND fetched_at = date('now')
        ORDER  BY rcept_dt DESC
        """, (code, since)).fetchall()
    return [dict(r) for r in rows]


# ── DART API 호출 ─────────────────────────────────────────────────────────────

def _get_corp_code(stock_code: str) -> Optional[str]:
    if not DART_API_KEY:
        return None
    try:
        res = requests.get(f"{DART_BASE}/company.json", params={
            "crtfc_key": DART_API_KEY, "stock_code": stock_code,
        }, timeout=TIMEOUT)
        data = res.json()
        if data.get("status") == "000":
            return data.get("corp_code")
    except Exception:
        pass
    return None


def _fetch_disclosure_list(corp_code: str, days: int = 60) -> list[dict]:
    """DART 공시 목록 조회 (전체 유형)."""
    if not DART_API_KEY:
        return []
    start = (date.today() - timedelta(days=days)).strftime("%Y%m%d")
    try:
        res = requests.get(f"{DART_BASE}/list.json", params={
            "crtfc_key":     DART_API_KEY,
            "corp_code":     corp_code,
            "bgn_de":        start,
            "page_count":    "100",
        }, timeout=TIMEOUT)
        return res.json().get("list", [])
    except Exception:
        return []


def _parse_amount(text: str) -> Optional[float]:
    """공시 본문/제목에서 금액(억원) 추출."""
    # 패턴: "1,234억", "12,345백만", "123,456,789원"
    patterns = [
        (r"([\d,]+)\s*억",    1.0),       # 억원 단위
        (r"([\d,]+)\s*백만",  0.01),      # 백만원 → 억원
        (r"([\d,]+)\s*천만",  0.1),       # 천만원 → 억원
        (r"([\d,]+)\s*만",    0.0001),    # 만원 → 억원
        (r"([\d,]+)\s*원",    1e-8),      # 원 단위
    ]
    for pattern, multiplier in patterns:
        m = re.search(pattern, text.replace(" ", ""))
        if m:
            try:
                amount = float(m.group(1).replace(",", "")) * multiplier
                if amount > 0:
                    return round(amount, 1)
            except ValueError:
                pass
    return None


def _is_key_disclosure(report_nm: str) -> bool:
    return any(kw in report_nm for kw in _KEY_DISCLOSURES)


# ── 공개 인터페이스 ───────────────────────────────────────────────────────────

def fetch_disclosures(
    stock_code: str,
    days: int = 60,
    revenue_25a: Optional[float] = None,  # 전년 매출 (억원), 비중 계산용
) -> list[DartDisclosure]:
    """
    주요 공시 목록 조회 + Catalyst 분석.
    DART API 키 없거나 오류 시 빈 리스트 반환.
    """
    _init_table()
    since = (date.today() - timedelta(days=days)).strftime("%Y%m%d")

    # 오늘 캐시 확인
    cached = _load_cached(stock_code, since)
    if cached:
        raw_list = cached
    else:
        corp_code = _get_corp_code(stock_code)
        if not corp_code:
            return []
        raw_list = _fetch_disclosure_list(corp_code, days)
        _save_disclosures(stock_code, [
            {
                "rcept_no": r.get("rcept_no", ""),
                "report_nm": r.get("report_nm", ""),
                "rcept_dt": r.get("rcept_dt", ""),
                "corp_name": r.get("corp_name", ""),
            }
            for r in raw_list
        ])

    results: list[DartDisclosure] = []
    for r in raw_list:
        report_nm = r.get("report_nm", "") if isinstance(r, dict) else r["report_nm"]
        if not _is_key_disclosure(report_nm):
            continue

        disc = DartDisclosure(
            rcept_no  = r.get("rcept_no", "") if isinstance(r, dict) else r["rcept_no"],
            report_nm = report_nm,
            rcept_dt  = r.get("rcept_dt", "") if isinstance(r, dict) else r["rcept_dt"],
            corp_name = r.get("corp_name", "") if isinstance(r, dict) else r["corp_name"],
        )

        # 제목에서 금액 추출
        amount = _parse_amount(report_nm)
        if amount:
            disc.amount = amount
            if revenue_25a and revenue_25a > 0:
                disc.revenue_pct = round(amount / revenue_25a * 100, 1)
                disc.is_major    = disc.revenue_pct >= 10.0

        results.append(disc)

    return results
