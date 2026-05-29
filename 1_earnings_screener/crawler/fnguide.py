"""
================================================
  FnGuide 크롤러
  - 연간 컨센서스 (매출·영업이익·OPM) 수집
  - 커버리지 증권사 수 확인
  - 컨센서스 월별 변화 추적용 데이터 수집
================================================

[수집 엔드포인트]
  SVD_Consensus.asp  : 연간 컨센서스 테이블 (25A / 26E)
  SVD_Finance.asp    : 실제 재무제표 (연결 기준 확인용)
  JSON /01_06/03_    : 목표주가 컨센서스 히스토리
"""

import random
import re
import time
import warnings
from dataclasses import dataclass, field
from typing import Optional

import pandas as pd
import requests
from bs4 import BeautifulSoup

warnings.filterwarnings("ignore")

from config import (
    ACTUAL_YEAR, ESTIMATE_YEAR,
    DELAY_MIN, DELAY_MAX,
    HEADERS, MIN_COVERAGE, TIMEOUT,
)

BASE_URL = "https://comp.fnguide.com"


# ── 데이터 모델 ────────────────────────────────────────────────────────────────

@dataclass
class ConsensusData:
    code: str
    name: str
    market_cap: float = 0.0          # 시가총액 (억원)

    # 매출액 (억원)
    revenue_25a: Optional[float] = None
    revenue_26e: Optional[float] = None

    # 영업이익 (억원)
    op_profit_25a: Optional[float] = None
    op_profit_26e: Optional[float] = None

    # 영업이익률 (%)
    opm_25a: Optional[float] = None
    opm_26e: Optional[float] = None

    # 컨센서스 메타
    analyst_count: int = 0

    # 1개월/3개월 전 영업이익 26E (DB에서 채워짐, 변화율 계산용)
    op_profit_26e_1m_ago: Optional[float] = None
    op_profit_26e_3m_ago: Optional[float] = None

    # 파생 지표 (screener에서 계산)
    revenue_growth: Optional[float] = None      # 매출 증가율 (%)
    op_profit_growth: Optional[float] = None    # 영업이익 증가율 (%)
    opm_improvement: Optional[float] = None     # OPM 개선폭 (% point)
    op_revision_1m: Optional[float] = None      # 최근 1개월 컨센 상향률 (%)
    op_revision_3m: Optional[float] = None      # 최근 3개월 컨센 상향률 (%)

    # 흑자전환 여부
    turnaround: bool = False

    # 밸류에이션 (네이버 금융에서 수집)
    fwd_per: Optional[float] = None
    peg: Optional[float] = None
    price: Optional[float] = None
    price_52w_high: Optional[float] = None
    price_52w_low: Optional[float] = None
    price_52w_pct: Optional[float] = None       # 현재가/52주고점 (%)
    price_surge_pct: Optional[float] = None     # 현재가/52주저점 - 1 (급등률 %)

    # 리스크 (DART + KRX)
    is_admin: bool = False
    debt_ratio: Optional[float] = None          # 부채비율 (%)
    audit_ok: Optional[bool] = None             # 감사의견 적정 여부

    # 수급 (KRX)
    foreign_net_3m: Optional[float] = None      # 외국인 3개월 순매수 (억원)
    institution_net_3m: Optional[float] = None  # 기관 3개월 순매수 (억원)

    def is_valid(self) -> bool:
        """수집된 핵심 데이터가 존재하는지 확인"""
        return (
            self.revenue_25a is not None
            and self.op_profit_25a is not None
            and self.revenue_26e is not None
            and self.op_profit_26e is not None
        )


# ── 유틸 ──────────────────────────────────────────────────────────────────────

def _delay():
    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))


def _get_soup(url: str, params: dict = None) -> Optional[BeautifulSoup]:
    try:
        res = requests.get(url, headers=HEADERS, params=params, timeout=TIMEOUT)
        res.raise_for_status()
        res.encoding = "utf-8"
        return BeautifulSoup(res.text, "lxml")
    except Exception as e:
        return None


def _parse_num(text: str) -> Optional[float]:
    if not text:
        return None
    text = text.strip().replace(",", "").replace(" ", "")
    if text in ("", "-", "N/A", "NA", "—"):
        return None
    try:
        return float(re.sub(r"[^\d.\-]", "", text))
    except ValueError:
        return None


def _safe_growth(new: float, old: float) -> Optional[float]:
    """증가율 계산. old가 0이거나 None이면 None 반환."""
    if old is None or new is None:
        return None
    if old == 0:
        return None
    return (new - old) / abs(old) * 100


# ── 유니버스 ──────────────────────────────────────────────────────────────────

def get_universe() -> pd.DataFrame:
    """
    KOSPI + KOSDAQ 전종목 수집 후 시총 필터 적용.
    Returns: DataFrame[code, name, market_cap(억)]
    """
    try:
        import FinanceDataReader as fdr
        kospi  = fdr.StockListing("KOSPI")[["Code", "Name", "Marcap"]].dropna()
        kosdaq = fdr.StockListing("KOSDAQ")[["Code", "Name", "Marcap"]].dropna()
        df = pd.concat([kospi, kosdaq], ignore_index=True)
        df.columns = ["code", "name", "market_cap"]
        df["market_cap"] = pd.to_numeric(df["market_cap"], errors="coerce").fillna(0) / 1e8
        df = df[df["market_cap"] >= 2000].copy()
        df["code"] = df["code"].astype(str).str.zfill(6)
        return df.reset_index(drop=True)
    except Exception as e:
        print(f"[유니버스] 수집 실패: {e}")
        return pd.DataFrame(columns=["code", "name", "market_cap"])


# ── FnGuide 컨센서스 파싱 ─────────────────────────────────────────────────────

def _find_year_col(headers: list[str], year: int, is_actual: bool) -> Optional[int]:
    """
    헤더 텍스트 목록에서 특정 연도의 컬럼 인덱스를 찾는다.
    FnGuide는 '2025/12A', '2026/12E', '2025(E)', '2026E' 등 다양한 형식을 사용.
    """
    year_str = str(year)
    # 실적(A) vs 추정(E) 마커
    actual_markers   = ["A", "실적", "a", "잠정"]
    estimate_markers = ["E", "추정", "e", "P", "F", "예상"]
    target_markers   = actual_markers if is_actual else estimate_markers

    for i, h in enumerate(headers):
        if year_str not in h:
            continue
        # 마커가 명시적으로 있으면 우선 매칭
        for m in target_markers:
            if m in h:
                return i
        # 마커 없이 연도만 있는 경우: 실적은 앞쪽, 추정은 뒤쪽 컬럼
        # (순서 기반 fallback은 아래 fallback 로직에서 처리)
    return None


def _parse_consensus_table(soup: BeautifulSoup, data: ConsensusData) -> bool:
    """
    SVD_Consensus.asp 페이지의 연간 컨센서스 테이블을 파싱.
    매출액, 영업이익, 영업이익률 행을 찾아 data에 채운다.
    Returns True if at least revenue or op_profit was found.
    """
    actual_year   = str(ACTUAL_YEAR)    # "2025"
    estimate_year = str(ESTIMATE_YEAR)  # "2026"

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue

        # 헤더에서 연도 컬럼 위치 파악
        header_cells = rows[0].find_all(["th", "td"])
        headers = [c.get_text(strip=True) for c in header_cells]

        col_25a = _find_year_col(headers, ACTUAL_YEAR,   is_actual=True)
        col_26e = _find_year_col(headers, ESTIMATE_YEAR, is_actual=False)

        # 마커 없는 경우 fallback: 연도가 포함된 순서로 처음/두번째
        if col_25a is None or col_26e is None:
            year_cols = [i for i, h in enumerate(headers) if actual_year in h or estimate_year in h]
            actual_cols   = [i for i in year_cols if actual_year   in headers[i]]
            estimate_cols = [i for i in year_cols if estimate_year in headers[i]]
            if col_25a is None and actual_cols:
                col_25a = actual_cols[-1]   # 가장 오른쪽 25 컬럼
            if col_26e is None and estimate_cols:
                col_26e = estimate_cols[0]  # 가장 왼쪽 26 컬럼

        if col_25a is None and col_26e is None:
            continue

        found_any = False
        for row in rows[1:]:
            cells = row.find_all(["th", "td"])
            if not cells:
                continue
            label = cells[0].get_text(strip=True)

            def val(col):
                if col is not None and col < len(cells):
                    return _parse_num(cells[col].get_text(strip=True))
                return None

            if re.search(r"매출(액)?", label) and "증가" not in label:
                if data.revenue_25a is None:
                    data.revenue_25a = val(col_25a)
                if data.revenue_26e is None:
                    data.revenue_26e = val(col_26e)
                found_any = True

            elif re.search(r"영업이익", label) and "률" not in label and "마진" not in label:
                if data.op_profit_25a is None:
                    data.op_profit_25a = val(col_25a)
                if data.op_profit_26e is None:
                    data.op_profit_26e = val(col_26e)
                found_any = True

            elif re.search(r"영업이익률|OPM|영업마진", label):
                if data.opm_25a is None:
                    data.opm_25a = val(col_25a)
                if data.opm_26e is None:
                    data.opm_26e = val(col_26e)

        if found_any:
            return True

    return False


def _parse_analyst_count(soup: BeautifulSoup) -> int:
    """페이지 전체 텍스트에서 커버리지 증권사 수 추출."""
    text = soup.get_text()

    # 패턴 1: "N개 기관", "추정기관수 N"
    for pat in [
        r'(\d+)\s*개\s*기관',
        r'추정기관수[^\d]*(\d+)',
        r'(\d+)\s*개\s*증권',
        r'컨센서스\s*\((\d+)',
    ]:
        m = re.search(pat, text)
        if m:
            return int(m.group(1))

    # 패턴 2: 증권사별 목표가 테이블에서 행 수 카운트
    count = 0
    for tag in soup.select("table"):
        for row in tag.find_all("tr"):
            cells = row.find_all("td")
            if cells and re.search(r"증권|투자|리서치|자산운용", cells[0].get_text()):
                count += 1
    return count


def _parse_consensus_revision(code: str) -> Optional[dict]:
    """
    FnGuide JSON API에서 목표주가 컨센서스 히스토리 수집.
    월별 평균 목표주가 변화를 통해 이익 컨센 변화 방향을 간접 확인.

    반환: {
        "avg_target": float,       # 현재 목표주가 평균
        "avg_target_prev": float,  # 직전 목표주가 평균
    }
    """
    url = f"{BASE_URL}/SVO2/json/data/01_06/03_A{code}.json"
    try:
        res = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        res.raise_for_status()
        body = res.json()

        avg     = None
        avg_pre = None
        for item in body if isinstance(body, list) else []:
            if isinstance(item, dict):
                avg     = avg     or _parse_num(str(item.get("AVG_PRC", "")))
                avg_pre = avg_pre or _parse_num(str(item.get("AVG_PRC_BF", "")))

        return {"avg_target": avg, "avg_target_prev": avg_pre}
    except Exception:
        return None


# ── 공개 인터페이스 ───────────────────────────────────────────────────────────

def fetch_consensus(code: str, name: str, market_cap: float = 0.0) -> Optional[ConsensusData]:
    """
    단일 종목의 FnGuide 컨센서스 데이터를 수집하여 ConsensusData로 반환.
    실패 또는 핵심 데이터 없으면 None 반환.
    """
    params = {
        "pGB":       "1",
        "gicode":    f"A{code}",
        "cID":       "",
        "MenuYn":    "Y",
        "ReportGB":  "",
        "NewMenuID": "160",
        "stkGb":     "701",
    }
    soup = _get_soup(f"{BASE_URL}/SVO2/asp/SVD_Consensus.asp", params)
    if not soup:
        return None

    data = ConsensusData(code=code, name=name, market_cap=market_cap)

    # 연간 컨센서스 테이블 파싱
    if not _parse_consensus_table(soup, data):
        return None

    # OPM 직접 계산 (테이블에 없을 경우)
    if data.opm_25a is None and data.revenue_25a and data.op_profit_25a and data.revenue_25a != 0:
        data.opm_25a = data.op_profit_25a / data.revenue_25a * 100
    if data.opm_26e is None and data.revenue_26e and data.op_profit_26e and data.revenue_26e != 0:
        data.opm_26e = data.op_profit_26e / data.revenue_26e * 100

    # 커버리지 증권사 수
    data.analyst_count = _parse_analyst_count(soup)

    # 컨센 변화 히스토리 (JSON API)
    revision = _parse_consensus_revision(code)
    # 목표주가 방향이 상향이면 이익 컨센도 상향 가능성 높음 (보조 지표)
    data._target_revision = revision  # storage에서 활용

    _delay()
    return data if data.is_valid() else None


def fetch_all(universe: pd.DataFrame, progress_callback=None) -> list[ConsensusData]:
    """
    유니버스 전종목 컨센서스 수집.
    progress_callback(current, total, name): 진도 표시용 콜백 (선택)
    """
    results = []
    total = len(universe)

    for i, row in enumerate(universe.itertuples(), 1):
        if progress_callback:
            progress_callback(i, total, row.name)

        data = fetch_consensus(row.code, row.name, row.market_cap)
        if data:
            results.append(data)

    return results
