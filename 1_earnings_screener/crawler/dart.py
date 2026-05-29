"""
================================================
  DART Open API 검증 크롤러
  - 최신 사업보고서/분기보고서에서 연결 실적 확인
  - FnGuide/네이버 수치의 신뢰도 교차 검증용
================================================

[API 키 설정]
  .env 파일에 DART_API_KEY=<발급받은키> 추가
  발급: https://opendart.fss.or.kr
"""

import re
from typing import Optional

import requests

from config import DART_API_KEY, HEADERS, TIMEOUT

DART_BASE = "https://opendart.fss.or.kr/api"


def _get_corp_code(stock_code: str) -> Optional[str]:
    """종목코드 → DART 고유번호(corp_code) 변환."""
    if not DART_API_KEY:
        return None
    url = f"{DART_BASE}/company.json"
    try:
        res = requests.get(url, params={
            "crtfc_key": DART_API_KEY,
            "stock_code": stock_code,
        }, timeout=TIMEOUT)
        data = res.json()
        if data.get("status") == "000":
            return data.get("corp_code")
    except Exception:
        pass
    return None


def _get_latest_report(corp_code: str) -> Optional[dict]:
    """가장 최근 사업보고서(11011) 또는 분기보고서(11012) 조회."""
    if not DART_API_KEY:
        return None
    url = f"{DART_BASE}/list.json"
    try:
        res = requests.get(url, params={
            "crtfc_key": DART_API_KEY,
            "corp_code":  corp_code,
            "bgn_de":     "20250101",
            "pblntf_ty":  "A",        # 정기공시
            "last_reprt_at": "Y",
            "page_count": "10",
        }, timeout=TIMEOUT)
        items = res.json().get("list", [])
        for item in items:
            if item.get("report_nm", "") in (
                "사업보고서", "반기보고서", "분기보고서"
            ):
                return item
    except Exception:
        pass
    return None


def _fetch_financial(corp_code: str, bsns_year: str) -> Optional[dict]:
    """
    연결재무제표 주요 항목 조회.
    반환: {"revenue": float, "op_profit": float}  단위: 억원
    """
    if not DART_API_KEY:
        return None
    url = f"{DART_BASE}/fnlttSinglAcntAll.json"
    try:
        res = requests.get(url, params={
            "crtfc_key": DART_API_KEY,
            "corp_code":  corp_code,
            "bsns_year":  bsns_year,
            "reprt_code": "11011",  # 사업보고서
            "fs_div":     "CFS",    # 연결재무제표
        }, timeout=TIMEOUT)
        items = res.json().get("list", [])

        revenue = op_profit = None
        for item in items:
            account = item.get("account_nm", "")
            amount_str = item.get("thstrm_amount", "0").replace(",", "")
            try:
                amount = float(amount_str) / 1e8  # 원 → 억
            except ValueError:
                continue

            if "매출" in account and revenue is None:
                revenue = amount
            elif "영업이익" in account and op_profit is None:
                op_profit = amount

        if revenue or op_profit:
            return {"revenue": revenue, "op_profit": op_profit}
    except Exception:
        pass
    return None


def verify(stock_code: str, expected_revenue_25a: float, expected_op_profit_25a: float,
           tolerance: float = 0.15) -> dict:
    """
    DART 실적과 FnGuide 수치를 교차 검증.
    tolerance: 허용 오차 비율 (기본 15%)

    반환: {
        "verified": bool,
        "dart_revenue": float | None,
        "dart_op_profit": float | None,
        "revenue_diff_pct": float | None,
        "op_profit_diff_pct": float | None,
    }
    """
    result = {
        "verified": True,  # DART 키 없으면 기본 통과
        "dart_revenue": None,
        "dart_op_profit": None,
        "revenue_diff_pct": None,
        "op_profit_diff_pct": None,
    }

    if not DART_API_KEY:
        return result

    corp_code = _get_corp_code(stock_code)
    if not corp_code:
        return result

    financials = _fetch_financial(corp_code, str(2025))
    if not financials:
        return result

    result["dart_revenue"]    = financials.get("revenue")
    result["dart_op_profit"]  = financials.get("op_profit")

    def pct_diff(dart_val, fn_val):
        if dart_val and fn_val and fn_val != 0:
            return abs(dart_val - fn_val) / abs(fn_val) * 100
        return None

    rev_diff = pct_diff(result["dart_revenue"],   expected_revenue_25a)
    op_diff  = pct_diff(result["dart_op_profit"], expected_op_profit_25a)

    result["revenue_diff_pct"]   = rev_diff
    result["op_profit_diff_pct"] = op_diff

    # 오차가 허용 범위를 초과하면 검증 실패
    if rev_diff is not None and rev_diff > tolerance * 100:
        result["verified"] = False
    if op_diff is not None and op_diff > tolerance * 100:
        result["verified"] = False

    return result


def fetch_risk(stock_code: str) -> dict:
    """
    감사의견 및 부채비율 조회.
    반환: {"audit_ok": bool|None, "debt_ratio": float|None}
    """
    result = {"audit_ok": None, "debt_ratio": None}
    if not DART_API_KEY:
        return result

    corp_code = _get_corp_code(stock_code)
    if not corp_code:
        return result

    # 감사의견
    try:
        res = requests.get(f"{DART_BASE}/fnlttAuditOpinion.json", params={
            "crtfc_key": DART_API_KEY,
            "corp_code": corp_code,
            "bsns_year": "2025",
            "reprt_code": "11011",
        }, timeout=TIMEOUT)
        for op in res.json().get("list", []):
            opinion = op.get("op_dp_nm", "")
            result["audit_ok"] = (opinion == "적정")
            break
    except Exception:
        pass

    # 부채비율 (부채총계 / 자본총계 * 100)
    try:
        res = requests.get(f"{DART_BASE}/fnlttSinglAcntAll.json", params={
            "crtfc_key": DART_API_KEY,
            "corp_code": corp_code,
            "bsns_year": "2025",
            "reprt_code": "11011",
            "fs_div": "CFS",
        }, timeout=TIMEOUT)
        liabilities = equity = None
        for item in res.json().get("list", []):
            acct = item.get("account_nm", "")
            try:
                amount = float(str(item.get("thstrm_amount", "0")).replace(",", "") or "0")
            except ValueError:
                continue
            if "부채총계" in acct or "부채합계" in acct:
                liabilities = amount
            elif ("자본총계" in acct or "자본합계" in acct) and equity is None:
                equity = amount
        if liabilities is not None and equity and equity != 0:
            result["debt_ratio"] = round(liabilities / equity * 100, 1)
    except Exception:
        pass

    return result
