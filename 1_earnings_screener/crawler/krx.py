"""
KRX 데이터 크롤러
- 관리종목 코드 목록
- 외국인/기관 투자자 수급 (최근 3개월 순매수)
"""

import requests
from datetime import date, timedelta
from typing import Optional

from config import TIMEOUT

_URL = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Referer":    "http://data.krx.co.kr",
    "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",
}


def fetch_admin_stocks() -> set:
    """KRX 관리종목 코드 목록 (6자리) 반환. 실패 시 빈 set."""
    try:
        res = requests.post(_URL, data={
            "bld":           "dbms/MDC/STAT/standard/MDCSTAT04302",
            "locale":        "ko_KR",
            "mktId":         "ALL",
            "csvxls_isNo":   "false",
        }, headers=_HEADERS, timeout=TIMEOUT)
        items = res.json().get("OutBlock_1", [])
        return {str(item.get("ISU_SRT_CD", "")).zfill(6)
                for item in items if item.get("ISU_SRT_CD")}
    except Exception:
        return set()


def _get_isu_code(code: str) -> Optional[str]:
    """6자리 종목코드 → KRX ISU 전체코드 (예: KR7005930003)."""
    try:
        res = requests.post(_URL, data={
            "bld":        "dbms/comm/finder/finder_stkisu",
            "locale":     "ko_KR",
            "searchText": code,
            "mktsel":     "ALL",
            "typeNo":     "0",
        }, headers=_HEADERS, timeout=TIMEOUT)
        for item in res.json().get("block1", []):
            if str(item.get("short_code", "")).zfill(6) == code:
                return item.get("full_code")
    except Exception:
        pass
    return None


def fetch_investor_flow(code: str, days: int = 90) -> dict:
    """
    최근 N일 외국인/기관 순매수 합계 (단위: 억원).
    반환: {"foreign_net_3m": float|None, "institution_net_3m": float|None}
    """
    result = {"foreign_net_3m": None, "institution_net_3m": None}
    isu = _get_isu_code(code)
    if not isu:
        return result

    end   = date.today()
    start = end - timedelta(days=days)

    try:
        res = requests.post(_URL, data={
            "bld":         "dbms/MDC/STAT/standard/MDCSTAT02303",
            "locale":      "ko_KR",
            "isuCd":       isu,
            "strtDd":      start.strftime("%Y%m%d"),
            "endDd":       end.strftime("%Y%m%d"),
            "money":       "1",
            "csvxls_isNo": "false",
        }, headers=_HEADERS, timeout=TIMEOUT)
        rows = res.json().get("output", [])

        def _net_sum(key: str) -> float:
            total = 0.0
            for r in rows:
                try:
                    total += float(str(r.get(key, "0")).replace(",", "") or "0")
                except (ValueError, TypeError):
                    pass
            return round(total / 1e8, 1)  # 원 → 억원

        if rows:
            result["foreign_net_3m"]     = _net_sum("FRGN_NETBID_TRDVAL")
            result["institution_net_3m"] = _net_sum("ORGN_NETBID_TRDVAL")
    except Exception:
        pass

    return result
