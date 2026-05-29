"""
================================================
  네이버 금융 폴백 크롤러
  FnGuide에서 데이터를 못 가져온 종목에 대해
  네이버 금융 종목 페이지에서 컨센서스 보완.
================================================
"""

import re
import time
import random
from typing import Optional

import requests
from bs4 import BeautifulSoup

from config import HEADERS, TIMEOUT, DELAY_MIN, DELAY_MAX, ACTUAL_YEAR, ESTIMATE_YEAR
from crawler.fnguide import ConsensusData, _parse_num, _delay

BASE_URL = "https://finance.naver.com"


def _get_soup(url: str) -> Optional[BeautifulSoup]:
    try:
        res = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        res.raise_for_status()
        res.encoding = "euc-kr"
        return BeautifulSoup(res.text, "lxml")
    except Exception:
        return None


def fetch_consensus_naver(code: str, name: str, market_cap: float = 0.0) -> Optional[ConsensusData]:
    """
    네이버 금융 종목 메인 페이지에서 연간 실적/추정 테이블 파싱.
    FnGuide 폴백용이므로 가능한 필드만 채운다.
    """
    url = f"{BASE_URL}/item/main.naver?code={code}"
    soup = _get_soup(url)
    if not soup:
        return None

    data = ConsensusData(code=code, name=name, market_cap=market_cap)
    actual_year   = str(ACTUAL_YEAR)
    estimate_year = str(ESTIMATE_YEAR)

    # 네이버 금융 "연간 실적" 테이블: id="content" > table.tb_type1
    for table in soup.select("table.tb_type1, table.type2"):
        rows = table.find_all("tr")
        if not rows:
            continue

        header = [th.get_text(strip=True) for th in rows[0].find_all(["th", "td"])]

        col_25a = next((i for i, h in enumerate(header) if actual_year   in h), None)
        col_26e = next((i for i, h in enumerate(header) if estimate_year in h), None)

        if col_25a is None or col_26e is None:
            continue

        for row in rows[1:]:
            cells = row.find_all(["th", "td"])
            if not cells:
                continue
            label = cells[0].get_text(strip=True)

            def val(col):
                if col is not None and col < len(cells):
                    return _parse_num(cells[col].get_text(strip=True))
                return None

            if "매출" in label and "증가" not in label:
                data.revenue_25a = data.revenue_25a or val(col_25a)
                data.revenue_26e = data.revenue_26e or val(col_26e)
            elif "영업이익" in label and "률" not in label:
                data.op_profit_25a = data.op_profit_25a or val(col_25a)
                data.op_profit_26e = data.op_profit_26e or val(col_26e)

    # OPM 계산
    if data.opm_25a is None and data.revenue_25a and data.op_profit_25a and data.revenue_25a != 0:
        data.opm_25a = data.op_profit_25a / data.revenue_25a * 100
    if data.opm_26e is None and data.revenue_26e and data.op_profit_26e and data.revenue_26e != 0:
        data.opm_26e = data.op_profit_26e / data.revenue_26e * 100

    # 목표주가에서 애널리스트 수 간접 추정
    target_em = soup.select_one("em._target_price")
    if target_em:
        data.analyst_count = data.analyst_count or 1   # 최소 1명

    _delay()
    return data if data.is_valid() else None


def fetch_valuation_naver(code: str) -> dict:
    """네이버 금융에서 현재주가, PER, 52주 고점 수집."""
    url = f"{BASE_URL}/item/main.naver?code={code}"
    soup = _get_soup(url)
    result = {
        "price": None, "fwd_per": None,
        "price_52w_high": None, "price_52w_low": None,
        "price_52w_pct": None, "price_surge_pct": None, "peg": None,
    }
    if not soup:
        return result

    # 현재주가
    for sel in ["strong#_nowVal", "p.no_today strong"]:
        el = soup.select_one(sel)
        if el:
            result["price"] = _parse_num(el.get_text(strip=True))
            break

    # 52주 최고 / 최저
    for th in soup.find_all("th"):
        label = th.get_text(strip=True)
        if "52주 최고" not in label and "52주 최저" not in label:
            continue
        td = th.find_next_sibling("td")
        if not td:
            continue
        text = td.get_text(strip=True)
        parts = text.split()
        if not parts:
            continue
        raw = _parse_num(parts[0])
        if "52주 최고" in label:
            result["price_52w_high"] = raw
        elif "52주 최저" in label:
            result["price_52w_low"] = raw

    # PER
    per_td = soup.select_one("td.td_per")
    if per_td:
        result["fwd_per"] = _parse_num(per_td.get_text(strip=True).replace("배", ""))

    # 52주 고점 대비 비율 (고점에 얼마나 가까운지)
    if result["price"] and result["price_52w_high"] and result["price_52w_high"] > 0:
        result["price_52w_pct"] = round(result["price"] / result["price_52w_high"] * 100, 1)

    # 52주 저점 대비 급등률 (저점에서 얼마나 올랐는지)
    if result["price"] and result["price_52w_low"] and result["price_52w_low"] > 0:
        result["price_surge_pct"] = round(
            (result["price"] - result["price_52w_low"]) / result["price_52w_low"] * 100, 1
        )

    _delay()
    return result
