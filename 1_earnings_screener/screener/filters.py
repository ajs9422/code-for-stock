"""
5개 축 스크리닝 필터 (DataFrame 기반)

각 필터는 pd.DataFrame을 받아 통과 종목만 반환.
데이터가 없는 필드는 보수적으로 통과 처리.

사용 순서 (권장):
    df = earnings_momentum_filter(df)   # 빠른 사전 필터 (API 없음)
    df = consensus_momentum_filter(df)  # DB 데이터 기반
    df = valuation_filter(df)           # 네이버 수집 후
    df = risk_exclusion_filter(df)      # DART + KRX 수집 후
    df = flow_filter(df)                # KRX 수집 후
"""

import pandas as pd
from crawler.fnguide import ConsensusData
from config import (
    MIN_REVENUE_GROWTH,
    MIN_OP_PROFIT_GROWTH,
    MIN_OPM_IMPROVEMENT,
    MAX_FWD_PER,
    MAX_PRICE_52W_PCT,
    MAX_SURGE_FROM_LOW,
    MAX_DEBT_RATIO,
)


def to_dataframe(items: list[ConsensusData]) -> pd.DataFrame:
    """list[ConsensusData] → pd.DataFrame 변환."""
    records = []
    for d in items:
        records.append({
            "code": d.code, "name": d.name, "market_cap": d.market_cap,
            "analyst_count": d.analyst_count,
            "revenue_25a": d.revenue_25a, "revenue_26e": d.revenue_26e,
            "revenue_growth": d.revenue_growth,
            "op_profit_25a": d.op_profit_25a, "op_profit_26e": d.op_profit_26e,
            "op_profit_growth": d.op_profit_growth, "turnaround": d.turnaround,
            "opm_25a": d.opm_25a, "opm_26e": d.opm_26e,
            "opm_improvement": d.opm_improvement,
            "op_revision_1m": d.op_revision_1m,
            "op_revision_3m": d.op_revision_3m,
            "fwd_per": d.fwd_per, "peg": d.peg,
            "price": d.price, "price_52w_high": d.price_52w_high,
            "price_52w_pct": d.price_52w_pct,
            "price_surge_pct": d.price_surge_pct,
            "is_admin": d.is_admin,
            "debt_ratio": d.debt_ratio, "audit_ok": d.audit_ok,
            "foreign_net_3m": d.foreign_net_3m,
            "institution_net_3m": d.institution_net_3m,
        })
    return pd.DataFrame(records)


def earnings_momentum_filter(df: pd.DataFrame) -> pd.DataFrame:
    """
    1축: 이익 모멘텀
    - 매출 YoY >= MIN_REVENUE_GROWTH
    - 영업이익 YoY >= MIN_OP_PROFIT_GROWTH 또는 흑자전환
    - OPM 개선 >= MIN_OPM_IMPROVEMENT
    """
    if df.empty:
        return df
    c1 = df["revenue_growth"].fillna(-999) >= MIN_REVENUE_GROWTH
    c2 = (df["turnaround"].fillna(False)
          | (df["op_profit_growth"].fillna(-999) >= MIN_OP_PROFIT_GROWTH))
    c3 = df["opm_improvement"].fillna(-999) >= MIN_OPM_IMPROVEMENT
    return df[c1 & c2 & c3].copy()


def consensus_momentum_filter(df: pd.DataFrame) -> pd.DataFrame:
    """
    2축: 컨센서스 모멘텀
    - 1M 컨센 상향 > 0% (데이터 없으면 통과)
    - 3M 컨센 상향 > 0% (데이터 없으면 통과)
    """
    if df.empty:
        return df
    has_1m = df["op_revision_1m"].notna()
    c1 = ~has_1m | (df["op_revision_1m"].fillna(-999) > 0)
    has_3m = df["op_revision_3m"].notna()
    c2 = ~has_3m | (df["op_revision_3m"].fillna(-999) > 0)
    return df[c1 & c2].copy()


def valuation_filter(df: pd.DataFrame) -> pd.DataFrame:
    """
    3축: 밸류에이션 + 과열/급등 제외
    - 12M Fwd PER < MAX_FWD_PER (데이터 없으면 통과)
    - 52주 고점 대비 현재가 < MAX_PRICE_52W_PCT — 고점 과열 종목 제외 (데이터 없으면 통과)
    - 52주 저점 대비 급등률 < MAX_SURGE_FROM_LOW — 단기 급등 종목 제외 (데이터 없으면 통과)
    """
    if df.empty:
        return df
    has_per = df["fwd_per"].notna() & (df["fwd_per"] > 0)
    c1 = ~has_per | (df["fwd_per"].fillna(0) < MAX_FWD_PER)

    has_52w = df["price_52w_pct"].notna()
    c2 = ~has_52w | (df["price_52w_pct"].fillna(0) < MAX_PRICE_52W_PCT)

    has_surge = df["price_surge_pct"].notna()
    c3 = ~has_surge | (df["price_surge_pct"].fillna(0) < MAX_SURGE_FROM_LOW)

    return df[c1 & c2 & c3].copy()


def risk_exclusion_filter(df: pd.DataFrame) -> pd.DataFrame:
    """
    4축: 리스크 제거
    - 관리종목 제외
    - 부채비율 < MAX_DEBT_RATIO (데이터 없으면 통과)
    - 감사의견 비적정 제외 (데이터 없으면 통과)
    """
    if df.empty:
        return df
    c1 = ~df["is_admin"].fillna(False)

    has_debt = df["debt_ratio"].notna()
    c2 = ~has_debt | (df["debt_ratio"].fillna(0) < MAX_DEBT_RATIO)

    has_audit = df["audit_ok"].notna()
    c3 = ~has_audit | (df["audit_ok"].fillna(True) == True)

    return df[c1 & c2 & c3].copy()


def flow_filter(df: pd.DataFrame) -> pd.DataFrame:
    """
    5축: 수급
    - 외국인 또는 기관 최근 3M 순매수 > 0 (데이터 없으면 통과)
    """
    if df.empty:
        return df
    has_flow = df["foreign_net_3m"].notna() | df["institution_net_3m"].notna()
    flow_pos = (
        (df["foreign_net_3m"].fillna(0) > 0)
        | (df["institution_net_3m"].fillna(0) > 0)
    )
    return df[~has_flow | flow_pos].copy()
