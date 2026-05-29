"""
5개 축 가중 점수화

가중치:
    이익 모멘텀    0.35
    컨센서스 모멘텀 0.20
    밸류에이션     0.20
    리스크         0.15
    수급           0.10

각 축 점수는 0~100 정규화 후 가중 합산.
"""

import pandas as pd

WEIGHTS = {
    "earnings":  0.35,
    "consensus": 0.20,
    "valuation": 0.20,
    "risk":      0.15,
    "flow":      0.10,
}


def _norm(s: pd.Series, lo: float, hi: float) -> pd.Series:
    """값을 [lo, hi] 구간 기준 0~100으로 정규화."""
    if hi == lo:
        return pd.Series(50.0, index=s.index)
    return ((s.clip(lo, hi) - lo) / (hi - lo) * 100).fillna(0)


def score(df: pd.DataFrame, top_n: int = None) -> pd.DataFrame:
    """
    5축 점수 계산 후 내림차순 정렬.
    top_n 지정 시 상위 N개만 반환.
    """
    df = df.copy()

    # 1축: 이익 모멘텀 (매출 30% + 영익 50% + OPM 20%)
    op_growth = df["op_profit_growth"].where(
        ~df["turnaround"].fillna(False), other=300.0
    ).fillna(0)
    df["score_earnings"] = (
        _norm(df["revenue_growth"].fillna(0),  10,  80) * 0.30
        + _norm(op_growth,                     30, 300) * 0.50
        + _norm(df["opm_improvement"].fillna(0), 2,  20) * 0.20
    ).clip(0, 100)

    # 2축: 컨센서스 모멘텀 (1M 60% + 3M 40%)
    df["score_consensus"] = (
        _norm(df["op_revision_1m"].fillna(0),  0, 20) * 0.60
        + _norm(df["op_revision_3m"].fillna(0), 0, 30) * 0.40
    ).clip(0, 100)

    # 3축: 밸류에이션 (낮을수록 고점수, PER 50% + PEG 30% + 52주위치 20%)
    per_score = _norm(30 - df["fwd_per"].fillna(15).clip(0, 30), 0, 30)
    peg_score = _norm(2  - df["peg"].fillna(1).clip(0, 2),       0,  2)
    w52_score = _norm(100 - df["price_52w_pct"].fillna(70).clip(0, 100), 0, 100)
    df["score_valuation"] = (
        per_score * 0.50 + peg_score * 0.30 + w52_score * 0.20
    ).clip(0, 100)

    # 4축: 리스크 (기본 100점에서 감점)
    admin_pen = df["is_admin"].fillna(False).astype(float) * 100
    audit_pen = (~df["audit_ok"].fillna(True)).astype(float) * 50
    debt_bonus = _norm(300 - df["debt_ratio"].fillna(150).clip(0, 300), 0, 300) * 0.20
    df["score_risk"] = (100 - admin_pen - audit_pen + debt_bonus).clip(0, 100)

    # 5축: 수급 (외국인 50% + 기관 50%)
    df["score_flow"] = (
        _norm(df["foreign_net_3m"].fillna(0),     -1e4, 1e4) * 0.50
        + _norm(df["institution_net_3m"].fillna(0), -1e4, 1e4) * 0.50
    ).clip(0, 100)

    # 총점
    df["total_score"] = (
        df["score_earnings"]  * WEIGHTS["earnings"]
        + df["score_consensus"] * WEIGHTS["consensus"]
        + df["score_valuation"] * WEIGHTS["valuation"]
        + df["score_risk"]      * WEIGHTS["risk"]
        + df["score_flow"]      * WEIGHTS["flow"]
    ).round(2)

    df = df.sort_values("total_score", ascending=False).reset_index(drop=True)
    if top_n:
        df = df.head(top_n)
    return df
