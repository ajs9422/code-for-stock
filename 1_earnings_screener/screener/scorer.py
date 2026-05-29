"""
================================================
  점수화 모듈
  필터를 통과한 종목을 투자 매력도 기준으로 정렬.
================================================

[점수 항목 및 가중치]
  매출 증가율      (20%)
  영업이익 증가율  (35%)
  OPM 개선폭       (25%)
  컨센 상향률      (20%)
"""

from dataclasses import dataclass
from typing import Optional

from crawler.fnguide import ConsensusData


@dataclass
class Score:
    code: str
    name: str
    total: float = 0.0
    revenue_score:    float = 0.0
    op_profit_score:  float = 0.0
    opm_score:        float = 0.0
    revision_score:   float = 0.0


def _normalize(value: Optional[float], low: float, high: float) -> float:
    """value를 [low, high] 범위 기준으로 0~100 점수로 정규화."""
    if value is None:
        return 0.0
    if high == low:
        return 50.0
    return min(100.0, max(0.0, (value - low) / (high - low) * 100))


def score(data: ConsensusData) -> Score:
    s = Score(code=data.code, name=data.name)

    # 매출 증가율: 20~100% 구간 정규화
    s.revenue_score   = _normalize(data.revenue_growth,  20,  100)

    # 영업이익 증가율: 100~500% (흑자전환은 300점 처리)
    op_growth = 300.0 if data.turnaround else data.op_profit_growth
    s.op_profit_score = _normalize(op_growth, 100, 500)

    # OPM 개선폭: 5~20%p 구간 정규화
    s.opm_score       = _normalize(data.opm_improvement,  5,   20)

    # 컨센 상향률: 0~20% 구간 정규화
    s.revision_score  = _normalize(data.op_revision_1m,   0,   20)

    s.total = (
        s.revenue_score   * 0.20
        + s.op_profit_score * 0.35
        + s.opm_score       * 0.25
        + s.revision_score  * 0.20
    )
    return s


def rank(candidates: list[ConsensusData]) -> list[tuple[ConsensusData, Score]]:
    """점수 계산 후 내림차순 정렬."""
    scored = [(data, score(data)) for data in candidates]
    scored.sort(key=lambda x: x[1].total, reverse=True)
    return scored
