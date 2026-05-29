"""
================================================
  스크리닝 필터 (AND 조건)
================================================

[필터 조건 — 모두 충족해야 통과]
  1) 매출 26E/25A 증가율 ≥ 10%
  2) 영업이익 26E/25A 증가율 ≥ 30%  또는  흑자전환 (25A<0, 26E>0)
  3) 영업이익률(OPM) 26E - 25A ≥ 2%p
  4) 최근 1개월 26E 영업이익 컨센서스 상향  (DB 스냅샷 비교)

[유니버스 전제조건 — 아래 미충족 종목은 필터 전 제외]
  - 시가총액 2,000억 이상
  - 커버리지 증권사 2곳 이상
"""

from dataclasses import dataclass
from typing import Optional

from config import (
    MIN_COVERAGE,
    MIN_MARKET_CAP,
    MIN_OPM_IMPROVEMENT,
    MIN_OP_PROFIT_GROWTH,
    MIN_REVENUE_GROWTH,
)
from crawler.fnguide import ConsensusData


@dataclass
class FilterResult:
    passed: bool
    reason: str = ""          # 탈락 이유 (디버그용)
    # 각 조건 통과 여부
    c1_revenue:    bool = False
    c2_op_profit:  bool = False
    c3_opm:        bool = False
    c4_revision:   bool = False


def _compute_metrics(data: ConsensusData) -> ConsensusData:
    """ConsensusData의 파생 지표(증가율·OPM 개선폭·컨센 변화율)를 채운다."""

    # 매출 증가율
    if data.revenue_25a is not None and data.revenue_26e is not None and data.revenue_25a != 0:
        data.revenue_growth = (data.revenue_26e - data.revenue_25a) / abs(data.revenue_25a) * 100

    # 영업이익 증가율 + 흑자전환 판정
    if data.op_profit_25a is not None and data.op_profit_26e is not None:
        if data.op_profit_25a < 0 and data.op_profit_26e > 0:
            data.turnaround = True
            data.op_profit_growth = None   # 흑자전환은 증가율로 표현 불가
        elif data.op_profit_25a != 0:
            data.op_profit_growth = (
                (data.op_profit_26e - data.op_profit_25a) / abs(data.op_profit_25a) * 100
            )

    # OPM 개선폭
    if data.opm_25a is not None and data.opm_26e is not None:
        data.opm_improvement = data.opm_26e - data.opm_25a

    # 1개월 컨센서스 상향률
    if data.op_profit_26e is not None and data.op_profit_26e_1m_ago is not None:
        if data.op_profit_26e_1m_ago != 0:
            data.op_revision_1m = (
                (data.op_profit_26e - data.op_profit_26e_1m_ago)
                / abs(data.op_profit_26e_1m_ago) * 100
            )

    return data


def apply(data: ConsensusData) -> FilterResult:
    """
    단일 종목에 AND 필터를 적용한다.
    유니버스 전제조건 → 4개 조건 순서로 검사.
    """
    # ── 전제조건: 유니버스 ────────────────────────────────
    if data.market_cap < MIN_MARKET_CAP:
        return FilterResult(False, f"시총 {data.market_cap:.0f}억 < {MIN_MARKET_CAP}억")

    if data.analyst_count < MIN_COVERAGE:
        return FilterResult(False, f"커버리지 {data.analyst_count}개 < {MIN_COVERAGE}개")

    # ── 파생 지표 계산 ────────────────────────────────────
    _compute_metrics(data)

    result = FilterResult(passed=False)

    # 조건 1: 매출 증가율 ≥ 20%
    result.c1_revenue = (
        data.revenue_growth is not None
        and data.revenue_growth >= MIN_REVENUE_GROWTH
    )

    # 조건 2: 영업이익 증가율 ≥ 100% 또는 흑자전환
    result.c2_op_profit = data.turnaround or (
        data.op_profit_growth is not None
        and data.op_profit_growth >= MIN_OP_PROFIT_GROWTH
    )

    # 조건 3: OPM 개선폭 ≥ 5%p
    result.c3_opm = (
        data.opm_improvement is not None
        and data.opm_improvement >= MIN_OPM_IMPROVEMENT
    )

    # 조건 4: 최근 1개월 컨센 상향 (DB 스냅샷이 없으면 조건 통과로 처리)
    if data.op_profit_26e_1m_ago is None:
        result.c4_revision = True   # 비교 데이터 없음 → 통과(보수적으로는 False도 가능)
    else:
        result.c4_revision = (
            data.op_revision_1m is not None
            and data.op_revision_1m > 0
        )

    # 탈락 이유 기록
    if not result.c1_revenue:
        rev_str = f"{data.revenue_growth:.1f}" if data.revenue_growth is not None else "N/A"
        result.reason += f"매출증가율={rev_str}%(<{MIN_REVENUE_GROWTH}%) "
    if not result.c2_op_profit:
        result.reason += f"영업이익증가율={data.op_profit_growth}%(<{MIN_OP_PROFIT_GROWTH}%) "
    if not result.c3_opm:
        result.reason += f"OPM개선={data.opm_improvement}%p(<{MIN_OPM_IMPROVEMENT}%p) "
    if not result.c4_revision:
        rev1m_str = f"{data.op_revision_1m:.1f}" if data.op_revision_1m is not None else "N/A"
        result.reason += f"컨센하향(1m={rev1m_str}%) "

    result.passed = (
        result.c1_revenue
        and result.c2_op_profit
        and result.c3_opm
        and result.c4_revision
    )
    return result


def run(candidates: list[ConsensusData]) -> list[tuple[ConsensusData, FilterResult]]:
    """전체 후보 종목에 필터 적용. 통과 종목만 반환."""
    passed = []
    for data in candidates:
        fr = apply(data)
        if fr.passed:
            passed.append((data, fr))
    return passed
