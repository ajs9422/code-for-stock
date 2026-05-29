"""
Catalyst 이벤트 점수화 및 타임라인 생성

점수 공식:
    score = 영향도(1~5) × 시기_가중치(0.5~1.0) × 확실성_가중치(0.5~1.0)

시기_가중치: 최근일수록 높음
    0~7일   → 1.0
    8~14일  → 0.85
    15~30일 → 0.70
    31~60일 → 0.55
    60일 초과 → 0.40

확실성_가중치:
    DART 공시 확정  → 1.0
    뉴스 + 키워드 3개↑ → 0.80
    뉴스 + 키워드 1~2개 → 0.65
    뉴스 단순 언급  → 0.50
"""

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from typing import Optional

from catalyst.dart_catalyst import DartDisclosure
from catalyst.keyword_matcher import MatchResult, CATEGORY_IMPACT


@dataclass
class CatalystEvent:
    source:       str             # "DART" | "뉴스"
    category:     str
    title:        str
    event_date:   date
    impact:       float           # 영향도 1~5
    timing_w:     float           # 시기 가중치
    certainty_w:  float           # 확실성 가중치
    score:        float = 0.0
    is_major:     bool  = False   # 매출의 10% 이상 (DART)
    amount:       Optional[float] = None  # 계약금액 (억원)
    revenue_pct:  Optional[float] = None  # 매출 대비 비중


@dataclass
class CatalystResult:
    code:                str
    catalyst_total:      float = 0.0
    events:              list[CatalystEvent] = field(default_factory=list)
    top_categories:      list[str]           = field(default_factory=list)
    summary:             str  = ""


def _timing_weight(event_date: date) -> float:
    days_ago = (date.today() - event_date).days
    if days_ago <= 7:   return 1.00
    if days_ago <= 14:  return 0.85
    if days_ago <= 30:  return 0.70
    if days_ago <= 60:  return 0.55
    return 0.40


def _certainty_weight(source: str, n_keywords: int) -> float:
    if source == "DART":
        return 1.00
    if n_keywords >= 3:
        return 0.80
    if n_keywords >= 1:
        return 0.65
    return 0.50


def _parse_date(date_str: str) -> Optional[date]:
    for fmt in ("%Y.%m.%d", "%Y%m%d", "%Y-%m-%d"):
        try:
            return datetime.strptime(date_str.split()[0], fmt).date()
        except ValueError:
            pass
    return None


def score_from_news(
    news_items: list[dict],
    match_results: dict[str, list[MatchResult]],
) -> list[CatalystEvent]:
    """뉴스 키워드 매칭 결과를 CatalystEvent 목록으로 변환."""
    events: list[CatalystEvent] = []

    for category, hits in match_results.items():
        for hit in hits:
            news_date = _parse_date(getattr(hit, "_news_date", ""))
            if not news_date:
                continue
            timing_w   = _timing_weight(news_date)
            certainty_w = _certainty_weight("뉴스", len(hit.matched_keywords))
            sc = hit.impact * timing_w * certainty_w

            events.append(CatalystEvent(
                source      = "뉴스",
                category    = category,
                title       = getattr(hit, "_title", ""),
                event_date  = news_date,
                impact      = hit.impact,
                timing_w    = timing_w,
                certainty_w = certainty_w,
                score       = round(sc, 3),
            ))

    return events


def score_from_dart(disclosures: list[DartDisclosure]) -> list[CatalystEvent]:
    """DART 공시를 CatalystEvent 목록으로 변환."""
    events: list[CatalystEvent] = []

    for d in disclosures:
        event_date = _parse_date(d.rcept_dt)
        if not event_date:
            continue

        # 공시 유형 → 카테고리 매핑
        if "계약" in d.report_nm or "수주" in d.report_nm:
            category = "수주계약"
        elif "취득" in d.report_nm or "처분" in d.report_nm or "주식" in d.report_nm:
            category = "M&A"
        elif "실적" in d.report_nm or "잠정" in d.report_nm:
            category = "실적이벤트"
        elif "양수" in d.report_nm or "양도" in d.report_nm:
            category = "M&A"
        else:
            category = "수주계약"

        impact      = CATEGORY_IMPACT.get(category, 3.0)
        timing_w    = _timing_weight(event_date)
        certainty_w = 1.0  # DART = 확정
        sc = impact * timing_w * certainty_w

        # 매출 10% 이상 공시는 가산점
        if d.is_major:
            sc *= 1.3

        events.append(CatalystEvent(
            source      = "DART",
            category    = category,
            title       = d.report_nm,
            event_date  = event_date,
            impact      = impact,
            timing_w    = timing_w,
            certainty_w = certainty_w,
            score       = round(sc, 3),
            is_major    = d.is_major,
            amount      = d.amount,
            revenue_pct = d.revenue_pct,
        ))

    return events


def aggregate(
    code: str,
    news_events: list[CatalystEvent],
    dart_events: list[CatalystEvent],
) -> CatalystResult:
    """뉴스 + DART 이벤트를 통합하여 종목 Catalyst 점수 산출."""
    all_events = sorted(
        news_events + dart_events,
        key=lambda e: e.score,
        reverse=True,
    )

    # 카테고리별 최고 점수만 합산 (중복 방지)
    best: dict[str, float] = {}
    for ev in all_events:
        if ev.score > best.get(ev.category, 0):
            best[ev.category] = ev.score

    total = round(sum(best.values()), 2)

    top_cats = sorted(best, key=best.get, reverse=True)[:3]

    # 요약 문장 생성
    parts = []
    for ev in all_events[:3]:
        if ev.source == "DART" and ev.is_major:
            parts.append(f"{ev.category}(DART 🔥 매출대비 {ev.revenue_pct:.1f}%)")
        elif ev.source == "DART":
            parts.append(f"{ev.category}(공시)")
        else:
            parts.append(ev.category)
    summary = " · ".join(dict.fromkeys(parts))  # 중복 제거

    return CatalystResult(
        code            = code,
        catalyst_total  = total,
        events          = all_events[:20],   # 상위 20개만
        top_categories  = top_cats,
        summary         = summary or "이벤트 없음",
    )


def build_timeline(events: list[CatalystEvent], months: int = 6) -> list[dict]:
    """향후 N개월 예정 이벤트 타임라인 (현재는 과거 이벤트 기반)."""
    cutoff = date.today() - timedelta(days=months * 30)
    recent = [
        {
            "date":     e.event_date.isoformat(),
            "source":   e.source,
            "category": e.category,
            "title":    e.title[:60],
            "score":    e.score,
            "is_major": e.is_major,
        }
        for e in events
        if e.event_date >= cutoff
    ]
    return sorted(recent, key=lambda x: x["date"], reverse=True)
