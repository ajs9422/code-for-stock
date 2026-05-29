"""
뉴스/공시 텍스트에서 Catalyst 카테고리 키워드 매칭
- 7개 카테고리 키워드 딕셔너리
- 매칭 결과: 카테고리 + 신뢰도(매칭 키워드 수)
"""

from dataclasses import dataclass, field

KEYWORDS: dict[str, list[str]] = {
    "제품출시": [
        "출시", "런칭", "신제품", "신작", "발매", "론칭",
        "오픈", "출품", "서비스 시작", "공개", "베타", "출시 예정",
    ],
    "수주계약": [
        "수주", "계약 체결", "공급계약", "납품", "MOU", "업무협약",
        "협력 계약", "파트너십", "공급", "수주 잔고", "장기계약",
        "단일판매", "공급 계약",
    ],
    "M&A": [
        "인수", "합병", "M&A", "지분 취득", "매각", "피인수",
        "합작법인", "JV", "자회사", "출자", "경영권", "인수합병",
        "타법인 주식",
    ],
    "실적이벤트": [
        "잠정실적", "실적 발표", "영업이익", "어닝", "실적 서프라이즈",
        "컨센서스 상회", "흑자전환", "매출 성장", "분기 실적",
        "연간 실적", "호실적", "어닝 서프라이즈", "깜짝 실적",
    ],
    "정책수혜": [
        "정부 지원", "보조금", "정책 수혜", "규제 완화", "인허가",
        "정부 사업", "세제 혜택", "국책 과제", "정책", "법안 통과",
        "승인", "허가", "선정",
    ],
    "글로벌진출": [
        "해외 진출", "수출", "미국 시장", "유럽 시장", "중국 진출",
        "글로벌 파트너", "현지법인", "해외 법인", "글로벌 출시",
        "북미", "유럽", "일본", "동남아", "해외",
    ],
    "신기술": [
        "AI", "인공지능", "특허 취득", "기술 개발", "R&D",
        "신기술", "연구개발", "혁신", "딥러닝", "반도체",
        "바이오", "신약", "임상", "플랫폼", "알고리즘",
    ],
}

# 카테고리별 기본 영향도 점수 (1~5)
CATEGORY_IMPACT: dict[str, float] = {
    "수주계약":   5.0,
    "M&A":        4.5,
    "실적이벤트": 4.0,
    "제품출시":   3.5,
    "글로벌진출": 3.0,
    "정책수혜":   2.5,
    "신기술":     2.0,
}


@dataclass
class MatchResult:
    category:       str
    matched_keywords: list[str] = field(default_factory=list)
    confidence:     float = 0.0   # 0~1 (매칭 키워드 수 / 전체 키워드 수)
    impact:         float = 0.0   # CATEGORY_IMPACT 기반


def match(text: str) -> list[MatchResult]:
    """
    텍스트에서 모든 카테고리 키워드 매칭.
    반환: 매칭된 카테고리 목록 (신뢰도 내림차순)
    """
    text_lower = text.lower()
    results: list[MatchResult] = []

    for category, keywords in KEYWORDS.items():
        matched = [kw for kw in keywords if kw.lower() in text_lower]
        if not matched:
            continue
        confidence = min(1.0, len(matched) / max(3, len(keywords) * 0.3))
        results.append(MatchResult(
            category=category,
            matched_keywords=matched,
            confidence=round(confidence, 3),
            impact=CATEGORY_IMPACT.get(category, 2.0),
        ))

    results.sort(key=lambda r: (r.confidence, r.impact), reverse=True)
    return results


def match_news_list(news_items: list[dict]) -> dict[str, list[MatchResult]]:
    """
    뉴스 목록 전체에 키워드 매칭 적용.
    반환: {category: [MatchResult, ...]}  (카테고리별 집계)
    """
    category_hits: dict[str, list[MatchResult]] = {cat: [] for cat in KEYWORDS}

    for item in news_items:
        text = item.get("title", "") + " " + item.get("summary", "")
        for result in match(text):
            result._news_date = item.get("news_date", "")
            result._title     = item.get("title", "")
            category_hits[result.category].append(result)

    # 빈 카테고리 제거
    return {cat: hits for cat, hits in category_hits.items() if hits}
