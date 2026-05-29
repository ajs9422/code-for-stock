"""
일별 스크리닝 리포트 생성기
- 종목별: 기본정보, 실적 테이블, 컨센 변화 추이, Catalyst 요약
- 전일 대비 신규 편입 / 이탈 하이라이트
"""

from datetime import date
from pathlib import Path
from typing import Optional

import pandas as pd

import storage.db as db
from config import ACTUAL_YEAR, ESTIMATE_YEAR, REPORT_DIR


def _fmt(v: Optional[float], suffix: str = "", decimals: int = 1) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return "N/A"
    return f"{v:,.{decimals}f}{suffix}"


def _arrow(v: Optional[float]) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return "▲" if v > 0 else ("▼" if v < 0 else "─")


def _catalyst(row: pd.Series) -> str:
    """점수 기여도 기반 Catalyst 요약 문장 생성."""
    items = []
    if row.get("turnaround"):
        items.append("흑자전환 기대")
    elif row.get("op_profit_growth", 0) and row["op_profit_growth"] >= 200:
        items.append(f"영업이익 {row['op_profit_growth']:+.0f}% 급증")
    elif row.get("op_profit_growth", 0):
        items.append(f"영업이익 {row['op_profit_growth']:+.0f}% 증가")

    if row.get("op_revision_1m", 0) and row["op_revision_1m"] > 5:
        items.append(f"컨센서스 1M {row['op_revision_1m']:+.1f}% 상향")
    if row.get("op_revision_3m", 0) and row["op_revision_3m"] > 10:
        items.append(f"컨센서스 3M {row['op_revision_3m']:+.1f}% 상향")

    score_e = row.get("score_earnings", 0) or 0
    score_f = row.get("score_flow", 0) or 0
    if score_f > 60:
        items.append("외국인/기관 수급 강세")
    if score_e > 70 and not items:
        items.append("실적 고성장 모멘텀")

    return " · ".join(items) if items else "복합 모멘텀"


def _consensus_trend(code: str) -> str:
    """DB 히스토리에서 영업이익 컨센 추이 텍스트 생성."""
    history = db.load_consensus_history(code, days=120)
    if not history:
        return ""
    # 월별 마지막 값만 추출 (최대 4개 포인트)
    monthly: dict[str, float] = {}
    for row in history:
        month = row["snap_date"][:7]
        if row["op_profit_26e"] is not None:
            monthly[month] = row["op_profit_26e"]
    points = list(monthly.items())[-4:]
    if len(points) < 2:
        return ""
    parts = [f"{m}: {v:,.0f}억" for m, v in points]
    return " → ".join(parts)


def generate(df: pd.DataFrame, run_date: date = None, top_n: int = 10,
             catalyst_map: dict = None) -> Path:
    """
    5축 스크리닝 결과로 일별 Markdown 리포트 생성.
    반환: 저장된 파일 경로
    """
    today = run_date or date.today()
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    out_path = REPORT_DIR / f"{today}_v2.md"

    prev_codes      = db.load_prev_results(today)
    prev_code_names = db.load_prev_result_names(today)
    today_codes = set(df["code"].tolist()) if not df.empty else set()
    new_entries = today_codes - prev_codes
    exits       = prev_codes - today_codes

    lines = []

    # ── 헤더 ────────────────────────────────────────────────────────────────────
    lines += [
        f"# 이익 추정치 급등 종목 스크리닝 — {today}",
        "",
        f"**통과 종목: {len(df)}개** (Top {top_n} 표시)",
        "",
    ]

    # ── 신규 편입 / 이탈 ────────────────────────────────────────────────────────
    if new_entries or exits:
        lines += ["## 전일 대비 변동", ""]
        if new_entries:
            new_names = df[df["code"].isin(new_entries)]["name"].tolist()
            lines.append(f"🆕 **신규 편입**: {', '.join(new_names)}")
        if exits:
            exit_names = [prev_code_names.get(c, c) for c in sorted(exits)]
            lines.append(f"⬇️ **이탈**: {', '.join(exit_names)}")
        lines.append("")

    if df.empty:
        lines.append("> 조건을 충족하는 종목이 없습니다.")
        out_path.write_text("\n".join(lines), encoding="utf-8")
        return out_path

    # ── 요약 테이블 ─────────────────────────────────────────────────────────────
    lines += [
        "## 통과 종목 요약",
        "",
        "| 순위 | 종목명 | 코드 | 매출증가 | 영익증가 | OPM개선 | 컨센1M | 총점 | 이익 | 컨센 | 밸류 | 리스크 | 수급 |",
        "|------|--------|------|---------|---------|--------|-------|------|------|------|------|-------|------|",
    ]
    for rank, row in enumerate(df.itertuples(), 1):
        op_str = "흑전" if getattr(row, "turnaround", False) else _fmt(getattr(row, "op_profit_growth", None), "%")
        new_mark = " 🆕" if row.code in new_entries else ""
        lines.append(
            f"| {rank} "
            f"| {row.name}{new_mark} "
            f"| {row.code} "
            f"| {_arrow(getattr(row, 'revenue_growth', None))} {_fmt(getattr(row, 'revenue_growth', None), '%')} "
            f"| {_arrow(getattr(row, 'op_profit_growth', None))} {op_str} "
            f"| {_arrow(getattr(row, 'opm_improvement', None))} {_fmt(getattr(row, 'opm_improvement', None), '%p')} "
            f"| {_arrow(getattr(row, 'op_revision_1m', None))} {_fmt(getattr(row, 'op_revision_1m', None), '%')} "
            f"| **{getattr(row, 'total_score', 0):.1f}** "
            f"| {getattr(row, 'score_earnings', 0):.0f} "
            f"| {getattr(row, 'score_consensus', 0):.0f} "
            f"| {getattr(row, 'score_valuation', 0):.0f} "
            f"| {getattr(row, 'score_risk', 0):.0f} "
            f"| {getattr(row, 'score_flow', 0):.0f} |"
        )
    lines.append("")

    # ── 종목별 상세 ─────────────────────────────────────────────────────────────
    lines += ["## 종목별 상세", ""]

    for rank, row in enumerate(df.itertuples(), 1):
        op_str   = "**흑자전환**" if getattr(row, "turnaround", False) else _fmt(getattr(row, "op_profit_growth", None), "%")
        new_mark = " 🆕 신규편입" if row.code in new_entries else ""

        lines += [
            f"### {rank}. {row.name} ({row.code}){new_mark}",
            "",
            f"- 시가총액: {_fmt(getattr(row, 'market_cap', None), '억', 0)}  |  커버리지: {getattr(row, 'analyst_count', 0)}개 증권사  |  **총점: {getattr(row, 'total_score', 0):.1f}점**",
            f"- Catalyst: _{_catalyst(row._asdict())}_",
            "",
            f"| 지표 | {ACTUAL_YEAR}A | {ESTIMATE_YEAR}E | 변화 |",
            f"|------|-------|-------|------|",
            f"| 매출액 (억원) | {_fmt(getattr(row, 'revenue_25a', None), '억', 0)} | {_fmt(getattr(row, 'revenue_26e', None), '억', 0)} | {_arrow(getattr(row, 'revenue_growth', None))} {_fmt(getattr(row, 'revenue_growth', None), '%')} |",
            f"| 영업이익 (억원) | {_fmt(getattr(row, 'op_profit_25a', None), '억', 0)} | {_fmt(getattr(row, 'op_profit_26e', None), '억', 0)} | {_arrow(getattr(row, 'op_profit_growth', None))} {op_str} |",
            f"| 영업이익률 (%) | {_fmt(getattr(row, 'opm_25a', None), '%')} | {_fmt(getattr(row, 'opm_26e', None), '%')} | {_arrow(getattr(row, 'opm_improvement', None))} {_fmt(getattr(row, 'opm_improvement', None), '%p')} |",
            "",
            f"**컨센서스 변화:** 1M {_arrow(getattr(row, 'op_revision_1m', None))} {_fmt(getattr(row, 'op_revision_1m', None), '%')}  |  3M {_arrow(getattr(row, 'op_revision_3m', None))} {_fmt(getattr(row, 'op_revision_3m', None), '%')}",
            "",
        ]

        # 컨센 추이
        trend = _consensus_trend(row.code)
        if trend:
            lines += [f"**영업이익 컨센 추이:** {trend}", ""]

        # 밸류에이션 / 리스크 / 수급
        val_parts = []
        if getattr(row, "fwd_per", None) and not pd.isna(row.fwd_per):
            val_parts.append(f"PER {row.fwd_per:.1f}배")
        if getattr(row, "price_52w_pct", None) and not pd.isna(row.price_52w_pct):
            val_parts.append(f"52주고점대비 {row.price_52w_pct:.1f}%")
        if val_parts:
            lines.append(f"**밸류에이션:** {' | '.join(val_parts)}")

        risk_parts = []
        if getattr(row, "debt_ratio", None) and not pd.isna(row.debt_ratio):
            risk_parts.append(f"부채비율 {row.debt_ratio:.0f}%")
        if getattr(row, "audit_ok", None) is not None and not pd.isna(row.audit_ok):
            risk_parts.append(f"감사의견 {'적정' if row.audit_ok else '⚠️비적정'}")
        if risk_parts:
            lines.append(f"**리스크:** {' | '.join(risk_parts)}")

        flow_parts = []
        if getattr(row, "foreign_net_3m", None) and not pd.isna(row.foreign_net_3m):
            flow_parts.append(f"외국인 3M {_arrow(row.foreign_net_3m)} {_fmt(row.foreign_net_3m, '억')}")
        if getattr(row, "institution_net_3m", None) and not pd.isna(row.institution_net_3m):
            flow_parts.append(f"기관 3M {_arrow(row.institution_net_3m)} {_fmt(row.institution_net_3m, '억')}")
        if flow_parts:
            lines.append(f"**수급:** {' | '.join(flow_parts)}")

        # ── Catalyst 섹션 ───────────────────────────────────────────────────
        cr = (catalyst_map or {}).get(row.code)
        if cr:
            lines += ["", f"**Catalyst 점수: {cr.catalyst_total:.2f}**  ({cr.summary})", ""]

            # 최근 30일 주요 이벤트
            from catalyst.catalyst_scorer import build_timeline
            timeline = build_timeline(cr.events, months=1)
            if timeline:
                lines += ["**최근 30일 주요 이벤트**", ""]
                lines += ["| 날짜 | 출처 | 카테고리 | 이벤트 | 점수 |",
                          "|------|------|---------|--------|------|"]
                for ev in timeline[:8]:
                    major = " 🔥" if ev.get("is_major") else ""
                    lines.append(
                        f"| {ev['date']} | {ev['source']} | {ev['category']}"
                        f" | {ev['title'][:40]}{major} | {ev['score']:.2f} |"
                    )
                lines.append("")

        lines.append("")

    # ── 푸터 ────────────────────────────────────────────────────────────────────
    lines += [
        "---",
        f"*생성일시: {today}  |  데이터 출처: FnGuide / 네이버금융 / DART / KRX*",
    ]

    out_path.write_text("\n".join(lines), encoding="utf-8")
    return out_path
