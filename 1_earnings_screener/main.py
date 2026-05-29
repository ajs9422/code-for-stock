"""
================================================
  이익 추정치 급등 종목 스크리너 — 오케스트레이션
================================================

[실행 방법]
  python main.py                        # 오늘 날짜, Top 10
  python main.py --date 20260423        # 특정 날짜
  python main.py --top 20               # Top N 변경

[전체 흐름]
  1. DB 초기화 (테이블 없으면 생성)
  2. FDR로 유니버스 수집 (KOSPI+KOSDAQ, 시총 2,000억↑)
  3. FnGuide에서 종목별 컨센서스 수집 (폴백: 네이버 금융)
  4. DB 스냅샷 저장 + 1M/3M 컨센 변화율 계산
  5. 5개 축 필터 순차 적용 (이익→컨센→밸류→리스크→수급)
  6. 5축 가중 점수화 → Top N 추출
  7. 일별 Markdown 리포트 저장 + 이메일 발송
  8. 완료 요약 출력

[5개 축 필터 조건]
  1) 이익 모멘텀:    매출 ≥ 10%, 영익 ≥ 20%(or 흑전), OPM ≥ 3%p
  2) 컨센 모멘텀:    1M 컨센 상향, 3M 상향 (없으면 통과)
  3) 밸류에이션:     Fwd PER < 35, 52주고점대비 < 95% (없으면 통과)
  4) 리스크 제거:    관리종목 제외, 부채비율 < 300%, 감사의견 적정
  5) 수급:           외국인/기관 3M 순매수 > 0 (없으면 통과)

[5축 가중치]
  이익 모멘텀 0.35 | 컨센 모멘텀 0.20 | 밸류에이션 0.20 | 리스크 0.15 | 수급 0.10
"""

import argparse
import sys
import time
from datetime import date, datetime

from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.table import Column, Table

import config
from config import REPORT_DIR
import storage.db as db
from crawler import fnguide, naver, krx
from crawler.dart import fetch_risk
from crawler.naver import fetch_valuation_naver
from screener import filters as new_filters
from screener import scoring
from screener.filter import _compute_metrics      # 파생 지표 계산 재사용
from report import daily_report, markdown
from catalyst import news_crawler, keyword_matcher, dart_catalyst, catalyst_scorer
from catalyst.catalyst_scorer import CatalystResult

console = Console(force_terminal=True)


# ── 필터 로그 ─────────────────────────────────────────────────────────────────

def _write_filter_log(
    run_date,
    universe_count: int,
    consensus_count: int,
    filter_log: list[tuple[str, int, int]],
) -> None:
    """5축 필터링 결과를 텍스트 로그로 저장 (조건 튜닝 참고용)."""
    from datetime import datetime
    REPORT_DIR.mkdir(parents=True, exist_ok=True)
    log_path = REPORT_DIR / f"{run_date}_filter_log.txt"

    now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    lines = [
        "=" * 52,
        "  5축 필터링 로그",
        "=" * 52,
        f"실행일시: {now}",
        f"기준일:   {run_date}",
        "",
        "[필터 조건]",
        f"  시총 최소:             {config.MIN_MARKET_CAP:,}억",
        f"  커버리지 최소:         {config.MIN_COVERAGE}개",
        f"  매출 증가율 최소:      {config.MIN_REVENUE_GROWTH:.1f}%",
        f"  영업이익 증가율 최소:  {config.MIN_OP_PROFIT_GROWTH:.1f}%",
        f"  OPM 개선 최소:         {config.MIN_OPM_IMPROVEMENT:.1f}%p",
        f"  Fwd PER 상한:          {config.MAX_FWD_PER:.1f}",
        f"  52주고점 대비 상한:    {config.MAX_PRICE_52W_PCT:.1f}%",
        f"  급등률 상한:           {config.MAX_SURGE_FROM_LOW:.1f}%",
        f"  부채비율 상한:         {config.MAX_DEBT_RATIO:.1f}%",
        "",
        "[필터링 결과]",
        f"  {'유니버스':<22} {universe_count:>5,}개",
        f"  {'컨센 수집 완료':<22} {consensus_count:>5,}개",
    ]
    for name, before, after in filter_log:
        drop = before - after
        lines.append(f"  {name:<22} {before:>5,} → {after:>4,}개  (탈락 {drop:,}개)")

    final = filter_log[-1][2] if filter_log else consensus_count
    lines += [
        "-" * 52,
        f"  {'최종 통과':<22} {final:>5,}개",
        "",
    ]

    log_path.write_text("\n".join(lines), encoding="utf-8")
    console.print(f"  [green]필터 로그:   {log_path}[/green]")


# ── CLI ──────────────────────────────────────────────────────────────────────

def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="이익 추정치 급등 종목 스크리너")
    p.add_argument("--date", type=str, default=None,
                   help="실행 기준 날짜 (YYYYMMDD, 기본: 오늘)")
    p.add_argument("--top",  type=int, default=10,
                   help="상위 몇 개 종목 추출 (기본: 10)")
    return p.parse_args()


# ── 컨센서스 수집 ─────────────────────────────────────────────────────────────

def collect_consensus(universe) -> list:
    """유니버스 전종목 컨센서스 수집. FnGuide 실패 시 네이버 폴백."""
    results = []
    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=32, no_wrap=True)),
        BarColumn(bar_width=20),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
        refresh_per_second=2,
        transient=False,
    ) as progress:
        task = progress.add_task("컨센서스 수집 중...", total=len(universe))
        for row in universe.itertuples():
            data = fnguide.fetch_consensus(row.code, row.name, row.market_cap)
            if data is None:
                data = naver.fetch_consensus_naver(row.code, row.name, row.market_cap)
            if data:
                results.append(data)
            progress.advance(task)
    return results


# ── 추가 데이터 수집 (밸류/리스크/수급) ──────────────────────────────────────

def enrich_candidates(candidates: list, admin_codes: set) -> list:
    """
    이익 모멘텀 통과 종목에 한해 밸류에이션/리스크/수급 데이터 추가 수집.
    관리종목 여부는 admin_codes set으로 일괄 처리.
    """
    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=32, no_wrap=True)),
        BarColumn(bar_width=20),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        console=console,
        refresh_per_second=2,
        transient=False,
    ) as progress:
        task = progress.add_task("추가 데이터 수집 중...", total=len(candidates))
        for data in candidates:
            # 관리종목 일괄 체크
            data.is_admin = data.code in admin_codes

            # 밸류에이션 (네이버)
            val = fetch_valuation_naver(data.code)
            data.fwd_per         = val.get("fwd_per")
            data.price           = val.get("price")
            data.price_52w_high  = val.get("price_52w_high")
            data.price_52w_low   = val.get("price_52w_low")
            data.price_52w_pct   = val.get("price_52w_pct")
            data.price_surge_pct = val.get("price_surge_pct")

            # 리스크 (DART)
            risk = fetch_risk(data.code)
            data.audit_ok   = risk.get("audit_ok")
            data.debt_ratio = risk.get("debt_ratio")

            # 수급 (KRX)
            flow = krx.fetch_investor_flow(data.code)
            data.foreign_net_3m     = flow.get("foreign_net_3m")
            data.institution_net_3m = flow.get("institution_net_3m")

            progress.advance(task)
    return candidates


# ── 결과 출력 ─────────────────────────────────────────────────────────────────

def print_summary(df, elapsed: float):
    import pandas as pd

    console.print()
    console.rule("[bold cyan]스크리닝 결과[/bold cyan]")
    console.print()

    if df is None or df.empty:
        console.print("  [yellow]조건을 충족하는 종목이 없습니다.[/yellow]")
        return

    table = Table(box=box.SIMPLE_HEAD, header_style="bold cyan", border_style="dim")
    table.add_column("순위",    width=4,  justify="right", style="dim")
    table.add_column("종목명",  width=14)
    table.add_column("코드",    width=8,  justify="center")
    table.add_column("매출증가", width=9, justify="right")
    table.add_column("영익증가", width=10, justify="right")
    table.add_column("컨센1M",  width=9,  justify="right")
    table.add_column("총점",    width=7,  justify="right", style="bold yellow")
    table.add_column("이익",    width=5,  justify="right", style="cyan")
    table.add_column("컨센",    width=5,  justify="right", style="cyan")
    table.add_column("밸류",    width=5,  justify="right", style="cyan")
    table.add_column("리스크",  width=6,  justify="right", style="cyan")
    table.add_column("수급",    width=5,  justify="right", style="cyan")

    def _f(v, suf="", dec=1):
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return "N/A"
        return f"{v:+.{dec}f}{suf}" if suf else f"{v:.{dec}f}"

    for rank, row in enumerate(df.itertuples(), 1):
        op_str = "흑전" if getattr(row, "turnaround", False) else _f(getattr(row, "op_profit_growth", None), "%")
        table.add_row(
            str(rank),
            getattr(row, "name", ""),
            getattr(row, "code", ""),
            _f(getattr(row, "revenue_growth", None), "%"),
            op_str,
            _f(getattr(row, "op_revision_1m", None), "%"),
            f"{getattr(row, 'total_score', 0):.1f}",
            f"{getattr(row, 'score_earnings', 0):.0f}",
            f"{getattr(row, 'score_consensus', 0):.0f}",
            f"{getattr(row, 'score_valuation', 0):.0f}",
            f"{getattr(row, 'score_risk', 0):.0f}",
            f"{getattr(row, 'score_flow', 0):.0f}",
        )

    console.print(table)
    console.print(
        f"  통과 종목: [green]{len(df)}개[/green]  "
        f"소요시간: [cyan]{elapsed:.0f}초[/cyan]"
    )
    console.print()


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    args = _parse_args()
    run_date = (datetime.strptime(args.date, "%Y%m%d").date()
                if args.date else date.today())
    top_n = args.top
    start = time.time()

    console.print()
    console.print(Panel.fit(
        "[bold cyan]이익 추정치 급등 종목 스크리너 (5축)[/bold cyan]\n"
        f"[dim]{config.ACTUAL_YEAR}A → {config.ESTIMATE_YEAR}E  |  기준일: {run_date}  |  Top {top_n}[/dim]",
        border_style="cyan",
    ))
    console.print()

    # ── 1. DB 초기화 ──────────────────────────────────────────────────────────
    db.init()
    console.print("  [dim]DB 초기화 완료[/dim]")

    # ── 2. 유니버스 수집 ──────────────────────────────────────────────────────
    console.rule("[cyan]유니버스 수집[/cyan]")
    universe = fnguide.get_universe()
    if universe.empty:
        console.print("[red]유니버스 수집 실패. 종료.[/red]")
        sys.exit(1)

    if config.TEST_MODE:
        universe = universe.head(5)
        console.print("  [yellow]테스트 모드: 5개 종목만 처리[/yellow]")

    console.print(f"  [green]총 {len(universe):,}개 종목 (시총 {config.MIN_MARKET_CAP:,}억↑)[/green]")
    console.print()

    # ── 3. 컨센서스 수집 ──────────────────────────────────────────────────────
    console.rule("[cyan]컨센서스 수집 (FnGuide → 네이버 폴백)[/cyan]")
    candidates = collect_consensus(universe)
    console.print(f"  [green]수집 완료: {len(candidates):,}개[/green]")
    console.print()

    if not candidates:
        console.print("[red]컨센서스 수집 결과가 없습니다. FnGuide/네이버 접근 실패 가능성 있음.[/red]")
        sys.exit(1)

    # ── 4. DB 스냅샷 저장 + 1M/3M 컨센 변화율 계산 ──────────────────────────
    db.save_snapshot(candidates, run_date)
    candidates = db.fill_revision(candidates)
    candidates = db.fill_revision_3m(candidates)
    console.print(f"  [dim]스냅샷 저장 완료 ({run_date})[/dim]")

    # 파생 지표 계산 (증가율·OPM 개선·컨센 상향률)
    for d in candidates:
        _compute_metrics(d)

    # ── 5. 5개 축 필터 ────────────────────────────────────────────────────────
    console.rule("[cyan]5축 필터링[/cyan]")
    df = new_filters.to_dataframe(candidates)
    total_before = len(df)

    filter_log: list[tuple[str, int, int]] = []

    filter_steps = [
        (new_filters.earnings_momentum_filter,  "1축 이익 모멘텀"),
        (new_filters.consensus_momentum_filter, "2축 컨센 모멘텀"),
    ]

    for fn, name in filter_steps:
        before = len(df)
        df = fn(df)
        filter_log.append((name, before, len(df)))
        console.print(f"  {name}: {before:,} → [green]{len(df):,}[/green]개")

    # 이익+컨센 통과 종목만 추가 데이터 수집 (API 효율화)
    if not df.empty:
        console.print()
        console.rule("[cyan]추가 데이터 수집 (밸류/리스크/수급)[/cyan]")
        passing_codes = set(df["code"].tolist())
        passing_items = [d for d in candidates if d.code in passing_codes]

        console.print("  [dim]관리종목 목록 조회 중...[/dim]")
        admin_codes = krx.fetch_admin_stocks()

        passing_items = enrich_candidates(passing_items, admin_codes)

        # 수집된 추가 데이터를 DataFrame에 반영
        enriched = {d.code: d for d in passing_items}
        for col in ("fwd_per", "price_52w_pct", "price_surge_pct", "is_admin",
                    "debt_ratio", "audit_ok", "foreign_net_3m", "institution_net_3m"):
            df[col] = df["code"].map(lambda c: getattr(enriched.get(c), col, None))
        console.print()

        # 나머지 3개 축 필터
        remaining_steps = [
            (new_filters.valuation_filter,      "3축 밸류에이션"),
            (new_filters.risk_exclusion_filter,  "4축 리스크 제거"),
            (new_filters.flow_filter,            "5축 수급"),
        ]
        for fn, name in remaining_steps:
            before = len(df)
            df = fn(df)
            filter_log.append((name, before, len(df)))
            console.print(f"  {name}: {before:,} → [green]{len(df):,}[/green]개")

    console.print(f"\n  전체: {total_before:,}개 → 최종 통과: [green]{len(df):,}개[/green]")
    console.print()

    # ── 6. 5축 점수화 → Top N ─────────────────────────────────────────────────
    if not df.empty:
        df = scoring.score(df, top_n=top_n)

    # ── 6.5. Catalyst 분석 ────────────────────────────────────────────────────
    catalyst_map: dict[str, CatalystResult] = {}
    if not df.empty:
        console.rule("[cyan]Catalyst 분석 (뉴스 + DART)[/cyan]")
        code_to_data = {d.code: d for d in candidates}
        with Progress(
            SpinnerColumn(),
            TextColumn("{task.description}", table_column=Column(width=36, no_wrap=True)),
            BarColumn(bar_width=20),
            MofNCompleteColumn(),
            TimeElapsedColumn(),
            console=console,
            refresh_per_second=2,
            transient=False,
        ) as progress:
            task = progress.add_task("Catalyst 분석 중...", total=len(df))
            for row in df.itertuples():
                code = row.code
                # 뉴스 크롤링 + 키워드 매칭
                news       = news_crawler.fetch_news(code, days=30)
                matches    = keyword_matcher.match_news_list(news)
                news_evts  = catalyst_scorer.score_from_news(news, matches)

                # DART 공시
                rev_25a    = getattr(code_to_data.get(code), "revenue_25a", None)
                discs      = dart_catalyst.fetch_disclosures(code, days=60, revenue_25a=rev_25a)
                dart_evts  = catalyst_scorer.score_from_dart(discs)

                result     = catalyst_scorer.aggregate(code, news_evts, dart_evts)
                catalyst_map[code] = result

                progress.update(task, advance=1,
                    description=f"Catalyst 분석 중... [{code}]")

        df["catalyst_total"] = df["code"].map(
            lambda c: catalyst_map.get(c, CatalystResult(c)).catalyst_total
        )
        df["catalyst_summary"] = df["code"].map(
            lambda c: catalyst_map.get(c, CatalystResult(c)).summary
        )
        console.print()

    # ── 7. 결과 저장 ──────────────────────────────────────────────────────────
    # 기존 DB 테이블 호환: list[tuple[ConsensusData, Score]] 형태로 변환
    if not df.empty:
        from screener.scorer import Score
        ranked_compat = []
        code_to_data = {d.code: d for d in candidates}
        for row in df.itertuples():
            data = code_to_data.get(row.code)
            if data:
                sc = Score(code=row.code, name=row.name, total=row.total_score)
                ranked_compat.append((data, sc))
        db.save_results(ranked_compat, run_date)
    else:
        ranked_compat = []

    # ── 8. 리포트 생성 ────────────────────────────────────────────────────────
    report_path = daily_report.generate(
        df if not df.empty else df.__class__(),
        run_date, top_n,
        catalyst_map=catalyst_map,
    )
    console.print(f"  [green]일별 리포트: {report_path}[/green]")

    # 기존 markdown + Excel + 이메일도 유지
    md_path = markdown.generate(ranked_compat, run_date)
    console.print(f"  [green]Markdown:    {md_path}[/green]")

    _write_filter_log(run_date, len(universe), len(candidates), filter_log)

    # ── 요약 출력 ─────────────────────────────────────────────────────────────
    elapsed = time.time() - start
    print_summary(df if not df.empty else None, elapsed)


if __name__ == "__main__":
    main()
