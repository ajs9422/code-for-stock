"""
================================================
  5. 전문가 추천 + 저평가 교집합 스크리너
================================================

[탐색 원리]
  유튜브 증권 채널과 네이버 증권 뉴스에서 전문가들이 자주 언급하는 종목을
  추출하고, KRX 재무지표 기반 저평가 필터와 교차 검증하여
  "전문가도 추천하고 실제로 저평가된" 종목만 골라낸다.

[탐색 절차]
  ① YouTube Data API로 증권 관련 채널 영상 제목 수집
     → 종목명 언급 횟수 집계 (API 키 없으면 이 단계 생략)
  ② 네이버 증권 뉴스 스크래핑
     → 기사 제목·본문에서 종목명 언급 빈도 집계
  ③ KRX 기본지표로 저평가 스크리닝 (1번 코드와 동일 기준)
     → PER·PBR·ROE 기반 종합 저평가점수 계산
  ④ 유튜브 언급 점수 + 뉴스 언급 점수 + 저평가 점수를 가중 합산
     → 최종 추천 점수 순 정렬

[가중치]
  유튜브 언급 빈도  (30%)
  뉴스 언급 빈도    (30%)
  저평가 점수       (40%)

[데이터 출처]
  YouTube Data API v3 (선택, 환경변수 YOUTUBE_API_KEY 필요)
  네이버 증권 뉴스 (finance.naver.com)
  pykrx KRX 기본지표 (_krx_utils.py 공용)
================================================
"""

import os
import sys
import time
import warnings
from datetime import datetime, timedelta, timezone

import pandas as pd
import requests
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
from rich.text import Text

from _krx_utils import suppress_output, build_stock_dicts, build_screener_df

try:
    from bs4 import BeautifulSoup
    BS4_OK = True
except ImportError:
    BS4_OK = False

warnings.filterwarnings("ignore")
console = Console()

# ──────────────────────────────────────────────
# 설정값
# ──────────────────────────────────────────────
YOUTUBE_API_KEY = os.getenv('YOUTUBE_API_KEY', '')

SEARCH_DAYS = 7
TOP_N       = 30

WEIGHTS = {"screener": 0.50, "youtube": 0.25, "news": 0.25}

YT_QUERIES = [
    "주식 추천 종목",
    "매수 추천 종목",
    "오늘의 추천 종목",
    "저평가 주식 추천",
    "급등 예상 종목",
]
YT_MAX_RESULTS = 50

NAVER_QUERIES = [
    "추천 종목",
    "매수 추천",
    "저평가 종목",
    "주식 추천",
    "급등 종목",
]


# ──────────────────────────────────────────────
# 텍스트 → 종목 언급 추출 (공통)
# ──────────────────────────────────────────────
def extract_mentions(texts: list, name_to_ticker: dict) -> dict:
    counts: dict = {}
    for text in texts:
        for name, ticker in name_to_ticker.items():
            if name in text:
                counts[ticker] = counts.get(ticker, 0) + 1
    return counts


def extract_mentions_with_reasons(
    texts: list, name_to_ticker: dict, source: str = ""
) -> tuple[dict, dict]:
    """언급 횟수 + 근거 텍스트 목록 반환"""
    counts: dict = {}
    reasons: dict = {}
    for text in texts:
        for name, ticker in name_to_ticker.items():
            if name in text:
                counts[ticker] = counts.get(ticker, 0) + 1
                reasons.setdefault(ticker, [])
                label = f"[{source}] {text[:60]}" if source else text[:60]
                if label not in reasons[ticker]:
                    reasons[ticker].append(label)
    return counts, reasons


def normalize_scores(count_dict: dict) -> dict:
    if not count_dict:
        return {}
    max_val = max(count_dict.values())
    if max_val == 0:
        return {k: 0.0 for k in count_dict}
    return {k: round(v / max_val * 100, 1) for k, v in count_dict.items()}


# ──────────────────────────────────────────────
# 모멘텀 (1개월 수익률)
# ──────────────────────────────────────────────
def fetch_momentum(tickers: list) -> dict:
    """FinanceDataReader로 최근 1개월 수익률(%) 계산"""
    try:
        import FinanceDataReader as fdr
    except ImportError:
        return {}

    start = (datetime.today() - timedelta(days=35)).strftime("%Y-%m-%d")
    momentum = {}
    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=32, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TextColumn("남은"),
        TimeRemainingColumn(),
        console=console,
        transient=False,
    ) as progress:
        task = progress.add_task("모멘텀 계산 중...", total=len(tickers))
        for ticker in tickers:
            progress.update(task, description=f"{ticker:<10} 모멘텀 계산")
            try:
                df = fdr.DataReader(ticker, start=start)
                if len(df) >= 2:
                    first = float(df["Close"].iloc[0])
                    last  = float(df["Close"].iloc[-1])
                    if first > 0:
                        momentum[ticker] = round((last - first) / first * 100, 1)
            except Exception:
                pass
            progress.advance(task)
            time.sleep(0.05)
    return momentum


# ──────────────────────────────────────────────
# YouTube
# ──────────────────────────────────────────────
def fetch_youtube_titles(api_key: str, queries: list, days: int) -> list:
    if not api_key:
        return []

    published_after = (
        datetime.now(timezone.utc) - timedelta(days=days)
    ).strftime('%Y-%m-%dT%H:%M:%SZ')

    titles = []
    for query in queries:
        try:
            resp = requests.get(
                "https://www.googleapis.com/youtube/v3/search",
                params={
                    "key": api_key,
                    "q": query,
                    "type": "video",
                    "publishedAfter": published_after,
                    "relevanceLanguage": "ko",
                    "regionCode": "KR",
                    "maxResults": YT_MAX_RESULTS,
                    "part": "snippet",
                },
                timeout=15,
            )
            if resp.status_code == 200:
                for item in resp.json().get("items", []):
                    title = item.get("snippet", {}).get("title", "")
                    if title:
                        titles.append(title)
            elif resp.status_code == 403:
                console.print("[yellow]  YouTube API 쿼터 초과 또는 키 오류[/yellow]")
                break
        except Exception as e:
            console.print(f"[yellow]  YouTube API 오류: {e}[/yellow]")
        time.sleep(0.5)

    return titles


# ──────────────────────────────────────────────
# 네이버 뉴스
# ──────────────────────────────────────────────
def fetch_naver_news_titles(queries: list) -> list:
    if not BS4_OK:
        console.print("[yellow]  beautifulsoup4 미설치 → pip install beautifulsoup4[/yellow]")
        return []

    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
        )
    }
    titles = []

    for query in queries:
        # 네이버 금융 뉴스 검색 (3페이지)
        for page in range(1, 4):
            try:
                resp = requests.get(
                    "https://finance.naver.com/news/news_search.naver",
                    params={"q": query, "pd": 2, "page": page},
                    headers=headers,
                    timeout=10,
                )
                soup = BeautifulSoup(resp.text, 'html.parser')
                found = False
                for sel in [
                    '.articleSubject a', '.articleTitle a',
                    'dd.articleTitle a', '.itemSubj a', 'ul.newsList li a',
                ]:
                    items = soup.select(sel)
                    if items:
                        titles.extend(
                            el.get_text(strip=True) for el in items
                            if len(el.get_text(strip=True)) >= 5
                        )
                        found = True
                        break
                if not found:
                    skip = {"더보기", "이전", "다음", "뉴스", "검색"}
                    for a in soup.find_all('a'):
                        txt = a.get_text(strip=True)
                        if len(txt) >= 5 and txt not in skip:
                            titles.append(txt)
            except Exception:
                pass
            time.sleep(0.3)

        # 네이버 통합 뉴스 검색 추가
        try:
            resp = requests.get(
                "https://search.naver.com/search.naver",
                params={"where": "news", "query": query + " 주식", "sort": "1", "pd": "4"},
                headers=headers,
                timeout=10,
            )
            soup = BeautifulSoup(resp.text, 'html.parser')
            for el in soup.select('.news_tit, .api_txt_lines.total_tit'):
                txt = el.get_text(strip=True)
                if len(txt) >= 5:
                    titles.append(txt)
        except Exception:
            pass
        time.sleep(0.3)

    return titles


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main():
    console.print()
    console.print(Panel.fit(
        "[bold cyan]전문가 언급 + 저평가 교집합 스크리너[/bold cyan]\n"
        "[dim]YouTube 영상 제목  ×  네이버 증권 뉴스  ×  KRX 기본지표 저평가[/dim]",
        border_style="cyan"
    ))
    console.print()

    # ① KRX 종목 사전
    console.print("[cyan]KRX 전종목 사전 빌드 중...[/cyan]")
    with suppress_output():
        name_to_ticker, ticker_to_name = build_stock_dicts()
    console.print(f"  [dim]{len(name_to_ticker):,}개 종목 로드 완료[/dim]")

    # ② 저평가 스크리너
    console.print()
    console.print("[cyan]① KRX 기본지표 수집 및 저평가 필터링...[/cyan]")
    screener_df = build_screener_df(console=console)
    if screener_df.empty:
        console.print("[red]스크리너 데이터 없음. 종료.[/red]")
        return

    # ③ YouTube 영상 제목 수집
    console.print()
    yt_scores: dict = {}
    if YOUTUBE_API_KEY:
        console.print(f"[cyan]② YouTube 영상 제목 수집 중... (최근 {SEARCH_DAYS}일, {len(YT_QUERIES)}개 쿼리)[/cyan]")
        yt_titles = fetch_youtube_titles(YOUTUBE_API_KEY, YT_QUERIES, SEARCH_DAYS)
        console.print(f"  [dim]수집된 영상 제목: {len(yt_titles)}개[/dim]")
        yt_raw, yt_reasons = extract_mentions_with_reasons(yt_titles, name_to_ticker, "YouTube")
        yt_scores = normalize_scores(yt_raw)
        top_yt    = sorted(yt_raw.items(), key=lambda x: -x[1])[:5]
        if top_yt:
            console.print(
                "  [dim]YouTube 상위 언급:[/dim] " +
                "  ".join(f"[green]{ticker_to_name.get(t, t)}[/green]({c}회)" for t, c in top_yt)
            )
    else:
        yt_reasons = {}
        console.print("[yellow]② YouTube API 키 없음 → YouTube 분석 제외[/yellow]")
        console.print("  [dim]환경변수 YOUTUBE_API_KEY 설정 시 자동 포함됩니다[/dim]")

    # ④ 네이버 뉴스 수집
    console.print()
    console.print(f"[cyan]③ 네이버 증권 뉴스 수집 중... (최근 {SEARCH_DAYS}일)[/cyan]")
    naver_titles = fetch_naver_news_titles(NAVER_QUERIES)
    console.print(f"  [dim]수집된 기사 제목: {len(naver_titles)}개[/dim]")
    naver_raw, naver_reasons = extract_mentions_with_reasons(naver_titles, name_to_ticker, "뉴스")
    naver_scores = normalize_scores(naver_raw)
    top_news     = sorted(naver_raw.items(), key=lambda x: -x[1])[:5]
    if top_news:
        console.print(
            "  [dim]뉴스 상위 언급:[/dim] " +
            "  ".join(f"[green]{ticker_to_name.get(t, t)}[/green]({c}회)" for t, c in top_news)
        )

    # ⑤ 복합 점수 계산
    console.print()
    console.print("[cyan]④ 복합 점수 계산...[/cyan]")

    if YOUTUBE_API_KEY:
        w_sc, w_yt, w_nw = WEIGHTS["screener"], WEIGHTS["youtube"], WEIGHTS["news"]
    else:
        w_sc, w_yt, w_nw = 0.60, 0.0, 0.40

    screener_df["yt_score"]   = screener_df["티커"].map(yt_scores).fillna(0)
    screener_df["news_score"] = screener_df["티커"].map(naver_scores).fillna(0)
    screener_df["복합점수"]   = (
        screener_df["종합_저평가점수"] * w_sc +
        screener_df["yt_score"]        * w_yt +
        screener_df["news_score"]      * w_nw
    ).round(1)
    # 교집합 보너스: YouTube·뉴스 양쪽 모두 언급 시 강조
    screener_df["최종점수"] = (
        screener_df["복합점수"] +
        (screener_df["yt_score"]   > 0).astype(int) * 3 +
        (screener_df["news_score"] > 0).astype(int) * 3
    ).round(1)

    result = screener_df.nlargest(TOP_N, "최종점수")[[
        "종목명", "티커", "현재가",
        "최종점수", "종합_저평가점수", "yt_score", "news_score",
        "PER", "PBR", "ROE", "시가총액_억", "시장",
    ]].reset_index(drop=True)
    result["시가총액_억"] = result["시가총액_억"].round(2)
    result["ROE"] = result["ROE"].round(1)
    result.index += 1

    yt_cnt   = int((result["yt_score"]   > 0).sum())
    news_cnt = int((result["news_score"] > 0).sum())
    both_cnt = int(((result["yt_score"] > 0) & (result["news_score"] > 0)).sum())
    console.print(f"  [dim]TOP {TOP_N} 중 YouTube 언급 {yt_cnt}개 · 뉴스 언급 {news_cnt}개 · 양쪽 모두 {both_cnt}개[/dim]")

    # 모멘텀 & 전문가 추천 사유 컬럼 추가
    console.print("[cyan]  모멘텀(1M) 계산 중...[/cyan]")
    result_tickers = result["티커"].tolist()
    momentum_map   = fetch_momentum(result_tickers)

    result["모멘텀(1M%)"] = result["티커"].map(lambda t: momentum_map.get(t, None))

    # result를 ticker 기준으로 빠르게 조회하기 위한 dict
    result_map = result.set_index("티커").to_dict("index")

    def build_reason(ticker: str) -> str:
        parts = []
        # ① 뉴스/YouTube 제목에서 언급된 경우 → 실제 제목 표시
        for r in yt_reasons.get(ticker, [])[:2]:
            parts.append(r)
        for r in naver_reasons.get(ticker, [])[:3]:
            parts.append(r)
        # ② 언급 없으면 → 펀더멘털 수치로 사유 자동 생성
        if not parts:
            row = result_map.get(ticker, {})
            per  = row.get("PER",  0)
            pbr  = row.get("PBR",  0)
            roe  = row.get("ROE",  0)
            subs = []
            if per > 0:
                subs.append(f"PER {per:.1f}배{'(저평가)' if per < 10 else ''}")
            if pbr > 0:
                subs.append(f"PBR {pbr:.2f}배{'(저평가)' if pbr < 1 else ''}")
            if roe != 0:
                subs.append(f"ROE {roe:.1f}%{'(우량)' if roe >= 15 else ''}")
            score = row.get("종합_저평가점수", 0)
            if score:
                subs.append(f"저평가점수 {score:.0f}점")
            parts = subs if subs else ["스크리너 저평가 기준 충족"]
        return " | ".join(parts)

    result["전문가추천사유"] = result["티커"].map(build_reason)

    # ⑥ 터미널 출력
    console.print()
    console.rule(f"[bold cyan] 전문가 언급 + 저평가 교집합 TOP {TOP_N} [/bold cyan]")
    console.print()

    table = Table(
        box=box.SIMPLE_HEAD, show_header=True,
        header_style="bold cyan", border_style="dim",
        row_styles=["", "dim"],
    )
    table.add_column("#",           style="dim", width=3,  justify="right")
    table.add_column("종목명",       width=14)
    table.add_column("티커",         width=8,  justify="center")
    table.add_column("현재가",       width=10, justify="right")
    table.add_column("최종\n점수",   width=7,  justify="right")
    table.add_column("저평가\n점수", width=7,  justify="right")
    table.add_column("YT\n언급",     width=6,  justify="right")
    table.add_column("뉴스\n언급",   width=6,  justify="right")
    table.add_column("PER",          width=6,  justify="right")
    table.add_column("PBR",          width=6,  justify="right")
    table.add_column("ROE%",         width=7,  justify="right")
    table.add_column("시장",         width=7,  justify="center")

    def cs(val, hi=70, mid=50):
        v = float(val)
        color = "bright_green" if v >= hi else "yellow" if v >= mid else "dim"
        return Text(f"{v:.0f}", style=color)

    def mention_cell(val):
        v = float(val)
        return Text(f"{v:.0f}", style="bright_green bold") if v > 0 else Text("-", style="dim")

    for i, row in result.iterrows():
        table.add_row(
            str(i),
            str(row["종목명"])[:12],
            str(row["티커"]),
            f"{int(row['현재가']):,}" if row.get("현재가", 0) > 0 else "-",
            cs(row["최종점수"]),
            cs(row["종합_저평가점수"]),
            mention_cell(row["yt_score"]),
            mention_cell(row["news_score"]),
            f"{row['PER']:.1f}",
            f"{row['PBR']:.2f}",
            f"{row['ROE']:.1f}",
            str(row.get("시장", "")),
        )

    console.print(table)
    console.print()

    w_txt = f"저평가 {w_sc:.0%}  ·  뉴스 {w_nw:.0%}"
    if YOUTUBE_API_KEY:
        w_txt = f"저평가 {w_sc:.0%}  ·  YouTube {w_yt:.0%}  ·  뉴스 {w_nw:.0%}"
    console.print(f"[dim]■ 가중치: {w_txt}   (교집합 보너스 +3점 × 채널수)[/dim]")
    console.print("[bold yellow]⚠  본 결과는 참고용이며 투자 조언이 아닙니다.[/bold yellow]")
    console.print()

    # ⑦ Excel 저장
    filename = f"5_전문가추천_저평가종목_{datetime.today().strftime('%Y_%m%d_%H%M')}.xlsx"
    result.to_excel(filename, index=True, index_label="순번")

    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font

        wb = load_workbook(filename)
        ws = wb.active
        name_col_idx = ticker_col_idx = None
        for cell in ws[1]:
            if cell.value == "종목명": name_col_idx = cell.column
            if cell.value == "티커":   ticker_col_idx = cell.column

        # 열 너비 자동 조정 (한글 2배 폭 반영)
        def _cell_width(val: str) -> int:
            return sum(2 if ord(c) > 127 else 1 for c in val)

        from openpyxl.styles import Alignment

        # 하이퍼링크 + 티커 컬럼 삭제
        if name_col_idx and ticker_col_idx:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                name_cell   = row[name_col_idx   - 1]
                ticker_cell = row[ticker_col_idx - 1]
                if ticker_cell.value:
                    code = str(ticker_cell.value).zfill(6)
                    name_cell.hyperlink = f"https://finance.naver.com/item/main.naver?code={code}"
                    name_cell.font = Font(color="0563C1", underline="single", bold=True)
            ws.delete_cols(ticker_col_idx)

        # 전체 가운데 정렬 + 숫자 포맷
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")
                if isinstance(cell.value, float):
                    cell.value = round(cell.value, 1)
                    cell.number_format = "#,##0.0"
                elif isinstance(cell.value, int):
                    cell.number_format = "#,##0"

        # 전문가추천사유 열 줄바꿈 설정
        reason_col_idx = None
        for cell in ws[1]:
            if cell.value == "전문가추천사유":
                reason_col_idx = cell.column
                break
        if reason_col_idx:
            ws.column_dimensions[ws.cell(1, reason_col_idx).column_letter].width = 120
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                row[reason_col_idx - 1].alignment = Alignment(horizontal="center", wrap_text=False)

        # 열 너비 재조정
        for col in ws.columns:
            col_letter = col[0].column_letter
            if ws.column_dimensions[col_letter].width == 120:
                continue
            max_len = max((_cell_width(str(cell.value or "")) for cell in col), default=0)
            ws.column_dimensions[col_letter].width = min(max(max_len + 2, 10), 80)

        wb.save(filename)
    except Exception as e:
        console.print(f"[yellow]하이퍼링크 추가 실패: {e}[/yellow]")

    console.print(f"[cyan]✓ 저장 완료: {filename}[/cyan]")
    console.print()


if __name__ == "__main__":
    main()
