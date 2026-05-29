"""
================================================
  멀티팩터 주식 스캐너 — 동작 원리
================================================

[전체 흐름 요약]
  유튜브 검색 → 자막 추출 → AI 종목 추출 → KRX 팩터 채점 → 엑셀 저장

[단계별 상세 설명]

  ① 유튜브 검색 (search_recent_videos)
     - SEARCH_KEYWORD("이권희 대표" 등)로 유튜브 검색 결과 HTML을 직접 요청
     - HTML에서 정규식으로 videoId(11자리) 추출 → 최신 N개 영상 ID 확보

  ② 유튜브 자막 → AI 종목 추출 (get_expert_picks_real)
     - YouTubeTranscriptApi로 각 영상의 한국어 자막 전체 텍스트 수집
     - Gemini(gemini-1.5-flash)에 자막을 넘겨 "추천/긍정 언급 종목명 리스트"를 요청
     - AI 응답에서 Python 리스트 패턴을 정규식으로 추출 후 ast.literal_eval로 안전 파싱
     - 결과: {종목명 → 해당 영상 URL 집합} 매핑 딕셔너리 구성

  ③ KRX 유니버스 로드
     - pykrx로 오늘 기준 KOSPI + KOSDAQ 전종목 코드 조회 (약 2,500개)
     - 장 마감·야간 등 KRX 응답이 없으면 코어 7개 종목으로 대체
     - 전종목 처리 시 소요 시간: 약 10~15분 (종목당 0.2초 딜레이)

  ④ 종목별 팩터 채점 (calculate_momentum_score)
     각 종목에 아래 4개 팩터를 적용해 0~100점 합산:

     팩터                  조건                              만점
     ─────────────────────────────────────────────────────────
     AI-유튜브 추천         ②에서 추출된 종목 목록에 포함       30
     외인 수급 유입         최근 30일 외국인 보유 비율 증가       30
     기술적 반등            현재가 > 5일 이동평균                20
     EPS 흑자               당일 기준 EPS > 0                   20
     ─────────────────────────────────────────────────────────
     최대 합계                                                  100

  ⑤ 엑셀 저장 (_build_excel)
     - 시트1 "스크리닝 결과": 팩터별 점수 컬럼 + 합계 내림차순, 종목명 클릭 시 네이버 링크
     - 시트2 "분석 조건":     이번 실행에 사용된 config 값 기록
     - output/YYYYMMDD_HHMM_최종리포트.xlsx 로 저장

[설정 파일]
  config.py → SEARCH_KEYWORD, MAX_VIDEOS, LOOKBACK_DAYS, MARKETS 등
  .env      → GEMINI_API_KEY, DART_API_KEY, EMAIL_USER/PASS/TO
"""

import ast
import os
import re
import time
import urllib.parse
import urllib.request
from datetime import datetime, timedelta

import pandas as pd
from google import genai
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
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
from youtube_transcript_api import YouTubeTranscriptApi

import config

# pykrx import 전에 KRX 자격증명 주입 (로그인 실패 메시지 방지)
if config.KRX_ID:
    os.environ.setdefault('KRX_ID', config.KRX_ID)
    os.environ.setdefault('KRX_PW', config.KRX_PW)

from pykrx import stock

# ==========================================
# 1. 환경 설정 및 API 키
# ==========================================
if not config.GEMINI_API_KEY:
    raise RuntimeError(".env 파일에 GEMINI_API_KEY가 설정되지 않았습니다.")

_gemini  = genai.Client(api_key=config.GEMINI_API_KEY)
console  = Console(force_terminal=True)

end_date   = datetime.now().strftime("%Y%m%d")
start_date = (datetime.now() - timedelta(days=config.LOOKBACK_DAYS)).strftime("%Y%m%d")

# ── 스타일 상수 ───────────────────────────────────────────────────────────────
_HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")   # 진한 파랑
_HEADER_FONT  = Font(bold=True, color="FFFFFF", size=10)
_LINK_FONT    = Font(color="0563C1", underline="single", bold=True, size=10)
_BASE_FONT    = Font(size=10)
_SCORE_FULL   = PatternFill("solid", fgColor="C6EFCE")   # 연두 — 만점
_SCORE_PART   = PatternFill("solid", fgColor="FFEB9C")   # 노랑 — 부분 점수
_SCORE_ZERO   = PatternFill("solid", fgColor="FFC7CE")   # 연빨 — 0점
_CENTER       = Alignment(horizontal="center", vertical="center")
_LEFT         = Alignment(horizontal="left",   vertical="center")
_THIN_BORDER  = Border(
    bottom=Side(style="thin", color="D9D9D9"),
)


def _col_width(text: str) -> float:
    w = 0.0
    for ch in str(text):
        w += 2.0 if ord(ch) > 127 else 1.0
    return max(w + 2, 6)


def _auto_width(ws):
    for col_cells in ws.columns:
        max_w = 6.0
        for cell in col_cells:
            if cell.value is not None:
                max_w = max(max_w, _col_width(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_w, 50)


def _score_fill(value: int, max_score: int) -> PatternFill:
    if value == 0:
        return _SCORE_ZERO
    if value >= max_score:
        return _SCORE_FULL
    return _SCORE_PART


# ==========================================
# 2. 유튜브 분석 모듈
# ==========================================
def _make_ytt() -> YouTubeTranscriptApi:
    """쿠키 파일(*cookies*.txt)이 있으면 requests.Session에 로드해서 YouTubeTranscriptApi 생성."""
    import requests
    import http.cookiejar
    candidates = sorted(config.BASE_DIR.glob("*cookies*.txt"))
    if candidates:
        cookies_path = candidates[0]
        session = requests.Session()
        cj = http.cookiejar.MozillaCookieJar(str(cookies_path))
        cj.load(ignore_discard=True, ignore_expires=True)
        session.cookies = cj
        return YouTubeTranscriptApi(http_client=session)
    return YouTubeTranscriptApi()


# 모듈 로드 시 ytt 인스턴스 1회 생성 (쿠키 로그 중복 방지)
_ytt = _make_ytt()


def get_expert_picks_real(video_id: str) -> list[str]:
    # 자막 가져오기 — 429(rate-limit) 시 최대 3회 재시도
    transcript_list = None
    for attempt in range(3):
        try:
            transcript_list = _ytt.fetch(video_id, languages=['ko', 'ko-KR', 'en'])
            break
        except Exception as yt_err:
            if "IpBlocked" in type(yt_err).__name__:
                wait = 30 * (attempt + 1)
                console.print(
                    f"  [yellow]YouTube 요청 한도 초과 — {wait}초 대기 후 재시도 ({attempt+1}/3)...[/yellow]"
                )
                time.sleep(wait)
            else:
                console.print(f"  [yellow][경고] 자막 가져오기 실패: {yt_err}[/yellow]")
                return []
    if transcript_list is None:
        console.print("  [yellow][경고] 자막 재시도 한도 초과, 영상 건너뜀[/yellow]")
        return []

    try:
        full_text = " ".join([item.text for item in transcript_list])

        prompt = (
            "다음 주식 유튜브 자막에서 추천/긍정 언급된 종목명만 "
            "파이썬 리스트 형태로 뽑아줘. 예: ['삼성전자', 'LG에너지솔루션']\n\n"
            f"자막: {full_text[:15000]}"
        )

        # Gemini 429 시 최대 3회 재시도
        for attempt in range(3):
            try:
                response = _gemini.models.generate_content(
                    model="gemini-flash-lite-latest", contents=prompt
                )
                break
            except Exception as api_err:
                err_str = str(api_err)
                if "429" in err_str:
                    m = re.search(r'retry in (\d+)', err_str)
                    wait = int(m.group(1)) + 2 if m else 60
                    console.print(f"  [yellow]Gemini API 한도 초과 — {wait}초 대기 후 재시도 ({attempt+1}/3)...[/yellow]")
                    time.sleep(wait)
                else:
                    raise
        else:
            console.print("  [yellow][경고] Gemini API 재시도 한도 초과, 영상 건너뜀[/yellow]")
            return []

        match = re.search(r'\[.*?\]', response.text, re.DOTALL)
        if not match:
            return []
        return ast.literal_eval(match.group())
    except Exception as e:
        console.print(f"  [yellow][경고] 유튜브 분석 실패: {e}[/yellow]")
        return []


# ==========================================
# 3. 팩터 스코어링 모듈
# ==========================================
def calculate_momentum_score(ticker: str, expert_picks_mapping: dict) -> dict | None:
    """
    4개 팩터를 독립적으로 채점해 각 점수와 합계를 반환.
    반환 dict 키: 종목명, 코드, 현재가, AI추천, 외인수급, 기술반등, EPS흑자, 합계점수, 유튜브링크
    """
    try:
        name = stock.get_market_ticker_name(ticker)
        if not name:
            return None
    except Exception:
        return None

    score_youtube    = 0
    score_foreign    = 0
    score_technical  = 0
    score_eps        = 0
    current_price    = 0
    yt_links         = ""

    try:
        # 팩터 1: AI 유튜브 추천 (30점)
        if name in expert_picks_mapping:
            score_youtube = 30
            yt_links = ", ".join(list(expert_picks_mapping[name]))

        # 팩터 2: 외인 수급 유입 (30점) — 30일간 외국인 보유비율 증가
        df_investor = stock.get_exhaustion_rates_of_foreign_investment(start_date, end_date, ticker)
        if (not df_investor.empty
                and df_investor['상장주수대비외국인비율'].iloc[-1]
                > df_investor['상장주수대비외국인비율'].iloc[0]):
            score_foreign = 30

        # 팩터 3: 기술적 반등 (20점) — 현재가 > 5일 이동평균
        df_ohlcv = stock.get_market_ohlcv(start_date, end_date, ticker)
        if not df_ohlcv.empty and len(df_ohlcv) >= 5:
            current_price = df_ohlcv['종가'].iloc[-1]
            if current_price > df_ohlcv['종가'].rolling(window=5).mean().iloc[-1]:
                score_technical = 20

        # 팩터 4: EPS 흑자 (20점)
        df_fundamental = stock.get_market_fundamental(end_date, end_date, ticker)
        if not df_fundamental.empty and df_fundamental['EPS'].iloc[0] > 0:
            score_eps = 20

    except Exception:
        pass

    return {
        "종목명":       name,
        "코드":         ticker,
        "현재가":       current_price,
        "AI추천(30)":   score_youtube,
        "외인수급(30)": score_foreign,
        "기술반등(20)": score_technical,
        "EPS흑자(20)":  score_eps,
        "합계점수":     score_youtube + score_foreign + score_technical + score_eps,
        "유튜브링크":   yt_links,
    }


# ==========================================
# 4. 엑셀 생성
# ==========================================
def _build_excel(df: pd.DataFrame) -> None:
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    file_path = config.OUTPUT_DIR / f"멀티팩터_최종리포트_{datetime.now().strftime('%Y_%m%d_%H%M')}.xlsx"

    wb = Workbook()

    # ── 시트 1: 스크리닝 결과 ─────────────────────────────────────────────────
    ws = wb.active
    ws.title = "스크리닝 결과"
    ws.freeze_panes = "A2"

    headers = [
        "순위", "종목명", "코드", "현재가",
        "AI추천\n(만점30)", "외인수급\n(만점30)", "기술반등\n(만점20)", "EPS흑자\n(만점20)",
        "합계점수\n(만점100)",
        "유튜브링크",
    ]
    ws.append(headers)
    ws.row_dimensions[1].height = 30

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill      = _HEADER_FILL
        cell.font      = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    df_sorted = df.sort_values("합계점수", ascending=False).reset_index(drop=True)

    for rank, row in enumerate(df_sorted.itertuples(), 1):
        score_total = getattr(row, "합계점수", 0)
        row_data = [
            rank,
            row.종목명,
            row.코드,
            int(row.현재가) if row.현재가 else 0,
            getattr(row, "AI추천(30)",   0),
            getattr(row, "외인수급(30)", 0),
            getattr(row, "기술반등(20)", 0),
            getattr(row, "EPS흑자(20)",  0),
            score_total,
            row.유튜브링크,
        ]
        ws.append(row_data)
        r = rank + 1

        for col_idx in range(1, len(row_data) + 1):
            cell = ws.cell(row=r, column=col_idx)
            cell.font      = _BASE_FONT
            cell.alignment = _CENTER
            cell.border    = _THIN_BORDER

        # 종목명 → 네이버 금융 하이퍼링크
        name_cell = ws.cell(row=r, column=2)
        name_cell.hyperlink = f"https://finance.naver.com/item/main.naver?code={row.코드}"
        name_cell.font      = _LINK_FONT
        name_cell.alignment = _LEFT

        # 유튜브 링크 왼쪽 정렬
        ws.cell(row=r, column=10).alignment = _LEFT

        # 점수 컬럼 조건부 색상 (5~9번 컬럼)
        score_cols   = [5, 6, 7, 8]
        score_maxes  = [30, 30, 20, 20]
        for col_idx, max_val in zip(score_cols, score_maxes):
            cell = ws.cell(row=r, column=col_idx)
            cell.fill = _score_fill(int(cell.value or 0), max_val)

        # 합계 점수 컬럼 (9번) — 굵게
        total_cell = ws.cell(row=r, column=9)
        total_cell.font = Font(bold=True, size=10)
        total_cell.fill = _score_fill(int(score_total), 100)

        # 현재가 천단위 포맷
        ws.cell(row=r, column=4).number_format = "#,##0"

    _auto_width(ws)

    # ── 시트 2: 분석 조건 ─────────────────────────────────────────────────────
    ws2 = wb.create_sheet("분석 조건")

    cond_headers = ["항목", "값"]
    ws2.append(cond_headers)
    ws2.row_dimensions[1].height = 22
    for col_idx in range(1, 3):
        cell = ws2.cell(row=1, column=col_idx)
        cell.fill      = _HEADER_FILL
        cell.font      = _HEADER_FONT
        cell.alignment = _CENTER

    conditions = [
        ("실행일시",          datetime.now().strftime("%Y-%m-%d %H:%M")),
        ("검색 키워드",       config.SEARCH_KEYWORD),
        ("분석 영상 수",      config.MAX_VIDEOS),
        ("데이터 조회 기간",  f"{config.LOOKBACK_DAYS}일"),
        ("스캔 시장",         ", ".join(config.MARKETS)),
        ("총 분석 종목 수",   len(df)),
        ("합계점수 만점",     100),
        ("  AI추천 배점",     30),
        ("  외인수급 배점",   30),
        ("  기술반등 배점",   20),
        ("  EPS흑자 배점",    20),
        ("기술반등 기준",     "현재가 > 5일 이동평균"),
        ("외인수급 기준",     f"최근 {config.LOOKBACK_DAYS}일 외국인 보유비율 증가"),
    ]
    for r_idx, (k, v) in enumerate(conditions, 2):
        ws2.cell(row=r_idx, column=1, value=k).font = Font(bold=True, size=10)
        ws2.cell(row=r_idx, column=2, value=v).font = _BASE_FONT
        for col_idx in range(1, 3):
            ws2.cell(row=r_idx, column=col_idx).alignment = _LEFT
            ws2.cell(row=r_idx, column=col_idx).border    = _THIN_BORDER

    _auto_width(ws2)

    wb.save(file_path)
    console.print(f"\n  [green]리포트 저장 완료: {file_path}[/green]")


# ==========================================
# 5. 유튜브 자체 검색 엔진
# ==========================================
def search_recent_videos(keyword: str = "이권희 대표 주식") -> list[str]:
    console.print(f"  [cyan]유튜브 검색: '[bold]{keyword}[/bold]'[/cyan]")
    query_string = urllib.parse.urlencode({"search_query": keyword})
    url = "https://www.youtube.com/results?" + query_string
    try:
        req  = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
        html = urllib.request.urlopen(req).read().decode('utf-8')
        video_ids  = re.findall(r'"videoId":"([^"]{11})"', html)
        unique_ids = list(dict.fromkeys(video_ids))
        return unique_ids[:config.MAX_VIDEOS]
    except Exception as e:
        console.print(f"  [red][검색 에러] 유튜브 통신 실패: {e}[/red]")
        return []


# ==========================================
# 6. 메인 실행 파이프라인
# ==========================================
def run_scanner(video_id_list: list[str]) -> pd.DataFrame:
    # 종목명 → 링크 집합 매핑
    stock_to_links: dict[str, set] = {}

    # ── 영상 분석 ──────────────────────────────────────────────────────────────
    console.rule("[cyan]유튜브 영상 AI 분석[/cyan]")
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
        task = progress.add_task("영상 분석 중...", total=len(video_id_list))
        for i, vid in enumerate(video_id_list, 1):
            progress.update(task, description=f"영상 분석 중... [{i}/{len(video_id_list)}]")
            picks     = get_expert_picks_real(vid)
            video_url = f"https://www.youtube.com/watch?v={vid}"
            for stock_name in picks:
                stock_to_links.setdefault(stock_name, set()).add(video_url)
            progress.advance(task)
            time.sleep(config.VIDEO_DELAY)

    console.print(f"  [green]AI 추출 종목: {len(stock_to_links):,}개[/green]")
    console.print()

    # ── KRX 유니버스 로드 ──────────────────────────────────────────────────────
    console.rule("[cyan]KRX 유니버스 로드[/cyan]")
    all_tickers: list[str] = []
    try:
        market_counts = []
        for market in config.MARKETS:
            tickers = stock.get_market_ticker_list(end_date, market=market)
            all_tickers += tickers
            market_counts.append(f"{market} {len(tickers):,}개")
        console.print(
            f"  [green]유니버스 로드 완료 — {' + '.join(market_counts)} = 총 {len(all_tickers):,}개[/green]"
        )
    except Exception:
        pass

    if not all_tickers:
        console.print("  [yellow]야간 모드: 코어 종목으로 대체 분석합니다.[/yellow]")
        all_tickers = config.FALLBACK_TICKERS

    console.print()

    # ── 종목 팩터 채점 ─────────────────────────────────────────────────────────
    console.rule("[cyan]종목별 팩터 채점[/cyan]")
    results = []
    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=36, no_wrap=True)),
        BarColumn(bar_width=20),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
        refresh_per_second=2,
        transient=False,
    ) as progress:
        task = progress.add_task("팩터 채점 중...", total=len(all_tickers))
        for ticker in all_tickers:
            progress.update(task, description=f"팩터 채점 중... [{ticker}]")
            res = calculate_momentum_score(ticker, stock_to_links)
            if res:
                results.append(res)
            progress.advance(task)
            time.sleep(config.TICKER_DELAY)

    df_results = pd.DataFrame(results)
    _build_excel(df_results)
    return df_results


def _print_result_table(df: pd.DataFrame, elapsed: float) -> None:
    console.print()
    console.rule("[bold cyan]스캐닝 결과[/bold cyan]")
    console.print()

    if df is None or df.empty:
        console.print("  [yellow]합계점수가 있는 종목이 없습니다.[/yellow]")
        return

    df_top = df.sort_values("합계점수", ascending=False).head(20)

    table = Table(box=box.SIMPLE_HEAD, header_style="bold cyan", border_style="dim")
    table.add_column("순위",     width=4,  justify="right", style="dim")
    table.add_column("종목명",   width=16)
    table.add_column("코드",     width=8,  justify="center")
    table.add_column("현재가",   width=10, justify="right")
    table.add_column("AI추천",   width=8,  justify="right", style="cyan")
    table.add_column("외인수급", width=8,  justify="right", style="cyan")
    table.add_column("기술반등", width=8,  justify="right", style="cyan")
    table.add_column("EPS흑자",  width=8,  justify="right", style="cyan")
    table.add_column("합계점수", width=8,  justify="right", style="bold yellow")

    for rank, row in enumerate(df_top.itertuples(), 1):
        table.add_row(
            str(rank),
            row.종목명,
            row.코드,
            f"{int(row.현재가):,}" if row.현재가 else "-",
            str(getattr(row, "AI추천(30)",   0)),
            str(getattr(row, "외인수급(30)", 0)),
            str(getattr(row, "기술반등(20)", 0)),
            str(getattr(row, "EPS흑자(20)",  0)),
            str(getattr(row, "합계점수",     0)),
        )

    console.print(table)
    console.print(
        f"  분석 종목: [green]{len(df):,}개[/green]  "
        f"소요시간: [cyan]{elapsed:.0f}초[/cyan]"
    )
    console.print()


if __name__ == "__main__":
    _start = time.time()

    console.print()
    console.print(Panel.fit(
        "[bold cyan]멀티팩터 주식 스캐너[/bold cyan]\n"
        f"[dim]키워드: {config.SEARCH_KEYWORD}  |  "
        f"시장: {', '.join(config.MARKETS)}  |  "
        f"영상: {config.MAX_VIDEOS}개[/dim]",
        border_style="cyan",
    ))
    console.print()

    console.rule("[cyan]유튜브 영상 검색[/cyan]")
    auto_video_list = search_recent_videos(config.SEARCH_KEYWORD)

    if auto_video_list:
        console.print(f"  [green]검색된 영상: {len(auto_video_list)}개[/green]")
        console.print()
        df_final = run_scanner(auto_video_list)
        _print_result_table(df_final, time.time() - _start)
    else:
        console.print("\n  [yellow]분석할 영상이 없습니다. 종료합니다.[/yellow]")
