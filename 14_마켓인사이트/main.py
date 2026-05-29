# -*- coding: utf-8 -*-
# =============================================================================
# 마켓인사이트 — 글로벌·국내 시황 AI 브리핑
#
# [동작 순서]
# 1. yfinance     → 글로벌 지수·환율·원자재·미국채금리 수집 (병렬)
# 2. pykrx        → KOSPI·KOSDAQ·투자자 수급·업종 등락 수집 (병렬)
# 3. Gemini+검색  → 글로벌 주요 이슈 5건 (Google Search grounding, 병렬)
# 4. Gemini+검색  → 한국 주요 이슈 5건 (Google Search grounding, 병렬)
# 5. Gemini 합성  → 시장 방향성·주목 업종·투자 포인트 종합
# 6. Excel 저장   → 3개 시트 (글로벌지표 / 국내수급 / AI시황분석)
# 7. summary.txt  → 이메일 본문 저장 (run_all 연동용)
# 8. 이메일 발송  → Excel 첨부 + 시황 요약 본문
#
# [필요 패키지]
# pip install yfinance pykrx google-genai python-dotenv pandas openpyxl rich
#
# [환경 변수 (.env)]
# GEMINI_API_KEY=...
# GMAIL_USER=... / GMAIL_APP_PASSWORD=...
# NOTIFY_EMAIL=... / NOTIFY_EMAIL_WIFE=...
# NOTIFY_EMAILS_EXTRA=...
# SEND_EMAIL=true
# =============================================================================

import os
import re
import sys
import smtplib
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO
from pathlib import Path
from zoneinfo import ZoneInfo

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

# ── .env 로드 (pykrx import 전에 KRX_ID/KRX_PW 환경변수 확보) ──────────────────
try:
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).parent / ".env", override=False)
except ImportError:
    _ef = Path(__file__).parent / ".env"
    if _ef.exists():
        for _ln in _ef.read_text(encoding="utf-8").splitlines():
            _ln = _ln.strip()
            if _ln and not _ln.startswith("#") and "=" in _ln:
                _k, _, _v = _ln.partition("=")
                os.environ.setdefault(_k.strip(), _v.strip())

import pandas as pd
import yfinance as yf
from pykrx import stock as krx
from google import genai
from google.genai import types
from rich.console import Console

# ── 환경 변수 ────────────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GMAIL_USER     = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW   = os.getenv("GMAIL_APP_PASSWORD", "")
SEND_EMAIL     = os.getenv("SEND_EMAIL", "true").lower() in ("1", "true", "yes")
_main_emails: list[str] = [
    e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")]
    if e.strip()
]
_extra_emails: list[str] = [
    e.strip()
    for e in re.split(r"[,;\n]", os.getenv("NOTIFY_EMAILS_EXTRA", ""))
    if e.strip()
]
NOTIFY_EMAILS: list[str] = _main_emails + _extra_emails

# ── 상수 ─────────────────────────────────────────────────────────────────────────
KST     = ZoneInfo("Asia/Seoul")
HERE    = Path(__file__).parent
OUT_DIR = HERE / "output"
OUT_DIR.mkdir(exist_ok=True)

console = Console()
gemini  = genai.Client(api_key=GEMINI_API_KEY) if GEMINI_API_KEY else None

# ── 글로벌 티커 정의 ──────────────────────────────────────────────────────────────
_TICKERS: dict[str, list[tuple[str, str]]] = {
    "미국증시": [
        ("^GSPC",    "S&P500"),
        ("^IXIC",    "NASDAQ"),
        ("^DJI",     "다우존스"),
        ("^VIX",     "VIX (공포지수)"),
        ("^RUT",     "러셀2000"),
    ],
    "아시아증시": [
        ("^N225",    "닛케이225"),
        ("^HSI",     "항셍"),
        ("000001.SS","상하이종합"),
        ("^TWII",    "대만가권"),
    ],
    "환율": [
        ("USDKRW=X", "달러/원"),
        ("JPYKRW=X", "엔/원"),
        ("EURKRW=X", "유로/원"),
        ("DX-Y.NYB", "달러인덱스(DXY)"),
        ("USDCNY=X", "달러/위안"),
    ],
    "원자재": [
        ("CL=F",  "WTI 원유"),
        ("BZ=F",  "브렌트 원유"),
        ("GC=F",  "금"),
        ("SI=F",  "은"),
        ("HG=F",  "구리"),
    ],
    "미국채금리": [
        ("^IRX",  "미 3개월물 (%)"),
        ("^FVX",  "미 5년물 (%)"),
        ("^TNX",  "미 10년물 (%)"),
        ("^TYX",  "미 30년물 (%)"),
    ],
}


# ─────────────────────────────────────────────────────────────────────────────────
# 1. 글로벌 데이터 수집 (yfinance)
# ─────────────────────────────────────────────────────────────────────────────────

def fetch_global_data() -> pd.DataFrame:
    console.log("[cyan]글로벌 지표 수집 중...[/cyan]")
    rows: list[dict] = []
    all_items   = [(cat, tkr, name) for cat, items in _TICKERS.items() for tkr, name in items]
    all_symbols = [tkr for _, tkr, _ in all_items]

    # 개별 Ticker 병렬 다운로드 (multi-level DataFrame 파싱 이슈 회피)
    prices: dict[str, tuple[float, float]] = {}

    def _fetch_one(symbol: str):
        try:
            df = yf.Ticker(symbol).history(period="5d", interval="1d", auto_adjust=True)
            df = df.dropna(subset=["Close"])
            if len(df) >= 2:
                return symbol, (float(df["Close"].iloc[-1]), float(df["Close"].iloc[-2]))
            elif len(df) == 1:
                v = float(df["Close"].iloc[-1])
                return symbol, (v, v)
        except Exception:
            pass
        return symbol, (float("nan"), float("nan"))

    with ThreadPoolExecutor(max_workers=10) as ex:
        for sym, pair in ex.map(_fetch_one, all_symbols):
            prices[sym] = pair

    for cat, tkr, name in all_items:
        curr, prev = prices.get(tkr, (float("nan"), float("nan")))
        if not (curr != curr):  # not NaN
            chg = curr - prev
            pct = chg / prev * 100 if prev else 0.0
        else:
            chg = pct = float("nan")

        rows.append({
            "구분":      cat,
            "종목":      name,
            "현재가":    round(curr, 4) if curr == curr else "-",
            "전일대비":  round(chg,  4) if chg  == chg  else "-",
            "등락률(%)": round(pct,  2) if pct  == pct  else "-",
        })

    console.log(f"[green]글로벌 지표 {len(rows)}개 완료[/green]")
    return pd.DataFrame(rows)


# ─────────────────────────────────────────────────────────────────────────────────
# 1b. 차트용 6개월 일봉 데이터 수집
# ─────────────────────────────────────────────────────────────────────────────────

_CHART_TICKERS: list[tuple[str, str]] = [
    ("^GSPC",    "S&P 500"),
    ("^IXIC",    "NASDAQ"),
    ("^DJI",     "다우존스"),
    ("USDKRW=X", "달러/원 환율"),
    ("CL=F",     "WTI 원유"),
    ("^TNX",     "미 10년 국채금리(%)"),
    ("LIT",      "리튬 (LIT ETF)"),
    ("MU",       "DRAM (Micron·MU)"),
    ("^KS11",    "KOSPI"),
    ("^KQ11",    "KOSDAQ"),
]


def fetch_chart_data() -> dict[str, pd.Series]:
    """6개월 일봉 종가 데이터를 병렬로 수집."""
    console.log("[cyan]차트 데이터 수집 중 (6개월 일봉)...[/cyan]")
    result: dict[str, pd.Series] = {}

    def _fetch(item: tuple[str, str]) -> tuple[str, pd.Series]:
        sym, name = item
        try:
            df = yf.Ticker(sym).history(period="1y", interval="1d", auto_adjust=True)
            if not df.empty:
                return name, df["Close"].dropna()
        except Exception:
            pass
        return name, pd.Series(dtype=float)

    with ThreadPoolExecutor(max_workers=8) as ex:
        for name, series in ex.map(_fetch, _CHART_TICKERS):
            result[name] = series

    console.log(f"[green]차트 데이터 {len(result)}개 완료[/green]")
    return result


# ─────────────────────────────────────────────────────────────────────────────────
# 2. 국내 데이터 수집 (pykrx)
# ─────────────────────────────────────────────────────────────────────────────────

def _last_biz_days(n: int = 2) -> list[str]:
    """최근 n 영업일 날짜 문자열 리스트 (오래된 순). 주말 스킵."""
    days, d = [], datetime.now(KST).date()
    while len(days) < n:
        if d.weekday() < 5:
            days.append(d.strftime("%Y%m%d"))
        d -= timedelta(days=1)
    return list(reversed(days))


def fetch_korea_data() -> dict:
    console.log("[cyan]국내 수급 데이터 수집 중...[/cyan]")
    biz = _last_biz_days(2)
    prev_d, today_d = biz[0], biz[-1]
    result: dict = {"indices": pd.DataFrame(), "investor": pd.DataFrame(), "sectors": pd.DataFrame()}

    # ── KOSPI / KOSDAQ / KOSPI200 지수 ──────────────────────────────────────────
    idx_map = [("KOSPI", "1001"), ("KOSDAQ", "2001"), ("KOSPI200", "1028")]
    idx_rows: list[dict] = []
    for mkt, code in idx_map:
        try:
            df = krx.get_index_ohlcv_by_date(prev_d, today_d, code)
            if df is None or df.empty:
                continue
            df = df.dropna(subset=[df.columns[3]])  # 종가 컬럼
            if len(df) == 0:
                continue
            cur_row = df.iloc[-1]
            prv_row = df.iloc[-2] if len(df) >= 2 else cur_row
            close_col = [c for c in df.columns if "종가" in c or "Close" in c]
            vol_col   = [c for c in df.columns if "거래대금" in c or "Value" in c]
            c_close = float(cur_row[close_col[0]]) if close_col else float("nan")
            p_close = float(prv_row[close_col[0]]) if close_col else float("nan")
            chg = c_close - p_close
            pct = chg / p_close * 100 if p_close else 0.0
            vol = float(cur_row[vol_col[0]]) if vol_col else 0
            idx_rows.append({
                "지수":          mkt,
                "종가":          round(c_close, 2),
                "전일대비":      round(chg, 2),
                "등락률(%)":     round(pct, 2),
                "거래대금(억원)": round(vol / 1e8, 0) if vol else "-",
            })
        except Exception as e:
            console.log(f"[yellow]{mkt} 조회 실패: {e}[/yellow]")
    result["indices"] = pd.DataFrame(idx_rows)

    # ── 투자자별 순매수 ──────────────────────────────────────────────────────────
    try:
        df_inv = krx.get_market_trading_value_by_investor(today_d, today_d, "KOSPI")
        if df_inv is not None and not df_inv.empty:
            result["investor"] = df_inv
    except Exception as e:
        console.log(f"[yellow]투자자 수급 조회 실패: {e}[/yellow]")

    # ── 업종별 등락 (KOSPI) ──────────────────────────────────────────────────────
    try:
        sector_tkrs = krx.get_index_ticker_list(market="KOSPI")
        sec_rows: list[dict] = []

        def _fetch_sector(tkr: str):
            try:
                name = krx.get_index_ticker_name(tkr)
                # 주요 지수 코드 제외 (전체·200·ETF 등)
                skip_kw = ["코스피200", "KOSPI200", "코스닥150", "KRX", "배당", "레버리지", "인버스", "F-"]
                if any(k in name for k in skip_kw) or len(name) < 2:
                    return None
                df_s = krx.get_index_ohlcv_by_date(prev_d, today_d, tkr)
                if df_s is None or len(df_s) < 2:
                    return None
                close_col = [c for c in df_s.columns if "종가" in c or "Close" in c]
                if not close_col:
                    return None
                c = float(df_s.iloc[-1][close_col[0]])
                p = float(df_s.iloc[-2][close_col[0]])
                pct = (c - p) / p * 100 if p else 0.0
                return {"업종": name, "등락률(%)": round(pct, 2)}
            except Exception:
                return None

        with ThreadPoolExecutor(max_workers=8) as ex:
            for r in ex.map(_fetch_sector, sector_tkrs):
                if r:
                    sec_rows.append(r)

        if sec_rows:
            result["sectors"] = (
                pd.DataFrame(sec_rows)
                .sort_values("등락률(%)", ascending=False)
                .reset_index(drop=True)
            )
    except Exception as e:
        console.log(f"[yellow]업종 조회 실패: {e}[/yellow]")

    console.log("[green]국내 수급 데이터 수집 완료[/green]")
    return result


# ─────────────────────────────────────────────────────────────────────────────────
# 3 & 4. Gemini 검색 grounding
# ─────────────────────────────────────────────────────────────────────────────────

def _gemini_search(prompt: str, label: str, _retries: int = 2, _delay: int = 15) -> str:
    if not gemini:
        return "(GEMINI_API_KEY 미설정)"
    import time as _time
    for attempt in range(1 + _retries):
        try:
            if attempt > 0:
                console.log(f"[yellow]Gemini 검색 재시도 {attempt}/{_retries}: {label}...[/yellow]")
            else:
                console.log(f"[cyan]Gemini 검색: {label}...[/cyan]")
            resp = gemini.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(
                    tools=[types.Tool(google_search=types.GoogleSearch())],
                    temperature=0.1,
                ),
            )
            text = resp.text or ""
            console.log(f"[green]Gemini 검색 완료: {label}[/green]")
            return text.strip()
        except Exception as e:
            err_str = str(e)
            console.log(f"[red]Gemini 검색 오류({label}): {err_str[:120]}[/red]")
            if attempt < _retries:
                console.log(f"[dim]{_delay}초 후 재시도...[/dim]")
                _time.sleep(_delay)
    return f"※ {label} 조회 실패 (Gemini API 일시 오류 — 잠시 후 재시도 바람)"


def fetch_global_news(today_str: str) -> str:
    prompt = (
        f"오늘({today_str}) 기준 글로벌 주요 증시·경제 이슈 5가지를 최신 뉴스에서 찾아 한국어로 정리해줘.\n\n"
        "각 이슈는 다음 형식으로:\n"
        "1. [이슈 제목]\n"
        "   - 내용: 2~3문장 핵심 요약\n"
        "   - 시장 영향: 긍정/부정/중립 + 영향 자산·섹터\n\n"
        "이슈는 미국 연준·금리, 지정학 리스크, 글로벌 경기, 원자재, 기술주 동향 등에서 선별."
    )
    return _gemini_search(prompt, "글로벌 이슈")


def fetch_korea_news(today_str: str) -> str:
    prompt = (
        f"오늘({today_str}) 기준 한국 증시 주요 이슈 5가지를 최신 뉴스에서 찾아 정리해줘.\n\n"
        "각 이슈는 다음 형식으로:\n"
        "1. [이슈 제목]\n"
        "   - 내용: 2~3문장 핵심 요약\n"
        "   - 관련 섹터/종목: 영향받는 업종 또는 대표 종목\n\n"
        "이슈는 외국인·기관 수급, 환율·금리, 정부 정책, 실적 발표, 업종 이슈 등에서 선별."
    )
    return _gemini_search(prompt, "한국 이슈")


# ─────────────────────────────────────────────────────────────────────────────────
# 5. Gemini 종합 합성
# ─────────────────────────────────────────────────────────────────────────────────

def synthesize(df_global: pd.DataFrame, korea: dict, global_news: str, korea_news: str) -> str:
    if not gemini:
        return "(GEMINI_API_KEY 미설정)"
    console.log("[cyan]Gemini 종합 분석 중...[/cyan]")

    # 정량 데이터 → 텍스트 요약
    lines = ["=== 글로벌 정량 데이터 ==="]
    for cat, grp in df_global.groupby("구분", sort=False):
        lines.append(f"\n[{cat}]")
        for _, row in grp.iterrows():
            pct = f"{row['등락률(%)']}%" if row["등락률(%)"] != "-" else ""
            lines.append(f"  {row['종목']}: {row['현재가']} ({pct})")

    if not korea["indices"].empty:
        lines.append("\n=== 국내 지수 ===")
        for _, row in korea["indices"].iterrows():
            lines.append(f"  {row['지수']}: {row['종가']} ({row['등락률(%)']}%)")

    if not korea["sectors"].empty:
        lines.append("\n=== 업종 상위 5 ===")
        for _, r in korea["sectors"].head(5).iterrows():
            lines.append(f"  {r['업종']}: {r['등락률(%)']}%")
        lines.append("=== 업종 하위 5 ===")
        for _, r in korea["sectors"].tail(5).iterrows():
            lines.append(f"  {r['업종']}: {r['등락률(%)']}%")

    quant = "\n".join(lines)

    prompt = f"""당신은 한국 주식시장 전문 애널리스트입니다.
아래 정량 데이터와 오늘의 주요 뉴스를 종합하여 시황 분석을 작성하세요.

{quant}

=== 글로벌 주요 이슈 ===
{global_news}

=== 한국 주요 이슈 ===
{korea_news}

다음 4개 섹션으로 작성하세요:

[섹션A] 시장 방향성 (리스크 온/오프)
- 오늘 시장의 전반적 분위기·방향
- 핵심 변수 2~3가지

[섹션B] 주목 업종·테마 Top5
- 오늘 주목할 업종/테마와 이유
- 관련 대표 종목

[섹션C] 투자 포인트 및 전략
- 단기(1~3일) 투자 포인트
- 주의할 리스크

[섹션D] 다음 주요 이벤트 예고
- 이번 주~다음 주 주요 이벤트 (FOMC, CPI, 실적 등)
- 예상 시장 영향"""

    import time as _time
    for attempt in range(3):
        try:
            if attempt > 0:
                console.log(f"[yellow]Gemini 합성 재시도 {attempt}/2...[/yellow]")
            resp = gemini.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt,
                config=types.GenerateContentConfig(temperature=0.2),
            )
            text = resp.text or ""
            console.log("[green]Gemini 종합 분석 완료[/green]")
            return text.strip()
        except Exception as e:
            console.log(f"[red]Gemini 합성 오류: {str(e)[:120]}[/red]")
            if attempt < 2:
                console.log("[dim]15초 후 재시도...[/dim]")
                _time.sleep(15)
    return "※ AI 종합 분석 실패 (Gemini API 일시 오류 — 잠시 후 재시도 바람)"


# ─────────────────────────────────────────────────────────────────────────────────
# 6. Excel 저장
# ─────────────────────────────────────────────────────────────────────────────────

def save_excel(
    df_global: pd.DataFrame,
    korea: dict,
    global_news: str,
    korea_news: str,
    synthesis: str,
    ts: str,
) -> Path:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    fname = OUT_DIR / f"마켓인사이트_{datetime.now().strftime('%Y_%m%d_%H%M')}.xlsx"
    wb = openpyxl.Workbook()

    # ── 공통 스타일 ──────────────────────────────────────────────────────────────
    HDR_FILL    = PatternFill("solid", fgColor="1F4E79")
    HDR_FONT    = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=10)
    BODY_FONT   = Font(name="맑은 고딕", size=10)
    TITLE_FONT  = Font(name="맑은 고딕", bold=True, size=12, color="1F4E79")
    SEC_FONT    = Font(name="맑은 고딕", bold=True, size=11, color="1F4E79")
    CENTER      = Alignment(horizontal="center", vertical="center", wrap_text=True)
    LEFT        = Alignment(horizontal="left",   vertical="center", wrap_text=True)
    THIN        = Border(
        left=Side(style="thin"), right=Side(style="thin"),
        top=Side(style="thin"), bottom=Side(style="thin"),
    )
    GREEN_FILL  = PatternFill("solid", fgColor="C6EFCE")
    RED_FILL    = PatternFill("solid", fgColor="FFC7CE")
    SEC_FILL    = PatternFill("solid", fgColor="D6E4F0")

    def _hdr(ws, row: int, cols: list, widths: list | None = None):
        for c, col in enumerate(cols, 1):
            cell = ws.cell(row=row, column=c, value=col)
            cell.fill, cell.font, cell.alignment, cell.border = HDR_FILL, HDR_FONT, CENTER, THIN
        if widths:
            for c, w in enumerate(widths, 1):
                ws.column_dimensions[get_column_letter(c)].width = w

    def _row(ws, row: int, vals: list, fills: list | None = None):
        for c, val in enumerate(vals, 1):
            cell = ws.cell(row=row, column=c, value=val)
            cell.font, cell.alignment, cell.border = BODY_FONT, CENTER, THIN
            if fills and c - 1 < len(fills) and fills[c - 1]:
                cell.fill = fills[c - 1]

    def _pct_fill(pct):
        try:
            v = float(pct)
            return GREEN_FILL if v > 0 else (RED_FILL if v < 0 else None)
        except Exception:
            return None

    # ════════════════════════════════════════════════════════════════════════════
    # Sheet 1: 글로벌지표
    # ════════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "글로벌지표"
    ws1.row_dimensions[1].height = 28
    t = ws1.cell(1, 1, f"글로벌 주요 지표  ·  {ts}")
    t.font, t.alignment = TITLE_FONT, CENTER
    ws1.merge_cells("A1:E1")

    _hdr(ws1, 2, ["구분", "종목", "현재가", "전일대비", "등락률(%)"],
         widths=[14, 20, 14, 14, 12])

    r = 3
    prev_cat = None
    for _, row in df_global.iterrows():
        cat_val = row["구분"] if row["구분"] != prev_cat else ""
        _row(ws1, r,
             [cat_val, row["종목"], row["현재가"], row["전일대비"], row["등락률(%)"]],
             [None, None, None, None, _pct_fill(row["등락률(%)"])])
        prev_cat = row["구분"]
        r += 1

    ws1.freeze_panes = "A3"

    # ════════════════════════════════════════════════════════════════════════════
    # Sheet 2: 국내수급
    # ════════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("국내수급")
    ws2.column_dimensions["A"].width = 20
    for col in ["B", "C", "D", "E", "F", "G", "H"]:
        ws2.column_dimensions[col].width = 14

    r = 1
    t = ws2.cell(r, 1, f"국내 수급 현황  ·  {ts}")
    t.font, t.alignment = TITLE_FONT, CENTER
    ws2.merge_cells(f"A{r}:F{r}")
    ws2.row_dimensions[r].height = 28
    r += 2

    # 지수 섹션
    ws2.cell(r, 1, "▌ 주요 지수").font = SEC_FONT
    r += 1
    if not korea["indices"].empty:
        cols = list(korea["indices"].columns)
        _hdr(ws2, r, cols)
        r += 1
        for _, row in korea["indices"].iterrows():
            pct_col = next((i for i, c in enumerate(cols) if "등락률" in c), None)
            fills = [None] * len(cols)
            if pct_col is not None:
                fills[pct_col] = _pct_fill(row.iloc[pct_col])
            _row(ws2, r, list(row), fills)
            r += 1
    r += 1

    # 투자자별 순매수
    ws2.cell(r, 1, "▌ 투자자별 순매수 (KOSPI, 억원)").font = SEC_FONT
    r += 1
    if not korea["investor"].empty:
        df_inv = korea["investor"]
        inv_cols = ["구분"] + list(df_inv.columns)[:7]
        _hdr(ws2, r, inv_cols)
        r += 1
        for idx_label, row_data in df_inv.iterrows():
            vals = [str(idx_label)] + [
                round(row_data.get(c, 0) / 1e8, 0) if isinstance(row_data.get(c), (int, float)) else row_data.get(c, "-")
                for c in list(df_inv.columns)[:7]
            ]
            _row(ws2, r, vals)
            r += 1
    else:
        ws2.cell(r, 1, "(데이터 없음)").font = BODY_FONT
        r += 1
    r += 1

    # 업종별 등락
    ws2.cell(r, 1, "▌ 업종별 등락률 (KOSPI)").font = SEC_FONT
    r += 1
    if not korea["sectors"].empty:
        _hdr(ws2, r, ["업종", "등락률(%)"], widths=[22, 12])
        r += 1
        for _, row in korea["sectors"].iterrows():
            _row(ws2, r, [row["업종"], row["등락률(%)"]],
                 [None, _pct_fill(row["등락률(%)"])])
            r += 1
    else:
        ws2.cell(r, 1, "(데이터 없음)").font = BODY_FONT
        r += 1

    # ════════════════════════════════════════════════════════════════════════════
    # Sheet 3: AI시황분석
    # ════════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("AI시황분석")
    ws3.column_dimensions["A"].width = 130

    r = 1
    t = ws3.cell(r, 1, f"AI 시황 분석  ·  {ts}")
    t.font, t.alignment = TITLE_FONT, LEFT
    ws3.row_dimensions[r].height = 28
    r += 2

    def _section(title: str, content: str):
        nonlocal r
        c = ws3.cell(r, 1, title)
        c.font, c.fill, c.alignment = SEC_FONT, SEC_FILL, LEFT
        ws3.row_dimensions[r].height = 22
        r += 1
        for line in content.strip().splitlines():
            cell = ws3.cell(r, 1, line)
            cell.font, cell.alignment = BODY_FONT, LEFT
            ws3.row_dimensions[r].height = 16
            r += 1
        r += 1

    _section("📰 글로벌 주요 이슈 (Gemini 검색)", global_news)
    _section("🇰🇷 한국 주요 이슈 (Gemini 검색)", korea_news)
    _section("🤖 AI 종합 시황 분석", synthesis)

    wb.save(fname)
    console.log(f"[green]Excel 저장 완료: {fname.name}[/green]")
    return fname


# ─────────────────────────────────────────────────────────────────────────────────
# 7. PDF 저장
# ─────────────────────────────────────────────────────────────────────────────────

def save_pdf(
    df_global: pd.DataFrame,
    korea: dict,
    global_news: str,
    korea_news: str,
    synthesis: str,
    ts: str,
    chart_data: dict | None = None,
) -> Path | None:
    try:
        from reportlab.lib import colors
        from reportlab.lib.enums import TA_CENTER, TA_LEFT
        from reportlab.lib.pagesizes import A4
        from reportlab.lib.styles import ParagraphStyle
        from reportlab.lib.units import cm
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        from reportlab.platypus import (
            HRFlowable, Image as RLImage, PageBreak, Paragraph,
            SimpleDocTemplate, Spacer, Table, TableStyle,
        )
    except ImportError:
        console.log("[yellow]reportlab 미설치 — PDF 생략 (pip install reportlab)[/yellow]")
        return None

    fname = OUT_DIR / f"마켓인사이트_{datetime.now().strftime('%Y_%m%d_%H%M')}.pdf"

    # ── 한글 폰트 등록 ──────────────────────────────────────────────────────────
    _WIN_FONT  = Path("C:/Windows/Fonts/malgun.ttf")
    _WIN_FONTB = Path("C:/Windows/Fonts/malgunbd.ttf")
    if _WIN_FONT.exists():
        try:
            pdfmetrics.registerFont(TTFont("Malgun",     str(_WIN_FONT)))
            pdfmetrics.registerFont(TTFont("MalgunBold", str(_WIN_FONTB)))
        except Exception:
            pass
        FONT, FONT_B = "Malgun", "MalgunBold"
    else:
        FONT, FONT_B = "Helvetica", "Helvetica-Bold"

    # ── 스타일 정의 ─────────────────────────────────────────────────────────────
    NAVY  = colors.HexColor("#1F4E79")
    BLUE  = colors.HexColor("#2E75B6")
    GREEN = colors.HexColor("#C6EFCE")
    RED   = colors.HexColor("#FFC7CE")
    ALT   = colors.HexColor("#F2F7FF")

    def S(name, **kw):
        kw.setdefault("fontName", FONT)
        return ParagraphStyle(name, **kw)

    sty = {
        "title":  S("title",  fontName=FONT_B, fontSize=22, leading=30, spaceAfter=8,  alignment=TA_CENTER, textColor=NAVY),
        "sub":    S("sub",    fontSize=11,     leading=16,  spaceAfter=4,  alignment=TA_CENTER, textColor=colors.grey),
        "h1":     S("h1",     fontName=FONT_B, fontSize=13,   spaceBefore=10, spaceAfter=4, textColor=NAVY),
        "h2":     S("h2",     fontName=FONT_B, fontSize=10.5, spaceBefore=6,  spaceAfter=3, textColor=BLUE),
        "body":   S("body",   fontSize=9.5,    leading=15,    spaceAfter=1),
        "disc":   S("disc",   fontSize=8,      leading=12,    textColor=colors.grey),
        "notice": S("notice", fontName=FONT_B, fontSize=9,    textColor=colors.HexColor("#7F0000")),
    }

    W = A4[0] - 4 * cm  # 사용 가능한 가로 폭

    # ── 테이블 공통 스타일 생성기 ────────────────────────────────────────────────
    def base_ts(n_rows: int):
        return [
            ("BACKGROUND", (0, 0), (-1, 0),      NAVY),
            ("TEXTCOLOR",  (0, 0), (-1, 0),      colors.white),
            ("FONTNAME",   (0, 0), (-1, 0),      FONT_B),
            ("FONTNAME",   (0, 1), (-1, n_rows), FONT),
            ("FONTSIZE",   (0, 0), (-1, -1),     8.5),
            ("ALIGN",      (0, 0), (-1, -1),     "CENTER"),
            ("VALIGN",     (0, 0), (-1, -1),     "MIDDLE"),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, ALT]),
            ("GRID",       (0, 0), (-1, -1),     0.3, colors.HexColor("#CCCCCC")),
            ("ROWHEIGHT",  (0, 0), (-1, -1),     16),
        ]

    def pct_color(val):
        try:
            return GREEN if float(val) > 0 else (RED if float(val) < 0 else None)
        except Exception:
            return None

    # ── Gemini 텍스트 → PDF 단락 변환 ───────────────────────────────────────────
    def text_to_paras(text: str) -> list:
        paras = []
        for line in text.strip().splitlines():
            line = line.strip()
            if not line:
                paras.append(Spacer(1, 4))
                continue
            # 섹션 제목 감지 ([섹션A] 등)
            is_sec = line.startswith("[섹션") or (line.startswith("[") and line.endswith("]"))
            is_num = len(line) > 1 and line[0].isdigit() and line[1] in "."
            style  = sty["h2"] if is_sec else (sty["notice"] if is_num else sty["body"])
            # reportlab 특수문자 이스케이프
            safe = line.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            paras.append(Paragraph(safe, style))
        return paras

    # ── 문서 구성 ────────────────────────────────────────────────────────────────
    doc   = SimpleDocTemplate(str(fname), pagesize=A4,
                               leftMargin=2*cm, rightMargin=2*cm,
                               topMargin=2*cm,  bottomMargin=2*cm)
    story = []

    # ── 1페이지 상단 간략 요약 헬퍼 ────────────────────────────────────────────
    def _brief_lines(text: str, max_items: int = 2) -> list[str]:
        """뉴스 본문에서 주요 이슈 제목 추출 (번호·별표 제거)."""
        items: list[str] = []
        for ln in text.splitlines():
            s = ln.strip()
            m = re.match(r'^\d+\.\s+\*{0,2}(.+?)\*{0,2}\s*$', s)
            if m:
                title = m.group(1).strip().rstrip(":")
                if len(title) > 4:
                    items.append(title)
                    if len(items) >= max_items:
                        break
        if not items:
            for ln in text.splitlines():
                s = ln.strip()
                if s and not s.startswith(("(", "※", "-", "·", "#", "[")):
                    items.append(s[:75])
                    if len(items) >= max_items:
                        break
        return items

    def _brief_table() -> "Table | None":
        g_lines = _brief_lines(global_news, max_items=4)
        k_lines = _brief_lines(korea_news,  max_items=4)
        if not g_lines and not k_lines:
            return None

        LT_BLUE = colors.HexColor("#EBF5FB")
        LT_GRN  = colors.HexColor("#EAF4EA")
        BORDER  = colors.HexColor("#2E75B6")

        def _safe(s: str) -> str:
            return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        rows_data, row_colors = [], []
        for i, ln in enumerate(g_lines):
            lbl = "🌍 글로벌" if i == 0 else ""
            rows_data.append([Paragraph(lbl, ParagraphStyle("bl", fontName=FONT_B, fontSize=8, textColor=NAVY)),
                               Paragraph(f"· {_safe(ln)}", ParagraphStyle("bc", fontName=FONT, fontSize=8, leading=12))])
            row_colors.append(LT_BLUE)
        for i, ln in enumerate(k_lines):
            lbl = "🇰🇷 국내" if i == 0 else ""
            rows_data.append([Paragraph(lbl, ParagraphStyle("kl", fontName=FONT_B, fontSize=8, textColor=colors.HexColor("#375623"))),
                               Paragraph(f"· {_safe(ln)}", ParagraphStyle("kc", fontName=FONT, fontSize=8, leading=12))])
            row_colors.append(LT_GRN)

        tbl = Table(rows_data, colWidths=[W * 0.15, W * 0.85])
        ts_b: list = [
            ("BOX",          (0, 0), (-1, -1), 0.7, BORDER),
            ("INNERGRID",    (0, 0), (-1, -1), 0.3, colors.HexColor("#CCDDEE")),
            ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING",  (0, 0), (-1, -1), 7),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING",   (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 4),
        ]
        for ri, bg in enumerate(row_colors):
            ts_b.append(("BACKGROUND", (0, ri), (-1, ri), bg))
        tbl.setStyle(TableStyle(ts_b))
        return tbl

    # ── KOSPI/KOSDAQ 1년 차트 (1페이지 하단용 compact 2-panel) ───────────────────
    def _build_kr_index_chart_png() -> bytes | None:
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import matplotlib.dates as mdates
        except ImportError:
            return None

        kospi_s  = (chart_data or {}).get("KOSPI",  pd.Series(dtype=float))
        kosdaq_s = (chart_data or {}).get("KOSDAQ", pd.Series(dtype=float))
        if kospi_s.empty and kosdaq_s.empty:
            return None

        try:
            matplotlib.rcParams["font.family"] = "Malgun Gothic"
        except Exception:
            pass
        matplotlib.rcParams["axes.unicode_minus"] = False

        pairs = [("KOSPI", kospi_s, "#1F4E79"), ("KOSDAQ", kosdaq_s, "#C00000")]
        fig, axes = plt.subplots(1, 2, figsize=(8.5, 2.6))
        fig.patch.set_facecolor("#F8F9FA")

        for ax, (name, series, color) in zip(axes, pairs):
            if series.empty:
                ax.text(0.5, 0.5, "데이터 없음", ha="center", va="center",
                        transform=ax.transAxes, fontsize=9)
                ax.set_title(name, fontsize=9, fontweight="bold")
                continue
            y_lo = series.min() * 0.97
            ax.fill_between(series.index, y_lo, series.values, alpha=0.12, color=color)
            ax.plot(series.index, series.values, color=color, linewidth=1.3, zorder=3)
            ax.set_ylim(bottom=y_lo)
            ax.set_title(f"{name} (1년)", fontsize=9, fontweight="bold",
                         color="#1F4E79", pad=3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
            ax.xaxis.set_major_locator(mdates.MonthLocator())
            for lbl in ax.get_xticklabels():
                lbl.set_rotation(30)
                lbl.set_ha("right")
            ax.tick_params(axis="both", labelsize=7, length=2)
            ax.set_facecolor("#FFFFFF")
            ax.grid(True, linestyle="--", alpha=0.45, linewidth=0.5, zorder=0)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)

            last_val  = series.iloc[-1]
            first_val = series.iloc[0]
            pct = (last_val - first_val) / first_val * 100 if first_val else 0
            sign = "+" if pct >= 0 else ""
            ax.annotate(
                f"{last_val:,.2f}  ({sign}{pct:.1f}%)",
                xy=(series.index[-1], last_val),
                xytext=(-5, 6), textcoords="offset points",
                fontsize=6.5,
                color="#C00000" if pct < 0 else "#375623",
                fontweight="bold", ha="right",
            )

        plt.tight_layout()
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight",
                    facecolor=fig.get_facecolor())
        plt.close(fig)
        buf.seek(0)
        return buf.read()

    # ── 1페이지 수급 미니 테이블 (지수 등락률 + 외인/기관/개인 순매수) ──────────────
    def _kr_supply_mini_table() -> "Table | None":
        idx_df = korea.get("indices",  pd.DataFrame())
        inv_df = korea.get("investor", pd.DataFrame())
        if idx_df.empty and inv_df.empty:
            return None

        LT_BLUE = colors.HexColor("#EBF5FB")
        LT_GRN  = colors.HexColor("#EAF4EA")
        LT_RED  = colors.HexColor("#FDECEA")
        GRAY_H  = colors.HexColor("#1F4E79")

        def _safe(s) -> str:
            return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

        def _p(txt, bold=False, color=None, size=8):
            fs = FONT_B if bold else FONT
            kw = {"fontName": fs, "fontSize": size, "leading": 12}
            if color:
                kw["textColor"] = color
            return Paragraph(_safe(txt), ParagraphStyle("_mp", **kw))

        # ── 지수 행 (KOSPI / KOSDAQ) ───────────────────────────
        idx_cells: list[tuple] = []
        for _, row in idx_df.iterrows():
            name = str(row.get("지수", ""))
            if name not in ("KOSPI", "KOSDAQ"):
                continue
            close = row.get("종가", "-")
            pct   = row.get("등락률(%)", 0)
            try:
                pct_f = float(pct)
            except Exception:
                pct_f = 0.0
            sign  = "+" if pct_f >= 0 else ""
            c_txt = colors.HexColor("#C00000") if pct_f < 0 else colors.HexColor("#375623")
            idx_cells.append((name, f"{close:,.2f}" if isinstance(close, float) else str(close),
                               f"{sign}{pct}%", pct_f))

        # ── 수급 행 (외국인 / 기관 / 개인) ─────────────────────
        net_col = None
        inv_rows: list[tuple] = []
        if not inv_df.empty:
            net_col = next((c for c in inv_df.columns if "순매수" in c), None)
            if net_col is None:
                net_col = inv_df.columns[-1] if len(inv_df.columns) else None
        if net_col:
            for label in ("외국인합계", "기관합계", "개인"):
                if label in inv_df.index:
                    raw = inv_df.loc[label, net_col]
                    val = round(raw / 1e8, 0) if isinstance(raw, (int, float)) else 0
                    inv_rows.append((label.replace("합계", ""), val))

        if not idx_cells and not inv_rows:
            return None

        # 테이블 열: 구분 | 지수1 종가·등락 | 지수2 종가·등락 | 외인 | 기관 | 개인
        n_idx = len(idx_cells)
        n_inv = len(inv_rows)
        header_cells = [_p("구분", bold=True)]
        for name, close, pct_s, _ in idx_cells:
            header_cells.append(_p(name, bold=True))
        for label, _ in inv_rows:
            header_cells.append(_p(label, bold=True))

        data_row_index = [_p("종가 / 등락률")]
        for name, close, pct_s, pct_f in idx_cells:
            c_txt = colors.HexColor("#C00000") if pct_f < 0 else colors.HexColor("#375623")
            data_row_index.append(_p(f"{close}  {pct_s}", color=c_txt, bold=True))
        for label, val in inv_rows:
            data_row_index.append(_p("순매수(억)"))

        data_row_values = [_p("")]
        for _ in idx_cells:
            data_row_values.append(_p(""))
        for label, val in inv_rows:
            sign  = "+" if val >= 0 else ""
            c_txt = colors.HexColor("#375623") if val >= 0 else colors.HexColor("#C00000")
            data_row_values.append(_p(f"{sign}{int(val):,}", bold=True, color=c_txt))

        # 단순 2행 레이아웃 대신 1행으로 통합
        n_cols = 1 + n_idx + n_inv
        col_w  = [W * 0.18] + [W * 0.22] * n_idx + [W * (0.60 / max(n_inv, 1))] * n_inv
        # n_idx와 n_inv에 따라 width 재배분
        if n_idx + n_inv > 0:
            rest = W * 0.82
            idx_w  = rest * 0.5 / max(n_idx, 1) if n_idx else 0
            inv_w  = rest * 0.5 / max(n_inv, 1) if n_inv else 0
            col_w  = [W * 0.18] + [idx_w] * n_idx + [inv_w] * n_inv

        combined_row = [_p("금일 현황", bold=True)]
        for name, close, pct_s, pct_f in idx_cells:
            c_txt = colors.HexColor("#C00000") if pct_f < 0 else colors.HexColor("#375623")
            combined_row.append(_p(f"{close}\n{pct_s}", bold=True, color=c_txt, size=8))
        for label, val in inv_rows:
            sign  = "+" if val >= 0 else ""
            c_txt = colors.HexColor("#375623") if val >= 0 else colors.HexColor("#C00000")
            combined_row.append(_p(f"순매수\n{sign}{int(val):,}억", bold=True, color=c_txt, size=8))

        tbl = Table([header_cells, combined_row], colWidths=col_w)
        ts_s: list = [
            ("BACKGROUND",  (0, 0), (-1, 0),  GRAY_H),
            ("TEXTCOLOR",   (0, 0), (-1, 0),  colors.white),
            ("FONTNAME",    (0, 0), (-1, 0),  FONT_B),
            ("BACKGROUND",  (0, 1), (-1, 1),  colors.white),
            ("FONTSIZE",    (0, 0), (-1, -1), 8),
            ("ALIGN",       (0, 0), (-1, -1), "CENTER"),
            ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
            ("GRID",        (0, 0), (-1, -1), 0.3, colors.HexColor("#CCCCCC")),
            ("ROWHEIGHT",   (0, 0), (-1, -1), 22),
            ("TOPPADDING",  (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        # 지수 열 배경
        for ci in range(1, 1 + n_idx):
            tbl_bg = LT_BLUE
            ts_s.append(("BACKGROUND", (ci, 1), (ci, 1), tbl_bg))
        # 수급 열 배경
        for ci in range(1 + n_idx, n_cols):
            ts_s.append(("BACKGROUND", (ci, 1), (ci, 1), LT_GRN))
        tbl.setStyle(TableStyle(ts_s))
        return tbl

    # ▌ 표지 + 글로벌 지표 (같은 페이지)
    story += [
        Spacer(1, 0.6*cm),
        Paragraph("AI 마켓인사이트", sty["title"]),
        Paragraph("글로벌 · 국내 시황 브리핑", sty["sub"]),
        Paragraph(ts, sty["sub"]),
        Spacer(1, 0.25*cm),
    ]

    brief = _brief_table()
    if brief:
        story += [brief, Spacer(1, 0.3*cm)]

    # ── KOSPI/KOSDAQ 1년 차트 ──────────────────────────────────────────────────
    if chart_data:
        kr_chart_png = _build_kr_index_chart_png()
        if kr_chart_png:
            kr_img = RLImage(BytesIO(kr_chart_png), width=W, height=W * 0.32)
            story += [kr_img, Spacer(1, 0.15*cm)]

    # ── 금일 지수 등락률 + 외인/기관/개인 수급 ────────────────────────────────────
    supply_tbl = _kr_supply_mini_table()
    if supply_tbl:
        story += [supply_tbl, Spacer(1, 0.25*cm)]

    story += [
        HRFlowable(width="100%", thickness=2, color=NAVY),
        Spacer(1, 0.2*cm),
        Paragraph(
            "본 자료는 투자 참고 목적으로만 제공되며, 특정 종목에 대한 매수·매도 권유가 아닙니다. "
            "투자 판단 및 그에 따른 손익은 전적으로 본인에게 있습니다. "
            "본 자료는 수신자 개인에게만 제공된 것으로, 외부 재배포를 금합니다.",
            sty["disc"],
        ),
        Spacer(1, 0.3*cm),
    ]

    # ▌ 1. 글로벌 지표
    story.append(Paragraph("1. 글로벌 주요 지표", sty["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=BLUE))
    story.append(Spacer(1, 0.3*cm))

    cw = [W*0.16, W*0.26, W*0.18, W*0.18, W*0.22]
    rows = [["구분", "종목", "현재가", "전일대비", "등락률(%)"]]
    ts_list = base_ts(len(df_global))
    prev_cat = None
    for i, (_, row) in enumerate(df_global.iterrows(), 1):
        cat = row["구분"] if row["구분"] != prev_cat else ""
        rows.append([cat, row["종목"], str(row["현재가"]), str(row["전일대비"]), str(row["등락률(%)"])])
        f = pct_color(row["등락률(%)"])
        if f:
            ts_list.append(("BACKGROUND", (4, i), (4, i), f))
        prev_cat = row["구분"]
    t = Table(rows, colWidths=cw)
    t.setStyle(TableStyle(ts_list))
    story += [t, PageBreak()]

    # ▌ 차트 페이지 (6개월 일봉)
    def _build_chart_png() -> bytes | None:
        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            import matplotlib.dates as mdates
        except ImportError:
            return None

        try:
            matplotlib.rcParams["font.family"] = "Malgun Gothic"
        except Exception:
            pass
        matplotlib.rcParams["axes.unicode_minus"] = False

        # KOSPI/KOSDAQ는 1페이지 전용 차트에서 별도 표시
        items   = [(k, v) for k, v in (chart_data or {}).items()
                   if k not in ("KOSPI", "KOSDAQ")]
        COLORS  = ["#2E75B6", "#C00000", "#70AD47", "#ED7D31", "#7030A0", "#1F4E79"]

        fig, axes = plt.subplots(4, 2, figsize=(8.5, 11.5))
        fig.patch.set_facecolor("#F8F9FA")
        fig.suptitle("주요 지표 1년 차트 (일봉)", fontsize=11,
                     fontweight="bold", color="#1F4E79", y=0.998)

        for idx, ax in enumerate(axes.flat):
            if idx >= len(items):
                ax.set_visible(False)
                continue
            name, series = items[idx]
            color = COLORS[idx % len(COLORS)]

            if series.empty:
                ax.text(0.5, 0.5, "데이터 없음", ha="center", va="center",
                        transform=ax.transAxes, fontsize=9)
                ax.set_title(name, fontsize=9, fontweight="bold", color="#1F4E79")
                continue

            y_lo = series.min() * 0.97
            ax.fill_between(series.index, y_lo, series.values, alpha=0.12, color=color)
            ax.plot(series.index, series.values, color=color, linewidth=1.3, zorder=3)
            ax.set_ylim(bottom=y_lo)
            ax.set_title(name, fontsize=9, fontweight="bold", color="#1F4E79", pad=3)
            ax.xaxis.set_major_formatter(mdates.DateFormatter("%m/%d"))
            ax.xaxis.set_major_locator(mdates.MonthLocator())
            for lbl in ax.get_xticklabels():
                lbl.set_rotation(30)
                lbl.set_ha("right")
            ax.tick_params(axis="both", labelsize=7, length=2)
            ax.set_facecolor("#FFFFFF")
            ax.grid(True, linestyle="--", alpha=0.45, linewidth=0.5, zorder=0)
            ax.spines["top"].set_visible(False)
            ax.spines["right"].set_visible(False)

            last_val  = series.iloc[-1]
            first_val = series.iloc[0]
            pct = (last_val - first_val) / first_val * 100 if first_val else 0
            sign = "+" if pct >= 0 else ""
            lbl_color = "#C00000" if pct < 0 else "#375623"
            ax.annotate(
                f"{last_val:,.2f}  ({sign}{pct:.1f}%)",
                xy=(series.index[-1], last_val),
                xytext=(-5, 6), textcoords="offset points",
                fontsize=6.5, color=lbl_color, fontweight="bold", ha="right",
            )

        plt.tight_layout(rect=[0, 0, 1, 0.995])
        buf = BytesIO()
        fig.savefig(buf, format="png", dpi=120, bbox_inches="tight",
                    facecolor=fig.get_facecolor())
        plt.close(fig)
        buf.seek(0)
        return buf.read()

    if chart_data:
        chart_png = _build_chart_png()
        if chart_png:
            img = RLImage(BytesIO(chart_png), width=W, height=W * 1.35)
            story.append(Paragraph("주요 지표 1년 차트 (일봉)", sty["h1"]))
            story.append(HRFlowable(width="100%", thickness=1, color=BLUE))
            story.append(Spacer(1, 0.25 * cm))
            story += [img, PageBreak()]

    # ▌ 2. 국내 수급
    story.append(Paragraph("2. 국내 수급 현황", sty["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=BLUE))
    story.append(Spacer(1, 0.3*cm))

    if not korea["indices"].empty:
        story.append(Paragraph("주요 지수", sty["h2"]))
        df_idx = korea["indices"]
        idx_cols = list(df_idx.columns)
        idx_rows = [idx_cols] + [list(r) for r in df_idx.itertuples(index=False)]
        idx_ts = base_ts(len(df_idx))
        pct_col = next((i for i, c in enumerate(idx_cols) if "등락률" in c), None)
        for ri, rv in enumerate(df_idx.itertuples(index=False), 1):
            if pct_col is not None:
                f = pct_color(list(rv)[pct_col])
                if f:
                    idx_ts.append(("BACKGROUND", (pct_col, ri), (pct_col, ri), f))
        t2 = Table(idx_rows, colWidths=[W / len(idx_cols)] * len(idx_cols))
        t2.setStyle(TableStyle(idx_ts))
        story += [t2, Spacer(1, 0.5*cm)]
    else:
        story.append(Paragraph("(국내 시장 데이터 없음 — 주말·공휴일)", sty["body"]))
        story.append(Spacer(1, 0.3*cm))

    if not korea["sectors"].empty:
        story.append(Paragraph("업종별 등락률 (KOSPI)", sty["h2"]))
        df_sec = korea["sectors"]
        sec_rows = [["업종", "등락률(%)"]] + [[r["업종"], r["등락률(%)"]] for _, r in df_sec.iterrows()]
        sec_ts = base_ts(len(df_sec))
        for ri, (_, rv) in enumerate(df_sec.iterrows(), 1):
            f = pct_color(rv["등락률(%)"])
            if f:
                sec_ts.append(("BACKGROUND", (1, ri), (1, ri), f))
        sec_ts.append(("ALIGN", (0, 0), (0, -1), "LEFT"))
        t3 = Table(sec_rows, colWidths=[W * 0.75, W * 0.25])
        t3.setStyle(TableStyle(sec_ts))
        story.append(t3)

    story.append(PageBreak())

    # ▌ 3. AI 시황 분석
    story.append(Paragraph("3. AI 시황 분석", sty["h1"]))
    story.append(HRFlowable(width="100%", thickness=1, color=BLUE))
    story.append(Spacer(1, 0.3*cm))

    story.append(Paragraph("글로벌 주요 이슈", sty["h2"]))
    story += text_to_paras(global_news)
    story.append(Spacer(1, 0.5*cm))

    story.append(Paragraph("한국 주요 이슈", sty["h2"]))
    story += text_to_paras(korea_news)
    story.append(PageBreak())

    story.append(Paragraph("AI 종합 시황 분석", sty["h2"]))
    story += text_to_paras(synthesis)

    doc.build(story)
    console.log(f"[green]PDF 저장 완료: {fname.name}[/green]")
    return fname


# ─────────────────────────────────────────────────────────────────────────────────
# 8 & 9. 이메일 발송 (summary.txt 항상 저장)
# ─────────────────────────────────────────────────────────────────────────────────

def _send_email(
    output_file: str,
    pdf_file: str | None,
    global_news: str,
    korea_news: str,
    synthesis: str,
) -> None:
    ts = datetime.now(KST).strftime("%Y-%m-%d %H:%M")

    # ── 플레인 텍스트 (HTML 미지원 클라이언트 fallback) ───────────────────────
    body_text = "\n".join([
        "=" * 60,
        "■ 주의사항 및 면책",
        "=" * 60,
        "  · 본 자료는 투자 참고 목적으로만 제공되며,",
        "    특정 종목에 대한 매수·매도 권유가 아닙니다.",
        "  · 투자 판단 및 그에 따른 손익은 전적으로 본인에게 있습니다.",
        "  · 본 자료는 수신자 개인에게만 제공된 것으로,",
        "    SNS·커뮤니티·단체방 등 외부 재배포를 금합니다.",
        "  · 무단 재배포로 인한 법적 책임은 재배포한 당사자에게 있습니다.",
        "=" * 60,
        "",
        "안녕하세요,",
        "",
        f"[{ts}] 마켓인사이트 시황 분석이 완료되었습니다.",
        "",
        "=" * 60,
        "■ 글로벌 주요 이슈",
        "=" * 60,
        global_news,
        "",
        "=" * 60,
        "■ 한국 주요 이슈",
        "=" * 60,
        korea_news,
        "",
        "=" * 60,
        "■ AI 종합 시황 분석",
        "=" * 60,
        synthesis,
        "",
        "상세 분석 결과는 첨부 Excel/PDF 파일을 확인해주세요.",
        "",
        "※ 이 메일은 자동 발송입니다.",
    ])

    # ── HTML 본문 ─────────────────────────────────────────────────────────────
    def _to_html(text: str) -> str:
        parts = []
        for line in text.strip().splitlines():
            line = line.strip()
            if not line:
                parts.append("<br>")
                continue
            safe = (line.replace("&", "&amp;")
                        .replace("<", "&lt;")
                        .replace(">", "&gt;"))
            if safe.startswith("[섹션") or (safe.startswith("[") and safe.endswith("]")):
                parts.append(
                    f'<p style="font-weight:bold;color:#2E75B6;margin:10px 0 3px 0;">{safe}</p>'
                )
            elif len(safe) > 1 and safe[0].isdigit() and safe[1] in ".)":
                parts.append(
                    f'<p style="margin:4px 0 4px 6px;"><b>{safe[:2]}</b>{safe[2:]}</p>'
                )
            elif safe[:1] in ("·", "•", "-", "▶"):
                parts.append(
                    f'<p style="margin:3px 0 3px 14px;">{safe}</p>'
                )
            else:
                parts.append(f'<p style="margin:4px 0;">{safe}</p>')
        return "\n".join(parts)

    def _section(icon: str, title: str, content: str, color: str) -> str:
        return f"""
  <div style="margin:20px 0;">
    <h3 style="background:{color};color:white;padding:10px 16px;
               border-radius:4px 4px 0 0;margin:0;font-size:14px;">{icon} {title}</h3>
    <div style="border:1px solid {color};border-top:none;padding:14px 16px;
                border-radius:0 0 4px 4px;background:#fafcff;
                line-height:1.75;font-size:13px;">
      {_to_html(content)}
    </div>
  </div>"""

    body_html = f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8"></head>
<body style="font-family:'Malgun Gothic',Arial,sans-serif;font-size:14px;
             color:#222;max-width:720px;margin:0 auto;padding:20px;">

  <!-- 헤더 -->
  <div style="background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);
              padding:24px 28px;border-radius:8px;margin-bottom:20px;">
    <div style="color:white;font-size:22px;font-weight:bold;letter-spacing:1px;">
      📊 AI 마켓인사이트
    </div>
    <div style="color:#BDD7EE;font-size:13px;margin-top:6px;">
      글로벌 · 국내 시황 브리핑 &nbsp;·&nbsp; {ts}
    </div>
  </div>

  <p style="margin:0 0 4px 0;">안녕하세요,</p>
  <p style="margin:0 0 16px 0;">오늘의 마켓인사이트 시황 분석이 완료되었습니다.</p>

  {_section("🌍", "글로벌 주요 이슈", global_news, "#2E75B6")}
  {_section("🇰🇷", "한국 주요 이슈", korea_news, "#375623")}
  {_section("🤖", "AI 종합 시황 분석", synthesis, "#1F4E79")}

  <!-- 첨부 안내 -->
  <div style="background:#EBF5FB;border-left:4px solid #2E75B6;padding:14px 18px;
              border-radius:4px;margin:20px 0;font-size:13px;">
    <strong>📎 첨부 파일 안내</strong>
    <ul style="margin:8px 0 0 0;padding-left:20px;line-height:1.9;">
      <li><b>[글로벌지표]</b> — S&amp;P500 · NASDAQ · VIX · 환율 · 원자재 · 미국채금리 전일 대비</li>
      <li><b>[국내수급]</b> — KOSPI · KOSDAQ · 투자자별 순매수 · 업종별 등락률</li>
      <li><b>[AI시황분석]</b> — 글로벌 · 국내 주요 이슈 + AI 종합 분석</li>
      <li><b>PDF</b> — 6개월 일봉 차트 포함 전체 브리핑 문서</li>
    </ul>
  </div>

  <!-- 면책 -->
  <div style="background:#F8F8F8;border:1px solid #DDD;border-radius:4px;
              padding:12px 16px;font-size:12px;color:#777;margin-top:20px;">
    · 본 자료는 투자 참고 목적으로만 제공되며, 특정 종목에 대한 매수·매도 권유가 아닙니다.<br>
    · 투자 판단 및 그에 따른 손익은 전적으로 본인에게 있습니다.<br>
    · 본 자료는 수신자 개인에게만 제공된 것으로, SNS·커뮤니티·단체방 등 외부 재배포를 금합니다.<br>
    · 무단 재배포로 인한 법적 책임은 재배포한 당사자에게 있습니다.<br><br>
    ※ 이 메일은 자동 발송입니다.
  </div>

</body>
</html>"""

    try:
        (Path(output_file).parent / "summary.txt").write_text(body_text, encoding="utf-8")
    except Exception:
        pass

    if not SEND_EMAIL:
        console.print("[dim]이메일 발송 생략 (SEND_EMAIL=false)[/dim]")
        return
    if not (GMAIL_USER and GMAIL_APP_PW):
        console.print("[yellow]발송 계정 미설정 (.env: GMAIL_USER / GMAIL_APP_PASSWORD)[/yellow]")
        return
    if not NOTIFY_EMAILS:
        console.print("[yellow]수신자 미설정 (.env: NOTIFY_EMAIL)[/yellow]")
        return

    # multipart/mixed → alternative(text+html) + 파일 첨부
    msg = MIMEMultipart("mixed")
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(_main_emails) if _main_emails else GMAIL_USER
    if _extra_emails:
        msg["Bcc"] = ", ".join(_extra_emails)
    msg["Subject"] = f"[마켓인사이트] {ts} — 글로벌·국내 시황 브리핑"

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(body_text, "plain", "utf-8"))
    alt.attach(MIMEText(body_html,  "html",  "utf-8"))
    msg.attach(alt)

    for attach_path in [output_file, pdf_file]:
        if not attach_path:
            continue
        p = Path(attach_path)
        if p.exists():
            with open(p, "rb") as f:
                part = MIMEApplication(f.read(), Name=p.name)
                part["Content-Disposition"] = f'attachment; filename="{p.name}"'
                msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PW)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAILS, msg.as_string())
        to_str  = ", ".join(_main_emails) if _main_emails else GMAIL_USER
        bcc_str = f" | Bcc {len(_extra_emails)}명" if _extra_emails else ""
        console.print(f"[green]이메일 발송 완료 → {to_str}{bcc_str}[/green]")
    except Exception as e:
        console.print(f"[red]이메일 발송 실패: {e}[/red]")


# ─────────────────────────────────────────────────────────────────────────────────
# 메인
# ─────────────────────────────────────────────────────────────────────────────────

def main() -> None:
    _start = datetime.now()
    ts     = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    today  = datetime.now(KST).strftime("%Y년 %m월 %d일")

    console.rule("[bold cyan]마켓인사이트[/bold cyan]")
    console.print(f"[dim]{ts}  글로벌·국내 시황 수집 시작[/dim]\n")

    # ── 5개 병렬 수집 ──────────────────────────────────────────────────────────
    with ThreadPoolExecutor(max_workers=5) as ex:
        f_global = ex.submit(fetch_global_data)
        f_charts = ex.submit(fetch_chart_data)
        f_korea  = ex.submit(fetch_korea_data)
        f_gnews  = ex.submit(fetch_global_news, today)
        f_knews  = ex.submit(fetch_korea_news,  today)

        df_global   = f_global.result()
        chart_data  = f_charts.result()
        korea_data  = f_korea.result()
        global_news = f_gnews.result()
        korea_news  = f_knews.result()

    # ── 합성 ──────────────────────────────────────────────────────────────────
    synthesis = synthesize(df_global, korea_data, global_news, korea_news)

    # ── Excel 저장 ────────────────────────────────────────────────────────────
    output_file = save_excel(df_global, korea_data, global_news, korea_news, synthesis, ts)

    # ── PDF 저장 ──────────────────────────────────────────────────────────────
    pdf_file = save_pdf(df_global, korea_data, global_news, korea_news, synthesis, ts,
                        chart_data=chart_data)

    # ── 이메일 발송 / summary.txt ─────────────────────────────────────────────
    _send_email(str(output_file), str(pdf_file) if pdf_file else None,
                global_news, korea_news, synthesis)

    elapsed = (datetime.now() - _start).total_seconds()
    m, s = divmod(int(elapsed), 60)
    console.print(f"\n[bold green]⏱ 완료  {m}분 {s}초[/bold green]")
    console.rule()


if __name__ == "__main__":
    main()
