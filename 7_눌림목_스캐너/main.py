"""
박스권 돌파 눌림목 스캐너  v3
=================================
조건 필터:
  일봉  60일선  횡보 또는 우상향
  주봉  20주선  우상향
  월봉   5월선  우상향

차트 출력:
  종목당 일봉 / 주봉 / 월봉 3개 차트 PNG 자동 저장

데이터:
  pykrx (기본)  또는  KRX OpenAPI (.env의 KRX_AUTH_KEY)

설치:
  pip install pykrx pandas numpy matplotlib mplfinance python-dotenv tabulate

사용:
  python scanner_v3.py                          # 전체 스캔
  python scanner_v3.py --market KOSPI           # KOSPI만
  python scanner_v3.py --chart-only 011170      # 특정 종목 차트만
  python scanner_v3.py --vol-mult 4 --top-n 15  # 조건 조정
"""

import argparse
import os
import re
import socket
import sys
import warnings
import time
from concurrent.futures import ThreadPoolExecutor, wait, FIRST_COMPLETED
from datetime import datetime, timedelta
from pathlib import Path

socket.setdefaulttimeout(15)   # pykrx HTTP 요청 최대 15초

import numpy as np
import pandas as pd
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import matplotlib.ticker as mticker
from rich.console import Console
from rich.progress import (
    BarColumn, MofNCompleteColumn, Progress,
    SpinnerColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn,
)
from rich.table import Column

warnings.filterwarnings("ignore")

console = Console()

# ── 출력 디렉토리 ─────────────────────────────────────────────────────────────
OUTPUT_DIR = Path(__file__).parent / "output"

# ── Excel 컬럼 순서 / 포맷 / 헤더 색상 (4_financial_scanner 스타일) ──────────────────
COLUMN_ORDER = [
    "종목명", "티커", "점수", "유형",
    "현재가", "돌파일", "돌파가", "고점", "눌림(%)",
    "박스상단", "박스하단", "박스범위(%)",
    "거래량배율", "시총(억)",
    "52주저점대비(%)",
    "일봉MA60기울기(%)", "주봉MA20기울기(%)", "월봉MA5기울기(%)",
]

_FMT_SLOPE    = '[Color10]+#,##0.00;[Red]-#,##0.00;0.00'
_FMT_PCT_SIGN = '[Color10]+#,##0.0;[Red]-#,##0.0;0.0'

_COL_FORMATS: dict[str, str] = {
    # ── 눌림목 컬럼 ──
    "점수":               "#,##0",
    "현재가":             "#,##0",
    "돌파가":             "#,##0",
    "고점":               "#,##0",
    "눌림(%)":            "#,##0.1",
    "박스상단":           "#,##0",
    "박스하단":           "#,##0",
    "박스범위(%)":        "#,##0.1",
    "거래량배율":         "#,##0.0",
    "시총(억)":           "#,##0",
    "52주저점대비(%)":    "#,##0.1",
    "일봉MA60기울기(%)":  _FMT_SLOPE,
    "주봉MA20기울기(%)":  _FMT_SLOPE,
    "월봉MA5기울기(%)":   _FMT_SLOPE,
    # ── 실적개선 컬럼 ──
    "종합점수":           "#,##0.0",
    "적정주가":           "#,##0",
    "상승여력1(%)":        _FMT_PCT_SIGN,
    "증권사목표주가":      "#,##0",
    "상승여력2(%)":        _FMT_PCT_SIGN,
    "25년매출(억)":       "#,##0",
    "26년매출E(억)":      "#,##0",
    "매출증가율(%)":      _FMT_PCT_SIGN,
    "25년영업익(억)":     "#,##0",
    "26년영업익E(억)":    "#,##0",
    "영업익증가율(%)":    _FMT_PCT_SIGN,
    "PER":                "#,##0.0",
    "12M PER":            "#,##0.0",
    "12M PER/PER":        "#,##0.00",
    "PSR":                "#,##0.00",
    "PBR":                "#,##0.00",
}

_HEADER_FILLS: dict[str, str] = {
    # ── 눌림목 컬럼 ──
    "종목명":             "1F4E79",
    "티커":               "1F4E79",
    "점수":               "7B3F00",
    "유형":               "7B3F00",
    "현재가":             "375623",
    "돌파일":             "375623",
    "돌파가":             "375623",
    "고점":               "375623",
    "눌림(%)":            "375623",
    "박스상단":           "1F3864",
    "박스하단":           "1F3864",
    "박스범위(%)":        "1F3864",
    "거래량배율":         "7F4F24",
    "시총(억)":           "4C4C4C",
    "52주저점대비(%)":    "5C3317",
    "일봉MA60기울기(%)":  "2E4057",
    "주봉MA20기울기(%)":  "2E4057",
    "월봉MA5기울기(%)":   "2E4057",
    # ── 실적개선 컬럼 ──
    "업종":               "1F4E79",
    "종합점수":           "7B3F00",
    "적정주가":           "375623",
    "상승여력1(%)":        "375623",
    "증권사목표주가":      "1D6A3A",
    "상승여력2(%)":        "1D6A3A",
    "25년매출(억)":       "7F4F24",
    "26년매출E(억)":      "7F4F24",
    "매출증가율(%)":      "7F4F24",
    "25년영업익(억)":     "1F3864",
    "26년영업익E(억)":    "1F3864",
    "영업익증가율(%)":    "1F3864",
    "PER":                "4C4C4C",
    "12M PER":            "2E4057",
    "12M PER/PER":        "2E4057",
    "PSR":                "4C4C4C",
    "PBR":                "4C4C4C",
}

# ── Excel 서식 적용 ───────────────────────────────────────────────────────────

def _format_excel(path: Path) -> None:
    """저장된 Excel 파일에 섹션별 헤더·컬럼별 숫자 포맷·하이퍼링크·열 너비 적용."""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font, PatternFill

        wb = load_workbook(path)

        def _cw(val):
            return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

        for ws in wb.worksheets:
            col_fmt_map: dict[int, str] = {}
            name_col = ticker_col = type_col = None

            for cell in ws[1]:
                col_name = cell.value or ""
                fill_hex = _HEADER_FILLS.get(col_name, "1F4E79")
                cell.fill      = PatternFill("solid", fgColor=fill_hex)
                cell.font      = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", wrap_text=True)
                if col_name in _COL_FORMATS:
                    col_fmt_map[cell.column] = _COL_FORMATS[col_name]
                if col_name == "종목명": name_col   = cell.column
                if col_name == "티커":   ticker_col = cell.column
                if col_name == "유형":   type_col   = cell.column

            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center")
                    fmt = col_fmt_map.get(cell.column)
                    if fmt and isinstance(cell.value, (int, float)):
                        cell.number_format = fmt
                    elif isinstance(cell.value, int):
                        cell.number_format = "#,##0"

            # 유형 컬럼 조건부 색상
            if type_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    cell = row[type_col - 1]
                    if cell.value == "바닥탈출":
                        cell.font = Font(bold=True, color="FF6B35")   # 주황
                        cell.fill = PatternFill("solid", fgColor="3D1A00")
                    elif cell.value == "우상향":
                        cell.font = Font(bold=True, color="3FB950")   # 초록
                        cell.fill = PatternFill("solid", fgColor="0D2818")

            if name_col and ticker_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    nc = row[name_col   - 1]
                    tc = row[ticker_col - 1]
                    if tc.value:
                        code6 = str(tc.value).zfill(6)
                        nc.hyperlink = (
                            f"https://finance.naver.com/item/main.naver?code={code6}"
                        )
                        nc.font = Font(color="0563C1", underline="single", bold=True)
                        tc.hyperlink = (
                            f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                            f"?pGB=1&gicode=A{code6}&cID=AA"
                            f"&MenuYn=Y&ReportGB=&NewMenuID=11&stkGb=&strResearchYN="
                        )
                        tc.font = Font(color="0563C1", underline="single")

            for col in ws.columns:
                header_w = _cw(col[0].value) + 3
                data_w   = max((_cw(cell.value) for cell in col[1:]), default=0) + 2
                ws.column_dimensions[col[0].column_letter].width = min(
                    max(header_w, data_w, 10), 40
                )

            ws.freeze_panes    = "A2"
            ws.auto_filter.ref = ws.dimensions

        wb.save(path)
    except Exception as e:
        console.print(f"[yellow]Excel 서식 적용 실패: {e}[/yellow]")


_COL_RENAME = {
    "ticker":        "티커",
    "name":          "종목명",
    "score":         "점수",
    "current_price": "현재가",
    "break_date":    "돌파일",
    "break_price":   "돌파가",
    "peak_price":    "고점",
    "pullback_pct":  "눌림(%)",
    "box_top":       "박스상단",
    "box_bottom":    "박스하단",
    "range_ratio":   "박스범위(%)",
    "vol_ratio":     "거래량배율",
    "cap_億":        "시총(억)",
    "from_52w_low":  "52주저점대비(%)",
    "pattern":       "유형",
    "d_slope":       "일봉MA60기울기(%)",
    "w_slope":       "주봉MA20기울기(%)",
    "m_slope":       "월봉MA5기울기(%)",
}


def _get_실적개선_df(df_눌림목: pd.DataFrame) -> pd.DataFrame:
    """눌림목 전체 종목의 실적개선 데이터 수집.
    1) 실적개선 output에서 left join
    2) 매칭 안 된 종목은 FnGuide에서 직접 수집
    """
    output_dir = Path(__file__).parent.parent / "4_financial_scanner" / "output"
    files = sorted(output_dir.glob("실적개선주_*_hl.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        files = sorted(output_dir.glob("실적개선주_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)

    base = df_눌림목[["종목명", "티커"]].copy()
    base["티커"] = base["티커"].astype(str).str.zfill(6)

    # 1. 실적개선 output left join
    if files:
        try:
            df_실적 = pd.read_excel(files[0], sheet_name="종합")
            df_실적["티커"] = df_실적["티커"].astype(str).str.zfill(6)
            extra_cols = [c for c in df_실적.columns if c not in ("종목명", "티커")]
            merged = base.merge(df_실적[["티커"] + extra_cols], on="티커", how="left")
        except Exception as e:
            console.print(f"[yellow]실적개선 파일 로드 실패: {e}[/yellow]")
            merged = base.copy()
    else:
        merged = base.copy()

    # 2. 매칭 안 된 종목 FnGuide 직접 수집
    score_col = "종합점수"
    missing_mask = merged[score_col].isna() if score_col in merged.columns else pd.Series(True, index=merged.index)
    missing_tickers = merged.loc[missing_mask, "티커"].tolist()

    if missing_tickers:
        console.print(f"  [dim]기본 밸류에이션 수집: {len(missing_tickers)}개 종목...[/dim]")
        cap_map = df_눌림목.set_index(
            df_눌림목["티커"].astype(str).str.zfill(6)
        )["시총(억)"].to_dict() if "시총(억)" in df_눌림목.columns else {}

        for ticker in missing_tickers:
            data = _fetch_partial_밸류(ticker, cap_map.get(ticker, 0))
            if not data:
                continue
            idx = merged.index[merged["티커"] == ticker][0]
            for col, val in data.items():
                if col not in merged.columns:
                    merged[col] = None
                merged.at[idx, col] = val

    return merged


def _fetch_partial_밸류(ticker: str, cap_억: float = 0.0) -> dict:
    """컨센서스가 없는 종목용 — FnGuide+Naver에서 기본 밸류 데이터만 수집."""
    import re
    import requests
    from bs4 import BeautifulSoup

    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"}

    try:
        res = requests.get(
            "https://comp.fnguide.com/SVO2/asp/SVD_Main.asp",
            params={"pGB": "1", "gicode": f"A{ticker}", "cID": "",
                    "MenuYn": "Y", "ReportGB": "D", "NewMenuID": "Y", "stkGb": "701"},
            headers=headers, timeout=10,
        )
        soup = BeautifulSoup(res.text, "lxml")
    except Exception:
        return {}

    def _parse(text: str):
        text = text.strip().replace(",", "")
        try:
            return float(re.sub(r"[^\d.\-]", "", text))
        except (ValueError, TypeError):
            return None

    result: dict = {}

    for aid, key in [("h_per", "PER"), ("h_12m", "12M PER"), ("h_pbr", "PBR")]:
        a = soup.find("a", id=aid)
        if a:
            dd = a.find_next_sibling("dd")
            if dd:
                v = _parse(dd.get_text(strip=True))
                if v is not None:
                    result[key] = v

    if "PER" in result and "12M PER" in result and result["PER"]:
        result["12M PER/PER"] = round(result["12M PER"] / result["PER"], 2)

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]
        if "목표주가" in header and "투자의견" in header:
            data_row = [c.get_text(strip=True) for c in rows[1].find_all(["th", "td"])]
            try:
                v = _parse(data_row[header.index("목표주가")])
                if v:
                    result["증권사목표주가"] = v
            except (IndexError, ValueError):
                pass
            break

    # 업종 정보
    stxt = soup.select_one("p.stxt_group")
    if stxt:
        import re as _re
        txt = stxt.get_text(" ", strip=True)
        m = _re.search(r"\|\s*FICS\s+(.+?)\s*\|", txt)
        if m:
            result["업종"] = m.group(1).strip()

    try:
        res2 = requests.get(
            f"https://m.stock.naver.com/api/stock/{ticker}/basic",
            headers={**headers, "Referer": "https://m.stock.naver.com/"},
            timeout=10,
        )
        if res2.status_code == 200:
            price_str = res2.json().get("closePrice", "").replace(",", "")
            if price_str:
                result["현재가"] = float(price_str)
    except Exception:
        pass

    if result.get("현재가") and result.get("증권사목표주가"):
        result["상승여력2(%)"] = round(
            (result["증권사목표주가"] - result["현재가"]) / result["현재가"] * 100, 1
        )

    return result


def _save_excel(df_result: pd.DataFrame) -> Path:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    ts        = datetime.now().strftime("%Y_%m%d_%H%M")
    file_name = OUTPUT_DIR / f"눌림목스캐너_{ts}.xlsx"

    df      = df_result.rename(columns=_COL_RENAME)
    ordered = [c for c in COLUMN_ORDER if c in df.columns]
    df      = df[ordered]

    df_실적 = _get_실적개선_df(df)

    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="눌림목", index=False)
        df_실적.to_excel(writer, sheet_name="실적개선", index=False)
        matched = df_실적["종합점수"].notna().sum() if "종합점수" in df_실적.columns else 0
        console.print(f"  [dim]실적개선 시트: 전체 {len(df_실적)}개 종목 (실적 데이터 {matched}개 매칭)[/dim]")

    _format_excel(file_name)
    return file_name


def _send_email(file_name: Path, df: pd.DataFrame) -> None:
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    if not (GMAIL_USER and GMAIL_APP_PW):
        console.print("[yellow]발송 계정 미설정 (.env: GMAIL_USER / GMAIL_APP_PASSWORD)[/yellow]")
        return
    if not NOTIFY_EMAILS:
        console.print("[yellow]수신자 미설정 (.env: NOTIFY_EMAIL / NOTIFY_EMAIL_WIFE)[/yellow]")
        return

    ts      = datetime.now().strftime("%Y-%m-%d %H:%M")
    n_stock = len(df)

    top10      = df.head(10) if not df.empty else df
    top_lines  = []
    for rank, (_, r) in enumerate(top10.iterrows(), 1):
        name = r.get("name", r.get("종목명", ""))
        pb   = r.get("pullback_pct", r.get("눌림(%)", 0))
        vr   = r.get("vol_ratio",    r.get("거래량배율", 0))
        sc   = r.get("score",        r.get("점수", 0))
        top_lines.append(
            f"  {rank:2d}. {str(name):<10}  점수:{sc:.0f}  눌림:{pb:.1f}%  거래량×{vr:.1f}"
        )

    body = "\n".join([
        "안녕하세요,",
        "",
        f"[{ts}] 눌림목 스캐너 분석이 완료되었습니다.",
        "",
        "■ 요약",
        f"  발굴 종목 수 : {n_stock}개",
        "",
        f"■ 점수 상위 {len(top_lines)}개 종목",
        *top_lines,
        "",
        "분석 결과 Excel 파일이 첨부되어 있습니다.",
        "",
        "※ 이 메일은 자동 발송입니다.",
    ])

    msg            = MIMEMultipart()
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(_main_emails) if _main_emails else GMAIL_USER
    if _extra_emails:
        msg["Bcc"] = ", ".join(_extra_emails)
    msg["Subject"] = f"[눌림목스캐너] {ts} 스캔 완료 — {n_stock}개 종목"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    if file_name.exists():
        with open(file_name, "rb") as f:
            part = MIMEApplication(f.read(), Name=file_name.name)
            part["Content-Disposition"] = f'attachment; filename="{file_name.name}"'
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PW)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAILS, msg.as_string())
        _to_str  = ', '.join(_main_emails) if _main_emails else GMAIL_USER
        _bcc_str = f" | Bcc: {len(_extra_emails)}명" if _extra_emails else ""
        console.print(f"[green]이메일 발송 완료 — To: {_to_str}{_bcc_str}[/green]")
        if _extra_emails:
            console.print(f"  [dim]Bcc: {', '.join(_extra_emails)}[/dim]")
    except Exception as e:
        console.print(f"[red]이메일 발송 실패: {e}[/red]")


# ── 한글 폰트 자동 설정 ──────────────────────────────────────────────────────
for _font in ["AppleGothic", "Malgun Gothic", "NanumGothic", "NanumBarunGothic", "DejaVu Sans"]:
    matplotlib.rcParams["font.family"] = [_font]
    break
matplotlib.rcParams["axes.unicode_minus"] = False

# ── .env 로드 (python-dotenv가 없으면 수동 파싱) ────────────────────────────
def _load_env():
    env_path = Path(".env")
    if not env_path.exists():
        # 상위 폴더에서도 탐색
        for p in [Path("../.env"), Path("../../.env")]:
            if p.exists():
                env_path = p
                break
    if env_path.exists():
        try:
            from dotenv import load_dotenv
            load_dotenv(env_path)
        except ImportError:
            # python-dotenv 없을 때 직접 파싱
            with open(env_path) as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#") and "=" in line:
                        k, v = line.split("=", 1)
                        os.environ.setdefault(k.strip(), v.strip())

_load_env()
KRX_AUTH_KEY  = os.getenv("KRX_AUTH_KEY", "")
GMAIL_USER    = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW  = os.getenv("GMAIL_APP_PASSWORD", "")
SEND_EMAIL    = os.getenv("SEND_EMAIL", "false").lower() == "true"
_extra_emails = [e.strip() for e in re.split(r"[,;]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
_main_emails  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
NOTIFY_EMAILS = _main_emails + _extra_emails

# ── 설정 기본값 ───────────────────────────────────────────────────────────────
DEFAULT = {
    # 박스권 / 돌파
    "box_days":        30,    # 박스권 탐지 기간 (거래일) — 1~2개월 눌림
    "box_tol":         0.25,  # 박스 허용 범위 ±25% (총 범위 ~56%)
    "vol_mult":        2.5,   # 돌파일 거래량 배수
    "break_lookback":  10,    # 돌파 후 허용 경과 거래일
    "pullback_min":    0.02,  # 눌림 최솟값 2%
    "pullback_max":    0.20,  # 눌림 최댓값 20%
    # MA 필터
    "daily_ma":        60,    # 일봉 이동평균 (60일)
    "daily_slope_min": -0.003,# 일봉 MA — 완만한 하락도 허용
    "weekly_ma":       20,    # 주봉 이동평균 (20주)
    "weekly_slope_min": 0.001,# 주봉 MA 기울기 최솟값 (완화)
    "monthly_ma":      5,     # 월봉 이동평균 (5월)
    "monthly_slope_min":0.002,# 월봉 MA 기울기 최솟값 (완화)
    # 기타
    "min_cap_億":      1000,
    "market":          "ALL",
    "top_n":           20,
    # 차트
    "chart_top":       10,
    "chart_dir":       "charts",
    "daily_bars":      120,   # 일봉 차트 표시 봉 수
    "weekly_bars":     52,    # 주봉
    "monthly_bars":    36,    # 월봉
}

CHART_DIR = Path(DEFAULT["chart_dir"])

# ─────────────────────────────────────────────────────────────────────────────
# 1. 데이터 수집
# ─────────────────────────────────────────────────────────────────────────────

def _pykrx_ohlcv(ticker: str, start: str, end: str) -> pd.DataFrame:
    from pykrx import stock
    df = stock.get_market_ohlcv(start, end, ticker)
    if df.empty:
        return pd.DataFrame()
    df.columns = ["open","high","low","close","volume","tv","pc","pp"]
    df = df[["open","high","low","close","volume"]].copy()
    df.index = pd.to_datetime(df.index)
    return df[df["volume"] > 0]


def _krx_api_ohlcv(ticker: str, start: str, end: str) -> pd.DataFrame:
    """
    KRX OpenAPI (data.krx.co.kr) — AUTH_KEY가 .env에 있을 때 사용.
    일별 시세 조회: /data/quotations/stocks/price-summary
    """
    import urllib.request, json, urllib.parse
    base = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"
    params = {
        "bld":    "dbms/MDC/STAT/standard/MDCSTAT01701",
        "locale": "ko_KR",
        "isuCd": ticker,
        "strtDd": start.replace("-",""),
        "endDd":  end.replace("-",""),
        "AUTH_KEY": KRX_AUTH_KEY,
    }
    url = base + "?" + urllib.parse.urlencode(params)
    try:
        with urllib.request.urlopen(url, timeout=10) as r:
            data = json.loads(r.read().decode())
        rows = data.get("output", [])
        if not rows:
            return pd.DataFrame()
        records = []
        for row in rows:
            records.append({
                "date":   pd.to_datetime(row.get("TRD_DD","").replace("/","-")),
                "open":   float(str(row.get("TDD_OPNPRC","0")).replace(",","")),
                "high":   float(str(row.get("TDD_HGPRC","0")).replace(",","")),
                "low":    float(str(row.get("TDD_LWPRC","0")).replace(",","")),
                "close":  float(str(row.get("TDD_CLSPRC","0")).replace(",","")),
                "volume": float(str(row.get("ACC_TRDVOL","0")).replace(",","")),
            })
        df = pd.DataFrame(records).set_index("date").sort_index()
        return df[df["volume"] > 0]
    except Exception:
        return pd.DataFrame()


def get_daily_ohlcv(ticker: str, days: int = 300) -> pd.DataFrame:
    end   = datetime.today()
    start = end - timedelta(days=int(days * 1.8))

    # FDR 우선 (Naver Finance 기반, 단일 요청으로 빠름)
    try:
        import FinanceDataReader as fdr
        df = fdr.DataReader(ticker, start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"))
        if not df.empty:
            df.index = pd.to_datetime(df.index)
            df.columns = [c.lower() for c in df.columns]
            df = df[["open", "high", "low", "close", "volume"]].copy()
            df = df.apply(pd.to_numeric, errors="coerce").dropna()
            return df[df["volume"] > 0]
    except Exception:
        pass

    # pykrx 폴백
    s, e = start.strftime("%Y%m%d"), end.strftime("%Y%m%d")
    if KRX_AUTH_KEY:
        df = _krx_api_ohlcv(ticker, s, e)
        if not df.empty:
            return df
    try:
        return _pykrx_ohlcv(ticker, s, e)
    except Exception:
        return pd.DataFrame()


def to_weekly(df: pd.DataFrame) -> pd.DataFrame:
    """일봉 → 주봉 (금요일 기준)."""
    if df.empty:
        return df
    w = df.resample("W-FRI").agg(
        open=("open","first"), high=("high","max"),
        low=("low","min"),    close=("close","last"),
        volume=("volume","sum")
    ).dropna()
    return w[w["volume"] > 0]


def to_monthly(df: pd.DataFrame) -> pd.DataFrame:
    """일봉 → 월봉 (월말 기준)."""
    if df.empty:
        return df
    m = df.resample("ME").agg(
        open=("open","first"), high=("high","max"),
        low=("low","min"),    close=("close","last"),
        volume=("volume","sum")
    ).dropna()
    return m[m["volume"] > 0]


def get_market_cap(ticker: str) -> float:
    try:
        from pykrx import stock
        end   = datetime.today()
        start = end - timedelta(days=7)   # 비거래일(주말·공휴일) 대비 7일 범위
        df = stock.get_market_cap(
            start.strftime("%Y%m%d"), end.strftime("%Y%m%d"), ticker
        )
        return float(df["시가총액"].iloc[-1]) / 1e8 if not df.empty else 0
    except Exception:
        return 0


def get_all_market_caps() -> dict[str, float]:
    """전체 종목 시총을 한 번에 조회 (억원 단위). FinanceDataReader 사용."""
    try:
        import FinanceDataReader as fdr
        df = fdr.StockListing("KRX")
        code_col = next((c for c in df.columns if c in ("Code", "Symbol")), None)
        cap_col  = next((c for c in df.columns if c in ("Marcap", "MarCap", "시가총액")), None)
        if code_col is None or cap_col is None:
            return {}
        result: dict[str, float] = {}
        for _, row in df.iterrows():
            code = str(row[code_col]).zfill(6)
            cap  = float(row[cap_col] or 0) / 1e8   # 원 → 억원
            result[code] = cap
        return result
    except Exception:
        return {}


def get_ticker_name(ticker: str) -> str:
    try:
        from pykrx import stock
        return stock.get_market_ticker_name(ticker)
    except Exception:
        return ticker


def get_ticker_list(market: str) -> list:
    from pykrx import stock
    today = datetime.today().strftime("%Y%m%d")
    if market.upper() == "KOSPI":
        return stock.get_market_ticker_list(today, market="KOSPI")
    elif market.upper() == "KOSDAQ":
        return stock.get_market_ticker_list(today, market="KOSDAQ")
    return (stock.get_market_ticker_list(today, market="KOSPI") +
            stock.get_market_ticker_list(today, market="KOSDAQ"))


# ─────────────────────────────────────────────────────────────────────────────
# 2. MA 기울기 판단
# ─────────────────────────────────────────────────────────────────────────────

def ma_slope(series: pd.Series, period: int, look: int = 4) -> tuple[float, bool, bool]:
    """
    MA(period)의 최근 look 구간 기울기 계산.

    반환:
        slope_pct  : look 구간 변화율 (%)
        is_rising  : 연속 상승 여부
        is_flat    : 횡보 여부 (|slope| < 0.5% 이면 횡보로 판단)
    """
    ma = series.rolling(period).mean().dropna()
    if len(ma) < look + 1:
        return 0.0, False, False
    window = ma.iloc[-(look + 1):]
    diffs  = window.diff().dropna()
    slope  = (window.iloc[-1] / window.iloc[0] - 1) * 100
    rising = bool((diffs > 0).all())
    flat   = abs(slope) < 0.5
    return float(slope), rising, flat


# ─────────────────────────────────────────────────────────────────────────────
# 3. 박스권 / 돌파 / 눌림목
# ─────────────────────────────────────────────────────────────────────────────

def detect_box(df: pd.DataFrame, box_days: int, tol: float):
    if len(df) < box_days + 5:
        return None, None, False
    w = df.iloc[-(box_days + 5):-5]
    hi, lo = w["high"].max(), w["low"].min()
    if lo == 0:
        return None, None, False
    return float(hi), float(lo), (hi - lo) / lo <= tol * 2


def detect_breakout(df: pd.DataFrame, box_top: float, vol_mult: float, lookback: int):
    if len(df) < 25:
        return None, None, None, None
    vol_ma = df["volume"].rolling(20).mean()
    recent = df.iloc[-lookback:]
    for i in range(len(recent)):
        row = recent.iloc[i]
        idx = recent.index[i]
        avg = vol_ma.loc[idx] if idx in vol_ma.index else None
        if avg is None or avg == 0:
            continue
        vr = row["volume"] / avg
        if row["close"] > box_top and vr >= vol_mult:
            after = df.loc[idx:]
            peak  = float(after["high"].max()) if not after.empty else float(row["close"])
            return idx, float(row["close"]), peak, float(vr)
    return None, None, None, None


def detect_pullback(df: pd.DataFrame, peak: float, pb_min: float, pb_max: float):
    if df.empty or peak == 0:
        return None, False, None
    cur = float(df["close"].iloc[-1])
    pb  = (peak - cur) / peak
    return pb, (pb_min <= pb <= pb_max), cur


def score_calc(vol_ratio, pb_pct, box_days, range_ratio,
               d_slope, w_slope, m_slope) -> int:
    vol_s   = min(vol_ratio / 6.0, 1.0) * 30
    ideal   = 0.075
    pull_s  = max(0, 1.0 - abs(pb_pct - ideal) / 0.075) * 20
    period_s= min(box_days / 90.0, 1.0) * 10
    tight_s = max(0, 1.0 - range_ratio / 0.16) * 10
    # MA 기울기 보너스 (각 10점)
    d_s = min(max(d_slope + 1, 0) / 2.0, 1.0) * 10  # 횡보 포함 허용
    w_s = min(w_slope / 1.5, 1.0) * 10
    m_s = min(m_slope / 2.0, 1.0) * 10
    return int(vol_s + pull_s + period_s + tight_s + d_s + w_s + m_s)


# ─────────────────────────────────────────────────────────────────────────────
# 4. 단일 종목 분석
# ─────────────────────────────────────────────────────────────────────────────

def analyze(ticker: str, cfg: dict, cap: float = 0.0) -> dict | None:
    try:
        # 시총 사전 조회 값이 있으면 OHLCV 다운로드 전에 조기 탈락
        if cap > 0 and cap < cfg["min_cap_億"]:
            return None

        # 일봉 (충분히 길게)
        df_d = get_daily_ohlcv(ticker, days=400)
        if df_d.empty or len(df_d) < cfg["box_days"] + 60:
            return None

        # 사전 조회 값이 없으면 개별 조회
        if cap <= 0:
            cap = get_market_cap(ticker)
        if cap < cfg["min_cap_億"]:
            return None

        # ── 일봉 60일 MA: 횡보 or 우상향 ──
        d_slope, d_rising, d_flat = ma_slope(df_d["close"], cfg["daily_ma"], look=4)
        if not (d_rising or d_flat):   # 하락이면 탈락
            return None
        if d_slope < cfg["daily_slope_min"] * 100:
            return None

        # ── 주봉 20주 MA: 우상향 ──
        df_w = to_weekly(df_d)
        if len(df_w) < cfg["weekly_ma"] + 5:
            return None
        w_slope, w_rising, _ = ma_slope(df_w["close"], cfg["weekly_ma"], look=4)
        if not w_rising or w_slope < cfg["weekly_slope_min"] * 100:
            return None

        # ── 월봉 5월 MA: 우상향 ──
        df_m = to_monthly(df_d)
        if len(df_m) < cfg["monthly_ma"] + 3:
            return None
        m_slope, m_rising, _ = ma_slope(df_m["close"], cfg["monthly_ma"], look=3)
        if not m_rising or m_slope < cfg["monthly_slope_min"] * 100:
            return None

        # ── 박스권 / 돌파 / 눌림목 (일봉 기준) ──
        box_top, box_bottom, is_box = detect_box(df_d, cfg["box_days"], cfg["box_tol"])
        if not is_box:
            return None
        range_ratio = (box_top - box_bottom) / box_bottom if box_bottom else 0

        brk_date, brk_price, peak, vol_ratio = detect_breakout(
            df_d, box_top, cfg["vol_mult"], cfg["break_lookback"])
        if brk_date is None:
            return None

        pb_pct, is_pb, cur = detect_pullback(
            df_d, peak, cfg["pullback_min"], cfg["pullback_max"])
        if not is_pb:
            return None

        sc = score_calc(vol_ratio, pb_pct, cfg["box_days"], range_ratio,
                        d_slope, w_slope, m_slope)

        # ── 유형 판별: 바닥탈출 vs 우상향 ──
        # 52주(최대 252 거래일) 최저가 대비 현재가 상승률
        days_252 = min(252, len(df_d))
        w52_low = float(df_d["low"].iloc[-days_252:].min())
        from_52w_low = (cur / w52_low - 1) * 100 if w52_low > 0 else 0.0
        # 일봉 MA120 기울기: 장기 추세가 여전히 하락/보합이면 바닥탈출
        ma120_slope, _, _ = ma_slope(df_d["close"], 120, look=4)
        is_bottom = from_52w_low >= 40.0 and ma120_slope < 2.0
        pattern = "바닥탈출" if is_bottom else "우상향"

        return dict(
            ticker=ticker, name=get_ticker_name(ticker),
            current_price=cur, break_date=brk_date.strftime("%m/%d"),
            break_price=brk_price, peak_price=peak,
            box_top=box_top, box_bottom=box_bottom,
            vol_ratio=vol_ratio, pullback_pct=pb_pct * 100,
            cap_億=cap, score=sc, range_ratio=range_ratio * 100,
            from_52w_low=from_52w_low, pattern=pattern,
            d_slope=d_slope, w_slope=w_slope, m_slope=m_slope,
            d_flat=d_flat,
        )
    except Exception:
        return None


# ─────────────────────────────────────────────────────────────────────────────
# 5. 차트 그리기 (공통 엔진)
# ─────────────────────────────────────────────────────────────────────────────

BG   = "#0d1117"
GRID = "#21262d"
UP   = "#f85149"
DN   = "#388bfd"
MA_C = {"ma1":"#e3b341","ma2":"#f0883e","ma3":"#a371f7","vol":"#e3b341"}
GREEN= "#3fb950"
AMBER= "#e3b341"
BLUE = "#388bfd"

def _candles(ax, df: pd.DataFrame):
    """캔들 그리기."""
    for i, (_, r) in enumerate(df.iterrows()):
        c = UP if r["close"] >= r["open"] else DN
        ax.plot([i,i],[r["low"],r["high"]], color=c, lw=0.7, alpha=0.85)
        lo = min(r["open"],r["close"])
        ht = max(abs(r["close"]-r["open"]), r["close"]*0.003)
        ax.add_patch(matplotlib.patches.Rectangle(
            (i-0.35, lo), 0.7, ht, facecolor=c, edgecolor=c, lw=0, alpha=0.88))


def _ma_lines(ax, closes: pd.Series, configs: list):
    """MA 선 그리기. configs = [(period, label, color, lw), ...]"""
    xs = np.arange(len(closes))
    for period, label, color, lw in configs:
        ma = closes.rolling(period).mean()
        ax.plot(xs, ma.values, color=color, lw=lw, label=label, alpha=0.92)


def _volume_bars(ax, df: pd.DataFrame, highlight_idx: int = -1):
    for i, (_, r) in enumerate(df.iterrows()):
        c = GREEN if i == highlight_idx else (UP if r["close"] >= r["open"] else DN)
        ax.bar(i, r["volume"]/1e4, color=c, width=0.7, alpha=0.75)


def _style_ax(ax):
    ax.set_facecolor(BG)
    ax.tick_params(colors="#8b949e", labelsize=8)
    for sp in ax.spines.values():
        sp.set_color(GRID)
    ax.grid(color=GRID, linestyle="--", lw=0.4, alpha=0.6)


def _x_labels(ax2, df: pd.DataFrame, fmt: str, step: int):
    xs = np.arange(len(df))
    ax2.set_xticks(xs[::step])
    ax2.set_xticklabels(
        [df.index[i].strftime(fmt) for i in range(0, len(df), step)],
        rotation=30, ha="right")


def _legend(ax, handles):
    ax.legend(handles=handles, loc="upper left", fontsize=8,
              facecolor="#161b22", edgecolor=GRID, labelcolor="#e6edf3", framealpha=0.92)


def _patch(color, label, alpha=1.0):
    return mpatches.Patch(color=color, label=label, alpha=alpha)


# ─── 일봉 차트 ────────────────────────────────────────────────────────────────
def chart_daily(df_full: pd.DataFrame, info: dict, cfg: dict, save_dir: Path) -> Path:
    df = df_full.tail(cfg["daily_bars"]).copy()
    xs = np.arange(len(df))

    fig, (ax1, ax2) = plt.subplots(2,1,figsize=(14,8),
        gridspec_kw={"height_ratios":[3,1]}, sharex=True)
    fig.patch.set_facecolor(BG)
    for ax in [ax1,ax2]: _style_ax(ax)

    _candles(ax1, df)
    _ma_lines(ax1, df["close"], [
        (5,  "MA5",  MA_C["ma1"], 0.9),
        (20, "MA20", MA_C["ma2"], 0.9),
        (60, "MA60", MA_C["ma3"], 1.4),
    ])

    # 박스권
    bt, bb = info.get("box_top"), info.get("box_bottom")
    if bt:
        ax1.axhline(bt, color=BLUE, lw=1.0, ls="--", alpha=0.75)
        ax1.text(len(df)-1, bt*1.003, f"박스상단 {bt:,.0f}",
                 color=BLUE, fontsize=7, ha="right")
    if bb:
        ax1.axhline(bb, color=BLUE, lw=0.7, ls=":", alpha=0.5)
    if bt and bb:
        ax1.axhspan(bb, bt, alpha=0.07, color=BLUE)

    # 돌파 마커 (화면 내 있으면)
    brk = info.get("break_date")
    if brk:
        match = [i for i,(idx,_) in enumerate(df.iterrows())
                 if idx.strftime("%m/%d") == brk]
        if match:
            bi = match[0]
            ax1.scatter(bi, df["close"].iloc[bi]*0.965,
                        marker="^", color=GREEN, s=110, zorder=6)
            ax1.annotate(f"돌파 ×{info['vol_ratio']:.1f}",
                         xy=(bi, df["close"].iloc[bi]),
                         xytext=(bi-3, df["close"].iloc[bi]*1.05),
                         fontsize=8, color=GREEN,
                         arrowprops=dict(arrowstyle="->",color=GREEN,lw=0.8))
            _volume_bars(ax2, df, highlight_idx=bi)
        else:
            _volume_bars(ax2, df)
    else:
        _volume_bars(ax2, df)

    # 눌림목 현재가 표시
    cur = info.get("current_price")
    pb  = info.get("pullback_pct", 0)
    if cur:
        ax1.axhline(cur, color=AMBER, lw=0.9, ls="-.", alpha=0.8)
        ax1.text(len(df)-1, cur*1.003,
                 f"현재가 {cur:,.0f}  (눌림 ▼{pb:.1f}%)",
                 color=AMBER, fontsize=8, ha="right",
                 bbox=dict(boxstyle="round,pad=0.25",
                           facecolor="#1f2937", edgecolor=AMBER, alpha=0.85))

    # MA60 상태 레이블
    d_slope = info.get("d_slope",0)
    d_flat  = info.get("d_flat", False)
    state   = "횡보" if d_flat else ("우상향 ↗" if d_slope >= 0 else "하락")
    ax1.text(1, df["high"].max()*0.998,
             f"MA60  {state}  기울기 {d_slope:+.2f}%",
             fontsize=9, color=MA_C["ma3"], fontweight="bold",
             bbox=dict(boxstyle="round,pad=0.3",
                       facecolor="#1a0a2e", edgecolor=MA_C["ma3"], alpha=0.85))

    vol_ma10 = df["volume"].rolling(10).mean() / 1e4
    ax2.plot(xs, vol_ma10.values, color=AMBER, lw=0.9, label="거래량 MA10")

    _x_labels(ax2, df, "%m/%d", max(1, len(df)//10))
    ax1.set_title(
        f"{info['name']}({info['ticker']})  일봉  |  점수 {info['score']}",
        color="#e6edf3", fontsize=12, fontweight="bold", pad=10)
    ax1.set_ylabel("주가 (원)", color="#8b949e", fontsize=9)
    ax2.set_ylabel("거래량 (만주)", color="#8b949e", fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"{x:,.0f}"))
    _legend(ax1, [
        _patch(MA_C["ma1"],"MA5"), _patch(MA_C["ma2"],"MA20"),
        _patch(MA_C["ma3"],"MA60 (횡보/우상향)"),
        _patch(BLUE,"박스권",0.4), _patch(GREEN,"돌파"),
        _patch(AMBER,"현재가 (눌림목)"),
    ])
    _legend(ax2, [_patch(AMBER,"거래량 MA10"), _patch(GREEN,"돌파일")])
    plt.tight_layout()

    fpath = save_dir / f"{info['ticker']}_{info['name']}_일봉.png"
    fig.savefig(fpath, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    return fpath


# ─── 주봉 차트 ────────────────────────────────────────────────────────────────
def chart_weekly(df_daily: pd.DataFrame, info: dict, cfg: dict, save_dir: Path) -> Path:
    df_w_full = to_weekly(df_daily)
    df = df_w_full.tail(cfg["weekly_bars"]).copy()
    xs = np.arange(len(df))

    fig, (ax1, ax2) = plt.subplots(2,1,figsize=(14,8),
        gridspec_kw={"height_ratios":[3,1]}, sharex=True)
    fig.patch.set_facecolor(BG)
    for ax in [ax1,ax2]: _style_ax(ax)

    _candles(ax1, df)
    _ma_lines(ax1, df["close"], [
        (5,  "MA5",  MA_C["ma1"], 0.9),
        (10, "MA10", MA_C["ma2"], 0.9),
        (20, "MA20 (우상향 ↗)", MA_C["ma3"], 1.5),
    ])

    # 박스권 (주봉 스케일)
    bt, bb = info.get("box_top"), info.get("box_bottom")
    if bt:
        ax1.axhline(bt, color=BLUE, lw=1.0, ls="--", alpha=0.7)
        ax1.text(len(df)-1, bt*1.003, f"박스상단", color=BLUE, fontsize=7, ha="right")
    if bb:
        ax1.axhline(bb, color=BLUE, lw=0.7, ls=":", alpha=0.45)
    if bt and bb:
        ax1.axhspan(bb, bt, alpha=0.07, color=BLUE)

    cur = info.get("current_price")
    if cur:
        ax1.axhline(cur, color=AMBER, lw=0.9, ls="-.", alpha=0.8)

    # MA20 기울기 레이블
    w_slope = info.get("w_slope", 0)
    ax1.text(1, df["high"].max()*0.998,
             f"주봉 MA20  우상향 ↗  기울기 {w_slope:+.2f}%",
             fontsize=9, color=MA_C["ma3"], fontweight="bold",
             bbox=dict(boxstyle="round,pad=0.3",
                       facecolor="#1a0a2e", edgecolor=MA_C["ma3"], alpha=0.85))

    _volume_bars(ax2, df)
    vol_ma = df["volume"].rolling(10).mean() / 1e4
    ax2.plot(xs, vol_ma.values, color=AMBER, lw=0.9, label="거래량 MA10")

    _x_labels(ax2, df, "%y/%m", max(1, len(df)//10))
    ax1.set_title(
        f"{info['name']}({info['ticker']})  주봉  |  MA20 기울기 {w_slope:+.2f}%",
        color="#e6edf3", fontsize=12, fontweight="bold", pad=10)
    ax1.set_ylabel("주가 (원)", color="#8b949e", fontsize=9)
    ax2.set_ylabel("거래량 (만주)", color="#8b949e", fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"{x:,.0f}"))
    _legend(ax1,[
        _patch(MA_C["ma1"],"MA5"), _patch(MA_C["ma2"],"MA10"),
        _patch(MA_C["ma3"],"MA20 ↗"), _patch(BLUE,"박스권",0.4),
        _patch(AMBER,"현재가"),
    ])
    plt.tight_layout()
    fpath = save_dir / f"{info['ticker']}_{info['name']}_주봉.png"
    fig.savefig(fpath, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    return fpath


# ─── 월봉 차트 ────────────────────────────────────────────────────────────────
def chart_monthly(df_daily: pd.DataFrame, info: dict, cfg: dict, save_dir: Path) -> Path:
    df_m_full = to_monthly(df_daily)
    df = df_m_full.tail(cfg["monthly_bars"]).copy()
    xs = np.arange(len(df))

    fig, (ax1, ax2) = plt.subplots(2,1,figsize=(14,8),
        gridspec_kw={"height_ratios":[3,1]}, sharex=True)
    fig.patch.set_facecolor(BG)
    for ax in [ax1,ax2]: _style_ax(ax)

    _candles(ax1, df)
    _ma_lines(ax1, df["close"], [
        (3, "MA3",  MA_C["ma1"], 0.9),
        (5, "MA5 (우상향 ↗)", MA_C["ma3"], 1.5),
        (10,"MA10", MA_C["ma2"], 0.9),
    ])

    cur = info.get("current_price")
    if cur:
        ax1.axhline(cur, color=AMBER, lw=0.9, ls="-.", alpha=0.8)

    # 장기 지지/저항 (박스권 주요 레벨)
    bt, bb = info.get("box_top"), info.get("box_bottom")
    if bt:
        ax1.axhline(bt, color=BLUE, lw=0.8, ls="--", alpha=0.6)
    if bb:
        ax1.axhline(bb, color=BLUE, lw=0.6, ls=":", alpha=0.4)

    m_slope = info.get("m_slope", 0)
    ax1.text(1, df["high"].max()*0.998,
             f"월봉 MA5  우상향 ↗  기울기 {m_slope:+.2f}%",
             fontsize=9, color=MA_C["ma3"], fontweight="bold",
             bbox=dict(boxstyle="round,pad=0.3",
                       facecolor="#1a0a2e", edgecolor=MA_C["ma3"], alpha=0.85))

    _volume_bars(ax2, df)
    vol_ma = df["volume"].rolling(6).mean() / 1e4
    ax2.plot(xs, vol_ma.values, color=AMBER, lw=0.9, label="거래량 MA6")

    _x_labels(ax2, df, "%Y/%m", max(1, len(df)//8))
    ax1.set_title(
        f"{info['name']}({info['ticker']})  월봉  |  MA5 기울기 {m_slope:+.2f}%",
        color="#e6edf3", fontsize=12, fontweight="bold", pad=10)
    ax1.set_ylabel("주가 (원)", color="#8b949e", fontsize=9)
    ax2.set_ylabel("거래량 (만주)", color="#8b949e", fontsize=9)
    ax1.yaxis.set_major_formatter(mticker.FuncFormatter(lambda x,_: f"{x:,.0f}"))
    _legend(ax1,[
        _patch(MA_C["ma1"],"MA3"), _patch(MA_C["ma3"],"MA5 ↗"),
        _patch(MA_C["ma2"],"MA10"), _patch(AMBER,"현재가"),
    ])
    plt.tight_layout()
    fpath = save_dir / f"{info['ticker']}_{info['name']}_월봉.png"
    fig.savefig(fpath, dpi=150, bbox_inches="tight", facecolor=BG)
    plt.close(fig)
    return fpath


# ─────────────────────────────────────────────────────────────────────────────
# 6. 3종 차트 일괄 생성
# ─────────────────────────────────────────────────────────────────────────────

def generate_charts(ticker: str, info: dict, cfg: dict, save_dir: Path) -> list[Path]:
    save_dir.mkdir(parents=True, exist_ok=True)
    ticker_dir = save_dir / f"{ticker}_{info['name']}"
    ticker_dir.mkdir(exist_ok=True)

    df_d = get_daily_ohlcv(ticker, days=500)   # 월봉까지 충분히
    if df_d.empty:
        print(f"    [WARN] {ticker} 데이터 없음")
        return []

    paths = []
    for fn, label in [
        (chart_daily,   "일봉"),
        (chart_weekly,  "주봉"),
        (chart_monthly, "월봉"),
    ]:
        try:
            p = fn(df_d, info, cfg, ticker_dir)
            paths.append(p)
            print(f"      {label} → {p.name}")
        except Exception as e:
            print(f"      [ERR] {label} 차트 실패: {e}")
    return paths


# ─────────────────────────────────────────────────────────────────────────────
# 7. 스캔 실행
# ─────────────────────────────────────────────────────────────────────────────

def run_scanner(cfg: dict) -> pd.DataFrame:
    console.print()
    console.rule("[bold cyan]박스권 돌파 눌림목 스캐너  v3[/bold cyan]")
    console.print(f"  시장        : [cyan]{cfg['market']}[/cyan]")
    console.print(f"  데이터      : {'KRX API + pykrx' if KRX_AUTH_KEY else 'pykrx'}")
    console.print(f"  박스 기간   : {cfg['box_days']}거래일  |  거래량 ×{cfg['vol_mult']:.1f}")
    console.print(f"  눌림목      : {cfg['pullback_min']*100:.0f}% ~ {cfg['pullback_max']*100:.0f}%")
    console.print(f"  일봉 MA{cfg['daily_ma']:02d}   : 횡보 또는 우상향  (기울기 ≥ {cfg['daily_slope_min']*100:.1f}%)")
    console.print(f"  주봉 MA{cfg['weekly_ma']:02d}   : 우상향  (기울기 ≥ {cfg['weekly_slope_min']*100:.1f}%)")
    console.print(f"  월봉 MA{cfg['monthly_ma']:02d}    : 우상향  (기울기 ≥ {cfg['monthly_slope_min']*100:.1f}%)")
    console.print(f"  시가총액    : {cfg['min_cap_億']:,}억 이상")
    console.print()

    tickers = get_ticker_list(cfg["market"])

    # 시총 일괄 조회 → min_cap 미달 종목 사전 제거
    console.print("  [dim]시가총액 일괄 조회 중...[/dim]")
    all_caps = get_all_market_caps()
    if all_caps and cfg["min_cap_億"] > 0:
        filtered = [t for t in tickers if all_caps.get(t, 0) >= cfg["min_cap_億"]]
        if filtered:
            tickers = filtered
            console.print(f"  시총 ≥ {cfg['min_cap_億']:,}억 필터 후: [bold]{len(tickers):,}[/bold]개")
        else:
            console.print(f"  [yellow]시총 사전 필터 매칭 없음 — 전체 {len(tickers):,}개 스캔[/yellow]")
            console.print(f"  [dim]티커 샘플: {tickers[:3]}  |  시총키 샘플: {list(all_caps.keys())[:3]}[/dim]")
    elif not all_caps:
        console.print(f"  [yellow]시총 일괄 조회 실패 — 전체 {len(tickers):,}개 스캔 (종목당 개별 조회)[/yellow]")
    console.print()

    total   = len(tickers)
    console.print(f"  총 [bold]{total:,}[/bold]개 종목 스캔 시작...")
    console.print()

    results = []

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=20, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TextColumn("남은"),
        TimeRemainingColumn(),
        console=console,
        transient=False,
    ) as progress:
        task = progress.add_task("스캔 중...", total=total)

        TICKER_TIMEOUT = 20  # 종목당 최대 대기 시간 (초)
        MAX_WORKERS    = 4
        WINDOW         = MAX_WORKERS       # 대기 없이 실행 중인 것만 유지 → 타임아웃 정확도 보장

        ticker_iter = iter(tickers)
        # pending: {future → (ticker, submit_time)}
        pending: dict = {}

        def _submit_next():
            t = next(ticker_iter, None)
            if t is None:
                return
            cap = all_caps.get(t, 0)
            f   = pool.submit(analyze, t, cfg, cap)
            pending[f] = (t, time.time())

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as pool:
            for _ in range(WINDOW):          # 초기 윈도우 채우기
                _submit_next()

            while pending:
                done, _ = wait(pending.keys(), timeout=1.0, return_when=FIRST_COMPLETED)

                for future in done:
                    ticker, _ = pending.pop(future)
                    try:
                        r = future.result()
                    except Exception:
                        r = None
                    if r:
                        results.append(r)
                        progress.update(task, description=f"[green]{r['name'][:12]}[/green] *")
                    else:
                        progress.update(task, description=f"[dim]{ticker}[/dim]")
                    progress.advance(task)
                    _submit_next()           # 완료되면 다음 종목 즉시 제출

                # 제출 시점 기준 25초 초과 future → 스킵 후 다음 제출
                now     = time.time()
                expired = [f for f, (_, st) in list(pending.items())
                           if now - st > TICKER_TIMEOUT]
                for future in expired:
                    ticker, _ = pending.pop(future)
                    progress.update(task, description=f"[yellow]T/O {ticker}[/yellow]")
                    progress.advance(task)
                    _submit_next()

    console.print()
    console.print(f"  [green]완료 -- {len(results)}개 종목 발굴[/green]")
    console.print()

    if not results:
        console.print("  [yellow]조건에 맞는 종목이 없습니다. --vol-mult 또는 --pullback-max 를 완화해보세요.[/yellow]")
        return pd.DataFrame()

    df_r = (pd.DataFrame(results)
            .sort_values("score", ascending=False)
            .head(cfg["top_n"])
            .reset_index(drop=True))
    return df_r


def print_results(df: pd.DataFrame):
    if df.empty:
        return
    try:
        from tabulate import tabulate
        d = df[[
            "ticker","name","pattern","current_price","break_date",
            "vol_ratio","pullback_pct","from_52w_low",
            "d_slope","w_slope","m_slope",
            "box_top","cap_億","score"
        ]].copy()
        d.columns = [
            "코드","종목명","유형","현재가","돌파일",
            "거래량×","눌림(%)","52주저점대비",
            "일봉MA60기울기","주봉MA20기울기","월봉MA5기울기",
            "박스상단","시총(억)","점수"
        ]
        for col in ["현재가","박스상단"]:
            d[col] = d[col].apply(lambda x: f"{x:,.0f}")
        d["거래량×"] = d["거래량×"].apply(lambda x: f"×{x:.1f}")
        d["눌림(%)"] = d["눌림(%)"].apply(lambda x: f"-{x:.1f}%")
        d["52주저점대비"] = d["52주저점대비"].apply(lambda x: f"+{x:.0f}%")
        for col in ["일봉MA60기울기","주봉MA20기울기","월봉MA5기울기"]:
            d[col] = d[col].apply(lambda x: f"{x:+.2f}%")
        d["시총(억)"] = d["시총(억)"].apply(lambda x: f"{x:,.0f}")
        print(tabulate(d, headers="keys", tablefmt="rounded_outline", showindex=True))
    except ImportError:
        print(df.to_string())

    print("""
  점수 = 거래량강도(30) + 눌림적절성(20) + 박스기간(10) + 박스타이트(10)
         + 일봉MA(10) + 주봉MA(10) + 월봉MA(10)   [합계 100점]

  매수 가이드:
    1) 박스 상단(전 저항 → 신규 지지) 부근까지 눌림 확인
    2) 일봉 캔들 양봉 + 거래량 증가 시 진입
    3) 손절: 박스 상단 하향 이탈  |  목표: 박스 폭만큼 상승
""")


# ─────────────────────────────────────────────────────────────────────────────
# 8. CLI
# ─────────────────────────────────────────────────────────────────────────────

def parse_args():
    p = argparse.ArgumentParser(
        description="박스권 돌파 눌림목 스캐너 v3 — 일봉/주봉/월봉 MA 필터 + 3종 차트")
    p.add_argument("--market",            default=DEFAULT["market"],          choices=["ALL","KOSPI","KOSDAQ"])
    p.add_argument("--box-days",          type=int,   default=DEFAULT["box_days"])
    p.add_argument("--vol-mult",          type=float, default=DEFAULT["vol_mult"])
    p.add_argument("--break-lookback",    type=int,   default=DEFAULT["break_lookback"])
    p.add_argument("--pullback-min",      type=float, default=DEFAULT["pullback_min"])
    p.add_argument("--pullback-max",      type=float, default=DEFAULT["pullback_max"])
    p.add_argument("--min-cap",           type=float, default=DEFAULT["min_cap_億"])
    p.add_argument("--top-n",             type=int,   default=DEFAULT["top_n"])
    # MA
    p.add_argument("--daily-ma",          type=int,   default=DEFAULT["daily_ma"],
                   help="일봉 MA 기간 (기본 60)")
    p.add_argument("--weekly-ma",         type=int,   default=DEFAULT["weekly_ma"],
                   help="주봉 MA 기간 (기본 20)")
    p.add_argument("--monthly-ma",        type=int,   default=DEFAULT["monthly_ma"],
                   help="월봉 MA 기간 (기본 5)")
    # 차트
    p.add_argument("--chart-top",         type=int,   default=DEFAULT["chart_top"],
                   help="상위 N개 차트 저장 (기본 10)")
    p.add_argument("--chart-dir",         type=str,   default=DEFAULT["chart_dir"])
    p.add_argument("--no-chart",          action="store_true", help="차트 생성 안 함")
    p.add_argument("--chart-only",        type=str,   default=None,
                   help="특정 종목 코드 — 스캔 없이 3종 차트만 생성")
    p.add_argument("--save",              action="store_true", help="결과 CSV 저장")
    return p.parse_args()


if __name__ == "__main__":
    args  = parse_args()
    CHART_DIR = Path(args.chart_dir)

    cfg = {
        "box_days":         args.box_days,
        "box_tol":          DEFAULT["box_tol"],
        "vol_mult":         args.vol_mult,
        "break_lookback":   args.break_lookback,
        "pullback_min":     args.pullback_min,
        "pullback_max":     args.pullback_max,
        "min_cap_億":        args.min_cap,
        "market":           args.market,
        "top_n":            args.top_n,
        "daily_ma":         args.daily_ma,
        "daily_slope_min":  DEFAULT["daily_slope_min"],
        "weekly_ma":        args.weekly_ma,
        "weekly_slope_min": DEFAULT["weekly_slope_min"],
        "monthly_ma":       args.monthly_ma,
        "monthly_slope_min":DEFAULT["monthly_slope_min"],
        "chart_top":        args.chart_top,
        "chart_dir":        args.chart_dir,
        "daily_bars":       DEFAULT["daily_bars"],
        "weekly_bars":      DEFAULT["weekly_bars"],
        "monthly_bars":     DEFAULT["monthly_bars"],
    }

    # ── 특정 종목 차트만 ──────────────────────────────────────────────────────
    if args.chart_only:
        tk   = args.chart_only
        name = get_ticker_name(tk)
        info = dict(ticker=tk, name=name, score=0,
                    d_slope=0, w_slope=0, m_slope=0,
                    d_flat=False, box_top=None, box_bottom=None,
                    vol_ratio=0, pullback_pct=0, current_price=None)
        console.print(f"\n  {tk} {name} — 일봉 / 주봉 / 월봉 차트 생성 중...")
        generate_charts(tk, info, cfg, CHART_DIR)
        sys.exit(0)

    # ── 전체 스캔 ─────────────────────────────────────────────────────────────
    df_result = run_scanner(cfg)
    print_results(df_result)

    if not df_result.empty:
        excel_path = _save_excel(df_result)
        console.print(f"[green]Excel 저장 완료: {excel_path}[/green]")
        console.print()

        if SEND_EMAIL:
            console.rule("[bold cyan]이메일 발송[/bold cyan]")
            _send_email(excel_path, df_result)
            console.print()
        else:
            console.print("[dim]이메일 발송 생략 (SEND_EMAIL=false)[/dim]")
            console.print()

    if not args.no_chart and not df_result.empty:
        top = df_result.head(cfg["chart_top"])
        console.print(f"\n  상위 {len(top)}개 종목 3종 차트 생성 중...\n")
        for _, row in top.iterrows():
            console.print(f"  [{row['score']}점] {row['ticker']} {row['name']}")
            generate_charts(row["ticker"], row.to_dict(), cfg, CHART_DIR)
        console.print(f"\n  완료 -- ./{args.chart_dir}/ 폴더 확인\n")
