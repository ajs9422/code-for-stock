"""
=============================================================
  주도 섹터 · 주도주 자동 탐색 시스템  v1.0
=============================================================
  구조:
    hot_sector/
    ├── main.py          ← 이 파일 (실행 진입점)
    ├── config.yaml      ← 모든 파라미터 설정
    ├── requirements.txt ← 의존성
    └── reports/         ← 결과 CSV/Excel 저장 (자동 생성)

  실행:
    python main.py                    # 오늘 기준 즉시 실행
    python main.py --date 20260425    # 특정 날짜 분석
    python main.py --schedule         # 스케줄 모드 (장중 자동 실행)
    python main.py --top 10           # 섹터 TOP 10 출력

  Claude Code 보완 포인트 (TODO 태그로 표시):
    TODO:SCORE  - 핫 스코어 가중치 로직 개선
    TODO:FILTER - 주도주 필터 조건 추가
    TODO:ALERT  - 텔레그램 알림 커스터마이징
    TODO:VISUAL - Streamlit 대시보드 연동
=============================================================
"""

from __future__ import annotations

import argparse
import logging
import os
import sys
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Optional

# Windows 콘솔 UTF-8 출력
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8")
if sys.stderr.encoding != "utf-8":
    sys.stderr.reconfigure(encoding="utf-8")

from dotenv import load_dotenv
import re as _re
load_dotenv(Path(__file__).parent / ".env")

import numpy as np
import pandas as pd
import yaml

# ── 선택적 import (없어도 핵심 기능은 동작) ──────────────────
try:
    from pykrx import stock as krx
    HAS_PYKRX = True
except ImportError:
    HAS_PYKRX = False

try:
    import pandas_ta as ta
    HAS_TA = True
except ImportError:
    HAS_TA = False

try:
    import telegram
    HAS_TELEGRAM = True
except ImportError:
    HAS_TELEGRAM = False

try:
    from apscheduler.schedulers.blocking import BlockingScheduler
    HAS_SCHEDULER = True
except ImportError:
    HAS_SCHEDULER = False

try:
    from rich.console import Console
    from rich.progress import (
        BarColumn, MofNCompleteColumn, Progress,
        SpinnerColumn, TextColumn, TimeElapsedColumn, TimeRemainingColumn,
    )
    from rich import box as rich_box
    from rich.table import Column as RichColumn, Table
    HAS_RICH = True
except ImportError:
    HAS_RICH = False

console = Console() if HAS_RICH else None


# ═══════════════════════════════════════════════════════════════
# 로거 설정
# ═══════════════════════════════════════════════════════════════

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("hot_sector.log", encoding="utf-8"),
    ],
)
log = logging.getLogger("hot_sector")


# ═══════════════════════════════════════════════════════════════
# 설정 로더
# ═══════════════════════════════════════════════════════════════

class Config:
    """config.yaml을 로드하고 속성으로 접근 가능하게 래핑"""

    CONFIG_PATH = Path(__file__).parent / "config.yaml"

    def __init__(self, override: dict = None):
        raw = self._load()
        if override:
            raw = self._deep_merge(raw, override)
        self._raw = raw

    def _load(self) -> dict:
        if not self.CONFIG_PATH.exists():
            log.warning(f"config.yaml 없음 → 기본값 사용: {self.CONFIG_PATH}")
            return {}
        with open(self.CONFIG_PATH, encoding="utf-8") as f:
            return yaml.safe_load(f) or {}

    @staticmethod
    def _deep_merge(base: dict, override: dict) -> dict:
        for k, v in override.items():
            if isinstance(v, dict) and isinstance(base.get(k), dict):
                base[k] = Config._deep_merge(base[k], v)
            else:
                base[k] = v
        return base

    def get(self, *keys, default=None):
        """cfg.get("scoring", "weight_return") 형식으로 접근"""
        node = self._raw
        for k in keys:
            if not isinstance(node, dict):
                return default
            node = node.get(k, default)
        return node


# ═══════════════════════════════════════════════════════════════
# 유틸리티
# ═══════════════════════════════════════════════════════════════

_KR_HOLIDAYS: set[str] = {
    # 2025
    "20250101", "20250128", "20250129", "20250130", "20250301",
    "20250505", "20250506", "20250506", "20250606", "20250815",
    "20251003", "20251006", "20251007", "20251008", "20251009", "20251225",
    # 2026
    "20260101", "20260216", "20260217", "20260218", "20260301",
    "20260501", "20260505", "20260606", "20260815",
    "20261003", "20261009", "20261225",
}


def get_recent_business_day(date: datetime = None) -> str:
    """가장 최근 거래일 (주말·한국 공휴일 제외)"""
    d = date or datetime.today()
    while d.weekday() >= 5 or d.strftime("%Y%m%d") in _KR_HOLIDAYS:
        d -= timedelta(days=1)
    return d.strftime("%Y%m%d")


def date_n_days_ago(n: int, from_date: str) -> str:
    base = datetime.strptime(from_date, "%Y%m%d")
    return (base - timedelta(days=n * 2)).strftime("%Y%m%d")  # 주말 여유


def _make_progress(label_width: int = 16) -> "Progress":
    """rich Progress 컨텍스트 매니저 반환 (rich 미설치 시 no-op 더미)."""
    if HAS_RICH:
        return Progress(
            SpinnerColumn(),
            TextColumn("{task.description}", table_column=RichColumn(width=label_width, no_wrap=True)),
            BarColumn(),
            MofNCompleteColumn(),
            TimeElapsedColumn(),
            TextColumn("남은"),
            TimeRemainingColumn(),
            console=console,
            transient=False,
        )

    class _Dummy:
        def __enter__(self): return self
        def __exit__(self, *a): pass
        def add_task(self, *a, **kw): return 0
        def update(self, *a, **kw): pass
        def advance(self, *a, **kw): pass

    return _Dummy()


def safe_call(func, *args, retries: int = 3, delay: float = 0.3, **kwargs):
    """재시도 포함 안전한 함수 호출"""
    for attempt in range(retries):
        try:
            result = func(*args, **kwargs)
            time.sleep(delay)
            return result
        except Exception as exc:
            log.warning(f"  재시도 {attempt+1}/{retries} [{func.__name__}]: {exc}")
            time.sleep(1.5)
    return None


def ensure_dir(path: str) -> Path:
    p = Path(path)
    p.mkdir(parents=True, exist_ok=True)
    return p


# ═══════════════════════════════════════════════════════════════
# STEP 1 — 데이터 수집
# ═══════════════════════════════════════════════════════════════

# KOSPI 업종 코드 매핑 (pykrx 1.2.7 / 실제 KRX 기준)
KOSPI_SECTOR_CODES: dict[str, str] = {
    "1005": "음식료·담배",  "1006": "섬유·의류",    "1007": "종이·목재",
    "1008": "화학",          "1009": "제약",          "1010": "비금속",
    "1011": "금속",          "1012": "기계·장비",     "1013": "전기전자",
    "1014": "의료·정밀기기", "1015": "운송장비·부품", "1016": "유통",
    "1017": "전기·가스",     "1018": "건설",          "1019": "운송·창고",
    "1020": "통신",          "1021": "금융",          "1024": "증권",
    "1025": "보험",          "1026": "일반서비스",     "1027": "제조",
}


def fetch_sector_performance(target_date: str, cfg: Config) -> pd.DataFrame:
    """[STEP 1-A] 업종별 등락률·장중변동폭 수집"""
    delay = cfg.get("api", "delay_sec", default=0.3)
    rows  = []
    items = list(KOSPI_SECTOR_CODES.items())

    with _make_progress() as progress:
        task = progress.add_task("[1-A] 섹터 수집", total=len(items))
        for code, name in items:
            progress.update(task, description=f"[1-A] {name[:8]:<10}")
            df = safe_call(
                krx.get_index_ohlcv_by_date,
                fromdate=target_date, todate=target_date, ticker=code,
                name_display=False, delay=delay,
            )
            if df is not None and not df.empty:
                r = df.iloc[-1]
                open_p  = r.get("시가", 0)
                close_p = r.get("종가", 0)
                high_p  = r.get("고가", close_p)
                low_p   = r.get("저가", close_p)
                rows.append({
                    "섹터코드":      code,
                    "섹터명":        name,
                    "시가":          open_p,
                    "종가":          close_p,
                    "등락률(%)":     round((close_p - open_p) / open_p * 100, 2) if open_p else 0,
                    "장중변동폭(%)": round((high_p - low_p) / low_p * 100, 2) if low_p else 0,
                })
            progress.advance(task)

    if not rows:
        log.warning("   → 업종 데이터 없음 (공휴일 또는 네트워크 오류)")
        return pd.DataFrame(columns=["섹터코드", "섹터명", "시가", "종가", "등락률(%)", "장중변동폭(%)"])
    df_out = pd.DataFrame(rows).sort_values("등락률(%)", ascending=False).reset_index(drop=True)
    log.info(f"   → {len(df_out)}개 업종 수집 완료")
    return df_out


def fetch_all_stocks(target_date: str, cfg: Config) -> pd.DataFrame:
    """[STEP 1-B] KOSPI + KOSDAQ 전종목 OHLCV + 시가총액"""
    log.info("[1-B] 전종목 주가 데이터 수집 중...")
    markets     = cfg.get("universe", "markets", default=["KOSPI", "KOSDAQ"])
    min_vol     = cfg.get("universe", "min_volume_billion", default=50)
    min_cap     = cfg.get("universe", "min_cap_billion", default=500)
    delay       = cfg.get("api", "delay_sec", default=0.3)
    frames = []

    for market in markets:
        df = safe_call(krx.get_market_ohlcv_by_ticker, target_date, market=market, delay=delay)
        if df is None or df.empty:
            log.warning(f"   [{market}] OHLCV 없음")
            continue

        df_cap = safe_call(krx.get_market_cap_by_ticker, target_date, market=market, delay=delay)
        if df_cap is not None and not df_cap.empty:
            # 시가총액은 get_market_ohlcv_by_ticker에 이미 포함 → 중복 제외
            cols_to_join = [c for c in ["시가총액", "상장주식수"]
                            if c in df_cap.columns and c not in df.columns]
            if cols_to_join:
                df = df.join(df_cap[cols_to_join], how="left")

        df["시장"]  = market
        df["티커"]  = df.index
        df["종목명"] = [
            safe_call(krx.get_market_ticker_name, t, delay=0) or t
            for t in df.index
        ]
        frames.append(df)
        log.info(f"   → [{market}] {len(df)}개 종목")

    if not frames:
        return pd.DataFrame()

    df_all = pd.concat(frames)
    df_all["거래대금(억)"]  = df_all.get("거래대금",  pd.Series(dtype=float)) / 1e8
    df_all["시가총액(억)"]  = df_all.get("시가총액",  pd.Series(dtype=float)) / 1e8

    df_all = df_all[
        (df_all["거래대금(억)"]  >= min_vol) &
        (df_all["시가총액(억)"]  >= min_cap)
    ].copy()

    df_all["등락률(%)"] = (
        (df_all["종가"] - df_all["시가"]) / df_all["시가"] * 100
    ).round(2).replace([np.inf, -np.inf], 0)

    log.info(f"   → 조건 통과: {len(df_all)}개 종목")
    return df_all


def fetch_volume_ma(tickers: list[str], target_date: str, cfg: Config) -> pd.Series:
    """[STEP 1-C] 20거래일 평균 거래대금 계산"""
    lookback = cfg.get("date", "lookback_days", default=30)
    delay    = cfg.get("api", "delay_sec", default=0.3)
    start    = date_n_days_ago(lookback, target_date)

    avg: dict[str, float] = {}
    with _make_progress() as progress:
        task = progress.add_task("[1-C] 거래대금 MA20", total=len(tickers))
        for ticker in tickers:
            progress.update(task, description=f"[1-C] {ticker}")
            df = safe_call(
                krx.get_market_ohlcv_by_date,
                fromdate=start, todate=target_date, ticker=ticker,
                delay=delay,
            )
            if df is not None and not df.empty:
                if "거래대금" in df.columns:
                    avg[ticker] = df["거래대금"].mean() / 1e8
                elif "종가" in df.columns and "거래량" in df.columns:
                    # pykrx 1.2.7: 거래대금 없음 → 종가×거래량으로 근사
                    avg[ticker] = (df["종가"] * df["거래량"]).mean() / 1e8
                else:
                    avg[ticker] = 0.0
            else:
                avg[ticker] = 0.0
            progress.advance(task)

    return pd.Series(avg, name="거래대금_MA20(억)")


def fetch_52w_high(tickers: list[str], target_date: str, cfg: Config) -> pd.Series:
    """[STEP 1-D] 52주 신고가 계산"""
    delay = cfg.get("api", "delay_sec", default=0.3)
    start = date_n_days_ago(250, target_date)   # 약 52주

    high52: dict[str, float] = {}
    with _make_progress() as progress:
        task = progress.add_task("[1-D] 52주 신고가", total=len(tickers))
        for ticker in tickers:
            progress.update(task, description=f"[1-D] {ticker}")
            df = safe_call(
                krx.get_market_ohlcv_by_date,
                fromdate=start, todate=target_date, ticker=ticker,
                delay=delay,
            )
            if df is not None and "고가" in df.columns:
                high52[ticker] = df["고가"].max()
            else:
                high52[ticker] = 0.0
            progress.advance(task)

    return pd.Series(high52, name="52주최고가")


def get_sector_ticker_map(cfg: Config) -> dict[str, list[str]]:
    """[STEP 1-E] 업종 코드별 소속 종목 목록"""
    delay = cfg.get("api", "delay_sec", default=0.3)
    sector_map: dict[str, list[str]] = {}
    items = list(KOSPI_SECTOR_CODES.items())

    with _make_progress() as progress:
        task = progress.add_task("[1-E] 섹터 매핑", total=len(items))
        for code, name in items:
            progress.update(task, description=f"[1-E] {name[:8]:<10}")
            tickers = safe_call(krx.get_index_portfolio_deposit_file, code, delay=delay)
            if tickers:
                sector_map[name] = list(tickers)
            progress.advance(task)

    total = sum(len(v) for v in sector_map.values())
    log.info(f"   → {len(sector_map)}개 섹터, {total}개 종목 매핑")
    return sector_map


# ═══════════════════════════════════════════════════════════════
# STEP 2 — 기술적 지표 계산
# ═══════════════════════════════════════════════════════════════

def calc_rsi(ticker: str, target_date: str, cfg: Config) -> float:
    """RSI 계산 (pandas_ta 사용)"""
    if not HAS_TA:
        return 50.0  # 기본값 반환

    period = cfg.get("leader_filter", "rsi_period", default=14)
    delay  = cfg.get("api", "delay_sec", default=0.3)
    start  = date_n_days_ago(period * 3, target_date)

    df = safe_call(
        krx.get_market_ohlcv_by_date,
        fromdate=start, todate=target_date, ticker=ticker,
        delay=delay,
    )
    if df is None or df.empty or "종가" not in df.columns:
        return 50.0

    rsi_series = ta.rsi(df["종가"], length=period)
    if rsi_series is None or rsi_series.dropna().empty:
        return 50.0
    return round(float(rsi_series.dropna().iloc[-1]), 1)


# ═══════════════════════════════════════════════════════════════
# STEP 3 — 핫 스코어 + 주도주 선별
# ═══════════════════════════════════════════════════════════════

def calc_hot_score(df: pd.DataFrame, cfg: Config) -> pd.DataFrame:
    """
    [STEP 3-A] 종목별 종합 핫 스코어 계산
    TODO:SCORE — 가중치 조정 또는 ML 스코어링으로 교체 가능
    """
    w_ret  = cfg.get("scoring", "weight_return",       default=0.35)
    w_vol  = cfg.get("scoring", "weight_volume_surge", default=0.35)
    w_mom  = cfg.get("scoring", "weight_momentum",     default=0.20)
    w_news = cfg.get("scoring", "weight_news",         default=0.10)

    df = df.copy()

    # 백분위 순위 (0~1)로 정규화
    df["rank_return"] = df["등락률(%)"].rank(pct=True)

    if "거래대금배율" in df.columns:
        df["rank_volume"] = df["거래대금배율"].rank(pct=True)
    else:
        df["rank_volume"] = 0.5

    if "52주근접도" in df.columns:
        df["rank_momentum"] = df["52주근접도"].rank(pct=True)
    else:
        df["rank_momentum"] = 0.5

    # TODO:SCORE 뉴스 스코어 추가 시 이 부분을 교체
    df["rank_news"] = 0.5

    df["hot_score"] = (
        df["rank_return"]   * w_ret  +
        df["rank_volume"]   * w_vol  +
        df["rank_momentum"] * w_mom  +
        df["rank_news"]     * w_news
    ).round(4)

    return df.sort_values("hot_score", ascending=False)


def filter_leaders(
    df_stocks: pd.DataFrame,
    sector_map: dict[str, list[str]],
    df_sector: pd.DataFrame,
    cfg: Config,
    target_date: Optional[str] = None,
    use_rsi: bool = False,
) -> pd.DataFrame:
    """
    [STEP 3-B] 섹터별 주도주 필터링
    TODO:FILTER — 수급(외국인/기관), MACD, 신고가 돌파 조건 추가 가능
    """
    rsi_min        = cfg.get("leader_filter", "rsi_min",            default=50)
    rsi_max        = cfg.get("leader_filter", "rsi_max",            default=75)
    high52_ratio   = cfg.get("leader_filter", "high52w_ratio_min",  default=0.80)
    vol_surge_min  = cfg.get("leader_filter", "volume_surge_min",   default=2.0)
    top_n_sectors  = cfg.get("filter",        "top_n_sectors",      default=5)
    top_n_stocks   = cfg.get("filter",        "top_n_stocks",       default=3)

    log.info("[3-B] 섹터별 주도주 필터링 중...")

    hot_sectors = df_sector.nlargest(top_n_sectors, "등락률(%)")
    leader_rows = []

    for _, s_row in hot_sectors.iterrows():
        sector_name = s_row["섹터명"]
        tickers_in  = sector_map.get(sector_name, [])
        if not tickers_in:
            continue

        df_sub = df_stocks[df_stocks["티커"].isin(tickers_in)].copy()
        if df_sub.empty:
            continue

        # ── 1차 필터: 거래대금 급증 ──────────────────────────
        if "거래대금배율" in df_sub.columns:
            df_sub = df_sub[df_sub["거래대금배율"] >= vol_surge_min]

        # ── 2차 필터: 52주 신고가 근접 ───────────────────────
        if "52주근접도" in df_sub.columns:
            df_sub = df_sub[df_sub["52주근접도"] >= high52_ratio]

        # ── 3차 필터: RSI ──────────────────────────────────────
        if use_rsi and HAS_TA and target_date:
            df_sub = df_sub[
                df_sub["티커"].apply(
                    lambda t: rsi_min <= calc_rsi(t, target_date, cfg) <= rsi_max
                )
            ]

        if df_sub.empty:
            df_sub = df_stocks[df_stocks["티커"].isin(tickers_in)].copy()  # 필터 완화

        # 핫 스코어 기준 TOP N
        leaders = df_sub.nlargest(top_n_stocks, "hot_score")
        for _, st_row in leaders.iterrows():
            leader_rows.append({
                "섹터명":         sector_name,
                "섹터등락률(%)":  s_row["등락률(%)"],
                "티커":           st_row["티커"],
                "종목명":         st_row.get("종목명", st_row["티커"]),
                "종가":           st_row.get("종가", 0),
                "등락률(%)":      st_row["등락률(%)"],
                "거래대금(억)":   round(st_row.get("거래대금(억)", 0), 1),
                "거래대금배율":   round(st_row.get("거래대금배율", 0), 2),
                "52주근접도":     round(st_row.get("52주근접도", 0), 3),
                "hot_score":      st_row.get("hot_score", 0),
                "시가총액(억)":   round(st_row.get("시가총액(억)", 0), 0),
                "시장":           st_row.get("시장", ""),
            })

    return pd.DataFrame(leader_rows)


# ═══════════════════════════════════════════════════════════════
# Excel 서식 정의 (4_financial_scanner 스타일)
# ═══════════════════════════════════════════════════════════════

_FMT_PCT_SIGN = '[Color10]+#,##0.00;[Red]-#,##0.00;0.00'

_COL_FORMATS_LEADER: dict[str, str] = {
    "hot_score":      "#,##0.0000",
    "종가":           "#,##0",
    "등락률(%)":      _FMT_PCT_SIGN,
    "섹터등락률(%)":  _FMT_PCT_SIGN,
    "거래대금(억)":   "#,##0.0",
    "거래대금배율":   "#,##0.00",
    "52주근접도":     "#,##0.00",
    "시가총액(억)":   "#,##0",
}

_HEADER_FILLS_LEADER: dict[str, str] = {
    "섹터명":         "1F4E79",
    "섹터등락률(%)":  "1F4E79",
    "티커":           "1F4E79",
    "종목명":         "1F4E79",
    "시장":           "1F4E79",
    "hot_score":      "7B3F00",
    "종가":           "375623",
    "등락률(%)":      "375623",
    "거래대금(억)":   "7F4F24",
    "거래대금배율":   "7F4F24",
    "52주근접도":     "1F3864",
    "시가총액(억)":   "4C4C4C",
}

_COL_FORMATS_SECTOR: dict[str, str] = {
    "등락률(%)":      _FMT_PCT_SIGN,
    "장중변동폭(%)":  "#,##0.00",
    "시가":           "#,##0",
    "종가":           "#,##0",
}

_HEADER_FILLS_SECTOR: dict[str, str] = {
    "섹터코드":       "1F4E79",
    "섹터명":         "1F4E79",
    "등락률(%)":      "375623",
    "장중변동폭(%)":  "7F4F24",
    "시가":           "4C4C4C",
    "종가":           "4C4C4C",
}

_LEADER_COLS = [
    "섹터명", "섹터등락률(%)", "티커", "종목명", "시장",
    "종가", "등락률(%)", "거래대금(억)", "거래대금배율",
    "52주근접도", "hot_score", "시가총액(억)",
]


def _apply_excel_styles(wb) -> None:
    """워크북 전체에 헤더 색상·숫자 포맷·하이퍼링크·열 너비·필터 적용."""
    from openpyxl.styles import Alignment, Font, PatternFill

    def _cw(val):
        return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

    for ws in wb.worksheets:
        use_sector   = (ws.title == "섹터등락률")
        col_formats  = _COL_FORMATS_SECTOR  if use_sector else _COL_FORMATS_LEADER
        header_fills = _HEADER_FILLS_SECTOR if use_sector else _HEADER_FILLS_LEADER

        col_fmt_map: dict[int, str] = {}
        name_col = ticker_col = None

        for cell in ws[1]:
            col_name = cell.value or ""
            fill_hex = header_fills.get(col_name, "1F4E79")
            cell.fill      = PatternFill("solid", fgColor=fill_hex)
            cell.font      = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            if col_name in col_formats:
                col_fmt_map[cell.column] = col_formats[col_name]
            if col_name == "종목명": name_col   = cell.column
            if col_name == "티커":   ticker_col = cell.column

        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")
                fmt = col_fmt_map.get(cell.column)
                if fmt and isinstance(cell.value, (int, float)):
                    cell.number_format = fmt
                elif isinstance(cell.value, int):
                    cell.number_format = "#,##0"

        if name_col and ticker_col:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                nc = row[name_col   - 1]
                tc = row[ticker_col - 1]
                if tc.value:
                    code6 = str(tc.value).zfill(6)
                    nc.hyperlink = f"https://finance.naver.com/item/main.naver?code={code6}"
                    nc.font = Font(color="0563C1", underline="single", bold=True)
                    tc.hyperlink = (
                        f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                        f"?pGB=1&gicode=A{code6}&cID=AA"
                        f"&MenuYn=Y&ReportGB=&NewMenuID=11&stkGb="
                    )
                    tc.font = Font(color="0563C1", underline="single")

        for col in ws.columns:
            header_w = _cw(col[0].value) + 3
            data_w   = max((_cw(cell.value) for cell in col[1:]), default=0) + 2
            ws.column_dimensions[col[0].column_letter].width = min(max(header_w, data_w, 10), 40)

        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions


def _write_leader_sheets(df_sector: pd.DataFrame, df_leaders: pd.DataFrame, writer) -> None:
    """ExcelWriter에 종합 + 섹터별 + 섹터등락률 시트 구성."""
    avail = [c for c in _LEADER_COLS if c in df_leaders.columns]
    df_s  = df_leaders.sort_values("hot_score", ascending=False).reindex(columns=avail)

    df_s.to_excel(writer, sheet_name="종합", index=False)

    for sector in df_s["섹터명"].dropna().unique():
        sname = (
            str(sector)[:31]
            .replace("/", "_").replace("\\", "_")
            .replace("*", "").replace(":", "").replace("?", "")
            .replace("[", "").replace("]", "")
            or "기타"
        )
        df_s[df_s["섹터명"] == sector].to_excel(writer, sheet_name=sname, index=False)

    df_sector.to_excel(writer, sheet_name="섹터등락률", index=False)


def build_excel_bytes(df_sector: pd.DataFrame, df_leaders: pd.DataFrame) -> bytes:
    """Streamlit 다운로드용: 메모리에서 서식 적용된 Excel bytes 반환."""
    import io
    from openpyxl import load_workbook

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        _write_leader_sheets(df_sector, df_leaders, writer)
    buf.seek(0)

    wb = load_workbook(buf)
    _apply_excel_styles(wb)

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ═══════════════════════════════════════════════════════════════
# STEP 4 — 출력 · 저장 · 알림
# ═══════════════════════════════════════════════════════════════

def print_report(df_sector: pd.DataFrame, df_leaders: pd.DataFrame,
                 target_date: str, cfg: Config) -> None:
    """콘솔 리포트 출력 (rich 설치 시 컬러 테이블)"""
    top_n    = cfg.get("filter", "top_n_sectors", default=5)
    dt_str   = f"{target_date[:4]}-{target_date[4:6]}-{target_date[6:]}"

    if not HAS_RICH:
        # ── fallback: 기존 plain-text 출력 ─────────────────────
        sep = "=" * 64
        print(f"\n{sep}\n  🔥 주도 섹터 · 주도주 리포트  기준일: {dt_str}\n{sep}")
        for i, row in enumerate(df_sector.nlargest(top_n, "등락률(%)").itertuples(), 1):
            pct = row._5
            print(f"  {i:2}. {row.섹터명:<12}  {pct:+.2f}%")
        if not df_leaders.empty:
            prev = None
            for _, r in df_leaders.iterrows():
                if r["섹터명"] != prev:
                    print(f"\n  ▶ {r['섹터명']}  (섹터 {r['섹터등락률(%)']:+.2f}%)")
                    prev = r["섹터명"]
                print(f"      {r['티커']}  {r['종목명']:<12}  {r['등락률(%)']:+.2f}%"
                      f"  {r['거래대금(억)']:.0f}억  score={r['hot_score']:.3f}")
        print(f"\n{sep}\n")
        return

    # ── rich 출력 ────────────────────────────────────────────────
    console.rule(f"[bold cyan]🔥 주도 섹터 · 주도주 리포트[/bold cyan]  [dim]{dt_str}[/dim]")
    if not HAS_TA:
        console.print("[yellow]⚠  pandas_ta 미설치 → RSI 필터 비활성[/yellow]")

    # 섹터 테이블
    tbl_s = Table(box=rich_box.SIMPLE_HEAD, header_style="bold cyan",
                  border_style="dim", title=f"[bold]HOT 섹터 TOP {top_n}[/bold]")
    tbl_s.add_column("순위", width=4,  justify="right", style="dim")
    tbl_s.add_column("섹터명",    width=14)
    tbl_s.add_column("등락률(%)", width=10, justify="right")

    for i, row in enumerate(df_sector.nlargest(top_n, "등락률(%)").itertuples(), 1):
        pct = row._5
        pct_str = f"[red]{pct:+.2f}%[/red]" if pct >= 0 else f"[blue]{pct:+.2f}%[/blue]"
        tbl_s.add_row(str(i), row.섹터명, pct_str)

    console.print()
    console.print(tbl_s)

    # 주도주 테이블
    if df_leaders.empty:
        console.print("[dim]  (주도주 없음 — 필터 조건 또는 섹터 매핑 확인)[/dim]\n")
        return

    tbl_l = Table(box=rich_box.SIMPLE_HEAD, header_style="bold cyan",
                  border_style="dim", title="[bold]섹터별 주도주[/bold]")
    tbl_l.add_column("섹터명",      width=14)
    tbl_l.add_column("종목명",      width=14)
    tbl_l.add_column("티커",        width=8,  style="dim")
    tbl_l.add_column("등락률(%)",   width=10, justify="right")
    tbl_l.add_column("거래대금(억)", width=11, justify="right")
    tbl_l.add_column("거래대금배율", width=10, justify="right")
    tbl_l.add_column("52W근접도",   width=9,  justify="right")
    tbl_l.add_column("hot_score",   width=10, justify="right", style="bold yellow")
    tbl_l.add_column("시장",        width=7)

    for _, r in df_leaders.iterrows():
        pct = r["등락률(%)"]
        pct_str = f"[red]{pct:+.2f}%[/red]" if pct >= 0 else f"[blue]{pct:+.2f}%[/blue]"
        tbl_l.add_row(
            r["섹터명"],
            r["종목명"],
            r["티커"],
            pct_str,
            f"{r['거래대금(억)']:.0f}",
            f"{r.get('거래대금배율', 0):.1f}x",
            f"{r.get('52주근접도', 0)*100:.0f}%",
            f"{r['hot_score']:.4f}",
            r.get("시장", ""),
        )

    console.print(tbl_l)
    console.print(
        f"[cyan]통과 종목: {len(df_leaders)}개  |  섹터: {df_leaders['섹터명'].nunique()}개[/cyan]\n"
    )


def save_report(df_sector: pd.DataFrame, df_leaders: pd.DataFrame,
                target_date: str, cfg: Config) -> Optional[Path]:
    """CSV / Excel 저장"""
    out_dir    = ensure_dir(cfg.get("output", "output_dir", default="./reports"))
    prefix     = cfg.get("output", "filename_prefix", default="sector_report")
    save_csv   = cfg.get("output", "save_csv",   default=True)
    save_excel = cfg.get("output", "save_excel", default=False)

    ts   = datetime.now().strftime("%Y_%m%d_%H%M")
    base = out_dir / f"{prefix}_{target_date}_{ts}"

    if save_csv:
        path = base.with_suffix(".csv")
        pd.concat([
            df_sector.assign(구분="섹터"),
            df_leaders.assign(구분="주도주"),
        ]).to_csv(path, index=False, encoding="utf-8-sig")
        log.info(f"  💾 CSV 저장: {path}")

    if save_excel:
        try:
            from openpyxl import load_workbook
            path = base.with_suffix(".xlsx")
            with pd.ExcelWriter(path, engine="openpyxl") as writer:
                _write_leader_sheets(df_sector, df_leaders, writer)
            wb = load_workbook(path)
            _apply_excel_styles(wb)
            wb.save(path)
            log.info(f"  💾 Excel 저장: {path}")
            return path
        except Exception as e:
            log.warning(f"  Excel 저장 실패 (openpyxl 필요): {e}")

    return None


def send_email_report(df_sector: pd.DataFrame, df_leaders: pd.DataFrame,
                      target_date: str, cfg: Config) -> None:
    """[STEP 4-EMAIL] 주도 섹터 · 주도주 결과를 이메일로 발송."""
    import smtplib
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    gmail_user = os.getenv("GMAIL_USER", "")
    gmail_pw   = os.getenv("GMAIL_APP_PASSWORD", "")
    if not gmail_user or not gmail_pw:
        log.warning("이메일 미설정 (.env: GMAIL_USER / GMAIL_APP_PASSWORD)")
        return

    _extra_recipients = [e.strip() for e in _re.split(r"[,;]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
    _main_recipients  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
    recipients        = _main_recipients + _extra_recipients
    if not recipients:
        log.warning("수신자 미설정 (.env: NOTIFY_EMAIL / NOTIFY_EMAIL_WIFE / NOTIFY_EMAILS_EXTRA)")
        return

    dt_str   = f"{target_date[:4]}-{target_date[4:6]}-{target_date[6:]}"
    top_n    = cfg.get("filter", "top_n_sectors", default=5)
    top_sect = df_sector.nlargest(top_n, "등락률(%)")

    sect_lines = []
    for i, row in enumerate(top_sect.itertuples(), 1):
        sect_lines.append(f"  {i}. {row.섹터명:<12}  {row._5:+.2f}%")

    leader_lines = []
    prev_sect = None
    for _, r in df_leaders.iterrows():
        if r["섹터명"] != prev_sect:
            leader_lines.append(f"\n  ▶ {r['섹터명']}  (섹터 {r['섹터등락률(%)']:+.2f}%)")
            prev_sect = r["섹터명"]
        surge = f"  거래대금×{r['거래대금배율']:.1f}" if r.get("거래대금배율", 0) > 0 else ""
        leader_lines.append(
            f"      {r['티커']}  {r['종목명']:<10}"
            f"  {r['등락률(%)']:+.2f}%"
            f"  {r['거래대금(억)']:.0f}억"
            f"{surge}"
            f"  score={r['hot_score']:.3f}"
        )

    body = "\n".join([
        f"안녕하세요,",
        f"",
        f"[{dt_str}] 주도 섹터 · 주도주 스캔이 완료되었습니다.",
        f"",
        f"■ HOT 섹터 TOP {top_n}",
        *sect_lines,
        f"",
        f"■ 섹터별 주도주",
        *leader_lines,
        f"",
        f"※ 이 메일은 자동 발송입니다.",
    ])

    msg = MIMEMultipart()
    msg["From"]    = gmail_user
    msg["To"]      = ", ".join(_main_recipients) if _main_recipients else gmail_user
    if _extra_recipients:
        msg["Bcc"] = ", ".join(_extra_recipients)
    msg["Subject"] = f"[주도주] {dt_str} 스캔 완료 — HOT {df_sector.nlargest(1,'등락률(%)').iloc[0]['섹터명']}"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail_user, gmail_pw)
            smtp.sendmail(gmail_user, recipients, msg.as_string())
        _to_str  = ', '.join(_main_recipients) if _main_recipients else gmail_user
        _bcc_str = f" | Bcc: {len(_extra_recipients)}명" if _extra_recipients else ""
        log.info(f"  이메일 발송 완료 — To: {_to_str}{_bcc_str}")
        if _extra_recipients:
            log.info(f"    Bcc: {', '.join(_extra_recipients)}")
    except Exception as e:
        log.warning(f"  이메일 발송 실패: {e}")


async def send_telegram_alert(df_leaders: pd.DataFrame, cfg: Config) -> None:
    """
    [STEP 4-ALERT] 텔레그램으로 주도주 알림 전송
    TODO:ALERT — 메시지 포맷, 차트 이미지 첨부 등 커스터마이징 가능
    """
    if not HAS_TELEGRAM:
        log.warning("python-telegram-bot 미설치 → 알림 스킵")
        return

    token   = cfg.get("alert", "telegram_token",   default="")
    chat_id = cfg.get("alert", "telegram_chat_id", default="")
    top_n   = cfg.get("alert", "alert_top_n",      default=5)

    if not token or not chat_id:
        log.warning("텔레그램 token / chat_id 설정 필요")
        return

    lines = ["🔥 *주도주 알림*\n"]
    for i, (_, r) in enumerate(df_leaders.head(top_n).iterrows(), 1):
        lines.append(
            f"{i}. [{r['섹터명']}] `{r['티커']}` {r['종목명']} "
            f"{r['등락률(%)']:+.2f}%  {r['거래대금(억)']:.0f}억"
        )

    bot = telegram.Bot(token=token)
    async with bot:
        await bot.send_message(
            chat_id=chat_id,
            text="\n".join(lines),
            parse_mode="Markdown",
        )
    log.info(f"  📬 텔레그램 알림 전송 완료 ({len(lines)-1}건)")


# ═══════════════════════════════════════════════════════════════
# 메인 파이프라인
# ═══════════════════════════════════════════════════════════════

def run_pipeline(cfg: Config, target_date: Optional[str] = None) -> dict:
    """
    전체 분석 파이프라인 실행
    반환값: {"sector": df_sector, "leaders": df_leaders}
    """
    if not HAS_PYKRX:
        log.error("pykrx 미설치 → pip install pykrx")
        sys.exit(1)

    t0 = time.time()
    raw_date = target_date or cfg.get("date", "target_date") or None
    date = get_recent_business_day(
        datetime.strptime(raw_date, "%Y%m%d") if raw_date else None
    )
    if raw_date and date != raw_date:
        log.info(f"  ⚠  {raw_date}는 거래일이 아님 → 직전 거래일 {date}로 조정")
    log.info(f"\n🚀 분석 시작 | 기준일: {date}")

    # ── STEP 1: 데이터 수집 ──────────────────────────────────
    df_sector = fetch_sector_performance(date, cfg)
    df_stocks = fetch_all_stocks(date, cfg)

    if df_stocks.empty:
        log.error("전종목 주가 수집 실패 — 날짜/네트워크 확인")
        return {}

    # 거래대금 이동평균 (상위 N종목)
    vol_top_n  = cfg.get("filter", "vol_ma_top_n", default=200)
    top_tickers = df_stocks.nlargest(vol_top_n, "거래대금(억)")["티커"].tolist()

    vol_ma  = fetch_volume_ma(top_tickers, date, cfg)
    high52  = fetch_52w_high(top_tickers, date, cfg)

    # df_stocks에 join
    df_stocks = df_stocks.set_index("티커", drop=False)
    df_stocks = df_stocks.join(vol_ma,  how="left")
    df_stocks = df_stocks.join(high52,  how="left")
    df_stocks = df_stocks.reset_index(drop=True)

    # 파생 컬럼 계산
    df_stocks["거래대금배율"] = (
        df_stocks["거래대금(억)"] / df_stocks["거래대금_MA20(억)"]
    ).replace([np.inf, -np.inf], 0).fillna(0).round(2)

    df_stocks["52주근접도"] = (
        df_stocks["종가"] / df_stocks["52주최고가"]
    ).replace([np.inf, -np.inf], 0).fillna(0).round(4)

    # 섹터-종목 매핑
    sector_map = get_sector_ticker_map(cfg)

    # ── STEP 2+3: 스코어 계산 + 주도주 선별 ─────────────────
    df_stocks  = calc_hot_score(df_stocks, cfg)
    df_leaders = filter_leaders(df_stocks, sector_map, df_sector, cfg, target_date=date)

    # ── STEP 4: 출력 · 저장 ──────────────────────────────────
    print_report(df_sector, df_leaders, date, cfg)

    if cfg.get("output", "save_csv", default=True):
        save_report(df_sector, df_leaders, date, cfg)

    # 텔레그램 알림
    if cfg.get("alert", "telegram_enabled", default=False) and not df_leaders.empty:
        import asyncio
        asyncio.run(send_telegram_alert(df_leaders, cfg))

    # 이메일 발송
    if cfg.get("alert", "email_enabled", default=False) and not df_leaders.empty:
        send_email_report(df_sector, df_leaders, date, cfg)

    elapsed = time.time() - t0
    log.info(f"  ⏱  완료: {elapsed:.1f}초\n")

    return {"sector": df_sector, "leaders": df_leaders}


# ═══════════════════════════════════════════════════════════════
# 스케줄러 (장중 자동 실행)
# ═══════════════════════════════════════════════════════════════

def run_scheduler(cfg: Config) -> None:
    """
    APScheduler 기반 장중 자동 실행
    TODO:VISUAL Streamlit 대시보드와 연동 시 이 함수를 백그라운드 스레드로 실행
    """
    if not HAS_SCHEDULER:
        log.error("APScheduler 미설치 → pip install apscheduler")
        sys.exit(1)

    run_times = cfg.get("scheduler", "run_times", default=["09:00", "12:00", "15:30"])
    scheduler = BlockingScheduler(timezone="Asia/Seoul")

    for t in run_times:
        h, m = map(int, t.split(":"))
        scheduler.add_job(
            run_pipeline,
            trigger="cron",
            hour=h, minute=m,
            args=[cfg],
            id=f"scan_{t}",
        )
        log.info(f"  📅 등록: {t} 실행")

    log.info("⏰ 스케줄러 시작 (Ctrl+C로 중지)")
    try:
        scheduler.start()
    except KeyboardInterrupt:
        log.info("스케줄러 중지")


# ═══════════════════════════════════════════════════════════════
# CLI 진입점
# ═══════════════════════════════════════════════════════════════

def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="주도 섹터 · 주도주 자동 탐색 시스템",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--date", "-d", type=str, default=None,
        help="분석 날짜 (YYYYMMDD). 기본값: 오늘 기준 최근 거래일",
    )
    parser.add_argument(
        "--schedule", "-s", action="store_true",
        help="스케줄 모드 활성화 (config.yaml scheduler.run_times 기준)",
    )
    parser.add_argument(
        "--top", "-t", type=int, default=None,
        help="출력할 섹터 TOP N (config.yaml 값 override)",
    )
    parser.add_argument(
        "--config", "-c", type=str, default=None,
        help="config.yaml 경로 override",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    # Config 로드
    override = {}
    if args.top:
        override = {"filter": {"top_n_sectors": args.top}}

    cfg = Config(override=override)
    if args.config:
        Config.CONFIG_PATH = Path(args.config)

    if args.schedule:
        run_scheduler(cfg)
    else:
        run_pipeline(cfg, target_date=args.date)


if __name__ == "__main__":
    main()
