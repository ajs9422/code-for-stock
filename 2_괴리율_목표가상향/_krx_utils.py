"""
공통 KRX 유틸리티 — 0_가치주_찾기.py / 4_전문가추천_저평가종목.py 공용
"""

import os
import sys
import contextlib
import time
import warnings
from datetime import datetime, timedelta

if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

os.environ.setdefault('KRX_ID', 'ajs9422')
os.environ.setdefault('KRX_PW', 'magma7608!')

import pandas as pd
from pykrx import stock

warnings.filterwarnings("ignore")

MARKETS        = ["KOSPI", "KOSDAQ"]
MIN_MARKET_CAP = 300
MIN_ROE        = 3.0
MAX_PER        = 50


# ── 출력 억제 ──────────────────────────────────
@contextlib.contextmanager
def suppress_output():
    with open(os.devnull, 'w') as devnull:
        with contextlib.redirect_stdout(devnull), contextlib.redirect_stderr(devnull):
            yield


# ── 날짜 유틸 ──────────────────────────────────
def get_trading_date(days_ago: int = 0) -> str:
    """영업일 기준 날짜 (YYYYMMDD)"""
    date = datetime.today() - timedelta(days=days_ago)
    while date.weekday() >= 5:
        date -= timedelta(days=1)
    return date.strftime("%Y%m%d")


# ── 종목 사전 ──────────────────────────────────
def build_stock_dicts(markets=None) -> tuple[dict, dict]:
    """
    KRX 전종목 사전 빌드 (3자 이상 종목명만)
    반환: name_to_ticker, ticker_to_name
    """
    if markets is None:
        markets = MARKETS
    today = get_trading_date(0)
    name_to_ticker: dict = {}
    ticker_to_name: dict = {}

    for mkt in markets:
        try:
            with suppress_output():
                tickers = stock.get_market_ticker_list(today, market=mkt)
            for t in tickers:
                with suppress_output():
                    name = stock.get_market_ticker_name(t)
                if name and len(name) >= 3:
                    name_to_ticker[name] = t
                    ticker_to_name[t] = name
        except Exception:
            pass

    return name_to_ticker, ticker_to_name


# ── KRX 기본지표 수집 ──────────────────────────
def fetch_krx_fundamentals(market: str) -> pd.DataFrame:
    today = get_trading_date(0)
    prev5 = get_trading_date(5)
    try:
        with suppress_output():
            df = stock.get_market_fundamental(today, market=market)
        if df.empty:
            with suppress_output():
                df = stock.get_market_fundamental(prev5, market=market)
    except Exception:
        return pd.DataFrame()

    df = df.reset_index()
    df.columns = [str(c) for c in df.columns]

    if "EPS" in df.columns and "BPS" in df.columns:
        bps = df["BPS"].replace(0, float("nan"))
        df["ROE"] = (df["EPS"] / bps * 100).fillna(0).clip(lower=0)

    try:
        with suppress_output():
            tickers = stock.get_market_ticker_list(today, market=market)
            names = {t: stock.get_market_ticker_name(t) for t in tickers}
        df["종목명"] = df["티커"].map(names)
    except Exception:
        df["종목명"] = df["티커"]

    return df


def fetch_market_cap(market: str) -> pd.DataFrame:
    today = get_trading_date(0)
    prev5 = get_trading_date(5)
    try:
        with suppress_output():
            df = stock.get_market_cap(today, market=market)
        if df.empty:
            with suppress_output():
                df = stock.get_market_cap(prev5, market=market)
    except Exception:
        return pd.DataFrame()
    df = df.reset_index()
    df.columns = [str(c) for c in df.columns]
    return df


# ── 1년치 OHLCV ───────────────────────────────
def fetch_ohlcv_1year(ticker: str) -> pd.DataFrame:
    """1년치 OHLCV, 실패 시 최대 2회 재시도"""
    end   = get_trading_date(0)
    start = get_trading_date(380)
    for attempt in range(3):
        try:
            df = stock.get_market_ohlcv_by_date(start, end, ticker)
            if df is not None and not df.empty:
                return df
        except Exception:
            pass
        time.sleep(1.5 * (attempt + 1))
    return pd.DataFrame()


# ── 점수 정규화 ────────────────────────────────
def score_metric(value, low_good: bool, low_val: float, high_val: float) -> float:
    if pd.isna(value):
        return 50.0
    v = float(value)
    if low_good:
        return min(100, max(0, (high_val - v) / (high_val - low_val) * 100))
    return min(100, max(0, (v - low_val) / (high_val - low_val) * 100))


# ── 저평가 필터링 + 기본 점수 ─────────────────
def build_screener_df(markets=None, console=None) -> pd.DataFrame:
    """
    KRX 기본지표 수집 → 저평가 필터링 → per/pbr/roe 점수 계산
    반환: 티커·종목명·PER·PBR·ROE·시가총액_억·현재가·종합_저평가점수 포함 DataFrame
    """
    if markets is None:
        markets = MARKETS

    def cprint(msg):
        if console:
            console.print(msg)

    dfs_fund, dfs_cap = [], []
    for mkt in markets:
        df_f = fetch_krx_fundamentals(mkt)
        if not df_f.empty:
            df_f["시장"] = mkt
            dfs_fund.append(df_f)
            cprint(f"  [dim]{mkt} 기본지표 {len(df_f):,}개[/dim]")

        df_c = fetch_market_cap(mkt)
        if not df_c.empty:
            dfs_cap.append(df_c)

    df_fund = pd.concat(dfs_fund, ignore_index=True) if dfs_fund else pd.DataFrame()
    df_cap  = pd.concat(dfs_cap,  ignore_index=True) if dfs_cap  else pd.DataFrame()

    if df_fund.empty:
        return pd.DataFrame()

    tc_f = "티커" if "티커" in df_fund.columns else df_fund.columns[0]
    tc_c = "티커" if "티커" in df_cap.columns  else df_cap.columns[0]

    df = pd.merge(
        df_fund,
        df_cap.drop_duplicates(subset=[tc_c])[[tc_c, "시가총액", "종가"]],
        on=tc_f, how="left"
    )

    for c in list(df.columns):
        if c.upper() in ("PER", "PBR", "ROE") and c != c.upper():
            df = df.rename(columns={c: c.upper()})
    for col in ("PER", "PBR", "ROE"):
        if col not in df.columns:
            df[col] = 0.0

    df["시가총액_억"] = df["시가총액"].fillna(0) / 1e8
    before = len(df)

    df = df[df["시가총액_억"] >= MIN_MARKET_CAP].copy()
    df = df[df["PER"]  > 0].copy()
    df = df[df["PBR"]  > 0].copy()
    df = df[df["ROE"]  >= MIN_ROE].copy()
    df = df[df["EPS"]  > 0].copy()
    df = df[df["PER"]  <= MAX_PER].copy()
    df = df[~((df["PER"] > 25) & (df["ROE"] < 5))].copy()
    df = df[~df["종목명"].str.contains("스팩|SPAC", case=False, na=False)].copy()

    cprint(f"  [dim]전체 {before:,}개 → 필터 후 {len(df):,}개 (스팩·적자·실적악화 제외)[/dim]")

    df["per_score"] = df["PER"].apply(lambda x: score_metric(x, True,  3,   40))
    df["pbr_score"] = df["PBR"].apply(lambda x: score_metric(x, True,  0.2,  5))
    df["roe_score"] = df["ROE"].apply(lambda x: score_metric(x, False, 0,   40))
    df["종합_저평가점수"] = (
        df["per_score"] * 0.40 +
        df["pbr_score"] * 0.30 +
        df["roe_score"] * 0.30
    ).round(1)

    df["현재가"] = df["종가"].fillna(0) if "종가" in df.columns else 0
    return df.rename(columns={tc_f: "티커"})
