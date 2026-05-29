"""
================================================
  1. 가치주 찾기 — KRX 전종목 저평가 스크리너
================================================

[탐색 원리]
  KRX에서 코스피·코스닥 전종목의 재무지표(PER·PBR·ROE·EPS)를 수집한 뒤,
  "실적은 좋은데 주가가 싼 종목"을 점수화하여 순위를 매긴다.

[필터 조건]
  ① 시가총액  300억 이상  (너무 작은 종목 제외)
  ② PER       0 초과 ~ 50 이하  (적자·고평가 제외)
  ③ PBR       0 초과  (순자산 대비 주가가 의미 있는 종목)
  ④ ROE       3% 이상  (수익성 최소 기준)
  ⑤ EPS       0 초과  (흑자 기업만)
  ⑥ PER > 25이면서 ROE < 5%인 종목 추가 제외
  ⑦ 스팩(SPAC) 종목 제외

[점수 계산 방식]
  - PER 점수  (40%) : PER이 낮을수록 고점수 (저평가)
  - PBR 점수  (30%) : PBR이 낮을수록 고점수 (자산 대비 저평가)
  - ROE 점수  (30%) : ROE가 높을수록 고점수 (수익성 우수)
  → 세 점수를 가중 합산한 "종합 저평가점수" 기준 상위 종목 추출

[데이터 출처]
  pykrx (KRX 공식) → 실패 시 FinanceDataReader 폴백
================================================
"""

import os
import sys
import contextlib
import time
import warnings
from datetime import datetime, timedelta

# Windows 터미널 UTF-8 설정
if sys.stdout.encoding != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')
if sys.stderr.encoding != 'utf-8':
    sys.stderr.reconfigure(encoding='utf-8')

# KRX 인증 (환경변수 없을 때 기본값 사용)
os.environ.setdefault('KRX_ID', 'ajs9422')
os.environ.setdefault('KRX_PW', 'magma7608!')

import pandas as pd
import requests
from pykrx import stock
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
from rich.table import Column
from rich.table import Table
from rich.text import Text


@contextlib.contextmanager
def suppress_output():
    with open(os.devnull, 'w') as devnull:
        with contextlib.redirect_stdout(devnull), contextlib.redirect_stderr(devnull):
            yield

warnings.filterwarnings("ignore")

console = Console()

# ──────────────────────────────────────────────
# 설정값
# ──────────────────────────────────────────────
TOP_N = 50                          # 스크리닝 결과 출력 개수
MARKETS = ["KOSPI", "KOSDAQ"]       # 분석 대상 시장
MIN_MARKET_CAP = 500                # 최소 시가총액 (억원)
MIN_ROE = 3.0                       # 최소 ROE (%) - 수익성 기준
MAX_PER = 50                        # 최대 PER - 과도한 고평가 제외
SCORE_WEIGHTS = {
    "per_score": 0.25,
    "pbr_score": 0.20,
    "roe_score": 0.20,
    "momentum_score": 0.20,
    "supply_score": 0.15,
}

# 테마별 키워드 매핑
THEMES = {
    "AI/인공지능":          ["AI", "인공지능", "딥러닝", "머신러닝", "LLM", "GPT", "뉴로"],
    "로봇/자동화":          ["로봇", "자동화", "협동로봇", "물류로봇", "드론", "무인"],
    "이차전지":             ["이차전지", "배터리", "전지", "양극재", "음극재", "전해질", "분리막", "리튬"],
    "반도체":               ["반도체", "HBM", "파운드리", "웨이퍼", "낸드", "DRAM", "칩", "패키징"],
    "바이오/제약":          ["바이오", "제약", "헬스케어", "신약", "의약", "임상", "항체", "진단"],
    "방산/우주항공":        ["방산", "방위", "우주", "항공", "미사일", "함정", "레이더", "위성"],
    "수소/신재생에너지":    ["수소", "신재생", "태양광", "풍력", "연료전지", "ESS", "에너지"],
    "전기차/자율주행":      ["전기차", "EV", "자율주행", "충전", "자동차부품", "모빌리티"],
    "K-컨텐츠/엔터":        ["엔터", "콘텐츠", "미디어", "게임", "드라마", "K팝", "웹툰", "OTT"],
    "조선/해운":            ["조선", "해운", "선박", "LNG선", "컨테이너선", "항만"],
    "건설/플랜트":          ["건설", "플랜트", "주택", "시공", "인프라", "건축"],
    "금융/핀테크":          ["금융", "은행", "보험", "핀테크", "증권", "자산운용", "카드"],
}




# ──────────────────────────────────────────────
# 날짜 유틸
# ──────────────────────────────────────────────
def get_trading_date(days_ago=0):
    """영업일 기준 날짜 반환"""
    date = datetime.today() - timedelta(days=days_ago)
    # 주말이면 금요일로 조정
    while date.weekday() >= 5:
        date -= timedelta(days=1)
    return date.strftime("%Y%m%d")


# ──────────────────────────────────────────────
# 데이터 수집
# ──────────────────────────────────────────────
def fetch_krx_fundamentals(market: str) -> pd.DataFrame:
    """KRX에서 PER, PBR, ROE, 시가총액 등 기본 지표 수집"""
    today = get_trading_date(0)
    prev5 = get_trading_date(5)

    console.print(f"  [dim]기준일: {today}[/dim]")

    try:
        df = stock.get_market_fundamental(today, market=market)
        if df.empty:
            df = stock.get_market_fundamental(prev5, market=market)
    except Exception as e:
        console.print(f"  [red]기본 지표 수집 오류: {e}[/red]")
        return pd.DataFrame()

    df = df.reset_index()
    df.columns = [str(c) for c in df.columns]

    # ROE 계산: pykrx fundamental에 ROE 없음 → EPS/BPS 로 역산
    if "EPS" in df.columns and "BPS" in df.columns:
        bps = df["BPS"].replace(0, float("nan"))
        df["ROE"] = (df["EPS"] / bps * 100).fillna(0).clip(lower=0)

    # 종목명 추가
    try:
        tickers = stock.get_market_ticker_list(today, market=market)
        names = {t: stock.get_market_ticker_name(t) for t in tickers}
        df["종목명"] = df["티커"].map(names)
    except Exception:
        df["종목명"] = df["티커"]

    return df


def fetch_market_cap(market: str) -> pd.DataFrame:
    """시가총액 및 현재가 수집"""
    today = get_trading_date(0)
    prev5 = get_trading_date(5)

    try:
        df = stock.get_market_cap(today, market=market)
        if df.empty:
            df = stock.get_market_cap(prev5, market=market)
    except Exception as e:
        console.print(f"  [red]시가총액 수집 오류: {e}[/red]")
        return pd.DataFrame()

    df = df.reset_index()
    df.columns = [str(c) for c in df.columns]
    return df


def _fetch_ohlcv_1year(ticker: str) -> pd.DataFrame:
    """1년치 OHLCV — 실패 시 최대 2회 재시도"""
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


# ──────────────────────────────────────────────
# 지표 계산
# ──────────────────────────────────────────────
def calc_momentum_and_range(ticker: str) -> dict:
    """
    OHLCV 1회 호출로 모멘텀 점수 + 52주 범위를 동시에 계산
    (기존 2회 호출 → 1회로 절감)
    """
    result = {"momentum_score": 50.0, "52주최고": 0, "52주최저": 0, "연간변동폭(%)": 0}
    df = _fetch_ohlcv_1year(ticker)
    if df is None or df.empty:
        return result

    close  = df["종가"]
    volume = df["거래량"]

    # 모멘텀 (20일·60일 수익률 + 거래량 증감)
    if len(close) >= 20:
        ret_20 = (close.iloc[-1] / close.iloc[-20] - 1) * 100
        ret_60 = (close.iloc[-1] / close.iloc[-60] - 1) * 100 if len(close) >= 60 else ret_20
        vol_recent = volume.iloc[-5:].mean()
        vol_past   = volume.iloc[-20:-5].mean()
        vol_ratio  = (vol_recent / vol_past - 1) * 100 if vol_past > 0 else 0
        s20  = min(100, max(0, (ret_20 + 20) / 40 * 100))
        s60  = min(100, max(0, (ret_60 + 30) / 60 * 100))
        svol = min(100, max(0, (vol_ratio + 50) / 100 * 100))
        result["momentum_score"] = round(s20 * 0.4 + s60 * 0.4 + svol * 0.2, 1)

    # 52주 고가/저가/변동폭
    h = int(df["고가"].max())
    l = int(df["저가"].min())
    result["52주최고"]      = h
    result["52주최저"]      = l
    result["연간변동폭(%)"] = round((h - l) / l * 100, 1) if l > 0 else 0

    return result


def calc_supply_score(ticker: str) -> float:
    """
    수급 점수 계산 (0~100)
    - 최근 15일 외국인·기관 합산 순매수 기반 — 실패 시 재시도
    """
    end   = get_trading_date(0)
    start = get_trading_date(15)
    for attempt in range(3):
        try:
            df = stock.get_market_trading_value_by_date(start, end, ticker)
            if df is None or df.empty:
                return 50.0
            net = sum(df[c].sum() for c in df.columns if "외국인" in str(c) or "기관합계" in str(c))
            return round(min(100, max(0, 50 + net / 1e8 / 50 * 30)), 1)
        except Exception:
            time.sleep(1.5 * (attempt + 1))
    return 50.0


def calc_target_gap(per: float, pbr: float, roe: float) -> float:
    """
    목표가 괴리율 추정 (%)
    - 섹터 평균 PER/PBR 대비 할인율 기반 추정
    - 실제 증권사 목표가 없이 펀더멘털로 역산
    """
    # 적정 PER 추정: ROE 기반 (PEG 유사 방식)
    fair_per = max(5, min(30, roe * 1.2)) if roe > 0 else 10
    fair_pbr = max(0.5, roe / 10) if roe > 0 else 1.0

    if per <= 0 or pbr <= 0:
        return 0.0

    gap_per = (fair_per / per - 1) * 100
    gap_pbr = (fair_pbr / pbr - 1) * 100

    return round((gap_per * 0.6 + gap_pbr * 0.4), 1)


def calc_earnings_revision(per: float, roe: float, pbr: float) -> float:
    """
    이익 상향폭 추정 (%)
    - ROE 트렌드 및 PER-PBR 관계 기반 추정
    """
    if per <= 0 or pbr <= 0 or roe <= 0:
        return 0.0

    # 내재 EPS 성장률 추정: ROE × (1 - 배당성향) 근사
    implied_growth = roe * 0.6

    # PER 대비 성장률 비교 (PEG < 1이면 상향 여력)
    peg = per / max(implied_growth, 1)
    revision = (1 / peg - 1) * 20  # 정규화

    return round(max(-20, min(40, revision)), 1)


def score_metric(value, low_good: bool, low_val, high_val) -> float:
    """메트릭을 0~100 점수로 변환"""
    if pd.isna(value):
        return 50.0
    val = float(value)
    if low_good:
        # 낮을수록 좋음 (PER, PBR)
        return min(100, max(0, (high_val - val) / (high_val - low_val) * 100))
    else:
        # 높을수록 좋음 (ROE)
        return min(100, max(0, (val - low_val) / (high_val - low_val) * 100))


# ──────────────────────────────────────────────
# 메인 스크리닝 로직
# ──────────────────────────────────────────────
def run_screener():
    theme_keywords = None  # 전체 종목 분석

    console.print()
    console.print(Panel.fit(
        "[bold cyan]저평가 종목 스크리너[/bold cyan]  [dim]Undervalue Detector · KRX 데이터[/dim]",
        border_style="cyan"
    ))
    console.print()

    # ── 1. 기초 데이터 수집 (코스피 + 코스닥) ──
    dfs_fund, dfs_cap = [], []
    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=36, no_wrap=True)),
        console=console,
        transient=False,
    ) as progress:
        for mkt in MARKETS:
            t1 = progress.add_task(f"[cyan]{mkt} 기본 지표 수집 중...", total=None)
            df_f = fetch_krx_fundamentals(mkt)
            if not df_f.empty:
                df_f["시장"] = mkt
                dfs_fund.append(df_f)
            progress.update(t1, description=f"[green]✓ {mkt} 기본 지표 {len(df_f)}개")

            t2 = progress.add_task(f"[cyan]{mkt} 시가총액 수집 중...", total=None)
            df_c = fetch_market_cap(mkt)
            if not df_c.empty:
                dfs_cap.append(df_c)
            progress.update(t2, description=f"[green]✓ {mkt} 시가총액 {len(df_c)}개")

    df_fund = pd.concat(dfs_fund, ignore_index=True) if dfs_fund else pd.DataFrame()
    df_cap  = pd.concat(dfs_cap,  ignore_index=True) if dfs_cap  else pd.DataFrame()

    if df_fund.empty:
        console.print("[red]데이터 수집 실패. 인터넷 연결 및 날짜를 확인하세요.[/red]")
        return

    # ── 2. 데이터 병합 ──
    ticker_col_fund = "티커" if "티커" in df_fund.columns else df_fund.columns[0]
    ticker_col_cap  = "티커" if "티커" in df_cap.columns  else df_cap.columns[0]

    df_cap_dedup = df_cap.drop_duplicates(subset=[ticker_col_cap])
    df = pd.merge(df_fund, df_cap_dedup[[ticker_col_cap, "시가총액", "종가"]], on=ticker_col_fund, how="left")

    # 컬럼 정리
    col_map = {}
    for c in df.columns:
        if "PER" in c.upper() or c == "PER": col_map[c] = "PER"
        elif "PBR" in c.upper() or c == "PBR": col_map[c] = "PBR"
        elif "ROE" in c.upper() or c == "ROE": col_map[c] = "ROE"
    df = df.rename(columns=col_map)

    for col in ["PER", "PBR", "ROE"]:
        if col not in df.columns:
            df[col] = 0.0

    # ── 3. 필터링 ──
    df["시가총액_억"] = (df["시가총액"].fillna(0) / 1e8) if "시가총액" in df.columns else 0

    before = len(df)
    df = df[df["시가총액_억"] >= MIN_MARKET_CAP].copy()
    df = df[df["PER"] > 0].copy()
    df = df[df["PBR"] > 0].copy()
    df = df[df["ROE"] >= MIN_ROE].copy()        # 최소 ROE 기준
    df = df[df["EPS"] > 0].copy()               # 흑자 종목만 (적자 제외)
    df = df[df["PER"] <= MAX_PER].copy()        # 고PER 제외 (실적 하락 우려)
    # PER 높고 ROE 낮으면 실적 개선 기대 낮음 → 제외
    df = df[~((df["PER"] > 25) & (df["ROE"] < 5))].copy()
    # 스팩(SPAC) 종목 제외
    df = df[~df["종목명"].str.contains("스팩|SPAC", case=False, na=False)].copy()

    # 테마 필터 (선택 시)
    if theme_keywords:
        pattern = "|".join(theme_keywords)
        df = df[df["종목명"].str.contains(pattern, case=False, na=False)].copy()

    console.print(f"  [dim]전체 {before:,}개 → 필터 후 {len(df):,}개 (스팩·적자·실적악화 제외)[/dim]")

    if df.empty:
        console.print("[red]필터링 후 종목이 없습니다.[/red]")
        return

    # ── 4. 점수 계산 ──
    df["per_score"] = df["PER"].apply(lambda x: score_metric(x, True, 3, 40))
    df["pbr_score"] = df["PBR"].apply(lambda x: score_metric(x, True, 0.2, 5))
    df["roe_score"] = df["ROE"].apply(lambda x: score_metric(x, False, 0, 40))

    df["목표가_괴리율"] = df.apply(
        lambda r: calc_target_gap(r["PER"], r["PBR"], r["ROE"]), axis=1
    )
    df["이익_상향폭"] = df.apply(
        lambda r: calc_earnings_revision(r["PER"], r["ROE"], r["PBR"]), axis=1
    )

    console.print()
    candidates = df.copy()
    ticker_col = ticker_col_fund
    tickers = candidates[ticker_col].tolist()

    momentum_scores = {}
    supply_scores = {}
    range_data = {}

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
        task = progress.add_task(f"[cyan]모멘텀·수급·52주 분석 중...", total=len(tickers))
        for ticker in tickers:
            name = candidates[candidates[ticker_col] == ticker]['종목명'].values[0] if '종목명' in candidates.columns else ticker
            progress.update(task, description=f"[cyan]분석 중: {str(name)[:18]:<20}")
            with suppress_output():
                mr = calc_momentum_and_range(ticker)
                supply_scores[ticker] = calc_supply_score(ticker)
            momentum_scores[ticker] = mr["momentum_score"]
            range_data[ticker] = {k: mr[k] for k in ("52주최고", "52주최저", "연간변동폭(%)")}
            progress.advance(task)
            time.sleep(0.3)

    candidates["모멘텀_점수"] = candidates[ticker_col].map(momentum_scores).fillna(50)
    candidates["수급_점수"]   = candidates[ticker_col].map(supply_scores).fillna(50)
    candidates["52주최고"]    = candidates[ticker_col].map(lambda t: range_data.get(t, {}).get("52주최고", 0))
    candidates["52주최저"]    = candidates[ticker_col].map(lambda t: range_data.get(t, {}).get("52주최저", 0))
    candidates["연간변동폭(%)"] = candidates[ticker_col].map(lambda t: range_data.get(t, {}).get("연간변동폭(%)", 0))

    # ── 5. 종합 저평가 점수 ──
    candidates["종합_저평가점수"] = (
        candidates["per_score"] * SCORE_WEIGHTS["per_score"] +
        candidates["pbr_score"] * SCORE_WEIGHTS["pbr_score"] +
        candidates["roe_score"] * SCORE_WEIGHTS["roe_score"] +
        candidates["모멘텀_점수"] * SCORE_WEIGHTS["momentum_score"] +
        candidates["수급_점수"] * SCORE_WEIGHTS["supply_score"]
    ).round(1)

    # ── 6. 투자등급 ──
    def get_rating(score):
        if score >= 75: return "★★★ 강력매수"
        if score >= 60: return "★★  매수"
        if score >= 45: return "★   중립"
        return "     관망"

    candidates["등급"] = candidates["종합_저평가점수"].apply(get_rating)

    # ── 7. 현재가 ──
    if "종가" in candidates.columns:
        candidates["현재가"] = candidates["종가"]
    elif "현재가" not in candidates.columns:
        candidates["현재가"] = 0

    result = candidates.nlargest(TOP_N, "종합_저평가점수")[[
        "종목명", ticker_col, "현재가", "52주최고", "52주최저", "연간변동폭(%)",
        "목표가_괴리율", "이익_상향폭",
        "수급_점수", "모멘텀_점수", "종합_저평가점수", "등급",
        "PER", "PBR", "ROE", "시가총액_억"
    ]].reset_index(drop=True)

    # ──────────────────────────────────────────
    # 출력
    # ──────────────────────────────────────────
    console.print()
    console.rule(f"[bold cyan] KOSPI+KOSDAQ 저평가 종목 TOP {TOP_N} [/bold cyan]")
    console.print()

    table = Table(
        box=box.SIMPLE_HEAD,
        show_header=True,
        header_style="bold cyan",
        border_style="dim",
        row_styles=["", "dim"],
    )

    table.add_column("#", style="dim", width=3, justify="right")
    table.add_column("종목명", width=14)
    table.add_column("티커", width=8, justify="center")
    table.add_column("현재가(원)", justify="right", width=11)
    table.add_column("목표가\n괴리율", justify="right", width=9)
    table.add_column("이익\n상향폭", justify="right", width=8)
    table.add_column("수급\n점수", justify="right", width=7)
    table.add_column("모멘텀\n점수", justify="right", width=8)
    table.add_column("종합\n점수", justify="right", width=7)
    table.add_column("등급", width=14)
    table.add_column("PER", justify="right", width=6)
    table.add_column("PBR", justify="right", width=6)
    table.add_column("ROE(%)", justify="right", width=7)

    def color_pct(val, good_positive=True):
        """수치 색상 처리"""
        try:
            v = float(val)
            if good_positive:
                color = "green" if v > 10 else "yellow" if v > 0 else "red"
            else:
                color = "green" if v > 0 else "red"
            sign = "+" if v > 0 else ""
            return Text(f"{sign}{v:.1f}%", style=color)
        except Exception:
            return Text(str(val))

    def color_score(val):
        v = float(val)
        color = "bright_green" if v >= 75 else "yellow" if v >= 60 else "white" if v >= 45 else "dim"
        return Text(f"{v:.0f}", style=color)

    def color_rating(val):
        s = str(val)
        if "강력" in s: return Text(s, style="bold bright_green")
        if "매수" in s: return Text(s, style="green")
        if "중립" in s: return Text(s, style="yellow")
        return Text(s, style="dim")

    for i, row in result.iterrows():
        table.add_row(
            str(i + 1),
            str(row["종목명"])[:12],
            str(row[ticker_col]),
            f"{int(row['현재가']):,}" if row["현재가"] > 0 else "-",
            color_pct(row["목표가_괴리율"]),
            color_pct(row["이익_상향폭"]),
            color_score(row["수급_점수"]),
            color_score(row["모멘텀_점수"]),
            color_score(row["종합_저평가점수"]),
            color_rating(row["등급"]),
            f"{row['PER']:.1f}",
            f"{row['PBR']:.2f}",
            f"{row['ROE']:.1f}",
        )

    console.print(table)

    # ── 범례 ──
    console.print()
    console.print("[dim]■ 점수 기준:[/dim]  "
                  "[bright_green]75+ 강력매수[/bright_green]  "
                  "[green]60+ 매수[/green]  "
                  "[yellow]45+ 중립[/yellow]  "
                  "[dim]45미만 관망[/dim]")
    console.print("[dim]■ 목표가 괴리율: 펀더멘털 기반 적정가 대비 현재가 할인율 추정치[/dim]")
    console.print("[dim]■ 이익 상향폭: ROE/PEG 기반 EPS 성장 여력 추정치[/dim]")
    console.print("[dim]■ 수급/모멘텀 점수: 최근 10~60일 외국인·기관 순매수 및 가격·거래량 기반[/dim]")
    console.print()
    console.print("[bold yellow]⚠  본 스크리너는 참고용이며 투자 조언이 아닙니다. 투자 손익은 투자자 본인에게 귀속됩니다.[/bold yellow]")
    console.print()

    # ── Excel 저장 (종합점수 내림차순) ──
    result = result.sort_values("종합_저평가점수", ascending=False).reset_index(drop=True)
    result.index += 1
    filename = f"1_가치주_찾기_{datetime.today().strftime('%Y_%m%d_%H%M')}.xlsx"
    result.to_excel(filename, index=True, index_label="순번")

    # ── 종목명 → 네이버 6개월 일봉 차트 하이퍼링크 ──
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font

        wb = load_workbook(filename)
        ws = wb.active

        def _cw(val):
            return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

        # 종목명 → 네이버 링크 + 티커 컬럼 삭제
        name_col_idx = ticker_link_col_idx = None
        for cell in ws[1]:
            if cell.value == "종목명":
                name_col_idx = cell.column
            if cell.value == ticker_col:
                ticker_link_col_idx = cell.column

        if name_col_idx and ticker_link_col_idx:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                name_cell   = row[name_col_idx - 1]
                ticker_cell = row[ticker_link_col_idx - 1]
                if ticker_cell.value:
                    code = str(ticker_cell.value).zfill(6)
                    url = f"https://finance.naver.com/item/main.naver?code={code}"
                    name_cell.hyperlink = url
                    name_cell.font = Font(color="0563C1", underline="single", bold=True)
            ws.delete_cols(ticker_link_col_idx)

        # 전체 가운데 정렬 + 숫자 포맷
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")
                if isinstance(cell.value, float):
                    cell.value = round(cell.value, 1)
                    cell.number_format = "#,##0.0"
                elif isinstance(cell.value, int):
                    cell.number_format = "#,##0"

        # 열 너비 재조정 (티커 삭제 후)
        for col in ws.columns:
            max_w = max((_cw(cell.value) for cell in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_w + 2, 10), 50)

        wb.save(filename)
    except Exception as e:
        console.print(f"[yellow]Excel 서식 적용 실패: {e}[/yellow]")

    console.print(f"[cyan]✓ 결과 저장 완료: {filename}[/cyan]")
    console.print(f"[dim]  → 종목명 클릭 시 네이버 금융 6개월 일봉 차트로 이동[/dim]")
    console.print()


# ──────────────────────────────────────────────
if __name__ == "__main__":
    run_screener()