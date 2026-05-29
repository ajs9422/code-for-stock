"""
=============================================================
  주도 섹터 · 주도주 대시보드  —  dashboard.py
  Streamlit 기반 실시간 분석 UI
=============================================================

  실행:
    streamlit run dashboard.py

  추가 설치:
    pip install streamlit plotly

  기능:
    - 섹터 등락률 히트맵 + 바 차트
    - 주도주 스코어 테이블 + 거래대금 급증 차트
    - 개별 종목 주가 / 거래량 차트 (클릭 드릴다운)
    - 날짜 선택 / 필터 사이드바
    - 자동 새로고침 (장중 모드)
=============================================================
"""

from __future__ import annotations

import time
from datetime import datetime, timedelta
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import yaml

# ── Plotly ─────────────────────────────────────────────────────
try:
    import plotly.express as px
    import plotly.graph_objects as go
    from plotly.subplots import make_subplots
except ImportError:
    st.error("plotly 미설치 → pip install plotly")
    st.stop()

# ── pykrx ──────────────────────────────────────────────────────
try:
    from pykrx import stock as krx
    HAS_PYKRX = True
except ImportError:
    HAS_PYKRX = False

# ── main.py 파이프라인 import ───────────────────────────────────
try:
    from main import (
        Config, run_pipeline,
        fetch_sector_performance, fetch_all_stocks,
        fetch_volume_ma, fetch_52w_high,
        get_sector_ticker_map, calc_hot_score,
        filter_leaders, get_recent_business_day,
        safe_call, KOSPI_SECTOR_CODES,
    )
    HAS_MAIN = True
except ImportError:
    HAS_MAIN = False


# ═══════════════════════════════════════════════════════════════
# 페이지 설정
# ═══════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="주도 섹터 · 주도주 대시보드",
    page_icon="🔥",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ── 커스텀 CSS ─────────────────────────────────────────────────
st.markdown("""
<style>
/* 전체 배경 */
[data-testid="stAppViewContainer"] {
    background: #0d0f14;
    color: #e8e6e0;
}
[data-testid="stSidebar"] {
    background: #13161d;
    border-right: 1px solid #1e2230;
}

/* 메트릭 카드 */
[data-testid="metric-container"] {
    background: #13161d;
    border: 1px solid #1e2230;
    border-radius: 10px;
    padding: 16px 20px;
}
[data-testid="stMetricValue"] {
    font-size: 1.6rem !important;
    font-weight: 700 !important;
    letter-spacing: -0.5px;
}

/* 섹션 헤더 */
.section-header {
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: #4a5568;
    margin: 28px 0 12px;
    border-bottom: 1px solid #1e2230;
    padding-bottom: 6px;
}

/* 주도주 테이블 행 */
.leader-row {
    display: flex;
    align-items: center;
    padding: 10px 16px;
    border-radius: 8px;
    margin-bottom: 6px;
    background: #13161d;
    border: 1px solid #1e2230;
    transition: border-color 0.15s;
}
.leader-row:hover { border-color: #2d3a50; }

/* 상승/하락 색 */
.pos { color: #f53d5b; }   /* 한국 기준 빨강=상승 */
.neg { color: #3b82f6; }

/* 배지 */
.badge {
    display: inline-block;
    font-size: 0.68rem;
    font-weight: 600;
    padding: 2px 8px;
    border-radius: 4px;
    background: #1e2230;
    color: #94a3b8;
    margin-left: 6px;
    letter-spacing: 0.04em;
}
.badge-hot { background: rgba(245,61,91,0.15); color: #f53d5b; }

/* 점수 바 */
.score-bar-bg {
    background: #1e2230;
    border-radius: 3px;
    height: 4px;
    width: 100%;
    margin-top: 4px;
}
.score-bar-fill {
    background: linear-gradient(90deg, #f53d5b, #ff8c42);
    border-radius: 3px;
    height: 4px;
}
</style>
""", unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# 설정 로더
# ═══════════════════════════════════════════════════════════════

@st.cache_resource
def load_config() -> "Config | dict":
    if HAS_MAIN:
        return Config()
    # main.py 없을 때 기본값
    cfg_path = Path(__file__).parent / "config.yaml"
    if cfg_path.exists():
        with open(cfg_path) as f:
            return yaml.safe_load(f)
    return {}


# ═══════════════════════════════════════════════════════════════
# 데이터 수집 (캐시)
# ═══════════════════════════════════════════════════════════════

@st.cache_data(ttl=1800, show_spinner=False)   # 30분 캐시
def load_sector_data(target_date: str) -> pd.DataFrame:
    """업종별 등락률 로드"""
    if not HAS_PYKRX:
        return _demo_sector_data()
    cfg = load_config() if HAS_MAIN else {}
    if HAS_MAIN:
        return fetch_sector_performance(target_date, cfg)
    return _fetch_sector_simple(target_date)


@st.cache_data(ttl=1800, show_spinner=False)
def load_stock_data(target_date: str) -> pd.DataFrame:
    """전종목 주가 + 핫 스코어 로드"""
    if not HAS_PYKRX:
        return _demo_stock_data()
    cfg = load_config() if HAS_MAIN else {}
    if HAS_MAIN:
        df = fetch_all_stocks(target_date, cfg)
        if df.empty:
            return _demo_stock_data()
        top_t = df.nlargest(150, "거래대금(억)")["티커"].tolist()
        vol_ma = fetch_volume_ma(top_t, target_date, cfg)
        h52    = fetch_52w_high(top_t, target_date, cfg)
        df = df.set_index("티커", drop=False).join(vol_ma).join(h52).reset_index(drop=True)
        df["거래대금배율"] = (df["거래대금(억)"] / df.get("거래대금_MA20(억)", 1)).replace([np.inf,-np.inf],0).fillna(0).round(2)
        df["52주근접도"]   = (df["종가"] / df.get("52주최고가", df["종가"])).replace([np.inf,-np.inf],0).fillna(0).round(4)
        df = calc_hot_score(df, cfg)
        return df
    return _demo_stock_data()


@st.cache_data(ttl=1800, show_spinner=False)
def load_leader_data(target_date: str, use_rsi: bool = False) -> pd.DataFrame:
    if not HAS_PYKRX or not HAS_MAIN:
        return _demo_leader_data()
    cfg = load_config()
    df_sector  = load_sector_data(target_date)
    df_stocks  = load_stock_data(target_date)
    sector_map = get_sector_ticker_map(cfg)
    return filter_leaders(df_stocks, sector_map, df_sector, cfg,
                          target_date=target_date, use_rsi=use_rsi)


@st.cache_data(ttl=900, show_spinner=False)
def load_stock_history(ticker: str, target_date: str, days: int = 60) -> pd.DataFrame:
    """개별 종목 주가 히스토리"""
    if not HAS_PYKRX:
        return _demo_history(days)
    base  = datetime.strptime(target_date, "%Y%m%d")
    start = (base - timedelta(days=days * 2)).strftime("%Y%m%d")
    df = safe_call(krx.get_market_ohlcv_by_date, fromdate=start, todate=target_date, ticker=ticker)
    return df if (df is not None and not df.empty) else _demo_history(days)


# ── 데모 데이터 (pykrx 미설치 시) ─────────────────────────────

def _demo_sector_data() -> pd.DataFrame:
    sectors = ["전기전자","기계","의약품","화학","금융업","건설업","철강금속",
               "운수장비","서비스업","유통업","음식료품","통신업"]
    np.random.seed(42)
    return pd.DataFrame({
        "섹터코드":      [str(1028+i) for i in range(len(sectors))],
        "섹터명":        sectors,
        "등락률(%)":     np.round(np.random.randn(len(sectors)) * 2.5, 2),
        "장중변동폭(%)": np.round(np.abs(np.random.randn(len(sectors))) * 1.5 + 0.5, 2),
        "시가":          np.random.randint(800, 1200, len(sectors)),
        "종가":          np.random.randint(800, 1200, len(sectors)),
    }).sort_values("등락률(%)", ascending=False).reset_index(drop=True)


def _demo_stock_data() -> pd.DataFrame:
    np.random.seed(99)
    n = 80
    names = [f"종목{i:03d}" for i in range(n)]
    return pd.DataFrame({
        "티커":          [f"{100000+i:06d}" for i in range(n)],
        "종목명":        names,
        "시장":          np.random.choice(["KOSPI","KOSDAQ"], n),
        "종가":          np.random.randint(5000, 200000, n),
        "등락률(%)":     np.round(np.random.randn(n) * 3, 2),
        "거래대금(억)":  np.round(np.random.exponential(300, n), 1),
        "시가총액(억)":  np.round(np.random.exponential(5000, n), 0),
        "거래대금배율":  np.round(np.random.exponential(2, n), 2),
        "52주근접도":    np.round(np.random.uniform(0.6, 1.0, n), 3),
        "hot_score":     np.round(np.random.uniform(0.3, 0.9, n), 4),
    })


def _demo_leader_data() -> pd.DataFrame:
    sectors = ["전기전자","기계","의약품","화학","금융업"]
    rows = []
    np.random.seed(7)
    for s in sectors:
        for j in range(3):
            rows.append({
                "섹터명":        s,
                "섹터등락률(%)": round(np.random.uniform(1,4), 2),
                "티커":          f"{200000+len(rows):06d}",
                "종목명":        f"{s[:2]}주도주{j+1}",
                "종가":          np.random.randint(10000, 300000),
                "등락률(%)":     round(np.random.uniform(2, 8), 2),
                "거래대금(억)":  round(np.random.uniform(200, 3000), 1),
                "거래대금배율":  round(np.random.uniform(2, 8), 2),
                "52주근접도":    round(np.random.uniform(0.85, 1.0), 3),
                "hot_score":     round(np.random.uniform(0.6, 0.95), 4),
                "시가총액(억)":  round(np.random.uniform(1000, 50000), 0),
                "시장":          np.random.choice(["KOSPI","KOSDAQ"]),
            })
    return pd.DataFrame(rows)


def _demo_history(days: int = 60) -> pd.DataFrame:
    np.random.seed(123)
    dates = pd.date_range(end=datetime.today(), periods=days, freq="B")
    price = 50000 + np.cumsum(np.random.randn(days) * 800)
    vol   = np.random.randint(500000, 5000000, days)
    return pd.DataFrame({
        "날짜": dates,
        "시가": price * np.random.uniform(0.99, 1.01, days),
        "고가": price * np.random.uniform(1.00, 1.03, days),
        "저가": price * np.random.uniform(0.97, 1.00, days),
        "종가": price,
        "거래량": vol,
    }).set_index("날짜")


# ═══════════════════════════════════════════════════════════════
# 차트 컴포넌트
# ═══════════════════════════════════════════════════════════════

CHART_BG   = "#0d0f14"
CHART_GRID = "#1a1e28"
CHART_TEXT = "#94a3b8"
COLOR_POS  = "#f53d5b"   # 상승 (한국 기준 빨강)
COLOR_NEG  = "#3b82f6"   # 하락


def _base_layout(title: str = "", height: int = 380) -> dict:
    return dict(
        title=dict(text=title, font=dict(size=13, color=CHART_TEXT), x=0.01),
        height=height,
        paper_bgcolor=CHART_BG,
        plot_bgcolor=CHART_BG,
        font=dict(family="'Pretendard', 'Noto Sans KR', sans-serif", color=CHART_TEXT, size=11),
        margin=dict(l=10, r=10, t=40, b=10),
        showlegend=False,
        xaxis=dict(gridcolor=CHART_GRID, linecolor=CHART_GRID, tickfont=dict(size=10)),
        yaxis=dict(gridcolor=CHART_GRID, linecolor=CHART_GRID, tickfont=dict(size=10)),
    )


def chart_sector_bar(df: pd.DataFrame, top_n: int = 12) -> go.Figure:
    """섹터 등락률 수평 바 차트"""
    df_plot = df.sort_values("등락률(%)").tail(top_n)
    colors  = [COLOR_POS if v >= 0 else COLOR_NEG for v in df_plot["등락률(%)"]]

    fig = go.Figure(go.Bar(
        x=df_plot["등락률(%)"],
        y=df_plot["섹터명"],
        orientation="h",
        marker=dict(color=colors, opacity=0.85, line=dict(width=0)),
        text=[f"{v:+.2f}%" for v in df_plot["등락률(%)"]],
        textposition="outside",
        textfont=dict(size=10, color=CHART_TEXT),
        hovertemplate="<b>%{y}</b><br>등락률: %{x:.2f}%<extra></extra>",
    ))
    layout = _base_layout("섹터별 등락률", height=420)
    layout["xaxis"].update(zeroline=True, zerolinecolor="#2d3a50", zerolinewidth=1)
    layout["yaxis"].update(gridcolor="rgba(0,0,0,0)")
    layout["margin"] = dict(l=10, r=60, t=40, b=10)
    fig.update_layout(**layout)
    return fig


def chart_sector_heatmap(df: pd.DataFrame) -> go.Figure:
    """섹터 등락률 트리맵 (크기=변동폭, 색=등락률)"""
    df = df.copy()
    df["크기"] = df["장중변동폭(%)"].clip(lower=0.1)
    fig = px.treemap(
        df,
        path=["섹터명"],
        values="크기",
        color="등락률(%)",
        color_continuous_scale=["#1d4ed8", "#13161d", "#dc2626"],
        color_continuous_midpoint=0,
        hover_data={"등락률(%)": ":.2f", "장중변동폭(%)": ":.2f"},
    )
    fig.update_traces(
        texttemplate="<b>%{label}</b><br>%{color:.2f}%",
        textfont=dict(size=12),
        marker=dict(line=dict(width=1.5, color=CHART_BG)),
    )
    fig.update_layout(
        height=380,
        paper_bgcolor=CHART_BG,
        margin=dict(l=0, r=0, t=10, b=0),
        coloraxis_showscale=False,
    )
    return fig


def chart_volume_surge(df: pd.DataFrame, top_n: int = 15) -> go.Figure:
    """거래대금 급증 종목 바 차트"""
    col = "거래대금배율" if "거래대금배율" in df.columns else "거래대금(억)"
    df_plot = df.nlargest(top_n, col)[["종목명", col, "등락률(%)"]].copy()
    colors  = [COLOR_POS if v >= 0 else COLOR_NEG for v in df_plot["등락률(%)"]]

    fig = go.Figure(go.Bar(
        x=df_plot["종목명"],
        y=df_plot[col],
        marker=dict(color=colors, opacity=0.85, line=dict(width=0)),
        text=[f"{v:.1f}x" if col == "거래대금배율" else f"{v:.0f}억" for v in df_plot[col]],
        textposition="outside",
        textfont=dict(size=9),
        hovertemplate="<b>%{x}</b><br>" + col + ": %{y:.2f}<extra></extra>",
    ))
    label = "거래대금 급증 배율 TOP " + str(top_n) if col == "거래대금배율" else f"거래대금 TOP {top_n}"
    layout = _base_layout(label, height=340)
    layout["xaxis"].update(tickangle=-35)
    fig.update_layout(**layout)
    return fig


def chart_hot_score_scatter(df: pd.DataFrame) -> go.Figure:
    """등락률 vs 거래대금배율 산점도 (hot_score = 버블 크기)"""
    col_x = "등락률(%)"
    col_y = "거래대금배율" if "거래대금배율" in df.columns else "거래대금(억)"
    col_s = "hot_score" if "hot_score" in df.columns else "거래대금(억)"

    df_plot = df.dropna(subset=[col_x, col_y]).copy()
    df_plot["_size"] = (df_plot[col_s].rank(pct=True) * 30 + 4).clip(4, 34)
    colors = [COLOR_POS if v >= 0 else COLOR_NEG for v in df_plot[col_x]]

    fig = go.Figure(go.Scatter(
        x=df_plot[col_x],
        y=df_plot[col_y],
        mode="markers",
        marker=dict(
            size=df_plot["_size"],
            color=colors,
            opacity=0.7,
            line=dict(width=0),
        ),
        text=df_plot["종목명"],
        hovertemplate="<b>%{text}</b><br>등락률: %{x:.2f}%<br>" + col_y + ": %{y:.2f}<extra></extra>",
    ))
    layout = _base_layout("등락률 vs 거래대금 급증", height=360)
    layout["xaxis"].update(title="등락률(%)", zeroline=True, zerolinecolor="#2d3a50")
    layout["yaxis"].update(title=col_y)
    fig.update_layout(**layout)
    return fig


def chart_candlestick(df_hist: pd.DataFrame, ticker: str, name: str) -> go.Figure:
    """캔들스틱 + 거래량"""
    fig = make_subplots(
        rows=2, cols=1,
        shared_xaxes=True,
        row_heights=[0.72, 0.28],
        vertical_spacing=0.03,
    )

    # 캔들
    fig.add_trace(go.Candlestick(
        x=df_hist.index,
        open=df_hist.get("시가", df_hist.get("Open", df_hist.iloc[:,0])),
        high=df_hist.get("고가", df_hist.get("High", df_hist.iloc[:,1])),
        low =df_hist.get("저가", df_hist.get("Low",  df_hist.iloc[:,2])),
        close=df_hist.get("종가",df_hist.get("Close",df_hist.iloc[:,3])),
        increasing=dict(line=dict(color=COLOR_POS, width=1), fillcolor=COLOR_POS),
        decreasing=dict(line=dict(color=COLOR_NEG, width=1), fillcolor=COLOR_NEG),
        name="주가",
        showlegend=False,
    ), row=1, col=1)

    # 거래량
    vol_col = "거래량" if "거래량" in df_hist.columns else "Volume"
    if vol_col in df_hist.columns:
        vol_colors = [
            COLOR_POS if c >= o else COLOR_NEG
            for c, o in zip(
                df_hist.get("종가", df_hist.get("Close", [0])),
                df_hist.get("시가", df_hist.get("Open",  [0])),
            )
        ]
        fig.add_trace(go.Bar(
            x=df_hist.index,
            y=df_hist[vol_col],
            marker=dict(color=vol_colors, opacity=0.7, line=dict(width=0)),
            name="거래량",
            showlegend=False,
        ), row=2, col=1)

    # 이동평균선 (5일, 20일)
    close = df_hist.get("종가", df_hist.get("Close"))
    if close is not None:
        for period, color in [(5, "#fbbf24"), (20, "#a78bfa")]:
            ma = close.rolling(period).mean()
            fig.add_trace(go.Scatter(
                x=df_hist.index, y=ma,
                mode="lines",
                line=dict(color=color, width=1.2, dash="dot"),
                name=f"MA{period}",
                showlegend=True,
            ), row=1, col=1)

    fig.update_layout(
        title=dict(text=f"{name} ({ticker})", font=dict(size=13, color=CHART_TEXT), x=0.01),
        height=460,
        paper_bgcolor=CHART_BG,
        plot_bgcolor=CHART_BG,
        font=dict(color=CHART_TEXT, size=10),
        margin=dict(l=10, r=10, t=40, b=10),
        xaxis_rangeslider_visible=False,
        legend=dict(
            orientation="h", yanchor="bottom", y=1.02,
            bgcolor="rgba(0,0,0,0)", font=dict(size=10),
        ),
    )
    for row in [1, 2]:
        fig.update_xaxes(gridcolor=CHART_GRID, linecolor=CHART_GRID, row=row, col=1)
        fig.update_yaxes(gridcolor=CHART_GRID, linecolor=CHART_GRID, row=row, col=1)
    return fig


# ═══════════════════════════════════════════════════════════════
# 사이드바
# ═══════════════════════════════════════════════════════════════

def render_sidebar() -> dict:
    with st.sidebar:
        st.markdown("### 🔥 주도 섹터 탐색")
        st.markdown("---")

        # 날짜 선택
        today = datetime.today()
        while today.weekday() >= 5:
            today -= timedelta(days=1)
        sel_date = st.date_input(
            "기준일",
            value=today,
            max_value=today,
            help="분석 기준일 (주말 자동 조정)",
        )
        target_date = sel_date.strftime("%Y%m%d")

        st.markdown("---")

        # 필터 설정
        top_n_sectors = st.slider("섹터 TOP N", 3, 15, 5)
        top_n_stocks  = st.slider("주도주 TOP N", 1, 10, 3)

        st.markdown("---")

        # 핫 스코어 가중치
        st.caption("핫 스코어 가중치")
        w_ret = st.slider("등락률", 0.0, 1.0, 0.35, 0.05)
        w_vol = st.slider("거래대금급증", 0.0, 1.0, 0.35, 0.05)
        w_mom = st.slider("52주모멘텀", 0.0, 1.0, 0.20, 0.05)
        total = w_ret + w_vol + w_mom
        if abs(total - 1.0) > 0.01:
            st.warning(f"가중치 합계: {total:.2f} (1.0 권장)")

        st.markdown("---")

        # RSI 필터
        st.caption("RSI 필터")
        use_rsi = st.checkbox(
            "RSI 필터 활성화",
            value=False,
            help="활성화 시 RSI가 설정 범위(기본 50~75)인 종목만 선별합니다.\n"
                 "⚠ 종목당 KRX 요청이 추가되어 속도가 느려집니다.",
        )
        if use_rsi:
            st.caption("⚠ RSI 필터 ON — 데이터 로딩 시간 증가")

        st.markdown("---")

        # 자동 새로고침
        auto_refresh = st.checkbox("자동 새로고침 (30분)", value=False)
        if auto_refresh:
            st.caption("30분마다 데이터 갱신")

        st.markdown("---")
        if not HAS_PYKRX:
            st.warning("⚠ pykrx 미설치\n데모 데이터로 표시 중")
        if not HAS_MAIN:
            st.info("ℹ main.py 없음\n독립 실행 모드")

    return {
        "target_date":   target_date,
        "top_n_sectors": top_n_sectors,
        "top_n_stocks":  top_n_stocks,
        "weights":       (w_ret, w_vol, w_mom),
        "use_rsi":       use_rsi,
        "auto_refresh":  auto_refresh,
    }


# ═══════════════════════════════════════════════════════════════
# 상단 메트릭 카드
# ═══════════════════════════════════════════════════════════════

def render_metrics(df_sector: pd.DataFrame, df_leaders: pd.DataFrame) -> None:
    c1, c2, c3, c4, c5 = st.columns(5)

    top_sector = df_sector.nlargest(1, "등락률(%)").iloc[0]
    bot_sector = df_sector.nsmallest(1, "등락률(%)").iloc[0]
    rising     = (df_sector["등락률(%)"] > 0).sum()
    falling    = (df_sector["등락률(%)"] <= 0).sum()
    avg_pct    = df_sector["등락률(%)"].mean()

    with c1:
        st.metric("🔥 HOT 섹터", top_sector["섹터명"],
                  delta=f"{top_sector['등락률(%)']:+.2f}%")
    with c2:
        st.metric("❄ COLD 섹터", bot_sector["섹터명"],
                  delta=f"{bot_sector['등락률(%)']:+.2f}%")
    with c3:
        st.metric("📈 상승 업종", f"{rising}개",
                  delta=f"하락 {falling}개")
    with c4:
        st.metric("📊 평균 등락률", f"{avg_pct:+.2f}%")
    with c5:
        st.metric("🏆 주도주 수", f"{len(df_leaders)}개",
                  delta=f"{df_leaders['섹터명'].nunique()}개 섹터")


# ═══════════════════════════════════════════════════════════════
# 주도주 카드 렌더링
# ═══════════════════════════════════════════════════════════════

def render_leader_cards(df_leaders: pd.DataFrame, top_n: int = 5) -> None:
    if df_leaders.empty:
        st.info("주도주 데이터가 없습니다.")
        return

    # 섹터별로 그룹화하여 탭 구성
    sectors = df_leaders["섹터명"].unique()[:top_n]
    tabs = st.tabs([f"🔥 {s}" for s in sectors])

    for tab, sector in zip(tabs, sectors):
        with tab:
            rows = df_leaders[df_leaders["섹터명"] == sector]
            for _, r in rows.iterrows():
                score_pct = int(r.get("hot_score", 0.5) * 100)
                surge_txt = f"거래대금 ×{r['거래대금배율']:.1f}" if r.get("거래대금배율", 0) > 0 else ""
                high52_txt = f"52W {r.get('52주근접도',0)*100:.0f}%" if r.get("52주근접도", 0) > 0 else ""

                st.markdown(f"""
                <div class="leader-row">
                  <div style="flex:1">
                    <span style="font-weight:600;font-size:0.95rem;">{r['종목명']}</span>
                    <span class="badge">{r['티커']}</span>
                    <span class="badge">{r['시장']}</span>
                    {'<span class="badge badge-hot">HOT</span>' if r.get('hot_score',0) > 0.7 else ''}
                  </div>
                  <div style="text-align:right;min-width:120px">
                    <div class="pos" style="font-size:1.1rem;font-weight:700">{r['등락률(%)']:+.2f}%</div>
                    <div style="font-size:0.75rem;color:#64748b">
                      {r['거래대금(억)']:.0f}억
                      {'&nbsp;·&nbsp;'+surge_txt if surge_txt else ''}
                      {'&nbsp;·&nbsp;'+high52_txt if high52_txt else ''}
                    </div>
                    <div class="score-bar-bg">
                      <div class="score-bar-fill" style="width:{score_pct}%"></div>
                    </div>
                  </div>
                </div>
                """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════
# 개별 종목 드릴다운
# ═══════════════════════════════════════════════════════════════

def render_stock_drilldown(df_stocks: pd.DataFrame, target_date: str) -> None:
    st.markdown('<div class="section-header">개별 종목 분석</div>', unsafe_allow_html=True)

    top_stocks = df_stocks.nlargest(100, "거래대금(억)")
    options = {
        f"{r['종목명']} ({r['티커']})": (r["티커"], r["종목명"])
        for _, r in top_stocks.iterrows()
    }

    col_sel, col_period = st.columns([3, 1])
    with col_sel:
        selected = st.selectbox("종목 선택", list(options.keys()), label_visibility="collapsed")
    with col_period:
        period = st.selectbox("기간", [20, 60, 120], index=1, label_visibility="collapsed")

    ticker, name = options[selected]

    with st.spinner("주가 데이터 로딩 중..."):
        df_hist = load_stock_history(ticker, target_date, days=period)

    if not df_hist.empty:
        st.plotly_chart(chart_candlestick(df_hist, ticker, name), use_container_width=True)

        # 기본 통계
        close_col = "종가" if "종가" in df_hist.columns else "Close"
        if close_col in df_hist.columns:
            c1, c2, c3, c4 = st.columns(4)
            c_series = df_hist[close_col]
            c1.metric("현재가",  f"{c_series.iloc[-1]:,.0f}원")
            c2.metric("기간 최고", f"{c_series.max():,.0f}원")
            c3.metric("기간 최저", f"{c_series.min():,.0f}원")
            c4.metric("기간 수익률", f"{(c_series.iloc[-1]/c_series.iloc[0]-1)*100:+.2f}%")
    else:
        st.warning("주가 데이터를 불러올 수 없습니다.")


# ═══════════════════════════════════════════════════════════════
# 메인 레이아웃
# ═══════════════════════════════════════════════════════════════

def main() -> None:
    # 사이드바
    opts = render_sidebar()
    target_date   = opts["target_date"]
    top_n_sectors = opts["top_n_sectors"]
    top_n_stocks  = opts["top_n_stocks"]
    use_rsi       = opts["use_rsi"]

    # 헤더
    st.markdown(
        f"## 🔥 주도 섹터 · 주도주 대시보드"
        f"<span style='font-size:0.8rem;color:#64748b;margin-left:12px'>"
        f"기준일 {target_date[:4]}-{target_date[4:6]}-{target_date[6:]}"
        f"{'  ·  데모 데이터' if not HAS_PYKRX else ''}"
        f"{'  ·  RSI 필터 ON' if use_rsi else ''}"
        f"</span>",
        unsafe_allow_html=True,
    )

    # 데이터 로드
    with st.spinner("데이터 수집 중..."):
        df_sector  = load_sector_data(target_date)
        df_stocks  = load_stock_data(target_date)
        df_leaders = load_leader_data(target_date, use_rsi=use_rsi)

    if df_sector.empty:
        st.error("섹터 데이터 없음 — 날짜 또는 네트워크 확인")
        return

    # ── 상단 메트릭 ──────────────────────────────────────────
    render_metrics(df_sector, df_leaders)
    st.markdown("")

    # ── Row 1: 섹터 바 차트 + 트리맵 ──────────────────────────
    col_bar, col_tree = st.columns([1, 1])
    with col_bar:
        st.plotly_chart(chart_sector_bar(df_sector, top_n=top_n_sectors),
                        use_container_width=True)
    with col_tree:
        st.plotly_chart(chart_sector_heatmap(df_sector),
                        use_container_width=True)

    # ── Row 2: 주도주 카드 + 산점도 ───────────────────────────
    st.markdown('<div class="section-header">섹터별 주도주</div>', unsafe_allow_html=True)
    col_leaders, col_scatter = st.columns([1.2, 1])
    with col_leaders:
        render_leader_cards(df_leaders, top_n=top_n_sectors)
    with col_scatter:
        st.plotly_chart(chart_hot_score_scatter(df_stocks),
                        use_container_width=True)

    # ── Row 3: 거래대금 급증 바 차트 ──────────────────────────
    st.markdown('<div class="section-header">거래대금 급증 종목</div>', unsafe_allow_html=True)
    st.plotly_chart(chart_volume_surge(df_stocks, top_n=15), use_container_width=True)

    # ── Row 4: 개별 종목 드릴다운 ─────────────────────────────
    if not df_stocks.empty:
        render_stock_drilldown(df_stocks, target_date)

    # ── 이메일 발송 / 엑셀 다운로드 ─────────────────────────
    st.markdown("---")
    col_email, col_excel, col_status = st.columns([1, 1, 2])
    with col_email:
        send_btn = st.button("📧 결과 이메일 발송", use_container_width=True)
    with col_excel:
        if HAS_MAIN and not df_leaders.empty:
            try:
                from main import build_excel_bytes
                excel_bytes = build_excel_bytes(df_sector, df_leaders)
                st.download_button(
                    label="💾 엑셀 다운로드",
                    data=excel_bytes,
                    file_name=f"주도주_{target_date}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.button("💾 엑셀 다운로드", disabled=True, use_container_width=True,
                          help=f"생성 실패: {e}")
        else:
            st.button("💾 엑셀 다운로드", disabled=True, use_container_width=True)
    if send_btn:
        if not HAS_MAIN:
            col_status.warning("main.py 없음 — 이메일 발송 불가")
        elif df_leaders.empty:
            col_status.warning("주도주 데이터 없음")
        else:
            with col_status:
                with st.spinner("발송 중..."):
                    try:
                        from main import send_email_report, Config as _Cfg
                        send_email_report(df_sector, df_leaders, target_date, load_config())
                        st.success("이메일 발송 완료")
                    except Exception as e:
                        st.error(f"발송 실패: {e}")

    # ── 원시 데이터 (접기) ────────────────────────────────────
    with st.expander("📋 원시 데이터 보기"):
        tab1, tab2, tab3 = st.tabs(["섹터 등락률", "전종목", "주도주"])
        with tab1:
            st.dataframe(df_sector, use_container_width=True, height=300)
        with tab2:
            cols = ["티커","종목명","시장","등락률(%)","거래대금(억)","거래대금배율","52주근접도","hot_score"]
            st.dataframe(
                df_stocks[[c for c in cols if c in df_stocks.columns]].nlargest(50, "hot_score"),
                use_container_width=True, height=300,
            )
        with tab3:
            st.dataframe(df_leaders, use_container_width=True, height=300)

    # ── 자동 새로고침 ─────────────────────────────────────────
    if opts["auto_refresh"]:
        time.sleep(1800)
        st.cache_data.clear()
        st.rerun()


if __name__ == "__main__":
    main()
