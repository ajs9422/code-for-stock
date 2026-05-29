"""
================================================
  6. 네이버 경제 뉴스 종목 언급 분석기
================================================

[탐색 원리]
  네이버 경제 뉴스에 자주 등장하는 종목은 시장의 관심이 집중된 종목이다.
  단순 언급 횟수뿐 아니라 기사의 긍정/부정 감성을 분석하여
  "좋은 이유로 자주 언급되는" 종목을 추려낸다.

[탐색 절차]
  ① 네이버 경제 뉴스 최신 기사 수집 (헤드라인 + 본문)
  ② KRX 전종목 이름 사전과 매칭하여 기사 내 종목명 자동 추출
  ③ 종목별 언급 횟수 집계
  ④ 기사 제목의 키워드로 긍정(상승·매수·호재) / 부정(하락·매도·악재) 감성 점수 산출
  ⑤ 언급 빈도 점수 + 감성 점수 합산 → 최종 순위 도출

[점수 계산 방식]
  - 언급 빈도 점수  : 뉴스에 언급된 횟수가 많을수록 고점수
  - 긍정 감성 점수  : 호재·매수·급등 등 긍정 키워드 포함 기사 비율
  - 부정 감성 페널티 : 악재·매도·급락 등 부정 키워드 포함 시 감점

[데이터 출처]
  네이버 경제 뉴스 (news.naver.com/section/101)
  FinanceDataReader (KRX 전종목 이름 사전)
================================================
"""

import time
import warnings
from collections import defaultdict
from datetime import datetime

import FinanceDataReader as fdr
import pandas as pd
import requests
from bs4 import BeautifulSoup
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

warnings.filterwarnings("ignore")
console = Console()

# ──────────────────────────────────────────────
# 설정
# ──────────────────────────────────────────────
CONFIG = {
    "max_pages": 10,            # 뉴스 페이지 수 (페이지당 ~20개)
    "min_mention": 1,           # 최소 언급 횟수 필터
    "top_n": 30,                # 최종 출력 종목 수
    "market": "ALL",          # KOSPI / KOSDAQ / ALL
    "delay": 0.3,               # 요청 간격 (초) - 서버 부하 방지
    "min_market_cap": 1000,     # 최소 시가총액 (억원)
}

# 긍정 키워드 (언급 시 점수 +)
POSITIVE_KEYWORDS = [
    "매수", "추천", "목표가", "상향", "호실적", "성장", "저평가",
    "기대", "강세", "반등", "급등", "신고가", "수혜", "모멘텀",
    "어닝서프라이즈", "깜짝실적", "컨센서스 상회", "주목", "유망",
]

# 부정 키워드 (언급 시 점수 -)
NEGATIVE_KEYWORDS = [
    "매도", "하향", "부진", "악재", "손실", "적자", "우려",
    "급락", "하락", "리스크", "경고", "불확실", "실망", "쇼크",
]

# 전문가/기관 신호 키워드 (가중치 1.5배)
EXPERT_KEYWORDS = [
    "애널리스트", "증권사", "리포트", "목표주가", "투자의견",
    "기관", "외국인", "펀드", "운용사", "전문가", "투자전략",
]

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                  "AppleWebKit/537.36 (KHTML, like Gecko) "
                  "Chrome/120.0.0.0 Safari/537.36",
    "Referer": "https://news.naver.com",
}


# ──────────────────────────────────────────────
# 1. KRX 종목 리스트 로드
# ──────────────────────────────────────────────
def load_krx_stocks(market: str) -> tuple[dict, pd.DataFrame]:
    """FinanceDataReader로 KRX 종목명 → 티커 매핑 + 전체 종목 DataFrame 반환"""
    console.print("  [cyan]KRX 종목 리스트 로딩 중...[/cyan]")

    df = fdr.StockListing("KRX")
    df = df[df["Code"].astype(str).str.match(r"^\d{6}$")].copy()

    if market != "ALL":
        df = df[df["Market"] == market]

    name_to_ticker = dict(zip(df["Name"], df["Code"]))
    console.print(f"  [green]✓ {len(name_to_ticker)}개 종목 로드 완료[/green]")
    return name_to_ticker, df


# ──────────────────────────────────────────────
# 2. 네이버 경제 뉴스 수집
# ──────────────────────────────────────────────
def fetch_naver_econ_news(page: int) -> list[dict]:
    """네이버 경제 뉴스 전체 (sid1=101)"""
    url = (
        f"https://news.naver.com/main/list.naver"
        f"?mode=LSD&mid=shm&sid1=101&page={page}"
    )
    try:
        res = requests.get(url, headers=HEADERS, timeout=10)
        res.encoding = "utf-8"
        soup = BeautifulSoup(res.text, "html.parser")

        articles = []
        for item in soup.select("ul.type06_headline li a, ul.type06 li a"):
            title = item.get_text(strip=True)
            href = item.get("href", "")
            if href and title and len(title) > 5:
                articles.append({"title": title, "url": href})
        return articles
    except Exception as e:
        console.print(f"  [red]경제 뉴스 수집 오류 (page {page}): {e}[/red]")
        return []


def fetch_naver_stock_news(page: int) -> list[dict]:
    """네이버 경제 > 증권 뉴스 (sid2=259)"""
    url = (
        f"https://news.naver.com/main/list.naver"
        f"?mode=LSD&mid=shm&sid1=101&sid2=259&page={page}"
    )
    try:
        res = requests.get(url, headers=HEADERS, timeout=10)
        res.encoding = "utf-8"
        soup = BeautifulSoup(res.text, "html.parser")

        articles = []
        for item in soup.select("ul.type06_headline li a, ul.type06 li a"):
            title = item.get_text(strip=True)
            href = item.get("href", "")
            if href and title and len(title) > 5:
                articles.append({"title": title, "url": href})
        return articles
    except Exception:
        return []


def fetch_article_body(url: str) -> str:
    """기사 본문 텍스트 추출"""
    try:
        res = requests.get(url, headers=HEADERS, timeout=8)
        res.encoding = "utf-8"
        soup = BeautifulSoup(res.text, "html.parser")

        for selector in ["#dic_area", "#newsct_article", "#articeBody", ".newsct_article", "#content"]:
            body = soup.select_one(selector)
            if body:
                return body.get_text(separator=" ", strip=True)

        paragraphs = soup.find_all("p")
        return " ".join(p.get_text(strip=True) for p in paragraphs)
    except Exception:
        return ""


# ──────────────────────────────────────────────
# 3. 종목명 추출 & 감성 분석
# ──────────────────────────────────────────────
def extract_stocks_from_text(text: str, name_to_ticker: dict) -> list[str]:
    """텍스트에서 KRX 종목명 추출"""
    found = []
    for name in name_to_ticker:
        if len(name) >= 2 and name in text:
            if name not in ["현대", "삼성", "LG", "SK", "한국", "대한", "국민"]:
                found.append(name)
    return found


def calc_sentiment_score(text: str) -> float:
    """텍스트 감성 점수 계산 (-100 ~ +100)"""
    pos = sum(text.count(kw) for kw in POSITIVE_KEYWORDS)
    neg = sum(text.count(kw) for kw in NEGATIVE_KEYWORDS)
    expert = sum(text.count(kw) for kw in EXPERT_KEYWORDS)

    multiplier = 1.5 if expert > 0 else 1.0
    score = (pos - neg) * multiplier * 10

    return max(-100, min(100, score))


def is_expert_article(title: str, body: str) -> bool:
    """전문가/기관 추천 기사 여부 판단"""
    combined = title + " " + body
    return any(kw in combined for kw in EXPERT_KEYWORDS)


# ──────────────────────────────────────────────
# 4. 펀더멘털 데이터 수집
# ──────────────────────────────────────────────
def fetch_fundamentals(tickers: list[str], stock_df: pd.DataFrame) -> pd.DataFrame:
    """네이버 금융 + FDR으로 PER/PBR/ROE/현재가/시총 수집"""
    naver_headers = {**HEADERS, "Referer": "https://finance.naver.com"}
    rows = []

    for ticker in tickers:
        row = stock_df[stock_df["Code"] == ticker]
        price  = int(row["Close"].iloc[0])  if not row.empty else 0
        mktcap = int(row["Marcap"].iloc[0]) if not row.empty else 0

        per, pbr, roe = 0.0, 0.0, 0.0
        try:
            res = requests.get(
                f"https://finance.naver.com/item/coinfo.naver?code={ticker}",
                headers=naver_headers, timeout=8,
            )
            res.encoding = "euc-kr"
            soup = BeautifulSoup(res.text, "html.parser")
            per_table = soup.select_one("table.per_table")
            if per_table:
                tds = per_table.find_all("td")
                # td[0]: PER(em[0]), EPS(em[1])
                # td[2]: N/A표시(em[0]), BPS(em[1])
                def _float(el):
                    try:
                        return float(el.get_text(strip=True).replace(",", ""))
                    except Exception:
                        return 0.0

                ems0 = tds[0].find_all("em") if len(tds) > 0 else []
                ems2 = tds[2].find_all("em") if len(tds) > 2 else []

                per = _float(ems0[0]) if ems0 else 0.0
                eps = _float(ems0[1]) if len(ems0) > 1 else 0.0
                # BPS: em 중 1000 이상인 숫자
                bps = 0.0
                for em in ems2:
                    v = _float(em)
                    if v >= 1000:
                        bps = v
                        break
                if bps > 0 and price > 0:
                    pbr = round(price / bps, 2)
                if bps > 0 and eps != 0:
                    roe = round(eps / bps * 100, 1)
        except Exception:
            pass

        rows.append({
            "티커":   ticker,
            "PER":    per,
            "PBR":    pbr,
            "ROE":    roe,
            "종가":   price,
            "시가총액": mktcap,
        })
        time.sleep(0.2)

    return pd.DataFrame(rows) if rows else pd.DataFrame()


# ──────────────────────────────────────────────
# 5. 종합 점수 계산
# ──────────────────────────────────────────────
def calc_total_score(mention_count: int, sentiment: float,
                     expert_count: int, per: float,
                     pbr: float, roe: float) -> float:
    """종합 점수 = 뉴스 점수 50% + 펀더멘털 점수 50%"""
    mention_score = min(100, mention_count * 15)
    sentiment_norm = (sentiment + 100) / 2
    expert_score = min(100, expert_count * 25)
    news_score = mention_score * 0.4 + sentiment_norm * 0.35 + expert_score * 0.25

    per_score = min(100, max(0, (40 - per) / 37 * 100)) if per > 0 else 50
    pbr_score = min(100, max(0, (5 - pbr) / 4.8 * 100)) if pbr > 0 else 50
    roe_score = min(100, max(0, roe / 40 * 100)) if roe > 0 else 30
    fund_score = per_score * 0.35 + pbr_score * 0.30 + roe_score * 0.35

    return round(news_score * 0.5 + fund_score * 0.5, 1)


# ──────────────────────────────────────────────
# 메인 실행
# ──────────────────────────────────────────────
def run():
    console.print()
    console.print(Panel.fit(
        "[bold cyan]네이버 경제 뉴스 종목 분석기[/bold cyan]\n"
        "[dim]전문가 추천 언급 + 저평가 필터 → 보물 종목 발굴[/dim]",
        border_style="cyan"
    ))
    console.print()

    # ── Step 1. KRX 종목 로드 ──
    console.rule("[cyan]Step 1. KRX 종목 리스트 로드[/cyan]")
    name_to_ticker, stock_df = load_krx_stocks(CONFIG["market"])
    console.print()

    # ── Step 2 & 3. 뉴스 수집 + 기사 분석 (단일 고정 진행바) ──
    console.rule("[cyan]Step 2-3. 네이버 경제 뉴스 수집 & 본문 분석[/cyan]")

    all_articles = []
    mention_count = defaultdict(int)
    sentiment_sum = defaultdict(float)
    expert_count  = defaultdict(int)

    total_collect = CONFIG["max_pages"] * 2  # 경제전체 + 증권

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

        # ── 뉴스 목록 수집 ──
        task_collect = progress.add_task(
            "뉴스 목록 수집 중...", total=total_collect
        )

        for page in range(1, CONFIG["max_pages"] + 1):
            articles = fetch_naver_econ_news(page)
            all_articles.extend(articles)
            progress.update(task_collect, advance=1,
                            description=f"경제 뉴스 수집 중... (page {page}/{CONFIG['max_pages']})")
            time.sleep(CONFIG["delay"])

        for page in range(1, CONFIG["max_pages"] + 1):
            articles = fetch_naver_stock_news(page)
            all_articles.extend(articles)
            progress.update(task_collect, advance=1,
                            description=f"증권 뉴스 수집 중... (page {page}/{CONFIG['max_pages']})")
            time.sleep(CONFIG["delay"])

        # 중복 제거
        seen_urls = set()
        unique_articles = []
        for a in all_articles:
            if a["url"] not in seen_urls:
                seen_urls.add(a["url"])
                unique_articles.append(a)

        progress.update(task_collect, description=f"[green]✓ 뉴스 수집 완료 ({len(unique_articles)}개)[/green]")
        progress.console.print(f"  [green]✓ 중복 제거 후 {len(unique_articles)}개 기사 수집[/green]")

        # ── 기사 본문 분석 ──
        task_analyze = progress.add_task(
            "기사 본문 분석 중...", total=len(unique_articles)
        )

        for idx, article in enumerate(unique_articles, 1):
            title = article["title"]
            body  = fetch_article_body(article["url"])
            full_text = title + " " + body

            stocks_found = extract_stocks_from_text(full_text, name_to_ticker)
            sentiment    = calc_sentiment_score(full_text)
            is_expert    = is_expert_article(title, body)

            for name in stocks_found:
                mention_count[name] += 1
                sentiment_sum[name] += sentiment
                if is_expert:
                    expert_count[name] += 1

            progress.update(
                task_analyze, advance=1,
                description=f"기사 분석 중... ({idx}/{len(unique_articles)})"
            )
            time.sleep(CONFIG["delay"])

        progress.update(task_analyze, description=f"[green]✓ 기사 분석 완료 ({len(mention_count)}개 종목 감지)[/green]")
        progress.console.print(f"  [green]✓ {len(mention_count)}개 종목 언급 감지[/green]")

    console.print()

    # ── Step 4. 필터링 ──
    console.rule("[cyan]Step 4. 언급 종목 필터링[/cyan]")

    filtered = {name: cnt for name, cnt in mention_count.items()
                if cnt >= CONFIG["min_mention"]}

    if not filtered:
        console.print("[yellow]언급 횟수 기준을 충족하는 종목이 없습니다. min_mention을 낮춰보세요.[/yellow]")
        return

    tickers_to_analyze = [name_to_ticker[n] for n in filtered if n in name_to_ticker]
    console.print(f"  [green]✓ 최소 {CONFIG['min_mention']}회 이상 언급 종목: {len(filtered)}개[/green]")
    console.print()

    # ── Step 5. 펀더멘털 수집 ──
    console.rule("[cyan]Step 5. 펀더멘털 지표 수집[/cyan]")
    fund_df = fetch_fundamentals(tickers_to_analyze, stock_df)
    console.print(f"  [green]✓ 펀더멘털 데이터 수집 완료[/green]")
    console.print()

    # ── Step 6. 종합 점수 계산 ──
    console.rule("[cyan]Step 6. 종합 점수 계산[/cyan]")

    rows = []

    for name, cnt in filtered.items():
        ticker = name_to_ticker.get(name)
        if not ticker:
            continue

        avg_sentiment = sentiment_sum[name] / cnt
        exp_cnt = expert_count[name]

        per, pbr, roe, price, mktcap = 0.0, 0.0, 0.0, 0, 0
        if fund_df is not None and not fund_df.empty and "티커" in fund_df.columns:
            row = fund_df[fund_df["티커"] == ticker]
            if not row.empty:
                r = row.iloc[0]
                per    = float(r.get("PER", 0) or 0)
                pbr    = float(r.get("PBR", 0) or 0)
                roe    = float(r.get("ROE", 0) or 0)
                price  = int(r.get("종가", 0) or 0)
                mktcap = int(r.get("시가총액", 0) or 0) // 100_000_000

        if mktcap > 0 and mktcap < CONFIG["min_market_cap"]:
            continue

        total_score = calc_total_score(cnt, avg_sentiment, exp_cnt, per, pbr, roe)

        rows.append({
            "종목명": name,
            "티커": ticker,
            "현재가": price,
            "언급횟수": cnt,
            "전문가언급": exp_cnt,
            "감성점수": round(avg_sentiment, 1),
            "PER": per,
            "PBR": pbr,
            "ROE(%)": roe,
            "시가총액(억)": mktcap,
            "종합점수": total_score,
        })

    if not rows:
        console.print("[yellow]조건을 만족하는 종목이 없습니다.[/yellow]")
        return

    result_df = (
        pd.DataFrame(rows)
        .sort_values("종합점수", ascending=False)
        .head(CONFIG["top_n"])
        .reset_index(drop=True)
    )

    def get_grade(score):
        if score >= 75: return "💎 보물"
        if score >= 60: return "⭐ 강력매수"
        if score >= 45: return "👍 매수"
        return "👀 관망"

    result_df["등급"] = result_df["종합점수"].apply(get_grade)

    console.print(f"  [green]✓ 최종 {len(result_df)}개 종목 도출[/green]")
    console.print()

    # ──────────────────────────────────────────
    # 출력
    # ──────────────────────────────────────────
    console.rule(f"[bold cyan] 📰 뉴스 언급 + 저평가 종목 TOP {CONFIG['top_n']} [/bold cyan]")
    console.print()

    table = Table(box=box.SIMPLE_HEAD, header_style="bold cyan",
                  border_style="dim", row_styles=["", "dim"])

    table.add_column("#",         width=3,  justify="right", style="dim")
    table.add_column("종목명",    width=14)
    table.add_column("티커",      width=8,  justify="center")
    table.add_column("현재가",    width=10, justify="right")
    table.add_column("언급\n횟수",   width=6,  justify="right")
    table.add_column("전문가\n언급",  width=7,  justify="right")
    table.add_column("감성\n점수",    width=7,  justify="right")
    table.add_column("PER",       width=6,  justify="right")
    table.add_column("PBR",       width=6,  justify="right")
    table.add_column("ROE%",      width=6,  justify="right")
    table.add_column("종합\n점수",    width=7,  justify="right")
    table.add_column("등급",      width=10)

    def sentiment_color(val):
        color = "bright_green" if val > 20 else "yellow" if val > 0 else "red"
        sign = "+" if val > 0 else ""
        return Text(f"{sign}{val:.0f}", style=color)

    def score_color(val):
        color = "bright_green" if val >= 75 else "green" if val >= 60 else "yellow" if val >= 45 else "dim"
        return Text(f"{val:.0f}", style=color)

    def grade_color(val):
        if "보물" in val:  return Text(val, style="bold bright_green")
        if "강력" in val:  return Text(val, style="green")
        if "매수" in val:  return Text(val, style="yellow")
        return Text(val, style="dim")

    for i, row in result_df.iterrows():
        table.add_row(
            str(i + 1),
            str(row["종목명"])[:12],
            str(row["티커"]),
            f"{int(row['현재가']):,}" if row["현재가"] > 0 else "-",
            str(int(row["언급횟수"])),
            Text(str(int(row["전문가언급"])),
                 style="bright_green" if row["전문가언급"] > 0 else "dim"),
            sentiment_color(row["감성점수"]),
            f"{row['PER']:.1f}" if row["PER"] > 0 else "-",
            f"{row['PBR']:.2f}" if row["PBR"] > 0 else "-",
            f"{row['ROE(%)']:.1f}" if row["ROE(%)"] > 0 else "-",
            score_color(row["종합점수"]),
            grade_color(row["등급"]),
        )

    console.print(table)
    console.print()

    # 보물 종목 별도 강조
    gems = result_df[result_df["등급"].str.contains("보물")]
    if not gems.empty:
        console.print(Panel(
            "\n".join(
                f"  💎 [bold bright_green]{r['종목명']}[/bold bright_green]  "
                f"[dim]{r['티커']}[/dim]  "
                f"언급 {int(r['언급횟수'])}회  "
                f"전문가 {int(r['전문가언급'])}회  "
                f"점수 [bright_green]{r['종합점수']:.0f}[/bright_green]"
                for _, r in gems.iterrows()
            ),
            title="[bold bright_green]💎 보물 종목[/bold bright_green]",
            border_style="bright_green"
        ))
        console.print()

    # 범례
    console.print("[dim]■ 감성점수: 기사 내 긍정/부정 키워드 분석 (-100 ~ +100)[/dim]")
    console.print("[dim]■ 전문가언급: 애널리스트·증권사·기관 관련 기사 내 언급 횟수[/dim]")
    console.print("[dim]■ 종합점수: 뉴스 점수 50% + 펀더멘털(PER·PBR·ROE) 50%[/dim]")
    console.print()
    console.print("[bold yellow]⚠  본 분석은 참고용이며 투자 조언이 아닙니다. 투자 손익은 투자자 본인에게 귀속됩니다.[/bold yellow]")
    console.print()

    # 엑셀 저장
    filename = f"6_NaverNewsStockFinder_{datetime.today().strftime('%Y_%m%d_%H%M')}.xlsx"
    result_df.to_excel(filename, index=False)
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font

        wb = load_workbook(filename)
        ws = wb.active

        def _cw(val):
            return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

        name_col_idx = ticker_col_idx = None
        for cell in ws[1]:
            if cell.value == "종목명": name_col_idx   = cell.column
            if cell.value == "티커":   ticker_col_idx = cell.column
        if name_col_idx and ticker_col_idx:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                nc = row[name_col_idx - 1]
                tc = row[ticker_col_idx - 1]
                if tc.value:
                    nc.hyperlink = f"https://finance.naver.com/item/main.naver?code={str(tc.value).zfill(6)}"
                    nc.font = Font(color="0563C1", underline="single", bold=True)
            ws.delete_cols(ticker_col_idx)

        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")
                if isinstance(cell.value, float):
                    cell.value = round(cell.value, 1)
                    cell.number_format = "#,##0.0"
                elif isinstance(cell.value, int):
                    cell.number_format = "#,##0"

        for col in ws.columns:
            max_w = max((_cw(cell.value) for cell in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_w + 2, 10), 50)

        wb.save(filename)
    except Exception as e:
        console.print(f"[yellow]Excel 서식 적용 실패: {e}[/yellow]")
    console.print(f"[cyan]✓ 결과 저장: {filename}[/cyan]")
    console.print()


if __name__ == "__main__":
    run()
