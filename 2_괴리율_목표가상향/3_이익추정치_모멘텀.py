"""
================================================
  3. 이익 추정치 모멘텀 분석기 (KRX 전종목)
================================================

[탐색 원리]
  애널리스트들이 목표주가 컨센서스를 올렸다는 것은 이익 추정치도 함께
  상향됐다는 신호다. 이익 추정치가 오르는 종목은 주가도 따라오르는
  경향이 있어 "모멘텀 투자"의 핵심 지표로 활용된다.

[탐색 절차]
  ① FinanceDataReader로 KRX 전종목 리스트 수집
  ② FnGuide에서 종목별 목표주가 컨센서스 현재값(AVG_PRC)과
     이전값(AVG_PRC_BF) 비교 → 상향률 계산
  ③ FnGuide SVD_Main에서 EPS·PER·추정기관 수 추가 수집
  ④ 조건을 충족하는 종목만 필터링 후 점수 순 정렬

[필터 조건]
  ① 시가총액   500억 이상
  ② 목표주가 컨센서스 상향률   5% 이상
  ③ 애널리스트 커버리지   1명 이상 (최소 1개 증권사 리포트 존재)

[데이터 출처]
  FinanceDataReader (전종목 리스트)
  FnGuide 컨센서스 API (comp.fnguide.com)
================================================
"""

import json
import re
import time
import warnings
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
    "min_market_cap": 500,   # 최소 시가총액 (억원)
    "min_revision":   10.0,  # 최소 상향률 (%)
    "top_n":          30,    # 최종 출력 종목 수
    "delay":          0.3,   # 요청 간격 (초)
    "market":         "ALL", # KOSPI / KOSDAQ / ALL
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Referer":         "https://comp.fnguide.com",
    "Accept-Language": "ko-KR,ko;q=0.9",
}


# ──────────────────────────────────────────────
# 1. KRX 전종목 로드
# ──────────────────────────────────────────────
def load_stocks() -> pd.DataFrame:
    """FinanceDataReader로 KRX 전종목 + 시총 필터"""
    console.print("  [cyan]KRX 전종목 로딩 중...[/cyan]")
    df = fdr.StockListing("KRX")
    df = df[df["Code"].astype(str).str.match(r"^\d{6}$")].copy()

    if CONFIG["market"] != "ALL":
        df = df[df["Market"] == CONFIG["market"]]

    df = df[df["Marcap"] >= CONFIG["min_market_cap"] * 1e8].copy()
    df = df.reset_index(drop=True)
    console.print(
        f"  [green]✓ {len(df)}개 종목 "
        f"(시총 {CONFIG['min_market_cap']}억↑, {CONFIG['market']})[/green]"
    )
    return df


# ──────────────────────────────────────────────
# 2. fnguide 데이터 수집
# ──────────────────────────────────────────────
def _parse_price(s: str) -> float:
    """'290,800' → 290800.0"""
    try:
        return float(re.sub(r"[^\d.]", "", str(s)))
    except Exception:
        return 0.0


def fetch_consensus(code: str) -> dict | None:
    """
    fnguide 03_ JSON + SVD_Main에서 컨센서스 수집

    반환:
        avg_prc     : 현재 목표주가 컨센서스
        avg_prc_bf  : 직전 목표주가 컨센서스
        revision    : 상향률 (%)
        analyst_cnt : 커버 애널리스트 수
        eps         : 현재 EPS 추정치
        per         : 현재 PER 추정치
    """
    gicode = f"A{code.zfill(6)}"

    # ── ① 03_ JSON : 목표주가 컨센서스 현재/직전 ──
    json_url = (
        f"https://comp.fnguide.com/SVO2/json/data/01_06/03_{gicode}.json"
    )
    try:
        r = requests.get(json_url, headers=HEADERS, timeout=10)
        r.encoding = "utf-8"
        data = json.loads(r.text.lstrip("\ufeff"))
        analysts = data.get("comp", [])
    except Exception:
        return None

    if not analysts:
        return None

    # 첫 번째 레코드에 집계값 포함 (AVG_PRC, AVG_PRC_BF)
    a = analysts[0]
    avg_cur = _parse_price(a.get("AVG_PRC", ""))
    avg_bef = _parse_price(a.get("AVG_PRC_BF", ""))

    if avg_cur <= 0 or avg_bef <= 0:
        return None

    revision = (avg_cur - avg_bef) / avg_bef * 100

    # ── ② SVD_Main : EPS / PER / 추정기관수 ──
    eps = per = 0.0
    analyst_cnt = len(analysts)
    try:
        main_url = (
            f"https://comp.fnguide.com/SVO2/ASP/SVD_Main.asp"
            f"?pGB=1&gicode={gicode}&cID=&MenuYn=Y&ReportGB=&NewMenuID=11&stkGb=701"
        )
        r2 = requests.get(main_url, headers=HEADERS, timeout=10)
        r2.encoding = "utf-8"
        soup = BeautifulSoup(r2.text, "html.parser")
        tables = soup.find_all("table")
        # table[7]: 투자의견 | 목표주가 | EPS | PER | 추정기관수
        if len(tables) > 7:
            rows = tables[7].find_all("tr")
            if len(rows) > 1:
                cells = [c.get_text(strip=True) for c in rows[1].find_all(["td", "th"])]
                # cells: [투자의견, 목표주가, EPS, PER, 추정기관수]
                if len(cells) >= 4:
                    eps = _parse_price(cells[2])
                    per = _parse_price(cells[3])
                if len(cells) >= 5:
                    try:
                        analyst_cnt = int(re.sub(r"[^\d]", "", cells[4]))
                    except Exception:
                        pass
    except Exception:
        pass

    return {
        "avg_prc":    round(avg_cur, 0),
        "avg_prc_bf": round(avg_bef, 0),
        "revision":   round(revision, 2),
        "analyst_cnt": analyst_cnt,
        "eps":        eps,
        "per":        per,
    }


# ──────────────────────────────────────────────
# 3. 전종목 순회
# ──────────────────────────────────────────────
def analyze_all(stocks_df: pd.DataFrame) -> list[dict]:
    results = []
    total   = len(stocks_df)

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=32, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console,
        transient=False,
    ) as progress:
        task = progress.add_task("fnguide 컨센서스 수집 중...", total=total)

        for _, row in stocks_df.iterrows():
            code   = str(row["Code"]).zfill(6)
            name   = str(row["Name"])
            market = str(row.get("Market", ""))
            close  = int(row.get("Close",  0) or 0)
            marcap = int(row.get("Marcap", 0) or 0) // 100_000_000

            cns = fetch_consensus(code)
            progress.advance(task)

            if cns is None or cns["revision"] < CONFIG["min_revision"]:
                time.sleep(CONFIG["delay"])
                continue

            results.append({
                "종목명":            name,
                "티커":              code,
                "시장":              market,
                "현재가":            close,
                "시가총액(억)":      marcap,
                "목표주가_현재":     int(cns["avg_prc"]),
                "목표주가_직전":     int(cns["avg_prc_bf"]),
                "컨센서스_상향률(%)": cns["revision"],
                "EPS추정":           cns["eps"],
                "PER추정":           cns["per"],
                "커버애널리스트":    cns["analyst_cnt"],
                "평가": (
                    "⭐⭐⭐" if cns["revision"] >= 30
                    else "⭐⭐" if cns["revision"] >= 20
                    else "⭐"
                ),
            })
            time.sleep(CONFIG["delay"])

        progress.update(
            task,
            description=f"[green]✓ 완료 — {len(results)}개 상향 종목[/green]",
        )

    return results


# ──────────────────────────────────────────────
# 4. 터미널 출력
# ──────────────────────────────────────────────
def print_results(df: pd.DataFrame) -> None:
    console.rule(
        f"[bold cyan] 목표주가 컨센서스 {CONFIG['min_revision']}%↑ 종목 "
        f"TOP {CONFIG['top_n']} [/bold cyan]"
    )
    console.print()

    table = Table(
        box=box.SIMPLE_HEAD, header_style="bold cyan",
        border_style="dim", row_styles=["", "dim"],
    )
    table.add_column("#",              width=3,  justify="right", style="dim")
    table.add_column("종목명",         width=16)
    table.add_column("티커",           width=8,  justify="center")
    table.add_column("시장",           width=7,  justify="center")
    table.add_column("현재가",         width=10, justify="right")
    table.add_column("시총(억)",       width=9,  justify="right")
    table.add_column("목표주가\n현재", width=10, justify="right")
    table.add_column("목표주가\n직전", width=10, justify="right")
    table.add_column("상향률(%)",      width=9,  justify="right")
    table.add_column("EPS추정",        width=9,  justify="right")
    table.add_column("PER추정",        width=7,  justify="right")
    table.add_column("애널\n리스트",   width=6,  justify="right")
    table.add_column("평가",           width=8)

    def rate_color(val: float) -> Text:
        color = "bright_green" if val >= 30 else "green" if val >= 20 else "yellow"
        return Text(f"+{val:.1f}%", style=color)

    for i, row in df.iterrows():
        table.add_row(
            str(i + 1),
            str(row["종목명"])[:14],
            str(row["티커"]),
            str(row["시장"]),
            f"{int(row['현재가']):,}" if row["현재가"] > 0 else "-",
            f"{int(row['시가총액(억)']):,}",
            f"{int(row['목표주가_현재']):,}",
            f"{int(row['목표주가_직전']):,}",
            rate_color(row["컨센서스_상향률(%)"]),
            f"{int(row['EPS추정']):,}" if row["EPS추정"] > 0 else "-",
            f"{row['PER추정']:.1f}"    if row["PER추정"]  > 0 else "-",
            str(int(row["커버애널리스트"])),
            str(row["평가"]),
        )

    console.print(table)
    console.print()


# ──────────────────────────────────────────────
# 5. Excel 저장
# ──────────────────────────────────────────────
def save_excel(df: pd.DataFrame, filename: str) -> None:
    df.to_excel(filename, index=False)
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font

        wb = load_workbook(filename)
        ws = wb.active

        # 열 너비 자동 조정 (한글 2배)
        def _w(val: str) -> int:
            return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

        # 종목명 → 네이버 링크 + 티커 컬럼 삭제
        name_col = ticker_col = None
        for cell in ws[1]:
            if cell.value == "종목명": name_col   = cell.column
            if cell.value == "티커":  ticker_col = cell.column

        if name_col and ticker_col:
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                t  = str(row[ticker_col - 1].value or "").zfill(6)
                nc = row[name_col - 1]
                nc.hyperlink = f"https://finance.naver.com/item/main.naver?code={t}"
                nc.font = Font(color="0563C1", underline="single", bold=True)
            ws.delete_cols(ticker_col)

        # 전체 가운데 정렬 + 숫자 포맷
        for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
            for cell in row:
                cell.alignment = Alignment(horizontal="center")
                if isinstance(cell.value, float):
                    cell.value = round(cell.value, 1)
                    cell.number_format = "#,##0.0"
                elif isinstance(cell.value, int):
                    cell.number_format = "#,##0"

        # 열 너비 재조정
        for col in ws.columns:
            max_w = max((_w(cell.value) for cell in col), default=0)
            ws.column_dimensions[col[0].column_letter].width = min(max(max_w + 2, 10), 40)

        wb.save(filename)
    except Exception as e:
        console.print(f"[yellow]Excel 서식 적용 실패: {e}[/yellow]")


# ──────────────────────────────────────────────
# 메인
# ──────────────────────────────────────────────
def main():
    console.print()
    console.print(Panel.fit(
        "[bold cyan]이익 추정치 모멘텀 분석기[/bold cyan]\n"
        "[dim]KRX 전종목 × fnguide 목표주가 컨센서스 상향률[/dim]\n"
        f"[dim]시총 {CONFIG['min_market_cap']}억↑  ·  상향률 {CONFIG['min_revision']}%↑  ·  TOP {CONFIG['top_n']}[/dim]",
        border_style="cyan",
    ))
    console.print()

    # Step 1. 종목 로드
    console.rule("[cyan]Step 1. KRX 종목 로드[/cyan]")
    stocks_df = load_stocks()
    est_min = len(stocks_df) * CONFIG["delay"] / 60
    console.print(f"  [dim]예상 소요시간: 약 {est_min:.0f}분[/dim]")
    console.print()

    # Step 2. fnguide 컨센서스 수집
    console.rule("[cyan]Step 2. fnguide 컨센서스 수집[/cyan]")
    raw = analyze_all(stocks_df)
    console.print()

    if not raw:
        console.print(f"[yellow]상향률 {CONFIG['min_revision']}% 이상 종목이 없습니다.[/yellow]")
        return

    # Step 3. 정렬 & TOP N
    result_df = (
        pd.DataFrame(raw)
        .sort_values("컨센서스_상향률(%)", ascending=False)
        .head(CONFIG["top_n"])
        .reset_index(drop=True)
    )

    # Step 4. 출력
    console.rule("[cyan]Step 3. 결과[/cyan]")
    print_results(result_df)

    console.print("[dim]■ 목표주가_현재 / 직전: fnguide 애널리스트 컨센서스 평균 (원)[/dim]")
    console.print("[dim]■ 상향률: (현재 컨센서스 - 직전 컨센서스) / 직전 × 100[/dim]")
    console.print("[dim]■ 목표주가 상향 ≈ 이익 추정치 상향의 선행 지표[/dim]")
    console.print("[bold yellow]⚠  본 결과는 참고용이며 투자 조언이 아닙니다.[/bold yellow]")
    console.print()

    # Step 5. Excel 저장
    filename = f"3_이익추정치_모멘텀_{datetime.now().strftime('%Y_%m%d_%H%M')}.xlsx"
    save_excel(result_df, filename)
    console.print(f"[cyan]✓ Excel 저장: {filename}[/cyan]")
    console.print()


if __name__ == "__main__":
    main()
