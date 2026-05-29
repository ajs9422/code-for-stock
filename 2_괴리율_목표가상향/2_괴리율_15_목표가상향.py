"""
================================================
  2. 괴리율 15% 이상 목표가 상향 종목 탐색
================================================

[탐색 원리]
  증권사 애널리스트가 최근 목표주가를 "상향"한 종목 중,
  현재 주가와 목표주가의 차이(괴리율)가 15% 이상인 종목을 발굴한다.
  목표주가 상향 = 기관의 확신이 담긴 신호로 해석.

[탐색 절차]
  ① 네이버 금융 리서치 리포트 목록 수집 (최근 30일)
  ② 제목에 "상향" / "↑" / "Up" 키워드가 포함된 리포트만 선별
  ③ FnGuide 컨센서스 페이지와 교차 검증 (교집합 필터)
  ④ 종목별 현재가 vs 목표주가 비교 → 괴리율 계산

[필터 조건]
  ① 리포트 발행일  최근 30일 이내
  ② 리포트 제목    상향/↑/Up 키워드 포함
  ③ 괴리율         15% 이상  (현재가 대비 목표주가 상승 여력)

[데이터 출처]
  네이버 금융 리서치 (finance.naver.com/research)
  FnGuide 컨센서스 (comp.fnguide.com)
================================================
"""

import re
from collections import Counter
from datetime import datetime, timedelta

import pandas as pd
import requests
from bs4 import BeautifulSoup
from rich.console import Console
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

HEADERS = {'User-Agent': 'Mozilla/5.0'}
console = Console()


def _parse_int(value):
    if not value:
        return 0
    cleaned = re.sub(r'[^\d]', '', value)
    try:
        return int(cleaned)
    except ValueError:
        return 0


def _parse_date(value):
    if not value:
        return None
    for fmt in ('%Y.%m.%d', '%Y-%m-%d', '%y.%m.%d'):
        try:
            return datetime.strptime(value.strip(), fmt).date()
        except ValueError:
            continue
    return None


def get_naver_report_list():
    url = 'https://finance.naver.com/research/company_list.naver'
    res = requests.get(url, headers=HEADERS, timeout=15)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, 'html.parser')
    table = soup.select_one('table.type_1')
    if table is None:
        return []

    report_list = []
    for row in table.find_all('tr'):
        cols = row.find_all('td')
        if len(cols) < 5:
            continue

        stock_link = cols[0].find('a')
        stock_name = stock_link.text.strip() if stock_link else cols[0].text.strip()
        stock_code = ''
        if stock_link and stock_link.get('href'):
            stock_code = stock_link['href'].split('code=')[-1]

        report_title = cols[1].text.strip()
        current_price = _parse_int(cols[3].text if len(cols) > 3 else '')
        report_date_text = cols[5].text.strip() if len(cols) > 5 else cols[-1].text.strip()
        report_date = _parse_date(report_date_text)

        report_list.append({
            '종목명': stock_name,
            '종목코드': stock_code,
            '제목': report_title,
            '현재가': current_price,
            '리포트일': report_date,
        })

    return report_list


def get_target_price(code):
    url = f'https://finance.naver.com/item/main.naver?code={code}'
    res = requests.get(url, headers=HEADERS, timeout=15)
    res.raise_for_status()

    soup = BeautifulSoup(res.text, 'html.parser')
    target = soup.select_one('em._target_price')
    if target and target.text:
        return _parse_int(target.text)

    for candidate in soup.select('span, em, td'):
        text = candidate.text.strip()
        if '목표' in text and any(ch.isdigit() for ch in text):
            found = re.search(r'([0-9,]+)', text)
            if found:
                return _parse_int(found.group(1))

    return 0


def get_fnguide_estimate_upgrade_codes(days=7, min_reports=2):
    url = 'https://comp.fnguide.com/SVO2/asp/SVD_Consensus.asp'
    try:
        res = requests.get(url, headers=HEADERS, timeout=15)
        res.raise_for_status()
        page = res.text

        codes = set()
        for line in page.splitlines():
            if '상향' in line and 'gicode=' in line:
                codes.update(re.findall(r'gicode=([0-9A-Z]+)', line))

        return codes
    except Exception:
        return set()


def get_upside_targets():
    print('🚀 기관의 확신이 담긴 종목을 분석 중입니다...')

    reports = get_naver_report_list()
    if not reports:
        print('⚠️ 네이버 리서치 데이터가 없습니다.')
        return pd.DataFrame()

    recent_threshold = datetime.now().date() - timedelta(days=30)
    upward_reports = [
        r for r in reports
        if r['종목코드']
        and r['리포트일']
        and r['리포트일'] >= recent_threshold
        and ('상향' in r['제목'] or '↑' in r['제목'] or 'Up' in r['제목'] or 'up' in r['제목'])
    ]

    counts = Counter((r['종목코드'], r['종목명']) for r in upward_reports)
    candidates = {
        code: name for (code, name), count in counts.items()
        if count >= 1
    }

    fnguide_codes = get_fnguide_estimate_upgrade_codes()
    if fnguide_codes:
        candidates = {code: name for code, name in candidates.items() if code in fnguide_codes}
        print(f'✅ FnGuide 교집합 {len(candidates)}개 종목 필터링 완료')
    else:
        print('⚠️ FnGuide 업그레이드 교집합을 가져오지 못했습니다. 네이버 리포트 기반 필터를 사용합니다.')

    treasure_list = []
    items = list(candidates.items())
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
        task = progress.add_task("목표가 조회 중...", total=len(items))
        for stock_code, stock_name in items:
            progress.update(task, description=f"{stock_name[:12]:<14} 조회 중")
            current_price = next((r['현재가'] for r in upward_reports if r['종목코드'] == stock_code), 0)
            target_price = get_target_price(stock_code)
            progress.advance(task)
            if target_price <= 0 or current_price <= 0:
                continue

            upside = ((target_price - current_price) / current_price) * 100
            if upside < 15:
                continue

            treasure_list.append({
            '종목명': stock_name,
            '종목코드': stock_code,
            '현재가': current_price,
            '목표가': target_price,
            '괴리율(%)': f'{upside:.2f}%',
            '리포트수(1주일)': counts[(stock_code, stock_name)],
            'FnGuide 교집합': stock_code in fnguide_codes,
        })

    return pd.DataFrame(treasure_list)


def _save_excel(df: pd.DataFrame, filename: str) -> None:
    df.to_excel(filename, index=False)
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font

        wb = load_workbook(filename)
        ws = wb.active

        def _cw(val):
            return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

        name_col_idx = ticker_col_idx = None
        for cell in ws[1]:
            if cell.value == "종목명":   name_col_idx   = cell.column
            if cell.value == "종목코드": ticker_col_idx = cell.column
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
        print(f'Excel 서식 적용 실패: {e}')


if __name__ == '__main__':
    try:
        final_targets = get_upside_targets()
        filename = f'2_괴리율_15_목표가상향_{datetime.now().strftime("%Y_%m%d_%H%M")}.xlsx'
        if not final_targets.empty:
            df_out = final_targets.sort_values(by='괴리율(%)', ascending=False).reset_index(drop=True)
            print(df_out.to_string())
        else:
            df_out = final_targets
            print('조건에 맞는 종목이 현재 없습니다. 시장 상황을 지켜보세요.')
        _save_excel(df_out, filename)
        print(f'Excel 저장 완료: {filename}')
    except Exception as e:
        print(f'Error: {e}')
