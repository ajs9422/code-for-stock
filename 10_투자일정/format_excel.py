"""
투자일정 엑셀 서식 적용

Sheet1 투자일정    : D-Day 색상, 종목 하이퍼링크, 영향도/방향 강조
Sheet2 종목인사이트: 주목도 색상, 종목명 하이퍼링크
Sheet3 요약        : 달력 형태 월별 일정표 (소형 분기보고서 제외)
"""
import copy
import glob
import os
import sys
from calendar import monthcalendar
from collections import defaultdict
from datetime import date as _date
from urllib.parse import quote_plus

import requests
import openpyxl
from concurrent.futures import ThreadPoolExecutor, as_completed
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ── 시트명 ────────────────────────────────────────────────────────────────────
SHEET_SCHEDULE = "투자일정"
SHEET_INSIGHT  = "종목인사이트"
SHEET_SUMMARY  = "요약"
SHEET_THEMATIC = "테마별일정"

# ── 공통 스타일 ───────────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN         = Side(style="thin", color="BBBBBB")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_HIGH     = PatternFill("solid", fgColor="C6EFCE")   # 영향도 7+ — 연초록
FILL_MID      = PatternFill("solid", fgColor="FFEB9C")   # 영향도 4~6 — 연노랑
FILL_POSITIVE = PatternFill("solid", fgColor="E2EFDA")   # 영향방향 긍정 — 연초록
FILL_NEGATIVE = PatternFill("solid", fgColor="FCE4D6")   # 영향방향 부정 — 연주황
FILL_URGENT   = PatternFill("solid", fgColor="FF0000")   # D-Day / D-7 이하 — 빨강
FILL_SOON     = PatternFill("solid", fgColor="FF9900")   # D-30 이하 — 주황
FILL_NEAR     = PatternFill("solid", fgColor="FFEB9C")   # D-60 이하 — 노랑

# ── 카테고리 색상 ─────────────────────────────────────────────────────────────
CAT_COLORS = {
    "기업실적":    ("FFF2CC", "7D6608"),
    "주주총회":    ("DDEBF7", "1F4E79"),
    "경제지표":    ("FCE4D6", "843C0C"),
    "글로벌이벤트": ("E2EFDA", "375623"),
    "공시":        ("EDE7F6", "4A148C"),
    "CEO·VIP방한": ("E8D5F5", "7030A0"),  # 연보라 배경 + 진보라 글자 — CEO/정상 방한
}

# ── 컬럼 너비 ─────────────────────────────────────────────────────────────────
WIDTHS_SCHEDULE = {
    "날짜": 12, "D-Day": 8, "이벤트명": 45, "카테고리": 14,
    "섹터": 18, "직접수혜종목": 22, "간접수혜종목": 28,
    "투자포인트": 55, "영향도": 8, "영향방향": 10,
    "지속기간": 10, "투자전략": 45, "출처": 10, "원본링크": 8,
}
WIDTHS_INSIGHT = {
    "종목명": 18, "관련이벤트수": 12, "직접수혜횟수": 12, "간접수혜횟수": 12,
    "주요관련이벤트": 50, "최근이벤트날짜": 14, "종합투자의견": 60, "주목도": 8,
}
WIDTHS_THEMATIC = {
    "테마": 20, "주요 일정 및 이벤트": 52, "대응 전략 Tip": 52,
}

# 테마별 (배경색, 글자색)
_THEME_STYLE: dict[str, tuple[str, str]] = {
    "반도체/AI":      ("1F4E79", "D6E4F0"),
    "제약/바이오":    ("375623", "E2EFDA"),
    "정치/외교/정책": ("7B2D00", "FCE4D6"),
    "금융/경제지표":  ("1A3A5C", "DDEBF7"),
    "빅테크/플랫폼":  ("4A148C", "EDE7F6"),
    "에너지/원자재":  ("7D6608", "FFF2CC"),
    "자동차/모빌리티":("404040", "F2F2F2"),
    "방산/우주":      ("1A3A1A", "E2EFE2"),
    "노동/파업":      ("C00000", "FFE0E0"),
}

WRAP_COLS_SCHEDULE = {"이벤트명", "투자포인트", "투자전략", "직접수혜종목", "간접수혜종목"}
WRAP_COLS_INSIGHT  = {"주요관련이벤트", "종합투자의견"}

# 투자일정 시트에서 행 단위로 분리할 종목 열
STOCK_EXPAND_COLS = {"직접수혜종목", "간접수혜종목"}


# ── 종목코드 캐시 ─────────────────────────────────────────────────────────────
_CODE_CACHE: dict[str, str] = {}   # {종목명: 종목코드 or ""}

def _lookup_stock_code(name: str) -> str:
    """Naver Finance 자동완성 API로 종목코드 조회. 없으면 빈 문자열 반환."""
    name = name.strip()
    if name in _CODE_CACHE:
        return _CODE_CACHE[name]
    try:
        url = (
            "https://ac.finance.naver.com/ac"
            f"?q={quote_plus(name)}&q_enc=UTF-8&st=111"
            "&r_format=json&r_enc=UTF-8&r_lt=111&r_unicode=0&r_escape=1"
        )
        resp = requests.get(url, timeout=5)
        data = resp.json()
        # 응답 구조: {"items": [[["종목명","코드","0","시장","KR",...], ...], ...]}
        for group in data.get("items", []):
            for item in group:
                if isinstance(item, list) and len(item) >= 2:
                    item_name = item[0]
                    item_code = item[1]
                    if item_name == name and item_code.isdigit():
                        _CODE_CACHE[name] = item_code
                        return item_code
    except Exception:
        pass
    _CODE_CACHE[name] = ""
    return ""


_STOCK_COLS = {"직접수혜종목", "간접수혜종목", "종목명"}

def _prefetch_stock_codes(worksheets: list):
    """여러 시트에서 종목명을 수집해 종목코드를 병렬 조회한 뒤 캐시에 채운다.
    스타일 적용 전에 한 번만 호출하면 이후 _stock_hyperlink 호출이 모두 캐시 히트된다."""
    names = set()
    for ws in worksheets:
        headers = [cell.value for cell in ws[1]]
        for col_i, h in enumerate(headers):
            if h not in _STOCK_COLS:
                continue
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                    min_col=col_i + 1, max_col=col_i + 1,
                                    values_only=True):
                val = str(row[0] or "").strip()
                if val and val not in ("nan", "None"):
                    for part in val.split(","):
                        part = part.strip()
                        if part and part not in _CODE_CACHE:
                            names.add(part)
    if not names:
        return
    with ThreadPoolExecutor(max_workers=12) as ex:
        futs = {ex.submit(_lookup_stock_code, n): n for n in names}
        for fut in as_completed(futs):
            fut.result()


# ── 공통 헬퍼 ─────────────────────────────────────────────────────────────────
def _apply_header(ws, col_idx: dict, widths: dict):
    ws.row_dimensions[1].height = 22
    for h, idx in col_idx.items():
        cell = ws.cell(row=1, column=idx)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(h, 14)


def _stock_hyperlink(cell, stock_name: str):
    """종목명에 네이버 금융 종목 페이지 하이퍼링크 적용."""
    if not stock_name or stock_name.strip() in ("", "nan", "None"):
        return
    code = _lookup_stock_code(stock_name.strip())
    if code:
        cell.hyperlink = f"https://finance.naver.com/item/main.naver?code={code}"
    else:
        cell.hyperlink = (
            f"https://search.naver.com/search.naver"
            f"?query={quote_plus(stock_name.strip() + ' 주가')}"
        )
    cell.font = Font(color="0563C1", underline="single")


# ── 종목 행 확장 (종목별 개별 하이퍼링크) ─────────────────────────────────────
def _expand_stock_rows(ws):
    """직접수혜종목·간접수혜종목 열의 여러 종목을 행별로 분리."""
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    stock_col_nums = {col_idx[h] for h in STOCK_EXPAND_COLS if h in col_idx}
    if not stock_col_nums:
        return

    data_rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if any(v is not None for v in row):
            data_rows.append(list(row))

    if not data_rows:
        return

    ws.delete_rows(2, ws.max_row)

    _EMPTY = {"", "nan", "None"}
    write_row = 2

    for data in data_rows:
        stocks_by_col: dict[int, list[str]] = {}
        for col_num in stock_col_nums:
            raw = str(data[col_num - 1] or "").strip()
            if raw and raw not in _EMPTY:
                stocks = [s.strip() for s in raw.split(",")
                          if s.strip() and s.strip() not in _EMPTY]
            else:
                stocks = []
            stocks_by_col[col_num] = stocks

        n_sub = max(1, max((len(v) for v in stocks_by_col.values()), default=1))

        for sub_i in range(n_sub):
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=write_row + sub_i, column=col_num)
                if col_num in stock_col_nums:
                    stocks = stocks_by_col.get(col_num, [])
                    cell.value = stocks[sub_i] if sub_i < len(stocks) else None
                elif sub_i == 0:
                    cell.value = data[col_num - 1]

        if n_sub > 1:
            for col_num in range(1, len(headers) + 1):
                if col_num not in stock_col_nums:
                    ws.merge_cells(
                        start_row=write_row, start_column=col_num,
                        end_row=write_row + n_sub - 1, end_column=col_num,
                    )

        write_row += n_sub


# ── Sheet1 투자일정 ───────────────────────────────────────────────────────────
def _style_schedule_ws(ws):
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    _apply_header(ws, col_idx, WIDTHS_SCHEDULE)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if all(c.value is None for c in row):
            continue

        row_height = 16
        for cell in row:
            h   = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val = str(cell.value).strip() if cell.value is not None else ""
            wrap = h in WRAP_COLS_SCHEDULE

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=wrap,
                shrink_to_fit=not wrap,
                horizontal="center" if h in {"D-Day", "영향도", "영향방향", "날짜",
                                              "카테고리", "출처", "지속기간"} else "left",
            )
            cell.border = BORDER

            # D-Day 표시 + 색상
            if h == "D-Day":
                try:
                    d = int(cell.value)
                    if d == 0:
                        cell.value = "D-Day"
                        cell.fill = FILL_URGENT
                        cell.font = Font(bold=True, color="FFFFFF")
                    elif d > 0:
                        cell.value = f"D-{d}"
                        if d <= 7:
                            cell.fill = FILL_URGENT
                            cell.font = Font(bold=True, color="FFFFFF")
                        elif d <= 30:
                            cell.fill = FILL_SOON
                            cell.font = Font(bold=True)
                        elif d <= 60:
                            cell.fill = FILL_NEAR
                    else:
                        cell.value = f"D+{abs(d)}"
                except (TypeError, ValueError):
                    pass

            # 카테고리 색상
            elif h == "카테고리" and val in CAT_COLORS:
                bg, fg = CAT_COLORS[val]
                cell.fill = PatternFill("solid", fgColor=bg)
                cell.font = Font(bold=True, color=fg)

            # 영향도 색상
            elif h == "영향도":
                try:
                    v = int(cell.value)
                    if v >= 7:
                        cell.fill = FILL_HIGH
                        cell.font = Font(bold=True)
                    elif v >= 4:
                        cell.fill = FILL_MID
                except (TypeError, ValueError):
                    pass

            # 영향방향 색상
            elif h == "영향방향":
                if val == "긍정":
                    cell.fill = FILL_POSITIVE
                    cell.font = Font(bold=True, color="375623")
                elif val == "부정":
                    cell.fill = FILL_NEGATIVE
                    cell.font = Font(bold=True, color="C65911")

            # 원본링크 하이퍼링크
            elif h == "원본링크" and val.startswith("http"):
                cell.hyperlink = val
                cell.value     = "링크"
                cell.font      = Font(color="0563C1", underline="single")

            # 종목 하이퍼링크 (행 분리 후 개별 적용)
            elif h in ("직접수혜종목", "간접수혜종목") and val and val not in ("nan", "None"):
                _stock_hyperlink(cell, val)

            if wrap and val:
                col_w      = WIDTHS_SCHEDULE.get(h, 14)
                lines      = max(1, int(len(val) * 2.2 / col_w) + 2)
                row_height = max(row_height, min(lines * 16, 300))

        ws.row_dimensions[row[0].row].height = row_height

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ── Sheet2 종목인사이트 ───────────────────────────────────────────────────────
def _style_insight_ws(ws):
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    _apply_header(ws, col_idx, WIDTHS_INSIGHT)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if all(c.value is None for c in row):
            continue

        row_height = 16
        for cell in row:
            h   = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val = str(cell.value).strip() if cell.value is not None else ""
            wrap = h in WRAP_COLS_INSIGHT

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=wrap,
                shrink_to_fit=not wrap,
                horizontal="center" if h in {"관련이벤트수", "직접수혜횟수", "간접수혜횟수",
                                              "최근이벤트날짜", "주목도"} else "left",
            )
            cell.border = BORDER

            # 종목명 하이퍼링크
            if h == "종목명" and val and val not in ("nan", "None"):
                _stock_hyperlink(cell, val)

            # 주목도 색상
            elif h == "주목도":
                try:
                    v = int(cell.value)
                    if v >= 7:
                        cell.fill = FILL_HIGH
                        cell.font = Font(bold=True)
                    elif v >= 4:
                        cell.fill = FILL_MID
                except (TypeError, ValueError):
                    pass

            if wrap and val:
                col_w      = WIDTHS_INSIGHT.get(h, 14)
                lines      = max(1, int(len(val) * 2.2 / col_w) + 2)
                row_height = max(row_height, min(lines * 16, 300))

        ws.row_dimensions[row[0].row].height = row_height

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ── Sheet 테마별일정 ──────────────────────────────────────────────────────────
def _style_thematic_ws(ws) -> None:
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    _apply_header(ws, col_idx, WIDTHS_THEMATIC)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if all(c.value is None for c in row):
            continue

        # 테마 이름으로 색상 결정
        theme_col = col_idx.get("테마", 1)
        theme_val = str(row[theme_col - 1].value or "").strip()
        fg, bg = _THEME_STYLE.get(theme_val, ("000000", "FFFFFF"))

        max_lines = 1
        for cell in row:
            h   = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val = str(cell.value or "")

            cell.border = BORDER

            if h == "테마":
                cell.fill      = PatternFill("solid", fgColor=bg)
                cell.font      = Font(bold=True, color=fg, size=11)
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            elif h == "주요 일정 및 이벤트":
                cell.fill      = PatternFill("solid", fgColor="FAFAFA")
                cell.font      = Font(size=10)
                cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")
                if val:
                    max_lines = max(max_lines, val.count("\n") + 1)

            elif h == "대응 전략 Tip":
                cell.fill      = PatternFill("solid", fgColor=bg)
                cell.font      = Font(size=10, color=fg)
                cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")

        # 행 높이: 이벤트 줄 수 기준
        ws.row_dimensions[row[0].row].height = max(40, max_lines * 15 + 10)

    ws.freeze_panes = "A2"


# ── Sheet3 요약 (달력) ────────────────────────────────────────────────────────

# 달력 셀 카테고리별 배경색
_CAL_CAT_FILL = {
    "경제지표":     PatternFill("solid", fgColor="FCE4D6"),
    "기업실적":     PatternFill("solid", fgColor="E2EFDA"),
    "글로벌이벤트": PatternFill("solid", fgColor="DDEBF7"),
    "주주총회":     PatternFill("solid", fgColor="EDE7F6"),
    "사용자지정":   PatternFill("solid", fgColor="FFF2CC"),
    "CEO·VIP방한":  PatternFill("solid", fgColor="E8D5F5"),  # 연보라 — CEO/정상 방한
}
_CAL_TODAY_FILL   = PatternFill("solid", fgColor="FFE066")
_CAL_WEEKEND_FILL = PatternFill("solid", fgColor="F5F5F5")
_CAL_EMPTY_FILL   = PatternFill("solid", fgColor="F9F9F9")
_CAL_MONTH_FILL   = PatternFill("solid", fgColor="1F4E79")
_CAL_DOW_FILL     = PatternFill("solid", fgColor="2E75B6")

_DOW_KR = ["월", "화", "수", "목", "금", "토", "일"]


def _read_calendar_events(ws_schedule) -> dict:
    """투자일정 시트에서 달력용 이벤트 읽기 + 필터.

    필터 규칙:
    - DART 공시(카테고리=="공시") 전체 제외
    - DART 분기/반기/사업보고서 중 영향도 < 7 제외 (소형 종목 보고서)
    - AI보완·사용자지정·네이버금융 전부 포함
    - 영향도 >= 7 이면 출처·카테고리 무관 포함
    반환: {date_str: [(이벤트명, 카테고리, 영향도), ...]} (영향도 내림차순)
    """
    headers = [cell.value for cell in ws_schedule[1]]
    col_map = {h: i for i, h in enumerate(headers) if h}

    events: dict = defaultdict(list)
    seen: set = set()

    for row_vals in ws_schedule.iter_rows(min_row=2, values_only=True):
        if not any(v is not None for v in row_vals):
            continue

        def _g(col):
            idx = col_map.get(col)
            return row_vals[idx] if idx is not None and idx < len(row_vals) else None

        date_str = str(_g("날짜") or "")[:10]
        name     = str(_g("이벤트명") or "").strip()
        if not date_str or len(date_str) != 10 or not name or name == "None":
            continue

        # 중복 방지 (행 분리로 같은 이벤트가 여러 행에 걸쳐 있을 수 있음)
        key = (date_str, name[:15])
        if key in seen:
            continue
        seen.add(key)

        source = str(_g("출처") or "")
        cat    = str(_g("카테고리") or "")
        try:
            influence = int(_g("영향도") or 0)
        except (TypeError, ValueError):
            influence = 0

        # 필터 로직
        # DART 공시·기업실적은 개별 회사 filing으로 summary 달력에 불필요
        # (주요 기업 실적은 AI보완 소스에서 별도 수집)
        if source == "DART" and cat in ("공시", "기업실적"):
            continue

        display = name[:30] + "…" if len(name) > 30 else name
        events[date_str].append((display, cat, influence))

    # 각 날짜 내 영향도 내림차순 정렬
    for ds in events:
        events[ds].sort(key=lambda x: -x[2])
    return events


def _build_calendar_ws(ws_schedule, wb) -> dict:
    """투자일정 시트를 읽어 달력 형태 '요약' 시트를 추가.
    Returns {(row, col): date_str} — 이벤트가 있는 달력 셀 좌표."""
    events_by_date = _read_calendar_events(ws_schedule)
    if not events_by_date:
        return {}

    today = _date.today()

    # 날짜 범위: 오늘 월 ~ 5개월 후 월 (이벤트 유무와 무관하게 고정)
    first_month = _date(today.year, today.month, 1)
    _em = today.month + 4
    last_display_month = _date(today.year + (_em - 1) // 12, (_em - 1) % 12 + 1, 1)

    cal_cell_dates: dict = {}  # {(row, col): date_str} — 하이퍼링크용

    # 요약 시트 생성 (기존 시트 교체)
    if SHEET_SUMMARY in wb.sheetnames:
        del wb[SHEET_SUMMARY]
    ws = wb.create_sheet(SHEET_SUMMARY, 0)  # 맨 앞에 배치

    row_cursor = 1
    cur_month  = first_month

    while (cur_month.year, cur_month.month) <= (last_display_month.year, last_display_month.month):
        y, m = cur_month.year, cur_month.month

        # 월 헤더 (7열 병합)
        ws.merge_cells(start_row=row_cursor, start_column=1,
                        end_row=row_cursor, end_column=7)
        c = ws.cell(row=row_cursor, column=1)
        c.value     = f"{y}년 {m}월"
        c.fill      = _CAL_MONTH_FILL
        c.font      = Font(color="FFFFFF", bold=True, size=13)
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border    = BORDER
        ws.row_dimensions[row_cursor].height = 26
        row_cursor += 1

        # 요일 헤더
        for j, dow in enumerate(_DOW_KR):
            c = ws.cell(row=row_cursor, column=j + 1)
            c.value     = dow
            c.fill      = _CAL_DOW_FILL
            c.font      = Font(color="FFFFFF", bold=True, size=10)
            c.alignment = HEADER_ALIGN
            c.border    = BORDER
        ws.row_dimensions[row_cursor].height = 20
        row_cursor += 1

        # 주별 행
        # width=20 기준 한글 한 줄에 들어가는 글자 수 / 9pt 폰트 1줄 높이(pt)
        _COL_KR = 9
        _LINE_PT = 13

        for week in monthcalendar(y, m):
            # 이 주에서 가장 많은 줄이 필요한 날의 예상 줄 수로 행 높이 결정
            max_lines = 1
            for d in week:
                if d == 0:
                    continue
                _ds = _date(y, m, d).strftime("%Y-%m-%d")
                _evs = events_by_date.get(_ds, [])
                day_lines = 1  # 날짜 숫자 줄
                for ev_name, _, _ in _evs[:6]:
                    n = len(ev_name) + 2  # "· " / "★ " prefix 포함
                    day_lines += max(1, (n + _COL_KR - 1) // _COL_KR)
                max_lines = max(max_lines, day_lines)
            row_h = max(50, max_lines * _LINE_PT + 8)

            for j, day_num in enumerate(week):
                c = ws.cell(row=row_cursor, column=j + 1)
                c.border = BORDER

                if day_num == 0:
                    c.fill = _CAL_EMPTY_FILL
                    continue

                d   = _date(y, m, day_num)
                ds  = d.strftime("%Y-%m-%d")
                evs = events_by_date.get(ds, [])
                if evs:
                    cal_cell_dates[(row_cursor, j + 1)] = ds

                # 셀 내용: 날짜 숫자 + 이벤트명 목록
                lines = [str(day_num)]
                for ev_name, _, ev_inf in evs[:6]:  # 최대 6건
                    prefix = "★" if ev_inf >= 8 else "·"
                    lines.append(f"{prefix} {ev_name}")
                c.value = "\n".join(lines)

                # 배경색 결정
                is_today   = (d == today)
                is_weekend = j >= 5  # 토(idx 5), 일(idx 6)
                if is_today:
                    c.fill = _CAL_TODAY_FILL
                elif evs:
                    primary_cat = evs[0][1]
                    c.fill = _CAL_CAT_FILL.get(primary_cat, PatternFill("solid", fgColor="FFFFFF"))
                elif is_weekend:
                    c.fill = _CAL_WEEKEND_FILL

                c.font      = Font(size=9, bold=bool(evs))
                c.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")

            ws.row_dimensions[row_cursor].height = row_h
            row_cursor += 1

        # 월 구분 여백
        ws.row_dimensions[row_cursor].height = 8
        row_cursor += 1

        # 다음 달로 이동
        cur_month = _date(y + 1, 1, 1) if m == 12 else _date(y, m + 1, 1)

    # 열 너비 균등 설정 (7일)
    for j in range(7):
        ws.column_dimensions[get_column_letter(j + 1)].width = 20

    # 범례
    leg_row = row_cursor + 1
    ws.merge_cells(start_row=leg_row, start_column=1,
                    end_row=leg_row, end_column=7)
    lc = ws.cell(row=leg_row, column=1)
    lc.value = (
        "범례:  ★ 영향도 8↑  · 일반  |  "
        "경제지표(연주황) / 기업실적(연초록) / 글로벌이벤트(연파랑) / 주주총회(연보라) / 사용자지정(연노랑) / CEO·VIP방한(연보라★)"
    )
    lc.font      = Font(size=8, color="666666", italic=True)
    lc.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[leg_row].height = 14

    return cal_cell_dates


# ── 요약 달력 하이퍼링크 (행 확장 후 호출) ──────────────────────────────────────
def _add_calendar_hyperlinks(ws_schedule, ws_summary, cal_cell_dates: dict):
    """요약 달력의 이벤트 셀 → 투자일정 시트의 해당 날짜 첫 행으로 하이퍼링크 연결."""
    if not cal_cell_dates:
        return

    headers = [cell.value for cell in ws_schedule[1]]
    col_map = {h: i + 1 for i, h in enumerate(headers) if h}
    date_col = col_map.get("날짜")
    if not date_col:
        return

    date_to_row: dict = {}
    for row in ws_schedule.iter_rows(min_row=2, max_row=ws_schedule.max_row):
        cell = row[date_col - 1]
        if cell.value:
            ds = str(cell.value)[:10]
            if ds not in date_to_row:
                date_to_row[ds] = cell.row

    for (r, c), date_str in cal_cell_dates.items():
        target_row = date_to_row.get(date_str)
        if target_row is None:
            continue
        cal_cell = ws_summary.cell(row=r, column=c)
        orig_font = cal_cell.font.copy()
        cal_cell.hyperlink = f"#'{ws_schedule.title}'!A{target_row}"
        cal_cell.font = orig_font  # 하이퍼링크 스타일 덮어쓰기 방지


# ── 실적개선주 시트 연동 ──────────────────────────────────────────────────────
def _add_실적개선_sheets(wb, ws_insight) -> None:
    """종목인사이트 시트의 종목을 4_financial_scanner 최신 Excel에서 찾아
    '종합'·'종합(분기포함)' 시트를 투자일정 워크북에 추가한다."""

    # 1. 종목인사이트 시트에서 종목명 수집
    headers = [cell.value for cell in ws_insight[1]]
    try:
        name_col_idx = headers.index("종목명")
    except ValueError:
        return

    insight_stocks: set[str] = set()
    for row in ws_insight.iter_rows(min_row=2, max_row=ws_insight.max_row, values_only=True):
        name = str(row[name_col_idx] or "").strip()
        if name and name not in ("nan", "None"):
            insight_stocks.add(name)

    if not insight_stocks:
        return

    # 2. 4_financial_scanner output 폴더에서 최신 파일 탐색 (_hl 우선, 없으면 일반)
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    output_dir_4 = os.path.join(base_dir, "4_financial_scanner", "output")

    def _latest(pattern):
        files = sorted([f for f in glob.glob(pattern) if "~$" not in f])
        return files[-1] if files else None

    src_path = (
        _latest(os.path.join(output_dir_4, "실적개선주_*_hl.xlsx"))
        or _latest(os.path.join(output_dir_4, "실적개선주_*.xlsx"))
    )
    if not src_path:
        print("[WARN] 4_financial_scanner/output 에 Excel 파일이 없어 종합 시트를 추가하지 않습니다.")
        return

    src_wb = openpyxl.load_workbook(src_path, data_only=True)

    # 3. "종합" → "종합(분기포함)" 순으로 시트 복사 (종목인사이트 행 기준 필터)
    for sheet_name in ("종합", "종합(분기포함)"):
        if sheet_name not in src_wb.sheetnames:
            continue
        src_ws = src_wb[sheet_name]

        src_headers = [cell.value for cell in src_ws[1]]
        try:
            src_name_idx = src_headers.index("종목명")
        except ValueError:
            continue

        # 종목인사이트에 있는 종목 행만 추출
        matched: list = []
        for row in src_ws.iter_rows(min_row=2, max_row=src_ws.max_row):
            cell_name = str(row[src_name_idx].value or "").strip()
            if cell_name in insight_stocks:
                matched.append(row)

        if not matched:
            continue

        # 기존 동명 시트 제거 후 재생성 (워크북 맨 뒤에 추가)
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws_new = wb.create_sheet(sheet_name)

        # 헤더 행 복사 (값 + 서식)
        ws_new.row_dimensions[1].height = src_ws.row_dimensions[1].height or 30
        for src_cell in src_ws[1]:
            dst = ws_new.cell(row=1, column=src_cell.column, value=src_cell.value)
            if src_cell.has_style:
                dst.font      = copy.copy(src_cell.font)
                dst.fill      = copy.copy(src_cell.fill)
                dst.alignment = copy.copy(src_cell.alignment)
                dst.border    = copy.copy(src_cell.border)

        # 데이터 행 복사 (값 + 서식 + 하이퍼링크)
        for dst_row_i, src_row in enumerate(matched, start=2):
            for src_cell in src_row:
                dst = ws_new.cell(row=dst_row_i, column=src_cell.column, value=src_cell.value)
                if src_cell.has_style:
                    dst.font          = copy.copy(src_cell.font)
                    dst.fill          = copy.copy(src_cell.fill)
                    dst.alignment     = copy.copy(src_cell.alignment)
                    dst.border        = copy.copy(src_cell.border)
                    dst.number_format = src_cell.number_format
                if src_cell.hyperlink:
                    dst.hyperlink = copy.copy(src_cell.hyperlink)

        # 열 너비 동기화
        for col_letter, dim in src_ws.column_dimensions.items():
            ws_new.column_dimensions[col_letter].width = dim.width

        ws_new.freeze_panes = "A2"
        ws_new.auto_filter.ref = ws_new.dimensions

        print(f"  → {sheet_name} 시트: {len(matched)}개 종목 (원본: {src_path})")

    src_wb.close()


# ── 진입점 ────────────────────────────────────────────────────────────────────
def format_file(path: str):
    wb = openpyxl.load_workbook(path)

    if SHEET_SCHEDULE in wb.sheetnames:
        cal_cell_dates = _build_calendar_ws(wb[SHEET_SCHEDULE], wb)   # 달력 먼저 (expand 전 원본 읽기)
        _expand_stock_rows(wb[SHEET_SCHEDULE])

        # 종목코드 병렬 프리페치 — 이후 _stock_hyperlink 는 전부 캐시 히트
        prefetch_sheets = [wb[SHEET_SCHEDULE]]
        if SHEET_INSIGHT in wb.sheetnames:
            prefetch_sheets.append(wb[SHEET_INSIGHT])
        _prefetch_stock_codes(prefetch_sheets)

        _style_schedule_ws(wb[SHEET_SCHEDULE])
        if SHEET_SUMMARY in wb.sheetnames:
            _add_calendar_hyperlinks(wb[SHEET_SCHEDULE], wb[SHEET_SUMMARY], cal_cell_dates)

    if SHEET_THEMATIC in wb.sheetnames:
        _style_thematic_ws(wb[SHEET_THEMATIC])

    if SHEET_INSIGHT in wb.sheetnames:
        _style_insight_ws(wb[SHEET_INSIGHT])
        _add_실적개선_sheets(wb, wb[SHEET_INSIGHT])

    # 시트 순서 정렬: 요약 → 테마별일정 → 투자일정 → 종목인사이트 → 나머지
    _SHEET_ORDER = [SHEET_SUMMARY, SHEET_THEMATIC, SHEET_SCHEDULE, SHEET_INSIGHT]
    ordered = [s for s in _SHEET_ORDER if s in wb.sheetnames]
    others  = [s for s in wb.sheetnames if s not in _SHEET_ORDER]
    for i, name in enumerate(ordered + others):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    base, ext = os.path.splitext(path)
    out_path  = f"{base}_fmt{ext}"
    wb.save(out_path)
    print(f"[OK] 포맷 완료: {out_path}")
    return out_path


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        base  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        files = sorted([
            f for f in glob.glob(os.path.join(base, "투자일정_*.xlsx"))
            if "~$" not in f and "_fmt" not in f
        ])
        if not files:
            print("[ERR] output 폴더에 xlsx 파일이 없습니다.")
            sys.exit(1)
        target = files[-1]
        print(f"대상 파일: {target}")

    format_file(target)
