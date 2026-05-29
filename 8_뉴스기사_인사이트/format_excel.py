"""
뉴스 인사이트 엑셀 서식 적용

Sheet1 종목뉴스        : 직접·간접수혜종목 하이퍼링크, 뉴스방향성 색상, 폭등전조 강조
Sheet2 투자인사이트    : 합성 분석 4섹션 (시장방향성 / 섹터테마 / 주목종목 / 악재경고)
Sheet3 시장일반        : 헤더/테두리/줄바꿈
Sheet4 종합            : 직접수혜종목 × 4_financial_scanner 재무데이터 (분기 컬럼 제외)
Sheet5 종합(분기포함)  : 종합 + 분기별 매출·영업익 YoY 컬럼

사용법: python format_excel.py [파일경로]  (생략 시 output 최신 파일 자동 선택)
"""
import sys
import glob
import os
from urllib.parse import quote, quote_plus

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# ── 시트명 ────────────────────────────────────────────────────────────────────
SHEET_INSIGHT = "투자인사이트"
SHEET_STOCKS  = "종목뉴스"
SHEET_MARKET  = "시장일반"
SHEET_BRIEF   = "종합"
SHEET_FULL    = "종합(분기포함)"

# ── 공통 스타일 ───────────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN         = Side(style="thin", color="BBBBBB")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

FILL_HIGH      = PatternFill("solid", fgColor="C6EFCE")  # 중요도 7+ / 상승여력 — 연초록
FILL_MID       = PatternFill("solid", fgColor="FFEB9C")  # 중요도 4~6 / 영업익증가율 — 연노랑
FILL_EXPLOSIVE = PatternFill("solid", fgColor="FCE4D6")  # 폭등전조유형 비어있지 않은 행 — 연주황
FILL_VALUE     = PatternFill("solid", fgColor="D6E4FC")  # 가치주신호 — 연파랑
FILL_PER_LOW   = PatternFill("solid", fgColor="DDEBF7")  # PER 낮음 — 연파랑
FILL_POSITIVE  = PatternFill("solid", fgColor="E2EFDA")  # 뉴스방향성 긍정 — 연초록
FILL_NEGATIVE  = PatternFill("solid", fgColor="FCE4D6")  # 뉴스방향성 부정 — 연주황
FILL_CAUTION   = PatternFill("solid", fgColor="FFF2CC")  # 악재경고 섹션 헤더 — 연노랑
FILL_SECTION   = PatternFill("solid", fgColor="2E4057")  # 투자인사이트 섹션 헤더 — 진남색
FILL_SUBSEC    = PatternFill("solid", fgColor="048A81")  # 투자인사이트 서브섹션 — 청록
FILL_CAUTION_HDR = PatternFill("solid", fgColor="843C0C")  # 악재경고 섹션 헤더 — 진주황
FONT_SECTION   = Font(color="FFFFFF", bold=True, size=11)
FONT_SUBSEC    = Font(color="FFFFFF", bold=True, size=10)
FONT_CAUTION   = Font(color="FFFFFF", bold=True, size=10)

# 4_financial_scanner과 동일한 수치 하이라이트 기준
PERF_MIN_UPSIDE2   = 20.0
PERF_MAX_PER       = 12.0
PERF_MIN_OP_GROWTH = 30.0

# ── 시트별 컬럼 너비 ──────────────────────────────────────────────────────────
WIDTHS_STOCKS = {
    "뉴스제목":       40, "출처":           10, "핵심이슈":       48,
    "뉴스방향성":     10,
    "직접수혜종목":   22, "간접수혜종목":   28, "공급망연관종목": 22, "부정영향종목": 18,
    "투자호재성격":   16, "폭등전조유형":   16, "단기폭발력":     10,
    "숨은수혜주_여부": 12, "가치주신호":   10,
    "인사이트":       70, "중요도":         8,  "원본링크":        8,
}
WIDTHS_MARKET = {
    "뉴스제목":     40, "출처":       10, "핵심이슈":     48,
    "뉴스방향성":   10,
    "투자호재성격": 16, "단기폭발력": 10, "인사이트":     70, "중요도": 8, "원본링크": 8,
}
WRAP_COLS = {"뉴스제목", "핵심이슈", "인사이트", "관련뉴스", "직접수혜종목", "간접수혜종목"}

# 종목뉴스 시트에서 행 단위로 분리할 종목 열 (각 종목에 개별 하이퍼링크 부여)
STOCK_EXPAND_COLS = {"직접수혜종목", "간접수혜종목"}

# ── 투자인사이트 시트 컬럼 너비 ──────────────────────────────────────────────
WIDTHS_INSIGHT = {
    # 섹션A 시장 방향성
    "시장 방향성 요약": 80, "투자 주의사항": 60,
    # 섹션B 유망 섹터/테마
    "순위": 8, "섹터/테마": 22, "투자 근거": 55, "핵심 촉매": 45,
    "직접수혜종목": 28, "간접수혜종목": 28, "중요도": 10,
    # 섹션C 주목 종목
    "종목명": 18, "발굴유형": 14, "공급망 위치": 14,
    "호재 근거": 55, "기사 언급횟수": 14, "리스크": 40,
    # 섹션D 악재 경고 종목
    "악재 내용": 55, "영향 강도": 12,
}


def _apply_header(ws, col_idx: dict, widths: dict):
    ws.row_dimensions[1].height = 22
    for h, idx in col_idx.items():
        cell = ws.cell(row=1, column=idx)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(h, 14)


def _fval(row, col):
    if not col:
        return None
    v = row[col - 1].value
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ── Sheet1 투자인사이트 ───────────────────────────────────────────────────────
def _style_insight_ws(ws, section_rows: tuple):
    """투자인사이트 시트 서식 — 4개 섹션을 startrow 기반으로 적용.

    section_rows: (n_summary, n_sectors, n_watch, n_caution)
    섹션A(시장방향성) → 섹션B(유망섹터) → 섹션C(주목종목) → 섹션D(악재경고)
    """
    # tuple 길이가 3이면 하위 호환 (n_caution=0)
    if len(section_rows) == 3:
        n_summary, n_sectors, n_watch = section_rows
        n_caution = 0
    else:
        n_summary, n_sectors, n_watch, n_caution = section_rows

    # 각 섹션 헤더 행 번호 (1-indexed)
    row_summary_hdr = 1
    row_sectors_hdr = row_summary_hdr + n_summary + 2
    row_watch_hdr   = row_sectors_hdr + n_sectors + 2
    row_caution_hdr = row_watch_hdr + n_watch + 2

    section_header_rows = {row_summary_hdr, row_sectors_hdr, row_watch_hdr}
    if n_caution > 0:
        section_header_rows.add(row_caution_hdr)

    # 컬럼 너비: 헤더 행에서 컬럼명 읽어 적용
    applied_cols: set[str] = set()
    for hdr_row in sorted(section_header_rows):
        for cell in ws[hdr_row]:
            h = str(cell.value or '').strip()
            col_letter = get_column_letter(cell.column)
            if h and col_letter not in applied_cols:
                ws.column_dimensions[col_letter].width = WIDTHS_INSIGHT.get(h, 16)
                applied_cols.add(col_letter)

    # 행별 서식
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        row_no = row[0].row
        is_section_hdr = row_no in section_header_rows
        is_empty       = all(c.value is None for c in row)

        if is_empty:
            ws.row_dimensions[row_no].height = 8
            continue

        if is_section_hdr:
            ws.row_dimensions[row_no].height = 22
            for cell in row:
                if row_no == row_summary_hdr:
                    fill, font = FILL_SECTION, FONT_SECTION
                elif row_no == row_caution_hdr:
                    fill, font = FILL_CAUTION_HDR, FONT_CAUTION
                else:
                    fill, font = FILL_SUBSEC, FONT_SUBSEC
                cell.fill      = fill
                cell.font      = font if cell.value else Font(color="FFFFFF")
                cell.alignment = HEADER_ALIGN
                cell.border    = BORDER
            continue

        # 데이터 행
        row_height = 16
        for cell in row:
            col_no  = cell.column
            sec_hdr = max(r for r in section_header_rows if r <= row_no)
            hdr_cells = list(ws[sec_hdr])
            h = hdr_cells[col_no - 1].value if col_no <= len(hdr_cells) else ''
            h = str(h or '').strip()

            val  = str(cell.value).strip() if cell.value is not None else ''
            wrap = h in {"시장 방향성 요약", "투자 주의사항", "투자 근거", "핵심 촉매",
                         "호재 근거", "리스크", "악재 내용"}

            cell.alignment = Alignment(
                vertical="center",
                wrap_text=wrap,
                horizontal="center" if h in {"순위", "중요도", "기사 언급횟수", "영향 강도"} else "left",
            )
            cell.border = BORDER

            if h == "중요도":
                try:
                    v = int(cell.value)
                    if v >= 7:
                        cell.fill = FILL_HIGH
                        cell.font = Font(bold=True)
                    elif v >= 4:
                        cell.fill = FILL_MID
                except (TypeError, ValueError):
                    pass

            if h == "영향 강도":
                try:
                    v = int(cell.value)
                    if v >= 7:
                        cell.fill = FILL_NEGATIVE
                        cell.font = Font(bold=True, color="C65911")
                except (TypeError, ValueError):
                    pass

            if wrap and val:
                col_w      = WIDTHS_INSIGHT.get(h, 16)
                lines      = max(1, int(len(val) * 2.2 / col_w) + 2)
                row_height = max(row_height, min(lines * 16, 300))

        ws.row_dimensions[row_no].height = row_height

    ws.freeze_panes = "A2"


# ── 종목뉴스 행 확장 (종목별 개별 하이퍼링크) ────────────────────────────────
def _expand_stock_rows(ws):
    """직접수혜종목·간접수혜종목 열의 여러 종목을 행별로 분리한다.

    콤마로 구분된 종목이 여러 개면 종목 수만큼 행을 확장하고,
    나머지 열은 수직 병합(merge)하여 각 종목에 개별 하이퍼링크를 붙일 수 있게 한다.
    """
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    stock_col_nums = {col_idx[h] for h in STOCK_EXPAND_COLS if h in col_idx}
    if not stock_col_nums:
        return

    # 데이터 읽기 (값만, 헤더 제외)
    data_rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=True):
        if any(v is not None for v in row):
            data_rows.append(list(row))

    if not data_rows:
        return

    # 기존 데이터 행 전체 삭제
    ws.delete_rows(2, ws.max_row)

    _EMPTY = {"", "nan", "None"}

    write_row = 2
    for data in data_rows:
        # 종목 열별 리스트 파싱
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

        # 서브 행 작성
        for sub_i in range(n_sub):
            for col_num in range(1, len(headers) + 1):
                cell = ws.cell(row=write_row + sub_i, column=col_num)
                if col_num in stock_col_nums:
                    stocks = stocks_by_col.get(col_num, [])
                    cell.value = stocks[sub_i] if sub_i < len(stocks) else None
                elif sub_i == 0:
                    cell.value = data[col_num - 1]
                # sub_i > 0의 비종목 열은 None 유지 → 병합으로 마스터 셀 표시

        # 비종목 열 수직 병합
        if n_sub > 1:
            for col_num in range(1, len(headers) + 1):
                if col_num not in stock_col_nums:
                    ws.merge_cells(
                        start_row=write_row, start_column=col_num,
                        end_row=write_row + n_sub - 1, end_column=col_num,
                    )

        write_row += n_sub


# ── Sheet2/3 공통 스타일 ──────────────────────────────────────────────────────
def _style_sheet(ws, widths: dict):
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}
    _apply_header(ws, col_idx, widths)

    col_score = col_idx.get("중요도")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        score_val = None
        if col_score:
            try:
                score_val = int(ws.cell(row=row[0].row, column=col_score).value)
            except (TypeError, ValueError):
                pass

        row_height = 16
        for cell in row:
            h    = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val  = str(cell.value).strip() if cell.value is not None else ""
            wrap = h in WRAP_COLS

            cell.alignment = Alignment(
                vertical="center", wrap_text=wrap, shrink_to_fit=not wrap,
                horizontal="center" if h in {"중요도", "단기폭발력", "뉴스방향성"} else "left",
            )
            cell.border = BORDER

            # 원본링크 → "링크" 하이퍼링크
            if h == "원본링크" and val.startswith("http"):
                cell.hyperlink = val
                cell.value     = "링크"
                cell.font      = Font(color="0563C1", underline="single")

            # 직접수혜종목 · 간접수혜종목 → 네이버 주가 검색 링크 (종목명+"주가")
            elif h in ("직접수혜종목", "간접수혜종목", "관련종목") and val and val not in ("", "nan"):
                first = val.split(",")[0].strip()
                if first:
                    cell.hyperlink = (
                        f"https://search.naver.com/search.naver"
                        f"?query={quote_plus(first + ' 주가')}"
                    )
                    cell.font = Font(color="0563C1", underline="single")

            # 중요도 · 단기폭발력 색상
            if h in ("중요도", "단기폭발력") and score_val is not None:
                try:
                    v = int(cell.value)
                    if v >= 7:
                        cell.fill = FILL_HIGH
                        cell.font = Font(bold=True)
                    elif v >= 4:
                        cell.fill = FILL_MID
                except (TypeError, ValueError):
                    pass

            # 폭등전조유형 — 비어있지 않으면 연주황 강조
            if h == "폭등전조유형" and val and val not in ("", "nan"):
                cell.fill = FILL_EXPLOSIVE
                cell.font = Font(bold=True, color="C65911")

            # 뉴스방향성 색상
            if h == "뉴스방향성":
                if val == "긍정":
                    cell.fill = FILL_POSITIVE
                    cell.font = Font(bold=True, color="375623")
                elif val == "부정":
                    cell.fill = FILL_NEGATIVE
                    cell.font = Font(bold=True, color="C65911")

            # 가치주신호 True — 연파랑
            if h == "가치주신호" and str(val).lower() == "true":
                cell.fill = FILL_VALUE
                cell.font = Font(bold=True, color="2E75B6")

            if wrap and val:
                col_w      = widths.get(h, 14)
                lines      = max(1, int(len(val) * 2.2 / col_w) + 2)
                row_height = max(row_height, min(lines * 16, 300))

        ws.row_dimensions[row[0].row].height = row_height

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ── Sheet4/5 종합·종합(분기포함) ─────────────────────────────────────────────
def _style_finance_ws(ws):
    """Sheet4/5 종합·종합(분기포함) — 6_유튜브_AI종목스캐너 실적정보 방식 스타일·하이라이팅."""
    _HEADER_BREAKS = {
        "종합점수":          "종합\n점수",
        "현재가":            "현재\n가",
        "적정주가":          "적정\n주가",
        "상승여력1(%)":      "상승여력\n1(%)",
        "증권사목표주가":    "증권사\n목표주가",
        "상승여력2(%)":      "상승여력\n2(%)",
        "25년매출(억)":      "25년\n매출(억)",
        "26년매출E(억)":     "26년\n매출E(억)",
        "매출증가율(%)":     "매출\n증가율(%)",
        "25년영업익(억)":    "25년\n영업익(억)",
        "26년영업익E(억)":   "26년\n영업익E(억)",
        "영업익증가율(%)":   "영업익\n증가율(%)",
        "25년1Q매출(억)":    "25년1Q\n매출(억)",
        "26년1Q매출(억)":    "26년1Q\n매출(억)",
        "1Q매출증가율(%)":   "1Q매출\n증가율(%)",
        "25년1Q영업익(억)":  "25년1Q\n영업익(억)",
        "26년1Q영업익(억)":  "26년1Q\n영업익(억)",
        "1Q영업익증가율(%)": "1Q영업익\n증가율(%)",
        "25년2Q매출(억)":    "25년2Q\n매출(억)",
        "26년2QE매출(억)":   "26년2QE\n매출(억)",
        "2Q매출증가율(%)":   "2Q매출\n증가율(%)",
        "25년2Q영업익(억)":  "25년2Q\n영업익(억)",
        "26년2QE영업익(억)": "26년2QE\n영업익(억)",
        "2Q영업익증가율(%)": "2Q영업익\n증가율(%)",
        "12MPER":            "12M\nPER",
        "12M PER/PER":       "12M PER\n/PER",
    }
    for cell in ws[1]:
        if cell.value in _HEADER_BREAKS:
            cell.value = _HEADER_BREAKS[cell.value]

    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    ws.row_dimensions[1].height = 60
    for cell in ws[1]:
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border    = BORDER

    def _cw(val):
        return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

    for col in ws.columns:
        col_letter   = col[0].column_letter
        header_lines = str(col[0].value or "").split('\n')
        hw           = max([_cw(line) for line in header_lines] + [0])
        data_w       = max((_cw(cell.value) for cell in col[1:]), default=0)
        content_w    = max(data_w, hw)
        h_val        = str(col[0].value or "").replace('\n', '')
        if h_val == "종목명":
            width = min(max(content_w + 2, 12), 22)
        elif h_val == "티커":
            width = 8
        elif h_val == "업종":
            width = min(max(int(content_w * 0.8) + 1, 8), 16)
        elif h_val == "관련뉴스":
            width = min(max(content_w + 2, 15), 35)
        else:
            width = min(max(content_w + 2, 10), 18)
        ws.column_dimensions[col_letter].width = width

    col_ticker = col_idx.get("티커")
    col_up2    = col_idx.get("상승여력\n2(%)")
    col_per    = col_idx.get("PER")
    col_per12  = col_idx.get("12M\nPER")
    col_opg    = col_idx.get("영업익\n증가율(%)")

    num_col_names = {
        "종합\n점수", "현재\n가", "적정\n주가", "상승여력\n1(%)", "증권사\n목표주가", "상승여력\n2(%)",
        "25년\n매출(억)", "26년\n매출E(억)", "매출\n증가율(%)", "25년\n영업익(억)",
        "26년\n영업익E(억)", "영업익\n증가율(%)",
        "25년1Q\n매출(억)", "26년1Q\n매출(억)", "1Q매출\n증가율(%)",
        "25년1Q\n영업익(억)", "26년1Q\n영업익(억)", "1Q영업익\n증가율(%)",
        "25년2Q\n매출(억)", "26년2QE\n매출(억)", "2Q매출\n증가율(%)",
        "25년2Q\n영업익(억)", "26년2QE\n영업익(억)", "2Q영업익\n증가율(%)",
        "PER", "12M\nPER", "12M PER\n/PER", "PSR", "PBR", "PFR",
    }
    num_cols = {col_idx[c] for c in num_col_names if c in col_idx}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 18

        for cell in row:
            h      = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val    = str(cell.value).strip() if cell.value is not None else ""
            is_num = cell.column in num_cols
            wrap   = h == "관련뉴스"

            # 문자열로 저장된 숫자를 실제 숫자 타입으로 변환 (정렬·필터 작동)
            if is_num and val and val not in ("-", "N/A", "nan", ""):
                try:
                    fv = float(val.replace(",", ""))
                    cell.value = int(fv) if fv == int(fv) else fv
                    val = str(cell.value)
                except (ValueError, OverflowError):
                    pass

            cell.alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=wrap,
            )
            cell.border = BORDER

            if h == "종목명" and val:
                code  = ""
                if col_ticker:
                    tc   = ws.cell(row=cell.row, column=col_ticker)
                    code = str(tc.value).strip() if tc.value else ""
                code6 = code.zfill(6) if code and code not in ("", "nan") else ""
                if code6 and code6 != "000000":
                    cell.hyperlink = f"https://finance.naver.com/item/main.naver?code={code6}"
                else:
                    cell.hyperlink = f"https://finance.naver.com/search/searchList.naver?query={quote(val)}"
                cell.font = Font(color="0563C1", underline="single", bold=True)

            elif h == "티커" and val and val not in ("", "nan"):
                code6 = str(val).strip().zfill(6)
                cell.hyperlink = (
                    f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                    f"?pGB=1&gicode=A{code6}&cID=AA&MenuYn=Y&ReportGB="
                    f"&NewMenuID=11&stkGb=&strResearchYN="
                )
                cell.font = Font(color="0563C1", underline="single")

        v = _fval(row, col_up2)
        if v is not None and v >= PERF_MIN_UPSIDE2:
            row[col_up2 - 1].fill = FILL_HIGH

        for col in [col_per, col_per12]:
            if col:
                v = _fval(row, col)
                if v is not None and 0 < v < PERF_MAX_PER:
                    row[col - 1].fill = FILL_PER_LOW

        v = _fval(row, col_opg)
        if v is not None and v >= PERF_MIN_OP_GROWTH:
            row[col_opg - 1].fill = FILL_MID

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def format_file(path: str, insight_section_rows: tuple = (1, 0, 0, 0)):
    wb = openpyxl.load_workbook(path)

    if SHEET_INSIGHT in wb.sheetnames:
        _style_insight_ws(wb[SHEET_INSIGHT], insight_section_rows)
    if SHEET_STOCKS in wb.sheetnames:
        _expand_stock_rows(wb[SHEET_STOCKS])   # 종목별 행 분리 → 개별 하이퍼링크 적용
        _style_sheet(wb[SHEET_STOCKS], WIDTHS_STOCKS)
    if SHEET_MARKET in wb.sheetnames:
        _style_sheet(wb[SHEET_MARKET], WIDTHS_MARKET)
    if SHEET_BRIEF in wb.sheetnames:
        _style_finance_ws(wb[SHEET_BRIEF])
    if SHEET_FULL in wb.sheetnames:
        _style_finance_ws(wb[SHEET_FULL])

    base, ext = os.path.splitext(path)
    out_path  = f"{base}_fmt{ext}"
    wb.save(out_path)
    print(f"[OK] 포맷 완료: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        base  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        files = sorted([
            f for f in glob.glob(os.path.join(base, "뉴스기사_인사이트_*.xlsx"))
            if "~$" not in f and "_fmt" not in f
        ])
        if not files:
            print("[ERR] output 폴더에 xlsx 파일이 없습니다.")
            sys.exit(1)
        target = files[-1]
        print(f"대상 파일: {target}")

    format_file(target)
