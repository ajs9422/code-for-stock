"""
생성된 유튜브 인사이트 엑셀 파일을 정리합니다.
- 4개 시트 보존 및 각각 포맷 적용
  · Sheet1        : 종목추천 (행 분리·병합, 하이퍼링크, 색상)
  · 투자관련       : 투자 인사이트 (헤더 스타일, 링크)
  · 종합           : 실적 지표 — 분기 컬럼 제외 (4번 실적개선 방식)
  · 종합(분기포함) : 실적 지표 — 분기 컬럼 포함
- 컬럼 순서: B·C 앞으로, A → C 위치
- 종목명 → 네이버증권 / 종목코드·티커 → FnGuide 하이퍼링크
사용법: python format_excel.py [파일경로]  (생략 시 output 최신 파일 자동 선택)
"""
import sys
import glob
import os
import re
from urllib.parse import quote
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


def _build_stock_map() -> tuple[dict, dict]:
    """KRX 전 종목 기준 종목명→티커 맵 (FDR 우선, 4_financial_scanner 보조)."""
    raw: dict[str, str] = {}
    try:
        import FinanceDataReader as fdr
        df = fdr.StockListing("KRX")
        name_col = next((c for c in df.columns if c in ("Name", "종목명")), None)
        code_col = next((c for c in df.columns if c in ("Code", "Symbol")), None)
        if name_col and code_col:
            for _, row in df.iterrows():
                n = str(row[name_col]).strip()
                c = str(row[code_col]).strip().zfill(6)
                if n and c != "000000":
                    raw[n] = c
    except Exception:
        pass
    try:
        base  = os.path.join(os.path.dirname(__file__), "..", "4_financial_scanner", "output")
        files = sorted([f for f in glob.glob(os.path.join(base, "*.xlsx")) if "_hl" not in f])
        if files:
            df2 = pd.read_excel(files[-1], sheet_name=0, usecols=["종목명", "티커"], dtype=str)
            for n, t in zip(df2["종목명"].str.strip(), df2["티커"].str.strip()):
                if n not in raw:
                    raw[n] = t
    except Exception:
        pass
    norm: dict[str, tuple[str, str]] = {}
    for name, code in raw.items():
        norm[name.replace(" ", "").lower()] = (code, name)
    return raw, norm


def _lookup_ticker(gemini_name: str, raw_map: dict, norm_map: dict) -> tuple[str, str]:
    if not gemini_name:
        return "", gemini_name
    if gemini_name in raw_map:
        return raw_map[gemini_name], gemini_name
    key = gemini_name.replace(" ", "").lower()
    if key in norm_map:
        code, canonical = norm_map[key]
        return code, canonical
    best_code, best_name, best_len = "", gemini_name, 0
    for map_key, (code, canonical) in norm_map.items():
        if len(map_key) < 2:
            continue
        if key in map_key or map_key in key:
            if len(map_key) > best_len:
                best_code, best_name, best_len = code, canonical, len(map_key)
    return best_code, best_name


STOCK_RAW_MAP, STOCK_NORM_MAP = _build_stock_map()

# ── 컬럼 너비 ───────────────────────────────────────────────────────────────
COLUMN_WIDTHS_MAIN = {
    "업로드일시": 16, "제목": 28, "채널": 20, "출연자": 16,
    "요약": 80, "종목명": 12, "종목코드": 10,
    "추천사유": 72, "링크": 8,
}
COLUMN_WIDTHS_INV = {
    "업로드일시": 16, "제목": 32, "채널": 12, "출연자": 16,
    "요약": 80, "링크": 8,
}
COLUMN_WIDTHS_PERF = {
    "종목명": 18, "티커": 10, "추천채널": 35, "업종": 20,
    "종합점수": 11, "현재가": 12, "적정주가": 12,
    "상승여력1(%)": 14, "증권사목표주가": 16, "상승여력2(%)": 14,
    "25년매출(억)": 15, "26년매출E(억)": 16, "매출증가율(%)": 14,
    "25년영업익(억)": 16, "26년영업익E(억)": 17, "영업익증가율(%)": 17,
    "전년Q매출(억)": 14, "최근Q매출(억)": 14, "Q매출YoY(%)": 12,
    "전년Q영업익(억)": 15, "최근Q영업익(억)": 15, "Q영업익YoY(%)": 13,
    "전년QE매출(억)": 14, "차기QE매출(억)": 14, "차기Q매출YoY(%)": 14,
    "전년QE영업익(억)": 15, "차기QE영업익(억)": 15, "차기Q영업익YoY(%)": 15,
    "PER": 10, "12M PER": 10, "12M PER/PER": 13, "PSR": 10, "PBR": 10, "PFR": 10,
}

# ── 컬럼 순서: B·C 앞으로, A → C 위치 ──────────────────────────────────────
COL_ORDER_MAIN = ["업로드일시", "제목", "채널", "출연자", "요약", "종목명", "종목코드", "추천사유", "링크"]
COL_ORDER_INV  = ["업로드일시", "제목", "채널", "출연자", "요약", "링크"]
_QUARTER_COLS = [
    "전년Q매출(억)", "최근Q매출(억)", "Q매출YoY(%)",
    "전년Q영업익(억)", "최근Q영업익(억)", "Q영업익YoY(%)",
    "전년QE매출(억)", "차기QE매출(억)", "차기Q매출YoY(%)",
    "전년QE영업익(억)", "차기QE영업익(억)", "차기Q영업익YoY(%)",
]
_BASE_PERF_COLS = [
    "종목명", "티커", "추천채널", "업종", "종합점수",
    "현재가", "적정주가", "상승여력1(%)", "증권사목표주가", "상승여력2(%)",
    "25년매출(억)", "26년매출E(억)", "매출증가율(%)",
    "25년영업익(억)", "26년영업익E(억)", "영업익증가율(%)",
]
_VALUATION_COLS = ["PER", "12M PER", "12M PER/PER", "PSR", "PBR", "PFR"]
COL_ORDER_PERF      = _BASE_PERF_COLS + _VALUATION_COLS
COL_ORDER_PERF_FULL = _BASE_PERF_COLS + _QUARTER_COLS + _VALUATION_COLS

WRAP_COLS  = {"제목", "요약", "추천사유"}
MERGE_COLS = {"채널", "업로드일시", "제목", "출연자", "요약", "링크"}

# ── 공통 스타일 상수 ──────────────────────────────────────────────────────────
HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT  = Font(color="FFFFFF", bold=True, size=10)
HEADER_ALIGN = Alignment(horizontal="center", vertical="center", wrap_text=True)
THIN         = Side(style="thin", color="BBBBBB")
BORDER       = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# Sheet1: 티커 매칭 여부 색상
FILL_FOUND    = PatternFill("solid", fgColor="C6EFCE")  # 연초록 — 티커 매칭
FILL_NOTFOUND = PatternFill("solid", fgColor="FFEB9C")  # 연노랑 — 티커 미확인

# 실적정보: 4번 실적개선 동일 기준·색상
PERF_MIN_UPSIDE2   = 20.0
PERF_MAX_PER       = 12.0
PERF_MIN_OP_GROWTH = 30.0
FILL_UPSIDE2   = PatternFill("solid", fgColor="C6EFCE")  # 연초록 — 상승여력2
FILL_PER_LOW   = PatternFill("solid", fgColor="DDEBF7")  # 연파랑 — PER
FILL_OP_GROWTH = PatternFill("solid", fgColor="FFEB9C")  # 연노랑 — 영업익증가율

GROUP_KEY = ["채널", "업로드일시", "제목"]


def _read_with_links(path: str, sheet_name: str | None = None) -> pd.DataFrame:
    """엑셀 읽기 + 링크 컬럼의 하이퍼링크 실제 URL 복원."""
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet_name] if sheet_name else wb.active
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_data = {}
        for cell, header in zip(row, headers):
            val = str(cell.value) if cell.value is not None else ""
            if header == "링크":
                hl = cell.hyperlink
                if hl and getattr(hl, "target", None) and str(hl.target).startswith("http"):
                    val = hl.target
            row_data[header] = val
        rows.append(row_data)
    wb.close()
    return pd.DataFrame(rows).fillna("")


def _split_and_group(df: pd.DataFrame) -> list[list[dict]]:
    """종목명 쉼표 분리 후 같은 영상 기준으로 그룹화."""
    flat = []
    for _, row in df.iterrows():
        names = [n.strip() for n in re.split(r"[,/]", str(row.get("종목명", "")))
                 if n.strip() not in ("", "nan")]
        codes = [c.strip() for c in re.split(r"[,/]", str(row.get("종목코드", "")))
                 if c.strip() not in ("", "nan")]
        if len(names) <= 1:
            flat.append(row.to_dict())
        else:
            for i, name in enumerate(names):
                r = row.to_dict()
                r["종목명"]   = name
                r["종목코드"] = codes[i] if i < len(codes) else ""
                flat.append(r)

    groups: list[list[dict]] = []
    prev_key  = None
    cur: list[dict] = []
    empty_key = tuple("" for _ in GROUP_KEY)
    for r in flat:
        key = tuple(str(r.get(k, "")) for k in GROUP_KEY)
        # 채널·업로드일시·제목이 모두 비어있으면 바로 위 그룹의 연속 행으로 처리
        if key == empty_key and cur:
            first = cur[0]
            for k in GROUP_KEY:
                if not r.get(k):
                    r[k] = first.get(k, "")
            cur.append(r)
        elif key == prev_key:
            cur.append(r)
        else:
            if cur:
                groups.append(cur)
            cur      = [r]
            prev_key = key
    if cur:
        groups.append(cur)
    return groups


def _normalize_stock(name: str, code: str) -> tuple[str, str]:
    if not code.strip():
        new_code, canonical = _lookup_ticker(name, STOCK_RAW_MAP, STOCK_NORM_MAP)
        return canonical, new_code
    return name, code.strip()


def _stock_hyperlink(name: str, code: str) -> tuple[str, str]:
    code6 = code.strip().zfill(6) if code.strip() else ""
    if code6 and code6 != "000000":
        return f"https://finance.naver.com/item/main.naver?code={code6}", code6
    return f"https://finance.naver.com/search/searchList.naver?query={quote(name)}", ""


def _fval(row, col: int | None):
    """행의 col 번째 셀을 float로 반환 (숫자가 아니면 None)."""
    if not col:
        return None
    v = row[col - 1].value
    return float(v) if isinstance(v, (int, float)) else None


def _apply_header(ws, col_idx: dict, widths: dict) -> None:
    ws.row_dimensions[1].height = 22
    for h, idx in col_idx.items():
        cell = ws.cell(row=1, column=idx)
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border    = BORDER
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(h, 14)


# ────────────────────────────────────────────────────────────────────────────
#  시트별 스타일 함수
# ────────────────────────────────────────────────────────────────────────────

def _style_main_ws(ws, groups: list[list[dict]]) -> None:
    """Sheet1 종목추천 스타일."""
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers)}

    _apply_header(ws, col_idx, COLUMN_WIDTHS_MAIN)

    data_row = 2
    for group in groups:
        group_size = len(group)
        start_row  = data_row
        end_row    = data_row + group_size - 1
        multi      = group_size > 1

        for g_idx, row_dict in enumerate(group):
            excel_row  = data_row + g_idx
            row_height = 14

            for h, cidx in col_idx.items():
                cell = ws.cell(row=excel_row, column=cidx)
                val  = str(cell.value).strip() if cell.value else ""
                wrap = h in WRAP_COLS

                cell.alignment = Alignment(vertical="center", wrap_text=wrap, shrink_to_fit=not wrap)
                cell.border    = BORDER

                if h == "종목명" and val:
                    code_val             = str(row_dict.get("종목코드", "")).strip()
                    clean_name, code_val = _normalize_stock(val, code_val)
                    cell.value           = clean_name
                    url, _               = _stock_hyperlink(clean_name, code_val)
                    cell.hyperlink       = url
                    cell.font            = Font(color="0563C1", underline="single", bold=True)
                    cell.fill            = FILL_FOUND if code_val else FILL_NOTFOUND

                    if code_val and "종목코드" in col_idx:
                        code6     = code_val.zfill(6)
                        code_cell = ws.cell(row=excel_row, column=col_idx["종목코드"])
                        code_cell.value     = code6
                        code_cell.hyperlink = (
                            f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                            f"?pGB=1&gicode=A{code6}&cID=AA&MenuYn=Y&ReportGB="
                            f"&NewMenuID=11&stkGb=&strResearchYN="
                        )
                        code_cell.font      = Font(color="0563C1", underline="single")
                        code_cell.alignment = Alignment(vertical="center")
                        code_cell.border    = BORDER
                        code_cell.fill      = FILL_FOUND

                elif h == "종목코드" and val:
                    code6 = val.zfill(6)
                    cell.hyperlink = (
                        f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                        f"?pGB=1&gicode=A{code6}&cID=AA&MenuYn=Y&ReportGB="
                        f"&NewMenuID=11&stkGb=&strResearchYN="
                    )
                    cell.font = Font(color="0563C1", underline="single")
                    cell.fill = FILL_FOUND

                elif h == "링크" and val:
                    cell.hyperlink = val
                    cell.value     = "링크"
                    cell.font      = Font(color="0563C1", underline="single")

                if wrap and val:
                    if multi and h in MERGE_COLS:
                        continue
                    col_w = COLUMN_WIDTHS_MAIN.get(h, 14)
                    lines = max(1, int(len(val) * 2.2 / col_w) + 1)
                    row_height = max(row_height, min(lines * 15, 100))

            ws.row_dimensions[excel_row].height = row_height

        if group_size > 1:
            for h in MERGE_COLS:
                if h not in col_idx:
                    continue
                cidx       = col_idx[h]
                col_letter = get_column_letter(cidx)
                ws.merge_cells(f"{col_letter}{start_row}:{col_letter}{end_row}")
                mc = ws.cell(row=start_row, column=cidx)
                mc.alignment = Alignment(vertical="center", wrap_text=(h in WRAP_COLS))
                mc.border    = BORDER
                if h == "링크":
                    val = str(mc.value).strip() if mc.value else ""
                    if val and val != "링크":
                        mc.hyperlink = val
                        mc.value     = "링크"
                        mc.font      = Font(color="0563C1", underline="single")

        data_row = end_row + 1

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _style_invest_ws(ws) -> None:
    """투자관련 시트 스타일."""
    headers = [cell.value for cell in ws[1]]
    col_idx = {h: i + 1 for i, h in enumerate(headers) if h}

    _apply_header(ws, col_idx, COLUMN_WIDTHS_INV)

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        row_height = 14
        for cell in row:
            h   = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val = str(cell.value).strip() if cell.value else ""
            wrap = h in {"제목", "요약"}
            cell.alignment = Alignment(vertical="center", wrap_text=wrap, shrink_to_fit=not wrap)
            cell.border    = BORDER
            if h == "링크" and val:
                cell.hyperlink = val
                cell.value     = "링크"
                cell.font      = Font(color="0563C1", underline="single")
            if wrap and val:
                col_w = COLUMN_WIDTHS_INV.get(h, 14)
                lines = max(1, int(len(val) * 2.2 / col_w) + 2)
                row_height = max(row_height, min(lines * 16, 120))
        ws.row_dimensions[row[0].row].height = row_height

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


def _style_perf_ws(ws) -> None:
    """실적정보 시트 — 4번 실적개선 방식 스타일·하이라이팅."""
    # 헤더에 줄바꿈 삽입
    _HEADER_BREAKS = {
        "종합점수":           "종합\n점수",
        "현재가":             "현재\n가",
        "적정주가":           "적정\n주가",
        "상승여력1(%)":       "상승여력\n1(%)",
        "증권사목표주가":     "증권사\n목표주가",
        "상승여력2(%)":       "상승여력\n2(%)",
        "25년매출(억)":       "25년\n매출(억)",
        "26년매출E(억)":      "26년\n매출E(억)",
        "매출증가율(%)":      "매출\n증가율(%)",
        "25년영업익(억)":     "25년\n영업익(억)",
        "26년영업익E(억)":    "26년\n영업익E(억)",
        "영업익증가율(%)":    "영업익\n증가율(%)",
        "전년Q매출(억)":      "전년Q\n매출(억)",
        "최근Q매출(억)":      "최근Q\n매출(억)",
        "Q매출YoY(%)":        "Q매출\nYoY(%)",
        "전년Q영업익(억)":    "전년Q\n영업익(억)",
        "최근Q영업익(억)":    "최근Q\n영업익(억)",
        "Q영업익YoY(%)":      "Q영업익\nYoY(%)",
        "전년QE매출(억)":     "전년QE\n매출(억)",
        "차기QE매출(억)":     "차기QE\n매출(억)",
        "차기Q매출YoY(%)":    "차기Q매출\nYoY(%)",
        "전년QE영업익(억)":   "전년QE\n영업익(억)",
        "차기QE영업익(억)":   "차기QE\n영업익(억)",
        "차기Q영업익YoY(%)":  "차기Q영업익\nYoY(%)",
        "12M PER":            "12M\nPER",
        "12M PER/PER":        "12M PER\n/PER",
    }
    for cell in ws[1]:
        if cell.value in _HEADER_BREAKS:
            cell.value = _HEADER_BREAKS[cell.value]

    headers  = [cell.value for cell in ws[1]]
    col_idx  = {h: i + 1 for i, h in enumerate(headers) if h}

    # 기본 헤더 스타일 및 행 높이 설정
    ws.row_dimensions[1].height = 60
    for cell in ws[1]:
        cell.fill      = HEADER_FILL
        cell.font      = HEADER_FONT
        cell.alignment = HEADER_ALIGN
        cell.border    = BORDER

    # 열 너비 자동 조정
    def _cw(val):
        return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

    for col in ws.columns:
        col_letter = col[0].column_letter
        header_lines = str(col[0].value or "").split('\n')
        hw = max([_cw(line) for line in header_lines] + [0])
        data_w = max((_cw(cell.value) for cell in col[1:]), default=0)
        content_w = max(data_w, hw)

        h_val = str(col[0].value).replace('\n', '')
        if h_val == "종목명":
            width = min(max(content_w + 2, 12), 22)
        elif h_val == "티커":
            width = 8
        elif h_val == "업종":
            width = min(max(int(content_w * 0.8) + 1, 8), 16)
        elif h_val == "추천채널":
            width = min(max(content_w + 2, 15), 35)
        else: # 숫자 컬럼
            width = min(max(content_w + 2, 10), 18)
        ws.column_dimensions[col_letter].width = width

    col_name   = col_idx.get("종목명")
    col_ticker = col_idx.get("티커")
    col_up2    = col_idx.get("상승여력\n2(%)")
    col_per    = col_idx.get("PER")
    col_per12  = col_idx.get("12M\nPER")
    col_opg    = col_idx.get("영업익\n증가율(%)")

    num_col_names = {
        "종합\n점수", "현재\n가", "적정\n주가", "상승여력\n1(%)", "증권사\n목표주가", "상승여력\n2(%)",
        "25년\n매출(억)", "26년\n매출E(억)", "매출\n증가율(%)", "25년\n영업익(억)",
        "26년\n영업익E(억)", "영업익\n증가율(%)",
        "전년Q\n매출(억)", "최근Q\n매출(억)", "Q매출\nYoY(%)",
        "전년Q\n영업익(억)", "최근Q\n영업익(억)", "Q영업익\nYoY(%)",
        "전년QE\n매출(억)", "차기QE\n매출(억)", "차기Q매출\nYoY(%)",
        "전년QE\n영업익(억)", "차기QE\n영업익(억)", "차기Q영업익\nYoY(%)",
        "PER", "12M\nPER", "12M PER\n/PER", "PSR", "PBR", "PFR",
    }
    num_cols = {col_idx[c] for c in num_col_names if c in col_idx}

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        ws.row_dimensions[row[0].row].height = 18

        for cell in row:
            h   = headers[cell.column - 1] if cell.column <= len(headers) else ""
            val = str(cell.value).strip() if cell.value is not None else ""
            wrap   = h == "추천채널"
            is_num = cell.column in num_cols

            cell.alignment = Alignment(
                vertical="center",
                horizontal="center",
                wrap_text=wrap,
            )
            cell.border = BORDER

            if h == "종목명" and val:
                code = ""
                if col_ticker:
                    tc   = ws.cell(row=cell.row, column=col_ticker)
                    code = str(tc.value).strip() if tc.value else ""
                code6 = code.zfill(6) if code and code != "nan" else ""
                if code6 and code6 != "000000":
                    cell.hyperlink = f"https://finance.naver.com/item/main.naver?code={code6}"
                else:
                    cell.hyperlink = f"https://finance.naver.com/search/searchList.naver?query={quote(val)}"
                cell.font = Font(color="0563C1", underline="single", bold=True)

            elif h == "티커" and val and val != "nan":
                code6 = str(val).strip().zfill(6)
                cell.hyperlink = (
                    f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                    f"?pGB=1&gicode=A{code6}&cID=AA&MenuYn=Y&ReportGB="
                    f"&NewMenuID=11&stkGb=&strResearchYN="
                )
                cell.font = Font(color="0563C1", underline="single")

        # 셀 색상 — 4번 실적개선 동일 기준
        v = _fval(row, col_up2)
        if v is not None and v >= PERF_MIN_UPSIDE2:
            row[col_up2 - 1].fill = FILL_UPSIDE2

        for col in [col_per, col_per12]:
            if col:
                v = _fval(row, col)
                if v is not None and v > 0 and v < PERF_MAX_PER:
                    row[col - 1].fill = FILL_PER_LOW

        v = _fval(row, col_opg)
        if v is not None and v >= PERF_MIN_OP_GROWTH:
            row[col_opg - 1].fill = FILL_OP_GROWTH

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions


# ────────────────────────────────────────────────────────────────────────────
#  메인 포맷 함수
# ────────────────────────────────────────────────────────────────────────────

def format_file(path: str):
    # 1. 원본 시트 목록 확인
    wb_src     = openpyxl.load_workbook(path, read_only=True)
    src_sheets = wb_src.sheetnames
    wb_src.close()
    main_sheet = src_sheets[0]
    has_invest = "투자관련" in src_sheets
    has_brief  = "종합" in src_sheets
    has_full   = "종합(분기포함)" in src_sheets

    # 2. Sheet1: 링크 복원 → 행 분리·그룹화 → 컬럼 순서 변경
    df_main = _read_with_links(path, main_sheet)
    groups  = _split_and_group(df_main)
    flat    = [row for group in groups for row in group]
    df_main = pd.DataFrame(flat).reset_index(drop=True)
    ordered = [c for c in COL_ORDER_MAIN if c in df_main.columns]
    df_main = df_main[ordered + [c for c in df_main.columns if c not in ordered]]

    # 3. 투자관련: 링크 복원 → 컬럼 순서 변경
    df_inv = None
    if has_invest:
        df_inv  = _read_with_links(path, "투자관련")
        ordered = [c for c in COL_ORDER_INV if c in df_inv.columns]
        df_inv  = df_inv[ordered + [c for c in df_inv.columns if c not in ordered]]

    # 4. 종합/종합(분기포함): 티커 zfill + 추천채널을 티커 바로 뒤로 이동
    # (4번 실적개선 시트는 헤더에 줄바꿈이 이미 저장돼 있으므로 원래 순서를 유지)
    def _prepare_perf_df(sheet_name: str):
        df = pd.read_excel(path, sheet_name=sheet_name, dtype={"티커": str})
        if "티커" in df.columns:
            df["티커"] = (
                df["티커"].fillna("").astype(str).str.strip()
                .apply(lambda x: x.zfill(6) if x and x != "nan" else x)
            )
        if "추천채널" in df.columns and "티커" in df.columns:
            cols = [c for c in df.columns if c != "추천채널"]
            cols.insert(cols.index("티커") + 1, "추천채널")
            df = df[cols]
        return df

    df_brief = _prepare_perf_df("종합") if has_brief else None
    df_full  = _prepare_perf_df("종합(분기포함)") if has_full else None

    # 5. 전체 시트를 output 파일로 저장 (Sheet1 → 유튜브_Summary 로 이름 고정)
    MAIN_SHEET_NAME = "유튜브_Summary"
    base, ext = os.path.splitext(path)
    out_path  = f"{base}_fmt{ext}"
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_main.to_excel(writer, sheet_name=MAIN_SHEET_NAME, index=False)
        if df_inv is not None:
            df_inv.to_excel(writer, sheet_name="투자관련", index=False)
        if df_brief is not None:
            df_brief.to_excel(writer, sheet_name="종합", index=False)
        if df_full is not None:
            df_full.to_excel(writer, sheet_name="종합(분기포함)", index=False)

    # 6. openpyxl 스타일 적용
    wb = openpyxl.load_workbook(out_path)
    _style_main_ws(wb[MAIN_SHEET_NAME], groups)
    if has_invest:
        _style_invest_ws(wb["투자관련"])
    if has_brief:
        _style_perf_ws(wb["종합"])
    if has_full:
        _style_perf_ws(wb["종합(분기포함)"])
    wb.save(out_path)

    # 7. xlwings AutoFit (Sheet1 행 높이 자동 조정)
    try:
        import xlwings as xw
        with xw.App(visible=False) as app:
            wb2 = app.books.open(os.path.abspath(out_path))
            ws2 = wb2.sheets[0]
            ws2.autofit("rows")

            # 세로 병합 셀 높이 보정 (autofit은 병합 셀을 처리 못하므로 수동 계산)
            wb_tmp = openpyxl.load_workbook(out_path)
            ws_tmp = wb_tmp.worksheets[0]
            for mr in ws_tmp.merged_cells.ranges:
                if mr.max_row <= mr.min_row:
                    continue
                cell = ws_tmp.cell(mr.min_row, mr.min_col)
                if not cell.value:
                    continue
                col_w = ws_tmp.column_dimensions[get_column_letter(mr.min_col)].width or 10
                chars = len(str(cell.value))
                lines = max(1, int(chars * 2.2 / col_w) + 1)
                needed = lines * 15
                current = sum(
                    ws2.range(f"A{r}").row_height or 15
                    for r in range(mr.min_row, mr.max_row + 1)
                )
                if current < needed:
                    first_h = ws2.range(f"A{mr.min_row}").row_height or 15
                    ws2.range(f"A{mr.min_row}").row_height = first_h + (needed - current)
            wb_tmp.close()

            wb2.save()
            wb2.close()
        print("[OK] 행 높이 자동 조정 완료")
    except Exception as e:
        print(f"[WARN] AutoFit 실패 (xlwings): {e}")

    print(f"[OK] 포맷 완료: {out_path}")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        target = sys.argv[1]
    else:
        base  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        files = sorted([f for f in glob.glob(os.path.join(base, "유튜브_AI종목스캐너_*.xlsx"))
                        if "~$" not in f and "_fmt" not in f])
        if not files:
            print("[ERR] output 폴더에 xlsx 파일이 없습니다.")
            sys.exit(1)
        target = files[-1]
        print(f"대상 파일: {target}")

    format_file(target)
