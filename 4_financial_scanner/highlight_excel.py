"""
════════════════════════════════════════════════════════════════
  실적개선주 스캐너 — Excel 조건부 하이라이팅
════════════════════════════════════════════════════════════════

사용법:
  python highlight_excel.py                  # output 폴더의 가장 최신 파일 자동 선택
  python highlight_excel.py 파일경로.xlsx    # 특정 파일 지정

════════════════════════════════════════════════════════════════
"""

# ══════════════════════════════════════════════
#   하이라이팅 조건  (None = 비활성화)
# ══════════════════════════════════════════════

# ① 상승여력2(%) — 증권사 목표주가 기준 상승여력
#   이 값 이상이면 해당 셀 강조
MIN_UPSIDE2: float | None = 20.0

# ② PER / 12M PER — 둘 중 하나라도 이 값 미만이면 해당 셀(들) 강조
#   None 으로 설정하면 이 조건 비활성화
MAX_PER: float | None = 12.0

# ③ 영업익증가율(%) — 이 값 이상이면 해당 셀 강조
MIN_OP_GROWTH: float | None = 30.0

# ④ PEG — 이 값 이하이면 해당 셀 강조 (성장 대비 저평가)
MAX_PEG: float | None = 0.5


# ══════════════════════════════════════════════
#   하이라이팅 색상  (6자리 HEX, RRGGBB)
# ══════════════════════════════════════════════

COLOR_UPSIDE2   = "C6EFCE"   # 연초록  — 상승여력2
COLOR_PER       = "DDEBF7"   # 연파랑  — PER / 12M PER
COLOR_OP_GROWTH = "FFEB9C"   # 연노랑  — 영업익증가율
COLOR_PEG       = "E2CFEF"   # 연보라  — PEG


# ─────────────────────────────────────────────────────────────
#   이하 수정 불필요
# ─────────────────────────────────────────────────────────────

import copy
import re
import sys
from pathlib import Path


def _latest_excel(output_dir: Path) -> Path | None:
    files = sorted(output_dir.glob("재무스캐너_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return files[0] if files else None


def _cell_val(row, col_idx) -> float | None:
    if col_idx is None:
        return None
    v = row[col_idx - 1].value
    return float(v) if isinstance(v, (int, float)) else None


def _norm(s) -> str:
    """공백·줄바꿈 모두 제거 — _HEADER_BREAKS 로 변환된 헤더도 매칭."""
    return re.sub(r"\s+", "", str(s or ""))


def _find_col(header: dict, name: str):
    """header dict 에서 name 에 해당하는 컬럼 인덱스 반환.

    main.py 의 _HEADER_BREAKS 가 헤더에 줄바꿈을 삽입하므로
    공백·줄바꿈을 제거한 뒤 비교한다.
    PEG 처럼 뒤에 부연 설명이 붙은 헤더는 전방 일치로 처리.
    """
    norm_name = _norm(name)
    # 1순위: 정규화 후 정확 일치
    for k, v in header.items():
        if _norm(k) == norm_name:
            return v
    # 2순위: 정규화 후 전방 일치 (PEG → "PEG(12MPER/이익증가율)" 등)
    for k, v in header.items():
        if _norm(k).startswith(norm_name):
            return v
    return None


def _passes_all(row, col_upside2, col_per, col_per_12m, col_op_growth) -> bool:
    """활성화된 모든 조건을 AND 로 검사."""
    checks: list[bool] = []

    if MIN_UPSIDE2 is not None and col_upside2:
        v = _cell_val(row, col_upside2)
        checks.append(v is not None and v >= MIN_UPSIDE2)

    if MAX_PER is not None and (col_per or col_per_12m):
        per_ok   = (v := _cell_val(row, col_per))     is not None and v > 0 and v < MAX_PER if col_per    else False
        per12_ok = (v := _cell_val(row, col_per_12m)) is not None and v > 0 and v < MAX_PER if col_per_12m else False
        checks.append(per_ok or per12_ok)

    if MIN_OP_GROWTH is not None and col_op_growth:
        v = _cell_val(row, col_op_growth)
        checks.append(v is not None and v >= MIN_OP_GROWTH)

    return bool(checks) and all(checks)


def _apply_highlight(ws) -> None:
    """워크시트 전체에 개별 조건 하이라이팅 적용."""
    from openpyxl.styles import PatternFill

    def _fill(color: str) -> PatternFill:
        return PatternFill("solid", fgColor=color)

    header = {cell.value: cell.column for cell in ws[1] if cell.value}
    col_upside2   = _find_col(header, "상승여력2(%)")
    col_per       = _find_col(header, "PER")
    col_per_12m   = _find_col(header, "12M PER")
    col_op_growth = _find_col(header, "영업익증가율(%)")
    col_peg       = _find_col(header, "PEG")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        if MIN_UPSIDE2 is not None and col_upside2:
            v = _cell_val(row, col_upside2)
            if v is not None and v >= MIN_UPSIDE2:
                row[col_upside2 - 1].fill = _fill(COLOR_UPSIDE2)

        if MAX_PER is not None:
            for col_idx in [col_per, col_per_12m]:
                if col_idx:
                    v = _cell_val(row, col_idx)
                    if v is not None and v > 0 and v < MAX_PER:
                        row[col_idx - 1].fill = _fill(COLOR_PER)

        if MIN_OP_GROWTH is not None and col_op_growth:
            v = _cell_val(row, col_op_growth)
            if v is not None and v >= MIN_OP_GROWTH:
                row[col_op_growth - 1].fill = _fill(COLOR_OP_GROWTH)

        if MAX_PEG is not None and col_peg:
            v = _cell_val(row, col_peg)
            if v is not None and v > 0 and v <= MAX_PEG:
                row[col_peg - 1].fill = _fill(COLOR_PEG)


def _build_and_sheet(wb, src_ws) -> None:
    """AND 조건 부합 종목만 모은 새 시트를 워크북 맨 앞에 추가."""
    header = {cell.value: cell.column for cell in src_ws[1] if cell.value}
    col_upside2   = _find_col(header, "상승여력2(%)")
    col_per       = _find_col(header, "PER")
    col_per_12m   = _find_col(header, "12M PER")
    col_op_growth = _find_col(header, "영업익증가율(%)")

    # AND 조건 부합 행 수집
    matched = [
        row for row in src_ws.iter_rows(min_row=2, max_row=src_ws.max_row)
        if _passes_all(row, col_upside2, col_per, col_per_12m, col_op_growth)
    ]

    sheet_name = "★AND조건"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]

    src_idx = wb.sheetnames.index(src_ws.title)
    if "종합(분기포함)" in wb.sheetnames:
        insert_idx = wb.sheetnames.index("종합(분기포함)") + 1
    else:
        insert_idx = src_idx + 1
    ws_and = wb.create_sheet(sheet_name, insert_idx)

    # 헤더 복사 (서식 포함)
    for src_cell in src_ws[1]:
        dst = ws_and.cell(row=1, column=src_cell.column, value=src_cell.value)
        if src_cell.has_style:
            dst.font      = copy.copy(src_cell.font)
            dst.fill      = copy.copy(src_cell.fill)
            dst.alignment = copy.copy(src_cell.alignment)

    # 데이터 행 복사 (하이라이팅 서식 + 하이퍼링크 + 값 포함)
    for dst_row_idx, src_row in enumerate(matched, start=2):
        for src_cell in src_row:
            dst = ws_and.cell(row=dst_row_idx, column=src_cell.column, value=src_cell.value)
            if src_cell.has_style:
                dst.font          = copy.copy(src_cell.font)
                dst.fill          = copy.copy(src_cell.fill)
                dst.alignment     = copy.copy(src_cell.alignment)
                dst.number_format = src_cell.number_format
            if src_cell.hyperlink:
                dst.hyperlink = copy.copy(src_cell.hyperlink)

    # 열 너비 / 행 높이 동기화
    for col in src_ws.column_dimensions:
        ws_and.column_dimensions[col].width = src_ws.column_dimensions[col].width
    ws_and.row_dimensions[1].height = src_ws.row_dimensions[1].height or 45

    ws_and.freeze_panes = "A2"
    ws_and.auto_filter.ref = ws_and.dimensions

    print(f"  → ★AND조건 시트: {len(matched)}개 종목")


def highlight(path: Path) -> None:
    from openpyxl import load_workbook

    wb = load_workbook(path)

    # 1. 모든 시트에 개별 조건 하이라이팅
    for ws in wb.worksheets:
        _apply_highlight(ws)

    # 2. 종합 시트 기준으로 AND 조건 시트 생성
    src = wb["종합"] if "종합" in wb.sheetnames else wb.worksheets[0]
    _build_and_sheet(wb, src)

    # 3. AND조건 시트는 생성 후에 만들어지므로 하이라이팅 별도 적용
    if "★AND조건" in wb.sheetnames:
        _apply_highlight(wb["★AND조건"])

    wb.save(path)
    print(f"저장 완료: {path}")
    return path


def main() -> None:
    if len(sys.argv) >= 2:
        target = Path(sys.argv[1])
        if not target.exists():
            print(f"파일을 찾을 수 없습니다: {target}")
            sys.exit(1)
    else:
        output_dir = Path(__file__).parent / "output"
        target = _latest_excel(output_dir)
        if target is None:
            print("output 폴더에 재무스캐너_*.xlsx 파일이 없습니다.")
            sys.exit(1)
        print(f"대상 파일: {target}")

    highlight(target)


if __name__ == "__main__":
    main()
