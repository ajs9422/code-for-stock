"""
================================================
  전체 분석 순차 실행 런처
  0번 ~ 5번 스크립트를 순서대로 실행
  생성된 Excel 파일은 output/ 폴더로 자동 이동
================================================
"""

import io
import os
import shutil
import smtplib
import subprocess
import sys
import time
import zipfile
from datetime import datetime
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path


from dotenv import load_dotenv
import re as _re
load_dotenv(Path(__file__).parent / ".env")

from rich import box
from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.text import Text

# Windows cp949 터미널에서 유니코드 출력 강제
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

console = Console(highlight=False)

SCRIPTS = [
    "1_가치주_찾기.py",
    "2_괴리율_15_목표가상향.py",
    "3_이익추정치_모멘텀.py",
    "4_수급의질과신용잔고분석.py",
    "5_전문가추천_저평가종목.py",
    "6_NaverNewsStockFinder.py",
]

BASE_DIR   = Path(__file__).parent
OUTPUT_DIR: Path = BASE_DIR / "output"  # main()에서 타임스탬프로 교체됨


def existing_xlsx() -> set[Path]:
    """현재 BASE_DIR 내 xlsx 파일 목록"""
    return set(BASE_DIR.glob("*.xlsx"))


def move_new_files(before: set[Path]) -> list[Path]:
    """실행 후 새로 생긴 xlsx 파일을 output/ 로 이동, 이동한 파일 목록 반환"""
    after   = existing_xlsx()
    new     = after - before
    moved   = []
    for f in sorted(new):
        dest = OUTPUT_DIR / f.name
        shutil.move(str(f), str(dest))
        moved.append(dest)
    return moved


def run_script(script: str) -> tuple[int, float, str, list[Path]]:
    """스크립트 실행 후 (종료코드, 소요초, 오류메시지, 저장파일목록) 반환"""
    path   = BASE_DIR / script
    before = existing_xlsx()
    start  = time.time()
    try:
        import os
        env = os.environ.copy()
        env["PYTHONUTF8"]       = "1"
        env["PYTHONIOENCODING"] = "utf-8"

        result = subprocess.run(
            [sys.executable, "-X", "utf8", str(path)],
            cwd=str(BASE_DIR),
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
            env=env,
        )
        elapsed = time.time() - start
        err_msg = result.stderr.strip()[-300:] if result.returncode != 0 else ""
        moved   = move_new_files(before)
        return result.returncode, elapsed, err_msg, moved
    except Exception as e:
        elapsed = time.time() - start
        return -1, elapsed, str(e), []


def fmt_time(seconds: float) -> str:
    m, s = divmod(int(seconds), 60)
    return f"{m}분 {s:02d}초" if m else f"{s}초"


def merge_excel_files(output_dir: Path, timestamp: str) -> Path | None:
    """output_dir 내 xlsx 파일을 시트별로 합쳐 통합 Excel 생성"""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Font, PatternFill
        from openpyxl.utils import get_column_letter
    except ImportError:
        console.print("[yellow]openpyxl 없음 - 통합 Excel 생성 생략[/yellow]")
        return None

    xlsx_files = sorted(output_dir.glob("*.xlsx"))
    if not xlsx_files:
        return None

    wb_merged = openpyxl.Workbook()
    wb_merged.remove(wb_merged.active)  # 기본 시트 제거

    for xlsx_path in xlsx_files:
        # 시트명: 파일명에서 날짜 postfix 제거 (예: 0_가치주_찾기)
        sheet_name = xlsx_path.stem
        for suffix in [f"_{timestamp}", f"_{timestamp[:8]}"]:
            sheet_name = sheet_name.replace(suffix, "")
        sheet_name = sheet_name[:31]  # Excel 시트명 31자 제한

        try:
            wb_src = openpyxl.load_workbook(xlsx_path)
            ws_src = wb_src.active
            ws_dst = wb_merged.create_sheet(title=sheet_name)

            for row in ws_src.iter_rows():
                for cell in row:
                    new_cell = ws_dst.cell(row=cell.row, column=cell.column, value=cell.value)
                    if cell.has_style:
                        from copy import copy
                        new_cell.font      = copy(cell.font)
                        new_cell.alignment = copy(cell.alignment)
                        new_cell.number_format = cell.number_format
                    if cell.hyperlink:
                        new_cell.hyperlink = cell.hyperlink

            # 열 너비 복사
            for col_letter, dim in ws_src.column_dimensions.items():
                ws_dst.column_dimensions[col_letter].width = dim.width

            # 헤더 배경색
            header_fill = PatternFill("solid", fgColor="1F4E79")
            header_font = Font(bold=True, color="FFFFFF")
            for cell in ws_dst[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")

        except Exception as e:
            console.print(f"[yellow]시트 추가 실패 ({xlsx_path.name}): {e}[/yellow]")

    if not wb_merged.sheetnames:
        return None

    merged_path = output_dir / f"통합분석_{timestamp}.xlsx"
    wb_merged.save(merged_path)
    return merged_path


def main():
    global OUTPUT_DIR
    OUTPUT_DIR = BASE_DIR / f"output_{datetime.now().strftime('%Y_%m%d_%H%M')}"
    OUTPUT_DIR.mkdir(exist_ok=True)

    console.print()
    console.print(Panel.fit(
        "[bold cyan]전체 분석 순차 실행[/bold cyan]\n"
        f"[dim]{len(SCRIPTS)}개 스크립트를 순서대로 실행합니다[/dim]\n"
        f"[dim]출력 파일 → output/[/dim]",
        border_style="cyan",
    ))
    console.print()

    results     = []
    total_start = time.time()

    for idx, script in enumerate(SCRIPTS, 1):
        console.rule(f"[cyan][{idx}/{len(SCRIPTS)}] {script}[/cyan]")
        console.print(f"  [dim]시작: {datetime.now().strftime('%H:%M:%S')}[/dim]")
        console.print(Panel.fit(
            f"[bold yellow]▶  [{idx}/{len(SCRIPTS)}]  {script}[/bold yellow]",
            border_style="yellow",
        ))

        code, elapsed, err, moved = run_script(script)

        if code == 0:
            console.print(f"  [green]OK 완료 ({fmt_time(elapsed)})[/green]")
            for f in moved:
                console.print(f"  [cyan]  → output/{f.name}[/cyan]")
            status = "OK 완료"
        else:
            console.print(f"  [red]NG 실패 ({fmt_time(elapsed)})[/red]")
            if err:
                console.print(f"  [red dim]{err}[/red dim]")
            status = "NG 실패"

        results.append((script, status, fmt_time(elapsed), code))
        console.print()

    # ── 최종 요약 ──
    total_elapsed = time.time() - total_start
    console.rule("[bold cyan]실행 결과 요약[/bold cyan]")
    console.print()

    table = Table(box=box.SIMPLE_HEAD, header_style="bold cyan", border_style="dim")
    table.add_column("#",        width=3,  justify="right", style="dim")
    table.add_column("스크립트", width=35)
    table.add_column("상태",     width=10, justify="center")
    table.add_column("소요시간", width=10, justify="right")

    for i, (script, status, elapsed, code) in enumerate(results, 1):
        status_text = (
            Text(status, style="green") if code == 0
            else Text(status, style="red")
        )
        table.add_row(str(i), script, status_text, elapsed)

    console.print(table)

    success = sum(1 for _, _, _, c in results if c == 0)
    fail    = len(results) - success
    console.print(
        f"  완료: [green]{success}개[/green]  "
        f"실패: [red]{fail}개[/red]  "
        f"총 소요시간: [cyan]{fmt_time(total_elapsed)}[/cyan]"
    )
    console.print(f"  저장 위치: [cyan]{OUTPUT_DIR}[/cyan]")
    console.print()

    # ── 통합 Excel 생성 ──
    timestamp = OUTPUT_DIR.name.replace("output_", "")
    merged = None
    if success > 0:
        console.rule("[bold cyan]통합 Excel 생성[/bold cyan]")
        merged = merge_excel_files(OUTPUT_DIR, timestamp)
        if merged:
            console.print(f"  [green]통합 Excel 저장 완료: {merged.name}[/green]")
            console.print(f"  [cyan]{merged}[/cyan]")
        console.print()

    # ── 이메일 발송 ──
    console.rule("[bold cyan]이메일 발송[/bold cyan]")
    send_result_email(OUTPUT_DIR, timestamp, results)
    console.print()


def send_result_email(output_dir: Path, timestamp: str, results: list) -> None:
    gmail_user    = os.environ.get("GMAIL_USER", "")
    gmail_pw      = os.environ.get("GMAIL_APP_PASSWORD", "")
    _extra        = [e.strip() for e in _re.split(r"[,;]", os.environ.get("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
    _main         = [e for e in [os.environ.get("NOTIFY_EMAIL", ""), os.environ.get("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
    notify_emails = _main + _extra

    if not (gmail_user and gmail_pw and notify_emails):
        console.print("[yellow]이메일 설정 없음 (.env 확인) — 발송 생략[/yellow]")
        return

    # output 폴더 zip 압축
    zip_path = output_dir.parent / f"분석결과_{timestamp}.zip"
    try:
        with zipfile.ZipFile(zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
            for f in sorted(output_dir.glob("*.xlsx")):
                zf.write(f, f.name)
        console.print(f"  [dim]ZIP 생성: {zip_path.name} ({zip_path.stat().st_size / 1024 / 1024:.1f} MB)[/dim]")
    except Exception as e:
        console.print(f"[red]ZIP 생성 실패: {e}[/red]")
        return

    # 메일 본문
    success = sum(1 for _, _, _, c in results if c == 0)
    fail    = len(results) - success
    body_lines = [
        f"안녕하세요,",
        f"",
        f"[{timestamp}] 주식 분석 자동 실행이 완료되었습니다.",
        f"",
        f"■ 실행 결과",
    ]
    for i, (script, status, elapsed, _) in enumerate(results, 1):
        body_lines.append(f"  {i}. {script}  →  {status}  ({elapsed})")
    body_lines += [
        f"",
        f"■ 요약: 성공 {success}개 / 실패 {fail}개",
        f"",
        f"분석 결과 Excel 파일이 첨부되어 있습니다.",
        f"",
        f"※ 이 메일은 자동 발송입니다.",
    ]

    msg = MIMEMultipart()
    msg["From"]    = gmail_user
    msg["To"]      = ", ".join(_main) if _main else gmail_user
    if _extra:
        msg["Bcc"] = ", ".join(_extra)
    msg["Subject"] = f"[주식분석] {timestamp} 자동 실행 완료 (성공 {success}/{len(results)})"
    msg.attach(MIMEText("\n".join(body_lines), "plain", "utf-8"))

    # ZIP 첨부
    with open(zip_path, "rb") as f:
        part = MIMEApplication(f.read(), Name=zip_path.name)
        part["Content-Disposition"] = f'attachment; filename="{zip_path.name}"'
        msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(gmail_user, gmail_pw)
            smtp.sendmail(gmail_user, notify_emails, msg.as_string())
        _to_str  = ', '.join(_main) if _main else gmail_user
        _bcc_str = f" | Bcc: {len(_extra)}명" if _extra else ""
        console.print(f"  [green]이메일 발송 완료 — To: {_to_str}{_bcc_str}[/green]")
        if _extra:
            console.print(f"  [dim]  Bcc: {', '.join(_extra)}[/dim]")
    except Exception as e:
        console.print(f"  [red]이메일 발송 실패: {e}[/red]")


if __name__ == "__main__":
    main()
