# -*- coding: utf-8 -*-
"""5개 분석 스크립트 통합 실행 + 단일 이메일 발송.

실행 위치: c:/works/code/22_run_all/
    python run_all.py

병렬화 (2단계):
    Phase 1 병렬: 4_financial_scanner + 10_투자일정 + 14_마켓인사이트
    Phase 2 병렬: 6_유튜브   + 8_뉴스기사  ← 반드시 4번 완료 후

    이유: 6번·8번 모두 4_financial_scanner/output/ 최신 xlsx 에서 종목코드 매핑을 읽으므로
    4번 완료 전에 실행하면 파일 손상 오류가 발생합니다.
"""

import os
import re
import sys
import time
import shutil
import smtplib
import threading
import subprocess
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo
from concurrent.futures import ThreadPoolExecutor, as_completed
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from rich.console import Console
from rich.live import Live
from rich.panel import Panel

# ─── 경로 ────────────────────────────────────────────────────────────────────────
HERE            = Path(__file__).parent          # 22_run_all/
ROOT            = HERE.parent                    # c:\works\code
COMBINED_OUTPUT = HERE / "output"                # 22_run_all/output/

# ─── .env 로드 ───────────────────────────────────────────────────────────────────
try:
    from dotenv import load_dotenv
    load_dotenv(HERE / ".env", override=False)
except ImportError:
    _ef = HERE / ".env"
    if _ef.exists():
        for _ln in _ef.read_text(encoding="utf-8").splitlines():
            _ln = _ln.strip()
            if _ln and not _ln.startswith("#") and "=" in _ln:
                _k, _, _v = _ln.partition("=")
                os.environ.setdefault(_k.strip(), _v.strip())

# ─── 이메일 설정 ─────────────────────────────────────────────────────────────────
GMAIL_USER   = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW = os.getenv("GMAIL_APP_PASSWORD", "")
EXTRA_EMAILS: list[str] = [
    e.strip()
    for e in re.split(r"[,;]", os.getenv("NOTIFY_EMAILS_EXTRA", ""))
    if e.strip()
]
MAIN_EMAILS: list[str] = [
    e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")]
    if e.strip()
]
NOTIFY_EMAILS: list[str] = MAIN_EMAILS + EXTRA_EMAILS

# ─── 프로젝트 정의 ───────────────────────────────────────────────────────────────
PROJECTS: list[dict] = [
    {"key": "4",  "name": "재무스캐너",             "dir": ROOT / "4_financial_scanner",                "uses_ollama": False, "skip_if_today": True},
    {"key": "6",  "name": "유튜브 AI 종목스캐너",  "dir": ROOT / "6_유튜브_AI종목스캐너",        "uses_ollama": False, "skip_if_today": True},
    {"key": "8",  "name": "뉴스기사 인사이트",     "dir": ROOT / "8_1_뉴스기사_인사이트_LocalLLM", "uses_ollama": True,  "skip_if_today": False},
    {"key": "10", "name": "투자일정",              "dir": ROOT / "10_투자일정",                  "uses_ollama": False, "skip_if_today": True},
    {"key": "14", "name": "마켓인사이트",           "dir": ROOT / "14_마켓인사이트",              "uses_ollama": False, "skip_if_today": True},
]
_KEY_ORDER  = {p["key"]: i for i, p in enumerate(PROJECTS)}
_ollama_sem = threading.Semaphore(1)   # Ollama GPU: 동시 1개 제한

# ─── Rich 콘솔 ───────────────────────────────────────────────────────────────────
console = Console()
_SPIN   = ["⠋", "⠙", "⠹", "⠸", "⠼", "⠴", "⠦", "⠧", "⠇", "⠏"]


# ─── 진도율 추적기 ───────────────────────────────────────────────────────────────

class RunTracker:
    """스레드 안전한 실행 상태 추적 + Rich 패널 렌더링."""

    PHASE1 = ("4", "10", "14")   # 14번: 독립 실행 (4번 output 불필요)
    PHASE2 = ("6", "8")

    # 경험적 예상 소요 시간 (초) — ETA 계산용
    _EST_SEC: dict[str, int] = {
        "4":  480,   # ~8분  재무 크롤링 + S-RIM
        "6":  420,   # ~7분  YouTube API + Gemini
        "8":  1200,  # ~20분 뉴스 크롤링 + LocalLLM (qwen3:14b, think:False) + Gemini
        "10": 180,   # ~3분  투자일정
        "14": 480,   # ~8분  yfinance + pykrx + Gemini
    }

    def __init__(self, projects: list[dict]) -> None:
        self._lock  = threading.Lock()
        self._tick  = 0
        self._start = datetime.now()
        self._st: dict[str, dict] = {
            p["key"]: {"name": p["name"], "status": "waiting", "t0": None, "t1": None}
            for p in projects
        }
        self._last_lines: dict[str, list[str]] = {p["key"]: [] for p in projects}
        self._progress:   dict[str, str]       = {p["key"]: "" for p in projects}

    def set_running(self, key: str) -> None:
        with self._lock:
            self._st[key].update(status="running", t0=datetime.now())

    def set_done(self, key: str, success: bool, skipped: bool = False) -> None:
        with self._lock:
            status = "skipped" if skipped else ("done" if success else "failed")
            self._st[key].update(status=status, t1=datetime.now())

    def set_progress(self, key: str, text: str) -> None:
        with self._lock:
            self._progress[key] = text

    def set_last_line(self, key: str, line: str) -> None:
        # PROGRESS:X/Y 마커는 진행률 표시 전용 — last_lines에 넣지 않음
        m = re.match(r'^PROGRESS:(\d+)/(\d+)$', line)
        if m:
            self.set_progress(key, f"{m.group(1)}/{m.group(2)}")
            return
        with self._lock:
            buf = self._last_lines[key]
            buf.append(line)
            if len(buf) > 3:
                buf.pop(0)

    def n_done(self) -> int:
        with self._lock:
            return sum(1 for s in self._st.values() if s["status"] in ("done", "failed", "skipped"))

    # ── 렌더링 헬퍼 ─────────────────────────────────────────────────────────────

    def _elapsed(self, s: dict) -> str:
        if not s["t0"]:
            return ""
        sec = int(((s["t1"] or datetime.now()) - s["t0"]).total_seconds())
        return f"{sec // 60}분 {sec % 60:02d}초"

    def _total_elapsed(self) -> str:
        sec = int((datetime.now() - self._start).total_seconds())
        return f"{sec // 60}분 {sec % 60:02d}초"

    def _remaining(self, s: dict, key: str) -> str:
        """실행 중인 프로젝트의 예상 잔여 시간 문자열."""
        est = self._EST_SEC.get(key, 0)
        if not s["t0"] or s["status"] != "running" or not est:
            return ""
        elapsed_sec = (datetime.now() - s["t0"]).total_seconds()
        left = est - elapsed_sec
        if left < 0:
            over = int(-left)
            m, sec = divmod(over, 60)
            return f"예상 +{m}분 {sec:02d}초 초과"
        if left < 30:
            return "곧 완료"
        m, sec = divmod(int(left), 60)
        return f"~{m}분 {sec:02d}초 남음"

    def _total_eta(self, st: dict) -> str:
        """현재 phase 기준 전체 예상 잔여 시간."""
        ph1_done = all(st[k]["status"] in ("done", "failed") for k in self.PHASE1)

        if not ph1_done:
            # Phase1 중 가장 오래 걸릴 작업 기준 잔여 + Phase2 예상치
            ph1_left = 0
            for k in self.PHASE1:
                s = st[k]
                if s["status"] == "running" and s["t0"]:
                    elapsed_sec = (datetime.now() - s["t0"]).total_seconds()
                    ph1_left = max(ph1_left, max(0, self._EST_SEC.get(k, 0) - elapsed_sec))
                elif s["status"] == "waiting":
                    ph1_left = max(ph1_left, self._EST_SEC.get(k, 0))
            ph2_left = max((self._EST_SEC.get(k, 0) for k in self.PHASE2), default=0)
            total_left = ph1_left + ph2_left
        else:
            ph2_left = 0
            for k in self.PHASE2:
                s = st[k]
                if s["status"] == "running" and s["t0"]:
                    elapsed_sec = (datetime.now() - s["t0"]).total_seconds()
                    ph2_left = max(ph2_left, max(0, self._EST_SEC.get(k, 0) - elapsed_sec))
                elif s["status"] == "waiting":
                    ph2_left = max(ph2_left, self._EST_SEC.get(k, 0))
            total_left = ph2_left

        if total_left <= 0:
            return "완료 임박"
        m, sec = divmod(int(total_left), 60)
        return f"~{m}분 {sec:02d}초"

    # ── Rich Panel 렌더링 ────────────────────────────────────────────────────────

    def render(self) -> Panel:
        self._tick += 1
        spin = _SPIN[self._tick % len(_SPIN)]

        with self._lock:
            st           = {k: dict(v) for k, v in self._st.items()}
            last_lines   = {k: list(v) for k, v in self._last_lines.items()}
            progress_map = dict(self._progress)

        n_total = len(st)
        n_done  = sum(1 for s in st.values() if s["status"] in ("done", "failed"))

        lines: list[str] = []

        for ph, keys, desc in [
            ("Phase 1", self.PHASE1, "4_financial_scanner  +  10_투자일정  +  14_마켓인사이트"),
            ("Phase 2", self.PHASE2, "6_유튜브  +  8_뉴스기사  ← 4번 완료 후"),
        ]:
            lines.append(f"  [bold cyan]▌{ph}[/bold cyan]  [dim]{desc}[/dim]")
            lines.append("")

            for key in keys:
                s      = st[key]
                status = s["status"]
                el     = self._elapsed(s)

                if status == "waiting":
                    icon = "[dim]⏳[/dim]"
                    stxt = "[dim]대기중[/dim]"
                elif status == "running":
                    icon = f"[yellow]{spin}[/yellow]"
                    prog = progress_map.get(key, "")
                    prog_prefix = f"[cyan bold]{prog}[/cyan bold] " if prog else ""
                    stxt = f"{prog_prefix}[yellow bold]실행중[/yellow bold]"
                elif status == "done":
                    icon = "[green]✓[/green]"
                    stxt = "[green bold]완료[/green bold]  "
                elif status == "skipped":
                    icon = "[dim]↩[/dim]"
                    stxt = "[dim]스킵 (오늘 완료)[/dim]"
                else:
                    icon = "[red]✗[/red]"
                    stxt = "[red bold]실패[/red bold]  "

                el_str  = f"  [dim]{el}[/dim]" if el else ""
                rem     = self._remaining(s, key)
                rem_color = "red" if rem and "초과" in rem else "magenta"
                rem_str = f"  [{rem_color}]{rem}[/{rem_color}]" if rem else ""
                lines.append(f"   [{key:>2}]  {s['name']:<24}  {icon}  {stxt}{el_str}{rem_str}")
                if status == "running":
                    recent = [l for l in last_lines.get(key, []) if l][-2:]
                    for ln in recent:
                        lines.append(f"         [dim]└ {ln[:60]}[/dim]")
                elif status == "waiting" and ph == "Phase 2":
                    ph1_done = all(st[k]["status"] in ("done", "failed") for k in self.PHASE1)
                    if not ph1_done:
                        lines.append(f"         [dim]↑ Phase 1 완료 후 시작[/dim]")

            lines.append("")

        # 전체 진행률 바
        BAR_W  = 30
        filled = int(n_done / n_total * BAR_W)
        bar    = "[cyan]" + "█" * filled + "[/cyan]" + "[dim]" + "░" * (BAR_W - filled) + "[/dim]"
        pct    = int(n_done / n_total * 100)

        eta_str = self._total_eta(st)
        lines.append("  [dim]" + "─" * 58 + "[/dim]")
        lines.append(
            f"  전체 진행  {bar}  [bold]{pct:3d}%[/bold]"
            f"  {n_done}/{n_total} 완료  ·  경과 [bold]{self._total_elapsed()}[/bold]"
            f"  ·  잔여 [magenta bold]{eta_str}[/magenta bold]"
        )

        return Panel(
            "\n".join(lines),
            title="[bold cyan]AI 투자 브리핑 통합 실행기[/bold cyan]",
            border_style="cyan",
            padding=(0, 1),
        )


# ─── 유틸 ────────────────────────────────────────────────────────────────────────

def _strip_ansi(text: str) -> str:
    return re.sub(r"\x1b\[[0-9;]*[mGKHFJK]", "", text)

def _extract_summary(stdout: str, n_tail: int = 60) -> str:
    lines = [l for l in _strip_ansi(stdout).splitlines() if l.strip()]
    return "\n".join(lines[-n_tail:])

def _fmt_elapsed(total_seconds: float) -> str:
    m, s = divmod(int(total_seconds), 60)
    return f"{m}분 {s}초"


def _has_today_output(project_dir: Path) -> tuple[bool, "Path | None", "Path | None"]:
    """오늘 날짜(YYYY_MMDD) output 파일이 존재하면 (True, xlsx, pdf) 반환."""
    today      = datetime.now().strftime("%Y_%m%d")
    output_dir = project_dir / "output"
    if not output_dir.exists():
        return False, None, None

    xlsx = None
    for f in sorted(output_dir.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
        if today in f.name and not any(s in f.name for s in ("_hl", "_raw")):
            xlsx = f
            break
    if xlsx is None:
        for f in sorted(output_dir.glob("*.xlsx"), key=lambda x: x.stat().st_mtime, reverse=True):
            if today in f.name:
                xlsx = f
                break

    pdf = None
    for f in sorted(output_dir.glob("*.pdf"), key=lambda x: x.stat().st_mtime, reverse=True):
        if today in f.name:
            pdf = f
            break

    return (xlsx is not None or pdf is not None), xlsx, pdf


# ─── 프로젝트 실행 ───────────────────────────────────────────────────────────────

def run_project(p: dict, tracker: RunTracker) -> dict:
    key         = p["key"]
    project_dir = Path(p["dir"])
    output_dir  = project_dir / "output"
    output_dir.mkdir(exist_ok=True)

    # ── 오늘 결과 이미 존재하면 스킵 ─────────────────────────────────────────────
    if p.get("skip_if_today", False):
        found, xlsx, pdf = _has_today_output(project_dir)
        if found:
            tracker.set_running(key)
            console.log(f"[{key}] {p['name']} — 오늘 결과 이미 존재, 스킵")
            tracker.set_done(key, True, skipped=True)
            COMBINED_OUTPUT.mkdir(exist_ok=True)
            for f in (xlsx, pdf):
                if f and f.exists():
                    try:
                        shutil.copy2(f, COMBINED_OUTPUT / f.name)
                    except Exception:
                        pass
            summary_txt = output_dir / "summary.txt"
            summary = summary_txt.read_text(encoding="utf-8") if summary_txt.exists() else ""
            return {"key": key, "name": p["name"], "success": True,
                    "excel": xlsx, "pdf": pdf, "summary": summary,
                    "elapsed": "스킵", "error": ""}

    before:     set[Path] = set(output_dir.glob("*.xlsx"))
    before_pdf: set[Path] = set(output_dir.glob("*.pdf"))

    env = {
        **os.environ,
        "SEND_EMAIL":       "false",
        "PYTHONIOENCODING": "utf-8",
        "PYTHONUTF8":       "1",
        "NO_COLOR":         "1",
        "COLUMNS":          "120",
    }

    tracker.set_running(key)
    console.log(f"[{key}] {p['name']} 시작")
    t0 = datetime.now()

    acquire = p.get("uses_ollama", False)
    if acquire:
        console.log(f"[{key}] Ollama 세마포어 대기 중...")
        _ollama_sem.acquire()
        console.log(f"[{key}] Ollama 실행 시작")

    stdout_lines: list[str] = []
    stderr_buf:   list[str] = []

    try:
        proc = subprocess.Popen(
            [sys.executable, "main.py"],
            cwd=str(project_dir),
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="replace",
        )

        def _read_stdout() -> None:
            assert proc.stdout
            for raw in proc.stdout:
                ln = raw.rstrip("\n")
                stdout_lines.append(ln)
                clean = _strip_ansi(ln).strip()
                if clean:
                    tracker.set_last_line(key, clean)
            proc.stdout.close()

        def _read_stderr() -> None:
            assert proc.stderr
            for raw in proc.stderr:
                stderr_buf.append(raw)
            proc.stderr.close()

        t_out = threading.Thread(target=_read_stdout, daemon=True)
        t_err = threading.Thread(target=_read_stderr, daemon=True)
        t_out.start()
        t_err.start()

        try:
            proc.wait(timeout=7200)
        except subprocess.TimeoutExpired:
            proc.kill()
            t_out.join(timeout=5)
            t_err.join(timeout=5)
            raise

        t_out.join()
        t_err.join()

    except subprocess.TimeoutExpired:
        elapsed = (datetime.now() - t0).total_seconds()
        tracker.set_done(key, False)
        console.log(f"[{key}] [red]타임아웃 (2시간 초과)[/red]")
        return {"key": key, "name": p["name"], "success": False,
                "excel": None, "pdf": None, "summary": "", "elapsed": _fmt_elapsed(elapsed),
                "error": "타임아웃 (2시간 초과)"}
    except Exception as exc:
        elapsed = (datetime.now() - t0).total_seconds()
        tracker.set_done(key, False)
        console.log(f"[{key}] [red]예외: {exc}[/red]")
        return {"key": key, "name": p["name"], "success": False,
                "excel": None, "pdf": None, "summary": "", "elapsed": _fmt_elapsed(elapsed),
                "error": str(exc)}
    finally:
        if acquire:
            _ollama_sem.release()

    elapsed = (datetime.now() - t0).total_seconds()

    # 신규 Excel 탐지
    after     = set(output_dir.glob("*.xlsx"))
    new_files = after - before
    excel: Path | None = (
        max(new_files, key=lambda f: f.stat().st_mtime) if new_files else
        max(after,     key=lambda f: f.stat().st_mtime) if after else None
    )

    # 신규 PDF 탐지 (14번 마켓인사이트 등)
    after_pdf = set(output_dir.glob("*.pdf"))
    new_pdfs  = after_pdf - before_pdf
    pdf: Path | None = (
        max(new_pdfs,  key=lambda f: f.stat().st_mtime) if new_pdfs else
        max(after_pdf, key=lambda f: f.stat().st_mtime) if after_pdf else None
    )

    success = proc.returncode == 0 and excel is not None
    tracker.set_done(key, success)

    if success:
        console.log(f"[{key}] [green]완료[/green] — {_fmt_elapsed(elapsed)}")
    else:
        console.log(f"[{key}] [red]실패 (rc={proc.returncode})[/red] — {_fmt_elapsed(elapsed)}")

    # 22_run_all/output/ 에 복사
    COMBINED_OUTPUT.mkdir(exist_ok=True)
    if success and excel:
        try:
            dest = COMBINED_OUTPUT / excel.name
            shutil.copy2(excel, dest)
            console.log(f"[{key}] Excel 복사 → output/{excel.name}")
        except Exception as exc:
            console.log(f"[{key}] [yellow]Excel 복사 실패: {exc}[/yellow]")
    if pdf and pdf.exists():
        try:
            dest_pdf = COMBINED_OUTPUT / pdf.name
            shutil.copy2(pdf, dest_pdf)
            console.log(f"[{key}] PDF 복사 → output/{pdf.name}")
        except Exception as exc:
            console.log(f"[{key}] [yellow]PDF 복사 실패: {exc}[/yellow]")

    _summary_txt = output_dir / "summary.txt"
    _stdout_text = "\n".join(stdout_lines)
    _summary = (
        _summary_txt.read_text(encoding="utf-8")
        if _summary_txt.exists()
        else _extract_summary(_stdout_text)
    )

    return {
        "key":     key,
        "name":    p["name"],
        "success": success,
        "excel":   excel,
        "pdf":     pdf,
        "summary": _summary,
        "elapsed": _fmt_elapsed(elapsed),
        "error":   _strip_ansi("".join(stderr_buf))[-3000:] if not success else "",
    }


# ─── 이메일 요약 포맷터 ──────────────────────────────────────────────────────────

_PROJECT_DESC: dict[str, str] = {
    "4":  "  실적 개선·저평가 종목 스크리닝 (매출·영업익 YoY + S-RIM 상승여력) → 종합/AND조건/메가테마/업종별 Excel",
    "6":  "  유튜브 투자 채널 AI 분석 (최근 영상 종목 추출 → 재무 데이터 매칭) → 추천 종목 Excel",
    "8":  "  뉴스 기사 AI 분석 (폭등후보·수혜·악재 분류 → 재무 데이터 매칭) → 주목 종목 Excel",
    "10": "  국내외 투자 이벤트 캘린더 (기업실적·글로벌 이슈·경제지표) → 주목 종목 매칭 Excel",
    "14": "  글로벌·국내 시황 AI 브리핑 (yfinance·pykrx + Gemini 분석) → PDF + Excel",
}

_BIG_COMPANIES: set[str] = {
    # 미국 빅테크 / 주요 대형주
    "애플", "마이크로소프트", "알파벳", "구글", "아마존", "메타", "엔비디아", "테슬라",
    "넷플릭스", "amd", "인텔", "tsmc", "브로드컴", "퀄컴", "jp모건", "골드만삭스",
    "버크셔", "비자", "마스터카드", "어도비", "세일즈포스", "오라클",
    # 한국 대기업
    "삼성전자", "sk하이닉스", "lg전자", "lg에너지솔루션", "현대차", "기아",
    "포스코", "sk텔레콤", "카카오", "네이버", "셀트리온", "삼성바이오로직스",
    "kb금융", "신한금융", "하나금융", "삼성sdi", "현대모비스", "삼성물산",
    "두산에너빌리티", "한화에어로스페이스", "sk이노베이션", "롯데케미칼",
}


def _is_big_co_event(line: str) -> bool:
    """투자일정 이벤트 라인이 글로벌이벤트/경제지표이거나 대형주 관련이면 True."""
    if any(kw in line for kw in ("[글로벌이벤트]", "[경제지표]", "[정책", "[금리")):
        return True
    low = line.lower()
    return any(co in low for co in _BIG_COMPANIES)


def _shorten_summary(summary: str, key: str) -> str:
    """면책/인사말 제거 후 핵심 결과 라인만 반환. 키별 상이한 후처리 적용."""
    desc = _PROJECT_DESC.get(key, "")

    if key == "14":
        return f"{desc}\n  PDF 첨부 참조 / 상단 시황 브리핑 확인"

    start = -1
    for marker in ("■ 분석 결과 요약", "■ 분석결과"):
        idx = summary.find(marker)
        if idx >= 0:
            start = idx
            break
    text = summary[start:] if start >= 0 else summary

    skip_kw = ("주의사항", "면책", "매수·매도", "투자 판단", "재배포", "안녕하세요",
                "자동 발송", "동작 방식", "이 메일은", "상세 분석")
    stop_kw = ("■ 첨부 엑셀 구성", "■ 첨부파일", "컬럼 구성:", "시트 구성:",
                "[유튜브_Summary]", "[투자관련]", "[종합]")

    lines: list[str] = []
    ranked_count = 0   # 번호 매긴 항목(1/5/8/10번 종목 리스트) 카운터
    note_count   = 0   # 8번 주목 종목 ([폭등후보] 등) 카운터
    event_count  = 0   # 10번 이벤트 카운터

    for ln in text.splitlines():
        s = ln.strip()
        if not s or s.startswith("====="):
            continue
        if any(kw in s for kw in stop_kw):
            break
        if any(kw in s for kw in skip_kw):
            continue

        # ── 4번: 상위 종목 5개로 제한 ────────────────────────────────────────
        if key == "4" and re.match(r"^\s+\d+\.", ln):
            ranked_count += 1
            if ranked_count > 5:
                continue

        # ── 8번: 주목 종목 5개로 제한 ────────────────────────────────────────
        if key == "8" and re.match(r"^\s+\[(폭등후보|직접수혜|간접수혜|악재)", ln):
            note_count += 1
            if note_count > 5:
                continue

        # ── 10번: 대형주·글로벌 이벤트만 ────────────────────────────────────
        if key == "10" and re.match(r"^\s+\[D[+-]", ln):
            if not _is_big_co_event(ln):
                continue
            event_count += 1
            if event_count > 8:
                continue

        display = ln.rstrip()
        if len(display) > 100:
            display = display[:97] + "..."
        lines.append(display)

    result = "\n".join(lines) if lines else "  (결과 없음)"
    return f"{desc}\n{result}" if desc else result


def _parse_market_header(summary: str) -> str:
    """14_마켓인사이트 summary → 글로벌·한국 시황 + 섹터 + 이벤트 헤더 블록."""
    if not summary:
        return ""

    lines = summary.splitlines()
    parts: list[str] = []

    # ── 1. 글로벌·미국 주요 이슈 (Top 3: 제목 + 시장 영향) ─────────────────────
    global_issues: list[tuple[str, str]] = []
    section = ""
    cur_title = ""
    for ln in lines:
        s = ln.strip()
        if "■ 글로벌 주요 이슈" in s:
            section = "global"; cur_title = ""; continue
        if "■ 한국 주요 이슈" in s or "■ AI 종합 시황 분석" in s:
            break
        if section != "global":
            continue
        m = re.match(r"^\d+\.\s+\*\*(.+?)\*\*", s)
        if m:
            cur_title = m.group(1)
        elif cur_title and "시장 영향" in s:
            impact = re.sub(r"\*+\s*시장 영향\s*:\s*", "", s).strip()
            global_issues.append((cur_title, impact[:120]))
            cur_title = ""
            if len(global_issues) >= 3:
                break

    if global_issues:
        parts.append("▶ 글로벌·미국 시장")
        for title, impact in global_issues:
            parts.append(f"  • {title}")
            if impact:
                parts.append(f"    → {impact}")
        parts.append("")

    # ── 2. 한국 주요 이슈 (Top 3: 제목 + 내용 첫 문장) ─────────────────────────
    korea_issues: list[tuple[str, str]] = []
    section = ""
    cur_title = ""
    for ln in lines:
        s = ln.strip()
        if "■ 한국 주요 이슈" in s:
            section = "korea"; cur_title = ""; continue
        if "■ AI 종합 시황 분석" in s:
            break
        if section != "korea":
            continue
        m = re.match(r"^\d+\.\s+\*\*(.+?)\*\*", s)
        if m:
            cur_title = m.group(1)
        elif cur_title and re.match(r"\*+\s*내용\s*:", s):
            content_raw = re.sub(r"\*+\s*내용\s*:\s*", "", s).strip()
            first_sent  = content_raw.split(". ")[0][:130]
            korea_issues.append((cur_title, first_sent))
            cur_title = ""
            if len(korea_issues) >= 3:
                break

    if korea_issues:
        parts.append("▶ 한국 시장")
        for title, content in korea_issues:
            parts.append(f"  • {title}")
            if content:
                parts.append(f"    → {content}")
        parts.append("")

    # ── 3. 섹션B — 주목 섹터 Top3 + 대표 종목 ───────────────────────────────────
    sectors: list[str] = []
    in_b = False
    cur_name = ""
    for ln in lines:
        if "[섹션B]" in ln:
            in_b = True; continue
        if not in_b:
            continue
        if ln.strip().startswith("### [") and "[섹션B]" not in ln:
            break
        m = re.match(r"^\s*\d+\.\s+\*\*(.+?)\*\*", ln)
        if m:
            cur_name = m.group(1).rstrip(":")
        if cur_name and "관련 대표 종목" in ln:
            raw    = re.sub(r".*관련 대표 종목[:\s*]+", "", ln).replace("**", "").strip()
            stocks = [x.strip() for x in re.split(r"[,·]", raw)][:3]
            sectors.append(f"  {cur_name}  /  {', '.join(stocks)}")
            cur_name = ""
            if len(sectors) >= 3:
                break

    if sectors:
        parts.append("▶ 주목 섹터·테마")
        parts.extend(sectors)
        parts.append("")

    # ── 4. 섹션D — 핵심 이벤트 Top4 ─────────────────────────────────────────────
    events: list[str] = []
    in_d = False
    for ln in lines:
        if "[섹션D]" in ln:
            in_d = True; continue
        if not in_d:
            continue
        if ln.strip().startswith("### [") and "[섹션D]" not in ln:
            break
        m = re.match(r"^\d+\.\s+\*\*(.+?)\*\*", ln.strip())
        if m:
            events.append(f"  · {m.group(1).rstrip(':')}")
            if len(events) >= 4:
                break

    if events:
        parts.append("▶ 이번 주 핵심 이벤트")
        parts.extend(events)

    if not parts:
        return ""

    sep = "─" * 60
    return "\n".join([sep, "■ 오늘의 시황 브리핑", sep, ""] + parts)


# ─── 이메일 발송 ─────────────────────────────────────────────────────────────────

def send_combined_email(results: list[dict], total_elapsed: float) -> None:
    if os.getenv("SEND_EMAIL", "true").lower() not in ("true", "1", "yes"):
        console.log("[dim]이메일 발송 생략 (SEND_EMAIL=false)[/dim]")
        return
    if not (GMAIL_USER and GMAIL_APP_PW):
        console.log("[yellow]이메일 발송 계정 미설정 (GMAIL_USER / GMAIL_APP_PASSWORD)[/yellow]")
        return
    if not NOTIFY_EMAILS:
        console.log("[yellow]수신자 미설정 (NOTIFY_EMAIL / NOTIFY_EMAIL_WIFE)[/yellow]")
        return

    n_ok    = sum(1 for r in results if r["success"])
    n_total = len(results)
    _KST    = ZoneInfo("Asia/Seoul")
    _now    = datetime.now(_KST)
    _days   = ["월", "화", "수", "목", "금", "토", "일"]
    ts      = _now.strftime("%Y-%m-%d %H:%M")
    date    = f"{_now.strftime('%Y-%m-%d')} ({_days[_now.weekday()]})"
    sep60   = "─" * 60

    # 첨부: 14번 PDF만, 나머지 Excel
    attach_files: list[Path] = []
    r14 = next((r for r in results if r["key"] == "14"), None)
    if r14 and r14.get("pdf") and Path(r14["pdf"]).exists():
        attach_files.append(Path(r14["pdf"]))
    for r in results:
        if r["key"] == "14":
            continue
        if r["success"] and r["excel"] and Path(r["excel"]).exists():
            attach_files.append(Path(r["excel"]))

    # 시황 헤더 (14번 summary 활용)
    market_header = _parse_market_header(r14["summary"]) if r14 and r14.get("summary") else ""

    # ── 플레인 텍스트 (HTML 미지원 클라이언트 fallback) ──────────────────────────
    attachment_lines = [f"  {i+1}. {p.name}" for i, p in enumerate(attach_files)]
    sections_txt: list[str] = []
    for i, r in enumerate(results):
        ok_mark = "✓" if r["success"] else "✗"
        label = f"[{i+1}/{n_total}] {r['name']}  ({r['elapsed']})  {ok_mark}"
        block = _shorten_summary(r["summary"], r["key"]) if r["success"] else f"  [실패] {(r['error'] or '')[:300]}"
        sections_txt.append(f"{label}\n{block}")

    body_txt_lines: list[str] = []
    if market_header:
        body_txt_lines += market_header.splitlines() + ["", ""]
    body_txt_lines += [
        sep60,
        f"■ 일간 AI 투자 브리핑  {ts}  |  총 {_fmt_elapsed(total_elapsed)}  |  성공 {n_ok}/{n_total}",
        sep60, "",
        *"\n\n".join(sections_txt).splitlines(), "",
        sep60,
        f"■ 첨부 파일 ({len(attach_files)}개)",
        *(attachment_lines if attachment_lines else ["  (없음)"]), "",
        sep60,
        "· 투자 참고 목적 자료 / 매수·매도 권유 아님 / 외부 재배포 금지",
        "· 투자 손익 책임은 본인에게 있습니다.",
        "※ 자동 발송",
    ]
    body_text = "\n".join(body_txt_lines)

    # ── HTML 본문 ─────────────────────────────────────────────────────────────────
    _PROJ_COLORS: dict[str, str] = {
        "4":  "#375623",
        "6":  "#843C0C",
        "8":  "#2E75B6",
        "10": "#4A235A",
        "14": "#1F4E79",
    }
    _PROJ_ICONS: dict[str, str] = {
        "4": "📊", "6": "📺", "8": "📰", "10": "📅", "14": "🌐",
    }

    def _safe(s: str) -> str:
        return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    def _summary_to_html(text: str) -> str:
        parts = []
        for ln in text.splitlines():
            s = ln.strip()
            if not s:
                continue
            safe = _safe(s)
            if safe.startswith("■") or safe.startswith("▶") or safe.startswith("▌"):
                parts.append(f'<p style="font-weight:bold;color:#1F4E79;margin:8px 0 3px 0;">{safe}</p>')
            elif safe.startswith(("•", "·", "-", "→", "  •", "  ·")):
                parts.append(f'<p style="margin:2px 0 2px 12px;">{safe}</p>')
            elif re.match(r"^\d+\.", safe):
                parts.append(f'<p style="margin:4px 0 2px 0;"><b>{safe[:2]}</b>{safe[2:]}</p>')
            elif safe.startswith("[") and "]" in safe:
                parts.append(f'<p style="margin:3px 0 2px 8px;color:#555;">{safe}</p>')
            else:
                parts.append(f'<p style="margin:3px 0;color:#444;">{safe}</p>')
        return "\n".join(parts) if parts else '<p style="color:#999;">(결과 없음)</p>'

    # ── 테이블 공통 헬퍼 ──────────────────────────────────────────────────────────

    def _th(color: str, extra: str = "") -> str:
        return (
            f'background:{color};color:white;padding:7px 8px;text-align:left;'
            f'font-size:11.5px;font-weight:bold;{extra}'
        )

    def _td(even: bool, extra: str = "") -> str:
        bg = "#F0F4F8" if even else "#FAFCFF"
        return f'padding:6px 8px;border-bottom:1px solid #E2E8F0;background:{bg};vertical-align:top;{extra}'

    def _badge(text: str, bg: str, fg: str = "white") -> str:
        return (
            f'<span style="background:{bg};color:{fg};padding:2px 6px;'
            f'border-radius:10px;font-size:10.5px;font-weight:bold;'
            f'white-space:nowrap;">{_safe(text)}</span>'
        )

    def _tbl_wrap(inner: str) -> str:
        return (
            '<div style="overflow-x:auto;-webkit-overflow-scrolling:touch;">'
            f'<table style="width:100%;border-collapse:collapse;font-size:12px;margin:6px 0;">'
            f'{inner}</table></div>'
        )

    def _stat_bar(html: str) -> str:
        return f'<p style="margin:4px 0 8px 0;font-size:12px;color:#555;">{html}</p>'

    # ── Project 4: 재무스캐너 테이블 ─────────────────────────────────────────────

    def _proj4_table_html(summary: str) -> str:
        color = _PROJ_COLORS["4"]
        stat_parts: list[str] = []
        rows: list[tuple[str, str, str, str, str]] = []
        in_stocks = False
        for ln in summary.splitlines():
            s = ln.strip()
            if not s:
                continue
            if "■ 종합점수 상위" in s:
                in_stocks = True
                continue
            if s.startswith("■"):
                in_stocks = False
            if in_stocks:
                m = re.match(
                    r'(\d+)\.\s+(\S+)\s+점수:([\d.]+)\s+영업익증가율:([\d.,]+)%\s+상승여력:([\d.,]+)%', s
                )
                if m:
                    rows.append(m.groups())
                    if len(rows) >= 10:
                        break
            elif re.match(r'^(통과 종목|업종 수)', s):
                stat_parts.append(_safe(s))

        stats = _stat_bar(" &nbsp;|&nbsp; ".join(stat_parts)) if stat_parts else ""
        if not rows:
            return stats + _summary_to_html(summary)

        th_s = _th(color)
        thead = (
            f'<thead><tr>'
            f'<th style="{th_s}width:28px;text-align:center;">#</th>'
            f'<th style="{th_s}">종목명</th>'
            f'<th style="{th_s}text-align:center;">점수</th>'
            f'<th style="{th_s}text-align:right;">영업익증가율</th>'
            f'<th style="{th_s}text-align:right;">상승여력</th>'
            f'</tr></thead>'
        )
        tbody = "<tbody>"
        for i, (rank, name, score, op_gr, upside) in enumerate(rows):
            sf = float(score)
            sc = "#1A7A1A" if sf >= 80 else "#B8860B" if sf >= 60 else "#C0392B"
            td = _td(i % 2 == 0)
            tbody += (
                f'<tr>'
                f'<td style="{_td(i%2==0)}text-align:center;color:#888;">#{rank}</td>'
                f'<td style="{td}font-weight:bold;">{_safe(name)}</td>'
                f'<td style="{td}text-align:center;color:{sc};font-weight:bold;">{score}</td>'
                f'<td style="{td}text-align:right;color:#1565C0;">{op_gr}%</td>'
                f'<td style="{td}text-align:right;color:#2E7D32;">{upside}%</td>'
                f'</tr>'
            )
        tbody += "</tbody>"
        score_note = (
            '<p style="font-size:11px;color:#888;margin:8px 2px 2px 2px;line-height:1.7;">'
            '<b style="color:#555;">종합점수 산출(0~100)</b> : '
            '영업익증가율 20% &middot; 매출증가율 15% &middot; 차기Q영업익YoY 10% &middot; '
            'PER/12M PER 20% &middot; PEG 10% &middot; 상승여력1 10% &middot; 상승여력2 10% &middot; PSR 5%'
            '</p>'
        )
        return stats + _tbl_wrap(thead + tbody) + score_note

    # ── Project 6: 유튜브 — 언급빈도 Index + 채널별 테이블 ─────────────────────

    def _proj6_table_html(summary: str) -> str:
        color = _PROJ_COLORS["6"]
        stat_parts: list[str] = []
        channel_stocks: dict[str, list[str]] = {}
        in_ch = False
        for ln in summary.splitlines():
            s = ln.strip()
            if not s:
                continue
            if any(kw in s for kw in ("[유튜브_Summary]", "[투자관련]", "[종합]")):
                break
            if "■ 채널별 추천 종목" in s:
                in_ch = True
                continue
            if s.startswith("■") and in_ch:
                break
            if not in_ch:
                if re.match(r'^(분석 영상|추천 종목)', s):
                    stat_parts.append(_safe(s))
                continue
            m = re.match(r'^\[([^\]]+)\]\s+(.+)', s)
            if m:
                ch = m.group(1)
                stocks = [x.strip().rstrip('.') for x in re.split(r'[,、]', m.group(2)) if x.strip()]
                if stocks:
                    channel_stocks[ch] = stocks

        stats = _stat_bar(" &nbsp;|&nbsp; ".join(stat_parts)) if stat_parts else ""

        # 언급빈도 집계
        mention: dict[str, list[str]] = {}
        for ch, stks in channel_stocks.items():
            for stk in stks:
                mention.setdefault(stk, []).append(ch)
        ranked = sorted(mention.items(), key=lambda x: (-len(x[1]), x[0]))
        max_cnt = max((len(v) for v in mention.values()), default=1)

        th_s = _th(color)
        # 언급빈도 Index 테이블 (상위 15종목)
        idx_tbody = "<tbody>"
        for i, (stk, chs) in enumerate(ranked[:15]):
            cnt = len(chs)
            bar_w = max(4, int(cnt / max_cnt * 50))
            bar = (
                f'<span style="display:inline-block;width:{bar_w}px;height:8px;'
                f'background:{color};border-radius:2px;vertical-align:middle;'
                f'margin-right:4px;opacity:0.75;"></span>'
            )
            ch_str = _safe(", ".join(chs[:5]) + ("…" if len(chs) > 5 else ""))
            td = _td(i % 2 == 0)
            idx_tbody += (
                f'<tr>'
                f'<td style="{td}font-weight:bold;">{_safe(stk)}</td>'
                f'<td style="{td}text-align:center;white-space:nowrap;">'
                f'{bar}<b style="color:{color};">{cnt}</b></td>'
                f'<td style="{td}color:#666;font-size:11px;">{ch_str}</td>'
                f'</tr>'
            )
        idx_tbody += "</tbody>"
        idx_thead = (
            f'<thead><tr>'
            f'<th style="{th_s}">종목명</th>'
            f'<th style="{th_s}text-align:center;width:80px;">언급채널수</th>'
            f'<th style="{th_s}">언급 채널</th>'
            f'</tr></thead>'
        )
        idx_title = (
            f'<p style="font-weight:bold;color:{color};margin:8px 0 4px 0;font-size:12.5px;">'
            f'📊 종목 언급빈도 Index (채널수 기준)</p>'
        )
        idx_table = idx_title + _tbl_wrap(idx_thead + idx_tbody) if ranked else ""

        # 채널별 상세 테이블
        ch_tbody = "<tbody>"
        for i, (ch, stks) in enumerate(channel_stocks.items()):
            td = _td(i % 2 == 0)
            stks_str = _safe(", ".join(stks[:8]) + ("…" if len(stks) > 8 else ""))
            ch_tbody += (
                f'<tr>'
                f'<td style="{td}font-weight:bold;white-space:nowrap;">{_safe(ch)}</td>'
                f'<td style="{td}text-align:center;color:#666;width:30px;">{len(stks)}</td>'
                f'<td style="{td}font-size:11.5px;">{stks_str}</td>'
                f'</tr>'
            )
        ch_tbody += "</tbody>"
        ch_thead = (
            f'<thead><tr>'
            f'<th style="{th_s}white-space:nowrap;">채널명</th>'
            f'<th style="{th_s}text-align:center;">종목수</th>'
            f'<th style="{th_s}">추천 종목</th>'
            f'</tr></thead>'
        )
        ch_title = (
            f'<p style="font-weight:bold;color:{color};margin:14px 0 4px 0;font-size:12.5px;">'
            f'📺 채널별 추천 종목</p>'
        )
        ch_table = (ch_title + _tbl_wrap(ch_thead + ch_tbody)) if channel_stocks else ""

        return stats + idx_table + ch_table

    # ── Project 8: 뉴스기사 — 주목종목 + 악재 + 상충 강조 ───────────────────────

    def _proj8_table_html(summary: str) -> str:
        color = _PROJ_COLORS["8"]
        stat_parts: list[str] = []
        pos_stocks: list[tuple[str, str, str]] = []   # (type, name, reason)
        neg_stocks: list[tuple[str, str]] = []          # (name, reason)
        section = ""
        for ln in summary.splitlines():
            s = ln.strip()
            if not s:
                continue
            if "■ 분석 결과 요약" in s:
                section = "stats"; continue
            if "■ 주목 종목" in s:
                section = "pos"; continue
            if "■ 악재 경고 종목" in s:
                section = "neg"; continue
            if s.startswith("■"):
                section = ""; continue
            if section == "stats":
                if re.match(r'^(분석 기사|직접수혜|주목 종목|악재 경고)', s):
                    stat_parts.append(_safe(s))
            elif section == "pos":
                m = re.match(
                    r'^\[(폭등후보|직접수혜|간접수혜|공급망연관|기타수혜)\]\s+(.+?)\s+[—\-–]\s+(.+)', s
                )
                if m:
                    pos_stocks.append((m.group(1), m.group(2).strip(), m.group(3).strip()[:80]))
            elif section == "neg":
                m = re.match(r'^(.+?)\s+[—\-–]\s+(.+)', s)
                if m:
                    neg_stocks.append((m.group(1).strip(), m.group(2).strip()[:80]))

        stats = _stat_bar(" &nbsp;|&nbsp; ".join(stat_parts)) if stat_parts else ""

        # 상충 탐지
        pos_names = {name.lower() for _, name, _ in pos_stocks}
        neg_names = {name.lower() for name, _ in neg_stocks}
        conflict = pos_names & neg_names

        conflict_html = ""
        if conflict:
            names_str = _safe(", ".join(sorted(conflict)))
            conflict_html = (
                '<div style="background:#FEF9E7;border-left:4px solid #F39C12;'
                'padding:10px 14px;margin:8px 0;border-radius:4px;font-size:12px;">'
                f'<b style="color:#E67E22;">⚡ 상충 종목 주의</b> — 폭등후보·수혜와 악재가 동시에 언급: <b>{names_str}</b>'
                '<br><span style="color:#999;font-size:11px;">양면적 정보가 존재하므로 첨부 Excel에서 추가 확인 권장</span>'
                '</div>'
            )

        type_colors = {
            "폭등후보":   ("#C0392B", "white"),
            "직접수혜":   ("#1565C0", "white"),
            "간접수혜":   ("#2980B9", "white"),
            "공급망연관": ("#7D3C98", "white"),
            "기타수혜":   ("#666666", "white"),
        }
        th_s = _th(color)

        # 주목 종목 테이블
        pos_tbody = "<tbody>"
        for i, (stype, name, reason) in enumerate(pos_stocks):
            bg, fg = type_colors.get(stype, ("#666", "white"))
            badge = _badge(stype, bg, fg)
            is_conf = name.lower() in conflict
            conf_badge = (" " + _badge("⚡상충", "#F39C12")) if is_conf else ""
            name_style = "font-weight:bold;color:#E67E22;" if is_conf else "font-weight:bold;"
            td = _td(i % 2 == 0)
            pos_tbody += (
                f'<tr>'
                f'<td style="{td}white-space:nowrap;">{badge}{conf_badge}</td>'
                f'<td style="{td}{name_style}">{_safe(name)}</td>'
                f'<td style="{td}color:#555;font-size:11px;">{_safe(reason)}</td>'
                f'</tr>'
            )
        pos_tbody += "</tbody>"
        pos_thead = (
            f'<thead><tr>'
            f'<th style="{th_s}width:90px;">유형</th>'
            f'<th style="{th_s}width:100px;">종목명</th>'
            f'<th style="{th_s}">분석 요지</th>'
            f'</tr></thead>'
        )
        pos_title = '<p style="font-weight:bold;color:#2E75B6;margin:8px 0 4px 0;font-size:12.5px;">📌 주목 종목</p>'
        pos_table = (pos_title + _tbl_wrap(pos_thead + pos_tbody)) if pos_stocks else ""

        # 악재 테이블
        neg_th = _th("#C0392B")
        neg_tbody = "<tbody>"
        for i, (name, reason) in enumerate(neg_stocks):
            is_conf = name.lower() in conflict
            conf_badge = (" " + _badge("⚡상충", "#F39C12")) if is_conf else ""
            name_style = "color:#C0392B;font-weight:bold;" if is_conf else "color:#C0392B;"
            td = _td(i % 2 == 0)
            neg_tbody += (
                f'<tr>'
                f'<td style="{td}{name_style}">{_safe(name)}{conf_badge}</td>'
                f'<td style="{td}color:#555;font-size:11px;">{_safe(reason)}</td>'
                f'</tr>'
            )
        neg_tbody += "</tbody>"
        neg_thead = (
            f'<thead><tr>'
            f'<th style="{neg_th}width:130px;">종목명</th>'
            f'<th style="{neg_th}">악재 내용</th>'
            f'</tr></thead>'
        )
        neg_title = '<p style="font-weight:bold;color:#C0392B;margin:14px 0 4px 0;font-size:12.5px;">⚠️ 악재 경고 종목</p>'
        neg_table = (neg_title + _tbl_wrap(neg_thead + neg_tbody)) if neg_stocks else ""

        return stats + conflict_html + pos_table + neg_table

    # ── Project 10: 투자일정 — 이벤트 + 주목종목 테이블 ─────────────────────────

    def _proj10_table_html(summary: str) -> str:
        color = _PROJ_COLORS["10"]
        stat_parts: list[str] = []
        events: list[tuple[str, str]] = []
        stocks: list[tuple[str, str, str]] = []
        section = ""
        for ln in summary.splitlines():
            s = ln.strip()
            if not s:
                continue
            if "■ 분석 결과 요약" in s:
                section = "stats"; continue
            if "■ 임박 이벤트" in s:
                section = "events"; continue
            if "■ 주목 종목" in s:
                section = "stocks"; continue
            if s.startswith("■"):
                section = ""; continue
            if section == "stats":
                if re.match(r'^(수집 이벤트|주목 종목)', s):
                    stat_parts.append(_safe(s))
            elif section == "events":
                m = re.match(r'\[(D[+-]\d+)\]\s+(.+)', s)
                if m:
                    events.append((m.group(1), m.group(2)[:70]))
                    if len(events) >= 8:
                        section = ""
            elif section == "stocks":
                m = re.match(r'(.+?)\s+\(이벤트\s+(\d+)건\)\s+[—\-–]\s+(.+)', s)
                if m:
                    stocks.append((m.group(1).strip(), m.group(2), m.group(3)[:70]))
                    if len(stocks) >= 10:
                        section = ""

        stats = _stat_bar(" &nbsp;|&nbsp; ".join(stat_parts)) if stat_parts else ""
        th_s = _th(color)

        # 이벤트 테이블
        ev_tbody = "<tbody>"
        for i, (dday, desc) in enumerate(events):
            td = _td(i % 2 == 0)
            dc = "#C0392B" if dday.startswith("D-") or dday == "D+0" else "#1565C0"
            ev_tbody += (
                f'<tr>'
                f'<td style="{td}text-align:center;font-weight:bold;color:{dc};white-space:nowrap;width:55px;">{_safe(dday)}</td>'
                f'<td style="{td}">{_safe(desc)}</td>'
                f'</tr>'
            )
        ev_tbody += "</tbody>"
        ev_thead = (
            f'<thead><tr>'
            f'<th style="{th_s}text-align:center;">D-Day</th>'
            f'<th style="{th_s}">이벤트</th>'
            f'</tr></thead>'
        )
        ev_title = f'<p style="font-weight:bold;color:{color};margin:8px 0 4px 0;font-size:12.5px;">📅 임박 이벤트</p>'
        ev_table = (ev_title + _tbl_wrap(ev_thead + ev_tbody)) if events else ""

        # 주목 종목 테이블
        st_tbody = "<tbody>"
        for i, (name, cnt, reason) in enumerate(stocks):
            td = _td(i % 2 == 0)
            st_tbody += (
                f'<tr>'
                f'<td style="{td}font-weight:bold;">{_safe(name)}</td>'
                f'<td style="{td}text-align:center;color:{color};font-weight:bold;white-space:nowrap;">{cnt}건</td>'
                f'<td style="{td}color:#555;font-size:11px;">{_safe(reason)}</td>'
                f'</tr>'
            )
        st_tbody += "</tbody>"
        st_thead = (
            f'<thead><tr>'
            f'<th style="{th_s}">종목명</th>'
            f'<th style="{th_s}text-align:center;width:55px;">이벤트수</th>'
            f'<th style="{th_s}">전망 요지</th>'
            f'</tr></thead>'
        )
        st_title = f'<p style="font-weight:bold;color:{color};margin:14px 0 4px 0;font-size:12.5px;">🎯 주목 종목</p>'
        st_table = (st_title + _tbl_wrap(st_thead + st_tbody)) if stocks else ""

        return stats + ev_table + st_table

    # ── 섹션 렌더러 (테이블 분기) ─────────────────────────────────────────────────

    def _proj_section_html(r: dict) -> str:
        key   = r["key"]
        color = _PROJ_COLORS.get(key, "#333")
        icon  = _PROJ_ICONS.get(key, "📌")
        mark  = "✓" if r["success"] else "✗"
        mark_color = "#A9DFBF" if r["success"] else "#F1948A"

        if r["success"]:
            if key == "4":
                body = _proj4_table_html(r["summary"])
            elif key == "6":
                body = _proj6_table_html(r["summary"])
            elif key == "8":
                body = _proj8_table_html(r["summary"])
            elif key == "10":
                body = _proj10_table_html(r["summary"])
            else:
                body = _summary_to_html(_shorten_summary(r["summary"], key))
        else:
            body = f'<p style="color:#C0392B;">[실패] {_safe((r["error"] or "")[:300])}</p>'

        name_safe = _safe(r["name"])
        return (
            '<div style="margin:16px 0;">'
            f'<div style="background:{color};color:white;padding:9px 16px;'
            f'border-radius:4px 4px 0 0;font-size:13px;font-weight:bold;'
            f'display:flex;justify-content:space-between;">'
            f'<span>{icon} {name_safe}</span>'
            f'<span style="color:{mark_color};">{mark}</span>'
            '</div>'
            f'<div style="border:1px solid {color};border-top:none;padding:12px 16px;'
            f'border-radius:0 0 4px 4px;background:#fafcff;font-size:12.5px;line-height:1.7;">'
            f'{body}'
            '</div>'
            '</div>'
        )

    proj_sections_html = "\n".join(
        _proj_section_html(r) for r in results if r["key"] != "14"
    )

    attach_items_html = "".join(
        f'<li style="margin:4px 0;"><b>{p.name}</b></li>'
        for p in attach_files
    )

    _ok_color   = "#A9DFBF" if n_ok == n_total else "#F1948A"
    _attach_li  = attach_items_html if attach_items_html else "<li>없음</li>"

    body_html = (
        '<!DOCTYPE html>\n<html lang="ko">\n<head>\n'
        '<meta charset="utf-8">\n'
        '<meta name="viewport" content="width=device-width,initial-scale=1.0">\n'
        '<style>\n'
        '  body{font-family:"Malgun Gothic",Arial,sans-serif;font-size:14px;'
        'color:#222;max-width:740px;width:100%;margin:0 auto;'
        'padding:20px;box-sizing:border-box;}\n'
        '  @media screen and (max-width:600px){\n'
        '    body{padding:10px !important;font-size:13px !important;}\n'
        '    .eh{padding:14px 12px !important;}\n'
        '    .eh-t{font-size:17px !important;}\n'
        '    .sh{padding:9px 10px !important;}\n'
        '    .sb{padding:8px 10px !important;font-size:12px !important;}\n'
        '    .mh{padding:12px 10px !important;}\n'
        '    table{font-size:11px !important;}\n'
        '    th,td{padding:5px 4px !important;}\n'
        '  }\n'
        '</style>\n</head>\n'
        '<body>\n\n'
        '  <!-- 헤더 -->\n'
        '  <div class="eh" style="background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);'
        'padding:22px 28px;border-radius:8px;margin-bottom:20px;">\n'
        f'    <div class="eh-t" style="color:white;font-size:20px;font-weight:bold;letter-spacing:0.5px;">\n'
        f'      \U0001f4c8 홍익선생 JS &nbsp;|&nbsp; Daily Brief\n'
        f'    </div>\n'
        f'    <div style="color:#BDD7EE;font-size:13px;margin-top:6px;">\n'
        f'      {date} &nbsp;\xb7&nbsp; {ts} &nbsp;\xb7&nbsp;\n'
        f'      <span style="color:{_ok_color};">성공 {n_ok}/{n_total}</span>\n'
        f'    </div>\n'
        f'  </div>\n\n'
        f'  <!-- 시황 브리핑 (14번) -->\n'
        + (
            '<div class="mh" style="background:#EBF5FB;border-left:4px solid #2E75B6;'
            'padding:16px 20px;border-radius:4px;margin-bottom:20px;font-size:13px;line-height:1.75;">'
            + _summary_to_html(market_header)
            + '</div>'
            if market_header else ""
        )
        + f'\n\n  <!-- 프로젝트별 결과 -->\n  {proj_sections_html}\n\n'
        '  <!-- 첨부 파일 -->\n'
        '  <div style="background:#F4F6F7;border-left:4px solid #AAB7B8;padding:14px 18px;'
        'border-radius:4px;margin:20px 0;font-size:13px;">\n'
        f'    <strong>\U0001f4ce 첨부 파일 ({len(attach_files)}개)</strong>\n'
        '    <ul style="margin:8px 0 0 0;padding-left:20px;line-height:1.8;">\n'
        f'      {_attach_li}\n'
        '    </ul>\n'
        '  </div>\n\n'
        '  <!-- 면책 -->\n'
        '  <div style="background:#F8F8F8;border:1px solid #DDD;border-radius:4px;'
        'padding:11px 15px;font-size:12px;color:#777;margin-top:16px;">\n'
        '    \xb7 투자 참고 목적 자료 / 매수\xb7매도 권유 아님 / 외부 재배포 금지<br>\n'
        '    \xb7 투자 손익 책임은 본인에게 있습니다.<br>\n'
        f'    ※ 자동 발송 ({ts})\n'
        '  </div>\n\n'
        '</body>\n</html>'
    )

    # ── HTML 저장 ─────────────────────────────────────────────────────────────────
    _date_str  = _now.strftime("%Y_%m%d")
    _html_path = COMBINED_OUTPUT / f"brief_{_date_str}.html"
    COMBINED_OUTPUT.mkdir(exist_ok=True)
    _html_path.write_text(body_html, encoding="utf-8")
    console.log(f"[green]HTML 저장: output/brief_{_date_str}.html[/green]")

    # ── 발송 ──────────────────────────────────────────────────────────────────────
    msg = MIMEMultipart("mixed")
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(MAIN_EMAILS) if MAIN_EMAILS else GMAIL_USER
    if EXTRA_EMAILS:
        msg["Bcc"] = ", ".join(EXTRA_EMAILS)
    _short_date = f"{_now.strftime('%m.%d')} ({_days[_now.weekday()]})"
    msg["Subject"] = f"홍익선생 JS | 📊 Daily Brief · {_short_date}"

    alt = MIMEMultipart("alternative")
    alt.attach(MIMEText(body_text, "plain", "utf-8"))
    alt.attach(MIMEText(body_html,  "html",  "utf-8"))
    msg.attach(alt)

    for fp in attach_files:
        with open(fp, "rb") as f:
            part = MIMEApplication(f.read(), Name=fp.name)
            part["Content-Disposition"] = f'attachment; filename="{fp.name}"'
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PW)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAILS, msg.as_string())
        to_str  = ", ".join(MAIN_EMAILS) if MAIN_EMAILS else GMAIL_USER
        bcc_str = f" | Bcc: {len(EXTRA_EMAILS)}명" if EXTRA_EMAILS else ""
        console.log(f"[green]이메일 발송 완료 - To: {to_str}{bcc_str}[/green]")
    except Exception as exc:
        console.log(f"[red]이메일 발송 실패: {exc}[/red]")


# ─── 메인 ────────────────────────────────────────────────────────────────────────

def main() -> None:
    overall_start = datetime.now()
    COMBINED_OUTPUT.mkdir(exist_ok=True)

    tracker = RunTracker(PROJECTS)
    phase1  = [p for p in PROJECTS if p["key"] in RunTracker.PHASE1]
    phase2  = [p for p in PROJECTS if p["key"] in RunTracker.PHASE2]
    results: list[dict] = []

    stop_refresh = threading.Event()

    def _refresh_loop(live: Live) -> None:
        while not stop_refresh.is_set():
            live.update(tracker.render(), refresh=True)
            time.sleep(0.4)
        live.update(tracker.render(), refresh=True)   # 최종 상태 반영

    with Live(tracker.render(), console=console, auto_refresh=False) as live:
        refresh_thread = threading.Thread(target=_refresh_loop, args=(live,), daemon=True)
        refresh_thread.start()

        # ── Phase 1: 4번 + 10번 + 14번 ──────────────────────────────────────────
        console.log(f"[cyan]Phase 1 시작:[/cyan] {' + '.join(p['name'] for p in phase1)}")
        with ThreadPoolExecutor(max_workers=len(phase1)) as ex:
            futs = {ex.submit(run_project, p, tracker): p for p in phase1}
            for f in as_completed(futs):
                results.append(f.result())

        r4 = next((r for r in results if r["key"] == "4"), None)
        if r4 and not r4["success"]:
            console.log("[yellow][경고] 4_financial_scanner 실패 — 6번·8번은 이전 output xlsx 참조[/yellow]")

        # ── Phase 2: 6번 + 8번 ──────────────────────────────────────────────────
        console.log(f"[cyan]Phase 2 시작:[/cyan] {' + '.join(p['name'] for p in phase2)}")
        with ThreadPoolExecutor(max_workers=len(phase2)) as ex:
            futs = {ex.submit(run_project, p, tracker): p for p in phase2}
            for f in as_completed(futs):
                results.append(f.result())

        stop_refresh.set()
        refresh_thread.join(timeout=1.0)

    # ── 결과 출력 ──────────────────────────────────────────────────────────────
    results.sort(key=lambda r: _KEY_ORDER[r["key"]])
    total_elapsed = (datetime.now() - overall_start).total_seconds()
    n_ok = sum(1 for r in results if r["success"])

    console.print()
    console.rule("[bold cyan]실행 결과[/bold cyan]")
    for r in results:
        mark = "[green]✓[/green]" if r["success"] else "[red]✗[/red]"
        console.print(f"  {mark}  [{r['key']:>2}]  {r['name']:<24}  {r['elapsed']}")
    console.print(f"\n  성공 {n_ok}/{len(results)}  ·  총 {_fmt_elapsed(total_elapsed)}")
    console.rule()

    if n_ok == 0:
        console.print("[red]전체 실패 — 이메일 발송 생략[/red]")
        for r in results:
            if r["error"]:
                console.print(f"  [{r['key']}] {r['error'][:300]}")
        return

    send_combined_email(results, total_elapsed)


if __name__ == "__main__":
    main()
