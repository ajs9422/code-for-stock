# -*- coding: utf-8 -*-
"""테스트 이메일 발송 — 기존 summary.txt 데이터 사용, ajs9422@gmail.com 단독 수신."""
import sys
import os
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
os.environ["SEND_EMAIL"] = "true"

import run_all  # noqa: E402  (main()은 __main__ 가드로 실행 안 됨)

# ── 수신자를 테스트 주소 하나로 제한
run_all.NOTIFY_EMAILS = ["ajs9422@gmail.com"]
run_all.MAIN_EMAILS   = ["ajs9422@gmail.com"]
run_all.EXTRA_EMAILS  = []

# ── 각 프로젝트 summary.txt + 최신 excel/pdf 로드
results: list[dict] = []
for p in run_all.PROJECTS:
    out = p["dir"] / "output"
    summary_path = out / "summary.txt"
    summary = summary_path.read_text(encoding="utf-8") if summary_path.exists() else ""

    excels = sorted(out.glob("*.xlsx"), key=lambda f: f.stat().st_mtime, reverse=True) if out.exists() else []
    pdfs   = sorted(out.glob("*.pdf"),  key=lambda f: f.stat().st_mtime, reverse=True) if out.exists() else []

    results.append({
        "key":     p["key"],
        "name":    p["name"],
        "success": bool(summary),
        "excel":   excels[0] if excels else None,
        "pdf":     pdfs[0]   if pdfs   else None,
        "summary": summary,
        "elapsed": "[테스트]",
        "error":   "",
    })

run_all.console.print("[cyan]테스트 이메일 발송 → ajs9422@gmail.com[/cyan]")
run_all.send_combined_email(results, 0)
