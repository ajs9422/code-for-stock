# -*- coding: utf-8 -*-
import os
import re
import sys
from datetime import datetime
from pathlib import Path

ROOT = Path("c:/works/code")

projects = [
    {"key": "4",  "name": "재무스캐너",           "dir": ROOT / "4_financial_scanner"},
    {"key": "6",  "name": "유튜브 AI 종목스캐너", "dir": ROOT / "6_유튜브_AI종목스캐너"},
    {"key": "8",  "name": "뉴스기사 인사이트",    "dir": ROOT / "8_뉴스기사_인사이트_LocalLLM"},
    {"key": "10", "name": "투자일정",             "dir": ROOT / "10_투자일정"},
    {"key": "14", "name": "마켓인사이트",          "dir": ROOT / "14_마켓인사이트"},
]


def read_summary(p: dict) -> str:
    f = p["dir"] / "output" / "summary.txt"
    return f.read_text(encoding="utf-8") if f.exists() else ""


_PROJECT_DESC: dict[str, str] = {
    "4":  "  실적 개선·저평가 종목 스크리닝 (매출·영업익 YoY + S-RIM 상승여력) → 종합/AND조건/메가테마/업종별 Excel",
    "6":  "  유튜브 투자 채널 AI 분석 (최근 영상 종목 추출 → 재무 데이터 매칭) → 추천 종목 Excel",
    "8":  "  뉴스 기사 AI 분석 (폭등후보·수혜·악재 분류 → 재무 데이터 매칭) → 주목 종목 Excel",
    "10": "  국내외 투자 이벤트 캘린더 (기업실적·글로벌 이슈·경제지표) → 주목 종목 매칭 Excel",
    "14": "  글로벌·국내 시황 AI 브리핑 (yfinance·pykrx + Gemini 분석) → PDF + Excel",
}

_BIG_COMPANIES: set[str] = {
    "애플", "마이크로소프트", "알파벳", "구글", "아마존", "메타", "엔비디아", "테슬라",
    "넷플릭스", "amd", "인텔", "tsmc", "브로드컴", "퀄컴", "jp모건", "골드만삭스",
    "버크셔", "비자", "마스터카드", "어도비", "세일즈포스", "오라클",
    "삼성전자", "sk하이닉스", "lg전자", "lg에너지솔루션", "현대차", "기아",
    "포스코", "sk텔레콤", "카카오", "네이버", "셀트리온", "삼성바이오로직스",
    "kb금융", "신한금융", "하나금융", "삼성sdi", "현대모비스", "삼성물산",
    "두산에너빌리티", "한화에어로스페이스", "sk이노베이션", "롯데케미칼",
}


def _is_big_co_event(line: str) -> bool:
    if any(kw in line for kw in ("[글로벌이벤트]", "[경제지표]", "[정책", "[금리")):
        return True
    low = line.lower()
    return any(co in low for co in _BIG_COMPANIES)


def _shorten_summary(summary: str, key: str) -> str:
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
    ranked_count = 0
    note_count   = 0
    event_count  = 0

    for ln in text.splitlines():
        s = ln.strip()
        if not s or s.startswith("====="):
            continue
        if any(kw in s for kw in stop_kw):
            break
        if any(kw in s for kw in skip_kw):
            continue

        if key == "4" and re.match(r"^\s+\d+\.", ln):
            ranked_count += 1
            if ranked_count > 5:
                continue

        if key == "8" and re.match(r"^\s+\[(폭등후보|직접수혜|간접수혜|악재)", ln):
            note_count += 1
            if note_count > 5:
                continue

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


results = []
for p in projects:
    summary = read_summary(p)
    results.append({
        "key": p["key"], "name": p["name"],
        "summary": summary, "success": bool(summary),
        "elapsed": "N분", "error": "",
    })

r14 = next(r for r in results if r["key"] == "14")
market_header = _parse_market_header(r14["summary"])

n_ok    = sum(1 for r in results if r["success"])
n_total = len(results)
from zoneinfo import ZoneInfo as _ZI
_KST  = _ZI("Asia/Seoul")
_now  = datetime.now(_KST)
_days = ["월", "화", "수", "목", "금", "토", "일"]
ts    = _now.strftime("%Y-%m-%d %H:%M")
date  = f"{_now.strftime('%Y-%m-%d')} ({_days[_now.weekday()]})"
sep60 = "─" * 60

sections: list[str] = []
for i, r in enumerate(results):
    ok_mark = "✓" if r["success"] else "✗"
    label = f"[{i+1}/{n_total}] {r['name']}  ({r['elapsed']})  {ok_mark}"
    block = _shorten_summary(r["summary"], r["key"])
    sections.append(f"{label}\n{block}")

body_lines: list[str] = []
if market_header:
    body_lines += market_header.splitlines()
    body_lines += ["", ""]
body_lines += [
    sep60,
    f"■ 일간 AI 투자 브리핑  {ts}  |  총 N분  |  성공 {n_ok}/{n_total}",
    sep60,
    "",
    *"\n\n".join(sections).splitlines(),
    "",
    sep60,
    "■ 첨부 파일 (5개)",
    "  1. 마켓인사이트_2026_0523_1734.pdf",
    "  2. financial_scanner_2026_0523_hl.xlsx",
    "  3. 유튜브_AI종목스캐너_2026_0523.xlsx",
    "  4. 뉴스기사_인사이트_2026_0523.xlsx",
    "  5. 투자일정_2026_0523.xlsx",
    "",
    sep60,
    "· 투자 참고 목적 자료 / 매수·매도 권유 아님 / 외부 재배포 금지",
    "· 투자 손익 책임은 본인에게 있습니다.",
    "※ 자동 발송",
]

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

body = "\n".join(body_lines)
print(body)
print()

# ── HTML 생성 ──────────────────────────────────────────────────────────────────

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

def _market_header_html(header_txt: str) -> str:
    if not header_txt:
        return ""
    return f"""
  <div style="background:#EBF5FB;border-left:4px solid #2E75B6;
              padding:16px 20px;border-radius:4px;margin-bottom:20px;
              font-size:13px;line-height:1.75;">
    {_summary_to_html(header_txt)}
  </div>"""

def _proj_section_html(r: dict) -> str:
    key   = r["key"]
    color = _PROJ_COLORS.get(key, "#333")
    icon  = _PROJ_ICONS.get(key, "📌")
    mark  = "✓" if r["success"] else "✗"
    body_content = (
        _summary_to_html(_shorten_summary(r["summary"], key))
        if r["success"]
        else f'<p style="color:#C0392B;">[실패] {_safe((r["error"] or "")[:300])}</p>'
    )
    return f"""
  <div style="margin:16px 0;">
    <div style="background:{color};color:white;padding:9px 16px;
                border-radius:4px 4px 0 0;font-size:13px;font-weight:bold;
                display:flex;justify-content:space-between;">
      <span>{icon} {r['name']}</span>
      <span style="color:{'#A9DFBF' if r['success'] else '#F1948A'};">{mark} {r['elapsed']}</span>
    </div>
    <div style="border:1px solid {color};border-top:none;padding:12px 16px;
                border-radius:0 0 4px 4px;background:#fafcff;
                font-size:12.5px;line-height:1.7;">
      {body_content}
    </div>
  </div>"""

proj_sections_html = "\n".join(
    _proj_section_html(r) for r in results if r["key"] != "14"
)

COMBINED_OUTPUT = Path(__file__).parent / "output"
COMBINED_OUTPUT.mkdir(exist_ok=True)

attach_files = sorted(COMBINED_OUTPUT.glob("*.xlsx")) + sorted(COMBINED_OUTPUT.glob("*.pdf"))
attach_items_html = "".join(
    f'<li style="margin:4px 0;"><b>{p.name}</b></li>'
    for p in attach_files
)

body_html = f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8"></head>
<body style="font-family:'Malgun Gothic',Arial,sans-serif;font-size:14px;
             color:#222;max-width:740px;margin:0 auto;padding:20px;">

  <!-- 헤더 -->
  <div style="background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);
              padding:22px 28px;border-radius:8px;margin-bottom:20px;">
    <div style="color:white;font-size:20px;font-weight:bold;letter-spacing:0.5px;">
      📈 홍익선생 JS &nbsp;|&nbsp; Daily Brief
    </div>
    <div style="color:#BDD7EE;font-size:13px;margin-top:6px;">
      {date} &nbsp;·&nbsp; {ts} &nbsp;·&nbsp;
      <span style="color:{'#A9DFBF' if n_ok == n_total else '#F1948A'};">
        성공 {n_ok}/{n_total}
      </span>
    </div>
  </div>

  <!-- 시황 브리핑 (14번) -->
  {_market_header_html(market_header)}

  <!-- 프로젝트별 결과 -->
  {proj_sections_html}

  <!-- 첨부 파일 -->
  <div style="background:#F4F6F7;border-left:4px solid #AAB7B8;padding:14px 18px;
              border-radius:4px;margin:20px 0;font-size:13px;">
    <strong>📎 output/ 파일 ({len(attach_files)}개)</strong>
    <ul style="margin:8px 0 0 0;padding-left:20px;line-height:1.8;">
      {attach_items_html if attach_items_html else '<li>없음</li>'}
    </ul>
  </div>

  <!-- 면책 -->
  <div style="background:#F8F8F8;border:1px solid #DDD;border-radius:4px;
              padding:11px 15px;font-size:12px;color:#777;margin-top:16px;">
    · 투자 참고 목적 자료 / 매수·매도 권유 아님 / 외부 재배포 금지<br>
    · 투자 손익 책임은 본인에게 있습니다.<br>
    ※ 자동 발송 ({ts})
  </div>

</body>
</html>"""

_date_str  = _now.strftime("%Y_%m%d")
_html_path = COMBINED_OUTPUT / f"brief_{_date_str}.html"
_html_path.write_text(body_html, encoding="utf-8")
print(f"✓ HTML 저장: output/brief_{_date_str}.html")
