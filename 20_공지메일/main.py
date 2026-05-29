# =============================================================================
# 공지 메일 발송기
#
# [개요]
# 지정된 수신자 목록으로 공지 이메일을 작성하고 발송합니다.
# HTML 본문 + 인라인 이미지 지원.
#
# [환경 변수 (.env)]
# GMAIL_USER=...             # 발송 Gmail 계정
# GMAIL_APP_PASSWORD=...     # Gmail 앱 비밀번호
# NOTIFY_EMAIL=...           # 주 수신자 이메일
# NOTIFY_EMAIL_WIFE=...      # 추가 수신자 이메일 (선택)
# NOTIFY_EMAILS_EXTRA=...    # 추가 수신자 목록 (쉼표 구분, 선택)
# SEND_EMAIL=true            # false 로 설정 시 발송 생략
# =============================================================================

import os
import re
import sys
import smtplib
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from datetime import datetime, timezone, timedelta
from pathlib import Path

from dotenv import load_dotenv
from rich.console import Console

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

console = Console(legacy_windows=False)

load_dotenv()

GMAIL_USER   = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW = os.getenv("GMAIL_APP_PASSWORD", "")
SEND_EMAIL   = os.getenv("SEND_EMAIL", "true").lower() in ("true", "1", "yes")

_extra_emails = [e.strip() for e in re.split(r"[,;\n]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
_main_emails  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
NOTIFY_EMAILS = _main_emails + _extra_emails

KST = timezone(timedelta(hours=9))


def send_notice(subject: str, body_text: str, body_html: str, inline_images: dict[str, str] | None = None) -> None:
    """공지 이메일 발송 (HTML + 인라인 이미지 지원).

    Args:
        subject: 이메일 제목
        body_text: 플레인 텍스트 본문 (HTML 미지원 클라이언트 대비)
        body_html: HTML 본문 (인라인 이미지는 <img src="cid:KEY"> 로 참조)
        inline_images: {"cid_key": "파일경로"} 형태의 인라인 이미지 맵
    """
    if not (GMAIL_USER and GMAIL_APP_PW):
        console.print("[yellow]발송 계정 미설정 (.env: GMAIL_USER / GMAIL_APP_PASSWORD)[/yellow]")
        return
    if not NOTIFY_EMAILS:
        console.print("[yellow]수신자 미설정 (.env: NOTIFY_EMAIL)[/yellow]")
        return

    # multipart/related: HTML + 인라인 이미지 묶음
    msg_related = MIMEMultipart("related")

    # multipart/alternative: 텍스트 / HTML 중 클라이언트가 선택
    msg_alt = MIMEMultipart("alternative")
    msg_alt.attach(MIMEText(body_text, "plain", "utf-8"))
    msg_alt.attach(MIMEText(body_html, "html", "utf-8"))
    msg_related.attach(msg_alt)

    # 인라인 이미지 첨부 (CID 방식)
    for cid_key, filepath in (inline_images or {}).items():
        p = Path(filepath)
        if p.exists():
            with open(p, "rb") as f:
                img = MIMEImage(f.read())
                img.add_header("Content-ID", f"<{cid_key}>")
                img.add_header("Content-Disposition", "inline", filename=p.name)
                msg_related.attach(img)
        else:
            console.print(f"[yellow][WARN] 이미지 없음: {filepath}[/yellow]")

    # 외부 봉투
    msg = MIMEMultipart("mixed")
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(_main_emails) if _main_emails else GMAIL_USER
    if _extra_emails:
        msg["Bcc"] = ", ".join(_extra_emails)
    msg["Subject"] = subject
    msg.attach(msg_related)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PW)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAILS, msg.as_string())
        _to_str  = ', '.join(_main_emails) if _main_emails else GMAIL_USER
        _bcc_str = f" | Bcc: {len(_extra_emails)}명" if _extra_emails else ""
        console.print(f"[green]✔ 이메일 발송 완료 — To: {_to_str}{_bcc_str}[/green]")
    except Exception as e:
        console.print(f"[red]이메일 발송 실패: {e}[/red]")


if __name__ == "__main__":
    ts = datetime.now(KST).strftime("%Y-%m-%d %H:%M")

    subject = "[공지] AI 투자 브리핑 서비스 안내 및 첨부 파일 활용 가이드"

    # ── 플레인 텍스트 ─────────────────────────────────────────────────────────
    body_text = "\n".join([
        "안녕하세요,",
        "",
        "매일 발송되는 AI 투자 브리핑 메일의 첨부 파일 활용 방법을 안내드립니다.",
        "",
        "=" * 60,
        "■ 서비스 업데이트 안내",
        "=" * 60,
        "",
        "기존에 뉴스·시황 분석에 사용하던 Google Gemini API의 비용이",
        "단기간에 급격히 증가하여 (5월 초 2주간 약 ₩30,800 발생),",
        "API 없이 자체 서버에서 구동되는 로컬 LLM(Local LLM, Ollama 기반)으로",
        "전면 교체하는 대규모 코드 업데이트를 완료하였습니다.",
        "",
        "  · 변경 전 : Google Gemini API (유료, 월 한도 초과 위험)",
        "  · 변경 후 : 로컬 LLM — Ollama 기반 qwen3:14b 모델 (API 비용 없음)",
        "",
        "분석 품질을 유지하면서 비용 부담 없이 서비스를 안정적으로 제공합니다.",
        "",
        "=" * 60,
        "■ 첨부 파일 구성 (총 5개)",
        "=" * 60,
        "",
        "① 마켓인사이트 PDF — 시황 브리핑",
        "",
        "  글로벌 주요 지표의 1년 일봉 차트를 한 페이지에 정리한 자료입니다.",
        "  S&P500 / NASDAQ / 다우존스 / 달러원 환율 / WTI 원유 / 미 10년 국채금리",
        "  리튬 (LIT ETF) / DRAM (Micron·MU) 가격 흐름까지 한눈에 확인하세요.",
        "  특히 리튬·DRAM 차트는 2차전지·반도체 사이클 파악에 유용합니다.",
        "  AI가 분석한 글로벌·국내 주요 이슈 및 주목 섹터·종목도 포함됩니다.",
        "",
        "② 재무스캐너 Excel — 실적 개선 · 저평가 종목 스크리닝",
        "",
        "  · PER / 12M PER(포워드 PER) 를 비교해 현재 밸류에이션을 확인하세요.",
        "    12M PER 가 낮을수록 향후 이익 대비 저평가 상태입니다.",
        "  · '분기포함' 탭에서 전년동기 대비 분기 매출·영업이익 상승률을 확인하세요.",
        "    실적 개선 추세가 이어지는 종목을 우선 검토합니다.",
        "  · 각 컬럼 헤더의 필터 버튼 → 내림차순 정렬로 상위 종목을 빠르게 추려보세요.",
        "",
        "③ 유튜브 AI 종목스캐너 Excel — 전문가 채널 분석",
        "",
        "  신뢰도 높은 투자 유튜브 채널 12개의 최신 영상을 AI가 분석합니다.",
        "  각 채널별 핵심 요약과 전문가들의 추천 종목·투자 조언을 정리해 드립니다.",
        "  여러 채널에서 중복 언급된 종목은 특히 주목할 필요가 있습니다.",
        "",
        "④ 뉴스기사 인사이트 Excel — 뉴스 기반 주목 종목",
        "",
        "  당일 주요 뉴스 기사를 AI가 분석해 종목별로 영향도를 분류합니다.",
        "  · 폭등후보 : 호재성 뉴스로 단기 급등 가능성이 있는 종목",
        "  · 직접수혜 : 해당 이슈의 직접적인 수혜 종목",
        "  · 간접수혜 : 연관 업종 또는 공급망 수혜 종목",
        "  · 악재     : 리스크 요인이 확인된 종목 (보유 중이라면 점검 필요)",
        "",
        "⑤ 투자일정 Excel — 국내외 주요 이벤트 캘린더",
        "",
        "  향후 일정별 주요 이벤트를 정리한 캘린더입니다.",
        "  기업 실적 발표 / 글로벌 경제지표 / 금리 결정 / 정책 이벤트 등을 포함합니다.",
        "  각 이벤트와 연관된 주목 종목도 함께 매칭되어 있습니다.",
        "  D-day 가 가까운 이벤트부터 확인하여 대응 전략을 미리 준비하세요.",
        "",
        "=" * 60,
        "■ 주의사항 및 면책",
        "=" * 60,
        "  · 본 서비스는 투자 참고 목적으로만 제공되며,",
        "    특정 종목에 대한 매수·매도 권유가 아닙니다.",
        "  · 투자 손익의 책임은 본인에게 있습니다.",
        "  · 본 메일은 수신자 개인에게만 제공된 것으로,",
        "    SNS·커뮤니티·단체방 등 외부 재배포를 금합니다.",
        "=" * 60,
        "",
        f"※ 이 메일은 자동 발송입니다. ({ts})",
    ])

    # ── HTML 본문 ─────────────────────────────────────────────────────────────
    def _section(color: str, icon: str, title: str, content: str) -> str:
        return f"""
  <div style="margin: 24px 0;">
    <div style="background:{color};color:white;padding:10px 16px;border-radius:6px 6px 0 0;
                font-size:14px;font-weight:bold;">
      {icon} {title}
    </div>
    <div style="border:1px solid {color};border-top:none;padding:14px 18px;
                border-radius:0 0 6px 6px;background:#fafcff;
                font-size:13px;line-height:1.85;">
      {content}
    </div>
  </div>"""

    # ── 섹션별 HTML 사전 계산 (f-string 내 """ 중첩 방지) ─────────────────────
    _s1 = _section("#1F4E79", "📄", "① 마켓인사이트 PDF — 시황 브리핑",
        "글로벌 주요 지표의 <strong>1년 일봉 차트</strong>를 한 페이지에 정리한 자료입니다."
        "<ul style='margin:8px 0 4px 0;padding-left:20px;'>"
        "<li>S&amp;P500 / NASDAQ / 다우존스 / 달러원 환율 / WTI 원유 / 미 10년 국채금리</li>"
        "<li><strong>리튬 (LIT ETF)</strong> / <strong>DRAM (Micron·MU)</strong> 가격 흐름 포함</li>"
        "</ul>"
        "리튬·DRAM 차트는 <strong>2차전지·반도체 사이클</strong> 파악에 특히 유용합니다.<br>"
        "AI가 분석한 글로벌·국내 주요 이슈, 주목 섹터·종목 브리핑도 함께 담겨 있습니다.")

    _s2 = _section("#375623", "📊", "② 재무스캐너 Excel — 실적 개선 · 저평가 종목",
        "<ul style='margin:6px 0;padding-left:20px;'>"
        "<li><strong>PER vs 12M PER(포워드 PER)</strong> 비교 — 12M PER가 낮을수록 향후 이익 대비 저평가 상태</li>"
        "<li><strong>'분기포함' 탭</strong> — 전년동기 대비 분기 매출·영업이익 상승률 확인,"
        " 실적 개선 추세가 이어지는 종목을 우선 검토</li>"
        "<li>컬럼 헤더 <strong>필터 버튼 → 내림차순 정렬</strong>로 상위 종목을 빠르게 추려볼 수 있습니다</li>"
        "</ul>")

    _s3 = _section("#843C0C", "📺", "③ 유튜브 AI 종목스캐너 Excel — 전문가 채널 분석",
        "신뢰도 높은 투자 유튜브 채널 <strong>12개</strong>의 최신 영상을 AI가 분석합니다."
        "<ul style='margin:8px 0 4px 0;padding-left:20px;'>"
        "<li>채널별 핵심 요약 및 전문가 추천 종목·투자 조언 정리</li>"
        "<li>여러 채널에서 <strong>중복 언급된 종목</strong>은 특히 주목할 필요가 있습니다</li>"
        "</ul>")

    _s4 = _section("#2E75B6", "📰", "④ 뉴스기사 인사이트 Excel — 뉴스 기반 주목 종목",
        "당일 주요 뉴스 기사를 AI가 분석해 종목별 영향도를 분류합니다."
        "<table style='border-collapse:collapse;width:100%;margin-top:10px;font-size:12.5px;'>"
        "<thead><tr style='background:#2E75B6;color:white;'>"
        "<th style='padding:6px 12px;border:1px solid #aaa;text-align:left;'>구분</th>"
        "<th style='padding:6px 12px;border:1px solid #aaa;text-align:left;'>의미</th>"
        "</tr></thead><tbody>"
        "<tr style='background:#f8f8f8;'><td style='padding:6px 12px;border:1px solid #ddd;font-weight:bold;color:#C00000;'>폭등후보</td><td style='padding:6px 12px;border:1px solid #ddd;'>호재성 뉴스로 단기 급등 가능성이 있는 종목</td></tr>"
        "<tr><td style='padding:6px 12px;border:1px solid #ddd;font-weight:bold;color:#375623;'>직접수혜</td><td style='padding:6px 12px;border:1px solid #ddd;'>해당 이슈의 직접적인 수혜 종목</td></tr>"
        "<tr style='background:#f8f8f8;'><td style='padding:6px 12px;border:1px solid #ddd;font-weight:bold;color:#70AD47;'>간접수혜</td><td style='padding:6px 12px;border:1px solid #ddd;'>연관 업종 또는 공급망 수혜 종목</td></tr>"
        "<tr><td style='padding:6px 12px;border:1px solid #ddd;font-weight:bold;color:#666;'>악재</td><td style='padding:6px 12px;border:1px solid #ddd;'>리스크 요인 확인 종목 (보유 중이라면 점검 필요)</td></tr>"
        "</tbody></table>")

    _s5 = _section("#4A235A", "📅", "⑤ 투자일정 Excel — 국내외 주요 이벤트 캘린더",
        "향후 일정별 주요 투자 이벤트를 정리한 캘린더입니다."
        "<ul style='margin:8px 0 4px 0;padding-left:20px;'>"
        "<li>기업 실적 발표 / 글로벌 경제지표 / 금리 결정 / 정책 이벤트 등 포함</li>"
        "<li>각 이벤트와 연관된 <strong>주목 종목</strong>도 함께 매칭</li>"
        "<li><strong>D-day 가 가까운 이벤트</strong>부터 확인해 대응 전략을 미리 준비하세요</li>"
        "</ul>")

    body_html = f"""<!DOCTYPE html>
<html lang="ko">
<head><meta charset="utf-8"></head>
<body style="font-family:'Malgun Gothic',Arial,sans-serif;font-size:14px;
             color:#222;max-width:720px;margin:0 auto;padding:20px;">

  <!-- 헤더 -->
  <div style="background:linear-gradient(135deg,#1F4E79 0%,#2E75B6 100%);
              padding:22px 28px;border-radius:8px;margin-bottom:24px;">
    <div style="color:white;font-size:20px;font-weight:bold;">
      📈 홍익선생 JS &nbsp;|&nbsp; AI 투자 브리핑 서비스 안내
    </div>
    <div style="color:#BDD7EE;font-size:13px;margin-top:6px;">
      첨부 파일 5종 구성 및 활용 가이드
    </div>
  </div>

  <p>안녕하세요,<br>
  매일 발송되는 <strong>AI 투자 브리핑 메일</strong>의 첨부 파일 활용 방법을 안내드립니다.</p>

  <!-- 서비스 업데이트 배경 -->
  <div style="background:#FEF9E7;border-left:4px solid #F39C12;
              padding:14px 18px;border-radius:4px;margin:20px 0;
              font-size:13px;line-height:1.8;">
    <p style="margin:0 0 8px 0;font-weight:bold;color:#B7770D;">🔧 서비스 업데이트 안내</p>
    기존에 사용하던 <strong>Google Gemini API</strong>의 비용이 단기간에 급격히 증가하여
    (5월 초 2주간 약 <strong style="color:#C0392B;">₩30,800</strong> 발생),
    API 없이 자체 서버에서 구동되는 <strong>로컬 LLM</strong>으로 전면 교체하는
    대규모 코드 업데이트를 완료하였습니다.
    <ul style="margin:10px 0 0 0;padding-left:20px;">
      <li>변경 전 : Google Gemini API (유료, 월 한도 초과 위험)</li>
      <li>변경 후 : <strong>로컬 LLM — Ollama 기반 qwen3:14b 모델</strong> (API 비용 없음)</li>
    </ul>
    <p style="margin:8px 0 0 0;">분석 품질을 유지하면서 비용 부담 없이 안정적으로 서비스를 제공합니다.</p>
  </div>

  {_s1}
  {_s2}
  {_s3}
  {_s4}
  {_s5}

  <!-- 면책 -->
  <div style="background:#F8F8F8;border:1px solid #DDD;border-radius:4px;
              padding:12px 16px;font-size:12px;color:#777;margin-top:24px;">
    · 본 서비스는 투자 참고 목적으로만 제공되며, 특정 종목에 대한 매수·매도 권유가 아닙니다.<br>
    · 투자 손익의 책임은 본인에게 있습니다.<br>
    · 본 메일은 수신자 개인에게만 제공된 것으로, SNS·커뮤니티·단체방 등 외부 재배포를 금합니다.<br>
    <br>※ 이 메일은 자동 발송입니다. ({ts})
  </div>

</body>
</html>"""

    # 발송 전 미리보기
    console.print(f"\n[bold cyan]{'='*60}[/bold cyan]")
    console.print(f"[bold]제목:[/bold] {subject}")
    console.print(f"[bold]수신:[/bold] {', '.join(_main_emails)}")
    console.print(f"[bold]Bcc:[/bold] {len(_extra_emails)}명")
    console.print(f"[bold cyan]{'='*60}[/bold cyan]\n")
    console.print(body_text)
    console.print(f"\n[bold cyan]{'='*60}[/bold cyan]\n")

    if SEND_EMAIL:
        send_notice(subject=subject, body_text=body_text, body_html=body_html)
    else:
        console.print("[dim]이메일 발송 생략 (SEND_EMAIL=false)[/dim]")
