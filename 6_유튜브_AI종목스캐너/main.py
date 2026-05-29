# =============================================================================
# 유튜브 채널 추천 종목 분석기
#
# [개요]
# 국내 주식 관련 유튜브 채널의 최신 영상을 자동으로 수집하고,
# 자막을 추출한 뒤 Gemini AI로 분석하여 추천 종목을 엑셀로 정리합니다.
#
# [동작 순서]
# 1. YouTube Data API로 각 채널의 20시간 이내 영상 조회 (채널당 최대 15개 조회)
# 2. YouTubeTranscriptApi로 영상 자막(한국어) 추출
# 3. Gemini AI에 자막을 전달하여 핵심 요약 + 추천 종목/사유를 JSON으로 분석
# 4. 4_financial_scanner/output/ 최신 xlsx에서 종목명→종목코드 매핑 조회
# 5. 결과를 DataFrame으로 정리 후 엑셀 파일로 저장 (하이퍼링크 포함)
#
# [분석 채널]
# - 달란트투자, 815머니톡, 삼프로TV, 김작가TV, 떠먹여주는TV, 경제야놀자,
#   딜사이트경제TV, ETF아는형, 815캠퍼스, 머니코믹스, 와이스트릿, 우리주식은안물어요
#
# [출력 파일]
# - 유튜브_인사이트_요약.xlsx (컬럼: 채널, 제목, 요약, 종목명, 종목코드, 추천사유, 링크)
#
# [필요 패키지]
# pip install google-api-python-client youtube-transcript-api google-generativeai python-dotenv openpyxl
#
# [환경 변수 (.env)]
# YOUTUBE_API_KEY=...   # Google Cloud Console에서 YouTube Data API v3 키 발급
# GEMINI_API_KEY=...    # Google AI Studio에서 Gemini API 키 발급
# =============================================================================

import os
import re
import sys
import glob
import json
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
import tempfile
import http.cookiejar

# Windows cp949 콘솔에서 한글/특수문자 깨짐 방지
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
from datetime import datetime, timezone, timedelta
import requests
import pandas as pd
from dotenv import load_dotenv
from googleapiclient.discovery import build
from youtube_transcript_api import YouTubeTranscriptApi
from google import genai
import openpyxl
from openpyxl.styles import Font
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

console = Console(legacy_windows=False)

load_dotenv()

YOUTUBE_API_KEY = os.getenv("YOUTUBE_API_KEY")
GEMINI_API_KEY  = os.getenv("GEMINI_API_KEY")

# 이메일 설정
GMAIL_USER    = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW  = os.getenv("GMAIL_APP_PASSWORD", "")
SEND_EMAIL    = os.getenv("SEND_EMAIL", "true").lower() in ("true", "1", "yes")
_extra_emails = [e.strip() for e in re.split(r"[,;]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
_main_emails  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
NOTIFY_EMAILS = _main_emails + _extra_emails

gemini = genai.Client(api_key=GEMINI_API_KEY)

# YouTube 자막 IP 차단 우회: 프로젝트 폴더의 cookies.txt 사용
_COOKIE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "www.youtube.com_cookies.txt")
def _make_yt_session() -> requests.Session:
    session = requests.Session()
    session.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    )
    if os.path.exists(_COOKIE_FILE):
        jar = http.cookiejar.MozillaCookieJar(_COOKIE_FILE)
        jar.load(ignore_discard=True, ignore_expires=True)
        session.cookies = jar
    return session

YT_SESSION = _make_yt_session()

CHANNELS = {
    "달란트투자":    "UCBM86JVoHLqg9irpR2XKvGw",
    "815머니톡":    "UCCG6BEYjfQMGzypJw2EJCDQ",
    "삼프로TV":     "UChlv4GSd7OQl3js-jkLOnFA",
    "김작가TV":     "UCvil4OAt-zShzkKHsg9EQAw",
    "떠먹여주는TV": "UC5dEgOV_mGqMHXizL1drtvA",
    "경제야놀자":   "UCgjsiInAgdceEq2lpRNNYvw",
    "딜사이트경제TV": "UCNi2OWpVGVBC2SuMBGXFzbA",
    "ETF아는형":   "UCIt_hQy30blHrVAS7QLTyTg",
    "815캠퍼스":   "UCnTZ9DHMo8oEoVIZrGlVoOA",
    "머니코믹스":        "UCJo6G1u0e_-wS-JQn3T-zEw",
    "와이스트릿":        "UCupslRq5jW95UGzPjOZz0FA",
    "우리주식은안물어요": "UC5kL3Ee6yY2bOtL2FRtdyuA",
}

# KRX 전 종목 기준 종목명 → 티커 맵 (FDR 우선, 4_financial_scanner 보조)
def _build_stock_map() -> tuple[dict, dict]:
    """
    반환: (raw_map, norm_map)
    raw_map  : 원본 종목명 → 티커
    norm_map : 소문자+공백제거 종목명 → (티커, 원본종목명)
    """
    raw: dict[str, str] = {}

    # 1순위: FDR StockListing (KRX 전 종목)
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

    # 2순위: 4_financial_scanner output (보조)
    try:
        base = os.path.join(os.path.dirname(__file__), "..", "4_financial_scanner", "output")
        files = sorted([f for f in glob.glob(os.path.join(base, "*.xlsx")) if "_hl" not in f])
        if files:
            df2 = pd.read_excel(files[-1], sheet_name=0, usecols=["종목명", "티커"], dtype=str)
            for n, t in zip(df2["종목명"].str.strip(), df2["티커"].str.strip()):
                if n not in raw:
                    raw[n] = t
    except Exception:
        pass

    # 정규화맵: 소문자+공백제거 → (티커, 원본명)
    norm: dict[str, tuple[str, str]] = {}
    for name, code in raw.items():
        key = name.replace(" ", "").lower()
        norm[key] = (code, name)

    return raw, norm


def _lookup_ticker(gemini_name: str, raw_map: dict, norm_map: dict) -> tuple[str, str]:
    """
    Gemini가 반환한 종목명으로 (티커, 정규종목명) 조회.
    못 찾으면 ('', gemini_name).
    전략: 원본 → 정규화 → 부분일치
    """
    if not gemini_name:
        return "", gemini_name

    # 1. 원본 그대로
    if gemini_name in raw_map:
        return raw_map[gemini_name], gemini_name

    # 2. 소문자+공백제거
    key = gemini_name.replace(" ", "").lower()
    if key in norm_map:
        code, canonical = norm_map[key]
        return code, canonical

    # 3. 부분 일치 (짧은 쪽이 긴 쪽에 포함되는지)
    best_code, best_name, best_len = "", gemini_name, 0
    for map_key, (code, canonical) in norm_map.items():
        if len(map_key) < 2:
            continue
        if key in map_key or map_key in key:
            if len(map_key) > best_len:
                best_code, best_name, best_len = code, canonical, len(map_key)
    return best_code, best_name


STOCK_RAW_MAP: dict = {}
STOCK_NORM_MAP: dict = {}


def _load_earnings_df() -> tuple[pd.DataFrame, pd.DataFrame]:
    """4_financial_scanner output 최신 파일의 '종합', '종합(분기포함)' 시트를 로드합니다."""
    base  = os.path.join(os.path.dirname(__file__), "..", "4_financial_scanner", "output")
    files = sorted([f for f in glob.glob(os.path.join(base, "*.xlsx"))
                    if "_hl" not in f])
    if not files:
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_brief = pd.read_excel(files[-1], sheet_name="종합", dtype=str)
        df_full  = pd.read_excel(files[-1], sheet_name="종합(분기포함)", dtype=str)
        return df_brief, df_full
    except Exception:
        return pd.DataFrame(), pd.DataFrame()



def get_latest_videos(channel_id, within_hours=20, max_fetch=15):
    """채널 업로드 플레이리스트로 최근 N시간 이내 영상 조회 (1 유닛, search 대비 100배 절약)"""
    try:
        youtube = build("youtube", "v3", developerKey=YOUTUBE_API_KEY)
        upload_playlist_id = "UU" + channel_id[2:]  # UC... → UU...
        response = youtube.playlistItems().list(
            playlistId=upload_playlist_id,
            part="snippet",
            maxResults=max_fetch,
        ).execute()
        cutoff = datetime.now(timezone.utc) - timedelta(hours=within_hours)
        videos = []
        for item in response.get("items", []):
            published_at = item["snippet"]["publishedAt"]
            pub_dt = datetime.fromisoformat(published_at.replace("Z", "+00:00"))
            if pub_dt >= cutoff:
                videos.append({
                    "video_id":    item["snippet"]["resourceId"]["videoId"],
                    "title":       item["snippet"]["title"],
                    "published_at": published_at,
                })
        return videos
    except Exception as e:
        console.print(f"[red][ERR] 영상 목록 조회 실패 ({channel_id}): {e}[/red]")
        return []


def get_transcript(video_id):
    """영상에서 자막을 추출합니다."""
    try:
        result = YouTubeTranscriptApi(http_client=YT_SESSION).fetch(video_id, languages=["ko"])
        return " ".join([t.text for t in result])
    except Exception as e:
        console.print(f"[yellow]  [WARN] 자막 없음 ({video_id}): {e}[/yellow]")
        return None


class _SilentLogger:
    """yt-dlp stderr 출력 완전 차단용 logger."""
    def debug(self, _): pass
    def warning(self, _): pass
    def error(self, _): pass


def analyze_from_audio(video_id, title):
    """자막 없는 영상 — 오디오 업로드 비용(영상당 ~45K 토큰)이 크므로 스킵."""
    console.print(f"  [dim]  [SKIP] 자막 없음, 오디오 분석 스킵: {title[:30]}[/dim]")
    return None
    import yt_dlp  # noqa: unreachable
    tmp_dir = tempfile.mkdtemp()
    audio_path = os.path.join(tmp_dir, f"{video_id}.m4a")
    try:
        # 오디오 다운로드 (ffmpeg 없이 m4a 직접 추출)
        ydl_opts = {
            "format": "bestaudio[ext=m4a]/bestaudio",
            "outtmpl": audio_path,
            "cookiefile": _COOKIE_FILE,
            "quiet": True,
            "no_warnings": True,
            "logger": _SilentLogger(),
            "extractor_args": {"youtube": {"skip": ["dash", "hls"]}},
        }
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([f"https://www.youtube.com/watch?v={video_id}"])

        # 실제 저장된 파일 찾기 (확장자가 다를 수 있음)
        saved = next(
            (os.path.join(tmp_dir, f) for f in os.listdir(tmp_dir) if f.startswith(video_id)),
            None,
        )
        if not saved:
            return None

        # Gemini Files API 업로드
        uploaded = gemini.files.upload(file=saved)

        prompt = f"""
아래는 유튜브 영상 '{title}'의 오디오입니다.
내용을 듣고 다음 JSON 형식으로만 응답해 주세요. 다른 텍스트는 포함하지 마세요.

{{
  "투자관련": true,
  "출연자": "홍길동, 김철수",
  "요약": "핵심 내용 3줄 이내 요약",
  "추천종목": [
    {{"종목명": "종목A", "사유": "추천 이유"}},
    {{"종목명": "종목B", "사유": "추천 이유"}}
  ]
}}

규칙:
- "투자관련": 주식·ETF·펀드·채권 등 금융투자와 관련된 내용이면 true, 관련 없으면 false
- "출연자": 영상에서 실제로 발표하거나 출연한 사람의 이름을 쉼표로 구분해 입력 (MC·진행자·게스트 등 실명으로 파악된 경우만, 없으면 빈 문자열)
- "추천종목"은 구체적으로 언급·추천한 국내 주식 종목만 포함하세요
"""
        for attempt in range(3):
            try:
                response = gemini.models.generate_content(
                    model="gemini-2.5-flash-lite",
                    contents=[uploaded, prompt],
                )
                break
            except Exception as e:
                err_str = str(e)
                if ("503" in err_str or "UNAVAILABLE" in err_str) and attempt < 2:
                    wait = 10 * (2 ** attempt)
                    console.print(f"[yellow]  >> 오디오 분석 일시 오류, {wait}초 대기 후 재시도...[/yellow]")
                    time.sleep(wait)
                else:
                    raise

        # 업로드 파일 정리
        gemini.files.delete(name=uploaded.name)

        text = response.text.strip()
        if text.startswith("```"):
            text = text.split("```")[1]
            if text.startswith("json"):
                text = text[4:]
        return json.loads(text)

    except Exception as e:
        console.print(f"[red]  [ERR] 오디오 분석 실패 ({video_id}): {e}[/red]")
        return None
    finally:
        # 임시 오디오 파일 정리
        for f in os.listdir(tmp_dir):
            os.remove(os.path.join(tmp_dir, f))
        os.rmdir(tmp_dir)


def _parse_gemini_response(text: str):
    text = text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text.strip())


def analyze_with_gemini(title, transcript, max_retries=3):
    """Gemini를 이용해 내용을 요약하고 종목을 JSON으로 추출합니다."""
    prompt = f"""
아래는 유튜브 영상 '{title}'의 자막 내용입니다.
이 내용을 바탕으로 다음 JSON 형식으로만 응답해 주세요. 다른 텍스트는 포함하지 마세요.

{{
  "투자관련": true,
  "출연자": "홍길동, 김철수",
  "요약": "핵심 내용 3줄 이내 요약",
  "추천종목": [
    {{"종목명": "종목A", "사유": "추천 이유"}},
    {{"종목명": "종목B", "사유": "추천 이유"}}
  ]
}}

규칙:
- "투자관련": 주식·ETF·펀드·채권 등 금융투자와 관련된 내용이면 true, 관련 없으면 false
- "출연자": 영상에서 실제로 발표하거나 출연한 사람의 이름을 쉼표로 구분해 입력 (MC·진행자·게스트 등 실명으로 파악된 경우만, 없으면 빈 문자열)
- "투자관련"이 false인 경우 "요약"과 "추천종목"은 빈 값으로 두세요
- "추천종목"은 영상에서 구체적으로 언급·추천한 국내 주식 종목만 포함하세요
- 종목을 특정하지 않은 일반적인 시장 전망이나 거시경제 내용은 추천종목을 빈 배열로 두세요

자막 내용:
{transcript[:8000]}
"""
    for attempt in range(max_retries):
        try:
            response = gemini.models.generate_content(model="gemini-2.5-flash-lite", contents=prompt)
            return _parse_gemini_response(response.text)
        except Exception as e:
            err_str = str(e)
            if "429" in err_str:
                import re
                m = re.search(r"retryDelay.*?(\d+)s", err_str)
                wait = int(m.group(1)) + 5 if m else 60
            elif "503" in err_str or "UNAVAILABLE" in err_str:
                wait = 10 * (2 ** attempt)   # 10s → 20s → 40s
            else:
                console.print(f"[red]  [ERR] Gemini 분석 실패: {e}[/red]")
                return None

            if attempt < max_retries - 1:
                console.print(f"[yellow]  >> Gemini 일시 오류, {wait}초 대기 후 재시도 ({attempt+1}/{max_retries})...[/yellow]")
                time.sleep(wait)
            else:
                console.print(f"[red]  [ERR] Gemini 재시도 {max_retries}회 실패: {e}[/red]")
                return None


KST = timezone(timedelta(hours=9))

def _to_kst(published_at: str) -> str:
    """YouTube publishedAt (UTC ISO 8601) → KST 문자열 'YYYY-MM-DD HH:MM'"""
    dt = datetime.fromisoformat(published_at.replace("Z", "+00:00"))
    return dt.astimezone(KST).strftime("%Y-%m-%d %H:%M")


def _format_excel(path):
    """저장된 엑셀 두 시트에 헤더 스타일·하이퍼링크·열 너비 적용."""
    from openpyxl.styles import Alignment, PatternFill

    HEADER_FILL_1 = PatternFill("solid", fgColor="1F4E79")  # 시트1: 진파랑
    HEADER_FILL_2 = PatternFill("solid", fgColor="375623")  # 시트2: 진초록
    HEADER_FONT   = Font(bold=True, color="FFFFFF")

    def _cw(val):
        return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

    def _apply_header(ws, fill):
        for cell in ws[1]:
            cell.fill      = fill
            cell.font      = HEADER_FONT
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

    def _autofit(ws):
        for col in ws.columns:
            w = max((_cw(cell.value) for cell in col), default=8) + 2
            ws.column_dimensions[col[0].column_letter].width = min(w, 60)
        ws.freeze_panes = "A2"

    wb = openpyxl.load_workbook(path)

    # ── 시트1: 종목추천 ──────────────────────────────────────
    ws1 = wb["종목추천"]
    _apply_header(ws1, HEADER_FILL_1)
    headers1 = [cell.value for cell in ws1[1]]
    name_col   = next((i+1 for i, h in enumerate(headers1) if h == "종목명"),   None)
    ticker_col = next((i+1 for i, h in enumerate(headers1) if h == "종목코드"), None)
    link_col1  = next((i+1 for i, h in enumerate(headers1) if h == "링크"),     None)

    for row in ws1.iter_rows(min_row=2, max_row=ws1.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        if ticker_col and name_col:
            tc = row[ticker_col - 1]
            nc = row[name_col   - 1]
            if tc.value and str(tc.value).strip():
                code6 = str(tc.value).strip().zfill(6)
                nc.hyperlink = f"https://finance.naver.com/item/main.naver?code={code6}"
                nc.font = Font(color="0563C1", underline="single", bold=True)
                tc.hyperlink = (
                    f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                    f"?pGB=1&gicode=A{code6}&cID=AA&MenuYn=Y&ReportGB="
                    f"&NewMenuID=11&stkGb=&strResearchYN="
                )
                tc.font = Font(color="0563C1", underline="single")
        if link_col1:
            lc = row[link_col1 - 1]
            if lc.value:
                lc.hyperlink = lc.value
                lc.font = Font(color="0563C1", underline="single")
                lc.value = "링크"
    _autofit(ws1)

    # ── 시트2: 투자관련 ──────────────────────────────────────
    ws2 = wb["투자관련"]
    _apply_header(ws2, HEADER_FILL_2)
    headers2  = [cell.value for cell in ws2[1]]
    link_col2 = next((i+1 for i, h in enumerate(headers2) if h == "링크"), None)

    for row in ws2.iter_rows(min_row=2, max_row=ws2.max_row):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        if link_col2:
            lc = row[link_col2 - 1]
            if lc.value:
                lc.hyperlink = lc.value
                lc.font = Font(color="0563C1", underline="single")
                lc.value = "링크"
    _autofit(ws2)

    wb.save(path)


def _send_email(output_file: str, results: list) -> None:
    """분석 결과 Excel을 첨부해 수신자 목록으로 이메일 전송."""
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from pathlib import Path

    ts      = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    df_res  = pd.DataFrame(results)
    n_total = len(df_res)
    n_stock = df_res["종목명"].replace("", pd.NA).dropna().nunique() if not df_res.empty else 0

    # 채널별 추천 종목 요약
    channel_lines = []
    if not df_res.empty:
        for ch, grp in df_res.groupby("채널", sort=False):
            stocks = [s for s in grp["종목명"].unique() if s]
            if stocks:
                channel_lines.append(f"  [{ch}] {', '.join(stocks)}")

    ch_list = " / ".join(CHANNELS.keys())

    body = "\n".join([
        f"{'='*60}",
        f"■ 주의사항 및 면책",
        f"{'='*60}",
        f"  · 본 자료는 투자 참고 목적으로만 제공되며,",
        f"    특정 종목에 대한 매수·매도 권유가 아닙니다.",
        f"  · 투자 판단 및 그에 따른 손익은 전적으로 본인에게 있습니다.",
        f"  · 본 자료는 수신자 개인에게만 제공된 것으로,",
        f"    SNS·커뮤니티·단체방 등 외부 재배포를 금합니다.",
        f"  · 무단 재배포로 인한 법적 책임은 재배포한 당사자에게 있습니다.",
        f"{'='*60}",
        f"",
        f"안녕하세요,",
        f"",
        f"[{ts}] 유튜브 AI 종목스캐너 분석이 완료되었습니다.",
        f"",
        f"{'='*60}",
        f"■ 분석 결과 요약",
        f"{'='*60}",
        f"  분석 영상 수 : {n_total}건",
        f"  추천 종목 수 : {n_stock}개 (중복 제외)",
        f"",
        f"■ 채널별 추천 종목",
        *(channel_lines if channel_lines else ["  (추천 종목 없음)"]),
        f"",
        f"{'='*60}",
        f"■ 첨부 엑셀 구성 (3개 시트)",
        f"{'='*60}",
        f"  [유튜브_Summary]",
        f"    종목을 추천한 영상 목록",
        f"    컬럼: 업로드일시 / 제목 / 채널 / 출연자 / 요약 / 종목명 / 종목코드 / 추천사유 / 링크",
        f"    · 종목명  → 네이버 금융 하이퍼링크",
        f"    · 종목코드 → FnGuide 실적 페이지 하이퍼링크",
        f"    · 동일 영상의 복수 종목은 A~D 컬럼 셀 병합",
        f"    · 티커 매칭 ✓ 연초록 / 미확인 연노랑",
        f"",
        f"  [투자관련]",
        f"    특정 종목 언급 없이 시장·매크로를 다룬 영상 목록",
        f"    컬럼: 업로드일시 / 제목 / 채널 / 요약 / 링크",
        f"",
        f"  [종합]",
        f"    추천 종목 중 실적 데이터(4_financial_scanner)와 매칭된 종목 — 분기 컬럼 제외",
        f"    컬럼: 종목명 / 티커 / 추천채널 / 업종 / 종합점수",
        f"          현재가 / 적정주가 / 상승여력 / 증권사목표주가",
        f"          매출·영업이익 증가율(25A→26E) / PER / 12M PER / PSR / PBR / PFR",
        f"    · 상승여력2 ≥ 20%  → 연초록",
        f"    · PER / 12M PER < 12 → 연파랑",
        f"    · 영업익증가율 ≥ 30% → 연노랑",
        f"",
        f"  [종합(분기포함)]",
        f"    종합 시트와 동일하되 분기 실적 컬럼 12개 추가",
        f"    최근 보고 분기 + 차기 추정 분기의 매출·영업이익 YoY 포함",
        f"",
        f"{'='*60}",
        f"■ 분석 채널 ({len(CHANNELS)}개)",
        f"{'='*60}",
        f"  {ch_list}",
        f"",
        f"■ 동작 방식",
        f"  YouTube Data API → 20시간 이내 업로드 영상만 수집 (채널당 최대 15개 조회)",
        f"  YouTubeTranscriptApi → 한국어 자막 추출",
        f"  자막 없는 영상 → yt-dlp 오디오 다운로드 → Gemini 전사",
        f"  Gemini AI (gemini-2.5-flash-lite) → 요약·추천종목 JSON 추출",
        f"  4_financial_scanner 최신 데이터와 티커 매칭 → 종합/종합(분기포함) 시트 생성",
        f"",
        f"상세 분석 결과는 첨부 Excel 파일을 확인해주세요.",
        f"",
        f"※ 이 메일은 자동 발송입니다.",
    ])

    try:
        (Path(output_file).parent / "summary.txt").write_text(body, encoding="utf-8")
    except Exception:
        pass

    if not SEND_EMAIL:
        console.print("[dim]이메일 발송 생략 (SEND_EMAIL=false)[/dim]")
        return
    if not (GMAIL_USER and GMAIL_APP_PW):
        console.print("[yellow]발송 계정 미설정 (.env: GMAIL_USER / GMAIL_APP_PASSWORD)[/yellow]")
        return
    if not NOTIFY_EMAILS:
        console.print("[yellow]수신자 미설정 (.env: NOTIFY_EMAIL / NOTIFY_EMAIL_WIFE)[/yellow]")
        return

    msg = MIMEMultipart()
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(_main_emails) if _main_emails else GMAIL_USER
    if _extra_emails:
        msg["Bcc"] = ", ".join(_extra_emails)
    msg["Subject"] = f"[유튜브 추천종목] {ts} 분석 완료 — {n_stock}개 종목"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    p = Path(output_file)
    if p.exists():
        with open(p, "rb") as f:
            part = MIMEApplication(f.read(), Name=p.name)
            part["Content-Disposition"] = f'attachment; filename="{p.name}"'
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(GMAIL_USER, GMAIL_APP_PW)
            smtp.sendmail(GMAIL_USER, NOTIFY_EMAILS, msg.as_string())
        _to_str  = ', '.join(_main_emails) if _main_emails else GMAIL_USER
        _bcc_str = f" | Bcc: {len(_extra_emails)}명" if _extra_emails else ""
        console.print(f"  [green]이메일 발송 완료 — To: {_to_str}{_bcc_str}[/green]")
        if _extra_emails:
            console.print(f"  [dim]  Bcc: {', '.join(_extra_emails)}[/dim]")
    except Exception as e:
        console.print(f"  [red]이메일 발송 실패: {e}[/red]")


# --- 메인 실행 ---
_stock_map_ex  = ThreadPoolExecutor(max_workers=1)
_stock_map_fut = _stock_map_ex.submit(_build_stock_map)
_start_time = datetime.now()
results = []          # 시트1: 종목 검출
related_results = []  # 시트2: 투자관련이지만 종목 없음
with Progress(
    SpinnerColumn(),
    TextColumn("{task.description}"),
    BarColumn(),
    MofNCompleteColumn(),
    TimeElapsedColumn(),
    TimeRemainingColumn(),
    console=console,
    transient=True,
) as progress:
    # 1단계: 영상 목록 수집
    t_collect = progress.add_task("영상 목록 수집 중...", total=len(CHANNELS))
    all_videos = []
    with ThreadPoolExecutor(max_workers=len(CHANNELS)) as ex:
        futs = {ex.submit(get_latest_videos, cid): name for name, cid in CHANNELS.items()}
        for fut in as_completed(futs):
            ch_name = futs[fut]
            for v in fut.result():
                all_videos.append((ch_name, v))
            progress.advance(t_collect)

    _raw, _norm = _stock_map_fut.result()
    STOCK_RAW_MAP.update(_raw)
    STOCK_NORM_MAP.update(_norm)
    _stock_map_ex.shutdown(wait=False)

    progress.update(t_collect, visible=False)  # 수집 완료 후 숨김

    # 2단계: 자막 추출 + Gemini 분석
    t_analyze = progress.add_task("분석 시작...", total=len(all_videos))
    for name, v in all_videos:
        progress.update(t_analyze, description=f"[{name}] {v['title'][:22]}")

        transcript = get_transcript(v["video_id"])
        time.sleep(2)  # YouTube 429 방지

        if transcript is not None:
            time.sleep(1)   # gemini-2.5-flash-lite 유료 → 여유 충분
            analysis = analyze_with_gemini(v["title"], transcript)
        else:
            # 자막 없는 영상 → 오디오 직접 분석
            progress.update(t_analyze, description=f"[{name}] 오디오 분석: {v['title'][:15]}")
            analysis = analyze_from_audio(v["video_id"], v["title"])
        if analysis is None:
            progress.advance(t_analyze)
            continue

        # 투자 무관 영상 건너뜀
        if not analysis.get("투자관련", True):
            progress.update(t_analyze, description=f"[dim][{name}] 투자무관 스킵[/dim]")
            progress.advance(t_analyze)
            continue

        link        = f"https://youtu.be/{v['video_id']}"
        summary     = analysis.get("요약", "")
        presenter   = analysis.get("출연자", "")
        uploaded_at = _to_kst(v["published_at"])
        stocks      = analysis.get("추천종목", [])

        if stocks:
            for stock in stocks:
                gemini_name = stock.get("종목명", "")
                stock_code, stock_name = _lookup_ticker(gemini_name, STOCK_RAW_MAP, STOCK_NORM_MAP)
                if not stock_name:
                    stock_name = gemini_name
                results.append({
                    "채널":      name,
                    "업로드일시": uploaded_at,
                    "제목":      v["title"],
                    "출연자":    presenter,
                    "요약":      summary,
                    "종목명":    stock_name,
                    "종목코드":  stock_code,
                    "추천사유":  stock.get("사유", ""),
                    "링크":      link,
                })
        else:
            # 투자관련이지만 종목 추천 없음 → 시트2
            related_results.append({
                "채널":      name,
                "업로드일시": uploaded_at,
                "제목":      v["title"],
                "출연자":    presenter,
                "요약":      summary,
                "링크":      link,
            })

        progress.advance(t_analyze)

output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
os.makedirs(output_dir, exist_ok=True)
timestamp   = datetime.now(KST).strftime("%Y_%m%d_%H%M")
output_file = os.path.join(output_dir, f"유튜브_AI종목스캐너_{timestamp}.xlsx")

# ── 실적정보 시트 준비 ────────────────────────────────────────────────────────
# 유튜브 추천 종목 코드 → 채널 목록 맵
_ticker_channels: dict[str, list[str]] = {}
for _r in results:
    _t = _r.get("종목코드", "").strip()
    _c = _r.get("채널", "")
    if _t:
        _ticker_channels.setdefault(_t, [])
        if _c not in _ticker_channels[_t]:
            _ticker_channels[_t].append(_c)

def _filter_matched(src: pd.DataFrame) -> pd.DataFrame:
    if src.empty or not _ticker_channels or "티커" not in src.columns:
        return pd.DataFrame()
    matched = src[src["티커"].isin(_ticker_channels)].copy()
    matched.insert(0, "추천채널", matched["티커"].map(
        lambda t: ", ".join(_ticker_channels.get(t, []))))
    return matched.reset_index(drop=True)

_earnings_brief_src, _earnings_full_src = _load_earnings_df()
df_earnings_brief = _filter_matched(_earnings_brief_src)
df_earnings_full  = _filter_matched(_earnings_full_src)

with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
    df_stocks = pd.DataFrame(results) if results else pd.DataFrame(
        columns=["채널","업로드일시","제목","출연자","요약","종목명","종목코드","추천사유","링크"])
    df_related = pd.DataFrame(related_results) if related_results else pd.DataFrame(
        columns=["채널","업로드일시","제목","출연자","요약","링크"])
    df_stocks.to_excel(writer, sheet_name="종목추천", index=False)
    df_related.to_excel(writer, sheet_name="투자관련", index=False)
    if not df_earnings_brief.empty:
        df_earnings_brief.to_excel(writer, sheet_name="종합", index=False)
    if not df_earnings_full.empty:
        df_earnings_full.to_excel(writer, sheet_name="종합(분기포함)", index=False)

# format_excel.py의 전체 포맷 적용 (종목추천 시트만 임시 파일로 처리 후 투자관련 시트 합산)
try:
    import importlib.util, sys as _sys

    _spec = importlib.util.spec_from_file_location(
        "format_excel",
        os.path.join(os.path.dirname(os.path.abspath(__file__)), "format_excel.py")
    )
    _mod = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_mod)

    # ① output_file에 이미 4개 시트(종목추천·투자관련·종합·종합(분기포함)) 모두 포함 → format_file 직접 적용
    _mod.format_file(output_file)
    fmt_path = os.path.splitext(output_file)[0] + "_fmt.xlsx"

    # ② 포맷 완료된 파일로 교체
    import shutil
    shutil.move(fmt_path, output_file)

    final_file = output_file
    output_file = final_file

except Exception as e:
    console.print(f"[yellow][WARN] 상세 포맷 실패: {e}[/yellow]")
    _format_excel(output_file)   # 폴백: 기본 서식만 적용

console.print(
    f"\n[bold green][OK] 분석 완료! "
    f"({len(CHANNELS)}개 채널에서 {len(all_videos)}개 동영상 분석)[/bold green]\n"
    f"{output_file}"
)

_send_email(output_file, results)

# 실행 시간 출력
_end_time = datetime.now()
_elapsed = _end_time - _start_time
_mins, _secs = divmod(int(_elapsed.total_seconds()), 60)
console.print(
    f"\n[bold green]⏱ 시작 {_start_time:%H:%M:%S} → 종료 {_end_time:%H:%M:%S}  "
    f"(총 {_mins}분 {_secs}초)[/bold green]"
)
