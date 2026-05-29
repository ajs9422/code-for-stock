# =============================================================================
# 투자일정 수집기 — 2달 이내 주요 투자 이벤트 자동 수집 · AI 관련주 매핑 · 엑셀 정리
#
# [개요]
# DART 공시 API · 네이버 금융 증시일정 · Gemini AI 지식 보완으로
# 향후 60일 이내 주요 투자 이벤트를 수집하고,
# Gemini AI로 이벤트별 관련주 매핑 + 종목별 투자의견을 엑셀 2개 시트로 정리합니다.
#
# [동작 순서]
# 1. 이벤트 수집 (3개 소스 병렬)
#    - DART 공시 API       : 최근 30일 주요사항보고 (주주총회·증자·배당 등)
#    - 네이버 금융 증시일정 : 배당락·권리락·코스피 이벤트 스크래핑
#    - Gemini AI 지식 보완 : FOMC·WWDC·MSCI 등 글로벌 반복 이벤트
# 2. 정규화 / 중복 제거 / D-Day 계산 (오늘 ~ +60일 필터)
# 3. Gemini AI 개별 분석 (이벤트별 관련주·투자포인트·영향도)
# 4. 종목 언급 집계 → Gemini AI 종합 투자의견 (Top 20 종목)
# 5. 엑셀 저장 (Sheet1 투자일정 / Sheet2 종목인사이트) + 서식 적용
# 6. 이메일 발송
#
# [출력 파일]
# - output/투자일정_YYYY_MMDD_HHmm.xlsx
#
# [필요 패키지]
# pip install requests beautifulsoup4 google-genai python-dotenv pandas openpyxl rich
#
# [환경 변수 (.env)]
# GEMINI_API_KEY=...
# DART_API_KEY=...       (없으면 DART 수집 스킵)
# GMAIL_USER=...
# GMAIL_APP_PASSWORD=...
# NOTIFY_EMAIL=...
# NOTIFY_EMAIL_WIFE=...
# NOTIFY_EMAILS_EXTRA=...
# SEND_EMAIL=true
# =============================================================================

import argparse
import glob
import importlib.util
import json
import os
import random
import re
import shutil
import sys
import threading
import time
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta, timezone

import pandas as pd
import requests
from bs4 import BeautifulSoup
from dotenv import find_dotenv, load_dotenv
from google import genai
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
from rich.table import Column

if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

console = Console(legacy_windows=False)

load_dotenv(find_dotenv())

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
if not GEMINI_API_KEY:
    raise ValueError(".env 파일에 GEMINI_API_KEY가 없습니다.")

DART_API_KEY  = os.getenv("DART_API_KEY", "")
GMAIL_USER    = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW  = os.getenv("GMAIL_APP_PASSWORD", "")
SEND_EMAIL    = os.getenv("SEND_EMAIL", "true").lower() in ("true", "1", "yes")
_extra_emails = [e.strip() for e in re.split(r"[,;\n]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
_main_emails  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
NOTIFY_EMAILS = _main_emails + _extra_emails

gemini_client = genai.Client(api_key=GEMINI_API_KEY)
GEMINI_MODEL  = "gemini-2.5-flash-lite"

TODAY       = date.today()
MONTH_START = date(TODAY.year, TODAY.month, 1)   # 이번 달 1일부터 포함
DAYS_AHEAD  = 150  # 달력 표시 범위(5개월)와 일치
DATE_END    = TODAY + timedelta(days=DAYS_AHEAD)
KST        = timezone(timedelta(hours=9))

GEMINI_CONCURRENCY = 10
_gemini_sem = threading.Semaphore(GEMINI_CONCURRENCY)

SHEET_SCHEDULE = "투자일정"
SHEET_INSIGHT  = "종목인사이트"
SHEET_THEMATIC = "테마별일정"

_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}


# ── Gemini 공통 유틸 ──────────────────────────────────────────────────────────

def _gemini_generate(prompt, max_retries=5):
    """Gemini API 호출. 일시 오류는 지수 백오프(최대 5회)로 재시도."""
    for attempt in range(max_retries):
        try:
            return gemini_client.models.generate_content(model=GEMINI_MODEL, contents=prompt)
        except Exception as e:
            err = str(e)
            retryable = any(c in err for c in ("503", "UNAVAILABLE", "429", "RESOURCE_EXHAUSTED"))
            if attempt < max_retries - 1 and retryable:
                wait = min(5 * (2 ** attempt), 30) + random.uniform(0, 2)
                console.print(f"[yellow][RETRY {attempt+1}/{max_retries}] {wait:.0f}초 대기...[/yellow]")
                time.sleep(wait)
            else:
                raise


def _extract_json(text: str):
    """Gemini 응답에서 JSON 블록 추출."""
    m = re.search(r'```(?:json)?\s*(\{.*?\}|\[.*?\])\s*```', text, re.DOTALL)
    if m:
        return m.group(1)
    m = re.search(r'(\[.*\]|\{.*\})', text, re.DOTALL)
    return m.group(1) if m else None


# ── 날짜 파싱 / 분류 헬퍼 ────────────────────────────────────────────────────

def _parse_date(text: str, ref_year: int = None) -> str:
    """날짜 문자열을 YYYY-MM-DD로 변환. 실패 시 빈 문자열."""
    text = text.strip().replace(" ", "")
    if not text:
        return ""
    year = ref_year or TODAY.year
    patterns = [
        (r"(\d{4})[.\-/](\d{1,2})[.\-/](\d{1,2})", lambda m: (int(m.group(1)), int(m.group(2)), int(m.group(3)))),
        (r"(\d{1,2})[.\-/](\d{1,2})$",              lambda m: (year, int(m.group(1)), int(m.group(2)))),
    ]
    for pat, extractor in patterns:
        m = re.search(pat, text)
        if m:
            try:
                y, mo, d = extractor(m)
                return date(y, mo, d).strftime("%Y-%m-%d")
            except ValueError:
                continue
    return ""

# ── 데이터 수집 ────────────────────────────────────────────────────────────────

def _collect_dart_events(days_back: int = 30) -> list:
    """DART 공시 API에서 최근 공시를 수집, 미래 이벤트로 변환."""
    if not DART_API_KEY:
        console.print("  [dim]DART_API_KEY 없음 → DART 수집 스킵[/dim]")
        return []

    bgn = (datetime.today() - timedelta(days=days_back)).strftime("%Y%m%d")
    end = datetime.today().strftime("%Y%m%d")

    events = []
    with requests.Session() as session:
        for ptype in ("B", "A"):
            url = (
                f"https://opendart.fss.or.kr/api/list.json"
                f"?crtfc_key={DART_API_KEY}&bgn_de={bgn}&end_de={end}"
                f"&pblntf_ty={ptype}&page_count=100"
            )
            try:
                resp = session.get(url, timeout=15)
                data = resp.json()
            except Exception as e:
                console.print(f"  [yellow][WARN] DART API {ptype} 실패: {e}[/yellow]")
                continue

            if data.get("status") != "000":
                console.print(f"  [yellow][WARN] DART 응답 이상({ptype}): {data.get('message', '')}[/yellow]")
                continue

            for item in data.get("list", []):
                rcept_dt = item.get("rcept_dt", "")
                try:
                    filing_date = datetime.strptime(rcept_dt, "%Y%m%d").date()
                except ValueError:
                    filing_date = TODAY

                corp   = item.get("corp_name", "")
                report = item.get("report_nm", "")
                rcept  = item.get("rcept_no", "")
                link   = f"https://dart.fss.or.kr/dsaf001/main.do?rcpNo={rcept}" if rcept else ""

                cat, future_days = "공시", 0
                if "주주총회" in report:
                    cat, future_days = "주주총회", 21
                elif "유상증자" in report or "전환사채" in report or "신주인수권" in report:
                    cat, future_days = "공시", 30
                elif any(k in report for k in ("사업보고서", "분기보고서", "반기보고서", "실적")):
                    cat = "기업실적"

                # 분기/사업보고서는 코스피·코스닥 주요 대기업만 포함 (소형사 제외)
                if cat == "기업실적" and not _is_major_corp(corp):
                    continue

                event_date = filing_date + timedelta(days=future_days)
                if event_date < MONTH_START or event_date > DATE_END:
                    continue

                events.append({
                    "날짜":    event_date.strftime("%Y-%m-%d"),
                    "이벤트명": f"[{corp}] {report}",
                    "카테고리": cat,
                    "출처":    "DART",
                    "원본링크": link,
                })

    console.print(f"  [dim]DART 공시[/dim] — {len(events)}건")
    return events


def _collect_naver_schedule() -> list:
    """네이버 금융 증시일정 스크래핑.

    [비고] 네이버 금융이 Next.js SPA로 전환(2025년 이후)되면서
    기존 /research/schedule_list.naver URL이 404 반환.
    중요 일정은 custom_events.json으로 직접 추가하거나 AI보완 소스로 대체됩니다.
    """
    console.print("  [dim]네이버금융[/dim] — 0건 [dim](URL 폐기됨 → custom_events.json으로 보완)[/dim]")
    return []


def _collect_known_events_via_ai() -> list:
    """Gemini 지식 기반으로 향후 60일 글로벌·국내 주요 투자 이벤트를 생성."""
    today_str = TODAY.strftime("%Y-%m-%d")
    end_str   = DATE_END.strftime("%Y-%m-%d")

    prompt = f"""주식 투자 애널리스트로서 {today_str}~{end_str} 주요 투자 이벤트를 최대 100개 알려주세요.

포함 항목:
- 경제지표: FOMC·ECB·BOJ·한은 금리결정, 미국 CPI·PPI·PCE·NFP·GDP, MSCI 리밸런싱, 코스피200 선물만기
- 기업실적: 엔비디아·애플·MS·메타·알파벳·아마존·테슬라·TSMC·삼성전자·SK하이닉스 등 주요 대형주 실적발표
- 글로벌이벤트: G7·G20·APEC, 미중·한미·한일 정상회담, ASCO·AACR 등 바이오학회, CES·MWC·Google I/O·WWDC 등 IT컨퍼런스
- 외교·정치: 주요국 정상 방한 (미국·중국·일본·EU·사우디 등), 한국 대통령 해외 순방,
  글로벌 빅테크·반도체 CEO 방한 (엔비디아 젠슨황·애플 팀쿡·MS 사티아나델라 등),
  무역협상·관세 협의·FTA, 북미 정상외교, 주요국 선거·정권교체,
  UN총회·WTO·NATO·OECD 정상급 회의

규칙: {today_str} 이전 날짜 제외, 중복 금지, 미확정은 이벤트명에 "(예정)" 표기
카테고리 분류: 경제지표/기업실적/글로벌이벤트/CEO·VIP방한 중 선택
  - CEO·VIP방한: 글로벌 빅테크·반도체·제조업 CEO 방한 또는 주요국 정상 방한 (LG·삼성 등 파트너십 논의 포함)

JSON 배열로만 응답:
[{{"날짜":"YYYY-MM-DD","이벤트명":"이벤트명","카테고리":"경제지표 또는 기업실적 또는 글로벌이벤트 또는 CEO·VIP방한","출처":"AI보완"}}]
"""
    events = []
    try:
        resp = _gemini_generate(prompt)
        raw  = _extract_json(resp.text)
        if raw:
            items = json.loads(raw)
            if isinstance(items, list):
                for item in items:
                    d = _parse_date(str(item.get("날짜", "")), TODAY.year)
                    if not d:
                        continue
                    events.append({
                        "날짜":    d,
                        "이벤트명": str(item.get("이벤트명", "")).strip(),
                        "카테고리": str(item.get("카테고리", "글로벌이벤트")).strip(),
                        "출처":    "AI보완",
                        "원본링크": "",
                    })
    except Exception as e:
        console.print(f"  [yellow][WARN] AI 이벤트 보완 실패: {e}[/yellow]")

    console.print(f"  [dim]AI보완[/dim] — {len(events)}건")
    return events


# 집중 섹터 (한국경제·매일경제 뉴스 필터링 기준)
_FOCUS_SECTORS = (
    "변압기·전력기기 / 이차전지(배터리·ESS) / 조선·해운 / "
    "방산·우주항공 / 반도체·AI·HBM / 원자력·수소에너지"
)

# 코스피/코스닥 주요 기업 키워드 — DART 기업실적 필터링 기준
_MAJOR_CORP_KW: tuple[str, ...] = (
    "삼성", "sk", "lg", "현대", "기아", "포스코", "롯데", "한화", "두산",
    "hd현대", "카카오", "네이버", "셀트리온", "kb금융", "신한", "하나금융",
    "우리금융", "한국전력", "한국가스공사", "kt", "s-oil", "gs", "cj", "ls",
    "고려아연", "한국항공우주", "한국타이어", "금호", "대한항공", "아시아나",
    "아모레퍼시픽", "하이브", "크래프톤", "넷마블", "엔씨소프트", "넥슨",
    "한미약품", "유한양행", "종근당", "대웅제약", "녹십자", "동아에스티",
    "효성", "oci", "코웨이", "bgf리테일", "이마트", "신세계", "현대백화점",
    "두산에너빌리티", "한화에어로", "한화오션", "현대중공업", "현대제철",
    "현대건설", "대우건설", "삼성중공업", "대우조선", "하림",
)


def _is_major_corp(corp_name: str) -> bool:
    """DART 이벤트 기업명이 코스피/코스닥 주요 대기업인지 판별."""
    c = corp_name.lower()
    return any(kw in c for kw in _MAJOR_CORP_KW)


def _parse_rss(url: str, max_items: int = 30) -> list[str]:
    """RSS URL에서 기사 제목 목록 반환. 실패 시 빈 리스트."""
    try:
        resp = requests.get(url, headers=_HEADERS, timeout=10)
        resp.raise_for_status()
        root = ET.fromstring(resp.content)
        titles: list[str] = []
        for item in root.iter("item"):
            t = (item.findtext("title") or "").strip()
            if t:
                titles.append(t)
                if len(titles) >= max_items:
                    break
        return titles
    except Exception:
        return []


def _news_titles_to_events(
    titles: list[str],
    source_name: str,
    prompt_context: str,
    category_hint: str = "글로벌이벤트",
) -> list:
    """뉴스 제목 목록을 Gemini로 구조화하여 이벤트 리스트 반환 (공통 헬퍼)."""
    if not titles:
        return []
    today_str = TODAY.strftime("%Y-%m-%d")
    end_str   = DATE_END.strftime("%Y-%m-%d")
    prompt = f"""아래는 {source_name}에서 수집한 최신 뉴스 제목입니다.
{prompt_context}

[뉴스 제목]
{chr(10).join(f'- {t}' for t in titles[:40])}

추출 조건:
- 단순 주가 등락·시황 뉴스는 제외하고, 정책 발표·회의·계약·실적·전시회 등 일정성 이벤트만 포함
- 날짜가 불명확하면 오늘({today_str}) 기준으로 적고 이벤트명에 "(속보)" 추가
- {today_str} ~ {end_str} 범위 이벤트만 포함
- 없으면 빈 배열 [] 반환

JSON 배열로만 응답:
[
  {{
    "날짜": "YYYY-MM-DD",
    "이벤트명": "구체적 이벤트명",
    "카테고리": "경제지표 또는 기업실적 또는 글로벌이벤트 또는 CEO·VIP방한"
  }}
]
카테고리 규칙: 글로벌 빅테크·반도체 CEO 방한 또는 주요국 정상 방한은 반드시 "CEO·VIP방한"으로 분류."""
    events: list = []
    try:
        resp = _gemini_generate(prompt)
        raw  = _extract_json(resp.text)
        if raw:
            for item in json.loads(raw):
                d = _parse_date(str(item.get("날짜", "")), TODAY.year)
                if not d:
                    continue
                events.append({
                    "날짜":    d,
                    "이벤트명": str(item.get("이벤트명", "")).strip(),
                    "카테고리": str(item.get("카테고리", category_hint)).strip(),
                    "출처":    source_name,
                    "원본링크": "",
                })
    except Exception as e:
        console.print(f"  [yellow][WARN] {source_name} 구조화 실패: {e}[/yellow]")
    return events


def _collect_korean_sector_news() -> list:
    """한국경제·매일경제 RSS에서 집중 섹터 뉴스 수집 → Gemini 구조화."""
    _FEEDS = [
        "https://www.hankyung.com/feed/all-news",
        "https://www.mk.co.kr/rss/30000001/",
    ]
    seen: set = set()
    titles: list[str] = []
    for url in _FEEDS:
        for t in _parse_rss(url):
            if t not in seen:
                seen.add(t)
                titles.append(t)

    if not titles:
        console.print("  [dim]국내뉴스[/dim] — 0건 [dim](RSS 수집 실패)[/dim]")
        return []

    events = _news_titles_to_events(
        titles,
        source_name="국내뉴스",
        prompt_context=(
            f"집중 섹터: {_FOCUS_SECTORS}\n"
            "위 섹터에 해당하는 이벤트만 추출하세요."
        ),
        category_hint="글로벌이벤트",
    )
    console.print(f"  [dim]국내뉴스[/dim] — {len(events)}건")
    return events


def _collect_global_news() -> list:
    """블룸버그·로이터 RSS에서 글로벌 매크로·외교 이슈 수집 → Gemini 구조화."""
    _FEEDS = [
        "https://feeds.bloomberg.com/markets/news.rss",
        "https://feeds.reuters.com/reuters/businessNews",
    ]
    seen: set = set()
    titles: list[str] = []
    for url in _FEEDS:
        for t in _parse_rss(url):
            if t not in seen:
                seen.add(t)
                titles.append(t)

    if not titles:
        console.print("  [dim]글로벌뉴스[/dim] — 0건 [dim](RSS 수집 실패)[/dim]")
        return []

    events = _news_titles_to_events(
        titles,
        source_name="글로벌뉴스",
        prompt_context=(
            "집중 영역: 연준(Fed) 정책·금리 / 미중 관계·무역협상 / "
            "유럽·일본 경제 / 글로벌 원자재·에너지 / 외국인 자금흐름\n"
            "한국 주식 시장에 영향을 미칠 거시경제·외교 이벤트만 추출하세요."
        ),
        category_hint="경제지표",
    )
    console.print(f"  [dim]글로벌뉴스[/dim] — {len(events)}건")
    return events


def _collect_investing_calendar() -> list:
    """인베스팅닷컴 경제 캘린더 POST API 수집 (JS 렌더링 우회)."""
    events = []
    try:
        headers = {
            **_HEADERS,
            "X-Requested-With": "XMLHttpRequest",
            "Referer": "https://www.investing.com/economic-calendar/",
        }
        resp = requests.post(
            "https://economic-calendar.investing.com/economic-calendar/",
            headers=headers,
            data={
                "timeZone":      "55",        # Asia/Seoul
                "timeFilter":    "timeOnly",
                "currentTab":    "custom",
                "dateFrom":      TODAY.strftime("%Y-%m-%d"),
                "dateTo":        (TODAY + timedelta(days=30)).strftime("%Y-%m-%d"),
                "submitFilters": "1",
                "limit_from":    "0",
            },
            timeout=10,
        )
        soup = BeautifulSoup(resp.text, "html.parser")
        for row in soup.select("tr.js-event-item")[:30]:
            date_el  = row.select_one("td.first.left.time")
            event_el = row.select_one("td.left.event")
            if not date_el or not event_el:
                continue
            date_str  = _parse_date(date_el.get_text(strip=True), TODAY.year)
            event_str = event_el.get_text(strip=True)
            if date_str and event_str:
                events.append({
                    "날짜":    date_str,
                    "이벤트명": event_str,
                    "카테고리": "경제지표",
                    "출처":    "Investing.com",
                    "원본링크": "",
                })
    except Exception:
        pass

    console.print(f"  [dim]Investing.com[/dim] — {len(events)}건")
    return events


def _collect_custom_events() -> list:
    """custom_events.json에서 사용자 정의 이벤트 로드.

    파일 형식:
    [
      {"날짜": "YYYY-MM-DD", "이벤트명": "...", "카테고리": "글로벌이벤트", "원본링크": ""}
    ]
    """
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "custom_events.json")
    if not os.path.exists(json_path):
        return []
    try:
        with open(json_path, encoding="utf-8") as f:
            items = json.load(f)
        events = []
        for item in items:
            d = _parse_date(str(item.get("날짜", "")), TODAY.year)
            if not d:
                continue
            events.append({
                "날짜":    d,
                "이벤트명": str(item.get("이벤트명", "")).strip(),
                "카테고리": str(item.get("카테고리", "글로벌이벤트")).strip(),
                "출처":    "사용자지정",
                "원본링크": str(item.get("원본링크", "")),
            })
        if events:
            console.print(f"  [dim]사용자지정[/dim] — {len(events)}건")
        return events
    except Exception as e:
        console.print(f"  [yellow][WARN] custom_events.json 로드 실패: {e}[/yellow]")
        return []


def _collect_labor_events() -> list:
    """네이버 뉴스에서 대기업 파업·노조·쟁의 기사 제목을 수집하고
    Gemini로 날짜·종목을 구조화하여 이벤트 목록 반환."""
    _LABOR_QUERIES = ["대기업 파업", "노조 파업", "쟁의행위 선언"]
    _MAJOR_CORPS   = (
        "삼성|LG|SK|현대|기아|포스코|롯데|한화|두산|HD현대|"
        "카카오|네이버|쿠팡|CJ|GS|코스트코|GM|르노|KT|KT&G"
    )

    # 1. 네이버 뉴스 검색 — 기사 제목 수집
    titles: list[str] = []
    seen: set[str] = set()
    for kw in _LABOR_QUERIES:
        try:
            resp = requests.get(
                "https://search.naver.com/search.naver",
                params={"where": "news", "query": kw, "sort": "1", "pd": "4"},
                headers=_HEADERS, timeout=10,
            )
            soup = BeautifulSoup(resp.text, "html.parser")
            for el in soup.select("a.news_tit")[:15]:
                title = el.get_text(strip=True)
                if title and title not in seen:
                    seen.add(title)
                    titles.append(title)
        except Exception:
            pass

    if not titles:
        console.print("  [dim]노동이슈[/dim] — 0건 [dim](뉴스 수집 실패)[/dim]")
        return []

    # 2. Gemini로 구조화
    today_str = TODAY.strftime("%Y-%m-%d")
    end_str   = DATE_END.strftime("%Y-%m-%d")
    prompt = f"""아래는 최근 수집된 파업·노조 관련 뉴스 제목입니다.
오늘({today_str})부터 {end_str} 사이에 예정되거나 진행 중인 파업·쟁의·노조 이벤트만 추출해주세요.

[뉴스 제목]
{chr(10).join(f"- {t}" for t in titles)}

조건:
- 아래 대기업 그룹 계열사 관련만 포함: {_MAJOR_CORPS}
- 이미 완전히 종료된 파업은 제외 (현재 진행 중이면 포함)
- 날짜 불명확 시 기사 맥락으로 추정하고 이벤트명에 "(예정)" 표기
- 오늘({today_str}) 이전 날짜는 포함하지 마세요

다음 JSON 배열 형식으로만 응답 (해당 이벤트 없으면 빈 배열 []):
[
  {{
    "날짜": "YYYY-MM-DD",
    "이벤트명": "회사명 파업/쟁의 내용 (예: 삼성전자 전국노조 총파업 시작)"
  }}
]
"""
    events: list = []
    try:
        resp = _gemini_generate(prompt)
        raw  = _extract_json(resp.text)
        if raw:
            items = json.loads(raw)
            if isinstance(items, list):
                for item in items:
                    d = _parse_date(str(item.get("날짜", "")), TODAY.year)
                    if not d:
                        continue
                    events.append({
                        "날짜":    d,
                        "이벤트명": str(item.get("이벤트명", "")).strip(),
                        "카테고리": "글로벌이벤트",
                        "출처":    "뉴스",
                        "원본링크": "",
                    })
    except Exception as e:
        console.print(f"  [yellow][WARN] 노동이슈 구조화 실패: {e}[/yellow]")

    console.print(f"  [dim]노동이슈[/dim] — {len(events)}건")
    return events


def _collect_diplomatic_events() -> list:
    """네이버 뉴스에서 정상회담·국빈방문·CEO 방한·대통령 순방 뉴스 수집 → Gemini 구조화.

    한국 증시에 직접 영향을 주는 외교·정치 이벤트를 포착한다.
    예) 젠슨황 방한, 미중 정상회담, 한미 정상회담, 대통령 순방 등
    """
    _QUERIES = [
        "정상회담 방한",
        "대통령 순방 일정",
        "CEO 방한 한국 방문",
        "국빈 방문 한국",
        "외교 회담 무역협상",
    ]
    seen: set[str] = set()
    titles: list[str] = []
    for kw in _QUERIES:
        try:
            resp = requests.get(
                "https://search.naver.com/search.naver",
                params={"where": "news", "query": kw, "sort": "1", "pd": "4"},
                headers=_HEADERS, timeout=10,
            )
            soup = BeautifulSoup(resp.text, "html.parser")
            for el in soup.select("a.news_tit")[:10]:
                title = el.get_text(strip=True)
                if title and title not in seen:
                    seen.add(title)
                    titles.append(title)
        except Exception:
            pass

    if not titles:
        console.print("  [dim]외교뉴스[/dim] — 0건 [dim](뉴스 수집 실패)[/dim]")
        return []

    events = _news_titles_to_events(
        titles,
        source_name="외교뉴스",
        prompt_context=(
            "아래 항목에 해당하는 이벤트만 추출하세요:\n"
            "- 외국 정상·국빈의 한국 방문 (미국·중국·일본·EU·사우디·UAE 등)\n"
            "- 한국 대통령의 해외 순방 일정\n"
            "- 글로벌 빅테크·반도체 CEO 방한 (엔비디아 젠슨황, 애플 팀쿡, MS 사티아나델라 등)\n"
            "- 주요 외교 회담 / 무역협상 / 관세 협의\n"
            "- G7·G20·APEC·UN총회 등 정상급 국제회의\n"
            "한국 주식 시장(특히 수출주·방산·반도체·조선)에 영향을 미치는 이벤트를 우선합니다.\n"
            "카테고리 규칙: 글로벌 빅테크·반도체·제조업 CEO 방한 또는 주요국 정상 방한은 반드시 "
            "\"CEO·VIP방한\"으로 분류. 그 외 외교·정상회담은 \"글로벌이벤트\"로 분류."
        ),
        category_hint="CEO·VIP방한",
    )
    console.print(f"  [dim]외교뉴스[/dim] — {len(events)}건")
    return events


def collect_all_events() -> list:
    """9개 소스 병렬 수집 (DART / AI보완 / 국내뉴스 / 글로벌뉴스 / 외교뉴스 / 노동이슈 / Investing.com / 사용자지정)."""
    console.print("[bold]▶ 투자 이벤트 수집 중...[/bold]")
    with ThreadPoolExecutor(max_workers=9) as ex:
        f_dart       = ex.submit(_collect_dart_events, 30)
        f_naver      = ex.submit(_collect_naver_schedule)
        f_ai         = ex.submit(_collect_known_events_via_ai)
        f_kr_news    = ex.submit(_collect_korean_sector_news)
        f_gl_news    = ex.submit(_collect_global_news)
        f_diplomatic = ex.submit(_collect_diplomatic_events)
        f_labor      = ex.submit(_collect_labor_events)
        f_investing  = ex.submit(_collect_investing_calendar)
        f_custom     = ex.submit(_collect_custom_events)
        raw = (
            f_dart.result()
            + f_naver.result()
            + f_ai.result()
            + f_kr_news.result()
            + f_gl_news.result()
            + f_diplomatic.result()
            + f_labor.result()
            + f_investing.result()
            + f_custom.result()
        )
    return raw


# ── 정규화 ────────────────────────────────────────────────────────────────────

def _normalize_events(raw: list) -> list:
    """날짜 검증 · D-Day 계산 · 범위 필터(오늘~+60일) · 중복 제거 · 정렬."""
    seen  = set()
    valid = []
    for ev in raw:
        date_str = ev.get("날짜", "")
        if not date_str:
            continue
        try:
            ev_date = date.fromisoformat(date_str)
        except ValueError:
            continue
        if ev_date < MONTH_START or ev_date > DATE_END:
            continue
        key = (date_str, ev.get("이벤트명", "")[:10])
        if key in seen:
            continue
        seen.add(key)
        valid.append({**ev, "D-Day": (ev_date - TODAY).days})

    valid.sort(key=lambda x: x["날짜"])
    return valid


# ── 분석 스킵 규칙 ────────────────────────────────────────────────────────────

_SKIP_SOURCES = {"DART"}
_SKIP_CATS    = {"공시", "기업실적"}

def _should_skip_analysis(ev: dict) -> bool:
    """Gemini 분석을 생략할 저영향 이벤트 판별.
    DART 공시·기업실적은 달력에서도 제외되며 수혜종목 매핑이 불필요."""
    return ev.get("출처", "") in _SKIP_SOURCES and ev.get("카테고리", "") in _SKIP_CATS

def _default_analysis(ev: dict) -> dict:
    """분석 생략 이벤트에 채울 기본값."""
    return {
        **ev,
        "섹터":        "",
        "영향방향":     "중립",
        "직접수혜종목": [],
        "간접수혜종목": [],
        "투자포인트":   "",
        "영향도":       2,
        "지속기간":     "단기(1주 이내)",
        "투자전략":     "",
    }


# ── Gemini 분석 ───────────────────────────────────────────────────────────────

def _analyze_events_batch(chunk: list) -> list:
    """여러 이벤트(배치)를 한 번의 Gemini API 호출로 분석하여 속도를 획기적으로 향상시킵니다."""
    if not chunk:
        return []

    events_text = ""
    for i, ev in enumerate(chunk):
        events_text += f"[ID: {i}]\n"
        events_text += f"이벤트명: {ev.get('이벤트명', '')}\n"
        events_text += f"날짜: {ev.get('날짜', '')}\n"
        events_text += f"카테고리: {ev.get('카테고리', '')}\n\n"

    prompt = f"""
당신은 주식 투자 전문 애널리스트입니다. 아래 여러 투자 이벤트를 개별적으로 분석하세요.

{events_text}
## 분석 지침
1. 각 이벤트가 한국 주식 시장에 미치는 영향을 분석하세요.
2. 직접수혜종목: 이 이벤트가 직접적 호재인 한국 상장 종목 (최대 5개)
3. 간접수혜종목: 공급망·경쟁구도·테마로 간접 수혜를 받는 종목 (최대 5개)
4. 글로벌 이벤트는 국내 수혜 산업군을 중심으로 분석하세요.

다음 JSON 배열 형식으로만 응답하세요. 입력받은 ID를 반드시 포함하여 매핑할 수 있게 하세요:
[
  {{
    "id": 0,
    "섹터": "주요 수혜 섹터명",
    "영향방향": "긍정 또는 부정 또는 중립",
    "직접수혜종목": ["종목1", "종목2"],
    "간접수혜종목": ["종목1", "종목2"],
    "투자포인트": "이 이벤트가 주가에 미치는 구체적 영향 (2문장)",
    "영향도": 1에서 10 사이 정수,
    "지속기간": "단기(1주 이내) 또는 중기(1~3개월) 또는 장기(3개월 이상)",
    "투자전략": "포지션 구축 타이밍과 주요 관심 종목 관점 (1~2문장)"
  }}
]
"""
    with _gemini_sem:
        resp = _gemini_generate(prompt)
    raw = _extract_json(resp.text)
    if not raw:
        raise ValueError("JSON 응답이 없습니다.")
    
    parsed = json.loads(raw)
    if not isinstance(parsed, list):
        raise ValueError("응답이 JSON 배열이 아닙니다.")

    results = []
    for item in parsed:
        idx = item.get("id")
        if idx is not None and 0 <= idx < len(chunk):
            ev = chunk[idx]
            item.update({
                "날짜":    ev["날짜"],
                "D-Day":   ev["D-Day"],
                "이벤트명": ev["이벤트명"],
                "카테고리": ev["카테고리"],
                "출처":    ev["출처"],
                "원본링크": ev.get("원본링크", ""),
            })
            results.append(item)
    return results


def _synthesize_events(analyzed: list) -> dict:
    """전체 분석 결과로 종목별 언급 집계 + Gemini 종합 투자의견 생성."""
    direct_count:  Counter = Counter()
    indirect_count: Counter = Counter()
    stock_events:  dict = defaultdict(list)
    stock_dates:   dict = defaultdict(list)

    for ev in analyzed:
        ev_name = ev.get("이벤트명", "")
        ev_date = ev.get("날짜", "")
        for s in ev.get("직접수혜종목", []):
            s = s.strip()
            if s and s not in ("nan", ""):
                direct_count[s] += 1
                stock_events[s].append(ev_name)
                stock_dates[s].append(ev_date)
        for s in ev.get("간접수혜종목", []):
            s = s.strip()
            if s and s not in ("nan", ""):
                indirect_count[s] += 1
                stock_events[s].append(ev_name)
                stock_dates[s].append(ev_date)

    all_stocks = set(direct_count) | set(indirect_count)
    score      = {s: direct_count[s] * 2 + indirect_count[s] for s in all_stocks}
    top20      = sorted(score, key=lambda x: score[x], reverse=True)[:20]

    if not top20:
        return {"stock_rows": []}

    stock_summary = [
        {
            "종목명":       s,
            "직접수혜횟수":  direct_count[s],
            "간접수혜횟수":  indirect_count[s],
            "주요관련이벤트": stock_events[s][:3],
        }
        for s in top20
    ]

    prompt = f"""
아래는 향후 60일 주요 투자 이벤트 분석에서 언급된 수혜 종목 목록입니다.
각 종목에 대해 향후 60일 관점의 간결한 종합 투자의견을 작성해주세요.

[종목 데이터]
{json.dumps(stock_summary, ensure_ascii=False, indent=2)}

다음 JSON 배열 형식으로만 응답하세요:
[
  {{
    "종목명": "종목명",
    "종합투자의견": "향후 60일 관점 투자 포인트 1~2문장",
    "주목도": 1에서 10 사이 정수
  }}
]
"""
    opinions = {}
    try:
        resp = _gemini_generate(prompt)
        raw  = _extract_json(resp.text)
        if raw:
            items = json.loads(raw)
            if isinstance(items, list):
                for item in items:
                    name = item.get("종목명", "")
                    if name:
                        opinions[name] = item
    except Exception as e:
        console.print(f"  [yellow][WARN] 종합 투자의견 생성 실패: {e}[/yellow]")

    stock_rows = []
    for s in top20:
        op     = opinions.get(s, {})
        dates  = sorted(set(stock_dates[s]))
        total  = direct_count[s] + indirect_count[s]
        stock_rows.append({
            "종목명":       s,
            "관련이벤트수":  total,
            "직접수혜횟수":  direct_count[s],
            "간접수혜횟수":  indirect_count[s],
            "주요관련이벤트": ", ".join(dict.fromkeys(stock_events[s]))[:150],
            "최근이벤트날짜": dates[-1] if dates else "",
            "종합투자의견":  op.get("종합투자의견", ""),
            "주목도":       op.get("주목도", min(10, total * 2)),
        })

    return {"stock_rows": stock_rows}


# ── 테마별 일정표 생성 ────────────────────────────────────────────────────────

def _synthesize_thematic_schedule(analyzed: list) -> list:
    """분석된 이벤트를 투자 테마별로 그룹화하고 대응 전략 팁 생성."""
    if not analyzed:
        return []

    events_text = "\n".join(
        f"- {ev.get('날짜','')} [{ev.get('카테고리','')}] {ev.get('이벤트명','')}"
        for ev in sorted(analyzed, key=lambda x: x.get("날짜", ""))
        if ev.get("이벤트명") and ev.get("날짜")
    )
    if not events_text:
        return []

    month_str = MONTH_START.strftime("%Y년 %m월")
    prompt = f"""당신은 주식 투자 전문 애널리스트입니다.
아래는 {month_str}부터 향후 5개월간의 주요 투자 이벤트 목록입니다.

[이벤트 목록]
{events_text}

위 이벤트들을 아래 테마별로 분류하고, 테마별 현장 대응 전략을 작성해주세요.
이벤트가 없는 테마는 생략하세요.

테마 분류 기준:
- 반도체/AI: 반도체·AI·데이터센터·GPU 관련 실적·컨퍼런스
- 제약/바이오: 바이오·제약 학회, 임상 발표, 의료기기 전시
- 정치/외교/정책: 정상회담·선거·무역협상·규제·관세
- 금융/경제지표: 금리결정·CPI·GDP 등 매크로 지표
- 빅테크/플랫폼: 빅테크 실적·IT 컨퍼런스·소프트웨어
- 에너지/원자재: 원유·LNG·2차전지·신재생에너지
- 자동차/모빌리티: 완성차 실적·EV·자율주행·모터쇼
- 방산/우주: 방위산업·우주항공·전시회
- 노동/파업: 파업·쟁의·노조 협상

각 테마는 아래 JSON 형식으로 응답 (이벤트 있는 테마만):
[
  {{
    "테마": "테마명",
    "이벤트목록": [
      "MM/DD: 이벤트명 (구체적으로, 날짜 포함)",
      "MM/DD: 이벤트명"
    ],
    "전략팁": "[이벤트 전] 구체적 전략 1~2문장\\n[이벤트 후] 구체적 전략 1~2문장"
  }}
]"""
    try:
        resp = _gemini_generate(prompt)
        raw  = _extract_json(resp.text)
        if raw:
            items = json.loads(raw)
            if isinstance(items, list):
                console.print(f"[green]✔[/green] 테마별 일정표 {len(items)}개 테마 생성\n")
                return items
    except Exception as e:
        console.print(f"  [yellow][WARN] 테마별 일정표 생성 실패: {e}[/yellow]")
    return []


# ── DataFrame 구성 ─────────────────────────────────────────────────────────────

SCHEDULE_COLS = [
    "날짜", "D-Day", "이벤트명", "카테고리", "섹터",
    "직접수혜종목", "간접수혜종목",
    "투자포인트", "영향도", "영향방향", "지속기간", "투자전략",
    "출처", "원본링크",
]

INSIGHT_COLS = [
    "종목명", "관련이벤트수", "직접수혜횟수", "간접수혜횟수",
    "주요관련이벤트", "최근이벤트날짜", "종합투자의견", "주목도",
]

THEMATIC_COLS = ["테마", "주요 일정 및 이벤트", "대응 전략 Tip"]

_LIST_COLS = {"직접수혜종목", "간접수혜종목"}


def _build_schedule_df(analyzed: list) -> pd.DataFrame:
    rows = []
    for ev in analyzed:
        row = {}
        for col in SCHEDULE_COLS:
            val = ev.get(col, "")
            if col in _LIST_COLS and isinstance(val, list):
                val = ", ".join(v.strip() for v in val if v.strip())
            row[col] = val
        rows.append(row)
    if not rows:
        return pd.DataFrame(columns=SCHEDULE_COLS)
    df = pd.DataFrame(rows)
    df["D-Day"] = pd.to_numeric(df["D-Day"], errors="coerce").fillna(0).astype(int)
    df["영향도"] = pd.to_numeric(df["영향도"], errors="coerce").fillna(0).astype(int)
    df = df.sort_values(["날짜", "영향도"], ascending=[True, False]).reset_index(drop=True)
    return df[SCHEDULE_COLS]


def _build_thematic_df(thematic_rows: list) -> pd.DataFrame:
    rows = []
    for item in thematic_rows:
        theme  = item.get("테마", "")
        events = item.get("이벤트목록", [])
        tips   = item.get("전략팁", "")
        if not theme or not events:
            continue
        rows.append({
            "테마":             theme,
            "주요 일정 및 이벤트": "\n".join(events),
            "대응 전략 Tip":      tips,
        })
    if not rows:
        return pd.DataFrame(columns=THEMATIC_COLS)
    return pd.DataFrame(rows, columns=THEMATIC_COLS)


def _build_stock_insight_df(synthesis: dict) -> pd.DataFrame:
    rows = synthesis.get("stock_rows", [])
    if not rows:
        return pd.DataFrame(columns=INSIGHT_COLS)
    df = pd.DataFrame(rows)
    df["주목도"] = pd.to_numeric(df.get("주목도", 0), errors="coerce").fillna(0).astype(int)
    df = df.sort_values("주목도", ascending=False).reset_index(drop=True)
    final = [c for c in INSIGHT_COLS if c in df.columns]
    return df[final]


# ── 이메일 발송 ───────────────────────────────────────────────────────────────

def _send_email(output_file: str, df_schedule: pd.DataFrame, df_insight: pd.DataFrame) -> None:
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from pathlib import Path

    ts    = datetime.now(KST).strftime("%Y-%m-%d %H:%M")
    n_ev  = len(df_schedule) if df_schedule is not None else 0
    n_stk = len(df_insight)  if df_insight  is not None else 0

    event_lines = []
    if df_schedule is not None and not df_schedule.empty:
        future_df = df_schedule[df_schedule["D-Day"] >= 0].sort_values("D-Day")
        for _, row in future_df.head(10).iterrows():
            d   = int(row.get("D-Day", 0))
            dt  = str(row.get("날짜", ""))
            nm  = str(row.get("이벤트명", ""))[:40]
            cat = str(row.get("카테고리", ""))
            tag = "D-Day" if d == 0 else f"D-{d}"
            event_lines.append(f"  {dt} [{tag}] [{cat}] {nm}")

    stock_lines = []
    if df_insight is not None and not df_insight.empty:
        for _, row in df_insight.head(10).iterrows():
            nm  = str(row.get("종목명", ""))
            cnt = int(row.get("관련이벤트수", 0))
            op  = str(row.get("종합투자의견", ""))[:50]
            stock_lines.append(f"  {nm} (이벤트 {cnt}건) — {op}")

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
        f"[{ts}] 투자일정 AI 분석이 완료되었습니다.",
        f"이번 달 1일부터 향후 5개월간의 주요 투자 이벤트를 수집·분석한 결과를 첨부합니다.",
        f"",
        f"{'='*60}",
        f"■ 분석 결과 요약",
        f"{'='*60}",
        f"  수집 이벤트 : {n_ev}건",
        f"  주목 종목   : {n_stk}개",
        f"",
        f"■ 향후 임박 이벤트 Top 10 (날짜순, 지난 일정 제외)",
        *(event_lines if event_lines else ["  (이벤트 없음)"]),
        f"",
        f"■ 주목 종목 Top 10",
        *(stock_lines if stock_lines else ["  (종목 없음)"]),
        f"",
        f"{'='*60}",
        f"■ 첨부 엑셀 구성",
        f"{'='*60}",
        f"  [요약]",
        f"    월별 달력 형태로 5개월치 이벤트를 한눈에 확인",
        f"    · 날짜 셀 클릭 → [투자일정] 시트의 해당 날짜 세부 일정으로 자동 이동",
        f"    · ★ 아이콘: 영향도 8 이상 고영향 이벤트",
        f"    · 색상 범례: 경제지표(연주황) / 기업실적(연초록) / 글로벌이벤트(연파랑)",
        f"                 주주총회(연보라) / 사용자지정(연노랑) / CEO·VIP방한(연보라★) / 오늘(노랑)",
        f"",
        f"  [테마별일정]",
        f"    반도체/AI·제약바이오·정치외교·금융경제지표 등 9개 테마별 이벤트 분류",
        f"    · 각 테마마다 [이벤트 전/후] 현장 대응 전략 포함",
        f"",
        f"  [투자일정]",
        f"    전체 이벤트 상세 목록 — 날짜순 정렬",
        f"    컬럼: 날짜 / D-Day / 이벤트명 / 카테고리 / 섹터",
        f"          직접수혜종목 / 간접수혜종목 / 투자포인트",
        f"          영향도(1~10) / 영향방향 / 지속기간 / 투자전략 / 출처",
        f"    · 종목명 클릭 → 네이버 금융 종목 페이지로 이동",
        f"    · D-Day ≤ 7: 빨강 / D-Day ≤ 30: 주황 / D-Day ≤ 60: 노랑",
        f"",
        f"  [종목인사이트]",
        f"    이벤트에서 언급된 수혜 종목 집계 및 AI 종합 투자의견",
        f"    컬럼: 종목명 / 관련이벤트수 / 직접·간접수혜횟수",
        f"          주요관련이벤트 / 최근이벤트날짜 / 종합투자의견 / 주목도",
        f"    · 주목도 7↑: 연초록 / 4~6: 연노랑",
        f"",
        f"  [종합] / [종합(분기포함)]",
        f"    주목 종목과 4_financial_scanner 데이터를 매칭한 펀더멘털 분석",
        f"    컬럼: 종목명 / 티커 / 업종 / 종합점수",
        f"          현재가 / 적정주가(S-RIM) / 상승여력 / 증권사목표주가",
        f"          매출·영업이익 증가율(25A→26E) / PER / 12M PER",
        f"    · 종합(분기포함): 분기 실적 컬럼 12개 추가",
        f"",
        f"{'='*60}",
        f"■ 데이터 수집 소스 (8개)",
        f"{'='*60}",
        f"  · DART 공시 API    : 주주총회·증자·코스피/코스닥 주요 대기업 실적보고 (소형사 제외)",
        f"  · Gemini AI 보완   : FOMC·정상회담·CEO방한·MSCI 등 반복 글로벌 이벤트",
        f"  · 국내 뉴스        : 한국경제·매일경제 (변압기·이차전지·조선·방산·반도체·원자력)",
        f"  · 글로벌 뉴스      : Bloomberg·Reuters (연준·미중관계·원자재)",
        f"  · 외교 뉴스        : 네이버뉴스 (정상회담·국빈방문·CEO방한·대통령순방)",
        f"  · 노동이슈         : 네이버뉴스 (대기업 파업·쟁의)",
        f"  · Investing.com    : 글로벌 경제지표 발표 일정",
        f"  · 사용자 지정      : custom_events.json 직접 입력",
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
        console.print("[yellow]발송 계정 미설정[/yellow]")
        return
    if not NOTIFY_EMAILS:
        console.print("[yellow]수신자 미설정[/yellow]")
        return

    msg = MIMEMultipart()
    msg["From"]    = GMAIL_USER
    msg["To"]      = ", ".join(_main_emails) if _main_emails else GMAIL_USER
    if _extra_emails:
        msg["Bcc"] = ", ".join(_extra_emails)
    msg["Subject"] = f"[투자일정] {ts} — 이벤트 {n_ev}건 / 주목종목 {n_stk}개"
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


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main():
    _start = datetime.now()
    console.print()
    console.print("════════════════════════════════════════════════════════════════")
    console.print("  [bold cyan]투자일정 수집기[/bold cyan]  — 향후 60일 이벤트 × AI 관련주 매핑")
    console.print(f"  오늘: {TODAY}  |  수집 범위: {MONTH_START} ~ {DATE_END}")
    console.print("════════════════════════════════════════════════════════════════")
    console.print()

    # 1단계: 이벤트 수집
    raw_events = collect_all_events()
    console.print(f"[green]✔[/green] 원시 수집 {len(raw_events)}건\n")

    # 2단계: 정규화
    events = _normalize_events(raw_events)
    if not events:
        console.print("[red]60일 이내 이벤트가 없습니다.[/red]")
        sys.exit(1)
    console.print(f"[green]✔[/green] 정규화 후 {len(events)}건 (중복 제거, 범위 필터)\n")

    # 3단계: 개별 Gemini 분석 (병렬) — 저영향 이벤트는 스킵
    to_analyze = [ev for ev in events if not _should_skip_analysis(ev)]
    skipped    = [_default_analysis(ev) for ev in events if _should_skip_analysis(ev)]
    if skipped:
        console.print(
            f"  [dim]스킵 {len(skipped)}건 (DART 공시·기업실적 — 기본값 적용)[/dim]"
        )
    console.print(
        f"[bold]▶ 이벤트별 AI 분석 중... "
        f"({len(to_analyze)}건 분석 / {len(skipped)}건 스킵, 병렬 {GEMINI_CONCURRENCY}건)[/bold]"
    )
    analyzed: list = list(skipped)   # 스킵 이벤트는 기본값으로 바로 추가
    failed:   list = []
    _lock    = threading.Lock()
    completed = [0]

    def _task(chunk: list):
        try:
            results = _analyze_events_batch(chunk)
            with _lock:
                analyzed.extend(results)
        except Exception as e:
            with _lock:
                for ev in chunk:
                    failed.append({"이벤트명": ev.get("이벤트명", ""), "error": str(e)})
            console.print(f"  [red]ERR[/red] 배치 분석 실패 — {str(e)[:60]}")
        finally:
            with _lock:
                completed[0] += len(chunk)

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=30, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TextColumn("남은"),
        TimeRemainingColumn(),
        console=console,
        transient=True,
    ) as progress:
        total = len(to_analyze)
        t_ev  = progress.add_task("AI 분석 중...", total=total)
        chunk_size = 8
        chunks = [to_analyze[i:i+chunk_size] for i in range(0, total, chunk_size)]
        with ThreadPoolExecutor(max_workers=GEMINI_CONCURRENCY) as ex:
            futures = {ex.submit(_task, chunk): chunk for chunk in chunks}
            for fut in as_completed(futures):
                fut.result()
                progress.update(t_ev, completed=completed[0],
                                description=f"[{completed[0]}/{total}] AI 분석 중...")

    for ev in sorted(analyzed, key=lambda x: x.get("날짜", "")):
        console.print(
            f"  [cyan]{ev.get('날짜','')}[/cyan]  "
            f"[dim]{ev.get('카테고리','')}[/dim]  "
            f"{ev.get('이벤트명','')[:40]}"
        )
    if failed:
        console.print(f"  [yellow]{len(failed)}건 분석 실패[/yellow]")

    if not analyzed:
        console.print("[red]분석 결과가 없습니다.[/red]")
        sys.exit(1)

    analyzed.sort(key=lambda x: (x.get("날짜", ""), -int(x.get("영향도", 0) or 0)))

    # 4단계: 종목 집계 + 종합 투자의견 / 테마별 일정표 (병렬)
    console.print(f"\n[bold]▶ 종목별 종합 투자의견 + 테마별 일정표 생성 중...[/bold]")
    synthesis      = {}
    thematic_rows  = []
    with ThreadPoolExecutor(max_workers=2) as ex:
        f_synthesis = ex.submit(_synthesize_events, analyzed)
        f_thematic  = ex.submit(_synthesize_thematic_schedule, analyzed)
        synthesis     = f_synthesis.result()
        thematic_rows = f_thematic.result()
    n_stocks = len(synthesis.get("stock_rows", []))
    console.print(f"[green]✔[/green] 주목 종목 {n_stocks}개 의견 생성\n")

    # 5단계: DataFrame 구성
    df_schedule  = _build_schedule_df(analyzed)
    df_insight   = _build_stock_insight_df(synthesis)
    df_thematic  = _build_thematic_df(thematic_rows)

    # 6단계: 엑셀 저장 + 서식 적용
    output_dir  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_dir, exist_ok=True)
    ts_str      = datetime.now().strftime("%Y_%m%d_%H%M")
    output_file = os.path.join(output_dir, f"투자일정_{ts_str}.xlsx")

    with Progress(SpinnerColumn(), TextColumn("{task.description}"),
                  console=console, transient=True) as p:
        p.add_task("엑셀 저장 및 서식 적용 중...", total=None)

        # raw 백업 저장 (--patch 모드의 기반 데이터)
        raw_backup = os.path.join(output_dir, f"투자일정_{ts_str}_raw.xlsx")
        with pd.ExcelWriter(raw_backup, engine="openpyxl") as writer:
            df_schedule.to_excel(writer, sheet_name=SHEET_SCHEDULE, index=False)
            df_insight.to_excel(writer,  sheet_name=SHEET_INSIGHT,  index=False)
            if not df_thematic.empty:
                df_thematic.to_excel(writer, sheet_name=SHEET_THEMATIC, index=False)

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_schedule.to_excel(writer, sheet_name=SHEET_SCHEDULE, index=False)
            df_insight.to_excel(writer,  sheet_name=SHEET_INSIGHT,  index=False)
            if not df_thematic.empty:
                df_thematic.to_excel(writer, sheet_name=SHEET_THEMATIC, index=False)

        fmt_warn = None
        try:
            spec = importlib.util.spec_from_file_location(
                "format_excel",
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "format_excel.py"),
            )
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            fmt_path = mod.format_file(output_file)
            shutil.move(fmt_path, output_file)
        except Exception as e:
            fmt_warn = str(e)

    # 완료 요약
    _end    = datetime.now()
    elapsed = _end - _start
    mins, s = divmod(int(elapsed.total_seconds()), 60)

    console.print(f"[bold green]✔ 완료[/bold green]")
    console.print(f"{'='*60}")
    console.print(
        f"  이벤트 [green]{len(analyzed)}건[/green]  |  "
        f"주목종목 [cyan]{n_stocks}개[/cyan]  |  "
        f"실패 [dim]{len(failed)}건[/dim]"
    )
    console.print(f"{'='*60}")
    console.print(f"  저장: [bold]{output_file}[/bold]")
    if fmt_warn:
        console.print(f"  [yellow][WARN] 서식 적용 실패: {fmt_warn}[/yellow]")
    console.print(
        f"  [bold green]⏱ {_start:%H:%M:%S} → {_end:%H:%M:%S}  "
        f"(총 {mins}분 {s}초)[/bold green]\n"
    )

    # 7단계: 이메일 발송
    _send_email(output_file, df_schedule, df_insight)


# ── 패치 모드 ─────────────────────────────────────────────────────────────────

def patch_mode():
    """custom_events.json의 신규 이벤트만 분석하여 최신 Excel에 빠르게 추가.

    전체 재실행(14분+) 없이 1~5건 신규 이벤트를 수십 초 내에 반영한다.
    종목인사이트 시트는 전체 재실행 시 갱신됨.
    """
    console.print()
    console.print("════════════════════════════════════════════════════════════════")
    console.print("  [bold cyan]투자일정 패치 모드[/bold cyan]  — custom_events.json 신규 이벤트 추가")
    console.print("════════════════════════════════════════════════════════════════")
    console.print()

    # 사용자 지정 이벤트 로드
    custom = _collect_custom_events()
    if not custom:
        console.print("[red]custom_events.json에 이벤트가 없습니다.[/red]")
        console.print(f"  파일 경로: {os.path.join(os.path.dirname(os.path.abspath(__file__)), 'custom_events.json')}")
        return

    output_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_dir, exist_ok=True)

    # 최신 raw 백업 파일 탐색
    raw_files = sorted([
        f for f in glob.glob(os.path.join(output_dir, "투자일정_*_raw.xlsx"))
        if "~$" not in f
    ])

    # 기존 이벤트 키 집합 (중복 방지)
    existing_keys: set = set()
    df_old = pd.DataFrame()
    df_insight_old = pd.DataFrame(columns=INSIGHT_COLS)

    if raw_files:
        latest_raw = raw_files[-1]
        console.print(f"  기준 파일: [dim]{latest_raw}[/dim]")
        try:
            df_old = pd.read_excel(latest_raw, sheet_name=SHEET_SCHEDULE, dtype=str)
            for _, row in df_old.iterrows():
                k = (str(row.get("날짜", ""))[:10], str(row.get("이벤트명", ""))[:10])
                existing_keys.add(k)
        except Exception as e:
            console.print(f"  [yellow][WARN] raw 백업 로드 실패: {e}[/yellow]")
        try:
            df_insight_old = pd.read_excel(latest_raw, sheet_name=SHEET_INSIGHT, dtype=str)
        except Exception:
            pass
    else:
        console.print("  [yellow]raw 백업 없음 — 신규 이벤트만으로 파일 생성[/yellow]")

    # 정규화 후 신규만 추출
    all_custom = _normalize_events(custom)
    new_events = [
        ev for ev in all_custom
        if (ev["날짜"][:10], ev["이벤트명"][:10]) not in existing_keys
    ]

    if not new_events:
        console.print("[yellow]custom_events.json의 이벤트가 이미 모두 반영되어 있습니다.[/yellow]")
        return

    console.print(f"[green]✔[/green] 신규 이벤트 {len(new_events)}건 감지\n")

    # 신규 이벤트 AI 분석
    console.print(f"[bold]▶ AI 분석 중... ({len(new_events)}건)[/bold]")
    analyzed_new = []
    chunk_size = 5
    chunks = [new_events[i:i+chunk_size] for i in range(0, len(new_events), chunk_size)]
    for chunk in chunks:
        try:
            results = _analyze_events_batch(chunk)
            analyzed_new.extend(results)
            for ev in chunk:
                console.print(f"  [green]✔[/green] {ev['이벤트명'][:50]}")
        except Exception as e:
            for ev in chunk:
                console.print(f"  [red]ERR[/red] {ev['이벤트명'][:40]} — {e}")

    if not analyzed_new:
        console.print("[red]분석 결과가 없습니다.[/red]")
        return

    # 신규 이벤트 DataFrame
    df_new = _build_schedule_df(analyzed_new)

    # 기존 + 신규 합산
    if not df_old.empty:
        # D-Day 재계산 (기존 이벤트도 오늘 기준으로 갱신)
        def _recalc(row):
            try:
                return (date.fromisoformat(str(row["날짜"])[:10]) - TODAY).days
            except Exception:
                return row.get("D-Day", 0)

        df_old["D-Day"] = df_old.apply(_recalc, axis=1)
        df_combined = pd.concat([df_old, df_new], ignore_index=True)
    else:
        df_combined = df_new

    df_combined["영향도"] = pd.to_numeric(df_combined.get("영향도", 0), errors="coerce").fillna(0).astype(int)
    df_combined["D-Day"]  = pd.to_numeric(df_combined.get("D-Day",  0), errors="coerce").fillna(0).astype(int)
    df_combined = df_combined.sort_values(["날짜", "영향도"], ascending=[True, False]).reset_index(drop=True)
    df_combined = df_combined[[c for c in SCHEDULE_COLS if c in df_combined.columns]]

    # 저장
    ts_str   = datetime.now().strftime("%Y_%m%d_%H%M")
    raw_path = os.path.join(output_dir, f"투자일정_{ts_str}_raw.xlsx")
    out_path = os.path.join(output_dir, f"투자일정_{ts_str}.xlsx")

    with pd.ExcelWriter(raw_path, engine="openpyxl") as writer:
        df_combined.to_excel(writer, sheet_name=SHEET_SCHEDULE, index=False)
        df_insight_old.to_excel(writer, sheet_name=SHEET_INSIGHT, index=False)

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        df_combined.to_excel(writer, sheet_name=SHEET_SCHEDULE, index=False)
        df_insight_old.to_excel(writer, sheet_name=SHEET_INSIGHT, index=False)

    # 서식 적용
    fmt_warn = None
    try:
        spec = importlib.util.spec_from_file_location(
            "format_excel",
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "format_excel.py"),
        )
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        fmt_path = mod.format_file(out_path)
        shutil.move(fmt_path, out_path)
    except Exception as e:
        fmt_warn = str(e)

    console.print(f"\n[bold green]✔ 패치 완료 — {len(analyzed_new)}건 추가[/bold green]")
    console.print(f"  저장: [bold]{out_path}[/bold]")
    if fmt_warn:
        console.print(f"  [yellow][WARN] 서식 적용 실패: {fmt_warn}[/yellow]")
    console.print(f"  [dim]종목인사이트는 python main.py 전체 재실행 시 갱신됩니다.[/dim]")


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="투자일정 수집기")
    parser.add_argument(
        "--patch", action="store_true",
        help="custom_events.json 신규 이벤트만 분석하여 최신 Excel에 추가 (빠른 업데이트)",
    )
    args = parser.parse_args()

    if args.patch:
        patch_mode()
    else:
        main()
