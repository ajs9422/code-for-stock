# =============================================================================
# 뉴스기사 인사이트 수집기 — 호재성 종목 · 관련 종목 발굴
#
# [개요]
# 7개 소스(네이버 경제/증권·4개 금융 RSS·KRX 공시)에서 최근 36시간 이내 뉴스를
# 자동 수집하고, Gemini AI로 개별 분석 + 종합 합성하여 엑셀 4개 시트로 정리합니다.
# 단순 종목 언급이 아니라 '호재성 기사'만 선별하고, 직접수혜 → 공급망 → 간접수혜
# 체인으로 관련 종목을 발굴하는 것이 핵심 목적입니다.
#
# [동작 순서]
# 1. 뉴스 수집 (7개 소스 병렬, 최대 ~130건, 중복 URL 제거)
#    - 네이버 경제 헤드라인 (sid1=101) : 36시간 이내, 최대 15건/일 (오늘·어제 페이지)
#    - 네이버 증권 섹션   (sid2=258)  : 36시간 이내, 최대 10건/일 (오늘·어제 페이지)
#    - 한국경제 RSS                   : 최대 20건
#    - 연합뉴스 경제 RSS              : 최대 20건
#    - 매일경제 RSS                   : 최대 20건
#    - 서울경제 RSS                   : 최대 20건
#    - KRX KIND 공시                  : 최대 15건 (대규모계약·지분변동 등 핵심만 필터)
#
# 2. 기사 본문 추출 (3단계 폴백)
#    - newspaper3k → requests + 네이버 본문 셀렉터 → og:description 메타태그
#
# 3. Gemini AI 개별 분석 (gemini-2.5-flash, 병렬 5건 동시)
#    - 제목 기반 사전 필터 (연예·스포츠·날씨 등 투자 무관 기사 스킵)
#    - 투자관련 여부 판단 → false 이면 스킵
#    - 중요도 3점 미만 AND 단기폭발력 5점 미만 → 스킵
#    - [핵심] 호재 종목과 악재 종목 분리:
#        직접수혜종목  : 기사가 명확히 호재인 상장사만
#        간접수혜종목  : 공급망·경쟁구도·정책 수혜로 오를 수 있는 연관사 (최소 2개 추론)
#        공급망연관종목: 직접수혜종목의 납품업체·고객사 중 상장사 (1차 공급망)
#        부정영향종목  : 이 뉴스로 악재를 받을 상장사
#    - 7대 폭등 전조 시그널 탐지 / 단기폭발력(1~10) / 뉴스방향성(긍정·부정·중립)
#    - 가치주 시그널: 적자→흑자 전환, 저PER+실적 급증, 목표가 대폭 상향
#
# 4. Gemini AI 합성 분석 (전체 결과 기반, 1회)
#    - 유망 섹터/테마 Top5 (직접수혜·간접수혜 종목 분리 포함)
#    - 주목 종목 Top10 (폭등후보 → 공급망연관 → 간접수혜 → 가치주 → 복수언급 순)
#    - 악재 경고 종목 (부정영향종목 중 복수 기사 언급 위주)
#    - 시장 방향성 요약 + 투자 주의사항
#
# 5. 엑셀 저장 (4개 시트) + 서식 적용 (format_excel.py)
#    - Sheet1 투자인사이트: 합성 분석 4섹션
#        섹션A 시장 방향성 / 섹션B 유망 섹터·테마 / 섹션C 주목 종목 / 섹션D 악재 경고
#    - Sheet2 종목뉴스   : 호재 기사 (발굴점수 = 단기폭발력×2 + 중요도 순)
#    - Sheet3 시장일반   : 거시경제·정책 뉴스
#    - Sheet4 재무제표   : 직접수혜종목 × 4_financial_scanner 재무 데이터 매칭
#
# [출력 파일]
# - output/뉴스기사_인사이트_YYYY_MMDD_HHhMMm.xlsx
#
# [비용]
# - Gemini 2.5 Flash 기준, ~130건 분석 + 합성 시 약 $0.08 (≈ 115원)
#
# [필요 패키지]
# pip install requests beautifulsoup4 newspaper3k google-genai
#             python-dotenv pandas openpyxl rich lxml_html_clean
#
# [환경 변수 (.env)]
# GEMINI_API_KEY=...   # Google AI Studio에서 발급
# =============================================================================

import json
import re
import sys
import time
import os
import importlib.util
import shutil
import xml.etree.ElementTree as ET
import random
import threading
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta, timezone
from email.utils import parsedate_to_datetime

import requests
from bs4 import BeautifulSoup
from newspaper import Article
from google import genai
import pandas as pd
from dotenv import load_dotenv
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

# Windows 콘솔 한글 깨짐 방지
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

console = Console(legacy_windows=False)

load_dotenv()

# ── Ollama (Local LLM) 설정 ──────────────────────────────────────────────────
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
OLLAMA_MODEL    = os.getenv("OLLAMA_MODEL", "qwen2.5:14b")
USE_LOCAL_LLM      = os.getenv("USE_LOCAL_LLM", "true").lower() in ("true", "1", "yes")
GEMINI_SYNTHESIS   = os.getenv("GEMINI_SYNTHESIS", "true").lower() in ("true", "1", "yes")

# ── Gemini 설정 ───────────────────────────────────────────────────────────────
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
if GEMINI_API_KEY:
    gemini = genai.Client(api_key=GEMINI_API_KEY)
else:
    gemini = None
    if not USE_LOCAL_LLM:
        raise ValueError(".env 파일에 GEMINI_API_KEY가 없습니다.")
GEMINI_MODEL      = "gemini-2.5-flash-lite"
GEMINI_MODEL_LITE = "gemini-2.5-flash-lite"

# 이메일 설정
GMAIL_USER    = os.getenv("GMAIL_USER", "")
GMAIL_APP_PW  = os.getenv("GMAIL_APP_PASSWORD", "")
SEND_EMAIL    = os.getenv("SEND_EMAIL", "true").lower() in ("true", "1", "yes")
_extra_emails = [e.strip() for e in re.split(r"[,;\n]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
_main_emails  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
NOTIFY_EMAILS = _main_emails + _extra_emails

NAVER_BASE_URL = "https://news.naver.com"

# ── RSS 소스 목록 (source_name, rss_url) ─────────────────────────────────────
RSS_SOURCES = [
    ("한국경제", "https://rss.hankyung.com/economy.xml"),
    ("연합뉴스", "https://www.yna.co.kr/rss/economy.xml"),
    ("매일경제", "https://www.mk.co.kr/rss/30100041/"),
    ("서울경제", "https://www.sedaily.com/RSS/Economy"),
]

# DART API는 키 없이 접근 불가 → KIND 스크래핑으로 대체
DART_SCRAPE_URL = "https://kind.krx.co.kr/disclosure/todaydisclosure.do"

SHEET_INSIGHT = "투자인사이트"
SHEET_STOCKS  = "종목뉴스"
SHEET_MARKET  = "시장일반"
SHEET_BRIEF   = "종합"
SHEET_FULL    = "종합(분기포함)"


# ── 뉴스 수집 ─────────────────────────────────────────────────────────────────

def get_naver_economy_news(days=2, limit_per_day=15):
    """네이버 경제 섹션 헤드라인 뉴스 수집 (36시간 이내 — 오늘·어제 날짜 페이지 커버)"""
    headers    = {"User-Agent": "Mozilla/5.0"}
    news_items = []
    seen_urls  = set()

    for d in range(days):
        date = (datetime.now() - timedelta(days=d)).strftime("%Y%m%d")
        url  = (f"{NAVER_BASE_URL}/main/list.naver"
                f"?mode=LSD&mid=sec&sid1=101&date={date}&page=1")
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            resp.raise_for_status()
        except requests.RequestException as e:
            console.print(f"[yellow][WARN] 경제헤드라인 {date} 실패: {e}[/yellow]")
            continue

        soup  = BeautifulSoup(resp.text, 'html.parser')
        count = 0

        # 1순위: 헤드라인 섹션
        for a in soup.select('ul.hdline_news dt a'):
            if count >= limit_per_day:
                break
            title = a.get_text(strip=True)
            href  = a.get('href', '')
            if not title or not href:
                continue
            link = href if href.startswith('http') else NAVER_BASE_URL + href
            if link in seen_urls:
                continue
            seen_urls.add(link)
            news_items.append({"title": title, "link": link, "date": date, "source": "네이버경제"})
            count += 1

        # 2순위: 일반 목록으로 보완
        if count < limit_per_day:
            for a in soup.select('dt:not(.photo) a'):
                if count >= limit_per_day:
                    break
                title = a.get_text(strip=True)
                href  = a.get('href', '')
                if not title or not href:
                    continue
                link = href if href.startswith('http') else NAVER_BASE_URL + href
                if link in seen_urls:
                    continue
                seen_urls.add(link)
                news_items.append({"title": title, "link": link, "date": date, "source": "네이버경제"})
                count += 1

        console.print(f"  [dim]네이버경제 {date}[/dim] — {count}건")

    return news_items, seen_urls


def get_naver_securities_news(days=2, limit_per_day=10, seen_urls=None):
    """네이버 뉴스 증권 섹션(sid2=258) 수집 — 종목/수급/실적 특화 (36시간 이내)"""
    if seen_urls is None:
        seen_urls = set()
    headers    = {"User-Agent": "Mozilla/5.0"}
    news_items = []

    for d in range(days):
        date = (datetime.now() - timedelta(days=d)).strftime("%Y%m%d")
        url  = (f"{NAVER_BASE_URL}/main/list.naver"
                f"?mode=LSD&mid=sec&sid1=101&sid2=258&date={date}&page=1")
        try:
            resp = requests.get(url, headers=headers, timeout=10)
            resp.raise_for_status()
        except requests.RequestException as e:
            console.print(f"[yellow][WARN] 네이버증권 {date} 실패: {e}[/yellow]")
            continue

        soup  = BeautifulSoup(resp.text, 'html.parser')
        count = 0
        for a in soup.select('dt:not(.photo) a'):
            if count >= limit_per_day:
                break
            title = a.get_text(strip=True)
            href  = a.get('href', '')
            if not title or not href:
                continue
            link = href if href.startswith('http') else NAVER_BASE_URL + href
            if link in seen_urls:
                continue
            seen_urls.add(link)
            news_items.append({"title": title, "link": link, "date": date, "source": "네이버증권"})
            count += 1

        console.print(f"  [dim]네이버증권 {date}[/dim] — {count}건")

    return news_items


def get_rss_news(source_name, rss_url, limit=20, seen_urls=None):
    """RSS 피드에서 최근 36시간 이내 뉴스 수집"""
    if seen_urls is None:
        seen_urls = set()
    headers    = {"User-Agent": "Mozilla/5.0"}
    news_items = []
    cutoff     = datetime.now() - timedelta(hours=36)

    try:
        resp = requests.get(rss_url, headers=headers, timeout=10)
        resp.raise_for_status()
        try:
            root = ET.fromstring(resp.content)
        except ET.ParseError:
            from lxml import etree as _lxml
            root = _lxml.fromstring(resp.content, _lxml.XMLParser(recover=True))
    except Exception as e:
        console.print(f"[yellow][WARN] {source_name} RSS 실패: {e}[/yellow]")
        return news_items

    items = root.findall('.//item')

    count = 0
    for item in items:
        if count >= limit:
            break

        title_el = item.find('title')
        link_el  = item.find('link')
        date_el  = item.find('pubDate')

        title = (title_el.text or '').strip() if title_el is not None else ''
        link  = (link_el.text or '').strip()  if link_el  is not None else ''

        # CDATA 래퍼 제거
        title = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', title, flags=re.DOTALL).strip()
        link  = re.sub(r'<!\[CDATA\[(.*?)\]\]>', r'\1', link,  flags=re.DOTALL).strip()

        if not title or not link:
            continue

        # 날짜 필터 (36시간 이내)
        date_str = datetime.now().strftime("%Y%m%d")
        if date_el is not None and date_el.text:
            try:
                pub_dt    = parsedate_to_datetime(date_el.text)
                pub_naive = pub_dt.replace(tzinfo=None)
                if pub_naive < cutoff:
                    continue
                date_str = pub_naive.strftime("%Y%m%d")
            except Exception:
                pass

        if link in seen_urls:
            continue
        seen_urls.add(link)
        news_items.append({"title": title, "link": link, "date": date_str, "source": source_name})
        count += 1

    console.print(f"  [dim]{source_name} RSS[/dim] — {count}건")
    return news_items


def get_kind_disclosures(limit=15, seen_urls=None):
    """KRX KIND 공시정보 수집 — 대규모계약·지분변동·최대주주변경 등 폭등 전조"""
    if seen_urls is None:
        seen_urls = set()
    news_items = []
    headers = {"User-Agent": "Mozilla/5.0"}
    # KIND 오늘의 공시 목록 (주요사항보고)
    try:
        url = ("https://kind.krx.co.kr/disclosure/todaydisclosure.do"
               "?method=searchTodayDisclosureSub&currentPageSize=30"
               "&pageIndex=1&orderMode=0&orderStat=D&forward=todaydisclosure_sub"
               "&chose_category=1&marketType=")
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        count = 0
        for row in soup.select('tr'):
            if count >= limit:
                break
            cols = row.select('td')
            if len(cols) < 3:
                continue
            a = cols[2].select_one('a') if len(cols) > 2 else None
            if not a:
                continue
            title = a.get_text(strip=True)
            href = a.get('href', '')
            # 투자 핵심 공시 키워드 필터
            keywords = ['대규모', '지분', '최대주주', '합병', '분할', '유상증자',
                        '전환사채', '신규사업', '수주', '계약', '특허', '인수',
                        '매각', '자사주', '배당', '실적', '흑자전환']
            if not any(k in title for k in keywords):
                continue
            company = cols[1].get_text(strip=True) if len(cols) > 1 else ''
            link = f"https://kind.krx.co.kr{href}" if href.startswith('/') else href
            if link in seen_urls or not link:
                continue
            seen_urls.add(link)
            full_title = f"[공시] {company} — {title}" if company else f"[공시] {title}"
            news_items.append({
                "title": full_title, "link": link,
                "date": datetime.now().strftime("%Y%m%d"), "source": "KRX공시"
            })
            count += 1
        console.print(f"  [dim]KRX공시[/dim] — {count}건")
    except Exception as e:
        console.print(f"  [yellow][WARN] KRX공시 수집 실패: {e}[/yellow]")
    return news_items


def _title_fingerprint(title: str) -> str:
    """제목 중복 검출용 지문 — 특수문자·공백 제거 후 앞 30자.

    같은 뉴스가 출처별로 제목이 약간 달라도 앞부분은 거의 동일하므로
    30자 슬라이스로 동일 기사를 효과적으로 탐지한다.
    """
    return re.sub(r'[^\w]', '', title)[:30]


def _title_jaccard(t1: str, t2: str) -> float:
    """두 제목의 단어 집합 Jaccard 유사도 (0~1)."""
    def _words(t):
        return set(w for w in re.sub(r'[^\w\s]', '', t).split() if len(w) > 1)
    w1, w2 = _words(t1), _words(t2)
    if not w1 or not w2:
        return 0.0
    return len(w1 & w2) / len(w1 | w2)


def collect_all_news():
    """7개 소스 병렬 수집 후 URL + 제목 유사도 기준 중복 제거.

    동일 기사가 여러 소스에 다른 URL로 올라오는 경우를 제목 지문(앞 30자)과
    Jaccard 유사도(≥0.7)로 탐지하여 하나만 남긴다.
    중복 시 먼저 수집된 기사(소스 우선순위 순)를 유지한다.
    """
    def _economy():
        items, _ = get_naver_economy_news(days=2, limit_per_day=15)
        return items

    with ThreadPoolExecutor(max_workers=7) as ex:
        fut_economy    = ex.submit(_economy)
        fut_securities = ex.submit(get_naver_securities_news, 2, 10)
        fut_rss        = [ex.submit(get_rss_news, s, u, 20) for s, u in RSS_SOURCES]
        fut_kind       = ex.submit(get_kind_disclosures, 15)

        all_batches = (
            [fut_economy.result(), fut_securities.result()]
            + [f.result() for f in fut_rss]
            + [fut_kind.result()]
        )

    seen_urls        = set()
    seen_fingerprints= set()     # 제목 앞 30자 지문
    kept_titles      = []        # Jaccard 비교용 기존 제목 목록
    all_news         = []
    dedup_count      = 0

    for batch in all_batches:
        for item in batch:
            if item['link'] in seen_urls:
                continue

            title = item.get('title', '')
            fp    = _title_fingerprint(title)

            # 1차: 지문(앞 30자) 완전 일치 → 즉시 스킵
            if fp in seen_fingerprints:
                dedup_count += 1
                continue

            # 2차: Jaccard 유사도 ≥ 0.7 → 유사 기사 스킵
            is_dup = any(_title_jaccard(title, t) >= 0.7 for t in kept_titles[-60:])
            if is_dup:
                dedup_count += 1
                continue

            seen_urls.add(item['link'])
            seen_fingerprints.add(fp)
            kept_titles.append(title)
            all_news.append(item)

    if dedup_count:
        console.print(f"  [dim]제목 중복 {dedup_count}건 제거 (URL 다른 동일 기사)[/dim]")
    return all_news


# ── 본문 추출 · JSON 파싱 ─────────────────────────────────────────────────────

def _extract_json(text):
    """Gemini 응답에서 JSON 블록 추출. 코드블록 → 전체 텍스트 순으로 시도."""
    match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', text, re.DOTALL)
    if match:
        return match.group(1)
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        return match.group(0)
    return None


_HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}

def _fetch_content(url):
    """기사 본문 추출 — 3단계 폴백.

    1) newspaper3k (7초 타임아웃, 200자 미만이면 폴백)
    2) requests + 네이버 본문 셀렉터 (#dic_area, .newsct_article, #articeBody, article)
    3) og:description 메타태그
    """
    try:
        article = Article(url, language='ko', request_timeout=7, fetch_images=False)
        article.download()
        article.parse()
        if len(article.text.strip()) > 200:
            return article.text[:1000]
    except Exception:
        pass

    try:
        resp = requests.get(url, headers=_HEADERS, timeout=10)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        body = (soup.select_one('#dic_area')
                or soup.select_one('.newsct_article')
                or soup.select_one('#articeBody')
                or soup.select_one('article'))
        if body:
            text = body.get_text(separator=' ', strip=True)
            if len(text) > 100:
                return text[:1000]
        og = soup.select_one('meta[property="og:description"]')
        if og and og.get('content', '').strip():
            return og['content'].strip()[:1000]
    except Exception:
        pass

    return ""


# ── Gemini 호출 (재시도 포함) ─────────────────────────────────────────────────

def _gemini_generate(prompt, max_retries=5, model=None, thinking=False):
    """Gemini API 호출. 503/429 등 일시 오류는 지수 백오프(최대 5회)로 재시도."""
    _model = model or GEMINI_MODEL
    config = {
        "max_output_tokens": 16384 if thinking else 2048,
        "temperature": 0.1,
        "thinking_config": {"thinking_budget": 8000} if thinking else {"thinking_budget": 0},
    }
    for attempt in range(max_retries):
        try:
            return gemini.models.generate_content(
                model=_model,
                contents=prompt,
                config=config,
            )
        except Exception as e:
            err_str = str(e)
            is_retryable = '503' in err_str or 'UNAVAILABLE' in err_str or '429' in err_str or 'RESOURCE_EXHAUSTED' in err_str
            if attempt < max_retries - 1 and is_retryable:
                wait = min(5 * (2 ** attempt), 30) + random.uniform(0, 2)
                console.print(f"[yellow][RETRY {attempt+1}/{max_retries}] {wait:.0f}초 대기 중...[/yellow]")
                time.sleep(wait)
            else:
                raise


# ── 제목 기반 사전 필터 (Gemini 호출 절감) ───────────────────────────────────

_SKIP_KEYWORDS = [
    '맛집', '레시피', '어린이날', '어버이날', '어머니날',
    '선물 추천', '할인 이벤트', '커피 전문점', '소비자 피해',
    '교통사고', '살인', '성폭력', '음주운전', '마약',
    '연예인', '아이돌', '방송 출연', '드라마', '예능',
    '스포츠', '야구', '축구', '농구', '배구', '골프',
    '날씨', '기상청', '태풍',
]

def _should_skip_by_title(title: str) -> bool:
    """제목 키워드로 투자 무관 기사를 걸러낸다 — Gemini 호출 전 선제 필터."""
    title_lower = title.lower()
    return any(kw in title_lower for kw in _SKIP_KEYWORDS)


# ── 병렬 분석용 세마포어 ─────────────────────────────────────────────────────
# Gemini: RPM 2000 기준으로 5건 동시 호출은 충분히 안전한 수준
# Ollama: OLLAMA_CONCURRENCY 환경변수로 조절 (기본 1, VRAM 여유분에 따라 2~3 가능)
_OLLAMA_CONCURRENCY = max(1, int(os.getenv("OLLAMA_CONCURRENCY", "1")))
GEMINI_CONCURRENCY = _OLLAMA_CONCURRENCY if USE_LOCAL_LLM else 5
_gemini_sem = threading.Semaphore(GEMINI_CONCURRENCY)

# 본문 로딩 동시 worker 수 (IO 바운드 — AI 분석 대기 중에도 병렬 실행)
_FETCH_WORKERS = 20


# ── Ollama 관련 ───────────────────────────────────────────────────────────────

class _OllamaResponse:
    """Gemini 응답 객체와 인터페이스를 통일하기 위한 래퍼."""
    def __init__(self, text: str):
        self.text = text


def _check_ollama() -> bool:
    """Ollama 서버 가동 여부 및 모델 존재 여부를 확인합니다."""
    try:
        resp = requests.get(f"{OLLAMA_BASE_URL}/api/tags", timeout=5)
        resp.raise_for_status()
        models = [m.get("name", "") for m in resp.json().get("models", [])]
        if not any(OLLAMA_MODEL in m for m in models):
            console.print(
                f"[yellow][WARN] Ollama에 '{OLLAMA_MODEL}' 모델이 없습니다. "
                f"→ ollama pull {OLLAMA_MODEL}[/yellow]"
            )
            return False
        return True
    except Exception as e:
        console.print(f"[yellow][WARN] Ollama 서버 미응답({OLLAMA_BASE_URL}): {e}[/yellow]")
        return False


_OLLAMA_TIMEOUT = 90   # think:False 기준 정상 응답 20~40초, 90초 초과 시 stuck으로 판단

def _ollama_generate(prompt: str, max_retries: int = 2) -> str | None:
    """Ollama /api/generate 호출. 성공 시 응답 텍스트, 실패 시 None."""
    url = f"{OLLAMA_BASE_URL}/api/generate"
    payload = {
        "model": OLLAMA_MODEL,
        "prompt": prompt,
        "stream": False,
        "format": "json",
        "think": False,  # Qwen3 thinking 모드 비활성화 (속도 향상, JSON 잘림 방지)
        "options": {"temperature": 0.1, "num_predict": 2048},
    }
    for attempt in range(max_retries):
        try:
            resp = requests.post(url, json=payload, timeout=_OLLAMA_TIMEOUT)
            resp.raise_for_status()
            return resp.json().get("response", "")
        except Exception as e:
            if attempt < max_retries - 1:
                wait = 5 * (attempt + 1)
                console.print(
                    f"[yellow][Ollama RETRY {attempt+1}/{max_retries}] {wait}초 대기: {e}[/yellow]"
                )
                time.sleep(wait)
            else:
                console.print(f"[red][Ollama] 최종 실패: {e}[/red]")
    return None


def _llm_generate(prompt, max_retries=5, model=None, thinking=False):
    """통합 LLM 호출: USE_LOCAL_LLM=true이면 Ollama 우선, 실패 시 Gemini 폴백."""
    if USE_LOCAL_LLM:
        result = _ollama_generate(prompt)
        if result is not None:
            return _OllamaResponse(result)
        if gemini is None:
            raise RuntimeError("Ollama 호출 실패 & GEMINI_API_KEY 미설정 — 폴백 불가")
        console.print("[yellow][Ollama→Gemini 폴백][/yellow]")
    return _gemini_generate(prompt, max_retries=max_retries, model=model, thinking=thinking)


# ── 개별 기사 분석 ────────────────────────────────────────────────────────────

_STOCK_NORM = re.compile(r'[\s㈜()（）주식회사\.]')

def _normalize_stock_name(name: str) -> str:
    return _STOCK_NORM.sub('', name).strip()


def _analyze_one(news, content=""):
    """단일 기사 Gemini 분석. content는 사전 prefetch된 본문."""
    content_note = content if content else "(본문 없음 — 제목만으로 분석)"

    prompt = f"""주식 투자 애널리스트로서 아래 뉴스를 분석해 호재성 종목을 발굴하세요.

[제목]: {news['title']}
[출처]: {news.get('source', '')}
[본문]: {content_note}

규칙:
- 투자관련: 기업·주식·산업·경제·정책·금융 등 넓게 해석하여 true. false는 연예·스포츠·날씨·사건사고 등 투자와 완전 무관한 경우만
- 카테고리: 특정 기업명·종목명·업종이 하나라도 언급되면 반드시 "종목관련". 거시경제·금리·환율만 다루고 특정 종목 언급이 전혀 없을 때만 "시장일반"
- 직접수혜종목: 이 뉴스가 호재인 상장사만 (악재 종목 절대 제외)
- 간접수혜종목: 공급망·경쟁구도·정책 수혜 연관 상장사 최소 2개 추론
- 공급망연관종목: 직접수혜종목의 납품업체·고객사 중 상장사
- 부정영향종목: 이 뉴스로 악재를 받을 상장사
- 폭등전조유형(해당 시): 제도정책빅뱅/독점최초계약/실적서프라이즈/대규모지분변동/신사업기술돌파/수급변화/M&A합병/CEO·VIP방한
  * CEO·VIP방한: 젠슨황(엔비디아)·팀쿡(애플)·사티아나델라(MS) 등 글로벌 빅테크·반도체 CEO, 또는 주요국 정상 방한으로 특정 기업·산업에 직접 호재가 되는 방문 이벤트
- 가치주신호 true: 적자→흑자, 저PER+실적급증, 목표가 대폭 상향

JSON으로만 응답 (한국 상장사만, ETF 제외):
{{
    "투자관련": true/false,
    "카테고리": "종목관련 또는 시장일반",
    "핵심이슈": "한 문장",
    "뉴스방향성": "긍정/부정/중립",
    "직접수혜종목": ["종목명1", "종목명2"],
    "간접수혜종목": ["종목명1", "종목명2"],
    "공급망연관종목": ["종목명1"],
    "부정영향종목": ["종목명1"],
    "숨은수혜주_여부": true/false,
    "가치주신호": true/false,
    "업종": "",
    "투자호재성격": "제도정책/독점최초/실적서프라이즈/지분변동/신사업기술/수급변화/M&A/CEO·VIP방한/기타",
    "폭등전조유형": "",
    "단기폭발력": 1~10,
    "매출증가율": null,
    "영업익증가율": null,
    "주요지표": "",
    "인사이트": "2~3문장",
    "중요도": 1~10
}}"""
    raw = None
    for _attempt in range(2):
        response = _llm_generate(prompt, model=GEMINI_MODEL_LITE, thinking=False)
        raw = _extract_json(response.text)
        if raw:
            break
        if _attempt == 0:
            console.print(f"  [yellow][JSON 재시도][/yellow] {news['title'][:35]}")
    if not raw:
        raise ValueError("JSON 블록을 찾을 수 없습니다.")

    analysis = json.loads(raw)
    analysis['중요도'] = int(analysis.get('중요도', 0))
    analysis['원본링크'] = news['link']
    analysis['뉴스제목'] = news['title']
    analysis['출처'] = news.get('source', '')

    # 하위 호환: 직접수혜종목을 기존 관련종목 필드로도 노출
    analysis['관련종목'] = (
        analysis.get('직접수혜종목', []) + analysis.get('간접수혜종목', [])
    )
    return analysis


# ── 합성 분석 ─────────────────────────────────────────────────────────────────

def _synthesize_insights(results):
    """전체 개별 분석 결과 → 섹터/테마/주목종목/악재경고 종합 인사이트 (Gemini 합성, 1회 호출)."""
    summary_data = [
        {
            "제목":           r.get('뉴스제목', '')[:50],
            "직접수혜종목":   r.get('직접수혜종목', []),
            "간접수혜종목":   r.get('간접수혜종목', []),
            "공급망연관종목": r.get('공급망연관종목', []),
            "부정영향종목":   r.get('부정영향종목', []),
            "폭등전조유형":   r.get('폭등전조유형', ''),
            "단기폭발력":     r.get('단기폭발력', 0),
            "가치주신호":     r.get('가치주신호', False),
            "중요도":         r.get('중요도', 0),
        }
        for r in results
    ]

    prompt = f"""
당신은 주식 투자 전문 애널리스트입니다.
아래는 최근 36시간 이내 경제·증권 뉴스 {len(results)}건의 AI 분석 결과입니다.
이를 종합하여 **호재성 기사 중심의** 투자 인사이트를 도출하세요.

[분석 데이터]
{json.dumps(summary_data, ensure_ascii=False, indent=2)}

## 🚨 핵심 지침 🚨

### A. 유망 섹터/테마 (최대 5개)
- **구조적 변화의 초기 징후** 테마를 우선 발굴하세요.
- 폭등전조유형이 있거나, 복수 기사에서 공통적으로 호재 신호가 나타난 테마에 주목하세요.

### B. 주목 종목 — 호재성 기사 기반 (최대 10개)
우선순위:
1. **폭등 후보**: 단기폭발력 7점 이상 OR 폭등전조유형 있는 직접수혜종목 (최우선)
2. **공급망 연관주**: 직접수혜종목의 핵심 납품업체·고객사 (공급망연관종목)
3. **간접 수혜주**: 동일 공급망·정책·테마로 올라갈 수 있는 종목 (간접수혜종목)
4. **가치주**: 가치주신호=true인 종목
5. 복수 기사 언급 + 고중요도 호재 종목
- **부정영향종목에 올라온 종목은 주목_종목에서 제외하세요.**
- 언급 횟수가 1회여도 단기폭발력이 극히 높으면 최상위 배치.

### C. 악재 경고 종목 (최대 5개)
- 부정영향종목으로 복수 기사에서 언급된 종목을 정리하세요.

### D. 한국 증시(KOSPI/KOSDAQ) 상장사 기준

다음 JSON 형식으로만 응답하세요:
{{
    "시장_방향성_요약": "현재 시장 전반 흐름 2~3문장 요약",
    "유망_섹터_테마": [
        {{
            "섹터_테마명": "섹터 또는 테마명",
            "근거": "왜 이 테마에 지금 돈이 몰릴 수 있는지(폭발력) 설명",
            "핵심_촉매": "이 섹터/테마를 움직이는 핵심 이슈",
            "직접수혜종목": ["종목1", "종목2"],
            "간접수혜종목": ["연관 공급망·테마 종목"],
            "중요도": 1~10
        }}
    ],
    "주목_종목": [
        {{
            "종목명": "종목명",
            "발굴유형": "폭등후보/공급망연관/간접수혜/가치주/복수언급 중 택1",
            "호재_근거": "구체적인 호재 내용 (어떤 뉴스의 어떤 내용이 왜 호재인지)",
            "언급횟수": 언급된 기사 수,
            "공급망_위치": "직접수혜 / 1차공급망 / 간접수혜 / 정책수혜 중 택1",
            "리스크": "단기 리스크 요인"
        }}
    ],
    "악재_경고_종목": [
        {{
            "종목명": "종목명",
            "악재_내용": "구체적인 악재 내용",
            "영향_강도": 1~10
        }}
    ],
    "투자_주의사항": "현재 시장에서 주의해야 할 리스크 요인 1~2문장"
}}
"""
    if GEMINI_SYNTHESIS and gemini is not None:
        response = _gemini_generate(prompt, thinking=True)
    else:
        response = _llm_generate(prompt, thinking=True)
    raw = _extract_json(response.text)
    if not raw:
        raise ValueError("합성 분석 JSON을 찾을 수 없습니다.")
    return json.loads(raw)


def _build_insight_dfs(synthesis):
    """합성 결과 dict → 4개 DataFrame (시장방향성, 섹터테마, 주목종목, 악재경고)."""
    df_summary = pd.DataFrame([{
        "시장 방향성 요약": synthesis.get('시장_방향성_요약', ''),
        "투자 주의사항":    synthesis.get('투자_주의사항', ''),
    }])

    sector_rows = [
        {
            "순위":       i,
            "섹터/테마":  item.get('섹터_테마명', ''),
            "투자 근거":  item.get('근거', ''),
            "핵심 촉매":  item.get('핵심_촉매', ''),
            "직접수혜종목": ", ".join(item.get('직접수혜종목', item.get('관련_종목', []))),
            "간접수혜종목": ", ".join(item.get('간접수혜종목', [])),
            "중요도":     item.get('중요도', 0),
        }
        for i, item in enumerate(synthesis.get('유망_섹터_테마', []), 1)
    ]
    df_sectors = pd.DataFrame(sector_rows) if sector_rows else pd.DataFrame(
        columns=["순위", "섹터/테마", "투자 근거", "핵심 촉매", "직접수혜종목", "간접수혜종목", "중요도"]
    )

    watch_rows = [
        {
            "종목명":        item.get('종목명', ''),
            "발굴유형":      item.get('발굴유형', ''),
            "공급망 위치":   item.get('공급망_위치', ''),
            "호재 근거":     item.get('호재_근거', item.get('이유', '')),
            "기사 언급횟수": item.get('언급횟수', 0),
            "리스크":        item.get('리스크', ''),
        }
        for item in synthesis.get('주목_종목', [])
    ]
    df_watch = pd.DataFrame(watch_rows) if watch_rows else pd.DataFrame(
        columns=["종목명", "발굴유형", "공급망 위치", "호재 근거", "기사 언급횟수", "리스크"]
    )

    caution_rows = [
        {
            "종목명":    item.get('종목명', ''),
            "악재 내용": item.get('악재_내용', ''),
            "영향 강도": item.get('영향_강도', 0),
        }
        for item in synthesis.get('악재_경고_종목', [])
    ]
    df_caution = pd.DataFrame(caution_rows) if caution_rows else pd.DataFrame(
        columns=["종목명", "악재 내용", "영향 강도"]
    )

    return df_summary, df_sectors, df_watch, df_caution


# ── 종목뉴스·시장일반 DataFrame ───────────────────────────────────────────────

_LIST_COLS = {'관련종목', '직접수혜종목', '간접수혜종목', '공급망연관종목', '부정영향종목'}

def _build_df(rows, col_order):
    if not rows:
        return pd.DataFrame(columns=col_order)
    df = pd.DataFrame(rows)
    for col in _LIST_COLS:
        if col in df.columns:
            df[col] = df[col].apply(
                lambda x: ", ".join(x) if isinstance(x, list) else str(x)
            )
    df['중요도'] = pd.to_numeric(df['중요도'], errors='coerce').fillna(0).astype(int)
    if '단기폭발력' in df.columns:
        df['단기폭발력'] = pd.to_numeric(df['단기폭발력'], errors='coerce').fillna(0).astype(int)
    # 발굴점수: 단기폭발력(×2 가중) + 중요도 → 폭등 전조 종목이 최상단에 노출
    if '단기폭발력' in df.columns:
        df['발굴점수'] = df['단기폭발력'] * 2 + df['중요도']
        sort_cols = ['발굴점수', '중요도']
    else:
        df['발굴점수'] = df['중요도']
        sort_cols = ['중요도']
    df = df.sort_values(by=sort_cols, ascending=False).reset_index(drop=True)
    final_cols = [c for c in col_order if c in df.columns]
    return df[final_cols]


# ── 종합/종합(분기포함) DataFrame ─────────────────────────────────────────────

def _load_earnings_df() -> tuple[pd.DataFrame, pd.DataFrame]:
    """4_financial_scanner output 최신 파일의 '종합', '종합(분기포함)' 시트를 로드합니다."""
    import glob as _glob
    base  = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '4_financial_scanner', 'output')
    files = sorted([f for f in _glob.glob(os.path.join(base, '*.xlsx')) if '_hl' not in f])
    if not files:
        return pd.DataFrame(), pd.DataFrame()
    try:
        df_brief = pd.read_excel(files[-1], sheet_name="종합", dtype=str)
        df_brief.columns = [str(c).replace('\n', '') for c in df_brief.columns]
        df_full  = pd.read_excel(files[-1], sheet_name="종합(분기포함)", dtype=str)
        df_full.columns  = [str(c).replace('\n', '') for c in df_full.columns]
        return df_brief, df_full
    except Exception:
        return pd.DataFrame(), pd.DataFrame()


def _build_earnings_sheets(stocks_all: list) -> tuple[pd.DataFrame, pd.DataFrame]:
    """직접수혜종목 × 4_financial_scanner '종합'/'종합(분기포함)' 시트 매칭."""
    stock_news: dict[str, list[str]] = {}
    for r in stocks_all:
        for name in r.get('직접수혜종목', []):
            name = name.strip()
            if name and name != 'nan':
                stock_news.setdefault(name, []).append(r.get('뉴스제목', ''))

    if not stock_news:
        return pd.DataFrame(), pd.DataFrame()

    df_brief_src, df_full_src = _load_earnings_df()
    if df_brief_src.empty:
        return pd.DataFrame(), pd.DataFrame()

    # 종목명 → 티커 맵 (정규화 포함)
    name_to_ticker: dict[str, str] = {}
    if '종목명' in df_brief_src.columns and '티커' in df_brief_src.columns:
        for _, row in df_brief_src.iterrows():
            raw    = str(row['종목명']).strip()
            ticker = str(row['티커']).strip()
            if raw and ticker and ticker != 'nan':
                name_to_ticker[raw] = ticker
                name_to_ticker[_normalize_stock_name(raw)] = ticker

    # 티커 → 관련뉴스 맵
    ticker_news: dict[str, list[str]] = {}
    for name, titles in stock_news.items():
        ticker = (name_to_ticker.get(name)
                  or name_to_ticker.get(_normalize_stock_name(name)))
        if ticker:
            prev = ticker_news.setdefault(ticker, [])
            for t in titles:
                if t not in prev:
                    prev.append(t)

    if not ticker_news:
        return pd.DataFrame(), pd.DataFrame()

    def _filter(src: pd.DataFrame) -> pd.DataFrame:
        if src.empty or '티커' not in src.columns:
            return pd.DataFrame()
        matched = src[src['티커'].isin(ticker_news)].copy()
        if matched.empty:
            return pd.DataFrame()
        matched.insert(0, '관련뉴스', matched['티커'].map(
            lambda t: " / ".join(ticker_news.get(t, [])[:3])))
        return matched.reset_index(drop=True)

    return _filter(df_brief_src), _filter(df_full_src)


KST = timezone(timedelta(hours=9))


def _send_email(output_file: str, stocks_all: list, df_watch, df_caution) -> None:
    """분석 결과 Excel을 첨부해 수신자 목록으로 이메일 전송."""
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from pathlib import Path

    ts = datetime.now(KST).strftime("%Y-%m-%d %H:%M")

    # 주목 종목 목록
    watch_lines = []
    if df_watch is not None and not df_watch.empty:
        for _, row in df_watch.head(10).iterrows():
            name  = str(row.get('종목명', ''))
            gtype = str(row.get('발굴유형', ''))
            basis = str(row.get('호재 근거', ''))[:50]
            if name:
                watch_lines.append(f"  [{gtype}] {name} — {basis}")

    # 악재 경고 목록
    caution_lines = []
    if df_caution is not None and not df_caution.empty:
        for _, row in df_caution.iterrows():
            name    = str(row.get('종목명', ''))
            content = str(row.get('악재 내용', ''))[:50]
            if name:
                caution_lines.append(f"  {name} — {content}")

    # 집계
    all_direct: set[str] = set()
    for r in stocks_all:
        for n in r.get('직접수혜종목', []):
            if n and n != 'nan':
                all_direct.add(n.strip())
    n_total   = len(stocks_all)
    n_direct  = len(all_direct)
    n_watch   = len(df_watch)   if df_watch   is not None else 0
    n_caution = len(df_caution) if df_caution is not None else 0

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
        f"[{ts}] 뉴스기사 인사이트 분석이 완료되었습니다.",
        f"",
        f"{'='*60}",
        f"■ 분석 결과 요약",
        f"{'='*60}",
        f"  분석 기사 수   : {n_total}건",
        f"  직접수혜 종목  : {n_direct}개 (중복 제외)",
        f"  주목 종목      : {n_watch}개",
        f"  악재 경고 종목 : {n_caution}개",
        f"",
        f"■ 주목 종목",
        *(watch_lines if watch_lines else ["  (주목 종목 없음)"]),
        f"",
        f"■ 악재 경고 종목",
        *(caution_lines if caution_lines else ["  (악재 경고 없음)"]),
        f"",
        f"{'='*60}",
        f"■ 첨부 엑셀 구성 (4개 시트)",
        f"{'='*60}",
        f"  [종목뉴스]",
        f"    호재성 기사 목록 (발굴점수 = 단기폭발력×2 + 중요도 순)",
        f"    컬럼: 뉴스제목 / 출처 / 핵심이슈 / 뉴스방향성",
        f"          직접수혜종목 / 간접수혜종목 / 공급망연관종목 / 부정영향종목",
        f"          투자호재성격 / 폭등전조유형 / 단기폭발력 / 인사이트 / 중요도",
        f"    · 폭등전조유형 있음 → 주황 강조",
        f"    · 단기폭발력 7↑ 연초록 / 4~6 연노랑",
        f"    · 가치주신호 true → 연파랑",
        f"",
        f"  [투자인사이트]",
        f"    합성 분석 결과 — 4개 섹션",
        f"    섹션A 시장 방향성 + 투자 주의사항",
        f"    섹션B 유망 섹터·테마 Top5 (직접수혜·간접수혜 종목 분리, 중요도 7↑ 연초록)",
        f"    섹션C 주목 종목 Top10 (폭등후보→공급망연관→간접수혜→가치주→복수언급)",
        f"    섹션D 악재 경고 종목 (영향강도 7↑ 주황 강조)",
        f"",
        f"  [시장일반]",
        f"    특정 종목 없이 금리·환율·정책·거시경제 흐름을 다루는 기사",
        f"",
        f"  [종합]",
        f"    직접수혜종목 중 4_financial_scanner 데이터와 매칭된 종목 — 분기 컬럼 제외",
        f"    컬럼: 관련뉴스 / 종목명 / 티커 / 업종 / 종합점수",
        f"          현재가 / 적정주가 / 상승여력 / 매출·영업익 증가율(25A→26E)",
        f"          PER / 12M PER / PSR / PBR / PFR",
        f"    · 상승여력2 ≥ 20%  → 연초록",
        f"    · PER / 12M PER < 12 → 연파랑",
        f"    · 영업익증가율 ≥ 30% → 연노랑",
        f"",
        f"  [종합(분기포함)]",
        f"    종합 시트와 동일하되 분기별 매출·영업익 YoY 컬럼 추가",
        f"",
        f"{'='*60}",
        f"■ 뉴스 수집 소스 (7개)",
        f"{'='*60}",
        f"  네이버경제 / 네이버증권 / 한국경제 / 연합뉴스 / 매일경제 / 서울경제 / KRX공시",
        f"",
        f"■ 동작 방식",
        f"  최근 36시간 이내 뉴스 중복 제거 후 약 100건 수집 (7개 소스 병렬)",
        f"  기사 본문 병렬 추출 (10건 동시) — newspaper3k → requests → og:description",
        f"  Gemini AI (gemini-2.5-flash) → 호재/악재 종목 분리·직접/공급망/간접 체인 발굴 (5건 병렬)",
        f"  전체 결과 합성 분석 1회 → 유망 섹터·주목 종목·악재 경고 종합",
        f"  4_financial_scanner 최신 데이터와 직접수혜종목 티커 매칭 → 재무제표 시트 생성",
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
    msg["Subject"] = f"[뉴스 인사이트] {ts} 분석 완료 — 주목 {n_watch}개 / 직접수혜 {n_direct}개"
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

if __name__ == "__main__":
    _start_time = datetime.now()
    console.print()
    console.print("════════════════════════════════════════════════════════════════")
    console.print("  [bold cyan]뉴스기사 인사이트 수집기[/bold cyan] — 호재성 종목 · 관련 종목 발굴")
    console.print("════════════════════════════════════════════════════════════════")

    if USE_LOCAL_LLM:
        console.print(f"  [cyan]AI 엔진[/cyan]: Local LLM ({OLLAMA_MODEL}) @ {OLLAMA_BASE_URL}")
        fallback_note = " + Gemini 폴백" if gemini else " (Gemini 폴백 미설정)"
        console.print(f"  [dim]동시 처리[/dim]: {GEMINI_CONCURRENCY}건{fallback_note}")
        synth_engine = "Gemini (thinking)" if (GEMINI_SYNTHESIS and gemini) else "Ollama"
        console.print(f"  [dim]합성 분석[/dim]: {synth_engine}")
        if not _check_ollama():
            console.print("[red]Ollama를 먼저 실행하세요: ollama serve[/red]")
            sys.exit(1)
    else:
        console.print(f"  [cyan]AI 엔진[/cyan]: Gemini ({GEMINI_MODEL})")
        console.print(f"  [dim]동시 처리[/dim]: {GEMINI_CONCURRENCY}건")
    console.print()

    # ── 1단계: 뉴스 수집 ──────────────────────────────────────────────────────
    console.print("[bold]▶ 뉴스 수집 중...[/bold]")
    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=20, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        console=console,
        transient=True
    ) as p:
        t = p.add_task("수집 중...", total=None)
        raw_news = collect_all_news()

    if not raw_news:
        console.print("[red]수집된 뉴스가 없습니다.[/red]")
        sys.exit(1)

    # 소스별 집계 출력
    src_counts = Counter(n['source'] for n in raw_news)
    src_summary = "  ".join(f"[dim]{s}[/dim] {c}건" for s, c in src_counts.items())
    console.print(f"[green]✔[/green] 총 {len(raw_news)}건 수집  ({src_summary})\n")

    # ── 2단계: 제목 기반 사전 필터링 (fetch 전에 스킵) ──────────────────────────
    pre_filtered = []
    title_skipped = 0
    for i, news in enumerate(raw_news):
        if _should_skip_by_title(news['title']):
            title_skipped += 1
        else:
            pre_filtered.append((i, news))
    if title_skipped:
        console.print(f"  [dim]제목 필터링으로 {title_skipped}건 사전 스킵[/dim]")

    # ── 3단계: 본문 로딩 + AI 분석 파이프라인 ────────────────────────────────────
    # 본문 로딩 완료 즉시 AI 분석 시작 — 두 단계를 겹쳐서 전체 시간 단축
    # • _FETCH_WORKERS(20)개가 IO 바운드 본문 로딩을 병렬 수행
    # • _gemini_sem이 AI 호출을 GEMINI_CONCURRENCY 건으로 제한
    # • AI 슬롯 대기 중에도 fetch는 계속 진행되어 항상 다음 기사 준비 완료
    console.print(
        f"[bold]▶ 기사 분석 중... "
        f"({len(pre_filtered)}건 | fetch {_FETCH_WORKERS}건 병렬 · AI {GEMINI_CONCURRENCY}건 병렬)[/bold]"
    )
    final_insights = []
    failed         = []
    _lock          = threading.Lock()
    completed      = [0]

    _ARTICLE_TIMEOUT = 120  # 기사 1개당 최대 120초 — 초과 시 강제 스킵

    def _analyze_task(idx, news, content=""):
        """세마포어로 AI 동시 호출 수 제어. 기사당 타임아웃 적용."""
        with _gemini_sem:
            result_box = [None]
            error_box  = [None]

            def _worker():
                try:
                    result_box[0] = _analyze_one(news, content=content)
                except Exception as exc:
                    error_box[0] = exc

            t = threading.Thread(target=_worker, daemon=True)
            t.start()
            t.join(timeout=_ARTICLE_TIMEOUT)

            if t.is_alive():
                with _lock:
                    failed.append({"title": news['title'], "error": f"타임아웃 ({_ARTICLE_TIMEOUT}초 초과) 스킵"})
                    completed[0] += 1
                    _done = completed[0]
                console.print(f"  [yellow]SKIP[/yellow] 타임아웃: {news['title'][:40]}")
                sys.stdout.write(f"PROGRESS:{_done}/{total}\n")
                sys.stdout.flush()
                return

            try:
                if error_box[0]:
                    raise error_box[0]
                analysis = result_box[0]
                with _lock:
                    if not analysis.get('투자관련', True):
                        failed.append({"title": news['title'], "error": "투자무관 스킵"})
                        console.print(f"  [dim]SKIP 투자무관: {news['title'][:45]}[/dim]")
                    elif analysis.get('중요도', 0) < 3 and int(analysis.get('단기폭발력', 0)) < 5:
                        failed.append({"title": news['title'], "error": "중요도 낮음 스킵"})
                        console.print(f"  [dim]SKIP 중요도낮음: {news['title'][:43]}[/dim]")
                    else:
                        final_insights.append(analysis)
                        console.print(f"  [green]OK[/green]  {news['title'][:50]}")
            except Exception as e:
                err_msg = str(e)
                with _lock:
                    failed.append({"title": news['title'], "error": err_msg})
                console.print(f"  [red]ERR[/red] {err_msg[:120]}")
            finally:
                with _lock:
                    completed[0] += 1
                    _done = completed[0]
                sys.stdout.write(f"PROGRESS:{_done}/{total}\n")
                sys.stdout.flush()

    def _fetch_and_analyze(idx, news):
        """본문 로딩 완료 즉시 AI 분석 투입 (파이프라인 단위 작업)."""
        content = _fetch_content(news['link'])
        _analyze_task(idx, news, content=content)

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=34, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TextColumn("남은"),
        TimeRemainingColumn(),
        console=console,
        transient=True,
    ) as progress:
        total = len(pre_filtered)
        t_analyze = progress.add_task("본문로딩+AI분석 중...", total=total)

        with ThreadPoolExecutor(max_workers=_FETCH_WORKERS) as ex:
            futures = {ex.submit(_fetch_and_analyze, idx, news): news for idx, news in pre_filtered}
            for fut in as_completed(futures):
                fut.result()
                done = completed[0]
                progress.update(t_analyze, completed=done,
                                description=f"[{done}/{total}] 로딩+분석 중...")
                if done % 5 == 0 or done == total:
                    ok  = len(final_insights)
                    sk  = len([f for f in failed if "스킵" in f["error"]])
                    err = len([f for f in failed if "스킵" not in f["error"]])
                    console.print(
                        f"  ├ [{done}/{total}] 분석 진행 — "
                        f"채택 {ok}건 / 스킵 {sk}건 / 오류 {err}건"
                    )

    CAT_STYLE = {"종목관련": "green", "시장일반": "yellow"}
    for r in final_insights:
        cat   = r.get('카테고리', '?')
        color = CAT_STYLE.get(cat, "white")
        src   = f"[dim][{r.get('출처','')}][/dim] " if r.get('출처') else ""
        console.print(f"  [{color}]{cat:6}[/{color}]  {src}{r['뉴스제목'][:38]}")
    for f in [x for x in failed if '스킵' in x['error']]:
        console.print(f"  [dim]스킵  ({f['error']})[/dim]  {f['title'][:35]}")
    for f in [x for x in failed if '스킵' not in x['error']]:
        console.print(f"  [red]실패  [/red]  {f['title'][:40]}")

    if not final_insights:
        console.print("\n[red]분석 결과가 없습니다.[/red]")
        sys.exit(1)

    # ── 4단계: 합성 분석 ──────────────────────────────────────────────────────
    console.print(f"\n[bold]▶ 섹터/테마/종목 종합 인사이트 합성 중...[/bold]")
    df_summary = df_sectors = df_watch = df_caution = None
    synthesis_warn = None

    with Progress(SpinnerColumn(), TextColumn("{task.description}"),
                  console=console, transient=True) as p:
        p.add_task("Gemini 합성 분석 중...", total=None)
        try:
            synthesis = _synthesize_insights(final_insights)
            df_summary, df_sectors, df_watch, df_caution = _build_insight_dfs(synthesis)
            console.print(f"[green]✔[/green] 합성 완료 — "
                          f"섹터/테마 {len(df_sectors)}개  주목종목 {len(df_watch)}개  "
                          f"악재경고 {len(df_caution)}개")
        except Exception as e:
            synthesis_warn = str(e)
            df_summary  = pd.DataFrame([{"시장 방향성 요약": "합성 분석 실패", "투자 주의사항": synthesis_warn}])
            df_sectors  = pd.DataFrame(columns=["순위", "섹터/테마", "투자 근거", "핵심 촉매", "직접수혜종목", "간접수혜종목", "중요도"])
            df_watch    = pd.DataFrame(columns=["종목명", "발굴유형", "공급망 위치", "호재 근거", "기사 언급횟수", "리스크"])
            df_caution  = pd.DataFrame(columns=["종목명", "악재 내용", "영향 강도"])
            console.print(f"[yellow][WARN] 합성 분석 실패: {synthesis_warn}[/yellow]")

    # ── 5단계: DataFrame 구성 ─────────────────────────────────────────────────
    stocks_all = [r for r in final_insights if r.get('카테고리') == '종목관련']
    market     = [r for r in final_insights if r.get('카테고리') == '시장일반']

    df_stocks        = _build_df(stocks_all, ["뉴스제목", "출처", "핵심이슈", "뉴스방향성", "직접수혜종목", "간접수혜종목", "공급망연관종목", "부정영향종목", "투자호재성격", "폭등전조유형", "단기폭발력", "숨은수혜주_여부", "가치주신호", "인사이트", "중요도", "원본링크"])
    df_market        = _build_df(market,     ["뉴스제목", "출처", "핵심이슈", "뉴스방향성", "투자호재성격", "단기폭발력", "인사이트", "중요도", "원본링크"])
    df_brief, df_full = _build_earnings_sheets(stocks_all)

    # ── 6단계: 저장 ───────────────────────────────────────────────────────────
    output_dir  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
    os.makedirs(output_dir, exist_ok=True)
    timestamp   = datetime.now().strftime("%Y_%m%d_%H%M")
    model_tag   = f"_{OLLAMA_MODEL.replace(':', '_')}" if USE_LOCAL_LLM else ""
    suffix      = f"_LocalLLM{model_tag}" if USE_LOCAL_LLM else ""
    output_file = os.path.join(output_dir, f"뉴스기사_인사이트_{timestamp}{suffix}.xlsx")

    with Progress(SpinnerColumn(), TextColumn("{task.description}"),
                  console=console, transient=True) as p:
        p.add_task("엑셀 저장 및 서식 적용 중...", total=None)

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            # 1. 종목뉴스 (사용자 요청으로 투자인사이트 왼쪽으로 배치)
            df_stocks.to_excel(writer,  sheet_name=SHEET_STOCKS,  index=False)
            
            # 2. 투자인사이트 (4개 섹션 통합)
            row_cursor = 0
            df_summary.to_excel(writer, sheet_name=SHEET_INSIGHT, index=False, startrow=row_cursor)
            row_cursor += len(df_summary) + 2
            df_sectors.to_excel(writer, sheet_name=SHEET_INSIGHT, index=False, startrow=row_cursor)
            row_cursor += len(df_sectors) + 2
            df_watch.to_excel(writer,   sheet_name=SHEET_INSIGHT, index=False, startrow=row_cursor)
            row_cursor += len(df_watch) + 2
            df_caution.to_excel(writer, sheet_name=SHEET_INSIGHT, index=False, startrow=row_cursor)

            # 3. 기타 시트
            df_market.to_excel(writer, sheet_name=SHEET_MARKET, index=False)
            if not df_brief.empty:
                df_brief.to_excel(writer, sheet_name=SHEET_BRIEF, index=False)
            if not df_full.empty:
                df_full.to_excel(writer,  sheet_name=SHEET_FULL,  index=False)

        # ── 7단계: 서식 적용 ──────────────────────────────────────────────────
        fmt_warn = None
        try:
            spec = importlib.util.spec_from_file_location(
                "format_excel",
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "format_excel.py")
            )
            mod = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(mod)
            mod.format_file(output_file,
                            insight_section_rows=(len(df_summary), len(df_sectors), len(df_watch), len(df_caution)))
            fmt_path = os.path.splitext(output_file)[0] + "_fmt.xlsx"
            shutil.move(fmt_path, output_file)
        except Exception as e:
            fmt_warn = str(e)

    # ── 완료 요약 ────────────────────────────────────────────────────────────
    console.print(f"\n[bold green]✔ 분석 완료[/bold green]")
    console.print(f"{'='*60}")
    console.print(
        f"  종목뉴스 [green]{len(stocks_all)}건[/green]  |  "
        f"시장일반 [yellow]{len(market)}건[/yellow]  |  "
        f"종합 [cyan]{len(df_brief)}건[/cyan]\n"
        f"  주목종목 [magenta]{len(df_watch)}개[/magenta]  |  "
        f"악재경고 [red]{len(df_caution)}개[/red]  |  "
        f"실패/스킵 [dim]{len(failed)}건[/dim]"
    )
    console.print(f"{'='*60}")
    console.print(f"  저장 경로: [bold]{output_file}[/bold]\n")
    if fmt_warn:
        console.print(f"[yellow][WARN] 서식 적용 실패: {fmt_warn}[/yellow]\n")

    _send_email(output_file, stocks_all, df_watch, df_caution)

    _end_time = datetime.now()
    _elapsed  = _end_time - _start_time
    _mins, _secs = divmod(int(_elapsed.total_seconds()), 60)
    console.print(
        f"[bold green]⏱ 시작 {_start_time:%H:%M:%S} → 종료 {_end_time:%H:%M:%S}  "
        f"(총 {_mins}분 {_secs}초)[/bold green]\n"
    )
