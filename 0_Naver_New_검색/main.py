# naver_news_selenium.py (0914 - URL+텍스트+본문메타 3단 보정판)
#
# 네이버 뉴스 키워드 검색기 (최신순 1페이지)
#
# [흐름]
#   1. 키워드 입력 → Selenium으로 네이버 뉴스 최신순 1페이지 수집
#   2. 구조 자동 감지: 구 렌더(a.news_tit) / 신 렌더(span.sds-comps-text-type-headline1)
#   3. 날짜 3단계 보정: URL 패턴 → 카드 텍스트 토큰 → 기사 본문 메타태그
#   4. 터미널 출력 + {키워드}_1page.xlsx 저장
#
# [주요 함수]
#   build_driver()         : headless Chrome 드라이버 생성
#   fetch_news()           : Selenium으로 뉴스 카드(제목·날짜텍스트·링크) 수집
#   ymd_best()             : URL·텍스트·본문 순으로 최적 날짜(YYYY-MM-DD) 결정
#   dedupe_news()          : URL 기준 중복 제거
#   print_news_to_console(): 컬러 터미널 출력

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from urllib.parse import quote_plus
from bs4 import BeautifulSoup
import pandas as pd
import re
import time
import os
import shutil
import textwrap
import sys
import requests
from datetime import datetime, timedelta, timezone

# (선택) 컬러 출력
try:
    from colorama import init as colorama_init, Fore, Style
    colorama_init(autoreset=True)
    HAS_COLOR = True
except Exception:
    HAS_COLOR = False

# ---- 전역 설정 ----
KST = timezone(timedelta(hours=9))
HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/124.0.0.0 Safari/537.36"),
    "Accept-Language": "ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7",
}

# 날짜 토큰(부분 매치) 정규식: 그룹을 나눠 우선순위 판단
DATE_TOKEN_RE = re.compile(
    r"(?:"
    r"(?P<ymd>\d{4}\s*[.\-]\s*\d{1,2}\s*[.\-]\s*\d{1,2})"                     # 2025. 9. 12 / 2025-09-12
    r"|(?P<rel>\d+\s*분\s*전|\d+\s*시간\s*전|\d+\s*일\s*전|\d+\s*주\s*전|어제|오늘|방금)"  # 상대표현
    r"|(?P<md>(?<!\d)(?:\d{1,2})\s*[.\-]\s*(?:\d{1,2})(?!\d|%))"               # MM.DD (바로 뒤에 % 금지)
    r")"
)

# 날짜 메타가 있을 법한 컨테이너(카드 내부만 탐색)
DATE_CONTAINER_SELECTORS = [".info_group", ".news_info", "span.info", "div.info", "time"]

def clean_filename(s: str) -> str:
    return re.sub(r'[\\/*?:"<>|]', "", s)

def norm_text(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").strip())

def build_driver(headless: bool = True) -> webdriver.Chrome:
    """항상 headless 기본값 True → 창이 보이지 않음"""
    opts = Options()
    if headless:
        opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    opts.add_argument("--disable-blink-features=AutomationControlled")
    opts.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    opts.add_experimental_option("useAutomationExtension", False)
    opts.add_argument("--log-level=3")
    opts.add_argument("accept-language=ko-KR,ko;q=0.9,en-US;q=0.8,en;q=0.7")
    opts.add_argument(f"--user-agent={HEADERS['User-Agent']}")
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=opts)

def print_step(msg: str):
    if HAS_COLOR:
        print(f"{Fore.CYAN}▶{Style.RESET_ALL} {msg}")
    else:
        print(f"▶ {msg}")

def print_progress(i: int, total: int, label: str = "진행"):
    width = 28
    ratio = 0 if total == 0 else i / total
    filled = int(ratio * width)
    bar = "█" * filled + "-" * (width - filled)
    pct = int(ratio * 100)
    sys.stdout.write(f"\r{label} [{bar}] {pct:3d}% ({i}/{total})")
    sys.stdout.flush()
    if i == total:
        sys.stdout.write("\n")
        sys.stdout.flush()

def dedupe_news(rows):
    seen = set()
    uniq = []
    for title, date_text, link in rows:
        key = (link.strip().lower() if link else "") or ("t:" + norm_text(title).lower())
        if key in seen:
            continue
        seen.add(key)
        uniq.append([title, date_text, link])
    return uniq

# ---------- 날짜 해석 ----------
def _strip_time_tail(s: str) -> str:
    # '오전 3:10' 같은 시간 꼬리 제거(날짜만 사용)
    return re.sub(r"(오전|오후)\s*\d{1,2}:\d{2}", "", s or "").strip()

def _parse_token_to_dt(token: str) -> datetime | None:
    token = token.strip()
    now = datetime.now(KST)
    if token in ("방금",):
        return now
    if token == "오늘":
        return now.replace(hour=0, minute=0, second=0, microsecond=0)
    if token == "어제":
        d = now - timedelta(days=1)
        return d.replace(hour=0, minute=0, second=0, microsecond=0)

    m = re.fullmatch(r"(\d+)\s*분\s*전", token)
    if m: return now - timedelta(minutes=int(m.group(1)))
    m = re.fullmatch(r"(\d+)\s*시간\s*전", token)
    if m: return now - timedelta(hours=int(m.group(1)))
    m = re.fullmatch(r"(\d+)\s*일\s*전", token)
    if m: return now - timedelta(days=int(m.group(1)))
    m = re.fullmatch(r"(\d+)\s*주\s*전", token)
    if m: return now - timedelta(weeks=int(m.group(1)))

    m = re.fullmatch(r"(\d{4})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})", token)
    if m:
        y, mo, d = map(int, m.groups())
        return datetime(y, mo, d, tzinfo=KST)

    m = re.fullmatch(r"(\d{1,2})\s*[.\-]\s*(\d{1,2})", token)
    if m:
        mo, d = map(int, m.groups())
        y = now.year
        try:
            dt = datetime(y, mo, d, tzinfo=KST)
        except ValueError:
            return None
        if (dt - now).days > 7:
            dt = dt.replace(year=y - 1)
        return dt

    return None

def _format_dt_ymd(dt: datetime) -> str:
    if isinstance(dt, datetime):
        if dt.tzinfo is not None:
            dt = dt.astimezone(KST)
        return dt.date().strftime("%Y-%m-%d")
    return ""

def ymd_from_url(url: str) -> str:
    """기사 URL에서 YYYY-MM-DD 추출(있으면 가장 먼저 사용)."""
    if not url:
        return ""
    u = url

    # /YYYY/MM/DD, /YYYY-MM-DD, /YYYY.MM.DD
    m = re.search(r"(20\d{2})[./-](\d{2})[./-](\d{2})", u)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            datetime(y, mo, d)
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except ValueError:
            pass

    # 연속 8자리 YYYYMMDD (한경 등)
    for m in re.finditer(r"(20\d{2})(\d{2})(\d{2})", u):
        y, mo, d = map(int, m.groups())
        try:
            datetime(y, mo, d)
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except ValueError:
            continue

    # 쿼리 param 류
    m = re.search(r"(?:date|regdate|input|published|ud|dt|tm|ts)=(\d{8})", u, re.I)
    if m:
        s = m.group(1)
        y, mo, d = int(s[:4]), int(s[4:6]), int(s[6:8])
        try:
            datetime(y, mo, d)
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except ValueError:
            pass
    return ""

def ymd_from_text(date_text: str) -> str:
    """카드의 원문 텍스트에서 날짜 토큰을 추출해 YYYY-MM-DD로."""
    s = _strip_time_tail(date_text or "")
    for m in DATE_TOKEN_RE.finditer(s):
        token = m.group(0).strip()
        if len(token) > 20:
            continue
        dt = _parse_token_to_dt(token)
        if isinstance(dt, datetime):
            return _format_dt_ymd(dt)
    return ""

# ---- 기사 본문 페이지에서 날짜 추출 ----
META_DATE_SELECTORS = [
    ('meta', {'property': 'article:published_time'}),
    ('meta', {'name': 'article:published_time'}),
    ('meta', {'property': 'og:published_time'}),
    ('meta', {'name': 'og:published_time'}),
    ('meta', {'property': 'article:modified_time'}),
    ('meta', {'property': 'og:updated_time'}),
    ('meta', {'name': 'pubdate'}),
    ('meta', {'name': 'publishdate'}),
    ('meta', {'name': 'ptime'}),
    ('meta', {'itemprop': 'datePublished'}),
    ('time', {'datetime': True}),
]

DATE_TEXT_HINT_SELECTORS = [
    '[class*="date"]',
    '[id*="date"]',
    'span.info',
    'div.info',
    'p.info',
]

_session = requests.Session()
_session.headers.update(HEADERS)

def _parse_absolute_date_string(s: str) -> str:
    """본문 메타/텍스트에서 흔한 절대 날짜(연도 포함)를 YYYY-MM-DD로 파싱."""
    if not s:
        return ""
    s = s.strip()

    # ISO 8601 류
    m = re.search(r"(20\d{2})-(\d{2})-(\d{2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            datetime(y, mo, d)
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except ValueError:
            pass

    # 2025.09.13 / 2025. 9. 13.
    m = re.search(r"(20\d{2})\s*[.\-]\s*(\d{1,2})\s*[.\-]\s*(\d{1,2})", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            datetime(y, mo, d)
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except ValueError:
            pass

    # 2025년 9월 13일
    m = re.search(r"(20\d{2})\s*년\s*(\d{1,2})\s*월\s*(\d{1,2})\s*일", s)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            datetime(y, mo, d)
            return f"{y:04d}-{mo:02d}-{d:02d}"
        except ValueError:
            pass

    return ""

def ymd_from_article(url: str, timeout=7) -> str:
    """기사 페이지를 열어 메타/타임/날짜텍스트에서 YYYY-MM-DD 추출."""
    if not url:
        return ""
    try:
        r = _session.get(url, timeout=timeout)
        if r.status_code != 200 or not r.text:
            return ""
        soup = BeautifulSoup(r.text, "html.parser")

        # 1) 메타/타임 태그
        for tag, attrs in META_DATE_SELECTORS:
            for node in soup.find_all(tag, attrs=attrs if isinstance(attrs, dict) else {}):
                content = node.get("content") or node.get("datetime") or ""
                ymd = _parse_absolute_date_string(content)
                if ymd:
                    return ymd

        # 2) 화면 표시 텍스트(클래스/아이디에 date 포함)
        for sel in DATE_TEXT_HINT_SELECTORS:
            for node in soup.select(sel):
                txt = node.get_text(" ", strip=True)
                ymd = _parse_absolute_date_string(txt)
                if ymd:
                    return ymd

        return ""
    except Exception:
        return ""

def ymd_best(date_text: str, url: str) -> str:
    """
    최종 의사결정:
    1) URL → 2) 카드텍스트(연/월/일 or 상대) → 3) 본문 메타/타임
    """
    # 1) URL
    ymd = ymd_from_url(url)
    if ymd:
        return ymd
    # 2) 텍스트
    ymd = ymd_from_text(date_text)
    if ymd:
        return ymd
    # 3) 기사 페이지
    return ymd_from_article(url)

# ---------- 날짜 텍스트 추출(카드 내부만) ----------
def extract_date_text_from_card(root_node, soup):
    candidates = []

    parent = root_node
    for _ in range(5):  # 상위 5단계
        if not parent:
            break
        for sel in DATE_CONTAINER_SELECTORS:
            for node in parent.select(sel):
                txt = (node.get("datetime") or node.get_text(" ", strip=True) or "").strip()
                if txt:
                    candidates.append(txt)
        parent = parent.parent

    # 보조 탐색(좁게)
    for sp in root_node.find_all_next("span", limit=60):
        txt = sp.get_text(" ", strip=True)
        if txt:
            candidates.append(txt)
    for sp in root_node.find_all_previous("span", limit=30):
        txt = sp.get_text(" ", strip=True)
        if txt:
            candidates.append(txt)

    # 후보 중 실제 날짜 토큰이 추출되는 첫 텍스트 반환(없으면 빈 문자열)
    for txt in candidates:
        if DATE_TOKEN_RE.search(_strip_time_tail(txt)):
            return txt
    return ""

# ---------- 수집 ----------
def fetch_news(keyword: str, headless: bool = True):
    driver = build_driver(headless=headless)
    try:
        q = quote_plus(keyword)
        url = f"https://search.naver.com/search.naver?where=news&query={q}&sm=tab_opt&sort=1&start=1"
        print_step("페이지 접속 중…")
        driver.get(url)

        print_step("페이지 로딩/스크롤 중…")
        for _ in range(3):
            time.sleep(1.0)
            driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")

        html = driver.page_source
        soup = BeautifulSoup(html, "html.parser")
        results = []

        # 1) 구 구조
        titles = soup.select("a.news_tit")
        if titles:
            print_step(f"'a.news_tit' 구조 감지: {len(titles)}건 수집 중…")
            total = len(titles)
            for i, a in enumerate(titles, start=1):
                title = a.get_text(strip=True)
                link = a.get("href", "")
                date_text = extract_date_text_from_card(a, soup)
                results.append([title, date_text, link])
                print_progress(i, total, label="수집")

        # 2) 신 구조
        if not results:
            spans = soup.select("span.sds-comps-text-type-headline1")
            print_step(f"'신규 렌더' 구조 감지: {len(spans)}건 수집 중…")
            total = len(spans)
            for i, span in enumerate(spans, start=1):
                a = span.find_parent("a")
                if not a:
                    print_progress(i, total, label="수집"); continue
                title = span.get_text(strip=True)
                link = a.get("href", "")
                date_text = extract_date_text_from_card(span, soup)
                results.append([title, date_text, link])
                print_progress(i, total, label="수집")

        if not results:
            with open("debug2.html", "w", encoding="utf-8") as f:
                f.write(html)
            try:
                driver.save_screenshot("debug2.png")
            except Exception:
                pass
        return results

    finally:
        driver.quit()

# ---------- 출력/저장 ----------
def print_news_to_console(data):
    term_width = shutil.get_terminal_size().columns
    term_width = max(term_width, 80)
    title_width = max(30, term_width - 20 - 6)

    print("\n===== 뉴스 결과 =====")
    for idx, (title, ymd, link) in enumerate(data, start=1):
        short = textwrap.shorten(title, width=title_width, placeholder="…")
        if HAS_COLOR:
            num_part = f"{Fore.CYAN}{idx:>2}.{Style.RESET_ALL}"
            title_part = f"{Fore.WHITE}{short}{Style.RESET_ALL}"
            date_part = f"{Fore.GREEN}{(ymd or '').strip()}{Style.RESET_ALL}"
            link_part = f"{Fore.BLUE}{link}{Style.RESET_ALL}"
        else:
            num_part = f"{idx:>2}."
            title_part = short
            date_part = (ymd or "").strip()
            link_part = link
        print(f"{num_part} {title_part}  |  {date_part}")
        print(f"    ↳ {link}")

def excel_width_from_cm(cm: float) -> float:
    pixels = cm * 37.7952755906
    return max(8, (pixels - 5) / 7.0)

if __name__ == "__main__":
    keyword = input("검색어를 입력하세요: ").strip()
    data = fetch_news(keyword, headless=True)

    if not data:
        print("뉴스 데이터를 찾을 수 없습니다.")
        if os.path.exists("debug2.html"):
            print("디버그 파일을 저장했습니다: debug2.html / debug2.png (있다면)")
    else:
        before = len(data)
        data = dedupe_news(data)
        after = len(data)
        print_step(f"중복 제거: {before}건 → {after}건")

        # ymd 일괄 계산 (URL/텍스트/본문메타 순, 한 번만 수행)
        print_step("날짜 계산 중…")
        total = len(data)
        processed = []
        for i, (title, date_text, link) in enumerate(data, start=1):
            ymd = ymd_best(date_text, link) or ""
            processed.append([title, ymd, link])
            print_progress(i, total, label="날짜")

        # 터미널 출력
        print_news_to_console(processed)

        # 엑셀 저장(날짜는 문자열 YYYY-MM-DD)
        rows_for_excel = processed

        df_to_save = pd.DataFrame(rows_for_excel, columns=["제목", "날짜", "링크"])
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "output")
        os.makedirs(out_dir, exist_ok=True)
        out = os.path.join(out_dir, clean_filename(keyword) + "_1page.xlsx")

        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            sheet_name = "뉴스"
            df_to_save.to_excel(writer, index=False, sheet_name=sheet_name)
            ws = writer.sheets[sheet_name]
            from openpyxl.utils import get_column_letter
            from openpyxl.styles import Font

            # 제목 열 ≈ 10cm
            ws.column_dimensions[get_column_letter(1)].width = excel_width_from_cm(10.0)
            # 날짜 열 폭 보기 좋게
            ws.column_dimensions[get_column_letter(2)].width = 12
            # 링크 열 하이퍼링크 처리
            ws.column_dimensions[get_column_letter(3)].width = 12
            for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
                cell = row[0]
                url = str(cell.value or "").strip()
                if url.startswith("http"):
                    cell.hyperlink = url
                    cell.value = "링크"
                    cell.font = Font(color="0563C1", underline="single")

        print(f"\n뉴스 제목과 날짜가 '{out}' 파일에 저장되었습니다.")
