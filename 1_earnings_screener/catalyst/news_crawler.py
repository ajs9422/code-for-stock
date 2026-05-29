"""
네이버 금융 뉴스 크롤러
- finance.naver.com/item/news_news.naver?code={code}&page={n}
- 최근 30일치 뉴스 제목 · 출처 · 날짜 수집
- SQLite catalyst_news 테이블에 저장 (중복 방지)
- 요청 간 1~2초 delay
"""

import random
import sqlite3
import time
from contextlib import contextmanager
from datetime import date, timedelta
from typing import Optional

import requests
from bs4 import BeautifulSoup

from config import DB_PATH, DELAY_MIN, DELAY_MAX, TIMEOUT

_NEWS_URL = "https://finance.naver.com/item/news_news.naver"
_HEADERS  = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer":       "https://finance.naver.com",
    "Accept-Language": "ko-KR,ko;q=0.9",
}


# ── DB ────────────────────────────────────────────────────────────────────────

@contextmanager
def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    finally:
        con.close()


def _init_table():
    with _conn() as con:
        con.execute("""
        CREATE TABLE IF NOT EXISTS catalyst_news (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            code        TEXT    NOT NULL,
            title       TEXT    NOT NULL,
            source      TEXT,
            news_date   TEXT,
            url         TEXT,
            fetched_at  TEXT    DEFAULT (date('now')),
            UNIQUE (code, title, news_date)
        )
        """)


def _save(code: str, items: list[dict]):
    if not items:
        return
    with _conn() as con:
        con.executemany("""
        INSERT OR IGNORE INTO catalyst_news
            (code, title, source, news_date, url)
        VALUES
            (:code, :title, :source, :news_date, :url)
        """, [{"code": code, **item} for item in items])


# ── 파싱 ──────────────────────────────────────────────────────────────────────

def _get_soup(page: int, code: str) -> Optional[BeautifulSoup]:
    try:
        res = requests.get(
            _NEWS_URL,
            params={"code": code, "page": page},
            headers=_HEADERS,
            timeout=TIMEOUT,
        )
        res.raise_for_status()
        res.encoding = "euc-kr"
        return BeautifulSoup(res.text, "lxml")
    except Exception:
        return None


def _parse_date(text: str) -> Optional[date]:
    """'2026.04.23 08:30' 또는 '04.23' → date 객체."""
    text = text.strip().split()[0]          # 시간 부분 제거
    parts = text.split(".")
    try:
        if len(parts) == 3:
            return date(int(parts[0]), int(parts[1]), int(parts[2]))
        if len(parts) == 2:
            today = date.today()
            return date(today.year, int(parts[0]), int(parts[1]))
    except ValueError:
        pass
    return None


def _parse_page(soup: BeautifulSoup, code: str) -> list[dict]:
    """type5 테이블에서 뉴스 행 파싱."""
    items = []
    table = soup.find("table", class_="type5")
    if not table:
        return items

    for tr in table.find_all("tr"):
        cells = tr.find_all("td")
        if len(cells) < 3:
            continue
        a = cells[0].find("a")
        if not a:
            continue
        title = a.get_text(strip=True)
        if not title:
            continue
        href = a.get("href", "")
        date_text = cells[2].get_text(strip=True)
        # 날짜 형식 검증 ('.' 포함, 숫자로 시작)
        if not date_text or not date_text[0].isdigit() or "." not in date_text:
            continue
        items.append({
            "title":     title,
            "source":    cells[1].get_text(strip=True),
            "news_date": date_text,
            "url":       (
                f"https://finance.naver.com{href}"
                if href.startswith("/") else href
            ),
        })
    return items


# ── 공개 인터페이스 ───────────────────────────────────────────────────────────

def fetch_news(code: str, days: int = 30) -> list[dict]:
    """
    네이버 금융 뉴스를 최근 N일치 크롤링 후 DB 저장.
    반환: 수집된 뉴스 목록 (최신순)
    """
    _init_table()
    cutoff    = date.today() - timedelta(days=days)
    all_items: list[dict] = []

    for page in range(1, 30):          # 최대 30페이지
        soup = _get_soup(page, code)
        if not soup:
            break

        raw = _parse_page(soup, code)
        if not raw:
            break

        batch: list[dict] = []
        stop = False
        for item in raw:
            parsed = _parse_date(item["news_date"])
            if parsed and parsed < cutoff:
                stop = True
                break
            batch.append(item)

        if batch:
            all_items.extend(batch)
            _save(code, batch)

        if stop or len(raw) < 10:      # 마지막 페이지이거나 cutoff 도달
            break

        time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))

    return all_items


def load_news(code: str, days: int = 30) -> list[dict]:
    """DB에서 최근 N일 뉴스 조회 (최신순)."""
    _init_table()
    since = str(date.today() - timedelta(days=days))
    with _conn() as con:
        rows = con.execute("""
        SELECT title, source, news_date, url
        FROM   catalyst_news
        WHERE  code = ? AND news_date >= ?
        ORDER  BY news_date DESC
        """, (code, since)).fetchall()
    return [dict(r) for r in rows]
