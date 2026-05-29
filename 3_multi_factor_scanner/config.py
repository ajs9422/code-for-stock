"""
================================================
  멀티팩터 스캐너 — 설정
================================================

[자주 변경하는 항목]
  SEARCH_KEYWORD  : 유튜브 검색 키워드 (분석할 전문가 채널/인물)
  MAX_VIDEOS      : 분석할 최신 영상 수
  LOOKBACK_DAYS   : 주가·수급 데이터 조회 기간 (일)
  MARKETS         : 스캔할 시장 목록 ("KOSPI" / "KOSDAQ")
"""

import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).parent

# ── API 키 ────────────────────────────────────
GEMINI_API_KEY     = os.getenv("GEMINI_API_KEY", "")
DART_API_KEY       = os.getenv("DART_API_KEY", "")
KRX_ID             = os.getenv("KRX_ID", "")
KRX_PW             = os.getenv("KRX_PW", "")

# ── 이메일 설정 ───────────────────────────────
import re as _re
EMAIL_USER = os.getenv("GMAIL_USER", "")
EMAIL_PASS = os.getenv("GMAIL_APP_PASSWORD", "")
_extra     = [e.strip() for e in _re.split(r"[,;]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
NOTIFY_EMAILS: list[str] = [e for e in [
    os.getenv("NOTIFY_EMAIL", ""),
    os.getenv("NOTIFY_EMAIL_WIFE", ""),
] if e.strip()] + _extra
EMAIL_TO = ", ".join(NOTIFY_EMAILS)

# ── 유튜브 검색 ───────────────────────────────
SEARCH_KEYWORD = "이권희 대표"   # 검색할 키워드 (인물명, 채널명 등)
MAX_VIDEOS     = 5               # 분석할 최신 영상 수

# ── 데이터 조회 기간 ──────────────────────────
LOOKBACK_DAYS = 30               # 주가·수급 데이터 조회 기간 (일)

# ── 스캔 유니버스 ─────────────────────────────
MARKETS = ["KOSPI", "KOSDAQ"]    # 스캔할 시장 목록

# 야간·장 마감으로 KRX 응답 없을 때 대체 종목
FALLBACK_TICKERS = [
    "005930",  # 삼성전자
    "012450",  # 한화에어로스페이스
    "484810",  # 이오테크닉스 (예시)
    "079550",  # LIG넥스원
    "000660",  # SK하이닉스
    "267260",  # 현대일렉트릭
    "005380",  # 현대차
]

# ── 크롤링 딜레이 ─────────────────────────────
TICKER_DELAY = 0.2   # 종목 간 딜레이 (초) — KRX 부하 방지
VIDEO_DELAY  = 1.0   # 영상 분석 간 딜레이 (초) — Gemini API 부하 방지

# ── 저장 경로 ─────────────────────────────────
OUTPUT_DIR = BASE_DIR / "output"
