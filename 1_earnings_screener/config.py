"""
================================================
  전역 설정
================================================
"""
import os
from pathlib import Path
from dotenv import load_dotenv

load_dotenv()

BASE_DIR = Path(__file__).parent

# ── 대상 연도 ──────────────────────────────────
ACTUAL_YEAR   = 2025   # 25A: 전년 실적
ESTIMATE_YEAR = 2026   # 26E: 당해년 컨센서스

# ── 유니버스 필터 ──────────────────────────────
MIN_MARKET_CAP   = 2000   # 시가총액 최소 (억원)
MIN_COVERAGE     = 2      # 커버 증권사 최소 수

# ── 스크리닝 AND 조건 ──────────────────────────
MIN_REVENUE_GROWTH   = 10.0    # 매출 증가율 최소 (%)
MIN_OP_PROFIT_GROWTH = 20.0    # 영업이익 증가율 최소 (%)
MIN_OPM_IMPROVEMENT  = 3.0     # 영업이익률 개선 최소 (% point)

# ── 밸류에이션 / 리스크 조건 ───────────────────
MAX_FWD_PER          = 35.0   # 선행 PER 상한
MAX_PRICE_52W_PCT    = 95.0   # 52주 고점 대비 현재가 상한 (%) — 고점 근처 과열 제외
MAX_SURGE_FROM_LOW   = 150.0  # 52주 저점 대비 급등률 상한 (%) — 급등 종목 제외
MAX_DEBT_RATIO       = 300.0  # 부채비율 상한 (%)

# ── 크롤링 ────────────────────────────────────
DELAY_MIN   = 1.0   # 요청 간 최소 딜레이 (초)
DELAY_MAX   = 2.0   # 요청 간 최대 딜레이 (초)
TIMEOUT     = 15    # 요청 타임아웃 (초)

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/124.0.0.0 Safari/537.36'
    ),
    'Accept-Language': 'ko-KR,ko;q=0.9',
    'Referer': 'https://comp.fnguide.com',
}

# ── 저장 경로 ─────────────────────────────────
DB_PATH     = BASE_DIR / "storage" / "screener.db"
REPORT_DIR  = BASE_DIR / "report" / "output"

# ── DART API ──────────────────────────────────
DART_API_KEY = os.getenv("DART_API_KEY", "")

# ── 이메일 설정 (Gmail SMTP_SSL) ──────────────
import re as _re
EMAIL_USER = os.getenv("GMAIL_USER", "")
EMAIL_PASS = os.getenv("GMAIL_APP_PASSWORD", "")
EXTRA_EMAILS: list[str] = [e.strip() for e in _re.split(r"[,;]", os.getenv("NOTIFY_EMAILS_EXTRA", "")) if e.strip()]
MAIN_EMAILS: list[str]  = [e for e in [os.getenv("NOTIFY_EMAIL", ""), os.getenv("NOTIFY_EMAIL_WIFE", "")] if e.strip()]
NOTIFY_EMAILS: list[str] = MAIN_EMAILS + EXTRA_EMAILS
EMAIL_TO = ", ".join(NOTIFY_EMAILS)

# ── 테스트 모드 ───────────────────────────────
TEST_MODE = os.getenv("TEST_MODE", "false").lower() == "true"  # true로 설정 시 5개 종목만 테스트
