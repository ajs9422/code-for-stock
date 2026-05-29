import os
from dotenv import load_dotenv

# .env 파일 로드 (상위 디렉토리 또는 현재 디렉토리)
load_dotenv()

# API Keys
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")

# 설정
DEFAULT_MONTHS_AHEAD = 2
OUTPUT_DIR = "output"

# 엑셀 파일명 형식
EXPORT_FILENAME_FORMAT = "투자일정_분석_{date}.xlsx"

# Gemini 모델 설정
GEMINI_MODEL_NAME = "gemini-2.0-flash"

# 크롤링 설정
NAVER_SCHEDULE_URL = "https://finance.naver.com/sise/sise_schedule.naver"
USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
