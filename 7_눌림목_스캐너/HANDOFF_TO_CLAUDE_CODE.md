# Claude Code 인수인계 — krx-pullback-scanner

## 프로젝트 목적
한국 주식(KRX) 에서 **박스권 돌파 + 거래량 급증 후 눌림목** 구간을 자동 발굴하는 스캐너.
발굴된 종목은 **일봉 / 주봉 / 월봉** 3종 차트 PNG 를 자동 저장.

---

## 디렉토리 구조 (목표)

```
krx-pullback-scanner/
├── .env                        ← KRX_AUTH_KEY 보관 (git 제외)
├── .gitignore
├── requirements.txt
├── scanner_v3.py               ← 메인 스캐너 (완성본 첨부)
├── charts/                     ← 차트 PNG 저장 폴더 (자동 생성)
│   └── {종목코드}_{종목명}/
│       ├── {코드}_{이름}_일봉.png
│       ├── {코드}_{이름}_주봉.png
│       └── {코드}_{이름}_월봉.png
└── results/                    ← CSV 결과 저장 폴더
    └── scan_v3_YYYYMMDD.csv
```

---

## .env 설정

```env
# KRX OpenAPI 인증키 (data.krx.co.kr 발급)
# 없으면 pykrx 자동 사용 (기능 동일, 속도 약간 느림)
KRX_AUTH_KEY=여기에_발급받은_키_입력
```

---

## requirements.txt

```
pykrx>=1.0.47
pandas>=2.0.0
numpy>=1.24.0
matplotlib>=3.7.0
mplfinance>=0.12.10b0
python-dotenv>=1.0.0
tabulate>=0.9.0
```

---

## 핵심 로직 요약

### 스캔 3단계 필터

| 단계 | 조건 | 설명 |
|------|------|------|
| 1 | 일봉 MA60 횡보·우상향 | 하락 추세 종목 제거 |
| 2 | 주봉 MA20 우상향 | 중기 추세 확인 |
| 3 | 월봉 MA5 우상향 | 장기 추세 확인 |
| 4 | 박스권 돌파 | 3~6개월 횡보 후 박스 상단 돌파 |
| 5 | 거래량 급증 | 돌파일 거래량 ≥ 20일 평균 × 3배 |
| 6 | 눌림목 | 돌파 후 고점 대비 3~15% 하락 구간 |

### MA 기울기 판단 방식 (ma_slope 함수)
```python
# 최근 4구간(봉) 동안 MA 연속 상승 AND 기울기 >= slope_min
window = ma.iloc[-(look+1):]
diffs  = window.diff().dropna()
slope  = (window.iloc[-1] / window.iloc[0] - 1) * 100
rising = (diffs > 0).all()
```

### 점수 계산 (100점 만점)
| 항목 | 배점 |
|------|------|
| 거래량 강도 | 30점 |
| 눌림 적절성 (5~10%가 최고) | 20점 |
| 박스 기간 (길수록 유리) | 10점 |
| 박스 타이트함 | 10점 |
| 일봉 MA60 기울기 보너스 | 10점 |
| 주봉 MA20 기울기 보너스 | 10점 |
| 월봉 MA5 기울기 보너스 | 10점 |

### 데이터 수집 우선순위
1. `.env` 의 `KRX_AUTH_KEY` 있으면 → KRX OpenAPI 사용
2. 없으면 → pykrx 자동 fallback

---

## 실행 방법

```bash
# 초기 설치
cd krx-pullback-scanner
pip install -r requirements.txt

# 전체 스캔 (상위 10개 3종 차트 자동 생성)
python scanner_v3.py

# 옵션 조합
python scanner_v3.py --market KOSPI --vol-mult 4.0 --chart-top 5 --save

# 특정 종목 차트만 즉시 확인
python scanner_v3.py --chart-only 011170
```

### 주요 CLI 파라미터

| 파라미터 | 기본값 | 설명 |
|----------|--------|------|
| `--market` | ALL | ALL / KOSPI / KOSDAQ |
| `--box-days` | 90 | 박스권 탐지 기간 (거래일) |
| `--vol-mult` | 3.0 | 돌파일 거래량 배수 |
| `--pullback-min` | 0.03 | 눌림 최솟값 (3%) |
| `--pullback-max` | 0.15 | 눌림 최댓값 (15%) |
| `--daily-ma` | 60 | 일봉 MA 기간 |
| `--weekly-ma` | 20 | 주봉 MA 기간 |
| `--monthly-ma` | 5 | 월봉 MA 기간 |
| `--chart-top` | 10 | 차트 저장 종목 수 |
| `--no-chart` | - | 차트 생성 안 함 |
| `--save` | - | 결과 CSV 저장 |

---

## Claude Code 에서 완료해야 할 작업

### 1순위 — 환경 완성
- [ ] `krx-pullback-scanner/` 폴더 생성
- [ ] `.env` 파일 생성 및 `KRX_AUTH_KEY` 입력
- [ ] `.gitignore` 생성 (`.env`, `charts/`, `results/`, `__pycache__/` 포함)
- [ ] `requirements.txt` 작성 및 `pip install`
- [ ] `scanner_v3.py` 배치

### 2순위 — 실제 데이터 검증
- [ ] `python scanner_v3.py --chart-only 011170` 으로 롯데케미칼 차트 정상 생성 확인
- [ ] KRX API 키 연결 확인 (없으면 pykrx로 동작 확인)
- [ ] 한글 폰트 설치 확인 (`pip install koreanize-matplotlib` 또는 Malgun Gothic)

### 3순위 — 기능 추가 (선택)
- [ ] 스캔 결과를 HTML 리포트로 출력 (차트 인라인 포함)
- [ ] 종목 알림 (텔레그램 or 이메일)
- [ ] 스케줄러 (장 마감 후 자동 실행 — `schedule` 라이브러리)
- [ ] PER/PBR 저평가 필터 추가 (실적 대비 저평가 종목 우선)

---

## 한글 폰트 설정 (차트 깨짐 방지)

```bash
# macOS
brew install font-nanum

# Ubuntu/Debian
apt-get install fonts-nanum
fc-cache -fv

# pip (가장 간단)
pip install koreanize-matplotlib
```

```python
# scanner_v3.py 상단에 추가 (pip 방법 사용 시)
import koreanize_matplotlib
```

---

## 주요 함수 맵

```
scanner_v3.py
│
├── 데이터 수집
│   ├── get_daily_ohlcv(ticker, days)     일봉 DataFrame 반환
│   ├── to_weekly(df)                     일봉→주봉 리샘플
│   ├── to_monthly(df)                    일봉→월봉 리샘플
│   └── get_ticker_list(market)           종목 코드 목록
│
├── 분석 로직
│   ├── ma_slope(series, period, look)    MA 기울기·방향 판단
│   ├── detect_box(df, box_days, tol)     박스권 상하단 탐지
│   ├── detect_breakout(df, ...)          거래량 돌파 감지
│   ├── detect_pullback(df, peak, ...)    눌림목 여부 판단
│   └── analyze(ticker, cfg)             단일 종목 종합 분석
│
├── 차트 생성
│   ├── chart_daily(df, info, cfg, dir)   일봉 PNG 저장
│   ├── chart_weekly(df, info, cfg, dir)  주봉 PNG 저장
│   ├── chart_monthly(df, info, cfg, dir) 월봉 PNG 저장
│   └── generate_charts(ticker, info, cfg, dir)  3종 일괄 저장
│
└── 실행
    ├── run_scanner(cfg)                  전체 종목 스캔
    └── __main__                          CLI 파싱 및 실행
```

---

## 매매 전략 요약 (코드 맥락 이해용)

```
[박스권 3~6개월 횡보]
        ↓
[거래량 급증 + 박스 상단 돌파]  ← 돌파 당일 매수 X
        ↓
[고점 형성 후 눌림목 (3~15% 하락)]  ← 여기서 매수 대기
        ↓
[박스 상단(신규 지지선) 부근 반등 확인]  ← 진입 시그널
        ↓
목표가: 박스 폭만큼 추가 상승
손절선: 박스 상단 하향 이탈
```
