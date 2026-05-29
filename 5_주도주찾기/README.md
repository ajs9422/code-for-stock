# 🔥 주도 섹터 · 주도주 자동 탐색 시스템

## 프로젝트 구조

```
hot_sector/
├── main.py            # 실행 진입점 (Step 1~4 통합)
├── config.yaml        # 모든 파라미터 설정 (여기만 수정)
├── requirements.txt   # 의존성
├── README.md
└── reports/           # 결과 저장 (자동 생성)
    └── sector_report_YYYYMMDD.csv
```

## 설치

```bash
pip install -r requirements.txt
```

## 실행

```bash
# 기본 실행 (오늘 기준 최근 거래일)
python main.py

# 특정 날짜 분석
python main.py --date 20260425

# 섹터 TOP 10 출력
python main.py --top 10

# 장중 자동 실행 (스케줄러)
python main.py --schedule
```

## 분석 파이프라인

```
[STEP 1] 데이터 수집
  1-A  업종별 등락률      pykrx 업종 지수
  1-B  전종목 OHLCV      KOSPI + KOSDAQ
  1-C  거래대금 MA20     20거래일 이동평균
  1-D  52주 신고가       근접도 계산용

[STEP 2] 기술적 지표
  RSI (pandas_ta)        과매도/과열 필터

[STEP 3] 스코어 + 필터
  핫 스코어             등락률 35% + 거래대금급증 35% + 모멘텀 20% + 뉴스 10%
  주도주 필터           거래대금배율, 52주근접도, RSI 범위

[STEP 4] 출력 · 알림
  콘솔 리포트
  CSV / Excel 저장
  텔레그램 알림 (선택)
```

## Claude Code 보완 포인트

코드 내 `TODO` 태그로 확장 포인트가 표시되어 있음:

| 태그 | 위치 | 내용 |
|------|------|------|
| `TODO:SCORE`  | `calc_hot_score()` | 가중치 조정, ML 스코어링 교체 |
| `TODO:FILTER` | `filter_leaders()` | 수급, MACD, 신고가 돌파 조건 추가 |
| `TODO:ALERT`  | `send_telegram_alert()` | 메시지 포맷, 차트 이미지 첨부 |
| `TODO:VISUAL` | `run_scheduler()` | Streamlit 대시보드 연동 |

### Claude Code 명령 예시

```bash
# 프로젝트 열고 전체 분석
claude "TODO:FILTER 주석 있는 함수에 외국인/기관 순매수 조건 추가해줘"

# 텔레그램 알림 개선
claude "send_telegram_alert 함수에 섹터 차트 이미지도 같이 보내도록 수정해줘"

# 스코어 로직 개선
claude "calc_hot_score 함수를 머신러닝 기반으로 바꾸고 학습 데이터 수집 코드도 추가해줘"
```

## config.yaml 주요 파라미터

```yaml
filter:
  top_n_sectors: 5        # HOT 섹터 몇 개 볼지
  top_n_stocks: 3         # 섹터당 주도주 몇 개

scoring:
  weight_return: 0.35     # 등락률 가중치
  weight_volume_surge: 0.35  # 거래대금 급증 가중치
  weight_momentum: 0.20   # 52주 신고가 가중치

leader_filter:
  rsi_min: 50             # RSI 최소값
  rsi_max: 75             # RSI 최대값
  volume_surge_min: 2.0   # 거래대금 최소 배율

alert:
  telegram_enabled: true  # 텔레그램 알림 켜기
  telegram_token: "..."   # BotFather 토큰
  telegram_chat_id: "..." # Chat ID
```
