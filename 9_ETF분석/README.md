# ETF 가치 분석 스캐너

단순 수익률이 아닌, ETF 구성 종목이 **실제로 저평가되어 있는지**를 분석합니다.  
S-RIM 적정주가 대비 상승여력을 구성 비중으로 가중 평균해 ETF 자체의 **가치 점수**를 산출합니다.

---

## 동작 흐름

```
1. FDR StockListing('ETF/KR')
   └─ 국내 상장 전체 ETF 목록 수집 (티커, 종목명, 시가총액, 보수율 등)

2. pykrx get_etf_portfolio_deposit_file
   └─ ETF별 구성 종목(PDF) 조회
   └─ 4_financial_scanner S-RIM 데이터와 티커 매칭 → 비중 가중 평균 상승여력 = 가치점수 산출

3. Gemini AI (gemini-2.5-flash-lite)
   └─ ETF 명칭 배치 분석 → 테마 자동 분류
      반도체 / AI·IT / 배당 / 채권 / 원자재·금 / 금융 /
      바이오·헬스케어 / 소비재 / 에너지 / 부동산·리츠 / 글로벌·해외 / 기타

4. 엑셀 저장 → 서식 적용 → (선택) 이메일 발송
```

---

## 엑셀 출력 구조

### 시트1 — 종합대시보드 (헤더 진파랑)

전체 ETF를 가치점수 내림차순으로 정렬합니다.

| 컬럼 | 설명 |
|---|---|
| 티커 | 6자리 종목코드 |
| 종목명 | ETF 명칭 (하이퍼링크 → 네이버 금융) |
| 테마 | Gemini AI 자동 분류 테마 |
| 시가총액 | ETF 시가총액 |
| 보수율 | 연 보수율 |
| 가치점수 | 구성 종목 비중 가중 평균 상승여력(%) |
| 매칭종목수 | S-RIM 데이터와 매칭된 구성 종목 수 |
| 구성종목수 | 전체 구성 종목 수 |
| 상위종목 | 비중 상위 3개 종목명 |

> 가치점수 ≥ 20% → 연초록 / ≥ 10% → 연노랑

### 시트2~ — 테마별 시트

각 테마에 해당하는 ETF만 필터링하여 가치점수 내림차순 정렬합니다.

- 반도체 / AI/IT / 배당 / 채권 / 원자재/금
- 금융 / 바이오/헬스케어 / 소비재 / 에너지
- 부동산/리츠 / 글로벌/해외 / 기타

---

## 가치점수 계산 방식

```
가치점수 = Σ (구성종목_비중 × 구성종목_상승여력1) / Σ 구성종목_비중

상승여력1 = (S-RIM 적정주가 - 현재가) / 현재가 × 100
           (4_financial_scanner/output 최신 파일 기준)
```

S-RIM 데이터와 매칭되지 않은 구성 종목은 가중 평균에서 제외됩니다.

---

## 설치

```bash
pip install FinanceDataReader pykrx google-genai python-dotenv openpyxl pandas rich xlwings
```

- `pykrx`: ETF 구성 종목(PDF) 조회용. KRX 로그인 필요 (KRX_ID / KRX_PW)
- `xlwings`: 엑셀 행 높이 AutoFit용 (Windows + Excel 설치 필요). 없어도 동작하며 경고만 출력됨.

---

## 환경 변수 (.env)

```ini
# API 키
GEMINI_API_KEY=...         # Google AI Studio

# KRX 로그인 (pykrx PDF 조회용)
KRX_ID=...
KRX_PW=...

# 이메일 발송 (Gmail)
GMAIL_USER=yourname@gmail.com
GMAIL_APP_PASSWORD=...     # Gmail 앱 비밀번호 (2단계 인증 필요)

# 수신자
NOTIFY_EMAIL=to@example.com
NOTIFY_EMAIL_WIFE=to2@example.com
NOTIFY_EMAILS_EXTRA=a@x.com, b@x.com

# 이메일 발송 여부 (true/false)
SEND_EMAIL=true
```

---

## 실행

```bash
python main.py
```

---

## API 사용량

| 항목 | 소비량 |
|---|---|
| pykrx KRX PDF | ETF 수만큼 호출 (국내 ETF 약 800~900개) |
| Gemini API | ETF 수 ÷ 40 회 (배치 처리) — 약 20~25회 호출 |

> 전체 실행 시간: 약 **20~30분** (PDF 조회 0.08초 간격 × ETF 수)

---

## 파일 구조

```
9_ETF분析/
├── main.py             # 메인 실행 스크립트
├── format_excel.py     # 엑셀 서식 적용 (헤더·색상·하이퍼링크·AutoFit)
├── .env                # API 키 및 설정 (git 제외)
└── output/             # 분석 결과 엑셀 저장 위치
```

---

## 이메일 발송

분석 완료 후 `SEND_EMAIL=true`이면 수신자 목록에 엑셀을 첨부하여 발송합니다.

- 제목: `[ETF 가치분석] YYYY-MM-DD HH:MM — 가치점수 산출 N개`
- 본문: 가치점수 상위 5개 ETF + 테마별 1위 ETF 요약
- 첨부: 결과 엑셀 파일

> Gmail 앱 비밀번호 발급: Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호
