# 유튜브 AI 종목스캐너

국내 주식 관련 유튜브 채널의 최신 영상을 자동 수집하고, 자막(또는 오디오)을 Gemini AI로 분석하여 추천 종목을 엑셀 파일로 정리합니다.

---

## 동작 흐름

```
1. YouTube Data API
   └─ 각 채널의 20시간 이내 영상 ID·제목·업로드일시 수집 (채널당 최대 15개 조회, 1유닛/채널)

2. YouTubeTranscriptApi
   └─ 영상별 한국어 자막 추출
   └─ 자막 없는 영상 → yt-dlp로 오디오 다운로드 → Gemini Files API로 전사

3. Gemini AI (gemini-2.5-flash-lite)
   └─ 자막/오디오를 분석해 JSON 반환
       {
         "투자관련": true/false,
         "요약": "핵심 내용 3줄 이내",
         "추천종목": [{"종목명": "...", "사유": "..."}]
       }
   └─ 투자 무관 영상 → 두 시트 모두 제외
   └─ 투자관련 + 종목 있음 → 시트1 (종목추천)
   └─ 투자관련 + 종목 없음 → 시트2 (투자관련)

4. 종목코드 매핑
   └─ ../4_financial_scanner/output/ 최신 xlsx에서 종목명 → 티커(6자리) 조회

5. 엑셀 저장 (2개 시트) → 서식 적용 → (선택) 이메일 발송
```

---

## 엑셀 출력 구조

### 시트1 — 종목추천 (헤더 진파랑)

구체적인 종목을 언급·추천한 영상만 포함합니다.

| 컬럼 | 설명 |
|---|---|
| 채널 | 유튜브 채널명 |
| 업로드일시 | 영상 업로드 시각 (KST, YYYY-MM-DD HH:MM) |
| 제목 | 영상 제목 |
| 요약 | Gemini 요약 (3줄 이내) |
| 종목명 | 추천 종목명 (하이퍼링크 → 네이버 금융) |
| 종목코드 | 6자리 티커 (하이퍼링크 → FnGuide) |
| 추천사유 | Gemini가 추출한 추천 이유 |
| 링크 | 영상 YouTube 링크 |

### 시트2 — 투자관련 (헤더 진초록)

투자·주식 관련 내용이지만 특정 종목을 언급하지 않은 영상입니다 (시장 전망, 매크로 분석 등).

| 컬럼 | 설명 |
|---|---|
| 채널 | 유튜브 채널명 |
| 업로드일시 | 영상 업로드 시각 (KST) |
| 제목 | 영상 제목 |
| 요약 | Gemini 요약 (3줄 이내) |
| 링크 | 영상 YouTube 링크 |

> 투자·주식과 무관한 영상(채널 공지, 브이로그 등)은 두 시트 모두에서 제외됩니다.

### 시트3 — 실적정보 (헤더 갈색, 선택)

시트1(종목추천)에 등장한 종목 중 `4_financial_scanner/output/` 최신 파일에 티커가 일치하는 종목만 표시됩니다. 해당 파일이 없거나 일치 종목이 없으면 시트3은 생략됩니다.

| 컬럼 | 설명 |
|---|---|
| 추천채널 | 해당 종목을 추천한 유튜브 채널명 |
| 종목명 | 종목명 |
| 티커 | 6자리 종목코드 (하이퍼링크 → 네이버 금융) |
| 업종 | 업종 분류 |
| 종합점수 | 실적 종합 점수 |
| 현재가 / 적정주가 / 상승여력 | 가격 및 목표가 |
| 매출·영업이익 증가율 | 실적 추정치 |
| PER / PBR / PSR | 밸류에이션 지표 |

---

## 분석 채널 (총 11개)

| 채널명 | 채널 ID |
|---|---|
| 달란트투자 | UCBM86JVoHLqg9irpR2XKvGw |
| 815머니톡 | UCCG6BEYjfQMGzypJw2EJCDQ |
| 삼프로TV | UChlv4GSd7OQl3js-jkLOnFA |
| 김작가TV | UCvil4OAt-zShzkKHsg9EQAw |
| 떠먹여주는TV | UC5dEgOV_mGqMHXizL1drtvA |
| 경제야놀자 | UCgjsiInAgdceEq2lpRNNYvw |
| 딜사이트경제TV | UCNi2OWpVGVBC2SuMBGXFzbA |
| ETF아는형 | UCIt_hQy30blHrVAS7QLTyTg |
| 815캠퍼스 | UCnTZ9DHMo8oEoVIZrGlVoOA |
| 머니코믹스 | UCJo6G1u0e_-wS-JQn3T-zEw |
| 와이스트릿 | UCupslRq5jW95UGzPjOZz0FA |

> 채널 추가/변경: `main.py`의 `CHANNELS` 딕셔너리에서 수정

---

## 설치

```bash
pip install google-api-python-client youtube-transcript-api google-generativeai \
            python-dotenv openpyxl pandas rich yt-dlp xlwings
```

- `xlwings`: 엑셀 행 높이 AutoFit용 (Windows + Excel 설치 필요). 없어도 동작하며 경고만 출력됨.
- `yt-dlp`: 자막 없는 영상의 오디오 분석용. 자막 있는 영상만 처리할 경우 없어도 무방.

---

## 환경 변수 (.env)

```ini
# API 키
YOUTUBE_API_KEY=...        # Google Cloud Console → YouTube Data API v3
GEMINI_API_KEY=...         # Google AI Studio (유료 결제 권장)

# 이메일 발송 (Gmail)
GMAIL_USER=yourname@gmail.com
GMAIL_APP_PASSWORD=...     # Gmail 앱 비밀번호 (2단계 인증 필요)

# 수신자
NOTIFY_EMAIL=to@example.com
NOTIFY_EMAIL_WIFE=to2@example.com
NOTIFY_EMAILS_EXTRA=a@x.com, b@x.com   # 추가 수신자 (쉼표 구분)

# 이메일 발송 여부 (true/false)
SEND_EMAIL=false
```

---

## 실행

```bash
python main.py
```

---

## API 한도

| 항목 | 한도 | 소비량 |
|---|---|---|
| YouTube Data API | 일 10,000유닛 | 11채널 × 1유닛 = **11유닛/실행** |
| YouTube 자막 스크래핑 | YouTube 서버 측 제한 | 20시간 이내 영상만 처리 — **하루 1~2회** 권장 |
| Gemini API | 유료 과금 | 20시간 이내 영상 수만큼 호출 |
| Gemini 503 | 일시적 과부하 | 자동 재시도 (10s → 20s → 40s) |
| Gemini 429 | 할당량 초과 | retryDelay 파싱 후 자동 재시도 |

> 12채널, 20시간 이내 업로드 영상만 분석 → 소요 시간은 영상 수에 따라 변동

---

## 파일 구조

```
6_유튜브_AI종목스캐너/
├── main.py                       # 메인 실행 스크립트
├── format_excel.py               # 엑셀 포맷 정리 (행 분리·병합·하이퍼링크·AutoFit)
├── .env                          # API 키 및 설정 (git 제외)
├── www.youtube.com_cookies.txt   # 자막 IP 차단 우회용 쿠키 (Netscape 형식)
└── output/                       # 분석 결과 엑셀 저장 위치
```

---

## 쿠키 파일 갱신 방법

자막 추출 시 YouTube IP 차단(403/429)이 발생하면 `www.youtube.com_cookies.txt`를 최신으로 교체합니다.

1. Chrome 확장 프로그램 **"Get cookies.txt LOCALLY"** 설치
2. `youtube.com` 로그인 상태에서 쿠키 내보내기 (Netscape 형식)
3. 내보낸 파일을 `www.youtube.com_cookies.txt`로 저장

---

## 이메일 발송

분석 완료 후 `SEND_EMAIL=true`이면 수신자 목록에 엑셀을 첨부하여 발송합니다.

- 제목: `[유튜브 추천종목] YYYY-MM-DD HH:MM 분석 완료 — N개 종목`
- 본문: 채널별 추천 종목 요약
- 첨부: 결과 엑셀 파일

> Gmail 앱 비밀번호 발급: Google 계정 → 보안 → 2단계 인증 → 앱 비밀번호
