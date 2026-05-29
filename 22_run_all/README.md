# AI 투자 브리핑 통합 실행기

5개 분석 스크립트를 2단계 병렬로 실행하고, 결과를 단일 이메일로 발송합니다.

## 프로젝트 구조

```
22_run_all/
├── run_all.py         # 통합 실행 (Phase 1 → Phase 2 → 이메일)
├── preview_email.py   # 이메일 미리보기 + 수동 발송 (분석 생략)
├── .env               # Gmail 계정·수신자 설정
├── README.md
└── output/            # 각 프로젝트 output 복사본 (자동 생성)
```

## 설치

```bash
pip install rich python-dotenv
```

## 실행

```bash
# 전체 분석 + 이메일 발송
python run_all.py

# 기존 output의 이메일 미리보기 (분석 없이 이메일만 발송)
python preview_email.py
```

## 실행 파이프라인

```
[Phase 1 — 병렬]
  [4]  4_financial_scanner     재무스캐너 (매출·영익 YoY + S-RIM)
  [10] 10_투자일정             국내외 투자 이벤트 캘린더
  [14] 14_마켓인사이트          글로벌·국내 시황 AI 브리핑 (PDF 생성)

        ↓ 4번 완료 후

[Phase 2 — 병렬]   ← 4_financial_scanner/output/ xlsx 필요
  [6]  6_유튜브_AI종목스캐너   유튜브 채널 AI 분석
  [8]  8_뉴스기사_인사이트      뉴스 기사 AI 분석 (LocalLLM)

        ↓

[이메일 발송]
  첨부: 마켓인사이트 PDF + 나머지 Excel 4개
  본문: 시황 브리핑 헤더 + 프로젝트별 요약
```

> [8] 뉴스기사 인사이트는 Ollama(LocalLLM)를 사용하며, GPU 세마포어로 동시 실행 1개로 제한됩니다.

## 스킵 로직

`4_financial_scanner`, `10_투자일정`, `14_마켓인사이트`는 오늘 날짜(`YYYY_MMDD`) output 파일이 이미 존재하면 분석을 건너뜁니다.
기존 파일을 `output/`에 복사하고 다음 단계로 진행합니다.

## .env 설정

```ini
GMAIL_USER=your@gmail.com
GMAIL_APP_PASSWORD=xxxx xxxx xxxx xxxx   # Google 앱 비밀번호

NOTIFY_EMAIL=주수신자@example.com
NOTIFY_EMAIL_WIFE=추가수신자@example.com  # 선택
NOTIFY_EMAILS_EXTRA=추가1@a.com,추가2@b.com  # 쉼표 구분, BCC 발송
```

이메일 발송을 끄려면:

```bash
SEND_EMAIL=false python run_all.py
```

## 예상 소요 시간

| 프로젝트 | 예상 |
|----------|------|
| 4_financial_scanner | ~8분 |
| 10_투자일정 | ~3분 |
| 14_마켓인사이트 | ~8분 |
| 6_유튜브 | ~7분 (Phase 2) |
| 8_뉴스기사 (LocalLLM) | ~20분 (Phase 2) |
| **전체 (Phase 1 + Phase 2)** | **약 30~35분** |

각 프로젝트의 타임아웃은 2시간입니다.
