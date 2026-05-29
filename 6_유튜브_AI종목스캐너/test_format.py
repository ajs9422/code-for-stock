"""
format_excel.py 빠른 검증 스크립트
- 4_실적개선 최신 파일에서 종합/종합(분기포함) 데이터를 5행 샘플링
- 더미 종목추천·투자관련 시트와 함께 test_input.xlsx 생성
- format_excel.py 실행 후 test_input_fmt.xlsx 오픈
"""
import os, glob, subprocess, sys
import pandas as pd

BASE = os.path.dirname(os.path.abspath(__file__))
EARNINGS_DIR = os.path.join(BASE, "..", "4_실적개선", "output")

# 1. 4번 최신 파일 찾기
files = sorted([f for f in glob.glob(os.path.join(EARNINGS_DIR, "*.xlsx")) if "_hl" not in f])
if not files:
    print("[ERR] 4_실적개선/output 에 xlsx 파일이 없습니다.")
    sys.exit(1)
src = files[-1]
print(f"소스: {src}")

# 2. 종합 / 종합(분기포함) 5행 샘플
df_brief = pd.read_excel(src, sheet_name="종합",         dtype=str).head(5)
df_full  = pd.read_excel(src, sheet_name="종합(분기포함)", dtype=str).head(5)

# 3. 추천채널 컬럼 삽입 (맨 앞)
for df in [df_brief, df_full]:
    df.insert(0, "추천채널", "테스트채널A, 테스트채널B")

# 4. 더미 종목추천 / 투자관련
df_stocks = pd.DataFrame([{
    "채널": "테스트채널A", "업로드일시": "2026-05-11 20:00",
    "제목": "테스트 영상 제목", "출연자": "홍길동",
    "요약": "테스트 요약입니다.", "종목명": df_brief.iloc[0].get("종목명", "샘플"),
    "종목코드": df_brief.iloc[0].get("티커", ""),
    "추천사유": "실적 개선 기대", "링크": "https://youtu.be/test",
}])
df_related = pd.DataFrame([{
    "채널": "테스트채널B", "업로드일시": "2026-05-11 19:00",
    "제목": "시장 전망 영상", "출연자": "김철수",
    "요약": "매크로 전망 요약.", "링크": "https://youtu.be/test2",
}])

# 5. 테스트 입력 파일 저장 (타임스탬프로 잠금 충돌 방지)
import time as _time
out = os.path.join(BASE, "output", f"test_input_{int(_time.time())}.xlsx")
with pd.ExcelWriter(out, engine="openpyxl") as w:
    df_stocks.to_excel(w, sheet_name="종목추천",       index=False)
    df_related.to_excel(w, sheet_name="투자관련",       index=False)
    df_brief.to_excel(w,  sheet_name="종합",            index=False)
    df_full.to_excel(w,   sheet_name="종합(분기포함)",  index=False)
print(f"테스트 입력 저장: {out}")

# 6. format_excel.py 실행
fmt_script = os.path.join(BASE, "format_excel.py")
ret = subprocess.run([sys.executable, fmt_script, out], cwd=BASE)
if ret.returncode != 0:
    print("[ERR] format_excel.py 실패")
    sys.exit(1)

# 7. 결과 파일 오픈
result = out.replace(".xlsx", "_fmt.xlsx")
if os.path.exists(result):
    print(f"[OK] 결과 파일: {result}")
    os.startfile(result)
else:
    print("[ERR] 결과 파일이 생성되지 않았습니다.")
