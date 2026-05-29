"""
════════════════════════════════════════════════════════════════
  재무스캐너  —  매출·영업이익 성장 & 상승여력
════════════════════════════════════════════════════════════════

■ 목적
  전년(25A) 대비 올해(26E) 실적이 개선되면서,
  증권사 목표주가 대비 현재가에 상승여력이 있는 종목을 발굴한다.
  단순 이익 성장뿐 아니라 매출도 함께 성장하는 구조적 개선주를 선별.

────────────────────────────────────────────────────────────────
■ 스크리닝 조건  (OR 적용 / 임계값은 config.py 에서 수정)
────────────────────────────────────────────────────────────────

  [검출 방식]
  ① 데이터 유효성 검사 (필수·고정) 를 통과한 뒤,
  ② 아래 조건 중 하나라도 만족하면 후보군에 편입한다.
  ③ 후보군 전체에 종합점수를 산출하여 내림차순 정렬한다.

  → OR 방식이므로 조건이 엄격해도 그물이 넓게 쳐진다.
    점수가 낮은 종목은 자연스럽게 하위로 밀리므로
    config.py 의 임계값은 "최소 관심 기준선" 역할을 한다.

  [필수 조건 — 데이터 유효성]
   · FnGuide 에서 25A / 26E 영업이익 모두 수집 가능한 종목

  [OR 조건 — 하나라도 만족하면 편입]

  ① 영업이익 절대 성장
       26E 영업이익 > 25A 영업이익
       → 이익이 한 푼이라도 늘어야 관심 대상.

  ② 영업이익 증가율 최소  [config: MIN_OP_PROFIT_GROWTH = 20.0 %]
       (26E 영업이익 - 25A 영업이익) / |25A 영업이익| × 100 ≥ 설정값
       → 시장 평균(5~10%)을 크게 상회하는 이익 성장 종목.
         흑자전환 종목은 증가율이 수백%이므로 자동 통과.
         설정값을 None 으로 바꾸면 이 조건 비활성화.

  ③ 매출 증가율 최소  [config: MIN_REVENUE_GROWTH = 10.0 %]
       (26E 매출 - 25A 매출) / |25A 매출| × 100 ≥ 설정값
       → 매출이 성장하는 종목 (이익만 늘고 매출 정체 = 비용 절감 효과).
         설정값을 None 으로 바꾸면 이 조건 비활성화.

  ④ 상승여력 최소  [config: MIN_UPSIDE_PCT = 15.0 %]
       (증권사 목표주가 - 현재가) / 현재가 × 100 ≥ 설정값
       → 현재가 대비 목표주가 괴리율이 큰 저평가 종목.
         목표주가가 없으면 이 조건에서는 편입 안 됨.

  ⑤ PSR 상한  [config: MAX_PSR = 2.0]
       시가총액(억) / 26E 매출액(억) ≤ 설정값
       → 매출 대비 시총이 낮아 실적 개선 시 반등 여력이 큰 종목.
         설정값을 None 으로 바꾸면 이 조건 비활성화.

  ⑥ PER 상한  [config: MAX_PER = 25.0]
       주가 / 주당순이익 ≤ 설정값
       → 이익 대비 주가가 낮은 밸류에이션 매력 종목.
         설정값을 None 으로 바꾸면 이 조건 비활성화.

────────────────────────────────────────────────────────────────
■ 종합점수 산출  (0 ~ 100점 / 가중 합산)
────────────────────────────────────────────────────────────────

  점수 = 각 지표를 0~100으로 정규화 후 가중치 합산.
  범위를 벗어난 극단값은 0점(하한) 또는 100점(상한)으로 처리.

  지표                  가중치   방향         정규화 범위
  ─────────────────────────────────────────────────────────
  영업익증가율(%)        20 %    높을수록↑    0 ~ 500 %
  매출증가율(%)          15 %    높을수록↑    0 ~ 100 %
  차기Q영업익YoY(%)      10 %    높을수록↑    0 ~ 300 %  (없으면 0점, 패널티 없음)
  PER/12M PER            20 %    낮을수록↑    5 ~ 40 배
  PEG                    10 %    낮을수록↑    0 ~ 3      (성장률≤0이면 0점)
  상승여력1(%)           10 %    높을수록↑    0 ~ 80 %   (S-RIM 적정주가 기준)
  상승여력2(%)           10 %    높을수록↑    0 ~ 80 %   (증권사 최고 목표주가 기준)
  PSR                     5 %    낮을수록↑    0 ~ 5 배
  ─────────────────────────────────────────────────────────
  합계                  100 %

────────────────────────────────────────────────────────────────
■ 출력 Excel 구성
────────────────────────────────────────────────────────────────

  시트 1  「종합」         : 전체 통과 종목 / 종합점수 내림차순
  시트 2  「★AND조건」    : 상승여력2·PER·영업익증가율 AND 조건 동시 충족 종목
  시트 3~ 「★메가테마」   : config.MEGA_THEMES 에 정의된 테마별 필터 시트
  시트 N~ 「업종명」       : 업종별 분리 / 동일하게 종합점수 내림차순
  마지막  「★눌림목」     : 7_눌림목_스캐너 최신 결과 자동 병합

  컬럼 순서:
    종목명 · 티커 · 업종 · 종합점수
    현재가 · 적정주가(S-RIM) · 상승여력1(%) · 증권사목표주가 · 상승여력2(%)
    25년매출(억) · 26년매출E(억) · 매출증가율(%)
    25년영업익(억) · 26년영업익E(억) · 영업익증가율(%)
    [분기 실적 — 실행일 기준 동적 분기명, 예: 1Q25/1Q26/2Q25/2Q26E]
    전년Q매출 · 최근Q매출 · Q매출YoY(%)
    전년Q영업익 · 최근Q영업익 · Q영업익YoY(%)
    전년QE매출 · 차기QE매출 · 차기Q매출YoY(%)
    전년QE영업익 · 차기QE영업익 · 차기Q영업익YoY(%)
    PER · 12M PER · 12M PER/PER · PSR · PBR · PFR

────────────────────────────────────────────────────────────────
■ 데이터 출처
────────────────────────────────────────────────────────────────

  FnGuide SVD_Main.asp  : 매출·영업이익 25A(실적) / 26E(추정)
                          분기 실적 (Net Quarter 테이블, 최근·차기)
                          목표주가·PER·12M PER·BPS·ROE·업종(FICS)
  네이버 금융 모바일 API : 현재가
  네이버 금융 메인 페이지: 최신 분기 기준 PBR
  DART OpenAPI           : 영업활동현금흐름·CAPEX → TTM FCF → PFR
                          (DART_API_KEY 설정 시 활성화)
  FinanceDataReader      : KRX 전종목 유니버스·업종·시가총액

════════════════════════════════════════════════════════════════
"""

import io
import json
import re
import time
import warnings
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Optional
from xml.etree import ElementTree as ET

import concurrent.futures
import threading
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
from bs4 import BeautifulSoup
from rich.console import Console
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn,
)
from rich.table import Column

warnings.filterwarnings("ignore")

import config
import highlight_excel

console = Console(force_terminal=True, legacy_windows=False)

# ── HTTP Session ──────────────────────────────────────────────────────────────────
http_session = requests.Session()
retries = Retry(total=3, backoff_factor=0.3, status_forcelist=[500, 502, 503, 504])
adapter = HTTPAdapter(pool_connections=20, pool_maxsize=20, max_retries=retries)
http_session.mount("http://", adapter)
http_session.mount("https://", adapter)


# ── 유틸 ──────────────────────────────────────────────────────────────────────

def _parse_num(text: str) -> Optional[float]:
    if not text:
        return None
    text = text.strip().replace(",", "").replace(" ", "")
    if text in ("", "-", "N/A", "NA", "—", "N/A"):
        return None
    try:
        return float(re.sub(r"[^\d.\-]", "", text))
    except ValueError:
        return None


def _get_soup(url: str, params: dict = None, encoding: str = "utf-8") -> Optional[BeautifulSoup]:
    try:
        res = http_session.get(url, headers=config.HEADERS, params=params, timeout=config.TIMEOUT)
        res.raise_for_status()
        time.sleep(config.DELAY)
        res.encoding = encoding
        return BeautifulSoup(res.text, "lxml")
    except Exception:
        return None


def _fetch_prev_year_close(code: str) -> Optional[float]:
    """ACTUAL_YEAR 마지막 거래일 종가 조회 (FinanceDataReader)."""
    try:
        import FinanceDataReader as fdr
        year = config.ACTUAL_YEAR
        df = fdr.DataReader(code, f"{year}-12-01", f"{year}-12-31")
        if not df.empty and "Close" in df.columns:
            return float(df["Close"].iloc[-1])
    except Exception:
        pass
    return None


# ── DART OpenAPI — corp_code 매핑 & FCF ─────────────────────────────────────
#
# DART API 키가 config.DART_API_KEY에 설정된 경우에만 동작.
# corpCode.xml: 전 상장사 corp_code 일괄 다운로드 (스캔 시작 시 1회)
# fnlttSinglAcntAll: 단일회사 전체 재무제표 (현금흐름표 포함)

_DART_CACHE = config.BASE_DIR / "dart_corp_codes.json"
_DART_FCF_CACHE = config.BASE_DIR / "dart_fcf_cache.json"
_DART_FCF_CACHE_DAYS = 7   # FCF 캐시 유효기간 (일) — 분기 공시 주기 고려
_fcf_cache: dict[str, dict] = {}   # {"종목코드": {"fcf": ..., "ts": ...}}


def _build_dart_corp_code_map() -> dict[str, str]:
    """DART corpCode.xml 다운로드 → {6자리 종목코드: 8자리 corp_code} 반환.

    성공 시 로컬 캐시(dart_corp_codes.json)에 저장.
    DART 서버 연결 실패 시 캐시에서 로드해 PFR 계산을 유지.
    """
    if not config.DART_API_KEY:
        return {}

    def _parse(content: bytes) -> dict[str, str]:
        zf = zipfile.ZipFile(io.BytesIO(content))
        root = ET.fromstring(zf.read(zf.namelist()[0]))
        return {
            (item.findtext("stock_code") or "").strip(): (item.findtext("corp_code") or "").strip()
            for item in root.findall("list")
            if len((item.findtext("stock_code") or "").strip()) == 6
        }

    # ── 다운로드 시도 ──────────────────────────────────────────
    for attempt in range(3):
        try:
            res = requests.get(
                "https://opendart.fss.or.kr/api/corpCode.xml",
                params={"crtfc_key": config.DART_API_KEY},
                headers={"User-Agent": config.HEADERS["User-Agent"]},
                timeout=30,
            )
            if res.status_code == 200:
                mapping = _parse(res.content)
                _DART_CACHE.write_text(json.dumps(mapping, ensure_ascii=False), encoding="utf-8")
                return mapping
        except Exception as e:
            if attempt < 2:
                wait = 2 ** attempt
                console.print(f"[yellow]DART corp_code 재시도 ({attempt + 1}/3), {wait}s 대기: {e}[/yellow]")
                time.sleep(wait)
            else:
                console.print(f"[yellow]DART 서버 연결 실패 — 캐시 확인 중...[/yellow]")

    # ── 캐시 폴백 ─────────────────────────────────────────────
    if _DART_CACHE.exists():
        try:
            mapping = json.loads(_DART_CACHE.read_text(encoding="utf-8"))
            mtime = datetime.fromtimestamp(_DART_CACHE.stat().st_mtime).strftime("%Y-%m-%d")
            console.print(f"[dim yellow]DART 캐시 사용 ({mtime} 기준, {len(mapping)}개)[/dim yellow]")
            return mapping
        except Exception:
            pass

    console.print("[yellow]DART corp_code 매핑 실패 — PFR 컬럼 비활성[/yellow]")
    return {}


def _fetch_dart_data(corp_code: str, stock_code: str = "") -> dict:
    """DART Annual(ACTUAL_YEAR) + Q1(ESTIMATE_YEAR) → TTM FCF (억원).

    P&L 분기 데이터는 FnGuide Net Quarter 테이블(_parse_fn_quarterly)로 대체.
    TTM FCF = Annual_FCF − Q1_prev_FCF + Q1_cur_FCF
    """
    empty = {"pfr_fcf": None}
    if not corp_code or not config.DART_API_KEY:
        return empty

    # 캐시 확인 — 유효기간 내면 API 호출 생략
    if stock_code and stock_code in _fcf_cache:
        cached = _fcf_cache[stock_code]
        cache_age = (datetime.now() - datetime.fromisoformat(cached["ts"])).days
        if cache_age < _DART_FCF_CACHE_DAYS:
            return {"pfr_fcf": cached["fcf"]}

    def _fetch_report(bsns_year: int, reprt_code: str) -> Optional[list]:
        for fs_div in ("CFS", "OFS"):
            try:
                r = http_session.get(
                    "https://opendart.fss.or.kr/api/fnlttSinglAcntAll.json",
                    params={
                        "crtfc_key":  config.DART_API_KEY,
                        "corp_code":  corp_code,
                        "bsns_year":  str(bsns_year),
                        "reprt_code": reprt_code,
                        "fs_div":     fs_div,
                    },
                    timeout=config.TIMEOUT,
                )
                d = r.json()
                if d.get("status") == "000" and d.get("list"):
                    return d["list"]
            except Exception:
                pass
        return None

    def _to_float(s: str) -> Optional[float]:
        s = s.replace(",", "").replace(" ", "")
        try:
            return float(s) if s else None
        except ValueError:
            return None

    def _parse_cf(items: list) -> dict:
        """현금흐름표에서 OCF·CAPEX만 파싱."""
        d: dict = {}
        for item in items:
            if "현금흐름표" not in item.get("sj_nm", ""):
                continue
            nm  = item.get("account_nm", "")
            cur = _to_float(item.get("thstrm_amount", ""))
            prv = _to_float(item.get("frmtrm_amount", ""))
            if re.search(r"영업활동.*현금흐름", nm) and "ocf_cur" not in d:
                d["ocf_cur"] = cur
                d["ocf_prv"] = prv
            elif "유형자산" in nm and "취득" in nm and "capex_cur" not in d:
                d["capex_cur"] = abs(cur or 0)
                d["capex_prv"] = abs(prv or 0)
        return d

    # Annual (ACTUAL_YEAR)
    a     = _parse_cf(_fetch_report(config.ACTUAL_YEAR, "11011") or [])
    ocf_a = a.get("ocf_cur")
    cap_a = a.get("capex_cur", 0)
    fcf_ann = (ocf_a - cap_a) / 1e8 if ocf_a is not None else None
    result  = {"pfr_fcf": round(fcf_ann, 0) if fcf_ann is not None else None}

    # Q1 (ESTIMATE_YEAR) → TTM if available
    q1_items = _fetch_report(config.ESTIMATE_YEAR, "11013")
    if q1_items:
        q     = _parse_cf(q1_items)
        ocf_q = q.get("ocf_cur")
        ocf_qp = q.get("ocf_prv")
        cap_q  = q.get("capex_cur", 0)
        cap_qp = q.get("capex_prv", 0)
        if ocf_a is not None and ocf_q is not None and ocf_qp is not None:
            fcf_ttm = (ocf_a - cap_a) - (ocf_qp - cap_qp) + (ocf_q - cap_q)
            result["pfr_fcf"] = round(fcf_ttm / 1e8, 0)

    # 캐시에 결과 저장
    if stock_code:
        _fcf_cache[stock_code] = {
            "fcf": result["pfr_fcf"],
            "ts": datetime.now().isoformat(),
        }

    return result


# ── 네이버 금융 메인 페이지 — PBR + 목표주가 ────────────────────────────────
#
# /item/main.naver 의 per_table 은 정적 HTML 로 최신 분기 BPS 기준 PBR 을 제공.
# 목표주가는 JS 동적 로딩이라 정적 파싱 불가 → FnGuide 컨센서스로 폴백.
# analyst.naver(개별 증권사 최고 목표주가) 요청을 이 요청으로 대체해
# 전체 요청 수를 유지한다.

def _fetch_naver_main(code: str) -> tuple[Optional[float], Optional[float]]:
    """네이버 금융 메인 페이지에서 PBR(최신 분기 BPS 기준)과 목표주가 반환.
    목표주가는 동적 로딩이므로 None 반환 → 호출부에서 FnGuide 컨센서스로 폴백.
    """
    soup = _get_soup(
        "https://finance.naver.com/item/main.naver",
        params={"code": code},
        encoding="euc-kr",
    )
    if soup is None:
        return None, None

    pbr: Optional[float] = None
    table = soup.find("table", class_="per_table")
    if table:
        for tr in table.find_all("tr"):
            cells = [td.get_text(strip=True) for td in tr.find_all(["th", "td"])]
            if len(cells) >= 2 and cells[0].startswith("PBR"):
                m = re.search(r"[\d,]+\.?\d*", cells[1])
                if m:
                    pbr = _parse_num(m.group())
                break

    return pbr, None  # 목표주가는 정적 HTML 에 없으므로 None


# ── FnGuide SVD_Main.asp 파싱 ────────────────────────────────────────────────
#
# SVD_Consensus.asp 의 재무 테이블은 JavaScript 로 동적 로딩되어 requests 로는
# 수집 불가. 대신 SVD_Main.asp 의 정적 HTML 에 25A/26E 연간 컨센서스 테이블이
# 포함되어 있어 여기서 매출·영업이익·목표주가·추정PER 을 한 번에 수집한다.

def _parse_fn_main(soup: BeautifulSoup) -> dict:
    """FnGuide SVD_Main.asp 에서 매출·영업이익(25A/26E)·목표주가·추정PER·업종 파싱."""
    result = {
        "rev_25a": None, "rev_26e": None,
        "op_25a":  None, "op_26e":  None, "op_27e": None,
        "target_price": None,
        "per": None, "per_12m": None, "per_sector": None,
        "pbr": None,
        "bps_26e": None, "roe_26e": None,
        "sector": None,
    }

    # 업종: FICS 세부 분류 우선, 없으면 KRX(코스피/코스닥) 광역 분류 사용
    # 예: "KSE  코스피 화학 코스피 화학 | FICS  개인생활용품 | K200 | NXT"
    # → FICS="개인생활용품"  (화장품 회사가 화학과 분리됨)
    stxt = soup.select_one("p.stxt_group")
    if stxt:
        txt = stxt.get_text(" ", strip=True)
        fics_m = re.search(r"\|\s*FICS\s+(.+?)\s*\|", txt)
        if fics_m:
            result["sector"] = fics_m.group(1).strip()
        else:
            krx_m = re.search(r"(?:코스피|코스닥)\s+(.+?)(?=\s+(?:코스피|코스닥)|\s*\|)", txt)
            if krx_m:
                result["sector"] = krx_m.group(1).strip()

    # PER / 12M PER: tip_in 링크(<a id="h_per/h_12m">) 다음 <dd> 에서 추출
    #   PER    = 전일 주가 / 최근 결산 EPS  (trailing)
    #   12M PER= 전일 주가 / 12개월 Forward EPS
    for aid, key in [("h_per", "per"), ("h_12m", "per_12m"), ("h_u_per", "per_sector"), ("h_pbr", "pbr")]:
        a_tag = soup.find("a", id=aid)
        if a_tag:
            dd_tag = a_tag.find_next_sibling("dd")
            if dd_tag:
                result[key] = _parse_num(dd_tag.get_text(strip=True))

    actual_year    = str(config.ACTUAL_YEAR)        # "2025"
    estimate_year  = str(config.ESTIMATE_YEAR)      # "2026"
    estimate_year2 = str(config.ESTIMATE_YEAR + 1)  # "2027"

    def _cells(row):
        return [c.get_text(strip=True) for c in row.find_all(["th", "td"])]

    for table in soup.find_all("table"):
        rows = table.find_all("tr")
        if len(rows) < 2:
            continue
        header = _cells(rows[0])

        # ① 투자의견/목표주가/EPS/PER/추정기관수 요약 테이블 — 컨센서스 평균 목표주가
        if "목표주가" in header and "투자의견" in header and len(rows) >= 2:
            data = _cells(rows[1])
            try:
                result["target_price"] = _parse_num(data[header.index("목표주가")])
            except (ValueError, IndexError):
                pass
            continue

        # ② IFRS(연결) Annual 재무 테이블 — 2열 헤더 구조
        #    rows[0] = ['IFRS(연결)', 'Annual', 'Net Quarter']  (colspan 포함)
        #    rows[1] = ['2023/12', '2024/12', '2025/12', '2026/12(E)', '2025/06', ...]
        #    rows[2+] = 데이터 행 (label, val1, val2, ...)
        if len(rows) < 3:
            continue
        if not (header and "IFRS(연결)" in header[0] and "Annual" in header):
            continue

        year_row = _cells(rows[1])
        col_actual = col_estimate = col_estimate2 = None
        for i, h in enumerate(year_row):
            # 25A: 실제년도 포함, 추정 마커 없음 — 첫 번째 매칭만 사용 (분기 중복 방지)
            if col_actual is None and actual_year in h and "(E)" not in h and "추정" not in h and "(P)" not in h:
                col_actual = i
            # 26E: 추정년도 포함, (E) 또는 추정 마커 있음, 잠정(P) 아님 — 첫 번째 매칭만
            if col_estimate is None and estimate_year in h and ("(E)" in h or "추정" in h) and "(P)" not in h:
                col_estimate = i
            # 27E: 2년 후 추정 — PEG 분모(2년 CAGR) 계산용
            if col_estimate2 is None and estimate_year2 in h and ("(E)" in h or "추정" in h) and "(P)" not in h:
                col_estimate2 = i

        if col_actual is None and col_estimate is None:
            continue

        # 데이터 행: cells[0]=label, cells[1+]=연도별 값 (year_row 인덱스 +1)
        for row in rows[2:]:
            cells = _cells(row)
            if not cells:
                continue
            label = cells[0]

            def _val(ci):
                if ci is None:
                    return None
                idx = ci + 1  # label 열(0)만큼 오프셋
                return _parse_num(cells[idx]) if idx < len(cells) else None

            if re.search(r"^매출(액)?$", label):
                result["rev_25a"] = result["rev_25a"] or _val(col_actual)
                result["rev_26e"] = result["rev_26e"] or _val(col_estimate)
            elif label == "영업이익":  # "영업이익(발표기준)", "영업이익률" 제외
                result["op_25a"] = result["op_25a"] or _val(col_actual)
                result["op_26e"] = result["op_26e"] or _val(col_estimate)
                result["op_27e"] = result["op_27e"] or _val(col_estimate2)
            elif label.startswith("ROE"):
                result["roe_26e"] = result["roe_26e"] or _val(col_estimate)
            elif label.startswith("BPS"):
                result["bps_26e"] = result["bps_26e"] or _val(col_estimate)

        if result["op_25a"] or result["rev_25a"]:
            break  # 연결 테이블 파싱 완료

    return result


# ── FnGuide SVD_Invest.asp — 개별 증권사 최고 목표주가 ────────────────────────

def _parse_fn_invest(soup: BeautifulSoup) -> Optional[float]:
    """FnGuide SVD_Invest.asp 개별 증권사 목표주가 리스트에서 최댓값 반환."""
    if soup is None:
        return None
    targets: list[float] = []
    for tbl in soup.find_all("table"):
        rows = tbl.find_all("tr")
        if len(rows) < 2:
            continue
        header = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]
        if "목표주가" not in header:
            continue
        if not any(kw in header for kw in ("증권사", "리서치", "기관")):
            continue
        tp_idx = header.index("목표주가")
        for row in rows[1:]:
            cells = [c.get_text(strip=True) for c in row.find_all(["th", "td"])]
            if tp_idx < len(cells):
                v = _parse_num(cells[tp_idx])
                if v and v > 0:
                    targets.append(v)
        if targets:
            break
    return max(targets) if targets else None


# ── FnGuide Net Quarter 분기 파싱 ───────────────────────────────────────────
#
# SVD_Main.asp의 IFRS(연결) Net Quarter 테이블 (정적 HTML, 추가 요청 불필요)
# 열 구조: 과거 4~5분기 실적(actual) + 최근 잠정(P) + 향후 추정(E) 2~3분기
# → 최근 분기(cur_q) & 차기 분기 추정(nxt_q_e) + 각각의 전년 동기 반환

def _parse_fn_quarterly(soup: BeautifulSoup) -> dict:
    """IFRS(연결) Net Quarter 테이블에서 최근·차기 분기 매출·영업이익 파싱.

    반환 키 (단위: 억원, FnGuide 원본값 그대로):
      cur_q_rev / cur_q_rev_yoy    — 최근 분기 매출 / 전년 동기 매출
      nxt_q_rev_e / nxt_q_rev_yoy  — 차기 분기 추정 매출 / 전년 동기 매출
      cur_q_op  / cur_q_op_yoy     — 최근 분기 영업익 / 전년 동기 영업익
      nxt_q_op_e / nxt_q_op_yoy   — 차기 분기 추정 영업익 / 전년 동기 영업익
    """
    empty = {
        "cur_q_rev": None, "cur_q_rev_yoy": None,
        "nxt_q_rev_e": None, "nxt_q_rev_yoy": None,
        "cur_q_op":  None, "cur_q_op_yoy":  None,
        "nxt_q_op_e": None, "nxt_q_op_yoy": None,
    }
    if not soup:
        return empty

    # IFRS(연결) Net Quarter 전용 테이블 탐색
    # — "IFRS(연결) | Annual | Net Quarter" 복합 테이블(Table 10)과 구분하기 위해
    #   첫 번째 행에 "Annual" 이 없고 두 번째 행에 "(P)" 또는 "(E)" 가 있는 테이블을 선택
    target = None
    for t in soup.find_all("table"):
        rows = t.find_all("tr")
        if len(rows) < 3:
            continue
        first = [c.get_text(strip=True) for c in rows[0].find_all(["th", "td"])]
        if not (any("IFRS(연결)" in x for x in first) and any("Net Quarter" in x for x in first)):
            continue
        if any("Annual" in x for x in first):
            continue  # 복합 Annual+Net Quarter 테이블 제외
        second = " ".join(c.get_text(strip=True) for c in rows[1].find_all(["th", "td"]))
        if "(P)" in second or "(E)" in second:
            target = t
            break
    if target is None:
        return empty

    rows = target.find_all("tr")
    if len(rows) < 3:
        return empty

    # 두 번째 tr — 분기 컬럼 라벨 파싱
    hdr_cells = [c.get_text(strip=True) for c in rows[1].find_all(["th", "td"])]
    col_info: list[tuple[Optional[str], str]] = []  # [(YYYY/MM, type)]
    for h in hdr_cells:
        ym = re.search(r"(\d{4}/\d{2})", h)
        ym_str = ym.group(1) if ym else None
        if "(P)" in h:
            typ = "provisional"
        elif "(E)" in h:
            typ = "estimate"
        else:
            typ = "actual"
        col_info.append((ym_str, typ))

    # 최근 분기: 마지막 실적/잠정 컬럼
    cur_idx: Optional[int] = None
    for i, (ym, typ) in enumerate(col_info):
        if typ in ("actual", "provisional") and ym:
            cur_idx = i

    # 차기 분기: 첫 번째 추정 컬럼
    nxt_idx: Optional[int] = next(
        (i for i, (ym, typ) in enumerate(col_info) if typ == "estimate" and ym), None
    )

    if cur_idx is None:
        return empty

    def _prev_yr_idx(idx: int) -> Optional[int]:
        """해당 컬럼의 전년 동기 컬럼 인덱스 반환."""
        ym = col_info[idx][0]
        if not ym:
            return None
        yr, mo = ym.split("/")
        prev_ym = f"{int(yr) - 1}/{mo}"
        return next((i for i, (y, _) in enumerate(col_info) if y == prev_ym), None)

    prev_cur_idx = _prev_yr_idx(cur_idx)
    prev_nxt_idx = _prev_yr_idx(nxt_idx) if nxt_idx is not None else None

    result = {**empty}

    for row in rows[2:]:
        cells = [c.get_text(strip=True) for c in row.find_all(["th", "td"])]
        if not cells:
            continue
        label = cells[0]

        def _v(idx: Optional[int]) -> Optional[float]:
            if idx is None:
                return None
            di = idx + 1  # 데이터는 cells[1] 부터 (cells[0] = 행 라벨)
            return _parse_num(cells[di]) if di < len(cells) else None

        if label == "매출액":
            result["cur_q_rev"]     = _v(cur_idx)
            result["cur_q_rev_yoy"] = _v(prev_cur_idx)
            result["nxt_q_rev_e"]   = _v(nxt_idx)
            result["nxt_q_rev_yoy"] = _v(prev_nxt_idx)
        elif label == "영업이익":
            result["cur_q_op"]      = _v(cur_idx)
            result["cur_q_op_yoy"]  = _v(prev_cur_idx)
            result["nxt_q_op_e"]    = _v(nxt_idx)
            result["nxt_q_op_yoy"]  = _v(prev_nxt_idx)

    return result


# ── 종목별 데이터 수집 ────────────────────────────────────────────────────────

def get_comprehensive_data(
    code: str,
    market_cap_eok: float = 0.0,
) -> Optional[dict]:
    """FnGuide SVD_Main.asp + 네이버 모바일 API에서 종목 데이터 수집.

    market_cap_eok: FDR 에서 받은 시가총액(억원) — PSR 계산에 사용.
    PFR(FCF 기반)은 필터 통과 후 _fetch_dart_data 로 후처리. (_marcap_eok 에 저장)

    데이터 소스:
    - FnGuide SVD_Main.asp: 매출·영업이익 25A/26E, 목표주가, 추정PER (정적 HTML)
    - 네이버 모바일 API  : 현재가 (JSON)
    - DART API          : 영업활동현금흐름·CAPEX → FCF → PFR
    """
    # 네이버 현재가·PBR을 FnGuide와 병렬 요청
    # FnGuide가 가장 오래 걸리므로 그 시간 동안 Naver 2건을 미리 처리해 소요 시간 단축
    price_box      = [None]
    pbr_naver_box  = [None]
    prev_close_box = [None]

    def _naver_price():
        try:
            res = http_session.get(
                f"https://m.stock.naver.com/api/stock/{code}/basic",
                headers={**config.HEADERS, "Referer": "https://m.stock.naver.com/"},
                timeout=config.TIMEOUT,
            )
            if res.status_code == 200:
                time.sleep(config.DELAY)
                price_box[0] = _parse_num(res.json().get("closePrice", ""))
        except Exception:
            pass

    def _naver_pbr():
        pbr_naver_box[0], _ = _fetch_naver_main(code)

    def _prev_close():
        prev_close_box[0] = _fetch_prev_year_close(code)

    max_target_box = [None]

    def _fn_invest():
        s = _get_soup(
            "https://comp.fnguide.com/SVO2/asp/SVD_Invest.asp",
            params={"pGB": "1", "gicode": f"A{code}", "cID": "",
                    "MenuYn": "Y", "ReportGB": "D", "NewMenuID": "133", "stkGb": "701"},
        )
        max_target_box[0] = _parse_fn_invest(s)

    t_price      = threading.Thread(target=_naver_price, daemon=True)
    t_pbr        = threading.Thread(target=_naver_pbr,   daemon=True)
    t_prev_close = threading.Thread(target=_prev_close,  daemon=True)
    t_invest     = threading.Thread(target=_fn_invest,   daemon=True)
    t_price.start()
    t_pbr.start()
    t_prev_close.start()
    t_invest.start()

    soup_fn = _get_soup(
        "https://comp.fnguide.com/SVO2/asp/SVD_Main.asp",
        params={
            "pGB": "1", "gicode": f"A{code}", "cID": "",
            "MenuYn": "Y", "ReportGB": "D", "NewMenuID": "Y", "stkGb": "701",
        },
    )

    t_price.join()
    t_pbr.join()
    t_prev_close.join()
    t_invest.join()

    if soup_fn is None:
        return None

    fn = _parse_fn_main(soup_fn)
    if fn["op_25a"] is None or fn["op_26e"] is None:
        return None

    price = price_box[0]

    per    = fn["per"]
    per_12m = fn["per_12m"]
    per_ratio = round(per_12m / per, 2) if per and per_12m and per != 0 else None

    # 적정주가 — S-RIM (사경인 방식)
    #   적정주가 = BPS(26E) × ROE(26E) / r
    #   ROE가 r 보다 작으면 BPS 이하로 산출 (정상)
    #   BPS·ROE 데이터 없으면 FnGuide 증권사 목표주가로 폴백
    bps = fn["bps_26e"]
    roe = fn["roe_26e"]   # % 단위, e.g. 45.56
    r   = config.SRIM_REQUIRED_RETURN
    if bps and roe and bps > 0:
        target = max(0, round(bps * (roe / 100) / r))
    else:
        target = fn["target_price"]

    pbr = pbr_naver_box[0] or fn["pbr"]
    broker_target = max_target_box[0] or fn["target_price"]
    upside1 = round((target        - price) / price * 100, 1) if price and target        is not None and price > 0 else None
    upside2 = round((broker_target - price) / price * 100, 1) if price and broker_target is not None and price > 0 else None

    op_25a, op_26e   = fn["op_25a"],  fn["op_26e"]
    op_27e           = fn["op_27e"]
    rev_25a, rev_26e = fn["rev_25a"], fn["rev_26e"]

    op_growth  = round((op_26e - op_25a) / abs(op_25a) * 100, 1) if op_25a and op_25a != 0 else None
    rev_growth = round((rev_26e - rev_25a) / abs(rev_25a) * 100, 1) if rev_25a and rev_25a != 0 and rev_26e else None

    # PEG 분모 — 시점 일관성 확보
    # 12M PER(선행) 분자에 맞춰 분모도 선행 다년 CAGR을 사용:
    #   27E 데이터 있음 → 2년 CAGR (25A→27E):  ((27E/25A)^0.5 - 1) × 100
    #   27E 없음        → 1년 성장률 (26E vs 25A) 으로 폴백
    if op_25a and op_25a > 0 and op_27e and op_27e > 0:
        _peg_growth: Optional[float] = round(((op_27e / op_25a) ** 0.5 - 1) * 100, 1)
    else:
        _peg_growth = op_growth

    # PSR = 시가총액(억) / 26E 매출(억)
    psr = round(market_cap_eok / rev_26e, 2) if market_cap_eok and rev_26e and rev_26e > 0 else None

    # 분기 비교 — FnGuide Net Quarter 테이블 (SVD_Main.asp, 추가 요청 없음)
    # PFR은 필터 통과 후 DART 후처리에서 계산 (_marcap_eok 로 전달)
    fq = _parse_fn_quarterly(soup_fn)

    def _yoy(cur: Optional[float], prv: Optional[float]) -> Optional[float]:
        if cur is not None and prv and prv != 0:
            return round((cur - prv) / abs(prv) * 100, 1)
        return None

    q_rev_yoy     = _yoy(fq["cur_q_rev"],   fq["cur_q_rev_yoy"])
    q_op_yoy      = _yoy(fq["cur_q_op"],    fq["cur_q_op_yoy"])
    nxt_rev_yoy   = _yoy(fq["nxt_q_rev_e"], fq["nxt_q_rev_yoy"])
    nxt_op_yoy    = _yoy(fq["nxt_q_op_e"],  fq["nxt_q_op_yoy"])

    return {
        "작년종가":       prev_close_box[0],
        "현재가":         price,
        "적정주가":       target,
        "상승여력1(%)":   upside1,
        "증권사목표주가": broker_target,
        "상승여력2(%)":   upside2,
        "25년매출(억)":       rev_25a,
        "26년매출E(억)":      rev_26e,
        "매출증가율(%)":      rev_growth,
        "25년영업익(억)":     op_25a,
        "26년영업익E(억)":    op_26e,
        "영업익증가율(%)":    op_growth,
        "최근Q매출(억)":      fq["cur_q_rev"],
        "전년Q매출(억)":      fq["cur_q_rev_yoy"],
        "Q매출YoY(%)":        q_rev_yoy,
        "최근Q영업익(억)":    fq["cur_q_op"],
        "전년Q영업익(억)":    fq["cur_q_op_yoy"],
        "Q영업익YoY(%)":      q_op_yoy,
        "차기QE매출(억)":     fq["nxt_q_rev_e"],
        "전년QE매출(억)":     fq["nxt_q_rev_yoy"],
        "차기Q매출YoY(%)":    nxt_rev_yoy,
        "차기QE영업익(억)":   fq["nxt_q_op_e"],
        "전년QE영업익(억)":   fq["nxt_q_op_yoy"],
        "차기Q영업익YoY(%)":  nxt_op_yoy,
        "PER":                per,
        "12M PER":            per_12m,
        "12M PER/PER":        per_ratio,
        "PSR":                psr,
        "PBR":                pbr,
        "PFR":                None,   # DART 후처리에서 채움
        "_sector":            fn["sector"],
        "_marcap_eok":        market_cap_eok,
        "_peg_growth":        _peg_growth,  # PEG 분모용 다년 CAGR (reindex 시 제거)
    }


# ── 스크리닝 필터 ────────────────────────────────────────────────────────────

def _passes_filter(data: dict) -> bool:
    """OR 조건 — 활성화된 조건 중 하나라도 만족하면 편입."""
    op_25a     = data["25년영업익(억)"]
    op_26e     = data["26년영업익E(억)"]
    op_growth  = data["영업익증가율(%)"]
    rev_growth = data["매출증가율(%)"]
    upside1    = data["상승여력1(%)"]
    upside2    = data["상승여력2(%)"]
    per        = data["PER"]
    psr        = data["PSR"]

    # 필수: 영업이익 데이터가 있어야 점수 산출 가능
    if op_25a is None or op_26e is None:
        return False

    # OR 조건 목록 — 하나라도 True면 편입
    checks = [
        op_26e > op_25a,                                                          # ① 영업이익 절대 성장
        config.MIN_OP_PROFIT_GROWTH is not None
            and op_growth is not None
            and op_growth >= config.MIN_OP_PROFIT_GROWTH,                         # ② 영업이익 증가율
        config.MIN_REVENUE_GROWTH is not None
            and rev_growth is not None
            and rev_growth >= config.MIN_REVENUE_GROWTH,                          # ③ 매출 증가율
        (upside1 is not None and upside1 >= config.MIN_UPSIDE_PCT)
            or (upside2 is not None and upside2 >= config.MIN_UPSIDE_PCT),        # ④ 상승여력
        config.MAX_PSR is not None and psr is not None and psr <= config.MAX_PSR, # ⑤ PSR 상한
        config.MAX_PER is not None and per is not None and per <= config.MAX_PER, # ⑥ PER 상한
    ]
    return any(checks)


# ── 종합점수 산출 ─────────────────────────────────────────────────────────────
#
# [가중치 설계]
#   영업익증가율   20%  — 연간 실적개선 핵심 지표
#   매출증가율     15%  — 이익 성장의 구조적 근거
#   차기Q영업익YoY 10%  — 분기 선행 모멘텀 (데이터 없으면 0점, 패널티 없음)
#   PER/12M PER    20%  — 이익 대비 밸류에이션 (inverse)
#   PEG            10%  — 성장 대비 밸류에이션 (inverse, 성장률≤0이면 0점)
#   상승여력1/2    10%  — S-RIM / 증권사 목표주가 대비 저평가 (각 10%)
#   PSR             5%  — 매출 대비 밸류에이션 보조 (inverse)
#
# [정규화 방식]
#   각 지표를 합리적 상·하한으로 clip 후 0~100 선형 정규화.
#   상한 초과는 100점, 하한 미달·NaN은 0점(역방향은 하한값=hi로 대체).

def _compute_scores(df: pd.DataFrame) -> pd.Series:
    """종합점수(0~100) 산출. 8개 지표 가중 합산."""

    def _norm(s: pd.Series, lo: float, hi: float, inverse: bool = False) -> pd.Series:
        filled = s.fillna(lo if not inverse else hi)
        clipped = filled.clip(lo, hi)
        if hi == lo:
            return pd.Series(50.0, index=s.index)
        scaled = (clipped - lo) / (hi - lo) * 100
        return (100 - scaled) if inverse else scaled

    # 12M PER 없는 종목은 PER 로 대체
    per_for_score = df["12M PER"].combine_first(df["PER"])

    # PEG 인라인 계산: 12M PER / _peg_growth (다년 CAGR 또는 1년 성장률)
    # 성장률 ≤ 0 이면 PEG 의미 없음 → NaN → inverse fill(hi=3) → 0점
    peg_growth = pd.to_numeric(df.get("_peg_growth", pd.Series(dtype=float)), errors="coerce")
    peg_raw = (per_for_score / peg_growth.where(peg_growth > 0))

    # 차기Q 영업익 YoY: 데이터 없는 종목은 NaN → 0점 (패널티 없음, 보너스 구조)
    nxt_q_op = pd.to_numeric(df.get("차기Q영업익YoY(%)", pd.Series(dtype=float)), errors="coerce")

    score = (
        _norm(df["영업익증가율(%)"],  lo=0,   hi=500)              * 0.20
      + _norm(df["매출증가율(%)"],    lo=0,   hi=100)              * 0.15
      + _norm(nxt_q_op,              lo=0,   hi=300)              * 0.10
      + _norm(per_for_score,         lo=5,   hi=40, inverse=True) * 0.20
      + _norm(peg_raw,               lo=0,   hi=3,  inverse=True) * 0.10
      + _norm(df["상승여력1(%)"],     lo=0,   hi=80)               * 0.10
      + _norm(df["상승여력2(%)"],     lo=0,   hi=80)               * 0.10
      + _norm(df["PSR"],             lo=0,   hi=5,  inverse=True) * 0.05
    )
    return score.round(1)


# ── Excel 컬럼별 숫자 포맷 정의 ──────────────────────────────────────────────

# 양수: 초록, 음수: 빨강 표기하는 % 컬럼 포맷
_FMT_PCT_SIGN  = '[Color10]+#,##0.0;[Red]-#,##0.0;0.0'
_FMT_SLOPE     = '[Color10]+#,##0.00;[Red]-#,##0.00;0.00'

_COL_FORMATS: dict[str, str] = {
    # ── 실적개선 컬럼 ──
    "종합점수":           "#,##0.0",
    "작년종가":           "#,##0",
    "현재가":             "#,##0",
    "현재가/작년종가":    "#,##0.000",
    "적정주가":           "#,##0",
    "상승여력1(%)":        _FMT_PCT_SIGN,
    "증권사목표주가":      "#,##0",
    "상승여력2(%)":        _FMT_PCT_SIGN,
    "25년매출(억)":       "#,##0",
    "26년매출E(억)":      "#,##0",
    "매출증가율(%)":      _FMT_PCT_SIGN,
    "25년영업익(억)":     "#,##0",
    "26년영업익E(억)":    "#,##0",
    "영업익증가율(%)":    _FMT_PCT_SIGN,
    "최근Q매출(억)":      "#,##0",
    "전년Q매출(억)":      "#,##0",
    "Q매출YoY(%)":        _FMT_PCT_SIGN,
    "최근Q영업익(억)":    "#,##0",
    "전년Q영업익(억)":    "#,##0",
    "Q영업익YoY(%)":      _FMT_PCT_SIGN,
    "차기QE매출(억)":     "#,##0",
    "전년QE매출(억)":     "#,##0",
    "차기Q매출YoY(%)":    _FMT_PCT_SIGN,
    "차기QE영업익(억)":   "#,##0",
    "전년QE영업익(억)":   "#,##0",
    "차기Q영업익YoY(%)":  _FMT_PCT_SIGN,
    "PER":                "#,##0.0",
    "12M PER":            "#,##0.0",
    "12M PER/PER":        "#,##0.00",
    "PSR":                "#,##0.00",
    "PBR":                "#,##0.00",
    "PFR":                "#,##0.00",
    "영익증가율-주가상승율(%p)": _FMT_PCT_SIGN,
    "PEG":                "#,##0.00",
    # ── 눌림목 컬럼 ──
    "점수":               "#,##0",
    "돌파가":             "#,##0",
    "고점":               "#,##0",
    "눌림(%)":            "#,##0.1",
    "박스상단":           "#,##0",
    "박스하단":           "#,##0",
    "박스범위(%)":        "#,##0.1",
    "거래량배율":         "#,##0.0",
    "시총(억)":           "#,##0",
    "52주저점대비(%)":    "#,##0.1",
    "일봉MA60기울기(%)":  _FMT_SLOPE,
    "주봉MA20기울기(%)":  _FMT_SLOPE,
    "월봉MA5기울기(%)":   _FMT_SLOPE,
}

# 섹션별 헤더 배경색
_HEADER_FILLS: dict[str, str] = {
    # ── 실적개선 컬럼 ──
    "종목명":             "1F4E79",
    "티커":               "1F4E79",
    "업종":               "1F4E79",
    "종합점수":           "7B3F00",
    "작년종가":           "1A4A2A",
    "현재가":             "375623",
    "현재가/작년종가":    "2A6B3A",
    "적정주가":           "375623",
    "상승여력1(%)":        "375623",
    "증권사목표주가":      "1D6A3A",
    "상승여력2(%)":        "1D6A3A",
    "25년매출(억)":       "7F4F24",
    "26년매출E(억)":      "7F4F24",
    "매출증가율(%)":      "7F4F24",
    "25년영업익(억)":     "1F3864",
    "26년영업익E(억)":    "1F3864",
    "영업익증가율(%)":    "1F3864",
    "최근Q매출(억)":      "5C3D11",
    "전년Q매출(억)":      "5C3D11",
    "Q매출YoY(%)":        "5C3D11",
    "최근Q영업익(억)":    "1A3A5C",
    "전년Q영업익(억)":    "1A3A5C",
    "Q영업익YoY(%)":      "1A3A5C",
    "차기QE매출(억)":     "7A2C0E",
    "전년QE매출(억)":     "7A2C0E",
    "차기Q매출YoY(%)":    "7A2C0E",
    "차기QE영업익(억)":   "0E2C5C",
    "전년QE영업익(억)":   "0E2C5C",
    "차기Q영업익YoY(%)":  "0E2C5C",
    "PER":                "4C4C4C",
    "12M PER":            "2E4057",
    "12M PER/PER":        "2E4057",
    "PSR":                "4C4C4C",
    "PBR":                "4C4C4C",
    "PFR":                "4C4C4C",
    "영익증가율-주가상승율(%p)": "5C2D0E",
    "PEG":                "2E2E5C",
    # ── 눌림목 컬럼 ──
    "점수":               "7B3F00",
    "유형":               "7B3F00",
    "돌파일":             "375623",
    "돌파가":             "375623",
    "고점":               "375623",
    "눌림(%)":            "375623",
    "박스상단":           "1F3864",
    "박스하단":           "1F3864",
    "박스범위(%)":        "1F3864",
    "거래량배율":         "7F4F24",
    "시총(억)":           "4C4C4C",
    "52주저점대비(%)":    "5C3317",
    "일봉MA60기울기(%)":  "2E4057",
    "주봉MA20기울기(%)":  "2E4057",
    "월봉MA5기울기(%)":   "2E4057",
}

# ── 분기 라벨 계산 ───────────────────────────────────────────────────────────

def _get_quarter_labels() -> dict[str, str]:
    """오늘 날짜 기준 분기 컬럼 헤더 rename 맵 반환.

    예) 2026-05-09 (Q2 진행중) → 최근 보고 분기 = 26년1Q, 차기 추정 = 26년2Q
    """
    from datetime import date
    today = date.today()
    month, year = today.month, today.year
    cur_q = (month - 1) // 3 + 1          # 현재 진행 중인 분기
    if cur_q == 1:
        rep_q, rep_y = 4, year - 1        # Q1이면 직전 연도 Q4가 최근 보고
    else:
        rep_q, rep_y = cur_q - 1, year    # 그 외엔 직전 분기

    nxt_q = rep_q % 4 + 1
    nxt_y = rep_y + (1 if rep_q == 4 else 0)

    ry, py = str(rep_y)[-2:], str(rep_y - 1)[-2:]
    ny, npy = str(nxt_y)[-2:], str(nxt_y - 1)[-2:]

    return {
        "전년Q매출(억)":     f"{py}년{rep_q}Q\n매출(억)",
        "최근Q매출(억)":     f"{ry}년{rep_q}Q\n매출(억)",
        "Q매출YoY(%)":       f"{rep_q}Q매출\n증가율(%)",
        "전년Q영업익(억)":   f"{py}년{rep_q}Q\n영업익(억)",
        "최근Q영업익(억)":   f"{ry}년{rep_q}Q\n영업익(억)",
        "Q영업익YoY(%)":     f"{rep_q}Q영업익\n증가율(%)",
        "전년QE매출(억)":    f"{npy}년{nxt_q}Q\n매출(억)",
        "차기QE매출(억)":    f"{ny}년{nxt_q}QE\n매출(억)",
        "차기Q매출YoY(%)":   f"{nxt_q}Q매출\n증가율(%)",
        "전년QE영업익(억)":  f"{npy}년{nxt_q}Q\n영업익(억)",
        "차기QE영업익(억)":  f"{ny}년{nxt_q}QE\n영업익(억)",
        "차기Q영업익YoY(%)": f"{nxt_q}Q영업익\n증가율(%)",
    }


# ── Excel 서식 적용 ───────────────────────────────────────────────────────────

def _format_excel(path: Path) -> None:
    """저장된 Excel 파일에 섹션별 헤더·컬럼별 숫자 포맷·하이퍼링크·열 너비 적용."""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Alignment, Font, PatternFill

        wb = load_workbook(path)

        def _cw(val):
            return sum(2 if ord(c) > 127 else 1 for c in str(val or ""))

        for ws in wb.worksheets:
            # 헤더 행 컬럼별 색상·폰트
            col_fmt_map: dict[int, str] = {}   # col_index → number_format
            name_col = ticker_col = None

            for cell in ws[1]:
                col_name = cell.value or ""
                fill_hex = _HEADER_FILLS.get(col_name, "1F4E79")
                cell.fill      = PatternFill("solid", fgColor=fill_hex)
                cell.font      = Font(bold=True, color="FFFFFF")
                cell.alignment = Alignment(horizontal="center", wrap_text=True)

                if col_name in _COL_FORMATS:
                    col_fmt_map[cell.column] = _COL_FORMATS[col_name]
                if col_name == "종목명": name_col   = cell.column
                if col_name == "티커":   ticker_col = cell.column

            # 데이터 행: 숫자 포맷 + 정렬
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center")
                    fmt = col_fmt_map.get(cell.column)
                    if fmt and isinstance(cell.value, (int, float)):
                        cell.number_format = fmt
                    elif isinstance(cell.value, int):
                        cell.number_format = "#,##0"

            # 유형 컬럼 조건부 색상 (눌림목 시트용)
            type_col = None
            for cell in ws[1]:
                if cell.value == "유형":
                    type_col = cell.column
                    break
            if type_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    cell = row[type_col - 1]
                    if cell.value == "바닥탈출":
                        cell.font = Font(bold=True, color="FF6B35")
                        cell.fill = PatternFill("solid", fgColor="3D1A00")
                    elif cell.value == "우상향":
                        cell.font = Font(bold=True, color="3FB950")
                        cell.fill = PatternFill("solid", fgColor="0D2818")

            # 종목명 → 네이버 하이퍼링크 / 티커 → FnGuide 하이퍼링크
            if name_col and ticker_col:
                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    nc = row[name_col   - 1]
                    tc = row[ticker_col - 1]
                    if tc.value:
                        code6 = str(tc.value).zfill(6)
                        nc.hyperlink = (
                            f"https://finance.naver.com/item/main.naver"
                            f"?code={code6}"
                        )
                        nc.font = Font(color="0563C1", underline="single", bold=True)
                        tc.hyperlink = (
                            f"https://comp.fnguide.com/SVO2/ASP/SVD_main.asp"
                            f"?pGB=1&gicode=A{code6}&cID=AA"
                            f"&MenuYn=Y&ReportGB=&NewMenuID=11&stkGb=&strResearchYN="
                        )
                        tc.font = Font(color="0563C1", underline="single")

            # 헤더 행 높이 — 긴 컬럼명 2줄 표시를 위해 고정 (넉넉하게 60)
            ws.row_dimensions[1].height = 60

            # 분기 컬럼 헤더를 오늘 날짜 기준 실제 분기명으로 교체
            q_rename = _get_quarter_labels()
            for cell in ws[1]:
                if cell.value in q_rename:
                    cell.value = q_rename[cell.value]

            # 헤더에 줄바꿈 삽입 — 단위 괄호 앞에서 끊어 2줄로 표시
            _HEADER_BREAKS = {
                "종합점수":           "종합\n점수",
                "작년종가":           "작년\n종가",
                "현재가":             "현재\n가",
                "현재가/작년종가":    "현재가/\n작년종가",
                "적정주가":           "적정\n주가",
                "상승여력1(%)":       "상승여력\n1(%)",
                "증권사목표주가":     "증권사목표주가\n(최고)",
                "상승여력2(%)":       "상승여력\n2(%)",
                "25년매출(억)":       "25년\n매출(억)",
                "26년매출E(억)":      "26년\n매출E(억)",
                "매출증가율(%)":      "매출\n증가율(%)",
                "25년영업익(억)":     "25년\n영업익(억)",
                "26년영업익E(억)":    "26년\n영업익E(억)",
                "영업익증가율(%)":    "영업익\n증가율(%)",
                "12M PER":            "12M\nPER",
                "12M PER/PER":        "12M PER\n/PER",
                "영익증가율-주가상승율(%p)": "영익증가율\n-주가상승율(%p)",
                "PEG":                "PEG\n(12MPER/이익증가율)",
            }
            for cell in ws[1]:
                if cell.value in _HEADER_BREAKS:
                    cell.value = _HEADER_BREAKS[cell.value]

            # 열 너비 자동 조정 — 데이터 및 헤더 셀 너비 고려 (마진 추가)
            for col in ws.columns:
                col_letter = col[0].column_letter
                col_idx    = col[0].column          # 1-based
                
                # 헤더 텍스트(줄바꿈 포함) 중 가장 긴 줄의 너비
                header_lines = str(col[0].value or "").split('\n')
                hw = max([_cw(line) for line in header_lines] + [0])
                
                # 데이터 행의 최대 너비
                data_w = max((_cw(cell.value) for cell in col[1:]), default=0)
                
                # 최종 기준 너비
                content_w = max(data_w, hw)

                if col_letter == "A":               # 종목명: 데이터 기준
                    width = min(max(content_w + 2, 12), 22)
                elif col_letter == "B":             # 티커: 고정 8
                    width = 8
                elif col_letter == "C":             # 업종: 약간 좁게
                    width = min(max(int(content_w * 0.8) + 1, 8), 16)
                elif 4 <= col_idx <= 35:            # 숫자 컬럼 (마진 2 추가, 최대 18)
                    width = min(max(content_w + 2, 10), 18)
                else:
                    width = min(max(content_w + 2, 8), 18)
                ws.column_dimensions[col_letter].width = width

            # 헤더 행 고정 + 자동 필터
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions

        wb.save(path)
    except Exception as e:
        console.print(f"[yellow]Excel 서식 적용 실패: {e}[/yellow]")


# ── 눌림목 스캐너 최신 결과 읽기 ────────────────────────────────────────────

def _get_latest_눌림목_df() -> pd.DataFrame:
    scanner_dir = Path(__file__).parent.parent / "7_눌림목_스캐너" / "output"
    files = sorted(scanner_dir.glob("눌림목스캐너_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        return pd.DataFrame()
    try:
        return pd.read_excel(files[0], sheet_name="눌림목")
    except Exception:
        return pd.DataFrame()


# ── 메인 ──────────────────────────────────────────────────────────────────────

def run_invest_scouter_with_sheets():
    import FinanceDataReader as fdr

    _start_time = datetime.now()
    console.print()
    console.print("[bold cyan]재무스캐너[/bold cyan] — 매출·영업익 성장 & 상승여력")
    console.print()

    # output 디렉토리를 시작 시점에 미리 생성
    config.OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    stocks = fdr.StockListing('KRX')

    # 업종·시총 컬럼 정리 (FDR 버전에 따라 컬럼명 상이)
    sector_col = next((c for c in stocks.columns if c in ("Sector", "Industry", "업종")), None)
    marcap_col = next((c for c in stocks.columns if c in ("Marcap", "MarCap", "시가총액")), None)
    keep_cols  = ["Code", "Name"] + ([sector_col] if sector_col else []) + ([marcap_col] if marcap_col else [])
    stocks = stocks[keep_cols].rename(columns={
        **({"업종": "업종"} if sector_col is None else {sector_col: "업종"}),
        **({"시총": "시총"} if marcap_col is None else {marcap_col: "시총"}),
    })
    if "업종" not in stocks.columns:
        stocks["업종"] = "기타"
    stocks["업종"] = stocks["업종"].fillna("기타")
    if "시총" not in stocks.columns:
        stocks["시총"] = 0
    stocks["시총"] = pd.to_numeric(stocks["시총"], errors="coerce").fillna(0)

    # 시총 최소 필터
    if config.MIN_MARKET_CAP > 0:
        stocks = stocks[stocks["시총"] / 1e8 >= config.MIN_MARKET_CAP]

    target_stocks = stocks.head(config.TEST_STOCKS) if config.TEST_STOCKS else stocks
    items = list(target_stocks.itertuples())

    # 메가 테마 names 에 등록된 종목명 → 재무 필터 면제 대상
    _mega_theme_names: set[str] = set()
    for theme_cfg in config.MEGA_THEMES.values():
        for n in theme_cfg.get("names", []):
            _mega_theme_names.add(n)

    console.print(f"[dim]대상 종목: {len(items)}개 | 테스트 모드: {'ON' if config.TEST_STOCKS else 'OFF'}[/dim]")
    console.print(f"[dim]메가테마 면제 종목: {len(_mega_theme_names)}개[/dim]")

    # DART corp_code 매핑 (PFR 계산용 — API 키 없으면 빈 dict, PFR=빈칸)
    dart_corp_map: dict[str, str] = {}
    if config.DART_API_KEY:
        console.print("[dim]DART corp_code 매핑 빌드 중...[/dim]")
        dart_corp_map = _build_dart_corp_code_map()
        console.print(f"[dim]DART 매핑 완료: {len(dart_corp_map)}개 종목[/dim]")
    else:
        console.print("[dim yellow]DART_API_KEY 미설정 — PFR 컬럼 비활성[/dim yellow]")

    console.print()

    final_list = []

    with Progress(
        SpinnerColumn(),
        TextColumn("{task.description}", table_column=Column(width=20, no_wrap=True)),
        BarColumn(),
        MofNCompleteColumn(),
        TimeElapsedColumn(),
        TextColumn("남은"),
        TimeRemainingColumn(),
        console=console,
        transient=False,
    ) as progress:
        task = progress.add_task("분석 시작...", total=len(items))

        def process_stock(row):
            code         = str(row.Code).zfill(6)
            name         = row.Name
            sector       = getattr(row, "업종", "기타") or "기타"
            marcap_eok   = getattr(row, "시총", 0) / 1e8

            progress.update(task, description=f"종목 수집 중...")
            data = get_comprehensive_data(code, market_cap_eok=marcap_eok)

            if not data:
                return None

            # 메가테마 종목 → 재무 필터 면제 / 일반 종목 → 필터 통과 필수
            is_theme = any(n in name or name in n for n in _mega_theme_names)
            if not is_theme and not _passes_filter(data):
                return None

            fn_sector = data.pop("_sector", None)
            actual_sector = (
                config.SECTOR_OVERRIDES.get(code)
                or fn_sector
                or sector
                or "기타"
            )
            result = {"종목명": name, "티커": code, "업종": actual_sector, **data}
            if is_theme and not _passes_filter({**data, "_sector": fn_sector}):
                result["_theme_bypass"] = True   # 필터 면제 표시
            return result

        # 동시 실행 스레드 수 — FnGuide/Naver 차단 방지를 위해 6으로 제한
        max_workers = 6
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = [executor.submit(process_stock, row) for row in items]
            
            for future in concurrent.futures.as_completed(futures):
                try:
                    result = future.result()
                    if result:
                        final_list.append(result)
                except Exception:
                    pass
                progress.advance(task)

    # ── DART 후처리: 필터 통과 종목만 PFR 계산 ──────────────────────────────────
    if config.DART_API_KEY and dart_corp_map and final_list:
        # FCF 캐시 로드
        global _fcf_cache
        if _DART_FCF_CACHE.exists():
            try:
                _fcf_cache = json.loads(_DART_FCF_CACHE.read_text(encoding="utf-8"))
                console.print(f"[dim]DART FCF 캐시 로드: {len(_fcf_cache)}개 종목[/dim]")
            except Exception:
                _fcf_cache = {}

        # 캐시로 커버 가능한 종목 수 계산
        cached_count = sum(
            1 for item in final_list
            if item["티커"] in _fcf_cache
            and (datetime.now() - datetime.fromisoformat(_fcf_cache[item["티커"]]["ts"])).days < _DART_FCF_CACHE_DAYS
        )
        fetch_count = len(final_list) - cached_count
        console.print(f"[dim]DART FCF 수집: {len(final_list)}개 중 캐시 {cached_count}개 / API {fetch_count}개[/dim]")
        with Progress(
            SpinnerColumn(),
            TextColumn("{task.description}", table_column=Column(width=20, no_wrap=True)),
            BarColumn(),
            MofNCompleteColumn(),
            TimeElapsedColumn(),
            TextColumn("남은"),
            TimeRemainingColumn(),
            console=console,
            transient=True,
        ) as prog:
            task2 = prog.add_task("PFR 계산...", total=len(final_list))
            
            def process_dart(item):
                corp_code = dart_corp_map.get(item["티커"], "")
                marcap    = item.pop("_marcap_eok", 0)
                if corp_code and marcap:
                    dart  = _fetch_dart_data(corp_code, stock_code=item["티커"])
                    fcf   = dart["pfr_fcf"]
                    if fcf and fcf != 0:
                        item["PFR"] = round(marcap / fcf, 2)
                return True

            with concurrent.futures.ThreadPoolExecutor(max_workers=10) as executor:
                dart_futures = [executor.submit(process_dart, item) for item in final_list]
                for f in concurrent.futures.as_completed(dart_futures):
                    try:
                        f.result()
                    except Exception:
                        pass
                    prog.advance(task2)

        # FCF 캐시 저장
        try:
            _DART_FCF_CACHE.write_text(json.dumps(_fcf_cache, ensure_ascii=False), encoding="utf-8")
            console.print(f"[dim]DART FCF 캐시 저장: {len(_fcf_cache)}개 종목[/dim]")
        except Exception:
            pass
    else:
        for item in final_list:
            item.pop("_marcap_eok", None)

    # ── 결과 저장 ──
    COLUMN_ORDER = [
        "종목명", "티커", "업종",
        "종합점수",
        "작년종가", "현재가", "현재가/작년종가", "적정주가", "상승여력1(%)", "증권사목표주가", "상승여력2(%)",
        "25년매출(억)", "26년매출E(억)", "매출증가율(%)",
        "25년영업익(억)", "26년영업익E(억)", "영업익증가율(%)",
        "전년Q매출(억)", "최근Q매출(억)", "Q매출YoY(%)",
        "전년Q영업익(억)", "최근Q영업익(억)", "Q영업익YoY(%)",
        "전년QE매출(억)", "차기QE매출(억)", "차기Q매출YoY(%)",
        "전년QE영업익(억)", "차기QE영업익(억)", "차기Q영업익YoY(%)",
        "PER", "12M PER", "12M PER/PER", "PSR", "PBR", "PFR",
        "영익증가율-주가상승율(%p)",
        "PEG",
    ]

    _QUARTER_COLS = {
        "전년Q매출(억)", "최근Q매출(억)", "Q매출YoY(%)",
        "전년Q영업익(억)", "최근Q영업익(억)", "Q영업익YoY(%)",
        "전년QE매출(억)", "차기QE매출(억)", "차기Q매출YoY(%)",
        "전년QE영업익(억)", "차기QE영업익(억)", "차기Q영업익YoY(%)",
    }
    COLUMN_ORDER_BRIEF = [c for c in COLUMN_ORDER if c not in _QUARTER_COLS]

    ts        = datetime.now().strftime("%Y_%m%d_%H%M")
    file_name = config.OUTPUT_DIR / f"재무스캐너_{ts}.xlsx"

    if not final_list:
        console.print("\n[yellow]조건에 맞는 종목이 없습니다. 빈 Excel을 저장합니다.[/yellow]")
        pd.DataFrame(columns=COLUMN_ORDER_BRIEF).to_excel(file_name, index=False)
        console.print(f"[dim]저장 위치: {file_name}[/dim]\n")
        return

    df = pd.DataFrame(final_list)

    # 종합점수 산출 후 내림차순 정렬
    df["종합점수"] = _compute_scores(df)
    df = df.sort_values(by="종합점수", ascending=False).reset_index(drop=True)

    # 작년 종가 기반 파생 컬럼
    prev_close_s = pd.to_numeric(df["작년종가"], errors="coerce")
    cur_price_s  = pd.to_numeric(df["현재가"],   errors="coerce")
    ratio = cur_price_s / prev_close_s.replace(0, float("nan"))
    df["현재가/작년종가"] = ratio.round(3)
    price_chg_pct = (ratio - 1) * 100
    df["영익증가율-주가상승율(%p)"] = (
        pd.to_numeric(df["영업익증가율(%)"], errors="coerce") - price_chg_pct
    ).round(1)

    # PEG = 12M PER(선행 분자) / 다년 CAGR(선행 분모) — 시점 일관성 확보
    # 분모 우선순위: ① 2년 CAGR (27E 데이터 있을 때) ② 1년 성장률(폴백)
    # 성장률이 0 이하면 PEG 무의미 → NaN 처리
    per_12m_s   = pd.to_numeric(df["12M PER"],     errors="coerce")
    peg_growth_s = pd.to_numeric(df["_peg_growth"], errors="coerce")
    df["PEG"] = (per_12m_s / peg_growth_s.where(peg_growth_s > 0)).round(2)

    df = df.reindex(columns=COLUMN_ORDER)

    # 종목명 → 티커 역방향 맵 (메가 테마 종목명 매핑용)
    name_to_code: dict[str, str] = {
        str(row.Name): str(row.Code).zfill(6)
        for row in stocks.itertuples()
    }

    with pd.ExcelWriter(file_name, engine="openpyxl") as writer:
        # ── 1. 종합 시트 (분기 컬럼 제외) ────────────────────────────────────
        df.reindex(columns=COLUMN_ORDER_BRIEF).to_excel(writer, sheet_name="종합", index=False)

        # ── 2. 종합(분기포함) 시트 (전체 컬럼) ───────────────────────────────
        df.to_excel(writer, sheet_name="종합(분기포함)", index=False)

        # ── 3. 메가 테마 시트 ─────────────────────────────────────────────────
        for theme_name, theme_cfg in config.MEGA_THEMES.items():
            keywords = theme_cfg.get("keywords", [])
            names    = theme_cfg.get("names", [])

            theme_codes: set[str] = set()
            for name in names:
                if name in name_to_code:
                    theme_codes.add(name_to_code[name])
                else:
                    for sname, code in name_to_code.items():
                        if name in sname:
                            theme_codes.add(code)
                            break

            mask = pd.Series(False, index=df.index)
            if keywords:
                kw_pat = "|".join(re.escape(k) for k in keywords)
                mask |= df["업종"].str.contains(kw_pat, na=False)
            if theme_codes:
                mask |= df["티커"].isin(theme_codes)

            theme_df = df[mask]
            if theme_df.empty:
                continue

            theme_df.to_excel(writer, sheet_name=f"★{theme_name}"[:31], index=False)

        # ── 4. 업종별 시트 ────────────────────────────────────────────────────
        for ind in df["업종"].dropna().unique():
            sheet_name = (
                str(ind)[:31]
                .replace("/", "_").replace("\\", "_")
                .replace("*", "").replace(":", "")
                .replace("?", "").replace("[", "").replace("]", "")
                or "기타"
            )
            df[df["업종"] == ind].to_excel(writer, sheet_name=sheet_name, index=False)

        # ── 5. 눌림목 스캐너 시트 ─────────────────────────────────────────────
        df_눌림목 = _get_latest_눌림목_df()
        if not df_눌림목.empty:
            df_눌림목.to_excel(writer, sheet_name="★눌림목", index=False)
            console.print(f"  [dim]★눌림목 시트: {len(df_눌림목)}개 종목 추가 (7_눌림목_스캐너 최신)[/dim]")
        else:
            console.print("  [dim yellow]★눌림목 시트: 7_눌림목_스캐너/output/ 파일 없음 — 생략[/dim yellow]")

    _format_excel(file_name)

    console.rule("[bold cyan]하이라이팅 & AND조건 시트[/bold cyan]")
    hl_file = highlight_excel.highlight(file_name)
    console.print()

    # ── 터미널 요약 (상위 10개) ──
    from rich import box
    from rich.table import Table

    top = df.head(10)
    tbl = Table(box=box.SIMPLE_HEAD, header_style="bold cyan", border_style="dim", title="[bold]Top 10 종목[/bold]")
    tbl.add_column("순위", width=4,  justify="right", style="dim")
    tbl.add_column("종목명",   width=14)
    tbl.add_column("업종",     width=16)
    tbl.add_column("종합점수", width=8,  justify="right", style="bold yellow")
    tbl.add_column("영업익증가율(%)", width=14, justify="right")
    tbl.add_column("상승여력2(%)",     width=12, justify="right")
    tbl.add_column("PSR", width=6, justify="right")

    for rank, (_, r) in enumerate(top.iterrows(), 1):
        tbl.add_row(
            str(rank),
            str(r["종목명"]),
            str(r["업종"]),
            f"{r['종합점수']:.1f}",
            f"{r['영업익증가율(%)']:.1f}%" if pd.notna(r["영업익증가율(%)"]) else "-",
            f"{r['상승여력2(%)']:.1f}%"     if pd.notna(r["상승여력2(%)"])     else "-",
            f"{r['PSR']:.2f}"             if pd.notna(r["PSR"])             else "-",
        )

    console.print()
    console.print(tbl)
    console.print(f"\n[green]저장 완료: {file_name}[/green]")
    console.print(
        f"[cyan]통과 종목: {len(df)}개  |  업종: {df['업종'].nunique()}개  |  시트: {df['업종'].nunique() + 1}개[/cyan]\n"
    )

    # 이메일 발송 / summary.txt 저장 (SEND_EMAIL=false 이면 summary.txt 만 저장)
    _send_email(hl_file, df)
    console.print()

    # 실행 시간 출력
    _end_time = datetime.now()
    _elapsed = _end_time - _start_time
    _mins, _secs = divmod(int(_elapsed.total_seconds()), 60)
    console.print(
        f"[bold green]⏱ 시작 {_start_time:%H:%M:%S} → 종료 {_end_time:%H:%M:%S}  "
        f"(총 {_mins}분 {_secs}초)[/bold green]\n"
    )


# ── 이메일 전송 ──────────────────────────────────────────────────────────────

def _send_email(file_name: Path, df: pd.DataFrame) -> None:
    """분석 결과 Excel을 첨부해 수신자 목록으로 전송."""
    import smtplib
    from email.mime.application import MIMEApplication
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText

    ts      = datetime.now().strftime("%Y-%m-%d %H:%M")
    n_stock = len(df)
    n_ind   = df["업종"].nunique() if not df.empty else 0

    # ── 메일 본문 ──
    top10 = df.head(10) if not df.empty else df
    top_lines = []
    for rank, (_, r) in enumerate(top10.iterrows(), 1):
        op_g  = f"{r['영업익증가율(%)']:.1f}%" if pd.notna(r.get("영업익증가율(%)")) else "-"
        up    = f"{r['상승여력2(%)']:.1f}%"     if pd.notna(r.get("상승여력2(%)"))     else "-"
        score = f"{r['종합점수']:.1f}"          if pd.notna(r.get("종합점수"))         else "-"
        top_lines.append(
            f"  {rank:2d}. {r['종목명']:<10}  점수:{score}  영업익증가율:{op_g}  상승여력:{up}"
        )

    q_labels = _get_quarter_labels()
    cur_q_rev  = q_labels.get("최근Q매출(억)",    "최근Q매출")
    cur_q_op   = q_labels.get("최근Q영업익(억)",  "최근Q영업익")
    nxt_q_rev  = q_labels.get("차기QE매출(억)",   "차기QE매출")
    nxt_q_op   = q_labels.get("차기QE영업익(억)", "차기QE영업익")

    body = "\n".join([
        f"{'='*60}",
        f"■ 주의사항 및 면책",
        f"{'='*60}",
        f"  · 본 자료는 투자 참고 목적으로만 제공되며,",
        f"    특정 종목에 대한 매수·매도 권유가 아닙니다.",
        f"  · 투자 판단 및 그에 따른 손익은 전적으로 본인에게 있습니다.",
        f"  · 본 자료는 수신자 개인에게만 제공된 것으로,",
        f"    SNS·커뮤니티·단체방 등 외부 재배포를 금합니다.",
        f"  · 무단 재배포로 인한 법적 책임은 재배포한 당사자에게 있습니다.",
        f"{'='*60}",
        f"",
        f"안녕하세요,",
        f"",
        f"[{ts}] 재무스캐너 분석이 완료되었습니다.",
        f"",
        f"{'='*60}",
        f"■ 분석 결과 요약",
        f"{'='*60}",
        f"  통과 종목 수 : {n_stock}개",
        f"  업종 수      : {n_ind}개",
        f"",
        f"■ 종합점수 상위 {len(top_lines)}개 종목",
        *(top_lines if top_lines else ["  (해당 종목 없음)"]),
        f"",
        f"{'='*60}",
        f"■ 첨부 엑셀 구성",
        f"{'='*60}",
        f"  [종합] 시트",
        f"    전체 통과 종목 / 종합점수 내림차순 정렬",
        f"    분기 실적 컬럼 제외 — 핵심 지표만 표시 (가로 폭 최소화)",
        f"",
        f"  [종합(분기포함)] 시트",
        f"    종합 시트와 동일하되 분기 실적 컬럼 12개 추가",
        f"    최근 보고 분기 + 차기 추정 분기의 매출·영업이익 YoY 포함",
        f"",
        f"  [★AND조건] 시트",
        f"    상승여력2·PER·영업익증가율 조건을 동시 충족하는 종목",
        f"",
        f"  [★메가테마] 시트 (설정된 테마별)",
        f"    조선·방산·로봇·2차전지·AI반도체·원전·전력망 등 주요 테마 종목",
        f"",
        f"  [업종별] 시트 (약 27개)",
        f"    업종별 분리 / 동일하게 종합점수 내림차순",
        f"",
        f"  [★눌림목] 시트",
        f"    7_눌림목_스캐너 최신 결과 자동 병합",
        f"",
        f"  컬럼 구성:",
        f"    종목명·티커·업종·종합점수",
        f"    현재가·적정주가(S-RIM)·상승여력1(%) / 증권사목표주가(최고)·상승여력2(%)",
        f"    25년매출·26년매출E·매출증가율(%)",
        f"    25년영업익·26년영업익E·영업익증가율(%)",
        f"    분기 실적 (최근: {cur_q_rev}·{cur_q_op} / 차기추정: {nxt_q_rev}·{nxt_q_op})",
        f"    PER·12M PER·12M PER/PER·PSR·PBR·PFR(잉여현금흐름)",
        f"",
        f"{'='*60}",
        f"■ 스크리닝 조건 (OR — 하나라도 만족 시 편입)",
        f"{'='*60}",
        f"  ① 영업이익 절대 성장   : 26E > 25A",
        f"  ② 영업이익증가율       : ≥ {config.MIN_OP_PROFIT_GROWTH}%",
        f"  ③ 매출증가율           : ≥ {config.MIN_REVENUE_GROWTH}%",
        f"  ④ 상승여력             : ≥ {config.MIN_UPSIDE_PCT}% (적정주가 또는 증권사 목표주가 기준)",
        f"  ⑤ PSR 상한             : ≤ {config.MAX_PSR}배",
        f"  ⑥ PER 상한             : ≤ {config.MAX_PER}배",
        f"",
        f"{'='*60}",
        f"■ 종합점수 산출 방식 (0~100점)",
        f"{'='*60}",
        f"  영업익증가율(%)        20%  (0~500% 정규화)",
        f"  매출증가율(%)          15%  (0~100% 정규화)",
        f"  차기Q영업익YoY(%)      10%  (0~300% 정규화, NaN=0점 패널티 없음)",
        f"  PER / 12M PER          20%  (5~40배, 낮을수록 고점수)",
        f"  PEG                    10%  (0~3배, 낮을수록 고점수)",
        f"  상승여력1(%)           10%  (0~80% / S-RIM 적정주가 기준)",
        f"  상승여력2(%)           10%  (0~80% / 증권사 최고 목표주가 기준)",
        f"  PSR                     5%  (0~5배, 낮을수록 고점수)",
        f"",
        f"{'='*60}",
        f"■ 동작 방식",
        f"{'='*60}",
        f"  FinanceDataReader      → KRX 전종목 유니버스·업종·시가총액 수집",
        f"  FnGuide SVD_Main.asp   → 연간(25A/26E) + 분기 실적·목표주가·PER·BPS·ROE 파싱",
        f"  네이버 금융 모바일 API → 현재가 조회",
        f"  네이버 금융 메인 페이지→ 최신 분기 PBR 파싱",
        f"  S-RIM(사경인 방식)     → BPS·ROE 기반 적정주가 산출 (없으면 증권사 목표주가 폴백)",
        f"  DART OpenAPI           → 영업활동현금흐름·CAPEX → TTM FCF → PFR 산출",
        f"                           (DART_API_KEY 설정 시 활성화, 필터 통과 종목만 적용)",
        f"",
        f"상세 분석 결과는 첨부 Excel 파일을 확인해주세요.",
        f"",
        f"※ 이 메일은 자동 발송입니다.",
    ])

    try:
        (file_name.parent / "summary.txt").write_text(body, encoding="utf-8")
    except Exception:
        pass

    if not config.SEND_EMAIL:
        console.print("[dim]이메일 발송 생략 (SEND_EMAIL=false)[/dim]")
        return
    if not (config.GMAIL_USER and config.GMAIL_APP_PW):
        console.print("[yellow]발송 계정 미설정 (.env: GMAIL_USER / GMAIL_APP_PASSWORD)[/yellow]")
        return
    if not config.NOTIFY_EMAILS:
        console.print("[yellow]수신자 미설정 (.env: NOTIFY_EMAIL / NOTIFY_EMAIL_WIFE)[/yellow]")
        return

    msg = MIMEMultipart()
    msg["From"]    = config.GMAIL_USER
    msg["To"]      = ", ".join(config.MAIN_EMAILS) if config.MAIN_EMAILS else config.GMAIL_USER
    if config.EXTRA_EMAILS:
        msg["Bcc"] = ", ".join(config.EXTRA_EMAILS)
    msg["Subject"] = f"[재무스캐너] {ts} 스캔 완료 — {n_stock}개 종목"
    msg.attach(MIMEText(body, "plain", "utf-8"))

    # Excel 첨부
    if file_name.exists():
        with open(file_name, "rb") as f:
            part = MIMEApplication(f.read(), Name=file_name.name)
            part["Content-Disposition"] = f'attachment; filename="{file_name.name}"'
            msg.attach(part)

    try:
        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(config.GMAIL_USER, config.GMAIL_APP_PW)
            smtp.sendmail(config.GMAIL_USER, config.NOTIFY_EMAILS, msg.as_string())
        _to_str  = ', '.join(config.MAIN_EMAILS) if config.MAIN_EMAILS else config.GMAIL_USER
        _bcc_str = f" | Bcc: {len(config.EXTRA_EMAILS)}명" if config.EXTRA_EMAILS else ""
        console.print(f"  [green]이메일 발송 완료 — To: {_to_str}{_bcc_str}[/green]")
        if config.EXTRA_EMAILS:
            console.print(f"  [dim]  Bcc: {', '.join(config.EXTRA_EMAILS)}[/dim]")
    except Exception as e:
        console.print(f"  [red]이메일 발송 실패: {e}[/red]")


if __name__ == "__main__":
    run_invest_scouter_with_sheets()
