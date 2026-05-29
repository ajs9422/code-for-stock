"""
================================================
  SQLite 저장소
  - 일별 컨센서스 스냅샷 저장 (시계열 추적)
  - 1개월 전 값 조회 (컨센 변화율 계산)
  - 스크리닝 결과 히스토리 저장
================================================

[테이블 구조]
  consensus_snapshot : 날짜·종목별 컨센서스 수치
  screening_result   : 날짜별 통과 종목 + 점수
"""

import sqlite3
from contextlib import contextmanager
from datetime import date, timedelta
from pathlib import Path
from typing import Optional

from config import DB_PATH
from crawler.fnguide import ConsensusData


# ── DB 연결 ───────────────────────────────────────────────────────────────────

@contextmanager
def _conn():
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    try:
        yield con
        con.commit()
    finally:
        con.close()


# ── 초기화 ────────────────────────────────────────────────────────────────────

def init():
    """테이블이 없으면 생성."""
    with _conn() as con:
        con.executescript("""
        CREATE TABLE IF NOT EXISTS consensus_snapshot (
            snap_date      TEXT NOT NULL,
            code           TEXT NOT NULL,
            name           TEXT,
            market_cap     REAL,
            revenue_25a    REAL,
            revenue_26e    REAL,
            op_profit_25a  REAL,
            op_profit_26e  REAL,
            opm_25a        REAL,
            opm_26e        REAL,
            analyst_count  INTEGER,
            PRIMARY KEY (snap_date, code)
        );

        CREATE TABLE IF NOT EXISTS screening_result (
            run_date        TEXT NOT NULL,
            code            TEXT NOT NULL,
            name            TEXT,
            total_score     REAL,
            revenue_growth  REAL,
            op_profit_growth REAL,
            opm_improvement REAL,
            op_revision_1m  REAL,
            turnaround      INTEGER,
            PRIMARY KEY (run_date, code)
        );
        """)


# ── 스냅샷 저장 ───────────────────────────────────────────────────────────────

def save_snapshot(items: list[ConsensusData], snap_date: date = None):
    """오늘 날짜로 컨센서스 스냅샷 저장 (이미 있으면 REPLACE)."""
    today = str(snap_date or date.today())
    rows = [
        (
            today, d.code, d.name, d.market_cap,
            d.revenue_25a, d.revenue_26e,
            d.op_profit_25a, d.op_profit_26e,
            d.opm_25a, d.opm_26e,
            d.analyst_count,
        )
        for d in items
    ]
    with _conn() as con:
        con.executemany("""
        INSERT OR REPLACE INTO consensus_snapshot
        (snap_date, code, name, market_cap,
         revenue_25a, revenue_26e, op_profit_25a, op_profit_26e,
         opm_25a, opm_26e, analyst_count)
        VALUES (?,?,?,?,?,?,?,?,?,?,?)
        """, rows)


# ── 1개월 전 스냅샷 조회 ──────────────────────────────────────────────────────

def load_1m_ago(codes: list[str]) -> dict[str, Optional[float]]:
    """
    약 30일 전 스냅샷에서 종목별 op_profit_26e 조회.
    반환: {code: op_profit_26e_1m_ago}
    """
    target_date = date.today() - timedelta(days=30)
    # 정확한 날짜가 없을 수 있으므로 ±3일 범위에서 가장 가까운 날짜 사용
    result: dict[str, Optional[float]] = {}

    with _conn() as con:
        for code in codes:
            row = con.execute("""
            SELECT op_profit_26e FROM consensus_snapshot
            WHERE code = ?
              AND snap_date BETWEEN ? AND ?
            ORDER BY ABS(JULIANDAY(snap_date) - JULIANDAY(?))
            LIMIT 1
            """, (
                code,
                str(target_date - timedelta(days=3)),
                str(target_date + timedelta(days=3)),
                str(target_date),
            )).fetchone()
            result[code] = row["op_profit_26e"] if row else None

    return result


def fill_revision(items: list[ConsensusData]) -> list[ConsensusData]:
    """items 리스트에 1개월 전 컨센서스 값을 채워 넣는다."""
    codes = [d.code for d in items]
    hist  = load_1m_ago(codes)
    for d in items:
        d.op_profit_26e_1m_ago = hist.get(d.code)
    return items


def load_3m_ago(codes: list[str]) -> dict[str, Optional[float]]:
    """약 90일 전 스냅샷에서 종목별 op_profit_26e 조회."""
    target_date = date.today() - timedelta(days=90)
    result: dict[str, Optional[float]] = {}
    with _conn() as con:
        for code in codes:
            row = con.execute("""
            SELECT op_profit_26e FROM consensus_snapshot
            WHERE code = ?
              AND snap_date BETWEEN ? AND ?
            ORDER BY ABS(JULIANDAY(snap_date) - JULIANDAY(?))
            LIMIT 1
            """, (
                code,
                str(target_date - timedelta(days=5)),
                str(target_date + timedelta(days=5)),
                str(target_date),
            )).fetchone()
            result[code] = row["op_profit_26e"] if row else None
    return result


def fill_revision_3m(items: list[ConsensusData]) -> list[ConsensusData]:
    """items에 3개월 전 컨센서스 값을 채워 op_revision_3m을 계산한다."""
    codes = [d.code for d in items]
    hist  = load_3m_ago(codes)
    for d in items:
        d.op_profit_26e_3m_ago = hist.get(d.code)
        if d.op_profit_26e is not None and d.op_profit_26e_3m_ago:
            if d.op_profit_26e_3m_ago != 0:
                d.op_revision_3m = (
                    (d.op_profit_26e - d.op_profit_26e_3m_ago)
                    / abs(d.op_profit_26e_3m_ago) * 100
                )
    return items


def load_prev_results(run_date: date = None) -> set[str]:
    """전날 스크리닝 통과 종목 코드 집합 반환 (신규/이탈 비교용)."""
    today = run_date or date.today()
    with _conn() as con:
        rows = con.execute("""
        SELECT DISTINCT code FROM screening_result
        WHERE run_date = (
            SELECT MAX(run_date) FROM screening_result WHERE run_date < ?
        )
        """, (str(today),)).fetchall()
    return {r["code"] for r in rows}


def load_prev_result_names(run_date: date = None) -> dict:
    """전날 스크리닝 통과 종목의 code→name 매핑 반환 (이탈 표시용)."""
    today = run_date or date.today()
    with _conn() as con:
        rows = con.execute("""
        SELECT code, name FROM screening_result
        WHERE run_date = (
            SELECT MAX(run_date) FROM screening_result WHERE run_date < ?
        )
        """, (str(today),)).fetchall()
    return {r["code"]: r["name"] for r in rows}


# ── 스크리닝 결과 저장 ────────────────────────────────────────────────────────

def save_results(ranked: list[tuple], run_date: date = None):
    """스크리닝 통과 종목과 점수를 저장."""
    today = str(run_date or date.today())
    rows = []
    for data, score_obj in ranked:
        rows.append((
            today, data.code, data.name,
            round(score_obj.total, 2),
            data.revenue_growth,
            data.op_profit_growth,
            data.opm_improvement,
            data.op_revision_1m,
            int(data.turnaround),
        ))
    with _conn() as con:
        con.executemany("""
        INSERT OR REPLACE INTO screening_result
        (run_date, code, name, total_score,
         revenue_growth, op_profit_growth, opm_improvement,
         op_revision_1m, turnaround)
        VALUES (?,?,?,?,?,?,?,?,?)
        """, rows)


# ── 히스토리 조회 ─────────────────────────────────────────────────────────────

def load_results(run_date: date = None) -> list[dict]:
    """특정 날짜(기본: 오늘)의 스크리닝 결과 조회."""
    today = str(run_date or date.today())
    with _conn() as con:
        rows = con.execute("""
        SELECT * FROM screening_result WHERE run_date = ?
        ORDER BY total_score DESC
        """, (today,)).fetchall()
    return [dict(r) for r in rows]


def load_consensus_history(code: str, days: int = 60) -> list[dict]:
    """특정 종목의 최근 N일 컨센서스 시계열 조회."""
    since = str(date.today() - timedelta(days=days))
    with _conn() as con:
        rows = con.execute("""
        SELECT snap_date, op_profit_26e, revenue_26e, opm_26e
        FROM consensus_snapshot
        WHERE code = ? AND snap_date >= ?
        ORDER BY snap_date
        """, (code, since)).fetchall()
    return [dict(r) for r in rows]
