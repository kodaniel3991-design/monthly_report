"""
월차보고 시스템 - 데이터베이스 관리
SQLite 기반, 단일 파일 저장
"""

import sqlite3
import os
from pathlib import Path

DB_PATH = Path(__file__).parent / "data" / "monthly_report.db"


def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    """DB 초기화 - 테이블 생성"""
    conn = get_conn()
    cur = conn.cursor()

    # ── 1. 사업계획 (연 1회 입력, 월별×항목) ────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS annual_plan (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        year        INTEGER NOT NULL,
        month       INTEGER NOT NULL,        -- 1~12
        item_code   TEXT    NOT NULL,        -- 항목코드 (e.g. 'sales_rkm')
        item_name   TEXT    NOT NULL,        -- 항목명 (e.g. '매출액_RKM')
        value       REAL    DEFAULT 0,
        updated_at  TEXT    DEFAULT (datetime('now','localtime')),
        UNIQUE(year, month, item_code)
    )""")

    # ── 2. 손익 실적 (매월 입력) ─────────────────────────────────────────
    # 공장별 4개: gimhae(김해), busan(부산), ulsan(울산), gimhae2(김해2)
    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_pl (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        year            INTEGER NOT NULL,
        month           INTEGER NOT NULL,
        -- 판매수량 (대)
        qty_gimhae      REAL DEFAULT 0,
        qty_busan       REAL DEFAULT 0,
        qty_ulsan       REAL DEFAULT 0,
        qty_gimhae2     REAL DEFAULT 0,
        -- 생산금액 (천원)
        prod_gimhae     REAL DEFAULT 0,
        prod_busan      REAL DEFAULT 0,
        prod_ulsan      REAL DEFAULT 0,
        prod_gimhae2    REAL DEFAULT 0,
        -- 매출액 - 생산품 (천원)
        sales_prod_gimhae   REAL DEFAULT 0,
        sales_prod_busan    REAL DEFAULT 0,
        sales_prod_ulsan    REAL DEFAULT 0,
        sales_prod_gimhae2  REAL DEFAULT 0,
        -- 매출액 - 외주품 (천원)
        sales_out_gimhae    REAL DEFAULT 0,
        sales_out_busan     REAL DEFAULT 0,
        sales_out_ulsan     REAL DEFAULT 0,
        sales_out_gimhae2   REAL DEFAULT 0,
        -- 변동비: 재고증감차 (천원)
        inv_diff_gimhae     REAL DEFAULT 0,
        inv_diff_busan      REAL DEFAULT 0,
        inv_diff_ulsan      REAL DEFAULT 0,
        inv_diff_gimhae2    REAL DEFAULT 0,
        -- 변동비: 재료비 (천원)
        material_gimhae     REAL DEFAULT 0,
        material_busan      REAL DEFAULT 0,
        material_ulsan      REAL DEFAULT 0,
        material_gimhae2    REAL DEFAULT 0,
        -- 변동비: 제조경비 - 복리후생비
        mfg_welfare_gimhae  REAL DEFAULT 0,
        mfg_welfare_busan   REAL DEFAULT 0,
        mfg_welfare_ulsan   REAL DEFAULT 0,
        mfg_welfare_gimhae2 REAL DEFAULT 0,
        -- 변동비: 제조경비 - 전력비
        mfg_power_gimhae    REAL DEFAULT 0,
        mfg_power_busan     REAL DEFAULT 0,
        mfg_power_ulsan     REAL DEFAULT 0,
        mfg_power_gimhae2   REAL DEFAULT 0,
        -- 변동비: 제조경비 - 운반비
        mfg_trans_gimhae    REAL DEFAULT 0,
        mfg_trans_busan     REAL DEFAULT 0,
        mfg_trans_ulsan     REAL DEFAULT 0,
        mfg_trans_gimhae2   REAL DEFAULT 0,
        -- 변동비: 제조경비 - 수선비
        mfg_repair_gimhae   REAL DEFAULT 0,
        mfg_repair_busan    REAL DEFAULT 0,
        mfg_repair_ulsan    REAL DEFAULT 0,
        mfg_repair_gimhae2  REAL DEFAULT 0,
        -- 변동비: 제조경비 - 소모품비
        mfg_supplies_gimhae  REAL DEFAULT 0,
        mfg_supplies_busan   REAL DEFAULT 0,
        mfg_supplies_ulsan   REAL DEFAULT 0,
        mfg_supplies_gimhae2 REAL DEFAULT 0,
        -- 변동비: 제조경비 - 지급수수료
        mfg_fee_gimhae      REAL DEFAULT 0,
        mfg_fee_busan       REAL DEFAULT 0,
        mfg_fee_ulsan       REAL DEFAULT 0,
        mfg_fee_gimhae2     REAL DEFAULT 0,
        -- 변동비: 제조경비 - 기타
        mfg_other_gimhae    REAL DEFAULT 0,
        mfg_other_busan     REAL DEFAULT 0,
        mfg_other_ulsan     REAL DEFAULT 0,
        mfg_other_gimhae2   REAL DEFAULT 0,
        -- 변동비: 판매운반비
        selling_trans_gimhae    REAL DEFAULT 0,
        selling_trans_busan     REAL DEFAULT 0,
        selling_trans_ulsan     REAL DEFAULT 0,
        selling_trans_gimhae2   REAL DEFAULT 0,
        -- 변동비: 상품매입
        merch_purchase_gimhae   REAL DEFAULT 0,
        merch_purchase_busan    REAL DEFAULT 0,
        merch_purchase_ulsan    REAL DEFAULT 0,
        merch_purchase_gimhae2  REAL DEFAULT 0,
        -- 고정비: 노무비 - 급료
        labor_salary_gimhae     REAL DEFAULT 0,
        labor_salary_busan      REAL DEFAULT 0,
        labor_salary_ulsan      REAL DEFAULT 0,
        labor_salary_gimhae2    REAL DEFAULT 0,
        -- 고정비: 노무비 - 임금
        labor_wage_gimhae       REAL DEFAULT 0,
        labor_wage_busan        REAL DEFAULT 0,
        labor_wage_ulsan        REAL DEFAULT 0,
        labor_wage_gimhae2      REAL DEFAULT 0,
        -- 고정비: 노무비 - 상여금
        labor_bonus_gimhae      REAL DEFAULT 0,
        labor_bonus_busan       REAL DEFAULT 0,
        labor_bonus_ulsan       REAL DEFAULT 0,
        labor_bonus_gimhae2     REAL DEFAULT 0,
        -- 고정비: 노무비 - 퇴충전입액
        labor_retire_gimhae     REAL DEFAULT 0,
        labor_retire_busan      REAL DEFAULT 0,
        labor_retire_ulsan      REAL DEFAULT 0,
        labor_retire_gimhae2    REAL DEFAULT 0,
        -- 고정비: 노무비 - 외주용역비
        labor_outsrc_gimhae     REAL DEFAULT 0,
        labor_outsrc_busan      REAL DEFAULT 0,
        labor_outsrc_ulsan      REAL DEFAULT 0,
        labor_outsrc_gimhae2    REAL DEFAULT 0,
        -- 고정비: 인건비 - 급료
        staff_salary_gimhae     REAL DEFAULT 0,
        staff_salary_busan      REAL DEFAULT 0,
        staff_salary_ulsan      REAL DEFAULT 0,
        staff_salary_gimhae2    REAL DEFAULT 0,
        -- 고정비: 인건비 - 상여금
        staff_bonus_gimhae      REAL DEFAULT 0,
        staff_bonus_busan       REAL DEFAULT 0,
        staff_bonus_ulsan       REAL DEFAULT 0,
        staff_bonus_gimhae2     REAL DEFAULT 0,
        -- 고정비: 인건비 - 퇴충전입액
        staff_retire_gimhae     REAL DEFAULT 0,
        staff_retire_busan      REAL DEFAULT 0,
        staff_retire_ulsan      REAL DEFAULT 0,
        staff_retire_gimhae2    REAL DEFAULT 0,
        -- 고정비: 제조경비 (대표 항목만)
        fix_depr_gimhae         REAL DEFAULT 0,  -- 감가상각비
        fix_depr_busan          REAL DEFAULT 0,
        fix_depr_ulsan          REAL DEFAULT 0,
        fix_depr_gimhae2        REAL DEFAULT 0,
        fix_lease_gimhae        REAL DEFAULT 0,  -- 지급임차료
        fix_lease_busan         REAL DEFAULT 0,
        fix_lease_ulsan         REAL DEFAULT 0,
        fix_lease_gimhae2       REAL DEFAULT 0,
        fix_outsrc_gimhae       REAL DEFAULT 0,  -- 외주가공비
        fix_outsrc_busan        REAL DEFAULT 0,
        fix_outsrc_ulsan        REAL DEFAULT 0,
        fix_outsrc_gimhae2      REAL DEFAULT 0,
        fix_other_gimhae        REAL DEFAULT 0,  -- 기타경비
        fix_other_busan         REAL DEFAULT 0,
        fix_other_ulsan         REAL DEFAULT 0,
        fix_other_gimhae2       REAL DEFAULT 0,
        -- 영업외
        non_op_income           REAL DEFAULT 0,  -- 영업외수익
        non_op_expense          REAL DEFAULT 0,  -- 영업외비용
        interest_income         REAL DEFAULT 0,  -- 이자수익
        interest_expense        REAL DEFAULT 0,  -- 이자비용
        updated_at              TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(year, month)
    )""")

    # ── 3. 인원 및 근무시간 (매월 입력) ──────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_labor (
        id                  INTEGER PRIMARY KEY AUTOINCREMENT,
        year                INTEGER NOT NULL,
        month               INTEGER NOT NULL,
        -- 인원 (명, 소수점 가능 - 월평균)
        mgmt_rkm            REAL DEFAULT 0,   -- 관리직 RKM
        mgmt_hkmc           REAL DEFAULT 0,   -- 관리직 HKMC
        prod_rkm            REAL DEFAULT 0,   -- 생산직 RKM
        prod_hkmc           REAL DEFAULT 0,   -- 생산직 HKMC
        -- 입퇴사
        hire_count          INTEGER DEFAULT 0,
        resign_count        INTEGER DEFAULT 0,
        -- 근무시간 (시간)
        work_hours_rkm      REAL DEFAULT 0,
        work_hours_hkmc     REAL DEFAULT 0,
        overtime_gimhae     REAL DEFAULT 0,   -- 잔업시간 김해
        overtime_busan      REAL DEFAULT 0,   -- 잔업시간 부산
        base_hours_gimhae   REAL DEFAULT 0,   -- 기본근로시간 김해
        base_hours_busan    REAL DEFAULT 0,   -- 기본근로시간 부산
        -- 상여금 (천원)
        bonus_prod_rkm      REAL DEFAULT 0,   -- 생산직 상여 RKM
        bonus_prod_hkmc     REAL DEFAULT 0,   -- 생산직 상여 HKMC
        -- 퇴직급여 (천원)
        retire_mgmt_rkm     REAL DEFAULT 0,   -- 사무직 퇴직 RKM
        retire_mgmt_hkmc    REAL DEFAULT 0,   -- 사무직 퇴직 HKMC
        retire_prod_rkm     REAL DEFAULT 0,   -- 생산직 퇴직 RKM
        retire_prod_hkmc    REAL DEFAULT 0,   -- 생산직 퇴직 HKMC
        updated_at          TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(year, month)
    )""")

    # ── 4. 업계동향 (매월 수시 입력) ─────────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS industry_news (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        year        INTEGER NOT NULL,
        month       INTEGER NOT NULL,
        company     TEXT NOT NULL,   -- 회사명 (e.g. '르노코리아', 'GM Korea', '현대자동차')
        headline    TEXT NOT NULL,   -- 기사 제목
        content     TEXT,            -- 기사 내용
        source      TEXT,            -- 출처 (e.g. '이지경제 04.02')
        seq         INTEGER DEFAULT 1,
        created_at  TEXT DEFAULT (datetime('now','localtime'))
    )""")

    # ── 5. 시장점유율 (매월 입력) ─────────────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_market_share (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        year        INTEGER NOT NULL,
        month       INTEGER NOT NULL,
        company     TEXT NOT NULL,   -- 현대, 기아, GM, 르노코리아, KG모빌리티
        share_pct   REAL DEFAULT 0,  -- 점유율 (%)
        UNIQUE(year, month, company)
    )""")

    # ── 6. TOP10 판매 모델 (매월 입력) ────────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_top_models (
        id          INTEGER PRIMARY KEY AUTOINCREMENT,
        year        INTEGER NOT NULL,
        month       INTEGER NOT NULL,
        rank        INTEGER NOT NULL,
        model_name  TEXT NOT NULL,
        company     TEXT NOT NULL,
        sales_qty   INTEGER DEFAULT 0,
        UNIQUE(year, month, rank)
    )""")

    # ── 7. 회계팀 자료 (RKM/HKMC 분리, 매월) ────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_acct (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        year            INTEGER NOT NULL,
        month           INTEGER NOT NULL,
        item_code       TEXT NOT NULL,       -- e.g. 'sales_rkm', 'va_hkmc'
        item_name       TEXT NOT NULL,
        value           REAL DEFAULT 0,
        updated_at      TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(year, month, item_code)
    )""")

    # ── 8. 운영실적 (서술형, 매월) ───────────────────────────────────────────
    cur.execute("""
    CREATE TABLE IF NOT EXISTS monthly_operations (
        id              INTEGER PRIMARY KEY AUTOINCREMENT,
        year            INTEGER NOT NULL,
        month           INTEGER NOT NULL,
        section         TEXT NOT NULL,       -- 섹션코드 (e.g. 'summary', 'sales', 'quality')
        section_name    TEXT NOT NULL,
        content         TEXT DEFAULT '',
        updated_at      TEXT DEFAULT (datetime('now','localtime')),
        UNIQUE(year, month, section)
    )""")

    conn.commit()
    conn.close()
    print(f"DB 초기화 완료: {DB_PATH}")


# ── CRUD 헬퍼 ──────────────────────────────────────────────────────────────

def save_monthly_pl(year: int, month: int, data: dict):
    conn = get_conn()
    cols = ", ".join(data.keys())
    placeholders = ", ".join(["?"] * len(data))
    updates = ", ".join([f"{k}=excluded.{k}" for k in data.keys()])
    sql = f"""
        INSERT INTO monthly_pl (year, month, {cols})
        VALUES (?, ?, {placeholders})
        ON CONFLICT(year, month) DO UPDATE SET {updates},
            updated_at=datetime('now','localtime')
    """
    conn.execute(sql, [year, month] + list(data.values()))
    conn.commit()
    conn.close()


def load_monthly_pl(year: int, month: int):
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM monthly_pl WHERE year=? AND month=?", (year, month)
    ).fetchone()
    conn.close()
    return dict(row) if row else {}


def save_monthly_labor(year: int, month: int, data: dict):
    conn = get_conn()
    cols = ", ".join(data.keys())
    placeholders = ", ".join(["?"] * len(data))
    updates = ", ".join([f"{k}=excluded.{k}" for k in data.keys()])
    sql = f"""
        INSERT INTO monthly_labor (year, month, {cols})
        VALUES (?, ?, {placeholders})
        ON CONFLICT(year, month) DO UPDATE SET {updates},
            updated_at=datetime('now','localtime')
    """
    conn.execute(sql, [year, month] + list(data.values()))
    conn.commit()
    conn.close()


def load_monthly_labor(year: int, month: int):
    conn = get_conn()
    row = conn.execute(
        "SELECT * FROM monthly_labor WHERE year=? AND month=?", (year, month)
    ).fetchone()
    conn.close()
    return dict(row) if row else {}


def load_all_months(year: int):
    """해당 연도 전체 월 데이터 조회 (누계 계산용)"""
    conn = get_conn()
    rows_pl = conn.execute(
        "SELECT * FROM monthly_pl WHERE year=? ORDER BY month", (year,)
    ).fetchall()
    rows_lb = conn.execute(
        "SELECT * FROM monthly_labor WHERE year=? ORDER BY month", (year,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows_pl], [dict(r) for r in rows_lb]


def save_industry_news(year: int, month: int, items: list[dict]):
    """업계동향 저장 (기존 데이터 삭제 후 재삽입)"""
    conn = get_conn()
    conn.execute(
        "DELETE FROM industry_news WHERE year=? AND month=?", (year, month)
    )
    for item in items:
        conn.execute("""
            INSERT INTO industry_news (year, month, company, headline, content, source, seq)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (year, month, item.get("company",""), item.get("headline",""),
              item.get("content",""), item.get("source",""), item.get("seq", 1)))
    conn.commit()
    conn.close()


def load_industry_news(year: int, month: int):
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM industry_news WHERE year=? AND month=? ORDER BY company, seq",
        (year, month)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def save_annual_plan(year: int, month: int, items: list[dict]):
    """사업계획 저장 (upsert)
    items: [{"item_code": "sales_rkm", "item_name": "매출액_RKM", "value": 123456}, ...]
    """
    conn = get_conn()
    for item in items:
        conn.execute("""
            INSERT INTO annual_plan (year, month, item_code, item_name, value)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(year, month, item_code) DO UPDATE SET
                item_name=excluded.item_name,
                value=excluded.value,
                updated_at=datetime('now','localtime')
        """, (year, month, item["item_code"], item["item_name"], item.get("value", 0)))
    conn.commit()
    conn.close()


def load_annual_plan(year: int, month: int = None) -> list[dict]:
    """사업계획 조회
    month=None이면 해당 연도 전체, month 지정하면 해당 월만
    """
    conn = get_conn()
    if month is None:
        rows = conn.execute(
            "SELECT * FROM annual_plan WHERE year=? ORDER BY month, item_code",
            (year,)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM annual_plan WHERE year=? AND month=? ORDER BY item_code",
            (year, month)
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def load_annual_plan_as_dict(year: int, month: int) -> dict:
    """사업계획을 {item_code: value} 딕셔너리로 반환"""
    rows = load_annual_plan(year, month)
    return {r["item_code"]: r["value"] for r in rows}


def save_monthly_acct(year: int, month: int, items: list[dict]):
    """회계팀 자료 저장 (upsert)"""
    conn = get_conn()
    for item in items:
        conn.execute("""
            INSERT INTO monthly_acct (year, month, item_code, item_name, value)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(year, month, item_code) DO UPDATE SET
                item_name=excluded.item_name, value=excluded.value,
                updated_at=datetime('now','localtime')
        """, (year, month, item["item_code"], item["item_name"], item.get("value", 0)))
    conn.commit()
    conn.close()


def load_monthly_acct(year: int, month: int) -> dict:
    """회계팀 자료 {item_code: value} 반환"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT item_code, value FROM monthly_acct WHERE year=? AND month=?",
        (year, month)
    ).fetchall()
    conn.close()
    return {r["item_code"]: r["value"] for r in rows}


def save_monthly_operations(year: int, month: int, sections: list[dict]):
    """운영실적 서술 저장"""
    conn = get_conn()
    for s in sections:
        conn.execute("""
            INSERT INTO monthly_operations (year, month, section, section_name, content)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(year, month, section) DO UPDATE SET
                section_name=excluded.section_name, content=excluded.content,
                updated_at=datetime('now','localtime')
        """, (year, month, s["section"], s["section_name"], s.get("content", "")))
    conn.commit()
    conn.close()


def load_monthly_operations(year: int, month: int) -> list[dict]:
    """운영실적 서술 조회"""
    conn = get_conn()
    rows = conn.execute(
        "SELECT * FROM monthly_operations WHERE year=? AND month=? ORDER BY section",
        (year, month)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


if __name__ == "__main__":
    init_db()
