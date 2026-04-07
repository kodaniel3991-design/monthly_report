"""
보고서 다운로드 페이지
모든 데이터 → 기존 템플릿 기반 Excel 파일 자동 생성
"""

import streamlit as st
import pandas as pd
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from database import (
    load_monthly_pl, load_monthly_labor,
    load_industry_news, get_conn
)
from calculator import (
    build_factory_pl_from_db, build_labor_input_from_db,
    calc_value_added, calc_labor_productivity_total,
    calc_labor_productivity_by_division, _sum_factories
)
from excel_generator import generate_excel, fill_labor_productivity_template, fill_industry_template
from flow_bar import render_flow_bar

st.set_page_config(page_title="보고서 다운로드", layout="wide")

render_flow_bar(current_step=6)

# ── 템플릿 경로 ──────────────────────────────────────────────────────────
TEMPLATE_DIR = Path(__file__).parent.parent / "data" / "templates"
LP_TEMPLATE = TEMPLATE_DIR / "노동생산성_템플릿.xlsx"
NEWS_TEMPLATE = TEMPLATE_DIR / "업계동향_템플릿.xls"

# ── 헤더 카드 ────────────────────────────────────────────────────────────
st.markdown(
    "<div style='background:white; border:0.5px solid #dddbd7; border-radius:8px; "
    "padding:18px 24px 14px; box-shadow:0 1px 3px rgba(85,73,64,0.08); margin-bottom:16px;'>"
    "<div style='font-family:Sora,Pretendard,sans-serif; font-size:18px; font-weight:700; "
    "color:#000; letter-spacing:-0.03em;'>보고서 다운로드</div>"
    "<div style='font-size:12px; color:#a8a9aa; margin-top:4px;'>"
    "입력된 데이터를 기존 양식과 동일한 Excel 파일로 생성합니다</div>"
    "</div>",
    unsafe_allow_html=True
)

# ── 연월 선택 ────────────────────────────────────────────────────────────
c1, c2, _ = st.columns([1, 1, 4])
with c1:
    year = st.selectbox("연도", range(2024, 2028), index=2)
with c2:
    month = st.selectbox("월", range(1, 13), index=2)

# ── 데이터 로드 ──────────────────────────────────────────────────────────
pl_data = load_monthly_pl(year, month)
labor_data = load_monthly_labor(year, month)
news_items = load_industry_news(year, month)

conn = get_conn()
top_models = [dict(r) for r in conn.execute(
    "SELECT * FROM monthly_top_models WHERE year=? AND month=? ORDER BY rank",
    (year, month)
).fetchall()]
market_share_rows = conn.execute(
    "SELECT company, share_pct FROM monthly_market_share WHERE year=? AND month=?",
    (year, month)
).fetchall()
conn.close()
market_share = {r["company"]: r["share_pct"] for r in market_share_rows}

# ── 데이터 상태 카드 ─────────────────────────────────────────────────────
def _status_chip(label, ok, detail=""):
    if ok:
        return (
            f"<div style='display:inline-flex; align-items:center; gap:6px; padding:8px 14px; "
            f"border-radius:6px; border:1px solid #879a77; background:#f0f3ee;'>"
            f"<span style='color:#55654a; font-size:13px; font-weight:500;'>&#10003; {label}</span>"
            f"<span style='color:#879a77; font-size:11px;'>{detail}</span>"
            f"</div>"
        )
    return (
        f"<div style='display:inline-flex; align-items:center; gap:6px; padding:8px 14px; "
        f"border-radius:6px; border:1px solid #dddbd7; background:#f5f4f2;'>"
        f"<span style='color:#a8a9aa; font-size:13px;'>{label}</span>"
        f"<span style='color:#c5c6c7; font-size:11px;'>미입력</span>"
        f"</div>"
    )

st.markdown(
    "<div style='display:flex; gap:8px; flex-wrap:wrap; margin:12px 0 8px;'>"
    + _status_chip("손익실적", bool(pl_data), "입력완료")
    + _status_chip("인원·노무비", bool(labor_data), "입력완료")
    + _status_chip("업계동향", bool(news_items), f"{len(news_items)}건")
    + "</div>",
    unsafe_allow_html=True
)

if not pl_data:
    st.error("손익실적 데이터가 없습니다. 손익실적 입력 후 다시 시도하세요.")
    st.stop()

# ── 계산 ─────────────────────────────────────────────────────────────────
gimhae = build_factory_pl_from_db(pl_data, "gimhae")
busan = build_factory_pl_from_db(pl_data, "busan")
ulsan = build_factory_pl_from_db(pl_data, "ulsan")
gimhae2 = build_factory_pl_from_db(pl_data, "gimhae2")

rkm = _sum_factories("RKM", gimhae, busan)
hkmc = _sum_factories("HKMC", ulsan, gimhae2)
total = _sum_factories("계", gimhae, busan, ulsan, gimhae2)

labor = build_labor_input_from_db(labor_data) if labor_data else None

def labor_cost_from_pl(pl_data, factory):
    return sum([
        pl_data.get(f"labor_salary_{factory}", 0) or 0,
        pl_data.get(f"labor_wage_{factory}", 0) or 0,
        pl_data.get(f"labor_bonus_{factory}", 0) or 0,
        pl_data.get(f"labor_retire_{factory}", 0) or 0,
        pl_data.get(f"labor_outsrc_{factory}", 0) or 0,
        pl_data.get(f"staff_salary_{factory}", 0) or 0,
        pl_data.get(f"staff_bonus_{factory}", 0) or 0,
        pl_data.get(f"staff_retire_{factory}", 0) or 0,
    ])

lc_total = sum(labor_cost_from_pl(pl_data, f)
               for f in ["gimhae", "busan", "ulsan", "gimhae2"])

if labor:
    retire_total = sum([
        labor_data.get("retire_mgmt_rkm", 0) or 0,
        labor_data.get("retire_mgmt_hkmc", 0) or 0,
        labor_data.get("retire_prod_rkm", 0) or 0,
        labor_data.get("retire_prod_hkmc", 0) or 0,
    ])
    lp_total = calc_labor_productivity_total(total, labor, lc_total, retire_total)
    lp_rkm, lp_hkmc = calc_labor_productivity_by_division(rkm, hkmc, labor, lc_total)
else:
    lp_total = lp_rkm = lp_hkmc = None

# ── 미리보기 ─────────────────────────────────────────────────────────────
st.divider()

tab_pl, tab_lp = st.tabs(["손익실적 요약", "노동생산성 요약"])

with tab_pl:
    data_preview = {
        "구분": ["김해", "부산", "RKM", "울산", "김해2", "HKMC", "합계"],
        "매출액": [gimhae.sales, busan.sales, rkm.sales,
                  ulsan.sales, gimhae2.sales, hkmc.sales, total.sales],
        "한계이익": [gimhae.contribution_margin, busan.contribution_margin, rkm.contribution_margin,
                    ulsan.contribution_margin, gimhae2.contribution_margin, hkmc.contribution_margin,
                    total.contribution_margin],
        "부가가치": [calc_value_added(f) for f in [gimhae, busan, rkm, ulsan, gimhae2, hkmc, total]],
        "영업이익": [gimhae.operating_profit, busan.operating_profit, rkm.operating_profit,
                    ulsan.operating_profit, gimhae2.operating_profit, hkmc.operating_profit,
                    total.operating_profit],
    }
    df = pd.DataFrame(data_preview).set_index("구분")
    st.dataframe(df.style.format("{:,.0f}"), use_container_width=True)

with tab_lp:
    if lp_total:
        lp_data = {
            "지표": ["부가가치율", "노동생산성(천원/인)", "근로소득배분율",
                     "인건비율", "1인당임금(천원)", "시간당임금(천원)"],
            "RKM": [f"{lp_rkm.value_added_ratio*100:.2f}%", f"{lp_rkm.labor_productivity:,.0f}",
                    f"{lp_rkm.labor_income_ratio*100:.2f}%", f"{lp_rkm.labor_cost_to_sales*100:.2f}%",
                    f"{lp_rkm.wage_per_person:,.0f}", f"{lp_rkm.hourly_wage:,.1f}"],
            "HKMC": [f"{lp_hkmc.value_added_ratio*100:.2f}%", f"{lp_hkmc.labor_productivity:,.0f}",
                     f"{lp_hkmc.labor_income_ratio*100:.2f}%", f"{lp_hkmc.labor_cost_to_sales*100:.2f}%",
                     f"{lp_hkmc.wage_per_person:,.0f}", f"{lp_hkmc.hourly_wage:,.1f}"],
            "합계": [f"{lp_total.value_added_ratio*100:.2f}%", f"{lp_total.labor_productivity:,.0f}",
                     f"{lp_total.labor_income_ratio*100:.2f}%", f"{lp_total.labor_cost_to_sales*100:.2f}%",
                     f"{lp_total.wage_per_person:,.0f}", f"{lp_total.hourly_wage:,.1f}"],
        }
        st.dataframe(pd.DataFrame(lp_data).set_index("지표"), use_container_width=True)
    else:
        st.info("인원·노무비 데이터 입력 후 노동생산성 지표가 표시됩니다.")

# ── 다운로드 섹션 ────────────────────────────────────────────────────────
st.divider()

st.markdown(
    "<div style='font-family:Sora,Pretendard,sans-serif; font-size:15px; font-weight:700; "
    "color:#554940; letter-spacing:-0.02em; margin-bottom:12px;'>Excel 파일 생성</div>",
    unsafe_allow_html=True
)

# 출력 옵션 — 카드형 체크박스
col_a, col_b = st.columns(2)
with col_a:
    include_lp = st.checkbox("노동생산성 시트 포함", value=bool(labor_data))
with col_b:
    include_news = st.checkbox("업계동향 시트 포함", value=bool(news_items))

# 템플릿 상태 표시
lp_template = None
news_template = None

if include_lp and labor_data:
    if LP_TEMPLATE.exists():
        lp_template = str(LP_TEMPLATE)
    else:
        lp_template = st.file_uploader("노동생산성 템플릿 (.xlsx)", type=["xlsx"],
                                        key="lp_template", label_visibility="collapsed")

if include_news and news_items:
    if NEWS_TEMPLATE.exists():
        news_template = str(NEWS_TEMPLATE)
    else:
        news_template = st.file_uploader("업계동향 템플릿 (.xls)", type=["xls"],
                                          key="news_template", label_visibility="collapsed")

# 템플릿 상태 칩
tmpl_chips = []
if include_lp:
    tmpl_chips.append(("노동생산성", bool(lp_template)))
if include_news:
    tmpl_chips.append(("업계동향", bool(news_template)))

if tmpl_chips:
    def _tmpl_chip(label, ok):
        border = "#879a77" if ok else "#dddbd7"
        bg = "#f0f3ee" if ok else "#f5f4f2"
        color = "#55654a" if ok else "#a8a9aa"
        icon = "&#10003;" if ok else "&#10007;"
        return (
            f"<span style='display:inline-flex; align-items:center; gap:4px; padding:4px 10px; "
            f"border-radius:4px; border:1px solid {border}; background:{bg}; "
            f"font-size:11px; color:{color};'>{icon} {label} 템플릿</span>"
        )
    chips_html = "".join(_tmpl_chip(label, ok) for label, ok in tmpl_chips)
    st.markdown(
        f"<div style='display:flex; gap:6px; margin:8px 0;'>{chips_html}</div>",
        unsafe_allow_html=True
    )

# ── 다운로드 버튼 ────────────────────────────────────────────────────────
if st.button("Excel 생성 및 다운로드", type="primary", use_container_width=True):
    with st.spinner("Excel 파일 생성 중..."):
        try:
            generated = []

            # (1) 월차보고 종합
            excel_bytes = generate_excel(
                year=year, month=month,
                pl_data=pl_data,
                labor_data=labor_data or {},
                lp_total=lp_total, lp_rkm=lp_rkm, lp_hkmc=lp_hkmc,
                labor_input=labor,
                news_items=news_items if include_news else [],
                top_models=[{"rank": r["rank"], "model": r["model_name"],
                             "company": r["company"], "qty": r["sales_qty"]}
                            for r in top_models] if include_news else [],
                market_share=market_share if include_news else {},
            )
            generated.append(("xlsx", f"월차보고_{year}년_{month:02d}월.xlsx", excel_bytes))

            # (2) 노동생산성 (템플릿 기반)
            if include_lp and labor_data and lp_template:
                lp_bytes = fill_labor_productivity_template(
                    template_path=lp_template,
                    pl_data=pl_data, labor_data=labor_data,
                    labor_input=labor, year=year, month=month,
                )
                generated.append(("xlsx", f"노동생산성_{year}년_{month:02d}월.xlsx", lp_bytes))

            # (3) 업계동향 (템플릿 기반)
            if include_news and news_items and news_template:
                news_bytes = fill_industry_template(
                    template_path=news_template,
                    year=year, month=month,
                    news_items=news_items,
                    top_models=[{"rank": r["rank"], "model_name": r["model_name"],
                                 "company": r["company"], "sales_qty": r["sales_qty"]}
                                for r in top_models] if top_models else [],
                    market_share=market_share,
                )
                generated.append(("xls", f"업계동향_{year}년_{month:02d}월.xls", news_bytes))

            # 다운로드 버튼 렌더링
            st.success(f"{len(generated)}개 파일 생성 완료")

            for i, (ext, filename, data) in enumerate(generated):
                mime = ("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        if ext == "xlsx" else "application/vnd.ms-excel")
                st.download_button(
                    label=f"📥 {filename}",
                    data=data,
                    file_name=filename,
                    mime=mime,
                    key=f"dl_{i}",
                    use_container_width=True,
                    type="primary",
                )

        except Exception as e:
            st.error(f"생성 실패: {e}")
            import traceback
            st.code(traceback.format_exc())
