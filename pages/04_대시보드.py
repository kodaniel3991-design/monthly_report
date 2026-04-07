"""
대시보드 - 노동생산성 계산 결과 조회
손익실적 + 인원·노무비 → 자동 계산된 6대 지표 표시
"""

import streamlit as st
import sys
from pathlib import Path
import plotly.graph_objects as go

sys.path.insert(0, str(Path(__file__).parent))
from database import load_monthly_pl, load_monthly_labor
from calculator import (
    FactoryPL, LaborInput,
    build_factory_pl_from_db, build_labor_input_from_db,
    calc_value_added, calc_labor_productivity_total,
    calc_labor_productivity_by_division, _sum_factories
)
from flow_bar import render_flow_bar

st.set_page_config(page_title="대시보드", layout="wide")

render_flow_bar(current_step=4)

# ── 헤더 카드 ────────────────────────────────────────────────────────────
st.markdown(
    "<div style='background:white; border:0.5px solid #dddbd7; border-radius:8px; "
    "padding:18px 24px 14px; box-shadow:0 1px 3px rgba(85,73,64,0.08); margin-bottom:16px;'>"
    "<div style='font-family:Sora,Pretendard,sans-serif; font-size:18px; font-weight:700; "
    "color:#000; letter-spacing:-0.03em;'>노동생산성 대시보드</div>"
    "<div style='font-size:12px; color:#a8a9aa; margin-top:4px;'>"
    "손익실적 + 인원·노무비 데이터 기반 자동 계산</div>"
    "</div>",
    unsafe_allow_html=True
)

# 메트릭 카드 스타일 (기본 st.metric 숨기고 커스텀 사용)
st.markdown("""
<style>
[data-testid="stMetric"] { position: relative; }
[data-testid="stMetricValue"] { font-size: 20px !important; }
[data-testid="stMetricDelta"] { position: absolute; top: 12px; right: 16px; }
</style>
""", unsafe_allow_html=True)

def _card(col, label, value, unit="천원", delta=None, delta_color="normal"):
    """커스텀 메트릭: 숫자 옆에 작은 단위 표기"""
    with col:
        delta_html = ""
        if delta:
            d_color = "#55654a" if delta_color == "normal" else "#dc2626"
            delta_html = (
                f"<span style='position:absolute; top:12px; right:16px; "
                f"font-size:11px; color:{d_color}; font-weight:500;'>{delta}</span>"
            )
        st.markdown(
            f"<div style='background:white; border:0.5px solid #dddbd7; border-radius:10px; "
            f"padding:16px 20px; min-height:88px; box-shadow:0 1px 3px rgba(85,73,64,0.08); "
            f"position:relative; transition: box-shadow 250ms, transform 250ms;'>"
            f"{delta_html}"
            f"<div style='font-size:13px; color:#73787c; margin-bottom:4px;'>{label}</div>"
            f"<div style='font-family:Sora,Pretendard,sans-serif; font-weight:700; "
            f"letter-spacing:-0.03em;'>"
            f"<span style='font-size:20px; color:#000;'>{value}</span>"
            f"<span style='font-size:11px; font-weight:400; color:#a8a9aa; "
            f"margin-left:4px;'>{unit}</span>"
            f"</div></div>",
            unsafe_allow_html=True
        )

c1, c2, _ = st.columns([1, 1, 4])
with c1:
    year = st.selectbox("연도", range(2024, 2028), index=2)
with c2:
    month = st.selectbox("월", range(1, 13), index=2)

pl_data = load_monthly_pl(year, month)
labor_data = load_monthly_labor(year, month)

if not pl_data:
    st.warning(f"{year}년 {month}월 손익실적 데이터가 없습니다. 먼저 손익실적을 입력하세요.")
    st.stop()

if not labor_data:
    st.warning(f"{year}년 {month}월 인원·노무비 데이터가 없습니다. 먼저 인원·노무비를 입력하세요.")
    st.stop()

# ── 데이터 계산 ──────────────────────────────────────────────────────────
gimhae = build_factory_pl_from_db(pl_data, "gimhae")
busan = build_factory_pl_from_db(pl_data, "busan")
ulsan = build_factory_pl_from_db(pl_data, "ulsan")
gimhae2 = build_factory_pl_from_db(pl_data, "gimhae2")

rkm = _sum_factories("RKM", gimhae, busan)
hkmc = _sum_factories("HKMC", ulsan, gimhae2)
total = _sum_factories("계", gimhae, busan, ulsan, gimhae2)

labor = build_labor_input_from_db(labor_data)

def sum_labor_cost(pl_data, factory):
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

labor_cost_total = sum(sum_labor_cost(pl_data, f) for f in ["gimhae", "busan", "ulsan", "gimhae2"])
retire_total = labor.retire_total

lp_total = calc_labor_productivity_total(total, labor, labor_cost_total, retire_total)
lp_rkm, lp_hkmc = calc_labor_productivity_by_division(rkm, hkmc, labor, labor_cost_total)

# ── 섹션 헤더 헬퍼 ───────────────────────────────────────────────────────
def section_header(title):
    st.markdown(
        f"<div style='font-family:Sora,Pretendard,sans-serif; font-size:15px; font-weight:700; "
        f"color:#554940; letter-spacing:-0.02em; margin:20px 0 12px; "
        f"padding-bottom:8px; border-bottom:1px solid #eceae7;'>{title}</div>",
        unsafe_allow_html=True
    )

# ── 섹션1: 손익 요약 ────────────────────────────────────────────────────
st.divider()
section_header(f"{year}년 {month}월 손익 요약")

m1, m2, m3, m4 = st.columns(4)
_card(m1, "매출액", f"{total.sales:,.0f}")
_card(m2, "한계이익", f"{total.contribution_margin:,.0f}", delta=f"{total.pct(total.contribution_margin):.1f}%")
_card(m3, "영업이익", f"{total.operating_profit:,.0f}",
      delta=f"{total.pct(total.operating_profit):.1f}%",
      delta_color="normal" if total.operating_profit >= 0 else "inverse")
_card(m4, "부가가치", f"{calc_value_added(total):,.0f}", delta=f"{lp_total.value_added_ratio*100:.2f}%")

# ── 섹션2: RKM vs HKMC ──────────────────────────────────────────────────
st.divider()
section_header("사업부별 비교")

col_rkm, col_hkmc = st.columns(2)

def show_division(col, label, pl, lp):
    with col:
        st.markdown(
            f"<div style='font-size:13px; font-weight:600; color:#554940; "
            f"margin-bottom:8px;'>{label}</div>",
            unsafe_allow_html=True
        )
        va = calc_value_added(pl)
        r1, r2 = st.columns(2)
        _card(r1, "매출액", f"{pl.sales:,.0f}")
        _card(r2, "부가가치", f"{va:,.0f}",
              delta=f"{va/pl.sales*100:.1f}%" if pl.sales else "-")
        _card(r1, "변동비", f"{pl.variable_cost:,.0f}",
              delta=f"{pl.pct(pl.variable_cost):.1f}%")
        _card(r2, "노동생산성", f"{lp.labor_productivity:,.0f}", unit="천원/인")
        _card(r1, "한계이익", f"{pl.contribution_margin:,.0f}",
              delta=f"{pl.pct(pl.contribution_margin):.1f}%")
        _card(r2, "근로소득배분율", f"{lp.labor_income_ratio*100:.1f}", unit="%")

show_division(col_rkm, "RKM (르노코리아 납품)", rkm, lp_rkm)
show_division(col_hkmc, "HKMC (현대기아 납품)", hkmc, lp_hkmc)

# ── 섹션3: 6대 노동생산성 지표 ──────────────────────────────────────────
st.divider()
section_header("노동생산성 6대 지표 (총괄)")

def pct_str(v):
    return f"{v*100:.2f}%"

indicators = [
    ("부가가치율",       pct_str(lp_total.value_added_ratio),     "%",      "附加價値 / 賣出額"),
    ("노동생산성",       f"{lp_total.labor_productivity:,.0f}",   "천원/인", "附加價値 / 종업원수"),
    ("근로소득배분율",   pct_str(lp_total.labor_income_ratio),    "%",      "勞務費 / 附加價値"),
    ("인건비율(매출)",   pct_str(lp_total.labor_cost_to_sales),   "%",      "人件費 / 賣出額"),
    ("1인당 임금수준",   f"{lp_total.wage_per_person:,.0f}",      "천원/인", "(勞務費+退職金) / 종업원"),
    ("시간당 임금",      f"{lp_total.hourly_wage:,.0f}",          "천원/h",  "勞務費 / 실작업시간"),
]

cols = st.columns(3)
for i, (name, val, unit, formula) in enumerate(indicators):
    _card(cols[i % 3], name, val, unit=unit)
    with cols[i % 3]:
        st.caption(formula)

# ── 섹션4: 공장별 매출 구조 차트 ─────────────────────────────────────────
st.divider()
section_header("공장별 매출 구조")

factories_data = {"김해": gimhae, "부산": busan, "울산": ulsan, "김해2": gimhae2}

fig = go.Figure()
names = list(factories_data.keys())
var_cost_vals = [f.variable_cost for f in factories_data.values()]
fix_cost_vals = [f.fixed_cost for f in factories_data.values()]
va_vals = [calc_value_added(f) for f in factories_data.values()]

fig.add_trace(go.Bar(name="변동비", x=names, y=var_cost_vals, marker_color="#d4c0b3"))
fig.add_trace(go.Bar(name="고정비", x=names, y=fix_cost_vals, marker_color="#bbc9b3"))
fig.add_trace(go.Bar(name="부가가치", x=names, y=va_vals, marker_color="#879a77"))

fig.update_layout(
    barmode="stack",
    xaxis_title="",
    yaxis_title="금액 (천원)",
    legend=dict(orientation="h", yanchor="bottom", y=1.02),
    height=320,
    margin=dict(t=40, b=20, l=60, r=20),
    plot_bgcolor="rgba(0,0,0,0)",
    paper_bgcolor="rgba(0,0,0,0)",
    font=dict(family="Pretendard, sans-serif", size=12),
    yaxis=dict(gridcolor="#eceae7", zerolinecolor="#dddbd7"),
)
st.plotly_chart(fig, use_container_width=True)

# ── 섹션5: 인원 현황 ────────────────────────────────────────────────────
st.divider()
section_header("인원 현황")

c1, c2, c3, c4 = st.columns(4)
_card(c1, "전체 종업원", f"{labor.total_employees:.0f}", unit="명")
_card(c2, "생산직", f"{labor.prod_employees:.0f}", unit="명")
_card(c3, "실작업시간", f"{labor.total_work_hours:,.1f}", unit="h")
_card(c4, "RKM / HKMC", f"{labor.rkm_ratio*100:.0f} / {labor.hkmc_ratio*100:.0f}", unit="%")
