"""
월차보고 시스템 - 메인 앱
진양오토모티브 김해 / 월별 경영실적 자동 집계
"""

import streamlit as st
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from database import init_db
from design_system import inject_design_system
from flow_bar import render_flow_bar

st.set_page_config(
    page_title="월차보고 시스템",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

init_db()
inject_design_system()

# ── 사이드바 ─────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown(
        "<div style='display:flex; align-items:center; gap:10px; margin-bottom:8px;'>"
        "<div style='width:32px; height:32px; background:#554940; border-radius:6px; "
        "display:flex; align-items:center; justify-content:center; "
        "font-size:14px; font-weight:800; color:white; flex-shrink:0;'>JY</div>"
        "<div>"
        "<div style='font-size:14px; font-weight:700; color:#000; line-height:1.2;'>월차보고 시스템</div>"
        "<div style='font-size:10px; color:#a8a9aa;'>진양오토모티브 김해</div>"
        "</div>"
        "</div>",
        unsafe_allow_html=True
    )
    st.markdown("<div style='height:1px; background:#dddbd7; margin:12px 0;'></div>",
                unsafe_allow_html=True)

render_flow_bar(current_step=-1)

# ── 메인: 사이드바에서 페이지 선택 안내 ──────────────────────────────────
st.markdown(
    "<div style='display:flex; align-items:center; justify-content:center; "
    "height:300px; color:#a8a9aa; font-size:14px;'>"
    "← 사이드바에서 업무 단계를 선택하세요"
    "</div>",
    unsafe_allow_html=True
)
