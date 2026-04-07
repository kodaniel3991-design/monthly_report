"""
전체 업무 흐름 — 세로형 사이드바 내비게이션
Streamlit 멀티페이지 앱용 (pages/ 디렉토리)
"""

import streamlit as st
from design_system import inject_design_system

WORKFLOW = [
    {"step": 0, "label": "사업계획",    "page": "pages/06_사업계획_입력.py",    "phase": "연초 1회",   "group": "데이터 입력"},
    {"step": 1, "label": "업계동향",    "page": "pages/03_업계동향_입력.py",    "phase": "월초",       "group": "데이터 입력"},
    {"step": 2, "label": "손익실적",    "page": "pages/01_손익실적_입력.py",    "phase": "월 마감",    "group": "데이터 입력"},
    {"step": 3, "label": "인원·노무비", "page": "pages/02_인원_노무비_입력.py", "phase": "동시 진행",  "group": "데이터 입력"},
    {"step": 4, "label": "노동생산성",  "page": "pages/04_대시보드.py",         "phase": "자동 계산",  "group": "결과 조회"},
    {"step": 5, "label": "운영실적",    "page": "pages/08_운영실적_입력.py",    "phase": "서술 작성",  "group": "결과 조회"},
    {"step": 6, "label": "보고서",      "page": "pages/05_보고서_다운로드.py",  "phase": "최종 완성",  "group": "결과 조회"},
]

_SIDEBAR_CSS = """
<style>
[data-testid="stSidebarNav"] { display: none !important; }
</style>
"""


def render_flow_bar(current_step: int):
    """사이드바에 세로형 업무 흐름 내비게이션을 렌더링합니다."""

    inject_design_system()

    with st.sidebar:
        # Streamlit 기본 멀티페이지 내비를 숨기고 커스텀으로 대체
        st.markdown(_SIDEBAR_CSS, unsafe_allow_html=True)

        st.markdown(
            "<div style='font-size:11px; font-weight:600; letter-spacing:0.08em; "
            "text-transform:uppercase; color:#a8a9aa; margin-bottom:8px;'>"
            "월차보고 업무 흐름</div>",
            unsafe_allow_html=True
        )

        current_group = None

        for w in WORKFLOW:
            # 그룹 헤더
            if w["group"] != current_group:
                current_group = w["group"]
                st.markdown(
                    f"<div style='font-size:12px; font-weight:700; color:#554940; "
                    f"padding:10px 0 4px; margin-top:2px;'>"
                    f"{current_group}</div>",
                    unsafe_allow_html=True
                )

            is_current = (w["step"] == current_step)

            if is_current:
                # 현재 단계: 강조
                st.markdown(
                    f"<div style='background:#eceae7; border-radius:6px; "
                    f"padding:10px 14px; margin:2px 0;'>"
                    f"<div style='font-size:13.5px; font-weight:600; color:#554940;'>"
                    f"{w['label']}</div>"
                    f"<div style='font-size:10px; color:#8a6e62; margin-top:2px;'>"
                    f"{w['phase']}</div>"
                    f"</div>",
                    unsafe_allow_html=True
                )
            else:
                # 다른 단계: st.page_link (멀티페이지에서 작동)
                st.page_link(
                    w["page"],
                    label=f"{w['label']}  ·  {w['phase']}",
                )

        st.markdown("<div style='height:1px; background:#dddbd7; margin:16px 0;'></div>",
                    unsafe_allow_html=True)
