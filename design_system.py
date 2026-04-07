"""
LUON AI Design System — Streamlit 적용 모듈
Taupe Primary (#554940) + Soft Green Accent (#879A77)
모든 페이지에서 inject_design_system()을 호출하여 적용
"""

import streamlit as st

DESIGN_CSS = """
<style>
/* ══════════════════════════════════════════════
   LUON AI Design System — CSS Variables
   Taupe Primary + Soft Green Accent (2025)
   ══════════════════════════════════════════════ */
@import url('https://fonts.googleapis.com/css2?family=DM+Mono:wght@400;500&family=Sora:wght@300;400;500;600;700;800&display=swap');

:root {
  /* PRIMARY — Taupe */
  --taupe-50:  #f5f0ec;
  --taupe-100: #e8ddd6;
  --taupe-200: #d4c0b3;
  --taupe-300: #bfa091;
  --taupe-400: #8a6e62;
  --taupe-500: #554940;
  --taupe-600: #433a33;
  --taupe-700: #312b27;
  --taupe-800: #1f1b18;

  /* ACCENT — Soft Green */
  --green-50:  #f0f3ee;
  --green-100: #d8e2d2;
  --green-200: #bbc9b3;
  --green-300: #9daf94;
  --green-400: #879a77;
  --green-500: #6e8060;
  --green-600: #55654a;
  --green-700: #3d4a35;

  /* NEUTRAL */
  --gray-50:  #f5f4f2;
  --gray-100: #eceae7;
  --gray-200: #dddbd7;
  --gray-300: #c5c6c7;
  --gray-400: #a8a9aa;
  --gray-500: #73787c;
  --gray-600: #5a5f62;

  /* SEMANTIC */
  --color-success: #879a77;
  --color-warning: #c9ad93;
  --color-error:   #dc2626;
  --color-info:    #d7e5f0;

  /* BACKGROUND */
  --bg-page:    #f5f4f2;
  --bg-surface: #ffffff;
  --bg-subtle:  #eceae7;

  /* TEXT */
  --text-primary:   #000000;
  --text-secondary: #73787c;
  --text-tertiary:  #a8a9aa;

  /* SHADOW */
  --shadow-sm: 0 1px 3px rgba(85,73,64,0.08), 0 1px 2px rgba(0,0,0,0.04);
  --shadow-md: 0 4px 12px rgba(85,73,64,0.12), 0 2px 4px rgba(0,0,0,0.05);
}

/* ── 전역 폰트 ── */
html, body, [class*="css"] {
    font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, sans-serif !important;
}
h1, h2, h3, h4, h5, h6,
[data-testid="stMarkdownContainer"] h1,
[data-testid="stMarkdownContainer"] h2,
[data-testid="stMarkdownContainer"] h3 {
    font-family: 'Sora', 'Pretendard', sans-serif !important;
    letter-spacing: -0.03em;
}
code, pre, [class*="mono"] {
    font-family: 'DM Mono', monospace !important;
}

/* ── 메인 컨테이너 ── */
.stApp {
    background-color: var(--bg-page) !important;
}
[data-testid="stAppViewBlockContainer"] {
    max-width: 1200px;
}

/* ── 입력 필드 컴팩트 ── */
[data-testid="stVerticalBlock"] > div {
    gap: 0.35rem !important;
}
[data-testid="stTextInput"],
[data-testid="stNumberInput"],
[data-testid="stSelectbox"] {
    margin-bottom: -8px !important;
}
.stTextInput > div > div > input,
.stNumberInput > div > div > input {
    height: 36px !important;
    padding: 6px 12px !important;
    font-size: 13px !important;
}
.stTextArea > div > div > textarea {
    font-size: 13px !important;
    padding: 8px 12px !important;
}

/* ── 사이드바 ── */
[data-testid="stSidebar"] {
    background-color: var(--bg-surface) !important;
    border-right: 0.5px solid var(--gray-200) !important;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {
    color: var(--text-secondary);
    font-size: 13.5px;
}

/* ── 버튼 (Primary = Taupe) ── */
.stButton > button[kind="primary"],
.stButton > button[data-testid="baseButton-primary"] {
    background-color: var(--taupe-500) !important;
    border-color: var(--taupe-500) !important;
    color: white !important;
    border-radius: 6px !important;
    font-weight: 600 !important;
    box-shadow: 0 2px 6px rgba(85,73,64,0.30) !important;
    transition: all 150ms cubic-bezier(0.16, 1, 0.3, 1) !important;
}
.stButton > button[kind="primary"]:hover,
.stButton > button[data-testid="baseButton-primary"]:hover {
    background-color: var(--taupe-600) !important;
    border-color: var(--taupe-600) !important;
    box-shadow: 0 4px 12px rgba(85,73,64,0.45) !important;
    transform: translateY(-1px);
}

/* ── 버튼 (Secondary/Ghost) ── */
.stButton > button[kind="secondary"],
.stButton > button[data-testid="baseButton-secondary"] {
    background-color: var(--bg-surface) !important;
    color: var(--text-secondary) !important;
    border: 1.5px solid var(--gray-300) !important;
    border-radius: 6px !important;
    font-weight: 500 !important;
}
.stButton > button[kind="secondary"]:hover,
.stButton > button[data-testid="baseButton-secondary"]:hover {
    background-color: var(--bg-subtle) !important;
    color: var(--text-primary) !important;
}

/* ── 입력 필드 ── */
.stTextInput > div > div > input,
.stNumberInput > div > div > input,
.stSelectbox > div > div,
.stTextArea > div > div > textarea {
    border: 1.5px solid var(--gray-300) !important;
    border-radius: 6px !important;
    font-size: 14px !important;
    transition: border-color 150ms, box-shadow 150ms !important;
}
.stTextInput > div > div > input:focus,
.stNumberInput > div > div > input:focus,
.stTextArea > div > div > textarea:focus {
    border-color: var(--taupe-500) !important;
    box-shadow: 0 0 0 3px rgba(85,73,64,0.12) !important;
}

/* ── 탭 ── */
.stTabs [data-baseweb="tab-list"] {
    gap: 2px;
    border-bottom: 1px solid var(--gray-200);
}
.stTabs [data-baseweb="tab"] {
    border-radius: 6px 6px 0 0 !important;
    font-weight: 500 !important;
    font-size: 13.5px !important;
    padding: 8px 16px !important;
}
.stTabs [aria-selected="true"] {
    background-color: var(--bg-surface) !important;
    border-bottom: 2px solid var(--taupe-500) !important;
    color: var(--taupe-600) !important;
    font-weight: 600 !important;
}

/* ── 메트릭 카드 ── */
[data-testid="stMetric"] {
    background: var(--bg-surface);
    border: 0.5px solid var(--gray-200);
    border-radius: 10px;
    padding: 16px 20px;
    box-shadow: var(--shadow-sm);
    transition: box-shadow 250ms, transform 250ms;
}
[data-testid="stMetric"]:hover {
    box-shadow: var(--shadow-md);
    transform: translateY(-2px);
}
[data-testid="stMetricLabel"] {
    font-size: 13px !important;
    color: var(--text-secondary) !important;
}
[data-testid="stMetricValue"] {
    font-family: 'Sora', 'Pretendard', sans-serif !important;
    font-weight: 700 !important;
    letter-spacing: -0.03em !important;
    color: var(--text-primary) !important;
}
[data-testid="stMetricDelta"] > div {
    font-size: 12px !important;
}

/* ── 데이터프레임 / 테이블 ── */
[data-testid="stDataFrame"] {
    border: 0.5px solid var(--gray-200) !important;
    border-radius: 10px !important;
    overflow: hidden !important;
    box-shadow: var(--shadow-sm) !important;
}

/* ── 구분선 ── */
hr {
    border-color: var(--gray-200) !important;
    margin: 16px 0 !important;
}

/* ── 알림 (Success = Green Accent) ── */
.stSuccess {
    background-color: var(--green-50) !important;
    border-left: 3px solid var(--green-400) !important;
    color: var(--green-700) !important;
}
.stInfo {
    background-color: #e7f5ff !important;
    border-left: 3px solid #8aa8c0 !important;
}
.stWarning {
    border-left: 3px solid var(--color-warning) !important;
}
.stError {
    border-left: 3px solid var(--color-error) !important;
}

/* ── 프로그레스 바 ── */
.stProgress > div > div > div {
    background-color: var(--green-400) !important;
}

/* ── Expander ── */
[data-testid="stExpander"] {
    border: 0.5px solid var(--gray-200) !important;
    border-radius: 10px !important;
    background: var(--bg-surface) !important;
}

/* ── 파일 업로더 ── */
[data-testid="stFileUploader"] {
    border: 1.5px dashed var(--gray-300) !important;
    border-radius: 10px !important;
    padding: 20px !important;
}
[data-testid="stFileUploader"]:hover {
    border-color: var(--taupe-400) !important;
    background: var(--taupe-50) !important;
}

/* ── 라디오 / 체크박스 액센트 ── */
input[type="radio"]:checked,
input[type="checkbox"]:checked {
    accent-color: var(--taupe-500) !important;
}

/* ── 셀렉트박스 ── */
[data-baseweb="select"] {
    border-radius: 6px !important;
}

/* ── 스크롤바 ── */
::-webkit-scrollbar { width: 8px; }
::-webkit-scrollbar-track { background: var(--bg-page); }
::-webkit-scrollbar-thumb { background: var(--gray-300); border-radius: 4px; }
::-webkit-scrollbar-thumb:hover { background: var(--gray-400); }

/* ── 페이지 타이틀 ── */
h1 {
    color: var(--text-primary) !important;
    font-weight: 700 !important;
    letter-spacing: -0.04em !important;
}

/* ── 캡션 ── */
.stCaption, [data-testid="stCaptionContainer"] {
    color: var(--text-tertiary) !important;
    font-size: 13px !important;
}
</style>
"""


def inject_design_system():
    """모든 페이지에서 호출하여 디자인 시스템 CSS를 주입합니다."""
    st.markdown(DESIGN_CSS, unsafe_allow_html=True)
