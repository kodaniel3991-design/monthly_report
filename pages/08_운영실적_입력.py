"""
운영실적 입력 (Step 6)
주요업무 추진실적 서술형 보고서 작성
손익 수치 + 업계동향 맥락을 토대로 작성
"""

import streamlit as st
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from database import save_monthly_operations, load_monthly_operations
from flow_bar import render_flow_bar

st.set_page_config(page_title="운영실적 입력", layout="wide")

render_flow_bar(current_step=5)

# ── 섹션 정의 ─────────────────────────────────────────────────────────────
SECTIONS = [
    ("summary",   "종합 요약",           "당월 경영 실적 종합 요약 (매출, 영업이익, 주요 이슈)"),
    ("sales",     "영업·매출 실적",       "사업부별 매출 실적, 수주 현황, 신규 거래처"),
    ("production","생산·품질 실적",       "공장별 생산 실적, 품질 지표, 불량률, 개선 활동"),
    ("cost",      "원가·경비 관리",       "원가 절감 실적, 경비 집행 현황, 예산 대비"),
    ("hr",        "인사·노무 현황",       "인원 변동, 교육 훈련, 노사 관계, 안전 사고"),
    ("investment","설비·투자",           "설비 투자 현황, 시설 보수, 신규 라인"),
    ("issues",    "당면 과제·계획",       "차월 중점 추진 사항, 리스크, 대응 계획"),
]


# ══════════════════════════════════════════════════════════════════════════
#  메인 화면
# ══════════════════════════════════════════════════════════════════════════

st.title("⑥ 운영실적 입력")
st.caption("주요업무 추진실적 서술형 보고서 | 손익 수치 + 업계동향을 참고하여 작성")

col1, col2, _ = st.columns([1, 1, 3])
with col1:
    year = st.selectbox("연도", range(2024, 2028), index=2, key="ops_year")
with col2:
    month = st.selectbox("월", range(1, 13), index=2, key="ops_month")

# 기존 데이터 로드
existing = load_monthly_operations(year, month)
existing_map = {r["section"]: r["content"] for r in existing}

st.divider()

# 작성 가이드
with st.expander("작성 가이드", expanded=False):
    st.markdown("""
    **작업 순서**
    1. 손익실적(②)과 노동생산성(⑤) 대시보드에서 당월 수치를 확인합니다
    2. 업계동향(①)에서 시장 맥락을 참고합니다
    3. 아래 섹션별로 서술형 실적을 작성합니다
    4. 저장 후 보고서 다운로드(⑦)에서 최종 취합합니다

    **작성 팁**
    - 정량적 실적(매출 ○○억, 전월 대비 +○%) 을 먼저 기술
    - 정성적 성과/이슈를 이어서 작성
    - 차월 계획까지 포함
    """)

# ── 섹션별 입력 ───────────────────────────────────────────────────────────
form_data = {}

for section_code, section_name, placeholder in SECTIONS:
    st.markdown(f"#### {section_name}")
    content = st.text_area(
        section_name,
        value=existing_map.get(section_code, ""),
        height=150,
        placeholder=placeholder,
        key=f"ops_{section_code}",
        label_visibility="collapsed",
    )
    form_data[section_code] = content

# ── 저장 ─────────────────────────────────────────────────────────────────
st.divider()

# 작성 현황 표시
filled = sum(1 for v in form_data.values() if v.strip())
total = len(SECTIONS)
st.progress(filled / total, text=f"작성 현황: {filled}/{total} 섹션")

col_btn, col_msg = st.columns([1, 3])
with col_btn:
    if st.button("저장", type="primary", use_container_width=True):
        items = []
        for section_code, section_name, _ in SECTIONS:
            items.append({
                "section": section_code,
                "section_name": section_name,
                "content": form_data.get(section_code, ""),
            })
        try:
            save_monthly_operations(year, month, items)
            st.success(f"{year}년 {month}월 운영실적 저장 완료!")
        except Exception as e:
            st.error(f"저장 실패: {e}")
