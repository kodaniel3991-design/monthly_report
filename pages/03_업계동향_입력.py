"""
업계동향 입력 페이지
르노코리아·GM·현대차 뉴스 + 국내판매 TOP10 + 시장점유율
"""

import streamlit as st
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
sys.path.insert(0, str(Path(__file__).parent))
from database import (
    save_industry_news, load_industry_news,
    get_conn
)
from flow_bar import render_flow_bar
from naver_news import search_news, search_news_with_summary
from danawa_scraper import scrape_danawa

st.set_page_config(page_title="업계동향 입력", layout="wide")

render_flow_bar(current_step=1)

st.markdown("**① 업계동향 입력** — 자동차 업계 주요 뉴스 및 시장 현황")

c1, c2 = st.columns([1, 1])
with c1:
    year  = st.selectbox("연도", range(2024, 2028), index=2)
with c2:
    month = st.selectbox("월",   range(1, 13),       index=2)

existing_news = load_industry_news(year, month)

def get_news(company, seq):
    for n in existing_news:
        if n["company"] == company and n["seq"] == seq:
            return n
    return {}

st.divider()

# ── 탭 구성 ───────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "르노코리아", "GM Korea", "현대자동차", "⚡ 업계이슈", "📊 판매현황·M/S"
])

news_items = []

def wrap_text(text, width=34):
    """보고서 양식에 맞게 텍스트 줄바꿈
    - 첫 줄: ①② 번호 + 내용 (들여쓰기 4칸 포함 약 34자)
    - 이후 줄: 왼쪽 정렬 약 34자
    - 단어 중간 안 잘림 (공백·마침표·쉼표 기준)
    """
    result = []
    for para in text.split("\n"):
        para = para.strip()
        if not para:
            continue
        is_first = True
        while len(para) > width:
            # 자를 위치 찾기: 뒤에서부터 자연스러운 끊김점
            cut = width
            # 마침표+공백, 쉼표+공백, 공백 순서로 우선
            best = -1
            for sep in [". ", ", ", " "]:
                pos = para[:width].rfind(sep)
                if pos > width * 0.4:
                    best = pos + len(sep)
                    break
            if best > 0:
                cut = best

            line = para[:cut].rstrip()
            result.append(line)
            para = para[cut:].lstrip()
            is_first = False
        if para:
            result.append(para)
    return "\n".join(result)

SEARCH_KEYWORDS = {
    "르노코리아": "르노코리아 자동차",
    "GM Korea":   "GM 한국 자동차",
    "현대자동차":  "현대자동차",
    "업계이슈":    "자동차 업계 동향",
}

COMPANIES = [
    (tab1, "르노코리아",  5),
    (tab2, "GM Korea",   5),
    (tab3, "현대자동차",  5),
    (tab4, "업계이슈",    5),
]

for tab, company, max_news in COMPANIES:
    with tab:
        # ── 뉴스 검색 ──
        search_col, btn_col = st.columns([3, 1])
        with search_col:
            keyword = st.text_input(
                "검색어",
                value=SEARCH_KEYWORDS.get(company, company),
                key=f"kw_{company}",
                label_visibility="collapsed",
                placeholder=f"{company} 관련 검색어 입력",
            )
        with btn_col:
            do_search = st.button("🔍 뉴스 검색", key=f"search_{company}",
                                   use_container_width=True)

        # 검색 실행
        search_key = f"sr_{company}"
        if do_search and keyword.strip():
            results = search_news_with_summary(keyword.strip(), display=10, max_sentences=5)
            st.session_state[search_key] = results

        # 검색 결과 표시 → 체크하면 자동 입력
        results = st.session_state.get(search_key, [])
        selected_key = f"sel_{company}"
        if selected_key not in st.session_state:
            st.session_state[selected_key] = {}

        if results:
            st.markdown(f"<small style='color:#8a6e62;'>검색 결과 {len(results)}건 — 사용할 뉴스를 선택하세요</small>",
                        unsafe_allow_html=True)
            for idx, r in enumerate(results):
                checked = st.checkbox(
                    f"**{r['title'][:60]}** · {r['source']}",
                    key=f"chk_{company}_{idx}",
                    value=idx in st.session_state[selected_key],
                )
                if checked and idx not in st.session_state[selected_key]:
                    st.session_state[selected_key][idx] = r
                    # session_state에 줄바꿈된 본문 직접 세팅
                    seq_num = len(st.session_state[selected_key])
                    st.session_state[f"ct_{company}_{seq_num}"] = wrap_text(f"  ① {r.get('summary', r['description'])}", width=34)
                    st.session_state[f"hl_{company}_{seq_num}"] = f"▷ {r['title']}"
                    st.session_state[f"src_{company}_{seq_num}"] = r["source"]
                    st.rerun()
                elif not checked and idx in st.session_state[selected_key]:
                    del st.session_state[selected_key][idx]
                    st.rerun()

        st.markdown("<div style='height:1px; background:#dddbd7; margin:8px 0;'></div>",
                    unsafe_allow_html=True)

        # ── 선택된 뉴스 + 기존 뉴스 편집 ──
        selected = st.session_state.get(selected_key, {})
        existing_count = len([n for n in existing_news if n["company"] == company])

        st.markdown(f"**등록된 뉴스** ({len(selected) + existing_count}건)")

        seq = 0
        # 검색에서 선택된 뉴스
        for idx, r in sorted(selected.items()):
            seq += 1
            col_a, col_b = st.columns([3, 1])
            with col_a:
                headline = st.text_input(
                    f"뉴스 {seq}",
                    value=f"▷ {r['title']}",
                    key=f"hl_{company}_{seq}",
                )
            with col_b:
                source = st.text_input(
                    "출처" if seq == 1 else " ",
                    value=r["source"],
                    key=f"src_{company}_{seq}",
                )
            # 본문 — ① 번호 + 45자 기준 줄바꿈
            default_content = wrap_text(f"  ① {r.get('summary', r['description'])}", width=34)
            content = st.text_area(
                "내용",
                value=default_content,
                key=f"ct_{company}_{seq}",
                height=120,
                label_visibility="collapsed",
                help="줄바꿈하여 ①②③ 등 번호를 붙여 작성하세요. 보고서에 그대로 출력됩니다.",
            )
            if headline.strip() and headline.strip() != "▷":
                news_items.append({
                    "company": company,
                    "headline": headline,
                    "content": content,
                    "source": source,
                    "seq": seq,
                })

        # 기존 DB 뉴스
        if not selected:
            for n in existing_news:
                if n["company"] != company:
                    continue
                seq += 1
                col_a, col_b = st.columns([3, 1])
                with col_a:
                    headline = st.text_input(
                        f"뉴스 {seq}",
                        value=n.get("headline", ""),
                        key=f"hl_{company}_{seq}",
                    )
                with col_b:
                    source = st.text_input(
                        "출처" if seq == 1 else " ",
                        value=n.get("source", ""),
                        key=f"src_{company}_{seq}",
                    )
                raw_content = n.get("content", "")
                # 불완전 문장(...) 제거 + 완료형 문장으로 마무리
                if raw_content:
                    import re as _re
                    # ...로 끝나는 줄 제거
                    lines = raw_content.split("\n")
                    clean_lines = [l for l in lines if not l.rstrip().endswith("...") and not l.rstrip().endswith("…")]
                    # 날짜/시간 라인 제거
                    clean_lines = [l for l in clean_lines if not _re.match(r'^[:\s]*\d{4}\.\d{2}\.\d{2}', l.strip())]
                    raw_content = "\n".join(clean_lines)
                    # 줄바꿈이 안 되어있으면 자동 줄바꿈
                    if "\n" not in raw_content:
                        raw_content = wrap_text(raw_content, width=34)
                content = st.text_area(
                    "내용",
                    value=raw_content,
                    key=f"ct_{company}_{seq}",
                    height=120,
                    label_visibility="collapsed",
                    help="줄바꿈하여 ①②③ 등 번호를 붙여 작성하세요. 보고서에 그대로 출력됩니다.",
                )
                if headline.strip() and headline.strip() != "▷":
                    news_items.append({
                        "company": company,
                        "headline": headline,
                        "content": content,
                        "source": source,
                        "seq": seq,
                    })

        # ── 보고서 미리보기 ──
        if news_items and any(n["company"] == company for n in news_items):
            with st.expander("📄 보고서 미리보기", expanded=False):
                for n in news_items:
                    if n["company"] != company:
                        continue
                    st.markdown(
                        f"<div style='margin-bottom:16px;'>"
                        f"<div style='font-size:14px; font-weight:700;'>"
                        f"{n['headline']} "
                        f"<span style='font-size:12px; font-weight:400; color:#8a6e62;'>"
                        f"〈{n['source']}〉</span></div>"
                        f"<div style='font-size:13px; color:#424547; line-height:1.8; "
                        f"margin-top:6px; padding-left:16px; white-space:pre-wrap;'>"
                        f"{n['content']}</div>"
                        f"</div>",
                        unsafe_allow_html=True
                    )

# ── 판매현황 탭 ───────────────────────────────────────────────────────────
with tab5:
    # ── 다나와 자동 스크래핑 (URL 2개 입력) ──
    if True:
        st.caption("다나와 auto.danawa.com에서 해당 월의 판매순위 + 점유율 페이지 URL을 입력하세요.")

        url_top = st.text_input(
            "판매순위 TOP10 페이지 URL",
            value=f"https://auto.danawa.com/newcar/?Work=record&Tab=Grand"
                  f"&Classify=~C,PC1,PC2,PC3,PC4,PC5,PS~,RU2,RU3,RU5,RM~,O~"
                  f"&Month={year}-{month:02d}-00&MonthTo=",
            key="danawa_url_top",
            help="다나와 > 신차 > 판매실적 > 종합 탭의 URL"
        )
        url_share = st.text_input(
            "시장점유율 페이지 URL",
            value=f"https://auto.danawa.com/newcar/?Work=record&Tab=BrandMaker"
                  f"&Month={year}-{month:02d}-00&MonthTo=",
            key="danawa_url_share",
            help="다나와 > 신차 > 판매실적 > 제조사별 탭의 URL"
        )

        do_scrape = st.button("다나와에서 자동 수집", key="scrape_danawa",
                               use_container_width=True, type="primary")

        if do_scrape:
            with st.spinner("다나와 데이터 수집 중..."):
                try:
                    danawa = scrape_danawa(year, month,
                                           url_top=url_top if url_top else None,
                                           url_share=url_share if url_share else None)
                    st.session_state["danawa_data"] = danawa
                    st.success(f"수집 완료! TOP {len(danawa['top10'])}개 모델, 총 {danawa['total_sales']:,}대")
                    st.rerun()
                except Exception as e:
                    st.error(f"수집 실패: {e}")

    danawa = st.session_state.get("danawa_data", None)

    # ── 시장점유율 (가로 배치) ──
    st.markdown("**시장점유율 (%)**")

    conn = get_conn()
    existing_ms = conn.execute(
        "SELECT * FROM monthly_market_share WHERE year=? AND month=?",
        (year, month)
    ).fetchall()
    conn.close()
    ms_map = {r["company"]: r["share_pct"] for r in existing_ms}

    # 다나와 데이터가 있으면 덮어쓰기
    if danawa and danawa.get("market_share"):
        ms_map = danawa["market_share"]

    ms_companies = ["현대", "기아", "GM", "르노코리아", "KG모빌리티"]
    ms_data = {}
    ms_cols = st.columns(len(ms_companies))
    for i, comp in enumerate(ms_companies):
        with ms_cols[i]:
            ms_data[comp] = st.number_input(
                comp, value=float(ms_map.get(comp, 0)),
                min_value=0.0, max_value=100.0, step=0.1,
                format="%.1f", key=f"ms_{comp}"
            )

    total_ms = sum(ms_data.values())
    color = "green" if abs(total_ms - 100) < 0.2 else "red"
    st.markdown(f"합계: :{color}[**{total_ms:.1f}%**]"
                + (" ✓" if abs(total_ms - 100) < 0.2 else " ← 100%가 되어야 합니다"))

    st.markdown("<div style='height:1px; background:#dddbd7; margin:8px 0;'></div>",
                unsafe_allow_html=True)

    # ── 국내 판매 TOP 10 ──
    st.markdown("**국내 판매 TOP 10** · 수입차 제외")

    conn = get_conn()
    existing_top = conn.execute(
        "SELECT * FROM monthly_top_models WHERE year=? AND month=? ORDER BY rank",
        (year, month)
    ).fetchall()
    conn.close()
    top_map = {r["rank"]: dict(r) for r in existing_top}

    # 다나와 데이터가 있으면 덮어쓰기
    if danawa and danawa.get("top10"):
        top_map = {}
        for t in danawa["top10"]:
            top_map[t["rank"]] = {
                "model_name": t["model"],
                "company": t["maker"],
                "sales_qty": t["sales"],
            }

    top_models = []
    ch1, ch2, ch3, ch4 = st.columns([0.3, 2, 2, 1])
    ch1.markdown("<small style='color:#a8a9aa;'>순위</small>", unsafe_allow_html=True)
    ch2.markdown("<small style='color:#a8a9aa;'>모델명</small>", unsafe_allow_html=True)
    ch3.markdown("<small style='color:#a8a9aa;'>소속사</small>", unsafe_allow_html=True)
    ch4.markdown("<small style='color:#a8a9aa;'>판매(대)</small>", unsafe_allow_html=True)
    for rank in range(1, 11):
        t = top_map.get(rank, {})
        c1t, c2t, c3t, c4t = st.columns([0.3, 2, 2, 1])
        with c1t:
            st.markdown(f"<div style='padding-top:4px; font-size:12px; font-weight:600; color:#8a6e62;'>{rank}</div>",
                        unsafe_allow_html=True)
        with c2t:
            model = st.text_input("m", value=t.get("model_name",""),
                                  key=f"tm_{rank}", label_visibility="collapsed")
        with c3t:
            comp  = st.text_input("c", value=t.get("company",""),
                                  key=f"tc_{rank}", label_visibility="collapsed")
        with c4t:
            qty   = st.number_input("q", value=int(t.get("sales_qty",0)),
                                    step=100, key=f"tq_{rank}",
                                    label_visibility="collapsed")
        if model:
            top_models.append({"rank": rank, "model": model,
                               "company": comp, "qty": qty})

# ── 저장 ─────────────────────────────────────────────────────────────────
st.divider()
if st.button("💾 전체 저장", type="primary"):
    try:
        # 업계동향 뉴스
        save_industry_news(year, month, news_items)

        # TOP10
        conn = get_conn()
        conn.execute("DELETE FROM monthly_top_models WHERE year=? AND month=?", (year, month))
        for t in top_models:
            conn.execute("""
                INSERT INTO monthly_top_models (year, month, rank, model_name, company, sales_qty)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (year, month, t["rank"], t["model"], t["company"], t["qty"]))

        # 시장점유율
        conn.execute("DELETE FROM monthly_market_share WHERE year=? AND month=?", (year, month))
        for comp, pct in ms_data.items():
            conn.execute("""
                INSERT INTO monthly_market_share (year, month, company, share_pct)
                VALUES (?, ?, ?, ?)
            """, (year, month, comp, pct))
        conn.commit()
        conn.close()

        st.success(f"✅ {year}년 {month}월 업계동향 저장 완료 ({len(news_items)}건)")
        pass
    except Exception as e:
        st.error(f"저장 실패: {e}")
