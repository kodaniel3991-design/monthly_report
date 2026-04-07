"""
다나와 자동차 판매현황 스크래핑
TOP10 모델별 판매량 + 제조사별 시장점유율 산출
"""

import ssl
import re
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.ssl_ import create_urllib3_context
from bs4 import BeautifulSoup


# 모델 → 제조사 매핑 (국내 주요 모델)
MODEL_MAKER = {
    "쏘렌토": "기아", "그랜저": "현대", "소나타 디 엣지": "현대", "소나타": "현대",
    "스포티지": "기아", "카니발": "기아", "아반떼": "현대", "셀토스": "기아",
    "디 올 뉴 셀토스": "기아", "필랑트": "르노코리아", "포터2": "현대", "포터": "현대",
    "EV3": "기아", "투싼": "현대", "G80": "현대", "코나": "현대",
    "싼타페": "현대", "팰리세이드": "현대", "K8": "기아", "모닝": "기아",
    "레이": "기아", "봉고3": "기아", "스타리아": "현대", "캐스퍼": "현대",
    "GV70": "현대", "GV80": "현대", "G70": "현대", "G90": "현대",
    "아이오닉5": "현대", "아이오닉 5": "현대", "아이오닉6": "현대", "아이오닉 6": "현대",
    "EV6": "기아", "EV9": "기아", "니로": "기아",
    "트레일블레이저": "한국GM", "트랙스 크로스오버": "한국GM", "이쿼녹스": "한국GM",
    "콜로라도": "한국GM", "타호": "한국GM",
    "그랑 콜레오스": "르노코리아", "아르카나": "르노코리아", "QM6": "르노코리아",
    "XM3": "르노코리아", "폴스타4": "르노코리아",
    "토레스": "KG모빌리티", "티볼리": "KG모빌리티", "코란도": "KG모빌리티",
    "렉스턴": "KG모빌리티", "액티언": "KG모빌리티",
}

# 제조사 → 통합명
MAKER_NORMALIZE = {
    "현대": "현대", "기아": "기아", "한국GM": "GM",
    "르노코리아": "르노코리아", "KG모빌리티": "KG모빌리티",
}


class _DanawaSSL(HTTPAdapter):
    def init_poolmanager(self, *a, **kw):
        ctx = create_urllib3_context()
        ctx.check_hostname = False
        ctx.verify_mode = ssl.CERT_NONE
        ctx.set_ciphers("DEFAULT@SECLEVEL=1")
        kw["ssl_context"] = ctx
        return super().init_poolmanager(*a, **kw)


def scrape_danawa(year: int, month: int,
                   url_top: str = None, url_share: str = None) -> dict:
    """
    다나와 자동차에서 월별 판매 데이터 스크래핑
    url_top: 판매순위 TOP10 페이지 URL (없으면 자동 생성)
    url_share: 시장점유율 페이지 URL (없으면 자동 생성)
    반환: {
        "top10": [{"rank", "model", "maker", "sales"}, ...],
        "market_share": {"현대": 47.4, "기아": 43.3, ...},
        "total_sales": 129000,
    }
    """
    sess = requests.Session()
    sess.mount("https://", _DanawaSSL())

    url = url_top or (
        f"https://auto.danawa.com/newcar/?Work=record&Tab=Grand"
        f"&Classify=~C,PC1,PC2,PC3,PC4,PC5,PS~,RU2,RU3,RU5,RM~,O~"
        f"&Month={year}-{month:02d}-00&MonthTo="
    )

    resp = sess.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=10)
    resp.encoding = "utf-8"
    soup = BeautifulSoup(resp.text, "html.parser")

    table = soup.find("table")
    if not table:
        return {"top10": [], "market_share": {}, "total_sales": 0}

    # 전체 모델 파싱
    all_models = []
    for tr in table.find_all("tr"):
        tds = tr.find_all("td")
        if len(tds) < 5:
            continue

        rank_text = tds[1].get_text(strip=True)
        if not rank_text.isdigit():
            continue

        rank = int(rank_text)
        model_name = tds[3].get_text(strip=True)

        # 판매량 — "10,870그래프로 보기" → 10870
        sales_raw = tds[4].get_text(strip=True)
        sales_num = re.sub(r"[^0-9]", "", sales_raw.split("그")[0])
        sales = int(sales_num) if sales_num else 0

        # 점유율
        share_text = tds[5].get_text(strip=True) if len(tds) > 5 else ""
        share = 0.0
        share_match = re.search(r"([\d.]+)%", share_text)
        if share_match:
            share = float(share_match.group(1))

        # 제조사 매핑
        maker = "기타"
        for key, val in MODEL_MAKER.items():
            if key in model_name:
                maker = val
                break

        all_models.append({
            "rank": rank,
            "model": model_name,
            "maker": maker,
            "sales": sales,
            "share": share,
        })

    # TOP 10
    top10 = [m for m in all_models if m["rank"] <= 10]

    # 제조사별 합산 (전체 모델 기준)
    maker_totals = {}
    for m in all_models:
        norm = MAKER_NORMALIZE.get(m["maker"], m["maker"])
        maker_totals[norm] = maker_totals.get(norm, 0) + m["sales"]

    total_sales = sum(maker_totals.values())

    # 점유율 계산
    market_share = {}
    if total_sales > 0:
        for mk, sv in maker_totals.items():
            market_share[mk] = round(sv / total_sales * 100, 1)

    return {
        "top10": top10,
        "market_share": market_share,
        "total_sales": total_sales,
    }


if __name__ == "__main__":
    data = scrape_danawa(2026, 3)
    print("=== TOP 10 ===")
    for t in data["top10"]:
        print(f"  {t['rank']}. {t['model']} ({t['maker']}) - {t['sales']:,}")
    print(f"\n=== 시장점유율 ===")
    for mk, pct in sorted(data["market_share"].items(), key=lambda x: -x[1]):
        print(f"  {mk}: {pct}%")
    print(f"  합계: {sum(data['market_share'].values()):.1f}%")
    print(f"  총 판매: {data['total_sales']:,}대")
