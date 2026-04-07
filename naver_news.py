"""
네이버 뉴스 검색 API 연동
"""

import os
import urllib.request
import urllib.parse
import json
import re
from pathlib import Path

# .env에서 키 로드
_env_path = Path(__file__).parent / ".env"
if _env_path.exists():
    for line in _env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and "=" in line and not line.startswith("#"):
            k, v = line.split("=", 1)
            os.environ.setdefault(k.strip(), v.strip())

CLIENT_ID = os.environ.get("NAVER_CLIENT_ID", "")
CLIENT_SECRET = os.environ.get("NAVER_CLIENT_SECRET", "")


def _clean_html(text: str) -> str:
    """HTML 태그 및 특수문자 제거"""
    text = re.sub(r"<[^>]+>", "", text)
    text = text.replace("&quot;", '"').replace("&amp;", "&")
    text = text.replace("&lt;", "<").replace("&gt;", ">")
    text = text.replace("&apos;", "'")
    return text.strip()


def _scrape_article(url: str, max_sentences: int = 5) -> str:
    """뉴스 URL에서 본문을 스크래핑하고 핵심 문장 추출 (최대 max_sentences줄)"""
    try:
        import ssl
        import requests as _req
        from requests.adapters import HTTPAdapter
        from urllib3.util.ssl_ import create_urllib3_context

        class _SSL(HTTPAdapter):
            def init_poolmanager(self, *a, **kw):
                ctx = create_urllib3_context()
                ctx.check_hostname = False
                ctx.verify_mode = ssl.CERT_NONE
                ctx.set_ciphers("DEFAULT@SECLEVEL=1")
                kw["ssl_context"] = ctx
                return super().init_poolmanager(*a, **kw)

        sess = _req.Session()
        sess.mount("https://", _SSL())

        resp = sess.get(url, headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"},
                        timeout=5, verify=False)
        resp.encoding = "utf-8"

        from bs4 import BeautifulSoup
        soup = BeautifulSoup(resp.text, "html.parser")

        # 네이버 뉴스 본문
        body = soup.select_one("#dic_area, #articleBodyContents, .article_body, #newsEndContents, "
                                "article, .news_end, #articeBody, .article-body")
        if not body:
            # 일반 뉴스 사이트 본문
            for sel in ["article", ".article", "#content", ".post-content", ".entry-content"]:
                body = soup.select_one(sel)
                if body:
                    break

        if not body:
            return ""

        # 불필요 태그 제거
        for tag in body.select("script, style, iframe, .ad, .promotion, .reporter_area, "
                                ".copyright, .byline, figure, .image, .photo"):
            tag.decompose()

        text = body.get_text("\n", strip=True)

        # 문장 분리 및 핵심 추출
        sentences = []
        for line in text.split("\n"):
            line = line.strip()
            if len(line) < 10:
                continue
            # 광고/기자/날짜/메타 등 불필요 라인 필터링
            skip_words = ["기자", "무단전재", "저작권", "ⓒ", "©", "제보", "구독", "댓글",
                          "좋아요", "공유", "카카오", "페이스북", "트위터", "URL", "클릭",
                          "입력", "수정", "송고", "뉴스1", "연합뉴스", "사진=", "사진 =",
                          "영상=", "취재=", "발행일", "등록일", "게시일",
                          "MBC", "KBS", "SBS", "JTBC", "YTN", "TV조선", "채널A",
                          "참 좋다", "뉴스데스크", "앵커", "리포트"]
            if any(w in line for w in skip_words):
                continue
            # 날짜/시간/메타 패턴 라인 제거
            if re.match(r'^[:\s]*\d{4}\.\d{2}\.\d{2}', line):
                continue
            if re.match(r'^[\d.\-:\s오전후]+$', line):
                continue
            # "2026년 04월 01일", "2026-04-07 (화)" 등 한글 날짜
            if re.match(r'^.*\d{4}년\s*\d{1,2}월\s*\d{1,2}일', line):
                continue
            if re.match(r'^.*\d{4}-\d{2}-\d{2}\s*\(', line):
                continue
            # [프로그램명] 패턴
            if re.match(r'^\[.+\]', line):
                continue
            # 라인 내 날짜/시간 잔여 제거
            line = re.sub(r':?\s*\d{4}\.\d{2}\.\d{2}\s*(오전|오후)?\s*\d{0,2}:?\d{0,2}', '', line)
            line = re.sub(r'\d{4}년\s*\d{1,2}월\s*\d{1,2}일', '', line)
            line = re.sub(r'\d{4}-\d{2}-\d{2}\s*\(\S\)', '', line)
            line = re.sub(r'\d{2}:\d{2}\s*(오전|오후)?', '', line)
            line = line.strip(' :')
            if len(line) < 15:
                continue
            # 마침표로 문장 분리
            for sent in re.split(r'(?<=다\.)\s*', line):
                sent = sent.strip()
                if len(sent) < 15:
                    continue
                # 불완전 문장 제거: ...으로 끝나거나, 완료형이 아닌 문장
                if sent.endswith("...") or sent.endswith("…"):
                    continue
                sentences.append(sent)

        # 서술형 요약: 완료형 문장(~다.)으로 끝나도록 조합
        max_chars = 34 * max_sentences  # 34자 × 5줄 = 170자
        combined = ""
        used = 0
        for sent in sentences:
            # 중복 방지
            if combined and sent[:15] in combined:
                continue
            if len(combined) + len(sent) + 1 > max_chars:
                break  # 잘리는 문장 넣지 않음 — 완료형 유지
            combined += (" " if combined else "") + sent
            used += 1
            if used >= max_sentences:
                break

        # 완료형 문장으로 끝나지 않으면 마지막 완료형까지 자르기
        if combined and not re.search(r'[다음됨임함짐][\.\!]?\s*$', combined):
            # 마지막 완료형 문장 끝 위치 찾기
            m = list(re.finditer(r'[다음됨임함짐]\.', combined))
            if m:
                combined = combined[:m[-1].end()]

        return combined.strip() if combined else ""

    except Exception:
        return ""


def search_news(query: str, display: int = 10, sort: str = "date") -> list[dict]:
    """
    네이버 뉴스 검색
    query: 검색어 (예: "르노코리아 판매")
    display: 결과 수 (최대 100)
    sort: "date" (최신순) 또는 "sim" (관련도순)
    반환: [{"title", "description", "link", "pubDate"}, ...]
    """
    if not CLIENT_ID or not CLIENT_SECRET:
        return []

    enc_query = urllib.parse.quote(query)
    url = f"https://openapi.naver.com/v1/search/news.json?query={enc_query}&display={display}&sort={sort}"

    req = urllib.request.Request(url)
    req.add_header("X-Naver-Client-Id", CLIENT_ID)
    req.add_header("X-Naver-Client-Secret", CLIENT_SECRET)

    try:
        with urllib.request.urlopen(req, timeout=5) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except Exception:
        return []

    results = []
    for item in data.get("items", []):
        # pubDate: "Mon, 07 Apr 2025 10:30:00 +0900" → "04.07"
        pub = item.get("pubDate", "")
        date_short = ""
        if pub:
            parts = pub.split()
            if len(parts) >= 4:
                months = {"Jan":"01","Feb":"02","Mar":"03","Apr":"04","May":"05","Jun":"06",
                          "Jul":"07","Aug":"08","Sep":"09","Oct":"10","Nov":"11","Dec":"12"}
                m = months.get(parts[2], "00")
                d = parts[1].rstrip(",")
                date_short = f"{m}.{d.zfill(2)}"

        # 출처 추출 (link에서 도메인)
        link = item.get("originallink", item.get("link", ""))
        source = ""
        try:
            from urllib.parse import urlparse
            domain = urlparse(link).netloc
            source = domain.replace("www.", "").split(".")[0]
        except Exception:
            pass

        results.append({
            "title": _clean_html(item.get("title", "")),
            "description": _clean_html(item.get("description", "")),
            "link": link,
            "pubDate": date_short,
            "source": f"{date_short} {source}" if source else date_short,
        })

    return results


def search_news_with_summary(query: str, display: int = 10, sort: str = "date",
                              max_sentences: int = 5) -> list[dict]:
    """뉴스 검색 + 각 기사 본문에서 핵심 5문장 요약 추가"""
    results = search_news(query, display, sort)
    for r in results:
        link = r.get("link", "")
        if link:
            summary = _scrape_article(link, max_sentences=max_sentences)
            r["summary"] = summary if summary else r["description"]
        else:
            r["summary"] = r["description"]
    return results
