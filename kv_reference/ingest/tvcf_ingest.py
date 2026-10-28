"""TVCF 우수 인쇄광고 목록을 KV 아틀라스 가져오기용 JSON으로 저장한다.

TVCF에 접속되는 PC에서 실행한다. 이미지 파일은 내려받지 않고 제목·광고주·날짜·상세 링크·썸네일 URL만 모은다.
결과 JSON을 KV 아틀라스의 "데이터 > JSON 그대로 가져오기"에 붙여 넣으면 공유 DB에 등록된다.

    python kv_reference/ingest/tvcf_ingest.py --pages 3 -o tvcf.json
    python kv_reference/ingest/tvcf_ingest.py --url "https://tvcf.co.kr/hot/prominent-print?page=1&size=100&sort_by=start_date&media_type=2&country_code=410&lang=ko"

TVCF 페이지 구조는 바뀔 수 있다. 결과가 비면 --debug 로 원본 HTML을 저장해 구조를 확인하고 extract_* 함수를 고친다.
TVCF 이용약관을 확인하고, 요청 간격(기본 1.5초)을 줄이지 않는다.
"""

import argparse
import json
import re
import sys
import time
from urllib.parse import parse_qs, urlencode, urljoin, urlparse, urlunparse

import requests
from bs4 import BeautifulSoup

DEFAULT_URL = "https://tvcf.co.kr/hot/prominent-print?page=1&size=100&sort_by=start_date&media_type=2&country_code=410&lang=ko"
BASE = "https://tvcf.co.kr"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36",
    "Accept-Language": "ko-KR,ko;q=0.9,en;q=0.8",
}
# TVCF country_code(ISO 3166 숫자) → 알파2
NUMERIC_TO_ALPHA2 = {"410": "KR", "392": "JP", "156": "CN", "840": "US", "826": "GB", "250": "FR", "276": "DE"}

TITLE_KEYS = ("title", "subject", "name", "ad_title", "adTitle")
BRAND_KEYS = ("advertiser", "advertiserName", "advertiser_name", "brand", "brandName", "company", "client")
AGENCY_KEYS = ("agency", "agencyName", "agency_name", "production", "productionName")
DATE_KEYS = ("start_date", "startDate", "onair_date", "onairDate", "date", "reg_date", "regDate", "created_at")
IMAGE_KEYS = ("thumbnail", "thumb", "thumbnail_url", "thumbnailUrl", "image", "image_url", "imageUrl", "img")
ID_KEYS = ("id", "idx", "seq", "ad_id", "adId", "code")


def page_url(url, page):
    u = urlparse(url)
    q = parse_qs(u.query)
    q["page"] = [str(page)]
    return urlunparse(u._replace(query=urlencode({k: v[0] for k, v in q.items()})))


def first(d, keys):
    for k in keys:
        v = d.get(k)
        if v not in (None, ""):
            return v
    return ""


def as_text(v):
    if isinstance(v, dict):
        return str(first(v, ("name", "title", "ko", "value")) or "")
    return str(v or "").strip()


def walk_json(node, found):
    """페이지에 박힌 JSON(Next.js __NEXT_DATA__ 등)에서 광고처럼 보이는 객체를 모은다."""
    if isinstance(node, dict):
        if first(node, TITLE_KEYS) and (first(node, BRAND_KEYS) or first(node, IMAGE_KEYS) or first(node, DATE_KEYS)):
            found.append(node)
        for v in node.values():
            walk_json(v, found)
    elif isinstance(node, list):
        for v in node:
            walk_json(v, found)


def extract_from_json(soup):
    found = []
    for script in soup.find_all("script"):
        text = script.string or ""
        if script.get("id") == "__NEXT_DATA__" or script.get("type") == "application/json":
            try:
                walk_json(json.loads(text), found)
            except ValueError:
                continue
        elif "self.__next_f" in text or "window.__" in text:
            for m in re.finditer(r"\{[^{}]*\"(?:title|subject)\"[^{}]*\}", text.replace('\\"', '"')):
                try:
                    walk_json(json.loads(m.group(0)), found)
                except ValueError:
                    continue
    items = []
    for d in found:
        ident = as_text(first(d, ID_KEYS))
        img = as_text(first(d, IMAGE_KEYS))
        items.append({
            "title": as_text(first(d, TITLE_KEYS)),
            "brand": as_text(first(d, BRAND_KEYS)),
            "agency": as_text(first(d, AGENCY_KEYS)),
            "date": as_text(first(d, DATE_KEYS))[:10],
            "imageUrl": urljoin(BASE, img) if img else "",
            "sourceUrl": urljoin(BASE, as_text(d.get("url") or d.get("link") or d.get("href"))) if (d.get("url") or d.get("link") or d.get("href")) else "",
            "tvcfId": ident,
        })
    return items


def extract_from_html(soup):
    """JSON이 없을 때: 이미지가 들어 있는 링크 카드를 광고 하나로 본다."""
    items = []
    for a in soup.select("a[href]"):
        img = a.find("img")
        if not img:
            continue
        text = " ".join(a.get_text(" ", strip=True).split())
        title = img.get("alt") or text
        if not title or len(title) < 2:
            continue
        date = re.search(r"(20\d{2})[.\-/](\d{1,2})[.\-/](\d{1,2})", text)
        items.append({
            "title": title.strip()[:200],
            "brand": "",
            "agency": "",
            "date": "%s-%02d-%02d" % (date.group(1), int(date.group(2)), int(date.group(3))) if date else "",
            "imageUrl": urljoin(BASE, img.get("src") or img.get("data-src") or ""),
            "sourceUrl": urljoin(BASE, a["href"]),
            "note": text[:300],
        })
    return items


def scrape(url, pages, delay, debug):
    country = NUMERIC_TO_ALPHA2.get(parse_qs(urlparse(url).query).get("country_code", [""])[0], "KR")
    session = requests.Session()
    session.headers.update(HEADERS)
    out, seen = [], set()
    for page in range(1, pages + 1):
        target = page_url(url, page)
        print("가져오는 중: %s" % target, file=sys.stderr)
        res = session.get(target, timeout=20)
        res.raise_for_status()
        if debug:
            with open("debug_tvcf_p%d.html" % page, "w", encoding="utf-8") as f:
                f.write(res.text)
        soup = BeautifulSoup(res.text, "lxml")
        items = extract_from_json(soup) or extract_from_html(soup)
        new = 0
        for it in items:
            key = it["sourceUrl"] or it["title"]
            if key in seen or not it["title"]:
                continue
            seen.add(key)
            it.update({"country": country, "medium": ["print"], "origin": "tvcf", "field": "브랜드"})
            if it["date"]:
                it["year"] = int(it["date"][:4])
            out.append(it)
            new += 1
        print("  %d건" % new, file=sys.stderr)
        if not new:
            break
        time.sleep(delay)
    return out


def main():
    ap = argparse.ArgumentParser(description="TVCF 우수 인쇄광고 목록 → KV 아틀라스 JSON")
    ap.add_argument("--url", default=DEFAULT_URL, help="TVCF 목록 URL (page 파라미터는 자동으로 바뀐다)")
    ap.add_argument("--pages", type=int, default=1, help="가져올 페이지 수")
    ap.add_argument("--delay", type=float, default=1.5, help="요청 간격(초), 1초 이상 유지")
    ap.add_argument("-o", "--output", default="tvcf_print.json")
    ap.add_argument("--debug", action="store_true", help="받은 HTML을 debug_tvcf_p*.html 로 저장")
    args = ap.parse_args()
    items = scrape(args.url, args.pages, max(1.0, args.delay), args.debug)
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False, indent=1)
    print("%d건을 %s 에 저장했습니다." % (len(items), args.output), file=sys.stderr)
    if not items:
        print("광고를 찾지 못했습니다. --debug 로 HTML을 저장해 구조를 확인하세요.", file=sys.stderr)


if __name__ == "__main__":
    main()
