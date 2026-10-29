"""Meta 광고 라이브러리 API에서 광고를 검색해 JSONL로 저장한다.

    META_ACCESS_TOKEN=... python kv_reference/ingest/meta_ad_library.py --terms "화장품,커피" --countries KR --limit 2000 -o meta.jsonl

- 액세스 토큰: Meta 개발자 계정에서 신원 확인을 마친 뒤 발급 (https://www.facebook.com/ads/library/api)
- ad_type=ALL 검색은 EU에 노출된 광고와 정치·사회 이슈 광고만 결과가 나온다. 한국 상업 광고는 대부분 조회되지 않는다.
- 소재 이미지는 저장하지 않고 ad_snapshot_url(로그인 필요)만 링크한다.
"""

import argparse
import json
import os
import sys
import time

import requests

from common import record, write_jsonl

API = "https://graph.facebook.com/v21.0/ads_archive"
FIELDS = "id,page_name,ad_creative_bodies,ad_creative_link_titles,ad_delivery_start_time,ad_delivery_stop_time,ad_snapshot_url,publisher_platforms,languages"


def search(token, term, countries, limit, delay):
    params = {"access_token": token, "search_terms": term, "ad_reached_countries": json.dumps(countries), "ad_type": "ALL", "fields": FIELDS, "limit": 100}
    url, got = API, 0
    while url and got < limit:
        res = requests.get(url, params=params if url == API else None, timeout=30)
        data = res.json()
        if "error" in data:
            sys.exit("Meta API 오류: " + data["error"].get("message", ""))
        for ad in data.get("data", []):
            got += 1
            yield ad
        url = data.get("paging", {}).get("next")
        time.sleep(delay)


def to_record(ad, country):
    title = (ad.get("ad_creative_link_titles") or [""])[0] or (ad.get("ad_creative_bodies") or [""])[0][:80] or (ad.get("page_name", "") + " 광고")
    start = (ad.get("ad_delivery_start_time") or "")[:10]
    return record(
        id="meta-" + ad["id"], title=title, brand=ad.get("page_name", ""), year=int(start[:4]) if start[:4].isdigit() else 0, date=start, country=country,
        field="제품·서비스", medium=["digital"], note=((ad.get("ad_creative_bodies") or [""])[0] or "")[:300],
        sourceUrl=ad.get("ad_snapshot_url") or "", origin="meta", license="Meta Ad Library API",
    )


def main():
    ap = argparse.ArgumentParser(description="Meta 광고 라이브러리 → KV 아틀라스 JSONL")
    ap.add_argument("--terms", required=True, help="쉼표로 구분한 검색어")
    ap.add_argument("--countries", default="KR")
    ap.add_argument("--limit", type=int, default=1000, help="검색어당 최대 건수")
    ap.add_argument("--delay", type=float, default=1.0)
    ap.add_argument("-o", "--output", default="meta.jsonl")
    a = ap.parse_args()
    token = os.environ.get("META_ACCESS_TOKEN")
    if not token:
        sys.exit("META_ACCESS_TOKEN 환경 변수가 필요합니다.")
    countries = [c.strip().upper() for c in a.countries.split(",") if c.strip()]
    rows, seen = [], set()
    for term in [t.strip() for t in a.terms.split(",") if t.strip()]:
        for ad in search(token, term, countries, a.limit, max(0.5, a.delay)):
            if ad["id"] not in seen:
                seen.add(ad["id"])
                rows.append(to_record(ad, countries[0] if len(countries) == 1 else "XX"))
        print("%s: 누적 %d건" % (term, len(rows)), file=sys.stderr)
    print("%d건 저장: %s" % (write_jsonl(a.output, rows), a.output), file=sys.stderr)


if __name__ == "__main__":
    main()
