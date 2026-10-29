"""Google 광고 투명성 센터 공개 데이터셋(BigQuery)에서 이미지 광고 소재를 JSONL로 가져온다.

    python kv_reference/ingest/google_ads_transparency.py --countries KR,JP,US --per-country 50000 --since 2026-01-01 -o gads.jsonl

필요한 것
- bq CLI 와 Google Cloud 인증 (서비스 계정 키: GOOGLE_APPLICATION_CREDENTIALS, 또는 gcloud auth)
- 쿼리 비용을 낼 프로젝트: --project 또는 gcloud 기본 프로젝트. 공개 데이터셋도 스캔한 바이트만큼 요금(월 1TB 무료)이 든다.

테이블 스키마를 먼저 읽고 있는 컬럼만 쓴다. 실행 전 --dry-run 으로 스캔량을 확인한다(기본 20GB 넘으면 --yes 없이는 멈춘다).
소재 이미지는 데이터셋에 없다. 각 레코드의 sourceUrl 이 투명성 센터의 소재 페이지다.
"""

import argparse
import json
import subprocess
import sys

from common import record, write_jsonl

TABLE = "bigquery-public-data.google_ads_transparency_center.creative_stats"


def bq(args, project):
    cmd = ["bq", "--format=json"] + (["--project_id=" + project] if project else []) + args
    res = subprocess.run(cmd, capture_output=True, text=True)
    if res.returncode != 0:
        sys.exit("bq 실패: " + (res.stderr or res.stdout).strip()[:800])
    return res.stdout


def schema(project):
    fields = json.loads(bq(["show", "--schema", TABLE.replace(".", ":", 1)], project))
    def walk(fs, prefix=""):
        out = {}
        for f in fs:
            out[prefix + f["name"]] = f
            if f.get("fields"):
                out.update(walk(f["fields"], prefix + f["name"] + "."))
        return out
    return walk(fields)


def pick(cols, *names):
    for n in names:
        if n in cols:
            return n
    return None


def build_sql(cols, country, since, limit):
    name = pick(cols, "advertiser_disclosed_name", "advertiser_legal_name", "advertiser_name")
    cid = pick(cols, "creative_id")
    url = pick(cols, "creative_page_url")
    fmt = pick(cols, "ad_format_type", "format")
    topic = pick(cols, "topic")
    reg = pick(cols, "region_stats")
    if not (name and cid and fmt and reg):
        sys.exit("예상한 컬럼이 없습니다. 스키마: " + ", ".join(sorted(cols)))
    rc = pick(cols, reg + ".region_code", reg + ".region")
    first = pick(cols, reg + ".first_shown", reg + ".first_shown_date")
    last = pick(cols, reg + ".last_shown", reg + ".last_shown_date")
    sel = [name + " AS advertiser", cid + " AS creative_id", fmt + " AS fmt", "r." + rc.split(".")[-1] + " AS region"]
    sel += [url + " AS url"] if url else []
    sel += [topic + " AS topic"] if topic else []
    sel += ["CAST(r." + first.split(".")[-1] + " AS STRING) AS first_shown"] if first else []
    sel += ["CAST(r." + last.split(".")[-1] + " AS STRING) AS last_shown"] if last else []
    where = ["UPPER(" + fmt + ") = 'IMAGE'", "r." + rc.split(".")[-1] + " = '" + country + "'"]
    if since and first:
        where.append("DATE(r." + first.split(".")[-1] + ") >= DATE('" + since + "')")
    order = " ORDER BY r." + last.split(".")[-1] + " DESC" if last else ""
    return "SELECT " + ", ".join(sel) + " FROM `" + TABLE + "`, UNNEST(" + reg + ") AS r WHERE " + " AND ".join(where) + order + " LIMIT " + str(int(limit))


def to_record(row, country):
    adv = (row.get("advertiser") or "").strip()
    first, last = (row.get("first_shown") or "")[:10], (row.get("last_shown") or "")[:10]
    return record(
        id="gads-" + row["creative_id"] + "-" + country.lower(), title=adv + " 이미지 광고", brand=adv, year=int(first[:4]) if first[:4].isdigit() else 0,
        date=last or first, country=country, field="제품·서비스", industry="기타", medium=["digital"],
        note="Google 광고 투명성 센터 · 노출 " + (first + " ~ " + last if first else "기간 미상") + (" · 주제 " + row["topic"] if row.get("topic") else ""),
        sourceUrl=row.get("url") or "https://adstransparency.google.com/", origin="gads", license="Google Ads Transparency Center public dataset",
    )


def main():
    ap = argparse.ArgumentParser(description="Google 광고 투명성 센터(BigQuery) → KV 아틀라스 JSONL")
    ap.add_argument("--countries", default="KR", help="쉼표로 구분한 ISO2 국가 코드")
    ap.add_argument("--per-country", type=int, default=20000)
    ap.add_argument("--since", default="", help="이 날짜(YYYY-MM-DD) 이후 처음 노출된 소재만")
    ap.add_argument("--project", default="")
    ap.add_argument("--max-gb", type=float, default=20.0, help="스캔량이 이보다 크면 --yes 없이는 멈춘다")
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--yes", action="store_true")
    ap.add_argument("-o", "--output", default="gads.jsonl")
    a = ap.parse_args()
    cols = schema(a.project)
    rows = []
    for c in [x.strip().upper() for x in a.countries.split(",") if x.strip()]:
        sql = build_sql(cols, c, a.since, a.per_country)
        dry = json.loads(bq(["query", "--use_legacy_sql=false", "--dry_run", sql], a.project) or "{}")
        gb = int(dry.get("statistics", {}).get("totalBytesProcessed", 0)) / 1e9
        print("%s: 예상 스캔 %.1f GB" % (c, gb), file=sys.stderr)
        if a.dry_run:
            print(sql, file=sys.stderr)
            continue
        if gb > a.max_gb and not a.yes:
            sys.exit("스캔량이 --max-gb(%.0f)를 넘습니다. 확인했다면 --yes 를 붙이세요." % a.max_gb)
        out = json.loads(bq(["query", "--use_legacy_sql=false", "--max_rows=" + str(a.per_country), sql], a.project) or "[]")
        rows += [to_record(r, c) for r in out]
        print("  %d건" % len(out), file=sys.stderr)
    if rows:
        print("%d건 저장: %s" % (write_jsonl(a.output, rows), a.output), file=sys.stderr)


if __name__ == "__main__":
    main()
