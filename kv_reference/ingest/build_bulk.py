"""수집 결과(JSONL, JSON 배열)를 모아 KV 아틀라스가 나눠 불러오는 대량 데이터 샤드로 만든다.

    python kv_reference/ingest/build_bulk.py museums.jsonl gads.jsonl tvcf_print.json -o kv_reference/bulk

출력
- bulk/manifest.json   샤드 목록(출처, 국가, 건수, 연도 범위)과 전체 집계
- bulk/<출처>-<국가>-<n>.json   {"rows": [[id, title, brand, agency, year, date, country, field, industry, medium, techniques, moods, note, sourceUrl, imageUrl], ...]}
샤드 하나는 최대 --shard-size 건(기본 40000, 약 8MB)이다. 아티팩트 파일 한도(파일당 16MB, 버전당 256MB) 안에서 약 100만 건까지 담을 수 있다.
같은 id는 나중 입력이 앞 입력을 덮어쓴다.
"""

import argparse
import collections
import json
import os
import sys

from common import read_jsonl

COLS = ["id", "title", "brand", "agency", "year", "date", "country", "field", "industry", "medium", "techniques", "moods", "note", "sourceUrl", "imageUrl"]


def load(path):
    if path.endswith(".jsonl"):
        yield from read_jsonl(path)
    else:
        data = json.load(open(path, encoding="utf-8"))
        yield from (data if isinstance(data, list) else data.get("items", []))


def row(r):
    j = lambda v: "|".join(v) if isinstance(v, list) else (v or "")
    return [r.get("id", ""), r.get("title", ""), r.get("brand", ""), r.get("agency", ""), int(r.get("year") or 0), r.get("date", ""),
            r.get("country") or "XX", r.get("field", ""), r.get("industry", ""), j(r.get("medium")), j(r.get("techniques")), j(r.get("moods")),
            r.get("note", ""), r.get("sourceUrl", ""), r.get("imageUrl", "")]


def main():
    ap = argparse.ArgumentParser(description="수집 데이터 → KV 아틀라스 대량 샤드")
    ap.add_argument("inputs", nargs="+")
    ap.add_argument("-o", "--out", default="kv_reference/bulk")
    ap.add_argument("--shard-size", type=int, default=40000)
    a = ap.parse_args()
    items = {}
    for p in a.inputs:
        n = 0
        for r in load(p):
            if not r.get("title"):
                continue
            if not r.get("id"):
                r["id"] = (r.get("origin") or "x") + "-" + str(abs(hash((r.get("title"), r.get("sourceUrl")))))
            items[r["id"]] = r
            n += 1
        print("%s: %d건" % (p, n), file=sys.stderr)
    groups = collections.defaultdict(list)
    for r in items.values():
        groups[(r.get("origin") or "user", r.get("country") or "XX")].append(r)
    os.makedirs(a.out, exist_ok=True)
    for f in os.listdir(a.out):
        if f.endswith(".json"):
            os.remove(os.path.join(a.out, f))
    shards = []
    for (origin, country), rs in sorted(groups.items(), key=lambda kv: (kv[0][0], -len(kv[1]))):
        rs.sort(key=lambda r: (r.get("date") or str(r.get("year") or "")), reverse=True)
        for i in range(0, len(rs), a.shard_size):
            part = rs[i:i + a.shard_size]
            name = "%s-%s-%d.json" % (origin, country.lower(), i // a.shard_size)
            with open(os.path.join(a.out, name), "w", encoding="utf-8") as f:
                json.dump({"rows": [row(r) for r in part]}, f, ensure_ascii=False, separators=(",", ":"))
            years = [int(r.get("year") or 0) for r in part if r.get("year")]
            shards.append({"file": name, "origin": origin, "country": country, "count": len(part), "years": [min(years), max(years)] if years else [],
                           "bytes": os.path.getsize(os.path.join(a.out, name))})
    by_origin = collections.Counter()
    for s in shards:
        by_origin[s["origin"]] += s["count"]
    manifest = {"cols": COLS, "total": len(items), "byOrigin": dict(by_origin), "shards": shards}
    json.dump(manifest, open(os.path.join(a.out, "manifest.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("샤드 %d개, 총 %d건 → %s" % (len(shards), len(items), a.out), file=sys.stderr)


if __name__ == "__main__":
    main()
