"""뮤지엄 공개 데이터(CC0)에서 포스터·광고물 레코드를 뽑아 JSONL로 저장한다.

    python kv_reference/ingest/museums.py --moma Artworks.csv --met MetObjects.csv -o museums.jsonl

- MoMA: https://github.com/MuseumofModernArt/collection (Artworks.csv, CC0, 이미지 미포함)
- The Met: https://github.com/metmuseum/openaccess (MetObjects.csv, CC0 메타데이터)
메타데이터만 쓰고 이미지는 링크하지 않는다. 각 레코드의 sourceUrl 이 뮤지엄 작품 페이지다.
"""

import argparse
import csv
import re
import sys

from common import country_from, industry_from, record, write_jsonl, year_from

csv.field_size_limit(10 ** 8)

AD_WORDS = re.compile(r"poster|advertis|billboard|trade card|broadside|affiche|plakat", re.I)
MET_EXCLUDE = re.compile(r"jar label|mummy label|decanter label|bottle ticket|baseball card", re.I)


def moma(path):
    for x in csv.DictReader(open(path, encoding="utf-8-sig")):
        text = x["Title"] + " " + x["Medium"]
        is_design_print = x["Classification"] == "Design" and re.search(r"offset|lithograph|screenprint|letterpress", x["Medium"], re.I)
        if not (AD_WORDS.search(text) or is_design_print):
            continue
        med = ["poster"] if re.search(r"poster", text, re.I) or is_design_print else ["print"]
        yield record(
            id="moma-" + x["ObjectID"], title=x["Title"], agency=x["Artist"], year=year_from(x["Date"]), country=country_from(x["Nationality"]),
            field="문화·공연" if re.search(r"exhibition|museum|film|theat|concert|festival", x["Title"], re.I) else "브랜드",
            industry=industry_from(x["Title"]), medium=med, note=(x["Medium"] + (" · " + x["Date"] if x["Date"] else "")),
            sourceUrl=x["URL"] or "https://www.moma.org/collection/works/" + x["ObjectID"], origin="museum", license="MoMA CC0 metadata",
        )


def met(path):
    for x in csv.DictReader(open(path, encoding="utf-8-sig")):
        text = x["Object Name"] + " " + x["Title"] + " " + x["Classification"]
        if not AD_WORDS.search(text) or MET_EXCLUDE.search(text):
            continue
        med = ["poster"] if re.search(r"poster", text, re.I) else ["print"]
        yield record(
            id="met-" + x["Object ID"], title=x["Title"] or x["Object Name"], agency=x["Artist Display Name"], year=year_from(x["Object Date"] or x["Object Begin Date"]),
            country=country_from(x["Country"], x["Culture"], x["Artist Nationality"]), field="브랜드", industry=industry_from(x["Title"] + " " + x["Tags"]),
            medium=med, note=" · ".join(v for v in (x["Object Name"], x["Object Date"], x["Medium"][:80]) if v),
            sourceUrl=x["Link Resource"], origin="museum", license="The Met CC0 metadata",
        )


def main():
    ap = argparse.ArgumentParser(description="뮤지엄 CC0 데이터 → KV 아틀라스 JSONL")
    ap.add_argument("--moma")
    ap.add_argument("--met")
    ap.add_argument("-o", "--output", default="museums.jsonl")
    a = ap.parse_args()
    rows = []
    if a.moma:
        rows += list(moma(a.moma))
    if a.met:
        rows += list(met(a.met))
    n = write_jsonl(a.output, rows)
    print("%d건 저장: %s" % (n, a.output), file=sys.stderr)


if __name__ == "__main__":
    main()
