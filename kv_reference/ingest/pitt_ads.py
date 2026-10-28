"""Pitt Image Ads Dataset(Hussain et al., CVPR 2017) 주석을 JSONL로 변환한다. 이미지는 쓰지 않는다.

    python kv_reference/ingest/pitt_ads.py --dir annotations_images/image -o pitt.jsonl

주석 파일: https://people.cs.pitt.edu/~kovashka/ads/ (annotations_images.zip). 이미지는 저자에게 요청해야 하며 연구 목적 사용이 조건이다.
읽는 파일: Topics.json, Sentiments.json, QA_Combined_Action_Reason_train.json(또는 QA_Action.json·QA_Reason.json), Slogans.json(있으면).
토픽·감정 번호의 이름은 같은 폴더나 상위 폴더의 readme 에서 읽고, 없으면 번호로 둔다.
"""

import argparse
import collections
import glob
import json
import os
import re
import sys

from common import record, write_jsonl

TOPIC_INDUSTRY = [
    (r"restaurant|fast food|cafe", "식음료"), (r"chocolate|cookie|candy|ice cream|chips|snack|cereal|soup|seasoning|condiment|soda|juice|water|milk|coffee|tea", "식음료"),
    (r"alcohol|beer|wine", "주류"), (r"car|automobile", "자동차·모빌리티"), (r"electronic", "IT·전자"), (r"phone|tv|internet service", "통신"),
    (r"financial|bank|insurance", "금융"), (r"education", "교육"), (r"software", "IT·전자"), (r"beauty|cosmetic", "화장품"),
    (r"health|medication", "제약·헬스케어"), (r"clothing|accessor", "패션·뷰티"), (r"game|toy", "게임"), (r"cleaning|home appliance|home improvement|baby", "리테일·생활"),
    (r"vacation|travel", "여행·항공"), (r"media|arts", "미디어·엔터"), (r"sport", "스포츠"), (r"shopping", "유통·이커머스"),
    (r"environment|animal|human rights|safety|smoking|domestic violence|self esteem|charit|political", "공공·캠페인"),
]
PSA = re.compile(r"environment|animal right|human right|safety|smoking|domestic violence|self esteem|charit|political", re.I)


def load(path):
    return json.load(open(path, encoding="utf-8")) if os.path.exists(path) else {}


def names_from_readme(folder):
    names = {"topic": {}, "sentiment": {}}
    for p in glob.glob(os.path.join(folder, "**", "*readme*"), recursive=True) + glob.glob(os.path.join(folder, "..", "*readme*")):
        text = open(p, encoding="utf-8", errors="ignore").read()
        sec = None
        for line in text.splitlines():
            low = line.lower()
            if "topic" in low and len(line) < 60:
                sec = "topic"
            elif "sentiment" in low and len(line) < 60:
                sec = "sentiment"
            m = re.match(r"\s*(\d{1,2})\s*[\.:\)\-]?\s*\"?([A-Za-z][^\"]{1,80})", line)
            if sec and m:
                names[sec].setdefault(m.group(1), m.group(2).strip().strip('",'))
    return names


def majority(values):
    vals = [str(v[0] if isinstance(v, list) else v) for v in (values or [])]
    return collections.Counter(vals).most_common(1)[0][0] if vals else ""


def main():
    ap = argparse.ArgumentParser(description="Pitt Image Ads 주석 → KV 아틀라스 JSONL")
    ap.add_argument("--dir", required=True, help="Topics.json 등이 있는 폴더")
    ap.add_argument("-o", "--output", default="pitt.jsonl")
    a = ap.parse_args()
    topics = load(os.path.join(a.dir, "Topics.json"))
    sents = load(os.path.join(a.dir, "Sentiments.json"))
    qa = load(os.path.join(a.dir, "QA_Combined_Action_Reason_train.json")) or load(os.path.join(a.dir, "QA_Action.json"))
    slogans = load(os.path.join(a.dir, "Slogans.json"))
    names = names_from_readme(a.dir)
    keys = set(topics) | set(qa)
    rows = []
    for k in sorted(keys):
        t = majority(topics.get(k))
        tname = names["topic"].get(t, "토픽 " + t if t else "")
        ind = next((i for pat, i in TOPIC_INDUSTRY if re.search(pat, tname, re.I)), "기타")
        ar = qa.get(k) or []
        ar_text = ar[0] if ar and isinstance(ar[0], str) else ""
        slogan = (slogans.get(k) or [""])[0] if slogans else ""
        s = majority(sents.get(k))
        rows.append(record(
            id="pitt-" + re.sub(r"[^0-9a-z]+", "-", k.lower()).strip("-"), title=(slogan or tname or "광고 이미지") + " · " + k,
            year=0, country="XX", field="공익·사회" if PSA.search(tname) else "제품·서비스", industry=ind, medium=["print"],
            note=" / ".join(x for x in (ar_text, ("감정: " + names["sentiment"].get(s, s)) if s else "") if x),
            sourceUrl="https://people.cs.pitt.edu/~kovashka/ads/", origin="pitt", license="Pitt Image Ads Dataset (research use)",
        ))
    print("%d건 저장: %s" % (write_jsonl(a.output, rows), a.output), file=sys.stderr)


if __name__ == "__main__":
    main()
