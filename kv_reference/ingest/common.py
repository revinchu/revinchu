"""수집 스크립트 공용: KV 아틀라스 레코드 정규화와 JSONL 입출력.

모든 수집기는 한 줄에 레코드 하나인 JSONL을 쓴다. build_bulk.py 가 이를 모아 아티팩트용 샤드로 만든다.
레코드 필드: id, title, brand, agency, year, date, country(ISO2, 다국가 XX), field, industry,
medium(list), techniques(list), moods(list), note, sourceUrl, origin, license
"""

import json
import re

# 국적·국가명 → ISO2 (데이터에 자주 나오는 것만)
COUNTRY_WORDS = {
    "american": "US", "united states": "US", "usa": "US", "british": "GB", "english": "GB", "united kingdom": "GB", "england": "GB",
    "french": "FR", "france": "FR", "german": "DE", "germany": "DE", "swiss": "CH", "switzerland": "CH", "dutch": "NL", "netherlands": "NL",
    "italian": "IT", "italy": "IT", "spanish": "ES", "spain": "ES", "polish": "PL", "poland": "PL", "russian": "RU", "russia": "RU",
    "japanese": "JP", "japan": "JP", "chinese": "CN", "china": "CN", "korean": "KR", "korea": "KR", "south korea": "KR",
    "swedish": "SE", "sweden": "SE", "danish": "DK", "denmark": "DK", "norwegian": "NO", "norway": "NO", "austrian": "AT", "austria": "AT",
    "belgian": "BE", "belgium": "BE", "canadian": "CA", "canada": "CA", "mexican": "MX", "mexico": "MX", "brazilian": "BR", "brazil": "BR",
    "argentine": "AR", "argentina": "AR", "australian": "AU", "australia": "AU", "portuguese": "PT", "portugal": "PT", "indian": "IN", "india": "IN",
}

# 영어 키워드 → 업종 (제목·설명에서 추정, 맞는 게 없으면 기타)
INDUSTRY_WORDS = [
    ("주류", r"\b(beer|wine|whisk(e)?y|brandy|gin|vodka|liquor|cognac|champagne|brewery|ale|sake)\b"),
    ("식음료", r"\b(coffee|tea|cocoa|chocolate|food|soup|biscuit|candy|soda|cola|milk|bread|restaurant|flour|sugar|baking)\b"),
    ("자동차·모빌리티", r"\b(automobile|motor|car|cars|tire|tyre|bicycle|cycle|airline|railway|railroad|train|shipping|steamship)\b"),
    ("여행·항공", r"\b(travel|tourism|tourist|hotel|resort|cruise|visit|holiday|airways)\b"),
    ("패션·뷰티", r"\b(fashion|clothing|shoes?|hat|corset|textile|dress|perfume|soap|cosmetic)\b"),
    ("제약·헬스케어", r"\b(medicine|pharmacy|remedy|pills?|tonic|health|hospital|cure)\b"),
    ("금융", r"\b(bank|insurance|savings|bonds?|loan)\b"),
    ("IT·전자", r"\b(radio|television|electric|computer|telephone|phonograph|camera|typewriter)\b"),
    ("미디어·엔터", r"\b(film|movie|cinema|theat(er|re)|circus|magazine|newspaper|music|concert|records?)\b"),
    ("문화·예술", r"\b(exhibition|museum|gallery|festival|opera|ballet|art)\b"),
    ("공공·캠페인", r"\b(war|election|vote|government|campaign|safety|army|navy|recruit|health service|olympic)\b"),
    ("리테일·생활", r"\b(store|shop|department|furniture|stove|sewing|household|cleaner)\b"),
]


def country_from(*texts):
    for t in texts:
        t = (t or "").lower().strip("() ")
        for part in re.split(r"[,;/|()]", t):
            part = part.strip()
            if part in COUNTRY_WORDS:
                return COUNTRY_WORDS[part]
    return "XX"


def industry_from(text):
    t = (text or "").lower()
    for name, pat in INDUSTRY_WORDS:
        if re.search(pat, t):
            return name
    return "기타"


def year_from(text):
    m = re.search(r"(1[5-9]\d\d|20\d\d)", text or "")
    return int(m.group(1)) if m else 0


def slug(s):
    return re.sub(r"[^a-z0-9가-힣]+", "-", str(s).lower()).strip("-")[:80] or "item"


def record(**kw):
    r = {
        "id": kw.get("id") or "", "title": (kw.get("title") or "").strip()[:200], "brand": (kw.get("brand") or "").strip()[:120],
        "agency": (kw.get("agency") or "").strip()[:160], "year": int(kw.get("year") or 0), "date": (kw.get("date") or "")[:10],
        "country": kw.get("country") or "XX", "field": kw.get("field") or "브랜드", "industry": kw.get("industry") or "기타",
        "medium": kw.get("medium") or ["print"], "techniques": kw.get("techniques") or [], "moods": kw.get("moods") or [],
        "note": (kw.get("note") or "").strip()[:300], "sourceUrl": kw.get("sourceUrl") or "", "origin": kw.get("origin") or "user",
        "license": kw.get("license") or "",
    }
    return r


def write_jsonl(path, rows):
    n = 0
    with open(path, "w", encoding="utf-8") as f:
        for r in rows:
            if r["title"]:
                f.write(json.dumps(r, ensure_ascii=False) + "\n")
                n += 1
    return n


def read_jsonl(path):
    with open(path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line:
                yield json.loads(line)
