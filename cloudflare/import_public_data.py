"""공공데이터(data.go.kr)로 실제 관광지·병원·비급여 가격을 가져와 D1 적재용 SQL을 만든다.

    python cloudflare/import_public_data.py attractions   # → cloudflare/.import/import_attractions.sql
    python cloudflare/import_public_data.py hospitals     # → cloudflare/.import/import_hospitals.sql
    python cloudflare/import_public_data.py partners      # → cloudflare/.import/import_partners.sql (숙소·K푸드·쇼핑)
    npx wrangler d1 execute wizx-medical --remote --file cloudflare/.import/import_attractions.sql

API 키는 cloudflare/set_data_key.py 로 %USERPROFILE%\\.wizx\\data_go_kr_key 에 저장한 값을 쓴다 (저장소에 넣지 않음).
"""
import hashlib
import json
import os
import pathlib
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import date

HERE = pathlib.Path(__file__).resolve().parent
OUT = HERE / ".import"  # build/ 는 build.py 가 지우므로 따로 둔다
TODAY = date.today().isoformat()
KEY_FILE = pathlib.Path(os.path.expanduser("~")) / ".wizx" / "data_go_kr_key"
sys.path[:0] = [str(HERE), str(HERE.parent)]
import name_i18n  # noqa: E402  병원·명소·항목 외국어 표기
from medical_platform.i18n import district as district_name  # noqa: E402

# 한국관광공사 「외래관광객조사」에서 외국인 방문 비율이 높은 주요 방문지 (서울·부산).
# (TourAPI 국문 제목, 영문 제목 앞부분, 지역, 순위) — 세부 정보·좌표·사진은 TourAPI 에서 조회한다.
ATTRACTIONS = [
    ("명동", "Myeong-dong", "seoul", 1), ("경복궁", "Gyeongbokgung Palace", "seoul", 2), ("남산서울타워", "N Seoul Tower", "seoul", 3),
    ("동대문디자인플라자(DDP)", "Dongdaemun Design Plaza", "seoul", 4), ("북촌한옥마을", "Bukchon Hanok Village", "seoul", 5),
    ("인사동", "Insa-dong", "seoul", 6), ("롯데월드 어드벤처", "Lotte World", "seoul", 7), ("남대문시장", "Namdaemun Market", "seoul", 8),
    ("창덕궁과 후원 [유네스코 세계유산]", "Changdeokgung Palace Complex", "seoul", 9), ("청계천", "Cheonggyecheon", "seoul", 10),
    ("해운대해수욕장", "Haeundae Beach", "busan", 1), ("광안리해수욕장", "Gwangalli Beach", "busan", 2),
    ("부산 감천문화마을", "Busan Gamcheon Culture Village", "busan", 3), ("부산 자갈치시장", "Jagalchi Market", "busan", 4),
    ("해동용궁사", "Haedong Yonggungsa Temple", "busan", 5), ("BIFF 광장", "BIFF Square", "busan", 6),
    ("태종대", "Taejongdae Cliffed Coast", "busan", 7), ("흰여울문화마을", "Huinnyeoul Culture Village", "busan", 8),
]
REGION_CODE = {"seoul": "11", "busan": "26"}  # 법정동 시도 코드 (TourAPI lDongRegnCd)


def api_key():
    if os.environ.get("DATA_GO_KR_KEY"):
        return os.environ["DATA_GO_KR_KEY"].strip()
    if KEY_FILE.exists():
        return KEY_FILE.read_text(encoding="ascii").strip()
    sys.exit("API 키가 없습니다: python cloudflare/set_data_key.py")


def call(url, **params):
    params = {"serviceKey": api_key(), "_type": "json", **params}
    req = urllib.request.Request(url + "?" + urllib.parse.urlencode(params), headers={"User-Agent": "wizx-import"})
    for attempt in range(3):
        try:
            with urllib.request.urlopen(req, timeout=30) as r:
                data = json.loads(r.read().decode("utf-8"))
            body = data.get("response", {}).get("body", {})
            items = (body.get("items") or {})
            items = items.get("item", []) if isinstance(items, dict) else []
            return items if isinstance(items, list) else [items]
        except (OSError, ValueError):
            if attempt == 2:
                raise
            time.sleep(1.5)


def tour(service, op, **params):
    return call(f"https://apis.data.go.kr/B551011/{service}/{op}", MobileOS="ETC", MobileApp="WIZX", **params)


def sql_str(v):
    return "NULL" if v is None else "'" + str(v).replace("'", "''") + "'"


def maps_url(ko_name, ko_addr, city):
    """Google 지도 링크. 이름만 넣으면 비슷한 곳이 여러 개 뜨므로 국문 이름 + 주소로 한 곳에 바로 연결한다.
    (visitkorea 상세 페이지는 TourAPI 콘텐츠 ID 로 주소를 만들 수 없다)"""
    return "https://www.google.com/maps/search/?api=1&query=" + urllib.parse.quote(f"{ko_name} {ko_addr or city}".strip())


def clean(text):
    return re.sub(r"<br\s*/?>", " ", text or "").strip()[:600]


def match_ko(title, region):
    """같은 시도에서 제목이 정확히 일치하는 관광지(12)·문화시설(14)·쇼핑(38)."""
    norm = lambda t: t.replace(" ", "")
    for i in tour("KorService2", "searchKeyword2", keyword=title, numOfRows=50):
        if str(i.get("lDongRegnCd")) == REGION_CODE[region] and norm(i["title"]) == norm(title)                 and str(i["contenttypeid"]) in ("12", "14", "38"):
            return i
    return None


def match_en(keyword, ko_title):
    """영문 제목에는 국문 이름이 괄호로 붙는다: 'Gyeongbokgung Palace (경복궁)'."""
    items = [i for i in tour("EngService2", "searchKeyword2", keyword=keyword, numOfRows=30)
             if str(i.get("contenttypeid")) in ("76", "78", "79")]
    tagged = [i for i in items if f"({ko_title})" in i["title"]]
    if tagged:
        return tagged[0]
    starts = sorted((i for i in items if i["title"].lower().startswith(keyword.lower())), key=lambda i: len(i["title"]))
    return starts[0] if starts else None


def en_title(title):
    title = re.sub(r"^\s*(.+?)\s*\(\s*\1\s*\)\s*$", r"\1", title.strip())  # 'T1 (T1)' -> 'T1'
    return re.sub(r"\s*\((?=[^()]*[가-힣]).*$", "", title).strip()


def attractions():
    rows = []
    for keyword, keyword_en, region, rank in ATTRACTIONS:
        ko = match_ko(keyword, region)
        if not ko:
            print(f"  ! 찾지 못함: {keyword}")
            continue
        cid = ko["contentid"]
        detail_ko = (tour("KorService2", "detailCommon2", contentId=cid) or [{}])[0]
        en = match_en(keyword_en, ko["title"])
        detail_en = (tour("EngService2", "detailCommon2", contentId=en["contentid"]) or [{}])[0] if en else {}
        image, license_ = ko.get("firstimage") or "", ko.get("cpyrhtDivCd") or ""
        if not image and en:
            image, license_ = en.get("firstimage") or "", en.get("cpyrhtDivCd") or ""
        if license_ not in ("Type1", "Type3"):  # 공공누리 1·3유형(출처표시)만 사진 사용
            image = ""
        if not image:  # 대표 사진이 없으면 갤러리에서 1·3유형 사진을 찾는다 (예: 남산서울타워)
            gallery = [g for g in tour("KorService2", "detailImage2", contentId=cid, imageYN="Y", numOfRows=20)
                       if g.get("cpyrhtDivCd") in ("Type1", "Type3") and g.get("originimgurl")]
            if gallery:
                image, license_ = gallery[0]["originimgurl"], gallery[0]["cpyrhtDivCd"]
        image = image.replace("http://", "https://")  # https 페이지에서 http 이미지는 차단된다
        names, overviews, addrs = {"ko": ko["title"]}, {"ko": clean(detail_ko.get("overview"))}, {"ko": ko.get("addr1", "")}
        if en:
            names["en"] = en_title(en["title"])
            overviews["en"] = clean(detail_en.get("overview"))
            addrs["en"] = en.get("addr1", "")
        url = maps_url(ko["title"], ko.get("addr1", ""), "서울" if region == "seoul" else "부산")
        names = {**names, **name_i18n.names(name_i18n.SPOTS, ko["title"])}  # zh·ja·vi 와 정리된 영문 이름
        rows.append((cid, region, rank, names, overviews, addrs, ko.get("mapy"), ko.get("mapx"), image, license_, url))
        print(f"  {region} #{rank} {ko['title']} [{ko['contenttypeid']}]" + (f" / {names['en']}" if en else " (영문 없음)"))
        time.sleep(0.2)

    OUT.mkdir(exist_ok=True)
    lines = ["UPDATE attractions SET active = 0;"]
    for cid, region, rank, names, overviews, addrs, lat, lng, image, license_, url in rows:
        vals = ", ".join([
            sql_str(cid), sql_str(region), str(rank), sql_str(json.dumps(names, ensure_ascii=False)),
            sql_str(json.dumps(overviews, ensure_ascii=False)), sql_str(json.dumps(addrs, ensure_ascii=False)),
            str(float(lat)) if lat else "NULL", str(float(lng)) if lng else "NULL", sql_str(image), sql_str(license_),
            sql_str("한국관광공사 TourAPI"), sql_str(url), sql_str(TODAY), "1"])
        lines.append(
            "INSERT INTO attractions (content_id, region, rank, name_i18n, overview_i18n, address_i18n, lat, lng, image_url, "
            "image_license, source, source_url, verified_at, active) VALUES (" + vals + ") "
            "ON CONFLICT(content_id) DO UPDATE SET region=excluded.region, rank=excluded.rank, name_i18n=excluded.name_i18n, "
            "overview_i18n=excluded.overview_i18n, address_i18n=excluded.address_i18n, lat=excluded.lat, lng=excluded.lng, "
            "image_url=excluded.image_url, image_license=excluded.image_license, source_url=excluded.source_url, "
            "verified_at=excluded.verified_at, active=1;")
    path = OUT / "import_attractions.sql"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{len(rows)}곳 → {path}")


# ---------------------------------------------------------------------------
# 병원 + 비급여 가격 (건강보험심사평가원)
# 비급여 공개 제도는 미용 목적 시술(보톡스·필러·쌍꺼풀 등)을 포함하지 않는다.
# 그래서 실제 공개 가격이 있는 진료군만 가져온다: 치과·안과·대형병원 검사·한방.
# ---------------------------------------------------------------------------
HOSP_URL = "https://apis.data.go.kr/B551182/hospInfoServicev2/getHospBasisList"
NPAY_URL = "https://apis.data.go.kr/B551182/nonPaymentDamtInfoService/getNonPaymentItemHospDtlList"
SIDO = {"seoul": "110000", "busan": "210000"}
HOSPITAL_TYPE = {"01": "tertiary", "11": "general", "21": "hospital", "41": "dental_hospital", "92": "korean_medicine_hospital"}
TYPE_EN = {"tertiary": "Tertiary hospital", "general": "General hospital", "hospital": "Hospital",
           "dental_hospital": "Dental hospital", "korean_medicine_hospital": "Korean medicine hospital"}
# 진료군: 병원 조회 조건, 가져올 비급여 항목(npayKorNm 포함 문자열), 도시별 최대 병원 수, 가격대 비교 기준 항목
GROUPS = {
    "dental": dict(query=dict(clCd="41"), items=["치과임플란트", "보철료/크라운/", "치아교정", "치석제거/전악"], limit=10,
                   ref="치과임플란트"),
    "ophthalmology": dict(query=dict(clCd="21", dgsbjtCd="12"), items=["시력교정술료", "조절성 인공수정체"], limit=10,
                          ref="라식"),
    "health_checkup": dict(query=dict(clCd="01"), items=[], exact=True, limit=20, ref="뇌 MRI"),
    "korean_medicine": dict(query=dict(clCd="92"), items=["약침술", "추나요법"], limit=10, ref="약침술"),
}
# 대형병원 검진성 검사: 항목명이 길고 비슷한 것이 많아 정확한 이름으로만 고른다 → (표시 국문, 표시 영문)
CHECKUP_ITEMS = {
    "자기공명영상진단료(MRI-기본검사)/뇌/일반": ("뇌 MRI", "Brain MRI"),
    "자기공명영상진단료(MRI-기본검사)/혈관-뇌혈관/일반": ("뇌혈관 MRA", "Brain vessel MRA"),
    "초음파검사료(진단초음파)/두경부-경부 초음파/갑상선·부갑상선": ("갑상선 초음파", "Thyroid ultrasound"),
    "초음파검사료(진단초음파)/흉부-유방·액와부 초음파/일반": ("유방 초음파", "Breast ultrasound"),
    "초음파검사료(진단초음파)/복부-복부 초음파/간·담낭·담도·비장·췌장-일반": ("상복부 초음파 (간·담낭·췌장 등)",
                                                              "Upper abdominal ultrasound (liver, gallbladder, pancreas)"),
    "초음파검사료(진단초음파)/혈관-두개외 혈관 도플러 초음파/경동맥": ("경동맥 초음파", "Carotid artery ultrasound"),
}
ITEM_EN = [  # 국문 항목명 조각 → 영문 (긴 것부터 치환)
    ("올세라믹", "all-ceramic"), ("기타", "other"), ("특수(내장기, 두개천골)추나", "special (visceral, craniosacral)"),
    ("치과임플란트(1치당)", "Dental implant (per tooth)"), ("크라운", "Crown"), ("치석제거", "Scaling"), ("전악", "full mouth"),
    ("치아교정", "Orthodontics"), ("레이저각막상피절삭성형술(라섹)", "LASEK"), ("레이저각막절삭성형술(라식)", "LASIK"),
    ("시력교정술료", "Vision correction"), ("조절성 인공수정체", "Premium intraocular lens (multifocal / presbyopia)"),
    ("두부-뇌", "brain"), ("일반", "standard"), ("정밀", "detailed"), ("갑상선·부갑상선", "thyroid & parathyroid"),
    ("두경부-경부 초음파", "neck ultrasound"), ("유방·액와", "breast & axilla"), ("유방", "breast"), ("상복부", "upper abdomen"),
    ("복부-", "abdomen "), ("약침술", "Pharmacopuncture"), ("경혈", "acupoint"), ("추나요법", "Chuna therapy"),
    ("단순추나", "simple"), ("복잡추나", "complex"), ("특수(탈구)추나", "special (dislocation)"),
]


def item_names(npay_name):
    """'치과의 보철료/치과임플란트(1치당)/Zirconia' → ('치과임플란트(1치당) · Zirconia', 'Dental implant (per tooth) · Zirconia')."""
    parts = [p.strip() for p in npay_name.split("/")][1:] or [npay_name]
    if parts[0] == "조절성 인공수정체":  # 제품명이 수십 개라 하나로 묶는다
        parts = parts[:1]
    ko = " · ".join(parts)
    en_parts = []
    for part in parts:
        for k, v in sorted(ITEM_EN, key=lambda kv: -len(kv[0])):
            part = part.replace(k, v)
        en_parts.append(part)
    en = " · ".join(en_parts)
    return ko, (ko if re.search("[가-힣]", en) else en)


def hira_hospitals(region, group):
    rows, page = [], 1
    while True:
        items = call(HOSP_URL, sidoCd=SIDO[region], numOfRows=500, pageNo=page, **group["query"])
        rows += items
        if len(items) < 500:
            return sorted(rows, key=lambda h: -int(h.get("drTotCnt") or 0))
        page += 1


def hira_prices(ykiho, group):
    """병원의 공개 비급여 중 원하는 항목만: 항목별 최저~최고가, 건수, 기준일."""
    found = {}
    for i in call(NPAY_URL, ykiho=ykiho, numOfRows=1000, pageNo=1):
        name = i.get("npayKorNm") or ""
        exact = CHECKUP_ITEMS.get(re.sub(r"\s*/\s*", "/", name).strip()) if group.get("exact") else None
        if not i.get("curAmt") or not (exact or any(w in name for w in group["items"])):
            continue
        ko, en = exact or item_names(name)
        f = found.setdefault(ko, {"en": en, "min": None, "max": None, "n": 0, "since": None})
        amt = int(i["curAmt"])
        f["min"] = amt if f["min"] is None else min(f["min"], amt)
        f["max"] = amt if f["max"] is None else max(f["max"], amt)
        f["n"] += 1
        since = str(i.get("adtFrDd") or "")
        if len(since) == 8:
            since = f"{since[:4]}-{since[4:6]}-{since[6:]}"
            f["since"] = max(f["since"] or since, since)
    return found


def hospitals():
    import hashlib
    picked = []  # (region, specialty, hospital row, prices)
    for region in SIDO:
        for spec, group in GROUPS.items():
            n = 0
            for h in hira_hospitals(region, group):
                prices = hira_prices(h["ykiho"], group)
                if prices:
                    picked.append((region, spec, h, prices))
                    n += 1
                    print(f"  {region} {spec}: {h['yadmNm']} ({len(prices)}개 항목)")
                if n >= group["limit"]:
                    break
                time.sleep(0.1)
    # 같은 진료군 안에서 기준 항목 최저가로 가격대(₩/₩₩/₩₩₩) 산정
    bands = {}
    for spec, group in GROUPS.items():
        ref = sorted((min(v["min"] for k, v in pr.items() if group["ref"] in k), h["ykiho"])
                     for _, s2, h, pr in picked if s2 == spec and any(group["ref"] in k for k in pr))
        for idx, (_, yk) in enumerate(ref):
            bands[(spec, yk)] = 1 + (3 * idx) // max(len(ref), 1)

    merged = {}
    for region, spec, h, prices in picked:  # 한 병원이 여러 진료군에 걸치면 합친다
        m = merged.setdefault(h["ykiho"], {"region": region, "h": h, "specs": [], "prices": {}})
        m["specs"].append(spec)
        m["prices"].update(prices)
    lines = ["UPDATE hospitals SET verification_status = 'archived' WHERE is_demo = 1;",
             "UPDATE partners SET active = 0 WHERE is_demo = 1;"]
    for yk, m in merged.items():
        h, slug = m["h"], "hira-" + hashlib.sha1(yk.encode()).hexdigest()[:10]
        htype = HOSPITAL_TYPE.get(str(h.get("clCd")), "hospital")
        since = max((p["since"] for p in m["prices"].values() if p["since"]), default=TODAY)
        band = round(sum(bands.get((s, yk), 2) for s in m["specs"]) / len(m["specs"]))
        dr = int(h.get("drTotCnt") or 0)
        city = "Seoul" if m["region"] == "seoul" else "Busan"
        gu, n = h.get("sgguCdNm", ""), len(m["prices"])
        t_zh, t_ja, t_vi = name_i18n.HOSPITAL_TYPE[htype]
        summary = {
            "ko": f"{district_name(gu, 'ko')}에 있는 {h.get('clCdNm', '')}입니다. 의사 {dr}명. "
                  f"건강보험심사평가원에 공개한 비급여 진료비 {n}개 항목을 확인할 수 있습니다.",
            "en": f"{TYPE_EN[htype]} in {district_name(gu, 'en')}, {city}, with {dr} doctors. "
                  f"{n} self-pay prices published with HIRA (Health Insurance Review & Assessment Service).",
            "zh": f"位于{'首尔' if city == 'Seoul' else '釜山'}{district_name(gu, 'zh')}的{t_zh}，医生{dr}名。"
                  f"可查看向韩国健康保险审查评价院（HIRA）公开的{n}项自费价格。",
            "ja": f"{'ソウル' if city == 'Seoul' else '釜山'}{district_name(gu, 'ja')}にある{t_ja}です。医師{dr}名。"
                  f"健康保険審査評価院（HIRA）に公開された保険外診療費{n}項目を確認できます。",
            "vi": f"{t_vi} tại {district_name(gu, 'vi')}, {city}, với {dr} bác sĩ. "
                  f"Có {n} mức giá tự chi trả được công bố với HIRA (Cơ quan Thẩm định Bảo hiểm Y tế Hàn Quốc).",
        }
        hnames = name_i18n.names(name_i18n.HOSPITALS, h["yadmNm"])
        if "en" not in hnames:
            print(f"  ! 외국어 병원명 없음 (name_i18n.HOSPITALS 에 추가): {h['yadmNm']}")
        lat, lng = float(h.get("YPos") or 0), float(h.get("XPos") or 0)
        vals = ", ".join([
            sql_str(slug), sql_str(h["yadmNm"]), sql_str(hnames.get("en", h["yadmNm"])), sql_str(htype), sql_str(m["region"]),
            sql_str(h.get("sgguCdNm", "")), sql_str(h.get("addr", "")), sql_str(h.get("addr", "")),
            str(lat) if lat else "NULL", str(lng) if lng else "NULL", sql_str(",".join(m["specs"])),
            sql_str(h.get("telno", "")), sql_str(h.get("hospUrl", "")), str(band),
            sql_str(json.dumps(summary, ensure_ascii=False)), "'published'", sql_str(since), "0", sql_str(TODAY), sql_str(TODAY),
            sql_str(json.dumps(hnames, ensure_ascii=False))])
        lines.append(
            "INSERT INTO hospitals (slug, name_ko, name_en, hospital_type, region, district, address_ko, address_en, lat, lng, "
            "specialties, phone, website, price_band, summary_i18n, verification_status, last_verified_at, is_demo, created_at, "
            "updated_at, name_i18n) VALUES (" + vals + ") ON CONFLICT(slug) DO UPDATE SET name_ko=excluded.name_ko, "
            "name_en=excluded.name_en, name_i18n=excluded.name_i18n, "
            "hospital_type=excluded.hospital_type, district=excluded.district, address_ko=excluded.address_ko, "
            "address_en=excluded.address_en, lat=excluded.lat, lng=excluded.lng, specialties=excluded.specialties, "
            "phone=excluded.phone, website=excluded.website, price_band=excluded.price_band, summary_i18n=excluded.summary_i18n, "
            "verification_status='published', last_verified_at=excluded.last_verified_at, updated_at=excluded.updated_at;")
        hid = f"(SELECT id FROM hospitals WHERE slug = {sql_str(slug)})"
        lines.append(f"DELETE FROM price_items WHERE hospital_id = {hid};")
        for ko, p in sorted(m["prices"].items()):
            spread = f"₩{p['min']:,}–₩{p['max']:,} ({p['n']})" if p["max"] != p["min"] else ""
            lines.append(
                "INSERT INTO price_items (hospital_id, item_name_en, item_name_ko, price_domestic_krw, price_foreign_krw, "
                f"includes, source, verified_at, name_i18n) VALUES ({hid}, {sql_str(p['en'])}, {sql_str(ko)}, {p['min']}, NULL, "
                f"{sql_str(spread)}, {sql_str('HIRA 비급여진료비 공개')}, {sql_str(p['since'] or TODAY)}, "
                f"{sql_str(json.dumps({'ko': ko, 'en': p['en'], **name_i18n.item_names(ko)}, ensure_ascii=False))});")
    OUT.mkdir(exist_ok=True)
    path = OUT / "import_hospitals.sql"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"병원 {len(merged)}곳 → {path}")


# ---------------------------------------------------------------------------
# 숙소·K푸드·쇼핑 (한국관광공사 TourAPI) — 병원·관광 중심지 반경 3km
# 계약 제휴사가 아니라 공개 정보이므로 partners.source='TourAPI' 로 구분하고 화면에도 출처를 밝힌다.
# ---------------------------------------------------------------------------
HUBS = [  # (지역, 이름, 경도, 위도)
    ("seoul", "명동·중구", 126.9850, 37.5636), ("seoul", "강남", 127.0276, 37.4979), ("seoul", "서초", 127.0080, 37.5013),
    ("seoul", "신촌", 126.9368, 37.5551), ("seoul", "종로·혜화", 126.9990, 37.5796), ("seoul", "잠실", 127.1000, 37.5133),
    ("busan", "해운대", 129.1604, 35.1587), ("busan", "서면", 129.0592, 35.1578), ("busan", "남포", 129.0327, 35.0980),
]
# 유형: (영문 서비스 contentTypeId, 국문 contentTypeId, 허용 분류(lclsSystm2), 거점당 개수)
PLACE_TYPES = {
    "hotel": ("80", "32", ("AC01",), 3),                       # 호텔
    "kfood": ("82", "39", ("FD01",), 3),                       # 한식
    "shopping": ("79", "38", ("SH01", "SH02", "SH05", "SH06"), 3),  # 백화점·면세점·특산/전문·시장
}
SKIP_SHOPS = ("GS25", "CU ", "세븐일레븐", "이마트24", "편의점", "7-Eleven", "emart24")


def ko_from_en(title):
    m = re.search(r"\s?\((.*[가-힣].*)\)\s*$", title)  # 'HOTEL QB (호텔큐비(호텔QB))' 처럼 괄호가 겹칠 수 있다
    return m.group(1).strip() if m else None


def partners():
    # 이미 '명소'로 소개한 곳(예: 자갈치시장)은 쇼핑 목록에서 뺀다
    seen, rows = {t.replace(" ", "") for t, *_ in ATTRACTIONS}, []
    for region, hub, x, y in HUBS:
        for ptype, (en_ct, ko_ct, cats, quota) in PLACE_TYPES.items():
            # 외국인 대상 영문 서비스 목록을 먼저, 모자라면 국문 서비스에서 사진 있는 곳으로 채운다
            cands = [("en", i) for i in tour("EngService2", "locationBasedList2", mapX=x, mapY=y, radius=3000,
                                              contentTypeId=en_ct, numOfRows=100, arrange="E")]
            cands += [("ko", i) for i in tour("KorService2", "locationBasedList2", mapX=x, mapY=y, radius=3000,
                                               contentTypeId=ko_ct, numOfRows=100, arrange="E")]
            taken = 0
            for lang, i in cands:
                if taken >= quota:
                    break
                title = i.get("title", "")
                ko_name = ko_from_en(title) if lang == "en" else title
                key = (ko_name or title).replace(" ", "")
                lat, lng = float(i.get("mapy") or 0), float(i.get("mapx") or 0)
                # 영문·국문 서비스에 같은 곳이 이름만 다르게 있을 수 있어 같은 유형 80m 이내는 하나로 본다
                near_dup = any(r["ptype"] == ptype and abs(float(r["lat"] or 0) - lat) < 0.0007 and abs(float(r["lng"] or 0) - lng) < 0.0009
                               for r in rows)
                if key in seen or near_dup or i.get("lclsSystm2") not in cats or not i.get("firstimage") \
                        or i.get("cpyrhtDivCd") not in ("Type1", "Type3") or any(s in title for s in SKIP_SHOPS):
                    continue
                detail = (tour("EngService2" if lang == "en" else "KorService2", "detailCommon2", contentId=i["contentid"]) or [{}])[0]
                names = {"ko": ko_name or en_title(title)}
                if lang == "en":
                    names["en"] = en_title(title)
                desc = {lang: clean(detail.get("overview"))[:300]}
                url = maps_url(names["ko"], i.get("addr1", ""), "서울" if region == "seoul" else "부산")
                rows.append(dict(cid=f"{lang}-{i['contentid']}", ptype=ptype, region=region, names=names, desc=desc,
                                 addr=i.get("addr1", ""), lat=i.get("mapy"), lng=i.get("mapx"),
                                 image=i["firstimage"].replace("http://", "https://"), url=url))
                seen.add(key)
                taken += 1
                print(f"  {hub} {ptype}: {names['ko']}" + (f" / {names['en']}" if 'en' in names else ""))
                time.sleep(0.1)

    lines = ["UPDATE partners SET active = 0 WHERE source = 'TourAPI';"]
    for r in rows:
        cid = sql_str(r["cid"])
        cols = {
            "partner_type": sql_str(r["ptype"]), "name": sql_str(r["names"]["ko"]), "region": sql_str(r["region"]),
            "lat": str(float(r["lat"])) if r["lat"] else "NULL", "lng": str(float(r["lng"])) if r["lng"] else "NULL",
            "description_i18n": sql_str(json.dumps(r["desc"], ensure_ascii=False)), "price_band": "0",
            "name_i18n": sql_str(json.dumps(r["names"], ensure_ascii=False)), "address": sql_str(r["addr"]),
            "image_url": sql_str(r["image"]), "source": "'TourAPI'", "source_url": sql_str(r["url"]),
            "verified_at": sql_str(TODAY), "active": "1", "is_demo": "0",
        }
        lines.append("UPDATE partners SET " + ", ".join(f"{k} = {v}" for k, v in cols.items()) + f" WHERE content_id = {cid};")
        lines.append(f"INSERT INTO partners (content_id, {', '.join(cols)}) SELECT {cid}, {', '.join(cols.values())} "
                     f"WHERE NOT EXISTS (SELECT 1 FROM partners WHERE content_id = {cid});")
    OUT.mkdir(exist_ok=True)
    path = OUT / "import_partners.sql"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"{len(rows)}곳 → {path}")


# ---------------------------------------------------------------------------
# 진료과 채우기: 외국인환자 유치기관 등록 의원 + 상급종합병원 진료과
# - 한국보건산업진흥원 「외국인환자 유치기관 현황」(.import/raw/khidi_foreign_patient_institutions.csv) 에 있고
# - 심평원에 해당 진료과로 등록된 의원 중, 의원 이름에 전문과목이 들어간 곳(= 그 과 전문의가 운영)만 고른다.
# - 상급종합병원은 심평원 진료과 목록으로 피부과·성형외과·정형외과·산부인과·흉부외과·방사선종양학과를 확인해 태그한다.
# 의원은 심평원 비급여 공개 대상이 아니라 가격이 없다 (미용 시술도 공개 대상 아님) → 화면에서 견적 문의로 안내.
# ---------------------------------------------------------------------------
KHIDI_CSV = HERE / ".import" / "raw" / "khidi_foreign_patient_institutions.csv"
CLINIC_GROUPS = {  # 진료과: (심평원 진료과목 코드, 의원 이름에 있어야 할 단어, 서울 수, 부산 수)
    "dermatology": ("14", ("피부과",), 12, 6),
    "plastic_surgery": ("08", ("성형외과",), 12, 6),
    "ophthalmology": ("12", ("안과",), 8, 4),
    "orthopedics": ("05", ("정형외과",), 8, 4),
    "fertility": ("10", ("산부인과", "여성의원", "여성병원"), 8, 4),
}
# 상급종합병원 진료과 태그: 사이트 진료과 → 심평원 진료과목 코드
TERTIARY_DEPTS = {"dermatology": "14", "plastic_surgery": "08", "orthopedics": "05", "fertility": "10",
                  "cardiology": "07", "oncology": "17"}


def _norm(name):
    return re.sub(r"[\s()（）·.,\-]|의료법인|재단법인|학교법인|사회복지법인", "", name or "")


def khidi_names():
    import csv
    with open(KHIDI_CSV, encoding="cp949", errors="replace") as f:
        return {_norm(r["상호명"]) for r in csv.DictReader(f) if r.get("상태") == "유지" and r.get("지역") in ("서울", "부산")}


def hira_all(region, **query):
    rows, page = [], 1
    while True:
        items = call(HOSP_URL, sidoCd=SIDO[region], numOfRows=1000, pageNo=page, **query)
        rows += items
        if len(items) < 1000:
            return rows
        page += 1


def registered(name, khidi):
    n = _norm(name)
    return n in khidi or any(k and (k in n or n in k) and min(len(k), len(n)) >= 4 for k in khidi)


def select_clinics():
    khidi = khidi_names()
    picked = []
    for spec, (code, words, n_seoul, n_busan) in CLINIC_GROUPS.items():
        for region, limit in (("seoul", n_seoul), ("busan", n_busan)):
            rows = [h for h in hira_all(region, clCd="31", dgsbjtCd=code)
                    if any(w in h["yadmNm"] for w in words) and _norm(h["yadmNm"]) in khidi]
            rows.sort(key=lambda h: (-int(h.get("drTotCnt") or 0), h["yadmNm"]))
            for h in rows[:limit]:
                picked.append((spec, region, h))
            print(f"  {region} {spec}: 등록·전문의 의원 {len(rows)}곳 중 {min(limit, len(rows))}곳")
    return picked


def clinics():
    sys.path.insert(0, str(HERE))
    khidi = khidi_names()
    picked = select_clinics()
    lines = []
    for spec, region, h in picked:
        slug = "hira-" + hashlib.sha1(h["ykiho"].encode()).hexdigest()[:10]
        names = name_i18n.clinic_names(h["yadmNm"], spec)
        if "en" not in names:
            print(f"  ! 영문 브랜드 없음 (name_i18n.CLINIC_BRANDS 에 추가): {h['yadmNm']}")
        gu, dr, city = h.get("sgguCdNm", ""), int(h.get("drTotCnt") or 0), "Seoul" if region == "seoul" else "Busan"
        t = name_i18n.CLINIC_TYPE[spec]
        summary = {
            "ko": f"{district_name(gu, 'ko')}에 있는 {t[0]}입니다. 의사 {dr}명. 한국보건산업진흥원에 등록된 외국인환자 유치기관입니다. "
                  "의원급·미용 시술은 비급여 가격 공개 대상이 아니어서 가격은 상담으로 확인해 드립니다.",
            "en": f"{t[1]} in {district_name(gu, 'en')}, {city}, with {dr} doctor(s), registered with KHIDI to treat international "
                  "patients. Clinic and cosmetic prices are not part of Korea's public price disclosure — ask us for a quote.",
            "zh": f"位于{'首尔' if region == 'seoul' else '釜山'}{district_name(gu, 'zh')}的{t[2]}，医生{dr}名，"
                  "已在韩国保健产业振兴院登记为外国患者招揽机构。诊所及美容项目不在公开价格范围内，请咨询报价。",
            "ja": f"{'ソウル' if region == 'seoul' else '釜山'}{district_name(gu, 'ja')}にある{t[3]}です。医師{dr}名。"
                  "韓国保健産業振興院に外国人患者誘致機関として登録されています。クリニック・美容施術は価格公開の対象外のため、お見積りはご相談ください。",
            "vi": f"{t[4]} tại {district_name(gu, 'vi')}, {city}, {dr} bác sĩ, đã đăng ký với KHIDI để tiếp nhận bệnh nhân quốc tế. "
                  "Giá phòng khám và thẩm mỹ không thuộc dữ liệu công bố — hãy hỏi chúng tôi để báo giá.",
        }
        lat, lng = float(h.get("YPos") or 0), float(h.get("XPos") or 0)
        vals = ", ".join([
            sql_str(slug), sql_str(h["yadmNm"]), sql_str(names.get("en", h["yadmNm"])), "'clinic'", sql_str(region),
            sql_str(gu), sql_str(h.get("addr", "")), sql_str(h.get("addr", "")), str(lat) if lat else "NULL",
            str(lng) if lng else "NULL", sql_str(spec), sql_str(h.get("telno", "")), sql_str(h.get("hospUrl", "")), "0",
            sql_str(json.dumps(summary, ensure_ascii=False)), "'published'", sql_str(TODAY), "0", sql_str(TODAY), sql_str(TODAY),
            sql_str(json.dumps(names, ensure_ascii=False)), "1"])
        lines.append(
            "INSERT INTO hospitals (slug, name_ko, name_en, hospital_type, region, district, address_ko, address_en, lat, lng, "
            "specialties, phone, website, price_band, summary_i18n, verification_status, last_verified_at, is_demo, created_at, "
            "updated_at, name_i18n, foreign_patient_registered) VALUES (" + vals + ") ON CONFLICT(slug) DO UPDATE SET "
            "name_ko=excluded.name_ko, name_en=excluded.name_en, name_i18n=excluded.name_i18n, district=excluded.district, "
            "address_ko=excluded.address_ko, lat=excluded.lat, lng=excluded.lng, specialties=excluded.specialties, "
            "phone=excluded.phone, website=excluded.website, summary_i18n=excluded.summary_i18n, "
            "foreign_patient_registered=1, verification_status='published', updated_at=excluded.updated_at;")
    # 상급종합병원 진료과 태그 (심평원 진료과목 목록으로 확인한 것만)
    for spec, code in TERTIARY_DEPTS.items():
        for region in SIDO:
            for h in hira_all(region, clCd="01", dgsbjtCd=code):
                slug = "hira-" + hashlib.sha1(h["ykiho"].encode()).hexdigest()[:10]
                lines.append(f"UPDATE hospitals SET specialties = specialties || ',{spec}' WHERE slug = {sql_str(slug)} "
                             f"AND ',' || specialties || ',' NOT LIKE '%,{spec},%';")
    # 이미 들어간 병원 중 외국인환자 유치기관 등록 여부 표시
    for row in re.findall(r"INSERT INTO hospitals .*?VALUES \('(hira-[0-9a-f]+)', '((?:[^']|'')*)'",
                          (OUT / "import_hospitals.sql").read_text(encoding="utf-8")):
        if registered(row[1].replace("''", "'"), khidi):
            lines.append(f"UPDATE hospitals SET foreign_patient_registered = 1 WHERE slug = {sql_str(row[0])};")
    path = OUT / "import_clinics.sql"
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"의원 {len(picked)}곳 + 진료과 태그 → {path}")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    what = sys.argv[1] if len(sys.argv) > 1 else ""
    if what == "attractions":
        attractions()
    elif what == "hospitals":
        hospitals()
    elif what == "partners":
        partners()
    elif what == "select-clinics":
        for spec, region, h in select_clinics():
            print(f"{spec}\t{region}\t{h['yadmNm']}\t{h.get('drTotCnt')}\t{h.get('sgguCdNm')}")
    elif what == "clinics":
        clinics()
    else:
        sys.exit(__doc__)
