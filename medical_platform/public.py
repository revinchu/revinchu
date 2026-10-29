"""해외 환자용 다국어 공개 사이트: 홈, 병원 지도·검색, 병원 상세, 컨시어지, 상담 신청."""
import json
import math
import re

from flask import (Blueprint, abort, current_app, g, redirect, render_template, request, session, url_for)

from . import crm, i18n, notify
from . import magazine as mag
from .db import get_db, now_iso

bp = Blueprint("public", __name__)

UTM_KEYS = ("utm_source", "utm_medium", "utm_campaign", "utm_content", "utm_term")


@bp.url_defaults
def add_lang(endpoint, values):
    if "lang" not in values and getattr(g, "lang", None) and current_app.url_map.is_endpoint_expecting(endpoint, "lang"):
        values["lang"] = g.lang


@bp.url_value_preprocessor
def pull_lang(endpoint, values):
    if values and "lang" in values:
        lang = values.pop("lang")
        if lang not in i18n.LANGUAGES:
            abort(404)
        g.lang = lang


@bp.before_request
def capture_attribution():
    """최초 유입(first-touch) 광고 파라미터를 세션에 저장해 상담 신청 시 리드에 연결한다."""
    if request.method != "GET":
        return
    current = session.get("attribution")
    has_utm = any(k in request.args for k in UTM_KEYS)
    if current is None or (has_utm and not current.get("utm_source")):
        session["attribution"] = {
            **{k: request.args.get(k, "")[:120] for k in UTM_KEYS},
            "landing_page": request.full_path[:300],
            "referrer": (request.referrer or "")[:300],
        }


@bp.app_context_processor
def inject_helpers():
    lang = getattr(g, "lang", i18n.DEFAULT_LANG)
    return {
        "lang": lang,
        "t": lambda key: i18n.translate(key, lang),
        "i18n": i18n,
        "site_url": notify.site_url,
        "magazine_categories": mag.CATEGORIES,
    }


def track(event, hospital_id=None):
    attr = session.get("attribution", {})
    db = get_db()
    db.execute(
        "INSERT INTO page_events (event, path, hospital_id, lang, utm_source, utm_campaign, created_at) VALUES (?,?,?,?,?,?,?)",
        (event, request.path[:200], hospital_id, g.lang, attr.get("utm_source", ""), attr.get("utm_campaign", ""), now_iso()),
    )
    db.commit()


def localized(json_text, lang):
    data = json.loads(json_text or "{}")
    return data.get(lang) or data.get("en") or ""


def partner_view(row, lang):
    """숙소·식당·쇼핑. 공개 정보(TourAPI)는 언어별 이름이 있고, 없으면 국문 이름을 쓴다."""
    names = json.loads(row["name_i18n"] or "{}")
    return dict(row, name=names.get(lang) or names.get("en") or row["name"],
                description=localized(row["description_i18n"], lang))


def distance_km(lat1, lng1, lat2, lng2):
    if None in (lat1, lng1, lat2, lng2):
        return 9999
    dlat, dlng = math.radians(lat2 - lat1), math.radians(lng2 - lng1)
    a = math.sin(dlat / 2) ** 2 + math.cos(math.radians(lat1)) * math.cos(math.radians(lat2)) * math.sin(dlng / 2) ** 2
    return 6371 * 2 * math.asin(math.sqrt(a))


def load_attractions(lang, region=None, limit=8):
    """외국인이 많이 찾는 관광지 (TourAPI 적재분). 언어별 이름·소개가 없으면 영어로 대체."""
    sql = "SELECT * FROM attractions WHERE active = 1"
    params = []
    if region:
        sql += " AND region = ?"
        params.append(region)
    sql += " ORDER BY rank, id LIMIT ?"
    params.append(limit)
    def text(field):
        data = json.loads(field or "{}")
        return data.get(lang) or data.get("en") or data.get("ko") or ""
    return [dict(r, name=text(r["name_i18n"]), overview=text(r["overview_i18n"]), address=text(r["address_i18n"]))
            for r in get_db().execute(sql, params).fetchall()]


@bp.app_template_global("hospital_name")
def hospital_name(row, lang):
    key = "hospital_name_i18n" if "hospital_name_i18n" in row.keys() else "name_i18n"  # 가격판 행은 병원명이 별칭으로 온다
    names = json.loads(row[key] or "{}") if key in row.keys() else {}
    return names.get(lang) or (row["name_ko"] if lang == "ko" else names.get("en") or row["name_en"])


@bp.app_template_global("item_name")
def item_name(p, lang):
    """가격 항목명: 언어별 표기 → 없으면 국문/영문."""
    names = json.loads(p["name_i18n"] or "{}") if "name_i18n" in p.keys() else {}
    return names.get(lang) or (p["item_name_ko"] if lang == "ko" else p["item_name_en"])


def hospital_view(row, lang, tags=None):
    h = dict(row)
    h["name"] = hospital_name(row, lang)
    h["district_name"] = i18n.district(h["district"], lang)
    h["alt_name"] = row["name_ko"] if lang != "ko" else hospital_name(row, "en")  # 현지에서 보여 줄 국문명 / 영문명
    h["address"] = h["address_ko"] if lang == "ko" else h["address_en"]
    h["summary"] = localized(h["summary_i18n"], lang)
    h["specialty_list"] = [s for s in h["specialties"].split(",") if s]
    h["language_list"] = [s for s in h["languages"].split(",") if s]
    h["tags"] = tags or []
    return h


def load_tags(hospital_ids):
    if not hospital_ids:
        return {}
    rows = []
    for i in range(0, len(hospital_ids), 90):  # Cloudflare D1 은 쿼리당 바인딩 파라미터 100개까지
        chunk = hospital_ids[i:i + 90]
        rows += get_db().execute(
            f"SELECT * FROM hospital_tags WHERE hospital_id IN ({','.join('?' * len(chunk))})", chunk).fetchall()
    out = {}
    for r in rows:
        out.setdefault(r["hospital_id"], []).append(dict(r))
    return out


def search_hospitals(args, lang):
    sql = "SELECT DISTINCT h.* FROM hospitals h LEFT JOIN hospital_tags t ON t.hospital_id = h.id WHERE h.verification_status = 'published'"
    params = []
    q = args.get("q", "").strip()
    if q:
        like = f"%{q}%"
        conds = ["h.name_en LIKE ?", "h.name_ko LIKE ?", "h.district LIKE ?", "h.address_ko LIKE ?",
                 "h.summary_i18n LIKE ?"]
        params += [like] * len(conds)
        # 진료과 이름(다국어)으로 검색한 경우 진료과 코드로도 매칭
        for code in i18n.SPECIALTIES:
            if any(q.lower() in n.lower() for n in i18n.T[f"sp_{code}"].values()):
                conds.append("h.specialties LIKE ?")
                params.append(f"%{code}%")
        sql += " AND (" + " OR ".join(conds) + ")"
    if args.get("specialty") in i18n.SPECIALTIES:
        sql += " AND (',' || h.specialties || ',') LIKE ?"
        params.append(f"%,{args['specialty']},%")
    if args.get("region") in i18n.REGIONS:
        sql += " AND h.region = ?"
        params.append(args["region"])
    if args.get("tag") in i18n.TAGS:
        sql += " AND t.tag = ?"
        params.append(args["tag"])
    if args.get("language") in i18n.LANGUAGES or args.get("language") in ("ru", "ar"):
        sql += " AND (',' || h.languages || ',') LIKE ?"
        params.append(f"%,{args['language']},%")
    if args.get("price") in ("1", "2", "3"):
        sql += " AND h.price_band = ?"
        params.append(int(args["price"]))
    sql += " ORDER BY h.name_en"
    rows = get_db().execute(sql, params).fetchall()
    tags = load_tags([r["id"] for r in rows])
    return [hospital_view(r, lang, tags.get(r["id"])) for r in rows]


def map_points(hospitals):
    # 병원이 많으면 url_for 반복 호출만으로도 무료 플랜 CPU 한도에 걸려, 주소는 문자열로 만든다
    base = url_for("public.hospitals")
    return [
        {"name": h["name"], "lat": h["lat"], "lng": h["lng"], "url": f"{base}/{h['slug']}", "district": h["district"]}
        for h in hospitals if h["lat"] is not None
    ]


PAGE_SIZE = 20


@bp.route("/")
def root():
    best = request.accept_languages.best_match(list(i18n.LANGUAGES)) or i18n.DEFAULT_LANG
    return redirect(url_for("public.home", lang=best, **request.args))


@bp.route("/<lang>/")
def home():
    hospitals = search_hospitals({}, g.lang)
    db = get_db()
    # 많은 병원이 공개한 항목부터, 항목별 최저가 병원 1곳씩
    price_rows = db.execute(
        """SELECT * FROM (
             SELECT p.*, h.slug, h.name_en, h.name_ko, h.name_i18n AS hospital_name_i18n, h.is_demo,
                    ROW_NUMBER() OVER (PARTITION BY p.item_name_ko ORDER BY p.price_domestic_krw, p.id) AS rn,
                    COUNT(*) OVER (PARTITION BY p.item_name_ko) AS hospital_count
             FROM price_items p JOIN hospitals h ON h.id = p.hospital_id
             WHERE h.verification_status = 'published' AND p.price_domestic_krw IS NOT NULL)
           WHERE rn = 1 ORDER BY hospital_count DESC, verified_at DESC LIMIT 5""").fetchall()
    tag_counts = {r["tag"]: r["n"] for r in db.execute(
        """SELECT t.tag, COUNT(DISTINCT t.hospital_id) AS n FROM hospital_tags t
           JOIN hospitals h ON h.id = t.hospital_id WHERE h.verification_status = 'published' GROUP BY t.tag""")}
    specialty_counts = {code: sum(code in h["specialty_list"] for h in hospitals) for code in i18n.SPECIALTIES}
    track("view_home")
    return render_template("public/home.html", hospitals=hospitals, points=map_points(hospitals),
                           price_rows=price_rows, tag_counts=tag_counts, specialty_counts=specialty_counts,
                           spots=load_attractions(g.lang, limit=8), articles=mag.listing(g.lang)[:4])


@bp.route("/<lang>/hospitals")
def hospitals():
    results = search_hospitals(request.args, g.lang)
    pages = max(1, -(-len(results) // PAGE_SIZE))
    page = min(max(request.args.get("page", 1, type=int), 1), pages)
    track("search")
    # 목록은 20곳씩 나눠 그리고(Workers CPU 절약), 지도에는 검색 결과 전체를 표시한다
    return render_template("public/hospitals.html", hospitals=results[(page - 1) * PAGE_SIZE:page * PAGE_SIZE],
                           total=len(results), page=page, pages=pages, points=map_points(results), args=request.args)


@bp.route("/<lang>/hospitals/<slug>")
def hospital_detail(slug):
    db = get_db()
    row = db.execute("SELECT * FROM hospitals WHERE slug = ? AND verification_status = 'published'", (slug,)).fetchone()
    if row is None:
        abort(404)
    h = hospital_view(row, g.lang, load_tags([row["id"]]).get(row["id"]))
    doctors = db.execute("SELECT * FROM doctors WHERE hospital_id = ? ORDER BY id", (row["id"],)).fetchall()
    prices = db.execute("SELECT * FROM price_items WHERE hospital_id = ? ORDER BY id", (row["id"],)).fetchall()
    # 병원에서 가까운 숙소 2곳·K푸드 2곳
    nearby = sorted((partner_view(p, g.lang) for p in db.execute(
        "SELECT * FROM partners WHERE active = 1 AND region = ? AND partner_type IN ('hotel', 'kfood')",
        (row["region"],)).fetchall()),
        key=lambda p: distance_km(row["lat"], row["lng"], p["lat"], p["lng"]))
    partners = [p for p in nearby if p["partner_type"] == "hotel"][:2] + [p for p in nearby if p["partner_type"] == "kfood"][:2]
    track("view_hospital", row["id"])
    return render_template("public/hospital_detail.html", h=h, doctors=doctors, prices=prices, partners=partners,
                           points=map_points([h]), spots=load_attractions(g.lang, region=row["region"], limit=3),
                           ld=hospital_ld(h, prices))


def hospital_ld(h, prices):
    """병원 구조화 데이터: 이름·주소·좌표·진료과·공개 가격 (출처: 심평원)."""
    base = notify.site_url()
    url = base + url_for("public.hospital_detail", slug=h["slug"])
    kind = "Dentist" if h["hospital_type"] == "dental_hospital" else "Hospital"
    ld = {"@context": "https://schema.org", "@type": kind, "@id": url + "#hospital", "name": h["name"],
          "alternateName": h["name_ko"], "url": url, "description": h["summary"],
          "address": {"@type": "PostalAddress", "streetAddress": h["address_ko"], "addressLocality": h["district_name"],
                      "addressRegion": i18n.translate("rg_" + h["region"], "en"), "addressCountry": "KR"},
          "medicalSpecialty": [i18n.translate("sp_" + sp, "en") for sp in h["specialty_list"]]}
    if h["lat"] and h["lng"]:
        ld["geo"] = {"@type": "GeoCoordinates", "latitude": h["lat"], "longitude": h["lng"]}
    if h.get("phone"):
        ld["telephone"] = h["phone"]
    if h.get("website"):
        ld["sameAs"] = [h["website"]]
    offers = [{"@type": "Offer", "name": item_name(p, g.lang), "price": p["price_domestic_krw"], "priceCurrency": "KRW",
               "validFrom": p["verified_at"]} for p in prices if p["price_domestic_krw"]]
    if offers:
        ld["makesOffer"] = offers
    return ld


@bp.route("/<lang>/concierge")
def concierge():
    rows = get_db().execute("SELECT * FROM partners WHERE active = 1 ORDER BY partner_type, name").fetchall()
    grouped = {pt: [] for pt in i18n.PARTNER_TYPES}
    for r in rows:
        grouped.setdefault(r["partner_type"], []).append(partner_view(r, g.lang))
    track("view_concierge")
    return render_template("public/concierge.html", grouped=grouped, spots=load_attractions(g.lang, limit=24))


@bp.route("/<lang>/how-it-works")
def guide():
    return render_template("public/guide.html")


@bp.route("/<lang>/privacy")
def privacy():
    return render_template("public/privacy.html")


# ---------------------------------------------------------------------------
# K-팝 뮤직비디오·드라마·영화 촬영지
# ---------------------------------------------------------------------------
FILMING_KINDS = ("mv", "drama", "movie")
SCHEMA_WORK = {"mv": "MusicVideoObject", "drama": "TVSeries", "movie": "Movie"}


def load_filming(lang, kind=None):
    spots = []
    for r in get_db().execute("SELECT * FROM filming_spots WHERE active = 1 ORDER BY region DESC, id").fetchall():
        if kind and kind not in r["kinds"].split(","):
            continue
        names = json.loads(r["name_i18n"] or "{}")
        works = []
        for w in json.loads(r["works_json"] or "[]"):
            if kind and w["kind"] != kind:
                continue
            # 장면 설명은 국문(원자료)·영문만 있다 → 중·일·베트남어는 작품명과 종류만 보여 준다
            works.append({"kind": w["kind"], "title": w["title"].get(lang) or w["title"]["en"],
                          "scene": w["scene"].get(lang, "") if lang in ("ko", "en") else ""})
        spots.append(dict(r, name=names.get(lang) or names.get("en") or r["place_key"], works=works))
    return spots


@bp.route("/<lang>/filming-locations")
def filming():
    kind = request.args.get("k")
    kind = kind if kind in FILMING_KINDS else None
    spots = load_filming(g.lang, kind)
    base, page = notify.site_url(), url_for("public.filming")
    points = [{"name": s["name"], "lat": s["lat"], "lng": s["lng"], "url": f"#spot-{s['id']}",
               "district": i18n.translate("rg_" + s["region"], g.lang)} for s in spots if s["lat"]]
    ld = {"@context": "https://schema.org", "@type": "CollectionPage", "inLanguage": g.lang,
          "name": i18n.translate("filming_title", g.lang), "description": i18n.translate("filming_sub", g.lang),
          "url": base + page,
          "mainEntity": {"@type": "ItemList", "itemListElement": [
              {"@type": "ListItem", "position": n, "item": {
                  "@type": "TouristAttraction", "name": s["name"], "url": f"{base}{page}#spot-{s['id']}",
                  "address": s["address"], "image": s["image_url"] or None,
                  "geo": {"@type": "GeoCoordinates", "latitude": s["lat"], "longitude": s["lng"]},
                  "subjectOf": [{"@type": SCHEMA_WORK[w["kind"]], "name": w["title"]} for w in s["works"]]}}
              for n, s in enumerate(spots, 1)]}}
    track("view_filming")
    return render_template("public/filming.html", spots=spots, kind=kind, points=points, ld=ld)


# ---------------------------------------------------------------------------
# 매거진 (K-의료·K-뷰티·K-팝·K-푸드·여행 정보 글)
# ---------------------------------------------------------------------------
@bp.route("/<lang>/magazine")
def magazine():
    cat = request.args.get("c")
    cat = cat if cat in mag.CATEGORIES else None
    items = mag.listing(g.lang, cat)
    base = notify.site_url()
    ld = {"@context": "https://schema.org", "@type": "CollectionPage", "inLanguage": g.lang,
          "name": i18n.translate("mag_title", g.lang), "description": i18n.translate("mag_sub", g.lang),
          "url": base + url_for("public.magazine"),
          "mainEntity": {"@type": "ItemList", "itemListElement": [
              {"@type": "ListItem", "position": n, "name": a["title"], "url": base + url_for("public.article", slug=a["slug"])}
              for n, a in enumerate(items, 1)]}}
    track("view_magazine")
    return render_template("public/magazine.html", items=items, cat=cat, ld=ld)


@bp.route("/<lang>/magazine/<slug>")
def article(slug):
    a = mag.get(slug, g.lang)
    if a is None:
        abort(404)
    same = [x for x in mag.listing(g.lang, a["category"]) if x["slug"] != slug]
    others = [x for x in mag.listing(g.lang) if x["slug"] != slug and x["category"] != a["category"]]
    track("view_article")
    return render_template("public/article.html", a=a, related=(same + others)[:3], languages=mag.languages(slug),
                           ld=article_ld(a))


def article_ld(a):
    """Article + BreadcrumbList + FAQPage 구조화 데이터 (AEO: 답변 엔진이 요약·FAQ 를 바로 인용할 수 있게)."""
    base, brand = notify.site_url(), current_app.config["BRAND_NAME"]
    url = base + url_for("public.article", slug=a["slug"], lang=a["lang"])
    graph = [
        {"@type": "Article", "@id": url + "#article", "headline": a["title"], "description": a["description"],
         "abstract": a["answer"], "image": [a["hero"]] if a["hero"] else [], "inLanguage": a["lang"],
         "datePublished": a["published"], "dateModified": a["updated"], "keywords": ", ".join(a["tags"]),
         "articleSection": i18n.translate("mag_" + a["category"], a["lang"]), "mainEntityOfPage": url,
         "author": {"@type": "Organization", "name": f"{brand} Editorial", "url": base + "/"},
         "publisher": {"@id": base + "/#org"}, "isAccessibleForFree": True,
         "speakable": {"@type": "SpeakableSpecification", "cssSelector": [".answer-box", "h1"]},
         "citation": [m.group(1) for m in (re.search(r"\((https?://[^)\s]+)\)", s) for s in a["sources"]) if m]},
        {"@type": "BreadcrumbList", "itemListElement": [
            {"@type": "ListItem", "position": 1, "name": brand, "item": base + url_for("public.home", lang=a["lang"])},
            {"@type": "ListItem", "position": 2, "name": i18n.translate("nav_magazine", a["lang"]),
             "item": base + url_for("public.magazine", lang=a["lang"])},
            {"@type": "ListItem", "position": 3, "name": a["title"], "item": url}]},
    ]
    if a["faq"]:
        graph.append({"@type": "FAQPage", "@id": url + "#faq", "mainEntity": [
            {"@type": "Question", "name": f["q"], "acceptedAnswer": {"@type": "Answer", "text": f["a"]}} for f in a["faq"]]})
    return {"@context": "https://schema.org", "@graph": graph}


# ---------------------------------------------------------------------------
# 검색엔진·AI 크롤러용: robots.txt, sitemap.xml, llms.txt (SEO/AEO/GEO)
# ---------------------------------------------------------------------------
AI_CRAWLERS = ["GPTBot", "OAI-SearchBot", "ChatGPT-User", "ClaudeBot", "Claude-User", "Claude-SearchBot", "anthropic-ai",
               "PerplexityBot", "Perplexity-User", "Google-Extended", "Applebot-Extended", "CCBot", "meta-externalagent",
               "Amazonbot", "DuckAssistBot", "MistralAI-User", "cohere-ai", "Bytespider"]
SEARCH_CRAWLERS = ["Googlebot", "Bingbot", "Yeti", "Daum", "Baiduspider", "YandexBot", "Applebot", "DuckDuckBot"]


def _text(body):
    return current_app.response_class(body, mimetype="text/plain")


@bp.route("/robots.txt")
def robots():
    base = notify.site_url()
    # 검색엔진과 AI 답변 엔진 모두 공개 페이지를 읽을 수 있게 열어 둔다 (관리자·상담 접수 결과는 제외)
    lines = ["# Public pages are open to search engines and AI answer engines. See /llms.txt for a summary.",
             "User-agent: *", "Allow: /", "Disallow: /admin", "Disallow: /*/consult/thanks/", ""]
    for bot in SEARCH_CRAWLERS + AI_CRAWLERS:
        lines += [f"User-agent: {bot}", "Allow: /", "Disallow: /admin", ""]
    lines += [f"Sitemap: {base}/sitemap.xml"]
    return _text("\n".join(lines) + "\n")


def _sitemap_groups():
    """(엔드포인트, 인자, 해당 언어 목록, 최종 수정일) — 언어별 URL 과 hreflang 대체 링크를 만든다."""
    langs = list(i18n.LANGUAGES)
    groups = [(ep, {}, langs, None) for ep in ("public.home", "public.hospitals", "public.magazine", "public.filming", "public.concierge",
                                               "public.guide", "public.consult", "public.privacy")]
    for slug, versions in mag.ARTICLES.items():
        updated = max(a["updated"] for a in versions.values())
        groups.append(("public.article", {"slug": slug}, sorted(versions), updated))
    rows = get_db().execute("SELECT slug, last_verified_at FROM hospitals WHERE verification_status = 'published'").fetchall()
    for r in rows:
        groups.append(("public.hospital_detail", {"slug": r["slug"]}, langs, r["last_verified_at"]))
    return groups


@bp.route("/sitemap.xml")
def sitemap():
    base = notify.site_url()
    out = ['<?xml version="1.0" encoding="UTF-8"?>',
           '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:xhtml="http://www.w3.org/1999/xhtml">']
    for ep, kw, langs, updated in _sitemap_groups():
        urls = {l: base + url_for(ep, lang=l, **kw) for l in langs}
        alts = "".join(f'<xhtml:link rel="alternate" hreflang="{l}" href="{u}"/>' for l, u in urls.items())
        alts += f'<xhtml:link rel="alternate" hreflang="x-default" href="{urls.get("en", next(iter(urls.values())))}"/>'
        lastmod = f"<lastmod>{updated[:10]}</lastmod>" if updated else ""
        for u in urls.values():
            out.append(f"<url><loc>{u}</loc>{lastmod}{alts}</url>")
    out.append("</urlset>")
    return current_app.response_class("\n".join(out), mimetype="application/xml")


@bp.route("/llms.txt")
def llms_txt():
    """llmstxt.org 형식: AI 가 사이트를 이해하고 인용하기 쉽게 핵심 페이지와 글을 정리한다."""
    base, brand = notify.site_url(), current_app.config["BRAND_NAME"]
    db = get_db()
    n_hosp = db.execute("SELECT COUNT(*) FROM hospitals WHERE verification_status = 'published'").fetchone()[0]
    n_price = db.execute("SELECT COUNT(*) FROM price_items p JOIN hospitals h ON h.id = p.hospital_id "
                         "WHERE h.verification_status = 'published'").fetchone()[0]
    lines = [f"# {brand}", "",
             f"> Multilingual guide to medical care and travel in Korea for international visitors: {n_hosp} hospitals in "
             f"Seoul and Busan with {n_price} self-pay prices published to HIRA (Korea's Health Insurance Review & Assessment "
             "Service), places foreign visitors love, and in-depth guides on K-medical, K-beauty, K-pop, K-food and travel.", "",
             "Content is available in English, Korean, Chinese (Simplified), Japanese and Vietnamese. "
             "Prices come from official public data with a reference date; information is not medical advice.", "",
             "## Key pages", "",
             f"- [Find hospitals]({base}/en/hospitals): map and filters by specialty, region and language",
             f"- [Travel & concierge]({base}/en/concierge): stays, K-food, shopping and sights near hospitals",
             f"- [K-pop, K-drama & film locations]({base}/en/filming-locations): music video, drama and movie filming spots in Seoul and Busan",
             f"- [How it works]({base}/en/how-it-works): consultation to aftercare in five steps",
             f"- [Free consultation]({base}/en/consult): coordinator replies in the visitor's language",
             f"- [Full text for LLMs]({base}/llms-full.txt)", "", "## Magazine", ""]
    for a in mag.listing("en"):
        lines.append(f"- [{a['title']}]({base}/en/magazine/{a['slug']}): {a['description']}")
    return _text("\n".join(lines) + "\n")


@bp.route("/llms-full.txt")
def llms_full_txt():
    base = notify.site_url()
    raw_files = mag._raw_files()
    parts = [f"# {current_app.config['BRAND_NAME']} — full text of guides (English)", ""]
    for a in mag.listing("en"):
        raw = raw_files.get(f"{a['slug']}.{a['lang']}.md", "")
        body = re.sub(r"^---\n.*?\n---\n", "", raw.replace("\r", ""), flags=re.S).replace("{lang}", "en")
        parts += [f"## {a['title']}", f"URL: {base}/en/magazine/{a['slug']}", f"Updated: {a['updated']}", "",
                  f"Quick answer: {a['answer']}", "", body.strip(), "", "---", ""]
    return _text("\n".join(parts))


def _int_or_none(value, lo, hi):
    try:
        v = int(value)
    except (TypeError, ValueError):
        return None
    return v if lo <= v <= hi else None


def parse_consult_form(form, lang):
    """폼 입력을 정규화하고 (data, consent, errors) 를 반환한다."""
    s = lambda k, n=200: (form.get(k) or "").strip()[:n]
    data = {
        "full_name": s("full_name", 120),
        "email": s("email", 200),
        "phone": s("phone", 40),
        "contact_channel": s("contact_channel", 20),
        "messenger_id": s("messenger_id", 120),
        "nationality": s("nationality", 60),
        "country_of_residence": s("country_of_residence", 60),
        "preferred_language": s("preferred_language", 5) or lang,
        "age_band": s("age_band", 10) if s("age_band", 10) in i18n.AGE_BANDS else "",
        "treatment_category": s("treatment_category", 40),
        "treatment_interest": s("treatment_interest", 300),
        "hospital_preferences": [p for p in form.getlist("hospital_preferences") if p in i18n.HOSPITAL_PREFS],
        "preferred_hospital_id": _int_or_none(form.get("preferred_hospital_id"), 1, 10**9),
        "budget_band": s("budget_band", 20) if s("budget_band", 20) in i18n.BUDGET_BANDS else "",
        "preferred_visit_date": s("preferred_visit_date", 10),
        "stay_nights": _int_or_none(form.get("stay_nights"), 0, 365),
        "travel_party_size": _int_or_none(form.get("travel_party_size"), 1, 50),
        "concierge_needs": [n for n in form.getlist("concierge_needs") if n in i18n.CONCIERGE_NEEDS],
        "accommodation_pref": s("accommodation_pref", 300),
        "food_pref": s("food_pref", 300),
        "itinerary_interests": s("itinerary_interests", 500),
        "vip_request": form.get("vip_request") == "1",
        "message": s("message", 2000),
        "source_channel": s("source_channel", 20) if s("source_channel", 20) in i18n.HEARD_OPTIONS else "",
        "site_language": lang,
    }
    consent = {
        "privacy_required": form.get("consent_privacy") == "1",
        "third_party_share": form.get("consent_third") == "1",
        "overseas_transfer": form.get("consent_overseas") == "1",
        "marketing": form.get("consent_marketing") == "1",
        "policy_version": current_app.config["PRIVACY_POLICY_VERSION"],
    }
    errors = []
    if (not data["full_name"] or data["contact_channel"] not in i18n.CONTACT_CHANNELS
            or data["preferred_language"] not in i18n.LANGUAGES
            or data["treatment_category"] not in i18n.SPECIALTIES + ["other"]):
        errors.append("err_required")
    if not data["email"] and not data["messenger_id"] and not data["phone"]:
        errors.append("err_contact")
    if data["email"] and not crm.EMAIL_RE.match(data["email"]):
        errors.append("err_email")
    if not consent["privacy_required"]:
        errors.append("err_consent")
    if data["preferred_hospital_id"]:
        exists = get_db().execute("SELECT 1 FROM hospitals WHERE id = ? AND verification_status = 'published'",
                                  (data["preferred_hospital_id"],)).fetchone()
        if not exists:
            data["preferred_hospital_id"] = None
    return data, consent, errors


@bp.route("/<lang>/consult", methods=["GET", "POST"])
def consult():
    db = get_db()
    hospital_rows = db.execute(
        "SELECT * FROM hospitals WHERE verification_status = 'published' ORDER BY name_en").fetchall()
    hospital_options = [hospital_view(r, g.lang) for r in hospital_rows]
    errors, form = [], request.form

    if request.method == "POST":
        # 봇 차단용 허니팟 필드: 사람은 보이지 않아 비워 둔다.
        if form.get("website"):
            return redirect(url_for("public.consult_thanks", ref="received"))
        data, consent, errors = parse_consult_form(form, g.lang)
        if not errors:
            lead_id = crm.create_lead(data, consent, session.get("attribution", {}))
            notify.notify_new_lead(lead_id)
            notify.send_patient_ack(lead_id)
            return redirect(url_for("public.consult_thanks", ref=lead_id[:8].upper()))

    prefill = {
        "hospital": request.args.get("hospital", type=int),
        "specialty": request.args.get("specialty", ""),
        "needs": request.args.getlist("need"),
    }
    return render_template("public/consult.html", hospitals=hospital_options, errors=errors, form=form,
                           prefill=prefill), (400 if errors else 200)


@bp.route("/<lang>/consult/thanks/<ref>")
def consult_thanks(ref):
    return render_template("public/thanks.html", ref=ref)
