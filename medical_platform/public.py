"""해외 환자용 다국어 공개 사이트: 홈, 병원 지도·검색, 병원 상세, 컨시어지, 상담 신청."""
import json

from flask import (Blueprint, abort, current_app, g, redirect, render_template, request, session, url_for)

from . import crm, i18n, notify
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


def hospital_view(row, lang, tags=None):
    h = dict(row)
    h["name"] = h["name_ko"] if lang == "ko" else h["name_en"]
    h["address"] = h["address_ko"] if lang == "ko" else h["address_en"]
    h["summary"] = localized(h["summary_i18n"], lang)
    h["specialty_list"] = [s for s in h["specialties"].split(",") if s]
    h["language_list"] = [s for s in h["languages"].split(",") if s]
    h["tags"] = tags or []
    return h


def load_tags(hospital_ids):
    if not hospital_ids:
        return {}
    rows = get_db().execute(
        f"SELECT * FROM hospital_tags WHERE hospital_id IN ({','.join('?' * len(hospital_ids))})",
        hospital_ids).fetchall()
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
    return [
        {"name": h["name"], "lat": h["lat"], "lng": h["lng"],
         "url": url_for("public.hospital_detail", slug=h["slug"]),
         "district": h["district"]}
        for h in hospitals if h["lat"] is not None
    ]


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
             SELECT p.*, h.slug, h.name_en, h.name_ko, h.is_demo,
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
                           spots=load_attractions(g.lang, limit=8))


@bp.route("/<lang>/hospitals")
def hospitals():
    results = search_hospitals(request.args, g.lang)
    track("search")
    return render_template("public/hospitals.html", hospitals=results, points=map_points(results), args=request.args)


@bp.route("/<lang>/hospitals/<slug>")
def hospital_detail(slug):
    db = get_db()
    row = db.execute("SELECT * FROM hospitals WHERE slug = ? AND verification_status = 'published'", (slug,)).fetchone()
    if row is None:
        abort(404)
    h = hospital_view(row, g.lang, load_tags([row["id"]]).get(row["id"]))
    doctors = db.execute("SELECT * FROM doctors WHERE hospital_id = ? ORDER BY id", (row["id"],)).fetchall()
    prices = db.execute("SELECT * FROM price_items WHERE hospital_id = ? ORDER BY id", (row["id"],)).fetchall()
    partners = db.execute(
        "SELECT * FROM partners WHERE active = 1 AND region = ? AND partner_type IN ('hotel', 'kfood') ORDER BY partner_type, name",
        (row["region"],)).fetchall()
    partners = [dict(p, description=localized(p["description_i18n"], g.lang)) for p in partners]
    track("view_hospital", row["id"])
    return render_template("public/hospital_detail.html", h=h, doctors=doctors, prices=prices, partners=partners,
                           points=map_points([h]), spots=load_attractions(g.lang, region=row["region"], limit=3))


@bp.route("/<lang>/concierge")
def concierge():
    rows = get_db().execute("SELECT * FROM partners WHERE active = 1 ORDER BY partner_type, name").fetchall()
    grouped = {pt: [] for pt in i18n.PARTNER_TYPES}
    for r in rows:
        grouped.setdefault(r["partner_type"], []).append(dict(r, description=localized(r["description_i18n"], g.lang)))
    track("view_concierge")
    return render_template("public/concierge.html", grouped=grouped, spots=load_attractions(g.lang, limit=24))


@bp.route("/<lang>/how-it-works")
def guide():
    return render_template("public/guide.html")


@bp.route("/<lang>/privacy")
def privacy():
    return render_template("public/privacy.html")


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
