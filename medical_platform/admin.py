"""내부 관리자: 대시보드, 환자 CRM, 병원 DB 관리, 제휴사, 직원 권한, 감사 로그.

권한
    admin        전체 기능
    coordinator  리드 상담·병원 연결·서비스 요청 처리, 병원 정보 열람
    marketer     대시보드·집계 지표만 (개인정보 마스킹, 상세·다운로드 불가)
"""
import csv
import io
import json
import re
from datetime import datetime, timedelta, timezone
from functools import wraps

from flask import (Blueprint, Response, abort, current_app, flash, g, redirect, render_template, request, session, url_for)
from werkzeug.security import check_password_hash, generate_password_hash

from . import crm, i18n, notify
from .db import get_db, now_iso

bp = Blueprint("admin", __name__, url_prefix="/admin")

VERIFICATION_STATUSES = [
    ("planned", "수집 예정"), ("collected", "자료 확보"), ("verified", "검증 완료"),
    ("published", "게시"), ("needs_update", "갱신 필요"), ("hidden", "비공개"),
]
HOSPITAL_TYPES = [("clinic", "의원"), ("hospital", "병원"), ("general_hospital", "종합병원"),
                  ("korean_medicine", "한방병원"), ("dental", "치과병원")]
ROLES = [("admin", "관리자"), ("coordinator", "코디네이터"), ("marketer", "마케터")]


@bp.app_template_filter("ko_label")
def ko_label(code, kind):
    table = {
        "status": crm.STATUS_LABELS,
        "referral": dict(crm.REFERRAL_STATUSES),
        "service": dict(crm.SERVICE_STATUSES),
        "verification": dict(VERIFICATION_STATUSES),
        "role": dict(ROLES),
        "priority": dict(crm.PRIORITIES),
        "htype": dict(HOSPITAL_TYPES),
    }[kind]
    return table.get(code, code)


@bp.app_template_filter("i18n_ko")
def i18n_ko(code, prefix):
    return i18n.translate(f"{prefix}_{code}", "ko") if code else ""


@bp.app_template_filter("dt")
def fmt_dt(value):
    if not value:
        return "-"
    try:
        d = datetime.fromisoformat(value).astimezone(timezone(timedelta(hours=9)))
        return d.strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return value


@bp.before_app_request
def load_staff():
    g.staff = None
    sid = session.get("staff_id")
    if sid:
        g.staff = get_db().execute("SELECT * FROM staff WHERE id = ? AND active = 1", (sid,)).fetchone()


def login_required(*roles):
    def deco(view):
        @wraps(view)
        def wrapped(*args, **kwargs):
            if g.staff is None:
                return redirect(url_for("admin.login", next=request.path))
            if roles and g.staff["role"] not in roles:
                abort(403)
            return view(*args, **kwargs)
        return wrapped
    return deco


def can_see_pii():
    return g.staff is not None and g.staff["role"] in ("admin", "coordinator")


@bp.app_context_processor
def admin_ctx():
    return {"staff": getattr(g, "staff", None), "can_see_pii": can_see_pii, "crm": crm,
            "VERIFICATION_STATUSES": VERIFICATION_STATUSES, "HOSPITAL_TYPES": HOSPITAL_TYPES, "ROLES": ROLES}


# ---------------------------------------------------------------------------
# 인증
# ---------------------------------------------------------------------------

@bp.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        row = get_db().execute("SELECT * FROM staff WHERE username = ? AND active = 1",
                               (request.form.get("username", ""),)).fetchone()
        if row and check_password_hash(row["password_hash"], request.form.get("password", "")):
            csrf = session.get("_csrf")
            session.clear()
            session["_csrf"] = csrf
            session["staff_id"] = row["id"]
            crm.audit(row["id"], "login")
            get_db().commit()
            nxt = request.args.get("next", "")
            return redirect(nxt if nxt.startswith("/admin") else url_for("admin.dashboard"))
        crm.audit(None, "login_failed", request.form.get("username", "")[:60])
        get_db().commit()
        flash("아이디 또는 비밀번호가 올바르지 않습니다.", "error")
    return render_template("admin/login.html")


@bp.route("/logout", methods=["POST"])
def logout():
    if g.staff:
        crm.audit(g.staff["id"], "logout")
        get_db().commit()
    session.pop("staff_id", None)
    return redirect(url_for("admin.login"))


# ---------------------------------------------------------------------------
# 대시보드
# ---------------------------------------------------------------------------

def _group(db, column, where="1=1", params=(), limit=10):
    return db.execute(
        f"SELECT COALESCE(NULLIF({column}, ''), '(미입력)') AS k, COUNT(*) AS n FROM leads WHERE {where} "
        f"GROUP BY k ORDER BY n DESC LIMIT {int(limit)}", params).fetchall()


@bp.route("/")
@login_required()
def dashboard():
    db = get_db()
    days = request.args.get("days", 30, type=int)
    since = (datetime.now(timezone.utc) - timedelta(days=days)).replace(microsecond=0).isoformat()
    where, params = "created_at >= ?", (since,)

    status_counts = {r["lead_status"]: r["n"] for r in db.execute(
        f"SELECT lead_status, COUNT(*) AS n FROM leads WHERE {where} GROUP BY lead_status", params)}
    total = sum(status_counts.values())
    # 퍼널: 각 단계에 '도달한' 리드 수 (현재 단계 이후 포함)
    funnel = []
    for i, st in enumerate(crm.FUNNEL):
        reached = sum(status_counts.get(s, 0) for s in crm.FUNNEL[i:]) + (status_counts.get("followup", 0))
        funnel.append((st, reached))
    valid = total - status_counts.get("invalid", 0)

    sla = db.execute(
        f"""SELECT AVG((julianday(first_response_at) - julianday(created_at)) * 1440) AS avg_min,
                   SUM(CASE WHEN (julianday(first_response_at) - julianday(created_at)) * 1440 <= ? THEN 1 ELSE 0 END) AS within,
                   COUNT(first_response_at) AS responded
            FROM leads WHERE {where}""", (crm.FIRST_RESPONSE_SLA_MIN, *params)).fetchone()
    overdue_cut = (datetime.now(timezone.utc) - timedelta(minutes=crm.FIRST_RESPONSE_SLA_MIN)).replace(microsecond=0).isoformat()
    overdue = db.execute(
        "SELECT COUNT(*) FROM leads WHERE lead_status = 'new' AND first_response_at IS NULL AND created_at < ?",
        (overdue_cut,)).fetchone()[0]
    today_end = (datetime.now(timezone.utc).replace(hour=23, minute=59)).isoformat()
    due_followups = db.execute(
        "SELECT COUNT(*) FROM leads WHERE next_contact_at IS NOT NULL AND next_contact_at <= ? AND lead_status NOT IN ('visited','lost','invalid')",
        (today_end,)).fetchone()[0]

    needs = {}
    for r in db.execute(f"SELECT concierge_needs, vip_request FROM leads WHERE {where}", params):
        for n in filter(None, r["concierge_needs"].split(",")):
            needs[n] = needs.get(n, 0) + 1
    prefs = {}
    for r in db.execute(f"SELECT hospital_preferences FROM leads WHERE {where}", params):
        for p in filter(None, r["hospital_preferences"].split(",")):
            prefs[p] = prefs.get(p, 0) + 1

    top_hospitals = db.execute(
        """SELECT h.name_ko, COUNT(e.id) AS views FROM page_events e JOIN hospitals h ON h.id = e.hospital_id
           WHERE e.event = 'view_hospital' AND e.created_at >= ? GROUP BY h.id ORDER BY views DESC LIMIT 8""",
        (since,)).fetchall()
    events = {r["event"]: r["n"] for r in db.execute(
        "SELECT event, COUNT(*) AS n FROM page_events WHERE created_at >= ? GROUP BY event", (since,))}
    visits = events.get("view_home", 0) + events.get("search", 0) + events.get("view_hospital", 0) + events.get("view_concierge", 0)

    return render_template(
        "admin/dashboard.html", days=days, total=total, valid=valid, status_counts=status_counts, funnel=funnel,
        sla=sla, overdue=overdue, due_followups=due_followups,
        by_language=_group(db, "preferred_language", where, params),
        by_country=_group(db, "country_of_residence", where, params),
        by_treatment=_group(db, "treatment_category", where, params),
        by_source=_group(db, "COALESCE(NULLIF(utm_source,''), source_channel)", where, params),
        by_campaign=_group(db, "utm_campaign", where, params),
        by_budget=_group(db, "budget_band", where, params),
        needs=sorted(needs.items(), key=lambda x: -x[1]), prefs=sorted(prefs.items(), key=lambda x: -x[1]),
        top_hospitals=top_hospitals, events=events, visits=visits,
    )


# ---------------------------------------------------------------------------
# 환자 CRM
# ---------------------------------------------------------------------------

LEAD_FILTERS = ("status", "lang", "assignee", "treatment", "q", "priority", "mine")


def _lead_query(args):
    sql = """SELECT l.*, s.name AS staff_name, h.name_ko AS hospital_name FROM leads l
             LEFT JOIN staff s ON s.id = l.assigned_staff_id
             LEFT JOIN hospitals h ON h.id = l.preferred_hospital_id WHERE 1=1"""
    params = []
    if args.get("status") == "open":
        sql += f" AND l.lead_status IN ({','.join('?' * len(crm.OPEN_STATUSES))})"
        params += list(crm.OPEN_STATUSES)
    elif args.get("status") in crm.STATUS_LABELS:
        sql += " AND l.lead_status = ?"
        params.append(args["status"])
    if args.get("lang"):
        sql += " AND l.preferred_language = ?"
        params.append(args["lang"])
    if args.get("treatment"):
        sql += " AND l.treatment_category = ?"
        params.append(args["treatment"])
    if args.get("priority"):
        sql += " AND l.priority = ?"
        params.append(args["priority"])
    if args.get("assignee"):
        sql += " AND l.assigned_staff_id = ?"
        params.append(args.get("assignee", type=int))
    if args.get("mine") == "1" and g.staff:
        sql += " AND l.assigned_staff_id = ?"
        params.append(g.staff["id"])
    if args.get("q") and can_see_pii():
        like = f"%{args['q']}%"
        sql += " AND (l.full_name LIKE ? OR l.email LIKE ? OR l.phone LIKE ? OR l.messenger_id LIKE ? OR l.id LIKE ?)"
        params += [like] * 4 + [f"{args['q'].lower()}%"]
    sql += " ORDER BY CASE l.priority WHEN 'high' THEN 0 WHEN 'normal' THEN 1 ELSE 2 END, l.created_at DESC"
    return sql, params


@bp.route("/leads")
@login_required()
def leads():
    db = get_db()
    sql, params = _lead_query(request.args)
    rows = db.execute(sql + " LIMIT 500", params).fetchall()
    staff_list = db.execute("SELECT id, name FROM staff WHERE active = 1 AND role != 'marketer' ORDER BY name").fetchall()
    return render_template("admin/leads.html", leads=rows, staff_list=staff_list, args=request.args,
                           sla_cut=(datetime.now(timezone.utc) - timedelta(minutes=crm.FIRST_RESPONSE_SLA_MIN)).isoformat())


@bp.route("/leads/export.csv")
@login_required("admin")
def leads_export():
    db = get_db()
    sql, params = _lead_query(request.args)
    rows = db.execute(sql, params).fetchall()
    cols = ["id", "created_at", "lead_status", "priority", "staff_name", "full_name", "email", "phone",
            "contact_channel", "messenger_id", "nationality", "country_of_residence", "preferred_language",
            "age_band", "treatment_category", "treatment_interest", "hospital_preferences", "hospital_name",
            "budget_band", "preferred_visit_date", "stay_nights", "travel_party_size", "concierge_needs",
            "accommodation_pref", "food_pref", "itinerary_interests", "vip_request", "source_channel",
            "utm_source", "utm_medium", "utm_campaign", "first_response_at", "last_contact_at", "next_contact_at"]
    buf = io.StringIO()
    w = csv.writer(buf)
    w.writerow(cols)
    for r in rows:
        w.writerow([r[c] if r[c] is not None else "" for c in cols])
    crm.audit(g.staff["id"], "leads_export", f"{len(rows)} rows")
    db.commit()
    return Response("﻿" + buf.getvalue(), mimetype="text/csv; charset=utf-8",
                    headers={"Content-Disposition": "attachment; filename=leads.csv"})


def _get_lead(lead_id):
    row = get_db().execute(
        """SELECT l.*, s.name AS staff_name, h.name_ko AS hospital_name FROM leads l
           LEFT JOIN staff s ON s.id = l.assigned_staff_id
           LEFT JOIN hospitals h ON h.id = l.preferred_hospital_id WHERE l.id = ?""", (lead_id,)).fetchone()
    if row is None:
        abort(404)
    return row


@bp.route("/leads/<lead_id>")
@login_required("admin", "coordinator")
def lead_detail(lead_id):
    db = get_db()
    lead = _get_lead(lead_id)
    crm.audit(g.staff["id"], "lead_view", lead_id)
    db.commit()
    return render_template(
        "admin/lead_detail.html", lead=lead,
        consent=db.execute("SELECT * FROM consents WHERE lead_id = ? ORDER BY id DESC LIMIT 1", (lead_id,)).fetchone(),
        activities=db.execute(
            """SELECT a.*, s.name AS staff_name FROM lead_activities a LEFT JOIN staff s ON s.id = a.staff_id
               WHERE a.lead_id = ? ORDER BY a.id DESC""", (lead_id,)).fetchall(),
        referrals=db.execute(
            """SELECT r.*, h.name_ko FROM referrals r JOIN hospitals h ON h.id = r.hospital_id
               WHERE r.lead_id = ? ORDER BY r.id""", (lead_id,)).fetchall(),
        services=db.execute(
            """SELECT o.*, p.name AS partner_name FROM service_orders o LEFT JOIN partners p ON p.id = o.partner_id
               WHERE o.lead_id = ? ORDER BY o.id""", (lead_id,)).fetchall(),
        hospitals=db.execute("SELECT id, name_ko FROM hospitals WHERE verification_status IN ('verified','published') ORDER BY name_ko").fetchall(),
        partners=db.execute("SELECT id, name, partner_type FROM partners WHERE active = 1 ORDER BY partner_type, name").fetchall(),
        staff_list=db.execute("SELECT id, name, languages FROM staff WHERE active = 1 AND role != 'marketer' ORDER BY name").fetchall(),
        duplicates=db.execute("SELECT id, created_at, lead_status FROM leads WHERE duplicate_of = ? OR id = ?",
                              (lead_id, lead["duplicate_of"] or "")).fetchall(),
    )


def _parse_local_dt(value):
    """관리자 화면의 KST datetime-local 입력을 UTC ISO 로 변환."""
    if not value:
        return None
    try:
        d = datetime.fromisoformat(value).replace(tzinfo=timezone(timedelta(hours=9)))
    except ValueError:
        return None
    return d.astimezone(timezone.utc).replace(microsecond=0).isoformat()


@bp.route("/leads/<lead_id>/action", methods=["POST"])
@login_required("admin", "coordinator")
def lead_action(lead_id):
    db = get_db()
    lead = _get_lead(lead_id)
    if lead["anonymized_at"]:
        flash("삭제(비식별화)된 리드는 수정할 수 없습니다.", "error")
        return redirect(url_for("admin.lead_detail", lead_id=lead_id))
    f = request.form
    action = f.get("action")
    sid = g.staff["id"]
    ts = now_iso()

    if action == "update":
        changes, notify_assignee = [], None
        status = f.get("lead_status")
        if status in crm.STATUS_LABELS and status != lead["lead_status"]:
            reason = f.get("disqualify_reason", "") if status in ("invalid", "lost") else ""
            db.execute("UPDATE leads SET lead_status = ?, disqualify_reason = ? WHERE id = ?", (status, reason, lead_id))
            changes.append(f"상태: {crm.STATUS_LABELS[lead['lead_status']]} → {crm.STATUS_LABELS[status]}"
                           + (f" ({reason})" if reason else ""))
        assignee = f.get("assigned_staff_id", type=int)
        if assignee and assignee != lead["assigned_staff_id"]:
            db.execute("UPDATE leads SET assigned_staff_id = ? WHERE id = ?", (assignee, lead_id))
            name = db.execute("SELECT name FROM staff WHERE id = ?", (assignee,)).fetchone()
            changes.append(f"담당자 변경: {name['name'] if name else assignee}")
            notify_assignee = assignee
        priority = f.get("priority")
        if priority in dict(crm.PRIORITIES) and priority != lead["priority"]:
            db.execute("UPDATE leads SET priority = ? WHERE id = ?", (priority, lead_id))
            changes.append(f"우선순위: {dict(crm.PRIORITIES)[priority]}")
        next_at = _parse_local_dt(f.get("next_contact_at"))
        if next_at != lead["next_contact_at"] and (next_at or f.get("next_contact_at") == ""):
            db.execute("UPDATE leads SET next_contact_at = ? WHERE id = ?", (next_at, lead_id))
            changes.append(f"다음 연락 예정: {fmt_dt(next_at) if next_at else '없음'}")
        if changes:
            db.execute("UPDATE leads SET updated_at = ? WHERE id = ?", (ts, lead_id))
            crm.log_activity(lead_id, "status", " / ".join(changes), sid)
        db.commit()
        if notify_assignee and notify_assignee != sid:
            notify.notify_assignment(lead_id, notify_assignee)

    elif action == "note":
        body = f.get("body", "").strip()[:4000]
        kind = "contact" if f.get("is_contact") == "1" else "note"
        if body:
            crm.log_activity(lead_id, kind, body, sid)
            if kind == "contact":
                db.execute("UPDATE leads SET last_contact_at = ?, first_response_at = COALESCE(first_response_at, ?), updated_at = ? WHERE id = ?",
                           (ts, ts, ts, lead_id))
                if lead["lead_status"] == "new":
                    db.execute("UPDATE leads SET lead_status = 'contacting' WHERE id = ?", (lead_id,))

    elif action == "add_referral":
        consent = db.execute("SELECT * FROM consents WHERE lead_id = ? ORDER BY id DESC LIMIT 1", (lead_id,)).fetchone()
        if not consent or not consent["third_party_share"] or consent["withdrawn_at"]:
            flash("제3자(병원) 제공 동의가 없습니다. 고객에게 동의를 받은 뒤 '동의 갱신'을 기록하고 연결하세요.", "error")
        else:
            hid = f.get("hospital_id", type=int)
            h = db.execute("SELECT name_ko FROM hospitals WHERE id = ?", (hid,)).fetchone()
            if h:
                db.execute("INSERT INTO referrals (lead_id, hospital_id, status, created_at, updated_at) VALUES (?,?,?,?,?)",
                           (lead_id, hid, "requested", ts, ts))
                crm.log_activity(lead_id, "status", f"병원 연결 추가: {h['name_ko']}", sid)
                if lead["lead_status"] in ("new", "contacting", "qualified"):
                    db.execute("UPDATE leads SET lead_status = 'matching' WHERE id = ?", (lead_id,))

    elif action == "update_referral":
        rid = f.get("referral_id", type=int)
        status = f.get("status")
        if status in dict(crm.REFERRAL_STATUSES):
            if status in ("sent", "replied", "quoted", "booked", "visited"):
                consent = db.execute("SELECT third_party_share, withdrawn_at FROM consents WHERE lead_id = ? ORDER BY id DESC LIMIT 1",
                                     (lead_id,)).fetchone()
                if not consent or not consent["third_party_share"] or consent["withdrawn_at"]:
                    flash("제3자 제공 동의가 없어 병원에 정보를 전달할 수 없습니다.", "error")
                    return redirect(url_for("admin.lead_detail", lead_id=lead_id))
            appt = _parse_local_dt(f.get("appointment_at"))
            db.execute("UPDATE referrals SET status = ?, quote_note = ?, appointment_at = COALESCE(?, appointment_at), updated_at = ? WHERE id = ? AND lead_id = ?",
                       (status, f.get("quote_note", "")[:1000], appt, ts, rid, lead_id))
            crm.log_activity(lead_id, "status", f"병원 연결 #{rid}: {dict(crm.REFERRAL_STATUSES)[status]}", sid)
            # 병원 단계가 리드 단계보다 앞서면 리드 상태를 끌어올린다
            sync = {"quoted": "consultation", "replied": "consultation", "booked": "booked", "visited": "visited"}
            target, current = sync.get(status), lead["lead_status"]
            if target and current in crm.FUNNEL and crm.FUNNEL.index(target) > crm.FUNNEL.index(current):
                db.execute("UPDATE leads SET lead_status = ? WHERE id = ?", (target, lead_id))

    elif action == "add_service":
        stype = f.get("service_type")
        if stype in i18n.CONCIERGE_NEEDS:
            db.execute("INSERT INTO service_orders (lead_id, service_type, partner_id, note, created_at, updated_at) VALUES (?,?,?,?,?,?)",
                       (lead_id, stype, f.get("partner_id", type=int), f.get("note", "")[:500], ts, ts))
            crm.log_activity(lead_id, "status", f"서비스 요청 추가: {i18n.translate('need_' + stype, 'ko')}", sid)

    elif action == "update_service":
        oid = f.get("order_id", type=int)
        status = f.get("status")
        if status in dict(crm.SERVICE_STATUSES):
            db.execute("UPDATE service_orders SET status = ?, partner_id = ?, note = ?, updated_at = ? WHERE id = ? AND lead_id = ?",
                       (status, f.get("partner_id", type=int), f.get("note", "")[:500], ts, oid, lead_id))
            crm.log_activity(lead_id, "status", f"서비스 요청 #{oid}: {dict(crm.SERVICE_STATUSES)[status]}", sid)

    elif action == "consent_update":
        # 상담 중 고객에게 추가 동의를 받은 경우 새 동의 레코드를 남긴다 (기존 기록은 보존)
        cur = db.execute(
            """INSERT INTO consents (lead_id, privacy_required, third_party_share, overseas_transfer, marketing,
               policy_version, ip, user_agent, created_at) VALUES (?,?,?,?,?,?,?,?,?)""",
            (lead_id, 1, int(f.get("third_party_share") == "1"), int(f.get("overseas_transfer") == "1"),
             int(f.get("marketing") == "1"), f"staff-recorded:{f.get('evidence', '')[:100]}", "", f"staff:{sid}", ts))
        db.execute("UPDATE leads SET consent_record_id = ? WHERE id = ?", (cur.lastrowid, lead_id))
        crm.log_activity(lead_id, "system", f"동의 갱신 기록 (근거: {f.get('evidence', '')[:100]})", sid)
        crm.audit(sid, "consent_update", lead_id)

    elif action == "withdraw_consent":
        db.execute("UPDATE consents SET withdrawn_at = ? WHERE lead_id = ? AND withdrawn_at IS NULL", (ts, lead_id))
        db.execute("UPDATE leads SET lead_status = 'lost', disqualify_reason = '동의 철회' WHERE id = ?", (lead_id,))
        crm.log_activity(lead_id, "system", "동의 철회 처리", sid)
        crm.audit(sid, "consent_withdraw", lead_id)

    elif action == "anonymize":
        if g.staff["role"] != "admin":
            abort(403)
        crm.anonymize_lead(lead_id, sid)
        flash("개인정보를 삭제(비식별화)했습니다.", "ok")
        return redirect(url_for("admin.lead_detail", lead_id=lead_id))

    db.commit()
    return redirect(url_for("admin.lead_detail", lead_id=lead_id))


# ---------------------------------------------------------------------------
# 병원 DB
# ---------------------------------------------------------------------------

@bp.route("/hospitals")
@login_required()
def hospitals():
    db = get_db()
    rows = db.execute(
        """SELECT h.*, (SELECT COUNT(*) FROM referrals r WHERE r.hospital_id = h.id) AS referral_count,
                  (SELECT COUNT(*) FROM page_events e WHERE e.hospital_id = h.id AND e.event = 'view_hospital') AS views
           FROM hospitals h ORDER BY h.verification_status, h.name_ko""").fetchall()
    today = now_iso()[:10]
    return render_template("admin/hospitals.html", hospitals=rows, today=today)


def _lines(text, n):
    """'a | b | c' 형식의 여러 줄 입력을 n개 필드 튜플 목록으로 변환."""
    out = []
    for line in (text or "").splitlines():
        if not line.strip():
            continue
        parts = [p.strip() for p in line.split("|")]
        out.append((parts + [""] * n)[:n])
    return out


def _int_krw(v):
    v = re.sub(r"[^\d]", "", v or "")
    return int(v) if v else None


@bp.route("/hospitals/new", methods=["GET", "POST"])
@bp.route("/hospitals/<int:hid>/edit", methods=["GET", "POST"])
@login_required("admin", "coordinator")
def hospital_edit(hid=None):
    db = get_db()
    row = db.execute("SELECT * FROM hospitals WHERE id = ?", (hid,)).fetchone() if hid else None
    if hid and row is None:
        abort(404)
    if request.method == "POST":
        if g.staff["role"] != "admin":
            abort(403)
        f = request.form
        ts = now_iso()
        slug = re.sub(r"[^a-z0-9-]", "-", (f.get("slug") or f.get("name_en", "")).lower()).strip("-")[:80]
        summary = json.dumps({"en": f.get("summary_en", ""), "ko": f.get("summary_ko", ""),
                              "zh": f.get("summary_zh", ""), "ja": f.get("summary_ja", ""),
                              "vi": f.get("summary_vi", "")}, ensure_ascii=False)
        fields = {
            "slug": slug, "name_ko": f.get("name_ko", "").strip(), "name_en": f.get("name_en", "").strip(),
            "hospital_type": f.get("hospital_type", "clinic"), "region": f.get("region", "seoul"),
            "district": f.get("district", ""), "address_ko": f.get("address_ko", ""), "address_en": f.get("address_en", ""),
            "lat": f.get("lat", type=float), "lng": f.get("lng", type=float),
            "specialties": ",".join(s for s in f.getlist("specialties") if s in i18n.SPECIALTIES),
            "languages": ",".join(x.strip() for x in f.get("languages", "").split(",") if x.strip()),
            "price_band": f.get("price_band", 2, type=int),
            "foreign_patient_registered": int(f.get("foreign_patient_registered") == "1"),
            "intl_coordinator": int(f.get("intl_coordinator") == "1"),
            "phone": f.get("phone", ""), "website": f.get("website", ""), "summary_i18n": summary,
            "is_sponsored": int(f.get("is_sponsored") == "1"),
            "verification_status": f.get("verification_status", "planned"),
            "last_verified_at": f.get("last_verified_at") or None, "next_review_at": f.get("next_review_at") or None,
            "is_demo": int(f.get("is_demo") == "1"), "updated_at": ts,
        }
        if not fields["name_ko"] or not fields["name_en"] or not slug:
            flash("병원명(국문·영문)은 필수입니다.", "error")
            return render_template("admin/hospital_edit.html", h=dict(fields, id=hid), tags_text=f.get("tags", ""),
                                   prices_text=f.get("prices", ""), doctors_text=f.get("doctors", ""), summary=json.loads(summary))
        try:
            if hid:
                db.execute(f"UPDATE hospitals SET {', '.join(k + ' = ?' for k in fields)} WHERE id = ?",
                           (*fields.values(), hid))
            else:
                fields["created_at"] = ts
                cur = db.execute(f"INSERT INTO hospitals ({', '.join(fields)}) VALUES ({','.join('?' * len(fields))})",
                                 tuple(fields.values()))
                hid = cur.lastrowid
        except Exception as exc:  # slug 중복 등
            db.rollback()
            flash(f"저장 실패: {exc}", "error")
            return redirect(request.path)

        verified = fields["last_verified_at"]
        db.execute("DELETE FROM hospital_tags WHERE hospital_id = ?", (hid,))
        for tag, evidence, source in _lines(f.get("tags"), 3):
            if tag in i18n.TAGS:
                db.execute("INSERT INTO hospital_tags (hospital_id, tag, evidence, source_url, verified_at) VALUES (?,?,?,?,?)",
                           (hid, tag, evidence, source, verified))
        db.execute("DELETE FROM price_items WHERE hospital_id = ?", (hid,))
        for en, ko, dom, fgn, inc, exc_, src in _lines(f.get("prices"), 7):
            if en:
                db.execute("""INSERT INTO price_items (hospital_id, item_name_en, item_name_ko, price_domestic_krw,
                              price_foreign_krw, includes, excludes, source, verified_at) VALUES (?,?,?,?,?,?,?,?,?)""",
                           (hid, en, ko, _int_krw(dom), _int_krw(fgn), inc, exc_, src, verified))
        db.execute("DELETE FROM doctors WHERE hospital_id = ?", (hid,))
        for ko, en, title, spec, cred, src in _lines(f.get("doctors"), 6):
            if ko or en:
                db.execute("""INSERT INTO doctors (hospital_id, name_ko, name_en, title, specialty, credentials, source_url, verified_at)
                              VALUES (?,?,?,?,?,?,?,?)""", (hid, ko, en, title, spec, cred, src, verified))
        crm.audit(g.staff["id"], "hospital_save", str(hid))
        db.commit()
        flash("저장했습니다.", "ok")
        return redirect(url_for("admin.hospital_edit", hid=hid))

    tags_text = prices_text = doctors_text = ""
    summary = {}
    if row:
        summary = json.loads(row["summary_i18n"] or "{}")
        tags_text = "\n".join(f"{t['tag']} | {t['evidence']} | {t['source_url']}" for t in
                              db.execute("SELECT * FROM hospital_tags WHERE hospital_id = ?", (hid,)))
        prices_text = "\n".join(
            f"{p['item_name_en']} | {p['item_name_ko']} | {p['price_domestic_krw'] or ''} | {p['price_foreign_krw'] or ''} | {p['includes']} | {p['excludes']} | {p['source']}"
            for p in db.execute("SELECT * FROM price_items WHERE hospital_id = ?", (hid,)))
        doctors_text = "\n".join(
            f"{d['name_ko']} | {d['name_en']} | {d['title']} | {d['specialty']} | {d['credentials']} | {d['source_url']}"
            for d in db.execute("SELECT * FROM doctors WHERE hospital_id = ?", (hid,)))
    return render_template("admin/hospital_edit.html", h=row, tags_text=tags_text, prices_text=prices_text,
                           doctors_text=doctors_text, summary=summary)


# ---------------------------------------------------------------------------
# 제휴사 / 직원 / 감사 로그
# ---------------------------------------------------------------------------

@bp.route("/partners", methods=["GET", "POST"])
@login_required()
def partners():
    db = get_db()
    if request.method == "POST":
        if g.staff["role"] != "admin":
            abort(403)
        f = request.form
        if f.get("action") == "toggle":
            db.execute("UPDATE partners SET active = 1 - active WHERE id = ?", (f.get("id", type=int),))
        elif f.get("name") and f.get("partner_type") in i18n.PARTNER_TYPES:
            desc = json.dumps({"en": f.get("desc_en", ""), "ko": f.get("desc_ko", "")}, ensure_ascii=False)
            db.execute("""INSERT INTO partners (partner_type, name, region, lat, lng, description_i18n, price_band, languages, website)
                          VALUES (?,?,?,?,?,?,?,?,?)""",
                       (f["partner_type"], f["name"], f.get("region", "seoul"), f.get("lat", type=float),
                        f.get("lng", type=float), desc, f.get("price_band", 2, type=int), f.get("languages", ""),
                        f.get("website", "")))
        crm.audit(g.staff["id"], "partner_save")
        db.commit()
        return redirect(url_for("admin.partners"))
    rows = db.execute(
        """SELECT p.*, (SELECT COUNT(*) FROM service_orders o WHERE o.partner_id = p.id) AS orders
           FROM partners p ORDER BY p.partner_type, p.name""").fetchall()
    open_orders = db.execute(
        """SELECT o.*, l.full_name, l.preferred_visit_date FROM service_orders o JOIN leads l ON l.id = o.lead_id
           WHERE o.status IN ('requested','in_progress') AND l.anonymized_at IS NULL ORDER BY o.created_at DESC LIMIT 100""").fetchall()
    return render_template("admin/partners.html", partners=rows, open_orders=open_orders)


@bp.route("/staff", methods=["GET", "POST"])
@login_required("admin")
def staff_admin():
    db = get_db()
    if request.method == "POST":
        f = request.form
        if f.get("action") == "toggle":
            target = f.get("id", type=int)
            if target != g.staff["id"]:
                db.execute("UPDATE staff SET active = 1 - active WHERE id = ?", (target,))
        elif f.get("action") == "create":
            if f.get("role") not in dict(ROLES) or len(f.get("password", "")) < 8 or not f.get("username"):
                flash("아이디, 8자 이상 비밀번호, 역할을 확인하세요.", "error")
                return redirect(url_for("admin.staff_admin"))
            try:
                db.execute("INSERT INTO staff (username, password_hash, name, role, languages, email, created_at) VALUES (?,?,?,?,?,?,?)",
                           (f["username"], generate_password_hash(f["password"]), f.get("name") or f["username"],
                            f["role"], f.get("languages", ""), f.get("email", "").strip(), now_iso()))
            except Exception:
                flash("이미 존재하는 아이디입니다.", "error")
                return redirect(url_for("admin.staff_admin"))
        elif f.get("action") == "email":
            email = f.get("email", "").strip()
            if email and not crm.EMAIL_RE.match(email):
                flash("이메일 주소 형식을 확인하세요.", "error")
                return redirect(url_for("admin.staff_admin"))
            db.execute("UPDATE staff SET email = ?, notify_email = ? WHERE id = ?",
                       (email, int(f.get("notify_email") == "1"), f.get("id", type=int)))
        elif f.get("action") == "test_mail":
            ok = notify.send("test", f.get("to", "").strip(), f"[{current_app.config['BRAND_NAME']}] 메일 발송 테스트",
                             f"발송 방식: {notify.backend()}\n이 메일이 보이면 알림 설정이 정상입니다.")
            flash("테스트 메일을 보냈습니다. 받은편지함(스팸함 포함)을 확인하세요." if ok else "발송 실패 — 아래 발송 기록의 오류를 확인하세요.",
                  "ok" if ok else "error")
        elif f.get("action") == "password":
            if len(f.get("password", "")) < 8:
                flash("비밀번호는 8자 이상이어야 합니다.", "error")
                return redirect(url_for("admin.staff_admin"))
            db.execute("UPDATE staff SET password_hash = ? WHERE id = ?",
                       (generate_password_hash(f["password"]), f.get("id", type=int)))
        crm.audit(g.staff["id"], f"staff_{f.get('action')}", f.get("username") or f.get("id", ""))
        db.commit()
        return redirect(url_for("admin.staff_admin"))
    rows = db.execute(
        f"""SELECT s.*, (SELECT COUNT(*) FROM leads l WHERE l.assigned_staff_id = s.id
               AND l.lead_status IN ({','.join('?' * len(crm.OPEN_STATUSES))})) AS open_count
            FROM staff s ORDER BY s.role, s.name""", crm.OPEN_STATUSES).fetchall()
    mail_logs = db.execute("SELECT * FROM email_logs ORDER BY id DESC LIMIT 50").fetchall()
    return render_template("admin/staff.html", staff_rows=rows, mail_logs=mail_logs, mail_backend=notify.backend(),
                           ops_recipients=notify.ops_recipients(), mail_from=notify.mail_from())


@bp.route("/audit")
@login_required("admin")
def audit_logs():
    rows = get_db().execute(
        """SELECT a.*, s.name FROM audit_logs a LEFT JOIN staff s ON s.id = a.staff_id
           ORDER BY a.id DESC LIMIT 500""").fetchall()
    return render_template("admin/audit.html", rows=rows)
