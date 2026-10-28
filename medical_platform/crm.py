"""리드 생성·자동 배정·활동 기록·감사 로그 등 CRM 도메인 로직."""
import re
import uuid
from datetime import datetime, timedelta, timezone

from flask import request

from .db import get_db, now_iso

# CRM 파이프라인 (사업기획서 5-1)
LEAD_STATUSES = [
    ("new", "신규 문의"),
    ("contacting", "최초 연락"),
    ("qualified", "유효 문의"),
    ("matching", "병원 연결"),
    ("consultation", "상담·견적"),
    ("booked", "예약 확정"),
    ("visited", "방문 완료"),
    ("followup", "사후관리"),
    ("lost", "보류·이탈"),
    ("invalid", "무효"),
]
STATUS_LABELS = dict(LEAD_STATUSES)
OPEN_STATUSES = ("new", "contacting", "qualified", "matching", "consultation", "booked")
FUNNEL = ["new", "contacting", "qualified", "matching", "consultation", "booked", "visited"]
DISQUALIFY_REASONS = ["허위 연락처", "중복 문의", "연락 불가", "단순 정보 탐색", "대상 외 진료", "기타"]

REFERRAL_STATUSES = [
    ("requested", "연결 요청"), ("sent", "병원 전달"), ("replied", "병원 회신"),
    ("quoted", "견적 전달"), ("booked", "예약 확정"), ("visited", "방문 완료"),
    ("cancelled", "취소"), ("declined", "병원 거절"),
]
SERVICE_STATUSES = [("requested", "요청"), ("in_progress", "진행 중"), ("confirmed", "확정"), ("cancelled", "취소")]

PRIORITIES = [("high", "높음"), ("normal", "보통"), ("low", "낮음")]

# 최초 사람 응대 목표 (SLA, 분)
FIRST_RESPONSE_SLA_MIN = 30

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def client_ip():
    forwarded = request.headers.get("CF-Connecting-IP") or request.headers.get("X-Forwarded-For") or request.remote_addr or ""
    return forwarded.split(",")[0].strip()


def audit(staff_id, action, target=""):
    get_db().execute(
        "INSERT INTO audit_logs (staff_id, action, target, ip, created_at) VALUES (?,?,?,?,?)",
        (staff_id, action, target, client_ip(), now_iso()),
    )


def log_activity(lead_id, kind, body, staff_id=None):
    get_db().execute(
        "INSERT INTO lead_activities (lead_id, staff_id, kind, body, created_at) VALUES (?,?,?,?,?)",
        (lead_id, staff_id, kind, body, now_iso()),
    )


def pick_staff(language):
    """선호 언어를 지원하는 코디네이터 중 미처리 리드가 가장 적은 사람에게 배정한다."""
    db = get_db()
    placeholders = ",".join("?" * len(OPEN_STATUSES))
    rows = db.execute(
        f"""SELECT s.id, s.languages,
                   (SELECT COUNT(*) FROM leads l WHERE l.assigned_staff_id = s.id
                      AND l.lead_status IN ({placeholders})) AS open_count
            FROM staff s WHERE s.active = 1 AND s.role IN ('coordinator', 'admin')
            ORDER BY (s.role = 'admin'), open_count, s.id""",
        OPEN_STATUSES,
    ).fetchall()
    speakers = [r for r in rows if language in (r["languages"] or "").split(",")]
    pool = speakers or rows
    return pool[0]["id"] if pool else None


def find_duplicate(email, messenger_id, days=30):
    since = (datetime.now(timezone.utc) - timedelta(days=days)).replace(microsecond=0).isoformat()
    db = get_db()
    if email:
        row = db.execute(
            "SELECT id FROM leads WHERE lower(email) = lower(?) AND created_at >= ? AND anonymized_at IS NULL ORDER BY created_at LIMIT 1",
            (email, since)).fetchone()
        if row:
            return row["id"]
    if messenger_id:
        row = db.execute(
            "SELECT id FROM leads WHERE messenger_id = ? AND created_at >= ? AND anonymized_at IS NULL ORDER BY created_at LIMIT 1",
            (messenger_id, since)).fetchone()
        if row:
            return row["id"]
    return None


def create_lead(data, consent, attribution):
    """검증이 끝난 폼 데이터로 리드·동의 기록·서비스 요청을 생성하고 lead_id 를 반환한다."""
    db = get_db()
    ts = now_iso()
    lead_id = str(uuid.uuid4())
    duplicate_of = find_duplicate(data["email"], data["messenger_id"])
    staff_id = pick_staff(data["preferred_language"])
    priority = "high" if data["vip_request"] or data["budget_band"] in ("10k_30k", "over_30k") else "normal"

    db.execute(
        """INSERT INTO leads (id, created_at, updated_at, full_name, email, phone, messenger_type, messenger_id,
           contact_channel, nationality, country_of_residence, preferred_language, age_band, treatment_category,
           treatment_interest, hospital_preferences, preferred_hospital_id, budget_band, preferred_visit_date,
           stay_nights, travel_party_size, concierge_needs, accommodation_pref, food_pref, itinerary_interests,
           vip_request, message, source_channel, utm_source, utm_medium, utm_campaign, utm_content, utm_term,
           landing_page, referrer, site_language, lead_status, priority, assigned_staff_id, duplicate_of)
           VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
        (lead_id, ts, ts, data["full_name"], data["email"], data["phone"], data["contact_channel"],
         data["messenger_id"], data["contact_channel"], data["nationality"], data["country_of_residence"],
         data["preferred_language"], data["age_band"], data["treatment_category"], data["treatment_interest"],
         ",".join(data["hospital_preferences"]), data["preferred_hospital_id"], data["budget_band"],
         data["preferred_visit_date"], data["stay_nights"], data["travel_party_size"],
         ",".join(data["concierge_needs"]), data["accommodation_pref"], data["food_pref"],
         data["itinerary_interests"], 1 if data["vip_request"] else 0, data["message"], data["source_channel"],
         attribution.get("utm_source", ""), attribution.get("utm_medium", ""), attribution.get("utm_campaign", ""),
         attribution.get("utm_content", ""), attribution.get("utm_term", ""), attribution.get("landing_page", ""),
         attribution.get("referrer", ""), data["site_language"], "new", priority, staff_id, duplicate_of),
    )
    cur = db.execute(
        """INSERT INTO consents (lead_id, privacy_required, third_party_share, overseas_transfer, marketing,
           policy_version, ip, user_agent, created_at) VALUES (?,?,?,?,?,?,?,?,?)""",
        (lead_id, 1, int(consent["third_party_share"]), int(consent["overseas_transfer"]), int(consent["marketing"]),
         consent["policy_version"], client_ip(), (request.user_agent.string or "")[:300], ts),
    )
    db.execute("UPDATE leads SET consent_record_id = ? WHERE id = ?", (cur.lastrowid, lead_id))

    needs = list(data["concierge_needs"])
    if data["vip_request"] and "vip" not in needs:
        needs.append("vip")
    for need in needs:
        db.execute(
            "INSERT INTO service_orders (lead_id, service_type, created_at, updated_at) VALUES (?,?,?,?)",
            (lead_id, need, ts, ts))
    if data["preferred_hospital_id"]:
        # 병원 전달은 담당자가 동의 확인 후 진행하므로, 희망 병원은 '연결 요청' 상태로만 기록한다.
        db.execute(
            "INSERT INTO referrals (lead_id, hospital_id, status, created_at, updated_at) VALUES (?,?,?,?,?)",
            (lead_id, data["preferred_hospital_id"], "requested", ts, ts))

    log_activity(lead_id, "system", "웹 상담 신청 접수" + (f" (중복 의심: {duplicate_of[:8]})" if duplicate_of else ""))
    if staff_id:
        log_activity(lead_id, "assign", f"자동 배정 (언어: {data['preferred_language']})")
    db.execute(
        "INSERT INTO page_events (event, path, lang, utm_source, utm_campaign, created_at) VALUES (?,?,?,?,?,?)",
        ("lead_submit", request.path, data["site_language"], attribution.get("utm_source", ""),
         attribution.get("utm_campaign", ""), ts))
    db.commit()
    return lead_id


def anonymize_lead(lead_id, staff_id):
    """삭제 요청 처리: 식별정보를 지우고 통계용 비식별 필드만 남긴다."""
    db = get_db()
    ts = now_iso()
    db.execute(
        """UPDATE leads SET full_name = '(삭제됨)', email = '', phone = '', messenger_id = '', message = '',
           treatment_interest = '', accommodation_pref = '', food_pref = '', itinerary_interests = '',
           referrer = '', anonymized_at = ?, updated_at = ? WHERE id = ?""",
        (ts, ts, lead_id))
    db.execute("UPDATE lead_activities SET body = '(삭제됨)' WHERE lead_id = ? AND kind IN ('note', 'contact')", (lead_id,))
    db.execute("UPDATE consents SET withdrawn_at = COALESCE(withdrawn_at, ?), ip = '', user_agent = '' WHERE lead_id = ?",
               (ts, lead_id))
    log_activity(lead_id, "system", "개인정보 삭제(비식별화) 처리", staff_id)
    audit(staff_id, "lead_anonymize", lead_id)
    db.commit()


def mask(value, keep=2):
    if not value:
        return ""
    if "@" in value:
        name, _, domain = value.partition("@")
        return name[:keep] + "***@" + domain
    return value[:keep] + "*" * max(len(value) - keep, 3)
