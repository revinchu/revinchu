"""이메일 알림.

발송 방식 (환경 변수)
    MAIL_BACKEND      smtp | resend | console  (미설정 시 RESEND_API_KEY → resend, SMTP_HOST → smtp, 그 외 console)
    MAIL_FROM         보내는 주소 (예: "브랜드명 <no-reply@내도메인.com>")
    NOTIFY_EMAIL      운영 알림 수신 주소, 쉼표로 여러 개 (예: ops@example.com)
    SITE_URL          메일 속 관리자 링크의 기준 주소 (예: https://내도메인.com)
    SMTP_HOST / SMTP_PORT(587) / SMTP_USER / SMTP_PASSWORD / SMTP_SSL(0|1) / SMTP_STARTTLS(1|0)
    RESEND_API_KEY

알림 종류
    new_lead     신규 상담 신청 → 운영 주소 + 배정 담당자
    patient_ack  신청자에게 접수 확인 (신청자 언어)
    assign       담당자 변경 → 새 담당자
    sla          30분 넘게 미응대인 신규 문의 → 운영 주소 + 담당자 (cron: python -m medical_platform.notify sla)

개인정보 최소화: 직원용 메일에는 연락처를 넣지 않고 관리자 화면 링크만 보낸다.
"""
import json
import os
import smtplib
import sys
import urllib.error
import urllib.request
from datetime import datetime, timedelta, timezone
from email.message import EmailMessage
from email.utils import parseaddr

from flask import current_app, has_request_context, request

from . import crm, i18n
from .db import get_db, now_iso


def backend():
    name = os.environ.get("MAIL_BACKEND", "").lower()
    if name:
        return name
    if os.environ.get("RESEND_API_KEY"):
        return "resend"
    if os.environ.get("SMTP_HOST"):
        return "smtp"
    return "console"


def ops_recipients():
    return [a.strip() for a in os.environ.get("NOTIFY_EMAIL", "").split(",") if a.strip()]


def mail_from():
    return os.environ.get("MAIL_FROM") or os.environ.get("SMTP_USER") or "no-reply@localhost"


def site_url():
    url = os.environ.get("SITE_URL")
    if not url and has_request_context():
        url = request.url_root
    return (url or "http://127.0.0.1:5000").rstrip("/")


def _deliver(to, subject, text, html=None):
    kind = backend()
    if kind == "console":
        current_app.logger.info("[mail:console] to=%s subject=%s\n%s", to, subject, text)
        return
    if kind == "resend":
        payload = {"from": mail_from(), "to": [to], "subject": subject, "text": text}
        if html:
            payload["html"] = html
        headers = {"Authorization": f"Bearer {os.environ['RESEND_API_KEY']}", "Content-Type": "application/json"}
        if sys.platform == "emscripten":  # Cloudflare Python Workers: 소켓 대신 fetch 사용
            import js
            from pyodide.ffi import run_sync, to_js
            init = to_js({"method": "POST", "headers": headers, "body": json.dumps(payload)},
                         dict_converter=js.Object.fromEntries)
            resp = run_sync(js.fetch("https://api.resend.com/emails", init))
            if not resp.ok:
                raise RuntimeError(f"Resend HTTP {resp.status}: {run_sync(resp.text())[:300]}")
            return
        req = urllib.request.Request("https://api.resend.com/emails", data=json.dumps(payload).encode(),
                                     method="POST", headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=10):
                pass
        except urllib.error.HTTPError as exc:
            raise RuntimeError(f"Resend HTTP {exc.code}: {exc.read()[:300].decode(errors='replace')}") from exc
        return
    if kind == "smtp" and sys.platform == "emscripten":
        raise RuntimeError("Cloudflare Workers에서는 SMTP를 쓸 수 없습니다. RESEND_API_KEY 를 설정하세요.")
    if kind == "smtp":
        msg = EmailMessage()
        msg["From"], msg["To"], msg["Subject"] = mail_from(), to, subject
        msg.set_content(text)
        if html:
            msg.add_alternative(html, subtype="html")
        host, port = os.environ["SMTP_HOST"], int(os.environ.get("SMTP_PORT", "587"))
        use_ssl = os.environ.get("SMTP_SSL") == "1" or port == 465
        with (smtplib.SMTP_SSL(host, port, timeout=10) if use_ssl else smtplib.SMTP(host, port, timeout=10)) as s:
            if not use_ssl and os.environ.get("SMTP_STARTTLS", "1") != "0":
                s.starttls()
            if os.environ.get("SMTP_USER"):
                s.login(os.environ["SMTP_USER"], os.environ.get("SMTP_PASSWORD", ""))
            s.send_message(msg)
        return
    raise RuntimeError(f"unknown MAIL_BACKEND: {kind}")


def send(event, to, subject, text, html=None, lead_id=None):
    """메일 1통을 보내고 결과를 email_logs 에 남긴다. 실패해도 예외를 올리지 않는다."""
    if not to or "@" not in parseaddr(to)[1]:
        return False
    status, error = "sent", ""
    try:
        _deliver(to, subject, text, html)
        if backend() == "console":
            status = "logged"
    except Exception as exc:  # 메일 장애가 상담 접수·CRM 작업을 막지 않도록 한다
        status, error = "failed", str(exc)[:500]
        current_app.logger.warning("mail send failed (%s → %s): %s", event, to, exc)
    get_db().execute(
        "INSERT INTO email_logs (event, recipient, subject, lead_id, status, error, created_at) VALUES (?,?,?,?,?,?,?)",
        (event, to, subject[:200], lead_id, status, error, now_iso()))
    get_db().commit()
    return status != "failed"


def _lead(lead_id):
    return get_db().execute(
        """SELECT l.*, s.name AS staff_name, s.email AS staff_email, s.notify_email AS staff_notify,
                  h.name_ko AS hospital_name
           FROM leads l LEFT JOIN staff s ON s.id = l.assigned_staff_id
           LEFT JOIN hospitals h ON h.id = l.preferred_hospital_id WHERE l.id = ?""", (lead_id,)).fetchone()


def _ko(code, prefix):
    return i18n.translate(f"{prefix}_{code}", "ko") if code else "-"


def _lead_summary(lead):
    needs = ", ".join(_ko(n, "need") for n in lead["concierge_needs"].split(",") if n) or "-"
    prefs = ", ".join(_ko(p, "pref") for p in lead["hospital_preferences"].split(",") if p) or "-"
    trip = " · ".join(x for x in [lead["preferred_visit_date"],
                                  f"{lead['stay_nights']}박" if lead["stay_nights"] else "",
                                  f"{lead['travel_party_size']}명" if lead["travel_party_size"] else ""] if x) or "-"
    rows = [
        ("접수번호", lead["id"][:8].upper()),
        ("상담 언어", i18n.LANGUAGES.get(lead["preferred_language"], lead["preferred_language"])),
        ("거주 국가", lead["country_of_residence"] or "-"),
        ("희망 진료", _ko(lead["treatment_category"], "sp")),
        ("희망 병원", lead["hospital_name"] or "추천 요청"),
        ("병원 선호", prefs),
        ("예산", _ko(lead["budget_band"], "bud")),
        ("방문 일정", trip),
        ("컨시어지", needs),
        ("VIP 의전", "요청" if lead["vip_request"] else "-"),
        ("유입", lead["utm_source"] or _ko(lead["source_channel"], "heard")),
        ("담당자", lead["staff_name"] or "미배정"),
    ]
    if lead["duplicate_of"]:
        rows.append(("중복 의심", lead["duplicate_of"][:8].upper()))
    return rows


def _staff_mail(title, intro, rows, link, link_label="관리자 화면에서 열기"):
    brand = current_app.config["BRAND_NAME"]
    text = f"{intro}\n\n" + "\n".join(f"{k}: {v}" for k, v in rows) + f"\n\n{link_label}: {link}\n\n— {brand} 알림 (연락처 등 개인정보는 관리자 화면에서만 확인하세요)"
    esc = lambda v: (str(v).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))
    trs = "".join(f'<tr><td style="padding:6px 16px 6px 0;color:#7a807c;white-space:nowrap">{esc(k)}</td>'
                  f'<td style="padding:6px 0">{esc(v)}</td></tr>' for k, v in rows)
    html = f"""<div style="font-family:-apple-system,'Apple SD Gothic Neo','Noto Sans KR',sans-serif;color:#1c1f1e;max-width:560px">
<p style="font-size:13px;color:#7a807c;margin:0 0 4px">{esc(brand)}</p>
<h2 style="font-size:20px;margin:0 0 12px">{esc(title)}</h2>
<p style="margin:0 0 16px">{esc(intro)}</p>
<table style="border-collapse:collapse;font-size:14px;border-top:2px solid #1c1f1e;width:100%">{trs}</table>
<p style="margin:20px 0"><a href="{esc(link)}" style="background:#2f5f50;color:#fff;padding:10px 16px;text-decoration:none;font-weight:600">{esc(link_label)}</a></p>
<p style="font-size:12px;color:#7a807c">연락처 등 개인정보는 관리자 화면에서만 확인하세요.</p></div>"""
    return text, html


def notify_new_lead(lead_id):
    lead = _lead(lead_id)
    if lead is None:
        return
    ref = lead["id"][:8].upper()
    flags = (" · VIP" if lead["vip_request"] else "") + (" · 중복 의심" if lead["duplicate_of"] else "")
    subject = (f"[{current_app.config['BRAND_NAME']}] 신규 문의 #{ref} · "
               f"{i18n.LANGUAGES.get(lead['preferred_language'], '')} · {_ko(lead['treatment_category'], 'sp')}{flags}")
    link = f"{site_url()}/admin/leads/{lead_id}"
    text, html = _staff_mail("신규 상담 신청", f"{crm.FIRST_RESPONSE_SLA_MIN}분 안에 최초 응대해 주세요.", _lead_summary(lead), link)
    recipients = set(ops_recipients())
    if lead["staff_email"] and lead["staff_notify"]:
        recipients.add(lead["staff_email"])
    for to in sorted(recipients):
        send("new_lead", to, subject, text, html, lead_id)


def send_patient_ack(lead_id):
    lead = _lead(lead_id)
    if lead is None or not lead["email"]:
        return
    lang = lead["preferred_language"] if lead["preferred_language"] in i18n.LANGUAGES else i18n.DEFAULT_LANG
    brand = current_app.config["BRAND_NAME"]
    ref = lead["id"][:8].upper()
    fmt = {"name": lead["full_name"], "ref": ref, "brand": brand, "min": crm.FIRST_RESPONSE_SLA_MIN}
    subject = i18n.translate("mail_ack_subject", lang).format(**fmt)
    body = i18n.translate("mail_ack_body", lang).format(**fmt)
    text = body + "\n\n" + i18n.translate("emergency_note", lang) + f"\n\n{brand} · {site_url()}/{lang}/"
    esc = lambda v: v.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    paras = "".join(f'<p style="margin:0 0 12px">{esc(p)}</p>' for p in body.split("\n\n"))
    html = f"""<div style="font-family:-apple-system,'Apple SD Gothic Neo','Noto Sans KR','Noto Sans SC','Noto Sans JP',sans-serif;color:#1c1f1e;max-width:560px">
<p style="font-family:Georgia,serif;font-size:20px;font-weight:600;margin:0 0 16px">{esc(brand)}</p>
{paras}<p style="font-family:Menlo,monospace;font-size:22px;letter-spacing:.12em;border:2px solid #1c1f1e;display:inline-block;padding:8px 14px">{ref}</p>
<p style="font-size:12px;color:#7a807c;margin-top:20px">{esc(i18n.translate('emergency_note', lang))}</p></div>"""
    send("patient_ack", lead["email"], subject, text, html, lead_id)


def notify_assignment(lead_id, staff_id):
    staff = get_db().execute("SELECT * FROM staff WHERE id = ?", (staff_id,)).fetchone()
    lead = _lead(lead_id)
    if not staff or not lead or not staff["email"] or not staff["notify_email"]:
        return
    ref = lead["id"][:8].upper()
    subject = f"[{current_app.config['BRAND_NAME']}] 담당 배정 #{ref} · {_ko(lead['treatment_category'], 'sp')}"
    text, html = _staff_mail("문의가 배정되었습니다", f"{staff['name']}님에게 문의가 배정되었습니다.",
                             _lead_summary(lead), f"{site_url()}/admin/leads/{lead_id}")
    send("assign", staff["email"], subject, text, html, lead_id)


def send_sla_alerts():
    """30분 넘게 최초 응대가 없는 신규 문의를 한 번씩만 알린다. 알린 건수를 반환."""
    db = get_db()
    cut = (datetime.now(timezone.utc) - timedelta(minutes=crm.FIRST_RESPONSE_SLA_MIN)).replace(microsecond=0).isoformat()
    rows = db.execute(
        """SELECT l.id FROM leads l WHERE l.lead_status = 'new' AND l.first_response_at IS NULL
           AND l.created_at < ? AND l.sla_alerted_at IS NULL AND l.anonymized_at IS NULL ORDER BY l.created_at""",
        (cut,)).fetchall()
    leads = [_lead(r["id"]) for r in rows]
    if not leads:
        return 0
    brand = current_app.config["BRAND_NAME"]
    base = site_url()

    def digest(items):
        rows_ = [(f"#{l['id'][:8].upper()}", f"{i18n.LANGUAGES.get(l['preferred_language'], '')} · "
                  f"{_ko(l['treatment_category'], 'sp')} · 담당 {l['staff_name'] or '미배정'} · 접수 {l['created_at'][:16].replace('T', ' ')} UTC")
                 for l in items]
        return _staff_mail(f"미응대 문의 {len(items)}건", f"{crm.FIRST_RESPONSE_SLA_MIN}분 넘게 최초 응대가 없는 문의입니다.",
                           rows_, f"{base}/admin/leads?status=new", "미응대 문의 보기")

    subject = f"[{brand}] SLA 초과 미응대 문의 {len(leads)}건"
    text, html = digest(leads)
    for to in ops_recipients():
        send("sla", to, subject, text, html)
    by_staff = {}
    for l in leads:
        if l["staff_email"] and l["staff_notify"]:
            by_staff.setdefault(l["staff_email"], []).append(l)
    for to, items in by_staff.items():
        if to in ops_recipients():
            continue
        t, h = digest(items)
        send("sla", to, f"[{brand}] 내 담당 미응대 문의 {len(items)}건", t, h)
    ts = now_iso()
    db.executemany("UPDATE leads SET sla_alerted_at = ? WHERE id = ?", [(ts, l["id"]) for l in leads])
    db.commit()
    return len(leads)


def main(argv):
    """cron·수동 실행용: python -m medical_platform.notify sla | test <email>"""
    from .app import create_app
    app = create_app()
    with app.app_context():
        if argv[:1] == ["sla"]:
            print(f"SLA 알림 대상: {send_sla_alerts()}건")
        elif argv[:1] == ["test"] and len(argv) > 1:
            ok = send("test", argv[1], f"[{app.config['BRAND_NAME']}] 메일 발송 테스트", f"발송 방식: {backend()}\n정상적으로 수신되었습니다.")
            print("성공" if ok else "실패 — email_logs 확인")
        else:
            print(__doc__)
            return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
