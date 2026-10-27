import re

import pytest

from medical_platform.app import create_app
from medical_platform.db import get_db


@pytest.fixture
def app(tmp_path):
    return create_app({"TESTING": True, "DATABASE": str(tmp_path / "test.db"), "SECRET_KEY": "test"})


@pytest.fixture
def client(app):
    return app.test_client()


def csrf(client, url):
    html = client.get(url).get_data(as_text=True)
    return re.search(r'name="_csrf" value="([^"]+)"', html).group(1)


def login(client, username="admin", password="admin1234"):
    token = csrf(client, "/admin/login")
    return client.post("/admin/login", data={"_csrf": token, "username": username, "password": password})


def submit_lead(client, lang="en", **overrides):
    token = csrf(client, f"/{lang}/consult")
    data = {
        "_csrf": token, "full_name": "Jane Doe", "email": "jane@example.com", "contact_channel": "email",
        "preferred_language": lang, "treatment_category": "dermatology", "consent_privacy": "1",
        "concierge_needs": ["hotel", "kfood"], "hospital_preferences": ["value"], "budget_band": "1k_3k",
    }
    data.update(overrides)
    return client.post(f"/{lang}/consult", data=data)


@pytest.mark.parametrize("lang", ["en", "ko", "zh", "ja", "vi"])
def test_public_pages_render_in_all_languages(client, lang):
    for path in ["/", "/hospitals", "/hospitals/demo-gangnam-skin", "/concierge", "/how-it-works", "/consult", "/privacy"]:
        resp = client.get(f"/{lang}{path}")
        assert resp.status_code == 200, path


def test_root_redirects_by_accept_language(client):
    resp = client.get("/", headers={"Accept-Language": "ja,en;q=0.5"})
    assert resp.status_code == 302 and resp.headers["Location"].startswith("/ja/")


def test_unknown_language_404(client):
    assert client.get("/xx/").status_code == 404


def test_hospital_filters(client):
    html = client.get("/en/hospitals?tag=vip").get_data(as_text=True)
    assert "Sinsa Plastic Surgery" in html and "Seocho Dental" not in html
    html = client.get("/en/hospitals?specialty=dental").get_data(as_text=True)
    assert "Seocho Dental" in html and "Sinsa Plastic" not in html
    html = client.get("/ko/hospitals?q=피부과").get_data(as_text=True)
    assert "강남 스킨" in html


def test_lead_submission_creates_lead_consent_and_services(client, app):
    client.get("/en/?utm_source=google&utm_campaign=jp_derm")
    resp = submit_lead(client, lang="ja", vip_request="1")
    assert resp.status_code == 302 and "/consult/thanks/" in resp.headers["Location"]
    with app.app_context():
        db = get_db()
        lead = db.execute("SELECT * FROM leads").fetchone()
        assert lead["utm_source"] == "google" and lead["utm_campaign"] == "jp_derm"
        assert lead["priority"] == "high"
        staff = db.execute("SELECT * FROM staff WHERE id = ?", (lead["assigned_staff_id"],)).fetchone()
        assert "ja" in staff["languages"]
        consent = db.execute("SELECT * FROM consents WHERE id = ?", (lead["consent_record_id"],)).fetchone()
        assert consent["privacy_required"] == 1 and consent["third_party_share"] == 0
        services = {r["service_type"] for r in db.execute("SELECT service_type FROM service_orders")}
        assert services == {"hotel", "kfood", "vip"}


def test_lead_requires_consent_and_contact(client, app):
    resp = submit_lead(client, consent_privacy="", email="")
    assert resp.status_code == 400
    with app.app_context():
        assert get_db().execute("SELECT COUNT(*) FROM leads").fetchone()[0] == 0


def test_honeypot_drops_bots(client, app):
    submit_lead(client, website="http://spam")
    with app.app_context():
        assert get_db().execute("SELECT COUNT(*) FROM leads").fetchone()[0] == 0


def test_duplicate_lead_flagged(client, app):
    submit_lead(client)
    submit_lead(client)
    with app.app_context():
        rows = get_db().execute("SELECT duplicate_of FROM leads ORDER BY created_at, rowid").fetchall()
        assert rows[0]["duplicate_of"] is None and rows[1]["duplicate_of"] is not None


def test_admin_requires_login(client):
    assert client.get("/admin/").status_code == 302


def test_admin_crm_flow(client, app):
    submit_lead(client, consent_third="1")
    login(client)
    assert client.get("/admin/").status_code == 200
    html = client.get("/admin/leads?status=open").get_data(as_text=True)
    assert "Jane Doe" in html
    with app.app_context():
        lead_id = get_db().execute("SELECT id FROM leads").fetchone()["id"]
    page = client.get(f"/admin/leads/{lead_id}")
    assert page.status_code == 200
    token = re.search(r'name="_csrf" value="([^"]+)"', page.get_data(as_text=True)).group(1)
    url = f"/admin/leads/{lead_id}/action"
    client.post(url, data={"_csrf": token, "action": "note", "body": "Called via email", "is_contact": "1"})
    client.post(url, data={"_csrf": token, "action": "add_referral", "hospital_id": "1"})
    with app.app_context():
        db = get_db()
        lead = db.execute("SELECT * FROM leads WHERE id = ?", (lead_id,)).fetchone()
        assert lead["first_response_at"] and lead["lead_status"] == "matching"
        ref = db.execute("SELECT id FROM referrals WHERE lead_id = ?", (lead_id,)).fetchone()
    client.post(url, data={"_csrf": token, "action": "update_referral", "referral_id": ref["id"], "status": "booked"})
    with app.app_context():
        assert get_db().execute("SELECT lead_status FROM leads WHERE id = ?", (lead_id,)).fetchone()[0] == "booked"
    export = client.get("/admin/leads/export.csv")
    assert export.status_code == 200 and "Jane Doe" in export.get_data(as_text=True)


def test_referral_blocked_without_third_party_consent(client, app):
    submit_lead(client)
    login(client)
    with app.app_context():
        lead_id = get_db().execute("SELECT id FROM leads").fetchone()["id"]
    token = csrf(client, f"/admin/leads/{lead_id}")
    client.post(f"/admin/leads/{lead_id}/action", data={"_csrf": token, "action": "add_referral", "hospital_id": "1"})
    with app.app_context():
        assert get_db().execute("SELECT COUNT(*) FROM referrals").fetchone()[0] == 0


def test_marketer_cannot_see_pii(client, app):
    submit_lead(client)
    login(client, "marketer", "market1234")
    html = client.get("/admin/leads").get_data(as_text=True)
    assert "Jane Doe" not in html and "jane@example.com" not in html
    with app.app_context():
        lead_id = get_db().execute("SELECT id FROM leads").fetchone()["id"]
    assert client.get(f"/admin/leads/{lead_id}").status_code == 403
    assert client.get("/admin/leads/export.csv").status_code == 403


def test_anonymize_removes_pii(client, app):
    submit_lead(client)
    login(client)
    with app.app_context():
        lead_id = get_db().execute("SELECT id FROM leads").fetchone()["id"]
    token = csrf(client, f"/admin/leads/{lead_id}")
    client.post(f"/admin/leads/{lead_id}/action", data={"_csrf": token, "action": "anonymize"})
    with app.app_context():
        lead = get_db().execute("SELECT * FROM leads WHERE id = ?", (lead_id,)).fetchone()
        assert lead["email"] == "" and lead["anonymized_at"]


def test_admin_hospital_create(client, app):
    login(client)
    token = csrf(client, "/admin/hospitals/new")
    resp = client.post("/admin/hospitals/new", data={
        "_csrf": token, "name_ko": "테스트의원", "name_en": "Test Clinic", "region": "seoul",
        "specialties": ["dental"], "verification_status": "published", "lat": "37.5", "lng": "127.0",
        "tags": "value | 공개 가격 비교 | https://example.com", "prices": "Scaling | 스케일링 | 50,000 | 50000 | | | site",
    })
    assert resp.status_code == 302
    html = client.get("/en/hospitals/test-clinic").get_data(as_text=True)
    assert "Test Clinic" in html and "₩50,000" in html


def test_post_without_csrf_rejected(client):
    assert client.post("/en/consult", data={"full_name": "x"}).status_code == 400
