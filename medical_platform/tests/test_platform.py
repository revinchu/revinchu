import re

import pytest

from medical_platform.app import create_app
from medical_platform.db import get_db


class FakeD1:
    """Cloudflare D1 바인딩 흉내: prepare/bind/all/batch 를 sqlite 로 처리한다 (D1처럼 자동 커밋)."""

    def __init__(self, path):
        import sqlite3
        self.conn = sqlite3.connect(path, isolation_level=None, check_same_thread=False)
        self.conn.execute("PRAGMA foreign_keys = ON")

    def prepare(self, sql):
        return FakeStmt(self.conn, sql)

    def batch(self, stmts):
        return [st.all() for st in stmts]


class FakeStmt:
    def __init__(self, conn, sql, params=()):
        self.conn, self.sql, self.params = conn, sql, params

    def bind(self, *params):
        return FakeStmt(self.conn, self.sql, params)

    def all(self):
        cur = self.conn.execute(self.sql, self.params)
        cols = [d[0] for d in cur.description or []]
        return {"results": [dict(zip(cols, r)) for r in cur.fetchall()], "meta": {"last_row_id": cur.lastrowid}}


@pytest.fixture(params=["sqlite", "d1"])
def app(request, tmp_path):
    config = {"TESTING": True, "DATABASE": str(tmp_path / "test.db"), "SECRET_KEY": "test"}
    if request.param == "d1":
        config.update(DB_BACKEND="d1", D1_BINDING=FakeD1(str(tmp_path / "d1.db")),
                      D1_OPTIONS={"run_sync": lambda value: value})
    return create_app(config)


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


# ---------------------------------------------------------------------------
# 이메일 알림
# ---------------------------------------------------------------------------

@pytest.fixture
def outbox(monkeypatch):
    from medical_platform import notify
    sent = []
    monkeypatch.setenv("MAIL_BACKEND", "smtp")
    monkeypatch.setenv("NOTIFY_EMAIL", "ops@example.com")
    monkeypatch.setattr(notify, "_deliver", lambda to, subject, text, html=None: sent.append((to, subject, text)))
    return sent


def test_new_lead_sends_ops_alert_and_patient_ack(client, app, outbox):
    submit_lead(client, lang="ja", email="yui@example.jp", preferred_language="ja", vip_request="1")
    recipients = {to for to, _, _ in outbox}
    assert recipients == {"ops@example.com", "yui@example.jp"}
    ops = next(m for m in outbox if m[0] == "ops@example.com")
    assert "VIP" in ops[1] and "/admin/leads/" in ops[2]
    assert "yui@example.jp" not in ops[2] and "Jane Doe" not in ops[2]  # 직원 메일에 연락처·이름 미포함
    ack = next(m for m in outbox if m[0] == "yui@example.jp")
    assert "ご相談を受け付けました" in ack[1]
    with app.app_context():
        assert get_db().execute("SELECT COUNT(*) FROM email_logs WHERE status = 'sent'").fetchone()[0] == 2


def test_assigned_staff_with_email_is_notified(client, app, outbox):
    with app.app_context():
        get_db().execute("UPDATE staff SET email = 'ja@wizx.test' WHERE username = 'coord_ja'")
        get_db().commit()
    submit_lead(client, lang="ja", preferred_language="ja")
    assert "ja@wizx.test" in {to for to, _, _ in outbox}


def test_mail_failure_does_not_block_lead(client, app, monkeypatch):
    from medical_platform import notify
    monkeypatch.setenv("MAIL_BACKEND", "smtp")
    monkeypatch.setenv("NOTIFY_EMAIL", "ops@example.com")

    def boom(*a, **k):
        raise OSError("connection refused")
    monkeypatch.setattr(notify, "_deliver", boom)
    resp = submit_lead(client)
    assert resp.status_code == 302
    with app.app_context():
        db = get_db()
        assert db.execute("SELECT COUNT(*) FROM leads").fetchone()[0] == 1
        assert db.execute("SELECT error FROM email_logs WHERE status = 'failed'").fetchone()[0] == "connection refused"


def test_sla_alert_sent_once(client, app, outbox):
    from medical_platform import notify
    submit_lead(client)
    outbox.clear()
    with app.app_context():
        get_db().execute("UPDATE leads SET created_at = '2020-01-01T00:00:00+00:00'")
        get_db().commit()
    with app.app_context():
        assert notify.send_sla_alerts() == 1
        assert notify.send_sla_alerts() == 0
    assert [to for to, _, _ in outbox] == ["ops@example.com"] and "SLA" in outbox[0][1]


def test_assignment_change_notifies_new_assignee(client, app, outbox):
    submit_lead(client, preferred_language="en")
    with app.app_context():
        db = get_db()
        db.execute("UPDATE staff SET email = 'zh@wizx.test' WHERE username = 'coord_zh'")
        db.commit()
        lead_id = db.execute("SELECT id FROM leads").fetchone()["id"]
        zh_id = db.execute("SELECT id FROM staff WHERE username = 'coord_zh'").fetchone()["id"]
    outbox.clear()
    login(client)
    token = csrf(client, f"/admin/leads/{lead_id}")
    client.post(f"/admin/leads/{lead_id}/action", data={"_csrf": token, "action": "update", "lead_status": "new",
                                                         "assigned_staff_id": zh_id, "priority": "normal"})
    assert [to for to, _, _ in outbox] == ["zh@wizx.test"]


def test_real_smtp_delivery(app, monkeypatch):
    if app.config["DB_BACKEND"] == "d1":
        pytest.skip("SMTP 경로는 DB 백엔드와 무관")
    """로컬 SMTP 서버로 실제 SMTP 대화를 검증한다."""
    import socket
    import threading
    from medical_platform import notify

    received = {}
    srv = socket.socket()
    srv.bind(("127.0.0.1", 0))
    srv.listen(1)
    port = srv.getsockname()[1]

    def serve():
        conn, _ = srv.accept()
        f = conn.makefile("rwb")
        w = lambda line: (f.write(line.encode() + b"\r\n"), f.flush())
        w("220 test")
        data_mode, lines = False, []
        for raw in f:
            line = raw.decode().rstrip("\r\n")
            if data_mode:
                if line == ".":
                    data_mode = False
                    received["data"] = "\n".join(lines)
                    w("250 queued")
                else:
                    lines.append(line)
                continue
            cmd = line.split(" ")[0].upper()
            if cmd == "EHLO":
                w("250 localhost")
            elif cmd == "DATA":
                data_mode = True
                w("354 go")
            elif cmd == "QUIT":
                w("221 bye")
                break
            else:
                if cmd == "RCPT":
                    received["rcpt"] = line
                w("250 ok")
        conn.close()

    t = threading.Thread(target=serve, daemon=True)
    t.start()
    monkeypatch.setenv("MAIL_BACKEND", "smtp")
    monkeypatch.setenv("SMTP_HOST", "127.0.0.1")
    monkeypatch.setenv("SMTP_PORT", str(port))
    monkeypatch.setenv("SMTP_STARTTLS", "0")
    monkeypatch.setenv("MAIL_FROM", "WIZ X Medical <no-reply@example.com>")
    with app.app_context():
        assert notify.send("test", "ops@example.com", "제목 테스트", "본문")
    t.join(5)
    assert "ops@example.com" in received["rcpt"] and "Subject:" in received["data"]


def test_production_seed_creates_only_admin(tmp_path):
    app = create_app({"TESTING": True, "DATABASE": str(tmp_path / "p.db"), "SECRET_KEY": "x",
                      "ADMIN_INITIAL_PASSWORD": "a-long-secret-pass"})
    with app.app_context():
        names = [r["username"] for r in get_db().execute("SELECT username FROM staff")]
    assert names == ["admin"]
    client = app.test_client()
    assert login(client, "admin", "admin1234").status_code == 200  # 기본 비밀번호 거부 (로그인 화면 재표시)
    assert login(client, "admin", "a-long-secret-pass").status_code == 302


def test_d1_row_behaves_like_sqlite_row():
    from medical_platform.db import D1Connection
    conn = D1Connection(FakeD1(":memory:"), run_sync=lambda v: v)
    conn.executescript("CREATE TABLE t (id INTEGER PRIMARY KEY, name TEXT, note TEXT); CREATE TABLE u (x INT)")
    cur = conn.execute("INSERT INTO t (name, note) VALUES (?, ?)", ("a", None))
    assert cur.lastrowid == 1
    row = conn.execute("SELECT id, name, note FROM t").fetchone()
    assert row[0] == 1 and row["name"] == "a" and row["note"] is None
    assert dict(row) == {"id": 1, "name": "a", "note": None}
    assert conn.execute("SELECT COUNT(*) FROM t").fetchone()[0] == 1
    assert {r[1] for r in conn.execute("PRAGMA table_info(t)")} == {"id", "name", "note"}
    conn.executemany("INSERT INTO u (x) VALUES (?)", [(1,), (2,)])
    assert [r["x"] for r in conn.execute("SELECT x FROM u ORDER BY x")] == [1, 2]
