"""SQLite 스키마, 연결 관리, 초기 데이터.

병원 DB와 환자(리드) DB는 별도 테이블로 분리하고, 리드-병원 관계는 referrals 로 관리한다.
"""
import json
import sqlite3
from datetime import datetime, timezone

from flask import current_app, g
from werkzeug.security import generate_password_hash

SCHEMA = """
CREATE TABLE IF NOT EXISTS staff (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    name TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('admin', 'coordinator', 'marketer')),
    languages TEXT NOT NULL DEFAULT '',
    active INTEGER NOT NULL DEFAULT 1,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS hospitals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    slug TEXT UNIQUE NOT NULL,
    name_ko TEXT NOT NULL,
    name_en TEXT NOT NULL,
    hospital_type TEXT NOT NULL DEFAULT 'clinic',
    region TEXT NOT NULL,
    district TEXT NOT NULL DEFAULT '',
    address_ko TEXT NOT NULL DEFAULT '',
    address_en TEXT NOT NULL DEFAULT '',
    lat REAL,
    lng REAL,
    specialties TEXT NOT NULL DEFAULT '',
    languages TEXT NOT NULL DEFAULT '',
    price_band INTEGER NOT NULL DEFAULT 2,
    foreign_patient_registered INTEGER NOT NULL DEFAULT 0,
    intl_coordinator INTEGER NOT NULL DEFAULT 0,
    phone TEXT NOT NULL DEFAULT '',
    website TEXT NOT NULL DEFAULT '',
    summary_i18n TEXT NOT NULL DEFAULT '{}',
    is_sponsored INTEGER NOT NULL DEFAULT 0,
    verification_status TEXT NOT NULL DEFAULT 'planned',
    last_verified_at TEXT,
    next_review_at TEXT,
    is_demo INTEGER NOT NULL DEFAULT 0,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS hospital_tags (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    hospital_id INTEGER NOT NULL REFERENCES hospitals(id) ON DELETE CASCADE,
    tag TEXT NOT NULL,
    evidence TEXT NOT NULL DEFAULT '',
    source_url TEXT NOT NULL DEFAULT '',
    verified_at TEXT
);

CREATE TABLE IF NOT EXISTS doctors (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    hospital_id INTEGER NOT NULL REFERENCES hospitals(id) ON DELETE CASCADE,
    name_ko TEXT NOT NULL,
    name_en TEXT NOT NULL,
    title TEXT NOT NULL DEFAULT '',
    specialty TEXT NOT NULL DEFAULT '',
    credentials TEXT NOT NULL DEFAULT '',
    source_url TEXT NOT NULL DEFAULT '',
    verified_at TEXT
);

CREATE TABLE IF NOT EXISTS price_items (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    hospital_id INTEGER NOT NULL REFERENCES hospitals(id) ON DELETE CASCADE,
    item_name_en TEXT NOT NULL,
    item_name_ko TEXT NOT NULL DEFAULT '',
    price_domestic_krw INTEGER,
    price_foreign_krw INTEGER,
    includes TEXT NOT NULL DEFAULT '',
    excludes TEXT NOT NULL DEFAULT '',
    source TEXT NOT NULL DEFAULT '',
    verified_at TEXT
);

CREATE TABLE IF NOT EXISTS partners (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    partner_type TEXT NOT NULL,
    name TEXT NOT NULL,
    region TEXT NOT NULL DEFAULT '',
    lat REAL,
    lng REAL,
    description_i18n TEXT NOT NULL DEFAULT '{}',
    price_band INTEGER NOT NULL DEFAULT 2,
    languages TEXT NOT NULL DEFAULT '',
    website TEXT NOT NULL DEFAULT '',
    is_demo INTEGER NOT NULL DEFAULT 0,
    active INTEGER NOT NULL DEFAULT 1
);

CREATE TABLE IF NOT EXISTS leads (
    id TEXT PRIMARY KEY,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL,
    full_name TEXT NOT NULL,
    email TEXT NOT NULL DEFAULT '',
    phone TEXT NOT NULL DEFAULT '',
    messenger_type TEXT NOT NULL DEFAULT '',
    messenger_id TEXT NOT NULL DEFAULT '',
    contact_channel TEXT NOT NULL,
    nationality TEXT NOT NULL DEFAULT '',
    country_of_residence TEXT NOT NULL DEFAULT '',
    preferred_language TEXT NOT NULL,
    age_band TEXT NOT NULL DEFAULT '',
    treatment_category TEXT NOT NULL,
    treatment_interest TEXT NOT NULL DEFAULT '',
    hospital_preferences TEXT NOT NULL DEFAULT '',
    preferred_hospital_id INTEGER REFERENCES hospitals(id),
    budget_band TEXT NOT NULL DEFAULT '',
    preferred_visit_date TEXT NOT NULL DEFAULT '',
    stay_nights INTEGER,
    travel_party_size INTEGER,
    concierge_needs TEXT NOT NULL DEFAULT '',
    accommodation_pref TEXT NOT NULL DEFAULT '',
    food_pref TEXT NOT NULL DEFAULT '',
    itinerary_interests TEXT NOT NULL DEFAULT '',
    vip_request INTEGER NOT NULL DEFAULT 0,
    message TEXT NOT NULL DEFAULT '',
    source_channel TEXT NOT NULL DEFAULT '',
    utm_source TEXT NOT NULL DEFAULT '',
    utm_medium TEXT NOT NULL DEFAULT '',
    utm_campaign TEXT NOT NULL DEFAULT '',
    utm_content TEXT NOT NULL DEFAULT '',
    utm_term TEXT NOT NULL DEFAULT '',
    landing_page TEXT NOT NULL DEFAULT '',
    referrer TEXT NOT NULL DEFAULT '',
    site_language TEXT NOT NULL DEFAULT '',
    lead_status TEXT NOT NULL DEFAULT 'new',
    priority TEXT NOT NULL DEFAULT 'normal',
    disqualify_reason TEXT NOT NULL DEFAULT '',
    assigned_staff_id INTEGER REFERENCES staff(id),
    first_response_at TEXT,
    last_contact_at TEXT,
    next_contact_at TEXT,
    duplicate_of TEXT,
    consent_record_id INTEGER,
    anonymized_at TEXT
);
CREATE INDEX IF NOT EXISTS idx_leads_status ON leads(lead_status);
CREATE INDEX IF NOT EXISTS idx_leads_email ON leads(email);

CREATE TABLE IF NOT EXISTS consents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id TEXT NOT NULL REFERENCES leads(id),
    privacy_required INTEGER NOT NULL,
    third_party_share INTEGER NOT NULL DEFAULT 0,
    overseas_transfer INTEGER NOT NULL DEFAULT 0,
    marketing INTEGER NOT NULL DEFAULT 0,
    policy_version TEXT NOT NULL,
    ip TEXT NOT NULL DEFAULT '',
    user_agent TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    withdrawn_at TEXT
);

CREATE TABLE IF NOT EXISTS lead_activities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id TEXT NOT NULL REFERENCES leads(id),
    staff_id INTEGER REFERENCES staff(id),
    kind TEXT NOT NULL,
    body TEXT NOT NULL,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS referrals (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id TEXT NOT NULL REFERENCES leads(id),
    hospital_id INTEGER NOT NULL REFERENCES hospitals(id),
    status TEXT NOT NULL DEFAULT 'requested',
    quote_note TEXT NOT NULL DEFAULT '',
    appointment_at TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS service_orders (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    lead_id TEXT NOT NULL REFERENCES leads(id),
    service_type TEXT NOT NULL,
    partner_id INTEGER REFERENCES partners(id),
    status TEXT NOT NULL DEFAULT 'requested',
    note TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS audit_logs (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    staff_id INTEGER REFERENCES staff(id),
    action TEXT NOT NULL,
    target TEXT NOT NULL DEFAULT '',
    ip TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS page_events (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    event TEXT NOT NULL,
    path TEXT NOT NULL DEFAULT '',
    hospital_id INTEGER,
    lang TEXT NOT NULL DEFAULT '',
    utm_source TEXT NOT NULL DEFAULT '',
    utm_campaign TEXT NOT NULL DEFAULT '',
    created_at TEXT NOT NULL
);
"""


def now_iso():
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(current_app.config["DATABASE"])
        g.db.row_factory = sqlite3.Row
        g.db.execute("PRAGMA foreign_keys = ON")
    return g.db


def close_db(_exc=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db(seed=True):
    db = get_db()
    db.executescript(SCHEMA)
    if seed and db.execute("SELECT COUNT(*) FROM staff").fetchone()[0] == 0:
        seed_db(db)
    db.commit()


# ---------------------------------------------------------------------------
# 초기 데이터
# 아래 병원·제휴사는 화면 확인용 "가상 샘플"이다. 실제 병원명·가격·명의·유명인 방문 정보가 아니며
# is_demo=1 로 표시되어 사이트에 'DEMO' 배지가 노출된다. 실제 운영 시 관리자 화면에서 교체할 것.
# ---------------------------------------------------------------------------

DEMO_HOSPITALS = [
    {
        "slug": "demo-gangnam-skin",
        "name_ko": "[샘플] 강남 스킨 클리닉",
        "name_en": "[Sample] Gangnam Skin Clinic",
        "hospital_type": "clinic",
        "region": "seoul", "district": "Gangnam",
        "address_ko": "서울 강남구 (샘플 주소)", "address_en": "Gangnam-gu, Seoul (sample address)",
        "lat": 37.4979, "lng": 127.0276,
        "specialties": "dermatology",
        "languages": "en,zh,ja",
        "price_band": 2, "foreign_patient_registered": 1, "intl_coordinator": 1,
        "summary": {
            "en": "Sample dermatology clinic near Gangnam Station focusing on laser and skin-care programs.",
            "ko": "강남역 인근 레이저·피부관리 중심 피부과 (샘플 데이터).",
        },
        "tags": [("value", "Published price list compared with 3 nearby clinics (sample)"),
                 ("foreign_favorite", "Foreign patient share disclosed by clinic (sample)")],
        "doctors": [("김샘플", "Dr. Sample Kim", "Director", "dermatology",
                     "Board-certified dermatologist (sample)")],
        "prices": [("Laser toning (1 session)", "레이저 토닝 1회", 150000, 150000, "Consultation, numbing cream", "Aftercare products"),
                   ("Skin booster (1 session)", "스킨부스터 1회", 300000, 300000, "Consultation", "Medication")],
        "status": "published",
    },
    {
        "slug": "demo-sinsa-plastic",
        "name_ko": "[샘플] 신사 성형외과",
        "name_en": "[Sample] Sinsa Plastic Surgery",
        "hospital_type": "clinic",
        "region": "seoul", "district": "Sinsa",
        "address_ko": "서울 강남구 신사동 (샘플 주소)", "address_en": "Sinsa-dong, Seoul (sample address)",
        "lat": 37.5163, "lng": 127.0203,
        "specialties": "plastic_surgery",
        "languages": "en,zh,ja,vi",
        "price_band": 3, "foreign_patient_registered": 1, "intl_coordinator": 1,
        "summary": {"en": "Sample plastic surgery clinic with a dedicated international lounge and VIP escort service.",
                    "ko": "외국인 전용 라운지와 VIP 의전을 제공하는 성형외과 (샘플 데이터)."},
        "tags": [("vip", "Private lounge and chauffeur service confirmed by clinic (sample)"),
                 ("celebrity", "Publicly reported visit, usage permission on file (sample)"),
                 ("foreign_favorite", "International patient counts disclosed (sample)")],
        "doctors": [("이샘플", "Dr. Sample Lee", "Chief Surgeon", "plastic_surgery",
                     "Board-certified plastic surgeon, 15 yrs (sample)")],
        "prices": [("Double eyelid surgery", "쌍꺼풀 수술", 1800000, 1800000, "Consultation, anesthesia, 1 follow-up", "Medication")],
        "status": "published",
    },
    {
        "slug": "demo-jongno-checkup",
        "name_ko": "[샘플] 종로 종합검진센터",
        "name_en": "[Sample] Jongno Health Screening Center",
        "hospital_type": "general_hospital",
        "region": "seoul", "district": "Jongno",
        "address_ko": "서울 종로구 (샘플 주소)", "address_en": "Jongno-gu, Seoul (sample address)",
        "lat": 37.5704, "lng": 126.9922,
        "specialties": "health_checkup,cardiology",
        "languages": "en,zh,ru",
        "price_band": 2, "foreign_patient_registered": 1, "intl_coordinator": 1,
        "summary": {"en": "Sample comprehensive health screening center with same-day results in English.",
                    "ko": "당일 영문 결과지를 제공하는 종합검진센터 (샘플 데이터)."},
        "tags": [("korean_favorite", "Listed in public screening-quality evaluation (sample)"),
                 ("value", "Package price compared with similar packages (sample)")],
        "doctors": [("박샘플", "Dr. Sample Park", "Head of Screening", "health_checkup",
                     "Family medicine specialist (sample)")],
        "prices": [("Premium health screening package", "프리미엄 종합검진", 1200000, 1250000,
                    "Lab tests, CT, endoscopy, report", "Sedation fee", )],
        "status": "published",
    },
    {
        "slug": "demo-seocho-dental",
        "name_ko": "[샘플] 서초 치과",
        "name_en": "[Sample] Seocho Dental",
        "hospital_type": "clinic",
        "region": "seoul", "district": "Seocho",
        "address_ko": "서울 서초구 (샘플 주소)", "address_en": "Seocho-gu, Seoul (sample address)",
        "lat": 37.4837, "lng": 127.0324,
        "specialties": "dental",
        "languages": "en,ja",
        "price_band": 1, "foreign_patient_registered": 0, "intl_coordinator": 0,
        "summary": {"en": "Sample dental clinic offering implants and whitening at published prices.",
                    "ko": "임플란트·미백 가격을 공개하는 치과 (샘플 데이터)."},
        "tags": [("value", "Lowest published implant price among sampled clinics (sample)")],
        "doctors": [("최샘플", "Dr. Sample Choi", "Director", "dental", "Prosthodontist (sample)")],
        "prices": [("Dental implant (1 tooth)", "임플란트 1개", 1100000, 1100000, "Fixture, crown", "Bone graft")],
        "status": "published",
    },
    {
        "slug": "demo-songpa-general",
        "name_ko": "[샘플] 송파 대학병원",
        "name_en": "[Sample] Songpa University Hospital",
        "hospital_type": "general_hospital",
        "region": "seoul", "district": "Songpa",
        "address_ko": "서울 송파구 (샘플 주소)", "address_en": "Songpa-gu, Seoul (sample address)",
        "lat": 37.5266, "lng": 127.1086,
        "specialties": "oncology,cardiology,orthopedics",
        "languages": "en,zh,ru,ar",
        "price_band": 3, "foreign_patient_registered": 1, "intl_coordinator": 1,
        "summary": {"en": "Sample tertiary hospital with an international healthcare center and specialist doctors.",
                    "ko": "국제진료센터와 전문 의료진을 갖춘 상급종합병원 (샘플 데이터)."},
        "tags": [("top_doctor", "Specialist with published academic record and award (sample)"),
                 ("korean_favorite", "Official accreditation listed (sample)"),
                 ("intl_certified", "International accreditation (sample)")],
        "doctors": [("정샘플", "Dr. Sample Jung", "Professor", "oncology",
                     "Oncology specialist, 100+ papers, national award (sample)")],
        "prices": [],
        "status": "published",
    },
    {
        "slug": "demo-busan-korean-medicine",
        "name_ko": "[샘플] 부산 한방병원",
        "name_en": "[Sample] Busan Korean Medicine Hospital",
        "hospital_type": "korean_medicine",
        "region": "busan", "district": "Haeundae",
        "address_ko": "부산 해운대구 (샘플 주소)", "address_en": "Haeundae-gu, Busan (sample address)",
        "lat": 35.1631, "lng": 129.1636,
        "specialties": "korean_medicine",
        "languages": "en,ja",
        "price_band": 1, "foreign_patient_registered": 1, "intl_coordinator": 0,
        "summary": {"en": "Sample Korean medicine hospital near Haeundae beach offering wellness programs.",
                    "ko": "해운대 인근 웰니스 프로그램 한방병원 (샘플 데이터)."},
        "tags": [("value", "Program price compared with similar programs (sample)")],
        "doctors": [],
        "prices": [("Acupuncture & herbal consultation", "침·한약 상담", 80000, 80000, "Consultation, 1 acupuncture", "Herbal medicine")],
        "status": "published",
    },
]

DEMO_PARTNERS = [
    ("hotel", "[Sample] Gangnam Medical Stay Residence", "seoul", 37.5010, 127.0250,
     {"en": "Serviced residence 5 min from Gangnam clinics, recovery-friendly rooms.", "ko": "강남 병원가 도보 5분, 회복 친화 객실 레지던스 (샘플)."}, 2),
    ("hotel", "[Sample] Jongno Heritage Hotel", "seoul", 37.5720, 126.9890,
     {"en": "Hotel near palaces, ideal for health-screening guests.", "ko": "고궁 인근, 검진 고객에게 적합한 호텔 (샘플)."}, 3),
    ("kfood", "[Sample] Seoul Royal Hanjeongsik", "seoul", 37.5740, 126.9850,
     {"en": "Traditional Korean course meal, soft-diet options after procedures.", "ko": "한정식 코스, 시술 후 부드러운 식단 가능 (샘플)."}, 3),
    ("kfood", "[Sample] Gangnam Korean BBQ House", "seoul", 37.4990, 127.0300,
     {"en": "Popular Korean BBQ with English/Chinese menus.", "ko": "영어·중국어 메뉴가 있는 고깃집 (샘플)."}, 2),
    ("interpreter", "[Sample] MediTalk Interpreters", "seoul", None, None,
     {"en": "Certified medical interpreters: EN, ZH, JA, VI, RU.", "ko": "의료통역 전문 (영·중·일·베·러) (샘플)."}, 2),
    ("transport", "[Sample] Incheon Airport Pickup", "seoul", None, None,
     {"en": "Airport pickup and clinic transfers with English-speaking drivers.", "ko": "공항 픽업·병원 이동 (샘플)."}, 2),
    ("tour", "[Sample] K-Culture Recovery Tour", "seoul", None, None,
     {"en": "Light-activity tours: palaces, markets, K-beauty shopping.", "ko": "회복기 가벼운 일정: 고궁·시장·K뷰티 쇼핑 (샘플)."}, 1),
    ("vip", "[Sample] Prestige Concierge", "seoul", None, None,
     {"en": "Private chauffeur, dedicated coordinator, luxury suites.", "ko": "전담 기사·코디네이터·스위트 객실 VIP 의전 (샘플)."}, 3),
]


def seed_db(db):
    ts = now_iso()
    today = ts[:10]
    staff = [
        ("admin", "admin1234", "관리자", "admin", "ko,en"),
        ("coord_en", "coord1234", "영어 코디네이터", "coordinator", "en"),
        ("coord_zh", "coord1234", "중국어 코디네이터", "coordinator", "zh"),
        ("coord_ja", "coord1234", "일본어 코디네이터", "coordinator", "ja,vi"),
        ("marketer", "market1234", "마케터", "marketer", ""),
    ]
    for username, pw, name, role, langs in staff:
        db.execute(
            "INSERT INTO staff (username, password_hash, name, role, languages, created_at) VALUES (?,?,?,?,?,?)",
            (username, generate_password_hash(pw), name, role, langs, ts),
        )

    for h in DEMO_HOSPITALS:
        cur = db.execute(
            """INSERT INTO hospitals (slug, name_ko, name_en, hospital_type, region, district, address_ko, address_en,
               lat, lng, specialties, languages, price_band, foreign_patient_registered, intl_coordinator,
               summary_i18n, verification_status, last_verified_at, is_demo, created_at, updated_at)
               VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,1,?,?)""",
            (h["slug"], h["name_ko"], h["name_en"], h["hospital_type"], h["region"], h["district"],
             h["address_ko"], h["address_en"], h["lat"], h["lng"], h["specialties"], h["languages"],
             h["price_band"], h["foreign_patient_registered"], h["intl_coordinator"],
             json.dumps(h["summary"], ensure_ascii=False), h["status"], today, ts, ts),
        )
        hid = cur.lastrowid
        for tag, evidence in h["tags"]:
            db.execute("INSERT INTO hospital_tags (hospital_id, tag, evidence, verified_at) VALUES (?,?,?,?)",
                       (hid, tag, evidence, today))
        for name_ko, name_en, title, spec, cred in h["doctors"]:
            db.execute(
                "INSERT INTO doctors (hospital_id, name_ko, name_en, title, specialty, credentials, verified_at) VALUES (?,?,?,?,?,?,?)",
                (hid, name_ko, name_en, title, spec, cred, today))
        for item in h["prices"]:
            en, ko, dom, fgn, inc, exc = item
            db.execute(
                """INSERT INTO price_items (hospital_id, item_name_en, item_name_ko, price_domestic_krw, price_foreign_krw,
                   includes, excludes, source, verified_at) VALUES (?,?,?,?,?,?,?,?,?)""",
                (hid, en, ko, dom, fgn, inc, exc, "Sample data", today))

    for ptype, name, region, lat, lng, desc, band in DEMO_PARTNERS:
        db.execute(
            """INSERT INTO partners (partner_type, name, region, lat, lng, description_i18n, price_band, is_demo)
               VALUES (?,?,?,?,?,?,?,1)""",
            (ptype, name, region, lat, lng, json.dumps(desc, ensure_ascii=False), band))
