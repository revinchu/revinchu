# Cloudflare 배포 (Workers + D1)

> 브랜드명 미정: 임시로 Worker·DB 이름은 `wizx-medical`, 화면 표시는 `BRAND_NAME=WIZ X Medical` 입니다.
> 이름이 정해지면 `wrangler.jsonc` 의 `BRAND_NAME`만 바꿔 다시 배포하면 사이트·메일 표기가 모두 바뀝니다.
> 주소는 커스텀 도메인을 연결하면 되므로 Worker 이름을 바꿀 필요는 없습니다.

Flask 앱(`medical_platform/`)을 **Cloudflare Python Workers**에서 그대로 실행합니다.
DB는 **Cloudflare D1**(SQLite 호환), 정적 파일은 **Workers Static Assets**, SLA 알림은 **Cron Trigger**로 동작합니다.
무료 플랜으로 시작할 수 있습니다 (Workers Free: 하루 10만 요청, D1 Free 포함).

```
요청 ─▶ Cloudflare ─┬─ /static/*  → build/public (정적 자산, Worker 미경유)
                     └─ 그 외      → worker.py → Flask(WSGI) → D1(DB)
cron(10분) ─▶ worker.py scheduled → 미응대 문의 이메일 알림
```

## 처음 한 번

```bash
cd cloudflare
pip install uv                        # uv 0.13.4 이상
npx wrangler login                    # 브라우저에서 Cloudflare 로그인 (또는 CLOUDFLARE_API_TOKEN 환경 변수)

npx wrangler d1 create wizx-medical      # 출력된 database_id 를 wrangler.jsonc 에 붙여넣기

# 시크릿 (값은 프롬프트에 입력 — 저장소에 남지 않음)
npx wrangler secret put SECRET_KEY              # 아무 긴 무작위 문자열
npx wrangler secret put ADMIN_INITIAL_PASSWORD  # 첫 관리자(admin) 비밀번호. 설정하면 데모 직원 계정은 만들지 않음
npx wrangler secret put NOTIFY_EMAIL            # 운영 알림 받을 주소 (쉼표로 여러 개)
npx wrangler secret put RESEND_API_KEY          # https://resend.com API 키
npx wrangler secret put MAIL_FROM               # 예: WIZ X Medical <no-reply@내도메인.com>
npx wrangler secret put SITE_URL                # 예: https://wizx-medical.<계정>.workers.dev (메일 속 링크용)
```

## 배포 (코드 바꿀 때마다)

```bash
python build.py                       # build/ 에 번들 생성 (템플릿·아이콘을 _embedded.py 로 포함)
uv run pywrangler deploy              # → https://wizx-medical.<계정>.workers.dev
```

첫 요청 때 D1에 테이블과 샘플 병원 데이터가 자동 생성됩니다. `/admin` 에서 `admin` / `ADMIN_INITIAL_PASSWORD` 로 로그인하세요.

## 로컬 실행 (Cloudflare 런타임)

```bash
printf 'SECRET_KEY=dev\nADMIN_INITIAL_PASSWORD=dev-pass\nMAIL_BACKEND=console\n' > .dev.vars
python build.py && uv run pywrangler dev   # http://localhost:8787 (로컬 D1 사용)
```

## 이메일 (Resend)

Workers 에서는 SMTP 소켓을 쓸 수 없어 HTTP API인 **Resend**로 보냅니다 (무료: 월 3,000통, 하루 100통).

- 도메인 인증 전: `MAIL_FROM=onboarding@resend.dev` 로 **Resend 가입 이메일 주소로만** 보낼 수 있습니다.
  운영 알림(NOTIFY_EMAIL)을 가입 주소로 받는 테스트는 바로 됩니다.
- 신청자 접수 확인 메일까지 보내려면 Resend 에서 도메인을 추가하고, 안내되는 DNS 레코드를 Cloudflare DNS 에 넣으세요.

## 커스텀 도메인

Cloudflare 대시보드 → Workers & Pages → wizx-medical → Settings → Domains & Routes → Add Custom Domain.
그다음 `SITE_URL` 시크릿을 새 주소로 바꿉니다.

## 실제 데이터 (공공데이터포털)

병원·비급여 가격은 건강보험심사평가원, 관광지는 한국관광공사 TourAPI 에서 가져옵니다.
data.go.kr 에서 `병원정보서비스`, `비급여진료비정보조회서비스`, `국문 관광정보 서비스_GW`, `영문 관광정보서비스_GW` 를 활용신청한 뒤:

```bash
python cloudflare/set_data_key.py                       # '일반 인증키(Decoding)' 를 복사해 둔 상태에서 실행 (클립보드에서 읽음)
python cloudflare/import_public_data.py attractions     # → cloudflare/.import/import_attractions.sql
python cloudflare/import_public_data.py hospitals       # → cloudflare/.import/import_hospitals.sql
npx wrangler d1 execute wizx-medical --remote --file cloudflare/.import/import_attractions.sql
npx wrangler d1 execute wizx-medical --remote --file cloudflare/.import/import_hospitals.sql
```

- 병원: 서울·부산의 치과병원, 안과 병원, 상급종합병원(검진성 검사), 한방병원 중 비급여 가격을 공개한 곳.
  비급여 공개 제도는 미용 목적 시술(보톡스·필러·쌍꺼풀 등)을 포함하지 않아 피부과·성형외과 가격은 공공데이터로 가져올 수 없습니다.
- 공공데이터에는 외국인 적용가가 없어 국내 공개가(최저~최고)만 표시하고, 외국인 가격은 예약 전 병원 확인으로 안내합니다.
- 다시 실행하면 같은 병원(요양기호 기준)은 갱신됩니다. 샘플 병원은 `archived`, 샘플 제휴사는 비활성으로 바뀝니다.

## 참고

- D1은 문장마다 자동 커밋됩니다. 백업: `npx wrangler d1 export wizx-medical --remote --output backup.sql`
- 로그: `npx wrangler tail` 또는 대시보드 Observability
