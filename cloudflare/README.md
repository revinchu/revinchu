# Cloudflare 배포 (Workers + D1)

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

npx wrangler d1 create mediprism      # 출력된 database_id 를 wrangler.jsonc 에 붙여넣기

# 시크릿 (값은 프롬프트에 입력 — 저장소에 남지 않음)
npx wrangler secret put SECRET_KEY              # 아무 긴 무작위 문자열
npx wrangler secret put ADMIN_INITIAL_PASSWORD  # 첫 관리자(admin) 비밀번호. 설정하면 데모 직원 계정은 만들지 않음
npx wrangler secret put NOTIFY_EMAIL            # 운영 알림 받을 주소 (쉼표로 여러 개)
npx wrangler secret put RESEND_API_KEY          # https://resend.com API 키
npx wrangler secret put MAIL_FROM               # 예: MEDIPRISM <no-reply@내도메인.com>
npx wrangler secret put SITE_URL                # 예: https://mediprism.<계정>.workers.dev (메일 속 링크용)
```

## 배포 (코드 바꿀 때마다)

```bash
python build.py                       # build/ 에 번들 생성 (템플릿·아이콘을 _embedded.py 로 포함)
uv run pywrangler deploy              # → https://mediprism.<계정>.workers.dev
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

Cloudflare 대시보드 → Workers & Pages → mediprism → Settings → Domains & Routes → Add Custom Domain.
그다음 `SITE_URL` 시크릿을 새 주소로 바꿉니다.

## 참고

- D1은 문장마다 자동 커밋됩니다. 백업: `npx wrangler d1 export mediprism --remote --output backup.sql`
- 로그: `npx wrangler tail` 또는 대시보드 Observability
