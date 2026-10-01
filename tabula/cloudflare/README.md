# Cloudflare Worker 실행·검증

앱 브라우저 코드는 기존 빌드를 사용하고, 이 폴더는 Wrangler가 별도 번들링합니다. Worker/DO 프로덕션 npm 의존성은 0개입니다. API·보관함 키·공유·용량·보안 한계는 [API_CONTRACT.md](API_CONTRACT.md)에 있습니다.

현재 공개 서비스 주소: [WIXEL 3](https://wixel-3.wizx.workers.dev). 주소 변경 후 HTTPS 접속 확인, 최초 배포 이력 및 운영 검증은 [12_공개배포와운영검증.md](../docs/codex/12_공개배포와운영검증.md)에 기록합니다.

## 환경

검증 환경은 Node 24.19.0, 외부 도구 폴더의 Wrangler 4.145.0입니다. Node 단위 테스트는 node:sqlite가 있는 Node 22.13 이상을 사용합니다. 작업 파일은 D:에 둡니다. 애플리케이션 package.json에 Wrangler를 추가하지 않습니다.

PowerShell (저장소의 tabula 폴더):
~~~powershell
$env:TEMP = 'D:\Codex\Temp'
$env:TMP = 'D:\Codex\Temp'
$env:npm_config_cache = 'D:\Codex\Caches\npm'
$env:XDG_CONFIG_HOME = 'D:\Codex\Config'
$env:WRANGLER_SEND_METRICS = 'false'
$wrangler = 'D:\Codex\Temp\wixel3-tools\node_modules\wrangler\bin\wrangler.js'

node build.mjs --cloud
node --test cloudflare/backend.test.js
node $wrangler types cloudflare/worker-configuration.d.ts --include-runtime=false --config cloudflare/wrangler.jsonc
node $wrangler deploy --dry-run --outdir ..\.local\worker-build --config cloudflare/wrangler.jsonc
node $wrangler dev --local --ip 127.0.0.1 --port 8787 --persist-to ..\.local\state --config cloudflare/wrangler.jsonc
~~~

다른 터미널에서:
~~~powershell
node cloudflare/integration.mjs
node cloudflare/recovery-integration.mjs # 로컬 전용: 버전 기록/복원·백업 원자 복원·충돌·정리
node cloudflare/limits-integration.mjs  # 선택: 20MiB 저장/다운로드와 1바이트 초과 413
~~~

integration.mjs는 임의의 합성 복구키로 로컬 API를 호출해 격리·충돌·청크·백업·게시를 확인한 뒤 생성한 개인 문서를 지우고 게시를 중지합니다. 실제 사용자 파일/키를 읽지 않습니다. 실패 시 정리되지 않은 합성 데이터가 로컬 .local/state에 남을 수 있습니다. 다른 서버를 검증하려면 WIXEL_WORKER_URL을 명시합니다. 공개 운영 서버에 대한 실행은 실제 저장 요청이 발생하므로 배포 담당자가 범위를 확인해야 합니다.

macOS/Linux에서도 node 명령은 동일하며 Wrangler 경로와 환경 변수 문법, 경로 구분자만 환경에 맞게 바꿉니다. 도구 설치 예시는 npm install --prefix /path/to/tools --no-save --package-lock=false wrangler@4.145.0입니다.

## 검증 범위와 운영

backend.test.js는 실제 SQLite와 합성 스트림으로 JSON 문법·UTF-8·청크 경계·CAS·원자성·동시 저장·버전 재사용 방지·용량 제한·백업 스냅숏·URL/리디렉션/응답 제한을 검사합니다. integration.mjs는 로컬 workerd의 실제 Durable Object와 HTTP 응답을 확인합니다. 런타임에 해당하는 환경 타입은 Wrangler가 생성한 worker-configuration.d.ts이며, 전체 런타임 타입을 복사하거나 앱에 @types/node를 추가하지 않습니다.

문서 업로드는 총 30초 제한입니다. 문서와 백업 출력은 ReadableStream으로 전달해 큰 파일을 RPC 값으로 한꺼번에 복제하지 않습니다. 과거 버전은 문서당20개/보관함100MiB 범위에서 보존하며, 보존 대상에서 빠진 청크도 조회 중이면 스트림 완료·취소 후 정리합니다. 인스턴스를 다시 열면 일반 미완료 업로드와 현재·과거 어느 쪽에서도 참조하지 않는 청크만 정리합니다.

백업 복원은 작은 manifest, 문서별 임시 업로드, 최종 원자 commit의 세 단계입니다. 서버가 보관함 전체 JSON을 메모리 객체로 펼치지 않습니다. 현재 문서100MiB와 과거본100MiB 외에 복원 준비용100MiB가 최대30분간 필요할 수 있습니다. SDK는 사용자가 선택한 이름의 교체만 수행하며, 새 연결로 전환되면 다음 업로드를 중단합니다. Node 서버는 capabilities가 없으므로 명확히 미지원으로 안내합니다. API·용량·삭제·결과 불확실 정책은 [API 계약](API_CONTRACT.md#온라인-버전-기록과-복원)을 참고하세요.

recovery-integration.mjs는 localhost/127.0.0.1만 허용하며 합성 보관함을 만들고 실제 HTTP로 과거본·If-Match 복원·잘못된 JSON·중간 충돌·여러 문서의 commit·백업을 확인한 뒤 해당 합성 문서와 과거본을 삭제합니다. backend.test.js에는 구 스키마의 추가 마이그레이션, 재시작, 읽기 스트림 유지/취소, 한도별 정리, 두 번째 문서 SQL 실패의 전체 롤백도 포함합니다. 공개 서버나 사용자 데이터의 검증 결과를 뜻하지 않습니다.

미존재 보관함·게시물의 읽기는 DDL을 실행하지 않습니다. 로컬 workerd는 그 조회에도 자체 SQLite/메타데이터 파일을 만들 수 있습니다. 플랫폼 내부 기록, 요청 실행과 SQL 읽기의 비용이 0이라고 보장하지 않습니다.

아직 게시한 적 없는 상태에서 임의 키로 쓰는 요청은 새 보관함을 만들 수 있으므로 대규모 공개 운영에는 추가 가입/봇 방어 및 계정 비용 관제가 필요합니다. Rate Limit binding은 세계 단일 카운터가 아닙니다. 무료/유료 계정의 CPU·DO·스토리지 한도는 배포 담당자가 확인해야 합니다. config는 요금제와 충돌할 수 있는 사용자 지정 CPU 값을 강제하지 않고 플랫폼 기본값을 사용합니다.

배포:
~~~powershell
node $wrangler deploy --config cloudflare/wrangler.jsonc
~~~

새 환경을 배포하면 v1 migration이 SQLite Durable Object 두 클래스를 생성합니다. 계정 인증 및 public workers.dev 서브도메인 설정은 배포 담당자가 수행합니다. 사용자 지정 도메인을 추가하면 동일 Origin API 동작, 자체 주소로 되돌아가는 프록시 요청, CSP를 해당 도메인에서도 다시 확인하세요. 한 환경의 바인딩/상태를 다른 환경에서 재사용하지 마세요.

Cloudflare 정적 자산은 API Worker를 거치지 않을 수 있으므로 앱 빌드의 _headers에도 보안 헤더가 필요합니다. Worker API 응답은 no-store와 nosniff, 엄격한 script-src 및 frame-ancestors CSP를 적용합니다.

## 배포 인증의 최소 범위

2026-10-01에 계정 이메일 인증을 확인한 뒤 아래 OAuth 범위로 로그인·최초 배포에 성공했습니다. Wrangler 4.145.0의 whoami와 배포 결과로 확인한 범위이며, 더 넓은 KV·D1·R2·zone 권한을 추가하지 않았습니다. 실제 배포 버전은 12 문서에 있습니다.

- 배포 기능 scope: workers_scripts:write. CLI가 설명하는 대상은 Worker scripts, Durable Objects, workers.dev 하위 도메인, triggers입니다.
- 계정 검색·선택과 whoami 확인: account:read, user:read. 이 둘은 계정 정보를 확인하기 위한 읽기 권한이며 배포 쓰기 권한은 아닙니다.
- Wrangler는 사용자 지정 --scopes에도 offline_access를 자동 추가합니다. 갱신 토큰 발급용이며 CLI 소스의 generateAuthUrl에서 확인했습니다.
- 현재 구성은 Workers Static Assets와 SQLite DO만 씁니다. workers:write의 광범위 권한, workers_kv:write, d1:write, pages:write, zone:read, workers_routes:write, R2/AI/Queues 권한을 미리 추가할 필요는 없습니다. Assets 업로드 세션의 공식 API 권한도 Workers Scripts Write입니다.

PowerShell에서 실행할 정확한 명령(저장소 tabula 기준):
~~~powershell
$env:XDG_CONFIG_HOME = 'D:\Codex\Config'
$wrangler = 'D:\Codex\Temp\wixel3-tools\node_modules\wrangler\bin\wrangler.js'
node $wrangler whoami
node build.mjs --cloud
node $wrangler deploy --dry-run --outdir ..\.local\worker-build --config cloudflare/wrangler.jsonc
node $wrangler deploy --config cloudflare/wrangler.jsonc
~~~

최초 인증 또는 필요한 재인증에만 다음 명령을 사용합니다. 위 XDG_CONFIG_HOME 설정을 유지합니다.

~~~powershell
node $wrangler login --scopes-list
node $wrangler login --scopes account:read user:read workers_scripts:write --use-keyring
node $wrangler whoami
~~~

이미 이 환경에서 인증을 완료했다면 login을 반복하지 말고 whoami로 확인합니다. 재인증이 필요한 경우에만 login의 브라우저 동의 절차를 진행합니다. XDG_CONFIG_HOME은 로그인·whoami·개발·배포 명령 모두 D:\Codex\Config로 유지합니다. 이 위치 아래 Wrangler가 관리하는 암호화 default.enc와 OS 키체인 사용을 확인했습니다. 암호화 파일·키체인 항목·복호화 키를 복사하거나 Git/전달 ZIP에 포함하지 않습니다. 여러 계정이 있으면 배포 대상 계정을 확인해 CLOUDFLARE_ACCOUNT_ID에 선택한 실제 계정 ID를 지정합니다. 토큰은 채팅/소스/문서에 붙여 넣지 않습니다.

OAuth scope와 계정 회원 역할은 구분해야 합니다. 최신 공식 권한 모델에서 **새 Worker 생성에는 Workers 제품 범위 Admin**, 이미 존재하는 Worker의 재배포에는 해당 Worker의 Editor가 필요합니다. DO는 구현 Worker의 권한을 따르며 별도 DO 역할을 추가하지 않습니다. 현재 배포가 요청하지 않는 사용자 지정 도메인/zone route를 나중에 추가할 때만 해당 zone의 Workers Routes Write 권한을 별도로 검토합니다. 이 권한들은 계정의 기존 역할과 결합되므로 실제 배포가 403이면 실패한 엔드포인트와 계정 역할을 확인하고 필요한 범위만 조정합니다.

공식 근거:
- [Workers 역할과 신규 생성/기존 배포 권한](https://developers.cloudflare.com/workers/authorization/workers/)
- [Durable Objects 권한](https://developers.cloudflare.com/workers/authorization/durable-objects/)
- [Assets 업로드 세션의 Workers Scripts Write 권한](https://developers.cloudflare.com/api/resources/workers/subresources/scripts/subresources/assets/subresources/upload/methods/create/)
- [Wrangler login 옵션](https://developers.cloudflare.com/workers/wrangler/commands/general/#login)
