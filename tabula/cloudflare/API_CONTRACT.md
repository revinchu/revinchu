# WIXEL 3 Cloudflare API 계약

브라우저 앱은 같은 Origin의 API를 사용합니다. 개발 서버의 단일 공유 폴더와 달리 복구키마다 SQLite Durable Object 보관함을 분리합니다. Cloudflare Worker는 기존 앱 번들러와 별도 Wrangler로 빌드하므로 이 폴더의 Worker 엔트리는 플랫폼이 요구하는 default export를 사용합니다. 앱 npm 의존성은 추가하지 않습니다.

## 개인 보관함

브라우저가 crypto.getRandomValues로 32바이트 복구키를 만들고 padding 없는 base64url 43자로 보관합니다. 개인 요청은 X-Wixel-Vault 헤더로 전달합니다. 서버는 SHA-256 해시를 DO 이름과 게시물 소유자 식별자로 사용합니다. 원문 키를 DB/로그/URL에 저장하지 않습니다. 복구키는 암호와 같은 전체 접근 권한입니다. 키를 잃으면 복구할 수 없으며 서비스 운영자나 이메일 인증으로 재설정하는 기능은 없습니다. 브라우저 저장소에 보관한 키와 문서는 같은 Origin의 스크립트가 읽을 수 있습니다. 서버 저장은 종단간 암호화가 아닙니다.

GET /api/health는 인증 없이 {ok:true,auth:false,vault:true,publish:true,maxDocumentBytes,maxVaultBytes,maxDocuments,maxPublications,maxProxyBytes}를 반환합니다. auth:false는 기존 단일 TABULA_TOKEN이 없다는 뜻이며 개인 API에는 반드시 복구키가 필요합니다. 키가 없거나 형식이 틀리면 401/VAULT_KEY_REQUIRED입니다.

| 경로 | 요청 | 응답 |
| --- | --- | --- |
| GET /api/files | 복구키 | [{name,revision,modified,size}] |
| GET /api/files/:name | 복구키 | 문서 JSON 원문; ETag: "revision", X-Wixel-Revision, X-Modified |
| PUT /api/files/:name | 복구키, Content-Type: application/json, If-Match | {ok,revision,modified,size} |
| DELETE /api/files/:name | 복구키, If-Match | {ok:true} |
| GET /api/usage | 복구키 | documents,bytes,publications 및 한도 |
| GET /api/backup | 복구키 | 아래 전체 개인 문서 백업 JSON 다운로드 |
| GET /api/publications | 복구키 | [{id,modified}] |

name은 URL 구성 시 encodeURIComponent로 인코딩합니다. 앞뒤 공백을 제거하며 제어 문자 없이 1~120자입니다. 새 문서는 If-Match: "0", 기존 문서는 마지막으로 읽거나 저장한 revision을 따옴표로 감싸 전송합니다. 서버 버전은 보관함 안에서 증가하며 삭제 후 같은 이름으로 다시 만들어도 재사용하지 않습니다. If-Match가 없으면 428/REVISION_REQUIRED, 다른 기기의 변경과 충돌하면 412/REVISION_CONFLICT 및 currentRevision을 반환합니다. 자동 덮어쓰기 대신 사용자가 최신 문서를 확인하거나 다른 이름으로 저장해야 합니다.

동시 업로드는 보관함당 1개로 제한하며 겹치면 429/UPLOAD_BUSY입니다. 오류 응답은 {error:"한국어 설명",code,...추가값}입니다. 429는 Retry-After: 60을 포함합니다. 저장은 완전한 UTF-8 JSON을 검증하고 완료 시 CAS와 SQLite 트랜잭션으로 원자적으로 교체합니다. 잘못된 JSON, 중단된 업로드와 충돌은 기존 문서를 바꾸지 않습니다.

개인 문서는 최대 50개, 문서 하나 20MiB, 합계 100MiB입니다. DB에는 1MiB 청크로 저장합니다. 제한을 넘으면 413입니다. JSON 문법 깊이는 128단계, 값 개수는 300만 개로 제한합니다. 백업 형식은 {"format":"wixel-vault-backup","version":1,"exportedAt":"ISO 날짜","documents":[{"name","revision","modified","data":문서JSON}]}입니다. 이 API는 다운로드만 지원하며 일괄 복원 엔드포인트는 없습니다.

## 명시적 읽기 전용 게시

| 경로 | 요청 | 응답 |
| --- | --- | --- |
| POST /api/publish | 복구키, application/json 문서 원문 | 201 {id,ok,revision,modified,size} |
| GET /api/published/:id | 인증 불필요 | 게시 문서 JSON 원문 및 ETag/X-Wixel-Revision |
| PUT /api/published/:id | 소유자 복구키, application/json, If-Match | {ok,revision,modified,size} |
| DELETE /api/published/:id | 소유자 복구키, If-Match | {ok:true} |

게시 id도 무작위 32바이트 base64url입니다. 링크를 아는 사람 누구나 읽을 수 있습니다. 개인 보관함 데이터는 게시할 때 복사한 별도 스냅숏이며 개인 문서 저장만으로 게시 내용이 바뀌지 않습니다. 게시 중지 후 GET은 404입니다. 이미 다운로드한 사본은 회수할 수 없습니다. 다른 키는 수정/중지할 수 없습니다.

보관함당 게시 링크 최대 20개이며 각 게시물 20MiB까지입니다. 게시 스냅숏은 개인 100MiB 한도와 별도로 최대 400MiB를 사용할 수 있습니다. 공개 GET은 검색색인/비밀보관을 보장하는 인증 방식이 아닙니다. 게시 목록 id는 복구키를 가진 소유자에게만 표시됩니다. 저장 도중 통신 결과가 불확실하면 목록에 예약 id가 남을 수 있습니다.

## 외부 데이터 가져오기

GET /api/fetch?url=...은 보관함 없이 공개 Google Sheets CSV 등 HTTPS 데이터를 가져옵니다. 응답은 실행되지 않는 text/plain이며 최대 5MiB, 전체 15초, 리디렉션 최대 5회입니다. 모든 단계에서 HTTPS/443, 사용자정보 없는 URL, 공개 도메인 형태를 확인하고 IP 리터럴·로컬 이름·현재 Worker 호스트를 차단합니다. 원래 요청의 쿠키·Authorization·복구키를 외부로 전달하지 않습니다.

Workers에는 Node DNS lookup 및 검증 IP 고정 연결 기능이 없으므로 Node 서버와 같은 DNS pin 검증을 구현했다고 주장하지 않습니다. Cloudflare의 외부 프록시 보안 모델에 의존하며 global_fetch_strictly_public을 켜서 같은 zone의 원본 서버 우회 연결도 막습니다. 로컬 workerd는 운영 네트워크와 다르므로 실제 DNS/네트워크 차단은 배포 환경에서 별도 확인해야 합니다.

GET /api/naver/keywordstool?hintKeywords=...은 복구키와 기존 X-Customer/X-API-Key/X-Secret 헤더가 필요합니다. 고정된 네이버 검색광고 API만 요청하며 키를 저장하지 않고 리디렉션을 허용하지 않습니다.

## 요청 제한과 운영 한계

모든 API는 Cross-Origin Origin 또는 Sec-Fetch-Site: cross-site를 거부합니다. Origin이 없는 CLI 요청은 허용하되 IP 제한을 적용합니다. CORS 헤더를 열지 않습니다. API는 IP당 분당 120회, 외부 가져오기는 IP당 분당 20회, 개인 API는 보관함당 분당 120회의 Rate Limit binding을 사용합니다. 공유 IP 사용자는 같은 한도를 공유할 수 있습니다.

Cloudflare Rate Limiting binding은 위치별 제한이며 전 세계 합계나 비용 상한을 보장하지 않습니다. 무료 계정의 저장량·요청·CPU·DO 한도와 공개 서비스 남용은 계정 운영자가 모니터링해야 합니다. 보관함 키를 무제한 생성하는 공격을 막는 회원 인증/Turnstile/결제 한도는 이 버전에 없습니다. Worker CPU 한도와 저장 용량 한도, 게시 스냅숏 중복 저장을 고려하세요.

소스에서 문서 본문, 복구키, API 비밀키를 로깅하지 않습니다. Wrangler invocation 로그와 trace는 명시적으로 비활성화하고 query string도 로그에서 제거하도록 설정하여 문서 이름/원격 URL을 불필요하게 남기지 않습니다. 이는 애플리케이션의 기록 정책이며 Cloudflare 플랫폼의 모든 내부 운영 기록이 없다는 뜻은 아닙니다.

공식 참고:
- [SQLite 저장 API와 트랜잭션](https://developers.cloudflare.com/durable-objects/api/sqlite-storage-api/)
- [DO 행·RPC·저장 한도](https://developers.cloudflare.com/durable-objects/platform/limits/)
- [ReadableStream/Response RPC](https://developers.cloudflare.com/workers/runtime-apis/rpc/)
- [Workers 보안 모델](https://developers.cloudflare.com/workers/reference/security-model/)
- [global_fetch_strictly_public](https://developers.cloudflare.com/workers/configuration/compatibility-flags/)
- [Rate Limiting binding 한계](https://developers.cloudflare.com/workers/runtime-apis/bindings/rate-limit/)
