# WIXEL 3 Cloudflare API 계약

브라우저 앱은 같은 Origin의 API를 사용합니다. 개발 서버의 단일 공유 폴더와 달리 복구키마다 SQLite Durable Object 보관함을 분리합니다. Cloudflare Worker는 기존 앱 번들러와 별도 Wrangler로 빌드하므로 이 폴더의 Worker 엔트리는 플랫폼이 요구하는 default export를 사용합니다. 앱 npm 의존성은 추가하지 않습니다.

## 개인 보관함

브라우저가 crypto.getRandomValues로 32바이트 복구키를 만들고 padding 없는 base64url 43자로 보관합니다. 개인 요청은 X-Wixel-Vault 헤더로 전달합니다. 서버는 SHA-256 해시를 DO 이름과 게시물 소유자 식별자로 사용합니다. 원문 키를 DB/로그/URL에 저장하지 않습니다. 복구키는 암호와 같은 전체 접근 권한입니다. 키를 잃으면 복구할 수 없으며 서비스 운영자나 이메일 인증으로 재설정하는 기능은 없습니다. 브라우저 저장소에 보관한 키와 문서는 같은 Origin의 스크립트가 읽을 수 있습니다. 서버 저장은 종단간 암호화가 아닙니다.

GET /api/health는 인증 없이 {ok:true,auth:false,vault:true,publish:true,versionHistory:true,backupImport:true,maxDocumentBytes,maxVaultBytes,maxDocuments,maxPublications,maxProxyBytes,maxVersions,maxHistoryBytes,importTimeoutMs}를 반환합니다. auth:false는 기존 단일 TABULA_TOKEN이 없다는 뜻이며 개인 API에는 반드시 복구키가 필요합니다. 키가 없거나 형식이 틀리면 401/VAULT_KEY_REQUIRED입니다. Node 서버와 이전 서버에 versionHistory/backupImport가 없으면 클라이언트는 해당 기능을 지원하지 않는다고 안내합니다.

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

개인 문서는 최대 50개, 문서 하나 32MiB, 현재 문서 합계 100MiB입니다. 이 한도는 전송하는 UTF-8 JSON 본문의 바이트 수이며, 압축 문서의 base64 포장도 포함합니다. 서버는 포장 안의 문서를 풀어 메모리에 올리지 않고 JSON 문법과 실제 스트림 바이트 수를 확인합니다. DB에는 1MiB 청크로 저장합니다. 제한을 넘으면 413입니다. JSON 문법 깊이는 128단계, 값 개수는 300만 개로 제한합니다. 백업 형식은 {"format":"wixel-vault-backup","version":1,"exportedAt":"ISO 날짜","documents":[{"name","revision","modified","data":문서JSON}]}입니다. 백업은 현재 개인 문서만 포함하며 과거 버전·게시물·복구키는 포함하지 않습니다.

## 온라인 버전 기록과 복원

| 경로 | 요청 | 응답 |
| --- | --- | --- |
| GET /api/versions?name=:name | 복구키 | {currentRevision,versions:[{revision,modified,size}],maxVersions,maxHistoryBytes} |
| GET /api/version?name=:name&revision=:revision | 복구키 | 해당 과거 문서 JSON, ETag/X-Wixel-Revision/X-Modified |
| POST /api/version?name=:name&revision=:revision | 복구키, 현재 문서의 If-Match | {ok,name,revision,modified,size} |

이름은 encodeURIComponent로 구성합니다. 목록은 과거본만 최신순이며, 현재 버전은 currentRevision입니다. 복원은 현재 문서를 과거 이력에 남기고 선택한 내용을 **새로운 revision**으로 저장합니다. 오래된 If-Match는 412이며 원문을 바꾸지 않습니다. 삭제하거나 정리된 과거본은 404/VERSION_NOT_FOUND입니다.

과거본은 문서당 최대 20개, 보관함 전체 100MiB의 별도 한도로 보관합니다. 최신 버전부터 문서별 개수와 전체 잔여 용량을 만족하는 항목을 남깁니다. 따라서 큰 문서는 20개보다 적게 보존될 수 있습니다. 현재 문서 100MiB와 합하면 보존 데이터는 최대 200MiB이며, 진행 중 읽기가 붙잡은 이전 청크와 업로드 임시 청크는 추가로 잠시 존재할 수 있습니다. 현재 문서 삭제는 그 문서의 모든 과거본도 삭제합니다. 휴지통 기능이 아닙니다. 읽기 중인 청크는 완료/취소 후 지워집니다. 공개 게시물에는 과거본을 남기지 않습니다.

기존 DB에는 versions/imports/import_items/import_chunks 테이블을 CREATE TABLE IF NOT EXISTS로 추가합니다. 기존 현재 문서와 revision을 변환하지 않습니다. 도입 전에 이미 삭제된 과거본을 복구할 수는 없습니다. 재시작 시 현재 또는 이력에 연결된 청크는 보존하고, 일반 업로드의 미완료 staging과 참조되지 않는 청크만 청소합니다.

## 백업의 원자적 일괄 복원

브라우저의 server.importBackup(backup,{overwrite,expectedRevisions,signal,onProgress})가 다음 순서를 처리합니다. overwrite 기본값은 false입니다. true는 사용자가 기존 이름의 교체를 명시적으로 선택한 경우만 사용하고, expectedRevisions에는 확인 화면에서 조회한 server.list() 결과를 전달합니다. **백업 안 revision은 대상 보관함의 권한이나 현재 버전으로 사용하지 않습니다.** 선택한 문서만 추가/교체하며 백업에 없는 현재 문서는 유지합니다.

| 경로 | 요청 | 응답 |
| --- | --- | --- |
| POST /api/imports | 복구키, application/json manifest | 201 {id,expires,documents:개수} |
| PUT /api/imports/:id/files/:name | 복구키, application/json 문서 원문, manifest와 같은 If-Match | {ok,name,size} |
| POST /api/imports/:id/commit | 복구키 | {ok,documents:[{name,revision,modified,size}]} |
| DELETE /api/imports/:id | 복구키 | {ok:true} |

manifest는 {format:"wixel-vault-backup",version:1,mode:"create" 또는 "replace",documents:[{name,expectedRevision}]}이며 32KiB/10초로 제한합니다. 문서는 1~50개, 정규화한 이름은 중복할 수 없습니다. create의 expectedRevision은 전부 0이어야 하므로 기존 이름을 덮어쓰지 않습니다. replace도 새 이름은 0, 기존 이름은 조회한 현재 revision이어야 합니다.

각 문서는 기존 저장과 같은 UTF-8·JSON 깊이/복잡도·32MiB·30초 검증을 거쳐 1MiB 청크로 임시 저장합니다. 임시 문서 합계는 100MiB입니다. 일부만 전송한 상태의 commit은 409/IMPORT_INCOMPLETE입니다. 모든 문서의 CAS와 최종 개수·용량을 다시 확인하고, 하나의 SQLite transactionSync 안에서 전체를 반영합니다. JSON 오류, 중단, 충돌, 용량 초과, SQL 오류는 현재 문서를 일부만 바꾸지 않습니다. 전송 완료된 상태는 DO 재시작 후에도 유지됩니다.

보관함별 준비 세션은 한 개이며 30분 후 만료됩니다. DO alarm 또는 다음 접근/재시작에서 만료 데이터를 청소합니다. 취소하면 준비 청크만 지웁니다. SDK는 실패 시 취소를 시도합니다. 최종 commit 뒤 통신이 끊기면 성공 여부가 불확실할 수 있으므로 IMPORT_RESULT_UNKNOWN을 안내하고 자동 재실행하지 않습니다. 현재 온라인 목록을 다시 확인해야 합니다. import ID는 같은 보관함에서만 사용할 수 있으며 URL에 복구키를 넣지 않습니다.

## 명시적 읽기 전용 게시

| 경로 | 요청 | 응답 |
| --- | --- | --- |
| POST /api/publish | 복구키, application/json 문서 원문 | 201 {id,ok,revision,modified,size} |
| GET /api/published/:id | 인증 불필요 | 게시 문서 JSON 원문 및 ETag/X-Wixel-Revision |
| PUT /api/published/:id | 소유자 복구키, application/json, If-Match | {ok,revision,modified,size} |
| DELETE /api/published/:id | 소유자 복구키, If-Match | {ok:true} |

게시 id도 무작위 32바이트 base64url입니다. 링크를 아는 사람 누구나 읽을 수 있습니다. 개인 보관함 데이터는 게시할 때 복사한 별도 스냅숏이며 개인 문서 저장만으로 게시 내용이 바뀌지 않습니다. 게시 중지 후 GET은 404입니다. 이미 다운로드한 사본은 회수할 수 없습니다. 다른 키는 수정/중지할 수 없습니다.

보관함당 게시 링크 최대 20개이며 각 게시물 32MiB까지입니다. 게시 스냅숏은 개인 100MiB 한도와 별도로 최대 640MiB를 사용할 수 있습니다. 공개 GET은 검색색인/비밀보관을 보장하는 인증 방식이 아닙니다. 게시 목록 id는 복구키를 가진 소유자에게만 표시됩니다. 저장 도중 통신 결과가 불확실하면 목록에 예약 id가 남을 수 있습니다.

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
