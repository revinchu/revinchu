# AI 도우미와 MCP 연결 (2026-10-12)

사용자가 요청한 양방향 연결을 구현했다. 데이터 → AI 도우미에서 다섯 AI API를 호출하고, 데이터 → MCP 연결에서 외부 앱의 문서 읽기·편집을 허용한다.

## 사용 방법

- **AI 도우미 (Alt+A+I+I)**: 온라인 개인 보관함 연결 후 ChatGPT·OpenAI, Claude, Perplexity, Grok·xAI, DeepSeek를 선택하고 해당 서비스의 API 키를 입력한다. 모델 ID는 계정에서 사용 가능한 모델로 바꿀 수 있다. 일반 채팅 구독과 API 사용은 별개다.
- 선택 셀 전송은 기본 해제다. 켜면 최대 500셀의 입력·저장 값을 전송하며 전체 파일은 전송하지 않는다. 답변의 변경 전·후를 보고 적용하면 단일 Workbook 트랜잭션으로 실행해 Ctrl+Z로 취소할 수 있다. 문서·시트·버전이 바뀐 뒤에는 적용을 거절한다.
- API 키는 탭 메모리에만 남긴다. localStorage·IndexedDB·문서·서버 저장소에 저장하지 않는다. 중계 시 고정된 공식 API에만 전달하며, 키 삭제와 요청 취소를 제공한다.
- **MCP 연결 (Alt+A+I+M)**: 외부 앱에 `https://wixel-3.wizx.workers.dev/api/mcp`를 입력한다. OAuth와 클라이언트 자동 등록(DCR)을 선택한다. CIMD 대신 자동 등록을 사용한다. 위셀 동의 화면에서 온라인 문서 하나와 요청된 권한을 확인한다. 보관함 복구키는 외부 앱에 전달하지 않는다.
- OAuth가 없는 MCP 클라이언트에는 문서별 토큰을 발급해 `Authorization: Bearer <토큰>` 헤더를 사용한다. URL에 넣지 않는다. 토큰은 발급 직후만 표시한다. MCP 연결 화면에서 해제할 수 있다.
- OAuth 액세스 토큰은 1시간, 연결·갱신 토큰·수동 토큰은 30일 후 만료된다. 갱신 시 회전하고 이전 토큰은 재사용할 수 없다. 보관함당 최대 20개 연결이다.

직접 MCP 연결은 각 AI 앱·플랜·관리자 설정의 지원 여부에 따라 달라진다. 모든 Grok·DeepSeek·Perplexity 채팅 화면에 사용자 지정 MCP 메뉴가 있다는 의미가 아니다. 위셀 내부의 다섯 API 경로와 구분한다. 계정별 실제 소비자 앱 연결과 유료 AI 응답은 API 키가 없어 미검증이다.

## MCP 도구

| 도구 | 동작 |
|---|---|
| `list_documents` | 허용한 온라인 문서만 이름·현재 버전 조회 |
| `describe_workbook` | 시트 이름·범위·보호·처리 제한 조회 |
| `read_range` | 시트 이름·A1 범위로 최대 1,000셀 읽기 |
| `write_cells` | 편집 권한에만 노출. 시트·최대 200셀·expectedRevision으로 저장 |

온라인 저장본에 연결한다. 브라우저의 미저장 내용은 포함하지 않는다. 문서명을 바꾸면 새 연결이 필요하다. 수식은 재계산하지 않고 저장 값임을 `valuesAreCached`로 명시한다. 기존 revision CAS와 이전 버전 기록을 이용해 충돌 후 자동 덮어쓰기를 하지 않는다.

MCP는 저장 본문 및 압축 해제 후 각각 8MiB, 10만 셀, 100시트까지다. 일반 위셀의 대용량 파일 한도는 유지한다. 필요한 시트를 별도 온라인 문서로 저장해 연결할 수 있다. 읽기 결과는 128,000자까지다. 압축 폭탄과 프레임의 거대한 배열·셀 run은 모델 생성 전에 차단한다. 보호·최종본·피벗/배열 수식 시트와 병합 하위 셀, 외부 참조·네트워크 수식은 AI 자동 편집을 제한한다.

## 구현

- `src/ai-ui.js`, `ai-cells.js`: 두 화면, 일회 표시 토큰, 변경 미리보기, 범위·보호 검증. `app.js`가 적용 직전 문서 정체성·버전·보호를 확인한다.
- `server/ai-provider.js`: 고정 공식 목적지, Claude Messages/나머지 Chat Completions 형식, 크기·55초 제한, upstream 오류 본문 비노출. 임의 URL을 프록시하지 않는다.
- `cloudflare/ai-auth.js`, `ai-state.js`: DCR·PKCE S256·동의·토큰·폐기. 새 `AiState` DO만 추가하는 `v2-ai` 마이그레이션이며 기존 문서 SQL은 변환·삭제하지 않는다.
- `cloudflare/mcp.js`: Streamable HTTP JSON 응답, MCP 2025-03-26/06-18/11-25, 도구 스키마·권한·revision. 기존 packed/v2 저장 모듈로 스타일·조건부 서식·메타데이터를 보존한다.
- `ai-connect.html/js/css`: 외부 스크립트 동의 화면. HttpOnly·Secure·SameSite 쿠키로 요청과 브라우저를 묶고 정확한 등록 콜백·client·resource·PKCE를 검증한다.
- 기존 개인 API는 같은 출처 제한을 유지한다. MCP/OAuth만 알려진 클라이언트 Origin을 허용하며 Origin 없는 서버 클라이언트도 인증한다. 모든 진입점에 요청 제한이 있다.
- API: `GET/POST /api/ai/connections`, `DELETE /api/ai/connections/:id`, `POST /api/ai/chat`, `/api/oauth/{register,authorize,pending,deny,token,revoke}`, `POST /api/ai/consent`, `/.well-known/oauth-authorization-server`, `/.well-known/oauth-protected-resource/api/mcp`.

## 검증

- Node 전체 2,649개: 본체 2,642 + modules 7 통과. 최초 리본 개수 고정 검증 364→366 변경 필요 1건을 수정한 후 전체 통과.
- Cloudflare 저장소 회귀 27개, 실제 로컬 workerd 기존 API 25확인 통과.
- 새 Node 테스트 9개: 범위·5개 API 형식·키/오류·PKCE·CSRF·토큰 만료/회전/폐기·문서 격리·읽기 전용·충돌·조건부 서식 왕복·압축 폭탄.
- `cloudflare/ai-integration.mjs`: 독립 합성 보관함으로 HTTP/OAuth/DO/읽기·편집·충돌·해제 41확인 통과. 문서·권한 정리.
- `tools/ai-ui-smoke.mjs`: Chromium에서 다섯 공급자, 선택 셀 전송, XSS 비실행, 변경 미리보기·적용·취소, 오래된 결과 거절, 키 비영속화, 토큰 발급·해제, 브라우저 OAuth 동의 통과. 최초 실행은 health 초기화 전 버튼 실행으로 실패해 초기화 대기를 추가했다.
- 리본 401명령 bad 0, 일반/Cloudflare 빌드·Wrangler dry-run 통과. 스크린샷은 `D:\Codex\Temp\wixel-ai-mcp`에 저장했다.
- 유료 모델 응답은 모의 응답으로 검사했다. 실제 업무 파일·사용자 복구키·API 키를 검증에 사용하지 않았다.

## 공식 참조

[MCP 전송](https://modelcontextprotocol.io/specification/2025-11-25/basic/transports), [MCP 인증](https://modelcontextprotocol.io/specification/2025-11-25/basic/authorization), [OpenAI MCP OAuth](https://developers.openai.com/plugins/build/auth), [OpenAI Chat Completions](https://developers.openai.com/api/reference/resources/chat/subresources/completions/methods/create), [Claude 원격 연결](https://support.claude.com/en/articles/11175166-get-started-with-custom-connectors-using-remote-mcp), [Claude Messages](https://platform.claude.com/docs/en/api/messages/create), [Perplexity 경로](https://docs.perplexity.ai/docs/agent-api/migrate-from-sonar/overview), [xAI API](https://docs.x.ai/developers/model-capabilities/legacy/chat-completions), [DeepSeek API](https://api-docs.deepseek.com/api/create-chat-completion/).

## 공개 배포

- 제품 커밋 `611f323`, 브랜치 `codex/ai-mcp-integrations`, 원격 push 완료.
- Cloudflare 버전 `d247f557-26f9-4807-a7d2-3d8850de6fe6`.
- 공개 주소 `https://wixel-3.wizx.workers.dev/`, MCP `https://wixel-3.wizx.workers.dev/api/mcp`.
- 자산 `wixel-fec2f0431b599554.js`, 5,979,892 bytes, SHA256 `fec2f0431b599554c95856593f290239681db01a8779a1319a978150b81cf952`.
- 공개 HTML·JS·동의 화면 HTML/JS/CSS·version.json 6파일 모두 로컬과 바이트 동일. health의 ai/mcp true, 기존 CSP 유지 확인.
- 공개 실제 API 41확인과 공개 Chromium 다섯 공급자 UI·변경안 적용/취소·stale 차단·키 비영속·토큰 해제·OAuth 동의 통과. AI 응답만 모의 처리하여 과금 호출을 하지 않았다. 합성 문서와 발급 권한은 검사 후 삭제했다.
- 공개 검증 JSON 및 스크린샷은 `D:\Codex\Temp\wixel-ai-mcp\public-verification.json`, `D:\Codex\Temp\wixel-ai-mcp\public`에 보관한다. 테스트 브라우저와 전용 로컬 workerd 종료.
