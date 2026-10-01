# WIXEL 3.0 후속 작업 지시문

아래는 후속 Codex 작업에 사용할 지시문입니다. 이 문서 자체가 현재 사용자 요청을 대신하지 않습니다.

공개 배포와 작은 합성 데이터 운영 검증을 완료했습니다. 주소·버전과 아직 검증하지 않은 운영 범위는 [12_공개배포와운영검증.md](12_공개배포와운영검증.md)를 확인하세요. 최신 키팁·QAT 기본값은 [13_단축키점검.md](13_단축키점검.md)에 있습니다. 이전 문서의 인증 대기 상태로 되돌려 판단하지 마세요.

```text
저장소 revinchu/revinchu의 codex/wixel-3-cloudflare 브랜치에서 시작하세요.
AGENTS.md, tabula/AGENTS.md, tabula/docs/codex/09_WIXEL3_구현과검증.md,
tabula/docs/codex/11_셀스타일과지우기.md, tabula/docs/codex/12_공개배포와운영검증.md, tabula/docs/codex/13_단축키점검.md,
tabula/docs/codex/14_차트와표시검증.md, tabula/cloudflare/API_CONTRACT.md를 읽고 실제 Git 상태를 확인하세요.
기존 사용자 데이터와 복구키를 출력·커밋하거나 임의로 업로드하지 마세요.

D:에 작업·임시·캐시를 두고 기존 작업 경로를 보존하세요.
앱 npm 의존성을 추가하지 말고 기존 순수 ES 모듈 번들 규칙을 지키세요.
Cloudflare 런타임과 Wrangler 개발 도구는 앱과 별도입니다.

1. cd tabula에서 npm test와 npm run build, npm run build:cloud를 실행하세요.
2. npm start와 별도 wrangler dev로 소스·Cloudflare 빌드를 검증하세요.
3. tools/smoke.mjs, tools/keys.mjs, tools/ui-regressions.mjs,
   tools/format-regressions.mjs, tools/security-regressions.mjs, tools/wixel3-ui.mjs,
   tools/cell-style-regressions.mjs, tools/clear-ui.mjs,
   tools/keytips.mjs, tools/chart-ui.mjs, tools/pivot-options.mjs, tools/grid-rendering.mjs,
   tools/vault-ui-integration.mjs 및 cloudflare/backend.test.js와 integration.mjs를 실행하세요.
   브라우저 도구는 WIXEL_URL과 PLAYWRIGHT_MODULE 환경 변수를 사용할 수 있습니다.
   셀 스타일·지우기 도구는 사용자 파일 대신 합성 문서를 검사하며,
   상세 범위는 tools/README.md, 최종 결과는 docs/codex/11_셀스타일과지우기.md를 확인하세요.
4. 실제 파일이 제공되면 고치기 전과 후 check.mjs/pvcmp.mjs/brcheck.mjs 결과를
   파일별로 비교하세요. 테스트 데이터를 사용자 업무 파일이라고 주장하지 마세요.
5. 새 서식은 화면·실행 취소·브라우저 저장·표준 XLSX 저장/재열기를 함께 검사하세요.
6. 공개 게시와 개인 보관함 경계, 복구키 노출, DOM 삽입 경로와 CSP를 보존하세요.
7. 실제 Excel 앱 검증, 비공개 Google Sheets OAuth, 공동 편집, 모든 3D/차트
   유형 지원 같은 미완료 범위를 완료했다고 보고하지 마세요.
8. 테스트 결과, 실제 변경 범위, 남은 제한을 한국어로 기록하고 커밋하세요.
   배포는 사용자가 허용한 Cloudflare 계정과 프로젝트에 한정하세요.
   Wrangler 명령은 XDG_CONFIG_HOME=D:\Codex\Config를 유지하고 기존 OS 키체인 인증을 사용하세요.
   공개 운영 검사는 작은 합성 보관함으로 순차 수행하며 실패 시에도 자기 문서·게시만 정리하세요.
   로컬 UI·스모크 도구를 공개 주소에 무조건 실행하거나 사용자 자료를 시험 데이터로 쓰지 마세요.
```
