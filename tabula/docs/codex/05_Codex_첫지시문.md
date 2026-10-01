# 05. Codex 에 붙여넣을 지시문 (복사해서 사용)

> 아래 예시는 `0a479d7` 최초 인계 시점 기준입니다. 이후 품질 개선을 받은 경우에는
> [07_품질개선결과.md](07_품질개선결과.md)와 [08_Codex_다음지시문.md](08_Codex_다음지시문.md)를 사용해 이미 고친 기능을 다시 만들지 마세요.

아래 A~D는 **다음 작업을 맡길 때 하나씩 선택해 붙여넣는 인계 예시**입니다. 문서를 읽는 것만으로 네 작업을 모두 실행하라는 뜻은 아닙니다.
이번 인계 보완은 검증 도구·문서·전달 묶음을 준비하는 작업이며, A의 P0 구현은 후속 작업입니다.

원본 인계는 `claude/offsro-excel-implementation-zmro3r`의 `f6cbc6e`입니다.
이 문서와 검증 도구 수정을 포함한 **`codex/handoff-validation`을 후속 작업의 기준 브랜치**로 사용하세요.
현재 실측 결과와 미검증 항목은 [`06_검증결과.md`](06_검증결과.md)에 있습니다.

## A. 첫 작업 (환경 확인 + 보안 P0-1 · P0-2)
```text
이 저장소의 tabula/는 WIXEL(엑셀과 같은 웹 스프레드시트)입니다.
이번 목표는 P0-1 웹 가져오기 중계 SSRF 강화와 P0-2 인증 기본값 개선입니다.

먼저 AGENTS.md, tabula/AGENTS.md, tabula/docs/codex/00~04와 06_검증결과.md,
수정 영역에 해당하는 CLAUDE.md 기술 노트를 읽으세요.
Git 상태·원격·기준 커밋을 확인하고 기존 변경을 보존한 상태에서
codex/handoff-validation의 최신 인계 커밋을 기준으로 새 작업 브랜치를 만드세요.
이 기준이 확인되지 않으면 추측해서 다른 브랜치로 바꾸지 말고 확인한 상태를 보고하세요.
원본 이력은 claude/offsro-excel-implementation-zmro3r의 f6cbc6e입니다.

기존 체크아웃 경로를 유지하고, 새 로컬 작업은 D:\Codex\Workspaces,
임시 파일은 D:\Codex\Temp, 의존성 캐시는 D:\Codex\Caches를 사용하세요.

1) 기준 상태 검증
   tabula에서 npm test와 npm run build를 실행하세요.
   실행 환경, 기준 커밋, 실제 통과·실패·건너뜀 개수를 기록하세요.
   176개 통과는 이전 인계 기록이며 고정된 현재 기대 개수가 아닙니다.
   실패가 있으면 변경 전 실패와 이번 변경으로 생긴 실패를 구분하세요.

2) docs/codex/02_보완항목내역서.md의 P0-1 구현
   - server.js의 URL/IP 검사와 요청 처리를 테스트 가능한 함수로 분리하고
     test/server.test.js에 회귀 테스트를 추가하세요. 모듈 import만으로 서버가 열리지 않게 하세요.
   - http/https와 포트 80/443만 허용하고, DNS가 반환한 모든 주소에서
     사설·루프백·링크로컬·CGNAT·IPv6 ULA·IPv4-mapped IPv6 등 차단 대상을 검사하세요.
   - 검사한 주소를 실제 소켓 연결에 사용하세요. 검사 뒤 다시 DNS를 조회하여
     다른 내부 주소로 접속하는 우회가 없어야 하며 Host와 HTTPS 인증서 검증은 유지하세요.
   - 리디렉션은 수동으로 최대 5회, 각 목적지를 연결 전에 다시 검사하세요.
   - 응답 본문은 스트리밍 중 크기를 제한하고 시간 제한을 유지하세요.
   - DNS와 전송을 테스트에서 제어해 재현 가능한 테스트를 작성하세요.
     정상 공개 주소, 공개/내부 혼합 DNS 응답, DNS 재조회 우회, 주소 표현 변형,
     제한 포트, 내부 주소로 향하는 리디렉션, 리디렉션 초과, 본문 초과를 포함하세요.
     함수 검사 외에 로컬 테스트 서버에 실제 소켓 요청을 보내 요청 경로에도
     정책이 적용되는지 확인하세요. 테스트용 우회가 운영 환경에 노출되면 안 됩니다.

3) 같은 문서의 P0-2 구현
   - 루프백 외부 수신 주소에서 TABULA_TOKEN이 없을 때의 동작을 명시적으로 정하고,
     /api/files, /api/publish, 게시본 변경·삭제, /api/fetch 같은 보호 API가
     무인증으로 노출되지 않게 하세요. 공개 읽기용 게시 링크의 정책도 문서에 명시하세요.
   - 네이버 검색광고 중계도 인증 뒤로 이동하세요.
   - 로컬 127.0.0.1 사용, 토큰 있음/없음/오류, 정상 토큰,
     네이버 중계의 무인증 거부와 앱의 토큰 입력 흐름을 검증하세요.
   - 서버 응답과 시작 정책을 실제 HTTP 요청 및 필요한 자식 프로세스 테스트로 검증하세요.
   - 서버 정책에 맞게 호스팅 가이드와 관련 안내를 갱신하세요.

규칙: npm 의존성 추가 금지. tabula/AGENTS.md의 import/export 및 빌드 규칙 준수.
Node 전용 서버와 테스트는 기존 Node 내장 모듈 사용 방식을 참고하세요.
사용자 메시지는 한국어이며, 사용자 업무 파일·토큰·검증용 개인정보는 커밋하지 마세요.
P0-1과 P0-2의 변경과 검증을 구분해서 진행하세요.

완료 시 npm test, npm run build와 추가한 서버 검증 결과를 보고하세요.
UI 변경이 있으면 tools/README.md에 따라 smoke도 실행하세요.
실행 명령·종료 코드·실제 결과, 변경 요약, 남은 한계를 함께 제시하세요.
P0-3 및 P0-4를 포함한 다른 항목의 완료 여부와 실제 Excel 검증 여부는 별도로 표시하세요.
```

## B. 엑셀 호환 작업 (실제 파일이 있을 때)
```text
AGENTS.md와 tabula/AGENTS.md를 읽고 codex/handoff-validation에서 새 작업 브랜치를 만드세요.
tabula/tools/README.md대로 Playwright를 --no-save --package-lock=false로 설치하고
캐시는 D:\Codex\Caches, 임시 검증 결과는 D:\Codex\Temp에 두세요.
테스트 서버는 HOST=127.0.0.1과 별도의 임시 TABULA_DATA를 사용하세요.

D:\Codex\Temp\wixel-samples의 사용자 xlsx/xlsm/xlsb 파일 목록과 SHA-256을 먼저 기록하세요.
파일이 없으면 실제 파일 검증을 했다고 가정하지 말고 필요한 파일을 알려 주세요.
코드를 바꾸기 전에 동일 파일마다 node tools/pvcmp.mjs와 node tools/check.mjs를 실행해
before.txt를 저장하세요. 그런 다음 docs/codex/02_보완항목내역서.md의
P1-1을 y60.xlsb부터 분석하고 수정하세요.
수정 뒤 같은 파일과 명령으로 after.txt를 만들고 파일별 불일치 수와 원인을 비교하세요.
TODAY/NOW/RAND 등의 차이는 자동 제외되지 않으므로 검토 근거를 따로 남기세요.
나빠진 파일이 하나라도 있으면 원인을 해결하기 전에는 커밋하지 마세요.
사용자 파일과 해당 파일의 값이 담긴 로그는 저장소나 전달 묶음에 넣지 마세요.
```

## C. 기능 추가 작업 (예: 피벗 테이블 이동)
```text
AGENTS.md와 tabula/AGENTS.md를 읽고 codex/handoff-validation에서 새 작업 브랜치를 만드세요.
임시 파일과 검증 결과는 D:\Codex\Temp에 두세요.
tabula/docs/codex/03_스킬가이드.md의 §2(명령 · 리본), §6(피벗)을 따라
'피벗 테이블 분석 › 동작 › 피벗 테이블 이동'을 구현하세요 (02_보완항목내역서 P1-2).
- 대화상자: 새 워크시트 / 기존 워크시트(위치 refInput)
- def.top/left/area 갱신 후 다시 그리기, 원래 자리는 지움, 실행 취소 가능
- 단위 테스트(가능한 부분) + node tools/smoke.mjs가 bad 0, 종료 코드 0
- npm test와 npm run build 성공
- WHATS_NEW와 CLAUDE.md(또는 docs/codex/01)에 한 줄 기록
실제 실행 결과와 변경 요약을 보고하세요.
```

## D. 큰 구조 정리 (P3-1 app.js 분할의 첫 단계)
```text
AGENTS.md와 tabula/AGENTS.md를 읽고 codex/handoff-validation에서 새 작업 브랜치를 만드세요.
임시 파일과 검증 결과는 D:\Codex\Temp에 두세요.
수정 전에 npm test, npm run build, node tools/smoke.mjs, node tools/keys.mjs를 실행해
결과와 건너뛴 검증 항목을 기록하세요.
tabula/docs/codex/02_보완항목내역서.md P3-1에 따라 src/app.js에서
피벗 UI 부분만 먼저 src/app-pivot-ui.js로 옮기세요.
공유 상태는 src/app-state.js로, 순환 import가 생기면 안 됩니다 (npm run build가 검사).
동작은 바뀌면 안 됩니다. 수정 후 같은 검증에서 실패 증가가 없어야 하고
스모크는 bad 0, 종료 코드 0이어야 합니다. keys의 SKIP을 통과 개수에 합치지 마세요.
이 단계의 변경 요약과 검증 결과를 보고하고, 전체 P3-1 완료와 구분하세요.
```
