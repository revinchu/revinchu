# 05. Codex 에 붙여넣을 지시문 (복사해서 사용)

## A. 첫 작업 (환경 확인 + 보안 P0)
```
이 저장소의 tabula/ 는 WIXEL(엑셀과 같은 웹 스프레드시트)입니다.
먼저 AGENTS.md, tabula/AGENTS.md, tabula/docs/codex/00~04 문서를 모두 읽으세요.
브랜치 claude/offsro-excel-implementation-zmro3r 에서 새 브랜치를 만들어 작업합니다.

1) cd tabula && npm test 를 실행해 176개가 모두 통과하는지 확인하고 결과를 보고하세요.
2) docs/codex/02_보완항목내역서.md 의 P0-1(웹 가져오기 중계 SSRF 강화)을 구현하세요.
   - server.js 의 검사 로직을 함수로 분리하고 Node 테스트(test/server.test.js)를 추가
   - DNS 로 해석한 모든 IP 검사, 리디렉션 수동 처리(최대 5회, 단계마다 검사), 포트 80/443 만
3) 같은 방식으로 P0-2(인증 기본값, 네이버 중계를 인증 뒤로)를 구현하세요.
규칙: npm 의존성 추가 금지, import/export 형식 제한(tabula/AGENTS.md), 사용자 메시지는 한국어.
끝나면 npm test, npm run build 결과와 변경 요약을 보고하세요.
```

## B. 엑셀 호환 작업 (실제 파일이 있을 때)
```
tabula/tools/README.md 대로 playwright 를 --no-save 로 설치하고 npm start 로 서버를 띄우세요.
~/wixel-samples/ 의 xlsx/xlsb 파일마다 node tools/pvcmp.mjs 와 node tools/check.mjs 를 돌려
before.txt 를 만든 뒤, 02_보완항목내역서.md 의 P1-1(y60.xlsb 부터)을 분석해 고치세요.
고친 뒤 after.txt 를 만들어 비교하고, 나빠진 파일이 하나라도 있으면 커밋하지 마세요.
사용자 파일은 절대 커밋하지 마세요.
```

## C. 기능 추가 작업 (예: 피벗 테이블 이동)
```
docs/codex/03_스킬가이드.md 의 §2(명령 · 리본), §6(피벗)을 따라
'피벗 테이블 분석 › 동작 › 피벗 테이블 이동' 을 구현하세요 (02_보완항목내역서 P1-2).
- 대화상자: 새 워크시트 / 기존 워크시트(위치 refInput)
- def.top/left/area 갱신 후 다시 그리기, 원래 자리는 지움, 실행 취소 가능
- 단위 테스트(가능한 부분) + node tools/smoke.mjs 가 bad 0
- WHATS_NEW 와 CLAUDE.md(또는 docs/codex/01) 에 한 줄 기록
```

## D. 큰 구조 정리 (P3-1 app.js 분할)
```
02_보완항목내역서.md P3-1 에 따라 src/app.js 에서 피벗 UI 부분만 먼저 src/app-pivot-ui.js 로 옮기세요.
공유 상태는 src/app-state.js 로, 순환 import 가 생기면 안 됩니다 (npm run build 가 검사).
동작은 바뀌면 안 됩니다: npm test, npm run build, node tools/smoke.mjs(bad 0), node tools/keys.mjs 결과가 이전과 같아야 합니다.
```
