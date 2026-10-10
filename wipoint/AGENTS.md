# AGENTS.md — WIPOINT (위포인트) 개발 규칙

WIPOINT 는 **PowerPoint 와 같은 브라우저 프레젠테이션 프로그램**입니다. WIXEL(`../tabula`, 엑셀)과 같은 방식으로 만들었습니다: 순수 ES 모듈, npm 의존성 0개, 빌드하면 HTML 한 파일, Cloudflare Workers 정적 자산으로 배포.

먼저 `docs/00_시작하기.md` → `docs/01_기술상세.md` 를 읽으세요.

## 명령
```bash
cd wipoint
npm start                 # http://127.0.0.1:5179 (개발 서버, src/ 그대로)
npm test                  # Node 내장 테스트 (현재 38개)
npm run build             # dist/index.html (한 파일) + dist/assets/
npm run build:cloud       # dist-cloudflare/ (Cloudflare Workers 용)
PLAYWRIGHT_MODULE=/경로/playwright/index.mjs node tools/smoke.mjs   # 모든 명령 실행 → "bad 0"
```

## 절대 규칙
1. **npm 의존성 추가 금지.** zip · XML · pptx 는 직접 구현(`src/zip.js`, `src/xml.js`, `src/pptx.js`).
2. **import/export 형식**: `import { a, b } from './x.js';` 와 `export const|let|function|class` 만. 동적 `import()`, `export { }`, `export default`, `import *`, 다른 폴더(`../`) import, 순환 import 금지 — `build.mjs` 와 `test/build.test.js` 가 검사합니다.
3. **UI 글자는 모두 한국어.**
4. **모듈끼리 직접 부르지 말고 `state.js` 를 거침**: 명령은 `register({...})` 로 등록하고 `run('이름')` 으로 실행, 다시 그리기는 `emit('change', {scope})`. (WIXEL 의 1.6만 줄 app.js 가 되지 않도록)
5. **DOM 없는 모듈**(model · themes · shapes · chart · render · pptx · smartart · templates · presets · zip · xml)에는 `document`/`window` 를 쓰지 않습니다 — Node 테스트가 돌아야 합니다.
6. 문서 변경은 항상 `change(() => …)` 안에서 (실행 취소 기록). 끌기 같은 연속 변경은 `S.history.record` 를 한 번만.
7. 사용자가 준 실제 업무 파일(.pptx 등)은 커밋하지 않습니다. 테스트 파일은 `test/fixtures/` 의 생성 파일만.

## 완료 기준 (Definition of Done)
- `npm test` 통과, `node tools/smoke.mjs` → `bad 0`, `npm run build` 와 `npm run build:cloud` 성공
- pptx 에 영향이 있으면: 왕복 테스트 추가 + LibreOffice 로 열어 확인 (`docs/03_스킬가이드.md` §5)
- `CLAUDE.md` 또는 `docs/01_기술상세.md` 에 한 줄 기록, 새 기능은 `commands.js` 의 `WHATS_NEW`
