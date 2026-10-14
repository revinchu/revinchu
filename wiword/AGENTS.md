# AGENTS.md — WIWORD (위워드) 개발 규칙

WIWORD 는 **Microsoft Word 와 같은 브라우저 문서 편집기**입니다. WIXEL(`../tabula`, 엑셀) · WIPOINT(`../wipoint`, 파워포인트)와 같은 방식으로 만들었습니다: 순수 ES 모듈, npm 의존성 0개, 빌드하면 HTML 한 파일.

먼저 `CLAUDE.md` → `docs/00_시작하기.md` 를 읽으세요.

## 명령
```bash
cd wiword
npm start                 # http://127.0.0.1:5180 (개발 서버 + AI 연결 중계)
npm test                  # Node 내장 테스트
WIWORD_CORPUS=/폴더 npm test   # 실제 문서 모음(.docx .hwp .hwpx .doc .pdf)도 열어 봄
npm run build             # dist/index.html (한 파일) + dist/assets/
node mcp.js               # MCP 서버 (stdio) — 보통은 AI 클라이언트가 실행
```

## 절대 규칙
1. **npm 의존성 추가 금지.** zip · XML · CFB · docx · hwp · hwpx · pdf 는 직접 구현.
2. **import/export 형식**: `import { a, b } from './x.js';` 와 `export const|let|function|class|async function` 만. 동적 `import()`, `export { }`, `export default`, `import *`, 빈 `import './x.js'`(→ `import { } from './x.js'`), 다른 폴더(`../`) import, 순환 import 금지.
3. **UI 글자는 모두 한국어.**
4. **모듈끼리 직접 부르지 말고 `state.js` 를 거침**: 명령은 `register({...})` 로 등록하고 `run('이름')`, 다시 그리기는 `emit(...)`.
5. **DOM 없는 모듈**(model · render · formats · docx · hwpx · hwp · doc97 · rtf · md · csv · htmlio · pdf · pdfdoc · ops · zip · xml · cfb · math · emf · wmf · tiff · fonts · symfont · webfonts · fontarchive · sample)에는 `document`/`window` 를 쓰지 않습니다 — Node 테스트와 `mcp.js` 가 그대로 씁니다.
6. 문서 변경은 항상 `change(() => …)` 안에서 (실행 취소 기록).
7. AI 도구는 `ops.js` 의 순수 함수로 만듭니다. `mcp.js` 와 `bridge.js` 가 같은 함수를 씁니다 — 한쪽에만 기능을 넣지 마세요.
8. **사용자가 준 실제 문서(.docx · .hwp · .pdf 등)는 절대 커밋하지 않습니다.** 외부 저장소의 예제 파일(hwp-parser 등)도 마찬가지 — 로컬 시험용으로만 쓰고 `WIWORD_CORPUS` 로 지정합니다.

## 완료 기준 (Definition of Done)
- `npm test` 통과, `npm run build` 성공, 브라우저 콘솔 오류 0
- 파일 형식에 영향이 있으면: 왕복 테스트 추가 + 실제 문서 모음으로 확인
- `CLAUDE.md` 에 한 줄 기록
