# AGENTS.md — WIXEL(위셀) 개발 지침 (Codex 용)

WIXEL 은 **Microsoft Excel(한국어판)의 화면과 동작을 그대로 따르는 브라우저 스프레드시트**입니다.
퍼포먼스 마케팅(네이버/카카오/구글 광고 리포트) 업무에 특화된 기능을 더했습니다.
목표: "엑셀 파일을 열어 엑셀과 같은 값 · 같은 모양으로 보이고, 고쳐서 다시 엑셀로 저장해도 깨지지 않는 것".

## 먼저 읽을 문서 (순서대로)
1. `docs/codex/00_시작하기.md` — 인수인계 개요, 실행 · 검증 방법
2. `docs/codex/01_기술상세내역서.md` — 구조 · 모듈 · 데이터 모델 · 엔진
3. `docs/codex/02_보완항목내역서.md` — 남은 일 (우선순위 · 완료 기준)
4. `docs/codex/03_스킬가이드.md` — 기능을 추가하는 정해진 방법 (레시피)
5. `docs/codex/04_서버호스팅가이드.md` — 서버 API · 배포 · 보안
6. 저장소 루트의 `CLAUDE.md` 의 "Tabula" 절 — **기술 노트 원본** (엑셀 동작을 맞추며 알아낸 세부 규칙 200여 개). 그 영역을 고치기 전에 해당 항목을 반드시 찾아 읽으세요.

## 명령
```bash
cd tabula
npm start        # http://localhost:5178  (server.js: 정적 파일 + 문서 저장 API)
npm test         # node --test test/*.test.js  (DOM 없는 모듈 단위 테스트)
npm run build    # dist/index.html 한 파일 배포본 (+ dist/assets/)
node tools/smoke.mjs   # (서버 실행 중) 모든 리본 명령을 실행해 오류가 없는지 — 브라우저 필요
```

## 절대 규칙 (어기면 빌드 · 배포가 깨집니다)
- **npm 의존성 추가 금지.** 순수 ES 모듈만. zip/inflate/XML/xlsx 파서도 직접 작성되어 있습니다.
- `build.mjs` 가 모듈을 한 파일로 묶으므로 **`import { a, b } from './x.js'` 와 `export const|let|function|class` 형태만** 쓰세요.
  `export default`, `export * from`, `export { a }`, 동적 `import()`, **순환 import 금지**.
- `src/fx-*.js`(수식 함수)는 `fxcore.js` 만 import 합니다 (`formula.js` import 금지 — 순환).
- 새 수식 함수는 `src/funcinfo.js` 에 설명을 넣어야 합니다 (테스트가 검사).
- 새 시트 속성은 `workbook.js` 의 `SHEET_PROPS`(필요하면 `CALC_NEUTRAL`)에 넣고 `shiftAxis`(행/열 삽입 · 삭제)에서 위치를 조정하세요.
- 되돌리기 가능한 변경은 모두 `wb.transact(fn, meta())` 안에서. 시트 속성은 `wb.setSheetProp`.
- `wb.invalidate()`(인수 없음 = 전체 재계산, 파일의 저장값 버림)는 쓰지 마세요. `wb.invalidate(si)` 또는 `invalidateStructure()`.
- 데이터 크기 배열에 `Math.max(...arr)` · `push(...arr)` 금지 (10만 개 이상에서 스택 오버플로) → `maxOf/minOf/pushAll`(fxcore.js).
- 100만 행 전체를 도는 반복 금지 → `wb.usedRange()` / `wb.extent()` 로 범위 제한. 큰 데이터는 열 블록(`block.js`).
- 핫 경로에서 셀 접근은 `CellMap` 의 `getRC/setRC/forEachRC/col(c)` (문자열 키 `'r,c'` 조회는 8배 느림).
- 수식 AST 는 같은 모양끼리 공유됩니다 — **AST 를 절대 변경하지 마세요**. 셀의 `dr/dc` 오프셋을 반영해 참조를 읽으세요.
- 오류 값은 hot path 에서 throw 하지 말고 **return** (Chrome 에서 throw/catch 1회 ≈ 3.5µs).
- UI 문자열은 한국어, 용어는 한국어판 엑셀과 같게 (예: 피벗 테이블 분석, 조건부 서식, 선택하여 붙여넣기).

## 작업 완료 기준 (Definition of Done)
1. `npm test` 전부 통과 + 새 동작에 대한 테스트 추가 (`test/*.test.js`, DOM 없는 모듈은 Node 에서 바로 테스트).
2. UI 를 바꿨다면 서버를 띄우고 `node tools/smoke.mjs` → `bad 0`.
3. xlsx 입출력을 바꿨다면 저장 → 다시 열기 왕복 테스트 추가 (`writeXlsx` → `readXlsx`).
4. 엑셀 호환을 바꿨다면 `tools/check.mjs`(수식) · `tools/pvcmp.mjs`(피벗)로 실제 파일을 **수정 전/후 비교**해 나빠진 파일이 없어야 합니다.
5. 사용자 기능이면 `src/app.js` 의 `WHATS_NEW` 와 `CLAUDE.md`(또는 `docs/codex/01_…`)에 한 줄 기록.
6. 모든 사용자 메시지는 한국어.

## 코드 위치 빠른 안내
| 무엇 | 어디 |
|------|------|
| 앱 전체 (명령 · 대화상자 · 키보드 · 리본 연결) | `src/app.js` (약 1.6만 줄, `run(cmd)` 명령 표 · `KEYTIPS` · `WHATS_NEW`) |
| 격자 그리기 (가상화) | `src/view.js` (`GridView`) |
| 리본 탭 정의 | `src/ribbon.js` (`TABS`), 아이콘 `src/icons.js` |
| 통합 문서 모델 · 재계산 · 실행 취소 | `src/workbook.js` |
| 수식 파서/평가기 | `src/formula.js`, 함수 `src/fx-*.js`, 공용 `src/fxcore.js` |
| 서식 코드 렌더 | `src/format.js` |
| xlsx 읽기/쓰기 | `src/xlsx.js` (xlsb `xlsb.js`, xls `xls.js`, ods `ods.js`) |
| 피벗 · 큐브 | `src/pivot.js`, `src/cube.js` |
| 차트 | `src/chart.js` |
| 서버 | `server.js` |
