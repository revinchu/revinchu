# 검증 도구 (tools/)

`npm test`(Node 단위 테스트) 외에, 실제 브라우저와 실제 엑셀 파일로 회귀를 잡는 스크립트들입니다.
엑셀 호환을 고칠 때는 **수정 전/후로 같은 파일을 돌려 숫자가 나빠지지 않았는지** 비교하는 것이 원칙입니다.

## 준비
```bash
cd tabula
npm i -D playwright      # 이 폴더의 스크립트만 씀 (앱 자체는 의존성 0 — package.json 에 넣지 말 것: --no-save 권장)
npx playwright install chromium
npm start &              # http://localhost:5178
```
- 다른 주소면 `WIXEL_URL=http://localhost:8080/`
- 전역 설치된 playwright 를 쓰려면 `PLAYWRIGHT_MODULE=/경로/playwright/index.mjs`
- 실제 업무 파일(xlsx/xlsb)은 저장소에 **커밋하지 마세요.** 로컬 폴더(예: `~/wixel-samples/`)에 두고 경로로 넘깁니다.

## 스크립트
| 파일 | 하는 일 | 사용법 | 성공 기준 |
|------|---------|--------|-----------|
| `smoke.mjs` | 앱의 모든 명령(`window.tabula.commands()`, 약 280개)을 샘플 데이터 위에서 실행해 콘솔 오류 · 예외 수집 | `node tools/smoke.mjs` | 마지막 줄 `bad 0` |
| `keys.mjs` | 엑셀에서 많이 쓰는 단축키 66개를 눌러 결과 확인 | `node tools/keys.mjs` | 실패 목록 없음 (Ctrl+Shift+1 은 한국어판처럼 `#,##0` 이 의도된 동작) |
| `check.mjs` | (Node, 브라우저 불필요) 파일의 모든 수식을 다시 계산해 **엑셀이 저장한 값과 비교** + xlsx 저장/다시 읽기 왕복 검사 | `node tools/check.mjs 파일.xlsx` | `formula mismatches: 0`, `roundtrip issues: 0` (TODAY/NOW/RAND 처럼 날짜 · 난수에 따라 바뀌는 수식은 예외) |
| `brcheck.mjs` | 같은 비교를 브라우저에서 (피벗을 다시 그린 뒤라 GETPIVOTDATA 등 포함) + 여는 시간 | `node tools/brcheck.mjs 파일.xlsx` | 불일치 0, 로딩 시간 기록 |
| `pvcmp.mjs` | 피벗 영역마다 **엑셀이 저장한 칸 값 vs WIXEL 이 다시 계산해 그린 값** | `node tools/pvcmp.mjs 파일.xlsx` | `pivots N differing 0` |

> `check.mjs` · `pvcmp.mjs` 는 xlsx/xlsm/xlsb 만 받습니다. 옛 `.xls`(BIFF8)는 `readXls` 를 쓰므로 브라우저(`brcheck.mjs`)로 확인하세요.

## 회귀 비교 방법 (권장)
```bash
for f in ~/wixel-samples/*.xls[xmb]; do echo "== $f"; node tools/pvcmp.mjs "$f" | tail -1; done > after.txt
git stash; (서버는 파일을 바로 읽으므로 재시작 불필요) 같은 명령 > before.txt; git stash pop
diff before.txt after.txt
```
나빠진 파일이 하나라도 있으면 원인을 찾을 때까지 커밋하지 않습니다.

## 마지막 측정값 (2026-10-01, 사용자 실제 파일 · 저장소에 없음)
- 피벗이 있는 실제 파일 50개: 대부분 `differing 0`.
- 남은 차이: g2.xlsb 3/30, m3.xlsb 2/19, s3.xlsx 1/1, v37.xlsx 6/56, w3.xlsb 1/8, x3.xlsx 1/12, y60.xlsb 27/62, y61.xlsb 5/10, y62.xlsb 3/8, y63.xlsb 4/5
  → `docs/codex/02_보완항목내역서.md` P1-1 참고 (xlsb 피벗이 대부분).
