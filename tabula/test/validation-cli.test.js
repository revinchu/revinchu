// 진단 CLI의 실제 프로세스 종료 코드와 입력 처리를 검사합니다.
import { test, after } from 'node:test';
import { strict as assert } from 'node:assert';
import { mkdtempSync, writeFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { spawnSync } from 'node:child_process';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
import { zip, unzip } from '../src/zip.js';

const tempRoot = resolve(tmpdir());
const dir = mkdtempSync(join(tempRoot, 'wixel-cli-'));
after(() => {
  assert.ok(resolve(dir).startsWith(tempRoot + sep));
  rmSync(dir, { recursive: true, force: true });
});
const run = (name, args = [], env = {}) => spawnSync(process.execPath, [fileURLToPath(new URL(`../tools/${name}.mjs`, import.meta.url)), ...args], {
  cwd: dir, encoding: 'utf8', timeout: 20000, env: { ...process.env, ...env },
});
const wb = new Workbook();
wb.transact(() => {
  wb.setInput(0, 0, 0, '7');
  wb.setInput(0, 1, 0, '5');
  wb.setInput(0, 2, 0, '=SUM(A1:A2)');
});
const good = join(dir, '합성 검증.XLSX');
const bytes = writeXlsx(wb);
writeFileSync(good, bytes);

test('검증 CLI: 인자 누락·옛 xls·파일 없음은 실행 오류로 종료', () => {
  for (const name of ['check', 'pvcmp']) {
    for (const [args, message] of [[[], /사용법/], [['옛 파일.xls'], /brcheck/], [[join(dir, '없음.xlsx')], /검사를 완료하지 못했습니다/]]) {
      const r = run(name, args, { PLAYWRIGHT_MODULE: '설치되지-않은-패키지' });
      assert.equal(r.status, 2, `${name}: ${r.stderr}`);
      assert.match(r.stderr, message);
      assert.doesNotMatch(r.stderr, /ERR_MODULE_NOT_FOUND/);
    }
  }
});

test('check CLI: 다른 작업 폴더·한글/공백 경로·대문자 확장자에서 합성 파일 통과', () => {
  const r = run('check', [good]);
  assert.equal(r.status, 0, r.stderr);
  assert.match(r.stdout, /formulas 1/);
  assert.match(r.stdout, /formula mismatches: 0/);
  assert.match(r.stdout, /roundtrip issues: 0/);
});

test('check CLI: 엑셀 저장값을 바꾼 합성 파일은 차이를 보고하고 실패', () => {
  const parts = unzip(bytes);
  const xml = new TextDecoder().decode(parts['xl/worksheets/sheet1.xml']);
  const changed = xml.replace('<v>12</v>', '<v>99</v>');
  assert.notEqual(changed, xml);
  parts['xl/worksheets/sheet1.xml'] = changed;
  const file = join(dir, '저장값 불일치.xlsx');
  writeFileSync(file, zip(parts));
  const r = run('check', [file]);
  assert.equal(r.status, 1, r.stderr);
  assert.match(r.stdout, /formula mismatches: 1/);
  assert.match(r.stdout, /app=12 xl=99/);
});

test('pvcmp CLI: 원본 피벗 누락·출력 영역 축소를 성공으로 처리하지 않음', () => {
  const pivotWb = new Workbook();
  pivotWb.transact(() => {
    [['매체', '비용'], ['네이버', '7'], ['구글', '5']].forEach((row, r) => row.forEach((v, c) => pivotWb.setInput(0, r, c, v)));
    pivotWb.setSheetProp(0, 'pivot', { name: '피벗1', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['매체'], cols: [], pages: [], filters: {}, values: [{ field: '비용', agg: 'sum' }], top: 0, left: 3, area: { r1: 0, c1: 3, r2: 3, c2: 4 } });
  });
  const file = join(dir, '합성 피벗.xlsx');
  writeFileSync(file, writeXlsx(pivotWb));
  // 브라우저 출력 경계만 대체합니다. 원본 XLSX 읽기와 비교/종료 코드는 실제 CLI를 실행합니다.
  const stub = join(dir, '브라우저 대역.mjs');
  writeFileSync(stub, `let calls = 0;
export const chromium = { launch: async () => ({
  newPage: async () => ({ on() {}, addInitScript() {}, goto() {}, reload() {}, waitForTimeout() {}, setInputFiles() {}, waitForFunction() {},
    evaluate: async () => ++calls === 1 ? undefined : JSON.parse(process.env.WIXEL_TEST_AREAS) }),
  close: async () => {}
}) };`);
  for (const [areas, message] of [
    [[], /피벗 또는 출력 영역이 누락/],
    [[{ si: 0, sheet: 'Sheet1', name: '피벗1', area: { r1: 0, c1: 3, r2: 2, c2: 4 }, vals: [] }], /피벗 영역 변경/],
  ]) {
    const r = run('pvcmp', [file], { PLAYWRIGHT_MODULE: stub, WIXEL_TEST_AREAS: JSON.stringify(areas) });
    assert.equal(r.status, 1, r.stderr);
    assert.match(r.stdout, message);
    assert.match(r.stdout, /differing 1/);
  }
});
