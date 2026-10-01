// 합성 XLSX 8조합의 실제 옵션 UI, 값 행, 클래식 전환, 취소/Undo/Redo, 저장 왕복.
// WIXEL_URL(기본 source 5180), PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH.
// 독립 브라우저 문맥만 사용하며 서버 쓰기를 차단합니다. 업무 파일은 읽지 않습니다.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { computePivot, resolvePivot, pivotSourceData } from '../src/pivot.js';

function fixture(layout, classic, showValuesRow) {
  const wb = new Workbook();
  wb.transact(() => {
    [['지역', '품목', '매출', '비용'], ['서울', 'A', 10, 1], ['서울', 'B', 20, 2], ['부산', 'A', 30, 3]].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
    wb.setSheetProp(0, 'pivot', {
      name: '표시옵션검증', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 3, c2: 3 },
      rows: ['지역', '품목'], cols: [], values: [{ field: '매출', name: '총매출' }, { field: '비용', name: '총비용' }],
      layout, classic, showValuesRow, top: 0, left: 6, area: { r1: 0, c1: 6, r2: 9, c2: 9 },
    });
  });
  return writeXlsx(wb);
}
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const errors = [], writes = [];
let checks = 0, cases = 0;
const eq = (a, b, label) => { assert.deepEqual(a, b, label); checks++; };
try {
  for (const layout of ['compact', 'tabular']) for (const classic of [false, true]) for (const showValuesRow of [false, true]) {
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    page.on('pageerror', (e) => errors.push(e.message));
    await context.route('**/*', (route) => {
      if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort(); }
      return route.continue();
    });
    const open = async () => {
      await page.evaluate(() => { window.tabula.selectCell(1, 6); window.tabula.run('pivotOptions'); });
      await page.locator('.dialog').getByRole('button', { name: '표시', exact: true }).click();
    };
    const state = () => page.evaluate(() => {
      const wb = window.tabula.wb(), d = wb.sheets[0].pivot;
      return { layout: d.layout, classic: d.classic, showValuesRow: d.showValuesRow, book: wb.serialize() };
    });
    const rendered = async (expectedLayout, expectedClassic, expectedShow) => {
      const s = await state();
      eq([s.layout, s.classic, s.showValuesRow], [expectedLayout, expectedClassic, expectedShow], '문서 옵션');
      const wb = new Workbook(s.book), d = wb.sheets[0].pivot, res = resolvePivot(pivotSourceData(wb, d), d), output = computePivot(res, res.def);
      eq(output.meta.headerRows, expectedClassic || expectedShow ? 2 : 1, '값 행 수');
      const n = output.meta.headerRows - 1, actual = Array.from({ length: output.meta.width }, (_, c) => wb.getRaw(0, n, 6 + c));
      eq(actual, output.grid[n].map((c) => c.raw), '화면 워크북에 실제 필드 머리글 기록');
      eq(wb.getValue(0, 1, 2), 10, '원본 데이터 보존');
      return s;
    };
    try {
      await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
      await page.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/');
      await page.waitForFunction(() => !!window.tabula?.wb());
      await page.locator('#fileInput').setInputFiles({ name: 'synthetic-pivot-display.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(fixture(layout, classic, showValuesRow)) });
      await page.waitForFunction(() => window.tabula.wb().sheets[0].pivot?.name === '표시옵션검증');
      await page.waitForFunction(() => !document.querySelector('.progress-overlay') && !window.tabula.wb().noUndo && window.tabula.wb().listeners.size > 0);
      await rendered(layout, classic, showValuesRow);
      await open();
      eq(await page.getByLabel('클래식 피벗 테이블 레이아웃 (표 안으로 필드 끌어 놓기)', { exact: true }).isChecked(), classic, '가져온 클래식 체크');
      eq(await page.getByLabel('값 행 표시', { exact: true }).isChecked(), showValuesRow, '가져온 값 행 체크');
      await page.getByLabel('값 행 표시', { exact: true }).setChecked(!showValuesRow);
      await page.locator('.dialog').getByRole('button', { name: '취소', exact: true }).click();
      await rendered(layout, classic, showValuesRow);
      await open();
      await page.getByLabel('값 행 표시', { exact: true }).setChecked(!showValuesRow);
      await page.locator('.dialog').getByRole('button', { name: '확인', exact: true }).click();
      await rendered(layout, classic, !showValuesRow);
      await page.evaluate(() => window.tabula.wb().undo());
      await rendered(layout, classic, showValuesRow);
      await page.evaluate(() => window.tabula.wb().redo());
      await rendered(layout, classic, !showValuesRow);
      await open();
      await page.getByLabel('클래식 피벗 테이블 레이아웃 (표 안으로 필드 끌어 놓기)', { exact: true }).setChecked(!classic);
      await page.locator('.dialog').getByRole('button', { name: '확인', exact: true }).click();
      const nextLayout = classic ? layout : 'tabular';
      const s = await rendered(nextLayout, !classic, !showValuesRow);
      await open();
      eq(await page.getByLabel('값 행 표시', { exact: true }).isChecked(), !showValuesRow, '실제 헤더와 독립인 옵션 재열기');
      await page.locator('.dialog').getByRole('button', { name: '취소', exact: true }).click();
      const back = new Workbook(readXlsx(writeXlsx(new Workbook(s.book))).data), d = back.sheets[0].pivot;
      eq([d.layout, d.classic, d.showValuesRow], [nextLayout, !classic, !showValuesRow], '편집 후 XLSX 왕복');
      cases++;
    } finally { await context.close(); }
  }
  eq(errors, [], '브라우저 오류'); eq(writes, [], '서버 쓰기');
  console.log(JSON.stringify({ ok: true, cases, checks, pageErrors: errors, blockedWrites: writes }));
} finally { await browser.close(); }
