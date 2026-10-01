// 합성 문서만 사용: 가져온 캐시 + 다른 시트 수식 의존성의 자동/수동 새로 고침과 Undo.
// WIXEL_URL, PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH. 서버 쓰기는 차단합니다.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';

function fixture(autoRefresh) {
  const wb = new Workbook();
  wb.addSheet('입력'); wb.addSheet('보고서');
  wb.transact(() => {
    [['지역', '금액'], ['서울', '=입력!A1']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v)));
    wb.setInput(1, 0, 0, '10');
    [['지역', '합계 : 금액'], ['서울', '10'], ['총합계', '10']].forEach((row, r) => row.forEach((v, c) => wb.setInput(2, r, c, v)));
    wb.setSheetProp(2, 'pivot', { name: '자동갱신검증', source: wb.sheets[0].name, range: { r1: 0, c1: 0, r2: 1, c2: 1 }, top: 0, left: 0, area: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['지역'], cols: [], values: [{ field: '금액', agg: 'sum' }], layout: 'tabular', autoRefresh });
  });
  return writeXlsx(wb);
}

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const errors = [], writes = []; let checks = 0;
const eq = (a, b) => { assert.deepEqual(a, b); checks++; };
try {
  for (const automatic of [true, false]) {
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    page.on('pageerror', (e) => errors.push(e.message));
    await context.route('**/*', (route) => {
      if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort(); }
      return route.continue();
    });
    try {
      await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
      await page.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/');
      await page.waitForFunction(() => !!window.tabula?.wb());
      await page.locator('#fileInput').setInputFiles({ name: 'synthetic-pivot-refresh.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(fixture(automatic)) });
      await page.waitForFunction(() => window.tabula.wb().sheets[2]?.pivot?.name === '자동갱신검증');
      await page.waitForFunction(() => !document.querySelector('.progress-overlay') && !window.tabula.wb().noUndo && window.tabula.wb().listeners.size > 0);
      await page.evaluate((automatic) => {
        const t = window.tabula, wb = t.wb();
        wb.pivotSnapshots = new Map([
          ['synthetic-cache', { rows: [['지역', '금액'], ['서울', 10]], ver: wb.sheets[0]._ev ?? 0 }],
          ['unrelated-cache', { rows: [['독립'], ['보존']] }],
        ]);
        wb.transact(() => wb.setSheetProp(2, 'pivot', { ...wb.sheets[2].pivot, snapshotId: 'synthetic-cache', autoRefresh: automatic }));
        t.switchSheet(2);
      }, automatic);
      await page.waitForTimeout(700); // 첫 원본 버전 관측(350ms 디바운스)
      eq(await page.evaluate(() => window.tabula.wb().getValue(2, 1, 1)), 10);
      await page.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(1, 0, 0, '20')); });
      eq(await page.evaluate(() => window.tabula.wb().getValue(0, 1, 1)), 20);
      if (automatic) {
        await page.waitForFunction(() => window.tabula.wb().getValue(2, 1, 1) === 20, null, { timeout: 5000 }).catch(async (e) => { console.error(await page.evaluate(() => { const w = window.tabula.wb(); return { def: w.sheets[2].pivot, versions: w.sheets.map((s, i) => w.sheetVersion(i)), values: [w.getValue(0, 1, 1), w.getValue(2, 1, 1)], caches: [...(w.pivotSnapshots?.keys() ?? [])] }; })); throw e; });
        eq(await page.evaluate(() => window.tabula.wb().pivotSnapshots.has('synthetic-cache')), false);
        eq(await page.evaluate(() => window.tabula.wb().pivotSnapshots.has('unrelated-cache')), true);
        await page.evaluate(() => window.tabula.wb().undo());
        await page.waitForTimeout(700);
        eq(await page.evaluate(() => [window.tabula.wb().getValue(1, 0, 0), window.tabula.wb().getValue(2, 1, 1)]), [10, 10]);
      } else {
        await page.waitForTimeout(700);
        eq(await page.evaluate(() => window.tabula.wb().getValue(2, 1, 1)), 10);
        eq(await page.evaluate(() => window.tabula.wb().pivotSnapshots.has('synthetic-cache')), true);
        await page.evaluate(() => window.tabula.run('refreshAll'));
        await page.waitForFunction(() => window.tabula.wb().getValue(2, 1, 1) === 20, null, { timeout: 5000 });
        eq(await page.evaluate(() => window.tabula.wb().getValue(2, 1, 1)), 20);
      }
    } finally { await context.close(); }
  }
  eq(errors, []); eq(writes, []);
  console.log(JSON.stringify({ ok: true, cases: 2, checks, pageErrors: errors, blockedWrites: writes }));
} finally { await browser.close(); }
