// 합성 XLSX를 실제 파일 열기 UI로 가져와 오류 표시 옵션·확인 후 결과·다시 저장을 검사합니다.
// source/cloud 공통: WIXEL_URL, PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH 사용. 사용자 파일·서버 쓰기 없음.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { pivotErrorDisplay } from '../src/pivot.js';

function fixture(flag, caption) {
  const wb = new Workbook();
  wb.transact(() => [['지역', '비용', '클릭'], ['오류행', '10', '0'], ['정상행', '30', '3']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v))));
  wb.transact(() => wb.setSheetProp(0, 'pivot', {
    name: '오류옵션검증', source: 'Sheet1', range: { r1: 0, c1: 0, r2: 2, c2: 2 }, rows: ['지역'], cols: [], values: [{ field: 'CPC', agg: 'sum' }],
    calcFields: [{ name: 'CPC', formula: '비용/클릭' }], layout: 'tabular', top: 0, left: 5, area: { r1: 0, c1: 5, r2: 3, c2: 6 },
  }));
  const files = unzip(writeXlsx(wb)), key = 'xl/pivotTables/pivotTable1.xml';
  files[key] = textOf(files[key]).replace(/ showError="[^"]*"| errorCaption="[^"]*"/g, '').replace('<pivotTableDefinition ', '<pivotTableDefinition ' + (flag === null ? '' : 'showError="' + flag + '" ') + (caption === null ? '' : 'errorCaption="' + caption + '" '));
  return zip(files);
}
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const errors = [], writes = [];
let cases = 0, checks = 0;
try {
  for (const [flag, caption, enabled] of [[null, null, false], [null, '-', false], ['0', '', false], ['0', '-', false], ['1', null, true], ['1', '', true], ['1', '-', true]]) {
    const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
    const page = await context.newPage();
    page.on('pageerror', (e) => errors.push(e.message));
    await context.route('**/*', (route) => {
      if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort('blockedbyclient'); }
      return route.continue();
    });
    try {
      await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
      await page.goto(process.env.WIXEL_URL || 'http://localhost:5178/');
      await page.waitForFunction(() => !!window.tabula?.wb());
      await page.locator('#fileInput').setInputFiles({ name: 'synthetic-pivot-options.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(fixture(flag, caption)) });
      await page.waitForFunction(() => window.tabula.wb().sheets[0].pivot?.name === '오류옵션검증' && /synthetic-pivot-options\.xlsx.*열었습니다/.test(document.getElementById('toast')?.textContent || '')); // 이름 설정 뒤 비동기 afterLoad까지 완료

      await page.evaluate(() => { window.tabula.selectCell(1, 5); window.tabula.run('pivotOptions'); });
      assert.equal(await page.getByLabel('오류 값 표시', { exact: true }).isChecked(), enabled); checks++;
      const options = page.getByRole('dialog', { name: '피벗 테이블 옵션', exact: true });
      await options.getByRole('button', { name: '확인', exact: true }).click();
      await options.waitFor({ state: 'detached' });
      const result = await page.evaluate(() => {
        const wb = window.tabula.wb(), def = wb.sheets[0].pivot, value = wb.getValue(0, 1, 6);
        return { enabled: def.errorShow, caption: def.errorCaption, value: value?.code ?? value, book: wb.serialize() };
      });
      assert.equal(result.enabled, enabled); checks++;
      assert.equal(result.value, enabled ? caption === '-' ? 0 : null : '#DIV/0!'); checks++;
      assert.equal(result.caption, caption ?? ''); checks++;
      const back = new Workbook(readXlsx(writeXlsx(new Workbook(result.book))).data);
      assert.equal(pivotErrorDisplay(back.sheets[0].pivot), enabled); checks++;
      assert.equal(back.sheets[0].pivot.errorCaption, result.caption); checks++;
      cases++;
    } finally { await context.close(); }
  }
  assert.deepEqual(errors, []); assert.deepEqual(writes, []);
  console.log(JSON.stringify({ ok: true, cases, checks, pageErrors: errors, blockedWrites: writes }));
} finally { await browser.close(); }
