import assert from 'node:assert/strict';
import { readFileSync, statSync, mkdirSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { createHash } from 'node:crypto';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

// Customer files are optional, read only, and stay local. Logs contain counts only.
const args = process.argv.slice(2), option = name => args.includes(name) ? args[args.indexOf(name) + 1] : null;
const actual = option('--file'), corrected = option('--corrected');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
if (actual || corrected) assert.ok(['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname), 'actual file checks use a local server');
const output = resolve(process.env.FORMULA_REPAIR_OUTPUT || '.local/iphone-formula-support/browser');
mkdirSync(output, { recursive: true });
const guard = path => { const s = statSync(path, { bigint: true }); return { size: String(s.size), mtime: String(s.mtimeNs), sha: createHash('sha256').update(readFileSync(path)).digest('hex') }; };
const guards = new Map([actual, corrected].filter(Boolean).map(path => [path, guard(path)]));
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.FORMULA_REPAIR_BROWSER || 'webkit';
assert.ok(['webkit', 'chromium'].includes(engine));
const browser = await pw[engine].launch();
let checks = 0;
const equal = (a, b) => { assert.deepEqual(a, b); checks++; };
const yes = value => { assert.ok(value); checks++; };
const seed = new Workbook({ sheets: [{ name: '검증', cells: {
  '0,0': { raw: 'red' }, '0,1': { raw: '10' }, '1,1': { raw: '30' },
  '0,2': { raw: '=IFERROR(IFS(A1="red",SUMIFSㅋ(B1:B2,A1:A2,"red"),TRUE,7),0)', cached: 0 },
  '1,2': { raw: '=IFERROR(IFS(A1="blue",SUMIFSㅋ(B1:B2,A1:A2,"red"),TRUE,7),0)', cached: 7 },
  '0,3': { raw: '=SUM(C1:C2)', cached: 7 },
} }] });
const fixture = Buffer.from(writeXlsx(seed));
try {
  const context = await browser.newContext({ ...pw.devices['iPhone 13'], acceptDownloads: true });
  const page = await context.newPage(), errors = [];
  page.on('pageerror', e => errors.push(e.message));
  await page.route('**/api/health', r => r.fulfill({ json: { ok: true, vault: false, auth: false, publish: false } }));
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.showSaveFilePicker = undefined; window.showOpenFilePicker = undefined; });
  await page.goto(url, { timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
  const importFile = async file => {
    await page.evaluate(() => { window.__formulaRepairBeforeOpen = tabula.wb(); });
    await page.locator('#fileInput').setInputFiles(file);
    await page.waitForFunction(() => tabula.wb() !== window.__formulaRepairBeforeOpen && !document.querySelector('.load-progress'), null, { timeout: 180000 });
    await page.evaluate(() => { delete window.__formulaRepairBeforeOpen; });
  };
  await importFile({ name: 'synthetic.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: fixture });
  const importing = page.getByRole('dialog', { name: '가져오기', exact: true });
  await importing.waitFor({ timeout: 60000 });
  yes((await importing.textContent()).includes('SUMIFSㅋ'));
  yes((await importing.textContent()).includes('수식 2개'));
  equal(await page.evaluate(() => { const w = tabula.wb(); return [w.getValue(0, 0, 2), w.getValue(0, 1, 2), w.getValue(0, 0, 3)]; }), [0, 7, 7]);
  await importing.getByRole('button', { name: '함수 이름 확인...', exact: true }).click();
  const review = page.getByRole('dialog', { name: '함수 이름 수정', exact: true });
  await review.waitFor();
  yes((await review.textContent()).includes('SUMIFSㅋ → SUMIFS'));
  const fitDialog = async () => {
    yes(await page.evaluate(() => [...document.querySelectorAll('.dialog')].every(d => { const r = d.getBoundingClientRect(); return r.left >= -1 && r.right <= innerWidth + 1 && r.top >= -1 && r.bottom <= innerHeight + 1; })));
  };
  await fitDialog();
  await review.getByRole('button', { name: '취소', exact: true }).click();
  equal(await page.evaluate(() => tabula.wb().getValue(0, 0, 2)), 0);
  await page.evaluate(() => tabula.run('formulaNameRepair'));
  await review.getByRole('button', { name: '수정 적용', exact: true }).click();
  equal(await page.evaluate(() => { const w = tabula.wb(); return [w.getValue(0, 0, 2), w.getValue(0, 1, 2), w.getValue(0, 0, 3)]; }), [10, 7, 17]);
  await page.evaluate(() => tabula.run('undo'));
  equal(await page.evaluate(() => tabula.wb().getValue(0, 0, 3)), 7);
  await page.evaluate(() => tabula.run('redo'));
  equal(await page.evaluate(() => tabula.wb().getValue(0, 0, 3)), 17);
  await page.evaluate(() => { void tabula.exportXlsx(); });
  const fallback = page.getByRole('dialog', { name: '파일로 저장', exact: true });
  await fallback.waitFor();
  const downloaded = page.waitForEvent('download');
  await fallback.getByRole('button', { name: '다운로드', exact: true }).click();
  const download = await downloaded;
  const path = join(output, `synthetic-${engine}.xlsx`);
  await download.saveAs(path);
  const roundtrip = readXlsx(new Uint8Array(readFileSync(path)));
  equal(roundtrip.warnings.filter(w => w.includes('확인이 필요한 함수 이름')).length, 0);
  equal(new Workbook(roundtrip.data).getValue(0, 0, 3), 17);
  await page.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setInput(0, 2, 2, '=SUMㅋ(1)')); });
  await page.evaluate(() => tabula.run('formulaNameRepair'));
  await page.evaluate(() => tabula.wb().transact(() => tabula.wb().setInput(0, 3, 0, 'new')));
  await review.getByRole('button', { name: '수정 적용', exact: true }).click();
  equal(await review.count(), 1);
  equal(await page.evaluate(() => tabula.wb().getRaw(0, 2, 2)), '=SUMㅋ(1)');
  await review.getByRole('button', { name: '취소', exact: true }).click();
  await page.evaluate(() => { tabula.wb().sheets[0].protect = { on: true }; });
  // Direct calculation-status/menu entry still previews a protected document safely.
  await page.evaluate(() => tabula.run('calculationStatus'));
  const status = page.getByRole('dialog', { name: '계산 상태 확인', exact: true });
  await status.getByRole('button', { name: '함수 이름 수정...', exact: true }).click();
  yes(await review.getByRole('button', { name: '수정 적용', exact: true }).isDisabled());
  await review.getByRole('button', { name: '취소', exact: true }).click();
  await page.evaluate(() => { tabula.wb().sheets[0].protect = undefined; tabula.selectCell(2, 0); });
  await page.locator('#formulaInput').fill('검토 중 입력');
  await page.evaluate(() => tabula.openNamedMenu('errorMenu', document.getElementById('formulaInput')));
  await page.getByText('함수 이름 수정...', { exact: true }).click();
  await review.waitFor();
  equal(await page.evaluate(() => tabula.wb().getRaw(0, 2, 0)), '검토 중 입력');
  await review.getByRole('button', { name: '취소', exact: true }).click();
  equal(await page.evaluate(() => tabula.wb().getRaw(0, 2, 0)), '검토 중 입력');
  if (actual) {
    await importFile(resolve(actual));
    await importing.waitFor({ timeout: 180000 });
    yes((await importing.textContent()).includes('SUMIFSㅋ'));
    yes((await importing.textContent()).includes('수식 372개'));
    await importing.getByRole('button', { name: '함수 이름 확인...', exact: true }).click();
    yes((await review.textContent()).includes('셀 수식 372개와 표 계산 열 수식 1개'));
    await fitDialog();
    await page.screenshot({ path: join(output, `actual-review-${engine}.png`) });
    const before = await page.evaluate(() => tabula.wb().sheets.flatMap(s => { const values = []; s.cells.forEachFormulaRC(c => { if (c.raw.includes('SUMIFSㅋ(')) values.push(c.raw); }); return values; }));
    equal(before.length, 372);
    await review.getByRole('button', { name: '수정 적용', exact: true }).click();
    equal(await page.evaluate(() => { let n = 0; tabula.wb().sheets.forEach(s => s.cells.forEachFormulaRC(c => { if (c.raw.includes('SUMIFSㅋ(')) n++; })); return n; }), 0);
    await page.evaluate(() => tabula.run('undo'));
    equal(await page.evaluate(() => tabula.wb().sheets.flatMap(s => { const values = []; s.cells.forEachFormulaRC(c => { if (c.raw.includes('SUMIFSㅋ(')) values.push(c.raw); }); return values; })), before);
    await page.evaluate(() => tabula.run('redo'));
    equal(await page.evaluate(() => { let n = 0; tabula.wb().sheets.forEach(s => s.cells.forEachFormulaRC(c => { if (c.raw.includes('SUMIFSㅋ(')) n++; })); return n; }), 0);
  }
  if (corrected) {
    await importFile(resolve(corrected));
    await page.waitForFunction(() => tabula.wb().sheets.length > 5, null, { timeout: 180000 });
    await page.waitForTimeout(2000);
    equal(await importing.count(), 0);
    equal(await page.evaluate(() => { let n = 0; tabula.wb().sheets.forEach(s => s.cells.forEachFormulaRC(c => { if (c.raw.includes('SUMIFSㅋ(')) n++; })); return n; }), 0);
    await page.screenshot({ path: join(output, `actual-repaired-${engine}.png`) });
  }
  equal(errors, []);
  await context.close();
  writeFileSync(join(output, `summary-${engine}.json`), JSON.stringify({ engine, checks, actual: !!actual, corrected: !!corrected, pageErrors: errors.length }, null, 2));
  console.log(JSON.stringify({ engine, checks, actual: !!actual, corrected: !!corrected, pageErrors: errors.length }));
} finally {
  await browser.close();
  for (const [path, before] of guards) assert.deepEqual(guard(path), before, 'read-only file guard');
}
