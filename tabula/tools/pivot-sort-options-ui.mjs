// Synthetic native-XLSX fixtures only; never sends document/API writes.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { computePivot, resolvePivot, pivotSourceData } from '../src/pivot.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = process.env.WIXEL_URL || 'http://127.0.0.1:5195/', origin = new URL(url).origin;
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw new Error('Use an isolated local test server.');
const out = process.env.WIXEL_PIVOT_SORT_OUT || `D:/Codex/Temp/wixel-upgrade-20261004/pivot-sort-${engine}`;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch(), results = [], assets = new Set();let checks = 0;
const eq = (actual, expected, label) => { assert.deepEqual(actual, expected, label);checks++; };
function fixture(values, order) {
  const wb = new Workbook();wb.sheets[0].name = 'Source';
  [['Item', 'Value'], ...values.map((v, i) => [v, i + 1])].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
  const si = wb.addSheet('Report');wb.sheets[si].pivot = { name: 'SortAudit', source: 'Source', range: { r1: 0, c1: 0, r2: values.length, c2: 1 }, rows: ['Item'], values: [{ field: 'Value', agg: 'sum' }], sort: { Item: { dir: 'asc' } }, ...(order ? { order: { Item: order } } : {}), top: 1, left: 1, layout: 'tabular' };
  return new Workbook(readXlsx(writeXlsx(wb)).data).serialize();
}
const run = (p, command) => p.evaluate(command => tabula.run(command), command);
const state = p => p.evaluate(() => tabula.wb().serialize());
const definition = p => p.evaluate(() => structuredClone(tabula.wb().sheets[1].pivot));
const labels = p => p.evaluate(() => { const w = tabula.wb(), d = w.sheets[1].pivot;return [1, 2, 3].map(i => w.getValue(1, d.top + i, d.left)); });
const dialog = (p, name) => p.getByRole('dialog', { name, exact: true });
async function menu(p) { await p.locator('.pbtn[data-k="rows"][data-f="Item"]').first().click();const m = p.locator('.pivot-filter-menu');await m.waitFor();return m; }
async function roundtrip(p, expected, enabled) {
  const wb = new Workbook(await state(p)), back = new Workbook(readXlsx(writeXlsx(wb)).data), d = back.sheets[1].pivot;
  const res = resolvePivot(pivotSourceData(back, d), d), actual = computePivot(res, res.def).grid.slice(1, -1).map(row => row[0].raw.replace(/^'/, ''));
  eq(actual, expected, 'native XLSX retains the UI result');
  if (enabled !== undefined) eq(d.customListSort !== false, enabled, 'native XLSX custom-list flag');
}
async function test(name, values, order, action) {
  const context = await browser.newContext({ viewport: { width: 1400, height: 1000 }, serviceWorkers: 'block' }), p = await context.newPage(), errors = [], writes = [], start = checks;
  p.setDefaultTimeout(15000);p.on('pageerror', e => errors.push(e.message));p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true;window.WIXEL_SKIP_START = true;localStorage.setItem('wixel.mobile-work.v1', 'off'); });
  await context.route('**/*', route => { const req = route.request(), u = new URL(req.url());if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.pathname);return route.abort(); }if (u.origin !== origin || /^\/api(?:\/|$)/.test(u.pathname)) return route.abort();return route.continue(); });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });await p.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
    for (const a of await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src')))) assets.add(a);
    await p.evaluate(data => { const t = tabula, w = t.wb();w.restore(data);t.switchSheet(1);t.selectCell(2, 1);t.run('pivotRefresh');w.undoStack = [];w.redoStack = [];window.__pivotSourceBefore = JSON.stringify(w.serialize().sheets[0].cells); }, fixture(values, order));
    await p.locator('.pbtn[data-k="rows"][data-f="Item"]').first().waitFor();
    await p.screenshot({ path: `${out}/${name}-before.png` });await action(p);
    eq(await p.evaluate(() => JSON.stringify(tabula.wb().serialize().sheets[0].cells) === window.__pivotSourceBefore), true, 'source cells unchanged');
    eq(errors, [], 'no page errors');eq(writes, [], 'no external/API writes');await p.screenshot({ path: `${out}/${name}-after.png` });
    results.push({ name, ok: true, checks: checks - start });console.log('OK ' + name);
  } catch (error) { await p.screenshot({ path: `${out}/${name}-failure.png` }).catch(() => {});results.push({ name, ok: false, error: error.stack, errors, writes, checks: checks - start });console.error('NG ' + name + ': ' + error.message); }
  finally { await context.close(); }
}
try {
  await test('imported-order-filter-descending', ['A', 'B', 'C'], ['B', 'A', 'C'], async p => {
    eq(await labels(p), ['B', 'A', 'C'], 'initial imported order is preserved');const before = await definition(p);
    const m = await menu(p);await m.getByRole('menuitem', { name: '텍스트 내림차순 정렬', exact: true }).click();
    eq(await labels(p), ['C', 'B', 'A'], 'filter-menu descending changes actual report cells');eq((await definition(p)).order?.Item, undefined, 'only explicitly sorted field snapshot invalidated');
    await roundtrip(p, ['C', 'B', 'A']);await run(p, 'undo');eq(await labels(p), ['B', 'A', 'C'], 'Undo restores original file order');eq((await definition(p)).order, before.order, 'Undo restores order metadata');
    await run(p, 'redo');eq(await labels(p), ['C', 'B', 'A'], 'Redo restores requested descending order');
  });
  await test('pivot-options-custom-list-toggle', ['Mon', 'Fri', 'Tue'], null, async p => {
    eq(await labels(p), ['Mon', 'Tue', 'Fri'], 'default custom weekday order');await run(p, 'pivotOptions');
    const d = dialog(p, '피벗 테이블 옵션');await d.getByRole('button', { name: '요약 및 필터', exact: true }).click();
    await d.getByRole('checkbox', { name: '정렬할 때 사용자 지정 목록 사용', exact: true }).uncheck();await d.getByRole('button', { name: '확인', exact: true }).click();
    eq(await labels(p), ['Fri', 'Mon', 'Tue'], 'disabling custom lists changes actual cells');eq((await definition(p)).customListSort, false, 'flag applied');await roundtrip(p, ['Fri', 'Mon', 'Tue'], false);
    await run(p, 'undo');eq(await labels(p), ['Mon', 'Tue', 'Fri'], 'Undo restores weekday order');await run(p, 'redo');eq(await labels(p), ['Fri', 'Mon', 'Tue'], 'Redo restores alphabetical order');
    await run(p, 'pivotOptions');const again = dialog(p, '피벗 테이블 옵션');await again.getByRole('button', { name: '요약 및 필터', exact: true }).click();eq(await again.getByRole('checkbox', { name: '정렬할 때 사용자 지정 목록 사용', exact: true }).isChecked(), false, 'reopened checkbox remains disabled');await again.getByRole('button', { name: '취소', exact: true }).click();
  });
  await test('sort-dialog-custom-list-toggle', ['Mon', 'Fri', 'Tue'], null, async p => {
    const m = await menu(p);await m.getByRole('menuitem', { name: '기타 정렬 옵션...', exact: true }).click();const d = dialog(p, '정렬(Item)');
    await d.getByRole('button', { name: '기타 옵션(R)...', exact: true }).click();const more = dialog(p, '기타 정렬 옵션(Item)');
    await more.getByRole('checkbox').uncheck();await more.getByRole('button', { name: '확인', exact: true }).click();
    await d.getByRole('button', { name: '확인', exact: true }).click();eq(await labels(p), ['Fri', 'Mon', 'Tue'], 'sort dialog routes its checkbox into the same engine');await roundtrip(p, ['Fri', 'Mon', 'Tue'], false);
  });
} finally {
  await browser.close();const result = { engine, url, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
  await writeFile(`${out}/result.json`, JSON.stringify(result, null, 2));console.log(JSON.stringify({ engine, cases: result.cases, passed: result.passed, checks, assets: result.assets, out }));if (result.passed !== result.cases || assets.size !== 1) process.exitCode = 1;
}
