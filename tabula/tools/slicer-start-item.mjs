// Synthetic startItem and runtime scroll probes. No private workbook input.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const origin = new URL(url).origin;
const local = ['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname);
if (!local && process.env.WIXEL_ALLOWED_TEST_URL !== url) throw new Error('Public synthetic tests require exact WIXEL_ALLOWED_TEST_URL.');
const browserName = process.env.WIXEL_BROWSER || 'chromium';
const out = path.resolve(repo, process.env.WIXEL_SLICER_START_OUT || '.local/large-xlsx-ipad/slicer-start-item/' + browserName);
if (!/^[dD]:/.test(out)) throw new Error('Outputs must be on D:.');
const filter = process.env.WIXEL_SLICER_START_FILTER || '';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
if (!['chromium', 'webkit', 'firefox'].includes(browserName)) throw new Error('Unsupported browser.');
await mkdir(out, { recursive: true });
const browser = await pw[browserName].launch();
const results = [], pageErrors = [], remoteWrites = [], blocked = [], assets = new Set(); let checks = 0;
function eq(a, b, reason) { checks++; assert.deepEqual(a, b, reason); }
function ok(value, reason) { checks++; assert.ok(value, reason); }
function near(a, b, reason) { checks++; assert.ok(Math.abs(a - b) <= 1.1, reason + ': actual=' + a + ', expected=' + b); }
const raf = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const list = p => p.locator('.obj .sl-items').first();
function data({ n = 80, columns = 2, startItem = 7, height = 24 } = {}) {
  const cells = { '0,0': { raw: 'Group' }, '0,1': { raw: 'Value' } };
  // Reverse source encounter order; startItem refers to sorted display ordinal.
  for (let i = 0; i < n; i++) { cells[(i + 1) + ',0'] = { raw: 'Item ' + String(n - i - 1).padStart(4, '0') }; cells[(i + 1) + ',1'] = { raw: String(n - i) }; }
  const pivot = { name: 'Start Pivot', source: 'Source', range: { r1: 0, c1: 0, r2: n, c2: 1 }, rows: [], cols: [], pages: [], values: [{ field: 'Value', agg: 'sum' }], top: 18, left: 0, layout: 'tabular', subtotals: 'none', grandRows: true, grandCols: true, autofit: false, preserveFormat: true, autoRefresh: false };
  const slicer = { id: 'start-slicer', caption: '합성 시작 위치', source: { kind: 'pivot', field: 'Group', pivots: [{ sheet: 'Report', name: 'Start Pivot' }] }, x: 35, y: 20, w: 350, h: 180, columns, buttonHeight: height, gap: 3, startItem, noDataLast: false, style: 'SlicerStyleLight1' };
  return { sheets: [{ name: 'Source', cells }, { name: 'Report', cells: {}, pivot, slicers: [slicer], colWidths: { 0: 160, 1: 140 } }] };
}
async function fixture(p, options = {}) {
  await p.evaluate(data => { const t = tabula, w = t.wb(); w.restore(data); t.switchSheet(1); t.selectCell(0, 0); t.gv().setZoom(100); t.run('pivotRefresh'); t.gv().layout(); w.undoStack = []; w.redoStack = []; window.__startSource = w.sheets[0]; window.__startVersion = w.sheetVersion(0); window.__startData = data; }, data(options));
  await raf(p); await list(p).waitFor();
}
async function metrics(p) {
  return p.evaluate(() => { const t = tabula, w = t.wb(), sl = w.sheets[1].slicers[0], m = t.gv().host.slicerModel(sl), node = [...document.querySelectorAll('.obj')].find(o => o.dataset.id === sl.id)?.querySelector('.sl-items'), buttons = [...node.querySelectorAll('.sl-item')], rect = node.getBoundingClientRect();
    const normalized = Number.isInteger(sl.startItem) && sl.startItem >= 0 && sl.startItem <= 4294967295 ? Math.min(sl.startItem, m.items.length - 1) : 0;
    const row = Math.floor(normalized / sl.columns), target = buttons[row * sl.columns];
    const pitch = sl.buttonHeight + sl.gap, total = Math.max(8, 8 + Math.ceil(m.items.length / sl.columns) * pitch - sl.gap), physical = Math.min(8000000, total), logicalMax = Math.max(0, total - node.clientHeight), physicalMax = Math.max(0, physical - node.clientHeight);
    const expected = node.dataset.virtual ? (logicalMax ? Math.min(logicalMax, row * pitch) * physicalMax / logicalMax : 0) : Math.min(Math.max(0, node.scrollHeight - node.clientHeight), target ? target.offsetTop - buttons[0].offsetTop : 0);
    const visible = buttons.filter(b => { const r = b.getBoundingClientRect(); return r.bottom > rect.top + 1 && r.top < rect.bottom - 1; }).map(b => ({ text: b.textContent, index: b.dataset.slIndex === undefined ? m.items.findIndex(i => i.text === b.textContent) : Number(b.dataset.slIndex), selected: b.classList.contains('on') }));
    return { startItem: sl.startItem, count: m.items.length, columns: sl.columns, clientHeight: node.clientHeight, scrollHeight: node.scrollHeight, top: node.scrollTop, expected, physical, total, virtual: node.dataset.virtual === 'true', mounted: buttons.length, visible, displayItems: m.items.map(i => i.text), selected: m.items.filter(i => i.selected).map(i => i.text), filters: structuredClone(w.sheets[1].pivot.filters ?? {}), sourceSame: w.sheets[0] === window.__startSource && w.sheetVersion(0) === window.__startVersion, sourceFirst: w.getValue(0, 1, 0), undo: w.undoStack.length, redo: w.redoStack.length, totalValue: w.getValue(1, w.sheets[1].pivot.area.r2, w.sheets[1].pivot.area.c2) }; });
}
async function initial(p, options, info) {
  await fixture(p, options); const m = await metrics(p); info.initial = m;
  near(m.top, m.expected, 'Saved startItem begins at its display row'); eq(m.virtual, options.n > 500, 'Virtual threshold follows actual display count'); eq(m.count, options.n, 'All source items retained');
  eq(m.displayItems[0], 'Item 0000', 'Display order differs from original source encounter order'); eq(m.sourceFirst, 'Item ' + String(options.n - 1).padStart(4, '0'), 'Original source encounter order remains unchanged'); eq(m.selected.length, options.n, 'Opening does not change item selection'); ok(m.sourceSame, 'Opening preserves source sheet and version');
  if (m.virtual) ok(m.mounted < 60, 'Virtual list mounts a bounded button window');
  if (m.total > 8000000) eq(m.scrollHeight, 8000000, 'Large logical extent uses the physical scroll cap');
  return m;
}
async function keepZero(p, options, info) {
  await initial(p, options, info);
  await list(p).evaluate(n => { n.scrollTop = 0; n.dispatchEvent(new Event('scroll')); }); await raf(p);
  eq((await metrics(p)).top, 0, 'User can return to the top');
  for (const zoom of [55, 100, 200]) { await p.evaluate(z => { tabula.gv().setZoom(z); tabula.gv().renderObjectsAll(); }, zoom); await raf(p); eq((await metrics(p)).top, 0, 'Explicit zero is retained at zoom ' + zoom); }
  await p.setViewportSize({ width: 1150, height: 850 }); await raf(p); eq((await metrics(p)).top, 0, 'Explicit zero retained after viewport resize');
  await p.evaluate(() => tabula.gv().setZoom(100)); await raf(p);
  await list(p).locator('.sl-item').filter({ hasText: /^Item 0000$/ }).click(); await raf(p); const picked = await metrics(p);
  eq(picked.filters.Group, ['Item 0000'], 'Mouse selection applies one exact pivot filter'); eq(picked.totalValue, 1, 'Pivot total matches independently known source value'); eq(picked.top, 0, 'Zero retained after changed filter rendering'); eq(picked.undo, 1, 'Filter creates one Undo transaction'); ok(picked.sourceSame, 'Filtering preserves source values and version');
  await p.evaluate(() => tabula.run('undo')); await raf(p); const undone = await metrics(p); eq(undone.selected.length, options.n, 'Undo restores full item selection'); eq(undone.top, 0, 'Zero retained after Undo');
  await p.evaluate(() => tabula.run('redo')); await raf(p); const redone = await metrics(p); eq(redone.filters.Group, ['Item 0000'], 'Redo restores exact filter'); eq(redone.top, 0, 'Zero retained after Redo'); info.after = redone;
}
async function test(name, fn) {
  if (filter && !filter.split('|').some(f => name.includes(f))) return;
  const context = await browser.newContext({ viewport: { width: 1450, height: 1000 }, serviceWorkers: 'block' }); const p = await context.newPage(); p.setDefaultTimeout(25000); const info = {}, errors = []; const start = checks;
  p.on('pageerror', e => { pageErrors.push({ name, message: e.message }); errors.push(e.message); }); p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {} } }); });
  await context.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { remoteWrites.push({ name, method: r.method(), path: u.pathname }); return route.abort(); } if (u.origin !== origin || /^\/api(?:\/|$)/.test(u.pathname)) { blocked.push({ name, path: u.pathname }); return route.abort(); } return route.continue(); });
  try { await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb()); for (const src of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => n.getAttribute('src')))) assets.add(src);
    await fn(p, info); eq(errors, [], 'No browser application exception'); await p.screenshot({ path: path.join(out, name + '.png') }); results.push({ name, pass: true, checks: checks - start, info });
  } catch (error) { await p.screenshot({ path: path.join(out, name + '-failed.png') }).catch(() => {}); results.push({ name, pass: false, checks: checks - start, error: error.stack, info }); }
  finally { await context.close(); }
  console.log(JSON.stringify({ name, pass: results.at(-1).pass, checks: results.at(-1).checks, error: results.at(-1).error?.split('\n')[0] }));
}
try {
  for (const [name, options] of [
    ['ordinary-one-column', { n: 80, columns: 1, startItem: 7 }], ['ordinary-two-columns', { n: 80, columns: 2, startItem: 7 }], ['ordinary-clamp-last', { n: 80, columns: 2, startItem: 4294967295 }],
    ['ordinary-invalid-negative', { n: 80, columns: 2, startItem: -4 }], ['ordinary-invalid-fraction', { n: 80, columns: 2, startItem: 2.5 }], ['ordinary-invalid-over-uint32', { n: 80, columns: 2, startItem: 4294967296 }],
    ['virtual-three-columns', { n: 600, columns: 3, startItem: 151 }], ['virtual-clamp-last', { n: 600, columns: 2, startItem: 4294967295 }], ['virtual-physical-scroll-cap', { n: 600, columns: 1, startItem: 450, height: 20000 }],
  ]) await test(name, (p, info) => initial(p, options, info));
  await test('ordinary-zero-preserved-filter-undo-zoom', (p, info) => keepZero(p, { n: 80, columns: 2, startItem: 40 }, info));
  await test('virtual-zero-preserved-filter-undo-zoom', (p, info) => keepZero(p, { n: 600, columns: 2, startItem: 400 }, info));
  await test('new-workbook-reused-id-reapplies-start', async (p, info) => { await initial(p, { n: 80, columns: 2, startItem: 40 }, info); await list(p).evaluate(n => n.scrollTop = 0); await raf(p);
    await p.evaluate(() => { window.__startPreviousBook = tabula.wb(); void tabula.newWorkbook({ name: 'Synthetic reset', build: () => window.__startData }); });
    const confirm = p.getByRole('button', { name: '새로 만들기', exact: true }); if (await confirm.isVisible()) await confirm.click();
    await p.waitForFunction(() => tabula.wb() !== window.__startPreviousBook); await raf(p); await p.evaluate(() => { tabula.switchSheet(1); tabula.run('pivotRefresh'); window.__startSource = tabula.wb().sheets[0]; window.__startVersion = tabula.wb().sheetVersion(0); }); await raf(p);
    const m = await metrics(p); near(m.top, m.expected, 'New Workbook reapplies saved initial row for the same slicer id'); ok(m.top > 0, 'Previous Workbook zero scroll is not inherited'); info.after = m;
  });
  await test('synthetic-xlsx-sync-async-and-file-open', async (p, info) => { const raw = data({ n: 80, columns: 2, startItem: 7 }), w = new Workbook(raw), bytes = writeXlsx(w), file = path.join(out, 'synthetic-start-item.xlsx'); await writeFile(file, bytes);
    for (const [name, result] of [['sync', readXlsx(bytes)], ['streaming', await readXlsxAsync(bytes, undefined, { streamThreshold: 0 })]]) { eq(result.data.sheets[1].slicers[0].startItem, 7, name + ' XLSX reader preserves startItem'); info[name] = { startItem: result.data.sheets[1].slicers[0].startItem }; }
    await p.locator('input[type="file"]').first().setInputFiles(file); await p.waitForFunction(() => tabula.wb().sheets.some(s => s.name === 'Report' && s.slicers?.length)); await p.evaluate(() => { tabula.switchSheet(1); window.__startSource = tabula.wb().sheets[0]; window.__startVersion = tabula.wb().sheetVersion(0); }); await raf(p);
    const m = await metrics(p); near(m.top, m.expected, 'Actual synthetic XLSX open applies initial startItem'); eq(m.startItem, 7, 'Opened browser model preserves startItem'); eq(m.count, 80, 'Opened browser retains every source item'); info.opened = m;
  });
} finally { await browser.close(); const result = { syntheticOnly: true, browser: browserName, url, assets: [...assets], passed: results.filter(r => r.pass).length, failed: results.filter(r => !r.pass).length, checks, pageErrors, remoteWrites, blocked, results }; await writeFile(path.join(out, 'result.json'), JSON.stringify(result, null, 2)); console.log(JSON.stringify({ browser: browserName, passed: result.passed, failed: result.failed, checks, pageErrors: pageErrors.length, remoteWrites: remoteWrites.length, assets: result.assets })); if (result.failed || pageErrors.length || remoteWrites.length) process.exitCode = 1; }
