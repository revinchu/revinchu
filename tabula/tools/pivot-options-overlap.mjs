// Native synthetic files and optional private original; all browser contexts are isolated.
// Private workbooks are opened locally only, never modified, uploaded, or sent to APIs.
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile, stat } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { computePivot, resolvePivot, pivotSourceData } from '../src/pivot.js';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const parsed = new URL(url), origin = parsed.origin;
const local = ['localhost', '127.0.0.1', '[::1]'].includes(parsed.hostname);
const real = process.env.REAL_WORKBOOK || '';
if (real && !local) throw new Error('Private workbooks require a local test URL.');
if (!local && process.env.WIXEL_ALLOWED_TEST_URL !== url) throw new Error('Public synthetic testing requires an exact WIXEL_ALLOWED_TEST_URL.');
const out = path.resolve(repo, process.env.PIVOT_LAYOUT_OUTPUT || '.local/pivot-options-overlap/source');
if (!/^[dD]:/.test(out)) throw new Error('All outputs must be on D:.');
await mkdir(out, { recursive: true });
const only = process.env.PIVOT_LAYOUT_FILTER || '';
const baseline = process.env.PIVOT_LAYOUT_BASELINE || '';
const baselineSources = new Map();
if (baseline) for (const name of ['app.js', 'xlsx.js', 'xlsb.js']) baselineSources.set('/src/' + name,
  execFileSync('git', ['-c', 'safe.directory=' + path.resolve(repo, '..').replaceAll(String.fromCharCode(92), '/'), '-C', repo, 'show', baseline + ':tabula/src/' + name], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }));

const refreshLeakBaseline = process.env.PIVOT_LAYOUT_REFRESH_LEAK_BASELINE === '1';
if (refreshLeakBaseline) {
  if (!local || baseline) throw new Error('The single-line refresh leak baseline requires a local source server without a Git baseline.');
  const current = await readFile(path.join(repo, 'src/app.js'), 'utf8');
  const fixedLine = /const cubes = targets\.map\(e => pivotSourceData\(wb, e\.def, \{ preserveSnapshot: true \}\)\?\.cube\)\.filter\(Boolean\);/;
  if (!fixedLine.test(current)) throw new Error('The fixed slicerRefresh source line was not found.');
  baselineSources.set('/src/app.js', current.replace(fixedLine, 'const cubes = targets.map(e => pivotSource(e.def)?.cube).filter(Boolean);'));
}

const pw = await import(process.env.PLAYWRIGHT_MODULE || 'file:///C:/Users/system777/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs');
const engine = process.env.WIXEL_BROWSER || 'chromium';
const browser = await pw[engine].launch(), results = [], assets = new Set();
let checks = 0;
const eq = (actual, expected, label) => { checks++; assert.deepEqual(actual, expected, label); };
const ok = (condition, label) => { checks++; assert.ok(condition, label); };
const run = (p, command) => p.evaluate(command => tabula.run(command), command);
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const GROUPS = ['Alpha group with a long report label', 'Beta group with a long report label'];
const HEADER = ['Group', 'Subgroup', 'Channel', 'Value', 'Cost', 'Filter'];
const SOURCE_ROWS = GROUPS.flatMap((g, gi) => ['One', 'Two'].flatMap((sub, sj) =>
  ['Web', 'Shop'].map((channel, ci) => [g, sub, channel, 100 + gi * 40 + sj * 10 + ci, 30 + gi + sj + ci, 'All'])));
function def(name, patch = {}) {
  return { name, source: 'Source', range: { r1: 0, c1: 0, r2: SOURCE_ROWS.length, c2: HEADER.length - 1 },
    rows: ['Group'], cols: [], values: [{ field: 'Value', agg: 'sum' }], top: 2, left: 1,
    layout: 'tabular', subtotals: 'none', autofit: false, preserveFormat: true, ...patch };
}
function gridOf(w, d) { const res = resolvePivot(pivotSourceData(w, d), d); return computePivot(res, res.def).grid; }
function bounds(w, d) {
  const grid = gridOf(w, d), cols = Math.max(...grid.map(row => row.length));
  return { r1: d.top ?? 0, c1: d.left ?? 0, r2: (d.top ?? 0) + grid.length - 1, c2: (d.left ?? 0) + cols - 1 };
}
function paint(w, si, d) {
  const grid = gridOf(w, d);d.area = bounds(w, d);
  grid.forEach((row, r) => row.forEach((cell, c) => { w.setInput(si, d.top + r, d.left + c, String(cell?.raw ?? '')); if (cell?.style) w.setStyle(si, d.top + r, d.left + c, cell.style); }));
}
function fixture({ autofit = false, expansion = null, neighbor = null, filtered = false, grandRows = true, multiSlicer = false, batchRefresh = false } = {}) {
  const w = new Workbook(); w.sheets[0].name = 'Source';
  [HEADER, ...SOURCE_ROWS].forEach((row, r) => row.forEach((v, c) => w.setInput(0, r, c, String(v))));
  const si = w.addSheet('Report'), d = def('MainPivot', { autofit, grandRows, ...(filtered ? { filters: { Group: [GROUPS[0]] } } : {}) });
  w.sheets[si].pivot = d;
  const current = bounds(w, d), grown = bounds(w, { ...d, ...(expansion === 'rows' ? { rows: ['Group', 'Subgroup'] } :
    expansion === 'cols' ? { cols: ['Channel'] } : expansion === 'values' ? { values: [...d.values, { field: 'Cost', agg: 'sum' }] } :
    expansion === 'pages' ? { pages: ['Filter'] } : expansion === 'filter' ? { filters: {} } :
    expansion === 'grand' ? { grandRows: true } : {}) });
  let neighborSi = si, nd = null;
  if (neighbor) {
    let top = d.top, left = d.left;
    if (neighbor === 'right') left = current.c2 + 1;
    else if (neighbor === 'below') top = current.r2 + 1;
    else if (neighbor === 'adjacent-right') left = grown.c2 + 1;
    else if (neighbor === 'adjacent-below') top = grown.r2 + 1;
    else if (neighbor === 'far') { top = 24; left = 15; }
    else if (neighbor === 'other-sheet') neighborSi = w.addSheet('Other Report');
    nd = def('NeighborPivot', { top, left, rows: ['Subgroup'], values: [{ field: 'Cost', agg: 'sum' }] });
    if (neighborSi === si) w.sheets[si].pivotsExtra = [nd]; else w.sheets[neighborSi].pivot = nd;
  }
  paint(w, si, d); if (nd) paint(w, neighborSi, nd);
  if (multiSlicer) {
    const linkedSi = w.addSheet('Linked Report'), linked = def('SafeLinkedPivot', { filters: { Group: [GROUPS[0]] } });
    w.sheets[linkedSi].pivot = linked;paint(w, linkedSi, linked);
    w.sheets[si].slicers = [{ id: 'linked-group', caption: 'Group', source: { kind: 'pivot', field: 'Group',
      pivots: [{ sheet: 'Linked Report', name: 'SafeLinkedPivot' }, { sheet: 'Report', name: 'MainPivot' }] },
      x: 520, y: 180, w: 370, h: 130, style: 'SlicerStyleLight1', columns: 1 }];
  }
  if (batchRefresh) {
    const early = def('EarlierRefreshPivot', { top: 2, left: 10, autofit: true });
    w.sheets[0].pivot = early;paint(w, 0, early);
    for (let c = 10; c < 15; c++) w.setColWidth(0, c, 47);
  }
  for (let c = 0; c < 22; c++) { w.setColWidth(si, c, 47); if (neighborSi !== si) w.setColWidth(neighborSi, c, 63); }
  return { buffer: Buffer.from(writeXlsx(w)), current, grown, neighborSi, native: new Workbook(readXlsx(writeXlsx(w)).data).serialize() };
}
const allDefs = p => p.evaluate(() => tabula.wb().sheets.flatMap((sh, si) => [sh.pivot, ...(sh.pivotsExtra ?? [])].filter(Boolean).map(d => ({ si, sheet: sh.name, name: d.name, autofit: d.autofit !== false, preserveFormat: d.preserveFormat !== false, area: structuredClone(d.area), rows: d.rows, cols: d.cols }))));
const bookState = p => p.evaluate(() => ({ sheets: tabula.wb().serialize().sheets.map(({ fileValues, ...sheet }) => sheet), undo: tabula.wb().undoStack.length, redo: tabula.wb().redoStack.length }));
const sheetState = (p, si = 1) => p.evaluate(si => { const w = tabula.wb();return { sheet: (() => { const { fileValues, ...sheet } = w.serialize().sheets[si];return sheet; })(), widths: Array.from({ length: 22 }, (_, c) => w.colWidth(si, c)), definition: structuredClone(w.sheets[si].pivot) }; }, si);

const snapshotState = p => p.evaluate(() => [...(tabula.wb().pivotSnapshots ?? [])].map(([id, snap]) => ({
  id, ver: snap.ver, sourceKey: snap.sourceKey, sourceSheet: tabula.wb().sheets.indexOf(snap.sourceSheet),
  watchVersion: snap.sourceWatch?.version, watchRange: structuredClone(snap.sourceWatch?.range), rows: JSON.stringify(snap.rows),
})));

const clearHistory = p => p.evaluate(() => { tabula.wb().undoStack = []; tabula.wb().redoStack = []; });
async function select(p, name = 'MainPivot', sheetName = 'Report') {
  await p.evaluate(({ name, sheetName }) => { const t = tabula, w = t.wb(), si = w.sheetIndexByName(sheetName), sh = w.sheets[si], d = [sh.pivot, ...(sh.pivotsExtra ?? [])].find(d => d?.name === name); t.switchSheet(si);t.selectCell(d.area?.r1 ?? d.top ?? 0, d.area?.c1 ?? d.left ?? 0);t.gv().setZoom(100);t.gv().renderAll(); }, { name, sheetName });
  await raf(p);
}
async function options(p, enabled, { confirm = true } = {}) {
  await run(p, 'pivotOptions');const d = p.getByRole('dialog', { name: '피벗 테이블 옵션', exact: true });
  const c = d.getByRole('checkbox', { name: '업데이트 시 열 자동 맞춤', exact: true });
  const actual = await c.isChecked();
  if (enabled !== undefined) await c.setChecked(enabled);
  await d.getByRole('button', { name: confirm ? '확인' : '취소', exact: true }).click();
  await d.waitFor({ state: 'detached' });await raf(p);return actual;
}
async function resetWidths(p, width = 47) {
  await p.evaluate(width => { const t = tabula, w = t.wb(); w.transact(() => { for (let c = 0; c < 22; c++) w.setColWidth(t.si, c, width); });t.gv().renderAll();w.undoStack = [];w.redoStack = []; }, width);await raf(p);
}
async function rowMenu(p, field = 'Group') {
  await p.locator('.pbtn[data-k="rows"][data-f="' + field + '"]').first().click();
  const m = p.locator('.pivot-filter-menu');await m.waitFor();return m;
}
async function sort(p) { const m = await rowMenu(p); await m.getByRole('menuitem', { name: '텍스트 내림차순 정렬', exact: true }).click();await raf(p); }
async function filter(p) { const m = await rowMenu(p); await m.getByRole('searchbox').fill('Alpha');await m.getByRole('button', { name: '확인', exact: true }).click();await raf(p); }
async function clearFilter(p) { const m = await rowMenu(p);await m.getByRole('menuitem', { name: '"Group"에서 필터 해제', exact: true }).click();await raf(p); }
async function addField(p, field, destination) {
  if (!(await p.locator('#pivotPane').isVisible())) await run(p, 'pivotFieldList');
  const node = p.locator('#pivotPane .pp-field').filter({ has: p.locator('[data-pivot-field="' + field + '"]') }).locator('.pp-fname');
  await node.click({ button: 'right' });
  await p.getByRole('menuitem', { name: new RegExp({ rows: '행 레이블에 추가', cols: '열 레이블에 추가', values: '값에 추가', pages: '보고서 필터에 추가' }[destination]) }).click();await raf(p);
}
async function changeGrand(p) {
  await run(p, 'pivotOptions');const d = p.getByRole('dialog', { name: '피벗 테이블 옵션', exact: true });
  await d.getByRole('button', { name: '요약 및 필터', exact: true }).click();
  await d.getByRole('checkbox', { name: '행 총합계 표시', exact: true }).check();
  await d.getByRole('button', { name: '확인', exact: true }).click();await raf(p);
}
async function openSynthetic(p, f, name) {
  await p.locator('#fileInput').setInputFiles({ name: 'synthetic-' + name + '.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: f.buffer });
  await p.waitForFunction(() => tabula.wb().sheets.some(sh => sh.pivot?.name === 'MainPivot') && /synthetic-.*\.xlsx.*열었습니다/.test(document.getElementById('toast')?.textContent || ''), null, { timeout: 60000 });
  await select(p);await clearHistory(p);
}
async function test(name, setup, action) {
  if (only && !name.includes(only)) return;
  const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, serviceWorkers: 'block' }), p = await context.newPage(), errors = [], writes = [], start = checks;
  p.setDefaultTimeout(15000);p.on('pageerror', e => errors.push(e.message));p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true;window.WIXEL_SKIP_START = true;localStorage.setItem('wixel.mobile-work.v1', 'off'); });
  await context.route('**/*', route => { const req = route.request(), u = new URL(req.url());if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.pathname);return route.abort(); }if (u.origin !== origin || /^\/api(?:\/|$)/.test(u.pathname)) return route.abort();if (baselineSources.has(u.pathname)) return route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: baselineSources.get(u.pathname) });return route.continue(); });
  const detail = {};
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });await p.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
    for (const a of await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src')))) assets.add(a);
    await setup(p, detail);await p.screenshot({ path: path.join(out, name + '-before.png') });await action(p, detail);
    eq(errors, [], 'no browser errors');eq(writes, [], 'no network document/API writes');
    await p.screenshot({ path: path.join(out, name + '-after.png') });results.push({ name, ok: true, checks: checks - start, ...detail });console.log('OK ' + name);
  } catch (error) {
    await p.screenshot({ path: path.join(out, name + '-failure.png') }).catch(() => {});
    results.push({ name, ok: false, checks: checks - start, error: error.stack, errors, writes, ...detail });console.error('NG ' + name + ': ' + error.message.split('\n')[0]);
  } finally { await context.close(); }
}
try {
  for (const enabled of [false, true]) for (const operation of ['sort', 'filter', 'refresh']) {
    const name = 'autofit-' + enabled + '-' + operation;
    await test(name, async (p, info) => { const f = fixture({ autofit: enabled });await openSynthetic(p, f, name);info.imported = await allDefs(p);eq(await options(p, undefined, { confirm: false }), enabled, 'native XLSX option checkbox');await resetWidths(p); },
      async (p, info) => {
        if (operation === 'refresh') { await p.evaluate(() => { const w = tabula.wb();w.transact(() => w.setInput(0, 1, 3, '999'));w.undoStack = [];w.redoStack = []; }); }
        const before = await bookState(p), a = await sheetState(p);
        if (operation === 'sort') await sort(p);else if (operation === 'filter') await filter(p);else { await run(p, 'pivotRefresh');await raf(p); }
        const after = await sheetState(p), history = await bookState(p);
        info.widths = { before: a.widths, after: after.widths };
        if (enabled) ok(after.widths.some((v, c) => v > a.widths[c] + 1), 'enabled autofit widens an overflowing pivot column');
        else eq(after.widths, a.widths, 'disabled autofit preserves every report column width');
        if (operation === 'sort') eq(after.definition.sort?.Group?.dir, 'desc', 'descending sort applied');
        if (operation === 'filter') eq(after.definition.filters?.Group, [GROUPS[0]], 'filter applied');
        if (operation === 'refresh') ok(JSON.stringify(after.sheet.cells) !== JSON.stringify(a.sheet.cells), 'refresh updates source-dependent report values');
        eq(history.sheets[0], before.sheets[0], 'source cells remain unchanged by pivot operation');eq(history.undo, before.undo + 1, 'operation has exactly one Undo entry');
        await run(p, 'undo');eq((await bookState(p)).sheets, before.sheets, 'Undo restores pivot cells, definitions, formats and dimensions');
        await run(p, 'redo');eq((await bookState(p)).sheets, history.sheets, 'Redo restores complete operation');
      });
  }
  await test('options-enable-disable-autofit', async (p, info) => { await openSynthetic(p, fixture(), 'options-toggle');info.before = (await allDefs(p))[0]; },
    async (p, info) => { const before = await bookState(p);eq(await options(p, true), false, 'initial checkbox disabled');let current = await sheetState(p);ok(current.widths.some(v => v > 48), 'enabling autofit immediately applies column fit');eq(await options(p, undefined, { confirm: false }), true, 'enabled checkbox persists');await run(p, 'undo');eq((await bookState(p)).sheets, before.sheets, 'Undo option restores dimensions and metadata');await options(p, true);eq(await options(p, false), true, 'checkbox can be disabled');await resetWidths(p);await sort(p);current = await sheetState(p);eq(current.widths, Array(22).fill(47), 'disabling option prevents later sort resize');eq(await options(p, undefined, { confirm: false }), false, 'disabled checkbox persists');info.after = current.definition; });

  const scenarios = [
    { name: 'row-expansion-right-blocked', expansion: 'rows', neighbor: 'right', blocked: true },
    { name: 'row-expansion-below-blocked', expansion: 'rows', neighbor: 'below', blocked: true },
    { name: 'column-expansion-right-blocked', expansion: 'cols', neighbor: 'right', blocked: true },
    { name: 'value-expansion-right-blocked', expansion: 'values', neighbor: 'right', blocked: true },
    { name: 'filter-expansion-below-blocked', expansion: 'filter', neighbor: 'below', filtered: true, blocked: true },
    { name: 'grand-expansion-below-blocked', expansion: 'grand', neighbor: 'below', grandRows: false, blocked: true },
    { name: 'row-expansion-boundary-adjacent', expansion: 'rows', neighbor: 'adjacent-below' },
    { name: 'column-expansion-boundary-adjacent', expansion: 'cols', neighbor: 'adjacent-right' },
    { name: 'row-expansion-far-neighbor', expansion: 'rows', neighbor: 'far' },
    { name: 'column-expansion-other-sheet', expansion: 'cols', neighbor: 'other-sheet' },
  ];
  for (const scenario of scenarios) await test(scenario.name, async (p, info) => { const f = fixture(scenario);await openSynthetic(p, f, scenario.name);info.predicted = { current: f.current, grown: f.grown, neighborSi: f.neighborSi };info.imported = await allDefs(p); },
    async (p, info) => {
      const before = await bookState(p), original = await sheetState(p), snapshotsBefore = await snapshotState(p);
      if (scenario.expansion === 'filter') await clearFilter(p);
      else if (scenario.expansion === 'grand') await changeGrand(p);
      else await addField(p, { rows: 'Subgroup', cols: 'Channel', values: 'Cost', pages: 'Filter' }[scenario.expansion], scenario.expansion);
      const after = await bookState(p);info.after = await allDefs(p);info.toast = await p.locator('#toast').textContent();
      if (scenario.blocked) {
        eq(after, before, 'blocked expansion preserves all sheets and Undo/Redo atomically');
        eq(await snapshotState(p), snapshotsBefore, 'blocked field/filter/option change preserves native snapshot registrations and contents');
        ok(/겹|중첩|overlap/i.test(info.toast || ''), 'visible overlap warning explains failed change');
        if (scenario.expansion === 'grand') { const d = p.getByRole('dialog', { name: '피벗 테이블 옵션', exact: true });eq(await d.isVisible(), true, 'failed option keeps dialog open for correction');await d.getByRole('button', { name: '취소', exact: true }).click(); }
        // A second write verifies the rejected layout did not leave stale ownership/capture caches.
        await run(p, 'pivotRefresh');await raf(p);const refreshed = await bookState(p);
        eq(refreshed.sheets, before.sheets, 'refresh after rejected edit preserves both original pivots');
      } else {
        const current = await sheetState(p);
        ok(JSON.stringify(current.definition.area) !== JSON.stringify(original.definition.area), 'allowed expansion actually grows the selected pivot');
        eq(current.definition.area, info.predicted.grown, 'result matches independently computed expanded bounds');
        for (let si = 0; si < before.sheets.length; si++) {
          if (si !== 1) eq(after.sheets[si], before.sheets[si], 'unrelated source or report sheet preserved');
        }
        const a = before.sheets[1], b = after.sheets[1], neighborDef = a.pivotsExtra?.[0];
        if (neighborDef) {
          eq(b.pivotsExtra, a.pivotsExtra, 'neighbor pivot definition preserved');
          const area = neighborDef.area, cells = sh => Object.fromEntries(Object.entries(sh.cells).filter(([key]) => { const [r, c] = key.split(',').map(Number);return r >= area.r1 && r <= area.r2 && c >= area.c1 && c <= area.c2; }));
          eq(cells(b), cells(a), 'neighbor pivot cells and formatting preserved');
        }
        eq(after.undo, before.undo + 1, 'allowed expansion has exactly one Undo entry');await run(p, 'undo');eq((await bookState(p)).sheets, before.sheets, 'Undo restores full pre-expansion book');
        await run(p, 'redo');eq((await bookState(p)).sheets, after.sheets, 'Redo restores accepted expansion');
      }
    });


  for (const operation of ['clear', 'toggle']) await test('multi-slicer-' + operation + '-atomic-block',
    async (p, info) => { const f = fixture({ expansion: 'filter', neighbor: 'below', filtered: true, multiSlicer: true });await openSynthetic(p, f, 'multi-slicer-' + operation);info.imported = await allDefs(p); },
    async (p, info) => {
      const before = await bookState(p), sl = p.locator('.obj[role="group"][aria-label="Group"]:visible');
      if (operation === 'clear') await sl.locator('.sl-clear').click();
      else await sl.locator('.sl-item').filter({ hasText: GROUPS[0] }).click();
      await raf(p);info.toast = await p.locator('#toast').textContent();info.after = await allDefs(p);
      eq(await bookState(p), before, 'a later conflicting linked pivot rejects all prior target changes and Undo');
      ok(/겹|중첩|overlap/i.test(info.toast || ''), 'linked slicer presents an overlap warning');
      const selected = await sl.locator('.sl-item.on').allTextContents();eq(selected.map(v => v.trim()), [GROUPS[0]], 'slicer selection remains unchanged after rejection');
    });
  await test('multi-slicer-clear-allowed-and-undo',
    async p => { await openSynthetic(p, fixture({ expansion: 'filter', neighbor: 'far', filtered: true, multiSlicer: true }), 'multi-slicer-allowed'); },
    async (p, info) => {
      const before = await bookState(p);await p.locator('.obj[role="group"][aria-label="Group"]:visible .sl-clear').click();await raf(p);
      const after = await bookState(p);info.after = await allDefs(p);
      for (const [si, name] of [[1, 'MainPivot'], [2, 'SafeLinkedPivot']]) {
        const sh = after.sheets[si], d = [sh.pivot, ...(sh.pivotsExtra ?? [])].find(d => d?.name === name);
        eq(d.filters?.Group, undefined, 'slicer clears every connected report filter');ok(d.area.r2 > before.sheets[si].pivot.area.r2, 'connected pivot really expands');
      }
      eq(after.sheets[1].pivotsExtra, before.sheets[1].pivotsExtra, 'unconnected nearby report remains intact');
      eq(after.undo, before.undo + 1, 'multi-report slicer is one Undo');await run(p, 'undo');eq((await bookState(p)).sheets, before.sheets, 'Undo restores every connected report and slicer');
      await run(p, 'redo');eq((await bookState(p)).sheets, after.sheets, 'Redo restores all linked report changes');
    });
  for (const command of ['pivotRefresh', 'refreshAll']) await test('batch-' + command + '-atomic-block',
    async (p, info) => { await openSynthetic(p, fixture({ neighbor: 'below', batchRefresh: true }), 'batch-' + command);await p.evaluate(() => { const w = tabula.wb();w.transact(() => w.setInput(0, 1, 0, 'Gamma new group'));w.undoStack = [];w.redoStack = []; });info.before = await allDefs(p); },
    async (p, info) => {
      const before = await bookState(p);await run(p, command);await raf(p);info.toast = await p.locator('#toast').textContent();info.after = await allDefs(p);
      eq(await bookState(p), before, 'batch rejection preserves earlier safe pivot, conflicting pivot, neighbor, dimensions, source edit and Undo');
      ok(/겹|중첩|overlap/i.test(info.toast || ''), 'batch refresh presents overlap warning');
    });
  await test('batch-refresh-allowed-and-undo',
    async p => { await openSynthetic(p, fixture({ neighbor: 'far', batchRefresh: true }), 'batch-refresh-allowed');await p.evaluate(() => { const w = tabula.wb();w.transact(() => w.setInput(0, 1, 0, 'Gamma new group'));w.undoStack = [];w.redoStack = []; }); },
    async (p, info) => {
      const before = await bookState(p);await run(p, 'pivotRefresh');await raf(p);const after = await bookState(p);info.after = await allDefs(p);
      for (const si of [0, 1]) ok(after.sheets[si].pivot.area.r2 > before.sheets[si].pivot.area.r2, 'fresh source grows both reports');
      eq(after.sheets[1].pivotsExtra, before.sheets[1].pivotsExtra, 'far unconnected report preserved');eq(after.undo, before.undo + 1, 'batch refresh has one Undo');
      await run(p, 'undo');eq((await bookState(p)).sheets, before.sheets, 'Undo restores all report cells and widths while keeping source edit');
      await run(p, 'redo');eq((await bookState(p)).sheets, after.sheets, 'Redo restores all refreshed reports');
    });


  await test('auto-refresh-collision-retry-same-source-version',
    async p => {
      await openSynthetic(p, fixture({ neighbor: 'below', batchRefresh: true }), 'auto-refresh-retry');
      for (const [name, sheetName] of [['EarlierRefreshPivot', 'Source'], ['MainPivot', 'Report']]) {
        await select(p, name, sheetName);await run(p, 'pivotOptions');const d = p.getByRole('dialog', { name: '피벗 테이블 옵션', exact: true });
        await d.getByRole('button', { name: '데이터', exact: true }).click();await d.getByRole('checkbox', { name: '원본 데이터가 바뀌면 자동 새로 고침 (WIXEL)', exact: true }).check();
        await d.getByRole('button', { name: '확인', exact: true }).click();
      }
      await select(p);await p.waitForTimeout(600);await clearHistory(p);
      await p.evaluate(() => { document.getElementById('toast').textContent = '';const w = tabula.wb();w.transact(() => w.setInput(0, 1, 0, 'Gamma new group')); });
    },
    async (p, info) => {
      const blocked = await bookState(p);await p.waitForFunction(() => /겹|중첩|overlap/i.test(document.getElementById('toast')?.textContent || ''), null, { timeout: 15000 });
      eq(await bookState(p), blocked, 'automatic collision leaves reports and source-edit Undo unchanged');
      info.blocked = await allDefs(p);await select(p, 'NeighborPivot', 'Report');await run(p, 'pivotMove');
      const d = p.getByRole('dialog', { name: '피벗 테이블 이동', exact: true });await d.getByRole('combobox').selectOption('existing');
      await d.getByLabel('위치', { exact: true }).fill('Report!$P$25');await d.getByRole('button', { name: '확인', exact: true }).click();await d.waitFor({ state: 'detached' });
      await p.waitForFunction(() => { const sh = tabula.wb().sheets[1];return sh.pivot.area.r2 === 6; }, null, { timeout: 15000 });
      const after = await bookState(p);info.after = await allDefs(p);
      eq(after.sheets[1].pivot.area.r2, blocked.sheets[1].pivot.area.r2 + 1, 'same source version is retried after space becomes available');
      eq(after.sheets[0].pivot.area.r2, blocked.sheets[0].pivot.area.r2 + 1, 'earlier connected automatic report updates only after successful batch');
      eq(after.sheets[1].pivotsExtra[0].left, 15, 'neighbor moves to chosen empty location');
      eq(after.sheets[0].cells['1,0'].raw, 'Gamma new group', 'retry keeps original source edit');
      eq(after.undo, blocked.undo + 1, 'automatic successful retry joins neighbor-move Undo');
      await run(p, 'undo');await p.waitForTimeout(600);eq((await bookState(p)).sheets, blocked.sheets, 'Undo returns reports to their safe state with source edit retained');
      await run(p, 'redo');await p.waitForTimeout(600);eq((await bookState(p)).sheets, after.sheets, 'Redo restores move and successful refresh');
    });


  await test('slicer-refresh-stale-snapshot-atomic-block',
    async (p, info) => {
      await openSynthetic(p, fixture({ neighbor: 'below', multiSlicer: true }), 'slicer-refresh-stale-snapshot');
      const sl = p.locator('.obj[role="group"][aria-label="Group"]:visible');
      // Open the real menu while the native cache is current. Display-only source reads keep
      // their normal invalidation behavior; this case isolates the refresh preparation read.
      await sl.locator('.sl-cap').click({ button: 'right' });
      await p.getByRole('menuitem', { name: /^새로 고침/ }).waitFor();
      info.setup = await p.evaluate(() => {
        const w = tabula.wb(), d = w.sheets[1].pivot, snap = w.pivotSnapshots?.get(d.snapshotId);
        if (!snap) throw new Error('Fixture must have a native registered pivot snapshot.');
        const currentBefore = w.pivotSnapshotCurrent({ ...snap }, d);
        // Mark only the source stale, without emitting a grid repaint or scheduling a refresh.
        // The menu action below is the real UI path under test; this is fixture preparation.
        w.setInput(0, 1, 0, 'Gamma new group');
        window.__slicerSnapshotRegistryBefore = w.pivotSnapshots;
        window.__slicerSnapshotReferencesBefore = new Map(w.pivotSnapshots);
        return { snapshotId: d.snapshotId, registered: w.pivotSnapshots.size, currentBefore,
          staleAfter: !w.pivotSnapshotCurrent({ ...snap }, d) };
      });
      eq(info.setup.currentBefore, true, 'native snapshot was current before source edit');
      eq(info.setup.staleAfter, true, 'source edit makes the registered native snapshot stale');
      ok(info.setup.registered > 0, 'a registered snapshot exists at the refresh boundary');
      info.before = await allDefs(p);
    },
    async (p, info) => {
      const before = await bookState(p), snapshotsBefore = await snapshotState(p);
      await p.getByRole('menuitem', { name: /^새로 고침/ }).click();await raf(p);
      info.toast = await p.locator('#toast').textContent();info.after = await allDefs(p);
      info.registrations = { before: snapshotsBefore.length, after: (await snapshotState(p)).length };
      eq(await bookState(p), before, 'failed slicer refresh preserves definitions, cells, formats, widths, source edit and Undo/Redo');
      eq(await snapshotState(p), snapshotsBefore, 'failed refresh preserves every stale snapshot registration, metadata and row content');
      eq(await p.evaluate(() => { const w = tabula.wb(), saved = window.__slicerSnapshotReferencesBefore;
        return w.pivotSnapshots === window.__slicerSnapshotRegistryBefore && w.pivotSnapshots.size === saved.size &&
          [...saved].every(([id, snap]) => w.pivotSnapshots.get(id) === snap); }), true, 'failed refresh preserves the original registry and registered object identities');
      ok(/겹|중첩|overlap/i.test(info.toast || ''), 'refresh reports neighboring pivot overlap');
      eq(await p.evaluate(() => tabula.wb().getRaw(0, 1, 0)), 'Gamma new group', 'source edit is retained after rejected refresh');
      await p.screenshot({ path: path.join(out, 'slicer-refresh-stale-snapshot-atomic-block-warning.png') });
    });

  if (real) {
    const initialStat = await stat(real), bytes = await readFile(real), raw = new Workbook(readXlsx(bytes).data);
    const expected = raw.sheets.flatMap((sh, si) => [sh.pivot, ...(sh.pivotsExtra ?? [])].filter(Boolean).map(d => ({ si, sheet: sh.name, name: d.name, autofit: d.autofit !== false, preserveFormat: d.preserveFormat !== false, area: d.area, rows: d.rows, cols: d.cols })));
    eq(expected.length, 14, 'original contains fourteen native pivots');
    eq(expected.map(d => d.autofit), Array(14).fill(false), 'original raw options have autofit disabled for all fourteen pivots');
    eq(expected.map(d => d.preserveFormat), Array(14).fill(true), 'original raw options preserve cell formatting');
    await writeFile(path.join(out, 'real-import-options.json'), JSON.stringify(expected, null, 2));
    const nativeBack = new Workbook(readXlsx(writeXlsx(raw)).data);
    const backOptions = nativeBack.sheets.flatMap((sh, si) => [sh.pivot, ...(sh.pivotsExtra ?? [])].filter(Boolean).map(d => ({ si, name: d.name, autofit: d.autofit !== false, preserveFormat: d.preserveFormat !== false })));
    eq(backOptions, expected.map(({ si, name, autofit, preserveFormat }) => ({ si, name, autofit, preserveFormat })), 'original fourteen options survive XLSB-to-XLSX native roundtrip');
    await test('private-original-options-and-search-sort', async (p, info) => {
      await p.locator('#fileInput').setInputFiles({ name: path.basename(real), mimeType: 'application/vnd.ms-excel.sheet.binary.macroEnabled.12', buffer: bytes });
      await p.waitForFunction(count => tabula.wb().sheets.reduce((n, sh) => n + Number(!!sh.pivot) + (sh.pivotsExtra?.length || 0), 0) === count && /열었습니다/.test(document.getElementById('toast')?.textContent || ''), expected.length, { timeout: 180000 });
      info.imported = await allDefs(p);eq(info.imported.map(({ area, ...d }) => d), expected.map(({ area, ...d }) => d), 'original pivot imported options match raw workbook');
      info.dialogs = [];
      for (const d of expected) { await select(p, d.name, d.sheet);const enabled = await options(p, undefined, { confirm: false });eq(enabled, d.autofit, 'original option checkbox for ' + d.name);info.dialogs.push({ name: d.name, sheet: d.sheet, autofit: enabled }); }
      const target = expected.find(d => /검색어/.test(d.sheet) && d.rows?.length && !d.autofit);
      ok(target, 'original has a disabled-autofit search-term pivot');info.target = target;await select(p, target.name, target.sheet);await clearHistory(p);
    }, async (p, info) => {
      const target = info.target, before = await p.evaluate(si => { const w = tabula.wb();return { widths: Array.from({ length: w.usedRange(si).cols }, (_, c) => w.colWidth(si, c)) }; }, target.si);
      await p.evaluate(name => window.__realTarget = name, target.name);
      const m = await rowMenu(p, target.rows[0]);await m.getByRole('menuitem', { name: '텍스트 내림차순 정렬', exact: true }).click();await raf(p);
      const after = await p.evaluate(({ si, name, columns }) => { const w = tabula.wb();return { widths: Array.from({ length: columns }, (_, c) => w.colWidth(si, c)), def: structuredClone([w.sheets[si].pivot, ...(w.sheets[si].pivotsExtra ?? [])].find(d => d.name === name)) }; }, { si: target.si, name: target.name, columns: before.widths.length });
      eq(after.widths, before.widths, 'original disabled-autofit sort preserves all source report widths');eq(after.def.sort?.[target.rows[0]]?.dir, 'desc', 'original requested descending sort applied');
      await run(p, 'undo');eq(await p.evaluate(({ si, columns }) => Array.from({ length: columns }, (_, c) => tabula.wb().colWidth(si, c)), { si: target.si, columns: before.widths.length }), before.widths, 'original Undo preserves widths');info.widths = { before: before.widths, after: after.widths };
    });
    const finalStat = await stat(real);eq({ size: finalStat.size, mtimeMs: finalStat.mtimeMs }, { size: initialStat.size, mtimeMs: initialStat.mtimeMs }, 'original file size and modification time unchanged');
  }
} finally {
  await browser.close();
  const summary = { url, baseline: baseline || (refreshLeakBaseline ? 'slicerRefresh-single-line-leak' : null), engine, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, original: real ? path.basename(real) : null, results };
  await writeFile(path.join(out, 'result.json'), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify({ url, baseline: baseline || (refreshLeakBaseline ? 'slicerRefresh-single-line-leak' : null), engine, cases: summary.cases, passed: summary.passed, checks, assets: summary.assets, out }));
  if (summary.passed !== summary.cases || !summary.cases) process.exitCode = 1;
}
