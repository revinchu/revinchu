// External cached-pivot UI regression. Synthetic data only; never reads business files.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';
import { makeImportedPivotSourceReference } from '../src/pivot-source-reference.js';
import { unzip, textOf } from '../src/zip.js';

const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/', origin = new URL(url).origin;
if (!['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname)) throw new Error('This synthetic regression is local only.');
const engine = process.env.WIXEL_BROWSER || 'chromium';
if (!['chromium', 'webkit', 'firefox'].includes(engine)) throw new Error('Unsupported browser.');
const out = path.resolve(repo, process.env.WIXEL_EXTERNAL_PIVOT_OUT || '.local/large-xlsx-ipad/external-pivot-cache/' + engine);
if (!/^[dD]:/.test(out)) throw new Error('Outputs must stay on D:.');
const only = process.env.WIXEL_EXTERNAL_PIVOT_FILTER || '';
await mkdir(out, { recursive: true });
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await pw[engine].launch();
const results = [], pageErrors = [], writes = [], externalRequests = [], blockedApi = [], assets = new Set(); let checks = 0;
const eq = (actual, expected, message) => { checks++; assert.deepEqual(actual, expected, message); };
const ok = (actual, message) => { checks++; assert.ok(actual, message); };
const raf = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const rows = [['Group', 'Value'], ['A', 10], ['B', 20]];
const externalPath = 'file:///D:/Synthetic/ExternalSource.xlsx';
function fixtureData({ metadataOnly = false, localOnly = false, mixed = false, saveData = true } = {}) {
  const def = { name: localOnly ? 'Local Summary' : 'External Summary', source: 'Source', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['Group'], cols: [], pages: [], values: [{ field: 'Value', agg: 'sum', name: 'Revenue' }], top: 4, left: 1, layout: 'tabular', subtotals: 'none', grandRows: true, grandCols: true, autofit: false, preserveFormat: true, autoRefresh: false, saveData,
    area: { r1: 4, c1: 1, r2: 7, c2: 2 }, buttons: [{ r: 4, c: 1, kind: 'rows', field: 'Group' }] };
  if (!localOnly) {
    def.cacheItemsId = 'saved'; def.snapshotId = 'saved';
    def.sourceReference = makeImportedPivotSourceReference({ ref: 'A1:B3', sheet: 'Source', name: null, external: externalPath }, def);
  }
  const source = { '0,0': { raw: 'Group' }, '0,1': { raw: 'Value' }, '1,0': { raw: localOnly ? 'A' : 'Local A' }, '1,1': { raw: '999' }, '2,0': { raw: localOnly ? 'B' : 'Local B' }, '2,1': { raw: '999' } };
  const report = {}, seed = (d, amount) => { [['Group', 'Revenue'], ['A', amount], ['B', amount * (localOnly ? 1 : 2)], ['합계', amount * (localOnly ? 2 : 3)]].forEach((row, r) => row.forEach((value, c) => { report[(d.top + r) + ',' + (d.left + c)] = { raw: String(value) }; })); };
  seed(def, localOnly ? 999 : 10);
  const sh = { name: 'Report', cells: report, pivot: def, colWidths: { 1: 165, 2: 145 } };
  if (mixed) {
    const local = { ...structuredClone(def), name: 'Local Summary', top: 14, area: { r1: 14, c1: 1, r2: 17, c2: 2 }, buttons: [{ r: 14, c: 1, kind: 'rows', field: 'Group' }] };
    delete local.sourceReference; delete local.snapshotId; delete local.cacheItemsId;
    sh.pivotsExtra = [local];
    [['Group', 'Revenue'], ['A', 999], ['B', 999], ['합계', 1998]].forEach((row, r) => row.forEach((value, c) => report[(14 + r) + ',' + (1 + c)] = { raw: String(value) }));
  }
  return { ...(!localOnly ? { pivotCacheItems: { saved: { missingItemsLimit: 1048576, fields: [{ name: 'Group', shared: ['A', 'B'], sharedTypes: 'ss' }, { name: 'Value', shared: [10, 20], sharedTypes: 'nn' }] } }, ...(!metadataOnly ? { pivotSnapshots: { saved: rows } } : {}) } : {}), sheets: [{ name: 'Source', cells: source }, sh] };
}
async function fixture(p, options = {}) {
  const data = fixtureData(options);
  await p.evaluate(data => {
    const t = tabula, w = t.wb(); w.restore(data); t.switchSheet(1); t.selectCell(4, 1); t.gv().setZoom(100); t.gv().layout();
    w.undoStack = []; w.redoStack = [];
    window.__externalState = () => {
      const t = tabula, w = t.wb(), sh = w.sheets[1], d = sh.pivot;
      return { book: JSON.stringify(w.serialize()), version: w.version, undo: w.undoStack.length, redo: w.redoStack.length,
        snapshotKeys: [...(w.pivotSnapshots?.keys() ?? [])], source: [w.getValue(0, 1, 1), w.getValue(0, 2, 1)],
        filters: structuredClone(d.filters ?? {}), total: w.getValue(1, d.area.r2, d.area.c2),
        visibleRows: Array.from({ length: d.area.r2 - d.area.r1 + 1 }, (_, r) => Array.from({ length: d.area.c2 - d.area.c1 + 1 }, (_, c) => w.getValue(1, d.area.r1 + r, d.area.c1 + c))),
        active: { ...t.active }, selection: { ...t.sel }, sourceVersion: w.sheetVersion(0) };
    };
  }, data);
  await raf(p); await p.locator('#cellEditor').focus(); return data;
}
const state = p => p.evaluate(() => window.__externalState());
const immutable = s => ({ book: s.book, version: s.version, undo: s.undo, redo: s.redo, snapshotKeys: s.snapshotKeys, source: s.source, total: s.total, sourceVersion: s.sourceVersion });
const content = s => ({ book: s.book, snapshotKeys: s.snapshotKeys, source: s.source, total: s.total });
async function command(p, name) { await p.evaluate(name => tabula.run(name), name); await raf(p); }
async function openFilter(p) {
  const point = await p.evaluate(() => { const t = tabula, d = t.wb().sheets[1].pivot, b = (d.buttons ?? []).find(b => b.kind === 'rows' && b.field === 'Group' && !b.sigma); if (!b) throw Error('Missing synthetic field header'); t.selectCell(b.r, b.c); t.gv().ensureVisible(b.r, b.c); return { r: b.r, c: b.c }; });
  await raf(p); await p.locator('#cellEditor').focus(); await p.keyboard.down('Alt'); await p.keyboard.press('ArrowDown'); await p.keyboard.up('Alt'); await raf(p);
  const menu = p.locator('.menu:has(.pivot-filter-menu)').first(); await menu.waitFor({ timeout: 4000 });
  eq((await state(p)).active, point, 'Actual field-menu keyboard opening keeps the selected header cell');
  return menu;
}
async function pickA(p) {
  const menu = await openFilter(p);
  const values = await menu.locator('.filter-window input[type=checkbox]').evaluateAll(nodes => nodes.map(n => n.getAttribute('aria-label')));
  eq(values, ['A', 'B'], 'Dropdown items come from retained external cache, not unrelated local data');
  const b = menu.getByRole('checkbox', { name: 'B', exact: true }); ok(await b.isChecked(), 'B begins selected'); await b.uncheck();
  ok(await menu.getByRole('checkbox', { name: 'A', exact: true }).isChecked(), 'A remains selected');
  await menu.getByRole('button', { name: '확인', exact: true }).click(); await raf(p);
  eq(await p.locator('.pivot-filter-menu').count(), 0, 'A changed filter attempts the actual confirmation path');
}
async function caseRun(name, fn) {
  if (only && !only.split('|').some(value => name.includes(value))) return;
  const context = await browser.newContext({ viewport: { width: 1440, height: 1020 }, serviceWorkers: 'block' }), p = await context.newPage(), info = {}, errors = []; const start = checks;
  p.setDefaultTimeout(25000);
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {} } }); window.__externalFetchAttempts = []; const original = window.fetch; window.fetch = function(input, init) { const raw = input instanceof Request ? input.url : String(input), u = new URL(raw, location.href), method = init?.method ?? (input instanceof Request ? input.method : 'GET'); if (u.origin !== location.origin || !['GET', 'HEAD', 'OPTIONS'].includes(method.toUpperCase())) window.__externalFetchAttempts.push({ url: u.href, method }); return original.call(this, input, init); }; });
  await context.route('**/*', route => {
    const r = route.request(), u = new URL(r.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push({ name, method: r.method(), path: u.pathname }); return route.abort(); }
    if (u.origin !== origin) { externalRequests.push({ name, url: r.url(), method: r.method() }); return route.abort(); }
    if (/^\/api(?:\/|$)/.test(u.pathname)) { blockedApi.push({ name, path: u.pathname }); return route.abort(); }
    return route.continue();
  });
  p.on('pageerror', e => { errors.push(e.message); pageErrors.push({ name, error: e.message }); }); p.on('dialog', d => d.dismiss());
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    for (const src of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => n.getAttribute('src')))) assets.add(src);
    await fn(p, info); eq(await p.evaluate(() => window.__externalFetchAttempts), [], 'No external or write fetch was attempted, including file: URLs rejected before network routing'); eq(errors, [], 'No application exception'); await p.screenshot({ path: path.join(out, name + '.png') }); results.push({ name, pass: true, checks: checks - start, info });
  } catch (error) { await p.screenshot({ path: path.join(out, name + '-failed.png') }).catch(() => {}); results.push({ name, pass: false, checks: checks - start, error: error.stack, info }); }
  finally { await context.close(); }
  console.log(JSON.stringify({ name, pass: results.at(-1).pass, checks: results.at(-1).checks, error: results.at(-1).error?.split('\n')[0] }));
}
function cachedResult(w) {
  const d = w.sheets[1].pivot, src = pivotSourceData(w, d, { preserveSnapshot: true });
  if (!src) return null;
  const resolved = resolvePivot(src, d), { grid } = computePivot(resolved, resolved.def);
  return { cacheOnly: src.cacheOnly === true, records: src.cube.n, values: Array.from({ length: src.cube.n }, (_, r) => src.cube.col(1).get(r)), total: Number(grid.at(-1).at(-1).raw) };
}
async function openSynthetic(p, file, name) {
  await p.locator('input[type=file]').first().setInputFiles(file);
  await p.waitForFunction(name => tabula.wb().sheets[1]?.pivot?.name === name, name, { timeout: 30000 });
  await p.evaluate(() => { tabula.switchSheet(1); tabula.selectCell(4, 1); tabula.gv().layout(); }); await raf(p);
}
try {
  await caseRun('full-record-filter-undo-redo', async (p, info) => {
    await fixture(p); const initial = await state(p); eq(initial.total, 30, 'Seeded report independently totals 10+20'); eq(initial.source, [999, 999], 'Local sheet deliberately disagrees with saved cache');
    await pickA(p); const after = await state(p); eq(after.filters, { Group: ['A'] }, 'Changed filter is committed'); eq(after.total, 10, 'Filtered value is saved external 10, not local 999'); eq(after.undo, initial.undo + 1, 'Filter is one Undo transaction'); eq(after.source, initial.source, 'Filter preserves local source cells'); eq(after.sourceVersion, initial.sourceVersion, 'Filter preserves local source version'); eq(after.snapshotKeys, ['saved'], 'Filter keeps saved records');
    await command(p, 'undo'); const undone = await state(p); eq(content(undone), content(initial), 'Undo restores exact cells, definition and saved cache'); eq(undone.redo, 1, 'Undo records one Redo');
    await command(p, 'redo'); const redone = await state(p); eq(content(redone), content(after), 'Redo restores exact filter result and cache'); info.after = { total: redone.total, filters: redone.filters, undo: redone.undo };
  });
  await caseRun('refresh-external-mixed-batch-atomic-refusal', async (p, info) => {
    await fixture(p, { mixed: true }); const before = await state(p);
    await command(p, 'pivotRefresh'); eq(immutable(await state(p)), immutable(before), 'Refresh-all rejects before clearing any records or rewriting the normal local pivot');
    ok(/원본이 이 문서에 없어/.test(await p.locator('#toast').innerText()), 'Refresh visibly explains missing external source');
    eq(await p.evaluate(() => tabula.wb().getValue(1, 17, 2)), 1998, 'Rejected mixed refresh preserves local pivot result'); info.after = { total: (await state(p)).total, undo: (await state(p)).undo };
  });
  await caseRun('same-name-local-edit-keeps-external-records', async (p, info) => {
    await fixture(p); const initial = await state(p);
    await p.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setInput(0, 1, 1, '1234')); }); await raf(p);
    const changed = await state(p); eq(changed.source, [1234, 999], 'Local same-name sheet is genuinely edited'); eq(changed.snapshotKeys, ['saved'], 'Unrelated local edit does not invalidate external records');
    await pickA(p); const after = await state(p); eq(after.total, 10, 'External computation still uses saved 10 after same-name local edit'); eq(after.source, changed.source, 'External filtering preserves local edit'); eq(after.undo, changed.undo + 1, 'External filter adds exactly one transaction');
    await command(p, 'undo'); eq(content(await state(p)), content(changed), 'Filter Undo retains the prior independent local edit'); await command(p, 'undo'); eq(content(await state(p)), content(initial), 'Second Undo restores original local cells and external cache'); info.after = { total: after.total, source: after.source };
  });
  await caseRun('metadata-only-filter-sort-refresh-atomic-refusal', async (p, info) => {
    await fixture(p, { metadataOnly: true, saveData: false }); const before = await state(p); eq(before.snapshotKeys, [], 'Fixture retains catalog without fabricated records');
    await pickA(p); eq(immutable(await state(p)), immutable(before), 'Actual changed selection is rejected without cells, filters, version or Undo changes');
    const menu = await openFilter(p); await menu.getByRole('menuitem', { name: '텍스트 내림차순 정렬', exact: true }).click(); await raf(p);
    eq(immutable(await state(p)), immutable(before), 'Metadata-only sort is also rejected atomically');
    await command(p, 'pivotRefresh'); eq(immutable(await state(p)), immutable(before), 'Metadata-only refresh preserves report and retained catalog'); ok(/원본이 이 문서에 없어/.test(await p.locator('#toast').innerText()), 'Refresh denial is visible'); info.after = { total: before.total, undo: before.undo };
  });
  await caseRun('normal-local-edit-filter-refresh-undo-redo', async (p, info) => {
    await fixture(p, { localOnly: true }); const initial = await state(p); eq(initial.total, 1998, 'Normal local fixture has its independent total');
    await p.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setInput(0, 1, 1, '1234')); }); await raf(p); const changed = await state(p);
    await command(p, 'pivotRefresh'); const refreshed = await state(p); eq(refreshed.total, 2233, 'Normal refresh reads actual local edited data'); eq(refreshed.undo, changed.undo + 1, 'Normal refresh remains Undoable');
    await command(p, 'undo'); eq(content(await state(p)), content(changed), 'Refresh Undo restores prior visible report'); await command(p, 'redo'); eq(content(await state(p)), content(refreshed), 'Refresh Redo restores newly computed local total');
    await pickA(p); eq((await state(p)).total, 1234, 'Normal local filtering still uses live source'); info.after = { total: (await state(p)).total };
  });
  await caseRun('full-cache-xlsx-roundtrip-file-open-and-filter', async (p, info) => {
    const w = new Workbook(fixtureData()), before = JSON.stringify(w.serialize()), bytes = writeXlsx(w), file = path.join(out, 'synthetic-external-cache.xlsx'); await writeFile(file, bytes);
    eq(JSON.stringify(w.serialize()), before, 'XLSX exporter preserves original model');
    const parts = unzip(bytes), cache = textOf(parts['xl/pivotCache/pivotCacheDefinition1.xml']); ok(cache.includes('type="external"') || /worksheetSource[^>]+r:id=/.test(cache), 'XLSX source retains external provenance');
    for (const [mode, read] of [['sync', readXlsx(bytes)], ['stream', await readXlsxAsync(bytes, undefined, { streamThreshold: 0 })], ['Blob', await readXlsxAsync(new Blob([bytes]), undefined, { streamThreshold: 0 })]]) {
      const back = new Workbook(read.data), got = cachedResult(back); eq(got, { cacheOnly: true, records: 2, values: [10, 20], total: 30 }, mode + ' roundtrip preserves external cache values'); eq(back.sheets[1].pivot.sourceReference.external, externalPath, mode + ' roundtrip preserves external source'); eq(back.getValue(0, 1, 1), 999, mode + ' same-name local source remains distinct');
    }
    await openSynthetic(p, file, 'External Summary');
    await p.evaluate(() => { const w = tabula.wb(); w.undoStack = []; w.redoStack = []; });
    // Install observation functions through a separate fixture page function, without replacing the loaded file.
    await p.evaluate(() => { window.__externalState = () => { const t = tabula, w = t.wb(), d = w.sheets[1].pivot; return { book: JSON.stringify(w.serialize()), version: w.version, undo: w.undoStack.length, redo: w.redoStack.length, snapshotKeys: [...(w.pivotSnapshots?.keys() ?? [])], source: [w.getValue(0, 1, 1), w.getValue(0, 2, 1)], filters: structuredClone(d.filters ?? {}), total: w.getValue(1, d.area.r2, d.area.c2), active: { ...t.active }, selection: { ...t.sel }, sourceVersion: w.sheetVersion(0) }; }; });
    eq((await state(p)).total, 30, 'Browser file open displays saved external total'); await pickA(p); eq((await state(p)).total, 10, 'Imported file filter computes external cached records'); const filtered = await state(p); await command(p, 'pivotRefresh'); eq(immutable(await state(p)), immutable(filtered), 'Imported external file refresh is atomically rejected');
    const saved = await p.evaluate(() => tabula.wb().serialize()), second = new Workbook(readXlsx(writeXlsx(new Workbook(saved))).data); eq(cachedResult(second).total, 10, 'Filtered external selection survives another XLSX save/read'); eq(second.sheets[1].pivot.filters.Group, ['A'], 'Filtered external selection metadata survives save'); info.after = { total: (await state(p)).total, xlsxBytes: bytes.length };
  });
  await caseRun('no-record-xlsx-preserves-definition-catalog-and-display', async (p, info) => {
    const w = new Workbook(fixtureData({ saveData: false })), bytes = writeXlsx(w), file = path.join(out, 'synthetic-external-metadata.xlsx'); await writeFile(file, bytes);
    const back = new Workbook(readXlsx(bytes).data), d = back.sheets[1].pivot;
    eq(cachedResult(back), null, 'No-record file never manufactures an empty UI calculation'); eq(back.pivotCacheItems[d.cacheItemsId].fields[0].shared, ['A', 'B'], 'Retained field catalog survives saveData=false'); eq(d.rows, ['Group'], 'Metadata-only saved pivot keeps assigned fields'); eq(d.sourceReference.external, externalPath, 'Metadata-only file keeps external provenance'); eq(back.getValue(1, 7, 2), 30, 'No-record save preserves visible cached result cells');
    const second = new Workbook(readXlsx(writeXlsx(back)).data); eq(second.sheets[1].pivot.rows, ['Group'], 'Exporter-only empty cache preserves fields on repeated save'); eq(second.getValue(1, 7, 2), 30, 'Repeated metadata-only save preserves display'); eq(cachedResult(second), null, 'Repeated save still has no invented records');
    await openSynthetic(p, file, 'External Summary');
    await p.evaluate(() => { const w = tabula.wb(); window.__externalState = () => { const t = tabula, w = t.wb(), d = w.sheets[1].pivot; return { book: JSON.stringify(w.serialize()), version: w.version, undo: w.undoStack.length, redo: w.redoStack.length, snapshotKeys: [...(w.pivotSnapshots?.keys() ?? [])], source: [w.getValue(0, 1, 1), w.getValue(0, 2, 1)], filters: structuredClone(d.filters ?? {}), total: w.getValue(1, d.area.r2, d.area.c2), active: { ...t.active }, sourceVersion: w.sheetVersion(0) }; }; });
    const before = await state(p); eq(before.total, 30, 'Metadata-only file open retains visible result'); info.opened = await p.evaluate(() => ({ rows: tabula.wb().sheets[1].pivot.rows, buttons: tabula.wb().sheets[1].pivot.buttons ?? [], renderedButtons: document.querySelectorAll('.pbtn').length }));
    info.limitations = ['Metadata-only XLSX has no worksheet filter arrows; actual header filtering remains unsupported. Direct retained-button refusal is tested separately.'];
    eq(info.opened.buttons, [], 'Known display limitation is recorded explicitly: no imported metadata-only header buttons'); eq(info.opened.renderedButtons, 0, 'Known display limitation also exists in the actual rendered sheet');
    const pane = p.locator('#pivotPane'); if (await pane.isVisible()) await command(p, 'pivotFieldList'); await command(p, 'pivotFieldList'); await pane.waitFor({ state: 'visible' }); ok(/원본 데이터를 찾을 수 없습니다/.test(await pane.innerText()), 'Actual metadata-only field list explains unavailable source'); eq(await pane.locator('[data-pivot-field]').count(), 0, 'Field list offers no calculation-changing controls without records'); eq(immutable(await state(p)), immutable(before), 'Opening unavailable-source field list preserves cells, definitions, catalog and Undo');
    await command(p, 'pivotRefresh'); eq(immutable(await state(p)), immutable(before), 'Actual imported metadata-only refresh is refused without mutation'); ok(/원본이 이 문서에 없어/.test(await p.locator('#toast').innerText()), 'Imported metadata-only refresh denial is visible'); info.after = { total: (await state(p)).total, xlsxBytes: bytes.length };
  });
} finally {
  await browser.close();
  try { eq(writes, [], 'No API/HTTP write request was attempted'); eq(externalRequests, [], 'External source URI was never fetched'); } catch (error) { results.push({ name: 'network-isolation', pass: false, error: error.stack }); }
  const summary = { syntheticOnly: true, privateWorkbookInput: false, url, browser: engine, assets: [...assets], passed: results.filter(r => r.pass).length, failed: results.filter(r => !r.pass).length, checks, pageErrors, writes, externalRequests, blockedApi, browserClosed: true, knownLimitations: [...new Set(results.flatMap(r => r.info?.limitations ?? []))], results };
  await writeFile(path.join(out, 'result.json'), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify({ passed: summary.passed, failed: summary.failed, checks, pageErrors: pageErrors.length, writes: writes.length, externalRequests: externalRequests.length, assets: summary.assets, out }));
  if (summary.failed || pageErrors.length) process.exitCode = 1;
}