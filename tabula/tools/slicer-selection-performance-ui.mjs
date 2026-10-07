// Synthetic slicer timing and correctness probes. No private input or document API writes.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const origin = new URL(url).origin;
const local = ['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname);
if (!local && process.env.WIXEL_ALLOWED_TEST_URL !== url) throw new Error('Public synthetic tests require an exact WIXEL_ALLOWED_TEST_URL.');
const out = path.resolve(repo, process.env.WIXEL_SLICER_PERF_OUT || '.local/slicer-selection-performance/source');
if (!/^[dD]:/.test(out)) throw new Error('Outputs must be on D:.');
const only = process.env.WIXEL_SLICER_PERF_FILTER || '';
const baseline = process.env.WIXEL_SLICER_PERF_BASELINE || '';
const sourceRows = Number(process.env.WIXEL_SLICER_PERF_ROWS || 240000);
assert.ok(Number.isInteger(sourceRows) && sourceRows >= 60 && sourceRows <= 500000, 'Bounded synthetic source row count');
const overrides = new Map();
if (baseline) {
  if (!local) throw new Error('Baseline memory routing is local only.');
  for (const name of ['app.js', 'workbook.js', 'cellmap.js', 'pivot.js', 'cube.js', 'pivot-field-items.js', 'view.js']) overrides.set('/src/' + name,
    execFileSync('git', ['-c', 'safe.directory=' + path.resolve(repo, '..').replaceAll(String.fromCharCode(92), '/'), '-C', repo, 'show', baseline + ':tabula/src/' + name], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }));
}
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await pw.chromium.launch();
await mkdir(out, { recursive: true });
const results = [], errors = [], writes = [], blocked = [], assets = new Set(); let checks = 0;
const eq = (a, b, message) => { checks++; assert.deepEqual(a, b, message); };
const ok = (value, message) => { checks++; assert.ok(value, message); };
const raf = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const state = p => p.evaluate(() => window.__slState());
const run = async (p, name) => { await p.evaluate(name => tabula.run(name), name); await raf(p); };
async function fixture(p, { n = sourceRows, linked = 4, reject = false, extraSlicer = false, duplicated = false, rowGroups = 4, channelGroups = 3 } = {}) {
  const setup = await p.evaluate(({ n, linked, reject, extraSlicer, duplicated, rowGroups, channelGroups }) => {
    const t = tabula, w = t.wb(), header = ['Region', 'Channel', 'Campaign', 'Week', 'Value', 'Cost'];
    const dims = [rowGroups, channelGroups, 6, 5].map((count, j) => ({ num: null, str: new Int32Array(n), dict: Array.from({ length: count }, (_, i) => ['Region ', 'Channel ', 'Campaign ', 'Week '][j] + (j < 2 ? String(i + 1).padStart(4, '0') : String.fromCharCode(65 + i))), fmt: null }));
    const nums = [4, 5].map(() => ({ num: new Float64Array(n), str: null, dict: [], fmt: null }));
    const totals = {}, weekTotals = {};
    for (let i = 0; i < n; i++) {
      dims[0].str[i] = Math.floor(i / 6) % rowGroups; dims[1].str[i] = Math.floor(i / 24) % channelGroups; dims[2].str[i] = i % 6; dims[3].str[i] = Math.floor(i / 72) % 5;
      nums[0].num[i] = i % 17 + 1; nums[1].num[i] = i % 11 + 2;
      const name = dims[2].dict[i % 6], week = dims[3].dict[dims[3].str[i]];
      totals[name] ??= { Value: 0, Cost: 0 }; weekTotals[week] ??= {};
      weekTotals[week][name] ??= { Value: 0, Cost: 0 };
      for (const [j, field] of ['Value', 'Cost'].entries()) { totals[name][field] += nums[j].num[i]; weekTotals[week][name][field] += nums[j].num[i]; }
    }
    const cells = Object.fromEntries(header.map((value, c) => ['0,' + c, { raw: value }]));
    const source = { name: 'Source', cells, blocks: [{ r0: 1, c0: 0, n, ver: 0, cols: [...dims, ...nums] }] };
    const mk = (index, patch = {}) => ({ name: 'Pivot ' + (index + 1), source: 'Source', range: { r1: 0, c1: 0, r2: n, c2: 5 },
      rows: reject && index === 0 ? ['Campaign', 'Region'] : index % 2 ? ['Channel'] : ['Region'], cols: [], pages: [],
      values: [{ field: index % 2 ? 'Cost' : 'Value', agg: 'sum' }], top: 18, left: 0, layout: 'tabular', subtotals: 'none',
      grandRows: true, grandCols: true, autofit: false, preserveFormat: true, autoRefresh: false,
      ...(reject ? { filters: { Campaign: ['Campaign A'] } } : {}), ...patch });
    const reports = Array.from({ length: linked }, (_, i) => ({ name: 'Report ' + (i + 1), cells: {}, pivot: mk(i), colWidths: { 0: 140, 1: 130, 2: 100 } }));
    w.restore({ sheets: [source, ...reports] }); t.switchSheet(1); t.selectCell(0, 0); t.gv().setZoom(100); t.run('pivotRefresh');
    if (reject) {
      const main = w.sheets[1].pivot;
      w.sheets[1].pivotsExtra = [mk(20, { name: 'Blocking Neighbor', rows: ['Channel'], top: main.area.r2 + 1, filters: {} })];
      t.run('pivotRefresh');
    }
    const links = reports.map(sh => ({ sheet: sh.name, name: sh.pivot.name }));
    if (reject) links.push(links.shift());
    if (duplicated) links.push(...links.map(link => ({ ...link })));
    const slicers = [{ id: 'campaign-slicer', caption: '캠페인', source: { kind: 'pivot', field: 'Campaign', pivots: links },
      x: 35, y: 20, w: 620, h: 142, columns: 3, style: 'SlicerStyleLight1', noDataLast: false }];
    if (extraSlicer) slicers.push({ id: 'week-slicer', caption: '주차', source: { kind: 'pivot', field: 'Week', pivots: links }, x: 690, y: 20, w: 380, h: 175, columns: 2, noDataLast: false });
    w.sheets[1].slicers = slicers; t.selectCell(0, 0); t.gv().layout(); w.undoStack = []; w.redoStack = [];
    window.__slExpected = { totals, weekTotals, n };
    window.__slSourceOriginal = { sheet: w.sheets[0], block: w.sheets[0].blocks[0], ver: w.sheetVersion(0) };
    return { n, linked, links: links.length, totals, weekTotals, sourceVersion: w.sheetVersion(0), initialAreas: reports.map(sh => w.sheets.find(s => s.name === sh.name).pivot.area) };
  }, { n, linked, reject, extraSlicer, duplicated, rowGroups, channelGroups });
  await raf(p); await p.locator('.obj[data-id="campaign-slicer"] .sl-item').first().waitFor();
  await p.evaluate(() => { window.__slTrace = []; window.__slTimings = []; window.__slLongTasks = []; });
  return setup;
}
async function selected(p, campaign = null, week = null) {
  const actual = await state(p);
  const selection = campaign === null ? Array.from({ length: 6 }, (_, i) => 'Campaign ' + String.fromCharCode(65 + i)) : campaign;
  eq(actual.selected['campaign-slicer'], selection, 'Visible slicer selected buttons match the filter');
  for (const def of actual.pivots.filter(d => d.name !== 'Blocking Neighbor')) {
    eq(def.filters.Campaign ?? null, campaign, 'Every connected pivot has the exact selected campaign filter');
    eq(def.filters.Week ?? null, week, 'Every connected pivot has the exact selected week filter');
    const expected = await p.evaluate(({ selection, week, field }) => selection.reduce((sum, name) => sum + (week ? week.reduce((v, k) => v + (__slExpected.weekTotals[k]?.[name]?.[field] ?? 0), 0) : __slExpected.totals[name][field]), 0), { selection, week, field: def.field });
    eq(def.total, expected, 'Independent synthetic source totals match pivot grand total');
  }
  ok(actual.sourceSame, 'Selection preserves the original source block and source version'); return actual;
}
const sl = (p, id = 'campaign-slicer') => p.locator('.obj[data-id="' + id + '"]');
async function pick(p, title, id = 'campaign-slicer', additive = false) {
  const before = (await state(p)).undo; const timingStart = await p.evaluate(() => __slTimings.length);
  await sl(p, id).locator('.sl-item').filter({ hasText: new RegExp('^' + title + '$') }).click({ modifiers: additive ? ['Control'] : [] }); await raf(p);
  const after = await state(p); eq(after.undo, before + 1, 'A changed slicer selection is exactly one Undo transaction'); ok(await sl(p, id).evaluate(node => node.classList.contains('sel')), 'Changed slicer has a visible object selection border');
  const metric = await p.evaluate(() => __slTrace.at(-1)); ok(metric?.finished, 'Mouse handler completion was measured');
  ok(metric.frameMs >= metric.handlerMs, 'Paint opportunity follows completed synchronous handler'); metric.layoutCalls = await p.evaluate(start => __slTimings.slice(start).filter(t => t.name === 'layout').length, timingStart); if (!baseline) eq(metric.layoutCalls, 1, 'A changed click produces exactly one layout'); return metric;
}
async function test(name, fn) {
  if (only && !only.split('|').some(filter => name.includes(filter))) return;
  const context = await browser.newContext({ viewport: { width: 1450, height: 1050 }, serviceWorkers: 'block' });
  const p = await context.newPage(), info = {}, pageErrors = [], caseWrites = []; p.setDefaultTimeout(20000);
  p.on('pageerror', e => { errors.push({ name, message: e.message }); pageErrors.push(e.message); }); p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {} } }); });
  await context.route('**/*', route => { const request = route.request(), u = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push({ name, method: request.method(), path: u.pathname }); caseWrites.push(u.pathname); return route.abort(); }
    if (u.origin !== origin || /^\/api(?:\/|$)/.test(u.pathname)) { blocked.push({ name, path: u.pathname }); return route.abort(); }
    if (overrides.has(u.pathname)) return route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: overrides.get(u.pathname) }); return route.continue(); });
  const start = checks;
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    for (const src of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => n.getAttribute('src')))) assets.add(src);
    await p.evaluate(() => {
      const t = tabula, w = t.wb(), g = t.gv(); window.__slTrace = []; window.__slTimings = []; window.__slLongTasks = [];
      window.__slState = () => ({ undo: w.undoStack.length, redo: w.redoStack.length, version: w.version,
        pivots: w.sheets.flatMap((sh, si) => [sh.pivot, ...(sh.pivotsExtra ?? [])].filter(Boolean).map(d => ({ si, name: d.name, filters: structuredClone(d.filters ?? {}), field: d.values[0].field, area: structuredClone(d.area), total: d.area ? w.getValue(si, d.area.r2, d.area.c2) : null }))),
        selected: Object.fromEntries(['campaign-slicer', 'week-slicer'].map(id => [id, [...document.querySelectorAll('.obj[data-id="' + id + '"] .sl-item.on')].map(n => n.textContent.trim())])),
        sourceSame: !!window.__slSourceOriginal && w.sheets[0] === __slSourceOriginal.sheet && w.sheets[0].blocks[0] === __slSourceOriginal.block && w.sheetVersion(0) === __slSourceOriginal.ver });
      window.__slReportState = () => w.sheets.map((sh, si) => si ? (() => { const { _sid, fileValues, ...data } = w.serializeSheet(si); return data; })() : null);
      for (const [object, names, owner] of [[w, ['transact', 'rangeRead'], 'wb'], [g, ['layout', 'renderAll', 'renderObjectsAll'], 'grid'], [g.host, ['slicerModel'], 'host']]) for (const name of names) {
        const original = object[name]; if (typeof original !== 'function') continue;
        object[name] = function (...args) { const start = performance.now(), before = w.undoStack.length;
          try { return original.apply(this, args); } finally { __slTimings.push({ owner, name, ms: performance.now() - start, before, after: w.undoStack.length }); } };
      }
      if (typeof PerformanceObserver === 'function') try { new PerformanceObserver(list => __slLongTasks.push(...list.getEntries().map(e => ({ start: e.startTime, duration: e.duration })))).observe({ type: 'longtask', buffered: true }); } catch {}
      const pending = new WeakMap();
      const finish = (e, phase) => { const record = pending.get(e); if (!record || record.finished) return;
        record.finished = true; record.phase = phase; record.handlerMs = performance.now() - record.at; record.prevented = e.defaultPrevented; record.after = __slState(); if (record.feedbackMs === undefined && JSON.stringify(record.after.selected[record.id]) !== JSON.stringify(record.before.selected[record.id])) record.feedbackMs = performance.now() - record.at;
        requestAnimationFrame(() => requestAnimationFrame(() => { record.frameMs = performance.now() - record.at; record.frameState = __slState(); })); };
      window.addEventListener('mousedown', e => { const button = e.target.closest?.('.sl-item,.sl-clear,.sl-multi'); if (!button) return;
        const record = { at: performance.now(), eventStamp: e.timeStamp, trusted: e.isTrusted, id: button.closest('.obj')?.dataset.id, key: button.dataset.k, title: button.textContent.trim(), action: button.className, additive: e.ctrlKey, before: __slState() };
        __slTrace.push(record); window.__slCurrent = record; pending.set(e, record); setTimeout(() => finish(e, 'async-fallback'), 0);
      }, true);
      g.viewEl.addEventListener('mousedown', e => finish(e, 'view-bubble-after-app')); window.addEventListener('mousedown', e => finish(e, 'window-bubble-after-app'));
      const observer = new MutationObserver(() => { const record = window.__slCurrent; if (record && record.feedbackMs === undefined && JSON.stringify(__slState().selected[record.id]) !== JSON.stringify(record.before.selected[record.id])) record.feedbackMs = performance.now() - record.at; });
      observer.observe(g.viewEl, { subtree: true, childList: true, attributes: true, attributeFilter: ['class'] });
    });
    await fn(p, info); await raf(p); await p.waitForTimeout(60);
    eq(pageErrors, [], 'No application exception'); eq(caseWrites, [], 'No API or remote writes');
    const data = await p.evaluate(() => ({ trace: __slTrace, timings: __slTimings, longTasks: __slLongTasks, final: __slState() }));
    results.push({ name, ok: true, checks: checks - start, info, ...data }); console.log('OK ' + name);
  } catch (error) {
    const data = await p.evaluate(() => ({ trace: window.__slTrace, timings: window.__slTimings, longTasks: window.__slLongTasks, final: window.__slState?.() })).catch(() => ({}));
    results.push({ name, ok: false, checks: checks - start, error: error.message, info, ...data }); console.error('NG ' + name + ': ' + error.message.split('\n')[0]); await p.screenshot({ path: path.join(out, 'failure-' + results.length + '.png') }).catch(() => {});
  } finally { await context.close(); }
}
try {
  for (const linked of [1, 4]) await test('large-source-first-and-warm-click-' + linked + '-linked-pivots', async (p, info) => {
    info.fixture = await fixture(p, { linked }); info.operations = [];
    for (const title of ['Campaign A', 'Campaign B', 'Campaign C', 'Campaign A', 'Campaign D', 'Campaign B']) { info.operations.push({ title, metric: await pick(p, title) }); await selected(p, [title]); }
    const before = await state(p); await sl(p).locator('.sl-clear').click(); await raf(p); eq((await state(p)).undo, before.undo + 1, 'Clear is one transaction'); await selected(p);
    await run(p, 'undo'); await selected(p, ['Campaign B']); await run(p, 'redo'); await selected(p);
    info.firstHandlerMs = info.operations[0].metric.handlerMs; info.warmHandlerMs = info.operations.slice(1).map(o => o.metric.handlerMs);
    info.scope = 'First interaction in a fresh context after initial pivot and slicer display; setup already reads the source. Timing is diagnostic and has no machine-dependent pass threshold.';
    await p.screenshot({ path: path.join(out, 'large-' + linked + '-linked-pivots.png') });
  });
  await test('below-rollup-threshold-repeated-selection', async (p, info) => {
    info.fixture = await fixture(p, { n: Math.min(sourceRows, 120000), linked: 4 }); info.operations = [];
    for (const title of ['Campaign A', 'Campaign B', 'Campaign C']) { info.operations.push(await pick(p, title)); await selected(p, [title]); }
  });
  await test('multi-slicer-cross-filter-and-additive-undo', async (p, info) => {
    info.fixture = await fixture(p, { n: Math.min(sourceRows, 120000), linked: 4, extraSlicer: true });
    await pick(p, 'Week B', 'week-slicer'); await selected(p, null, ['Week B']);
    await pick(p, 'Campaign A'); await selected(p, ['Campaign A'], ['Week B']);
    await pick(p, 'Campaign C', 'campaign-slicer', true); await selected(p, ['Campaign A', 'Campaign C'], ['Week B']);
    await run(p, 'undo'); await selected(p, ['Campaign A'], ['Week B']); await run(p, 'redo'); await selected(p, ['Campaign A', 'Campaign C'], ['Week B']);
    await pick(p, 'Campaign F'); await selected(p, ['Campaign F'], ['Week B']);
  });
  await test('same-task-rapid-selection-delivers-every-change', async (p, info) => {
    info.fixture = await fixture(p, { n: Math.min(sourceRows, 120000), linked: 4 });
    info.operations = await p.evaluate(() => ['Campaign A', 'Campaign B', 'Campaign C', 'Campaign A', 'Campaign D', 'Campaign E', 'Campaign F', 'Campaign A'].map(title => {
      const before = __slState().undo, at = performance.now(), button = [...document.querySelectorAll('.obj[data-id="campaign-slicer"] .sl-item')].find(b => b.textContent.trim() === title);
      const event = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 }); button.dispatchEvent(event); return { title, before, after: __slState().undo, ms: performance.now() - at, prevented: event.defaultPrevented };
    })); await raf(p); eq((await state(p)).undo, 8, 'All eight changed selections survive queued renders'); ok(info.operations.every(o => o.after === o.before + 1), 'No same-task selection is lost'); await selected(p, ['Campaign A']);
    for (const title of ['Campaign F', 'Campaign E', 'Campaign D', 'Campaign A', 'Campaign C', 'Campaign B', 'Campaign A', null]) { await run(p, 'undo'); await selected(p, title ? [title] : null); }
  });
  await test('trusted-busy-queued-selection-last-result-and-history', async (p, info) => {
    info.fixture = await fixture(p, { n: Math.min(sourceRows, 120000), linked: 4 }); const titles = ['Campaign A', 'Campaign B', 'Campaign C', 'Campaign D', 'Campaign E', 'Campaign F', 'Campaign A', 'Campaign B'];
    const points = await sl(p).locator('.sl-item').evaluateAll(nodes => Object.fromEntries(nodes.map(n => { const r = n.getBoundingClientRect(); return [n.textContent.trim(), { x: r.x + r.width / 2, y: r.y + r.height / 2 }]; })));
    const cdp = await p.context().newCDPSession(p); await p.evaluate(() => setTimeout(() => { const end = performance.now() + 180; while (performance.now() < end) {} }, 0));
    const requests = [], start = Date.now(); for (const title of titles) for (const type of ['mousePressed', 'mouseReleased']) requests.push(cdp.send('Input.dispatchMouseEvent', { type, ...points[title], button: 'left', clickCount: 1 }));
    await Promise.all(requests); await raf(p); info.wallMs = Date.now() - start;
    const trace = await p.evaluate(() => __slTrace); eq(trace.length, 8, 'Browser delivers all eight requested mousedown events'); ok(trace.every(e => e.trusted), 'Queued events use browser trusted dispatch'); eq((await state(p)).undo, 8, 'Each queued changed selection is one history entry'); await selected(p, ['Campaign B']);
    info.limit = 'Protocol input during a deliberately busy main thread does not establish physical device event loss.';
  });
  await test('duplicated-connections-apply-each-pivot-once', async (p, info) => {
    info.fixture = await fixture(p, { n: 2400, linked: 4, duplicated: true }); await pick(p, 'Campaign C'); await selected(p, ['Campaign C']);
    const counts = await p.evaluate(() => ({ linked: tabula.wb().sheets[1].slicers[0].source.pivots.length, actual: __slState().pivots.length })); eq(counts, { linked: 8, actual: 4 }, 'Fixture has eight connections pointing to four unique pivots');
    await run(p, 'undo'); await selected(p); eq((await state(p)).undo, 0, 'One Undo restores all deduplicated targets');
  });
  await test('linked-pivot-overlap-rejection-is-atomic', async (p, info) => {
    info.fixture = await fixture(p, { n: 2400, linked: 3, reject: true }); await selected(p, ['Campaign A']);
    const before = await p.evaluate(() => ({ reports: __slReportState(), state: __slState(), snapshots: [...(tabula.wb().pivotSnapshots ?? [])] }));
    await sl(p).locator('.sl-clear').click(); await raf(p);
    const after = await p.evaluate(() => ({ reports: __slReportState(), state: __slState(), snapshots: [...(tabula.wb().pivotSnapshots ?? [])] }));
    eq(after, before, 'Rejected clear preserves all pivot definitions, cells, formatting, snapshots, source and Undo/Redo');
    info.toast = await p.locator('#toast').textContent(); ok(/겹/.test(info.toast), 'Overlap rejection is visible'); ok(await sl(p).evaluate(node => node.classList.contains('sel')), 'Overlap fallback paints the object selection border'); await selected(p, ['Campaign A']);
    await p.screenshot({ path: path.join(out, 'overlap-rejection.png') });
  });

  await test('large-output-six-linked-pivots-first-and-warm', async (p, info) => {
    info.fixture = await fixture(p, { n: sourceRows, linked: 6, rowGroups: 800, channelGroups: 200 }); info.operations = [];
    for (const title of ['Campaign A', 'Campaign B', 'Campaign A']) { info.operations.push(await pick(p, title)); await selected(p, [title]); }
    await run(p, 'undo'); await selected(p, ['Campaign B']); await run(p, 'redo'); await selected(p, ['Campaign A']);
  });
  await test('no-op-clear-keeps-object-selected-without-history', async (p, info) => {
    info.fixture = await fixture(p, { n: 2400, linked: 3 });
    const before = await p.evaluate(() => ({ reports: __slReportState(), state: __slState() }));
    await sl(p).locator('.sl-clear').click(); await raf(p);
    const after = await p.evaluate(() => ({ reports: __slReportState(), state: __slState() }));
    eq(after, before, 'Clearing an already clear slicer edits no definition, cell or history');
    eq(await p.evaluate(() => tabula.gv().host.state().chartSel), 'campaign-slicer', 'No-op still selects the slicer object'); ok(await sl(p).evaluate(node => node.classList.contains('sel')), 'No-op fallback paints the object selection border');
    ok(await sl(p).isVisible(), 'No-op fallback leaves slicer visible'); await selected(p);
  });
  await test('protected-connected-report-denial-keeps-selection-and-history', async (p, info) => {
    info.fixture = await fixture(p, { n: 2400, linked: 3 });
    await p.evaluate(() => { const t = tabula; t.wb().sheets[3].protect = { on: true, allow: { pivotTables: false } }; t.gv().layout(); });
    const before = await p.evaluate(() => ({ reports: __slReportState(), state: __slState() }));
    await sl(p).locator('.sl-item').filter({ hasText: /^Campaign A$/ }).click(); await raf(p);
    const after = await p.evaluate(() => ({ reports: __slReportState(), state: __slState() }));
    eq(after, before, 'Protection on a later linked report rejects the entire filter change');
    eq(await p.evaluate(() => tabula.gv().host.state().chartSel), 'campaign-slicer', 'Rejected filtering still selects the slicer object'); ok(await sl(p).evaluate(node => node.classList.contains('sel')), 'Protection fallback paints the object selection border');
    info.toast = await p.locator('#toast').textContent(); ok(/권한/.test(info.toast), 'Protection denial is visible'); await selected(p);
  });
  await test('broken-link-clear-keeps-object-selected-without-history', async (p, info) => {
    info.fixture = await fixture(p, { n: 2400, linked: 1 });
    await p.evaluate(() => { const t = tabula; t.wb().sheets[1].slicers[0].source.pivots = [{ sheet: 'Missing Report', name: 'Missing Pivot' }]; t.gv().layout(); });
    ok(await sl(p).locator('.sl-broken').isVisible(), 'Broken source is displayed');
    const before = await p.evaluate(() => ({ reports: __slReportState(), state: __slState() }));
    await sl(p).locator('.sl-clear').click(); await raf(p);
    const after = await p.evaluate(() => ({ reports: __slReportState(), state: __slState() }));
    eq(after, before, 'Broken source fallback does not mutate cells or history');
    eq(await p.evaluate(() => tabula.gv().host.state().chartSel), 'campaign-slicer', 'Broken slicer remains selectable'); ok(await sl(p).evaluate(node => node.classList.contains('sel')), 'Broken source fallback paints the object selection border');
    ok(await sl(p).locator('.sl-broken').isVisible(), 'Broken source feedback is retained');
  });
  await test('selected-single-click-clears-and-keyboard-clear-is-one-render', async (p, info) => {
    info.fixture = await fixture(p, { n: 2400, linked: 3 }); await pick(p, 'Campaign A'); await selected(p, ['Campaign A']);
    await pick(p, 'Campaign A'); await selected(p); await pick(p, 'Campaign C'); await selected(p, ['Campaign C']);
    const start = await p.evaluate(() => __slTimings.length), before = await state(p);
    await p.keyboard.press('Alt+c'); await raf(p); await selected(p); eq((await state(p)).undo, before.undo + 1, 'Alt+C clears selected slicer in one Undo');
    const layouts = await p.evaluate(start => __slTimings.slice(start).filter(t => t.name === 'layout').length, start);
    info.keyboardLayouts = layouts; if (!baseline) eq(layouts, 1, 'Changed keyboard clear produces exactly one layout');
    await run(p, 'undo'); await selected(p, ['Campaign C']);
  });

} finally {
  await browser.close();
  const summary = { url, baseline: baseline || null, syntheticOnly: true, rows: sourceRows, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, pageErrors: errors, remoteWrites: writes, blockedRequests: blocked, results };
  await writeFile(path.join(out, 'result.json'), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify({ cases: summary.cases, passed: summary.passed, checks, assets: summary.assets, pageErrors: errors.length, remoteWrites: writes.length, out }));
}
if (results.some(result => !result.ok) || errors.length || writes.length) process.exitCode = 1;
