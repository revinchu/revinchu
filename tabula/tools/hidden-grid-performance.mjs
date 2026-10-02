import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 합성 검사 전용입니다.');
const ref = process.env.WIXEL_PERF_BASELINE_REF, root = fileURLToPath(new URL('../../', import.meta.url));
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1400, height: 950 } });
const page = await context.newPage(), errors = [], writes = [], results = [];
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => {
  const request = route.request(), u = new URL(request.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method()); return route.abort(); }
  if (u.origin !== new URL(url).origin && !['data:', 'blob:'].includes(u.protocol)) return route.abort();
  return route.continue();
});
if (ref) for (const path of ['src/view.js', 'src/axis.js']) {
  const body = execFileSync('git', ['-c', 'safe.directory=' + resolve(root).replaceAll('\\', '/'), 'show', ref + ':tabula/' + path], { cwd: root, encoding: 'utf8' });
  await context.route('**/' + path, route => route.fulfill({ status: 200, contentType: 'text/javascript', body }));
}
let checks = 0;
const eq = (a, b, message) => { assert.deepEqual(a, b, message); checks++; };
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForFunction(() => !!window.tabula?.gv());
  for (const cfg of [
    { name: 'filtered-1million', n: 1000000, every: 25000, notes: false, freeze: false, zoom: 100 },
    { name: 'filtered-1million-visible-notes', n: 1000000, every: 25000, notes: true, freeze: false, zoom: 100 },
    { name: 'hidden-rows-cols-freeze-zoom-scroll', n: 60000, every: 500, notes: true, freeze: true, zoom: 125 },
  ]) {
    const result = await page.evaluate(cfg => {
      const t = window.tabula, wb = t.wb(), gv = t.gv(), cells = {}, bits = new Uint8Array(cfg.n); bits.fill(1);
      let count = cfg.n;
      for (let r = 0; r < cfg.n; r += cfg.every) { bits[r] = 0; count--; for (let c = 0; c < 18; c++) cells[r + ',' + c] = { raw: 'R' + r + 'C' + c, ...(cfg.notes && c === 2 ? { comment: '합성 메모 ' + r } : {}) }; }
      // 숨긴 행/열의 메모는 표시하지 않아야 한다.
      cells['1,2'] = { raw: '숨김', comment: '숨긴 행 메모' }; cells['0,3'] = { raw: '숨김', comment: '숨긴 열 메모' };
      wb.load({ sheets: [{ name: '합성 숨김 성능', cells, defRowH: 23, defColW: 80, hiddenCols: { 3: true, 4: true }, filter: { r1: 0, c1: 0, r2: cfg.n - 1, c2: 17, hidden: { __bits: bits, start: 0, count } }, freeze: cfg.freeze ? { rows: 1, cols: 1 } : {}, noteVisibility: cfg.notes ? { all: true } : null }] });
      t.selectCell(0, 0); gv.setZoom(cfg.zoom); gv.layout(); gv.setScroll(cfg.freeze ? 160 : 0, cfg.freeze ? 100 : 0); gv.renderAll();
      const times = []; for (let i = 0; i < 7; i++) { const start = performance.now(); gv.renderAll(); const elapsed = performance.now() - start; if (i >= 2) times.push(elapsed); }
      times.sort((a,b) => a-b);
      let calls = 0; const size = gv.rows.size; gv.rows.size = function(i) { calls++; return size.call(this, i); }; gv.renderAll(); gv.rows.size = size;
      const rendered = [...document.querySelectorAll('.pane .c')].map(el => [Number(el.dataset.r), Number(el.dataset.c), el.textContent]);
      const notes = [...document.querySelectorAll('.cell-note-visible')].map(el => [Number(el.dataset.noteR), Number(el.dataset.noteC), el.textContent]);
      const heads = [...document.querySelectorAll('.row-head .hr')].map(el => Number(el.textContent) - 1);
      const first = rendered.find(([r, c]) => r >= gv.firstVisibleRow() + cfg.every && c === 6);
      const rect = first ? gv.clientRect({ r1: first[0], c1: first[1], r2: first[0], c2: first[1] }) : null;
      const hit = rect ? gv.hitTest(rect.left + 8, rect.top + 8) : null;
      return { name: cfg.name, medianMs: +times[2].toFixed(2), rowSizeCalls: calls, rendered, notes, heads, hit, first, firstVisible: gv.firstVisibleRow(), value: wb.getValue(0, cfg.every, 2), filterCount: wb.sheets[0].filter.hidden.count };
    }, cfg);
    eq(result.rendered.every(([r,c]) => r % cfg.every === 0 && c !== 3 && c !== 4), true, cfg.name + ' hidden cells omitted');
    eq(result.heads.every(r => r % cfg.every === 0), true, cfg.name + ' visible row headers');
    eq(result.notes.every(([r,c]) => r % cfg.every === 0 && c === 2), true, cfg.name + ' visible notes only');
    eq(result.notes.length > 0, cfg.notes, cfg.name + ' notes enabled');
    eq(result.value, 'R' + cfg.every + 'C2', cfg.name + ' value unchanged');
    eq(result.filterCount, cfg.n - Math.ceil(cfg.n / cfg.every), cfg.name + ' source hidden bitmap unchanged');
    eq([result.hit.r, result.hit.c], result.first.slice(0,2), cfg.name + ' hit-test geometry');
    if (!ref) { assert.ok(result.rowSizeCalls < 30000, cfg.name + ' bounded visible work: ' + result.rowSizeCalls); checks++; }
    const signature = createHash('sha256').update(JSON.stringify([result.rendered,result.notes,result.heads])).digest('hex');
    results.push({ name: cfg.name, medianMs: result.medianMs, rowSizeCalls: result.rowSizeCalls, cells: result.rendered.length, notes: result.notes.length, signature });
  }
  const changed = await page.evaluate(() => {
    const t = window.tabula, wb = t.wb(), gv = t.gv(), firstAxis = gv.rows;
    const oldFilter = wb.sheets[0].filter;
    wb.transact(() => wb.setSheetProp(0, 'filter', { ...oldFilter, hidden: {} }), { si: 0, sel: { ...t.sel }, active: { ...t.active }, selKind: 'cells' }); gv.layout();
    const clearAxis = gv.rows, shown = clearAxis.size(501) > 0;
    t.run('undo'); const undoAxis = gv.rows, hiddenAgain = undoAxis.size(501) === 0;
    t.run('redo'); const redoAxis = gv.rows, reshown = redoAxis.size(501) > 0;
    t.selectRange({ r1: 1, c1: 0, r2: 2, c2: 0 }); t.run('hideRows');
    const hiddenAxis = gv.rows, rowsHidden = hiddenAxis.size(1) === 0 && hiddenAxis.size(2) === 0;
    t.run('undo'); const visibleAgain = gv.rows.size(1) > 0 && gv.rows.size(2) > 0;
    return { newAxes: firstAxis !== clearAxis && clearAxis !== undoAxis && undoAxis !== redoAxis && redoAxis !== hiddenAxis && hiddenAxis !== gv.rows, shown, hiddenAgain, reshown, rowsHidden, visibleAgain };
  });
  for (const [name, value] of Object.entries(changed)) eq(value, true, 'filter/hide/Undo refresh: ' + name);
  eq(errors, [], 'browser errors'); eq(writes, [], 'remote writes');
  console.log(JSON.stringify({ baseline: ref || null, url, checks, results }, null, 2));
} finally { await context.close(); await browser.close(); }
