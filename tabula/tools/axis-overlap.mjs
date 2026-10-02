import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 합성 검사 전용입니다.');
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1400, height: 950 } });
const page = await context.newPage(), errors = [], writes = [];
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => { const q = route.request(), u = new URL(q.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { writes.push(q.method()); return route.abort(); } if (u.origin !== new URL(url).origin || u.pathname.startsWith('/api/')) return route.abort(); return route.continue(); });
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 30000 }); await page.waitForFunction(() => window.tabula?.gv());
  const results = await page.evaluate(() => {
    const t = window.tabula, wb = t.wb(), gv = t.gv(), cells = {};
    for (let r = 0; r < 520; r++) for (let c = 0; c < 7; c++) cells[r + ',' + c] = { raw: String(r * 10 + c) };
    const hidden = (start, length, period) => { const data = Uint8Array.from({ length }, (_, i) => +(i % period !== 0)); return { start, __bits: data, count: data.reduce((a,b) => a+b,0) }; };
    const table = (name, r1, c1, r2, h) => ({ id: name, name, r1, c1, r2, c2: c1 + 2, header: true, columns: [0,1,2].map(i => ({ name: name + i })), filter: { criteria: {}, hidden: h } });
    const a = hidden(1,320,5), b = hidden(51,449,7), beforeA = Array.from(a.__bits), beforeB = Array.from(b.__bits);
    wb.load({ sheets: [{ name: '두 표 합성 필터', cells, defRowH: 23, defColW: 90, rowHeights: { 0: 31, 10: 36, 80: 27, 450: 40 }, hiddenRows: { 500: true }, freeze: { rows: 1, cols: 1 }, tables: [table('왼쪽표',0,0,320,a),table('오른쪽표',50,4,499,b)] }] });
    t.selectCell(0,0); gv.setZoom(125); gv.layout();
    const check = (name, scroll) => {
      const s = wb.sheets[0]; let y = 0, checks = 0;
      for (let r = 0; r < 520; r++) {
        const hidden = !!s.hiddenRows[r] || s.tables.some(t => { const h=t.filter.hidden; return !!h.__bits && r >= h.start && r-h.start < h.__bits.length && !!h.__bits[r-h.start]; });
        const h = hidden ? 0 : s.rowHeights[r] ?? 23;
        if (gv.rows.size(r) !== h || gv.rows.pos(r) !== y) throw Error(name + ' size/position row ' + r); checks += 2;
        if (h > 0) { if (gv.rows.indexAt(y+h/2) !== r) throw Error(name + ' indexAt row ' + r); checks++; }
        y += h;
      }
      gv.setScroll(0, scroll); gv.renderAll();
      const first = [...document.querySelectorAll('.pane-br .c')].find(el => { const r=Number(el.dataset.r),c=Number(el.dataset.c),rect=el.getBoundingClientRect(), pane=el.closest('.pane').getBoundingClientRect();return r>gv.firstVisibleRow() && c===5 && rect.top>=pane.top && rect.bottom<pane.bottom; });
      if (!first) throw Error(name + ' visible cell missing');
      const rect = first.getBoundingClientRect(), hit = gv.hitTest(rect.left+5,rect.top+5);
      if (hit.r !== Number(first.dataset.r) || hit.c !== 5) throw Error(name + ' scrolled hit mismatch'); checks++;
      const cellRect = gv.clientRect({ r1: hit.r, c1: 5, r2: hit.r, c2: 5 });
      if (Math.abs(cellRect.top-rect.top) > 2) throw Error(name + ' rendered coordinate mismatch'); checks++;
      if (wb.getValue(0,510,6) !== 5106) throw Error(name + ' raw value changed'); checks++;
      return { name, checks, height: y, bits: gv.rows.bits.length, hitRow: hit.r };
    };
    const results = [check('overlapping table filters',200)];
    wb.transact(() => wb.setSheetProp(0,'tables',wb.sheets[0].tables.map((table,i)=>i===1?{...table,filter:{...table.filter,hidden:{}}}:table)), { si:0,sel:{...t.sel},active:{...t.active},selKind:'cells' }); gv.layout();
    results.push(check('second filter cleared',450));
    t.run('undo'); results.push(check('Undo restores both overlapping filters',200));
    t.run('redo'); results.push(check('Redo clears second filter again',450));
    if (JSON.stringify(Array.from(a.__bits)) !== JSON.stringify(beforeA) || JSON.stringify(Array.from(b.__bits)) !== JSON.stringify(beforeB)) throw Error('source bitmap mutated');
    return results;
  });
  assert.deepEqual(errors,[]); assert.deepEqual(writes,[]);
  assert.equal(results[0].bits,1); assert.equal(results[0].height,results[2].height); assert.equal(results[1].height,results[3].height); assert.ok(results[1].height>results[0].height);
  console.log(JSON.stringify({ url, scenarios: results.length, checks: results.reduce((n,r)=>n+r.checks,0)+7, results, pageErrors:errors.length, writes:writes.length },null,2));
} finally { await context.close(); await browser.close(); }
