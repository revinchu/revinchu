import assert from 'node:assert/strict';
import { inflateSync } from 'node:zlib';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { resolveGridBorders } from '../src/grid-lines.js';

// 합성 시트만 사용. PNG의 실제 장치 픽셀을 읽어 확대/축소 선 두께를 검사합니다.
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const target = process.env.WIXEL_URL || 'http://localhost:5178/';
const output = process.env.GRID_SCREENSHOTS;
if (output) await mkdir(output, { recursive: true });
function png(bytes) {
  let w, h, type; const parts = [];
  for (let p = 8; p < bytes.length;) {
    const n = bytes.readUInt32BE(p), t = bytes.toString('ascii', p + 4, p + 8), d = bytes.subarray(p + 8, p + 8 + n);
    if (t === 'IHDR') { w = d.readUInt32BE(0); h = d.readUInt32BE(4); type = d[9]; assert.equal(d[8], 8); }
    if (t === 'IDAT') parts.push(d);
    p += n + 12;
  }
  assert.ok(type === 2 || type === 6);
  const bpp = type === 6 ? 4 : 3, raw = inflateSync(Buffer.concat(parts)), out = Buffer.alloc(w * h * bpp), stride = w * bpp;
  const paeth = (a, b, c) => { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); return pa <= pb && pa <= pc ? a : pb <= pc ? b : c; };
  let k = 0;
  for (let y = 0; y < h; y++) { const f = raw[k++]; assert.ok(f <= 4); for (let x = 0; x < stride; x++) {
    const i = y * stride + x, a = x >= bpp ? out[i - bpp] : 0, b = y ? out[i - stride] : 0, c = y && x >= bpp ? out[i - stride - bpp] : 0;
    out[i] = (raw[k++] + (f === 0 ? 0 : f === 1 ? a : f === 2 ? b : f === 3 ? Math.floor((a + b) / 2) : paeth(a, b, c))) & 255;
  } }
  return { at(x, y) { return [...out.subarray((Math.floor(y) * w + Math.floor(x)) * bpp, (Math.floor(y) * w + Math.floor(x)) * bpp + 3)]; } };
}
let checks = 0, cases = 0; const errors = [], writes = [], timings = [], runs = [];
function eq(a, b, msg) { assert.deepEqual(a, b, msg); checks++; }
function scan(im, x, y, vertical, radius) {
  const a = []; for (let k = -radius; k <= radius; k++) a.push(im.at(x + (vertical ? k : 0), y + (vertical ? 0 : k)));
  return a;
}
function blackWidth(im, rect, side, dpr, expected, label) {
  const vertical = side === 'left' || side === 'right';
  const x = (vertical ? rect[side] : rect.left + rect.width / 2) * dpr;
  const y = (vertical ? rect.top + rect.height / 2 : rect[side]) * dpr;
  const pixels = scan(im, x, y, vertical, expected + 2), dark = pixels.filter(p => Math.max(...p) < 32).length;
  eq(dark, expected, label + ' actual black width: ' + JSON.stringify(pixels));
  if (side === 'bottom' || side === 'right') eq(pixels.every(p => p.every(v => v === 0) || p.every(v => v === 255)), true, label + ' no antialias fringe');
}
const browser = await chromium.launch();
try {
  for (const dpr of [1, 1.25, 2]) {
    const context = await browser.newContext({ viewport: { width: 1600, height: 1100 }, deviceScaleFactor: dpr });
    const page = await context.newPage();
    page.on('pageerror', e => errors.push(e.message));
    await context.route('**/*', route => {
      if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort('blockedbyclient'); }
      return route.continue();
    });
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(target); await page.waitForFunction(() => !!window.tabula?.gv());
    await page.evaluate(() => {
      const wb = window.tabula.wb(); wb.load({ sheets: [{ name: '합성 선 검증', cells: {} }] });
      wb.transact(() => {
        wb.setSheetProp(0, 'defColW', 73); wb.setSheetProp(0, 'defRowH', 23);
        const border = { fill: '#ffffff', bt: true, bb: true, bl: true, br: true };
        wb.setStyle(0, 0, 0, border);
        for (let r = 2; r <= 4; r++) for (let c = 1; c <= 3; c++) wb.setStyle(0, r, c, { fill: '#ffffff' });
        for (let r = 6; r <= 10; r++) for (let c = 1; c <= 9; c++) {
          const kind = c <= 3 ? 'thin' : c <= 5 ? 'medium' : c <= 7 ? 'thick' : 'double';
          wb.setStyle(0, r, c, { ...border, bts: kind, bbs: kind, bls: kind, brs: kind });
        }
        wb.setSheetProp(0, 'merges', [{ r1: 13, c1: 1, r2: 14, c2: 3 }]); wb.setStyle(0, 13, 1, border);
        wb.setSheetProp(0, 'filter', { r1: 6, c1: 1, r2: 10, c2: 3 });
      });
      document.querySelectorAll('.overlay').forEach(x => x.style.visibility = 'hidden');
    });
    for (const theme of ['white', 'black']) for (const zoom of [50, 75, 100, 125]) {
      const info = await page.evaluate(({ theme, zoom }) => {
        document.documentElement.className = 'theme-' + theme;
        const gv = window.tabula.gv(); gv.setZoom(zoom); gv.setScroll(0, 0);
        const t = performance.now(); gv.renderAll(); const ms = performance.now() - t;
        const rect = (r, c) => document.querySelector(`.pane-br .c[data-r="${r}"][data-c="${c}"]`).getBoundingClientRect().toJSON();
        const f = document.querySelector('.pane-br .fbtn'), fr = f.getBoundingClientRect();
        return { ms, thin: rect(8, 2), medium: rect(8, 4), thick: rect(8, 6), double: rect(8, 8), origin: rect(0, 0), fill: rect(2, 1), fillLast: rect(4, 3), merged: rect(13, 1), view: gv.viewEl.getBoundingClientRect().toJSON(), hw: gv.hw, hh: gv.hh, pe: getComputedStyle(document.querySelector('.cell-borders')).pointerEvents, button: document.elementFromPoint(fr.left + fr.width / 2, fr.top + fr.height / 2)?.closest('.fbtn') === f, hit: gv.hitTest(rect(8, 2).left + 10 * gv.z, rect(8, 2).top + 10 * gv.z) };
      }, { theme, zoom });
      await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
      const bytes = await page.screenshot(), im = png(bytes), scale = dpr * zoom / 100, label = `${theme} DPR${dpr} zoom${zoom}`;
      if (output) await writeFile(join(output, `grid-${theme}-dpr${dpr}-zoom${zoom}.png`), bytes);
      timings.push(info.ms); cases++;
      for (const [kind, width] of [['thin', 1], ['medium', 2], ['thick', 3]]) for (const side of ['bottom', 'right']) blackWidth(im, info[kind], side, dpr, Math.max(1, Math.round(width * scale)), label + ' ' + kind + ' ' + side);
      blackWidth(im, info.origin, 'top', dpr, Math.max(1, Math.round(scale)), label + ' first row');
      blackWidth(im, info.origin, 'left', dpr, Math.max(1, Math.round(scale)), label + ' first column');
      blackWidth(im, info.merged, 'bottom', dpr, Math.max(1, Math.round(scale)), label + ' merged');
      const double = scan(im, (info.double.left + info.double.width / 2) * dpr, info.double.bottom * dpr, false, Math.max(3, Math.round(3 * scale)) + 2);
      eq(double.filter(p => Math.max(...p) < 32).length, 2 * Math.max(1, Math.floor(Math.max(3, Math.round(3 * scale)) / 3)), label + ' double two strokes');
      let bands = 0, on = false; for (const p of double) { const v = Math.max(...p) < 32; if (v && !on) bands++; on = v; } eq(bands, 2, label + ' double gap');
      for (const [r, side] of [[info.fill, 'top'], [info.fill, 'left'], [info.fillLast, 'bottom'], [info.fillLast, 'right']]) {
        const vertical = side === 'left' || side === 'right';
        const x = (vertical ? r[side] : r.left + r.width / 2) * dpr, y = (vertical ? r.top + r.height / 2 : r[side]) * dpr;
        eq(scan(im, x, y, vertical, 2).every(p => p.every(v => v === 255)), true, label + ' filled boundary ' + side + ' ' + JSON.stringify(scan(im, x, y, vertical, 2))); 
      }
      const gx = (info.view.left + (info.hw + 11 * 73 + 30) * zoom / 100) * dpr;
      const gy = (info.view.top + (info.hh + 3 * 23) * zoom / 100) * dpr;
      const grid = scan(im, gx, gy, false, 3);
      eq(grid.filter(p => p.some(v => v < 250)).length, 1, label + ' grid one device pixel');
      eq(info.pe, 'none', label + ' transparent hit testing'); eq(info.button, true, label + ' filter reachable'); eq([info.hit.zone, info.hit.r, info.hit.c], ['cell', 8, 2], label + ' hit test');
      runs.push({ dpr, zoom, theme, thinDevicePixels: Math.max(1, Math.round(scale)), gridDevicePixels: 1 });
    }
    // 숨김 축/고정/분수 스크롤의 실제 표시·클릭 좌표 회귀.
    const extra = await page.evaluate(() => {
      const wb = window.tabula.wb(), gv = window.tabula.gv();
      wb.transact(() => { wb.setSheetProp(0, 'hiddenRows', { 7: true }); wb.setSheetProp(0, 'hiddenCols', { 2: true }); wb.setSheetProp(0, 'freeze', { rows: 2, cols: 1 }); });
      gv.setZoom(75); gv.renderAll(); gv.setScroll(14.5, 17.25); gv.update(true);
      document.querySelectorAll('.overlay').forEach(x => x.style.visibility = '');
      window.tabula.selectCell(8, 3); gv.renderOverlays();
      const cell = document.querySelector('.pane-br .c[data-r="8"][data-c="3"]'), r = cell.getBoundingClientRect();
      const sel = document.querySelector('.pane-br .sel-border').getBoundingClientRect(), hit = gv.hitTest(r.left + r.width / 2, r.top + r.height / 2);
      return { hidden: document.querySelectorAll('.c[data-r="7"],.c[data-c="2"]').length, panes: gv.panes.filter(p => p.win).length, hit: [hit.r, hit.c], selection: Math.abs(sel.left - r.left) < 2 && Math.abs(sel.top - r.top) < 2, rect: r.toJSON(), pointer: getComputedStyle(document.querySelector('.pane-br .cell-borders')).pointerEvents };
    });
    eq(extra.hidden, 0, 'hidden axes absent'); eq(extra.panes, 4, 'freeze panes'); eq(extra.hit, [8, 3], 'scrolled hit'); eq(extra.selection, true, 'selection aligned'); eq(extra.pointer, 'none', 'scrolled pointer');
    await page.evaluate(() => document.querySelectorAll('.overlay').forEach(x => x.style.visibility = 'hidden'));
    const scrolled = png(await page.screenshot());
    blackWidth(scrolled, extra.rect, 'bottom', dpr, Math.max(1, Math.round(dpr * .75)), 'fractional scroll hidden/freeze DPR' + dpr);
    await page.keyboard.press('F2');
    const editor = await page.locator('#cellEditor').evaluate(el => ({ idle: el.classList.contains('idle'), rect: el.getBoundingClientRect().toJSON() }));
    eq(editor.idle, false, 'edit starts'); eq(Math.abs(editor.rect.left - extra.rect.left) < 4 && Math.abs(editor.rect.top - extra.rect.top) < 4, true, 'editor aligned');
    await page.keyboard.press('Escape');
    const noGrid = await page.evaluate(() => { const wb = window.tabula.wb(), gv = window.tabula.gv(); wb.transact(() => wb.setSheetProp(0, 'noGrid', true)); gv.renderAll(); return { grid: document.querySelectorAll('.gl svg').length, borders: document.querySelectorAll('.cell-borders rect').length }; });
    eq(noGrid.grid, 0, 'noGrid hides only grid'); eq(noGrid.borders > 0, true, 'noGrid retains cell borders');
    await context.close();
  }
  eq(errors, [], 'page errors'); eq(writes, [], 'server writes');
  timings.sort((a, b) => a - b);
  const edges = [{ vertical: false, at: 20, start: 0, end: 20000, width: 1, pattern: 'solid', color: '#000' }];
  for (let i = 0; i < 20000; i++) edges.push({ ...edges[0], start: i, end: i + 1, width: i % 2 ? 1 : 2 });
  const start = performance.now(), resolved = resolveGridBorders(edges), resolverMs = performance.now() - start;
  eq(resolved.length, 20000, 'long merged boundary segments');
  console.log(JSON.stringify({ ok: true, cases, checks, runs, mergedBoundary: { edges: edges.length, segments: resolved.length, resolverMs }, renderMs: { min: timings[0], median: timings[Math.floor(timings.length / 2)], max: timings.at(-1) }, pageErrors: errors, blockedWrites: writes }));
} finally { await browser.close(); }
