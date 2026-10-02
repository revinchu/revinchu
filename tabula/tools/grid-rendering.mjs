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
let checks = 0, cases = 0, failureCount = 0; const errors = [], writes = [], timings = [], runs = [], failures = [];
function eq(a, b, msg) {
  checks++;
  try { assert.deepEqual(a, b); return true; }
  catch { failureCount++; if (failures.length < 20) failures.push({ message: msg, actual: a, expected: b }); return false; }
}
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
// 기대 색과 thin=1px는 렌더러의 폭 계산식을 가져오지 않은 독립 기준이다.
function verifyPaintedBorders(im, info, dpr, label) {
  const is = (pixel, rgb) => pixel.length === 3 && pixel.every((n, i) => n === rgb[i]);
  const black = [0, 0, 0], gray = [208, 206, 206];
  eq(info.overflow, [true, true, true], label + ' filled text uses overflow layer');
  eq(info.fills >= 18, true, label + ' solid/pattern/gradient fill layers exist');
  const y1 = Math.ceil(info.top.top * dpr) + 1, y2 = Math.floor(info.last.bottom * dpr) - 1;
  const blackColumns = [];
  for (const rect of info.vertical) {
    const center = Math.floor(rect.right * dpr), hits = [];
    for (let x = center - 2; x <= center + 1; x++) {
      let n = 0; for (let y = y1; y <= y2; y++) if (is(im.at(x, y), black)) n++;
      if (n > (y2 - y1 + 1) * .8) hits.push({ x, n });
    }
    eq(hits.length, 1, label + ' black vertical is one physical pixel');
    if (hits.length) {
      eq(hits[0].n, y2 - y1 + 1, label + ' black vertical uninterrupted at every row/crossing');
      blackColumns.push(hits[0].x);
    }
  }
  const x1 = Math.ceil(info.top.left * dpr) + 1, x2 = Math.floor(info.last.right * dpr) - 1;
  for (const rect of info.horizontal) {
    const center = Math.floor(rect.bottom * dpr), sampleX = Math.floor((rect.left + rect.width / 2) * dpr), ys = [];
    for (let y = center - 2; y <= center + 1; y++) if (is(im.at(sampleX, y), gray)) ys.push(y);
    eq(ys.length, 1, label + ' gray horizontal is one physical pixel');
    if (!ys.length) continue;
    const bad = [];
    for (let x = x1; x <= x2; x++) {
      const expected = blackColumns.includes(x) ? black : gray;
      if (!is(im.at(x, ys[0]), expected)) { if (bad.length < 10) bad.push({ x, y: ys[0], actual: im.at(x, ys[0]), expected }); }
    }
    eq(bad, [], label + ' entire horizontal and dark intersections stay visible');
  }
}
const browser = await chromium.launch();
try {
  for (const dpr of [1, 1.25, 1.5, 2]) {
    const context = await browser.newContext({ viewport: { width: 1600, height: 1100 }, deviceScaleFactor: dpr });
    const page = await context.newPage();
    page.on('pageerror', e => errors.push(e.message));
    await context.route('**/*', route => {
      if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort('blockedbyclient'); }
      const u = new URL(route.request().url());
      if (u.origin !== new URL(target).origin || u.pathname.startsWith('/api/')) return route.abort('blockedbyclient');
      return route.continue();
    });
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => !!window.tabula?.gv());
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
        // 글자가 있는 .ovf 셀의 배경이 선을 덮는 회귀: 단색·무늬·그라데이션을 나란히 배치.
        const fills = [{ fill: '#fce4d6' }, { fill: '#e2f0d9', pattern: 'lightDown', patternColor: '#a9d18e' }, { gradient: { deg: 0, stops: [[0, '#ddebf7'], [1, '#bdd7ee']] } }];
        for (let r = 16; r <= 21; r++) for (let c = 1; c <= 6; c++) {
          wb.setStyle(0, r, c, { ...fills[Math.floor((c - 1) / 2)], align: 'left', color: '#000000', bb: true, bbc: '#D0CECE', ...(c % 2 ? { br: true, brc: '#000000' } : {}) });
          if (c % 2) wb.setInput(0, r, c, 'Text');
        }
        const patterns = ['hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot'];
        patterns.forEach((kind, i) => wb.setStyle(0, 2 + Math.floor(i / 5) * 2, 5 + i % 5, { fill: '#ffffff', bb: true, bbs: kind, bbc: '#11223' + i }));
        wb.setSheetProp(0, 'merges', [{ r1: 13, c1: 1, r2: 14, c2: 3 }]); wb.setStyle(0, 13, 1, border);
        wb.setSheetProp(0, 'filter', { r1: 6, c1: 1, r2: 10, c2: 3 });
      });
      document.querySelectorAll('.overlay').forEach(x => x.style.visibility = 'hidden');
    });
    for (const theme of ['white', 'black']) for (const zoom of [50, 75, 80, 100, 125]) {
      const info = await page.evaluate(({ theme, zoom }) => {
        document.documentElement.className = 'theme-' + theme;
        const gv = window.tabula.gv(); gv.setZoom(zoom); gv.setScroll(0, 0);
        const t = performance.now(); gv.renderAll(); const ms = performance.now() - t;
        const rect = (r, c) => document.querySelector(`.pane-br .c[data-r="${r}"][data-c="${c}"]`).getBoundingClientRect().toJSON();
        const f = document.querySelector('.pane-br .fbtn'), fr = f.getBoundingClientRect();
        const patterns = ['hair', 'dotted', 'dashed', 'dashDot', 'dashDotDot', 'mediumDashed', 'mediumDashDot', 'mediumDashDotDot', 'slantDashDot'].map((kind, i) => { const path = document.querySelector(`.pane-br .cell-borders path[stroke="#11223${i}"]`); return { kind, dash: path?.getAttribute('stroke-dasharray')?.split(' ').map(Number), width: Number(path?.getAttribute('stroke-width')) }; });
        return { ms, patterns, painted: { top: rect(16, 1), last: rect(21, 6), vertical: [1, 3, 5].map(c => rect(16, c)), horizontal: [16, 17, 18, 19, 20, 21].map(r => rect(r, 2)), overflow: [1, 3, 5].map(c => document.querySelector(`.pane-br .c[data-r="16"][data-c="${c}"]`).classList.contains('ovf')), fills: document.querySelectorAll('.pane-br .cell-fill').length }, thin: rect(8, 2), medium: rect(8, 4), thick: rect(8, 6), double: rect(8, 8), origin: rect(0, 0), fill: rect(2, 1), fillLast: rect(4, 3), merged: rect(13, 1), view: gv.viewEl.getBoundingClientRect().toJSON(), hw: gv.hw, hh: gv.hh, pe: getComputedStyle(document.querySelector('.cell-borders')).pointerEvents, button: document.elementFromPoint(fr.left + fr.width / 2, fr.top + fr.height / 2)?.closest('.fbtn') === f, hit: gv.hitTest(rect(8, 2).left + 10 * gv.z, rect(8, 2).top + 10 * gv.z) };
      }, { theme, zoom });
      await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
      const bytes = await page.screenshot(), im = png(bytes), scale = dpr * zoom / 100, label = `${theme} DPR${dpr} zoom${zoom}`;
      if (output) { await writeFile(join(output, `grid-${theme}-dpr${dpr}-zoom${zoom}.png`), bytes); await writeFile(join(output, `grid-${theme}-dpr${dpr}-zoom${zoom}.json`), JSON.stringify({ dpr, zoom, theme, info })); }
      timings.push(info.ms); cases++;
      verifyPaintedBorders(im, info.painted, dpr, label);
      const dashPatterns = Object.fromEntries(info.patterns.map(p => [p.kind, p.dash?.map(n => Number((n / p.width).toFixed(5)))]));
      eq(info.patterns.every(p => p.dash?.length && p.width > 0), true, label + ' all nine line patterns use SVG strokes');
      eq(dashPatterns.hair, [1, 1], label + ' hair distinct from dotted');
      eq(dashPatterns.dotted, [1, 2], label + ' dotted spacing');
      eq(dashPatterns.dashDot?.length, 4, label + ' dash-dot has dash plus dot');
      eq(dashPatterns.dashDotDot?.length, 6, label + ' dash-dot-dot has two dots');
      eq(dashPatterns.mediumDashDot?.length, 4, label + ' medium dash-dot');
      eq(dashPatterns.mediumDashDotDot?.length, 6, label + ' medium dash-dot-dot');
      eq(JSON.stringify(dashPatterns.slantDashDot) !== JSON.stringify(dashPatterns.mediumDashDot), true, label + ' slant pattern distinct');
      for (const [kind, width] of [['thin', 1], ['medium', 2], ['thick', 3]]) for (const side of ['bottom', 'right']) blackWidth(im, info[kind], side, dpr, kind === 'thin' ? 1 : Math.max(1, Math.round(width * scale)), label + ' ' + kind + ' ' + side);
      blackWidth(im, info.origin, 'top', dpr, 1, label + ' first row');
      blackWidth(im, info.origin, 'left', dpr, 1, label + ' first column');
      blackWidth(im, info.merged, 'bottom', dpr, 1, label + ' merged');
      const double = scan(im, (info.double.left + info.double.width / 2) * dpr, info.double.bottom * dpr, false, 5);
      const doublePixels = double.map((p, i) => Math.max(...p) < 32 ? i : -1).filter(i => i >= 0);
      eq(doublePixels.length, 2, label + ' double has two one-pixel strokes');
      eq(doublePixels[1] - doublePixels[0], 2, label + ' double has exactly one blank pixel between strokes (total 3px)');
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
      runs.push({ dpr, zoom, theme, thinDevicePixels: 1, gridDevicePixels: 1 });
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
      return { hidden: document.querySelectorAll('.c[data-r="7"],.c[data-c="2"]').length, panes: gv.panes.filter(p => p.win).length, hit: [hit.r, hit.c], selection: Math.abs(sel.left - r.left) < 2 && Math.abs(sel.top - r.top) < 2, rect: r.toJSON(), svg: document.querySelector('.pane-br .cell-borders').getBoundingClientRect().toJSON(), paths: [...document.querySelectorAll('.pane-br .cell-borders path')].map(p => ({ d: p.getAttribute('d'), stroke: p.getAttribute('stroke'), width: p.getAttribute('stroke-width') })), pointer: getComputedStyle(document.querySelector('.pane-br .cell-borders')).pointerEvents };
    });
    eq(extra.hidden, 0, 'hidden axes absent'); eq(extra.panes, 4, 'freeze panes'); eq(extra.hit, [8, 3], 'scrolled hit'); eq(extra.selection, true, 'selection aligned'); eq(extra.pointer, 'none', 'scrolled pointer');
    await page.evaluate(() => document.querySelectorAll('.overlay').forEach(x => x.style.visibility = 'hidden'));
    const scrolledBytes = await page.screenshot(), scrolled = png(scrolledBytes);
    if (output) { await writeFile(join(output, `grid-scrolled-dpr${dpr}.png`), scrolledBytes); await writeFile(join(output, `grid-scrolled-dpr${dpr}.json`), JSON.stringify({ dpr, zoom: 75, info: extra })); }
    blackWidth(scrolled, extra.rect, 'bottom', dpr, 1, 'fractional scroll hidden/freeze DPR' + dpr);
    await page.keyboard.press('F2');
    const editor = await page.locator('#cellEditor').evaluate(el => ({ idle: el.classList.contains('idle'), rect: el.getBoundingClientRect().toJSON() }));
    eq(editor.idle, false, 'edit starts'); eq(Math.abs(editor.rect.left - extra.rect.left) < 4 && Math.abs(editor.rect.top - extra.rect.top) < 4, true, 'editor aligned');
    await page.keyboard.press('Escape');
    const noGrid = await page.evaluate(() => { const wb = window.tabula.wb(), gv = window.tabula.gv(); wb.transact(() => wb.setSheetProp(0, 'noGrid', true)); gv.renderAll(); return { grid: document.querySelectorAll('.gl svg').length, borders: document.querySelectorAll('.cell-borders rect, .cell-borders path').length }; });
    eq(noGrid.grid, 0, 'noGrid hides only grid'); eq(noGrid.borders > 0, true, 'noGrid retains cell borders');
    await context.close();
  }
  eq(errors, [], 'page errors'); eq(writes, [], 'server writes');
  timings.sort((a, b) => a - b);
  const edges = [{ vertical: false, at: 20, start: 0, end: 20000, width: 1, pattern: 'solid', color: '#000' }];
  for (let i = 0; i < 20000; i++) edges.push({ ...edges[0], start: i, end: i + 1, width: i % 2 ? 1 : 2 });
  const start = performance.now(), resolved = resolveGridBorders(edges), resolverMs = performance.now() - start;
  eq(resolved.length, 20000, 'long merged boundary segments');
  console.log(JSON.stringify({ ok: failureCount === 0, failureCount, failures, cases, checks, runs, mergedBoundary: { edges: edges.length, segments: resolved.length, resolverMs }, renderMs: { min: timings[0], median: timings[Math.floor(timings.length / 2)], max: timings.at(-1) }, pageErrors: errors, blockedWrites: writes }));
  assert.equal(failureCount, 0, 'grid rendering assertions failed; see JSON failures above');
} finally { await browser.close(); }
