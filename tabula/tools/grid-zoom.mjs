import assert from 'node:assert/strict';
import { inflateSync } from 'node:zlib';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
// 실제 확대/축소 입력 뒤 강제 renderAll/update 호출 없이 화면 픽셀을 검사한다.
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const target = process.env.WIXEL_URL || 'http://localhost:5178/';
const output = process.env.GRID_ZOOM_SCREENSHOTS;
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
const errors = [], writes = [], failures = [], runs = [];
let checks = 0, failed = 0;
function eq(actual, expected, message) {
  checks++;
  try { assert.deepEqual(actual, expected); } catch { failed++; if (failures.length < 30) failures.push({ message, actual, expected }); }
}
const countStroke = (image, rect, side, dpr) => {
  const vertical = side === 'left' || side === 'right';
  const x = (vertical ? rect[side] : (rect.left + rect.right) / 2) * dpr;
  const y = (vertical ? (rect.top + rect.bottom) / 2 : rect[side]) * dpr;
  const pixels = [];
  for (let i = -6; i <= 6; i++) pixels.push(image.at(x + (vertical ? i : 0), y + (vertical ? 0 : i)));
  return pixels.map((p, i) => p.length === 3 && Math.max(...p) < 32 ? i : -1).filter(i => i >= 0);
};
const browser = await chromium.launch();
try {
  for (const dpr of [1, 1.25, 1.5, 2]) {
    const context = await browser.newContext({ viewport: { width: 1800, height: 1200 }, deviceScaleFactor: dpr });
    await context.route('**/*', route => {
      const r = route.request(), u = new URL(r.url());
      if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); }
      if (u.origin !== new URL(target).origin || u.pathname.startsWith('/api/')) return route.abort();
      return route.continue();
    });
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    const page = await context.newPage(); page.on('pageerror', e => errors.push(e.message));
    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => !!window.tabula?.gv());
    for (const freeze of [{ name: 'none', rows: 0, cols: 0 }, { name: 'rows', rows: 1, cols: 0 }, { name: 'cols', rows: 0, cols: 1 }, { name: 'both', rows: 1, cols: 1 }]) {
      await page.evaluate(freeze => {
        const t = window.tabula, wb = t.wb();
        wb.load({ sheets: [{ name: '합성 확대 검사', cells: {} }] });
        wb.transact(() => {
          wb.setSheetProp(0, 'defColW', 80); wb.setSheetProp(0, 'defRowH', 22);
          wb.setSheetProp(0, 'freeze', { rows: freeze.rows, cols: freeze.cols });
          const borders = { fill: '#ffffff', bt: true, bb: true, bl: true, br: true };
          wb.setStyle(0, 0, 0, borders);
          wb.setStyle(0, 2, 1, borders);
          wb.setStyle(0, 2, 3, { ...borders, bts: 'double', bbs: 'double', bls: 'double', brs: 'double' });
          for (let c = 0; c < 3; c++) { wb.setInput(0, 4, c, 'ABC123'); wb.setStyle(0, 4, c, { fill: '#ffffff', size: 9, valign: 'middle', align: ['left', 'center', 'right'][c] }); }
        });
        // 기준100%를 처음 한 번만 그린다. 이후 slider change에는 강제 다시 그리기를 하지 않는다.
        const slider = document.querySelector('#zoomSlider'); slider.value = '100'; slider.dispatchEvent(new Event('input', { bubbles: true }));
        t.gv().setScroll(0, 0); t.gv().renderAll();
        document.querySelectorAll('.overlay').forEach(el => el.style.visibility = 'hidden');
      }, freeze);
      let baseline;
      for (const [step, zoom] of [100, 400, 100, 200, 50, 115].entries()) {
        const started = performance.now();
        await page.locator('#zoomSlider').evaluate((slider, zoom) => {
          slider.value = String(zoom);
          slider.dispatchEvent(new Event('input', { bubbles: true }));
          slider.dispatchEvent(new Event('change', { bubbles: true }));
        }, zoom);
        await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
        const info = await page.evaluate(() => {
          const gv = window.tabula.gv(), visible = (r, c) => [...document.querySelectorAll(`.c[data-r="${r}"][data-c="${c}"]`)].find(el => el.getBoundingClientRect().width);
          const rect = (r, c) => visible(r, c).getBoundingClientRect().toJSON();
          const text = [0, 1, 2].map(c => {
            const el = visible(4, c), span = el.querySelector(':scope > span'), box = el.getBoundingClientRect(), sr = span.getBoundingClientRect(), cs = getComputedStyle(el);
            return { col: c, dx: (sr.left - box.left) / gv.z, dy: (sr.top - box.top) / gv.z, right: (box.right - sr.right) / gv.z, width: sr.width / gv.z, height: sr.height / gv.z, padding: [cs.paddingTop, cs.paddingRight, cs.paddingBottom, cs.paddingLeft], align: cs.justifyContent };
          });
          const headers = [document.querySelector('.corner'), [...document.querySelectorAll('.hc')].find(el => el.textContent === 'C'), [...document.querySelectorAll('.hr')].find(el => el.textContent === '3')].map(el => el.getBoundingClientRect().toJSON());
          return { headers, zoom: gv.z * 100, origin: rect(0, 0), thin: rect(2, 1), double: rect(2, 3), text, panes: gv.panes.filter(p => p.win).length, pathWidths: [...document.querySelectorAll('.cell-borders path')].slice(0, 15).map(el => Number(el.getAttribute('stroke-width'))) };
        });
        const bytes = await page.screenshot(), image = png(bytes), label = `DPR${dpr} ${freeze.name} step${step} zoom${zoom}`;
        if (output) { const name = `zoom-dpr${dpr}-${freeze.name}-${step}-${zoom}`; await writeFile(join(output, name + '.png'), bytes); await writeFile(join(output, name + '.json'), JSON.stringify(info)); }
        eq(Math.abs(info.zoom - zoom) < 1e-9, true, label + ' slider applied');
        for (const side of ['bottom', 'right']) {
          eq(countStroke(image, info.thin, side, dpr).length, 1, label + ' thin ' + side + ' = 1 physical pixel');
          const pixels = countStroke(image, info.double, side, dpr);
          eq(pixels.length, 2, label + ' double two one-pixel strokes ' + side);
          eq(pixels[1] - pixels[0], 2, label + ' double one-pixel gap ' + side);
        }
        for (const side of ['top', 'left']) eq(countStroke(image, info.origin, side, dpr).length, 1, label + ' origin/header ' + side + ' = 1 physical pixel');
        if (zoom === 400) for (const [i, rect] of info.headers.entries()) for (const side of ['bottom', 'right']) {
          const vertical = side === 'right', x = (vertical ? rect.right : rect.left + rect.width / 4) * dpr, y = (vertical ? rect.top + rect.height / 4 : rect.bottom) * dpr;
          let gray = 0; for (let k = -5; k <= 5; k++) { const p = image.at(x + (vertical ? k : 0), y + (vertical ? 0 : k)); if (p.length === 3 && p.every(v => v === 208)) gray++; }
          eq(gray, 1, label + ' gray header ' + i + ' ' + side + ' one physical pixel');
        }
        if (!baseline) baseline = info.text;
        for (let c = 0; c < 3; c++) {
          const text = info.text[c], base = baseline[c];
          for (const k of ['dx', 'dy', 'right', 'width', 'height']) eq(Math.abs(text[k] - base[k]) < .05, true, label + ' text ' + c + ' invariant CSS ' + k);
          eq(text.padding, base.padding, label + ' text ' + c + ' padding unchanged');
          eq(text.padding, ['1px', '3px', '1px', c === 0 ? '2px' : '3px'], label + ' text ' + c + ' Excel-aligned horizontal/middle insets');
          eq(text.dx >= 0 && text.right >= 0, true, label + ' text ' + c + ' remains inside cell');
        }
        runs.push({ dpr, freeze: freeze.name, step, zoom, elapsedMs: Math.round((performance.now() - started) * 10) / 10 });
      }
    }
    await context.close();
  }
  eq(errors, [], 'page errors'); eq(writes, [], 'remote writes');
  console.log(JSON.stringify({ ok: failed === 0, cases: runs.length, checks, failed, failures, pageErrors: errors, blockedWrites: writes, runs }));
  assert.equal(failed, 0, 'zoom regression failed; see JSON above');
} finally { await browser.close(); }
