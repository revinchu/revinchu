import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { readFileSync } from 'node:fs';
import { join, resolve, extname, sep } from 'node:path';

// 화면 중간에서 찾은 가로/세로 검정 띠의 교차 사각형이 실제 PNG에서도
// 빈 픽셀 없이 이어지는지 검사한다. 제품의 선 끝점/폭 계산은 사용하지 않는다.
const playwright = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const target = process.env.WIXEL_URL || 'http://localhost:5178/';
const output = process.env.CORNER_SCREENSHOTS;
if (output) await mkdir(output, { recursive: true });
const engines = (process.env.CORNER_ENGINES || 'chromium').split(',');
let cases = 0, checks = 0, failureCount = 0;
const failures = [], errors = [], writes = [];
const eq = (actual, expected, label) => {
  checks++;
  try { assert.deepEqual(actual, expected); }
  catch { failureCount++; if (failures.length < 40) failures.push({ label, actual, expected }); }
};
for (const engine of engines) {
  const browser = await playwright[engine].launch();
  try {
    for (const dpr of [1, 1.25, 1.5, 2]) {
      const context = await browser.newContext({ viewport: { width: 1500, height: 1150 }, deviceScaleFactor: dpr });
      const page = await context.newPage();
      page.on('pageerror', e => errors.push(e.message));
      await context.route('**/*', route => {
        if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort(); }
        const u = new URL(route.request().url());
        if (u.origin !== new URL(target).origin || u.pathname.startsWith('/api/')) return route.abort();
        if (process.env.WIXEL_LOCAL_ASSETS) {
          const root = resolve(process.env.WIXEL_LOCAL_ASSETS), asset = decodeURIComponent(u.pathname).replace(/^\//, '') || 'index.html';
          if (!/^(index\.html|wixel-[a-f0-9]{16}\.js|styles\.css|src\/.+\.js|assets\/.+|icons\/.+|manifest\.webmanifest|favicon[^/]*|apple-touch-icon\.png)$/.test(asset)) return route.abort();
          const file = resolve(root, asset);
          if (!file.startsWith(root + sep)) return route.abort();
          try { return route.fulfill({ body: readFileSync(file), contentType: ({ '.html': 'text/html', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png' })[extname(file)] || 'application/octet-stream' }); }
          catch { return route.fulfill({ status: 404, body: '' }); }
        }
        return route.continue();
      });
      await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
      await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 120000 });
      await page.waitForFunction(() => !!window.tabula?.gv(), null, { timeout: 120000 });
      await page.evaluate(() => {
        const wb = window.tabula.wb(); wb.load({ sheets: [{ name: '외곽선 검증', cells: {} }] });
        window.cornerTables = [
          { r: 2, c: 2, h: 'thin', v: 'thin' }, { r: 2, c: 7, h: 'medium', v: 'medium' },
          { r: 8, c: 2, h: 'thick', v: 'thick' }, { r: 8, c: 7, h: 'medium', v: 'thick' },
        ];
        wb.transact(() => {
          wb.setSheetProp(0, 'defColW', 60); wb.setSheetProp(0, 'defRowH', 23);
          wb.setSheetProp(0, 'noGrid', true);
          const merges = [];
          for (const t of window.cornerTables) {
            for (let r = t.r; r <= t.r + 3; r++) for (let c = t.c; c <= t.c + 3; c++) {
              wb.setStyle(0, r, c, {
                fill: r === t.r ? '#002060' : '#ffffff', color: r === t.r ? '#ffffff' : '#000000',
                bt: r === t.r, bts: t.h, bb: true, bbs: r === t.r + 3 ? t.h : 'dotted',
                bl: c === t.c, bls: t.v, br: true, brs: c === t.c + 3 ? t.v : 'dotted',
              });
            }
            wb.setInput(0, t.r, t.c, '키워드'); wb.setInput(0, t.r + 1, t.c, '테두리');
            merges.push({ r1: t.r, r2: t.r, c1: t.c, c2: t.c + 2 });
          }
          wb.setSheetProp(0, 'merges', merges);
        });
        document.querySelectorAll('.overlay').forEach(e => e.style.visibility = 'hidden');
      });
      for (const state of ['plain', 'hidden-freeze-scroll']) {
        if (state !== 'plain') await page.evaluate(() => {
          const wb = window.tabula.wb();
          wb.transact(() => {
            wb.setSheetProp(0, 'hiddenRows', { 4: true }); wb.setSheetProp(0, 'hiddenCols', { 3: true });
            wb.setSheetProp(0, 'freeze', { rows: 1, cols: 1 });
          });
        });
        for (const zoom of [50, 75, 80, 100, 125, 200]) {
          const rectangles = await page.evaluate(({ zoom, state }) => {
            const gv = window.tabula.gv(); gv.setZoom(zoom); gv.setScroll(state === 'plain' ? 0 : 7.25, state === 'plain' ? 0 : 3.75); gv.renderAll();
            const rect = (r, c) => document.querySelector(`.pane-br .c[data-r="${r}"][data-c="${c}"]`).getBoundingClientRect();
            return window.cornerTables.map(t => {
              const a = rect(t.r, t.c), b = rect(t.r + 3, t.c + 3);
              return { left: a.left, top: a.top, right: b.right, bottom: b.bottom, h: t.h, v: t.v };
            });
          }, { zoom, state });
          await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
          const bytes = await page.screenshot(), label = `${engine}-${state}-dpr${dpr}-zoom${zoom}`;
          if (output && dpr === 1 && zoom === 100) await writeFile(join(output, label + '.png'), bytes);
          // PNG 재디코딩만 브라우저 Canvas로 수행. 좌표 기대값은 DOM 표 범위의
          // 중간에서 찾은 실제 띠이며 SVG path 속성을 읽지 않는다.
          const measured = await page.evaluate(async ({ base64, rectangles, dpr }) => {
            const img = new Image(); img.src = 'data:image/png;base64,' + base64; await img.decode();
            const canvas = document.createElement('canvas'); canvas.width = img.width; canvas.height = img.height;
            const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0);
            const { data } = ctx.getImageData(0, 0, img.width, img.height);
            const black = (x, y) => { const i = (y * img.width + x) * 4; return x >= 0 && y >= 0 && x < img.width && y < img.height && data[i] < 24 && data[i + 1] < 24 && data[i + 2] < 24; };
            return rectangles.map(t => {
              const xmid = Math.round((t.left + (t.right - t.left) * .37) * dpr), ymid = Math.round((t.top + (t.bottom - t.top) * .61) * dpr);
              const bands = {};
              for (const side of ['left', 'right', 'top', 'bottom']) {
                const vertical = side === 'left' || side === 'right', center = Math.round(t[side] * dpr);
                bands[side] = [];
                for (let p = center - 15; p <= center + 15; p++) if (black(vertical ? p : xmid, vertical ? ymid : p)) bands[side].push(p);
              }
              const corners = [];
              for (const h of ['top', 'bottom']) for (const v of ['left', 'right']) {
                let gaps = 0;
                for (const x of bands[v]) for (const y of bands[h]) if (!black(x, y)) gaps++;
                corners.push({ name: h + '-' + v, gaps });
              }
              return { widths: Object.fromEntries(Object.entries(bands).map(([k, v]) => [k, v.length])), corners };
            });
          }, { base64: bytes.toString('base64'), rectangles, dpr });
          for (let i = 0; i < measured.length; i++) {
            for (const [side, width] of Object.entries(measured[i].widths)) eq(width > 0, true, `${label} table${i} ${side} exists`);
            for (const c of measured[i].corners) eq(c.gaps, 0, `${label} table${i} ${c.name} has no gap`);
          }
          cases++;
        }
      }
      await context.close();
    }
  } finally { await browser.close(); }
}
eq(errors, [], 'page errors'); eq(writes, [], 'server writes');
console.log(JSON.stringify({ ok: !failureCount, engines, cases, checks, failureCount, failures, errors, writes }));
assert.equal(failureCount, 0, '외곽 모서리 픽셀 검사 실패');
