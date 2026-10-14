// 합성 보고서만 사용해 배율 전환 뒤 SVG 글자의 실제 PNG 픽셀까지 확인한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5184/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw new Error('로컬 서버에서 실행하세요.');
const out = process.env.REPORT_CHART_OUT || 'D:/Codex/Temp/wixel-report-chart';
await mkdir(out, { recursive: true });
const engines = (process.env.REPORT_CHART_ENGINES || 'chromium,firefox,webkit').split(',');
let cases = 0, checks = 0;
const errors = [], writes = [];
for (const engine of engines) {
  const browser = await pw[engine].launch();
  try {
    for (const dpr of [1, 1.25, 2]) {
      const context = await browser.newContext({ viewport: { width: 1700, height: 1100 }, deviceScaleFactor: dpr });
      await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
      await context.route('**/*', r => {
        const u = new URL(r.request().url());
        if (!['GET', 'HEAD', 'OPTIONS'].includes(r.request().method())) { writes.push(r.request().method()); return r.abort(); }
        return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? r.continue() : r.abort();
      });
      const p = await context.newPage(); p.on('pageerror', e => errors.push(e.message));
      await p.goto(url); await p.waitForFunction(() => window.tabula?.gv());
      await p.evaluate(() => {
        const categories = Array.from({ length: 31 }, (_, i) => '22-05-' + String(i + 1).padStart(2, '0'));
        const chart = { id: 'schedule-test', type: 'combo', x: 30, y: 40, w: 1150, h: 290, gap: 219, overlap: -27, legend: 'b', gridY: false,
          axes: { x: { type: 'date' }, y: { min: 0, max: 600000, major: 100000, numFmt: '#,##0' }, y2: { min: 0, max: .03, major: .005, numFmt: '0.00%' } },
          snapshotData: { categories, series: [
            { name: '노출', type: 'column', color: '#4f81bd', values: categories.map((_, i) => i < 26 ? 40000 + i * 17000 : 0), pointLabels: { 0: { labels: true } }, numFmt: '#,##0' },
            { name: '클릭', type: 'column', color: '#c0504d', values: categories.map(() => 800) },
            { name: 'CTR', type: 'line', axis: 1, color: '#9bbb59', marker: 'none', values: categories.map((_, i) => i < 26 ? .015 + i % 7 * .001 : 0) },
          ] } };
        const w = tabula.wb(); w.restore({ sheets: [
          { name: 'Summary', zoom: 85, cells: {}, charts: [{ ...chart, id: 'summary-test' }] },
          { name: 'Schedule', zoom: 80, cells: {}, freeze: { rows: 19, cols: 8 }, charts: [chart] },
        ] });
        tabula.switchSheet(1); tabula.switchSheet(0);
      });
      await p.evaluate(() => document.fonts.ready);
      // Rasterize the old sheet first: DOM bounds alone miss Chromium's stale text paint scale.
      await p.screenshot();
      await p.evaluate(() => tabula.switchSheet(1));
      for (const zoom of [80, 100, 125, 85, 80]) {
        await p.evaluate(z => { tabula.gv().setZoom(z); tabula.gv().setScroll(0, 0); }, zoom);
        await p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
        const rects = await p.locator('.chart[data-id="schedule-test"] svg').first().evaluate(svg => {
          const s = svg.getBoundingClientRect();
          return [...svg.querySelectorAll('[data-el="axis-y2"] text,[data-el="legend"] text,[data-el="axis-x"] text,[data-el="label"]')].map(e => {
            const r = e.getBoundingClientRect();
            return { text: e.textContent, x: r.x, y: r.y, w: r.width, h: r.height, inside: r.x >= s.x - 1 && r.y >= s.y - 1 && r.right <= s.right + 1 && r.bottom <= s.bottom + 1, lines: e.querySelectorAll('tspan').length };
          });
        });
        assert.equal(rects.length, 42, '31 dates + 7 secondary ticks + 3 legend names + 1 point label'); checks++;
        for (const r of rects) { assert.ok(r.inside, r.text + ' is inside chart'); assert.equal(r.lines, 0, 'date does not wrap'); checks += 2; }
        const png = await p.screenshot(), name = `${engine}-dpr${dpr}-zoom${zoom}-${cases}`;
        const counts = await p.evaluate(async ({ base64, rects, dpr }) => {
          const img = new Image(); img.src = 'data:image/png;base64,' + base64; await img.decode();
          const canvas = document.createElement('canvas'); canvas.width = img.width; canvas.height = img.height;
          const ctx = canvas.getContext('2d'); ctx.drawImage(img, 0, 0);
          const data = ctx.getImageData(0, 0, img.width, img.height).data;
          return rects.map(r => {
            let ink = 0;
            for (let y = Math.max(0, Math.floor(r.y * dpr)); y < Math.min(img.height, Math.ceil((r.y + r.h) * dpr)); y++) for (let x = Math.max(0, Math.floor(r.x * dpr)); x < Math.min(img.width, Math.ceil((r.x + r.w) * dpr)); x++) {
              const k = (y * img.width + x) * 4;
              if (Math.max(data[k], data[k + 1], data[k + 2]) < 200) ink++;
            }
            return ink;
          });
        }, { base64: png.toString('base64'), rects, dpr });
        if (dpr === 1 || counts.some(n => n < 4)) await writeFile(`${out}/${name}.png`, png);
        counts.forEach((n, i) => { assert.ok(n >= 4, `${name}: '${rects[i].text}' must have painted pixels in its own bounds (${n})`); checks++; });
        cases++;
      }
      await context.close();
    }
  } finally { await browser.close(); }
}
assert.deepEqual(errors, []); assert.deepEqual(writes, []); checks += 2;
console.log(JSON.stringify({ ok: true, engines, cases, checks, errors, writes }));
