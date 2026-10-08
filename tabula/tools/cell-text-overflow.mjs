import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { createHash } from 'node:crypto';
import { inflateSync } from 'node:zlib';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

// Synthetic display test. Read actual PNG pixels, not the production overflow helper.
const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const output = process.env.CELL_TEXT_OVERFLOW_OUTPUT;
if (output) await mkdir(output, { recursive: true });
const target = process.env.WIXEL_URL || 'http://127.0.0.1:5193/', targetUrl = new URL(target);
const local = ['localhost', '127.0.0.1'].includes(targetUrl.hostname);
assert.ok(local || (target === process.env.WIXEL_ALLOWED_TEST_URL && targetUrl.origin === 'https://wixel-3.wizx.workers.dev'), 'local or explicitly approved synthetic deployment only');
const baselinePath = process.env.TEXT_OVERFLOW_BASELINE_VIEW;
const baselineBody = baselinePath ? await readFile(baselinePath, 'utf8') : null;
const fingerprint = async () => Object.fromEntries(await Promise.all(['src/view.js', 'styles.css', 'src/axis.js'].map(async p => [p, createHash('sha256').update(await readFile(join(repo, p))).digest('hex')])));
const before = await fingerprint();
const playwright = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browserName = process.env.TEXT_OVERFLOW_BROWSER || 'chromium';
assert.ok(['chromium', 'webkit', 'firefox'].includes(browserName), 'official Playwright browser only');

function png(bytes) {
  let width, height, type; const parts = [];
  for (let p = 8; p < bytes.length;) {
    const n = bytes.readUInt32BE(p), name = bytes.toString('ascii', p + 4, p + 8), d = bytes.subarray(p + 8, p + 8 + n);
    if (name === 'IHDR') { width = d.readUInt32BE(0); height = d.readUInt32BE(4); type = d[9]; assert.equal(d[8], 8); }
    if (name === 'IDAT') parts.push(d); p += n + 12;
  }
  assert.ok(type === 2 || type === 6);
  const bpp = type === 6 ? 4 : 3, stride = width * bpp;
  const raw = inflateSync(Buffer.concat(parts)), pixels = Buffer.alloc(width * height * bpp);
  const paeth = (a, b, c) => { const p = a + b - c, pa = Math.abs(p - a), pb = Math.abs(p - b), pc = Math.abs(p - c); return pa <= pb && pa <= pc ? a : pb <= pc ? b : c; };
  let k = 0;
  for (let y = 0; y < height; y++) {
    const f = raw[k++]; assert.ok(f <= 4);
    for (let x = 0; x < stride; x++) {
      const i = y * stride + x, a = x >= bpp ? pixels[i - bpp] : 0, b = y ? pixels[i - stride] : 0, c = y && x >= bpp ? pixels[i - stride - bpp] : 0;
      pixels[i] = (raw[k++] + (f === 0 ? 0 : f === 1 ? a : f === 2 ? b : f === 3 ? Math.floor((a + b) / 2) : paeth(a, b, c))) & 255;
    }
  }
  return { width, height, rgb(x, y) { const i = (Math.floor(y) * width + Math.floor(x)) * bpp; return [pixels[i], pixels[i + 1], pixels[i + 2]]; } };
}
function redPixels(im, rect, dpr) {
  if (!rect) return 0;
  let count = 0;
  const left = Math.max(0, Math.ceil((rect.left + 1) * dpr)), right = Math.min(im.width, Math.floor((rect.right - 1) * dpr));
  const top = Math.max(0, Math.ceil((rect.top + 1) * dpr)), bottom = Math.min(im.height, Math.floor((rect.bottom - 1) * dpr));
  for (let y = top; y < bottom; y++) for (let x = left; x < right; x++) {
    const [r, g, b] = im.rgb(x, y); if (r > 130 && g < 100 && b < 100) count++;
  }
  return count;
}
const fixtures = [
  { id: 'left-hidden-C-filled-D', source: 1, barrier: 3, hidden: [2], hiDpr: true },
  { id: 'left-visible-gap-filled-D', source: 1, barrier: 3, probe: 2, hiDpr: true },
  { id: 'left-hidden-valued-C-filled-D', source: 1, barrier: 3, hidden: [2], extras: [[2, 'HIDDEN']] },
  { id: 'left-multiple-hidden-gap', source: 1, barrier: 6, hidden: [2, 3, 5] },
  { id: 'left-immediate-filled', source: 1, barrier: 2 },
  { id: 'left-empty-formula', source: 1, barrier: 2, barrierRaw: '=""' },
  { id: 'left-zero', source: 1, barrier: 2, barrierRaw: '0', value: 0 },
  { id: 'left-false', source: 1, barrier: 2, barrierRaw: '=FALSE()', value: false },
  { id: 'left-merge-after-gap', source: 1, barrier: 3, merges: [{ r1: 3, c1: 3, r2: 3, c2: 4 }] },
  { id: 'left-hidden-gap-merge', source: 1, barrier: 3, hidden: [2], merges: [{ r1: 3, c1: 3, r2: 3, c2: 4 }] },
  { id: 'right-hidden-gap-filled-B', source: 3, barrier: 1, hidden: [2], align: 'right', hiDpr: true },
  { id: 'right-visible-gap-filled-B', source: 3, barrier: 1, align: 'right', probe: 2 },
  { id: 'right-immediate-filled', source: 3, barrier: 2, align: 'right' },
  { id: 'center-two-gaps-filled', source: 3, barrier: 5, align: 'center', extras: [[1, 'LEFT TARGET']], secondBarrier: 1, probe: 4 },
  { id: 'center-left-occupied-right-empty', source: 3, barrier: 2, align: 'center', probe: 4, hiDpr: true },
  { id: 'center-right-occupied-left-empty', source: 3, barrier: 4, align: 'center', probe: 2 },
  { id: 'left-wrap', source: 1, barrier: 3, hidden: [2], style: { wrap: true } },
  { id: 'left-shrink', source: 1, barrier: 3, hidden: [2], style: { shrink: true } },
  { id: 'hidden-occupied-does-not-block', source: 1, barrier: 4, hidden: [2], extras: [[2, 'HIDDEN']], probe: 3 },
  { id: 'fill-only-does-not-block', source: 1, barrier: 3, probe: 2, probeStyle: { fill: '#e2f0d9' } },
  { id: 'merged-source-stays-clipped', source: 1, barrier: 3, merges: [{ r1: 3, c1: 1, r2: 3, c2: 2 }] },
  { id: 'zero-spill-child-after-gap', source: 1, barrier: 3, spill: [2, 3, '={1;0}'], value: 0 },
  { id: 'false-spill-child-after-gap', source: 1, barrier: 3, spill: [1, 3, '={1;2;FALSE}'], value: false },
  { id: 'freeze-hidden-gap', source: 1, barrier: 3, hidden: [2], freeze: { rows: 2, cols: 1 }, hiDpr: true },
  { id: 'scrolled-hidden-occupied-source', source: 1, barrier: 4, hidden: [2], extras: [[2, 'HIDDEN']], probe: 3, scroll: 150 },
];
let checks = 0; const failures = [], errors = [], blocked = [], runs = [];
function eq(actual, expected, message) {
  checks++;
  try { assert.deepEqual(actual, expected); }
  catch { failures.push({ message, actual, expected }); }
}
const browser = await playwright[browserName].launch(); let exception;
try {
  for (const dpr of [1, 1.5]) {
    const context = await browser.newContext({ viewport: { width: 1280, height: 850 }, deviceScaleFactor: dpr, serviceWorkers: 'block' });
    await context.route('**/*', route => {
      const request = route.request(), u = new URL(request.url());
      if (u.origin !== targetUrl.origin || request.method() !== 'GET' || u.pathname.startsWith('/api/')) {
        blocked.push({ url: request.url(), method: request.method() }); return route.abort('blockedbyclient');
      }
      if (baselineBody && u.pathname === '/src/view.js') return route.fulfill({ contentType: 'text/javascript', body: baselineBody });
      return route.continue();
    });
    const page = await context.newPage(); page.on('pageerror', e => errors.push(e.message));
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => !!window.tabula?.gv());
    for (const fixture of fixtures.filter(f => dpr === 1 || f.hiDpr)) for (const zoom of [75, 100, 150]) {
      const info = await page.evaluate(({ fixture, zoom }) => {
        const tab = window.tabula, wb = tab.wb(), gv = tab.gv(), r = 3;
        wb.load({ sheets: [{ name: '합성 텍스트 경계', cells: {}, defColW: 70, defRowH: 34 }] });
        wb.transact(() => {
          wb.setSheetProp(0, 'hiddenCols', Object.fromEntries((fixture.hidden || []).map(c => [c, true])));
          if (fixture.merges) wb.setSheetProp(0, 'merges', fixture.merges);
          if (fixture.freeze) wb.setSheetProp(0, 'freeze', fixture.freeze);
          for (let c = 0; c <= 26; c++) wb.setStyle(0, r, c, { fill: '#ffffff', size: 15, color: '#0033cc', bb: true, bbc: '#202020' });
          wb.setInput(0, r, fixture.source, 'MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM');
          wb.setStyle(0, r, fixture.source, { fill: '#ffffff', size: 15, color: '#d80000', align: fixture.align || 'left', ...fixture.style });
          if (fixture.spill) wb.setInput(0, ...fixture.spill);
          else wb.setInput(0, r, fixture.barrier, fixture.barrierRaw ?? 'TARGET');
          for (const [c, raw] of fixture.extras || []) wb.setInput(0, r, c, raw);
          if (fixture.probeStyle) wb.setStyle(0, r, fixture.probe, { fill: '#ffffff', size: 15, ...fixture.probeStyle });
        });
        if (fixture.spill) wb.getValue(0, fixture.spill[0], fixture.spill[1]);
        const version = wb.version, sourceRaw = wb.getCell(0, r, fixture.source).raw;
        tab.selectCell(0, 0); gv.setZoom(zoom); gv.setScroll(fixture.scroll || 0, 0); gv.renderAll();
        document.querySelectorAll('.overlay').forEach(e => { e.style.visibility = 'hidden'; });
        const find = c => [...document.querySelectorAll('.c[data-r="' + r + '"][data-c="' + c + '"]')].find(e => e.closest('.pane').getBoundingClientRect().width > 0);
        const rect = c => find(c)?.getBoundingClientRect().toJSON() ?? null, source = find(fixture.source);
        return { source: rect(fixture.source), barrier: rect(fixture.barrier), secondBarrier: fixture.secondBarrier == null ? null : rect(fixture.secondBarrier), probe: fixture.probe == null ? null : rect(fixture.probe), className: source?.className ?? null, clipPath: source ? getComputedStyle(source).clipPath : null, sourceRaw, versionBefore: version, versionAfter: wb.version, sourceRawAfter: wb.getCell(0, r, fixture.source).raw, hiddenAbsent: (fixture.hidden || []).every(c => !document.querySelector('.c[data-r="' + r + '"][data-c="' + c + '"]')), barrierRaw: wb.getCell(0, r, fixture.barrier)?.raw ?? null, barrierValue: wb.getValue(0, r, fixture.barrier), scroll: gv.sx };
      }, { fixture, zoom });
      await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
      const bytes = await page.screenshot(), im = png(bytes), label = fixture.id + ' DPR' + dpr + ' zoom' + zoom;
      const redBarrier = redPixels(im, info.barrier, dpr), redSecond = redPixels(im, info.secondBarrier, dpr), redProbe = redPixels(im, info.probe, dpr);
      eq(info.barrier != null, true, label + ' real barrier DOM exists');
      eq(info.hiddenAbsent, true, label + ' hidden columns absent');
      eq(info.sourceRawAfter, info.sourceRaw, label + ' render preserves source value');
      eq(info.versionAfter, info.versionBefore, label + ' render does not mutate workbook');
      eq(redBarrier, 0, label + ' no red source pixels inside occupied barrier');
      if (fixture.secondBarrier != null) eq(redSecond, 0, label + ' opposite occupied barrier also clips');
      if (fixture.probe != null) eq(redProbe > 0, true, label + ' allowed empty visible cell really receives source pixels');
      if (Object.hasOwn(fixture, 'value')) eq(info.barrierValue, fixture.value, label + ' zero/FALSE/spill value is real');
      runs.push({ id: fixture.id, dpr, zoom, redBarrier, redSecond, redProbe, info });
      if (output && dpr === 1 && zoom === 100 && ['left-hidden-C-filled-D', 'left-visible-gap-filled-D', 'center-left-occupied-right-empty', 'scrolled-hidden-occupied-source'].includes(fixture.id)) await writeFile(join(output, fixture.id + '.png'), bytes);
    }
    await context.close();
  }
} catch (e) { exception = { name: e.name, message: e.message, stack: e.stack }; }
finally { await browser.close(); }
const after = await fingerprint(), sourceUnchanged = JSON.stringify(before) === JSON.stringify(after);
eq(errors, [], 'no page errors'); eq(blocked.filter(x => x.method !== 'GET'), [], 'no API/external writes');
if (!baselineBody) eq(sourceUnchanged, true, 'production source remained frozen during test');
const result = { browser: browserName, cases: runs.length, checks, bad: failures.length, failures, errors, blocked, exception, before, after, sourceUnchanged, baseline: baselineBody ? { path: baselinePath, sha256: createHash('sha256').update(baselineBody).digest('hex') } : null, runs, browserClosed: true, customerFiles: 0, note: 'Excel oracle is a separate synthetic PDF reference; these PNG assertions validate WIXEL clipping, not universal Excel screen pixel parity.' };
if (output) await writeFile(join(output, 'result.json'), JSON.stringify(result, null, 2));
console.log(JSON.stringify({ browser: browserName, cases: result.cases, checks, bad: failures.length, errors, exception, browserClosed: true, failures: failures.slice(0, 12) }, null, 2));
if (failures.length || exception) process.exitCode = 1;
