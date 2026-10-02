// 합성 문서만 사용한다. 외부/서버 API 요청과 실제 파일 저장은 차단한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
import { inflateSync } from 'node:zlib';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(base).hostname)) throw new Error('로컬 합성 검사 전용입니다.');
const output = process.env.PRINT_PAGINATION_OUTPUT || 'D:/Codex/Temp/wixel-print-pagination';
await mkdir(output, { recursive: true });
function pngPixels(bytes) {
  let width, height, channels; const parts = [];
  for (let p = 8; p < bytes.length;) { const n = bytes.readUInt32BE(p), type = bytes.toString('ascii', p + 4, p + 8), data = bytes.subarray(p + 8, p + 8 + n); if (type === 'IHDR') { width = data.readUInt32BE(0); height = data.readUInt32BE(4); channels = data[9] === 6 ? 4 : 3; assert.ok(data[8] === 8 && [2, 6].includes(data[9])); } if (type === 'IDAT') parts.push(data); p += n + 12; }
  const raw = inflateSync(Buffer.concat(parts)), stride = width * channels, pixels = Buffer.alloc(stride * height); let at = 0;
  const paeth = (a, b, c) => { const p = a + b - c, da = Math.abs(p - a), db = Math.abs(p - b), dc = Math.abs(p - c); return da <= db && da <= dc ? a : db <= dc ? b : c; };
  for (let y = 0; y < height; y++) { const filter = raw[at++]; assert.ok(filter <= 4); for (let x = 0; x < stride; x++) { const i = y * stride + x, a = x >= channels ? pixels[i - channels] : 0, b = y ? pixels[i - stride] : 0, c = y && x >= channels ? pixels[i - stride - channels] : 0; pixels[i] = (raw[at++] + (filter === 1 ? a : filter === 2 ? b : filter === 3 ? Math.floor((a + b) / 2) : filter === 4 ? paeth(a, b, c) : 0)) & 255; } }
  return { ink(rect) { let count = 0; for (let y = Math.ceil(rect.y); y < Math.floor(rect.y + rect.h); y++) for (let x = Math.ceil(rect.x + 2); x < Math.floor(rect.x + rect.w - 2); x++) { const i = (y * width + x) * channels; if (pixels[i] < 50 && pixels[i + 1] < 50 && pixels[i + 2] < 50) count++; } return count; } };
}
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1440, height: 1050 } }), page = await context.newPage();
page.setDefaultTimeout(20000); const errors = [], writes = [], results = [];
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method() + ' ' + u.pathname); return route.abort(); } return u.origin === new URL(base).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort(); });
await context.addInitScript(() => {
  window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; window.__printed = 0; window.__saved = null;
  window.print = () => { window.__printed++; };
  window.showSaveFilePicker = async options => ({ name: options.suggestedName, createWritable: async () => ({
    write: async blob => { window.__saved = { name: options.suggestedName, type: blob.type, size: blob.size, text: await blob.text() }; }, close: async () => {}, abort: async () => {},
  }) });
});
const pg = { paper: 1, margins: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 }, scale: 100 };
const dialog = () => page.getByRole('dialog', { name: '인쇄 미리보기', exact: true });
async function close() { if (await dialog().count()) await dialog().getByRole('button', { name: '닫기', exact: true }).last().click(); }
async function reset(settings, extra = {}) {
  await close();
  await page.evaluate(({ settings, extra }) => {
    const t = tabula, w = t.wb(); if (t.si !== 0) t.switchSheet(0);
    const cells = {}; for (let r = 0; r < 30; r++) for (let c = 0; c < 20; c++) cells[r + ',' + c] = { raw: `R${r}C${c}` };
    w.restore({ sheets: [{ name: '초기화', cells: {} }, { name: '합성 출력', cells, defRowH: 100, defColW: 100, page: settings, view: { mode: 'pageBreakPreview' }, ...extra }] });
    t.switchSheet(1); w.undoStack = []; w.redoStack = []; window.__beforePrint = JSON.stringify(w.serialize()); window.__saved = null;
    t.run('print');
  }, { settings, extra });
  await dialog().waitFor();
}
async function pages() {
  await dialog().getByRole('button', { name: '인쇄', exact: true }).click();
  return page.locator('#printArea > .wixel-print-page').evaluateAll(nodes => nodes.map(node => ({
    columns: [...node.querySelectorAll('col')].map(c => Number(c.dataset.printCol)),
    head: [...node.querySelectorAll('thead tr[data-print-row]')].map(r => Number(r.dataset.printRow)),
    rows: [...node.querySelectorAll('tbody tr[data-print-row]')].map(r => Number(r.dataset.printRow)),
    text: node.textContent, objects: [...node.querySelectorAll('[data-print-object]')].map(n => n.dataset.printObject),
    merges: [...node.querySelectorAll('td[colspan],td[rowspan]')].filter(n => n.colSpan > 1 || n.rowSpan > 1).map(n => ({ text: n.textContent, cols: n.colSpan, rows: n.rowSpan })),
  })));
}
async function test(name, fn) { try { await fn(); results.push({ name, ok: true }); console.log('PASS ' + name); } catch (e) { results.push({ name, ok: false, error: e.message }); console.log('FAIL ' + name + '\n' + e.stack); } }
try {
  await page.goto(base, { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => window.tabula?.wb());
  await test('복수 영역은 별도 쪽 · 수동 나누기 · 전역 쪽 번호', async () => {
    await reset({ ...pg, areas: [{ r1: 0, c1: 0, r2: 3, c2: 2 }, { r1: 6, c1: 5, r2: 8, c2: 7 }], rowBreaks: [2], header: '&P / &N' });
    const all = await pages(); assert.equal(all.length, 3);
    assert.deepEqual(all.map(p => p.rows), [[0, 1], [2, 3], [6, 7, 8]]);
    assert.deepEqual(all.map(p => p.columns), [[0, 1, 2], [0, 1, 2], [5, 6, 7]]);
    all.forEach((p, i) => assert.ok(p.text.startsWith(`${i + 1} / 3`)));
    assert.equal(await page.evaluate(() => JSON.stringify(tabula.wb().serialize()) === __beforePrint), true);
  });
  await test('HTML은 미리보기와 동일한 영역·본문·쪽 번호', async () => {
    const expected = await pages(); await close(); await page.evaluate(() => { __saved = null; tabula.run('exportHtml'); });
    await page.waitForFunction(() => window.__saved?.type.startsWith('text/html'));
    const actual = await page.evaluate(() => { const doc = new DOMParser().parseFromString(__saved.text, 'text/html'); return [...doc.querySelectorAll('.wixel-print-page')].map(n => ({ columns: [...n.querySelectorAll('col')].map(c => Number(c.dataset.printCol)), rows: [...n.querySelectorAll('tbody tr')].map(r => Number(r.dataset.printRow)), text: n.textContent })); });
    assert.deepEqual(actual, expected.map(({ columns, rows, text }) => ({ columns, rows, text })));
  });
  await test('맞춤은 수동 나누기를 무시 · 저장 설정과 화면 계획 일치', async () => {
    const area = { r1: 0, c1: 0, r2: 7, c2: 7 }, settings = { ...pg, area, rowBreaks: [3], colBreaks: [3] };
    await reset(settings); assert.equal((await pages()).length, 4);
    await reset({ ...settings, fitW: 1, fitH: 1 }); assert.equal((await pages()).length, 1);
    assert.equal(await page.evaluate(() => tabula.gv().printLayout().manualIgnored), true);
    assert.deepEqual(await page.evaluate(() => tabula.wb().sheets[tabula.si].page.rowBreaks), [3]);
  });
  await test('영역 밖 반복 행·열 · 숨김 · 병합 · 개체를 같은 쪽에 보존', async () => {
    await reset({ ...pg, area: { r1: 3, c1: 3, r2: 8, c2: 8 }, titleRows: [0, 0], titleCols: [0, 0], rowBreaks: [6], colBreaks: [6] }, {
      hiddenRows: { 4: true }, hiddenCols: { 4: true }, merges: [{ r1: 3, c1: 3, r2: 3, c2: 5 }],
      shapes: [{ id: 'body-shape', kind: 'rect', x: 320, y: 320, w: 80, h: 80, fill: '#ff0000' }, { id: 'never-print', kind: 'rect', x: 0, y: 0, w: 40, h: 40, noPrint: true }],
    });
    const all = await pages(); assert.equal(all.length, 4);
    assert.deepEqual(all.map(p => p.rows), [[3, 5], [6, 7, 8], [3, 5], [6, 7, 8]]);
    assert.deepEqual(all.map(p => p.columns), [[0, 3, 5], [0, 3, 5], [0, 6, 7, 8], [0, 6, 7, 8]]);
    assert.ok(all.every(p => p.head.join(',') === '0')); assert.ok(all.every(p => !p.objects.includes('never-print')));
    assert.deepEqual(all[0].merges, [{ text: 'R3C3', cols: 2, rows: 1 }]); assert.ok(all[0].objects.includes('body-shape'));
    assert.equal(all.filter(p => p.objects.includes('body-shape')).length, 1);
    await page.screenshot({ path: join(output, 'repeated-hidden-merged.png') });
  });
  await test('PDF는 동일 계획의 실제 4페이지 · 저장 원본 불변', async () => {
    await dialog().getByRole('button', { name: 'PDF 파일 저장', exact: true }).click();
    await page.waitForFunction(() => window.__saved?.type === 'application/pdf', null, { timeout: 60000 });
    const pdf = await page.evaluate(() => ({ size: __saved.size, header: __saved.text.slice(0, 8), count: /\/Count (\d+)/.exec(__saved.text)?.[1], images: (__saved.text.match(/\/Filter \/DCTDecode/g) || []).length }));
    assert.equal(pdf.header, '%PDF-1.4'); assert.equal(pdf.count, '4'); assert.equal(pdf.images, 4); assert.ok(pdf.size > 10000);
    assert.equal(await page.evaluate(() => JSON.stringify(tabula.wb().serialize()) === __beforePrint), true);
  });
  await test('본문 높이는 여백만 제외 · 기본 머리글 없음 · 지정 머리글은 여백 내부', async () => {
    await reset({ ...pg, area: { r1: 0, c1: 0, r2: 4, c2: 4 }, margins: { top: 1, bottom: .5, left: .25, right: .75, header: .2, footer: .2 }, footer: '&P' });
    const geometry = await dialog().locator('.wixel-print-page').evaluate(node => ({ bodyH: parseFloat(node.querySelector('.wixel-print-body').style.height), headerText: node.querySelector('.wixel-print-hf').textContent, headerTop: parseFloat(node.querySelector('.wixel-print-hf').style.top), footerBottom: parseFloat(node.lastElementChild.style.bottom) }));
    assert.deepEqual(geometry, { bodyH: 912, headerText: '', headerTop: 19.2, footerBottom: 19.2 });
  });
  await test('오른쪽 먼저 출력 순서도 실제 인쇄 쪽에 반영', async () => {
    await reset({ ...pg, area: { r1: 0, c1: 0, r2: 11, c2: 11 }, order: 'overThenDown' });
    assert.deepEqual((await pages()).map(p => [p.rows[0], p.columns[0]]), [[0, 0], [0, 8], [10, 0], [10, 8]]);
  });
  await test('큰 글꼴·줄 바꿈·병합은 행을 늘리지 않고 셀 안에서 잘라 표시', async () => {
    const cells = {};
    for (let r = 0; r <= 10; r++) for (let c = 0; c < 3; c++) cells[r + ',' + c] = { raw: r === 10 ? '마지막 행' : '큰 글꼴과 긴 줄 바꿈 내용', style: { size: 50, font: 'Arial', wrap: true, fill: r === 10 ? '#00ff00' : '#ffe699' } };
    await reset({ ...pg, area: { r1: 1, c1: 0, r2: 10, c2: 2 }, titleRows: [0, 0], margins: { top: 4, bottom: 4, left: .25, right: .25, header: .2, footer: .2 } }, {
      cells, defRowH: 18, merges: [{ r1: 4, c1: 1, r2: 5, c2: 1 }],
    });
    const rows = await dialog().locator('.wixel-print-page').evaluate(node => {
      const zoom = Number(node.style.zoom), body = node.querySelector('.wixel-print-body').getBoundingClientRect();
      return [...node.querySelectorAll('tr')].map(row => { const rect = row.getBoundingClientRect(), first = row.querySelector('.wixel-print-cell-content'), css = getComputedStyle(first); return { r: Number(row.dataset.printRow), height: rect.height / zoom, inPage: rect.bottom <= body.bottom + .1,
        x: rect.x, y: rect.y, w: rect.width, h: rect.height, wrapperH: first.getBoundingClientRect().height / zoom, leading: parseFloat(css.lineHeight) / parseFloat(css.fontSize),
        clipped: [...row.querySelectorAll('.wixel-print-cell-content')].every(n => getComputedStyle(n).overflow === 'hidden') }; });
    });
    assert.equal(rows.length, 11); assert.ok(rows.every(r => Math.abs(r.height - 18) < .1 && r.inPage && r.clipped), JSON.stringify(rows));
    const all = await pages(); assert.equal(all.length, 1); assert.match(all[0].text, /마지막 행/);
    assert.deepEqual(all[0].merges.map(m => [m.cols, m.rows]), [[1, 2]]);
    const pixels = pngPixels(await page.screenshot({ path: join(output, 'fixed-row-large-font.png') })), inks = rows.map(r => pixels.ink(r));
    assert.ok(rows.every(r => Math.abs(r.wrapperH - 18) < .1 && Math.abs(r.leading - 1.2) < .01));
    assert.ok(inks.every(n => n > 0), '각 행의 잘린 글자도 실제 픽셀에 남아야 함: ' + inks.join(','));
    console.log('고정 행별 검정 픽셀: ' + inks.join(','));
  });
  assert.deepEqual(errors, []); assert.deepEqual(writes, []);
} finally {
  await writeFile(join(output, 'results.json'), JSON.stringify({ results, errors, writes }, null, 2));
  await browser.close();
}
const failed = results.filter(r => !r.ok).length;
console.log(`print-pagination: ${results.length - failed}/${results.length}, pageErrors ${errors.length}, writes ${writes.length}`);
if (failed || errors.length || writes.length) process.exitCode = 1;
