// Narrow numeric cells must not force a hash wider than the available content area.
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile, stat } from 'node:fs/promises';
import { resolve, join } from 'node:path';
import { pathToFileURL } from 'node:url';
import { execFileSync } from 'node:child_process';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const local = ['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname);
if (!local && process.env.WIXEL_ALLOWED_TEST_URL !== url) throw Error('공개 합성 검사는 WIXEL_ALLOWED_TEST_URL에 정확한 대상 URL을 지정하세요.');
if (process.env.REAL_WORKBOOK && !local) throw Error('사용자 실제 원본은 로컬 검증 서버에서만 열 수 있습니다.');
const output = resolve(process.env.NARROW_COLUMN_OUTPUT || (process.env.NARROW_COLUMN_BASELINE ? '.local/narrow-column/baseline' : '.local/narrow-column'));
await mkdir(output, { recursive: true });
const repository = resolve('..'), baseline = process.env.NARROW_COLUMN_BASELINE ? execFileSync('git', ['-c', 'safe.directory=' + repository, '-C', repository, 'show', process.env.NARROW_COLUMN_BASELINE + ':tabula/src/view.js'], { encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 }) : null;
const fallback = 'C:/Users/system777/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || pathToFileURL(fallback).href);
const zooms = (process.env.NARROW_COLUMN_ZOOMS || '25,55,100,200').split(',').map(Number);
if (zooms.some(z => !Number.isFinite(z) || z < 25 || z > 400)) throw Error('검사 배율은 25~400입니다.');
const widths = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 18, 20, 32, 48, 64, 96, 160];
const cases = [
  { label: '큰 숫자', raw: '1234567890123.45', style: { numFmt: 'number', decimals: 2 }, overflow: true },
  { label: '음수', raw: '-1234567890123.45', style: { numFmt: 'number', decimals: 2 }, overflow: true },
  { label: '날짜', raw: '45540', style: { numFmt: 'custom', code: 'yyyy/mm/dd' }, overflow: true },
  { label: '음수 날짜', raw: '-1', style: { numFmt: 'date' }, overflow: true },
  { label: '백분율', raw: '123456789.12', style: { numFmt: 'percent', decimals: 2 }, overflow: true },
  { label: 'General 소수', raw: '3.14159265358979', style: { numFmt: 'general' }, general: 'decimal' },
  { label: 'General 지수', raw: '1234567890123', style: { numFmt: 'general' }, general: 'exponent' },
  { label: '텍스트', raw: 'TEXT-123456789', style: { numFmt: 'general' }, expected: 'TEXT-123456789' },
  { label: '오류', raw: '=1/0', style: { numFmt: 'general' }, expected: '#DIV/0!' },
  { label: '맑은 고딕 숫자', raw: '1234567890123.45', style: { numFmt: 'number', decimals: 2, font: '맑은 고딕', size: 10 }, overflow: true },
  { label: '맑은 고딕 굵은 숫자', raw: '1234567890123.45', style: { numFmt: 'number', decimals: 2, font: '맑은 고딕', size: 10, bold: true }, overflow: true },
  { label: '짧은 General 3', raw: '3', style: { numFmt: 'general' }, shortExpected: '3' },
  { label: '짧은 General 0', raw: '0', style: { numFmt: 'general' }, shortExpected: '0' },
  { label: '숫자 모양 텍스트', raw: "'1234567890123456", style: { numFmt: 'general' }, expected: '1234567890123456' },
];
const cells = {}, colWidths = { 0: 150 };
widths.forEach((width, index) => { colWidths[index + 1] = width; cells['0,' + (index + 1)] = { raw: String(width), style: { bold: true, size: 8 } }; });
for (let r = 0; r < cases.length; r++) {
  cells[(r + 1) + ',0'] = { raw: cases[r].label };
  for (let c = 0; c < widths.length; c++) cells[(r + 1) + ',' + (c + 1)] = {
    raw: cases[r].raw, style: { font: 'Arial', size: 9, ...(cases[r].expected ? {} : { wrap: true }), ...cases[r].style },
  };
  cells[(r + 1) + ',' + (widths.length + 1)] = { raw: '끝' };
}
const fixture = join(output, 'narrow-numbers.xlsx');
await writeFile(fixture, writeXlsx(new Workbook({ defaultFont: { name: 'Arial', size: 9 }, sheets: [{ name: '좁은 열 검증', cells, colWidths, defRowH: 24 }] })));
const browser = await chromium.launch(), results = [], metrics = [], errors = [], writes = [];
let checks = 0, failure;
const ok = (value, message) => { assert.ok(value, message); checks++; };
const eq = (actual, expected, message) => { assert.deepEqual(actual, expected, message); checks++; };
async function isolated() {
  const context = await browser.newContext({ viewport: { width: 1700, height: 1000 }, deviceScaleFactor: 1 });
  const page = await context.newPage(); page.setDefaultTimeout(15000);
  page.on('pageerror', error => errors.push(error.message));
  await context.route('**/*', route => {
    const request = route.request(), target = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method()); return route.abort(); }
    if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) return route.abort();
    return route.continue();
  });
  if (baseline) await context.route('**/src/view.js', route => route.fulfill({ contentType: 'text/javascript', body: baseline }));
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
  return { context, page };
}
async function settle(page) {
  await page.evaluate(() => document.fonts.ready);
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
}
async function zoom(page, percent) {
  await page.locator('#zoomSlider').evaluate((input, percent) => { input.value = String(percent); input.dispatchEvent(new Event('input', { bubbles: true })); }, percent);
  await page.evaluate(() => { tabula.gv().setScroll(0, 0); tabula.gv().renderAll(); }); await settle(page);
}
async function test(name, fn) {
  try { await fn(); results.push({ name, ok: true }); console.log('OK ' + name); }
  catch (error) { results.push({ name, ok: false, error: error.message }); console.error('NG ' + name + ': ' + error.stack); }
}
async function measure(page, rows, cols) {
  return page.evaluate(({ rows, cols }) => {
    const t = tabula, g = t.gv(), w = t.wb(), si = t.si, canvas = document.createElement('canvas'), ctx = canvas.getContext('2d');
    const result = [];
    for (const r of rows) for (const c of cols) {
      const width = g.cols.size(c), node = [...document.querySelectorAll('.c[data-r="' + r + '"][data-c="' + c + '"]')].find(node => node.getBoundingClientRect().width > 0);
      if (!node) { result.push({ r, c, width, hidden: width === 0, absent: true }); continue; }
      const style = getComputedStyle(node), span = node.querySelector(':scope > span'), text = span?.textContent ?? '';
      const font = style.fontWeight + ' ' + style.fontSize + ' ' + style.fontFamily;
      ctx.font = font;
      const hashWidth = ctx.measureText('#').width, padding = parseFloat(style.paddingLeft) + parseFloat(style.paddingRight);
      const rect = node.getBoundingClientRect(), spanRect = span?.getBoundingClientRect();
      const value = w.getValue(si, r, c), sourceStyle = w.styleAt(si, r, c);
      result.push({ r, c, width, room: Math.max(0, width - padding), hashWidth, measuredText: ctx.measureText(text).width, text,
        numeric: typeof value === 'number', valueWidth: ctx.measureText(String(value)).width, raw: w.getRaw(si, r, c), shrink: !!sourceStyle.shrink,
        classes: node.className, whiteSpace: style.whiteSpace, overflow: style.overflow,
        spanWhiteSpace: span ? getComputedStyle(span).whiteSpace : null,
        spanWidth: spanRect ? spanRect.width / g.z : 0, spanHeight: spanRect ? spanRect.height / g.z : 0,
        cellHeight: rect.height / g.z, zoom: g.z * 100 });
    }
    return result;
  }, { rows, cols });
}
try {
  const { context, page } = await isolated();
  try {
    await page.locator('#fileInput').setInputFiles(fixture);
    await page.waitForFunction(() => tabula.wb().sheets[0]?.name === '좁은 열 검증' && !document.querySelector('.load-progress'), null, { timeout: 60000 });
    await settle(page);
    for (const percent of zooms) await test('합성 숫자·날짜·텍스트·오류 너비 0~20px 및 넓은 열 / ' + percent + '%', async () => {
      await zoom(page, percent);
      const snapshot = await measure(page, cases.map((_, index) => index + 1), widths.map((_, index) => index + 1));
      metrics.push({ kind: 'synthetic', zoom: percent, cells: snapshot });
      for (const cell of snapshot) {
        const sample = cases[cell.r - 1], label = sample.label + '/' + widths[cell.c - 1] + 'px/' + percent + '%';
        if (cell.hidden) { eq(cell.absent, true, label + ' 숨긴 0px열 DOM 생략'); continue; }
        ok(!cell.absent, label + ' 셀 DOM 표시');
        if (sample.expected !== undefined) {
          eq(cell.text, sample.expected, label + ' 문자열과 오류 텍스트 보존');
          ok(!/^#+$/.test(cell.text), label + ' 텍스트는 숫자 폭 경고로 변환하지 않음');
        } else {
          eq(cell.numeric, true, label + ' 원본 숫자 타입');
          ok(!cell.classes.split(' ').includes('wrap'), label + ' 숫자 자동 줄바꿈 방지');
          ok(['nowrap', 'pre'].includes(cell.spanWhiteSpace), label + ' 숫자 자동 줄바꿈 없는 한 줄 표시');
          if (sample.overflow && cell.room + .01 < cell.hashWidth) eq(cell.text, '', label + ' 넘친 숫자의 #도 안 들어가면 공백');
          if (cell.text && !/^#+$/.test(cell.text)) {
            ok(cell.measuredText <= cell.room + .08, label + ' 표시 숫자 폭이 실제 내용폭 안에 들어감');
            ok(cell.spanWidth <= cell.room + .08, label + ' DOM 숫자 폭도 실제 내용폭 안에 들어감');
          }
          if (sample.shortExpected && cell.valueWidth <= cell.room + .01) eq(cell.text, sample.shortExpected, label + ' 실제로 들어가는 짧은 숫자는 보존');
          if (sample.shortExpected && cell.valueWidth > cell.room + .01 && cell.room + .01 < cell.hashWidth) eq(cell.text, '', label + ' 숫자도 #도 들어갈 수 없을 때 공백');
          if (/^#+$/.test(cell.text)) {
            ok(cell.text.length >= 1, label + ' 들어갈 때는 # 유지');
            ok(cell.measuredText <= cell.room + .08, label + ' # 문자가 실제 셀 내용폭 안에 들어감');
            ok(cell.spanWidth <= cell.room + .08, label + ' DOM # 너비도 내용폭 안에 들어감');
            eq(cell.overflow, 'hidden', label + ' #가 인접 열로 새지 않음');
          }
          if (sample.overflow && widths[cell.c - 1] === 20) ok(/^#+$/.test(cell.text), label + ' 충분히 좁은 20px열에서는 적정 # 표시');
        }
      }
      const at = (label, width) => snapshot.find(cell => cases[cell.r - 1].label === label && widths[cell.c - 1] === width);
      ok(/^3(?:\.|$)/.test(at('General 소수', 64).text), 'General 소수는 좁아져도 반올림 숫자 표시');
      ok(/E\+\d+/.test(at('General 지수', 64).text), 'General 큰 숫자는 적절한 지수 표시');
      eq(at('큰 숫자', 160).text, '1,234,567,890,123.45', '넓은 열의 원래 숫자 표시 보존');
      if ([55, 100].includes(percent)) await page.screenshot({ path: join(output, 'synthetic-' + percent + '.png') });
    });
  } finally { await context.close(); }
  if (process.env.REAL_WORKBOOK) await test('사용자 원본 [구글] C열을 로컬로 열어 좁은 숫자 표시 확인', async () => {
    const file = resolve(process.env.REAL_WORKBOOK), before = await stat(file);
    let parsed = readXlsx(new Uint8Array(await readFile(file)));
    const si = parsed.data.sheets.findIndex(s => s.name.includes('구글'));
    ok(si >= 0, '구글 시트가 원본에 존재');
    const source = parsed.data.sheets[si], metadata = { file, sheet: source.name, si, colWidth: source.colWidths[2] ?? source.defColW, font: parsed.data.defaultFont };
    parsed = null;
    const { context, page } = await isolated();
    try {
      await page.locator('#fileInput').setInputFiles(file);
      await page.waitForFunction(name => tabula.wb().sheets.some(s => s.name === name) && !document.querySelector('.load-progress'), metadata.sheet, { timeout: 300000 });
      await page.evaluate(si => { tabula.switchSheet(si); tabula.gv().setScroll(0, 0); tabula.gv().renderAll(); }, si); await settle(page);
      for (const percent of zooms) {
        await zoom(page, percent);
        await page.evaluate(() => { const g = tabula.gv(); tabula.selectCell(105, 2); g.setScroll(0, Math.max(0, g.rows.pos(104) - g.boundaryY)); g.renderAll(); }); await settle(page);
        const snapshot = await measure(page, Array.from({ length: 24 }, (_, r) => r + 105), [2]);
        metrics.push({ kind: 'real', zoom: percent, metadata,
          narrowCandidates: snapshot.filter(c => !c.absent && c.numeric && !c.shrink && c.room + .01 < c.hashWidth).length,
          displayedHashCells: snapshot.filter(c => /^#+$/.test(c.text)).length,
          blankNumericCells: snapshot.filter(c => c.numeric && c.text === '').length, cells: snapshot });
        if ([55, 100].includes(percent)) await page.screenshot({ path: join(output, 'real-google-' + percent + '.png') });
        let narrowNumeric = 0;
        for (const cell of snapshot.filter(c => !c.absent && c.numeric && !c.shrink)) {
          if (cell.room + .01 < cell.hashWidth) { eq(cell.text, '', metadata.sheet + '!C' + (cell.r + 1) + '/' + percent + '% 공백'); narrowNumeric++; }
          if (/^#+$/.test(cell.text)) ok(cell.measuredText <= cell.room + .08, metadata.sheet + '!C' + (cell.r + 1) + ' 폭 안의 #');
        }
        eq(narrowNumeric, 24, '구글 C106:C129의 실제 숫자 24개를 검증함 / ' + percent + '%');
        console.log('OK real ' + metadata.sheet + '!C106:C129 ' + percent + '% blank=' + narrowNumeric);
      }
      await zoom(page, 100);
      await page.evaluate(() => { const t = tabula, w = t.wb(), g = t.gv(); w.undoStack = []; w.redoStack = []; w.transact(() => w.setColWidth(t.si, 2, 80)); g.layout(); g.setScroll(0, Math.max(0, g.rows.pos(104) - g.boundaryY)); g.renderAll(); }); await settle(page);
      const widened = await measure(page, Array.from({ length: 24 }, (_, r) => r + 105), [2]);
      for (let index = 0; index < widened.length; index++) {
        eq(widened[index].text, String(index), '구글 C' + (index + 106) + ' 너비 80px에서 원래 값 표시 복원');
        eq(widened[index].raw, String(index), '구글 C' + (index + 106) + ' 원본 값은 좁은 표시에도 보존');
      }
      metrics.push({ kind: 'real-wide', metadata, cells: widened });
      await page.screenshot({ path: join(output, 'real-google-widened-100.png') });
      await page.evaluate(() => { tabula.run('undo'); const g = tabula.gv(); g.layout(); g.setScroll(0, Math.max(0, g.rows.pos(104) - g.boundaryY)); g.renderAll(); }); await settle(page);
      const restored = await measure(page, Array.from({ length: 24 }, (_, r) => r + 105), [2]);
      for (const cell of restored) eq(cell.text, '', '폭 확장 Undo 후 C' + (cell.r + 1) + ' 공백 표시 복원');
      metrics.push({ kind: 'real-restored', metadata, cells: restored });
      const after = await stat(file); eq(after.size, before.size, '원본 파일 크기 보존'); eq(after.mtimeMs, before.mtimeMs, '원본 파일 수정 시각 보존');
    } finally { await context.close(); }
  });
  eq(errors, [], '페이지 오류 없음'); eq(writes, [], 'API와 외부 쓰기 없음');
} catch (error) { failure = error; }
finally {
  await writeFile(join(output, 'narrow-column-metrics.json'), JSON.stringify({ url, checks, results, metrics, errors, writes, error: failure?.message }, null, 2));
  await browser.close();
}
console.log(JSON.stringify({ checks, total: results.length, passed: results.filter(r => r.ok).length, failed: results.filter(r => !r.ok).length, output }));
if (failure) throw failure;
assert.equal(results.every(r => r.ok), true, '좁은 열의 숫자 표시 회귀 검사');
