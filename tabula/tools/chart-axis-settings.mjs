// 합성 차트만 사용하는 축 서식 UI/Undo/XLSX 검사. OS 파일 선택은 메모리 저장으로 대체한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { readXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, descendants } from '../src/xml.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const origin = new URL(url).origin;
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw new Error('로컬 서버에서만 합성 검사를 실행하세요.');
const out = process.env.WIXEL_AXIS_SETTINGS_OUT || 'D:/Codex/Temp/wixel-axis-settings/' + engine;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch(), results = [], assets = new Set(); let checks = 0;
const eq = (a, b, m) => { checks++; assert.deepEqual(a, b, m); };
const ok = (v, m) => { checks++; assert.ok(v, m); };
const run = (p, cmd) => p.evaluate(cmd => tabula.run(cmd), cmd);
const model = p => p.evaluate(() => structuredClone(tabula.wb().sheets[0].charts[0]));
const state = p => p.evaluate(() => ({ chart: structuredClone(tabula.wb().sheets[0].charts[0]), cells: [...tabula.wb().sheets[0].cells].map(([k, v]) => [k, v.raw]), undo: tabula.wb().undoStack.length }));
const pane = p => p.getByRole('dialog').filter({ has: p.locator('.cfp') });
const control = (p, name) => pane(p).getByLabel(name, { exact: true });
const number = async (p, name, value) => { await control(p, name).fill(String(value)); await control(p, name).press('Tab'); };
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function fixture(p, type) {
  await p.evaluate(type => {
    const cells = {}; [['광고', '매출'], ...['브랜드검색_모바일', '브랜드검색_PC', '쇼핑검색_모바일', '쇼핑검색_PC', '파워링크_모바일', '파워링크_PC'].map((s, i) => [s, String((i + 1) * 10)])].forEach((row, r) => row.forEach((raw, c) => cells[r + ',' + c] = { raw }));
    const w = tabula.wb(); w.restore({ sheets: [{ name: '축 검사', cells, charts: [{ id: 'axis', type, x: 30, y: 25, w: 650, h: 390, range: { r1: 0, c1: 0, r2: 6, c2: 1 }, title: '합성 차트', legend: 'none', axes: { x: { title: '광고 이름' }, y: { title: '금액', min: 0, max: 70, major: 10 } } }] }, { name: '빈 시트', cells: {} }] });
    tabula.switchSheet(1); tabula.switchSheet(0); tabula.gv().setZoom(100); w.undoStack = []; w.redoStack = [];
  }, type); await raf(p);
}
async function open(p, full) {
  // WebKit's SVG text bounding rectangle can ignore text-anchor=end. Click a
  // measured glyph in screen coordinates, without forcing through another shape.
  const point = await p.locator('.pane-br .obj[data-id="axis"] [data-el="axis-x"] text').first().evaluate(node => {
    const b = node.getExtentOfChar(Math.floor(node.getNumberOfChars() / 2)), m = node.getScreenCTM();
    const p = new DOMPoint(b.x + b.width / 2, b.y + b.height / 2).matrixTransform(m);
    return { x: p.x, y: p.y, hit: document.elementFromPoint(p.x, p.y)?.closest('[data-el]')?.getAttribute('data-el') };
  });
  eq(point.hit, 'axis-x', '실제 글자 지점에 축이 노출되어 있음');
  await p.mouse.click(point.x, point.y);
  await run(p, 'chartFormat'); await control(p, '서식을 지정할 차트 요소').waitFor();
  eq(await control(p, '서식을 지정할 차트 요소').inputValue(), 'axis-x', '실제 축 클릭은 항목 축 서식을 선택');
  if (full) await pane(p).getByRole('button', { name: '차트 전체 옵션…', exact: true }).click();
}
async function exported(p) {
  const bytes = Uint8Array.from(await p.evaluate(async () => { await tabula.exportXlsx('합성 축 검사', 'xlsx'); return window.__axisSaved; }));
  ok(bytes.length > 100, '메모리 XLSX 저장 완료');
  const xml = parseXml(textOf(unzip(bytes)['xl/charts/chart1.xml']));
  return { chart: readXlsx(bytes).data.sheets[0].charts[0], axis: descendants(xml, 'catAx')[0] };
}
async function test(name, fn, mobile = false) {
  const c = await browser.newContext({ viewport: { width: 1200, height: 900 }, serviceWorkers: 'block', hasTouch: mobile });
  const p = await c.newPage(), errors = [], writes = [], begin = checks;
  p.setDefaultTimeout(15000); p.on('pageerror', e => errors.push(e.message));
  await c.addInitScript(mobile => {
    window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
    localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', mobile ? 'on' : 'off');
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: async () => ({ kind: 'file', name: '합성 축 검사.xlsx', async createWritable() { return { async write(blob) { window.__axisSaved = Array.from(new Uint8Array(await blob.arrayBuffer())); }, async close() {}, async abort() {} }; } }) });
  }, mobile);
  await c.route('**/*', r => { const q = r.request(), u = new URL(q.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { writes.push(q.method() + ' ' + u.pathname); return r.abort(); } return u.origin === origin && !u.pathname.startsWith('/api/') ? r.continue() : r.abort(); });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    for (const src of await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src')))) assets.add(src);
    await fn(p); eq(errors, [], '페이지 오류 없음'); eq(writes, [], '원격 쓰기 없음');
    results.push({ name, ok: true, checks: checks - begin }); console.log('OK ' + name);
  } catch (e) { results.push({ name, ok: false, checks: checks - begin, error: e.stack, errors, writes }); console.error('NG ' + name + ' ' + e.stack); await p.screenshot({ path: out + '/failure-' + results.length + '.png' }).catch(() => {}); }
  finally { await c.close(); }
}
try {
  for (const type of ['column', 'bar']) for (const full of [false, true]) await test(type + (full ? ' 전체 옵션·모바일' : ' 선택 축 옵션') + ' 간격·각도·Undo·XLSX', async p => {
    await fixture(p, type); const before = await state(p); await open(p, full);
    eq(await control(p, '레이블 간격').inputValue(), 'auto'); eq(await control(p, '텍스트 방향').inputValue(), 'auto');
    eq(await control(p, '간격 단위').count(), 0); eq(await control(p, '사용자 지정 각도(°)').count(), 0);
    await control(p, '레이블 간격').selectOption('manual'); eq((await model(p)).axes.x.labelInterval, 1); eq((await state(p)).undo, before.undo + 1);
    await number(p, '간격 단위', 3); eq((await model(p)).axes.x.labelInterval, 3);
    await run(p, 'undo'); await raf(p); eq((await model(p)).axes.x.labelInterval, 1, '간격 Undo');
    await run(p, 'redo'); await raf(p); eq((await model(p)).axes.x.labelInterval, 3, '간격 Redo');
    await control(p, '텍스트 방향').selectOption('manual'); eq((await model(p)).axes.x.labelRotation, 0, '명시 0°');
    await number(p, '사용자 지정 각도(°)', -30); eq((await model(p)).axes.x.labelRotation, -30);
    const selected = await model(p); eq(selected.axes.x.title, before.chart.axes.x.title); eq(selected.axes.y, before.chart.axes.y, '다른 축 설정 유지'); eq((await state(p)).cells, before.cells, '원본 데이터 유지');
    const saved = await exported(p); eq(saved.chart.axes.x, selected.axes.x, 'XLSX 설정 왕복'); eq(child(saved.axis, 'tickLblSkip')?.attrs.val, '3'); eq(child(child(saved.axis, 'txPr'), 'bodyPr')?.attrs.rot, '-1800000');
    await p.screenshot({ path: out + '/' + type + (full ? '-full' : '-selected') + '.png' });
    await control(p, '레이블 간격').selectOption('auto'); await control(p, '텍스트 방향').selectOption('auto');
    eq((await model(p)).axes.x.labelInterval, undefined); eq((await model(p)).axes.x.labelRotation, undefined); eq((await state(p)).cells, before.cells);
    const automatic = await exported(p); eq(child(automatic.axis, 'tickLblSkip'), null); eq(child(automatic.axis, 'txPr'), null, '자동 방향은 표준XML 속성 생략');
  }, full);
  for (const type of ['scatter', 'bubble']) await test(type + ' 숫자 가로축에는 항목 간격 설정을 노출하지 않음', async p => {
    await fixture(p, type); await open(p, false);
    eq(await control(p, '레이블 간격').count(), 0); eq(await control(p, '텍스트 방향').count(), 0);
    eq(await control(p, '최소값').count(), 1); eq(await control(p, '주 단위').count(), 1);
    await pane(p).getByRole('button', { name: '차트 전체 옵션…', exact: true }).click();
    eq(await control(p, '레이블 간격').count(), 0); eq(await control(p, '텍스트 방향').count(), 0);
  });
  await test('열린 축 서식에서 시트 보호 후 변경은 모델을 바꾸지 않음', async p => {
    await fixture(p, 'column'); await open(p, false); await p.evaluate(() => { tabula.wb().sheets[0].protect = { on: true, allow: {} }; }); const before = await state(p);
    await control(p, '레이블 간격').selectOption('manual'); eq(await state(p), before, '보호 상태에서 간격 변경 거부');
  });
} finally {
  await browser.close(); const summary = { engine, url, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
  await writeFile(out + '/result.json', JSON.stringify(summary, null, 2)); console.log(JSON.stringify({ ...summary, results: undefined, out })); if (summary.passed !== summary.cases) process.exitCode = 1;
}
