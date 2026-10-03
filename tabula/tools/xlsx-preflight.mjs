// Excel 저장 전 확인의 실제 UI 검사. 파일 선택/쓰기·외부 API는 메모리에서만 모의합니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5191/');
assert.ok(['127.0.0.1', 'localhost', '[::1]'].includes(url.hostname), '로컬 서버만 허용');
const out = process.env.WIXEL_PREFLIGHT_OUT || 'D:/Codex/Temp/wixel-xlsx-preflight/' + engine;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch(), results = [], assets = new Set();
let checks = 0;
function eq(a, b, note) { assert.deepEqual(a, b, note); checks++; }
function ok(a, note) { assert.ok(a, note); checks++; }
const d = p => p.getByRole('dialog', { name: 'Excel 호환성 확인', exact: true });
const ctrlSave = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+s'); };
const snap = p => p.evaluate(() => ({ json: JSON.stringify(tabula.wb().serialize()), version: tabula.wb().version, undo: tabula.wb().undoStack.length, title: document.getElementById('docTitle').textContent }));
const mock = p => p.evaluate(() => window.__saveMock);
const quiet = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function noWrite(p) { await quiet(p); const m = await mock(p); eq(m.calls.length, 0); eq(m.events, []); }
const completed = p => p.waitForFunction(() => window.__saveMock.events.includes('close'));
const savedBook = async p => new Workbook(readXlsx(Uint8Array.from((await mock(p)).bytes)).data);
async function setupWarning(p, type = 'table') {
  await p.evaluate(type => { const w = tabula.wb(); if (type === 'log') w.sheets[0].charts = [{ id: 'synthetic-chart', type: 'column', x: 30, y: 30, w: 420, h: 280, axes: { y: { logBase: 10 } }, series: [{ name: { text: '합성' }, cache: [1, 10, 100], catCache: ['가', '나', '다'] }] }];
    else { w.props = { ...w.props, title: '기존 메타', xlsxImportWarnings: ['dataTableValuesOnly'] }; w.calculation = { mode: 'manual', iterate: true, fullPrecision: false, calcOnSave: false }; w.manualCalc = true; }
  }, type);
}
function tableFile() {
  const w = new Workbook({ sheets: [{ name: 'TABLE 합성', cells: { '0,0': { raw: '2' }, '0,1': { raw: '=A1*2' }, '1,0': { raw: '2' }, '1,1': { raw: '4' }, '2,0': { raw: '3' }, '2,1': { raw: '6' } } }] });
  const files = unzip(writeXlsx(w)); files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml']).replace(/(<c\b[^>]*r="B2"[^>]*>)/, '$1<f t="dataTable" ref="B2:B3" r1="A1"/>');
  return Buffer.from(zip(files));
}
async function test(name, fn) {
  if (process.env.WIXEL_CASE && !name.includes(process.env.WIXEL_CASE)) return;
  const context = await browser.newContext({ viewport: { width: 1360, height: 920 }, acceptDownloads: true, serviceWorkers: 'block' });
  const p = await context.newPage(), errors = [], unexpected = [], downloads = []; p.setDefaultTimeout(18000);
  p.on('pageerror', e => errors.push(e.message)); p.on('download', x => downloads.push(x.suggestedFilename()));
  await context.route('**/*', r => { const q = r.request(), u = new URL(q.url());
    if (u.origin === url.origin && !u.pathname.startsWith('/api/') && ['GET', 'HEAD'].includes(q.method())) return r.continue();
    if (u.pathname === '/api/health') return r.fulfill({ status: 404, contentType: 'application/json', body: '{}' });
    unexpected.push(q.method() + ' ' + u.pathname); return r.abort();
  });
  await context.addInitScript(() => {
    window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
    localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off');
    localStorage.setItem('wixel.options', JSON.stringify({ saveConfirm: false, userName: '' }));
    localStorage.setItem('tabula.workbook.v1', JSON.stringify({ docName: '사전 확인 합성', autosave: false, remoteDoc: false, workbook: { sheets: [{ name: '기존', cells: { '0,0': { raw: '42' }, '1,0': { raw: '=A1*2', cached: 84 } } }] } }));
    const m = window.__saveMock = { calls: [], events: [], bytes: null, abort: false };
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: async options => { m.calls.push({ name: options.suggestedName, active: navigator.userActivation?.isActive });
      if (m.abort) throw new DOMException('합성 선택 취소', 'AbortError');
      return { kind: 'file', name: options.suggestedName, createWritable: async () => ({ write: async b => { m.events.push('write'); m.bytes = Array.from(new Uint8Array(await b.arrayBuffer())); }, close: async () => m.events.push('close'), abort: async () => m.events.push('abort') }) };
    } });
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {} } });
  });
  try {
    await p.goto(url.href, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb()?.getValue(0, 0, 0) === 42);
    for (const asset of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => new URL(n.src).pathname).filter(n => /wixel-[a-f0-9]+\.js$/.test(n)))) assets.add(asset);
    await fn(p); eq(errors, [], '페이지 오류'); eq(unexpected, [], '외부/API 쓰기 없음'); eq(downloads, [], '자동 다운로드 없음');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (e) { results.push({ name, ok: false, error: e.message, errors, unexpected }); console.error('NG ' + name + ': ' + e.stack); await p.screenshot({ path: out + '/failure-' + results.length + '.png' }).catch(() => {}); }
  finally { await context.close(); }
}
try {
  await test('경고 없는 정상 Ctrl+S는 picker 한 번과 값·수식 저장', async p => {
    const before = await snap(p); await ctrlSave(p); await completed(p); const m = await mock(p), w = await savedBook(p);
    eq(m.calls.length, 1); eq(m.events, ['write', 'close']); ok(m.calls[0].active); eq(w.getValue(0, 0, 0), 42); eq(w.getRaw(0, 1, 0), '=A1*2'); eq(w.getValue(0, 1, 0), 84); eq((await snap(p)).undo, before.undo);
    eq(await d(p).count(), 0);
  });
  for (const axis of ['행', '열']) await test(axis + ' 블록 한도 초과는 picker 이전 차단·값과 Undo 보존', async p => {
    await p.evaluate(axis => { tabula.wb().sheets[0].blocks.push(axis === '행'
      ? { r0: 1048575, c0: 1, n: 2, ver: 0, cols: [{ num: new Float64Array([11, 22]), str: null, dict: [], fmt: null }] }
      : { r0: 2, c0: 16383, n: 1, ver: 0, cols: [11, 22].map(n => ({ num: new Float64Array([n]), str: null, dict: [], fmt: null })) }); }, axis);
    const before = await snap(p); await ctrlSave(p); const alert = p.getByRole('dialog', { name: 'Excel 저장 범위 초과', exact: true }); await alert.waitFor();
    const text = await alert.innerText(); ok(text.includes('1,048,576') && text.includes('16,384') && text.includes('.wixel')); await noWrite(p); eq(await snap(p), before);
    await alert.getByRole('button', { name: '확인', exact: true }).click(); await noWrite(p);
  });
  await test('TABLE·반복 계산·표시 정밀도 경고 취소는 원본 불변', async p => {
    await setupWarning(p); const before = await snap(p); await ctrlSave(p); await d(p).waitFor(); const text = await d(p).innerText();
    ok(text.includes('TABLE') && text.includes('반복 계산') && text.includes('표시된 정밀도')); await noWrite(p); await d(p).getByRole('button', { name: '취소', exact: true }).click(); await noWrite(p); eq(await snap(p), before);
  });
  await test('로그 축 경고 승인 후 표준 XLSX 저장·재읽기', async p => {
    await setupWarning(p, 'log'); const before = await snap(p); await ctrlSave(p); await d(p).waitFor(); ok((await d(p).innerText()).includes('선형 축')); await noWrite(p);
    await d(p).getByRole('button', { name: '확인', exact: true }).click(); await completed(p); const m = await mock(p), w = await savedBook(p);
    eq(m.calls.length, 1); ok(m.calls[0].active, '확인 클릭 사용자 활성화'); eq(w.sheets[0].charts[0].axes.y.logBase, 10); eq(w.getValue(0, 0, 0), 42); eq((await snap(p)).undo, before.undo);
  });
  await test('경고 대기 중 셀 편집은 승인해도 저장 취소', async p => {
    await setupWarning(p); await ctrlSave(p); await d(p).waitFor();
    await p.evaluate(() => tabula.wb().transact(() => tabula.wb().setInput(0, 0, 0, '99'))); const changed = await snap(p);
    await d(p).getByRole('button', { name: '확인', exact: true }).click(); await noWrite(p); eq(await snap(p), changed);
  });
  await test('경고 대기 중 문서 교체는 새 문서 보존', async p => {
    await setupWarning(p); await ctrlSave(p); await d(p).waitFor();
    await p.locator('#fileInput').setInputFiles({ name: '교체 합성.wixel', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify({ docName: '교체 합성', workbook: { sheets: [{ name: '새 문서', cells: { '0,0': { raw: '123' } } }] } })) });
    await p.waitForFunction(() => tabula.wb().getValue(0, 0, 0) === 123); const changed = await snap(p);
    await d(p).getByRole('button', { name: '확인', exact: true }).click(); await noWrite(p); eq(await snap(p), changed);
  });
  await test('TABLE 시트 추가는 경고 합집합·1 Undo·Redo·XLSX 재읽기', async p => {
    await p.evaluate(() => { tabula.wb().props = { title: '기존 속성' }; }); const before = await snap(p);
    const chooser = p.waitForEvent('filechooser'); await p.evaluate(() => tabula.run('importCsv'));
    await (await chooser).setFiles({ name: '합성 TABLE.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: tableFile() });
    await p.getByRole('dialog', { name: '가져오기', exact: true }).waitFor(); await p.keyboard.press('Escape');
    const imported = await p.evaluate(() => ({ count: tabula.wb().sheets.length, props: tabula.wb().props, values: [tabula.wb().getValue(1, 1, 1), tabula.wb().getValue(1, 2, 1)], undo: tabula.wb().undoStack.length }));
    eq(imported, { count: 2, props: { title: '기존 속성', xlsxImportWarnings: ['dataTableValuesOnly'] }, values: [4, 6], undo: before.undo + 1 });
    await p.evaluate(() => tabula.run('undo')); eq((await snap(p)).json, before.json);
    await p.evaluate(() => tabula.run('redo')); eq(await p.evaluate(() => tabula.wb().props.xlsxImportWarnings), ['dataTableValuesOnly']);
    await ctrlSave(p); await d(p).waitFor(); ok((await d(p).innerText()).includes('TABLE')); await d(p).getByRole('button', { name: '확인', exact: true }).click(); await completed(p);
    const w = await savedBook(p); eq(w.props.xlsxImportWarnings, ['dataTableValuesOnly']); eq(w.props.title, '기존 속성'); eq(w.getValue(1, 1, 1), 4); eq(w.getValue(1, 2, 1), 6);
  });
  await test('범위 초과 문서는 .wixel 저장으로 전체 블록 보관 가능', async p => {
    await p.evaluate(() => { tabula.wb().sheets[0].blocks.push({ r0: 1048576, c0: 0, n: 2, ver: 0, cols: [{ num: new Float64Array([71, 72]), str: null, dict: [], fmt: null }] }); tabula.run('saveAs'); });
    const dialog = p.getByRole('dialog', { name: '다른 이름으로 저장', exact: true });
    await dialog.getByRole('combobox', { name: /^파일 형식/ }).selectOption('wixel'); await dialog.getByRole('button', { name: '저장', exact: true }).click(); await completed(p);
    const m = await mock(p); eq(m.calls.length, 1); ok(m.calls[0].name.endsWith('.wixel')); const data = JSON.parse(Buffer.from(m.bytes).toString('utf8')), w = new Workbook(data.workbook);
    eq([w.getValue(0, 1048576, 0), w.getValue(0, 1048577, 0)], [71, 72]); eq(w.getValue(0, 0, 0), 42);
  });
} finally { await browser.close(); }
const result = { url: url.href, assets: [...assets], engine, total: results.length, passed: results.filter(x => x.ok).length, checks, failed: results.filter(x => !x.ok), results, note: '로컬 합성 문서만. 모의 파일 picker·메모리 bytes; 실제 OS 저장창/사용자 파일/원격 쓰기 없음.' };
await writeFile(out + '/result.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify(result)); if (result.failed.length) process.exitCode = 1;
