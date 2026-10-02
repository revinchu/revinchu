// 합성 슬라이서 우클릭·보호·Undo 회귀. 업무 파일/사용자 탭/원격 쓰기를 사용하지 않습니다.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const browser = await chromium.launch(), results = []; let checks = 0;
const eq = (a, b, msg) => { assert.deepEqual(a, b, msg); checks++; };
const ok = (a, msg) => { assert.ok(a, msg); checks++; };
const menu = p => p.locator('#menuLayer > .menu[data-level="0"]');
const item = (p, name) => menu(p).getByRole('menuitem', { name });
const dialog = (p, name) => p.getByRole('dialog', { name, exact: true });
const sl = p => p.evaluate(() => window.tabula.wb().sheets[0].slicers?.[0]);
const undoCount = p => p.evaluate(() => window.tabula.wb().undoStack.length);
const command = (p, id) => p.evaluate(id => window.tabula.run(id), id);
const obj = p => p.locator('.obj[data-id="sl"]').first();
function fixture(pivots = false) {
  const data = { sheets: [{ name: '자료', cells: { '0,0': { raw: '지역' }, '0,1': { raw: '매출' }, '1,0': { raw: '서울' }, '1,1': { raw: '10' }, '2,0': { raw: '부산' }, '2,1': { raw: '20' }, '3,0': { raw: '서울' }, '3,1': { raw: '30' } }, tables: [{ id: 't1', name: '표1', r1: 0, c1: 0, r2: 3, c2: 1, header: true, filter: { criteria: {}, hidden: {} } }], slicers: [{ id: 'sl', caption: '지역', name: '지역슬라이서', source: { kind: 'table', table: '표1', column: '지역' }, x: 420, y: 100, w: 192, h: 240, z: 2 }], shapes: [{ id: 'shape', kind: 'rect', x: 700, y: 100, w: 60, h: 60, z: 1 }] }, { name: '보고서', cells: {} }, { name: '다른 원본', cells: { '0,0': { raw: '지역' }, '0,1': { raw: '매출' }, '1,0': { raw: '제주' }, '1,1': { raw: '99' } } }] };
  if (pivots) {
    const def = (name, source, top) => ({ name, source, range: { r1: 0, c1: 0, r2: source === '자료' ? 3 : 1, c2: 1 }, rows: ['지역'], cols: [], values: [{ field: '매출', name: '총매출' }], top, left: 0, area: { r1: top, c1: 0, r2: top + 4, c2: 1 } });
    data.sheets[1].pivot = def('피벗1', '자료', 0); data.sheets[1].pivotsExtra = [def('피벗2', '자료', 7), def('다른피벗', '다른 원본', 14)];
    data.sheets[0].slicers[0].source = { kind: 'pivot', field: '지역', pivots: [{ sheet: '보고서', name: '피벗1' }] };
  }
  return data;
}
async function open(p) { await obj(p).locator('.sl-cap').click({ button: 'right' }); await menu(p).waitFor(); }
async function confirm(p, title) { await dialog(p, title).getByRole('button', { name: '확인', exact: true }).click(); }
async function test(name, fn, { pivots = false, readonly = false, snapshots = false } = {}) {
  if (process.env.WIXEL_SLICER_FILTER && !name.includes(process.env.WIXEL_SLICER_FILTER)) return;
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }), p = await context.newPage(), errors = [], writes = [], start = checks;
  p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); } if (u.origin !== new URL(url).origin || u.pathname.startsWith('/api/')) return route.abort(); return route.continue(); });
  try {
    await p.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.__slicerClipboard = ''; Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { writeText: async t => { window.__slicerClipboard = t; }, readText: async () => window.__slicerClipboard, write: async () => {}, read: async () => [] } }); });
    const f = fixture(pivots); if (snapshots) { f.sheets[1].pivot.snapshotId = 'linked'; f.sheets[1].pivotsExtra[1].snapshotId = 'other'; f.pivotSnapshots = { linked: [['지역', '매출'], ['옛 지역', 5]], other: [['지역', '매출'], ['다른 옛 지역', 9]] }; } const target = readonly ? url.split('#')[0] + '#view=' + gzipSync(JSON.stringify({ docName: '슬라이서 합성', workbook: f, view: { headers: true } })).toString('base64url') : url;
    await p.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    if (readonly) await p.locator('.view-bar').waitFor();
    else await p.evaluate(data => { const t = window.tabula, w = t.wb(); w.restore(data); w.setSnapshots(data); t.switchSheet(0); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; }, f);
    await obj(p).waitFor(); await fn(p); eq(errors, [], '페이지 오류 없음'); eq(writes, [], '원격 쓰기 없음'); results.push({ name, ok: true, checks: checks - start }); console.log(`OK ${name} (${checks - start})`);
  } catch (e) { const ui = await p.evaluate(() => ({ toast: document.querySelector('#toast')?.innerText, menus: [...document.querySelectorAll('#menuLayer > .menu')].map(n => n.innerText), dialogs: [...document.querySelectorAll('.dialog')].map(n => n.innerText) })).catch(() => null); results.push({ name, ok: false, checks: checks - start, error: e.message, errors, ui }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
try {
  await test('우클릭 메뉴·정렬·다중 선택·Undo', async p => {
    const original = await sl(p); await open(p);
    for (const name of [/새로 고침/, /오름차순/, /내림차순/, /필터 지우기/, /다중 선택/, /보고서 연결/, /제거/, /그룹화/, /맨 앞으로/, /맨 뒤로/, /매크로 지정/, /대체 텍스트/, /크기 및 속성/, /슬라이서 설정/]) eq(await item(p, name).count(), 1);
    eq(await item(p, /그룹화/).isDisabled(), true); eq(await item(p, /보고서 연결/).isDisabled(), true);
    if (process.env.WIXEL_SLICER_SCREENSHOT) await p.screenshot({ path: process.env.WIXEL_SLICER_SCREENSHOT });
    await item(p, /내림차순/).click(); eq((await sl(p)).sort, 'desc'); eq(await undoCount(p), 1); await command(p, 'undo'); eq(await sl(p), original);
    await open(p); await item(p, /다중 선택/).click(); eq((await sl(p)).multi, true); await command(p, 'undo'); eq(await sl(p), original);
  });
  await test('실제 항목 필터·필터 지우기·Undo', async p => {
    await obj(p).locator('.sl-item').filter({ hasText: /^서울$/ }).click();
    eq(await p.evaluate(() => window.tabula.wb().sheets[0].tables[0].filter.criteria[0]), ['서울']);
    await open(p); eq(await item(p, /필터 지우기/).isDisabled(), false); await item(p, /필터 지우기/).click();
    eq(await p.evaluate(() => window.tabula.wb().sheets[0].tables[0].filter.criteria), {}); await command(p, 'undo');
    eq(await p.evaluate(() => window.tabula.wb().sheets[0].tables[0].filter.criteria[0]), ['서울']);
  });
  await test('크기·위치·잠금 초안 취소와 cm 저장·한 번 Undo', async p => {
    const original = await sl(p); await open(p); await item(p, /크기 및 속성/).click(); let d = dialog(p, '슬라이서 크기 및 속성');
    await d.getByLabel('너비(cm)', { exact: true }).fill('2.54'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await sl(p), original); eq(await undoCount(p), 0);
    await open(p); await item(p, /크기 및 속성/).click(); d = dialog(p, '슬라이서 크기 및 속성'); await d.getByLabel('너비(cm)', { exact: true }).fill('2.54'); await d.locator('select[name=placement]').selectOption('absolute'); await d.getByLabel('잠금', { exact: true }).uncheck(); await confirm(p, '슬라이서 크기 및 속성');
    eq((await sl(p)).w, 96); eq((await sl(p)).h, 240); eq((await sl(p)).placement, 'absolute'); eq((await sl(p)).locked, false); eq(await undoCount(p), 1); await command(p, 'undo'); eq(await sl(p), original);
  });
  await test('대체 텍스트·매크로 이름 편집과 기존 값 재열기·Undo', async p => {
    await open(p); await item(p, /대체 텍스트/).click(); await dialog(p, '슬라이서 대체 텍스트').locator('textarea[name=alt]').fill('지역별 매출 필터'); await confirm(p, '슬라이서 대체 텍스트'); eq((await sl(p)).alt, '지역별 매출 필터');
    await open(p); await item(p, /대체 텍스트/).click(); eq(await dialog(p, '슬라이서 대체 텍스트').locator('textarea[name=alt]').inputValue(), '지역별 매출 필터'); await dialog(p, '슬라이서 대체 텍스트').getByRole('button', { name: '취소', exact: true }).click();
    await open(p); await item(p, /매크로 지정/).click(); ok((await dialog(p, '매크로 지정').innerText()).includes('VBA 코드를 실행하지 않습니다')); await dialog(p, '매크로 지정').getByLabel(/매크로 이름/).fill('Module1.RefreshReport'); await confirm(p, '매크로 지정'); eq((await sl(p)).macro, 'Module1.RefreshReport'); await command(p, 'undo'); eq((await sl(p)).macro, undefined); eq((await sl(p)).alt, '지역별 매출 필터');
  });
  await test('설정·순서 하위 메뉴·제거는 각각 한 번 Undo', async p => {
    const original = await sl(p); await open(p); await item(p, /슬라이서 설정/).click(); const d = dialog(p, '슬라이서 설정'); await d.getByLabel('캡션', { exact: true }).fill('변경된 지역'); await confirm(p, '슬라이서 설정'); eq((await sl(p)).caption, '변경된 지역'); await command(p, 'undo'); eq(await sl(p), original);
    await open(p); await item(p, /맨 뒤로 보내기/).click(); await p.locator('#menuLayer > .menu[data-level="1"]').getByRole('menuitem', { name: '맨 뒤로 보내기', exact: true }).click(); ok(await p.evaluate(() => window.tabula.wb().sheets[0].slicers[0].z < window.tabula.wb().sheets[0].shapes[0].z)); await command(p, 'undo'); eq(await sl(p), original);
    await open(p); await item(p, /제거/).click(); eq(await sl(p), undefined); await command(p, 'undo'); eq(await sl(p), original);
  });
  await test('열린 설정 대상이 다른 시트로 바뀌면 저장 차단', async p => {
    const original = await sl(p); await open(p); await item(p, /대체 텍스트/).click(); await dialog(p, '슬라이서 대체 텍스트').locator('textarea[name=alt]').fill('저장하면 안 되는 초안'); await p.evaluate(() => window.tabula.switchSheet(1)); await confirm(p, '슬라이서 대체 텍스트'); eq(await sl(p), original); eq(await undoCount(p), 0); eq(await dialog(p, '슬라이서 대체 텍스트').count(), 1);
  });
  await test('보호 시트는 개체 편집을 막고 허용된 필터만 작동', async p => {
    await p.evaluate(() => { const w = window.tabula.wb(); w.sheets[0].protect = { on: true, allow: { autoFilter: true } }; window.tabula.gv().renderAll(); });
    await open(p); eq(await item(p, /슬라이서 설정/).isDisabled(), true); eq(await item(p, /제거/).isDisabled(), true); eq(await item(p, /새로 고침/).isDisabled(), false); await item(p, /다중 선택/).click(); eq((await sl(p)).multi, true);
    await p.evaluate(() => { window.tabula.wb().sheets[0].protect = { on: true, allow: {} }; window.tabula.gv().renderAll(); }); await open(p); eq(await item(p, /다중 선택/).isDisabled(), true); eq(await item(p, /새로 고침/).isDisabled(), true);
  });
  await test('공개 읽기 전용은 복사 외 편집·필터를 막음', async p => {
    await open(p); for (const name of [/새로 고침/, /다중 선택/, /내림차순/, /제거/, /크기 및 속성/, /매크로 지정/]) eq(await item(p, name).isDisabled(), true); eq(await item(p, /^복사/).isDisabled(), false); eq(await undoCount(p), 0);
  }, { readonly: true });
  await test('보고서 연결은 같은 원본만 허용·선택 전파·Undo', async p => {
    await obj(p).locator('.sl-item').filter({ hasText: /^서울$/ }).click(); const original = await sl(p), before = await undoCount(p); await open(p); await item(p, /보고서 연결/).click(); const d = dialog(p, '보고서 연결 (지역)');
    const row = name => d.locator('label.conn-row').filter({ hasText: name }).getByRole('checkbox'); eq(await row('피벗1').isChecked(), true); eq(await row('피벗2').isEnabled(), true); eq(await row('다른피벗').isDisabled(), true); await row('피벗2').check(); await confirm(p, '보고서 연결 (지역)');
    eq((await sl(p)).source.pivots, [{ sheet: '보고서', name: '피벗1' }, { sheet: '보고서', name: '피벗2' }]); eq(await p.evaluate(() => window.tabula.wb().sheets[1].pivotsExtra[0].filters.지역), ['서울']); eq(await undoCount(p), before + 1); await command(p, 'undo'); eq(await sl(p), original);
  }, { pivots: true });
  await test('표 새로 고침은 원본 변경 후 필터 숨김을 다시 계산', async p => {
    await obj(p).locator('.sl-item').filter({ hasText: /^서울$/ }).click(); await p.evaluate(() => window.tabula.wb().setInput(0, 2, 0, '서울')); await open(p); await item(p, /새로 고침/).click();
    eq(await p.evaluate(() => window.tabula.wb().sheets[0].tables[0].filter.criteria[0]), ['서울']); eq(await obj(p).locator('.sl-item').count(), 1); eq(await p.evaluate(() => !!window.tabula.wb().sheets[0].tables[0].filter.hidden?.[2]), false);
  });
  await test('피벗 새로 고침은 연결된 저장 캐시만 갱신', async p => {
    eq(await obj(p).locator('.sl-item').allTextContents(), ['옛 지역']); const before = await undoCount(p); await open(p); await item(p, /새로 고침/).click();
    eq((await obj(p).locator('.sl-item').allTextContents()).sort(), ['부산', '서울']); eq(await p.evaluate(() => window.tabula.wb().pivotSnapshots.has('linked')), false); eq(await p.evaluate(() => window.tabula.wb().pivotSnapshots.has('other')), true); eq(await undoCount(p), before + 1);
  }, { pivots: true, snapshots: true });
  await test('열린 크기 창에서 보호가 켜지면 저장 거절', async p => {
    const original = await sl(p); await open(p); await item(p, /크기 및 속성/).click(); await dialog(p, '슬라이서 크기 및 속성').getByLabel('너비(cm)', { exact: true }).fill('3'); await p.evaluate(() => { window.tabula.wb().sheets[0].protect = { on: true, allow: {} }; }); await confirm(p, '슬라이서 크기 및 속성'); eq(await sl(p), original); eq(await undoCount(p), 0); eq(await dialog(p, '슬라이서 크기 및 속성').count(), 1);
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ cases: results.length, passed: results.filter(r => r.ok).length, checks, results }, null, 2));
if (results.some(r => !r.ok)) process.exitCode = 1;
