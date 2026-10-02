// 실제 메뉴로 합성 피벗을 만들고 보고서 필터 배치·서식·선택·XLSX 재열기를 검사합니다.
// 사용자 문서/클립보드/원격 API를 읽거나 쓰지 않습니다. 실제 Excel 합성 파일 비교는 선택 인수입니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile, access } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_PIVOT_REPORT_OUT || 'D:/Codex/Temp/wixel-pivot-report-layout/source';
const only = process.env.WIXEL_PIVOT_REPORT_FILTER || '';
const native = process.env.WIXEL_PIVOT_REPORT_NATIVE_FIXTURE || '';
await mkdir(out, { recursive: true });
const browser = await chromium.launch(), results = [], errors = [], writes = [], blocked = [], measurements = [], assets = new Set();
let checks = 0;
const eq = (a, b, m) => { checks++; assert.deepEqual(a, b, m); }, ok = (a, m) => { checks++; assert.ok(a, m); };
const run = (p, id) => p.evaluate(id => window.tabula.run(id), id);
const def = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets[window.tabula.si].pivot));
const model = p => p.evaluate(() => { const t = window.tabula, w = t.wb(); return { doc: w.serialize(), undo: w.undoStack.length }; });
const dialog = (p, name) => p.getByRole('dialog', { name, exact: true });
const menu = p => p.locator('.filter-menu.pivot-filter-menu');
const confirm = d => d.getByRole('button', { name: '확인', exact: true });
const cancel = d => d.getByRole('button', { name: '취소', exact: true });
const raft = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function create(p, pages = ['지역', '채널', '연도'], body = false) {
  for (let i = 0; i < 3; i++) await p.keyboard.press('Escape');
  await p.evaluate(() => {
    const t = window.tabula, w = t.wb(); if (t.si) t.switchSheet(0); const cells = {};
    [['지역', '채널', '연도', '상품', '매출'], ['서울', '온라인', 2025, 'A', 100], ['부산', '오프라인', 2026, 'B', 200], ['서울', '오프라인', 2026, 'B', 300], ['대구', '온라인', 2025, 'A', 400]].forEach((row, r) => row.forEach((v, c) => { cells[r + ',' + c] = { raw: String(v) }; }));
    w.restore({ sheets: [{ name: '합성 원본', cells }] }); t.gv().setZoom(100); t.gv().layout(); t.gv().renderAll(); t.selectRange({ r1: 0, c1: 0, r2: 4, c2: 4 }); t.run('insertPivot');
  });
  await confirm(dialog(p, '피벗 테이블 만들기')).click();
  for (const field of pages) await addField(p, field, '보고서 필터에 추가');
  if (body) for (const field of ['상품', '매출']) await p.locator('#pivotPane').getByRole('checkbox', { name: field, exact: true }).check();
  await p.evaluate(() => { const w = window.tabula.wb(); w.undoStack = []; w.redoStack = []; }); await raft(p);
}
async function addField(p, field, action) {
  await p.locator('#pivotPane .pp-field').filter({ has: p.getByRole('checkbox', { name: field, exact: true }) }).click({ button: 'right' });
  await p.getByRole('menuitem', { name: action, exact: true }).click();
}
async function options(p, order, wrap, apply = true) {
  await run(p, 'pivotOptions'); const d = dialog(p, '피벗 테이블 옵션'); await d.waitFor();
  if (order !== undefined) await d.locator('.opt-box select').first().selectOption(order);
  if (wrap !== undefined) await d.locator('.opt-box input[type=number]').nth(1).fill(String(wrap));
  if (apply) { await confirm(d).click(); await d.waitFor({ state: 'detached' }); await raft(p); }
  return d;
}
async function snapshot(p, label) {
  const state = await p.evaluate(() => {
    const t = window.tabula, w = t.wb(), s = w.sheets[t.si], d = s.pivot;
    const fields = (d.buttons || []).filter(b => b.kind === 'page').map(b => {
      const button = [...document.querySelectorAll('.pbtn[data-k="page"]')].find(n => n.dataset.f === b.field && n.getClientRects().length), node = [...document.querySelectorAll('.c')].find(n => +n.dataset.r === b.r && +n.dataset.c === b.c && n.getClientRects().length);
      const r = button?.getBoundingClientRect(); let textRect = null; const text = node?.querySelector('span');
      if (text?.firstChild) { const range = document.createRange(); range.selectNodeContents(text); const q = range.getBoundingClientRect(); textRect = { left: q.left, right: q.right, top: q.top, bottom: q.bottom }; }
      return { field: b.field, r: b.r, c: b.c, label: w.getRaw(t.si, b.r, b.c - 1), value: w.getRaw(t.si, b.r, b.c), labelStyle: w.styleAt(t.si, b.r, b.c - 1), valueStyle: w.styleAt(t.si, b.r, b.c), cssAlign: node && getComputedStyle(node).textAlign, button: r ? { left: r.left, right: r.right, top: r.top, bottom: r.bottom } : null, text: textRect };
    });
    return { def: structuredClone(d), fields, source: Object.fromEntries(w.sheets[0].cells.entries()), undo: w.undoStack.length };
  }); measurements.push({ label, ...state }); return state;
}
async function placement(p, expected) {
  const d = await def(p), actual = d.buttons.filter(b => b.kind === 'page').map(b => [b.field, b.r - d.top, b.c - d.left - 1]); eq(actual.sort((a,b)=>a[0].localeCompare(b[0])), [...expected].sort((a,b)=>a[0].localeCompare(b[0])), '필드 이름과 필터 레이블 위치');
  for (const b of d.buttons.filter(b => b.kind === 'page')) eq(await p.locator('.pbtn[data-k="page"][data-f="' + b.field + '"]').count(), 1, b.field + ' 필터 단추 하나');
}
async function openFilter(p, field = '지역') { const b = p.locator('.pbtn[data-k="page"][data-f="' + field + '"]').first(); await b.scrollIntoViewIfNeeded(); await b.click(); await menu(p).waitFor(); return menu(p); }
async function contextField(p, field) {
  await p.evaluate(field => { const t = window.tabula, d = t.wb().sheets[t.si].pivot, b = d.buttons.find(b => b.kind === 'page' && b.field === field); t.selectCell(b.r, b.c - 1); }, field);
  await p.locator('#cellEditor').focus(); await p.keyboard.press('Shift+F10'); return p.locator('#menuLayer>.menu[data-context-kind="pivot"]');
}
async function shot(p, name) { await p.screenshot({ path: out + '/' + name + '.png' }); }
async function test(name, fn) {
  if (only && !name.includes(only)) return;
  const c = await browser.newContext({ viewport: { width: 1440, height: 1000 } }), p = await c.newPage(), pageErrors = [], pageWrites = [];
  p.setDefaultTimeout(10000); p.on('pageerror', e => { pageErrors.push(e.message); errors.push({ name, error: e.message }); }); p.on('dialog', d => d.type() === 'beforeunload' ? d.accept() : d.dismiss());
  await c.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {}, read: async () => [], write: async () => {} } }); });
  await c.route('**/*', r => { const q = r.request(), u = new URL(q.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(q.method())) { writes.push(q.method() + ' ' + u.pathname); pageWrites.push(u.pathname); return r.abort(); } if (u.origin !== new URL(url).origin || u.pathname.startsWith('/api/')) { blocked.push(u.pathname); return r.abort(); } return r.continue(); });
  try { await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb()); for (const s of await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src')))) assets.add(s); await create(p); await fn(p); eq(pageErrors, [], '페이지 오류'); eq(pageWrites, [], '원격 쓰기'); results.push({ name, ok: true }); console.log('OK ' + name); }
  catch (e) { results.push({ name, ok: false, error: e.message }); console.error('NG ' + name + ': ' + e.stack); await shot(p, 'failure-' + results.length).catch(() => {}); }
  finally { await c.close(); }
}
try {
  await test('UI로 필터 1·2·3개를 추가하면 이름·값·단추가 정확히 배치된다', async p => {
    await create(p, []);
    for (const [field, index] of [['지역', 0], ['채널', 1], ['연도', 2]]) {
      await addField(p, field, '보고서 필터에 추가'); const s = await snapshot(p, 'filters-only-' + (index + 1)); eq(s.fields.length, index + 1); if (index === 0) eq(s.def.top, 0, 'A3 새 피벗에 첫 필터를 넣으면 A1/B1을 사용'); eq(s.fields[index].value, '(모두)'); eq(s.fields[index].label, field); ok(['left','start'].includes(s.fields[index].cssAlign), '텍스트는 가운데 시트 기본값과 무관하게 왼쪽 정렬');
      for (const f of s.fields) { ok(f.button && f.text, '실제 셀·단추 렌더'); ok(f.text.right <= f.button.left + 1, '(모두)와 필터 단추 겹침 없음 ' + JSON.stringify(f)); }
      await shot(p, 'created-' + (index + 1) + '-filters');
    }
  });
  await test('행·값 추가와 제거는 보고서 필터의 서식·항목을 바꾸지 않는다', async p => {
    const before = await snapshot(p, 'filter-style-before-body');
    for (const field of ['상품', '매출']) await p.locator('#pivotPane').getByRole('checkbox', { name: field, exact: true }).check();
    let after = await snapshot(p, 'filter-style-with-body'); eq(after.fields.map(f => [f.labelStyle, f.valueStyle]), before.fields.map(f => [f.labelStyle, f.valueStyle])); ok((await def(p)).area.r2 > before.def.area.r2); await shot(p, 'filters-with-body');
    for (const field of ['상품', '매출']) await p.locator('#pivotPane').getByRole('checkbox', { name: field, exact: true }).uncheck();
    after = await snapshot(p, 'filter-style-after-remove'); eq(after.fields.map(f => [f.labelStyle, f.valueStyle]), before.fields.map(f => [f.labelStyle, f.valueStyle])); eq(after.def.area, before.def.area); eq(after.source, before.source);
  });
  await test('행 우선과 열 우선·줄 바꿈 0/2 옵션은 실제 셀 위치와 Undo에 반영된다', async p => {
    const samples = [['down', 0, [['지역', 0, 0], ['채널', 1, 0], ['연도', 2, 0]]], ['down', 2, [['지역', 0, 0], ['채널', 1, 0], ['연도', 0, 3]]], ['over', 2, [['지역', 0, 0], ['채널', 0, 3], ['연도', 1, 0]]], ['over', 0, [['지역', 0, 0], ['채널', 0, 3], ['연도', 0, 6]]]];
    for (const [order, wrap, expected] of samples) { const before = await def(p), count = await p.evaluate(() => window.tabula.wb().undoStack.length); await options(p, order, wrap); await placement(p, expected); eq(await p.evaluate(() => window.tabula.wb().undoStack.length), count + 1); await shot(p, order + '-' + wrap); await run(p, 'undo'); eq(await def(p), before); }
  });
  await test('배치 변경은 이전 필터 셀만 지우고 빈 열·오른쪽 사용자 내용을 보존한다', async p => {
    await create(p, ['지역', '채널', '연도'], true); await options(p, 'over', 2);
    await p.evaluate(() => { const t = window.tabula, w = t.wb(), d = w.sheets[t.si].pivot; w.transact(() => { w.setInput(t.si, d.top, d.left + 2, '필터 사이 보존'); w.setInput(t.si, d.top, d.left + 5, '오른쪽 보존'); w.setInput(t.si, d.top + 5, d.left + 2, '본문 오른쪽 보존'); }); });
    const coords = await p.evaluate(() => { const d = window.tabula.wb().sheets[window.tabula.si].pivot; return [[0,2],[0,5],[5,2]].map(([r,c])=>[d.top+r,d.left+c]); });
    const marker = () => p.evaluate(coords => { const t = window.tabula, w = t.wb(); return coords.map(([r,c])=>w.getRaw(t.si,r,c)); }, coords);
    const before = await marker(); await options(p, 'down', 0); eq(await marker(), before); const d = await def(p); eq(await p.evaluate(d => [window.tabula.wb().getRaw(window.tabula.si, d.top, d.left + 3), window.tabula.wb().getRaw(window.tabula.si, d.top, d.left + 4)], d), ['', ''], '이전 가로 필터 잔상 없음'); await run(p, 'undo'); eq(await marker(), before); await placement(p, [['지역', 0, 0], ['채널', 0, 3], ['연도', 1, 0]]);
  });
  await test('실행 취소 뒤 다른 배치와 다시 실행은 이전 필터 셀 잔상을 남기지 않는다', async p => {
    const checkCells = async () => {
      const state = await p.evaluate(() => {
        const t = window.tabula, w = t.wb(), d = w.sheets[t.si].pivot, expected = [], actual = [];
        for (const b of d.buttons.filter(b => b.kind === 'page')) expected.push([b.r, b.c - 1, d.fieldCaptions?.[b.field] || b.field], [b.r, b.c, '(모두)']);
        for (let r = 0; r < 10; r++) for (let c = 0; c < 10; c++) { const value = w.getRaw(t.si, r, c); if (value !== '') actual.push([r, c, value]); }
        return { expected: expected.sort((a,b)=>a[0]-b[0]||a[1]-b[1]), actual };
      });
      eq(state.actual, state.expected, '새 필터 위치 밖의 이전 레이블·값이 없음');
    };
    await options(p, 'over', 0); await checkCells(); await run(p, 'undo'); await checkCells();
    await options(p, 'over', 2); await checkCells(); await run(p, 'undo'); await checkCells(); await run(p, 'redo'); await checkCells(); await shot(p, 'undo-rearrange-no-ghosts');
  });
  await test('필터 사이 사용자 병합과 설명은 옵션 확인·재배치·Undo·Redo 후 보존된다', async p => {
    await options(p, 'over', 2);
    const before = await p.evaluate(() => {
      const t = window.tabula, w = t.wb(), d = w.sheets[t.si].pivot, r = d.top, c = d.left + 2;
      w.transact(() => { w.setInput(t.si, r, c, '사용자 설명'); w.merge(t.si, r, c, r + 1, c); });
      return { merges: structuredClone(w.sheets[t.si].merges), r, c, raw: w.getRaw(t.si, r, c) };
    });
    const checkMerge = async () => eq(await p.evaluate(({r,c}) => { const t=window.tabula,w=t.wb(); return {merges:structuredClone(w.sheets[t.si].merges),raw:w.getRaw(t.si,r,c)}; }, before), {merges:before.merges,raw:before.raw}, '비소유 간격 셀의 병합·설명 유지');
    await options(p); await checkMerge(); await run(p, 'undo'); await checkMerge(); await run(p, 'redo'); await checkMerge(); await options(p, 'over', 0); await checkMerge(); await run(p, 'undo'); await checkMerge(); await shot(p, 'gap-merge-preserved');
  });
  await test('두 번째 가로 보고서 필터는 자기 필드를 검색·선택한다', async p => {
    await options(p, 'over', 2); const before = await def(p), m = await openFilter(p, '채널'); await m.getByRole('searchbox').fill('오프라인'); await m.getByRole('option', { name: '오프라인', exact: true }).click(); await confirm(m).click(); eq((await def(p)).filters.채널, ['오프라인']); eq((await def(p)).filters?.지역, undefined); const s = await snapshot(p, 'second-filter-selected'); eq(s.fields.find(f => f.field === '채널').value, '오프라인'); await run(p, 'undo'); eq(await def(p), before);
  });
  await test('가로 필터 우클릭의 필드 이름 변경은 올바른 필드와 옵션 확인 뒤에도 유지된다', async p => {
    await options(p, 'over', 2); let m = await contextField(p, '채널'); await m.getByRole('menuitem', { name: /^필드 설정/ }).click(); let d = dialog(p, '필드 설정'); await d.getByRole('textbox').fill('판매 경로'); await confirm(d).click(); eq((await def(p)).fieldCaptions.채널, '판매 경로'); await options(p); const state = await snapshot(p, 'caption-preserved'); eq(state.def.fieldCaptions.채널, '판매 경로'); eq(state.fields.find(f => f.field === '채널').label, '판매 경로'); m = await contextField(p, '채널'); await m.getByRole('menuitem', { name: /제거\(E\)/ }).click(); eq((await def(p)).pages, ['지역', '연도']); await run(p, 'undo'); eq((await def(p)).pages, ['지역', '채널', '연도']);
  });
  await test('단일·다중 항목과 모두 선택의 표시가 필터만 있는 경우에도 일치한다', async p => {
    let m = await openFilter(p); await m.getByRole('option', { name: '서울', exact: true }).click(); await confirm(m).click(); eq((await snapshot(p, 'single')).fields[0].value, '서울');
    m = await openFilter(p); await m.getByRole('checkbox', { name: '여러 항목 선택', exact: true }).check(); await m.getByRole('checkbox', { name: '부산', exact: true }).check(); await confirm(m).click(); eq((await snapshot(p, 'multi')).fields[0].value, '(다중 항목)');
    m = await openFilter(p); await m.getByRole('checkbox', { name: '표시된 항목 모두 선택', exact: true }).check(); await confirm(m).click(); eq((await snapshot(p, 'all-multi')).fields[0].value, '(모두)');
  });
  await test('보고서 필터 검색 추가·취소·Undo는 선택과 배치를 보존한다', async p => {
    await options(p, 'over', 2); let m = await openFilter(p); await m.getByRole('option', { name: '서울', exact: true }).click(); await confirm(m).click(); const before = await def(p);
    m = await openFilter(p); await m.getByRole('checkbox', { name: '여러 항목 선택', exact: true }).check(); await m.getByRole('searchbox').fill('부산'); await m.getByRole('checkbox', { name: '필터에 현재 선택한 내용 추가', exact: true }).check(); await cancel(m).click(); eq(await def(p), before);
    m = await openFilter(p); await m.getByRole('checkbox', { name: '여러 항목 선택', exact: true }).check(); await m.getByRole('searchbox').fill('부산'); await m.getByRole('checkbox', { name: '필터에 현재 선택한 내용 추가', exact: true }).check(); await confirm(m).click(); eq((await def(p)).filters.지역.slice().sort(), ['부산', '서울']); await placement(p, [['지역', 0, 0], ['채널', 0, 3], ['연도', 1, 0]]); await run(p, 'undo'); eq(await def(p), before);
  });
  await test('옵션 취소와 잘못된 줄 바꿈 수는 기존 필터 배치를 보존한다', async p => {
    const before = await def(p); let d = await options(p, 'over', 2, false); await cancel(d).click(); eq(await def(p), before);
    for (const invalid of ['-1', '1.5', '256']) { d = await options(p, 'over', invalid, false); await confirm(d).click(); eq(await d.isVisible(), true); eq(await def(p), before); await cancel(d).click(); }
  });
  await test('가로 필터 제거 뒤 잔상을 지우고 한 번 Undo로 복원한다', async p => {
    await options(p, 'over', 0); const before = await def(p), m = await contextField(p, '연도'); await m.getByRole('menuitem', { name: /제거\(E\)/ }).click(); eq((await def(p)).pages, ['지역', '채널']); eq(await p.evaluate(before => [window.tabula.wb().getRaw(window.tabula.si, before.top, before.left + 6), window.tabula.wb().getRaw(window.tabula.si, before.top, before.left + 7)], before), ['', '']); await run(p, 'undo'); eq(await def(p), before);
  });
  await test('XLSX 저장 후 실제 파일 다시 열기는 배치·다중 모드·필터·원본 값을 보존한다', async p => {
    await create(p, ['지역', '채널', '연도'], true); await options(p, 'over', 2); let m = await openFilter(p); await m.getByRole('option', { name: '서울', exact: true }).click(); await confirm(m).click(); m = await openFilter(p); await m.getByRole('checkbox', { name: '여러 항목 선택', exact: true }).check(); await confirm(m).click();
    const before = await snapshot(p, 'before-save');
    await p.evaluate(() => { window.__pivotFile = null; window.showSaveFilePicker = async () => ({ name: '보고서 필터 합성.xlsx', createWritable: async () => ({ write: async value => { const blob = value instanceof Blob ? value : new Blob([value]); window.__pivotFile = [...new Uint8Array(await blob.arrayBuffer())]; }, close: async () => {}, abort: async () => {} }) }); });
    await p.evaluate(() => window.tabula.exportXlsx('보고서 필터 합성', 'xlsx')); await p.waitForFunction(() => window.__pivotFile?.length > 0); const bytes = await p.evaluate(() => window.__pivotFile); await writeFile(out + '/report-filters.xlsx', Buffer.from(bytes));
    await p.locator('#fileInput').setInputFiles({ name: '보고서 필터 합성.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(bytes) }); await p.waitForFunction(() => window.tabula.wb().sheets.some(s => s.pivot?.pageOrder === 'over'));
    await p.evaluate(() => { const t = window.tabula; t.switchSheet(t.wb().sheets.findIndex(s => s.pivot)); const d = t.wb().sheets[t.si].pivot; t.selectCell(d.top, d.left); });
    const after = await snapshot(p, 'after-reopen'); eq(after.def.pages, before.def.pages); eq(after.def.filters, before.def.filters); eq(after.def.pages.map(f=>!!after.def.pageMulti?.[f]), before.def.pages.map(f=>!!before.def.pageMulti?.[f])); eq(after.def.pageOrder, 'over'); eq(after.def.pageWrap, 2); eq(after.fields.map(f => [f.field, f.r, f.c, f.value]), before.fields.map(f => [f.field, f.r, f.c, f.value]));
    eq(await p.evaluate(() => { const w = window.tabula.wb(); return Array.from({ length: 5 }, (_, r) => Array.from({ length: 5 }, (_, c) => w.getValue(0, r, c))); }), [['지역','채널','연도','상품','매출'],['서울','온라인',2025,'A',100],['부산','오프라인',2026,'B',200],['서울','오프라인',2026,'B',300],['대구','온라인',2025,'A',400]]); await shot(p, 'reopened');
  });
  for (const width of [390, 320]) await test(width + 'px 화면에서 보고서 필터 검색·확인과 옵션 취소에 접근한다', async p => {
    await p.setViewportSize({ width, height: 844 }); await p.evaluate(() => { if (!window.tabula.mobile().active) window.tabula.run('mobileWorkMode'); }); if (await p.locator('#pivotPane').isVisible()) await p.locator('#pivotPane .pp-head button').click();
    const m = await openFilter(p); await m.getByRole('searchbox').fill('서울'); await m.getByRole('option', { name: '서울', exact: true }).click();
    for (const n of [m.getByRole('searchbox'), confirm(m), cancel(m)]) { await n.scrollIntoViewIfNeeded(); const box = await n.boundingBox(); ok(box && box.x >= -1 && box.x + box.width <= width + 1 && box.y >= -1 && box.y + box.height <= 845); }
    await shot(p, 'mobile-' + width); await confirm(m).click(); eq((await def(p)).filters.지역, ['서울']); await run(p, 'pivotOptions'); const d = dialog(p, '피벗 테이블 옵션'); await cancel(d).scrollIntoViewIfNeeded(); const box = await cancel(d).boundingBox(); ok(box.x >= -1 && box.x + box.width <= width + 1 && box.y + box.height <= 845); await cancel(d).click();
  });
  if (native) {
    await access(native);
    await test('별도 Excel로 재저장한 합성 파일의 필터 서식은 새 피벗과 일치한다', async p => {
      await create(p, ['지역']); const fresh = await snapshot(p, 'fresh-native-compare'); await p.locator('#fileInput').setInputFiles(native); await p.waitForFunction(() => window.tabula.wb().sheets.some(s => s.pivot?.name === 'ReportFilter')); await p.evaluate(() => { const t = window.tabula; t.switchSheet(t.wb().sheets.findIndex(s => s.pivot)); const d = t.wb().sheets[t.si].pivot; t.selectCell(d.top, d.left); }); const imported = await snapshot(p, 'native-import'); eq(imported.fields.length, 1); ok(['left','start'].includes(imported.fields[0].cssAlign));
      const style = s => ({ fill: s.fill || null, bb: !!s.bb, bbc: s.bbc || null, bold: !!s.bold }); eq(style(imported.fields[0].valueStyle), style(fresh.fields[0].valueStyle)); eq(style(imported.fields[0].labelStyle), style(fresh.fields[0].labelStyle)); await shot(p, 'native-import');
    });
  }
} finally {
  await browser.close(); const result = { url, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, nativeFixture: native || null, pageErrors: errors, remoteWrites: writes, blockedRequests: blocked, results };
  await writeFile(out + '/result.json', JSON.stringify(result, null, 2)); await writeFile(out + '/measurements.json', JSON.stringify(measurements, null, 2)); console.log(JSON.stringify(result)); if (results.some(r => !r.ok) || errors.length || writes.length) process.exitCode = 1;
}
