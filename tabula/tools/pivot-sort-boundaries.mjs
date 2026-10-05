// 합성 피벗 정렬 경계 회귀: 로컬/명시한 공개 읽기, 외부/API 쓰기 차단.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { resolvePivot, computePivot, pivotSourceData } from '../src/pivot.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = process.env.WIXEL_URL || 'http://127.0.0.1:5195/', origin = new URL(url).origin;
const publicCheck = process.env.WIXEL_PIVOT_SORT_PUBLIC === '1' && origin === 'https://wixel-3.wizx.workers.dev';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname) && !publicCheck) throw Error('로컬 또는 명시한 위셀 공개 주소만 검사합니다.');
const out = process.env.WIXEL_PIVOT_SORT_OUT || `D:/Codex/Temp/wixel-pivot-sort-20261005/${engine}`, filter = process.env.WIXEL_PIVOT_SORT_FILTER || '';
await mkdir(out, { recursive: true });
const browserTemp = out + '/browser-temp';await mkdir(browserTemp, { recursive: true });process.env.TEMP = browserTemp;process.env.TMP = browserTemp;
const browser = await pw[engine].launch(), results = [], assets = new Set();let checks = 0;
const eq = (a, b, label) => { assert.deepEqual(a, b, label);checks++; };
const source = [['Region', 'Item', 'Channel', 'Value'], ['East', 'Alpha', 'Direct', 10], ['East', 'Beta', 'Partner', 30], ['North', 'Alpha', 'Direct', 5], ['North', 'Beta', 'Partner', 15], ['West', 'Alpha', 'Direct', 40], ['West', 'Beta', 'Partner', 20]];
function fixture({ nested = false, layout = 'tabular', adjacent = false, columns = false, scoped = false, literalTotal = false, grandCaption } = {}) {
  const w = new Workbook();w.sheets[0].name = 'Source';
  const data = source.map((row, r) => row.map((v, c) => literalTotal && c === 0 && v === 'West' ? '총합계' : scoped && r > 0 && c === 3 ? [70, 0, 40, 0, 20, 100][r - 1] : v));
  data.forEach((row, r) => row.forEach((v, c) => w.setInput(0, r, c, String(v))));
  const si = w.addSheet('Report'), d = { name: 'BoundaryPivot', source: 'Source', range: { r1: 0, c1: 0, r2: 6, c2: 3 }, rows: nested ? ['Region', 'Item'] : ['Region'], cols: columns ? ['Channel'] : [], values: [{ field: 'Value', agg: 'sum', name: 'Revenue' }], layout, subtotals: true, top: 1, left: 1, customListSort: false };
  if (grandCaption !== undefined) d.grandCaption = grandCaption;
  w.sheets[si].pivot = d;
  w.sheets[si].slicers = [{ id: 'channel', caption: 'Channel', source: { kind: 'pivot', field: 'Channel', pivots: [{ sheet: 'Report', name: d.name }] }, x: 650, y: 30, w: 185, h: 150, style: 'SlicerStyleLight1' }];
  w.setInput(si, 15, 1, 'RAW HEADER');w.setInput(si, 16, 1, 'Raw Z');w.setInput(si, 16, 2, '999');w.setInput(si, 17, 1, 'Raw A');w.setInput(si, 17, 2, '-999');
  if (adjacent) for (let r = 1;r <= 5;r++) w.setInput(si, r, 3, r === 1 ? 'Adjacent source' : String(200 - r));
  return new Workbook(readXlsx(writeXlsx(w)).data).serialize();
}
const run = (p, command) => p.evaluate(command => tabula.run(command), command);
const settle = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
async function state(p) { return p.evaluate(() => {
  const w = tabula.wb(), d = w.sheets[1].pivot, a = d.area;
  const grid = [];for (let r = a.r1;r <= a.r2;r++) { const row = [];for (let c = a.c1;c <= a.c2;c++) row.push(w.getValue(1, r, c));grid.push(row); }
  return { def: structuredClone(d), grid, source: [...w.sheets[0].cells].map(([k, v]) => [k, v.raw]), outside: [...w.sheets[1].cells].filter(([k]) => { const [r, c] = k.split(',').map(Number);return r < a.r1 || r > a.r2 || c < a.c1 || c > a.c2; }).map(([k, v]) => [k, v.raw]), undo: w.undoStack.length, redo: w.redoStack.length };
}); }
async function select(p, r, c, range) { await p.evaluate(({ r, c, range }) => range ? tabula.selectRange(range, 'cells', { r, c }) : tabula.selectCell(r, c), { r, c, range }); }
async function history(p, before, after) {
  eq(after.undo, 1, '정렬은 한 번 실행 취소');await run(p, 'undo');await settle(p);let s = await state(p);eq(s.grid, before.grid, 'Undo 실제 피벗 셀 복원');eq(s.def, before.def, 'Undo 피벗 정의 복원');
  await run(p, 'redo');await settle(p);s = await state(p);eq(s.grid, after.grid, 'Redo 실제 피벗 셀 복원');eq(s.def, after.def, 'Redo 피벗 정의 복원');
}
async function roundtrip(p, expected) {
  const w = new Workbook(await p.evaluate(() => tabula.wb().serialize())), back = new Workbook(readXlsx(writeXlsx(w)).data), d = back.sheets[1].pivot;
  const res = resolvePivot(pivotSourceData(back, d), d), grid = computePivot(res, res.def).grid.map(row => row.map(c => c.raw === '' ? null : /^-?\d+(\.\d+)?$/.test(c.raw) ? Number(c.raw) : c.raw.replace(/^'/, '')));
  eq(grid, expected.grid, 'XLSX 저장 후 다시 계산한 행·합계 순서 보존');eq(d.sort, expected.def.sort, 'XLSX 정렬 기준 보존');
}
async function valueSortResult(p, before, dir, field = 'Region', nested = false, layout = 'tabular') {
  const s = await state(p);
  eq(s.grid.at(-1).at(-1), 120, '총합계 값은 마지막 행에 유지');eq(s.grid.at(-1)[0], '총합계', '총합계 라벨은 마지막 행에 유지');
  eq(s.def.sort?.[field]?.dir, dir, '피벗 정의의 정렬 방향 갱신');eq(s.def.sort?.[field]?.by, 0, '선택한 값 필드 기준 정렬');
  if (!nested) eq(s.grid.slice(1, -1).map(row => row[0]), dir === 'asc' ? ['North', 'East', 'West'] : ['West', 'East', 'North'], '항목과 값이 함께 정렬');
  else if (layout === 'tabular') {
    eq(s.grid.slice(1, -1), dir === 'desc' ? [['East', 'Beta', 30], [null, 'Alpha', 10], ['East 요약', null, 40], ['North', 'Beta', 15], [null, 'Alpha', 5], ['North 요약', null, 20], ['West', 'Alpha', 40], [null, 'Beta', 20], ['West 요약', null, 60]] : [['East', 'Alpha', 10], [null, 'Beta', 30], ['East 요약', null, 40], ['North', 'Alpha', 5], [null, 'Beta', 15], ['North 요약', null, 20], ['West', 'Beta', 20], [null, 'Alpha', 40], ['West 요약', null, 60]], '각 상위 그룹 내부만 값 정렬하고 부분합은 그룹 아래에 유지');
  } else eq(s.grid.slice(1, -1).map(row => row.at(-1)), dir === 'desc' ? [40, 30, 10, 20, 15, 5, 60, 40, 20] : [40, 10, 30, 20, 5, 15, 60, 20, 40], '상위 그룹 합계와 자식 항목 경계 유지');
  await history(p, before, s);await roundtrip(p, s);return s;
}
async function chooseDialogSort(p, dir, by = '0') {
  const d = p.getByRole('dialog', { name: /^정렬(?:\(|$)/ });await d.waitFor();
  if (await d.locator('input[name="pvsort"]').count()) { await d.locator(`input[name="pvsort"][value="${dir}"]`).check();await d.locator('select:not(:disabled)').selectOption(by); }
  else await d.locator('.sort-row:not(.sort-head) select').nth(2).selectOption(dir);
  await d.getByRole('button', { name: '확인', exact: true }).click();await settle(p);
}
async function contextSort(p, dir) {
  await p.locator('#cellEditor').focus();await p.keyboard.press('Shift+F10');
  const menu = p.locator('#menuLayer>.menu[data-context-kind="pivot"]');await menu.waitFor();await menu.getByRole('menuitem', { name: /^정렬/ }).hover();
  await p.locator('#menuLayer .menu').getByRole('menuitem', { name: dir === 'asc' ? '오름차순 정렬(S)' : '내림차순 정렬(O)', exact: true }).click();await settle(p);
}
async function test(name, options, action) {
  if (filter && !name.includes(filter)) return;
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, serviceWorkers: 'block' }), p = await context.newPage(), errors = [], writes = [], start = checks;let before, after;
  p.setDefaultTimeout(15000);p.on('pageerror', e => errors.push(e.message));p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true;window.WIXEL_SKIP_START = true;localStorage.setItem('wixel.mobile-work.v1', 'off'); });
  await context.route('**/*', route => { const req = route.request(), u = new URL(req.url());if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.pathname);return route.abort(); }if (u.origin !== origin || /^\/api(?:\/|$)/.test(u.pathname)) return route.abort();return route.continue(); });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });await p.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
    for (const a of await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src')))) assets.add(a);
    await p.evaluate(data => { const t = tabula, w = t.wb();w.restore(data);t.switchSheet(1);t.selectCell(2, 2);t.run('pivotRefresh');w.undoStack = [];w.redoStack = []; }, fixture(options));await settle(p);
    before = await state(p);await p.screenshot({ path: `${out}/${name}-before.png` });await action(p, before);after = await state(p);
    eq(after.source, before.source, '원본 셀 변경 없음');eq(after.outside, before.outside, '피벗 밖 원본/메모 셀 변경 없음');eq(errors, [], '브라우저 오류 없음');eq(writes, [], 'API·외부 쓰기 없음');
    await p.screenshot({ path: `${out}/${name}-after.png` });results.push({ name, ok: true, checks: checks - start, before, after });console.log('OK ' + name);
  } catch (error) { after = await state(p).catch(() => null);await p.screenshot({ path: `${out}/${name}-failure.png` }).catch(() => {});results.push({ name, ok: false, error: error.stack, errors, writes, checks: checks - start, before, after });console.error('NG ' + name + ': ' + error.message); }
  finally { await context.close(); }
}
try {
  for (const dir of ['asc', 'desc']) await test('quick-' + dir, {}, async (p, b) => { await run(p, dir === 'asc' ? 'sortAsc' : 'sortDesc');await settle(p);await valueSortResult(p, b, dir); });
  await test('ribbon-desc', {}, async (p, b) => { await p.locator('[data-ribbon-tab="data"]').click();await p.locator('[data-ribbon-command="sortDesc"]').click();await settle(p);await valueSortResult(p, b, 'desc'); });
  await test('dialog-desc', {}, async (p, b) => { await run(p, 'sortDialog');await chooseDialogSort(p, 'desc');await valueSortResult(p, b, 'desc'); });
  await test('value-column-selection-desc', {}, async (p, b) => { await select(p, 2, 2, { r1: 1, c1: 2, r2: 5, c2: 2 });await run(p, 'sortDesc');await settle(p);await valueSortResult(p, b, 'desc'); });
  await test('adjacent-raw-auto-region-desc', { adjacent: true }, async (p, b) => { await run(p, 'sortDesc');await settle(p);await valueSortResult(p, b, 'desc'); });
  await test('context-value-desc', {}, async (p, b) => { await contextSort(p, 'desc');await valueSortResult(p, b, 'desc'); });
  await test('header-filter-label-desc', {}, async (p, b) => {
    await p.locator('.pbtn[data-k="rows"][data-f="Region"]').first().click();await p.locator('.pivot-filter-menu').getByRole('menuitem', { name: '텍스트 내림차순 정렬', exact: true }).click();await settle(p);
    const s = await state(p);eq(s.grid.slice(1, -1).map(row => row[0]), ['West', 'North', 'East'], '머리글 필터 메뉴의 레이블 정렬');eq(s.grid.at(-1), ['총합계', 120], '머리글 정렬의 총합계 고정');eq(s.def.sort.Region, { dir: 'desc' }, '레이블 기준만 저장');await history(p, b, s);await roundtrip(p, s);
  });
  for (const layout of ['tabular', 'compact', 'outline']) for (const dir of ['asc', 'desc']) await test('nested-' + layout + '-' + dir, { nested: true, layout }, async (p, b) => {
    await select(p, layout === 'tabular' ? 2 : 3, layout === 'compact' ? 2 : 3);await run(p, dir === 'asc' ? 'sortAsc' : 'sortDesc');await settle(p);await valueSortResult(p, b, dir, 'Item', true, layout);
  });
  for (const command of ['sortAsc', 'sortDesc', 'sortDialog']) await test('mixed-raw-pivot-' + command, {}, async (p, b) => {
    await select(p, 16, 2, { r1: 1, c1: 1, r2: 17, c2: 2 });await run(p, command);await settle(p);
    if (command === 'sortDialog' && await p.locator('.sort-levels').count()) await chooseDialogSort(p, 'desc');
    const s = await state(p);eq(s.grid, b.grid, '피벗과 일반 셀 혼합 범위는 피벗 셀 재배열 금지');eq(s.outside, b.outside, '혼합 범위의 일반 셀도 재배열 금지');eq(s.undo, 0, '거부한 혼합 정렬의 실행 취소 기록 없음');
  });
  for (const path of ['quick', 'ribbon', 'dialog', 'context']) await test('opposite-axis-scope-' + path, { columns: true, scoped: true }, async (p, b) => {
    await select(p, 3, 2);
    if (path === 'context') await contextSort(p, 'desc');
    else if (path === 'dialog') { await run(p, 'sortDialog');await chooseDialogSort(p, 'desc'); }
    else if (path === 'ribbon') { await p.locator('[data-ribbon-tab="data"]').click();await p.locator('[data-ribbon-command="sortDesc"]').click(); }
    else await run(p, 'sortDesc');
    await settle(p);const s = await state(p);
    eq(s.grid.slice(2, -1), [['East', 70, 0, 70], ['North', 40, 0, 40], ['West', 20, 100, 120]], '선택한 Direct 값은 전체 합계와 다른 행 순서로 정렬');
    eq(s.def.sort.Region, { dir: 'desc', by: 0, at: [['Channel', 'Direct']] }, '반대 축의 실제 항목 경로 저장');eq(s.grid.at(-1), ['총합계', 130, 100, 230], '열별 합계 및 전체 합계는 마지막 행 유지');await history(p, b, s);await roundtrip(p, s);
  });
  for (const path of ['quick', 'context']) await test('grand-row-column-sort-' + path, { columns: true, scoped: true }, async (p, b) => {
    await select(p, 6, 2);if (path === 'context') await contextSort(p, 'asc');else await run(p, 'sortAsc');await settle(p);const s = await state(p);
    eq(s.grid, [['Revenue', 'Channel', null, null], ['Region', 'Partner', 'Direct', '총합계'], ['East', 0, 70, 70], ['North', 0, 40, 40], ['West', 100, 20, 120], ['총합계', 100, 130, 230]], '합계 행의 값 정렬은 열 항목만 이동하고 끝 합계 열 유지');
    eq(s.def.sort.Channel, { dir: 'asc', by: 0 }, '열 필드를 전체 행 합계로 정렬');await history(p, b, s);await roundtrip(p, s);
  });
  await test('grand-corner-no-sort', { columns: true, scoped: true }, async (p, b) => {
    await select(p, 6, 4);await run(p, 'sortDesc');await settle(p);const s = await state(p);eq(s.grid, b.grid, '행·열 합계 교차점은 정렬 대상 항목 아님');eq(s.def, b.def, '합계 교차점 정의 유지');eq(s.undo, 0, '합계 교차점에 실행 취소 기록 없음');
  });
  for (const custom of [false, true]) await test('literal-total-item-' + (custom ? 'custom-caption' : 'default-caption'), { literalTotal: true, ...(custom ? { grandCaption: '전체 매출' } : {}) }, async (p, b) => {
    await run(p, 'sortDesc');await settle(p);const s = await state(p);eq(s.grid, [['Region', 'Revenue'], ['총합계', 60], ['East', 40], ['North', 20], [custom ? '전체 매출' : '총합계', 120]], '총합계라는 실제 항목도 값으로 정렬하며 실제 총합계는 역할로 구분');await history(p, b, s);await roundtrip(p, s);
  });
  for (const permit of [true, false]) for (const command of ['sortAsc', 'sortDesc', 'sortDialog']) await test('protected-pivot-' + (permit ? 'allowed-' : 'denied-') + command, {}, async (p, b) => {
    await p.evaluate(permit => { const w = tabula.wb();w.sheets[1].protect = { on: true, allow: { selectLocked: true, selectUnlocked: true, pivotTables: permit, sort: !permit } }; }, permit);
    await run(p, command);if (permit && command === 'sortDialog') await chooseDialogSort(p, 'desc');await settle(p);
    if (permit) await valueSortResult(p, b, command === 'sortAsc' ? 'asc' : 'desc');
    else { const s = await state(p);eq(s.grid, b.grid, '일반 정렬만 허용해도 피벗 권한 없으면 변경 금지');eq(s.def, b.def, '보호된 피벗 정의 유지');eq(s.undo, 0, '거부된 정렬 실행 취소 없음'); }
  });
  for (const mode of ['protected', 'marked-final']) await test('ordinary-dialog-late-' + mode, {}, async (p, b) => {
    await select(p, 16, 2, { r1: 15, c1: 1, r2: 17, c2: 2 });await run(p, 'sortDialog');const d = p.getByRole('dialog', { name: '정렬', exact: true });await d.waitFor();
    await p.evaluate(mode => { const w = tabula.wb();if (mode === 'protected') w.sheets[1].protect = { on: true, allow: { selectLocked: true, selectUnlocked: true, sort: false } };else w.props = { ...w.props, markedFinal: true }; }, mode);
    await chooseDialogSort(p, 'desc');const s = await state(p);eq(s.grid, b.grid, '늦은 권한 변경 후 피벗 셀 보존');eq(s.outside, b.outside, '열린 일반 정렬 창도 권한 재확인 후 원본 변경 거부');eq(s.undo, 0, '권한 변경 뒤 정렬 실행 취소 기록 없음');
  });
  await test('slicer-refresh-sorted-report', {}, async (p, b) => {
    await run(p, 'sortDesc');await settle(p);await valueSortResult(p, b, 'desc');
    const sl = p.locator('.pane-br .obj.slicer').first();await sl.locator('.sl-item').filter({ hasText: /^Direct$/ }).click();await settle(p);
    let s = await state(p);eq(s.grid, [['Region', 'Revenue'], ['West', 40], ['East', 10], ['North', 5], ['총합계', 55]], '슬라이서 선택 후 정렬 정의로 재계산');eq(s.def.sort.Region, { dir: 'desc', by: 0 }, '슬라이서가 정렬 설정 보존');
    await run(p, 'pivotRefresh');await settle(p);eq((await state(p)).grid, s.grid, '명시 새로 고침 후 정렬·총합계 일치');
    await sl.locator('.sl-clear').click();await settle(p);s = await state(p);eq(s.grid, [['Region', 'Revenue'], ['West', 60], ['East', 40], ['North', 20], ['총합계', 120]], '슬라이서 해제 후 정렬된 전체 보고서 복원');await roundtrip(p, s);
  });
} finally {
  await browser.close();const result = { engine, url, publicCheck, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
  await writeFile(`${out}/result.json`, JSON.stringify(result, null, 2));console.log(JSON.stringify({ engine, cases: result.cases, passed: result.passed, checks, assets: result.assets, out }));if (result.passed !== result.cases || assets.size !== 1) process.exitCode = 1;
}
