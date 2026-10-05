// Excel 단축키 회귀: 합성 문서만 사용하고 API·외부·쓰기 요청을 차단합니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { pathToFileURL } from 'node:url';
const engine = process.env.WIXEL_BROWSER || 'chromium';
const url = process.env.WIXEL_URL || 'http://localhost:5195/';
const address = new URL(url), origin = address.origin;
const publicRun = process.env.WIXEL_SHORTCUT_PUBLIC === '1' && origin === 'https://wixel-3.wizx.workers.dev';
assert.ok(['localhost', '127.0.0.1', '[::1]'].includes(address.hostname) || publicRun, '로컬 서버 또는 승인된 합성 공개 검사만 허용');
const out = process.env.WIXEL_SHORTCUT_OUT || `D:/Codex/Temp/wixel-excel-shortcuts-20261005/${engine}`;
assert.match(out, /^D:[/\\]/i, '검사 결과는 D:에 저장');
await mkdir(out, { recursive: true });
const filter = process.env.WIXEL_SHORTCUT_FILTER || '';
const modulePath = process.env.PLAYWRIGHT_MODULE;
const pw = await import(modulePath ? (/^[A-Za-z]:[/\\]/.test(modulePath) ? pathToFileURL(modulePath).href : modulePath) : 'playwright');
const browser = await pw[engine].launch();
const results = []; let checks = 0;
const eq = (actual, expected, message) => { checks++; assert.deepEqual(actual, expected, message); };
const ok = (actual, message) => { checks++; assert.ok(actual, message); };
const seq = async (p, letters, { held = false, start = true } = {}) => {
  if (start) { if (held) await p.keyboard.down('Alt'); else await p.keyboard.press('Alt'); }
  for (const key of letters) await p.keyboard.press(key);
  if (held) await p.keyboard.up('Alt');
};
const run = (p, command) => p.evaluate(command => tabula.run(command), command);
const state = p => p.evaluate(() => ({
  filter: tabula.wb().sheets[0].filter ?? null,
  raw: tabula.wb().getRaw(0, 1, 1), editor: document.getElementById('cellEditor').value,
  editing: !document.getElementById('cellEditor').classList.contains('idle'),
  keytip: document.body.dataset.keytipSequence ?? null,
  dialogs: [...document.querySelectorAll('.dialog')].map(n => n.getAttribute('aria-label') || n.querySelector('.dialog-title')?.textContent || n.textContent.slice(0, 30)),
  undo: tabula.wb().undoStack.length, sheets: tabula.wb().sheets.length,
  charts: tabula.wb().sheets.reduce((n, s) => n + (s.charts?.length || 0), 0),
}));
async function fixture(p, table = false) {
  await p.evaluate(table => {
    const t = tabula, w = t.wb(), cells = {};
    [['분류', '수량', '금액'], ['A', '10', '100'], ['B', '20', '200'], ['C', '30', '300']].forEach((row, r) => row.forEach((raw, c) => cells[r + ',' + c] = { raw }));
    const sheets = [{ name: '합성 단축키', cells, shapes: [{ id: 'test-shape', kind: 'rect', text: '검사 도형', x: 380, y: 100, w: 200, h: 110 }] }];
    if (table) {
      sheets[0].tables = [{ id: 'test-table', name: '검사표', r1: 0, c1: 0, r2: 3, c2: 2, header: true, style: 'TableStyleMedium2', filter: { criteria: { 0: ['A'] }, hidden: { 2: true, 3: true } } }];
      sheets[0].slicers = [{ id: 'test-slicer', name: '검사 슬라이서', caption: '분류', source: { kind: 'table', table: 'test-table', column: 0 }, x: 620, y: 100, w: 180, h: 220 }];
    }
    w.restore({ sheets }); w.undoStack = []; w.redoStack = [];
    t.gv().layout(); t.gv().renderAll(); t.selectCell(1, 1);
  }, table);
  await p.locator('#cellEditor').focus();
}
async function test(name, body, { table = false } = {}) {
  if (filter && !filter.split('|').some(part => name.includes(part))) return;
  const start = checks, context = await browser.newContext({ viewport: { width: 1440, height: 960 }, serviceWorkers: 'block' });
  const p = await context.newPage(), errors = [], writes = [];
  p.setDefaultTimeout(6000); p.on('pageerror', e => errors.push(e.message));
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await context.route('**/*', route => {
    const req = route.request(), target = new URL(req.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.url()); return route.abort(); }
    return target.origin === origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await p.waitForFunction(() => !!window.tabula?.gv()); await fixture(p, table);
    await body(p); eq(errors, [], '브라우저 오류 없음'); eq(writes, [], '원격 쓰기 없음');
    results.push({ name, ok: true, checks: checks - start }); console.log('OK ' + name);
  } catch (error) {
    results.push({ name, ok: false, checks: checks - start, error: error.stack, state: await state(p).catch(() => null), errors, writes });
    console.error('NG ' + name + ': ' + error.message);
    await p.screenshot({ path: `${out}/failure-${results.length}.png` }).catch(() => {});
  } finally { await context.close(); }
}
try {
  await test('legacy-dff-filter-toggle-undo', async p => {
    await seq(p, 'dff'); const first = await state(p);
    eq(!!first.filter, true, 'Alt D F F 필터 생성'); eq(first.editing, false, '키 입력 오염 없음'); eq(first.raw, '10', '활성 셀 유지');
    await p.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setSheetProp(0, 'filter', { ...w.sheets[0].filter, criteria: { 0: ['A'] }, hidden: { 2: true, 3: true } })); });
    const filtered = (await state(p)).filter; await seq(p, 'dff'); eq((await state(p)).filter, null, '기존 필터 해제');
    await p.keyboard.press('Control+z'); eq((await state(p)).filter, filtered, 'Undo는 조건과 숨긴 행 복원');
  });
  await test('modern-filter-and-table-button-state', async p => {
    await seq(p, 'at'); eq(!!(await state(p)).filter, true, 'Alt A T 필터 생성');
    await p.keyboard.press('Control+Shift+l'); eq((await state(p)).filter, null, 'Ctrl Shift L 동일 필터 해제');
    await fixture(p, true); const before = await p.evaluate(() => structuredClone(tabula.wb().sheets[0].tables[0].filter));
    await run(p, 'tblFilter'); const hiddenButtons = await p.evaluate(() => tabula.wb().sheets[0].tables[0].filter);
    eq(hiddenButtons.criteria, before.criteria, '디자인 필터 단추는 조건 보존'); eq(hiddenButtons.hidden, before.hidden, '디자인 필터 단추는 숨긴 행 보존');
    await p.locator('#cellEditor').focus(); await seq(p, 'dff');
    eq(await p.evaluate(() => tabula.wb().sheets[0].tables[0].filter), null, '레거시 자동필터는 표 조건도 해제');
  });
  await test('legacy-clear-advanced-and-modern-reapply', async p => {
    await seq(p, 'dff');
    await p.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setSheetProp(0, 'filter', { ...w.sheets[0].filter, criteria: { 0: ['A'] }, hidden: { 2: true, 3: true } })); });
    await p.evaluate(() => tabula.wb().transact(() => tabula.wb().setInput(0, 2, 0, 'A')));
    await p.keyboard.press('Control+Alt+l'); eq((await state(p)).filter.hidden, { 3: true }, '다시 적용은 변경된 원본으로 숨긴 행 갱신');
    const before = (await state(p)).filter; await seq(p, 'dfs');
    eq((await state(p)).filter.criteria, {}, 'Alt D F S는 조건만 지움'); eq((await state(p)).filter.hidden, {}, '모두 표시');
    await p.keyboard.press('Control+z'); eq((await state(p)).filter, before, '지우기 Undo는 이전 조건 복원');
    await seq(p, 'dfa'); await p.getByRole('dialog', { name: '고급 필터', exact: true }).waitFor();
    eq((await state(p)).filter, before, '고급 필터 창을 여는 것만으로 조건 변경 없음');
  });
  await test('legacy-alt-held-korean-and-delayed-ime', async p => {
    await seq(p, 'dff', { held: true }); eq(!!(await state(p)).filter, true, 'Alt를 계속 누른 DFF');
    await seq(p, '', {});
    for (const [key, code] of [['ㅇ', 'KeyD'], ['ㄹ', 'KeyF'], ['ㄹ', 'KeyF']]) await p.locator('#cellEditor').dispatchEvent('keydown', { key, code, bubbles: true, cancelable: true });
    eq((await state(p)).filter, null, '한글 물리 키 DFF');
    await p.locator('#cellEditor').evaluate(e => { e.value = 'ㄹ'; e.dispatchEvent(new CompositionEvent('compositionend', { data: 'ㄹ', bubbles: true })); e.dispatchEvent(new InputEvent('input', { data: 'ㄹ', inputType: 'insertCompositionText', bubbles: true })); });
    eq((await state(p)).editor, '', '종료 뒤 지연된 조합 입력 차단'); eq((await state(p)).raw, '10', '원래 셀 보존');
    await p.keyboard.type('next'); await p.keyboard.press('Enter'); eq((await state(p)).raw, 'next', '다음 정상 입력 보존');
  });
  await test('legacy-backspace-escape-invalid-key', async p => {
    await seq(p, 'df'); await p.keyboard.press('Backspace'); await seq(p, 'ff', { start: false });
    eq(!!(await state(p)).filter, true, '레거시 이전 단계 후 재입력');
    await seq(p, 'd'); await p.keyboard.press('x'); eq((await state(p)).editing, false, '잘못된 레거시 키가 셀에 새지 않음');
    for (let i = 0; i < 4; i++) await p.keyboard.press('Escape');
    eq((await state(p)).keytip, null, 'Esc로 레거시 탐색 종료'); eq(!!(await state(p)).filter, true, '취소는 필터 유지');
    await p.keyboard.type('normal'); await p.keyboard.press('Enter'); eq((await state(p)).raw, 'normal', '취소 뒤 정상 입력');
  });
  await test('end-mode-navigation-extension-and-cancel', async p => {
    await p.keyboard.press('End'); await p.keyboard.press('ArrowDown');
    eq(await p.evaluate(() => tabula.active), { r: 3, c: 1 }, 'End 다음 방향키는 데이터 끝으로 이동');
    await p.keyboard.press('ArrowUp'); eq(await p.evaluate(() => tabula.active), { r: 2, c: 1 }, '한 번 이동 후 끝 모드 해제');
    await p.evaluate(() => tabula.selectCell(1, 1)); await p.keyboard.press('Shift+End'); await p.keyboard.press('Shift+ArrowDown');
    eq(await p.evaluate(() => tabula.sel), { r1: 1, c1: 1, r2: 3, c2: 1 }, 'Shift End 방향키 선택 확장');
    await p.evaluate(() => tabula.selectCell(1, 1)); await p.keyboard.press('End'); await p.keyboard.press('Escape'); await p.keyboard.press('ArrowDown');
    eq(await p.evaluate(() => tabula.active), { r: 2, c: 1 }, 'Esc는 끝 모드 취소');
    await p.keyboard.press('End'); await p.locator('.c[data-r="1"][data-c="1"]').first().click(); await p.keyboard.press('ArrowDown');
    eq(await p.evaluate(() => tabula.active), { r: 2, c: 1 }, '포인터 선택은 끝 모드 취소');
  });
  await test('cell-and-formula-editing-guards', async p => {
    for (const target of ['#cellEditor', '#formulaInput']) {
      await p.keyboard.press('F2'); if (target === '#formulaInput') await p.locator(target).focus();
      await seq(p, 'dff'); eq((await state(p)).filter, null, '편집 중 자동필터 명령 실행 안 함');
      eq((await state(p)).keytip, null, '편집 중 리본 탐색 시작 안 함'); await p.keyboard.press('Escape');
    }
    eq((await state(p)).raw, '10', '편집 취소 후 원래 셀');
  });
  await test('dialog-menu-input-isolation', async p => {
    await p.keyboard.press('Control+g'); const dialog = p.getByRole('dialog').last(); await dialog.waitFor();
    const before = await state(p); await dialog.locator('input').first().focus(); await p.keyboard.press('Control+Shift+g');
    eq((await state(p)).dialogs, before.dialogs, '대화상자 Ctrl Shift G는 새 전역 창을 열지 않음');
    await p.keyboard.press('Escape'); await seq(p, 'at'); await p.evaluate(() => tabula.selectCell(0, 0)); await p.locator('#cellEditor').focus(); await p.keyboard.press('Alt+ArrowDown');
    const menu = p.getByRole('menu').first(); await menu.waitFor(); const search = menu.locator('input[type=search],input[type=text]').first(); await search.fill('A');
    await search.press('Control+a'); eq((await state(p)).raw, '10', '메뉴 검색 선택은 셀 데이터 보존');
    eq(!!(await state(p)).filter, true, '메뉴 검색 키는 필터 유지'); eq(await search.inputValue(), 'A', '메뉴 검색 내용 보존');
  });
  await test('cell-context-c-only-copies', async p => {
    await p.keyboard.press('Shift+F10'); await p.getByRole('menu').first().waitFor();
    await p.keyboard.press('c'); eq((await state(p)).raw, '10', 'C는 내용 삭제 금지');
    eq(await p.getByRole('menu').count(), 0, 'C 한 번으로 복사 메뉴 실행');
    await p.evaluate(() => tabula.selectCell(6, 1)); await p.locator('#cellEditor').focus(); await p.keyboard.press('Enter');
    eq(await p.evaluate(() => tabula.wb().getRaw(0, 6, 1)), '10', 'C가 실제 셀 복사를 수행');
  });
  await test('alt-shift-f1-inserts-sheet', async p => {
    await p.keyboard.press('Alt+Shift+F1'); eq((await state(p)).sheets, 2, '새 시트 삽입'); eq((await state(p)).charts, 0, '차트 삽입과 구분');
  });
  await test('ctrl-shift-f1-fullscreen', async p => {
    await p.evaluate(() => { document.documentElement.requestFullscreen = async () => { throw new Error('합성 전체 화면 fallback'); }; });
    await p.keyboard.press('Control+Shift+F1'); eq(await p.locator('body').evaluate(e => e.classList.contains('wixel-fullscreen')), true, '전체 화면 전환');
    eq(await p.locator('#ribbon').evaluate(e => e.classList.contains('collapsed')), false, '리본 접기로 잘못 연결 안 됨');
  });
  await test('ctrl-shift-g-workbook-statistics', async p => {
    await p.keyboard.press('Control+Shift+g'); await p.getByRole('dialog', { name: '통합 문서 통계', exact: true }).waitFor();
    eq(await p.getByRole('dialog').count(), 1, '통계 창 한 개');
  });
  await test('ctrl-shift-s-saveas-and-ctrl-f12-open', async p => {
    await p.keyboard.press('Control+Shift+s'); await p.getByRole('dialog', { name: '다른 이름으로 저장', exact: true }).waitFor();
    eq((await state(p)).raw, '10', '다른 이름으로 저장 창은 문서 값 보존'); await p.keyboard.press('Escape');
    await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+F12');
    await p.locator('.backstage').waitFor(); eq(await p.locator('.backstage').getByText('하던 일을 이어가세요', { exact: true }).count(), 1, 'Ctrl F12 문서 열기 화면');
  });
  await test('ribbon-focus-saveas-and-korean-physical-key', async p => {
    const button = p.locator('[data-ribbon-tab="home"]'); await button.focus(); await p.keyboard.press('Control+Shift+s');
    await p.getByRole('dialog', { name: '다른 이름으로 저장', exact: true }).waitFor();
    eq((await state(p)).raw, '10', '리본 포커스 Ctrl Shift S는 다른 이름으로 저장'); await p.keyboard.press('Escape');
    await button.focus(); await button.dispatchEvent('keydown', { key: 'ㄴ', code: 'KeyS', ctrlKey: true, shiftKey: true, bubbles: true, cancelable: true });
    await p.getByRole('dialog', { name: '다른 이름으로 저장', exact: true }).waitFor(); eq((await state(p)).raw, '10', '한글 자판 리본 포커스 저장 단축키');
  });
  await test('ctrl-alt-zoom-without-browser-edit', async p => {
    const z = await p.evaluate(() => tabula.gv().z); await p.keyboard.press('Control+Alt+=');
    ok((await p.evaluate(() => tabula.gv().z)) > z, 'Ctrl Alt + 확대'); await p.keyboard.press('Control+Alt+-');
    eq(await p.evaluate(() => tabula.gv().z), z, 'Ctrl Alt - 원래 배율'); eq((await state(p)).raw, '10', '배율은 셀 값 보존');
  });
  await test('unsupported-modified-keys-do-not-run-base', async p => {
    const before = await p.evaluate(() => tabula.wb().styleAt(0, 1, 1));
    await p.keyboard.press('Control+Shift+b'); eq(await p.evaluate(() => tabula.wb().styleAt(0, 1, 1)), before, '미등록 Ctrl Shift B가 Ctrl B로 오동작 안 함');
    await p.keyboard.press('Control+Shift+F2'); eq(await p.getByRole('dialog').count(), 0, '미지원 댓글 단축키가 인쇄 창을 열지 않음');
    eq(await p.locator('.print-preview').count(), 0, '인쇄 미리 보기 미실행');
  });
  await test('shape-ctrl1-context', async p => {
    await p.locator('.obj[data-id=test-shape]').click({ position: { x: 3, y: 3 } }); await p.keyboard.press('Control+1');
    await p.locator('.shape-format-pane').waitFor(); eq(await p.getByRole('dialog', { name: '셀 서식', exact: true }).count(), 0, '도형 Ctrl1은 셀 서식 아님');
    eq((await state(p)).raw, '10', '도형 서식은 셀 값 보존');
  });
  await test('slicer-ctrl1-context', async p => {
    await p.locator('.obj[data-id=test-slicer]').click({ position: { x: 3, y: 3 } }); await p.keyboard.press('Control+1');
    await p.getByRole('dialog', { name: '슬라이서 크기 및 속성', exact: true }).waitFor();
    eq(await p.getByRole('dialog', { name: '셀 서식', exact: true }).count(), 0, '슬라이서 Ctrl1은 셀 서식 아님');
  }, { table: true });
} finally {
  await browser.close();
  const summary = { engine, url, cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
  await writeFile(`${out}/result.json`, JSON.stringify(summary, null, 2)); console.log(JSON.stringify({ ...summary, results: undefined, out }));
  if (!results.length || results.some(r => !r.ok)) process.exitCode = 1;
}