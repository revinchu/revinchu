// 셀 우클릭 실제 UI 회귀. 합성 문서·격리 브라우저·메모리 클립보드만 사용한다.
// WIXEL_URL=소스/번들, WIXEL_CELL_CONTEXT_FILTER=검사 이름 일부(|로 여러 조건).
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const browser = await chromium.launch(), results = [];
let checks = 0;
const eq = (a, b, message) => { assert.deepEqual(a, b, message); checks++; };
const ok = (v, message) => { assert.ok(v, message); checks++; };
const menu = p => p.locator('#menuLayer > .menu[data-level="0"]');
const sub = p => p.locator('#menuLayer > .menu[data-level="1"]');
const item = (p, name) => menu(p).getByRole('menuitem', { name });
const dialog = (p, name) => p.getByRole('dialog', { name, exact: true });
const run = (p, command) => p.evaluate(command => window.tabula.run(command), command);
const state = p => p.evaluate(() => ({ data: window.tabula.wb().serialize(), undo: window.tabula.wb().undoStack.length }));
const selection = p => p.evaluate(() => ({ sel: window.tabula.sel, active: window.tabula.active }));
const rows = p => p.evaluate(() => [0, 1, 2, 3].map(r => [0, 1, 2].map(c => window.tabula.wb().getValue(0, r, c))));
const fixture = () => ({ sheets: [{ name: '셀 메뉴 합성', cells: {
  '0,0': { raw: '이름', style: { bold: true } }, '0,1': { raw: '수량' }, '0,2': { raw: '분류' },
  '1,0': { raw: '가' }, '1,1': { raw: '30' }, '1,2': { raw: '승인', style: { fill: '#ffeecc', color: '#cc0000' } },
  '2,0': { raw: '나' }, '2,1': { raw: '10' }, '2,2': { raw: '검토' },
  '3,0': { raw: '다' }, '3,1': { raw: '20' }, '3,2': { raw: '승인', style: { fill: '#ffeecc', color: '#cc0000' } },
  '5,0': { raw: '=SUM(B2:B4)' }, '5,4': { raw: '범위 밖 보존' },
} }, { name: '다른 시트', cells: { '0,0': { raw: '대상' } } }] });
async function rightClick(p, r = 1, c = 1) {
  const point = await p.evaluate(({ r, c }) => {
    const g = window.tabula.gv(); g.ensureVisible(r, c); g.renderAll();
    const b = g.clientRect({ r1: r, c1: c, r2: r, c2: c }); return { x: (b.left + b.right) / 2, y: (b.top + b.bottom) / 2 };
  }, { r, c });
  await p.mouse.click(point.x, point.y, { button: 'right' }); await menu(p).waitFor();
}
async function submenu(p, name) {
  if (await p.evaluate(() => document.body.classList.contains('mobile-work-mode'))) await item(p, name).tap();
  else await item(p, name).click();
  await sub(p).waitFor(); return sub(p);
}
async function undo(p, before) { await run(p, 'undo'); eq(await state(p), before, '한 번 Undo로 원래 문서/이력 복원'); }
async function test(name, fn, { readonly = false, viewport = { width: 1440, height: 1000 }, mobile = false } = {}) {
  if (process.env.WIXEL_CELL_CONTEXT_FILTER && !process.env.WIXEL_CELL_CONTEXT_FILTER.split('|').some(part => name.includes(part))) return;
  const context = await browser.newContext({ viewport, isMobile: mobile, hasTouch: mobile }), p = await context.newPage();
  const errors = [], writes = [], blocked = []; const start = checks;
  p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => {
    const request = route.request(), target = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.url()); return route.abort(); }
    if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) { blocked.push(request.url()); return route.abort(); }
    return route.continue();
  });
  try {
    await p.addInitScript(() => {
      window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.__cellClipboard = '';
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
        writeText: async text => { window.__cellClipboard = text; }, readText: async () => window.__cellClipboard,
        write: async () => {}, read: async () => [],
      } });
    });
    const target = readonly ? url.split('#')[0] + '#view=' + gzipSync(JSON.stringify({ docName: '공개 합성', workbook: fixture(), view: { headers: true } })).toString('base64url') : url;
    await p.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => !!window.tabula?.wb());
    if (readonly) await p.locator('.view-bar').waitFor();
    else await p.evaluate(data => { const t = window.tabula, w = t.wb(); w.restore(data); t.switchSheet(0); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; }, fixture());
    await p.evaluate(() => window.tabula.selectCell(1, 1));
    await fn(p); eq(errors, [], '페이지 오류 없음'); eq(writes, [], '원격 쓰기 없음');
    results.push({ name, ok: true, checks: checks - start, blockedRequests: blocked.length }); console.log(`OK ${name} (${checks - start})`);
  } catch (e) { const ui = await p.evaluate(() => ({ active: document.activeElement?.outerHTML?.slice(0, 350), toast: document.querySelector('#toast')?.textContent, dialogs: [...document.querySelectorAll('.dialog')].map(n => n.innerText), menus: [...document.querySelectorAll('#menuLayer > .menu')].map(n => ({ level: n.dataset.level, text: n.innerText, style: n.getAttribute('style'), hidden: n.hidden })) })).catch(() => null); results.push({ name, ok: false, checks: checks - start, error: e.message, errors, writes, ui }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
try {
  await test('실제 우클릭은 기존 다중 선택을 유지하고 범위 밖은 새 셀만 선택', async p => {
    await p.evaluate(() => window.tabula.selectRange({ r1: 1, c1: 0, r2: 3, c2: 2 }));
    const before = await state(p), selected = await selection(p); await rightClick(p, 2, 1);
    eq(await selection(p), selected);
    if (process.env.WIXEL_CELL_CONTEXT_DESKTOP_SCREENSHOT) await p.screenshot({ path: process.env.WIXEL_CELL_CONTEXT_DESKTOP_SCREENSHOT });
    await p.keyboard.press('Escape'); eq(await state(p), before);
    await rightClick(p, 5, 4); eq((await selection(p)).sel, { r1: 5, c1: 4, r2: 5, c2: 4 });
    await p.keyboard.press('Escape'); eq(await state(p), before);
  });
  await test('정렬 하위 메뉴는 행의 결합·머리글·범위 밖을 보존하고 Undo', async p => {
    const before = await state(p); await rightClick(p); await submenu(p, /^정렬/); await sub(p).getByRole('menuitem', { name: /오름차순/ }).click();
    eq(await rows(p), [['이름', '수량', '분류'], ['나', 10, '검토'], ['다', 20, '승인'], ['가', 30, '승인']]);
    eq(await p.evaluate(() => window.tabula.wb().getValue(0, 5, 0)), 60); eq((await state(p)).undo, before.undo + 1); await undo(p, before);
    await rightClick(p); await submenu(p, /^정렬/); await sub(p).getByRole('menuitem', { name: /내림차순/ }).click();
    eq((await rows(p)).slice(1).map(row => row[1]), [30, 20, 10]); await undo(p, before);
  });
  await test('사용자 지정 정렬 창 취소는 셀·Undo를 바꾸지 않는다', async p => {
    const before = await state(p); await rightClick(p); await submenu(p, /^정렬/); await sub(p).getByRole('menuitem', { name: /사용자 지정 정렬/ }).click();
    await dialog(p, '정렬').getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), before);
  });
  await test('필터 사용 메뉴는 데이터 범위에 필터를 만들고 한 번 Undo', async p => {
    const before = await state(p); await rightClick(p); await submenu(p, /^필터/); await sub(p).getByRole('menuitem', { name: /필터 사용/ }).click();
    const filter = await p.evaluate(() => window.tabula.wb().sheets[0].filter);
    eq([filter.r1, filter.c1, filter.r2, filter.c2], [0, 0, 3, 2]); eq(filter.criteria, {}); eq((await state(p)).undo, before.undo + 1); await undo(p, before);
  });
  await test('새 메모·편집·삭제는 실제 셀만 변경하고 각각 Undo', async p => {
    const before = await state(p); await rightClick(p); await item(p, /메모 삽입/).click();
    let d = dialog(p, '메모 - B2'); await d.getByRole('textbox').fill('합성 메모'); await d.getByRole('button', { name: '저장', exact: true }).click();
    eq(await p.evaluate(() => window.tabula.wb().getCell(0, 1, 1).comment), '합성 메모'); const written = await state(p);
    await rightClick(p); await item(p, /메모 편집/).click(); d = dialog(p, '메모 - B2'); await d.getByRole('textbox').fill('취소할 변경'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), written);
    await rightClick(p); await item(p, /메모 삭제/).click(); eq(await p.evaluate(() => window.tabula.wb().getCell(0, 1, 1).comment || ''), ''); await undo(p, written); await undo(p, before);
  });
  await test('셀 서식 F 접근·취소와 미니 굵게는 선택 범위 및 Undo를 보존', async p => {
    await p.evaluate(() => window.tabula.selectRange({ r1: 1, c1: 0, r2: 2, c2: 1 })); const before = await state(p), selected = await selection(p);
    await rightClick(p); await p.keyboard.press('f'); await dialog(p, '셀 서식').getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), before);
    await rightClick(p); const toolbar = p.locator('.context-mini-toolbar');
    ok(await toolbar.locator('.access-key-hint').evaluateAll(nodes => nodes.length === 0), '좁은 아이콘 단추에 상시 접근키 문구를 겹쳐 표시하지 않음');
    await p.keyboard.press('Tab'); ok(await toolbar.evaluate(el => el.contains(document.activeElement)), '메뉴에서 Tab으로 미니 서식 진입');
    for (let i = 0; i < 20 && await p.evaluate(() => document.activeElement?.dataset.miniCommand !== 'bold'); i++) await p.keyboard.press('Tab');
    eq(await p.evaluate(() => document.activeElement?.dataset.miniCommand), 'bold');
    await p.keyboard.press('Alt'); ok(await p.locator('.access-key-layer .access-key-badge').count() > 0, 'Alt 별도 접근키 배지는 유지'); await p.keyboard.press('Alt');
    await p.keyboard.press('Space'); eq(await selection(p), selected);
    eq(await p.evaluate(() => [1, 2].flatMap(r => [0, 1].map(c => !!window.tabula.wb().styleAt(0, r, c).bold))), [true, true, true, true]);
    await p.keyboard.press('Escape'); await undo(p, before);
  });
  await test('링크 삽입·편집 취소·링크만 지우기는 숫자 값과 서식을 보존', async p => {
    const before = await state(p); await rightClick(p); await item(p, /^링크/).click(); let d = dialog(p, '하이퍼링크 삽입');
    await d.getByLabel('주소 (웹 주소 또는 #시트!A1)', { exact: true }).fill('#다른 시트!A1'); await d.getByRole('button', { name: '확인', exact: true }).click();
    eq(await p.evaluate(() => window.tabula.wb().getValue(0, 1, 1)), 30); const linked = await state(p);
    await rightClick(p); await item(p, /링크 편집/).click(); await dialog(p, '하이퍼링크 삽입').getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), linked);
    await rightClick(p); await item(p, /하이퍼링크 지우기/).click();
    const cell = (await state(p)).data.sheets[0].cells['1,1']; eq(cell.link, undefined); eq(cell.style, linked.data.sheets[0].cells['1,1'].style); eq(cell.raw, linked.data.sheets[0].cells['1,1'].raw);
    await undo(p, linked); await undo(p, before);
  });
  await test('Shift+F10·방향키 하위 메뉴·Escape는 편집기와 문서를 오염시키지 않는다', async p => {
    const before = await state(p); await p.locator('#cellEditor').focus(); await p.keyboard.press('Shift+F10'); await menu(p).waitFor();
    await item(p, /^정렬/).focus(); await p.keyboard.press('ArrowRight'); await sub(p).waitFor(); await p.keyboard.press('ArrowDown'); await p.keyboard.press('ArrowLeft');
    eq(await sub(p).count(), 0); ok(await item(p, /^정렬/).evaluate(n => n === document.activeElement));
    await p.keyboard.press('Escape'); eq(await menu(p).count(), 0); eq(await p.evaluate(() => document.activeElement.id), 'cellEditor'); eq(await state(p), before);
    await p.keyboard.press('ContextMenu'); await menu(p).waitFor(); await p.keyboard.press('c'); eq(await p.evaluate(() => window.__cellClipboard), '30'); eq(await state(p), before);
  });
  await test('선택한 값·셀 색·글꼴 색 필터는 일치 행만 남기고 현재 열 해제·Undo', async p => {
    const hidden = () => p.evaluate(() => { const f = window.tabula.wb().sheets[0].filter; return [1, 2, 3].map(r => !!(f.hidden?.__bits ? f.hidden.__bits[r - f.hidden.start] : f.hidden?.[r])); });
    for (const label of [/선택한 셀 값으로/, /선택한 셀 색으로/, /선택한 셀 글꼴 색으로/]) {
      const before = await state(p); await rightClick(p, 1, 2); await submenu(p, /^필터/); await sub(p).getByRole('menuitem', { name: label }).click();
      eq(await hidden(), [false, true, false], String(label)); const filtered = await state(p); eq(filtered.undo, before.undo + 1);
      await rightClick(p, 1, 2); await submenu(p, /^필터/); await sub(p).getByRole('menuitem', { name: /현재 열에서 필터 해제/ }).click(); eq(await hidden(), [false, false, false]);
      await undo(p, filtered); await undo(p, before);
    }
  });
  await test('셀 색·글꼴 색 위에 배치는 행 단위 안정 정렬하며 Undo', async p => {
    for (const label of [/선택한 셀 색을/, /선택한 셀 글꼴 색을/]) {
      const before = await state(p); await rightClick(p, 1, 2); await submenu(p, /^정렬/); await sub(p).getByRole('menuitem', { name: label }).click();
      eq((await rows(p)).slice(1).map(row => row[0]), ['가', '다', '나']); eq((await rows(p)).slice(1).map(row => row[1]), [30, 20, 10]); await undo(p, before);
    }
  });
  await test('메모 표시·숨기기 메뉴는 원문 보존 및 Undo', async p => {
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setComment(0, 1, 1, '표시할 합성 메모')); w.undoStack = []; });
    const before = await state(p); await rightClick(p); await item(p, /메모 표시\/숨기기/).click();
    const visible = await state(p); eq(visible.data.sheets[0].noteVisibility.states['1,1'], true); eq(visible.data.sheets[0].cells['1,1'].comment, '표시할 합성 메모');
    await rightClick(p); await item(p, /메모 표시\/숨기기/).click(); eq((await state(p)).data.sheets[0].noteVisibility.states['1,1'], false); await undo(p, visible); await undo(p, before);
  });
  await test('이름 정의는 선택 범위를 참조하고 취소·저장·Undo', async p => {
    await p.evaluate(() => window.tabula.selectRange({ r1: 1, c1: 0, r2: 3, c2: 2 })); const before = await state(p);
    await rightClick(p, 2, 1); await item(p, /이름 정의/).click(); let d = dialog(p, '새 이름');
    eq(await d.getByLabel('참조 대상', { exact: true }).inputValue(), "='셀 메뉴 합성'!$A$2:$C$4");
    await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), before);
    await rightClick(p, 2, 1); await item(p, /이름 정의/).click(); d = dialog(p, '새 이름'); await d.getByLabel('이름', { exact: true }).fill('합성선택'); await d.getByRole('button', { name: '확인', exact: true }).click();
    const names = (await state(p)).data.names; eq(names.length, 1); eq(names[0].name, '합성선택'); eq(names[0].ref, "='셀 메뉴 합성'!$A$2:$C$4"); await undo(p, before);
  });
  await test('인접 텍스트 목록은 숫자·수식 제외/중복 제거·검색·선택·Undo', async p => {
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => { for (const [r, raw] of [[1, '먼저'], [2, '42'], [3, '중복'], [4, '=\"수식 제외\"'], [5, '중복'], [7, '나중'], [9, '빈칸 너머 제외']]) w.setInput(0, r, 4, raw); }); w.undoStack = []; window.tabula.gv().renderAll(); });
    const before = await state(p); await rightClick(p, 6, 4); await item(p, /드롭다운 목록/).click();
    const picker = p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true }); await picker.waitFor();
    eq((await picker.getByRole('menuitem').allTextContents()).sort(), ['나중', '먼저', '중복'].sort());
    await picker.getByRole('searchbox', { name: '목록 검색' }).fill('중복'); eq(await picker.getByRole('menuitem').count(), 1); await picker.getByRole('menuitem', { name: '중복', exact: true }).click();
    eq(await p.evaluate(() => window.tabula.wb().getValue(0, 6, 4)), '중복'); eq((await state(p)).undo, before.undo + 1); await undo(p, before);
  });
  await test('Alt+↓ 목록에서 키보드로 다음 항목 선택·Enter·Undo', async p => {
    await p.evaluate(() => window.tabula.selectCell(2, 2)); const before = await state(p); await p.locator('#cellEditor').focus(); await p.keyboard.press('Alt+ArrowDown');
    const picker = p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true }); await picker.waitFor();
    eq(await p.evaluate(() => document.activeElement.getAttribute('aria-label')), '목록 검색');
    await p.keyboard.press('ArrowDown'); const first = await p.evaluate(() => document.activeElement.textContent);
    await p.keyboard.press('ArrowDown'); const second = await p.evaluate(() => document.activeElement.textContent); ok(first !== second, '목록 내부 방향키가 다음 항목으로 이동');
    await p.keyboard.press('Enter'); await picker.waitFor({ state: 'detached' }); eq(await p.evaluate(() => window.tabula.wb().getValue(0, 2, 2)), second.replace(/ ✓$/, '')); await undo(p, before);
  });
  await test('데이터 유효성 목록은 인접 값보다 우선하고 취소로 값 불변', async p => {
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'validations', [{ r1: 2, c1: 2, r2: 2, c2: 2, type: 'list', f1: '"진행,완료"', showDropdown: true }])); w.undoStack = []; });
    const before = await state(p); await rightClick(p, 2, 2); await item(p, /드롭다운 목록/).click();
    const picker = p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true }); eq(await picker.getByRole('menuitem').allTextContents(), ['진행', '완료']);
    await p.keyboard.press('Escape'); eq(await state(p), before);
  });
  await test('윗주 편집은 구간·PHONETIC 값·표시 전환·취소와 한 번 Undo를 보존', async p => {
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => { w.setInput(0, 7, 0, '東京 韓國'); w.setInput(0, 7, 1, '=PHONETIC(A8)'); }); w.undoStack = []; window.tabula.gv().renderAll(); });
    const before = await state(p); await rightClick(p, 7, 0); await item(p, /윗주 편집/).click();
    let d = dialog(p, '윗주 편집'); await d.getByLabel('윗주', { exact: true }).first().fill('취소'); await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), before);
    await rightClick(p, 7, 0); await item(p, /윗주 편집/).click(); d = dialog(p, '윗주 편집');
    await d.getByLabel('윗주 시작 글자', { exact: true }).first().fill('1'); await d.getByLabel('윗주 끝 글자', { exact: true }).first().fill('2'); await d.getByLabel('윗주', { exact: true }).first().fill('とうきょう');
    await d.getByRole('button', { name: '윗주 구간 추가', exact: true }).click();
    await d.getByLabel('윗주 시작 글자', { exact: true }).nth(1).fill('4'); await d.getByLabel('윗주 끝 글자', { exact: true }).nth(1).fill('5'); await d.getByLabel('윗주', { exact: true }).nth(1).fill('かんこく');
    await d.getByRole('button', { name: '확인', exact: true }).click();
    const edited = await state(p); eq(edited.undo, before.undo + 1); eq(edited.data.sheets[0].cells['7,0'].phonetic.runs, [{ sb: 0, eb: 2, text: 'とうきょう' }, { sb: 3, eb: 5, text: 'かんこく' }]);
    eq(await p.evaluate(() => window.tabula.wb().getValue(0, 7, 1)), 'とうきょう かんこく');
    await rightClick(p, 7, 0); await item(p, /윗주 필드 표시/).click(); eq((await state(p)).data.sheets[0].cells['7,0'].phonetic.visible, false); await undo(p, edited); await undo(p, before);
    for (const [r, c] of [[1, 1], [7, 1]]) { await rightClick(p, r, c); eq(await item(p, /윗주 편집/).isDisabled(), true, '숫자·수식 셀은 윗주 편집 불가'); await p.keyboard.press('Escape'); }
  });
  await test('표/범위 데이터 가져오기 우클릭 연결·취소는 원본과 Undo를 보존', async p => {
    const before = await state(p); await rightClick(p); await item(p, /표\/범위에서 데이터 가져오기/).click();
    const d = dialog(p, '표/범위에서 데이터 가져오기'); await d.waitFor(); eq(await d.locator('.range-query-editor').count(), 1);
    await d.getByRole('button', { name: '취소', exact: true }).click(); eq(await state(p), before);
  });
  await test('열린 드롭다운도 뒤늦은 시트 보호·문서 전환에서 값을 덮지 않는다', async p => {
    await rightClick(p, 2, 2); await item(p, /드롭다운 목록/).click(); let picker = p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true });
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: {} })); }); const protectedState = await state(p);
    await picker.getByRole('menuitem', { name: '승인', exact: true }).click(); eq(await state(p), protectedState);
    await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'protect', null)); });
    await rightClick(p, 2, 2); await item(p, /드롭다운 목록/).click(); picker = p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true });
    await p.evaluate(() => window.tabula.switchSheet(1)); const switched = await state(p);
    await picker.getByRole('menuitem', { name: '승인', exact: true }).click(); eq(await state(p), switched); eq(await p.evaluate(() => window.tabula.si), 1);
  });
  await test('열린 드롭다운의 유효성 규칙이 추가·변경되면 이전 항목 입력을 거절', async p => {
    for (const previous of [null, '"진행,완료"']) {
      await p.evaluate(previous => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'validations', previous ? [{ r1: 2, c1: 2, r2: 2, c2: 2, type: 'list', f1: previous }] : [])); }, previous);
      await rightClick(p, 2, 2); await item(p, /드롭다운 목록/).click(); const picker = p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true });
      await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'validations', [{ r1: 2, c1: 2, r2: 2, c2: 2, type: 'list', f1: '"신규"' }])); });
      const changed = await state(p); await picker.getByRole('menuitem', { name: previous ? '완료' : '승인', exact: true }).click();
      eq(await state(p), changed, '예전 목록 선택으로 셀·규칙·Undo를 바꾸지 않음'); eq(await p.evaluate(() => window.tabula.wb().getValue(0, 2, 2)), '검토');
      ok(/변경되었습니다/.test(await p.locator('#toast').textContent()), '오래된 목록을 다시 열도록 안내');
      await rightClick(p, 2, 2); await item(p, /드롭다운 목록/).click(); eq(await p.getByRole('menu', { name: '드롭다운 목록에서 선택', exact: true }).getByRole('menuitem').allTextContents(), ['신규']); await p.keyboard.press('Escape');
    }
  });
  for (const readonly of [false, true]) await test(readonly ? '공개 읽기 전용 메뉴는 변경을 거절하고 복사만 허용' : '시트 보호는 하위 정렬·필터·메모·링크·셀 서식을 차단', async p => {
    if (!readonly) await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true } })); w.undoStack = []; });
    const before = await state(p); await rightClick(p);
    for (const label of [/잘라내기/, /내용 지우기/, /셀 서식/, /메모 삽입/, /^링크/, /드롭다운 목록/, /이름 정의/]) eq(await item(p, label).isDisabled(), true, String(label));
    await submenu(p, /^정렬/); for (const n of await sub(p).getByRole('menuitem').all()) eq(await n.isDisabled(), true);
    await p.keyboard.press('Escape'); await rightClick(p); await submenu(p, /^필터/); for (const n of await sub(p).getByRole('menuitem').all()) eq(await n.isDisabled(), true);
    await p.keyboard.press('Escape'); await rightClick(p); await item(p, /^복사/).click(); eq(await p.evaluate(() => window.__cellClipboard), '30'); eq(await state(p), before);
  }, { readonly });
  await test('320px 모바일 메뉴와 하위 메뉴가 화면 안에서 끝까지 접근 가능', async p => {
    const before = await state(p); await rightClick(p, 1, 1);
    const bounds = async locator => locator.evaluate(el => { const r = el.getBoundingClientRect(); return { l: r.left, t: r.top, r: r.right, b: r.bottom, w: innerWidth, h: innerHeight }; });
    for (const target of [menu(p), p.locator('.context-mini-toolbar')]) { const b = await bounds(target); ok(b.l >= -1 && b.r <= b.w + 1 && b.t >= -1 && b.b <= b.h + 1, JSON.stringify(b)); }
    await p.keyboard.press('End'); ok(await menu(p).evaluate(el => { const a = document.activeElement.getBoundingClientRect(), b = el.getBoundingClientRect(); return a.top >= b.top - 1 && a.bottom <= b.bottom + 1; }));
    await submenu(p, /^정렬/); const b = await bounds(sub(p)); ok(b.l >= -1 && b.r <= b.w + 1 && b.t >= -1 && b.b <= b.h + 1, JSON.stringify(b));
    await p.keyboard.press('Escape'); await p.keyboard.press('Escape'); eq(await state(p), before);
    if (process.env.WIXEL_CELL_CONTEXT_SCREENSHOT) { await rightClick(p); await p.screenshot({ path: process.env.WIXEL_CELL_CONTEXT_SCREENSHOT }); }
  }, { viewport: { width: 320, height: 680 }, mobile: true });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(r => r.ok).length, checks, bad: results.filter(r => !r.ok), results }, null, 2));
if (results.some(r => !r.ok)) process.exitCode = 1;
