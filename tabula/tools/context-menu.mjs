// 우클릭 메뉴 회귀: 새 브라우저 컨텍스트와 합성 문서만 사용하며 원격 쓰기는 차단한다.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
import { MAX_ROWS, MAX_COLS } from '../src/formula.js';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [];
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const fixture = { sheets: [{ name: '우클릭 합성 문서', cells: {
  '0,0': { raw: '이름' }, '0,1': { raw: '수량' }, '1,0': { raw: '가' }, '1,1': { raw: '10' },
  '2,0': { raw: '나' }, '2,1': { raw: '20' }, '3,0': { raw: '다' }, '3,1': { raw: '=SUM(B2:B3)' },
} }] };
async function test(name, check, options = {}) {
  const context = await browser.newContext({ viewport: options.viewport || { width: 1440, height: 1000 } });
  const p = await context.newPage(), errors = [], writes = []; p.setDefaultTimeout(10000);
  p.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  try {
    await p.addInitScript(() => {
      window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
      window.__mockClipboard = '';
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
        writeText: async text => { window.__mockClipboard = text; }, readText: async () => window.__mockClipboard,
      } });
    });
    const target = options.readonly ? url.split('#')[0] + '#view=' + gzipSync(JSON.stringify({ docName: '읽기 전용 합성', workbook: fixture, view: { headers: true } })).toString('base64url') : url;
    await p.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    if (options.readonly) { await p.locator('.view-bar').waitFor(); await p.evaluate(() => window.tabula.selectCell(1, 1)); }
    else await p.evaluate(data => {
      const t = window.tabula, w = t.wb(); w.restore(data); w.undoStack = []; w.redoStack = [];
      t.gv().layout(); t.gv().renderAll(); t.selectCell(1, 1);
    }, fixture);
    await p.locator('#cellEditor').focus(); await check(p);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) {
    results.push({ name, ok: false, error: e.message, pageErrors: errors, blockedWrites: writes }); console.error(`NG ${name}: ${e.stack}`);
  } finally { await context.close(); }
}
const selection = p => p.evaluate(() => ({ sel: window.tabula.sel, active: window.tabula.active }));
const rootMenu = p => p.locator('#menuLayer > .menu').first();
const menuItem = (p, name) => rootMenu(p).getByRole('menuitem', { name });
const snapshot = p => p.evaluate(() => {
  const w = window.tabula.wb(); return { workbook: w.serialize(), undo: w.undoStack.length };
});
const dimensions = p => p.evaluate(() => {
  const w = window.tabula.wb(), s = w.sheets[0];
  return { rows: [0, 1, 2, 3].map(r => w.rowHeight(0, r)), cols: [0, 1, 2, 3].map(c => w.colWidth(0, c)),
    hiddenRows: { ...s.hiddenRows }, hiddenCols: { ...s.hiddenCols }, undo: w.undoStack.length };
});
async function finishSize(p, name, value, cancel = false) {
  const d = p.getByRole('dialog', { name, exact: true }); await d.waitFor(); await d.getByRole('spinbutton').fill(String(value));
  await d.getByRole('button', { name: cancel ? '취소' : '확인', exact: true }).click(); await d.waitFor({ state: 'hidden' });
}
async function selectComments(p) {
  await p.evaluate(() => {
    const t = window.tabula, w = t.wb(); w.transact(() => { w.setComment(0, 0, 0, '첫 메모'); w.setComment(0, 2, 0, '둘째 메모'); });
    w.undoStack = []; w.redoStack = []; t.selectCell(0, 0);
  });
  await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+Shift+o');
}
async function clickCell(p, r, c, zone = 'cell', button = 'right') {
  const point = await p.evaluate(({ r, c, zone }) => {
    const g = window.tabula.gv(), rect = g.clientRect({ r1: r, c1: c, r2: r, c2: c });
    const view = g.viewEl.getBoundingClientRect();
    return { x: zone === 'row' ? view.left + g.hw * g.z / 2 : (rect.left + rect.right) / 2,
      y: zone === 'col' ? view.top + g.hh * g.z / 2 : (rect.top + rect.bottom) / 2 };
  }, { r, c, zone });
  await p.mouse.click(point.x, point.y, { button });
  if (button === 'right') await rootMenu(p).waitFor();
}

try {
  await test('행 전체 선택 뒤 셀 본문 우클릭은 행 메뉴와 선택 범위를 유지', async p => {
    await p.keyboard.press('Shift+Space'); const before = await selection(p);
    assert.deepEqual(before.sel, { r1: 1, c1: 0, r2: 1, c2: MAX_COLS - 1 });
    await clickCell(p, 1, 2); assert.deepEqual(await selection(p), before);
    assert.equal(await menuItem(p, /^행 높이.*\.\.\./).count(), 1); assert.equal(await menuItem(p, /셀 서식/).count(), 1);
    if (process.env.CONTEXT_MENU_SCREENSHOT) await p.screenshot({ path: process.env.CONTEXT_MENU_SCREENSHOT, fullPage: true });
  });
  await test('열 전체 선택 뒤 셀 본문 우클릭은 열 메뉴와 선택 범위를 유지', async p => {
    await p.keyboard.press('Control+Space'); const before = await selection(p);
    assert.deepEqual(before.sel, { r1: 0, c1: 1, r2: MAX_ROWS - 1, c2: 1 });
    await clickCell(p, 3, 1); assert.deepEqual(await selection(p), before);
    assert.equal(await menuItem(p, /^열 너비.*\.\.\./).count(), 1); assert.equal(await menuItem(p, /셀 서식/).count(), 1);
  });
  await test('행·열 머리글 메뉴에 셀 서식이 있고 Shift+F10도 선택 종류를 반영', async p => {
    await clickCell(p, 2, 1, 'row'); assert.equal(await menuItem(p, /셀 서식/).count(), 1);
    await p.keyboard.press('Escape'); await p.locator('#cellEditor').focus(); await p.keyboard.press('Shift+F10');
    assert.equal(await menuItem(p, /^행 높이.*\.\.\./).count(), 1); await p.keyboard.press('Escape');
    await clickCell(p, 1, 2, 'col'); assert.equal(await menuItem(p, /셀 서식/).count(), 1);
    await p.keyboard.press('Escape'); await p.locator('#cellEditor').focus(); await p.keyboard.press('ContextMenu');
    assert.equal(await menuItem(p, /^열 너비.*\.\.\./).count(), 1);
  });
  await test('선택 밖 우클릭만 대상 셀을 변경하며 일반 범위 우클릭은 선택을 유지', async p => {
    await p.evaluate(() => window.tabula.selectRange({ r1: 1, c1: 0, r2: 3, c2: 2 })); const before = await selection(p);
    await clickCell(p, 2, 1); assert.deepEqual(await selection(p), before); await p.keyboard.press('Escape');
    await clickCell(p, 4, 3); assert.deepEqual(await selection(p), { sel: { r1: 4, c1: 3, r2: 4, c2: 3 }, active: { r: 4, c: 3 } });
  });
  await test('행 높이 R·열 너비 W는 취소 시 무변경, 확인 시 한 번 실행 취소', async p => {
    const before = await dimensions(p); await clickCell(p, 1, 1, 'row');
    await p.keyboard.press('r'); await finishSize(p, '행 높이', 36, true); assert.deepEqual(await dimensions(p), before);
    await clickCell(p, 1, 1, 'row'); await p.keyboard.press('r'); await finishSize(p, '행 높이', 36);
    assert.deepEqual((await dimensions(p)).rows, [before.rows[0], 48, before.rows[2], before.rows[3]]);
    assert.equal((await dimensions(p)).undo, before.undo + 1); await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
    await clickCell(p, 1, 2, 'col'); await p.keyboard.press('w'); await finishSize(p, '열 너비', 20, true); assert.deepEqual(await dimensions(p), before);
    await clickCell(p, 1, 2, 'col'); await p.keyboard.press('w'); await finishSize(p, '열 너비', 20);
    assert.deepEqual((await dimensions(p)).cols, [before.cols[0], before.cols[1], 165, before.cols[3]]);
    assert.equal((await dimensions(p)).undo, before.undo + 1); await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
  });
  await test('행·열 크기 범위 초과 입력은 문서 변경 없이 대화상자에서 다시 입력', async p => {
    for (const [zone, key, title, value] of [['row', 'r', '행 높이', '410'], ['col', 'w', '열 너비', '256']]) {
      const before = await snapshot(p); await clickCell(p, 1, 1, zone); await p.keyboard.press(key);
      const d = p.getByRole('dialog', { name: title, exact: true }); await d.waitFor();
      await d.getByRole('spinbutton').fill(value); await d.getByRole('button', { name: '확인', exact: true }).click();
      assert.equal(await d.isVisible(), true); assert.deepEqual(await snapshot(p), before);
      await d.getByRole('button', { name: '취소', exact: true }).click();
    }
  });
  await test('크기 0은 기존 치수를 보존하며 숨기고 양수 입력은 숨김 해제, Undo로 모두 복원', async p => {
    await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => { w.setRowHeight(0, 1, 52); w.setColWidth(0, 1, 155); }); w.undoStack = []; w.redoStack = []; t.gv().layout(); t.gv().renderAll(); });
    for (const [zone, key, title, positive, expected, sizes, hidden] of [
      ['row', 'r', '행 높이', 45, 60, 'rows', 'hiddenRows'], ['col', 'w', '열 너비', 20, 165, 'cols', 'hiddenCols'],
    ]) {
      const before = await dimensions(p); await clickCell(p, 1, 1, zone); await p.keyboard.press(key); await finishSize(p, title, 0);
      let current = await dimensions(p); assert.deepEqual(current[sizes], before[sizes]); assert.equal(current[hidden][1], true);
      await p.locator('#cellEditor').focus(); await p.keyboard.press('Shift+F10'); await p.keyboard.press(key); await finishSize(p, title, positive);
      current = await dimensions(p); assert.equal(current[sizes][1], expected); assert.equal(current[hidden][1], undefined);
      await p.keyboard.press('Control+z'); current = await dimensions(p); assert.equal(current[hidden][1], true); assert.deepEqual(current[sizes], before[sizes]);
      await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
    }
  });
  await test('수동 행 높이 뒤 자동 맞춤은 줄 바꿈 내용을 반영하고 수동 고정을 해제, Undo 가능', async p => {
    await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => { w.setInput(0, 1, 0, '첫 줄\n둘째 줄\n셋째 줄'); w.setStyle(0, 1, 0, { wrap: true }); w.setRowHeight(0, 1, 96); }); w.undoStack = []; w.redoStack = []; t.gv().layout(); t.gv().renderAll(); });
    const before = await snapshot(p); await clickCell(p, 1, 1, 'row'); await menuItem(p, /^행 높이 자동 맞춤/).click();
    const current = await dimensions(p); assert.ok(current.rows[1] > current.rows[0] && current.rows[1] < 96);
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].rowManual[1]), undefined); assert.equal(current.undo, 1);
    await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
  });
  await test('20,000개를 넘는 일부 행 선택은 크기·숨김을 일부만 처리하지 않고 거부', async p => {
    await p.evaluate(max => window.tabula.selectRange({ r1: 1, c1: 0, r2: 20001, c2: max - 1 }, 'rows'), MAX_COLS);
    const before = await snapshot(p); await clickCell(p, 1, 1); await p.keyboard.press('r');
    assert.equal(await p.getByRole('dialog', { name: '행 높이', exact: true }).count(), 0); assert.deepEqual(await snapshot(p), before);
    assert.match(await p.locator('body').innerText(), /범위를 줄여/);
    await clickCell(p, 1, 1); await p.keyboard.press('h'); assert.deepEqual(await snapshot(p), before);
  });
  await test('행·열 메뉴의 H/U 숨기기·숨기기 취소와 Undo는 실제 크기와 선택을 보존', async p => {
    const before = await dimensions(p); await clickCell(p, 1, 1, 'row'); const selected = await selection(p); await p.keyboard.press('h');
    assert.equal((await dimensions(p)).hiddenRows[1], true); assert.deepEqual(await selection(p), selected);
    await p.locator('#cellEditor').focus(); await p.keyboard.press('Shift+F10'); await p.keyboard.press('u'); assert.deepEqual((await dimensions(p)).hiddenRows, before.hiddenRows);
    await p.keyboard.press('Control+z'); assert.equal((await dimensions(p)).hiddenRows[1], true); await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
    await clickCell(p, 1, 2, 'col'); const selectedCol = await selection(p); await p.keyboard.press('h');
    assert.equal((await dimensions(p)).hiddenCols[2], true); assert.deepEqual(await selection(p), selectedCol);
    await p.locator('#cellEditor').focus(); await p.keyboard.press('ContextMenu'); await p.keyboard.press('u'); assert.deepEqual((await dimensions(p)).hiddenCols, before.hiddenCols);
    await p.keyboard.press('Control+z'); assert.equal((await dimensions(p)).hiddenCols[2], true); await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
  });
  await test('행·열 메뉴는 T/C/S/I/D/N/F/R 또는 W/H/U 접근키를 명시하고 F로 셀 서식을 연다', async p => {
    for (const zone of ['row', 'col']) {
      await clickCell(p, 1, 1, zone);
      for (const key of ['t', 'c', 's', 'i', 'd', 'n', 'f', zone === 'row' ? 'r' : 'w', 'h', 'u']) {
        assert.equal(await rootMenu(p).locator(`:scope > [data-access-key="${key}"]`).count(), 1, zone + ' 접근키 ' + key);
      }
      const before = await snapshot(p); await p.keyboard.press('f');
      const d = p.getByRole('dialog', { name: '셀 서식', exact: true }); await d.waitFor(); await d.getByRole('button', { name: '취소', exact: true }).click();
      assert.deepEqual(await snapshot(p), before);
    }
  });
  await test('비연속 메모 셀의 우클릭 지우기는 중간 미선택 셀을 보존하고 Undo로 복원', async p => {
    await selectComments(p); const before = await snapshot(p), selected = await selection(p);
    await clickCell(p, 2, 0); assert.deepEqual(await selection(p), selected); await menuItem(p, /내용 지우기/).click();
    assert.deepEqual(await p.evaluate(() => [0, 1, 2].map(r => window.tabula.wb().getRaw(0, r, 0))), ['', '가', '']);
    await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
  });
  await test('비연속 선택의 빈틈 우클릭은 그 셀만 새로 선택하며 다른 셀을 지우지 않는다', async p => {
    await selectComments(p); await clickCell(p, 1, 0);
    assert.deepEqual(await selection(p), { sel: { r1: 1, c1: 0, r2: 1, c2: 0 }, active: { r: 1, c: 0 } });
    await menuItem(p, /내용 지우기/).click(); assert.deepEqual(await p.evaluate(() => [0, 1, 2].map(r => window.tabula.wb().getRaw(0, r, 0))), ['이름', '', '나']);
  });
  await test('행 삽입·삭제 접근키는 실제 셀을 이동하며 각각 한 번 Undo로 복원', async p => {
    const before = await snapshot(p); await clickCell(p, 1, 1, 'row'); await p.keyboard.press('i');
    assert.deepEqual(await p.evaluate(() => [0, 1, 2, 3].map(r => window.tabula.wb().getRaw(0, r, 0))), ['이름', '', '가', '나']);
    assert.equal((await snapshot(p)).undo, before.undo + 1); await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
    await clickCell(p, 1, 1, 'row'); await p.keyboard.press('d');
    assert.deepEqual(await p.evaluate(() => [0, 1, 2].map(r => window.tabula.wb().getRaw(0, r, 0))), ['이름', '나', '다']);
    assert.equal((await snapshot(p)).undo, before.undo + 1); await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
  });
  await test('일반 셀 숨기기 하위 메뉴는 키보드 진입·뒤로 가기·행 숨기기를 지원', async p => {
    const before = await dimensions(p); await clickCell(p, 1, 1); await p.keyboard.press('h');
    let sub = p.locator('#menuLayer > .menu[data-level="1"]'); await sub.waitFor();
    assert.equal(await sub.getByRole('menuitem', { name: /^행 숨기기(?:\(|$)/ }).count(), 1);
    assert.equal(await sub.getByRole('menuitem', { name: /^열 숨기기(?:\(|$)/ }).count(), 1);
    await p.keyboard.press('ArrowLeft'); assert.equal(await sub.count(), 0);
    assert.equal(await rootMenu(p).locator('[data-access-key="h"]').evaluate(el => el === document.activeElement), true);
    await p.keyboard.press('ArrowRight'); await sub.waitFor(); await p.keyboard.press('ArrowDown'); await p.keyboard.press('ArrowDown');
    assert.match(await p.evaluate(() => document.activeElement.textContent), /^행 숨기기/); await p.keyboard.press('Enter');
    assert.equal((await dimensions(p)).hiddenRows[1], true); await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
  });
  await test('비연속 선택의 행 높이 변경은 선택한 두 행만 바꾸고 Undo로 복원', async p => {
    await selectComments(p); const before = await dimensions(p); await clickCell(p, 2, 0); await p.keyboard.press('h'); await p.keyboard.press('r');
    await finishSize(p, '행 높이', 36);
    assert.deepEqual((await dimensions(p)).rows, [48, before.rows[1], 48, before.rows[3]]);
    assert.equal((await dimensions(p)).undo, before.undo + 1); await p.keyboard.press('Control+z'); assert.deepEqual(await dimensions(p), before);
  });
  await test('복사 C·선택하여 붙여넣기 S 취소·붙여넣기 옵션 값은 모의 클립보드만 사용', async p => {
    await clickCell(p, 3, 1); await p.keyboard.press('c'); assert.equal(await p.evaluate(() => window.__mockClipboard), '30');
    const before = await snapshot(p); await clickCell(p, 4, 3); await p.keyboard.press('s');
    assert.match(await p.evaluate(() => document.activeElement.textContent), /선택하여 붙여넣기/);
    await p.keyboard.press('s'); assert.match(await p.evaluate(() => document.activeElement.textContent), /윗주 필드 표시/);
    await p.keyboard.press('Escape'); assert.deepEqual(await snapshot(p), before, '중복 접근키 순환 후 Escape는 문서를 바꾸지 않음');
    await clickCell(p, 4, 3); await p.keyboard.press('s'); await p.keyboard.press('Enter');
    const d = p.getByRole('dialog', { name: '선택하여 붙여넣기', exact: true }); await d.waitFor(); await d.getByRole('button', { name: '취소', exact: true }).click();
    assert.deepEqual(await snapshot(p), before); await clickCell(p, 4, 3); await menuItem(p, '붙여넣기 옵션').click();
    const sub = p.locator('#menuLayer > .menu[data-level="1"]'); await sub.getByRole('menuitem', { name: /^값\(V\)/ }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 4, 3)), 30);
    assert.equal(await p.evaluate(() => window.tabula.wb().getCell(0, 4, 3)?.fx), undefined);
    assert.equal((await snapshot(p)).undo, before.undo + 1); await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
  });
  await test('보호된 시트의 행·열 크기/숨김/삭제/서식은 메뉴에서도 차단', async p => {
    await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true } })); w.undoStack = []; t.gv().renderAll(); });
    const before = await snapshot(p);
    for (const zone of ['row', 'col']) {
      await clickCell(p, 1, 1, zone);
      for (const key of ['i', 'd', 'n', 'f', zone === 'row' ? 'r' : 'w', 'h', 'u']) {
        const item = rootMenu(p).locator(`:scope > [data-access-key="${key}"]`); assert.equal(await item.count(), 1);
        assert.equal(await item.isDisabled(), true, '보호된 ' + zone + ' 메뉴 ' + key);
        await p.keyboard.press(key); assert.deepEqual(await snapshot(p), before, key + ' 접근키 우회 금지');
      }
      await p.keyboard.press('Escape');
    }
  });
  await test('공개 읽기 전용 문서의 우클릭 변경 명령과 미니 서식은 차단되고 복사는 허용', async p => {
    const before = await snapshot(p); await p.keyboard.press('Shift+Space'); await clickCell(p, 1, 1);
    for (const key of ['t', 'i', 'd', 'n', 'f', 'r', 'h', 'u']) {
      const item = rootMenu(p).locator(`:scope > [data-access-key="${key}"]`); assert.equal(await item.count(), 1);
      assert.equal(await item.isDisabled(), true, '읽기 전용 접근키 ' + key); await p.keyboard.press(key);
    }
    const toolbar = p.getByRole('toolbar', { name: '미니 서식 도구 모음', exact: true }); await toolbar.waitFor();
    assert.equal(await toolbar.locator('button:not(:disabled), input:not(:disabled), select:not(:disabled)').count(), 0);
    assert.equal(await rootMenu(p).locator('[data-access-key="c"]').isDisabled(), false); await p.keyboard.press('c');
    assert.match(await p.evaluate(() => window.__mockClipboard), /가\t10/); assert.deepEqual(await snapshot(p), before);
  }, { readonly: true });
  await test('미니 서식은 선택을 유지하고 굵게 한 번 Undo, Tab 이동·Esc 종료·포커스 복원', async p => {
    await p.evaluate(() => window.tabula.selectRange({ r1: 1, c1: 0, r2: 2, c2: 1 })); const before = await snapshot(p), selected = await selection(p);
    await clickCell(p, 2, 1); const toolbar = p.getByRole('toolbar', { name: '미니 서식 도구 모음', exact: true }); await toolbar.waitFor();
    await p.keyboard.press('Tab'); assert.ok(await toolbar.evaluate(el => el.contains(document.activeElement)), '메뉴 Tab → 미니 서식');
    await toolbar.locator('[data-mini-command="bold"]').click(); assert.deepEqual(await selection(p), selected);
    assert.deepEqual(await p.evaluate(() => [1, 2].flatMap(r => [0, 1].map(c => window.tabula.wb().styleAt(0, r, c).bold))), [true, true, true, true]);
    assert.equal((await snapshot(p)).undo, before.undo + 1);
    await p.keyboard.press('Escape'); assert.equal(await toolbar.count(), 0); assert.equal(await rootMenu(p).count(), 0);
    assert.equal(await p.evaluate(() => document.activeElement.id), 'cellEditor'); await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
  });
  await test('작은 화면에서도 메뉴 마지막 항목과 미니 서식이 겹치지 않고 키보드로 접근', async p => {
    await clickCell(p, 1, 1, 'row'); const toolbar = p.getByRole('toolbar', { name: '미니 서식 도구 모음', exact: true }); await toolbar.waitFor();
    const rects = await p.evaluate(() => {
      const read = e => { const r = e.getBoundingClientRect(); return { left: r.left, top: r.top, right: r.right, bottom: r.bottom }; };
      return { menu: read(document.querySelector('#menuLayer > .menu')), toolbar: read(document.querySelector('.context-mini-toolbar')), width: innerWidth, height: innerHeight };
    });
    for (const rect of [rects.menu, rects.toolbar]) { assert.ok(rect.left >= 0 && rect.top >= 0 && rect.right <= rects.width && rect.bottom <= rects.height, JSON.stringify(rects)); }
    assert.ok(rects.toolbar.bottom <= rects.menu.top || rects.menu.bottom <= rects.toolbar.top, '메뉴와 미니 도구 모음 겹침 없음');
    await p.keyboard.press('End');
    assert.ok(await rootMenu(p).evaluate(menu => {
      const last = [...menu.querySelectorAll(':scope > .menu-item:not(:disabled)')].at(-1), r = last.getBoundingClientRect(), m = menu.getBoundingClientRect();
      return document.activeElement === last && r.top >= m.top && r.bottom <= m.bottom + 1;
    }), 'End로 마지막 메뉴 항목이 화면 안에 노출');
    await p.keyboard.press('Shift+Tab'); assert.ok(await toolbar.evaluate(el => el.contains(document.activeElement))); await p.keyboard.press('Escape');
    assert.equal(await toolbar.count(), 0); assert.equal(await rootMenu(p).count(), 0);
  }, { viewport: { width: 760, height: 520 } });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(x => x.ok).length, bad: results.filter(x => !x.ok) }, null, 2));
if (results.some(x => !x.ok)) process.exitCode = 1;
