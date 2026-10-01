// 우클릭 데이터 보호 회귀. 새 컨텍스트·합성 문서만 사용하고 서버 쓰기는 차단한다.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [];
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
let checks = 0;
const equal = (actual, expected, message) => { assert.deepEqual(actual, expected, message); checks++; };
const fixture = () => ({ sheets: [{ name: '안전 회귀 합성', cells: {
  '0,0': { raw: '123', style: { bold: true } }, '0,1': { raw: '456' }, '0,2': { raw: '789' },
} }] });
const snapshot = (page) => page.evaluate(() => ({ data: window.tabula.wb().serialize(), undo: window.tabula.wb().undoStack.length }));
const values = (page) => page.evaluate(() => [0, 1, 2].map(c => window.tabula.wb().getValue(0, 0, c)));
async function test(name, fn, { data = fixture(), readonly = false } = {}) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage(), errors = [], writes = [], start = checks;
  page.setDefaultTimeout(10000);
  page.on('pageerror', error => errors.push(error.message));
  await context.route('**/*', route => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  try {
    await page.addInitScript(() => {
      window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.__safetyClipboard = '';
      Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
        writeText: async text => { window.__safetyClipboard = text; }, readText: async () => window.__safetyClipboard,
      } });
    });
    const target = readonly ? url.split('#')[0] + '#view=' + gzipSync(JSON.stringify({ docName: '읽기 전용 합성', workbook: data, view: { headers: true } })).toString('base64url') : url;
    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => !!window.tabula?.wb());
    if (readonly) await page.locator('.view-bar').waitFor();
    else await page.evaluate(data => {
      const t = window.tabula, w = t.wb(); w.restore(data); w.undoStack = []; w.redoStack = [];
      t.gv().layout(); t.gv().renderAll();
    }, data);
    await page.evaluate(() => window.tabula.selectCell(0, 0));
    await fn(page);
    equal(errors, [], '페이지 오류'); equal(writes, [], '서버 쓰기 요청');
    results.push({ name, ok: true, checks: checks - start }); console.log(`OK ${name} (${checks - start})`);
  } catch (error) {
    results.push({ name, ok: false, checks: checks - start, error: error.message, errors, writes });
    console.error(`NG ${name}: ${error.stack}`);
  } finally { await context.close(); }
}
async function checkBlockedInsertDelete(page, { allowGuardedDialog = false } = {}) {
  const before = await snapshot(page);
  await page.locator('#cellEditor').focus(); await page.keyboard.press('Shift+F10');
  equal(await page.getByRole('menuitem', { name: /^삽입/ }).isDisabled(), true, '삽입 메뉴 비활성');
  equal(await page.getByRole('menuitem', { name: /^삭제/ }).isDisabled(), true, '삭제 메뉴 비활성');
  await page.keyboard.press('Escape');
  for (const cmd of ['insertMenuKey', 'deleteMenuKey']) {
    await page.evaluate(cmd => window.tabula.run(cmd), cmd);
    // 잠금 해제 셀은 키보드 별칭이 창을 열 수 있으나 확인 시점에도 보호 검사를 해야 한다.
    if (allowGuardedDialog && await page.locator('.shift-dlg').count()) {
      await page.getByRole('dialog', { name: cmd === 'insertMenuKey' ? '삽입' : '삭제', exact: true }).getByRole('button', { name: '확인', exact: true }).click();
    }
    equal(await page.locator('.shift-dlg').count(), 0, `${cmd}는 셀 밀기 창을 열지 않는다`);
    equal(await snapshot(page), before, `${cmd}는 문서와 Undo를 바꾸지 않는다`);
    const dismiss = page.getByRole('dialog').getByRole('button', { name: '확인', exact: true });
    if (await dismiss.count()) await dismiss.last().click();
  }
}

try {
  const protectedData = fixture(); protectedData.sheets[0].protect = { on: true, allow: {} };
  await test('보호된 셀의 삽입·삭제는 메뉴와 명령 양쪽에서 차단', checkBlockedInsertDelete, { data: protectedData });
  await test('읽기 전용 게시 사본의 삽입·삭제는 메뉴와 명령 양쪽에서 차단', checkBlockedInsertDelete, { readonly: true });
  await test('열린 셀 밀기 창도 확인 시점의 보호 상태를 다시 검사', async page => {
    for (const [cmd, title] of [['insertMenuKey', '삽입'], ['deleteMenuKey', '삭제']]) {
      await page.evaluate(() => { window.tabula.wb().sheets[0].protect = null; });
      await page.evaluate(cmd => window.tabula.run(cmd), cmd);
      const dialog = page.getByRole('dialog', { name: title, exact: true }); await dialog.waitFor();
      await page.evaluate(() => { window.tabula.wb().sheets[0].protect = { on: true, allow: {} }; });
      const before = await snapshot(page);
      await dialog.getByRole('button', { name: '확인', exact: true }).click();
      equal(await snapshot(page), before, `${title} 확인은 보호된 값을 이동하지 않는다`);
      equal(await values(page), [123, 456, 789]);
    }
  });
  await test('잠금 해제된 셀 밀기도 잠긴 이웃을 바꾸지 않으며 허용된 전체 행 삽입은 유지', async page => {
    await page.evaluate(() => {
      const t = window.tabula, w = t.wb(); w.sheets[0].cells.getRC(0, 0).style.locked = false;
      w.sheets[0].protect = { on: true, allow: {} }; t.selectCell(0, 0);
    });
    await checkBlockedInsertDelete(page, { allowGuardedDialog: true });
    for (const [cmd, title] of [['insertMenuKey', '삽입'], ['deleteMenuKey', '삭제']]) {
      await page.evaluate(() => { window.tabula.wb().sheets[0].protect = null; });
      await page.evaluate(cmd => window.tabula.run(cmd), cmd);
      const dialog = page.getByRole('dialog', { name: title, exact: true }); await dialog.waitFor();
      await page.evaluate(() => { window.tabula.wb().sheets[0].protect = { on: true, allow: {} }; });
      const before = await snapshot(page);
      await dialog.getByRole('button', { name: '확인', exact: true }).click();
      equal(await snapshot(page), before, `${title}로 잠긴 이웃 셀도 이동하지 않는다`);
      equal(await values(page), [123, 456, 789]);
    }
    await page.evaluate(() => { window.tabula.wb().sheets[0].protect = { on: true, allow: { insertRows: true } }; });
    const before = await snapshot(page);
    await page.locator('#cellEditor').focus(); await page.keyboard.press('Shift+Space');
    await page.evaluate(() => window.tabula.run('insertMenuKey'));
    equal(await values(page), [null, null, null], '명시적으로 허용한 전체 행 삽입');
    equal(await page.evaluate(() => [0, 1, 2].map(c => window.tabula.wb().getValue(0, 1, c))), [123, 456, 789]);
    await page.evaluate(() => window.tabula.run('undo'));
    equal(await snapshot(page), before, '허용한 전체 행 삽입을 되돌리면 원래 문서로 복구');
  });
  await test('잘라낸 셀의 선택 붙여넣기는 원본 삭제 없이 거절하고 일반 이동은 유지', async page => {
    await page.evaluate(() => { const t = window.tabula; t.selectCell(0, 0); t.run('cut'); t.selectCell(0, 1); });
    const before = await snapshot(page);
    for (const cmd of ['pasteFormats', 'pasteFormulas', 'pasteValuesKey', 'pasteTranspose']) {
      await page.evaluate(cmd => window.tabula.run(cmd), cmd);
      equal(await snapshot(page), before, `${cmd} 원본·대상·Undo 유지`);
      equal(await values(page), [123, 456, 789]);
    }
    await page.evaluate(() => window.tabula.run('paste'));
    await page.waitForFunction(() => window.tabula.wb().getValue(0, 0, 1) === 123);
    equal(await values(page), [null, 123, 789]);
    await page.evaluate(() => window.tabula.run('undo'));
    equal(await values(page), [123, 456, 789]);
  });
  await test('셀 값 잠금·서식 허용인 시트에서 복사한 서식만 붙여넣기 가능', async page => {
    await page.evaluate(() => {
      const t = window.tabula, w = t.wb(); w.sheets[0].protect = { on: true, allow: { formatCells: true } };
      t.selectCell(0, 0); t.run('copy'); t.selectCell(0, 1); t.run('pasteFormats');
    });
    equal(await values(page), [123, 456, 789]);
    equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 1).bold), true);
    await page.evaluate(() => window.tabula.run('undo'));
    equal(await page.evaluate(() => !!window.tabula.wb().styleAt(0, 0, 1).bold), false);
  });
  await test('붙여넣기 확장 범위에 잠긴 셀이 있으면 대상 전체를 변경하지 않는다', async page => {
    await page.evaluate(() => {
      const t = window.tabula, w = t.wb();
      t.selectRange({ r1: 0, c1: 0, r2: 0, c2: 1 }); t.run('copy');
      w.sheets[0].cells.getRC(0, 1).style = { locked: false }; w.sheets[0].protect = { on: true, allow: {} }; t.selectCell(0, 1);
    });
    const before = await snapshot(page); await page.evaluate(() => window.tabula.run('pasteFormulas'));
    equal(await snapshot(page), before); equal(await values(page), [123, 456, 789]);
  });
  const cacheData = { sheets: [{ name: '치수 캐시 합성', fileValues: true, cells: {
    '0,0': { raw: '2' }, '0,1': { raw: '=NO_SUCH_FN(A1)', cached: 20 },
  } }] };
  await test('숨김 변화 없는 행 높이·열 너비 변경과 Undo는 가져온 계산 캐시를 보존', async page => {
    const status = () => page.evaluate(() => {
      const w = window.tabula.wb(); return { value: w.getValue(0, 0, 1), status: w.getCalculationStatus(0, 0, 1).status };
    });
    equal(await status(), { value: 20, status: 'cached' });
    for (const [cmd, title, value] of [['rowHeight', '행 높이', '30'], ['colWidth', '열 너비', '20']]) {
      const before = await snapshot(page); await page.evaluate(cmd => window.tabula.run(cmd), cmd);
      const dialog = page.getByRole('dialog', { name: title, exact: true }); await dialog.waitFor();
      await dialog.getByRole('spinbutton').fill(value); await dialog.getByRole('button', { name: '확인', exact: true }).click();
      equal(await status(), { value: 20, status: 'cached' }, `${title} 변경은 계산 입력이 아니다`);
      equal(await page.evaluate(() => window.tabula.wb().undoStack.length), before.undo + 1);
      equal(await page.evaluate(() => ({ rows: window.tabula.wb().sheets[0].hiddenRows, cols: window.tabula.wb().sheets[0].hiddenCols })), { rows: {}, cols: {} });
      await page.evaluate(() => window.tabula.run('undo'));
      equal(await status(), { value: 20, status: 'cached' }); equal(await snapshot(page), before);
    }
  }, { data: cacheData });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, scenarios: results.length, passed: results.filter(r => r.ok).length, checks, results }, null, 2));
if (results.some(r => !r.ok)) process.exitCode = 1;
