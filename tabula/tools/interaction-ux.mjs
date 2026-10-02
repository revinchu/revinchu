// 합성 문서만 사용하며 실제 서버 쓰기를 차단한다. 소스/최종 번들 공통 검사.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const results = [];
async function test(name, run) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 980 } });
  const page = await context.newPage(); page.setDefaultTimeout(10000);
  const errors = [], writes = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await context.route('**/*', (route) => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/');
    await page.waitForFunction(() => window.tabula?.wb());
    await page.evaluate(() => {
      const t = window.tabula, wb = t.wb(), cells = {};
      [['지역', '부서', '매출'], ['서울', '영업', '100'], ['부산', '개발', '200'], ['대구', '영업', '300'], ['서울2', '개발', '400']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
      wb.restore({ sheets: [{ name: '검사 자료', cells }] }); wb.undoStack = []; wb.redoStack = [];
      t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
    });
    await run(page);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
const run = (p, command) => p.evaluate((cmd) => window.tabula.run(cmd), command);
const snapshot = (p) => p.evaluate(() => window.tabula.wb().serialize());
const openFilter = async (p) => {
  await p.evaluate(() => window.tabula.selectCell(0, 0));
  await p.keyboard.press('Alt+ArrowDown'); return p.locator('.filter-menu');
};
const pivot = (p, unnamed = false) => p.evaluate((unnamed) => {
  const t = window.tabula, w = t.wb();
  w.transact(() => {
    w.addSheet('피벗 검사');
    w.setSheetProp(1, 'pivot', { ...(unnamed ? {} : { name: '검사피벗' }), source: '검사 자료', range: { r1: 0, c1: 0, r2: 4, c2: 2 }, rows: ['지역'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 0, left: 0 });
  });
  t.switchSheet(1); t.run('pivotRefresh'); t.selectCell(1, 0);
}, unnamed);
const makeTable = async (p) => { await run(p, 'createTable'); await p.getByRole('dialog', { name: '표 만들기', exact: true }).getByRole('button', { name: '확인', exact: true }).click(); };
try {
  await test('일반 필터: 검색을 지워도 임시 체크 선택 보존·취소 무변경', async (p) => {
    await run(p, 'toggleFilter'); const before = await snapshot(p), menu = await openFilter(p);
    await menu.getByRole('checkbox', { name: '표시된 항목 모두 선택', exact: true }).uncheck();
    await menu.getByRole('checkbox', { name: '서울', exact: true }).check();
    await menu.getByRole('searchbox').fill('부산');
    assert.equal(await menu.getByRole('checkbox', { name: '부산', exact: true }).isChecked(), true);
    await menu.getByRole('searchbox').fill('');
    assert.equal(await menu.getByRole('checkbox', { name: '서울', exact: true }).isChecked(), true);
    assert.equal(await menu.getByRole('checkbox', { name: '부산', exact: true }).isChecked(), false);
    await menu.getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await snapshot(p), before);
  });
  await test('일반 필터: 검색 결과를 현재 선택에 추가·Enter 적용·한 번 실행 취소', async (p) => {
    await run(p, 'toggleFilter'); const before = await snapshot(p), menu = await openFilter(p);
    await menu.getByRole('checkbox', { name: '표시된 항목 모두 선택', exact: true }).uncheck();
    await menu.getByRole('checkbox', { name: '서울', exact: true }).check();
    await menu.getByRole('searchbox').fill('부산');
    await menu.getByRole('checkbox', { name: '필터에 현재 선택한 내용 추가' }).check();
    await menu.getByRole('searchbox').press('Enter');
    assert.deepEqual(await p.evaluate(() => window.tabula.wb().sheets[0].filter.criteria[0].sort()), ['부산', '서울']);
    assert.deepEqual(await p.evaluate(() => window.tabula.wb().sheets[0].filter.hidden), { 3: true, 4: true });
    assert.equal(await p.evaluate(() => document.getElementById('cellEditor').classList.contains('idle')), true, '필터 Enter가 격자 편집으로 새어 나감');
    await run(p, 'undo'); assert.deepEqual(await snapshot(p), before);
  });
  await test('일반 필터: 빈 검색 결과는 적용하지 않고 검색·체크 목록 방향키 지원', async (p) => {
    await run(p, 'toggleFilter'); const before = await snapshot(p), menu = await openFilter(p);
    await menu.getByRole('searchbox').fill('없는항목'); await menu.getByRole('searchbox').press('Enter');
    assert.equal(await menu.isVisible(), true); assert.match(await menu.innerText(), /검색 결과가 없습니다/);
    assert.deepEqual(await snapshot(p), before);
    await menu.getByRole('searchbox').fill('서울'); await menu.getByRole('searchbox').press('ArrowDown');
    await p.keyboard.press('ArrowDown'); assert.equal(await p.evaluate(() => document.activeElement.getAttribute('aria-label')), '서울');
    await p.keyboard.press('Space'); assert.equal(await menu.getByRole('checkbox', { name: '서울', exact: true }).isChecked(), false);
  });
  await test('피벗 필터: 검색 선택 추가·Enter 적용·취소 및 실행 취소', async (p) => {
    await pivot(p); const before = await snapshot(p);
    await p.locator('#pivotPane .pp-field').filter({ has: p.locator('[data-pivot-field="지역"]') }).hover();
    await p.locator('#pivotPane .pp-field').filter({ has: p.locator('[data-pivot-field="지역"]') }).locator('.pp-drop').click();
    let menu = p.locator('.filter-menu');
    await menu.getByRole('checkbox', { name: '표시된 항목 모두 선택', exact: true }).uncheck();
    await menu.getByRole('checkbox', { name: '서울', exact: true }).check();
    await menu.getByRole('searchbox').fill('부산'); await menu.getByRole('checkbox', { name: '필터에 현재 선택한 내용 추가' }).check();
    await menu.getByRole('searchbox').press('Enter');
    assert.deepEqual(await p.evaluate(() => window.tabula.wb().sheets[1].pivot.filters.지역.sort()), ['부산', '서울']);
    await run(p, 'undo'); assert.deepEqual(await snapshot(p), before);
    await p.locator('#pivotPane .pp-field').filter({ has: p.locator('[data-pivot-field="지역"]') }).hover();
    await p.locator('#pivotPane .pp-field').filter({ has: p.locator('[data-pivot-field="지역"]') }).locator('.pp-drop').click();
    menu = p.locator('.filter-menu'); await menu.getByRole('searchbox').fill('부산'); await menu.getByRole('button', { name: '취소', exact: true }).click();
    assert.deepEqual(await snapshot(p), before);
  });
  await test('피벗 필드 검색·체크 후 검색어와 포커스 유지·키보드 영역 이동', async (p) => {
    await pivot(p); const search = p.getByRole('searchbox', { name: '피벗 필드 검색' });
    await search.fill('부서'); await search.press('ArrowDown'); await p.keyboard.press('Space');
    assert.equal(await search.inputValue(), '부서');
    assert.equal(await p.evaluate(() => document.activeElement.dataset.pivotField), '부서');
    assert.equal(await p.locator('#pivotPane .pp-fields input[type=checkbox]').count(), 1);
    assert.equal(await p.locator('#pivotPane [data-pivot-field="부서"]').isChecked(), true);
    await p.keyboard.press('Alt+ArrowDown'); await p.getByRole('menuitem', { name: '열 레이블에 추가', exact: true }).click();
    assert.deepEqual(await p.evaluate(() => window.tabula.wb().sheets[1].pivot.cols), ['부서']);
    assert.equal(await search.inputValue(), '부서');
    await search.focus(); await search.press('Escape'); assert.equal(await search.inputValue(), '');
    await search.press('ArrowDown'); await p.keyboard.press('End'); assert.equal(await p.evaluate(() => document.activeElement.dataset.pivotField), '매출');
    await p.screenshot({ path: process.env.WIXEL_INTERACTION_SCREENSHOT || 'D:/Codex/Temp/wixel3-interaction-ux.png' });
  });
  await test('슬라이서 삽입: 검색 선택 누적·검색 결과 선택 해제·취소·한 번 실행 취소', async (p) => {
    await pivot(p, true); const before = await snapshot(p); await run(p, 'insertSlicer');
    let d = p.getByRole('dialog', { name: '슬라이서 삽입', exact: true });
    await d.getByRole('searchbox').fill('지'); await d.getByRole('button', { name: '검색 결과 모두 선택', exact: true }).click();
    await d.getByRole('searchbox').fill('부'); await d.getByRole('button', { name: '검색 결과 모두 선택', exact: true }).click();
    assert.match(await d.locator('[role=status]').innerText(), /2개 선택/);
    await p.screenshot({ path: process.env.WIXEL_SLICER_PICKER_SCREENSHOT || 'D:/Codex/Temp/wixel3-slicer-picker.png' });
    await d.getByRole('button', { name: '검색 결과 선택 해제', exact: true }).click();
    assert.match(await d.locator('[role=status]').innerText(), /1개 선택/);
    await d.getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await snapshot(p), before);
    await run(p, 'insertSlicer'); d = p.getByRole('dialog', { name: '슬라이서 삽입', exact: true });
    await d.getByRole('checkbox', { name: '지역', exact: true }).check(); await d.getByRole('checkbox', { name: '부서', exact: true }).check();
    await d.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[1].slicers.length), 2);
    await run(p, 'undo'); assert.deepEqual(await snapshot(p), before);
  });
  await test('표 만들기: 다른 시트 참조는 거부하고 현재 시트의 따옴표 참조 허용', async (p) => {
    const before = await snapshot(p); await run(p, 'createTable');
    const d = p.getByRole('dialog', { name: '표 만들기', exact: true });
    await d.locator('input[type=text]').fill('다른시트!A1:C5'); await d.getByRole('button', { name: '확인', exact: true }).click();
    assert.deepEqual(await snapshot(p), before); assert.match(await p.locator('.dialog').last().innerText(), /현재 시트/);
    await p.locator('.dialog').last().getByRole('button', { name: '확인', exact: true }).click();
    await d.locator('input[type=text]').fill("'검사 자료'!$A$1:$C$5"); await d.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].tables.length), 1);
    await run(p, 'undo'); assert.deepEqual(await snapshot(p), before);
  });
  await test('표 크기 조정: 병합·분산 결과 충돌 거부·정상 확장 한 번 실행 취소', async (p) => {
    await makeTable(p);
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => { w.merge(0, 6, 0, 6, 1); w.setInput(0, 8, 0, '=SEQUENCE(2,2)'); }); window.tabula.selectCell(1, 0); });
    let before = await snapshot(p);
    for (const [ref, text] of [['A1:C7', /병합/], ['A1:C10', /분산/]]) {
      if (ref === 'A1:C10') {
        await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'merges', [])); });
        before = await snapshot(p);
      }
      await run(p, 'resizeTable'); const d = p.getByRole('dialog', { name: '표 크기 조정', exact: true });
      await d.locator('input').fill(ref); await d.getByRole('button', { name: '확인', exact: true }).click();
      assert.match(await p.locator('.dialog').last().innerText(), text); assert.deepEqual(await snapshot(p), before);
      await p.keyboard.press('Escape'); await p.keyboard.press('Escape');
    }
    await run(p, 'resizeTable'); const d = p.getByRole('dialog', { name: '표 크기 조정', exact: true });
    await d.locator('input').fill("'검사 자료'!A1:C6"); await d.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].tables[0].r2), 5);
    await run(p, 'undo'); assert.deepEqual(await snapshot(p), before);
  });
  const failed = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ tests: results.length, failed, results }, null, 2));
  if (failed) process.exitCode = 1;
} finally { await browser.close(); }
