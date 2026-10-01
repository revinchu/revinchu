// UI 회귀 검사. npm start 후 실행하며 서버 문서 자동 저장/복원은 격리합니다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const results = [];
let screenshotSaved = false;
const test = async (name, fn, { server = false } = {}) => {
  const page = await browser.newPage({ viewport: { width: 1440, height: 900 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  try {
    await page.addInitScript(() => { window.WIXEL_SKIP_START = true; });
    if (!server) await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    else {
      await page.route('**/api/health', (route) => route.fulfill({ json: { ok: true, auth: true } }));
      await page.route('**/api/files**', (route) => route.fulfill({ json: [] }));
    }
    await page.goto(process.env.WIXEL_URL || 'http://localhost:5178/');
    await page.waitForFunction(() => !!window.tabula);
    await fn(page);
    assert.deepEqual(errors, [], '페이지 오류');
    results.push({ name, ok: true });
    console.log(`OK ${name}`);
  } catch (e) {
    results.push({ name, ok: false });
    console.error(`NG ${name}: ${e.stack}`);
  } finally { await page.close(); }
};
const openUi = (page, fn) => page.evaluate(async (source) => { const ui = await import('/src/ui.js'); return await new Function('ui', `return (${source})(ui)`)(ui); }, fn.toString());
const pivotFixture = (page) => page.evaluate(() => {
  const t = window.tabula; const wb = t.wb();
  wb.transact(() => {
    [['부서', '매출'], ['서울', '100'], ['부산', '200'], ['서울', '50']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v)));
    wb.addSheet('보고서'); wb.addSheet('대상');
    wb.setSheetProp(1, 'pivot', { name: '검증피벗', source: wb.sheets[0].name, range: { r1: 0, c1: 0, r2: 3, c2: 1 }, rows: ['부서'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 2, left: 1 });
  });
  t.switchSheet(1); t.run('pivotRefresh'); t.selectCell(2, 1);
  wb.transact(() => {
    wb.setInput(0, 7, 0, "='보고서'!C4");
    wb.setInput(0, 8, 0, '=GETPIVOTDATA("매출",보고서!B3)');
    wb.setStyle(1, 3, 2, { fill: '#ffff00' });
    wb.setSheetProp(0, 'charts', [{ id: 'ui-chart', type: 'column', x: 500, y: 20, w: 300, h: 200, pivot: { sheet: '보고서', name: '검증피벗' } }]);
    wb.setSheetProp(0, 'slicers', [{ id: 'ui-slicer', name: '지역', x: 850, y: 20, w: 180, h: 200, source: { kind: 'pivot', field: '부서', pivots: [{ sheet: '보고서', name: '검증피벗' }] } }]);
  });
  const area = wb.sheets[1].pivot.area;
  const cells = [];
  for (let r = area.r1; r <= area.r2; r++) for (let c = area.c1; c <= area.c2; c++) cells.push([r - area.r1, c - area.c1, wb.getCell(1, r, c)?.raw ?? '']);
  return { area, cells, value: wb.getValue(0, 7, 0), total: wb.getValue(0, 8, 0) };
});
const movePivot = async (page, location) => {
  await page.evaluate(() => window.tabula.run('pivotMove'));
  const dialog = page.getByRole('dialog', { name: '피벗 테이블 이동', exact: true });
  if (location) {
    await dialog.locator('select').selectOption('existing');
    await dialog.locator('.ref-input').fill(location);
  }
  if (process.env.WIXEL_UI_SCREENSHOT && !screenshotSaved) { await page.screenshot({ path: process.env.WIXEL_UI_SCREENSHOT }); screenshotSaved = true; }
  await dialog.getByRole('button', { name: '확인', exact: true }).click();
  return dialog;
};

try {
  await test('메모 비연속 선택: Delete와 Ctrl+Enter가 사이의 일반 셀을 보존', async (page) => {
    await page.evaluate(() => {
      const t = window.tabula; const wb = t.wb();
      wb.transact(() => { for (let r = 0; r < 3; r++) wb.setInput(0, r, 0, String(r + 1)); wb.setComment(0, 0, 0, '첫째'); wb.setComment(0, 2, 0, '셋째'); });
      t.selectCell(0, 0);
    });
    await page.keyboard.press('Control+Shift+o'); await page.keyboard.press('Delete');
    assert.deepEqual(await page.evaluate(() => [0, 1, 2].map((r) => window.tabula.wb().getCell(0, r, 0)?.raw)), ['', '2', '']);
    await page.keyboard.press('Control+z'); await page.keyboard.press('Control+Shift+o');
    await page.keyboard.type('새값'); await page.keyboard.press('Control+Enter');
    assert.deepEqual(await page.evaluate(() => [0, 1, 2].map((r) => window.tabula.wb().getCell(0, r, 0)?.raw)), ['새값', '2', '새값']);
  });
  await test('비연속 선택의 리본·우클릭 지우기 4종과 서식은 선택 셀에만 적용', async (page) => {
    for (const command of ['clearContents', 'clearAll', 'clearFormats', 'clearComments', 'bold']) {
      const cells = await page.evaluate((cmd) => {
        const t = window.tabula; const wb = t.wb();
        wb.transact(() => {
          for (let r = 0; r < 3; r++) wb.setCellData(0, r, 0, { raw: String(r + 1), style: { fill: '#abcdef' }, ...(r !== 1 ? { comment: '선택 메모' } : {}) });
        });
        t.selectCell(0, 0); t.run('selectComments');
        // 선택 완료 뒤 가운데에 새 메모가 생겨도 기존 선택 범위가 넓어지면 안 됩니다.
        if (cmd === 'clearComments') wb.transact(() => wb.setComment(0, 1, 0, '보존 메모'));
        t.run(cmd);
        return [0, 1, 2].map((r) => { const c = wb.getCell(0, r, 0); return c ? { raw: c.raw, style: c.style ?? {}, comment: c.comment ?? '' } : null; });
      }, command);
      assert.equal(cells[1].raw, '2', command);
      assert.equal(cells[1].style.fill, '#abcdef', command);
      assert.equal(cells[1].style.bold, undefined, command);
      if (command === 'clearContents') assert.deepEqual([cells[0].raw, cells[2].raw], ['', '']);
      if (command === 'clearAll') assert.deepEqual([cells[0], cells[2]], [null, null]);
      if (command === 'clearFormats') assert.deepEqual([cells[0].style.fill, cells[2].style.fill], [undefined, undefined]);
      if (command === 'clearComments') { assert.equal(cells[1].comment, '보존 메모'); assert.deepEqual([cells[0].comment, cells[2].comment], ['', '']); }
      if (command === 'bold') assert.deepEqual([cells[0].style.bold, cells[2].style.bold], [true, true]);
    }
  });
  await test('모달 Tab 순환·중첩 포커스 복원·한글 조합 Enter 보호', async (page) => {
    await openUi(page, (ui) => { window.uiCount = 0; ui.openDialog({ title: '부모 창', body: ui.el('input', { id: 'parent-input' }), buttons: [{ label: '확인', primary: true, action: () => { window.uiCount++; } }, { label: '취소' }] }); });
    const parent = page.getByRole('dialog', { name: '부모 창' });
    await parent.getByRole('button', { name: '취소', exact: true }).focus(); await page.keyboard.press('Tab');
    assert.equal(await page.evaluate(() => document.activeElement.title), '닫기');
    await page.keyboard.press('Shift+Tab'); assert.equal(await page.evaluate(() => document.activeElement.textContent), '취소');
    await parent.locator('input').focus();
    await parent.locator('input').dispatchEvent('keydown', { key: 'Enter', isComposing: true });
    assert.equal(await page.evaluate(() => window.uiCount), 0);
    await openUi(page, (ui) => { ui.alertDialog('안내 창', '검증').then(() => { window.alertDone = true; }); });
    await page.keyboard.press('Escape');
    assert.equal(await page.evaluate(() => window.alertDone), true);
    assert.equal(await page.evaluate(() => document.activeElement.id), 'parent-input');
  });
  await test('비동기 확인: 완료 대기·중복 실행 차단·false 유지·오류 표시', async (page) => {
    await openUi(page, (ui) => { window.actionCount = 0; ui.openDialog({ title: '비동기 창', body: ui.el('input'), buttons: [{ label: '확인', primary: true, action: () => { window.actionCount++; return new Promise((resolve, reject) => { window.resolveAction = resolve; window.rejectAction = reject; }); } }] }); });
    const dialog = page.getByRole('dialog', { name: '비동기 창' });
    await dialog.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await dialog.getAttribute('aria-busy'), 'true');
    await page.keyboard.press('Enter'); assert.equal(await page.evaluate(() => window.actionCount), 1);
    await page.evaluate(() => window.resolveAction(false));
    await page.waitForFunction(() => !document.querySelector('[aria-busy="true"]'));
    assert.equal(await dialog.count(), 1);
    await dialog.getByRole('button', { name: '확인', exact: true }).click();
    await page.evaluate(() => window.rejectAction(new Error('검증 오류')));
    await dialog.getByRole('alert').waitFor({ state: 'visible' });
    assert.match(await dialog.getByRole('alert').textContent(), /검증 오류/);
    await page.keyboard.press('Escape'); assert.equal(await dialog.count(), 0);
  });
  await test('메뉴 키보드 탐색·하위 메뉴·선택·입력 칸 방향키 보존', async (page) => {
    await openUi(page, (ui) => { window.menuValue = ''; ui.openMenu({ x: 200, y: 200 }, [{ label: '사용 불가', disabled: true }, { label: '첫째', action: () => { window.menuValue = '첫째'; } }, { label: '하위', submenu: [{ label: '선택', action: () => { window.menuValue = '하위 선택'; } }] }]); });
    assert.equal(await page.evaluate(() => document.activeElement.textContent.trim()), '첫째');
    await page.keyboard.press('End'); await page.keyboard.press('ArrowRight');
    assert.equal(await page.evaluate(() => document.activeElement.textContent.trim()), '선택');
    await page.keyboard.press('ArrowLeft'); assert.match(await page.evaluate(() => document.activeElement.textContent), /하위/);
    await page.keyboard.press('ArrowRight'); await page.keyboard.press('Enter');
    assert.equal(await page.evaluate(() => window.menuValue), '하위 선택');
    await openUi(page, (ui) => { ui.openMenu({ x: 200, y: 200 }, [{ node: ui.el('input', { id: 'menu-input', value: 'abc' }) }, { label: '선택' }]); });
    await page.locator('#menu-input').evaluate((el) => el.setSelectionRange(2, 2));
    await page.keyboard.press('ArrowLeft'); assert.equal(await page.locator('#menu-input').evaluate((el) => el.selectionStart), 1);
    await page.keyboard.press('Escape'); assert.equal(await page.locator('.menu').count(), 0);
  });
  await test('피벗 다른 시트 이동: 값·서식·참조·차트·슬라이서 및 실행 취소/다시 실행', async (page) => {
    const before = await pivotFixture(page);
    await movePivot(page, '대상!$H$5');
    const after = await page.evaluate(() => {
      const wb = window.tabula.wb(); const d = wb.sheets[2].pivot;
      return { source: wb.sheets[1].pivot, def: d, style: wb.styleAt(2, 5, 8), raw: wb.getCell(0, 7, 0).raw, value: wb.getValue(0, 7, 0), total: wb.getValue(0, 8, 0), chart: wb.sheets[0].charts[0].pivot, slicer: wb.sheets[0].slicers[0].source.pivots[0] };
    });
    assert.equal(after.source, null); assert.equal(after.def.top, 4); assert.equal(after.def.left, 7);
    assert.equal(after.style.fill, '#ffff00'); assert.equal(after.value, before.value); assert.equal(after.total, before.total);
    assert.match(after.raw, /대상.*I6/); assert.deepEqual(after.chart, { sheet: '대상', name: '검증피벗' }); assert.deepEqual(after.slicer, after.chart);
    assert.deepEqual(await page.evaluate((cells) => cells.map(([r, c]) => window.tabula.wb().getCell(2, 4 + r, 7 + c)?.raw ?? ''), before.cells), before.cells.map((x) => x[2]));
    await page.keyboard.press('Control+z');
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets[1].pivot.name), '검증피벗');
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets[2].pivot), null);
    await page.keyboard.press('Control+y'); assert.equal(await page.evaluate(() => window.tabula.wb().sheets[2].pivot.top), 4);
  });
  await test('피벗 새 워크시트 이동과 실행 취소', async (page) => {
    await pivotFixture(page); await movePivot(page);
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets.length), 4);
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets[3].pivot.top), 0);
    await page.keyboard.press('Control+z'); assert.equal(await page.evaluate(() => window.tabula.wb().sheets.length), 3);
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets[1].pivot.name), '검증피벗');
  });
  await test('피벗 겹치는 자체 영역 이동과 다른 데이터 덮어쓰기 차단', async (page) => {
    const before = await pivotFixture(page); await movePivot(page, '보고서!C3');
    assert.deepEqual(await page.evaluate((cells) => cells.map(([r, c]) => window.tabula.wb().getCell(1, 2 + r, 2 + c)?.raw ?? ''), before.cells), before.cells.map((x) => x[2]));
    await page.evaluate(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setInput(2, 9, 9, '보존')); });
    const dialog = await movePivot(page, '대상!J10');
    assert.match(await dialog.textContent(), /데이터가 있습니다/);
    assert.equal(await page.evaluate(() => window.tabula.wb().getValue(2, 9, 9)), '보존');
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets[1].pivot.left), 2);
  });
  await test('피벗 이동이 동적 배열 분산 결과와 앵커 수식을 보존', async (page) => {
    await pivotFixture(page);
    await page.evaluate(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setInput(2, 4, 7, '=SEQUENCE(10,3)')); });
    const dialog = await movePivot(page, '대상!I6');
    assert.match(await dialog.textContent(), /동적 배열.*분산/);
    assert.deepEqual(await page.evaluate(() => {
      const wb = window.tabula.wb();
      return [wb.getValue(2, 4, 7), wb.getValue(2, 5, 8), wb.getCell(2, 4, 7).raw, wb.sheets[1].pivot.name, wb.sheets[2].pivot];
    }), [1, 5, '=SEQUENCE(10,3)', '검증피벗', null]);
  });
  await test('피벗 이동 후 일반 차트의 셀 참조 계열 이름·값 및 실행 취소 보존', async (page) => {
    await pivotFixture(page);
    const before = await page.evaluate(async () => {
      const { chartModelData } = await import('/src/chart.js');
      const wb = window.tabula.wb();
      wb.transact(() => wb.setSheetProp(0, 'charts', [{ id: 'name-test', type: 'column', sheet: '보고서', x: 500, y: 20, w: 300, h: 200,
        series: [{ name: { ref: { sheet: '보고서', r1: 2, c1: 2, r2: 2, c2: 2 } }, cat: { sheet: '보고서', r1: 3, c1: 1, r2: 4, c2: 1 }, val: { sheet: '보고서', r1: 3, c1: 2, r2: 4, c2: 2 } }] }]));
      return chartModelData(wb, 0, wb.sheets[0].charts[0]).series[0];
    });
    await movePivot(page, '대상!H5');
    const after = await page.evaluate(async () => {
      const { chartModelData } = await import('/src/chart.js'); const wb = window.tabula.wb(); const ch = wb.sheets[0].charts[0];
      return { data: chartModelData(wb, 0, ch).series[0], ref: ch.series[0].name.ref };
    });
    assert.deepEqual(after.data, before);
    assert.deepEqual(after.ref, { sheet: '대상', r1: 4, c1: 8, r2: 4, c2: 8 });
    await page.keyboard.press('Control+z');
    assert.equal(await page.evaluate(() => window.tabula.wb().sheets[0].charts[0].series[0].name.ref.sheet), '보고서');
  });
  await test('웹 가져오기 모달 위 서버 암호 입력 후 재시도', async (page) => {
    let requests = 0;
    await page.route('**/api/fetch?**', (route) => {
      requests++;
      return route.request().headers()['x-tabula-token'] === 'test-token'
        ? route.fulfill({ status: 200, contentType: 'text/plain; charset=utf-8', body: '항목,금액\n검증,10' })
        : route.fulfill({ status: 401, json: { error: '서버 암호가 필요합니다' } });
    });
    await page.evaluate(() => window.tabula.run('webData'));
    const dialog = page.getByRole('dialog', { name: '웹에서', exact: true });
    await dialog.locator('input[type=url]').fill('https://example.com/test.csv');
    await dialog.getByRole('button', { name: '이동', exact: true }).click();
    const auth = page.getByRole('dialog', { name: '서버 암호', exact: true });
    await auth.locator('input').fill('test-token'); await auth.getByRole('button', { name: '확인', exact: true }).click();
    await dialog.locator('.web-item').waitFor();
    assert.equal(requests, 2); assert.equal(await auth.count(), 0);
    assert.match(await dialog.locator('.web-preview').textContent(), /검증10/);
  }, { server: true });
  await test('네이버 중계 인증 헤더·잘못된 암호 재입력·결과 표시', async (page) => {
    const tokens = [];
    await page.route('**/api/naver/keywordstool?**', (route) => {
      const headers = route.request().headers();
      tokens.push(headers['x-tabula-token'] ?? '');
      assert.equal(headers['x-customer'], '123');
      return headers['x-tabula-token'] === 'naver-test-token'
        ? route.fulfill({ json: { keywordList: [{ relKeyword: '연관검색어', monthlyPcQcCnt: 100, monthlyMobileQcCnt: 200 }] } })
        : route.fulfill({ status: 401, json: { error: '서버 암호가 필요합니다' } });
    });
    await page.evaluate(async () => {
      const { newShape } = await import('/src/shapes.js');
      const t = window.tabula; const wb = t.wb();
      wb.transact(() => {
        wb.setInput(0, 5, 2, '검색어'); wb.setInput(0, 16, 2, '123'); wb.setInput(0, 17, 2, 'dummy-key'); wb.setInput(0, 18, 2, 'dummy-secret');
        wb.setSheetProp(0, 'shapes', [{ ...newShape('rect', { x: 400, y: 100, w: 180, h: 60 }), id: 'naver-test', text: '연관검색어', macro: 'GetNaverAdKeyword' }]);
      });
      t.gv().renderAll();
    });
    await page.locator('[data-id="naver-test"]').first().click();
    const auth = page.getByRole('dialog', { name: '서버 암호', exact: true });
    await auth.locator('input').fill('wrong-token'); await auth.getByRole('button', { name: '확인', exact: true }).click();
    await page.waitForFunction(() => document.querySelector('[data-server-token] input')?.value === '');
    await auth.locator('input').fill('naver-test-token'); await auth.getByRole('button', { name: '확인', exact: true }).click();
    await page.waitForFunction(() => window.tabula.wb().getValue(0, 5, 6) === '연관검색어');
    assert.deepEqual(tokens, ['', 'wrong-token', 'naver-test-token']);
  }, { server: true });
  const bad = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ total: results.length, ok: results.length - bad, bad }));
  if (bad) process.exitCode = 1;
} finally { await browser.close(); }
