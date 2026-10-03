// 소스 UI 모듈의 미니 도구 모음/메뉴 결합 계약. 실제앱 선택범위 검증은 context-menu.mjs에서 수행한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
assert.ok(['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname));
const browser = await chromium.launch(), results = [];
async function fixture(page, { x = 500, y = 300, readonly = false, rows = 10 } = {}) {
  await page.evaluate(async ({ x, y, readonly, rows }) => {
    const ui = await import('/src/ui.js'), { createContextMiniToolbar } = await import('/src/context-mini-toolbar.js');
    const model = window.miniModel = { font: '맑은 고딕', size: 11, bold: false, italic: false, underline: false, align: 'left', calls: [], readonly };
    const bar = createContextMiniToolbar({ getStyle: () => model, readonly: () => model.readonly,
      onCommand: (cmd, value) => { model.calls.push([cmd, value]); if (['bold', 'italic', 'underline'].includes(cmd)) model[cmd] = !model[cmd]; if (cmd === 'fontFamily') model.font = value; if (cmd === 'fontSize') model.size = value; },
      openNamedMenu: (name, anchor) => { model.calls.push(['menu', name]); ui.openMenu(anchor, [{ label: '색 적용', action: () => model.calls.push(['color', 'red']) }]); },
    });
    window.miniBar = bar; window.miniMenu = ui.openMenu({ x, y }, [
      { label: '복사', accessKey: 'c', action: () => model.calls.push(['copy']) },
      { label: '숨기기', accessKey: 'h', action: () => model.calls.push(['hide']) },
      ...Array.from({ length: rows - 2 }, (_, i) => ({ label: `합성 항목 ${i + 1}`, accessKey: i === rows - 3 ? 'd' : undefined, action: () => model.calls.push(['item', i]) })),
    ], { minWidth: 260, toolbar: bar });
  }, { x, y, readonly, rows });
}
async function test(name, fn, viewport = { width: 1440, height: 1000 }) {
  const context = await browser.newContext({ viewport }), page = await context.newPage(), errors = [], writes = [];
  page.setDefaultTimeout(10000); page.on('pageerror', (e) => errors.push(e.message));
  await context.route('**/*', (route) => { const req = route.request(); if (!['GET', 'HEAD'].includes(req.method())) { writes.push(req.method() + ' ' + req.url()); return route.abort(); } return route.continue(); });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => !!window.tabula);
    await fn(page); assert.deepEqual(errors, []); assert.deepEqual(writes, []); results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, errors, writes }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
const state = (p) => p.evaluate(() => window.miniModel);
try {
  await test('미니 굵게·기울임꼴 연속 적용은 메뉴 유지 및 눌림 상태 갱신', async (p) => {
    await fixture(p); const bold = p.locator('[data-mini-command="bold"]');
    await bold.click(); assert.equal(await bold.getAttribute('aria-pressed'), 'true');
    await p.locator('[data-mini-command="italic"]').click(); await bold.click();
    assert.equal(await bold.getAttribute('aria-pressed'), 'false'); assert.equal(await p.getByRole('menu').count(), 1); assert.equal(await p.getByRole('toolbar', { name: '미니 서식 도구 모음' }).count(), 1);
    assert.deepEqual((await state(p)).calls.map((x) => x[0]), ['bold', 'italic', 'bold']);
  });
  await test('글꼴·크기 입력은 Enter 적용 후 blur 때 중복 명령 없음', async (p) => {
    await fixture(p); const bar = p.getByRole('toolbar', { name: '미니 서식 도구 모음' });
    await bar.getByLabel('글꼴', { exact: true }).fill('Arial'); await p.keyboard.press('Enter');
    await bar.getByLabel('글꼴 크기', { exact: true }).fill('18'); await p.keyboard.press('Enter');
    await p.locator('[data-mini-command="bold"]').click();
    assert.deepEqual((await state(p)).calls.map((a) => a.map((v) => v ?? null)), [['fontFamily', 'Arial'], ['fontSize', 18], ['bold', null]]);
    assert.equal(await p.getByRole('menu').count(), 1);
  });
  await test('기존 색 메뉴 연결은 보조 막대까지 교체하며 선택 명령 정확', async (p) => {
    await fixture(p); const anchor = await p.locator('[data-mini-command="fillColor"]').boundingBox(); await p.locator('[data-mini-command="fillColor"]').click();
    const menu = await p.getByRole('menu').boundingBox(); assert.ok(Math.abs(menu.x - anchor.x) <= 1 && Math.abs(menu.y - anchor.y - anchor.height - 2) <= 1, '제거된 미니 단추의 위치에서 새 팔레트가 열려야 함');
    assert.equal(await p.locator('.context-mini-toolbar').count(), 0); await p.getByRole('menuitem', { name: '색 적용' }).click();
    assert.deepEqual((await state(p)).calls, [['menu', 'fillColor'], ['color', 'red']]); assert.equal(await p.getByRole('menu').count(), 0);
  });
  await test('Tab/ShiftTab 도구-메뉴 전환, 도구 방향키, Escape 함께 종료', async (p) => {
    await fixture(p); await p.keyboard.press('Tab'); assert.equal(await p.evaluate(() => document.activeElement.getAttribute('aria-label')), '글꼴');
    await p.keyboard.press('Shift+Tab'); assert.equal(await p.evaluate(() => document.activeElement.textContent.trim()), '합성 항목 8');
    await p.keyboard.press('Tab'); await p.keyboard.press('Tab'); await p.keyboard.press('Tab'); await p.keyboard.press('Tab'); await p.keyboard.press('Tab');
    assert.equal(await p.evaluate(() => document.activeElement.dataset.miniCommand), 'growFont'); await p.keyboard.press('ArrowRight'); assert.equal(await p.evaluate(() => document.activeElement.dataset.miniCommand), 'shrinkFont');
    await p.keyboard.press('Escape'); assert.equal(await p.getByRole('menu').count(), 0); assert.equal(await p.locator('.context-mini-toolbar').count(), 0);
  });
  await test('메뉴 명시 C는 미니 자동 키보다 우선하며 미니 힌트·배지는 숨김', async (p) => {
    await fixture(p); await p.keyboard.press('Alt'); assert.equal(await p.locator('.access-key-badge').count(), 10);
    assert.equal(await p.locator('.context-mini-toolbar .access-key-hint').count(), 0);
    const miniRect = await p.locator('.context-mini-toolbar').boundingBox();
    assert.equal(await p.locator('.access-key-badge').evaluateAll((nodes,b) => nodes.filter(n => { const r=n.getBoundingClientRect(),x=r.x+r.width/2,y=r.y+r.height/2;return x>=b.x&&x<=b.x+b.width&&y>=b.y&&y<=b.y+b.height; }).length, miniRect), 0);
    assert.equal(await p.locator('.context-mini-toolbar [data-resolved-access-key="c"]').count(), 0);
    await p.keyboard.press('c'); assert.deepEqual((await state(p)).calls, [['copy']]); assert.equal(await p.locator('.context-mini-toolbar').count(), 0);
  });
  await test('읽기전용 비활성·해제 refresh·동적 권한 거절', async (p) => {
    await fixture(p, { readonly: true }); assert.equal(await p.locator('.context-mini-toolbar button:enabled,.context-mini-toolbar input:enabled').count(), 0);
    await p.evaluate(() => { window.miniModel.readonly = false; window.miniBar.refresh(); }); await p.locator('[data-mini-command="bold"]').click();
    await p.evaluate(() => { window.miniModel.readonly = true; }); await p.locator('[data-mini-command="italic"]').click(); assert.deepEqual((await state(p)).calls.map((a) => a.map((v) => v ?? null)), [['bold', null]]);
  });
  await test('외부 메뉴 DOM 제거와 바깥 클릭에서 보조 막대 잔류 없음', async (p) => {
    await fixture(p); await p.evaluate(() => window.miniMenu.remove()); await p.waitForFunction(() => !document.querySelector('.context-mini-toolbar'));
    await fixture(p); await p.evaluate(async () => { const ui = await import('/src/ui.js'); window.miniMenu = ui.openMenu({ x: 450, y: 350 }, [{ label: '재사용 메뉴' }], { toolbar: window.miniBar }); await new Promise(requestAnimationFrame); });
    assert.equal(await p.locator('.context-mini-toolbar').count(), 1, '같은 toolbar 재사용 뒤 이전 메뉴 observer가 제거하면 안 됨');
    await fixture(p); await p.locator('.titlebar').click({ position: { x: 5, y: 5 } }); assert.equal(await p.locator('.context-mini-toolbar').count(), 0); assert.equal(await p.getByRole('menu').count(), 0);
  });
  await test('좁은 화면의 스크롤 밖 명시 D는 자동 미니 명령 대신 해당 메뉴 실행', async (p) => {
    await fixture(p, { x: 300, y: 220, rows: 22 }); await p.keyboard.press('Alt+d');
    assert.deepEqual((await state(p)).calls, [['item', 19]]); assert.equal(await p.locator('.context-mini-toolbar').count(), 0);
  }, { width: 320, height: 240 });
  for (const viewport of [{ width: 1440, height: 1000 }, { width: 320, height: 240 }, { width: 240, height: 180 }]) {
    await test(`화면 ${viewport.width}×${viewport.height}: 네 모서리 비겹침·잘림 없음 및 resize`, async (p) => {
      for (const [x, y] of [[2, 2], [viewport.width - 2, 2], [2, viewport.height - 2], [viewport.width - 2, viewport.height - 2]]) {
        await fixture(p, { x, y, rows: 22 });
        const rects = await p.evaluate(() => [window.miniBar, window.miniMenu].map((n) => ({ ...n.getBoundingClientRect().toJSON(), scroll: n.scrollHeight > n.clientHeight })));
        for (const rect of rects) assert.ok(rect.left >= 0 && rect.top >= 0 && rect.right <= viewport.width && rect.bottom <= viewport.height, JSON.stringify(rect));
        assert.ok(rects[0].bottom < rects[1].top, '미니와 메뉴 겹침');
        if (viewport.height < 300) assert.equal(rects[1].scroll, true);
        await p.keyboard.press('Alt'); assert.equal(await p.locator('.context-mini-toolbar [data-resolved-access-key="d"]').count(), 0, '스크롤 밖 메뉴 D 예약'); await p.keyboard.press('Escape');
      }
      await p.setViewportSize({ width: 360, height: 280 });
      assert.ok((await p.locator('.context-mini-toolbar').boundingBox()).x >= 0);
    }, viewport);
  }
} finally { await browser.close(); }
console.log(JSON.stringify({ sourceOnly: true, total: results.length, passed: results.filter((r) => r.ok).length, results }, null, 2));
if (results.some((r) => !r.ok)) process.exitCode = 1;
