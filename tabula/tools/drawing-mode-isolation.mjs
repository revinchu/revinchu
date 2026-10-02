// 도형/테두리/서식 복사 모드 격리. 합성 문서, 실제 메뉴 클릭·마우스 드래그.
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5178/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 소스/번들 서버에서 실행하세요.');
const baseline = process.env.DRAWING_BASELINE_REF;
const baselineApp = baseline ? execFileSync('git', ['-c', `safe.directory=${fileURLToPath(new URL('../../', import.meta.url)).replace(/\\/g, '/').replace(/\/$/, '')}`, 'show', `${baseline}:tabula/src/app.js`], { cwd: fileURLToPath(new URL('../../', import.meta.url)), encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 }) : null;
const browser = await chromium.launch(), results = [];
const labels = { outline: '테두리 그리기', grid: '테두리 눈금 그리기', erase: '테두리 지우기' };
const target = { r1: 14, c1: 8, r2: 17, c2: 10 };
const secondTarget = { r1: 20, c1: 12, r2: 23, c2: 14 };
const flags = p => p.evaluate(() => ({ border: document.querySelector('#gridView').classList.contains('border-draw'), shape: document.querySelector('#gridView').classList.contains('drawing-mode'), painter: document.querySelector('#gridView').classList.contains('painting') }));
const snapshot = p => p.evaluate(() => ({ cells: tabula.wb().serialize().sheets.map(s => s.cells), undo: tabula.wb().undoStack.length }));
const shapes = p => p.evaluate(() => structuredClone(tabula.wb().sheets[0].shapes));
async function fixture(p, { borders = false, protectedSecond = false, object = false } = {}) {
  await p.evaluate(({ borders, protectedSecond, object }) => {
    const t = tabula, w = t.wb(), cells = { '0,0': { raw: '보존', style: { bold: true, fill: '#fff2cc' } }, '0,1': { raw: '=1+1', cached: 2 } };
    if (borders) for (const rg of [{ r1: 14, c1: 8, r2: 17, c2: 10 }, { r1: 20, c1: 12, r2: 23, c2: 14 }]) for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) cells[`${r},${c}`] = { raw: '', style: { bt: true, bb: true, bl: true, br: true, btc: '#123456' } };
    w.restore({ sheets: [{ name: '그리기 합성', cells, shapes: object ? [{ id: 'existing-shape', kind: 'rect', x: 130, y: 80, w: 140, h: 80, fill: '#4472c4', stroke: '#000000', strokeWidth: 1 }] : [] }, { name: '다른 시트', cells: { '0,0': { raw: '둘째' } }, ...(protectedSecond ? { protect: { on: true, allow: { selectLocked: true, selectUnlocked: true, formatCells: false, objects: false } } } : {}) }] });
    t.switchSheet(1); t.switchSheet(0); w.undoStack = []; w.redoStack = []; t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
  }, { borders, protectedSecond, object });
}
async function menu(p, label) {
  await p.locator('[data-ribbon-tab="home"]').click();
  await p.locator('[data-ribbon-menu="borders"] .caret').click();
  await p.getByRole('menuitem', { name: new RegExp(`^${label}(?:\\s|$)`) }).click();
}
async function border(p, mode) { await menu(p, labels[mode]); assert.equal((await flags(p)).border, true, '테두리 펜 활성화'); }
async function pick(p, label) {
  await p.locator('[data-ribbon-tab="insert"]').click();
  await p.locator('[data-ribbon-menu="shapes"]').click();
  await p.locator(`.shape-btn[title="${label}"]`).click();
}
async function drag(p, a, b) { await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(b.x, b.y, { steps: 8 }); await p.mouse.up(); }
async function cellPoint(p, r, c) { return p.evaluate(([r, c]) => { const x = tabula.gv().clientRect({ r1: r, c1: c, r2: r, c2: c }); return { x: (x.left + x.right) / 2, y: (x.top + x.bottom) / 2 }; }, [r, c]); }
async function select(p, rg = target) {
  await drag(p, await cellPoint(p, rg.r1, rg.c1), await cellPoint(p, rg.r2, rg.c2));
  assert.deepEqual(await p.evaluate(() => ({ ...tabula.sel })), rg, '실제 마우스가 의도한 셀 범위를 선택');
}
async function draw(p, label) {
  await pick(p, label); const box = await p.locator('#gridView').boundingBox();
  await drag(p, { x: box.x + 170, y: box.y + 100 }, { x: box.x + 330, y: box.y + 190 });
  assert.equal((await shapes(p)).length, 1, '실제 드래그로 도형 하나 생성');
}
async function undo(p) { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+z'); }
async function test(name, fn) {
  if (process.env.DRAWING_TEST_FILTER && !name.includes(process.env.DRAWING_TEST_FILTER)) return;
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } }), p = await context.newPage(), errors = [], writes = [];
  p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message)); p.on('dialog', d => d.dismiss());
  await context.route('**/*', route => {
    const req = route.request(), u = new URL(req.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.pathname); return route.abort(); }
    return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  if (baselineApp) await context.route('**/src/app.js', r => r.fulfill({ contentType: 'text/javascript', body: baselineApp }));
  try {
    await context.addInitScript(() => { window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    await fn(p); assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '외부 쓰기 시도');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (e) { results.push({ name, ok: false, error: e.message, errors, writes }); console.error('NG ' + name + ': ' + e.stack); }
  finally { await context.close(); }
}
try {
  for (const mode of ['outline', 'grid', 'erase']) {
    await test(`${mode} → 도형 완성 → 범위 선택은 셀·Undo 불변`, async p => {
      await fixture(p, { borders: mode === 'erase' }); const before = await snapshot(p); await border(p, mode); await draw(p, '사각형');
      const afterShape = await snapshot(p); assert.deepEqual(afterShape.cells, before.cells); assert.equal(afterShape.undo, 1);
      // 먼저 실제 선택 결과를 비교하여 모드 클래스만 바뀐 가짜 성공을 피한다.
      await select(p); assert.deepEqual(await snapshot(p), afterShape, '범위 선택이 셀 서식이나 Undo를 변경'); assert.equal((await flags(p)).border, false);
      await undo(p); assert.equal((await shapes(p)).length, 0); assert.deepEqual((await snapshot(p)).cells, before.cells);
    });
    await test(`${mode} → 자유곡선 취소 → 범위 선택은 셀·Undo 불변`, async p => {
      await fixture(p, { borders: mode === 'erase' }); const before = await snapshot(p); await border(p, mode); await pick(p, '자유형: 자유곡선'); await p.keyboard.press('Escape');
      await select(p); assert.deepEqual(await snapshot(p), before); assert.equal((await flags(p)).border, false); assert.equal((await shapes(p)).length, 0);
    });
    await test(`명시 ${mode} 펜은 한 번 적용·다음 선택 불변·재활성화·Undo`, async p => {
      await fixture(p, { borders: mode === 'erase' }); const before = await snapshot(p); await border(p, mode); await select(p);
      const after = await snapshot(p); assert.notDeepEqual(after.cells, before.cells); assert.equal(after.undo, 1);
      const style = await p.evaluate(() => tabula.wb().getCell(0, 14, 8)?.style ?? {});
      if (mode === 'erase') for (const side of ['bt', 'bb', 'bl', 'br']) assert.ok(!style[side], side);
      else { assert.equal(style.bt, true); assert.equal(style.bl, true); if (mode === 'grid') assert.equal(style.br, true); }
      const point = await cellPoint(p, 21, 13); await p.mouse.click(point.x, point.y);
      assert.deepEqual(await snapshot(p), after, '한 번 적용 뒤 다른 셀 클릭이 서식이나 Undo를 변경');
      await select(p, secondTarget); assert.deepEqual(await snapshot(p), after, '한 번 적용 뒤 다른 범위 드래그가 서식이나 Undo를 변경'); assert.equal((await flags(p)).border, false);
      await border(p, mode); await select(p, secondTarget); const second = await snapshot(p); assert.notDeepEqual(second.cells, after.cells); assert.equal(second.undo, 2); assert.equal((await flags(p)).border, false);
      await undo(p); assert.deepEqual(await snapshot(p), after, '재활성화 적용 한 번 Undo');
      await undo(p); assert.deepEqual(await snapshot(p), before, '첫 적용 한 번 Undo');
    });
  }
  await test('테두리 눈금 → 자유곡선 완성 → 다음 선택 불변', async p => {
    await fixture(p); await border(p, 'grid'); await draw(p, '자유형: 자유곡선'); const before = await snapshot(p); await select(p); assert.deepEqual(await snapshot(p), before); assert.equal((await shapes(p))[0].kind, 'scribble');
  });
  await test('일회성 모든 테두리는 기존 펜을 종료', async p => {
    await fixture(p); await border(p, 'outline'); await menu(p, '모든 테두리'); const before = await snapshot(p); assert.equal(before.undo, 1); await select(p); assert.deepEqual(await snapshot(p), before); assert.equal((await flags(p)).border, false);
  });
  await test('선 스타일 선택만으로 펜·서식이 바뀌지 않고 명시 펜에 선택 스타일 적용', async p => {
    await fixture(p); const before = await snapshot(p); await menu(p, '선 스타일...');
    await p.getByRole('menuitem', { name: /^보통 실선(?:\s|$)/ }).click();
    assert.deepEqual(await snapshot(p), before, '선 스타일 선택만으로 셀 변경');
    await select(p, secondTarget); assert.deepEqual(await snapshot(p), before, '선 스타일 선택 뒤 일반 드래그'); assert.equal((await flags(p)).border, false);
    await border(p, 'outline'); await select(p); const style = await p.evaluate(() => tabula.wb().getCell(0, 14, 8)?.style ?? {});
    assert.equal(style.bts, 'medium'); assert.equal(style.bls, 'medium'); assert.equal((await flags(p)).border, false);
    await undo(p); assert.deepEqual(await snapshot(p), before);
  });
  await test('모든 테두리와 주 버튼 반복은 현재 범위에만 한 번 적용', async p => {
    await fixture(p); const before = await snapshot(p); await select(p); await menu(p, '모든 테두리'); const first = await snapshot(p); assert.equal(first.undo, 1);
    await select(p, secondTarget); assert.deepEqual(await snapshot(p), first, '일반 메뉴 적용 뒤 다음 선택 불변');
    await p.locator('[data-ribbon-command="borderLast"]').click({ position: { x: 7, y: 8 } }); const second = await snapshot(p); assert.equal(second.undo, 2);
    const style = await p.evaluate(() => tabula.wb().getCell(0, 20, 12)?.style ?? {}); for (const k of ['bt', 'bb', 'bl', 'br']) assert.equal(style[k], true);
    const point = await cellPoint(p, 25, 16); await p.mouse.click(point.x, point.y); assert.deepEqual(await snapshot(p), second); assert.equal((await flags(p)).border, false);
    await undo(p); assert.deepEqual(await snapshot(p), first); await undo(p); assert.deepEqual(await snapshot(p), before);
  });
  await test('펜 → 시트 이동 → 다음 선택은 양 시트·Undo 불변', async p => {
    await fixture(p); await border(p, 'outline'); const before = await snapshot(p); await p.locator('.sheet-tab').filter({ hasText: '다른 시트' }).click(); assert.equal(await p.evaluate(() => tabula.si), 1); await select(p); assert.deepEqual(await snapshot(p), before); assert.equal((await flags(p)).border, false);
  });
  await test('펜 → 기존 개체 선택 → 다음 범위는 셀·Undo 불변', async p => {
    await fixture(p, { object: true }); await border(p, 'outline'); const before = await snapshot(p); await p.locator('.obj[data-id="existing-shape"]').first().click({ position: { x: 50, y: 30 } }); await select(p); assert.deepEqual(await snapshot(p), before); assert.equal((await flags(p)).border, false);
  });
  await test('도형 모드 → 테두리 펜은 도형 생성 없이 테두리만 적용', async p => {
    await fixture(p); await pick(p, '사각형'); await border(p, 'grid'); assert.equal((await flags(p)).shape, false); await select(p); assert.equal((await shapes(p)).length, 0); assert.equal((await snapshot(p)).undo, 1);
  });
  await test('서식 복사 → 테두리 펜 상호 배제', async p => {
    await fixture(p); await p.locator('[data-ribbon-command="painter"]').click(); assert.equal((await flags(p)).painter, true); await border(p, 'outline'); assert.equal((await flags(p)).painter, false); await select(p);
    const style = await p.evaluate(() => tabula.wb().getCell(0, 14, 8)?.style ?? {}); assert.equal(style.bt, true); assert.ok(!style.bold); assert.notEqual(style.fill, '#fff2cc');
  });
  await test('서식 복사 → 보호 시트 이동 → 드래그는 서식·Undo 불변', async p => {
    await fixture(p, { protectedSecond: true }); await p.locator('[data-ribbon-command="painter"]').click(); assert.equal((await flags(p)).painter, true); const before = await snapshot(p);
    await p.locator('.sheet-tab').filter({ hasText: '다른 시트' }).click(); await select(p); assert.deepEqual(await snapshot(p), before);
  });
  await test('일반 선택 드래그 중 UI로 펜 활성화해도 현재 드래그에는 미적용', async p => {
    await fixture(p); const before = await snapshot(p), a = await cellPoint(p, 14, 8), b = await cellPoint(p, 17, 10);
    // 마우스를 누른 채 실제 DOM 버튼의 키보드 활성화 경로로 메뉴를 연다.
    await p.locator('[data-ribbon-menu="borders"] .caret').click();
    await p.keyboard.press('Escape'); await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(b.x, b.y, { steps: 4 });
    await p.evaluate(() => document.querySelector('[data-ribbon-menu="borders"] .caret').click());
    await p.getByRole('menuitem', { name: /^테두리 그리기(?:\s|$)/ }).evaluate(e => e.click()); await p.mouse.up();
    assert.deepEqual(await snapshot(p), before, 'mousedown 때 펜이 없었던 선택'); assert.equal((await flags(p)).border, true);
  });
  await test('테두리 눈금 드래그 중 Escape 취소 뒤 mouseup은 셀·Undo 불변', async p => {
    await fixture(p); const before = await snapshot(p); await border(p, 'grid');
    const a = await cellPoint(p, 14, 8), b = await cellPoint(p, 17, 10);
    await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(b.x, b.y, { steps: 6 });
    await p.keyboard.press('Escape'); await p.mouse.up();
    assert.deepEqual(await snapshot(p), before); assert.equal((await flags(p)).border, false);
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, baseline: baseline || null, total: results.length, passed: results.filter(x => x.ok).length, failed: results.filter(x => !x.ok).length, results }, null, 2));
if (!results.length || results.some(x => !x.ok)) process.exitCode = 1;
