// Isolated palette cancellation/permission regression; no remote document writes.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const origin = new URL(url).origin;
if (!['localhost', '127.0.0.1'].includes(new URL(url).hostname)) throw new Error('로컬 합성 검증 서버를 사용하세요.');
const browser = await chromium.launch(), results = [];
async function test(name, fn) {
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 } });
  const page = await context.newPage(), errors = [], blockedWrites = [], unexpectedReads = [];
  page.setDefaultTimeout(15000);
  page.on('pageerror', error => errors.push(error.message));
  await context.route('**/*', route => {
    const request = route.request(), u = new URL(request.url());
    if (!['GET', 'HEAD'].includes(request.method())) { blockedWrites.push(request.method() + ' ' + u.pathname); return route.abort(); }
    if (u.origin === origin && !u.pathname.startsWith('/api/')) return route.continue();
    if (u.origin === origin && u.pathname === '/api/health') return route.fulfill({ status: 503, json: { ok: false } });
    unexpectedReads.push(u.origin + u.pathname); return route.abort();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => window.tabula?.wb());
    await page.evaluate(() => {
      const w = tabula.wb();
      w.restore({ sheets: [{ name: '합성 A', cells: { '0,0': { raw: '보존' } } }, { name: '합성 B', cells: { '0,0': { raw: '보존' } } }] });
      tabula.switchSheet(0); w.undoStack = []; w.redoStack = []; tabula.gv().layout(); tabula.run('drawingPalette');
    });
    await fn(page);
    assert.deepEqual(errors, []); assert.deepEqual(blockedWrites, []); assert.deepEqual(unexpectedReads, []);
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) {
    results.push({ name, ok: false, error: error.message, errors, blockedWrites, unexpectedReads }); console.error('NG ' + name + ': ' + error.stack);
  } finally { await context.close(); }
}
const state = page => page.evaluate(() => ({ shapes: tabula.wb().sheets.map(s => structuredClone(s.shapes || [])), cells: tabula.wb().sheets.map(s => [...s.cells]), undo: tabula.wb().undoStack.length }));
async function beginStroke(page) {
  await page.mouse.move(280, 360); await page.mouse.down(); await page.mouse.move(400, 420, { steps: 6 });
}
try {
  await test('사각형 드래그 중 시트 전환은 원래 시트 임시 도형과 Undo를 남기지 않음', async page => {
    const before = await state(page);
    await page.getByRole('button', { name: '사각형', exact: true }).click(); await beginStroke(page);
    assert.equal((await state(page)).shapes[0].length, 1, '검증 전에 사각형 임시 도형이 실제 생성되어야 함');
    await page.evaluate(() => tabula.switchSheet(1)); await page.mouse.up();
    assert.deepEqual(await state(page), before);
    await page.evaluate(() => tabula.switchSheet(0)); assert.deepEqual(await state(page), before);
  });
  await test('펜 드래그 도중 최종본 전환은 확정하지 않고 새 편집 가능한 획은 정상 Undo', async page => {
    const before = await state(page); await beginStroke(page);
    // Synthetic policy transition while a real mouse drag is in progress.
    await page.evaluate(() => { tabula.wb().props = { ...tabula.wb().props, markedFinal: true }; });
    await page.mouse.up(); assert.deepEqual(await state(page), before);
    assert.equal(await page.evaluate(() => !!tabula.wb().props.markedFinal), true);
    await page.evaluate(() => { tabula.wb().props = { ...tabula.wb().props, markedFinal: undefined }; });
    await page.getByRole('button', { name: '펜', exact: true }).click(); await beginStroke(page); await page.mouse.up();
    const drawn = await state(page); assert.equal(drawn.shapes[0].length, 1); assert.equal(drawn.shapes[0][0].ink, true); assert.equal(drawn.undo, before.undo + 1);
    await page.evaluate(() => tabula.run('undo')); assert.deepEqual(await state(page), before);
  });
  await test('획 지우개는 일반 도형을 보존하고 ink 획만 삭제·Undo 복원', async page => {
    await page.evaluate(() => {
      const w = tabula.wb();
      w.setSheetProp(0, 'shapes', [{ id: 'ordinary', kind: 'rect', x: 200, y: 120, w: 180, h: 80, fill: '#4472c4', text: '일반 도형' }]);
      w.undoStack = []; w.redoStack = []; tabula.gv().renderAll();
    });
    await page.getByRole('button', { name: '획 지우개', exact: true }).click(); const before = await state(page);
    await page.locator('.obj[data-id="ordinary"]').first().click(); assert.deepEqual(await state(page), before);
    await page.getByRole('button', { name: '펜', exact: true }).click();
    await page.mouse.move(560, 490); await page.mouse.down(); await page.mouse.move(680, 540, { steps: 6 }); await page.mouse.up();
    const drawn = await state(page), ink = drawn.shapes[0].find(o => o.ink); assert.ok(ink); assert.equal(drawn.shapes[0].length, 2);
    await page.getByRole('button', { name: '획 지우개', exact: true }).click(); await page.locator(`.obj[data-id="${ink.id}"]`).first().click();
    const erased = await state(page); assert.equal(erased.shapes[0].length, 1); assert.equal(erased.shapes[0][0].id, 'ordinary'); assert.equal(erased.undo, drawn.undo + 1);
    await page.evaluate(() => tabula.run('undo')); assert.deepEqual(await state(page), drawn);
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, passed: results.filter(r => r.ok).length, failed: results.filter(r => !r.ok).length, results }, null, 2));
if (results.some(r => !r.ok)) process.exitCode = 1;
