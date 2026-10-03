// 차트 옆 편집 버튼 겹침 회귀. 합성 문서·새 브라우저만 사용하고 외부 요청/API 쓰기를 막는다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium';
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const origin = new URL(url).origin;
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw new Error('합성 차트 검사는 로컬 서버에서만 실행하세요.');
const out = process.env.WIXEL_CHART_CONTROLS_OUT || 'D:/Codex/Temp/wixel-chart-controls/' + engine;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch(), results = [], assets = new Set(); let checks = 0;
const eq = (a, b, message) => { assert.deepEqual(a, b, message); checks++; };
const ok = (value, message) => { assert.ok(value, message); checks++; };
const object = (p, id) => p.locator('.pane-br .obj[data-id="' + id + '"]');
const side = p => object(p, 'back').locator('.ch-side');
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const saved = p => p.evaluate(() => JSON.stringify(tabula.wb().serialize()));
const depth = p => p.evaluate(() => tabula.wb().undoStack.length);
async function fixture(p, kind = 'chart') {
  await p.evaluate(kind => {
    const w = tabula.wb(), range = { r1: 0, c1: 0, r2: 3, c2: 1 };
    const chart = { id: 'back', x: 40, y: 40, w: 360, h: 230, z: 1, type: 'column', range, title: '뒤 차트', legend: 'b' };
    const front = { id: 'front', x: 380, y: 40, w: 300, h: 250, z: 2 };
    const s = { name: '합성 차트', cells: { '0,0': {raw:'분류'}, '0,1': {raw:'값'}, '1,0': {raw:'가'}, '1,1': {raw:'12'}, '2,0': {raw:'나'}, '2,1': {raw:'7'}, '3,0': {raw:'다'}, '3,1': {raw:'18'} }, charts: [chart] };
    if (kind === 'chart') s.charts.push({ ...chart, ...front, title: '앞 차트' });
    if (kind === 'picture') {
      const cv = document.createElement('canvas'); cv.width = 30; cv.height = 25;
      const c = cv.getContext('2d'); c.fillStyle = '#ff6600'; c.fillRect(0,0,30,25);
      s.images = [{ ...front, src: cv.toDataURL() }];
    }
    if (kind === 'shape') s.shapes = [{ ...front, kind: 'rect', fill: '#ff6600', stroke: '#880000', rot: 8 }];
    if (kind === 'slicer') {
      s.tables = [{ id: 'table', name: '합성표', ...range, header: true, filter: { criteria: {}, hidden: {} } }];
      s.slicers = [{ ...front, caption: '분류', source: { kind: 'table', table: '합성표', column: '분류' } }];
    }
    w.restore({ sheets: [s, { name: '빈 시트', cells: {} }] });
    tabula.switchSheet(1); tabula.switchSheet(0); tabula.gv().setZoom(100);
    w.undoStack = []; w.redoStack = [];
  }, kind);
  await raf(p);
  if (kind === 'picture') await p.waitForFunction(() => [...document.querySelectorAll('.obj.pic img')].every(im => im.complete && im.naturalWidth));
}
async function chooseBack(p) {
  await object(p, 'back').click({ position: { x: 25, y: 8 } });
  await side(p).waitFor(); await raf(p);
  eq(await p.evaluate(() => tabula.chartSel), 'back', '뒤 차트 실제 선택');
}
async function hitControls(p) {
  const hits = await side(p).locator('.ch-sb').evaluateAll(nodes => nodes.map(node => {
    const r = node.getBoundingClientRect(), hit = document.elementFromPoint(r.x + r.width / 2, r.y + r.height / 2);
    return { action: node.dataset.a, direct: hit?.closest('.ch-sb') === node, owner: hit?.closest('.obj')?.dataset.id };
  }));
  eq(hits.map(h => h.action), ['elements', 'styles', 'filter'], '세 편집 버튼');
  for (const hit of hits) { eq(hit.direct, true, hit.action + ' 클릭 지점 노출'); eq(hit.owner, 'back', hit.action + ' 선택 차트 소유'); }
  return hits;
}
async function contentTop(p) {
  return object(p, 'back').evaluate(n => {
    const r = n.getBoundingClientRect(), hit = document.elementFromPoint(r.right - 10, r.top + r.height / 2);
    return hit?.closest('.obj')?.dataset.id;
  });
}
async function menus(p) {
  for (const [action, panel] of [['elements', '.ce-pop'], ['styles', '.cs-pop .cs-tabs'], ['filter', '.cf-pop']]) {
    await side(p).locator('[data-a="' + action + '"]').click();
    await p.locator('#menuLayer ' + panel).waitFor();
    eq(await p.evaluate(() => tabula.chartSel), 'back', action + ' 메뉴가 뒤 차트를 유지');
    await p.keyboard.press('Escape'); await p.locator('#menuLayer > .menu').waitFor({ state: 'detached' });
  }
}
async function contextFront(p, id) {
  await object(p, id).click({ button: 'right', position: { x: id === 'front' ? 80 : 10, y: 4 } });
  const menu = p.locator('#menuLayer > .menu[data-level="0"]');
  await menu.getByRole('menuitem', { name: '맨 앞으로 가져오기', exact: true }).click();
  await p.locator('#menuLayer > .menu[data-level="1"]').getByRole('menuitem', { name: '맨 앞으로 가져오기', exact: true }).click();
  await raf(p);
}
async function test(name, run) {
  const context = await browser.newContext({ viewport: { width: 1400, height: 950 }, deviceScaleFactor: 1, serviceWorkers: 'block' });
  const p = await context.newPage(), errors = [], writes = [], start = checks;
  p.setDefaultTimeout(15000); p.on('pageerror', e => errors.push(e.message)); p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel.mobile-work.v1', 'off'); });
  await context.route('**/*', r => {
    const request = r.request(), target = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method() + ' ' + target.pathname); return r.abort(); }
    return target.origin === origin && !/^\/api(?:\/|$)/.test(target.pathname) ? r.continue() : r.abort();
  });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await p.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
    for (const asset of await p.locator('script[src]').evaluateAll(ns => ns.map(n => n.getAttribute('src')))) assets.add(asset);
    await run(p); eq(errors, [], '페이지 오류 없음'); eq(writes, [], '원격 쓰기 없음');
    results.push({ name, ok: true, checks: checks - start, pageErrors: errors, blockedWrites: writes }); console.log('OK ' + name + ' (' + (checks - start) + ')');
  } catch (e) {
    await p.screenshot({ path: out + '/failure-' + results.length + '.png' }).catch(() => {});
    results.push({ name, ok: false, error: e.stack, errors, writes, checks: checks - start }); console.error('NG ' + name + ': ' + e.stack);
  } finally { await context.close(); }
}
try {
  for (const [kind, label] of [['chart','앞 차트와'], ['picture','그림과'], ['shape','회전 도형과'], ['slicer','슬라이서와']]) {
    await test(label + ' 겹쳐도 세 메뉴 사용·개체 순서 보존', async p => {
      await fixture(p, kind); await chooseBack(p);
      const before = await saved(p), undo = await depth(p);
      eq(await contentTop(p), 'front', '선택 차트 본문은 앞 개체 아래 유지');
      await hitControls(p); await menus(p); await hitControls(p);
      eq(await contentTop(p), 'front', '메뉴 후에도 본문 겹침 순서 유지');
      eq(await saved(p), before, '메뉴 열기·닫기가 셀/개체 모델을 바꾸지 않음'); eq(await depth(p), undo, '메뉴 사용이 Undo를 추가하지 않음');
      await p.screenshot({ path: out + '/' + kind + '-overlap.png' });
      const grid = await p.locator('#gridView').boundingBox();
      await p.mouse.click(grid.x + 100, grid.y + 520);
      eq(await p.evaluate(() => tabula.chartSel), null, '빈 셀 클릭으로 선택 해제'); eq(await p.locator('.ch-side').count(), 0, '선택 해제시 버튼 숨김');
      eq(await saved(p), before, '선택 해제도 모델 불변'); eq(await depth(p), undo, '선택 해제도 Undo 불변');
    });
  }
  await test('명시적 맨 앞으로·Undo/Redo는 본문 순서를 정상 변경', async p => {
    await fixture(p); await chooseBack(p); const before = await saved(p);
    await contextFront(p, 'back'); eq(await contentTop(p), 'back', '명시 명령은 차트 본문을 앞으로 이동'); eq(await depth(p), 1, '순서 변경 한 번 Undo');
    await hitControls(p);
    await p.evaluate(() => tabula.run('undo')); await raf(p);
    eq(await saved(p), before, 'Undo가 원래 문서 복원'); eq(await contentTop(p), 'front', 'Undo가 본문 순서 복원');
    await p.evaluate(() => tabula.run('redo')); await raf(p); eq(await contentTop(p), 'back', 'Redo가 앞으로 이동 복원');
    await contextFront(p, 'front'); eq(await contentTop(p), 'front', '다른 차트의 앞으로 명령도 유지');
    await chooseBack(p); await hitControls(p); eq(await contentTop(p), 'front', '재선택은 본문을 앞으로 보내지 않음');
  });
  await test('배율·스크롤·복수 선택에서도 버튼과 개체 선택 구분', async p => {
    await fixture(p); await chooseBack(p); const before = await saved(p);
    const slider = p.locator('#zoomSlider');
    await slider.evaluate(el => { el.value = '150'; el.dispatchEvent(new Event('input', { bubbles:true })); el.dispatchEvent(new Event('change', { bubbles:true })); });
    await raf(p); await hitControls(p); eq(await contentTop(p), 'front', '150% 본문 순서');
    await p.evaluate(() => tabula.gv().setScroll(20, 20)); await raf(p); await hitControls(p);
    await object(p, 'front').click({ position: { x: 150, y: 160 }, modifiers: ['Control'] });
    eq(await p.locator('.ch-side').count(), 0, '복수 개체 선택은 단일 차트 버튼 숨김');
    eq(await p.evaluate(() => tabula.gv().host.state().objMulti.size), 1, 'Ctrl 클릭으로 두 개체 선택');
    eq(await depth(p), 0, '배율·스크롤·복수 선택은 Undo 불변');
    const after = JSON.parse(await saved(p)), original = JSON.parse(before);
    eq(after.sheets[0].charts, original.sheets[0].charts, '배율·스크롤 중 차트 위치/순서 불변'); eq(after.sheets[0].cells, original.sheets[0].cells, '셀 원본 불변');
  });
} finally {
  await browser.close();
  const result = { engine, url, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
  await writeFile(out + '/result.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify({engine,assets:result.assets,cases:result.cases,passed:result.passed,checks,out}));
  if (result.passed !== result.cases) process.exitCode = 1;
}