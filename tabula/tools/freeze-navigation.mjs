import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://localhost:5178/';
const output = process.env.FREEZE_NAV_SCREENSHOTS;
if (output) await mkdir(output, { recursive: true });
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1800, height: 1000 }, deviceScaleFactor: 1 });
const page = await context.newPage();
page.setDefaultTimeout(12000);
const results = [], pageErrors = [], writes = [], metrics = [];
page.on('pageerror', error => pageErrors.push(error.message));
await context.route('**/*', route => {
  const request = route.request(), url = new URL(request.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method() + ' ' + url.pathname); return route.abort(); }
  if (url.origin !== new URL(base).origin || url.pathname.startsWith('/api/')) return route.abort();
  return route.continue();
});
await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
const frame = () => page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));

async function reset(options = {}) {
  for (let i = 0; i < 3; i++) await page.keyboard.press('Escape');
  await page.setViewportSize(options.viewport || { width: 1800, height: 1000 });
  await page.evaluate(options => {
    const t = tabula;
    if (t.si !== 0) t.switchSheet(0);
    const cells = {};
    for (const [first, last] of [[0, 80], [290, 350], [875, 901]]) {
      for (let r = first; r < last; r++) for (let c = 0; c < 18; c++) cells[r + ',' + c] = { raw: `R${r + 1}C${c + 1}` };
    }
    const s = { name: '틀 고정 합성', defRowH: 22, defColW: 64, cells,
      freeze: options.freeze || { rows: 14, cols: 0 }, zoom: options.zoom || 100,
      hiddenRows: options.hiddenRows || {}, hiddenCols: options.hiddenCols || {},
      view: options.view || { top: 14, left: 0, r: 14, c: 5 } };
    const w = t.wb();
    w.restore({ sheets: [{ name: '초기화', cells: {} }, s] });
    w.undoStack = []; w.redoStack = [];
    // A newly created target sheet has no cached selection, so this exercises showSheetStart.
    t.switchSheet(1);
    window.__freezeFixture = JSON.stringify({ freeze: w.sheets[1].freeze, hiddenRows: w.sheets[1].hiddenRows, hiddenCols: w.sheets[1].hiddenCols });
  }, options);
  await page.evaluate(() => document.fonts.ready);
  await frame();
  await page.locator('#cellEditor').focus();
}
async function slider(value) {
  await page.locator('#zoomSlider').evaluate((node, value) => {
    node.value = String(value);
    node.dispatchEvent(new Event('input', { bubbles: true }));
    node.dispatchEvent(new Event('change', { bubbles: true }));
  }, value);
  await frame();
}
async function state() {
  return page.evaluate(() => {
    const t = tabula, g = t.gv(), active = { ...t.active }, s = t.wb().sheets[t.si];
    const area = { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
    const rect = g.clientRect(area), viewport = g.viewEl.getBoundingClientRect().toJSON();
    const candidates = [...document.querySelectorAll(`.c[data-r="${active.r}"][data-c="${active.c}"]`)];
    const node = candidates.find(node => {
      const r = node.getBoundingClientRect(), pane = g.panes.find(p => p.el.contains(node));
      if (!pane || getComputedStyle(pane.el).display === 'none') return false;
      const p = pane.el.getBoundingClientRect();
      return r.width > 0 && r.height > 0 && r.right > p.left && r.left < p.right && r.bottom > p.top && r.top < p.bottom;
    });
    const pane = node ? g.panes.find(p => p.el.contains(node)).el.getBoundingClientRect().toJSON() : null;
    const actual = node?.getBoundingClientRect();
    const x = ((actual || rect).left + (actual || rect).right) / 2, y = ((actual || rect).top + (actual || rect).bottom) / 2;
    const hit = g.hitTest(x, y);
    const borders = [...document.querySelectorAll('.sel-border')].some(n => {
      const r = n.getBoundingClientRect();
      return r.width > 0 && r.height > 0 && x >= r.left - 2 && x <= r.right + 2 && y >= r.top - 2 && y <= r.bottom + 2;
    });
    const frozenRows = [...new Set([...document.querySelectorAll('.c[data-c="5"]')].filter(n => {
      const p = g.panes.find(p => p.el.contains(n));
      if (!p || p.scrollY || getComputedStyle(p.el).display === 'none') return false;
      const r = n.getBoundingClientRect(), b = p.el.getBoundingClientRect();
      return r.height > 0 && r.bottom > b.top + 1 && r.top < b.bottom - 1;
    }).map(n => Number(n.dataset.r)))].sort((a, b) => a - b);
    return { active, selection: { ...t.sel }, rect, viewport, pane, dom: node?.getBoundingClientRect().toJSON() || null,
      hit: { r: hit.r, c: hit.c, zone: hit.zone }, border: borders, frozenRows,
      fr: g.fr, fc: g.fc, frozenTop: g.frozenTop, frozenLeft: g.frozenLeft,
      frozenH: g.frozenH, frozenW: g.frozenW, sx: g.sx, sy: g.sy, z: g.z,
      model: { freeze: s.freeze, hiddenRows: s.hiddenRows, hiddenCols: s.hiddenCols },
      unchanged: window.__freezeFixture === JSON.stringify({ freeze: s.freeze, hiddenRows: s.hiddenRows, hiddenCols: s.hiddenCols }) };
  });
}
async function visible(label, expected) {
  await frame();
  const s = await state();
  metrics.push({ label, ...s });
  if (expected) assert.deepEqual(s.active, expected, label + ': 활성 셀');
  assert.ok(s.dom && s.pane, label + ': 화면에 그려진 활성 셀');
  for (const edge of ['left', 'top', 'right', 'bottom']) assert.ok(Math.abs(s.dom[edge] - s.rect[edge]) <= 2, label + ': 실제 셀 DOM과 좌표 ' + edge);
  assert.ok(s.rect.top >= s.pane.top - 2 && s.rect.bottom <= s.pane.bottom + 2, label + ': 활성 셀의 세로 경계');
  assert.ok(s.rect.left >= s.pane.left - 2 && s.rect.right <= s.pane.right + 2, label + ': 활성 셀의 가로 경계');
  assert.ok(s.rect.top >= s.viewport.top - 2 && s.rect.bottom <= s.viewport.bottom + 2, label + ': 격자 세로 영역');
  assert.ok(s.rect.left >= s.viewport.left - 2 && s.rect.right <= s.viewport.right + 2, label + ': 격자 가로 영역');
  assert.deepEqual(s.hit, { ...s.active, zone: 'cell' }, label + ': 좌표와 클릭 대상');
  assert.ok(s.border, label + ': 선택 테두리');
  assert.ok(s.unchanged, label + ': 저장된 고정·숨김 설정 보존');
  return s;
}
async function keys(key, count, prefix) {
  for (let i = 0; i < count; i++) { await page.keyboard.press(key); await visible(prefix + ' ' + (i + 1)); }
}
async function screenshot(name) { if (output) await page.screenshot({ path: join(output, name + '.png') }); }
async function test(name, fn) {
  try { await fn(); results.push({ name, ok: true }); console.log('OK ' + name); }
  catch (error) { results.push({ name, ok: false, error: error.message }); console.error('NG ' + name + ': ' + error.message); }
}
const hiddenOrigin = {};
for (let r = 294; r <= 305; r++) hiddenOrigin[r] = true;
for (let r = 308; r <= 328; r++) hiddenOrigin[r] = true;
const originOptions = { freeze: { rows: 14, cols: 0, top: 294 }, hiddenRows: hiddenOrigin, zoom: 330,
  view: { top: 329, left: 0, r: 329, c: 5 } };
try {
  await page.goto(base, { waitUntil: 'commit', timeout: 60000 });
  await page.waitForFunction(() => !!window.tabula?.gv(), null, { timeout: 60000 });
  await test('저장된 고정 원점과 숨김 행으로 307·308행 및 330행을 표시', async () => {
    await reset(originOptions);
    const s = await visible('원점 복원', { r: 329, c: 5 });
    assert.equal(s.frozenTop, 294); assert.equal(s.fr, 308); assert.equal(s.frozenH, 44);
    assert.deepEqual(s.frozenRows, [306, 307]);
    await screenshot('origin-307-308-body-330');
  });
  await test('저장된 본문 시작 884행과 활성 셀 886행을 고정 원점과 별도로 복원', async () => {
    await reset({ ...originOptions, view: { top: 883, left: 0, r: 885, c: 4 } });
    const s = await visible('본문 884행', { r: 885, c: 4 });
    assert.deepEqual(s.frozenRows, [306, 307]);
    const bodyFirst = await page.evaluate(() => {
      const g = tabula.gv(), rows = [...document.querySelectorAll('.c[data-c="4"]')].filter(n => {
        const p = g.panes.find(p => p.el.contains(n)); if (!p?.scrollY || getComputedStyle(p.el).display === 'none') return false;
        const r = n.getBoundingClientRect(), b = p.el.getBoundingClientRect(); return r.height > 0 && r.bottom > b.top + 1 && r.top < b.bottom - 1;
      }).map(n => Number(n.dataset.r));
      return Math.min(...rows);
    });
    assert.equal(bodyFirst, 883);
  });
  await test('330% 고정 원점 문서에서 아래·위 방향키가 활성 셀을 계속 화면에 표시', async () => {
    await reset(originOptions);
    await keys('ArrowDown', 15, '원점 아래 이동');
    await keys('ArrowUp', 15, '원점 위 이동');
    await visible('원점 복귀', { r: 329, c: 5 });
    await page.keyboard.press('Control+Home'); await visible('본문 처음으로', { r: 329, c: 0 });
    for (const r of [307, 306, 293]) { await page.keyboard.press('ArrowUp'); await visible('고정 원점 위 이동 ' + (r + 1), { r, c: 0 }); }
    assert.equal((await state()).fr, 0);
    await page.evaluate(() => { tabula.switchSheet(0); tabula.switchSheet(1); });
    await visible('원점 위 선택의 시트 왕복', { r: 293, c: 0 });
    await page.evaluate(() => tabula.selectCell(306, 0));
    const restored = await visible('고정 영역 복귀', { r: 306, c: 0 });
    assert.equal(restored.fr, 308); assert.deepEqual(restored.frozenRows, [306, 307]);
  });
  await test('화면보다 큰 14행 고정은 일시 해제되어 F13과 아래 셀로 이동 가능', async () => {
    await reset({ zoom: 330, view: { top: 0, left: 0, r: 0, c: 5 } });
    assert.equal((await state()).fr, 0);
    await keys('ArrowDown', 16, '큰 고정 행');
    await visible('F17', { r: 16, c: 5 });
    await screenshot('oversized-rows-f17');
  });
  await test('화면보다 큰 고정 열은 오른쪽·왼쪽 방향키로 접근 가능', async () => {
    await reset({ freeze: { rows: 0, cols: 8 }, zoom: 330, view: { top: 0, left: 0, r: 2, c: 0 } });
    assert.equal((await state()).fc, 0);
    await keys('ArrowRight', 12, '큰 고정 열 오른쪽');
    await keys('ArrowLeft', 12, '큰 고정 열 왼쪽');
    await visible('A3 복귀', { r: 2, c: 0 });
  });
  await test('두 축의 큰 고정 영역에서도 세로·가로 이동과 클릭 대상이 일치', async () => {
    await reset({ freeze: { rows: 14, cols: 8 }, zoom: 330, view: { top: 0, left: 0, r: 0, c: 0 } });
    const s = await state(); assert.equal(s.fr, 0); assert.equal(s.fc, 0);
    await keys('ArrowDown', 16, '양축 아래'); await keys('ArrowRight', 12, '양축 오른쪽');
    const at = await visible('양축 끝', { r: 16, c: 12 });
    await page.mouse.click((at.dom.left + at.dom.right) / 2, (at.dom.top + at.dom.bottom) / 2);
    await visible('양축 클릭', { r: 16, c: 12 });
  });
  await test('확대 후 일시 해제된 고정 설정은 100% 축소 시 자동 복원', async () => {
    await reset({ zoom: 100, view: { top: 14, left: 0, r: 16, c: 5 } }); assert.equal((await state()).fr, 14);
    await slider(330); assert.equal((await visible('확대 F17', { r: 16, c: 5 })).fr, 0);
    await slider(100);
    const s = await visible('축소 F17', { r: 16, c: 5 }); assert.equal(s.fr, 14); assert.equal(s.frozenH, 308);
  });
  await test('창 높이를 늘리고 줄이면 고정 표시를 복원하며 원본 설정을 유지', async () => {
    await reset({ zoom: 330, view: { top: 14, left: 0, r: 16, c: 5 } }); assert.equal((await state()).fr, 0);
    await page.setViewportSize({ width: 1800, height: 1800 }); await frame();
    await visible('창 확대 직후');
    await page.keyboard.press('ArrowDown'); await frame();
    assert.equal((await visible('큰 창')).fr, 14);
    await page.setViewportSize({ width: 1800, height: 1000 }); await frame();
    await visible('창 축소 직후');
    await page.keyboard.press('ArrowDown'); await frame();
    assert.equal((await visible('작은 창')).fr, 0);
  });
  await test('숨긴 행은 고정 영역의 높이 계산에서 제외하여 불필요한 해제를 방지', async () => {
    const hiddenRows = {}; for (let r = 0; r < 12; r++) hiddenRows[r] = true;
    await reset({ zoom: 330, hiddenRows, view: { top: 14, left: 0, r: 14, c: 5 } });
    const s = await visible('숨긴 고정 행', { r: 14, c: 5 });
    assert.equal(s.fr, 14); assert.equal(s.frozenH, 44); assert.deepEqual(s.frozenRows, [12, 13]);
  });
  await test('0이 아닌 행·열 원점을 가진 고정 창의 좌표와 셀 클릭이 일치', async () => {
    await reset({ ...originOptions, freeze: { rows: 14, cols: 2, top: 294, left: 3 }, hiddenCols: { 3: true },
      view: { top: 329, left: 8, r: 329, c: 8 } });
    const s = await visible('양축 원점', { r: 329, c: 8 });
    assert.equal(s.fc, 5); assert.equal(s.frozenLeft, 3); assert.equal(s.frozenW, 64);
    await page.mouse.click((s.dom.left + s.dom.right) / 2, (s.dom.top + s.dom.bottom) / 2);
    await visible('양축 원점 클릭', { r: 329, c: 8 });
  });
  await test('큰 고정 영역의 F2 편집·실행 취소·다시 실행은 값과 선택 표시를 보존', async () => {
    await reset({ zoom: 330, view: { top: 16, left: 0, r: 16, c: 5 } });
    const before = await page.evaluate(() => tabula.wb().getRaw(tabula.si, 16, 5));
    await page.keyboard.press('F2'); await page.locator('#cellEditor').fill('수정 합성 값'); await page.keyboard.press('Enter');
    assert.equal(await page.evaluate(() => tabula.wb().getRaw(tabula.si, 16, 5)), '수정 합성 값');
    await visible('편집 후');
    await page.evaluate(() => tabula.run('undo'));
    assert.equal(await page.evaluate(() => tabula.wb().getRaw(tabula.si, 16, 5)), before); await visible('실행 취소');
    await page.evaluate(() => tabula.run('redo'));
    assert.equal(await page.evaluate(() => tabula.wb().getRaw(tabula.si, 16, 5)), '수정 합성 값'); await visible('다시 실행');
  });
  await test('Shift 방향키로 화면 아래까지 확장한 선택 영역의 끝이 보임', async () => {
    await reset({ zoom: 330, view: { top: 0, left: 0, r: 0, c: 5 } });
    // The anchor stays at F1 in an extended selection; verify the moving edge directly.
    for (let i = 0; i < 16; i++) await page.keyboard.press('Shift+ArrowDown');
    const m = await page.evaluate(() => {
      const t = tabula, g = t.gv(), s = t.sel, r = g.clientRect({ r1: s.r2, c1: s.c2, r2: s.r2, c2: s.c2 }), v = g.viewEl.getBoundingClientRect();
      return { selection: { ...s }, rect: r, bottom: v.bottom, hit: g.hitTest((r.left + r.right) / 2, (r.top + r.bottom) / 2) };
    });
    assert.deepEqual(m.selection, { r1: 0, c1: 5, r2: 16, c2: 5 });
    assert.ok(m.rect.bottom <= m.bottom + 2); assert.equal(m.hit.r, 16); assert.equal(m.hit.c, 5);
    assert.ok((await state()).unchanged);
  });
  const summary = { total: results.length, passed: results.filter(r => r.ok).length, failed: results.filter(r => !r.ok).length, results, pageErrors, blockedWrites: writes };
  console.log(JSON.stringify(summary));
  if (output) await writeFile(join(output, 'freeze-navigation.json'), JSON.stringify({ ...summary, metrics }, null, 2));
  assert.equal(summary.failed, 0); assert.deepEqual(pageErrors, []); assert.deepEqual(writes, []);
} finally { await browser.close(); }
