// 차트 요소 직접 선택·이동·삭제·서식: 합성 문서와 격리 브라우저만 사용한다.
// WIXEL_URL(소스/번들/공개), WIXEL_CHART_DIRECT_FILTER, WIXEL_CHART_DIRECT_SCREENSHOT 선택 가능.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const browser = await chromium.launch(), results = [], filter = process.env.WIXEL_CHART_DIRECT_FILTER;
async function test(name, fn) {
  if (filter && !name.includes(filter)) return;
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 }, acceptDownloads: false });
  const p = await context.newPage(), errors = [], writes = []; p.setDefaultTimeout(10000);
  p.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => {
    const r = route.request(), target = new URL(r.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); }
    return target.origin === new URL(url).origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
    await fixture(p); await fn(p); await cellsUnchanged(p);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors, blockedWrites: writes }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
const current = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets[0].charts[0]));
const state = p => p.evaluate(() => structuredClone(window.tabula.gv().host.state().chartPart));
const depth = p => p.evaluate(() => window.tabula.wb().undoStack.length);
const command = (p, cmd) => p.evaluate(cmd => window.tabula.run(cmd), cmd);
async function fixture(p, patch = {}, zoom = 100) {
  await p.evaluate(({ patch, zoom }) => {
    const t = window.tabula, w = t.wb(), cells = {};
    [['분류', '매출', '비용', '목표'], ['가', '10', '30', '15'], ['나', '20', '20', '25'], ['다', '30', '10', '35']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
    cells['6,0'] = { raw: '=1+1' };
    w.restore({ sheets: [{ name: '차트 합성', cells }, { name: '다른 합성', cells: { '0,0': { raw: '다른 시트 보존' } } }] }); t.switchSheet(0);
    t.gv().setZoom(zoom); t.gv().layout(); t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 3 }); t.run('chartColumn');
    w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], x: 200, y: 60, w: 600, h: 360, title: '합성 제목', legend: 'b', ...patch }]));
    t.gv().renderObjectsAll(); w.undoStack = []; w.redoStack = [];
    window.__chartCellSource = w;
    window.__chartCellsBefore = JSON.stringify(w.serialize().sheets.map(s => s.cells));
  }, { patch, zoom });
}
const cellsUnchanged = p => p.evaluate(() => {
  if (JSON.stringify(window.__chartCellSource.serialize().sheets.map(s => s.cells)) !== window.__chartCellsBefore) throw Error('차트 편집이 원본 셀/수식을 변경함');
});
const element = (p, kind) => p.locator(`.obj.chart [data-el="${kind}"]${kind === 'legend' ? ' text' : ''}`).first();
const point = (p, s = 0, n = 1) => p.locator(`.obj.chart svg [data-s="${s}"][data-p="${n}"]:not(text)`).first();
const center = async locator => { const b = await locator.boundingBox(); assert.ok(b, '요소 화면 위치'); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
async function drag(p, target, dx, dy, cancel = false) {
  const a = typeof target.x === 'number' ? target : await center(target);
  await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(a.x + dx, a.y + dy, { steps: 8 });
  if (cancel) await p.keyboard.press('Escape'); await p.mouse.up();
}
const undo = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+z'); };
const redo = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+y'); };
const pane = p => p.getByRole('dialog').filter({ has: p.locator('.cfp, .chart-selection-pane') });
const picker = p => pane(p).getByRole('combobox', { name: '서식을 지정할 차트 요소' });
const color = (loc, value) => loc.evaluate((el, value) => { el.value = value; el.dispatchEvent(new Event('change', { bubbles: true })); }, value);
async function piePosition(p, index = 1) {
  return point(p, 0, index).evaluate(el => {
    const d = el.dataset, angle = Number(d.pieAngle), radius = Number(d.pieR), squash = Number(d.pieSquash || 1);
    const v = new DOMPoint(Number(d.pieCx) + Math.cos(angle) * radius * .4, Number(d.pieCy) + Math.sin(angle) * radius * .4 * squash).matrixTransform(el.ownerSVGElement.getScreenCTM());
    return { x: v.x, y: v.y, dx: Math.cos(angle) * 45, dy: Math.sin(angle) * 45 * squash };
  });
}
try {
  for (const kind of ['title', 'legend']) {
    await test(`${kind}: 요소만 이동·차트 위치 보존·한 번 Undo/Redo`, async p => {
      const before = await current(p); await drag(p, element(p, kind), 45, 28); const after = await current(p);
      assert.deepEqual([after.x, after.y, after.w, after.h], [before.x, before.y, before.w, before.h]);
      assert.ok(Number.isFinite(after[kind + 'Layout']?.x) && Number.isFinite(after[kind + 'Layout']?.y), '수동 위치 저장');
      assert.equal(await depth(p), 1); await undo(p); assert.deepEqual(await current(p), before); await redo(p); assert.deepEqual(await current(p), after);
    });
    await test(`${kind}: 이동 중 Esc는 위치·Undo 기록을 바꾸지 않음`, async p => {
      const before = await current(p); await drag(p, element(p, kind), 55, 25, true); assert.deepEqual(await current(p), before); assert.equal(await depth(p), 0);
    });
    await test(`${kind}: 선택 Delete는 해당 요소만 제거·차트와 원본 셀 유지`, async p => {
      const before = await current(p); await element(p, kind).click(); assert.equal((await state(p)).kind, kind); await p.keyboard.press('Delete');
      const after = await current(p); assert.ok(after, '차트 전체를 삭제하면 안 됨'); assert.equal(kind === 'title' ? after.title : after.legend, kind === 'title' ? '' : 'none');
      await undo(p); assert.deepEqual(await current(p), before);
    });
    await test(`${kind}: Ctrl+1은 선택한 차트 요소의 서식 패널을 엶`, async p => {
      await element(p, kind).click(); await p.keyboard.press('Control+1'); await pane(p).waitFor();
      assert.equal(await p.getByRole('dialog', { name: '셀 서식', exact: true }).count(), 0);
      assert.match(await picker(p).locator('option:checked').innerText(), kind === 'title' ? /제목/ : /범례/);
    });
  }
  await test('계열→데이터 요소 클릭과 열린 패널의 선택 대상 동기화·개별 색·Undo', async p => {
    await point(p, 1, 1).click(); assert.equal((await state(p)).kind, 'series'); await p.keyboard.press('Control+1'); await pane(p).waitFor(); assert.equal(await picker(p).inputValue(), 'series:1');
    await color(pane(p).getByLabel('계열 채우기 색', { exact: true }), '#123456');
    assert.equal((await current(p)).seriesFmt[1].color, '#123456');
    await point(p, 1, 1).click(); assert.deepEqual((({ kind, s, p: n }) => [kind, s, n])(await state(p)), ['point', 1, 1]);
    assert.equal(await picker(p).inputValue(), 'point:1:1');
    const before = await current(p); await color(pane(p).getByLabel('선택한 요소 색', { exact: true }), '#ff3300');
    const after = await current(p); assert.equal(after.seriesFmt[1].pointColors[1], '#ff3300'); assert.equal(after.seriesFmt[1].color, '#123456');
    await pane(p).getByRole('button', { name: '닫기', exact: true }).click(); await undo(p); assert.deepEqual(await current(p), before);
  });
  await test('데이터 요소 더블클릭은 해당 요소 서식을 열고 차트 편집 창을 열지 않음', async p => {
    await point(p, 0, 1).dblclick(); await pane(p).waitFor(); assert.equal(await picker(p).inputValue(), 'point:0:1');
    assert.equal(await p.getByRole('dialog', { name: '차트 편집', exact: true }).count(), 0);
  });
  for (const selected of ['series', 'point']) await test(`${selected}: Delete는 Excel처럼 해당 계열만 제거·Undo`, async p => {
    const before = await current(p); await point(p, 1, 1).click(); if (selected === 'point') await point(p, 1, 1).click();
    assert.equal((await state(p)).kind, selected); await p.keyboard.press('Delete'); const after = await current(p);
    assert.ok(after, '차트 유지'); assert.deepEqual(after.hiddenSeries, [1]);
    assert.equal(await p.locator('.obj.chart svg [data-s="1"]').count(), 0); assert.ok(await p.locator('.obj.chart svg [data-s="0"]').count());
    await undo(p); assert.deepEqual(await current(p), before);
  });
  await test('마지막 계열 삭제는 빈 차트 틀을 유지·차트 영역 Delete와 구분', async p => {
    await fixture(p, { range: { r1: 0, c1: 0, r2: 3, c2: 1 } }); await point(p, 0, 1).click(); await p.keyboard.press('Delete'); assert.ok(await current(p));
    assert.equal(await p.locator('.obj.chart svg [data-s]').count(), 0); await undo(p);
    const box = p.locator('.obj.chart').first(); await box.click({ position: { x: 4, y: 4 } }); assert.equal(await state(p), null); await p.keyboard.press('Delete'); assert.equal(await current(p), undefined);
    await undo(p); assert.ok(await current(p));
  });
  for (const threeD of [false, true]) for (const zoom of [50, 100, 150]) await test(`원형 ${threeD ? '3D' : '2D'} ${zoom}%: 개별 조각 분리·다른 조각/차트 위치 보존·Undo`, async p => {
    await fixture(p, { type: 'pie', threeD, range: { r1: 0, c1: 0, r2: 3, c2: 1 }, labels: false }, zoom);
    let a = await piePosition(p); await p.mouse.click(a.x, a.y); await p.mouse.click(a.x, a.y); assert.equal((await state(p)).kind, 'point');
    const before = await current(p); a = await piePosition(p); await drag(p, a, a.dx, a.dy); const after = await current(p);
    assert.ok(after.seriesFmt?.[0]?.pointExplosion?.[1] > 0, '선택 조각의 분리량'); assert.equal(Object.keys(after.seriesFmt[0].pointExplosion).length, 1);
    assert.deepEqual([after.x, after.y, after.w, after.h], [before.x, before.y, before.w, before.h]); assert.equal(await depth(p), 1);
    assert.doesNotMatch(await p.locator('.obj.chart svg').first().evaluate(el => el.outerHTML), /NaN|Infinity/);
    if (threeD && zoom === 100 && process.env.WIXEL_CHART_DIRECT_SCREENSHOT) await p.screenshot({ path: process.env.WIXEL_CHART_DIRECT_SCREENSHOT });
    await undo(p); assert.deepEqual(await current(p), before); await redo(p); assert.deepEqual(await current(p), after);
  });
  await test('원형 조각 분리 중 Esc는 분리값·Undo 기록을 취소', async p => {
    await fixture(p, { type: 'pie', range: { r1: 0, c1: 0, r2: 3, c2: 1 }, labels: false }); const a = await piePosition(p); await p.mouse.click(a.x, a.y); await p.mouse.click(a.x, a.y);
    const before = await current(p); await drag(p, a, a.dx, a.dy, true); assert.deepEqual(await current(p), before); assert.equal(await depth(p), 0);
  });
  await test('숨긴 계열 뒤의 데이터 요소 선택은 원래 계열 번호를 유지', async p => {
    await fixture(p, { hiddenSeries: [0], seriesFmt: [{ color: '#123456' }, { color: '#654321' }] }); await point(p, 1, 2).click(); await point(p, 1, 2).click(); await p.keyboard.press('Control+1');
    await pane(p).waitFor(); assert.equal(await picker(p).inputValue(), 'point:1:2'); await color(pane(p).getByLabel('선택한 요소 색', { exact: true }), '#ff3300');
    const after = await current(p); assert.equal(after.seriesFmt[0].color, '#123456'); assert.equal(after.seriesFmt[1].pointColors[2], '#ff3300'); assert.deepEqual(after.hiddenSeries, [0]);
  });
  await test('시트 보호는 선택 요소 이동·삭제·열린 서식 패널 변경을 막음', async p => {
    await element(p, 'title').click(); await p.keyboard.press('Control+1'); await pane(p).waitFor();
    await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { objects: false } })); }); const before = await current(p);
    const titleInput = pane(p).getByLabel('제목 텍스트', { exact: true }); if (await titleInput.isEnabled()) { await titleInput.fill('보호 위반'); await titleInput.dispatchEvent('change'); }
    await pane(p).getByRole('button', { name: '닫기', exact: true }).click(); await drag(p, element(p, 'title'), 30, 20); await p.keyboard.press('Delete'); assert.deepEqual(await current(p), before);
  });
  await test('다른 시트로 전환하면 열린 차트 서식이 이전·새 시트를 바꾸지 않음', async p => {
    await point(p, 0, 1).click(); await p.keyboard.press('Control+1'); await pane(p).waitFor(); const before = await current(p);
    await p.evaluate(() => window.tabula.switchSheet(1)); if (await pane(p).count()) {
      const input = pane(p).getByLabel('계열 채우기 색', { exact: true }); if (await input.count() && await input.isEnabled()) await color(input, '#ff00ff');
    }
    assert.deepEqual(await current(p), before); assert.equal(await p.evaluate(() => window.tabula.wb().sheets[1].charts.length), 0);
  });
  for (const kind of ['title', 'legend']) await test(`${kind}: 방향키1px·Ctrl방향키10px 이동·Undo`, async p => {
    await fixture(p, { titleLayout: { x: .2, y: .2 }, legendLayout: { x: .5, y: .5 } });
    const before = await current(p); await element(p, kind).click(); await p.keyboard.press('ArrowRight'); const right = await current(p);
    await p.keyboard.press('Control+ArrowUp'); const down = await current(p);
    assert.ok(right[kind + 'Layout']); assert.ok(Math.abs((down[kind + 'Layout'].y - right[kind + 'Layout'].y) * before.h + 10) < .01);
    assert.deepEqual([down.x, down.y], [before.x, before.y]); await undo(p); assert.deepEqual(await current(p), right); await undo(p); assert.deepEqual(await current(p), before);
  });
  await test('꺾은선 개별 표식 색은 계열 선 색과 다른 점을 보존', async p => {
    await fixture(p, { type: 'line', range: { r1: 0, c1: 0, r2: 3, c2: 1 }, marker: 'circle', seriesFmt: [{ color: '#123456', marker: 'circle' }] });
    await point(p, 0, 1).click(); await point(p, 0, 1).click(); await p.keyboard.press('Control+1'); await pane(p).waitFor();
    await color(pane(p).getByLabel('선택한 표식 색', { exact: true }), '#ff3300'); const after = await current(p);
    assert.equal(after.seriesFmt[0].color, '#123456'); assert.deepEqual(after.seriesFmt[0].pointColors, { 1: '#ff3300' });
    assert.equal(await point(p, 0, 1).getAttribute('fill'), '#ff3300'); assert.notEqual(await point(p, 0, 0).getAttribute('fill'), '#ff3300');
  });
  await test('공개 읽기 전용 합성 문서는 제목 이동·Delete·서식 변경으로 바뀌지 않음', async p => {
    const workbook = await p.evaluate(() => window.tabula.wb().serialize());
    const target = url.split('#')[0] + '#view=' + gzipSync(JSON.stringify({ docName: '읽기 전용 차트 합성', workbook })).toString('base64url');
    await p.goto('about:blank'); await p.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.locator('.view-bar').waitFor();
    await p.evaluate(() => { const w = window.tabula.wb(); window.__chartCellSource = w; window.__chartCellsBefore = JSON.stringify(w.serialize().sheets.map(s => s.cells)); });
    const before = await current(p); await drag(p, element(p, 'title'), 35, 20); await element(p, 'title').click(); await p.keyboard.press('Delete'); assert.deepEqual(await current(p), before);
    await p.keyboard.press('Control+1');
    if (await pane(p).count()) { const input = pane(p).getByLabel('제목 텍스트', { exact: true }); if (await input.count() && await input.isEnabled()) await color(input, '읽기 전용 위반'); }
    assert.deepEqual(await current(p), before);
  });
  await test('다른 문서 열기 뒤 이전 패널 이벤트는 새 문서나 이전 차트를 바꾸지 않음', async p => {
    await point(p, 0, 1).click(); await p.keyboard.press('Control+1'); await pane(p).waitFor();
    const oldInput = await pane(p).getByLabel('계열 채우기 색', { exact: true }).elementHandle();
    await p.evaluate(() => { window.__oldChart = window.tabula.wb().sheets[0].charts[0]; window.__oldChartBefore = JSON.stringify(window.__oldChart); });
    await p.setInputFiles('#fileInput', { name: '새 합성 문서.wixel', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify({ sheets: [{ name: '새 합성', cells: { '0,0': { raw: '새 문서 보존' } } }] })) });
    await p.waitForFunction(() => window.tabula.wb().sheets[0].name === '새 합성');
    await oldInput.evaluate(el => { el.value = '#ff00ff'; el.dispatchEvent(new Event('change', { bubbles: true })); });
    const same = await p.evaluate(() => {
      const w = window.tabula.wb(); if (w.getRaw(0, 0, 0) !== '새 문서 보존' || w.sheets[0].charts.length) throw Error('새 문서 변경됨');
      window.__chartCellSource = w; window.__chartCellsBefore = JSON.stringify(w.serialize().sheets.map(s => s.cells));
      return JSON.stringify(window.__oldChart) === window.__oldChartBefore;
    }); assert.equal(same, true, '이전 차트도 유지');
  });
  const failed = results.filter(r => !r.ok).length; console.log(JSON.stringify({ url, tests: results.length, failed, results }, null, 2)); if (failed) process.exitCode = 1;
} finally { await browser.close(); }
