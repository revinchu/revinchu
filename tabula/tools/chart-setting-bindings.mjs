// 합성 차트만 사용하는 실제 설정→SVG 회귀. 외부 및 API 쓰기를 차단합니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_CHART_BINDINGS_OUT || 'D:/Codex/Temp/wixel-pivot-audit/chart-settings';
await mkdir(out, { recursive: true });
const browser = await chromium.launch(), results = []; let checks = 0;
const eq = (a, b, label) => { checks++; assert.deepEqual(a, b, label); };
const ok = (a, label) => { checks++; assert.ok(a, label); };
const run = (p, cmd) => p.evaluate(cmd => window.tabula.run(cmd), cmd);
const chart = p => p.locator('.obj.chart').first();
const pane = p => p.getByRole('dialog').filter({ has: p.locator('.chart-selection-pane') });
const model = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets[0].charts[0]));
async function fixture(p, patch) {
  await p.evaluate(patch => {
    const t = window.tabula, w = t.wb(), cells = {};
    [['가로값', '금액'], [1, 10], [2, 40], [3, 20]].forEach((row, r) => row.forEach((v, c) => cells[r + ',' + c] = { raw: String(v) }));
    w.restore({ sheets: [{ name: '차트 설정 합성', cells }] }); t.switchSheet(0); t.gv().layout();
    t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 1 }); t.run('chartColumn');
    w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], x: 120, y: 60, w: 560, h: 360, legend: 'none', ...patch }]));
    t.gv().renderObjectsAll(); w.undoStack = []; w.redoStack = [];
    window.__bindingCells = JSON.stringify(w.serialize().sheets[0].cells);
  }, patch);
}
async function test(name, fn) {
  const ctx = await browser.newContext({ viewport: { width: 1450, height: 1000 } }), p = await ctx.newPage(), errors = [], writes = [];
  p.setDefaultTimeout(12000); p.on('pageerror', e => errors.push(e.message));
  await ctx.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await ctx.route('**/*', route => { const r = route.request(), target = new URL(r.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(target.pathname); return route.abort(); } return target.origin === new URL(url).origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort(); });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded' }); await p.waitForFunction(() => !!window.tabula?.wb());
    await fn(p);
    eq(await p.evaluate(() => JSON.stringify(window.tabula.wb().serialize().sheets[0].cells)), await p.evaluate(() => window.__bindingCells), '원본 셀 불변');
    eq(errors, [], '페이지 오류'); eq(writes, [], '원격 쓰기'); results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.stack }); await p.screenshot({ path: out + '/failure-' + results.length + '.png' }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await ctx.close(); }
}
try {
  await test('분산형 계열의 곡선·파선·표식 없음은 실제 그림에 반영되고 Undo', async p => {
    await fixture(p, { type: 'scatter', scatterStyle: 'lineMarker' }); const before = await model(p);
    await chart(p).locator('svg circle[data-p="2"]').first().click(); await p.keyboard.press('Control+1'); await pane(p).waitFor();
    const path = () => chart(p).locator('path[data-s="0"]:not([data-p])').first();
    ok(!(await path().getAttribute('d')).includes('C'), '처음 직선');
    await pane(p).getByRole('checkbox', { name: '부드러운 선', exact: true }).check(); ok((await path().getAttribute('d')).includes('C'), '곡선 반영');
    await pane(p).getByRole('combobox', { name: '선 종류', exact: true }).selectOption('dashDot'); ok(await path().getAttribute('stroke-dasharray'), '파선 반영');
    await pane(p).getByRole('combobox', { name: '표식 모양', exact: true }).selectOption('none'); eq(await chart(p).locator('svg [data-p]').count(), 0, '표식 없음');
    await p.keyboard.press('Escape'); for (let i = 0; i < 3; i++) await run(p, 'undo'); eq(await model(p), before, 'Undo');
  });
  await test('깔때기 개별 요소 색과 레이블 표시·이름을 실제 편집하고 원본 보존', async p => {
    await fixture(p, { type: 'funnel', labels: undefined });
    const point = () => chart(p).locator('rect[data-s="0"][data-p="1"]').first();
    await point().click({ position: { x: 30, y: 20 } }); await point().click({ position: { x: 30, y: 20 } }); await p.keyboard.press('Control+1'); await pane(p).waitFor();
    eq(await pane(p).getByRole('combobox', { name: '서식을 지정할 차트 요소' }).inputValue(), 'point:0:1');
    await pane(p).getByLabel('선택한 요소 색', { exact: true }).evaluate(el => { el.value = '#cc2200'; el.dispatchEvent(new Event('change', { bubbles: true })); });
    eq(await point().getAttribute('fill'), '#cc2200');
    await pane(p).getByRole('combobox', { name: '서식을 지정할 차트 요소' }).selectOption('label:0');
    eq(await pane(p).getByRole('checkbox', { name: '값 표시', exact: true }).isChecked(), true, '깔때기 기본 표시와 일치');
    eq(await pane(p).getByRole('checkbox', { name: '레이블 굵게', exact: true }).isChecked(), true);
    eq(await pane(p).getByRole('combobox', { name: '레이블 위치', exact: true }).count(), 0, '미지원 위치 메뉴 숨김');
    await pane(p).getByRole('checkbox', { name: '값 표시', exact: true }).uncheck(); eq(await chart(p).locator('[data-el="label"]').count(), 0);
    await pane(p).getByRole('checkbox', { name: '항목 이름', exact: true }).check(); eq(await chart(p).locator('[data-el="label"]').count(), 3);
    eq(await chart(p).locator('[data-el="label"][data-p="1"]').textContent(), '2');
    await p.screenshot({ path: out + '/funnel-settings.png' });
    await pane(p).getByRole('button', { name: '계열 데이터 레이블 삭제', exact: true }).click(); eq(await chart(p).locator('[data-el="label"]').count(), 0);
  });
} finally { await browser.close(); }
const result = { url, cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
await writeFile(out + '/results.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify(result));
process.exitCode = results.some(r => !r.ok) ? 1 : 0;
