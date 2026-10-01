// 소스/번들 공용. 격리된 합성 통합 문서만 쓰며 서버 쓰기 요청은 차단합니다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [];
const command = (p, c) => p.evaluate((c) => window.tabula.run(c), c);
const current = (p) => p.evaluate(() => window.tabula.wb().sheets[0].charts[0]);
const chartSvg = (p) => p.locator('.obj.chart svg').first();
const pick = async (d, text) => {
  const picker = d.getByRole('combobox', { name: '서식을 지정할 차트 요소' });
  await picker.selectOption(await picker.getByRole('option', { name: text, exact: true }).getAttribute('value'));
};
const fixture = (p, chart, rows = [['분포', '금액'], ['가', 1], ['나', 2], ['다', 3], ['라', 4], ['마', 5], ['바', 6], ['사', 30]]) => p.evaluate(({ chart, rows }) => {
  const t = window.tabula, w = t.wb(), cells = {};
  rows.forEach((row, r) => row.forEach((v, c) => { cells[`${r},${c}`] = { raw: String(v) }; }));
  w.restore({ sheets: [{ name: '통계 차트 검사', cells }] });
  t.gv().layout(); t.selectRange({ r1: 0, c1: 0, r2: rows.length - 1, c2: rows[0].length - 1 }); t.run('chartColumn');
  w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], w: 640, h: 430, labels: false, ...chart }]));
  w.undoStack = []; w.redoStack = []; t.gv().renderObjectsAll();
}, { chart, rows });
async function test(name, run) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }), page = await context.newPage(), errors = [], writes = [];
  page.setDefaultTimeout(15000); page.on('pageerror', (e) => errors.push(e.message));
  await context.route('**/*', (route) => { if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); } return route.continue(); });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/', { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
    await run(page); assert.deepEqual(errors, []); assert.deepEqual(writes, []);
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.message, errors, writes }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
async function format(p) { await command(p, 'chartFormat'); return p.getByRole('dialog', { name: '차트 서식', exact: true }); }
const number = async (d, name, v) => { const el = d.getByRole('spinbutton', { name, exact: true }); await el.fill(String(v)); await el.press('Tab'); };
try {
  await test('상자수염 옵션: 내부점·이상값·평균·사분위 즉시 반영과 실행 취소', async (p) => {
    await fixture(p, { type: 'boxWhisker' }); const before = await current(p), d = await format(p); await pick(d, '계열 옵션');
    assert.equal(await chartSvg(p).locator('[data-box=outlier]').count(), 1);
    await d.getByRole('checkbox', { name: '이상값 표시', exact: true }).uncheck(); assert.equal(await chartSvg(p).locator('[data-box=outlier]').count(), 0);
    await d.getByRole('checkbox', { name: '내부 데이터 요소 표시', exact: true }).check(); assert.equal(await chartSvg(p).locator('[data-box=inner]').count(), 6);
    await d.getByRole('checkbox', { name: '평균 표식 표시', exact: true }).uncheck(); assert.equal(await chartSvg(p).locator('[data-box=mean]').count(), 0);
    await d.getByRole('combobox', { name: '사분위수 계산', exact: true }).selectOption('exclusive');
    assert.equal(await chartSvg(p).locator('[data-box=quartiles]').getAttribute('data-q1'), '2'); assert.equal(await chartSvg(p).locator('[data-box=quartiles]').getAttribute('data-q3'), '6');
    assert.equal(await d.getByRole('combobox', { name: '추세선', exact: true }).count(), 0, '작동하지 않는 추세선 컨트롤을 노출하지 않음');
    await d.getByRole('button', { name: '닫기', exact: true }).click(); for (let i = 0; i < 4; i++) await command(p, 'undo'); assert.deepEqual(await current(p), before);
  });
  await test('폭포 축: 범위·주 단위·역순·제목·표시 형식·숨김', async (p) => {
    await fixture(p, { type: 'waterfall' }); const d = await format(p); await pick(d, '세로(값) 축');
    await number(d, '최소값', 0); await number(d, '최대값', 60); await number(d, '주 단위', 20);
    await d.getByRole('textbox', { name: '표시 형식', exact: true }).fill('0.00'); await d.getByRole('textbox', { name: '표시 형식', exact: true }).press('Tab');
    assert.deepEqual(await chartSvg(p).locator('[data-axis=y]').allTextContents(), ['0.00', '20.00', '40.00', '60.00']);
    await d.getByRole('checkbox', { name: '값을 거꾸로', exact: true }).check();
    const ticks = await chartSvg(p).locator('[data-axis=y]').evaluateAll((es) => es.map((e) => +e.getAttribute('y'))); assert.ok(ticks[0] < ticks.at(-1));
    const title = d.getByRole('textbox', { name: '제목', exact: true }); await title.fill('검증 값 축'); await title.press('Tab'); assert.equal(await chartSvg(p).locator('[data-axis-title=y]').textContent(), '검증 값 축');
    await d.getByRole('checkbox', { name: '표시', exact: true }).uncheck(); assert.equal(await chartSvg(p).locator('[data-axis=y]').count(), 0);
    assert.equal(await chartSvg(p).locator('[data-axis-title=y]').count(), 0);
  });
  for (const volume of [false, true]) await test(`${volume ? '거래량 주식' : '파레토'} 보조 축은 실제 별도 숫자 축을 제어`, async (p) => {
    const rows = volume ? [['일', '거래량', '고가', '저가', '종가'], ['월', 1000000, 12, 5, 8], ['화', 2000000, 18, 6, 10]] : [['항목', '빈도'], ['a', 6], ['b', 3], ['c', 1]];
    await fixture(p, { type: volume ? 'stock' : 'pareto', volume, ohlc: false }, rows); const d = await format(p); await pick(d, '보조 값 축');
    await number(d, '주 단위', volume ? 2 : .2); await d.getByRole('checkbox', { name: '값을 거꾸로', exact: true }).check();
    const ticks = await chartSvg(p).locator('[data-axis=y2]').evaluateAll((es) => es.map((e) => +e.getAttribute('y'))); assert.ok(ticks.length > 2 && ticks[0] < ticks.at(-1));
    await d.getByRole('checkbox', { name: '표시', exact: true }).uncheck(); assert.equal(await chartSvg(p).locator('[data-axis=y2]').count(), 0); assert.ok(await chartSvg(p).locator('[data-axis=y]').count() > 0);
  });
  await test('표면형: 깊이·직각 축·원근감이 실제 면 기하를 변경', async (p) => {
    await fixture(p, { type: 'surface', surfaceStyle: 'surface', threeD: true }, [['행', 'a', 'b', 'c'], ['r1', 0, 2, 6], ['r2', 4, 9, 5], ['r3', 7, 3, 1]]);
    const d = await format(p); await d.getByRole('tab', { name: '3차원 회전', exact: true }).click();
    const geometry = () => chartSvg(p).locator('[data-band]').evaluateAll((es) => es.map((e) => e.getAttribute('points')).join('|'));
    const a = await geometry(); await number(d, '깊이(%)', 300); const b = await geometry(); assert.notEqual(a, b);
    await d.getByRole('checkbox', { name: '직각 축', exact: true }).uncheck(); const c = await geometry(); assert.notEqual(b, c);
    await number(d, '원근감', 180); assert.notEqual(c, await geometry());
    if (process.env.WIXEL_CHART_STAT_SCREENSHOT) await p.screenshot({ path: process.env.WIXEL_CHART_STAT_SCREENSHOT });
  });
  await test('통계 데이터 표: 화면 반영·Excel 표시 제한 안내·지원하지 않는 차트는 숨김', async (p) => {
    await fixture(p, { type: 'boxWhisker' }); let d = await format(p); await pick(d, '범례 · 레이블');
    await d.getByRole('checkbox', { name: '데이터 표', exact: true }).check(); assert.equal(await chartSvg(p).locator('[data-el=dataTable]').count(), 1);
    await d.getByRole('button', { name: '닫기', exact: true }).click(); d = await format(p); await pick(d, '범례 · 레이블'); assert.match(await d.innerText(), /Excel의 데이터 표로 표시되지는 않습니다/);
    await d.getByRole('button', { name: '닫기', exact: true }).click(); await fixture(p, { type: 'sunburst' }); d = await format(p); await pick(d, '범례 · 레이블');
    assert.equal(await d.getByRole('checkbox', { name: '데이터 표', exact: true }).count(), 0);
  });
  const failed = results.filter((r) => !r.ok).length; console.log(JSON.stringify({ tests: results.length, failed, results }, null, 2)); if (failed) process.exitCode = 1;
} finally { await browser.close(); }
