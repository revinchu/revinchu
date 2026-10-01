// 실제 사용자 파일·보관함과 분리한 합성 차트의 종류/축 편집 회귀.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [], results = [];
page.setDefaultTimeout(10000);
page.on('pageerror', (e) => errors.push(e.message));
const ev = (fn, arg) => page.evaluate(fn, arg);
const current = () => ev(() => window.tabula.wb().sheets[0].charts[0]);
const fixture = async (patch = {}) => {
  for (let i = 0; i < 6; i++) await page.keyboard.press('Escape');
  await ev((patch) => {
    const t = window.tabula, w = t.wb(); t.switchSheet(0);
    const cells = {};
    [['월', '매출', '방문', '전환율'], ['1월', '120', '1000', '.02'], ['2월', '180', '1500', '.03'], ['3월', '240', '1100', '.04']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
    w.restore({ sheets: [{ name: '차트 검사', cells }] }); t.gv().layout(); t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 3 }); t.run('chartColumn');
    w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], type: 'combo', title: '월별 매출과 전환율', seriesFmt: [{ type: 'column', axis: 0, color: '#336699' }, { type: 'area', axis: 0, labels: true }, { type: 'line', axis: 1, marker: 'diamond' }], ...patch }]));
    w.undoStack = []; w.redoStack = []; t.gv().renderObjectsAll();
  }, patch);
};
const open = async () => { await ev(() => window.tabula.run('chartChangeType')); return page.getByRole('dialog', { name: '차트 종류 변경', exact: true }); };
const type = (i) => page.locator(`[data-combo-type="${i}"]`);
const axis = (i) => page.locator(`[data-combo-axis="${i}"]`);
const test = async (name, fn) => {
  const before = errors.length;
  try { await fn(); assert.deepEqual(errors.slice(before), []); results.push({ name, ok: true }); console.log(`OK ${name}`); }
  catch (e) { results.push({ name, ok: false }); console.error(`NG ${name}: ${e.stack}`); }
};
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/', { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => !!window.tabula?.wb(), null, { timeout: 60000 });
  await test('기존 계열별 종류·축 복원 / 변경 즉시 미리보기 / 취소 보존', async () => {
    await fixture(); const before = await current(); const dialog = await open();
    assert.deepEqual(await page.locator('[data-combo-type]').evaluateAll((els) => els.map((s) => s.value)), ['column', 'area', 'line']);
    assert.deepEqual(await page.locator('[data-combo-axis]').evaluateAll((els) => els.map((s) => s.value)), ['0', '0', '1']);
    const preview = await page.locator('.cg-prev svg').evaluate((s) => s.outerHTML);
    await type(0).selectOption('line'); await axis(2).selectOption('0');
    assert.notEqual(await page.locator('.cg-prev svg').evaluate((s) => s.outerHTML), preview);
    assert.deepEqual(await current(), before); assert.equal(await ev(() => window.tabula.wb().undoStack.length), 0);
    await axis(2).selectOption('1');
    await page.screenshot({ path: process.env.WIXEL_CHART_SCREENSHOT || 'D:/Codex/Temp/wixel3-combo-chart.png' });
    await dialog.getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await current(), before);
  });
  await test('여러 계열 종류·축을 한 번에 확정 / 1회 실행 취소·다시 실행', async () => {
    await fixture(); const before = await current(); const dialog = await open();
    await type(0).selectOption('line'); await axis(0).selectOption('1'); await type(1).selectOption('column'); await axis(2).selectOption('0');
    await dialog.getByRole('button', { name: '확인', exact: true }).click();
    const after = await current(); assert.deepEqual(after.seriesFmt.map((f) => [f.type, f.axis]), [['line', 1], ['column', 0], ['line', 0]]);
    assert.equal(after.seriesFmt[0].color, '#336699'); assert.equal(after.seriesFmt[1].labels, true); assert.equal(after.seriesFmt[2].marker, 'diamond');
    assert.equal(await ev(() => window.tabula.wb().undoStack.length), 1);
    await ev(() => window.tabula.run('undo')); assert.deepEqual(await current(), before);
    await ev(() => window.tabula.run('redo')); assert.deepEqual(await current(), after);
  });
  await test('범위 없는 가져온 계열 / 숨긴 계열도 원래 번호로 편집', async () => {
    await fixture({ range: null, series: [{ name: { text: '목표' }, cache: [100, 110, 120] }, { name: { text: '실적' }, cache: [80, 130, 140] }, { name: { text: '비율' }, cache: [.8, 1.18, 1.16] }], hiddenSeries: [0] });
    const dialog = await open(); assert.equal(await page.locator('[data-combo-type]').count(), 3); assert.match(await dialog.innerText(), /목표.*숨김/);
    await type(1).selectOption('line'); await axis(1).selectOption('1'); await dialog.getByRole('button', { name: '확인', exact: true }).click();
    const c = await current(); assert.equal(c.range, null); assert.equal(c.series.length, 3); assert.deepEqual(c.hiddenSeries, [0]); assert.deepEqual([c.seriesFmt[1].type, c.seriesFmt[1].axis], ['line', 1]); assert.equal(c.seriesFmt[0].type, 'column');
  });
  await test('동일 종류의 기본축·보조축 가져온 차트도 콤보 편집 표시', async () => {
    await fixture({ type: 'column', seriesFmt: [{ type: 'column', axis: 0 }, { type: 'column', axis: 1 }, { type: 'column', axis: 1 }] });
    const dialog = await open(); assert.equal(await page.locator('[data-combo-type]').count(), 3); assert.equal(await axis(1).inputValue(), '1');
    await dialog.getByRole('button', { name: '취소', exact: true }).click();
  });
  await test('기본축·보조축 콤보 견본 선택과 사용자 설정 보존', async () => {
    await fixture(); const dialog = await open();
    const choices = dialog.locator('.cg-sub'); assert.equal(await choices.count(), 2);
    await dialog.getByRole('button', { name: '묶은 세로 막대형 - 꺾은선형, 기본 축', exact: true }).click(); assert.deepEqual(await page.locator('[data-combo-axis]').evaluateAll((els) => els.map((s) => s.value)), ['0', '0', '0']);
    await dialog.getByRole('button', { name: '묶은 세로 막대형 - 꺾은선형, 보조 축', exact: true }).click(); assert.deepEqual(await page.locator('[data-combo-axis]').evaluateAll((els) => els.map((s) => s.value)), ['0', '0', '1']);
    await type(1).selectOption('area'); await axis(1).selectOption('1'); await dialog.getByRole('button', { name: '확인', exact: true }).click();
    const reopened = await open(); assert.equal(await type(1).inputValue(), 'area'); assert.equal(await axis(1).inputValue(), '1'); await reopened.getByRole('button', { name: '취소', exact: true }).click();
  });
  await test('콤보에서 일반 꺾은선형으로 변경 시 계열 종류·축 재설정', async () => {
    await fixture(); const dialog = await open(); await dialog.getByRole('button', { name: '꺾은선형', exact: true }).first().click();
    await dialog.getByRole('button', { name: '확인', exact: true }).click(); const c = await current();
    assert.equal(c.type, 'line'); assert.equal(c.threeD, false); assert.equal(c.seriesFmt.some((f) => f.type !== undefined || f.axis !== undefined), false); assert.equal(c.seriesFmt[0].color, '#336699');
  });
  await test('차트 삽입에서도 계열별 콤보 설정 확정', async () => {
    await fixture(); await ev(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'charts', [])); t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 3 }); t.run('insertChartAll'); });
    const dialog = page.getByRole('dialog', { name: '차트 삽입', exact: true }); await dialog.getByRole('button', { name: '모든 차트', exact: true }).click(); await dialog.getByRole('button', { name: '콤보', exact: true }).click();
    await type(0).selectOption('area'); await axis(2).selectOption('0'); await dialog.getByRole('button', { name: '확인', exact: true }).click();
    const c = await current(); assert.equal(c.type, 'combo'); assert.equal(c.seriesFmt[0].type, 'area'); assert.equal(c.seriesFmt[2].axis, 0);
  });
  await test('기존 차트 편집: 숨긴 계열 번호와 보조축→기본축 0 보존', async () => {
    await fixture({ seriesFmt: [], hiddenSeries: [0] });
    await page.locator('.obj.chart').first().dblclick({ position: { x: 50, y: 30 } });
    const dialog = page.getByRole('dialog', { name: '차트 편집', exact: true });
    assert.equal(await dialog.getByRole('combobox', { name: / 축$/ }).count(), 3);
    const last = dialog.getByRole('combobox', { name: '전환율 축', exact: true }); assert.equal(await last.inputValue(), '1');
    await last.selectOption('0'); await dialog.getByRole('button', { name: '확인', exact: true }).click();
    const c = await current(); assert.equal(c.seriesFmt[2].axis, 0); assert.deepEqual(c.hiddenSeries, [0]);
  });
  const bad = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ total: results.length, ok: results.length - bad, bad, pageErrors: errors.length }));
  if (bad || errors.length) { if (errors.length) console.error(errors); process.exitCode = 1; }
} finally { await browser.close(); }
