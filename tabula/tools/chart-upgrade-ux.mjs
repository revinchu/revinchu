// 합성 차트로만 갤러리/서식 패널을 검사한다. 소스 및 최종 번들 공용.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const browser = await chromium.launch();
const results = [];
async function test(name, run) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage(); page.setDefaultTimeout(10000);
  const errors = [], writes = [];
  page.on('pageerror', (error) => errors.push(error.stack || error.message));
  await context.route('**/*', (route) => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
    const target = new URL(route.request().url());
    if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) return route.abort();
    return route.continue();
  });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
    await fixture(page); await run(page);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '서버 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors }); console.error(`NG ${name}: ${error.stack}`); }
  finally { await context.close(); }
}
const current = (p) => p.evaluate(() => window.tabula.wb().sheets[0].charts[0]);
const command = (p, cmd) => p.evaluate((cmd) => window.tabula.run(cmd), cmd);
const fixture = (p, patch = {}) => p.evaluate((patch) => {
  const t = window.tabula, w = t.wb(), cells = {};
  [['월', '매출', '방문', '전환율'], ['1월', '120', '1000', '.02'], ['2월', '180', '1500', '.03'], ['3월', '240', '1100', '.04']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
  w.restore({ sheets: [{ name: '차트 UX 검사', cells }] });
  t.gv().layout(); t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 3 }); t.run('chartColumn');
  w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], title: '월별 성과', ...patch }]));
  w.undoStack = []; w.redoStack = []; t.gv().renderObjectsAll();
}, patch);
const gallery = async (p) => { await command(p, 'chartChangeType'); return p.getByRole('dialog', { name: '차트 종류 변경', exact: true }); };
const format = async (p) => {
  await command(p, 'chartFormat');
  // 서식 명령은 현재 선택한 요소의 창을 연다. 기존 전체 범주 검사는 명시적 전환 뒤 실행한다.
  const selected = p.getByRole('dialog', { name: '차트 영역 서식', exact: true });
  await selected.waitFor();
  assert.equal(await selected.getByRole('combobox', { name: '서식을 지정할 차트 요소' }).inputValue(), 'chart');
  await selected.getByRole('button', { name: '차트 전체 옵션…', exact: true }).click();
  const all = p.getByRole('dialog', { name: '차트 서식', exact: true });
  await all.waitFor(); return all;
};
const category = (d, name) => d.locator('.cg-cat').filter({ hasText: new RegExp(`^${name}$`) });
const optionFor = (picker, label) => picker.locator('option').filter({ hasText: new RegExp(`^${label}$`) }).getAttribute('value');
const pickerSection = async (d, name) => { const picker = d.getByRole('combobox', { name: '서식을 지정할 차트 요소' }); await picker.selectOption(await optionFor(picker, name)); };
const changeNumber = async (locator, value) => { await locator.fill(String(value)); await locator.press('Tab'); };
try {
  await test('갤러리 한글 캡션·검색·검색 결과 없음에서 정상 안내 복귀', async (p) => {
    const before = await current(p), d = await gallery(p), search = d.getByRole('searchbox', { name: '차트 종류 검색' });
    await search.fill('표식');
    assert.deepEqual(await d.locator('.cg-cat:visible').allTextContents(), ['꺾은선형', '분산형', '방사형']);
    const captions = await d.locator('.cg-sub:visible .cg-sub-name').allTextContents();
    assert.ok(captions.length >= 3 && captions.every((s) => s.includes('표식')));
    await d.locator('.cg-sub:visible').last().click(); assert.match(await d.locator('.cg-label').innerText(), /100%/);
    await search.fill('없는차트999'); assert.equal(await d.locator('.cg-cat:visible').count(), 0); assert.match(await d.locator('.cg-data-guide').innerText(), /일치하는 차트가 없습니다/);
    const confirm = d.getByRole('button', { name: '확인', exact: true });
    if (await confirm.isEnabled()) await confirm.click();
    assert.equal(await d.isVisible(), true, '검색 결과가 없는데 이전 유형을 확정함'); assert.deepEqual(await current(p), before);
    await search.fill(''); assert.ok(await d.locator('.cg-cat:visible').count() > 10); assert.doesNotMatch(await d.locator('.cg-data-guide').innerText(), /일치하는 차트가 없습니다/);
    await p.screenshot({ path: process.env.WIXEL_CHART_UPGRADE_SCREENSHOT || 'D:/Codex/Temp/wixel3-chart-upgrade.png' });
    await d.getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await current(p), before);
  });
  await test('차트 분류·하위 종류 방향키와 Home/End 실제 선택', async (p) => {
    const d = await gallery(p); await category(d, '세로 막대형').focus(); await p.keyboard.press('ArrowDown');
    assert.equal(await d.locator('.cg-cat.on').innerText(), '꺾은선형');
    await p.keyboard.press('Home'); assert.equal(await d.locator('.cg-cat.on').innerText(), '세로 막대형');
    await d.locator('.cg-sub').first().focus(); await p.keyboard.press('ArrowRight');
    assert.equal(await d.locator('.cg-sub[aria-pressed=true]').getAttribute('title'), '누적 세로 막대형');
    await p.keyboard.press('End'); assert.equal(await d.locator('.cg-sub[aria-pressed=true]').getAttribute('title'), '깊이 축 피라미드 세로 막대형');
    await d.getByRole('button', { name: '확인', exact: true }).click(); const c = await current(p);
    assert.equal(c.threeD, true); assert.equal(c.grouping, 'standard'); assert.equal(c.barShape, 'pyramid');
  });
  await test('고른 하위 유형의 재열기 유지: 표식 누적·3D·거품·주식', async (p) => {
    for (const [patch, name] of [
      [{ type: 'line', grouping: 'percentStacked', marker: 'circle' }, '표식이 있는 100% 기준 누적 꺾은선형'],
      [{ type: 'column', threeD: true, grouping: 'stacked' }, '3차원 누적 세로 막대형'],
      [{ type: 'bubble', threeD: true }, '3차원 효과가 있는 거품형'],
      [{ type: 'stock', volume: false, ohlc: false }, '고가-저가-종가'],
    ]) {
      await fixture(p, patch); const before = await current(p), d = await gallery(p);
      assert.equal(await d.locator('.cg-sub[aria-pressed=true]').getAttribute('title'), name);
      await d.getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await current(p), before);
    }
  });
  await test('거래량·시가 주식형: 3계열은 거부하고 5계열은 저장', async (p) => {
    const before = await current(p); let d = await gallery(p);
    await category(d, '주식형').click(); await d.getByRole('button', { name: '거래량-시가-고가-저가-종가', exact: true }).click();
    assert.match(await d.locator('.cg-data-guide').innerText(), /5개.*현재 3개/);
    await d.getByRole('button', { name: '확인', exact: true }).click(); assert.equal(await d.isVisible(), true); assert.deepEqual(await current(p), before);
    await d.getByRole('button', { name: '취소', exact: true }).click();
    await fixture(p, { range: null, series: ['거래량', '시가', '고가', '저가', '종가'].map((name, i) => ({ name: { text: name }, cache: [[500, 600, 700], [10, 20, 30], [15, 25, 35], [5, 15, 25], [12, 22, 32]][i] })) });
    d = await gallery(p); await category(d, '주식형').click(); await d.getByRole('button', { name: '거래량-시가-고가-저가-종가', exact: true }).click();
    assert.equal(await d.locator('.cg-data-guide').evaluate((el) => el.classList.contains('error')), false);
    await d.getByRole('button', { name: '확인', exact: true }).click(); const c = await current(p);
    assert.equal(c.type, 'stock'); assert.equal(c.volume, true); assert.equal(c.ohlc, true);
  });
  await test('표면형 4종의 미리 보기·실제 모델·재열기 유지', async (p) => {
    const svg = new Set();
    for (const [name, style, threeD] of [['3차원 표면형', 'surface', true], ['3차원 표면형(골격형)', 'wireframe', true], ['등고선형', 'contour', false], ['등고선형(골격형)', 'wireframeContour', false]]) {
      await fixture(p); let d = await gallery(p); await category(d, '표면형').click(); assert.equal(await d.locator('.cg-sub').count(), 4);
      await d.getByRole('button', { name, exact: true }).click(); svg.add(await d.locator('.cg-prev svg').evaluate((el) => el.outerHTML));
      await d.getByRole('button', { name: '확인', exact: true }).click(); const c = await current(p); assert.equal(c.surfaceStyle, style); assert.equal(c.threeD, threeD);
      d = await gallery(p); assert.equal(await d.locator('.cg-sub[aria-pressed=true]').getAttribute('title'), name); await d.getByRole('button', { name: '취소', exact: true }).click();
    }
    assert.equal(svg.size, 4, '표면형 하위 종류 미리 보기가 서로 같음');
  });
  await test('요소 선택: 숨긴 계열 0 이후 실제 계열 번호·색·기본 축 0 유지', async (p) => {
    await fixture(p, { type: 'combo', hiddenSeries: [0], seriesFmt: [{ color: '#123456', axis: 0 }, { color: '#654321', axis: 1 }, { axis: 1 }] });
    const before = await current(p), d = await format(p), picker = d.getByRole('combobox', { name: '서식을 지정할 차트 요소' });
    assert.match(await picker.locator('option[value="series:0"]').innerText(), /매출.*숨김/);
    await picker.selectOption('series:1'); const section = d.locator('[data-format-series="1"]');
    assert.equal(await d.locator('[data-format-series="0"]').isVisible(), false); assert.equal(await section.isVisible(), true);
    await section.getByRole('combobox', { name: '축', exact: true }).selectOption('0');
    await section.getByLabel('색', { exact: true }).evaluate((el) => { el.value = '#ff3300'; el.dispatchEvent(new Event('change', { bubbles: true })); });
    const after = await current(p); assert.equal(after.seriesFmt[1].axis, 0); assert.equal(after.seriesFmt[1].color, '#ff3300'); assert.deepEqual(after.seriesFmt[0], before.seriesFmt[0]);
    await d.getByRole('button', { name: '닫기', exact: true }).click(); await command(p, 'undo'); await command(p, 'undo'); assert.deepEqual(await current(p), before);
  });
  await test('계열 선택 후 서식 탭 방향키로 전체 범주 복귀', async (p) => {
    const d = await format(p), picker = d.getByRole('combobox', { name: '서식을 지정할 차트 요소' });
    await picker.selectOption('series:1'); await d.getByRole('tab', { name: '차트 옵션', exact: true }).focus(); await p.keyboard.press('ArrowRight');
    assert.equal(await picker.inputValue(), 'all'); assert.equal(await d.getByRole('spinbutton', { name: '너비(px)', exact: true }).isVisible(), true);
    await p.keyboard.press('Home'); assert.equal(await d.locator('[data-format-series="0"]').isVisible(), true); assert.equal(await d.locator('[data-format-series="1"]').isVisible(), true);
  });
  await test('폭포 차트 합계 지정·데이터 요소 전환·렌더 변경·한 번 실행 취소', async (p) => {
    await fixture(p, { type: 'waterfall', totals: [] }); const before = await current(p), d = await format(p);
    await pickerSection(d, '계열 옵션'); await d.getByRole('combobox', { name: '데이터 요소', exact: true }).selectOption('1');
    const oldSvg = await p.locator('.obj.chart svg').first().evaluate((el) => el.outerHTML);
    await d.getByRole('checkbox', { name: '합계로 설정', exact: true }).check(); assert.deepEqual((await current(p)).totals, [1]);
    assert.notEqual(await p.locator('.obj.chart svg').first().evaluate((el) => el.outerHTML), oldSvg);
    await d.getByRole('combobox', { name: '데이터 요소', exact: true }).selectOption('0'); assert.equal(await d.getByRole('checkbox', { name: '합계로 설정', exact: true }).isChecked(), false);
    await d.getByRole('combobox', { name: '데이터 요소', exact: true }).selectOption('1'); assert.equal(await d.getByRole('checkbox', { name: '합계로 설정', exact: true }).isChecked(), true);
    await d.getByRole('button', { name: '닫기', exact: true }).click(); await command(p, 'undo'); assert.deepEqual(await current(p), before);
  });
  await test('보조 축의 표시·역순·주 단위가 기본 축과 계열을 보존', async (p) => {
    await fixture(p, { type: 'combo', seriesFmt: [{ axis: 0 }, { axis: 0 }, { axis: 1 }], axes: { y: { min: 0, max: 2000, major: 500 }, y2: { min: 0, max: .1 } } });
    const before = await current(p), d = await format(p); await pickerSection(d, '보조 값 축');
    await d.getByRole('checkbox', { name: '표시', exact: true }).uncheck();
    await d.getByRole('checkbox', { name: '값을 거꾸로', exact: true }).check(); await changeNumber(d.getByRole('spinbutton', { name: '주 단위', exact: true }), .02);
    const c = await current(p); assert.deepEqual(c.axes.y2, { min: 0, max: .1, hide: true, reverse: true, major: .02 }); assert.deepEqual(c.axes.y, before.axes.y); assert.deepEqual(c.seriesFmt, before.seriesFmt);
    await d.getByRole('button', { name: '닫기', exact: true }).click(); for (let i = 0; i < 3; i++) await command(p, 'undo'); assert.deepEqual(await current(p), before);
  });
  const bad = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ tests: results.length, failed: bad, results }, null, 2)); if (bad) process.exitCode = 1;
} finally { await browser.close(); }
