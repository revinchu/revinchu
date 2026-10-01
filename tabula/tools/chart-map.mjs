// 네트워크를 끊은 브라우저에서 합성 국가/지역 데이터의 지도 표시를 확인합니다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }), page = await context.newPage();
const errors = [], writes = []; let checks = 0;
const eq = (a, b, message) => { assert.deepEqual(a, b, message); checks++; };
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => { if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort('blockedbyclient'); } return route.continue(); });
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(process.env.WIXEL_URL || 'http://127.0.0.1:5180/', { timeout: 60000 }); await page.waitForFunction(() => !!window.tabula?.wb());
  await context.setOffline(true);
  await page.evaluate(() => {
    const t = window.tabula, w = t.wb(), cells = {};
    [['국가/지역', '증감'], ['한국', '100'], ['KR', '25'], ['미국', '-20'], ['영국', '0'], ['싱가포르', '35'], ['프랑스', '70'], ['노르웨이', '40'], ['서울', '12'], ['일본', '잘못된 값']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
    w.restore({ sheets: [{ name: '지도 검증', cells }] }); t.gv().renderAll(); t.selectRange({ r1: 0, c1: 0, r2: 9, c2: 1 }); t.run('chartColumn');
    w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], type: 'map', title: '국가별 증감 · 오프라인 지도', x: 90, y: 40, w: 900, h: 530, legend: 'b' }])); t.selectCell(0, 0); t.gv().renderObjectsAll();
  });
  const chart = page.locator('.obj.chart').first();
  eq(await chart.locator('path[data-country]').count(), 177); eq(await chart.locator('circle[data-map-point]').count(), 37);
  for (const [id, value] of [['KOR', '125'], ['USA', '-20'], ['GBR', '0'], ['SGP', '35'], ['FRA', '70'], ['NOR', '40']]) eq(await chart.locator(`[data-country="${id}"][data-value]`).first().getAttribute('data-value'), value, id);
  eq(await chart.locator('[data-country="GBR"]').first().getAttribute('fill') !== await chart.locator('[data-country="JPN"]').first().getAttribute('fill'), true, '0 is not missing');
  eq((await chart.locator('[data-map="unmatched"]').textContent()).includes('인식 못한 지역 1개'), true); eq((await chart.locator('[data-map="unmatched"]').textContent()).includes('숫자 아닌 값·합계 오류 1개'), true);
  eq(await chart.locator('[data-el="legend"]').count(), 1); eq((await chart.locator('[data-country="KOR"] title').textContent()).includes('2개 행 합계'), true);
  eq(await chart.locator('[data-country="SGP"] title').textContent(), '싱가포르 (SG): 35 · 작은 국가/지역 위치 표시');
  await chart.screenshot({ path: process.env.WIXEL_MAP_SCREENSHOT || 'D:/Codex/Temp/wixel3-country-map.png' });
  await page.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'charts', w.sheets[0].charts.map(c => ({ ...c, mapLowColor: '#ff0000', mapMidColor: '#ffff00', mapHighColor: '#0000ff' })))); t.gv().renderObjectsAll(); });
  eq(await chart.evaluate(el => ['USA', 'GBR', 'KOR'].map(id => el.querySelector('[data-country="' + id + '"]').getAttribute('fill'))), ['#ff0000', '#ffff00', '#0000ff'], 'map low/mid/high colors render at negative/zero/positive values');
  // 기존 지도 옵션을 숨김으로 바꿨을 때 색상 범례만 사라지고 지도는 남습니다.
  await page.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'charts', w.sheets[0].charts.map(c => ({ ...c, legend: 'none' })))); t.gv().renderObjectsAll(); });
  eq(await chart.locator('[data-el="legend"]').count(), 0); eq(await chart.locator('path[data-country]').count(), 177);
  eq(errors, [], 'pageErrors'); eq(writes, [], 'server writes');
  console.log(JSON.stringify({ ok: true, checks, offline: true, polygonRegions: 177, markers: 37, pageErrors: errors, blockedWrites: writes }));
} finally { await browser.close(); }
