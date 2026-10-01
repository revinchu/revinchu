// 범례 회귀: 실제 피벗 계산과 차트 명령/UI, 독립 컨텍스트의 합성 문서만 사용한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [];
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const names = ['브랜드 광고', '신규 광고'];

async function test(name, path, threeD, options = {}) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const p = await context.newPage(), pageErrors = [], blockedWrites = []; p.setDefaultTimeout(10000);
  p.on('pageerror', error => pageErrors.push(error.message));
  await context.route('**/*', route => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { blockedWrites.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  const observation = { name, path, threeD };
  try {
    await p.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
    observation.source = await p.evaluate(({ path, names }) => {
      const t = window.tabula, w = t.wb(), cells = {};
      const extra = path === 'numericValues' || path === 'numericCost';
      const rows = [['캠페인', '광고 세트', '노출'], ['검색 광고', names[0], '120'], ['검색 광고', names[1], '300'], ['검색 광고', names[0], '80']];
      if (extra) ['비용', '40', '150', '10'].forEach((value, i) => rows[i].push(value));
      rows.forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
      w.restore({ sheets: [{ name: '광고 원본', cells }] });
      w.transact(() => {
        w.addSheet('피벗 보고서');
        w.setSheetProp(1, 'pivot', { name: '광고 세트별 노출', source: '광고 원본', range: { r1: 0, c1: 0, r2: 3, c2: extra ? 3 : 2 },
          rows: path === 'multiLabel' ? ['캠페인', '광고 세트'] : ['광고 세트'], cols: [], values: [{ field: '노출', agg: 'sum' }],
          layout: 'tabular', subtotals: false, grandRows: false, grandCols: false, top: 0, left: 0 });
        if (extra) w.setSheetProp(1, 'pivot', { ...w.sheets[1].pivot, values: [{ field: '노출', agg: 'sum' }, { field: '비용', agg: 'sum' }] });
      });
      t.switchSheet(1); t.run('pivotRefresh');
      const area = w.sheets[1].pivot.area;
      const valueCol = path === 'numericValues' ? area.c1 + 1 : area.c2;
      const selected = path.startsWith('numeric') ? { r1: area.r1 + 1, c1: valueCol, r2: area.r2, c2: valueCol } : { ...area };
      t.selectRange(selected); w.undoStack = []; w.redoStack = []; t.gv().layout(); t.gv().renderAll();
      return { area, selected, rows: Array.from({ length: area.r2 - area.r1 + 1 }, (_, i) => Array.from({ length: area.c2 - area.c1 + 1 }, (_, j) => w.getValue(1, area.r1 + i, area.c1 + j))) };
    }, { path, names });
    const before = await p.evaluate(() => window.tabula.wb().serialize());
    if (path === 'pivotChart') await p.evaluate(() => window.tabula.run('insertPivotChart'));
    if (threeD || path === 'pivotChart') {
      const changing = path === 'pivotChart';
      await p.evaluate(cmd => window.tabula.run(cmd), changing ? 'chartChangeType' : 'insertChartAll');
      const d = p.getByRole('dialog', { name: changing ? '차트 종류 변경' : '차트 삽입', exact: true }); await d.waitFor();
      if (!changing) await d.getByRole('button', { name: '모든 차트', exact: true }).click();
      await d.locator('.cg-cat').filter({ hasText: /^원형$/ }).click();
      await d.locator(`.cg-sub[aria-label="${threeD ? '3차원 원형' : '원형'}"]`).click();
      observation.previewLegend = await d.locator('.cg-prev [data-el="legend"] text').allTextContents();
      await d.getByRole('button', { name: '확인', exact: true }).click(); await d.waitFor({ state: 'hidden' });
    } else await p.evaluate(cmd => window.tabula.run(cmd), path === 'columnLegend' ? 'chartColumn' : 'chartPie');
    if (options.legacy) {
      observation.legacyModelUnchanged = await p.evaluate(selected => {
        const t = window.tabula, w = t.wb(), current = w.sheets[1].charts.at(-1);
        const legacy = { ...current, range: selected }; delete legacy.series; delete legacy.pivot;
        w.setSheetProp(1, 'charts', [legacy]); const serialized = JSON.stringify(w.sheets[1].charts[0]);
        t.gv().renderObjectsAll(); return serialized === JSON.stringify(w.sheets[1].charts[0]);
      }, observation.source.selected);
    }
    observation.chart = await p.evaluate(() => window.tabula.wb().sheets[1].charts.at(-1));
    const chart = p.locator(`.obj.chart[data-id="${observation.chart.id}"]`); await chart.waitFor();
    observation.legend = await chart.locator('[data-el="legend"] text').allTextContents();
    console.log('OBS ' + JSON.stringify(observation));
    if (process.env.CHART_LEGEND_SCREENSHOT && path === 'numeric' && threeD) await p.screenshot({ path: process.env.CHART_LEGEND_SCREENSHOT, fullPage: true });
    assert.equal(observation.chart.type, path === 'columnLegend' ? 'column' : 'pie'); assert.equal(!!observation.chart.threeD, threeD);
    assert.deepEqual(observation.legend, path === 'columnLegend' ? ['합계 : 노출'] : names, '범례가 원본 항목/계열 이름과 일치해야 함');
    if (observation.previewLegend) assert.deepEqual(observation.previewLegend, names, '갤러리 미리보기와 삽입 결과 범례 일치');
    if (path === 'pivotChart') assert.equal(observation.chart.pivot?.name, '광고 세트별 노출');
    if (options.legacy) assert.equal(observation.legacyModelUnchanged, true, '기존 범위 차트 렌더링이 저장 모델을 변경하면 안 됨');
    if (path === 'numericValues' || path === 'numericCost') {
      assert.equal(observation.chart.series?.length, 1, '선택하지 않은 지표를 계열에 포함하지 않음');
      const val = observation.chart.series[0].val;
      assert.deepEqual({ r1: val.r1, c1: val.c1, r2: val.r2, c2: val.c2 }, observation.source.selected, '선택한 숫자 범위 그대로 참조');
      const actual = await p.evaluate(ref => { const w = window.tabula.wb(), si = ref.sheet ? w.sheetIndexByName(ref.sheet) : 1; return [ref.r1, ref.r2].map(r => w.getValue(si, r, ref.c1)); }, val);
      assert.deepEqual(actual, path === 'numericCost' ? [50, 150] : [200, 300]);
      assert.match(observation.chart.title, path === 'numericCost' ? /비용/ : /노출/);
    }
    if (path === 'columnLegend') {
      await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.setSheetProp(1, 'charts', [{ ...w.sheets[1].charts[0], legend: 'none' }]); t.gv().renderObjectsAll(); });
      assert.equal(await chart.locator('[data-el="legend"]').count(), 0, '명시 범례 없음은 유지');
    } else if (!options.legacy) {
      await p.evaluate(() => window.tabula.run('undo')); if (path === 'pivotChart') await p.evaluate(() => window.tabula.run('undo'));
      assert.deepEqual(await p.evaluate(() => window.tabula.wb().serialize()), before, '차트 삽입 실행 취소는 원래 문서를 복원');
    }
    assert.deepEqual(pageErrors, [], '페이지 오류'); assert.deepEqual(blockedWrites, [], '원격 쓰기');
    results.push({ ...observation, ok: true }); console.log(`OK ${name}`);
  } catch (error) { results.push({ ...observation, ok: false, error: error.message, pageErrors, blockedWrites }); console.error(`NG ${name}: ${error.stack}`); }
  finally { await context.close(); }
}
try {
  for (const [path, label] of [['numeric', '피벗 숫자 열만 선택'], ['multiLabel', '피벗 행 레이블 두 열 포함'], ['pivotChart', '피벗 차트 삽입 경로']]) {
    for (const threeD of [false, true]) await test(`${label} → ${threeD ? '3D 원형' : '원형'} 범례`, path, threeD);
  }
  await test('기존 숫자 범위만 저장한 3D 원형은 모델 변경 없이 범례 복구', 'numeric', true, { legacy: true });
  await test('단일 세로 막대 계열은 기본 범례 표시·명시 none 숨김 유지', 'columnLegend', false);
  await test('두 지표 피벗에서 노출만 선택하면 비용을 계열에 포함하지 않음', 'numericValues', false);
  await test('두 지표 피벗에서 비용만 선택하면 노출을 계열에 포함하지 않음', 'numericCost', true);
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(item => item.ok).length, bad: results.filter(item => !item.ok).length, results }, null, 2));
if (results.some(item => !item.ok)) process.exitCode = 1;
