import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { pivotSourceData, resolvePivot, computePivot, pivotChartData } from '../src/pivot.js';
import { inferPivotCategorySeries } from '../src/chart-source.js';
import { chartModelData } from '../src/chart.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';

const basicRows = [['광고', '클릭수', '비용'], ['네이버', 879, 12000], ['구글', 5379, 80000]];
function fixture(rows = basicRows, patch = {}) {
  const wb = new Workbook();
  wb.transact(() => {
    wb.addSheet('피벗 결과');
    rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v))));
  });
  const def = { name: '광고 요약', source: 'Sheet1', range: { r1: 0, c1: 0, r2: rows.length - 1, c2: rows[0].length - 1 }, rows: ['광고'], cols: [], values: [{ field: '클릭수', agg: 'sum' }, { field: '비용', agg: 'sum' }], layout: 'tabular', top: 0, left: 0, ...patch };
  const render = () => {
    const source = pivotSourceData(wb, def), res = resolvePivot(source, def);
    const { grid, meta } = computePivot(res, res.def);
    let width = 0;
    for (const row of grid) width = Math.max(width, row.length);
    def.area = { r1: def.top, c1: def.left, r2: def.top + grid.length - 1, c2: def.left + width - 1 };
    wb.transact(() => {
      grid.forEach((row, r) => row.forEach((cell, c) => wb.setCellData(1, def.top + r, def.left + c, cell)));
      wb.setSheetProp(1, 'pivot', def);
    });
    return { grid, meta, data: pivotChartData(source, def) };
  };
  return { wb, def, render, ...render() };
}
const chart = (range, patch = {}) => ({ id: 'inferred', type: 'pie', range, x: 30, y: 30, w: 600, h: 360, ...patch });
const ref = (c, r1 = 1, r2 = 2) => ({ sheet: '피벗 결과', r1, c1: c, r2, c2: c });
const shape = (m) => ({ categories: m.categories, series: m.series.map((s) => ({ name: s.name, values: s.values })) });

test('피벗 숫자 879/5379만 선택해도 범주는 광고명 참조이며 값 범위를 확장하지 않음', () => {
  const { wb, def } = fixture();
  const ch = chart({ r1: 1, c1: 1, r2: 2, c2: 1 });
  const before = structuredClone({ ch, def });
  const series = inferPivotCategorySeries(wb, 1, ch);
  assert.deepEqual(series, [{ name: { ref: ref(1, 0, 0) }, cat: ref(0), val: ref(1) }]);
  const model = chartModelData(wb, 1, ch);
  assert.deepEqual(model.categories, ['구글', '네이버']);
  assert.deepEqual(model.series[0].values, [5379, 879]);
  assert.deepEqual({ ch, def }, before, '원래 차트와 피벗 정의를 변경하지 않음');
  for (const type of ['column', 'bar', 'line', 'area', 'combo', 'pie', 'doughnut', 'radar', 'pieOfPie', 'barOfPie', 'waterfall']) {
    assert.deepEqual(inferPivotCategorySeries(wb, 1, { ...ch, type }), series, type);
  }
  assert.deepEqual(inferPivotCategorySeries(wb, 1, { ...ch, threeD: true }), series);
});

test('두 지표 또는 뒤쪽 비용만 선택해도 클릭 수를 범주로 오인하지 않음', () => {
  const { wb } = fixture();
  const cost = inferPivotCategorySeries(wb, 1, chart({ r1: 1, c1: 2, r2: 2, c2: 2 }));
  assert.deepEqual(cost[0].cat, ref(0));
  assert.deepEqual(cost[0].val, ref(2));
  const both = inferPivotCategorySeries(wb, 1, chart({ r1: 1, c1: 1, r2: 2, c2: 2 }, { type: 'column' }));
  assert.equal(both.length, 2);
  assert.deepEqual(both.map((s) => s.cat), [ref(0), ref(0)]);
  assert.deepEqual(both.map((s) => s.val), [ref(1), ref(2)]);
});

test('행 부분선택과 머리글·총합계 포함 선택은 선택 안의 잎 행만 연결', () => {
  const { wb, def } = fixture();
  const subset = inferPivotCategorySeries(wb, 1, chart({ r1: 2, c1: 1, r2: 2, c2: 1 }));
  assert.deepEqual(subset[0].cat, ref(0, 2, 2));
  assert.deepEqual(subset[0].val, ref(1, 2, 2));
  const all = inferPivotCategorySeries(wb, 1, chart({ r1: 0, c1: 1, r2: def.area.r2, c2: 1 }));
  assert.deepEqual(all[0].val, ref(1));
  assert.equal(inferPivotCategorySeries(wb, 1, chart({ r1: 0, c1: 1, r2: 0, c2: 1 })), null);
  assert.equal(inferPivotCategorySeries(wb, 1, chart({ r1: def.area.r2, c1: 1, r2: def.area.r2, c2: 1 })), null);
});

test('다층 표형·압축·개요의 마지막 레이블 열 사용, 부분합 사이 불연속 선택 거부', () => {
  const rows = [['상위', '광고', '클릭수', '비용'], ['봄', '네이버', 879, 12000], ['봄', '구글', 5379, 80000], ['여름', '카카오', 24, 900]];
  for (const layout of ['tabular', 'compact', 'outline']) {
    const f = fixture(rows, { rows: ['상위', '광고'], layout });
    const [r1, r2, last] = f.data.rows, c = f.meta.labelCols + 1;
    const ch = chart({ r1, c1: c, r2, c2: c });
    const series = inferPivotCategorySeries(f.wb, 1, ch);
    assert.deepEqual(series[0].cat, ref(f.meta.labelCols - 1, r1, r2), layout);
    assert.deepEqual(chartModelData(f.wb, 1, ch).categories, ['구글', '네이버'], layout);
    assert.equal(inferPivotCategorySeries(f.wb, 1, chart({ r1, c1: c, r2: last, c2: c })), null, `${layout}: 중간 요약 행을 참조하지 않음`);
  }
});

test('열 그룹의 부분합·총합계와 인접한 값 열도 실제 col 메타데이터로 판별', () => {
  const rows = [['광고', '분기', '월', '클릭수', '비용'], ['구글', '상반기', '1월', 5379, 80000], ['네이버', '상반기', '2월', 879, 12000], ['구글', '하반기', '7월', 10, 100]];
  const f = fixture(rows, { cols: ['분기', '월'] });
  const validCols = new Set(f.data.series.map((s) => s.col));
  assert.ok(validCols.size > 2);
  assert.ok(f.meta.width > validCols.size + f.meta.labelCols, '부분합 또는 총합계 열이 있는 fixture');
  for (let c = f.meta.labelCols; c < f.meta.width; c++) {
    const ch = chart({ r1: f.data.rows[0], c1: c, r2: f.data.rows.at(-1), c2: c });
    const got = inferPivotCategorySeries(f.wb, 1, ch);
    if (validCols.has(c)) {
      assert.equal(got[0].cat.c1, 0);
      assert.equal(got[0].val.c1, c);
      assert.equal(chartModelData(f.wb, 1, ch).series[0].name, f.data.series.find((s) => s.col === c).name);
    } else assert.equal(got, null, `합계 열 ${c} 제외`);
  }
});

test('일반 셀·명시 계열·피벗 차트·스냅샷·행 방향·수치 XY 종류는 그대로 유지', () => {
  const { wb, def } = fixture();
  const ch = chart({ r1: 1, c1: 1, r2: 2, c2: 1 });
  for (const patch of [{ series: [] }, { series: [{ val: ref(1) }] }, { pivot: { name: def.name } }, { snapshotData: {} }, { byRows: true }, { type: 'scatter' }, { type: 'bubble' }, { type: 'surface' }, { type: 'stock' }]) {
    assert.equal(inferPivotCategorySeries(wb, 1, { ...ch, ...patch }), null);
  }
  assert.equal(inferPivotCategorySeries(wb, 0, ch), null, '주변 일반 셀 레이블은 추측하지 않음');
  assert.equal(inferPivotCategorySeries(wb, 1, { ...ch, range: { ...ch.range, c1: 0 } }), null, '이미 레이블을 포함한 범위는 유지');
  assert.equal(inferPivotCategorySeries(wb, 1, { ...ch, range: { ...ch.range, r2: 50 } }), null);
  assert.equal(inferPivotCategorySeries(wb, 1, { ...ch, sheet: '없는 시트' }), null);
  assert.deepEqual(inferPivotCategorySeries(wb, 0, { ...ch, sheet: '피벗 결과' }), inferPivotCategorySeries(wb, 1, ch));
});

test('값 필드가 행에 놓인 복합 피벗은 지표 이름을 항목으로 추론하지 않음', () => {
  const f = fixture(basicRows, { valuesOnRows: true });
  assert.equal(f.data.series.length, 1);
  assert.ok(f.data.rows.length > 2, '각 광고 아래 클릭 수/비용이 별도 행');
  const c = f.data.series[0].col;
  assert.equal(inferPivotCategorySeries(f.wb, 1, chart({ r1: f.data.rows[0], c1: c, r2: f.data.rows.at(-1), c2: c })), null);
});

test('숫자·빈 레이블도 참조를 유지하고 원본 이름 수정·피벗 갱신을 새 범주로 읽음', () => {
  const f = fixture();
  const ch = chart({ r1: 1, c1: 1, r2: 2, c2: 1 });
  const explicit = { ...ch, series: inferPivotCategorySeries(f.wb, 1, ch) };
  f.wb.transact(() => { f.wb.setInput(1, 1, 0, ''); f.wb.setInput(1, 2, 0, '42'); });
  assert.deepEqual(chartModelData(f.wb, 1, explicit).categories, ['', '42']);
  f.wb.transact(() => f.wb.setInput(0, 2, 0, '다음'));
  f.render();
  assert.deepEqual(chartModelData(f.wb, 1, ch).categories, ['네이버', '다음']);
  assert.deepEqual(chartModelData(f.wb, 1, explicit).categories, ['네이버', '다음']);
});

test('피벗 위치·다른 시트 차트·추가 피벗 지원, 겹친 정의는 추측하지 않음', () => {
  const f = fixture(basicRows, { top: 4, left: 6 });
  const ch = chart({ r1: 5, c1: 8, r2: 6, c2: 8 }, { sheet: '피벗 결과' });
  const result = inferPivotCategorySeries(f.wb, 0, ch);
  assert.deepEqual(result[0].cat, ref(6, 5, 6));
  assert.deepEqual(result[0].val, ref(8, 5, 6));
  f.wb.transact(() => { f.wb.setSheetProp(1, 'pivot', null); f.wb.setSheetProp(1, 'pivotsExtra', [f.def]); });
  assert.deepEqual(inferPivotCategorySeries(f.wb, 0, ch), result);
  f.wb.transact(() => f.wb.setSheetProp(1, 'pivot', { ...f.def, name: '겹침' }));
  assert.equal(inferPivotCategorySeries(f.wb, 0, ch), null);
});

for (const explicit of [false, true]) test(`XLSX ${explicit ? '새 명시 계열' : '기존 범위'} 차트: 표준 cat 참조 및 값·범례 왕복`, () => {
  const { wb } = fixture();
  const ch = chart({ r1: 1, c1: 2, r2: 2, c2: 2 }, { threeD: true });
  if (explicit) ch.series = inferPivotCategorySeries(wb, 1, ch);
  const before = shape(chartModelData(wb, 1, ch));
  wb.transact(() => wb.setSheetProp(1, 'charts', [ch]));
  const bytes = writeXlsx(wb), files = unzip(bytes);
  const tree = parseXml(textOf(files['xl/charts/chart1.xml']));
  const cat = descendants(tree, 'cat')[0], val = descendants(tree, 'val')[0];
  assert.equal(descendants(cat, 'f')[0].text, "'피벗 결과'!$A$2:$A$3");
  assert.equal(descendants(val, 'f')[0].text, "'피벗 결과'!$C$2:$C$3");
  const restored = new Workbook(readXlsx(bytes).data);
  const saved = restored.sheets[1].charts[0];
  assert.deepEqual(saved.series[0].cat, ref(0));
  assert.deepEqual(saved.series[0].val, ref(2));
  assert.deepEqual(shape(chartModelData(restored, 1, saved)), before);
  assert.equal(ch.series != null, explicit, 'writer가 원래 차트를 변경하지 않음');
});
