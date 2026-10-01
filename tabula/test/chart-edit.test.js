import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chartSeriesPatch, chartPointColorPatch, chartExplosionPatch, chartPartDeletePatch, chartLayoutAfterDrag, chartExplosionAfterDrag } from '../src/chart-edit.js';
import { Workbook } from '../src/workbook.js';

test('필터된 계열의 원본 번호로 색·조각 분리를 편집하고 이웃 서식과 원본을 보존', () => {
  const chart = { hiddenSeries: [0], seriesFmt: [{ color: '#123456' }, { pointColors: { 2: '#abcdef' }, pointExplosion: { 1: 70 }, lineWidth: 3 }] };
  const original = structuredClone(chart);
  const colored = { ...chart, ...chartPointColorPatch(chart, 1, 0, '#ff0000') };
  assert.deepEqual(colored.seriesFmt[1].pointColors, { 0: '#ff0000', 2: '#abcdef' });
  const separated = { ...colored, ...chartExplosionPatch(colored, { kind: 'point', s: 1, p: 1 }, 0) };
  assert.equal(separated.seriesFmt[1].pointExplosion[1], 0);
  assert.equal(separated.seriesFmt[1].lineWidth, 3);
  assert.deepEqual(chart, original);
  assert.deepEqual(chartSeriesPatch(chart, -1, {}), {});
  assert.deepEqual(chartExplosionPatch(chart, { s: 1 }, NaN), {});
});
test('제목·범례 삭제는 해당 요소만, 점 Delete는 Excel처럼 계열만 제거', () => {
  const chart = { title: '매출', titleLayout: { x: .3, y: .4 }, legend: 'b', hiddenSeries: [1] };
  assert.deepEqual(chartPartDeletePatch(chart, { kind: 'title' }), { title: '', titleLayout: undefined });
  assert.deepEqual(chartPartDeletePatch(chart, { kind: 'legend' }), { legend: 'none', legendLayout: undefined });
  assert.deepEqual(chartPartDeletePatch(chart, { kind: 'point', s: 0, p: 2 }), { hiddenSeries: [1, 0] });
  assert.deepEqual(chartPartDeletePatch(chart, { kind: 'series', s: 1 }), { hiddenSeries: [1] });
  assert.equal(chartPartDeletePatch(chart, { kind: 'plot' }), null);
});
test('위치 변경은 차트 크기로 정규화하며 요소가 경계 밖으로 나가지 않는다', () => {
  const layout = { x: .2, y: .3 };
  const moved = chartLayoutAfterDrag(layout, [50, 40], 500, 400, { w: .3, h: .1 });
  assert.ok(Math.abs(moved.x - .3) < 1e-12 && Math.abs(moved.y - .4) < 1e-12);
  assert.deepEqual(chartLayoutAfterDrag(layout, [1000, -1000], 500, 400, { w: .3, h: .1 }), { x: .7, y: 0 });
  assert.deepEqual(layout, { x: .2, y: .3 });
});
test('원형 조각 드래그는 3D 투영을 역보정하고 접선 방향 움직임은 분리율을 바꾸지 않는다', () => {
  const geometry = { angle: Math.PI / 2, r: 100, squash: .5 };
  assert.ok(Math.abs(chartExplosionAfterDrag(20, [0, 25], geometry) - 70) < 1e-8);
  assert.ok(Math.abs(chartExplosionAfterDrag(20, [60, 0], geometry) - 20) < 1e-8);
  assert.equal(chartExplosionAfterDrag(20, [0, -100], geometry), 0);
  assert.equal(chartExplosionAfterDrag(20, [0, 1000], geometry), 400);
});
test('차트 요소 변경과 계열 삭제는 한 번 Undo로 복원하고 수식·셀 값을 바꾸지 않는다', () => {
  const wb = new Workbook({ sheets: [{ name: '합성', cells: { '0,0': { raw: '=1+1', cached: 2 } }, charts: [{ id: 'c', title: '매출', w: 500, h: 300, seriesFmt: [{ color: '#123456' }] }] }] });
  const before = structuredClone(wb.sheets[0].charts[0]), raw = wb.getRaw(0, 0, 0);
  const apply = patch => wb.transact(() => wb.setSheetProp(0, 'charts', [{ ...wb.sheets[0].charts[0], ...patch }]));
  apply(chartExplosionPatch(before, { kind: 'point', s: 0, p: 2 }, 80));
  assert.equal(wb.sheets[0].charts[0].seriesFmt[0].pointExplosion[2], 80);
  wb.undo(); assert.deepEqual(wb.sheets[0].charts[0], before); wb.redo();
  apply(chartPartDeletePatch(wb.sheets[0].charts[0], { kind: 'point', s: 0, p: 2 }));
  assert.equal(wb.sheets[0].charts.length, 1); assert.deepEqual(wb.sheets[0].charts[0].hiddenSeries, [0]);
  assert.equal(wb.getRaw(0, 0, 0), raw); wb.undo(); assert.equal(wb.sheets[0].charts[0].hiddenSeries, undefined);
});
