import { test } from 'node:test';
import assert from 'node:assert/strict';
import { chartPartDeletePatch } from '../src/chart-edit.js';
import { renderChartSvg } from '../src/chart.js';

test('축 Delete는 범위/다른 축 서식을 유지하고 SVG 눈금만 숨긴다', () => {
  const chart = { type: 'column', w: 500, h: 300, axes: { y: { min: 0, max: 100, major: 20 }, x: { title: '분류' } } };
  const next = { ...chart, ...chartPartDeletePatch(chart, { kind: 'axis-y' }) };
  assert.deepEqual(next.axes.y, { min: 0, max: 100, major: 20, hide: true }); assert.deepEqual(next.axes.x, chart.axes.x);
  const data = { categories: ['가'], series: [{ name: '매출', values: [20] }] };
  assert.match(renderChartSvg(chart, data), /data-el="axis-y"/); assert.doesNotMatch(renderChartSvg(next, data), /data-el="axis-y"/);
});
test('계층 레이블 Delete는 기본 항목명과 명시 값/백분율/계열명 모두 제거한다', () => {
  const chart = { type: 'sunburst', w: 500, h: 350, legend: 'none', seriesFmt: [{ labels: true, catName: true, serName: true, pct: true }] };
  const patch = chartPartDeletePatch(chart, { kind: 'label', s: 0 });
  assert.deepEqual(patch.seriesFmt[0], { labels: false, catName: false, serName: false, pct: false });
  const data = { categories: ['가', '나'], series: [{ name: '매출', values: [20, 30], ...patch.seriesFmt[0], _fi: 0 }] };
  assert.doesNotMatch(renderChartSvg({ ...chart, ...patch }, data), /data-el="label"/);
});
