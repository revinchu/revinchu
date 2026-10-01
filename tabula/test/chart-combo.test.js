import test from 'node:test';
import assert from 'node:assert/strict';
import { chartStackValues, resolveChart, renderChartSvg, CHART_GALLERY } from '../src/chart.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { parseXml, descendants } from '../src/xml.js';

const rows = [['항목', '매출', '비율'], ['A', 1000, 10], ['B', 2000, 20], ['C', 1500, 15]];
const api = { range: () => rows };
const rects = (svg, si) => descendants(parseXml(svg), 'rect').filter((n) => n.attrs['data-s'] === String(si));

test('콤보 누적: 축과 계열 종류별로 독립 누적하고 양수·음수를 분리', () => {
  const data = [
    { type: 'column', axis: 0, values: [10, -5, 0, null] },
    { type: 'column', axis: 0, values: [30, -15, 0, 4] },
    { type: 'column', axis: 1, values: [1000, -200, 0, 8] },
    { type: 'line', axis: 0, values: [2, 4, 0, 6] },
  ];
  const before = structuredClone(data), got = chartStackValues(data, 'stacked', 4);
  assert.deepEqual(got[1]._lower, [10, -5, 0, 0]);
  assert.deepEqual(got[1]._upper, [40, -20, 0, 4]);
  assert.deepEqual(got[2]._lower, [0, 0, 0, 0]);
  assert.deepEqual(got[2]._upper, [1000, -200, 0, 8]);
  assert.deepEqual(got[3]._upper, [2, 4, 0, 6]);
  assert.deepEqual(data, before, '원본 계열 값과 서식을 변경하지 않음');
});

test('100% 누적: 보조축의 큰 값이 주축의 비율 분모를 바꾸지 않음', () => {
  const got = chartStackValues([
    { type: 'column', values: [10, -10, 0, null] },
    { type: 'column', values: [30, 30, 0, null] },
    { type: 'column', axis: 'secondary', values: [1e12, -1e12, 0, null] },
    { type: 'line', axis: 0, values: [500, 500, 0, null] },
  ], 'percentStacked', 4);
  assert.deepEqual(got[0]._upper, [0.25, -0.25, 0, null]);
  assert.deepEqual(got[1]._upper, [1, 0.75, 0, null]);
  assert.deepEqual(got[2]._upper, [1, -1, 0, null]);
  assert.deepEqual(got[3]._upper, [1, 1, 0, null]);
  assert.equal(got[2].axis, 1);
});

test('콤보 기본 축·보조 축 프리셋과 명시 축 0을 구분', () => {
  assert.deepEqual(resolveChart({ type: 'combo' }, api).series.map((s) => [s.type, s.axis]), [['column', 0], ['line', 1]]);
  assert.deepEqual(resolveChart({ type: 'combo', comboAxis: 'primary' }, api).series.map((s) => [s.type, s.axis]), [['column', 0], ['line', 0]]);
  assert.equal(resolveChart({ type: 'combo', seriesFmt: [{}, { type: 'line', axis: 0 }] }, api).series[1].axis, 0);
  assert.equal(resolveChart({ type: 'combo', seriesFmt: [{ axis: 'secondary' }, {}] }, api).series[0].axis, 1);
});

test('보조 축 막대는 자기 축의 0에서 시작하고 축 반전을 독립 적용', () => {
  const data = { categories: ['A', 'B'], series: [
    { name: '주축', type: 'line', values: [100, 200], axis: 0 },
    { name: '보조축', type: 'column', values: [10, 20], axis: 1 },
  ] };
  const ch = { type: 'combo', w: 500, h: 320, legend: 'none', axes: { y: { min: 100, max: 200 }, y2: { min: 0, max: 20 } } };
  const normal = rects(renderChartSvg(ch, data), 1);
  assert.equal(normal.length, 2);
  for (const n of normal) {
    assert.ok(Number(n.attrs.y) >= 0);
    assert.ok(Number(n.attrs.y) + Number(n.attrs.height) <= 320);
  }
  const reversed = rects(renderChartSvg({ ...ch, axes: { ...ch.axes, y2: { ...ch.axes.y2, reverse: true } } }, data), 1);
  assert.ok(Number(reversed[0].attrs.y) < Number(normal[0].attrs.y));
  assert.equal(reversed[0].attrs.height, normal[0].attrs.height);
});

test('보조 축만 있는 차트·영 값·큰 값·잘못된 수동 축 범위도 유효 SVG', () => {
  for (const values of [[0, 0], [-1e12, 1e12], [10, 20]]) {
    for (const grouping of ['clustered', 'stacked', 'percentStacked']) {
      const ch = { type: 'combo', w: 500, h: 320, grouping, axes: { y2: { min: 0, max: 0, title: '보조 축' } } };
      const svg = renderChartSvg(ch, { categories: ['A', 'B'], series: [{ name: '계열', type: 'column', axis: 'secondary', values }] });
      assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
      assert.equal(parseXml(svg).name, 'svg');
      assert.match(svg, /보조 축/);
      assert.ok(descendants(parseXml(svg), 'text').some((n) => /^rotate\(90 /.test(n.attrs.transform ?? '')), '보조 축만 있어도 축 제목은 오른쪽');
    }
  }
});

test('다중 계열 도넛: 세 계열이 서로 다른 링이며 완전한 링도 앞 계열을 가리지 않음', () => {
  const data = { categories: ['A'], series: [1, 2, 3].map((v, i) => ({ name: `계열${i}`, values: [v], labels: false })) };
  const svg = renderChartSvg({ type: 'doughnut', w: 500, h: 320 }, data);
  const paths = descendants(parseXml(svg), 'path').filter((n) => n.attrs['data-s'] !== undefined);
  assert.deepEqual(paths.map((n) => n.attrs['data-s']), ['0', '1', '2']);
  assert.ok(paths.every((n) => n.attrs['fill-rule'] === 'evenodd'));
  const radii = paths.map((n) => Number(n.attrs.d.match(/A([\d.]+),/)[1]));
  assert.ok(radii[0] < radii[1] && radii[1] < radii[2]);
  assert.equal(descendants(parseXml(svg), 'circle').length, 0);
});

test('도넛 계열은 XLSX 저장·다시 열기 후에도 모두 렌더링', () => {
  const wb = new Workbook();
  rows.forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value))));
  wb.sheets[0].charts = [{ id: 'rings', type: 'doughnut', w: 500, h: 320, x: 0, y: 0, range: { r1: 0, c1: 0, r2: 3, c2: 2 } }];
  const back = new Workbook(readXlsx(writeXlsx(wb)).data), ch = back.sheets[0].charts[0];
  const data = resolveChart(ch, { values: (ref) => {
    const out = [];
    for (let r = ref.r1; r <= ref.r2; r++) { const row = []; for (let c = ref.c1; c <= ref.c2; c++) row.push(back.getValue(0, r, c)); out.push(row); }
    return out;
  }, range: () => rows });
  assert.equal(ch.type, 'doughnut');
  assert.equal(data.series.length, 2);
  const svg = renderChartSvg(ch, data);
  assert.match(svg, /data-s="0"/);
  assert.match(svg, /data-s="1"/);
});

test('갤러리의 실제 변형: 표식 100% 꺾은선·쪼개진 도넛·기본 축 콤보', () => {
  const variants = CHART_GALLERY.flatMap(([, entries]) => entries).map(([, patch]) => patch);
  assert.ok(variants.some((p) => p.type === 'line' && p.grouping === 'percentStacked' && p.marker === 'circle'));
  assert.ok(variants.some((p) => p.type === 'doughnut' && p.explode > 0));
  assert.ok(variants.some((p) => p.type === 'combo' && p.comboAxis === 'primary'));
  assert.ok(variants.some((p) => p.type === 'combo' && p.comboAxis === 'secondary'));
});
