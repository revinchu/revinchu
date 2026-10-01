import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { chartData, chartModelData, renderChartSvg, chartValueScale, boxWhiskerStats } from '../src/chart.js';
import { surfaceGeometry, surfaceProjector } from '../src/chart-advanced.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { parseXml, descendants } from '../src/xml.js';
import { unzip, textOf } from '../src/zip.js';

const near = (a, b, epsilon = 1e-6) => assert.ok(Math.abs(a - b) < epsilon, `${a} ≈ ${b}`);
const data = { categories: ['가', '나', '다'], series: [{ name: '금액', values: [10, -4, 6] }] };
const render = (type, opts = {}, d = data) => parseXml(renderChartSvg({ type, w: 640, h: 440, legend: 'none', labels: false, ...opts }, d));
const nodes = (xml, tag, attr, value) => descendants(xml, tag).filter((n) => n.attrs[attr] === value);
const plots = (xml, type) => nodes(xml, 'g', 'data-plot', type)[0];
const chartShapes = (xml, type, tag = 'rect') => descendants(plots(xml, type), tag);
const statData = { categories: ['1', '2', '3', '4', '5', '6', '7'], series: [{ name: '분포', values: [1, 2, 3, 4, 5, 6, 30] }] };

test('표면 와이어프레임: 수동 값 축을 벗어난 선분도 경계에서 보간하여 자름', () => {
  const d = { categories: ['a', 'b'], series: [{ name: 'r1', values: [-10, 10] }, { name: 'r2', values: [20, 30] }] };
  const chart = { type: 'surface', surfaceStyle: 'wireframe', axes: { y: { min: 0, max: 5 } } }, g = surfaceGeometry(d, chart);
  assert.equal(g.mesh.length, 2);
  assert.deepEqual(g.mesh[0], [{ x: .5, y: 0, z: 0 }, { x: .75, y: 0, z: 5 }]);
  for (const line of g.mesh) for (const p of line) assert.ok(p.z >= 0 && p.z <= 5);
  const xml = render('surface', chart, d);
  assert.equal(nodes(xml, 'path', 'data-surface-line', 'mesh').length, 2);
});

test('상자수염: 홀수 중앙값 포함/제외와 짝수 표본의 절반 중앙값', () => {
  const odd = [1, 2, 4, 7, 10, 20, 30], even = [1, 2, 4, 7, 10, 20];
  const inc = boxWhiskerStats(odd), exc = boxWhiskerStats(odd, 'exclusive');
  assert.deepEqual([inc.q1, inc.med, inc.q3], [3, 7, 15]);
  assert.deepEqual([exc.q1, exc.med, exc.q3], [2, 7, 20]);
  for (const method of ['inclusive', 'exclusive']) {
    const s = boxWhiskerStats(even, method); assert.deepEqual([s.q1, s.med, s.q3], [2, 5.5, 10]);
  }
  assert.deepEqual(odd, [1, 2, 4, 7, 10, 20, 30], '입력을 정렬/변경하지 않음');
});

test('상자수염: Tukey 1.5 IQR 수염과 이상값, 빈/단일/동일 표본', () => {
  const s = boxWhiskerStats([1, 2, 3, 4, 5, 6, 30, null, '50', NaN]);
  assert.deepEqual([s.q1, s.q3, s.lo, s.hi], [2.5, 5.5, 1, 6]);
  assert.deepEqual(s.out, [30]); assert.deepEqual(s.inside, [1, 2, 3, 4, 5, 6]);
  near(s.mean, 51 / 7); assert.equal(s.count, 7);
  assert.equal(boxWhiskerStats([null, NaN, '']), null);
  for (const method of ['inclusive', 'exclusive']) for (const values of [[9], [9, 9], [9, 9, 9]]) {
    const s = boxWhiskerStats(values, method); assert.deepEqual([s.q1, s.med, s.q3, s.lo, s.hi], [9, 9, 9, 9, 9]); assert.deepEqual(s.out, []);
  }
});

test('상자수염: 평균/내부점/이상값 옵션은 실제 SVG 요소와 값에 반영', () => {
  const base = render('boxWhisker', {}, statData), hidden = render('boxWhisker', { showMean: false, showOutliers: false, showInnerPoints: true }, statData);
  assert.equal(nodes(base, 'path', 'data-box', 'mean').length, 1);
  assert.equal(nodes(base, 'circle', 'data-box', 'outlier')[0].attrs['data-value'], '30');
  assert.equal(nodes(base, 'circle', 'data-box', 'inner').length, 0);
  assert.equal(nodes(hidden, 'path', 'data-box', 'mean').length, 0);
  assert.equal(nodes(hidden, 'circle', 'data-box', 'outlier').length, 0);
  assert.deepEqual(nodes(hidden, 'circle', 'data-box', 'inner').map((n) => +n.attrs['data-value']), [1, 2, 3, 4, 5, 6]);
  assert.equal(nodes(base, 'line', 'data-box', 'median')[0].attrs.y1, nodes(hidden, 'line', 'data-box', 'median')[0].attrs.y1, '표시 옵션으로 원자료 축 범위가 바뀌지 않음');
  const exc = render('boxWhisker', { quartileMethod: 'exclusive' }, statData);
  assert.equal(nodes(exc, 'rect', 'data-box', 'quartiles')[0].attrs['data-q1'], '2');
  assert.equal(nodes(exc, 'rect', 'data-box', 'quartiles')[0].attrs['data-q3'], '6');
});

test('통계 값 축: 수동 범위·주 단위·잘못된 범위·작은 주 단위의 유한 처리', () => {
  const s = chartValueScale([2, 9], { min: -2, max: 12, major: 2, numFmt: '0.00', reverse: true });
  assert.deepEqual(s.ticks, [-2, 0, 2, 4, 6, 8, 10, 12]); assert.equal(s.reverse, true); assert.equal(s.code, '0.00');
  for (const cfg of [{ min: 10, max: -10 }, { min: 0, max: 0 }, { major: 1e-99 }]) {
    const scale = chartValueScale([0, 4, -6], cfg); assert.ok(scale.max > scale.min); assert.ok(scale.ticks.length <= 201); assert.ok(scale.ticks.every(Number.isFinite));
  }
});

for (const type of ['waterfall', 'histogram', 'pareto', 'boxWhisker', 'stock']) {
  test(`${type}: 축 hide/제목/주 단위/표시형식 및 역순 막대 기하`, () => {
    const d = type === 'boxWhisker' ? statData : type === 'stock' ? { categories: ['a', 'b'], series: [{ values: [12, 16] }, { values: [3, 4] }, { values: [8, 10] }] } : data;
    const axes = { y: { min: -10, max: 40, major: 10, numFmt: '0.00', title: '금액 축' }, x: { title: '항목 축' } };
    const normal = render(type, { axes, gridX: true }, d), reversed = render(type, { axes: { ...axes, y: { ...axes.y, reverse: true } } }, d);
    assert.deepEqual(nodes(normal, 'text', 'data-axis', 'y').map((n) => n.text), ['-10.00', '0.00', '10.00', '20.00', '30.00', '40.00']);
    assert.deepEqual(nodes(normal, 'text', 'data-axis-title', 'y').map((n) => n.text), ['금액 축']);
    assert.deepEqual(nodes(normal, 'text', 'data-axis-title', 'x').map((n) => n.text), ['항목 축']);
    assert.ok(nodes(normal, 'line', 'data-grid', 'x').length > 0);
    const a = nodes(normal, 'text', 'data-axis', 'y'), b = nodes(reversed, 'text', 'data-axis', 'y');
    assert.ok(+a[0].attrs.y > +a.at(-1).attrs.y); assert.ok(+b[0].attrs.y < +b.at(-1).attrs.y);
    const rectangles = chartShapes(reversed, type); for (const r of rectangles) assert.ok(+r.attrs.height >= 0);
    if (rectangles.length) { const orig = chartShapes(normal, type); orig.forEach((r, i) => near(+r.attrs.height, +rectangles[i].attrs.height, .11)); }
    const hidden = render(type, { axes: { x: { hide: true, title: '감춰진 항목' }, y: { ...axes.y, hide: true }, y2: { hide: true } }, gridY: false }, d);
    assert.equal(nodes(hidden, 'text', 'data-axis', 'x').length, 0); assert.equal(nodes(hidden, 'text', 'data-axis', 'y').length, 0);
    assert.equal(nodes(hidden, 'text', 'data-axis-title', 'y').length, 0); assert.equal(nodes(hidden, 'line', 'data-grid', 'y').length, 0);
  });
}

test('폭포: 양수/음수/합계 금액과 항목 역순·간격을 보존', () => {
  const d = { categories: ['시작', '증가', '감소', '합계'], series: [{ name: '금액', values: [20, 8, -10, 18] }] };
  const a = chartShapes(render('waterfall', { totals: [0, 3], gap: 0 }, d), 'waterfall');
  const b = chartShapes(render('waterfall', { totals: [0, 3], gap: 300, axes: { x: { reverse: true } } }, d), 'waterfall');
  assert.deepEqual(a.map((n) => +n.attrs['data-value']), [20, 8, -10, 18]);
  assert.ok(+a[0].attrs.x < +a.at(-1).attrs.x); assert.ok(+b[0].attrs.x > +b.at(-1).attrs.x);
  near(+a[0].attrs.width / +b[0].attrs.width, 4, .02);
});

test('파레토: 정렬 합계·누적 보조 축의 25% 눈금/역순/제목', () => {
  const d = { categories: ['A', 'B', 'C'], series: [{ name: '빈도', values: [1, 6, 3] }] };
  const xml = render('pareto', { axes: { y2: { title: '누적률', reverse: true, major: .25, numFmt: '0.0%' } } }, d);
  assert.deepEqual(chartShapes(xml, 'pareto').map((n) => +n.attrs['data-value']), [6, 3, 1]);
  const ticks = nodes(xml, 'text', 'data-axis', 'y2'); assert.deepEqual(ticks.map((n) => n.text), ['0.0%', '25.0%', '50.0%', '75.0%', '100.0%']);
  assert.ok(+ticks[0].attrs.y < +ticks.at(-1).attrs.y); assert.equal(nodes(xml, 'text', 'data-axis-title', 'y2')[0].text, '누적률');
  const p = nodes(xml, 'path', 'data-pareto', 'cumulative')[0].attrs.d.match(/[-\d.]+/g).map(Number);
  assert.ok(p[1] < p[3] && p[3] < p[5], '역순 누적선이 아래쪽 100%로 진행');
});

test('특수 차트의 데이터 표는 집계값/통계/거래량과 가격을 실제 표시', () => {
  for (const type of ['waterfall', 'histogram', 'pareto', 'boxWhisker', 'stock']) {
    const d = type === 'boxWhisker' ? statData : data;
    const xml = render(type, { dataTable: true }, d), table = nodes(xml, 'g', 'data-el', 'dataTable')[0];
    assert.ok(table); assert.ok(descendants(table, 'text').length >= 4);
    assert.equal(nodes(render(type, {}, d), 'g', 'data-el', 'dataTable').length, 0);
    const bounds = descendants(table, 'rect')[0].attrs; assert.ok(+bounds.y + +bounds.height <= 440);
  }
  const table = nodes(render('boxWhisker', { dataTable: true }, statData), 'g', 'data-el', 'dataTable')[0];
  assert.ok(descendants(table, 'text').some((n) => n.text === 'Q1')); assert.ok(descendants(table, 'text').some((n) => n.text === '2.5'));
});

const surfaceData = chartData([['Y/X', 'a', 'b', 'c'], ['r1', 0, 2, 6], ['r2', 4, 9, 5], ['r3', 7, 3, 1]], 'surface');
test('표면 직각 투영: X/높이 축 수직 유지, 깊이는 깊이 축만 비례 확대', () => {
  const g = surfaceGeometry(surfaceData), a = surfaceProjector(g, { view3D: { rotX: 30, rotY: 40, depthPercent: 100 } }), b = surfaceProjector(g, { view3D: { rotX: 30, rotY: 40, depthPercent: 300 } });
  const dx = (p, i, j) => p.project(j).map((v, k) => v - p.project(i)[k]);
  const origin = { x: 0, y: 0, z: g.min }, x = { ...origin, x: 2 }, y = { ...origin, y: 2 }, z = { ...origin, z: g.max };
  near(dx(a, origin, x)[1], 0); near(dx(a, origin, z)[0], 0);
  near(dx(b, origin, y)[0], dx(a, origin, y)[0] * 3); near(dx(b, origin, y)[1], dx(a, origin, y)[1] * 3);
  for (let i = 0; i < 2; i++) near(dx(a, origin, z)[i], dx(b, origin, z)[i]);
});

test('표면 원근 투영: 멀리 있는 동일 길이를 작게 그리고 극단 회전/깊이도 유한', () => {
  const g = surfaceGeometry(surfaceData), view3D = { rotX: 20, rotY: 0, rAngAx: false, perspective: 180 }, p = surfaceProjector(g, { view3D });
  const length = (y) => { const a = p.project({ x: 0, y, z: g.min }), b = p.project({ x: 2, y, z: g.min }); return Math.hypot(b[0] - a[0], b[1] - a[1]); };
  assert.ok(length(2) < length(0), '뒤쪽 X 간격이 앞쪽보다 작음');
  const polygons = (options) => descendants(render('surface', options, surfaceData), 'polygon').map((n) => n.attrs.points).join('|');
  const base = polygons({ view3D });
  for (const patch of [{ perspective: 20 }, { depthPercent: 300 }, { rAngAx: true }]) assert.notEqual(polygons({ view3D: { ...view3D, ...patch } }), base);
  for (const rotX of [-90, 0, 90]) for (const rotY of [0, 90, 180, 270]) {
    const xml = renderChartSvg({ type: 'surface', w: 640, h: 440, view3D: { rotX, rotY, depthPercent: 2000, perspective: 240, rAngAx: false } }, surfaceData);
    assert.doesNotMatch(xml, /NaN|Infinity/);
    for (const poly of descendants(parseXml(xml), 'polygon')) { const nums = poly.attrs.points.split(/[ ,]+/).map(Number); nums.forEach((v, i) => assert.ok(v >= 0 && v <= (i % 2 ? 440 : 640))); }
  }
});

test('상자수염 XLSX: 기본 평균/명시 옵션 저장 후 표시와 통계값까지 보존', () => {
  for (const options of [{}, { showMean: false, showOutliers: false, showInnerPoints: true, quartileMethod: 'exclusive', dataTable: true, axes: { y: { reverse: true, title: '분포', numFmt: '0.00' } } }]) {
    const wb = new Workbook(); wb.sheets[0].charts = [{ type: 'boxWhisker', id: 'box', x: 0, y: 0, w: 640, h: 440, ...options, series: [{ name: { text: '분포' }, cache: statData.series[0].values }] }];
    const chart = wb.sheets[0].charts[0], bytes = writeXlsx(wb), loaded = new Workbook(readXlsx(bytes).data), restored = loaded.sheets[0].charts[0];
    assert.equal(restored.showMean, options.showMean !== false);
    const before = parseXml(renderChartSvg(chart, chartModelData(wb, 0, chart))), after = parseXml(renderChartSvg(restored, chartModelData(loaded, 0, restored)));
    for (const [tag, kind] of [['path', 'mean'], ['circle', 'outlier'], ['circle', 'inner'], ['rect', 'quartiles']]) assert.deepEqual(nodes(after, tag, 'data-box', kind).map((n) => [n.attrs['data-value'], n.attrs['data-q1'], n.attrs['data-q3']]), nodes(before, tag, 'data-box', kind).map((n) => [n.attrs['data-value'], n.attrs['data-q1'], n.attrs['data-q3']]));
    if (process.env.CHARTEX_FIXTURE_DIR) {
      mkdirSync(process.env.CHARTEX_FIXTURE_DIR, { recursive: true }); const name = options.showMean === false ? 'box-stat-options' : 'box-stat-default';
      writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, `${name}.xlsx`), bytes);
      writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, `${name}.xml`), textOf(unzip(bytes)['xl/charts/chart1.xml']));
    }
  }
});
