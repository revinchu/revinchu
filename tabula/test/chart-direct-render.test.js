import test from 'node:test';
import assert from 'node:assert/strict';
import { renderChartSvg, resolveChart, filterChart, chartPointIndex, chartPointExplosion } from '../src/chart.js';
import { parseXml, descendants } from '../src/xml.js';

const data = () => ({ categories: ['첫 항목', '둘째 항목', '셋째 항목', '넷째 항목'], series: [{ name: '매출', values: [10, 20, 30, 40], x: [1, 2, 3, 4], size: [1, 2, 3, 4], labels: false }] });
const draw = (chart = {}, source = data()) => parseXml(renderChartSvg({ type: 'pie', w: 500, h: 320, ...chart }, source));
const shapes = (svg) => ['path', 'rect', 'circle', 'ellipse', 'polygon'].flatMap((name) => descendants(svg, name));
const pointShapes = (svg) => shapes(svg).filter((n) => n.attrs['data-p'] !== undefined);
const group = (svg, kind) => descendants(svg, 'g').find((n) => n.attrs['data-el'] === kind);
const piePoints = (svg) => pointShapes(svg).filter((n) => n.attrs['data-pie-r'] !== undefined);
const near = (a, b, message) => assert.ok(Math.abs(a - b) < 1e-6, `${message}: ${a} ≠ ${b}`);

test('수동 제목·범례는 차트 크기 기준 좌상단이며 그림 영역 뒤로 가리지 않는다', () => {
  for (const [w, h] of [[500, 320], [1000, 640]]) {
    const svg = draw({ w, h, title: '매출 보고서', titleLayout: { x: .15, y: .2, w: .4, h: .1 }, legend: 'r', legendLayout: { x: .6, y: .35, w: .3, h: .5 }, plotFill: '#ddd' });
    const title = group(svg, 'title'), legend = group(svg, 'legend');
    assert.equal(title.attrs.transform, `translate(${w * .15},${h * .2})`);
    assert.equal(legend.attrs.transform, `translate(${w * .6},${h * .35})`);
    for (const node of [title, legend]) {
      const box = descendants(node, 'rect')[0], text = descendants(node, 'text')[0];
      assert.equal(box.attrs.x, '0'); assert.equal(box.attrs.y, '0');
      assert.equal(text.attrs.y, '0'); assert.equal(text.attrs['dominant-baseline'], 'text-before-edge');
    }
    const html = renderChartSvg({ w, h, type: 'pie', title: '매출', titleLayout: { x: .1, y: .2 }, legendLayout: { x: .4, y: .2 }, plotFill: '#ddd' }, data());
    assert.ok(html.indexOf('data-el="title"') > html.lastIndexOf('data-pie-r'));
    assert.ok(html.indexOf('data-el="legend"') > html.lastIndexOf('data-pie-r'));
  }
});

test('잘못된 수동 위치는 자동 위치로 복원하고 명시 범례 없음은 유지한다', () => {
  for (const layout of [{ x: NaN, y: .2 }, { x: .2 }, { x: Infinity, y: .1 }]) {
    const svg = draw({ title: '자동 제목', titleLayout: layout, legendLayout: layout });
    assert.equal(group(svg, 'title').attrs.transform, undefined);
    assert.equal(group(svg, 'legend').attrs.transform, undefined);
  }
  assert.equal(group(draw({ legend: 'none', legendLayout: { x: 0, y: 0 } }), 'legend'), undefined);
});

test('계열·범주 필터 후 원본 point/series 인덱스와 요소별 색은 유지된다', () => {
  const source = data(); source.series.push({ ...source.series[0], name: '이익', pointColors: { 0: '#a00', 2: '#0a0', 3: '#00a' } });
  const before = structuredClone(source), filtered = filterChart(source, { hiddenSeries: [0], hiddenCats: [1] });
  assert.equal(filtered.series[0]._fi, 1);
  assert.deepEqual(filtered.series[0]._pi, [0, 2, 3]);
  assert.deepEqual(filtered.series[0].values, [10, 30, 40]);
  assert.equal(chartPointIndex(filtered.series[0], 1), 2);
  for (const [type, extra] of [['column', {}], ['bar', {}], ['pie', {}], ['pie', { threeD: true }], ['doughnut', {}], ['line', {}], ['scatter', {}], ['radar', { radarStyle: 'marker' }], ['bubble', {}]]) {
    const svg = draw({ type, ...extra }, filtered), points = pointShapes(svg);
    assert.ok(points.length >= 3, type);
    assert.deepEqual([...new Set(points.map((n) => n.attrs['data-p']))].sort(), ['0', '2', '3'], type);
    assert.ok(points.every((n) => n.attrs['data-s'] === '1'), type);
    assert.ok(points.filter((n) => n.attrs['data-p'] === '2').every((n) => n.attrs.fill === '#0a0'), type);
  }
  assert.deepEqual(source, before, '필터와 렌더는 원본 값·색 객체를 변경하지 않음');
});

test('원형 분리 우선순위는 개별 요소→계열→차트이며 0과 400 경계를 구분한다', () => {
  const chart = { explode: 12 }, series = { explode: 30, pointExplosion: { 0: 0, 2: 999, 3: -10 }, _pi: [0, 2, 3, 5] };
  assert.equal(chartPointExplosion(chart, series, 0), 0);
  assert.equal(chartPointExplosion(chart, series, 1), 400);
  assert.equal(chartPointExplosion(chart, series, 2), 0);
  assert.equal(chartPointExplosion(chart, series, 3), 30);
  assert.equal(chartPointExplosion(chart, {}, 0), 12);
  assert.equal(chartPointExplosion(chart, { pointExplosion: { 0: NaN } }, 0), 0);
});

test('2D 원형·도넛 개별 조각은 정확한 반경 비율만큼 이동하고 드래그 metadata를 제공한다', () => {
  for (const type of ['pie', 'doughnut']) {
    const source = data(); source.series[0].pointExplosion = { 1: 80 };
    const before = structuredClone(source), svg = draw({ type }, source);
    const groups = descendants(svg, 'g').filter((n) => /^translate/.test(n.attrs.transform ?? '') && piePoints(n).length);
    assert.equal(groups.length, 1, type);
    const point = piePoints(groups[0])[0], a = Number(point.attrs['data-pie-angle']), r = Number(point.attrs['data-pie-r']);
    assert.equal(point.attrs['data-p'], '1');
    const [x, y] = groups[0].attrs.transform.match(/[-\d.]+/g).map(Number);
    assert.ok(Math.abs(x - Math.cos(a) * r * .8) < .006);
    assert.ok(Math.abs(y - Math.sin(a) * r * .8) < .006);
    assert.equal(point.attrs['data-pie-squash'], '1');
    assert.ok(piePoints(svg).every((n) => Number.isFinite(Number(n.attrs['data-pie-cx'])) && Number.isFinite(Number(n.attrs['data-pie-cy']))));
    assert.deepEqual(source, before);
  }
});

test('3D 원형 개별 조각은 투영된 중심과 Y 압축을 기준으로 이동한다', () => {
  const source = data(); source.series[0].pointExplosion = { 1: 75 };
  const svg = draw({ threeD: true, view3D: { rotX: 30, rotY: 20 } }, source);
  const paths = piePoints(svg).filter((n) => n.name === 'path' && n.attrs['data-p'] === '1' && n.attrs.stroke === '#fff');
  assert.equal(paths.length, 1);
  const n = paths[0], a = Number(n.attrs['data-pie-angle']), r = Number(n.attrs['data-pie-r']), squash = Number(n.attrs['data-pie-squash']);
  const [x, y] = n.attrs.d.match(/^M([-\d.]+),([-\d.]+)/).slice(1).map(Number);
  assert.ok(Math.abs(x - (Number(n.attrs['data-pie-cx']) + Math.cos(a) * r * .75)) < .006);
  assert.ok(Math.abs(y - (Number(n.attrs['data-pie-cy']) + Math.sin(a) * r * .75 * squash)) < .006);
  near(squash, .5, '3D 투영');
});

test('다중 도넛은 같은 외곽 기준 반경을 쓰고 계열별·요소별 분리를 독립 적용한다', () => {
  const source = data(); source.series[0].explode = 20;
  source.series.push({ ...structuredClone(source.series[0]), name: '둘째 계열', explode: 40, pointExplosion: { 2: 0 } });
  const svg = draw({ type: 'doughnut' }, source), points = piePoints(svg);
  const radii = [0, 1].map((si) => [...new Set(points.filter((n) => n.attrs['data-s'] === String(si)).map((n) => Number(n.attrs['data-pie-r'])))]);
  assert.equal(radii[0].length, 1); assert.equal(radii[1].length, 1); assert.ok(radii[0][0] < radii[1][0]);
  const shifted = descendants(svg, 'g').filter((g) => /^translate/.test(g.attrs.transform ?? '') && piePoints(g).length);
  assert.equal(shifted.length, 7, '첫 계열4개 + 둘째 계열3개만 분리');
});

test('명시 원형 계열 색은 전체 조각·범례에 적용하되 요소 색이 우선한다', () => {
  for (const type of ['pie', 'doughnut']) {
    const chart = { type, seriesFmt: [{ color: '#123456', pointColors: { 1: '#abcdef' } }] };
    const resolved = resolveChart(chart, { range: () => [['항목', '값'], ['A', 10], ['B', 20], ['C', 30]] });
    const svg = draw(chart, resolved);
    assert.deepEqual(piePoints(svg).map((n) => n.attrs.fill), ['#123456', '#abcdef', '#123456']);
    assert.deepEqual(descendants(group(svg, 'legend'), 'rect').map((n) => n.attrs.fill), ['#123456', '#abcdef', '#123456']);
  }
  assert.ok(new Set(piePoints(draw()).map((n) => n.attrs.fill)).size > 1, '서식이 없으면 기존 요소별 팔레트 유지');
});

test('전체 계열 숨김·0 값·큰 분리도 유효 SVG이고 제목 수동 위치는 유지된다', () => {
  for (const type of ['pie', 'doughnut']) {
    const source = data(); source.series[0].pointExplosion = { 0: 400, 1: 400 };
    source.series[0].values = [0, 40, 0, 0];
    const svg = renderChartSvg({ type, w: 120, h: 90, title: '보고서', titleLayout: { x: .1, y: .2 } }, source);
    assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
    assert.equal(piePoints(parseXml(svg)).length, 1);
  }
  const svg = draw({ title: '빈 차트', titleLayout: { x: .2, y: .3 } }, filterChart(data(), { hiddenSeries: [0] }));
  assert.ok(group(svg, 'title'));
  assert.equal(piePoints(svg).length, 0);
});

test('기본형 데이터 레이블은 대응 원본 요소를 선택할 수 있는 표식을 가진다', () => {
  const source = data(); source.series[0].labels = true;
  const filtered = filterChart(source, { hiddenCats: [1] });
  for (const type of ['pie', 'column', 'bar', 'line', 'scatter']) {
    const svg = draw({ type }, filtered);
    const labels = descendants(svg, 'text').filter((n) => n.attrs['data-el'] === 'label');
    assert.deepEqual(labels.map((n) => n.attrs['data-p']).sort(), ['0', '2', '3'], type);
    assert.ok(labels.every((n) => n.attrs['data-s'] === '0'));
  }
});
