import test from 'node:test';
import assert from 'node:assert/strict';
import { chartData, resolveChart, renderChartSvg, CHART_GALLERY, CHART_TYPES } from '../src/chart.js';
import { sunburstNodes, annularSector, splitPieData, surfaceGeometry, stockVolumeData } from '../src/chart-advanced.js';
import { parseXml, descendants } from '../src/xml.js';

const near = (a, b, epsilon = 1e-9) => assert.ok(Math.abs(a - b) < epsilon, `${a} ≈ ${b}`);
const area = (points) => Math.abs(points.reduce((sum, p, i) => { const q = points[(i + 1) % points.length]; return sum + p.x * q.y - q.x * p.y; }, 0)) / 2;
const hierarchyRows = [['상위', '하위', '값'], ['A', '공통', 20], ['A', '전용', 30], ['B', '공통', 50]];
const surfaceRows = [['Y/X', 0, 10, 20], [0, 0, 10, 20], [10, 10, 20, 30], [20, 20, 30, 40]];

test('선버스트: 같은 이름의 자식도 부모별로 분리하고 부모 각도 안에 배치', () => {
  const data = chartData(hierarchyRows, 'sunburst'), { nodes, total, depth } = sunburstNodes(data);
  assert.equal(total, 100); assert.equal(depth, 2);
  assert.deepEqual(nodes.filter((n) => n.depth === 1).map((n) => [n.name, n.value]), [['A', 50], ['B', 50]]);
  assert.equal(nodes.filter((n) => n.name === '공통').length, 2);
  for (const parent of nodes.filter((n) => n.depth === 1)) {
    near(parent.end - parent.start, Math.PI);
    for (const child of nodes.filter((n) => n.depth === 2 && n.rootIndex === parent.rootIndex)) assert.ok(child.start >= parent.start && child.end <= parent.end + 1e-10);
  }
  const leaf = nodes.find((n) => n.name === '전용'); near(leaf.end - leaf.start, 2 * Math.PI * .3);
});

test('선버스트: 숫자 계층·빈 하위 가지·0/음수/빈 값을 구분하고 입력을 보존', () => {
  const rows = [['연도', '분류', '값'], [2025, '', 10], [2025, 'A', 20], [2026, 'B', 0], [2026, 'C', -2], [2026, 'D', null]];
  const before = structuredClone(rows), data = chartData(rows, 'sunburst'), result = sunburstNodes(data);
  assert.equal(result.total, 30);
  assert.equal(result.nodes.find((n) => n.name === '2025').own, 10);
  assert.equal(result.nodes.find((n) => n.name === 'A').value, 20);
  assert.deepEqual(rows, before);
  const svg = renderChartSvg({ type: 'sunburst', w: 500, h: 360 }, data);
  assert.match(svg, /data-depth="2"/); assert.doesNotMatch(svg, /NaN|Infinity/);
});

test('원형 섹터: 단일 항목의 완전한 원과 링은 두 호로 닫힘', () => {
  const circle = annularSector(50, 50, 0, 40, 0, Math.PI * 2), ring = annularSector(50, 50, 20, 40, 0, Math.PI * 2);
  assert.equal((circle.match(/A/g) ?? []).length, 2);
  assert.equal((ring.match(/A/g) ?? []).length, 4);
  assert.equal(annularSector(0, 0, 1, 2, 0, 0), '');
});

test('표면 데이터: 숫자 축 머리글과 헤더 없는 행렬, 행/열 전환을 유지', () => {
  const data = chartData(surfaceRows, 'surface');
  assert.deepEqual(data.categories, ['0', '10', '20']);
  assert.deepEqual(data.series.map((s) => s.name), ['0', '10', '20']);
  assert.deepEqual(data.series[1].values, [10, 20, 30]);
  assert.deepEqual(chartData([[1, 2], [3, 4]], 'surface').series.map((s) => s.values), [[1, 2], [3, 4]]);
  assert.deepEqual(chartData([['', 'a', 'b'], ['r1', 1, 2], ['r2', 3, 4]], 'surface', true).series.map((s) => s.values), [[1, 3], [2, 4]]);
});

test('등고선: 평면의 등고선 위치를 값으로 보간하고 색대 면적 합계를 보존', () => {
  const data = { categories: ['0', '1'], series: [{ values: [0, 10] }, { values: [0, 10] }] };
  const g = surfaceGeometry(data, { bandCount: 5 });
  near(g.faces.reduce((sum, face) => sum + area(face.points), 0), 1);
  assert.deepEqual([...new Set(g.lines.map((l) => l.value))], [2, 4, 6, 8]);
  for (const line of g.lines) for (const point of line.points) { near(point.x, line.value / 10); near(point.z, line.value); }
  for (const face of g.faces) for (const point of face.points) assert.ok(point.z >= g.levels[face.band] - 1e-9 && point.z <= g.levels[face.band + 1] + 1e-9);
});

test('표면: 상수 면을 중복 채우지 않고 빈 격자를 연결하지 않음', () => {
  const data = { categories: ['a', 'b'], series: [{ values: [5, 5] }, { values: [5, 5] }] };
  const g = surfaceGeometry(data);
  assert.equal(g.faces.length, 2); near(g.faces.reduce((sum, face) => sum + area(face.points), 0), 1); assert.equal(g.lines.length, 0);
  const hole = surfaceGeometry({ categories: ['a', 'b'], series: [{ values: [1, null] }, { values: [2, 3] }] });
  assert.equal(hole.faces.length, 0); assert.equal(hole.lines.length, 0);
});

test('표면 네 변형은 채움·메시·등고선 및 회전이 실제로 다름', () => {
  const data = chartData(surfaceRows, 'surface'), rendered = {};
  for (const style of ['surface', 'wireframe', 'contour', 'wireframeContour']) {
    const svg = renderChartSvg({ type: 'surface', surfaceStyle: style, w: 540, h: 360 }, data), xml = parseXml(svg);
    assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
    assert.equal(descendants(xml, 'polygon').length > 0, !/wireframe/i.test(style));
    if (/contour/i.test(style) || /wireframe/i.test(style)) assert.match(svg, new RegExp(`data-surface-line="${/contour/i.test(style) ? 'contour' : 'mesh'}"`));
    else assert.doesNotMatch(svg, /data-surface-line="mesh"/, '불투명 표면 위로 뒤쪽 골격을 덧그리지 않음');
    rendered[style] = svg;
  }
  assert.notEqual(rendered.surface, renderChartSvg({ type: 'surface', surfaceStyle: 'surface', w: 540, h: 360, view3D: { rotY: 135 } }, data));
  assert.doesNotMatch(rendered.contour, /data-3d=/);
});

test('대조 원/막대 분할: 위치·값·백분율·선택 번호와 합계를 보존', () => {
  const values = [60, 20, 10, 5, 5], before = [...values];
  for (const chart of [{ splitType: 'position', splitPos: 2 }, { splitType: 'value', splitPos: 10 }, { splitType: 'percent', splitPos: 10 }, { splitType: 'custom', splitPoints: [3, 4] }]) {
    const s = splitPieData(values, chart);
    assert.deepEqual(s.secondary.map((p) => p.i), [3, 4]); assert.equal(s.other, 10); assert.equal(s.total, 100);
    assert.equal(s.primary.reduce((sum, p) => sum + p.value, 0), 100);
    assert.equal(s.primary.at(-1).i, -1);
  }
  assert.deepEqual(values, before);
  assert.equal(splitPieData([10], {}).secondary.length, 0);
});

test('대조 원형과 대조 막대는 기타 합계·원본 점 번호·보조 표현을 다르게 렌더', () => {
  const data = { categories: ['A', 'B', 'C', 'D'], series: [{ name: '금액', values: [60, 20, 15, 5] }] };
  for (const type of ['pieOfPie', 'barOfPie']) {
    const svg = renderChartSvg({ type, splitPos: 2, w: 500, h: 320 }, data), xml = parseXml(svg);
    const secondary = descendants(xml, type === 'pieOfPie' ? 'path' : 'rect').filter((n) => n.attrs['data-ofpie'] === 'secondary');
    assert.equal(secondary.length, 2); assert.deepEqual(secondary.map((n) => n.attrs['data-p']), ['2', '3']);
    assert.deepEqual(secondary.map((n) => Number(n.attrs['data-value'])), [15, 5]);
    assert.equal(descendants(xml, 'line').filter((n) => n.attrs['data-ofpie'] === 'connector').length, 2);
    assert.doesNotMatch(svg, /NaN|Infinity/);
  }
});

test('거래량 주식: 4/5계열의 열 역할을 가격과 분리하고 두 축으로 그리기', () => {
  for (const ohlc of [false, true]) {
    const names = ohlc ? ['거래량', '시가', '고가', '저가', '종가'] : ['거래량', '고가', '저가', '종가'];
    const values = ohlc ? [[1000000, 2000000], [100, 110], [120, 140], [90, 95], [110, 105]] : [[1000000, 2000000], [120, 140], [90, 95], [110, 105]];
    const data = { categories: ['월', '화'], series: names.map((name, i) => ({ name, values: values[i] })) };
    const stock = stockVolumeData(data.series, { volume: true, ohlc });
    assert.equal(stock.high.name, '고가'); assert.equal(stock.close.name, '종가'); assert.equal(stock.volume.name, '거래량');
    const svg = renderChartSvg({ type: 'stock', volume: true, ohlc, w: 540, h: 360 }, data), xml = parseXml(svg);
    assert.equal(descendants(xml, 'rect').filter((n) => n.attrs['data-stock'] === 'volume').length, 2);
    assert.equal(descendants(xml, 'line').filter((n) => n.attrs['data-stock'] === 'highLow').length, 2);
    assert.equal(descendants(xml, 'rect').filter((n) => n.attrs['data-stock'] === 'openClose').length, ohlc ? 2 : 0);
    for (const n of descendants(xml, 'line').filter((n) => n.attrs['data-stock'] === 'highLow')) assert.ok(Math.abs(Number(n.attrs.y1) - Number(n.attrs.y2)) > 30, '거래량이 가격 축 범위를 지배하지 않음');
    assert.doesNotMatch(svg, /NaN|Infinity/);
  }
});

test('참조 없는 계열의 항목 캐시와 계층 캐시도 유지', () => {
  const chart = { type: 'sunburst', series: [{ name: { text: '값' }, cache: [2, 3], catCache: ['A', 'B'], catLevels: [[{ text: '상위', start: 0, end: 1 }]] }] };
  const data = resolveChart(chart, {});
  assert.deepEqual(data.categories, ['A', 'B']); assert.deepEqual(data.catLevels, chart.series[0].catLevels);
  assert.equal(sunburstNodes(data).nodes.find((n) => n.name === '상위').value, 5);
});

test('3D 거품은 X/Y/면적 크기를 유지하고 0·음수 크기 기본값을 구분', () => {
  const data = { categories: ['a', 'b', 'c', 'd'], series: [{ name: '거품', x: [1, 2, 3, 4], values: [1, 2, 3, 4], size: [1, 4, 0, -9] }] };
  const circles = (threeD) => descendants(parseXml(renderChartSvg({ type: 'bubble', threeD, w: 500, h: 340 }, data)), 'circle');
  const plain = circles(false), sphere = circles(true);
  assert.equal(plain.length, 2); assert.equal(sphere.length, 2);
  for (let i = 0; i < 2; i++) for (const k of ['cx', 'cy', 'r']) assert.equal(plain[i].attrs[k], sphere[i].attrs[k]);
  for (const n of sphere) for (const [key, max] of [['cx', 500], ['cy', 340]]) { const p = Number(n.attrs[key]), r = Number(n.attrs.r); assert.ok(p - r >= 0 && p + r <= max, '축 끝 거품이 차트 가장자리에서 잘리지 않음'); }
  near(Number(sphere[1].attrs.r) / Number(sphere[0].attrs.r), 2, .01);
  assert.ok(sphere.every((n) => n.attrs['data-3d'] === 'bubble' && n.attrs.fill.startsWith('url(')));
});

test('주식형 갤러리는 OHLC에 부족한 계열로도 미리보기 오류가 없음', () => {
  const data = { categories: ['월'], series: [{ name: '고가', values: [5] }, { name: '저가', values: [2] }, { name: '종가', values: [3] }] };
  assert.doesNotThrow(() => renderChartSvg({ type: 'stock', ohlc: true, w: 240, h: 160 }, data));
  assert.doesNotThrow(() => renderChartSvg({ type: 'stock', ohlc: false, w: 240, h: 160 }, { ...data, series: data.series.slice(0, 2) }));
});

test('트리맵 반복 상위 항목은 하나의 묶음과 범례로 합쳐짐', () => {
  const data = chartData([['지역', '나라', '값'], ['아시아', '한국', 10], ['미주', '미국', 7], ['아시아', '일본', 5]], 'treemap');
  const texts = descendants(parseXml(renderChartSvg({ type: 'treemap', w: 640, h: 420 }, data)), 'text');
  assert.equal(texts.filter((t) => t.text === '아시아').length, 2, '묶음 제목 1 + 범례 1');
  assert.equal(texts.filter((t) => t.text === '미주').length, 2);
  for (const c of data.categories) assert.equal(texts.filter((t) => t.text === c).length, 1);
});

test('고급 갤러리는 구현된 유형과 최소 9개의 추가 변형을 노출', () => {
  const variants = CHART_GALLERY.flatMap(([, entries]) => entries).map(([, p]) => p);
  for (const type of ['sunburst', 'surface', 'pieOfPie', 'barOfPie']) assert.ok(CHART_TYPES.some((t) => t.id === type));
  assert.equal(variants.filter((p) => p.type === 'surface').length, 4);
  assert.equal(variants.filter((p) => p.type === 'stock' && p.volume).length, 2);
  assert.ok(variants.some((p) => p.type === 'sunburst'));
  assert.ok(variants.some((p) => p.type === 'pieOfPie'));
  assert.ok(variants.some((p) => p.type === 'barOfPie'));
});
