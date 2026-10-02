import { test } from 'node:test';
import assert from 'node:assert/strict';
import { hierarchyCategories, buildChartHierarchy, hierarchyRects, hierarchyLegend, hierarchyNodeColor, fitHierarchyLabel } from '../src/chart-hierarchy.js';
import { sunburstNodes } from '../src/chart-advanced.js';
import { chartData, resolveChart, filterChart, renderChartSvg } from '../src/chart.js';
import { parseXml, descendants } from '../src/xml.js';

const rows = [['지역', '나라', '상품', '값'], ['아시아', '한국', '광고', 30], ['', '', '분석', 20], ['미주', '미국', '광고', 50]];
const near = (actual, expected, eps = 1e-8) => assert.ok(Math.abs(actual - expected) <= eps, `${actual} ≈ ${expected}`);
const nodesOf = (svg, tag = 'path') => descendants(parseXml(svg), tag).filter((n) => n.attrs['data-node']);
const draw = (type, fmt = {}, options = {}, input = chartData(rows, type)) => {
  const chart = { type, w: 800, h: 600, seriesFmt: [fmt], ...options };
  const data = { ...input, series: input.series.map((s) => ({ ...s, ...fmt, _fi: 0 })) };
  return renderChartSvg(chart, data);
};

test('계층의 sparse 부모·명시된 반복 부모·짧은 가지를 서로 구분한다', () => {
  const input = [['A', '중간', '첫째'], ['A', '', '둘째'], ['', '', '셋째'], ['B', '', '넷째'], ['B', '', '']];
  const before = structuredClone(input), data = hierarchyCategories(input);
  assert.deepEqual(data.paths, [['A', '중간', '첫째'], ['A', '중간', '둘째'], ['A', '중간', '셋째'], ['B', '', '넷째'], ['B', '', '']]);
  assert.deepEqual(input, before);
  const h = buildChartHierarchy({ categories: data.categories, hierarchyPaths: data.paths, series: [{ values: [1, 2, 3, 4, 5] }] });
  assert.equal(h.nodes.find((n) => n.key === '["B"]').own, 5);
  assert.equal(h.nodes.find((n) => n.key === '["B","","넷째"]').depth, 3, '중간 빈 단계를 지워 다른 가지와 합치지 않는다');
});

test('단 한 행의 여러 계층 및 숫자 계층도 다단계로 유지한다', () => {
  const data = chartData([['연도', '월', '값'], [2026, 10, 7]], 'sunburst');
  assert.deepEqual(buildChartHierarchy(data).nodes.map((n) => [n.name, n.depth, n.value]), [['2026', 1, 7], ['10', 2, 7]]);
  const explicit = resolveChart({ type: 'treemap', series: [{ cat: 'cats', val: 'vals' }] }, { values: (r) => r === 'cats' ? [[2026, 10]] : [[7]] });
  assert.equal(buildChartHierarchy(explicit).depth, 2);
});

test('범위 트리맵의 숫자 상위 계층/마지막 값 열을 재추론으로 훼손하지 않는다', () => {
  const source = [['연도', '제품', '값'], [2025, '가', 4], [2026, '나', 6]];
  const data = resolveChart({ type: 'treemap' }, { range: () => source });
  assert.deepEqual(data.series[0].values, [4, 6]);
  assert.deepEqual(buildChartHierarchy(data).roots.map((n) => n.name), ['2025', '2026']);
});

test('항목 필터는 계층·원본 번호·같은 부모 색을 보존한다', () => {
  const base = chartData(rows, 'sunburst'), before = structuredClone(base);
  const filtered = filterChart(base, { hiddenCats: [0] }), h = sunburstNodes(filtered);
  assert.equal(h.depth, 3); assert.equal(h.total, 70);
  assert.deepEqual(h.roots.map((n) => [n.name, n.value]), [['아시아', 20], ['미주', 50]]);
  assert.deepEqual(h.nodes.find((n) => n.name === '분석').points, [1]);
  assert.deepEqual(base, before);
  const onlyAmerica = filterChart(base, { hiddenCats: [0, 1] });
  assert.equal(buildChartHierarchy(onlyAmerica).roots[0].rootIndex, 1);
  assert.equal(hierarchyLegend(onlyAmerica, {}, ['red', 'blue'])[0].color, 'blue');
});

test('같은 경로의 반복 값만 합산하고 부모의 범위/자식 각도 합계를 보존한다', () => {
  const data = chartData([['부모', '항목', '값'], ['A', '같음', 2], ['A', '같음', 3], ['B', '같음', 5]], 'sunburst');
  const h = sunburstNodes(data), a = h.nodes.find((n) => n.key === '["A","같음"]');
  assert.equal(a.value, 5); assert.deepEqual(a.points, [0, 1]);
  assert.equal(h.nodes.filter((n) => n.name === '같음').length, 2);
  for (const parent of h.roots) {
    near(parent.end - parent.start, Math.PI);
    for (const child of parent.children) { assert.ok(child.start >= parent.start); assert.ok(child.end <= parent.end + 1e-9); }
  }
});

test('0/음수/빈 값은 합계에서 제외하고 이유/상한을 명시한다', () => {
  const data = { categories: ['a', 'b', 'c', 'd', 'e'], series: [{ values: [5, 0, -4, null, NaN] }] };
  const h = buildChartHierarchy(data);
  assert.equal(h.total, 5); assert.deepEqual(h.ignored, { zero: 1, negative: 1, missing: 2 });
  const tooMany = { categories: [], series: [{ values: Array(30001).fill(1) }] };
  assert.match(buildChartHierarchy(tooMany).error, /3만/);
  assert.match(renderChartSvg({ type: 'sunburst' }, tooMany), /3만/);
});

test('선버스트 단일 계층도 중심 구멍을 유지하고 회전/완전한 링을 그린다', () => {
  const data = { categories: ['하나'], series: [{ values: [5] }] };
  const svg = draw('sunburst', {}, {}, data), path = nodesOf(svg)[0];
  assert.equal((path.attrs.d.match(/A/g) ?? []).length, 4, '전체 원은 외곽2호+안쪽2호');
  assert.notEqual(nodesOf(draw('sunburst', {}, { firstAngle: 60 }, data))[0].attrs.d, path.attrs.d);
  assert.doesNotMatch(svg, /NaN|Infinity|undefined/);
});

for (const type of ['sunburst', 'treemap']) {
  test(`${type}: 원본 leaf 선택과 부모 노드 선택은 다른 키이며 leaf 색이 부모로 번지지 않는다`, () => {
    const data = filterChart(chartData(rows, type), { hiddenCats: [0] });
    const svg = draw(type, { pointColors: { 1: '#123456' } }, {}, data), elements = nodesOf(svg, type === 'sunburst' ? 'path' : 'rect');
    const leaf = elements.find((n) => n.attrs['data-node'] === '["아시아","한국","분석"]');
    assert.equal(leaf.attrs['data-p'], '1'); assert.equal(leaf.attrs.fill, '#123456');
    const parent = elements.find((n) => n.attrs['data-node'] === '["아시아"]');
    assert.equal(parent.attrs['data-p'], undefined); assert.notEqual(parent.attrs.fill, '#123456');
    assert.ok(elements.every((n) => n.attrs['data-s'] === '0'));
  });
  test(`${type}: 계열/부모/하위 직접색/원본 점색 우선순위와 레이블 옵션이 실제 SVG에 반영된다`, () => {
    const h = buildChartHierarchy(chartData(rows, type)), parent = h.roots[0], leaf = h.nodes.find((n) => n.name === '분석');
    const fmt = { hierarchyColors: { [parent.key]: '#abcdef' }, pointColors: { 1: '#123456' }, catName: false, labels: true, serName: true, numFmt: '0.0', labelColor: '#13579b', labelSize: 7, labelBold: true };
    assert.equal(hierarchyNodeColor(parent, fmt, {}, ['red']), '#abcdef');
    assert.equal(hierarchyNodeColor(leaf, fmt, {}, ['red']), '#123456');
    const svg = draw(type, fmt, { treemapLabelLayout: 'none' });
    const texts = descendants(parseXml(svg), 'text').filter((n) => n.attrs['data-el'] === 'label');
    assert.ok(texts.length); assert.ok(texts.every((n) => n.attrs.fill === '#13579b' && n.attrs['font-weight'] === '700'));
    assert.ok(texts.some((n) => descendants(n, 'title').some((t) => /값 · \d+\.0/.test(t.text))));
    assert.equal(descendants(parseXml(draw(type, { labels: false, catName: false }, { treemapLabelLayout: 'none' })), 'text').filter((n) => n.attrs['data-el'] === 'label').length, 0);
  });
}

test('트리맵 부모 레이블 banner/overlapping/none은 실제 배치가 다르고 leaf 면적은 양수이다', () => {
  const shapes = {};
  for (const mode of ['banner', 'overlapping', 'none']) {
    const svg = draw('treemap', {}, { treemapLabelLayout: mode });
    const texts = descendants(parseXml(svg), 'text').filter((n) => n.attrs['data-el'] === 'label' && n.attrs['data-depth'] === '1');
    assert.equal(texts.length, mode === 'none' ? 0 : 2);
    const rects = nodesOf(svg, 'rect').filter((n) => n.attrs['data-p']);
    assert.ok(rects.every((r) => Number(r.attrs.width) > 0 && Number(r.attrs.height) > 0));
    shapes[mode] = rects[0].attrs;
  }
  assert.ok(Number(shapes.banner.y) > Number(shapes.overlapping.y));
  assert.equal(shapes.overlapping.y, shapes.none.y);
});

test('트리맵 배치: 면적비·경계·서로 겹치지 않음·원본 불변을 독립 계산으로 검증', () => {
  const items = Array.from({ length: 29 }, (_, i) => ({ value: ((i * 17) % 31) + 1 })), before = structuredClone(items), box = { x: 13, y: 29, w: 701, h: 433 };
  const rects = hierarchyRects(items, box), total = items.reduce((s, n) => s + n.value, 0);
  assert.equal(rects.length, items.length);
  near(rects.reduce((s, r) => s + r.w * r.h, 0), box.w * box.h, 1e-6);
  for (const r of rects) { near(r.w * r.h / (box.w * box.h), r.node.value / total); assert.ok(r.x >= box.x && r.y >= box.y); assert.ok(r.x + r.w <= box.x + box.w + 1e-7 && r.y + r.h <= box.y + box.h + 1e-7); }
  for (let i = 0; i < rects.length; i++) for (let j = i + 1; j < rects.length; j++) { const a = rects[i], b = rects[j]; assert.ok(Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x) < 1e-7 || Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y) < 1e-7); }
  assert.deepEqual(items, before);
});

test('계층명/색의 XML 특수문자는 속성/툴팁으로 탈출하지 않는다', () => {
  const data = chartData([['부모', '자식', '값'], ['A"<&', 'B', 3]], 'sunburst');
  const svg = draw('sunburst', { pointColors: { 0: 'red" onload="bad' } }, {}, data);
  const paths = nodesOf(svg);
  assert.equal(paths[1].attrs.onload, undefined);
  assert.equal(paths[0].attrs['data-node'], '["A\\\"<&"]');
});

test('긴 한글 레이블은 공간에 맞게 줄이되 전체 문구는 툴팁에 유지한다', () => {
  const name = '대한민국 온라인 광고 월별 분석';
  assert.equal(fitHierarchyLabel(name, 48, 12), '대한민…');
  assert.equal(fitHierarchyLabel('A', 30, 12), 'A');
  const data = { categories: [name], series: [{ values: [1] }] };
  const labels = descendants(parseXml(draw('sunburst', {}, {}, data)), 'text').filter((n) => n.attrs['data-el'] === 'label');
  assert.equal(descendants(labels[0], 'title')[0].text, name);
  assert.ok(labels[0].text.endsWith('…'));
});
