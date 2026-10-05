import test from 'node:test';
import assert from 'node:assert/strict';
import { chartTextStyle, chartTextFormatPatch, chartTextTargetSupported, chartTextFontNames } from '../src/chart-text-format.js';
import { renderChartSvg, resolveChart } from '../src/chart.js';
import { parseXml, descendants } from '../src/xml.js';
import { Workbook } from '../src/workbook.js';

const chart = (extra = {}) => ({ id: 'text-chart', type: 'column', w: 700, h: 460, title: '월간 실적', legend: 'b', labels: true, axes: { x: { title: '기간' }, y: { title: '매출' } }, series: [{ name: { text: '매출' }, catCache: ['첫째', '둘째', '셋째'], cache: [12, 24, 36] }], ...extra });
const draw = c => parseXml(renderChartSvg(c, resolveChart(c, {})));
const apply = (c, part, delta) => ({ ...c, ...chartTextFormatPatch(c, part, delta) });
const texts = root => descendants(root, 'text');
const partTexts = (root, kind) => {
  const group = descendants(root, 'g').find(n => n.attrs['data-el'] === kind);
  return group ? texts(group) : texts(root).filter(n => n.attrs['data-el'] === kind);
};
const checkText = (node, size, color, family) => {
  assert.ok(node, '렌더에 선택 대상 텍스트가 있어야 함');
  if (size !== undefined) assert.equal(Number(node.attrs['font-size']), Math.round(size * 4 / 3 * 10) / 10);
  if (color !== undefined) assert.equal(node.attrs.fill, color);
  if (family !== undefined) assert.ok(node.attrs['font-family']?.includes(family), node.attrs['font-family']);
};

test('제목 글꼴·크기·색·강조 변경은 실제 SVG에만 해당 요소로 적용되고 모델은 불변이다', () => {
  const c = chart({ titleSize: 14, legendColor: '#113355' }), before = structuredClone(c);
  const changed = apply(c, { kind: 'title' }, { font: 'Noto Serif KR', size: 22, color: '#cc2244', bold: true, italic: true, underline: true });
  const svg = draw(changed), title = partTexts(svg, 'title')[0];
  checkText(title, 22, '#cc2244', 'Noto Serif KR');
  assert.equal(title.attrs['font-weight'], '700'); assert.equal(title.attrs['font-style'], 'italic'); assert.equal(title.attrs['text-decoration'], 'underline');
  assert.equal(title.text, '월간 실적'); checkText(partTexts(svg, 'legend')[0], 9, '#113355');
  assert.deepEqual(c, before); assert.deepEqual(changed.series, before.series); assert.deepEqual(changed.axes, before.axes);
});

test('기본·보조 축 눈금과 각각의 축 제목은 독립된 글꼴 대상으로 렌더된다', () => {
  let c = chart({ type: 'combo', series: [...chart().series, { name: { text: '비율' }, catCache: ['첫째', '둘째', '셋째'], cache: [.1, .2, .3] }], seriesFmt: [{ type: 'column' }, { type: 'line', axis: 'secondary' }], axes: { x: { title: '기간' }, y: { title: '매출' }, y2: { title: '비율' } } });
  const parts = ['axis-x', 'axis-y', 'axis-y2', 'axis-title-x', 'axis-title-y', 'axis-title-y2'];
  parts.forEach((kind, i) => { c = apply(c, { kind }, { font: 'Georgia', size: 10 + i, color: '#11223' + i, italic: true }); });
  const svg = draw(c);
  parts.forEach((kind, i) => {
    const nodes = partTexts(svg, kind).filter(n => !n.attrs['data-el']?.startsWith('axis-title-') || n.attrs['data-el'] === kind);
    assert.ok(nodes.length, kind); nodes.forEach(n => { checkText(n, 10 + i, '#11223' + i, 'Georgia'); assert.equal(n.attrs['font-style'], 'italic'); });
  });
  checkText(partTexts(svg, 'title')[0], 14); checkText(partTexts(svg, 'legend')[0], 9);
});

test('범례의 글자색은 계열 범례 표지의 색과 다른 속성이다', () => {
  const c = chart({ seriesFmt: [{ color: '#008866' }] });
  const next = apply(c, { kind: 'legend' }, { font: '나눔명조', size: 16, color: '#ffffff', bold: true });
  const svg = draw(next), group = descendants(svg, 'g').find(n => n.attrs['data-el'] === 'legend');
  checkText(texts(group)[0], 16, '#ffffff', 'NanumMyeongjo');
  assert.equal(descendants(group, 'rect')[0].attrs.fill, '#008866');
  assert.equal(next.seriesFmt[0].color, '#008866');
});

test('데이터 레이블의 개별 항목 서식은 필터 뒤에도 원래 항목 번호에 연결된다', () => {
  let c = chart({ hiddenCats: [0], seriesFmt: [{ color: '#008866', labelColor: '#112233', labelSize: 11 }] });
  c = apply(c, { kind: 'label', s: 0, p: 2 }, { font: 'Georgia', size: 18, color: '#992244', bold: true, italic: true });
  const labels = partTexts(draw(c), 'label');
  assert.deepEqual(labels.map(n => n.attrs['data-p']), ['1', '2']);
  checkText(labels[0], 11, '#112233'); checkText(labels[1], 18, '#992244', 'Georgia');
  assert.equal(c.seriesFmt[0].labelColor, '#112233');
  assert.deepEqual(c.seriesFmt[0].pointLabelStyles[2], { font: 'Georgia', size: 18, color: '#992244', bold: true, italic: true });
  c = apply(c, { kind: 'label', s: 0 }, { color: '#abcdef', italic: false });
  const after = partTexts(draw(c), 'label'); after.forEach(n => { assert.equal(n.attrs.fill, '#abcdef'); assert.equal(n.attrs['font-style'], 'normal'); });
  checkText(after[1], 18, '#abcdef', 'Georgia');
});

test('차트 전체 글꼴은 하위 요소의 기존 오버라이드를 교체하되 값·배치·선색은 보존한다', () => {
  const c = chart({ font: 'Arial', titleFont: 'Georgia', legendBold: false, titleSize: 22, axes: { x: { title: '기간', font: 'Calibri', color: '#123456' }, y: { title: '매출', titleSize: 24 } }, seriesFmt: [{ labelFont: 'Arial', labelSize: 11, color: '#006633', pointLabelStyles: { 1: { font: 'Verdana', size: 12, color: '#aa0000' } } }] });
  const before = structuredClone(c), changed = apply(c, null, { font: 'Noto Sans KR', size: 13, color: '#664422', bold: true, italic: true });
  for (const n of texts(draw(changed))) { checkText(n, 13, '#664422', 'Noto Sans KR'); assert.equal(n.attrs['font-weight'], '700'); assert.equal(n.attrs['font-style'], 'italic'); }
  assert.deepEqual(c, before); assert.deepEqual(changed.series, c.series); assert.equal(changed.w, c.w); assert.equal(changed.seriesFmt[0].color, '#006633');
  assert.equal(changed.seriesFmt[0].cache, undefined, '큰 원본 데이터 배열을 서식에 복제하지 않음');
});

for (const type of ['column', 'bar', 'line', 'area', 'pie', 'doughnut', 'scatter', 'bubble', 'radar', 'waterfall', 'histogram', 'funnel', 'treemap', 'sunburst']) test(`${type}: 데이터 레이블 글꼴과 기울임은 실제 SVG에 적용된다`, () => {
  const c = apply(chart({ type, axes: undefined }), { kind: 'label', s: 0 }, { font: 'Georgia', size: 10, color: '#663399', italic: true });
  const labels = partTexts(draw(c), 'label');
  assert.ok(labels.length, type); for (const node of labels) { checkText(node, 10, '#663399', 'Georgia'); assert.equal(node.attrs['font-style'], 'italic'); }
});

test('데이터 표의 텍스트는 축이나 범례를 바꾸지 않고 따로 서식을 갖는다', () => {
  const c = apply(chart({ dataTable: true }), { kind: 'dataTable' }, { font: 'Georgia', size: 14, color: '#224466', strike: true });
  const svg = draw(c), rows = partTexts(svg, 'dataTable');
  assert.ok(rows.length > 3); rows.forEach(n => { checkText(n, 14, '#224466', 'Georgia'); assert.equal(n.attrs['text-decoration'], 'line-through'); });
  checkText(partTexts(svg, 'legend')[0], 9);
});

test('차트 텍스트 변경 실행 취소·재실행은 셀 데이터와 독립적으로 보존된다', () => {
  const wb = new Workbook(), c = chart(); wb.setInput(0, 0, 0, '셀 데이터'); wb.sheets[0].charts = [c];
  const changed = apply(c, { kind: 'title' }, { font: 'Georgia', size: 25, color: '#cc3300' });
  wb.transact(() => wb.setSheetProp(0, 'charts', [changed]));
  checkText(partTexts(draw(wb.sheets[0].charts[0]), 'title')[0], 25, '#cc3300', 'Georgia');
  wb.undo(); assert.deepEqual(wb.sheets[0].charts[0], c); assert.equal(wb.getValue(0, 0, 0), '셀 데이터');
  wb.redo(); checkText(partTexts(draw(wb.sheets[0].charts[0]), 'title')[0], 25, '#cc3300', 'Georgia'); assert.equal(wb.getValue(0, 0, 0), '셀 데이터');
});

test('서식 입력은 안전하게 이스케이프하고 잘못된 크기나 비텍스트 대상은 무시한다', () => {
  const c = chart();
  for (const kind of ['series', 'point', 'node', 'plot']) { assert.equal(chartTextTargetSupported(c, { kind }), false); assert.equal(chartTextStyle(c, { kind }), null); assert.equal(chartTextFormatPatch(c, { kind }, { font: 'Arial' }), null); }
  assert.equal(chartTextTargetSupported(c, { kind: 'title', id: 'different' }), false);
  for (const size of [0, -1, Infinity, NaN, 10000]) assert.deepEqual(chartTextFormatPatch(c, { kind: 'title' }, { size }), {});
  const svg = draw(apply(c, { kind: 'title' }, { font: '"><script>bad</script>', color: '"><script>bad</script>' }));
  assert.equal(descendants(svg, 'script').length, 0); assert.equal(partTexts(svg, 'title')[0].text, c.title);
});

test('원본 계열·포인트 글꼴도 수집하고 주석·캐시 데이터는 서식 패치에 복사하지 않는다', () => {
  const c = chart({ font: 'Arial', series: [{ ...chart().series[0], labelFont: 'Georgia', pointLabelStyles: { 0: { font: 'Verdana', size: 15 } } }], seriesFmt: [{ labelColor: '#aa3344' }], axes: { y: { font: 'Noto Sans KR', titleFont: 'Noto Serif KR' } } });
  assert.deepEqual(new Set(chartTextFontNames(c)), new Set(['Arial', 'Georgia', 'Verdana', 'Noto Sans KR', 'Noto Serif KR']));
  assert.equal(chartTextStyle(c, { kind: 'label', s: 0 }).font, 'Georgia'); assert.equal(chartTextStyle(c, { kind: 'label', s: 0, p: 0 }).size, 15);
  const next = apply(c, { kind: 'label', s: 0 }, { italic: true });
  assert.equal(next.seriesFmt[0].cache, undefined); assert.equal(next.seriesFmt[0].pointLabelStyles[0].font, 'Verdana');
});
