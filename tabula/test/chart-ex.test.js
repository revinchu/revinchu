import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { chartModelData, chartData } from '../src/chart.js';
import { isChartEx, writeChartEx, readChartEx, CHARTEX_NS, CHARTEX_CONTENT, CHARTEX_REL } from '../src/chart-ex.js';

const rows = [['분류', '금액', '이익'], ['한국', 10, 2], ['일본', 5, 4], ['미국', 7, 6]];
const hierarchy = [['지역', '상세', '금액'], ['아시아', '한국', 10], ['아시아', '일본', 5], ['미주', '미국', 7]];
function book(type, options = {}, cells = rows) {
  const wb = new Workbook();
  wb.transact(() => cells.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  wb.sheets[0].charts = [{ id: 'cx1', type, title: '검증 차트', range: { r1: 0, c1: 0, r2: cells.length - 1, c2: cells[0].length - 1 }, x: 0, y: 0, w: 600, h: 360, ...options }];
  return wb;
}
const path = 'xl/charts/chart1.xml';
function noExtension(files) {
  const xml = textOf(files[path]).replace(/<cx:extLst>[\s\S]*?<\/cx:extLst>/g, '');
  return { ...files, [path]: xml };
}
function reopen(files) {
  const wb = new Workbook(readXlsx(zip(files)).data);
  return { wb, chart: wb.sheets[0].charts[0] };
}

for (const type of ['waterfall', 'funnel', 'histogram', 'pareto', 'treemap', 'sunburst', 'boxWhisker', 'map']) {
  test(`ChartEx ${type}: 표준 종류·참조·캐시를 확장 정보 없이 읽고 XLSX 재저장`, () => {
    const cells = ['treemap', 'sunburst'].includes(type) ? hierarchy : type === 'boxWhisker' ? rows : rows.map((r) => r.slice(0, 2));
    const wb = book(type, {}, cells), before = chartModelData(wb, 0, wb.sheets[0].charts[0]);
    const files = unzip(writeXlsx(wb));
    const xml = textOf(files[path]), native = parseXml(xml);
    assert.ok(xml.includes(CHARTEX_NS));
    assert.ok(textOf(files['[Content_Types].xml']).includes(CHARTEX_CONTENT));
    assert.ok(textOf(files['xl/drawings/_rels/drawing1.xml.rels']).includes(CHARTEX_REL));
    assert.ok(textOf(files['xl/drawings/drawing1.xml']).includes(CHARTEX_NS));
    assert.equal(descendants(native, 'barChart').length, 0, '막대 대체 차트 아님');
    const { wb: loaded, chart } = reopen(noExtension(files));
    assert.ok(chart, '표준 확장 차트가 도형에서 발견됨'); assert.equal(chart.type, type);
    const after = chartModelData(loaded, 0, chart);
    assert.deepEqual(after.series.map((s) => s.values), before.series.map((s) => s.values));
    assert.deepEqual(after.series.map((s) => s.name), before.series.map((s) => s.name));
    if (type !== 'histogram') assert.deepEqual(after.categories, before.categories);
    if (['sunburst', 'treemap'].includes(type)) assert.deepEqual(after.catLevels, before.catLevels);
    assert.equal(reopen(unzip(writeXlsx(loaded))).chart.type, type);
    const literal = readChartEx(native);
    assert.deepEqual(literal.series.map((s) => s.cache), before.series.map((s) => s.values));
    if (process.env.CHARTEX_FIXTURE_DIR) {
      mkdirSync(process.env.CHARTEX_FIXTURE_DIR, { recursive: true });
      writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, `${type}.xml`), xml);
      writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, `${type}.xlsx`), writeXlsx(wb));
    }
  });
}

test('ChartEx 표준 옵션: 폭포 합계·연결선, 구간 크기/개수, 상자 평균·사분위', () => {
  const options = [
    ['waterfall', { totals: [0, 2], connectors: false }],
    ['histogram', { binWidth: 2 }], ['histogram', { binCount: 4 }],
    ['boxWhisker', { showMean: true, showOutliers: false, showInnerPoints: true, quartileMethod: 'exclusive' }],
  ];
  for (const [type, props] of options) {
    const files = noExtension(unzip(writeXlsx(book(type, props))));
    const chart = reopen(files).chart;
    for (const [key, value] of Object.entries(props)) assert.deepEqual(chart[key], value);
    if (process.env.CHARTEX_FIXTURE_DIR) writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, `${type}-${Object.keys(props)[0]}.xml`), typeof files[path] === 'string' ? files[path] : textOf(files[path]));
  }
});

test('ChartEx 표준 서식: 제목·범례·배경·계열·점 색과 축 설정 보존', () => {
  const options = { titleSize: 18, titleBold: true, titleColor: '#112233', legend: 'r', legendSize: 9, legendBold: true, legendColor: '#334455', fill: '#fafafa', plotFill: '#efefef', border: '#778899', labels: true, seriesFmt: [{ color: '#aa0000', pointColors: { 1: '#00aa00' }, numFmt: '0.0' }], axes: { x: { hide: true, title: '구간' }, y: { min: -5, max: 20, major: 5, title: '금액', numFmt: '0.00' } } };
  const files = noExtension(unzip(writeXlsx(book('waterfall', options))));
  const chart = reopen(files).chart;
  for (const key of ['titleSize', 'titleBold', 'titleColor', 'legend', 'legendSize', 'legendBold', 'legendColor', 'fill', 'plotFill', 'border', 'axes']) assert.deepEqual(chart[key], options[key]);
  assert.deepEqual(chart.seriesFmt[0], { ...options.seriesFmt[0], labels: true });
  if (process.env.CHARTEX_FIXTURE_DIR) writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, 'format.xml'), typeof files[path] === 'string' ? files[path] : textOf(files[path]));
});

test('ChartEx 외부 표준 계층 캐시: 빈 하위 항목·같은 자식 이름의 부모 경계 보존', () => {
  const data = { categories: ['A', 'B', ''], catLevels: [[{ text: '같음', start: 0, end: 1 }, { text: '같음', start: 2, end: 2 }], [{ text: '부모1', start: 0, end: 1 }, { text: '부모2', start: 2, end: 2 }]], series: [{ name: '값', values: [1, null, 3] }] };
  const xml = writeChartEx({ type: 'sunburst' }, data).replace(/<cx:extLst>[\s\S]*?<\/cx:extLst>/g, '');
  const chart = readChartEx(parseXml(xml));
  assert.deepEqual(chart.series[0].cache, [1, null, 3]);
  assert.deepEqual(chart.series[0].catCache, data.categories);
  assert.deepEqual(chart.series[0].catLevels, data.catLevels);
  assert.equal(child(kids(child(parseXml(xml), 'chartData'), 'data')[0], 'numDim').attrs.type, 'size');
});

test('ChartEx 숨긴 계열은 표준 hidden에 저장되고 원래 계열 번호/색 유지', () => {
  const files = noExtension(unzip(writeXlsx(book('boxWhisker', { hiddenSeries: [0], seriesFmt: [{ color: '#FF0000' }, { color: '#0000ff' }] }))));
  const { chart } = reopen(files);
  assert.equal(chart.series.length, 2); assert.deepEqual(chart.hiddenSeries, [0]);
  assert.equal(chart.seriesFmt[1].color, '#0000ff');
});

test('트리맵은 숫자 범주를 포함해 마지막 값열과 계층을 사용', () => {
  const d = chartData([['상위', '하위', '값'], [2025, '상반기', 10], [2026, '하반기', 20]], 'treemap');
  assert.equal(d.series.length, 1); assert.deepEqual(d.series[0].values, [10, 20]);
  assert.deepEqual(d.categories, ['상반기', '하반기']); assert.deepEqual(d.catLevels[0].map((s) => s.text), ['2025', '2026']);
});

test('ChartEx 가로 계층 표의 이름·범주·값 참조는 전치 방향 유지', () => {
  const horizontal = hierarchy[0].map((_, c) => hierarchy.map((r) => r[c]));
  for (const type of ['treemap', 'sunburst']) {
    const wb = book(type, { byRows: true }, horizontal), before = chartModelData(wb, 0, wb.sheets[0].charts[0]);
    const files = noExtension(unzip(writeXlsx(wb)));
    const { wb: loaded, chart } = reopen(files), after = chartModelData(loaded, 0, chart);
    assert.deepEqual(after.series.map((s) => [s.name, s.values]), before.series.map((s) => [s.name, s.values]));
    assert.deepEqual(after.categories, before.categories); assert.deepEqual(after.catLevels, before.catLevels);
    if (process.env.CHARTEX_FIXTURE_DIR) writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, `${type}-horizontal.xml`), files[path]);
  }
});

test('지도 ChartEx는 자체 확장 없이도 표준 최소·중간·최대 색을 보존', () => {
  const colors = { mapLowColor: '#001122', mapMidColor: '#334455', mapHighColor: '#667788' };
  const files = noExtension(unzip(writeXlsx(book('map', colors))));
  const { chart } = reopen(files);
  for (const [key, value] of Object.entries(colors)) assert.equal(chart[key], value);
  if (process.env.CHARTEX_FIXTURE_DIR) writeFileSync(join(process.env.CHARTEX_FIXTURE_DIR, 'map-colors.xml'), files[path]);
});

test('ChartEx 낯선 레이아웃·빈 차트는 지원된다고 오인하지 않음', () => {
  assert.equal(isChartEx({ type: 'not-chart' }), false);
  assert.equal(readChartEx(parseXml(`<cx:chartSpace xmlns:cx="${CHARTEX_NS}"><cx:chartData/><cx:chart><cx:plotArea><cx:plotAreaRegion><cx:series layoutId="unknown"/></cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>`)), null);
});

test('ChartEx 표준에 없는 위셀 서식도 허용된 확장으로 왕복 보존', () => {
  const options = { axes: { x: { reverse: true, hide: true }, y: { reverse: true, min: -10, max: 30, major: 5, numFmt: '0.0', log: 10 } }, seriesFmt: [{ color: '#123456', labels: false, labelPos: 'outEnd', lineWidth: 3, dash: 'dash', marker: 'diamond', markerSize: 9, trend: { type: 'linear', equation: true, r2: true }, outline: '#654321', pointColors: { 1: '#778899' } }], labels: true, dataTable: true, gap: 80, legend: 't' };
  const { wb, chart } = reopen(unzip(writeXlsx(book('waterfall', options))));
  for (const [key, value] of Object.entries(options)) assert.deepEqual(chart[key], value, key);
  const twice = reopen(unzip(writeXlsx(wb))).chart;
  for (const [key, value] of Object.entries(options)) assert.deepEqual(twice[key], value, key);
});

test('ChartEx 외부 확장 데이터는 종류·데이터 참조·프로토타입을 덮어쓰지 않음', () => {
  const xml = writeChartEx({ type: 'funnel' }, { categories: ['A'], series: [{ name: '값', values: [1] }] });
  const root = parseXml(xml), props = descendants(root, 'props')[0];
  props.attrs.json = '{"type":"unknown","series":[{"val":"secret"}],"axes":{"__proto__":{"polluted":1},"y":{"min":0}},"labels":"wrong","hiddenSeries":[-1]}';
  const chart = readChartEx(root);
  assert.equal(chart.type, 'funnel'); assert.deepEqual(chart.series[0].cache, [1]);
  assert.equal(chart.labels, undefined); assert.equal(chart.hiddenSeries, undefined);
  assert.deepEqual(chart.axes, { y: { min: 0 } }); assert.equal({}.polluted, undefined);
});
