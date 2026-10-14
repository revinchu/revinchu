import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, digitWidth, width2pxM } from '../src/xlsx.js';
import { chartModelData, renderChartSvg } from '../src/chart.js';
import { categoryAxisLayout } from '../src/chart-axis-labels.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, descendants } from '../src/xml.js';
import { formatCode } from '../src/format.js';

const chartPath = 'xl/charts/chart1.xml';
const xmlOf = files => typeof files[chartPath] === 'string' ? files[chartPath] : textOf(files[chartPath]);
const ref = (c, r1 = 1, r2 = 31) => ({ r1, c1: c, r2, c2: c });
function fixture() {
  const wb = new Workbook(); wb.defaultFont = { name: '돋움', size: 11 };
  ['날짜', '노출', 'CTR'].forEach((v, c) => wb.setInput(0, 0, c, v));
  for (let r = 1; r <= 31; r++) {
    wb.setInput(0, r, 0, String(44681 + r));
    wb.setInput(0, r, 1, r < 27 ? String(r === 17 ? 502189 : 47475 + r * 43) : '=""');
    wb.setInput(0, r, 2, r < 27 ? String(.01 + r % 7 * .002) : '=""');
  }
  const chart = { id: 'report-chart', type: 'combo', x: 0, y: 0, w: 1200, h: 290, gap: 150, legend: 'b', gridY: false,
    axes: { x: { type: 'date', numFmt: 'yy/mm/dd', sourceLinked: true, lineColor: '#d9d9d9', lineWidth: 1 }, y: { numFmt: '#,##0', sourceLinked: true, lineNone: true }, y2: { numFmt: '0.00%', sourceLinked: true, lineNone: true } },
    series: [1, 2].map(c => ({ name: { ref: ref(c, 0, 0) }, cat: ref(0), val: ref(c) })),
    seriesFmt: [{ type: 'column', color: '#4f81bd', labels: false, numFmt: '#,##0', pointLabels: { 0: { labels: true, labelPos: 'outEnd' } } }, { type: 'line', axis: 1, color: '#9bbb59', marker: 'none', numFmt: '0.00%' }] };
  wb.sheets[0].charts = [chart]; return wb;
}
function nativeRound(wb) {
  const files = unzip(writeXlsx(wb));
  files[chartPath] = xmlOf(files).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
  return { files, wb: new Workbook(readXlsx(zip(files)).data) };
}
const draw = (wb) => parseXml(renderChartSvg(wb.sheets[0].charts[0], chartModelData(wb, 0, wb.sheets[0].charts[0])));
const group = (root, el) => descendants(root, 'g').find(g => g.attrs['data-el'] === el);

test('묶은 막대의 간격은 클러스터가 아닌 개별 막대 너비의 백분율이다', () => {
  const svg = parseXml(renderChartSvg({ type: 'column', w: 800, h: 300, gap: 200, overlap: 0, legend: 'none' }, {
    categories: ['A', 'B'], series: [{ name: '첫째', values: [10, 20] }, { name: '둘째', values: [15, 25] }],
  }));
  const bars = descendants(svg, 'rect').filter(n => n.attrs['data-s'] === '0');
  const band = Number(bars[1].attrs.x) - Number(bars[0].attrs.x);
  assert.ok(Math.abs(Number(bars[0].attrs.width) + 1 - band / 4) < .1, '막대 두 개와 200% 간격이 한 범주 폭을 채운다');
});

test('돋움 11pt 열 너비는 독립 Excel 측정의 9px MDW와 일치한다', () => {
  for (const name of ['돋움', 'Dotum']) assert.equal(digitWidth({ name, size: 11 }), 9);
  assert.deepEqual([1.77734375, 12, 22.88671875, 14.109375].map(w => width2pxM(w, 9)), [16, 108, 206, 127]);
  const wb = fixture(); wb.sheets[0].colWidths = { 0: 16, 1: 108, 2: 206 };
  assert.deepEqual(nativeRound(wb).wb.sheets[0].colWidths, wb.sheets[0].colWidths);
});

test('날짜 축·연결된 숫자 서식·숨긴 축선·기본 간격도 확장 없이 저장된다', () => {
  const original = fixture(), { files, wb } = nativeRound(original), xml = xmlOf(files), ch = wb.sheets[0].charts[0];
  assert.match(xml, /<c:dateAx>/); assert.match(xml, /<c:cat><c:numRef>/);
  assert.match(xml, /<c:formatCode>yy\/mm\/dd<\/c:formatCode>/);
  assert.equal(ch.axes.x.type, 'date'); assert.equal(ch.axes.y2.numFmt, '0.00%');
  assert.equal(ch.axes.y2.sourceLinked, true); assert.equal(ch.axes.y2.lineNone, true);
  assert.equal(ch.gap, 150); assert.equal(ch.gridY, false);
  assert.deepEqual(nativeRound(wb).wb.sheets[0].charts[0].axes, ch.axes);
});

test('수식의 빈 문자열은 0으로 이어지되 실제 빈 셀과 계열 이름은 빈 값이다', () => {
  const wb = fixture(), ch = wb.sheets[0].charts[0];
  wb.setInput(0, 0, 1, '=""'); wb.setInput(0, 30, 1, '');
  const data = chartModelData(wb, 0, ch);
  assert.equal(data.series[0].name, ''); assert.equal(data.series[0].values[26], 0);
  assert.equal(data.series[0].values[29], null); assert.equal(data.series[0].values[30], 0);
  assert.equal(wb.getValue(0, 27, 1), '', '원본 셀 값은 바꾸지 않는다');
});

test('첫 점에만 표시한 레이블은 계열 전체 표시가 꺼져 있어도 왕복·렌더된다', () => {
  const { wb } = nativeRound(fixture()), ch = wb.sheets[0].charts[0];
  assert.deepEqual(ch.seriesFmt[0].pointLabels, { 0: { labels: true, labelPos: 'outEnd' } });
  const labels = descendants(draw(wb), 'text').filter(n => n.attrs['data-el'] === 'label');
  assert.equal(labels.length, 1); assert.equal(labels[0].attrs['data-p'], '0'); assert.equal(labels[0].text, '47,518');
  ch.seriesFmt[0].pointLabels[0].hide = true;
  assert.equal(descendants(draw(nativeRound(wb).wb), 'text').filter(n => n.attrs['data-el'] === 'label').length, 0);
});

test('31개 날짜는 단어 중간 줄바꿈 없이 회전하고 0~3% 축과 10만 단위 눈금이 표시된다', () => {
  const wb = fixture(), root = draw(wb);
  const x = descendants(group(root, 'axis-x'), 'text');
  assert.equal(x.length, 31); assert.ok(x.every(n => !child(n, 'tspan') && /rotate\(-45\)/.test(n.attrs.transform)));
  const y = descendants(group(root, 'axis-y'), 'text').map(n => n.text), y2 = descendants(group(root, 'axis-y2'), 'text').map(n => n.text);
  assert.deepEqual(y, ['0', '100,000', '200,000', '300,000', '400,000', '500,000', '600,000']);
  assert.equal(y2[0], '0.00%'); assert.equal(y2.at(-1), '2.50%');
  const manual = categoryAxisLayout(['긴 날짜 항목'], { length: 30, depth: 120, noWrap: true, rotation: 0 });
  assert.equal(manual.angle, 0, '명시한 각도는 자동 회전보다 우선한다');
});

test('단일 보조 계열도 0 기준 4.50% 자동 눈금, 숨긴 선, 수동 범례 높이를 지킨다', () => {
  const wb = fixture(), ch = wb.sheets[0].charts[0];
  ch.w = 668; ch.h = 290; delete ch.axes.x.type; ch.series.forEach(s => { s.cat = ref(0, 1, 1); s.val = ref(s.val.c1, 1, 1); });
  wb.setInput(0, 1, 0, '배너'); wb.setInput(0, 1, 1, '516608'); wb.setInput(0, 1, 2, '.0412');
  ch.seriesFmt[1].lineNone = true;
  ch.legendLayout = { x: .3912, y: .8923, w: .2212, h: .0757 }; ch.legendOverlay = false;
  ch.series[0].name = { text: ' 실제노출량 ' };
  const back = nativeRound(wb).wb, root = draw(back);
  assert.equal(descendants(group(root, 'axis-y2'), 'text').at(-1).text, '4.50%');
  assert.ok(descendants(root, 'path').some(p => p.attrs['data-s'] === '1' && p.attrs.stroke === 'none'));
  assert.deepEqual(descendants(group(root, 'legend'), 'text').map(n => n.text.trim()), ['실제노출량']);
  ch.hiddenLegend = [0]; assert.equal(nativeRound(wb).wb.sheets[0].charts[0].hiddenLegend[0], 0);
});

test('가져온 빈 차트는 빈 차트 영역을 유지하며 안내 문구를 덧붙이지 않는다', () => {
  const svg = renderChartSvg({ type: 'column', w: 500, h: 280, series: [] }, { categories: [], series: [] });
  assert.doesNotMatch(svg, /표시할 숫자/); assert.match(svg, /<rect/);
});

test('한국어 Excel 날짜 구분자와 명시한 슬래시·로캘·분수를 구별한다', () => {
  assert.equal(formatCode(44682, 'yy/mm/dd').text, '22-05-01');
  assert.equal(formatCode(44682, 'yy"/"mm"/"dd').text, '22/05/01');
  assert.equal(formatCode(44682, 'yy\\/mm\\/dd').text, '22/05/01');
  assert.equal(formatCode(44682, '[$-409]yy/mm/dd').text, '22/05/01');
  assert.equal(formatCode(44682, '[$-412]yy/mm/dd').text, '22-05-01');
  assert.equal(formatCode(.5, '# ?/?').text.trim(), '1/2');
});

test('선 없는 주 눈금선은 XML 요소 존재만으로 다시 표시하지 않는다', () => {
  const { files } = nativeRound(fixture());
  files[chartPath] = xmlOf(files).replace('<c:axPos val="l"/>', '<c:axPos val="l"/><c:majorGridlines><c:spPr><a:ln><a:noFill/></a:ln></c:spPr></c:majorGridlines>');
  assert.equal(readXlsx(zip(files)).data.sheets[0].charts[0].gridY, false);
});
