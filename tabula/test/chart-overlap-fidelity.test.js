import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { chartModelData, renderChartSvg } from '../src/chart.js';
import { chartLabelPlacer } from '../src/chart-label-placement.js';
import { chartAnchorBox } from '../src/chart-drawing-anchor.js';
import { captureDrawingAnchors, reflowDrawingAnchors } from '../src/drawing-anchor.js';
import { Axis } from '../src/axis.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';

const ref = c => ({ r1: 1, c1: c, r2: 3, c2: c });
function book() {
  const w = new Workbook();
  for (let r = 1; r <= 3; r++) { w.setInput(0, r, 0, '항목' + r); w.setInput(0, r, 1, String(r * 10)); }
  w.sheets[0].charts = [{ id: 'c', type: 'column', x: 20, y: 20, w: 600, h: 300, series: [{ name: { text: '노출' }, cat: ref(0), val: ref(1) }] }];
  return w;
}
const round = w => {
  const f = unzip(writeXlsx(w));
  f['xl/charts/chart1.xml'] = textOf(f['xl/charts/chart1.xml']).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
  return new Workbook(readXlsx(zip(f)).data);
};
const group = (s, kind) => descendants(s, 'g').find(g => g.attrs['data-el'] === kind);

test('수동 그림 영역과 가로로 넓은 오른쪽 범례가 표준 XML에서 보존된다', () => {
  const w = book(), c = w.sheets[0].charts[0];
  Object.assign(c, { legend: 'r', legendLayout: { x: .65, y: .02, w: .33, h: .08 }, plotLayout: { x: .08, y: .16, w: .84, h: .65, target: 'inner' }, plotAreaFormat: { fillMode: 'solid', fill: '#fff' } });
  c.series = ['노출', '클릭', 'CTR'].map(name => ({ ...c.series[0], name: { text: name } }));
  const b = round(w), ch = b.sheets[0].charts[0];
  assert.deepEqual(ch.plotLayout, c.plotLayout);
  const svg = parseXml(renderChartSvg(ch, chartModelData(b, 0, ch)));
  const legend = descendants(group(svg, 'legend'), 'text');
  assert.deepEqual(legend.map(n => n.text), ['노출', '클릭', 'CTR']);
  assert.equal(new Set(legend.map(n => n.attrs.y)).size, 1);
  const plot = descendants(group(svg, 'plot'), 'rect')[0];
  assert.equal(Number(plot.attrs.x), 48); assert.equal(Number(plot.attrs.y), 48);
});

test('숨긴 행의 축 글자와 값은 화면에서 제외하고 XLSX 캐시는 모두 보존한다', () => {
  const w = book(), c = w.sheets[0].charts[0]; c.visibleOnly = true;
  w.sheets[0].hiddenRows = { 2: true };
  const visible = chartModelData(w, 0, c);
  assert.deepEqual(visible.categories, ['항목1', '항목3']);
  assert.deepEqual(visible.series[0].values, [10, 30]);
  assert.deepEqual(visible.series[0]._pi, [0, 2]);
  assert.deepEqual(chartModelData(w, 0, c, { sample: false }).series[0].values, [10, 20, 30]);
  const b = round(w); assert.equal(b.sheets[0].charts[0].visibleOnly, true);
  c.visibleOnly = false; assert.equal(round(w).sheets[0].charts[0].visibleOnly, false);
  assert.deepEqual(chartModelData(w, 0, c).series[0].values, [10, 20, 30]);
});

test('표시할 점이 없는 가져온 차트도 범례와 축을 유지하고 빈 범주 글자는 만들지 않는다', () => {
  const c = { type: 'combo', w: 900, h: 230, series: [{}, {}, {}], legend: 'r', legendLayout: { x: .6, y: .02, w: .36, h: .1 }, plotLayout: { x: .08, y: .16, w: .84, h: .65, target: 'inner' } };
  const data = { categories: ['', '', ''], series: ['Imps', 'Clicks', 'CTR'].map(name => ({ name, values: [null, null, null] })) };
  const svg = parseXml(renderChartSvg(c, data));
  assert.equal(descendants(group(svg, 'legend'), 'text').length, 3);
  assert.ok(group(svg, 'axis-y')); assert.equal(descendants(group(svg, 'axis-x'), 'text').length, 0);
});

test('자동 데이터 레이블은 충돌할 때만 이동하며 지정한 안쪽 위치는 유지한다', () => {
  const place = chartLabelPlacer({ x: 0, y: 3, w: 600, h: 250 });
  assert.deepEqual(place(300, 20, '570,169', 12), { x: 300, y: 20 });
  const percent = place(300, 21, '4.14%', 12);
  assert.ok(Math.abs(percent.x - 300) > 30 || Math.abs(percent.y - 21) > 15);
  assert.deepEqual(place(500, 80, '1.00%', 12), { x: 500, y: 80 });
  assert.deepEqual(place(300, 20, '고정', 12, 'middle', false), { x: 300, y: 20 });
  const svg = parseXml(renderChartSvg({ type: 'combo', w: 669, h: 290, legend: 'none', axes: { y: { min: 0, max: 600000, major: 100000 }, y2: { min: 0, max: .045, major: .005 } } }, { categories: ['배너'], series: [{ name: '노출', values: [570169], type: 'column', labels: true }, { name: 'CTR', values: [.0414], type: 'line', axis: 1, labels: true, labelPos: 'above', numFmt: '0.00%' }] }));
  const labels = descendants(svg, 'text').filter(n => n.attrs['data-el'] === 'label');
  assert.equal(labels.length, 2);
  assert.ok(Math.abs(Number(labels[0].attrs.x) - Number(labels[1].attrs.x)) > 30 || Math.abs(Number(labels[0].attrs.y) - Number(labels[1].attrs.y)) > 15, '가져온 above 레이블도 충돌을 피한다');
});

test('숨긴 행 안의 차트는 높이 0과 원래 앵커를 왕복하고 숨기기 해제로 복원된다', () => {
  const w = book(), files = unzip(writeXlsx(w));
  files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml']).replace('<sheetData>', '<sheetData><row r="1" ht="180" customHeight="1" hidden="1"/>');
  const marker = off => `<xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>${off}</xdr:rowOff>`;
  files['xl/drawings/drawing1.xml'] = textOf(files['xl/drawings/drawing1.xml']).replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, `<xdr:from>${marker(47625)}</xdr:from>`).replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, `<xdr:to>${marker(1905000).replace('<xdr:col>0', '<xdr:col>7')}</xdr:to>`);
  const b = new Workbook(readXlsx(zip(files)).data), c = b.sheets[0].charts[0];
  assert.equal(c.h, 0); assert.ok(c.collapsedAnchor);
  const back = round(b).sheets[0].charts[0];
  assert.equal(back.h, 0); assert.deepEqual(back.collapsedAnchor, c.collapsedAnchor);
  const cols = new Axis(80, {}, [], 100), hidden = new Axis(20, { 0: 240 }, [{ 0: true }], 100), shown = new Axis(20, { 0: 240 }, [], 100);
  assert.equal(chartAnchorBox(c, cols, hidden).h, 0);
  assert.equal(chartAnchorBox(c, cols, shown).h, 195);
  const sheet = { charts: [c] };
  const next = reflowDrawingAnchors(sheet, captureDrawingAnchors(sheet, cols, hidden), cols, shown, { moveOnly: true }).charts[0];
  assert.equal(next.h, 195); assert.equal(next.collapsedAnchor, undefined);
  const shiftedRows = new Axis(20, { 2: 240 }, [{ 2: true }], 100);
  const shifted = reflowDrawingAnchors(sheet, captureDrawingAnchors(sheet, cols, hidden), cols, shiftedRows, { shift: { axis: 'row', index: 0, count: 2 } }).charts[0];
  assert.equal(shifted.h, 0); assert.equal(shifted.collapsedAnchor.from.row, 2);
  const restored = reflowDrawingAnchors({ charts: [shifted] }, captureDrawingAnchors({ charts: [shifted] }, cols, shiftedRows), cols, new Axis(20, { 2: 240 }, [], 100)).charts[0];
  assert.equal(restored.h, 195);
});
