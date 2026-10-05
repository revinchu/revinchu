import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, descendants } from '../src/xml.js';

const path = 'xl/charts/chart1.xml';
const xmlText = value => typeof value === 'string' ? value : textOf(value);
const read = files => readXlsx(zip(files)).data.sheets[0].charts[0];
function book(patch = {}) {
  const wb = new Workbook();
  wb.transact(() => [['항목', '금액'], ['첫 번째 광고', '10'], ['두 번째 광고', '20'], ['세 번째 광고', '30']].forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v))));
  wb.sheets[0].charts = [{ id: 'axis-label-test', type: 'column', x: 0, y: 0, w: 420, h: 280, range: { r1: 0, c1: 0, r2: 3, c2: 1 }, ...patch }];
  return wb;
}
function withoutExtension(files) {
  files[path] = xmlText(files[path]).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
  return files;
}
const category = files => descendants(parseXml(xmlText(files[path])), 'catAx')[0];

test('항목 축 간격·회전은 표준 XML로 저장되고 WIXEL 확장 없이 왕복된다', () => {
  for (const type of ['column', 'bar', 'line', 'area', 'combo']) {
    const axes = { x: { title: '광고 이름', labelInterval: 1, labelRotation: -45 }, y: { title: '금액', min: 0, max: 40, major: 10 } };
    const files = withoutExtension(unzip(writeXlsx(book({ type, axisSize: 11, axes }))));
    const ax = category(files);
    assert.equal(child(ax, 'tickLblSkip')?.attrs.val, '1', type);
    assert.equal(child(child(ax, 'txPr'), 'bodyPr')?.attrs.rot, '-2700000', type);
    assert.equal(descendants(child(ax, 'txPr'), 'defRPr')[0]?.attrs.sz, '1100', type);
    const chart = read(files);
    const nativeAxes={x:{...axes.x,size:11},y:{...axes.y,size:11}};
    assert.deepEqual(chart.axes, nativeAxes, type);
    assert.equal(chart.axisSize, 11, type);
    assert.deepEqual(read(withoutExtension(unzip(writeXlsx(book(chart))))).axes, nativeAxes, type);
  }
});

test('자동 간격·방향은 생략하고 명시한 0°·1 간격은 자동으로 바꾸지 않는다', () => {
  const auto = withoutExtension(unzip(writeXlsx(book({ axes: { x: { title: '기간' } } }))));
  assert.equal(child(category(auto), 'tickLblSkip'), null);
  assert.equal(child(category(auto), 'txPr'), null);
  assert.deepEqual(read(auto).axes.x, { title: '기간' });
  const manual = withoutExtension(unzip(writeXlsx(book({ axes: { x: { labelInterval: 1, labelRotation: 0 } } }))));
  assert.equal(child(child(category(manual), 'txPr'), 'bodyPr')?.attrs.rot, '0');
  assert.deepEqual(read(manual).axes.x, { labelInterval: 1, labelRotation: 0 });
});

test('Excel에서 바꾼 표준 간격·양수/음수 각도를 다시 읽는다', () => {
  const files = withoutExtension(unzip(writeXlsx(book({ axes: { x: { title: '항목', labelInterval: 1, labelRotation: -45 } } }))));
  files[path] = xmlText(files[path]).replace('<c:tickLblSkip val="1"/>', '<c:tickLblSkip val="3"/>').replace('rot="-2700000"', 'rot="1800000"');
  assert.deepEqual(read(files).axes.x, { title: '항목', labelInterval: 3, labelRotation: 30 });
});

test('날짜 축의 기존 제목·글꼴과 명시 회전을 읽되 날짜 단위 기능을 추가하지 않는다', () => {
  const files = withoutExtension(unzip(writeXlsx(book({ axisSize: 10, axes: { x: { title: '날짜', labelRotation: 0 } } }))));
  files[path] = xmlText(files[path]).replace(/c:catAx/g, 'c:dateAx');
  const chart = read(files);
  assert.deepEqual(chart.axes.x, { title: '날짜', size: 10, labelRotation: 0 });
  assert.equal(chart.axisSize, 10);
});

test('잘못된 간격·회전은 자동으로 처리하고 ±90°는 보존한다', () => {
  for (const invalid of [0, -2, 1.5, NaN, Infinity]) {
    const files = withoutExtension(unzip(writeXlsx(book({ axes: { x: { labelInterval: invalid, labelRotation: Infinity } } }))));
    assert.equal(child(category(files), 'tickLblSkip'), null);
    assert.equal(child(category(files), 'txPr'), null);
    assert.equal(read(files).axes?.x?.labelInterval, undefined);
  }
  for (const angle of [-90, 90]) {
    const files = withoutExtension(unzip(writeXlsx(book({ axes: { x: { labelRotation: angle } } }))));
    assert.equal(read(files).axes.x.labelRotation, angle);
  }
  const files = withoutExtension(unzip(writeXlsx(book({ axes: { x: { labelInterval: 2, labelRotation: -45 } } }))));
  files[path] = xmlText(files[path]).replace('tickLblSkip val="2"', 'tickLblSkip val="0"').replace('rot="-2700000"', 'rot="not-an-angle"');
  assert.equal(read(files).axes?.x?.labelInterval, undefined);
  assert.equal(read(files).axes?.x?.labelRotation, undefined);
});

test('숫자 가로축에는 항목 간격·각도를 내보내지 않는다', () => {
  for (const type of ['scatter', 'bubble']) {
    const files = withoutExtension(unzip(writeXlsx(book({ type, axes: { x: { min: 0, max: 50, labelInterval: 2, labelRotation: -45 } } }))));
    assert.equal(category(files), undefined);
    assert.equal(descendants(parseXml(xmlText(files[path])), 'tickLblSkip').length, 0);
    assert.equal(read(files).axes.x.labelInterval, undefined);
    assert.equal(read(files).axes.x.labelRotation, undefined);
  }
});

test('ChartEx의 항목 간격·각도는 기존 보완 확장으로 왕복 유지한다', () => {
  for (const type of ['waterfall', 'boxWhisker', 'histogram', 'pareto']) {
    const chart = read(unzip(writeXlsx(book({ type, axes: { x: { title: '항목', labelInterval: 2, labelRotation: -45 } } }))));
    assert.equal(chart.axes.x.labelInterval, 2, type);
    assert.equal(chart.axes.x.labelRotation, -45, type);
    assert.equal(chart.axes.x.title, '항목', type);
  }
});
