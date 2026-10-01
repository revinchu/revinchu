import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';

function book(chart) {
  const wb = new Workbook();
  const rows = [['월', '금액', '비율', '추가'], ['1월', '10', '0.2', '4'], ['2월', '20', '0.4', '5'], ['3월', '30', '0.5', '6'], ['4월', '40', '0.6', '7']];
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, v))));
  wb.sheets[0].charts = [{ id: 'combo-test', type: 'combo', range: { r1: 0, c1: 0, r2: 4, c2: 3 }, x: 0, y: 0, w: 600, h: 360, ...chart }];
  return wb;
}
const back = (files) => readXlsx(zip(files)).data.sheets[0].charts[0];
const path = 'xl/charts/chart1.xml';
const axes = { x: { title: '기간', hide: true }, y: { title: '금액', min: 0, max: 50, major: 10, hide: true, reverse: true, numFmt: '#,##0' }, y2: { title: '비율', min: 0, max: 1, major: 0.2, hide: true, numFmt: '0%' } };

test('XLSX 콤보: 계열별 종류·순서·주축/보조축과 축 서식을 왕복 보존', () => {
  const seriesFmt = [{ type: 'line', axis: 1 }, { type: 'column', axis: 0 }, { type: 'area', axis: 0 }];
  const files = unzip(writeXlsx(book({ seriesFmt, axes })));
  const chart = back(files);
  assert.equal(chart.type, 'combo');
  assert.deepEqual(chart.seriesFmt.map((s) => ({ type: s.type, axis: s.axis ?? 0 })), seriesFmt);
  assert.deepEqual(chart.axes, axes);
  const twice = back(unzip(writeXlsx(book({ ...chart }))));
  assert.deepEqual(twice.seriesFmt.map((s) => ({ type: s.type, axis: s.axis ?? 0 })), seriesFmt);
  assert.deepEqual(twice.axes, axes);
});

test('모든 계열을 보조축에 지정해도 XLSX가 주축으로 바꾸지 않음', () => {
  const seriesFmt = Array.from({ length: 3 }, () => ({ type: 'line', axis: 1 }));
  const files = unzip(writeXlsx(book({ seriesFmt, axes })));
  const plot = descendants(parseXml(textOf(files[path])), 'plotArea')[0];
  assert.deepEqual(kids(child(plot, 'lineChart'), 'axId').map((n) => n.attrs.val), ['333333333', '444444444']);
  const chart = back(files);
  assert.equal(chart.series.length, 3);
  assert.deepEqual(chart.seriesFmt.map((s) => s.axis), [1, 1, 1]);
  assert.deepEqual(chart.axes.y2, axes.y2);
  // Excel이 사용하지 않는 주축을 생략한 파일도 그룹이 하나라는 이유로 보조축을 버리지 않습니다.
  files[path] = textOf(files[path]).replace(/<c:catAx><c:axId val="111111111"\/>[\s\S]*?<\/c:catAx>/, '').replace(/<c:valAx><c:axId val="222222222"\/>[\s\S]*?<\/c:valAx>/, '');
  const onlySecondary = back(files);
  assert.deepEqual(onlySecondary.seriesFmt.map((s) => s.axis), [1, 1, 1]);
  assert.equal(onlySecondary.axes.y, undefined);
  assert.deepEqual(onlySecondary.axes.y2, axes.y2);
});

test('같은 종류의 계열도 서로 다른 축을 유지', () => {
  const seriesFmt = [{ type: 'column', axis: 0 }, { type: 'column', axis: 1 }, { type: 'column', axis: 0 }];
  const chart = back(unzip(writeXlsx(book({ seriesFmt, axes }))));
  assert.equal(chart.type, 'combo', '같은 종류의 계열도 WIXEL 콤보 편집 상태를 보존');
  assert.deepEqual(chart.seriesFmt.map((s) => ({ type: s.type, axis: s.axis ?? 0 })), seriesFmt);
  assert.deepEqual(chart.axes, axes);
});

test('콤보 주축 기본 preset 설정이 XLSX 확장에 유지', () => {
  const chart = back(unzip(writeXlsx(book({ comboAxis: 'primary', seriesFmt: [{ type: 'column', axis: 0 }, { type: 'column', axis: 0 }, { type: 'line', axis: 0 }] }))));
  assert.equal(chart.comboAxis, 'primary');
  assert.deepEqual(chart.seriesFmt.map((s) => s.axis ?? 0), [0, 0, 0]);
});
