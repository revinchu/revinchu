import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { chartModelData, renderChartSvg } from '../src/chart.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants, child } from '../src/xml.js';

test('여러 거품 계열: Y/크기 열 쌍과 숫자 X 참조를 XLSX에 보존', () => {
  const rows = [['X', 'Y A', '크기 A', 'Y B', '크기 B'], [1, 2, 10, 6, 40], [4, 7, 20, 9, 70], [9, 12, 30, 11, 100]], wb = new Workbook();
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => wb.setInput(0, r, c, String(v)))));
  wb.sheets[0].charts = [{ id: 'b', type: 'bubble', x: 0, y: 0, w: 600, h: 400, range: { r1: 0, c1: 0, r2: 3, c2: 4 } }];
  const data = chartModelData(wb, 0, wb.sheets[0].charts[0]), bytes = writeXlsx(wb), loaded = new Workbook(readXlsx(bytes).data);
  const got = chartModelData(loaded, 0, loaded.sheets[0].charts[0]);
  for (const key of ['name', 'values', 'x', 'size']) assert.deepEqual(got.series.map((s) => s[key]), data.series.map((s) => s[key]), key);
  const ser = descendants(parseXml(textOf(unzip(bytes)['xl/charts/chart1.xml'])), 'ser');
  assert.equal(descendants(child(ser[1], 'yVal'), 'f')[0].text, 'Sheet1!$D$2:$D$4');
  assert.equal(descendants(child(ser[1], 'bubbleSize'), 'f')[0].text, 'Sheet1!$E$2:$E$4');
  assert.ok(child(child(ser[0], 'xVal'), 'numRef'));
  loaded.setInput(0, 1, 4, '250');
  assert.equal(chartModelData(loaded, 0, loaded.sheets[0].charts[0]).series[1].size[0], 250, '크기 참조는 이후 셀 편집도 반영');
});

test('거품 리터럴 X/크기: 참조 없는 표준 numLit도 반복 저장 뒤 유지', () => {
  let wb = new Workbook();
  wb.sheets[0].charts = [{ id: 'b', type: 'bubble', x: 0, y: 0, w: 600, h: 400, series: [{ name: { text: '리터럴' }, cache: [3, 7, 11], xCache: [2, 6, 15], sizeCache: [9, 25, 100] }] }];
  for (let i = 0; i < 2; i++) {
    wb = new Workbook(readXlsx(writeXlsx(wb)).data); const s = chartModelData(wb, 0, wb.sheets[0].charts[0]).series[0];
    assert.deepEqual(s.values, [3, 7, 11]); assert.deepEqual(s.x, [2, 6, 15]); assert.deepEqual(s.size, [9, 25, 100]);
  }
});

test('3차원 꺾은선: 필수 세 번째 계열 축을 연결하여 저장', () => {
  const wb = new Workbook();
  wb.sheets[0].charts = [{ id: 'l', type: 'line', threeD: true, x: 0, y: 0, w: 600, h: 400, series: [{ name: { text: 'A' }, cache: [1, 2, 3] }, { name: { text: 'B' }, cache: [4, 2, 6] }] }];
  const root = parseXml(textOf(unzip(writeXlsx(wb))['xl/charts/chart1.xml'])), group = descendants(root, 'line3DChart')[0];
  assert.equal(descendants(group, 'axId').length, 3);
  const z = descendants(root, 'serAx')[0]; assert.ok(z);
  assert.equal(child(z, 'axId').attrs.val, descendants(group, 'axId')[2].attrs.val);
});

test('쪼개진 원형/도넛/3D 원형: 표준 계열 explosion으로 저장하고 자체 확장 없이 복원', () => {
  for (const options of [{ type: 'pie' }, { type: 'doughnut' }, { type: 'pie', threeD: true }]) {
    const wb = new Workbook(); wb.sheets[0].charts = [{ id: 'p', x: 0, y: 0, w: 600, h: 400, ...options, explode: 12, series: [{ name: { text: 'A' }, cache: [1, 2, 3] }] }];
    const files = unzip(writeXlsx(wb)), path = 'xl/charts/chart1.xml';
    const root = parseXml(textOf(files[path])), ser = descendants(root, 'ser')[0];
    assert.equal(child(ser, 'explosion').attrs.val, '12');
    assert.ok(descendants(ser, 'dPt').every((p) => !child(p, 'explosion')), '중복 분리량을 더하지 않음');
    files[path] = textOf(files[path]).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '');
    const loaded = new Workbook(readXlsx(zip(files)).data); assert.equal(loaded.sheets[0].charts[0].explode, 12);
  }
});

test('Excel 재저장의 일반 분산형/방사형: 계열별 noFill/marker=none을 실제 종류에 반영', () => {
  for (const type of ['scatter', 'radar']) {
    const wb = new Workbook(); wb.sheets[0].charts = [{ id: 'n', type, x: 0, y: 0, w: 600, h: 400, legend: 'none', series: [{ name: { text: 'A' }, cache: [1, 4, 2], xCache: [2, 5, 9] }], seriesFmt: [{ color: '#123456' }] }];
    const files = unzip(writeXlsx(wb)), path = 'xl/charts/chart1.xml';
    files[path] = textOf(files[path]).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, '').replace('<c:scatterStyle val="marker"/>', '<c:scatterStyle val="lineMarker"/>').replace('<c:radarStyle val="standard"/>', '<c:radarStyle val="marker"/>');
    const loaded = new Workbook(readXlsx(zip(files)).data), c = loaded.sheets[0].charts[0], svg = parseXml(renderChartSvg(c, chartModelData(loaded, 0, c)));
    if (type === 'scatter') { assert.equal(c.scatterStyle, 'marker'); assert.equal(descendants(svg, 'path').filter((p) => p.attrs.stroke === '#123456').length, 0); }
    else { assert.notEqual(c.radarStyle, 'marker'); assert.equal(descendants(svg, 'circle').length, 0); }
  }
});
