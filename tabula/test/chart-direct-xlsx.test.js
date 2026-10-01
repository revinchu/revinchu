import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { chartModelData } from '../src/chart.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';

const path = 'xl/charts/chart1.xml';
const rows = [['항목', '첫 계열', '둘째 계열'], ['가', 10, 100], ['나', 20, 200], ['다', 30, 300]];
function book(chart = {}) {
  const wb = new Workbook();
  wb.transact(() => rows.forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value)))));
  wb.sheets[0].charts = [{ id: 'direct', type: 'pie', x: 0, y: 0, w: 600, h: 400, title: '직접 편집', legend: 'b', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, ...chart }];
  return wb;
}
function filesFor(wb, strip = true) {
  const files = unzip(writeXlsx(wb));
  if (strip) for (const name of Object.keys(files)) if (name.endsWith('.xml')) files[name] = new TextEncoder().encode(textOf(files[name]).replace(/<(?:\w+:)?extLst\b[^>]*>[\s\S]*?<\/(?:\w+:)?extLst>/g, ''));
  return files;
}
const back = files => readXlsx(zip(files)).data.sheets[0].charts[0];
const dataOf = files => { const wb = new Workbook(readXlsx(zip(files)).data); return chartModelData(wb, 0, wb.sheets[0].charts[0]); };
const close = (a, b) => assert.ok(Math.abs(a - b) < 1e-9, `${a} != ${b}`);

test('차트 제목·범례 edge 위치와 크기는 전용 확장 없이 표준 manualLayout으로 왕복한다', () => {
  const titleLayout = { x: .13, y: .18 }, legendLayout = { x: .62, y: .34, w: .25, h: .4 };
  const files = filesFor(book({ titleLayout, legendLayout })), xml = textOf(files[path]), chart = back(files);
  assert.equal((xml.match(/<c:manualLayout>/g) ?? []).length, 2);
  assert.equal((xml.match(/<c:xMode val="edge"\/>/g) ?? []).length, 2);
  assert.match(xml, /<c:wMode val="factor"\/><c:hMode val="factor"\/>/);
  assert.deepEqual(chart.titleLayout, titleLayout); assert.deepEqual(chart.legendLayout, legendLayout);
});

test('외부 manualLayout의 edge 오른쪽/아래 좌표는 크기로 변환하고 factor 이동은 추측하지 않는다', () => {
  const files = filesFor(book({ titleLayout: { x: .1, y: .2, w: .6, h: .5 } }));
  files[path] = textOf(files[path]).replace('wMode val="factor"', 'wMode val="edge"').replace('hMode val="factor"', 'hMode val="edge"');
  const chart = back(files); close(chart.titleLayout.w, .5); close(chart.titleLayout.h, .3);
  files[path] = files[path].replace('xMode val="edge"', 'xMode val="factor"');
  assert.equal(back(files).titleLayout, undefined);
});

test('계열/점 분리와 점 색은 하나의 dPt에 합쳐지며 0·400 경계도 표준 왕복한다', () => {
  for (const type of ['pie', 'doughnut', 'pieOfPie', 'barOfPie']) {
    const source = { type, explode: 35, seriesFmt: [{ explode: 0, pointExplosion: { 0: 0, 1: 85, 2: 400 }, pointColors: { 1: '#123456' } }, { explode: 60, pointExplosion: { 2: 90 } }] };
    const wb = book(source), before = wb.serialize(), files = filesFor(wb), root = parseXml(textOf(files[path]));
    const series = descendants(root, 'ser');
    assert.equal(series.length, 2);
    assert.equal(child(series[0], 'explosion').attrs.val, '0'); assert.equal(child(series[1], 'explosion').attrs.val, '60');
    assert.equal(kids(series[0], 'dPt').length, 3); assert.equal(new Set(kids(series[0], 'dPt').map(p => child(p, 'idx').attrs.val)).size, 3);
    const chart = back(files); assert.deepEqual(chart.seriesFmt[0].pointExplosion, source.seriesFmt[0].pointExplosion);
    assert.equal(chart.seriesFmt[0].explode, 0); assert.equal(chart.seriesFmt[1].explode, 60); assert.equal(chart.seriesFmt[0].pointColors[1], '#123456');
    assert.deepEqual(wb.serialize(), before, '파일 저장이 원본 셀/차트 모델을 변경하지 않음');
  }
});

test('첫 계열 삭제 뒤 남은 원본 fi의 이름/범주/값 참조와 색·분리 형식을 저장 순서로 맞춘다', () => {
  for (const explicit of [false, true]) for (const strip of [false, true]) {
    const series = explicit ? [1, 2].map(c => ({ name: { ref: { r1: 0, c1: c, r2: 0, c2: c } }, cat: { r1: 1, c1: 0, r2: 3, c2: 0 }, val: { r1: 1, c1: c, r2: 3, c2: c } })) : undefined;
    const wb = book({ type: 'doughnut', series, hiddenSeries: [0], seriesFmt: [{ color: '#ff0000', explode: 5 }, { color: '#00aa00', explode: 50, pointExplosion: { 1: 80 }, pointColors: { 2: '#abcdef' } }] });
    const files = filesFor(wb, strip), chart = back(files), data = dataOf(files), ser = descendants(parseXml(textOf(files[path])), 'ser');
    assert.equal(ser.length, 1); assert.equal(chart.series.length, 1); assert.equal(chart.hiddenSeries, undefined);
    assert.equal(descendants(child(ser[0], 'val'), 'f')[0].text, "Sheet1!$C$2:$C$4");
    assert.equal(data.series.length, 1); assert.equal(data.series[0].name, '둘째 계열'); assert.deepEqual(data.series[0].values, [100, 200, 300]);
    assert.deepEqual(data.categories, ['가', '나', '다']); assert.equal(chart.seriesFmt[0].explode, 50);
    assert.equal(chart.seriesFmt[0].pointExplosion[1], 80); assert.equal(chart.seriesFmt[0].pointColors[2].toLowerCase(), '#abcdef');
  }
});

test('마지막 계열을 삭제해도 빈 차트를 유지하고 재열기에서 원본 셀을 다시 그리지 않는다', () => {
  for (const type of ['pie', 'line', 'column']) {
    const wb = book({ type, hiddenSeries: [0, 1], titleLayout: { x: .2, y: .1 } }), before = wb.serialize();
    const files = filesFor(wb, false), chart = back(files), data = dataOf(files);
    assert.ok(chart); assert.equal(chart.type, type); assert.equal(chart.series.length, 0); assert.equal(chart.range, undefined); assert.equal(chart.hiddenSeries, undefined);
    assert.equal(data.series.length, 0); assert.deepEqual(chart.titleLayout, { x: .2, y: .1 });
    assert.deepEqual(wb.serialize(), before);
  }
});

test('ChartEx 숨긴 계열은 native hidden으로 한 번만 반영하며 원본 참조를 유지한다', () => {
  for (const hiddenSeries of [[0], [0, 1]]) {
    const files = filesFor(book({ type: 'boxWhisker', hiddenSeries, seriesFmt: [{ color: '#ff0000' }, { color: '#00aa00' }] })), chart = back(files), data = dataOf(files);
    assert.equal(chart.series.length, 2); assert.deepEqual(chart.hiddenSeries, hiddenSeries);
    assert.equal(data.series.length, 2 - hiddenSeries.length);
    if (data.series.length) { assert.equal(data.series[0].name, '둘째 계열'); assert.deepEqual(data.series[0].values, [100, 200, 300]); }
  }
});

test('ChartEx 직접 위치는 WIXEL 보조 확장에서 보존하고 native manualLayout으로 오인하지 않는다', () => {
  const titleLayout = { x: .12, y: .23 }, legendLayout = { x: .62, y: .44, w: .2, h: .3 };
  const wb = book({ type: 'waterfall', titleLayout, legendLayout }), files = filesFor(wb, false), chart = back(files);
  assert.deepEqual(chart.titleLayout, titleLayout); assert.deepEqual(chart.legendLayout, legendLayout);
  assert.doesNotMatch(textOf(files[path]), /manualLayout/);
  const native = back(filesFor(wb)); assert.equal(native.titleLayout, undefined); assert.equal(native.legendLayout, undefined);
});

test('직접 편집의 위치/분리/계열삭제를 Undo/Redo한 현재 상태가 저장된다', () => {
  const wb = book(), original = structuredClone(wb.sheets[0].charts);
  wb.transact(() => wb.setSheetProp(0, 'charts', [{ ...original[0], hiddenSeries: [0], titleLayout: { x: .21, y: .31 }, seriesFmt: [{}, { explode: 70, pointExplosion: { 1: 100 } }] }]));
  wb.undo(); assert.deepEqual(wb.sheets[0].charts, original); assert.equal(back(filesFor(wb)).series.length, 2);
  wb.redo(); const chart = back(filesFor(wb)); assert.equal(chart.series.length, 1); assert.deepEqual(chart.titleLayout, { x: .21, y: .31 }); assert.equal(chart.seriesFmt[0].pointExplosion[1], 100);
  for (let r = 1; r < rows.length; r++) for (let c = 1; c < 3; c++) assert.equal(wb.getValue(0, r, c), rows[r][c]);
});

test('범주 숨김은 원본 참조/캐시/점 번호를 유지하고 WIXEL에서만 한 번 적용한다', () => {
  const wb = book({ range: { r1: 0, c1: 0, r2: 3, c2: 1 }, hiddenCats: [0], seriesFmt: [{ pointExplosion: { 1: 85, 2: 140 }, pointColors: { 1: '#123456', 2: '#abcdef' } }] });
  const before = wb.serialize();
  for (const strip of [false, true]) {
    const files = filesFor(wb, strip), root = parseXml(textOf(files[path])), ser = descendants(root, 'ser')[0], chart = back(files), data = dataOf(files);
    const values = descendants(child(ser, 'val'), 'pt').map(pt => Number(child(pt, 'v').text));
    assert.deepEqual(values, [10, 20, 30], '원본 numRef를 가진 Excel 캐시도 전체 값을 유지');
    assert.equal(descendants(child(ser, 'val'), 'f')[0].text, 'Sheet1!$B$2:$B$4');
    assert.deepEqual(chart.seriesFmt[0].pointExplosion, { 1: 85, 2: 140 });
    assert.equal(chart.seriesFmt[0].pointColors[1].toLowerCase(), '#123456'); assert.equal(chart.seriesFmt[0].pointColors[2].toLowerCase(), '#abcdef');
    assert.deepEqual(data.categories, strip ? ['가', '나', '다'] : ['나', '다']);
    assert.deepEqual(data.series[0].values, strip ? [10, 20, 30] : [20, 30]);
    const twice = new Workbook(readXlsx(zip(files)).data), again = back(filesFor(twice, strip));
    assert.deepEqual(again.seriesFmt[0].pointExplosion, { 1: 85, 2: 140 });
  }
  assert.deepEqual(wb.serialize(), before);
});
