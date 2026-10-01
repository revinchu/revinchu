import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';

const base = { id: 's1', kind: 'rect', name: '효과 도형', x: 20, y: 30, w: 220, h: 130, fill: '#4472c4', stroke: '#123456', text: '텍스트', color: '#102030' };
const make = (shapes) => { const wb = new Workbook(); wb.transact(() => wb.setSheetProp(0, 'shapes', shapes)); return wb; };
const saved = (shapes) => {
  const bytes = writeXlsx(make(shapes)), files = unzip(bytes);
  return { bytes, files, xml: textOf(files['xl/drawings/drawing1.xml']), shapes: readXlsx(bytes).data.sheets[0].shapes };
};
const close = (actual, expected) => assert.ok(Math.abs(actual - expected) <= 0.0002, `${actual} ≠ ${expected}`);

test('도형: 표준 DrawingML 그라데이션·알파·그림자·네온·부드러운 가장자리 왕복', () => {
  const shape = { ...base, fillOpacity: 0.35, strokeOpacity: 0.7, grad: { ang: 135, stops: [[0, '#112233'], [0.4, '#aabbcc'], [1, '#ff7700']] }, shadow: { dx: -4, dy: 3, blur: 6, color: '#090807', opacity: 0.2 }, glow: { size: 7, color: '#ff2200', opacity: 0.45 }, soft: 2.5 };
  const { xml, shapes: [back] } = saved([shape]);
  assert.match(xml, /<a:gradFill/);
  assert.match(xml, /<a:lin ang="8100000"/);
  assert.match(xml, /<a:alpha val="35000"/);
  assert.match(xml, /<a:glow rad="66675"/);
  assert.match(xml, /<a:outerShdw blurRad="57150" dist="47625"/);
  assert.match(xml, /<a:softEdge rad="23813"/);
  assert.ok(xml.indexOf('<a:glow') < xml.indexOf('<a:outerShdw'));
  assert.ok(xml.indexOf('<a:outerShdw') < xml.indexOf('<a:softEdge'));
  assert.deepEqual(back.grad, shape.grad);
  assert.equal(back.fillOpacity, 0.35);
  assert.equal(back.strokeOpacity, 0.7);
  assert.equal(back.name, shape.name);
  for (const k of ['dx', 'dy', 'blur', 'opacity']) close(back.shadow[k], shape.shadow[k]);
  assert.equal(back.shadow.color, shape.shadow.color);
  assert.deepEqual(back.glow, shape.glow);
  close(back.soft, 2.5);
});

test('도형: 완전 투명 단색·선 투명도·연결선 대시와 개별 그라데이션 중지점 알파', () => {
  const shape = { ...base, grad: { ang: 270, stops: [[0, '#112233', 0], [1, '#aabbcc', 0.8]] } };
  const { xml, shapes } = saved([{ ...base, fillOpacity: 0 }, { ...base, kind: 'line', strokeOpacity: 0, dash: 'dashDot', arrow: 'both' }, shape]);
  assert.match(xml, /<a:alpha val="0"/);
  assert.match(xml, /<a:prstDash val="dashDot"/);
  assert.equal(shapes[0].fillOpacity, 0);
  assert.equal(shapes[1].strokeOpacity, 0);
  assert.equal(shapes[1].dash, 'dashDot');
  assert.equal(shapes[1].arrow, 'both');
  assert.deepEqual(shapes[2].grad, shape.grad);
  assert.equal(shapes[2].fillOpacity, undefined);
});

test('도형: 글꼴·굵게·기울임·밑줄·여백·줄 바꿈·세로 맞춤과 빈 텍스트 왕복', () => {
  const shape = { ...base, font: '맑은 고딕', size: 17, bold: true, italic: true, underline: true, pad: [1, 2, 3, 4], nowrap: true, align: 'right', valign: 'bottom' };
  const { xml, shapes } = saved([shape, { ...shape, text: '', kind: 'textbox' }]);
  assert.match(xml, /i="1" u="sng"/);
  assert.match(xml, /<a:latin typeface="맑은 고딕"/);
  assert.match(xml, /wrap="none"/);
  assert.match(xml, /tIns="9525" rIns="19050" bIns="28575" lIns="38100"/);
  for (const back of shapes) for (const k of ['font', 'size', 'bold', 'italic', 'underline', 'pad', 'nowrap', 'align', 'valign', 'color']) assert.deepEqual(back[k], shape[k], k);
});

test('도형: 표준 XML alphaMod/alphaOff와 혼합 텍스트 조각을 다시 저장해도 유지', () => {
  const { files, xml } = saved([{ ...base, paras: [{ runs: [{ t: '강조', b: true, i: true, u: true, font: 'Arial', color: '#ff0000' }, { t: '일반', b: false, i: false, u: false, font: '맑은 고딕', color: '#000000' }] }] }]);
  files['xl/drawings/drawing1.xml'] = new TextEncoder().encode(xml.replace('<a:srgbClr val="4472C4"></a:srgbClr>', '<a:srgbClr val="4472C4"><a:alpha val="80000"/><a:alphaMod val="50000"/><a:alphaOff val="10000"/></a:srgbClr>'));
  const one = readXlsx(zip(files)).data.sheets[0].shapes[0];
  assert.equal(one.fillOpacity, 0.5);
  const two = saved([one]).shapes[0];
  assert.equal(two.fillOpacity, 0.5);
  assert.equal(two.paras[0].runs[0].i, true);
  assert.equal(two.paras[0].runs[1].i, false);
  assert.equal(two.paras[0].runs[1].u, false);
  assert.equal(two.paras[0].runs[1].b, false);
  assert.equal(two.paras[0].runs[1].font, '맑은 고딕');
});

test('차트 제목: WIXEL 확장 없이 표준 XML만으로 글자 크기·색·굵기를 유지', () => {
  const wb = make([]);
  wb.transact(() => {
    wb.setInput(0, 0, 0, '매체'); wb.setInput(0, 0, 1, '매출');
    wb.setInput(0, 1, 0, '검색'); wb.setInput(0, 1, 1, '100');
    wb.setSheetProp(0, 'charts', [{ id: 'c1', type: 'column', x: 100, y: 100, w: 480, h: 280, title: '월별 매출', titleSize: 22, titleBold: true, titleColor: '#ab1234', range: { r1: 0, c1: 0, r2: 1, c2: 1 } }]);
  });
  const files = unzip(writeXlsx(wb));
  const xml = textOf(files['xl/charts/chart1.xml']);
  assert.match(xml, /<a:defRPr sz="2200" b="1">/);
  assert.match(xml, /<a:srgbClr val="AB1234"/);
  files['xl/charts/chart1.xml'] = new TextEncoder().encode(xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g, ''));
  const back = readXlsx(zip(files)).data.sheets[0].charts[0];
  assert.equal(back.titleSize, 22);
  assert.equal(back.titleBold, true);
  assert.equal(back.titleColor, '#ab1234');
});
