import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';

const drawingPath = 'xl/drawings/drawing1.xml';
const base = { id: 'shape', kind: 'rect', x: 20, y: 30, w: 220, h: 130, fill: '#4472c4', stroke: '#123456', text: '인쇄·보호 합성 도형' };
const make = shapes => {
  const wb = new Workbook(); wb.transact(() => wb.setSheetProp(0, 'shapes', shapes));
  const files = unzip(writeXlsx(wb));
  // WIXEL 전용 확장으로 왕복이 우연히 통과하지 않도록 모든 extLst를 제거한다.
  for (const name of Object.keys(files)) if (name.endsWith('.xml')) {
    files[name] = new TextEncoder().encode(textOf(files[name]).replace(/<(?:\w+:)?extLst\b[^>]*>[\s\S]*?<\/(?:\w+:)?extLst>/g, ''));
  }
  return files;
};
const readShapes = files => readXlsx(zip(files)).data.sheets[0].shapes;

test('도형 인쇄·잠금은 전용 확장 없이 표준 clientData의 명시적 true/false를 왕복한다', () => {
  const source = [
    { ...base, noPrint: true, locked: false },
    { ...base, id: 'defaults', noPrint: false, locked: true },
    { ...base, id: 'unlocked', noPrint: false, locked: false },
    { ...base, id: 'hidden-print', noPrint: true, locked: true },
  ];
  const files = make(source), xml = textOf(files[drawingPath]);
  assert.doesNotMatch(xml, /extLst|wx:|wixel/i);
  assert.match(xml, /<xdr:clientData fPrintsWithSheet="0" fLocksWithSheet="0"\/>/);
  assert.match(xml, /<xdr:clientData fPrintsWithSheet="1" fLocksWithSheet="1"\/>/);
  const back = readShapes(files);
  for (let i = 0; i < source.length; i++) {
    assert.equal(back[i].noPrint, source[i].noPrint); assert.equal(back[i].locked, source[i].locked);
    assert.equal(back[i].text, source[i].text); assert.equal(back[i].fill, source[i].fill);
  }
});

test('clientData가 생략한 기본 인쇄·잠금 상태는 기존 모델 기본값을 유지한다', () => {
  const files = make([{ ...base }]);
  assert.match(textOf(files[drawingPath]), /<xdr:clientData\/>/);
  let [shape] = readShapes(files);
  assert.equal(shape.noPrint, undefined); assert.equal(shape.locked, undefined);
  files[drawingPath] = new TextEncoder().encode(textOf(files[drawingPath]).replace('<xdr:clientData/>', ''));
  [shape] = readShapes(files);
  assert.equal(shape.noPrint, undefined); assert.equal(shape.locked, undefined);
});

test('외부 도형의 true/false 문자열과 oneCell/absolute anchor도 인쇄·잠금을 복원한다', () => {
  for (const anchor of ['oneCellAnchor', 'absoluteAnchor']) for (const flag of ['false', 'true']) {
    const files = make([{ ...base }]);
    let xml = textOf(files[drawingPath]);
    xml = xml.replace('<xdr:twoCellAnchor editAs="twoCell">', `<xdr:${anchor}>`).replace('</xdr:twoCellAnchor>', `</xdr:${anchor}>`)
      .replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, '<xdr:ext cx="2095500" cy="1238250"/>')
      .replace('<xdr:clientData/>', `<xdr:clientData fPrintsWithSheet="${flag}" fLocksWithSheet="${flag}"/>`);
    if (anchor === 'absoluteAnchor') xml = xml.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, '<xdr:pos x="190500" y="285750"/>');
    files[drawingPath] = new TextEncoder().encode(xml);
    const [shape] = readShapes(files);
    assert.equal(shape.noPrint, flag === 'false'); assert.equal(shape.locked, flag === 'true'); assert.equal(shape.placement, anchor === 'oneCellAnchor' ? 'oneCell' : 'absolute');
    const saved = make([shape]), [again] = readShapes(saved);
    assert.equal(again.noPrint, flag === 'false'); assert.equal(again.locked, flag === 'true');
  }
});
