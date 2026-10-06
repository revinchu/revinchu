import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';

const EMU = 9525;
const close = (actual, expected, emus = 1) => assert.ok(Math.abs(actual - expected) <= emus / EMU + 1e-8, `${actual} ≈ ${expected}`);
const source = { kind: 'table', table: 'Data', column: 'Region' };
function fixture(slicers) {
  const wb = new Workbook();
  wb.transact(() => {
    [['Region', 'Value'], ['A', 1], ['B', 2]].forEach((row, r) => row.forEach((value, c) => wb.setInput(0, r, c, String(value))));
    wb.setSheetProp(0, 'tables', [{ id: 't1', name: 'Data', r1: 0, c1: 0, r2: 2, c2: 1, header: true }]);
    wb.setSheetProp(0, 'slicers', slicers);
  });
  return wb;
}
const slicer = (patch = {}) => ({ id: 's1', caption: '지역', source, x: 328.125, y: 29.375, w: 184.375, h: 171.125, ...patch });
function assertBox(actual, expected, emus = 1) {
  for (const key of ['x', 'y', 'w', 'h']) close(actual[key], expected[key], emus);
}

test('explicit OOXML widths differing by one pixel retain cumulative slicer cell anchors without customWidth', () => {
  const original = slicer({ x: 638.125 }), wb = fixture([original]);
  wb.sheets[0].defColW = 64;
  for (let c = 0; c < 10; c++) wb.sheets[0].colWidths[c] = 63;
  const files = unzip(writeXlsx(wb));
  files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml']).replaceAll(' customWidth="1"', '');
  const imported = new Workbook(readXlsx(zip(files)).data);
  for (let c = 0; c < 10; c++) assert.equal(imported.sheets[0].colWidths[c], 63, `column ${c}`);
  assertBox(imported.sheets[0].slicers[0], original);
  assertBox(readXlsx(writeXlsx(imported)).data.sheets[0].slicers[0], original);
});

test('a slicer group inside AlternateContent resolves every child against group coordinates', () => {
  const originals = [slicer({ objectGroup: 'g', x: 300.125, y: 10.375, w: 180.25, h: 200.5 }), slicer({ id: 's2', caption: '두 번째', objectGroup: 'g', x: 510.375, y: 35.625, w: 210.5, h: 175.25 })];
  const files = unzip(writeXlsx(fixture(originals)));
  files['xl/drawings/drawing1.xml'] = textOf(files['xl/drawings/drawing1.xml'])
    .replace('<xdr:grpSp>', '<mc:AlternateContent><mc:Choice Requires="x14" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><xdr:grpSp>')
    .replace('</xdr:grpSp>', '</xdr:grpSp></mc:Choice><mc:Fallback/></mc:AlternateContent>');
  const imported = new Workbook(readXlsx(zip(files)).data), actual = imported.sheets[0].slicers;
  assert.equal(actual.length, 2);
  assert.ok(actual[0].objectGroup); assert.equal(actual[0].objectGroup, actual[1].objectGroup);
  for (let i = 0; i < originals.length; i++) assertBox(actual[i], originals[i], 2);
  const roundtrip = readXlsx(writeXlsx(imported)).data.sheets[0].slicers;
  for (let i = 0; i < originals.length; i++) assertBox(roundtrip[i], originals[i], 2);
});

for (const placement of ['twoCell', 'oneCell', 'absolute']) test(`slicer ${placement} anchors retain exact EMU geometry through XLSX roundtrip with custom and hidden cells`, () => {
  const original = slicer({ placement }), wb = fixture([original]);
  wb.sheets[0].defColW = 64; wb.sheets[0].defRowH = 20;
  wb.sheets[0].colWidths = { 0: 37, 1: 143, 3: 81 };
  wb.sheets[0].rowHeights = { 0: 25, 2: 61, 3: 17 };
  wb.sheets[0].hiddenCols = { 2: true }; wb.sheets[0].hiddenRows = { 1: true };
  const files = unzip(writeXlsx(wb));
  let drawing = textOf(files['xl/drawings/drawing1.xml']);
  if (placement === 'oneCell') drawing = drawing.replace('<xdr:twoCellAnchor editAs="oneCell">', '<xdr:oneCellAnchor>')
    .replace(/<xdr:to>[\s\S]*?<\/xdr:to>/, `<xdr:ext cx="${Math.round(original.w * EMU)}" cy="${Math.round(original.h * EMU)}"/>`)
    .replace('</xdr:twoCellAnchor>', '</xdr:oneCellAnchor>');
  if (placement === 'absolute') drawing = drawing.replace('<xdr:twoCellAnchor editAs="absolute">', '<xdr:absoluteAnchor>')
    .replace(/<xdr:from>[\s\S]*?<\/xdr:to>/, `<xdr:pos x="${Math.round(original.x * EMU)}" y="${Math.round(original.y * EMU)}"/><xdr:ext cx="${Math.round(original.w * EMU)}" cy="${Math.round(original.h * EMU)}"/>`)
    .replace('</xdr:twoCellAnchor>', '</xdr:absoluteAnchor>');
  files['xl/drawings/drawing1.xml'] = drawing;
  const imported = new Workbook(readXlsx(zip(files)).data), actual = imported.sheets[0].slicers[0];
  assertBox(actual, original); assert.equal(actual.placement ?? 'oneCell', placement);
  const roundtrip = readXlsx(writeXlsx(imported)).data.sheets[0].slicers[0];
  assertBox(roundtrip, original); assert.equal(roundtrip.placement ?? 'oneCell', placement);
});
