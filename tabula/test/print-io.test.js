import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, parsePrintAreas } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const range = (r1, c1, r2, c2) => ({ r1, c1, r2, c2 });
const maxRow = 1048575, maxCol = 16383;
const restore = bytes => new Workbook(readXlsx(bytes).data);
function book() {
  const wb = new Workbook(); wb.sheets[0].name = "매출, O'Brien";
  wb.transact(() => { wb.setInput(0, 0, 0, '합성'); wb.setInput(0, 19, 5, '=40+2'); });
  return wb;
}
function rewrite(bytes, path, fn) { const files = unzip(bytes); files[path] = fn(textOf(files[path])); return zip(files); }

test('print-area parser keeps ordered unions, whole axes, quoted commas/apostrophes and reversed ranges', () => {
  assert.deepEqual(parsePrintAreas("='매출, O''Brien'!$C$9:$A$2,'매출, O''Brien'!$E:$F,'매출, O''Brien'!$10:$12", "매출, O'Brien"),
    [range(1, 0, 8, 2), range(0, 4, maxRow, 5), range(9, 0, 11, maxCol)]);
  assert.deepEqual(parsePrintAreas('a1,c3:d4,A1', 's'), [range(0, 0, 0, 0), range(2, 2, 3, 3)]);
  assert.deepEqual(parsePrintAreas('(S!A1:B2,S!D1:E2)', 's'), [range(0, 0, 1, 1), range(0, 3, 1, 4)]);
});

test('strict print-area validation never silently keeps only the valid half of a user entry', () => {
  for (const bad of ['', 'A1,invalid', 'A1,Other!B2', "'[external.xlsx]S'!A1", 'A0:B2', 'XFE1', 'A1048577', '0:2', 'A1,', "'S!A1", 'A1 B2']) {
    assert.deepEqual(parsePrintAreas(bad, 'S'), [], bad);
  }
  assert.deepEqual(parsePrintAreas('A1,Other!B2,C3', 'S', { strict: false }), [range(0, 0, 0, 0), range(2, 2, 2, 2)]);
});

test('multiple print areas and repeated title rows/columns survive standard XLSX without private metadata', () => {
  const wb = book(), areas = [range(0, 0, 19, 2), range(3, 4, 11, 5)];
  wb.sheets[0].page = { area: areas[0], areas, titleRows: [0, 1], titleCols: [0, 0], rowBreaks: [7, 15], colBreaks: [2, 5] };
  const before = wb.serialize(), bytes = writeXlsx(wb), files = unzip(bytes);
  const defined = kids(child(parseXml(textOf(files['xl/workbook.xml'])), 'definedNames'), 'definedName');
  assert.equal(defined.find(n => n.attrs.name === '_xlnm.Print_Area').text, "'매출, O''Brien'!$A$1:$C$20,'매출, O''Brien'!$E$4:$F$12");
  const s = restore(bytes).sheets[0];
  assert.deepEqual(s.page.areas, areas); assert.deepEqual(s.page.area, areas[0]);
  assert.deepEqual(s.page.titleRows, [0, 1]); assert.deepEqual(s.page.titleCols, [0, 0]);
  assert.deepEqual(s.page.rowBreaks, [7, 15]); assert.deepEqual(s.page.colBreaks, [2, 5]);
  assert.deepEqual(wb.serialize(), before, 'export does not change model or cells');
  assert.equal(restore(bytes).getValue(0, 19, 5), 42);
});

test('full-column and full-row print areas use standard compact references', () => {
  const wb = book(); wb.sheets[0].page = { areas: [range(0, 1, maxRow, 3), range(4, 0, 8, maxCol)] };
  const bytes = writeXlsx(wb), text = textOf(unzip(bytes)['xl/workbook.xml']);
  assert.ok(text.includes('!$B:$D')); assert.ok(text.includes('!$5:$9'));
  assert.deepEqual(restore(bytes).sheets[0].page.areas, wb.sheets[0].page.areas);
});

test('view mode coexists with frozen origin, active pane and scrolling selection', () => {
  for (const mode of ['normal', 'pageBreakPreview', 'pageLayout']) {
    const wb = book(), s = wb.sheets[0]; s.freeze = { rows: 3, cols: 2, top: 10, left: 4 };
    s.view = { mode, top: 30, left: 8, r: 32, c: 9, activePane: 'bottomRight' };
    const bytes = writeXlsx(wb), back = restore(bytes).sheets[0];
    assert.deepEqual(back.view, s.view); assert.deepEqual(back.freeze, s.freeze);
    assert.match(textOf(unzip(bytes)['xl/worksheets/sheet1.xml']), new RegExp(`view="${mode}"`));
  }
});

test('only the workbook-zero sheet view and manual valid breaks are imported', () => {
  const wb = book();
  const bytes = rewrite(writeXlsx(wb), 'xl/worksheets/sheet1.xml', xml => xml
    .replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, '<sheetViews><sheetView workbookViewId="1" view="pageLayout"/><sheetView workbookViewId="0" view="pageBreakPreview"/></sheetViews>')
    .replace('</worksheet>', '<rowBreaks count="5"><brk id="7" man="1"/><brk id="7" man="true"/><brk id="12" man="0"/><brk id="0" man="1"/><brk id="1048576" man="1"/></rowBreaks><colBreaks count="2"><brk id="3" man="1"/><brk id="6"/></colBreaks></worksheet>'));
  const s = restore(bytes).sheets[0];
  assert.equal(s.view.mode, 'pageBreakPreview'); assert.deepEqual(s.page.rowBreaks, [7]); assert.deepEqual(s.page.colBreaks, [3]);
  const root = parseXml(textOf(unzip(writeXlsx(restore(bytes)))['xl/worksheets/sheet1.xml']));
  assert.deepEqual(kids(child(root, 'rowBreaks'), 'brk').map(x => x.attrs), [{ id: '7', min: '0', max: '16383', man: '1' }]);
});

// Synthetic XLSB records: 7-bit record IDs and lengths, no external/user workbook data.
const u32 = (...values) => { const b = Buffer.alloc(values.length * 4); values.forEach((v, i) => b.writeUInt32LE(v, i * 4)); return b; };
const wide = text => Buffer.concat([u32(text.length), Buffer.from(text, 'utf16le')]);
const variable = value => { const out = []; do { const b = value & 127; value >>>= 7; out.push(b | (value ? 128 : 0)); } while (value); return Buffer.from(out); };
const br = (id, data = Buffer.alloc(0)) => Buffer.concat([variable(id), variable(data.length), data]);
function xlsb(mode) {
  const view = Buffer.alloc(30); view.writeUInt16LE(0x1c); view.writeInt32LE(mode, 2); view.writeInt32LE(4, 6); view.writeInt32LE(2, 10); view.writeUInt16LE(125, 18);
  return zip({
    '_rels/.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.bin"/></Relationships>',
    'xl/workbook.bin': br(156, Buffer.concat([u32(0, 1), wide('rId1'), wide('S')])),
    'xl/_rels/workbook.bin.rels': '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.bin"/></Relationships>',
    'xl/worksheets/sheet1.bin': Buffer.concat([br(137, view), br(392, u32(2)), br(396, u32(7, 0, 16383, 1, 0)), br(396, u32(12, 0, 16383, 0, 0)), br(393), br(394, u32(1)), br(396, u32(3, 0, 1048575, 1, 0)), br(395)]),
  });
}

test('XLSB binary view enum and explicit breaks use the same model and standard XLSX output', () => {
  for (const [index, mode] of ['normal', 'pageBreakPreview', 'pageLayout'].entries()) {
    const wb = restore(xlsb(index)), s = wb.sheets[0];
    assert.equal(s.view.mode, mode); assert.equal(s.view.top, 4); assert.equal(s.view.left, 2); assert.equal(s.zoom, 125);
    assert.deepEqual(s.page.rowBreaks, [7]); assert.deepEqual(s.page.colBreaks, [3]);
    const back = restore(writeXlsx(wb)).sheets[0];
    assert.equal(back.view.mode, mode); assert.deepEqual(back.page.rowBreaks, [7]); assert.deepEqual(back.page.colBreaks, [3]);
  }
});
