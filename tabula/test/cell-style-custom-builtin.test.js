import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { importCellStyleList } from '../src/cell-style.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';
const include = { number: true, alignment: true, font: true, border: true, fill: true, protection: true };
const fixtures = () => [
  { name: 'BuiltinUnset', builtinId: 5, style: { bold: true }, include },
  { name: 'BuiltinFalse', builtinId: 6, customBuiltin: false, style: { italic: true }, include },
  { name: 'BuiltinTrue', builtinId: 8, customBuiltin: true, style: { color: '#123456' }, include },
];
const attrs = bytes => kids(child(parseXml(textOf(unzip(bytes)['xl/styles.xml'])), 'cellStyles'), 'cellStyle').map(n => n.attrs);

test('XLSX builtin styles preserve absent/false/true customBuiltin across two saves without altering definitions', () => {
  let book = new Workbook(); book.cellStyles = fixtures();
  book.setInput(0, 0, 0, '12'); book.setStyle(0, 0, 0, { bold: true, cellStyleName: 'BuiltinUnset' });
  for (let pass = 0; pass < 2; pass++) {
    const saved = writeXlsx(book), xml = attrs(saved);
    assert.equal(Object.hasOwn(xml.find(a => a.name === 'BuiltinUnset'), 'customBuiltin'), false);
    assert.equal(xml.find(a => a.name === 'BuiltinFalse').customBuiltin, '0');
    assert.equal(xml.find(a => a.name === 'BuiltinTrue').customBuiltin, '1');
    book = new Workbook(readXlsx(saved).data);
    assert.deepEqual(book.cellStyles, fixtures());
    assert.equal(book.styleAt(0, 0, 0).cellStyleName, 'BuiltinUnset'); assert.equal(book.styleAt(0, 0, 0).bold, true);
  }
});

test('style list import retains only explicit customBuiltin booleans and never invents customization', () => {
  const source = [...fixtures(), ...[null, 'true', 1].map((flag, n) => ({ name: 'Invalid' + n, builtinId: 10 + n, customBuiltin: flag, style: {} }))];
  const imported = importCellStyleList(source);
  assert.deepEqual(imported.slice(0, 3), fixtures());
  for (const def of imported.slice(3)) assert.equal(Object.hasOwn(def, 'customBuiltin'), false);
  assert.equal(Object.hasOwn(imported[0], 'customBuiltin'), false);
  imported[1].customBuiltin = true; assert.equal(source[1].customBuiltin, false);
});

test('an explicitly modified builtin stays customized through style import, XLSX and Undo/Redo', () => {
  const book = new Workbook(); book.cellStyles = fixtures();
  const modified = importCellStyleList([{ ...book.cellStyles[0], style: { bold: true, color: '#654321' }, customBuiltin: true }])[0];
  book.transact(() => book.setCellStyles([modified, ...book.cellStyles.slice(1)]));
  const check = (flag, color) => { const back = new Workbook(readXlsx(writeXlsx(book)).data); assert.equal(back.cellStyles[0].customBuiltin, flag); assert.equal(back.cellStyles[0].style.color, color); };
  check(true, '#654321'); book.undo(); check(undefined, undefined); book.redo(); check(true, '#654321');
});
