import { test } from 'node:test';
import assert from 'node:assert/strict';
import { convertXlsb } from '../src/xlsb.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx, writeXlsxBlobAsync } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const u32 = (...xs) => { const b = Buffer.alloc(xs.length * 4); xs.forEach((x, i) => b.writeUInt32LE(x >>> 0, i * 4)); return b; };
const u16 = x => { const b = Buffer.alloc(2); b.writeUInt16LE(x); return b; };
const wide = s => Buffer.concat([u32(s.length), Buffer.from(s, 'utf16le')]);
const variable = x => { const a = []; do { const b = x & 127; x >>>= 7; a.push(b | (x ? 128 : 0)); } while (x); return Buffer.from(a); };
const rec = (id, b = Buffer.alloc(0)) => Buffer.concat([variable(id), variable(b.length), b]);
// MS-XLSB 2.4.283: flags (2 bytes), ctse (4 bytes), XLNullableWideString.
const definitions = [{ name: 'PivotStyleMedium8 2', flags: 2 }, { name: 'Table custom', flags: 4 }, { name: '양쪽 서식', flags: 6 }, { name: '슬라이서 보조', flags: 0 }, { name: 'P', flags: 2 }];
function convertedStyles() {
  const chunks = [rec(508, Buffer.concat([u32(definitions.length), wide('TableStyleMedium2'), wide('PivotStyleLight16')]))];
  for (const { name, flags } of definitions) chunks.push(rec(510, Buffer.concat([u16(flags), u32(1), wide(name)])), rec(512, u32(0, 1, 0)), rec(511));
  const files = { 'xl/workbook.bin': new Uint8Array(), 'xl/styles.bin': Buffer.concat(chunks),
    'xl/_rels/workbook.bin.rels': Buffer.from('<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.bin"/></Relationships>') };
  for (const _ of convertXlsb(files)) {}
  return textOf(files['xl/styles.bin']);
}

test('XLSB custom table style names start at the documented byte offset and preserve target flags', () => {
  const styles = kids(child(parseXml(convertedStyles()), 'tableStyles'), 'tableStyle');
  assert.equal(styles.length, definitions.length);
  styles.forEach((style, i) => {
    const { name, flags } = definitions[i];
    assert.equal(style.attrs.name, name);
    assert.equal(style.attrs.pivot, flags & 2 ? '1' : '0');
    assert.equal(style.attrs.table, flags & 4 ? '1' : '0');
    assert.equal(style.attrs.count, '1');
    assert.equal(child(style, 'tableStyleElement').attrs.dxfId, '0');
  });
});

for (const asynchronous of [false, true]) test(`imported XLSB custom pivot style resolves by its full name after XLSX save (${asynchronous ? 'async' : 'sync'})`, async () => {
  const wb = new Workbook({ sheets: [{ name: 'Data', cells: { '0,0': { raw: 'Group' }, '0,1': { raw: 'Value' }, '1,0': { raw: 'A' }, '1,1': { raw: '4' } } }, { name: 'Report', cells: {}, pivot: { name: 'Pivot1', source: 'Data', range: { r1: 0, c1: 0, r2: 1, c2: 1 }, rows: ['Group'], values: [{ field: 'Value', agg: 'sum' }], top: 0, left: 0, style: definitions[0].name } }] });
  wb.setObjectStyles({ tables: [{ name: definitions[0].name, pivot: true, table: false, elements: [{ type: 'wholeTable', style: { fill: '#123ABC' } }] }] });
  const files = unzip(writeXlsx(wb));
  const tableStyles = convertedStyles().match(/<tableStyles\b[^>]*>[\s\S]*?<\/tableStyles>/)[0];
  files['xl/styles.xml'] = textOf(files['xl/styles.xml']).replace(/<tableStyles\b[^>]*>[\s\S]*?<\/tableStyles>/, tableStyles);
  const bytes = zip(files), data = asynchronous ? (await readXlsxAsync(bytes)).data : readXlsx(bytes).data;
  const imported = new Workbook(data), out = asynchronous ? new Uint8Array(await (await writeXlsxBlobAsync(imported)).arrayBuffer()) : writeXlsx(imported);
  const back = readXlsx(out).data;
  const pivot = back.sheets[1].pivot, style = back.objectStyles.tables.find(s => s.name === pivot.style);
  assert.equal(pivot.style, definitions[0].name);
  assert.ok(style, 'pivot must resolve its named custom style');
  assert.equal(style.pivot, true); assert.equal(style.table, false);
  assert.equal(style.elements[0].style.fill.toLowerCase(), '#123abc');
  for (const { name, flags } of definitions) {
    const def = back.objectStyles.tables.find(s => s.name === name);
    assert.ok(def, name); assert.equal(def.pivot, !!(flags & 2)); assert.equal(def.table, !!(flags & 4));
  }
});
