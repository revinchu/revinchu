import { test } from 'node:test';
import assert from 'node:assert/strict';
import { styleForCode, fileCode, formatCode } from '../src/format.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxBlobAsync } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';

const codes = ['0.00_ ', '0.00* ', '0.00\\ ', ' 0.00', '0.00 ', ' 0.00 ', '0.00;[Red](0.00)_ ', '0"끝 공백 "'];

test('Excel format operators keep their whitespace operands during import', () => {
  for (const code of codes) {
    const style = styleForCode(code);
    assert.equal(style.numFmt, 'custom', code);
    assert.equal(style.code, code, code);
    assert.equal(fileCode(code), code, code);
  }
  for (const code of ['General', ' General ', ' G/표준 ', '', '   ']) assert.equal(styleForCode(code).numFmt, undefined);
});

test('legacy format repair only completes a missing operator operand outside literals', () => {
  for (const code of ['0.00_', '0.00*', '0.00\\', '0.00;[Red]0_', '0"_"_']) assert.equal(fileCode(code), code + ' ');
  for (const code of [null, undefined, '']) assert.equal(fileCode(code), code);
  for (const code of ['0\\_', '0\\*', '0\\\\', '0"_"', '0"*"', '0"\\"', '0"unfinished_', '[unfinished_', '0_"', '0_[', '0.00_ ', '0.00* ', '0.00\\ ']) assert.equal(fileCode(code), code, code);
  assert.equal(fileCode('[빨강]0_'), '[Red]0_ ');
  assert.equal(fileCode('0"[빨강]"_'), '0"[빨강]"_ ');
});

for (const blob of [false, true]) test(`format whitespace survives XLSX import and re-export (${blob ? 'Blob' : 'sync'})`, async () => {
  const cells = Object.fromEntries(codes.map((code, r) => [`${r},0`, { raw: '12.5', style: { numFmt: 'custom', code } }]));
  const source = new Workbook({ sheets: [{ name: 'Formats', cells }] });
  const imported = new Workbook(readXlsx(writeXlsx(source)).data);
  for (let r = 0; r < codes.length; r++) assert.equal(imported.sheets[0].cells.get(`${r},0`).style.code, codes[r]);
  const bytes = blob ? new Uint8Array(await (await writeXlsxBlobAsync(imported)).arrayBuffer()) : writeXlsx(imported);
  const formats = kids(child(parseXml(textOf(unzip(bytes)['xl/styles.xml'])), 'numFmts'), 'numFmt').map(f => f.attrs.formatCode);
  for (const code of codes) assert.ok(formats.includes(code), code);
  const reread = new Workbook(readXlsx(bytes).data);
  for (let r = 0; r < codes.length; r++) {
    const c = reread.sheets[0].cells.get(`${r},0`);
    assert.equal(c.style.code, codes[r]);
    assert.equal(reread.getValue(0, r, 0), 12.5);
    assert.deepEqual(formatCode(12.5, c.style.code), formatCode(12.5, codes[r]));
  }
});

test('legacy WIXEL dangling-space format exports as a complete Excel format', async () => {
  const wb = new Workbook({ sheets: [{ name: 'Legacy', cells: { '0,0': { raw: '12.5', style: { numFmt: 'custom', code: '0.00_' } } } }] });
  const bytes = new Uint8Array(await (await writeXlsxBlobAsync(wb)).arrayBuffer());
  assert.match(textOf(unzip(bytes)['xl/styles.xml']), /formatCode="0\.00_ "/);
  assert.equal(wb.sheets[0].cells.get('0,0').style.code, '0.00_', 'export must not modify the source workbook');
  assert.equal(new Workbook(readXlsx(bytes).data).sheets[0].cells.get('0,0').style.code, '0.00_ ');
});
