import test from 'node:test';
import assert from 'node:assert/strict';
import { createXmlChunks } from '../src/xml-chunks.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, writeXlsxAsync, readXlsx } from '../src/xlsx.js';

test('XML builder bounds every string and preserves all text across surrogate boundaries', () => {
  const builder = createXmlChunks(7), values = ['<r>한글', '😀'.repeat(13), '\ud83d', '\ude01</r>'];
  for (const value of values) builder.push(value);
  const chunks = builder.finish();
  assert.equal(builder.count, values.length);
  assert.ok(chunks.length > 2 && chunks.every(x => x.length <= 7));
  assert.equal(chunks.join(''), values.join(''));
  assert.equal(textOf(unzip(zip({ 't.xml': ['<xml>', chunks, '</xml>'] }))['t.xml']), '<xml>' + values.join('') + '</xml>');
});

for (const async of [false, true]) test(`worksheet larger than a chunk never joins all rows and retains blanks/styles/formulas: async=${async}`, async () => {
  const wb = new Workbook({ sheets: [{ name: '검사', cells: {}, cellRuns: [
    [0, 0, 24000, { raw: '', style: { fill: '#abcdef', bold: true } }],
    [0, 1, 24000, { raw: '', style: { border: { bottom: { color: '#123456', style: 'thin' } } } }],
  ] }] });
  wb.setInput(0, 3, 2, '12'); wb.setInput(0, 3, 3, '=C4*2');
  wb.setCellData(0, 9, 2, { raw: '한글😀', comment: '메모', link: 'https://example.invalid' });
  const join = Array.prototype.join;
  Array.prototype.join = function (separator) {
    if (separator === '') {
      let length = 0;
      for (let i = 0; i < this.length; i++) {
        if (typeof this[i] !== 'string') break;
        length += this[i].length;
        if (length > (1 << 20)) throw new Error('whole XML join');
      }
    }
    return join.call(this, separator);
  };
  let bytes;
  try { bytes = async ? await writeXlsxAsync(wb) : writeXlsx(wb); }
  finally { Array.prototype.join = join; }
  const xml = textOf(unzip(bytes)['xl/worksheets/sheet1.xml']);
  assert.ok(xml.length > (1 << 20));
  assert.equal((xml.match(/<row /g) ?? []).length, 24000);
  assert.match(xml, /<c r="A24000" s="\d+"\/>/);
  assert.match(xml, /<c r="B24000" s="\d+"\/>/);
  const back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.getRaw(0, 3, 3), '=C4*2'); assert.equal(back.getValue(0, 3, 3), 24);
  assert.equal(back.styleAt(0, 23999, 0).fill, '#abcdef');
  assert.equal(back.getRaw(0, 9, 2), '한글😀'); assert.equal(back.getCell(0, 9, 2).comment, '메모');
});
