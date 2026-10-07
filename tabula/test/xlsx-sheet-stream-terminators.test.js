import test from 'node:test';
import assert from 'node:assert/strict';
import { sheetXmlStream } from '../src/xlsx-sheet-stream.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
const enc = new TextEncoder();
function* chunks(xml, size) { const bytes = enc.encode(xml); for (let at = 0; at < bytes.length; at += size) yield bytes.subarray(at, at + size); }

for (const prefix of ['', 'p:']) test('row closing XML whitespace and PI terminators remain literal row data: prefix=' + prefix, () => {
  const name = prefix + 'row', first = '<' + name + ' r="1"><c><![CDATA[</' + name + '>]]></c><!-- </' + name + '> --><?audit ignored </' + name + '> ?><c/></' + name + ' \r\n\t>', second = '<' + name + ' r="2"/>';
  const head = '<worksheet><cols/>', tail = '<mergeCells/></worksheet>', xml = head + '<sheetData>' + first + second + '</sheetData>' + tail;
  for (const size of [1, 2, 7, 31, 65536]) { const stream = sheetXmlStream(chunks(xml, size)); assert.deepEqual([...stream.rows()], [first, second]); assert.equal(stream.rest, head + '<sheetData/>' + tail); }
});

test('unfinished PI/closing tags fail and a wrong QName suffix is not accepted as a row close', () => {
  for (const body of ['<row><?audit </row>', '<row></row ', '<row></rowOther>', '<row><!-- </row>', '<row><![CDATA[</row>']) {
    assert.throws(() => [...sheetXmlStream(chunks('<worksheet><sheetData>' + body, 1)).rows()], /끝까지/);
  }
});

test('byte/Blob XLSX stream keeps cell values with whitespace close or PI/comment/CDATA fake row closes', async () => {
  const original = new Workbook({ sheets: [{ name: 'Synthetic', cells: { '0,0': { raw: '42' }, '1,0': { raw: '7' } } }] }), parts = unzip(writeXlsx(original)), path = 'xl/worksheets/sheet1.xml', base = textOf(parts[path]);
  for (const prefix of ['', 'x:']) for (const mode of ['space', 'pi', 'comment', 'cdata']) {
    let xml = base;
    if (prefix) xml = xml.replace(/<(\/?)(worksheet|sheetData|row|c|v|f|is|t)(?=[\s/>])/g, '<$1x:$2').replace('<x:worksheet ', '<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ');
    const close = '</' + prefix + 'row>';
    if (mode === 'space') xml = xml.replaceAll(close, '</' + prefix + 'row \t>');
    else xml = xml.replace(close, (mode === 'pi' ? '<?audit fake ' + close + ' ?>' : mode === 'comment' ? '<!-- fake ' + close + ' -->' : '<![CDATA[fake ' + close + ']]>') + close);
    const bytes = zip({ ...parts, [path]: enc.encode(xml) });
    if (prefix) assert.equal(new Workbook(readXlsx(bytes).data).getValue(0, 0, 0), 42);
    for (const input of [bytes, new Blob([bytes])]) { const w = new Workbook((await readXlsxAsync(input, null, { streamThreshold: 0 })).data); assert.equal(w.getValue(0, 0, 0), 42, prefix + ':' + mode); assert.equal(w.getValue(0, 1, 0), 7); }
  }
});
