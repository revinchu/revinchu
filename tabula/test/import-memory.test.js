import test from 'node:test';
import assert from 'node:assert/strict';
import { readSharedStrings } from '../src/shared-strings.js';
import { parseXml, kids, allText } from '../src/xml.js';
import { readPhonetic } from '../src/phonetic.js';
import { unzip, zipAsync, prepareZipEntry, textOf } from '../src/zip.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';
import { convertXlsb } from '../src/xlsb.js';
const enc = new TextEncoder();
function consume(generator) { for (;;) { const step = generator.next(); if (step.done) return step.value; } }
const corpus = ['<si/>', '<si><t xml:space="preserve"> 앞 뒤 </t></si>', '<si><r><t>가😀</t></r><r><t>&amp;_x000A_나</t></r></si>', '<si><t><![CDATA[글자 </si><si> 끝]]></t></si>', '<si><!-- <si><t>무시</t></si> --><t>&lt;태그&gt;&#13;줄\r\n바꿈_x005F_x0041_</t></si>', '<si><t>東京</t><rPh sb="0" eb="2"><t>とうきょう</t></rPh><phoneticPr fontId="0" alignment="center"/></si>'];
test('shared string chunks retain Unicode, whitespace, escapes, rich runs and phonetics', () => {
  const xml = '<?xml version="1.0"?><sst xmlns="urn:test">' + corpus.join('') + '</sst>', fonts = [{ font: '맑은 고딕', size: 8 }];
  const old = kids(parseXml(xml), 'si'), expected = old.map(allText);
  for (const chunk of [64, 65, 67, 71, 97, 128, 512, 1048576]) {
    const result = consume(readSharedStrings(enc.encode(xml), fonts, chunk));
    assert.deepEqual(result.strings, expected, `chunk=${chunk}`);
    for (let i = 0; i < old.length; i++) assert.deepEqual(result.phonetics[i], readPhonetic(old[i], expected[i], fonts));
  }
});
test('prefixed shared strings and greater-than characters in attributes preserve item boundaries', () => {
  const xml = '<s:sst xmlns:s="urn:test"><s:si note="a > b"><s:t>첫째</s:t></s:si><s:si/><s:si><s:t>둘째😀</s:t></s:si><extLst><si><t>확장</t></si></extLst></s:sst>';
  const result = consume(readSharedStrings(enc.encode(xml), [], 64));
  assert.deepEqual(result.strings, ['첫째', '', '둘째😀']);
});
test('large shared string tables yield progress without an all-item XML tree', () => {
  const xml = '<sst>' + '<si><t>반복😀</t></si>'.repeat(5000) + '</sst>', iter = readSharedStrings(enc.encode(xml), [], 257);
  const progress = []; let result;
  for (;;) { const step = iter.next(); if (step.done) { result = step.value; break; } progress.push(step.value); }
  assert.deepEqual(progress, [2048, 4096]); assert.equal(result.strings.length, 5000); assert.equal(result.strings[4999], '반복😀'); assert.equal(result.phonetics.length, 0);
});
test('a truncated shared string is rejected instead of silently losing the trailing item', () => {
  assert.throws(() => consume(readSharedStrings(enc.encode('<sst><si><t>미완성'), [], 64)), /XML 항목/);
});
test('ZIP entries are prepared individually and unrelated sheet data stays compressed', async () => {
  const bytes = await zipAsync({ 'one.xml': '가'.repeat(10000), 'two.xml': '나'.repeat(10000) });
  const files = unzip(bytes);
  assert.equal(typeof Object.getOwnPropertyDescriptor(files, 'one.xml').get, 'function');
  assert.equal(typeof Object.getOwnPropertyDescriptor(files, 'two.xml').get, 'function');
  const [first, again] = await Promise.all([prepareZipEntry(files, 'one.xml'), prepareZipEntry(files, 'one.xml')]);
  assert.equal(first, again); assert.equal(textOf(first), '가'.repeat(10000));
  assert.equal(typeof Object.getOwnPropertyDescriptor(files, 'two.xml').get, 'function');
  delete files['one.xml'];
  assert.equal(textOf(await prepareZipEntry(files, 'two.xml')), '나'.repeat(10000));
});
test('an overwritten ZIP entry is not replaced by late decompression', async () => {
  const files = unzip(await zipAsync({ 'one.xml': '가'.repeat(10000) }));
  const pending = prepareZipEntry(files, 'one.xml'); const replacement = enc.encode('바뀐 내용'); files['one.xml'] = replacement;
  assert.equal(await pending, replacement); assert.equal(files['one.xml'], replacement);
});
test('sync and async XLSX paths retain values, formulas and sheet metadata after release', async () => {
  const wb = new Workbook({ sheets: [{ name: '첫째', cells: { '0,0': { raw: '한국어😀' }, '1,0': { raw: '=1+2' } }, freeze: { rows: 1, cols: 1 }, colWidths: { 1: 123 } }, { name: '둘째', cells: { '0,0': { raw: '마지막' }, '0,1': { raw: "='첫째'!A2" } }, state: 'hidden' }] });
  const bytes = await zipAsync(unzip(writeXlsx(wb))), progress = [];
  const sync = new Workbook(readXlsx(bytes).data), asyncBook = new Workbook((await readXlsxAsync(bytes, p => progress.push(p))).data);
  assert.deepEqual(asyncBook.serialize(), sync.serialize());
  assert.equal(asyncBook.getValue(0, 0, 0), '한국어😀'); assert.equal(asyncBook.getValue(1, 0, 1), 3);
  assert.ok(progress.some(p => p.msg.includes('시트 준비'))); assert.ok(progress.every(p => !('preparePart' in p)));
});
test('async import checks progress callbacks before the next sheet decompression', async () => {
  const bytes = writeXlsx(new Workbook()), cancellation = new Error('문서 전환');
  await assert.rejects(readXlsxAsync(bytes, p => { if (p.msg.includes('시트 준비')) throw cancellation; }), e => e === cancellation);
});
test('lazy XLSB conversion does not retain unvisited binary sheets', () => {
  const files = { 'xl/workbook.bin': new Uint8Array(), 'xl/worksheets/sheet1.bin': new Uint8Array(), 'xl/worksheets/sheet2.bin': new Uint8Array(), 'xl/_rels/workbook.bin.rels': enc.encode('<Relationships><Relationship Id="r1" Type="urn:worksheet" Target="worksheets/sheet1.bin"/><Relationship Id="r2" Type="urn:worksheet" Target="worksheets/sheet2.bin"/></Relationships>') };
  const original = files['xl/worksheets/sheet2.bin']; consume(convertXlsb(files, { lazySheets: true }));
  assert.equal(files.__xlsb.rows.size, 0); assert.equal(files['xl/worksheets/sheet2.bin'], original);
  files.__xlsb.prepareSheet('xl/worksheets/sheet1.bin');
  assert.match(textOf(files['xl/worksheets/sheet1.bin']), /<worksheet/); assert.equal(files['xl/worksheets/sheet2.bin'], original);
});

test('native decompression failure cancels and unlocks its reader before JS fallback', async () => {
  const bytes = await zipAsync({ 'one.xml': '가'.repeat(10000) }), files = unzip(bytes);
  const OriginalBlob = globalThis.Blob, OriginalStream = globalThis.DecompressionStream;
  let cancelled = false, released = false;
  const reader = { read: async () => ({ value: new Uint8Array(40000), done: false }), cancel: async () => { cancelled = true; }, releaseLock: () => { released = true; } };
  try {
    globalThis.Blob = class { stream() { return { pipeThrough: () => ({ getReader: () => reader }) }; } };
    globalThis.DecompressionStream = class {};
    assert.equal(textOf(await prepareZipEntry(files, 'one.xml')), '가'.repeat(10000));
    assert.equal(cancelled, true); assert.equal(released, true);
  } finally { globalThis.Blob = OriginalBlob; globalThis.DecompressionStream = OriginalStream; }
});


test('retained worksheet formulas and string caches preserve exact UTF-16 text', async () => {
  const values = ['한글😀', '앞\n뒤', '"\\끝', '_x0041_'];
  const sheets = [{ name: 'Data', cells: Object.fromEntries(values.flatMap((value, r) => [[`${r},0`, {raw:"'" + value}], [`${r},1`, {raw:'="' + value.replace(/"/g, '""') + '"'}]])) }];
  const original = new Workbook({ sheets });
  const bytes = writeXlsx(original);
  const loaded = new Workbook((await readXlsxAsync(bytes)).data);
  for (let r=0;r<values.length;r++) {
    assert.equal(loaded.getValue(0,r,0), original.getValue(0,r,0));
    assert.equal(loaded.getCell(0,r,1).raw, original.getCell(0,r,1).raw);
    assert.equal(loaded.getValue(0,r,1), original.getValue(0,r,1));
  }
  const files=unzip(bytes);
  files['xl/worksheets/sheet1.xml']=enc.encode('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>_xD800_단독_xDFFF_</t></is></c><c r="B1" t="str"><f>"_xD800_단독_xDFFF_"</f><v>_xD800_단독_xDFFF_</v></c></row></sheetData></worksheet>');
  const escaped=new Workbook((await readXlsxAsync(await zipAsync(files))).data);
  assert.equal(escaped.getValue(0,0,0),'\ud800단독\udfff');
  assert.equal(escaped.getCell(0,0,1).raw,'="\ud800단독\udfff"');
  assert.equal(escaped.getValue(0,0,1),'\ud800단독\udfff');
});
