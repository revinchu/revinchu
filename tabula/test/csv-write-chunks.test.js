import test from 'node:test';
import assert from 'node:assert/strict';
import { toDelimited, toDelimitedChunks, parseDelimited, CsvStream } from '../src/csv.js';

function decode(chunks) { const decoder = new TextDecoder(); let text = ''; for (const part of chunks) text += decoder.decode(part, { stream: true }); return text + decoder.decode(); }

test('lazy CSV/TSV byte output matches legacy quoting and row separators', () => {
  const rows = [['한글😀', '쉼표,', '"따옴표"', '탭\t', '줄\n바꿈', 'CR\rLF\r\n', ''], ['', '끝']];
  for (const delim of [',', '\t', ';']) for (const eol of ['\r\n', '\n']) {
    assert.equal(decode(toDelimitedChunks(rows, delim, eol)), toDelimited(rows, delim, eol));
    const BOM = [...toDelimitedChunks(rows, delim, eol, { bom: true })];
    assert.deepEqual([...BOM[0].subarray(0, 3)], [239, 187, 191]);
    assert.equal(new TextDecoder('utf-8', { ignoreBOM: true }).decode(Buffer.concat(BOM)), '\ufeff' + toDelimited(rows, delim, eol));
  }
  assert.equal(toDelimited([['a\rb']]), '"a\rb"');
});

test('empty row and file output preserve exact legacy separators and optional BOM', () => {
  for (const rows of [[], [[]], [[], []], [['']], [['', '']], [[], ['a'], []]]) assert.equal(decode(toDelimitedChunks(rows)), toDelimited(rows));
  assert.equal([...toDelimitedChunks([])].length, 0);
  assert.deepEqual([...toDelimitedChunks([], ',', '\r\n', { bom: true })][0], new Uint8Array([239, 187, 191]));
});

test('CSV coalesces tiny fields yet leaves the remaining lazy source unconsumed', () => {
  let produced = 0, closed = false;
  function* rows() { try { for (let i = 0; i < 1000000; i++) { produced++; yield ['row', String(i)]; } } finally { closed = true; } }
  const chunks = toDelimitedChunks(rows()); const first = chunks.next();
  assert.equal(first.value.byteLength, 65536); assert.ok(produced < 10000);
  chunks.return(); assert.equal(closed, true);
});

test('very large quoted fields and many columns remain bounded and exactly escaped', () => {
  const rows = [['"한글😀,\r\n'.repeat(100000), 'tail'], Array.from({ length: 40000 }, (_, i) => '열' + i)];
  const chunks = [...toDelimitedChunks(rows)];
  assert.ok(chunks.length > 10);
  for (const part of chunks) { assert.ok(part.length <= 256 * 1024); assert.ok(new TextDecoder().decode(part).length <= 65536); }
  assert.equal(decode(chunks), toDelimited(rows));
});

test('UTF-16 surrogate pairs split at escape and encoding boundaries remain UTF-8 intact', () => {
  for (const prefix of [32767, 32768, 65534, 65535, 65536, 131070]) {
    const rows = [['a'.repeat(prefix) + '😀끝', '\ud800홀로', '"' + 'b'.repeat(prefix) + '😀"']];
    const expected = new TextEncoder().encode(toDelimited(rows));
    const output = [...toDelimitedChunks(rows)];
    assert.deepEqual(Buffer.concat(output), Buffer.from(expected));
    const decoded = decode(output); assert.ok(decoded.includes('😀끝')); assert.ok(decoded.includes('😀""'));
  }
});

test('streamed CSV import sees the same normalized cell values as ordinary parsing', () => {
  const rows = [['A', 'B'], ['한글😀', 'a,b'], ['"', 'CR\r\nLF\n'], ['zero', '0']];
  const parsed = [], parser = new CsvStream(',', row => parsed.push(row)), decoder = new TextDecoder();
  for (const part of toDelimitedChunks(rows, ',', '\r\n', { bom: true })) parser.push(decoder.decode(part, { stream: true }));
  parser.push(decoder.decode()); parser.end();
  assert.deepEqual(parsed, parseDelimited(toDelimited(rows)));
});
