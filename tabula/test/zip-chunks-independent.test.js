import test from 'node:test';
import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';
import { zip, zipAsync } from '../src/zip.js';
import { createXmlChunks } from '../src/xml-chunks.js';

// Independent ZIP reader: do not use the production inflater, CRC or directory reader.
function oracleCrc(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function inspect(bytes, expected, methods) {
  assert.ok(bytes instanceof Uint8Array);
  const data = Buffer.from(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const end = data.length - 22;
  assert.equal(data.readUInt32LE(end), 0x06054b50);
  assert.equal(data.readUInt16LE(end + 8), expected.size);
  assert.equal(data.readUInt16LE(end + 10), expected.size);
  assert.equal(data.readUInt16LE(end + 20), 0);
  const centralStart = data.readUInt32LE(end + 16);
  assert.equal(centralStart + data.readUInt32LE(end + 12), end);
  let cd = centralStart, previousEnd = 0;
  for (const [name, original] of expected) {
    assert.equal(data.readUInt32LE(cd), 0x02014b50);
    const flags = data.readUInt16LE(cd + 8), method = data.readUInt16LE(cd + 10);
    const crc = data.readUInt32LE(cd + 16), packedSize = data.readUInt32LE(cd + 20), size = data.readUInt32LE(cd + 24);
    const nameSize = data.readUInt16LE(cd + 28), extra = data.readUInt16LE(cd + 30), comment = data.readUInt16LE(cd + 32);
    const local = data.readUInt32LE(cd + 42);
    assert.equal(data.subarray(cd + 46, cd + 46 + nameSize).toString('utf8'), name);
    assert.equal(flags & 0x800, 0x800, 'UTF-8 file name flag');
    assert.ok(methods.includes(method), 'expected compression method');
    assert.equal(local, previousEnd, 'local offsets are exact, without overlap or gaps');
    assert.equal(data.readUInt32LE(local), 0x04034b50);
    assert.equal(data.readUInt16LE(local + 6), flags);
    assert.equal(data.readUInt16LE(local + 8), method);
    assert.equal(data.readUInt32LE(local + 14), crc);
    assert.equal(data.readUInt32LE(local + 18), packedSize);
    assert.equal(data.readUInt32LE(local + 22), size);
    const localNameSize = data.readUInt16LE(local + 26), localExtra = data.readUInt16LE(local + 28);
    assert.equal(data.subarray(local + 30, local + 30 + localNameSize).toString('utf8'), name);
    const bodyStart = local + 30 + localNameSize + localExtra;
    const packed = data.subarray(bodyStart, bodyStart + packedSize);
    const actual = method === 8 ? inflateRawSync(packed) : packed;
    assert.equal(size, original.length);
    assert.deepEqual(actual, original);
    assert.equal(crc, oracleCrc(original), 'independent bitwise CRC agrees');
    previousEnd = bodyStart + packedSize;
    cd += 46 + nameSize + extra + comment;
  }
  assert.equal(previousEnd, centralStart);
  assert.equal(cd, end);
}
function fixture() {
  const backing = Uint8Array.from([98, 1, 0, 255, 7, 99]);
  const view = backing.subarray(1, 5);
  return {
    entries: {
      '한글/😀.xml': ['<r>가', ['\ud83d', '', ['\ude00']], '끝</r>'],
      'binary.bin': ['앞', view, '뒤'],
      'isolated.xml': ['\ud83d', view, '\ude00', '종료\ud83d'],
      'empty.xml': [[], '', new Uint8Array()],
    },
    expected: new Map([
      ['한글/😀.xml', Buffer.from('<r>가😀끝</r>')],
      ['binary.bin', Buffer.concat([Buffer.from('앞'), Buffer.from(view), Buffer.from('뒤')])],
      ['isolated.xml', Buffer.concat([Buffer.from('\ud83d'), Buffer.from(view), Buffer.from('\ude00종료\ud83d')])],
      ['empty.xml', Buffer.alloc(0)],
    ]),
  };
}
test('chunked sync ZIP has independent headers, CRC, UTF-8 and byte-view fidelity', () => {
  const { entries, expected } = fixture();
  inspect(zip(entries), expected, [0]);
  assert.equal(entries['한글/😀.xml'][1][0], '\ud83d', 'input chunks are not mutated');
});
test('chunked async ZIP streams compressed Unicode and uses correct original sizes', async () => {
  const text = '<row>한글😀&amp;é</row>'.repeat(600);
  const split = text.indexOf('😀') + 1;
  const entries = { 'large.xml': [text.slice(0, split), [text.slice(split)]], 'tail.xml': ['끝'] };
  const progress = [];
  const bytes = await zipAsync(entries, value => progress.push(value));
  inspect(bytes, new Map([['large.xml', Buffer.from(text)], ['tail.xml', Buffer.from('끝')]]), [0, 8]);
  assert.deepEqual(progress, [0.5, 1]);
  let supported = false; try { new CompressionStream('deflate-raw'); supported = true; } catch {}
  if (supported) { const view = Buffer.from(bytes); const central = view.readUInt32LE(view.length - 6); assert.equal(view.readUInt16LE(central + 10), 8, 'native compressed path is exercised'); }
});
test('native compression absent or unsupported preserves every chunk in stored fallback', async () => {
  const original = globalThis.CompressionStream;
  try {
    for (const implementation of [undefined, class { constructor() { throw new TypeError('format unsupported'); } }]) {
      globalThis.CompressionStream = implementation;
      const { entries, expected } = fixture();
      entries['large.xml'] = ['가😀'.repeat(2000)];
      expected.set('large.xml', Buffer.from('가😀'.repeat(2000)));
      inspect(await zipAsync(entries), expected, [0]);
    }
  } finally { globalThis.CompressionStream = original; }
});
test('runtime compression failure propagates and is not silently converted to stored data', async () => {
  const original = globalThis.CompressionStream;
  const failure = new Error('independent compression failure');
  try {
    globalThis.CompressionStream = class { constructor() { return new TransformStream({ transform() { throw failure; } }); } };
    await assert.rejects(zipAsync({ 'large.xml': ['가'.repeat(10000)] }), error => error === failure);
  } finally { globalThis.CompressionStream = original; }
});
test('progress callback cancellation stops the operation and propagates its exact error', async () => {
  const stop = new Error('stop export'); let called = 0;
  await assert.rejects(zipAsync({ 'one.xml': ['one'], 'two.xml': ['two'] }, () => { called++; throw stop; }), error => error === stop);
  assert.equal(called, 1);
});

test('bounded XML builder preserves token order even when a surrogate pair spans chunks', async () => {
  const chunks = createXmlChunks(7), tokens = ['<r>', '가😀나'.repeat(11), '', '</r>'];
  for (const token of tokens) chunks.push(token);
  const parts = chunks.finish();
  assert.equal(chunks.count, tokens.length);
  assert.ok(parts.length > 1 && parts.every(part => part.length <= 7));
  assert.equal(parts.join(''), tokens.join(''));
  const expected = new Map([['small.xml', Buffer.from(tokens.join(''))]]);
  inspect(zip({ 'small.xml': parts }), expected, [0]);
  inspect(await zipAsync({ 'small.xml': parts }), expected, [0, 8]);
});
test('consume only releases completed owned entries; default ZIP input remains reusable', async () => {
  const entries = { first: ['가😀'], second: ['마지막'] };
  const expected = new Map([['first', Buffer.from('가😀')], ['second', Buffer.from('마지막')]]);
  inspect(zip(entries), expected, [0]);
  assert.equal(Object.keys(entries).length, 2);
  const counts = [];
  inspect(await zipAsync(entries, () => counts.push(Object.keys(entries).length), { consume: true }), expected, [0, 8]);
  assert.deepEqual(counts, [1, 0]);
  assert.deepEqual(entries, {});
  const stopped = { first: ['one'], second: ['two'] }, stop = new Error('stop after one');
  await assert.rejects(zipAsync(stopped, () => { throw stop; }, { consume: true }), error => error === stop);
  assert.deepEqual(stopped, { second: ['two'] });
});
