import test from 'node:test';
import assert from 'node:assert/strict';
import { deflateRawSync, inflateRawSync, constants } from 'node:zlib';
import { zipEntryChunks, zipEntryInfo, unzip, prepareZipEntry } from '../src/zip.js';
import { inflateChunks } from '../src/inflate-chunks.js';

// ZIP headers and CRC oracle are independent of production writer/directory helpers.
function checksum(data) {
  let crc = 0xffffffff;
  for (const value of data) {
    crc ^= value;
    for (let i = 0; i < 8; i++) crc = (crc >>> 1) ^ (crc & 1 ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}
function archive(data, { method = 8, packed = deflateRawSync(data), size = data.length, crc = checksum(data) } = {}) {
  const name = Buffer.from('한글.xml'); if (method === 0) packed = data;
  const local = Buffer.alloc(30 + name.length), central = Buffer.alloc(46 + name.length), end = Buffer.alloc(22);
  local.writeUInt32LE(0x04034b50); local.writeUInt16LE(20, 4); local.writeUInt16LE(0x800, 6); local.writeUInt16LE(method, 8);
  local.writeUInt32LE(crc, 14); local.writeUInt32LE(packed.length, 18); local.writeUInt32LE(size, 22); local.writeUInt16LE(name.length, 26); name.copy(local, 30);
  central.writeUInt32LE(0x02014b50); central.writeUInt16LE(20, 4); central.writeUInt16LE(20, 6); central.writeUInt16LE(0x800, 8); central.writeUInt16LE(method, 10);
  central.writeUInt32LE(crc, 16); central.writeUInt32LE(packed.length, 20); central.writeUInt32LE(size, 24); central.writeUInt16LE(name.length, 28); name.copy(central, 46);
  end.writeUInt32LE(0x06054b50); end.writeUInt16LE(1, 8); end.writeUInt16LE(1, 10); end.writeUInt32LE(central.length, 12); end.writeUInt32LE(local.length + packed.length, 16);
  return Buffer.concat([local, packed, central, end]);
}
const entry = '한글.xml';
function collect(chunks) { return Buffer.concat(Array.from(chunks, value => Buffer.from(value))); }
class Bits {
  constructor() { this.bytes = []; this.value = 0; this.n = 0; }
  bits(value, n) { for (let i = 0; i < n; i++) { this.value |= ((value >>> i) & 1) << this.n; if (++this.n === 8) { this.bytes.push(this.value); this.value = 0; this.n = 0; } } }
  code(value, n) { for (let i = n - 1; i >= 0; i--) this.bits((value >>> i) & 1, 1); }
  literal(symbol) {
    if (symbol < 144) this.code(0x30 + symbol, 8);
    else if (symbol < 256) this.code(0x190 + symbol - 144, 9);
    else if (symbol < 280) this.code(symbol - 256, 7);
    else this.code(0xc0 + symbol - 280, 8);
  }
  finish() { if (this.n) this.bytes.push(this.value); return Buffer.from(this.bytes); }
}
function fixedPayload(write) { const bits = new Bits(); bits.bits(1, 1); bits.bits(1, 2); write(bits); bits.literal(256); return bits.finish(); }
function runPayload(repeats) { return fixedPayload(b => { b.literal(97); for (let i = 0; i < repeats; i++) { b.literal(285); b.code(0, 5); } }); }
function literalsOnlyDynamic({ eob = true } = {}) {
  const b = new Bits(); b.bits(1, 1); b.bits(2, 2); b.bits(0, 5); b.bits(0, 5); b.bits(14, 4);
  const order = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1];
  for (const symbol of order) b.bits(symbol === 0 || symbol === 1 ? 1 : 0, 3);
  for (let i = 0; i < 258; i++) b.bits(i === 65 || (eob && i === 256) ? 1 : 0, 1);
  b.bits(0, 1); b.bits(1, 1); return b.finish();
}

test('ZIP chunks match native zlib for stored, fixed and dynamic Huffman blocks', () => {
  const source = Buffer.from('가나다😀\r\n<&>반복 0123456789'.repeat(12000));
  for (const options of [{ level: 0 }, { strategy: constants.Z_FIXED, level: 9 }, { level: 9 }]) {
    const packed = deflateRawSync(source, options), bytes = archive(source, { packed });
    const files = unzip(bytes), getter = Object.getOwnPropertyDescriptor(files, entry).get;
    assert.ok(getter); assert.deepEqual(collect(zipEntryChunks(files, entry)), inflateRawSync(packed));
    assert.equal(Object.getOwnPropertyDescriptor(files, entry).get, getter, 'streaming does not cache a full expanded value');
    assert.deepEqual(zipEntryInfo(files, entry), { size: source.length, compressedSize: packed.length, method: 8, crc: checksum(source) });
  }
  const dynamic = deflateRawSync(source, { level: 9 }); assert.equal((dynamic[0] >>> 1) & 3, 2, 'independent vector exercises dynamic tables');
});

test('ZIP stored entries and zero-length entries preserve values and bounded chunk ownership', () => {
  for (const method of [0, 8]) for (const source of [Buffer.alloc(0), Buffer.from('é한글😀'.repeat(300))]) {
    const files = unzip(archive(source, { method })), before = method === 0 ? files[entry].slice() : null;
    const chunks = Array.from(zipEntryChunks(files, entry, { chunkSize: 17 }));
    assert.deepEqual(collect(chunks), source); assert.ok(chunks.every(chunk => chunk.byteLength <= 17 && chunk.buffer.byteLength <= 17));
    assert.equal(new Set(chunks.map(chunk => chunk.buffer)).size, chunks.length);
    if (before && chunks.length) { chunks[0].fill(0); assert.deepEqual(files[entry], before, 'stored chunks do not alias the archive'); }
  }
});

test('chunk boundaries preserve UTF-8 XML, distance-one overlapping matches and retained previous chunks', () => {
  const source = Buffer.from('<s:worksheet>한글😀&amp;</s:worksheet>'.repeat(30));
  for (const chunkSize of [1, 2, 3, 7, 64, 65536]) {
    const files = unzip(archive(source)), decoder = new TextDecoder(); let text = '';
    for (const chunk of zipEntryChunks(files, entry, { chunkSize })) text += decoder.decode(chunk, { stream: true });
    text += decoder.decode(); assert.equal(text, source.toString('utf8'));
  }
  const packed = runPayload(1200), original = inflateRawSync(packed);
  const chunks = inflateChunks(packed, { chunkSize: 7 }); const first = chunks.next().value;
  const retained = first.slice(); const second = chunks.next().value;
  assert.deepEqual(first, retained); assert.notEqual(first.buffer, second.buffer);
  first.fill(42);
  const remainder = collect(chunks); assert.ok(remainder.every(value => value === 97), 'mutating a yielded chunk never changes the history window');
  assert.equal(7 + 7 + remainder.length, original.length);
});

test('the maximum 32768-byte distance remains correct across output and history-window boundaries', () => {
  const original = Buffer.alloc(32768 + 258 + 258);
  for (let i = 0; i < 32768; i++) original[i] = (i * 131 + (i >>> 7)) & 255;
  original.copy(original, 32768, 0, 258); original.copy(original, 32768 + 258, 258, 516);
  const packed = fixedPayload(b => {
    for (let i = 0; i < 32768; i++) b.literal(original[i]);
    for (let i = 0; i < 2; i++) { b.literal(285); b.code(29, 5); b.bits(8191, 13); }
  });
  assert.deepEqual(inflateRawSync(packed), original, 'native oracle validates the crafted maximum-distance vector');
  for (const chunkSize of [31, 32767, 65536]) assert.deepEqual(collect(inflateChunks(packed, { chunkSize })), original);
});

test('dynamic literal-only blocks permit an empty distance table but require an end code', () => {
  const valid = literalsOnlyDynamic(); assert.deepEqual(inflateRawSync(valid), Buffer.from('A'));
  assert.deepEqual(collect(inflateChunks(valid)), Buffer.from('A'));
  assert.throws(() => collect(inflateChunks(literalsOnlyDynamic({ eob: false }))), /종료 코드/);
});

test('full expansion is never allocated, even for high-ratio multi-megabyte output', () => {
  const packed = runPayload(66000), size = 1 + 258 * 66000, Native = globalThis.Uint8Array;
  const sizes = []; let total = 0, maxChunk = 0;
  globalThis.Uint8Array = new Proxy(Native, { construct(Target, args) {
    if (typeof args[0] === 'number') { sizes.push(args[0]); assert.ok(args[0] <= 65536, 'no output-size allocation is permitted'); }
    return Reflect.construct(Target, args);
  } });
  try {
    for (const chunk of inflateChunks(packed, { sizeLimit: size })) {
      total += chunk.length; maxChunk = Math.max(maxChunk, chunk.length);
      assert.equal(chunk[0], 97); assert.equal(chunk[chunk.length - 1], 97);
    }
  } finally { globalThis.Uint8Array = Native; }
  assert.equal(total, size); assert.equal(maxChunk, 65536); assert.ok(sizes.includes(32768)); assert.ok(sizes.includes(65536));
});

test('consuming twice or returning early keeps lazy entries reusable and the input unchanged', () => {
  const source = Buffer.from('bounded output '.repeat(20000)), bytes = archive(source), untouched = bytes.slice(), files = unzip(bytes);
  const getter = Object.getOwnPropertyDescriptor(files, entry).get, stream = zipEntryChunks(files, entry);
  assert.equal(stream.next().value.length, 65536); assert.equal(stream.return().done, true);
  assert.equal(Object.getOwnPropertyDescriptor(files, entry).get, getter);
  assert.deepEqual(collect(zipEntryChunks(files, entry)), source); assert.deepEqual(collect(zipEntryChunks(files, entry)), source);
  assert.equal(Object.getOwnPropertyDescriptor(files, entry).get, getter); assert.deepEqual(bytes, untouched);
  const info = zipEntryInfo(files, entry); info.size = 0; assert.equal(zipEntryInfo(files, entry).size, source.length, 'metadata callers cannot mutate reader state');
});

test('prepared full entries and plain byte maps preserve the new streaming contract', async () => {
  const source = Buffer.from('previous API is preserved '.repeat(3000)), files = unzip(archive(source));
  await prepareZipEntry(files, entry); const value = files[entry];
  assert.deepEqual(collect(zipEntryChunks(files, entry)), source); assert.equal(files[entry], value);
  assert.deepEqual(collect(zipEntryChunks({ [entry]: source }, entry, { chunkSize: 1024 })), source);
  assert.equal(zipEntryInfo(files, 'missing'), null); delete files[entry]; assert.equal(zipEntryInfo(files, entry), null);
});

test('size mismatch and CRC failure are rejected without caching a partial expanded entry', () => {
  const source = Buffer.from('CRC and size '.repeat(12000));
  for (const method of [0, 8]) for (const options of [{ size: source.length - 1 }, { size: source.length + 1 }, { crc: checksum(source) ^ 1 }]) {
    const files = unzip(archive(source, { ...options, method })), before = Object.getOwnPropertyDescriptor(files, entry);
    assert.throws(() => collect(zipEntryChunks(files, entry, { chunkSize: 1024 })), /크기|CRC/);
    const after = Object.getOwnPropertyDescriptor(files, entry); assert.equal(after.get, before.get); assert.equal(after.value, before.value);
  }
});

test('stored blocks check inverse lengths and payload truncation; deflate checks trailing payload', () => {
  for (const packed of [Buffer.from([1, 1, 0, 0, 0, 65]), Buffer.from([1, 3, 0, 252, 255, 65])]) assert.throws(() => collect(inflateChunks(packed)), /길이|잘렸/);
  const source = Buffer.from('complete raw stream'), packed = deflateRawSync(source);
  assert.throws(() => collect(inflateChunks(packed.subarray(0, packed.length - 1))), /잘렸/);
  assert.throws(() => collect(inflateChunks(Buffer.concat([packed, Buffer.from([0])]))), /불필요/);
});

test('reserved block, literal and distance codes and references before the window are rejected', () => {
  const invalid = [Buffer.from([7]), fixedPayload(b => b.literal(286)), fixedPayload(b => { b.literal(65); b.literal(257); b.code(30, 5); }), fixedPayload(b => { b.literal(257); b.code(0, 5); })];
  for (const bytes of invalid) assert.throws(() => collect(inflateChunks(bytes)), /형식|길이|거리/);
});

test('invalid dynamic code tables fail before output is exposed', () => {
  for (const counts of [[1, 1, 1, 0], [0, 0, 0, 2]]) {
    const b = new Bits(); b.bits(1, 1); b.bits(2, 2); b.bits(0, 5); b.bits(0, 5); b.bits(0, 4);
    for (const count of counts) b.bits(count, 3);
    assert.throws(() => collect(inflateChunks(b.finish())), /허프만/);
  }
});

test('ZIP streaming rejects missing/non-byte entries and invalid chunk sizes without invoking custom getters', () => {
  const source = Buffer.from('checked'), files = unzip(archive(source));
  for (const chunkSize of [0, -1, 0.5, NaN, Infinity, (1 << 20) + 1]) assert.throws(() => collect(zipEntryChunks(files, entry, { chunkSize })), RangeError);
  assert.throws(() => collect(zipEntryChunks(files, 'missing')), /찾을/);
  assert.throws(() => collect(zipEntryChunks({ [entry]: 'text' }, entry)), TypeError);
  const custom = {}; Object.defineProperty(custom, entry, { get() { throw Error('must not run'); } });
  assert.throws(() => collect(zipEntryChunks(custom, entry)), TypeError);
});


test('mixed nonfinal stored/fixed blocks preserve prefetched bytes and the history window', () => {
  const match = fixedPayload(b => { b.literal(257); b.code(0, 5); });
  const one = Buffer.concat([Buffer.from([0, 3, 0, 252, 255, 65, 65, 65]), match]);
  const b = new Bits(); b.bits(0, 1); b.bits(1, 2); b.literal(65); b.literal(256); b.bits(1, 1); b.bits(0, 2);
  while (b.n) b.bits(0, 1); b.bits(3, 16); b.bits(0xfffc, 16); b.bits(66, 8); b.bits(67, 8); b.bits(68, 8);
  for (const packed of [one, b.finish()]) assert.deepEqual(collect(inflateChunks(packed, { chunkSize: 2 })), inflateRawSync(packed));
});

test('native-zlib vectors with short final codes, all strategies and truncated final input match the oracle', () => {
  let seed = 0x1290ab;
  const random = () => { seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed; };
  for (let i = 0; i < 90; i++) {
    const source = Buffer.alloc(i < 8 ? i : random() % 6000);
    for (let j = 0; j < source.length; j++) source[j] = i % 3 === 0 ? random() & 255 : i % 3 === 1 ? j % 17 : 65;
    const strategy = [constants.Z_DEFAULT_STRATEGY, constants.Z_FIXED, constants.Z_HUFFMAN_ONLY, constants.Z_RLE][i % 4];
    const packed = deflateRawSync(source, { level: i % 10, strategy });
    assert.deepEqual(collect(inflateChunks(packed, { chunkSize: 31 })), source);
    assert.throws(() => collect(inflateChunks(packed.subarray(0, packed.length - 1))), /잘렸|허프만|길이/);
    assert.throws(() => collect(inflateChunks(Buffer.concat([packed, Buffer.from([0])]))), /불필요/);
  }
});
