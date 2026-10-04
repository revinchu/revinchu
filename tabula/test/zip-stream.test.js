import test from 'node:test';
import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';
import { createZipStreamWriter } from '../src/zip-stream.js';
import { zipCentralRecord, zipEndRecords, zipDescriptorRecord } from '../src/zip-records.js';
import { unzip, textOf } from '../src/zip.js';

// Bitwise oracle differs from the table-based production CRC and works on Node 18.
function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) { crc ^= byte; for (let bit = 0; bit < 8; bit++) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0); }
  return (crc ^ 0xffffffff) >>> 0;
}

// Independent Buffer + native zlib oracle, not the production ZIP reader.
function inspect(chunks, expected) {
  const data = Buffer.concat(chunks.map(chunk => Buffer.from(chunk))), end = data.length - 22;
  assert.equal(data.readUInt32LE(end), 0x06054b50);
  let count = data.readUInt16LE(end + 10), start = data.readUInt32LE(end + 16), dirSize = data.readUInt32LE(end + 12);
  if (count === 0xffff) {
    assert.equal(data.readUInt32LE(end - 20), 0x07064b50);
    const p = Number(data.readBigUInt64LE(end - 12)); assert.equal(data.readUInt32LE(p), 0x06064b50);
    count = Number(data.readBigUInt64LE(p + 32)); start = Number(data.readBigUInt64LE(p + 48)); dirSize = Number(data.readBigUInt64LE(p + 40));
  }
  assert.equal(count, expected.size); let p = start, previousEnd = 0;
  for (const [name, value] of expected) {
    assert.equal(data.readUInt32LE(p), 0x02014b50);
    const method = data.readUInt16LE(p + 10), sum = data.readUInt32LE(p + 16), nameLen = data.readUInt16LE(p + 28), extraLen = data.readUInt16LE(p + 30);
    let packed = data.readUInt32LE(p + 20), size = data.readUInt32LE(p + 24), local = data.readUInt32LE(p + 42);
    const zip64 = size === 0xffffffff;
    if (extraLen) {
      let extra = p + 46 + nameLen; assert.equal(data.readUInt16LE(extra), 1); extra += 4;
      if (size === 0xffffffff) { size = Number(data.readBigUInt64LE(extra)); extra += 8; }
      if (packed === 0xffffffff) { packed = Number(data.readBigUInt64LE(extra)); extra += 8; }
      if (local === 0xffffffff) local = Number(data.readBigUInt64LE(extra));
    }
    assert.equal(data.subarray(p + 46, p + 46 + nameLen).toString('utf8'), name);
    assert.equal(local, previousEnd); assert.equal(data.readUInt16LE(local + 6), 0x0808);
    const payload = local + 30 + data.readUInt16LE(local + 26) + data.readUInt16LE(local + 28);
    const raw = data.subarray(payload, payload + packed), actual = method === 8 ? inflateRawSync(raw) : raw;
    assert.deepEqual(actual, Buffer.from(value)); assert.equal(size, actual.length); assert.equal(sum, crc32(actual));
    const descriptor = payload + packed; assert.equal(data.readUInt32LE(descriptor), 0x08074b50); assert.equal(data.readUInt32LE(descriptor + 4), sum);
    if (zip64) { assert.equal(data.readBigUInt64LE(descriptor + 8), BigInt(packed)); assert.equal(data.readBigUInt64LE(descriptor + 16), BigInt(size)); }
    else { assert.equal(data.readUInt32LE(descriptor + 8), packed); assert.equal(data.readUInt32LE(descriptor + 12), size); }
    previousEnd = descriptor + (zip64 ? 24 : 16); p += 46 + nameLen + extraLen;
  }
  assert.equal(previousEnd, start); assert.equal(p, start + dirSize); return data;
}
function memorySink() { const chunks = []; return { chunks, async write(bytes) { chunks.push(bytes.slice()); } }; }

test('disk ZIP streams ZIP32/ZIP64, compressed/stored, Unicode, empty and async/Blob entries', async () => {
  for (const zip64 of [false, true]) for (const compress of [false, true]) {
    const sink = memorySink(), writer = createZipStreamWriter(sink, { zip64, compress });
    async function* text() { yield '<a>가'; yield ['\ud83d', ['\ude00']]; yield '나</a>'; }
    await writer.add('한글.xml', text()); await writer.add('empty', []); await writer.add('blob', new Blob(['바이트']));
    const meta = await writer.finish(); assert.equal(meta.entries, 3); assert.equal(meta.zip64, zip64);
    const data = inspect(sink.chunks, new Map([['한글.xml', '<a>가😀나</a>'], ['empty', ''], ['blob', '바이트']]));
    assert.equal(meta.bytesWritten, data.length);
    assert.equal(textOf(unzip(data)['한글.xml']), '<a>가😀나</a>');
  }
});

test('slow sink applies backpressure and receives bounded chunks before the producer finishes', async () => {
  let produced = 0, writes = 0, firstPayload = null, maxChunk = 0;
  async function* source() { for (let i = 0; i < 50; i++) { produced++; yield new Uint8Array(1 << 18).fill(i); } }
  const writer = createZipStreamWriter({ async write(bytes) {
    writes++; maxChunk = Math.max(maxChunk, bytes.length);
    if (bytes.length === 1 << 18 && firstPayload === null) firstPayload = produced;
    await new Promise(resolve => setImmediate(resolve));
  } }, { compress: false });
  await writer.add('large', source()); const result = await writer.finish();
  assert.equal(produced, 50); assert.equal(firstPayload, 1); assert.ok(writes > 50); assert.ok(maxChunk <= 1 << 18); assert.ok(result.bytesWritten > 50 * (1 << 18));
});

test('source/sink/abort failures close the source and permanently reject further output', async () => {
  for (const mode of ['source', 'sink', 'abort']) {
    const failure = new Error(mode), abort = new AbortController(); let closed = false, writes = 0;
    async function* source() { try { yield new Uint8Array(1 << 18); if (mode === 'source') throw failure; yield new Uint8Array(1 << 18); } finally { closed = true; } }
    const writer = createZipStreamWriter({ async write() { if (++writes === 2) { if (mode === 'sink') throw failure; if (mode === 'abort') abort.abort(failure); } } }, { compress: false, signal: abort.signal });
    await assert.rejects(writer.add('fail', source()), error => error === failure); assert.equal(closed, true);
    const oldWrites = writes; await assert.rejects(writer.add('later', ''), error => error === failure); await assert.rejects(writer.finish(), error => error === failure); assert.equal(writes, oldWrites);
  }
});

test('writer abort, duplicate names and in-flight operations cannot yield a valid-looking partial archive', async () => {
  const sink = memorySink(), writer = createZipStreamWriter(sink, { compress: false });
  await writer.add('one', '1'); await assert.rejects(writer.add('one', '2'), /순서대로/);
  let release; const gate = new Promise(resolve => { release = resolve; }); let entered;
  const start = new Promise(resolve => { entered = resolve; });
  async function* source() { entered(); await gate; yield 'pending'; }
  const pending = writer.add('two', source()); await start;
  await assert.rejects(writer.finish(), /완료되지/); await assert.rejects(writer.add('three', ''), /순서대로/);
  const failure = new Error('cancel'); const cancel = writer.abort(failure); release(); await cancel;
  await assert.rejects(pending, error => error === failure); await assert.rejects(writer.finish(), error => error === failure);
});

test('ZIP64 directory/descriptor encodes 20 GiB offsets and sizes without 32-bit wraparound', () => {
  const huge = 20 * 1024 ** 3 + 137, offset = huge + 42;
  const cd = Buffer.from(zipCentralRecord({ name: 'large.xml', method: 0, crc: 123, size: huge, compressedSize: huge, offset, zip64: true }));
  const extra = 46 + Buffer.byteLength('large.xml'); assert.equal(cd.readUInt16LE(extra), 1); assert.equal(cd.readUInt16LE(extra + 2), 24);
  assert.equal(cd.readBigUInt64LE(extra + 4), BigInt(huge)); assert.equal(cd.readBigUInt64LE(extra + 12), BigInt(huge)); assert.equal(cd.readBigUInt64LE(extra + 20), BigInt(offset));
  const dd = Buffer.from(zipDescriptorRecord(123, huge, huge, true)); assert.equal(dd.readBigUInt64LE(8), BigInt(huge)); assert.equal(dd.readBigUInt64LE(16), BigInt(huge));
  const [end, locator, classic] = zipEndRecords(70000, 12345678, offset).map(Buffer.from);
  assert.equal(end.readBigUInt64LE(32), 70000n); assert.equal(end.readBigUInt64LE(48), BigInt(offset)); assert.equal(locator.readBigUInt64LE(8), BigInt(offset + 12345678)); assert.equal(classic.readUInt16LE(10), 0xffff);
  assert.throws(() => zipEndRecords(1, 20, Number.MAX_SAFE_INTEGER + 1), /정확하게/);
  assert.throws(() => zipDescriptorRecord(0, huge, huge, false), /ZIP32/);
});

test('streamed native zlib fallback strips framing across one-byte output boundaries', async () => {
  const Native = globalThis.CompressionStream;
  try {
    globalThis.CompressionStream = class { constructor(format) {
      if (format === 'deflate-raw') throw new TypeError('unsupported');
      const stream = new Native(format); return { writable: stream.writable, readable: stream.readable.pipeThrough(new TransformStream({
        transform(bytes, controller) { for (const byte of bytes) controller.enqueue(Uint8Array.of(byte)); },
      })) };
    } };
    const sink = memorySink(), writer = createZipStreamWriter(sink); const text = '가😀'.repeat(2000);
    await writer.add('fallback', [text]); await writer.finish(); inspect(sink.chunks, new Map([['fallback', text]]));
  } finally { globalThis.CompressionStream = Native; }
});


test('ZIP64 reader rejects truncated metadata, invalid offsets and unsafe 64-bit numbers', async () => {
  const sink = memorySink(), writer = createZipStreamWriter(sink, { compress: false });
  await writer.add('entry', 'content'); await writer.finish(); const valid = Buffer.concat(sink.chunks);
  const end = valid.length - 22, zip64End = valid.length - 98, cd = Number(valid.readBigUInt64LE(zip64End + 48)), extra = cd + 46 + 5;
  const changes = [
    bytes => bytes.writeUInt32LE(0, end - 20),
    bytes => bytes.writeBigUInt64LE(BigInt(bytes.length + 1), end - 12),
    bytes => bytes.writeBigUInt64LE(0xffffffffffffffffn, zip64End + 48),
    bytes => bytes.writeBigUInt64LE(0xffffn, zip64End + 32),
    bytes => bytes.writeUInt16LE(0, extra),
    bytes => bytes.writeUInt16LE(8, extra + 2),
    bytes => bytes.writeBigUInt64LE(BigInt(bytes.length + 1), extra + 12),
  ];
  for (const mutate of changes) { const bad = Buffer.from(valid); mutate(bad); assert.throws(() => unzip(bad), /ZIP64/); }
  assert.equal(textOf(unzip(valid).entry), 'content');
});

test('65,535-entry ZIP64 archive has an accurate count and reopens every entry', async () => {
  const sink = memorySink(), writer = createZipStreamWriter(sink, { compress: false });
  for (let i = 0; i < 65535; i++) await writer.add(String(i), '');
  const meta = await writer.finish(), data = Buffer.concat(sink.chunks); assert.equal(meta.entries, 65535);
  const files = unzip(data); assert.equal(Object.keys(files).length, 65535); assert.equal(files['65534'].length, 0);
  // Python/other producers can use a classic EOCD at exactly the count boundary.
  const zip64End = data.length - 98, tail = Buffer.alloc(22), classic = Buffer.concat([data.subarray(0, zip64End), tail]);
  const end = classic.length - 22; classic.writeUInt32LE(0x06054b50, end); classic.writeUInt16LE(0xffff, end + 8); classic.writeUInt16LE(0xffff, end + 10);
  classic.writeUInt32LE(Number(data.readBigUInt64LE(zip64End + 40)), end + 12); classic.writeUInt32LE(Number(data.readBigUInt64LE(zip64End + 48)), end + 16);
  assert.equal(Object.keys(unzip(classic)).length, 65535);
});

test('stream compression failure preserves its cause and closes the lazy source', async () => {
  const Native = globalThis.CompressionStream, failure = new Error('compression failure'); let closed = false;
  async function* source() { try { for (let i = 0; i < 20; i++) yield new Uint8Array(1 << 18); } finally { closed = true; } }
  try {
    globalThis.CompressionStream = class { constructor() { return new TransformStream({ transform() { throw failure; } }); } };
    const writer = createZipStreamWriter(memorySink()); await assert.rejects(writer.add('bad', source()), error => error === failure); assert.equal(closed, true);
    await assert.rejects(writer.finish(), error => error === failure);
  } finally { globalThis.CompressionStream = Native; }
});


function seekSink({failSeek=false}={}) {
  let bytes=new Uint8Array(1<<20),position=0,extent=0;const seeks=[];
  return {seeks,get bytes(){return bytes.slice(0,extent);},async seek(at){if(failSeek)throw new Error('seek failed');assert.ok(at>=0&&at<=extent);seeks.push(at);position=at;},async write(chunk){if(position+chunk.length>bytes.length){const next=new Uint8Array(Math.max(position+chunk.length,bytes.length*2));next.set(bytes);bytes=next;}bytes.set(chunk,position);position+=chunk.length;extent=Math.max(extent,position);}};
}
test('automatic seek ZIP patches local CRC and real sizes, emits classic central/end for small archives',async()=>{
  for(const compress of [false,true]){
    const sink=seekSink(),writer=createZipStreamWriter(sink,{zip64:'auto',compress});
    await writer.add('one.xml','가😀'.repeat(1000));await writer.add('two.xml','tail');const meta=await writer.finish(),data=Buffer.from(sink.bytes);
    assert.equal(meta.zip64,false);assert.equal(meta.bytesWritten,data.length);assert.equal(sink.seeks.length,4);
    assert.equal(data.readUInt16LE(4),20);assert.equal(data.readUInt16LE(6),0x800);assert.equal(data.readUInt32LE(22),Buffer.byteLength('가😀'.repeat(1000)));
    const start=30+data.readUInt16LE(26)+data.readUInt16LE(28),compressed=data.subarray(start,start+data.readUInt32LE(18)),value=compress?inflateRawSync(compressed):compressed;
    assert.equal(data.readUInt32LE(14),crc32(value));assert.equal(value.toString(),'가😀'.repeat(1000));
    const end=data.length-22;assert.equal(data.readUInt16LE(end+10),2);const cd=data.readUInt32LE(end+16);assert.equal(data.readUInt16LE(cd+8),0x800);assert.equal(data.readUInt16LE(cd+30),0);
    const files=unzip(data);assert.equal(textOf(files['one.xml']),'가😀'.repeat(1000));assert.equal(textOf(files['two.xml']),'tail');
  }
});
test('automatic nonseek ZIP stays classic ZIP32 and propagates header seek failure',async()=>{
  const sink=memorySink(),writer=createZipStreamWriter(sink,{zip64:'auto'});await writer.add('a','value');const result=await writer.finish();assert.equal(result.zip64,false);inspect(sink.chunks,new Map([['a','value']]));
  const failing=createZipStreamWriter(seekSink({failSeek:true}),{zip64:'auto'});await assert.rejects(failing.add('a','value'),/seek failed/);await assert.rejects(failing.finish(),/seek failed/);
});
