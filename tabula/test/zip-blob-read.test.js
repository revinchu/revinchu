import test from 'node:test';
import assert from 'node:assert/strict';
import { deflateRawSync } from 'node:zlib';
import { unzip, unzipBlob, zipEntryInfo, zipEntryChunks, prepareZipEntryRaw, prepareZipEntry, deleteZipEntry } from '../src/zip.js';
import { blobZipEntries } from '../src/zip-blob-read.js';
function crc(data) {
  let value = 0xffffffff;
  for (const byte of data) { value ^= byte; for (let i = 0; i < 8; i++) value = (value >>> 1) ^ (value & 1 ? 0xedb88320 : 0); }
  return (value ^ 0xffffffff) >>> 0;
}
function makeZip(entries, { zip64 = false, comment = Buffer.alloc(0) } = {}) {
  const locals = [], centrals = [], offsets = new Map(); let offset = 0;
  for (const entry of entries) {
    const { name, data, method = 8 } = entry, nameBytes = Buffer.from(name), packed = method === 8 ? deflateRawSync(data) : data;
    const extraLocal = zip64 ? 20 : 0, extraCentral = zip64 ? 28 : 0, local = Buffer.alloc(30 + nameBytes.length + extraLocal), central = Buffer.alloc(46 + nameBytes.length + extraCentral);
    local.writeUInt32LE(0x04034b50); local.writeUInt16LE(zip64 ? 45 : 20, 4); local.writeUInt16LE(0x800, 6); local.writeUInt16LE(method, 8); local.writeUInt32LE(crc(data), 14);
    local.writeUInt32LE(zip64 ? 0xffffffff : packed.length, 18); local.writeUInt32LE(zip64 ? 0xffffffff : data.length, 22); local.writeUInt16LE(nameBytes.length, 26); local.writeUInt16LE(extraLocal, 28); nameBytes.copy(local, 30);
    if (zip64) { const p = 30 + nameBytes.length; local.writeUInt16LE(1, p); local.writeUInt16LE(16, p + 2); local.writeBigUInt64LE(BigInt(data.length), p + 4); local.writeBigUInt64LE(BigInt(packed.length), p + 12); }
    central.writeUInt32LE(0x02014b50); central.writeUInt16LE(zip64 ? 45 : 20, 4); central.writeUInt16LE(zip64 ? 45 : 20, 6); central.writeUInt16LE(0x800, 8); central.writeUInt16LE(method, 10); central.writeUInt32LE(entry.crc ?? crc(data), 16);
    central.writeUInt32LE(zip64 ? 0xffffffff : packed.length, 20); central.writeUInt32LE(zip64 ? 0xffffffff : data.length, 24); central.writeUInt16LE(nameBytes.length, 28); central.writeUInt16LE(extraCentral, 30); central.writeUInt32LE(zip64 ? 0xffffffff : offset, 42); nameBytes.copy(central, 46);
    if (zip64) { const p = 46 + nameBytes.length; central.writeUInt16LE(1, p); central.writeUInt16LE(24, p + 2); central.writeBigUInt64LE(BigInt(data.length), p + 4); central.writeBigUInt64LE(BigInt(packed.length), p + 12); central.writeBigUInt64LE(BigInt(offset), p + 20); }
    offsets.set(name, { local: offset, payload: offset + local.length, packed: packed.length }); locals.push(local, packed); centrals.push(central); offset += local.length + packed.length;
  }
  const directory = Buffer.concat(centrals), ends = [], end = Buffer.alloc(22 + comment.length);
  if (zip64) {
    const record = Buffer.alloc(56), locator = Buffer.alloc(20);
    record.writeUInt32LE(0x06064b50); record.writeBigUInt64LE(44n, 4); record.writeUInt16LE(45, 12); record.writeUInt16LE(45, 14); record.writeBigUInt64LE(BigInt(entries.length), 24); record.writeBigUInt64LE(BigInt(entries.length), 32); record.writeBigUInt64LE(BigInt(directory.length), 40); record.writeBigUInt64LE(BigInt(offset), 48);
    locator.writeUInt32LE(0x07064b50); locator.writeBigUInt64LE(BigInt(offset + directory.length), 8); locator.writeUInt32LE(1, 16); ends.push(record, locator);
  }
  end.writeUInt32LE(0x06054b50); end.writeUInt16LE(zip64 ? 0xffff : entries.length, 8); end.writeUInt16LE(zip64 ? 0xffff : entries.length, 10); end.writeUInt32LE(zip64 ? 0xffffffff : directory.length, 12); end.writeUInt32LE(zip64 ? 0xffffffff : offset, 16); end.writeUInt16LE(comment.length, 20); comment.copy(end, 22);
  return { bytes: Buffer.concat([...locals, directory, ...ends, end]), offsets, directoryOffset: offset, directorySize: directory.length };
}
function tracked(bytes) {
  const reads = []; return { size: bytes.length, reads,
    arrayBuffer() { throw Error('The original Blob must never be fully read.'); },
    slice(start, end) { reads.push({ start, end, size: end - start }); return new Blob([bytes.subarray(start, end)]); },
  };
}
const collect = chunks => Buffer.concat(Array.from(chunks, chunk => Buffer.from(chunk)));
function fixture() {
  return makeZip([
    { name: 'xl/workbook.xml', data: Buffer.from('<workbook>한글</workbook>') },
    { name: 'xl/worksheets/sheet1.xml', data: Buffer.from('가나다😀'.repeat(10000)) },
    { name: 'xl/pivotCache/pivotCacheRecords1.xml', data: Buffer.from('<r><n v="1"/></r>'.repeat(5000)) },
    { name: 'xl/media/image1.bin', data: Buffer.alloc(180000, 137), method: 0 },
  ]);
}

test('Blob indexing reads only end records and central directory, leaving every payload lazy', async () => {
  const original = fixture(), blob = tracked(original.bytes), entries = await blobZipEntries(blob);
  assert.equal(entries.length, 4); assert.equal(blob.reads.length, 2);
  assert.deepEqual(blob.reads[1], { start: original.directoryOffset, end: original.directoryOffset + original.directorySize, size: original.directorySize });
  assert.ok(blob.reads.every(read => read.size <= 65557));
  assert.equal(entries[1].size, Buffer.byteLength('가나다😀'.repeat(10000)));
  const at = blob.reads.length, raw = await entries[1].loadRaw(), part = original.offsets.get(entries[1].name);
  assert.equal(raw.length, part.packed); assert.deepEqual(blob.reads.slice(at).map(read => read.size), [30, part.packed]);
});

test('default Blob preload keeps worksheets/records compressed until requested and supports stored media', async () => {
  const original = fixture(), blob = tracked(original.bytes), progress = [], files = await unzipBlob(blob, { onProgress: value => progress.push(value) });
  assert.equal(progress.length, 4); assert.equal(progress.at(-1).p, 1);
  for (const name of ['xl/worksheets/sheet1.xml', 'xl/pivotCache/pivotCacheRecords1.xml']) {
    assert.throws(() => files[name], /비동기 준비/);
    assert.ok(zipEntryInfo(files, name));
    const before = blob.reads.length; await prepareZipEntryRaw(files, name);
    assert.deepEqual(blob.reads.slice(before).map(read => read.size), [30, original.offsets.get(name).packed]);
    const getter = Object.getOwnPropertyDescriptor(files, name).get;
    const expected = name.includes('worksheets') ? Buffer.from('가나다😀'.repeat(10000)) : Buffer.from('<r><n v="1"/></r>'.repeat(5000));
    assert.deepEqual(collect(zipEntryChunks(files, name, { chunkSize: 65536 })), expected);
    assert.equal(Object.getOwnPropertyDescriptor(files, name).get, getter, 'expanded worksheet/record remains uncached');
  }
  assert.equal(new TextDecoder().decode(files['xl/workbook.xml']), '<workbook>한글</workbook>');
  assert.deepEqual(collect(zipEntryChunks(files, 'xl/media/image1.bin')), Buffer.alloc(180000, 137));
  assert.equal(progress.at(-1).preloadedBytes, original.offsets.get('xl/workbook.xml').packed + 180000);
});

test('raw preparation deduplicates concurrent reads; full prepare and synchronous byte ZIP stay compatible', async () => {
  const { bytes } = fixture(), blob = tracked(bytes), files = await unzipBlob(blob, { preload: () => false }), name = 'xl/worksheets/sheet1.xml';
  const before = blob.reads.length, [a, b] = await Promise.all([prepareZipEntryRaw(files, name), prepareZipEntryRaw(files, name)]);
  assert.equal(a, b); assert.equal(blob.reads.length - before, 2);
  const loaded = blob.reads.length; await prepareZipEntry(files, name); assert.equal(blob.reads.length, loaded);
  assert.deepEqual(Buffer.from(files[name]), Buffer.from('가나다😀'.repeat(10000)));
  assert.deepEqual(collect(zipEntryChunks(files, name)), Buffer.from(files[name]));
  const sync = unzip(bytes); assert.deepEqual(Buffer.from(sync[name]), Buffer.from(files[name]));
  assert.ok(Object.getOwnPropertyDescriptor(sync, 'xl/pivotCache/pivotCacheRecords1.xml').get);
});

test('deleting a consumed Blob entry removes its raw accessor, metadata and future preparation', async () => {
  const { bytes } = fixture(), files = await unzipBlob(tracked(bytes), { preload: () => false }), name = 'xl/worksheets/sheet1.xml';
  await prepareZipEntryRaw(files, name); collect(zipEntryChunks(files, name));
  assert.equal(deleteZipEntry(files, name), true); assert.equal(Object.hasOwn(files, name), false); assert.equal(zipEntryInfo(files, name), null);
  await assert.rejects(prepareZipEntryRaw(files, name), TypeError); assert.throws(() => collect(zipEntryChunks(files, name)), /찾을/);
  assert.equal(deleteZipEntry(files, name), true);
});

test('deletion during an async raw read rejects cancellation and cannot resurrect the entry', async () => {
  const { bytes, offsets } = fixture(), base = tracked(bytes), name = 'xl/worksheets/sheet1.xml', target = offsets.get(name); let resolveRead;
  const blob = { size: base.size, slice(start, end) {
    const sliced = base.slice(start, end);
    if (start === target.payload) return { arrayBuffer: () => new Promise(resolve => { resolveRead = async () => resolve(await sliced.arrayBuffer()); }) };
    return sliced;
  } };
  const files = await unzipBlob(blob, { preload: () => false }), pending = prepareZipEntryRaw(files, name);
  while (!resolveRead) await new Promise(resolve => setImmediate(resolve));
  deleteZipEntry(files, name); await resolveRead(); await assert.rejects(pending, /취소/); assert.equal(Object.hasOwn(files, name), false);
});

test('Blob ZIP64 sizes/offsets and maximum comments use bounded end/index reads', async () => {
  const comment = Buffer.alloc(65535, 88); comment.writeUInt32LE(0x06054b50, 1000);
  const original = makeZip([{ name: '파일.xml', data: Buffer.from('큰 목록 한글') }], { zip64: true, comment });
  const blob = tracked(original.bytes), files = await unzipBlob(blob, { preload: () => false });
  assert.ok(blob.reads.every(read => read.size <= 65557));
  await prepareZipEntryRaw(files, '파일.xml'); assert.equal(collect(zipEntryChunks(files, '파일.xml')).toString('utf8'), '큰 목록 한글');
  assert.deepEqual(zipEntryInfo(files, '파일.xml'), { size: Buffer.byteLength('큰 목록 한글'), compressedSize: original.offsets.get('파일.xml').packed, method: 8, crc: crc(Buffer.from('큰 목록 한글')) });
});

test('optional read budgets reject before payload allocations while omitted budgets preserve compatibility', async () => {
  const original = fixture();
  await assert.rejects(unzipBlob(tracked(original.bytes), { maxDirectoryBytes: original.directorySize - 1 }), /목록.*예산/);
  await assert.rejects(unzipBlob(tracked(original.bytes), { maxPreloadBytes: 1 }), /사전 읽기.*예산/);
  const blob = tracked(original.bytes), files = await unzipBlob(blob, { preload: () => false, maxEntryBytes: 1 }), count = blob.reads.length;
  await assert.rejects(prepareZipEntryRaw(files, 'xl/media/image1.bin'), /항목.*예산/); assert.equal(blob.reads.length, count, 'budget rejects before the local header/payload read');
  const compatible = await unzipBlob(tracked(original.bytes)); assert.equal(compatible['xl/media/image1.bin'].length, 180000);
  for (const limit of [NaN, -1, 0.5]) await assert.rejects(unzipBlob(tracked(original.bytes), { maxPreloadBytes: limit }), RangeError);
});

test('Blob directory, local header and CRC corruption are rejected at their consumption boundary', async () => {
  const original = fixture();
  const badDirectory = Buffer.from(original.bytes); badDirectory.writeUInt32LE(0, original.directoryOffset);
  await assert.rejects(unzipBlob(tracked(badDirectory), { preload: () => false }), /목록/);
  const truncated = original.bytes.subarray(0, original.bytes.length - 1); await assert.rejects(unzipBlob(tracked(truncated)), /끝부분/);
  const badLocal = Buffer.from(original.bytes), name = 'xl/worksheets/sheet1.xml'; badLocal.writeUInt32LE(0, original.offsets.get(name).local);
  const files = await unzipBlob(tracked(badLocal), { preload: () => false }); await assert.rejects(prepareZipEntryRaw(files, name), /머리글/);
  const badCrc = makeZip([{ name, data: Buffer.from('원문'), crc: 1 }]), corrupted = await unzipBlob(tracked(badCrc.bytes), { preload: () => false });
  await prepareZipEntryRaw(corrupted, name); assert.throws(() => collect(zipEntryChunks(corrupted, name)), /CRC/);
});

test('progress cancellation stops metadata preparation and custom preload may select exactly one part', async () => {
  const original = fixture(), stop = Error('cancel index preparation'), blob = tracked(original.bytes); let called = 0;
  await assert.rejects(unzipBlob(blob, { onProgress() { called++; throw stop; } }), error => error === stop); assert.equal(called, 1);
  const custom = tracked(original.bytes), target = 'xl/pivotCache/pivotCacheRecords1.xml', files = await unzipBlob(custom, { preload: name => name === target });
  assert.deepEqual(collect(zipEntryChunks(files, target)), Buffer.from('<r><n v="1"/></r>'.repeat(5000)));
  assert.throws(() => files['xl/workbook.xml'], /비동기 준비/);
});
