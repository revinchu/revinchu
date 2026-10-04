// Node 22+: opt-in ZIP layer stress (synthetic bytes, not a workbook capacity claim).
// node tools/zip-stream-stress.mjs D:/Codex/Temp/zip-stress.zip 20
import { open, stat } from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import { createInflateRaw, crc32 } from 'node:zlib';
import { createZipStreamWriter } from '../src/zip-stream.js';
const path = process.argv[2], gib = Number(process.argv[3] ?? 1);
if (!path || !Number.isFinite(gib) || gib <= 0 || gib > 20) throw new Error('출력 경로와 0~20 GiB 범위를 지정해 주세요.');
const stored = process.argv[4] === 'stored';
const bytes = Math.floor(gib * 1024 ** 3), handle = await open(path, 'wx'), started = performance.now();
let produced = 0, expectedCrc = 0, peakRss = process.memoryUsage().rss, nextReport = 1024 ** 3, writes = 0, largestWrite = 0;
const chunk = new Uint8Array(1 << 18);
async function* source() {
  while (produced < bytes) {
    const value = chunk.subarray(0, Math.min(chunk.length, bytes - produced)); produced += value.length; expectedCrc = crc32(value, expectedCrc);
    if (produced >= nextReport) { peakRss = Math.max(peakRss, process.memoryUsage().rss); console.log(JSON.stringify({ producedGiB: produced / 1024 ** 3, elapsedSeconds: (performance.now() - started) / 1000, rssMiB: peakRss / 1024 ** 2 })); nextReport += 1024 ** 3; }
    yield value;
  }
}
let result;
try {
  const writer = createZipStreamWriter({ async write(value) { writes++; largestWrite = Math.max(largestWrite, value.length); for (let at = 0; at < value.length;) { const written = await handle.write(value, at, value.length - at); if (!written.bytesWritten) throw new Error('파일 쓰기가 진행되지 않았습니다.'); at += written.bytesWritten; } } }, { compress: !stored });
  await writer.add('synthetic.bin', source()); if (stored) await writer.add('tail.txt', 'tail'); result = await writer.finish(); await handle.sync();
} finally { await handle.close(); }
// Independent Buffer metadata + native inflater/CRC, retaining only streamed chunks.
const check = await open(path, 'r'), size = (await stat(path)).size, tail = Buffer.alloc(98);
await check.read(tail, 0, tail.length, size - tail.length);
if (tail.readUInt32LE(0) !== 0x06064b50) throw new Error('ZIP64 종료 레코드가 없습니다.');
const directoryOffset = Number(tail.readBigUInt64LE(48)), header = Buffer.alloc(128);
await check.read(header, 0, header.length, directoryOffset);
const nameLength = header.readUInt16LE(28), extra = 46 + nameLength;
const uncompressed = Number(header.readBigUInt64LE(extra + 4)), compressed = Number(header.readBigUInt64LE(extra + 12));
const local = Buffer.alloc(30); await check.read(local, 0, 30, 0); await check.close();
const start = 30 + local.readUInt16LE(26) + local.readUInt16LE(28);
let decoded = 0, checksum = 0;
const input = createReadStream(path, { start, end: start + compressed - 1 });
for await (const value of stored ? input : input.pipe(createInflateRaw())) { decoded += value.length; checksum = crc32(value, checksum); }
if (decoded !== bytes || uncompressed !== bytes || checksum !== expectedCrc || checksum !== header.readUInt32LE(16)) throw new Error('독립 ZIP 검증이 실패했습니다.');
if (stored) {
  const tailCheck = await open(path, 'r'), central = Buffer.alloc(128), next = directoryOffset + 46 + nameLength + header.readUInt16LE(30);
  await tailCheck.read(central, 0, central.length, next); const extraStart = 46 + central.readUInt16LE(28);
  const recordedOffset = central.readUInt32LE(42);
  if (bytes >= 0xffffffff && recordedOffset !== 0xffffffff) throw new Error('두 번째 항목의 ZIP64 위치가 없습니다.');
  const location = recordedOffset === 0xffffffff ? Number(central.readBigUInt64LE(extraStart + 20)) : recordedOffset, localTail = Buffer.alloc(30);
  await tailCheck.read(localTail, 0, 30, location); const value = Buffer.alloc(4);
  await tailCheck.read(value, 0, 4, location + 30 + localTail.readUInt16LE(26) + localTail.readUInt16LE(28)); await tailCheck.close();
  if (value.toString() !== 'tail') throw new Error('4GiB 이후 항목 위치가 잘못되었습니다.');
}
console.log(JSON.stringify({ ok: true, path, sourceBytes: bytes, diskBytes: size, writes, largestWrite, peakRssMiB: peakRss / 1024 ** 2, seconds: (performance.now() - started) / 1000, ...result }));
