// Actual local disk test. Generated data is deleted after hash verification.
// Usage: node tools/file-save-disk.mjs D:/Codex/Temp/wixel-large-save [bytes]
import { open, mkdir, rename, rm, writeFile } from 'node:fs/promises';
import { createReadStream } from 'node:fs';
import { createHash, randomUUID } from 'node:crypto';
import { resolve, join, dirname } from 'node:path';
import { writeFileHandle } from '../src/file-save.js';

const directory = resolve(process.argv[2] || 'D:/Codex/Temp/wixel-large-save');
const size = Number(process.argv[3] || 1024 ** 3);
if (!Number.isSafeInteger(size) || size < 1) throw new Error('Positive safe integer bytes required');
await mkdir(directory, { recursive: true });
const name = 'file-save-' + randomUUID(), destination = join(directory, name + '.bin'), temporary = destination + '.partial';
for (const path of [destination, temporary]) if (dirname(path) !== directory) throw new Error('Unsafe generated output path');
const expected = createHash('sha256'), chunk = new Uint8Array(1 << 20), view = new DataView(chunk.buffer);
for (let i = 0; i < chunk.length; i++) chunk[i] = ((i * 31 + 17) ^ ((i >>> 8) * 7)) & 255;
let file, at = 0, calls = 0, maximumChunk = 0, committed = false;
const startRss = process.memoryUsage().rss; let maximumRss = startRss;
const target = { async createWritable() {
  file = await open(temporary, 'wx');
  return {
    async write(bytes) {
      calls++; maximumChunk = Math.max(maximumChunk, bytes.byteLength);
      maximumRss = Math.max(maximumRss, process.memoryUsage().rss);
      let offset = 0;
      while (offset < bytes.length) {
        const result = await file.write(bytes, offset, bytes.length - offset, at);
        if (result.bytesWritten < 1) throw new Error('Disk write made no progress');
        offset += result.bytesWritten; at += result.bytesWritten;
      }
    },
    async close() { await file.sync(); await file.close(); file = null; await rename(temporary, destination); committed = true; },
    async abort() { await file?.close(); file = null; await rm(temporary, { force: true }); }
  };
} };
let report;
try {
  const started = performance.now();
  const result = await writeFileHandle(target, async sink => {
    for (let position = 0; position < size; position += chunk.length) {
      view.setUint32(0, position / chunk.length, true);
      const part = chunk.subarray(0, Math.min(chunk.length, size - position));
      expected.update(part); await sink.write(part);
    }
  });
  const writeMs = performance.now() - started;
  const stored = createHash('sha256'), readStarted = performance.now(); let readBytes = 0;
  for await (const bytes of createReadStream(destination, { highWaterMark: 1 << 20 })) { stored.update(bytes); readBytes += bytes.length; }
  const expectedHash = expected.digest('hex'), storedHash = stored.digest('hex');
  if (result.bytesWritten !== size || at !== size || readBytes !== size || expectedHash !== storedHash) throw new Error('Saved bytes/hash mismatch');
  report = { kind: 'actual-local-disk-sequential-write-and-full-reread', bytes: size, bytesWritten: result.bytesWritten, readBytes,
    sha256: storedHash, match: true, nativeWrites: calls, maxWriteBytes: maximumChunk, writeMs: Math.round(writeMs),
    rereadMs: Math.round(performance.now() - readStarted), startRssMiB: Math.round(startRss / 1048576), maxWriteRssMiB: Math.round(maximumRss / 1048576),
    retainedDataFile: false, caveat: 'Node filesystem target adapter; not native browser picker, workbook export, or iPad testing.' };
} finally {
  await file?.close().catch(() => {});
  await rm(temporary, { force: true });
  if (committed) await rm(destination, { force: true });
}
await writeFile(join(directory, 'disk-save-result.json'), JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify(report));
