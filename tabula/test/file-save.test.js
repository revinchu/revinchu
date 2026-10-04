import test from 'node:test';
import assert from 'node:assert/strict';
import { writeFileHandle, isFileSaveSource } from '../src/file-save.js';
import { createWixelFileStream, readWixelFile } from '../src/wixel-file.js';
import { Workbook } from '../src/workbook.js';

function target(options = {}) {
  const log = { writes: 0, maxChunk: 0, bytes: 0, closes: 0, aborts: 0, active: 0, chunks: [] };
  return { log, async createWritable() {
    if (options.openError) throw options.openError;
    return {
      async write(chunk) {
        assert.equal(log.active++, 0, 'Writes must apply backpressure instead of overlapping');
        try {
          await Promise.resolve();
          if (options.writeError) throw options.writeError;
          log.writes++; log.bytes += chunk.byteLength; log.maxChunk = Math.max(log.maxChunk, chunk.byteLength);
          if (options.collect) log.chunks.push(chunk.slice());
          options.onWrite?.(log);
        } finally { log.active--; }
      },
      async close() { if (options.closeError) throw options.closeError; log.closes++; },
      async abort() { log.aborts++; }
    };
  } };
}

test('direct file writing supports Blob, bytes, streams and async byte producers', async () => {
  const data = new TextEncoder().encode('한글😀,stream');
  const sources = [new Blob([data]), data, data.buffer, new DataView(data.buffer), [data.subarray(0, 3), data.subarray(3)],
    new ReadableStream({ start(c) { c.enqueue(data); c.close(); } }), async sink => { await sink.write(data.subarray(0, 3)); await sink.write(data.subarray(3)); }];
  for (const source of sources) {
    assert.equal(isFileSaveSource(source), true);
    const handle = target({ collect: true });
    const result = await writeFileHandle(handle, source);
    assert.equal(result.bytesWritten, data.length); assert.equal(handle.log.closes, 1); assert.equal(handle.log.aborts, 0);
    assert.deepEqual(new Uint8Array(await new Blob(handle.log.chunks).arrayBuffer()), data);
  }
  for (const source of [null, {}, 3, 'not bytes']) assert.equal(isFileSaveSource(source), false);
});

test('direct writer bounds a large source chunk before each native write', async () => {
  const handle = target(), bytes = new Uint8Array((3 << 20) + 17);
  const values = [];
  const result = await writeFileHandle(handle, bytes, { onProgress(written, total) { assert.equal(total, bytes.byteLength); values.push(written); } });
  assert.equal(result.bytesWritten, bytes.byteLength); assert.equal(handle.log.writes, 4); assert.equal(handle.log.maxChunk, 1 << 20);
  assert.deepEqual(values, [1 << 20, 2 << 20, 3 << 20, bytes.length]);
});

test('count-only 20 GiB producer crosses 32-bit sizes without allocating the file', async () => {
  const handle = target(), chunk = new Uint8Array(1 << 20), count = 20 * 1024;
  let produced = 0, latest = 0;
  const result = await writeFileHandle(handle, async sink => {
    for (let i = 0; i < count; i++) { assert.equal(handle.log.writes, produced); await sink.write(chunk); produced++; }
    await sink.write(chunk.subarray(0, 7));
  }, { onProgress(n) { assert.ok(n > latest); latest = n; } });
  assert.equal(result.bytesWritten, 20 * 1024 ** 3 + 7); assert.equal(latest, result.bytesWritten);
  assert.equal(handle.log.writes, count + 1); assert.equal(handle.log.closes, 1); assert.equal(handle.log.aborts, 0);
});

test('producer failure aborts the native temporary file instead of committing partial data', async () => {
  const handle = target(), failure = new Error('source failed');
  await assert.rejects(writeFileHandle(handle, async sink => { await sink.write(new Uint8Array(1 << 20)); throw failure; }), error => error === failure);
  assert.equal(handle.log.bytes, 1 << 20); assert.equal(handle.log.aborts, 1); assert.equal(handle.log.closes, 0);
});

test('disk failures cancel the source iterator and preserve the original exception', async () => {
  const failure = new Error('disk full'), handle = target({ writeError: failure }); let cancelled = false;
  async function* parts() { try { yield new Uint8Array(1 << 20); assert.fail('Must stop after disk failure'); } finally { cancelled = true; } }
  await assert.rejects(writeFileHandle(handle, parts()), error => error === failure);
  assert.equal(cancelled, true); assert.equal(handle.log.aborts, 1); assert.equal(handle.log.closes, 0);
});

test('document change at the last write aborts before closing', async () => {
  let current = true; const handle = target({ onWrite() { current = false; } });
  await assert.rejects(writeFileHandle(handle, new Uint8Array([1]), { isCurrent: () => current }), error => error.code === 'FILE_SAVE_ABORT');
  assert.equal(handle.log.aborts, 1); assert.equal(handle.log.closes, 0);
});

test('failed target acquisition cancels an unconsumed readable stream', async () => {
  const failure = new Error('permission denied'), handle = target({ openError: failure }); let cancelled = false;
  const stream = new ReadableStream({ cancel() { cancelled = true; } }, { highWaterMark: 0 });
  await assert.rejects(writeFileHandle(handle, stream), error => error === failure);
  assert.equal(cancelled, true); assert.equal(handle.log.writes, 0); assert.equal(handle.log.aborts, 0);
});

test('abort signals and closing failures cannot commit a partial file', async () => {
  const abort = new AbortController(), failure = new Error('cancelled');
  const handle = target({ onWrite() { abort.abort(failure); } });
  await assert.rejects(writeFileHandle(handle, new Uint8Array([1]), { signal: abort.signal }), error => error === failure);
  assert.equal(handle.log.aborts, 1); assert.equal(handle.log.closes, 0);
  const closing = target({ closeError: failure });
  await assert.rejects(writeFileHandle(closing, new Uint8Array([1])), error => error === failure);
  assert.equal(closing.log.aborts, 1); assert.equal(closing.log.closes, 0);
});

test('empty files, invalid chunks and swallowed producer failures do not commit', async () => {
  for (const source of [new Blob(), [new Uint8Array()], ['invalid'], async sink => { try { await sink.write('bad'); } catch { /* Producer cannot turn failure into success. */ } }]) {
    const handle = target(); await assert.rejects(writeFileHandle(handle, source)); assert.equal(handle.log.closes, 0); assert.equal(handle.log.aborts, 1);
  }
});

test('WIXEL streaming writes no intermediate Blob and round-trips cell and workbook metadata', async () => {
  const book = new Workbook({ sheets: [{ name: '검증', cells: { '0,0': { raw: '한글😀', style: { fill: '#aabbcc', bold: true } } },
    page: { printArea: 'A1:C20' }, slicers: [{ id: 'synthetic', style: '내 스타일', selected: ['가'] }] }] });
  const NativeBlob = globalThis.Blob;
  for (const gzip of [false, true]) {
    const handle = target({ collect: true });
    try {
      globalThis.Blob = class extends NativeBlob { constructor() { throw new Error('Intermediate Blob forbidden'); } };
      await writeFileHandle(handle, createWixelFileStream(book, { docName: '직접 저장' }, { gzip }));
    } finally { globalThis.Blob = NativeBlob; }
    const data = await readWixelFile(new Blob(handle.log.chunks));
    assert.equal(data.docName, '직접 저장');
    assert.deepEqual(new Workbook(data.workbook).serialize(), book.serialize());
    assert.equal(handle.log.closes, 1); assert.equal(handle.log.aborts, 0);
  }
});

test('WIXEL mutation during direct streaming aborts the target', async () => {
  const book = new Workbook(); for (let r = 0; r < 1000; r++) book.setInput(0, r, 0, '합성 행 ' + r);
  const version = book.version, handle = target({ onWrite() { book.version++; } });
  await assert.rejects(writeFileHandle(handle, createWixelFileStream(book, {}, { gzip: false }), { isCurrent: () => book.version === version }), error => error.code === 'WIXEL_FILE_ABORT' || error.code === 'FILE_SAVE_ABORT');
  assert.equal(handle.log.closes, 0); assert.equal(handle.log.aborts, 1);
});


test('small gzip and ZIP chunks share bounded writes and do not retain reused source arrays', async () => {
  const handle = target({ collect: true }), source = new Uint8Array(16384);
  await writeFileHandle(handle, async sink => {
    for (let i = 0; i < 129; i++) { source.fill(i); await sink.write(source); }
    source.fill(255);
  });
  assert.equal(handle.log.writes, 3); assert.equal(handle.log.maxChunk, 1 << 20);
  const bytes = new Uint8Array(await new Blob(handle.log.chunks).arrayBuffer());
  for (let i = 0; i < 129; i++) { assert.equal(bytes[i * source.length], i); assert.equal(bytes[(i + 1) * source.length - 1], i); }
});


function seekTarget({ seekError, onSeek, collect = true } = {}) {
  const data = new Uint8Array(4 << 20), events = []; let position = 0, extent = 0, closes = 0, aborts = 0;
  return { data, events, stats: () => ({ position, extent, closes, aborts }), async createWritable() { return {
    async write(chunk) { if (collect) data.set(chunk, position); events.push(['write', position, chunk.length]); position += chunk.length; extent = Math.max(extent, position); },
    async seek(at) { events.push(['seek', at]); if (seekError) throw seekError; position = at; onSeek?.(); },
    async close() { closes++; }, async abort() { aborts++; },
  }; } };
}

test('seekable file writer flushes buffered bytes, patches a header and returns extent rather than cumulative writes', async () => {
  const handle = seekTarget(), progress = [];
  const result = await writeFileHandle(handle, async sink => {
    assert.equal(typeof sink.seek, 'function'); await sink.write(Uint8Array.from([1,2,3,4,5,6]));
    assert.equal(sink.position, 6); assert.equal(sink.extent, 6); await sink.seek(1);
    await sink.write(Uint8Array.from([9,8])); assert.equal(sink.position, 3); assert.equal(sink.extent, 6);
    await sink.seek(6); await sink.write(Uint8Array.from([7]));
  }, { onProgress(n) { progress.push(n); } });
  assert.deepEqual(Array.from(handle.data.subarray(0,7)), [1,9,8,4,5,6,7]); assert.equal(result.bytesWritten, 7);
  assert.deepEqual(progress,[6,7]); assert.deepEqual(handle.events.slice(0,4), [['write',0,6],['seek',1],['write',1,2],['seek',6]]); assert.equal(handle.stats().closes,1);
});
test('seekable file writer does not expose seeking when the native target has no seek', async () => {
  await writeFileHandle(target(), async sink => { assert.equal(sink.seek, undefined); await sink.write(Uint8Array.of(1)); });
});
test('seekable file writer rejects holes, negative/fractional/unsafe positions and cannot swallow a seek failure', async () => {
  for (const position of [-1, 4, .5, Infinity, Number.MAX_SAFE_INTEGER + 1]) {
    const handle = seekTarget(); await assert.rejects(writeFileHandle(handle, async sink => {
      await sink.write(Uint8Array.from([1,2,3])); try { await sink.seek(position); } catch {};
    }), /범위/); assert.equal(handle.stats().closes,0); assert.equal(handle.stats().aborts,1);
  }
  const failure = new Error('native seek failed'), handle = seekTarget({seekError:failure});
  await assert.rejects(writeFileHandle(handle, async sink => {await sink.write(Uint8Array.of(1)); await sink.seek(0);}), error => error===failure);
  assert.equal(handle.stats().closes,0); assert.equal(handle.stats().aborts,1);
});
test('seekable file writer aborts on document replacement during seek', async () => {
  let current=true;const handle=seekTarget({onSeek(){current=false;}});
  await assert.rejects(writeFileHandle(handle,async sink=>{await sink.write(Uint8Array.of(1));await sink.seek(0);},{isCurrent:()=>current}),error=>error.code==='FILE_SAVE_ABORT');
  assert.equal(handle.stats().closes,0);assert.equal(handle.stats().aborts,1);
});
test('seekable file writer rejects concurrent seeks/writes', async () => {
  const handle=seekTarget();await assert.rejects(writeFileHandle(handle,async sink=>{await sink.write(Uint8Array.of(1));const first=sink.seek(0);await assert.rejects(sink.write(Uint8Array.of(2)),/순서대로/);await first;}),/순서대로/);
  assert.equal(handle.stats().closes,0);assert.equal(handle.stats().aborts,1);
});
test('seekable file writer keeps exact >4GiB positions and extent without allocating a file',async()=>{
  const handle=seekTarget({collect:false}),chunk=new Uint8Array(1<<20),size=5*1024**3;
  const result=await writeFileHandle(handle,async sink=>{for(let i=0;i<5120;i++)await sink.write(chunk);await sink.seek(size-3);await sink.write(Uint8Array.from([1,2,3]));await sink.seek(0);await sink.write(Uint8Array.of(9));});
  assert.equal(result.bytesWritten,size);assert.equal(handle.stats().extent,size);assert.equal(handle.stats().closes,1);
});
