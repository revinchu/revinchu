// Isolated sequential ZIP writer; callers choose it explicitly for a disk sink.
import { zipCount, zipStreamLocalRecord, zipSeekLocalRecord, zipDescriptorRecord, zipCentralRecord, zipEndRecords } from './zip-records.js';
const enc = new TextEncoder(), ZIP32_MAX = 0xffffffff;
const crcTable = new Uint32Array(256);
for (let n = 0; n < 256; n++) { let c = n; for (let bit = 0; bit < 8; bit++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; crcTable[n] = c; }
function crcUpdate(crc, bytes) { for (let i = 0; i < bytes.length; i++) crc = crcTable[(crc ^ bytes[i]) & 255] ^ (crc >>> 8); return crc; }
const pause = () => new Promise(resolve => setTimeout(resolve, 0));

async function* zipStreamParts(value) {
  if (typeof value === 'string' || value instanceof Uint8Array) yield value;
  else if (typeof Blob === 'function' && value instanceof Blob) yield* zipStreamParts(value.stream());
  else if (value?.[Symbol.asyncIterator] || value?.[Symbol.iterator]) { for await (const part of value) yield* zipStreamParts(part); }
  else if (value?.getReader) {
    const reader = value.getReader(); let complete = false;
    try { for (;;) { const next = await reader.read(); if (next.done) { complete = true; break; } yield* zipStreamParts(next.value); } }
    finally { if (!complete) { try { await reader.cancel(); } catch {} } reader.releaseLock(); }
  } else throw new TypeError('ZIP 항목은 문자열 또는 바이트 조각이어야 합니다.');
}

function* zipStreamSyncChunks(content) {
  function* parts(value) {
    if (typeof value === 'string' || value instanceof Uint8Array) yield value;
    else if (value && typeof value[Symbol.iterator] === 'function') { for (const part of value) yield* parts(part); }
    else throw new TypeError('ZIP 항목은 문자열 또는 바이트 조각이어야 합니다.');
  }
  const step = 1 << 18;
  let pending = [], length = 0, high = '';
  const encode = () => {
    let text = high + pending.join(''); pending = []; length = 0; high = '';
    const last = text.charCodeAt(text.length - 1);
    if (last >= 0xd800 && last <= 0xdbff) { high = text.slice(-1); text = text.slice(0, -1); }
    return text ? enc.encode(text) : null;
  };
  for (const part of parts(content)) {
    if (typeof part === 'string') {
      for (let offset = 0; offset < part.length;) {
        const end = Math.min(part.length, offset + step - length);
        pending.push(part.slice(offset, end)); length += end - offset; offset = end;
        if (length === step) { const bytes = encode(); if (bytes) yield bytes; }
      }
    } else {
      if (length) { const bytes = encode(); if (bytes) yield bytes; }
      if (high) { yield enc.encode(high); high = ''; }
      for (let offset = 0; offset < part.length; offset += step) yield part.subarray(offset, offset + step);
    }
  }
  if (length) { const bytes = encode(); if (bytes) yield bytes; }
  if (high) yield enc.encode(high);
}

/** Async input keeps surrogate pairs intact across text fragments. */
async function* zipStreamChunks(content) {
  // XML producers are synchronous and often yield one tiny cell token at a time.
  // Coalesce them before crossing an async boundary; never await every token.
  if (typeof content === 'string' || content instanceof Uint8Array || content?.[Symbol.iterator]) {
    yield* zipStreamSyncChunks(content); return;
  }
  const step = 1 << 18; let pending = [], length = 0, high = '';
  const encode = () => {
    let text = high + pending.join(''); pending = []; length = 0; high = '';
    const last = text.charCodeAt(text.length - 1);
    if (last >= 0xd800 && last <= 0xdbff) { high = text.slice(-1); text = text.slice(0, -1); }
    return text ? enc.encode(text) : null;
  };
  for await (const part of zipStreamParts(content)) {
    if (typeof part === 'string') {
      for (let offset = 0; offset < part.length;) {
        const end = Math.min(part.length, offset + step - length);
        pending.push(part.slice(offset, end)); length += end - offset; offset = end;
        if (length === step) { const bytes = encode(); if (bytes) yield bytes; }
      }
    } else {
      if (length) { const bytes = encode(); if (bytes) yield bytes; }
      if (high) { yield enc.encode(high); high = ''; }
      for (let offset = 0; offset < part.length; offset += step) yield part.subarray(offset, offset + step);
    }
  }
  if (length) { const bytes = encode(); if (bytes) yield bytes; }
  if (high) yield enc.encode(high);
}

/** The caller owns sink.close/abort. Only central records, not bodies, are retained. */
export function createZipStreamWriter(sink, { zip64 = true, compress = true, signal } = {}) {
  if (!sink || typeof sink.write !== 'function') throw new TypeError('ZIP 저장 스트림이 필요합니다.');
  if (zip64 !== true && zip64 !== false && zip64 !== 'auto') throw new TypeError('ZIP64 저장 모드가 올바르지 않습니다.');
  const patchHeaders = zip64 === 'auto' && typeof sink.seek === 'function', allowZip64 = zip64 === true || patchHeaders;
  const entries = [], names = new Set(), now = new Date();
  let usedZip64 = zip64 === true;
  let offset = 0, finished = false, writing = false, failure = null, activeIterator = null, activeReader = null;
  const check = () => {
    if (failure) throw failure;
    if (signal?.aborted) throw signal.reason ?? new DOMException('저장을 취소했습니다.', 'AbortError');
  };
  const count = n => { zipCount(n); if (!allowZip64 && n >= ZIP32_MAX) throw new RangeError('ZIP32 크기를 넘었습니다. ZIP64로 저장해 주세요.'); return n; };
  const write = async bytes => { check(); const next = count(offset + bytes.length); await sink.write(bytes); check(); offset = next; };
  return {
    get bytesWritten() { return offset; },
    async add(name, content) {
      check();
      if (finished || writing || names.has(name)) throw new Error('ZIP 항목은 순서대로 한 번씩 저장해야 합니다.');
      if (!allowZip64 && entries.length >= 0xfffe) throw new RangeError('ZIP32 항목 수를 넘었습니다. ZIP64로 저장해 주세요.');
      writing = true;
      const iterator = zipStreamChunks(content); activeIterator = iterator;
      let crc = 0xffffffff, size = 0, compressedSize = 0, sincePause = 0, reader;
      try {
        let transform, zlibWrapped = false;
        if (compress && typeof CompressionStream === 'function' && typeof ReadableStream === 'function') {
          try { transform = new CompressionStream('deflate-raw'); }
          catch { try { transform = new CompressionStream('deflate'); zlibWrapped = true; } catch {} }
        }
        const method = transform ? 8 : 0, entryOffset = offset;
        await write(patchHeaders ? zipSeekLocalRecord(name, method, 0, 0, 0, now) : zipStreamLocalRecord(name, method, zip64 === true, now));
        const take = async () => {
          check(); const next = await iterator.next(); check();
          if (!next.done) {
            crc = crcUpdate(crc, next.value); size = count(size + next.value.length); sincePause += next.value.length;
            if (sincePause >= 4 << 20) { sincePause = 0; await pause(); check(); }
          }
          return next;
        };
        const payload = async bytes => { if (!bytes.length) return; compressedSize = count(compressedSize + bytes.length); await write(bytes); };
        if (!transform) {
          for (;;) { const next = await take(); if (next.done) break; await payload(next.value); }
        } else {
          const input = new ReadableStream({
            async pull(controller) { const next = await take(); if (next.done) controller.close(); else controller.enqueue(next.value); },
            async cancel() { await iterator.return?.(); },
          }, { highWaterMark: 0 });
          reader = input.pipeThrough(transform).getReader(); activeReader = reader;
          const header = new Uint8Array(2); let headerSize = 0, tail = new Uint8Array();
          for (;;) {
            check(); const next = await reader.read(); if (next.done) break;
            let bytes = next.value;
            if (zlibWrapped) {
              if (headerSize < 2) {
                const n = Math.min(2 - headerSize, bytes.length); header.set(bytes.subarray(0, n), headerSize); headerSize += n; bytes = bytes.subarray(n);
                if (headerSize === 2 && ((header[0] & 15) !== 8 || header[0] >>> 4 > 7 || header[1] & 32 || ((header[0] << 8) | header[1]) % 31)) throw new Error('ZLIB 압축 헤더가 올바르지 않습니다.');
              }
              // Delay only the four Adler-32 bytes, even with tiny output chunks.
              const flush = tail.length + bytes.length - 4;
              if (flush > 0) {
                const fromTail = Math.min(flush, tail.length); await payload(tail.subarray(0, fromTail));
                await payload(bytes.subarray(0, flush - fromTail));
                const remain = new Uint8Array(4), old = tail.subarray(fromTail); remain.set(old); remain.set(bytes.subarray(flush - fromTail), old.length); tail = remain;
              } else if (bytes.length) { const remain = new Uint8Array(tail.length + bytes.length); remain.set(tail); remain.set(bytes, tail.length); tail = remain; }
            } else await payload(bytes);
          }
          if (zlibWrapped && (headerSize !== 2 || tail.length !== 4 || compressedSize < 2)) throw new Error('ZLIB 압축 데이터가 잘렸습니다.');
        }
        const checksum = (crc ^ 0xffffffff) >>> 0;
        const largeEntry = size >= ZIP32_MAX || compressedSize >= ZIP32_MAX;
        usedZip64 ||= largeEntry || entryOffset >= ZIP32_MAX;
        if (patchHeaders) {
          check(); await sink.seek(entryOffset); check();
          await sink.write(zipSeekLocalRecord(name, method, checksum, compressedSize, size, now)); check();
          await sink.seek(offset); check();
        } else await write(zipDescriptorRecord(checksum, compressedSize, size, zip64 === true));
        entries.push(zipCentralRecord({ name, method, crc: checksum, size, compressedSize, offset: entryOffset, zip64: zip64 === true || largeEntry, descriptor: !patchHeaders }, now)); names.add(name);
      } catch (error) {
        failure ??= error;
        if (reader) { try { await reader.cancel(error); } catch {} }
        throw error;
      } finally {
        if (reader) reader.releaseLock(); await iterator.return?.(); activeIterator = null; activeReader = null; writing = false;
      }
    },
    async finish() {
      check(); if (finished || writing) throw new Error('ZIP 저장이 완료되지 않았거나 이미 끝났습니다.');
      finished = true; const start = offset;
      try {
        for (const entry of entries) await write(entry);
        usedZip64 ||= entries.length >= 0xffff || offset - start >= ZIP32_MAX || start >= ZIP32_MAX;
        for (const end of zipEndRecords(entries.length, offset - start, start, zip64 === true)) await write(end);
        return { bytesWritten: offset, entries: entries.length, zip64: usedZip64 };
      } catch (error) { failure ??= error; throw error; }
    },
    async abort(reason = new DOMException('저장을 취소했습니다.', 'AbortError')) {
      failure ??= reason;
      if (activeReader) { try { await activeReader.cancel(reason); } catch {} }
      if (activeIterator) { try { await activeIterator.return?.(); } catch {} }
    },
  };
}
