// File System Access targets are selected by the UI before serialization.
// This module owns only the atomic byte write, not permissions or download UI.
const WRITE_CHUNK = 1 << 20;
const sourceError = () => new TypeError('저장할 파일 데이터가 올바르지 않습니다.');
const abortError = () => Object.assign(new Error('저장 중 문서가 변경되거나 작업이 취소되었습니다.'), { code: 'FILE_SAVE_ABORT' });

function check(options) {
  if (options.signal?.aborted) throw options.signal.reason instanceof Error ? options.signal.reason : abortError();
  if (options.isCurrent?.() === false) throw abortError();
  options.assertCurrent?.();
}
function bytes(value) {
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  throw sourceError();
}
export function isFileSaveSource(source) {
  return typeof source === 'function' || source instanceof Blob || source instanceof ArrayBuffer || ArrayBuffer.isView(source) ||
    !!source && (typeof source.getReader === 'function' || typeof source[Symbol.asyncIterator] === 'function' || typeof source !== 'string' && typeof source[Symbol.iterator] === 'function');
}

async function* sourceChunks(source) {
  if (source instanceof Blob) {
    for (let at = 0; at < source.size; at += WRITE_CHUNK) yield new Uint8Array(await source.slice(at, at + WRITE_CHUNK).arrayBuffer());
  } else if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) yield bytes(source);
  else if (typeof source?.getReader === 'function') {
    const reader = source.getReader(); let ended = false;
    try {
      for (;;) { const next = await reader.read(); if (next.done) { ended = true; break; } yield bytes(next.value); }
    } finally {
      if (!ended) { try { await reader.cancel(); } catch { /* Keep the write failure. */ } }
      reader.releaseLock();
    }
  } else if (source && typeof source !== 'string' && (typeof source[Symbol.asyncIterator] === 'function' || typeof source[Symbol.iterator] === 'function')) {
    for await (const chunk of source) yield bytes(chunk);
  } else throw sourceError();
}

/**
 * Writes Blob/bytes/ReadableStream/Iterable bytes, or async sink => producer(sink).
 * Every sink.write is awaited; small chunks share one bounded 1 MiB buffer.
 * Only the owner closes/aborts the native target.
 * No 32-bit/file-size cap or whole-file Blob. onProgress(bytesWritten, total?).
 * The native writable commits on close; a failed/cancelled write aborts it.
 */
export async function writeFileHandle(handle, source, options = {}) {
  if (!handle || typeof handle.createWritable !== 'function' || !isFileSaveSource(source)) throw sourceError();
  let writable, iterator, pending = null, failure = null, bytesWritten = 0, position = 0, committed = false, buffer = null, buffered = 0;
  const total = source instanceof Blob ? source.size : source instanceof ArrayBuffer || ArrayBuffer.isView(source) ? source.byteLength : undefined;
  const writeBytes = async chunk => {
    check(options);
    if (!Number.isSafeInteger(position + chunk.byteLength)) throw new RangeError('파일 크기를 정확히 계산할 수 없습니다.');
    await writable.write(chunk);
    position += chunk.byteLength;
    const extent = Math.max(bytesWritten, position);
    if (extent !== bytesWritten) { bytesWritten = extent; options.onProgress?.(bytesWritten, total); }
  };
  const flush = async () => { if (buffered) { await writeBytes(buffer.subarray(0, buffered)); buffered = 0; } };
  const sink = {
    get position() { return position + buffered; },
    get extent() { return Math.max(bytesWritten, position + buffered); },
    write(value) {
      if (failure) return Promise.reject(failure);
      if (pending) { failure = new Error('파일 데이터는 한 조각씩 순서대로 저장해야 합니다.'); return Promise.reject(failure); }
      const job = (async () => {
        const part = bytes(value);
        for (let at = 0; at < part.byteLength;) {
          check(options);
          if (buffered || part.byteLength - at < WRITE_CHUNK) {
            buffer ??= new Uint8Array(WRITE_CHUNK);
            const count = Math.min(WRITE_CHUNK - buffered, part.byteLength - at);
            buffer.set(part.subarray(at, at + count), buffered); buffered += count; at += count;
            if (buffered === WRITE_CHUNK) await flush();
          } else { await writeBytes(part.subarray(at, at + WRITE_CHUNK)); at += WRITE_CHUNK; }
        }
      })();
      pending = job;
      return job.catch(error => { failure ??= error; throw error; }).finally(() => { if (pending === job) pending = null; });
    }
  };
  try {
    check(options);
    writable = await handle.createWritable();
    check(options);
    if (typeof writable.seek === 'function') sink.seek = target => {
      if (failure) return Promise.reject(failure);
      if (pending) { failure = new Error('파일 데이터와 위치는 순서대로 변경해야 합니다.'); return Promise.reject(failure); }
      const job = (async () => {
        check(options);
        if (!Number.isSafeInteger(target) || target < 0 || target > Math.max(bytesWritten, position + buffered)) throw new RangeError('파일 위치가 저장된 범위를 벗어났습니다.');
        await flush(); check(options); await writable.seek(target); check(options); position = target;
      })();
      pending = job;
      return job.catch(error => { failure ??= error; throw error; }).finally(() => { if (pending === job) pending = null; });
    };
    if (typeof source === 'function') await source(sink);
    else {
      iterator = sourceChunks(source);
      for (;;) { check(options); const next = await iterator.next(); if (next.done) break; await sink.write(next.value); }
    }
    if (pending) await pending;
    if (failure) throw failure;
    check(options);
    await flush();
    check(options);
    if (!bytesWritten) throw new Error('저장할 파일을 생성하지 못했습니다.');
    await writable.close(); committed = true;
    return { bytesWritten };
  } catch (error) {
    if (pending) { try { await pending; } catch { /* Keep the original failure. */ } }
    if (writable && !committed) { try { await writable.abort(error); } catch { /* A failed native stream can already be closed. */ } }
    throw error;
  } finally {
    try { await iterator?.return(); } catch { /* Preserve the write result. */ }
    // A picker/createWritable failure can occur before the stream is read.
    if (!committed && typeof source?.cancel === 'function') { try { await source.cancel(); } catch { /* Already locked/failed. */ } }
  }
}
