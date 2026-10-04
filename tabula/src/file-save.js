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

// Cancellation must settle a paused read without waiting for another chunk.
// Late settlements are observed and cannot resume writing after cancellation.
function readWithAbort(promise, options) {
  const signal = options.signal;
  if (!signal) return promise;
  return new Promise((resolve, reject) => {
    const abort = () => {
      signal.removeEventListener('abort', abort);
      reject(signal.reason instanceof Error ? signal.reason : abortError());
    };
    signal.addEventListener('abort', abort, { once: true });
    Promise.resolve(promise).then(resolve, reject).finally(() => signal.removeEventListener('abort', abort));
    if (signal.aborted) abort();
  });
}
// A producer's cancellation hook can stall or reject. Invoke it, but do not let
// it hide a disk failure or delay aborting the native temporary file.
function cleanup(action) { try { Promise.resolve(action()).catch(() => {}); } catch { /* Preserve the original outcome. */ } }

async function* sourceChunks(source, options) {
  if (source instanceof Blob) {
    for (let at = 0; at < source.size; at += WRITE_CHUNK) yield new Uint8Array(await readWithAbort(source.slice(at, at + WRITE_CHUNK).arrayBuffer(), options));
  } else if (source instanceof ArrayBuffer || ArrayBuffer.isView(source)) yield bytes(source);
  else if (typeof source?.getReader === 'function') {
    const reader = source.getReader(); let ended = false;
    try {
      for (;;) { const next = await readWithAbort(reader.read(), options); check(options); if (next.done) { ended = true; break; } yield bytes(next.value); }
    } finally {
      if (!ended) cleanup(() => reader.cancel(options.signal?.reason));
      reader.releaseLock();
    }
  } else if (source && typeof source !== 'string' && (typeof source[Symbol.asyncIterator] === 'function' || typeof source[Symbol.iterator] === 'function')) {
    const iterator = typeof source[Symbol.asyncIterator] === 'function' ? source[Symbol.asyncIterator]() : source[Symbol.iterator]();
    let ended = false;
    try {
      for (;;) { const next = await readWithAbort(iterator.next(), options); check(options); if (next.done) { ended = true; break; } yield bytes(next.value); }
    } finally { if (!ended) cleanup(() => iterator.return?.()); }
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
    if (typeof source === 'function') await readWithAbort(source(sink), options);
    else {
      iterator = sourceChunks(source, options);
      for (;;) { check(options); const next = await iterator.next(); if (next.done) break; await sink.write(next.value); }
    }
    if (pending) await pending;
    if (failure) throw failure;
    check(options);
    await flush();
    check(options);
    if (!bytesWritten) throw new Error('저장할 파일을 생성하지 못했습니다.');
    options.onCommitting?.();
    await writable.close(); committed = true;
    return { bytesWritten };
  } catch (error) {
    failure ??= error;
    if (pending) { try { await pending; } catch { /* Keep the original failure. */ } }
    if (writable && !committed) { try { await writable.abort(error); } catch { /* A failed native stream can already be closed. */ } }
    throw error;
  } finally {
    try { await iterator?.return(); } catch { /* Preserve the write result. */ }
    // A picker/createWritable failure can occur before the stream is read.
    if (!committed && typeof source?.cancel === 'function') cleanup(() => source.cancel(options.signal?.reason));
  }
}
