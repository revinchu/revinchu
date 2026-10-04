// ZIP 읽기/쓰기 (xlsx 용). 외부 라이브러리 없이 inflate(RFC 1951)를 직접 구현.
// 쓰기는 무압축(stored) 방식 — 모든 스프레드시트 프로그램이 읽을 수 있음.

// ───────────── inflate ─────────────
const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];

function huffman(lengths) {
  const counts = new Uint16Array(16);
  for (const l of lengths) counts[l]++;
  counts[0] = 0;
  const offs = new Uint16Array(16);
  for (let len = 1; len < 15; len++) offs[len + 1] = offs[len] + counts[len];
  const symbols = new Uint16Array(lengths.length);
  for (let s = 0; s < lengths.length; s++) if (lengths[s]) symbols[offs[lengths[s]]++] = s;
  return { counts, symbols };
}

let FIXED = null;
function fixedTables() {
  if (!FIXED) {
    const l = new Uint8Array(288);
    l.fill(8, 0, 144); l.fill(9, 144, 256); l.fill(7, 256, 280); l.fill(8, 280, 288);
    FIXED = { lit: huffman(l), dist: huffman(new Uint8Array(30).fill(5)) };
  }
  return FIXED;
}

export function inflate(data, sizeHint = 0) {
  let pos = 0;
  let bitbuf = 0;
  let bitcnt = 0;
  let out = new Uint8Array(Math.max(1024, sizeHint));
  let outPos = 0;
  const ensure = (n) => {
    if (outPos + n <= out.length) return;
    let size = out.length * 2;
    while (size < outPos + n) size *= 2;
    const next = new Uint8Array(size);
    next.set(out);
    out = next;
  };
  const bits = (n) => {
    while (bitcnt < n) {
      if (pos >= data.length) throw new Error('압축 데이터가 손상되었습니다');
      bitbuf |= data[pos++] << bitcnt;
      bitcnt += 8;
    }
    const v = bitbuf & ((1 << n) - 1);
    bitbuf >>>= n;
    bitcnt -= n;
    return v;
  };
  const decode = (h) => {
    let code = 0;
    let first = 0;
    let index = 0;
    for (let len = 1; len < 16; len++) {
      code |= bits(1);
      const count = h.counts[len];
      if (code - count < first) return h.symbols[index + (code - first)];
      index += count;
      first += count;
      first <<= 1;
      code <<= 1;
    }
    throw new Error('잘못된 허프만 코드');
  };
  const codes = (lit, dist) => {
    for (;;) {
      const sym = decode(lit);
      if (sym < 256) {
        ensure(1);
        out[outPos++] = sym;
      } else if (sym === 256) {
        return;
      } else {
        const li = sym - 257;
        if (li >= 29) throw new Error('잘못된 길이 코드');
        const len = LEN_BASE[li] + bits(LEN_EXTRA[li]);
        const di = decode(dist);
        if (di >= 30) throw new Error('잘못된 거리 코드');
        const d = DIST_BASE[di] + bits(DIST_EXTRA[di]);
        if (d > outPos) throw new Error('잘못된 거리');
        ensure(len);
        for (let i = 0; i < len; i++, outPos++) out[outPos] = out[outPos - d];
      }
    }
  };
  let last = 0;
  while (!last) {
    last = bits(1);
    const type = bits(2);
    if (type === 0) {
      bitbuf = 0;
      bitcnt = 0;
      const len = data[pos] | (data[pos + 1] << 8);
      pos += 4;
      ensure(len);
      out.set(data.subarray(pos, pos + len), outPos);
      pos += len;
      outPos += len;
    } else if (type === 1) {
      const f = fixedTables();
      codes(f.lit, f.dist);
    } else if (type === 2) {
      const nlen = bits(5) + 257;
      const ndist = bits(5) + 1;
      const ncode = bits(4) + 4;
      const cl = new Uint8Array(19);
      for (let i = 0; i < ncode; i++) cl[CL_ORDER[i]] = bits(3);
      const clh = huffman(cl);
      const lengths = new Uint8Array(nlen + ndist);
      for (let i = 0; i < nlen + ndist;) {
        const sym = decode(clh);
        if (sym < 16) { lengths[i++] = sym; continue; }
        let rep;
        let val = 0;
        if (sym === 16) {
          if (!i) throw new Error('잘못된 반복 코드');
          val = lengths[i - 1];
          rep = 3 + bits(2);
        } else if (sym === 17) rep = 3 + bits(3);
        else rep = 11 + bits(7);
        if (i + rep > nlen + ndist) throw new Error('코드 길이가 너무 깁니다');
        while (rep--) lengths[i++] = val;
      }
      codes(huffman(lengths.subarray(0, nlen)), huffman(lengths.subarray(nlen)));
    } else {
      throw new Error('잘못된 블록 형식');
    }
  }
  return out.subarray(0, outPos);
}

// ───────────── CRC32 ─────────────
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();

export function crc32(data) {
  return crcEnd(crcUpdate(0xffffffff, data, 0, data.length));
}
function crcUpdate(c, data, from, to) {
  for (let i = from; i < to; i++) c = CRC_TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
  return c;
}
const crcEnd = (c) => (c ^ 0xffffffff) >>> 0;
const pause = () => new Promise((res) => setTimeout(res, 0));

// ───────────── ZIP ─────────────
const enc = new TextEncoder();
const dec = new TextDecoder();

/** ZIP 바이트 → { 경로: Uint8Array } (압축은 처음 읽을 때 풂) */
export function unzip(bytes) {
  const files = {};
  for (const e of zipEntries(bytes)) {
    if (e.method === 0) files[e.name] = e.raw;
    else lazyInflate(files, e);
  }
  return files;
}

const pendingZipEntries = new WeakMap();
function lazyInflate(files, e) {
  let pending = pendingZipEntries.get(files);
  if (!pending) pendingZipEntries.set(files, (pending = new Map()));
  pending.set(e.name, e);
  Object.defineProperty(files, e.name, {
    enumerable: true, configurable: true,
    get() {
      const v = inflate(e.raw, e.size);
      pending.delete(e.name);
      Object.defineProperty(files, e.name, { value: v, enumerable: true, configurable: true, writable: true });
      return v;
    },
    set(v) { pending.delete(e.name); Object.defineProperty(files, e.name, { value: v, enumerable: true, configurable: true, writable: true }); },
  });
}

/** 지정한 ZIP 항목만 해제합니다. 압축을 푼 다른 시트를 미리 보유하지 않습니다. */
export async function prepareZipEntry(files, name) {
  const descriptor = Object.getOwnPropertyDescriptor(files, name);
  const pending = pendingZipEntries.get(files), entry = pending?.get(name);
  if (!descriptor?.get || !entry) return files[name];
  if (entry.preparing) return entry.preparing;
  entry.preparing = (async () => {
    let value;
    if (typeof DecompressionStream === 'function') {
      try {
        const reader = new Blob([entry.raw]).stream().pipeThrough(new DecompressionStream('deflate-raw')).getReader();
        // Response.arrayBuffer()의 전체 출력 복사 대신 ZIP에 기록된 크기에 바로 채웁니다.
        const bytes = new Uint8Array(entry.size);
        let position = 0;
        try {
          for (;;) {
            const chunk = await reader.read();
            if (chunk.done) break;
            if (position + chunk.value.length > bytes.length) throw new Error('ZIP 항목 크기가 올바르지 않습니다.');
            bytes.set(chunk.value, position); position += chunk.value.length;
          }
          if (position !== bytes.length) throw new Error('ZIP 항목 크기가 올바르지 않습니다.');
          value = bytes;
        } catch (error) {
          try { await reader.cancel(error); } catch { /* 이미 실패한 스트림 */ }
          throw error;
        } finally { reader.releaseLock(); }
      } catch { /* 내장 압축 해제가 없는 브라우저와 같은 순수 JS 경로 */ }
    }
    value ??= inflate(entry.raw, entry.size);
    // 기다리는 동안 동기 읽기나 변환기가 이미 값을 바꿨으면 그 결과를 유지합니다.
    if (Object.getOwnPropertyDescriptor(files, name)?.get !== descriptor.get) return files[name];
    Object.defineProperty(files, name, { value, enumerable: true, configurable: true, writable: true });
    pending.delete(name);
    return value;
  })();
  return entry.preparing;
}

/** 선택한 큰 항목을 미리 풉니다. 필요한 항목만 순차로 읽을 때는 prepareZipEntry를 사용합니다. */
export async function unzipAsync(bytes, pre = (name, size) => size > 1 << 20) {
  const files = unzip(bytes), pending = pendingZipEntries.get(files);
  if (pending) for (const [name, entry] of pending) if (pre(name, entry.size)) await prepareZipEntry(files, name);
  return files;
}

function zipEntries(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  }
  if (eocd < 0) {
    // 앞은 ZIP 인데 끝의 목록이 없음: 내려받기 · 올리기가 덜 끝나 잘린 파일
    if (bytes.length > 4 && dv.getUint32(0, true) === 0x04034b50) throw new Error('파일 끝부분이 잘려 있어 열 수 없습니다 (내려받기나 올리기가 덜 끝난 파일). 원본 파일을 다시 받아 열어 주세요.');
    // 문서 보안(DRM) 프로그램이 암호화한 파일: 머리에 제품 표시가 있음 (Fasoo DRMONE, SoftCamp, MarkAny …)
    const head = String.fromCharCode(...bytes.subarray(0, Math.min(1024, bytes.length))).toUpperCase();
    if (/DRMONE|FASOO|SCDSA|SOFTCAMP|MARKANY|DOCUMENT SECURITY|ENCRYPTED AND PROTECTED/.test(head)) throw new Error('문서 보안(DRM)으로 암호화된 파일이라 열 수 없습니다. 회사 보안 프로그램에서 암호화를 해제(반출)한 파일을 열어 주세요.');
    throw new Error('ZIP 파일이 아닙니다');
  }
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  const out = [];
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP 디렉터리가 손상되었습니다');
    const method = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const local = dv.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    p += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith('/')) continue;
    const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
    const raw = bytes.subarray(start, start + compSize);
    if (method !== 0 && method !== 8) throw new Error(`지원하지 않는 압축 방식(${method}): ${name}`);
    out.push({ name, method, raw, size });
  }
  return out;
}

/** 문자열/바이트/반복자 조각을 제한된 UTF-8 청크로 읽습니다. */
function* zipByteChunks(content) {
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

const ZIP32_MAX = 0xffffffff;
function zipSize(n) {
  if (!Number.isSafeInteger(n) || n < 0 || n >= ZIP32_MAX) throw new RangeError('ZIP32로 저장할 수 있는 파일 크기를 넘었습니다.');
  return n;
}
function zipStoredEntry(name, content) {
  let crc = 0xffffffff, size = 0;
  const body = [];
  for (const bytes of zipByteChunks(content)) { crc = crcUpdate(crc, bytes, 0, bytes.length); size = zipSize(size + bytes.length); body.push(bytes); }
  return { name, crc: crcEnd(crc), size, compressedSize: size, method: 0, body };
}

/** { 경로: 문자열|Uint8Array|중첩 조각 배열 } → ZIP 바이트 (무압축).
 * consume은 호출자가 소유권을 넘긴 내부 항목만 완료 후 해제합니다. 기본 입력은 불변입니다. */
export function zip(entries, { consume = false } = {}) {
  const list = [];
  for (const name of Object.keys(entries)) {
    list.push(zipStoredEntry(name, entries[name]));
    if (consume) delete entries[name];
  }
  return zipBuild(list);
}

async function zipAsyncEntry(name, content) {
  let transform, zlibWrapped = false;
  if (typeof CompressionStream === 'function' && typeof ReadableStream === 'function') {
    // Try native zlib when only raw deflate is unsupported. Runtime failures
    // still propagate; consuming the source again could silently lose data.
    try { transform = new CompressionStream('deflate-raw'); }
    catch { try { transform = new CompressionStream('deflate'); zlibWrapped = true; } catch { /* 저장 방식으로 대체 */ } }
  }
  const iterator = zipByteChunks(content);
  let crc = 0xffffffff, size = 0, sincePause = 0;
  const take = async () => {
    const next = iterator.next();
    if (!next.done) {
      crc = crcUpdate(crc, next.value, 0, next.value.length);
      size = zipSize(size + next.value.length); sincePause += next.value.length;
      if (sincePause >= 4 << 20) { sincePause = 0; await pause(); }
    }
    return next;
  };
  const body = [];
  let compressedSize = 0;
  if (!transform) {
    try { for (;;) { const next = await take(); if (next.done) break; body.push(next.value); } }
    finally { iterator.return?.(); }
    return { name, crc: crcEnd(crc), size, compressedSize: size, method: 0, body };
  }
  const input = new ReadableStream({
    async pull(controller) { const next = await take(); if (next.done) controller.close(); else controller.enqueue(next.value); },
    cancel() { iterator.return?.(); },
  });
  const reader = input.pipeThrough(transform).getReader();
  const zlibHeader = new Uint8Array(2); let headerSize = 0;
  try {
    for (;;) {
      const next = await reader.read(); if (next.done) break;
      let bytes = next.value;
      if (zlibWrapped && headerSize < 2) {
        const count = Math.min(2 - headerSize, bytes.length);
        zlibHeader.set(bytes.subarray(0, count), headerSize); headerSize += count; bytes = bytes.subarray(count);
        if (headerSize === 2 && ((zlibHeader[0] & 15) !== 8 || zlibHeader[0] >>> 4 > 7 || zlibHeader[1] & 32 || ((zlibHeader[0] << 8) | zlibHeader[1]) % 31)) throw new Error('ZLIB 압축 헤더가 올바르지 않습니다.');
      }
      if (bytes.length) { compressedSize = zipSize(compressedSize + bytes.length); body.push(bytes); }
    }
    if (zlibWrapped) {
      if (headerSize !== 2 || compressedSize < 6) throw new Error('ZLIB 압축 데이터가 잘렸습니다.');
      // ZIP stores RFC1951 payload only. Trim the Adler-32 trailer in place,
      // even when its four bytes span several native output chunks.
      for (let remaining = 4; remaining;) {
        const last = body.pop(), remove = Math.min(remaining, last.length);
        if (remove < last.length) body.push(last.subarray(0, last.length - remove));
        remaining -= remove;
      }
      compressedSize -= 4;
    }
  } catch (error) {
    try { await reader.cancel(error); } catch { /* 오류가 난 스트림도 잠금은 해제합니다. */ }
    throw error;
  } finally { reader.releaseLock(); iterator.return?.(); }
  return { name, crc: crcEnd(crc), size, compressedSize, method: 8, body };
}

/** 한 항목을 완료한 뒤에 다음 항목을 생성합니다. 입력 XML 전체를 보관하지 않습니다. */
export function createZipWriter() {
  const list = [], names = new Set(); let finished = false;
  return {
    add(name, content) {
      if (finished || names.has(name)) throw new Error('ZIP 항목을 중복 저장할 수 없습니다.');
      list.push(zipStoredEntry(name, content)); names.add(name);
    },
    finish({ blob = false, first = null } = {}) {
      if (finished) throw new Error('ZIP 저장이 이미 완료되었습니다.');
      finished = true; return zipBuild(zipFirst(list, first), blob);
    },
  };
}

/** 비동기 ZIP: 항목은 하나씩 압축하고 마지막에 Blob으로 반환할 수 있습니다. */
export function createZipAsyncWriter() {
  const list = [], names = new Set(); let finished = false, writing = false;
  return {
    async add(name, content) {
      if (finished || writing || names.has(name)) throw new Error('ZIP 항목은 순서대로 한 번씩 저장해야 합니다.');
      writing = true;
      try { list.push(await zipAsyncEntry(name, content)); names.add(name); }
      finally { writing = false; }
    },
    finish({ blob = false, first = null } = {}) {
      if (finished || writing) throw new Error('ZIP 저장이 완료되지 않았거나 이미 끝났습니다.');
      finished = true; return zipBuild(zipFirst(list, first), blob);
    },
  };
}
function zipFirst(list, first) {
  const index = first ? list.findIndex(entry => entry.name === first) : -1;
  if (index > 0) list.unshift(list.splice(index, 1)[0]);
  return list;
}

/** 전체 XML/UTF-8 입력 복사 없이 청크를 압축하고 CRC와 크기를 누적합니다. */
export async function zipAsync(entries, onProgress, { consume = false, blob = false } = {}) {
  const writer = createZipAsyncWriter(), names = Object.keys(entries);
  for (let i = 0; i < names.length; i++) {
    const name = names[i];
    await writer.add(name, entries[name]);
    if (consume) delete entries[name];
    onProgress?.((i + 1) / names.length);
    await pause();
  }
  return writer.finish({ blob });
}

function zipBuild(list, blob = false) {
  if (list.length >= 0xffff) throw new RangeError('ZIP32로 저장할 수 있는 항목 수를 넘었습니다.');
  const parts = [], central = [];
  let offset = 0;
  const now = new Date();
  const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  for (const { name, crc, size, compressedSize, method, body } of list) {
    const nameBytes = enc.encode(name);
    if (nameBytes.length > 0xffff) throw new RangeError('ZIP 항목 이름이 너무 깁니다.');
    zipSize(size); zipSize(compressedSize); zipSize(offset);
    const local = new Uint8Array(30 + nameBytes.length), lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true); lv.setUint16(4, 20, true); lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, method, true); lv.setUint16(10, time, true); lv.setUint16(12, date, true);
    lv.setUint32(14, crc, true); lv.setUint32(18, compressedSize, true); lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true); local.set(nameBytes, 30);
    const cd = new Uint8Array(46 + nameBytes.length), cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true); cv.setUint16(4, 20, true); cv.setUint16(6, 20, true); cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, method, true); cv.setUint16(12, time, true); cv.setUint16(14, date, true);
    cv.setUint32(16, crc, true); cv.setUint32(20, compressedSize, true); cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true); cv.setUint32(42, offset, true); cd.set(nameBytes, 46);
    parts.push(local); for (const chunk of body) parts.push(chunk);
    central.push(cd); offset = zipSize(offset + local.length + compressedSize);
  }
  const cdSize = zipSize(central.reduce((s, c) => s + c.length, 0));
  const end = new Uint8Array(22), ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true); ev.setUint16(8, central.length, true); ev.setUint16(10, central.length, true);
  ev.setUint32(12, cdSize, true); ev.setUint32(16, offset, true);
  zipSize(offset + cdSize + end.length);
  if (blob) return new Blob([...parts, ...central, end], { type: 'application/zip' });
  const out = new Uint8Array(offset + cdSize + end.length);
  let p = 0;
  for (const part of parts) { out.set(part, p); p += part.length; }
  for (const part of central) { out.set(part, p); p += part.length; }
  out.set(end, p);
  return out;
}

export const textOf = (bytes) => (bytes ? dec.decode(bytes) : null);
