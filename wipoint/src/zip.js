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

/** 큰 데이터의 CRC 를 조금씩 (화면이 멈추지 않게) */
async function crc32Async(data) {
  let c = 0xffffffff;
  const step = 4 << 20;
  for (let i = 0; i < data.length; i += step) {
    c = crcUpdate(c, data, i, Math.min(data.length, i + step));
    if (i + step < data.length) await pause();
  }
  return crcEnd(c);
}

/** 긴 문자열 → UTF-8 바이트를 조각으로 (한 번에 수십 MB 를 바꾸면 화면이 멈춤) */
async function encodeAsync(text) {
  const step = 4 << 20;
  if (text.length <= step) return enc.encode(text);
  const parts = [];
  let n = 0;
  for (let i = 0; i < text.length; i += step) {
    let end = Math.min(text.length, i + step);
    const code = text.charCodeAt(end - 1);
    if (end < text.length && code >= 0xd800 && code <= 0xdbff) end++; // 서로게이트 쌍을 자르지 않음
    const b = enc.encode(text.slice(i, end));
    parts.push(b);
    n += b.length;
    i = end - step;
    await pause();
  }
  const out = new Uint8Array(n);
  let p = 0;
  for (const b of parts) { out.set(b, p); p += b.length; }
  return out;
}

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

function lazyInflate(files, e) {
  Object.defineProperty(files, e.name, {
    enumerable: true, configurable: true,
    get() {
      const v = inflate(e.raw, e.size);
      Object.defineProperty(files, e.name, { value: v, enumerable: true, configurable: true, writable: true });
      return v;
    },
    set(v) { Object.defineProperty(files, e.name, { value: v, enumerable: true, configurable: true, writable: true }); },
  });
}

/** 큰 항목은 브라우저 내장 압축 해제(DecompressionStream)로 미리 풂 — 나머지는 unzip 과 같음 */
export async function unzipAsync(bytes, pre = (name, size) => size > 1 << 20) {
  const files = {};
  const native = typeof DecompressionStream === 'function';
  for (const e of zipEntries(bytes)) {
    if (e.method === 0) { files[e.name] = e.raw; continue; }
    if (native && pre(e.name, e.size)) {
      try {
        const stream = new Blob([e.raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        files[e.name] = new Uint8Array(await new Response(stream).arrayBuffer());
        continue;
      } catch { /* 아래 방식으로 */ }
    }
    lazyInflate(files, e);
  }
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
  const out = [];
  for (const e of centralEntries(bytes, dv.getUint32(eocd + 16, true), dv.getUint16(eocd + 10, true))) {
    const start = e.local + 30 + dv.getUint16(e.local + 26, true) + dv.getUint16(e.local + 28, true);
    out.push({ name: e.name, method: e.method, raw: bytes.subarray(start, start + e.compSize), size: e.size });
  }
  return out;
}

const MAX32 = 0xffffffff;
/** 중앙 디렉터리 항목들 (ZIP64 extra 포함) → [{ name, method, compSize, size, local }] */
function centralEntries(bytes, p, count) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const out = [];
  for (let n = 0; n < count && p + 46 <= bytes.length; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP 디렉터리가 손상되었습니다');
    const method = dv.getUint16(p + 10, true);
    let compSize = dv.getUint32(p + 20, true);
    let size = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    let local = dv.getUint32(p + 42, true);
    const name = dec.decode(bytes.subarray(p + 46, p + 46 + nameLen));
    if (size === MAX32 || compSize === MAX32 || local === MAX32) {
      // ZIP64: 원래 크기 · 압축 크기 · 위치가 extra(0x0001)에 8바이트씩 (값이 MAX 인 것만, 이 순서로)
      for (let q = p + 46 + nameLen, end = q + extraLen; q + 4 <= end;) {
        const id = dv.getUint16(q, true);
        const len = dv.getUint16(q + 2, true);
        if (id === 1) {
          let k = q + 4;
          const u64 = () => { const v = Number(dv.getBigUint64(k, true)); k += 8; return v; };
          if (size === MAX32) size = u64();
          if (compSize === MAX32) compSize = u64();
          if (local === MAX32) local = u64();
        }
        q += 4 + len;
      }
    }
    p += 46 + nameLen + extraLen + commentLen;
    if (name.endsWith('/')) continue;
    if (method !== 0 && method !== 8) throw new Error(`지원하지 않는 압축 방식(${method}): ${name}`);
    out.push({ name, method, compSize, size, local, nameLen });
  }
  return out;
}

/**
 * 큰 파일 (수 GB): 파일 전체를 메모리에 올리지 않고 끝의 목록만 읽은 뒤 항목을 조각으로 읽음.
 *   lazy(name, size) 가 참인 항목(동영상 · 큰 그림)은 Blob 으로 둠 — 무압축이면 파일 조각 그대로(복사 없음),
 *   압축이면 브라우저 스트림으로 풂. 나머지(XML 등)는 Uint8Array (압축은 처음 읽을 때 풂).
 *   onProgress(0..1) — 항목을 읽을 때마다
 */
export async function openZipFile(file, { lazy = () => false, onProgress } = {}) {
  const read = async (a, b) => new Uint8Array(await file.slice(a, b).arrayBuffer());
  const total = file.size;
  const tailStart = Math.max(0, total - 65557 - 20);
  const tail = await read(tailStart, total);
  const tv = new DataView(tail.buffer);
  let eocd = -1;
  for (let i = tail.length - 22; i >= 0; i--) if (tv.getUint32(i, true) === 0x06054b50) { eocd = i; break; }
  if (eocd < 0) { zipEntries(await read(0, Math.min(4096, total))); throw new Error('ZIP 파일이 아닙니다'); }
  let count = tv.getUint16(eocd + 10, true);
  let cdSize = tv.getUint32(eocd + 12, true);
  let cdOff = tv.getUint32(eocd + 16, true);
  if (count === 0xffff || cdSize === MAX32 || cdOff === MAX32) {
    const loc = eocd - 20;
    if (loc >= 0 && tv.getUint32(loc, true) === 0x07064b50) {
      const at = Number(tv.getBigUint64(loc + 8, true));
      const r = await read(at, at + 56);
      const v = new DataView(r.buffer);
      if (v.getUint32(0, true) === 0x06064b50) { count = Number(v.getBigUint64(32, true)); cdSize = Number(v.getBigUint64(40, true)); cdOff = Number(v.getBigUint64(48, true)); }
    }
  }
  const list = centralEntries(await read(cdOff, cdOff + cdSize), 0, count);
  const native = typeof DecompressionStream === 'function';
  const files = {};
  let done = 0;
  const one = async (e) => {
    const big = lazy(e.name, e.size);
    // 지역 머리글의 extra 길이는 중앙 목록과 다를 수 있어 넉넉히 읽음 (작은 항목은 한 번에)
    const headLen = 30 + e.nameLen + 1024;
    const chunk = await read(e.local, e.local + headLen + (big ? 0 : e.compSize));
    const lv = new DataView(chunk.buffer);
    if (lv.getUint32(0, true) !== 0x04034b50) throw new Error('ZIP 항목이 손상되었습니다');
    const off = 30 + lv.getUint16(26, true) + lv.getUint16(28, true);
    const start = e.local + off;
    if (big) {
      const part = file.slice(start, start + e.compSize);
      files[e.name] = e.method === 0 ? part : await new Response(part.stream().pipeThrough(new DecompressionStream('deflate-raw'))).blob();
    } else {
      const raw = off + e.compSize <= chunk.length ? chunk.subarray(off, off + e.compSize) : await read(start, start + e.compSize);
      if (e.method === 0) files[e.name] = raw;
      else if (native && e.size > 1 << 20) files[e.name] = new Uint8Array(await new Response(new Blob([raw]).stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer());
      else lazyInflate(files, { name: e.name, raw, size: e.size });
    }
    onProgress?.(++done / list.length);
  };
  // 조금씩 동시에 (디스크 · 화면이 멈추지 않게)
  for (let i = 0; i < list.length; i += 16) await Promise.all(list.slice(i, i + 16).map(one));
  return files;
}

/** { 경로: Uint8Array|string } → ZIP 바이트 (무압축) */
export function zip(entries) {
  const list = Object.entries(entries).map(([name, content]) => {
    const data = typeof content === 'string' ? enc.encode(content) : content;
    return { name, data, crc: crc32(data), size: data.length, method: 0, body: data };
  });
  return zipBuild(list);
}

/**
 * 큰 항목은 브라우저 내장 압축(deflate)으로 줄여서 ZIP 만들기 — 압축은 화면을 멈추지 않음.
 * 항목은 문자열 · Uint8Array · Blob (동영상 등 큰 미디어: 조각으로 CRC 만 계산하고 복사 없이 그대로 담음).
 * 결과는 Blob (4GB 넘으면 ZIP64).
 */
export async function zipAsync(entries, onProgress) {
  const native = typeof CompressionStream === 'function';
  const list = [];
  const all = Object.entries(entries);
  for (let i = 0; i < all.length; i++) {
    const [name, content] = all[i];
    if (typeof Blob === 'function' && content instanceof Blob) {
      let c = 0xffffffff;
      const step = 16 << 20;
      for (let a = 0; a < content.size; a += step) {
        const part = new Uint8Array(await content.slice(a, Math.min(content.size, a + step)).arrayBuffer());
        c = crcUpdate(c, part, 0, part.length);
        onProgress?.((i + a / content.size) / all.length);
        await pause();
      }
      list.push({ name, crc: crcEnd(c), size: content.size, method: 0, body: content });
      continue;
    }
    const data = typeof content === 'string' ? await encodeAsync(content) : content;
    const e = { name, crc: await crc32Async(data), size: data.length, method: 0, body: data };
    if (native && data.length > 4096 && !/\.(jpe?g|png|gif|webp|mp4|m4v|mov|webm|mp3|m4a|aac|ogg|zip|xlsx|fntdata)$/i.test(name)) {
      try {
        const stream = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
        const comp = new Uint8Array(await new Response(stream).arrayBuffer());
        if (comp.length < data.length) { e.method = 8; e.body = comp; }
      } catch { /* 무압축 */ }
    }
    list.push(e);
    onProgress?.((i + 1) / all.length);
    await pause();
  }
  return new Blob(zipParts(list));
}

/** 동기 ZIP (작은 문서 · 테스트) */
function zipBuild(list) {
  const all = zipParts(list);
  const out = new Uint8Array(all.reduce((s, a) => s + a.length, 0));
  let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

export function zipParts(list) {
  const parts = [];
  const central = [];
  let offset = 0;
  const now = new Date();
  const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  const lenOf = (b) => b.length ?? b.size;
  let z64any = false;
  for (const { name, crc, size, method, body } of list) {
    const nameBytes = enc.encode(name);
    const comp = lenOf(body);
    const bigSize = size >= MAX32 || comp >= MAX32;
    const bigOff = offset >= MAX32;
    // ZIP64 extra: 지역 머리글은 크기 둘, 중앙 목록은 필요한 것만 (크기 · 크기 · 위치)
    const lx = bigSize ? 20 : 0;
    const cxVals = [...(bigSize ? [size, comp] : []), ...(bigOff ? [offset] : [])];
    const cx = cxVals.length ? 4 + cxVals.length * 8 : 0;
    if (bigSize || bigOff) z64any = true;
    const local = new Uint8Array(30 + nameBytes.length + lx);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, bigSize ? 45 : 20, true);
    lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, method, true);
    lv.setUint16(10, time, true);
    lv.setUint16(12, date, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, bigSize ? MAX32 : comp, true);
    lv.setUint32(22, bigSize ? MAX32 : size, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, lx, true);
    local.set(nameBytes, 30);
    if (bigSize) {
      const q = 30 + nameBytes.length;
      lv.setUint16(q, 1, true); lv.setUint16(q + 2, 16, true);
      lv.setBigUint64(q + 4, BigInt(size), true); lv.setBigUint64(q + 12, BigInt(comp), true);
    }
    const cd = new Uint8Array(46 + nameBytes.length + cx);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, cx ? 45 : 20, true);
    cv.setUint16(6, cx ? 45 : 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, method, true);
    cv.setUint16(12, time, true);
    cv.setUint16(14, date, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, bigSize ? MAX32 : comp, true);
    cv.setUint32(24, bigSize ? MAX32 : size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, cx, true);
    cv.setUint32(42, bigOff ? MAX32 : offset, true);
    cd.set(nameBytes, 46);
    if (cx) {
      const q = 46 + nameBytes.length;
      cv.setUint16(q, 1, true); cv.setUint16(q + 2, cxVals.length * 8, true);
      cxVals.forEach((v, k) => cv.setBigUint64(q + 4 + k * 8, BigInt(v), true));
    }
    parts.push(local, body);
    central.push(cd);
    offset += local.length + comp;
  }
  const cdSize = central.reduce((s, c) => s + c.length, 0);
  const tailParts = [];
  if (z64any || offset >= MAX32 || central.length >= 0xffff) {
    const e64 = new Uint8Array(56 + 20);
    const v = new DataView(e64.buffer);
    v.setUint32(0, 0x06064b50, true);
    v.setBigUint64(4, 44n, true);
    v.setUint16(12, 45, true); v.setUint16(14, 45, true);
    v.setBigUint64(24, BigInt(central.length), true);
    v.setBigUint64(32, BigInt(central.length), true);
    v.setBigUint64(40, BigInt(cdSize), true);
    v.setBigUint64(48, BigInt(offset), true);
    v.setUint32(56, 0x07064b50, true);
    v.setBigUint64(64, BigInt(offset + cdSize), true);
    v.setUint32(72, 1, true);
    tailParts.push(e64);
  }
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, Math.min(central.length, 0xffff), true);
  ev.setUint16(10, Math.min(central.length, 0xffff), true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, Math.min(offset, MAX32), true);
  return [...parts, ...central, ...tailParts, end];
}

export const textOf = (bytes) => (bytes ? dec.decode(bytes) : null);
