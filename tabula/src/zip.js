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

/** { 경로: Uint8Array|string } → ZIP 바이트 (무압축) */
export function zip(entries) {
  const list = Object.entries(entries).map(([name, content]) => {
    const data = typeof content === 'string' ? enc.encode(content) : content;
    return { name, data, crc: crc32(data), size: data.length, method: 0, body: data };
  });
  return zipBuild(list);
}

/** 큰 항목은 브라우저 내장 압축(deflate)으로 줄여서 ZIP 만들기 — 압축은 화면을 멈추지 않음 */
export async function zipAsync(entries, onProgress) {
  const native = typeof CompressionStream === 'function';
  const list = [];
  const all = Object.entries(entries);
  for (let i = 0; i < all.length; i++) {
    const [name, content] = all[i];
    const data = typeof content === 'string' ? await encodeAsync(content) : content;
    const e = { name, crc: await crc32Async(data), size: data.length, method: 0, body: data };
    if (native && data.length > 4096) {
      try {
        const stream = new Blob([data]).stream().pipeThrough(new CompressionStream('deflate-raw'));
        const comp = new Uint8Array(await new Response(stream).arrayBuffer());
        if (comp.length < data.length) { e.method = 8; e.body = comp; }
      } catch { /* 무압축 */ }
    }
    list.push(e);
    onProgress?.((i + 1) / all.length);
    await new Promise((res) => setTimeout(res, 0));
  }
  return zipBuild(list);
}

function zipBuild(list) {
  const parts = [];
  const central = [];
  let offset = 0;
  const now = new Date();
  const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
  const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
  for (const { name, crc, size, method, body } of list) {
    const nameBytes = enc.encode(name);
    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0x0800, true);
    lv.setUint16(8, method, true);
    lv.setUint16(10, time, true);
    lv.setUint16(12, date, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, body.length, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    const cd = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(cd.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0x0800, true);
    cv.setUint16(10, method, true);
    cv.setUint16(12, time, true);
    cv.setUint16(14, date, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, body.length, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint32(42, offset, true);
    cd.set(nameBytes, 46);
    parts.push(local, body);
    central.push(cd);
    offset += local.length + body.length;
  }
  const cdSize = central.reduce((s, c) => s + c.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(8, central.length, true);
  ev.setUint16(10, central.length, true);
  ev.setUint32(12, cdSize, true);
  ev.setUint32(16, offset, true);
  const all = [...parts, ...central, end];
  const out = new Uint8Array(all.reduce((s, a) => s + a.length, 0));
  let p = 0;
  for (const a of all) { out.set(a, p); p += a.length; }
  return out;
}

export const textOf = (bytes) => (bytes ? dec.decode(bytes) : null);
