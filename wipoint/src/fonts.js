// 포함된 글꼴 (DOM 없음): pptx 의 ppt/fonts/*.fntdata = EOT(Embedded OpenType)
// EOT 머리글 → (XOR 0x50 풀기) → MTX 압축이면 LZCOMP(적응 허프만 + LZ77) 세 덩어리 풀기 → CTF(압축 글리프)를 TrueType 로 되돌림
// 형식 설명: W3C MTX Submission (MicroType Express) · EOT Submission. 알고리즘을 그대로 따라 새로 작성함.

// ───────────── 비트 읽기 (상위 비트부터) ─────────────
class BitIn {
  constructor(b) { this.b = b; this.i = 0; this.buf = 0; this.left = 0; }
  bit() {
    if (this.left === 0) {
      if (this.i >= this.b.length) throw new Error('MTX: 데이터가 끝났습니다');
      this.buf = this.b[this.i++];
      this.left = 8;
    }
    this.left--;
    return (this.buf >> this.left) & 1;
  }
  bits(n) { let v = 0; for (let k = 0; k < n; k++) v = (v << 1) | this.bit(); return v >>> 0; }
}

// ───────────── 적응 허프만 (MTX AHUFF) ─────────────
class AdaptiveHuffman {
  constructor(input, range) {
    this.in = input;
    const n = 2 * range;
    this.up = new Int32Array(n);
    this.left = new Int32Array(n);
    this.right = new Int32Array(n);
    this.code = new Int32Array(n).fill(-1);
    this.weight = new Float64Array(n);
    this.index = new Int32Array(range);
    for (let i = 2; i < n; i++) { this.up[i] = i >> 1; this.weight[i] = 1; }
    for (let i = 1; i < range; i++) { this.left[i] = 2 * i; this.right[i] = 2 * i + 1; }
    for (let i = 0; i < range; i++) { this.code[range + i] = i; this.left[range + i] = -1; this.right[range + i] = -1; this.index[i] = range + i; }
    // 안쪽 마디 무게 = 두 자식의 합 (자식 번호가 항상 더 큼)
    for (let i = range - 1; i >= 1; i--) this.weight[i] = this.weight[this.left[i]] + this.weight[this.right[i]];
    if (range > 256 && range < 512) {
      // 글자 모델: 복사 기호와 DUP 기호를 처음부터 자주 나오는 것으로
      this.bump(this.index[256]);
      this.bump(this.index[257]);
      for (let i = 0; i < 12; i++) this.bump(this.index[range - 3]);
      for (let i = 0; i < 6; i++) this.bump(this.index[range - 2]);
    } else {
      for (let k = 0; k < 2; k++) for (let i = 0; i < range; i++) this.bump(this.index[i]);
    }
  }
  swap(a, b) {
    const { up, left, right, code, weight, index } = this;
    const ua = up[a];
    const ub = up[b];
    let t = left[a]; left[a] = left[b]; left[b] = t;
    t = right[a]; right[a] = right[b]; right[b] = t;
    t = code[a]; code[a] = code[b]; code[b] = t;
    t = weight[a]; weight[a] = weight[b]; weight[b] = t;
    up[a] = ua;
    up[b] = ub;
    for (const x of [a, b]) {
      if (code[x] < 0) { up[left[x]] = x; up[right[x]] = x; } else index[code[x]] = x;
    }
  }
  /** 마디 a 와 조상들의 무게 +1 (형제 규칙을 지키도록 같은 무게의 가장 앞 마디와 바꿈) */
  bump(a) {
    const { weight, up } = this;
    while (a !== 1) {
      const w = weight[a];
      let b = a - 1;
      if (weight[b] === w) {
        do b--; while (weight[b] === w);
        b++;
        if (b > 1) { this.swap(a, b); a = b; }
      }
      weight[a] = w + 1;
      a = up[a];
    }
    weight[1]++;
  }
  read() {
    let a = 1;
    let sym;
    do { a = this.in.bit() ? this.right[a] : this.left[a]; sym = this.code[a]; } while (sym < 0);
    this.bump(a);
    return sym;
  }
}

// ───────────── LZCOMP ─────────────
const PRELOAD = 2 * 32 * 96 + 4 * 256;
function lzcomp(data, version) {
  const bin = new BitIn(data);
  const runLength = version === 1 ? false : !!bin.bit();
  const dist = new AdaptiveHuffman(bin, 8);
  const len = new AdaptiveHuffman(bin, 8);
  const outLen = bin.bits(24);
  let ranges = 1;
  while (1 + (2 ** (3 * ranges)) - 1 < outLen) ranges++;
  const DUP2 = 256 + 8 * ranges;
  const DUP4 = DUP2 + 1;
  const DUP6 = DUP2 + 2;
  const sym = new AdaptiveHuffman(bin, DUP6 + 1);
  // 미리 채운 창 (앞부분도 복사 항목을 쓸 수 있게)
  const buf = new Uint8Array(PRELOAD + outLen);
  let i = 0;
  for (let k = 0; k < 32; k++) for (let j = 0; j < 96; j++) { buf[i++] = k; buf[i++] = j; }
  for (let j = 0; i < PRELOAD && j < 256; j++) { buf[i++] = j; buf[i++] = j; buf[i++] = j; buf[i++] = j; }
  // 출력 (run-length 풀기 포함)
  let out = new Uint8Array(Math.max(16, outLen));
  let n = 0;
  const push = (v) => { if (n >= out.length) { const o = new Uint8Array(out.length * 2); o.set(out); out = o; } out[n++] = v; };
  let rlState = 100;
  let escape = 0;
  let count = 0;
  const emit = runLength ? (v) => {
    if (rlState === 0) { if (v === escape) rlState = 1; else push(v); } else if (rlState === 1) { count = v; if (count === 0) { push(escape); rlState = 0; } else rlState = 2; } else if (rlState === 2) { for (let k = 0; k < count; k++) push(v); rlState = 0; } else { escape = v; rlState = 0; }
  } : push;
  let pos = PRELOAD;
  const end = PRELOAD + outLen;
  while (pos < end) {
    const s = sym.read();
    if (s < 256) { buf[pos++] = s; emit(s); continue; }
    if (s === DUP2 || s === DUP4 || s === DUP6) { const v = buf[pos - (s === DUP2 ? 2 : s === DUP4 ? 4 : 6)]; buf[pos++] = v; emit(v); continue; }
    // 복사 항목: 길이 (3비트 묶음, 위 비트 = 계속) · 거리 (3비트 × 범위 수)
    let bits = s - 256;
    const nr = Math.floor(bits / 8) + 1;
    bits %= 8;
    let length = 0;
    for (;;) {
      const done = (bits & 4) === 0;
      length = length * 4 + (bits & 3);
      if (done) break;
      bits = len.read();
    }
    length += 2;
    let distance = 0;
    for (let k = 0; k < nr; k++) distance = distance * 8 + dist.read();
    distance += 1;
    if (distance >= 512) length++;
    const start = pos - distance - length + 1;
    if (start < 0 || pos + length > end) throw new Error('MTX: 잘못된 복사 항목');
    for (let k = 0; k < length; k++) { const v = buf[start + k]; buf[pos++] = v; emit(v); }
  }
  return out.subarray(0, n);
}

/** MTX 덩어리 → CTF 세 흐름 */
export function unpackMtx(b) {
  const version = b[0];
  const u24 = (o) => (b[o] << 16) | (b[o + 1] << 8) | b[o + 2];
  const off = [10, u24(4), u24(7)];
  const ends = [off[1], off[2], b.length];
  return off.map((o, k) => {
    if (o > b.length || ends[k] > b.length || ends[k] < o) throw new Error('MTX: 덩어리 위치가 잘못되었습니다');
    return lzcomp(b.subarray(o, ends[k]), version);
  });
}

// ───────────── CTF → TrueType ─────────────
class Rd {
  constructor(b, p = 0) { this.b = b; this.p = p; }
  u8() { if (this.p >= this.b.length) throw new Error('CTF: 데이터가 끝났습니다'); return this.b[this.p++]; }
  u16() { return (this.u8() << 8) | this.u8(); }
  s16() { const v = this.u16(); return v & 0x8000 ? v - 0x10000 : v; }
  u32() { return ((this.u16() << 16) | this.u16()) >>> 0; }
  bytes(n) { if (this.p + n > this.b.length) throw new Error('CTF: 데이터가 끝났습니다'); const r = this.b.subarray(this.p, this.p + n); this.p += n; return r; }
  u255() { const c = this.u8(); if (c === 253) return this.u16(); if (c === 255) return 253 + this.u8(); if (c === 254) return 506 + this.u8(); return c; }
  s255() {
    let c = this.u8();
    if (c === 253) return this.s16();
    let sign = 1;
    if (c === 250) { sign = -1; c = this.u8(); }
    const v = c === 255 ? 250 + this.u8() : c === 254 ? 500 + this.u8() : c;
    return sign * v;
  }
}
class Wr {
  constructor(n = 1024) { this.b = new Uint8Array(n); this.p = 0; }
  need(k) { if (this.p + k > this.b.length) { const o = new Uint8Array(Math.max(this.b.length * 2, this.p + k)); o.set(this.b); this.b = o; } }
  u8(v) { this.need(1); this.b[this.p++] = v & 255; }
  u16(v) { this.need(2); this.b[this.p++] = (v >> 8) & 255; this.b[this.p++] = v & 255; }
  u32(v) { this.u16((v >>> 16) & 0xffff); this.u16(v & 0xffff); }
  put(a) { this.need(a.length); this.b.set(a, this.p); this.p += a.length; }
  set16(at, v) { this.b[at] = (v >> 8) & 255; this.b[at + 1] = v & 255; }
  done() { return this.b.subarray(0, this.p); }
}

/** 좌표 세 쌍 부호화 표 (128 가지: 바이트 수, x 비트, y 비트, Δx, Δy, x 부호, y 부호) */
export const TRIPLETS = (() => {
  const t = [];
  const sg = (j) => [j & 1 ? 1 : -1, j & 2 ? 1 : -1];
  for (let i = 0; i < 10; i++) t.push([2, 0, 8, 0, (i >> 1) * 256, 0, i & 1 ? 1 : -1]);
  for (let i = 0; i < 10; i++) t.push([2, 8, 0, (i >> 1) * 256, 0, i & 1 ? 1 : -1, 0]);
  const d4 = [1, 17, 33, 49];
  for (let j = 0; j < 64; j++) { const [xs, ys] = sg(j); t.push([2, 4, 4, d4[j >> 4], d4[(j >> 2) & 3], xs, ys]); }
  const d8 = [1, 257, 513];
  for (let j = 0; j < 36; j++) { const [xs, ys] = sg(j); t.push([3, 8, 8, d8[Math.floor(j / 12)], d8[Math.floor(j / 4) % 3], xs, ys]); }
  for (let j = 0; j < 4; j++) { const [xs, ys] = sg(j); t.push([4, 12, 12, 0, 0, xs, ys]); }
  for (let j = 0; j < 4; j++) { const [xs, ys] = sg(j); t.push([5, 16, 16, 0, 0, xs, ys]); }
  return t;
})();

/** 밀어 넣기 명령 (PUSHB/PUSHW) 되살리기 — 홉 코드 0xFB · 0xFC */
function pushInstructions(r, out, count) {
  const vals = [];
  let left = count;
  while (left > 0) {
    const c = r.b[r.p];
    if (c === 0xfb || c === 0xfc) {
      r.p++;
      const a = vals[vals.length - 2];
      if (a === undefined) throw new Error('CTF: 잘못된 홉 코드');
      if (c === 0xfb) { vals.push(a, r.s255(), a); left -= 3; } else { vals.push(a, r.s255(), a, r.s255(), a); left -= 5; }
    } else { vals.push(r.s255()); left--; }
  }
  // 같은 종류(바이트/워드)끼리 묶어 쓰기, 255개마다 끊기
  let i = 0;
  while (i < vals.length) {
    const word = !(vals[i] >= 0 && vals[i] < 256);
    let j = i;
    while (j < vals.length && j - i < 255 && !(vals[j] >= 0 && vals[j] < 256) === word) j++;
    const n = j - i;
    if (n < 8) out.u8((word ? 0xb8 : 0xb0) | (n - 1)); else { out.u8(word ? 0x41 : 0x40); out.u8(n); }
    for (let k = i; k < j; k++) { if (word) out.u16(vals[k] & 0xffff); else out.u8(vals[k]); }
    i = j;
  }
}

function glyphOut(s0, s1, s2, out) {
  let nc = s0.s16();
  if (nc < 0) {
    // 복합 글리프: 대부분 그대로
    out.u16(0xffff);
    for (let k = 0; k < 4; k++) out.u16(s0.u16());
    let flags;
    do {
      flags = s0.u16();
      out.u16(flags);
      out.put(s0.bytes(2));
      out.put(s0.bytes(flags & 1 ? 4 : 2));
      out.put(s0.bytes(flags & 0x80 ? 8 : flags & 0x40 ? 4 : flags & 8 ? 2 : 0));
    } while (flags & 0x20);
    if (flags & 0x100) {
      const at = out.p;
      out.u16(0);
      const pc = s0.u255();
      pushInstructions(s1, out, pc);
      const cs = s0.u255();
      out.put(s2.bytes(cs));
      out.set16(at, out.p - at - 2);
    }
    return;
  }
  let bbox = null;
  if (nc === 0x7fff) { nc = s0.s16(); bbox = [s0.s16(), s0.s16(), s0.s16(), s0.s16()]; }
  if (nc === 0) return;
  out.u16(nc);
  const bboxAt = out.p;
  for (let k = 0; k < 4; k++) out.u16(bbox ? bbox[k] & 0xffff : 0);
  let total = 0;
  for (let c = 0; c < nc; c++) { total = c === 0 ? 1 + s0.u255() : total + s0.u255(); out.u16(total - 1); }
  const flags = s0.bytes(total);
  const xs = new Int32Array(total);
  const ys = new Int32Array(total);
  let x = 0;
  let y = 0;
  let minX = 32767; let minY = 32767; let maxX = -32768; let maxY = -32768;
  for (let i = 0; i < total; i++) {
    const [nb, xb, yb, dx, dy, sx, sy] = TRIPLETS[flags[i] & 0x7f];
    const raw = s0.bytes(nb - 1);
    let acc = 0;
    for (const v of raw) acc = acc * 256 + v;
    const totalBits = (nb - 1) * 8;
    const vx = Math.floor(acc / 2 ** (totalBits - xb)) % 2 ** xb;
    const vy = acc % 2 ** yb;
    xs[i] = sx * (vx + dx);
    ys[i] = sy * (vy + dy);
    x += xs[i];
    y += ys[i];
    minX = Math.min(minX, x); maxX = Math.max(maxX, x); minY = Math.min(minY, y); maxY = Math.max(maxY, y);
  }
  // 명령
  const lenAt = out.p;
  out.u16(0);
  const pc = s0.u255();
  pushInstructions(s1, out, pc);
  const cs = s0.u255();
  out.put(s2.bytes(cs));
  out.set16(lenAt, out.p - lenAt - 2);
  // 플래그 · 좌표 (TrueType 형식)
  for (let i = 0; i < total; i++) {
    let f = flags[i] & 0x80 ? 0 : 1;
    const ax = xs[i];
    const ay = ys[i];
    if (i > 0 && ax === 0) f |= 0x10; else if (ax > -256 && ax < 0) f |= 0x02; else if (ax >= 0 && ax < 256) f |= 0x12;
    if (i > 0 && ay === 0) f |= 0x20; else if (ay > -256 && ay < 0) f |= 0x04; else if (ay >= 0 && ay < 256) f |= 0x24;
    out.u8(f);
  }
  for (const arr of [xs, ys]) {
    for (let i = 0; i < total; i++) {
      const v = arr[i];
      if (i > 0 && v === 0) continue;
      if (v > -256 && v < 256) out.u8(Math.abs(v)); else out.u16(v & 0xffff);
    }
  }
  if (!bbox) { out.set16(bboxAt, minX & 0xffff); out.set16(bboxAt + 2, minY & 0xffff); out.set16(bboxAt + 4, maxX & 0xffff); out.set16(bboxAt + 6, maxY & 0xffff); }
}

/** CTF 세 흐름 → TrueType (sfnt) 바이트 */
export function ctfToSfnt([c0, c1, c2]) {
  const r = new Rd(c0);
  const scalar = r.u32();
  const num = r.u16();
  r.p += 6;
  const dir = [];
  for (let i = 0; i < num; i++) {
    const tag = String.fromCharCode(...r.bytes(4));
    r.p += 4; // 검사합
    const off = r.u32();
    const len = r.u32();
    if (tag === 'hdmx' || tag === 'VDMX') continue; // MTX 압축 형식 → 빼도 됨 (선택 표)
    dir.push({ tag, off, len });
  }
  const tables = new Map();
  for (const t of dir) {
    if (t.tag === 'glyf' || t.tag === 'loca') continue;
    if (t.tag === 'cvt ') {
      const cr = new Rd(c0, t.off);
      const n = cr.u16();
      const w = new Wr(n * 2);
      let last = 0;
      for (let i = 0; i < n; i++) {
        const code = cr.u8();
        let v;
        if (code >= 248) v = 238 * (code - 247) + cr.u8();
        else if (code >= 239) v = -(238 * (code - 239) + cr.u8());
        else if (code === 238) v = cr.s16();
        else v = code;
        last = ((last + v) << 16) >> 16;
        w.u16(last & 0xffff);
      }
      tables.set('cvt ', w.done());
      continue;
    }
    tables.set(t.tag, c0.slice(t.off, t.off + t.len));
  }
  const head = tables.get('head');
  const maxp = tables.get('maxp');
  if (!head || !maxp) throw new Error('CTF: head/maxp 표가 없습니다');
  head.fill(0, 8, 12);
  const longLoca = ((head[50] << 8) | head[51]) !== 0;
  const numGlyphs = (maxp[4] << 8) | maxp[5];
  const g = dir.find((t) => t.tag === 'glyf');
  if (g) {
    const s0 = new Rd(c0, g.off);
    const s1 = new Rd(c1);
    const s2 = new Rd(c2);
    const out = new Wr(Math.max(1024, g.len * 3));
    const loca = new Wr((numGlyphs + 1) * 4);
    if (longLoca) loca.u32(0); else loca.u16(0);
    for (let i = 0; i < numGlyphs; i++) {
      glyphOut(s0, s1, s2, out);
      if (out.p % 2) out.u8(0);
      if (longLoca) loca.u32(out.p); else loca.u16(out.p / 2);
    }
    tables.set('glyf', out.done());
    tables.set('loca', loca.done());
  }
  return buildSfnt(scalar, tables);
}

const sum32 = (b) => {
  let s = 0;
  for (let i = 0; i < b.length; i += 4) s = (s + (((b[i] << 24) | ((b[i + 1] ?? 0) << 16) | ((b[i + 2] ?? 0) << 8) | (b[i + 3] ?? 0)) >>> 0)) >>> 0;
  return s;
};
function buildSfnt(scalar, tables) {
  const tags = [...tables.keys()].sort();
  const n = tags.length;
  let es = 0;
  while (2 ** (es + 1) <= n) es++;
  const sr = 2 ** es * 16;
  let size = 12 + 16 * n;
  for (const t of tags) size += (tables.get(t).length + 3) & ~3;
  const out = new Uint8Array(size);
  const w = new Wr(0);
  w.b = out;
  w.u32(scalar);
  w.u16(n); w.u16(sr); w.u16(es); w.u16(n * 16 - sr);
  let off = 12 + 16 * n;
  let headOff = -1;
  for (const t of tags) {
    const d = tables.get(t);
    for (let k = 0; k < 4; k++) w.u8(t.charCodeAt(k));
    w.u32(sum32(d));
    w.u32(off);
    w.u32(d.length);
    out.set(d, off);
    if (t === 'head') headOff = off;
    off += (d.length + 3) & ~3;
  }
  if (headOff >= 0) {
    const adj = (0xb1b0afba - sum32(out)) >>> 0;
    out[headOff + 8] = adj >>> 24; out[headOff + 9] = (adj >> 16) & 255; out[headOff + 10] = (adj >> 8) & 255; out[headOff + 11] = adj & 255;
  }
  return out;
}

// ───────────── EOT ─────────────
const u32le = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;
const isSfnt = (b) => b.length > 12 && ((b[0] === 0 && b[1] === 1 && b[2] === 0 && b[3] === 0) || String.fromCharCode(b[0], b[1], b[2], b[3]) === 'OTTO' || String.fromCharCode(b[0], b[1], b[2], b[3]) === 'true' || String.fromCharCode(b[0], b[1], b[2], b[3]) === 'wOFF' || String.fromCharCode(b[0], b[1], b[2], b[3]) === 'wOF2');

/**
 * .fntdata (EOT) / .ttf / .otf → { data: sfnt 바이트, family, compressed } — 읽을 수 없으면 Error
 */
export function decodeEmbeddedFont(bytes) {
  if (isSfnt(bytes)) return { data: bytes, family: sfntFamily(bytes), compressed: false };
  if (bytes.length < 82 || ((bytes[34] | (bytes[35] << 8)) !== 0x504c)) throw new Error('글꼴 형식을 알 수 없습니다');
  const total = Math.min(u32le(bytes, 0), bytes.length);
  const dataSize = u32le(bytes, 4);
  const flags = u32le(bytes, 12);
  const start = total - dataSize;
  if (start < 0 || dataSize <= 0) throw new Error('글꼴 데이터가 손상되었습니다');
  let data = bytes.slice(start, total);
  if (flags & 0x10000000) for (let i = 0; i < data.length; i++) data[i] ^= 0x50;
  const compressed = !!(flags & 0x4);
  if (compressed) data = ctfToSfnt(unpackMtx(data));
  if (!isSfnt(data)) throw new Error('글꼴 데이터가 손상되었습니다');
  // EOT 머리글의 글꼴 이름 (UTF-16LE): 82 바이트 뒤 FamilyName
  let family = sfntFamily(data);
  try {
    const n = bytes[82] | (bytes[83] << 8);
    if (n > 0 && n < 512) { let s = ''; for (let i = 0; i < n; i += 2) s += String.fromCharCode(bytes[84 + i] | (bytes[85 + i] << 8)); family = s.replace(/\0+$/, '') || family; }
  } catch { /* 무시 */ }
  return { data, family, compressed };
}

/** sfnt 의 name 표에서 글꼴 집합 이름 (ID 1, 한국어 우선) */
export function sfntFamily(b) {
  try {
    const n = (b[4] << 8) | b[5];
    for (let i = 0; i < n; i++) {
      const o = 12 + i * 16;
      if (String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]) !== 'name') continue;
      const t = ((b[o + 8] << 24) | (b[o + 9] << 16) | (b[o + 10] << 8) | b[o + 11]) >>> 0;
      const count = (b[t + 2] << 8) | b[t + 3];
      const strOff = t + ((b[t + 4] << 8) | b[t + 5]);
      let best = null;
      for (let k = 0; k < count; k++) {
        const r = t + 6 + k * 12;
        const pid = (b[r] << 8) | b[r + 1];
        const lang = (b[r + 4] << 8) | b[r + 5];
        const nid = (b[r + 6] << 8) | b[r + 7];
        const len = (b[r + 8] << 8) | b[r + 9];
        const off = (b[r + 10] << 8) | b[r + 11];
        if (nid !== 1) continue;
        let s = '';
        if (pid === 3 || pid === 0) for (let j = 0; j < len; j += 2) s += String.fromCharCode((b[strOff + off + j] << 8) | b[strOff + off + j + 1]);
        else for (let j = 0; j < len; j++) s += String.fromCharCode(b[strOff + off + j]);
        const score = (pid === 3 && lang === 0x412 ? 3 : pid === 3 ? 2 : 1);
        if (!best || score > best.score) best = { s, score };
      }
      return best?.s ?? '';
    }
  } catch { /* 무시 */ }
  return '';
}
