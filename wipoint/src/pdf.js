// PDF 열기 (DOM 없음): 쪽마다 슬라이드 하나, 글은 편집할 수 있는 텍스트 상자, 도형 · 선은 도형/자유형, 그림은 그림 개체
// 직접 구현: 객체 찾기(xref 없이 훑기 + 개체 스트림), Flate/ASCIIHex/ASCII85/LZW/RunLength + PNG 예측자, 내용 스트림 해석,
// 글꼴(ToUnicode CMap · WinAnsi/Differences · Type0/CID 너비), 그림(JPEG 그대로, 나머지는 RGBA → encodeImage)
import { inflate } from './zip.js';
import { newPresentation, uid } from './model.js';

const latin1 = new TextDecoder('latin1');
const PT2PX = 96 / 72;

// ───────────── 기본 값 ─────────────
class Ref { constructor(num, gen) { this.num = num; this.gen = gen; } }
class Name { constructor(n) { this.n = n; } }
class Cmd { constructor(c) { this.c = c; } }
class Stream { constructor(dict, start, end, bytes) { this.dict = dict; this.start = start; this.end = end; this.bytes = bytes; } }
const isName = (v, n) => v instanceof Name && (n == null || v.n === n);

const WS = new Uint8Array(256);
for (const c of [0, 9, 10, 12, 13, 32]) WS[c] = 1;
const DELIM = new Uint8Array(256);
for (const c of '()<>[]{}/%') DELIM[c.charCodeAt(0)] = 1;

// ───────────── 낱말 읽기 ─────────────
class Lexer {
  constructor(b, pos = 0, end = b.length) { this.b = b; this.pos = pos; this.end = end; }
  skipWs() {
    const b = this.b;
    while (this.pos < this.end) {
      const c = b[this.pos];
      if (WS[c]) this.pos++;
      else if (c === 37) { while (this.pos < this.end && b[this.pos] !== 10 && b[this.pos] !== 13) this.pos++; }
      else break;
    }
  }
  /** 다음 값 (배열 · 사전 포함). 연산자는 Cmd, 끝이면 undefined */
  next() {
    this.skipWs();
    const b = this.b;
    if (this.pos >= this.end) return undefined;
    const c = b[this.pos];
    if (c === 40) return this.litString();
    if (c === 60) {
      if (b[this.pos + 1] === 60) { this.pos += 2; return this.dict(); }
      return this.hexString();
    }
    if (c === 62 && b[this.pos + 1] === 62) { this.pos += 2; return new Cmd('>>'); }
    if (c === 91) { this.pos++; return this.array(); }
    if (c === 93) { this.pos++; return new Cmd(']'); }
    if (c === 47) return this.name();
    if (c === 123 || c === 125) { this.pos++; return new Cmd(String.fromCharCode(c)); }
    // 숫자 · 낱말
    let s = this.pos;
    while (this.pos < this.end && !WS[b[this.pos]] && !DELIM[b[this.pos]]) this.pos++;
    if (this.pos === s) { this.pos++; return new Cmd(String.fromCharCode(c)); }
    const t = latin1.decode(b.subarray(s, this.pos));
    if (/^[+-]?(\d+\.?\d*|\.\d+)$/.test(t)) return parseFloat(t);
    if (/^[+-]?\d*\.?\d*[+-]\d*$/.test(t)) return parseFloat(t) || 0; // 잘못된 숫자 (1.5-3 등)
    if (t === 'true') return true;
    if (t === 'false') return false;
    if (t === 'null') return null;
    return new Cmd(t);
  }
  name() {
    const b = this.b;
    this.pos++;
    const s = this.pos;
    while (this.pos < this.end && !WS[b[this.pos]] && !DELIM[b[this.pos]]) this.pos++;
    let raw = latin1.decode(b.subarray(s, this.pos));
    if (raw.includes('#')) raw = raw.replace(/#([0-9a-fA-F]{2})/g, (m, h) => String.fromCharCode(parseInt(h, 16)));
    return new Name(raw);
  }
  litString() {
    const b = this.b;
    this.pos++;
    const out = [];
    let depth = 1;
    while (this.pos < this.end) {
      let c = b[this.pos++];
      if (c === 92) {
        c = b[this.pos++];
        const map = { 110: 10, 114: 13, 116: 9, 98: 8, 102: 12 };
        if (map[c] != null) out.push(map[c]);
        else if (c >= 48 && c <= 55) {
          let v = c - 48;
          for (let k = 0; k < 2 && b[this.pos] >= 48 && b[this.pos] <= 55; k++) v = v * 8 + b[this.pos++] - 48;
          out.push(v & 255);
        } else if (c === 13) { if (b[this.pos] === 10) this.pos++; } else if (c === 10) { /* 줄 이음 */ } else out.push(c);
        continue;
      }
      if (c === 40) depth++;
      else if (c === 41 && --depth === 0) break;
      out.push(c);
    }
    return Uint8Array.from(out);
  }
  hexString() {
    const b = this.b;
    this.pos++;
    const out = [];
    let hi = -1;
    while (this.pos < this.end) {
      const c = b[this.pos++];
      if (c === 62) break;
      const v = c >= 48 && c <= 57 ? c - 48 : c >= 65 && c <= 70 ? c - 55 : c >= 97 && c <= 102 ? c - 87 : -1;
      if (v < 0) continue;
      if (hi < 0) hi = v; else { out.push(hi * 16 + v); hi = -1; }
    }
    if (hi >= 0) out.push(hi * 16);
    return Uint8Array.from(out);
  }
  array() {
    const out = [];
    for (;;) {
      const v = this.next();
      if (v === undefined || (v instanceof Cmd && v.c === ']')) break;
      if (v instanceof Cmd && v.c === 'R' && out.length >= 2) { const g = out.pop(); const n = out.pop(); out.push(new Ref(n, g)); continue; }
      out.push(v);
    }
    return out;
  }
  dict() {
    const d = {};
    for (;;) {
      const k = this.next();
      if (k === undefined || (k instanceof Cmd && k.c === '>>')) break;
      if (!(k instanceof Name)) continue;
      let v = this.next();
      if (typeof v === 'number') {
        // n g R ?
        const save = this.pos;
        const g = this.next();
        if (typeof g === 'number') {
          const r = this.next();
          if (r instanceof Cmd && r.c === 'R') v = new Ref(v, g);
          else this.pos = save;
        } else this.pos = save;
      }
      if (v instanceof Cmd && v.c === '>>') { d[k.n] = null; break; }
      d[k.n] = v;
    }
    return d;
  }
}

// ───────────── 문서 ─────────────
class PdfDoc {
  constructor(bytes) {
    this.b = bytes;
    this.offsets = new Map(); // num → byte offset (마지막 정의)
    this.inStm = new Map();   // num → [stmNum, index]
    this.cache = new Map();
    this.text = latin1.decode(bytes);
    this.scan();
  }
  scan() {
    const re = /(\d+)\s+(\d+)\s+obj\b/g;
    let m;
    const stms = [];
    while ((m = re.exec(this.text))) {
      this.offsets.set(Number(m[1]), m.index + m[0].length);
    }
    // 개체 스트림 (PDF 1.5+): /Type /ObjStm
    for (const [num, off] of this.offsets) {
      const head = this.text.slice(off, off + 400);
      if (/\/Type\s*\/ObjStm/.test(head)) stms.push(num);
    }
    for (const sn of stms) {
      const st = this.get(new Ref(sn, 0));
      if (!(st instanceof Stream)) continue;
      let data;
      try { data = this.decode(st); } catch { continue; }
      const n = this.val(st.dict.N) ?? 0;
      const lx = new Lexer(data);
      for (let i = 0; i < n; i++) {
        const num = lx.next();
        lx.next();
        if (typeof num !== 'number') break;
        if (!this.offsets.has(num) || (this.inStm.get(num)?.[0] ?? -1) < sn) this.inStm.set(num, [sn, i]);
      }
    }
    // 트레일러
    this.trailer = {};
    const tre = /trailer\s*<</g;
    let tm;
    while ((tm = tre.exec(this.text))) { const lx = new Lexer(this.b, tm.index + tm[0].length - 2); const d = lx.next(); if (d && typeof d === 'object' && !(d instanceof Uint8Array)) Object.assign(this.trailer, d); }
    if (!this.trailer.Root) {
      for (const [num, off] of this.offsets) {
        const head = this.text.slice(off, off + 600);
        if (/\/Type\s*\/XRef/.test(head)) { const o = this.get(new Ref(num, 0)); if (o instanceof Stream) { if (o.dict.Root) this.trailer.Root = o.dict.Root; if (o.dict.Encrypt) this.trailer.Encrypt = o.dict.Encrypt; } }
      }
    }
    if (!this.trailer.Root) {
      for (const num of [...this.offsets.keys(), ...this.inStm.keys()]) {
        const o = this.get(new Ref(num, 0));
        if (o && typeof o === 'object' && isName(o.Type, 'Catalog')) { this.trailer.Root = new Ref(num, 0); break; }
      }
    }
  }
  /** 참조 풀기 */
  get(ref) {
    if (!(ref instanceof Ref)) return ref;
    if (this.cache.has(ref.num)) return this.cache.get(ref.num);
    this.cache.set(ref.num, null); // 순환 방지
    let v = null;
    const off = this.offsets.get(ref.num);
    const inStm = this.inStm.get(ref.num);
    if (inStm && (off == null || inStm)) v = this.fromObjStm(inStm[0], inStm[1]) ?? (off != null ? this.parseAt(off) : null);
    else if (off != null) v = this.parseAt(off);
    this.cache.set(ref.num, v);
    return v;
  }
  val(v) { return v instanceof Ref ? this.get(v) : v; }
  parseAt(off) {
    const lx = new Lexer(this.b, off);
    const v = lx.next();
    if (v && typeof v === 'object' && !(v instanceof Uint8Array) && !Array.isArray(v) && !(v instanceof Name) && !(v instanceof Cmd)) {
      const save = lx.pos;
      lx.skipWs();
      if (this.text.startsWith('stream', lx.pos)) {
        let s = lx.pos + 6;
        if (this.b[s] === 13) s++;
        if (this.b[s] === 10) s++;
        let len = v.Length;
        if (len instanceof Ref) {
          // 길이가 다른 개체에 있으면 (순환 위험 없이) 직접 읽기
          const lo = this.offsets.get(len.num);
          len = lo != null ? new Lexer(this.b, lo).next() : null;
        }
        let e = typeof len === 'number' ? s + len : -1;
        if (e < 0 || e > this.b.length || !/^\s*endstream/.test(this.text.slice(e, e + 20))) {
          const k = this.text.indexOf('endstream', s);
          e = k < 0 ? this.b.length : k;
          while (e > s && (this.b[e - 1] === 10 || this.b[e - 1] === 13)) e--;
        }
        return new Stream(v, s, e, this.b);
      }
      lx.pos = save;
    }
    return v;
  }
  fromObjStm(sn, idx) {
    const key = `stm${sn}`;
    let parsed = this.cache.get(key);
    if (!parsed) {
      const st = this.get(new Ref(sn, 0));
      if (!(st instanceof Stream)) return null;
      let data;
      try { data = this.decode(st); } catch { return null; }
      const n = this.val(st.dict.N) ?? 0;
      const first = this.val(st.dict.First) ?? 0;
      const lx = new Lexer(data);
      const offs = [];
      for (let i = 0; i < n; i++) { const num = lx.next(); const o = lx.next(); offs.push([num, o]); }
      parsed = { data, first, offs };
      this.cache.set(key, parsed);
    }
    const o = parsed.offs[idx];
    if (!o) return null;
    return new Lexer(parsed.data, parsed.first + o[1]).next();
  }
  /** 스트림 풀기 (그림 압축 DCT/JPX 등은 그 앞 단계까지) */
  decode(st, { stopAtImage = false } = {}) {
    let data = st.bytes.subarray(st.start, st.end);
    let filters = this.val(st.dict.Filter ?? st.dict.F);
    let parms = this.val(st.dict.DecodeParms ?? st.dict.DP);
    if (!filters) return data;
    if (!Array.isArray(filters)) { filters = [filters]; parms = [parms]; }
    for (let i = 0; i < filters.length; i++) {
      const f = this.val(filters[i])?.n;
      const p = this.val(Array.isArray(parms) ? parms[i] : parms) ?? {};
      if (stopAtImage && (f === 'DCTDecode' || f === 'DCT' || f === 'JPXDecode' || f === 'CCITTFaxDecode' || f === 'JBIG2Decode')) { data.imageFilter = f; return Object.assign(data, { imageFilter: f }); }
      if (f === 'FlateDecode' || f === 'Fl') data = predict(flate(data), p, this);
      else if (f === 'LZWDecode' || f === 'LZW') data = predict(lzw(data, this.val(p.EarlyChange) ?? 1), p, this);
      else if (f === 'ASCIIHexDecode' || f === 'AHx') data = new Lexer(Uint8Array.from([60, ...data, 62])).hexString();
      else if (f === 'ASCII85Decode' || f === 'A85') data = ascii85(data);
      else if (f === 'RunLengthDecode' || f === 'RL') data = runLength(data);
      else if (f === 'Crypt') { /* 항등 */ } else throw new Error(`지원하지 않는 압축: ${f}`);
    }
    return data;
  }
}

function flate(data) {
  let s = 0;
  if (data.length > 2 && (data[0] & 0x0f) === 8 && ((data[0] << 8) | data[1]) % 31 === 0) s = 2;
  try { return inflate(data.subarray(s), data.length * 4); } catch (e) {
    // 끝이 잘린 스트림: 읽힌 데까지 (inflate 가 던지면 빈 것)
    return new Uint8Array(0);
  }
}
function predict(data, p, doc) {
  const pred = doc.val(p.Predictor) ?? 1;
  if (pred < 2) return data;
  const colors = doc.val(p.Colors) ?? 1;
  const bpc = doc.val(p.BitsPerComponent) ?? 8;
  const cols = doc.val(p.Columns) ?? 1;
  const bpp = Math.max(1, (colors * bpc) >> 3);
  const row = (cols * colors * bpc + 7) >> 3;
  if (pred === 2) {
    const out = data.slice();
    if (bpc === 8) for (let r = 0; r + row <= out.length; r += row) for (let i = bpp; i < row; i++) out[r + i] = (out[r + i] + out[r + i - bpp]) & 255;
    return out;
  }
  const rows = Math.floor(data.length / (row + 1));
  const out = new Uint8Array(rows * row);
  let prev = new Uint8Array(row);
  for (let r = 0; r < rows; r++) {
    const t = data[r * (row + 1)];
    const src = data.subarray(r * (row + 1) + 1, (r + 1) * (row + 1));
    const cur = out.subarray(r * row, (r + 1) * row);
    for (let i = 0; i < row; i++) {
      const a = i >= bpp ? cur[i - bpp] : 0;
      const b = prev[i];
      const c = i >= bpp ? prev[i - bpp] : 0;
      let v = src[i];
      if (t === 1) v += a;
      else if (t === 2) v += b;
      else if (t === 3) v += (a + b) >> 1;
      else if (t === 4) { const pp = a + b - c; const pa = Math.abs(pp - a); const pb = Math.abs(pp - b); const pc = Math.abs(pp - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      cur[i] = v & 255;
    }
    prev = cur;
  }
  return out;
}
function lzw(data, early) {
  const out = [];
  let dict = [];
  const reset = () => { dict = []; for (let i = 0; i < 256; i++) dict.push([i]); dict.push(null, null); };
  reset();
  let bits = 9;
  let buf = 0;
  let cnt = 0;
  let prev = null;
  for (let i = 0; i < data.length; i++) {
    buf = (buf << 8) | data[i];
    cnt += 8;
    while (cnt >= bits) {
      const code = (buf >> (cnt - bits)) & ((1 << bits) - 1);
      cnt -= bits;
      if (code === 256) { reset(); bits = 9; prev = null; continue; }
      if (code === 257) return Uint8Array.from(out);
      let entry = dict[code];
      if (!entry && prev) entry = [...prev, prev[0]];
      if (!entry) return Uint8Array.from(out);
      for (const v of entry) out.push(v);
      if (prev) dict.push([...prev, entry[0]]);
      prev = entry;
      const n = dict.length + early;
      bits = n >= 2048 ? 12 : n >= 1024 ? 11 : n >= 512 ? 10 : 9;
    }
  }
  return Uint8Array.from(out);
}
function ascii85(data) {
  const out = [];
  let tuple = [];
  for (let i = 0; i < data.length; i++) {
    const c = data[i];
    if (c === 126) break;
    if (WS[c]) continue;
    if (c === 122 && !tuple.length) { out.push(0, 0, 0, 0); continue; }
    tuple.push(c - 33);
    if (tuple.length === 5) {
      let v = 0;
      for (const t of tuple) v = v * 85 + t;
      out.push((v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255);
      tuple = [];
    }
  }
  if (tuple.length) {
    const n = tuple.length;
    while (tuple.length < 5) tuple.push(84);
    let v = 0;
    for (const t of tuple) v = v * 85 + t;
    const bytes = [(v >>> 24) & 255, (v >>> 16) & 255, (v >>> 8) & 255, v & 255];
    out.push(...bytes.slice(0, n - 1));
  }
  return Uint8Array.from(out);
}
function runLength(data) {
  const out = [];
  for (let i = 0; i < data.length;) {
    const n = data[i++];
    if (n === 128) break;
    if (n < 128) { for (let k = 0; k <= n; k++) out.push(data[i++]); } else { const v = data[i++]; for (let k = 0; k < 257 - n; k++) out.push(v); }
  }
  return Uint8Array.from(out);
}

// ───────────── 글꼴 ─────────────
const AGL = {
  space: ' ', exclam: '!', quotedbl: '"', numbersign: '#', dollar: '$', percent: '%', ampersand: '&', quotesingle: "'", quoteright: '’', quoteleft: '‘', parenleft: '(', parenright: ')', asterisk: '*', plus: '+', comma: ',', hyphen: '-', minus: '−', period: '.', slash: '/',
  zero: '0', one: '1', two: '2', three: '3', four: '4', five: '5', six: '6', seven: '7', eight: '8', nine: '9', colon: ':', semicolon: ';', less: '<', equal: '=', greater: '>', question: '?', at: '@',
  bracketleft: '[', backslash: '\\', bracketright: ']', asciicircum: '^', underscore: '_', grave: '`', braceleft: '{', bar: '|', braceright: '}', asciitilde: '~', bullet: '•', endash: '–', emdash: '—', quotedblleft: '“', quotedblright: '”', ellipsis: '…', copyright: '©', registered: '®', trademark: '™', degree: '°', periodcentered: '·', middot: '·', multiply: '×', divide: '÷', Euro: '€', fi: 'fi', fl: 'fl', nbspace: ' ', uni00A0: ' ',
};
const SYMBOL_BULLET = { 0xa7: '▪', 0xb7: '•', 0xfc: '✓', 0xd8: '➢', 0x76: '❖', 0x6c: '●', 0x6e: '■', 0x71: '❑', 0x75: '◆' };
function glyphToUni(g) {
  if (AGL[g]) return AGL[g];
  if (g.length === 1) return g;
  let m = /^uni([0-9A-Fa-f]{4,})$/.exec(g);
  if (m) { let s = ''; for (let i = 0; i + 4 <= m[1].length; i += 4) s += String.fromCharCode(parseInt(m[1].slice(i, i + 4), 16)); return s; }
  m = /^u([0-9A-Fa-f]{4,6})$/.exec(g);
  if (m) return String.fromCodePoint(parseInt(m[1], 16));
  m = /^([a-zA-Z]+)\.(sc|alt|\w+)$/.exec(g);
  if (m) return glyphToUni(m[1]);
  return null;
}
const winAnsi = new TextDecoder('windows-1252');
let macRoman = null;
try { macRoman = new TextDecoder('macintosh'); } catch { /* 없음 */ }

function utf16be(bytes) {
  let s = '';
  for (let i = 0; i + 1 < bytes.length; i += 2) s += String.fromCharCode((bytes[i] << 8) | bytes[i + 1]);
  return s;
}
const bytesNum = (b) => { let v = 0; for (const x of b) v = v * 256 + x; return v; };

/** CMap (ToUnicode / 인코딩) 읽기 → { map: Map(code→string|cid), ranges: [{lo, hi, n}] } */
function parseCMap(data, toUnicode) {
  const lx = new Lexer(data);
  const map = new Map();
  const ranges = [];
  const stack = [];
  for (;;) {
    const v = lx.next();
    if (v === undefined) break;
    if (!(v instanceof Cmd)) { stack.push(v); continue; }
    const op = v.c;
    if (op === 'endcodespacerange') {
      for (let i = 0; i + 1 < stack.length; i += 2) if (stack[i] instanceof Uint8Array) ranges.push({ lo: bytesNum(stack[i]), hi: bytesNum(stack[i + 1]), n: stack[i].length });
    } else if (op === 'endbfchar' || op === 'endcidchar') {
      for (let i = 0; i + 1 < stack.length; i += 2) {
        const src = stack[i];
        const dst = stack[i + 1];
        if (!(src instanceof Uint8Array)) continue;
        map.set(bytesNum(src), toUnicode ? (dst instanceof Uint8Array ? utf16be(dst) : dst instanceof Name ? glyphToUni(dst.n) ?? '' : String.fromCharCode(dst)) : dst);
      }
    } else if (op === 'endbfrange' || op === 'endcidrange') {
      for (let i = 0; i + 2 < stack.length; i += 3) {
        const lo = bytesNum(stack[i]);
        const hi = bytesNum(stack[i + 1]);
        const dst = stack[i + 2];
        if (hi - lo > 65535) continue;
        if (Array.isArray(dst)) { for (let c = lo; c <= hi; c++) { const d = dst[c - lo]; if (d instanceof Uint8Array) map.set(c, utf16be(d)); } } else if (dst instanceof Uint8Array && toUnicode) {
          const base = utf16be(dst);
          const last = base.charCodeAt(base.length - 1);
          for (let c = lo; c <= hi; c++) map.set(c, base.slice(0, -1) + String.fromCharCode(last + (c - lo)));
        } else if (typeof dst === 'number') { for (let c = lo; c <= hi; c++) map.set(c, dst + (c - lo)); }
      }
    }
    if (op.startsWith('end') || op.startsWith('begin') || op === 'def' || op === 'usecmap') stack.length = 0;
  }
  return { map, ranges };
}

const FONT_NAMES = [
  [/malgun|맑은/i, '맑은 고딕'], [/nanumsquare/i, '나눔스퀘어'], [/nanumbarungothic/i, '나눔바른고딕'], [/nanumgothic|나눔고딕/i, '나눔고딕'], [/nanummyeongjo/i, '나눔명조'], [/pretendard/i, 'Pretendard'],
  [/notosans(kr|cjk)|noto sans (kr|cjk)|sourcehansans|본고딕/i, 'Noto Sans KR'], [/notoserif(kr|cjk)|sourcehanserif/i, 'Noto Serif KR'], [/gulim|굴림/i, '굴림'], [/dotum|돋움/i, '돋움'], [/batang|바탕/i, '바탕'], [/gungsuh|궁서/i, '궁서'],
  [/wenquanyi|droidsansfallback|unbatang|undotum|baekmuk/i, '맑은 고딕'], [/arialblack/i, 'Arial Black'], [/arial|helvetica|liberationsans/i, 'Arial'], [/calibri|carlito/i, 'Calibri'], [/cambria|caladea/i, 'Cambria'], [/times|liberationserif/i, 'Times New Roman'], [/georgia/i, 'Georgia'], [/segoe/i, 'Segoe UI'], [/tahoma/i, 'Tahoma'], [/verdana/i, 'Verdana'], [/consolas/i, 'Consolas'], [/courier|liberationmono/i, 'Courier New'], [/trebuchet/i, 'Trebuchet MS'], [/impact/i, 'Impact'], [/dejavusans/i, 'Arial'],
];
function fontFamily(base) {
  const n = base.replace(/^[A-Z]{6}\+/, '');
  for (const [re, f] of FONT_NAMES) if (re.test(n.replace(/[\s_-]/g, ''))) return f;
  return n.replace(/[,-](Bold|Italic|Oblique|Regular|Medium|Light|Black|Heavy|SemiBold|ExtraBold|Thin|BoldItalic|BoldOblique|Roman|Book|MT|PS|PSMT)+$/i, '').replace(/(PSMT|MT|PS)$/, '') || n;
}

class Font {
  constructor(doc, dict) {
    this.doc = doc;
    const d = dict ?? {};
    const v = (x) => doc.val(x);
    this.subtype = v(d.Subtype)?.n ?? 'Type1';
    const base = v(d.BaseFont)?.n ?? '';
    this.base = base;
    let desc = v(d.FontDescriptor);
    this.type0 = this.subtype === 'Type0';
    this.widths = new Map();
    this.dw = 1000;
    this.fm = 0.001;
    if (this.type0) {
      const df = v(v(d.DescendantFonts)?.[0]) ?? {};
      desc = v(df.FontDescriptor) ?? desc;
      this.dw = v(df.DW) ?? 1000;
      const w = v(df.W) ?? [];
      for (let i = 0; i < w.length;) {
        const a = v(w[i]);
        const b = v(w[i + 1]);
        if (Array.isArray(b)) { b.forEach((x, k) => this.widths.set(a + k, v(x))); i += 2; } else { const ww = v(w[i + 2]); for (let c = a; c <= b && c - a < 65536; c++) this.widths.set(c, ww); i += 3; }
      }
      const enc = v(d.Encoding);
      this.codeLen = null;
      this.cidMap = null;
      this.encName = enc instanceof Name ? enc.n : '';
      if (enc instanceof Stream) { try { const cm = parseCMap(doc.decode(enc), false); this.ranges = cm.ranges; this.cidMap = cm.map; } catch { /* 무시 */ } }
      this.ucs2 = /UCS2|UTF16/.test(this.encName);
    } else {
      const first = v(d.FirstChar) ?? 0;
      (v(d.Widths) ?? []).forEach((x, i) => this.widths.set(first + i, v(x)));
      if (this.subtype === 'Type3') { const fm = v(d.FontMatrix); if (fm) this.fm = fm[0]; this.dw = 0; } else this.dw = v(desc?.MissingWidth) ?? 500;
      // 인코딩
      this.enc = new Array(256).fill(null);
      const enc = v(d.Encoding);
      const baseEnc = enc instanceof Name ? enc.n : v(enc?.BaseEncoding)?.n;
      this.symbolic = /Symbol|Wingdings|Webdings|ZapfDingbats/i.test(base) || ((v(desc?.Flags) ?? 0) & 4 && !baseEnc && !(enc && enc.Differences));
      this.macRoman = baseEnc === 'MacRomanEncoding';
      const diff = enc && !(enc instanceof Name) ? v(enc.Differences) : null;
      if (diff) { let c = 0; for (const x of diff) { const y = v(x); if (typeof y === 'number') c = y; else if (y instanceof Name) this.enc[c++] = glyphToUni(y.n); } }
    }
    this.bold = /bold|black|heavy|semibold|extrabold|demi|-bd\b|,bd/i.test(base) || ((v(desc?.FontWeight) ?? 400) >= 600) || ((v(desc?.Flags) ?? 0) & 0x40000) !== 0;
    this.italic = /italic|oblique/i.test(base) || ((v(desc?.ItalicAngle) ?? 0) !== 0);
    this.family = fontFamily(base);
    const tu = v(d.ToUnicode);
    this.toUni = null;
    if (tu instanceof Stream) { try { const cm = parseCMap(doc.decode(tu), true); this.toUni = cm.map; if (this.type0 && !this.ranges?.length && cm.ranges.length) this.ranges = cm.ranges; } catch { /* 무시 */ } }
    this.vertical = this.type0 && /-V$/.test(this.encName ?? '');
  }
  /** 글자 바이트 → [{code, uni, w(글자 공간 1/1000), space}] */
  decode(bytes) {
    const out = [];
    if (!this.type0) {
      for (const c of bytes) out.push({ code: c, uni: this.uniOf(c), w: this.widths.get(c) ?? this.dw, space: c === 32 });
      return out;
    }
    for (let i = 0; i < bytes.length;) {
      let n = 2;
      if (this.ranges?.length) {
        n = 0;
        for (const len of [1, 2, 3, 4]) {
          if (i + len > bytes.length) break;
          const code = bytesNum(bytes.subarray(i, i + len));
          if (this.ranges.some((r) => r.n === len && code >= r.lo && code <= r.hi)) { n = len; break; }
        }
        if (!n) n = Math.min(2, bytes.length - i) || 1;
      }
      const code = bytesNum(bytes.subarray(i, i + n));
      i += n;
      const cid = this.cidMap?.get(code) ?? code;
      let uni = this.toUni?.get(code);
      if (uni == null && this.ucs2) uni = String.fromCharCode(code);
      out.push({ code, uni: uni ?? '', w: this.widths.get(cid) ?? this.dw, space: n === 1 && code === 32 });
    }
    return out;
  }
  uniOf(c) {
    const t = this.toUni?.get(c);
    if (t != null) return t;
    if (this.enc[c] != null) return this.enc[c];
    if (this.symbolic) return SYMBOL_BULLET[c] ?? (c >= 32 && c < 127 ? String.fromCharCode(c) : '•');
    if (this.macRoman && macRoman) return macRoman.decode(Uint8Array.of(c));
    return c < 32 ? '' : winAnsi.decode(Uint8Array.of(c));
  }
}

// ───────────── 색 ─────────────
const clamp = (v) => Math.max(0, Math.min(255, Math.round(v)));
const hex = (r, g, b) => `#${[r, g, b].map((x) => clamp(x).toString(16).padStart(2, '0')).join('')}`.toUpperCase();
function cmyk(c, m, y, k) { return [255 * (1 - c) * (1 - k), 255 * (1 - m) * (1 - k), 255 * (1 - y) * (1 - k)]; }

/** 색 공간 → { n, toRgb(comps[0..1]) } */
function colorSpace(doc, cs, res) {
  cs = doc.val(cs);
  if (cs instanceof Name) {
    const nm = cs.n;
    if (nm === 'DeviceGray' || nm === 'G' || nm === 'CalGray') return { n: 1, rgb: (v) => [v[0] * 255, v[0] * 255, v[0] * 255] };
    if (nm === 'DeviceRGB' || nm === 'RGB' || nm === 'CalRGB') return { n: 3, rgb: (v) => [v[0] * 255, v[1] * 255, v[2] * 255] };
    if (nm === 'DeviceCMYK' || nm === 'CMYK') return { n: 4, rgb: (v) => cmyk(v[0], v[1], v[2], v[3]) };
    if (nm === 'Pattern') return { n: 0, pattern: true, rgb: () => [128, 128, 128] };
    const named = doc.val(doc.val(res?.ColorSpace)?.[nm]);
    if (named) return colorSpace(doc, named, res);
    return { n: 3, rgb: (v) => [v[0] * 255, (v[1] ?? v[0]) * 255, (v[2] ?? v[0]) * 255] };
  }
  if (Array.isArray(cs)) {
    const kind = doc.val(cs[0])?.n;
    if (kind === 'ICCBased') { const st = doc.val(cs[1]); const n = doc.val(st?.dict?.N) ?? 3; return colorSpace(doc, new Name(n === 1 ? 'DeviceGray' : n === 4 ? 'DeviceCMYK' : 'DeviceRGB'), res); }
    if (kind === 'CalRGB' || kind === 'Lab') return colorSpace(doc, new Name('DeviceRGB'), res);
    if (kind === 'CalGray') return colorSpace(doc, new Name('DeviceGray'), res);
    if (kind === 'Indexed' || kind === 'I') {
      const base = colorSpace(doc, cs[1], res);
      const hival = doc.val(cs[2]);
      let lut = doc.val(cs[3]);
      if (lut instanceof Stream) lut = doc.decode(lut);
      return { n: 1, indexed: true, hival, base, lut, rgb: (v) => { const i = Math.round(v[0]); const o = i * base.n; return base.rgb(Array.from({ length: base.n }, (_, k) => (lut[o + k] ?? 0) / 255)); } };
    }
    if (kind === 'Separation' || kind === 'DeviceN') {
      const alt = colorSpace(doc, cs[kind === 'Separation' ? 2 : 2], res);
      const fn = doc.val(cs[3]);
      const nIn = kind === 'Separation' ? 1 : (doc.val(cs[1])?.length ?? 1);
      return { n: nIn, rgb: (v) => { const out = evalFunction(doc, fn, v); return alt.rgb(out.length ? out : [1 - v[0], 1 - v[0], 1 - v[0]]); } };
    }
    if (kind === 'Pattern') return { n: 0, pattern: true, rgb: () => [128, 128, 128] };
  }
  return colorSpace(doc, new Name('DeviceRGB'), res);
}

/** PDF 함수 (형식 2 지수 보간 · 3 이어 붙이기 · 0 표본 근사) */
function evalFunction(doc, fn, input) {
  fn = doc.val(fn);
  if (Array.isArray(fn)) return fn.flatMap((f) => evalFunction(doc, f, input));
  const d = fn instanceof Stream ? fn.dict : fn;
  if (!d) return [];
  const t = doc.val(d.FunctionType);
  const x = input[0] ?? 0;
  if (t === 2) {
    const c0 = doc.val(d.C0) ?? [0];
    const c1 = doc.val(d.C1) ?? [1];
    const n = doc.val(d.N) ?? 1;
    const xn = Math.pow(Math.max(0, x), n);
    return c0.map((a, i) => a + xn * ((c1[i] ?? a) - a));
  }
  if (t === 3) {
    const fns = doc.val(d.Functions) ?? [];
    const bounds = doc.val(d.Bounds) ?? [];
    const domain = doc.val(d.Domain) ?? [0, 1];
    const enc = doc.val(d.Encode) ?? [];
    let k = 0;
    while (k < bounds.length && x >= bounds[k]) k++;
    const lo = k === 0 ? domain[0] : bounds[k - 1];
    const hi = k === bounds.length ? domain[1] : bounds[k];
    const e0 = enc[2 * k] ?? 0;
    const e1 = enc[2 * k + 1] ?? 1;
    const xx = hi > lo ? e0 + (x - lo) / (hi - lo) * (e1 - e0) : e0;
    return evalFunction(doc, fns[k], [xx]);
  }
  if (t === 0 && fn instanceof Stream) {
    try {
      const data = doc.decode(fn);
      const range = doc.val(d.Range) ?? [0, 1, 0, 1, 0, 1];
      const nOut = range.length / 2;
      const size = doc.val(d.Size)?.[0] ?? 2;
      const bps = doc.val(d.BitsPerSample) ?? 8;
      if (bps !== 8) return range.filter((_, i) => i % 2 === 0);
      const idx = Math.min(size - 1, Math.max(0, Math.round(x * (size - 1))));
      return Array.from({ length: nOut }, (_, k) => range[2 * k] + (data[idx * nOut + k] ?? 0) / 255 * (range[2 * k + 1] - range[2 * k]));
    } catch { return []; }
  }
  return [];
}

// ───────────── 행렬 ─────────────
const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const apply = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

// ───────────── 그림 ─────────────
/** PNG 인코더 (무압축 deflate, DOM 없음 — 브라우저에서는 encodeImage 로 canvas 사용) */
export function pngDataUrl(w, h, rgba) {
  const raw = new Uint8Array((w * 4 + 1) * h);
  for (let y = 0; y < h; y++) { raw[y * (w * 4 + 1)] = 0; raw.set(rgba.subarray(y * w * 4, (y + 1) * w * 4), y * (w * 4 + 1) + 1); }
  const blocks = [];
  for (let i = 0; i < raw.length || i === 0; i += 65535) {
    const part = raw.subarray(i, Math.min(raw.length, i + 65535));
    const last = i + 65535 >= raw.length ? 1 : 0;
    blocks.push(Uint8Array.of(last, part.length & 255, part.length >> 8, ~part.length & 255, (~part.length >> 8) & 255), part);
    if (!raw.length) break;
  }
  let a = 1;
  let b = 0;
  for (const v of raw) { a = (a + v) % 65521; b = (b + a) % 65521; }
  const z = concat([Uint8Array.of(0x78, 0x01), ...blocks, Uint8Array.of(b >> 8, b & 255, a >> 8, a & 255)]);
  const chunk = (type, data) => {
    const t = new TextEncoder().encode(type);
    const len = Uint8Array.of(data.length >>> 24, (data.length >> 16) & 255, (data.length >> 8) & 255, data.length & 255);
    const crc = crc32(concat([t, data]));
    return concat([len, t, data, Uint8Array.of(crc >>> 24, (crc >> 16) & 255, (crc >> 8) & 255, crc & 255)]);
  };
  const ihdr = Uint8Array.of(w >>> 24, (w >> 16) & 255, (w >> 8) & 255, w & 255, h >>> 24, (h >> 16) & 255, (h >> 8) & 255, h & 255, 8, 6, 0, 0, 0);
  const png = concat([Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10), chunk('IHDR', ihdr), chunk('IDAT', z), chunk('IEND', new Uint8Array(0))]);
  return `data:image/png;base64,${b64(png)}`;
}
function concat(parts) { let n = 0; for (const p of parts) n += p.length; const o = new Uint8Array(n); let k = 0; for (const p of parts) { o.set(p, k); k += p.length; } return o; }
const CRC = (() => { const t = new Uint32Array(256); for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
function crc32(d) { let c = 0xffffffff; for (const v of d) c = CRC[(c ^ v) & 255] ^ (c >>> 8); return (c ^ 0xffffffff) >>> 0; }
function b64(bytes) {
  if (typeof Buffer !== 'undefined') return Buffer.from(bytes.buffer, bytes.byteOffset, bytes.length).toString('base64');
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** 그림 XObject → { url, w, h } | null */
function imageOf(doc, st, res, encodeImage, fillRgb) {
  const d = st.dict;
  const v = (x) => doc.val(x);
  const w = v(d.Width ?? d.W);
  const h = v(d.Height ?? d.H);
  if (!w || !h) return null;
  const mask = v(d.ImageMask ?? d.IM) === true;
  let data;
  try { data = doc.decode(st, { stopAtImage: true }); } catch { return null; }
  const smask = v(d.SMask);
  if (data.imageFilter === 'DCTDecode' || data.imageFilter === 'DCT') {
    const url = `data:image/jpeg;base64,${b64(data)}`;
    // 투명 가리개가 있는 JPEG: 불러온 뒤(비동기) 합성하도록 alpha 를 함께 넘김
    if (smask instanceof Stream) { const alpha = alphaOf(doc, smask, w, h); if (alpha) return { url, w, h, alpha }; }
    return { url, w, h };
  }
  if (data.imageFilter) return null; // JPX · JBIG2 · CCITT
  const bpc = mask ? 1 : v(d.BitsPerComponent ?? d.BPC) ?? 8;
  const cs = mask ? null : colorSpace(doc, d.ColorSpace ?? d.CS ?? new Name('DeviceGray'), res);
  const n = mask ? 1 : cs.n || 1;
  const decodeArr = v(d.Decode ?? d.D);
  const rgba = new Uint8ClampedArray(w * h * 4);
  const rowBytes = (w * n * bpc + 7) >> 3;
  const maxV = (1 << bpc) - 1;
  const comps = new Array(n);
  for (let y = 0; y < h; y++) {
    const ro = y * rowBytes;
    for (let x = 0; x < w; x++) {
      for (let k = 0; k < n; k++) {
        const bit = (x * n + k) * bpc;
        let val;
        if (bpc === 8) val = data[ro + x * n + k] ?? 0;
        else if (bpc === 16) val = data[ro + (x * n + k) * 2] ?? 0;
        else val = ((data[ro + (bit >> 3)] ?? 0) >> (8 - bpc - (bit & 7))) & maxV;
        const mv = bpc === 16 ? 255 : maxV;
        if (cs?.indexed) comps[k] = val;
        else {
          let f = val / mv;
          if (decodeArr) f = decodeArr[2 * k] + f * (decodeArr[2 * k + 1] - decodeArr[2 * k]);
          comps[k] = f;
        }
      }
      const o = (y * w + x) * 4;
      if (mask) {
        const on = decodeArr && decodeArr[0] === 1 ? comps[0] >= 0.5 : comps[0] < 0.5;
        rgba[o] = fillRgb[0]; rgba[o + 1] = fillRgb[1]; rgba[o + 2] = fillRgb[2]; rgba[o + 3] = on ? 255 : 0;
        continue;
      }
      const [r, g, b] = cs.rgb(comps);
      rgba[o] = r; rgba[o + 1] = g; rgba[o + 2] = b; rgba[o + 3] = 255;
    }
  }
  if (smask instanceof Stream) {
    const a = alphaOf(doc, smask, w, h);
    if (a) for (let i = 0; i < w * h; i++) rgba[i * 4 + 3] = a[i];
  } else if (Array.isArray(v(d.Mask)) && !mask) {
    // 색 키 가리기
    const m = v(d.Mask);
    for (let i = 0; i < w * h; i++) {
      const px = [rgba[i * 4], rgba[i * 4 + 1], rgba[i * 4 + 2]];
      if (n === 1 && data[i] >= m[0] && data[i] <= m[1]) rgba[i * 4 + 3] = 0;
      else if (n === 3 && px.every((c, k) => c >= m[2 * k] * (255 / maxV) && c <= m[2 * k + 1] * (255 / maxV))) rgba[i * 4 + 3] = 0;
    }
  }
  const hasAlpha = mask || smask instanceof Stream || Array.isArray(v(d.Mask));
  const url = encodeImage ? encodeImage(w, h, rgba, hasAlpha) : pngDataUrl(w, h, new Uint8Array(rgba.buffer));
  return url ? { url, w, h } : null;
}
function alphaOf(doc, sm, w, h) {
  const d = sm.dict;
  const sw = doc.val(d.Width);
  const sh = doc.val(d.Height);
  const bpc = doc.val(d.BitsPerComponent) ?? 8;
  const dec = doc.val(d.Decode);
  const inv = Array.isArray(dec) && dec[0] > dec[1];
  let data;
  try { data = doc.decode(sm, { stopAtImage: true }); } catch { return null; }
  if (data.imageFilter || !sw || !sh || ![1, 2, 4, 8].includes(bpc)) return null;
  const row = (sw * bpc + 7) >> 3;
  const max = (1 << bpc) - 1;
  const a = new Uint8Array(w * h);
  for (let y = 0; y < h; y++) {
    const ro = Math.floor(y * sh / h) * row;
    for (let x = 0; x < w; x++) {
      const sx = Math.floor(x * sw / w);
      let v = bpc === 8 ? data[ro + sx] ?? 255 : ((((data[ro + ((sx * bpc) >> 3)] ?? 255) >> (8 - bpc - ((sx * bpc) & 7))) & max) * 255) / max;
      if (inv) v = 255 - v;
      a[y * w + x] = v;
    }
  }
  return a;
}

// ───────────── 쪽 해석 ─────────────
const MAX_OBJECTS = 1500;

class PageReader {
  constructor(doc, page, opts) {
    this.doc = doc;
    this.opts = opts;
    this.items = [];      // 칠한 순서: {kind:'path'|'image'|'text', ...}
    this.fonts = new Map();
    this.warn = opts.warn;
    const v = (x) => doc.val(x);
    const box = (v(page.CropBox) ?? v(page.MediaBox) ?? [0, 0, 612, 792]).map(v);
    const [x0, y0, x1, y1] = [Math.min(box[0], box[2]), Math.min(box[1], box[3]), Math.max(box[0], box[2]), Math.max(box[1], box[3])];
    const rot = (((v(page.Rotate) ?? 0) % 360) + 360) % 360;
    this.pw = (rot % 180 ? y1 - y0 : x1 - x0);
    this.ph = (rot % 180 ? x1 - x0 : y1 - y0);
    // 사용자 공간 → 슬라이드 px (y 아래로)
    const s = opts.scale;
    let base;
    if (rot === 90) base = [0, s, s, 0, -y0 * s, -x0 * s];
    else if (rot === 180) base = [-s, 0, 0, s, x1 * s, -y0 * s];
    else if (rot === 270) base = [0, -s, -s, 0, y1 * s, x1 * s];
    else base = [s, 0, 0, -s, -x0 * s, y1 * s];
    this.base = base;
  }
  font(res, name) {
    const fd = this.doc.val(this.doc.val(res?.Font)?.[name]);
    const key = fd ?? name;
    if (!this.fonts.has(key)) this.fonts.set(key, fd ? new Font(this.doc, fd) : null);
    return this.fonts.get(key);
  }
  run(content, res, ctm, depth = 0) {
    if (depth > 8) return;
    const doc = this.doc;
    const v = (x) => doc.val(x);
    const lx = new Lexer(content);
    const ops = [];
    let gs = { ctm, fill: [0, 0, 0], stroke: [0, 0, 0], fa: 1, sa: 1, lw: 1, dash: null, fcs: null, scs: null, fillPat: null, font: null, size: 12, tc: 0, tw: 0, th: 1, tl: 0, rise: 0, tr: 0 };
    const stack = [];
    let path = [];
    let cur = null;
    let start = null;
    let tm = null;
    let tlm = null;
    const pt = (x, y) => apply(gs.ctm, x, y);
    for (;;) {
      const t = lx.next();
      if (t === undefined) break;
      if (!(t instanceof Cmd)) { ops.push(t); continue; }
      const op = t.c;
      const a = ops;
      const n = (i) => (typeof a[i] === 'number' ? a[i] : 0);
      switch (op) {
        case 'q': stack.push({ ...gs }); break;
        case 'Q': if (stack.length) gs = stack.pop(); break;
        case 'cm': gs.ctm = mul([n(0), n(1), n(2), n(3), n(4), n(5)], gs.ctm); break;
        case 'w': gs.lw = n(0); break;
        case 'd': gs.dash = Array.isArray(a[0]) && a[0].length ? a[0] : null; break;
        case 'gs': {
          const eg = v(v(res?.ExtGState)?.[a[0]?.n]);
          if (eg) { if (v(eg.ca) != null) gs.fa = v(eg.ca); if (v(eg.CA) != null) gs.sa = v(eg.CA); if (v(eg.LW) != null) gs.lw = v(eg.LW); }
          break;
        }
        // 색
        case 'g': gs.fill = [n(0) * 255, n(0) * 255, n(0) * 255]; gs.fillPat = null; gs.fcs = null; break;
        case 'G': gs.stroke = [n(0) * 255, n(0) * 255, n(0) * 255]; gs.scs = null; break;
        case 'rg': gs.fill = [n(0) * 255, n(1) * 255, n(2) * 255]; gs.fillPat = null; gs.fcs = null; break;
        case 'RG': gs.stroke = [n(0) * 255, n(1) * 255, n(2) * 255]; gs.scs = null; break;
        case 'k': gs.fill = cmyk(n(0), n(1), n(2), n(3)); gs.fillPat = null; gs.fcs = null; break;
        case 'K': gs.stroke = cmyk(n(0), n(1), n(2), n(3)); gs.scs = null; break;
        case 'cs': gs.fcs = colorSpace(doc, a[0], res); gs.fill = gs.fcs.rgb(new Array(gs.fcs.n).fill(0)); gs.fillPat = null; break;
        case 'CS': gs.scs = colorSpace(doc, a[0], res); gs.stroke = gs.scs.rgb(new Array(gs.scs.n).fill(0)); break;
        case 'sc': case 'scn': {
          const last = a[a.length - 1];
          if (last instanceof Name) { gs.fillPat = this.pattern(res, last.n, gs); if (gs.fillPat?.color) gs.fill = gs.fillPat.color; break; }
          const nums = a.filter((x) => typeof x === 'number');
          gs.fill = (gs.fcs ?? (nums.length === 1 ? colorSpace(doc, new Name('DeviceGray')) : nums.length === 4 ? colorSpace(doc, new Name('DeviceCMYK')) : colorSpace(doc, new Name('DeviceRGB')))).rgb(nums);
          gs.fillPat = null;
          break;
        }
        case 'SC': case 'SCN': {
          const nums = a.filter((x) => typeof x === 'number');
          if (nums.length) gs.stroke = (gs.scs ?? (nums.length === 1 ? colorSpace(doc, new Name('DeviceGray')) : nums.length === 4 ? colorSpace(doc, new Name('DeviceCMYK')) : colorSpace(doc, new Name('DeviceRGB')))).rgb(nums);
          break;
        }
        // 경로
        case 'm': cur = pt(n(0), n(1)); start = cur; path.push(['M', ...cur]); break;
        case 'l': cur = pt(n(0), n(1)); path.push(['L', ...cur]); break;
        case 'c': { const p1 = pt(n(0), n(1)); const p2 = pt(n(2), n(3)); cur = pt(n(4), n(5)); path.push(['C', ...p1, ...p2, ...cur]); break; }
        case 'v': { const p2 = pt(n(0), n(1)); const p1 = cur ?? p2; cur = pt(n(2), n(3)); path.push(['C', ...p1, ...p2, ...cur]); break; }
        case 'y': { const p1 = pt(n(0), n(1)); cur = pt(n(2), n(3)); path.push(['C', ...p1, ...cur, ...cur]); break; }
        case 'h': path.push(['Z']); if (start) cur = start; break;
        case 're': {
          const [x, y, w, h] = [n(0), n(1), n(2), n(3)];
          const p0 = pt(x, y); const p1 = pt(x + w, y); const p2 = pt(x + w, y + h); const p3 = pt(x, y + h);
          path.push(['M', ...p0], ['L', ...p1], ['L', ...p2], ['L', ...p3], ['Z']);
          path.rect = path.length === 5 && Math.abs(gs.ctm[1]) < 1e-6 && Math.abs(gs.ctm[2]) < 1e-6;
          cur = p0; start = p0;
          break;
        }
        case 'f': case 'F': case 'f*': case 'B': case 'B*': case 'b': case 'b*': case 'S': case 's': {
          const fill = op !== 'S' && op !== 's';
          const stroke = op === 'S' || op === 's' || op.startsWith('B') || op.startsWith('b');
          if (op === 'b' || op === 'b*' || op === 's') path.push(['Z']);
          if (path.length) this.addPath(path, gs, fill, stroke, op.endsWith('*'));
          path = [];
          break;
        }
        case 'n': path = []; break;
        case 'W': case 'W*': break;
        case 'sh': {
          // 그러데이션 칠하기 (영역은 잘라내기 경로라 알 수 없음 → 건너뜀)
          break;
        }
        // 글
        case 'BT': tm = [1, 0, 0, 1, 0, 0]; tlm = tm; break;
        case 'ET': tm = null; break;
        case 'Tf': gs.font = this.font(res, a[0]?.n); gs.size = n(1); break;
        case 'Tc': gs.tc = n(0); break;
        case 'Tw': gs.tw = n(0); break;
        case 'Tz': gs.th = n(0) / 100; break;
        case 'TL': gs.tl = n(0); break;
        case 'Ts': gs.rise = n(0); break;
        case 'Tr': gs.tr = n(0); break;
        case 'Td': tlm = mul([1, 0, 0, 1, n(0), n(1)], tlm ?? [1, 0, 0, 1, 0, 0]); tm = tlm; break;
        case 'TD': gs.tl = -n(1); tlm = mul([1, 0, 0, 1, n(0), n(1)], tlm ?? [1, 0, 0, 1, 0, 0]); tm = tlm; break;
        case 'Tm': tlm = [n(0), n(1), n(2), n(3), n(4), n(5)]; tm = tlm; break;
        case 'T*': tlm = mul([1, 0, 0, 1, 0, -gs.tl], tlm ?? [1, 0, 0, 1, 0, 0]); tm = tlm; break;
        case 'Tj': case 'TJ': case "'": case '"': {
          if (op === "'" || op === '"') { if (op === '"') { gs.tw = n(0); gs.tc = n(1); } tlm = mul([1, 0, 0, 1, 0, -gs.tl], tlm ?? [1, 0, 0, 1, 0, 0]); tm = tlm; }
          if (!tm) tm = [1, 0, 0, 1, 0, 0];
          const arg = a[a.length - 1];
          tm = this.showText(op === 'TJ' ? (Array.isArray(arg) ? arg : []) : [arg], gs, tm);
          break;
        }
        case 'Do': {
          const xo = v(v(res?.XObject)?.[a[0]?.n]);
          if (!(xo instanceof Stream)) break;
          const sub = v(xo.dict.Subtype)?.n;
          if (sub === 'Image') this.addImage(xo, res, gs);
          else if (sub === 'Form') {
            let data;
            try { data = doc.decode(xo); } catch { break; }
            const m = v(xo.dict.Matrix)?.map(v) ?? [1, 0, 0, 1, 0, 0];
            const saved = { ...gs };
            this.run(data, v(xo.dict.Resources) ?? res, mul(m, gs.ctm), depth + 1);
            gs = saved;
          }
          break;
        }
        case 'BI': {
          // 인라인 그림: ID … EI
          const dict = {};
          for (let i = 0; i + 1 < a.length; i += 2) if (a[i] instanceof Name) dict[a[i].n] = a[i + 1];
          let p = lx.pos;
          const bts = lx.b;
          // 사전 끝까지 (ID 앞)
          for (;;) { const k = lx.next(); if (k === undefined || (k instanceof Cmd && k.c === 'ID')) break; if (k instanceof Name) dict[k.n] = lx.next(); }
          p = lx.pos + 1;
          let e = p;
          while (e < bts.length - 2 && !(bts[e] === 69 && bts[e + 1] === 73 && (WS[bts[e + 2]] || e + 2 === bts.length) && WS[bts[e - 1]])) e++;
          lx.pos = e + 2;
          const st = new Stream(dict, p, e, bts);
          this.addImage(st, res, gs);
          break;
        }
        default:
      }
      ops.length = 0;
    }
  }
  pattern(res, name, gs) {
    const doc = this.doc;
    const v = (x) => doc.val(x);
    let pat = v(v(res?.Pattern)?.[name]);
    const d = pat instanceof Stream ? pat.dict : pat;
    if (!d) return null;
    if (v(d.PatternType) === 2) {
      const sh = v(d.Shading);
      const sd = sh instanceof Stream ? sh.dict : sh;
      if (!sd) return null;
      const cs = colorSpace(doc, sd.ColorSpace, res);
      const fn = sd.Function;
      const c0 = cs.rgb(evalFunction(doc, fn, [0]));
      const c1 = cs.rgb(evalFunction(doc, fn, [1]));
      const cm = cs.rgb(evalFunction(doc, fn, [0.5]));
      const type = v(sd.ShadingType);
      const coords = v(sd.Coords) ?? [0, 0, 1, 0];
      const m = mul(v(d.Matrix)?.map(v) ?? [1, 0, 0, 1, 0, 0], this.base);
      let angle = 90;
      if (type === 2) { const [ax, ay] = apply(m, coords[0], coords[1]); const [bx, by] = apply(m, coords[2], coords[3]); angle = Math.round((Math.atan2(by - ay, bx - ax) * 180) / Math.PI); }
      return { color: cm, grad: { type: 'gradient', angle: ((angle % 360) + 360) % 360, path: type === 3, stops: [[0, hex(...c0)], [1, hex(...c1)]] } };
    }
    return { color: [200, 200, 200] };
  }
  addPath(path, gs, fill, stroke, evenOdd) {
    const m = mul(gs.ctm, [1, 0, 0, 1, 0, 0]);
    // 장치 좌표(사용자 단위) → 슬라이드
    const cmds = path.map(([op, ...p]) => {
      const q = [];
      for (let i = 0; i < p.length; i += 2) { const [x, y] = apply(this.base, p[i], p[i + 1]); q.push(x, y); }
      return [op, ...q];
    });
    let minX = Infinity; let minY = Infinity; let maxX = -Infinity; let maxY = -Infinity;
    for (const [, ...p] of cmds) for (let i = 0; i < p.length; i += 2) { minX = Math.min(minX, p[i]); maxX = Math.max(maxX, p[i]); minY = Math.min(minY, p[i + 1]); maxY = Math.max(maxY, p[i + 1]); }
    if (!Number.isFinite(minX)) return;
    const isRect = path.rect || rectPath(cmds);
    const scaleLw = Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) * this.opts.scale;
    const lw = Math.max(0.5, (gs.lw || 1) * scaleLw);
    this.items.push({ kind: 'path', cmds, rect: isRect, box: { x: minX, y: minY, w: maxX - minX, h: maxY - minY }, fill: fill ? (gs.fillPat?.grad ?? hex(...gs.fill)) : null, fillAlpha: gs.fa, stroke: stroke ? hex(...gs.stroke) : null, lw, dash: gs.dash, evenOdd });
  }
  addImage(st, res, gs) {
    const m = mul(gs.ctm, this.base);
    const pts = [apply(m, 0, 0), apply(m, 1, 0), apply(m, 0, 1), apply(m, 1, 1)];
    const xs = pts.map((p) => p[0]);
    const ys = pts.map((p) => p[1]);
    const box = { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
    if (box.w < 0.5 || box.h < 0.5) return;
    const key = st.start;
    let img = this.opts.imgCache.get(key);
    if (img === undefined) {
      img = imageOf(this.doc, st, res, this.opts.encodeImage, gs.fill);
      this.opts.imgCache.set(key, img);
    }
    if (!img) { this.warn('일부 그림 형식(JPEG 2000 · 팩스 등)은 열지 못했습니다'); return; }
    // 뒤집기: 이미지 공간의 (0,0)=왼쪽 아래
    const flipV = m[3] < 0 ? false : m[3] > 0;
    const flipH = m[0] < 0;
    const rot = Math.abs(m[1]) > 1e-6 || Math.abs(m[2]) > 1e-6 ? (Math.atan2(m[1], m[0]) * 180) / Math.PI : 0;
    this.items.push({ kind: 'image', img, box, flipV, flipH, rot: Math.abs(rot) > 0.5 ? rot : 0, alpha: gs.fa });
  }
  showText(arr, gs, tm) {
    const f = gs.font;
    if (!f) return tm;
    const size = gs.size;
    let text = '';
    let startTm = null;
    let advance = 0; // 글자 공간 x
    const flush = () => {
      if (!text || !startTm) { text = ''; startTm = null; return; }
      this.addText(text, gs, startTm, advance);
      text = '';
      startTm = null;
    };
    for (const it of arr) {
      if (typeof it === 'number') {
        const tx = (-it / 1000) * size * gs.th;
        if (it < -250 && text && !text.endsWith(' ')) text += ' ';
        if (it > 600 * 2) flush();
        tm = mul([1, 0, 0, 1, tx, 0], tm);
        if (startTm) advance += tx;
        continue;
      }
      if (!(it instanceof Uint8Array)) continue;
      for (const g of f.decode(it)) {
        if (!startTm) { startTm = tm; advance = 0; }
        const w0 = f.subtype === 'Type3' ? g.w * f.fm : g.w / 1000;
        const tx = (w0 * size + gs.tc + (g.space ? gs.tw : 0)) * gs.th;
        text += g.uni;
        tm = mul([1, 0, 0, 1, tx, 0], tm);
        advance += tx;
      }
    }
    flush();
    return tm;
  }
  addText(text, gs, tm, advance) {
    if (gs.tr === 3 || gs.tr === 7) return; // 보이지 않는 글 (OCR 층)
    const clean = text.replace(/[\u0000-\u0008\u000b-\u001f�]/g, '');
    if (!clean.trim()) return;
    const trm = mul(mul([gs.size * gs.th, 0, 0, gs.size, 0, gs.rise], tm), gs.ctm);
    const full = mul(trm, this.base);
    const sizePx = Math.hypot(full[2], full[3]);
    if (sizePx < 1) return;
    const angle = (Math.atan2(full[1], full[0]) * 180) / Math.PI;
    const [x0, y0] = apply(this.base, ...apply(mul(mul([1, 0, 0, 1, 0, gs.rise], tm), gs.ctm), 0, 0));
    const [x1, y1] = apply(this.base, ...apply(mul(mul([1, 0, 0, 1, advance, gs.rise], tm), gs.ctm), 0, 0));
    const f = gs.font;
    this.items.push({ kind: 'text', text: clean, x0, y0, x1, y1, sizePx, angle, color: hex(...gs.fill), alpha: gs.fa, font: f.family, b: f.bold, i: f.italic, stroke: gs.tr === 1 || gs.tr === 2 });
  }
}

/** 축에 나란한 사각형 경로인지 (M + L 3~4개 + Z) */
function rectPath(cmds) {
  const pts = [];
  for (const [op, ...p] of cmds) {
    if (op === 'Z') continue;
    if (op === 'C') return false;
    if (op === 'M' && pts.length) return false;
    pts.push(p);
  }
  if (pts.length === 5 && Math.abs(pts[4][0] - pts[0][0]) < 0.01 && Math.abs(pts[4][1] - pts[0][1]) < 0.01) pts.pop();
  if (pts.length !== 4) return false;
  const e = 0.01;
  for (let i = 0; i < 4; i++) {
    const [ax, ay] = pts[i];
    const [bx, by] = pts[(i + 1) % 4];
    if (Math.abs(ax - bx) > e && Math.abs(ay - by) > e) return false;
  }
  const xs = new Set(pts.map((q) => Math.round(q[0] * 100)));
  const ys = new Set(pts.map((q) => Math.round(q[1] * 100)));
  return xs.size === 2 && ys.size === 2;
}

/** 같은 경로를 채우기 한 번 + 선 한 번으로 그린 것 (LibreOffice · Word 등) → 도형 하나 */
function joinFillStroke(items) {
  const out = [];
  for (const it of items) {
    const prev = out[out.length - 1];
    if (it.kind === 'path' && prev?.kind === 'path' && prev.fill && !prev.stroke && !it.fill && it.stroke && prev.cmds.length === it.cmds.length &&
      prev.cmds.every((c, i) => c[0] === it.cmds[i][0] && c.every((x, k) => k === 0 || Math.abs(x - it.cmds[i][k]) < 0.05))) {
      out[out.length - 1] = { ...prev, stroke: it.stroke, lw: it.lw, dash: it.dash };
      continue;
    }
    out.push(it);
  }
  return out;
}

// ───────────── 글 조각 → 텍스트 상자 ─────────────
function groupText(runs) {
  // 1) 한 줄로 이어 붙이기
  const lines = [];
  for (const r of runs) {
    const flat = Math.abs(r.angle) < 0.5;
    let line = null;
    for (let k = lines.length - 1; k >= Math.max(0, lines.length - 6); k--) {
      const l = lines[k];
      if (!flat || Math.abs(l.angle - r.angle) > 0.5) continue;
      const s = Math.max(l.size, r.sizePx);
      if (Math.abs(l.y - r.y0) > s * 0.35) continue;
      const gap = r.x0 - l.x1;
      if (gap < -s * 0.6 || gap > s * 2.2) continue;
      line = l;
      if (gap > s * 0.18 && !l.runs[l.runs.length - 1].text.endsWith(' ') && !/^[\s,.:;!?)\]}%·、。，．）」』]/.test(r.text)) l.runs.push({ ...r, text: ' ' });
      break;
    }
    if (!line) { line = { runs: [], x0: r.x0, x1: r.x1, y: r.y0, size: r.sizePx, angle: r.angle }; lines.push(line); }
    line.runs.push(r);
    line.x1 = Math.max(line.x1, r.x1);
    line.x0 = Math.min(line.x0, r.x0);
    line.size = Math.max(line.size, r.sizePx);
  }
  // 2) 줄 → 상자 (왼쪽/가운데/오른쪽이 맞고 줄 간격이 일정한 줄들)
  const blocks = [];
  for (const l of lines) {
    let blk = null;
    for (let k = blocks.length - 1; k >= Math.max(0, blocks.length - 4); k--) {
      const b = blocks[k];
      const last = b.lines[b.lines.length - 1];
      if (Math.abs(b.angle - l.angle) > 0.5 || Math.abs(l.angle) > 0.5) continue;
      const gap = l.y - last.y;
      const s = Math.max(last.size, l.size);
      if (gap <= s * 0.6 || gap > s * 1.9) continue;
      if (l.size / last.size > 1.35 || last.size / l.size > 1.35) continue;
      if (b.lines.length > 1 && Math.abs(gap - b.gap) > s * 0.25) continue;
      const left = Math.abs(l.x0 - b.x0) < s * 0.9;
      const center = Math.abs((l.x0 + l.x1) / 2 - (b.x0 + b.x1) / 2) < s * 0.8;
      const right = Math.abs(l.x1 - b.x1) < s * 0.8;
      // 글머리 기호가 있는 줄 다음의 내어 쓴 줄
      const hang = l.x0 > b.x0 && l.x0 - b.x0 < s * 2.5 && /^[•▪■●◆❖➢✓·\-–]\s?/.test(b.lines[0].runs[0]?.text ?? '');
      if (!left && !center && !right && !hang) continue;
      // 가로로 겹쳐야 함
      if (l.x0 > b.x1 + s || l.x1 < b.x0 - s) continue;
      blk = b;
      if (b.lines.length === 1) b.gap = gap;
      b.align = left || hang ? b.align : center ? 'ctr' : 'r';
      break;
    }
    if (!blk) { blk = { lines: [], x0: l.x0, x1: l.x1, angle: l.angle, align: 'l', gap: 0 }; blocks.push(blk); }
    blk.lines.push(l);
    blk.x0 = Math.min(blk.x0, l.x0);
    blk.x1 = Math.max(blk.x1, l.x1);
  }
  return blocks;
}

function blockObject(b) {
  const first = b.lines[0];
  const s = first.size;
  const lineH = b.lines.length > 1 ? b.gap : s * 1.2;
  const lsPt = lineH / PT2PX;
  // 줄 높이 lineH 안에서 글자 기준선 위치 ≈ (lineH − s)/2 + 0.88s
  const top = first.y - ((lineH - s) / 2 + 0.88 * s);
  const height = lineH * b.lines.length;
  const pad = s * 0.08;
  const paras = b.lines.map((l) => {
    const runs = [];
    for (const r of l.runs) {
      const props = { size: Math.round((r.sizePx / PT2PX) * 2) / 2 };
      if (r.font) props.font = r.font;
      if (r.b) props.b = true;
      if (r.i) props.i = true;
      if (r.color && r.color !== '#000000') props.color = r.color;
      const prev = runs[runs.length - 1];
      if (prev && ['size', 'font', 'b', 'i', 'color'].every((k) => prev[k] === props[k])) prev.t += r.text;
      else runs.push({ t: r.text, ...props });
    }
    return { runs, align: b.align, lvl: 0, lineSpacingPt: Math.round(lsPt * 10) / 10, spcBef: 0, spcAft: 0 };
  });
  return {
    id: uid(), type: 'shape', shape: 'rect', txBox: true, fill: null, line: null,
    x: Math.round(b.x0 - pad), y: Math.round(top), w: Math.max(4, Math.round(b.x1 - b.x0 + pad * 2 + s * 0.3)), h: Math.max(4, Math.round(height)), rot: Math.abs(b.angle) > 0.5 ? Math.round(b.angle * 10) / 10 : 0,
    text: { paras, anchor: 't', insets: [0, 0, 0, 0], wrap: false, autofit: 'none' },
  };
}

function rotatedTextObject(r) {
  // 회전한 글 조각 하나 → 회전한 텍스트 상자 (중심 기준)
  const len = Math.hypot(r.x1 - r.x0, r.y1 - r.y0);
  const s = r.sizePx;
  const w = len + s * 0.3;
  const h = s * 1.2;
  const rad = (r.angle * Math.PI) / 180;
  // 기준선 시작점에서 글 상자 중심
  const cx = r.x0 + Math.cos(rad) * (len / 2) + Math.sin(rad) * (0.98 * s - h / 2) * -1;
  const cy = r.y0 + Math.sin(rad) * (len / 2) - Math.cos(rad) * (0.98 * s - h / 2);
  const props = { size: Math.round((s / PT2PX) * 2) / 2, ...(r.font ? { font: r.font } : {}), ...(r.b ? { b: true } : {}), ...(r.i ? { i: true } : {}), ...(r.color !== '#000000' ? { color: r.color } : {}) };
  return {
    id: uid(), type: 'shape', shape: 'rect', txBox: true, fill: null, line: null,
    x: Math.round(cx - w / 2), y: Math.round(cy - h / 2), w: Math.round(w), h: Math.round(h), rot: Math.round(r.angle * 10) / 10,
    text: { paras: [{ runs: [{ t: r.text, ...props }], align: 'l', lvl: 0, lineSpacingPt: (h / PT2PX), spcBef: 0, spcAft: 0 }], anchor: 't', insets: [0, 0, 0, 0], wrap: false, autofit: 'none' },
  };
}

function pathObject(p, pageW, pageH) {
  const { box } = p;
  const w = Math.max(box.w, 0.5);
  const h = Math.max(box.h, 0.5);
  const fill = p.fill ? (typeof p.fill === 'string' ? { type: 'solid', color: p.fill, ...(p.fillAlpha < 1 ? { alpha: p.fillAlpha } : {}) } : p.fill) : null;
  const line = p.stroke ? { color: p.stroke, width: Math.round(p.lw * 100) / 100, dash: p.dash ? 'dash' : 'solid' } : null;
  const base = { id: uid(), type: 'shape', fill, line, rot: 0, x: Math.round(box.x * 100) / 100, y: Math.round(box.y * 100) / 100, w: Math.round(w * 100) / 100, h: Math.round(h * 100) / 100 };
  if (p.rect) return { ...base, shape: 'rect' };
  // 가로/세로 직선 하나
  const segs = p.cmds.filter((c) => c[0] !== 'Z');
  if (!p.fill && segs.length === 2 && segs[0][0] === 'M' && segs[1][0] === 'L') {
    const [, ax, ay] = segs[0];
    const [, bx, by] = segs[1];
    return { ...base, shape: 'line', flipH: bx < ax !== by < ay && Math.abs(bx - ax) > 0.01 && Math.abs(by - ay) > 0.01, flipV: false, fill: null };
  }
  const cmds = p.cmds.map(([op, ...v]) => [op, ...v.map((x, i) => Math.round((i % 2 ? x - box.y : x - box.x) * 100) / 100)]);
  return { ...base, shape: 'rect', path: [{ w: base.w, h: base.h, cmds }], openPath: !p.fill || undefined };
}

/** 같은 모양의 연속 경로 묶기 (차트 · 그림 글꼴 등으로 개체가 너무 많을 때) */
function mergePaths(items) {
  const out = [];
  for (const it of items) {
    const prev = out[out.length - 1];
    if (it.kind === 'path' && prev?.kind === 'path' && !prev.rect && !it.rect && prev.fill === it.fill && prev.stroke === it.stroke && prev.lw === it.lw && typeof it.fill !== 'object') {
      prev.cmds.push(...it.cmds);
      const b = prev.box;
      const x0 = Math.min(b.x, it.box.x); const y0 = Math.min(b.y, it.box.y);
      const x1 = Math.max(b.x + b.w, it.box.x + it.box.w); const y1 = Math.max(b.y + b.h, it.box.y + it.box.h);
      prev.box = { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
      continue;
    }
    if (it.kind === 'path' && prev?.kind === 'path' && prev.rect && it.rect && prev.fill === it.fill && prev.stroke === it.stroke && typeof it.fill !== 'object' && out.length > MAX_OBJECTS / 2) { prev.rect = false; prev.cmds.push(...it.cmds); continue; }
    out.push(it.kind === 'path' ? { ...it, cmds: [...it.cmds] } : it);
  }
  return out;
}

/**
 * PDF 바이트 → { pres, warnings }
 * opts.encodeImage(w, h, rgbaClampedArray, hasAlpha) → dataURL (브라우저는 canvas, 없으면 PNG 직접)
 */
export function readPdf(bytes, opts = {}) {
  if (!(bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46)) {
    const head = latin1.decode(bytes.subarray(0, 1024));
    if (!head.includes('%PDF')) throw new Error('PDF 파일이 아닙니다');
  }
  const doc = new PdfDoc(bytes);
  if (doc.trailer.Encrypt) throw new Error('암호로 보호된 PDF 는 열 수 없습니다. PDF 프로그램에서 암호를 해제한 뒤 다시 열어 주세요.');
  const root = doc.val(doc.trailer.Root);
  if (!root) throw new Error('PDF 구조를 읽을 수 없습니다 (손상된 파일)');
  const pages = [];
  const walk = (node, inh, depth = 0) => {
    node = doc.val(node);
    if (!node || depth > 32) return;
    const type = doc.val(node.Type)?.n;
    const next = { Resources: node.Resources ?? inh.Resources, MediaBox: node.MediaBox ?? inh.MediaBox, CropBox: node.CropBox ?? inh.CropBox, Rotate: node.Rotate ?? inh.Rotate };
    if (type === 'Page' || (!node.Kids && node.Contents)) { pages.push({ ...node, ...next }); return; }
    for (const k of doc.val(node.Kids) ?? []) walk(k, next, depth + 1);
  };
  walk(root.Pages, {});
  if (!pages.length) throw new Error('PDF 에 쪽이 없습니다');
  const maxPages = opts.maxPages ?? 300;
  const warnings = new Set();
  const warn = (m) => warnings.add(m);
  if (pages.length > maxPages) warn(`쪽이 많아 처음 ${maxPages}쪽만 열었습니다`);
  // 슬라이드 크기 = 첫 쪽 크기 (pt → px)
  const first = new PageReader(doc, pages[0], { scale: PT2PX, warn, imgCache: new Map() });
  const pres = newPresentation({ firstLayout: null });
  pres.size = { w: Math.round(first.pw * PT2PX), h: Math.round(first.ph * PT2PX) };
  const imgCache = new Map();
  const mediaByUrl = new Map();
  const jobs = [];
  let textChars = 0;
  for (const page of pages.slice(0, maxPages)) {
    const pr0 = new PageReader(doc, page, { scale: PT2PX, warn, imgCache, encodeImage: opts.encodeImage });
    // 다른 크기의 쪽은 슬라이드 크기에 맞춤
    const sc = Math.min(pres.size.w / (pr0.pw * PT2PX), pres.size.h / (pr0.ph * PT2PX));
    const pr = Math.abs(sc - 1) > 0.01 ? new PageReader(doc, page, { scale: PT2PX * sc, warn, imgCache, encodeImage: opts.encodeImage }) : pr0;
    const contents = doc.val(page.Contents);
    const list = Array.isArray(contents) ? contents.map((c) => doc.val(c)) : [contents];
    const parts = [];
    for (const c of list) { if (c instanceof Stream) { try { parts.push(doc.decode(c), Uint8Array.of(10)); } catch (e) { warn(e.message); } } }
    const data = concat(parts);
    try { pr.run(data, doc.val(page.Resources), [1, 0, 0, 1, 0, 0]); } catch (e) { warn(`일부 내용을 읽지 못했습니다: ${e.message}`); }
    let items = joinFillStroke(pr.items);
    if (items.length > MAX_OBJECTS) items = mergePaths(items);
    const slide = { id: uid('s'), layout: 'blank', bg: { type: 'solid', color: '#FFFFFF' }, hidden: false, transition: null, notes: '', objects: [], anims: [], hideDecor: true };
    // 맨 처음 쪽 전체를 덮는 사각형 = 배경
    const W = pres.size.w;
    const H = pres.size.h;
    let start = 0;
    while (start < items.length && items[start].kind === 'path' && items[start].rect && typeof items[start].fill === 'string' && !items[start].stroke && items[start].box.x <= 1 && items[start].box.y <= 1 && items[start].box.w >= W - 2 && items[start].box.h >= H - 2) { slide.bg = { type: 'solid', color: items[start].fill }; start++; }
    // 글은 묶어서 (칠한 순서상 각 상자의 첫 조각 자리에)
    const texts = items.filter((it) => it.kind === 'text');
    textChars += texts.reduce((n, t) => n + t.text.trim().length, 0);
    const flat = texts.filter((t) => Math.abs(t.angle) < 0.5);
    const blocks = groupText(flat);
    const blockOf = new Map();
    for (const b of blocks) blockOf.set(b.lines[0].runs[0], b);
    for (const it of items.slice(start)) {
      if (it.kind === 'path') {
        if (it.box.w < 0.05 && it.box.h < 0.05) continue;
        if (!it.fill && !it.stroke) continue;
        slide.objects.push(pathObject(it, W, H));
      } else if (it.kind === 'image') {
        let id = mediaByUrl.get(it.img);
        if (!id) {
          id = uid('m');
          mediaByUrl.set(it.img, id);
          pres.media[id] = it.img.url;
          if (it.img.alpha) jobs.push({ media: id, url: it.img.url, alpha: it.img.alpha, w: it.img.w, h: it.img.h });
        }
        slide.objects.push({ id: uid(), type: 'image', media: id, x: Math.round(it.box.x), y: Math.round(it.box.y), w: Math.round(it.box.w), h: Math.round(it.box.h), rot: 0, ...(it.flipV ? { flipV: true } : {}), ...(it.flipH ? { flipH: true } : {}) });
      } else if (it.kind === 'text') {
        if (Math.abs(it.angle) >= 0.5) { slide.objects.push(rotatedTextObject(it)); continue; }
        const b = blockOf.get(it);
        if (b) slide.objects.push(blockObject(b));
      }
    }
    if (slide.objects.length > MAX_OBJECTS * 2) warn('복잡한 쪽은 도형을 묶어서 열었습니다');
    pres.slides.push(slide);
  }
  if (!textChars) warn('글자가 그림으로 된 PDF(스캔 문서)는 글을 편집할 수 없습니다');
  const info = doc.val(doc.trailer.Info);
  const str = (v) => { v = doc.val(v); if (!(v instanceof Uint8Array)) return ''; return v[0] === 0xfe && v[1] === 0xff ? utf16be(v.subarray(2)) : latin1.decode(v); };
  if (info) pres.props = { ...pres.props, title: str(info.Title), author: str(info.Author) };
  // jobs: 투명 가리개가 있는 JPEG — 브라우저에서 합성 (fileio.js)
  return { pres, warnings: [...warnings], jobs };
}
