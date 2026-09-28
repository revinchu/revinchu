// 열 블록 (DOM 없음): 행이 아주 많은 데이터 영역을 셀 객체 대신 열마다 형식화 배열로 저장
//
// block = { r0, c0, n, ver, cols: [BlockCol] }   — 시트의 행 r0 ~ r0+n-1, 열 c0 ~ c0+cols.length-1
// BlockCol = { num: Float64Array | null  (숫자, 아니면 NaN)
//              str: Int32Array | null    (글자 · 논리값 · 오류의 사전 번호, 없으면 -1)
//              dict: [값]                (글자 · true/false · { error: '#N/A' })
//              fmt: 서식 객체 | null      (열 전체 서식 — 다른 서식의 셀은 일반 셀로 따로 둠) }
// 1,000만 행 × 열 10개 ≈ 수백 MB (셀 객체로는 수십 GB). 수식 · 메모 · 다른 서식이 있는 칸은 일반 셀(시트의 cells)에 두고,
// 그 칸은 블록에서 빈 칸으로 둔다 (일반 셀이 항상 우선).

/** 이 행 수보다 많으면 파일을 열 때 블록으로 */
export const BLOCK_MIN_ROWS = 50000;

const NUM_RE = /^-?(?:\d+\.?\d*|\.\d+)(?:[eE][-+]?\d+)?$/;

/** 열 하나를 채우는 도구 (크기를 모르면 두 배씩 늘림) */
export class ColBuilder {
  constructor(cap = 1024) {
    this.cap = Math.max(16, cap);
    this.num = null;
    this.str = null;
    this.dict = [];
    this.index = new Map();
    this.n = 0;
  }

  grow(need) {
    let cap = this.cap;
    while (cap < need) cap *= 2;
    if (cap === this.cap) return;
    if (this.num) { const a = new Float64Array(cap).fill(NaN); a.set(this.num); this.num = a; }
    if (this.str) { const a = new Int32Array(cap).fill(-1); a.set(this.str); this.str = a; }
    this.cap = cap;
  }

  /** i 번째 값 (숫자 · 글자 · 논리값 · { error } · null) */
  set(i, v) {
    if (i >= this.cap) this.grow(i + 1);
    if (i >= this.n) this.n = i + 1;
    if (v === null || v === undefined || v === '') return;
    if (typeof v === 'number') {
      if (!Number.isFinite(v)) return;
      if (!this.num) this.num = new Float64Array(this.cap).fill(NaN);
      this.num[i] = v;
      return;
    }
    if (!this.str) this.str = new Int32Array(this.cap).fill(-1);
    const k = typeof v === 'object' ? `\u0000e${v.error}` : typeof v === 'boolean' ? `\u0000b${v}` : v;
    let code = this.index.get(k);
    if (code === undefined) { code = this.dict.length; this.dict.push(v); this.index.set(k, code); }
    this.str[i] = code;
  }

  /** 끝: 길이 n 으로 자른 열 */
  finish(n = this.n, fmt = null) {
    const cut = (a) => (a ? (a.length === n ? a : a.slice(0, n)) : null);
    if (n > this.cap) this.grow(n);
    return { num: cut(this.num), str: cut(this.str), dict: this.dict, fmt };
  }
}

/** CSV · 붙여넣기 글자 → 값 (숫자면 숫자) */
export function textValue(s) {
  if (s === '') return null;
  if (NUM_RE.test(s)) { const n = Number(s); if (Number.isFinite(n) && !(s.length > 1 && s[0] === '0' && s[1] !== '.')) return n; }
  if (s === 'TRUE' || s === 'FALSE') return s === 'TRUE';
  return s;
}

export const inBlock = (b, r, c) => r >= b.r0 && r < b.r0 + b.n && c >= b.c0 && c < b.c0 + b.cols.length;

/** 블록 칸의 값 (빈 칸은 null) */
export function blockValue(b, r, c) {
  const col = b.cols[c - b.c0];
  const i = r - b.r0;
  if (col.str) { const s = col.str[i]; if (s >= 0) return col.dict[s]; }
  if (col.num) { const v = col.num[i]; if (v === v) return v; }
  return null;
}

/** 블록 칸에 값 쓰기 (null 이면 비움) */
export function blockSet(b, r, c, v) {
  const col = b.cols[c - b.c0];
  const i = r - b.r0;
  b.ver = (b.ver ?? 0) + 1;
  if (col.str) col.str[i] = -1;
  if (col.num) col.num[i] = NaN;
  if (v === null || v === undefined || v === '') return;
  if (typeof v === 'number') {
    col.num ??= new Float64Array(b.n).fill(NaN);
    col.num[i] = v;
    return;
  }
  col.str ??= new Int32Array(b.n).fill(-1);
  col.index ??= new Map(col.dict.map((x, k) => [typeof x === 'object' && x ? `\u0000e${x.error}` : typeof x === 'boolean' ? `\u0000b${x}` : x, k]));
  const key = typeof v === 'object' ? `\u0000e${v.error}` : typeof v === 'boolean' ? `\u0000b${v}` : v;
  let code = col.index.get(key);
  if (code === undefined) { code = col.dict.length; col.dict.push(v); col.index.set(key, code); }
  col.str[i] = code;
}

/** 블록 복사본 (실행 취소 기록용) */
export function blockClone(b) {
  return {
    r0: b.r0, c0: b.c0, n: b.n, ver: b.ver ?? 0,
    cols: b.cols.map((c) => ({ num: c.num ? c.num.slice() : null, str: c.str ? c.str.slice() : null, dict: [...c.dict], fmt: c.fmt ?? null })),
  };
}

/**
 * 행/열 삽입 · 삭제에 맞춰 블록 조정 (count > 0 삽입, < 0 삭제). 블록이 없어지면 null
 */
export function blockShift(b, axis, index, count) {
  if (axis === 'row') {
    const end = b.r0 + b.n;
    if (count > 0) {
      if (index <= b.r0) { b.r0 += count; return b; }
      if (index >= end) return b;
      const at = index - b.r0;
      const n2 = b.n + count;
      for (const col of b.cols) {
        if (col.num) { const a = new Float64Array(n2).fill(NaN); a.set(col.num.subarray(0, at)); a.set(col.num.subarray(at), at + count); col.num = a; }
        if (col.str) { const a = new Int32Array(n2).fill(-1); a.set(col.str.subarray(0, at)); a.set(col.str.subarray(at), at + count); col.str = a; }
      }
      b.n = n2;
      b.ver = (b.ver ?? 0) + 1;
      return b;
    }
    const d1 = index;
    const d2 = index - count; // 지우는 행 [d1, d2)
    if (d2 <= b.r0) { b.r0 += count; return b; }
    if (d1 >= end) return b;
    const a1 = Math.max(d1, b.r0) - b.r0;
    const a2 = Math.min(d2, end) - b.r0;
    const cut = a2 - a1;
    const n2 = b.n - cut;
    if (n2 <= 0) return null;
    for (const col of b.cols) {
      if (col.num) { const a = new Float64Array(n2); a.set(col.num.subarray(0, a1)); a.set(col.num.subarray(a2), a1); col.num = a; }
      if (col.str) { const a = new Int32Array(n2); a.set(col.str.subarray(0, a1)); a.set(col.str.subarray(a2), a1); col.str = a; }
    }
    b.r0 = d1 < b.r0 ? d1 : b.r0;
    b.n = n2;
    b.ver = (b.ver ?? 0) + 1;
    return b;
  }
  const w = b.cols.length;
  const end = b.c0 + w;
  if (count > 0) {
    if (index <= b.c0) { b.c0 += count; return b; }
    if (index >= end) return b;
    const at = index - b.c0;
    b.cols.splice(at, 0, ...Array.from({ length: count }, () => ({ num: null, str: null, dict: [], fmt: null })));
    b.ver = (b.ver ?? 0) + 1;
    return b;
  }
  const d1 = index;
  const d2 = index - count;
  if (d2 <= b.c0) { b.c0 += count; return b; }
  if (d1 >= end) return b;
  const a1 = Math.max(d1, b.c0) - b.c0;
  const a2 = Math.min(d2, end) - b.c0;
  b.cols.splice(a1, a2 - a1);
  if (!b.cols.length) return null;
  b.c0 = d1 < b.c0 ? d1 : b.c0;
  b.ver = (b.ver ?? 0) + 1;
  return b;
}

// ───────────── JSON 저장 (형식화 배열 → base64) ─────────────
const TA = { Float64Array, Int32Array, Uint8Array, Uint32Array };
function toB64(u8) {
  let s = '';
  const step = 0x8000;
  for (let i = 0; i < u8.length; i += step) s += String.fromCharCode.apply(null, u8.subarray(i, i + step));
  return btoa(s);
}
function fromB64(b64) {
  const s = atob(b64);
  const u8 = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) u8[i] = s.charCodeAt(i);
  return u8;
}
/** JSON.stringify(값, jsonReplacer) — 형식화 배열도 저장 */
export function jsonReplacer(k, v) {
  if (ArrayBuffer.isView(v) && TA[v.constructor.name]) {
    return { __ta: v.constructor.name, b64: toB64(new Uint8Array(v.buffer, v.byteOffset, v.byteLength)) };
  }
  return v;
}
/** JSON.parse(글자, jsonReviver) */
export function jsonReviver(k, v) {
  if (v && typeof v === 'object' && v.__ta && TA[v.__ta]) {
    const u8 = fromB64(v.b64);
    return new TA[v.__ta](u8.buffer, u8.byteOffset, u8.byteLength / TA[v.__ta].BYTES_PER_ELEMENT);
  }
  return v;
}

// ───────────── 셀 입력 글자 (편집 · 수식 입력줄에 보이는 글자) ─────────────
const pad = (n) => String(n).padStart(2, '0');
/** 블록 값 + 열 서식 → 셀 입력 글자 (날짜 서식이면 2026-07-01) */
export function rawOf(v, fmt, textRaw) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') {
    const f = fmt?.numFmt;
    if ((f === 'date' || f === 'longdate') && Number.isInteger(v) && v > 0) {
      const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
      return `${d.getUTCFullYear()}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
    }
    return String(v);
  }
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'object') return v.error;
  return textRaw ? textRaw(v) : v;
}
