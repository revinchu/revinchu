// 열 블록 (DOM 없음): 행이 아주 많은 데이터 영역을 셀 객체 대신 열마다 형식화 배열로 저장
//
// block = { r0, c0, n, ver, cols: [BlockCol], perm? }   — 시트의 행 r0 ~ r0+n-1, 열 c0 ~ c0+cols.length-1
// perm: 정렬 결과 (보이는 행 i → 저장 위치 perm[i]). 정렬은 데이터를 옮기지 않고 이 순서만 바꿈 (실행 취소도 즉시)
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

/** 보이는 행 i → 저장 위치 */
export const physRow = (b, i) => (b.perm ? b.perm[i] : i);

/** 블록 칸의 값 (빈 칸은 null) */
export function blockValue(b, r, c) {
  const col = b.cols[c - b.c0];
  const i = b.perm ? b.perm[r - b.r0] : r - b.r0;
  if (col.str) { const s = col.str[i]; if (s >= 0) return col.dict[s]; }
  if (col.num) { const v = col.num[i]; if (v === v) return v; }
  return null;
}

/** 블록 칸에 값 쓰기 (null 이면 비움) */
export function blockSet(b, r, c, v) {
  const col = b.cols[c - b.c0];
  const i = b.perm ? b.perm[r - b.r0] : r - b.r0;
  b.ver = (b.ver ?? 0) + 1;
  b.dver = (b.dver ?? 0) + 1; // 값이 바뀜 (정렬 순서만 바뀌면 dver 는 그대로)
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

/** 기존 JSON 저장본의 typed array는 숫자 키 객체입니다. null 숫자는 원래 NaN(빈칸).
 * 살아 있는 typed array의 복사 경로와 JSON 복원 경로를 구분합니다.
 */
function restoreBlockArray(value, Type, length, empty) {
  if (ArrayBuffer.isView(value)) return value.slice();
  const result = new Type(length).fill(empty);
  for (const key of Object.keys(value)) {
    const i = Number(key), v = value[key];
    if (Number.isSafeInteger(i) && i >= 0 && i < length && String(i) === key && typeof v === 'number' && Number.isFinite(v)) result[i] = v;
  }
  return result;
}

/** 블록 복사본 (실행 취소 기록용) */
export function blockClone(b) {
  return {
    r0: b.r0, c0: b.c0, n: b.n, ver: b.ver ?? 0, ...(b.perm ? { perm: restoreBlockArray(b.perm, Int32Array, b.n, 0) } : {}),
    cols: b.cols.map((c) => ({ num: c.num ? restoreBlockArray(c.num, Float64Array, b.n, NaN) : null, str: c.str ? restoreBlockArray(c.str, Int32Array, b.n, -1) : null, dict: [...c.dict], fmt: c.fmt ?? null })),
  };
}

/**
 * 행/열 삽입 · 삭제에 맞춰 블록 조정 (count > 0 삽입, < 0 삭제). 블록이 없어지면 null
 */
export function blockShift(b, axis, index, count) {
  if (axis === 'row') {
    materialize(b); // 정렬 순서가 있으면 먼저 실제 순서로
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
      b.dver = (b.dver ?? 0) + 1;
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
    b.dver = (b.dver ?? 0) + 1;
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
    b.dver = (b.dver ?? 0) + 1;
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

// ───────────── 정렬 (천만 행도 O(n)) ─────────────
/**
 * 블록 열의 행 a ~ a+n-1 을 정렬하는 순서(상대 번호 Uint32Array). 빈 칸은 항상 끝, 같은 값은 원래 순서(안정 정렬)
 * 글자 열: 사전 항목 순위로 계수 정렬 · 숫자 열: 64비트 기수 정렬 · 섞인 열: 비교 정렬
 * cmp(x, y): 사전 값 비교 (엑셀 순서: 숫자 < 글자 < 논리값, 오류는 끝)
 */
export function sortOrder(col, a, n, asc, cmp) {
  const { num, str, dict } = col;
  const blank = (i) => !(str && str[a + i] >= 0) && !(num && num[a + i] === num[a + i]);
  if (str && !hasNum(num, a, n)) return countingOrder(str, dict, a, n, asc, cmp);
  if (num && !hasStr(str, a, n)) return radixOrder(num, a, n, asc);
  // 숫자와 글자가 섞인 열
  const idx = new Uint32Array(n);
  for (let i = 0; i < n; i++) idx[i] = i;
  const val = (i) => { const s = str ? str[a + i] : -1; return s >= 0 ? dict[s] : num[a + i]; };
  const errOf = (v) => v && typeof v === 'object';
  idx.sort((x, y) => {
    const bx = blank(x);
    const by = blank(y);
    if (bx || by) return bx === by ? x - y : bx ? 1 : -1;
    const vx = val(x);
    const vy = val(y);
    const ex = errOf(vx);
    const ey = errOf(vy);
    if (ex || ey) return ex === ey ? x - y : ex ? 1 : -1;
    const c = cmp(vx, vy);
    return (asc ? c : -c) || x - y;
  });
  return idx;
}
const hasNum = (num, a, n) => { if (!num) return false; for (let i = 0; i < n; i++) if (num[a + i] === num[a + i]) return true; return false; };
const hasStr = (str, a, n) => { if (!str) return false; for (let i = 0; i < n; i++) if (str[a + i] >= 0) return true; return false; };

function countingOrder(str, dict, a, n, asc, cmp) {
  // 사전 항목 순위 (오류는 끝)
  const order = dict.map((_, i) => i).sort((x, y) => {
    const ex = dict[x] && typeof dict[x] === 'object';
    const ey = dict[y] && typeof dict[y] === 'object';
    if (ex || ey) return ex === ey ? 0 : ex ? 1 : -1;
    const c = cmp(dict[x], dict[y]);
    return asc ? c : -c;
  });
  const K = dict.length;
  const rank = new Int32Array(K);
  // 오류는 내림차순이어도 끝에
  let r = 0;
  for (const i of order) rank[i] = r++;
  const counts = new Uint32Array(K + 2);
  for (let i = 0; i < n; i++) { const s = str[a + i]; counts[(s >= 0 ? rank[s] : K) + 1]++; }
  for (let k = 1; k < counts.length; k++) counts[k] += counts[k - 1];
  const out = new Uint32Array(n);
  for (let i = 0; i < n; i++) { const s = str[a + i]; out[counts[s >= 0 ? rank[s] : K]++] = i; }
  return out;
}

/** 숫자 기수 정렬 (IEEE 754 비트를 정렬 가능한 부호 없는 정수로 바꿔 16비트씩 4번, 안정) */
function radixOrder(num, a, n, asc) {
  const hi = new Uint32Array(n);
  const lo = new Uint32Array(n);
  const rows = new Uint32Array(n); // k 번째 숫자 칸의 원래 행
  const f = new Float64Array(1);
  const u = new Uint32Array(f.buffer);
  const blanks = [];
  let m = 0;
  for (let i = 0; i < n; i++) {
    const v = num[a + i];
    if (v !== v) { blanks.push(i); continue; }
    f[0] = v === 0 ? 0 : v; // -0 → 0
    let h = u[1];
    let l = u[0];
    if (h & 0x80000000) { h = ~h >>> 0; l = ~l >>> 0; } else h = (h | 0x80000000) >>> 0;
    if (!asc) { h = ~h >>> 0; l = ~l >>> 0; }
    hi[m] = h;
    lo[m] = l;
    rows[m++] = i;
  }
  const cnt = new Uint32Array(65537);
  let cur = new Uint32Array(m);
  for (let k = 0; k < m; k++) cur[k] = k;
  let nxt = new Uint32Array(m);
  for (let d = 0; d < 4; d++) {
    // 자리(16비트)마다 계수 정렬 — 자리 값을 미리 한 배열에 뽑아 두면 반복문이 단순해져 빠름
    const src = d < 2 ? lo : hi;
    const shift = d % 2 ? 16 : 0;
    cnt.fill(0);
    let same = true;
    const first = (src[cur[0]] >>> shift) & 0xffff;
    for (let k = 0; k < m; k++) { const v = (src[cur[k]] >>> shift) & 0xffff; cnt[v + 1]++; if (v !== first) same = false; }
    if (same) continue; // 이 자리는 모두 같음 (예: 정수만 있으면 아래 자리들)
    for (let q = 1; q < cnt.length; q++) cnt[q] += cnt[q - 1];
    for (let k = 0; k < m; k++) { const x = cur[k]; nxt[cnt[(src[x] >>> shift) & 0xffff]++] = x; }
    const t = cur; cur = nxt; nxt = t;
  }
  const out = new Uint32Array(n);
  for (let k = 0; k < m; k++) out[k] = rows[cur[k]];
  for (let k = 0; k < blanks.length; k++) out[m + k] = blanks[k];
  return out;
}

/** 블록의 행 a ~ a+n-1, 열 c1~c2(블록 안 번호)를 order 순서로 재배치 (inverse 면 되돌리기) */
export function blockPermute(b, a, n, j1, j2, order, inverse = false) {
  for (let j = j1; j <= j2; j++) {
    const col = b.cols[j];
    for (const k of ['num', 'str']) {
      const arr = col[k];
      if (!arr) continue;
      const part = arr.slice(a, a + n);
      if (!inverse) for (let i = 0; i < n; i++) arr[a + i] = part[order[i]];
      else for (let i = 0; i < n; i++) arr[a + order[i]] = part[i];
    }
  }
  b.ver = (b.ver ?? 0) + 1;
  b.dver = (b.dver ?? 0) + 1;
}

/** 정렬 순서(perm)를 실제 데이터 순서로 반영하고 perm 을 없앰 */
export function materialize(b) {
  if (!b.perm) return;
  const p = b.perm;
  for (const col of b.cols) {
    for (const k of ['num', 'str']) {
      const arr = col[k];
      if (!arr) continue;
      const out = new arr.constructor(arr.length);
      for (let i = 0; i < p.length; i++) out[i] = arr[p[i]];
      col[k] = out;
    }
  }
  b.perm = null;
  b.ver = (b.ver ?? 0) + 1;
  b.dver = (b.dver ?? 0) + 1;
}

/** 보이는 순서의 열 배열 (행 a ~ a+n-1). 정렬 순서가 없으면 복사하지 않음 */
export function logicalCol(b, j, a, n) {
  const col = b.cols[j];
  if (!b.perm) return { num: col.num, str: col.str, dict: col.dict, a };
  const p = b.perm;
  const gather = (arr) => {
    if (!arr) return null;
    const out = new arr.constructor(n);
    for (let i = 0; i < n; i++) out[i] = arr[p[a + i]];
    return out;
  };
  return { num: gather(col.num), str: gather(col.str), dict: col.dict, a: 0 };
}

/**
 * 블록 전체 너비의 행 a ~ a+n-1 을 order(상대 번호) 순서로 — 정렬 순서만 바꿈. 되돌리기용 이전 순서를 돌려줌
 */
export function reorderRows(b, a, n, order) {
  if (!b.perm) { b.perm = new Uint32Array(b.n); for (let i = 0; i < b.n; i++) b.perm[i] = i; }
  const p = b.perm;
  const prev = p.slice(a, a + n);
  for (let i = 0; i < n; i++) p[a + i] = prev[order[i]];
  b.ver = (b.ver ?? 0) + 1;
  return prev;
}

/** reorderRows 되돌리기 / 다시 하기 */
export function setRowOrder(b, a, seg) {
  if (!b.perm) { b.perm = new Uint32Array(b.n); for (let i = 0; i < b.n; i++) b.perm[i] = i; }
  b.perm.set(seg, a);
  b.ver = (b.ver ?? 0) + 1;
}
