// 열 기반(columnar) 분석 엔진 (DOM 없음) — 피벗 테이블 · 슬라이서 · 피벗 차트가 함께 씀
//
// 원본(행 배열 또는 시트의 열 블록)을 열마다 형식화 배열(typed array)로 바꿔 두고,
//   · 차원(행/열/필터 필드)은 사전 부호화(dictionary encoding): 값 → 정수 코드 (Int32Array)
//   · 측정값(값 필드)은 Float64Array (숫자가 아니면 NaN) + 비어 있지 않음 표시(Uint8Array)
// 필터는 코드별 허용표(Uint8Array)로 한 번에 걸러 행 번호 목록(Uint32Array)을 만들고,
// 그룹화는 코드들을 한 정수로 합친 키로 (작으면 직접 주소표, 크면 해시) 한 번 훑어 누적한다.
// 문자열을 만들거나 행마다 객체를 만들지 않으므로 수백만~천만 행도 한 번 훑는 시간에 끝난다.
import { formatGeneral } from './format.js';

export const EMPTY = '(비어 있음)';
// 셀 안 그림은 주소로 구별 (피벗 항목 · 슬라이서에서 그림을 그대로 보여 줌)
export const IMG_KEY = '\u0000img:';
const imageByKey = new Map();

/** 셀 값 → 항목 키 (빈 칸은 EMPTY, 오류는 코드 글자, 그림은 IMG_KEY+주소) */
export const keyOf = (v) => {
  if (v === null || v === undefined || v === '') return EMPTY;
  if (typeof v === 'object') {
    if (v.type === 'image') { const k = `${IMG_KEY}${v.src}`; if (!imageByKey.has(k)) imageByKey.set(k, v); return k; }
    return String(v.code);
  }
  return v;
};
/** 그림 항목 키 → 그림 (없으면 null) */
export const imageOfKey = (k) => (typeof k === 'string' && k.startsWith(IMG_KEY) ? imageByKey.get(k) ?? { type: 'image', src: k.slice(IMG_KEY.length), alt: '' } : null);

/** 슬라이서·필터에서 쓰는 항목 글자 */
export const itemText = (v) => {
  if (v === null || v === undefined || v === '') return EMPTY;
  if (typeof v === 'number') return formatGeneral(v);
  if (typeof v === 'object') return v.type === 'image' ? v.alt || `그림 ${String(v.src).slice(-12)}` : v.code;
  const img = imageOfKey(v);
  return img ? img.alt || `그림 ${String(img.src).slice(-12)}` : String(v);
};

const collator = new Intl.Collator('ko');
/** 항목 정렬: 숫자 → 글자(한국어 순) → 빈 칸 */
export function sortKeys(keys) {
  return keys.sort((a, b) => {
    if (a === EMPTY) return 1;
    if (b === EMPTY) return -1;
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'number') return -1;
    if (typeof b === 'number') return 1;
    return collator.compare(String(a), String(b));
  });
}

/** 키 → 구별용 글자 (숫자 1 과 글자 '1' 을 다르게) */
export const kk = (k) => `${typeof k}:${k}`;

// ───────────── 열 ─────────────
/**
 * 열 하나. get(i) 로 값을 읽을 수 있으면 나머지(차원 코드 · 숫자 배열)는 처음 필요할 때 한 번 만든다.
 * 시트 블록처럼 이미 형식화 배열을 가진 열은 num/str/dict 를 넘겨 바로 씀.
 */
export class Column {
  constructor(n, get, pre = null) {
    this.n = n;
    this.get = get;
    this._dim = pre?.dim ?? null;
    this._num = pre?.num ?? null;
    this._ne = pre?.ne ?? null;
  }

  /** 차원: { codes: Int32Array, keys: [키], empty: 빈 칸 코드 | -1 } */
  dim() {
    if (this._dim) return this._dim;
    const { n, get } = this;
    const codes = new Int32Array(n);
    const keys = [];
    const numMap = new Map();
    const strMap = new Map();
    const other = new Map();
    let empty = -1;
    for (let i = 0; i < n; i++) {
      const v = get(i);
      let c;
      if (typeof v === 'number') {
        c = numMap.get(v);
        if (c === undefined) { c = keys.length; keys.push(v); numMap.set(v, c); }
      } else if (v === null || v === undefined || v === '') {
        if (empty < 0) { empty = keys.length; keys.push(EMPTY); strMap.set(EMPTY, empty); }
        c = empty;
      } else if (typeof v === 'string') {
        c = strMap.get(v);
        if (c === undefined) { c = keys.length; keys.push(v); strMap.set(v, c); }
      } else {
        const k = keyOf(v);
        const map = typeof k === 'string' ? strMap : other;
        c = map.get(k);
        if (c === undefined) { c = keys.length; keys.push(k); map.set(k, c); if (k === EMPTY) empty = c; }
      }
      codes[i] = c;
    }
    this._dim = { codes, keys, empty };
    return this._dim;
  }

  /** 숫자 배열 (숫자가 아니면 NaN) */
  num() {
    if (this._num) return this._num;
    const { n, get } = this;
    const out = new Float64Array(n);
    for (let i = 0; i < n; i++) {
      const v = get(i);
      out[i] = typeof v === 'number' && Number.isFinite(v) ? v : NaN;
    }
    this._num = out;
    return out;
  }

  /** 비어 있지 않은 칸 (개수 집계용) */
  ne() {
    if (this._ne) return this._ne;
    const { n, get } = this;
    const out = new Uint8Array(n);
    for (let i = 0; i < n; i++) { const v = get(i); out[i] = v !== null && v !== undefined && v !== '' ? 1 : 0; }
    this._ne = out;
    return out;
  }

  /** 코드별 표시 글자 (필터 · 슬라이서 비교용) */
  texts() {
    const d = this.dim();
    d.texts ??= d.keys.map((k) => itemText(k));
    return d.texts;
  }
}

// ───────────── 큐브 (원본 하나) ─────────────
/**
 * 큐브: { n, header: [이름], col(j): Column, rowOf(i): 원본 행 번호, row(i): 값 배열 }
 */
export class Cube {
  constructor(n, header, makeCol, rowAt) {
    this.n = n;
    this.header = header;
    this.cols = new Array(header.length);
    this.makeCol = makeCol;
    this.rowAt = rowAt;
    this.memo = new Map(); // 필터 · 그룹 결과 재사용
  }

  col(j) {
    if (j < 0 || j >= this.header.length) return null;
    return (this.cols[j] ??= this.makeCol(j));
  }

  /** 큐브 행 i 의 값 배열 (머리글 순서) */
  row(i) { return this.rowAt(i); }

  memoGet(key, make) {
    let v = this.memo.get(key);
    if (v === undefined) {
      if (this.memo.size > 200) this.memo.clear();
      v = make();
      this.memo.set(key, v);
    }
    return v;
  }
}

const cubeByRows = new WeakMap();
/** 행 배열(첫 행 = 머리글) → 큐브. 모든 칸이 빈 행은 뺌. 같은 배열이면 같은 큐브 */
export function cubeFromRows(rows) {
  let c = cubeByRows.get(rows);
  if (c) return c;
  const header = (rows[0] ?? []).map((h, i) => (h === null || h === '' || h === undefined ? `열${i + 1}` : String(h)));
  const idx = [];
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r];
    if (!row) continue;
    let any = false;
    for (let j = 0; j < row.length; j++) { const v = row[j]; if (v !== null && v !== '' && v !== undefined) { any = true; break; } }
    if (any) idx.push(r);
  }
  const base = Uint32Array.from(idx);
  const n = base.length;
  c = new Cube(n, header, (j) => new Column(n, (i) => rows[base[i]][j] ?? null), (i) => rows[base[i]]);
  cubeByRows.set(rows, c);
  return c;
}

// ───────────── 필터 ─────────────
/**
 * filters: [[열 번호, Set(보이는 항목 글자)]] → 통과한 행 번호 (Uint32Array). 필터가 없으면 null (= 모든 행)
 * 코드마다 한 번만 판정해 두고 코드 배열만 훑는다.
 */
export function filterRows(cube, filters) {
  const fs = filters.filter(([j]) => j >= 0 && j < cube.header.length);
  if (!fs.length) return null;
  const key = `f\u0001${JSON.stringify(fs.map(([j, set]) => [j, [...set].sort()]))}`;
  return cube.memoGet(key, () => {
    const tests = fs.map(([j, set]) => {
      const col = cube.col(j);
      const texts = col.texts();
      const ok = new Uint8Array(texts.length);
      for (let c = 0; c < texts.length; c++) ok[c] = set.has(texts[c]) ? 1 : 0;
      return { codes: col.dim().codes, ok };
    });
    const n = cube.n;
    const out = new Uint32Array(n);
    let m = 0;
    if (tests.length === 1) {
      const { codes, ok } = tests[0];
      for (let i = 0; i < n; i++) if (ok[codes[i]]) out[m++] = i;
    } else {
      const T = tests.length;
      outer: for (let i = 0; i < n; i++) {
        for (let t = 0; t < T; t++) if (!tests[t].ok[tests[t].codes[i]]) continue outer;
        out[m++] = i;
      }
    }
    return out.slice(0, m);
  });
}

// ───────────── 그룹화 · 누적 ─────────────
/**
 * 행 번호 목록(sel, null 이면 전체)을 차원 열(dims) 조합으로 묶고 측정 열(measures)을 누적.
 * measures: [{ col: 열 번호 | -1(항상 1), need: { sq, mm, prod } }]
 * → { G, codes: Int32Array(G*D) (그룹별 차원 코드), stats: [{ count, nums, sum, sq?, min?, max?, prod? }] }
 */
export function groupAggregate(cube, sel, dims, measures) {
  const D = dims.length;
  const n = sel ? sel.length : cube.n;
  const dimCodes = dims.map((j) => cube.col(j).dim().codes);
  const cards = dims.map((j) => cube.col(j).dim().keys.length || 1);
  let space = 1;
  for (const c of cards) space *= c;
  const strides = new Array(D);
  { let s = 1; for (let d = D - 1; d >= 0; d--) { strides[d] = s; s *= cards[d]; } }

  // 그룹 번호: 조합 수가 작으면 직접 주소표, 아니면 해시
  const direct = space <= (1 << 22) ? new Int32Array(space).fill(-1) : null;
  const hash = direct ? null : new Map();
  let cap = 1024;
  let G = 0;
  let gcodes = new Int32Array(cap * Math.max(1, D));
  const M = measures.length;
  const mcols = measures.map((m) => (m.col >= 0 ? cube.col(m.col) : null));
  const numArr = mcols.map((c) => (c ? c.num() : null));
  const neArr = mcols.map((c) => (c ? c.ne() : null));
  const mk = (on) => (on ? new Float64Array(cap) : null);
  const st = measures.map((m) => ({ count: mk(true), nums: mk(true), sum: mk(true), sq: mk(m.need?.sq), min: mk(m.need?.mm), max: mk(m.need?.mm), prod: mk(m.need?.prod) }));
  const grow = () => {
    const ncap = cap * 2;
    const g2 = new Int32Array(ncap * Math.max(1, D));
    g2.set(gcodes);
    gcodes = g2;
    for (const s of st) {
      for (const k of ['count', 'nums', 'sum', 'sq', 'min', 'max', 'prod']) {
        if (!s[k]) continue;
        const a = new Float64Array(ncap);
        a.set(s[k]);
        s[k] = a;
      }
    }
    cap = ncap;
  };
  const newGroup = (row) => {
    if (G >= cap) grow();
    const g = G++;
    for (let d = 0; d < D; d++) gcodes[g * D + d] = dimCodes[d][row];
    for (const s of st) {
      if (s.min) s.min[g] = Infinity;
      if (s.max) s.max[g] = -Infinity;
      if (s.prod) s.prod[g] = 1;
    }
    return g;
  };

  for (let t = 0; t < n; t++) {
    const row = sel ? sel[t] : t;
    let key = 0;
    for (let d = 0; d < D; d++) key += dimCodes[d][row] * strides[d];
    let g;
    if (direct) {
      g = direct[key];
      if (g < 0) { g = newGroup(row); direct[key] = g; }
    } else {
      g = hash.get(key);
      if (g === undefined) { g = newGroup(row); hash.set(key, g); }
    }
    for (let m = 0; m < M; m++) {
      const s = st[m];
      const num = numArr[m];
      if (!num) { s.count[g]++; s.nums[g]++; s.sum[g]++; if (s.sq) s.sq[g]++; if (s.min) { if (1 < s.min[g]) s.min[g] = 1; if (1 > s.max[g]) s.max[g] = 1; } continue; }
      if (neArr[m][row]) s.count[g]++;
      const v = num[row];
      if (v === v) {
        s.nums[g]++;
        s.sum[g] += v;
        if (s.sq) s.sq[g] += v * v;
        if (s.min) { if (v < s.min[g]) s.min[g] = v; if (v > s.max[g]) s.max[g] = v; }
        if (s.prod) s.prod[g] *= v;
      }
    }
  }
  return { G, D, codes: gcodes, stats: st, dims };
}

/** 그룹 g 의 측정 m 누적값 → 피벗 누적 객체 { count, nums, sum, sq, min, max, prod } */
export function groupAcc(agg, m, g) {
  const s = agg.stats[m];
  return {
    count: s.count[g], nums: s.nums[g], sum: s.sum[g], sq: s.sq ? s.sq[g] : 0,
    min: s.min ? s.min[g] : Infinity, max: s.max ? s.max[g] : -Infinity, prod: s.prod ? s.prod[g] : 1,
  };
}

/**
 * 슬라이서 항목: 필드 열 j 의 모든 항목과, 다른 필터(others)를 통과한 행에 그 항목이 있는지
 * → { keys: [키], texts: [글자], has: Uint8Array }
 */
export function itemStats(cube, j, others) {
  const col = cube.col(j);
  const { codes, keys } = col.dim();
  const texts = col.texts();
  const sel = filterRows(cube, others);
  const key = `s\u0001${j}\u0001${JSON.stringify(others.map(([k, set]) => [k, [...set].sort()]))}`;
  const has = cube.memoGet(key, () => {
    const out = new Uint8Array(keys.length);
    let left = keys.length;
    if (sel) {
      for (let t = 0; t < sel.length && left; t++) { const c = codes[sel[t]]; if (!out[c]) { out[c] = 1; left--; } }
    } else {
      for (let i = 0; i < codes.length && left; i++) { const c = codes[i]; if (!out[c]) { out[c] = 1; left--; } }
    }
    return out;
  });
  return { keys, texts, has };
}
