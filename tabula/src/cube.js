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
    // 빈 글자("")는 빈 칸이 아님: 엑셀 피벗의 개수도 셈 (빈 칸은 null)
    for (let i = 0; i < n; i++) { const v = get(i); out[i] = v !== null && v !== undefined ? 1 : 0; }
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
  // dims: 열 번호 또는 Column (그룹화한 열)
  const dcols = dims.map((j) => (typeof j === 'number' ? cube.col(j) : j));
  const dimCodes = dcols.map((c) => c.dim().codes);
  const cards = dcols.map((c) => c.dim().keys.length || 1);
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
  const neArr = mcols.map((c) => (c && !c.numericOnly ? c.ne() : null));
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
      const ne = neArr[m];
      const v = num[row];
      if (ne ? ne[row] : v === v) s.count[g]++;
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
  const key = `s\u0001${j}\u0001${JSON.stringify(others.map(([k, set]) => [k, [...set].sort()]))}`;
  const has = cube.memoGet(key, () => {
    const fast = itemStatsRollup(cube, j, others);
    if (fast) return fast;
    const sel = filterRows(cube, others);
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

/**
 * 시트 열 블록의 열(num/str/dict) 중 행 a ~ a+n-1 → 큐브 열. 숫자만 있는 열은 복사 없이 그대로 씀
 */
export function blockColumn(bc, a, n) {
  const { num, str, dict } = bc;
  const val = (x) => (x && typeof x === 'object' ? { code: x.error } : x);
  const col = new Column(n, (i) => {
    if (str) { const s = str[a + i]; if (s >= 0) return val(dict[s]); }
    if (num) { const v = num[a + i]; if (v === v) return v; }
    return null;
  });
  col.num = () => {
    if (col._num) return col._num;
    if (num && !str) col._num = num.subarray(a, a + n);
    else if (!num) col._num = new Float64Array(n).fill(NaN);
    else {
      const out = num.slice(a, a + n);
      for (let i = 0; i < n; i++) if (str[a + i] >= 0) out[i] = NaN;
      col._num = out;
    }
    return col._num;
  };
  col.numericOnly = !str; // 숫자뿐인 열: 비어 있지 않음 = 숫자 있음 (따로 배열을 만들지 않음)
  col.ne = () => {
    if (col._ne) return col._ne;
    const out = new Uint8Array(n);
    for (let i = 0; i < n; i++) {
      if (str && str[a + i] >= 0) out[i] = 1;
      else if (num) { const v = num[a + i]; out[i] = v === v ? 1 : 0; }
    }
    col._ne = out;
    return out;
  };
  col.dim = () => {
    if (col._dim) return col._dim;
    const codes = new Int32Array(n);
    const keys = [];
    const byDict = str ? new Int32Array(dict.length).fill(-1) : null;
    const strMap = new Map(); // 사전에 같은 글자가 둘 있어도 한 항목
    const numMap = new Map();
    let empty = -1;
    for (let i = 0; i < n; i++) {
      let c;
      const s = str ? str[a + i] : -1;
      if (s >= 0) {
        c = byDict[s];
        if (c < 0) {
          const k = keyOf(val(dict[s]));
          const km = typeof k === 'string' ? strMap : numMap;
          c = km.get(k);
          if (c === undefined) { c = keys.length; keys.push(k); km.set(k, c); if (k === EMPTY) empty = c; }
          byDict[s] = c;
        }
      } else {
        const v = num ? num[a + i] : NaN;
        if (v === v) {
          c = numMap.get(v);
          if (c === undefined) { c = keys.length; keys.push(v); numMap.set(v, c); }
        } else {
          if (empty < 0) { empty = keys.length; keys.push(EMPTY); strMap.set(EMPTY, empty); }
          c = empty;
        }
      }
      codes[i] = c;
    }
    col._dim = { codes, keys, empty };
    return col._dim;
  };
  return col;
}

// ───────────── 그룹화 (날짜 · 숫자 구간) ─────────────
export const GROUP_BY = [
  { id: 'years', label: '연' },
  { id: 'quarters', label: '분기' },
  { id: 'months', label: '월' },
  { id: 'yearMonth', label: '연-월' },
  { id: 'days', label: '일' },
  { id: 'number', label: '숫자 구간' },
];
const serialDate = (v) => new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000);
const p2 = (n) => String(n).padStart(2, '0');
const isoDay = (v) => { const d = serialDate(v); return `${d.getUTCFullYear()}-${p2(d.getUTCMonth() + 1)}-${p2(d.getUTCDate())}`; };
/** 원래 값 → 그룹 키 (숫자가 아니면 그대로) */
export function groupKey(v, spec) {
  if (typeof v !== 'number') return v;
  if (spec.by === 'number') {
    const size = Number(spec.size) || 10;
    const start = Number.isFinite(Number(spec.start)) ? Number(spec.start) : 0;
    const lo = start + Math.floor((v - start) / size) * size;
    return `${lo}-${lo + size - 1}`;
  }
  // 엑셀: 시작 날짜보다 앞 · 끝 날짜보다 뒤는 '<2025-11-01' · '>2026-05-01' 항목으로
  if (spec.start !== undefined && v < Math.floor(spec.start)) return `<${isoDay(spec.start)}`;
  if (spec.end !== undefined && Math.floor(v) > Math.floor(spec.end)) return `>${isoDay(spec.end)}`;
  const d = serialDate(v);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;
  switch (spec.by) {
    case 'mdays': return `${m}월${d.getUTCDate()}일`; // 엑셀의 '일' 그룹 (연도 구분 없음)
    case 'years': return y;
    case 'quarters': return `${Math.floor((m - 1) / 3) + 1}분기`;
    case 'months': return `${m}월`;
    case 'yearMonth': return `${y}-${p2(m)}`;
    default: return `${y}-${p2(m)}-${p2(d.getUTCDate())}`;
  }
}
/** 그룹 키 정렬 순서 (월 · 분기 · 구간은 숫자 순서) */
export function groupRank(k, spec) {
  if (typeof k === 'number') return k;
  if (k === EMPTY) return Infinity;
  if (typeof k === 'string' && k[0] === '<') return -Infinity;
  if (typeof k === 'string' && k[0] === '>') return Infinity;
  if (spec.by === 'mdays') { const m = /^(\d+)월(\d+)일$/.exec(k); return m ? +m[1] * 100 + +m[2] : Infinity; }
  if (spec.by === 'months' || spec.by === 'quarters' || spec.by === 'number') { const n = parseFloat(k); return Number.isFinite(n) ? n : Infinity; }
  return null; // 글자 순서
}

/** 큐브 열 j 를 spec 으로 묶은 열 (차원만). 원래 열의 서로 다른 값마다 한 번만 계산 */
export function groupedColumn(cube, j, spec) {
  const key = `g\u0001${j}\u0001${JSON.stringify(spec)}`;
  return cube.memoGet(key, () => {
    const base = cube.col(j);
    const col = new Column(base.n, (i) => groupKey(base.get(i), spec));
    col.rollupKey = `g${j}:${JSON.stringify(spec)}`;
    col.dim = () => {
      if (col._dim) return col._dim;
      const bd = base.dim();
      const keys = [];
      const index = new Map();
      const map = new Int32Array(bd.keys.length);
      let empty = -1;
      bd.keys.forEach((k, c) => {
        const g = k === EMPTY ? EMPTY : groupKey(k, spec);
        const kk2 = `${typeof g}:${g}`;
        let code = index.get(kk2);
        if (code === undefined) { code = keys.length; keys.push(g); index.set(kk2, code); if (g === EMPTY) empty = code; }
        map[c] = code;
      });
      const codes = new Int32Array(base.n);
      const bc = bd.codes;
      for (let i = 0; i < codes.length; i++) codes[i] = map[bc[i]];
      col._dim = { codes, keys, empty };
      return col._dim;
    };
    return col;
  });
}

// ───────────── 요약 캐시 (롤업) ─────────────
// 피벗 · 슬라이서가 쓰는 필드는 몇 개뿐이라, 원본(수백만 행)을 그 필드 조합으로 한 번 미리 묶어 두면
// 이후 슬라이서 클릭 · 피벗 변경은 묶인 행(보통 수천 개)만 다시 모으면 된다. 필요한 필드가 늘면 합쳐서 다시 만든다.
const ROLLUP_MIN_ROWS = 200000;
const STAT_KEYS = ['count', 'nums', 'sum', 'sq', 'min', 'max', 'prod'];

/** 차원 열 설명 → 고유 글자 (열 번호 또는 그룹화한 열) */
const dimKey = (d) => (typeof d === 'number' ? `b${d}` : d.rollupKey);

/**
 * 필터 + 그룹화 + 누적을 한 번에: groupAggregate 와 같은 모양의 결과
 * filters: [[열 번호, Set(보이는 항목 글자)]], dims: 열 번호 | 그룹화 열(rollupKey 있음), measures: [{ col, need }]
 */
export function aggregateQuery(cube, filters, dims, measures) {
  const fs = filters.filter(([j]) => j >= 0 && j < cube.header.length);
  if (cube.n < ROLLUP_MIN_ROWS || dims.some((d) => typeof d !== 'number' && !d.rollupKey)) {
    return groupAggregate(cube, filterRows(cube, fs), dims, measures);
  }
  const need = new Map();
  for (const d of dims) need.set(dimKey(d), d);
  for (const [j] of fs) need.set(dimKey(j), j);
  const ru = ensureRollup(cube, need, measures);
  if (!ru) return groupAggregate(cube, filterRows(cube, fs), dims, measures);
  return queryRollup(cube, ru, fs, dims, measures);
}

/** 여러 피벗 · 슬라이서가 쓸 필드를 미리 모아 롤업을 한 번에 만듦 (필드마다 다시 만들지 않게) */
export function planRollup(cube, dims, measures) {
  if (cube.n < ROLLUP_MIN_ROWS) return;
  const need = new Map();
  for (const d of dims) need.set(dimKey(d), d);
  ensureRollup(cube, need, measures);
}

function ensureRollup(cube, need, measures) {
  let ru = cube.rollup;
  const slotKey = (m) => `${m.col}`;
  const hasAll = ru && [...need.keys()].every((k) => ru.dimIndex.has(k))
    && measures.every((m) => {
      const s = ru.slotIndex.get(slotKey(m));
      if (s === undefined) return false;
      if (!ru.agg) return true; // 조합이 너무 많아 롤업을 쓰지 않기로 한 경우 (측정값과 무관)
      const st = ru.agg.stats[s];
      return (!m.need?.sq || st.sq) && (!m.need?.mm || st.min) && (!m.need?.prod || st.prod);
    });
  if (hasAll) return ru.useless ? null : ru;
  // 필드 · 측정값을 합쳐 다시 만듦
  const dimMap = new Map(ru ? ru.dimList.map((d) => [dimKey(d), d]) : []);
  for (const [k, d] of need) dimMap.set(k, d);
  const slotMap = new Map(ru ? ru.slots.map((m) => [slotKey(m), { col: m.col, need: { ...m.need } }]) : []);
  for (const m of measures) {
    const cur = slotMap.get(slotKey(m));
    if (!cur) slotMap.set(slotKey(m), { col: m.col, need: { ...(m.need ?? {}) } });
    else for (const k of ['sq', 'mm', 'prod']) if (m.need?.[k]) cur.need[k] = true;
  }
  const dimList = [...dimMap.values()];
  const slots = [...slotMap.values()];
  // 조합 수가 원본의 1/3 을 넘을 것 같으면 롤업 대신 바로 계산
  let space = 1;
  for (const d of dimList) space *= (typeof d === 'number' ? cube.col(d) : d).dim().keys.length || 1;
  const useless = space > cube.n / 3 && dimList.length > 1;
  const agg = useless ? null : groupAggregate(cube, null, dimList, slots);
  ru = {
    dimList, slots, agg, useless: useless || agg.G > cube.n / 3,
    dimIndex: new Map(dimList.map((d, i) => [dimKey(d), i])),
    slotIndex: new Map(slots.map((m, i) => [slotKey(m), i])),
  };
  cube.rollup = ru;
  cube.memo.clear();
  return ru.useless ? null : ru;
}

/** 롤업 묶음만 훑어 필터 · 그룹화 · 누적 */
function queryRollup(cube, ru, fs, dims, measures) {
  const { agg } = ru;
  const RD = agg.D;
  const rc = agg.codes;
  const tests = fs.map(([j, set]) => {
    const col = cube.col(j);
    const texts = col.texts();
    const ok = new Uint8Array(texts.length);
    for (let c = 0; c < texts.length; c++) ok[c] = set.has(texts[c]) ? 1 : 0;
    return { pos: ru.dimIndex.get(dimKey(j)), ok };
  });
  const D = dims.length;
  const pos = dims.map((d) => ru.dimIndex.get(dimKey(d)));
  const cards = dims.map((d) => (typeof d === 'number' ? cube.col(d) : d).dim().keys.length || 1);
  const strides = new Array(D);
  { let s = 1; for (let d = D - 1; d >= 0; d--) { strides[d] = s; s *= cards[d]; } }
  const slotOf = measures.map((m) => ru.slotIndex.get(`${m.col}`));
  const index = new Map();
  const groupsOf = [];
  const out = { G: 0, D, codes: [], stats: measures.map((m) => { const st = {}; for (const k of STAT_KEYS) st[k] = (k === 'sq' && !m.need?.sq) || ((k === 'min' || k === 'max') && !m.need?.mm) || (k === 'prod' && !m.need?.prod) ? null : []; return st; }), dims };
  for (let g = 0; g < agg.G; g++) {
    let pass = true;
    for (const t of tests) if (!t.ok[rc[g * RD + t.pos]]) { pass = false; break; }
    if (!pass) continue;
    let key = 0;
    for (let d = 0; d < D; d++) key += rc[g * RD + pos[d]] * strides[d];
    let q = index.get(key);
    if (q === undefined) {
      q = out.G++;
      index.set(key, q);
      for (let d = 0; d < D; d++) out.codes.push(rc[g * RD + pos[d]]);
      out.stats.forEach((st) => {
        st.count.push(0); st.nums.push(0); st.sum.push(0);
        if (st.sq) st.sq.push(0);
        if (st.min) { st.min.push(Infinity); st.max.push(-Infinity); }
        if (st.prod) st.prod.push(1);
      });
      groupsOf.push(q);
    }
    for (let m = 0; m < measures.length; m++) {
      const a = out.stats[m];
      const b = agg.stats[slotOf[m]];
      a.count[q] += b.count[g];
      a.nums[q] += b.nums[g];
      a.sum[q] += b.sum[g];
      if (a.sq) a.sq[q] += b.sq[g];
      if (a.min) { if (b.min[g] < a.min[q]) a.min[q] = b.min[g]; if (b.max[g] > a.max[q]) a.max[q] = b.max[g]; }
      if (a.prod) a.prod[q] *= b.prod[g];
    }
  }
  return out;
}

/** 슬라이서 항목의 '데이터 있음' 을 롤업으로 (없으면 null → 원본으로) */
export function itemStatsRollup(cube, j, others) {
  const ru = cube.n >= ROLLUP_MIN_ROWS ? cube.rollup : null;
  if (!ru || ru.useless) return null;
  const k = dimKey(j);
  if (!ru.dimIndex.has(k) || others.some(([o]) => !ru.dimIndex.has(dimKey(o)))) return null;
  const { agg } = ru;
  const RD = agg.D;
  const p = ru.dimIndex.get(k);
  const tests = others.filter(([o]) => o >= 0).map(([o, set]) => {
    const texts = cube.col(o).texts();
    const ok = new Uint8Array(texts.length);
    for (let c = 0; c < texts.length; c++) ok[c] = set.has(texts[c]) ? 1 : 0;
    return { pos: ru.dimIndex.get(dimKey(o)), ok };
  });
  const has = new Uint8Array(cube.col(j).dim().keys.length);
  for (let g = 0; g < agg.G; g++) {
    let pass = true;
    for (const t of tests) if (!t.ok[agg.codes[g * RD + t.pos]]) { pass = false; break; }
    if (pass) has[agg.codes[g * RD + p]] = 1;
  }
  return has;
}
