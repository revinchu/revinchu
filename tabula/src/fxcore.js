// 수식 엔진 공통: 오류 값 · 배열(Range) · 참조 · 값 변환 · 조건 · 날짜 · 배열 브로드캐스트 (DOM 없음)
import { formatGeneral, parseInput, dateParts, serialOf } from './format.js';

// ───────────────────────── 오류 값 ─────────────────────────
export class FormulaError {
  constructor(code) { this.code = code; }
  toString() { return this.code; }
}
export const ERR = {
  DIV0: new FormulaError('#DIV/0!'),
  VALUE: new FormulaError('#VALUE!'),
  REF: new FormulaError('#REF!'),
  NAME: new FormulaError('#NAME?'),
  NUM: new FormulaError('#NUM!'),
  NA: new FormulaError('#N/A'),
  NULL: new FormulaError('#NULL!'),
  CIRC: new FormulaError('#CIRC!'),
  SPILL: new FormulaError('#SPILL!'),
  CALC: new FormulaError('#CALC!'),
  BUSY: new FormulaError('#BUSY!'),
};
export const ERR_BY_CODE = Object.fromEntries(Object.values(ERR).map((e) => [e.code, e]));
export const isError = (v) => v instanceof FormulaError;
/** ERROR.TYPE 번호 */
export const ERROR_TYPE = { '#NULL!': 1, '#DIV/0!': 2, '#VALUE!': 3, '#REF!': 4, '#NAME?': 5, '#NUM!': 6, '#N/A': 7, '#SPILL!': 9, '#CALC!': 14 };

// ───────────────────────── 값 종류 ─────────────────────────
/** 2차원 배열 값. 셀 범위에서 왔으면 ref 에 { sheet, r1, c1, r2, c2 } */
/** 셀 안 그림 (셀에 배치한 그림 · IMAGE 함수 결과) — sizing 0: 맞춤, 1: 채우기, 2: 원래 크기, 3: 사용자 지정(h, w) */
export class CellImage {
  constructor({ src, alt = '', sizing = 0, h = null, w = null }) {
    this.type = 'image';
    this.src = src;
    this.alt = alt ?? '';
    this.sizing = sizing ?? 0;
    this.h = h;
    this.w = w;
  }

  toString() { return this.alt || ''; }
}

export class Range {
  constructor(rows, ref = null) { this.rows = rows; this.ref = ref; }
  get height() { return this.rows.length; }
  get width() { return this.rows[0]?.length ?? 0; }
  *values() { for (const row of this.rows) yield* row; }
  at(r, c) { return this.rows[r]?.[c]; }
}

/** 참조 (OFFSET·INDIRECT·INDEX 결과, 범위 연산자 등). sheet: 시트 이름 또는 null(수식이 있는 시트) */
export class RefValue {
  constructor(sheet, r1, c1, r2 = r1, c2 = c1) {
    this.sheet = sheet;
    this.r1 = Math.min(r1, r2); this.c1 = Math.min(c1, c2);
    this.r2 = Math.max(r1, r2); this.c2 = Math.max(c1, c2);
  }
  get single() { return this.r1 === this.r2 && this.c1 === this.c2; }
}

/** LAMBDA 값 */
export class Lambda {
  constructor(params, body, ctx) { this.params = params; this.body = body; this.ctx = ctx; }
}

// ───────────────────────── 값 변환 ─────────────────────────
export function scalar(v) {
  if (v instanceof Range) {
    if (v.height === 1 && v.width === 1) v = v.rows[0][0];
    else throw ERR.VALUE;
  }
  if (v instanceof Lambda) throw ERR.CALC;
  if (isError(v)) throw v;
  return v;
}

/** 숫자 모양 텍스트 → 숫자 (쉼표·%·통화·날짜·시간 포함). 아니면 null */
export function parseNumberText(s) {
  const t = String(s).trim();
  if (t === '') return null;
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return Number(t);
  const p = parseInput(t);
  return typeof p.value === 'number' ? p.value : null;
}

export function toNum(v) {
  v = scalar(v);
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = parseNumberText(String(v));
  if (n === null) throw ERR.VALUE;
  return n;
}

export function toStr(v) {
  v = scalar(v);
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return numberText(v);
  return String(v);
}

/**
 * 숫자 → 글자 (& · CONCAT · LEN 등): 엑셀처럼 유효 숫자 15자리 (화면의 "일반" 11자리가 아님).
 * =1/3&"" → 0.333333333333333, 1E+20, 1.23456789012346E+17, 1E-10
 */
export function numberText(n) {
  if (!Number.isFinite(n)) return formatGeneral(n);
  if (n === 0) return '0';
  const p = Number(n.toPrecision(15));
  const a = Math.abs(p);
  if (a >= 1e15 || a < 1e-9) {
    const [m, e] = p.toExponential(14).split('e');
    const mant = m.includes('.') ? m.replace(/0+$/, '').replace(/\.$/, '') : m;
    const ex = Number(e);
    return `${mant}E${ex < 0 ? '-' : '+'}${String(Math.abs(ex)).padStart(2, '0')}`;
  }
  if (a >= 1e-6) return String(p);
  // 0.000001 보다 작으면 JS 는 지수로 쓰므로 소수로 풀어 씀
  return p.toFixed(Math.min(20, 15 - Math.floor(Math.log10(a)) - 1)).replace(/0+$/, '').replace(/\.$/, '');
}

export function toBool(v) {
  v = scalar(v);
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined) return false;
  const u = String(v).toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  throw ERR.VALUE;
}

export const toInt = (v) => Math.trunc(toNum(v));
export const optNum = (v, d) => (v === undefined || v === null ? d : toNum(v));
export const optInt = (v, d) => (v === undefined || v === null ? d : toInt(v));
export const optBool = (v, d) => (v === undefined || v === null ? d : toBool(v));

export function checkNum(n) {
  if (!Number.isFinite(n)) throw ERR.NUM;
  return n;
}

export const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);

// 한국어 엑셀의 글자 순서: 기호 < 숫자 < 한글 < 영문 < 그 밖의 문자 (대소문자 무시)
// (근사 일치 MATCH · LOOKUP 의 이진 탐색, 비교 연산자, 정렬이 모두 이 순서를 씀)
const charClass = (c) => (c >= 0x30 && c <= 0x39 ? 1
  : (c >= 0xac00 && c <= 0xd7a3) || (c >= 0x1100 && c <= 0x11ff) || (c >= 0x3130 && c <= 0x318f) ? 2
    : (c >= 0x61 && c <= 0x7a) || (c >= 0xc0 && c <= 0x24f) ? 3
      : c < 0x80 || (c >= 0x2000 && c <= 0x2bff) || (c >= 0x3000 && c <= 0x303f) || (c >= 0xff00 && c <= 0xff0f) ? 0 : 4);
/** 글자 비교 (대소문자 무시, 한국어 엑셀 순서) */
export function compareText(a, b) {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  if (x === y) return 0;
  const n = Math.min(x.length, y.length);
  for (let i = 0; i < n; i++) {
    const p = x.charCodeAt(i);
    const q = y.charCodeAt(i);
    if (p === q) continue;
    // 둘 다 ASCII 가 아닌 같은 종류면 코드 순서, 아니면 종류 순서
    const cp = charClass(p);
    const cq = charClass(q);
    if (cp !== cq) return cp - cq;
    return p < q ? -1 : 1;
  }
  return x.length - y.length < 0 ? -1 : 1;
}

// 엑셀의 정렬(데이터 정렬 · 피벗 · 필터 목록)은 수식 비교와 달리 영문이 한글보다 앞
const SORT_COLLATOR = new Intl.Collator('en', { sensitivity: 'accent' });
/** 정렬용 비교: 숫자 < 문자 < 논리값, 문자는 엑셀 정렬 순서 */
export function compareSortValues(a, b) {
  if (typeof a === 'string' && typeof b === 'string') return SORT_COLLATOR.compare(a, b);
  return compareValues(a, b);
}

/** 엑셀 비교 규칙: 숫자 < 문자 < 논리값, 문자는 대소문자 무시 */
export function compareValues(a, b) {
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'string') return compareText(a, b);
  if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
  if (a === b) return 0;
  // 엑셀은 숫자를 유효 숫자 15자리로 비교 (3977975.925 = 3977975.9249999993)
  const x = r15(a);
  const y = r15(b);
  return x < y ? -1 : x > y ? 1 : 0;
}

/** 수식 계산 중 닫힌 외부 통합 문서 범위를 *IF(S) 가 읽었는지 (workbook.evalCell 이 칸마다 확인) */
export const CLOSED_BOOK = { hit: false };

// ───────────────────────── 인수 모으기 ─────────────────────────
/** SUM 계열 규칙: 범위 안에서는 숫자만, 직접 인수는 숫자로 변환 */
export function collectNums(args, { errors = true } = {}) {
  // 직접 적은 오류 인수(#REF! 등)가 범위 안의 오류보다 먼저 (엑셀: SUM(A1,#REF!) 에서 A1 이 #DIV/0! 이어도 #REF!)
  if (errors) for (const a of args) if (isError(a)) throw a;
  const out = [];
  for (const a of args) {
    if (a instanceof Range) {
      for (const v of a.values()) {
        if (isError(v)) { if (errors) throw v; continue; }
        if (typeof v === 'number') out.push(v);
      }
    } else if (a !== null && a !== undefined) {
      out.push(toNum(a));
    }
  }
  return out;
}

/** ...A 계열 규칙: 범위 안의 텍스트는 0, 논리값은 1/0 */
export function collectNumsA(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) {
      for (const v of a.values()) {
        if (isError(v)) throw v;
        if (typeof v === 'number') out.push(v);
        else if (typeof v === 'boolean') out.push(v ? 1 : 0);
        else if (typeof v === 'string') out.push(0);
      }
    } else if (a !== null && a !== undefined) out.push(toNum(a));
  }
  return out;
}

export function flat(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) out.push(...a.values());
    else out.push(a);
  }
  return out;
}

export function asRange(v) {
  if (v instanceof Range) return v;
  return new Range([[v === undefined ? null : v]]);
}

export const matrix = (v) => asRange(v).rows;
export const arr = (rows) => new Range(rows);

// ───────────────────────── 조건 (COUNTIF 등) ─────────────────────────
export function wildcardRegex(pattern) {
  let re = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    // 엑셀: ~ 는 뒤의 * ? ~ 만 글자로 만듦 ('00:00~01:00' 의 ~ 는 그냥 글자)
    if (ch === '~' && i + 1 < pattern.length && '*?~'.includes(pattern[i + 1])) re += escapeRe(pattern[++i]);
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += escapeRe(ch);
  }
  return new RegExp(`^${re}$`, 'is');
}

export function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** 엑셀은 조건 비교에서 숫자를 유효 숫자 15자리로 봄 (0.1+0.2 와 0.3 이 같음) */
export const r15 = (x) => (Number.isInteger(x) || !Number.isFinite(x) ? x : Number(x.toPrecision(15)));

/**
 * 같은 범위를 조건만 바꿔 수만 번 세는 경우(=COUNTIF($L$5:$L$4000, A1) 을 채운 열): 범위 값마다 개수를 한 번 세어 두고
 * 같음(=) · 다름(<>) 조건은 해시로 바로 답함. 범위 행 배열(범위 캐시가 공유)마다 한 번 만듦.
 */
const countIdxMemo = new WeakMap();
function countIndex(rows) {
  let idx = countIdxMemo.get(rows);
  if (idx) return idx;
  idx = { num: new Map(), str: new Map(), blank: 0, nul: 0, t: 0, f: 0, total: 0, numCount: 0, sorted: null };
  for (const row of rows) {
    for (const v of row) {
      idx.total++;
      if (typeof v === 'number') { const k = r15(v); idx.num.set(k, (idx.num.get(k) ?? 0) + 1); idx.numCount++; }
      else if (typeof v === 'string') {
        if (v === '') idx.blank++;
        const k = v.toLowerCase();
        idx.str.set(k, (idx.str.get(k) ?? 0) + 1);
        const p = parseNumberText(v);
        if (p !== null) { const q = r15(p); idx.num.set(q, (idx.num.get(q) ?? 0) + 1); }
      } else if (v === null || v === undefined) { idx.blank++; idx.nul++; } else if (v === true) idx.t++;
      else if (v === false) idx.f++;
    }
  }
  countIdxMemo.set(rows, idx);
  return idx;
}

/** 조건 개수를 해시로 셀 수 있으면 개수, 아니면 null (와일드카드 · 크기 비교 · 오류 값이 섞인 범위는 null) */
export function fastCount(range, crit) {
  const rows = range?.rows;
  if (!rows || rows.length * (rows[0]?.length ?? 0) < 64) return null;
  if (isError(crit) || (crit instanceof Range && crit.height === 1 && crit.width === 1 && isError(crit.rows[0][0]))) return null;
  crit = scalar(crit);
  if (typeof crit === 'number') return countIndex(rows).num.get(r15(crit)) ?? 0;
  if (typeof crit === 'boolean') { const x = countIndex(rows); return crit ? x.t : x.f; }
  if (typeof crit !== 'string') return null;
  // 크기 비교(> < >= <=) + 숫자: 정렬해 둔 숫자 배열에서 이분 탐색 (순위 수식 =COUNTIF($B$2:$B$9999,">"&B2) 가 O(log n))
  const cm = /^(<=|>=|<(?!>)|>)([\s\S]*)$/.exec(crit);
  if (cm) {
    let num = parseNumberText(cm[2]);
    if (num === null) return null;
    num = r15(num);
    const x = countIndex(rows);
    if (!x.sorted) {
      const a = new Float64Array(x.numCount);
      let k = 0;
      for (const row of rows) for (const v of row) if (typeof v === 'number') a[k++] = r15(v);
      a.sort();
      x.sorted = a;
    }
    const a = x.sorted;
    // lower(v): v 보다 작은 개수, upper(v): v 이하 개수
    const bound = (v, le) => { let lo = 0; let hi = a.length; while (lo < hi) { const mid = (lo + hi) >> 1; if (le ? a[mid] <= v : a[mid] < v) lo = mid + 1; else hi = mid; } return lo; };
    switch (cm[1]) {
      case '<': return bound(num, false);
      case '<=': return bound(num, true);
      case '>': return a.length - bound(num, true);
      default: return a.length - bound(num, false);
    }
  }
  const m = /^(<>|=)?([\s\S]*)$/.exec(crit);
  if (!m || /[*?~]/.test(m[2])) return null;
  const neg = m[1] === '<>';
  const operand = m[2];
  const x = countIndex(rows);
  let eq;
  const num = parseNumberText(operand);
  const up = operand.toUpperCase();
  if (num !== null) eq = x.num.get(r15(num)) ?? 0;
  else if (up === 'TRUE' || up === 'FALSE') eq = up === 'TRUE' ? x.t : x.f;
  else if (operand === '') eq = x.blank;
  else eq = x.str.get(operand.toLowerCase()) ?? 0;
  return neg ? x.total - eq : eq;
}

/** COUNTIF/SUMIF 조건 → 판별 함수 */
export function makeCriteria(crit) {
  // 오류 값 조건: 같은 오류가 든 칸만 (엑셀: =SUMIF(A:A, #N/A) 는 오류를 돌려주지 않음)
  const e = crit instanceof Range && crit.height === 1 && crit.width === 1 ? crit.rows[0][0] : crit;
  if (isError(e)) return (v) => isError(v) && v.code === e.code;
  crit = scalar(crit);
  if (typeof crit === 'number') {
    const c = r15(crit);
    return (v) => (typeof v === 'number' && r15(v) === c) || (typeof v === 'string' && r15(parseNumberText(v) ?? NaN) === c);
  }
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => typeof v === 'number' && v === 0; // 빈 셀 조건은 0 과 같음 (엑셀 규칙)
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(String(crit));
  const op = m[1] || '=';
  const operand = m[2];
  let num = parseNumberText(operand);
  if (num !== null) {
    num = r15(num);
    const cmp = (v) => {
      if (typeof v === 'number') return r15(v);
      if (typeof v === 'string') { const p = parseNumberText(v); return p === null ? null : r15(p); }
      return null;
    };
    switch (op) {
      case '=': return (v) => cmp(v) === num;
      case '<>': return (v) => cmp(v) !== num;
      case '<': return (v) => typeof v === 'number' && r15(v) < num;
      case '>': return (v) => typeof v === 'number' && r15(v) > num;
      case '<=': return (v) => typeof v === 'number' && r15(v) <= num;
      case '>=': return (v) => typeof v === 'number' && r15(v) >= num;
    }
  }
  const upper = operand.toUpperCase();
  if (upper === 'TRUE' || upper === 'FALSE') {
    const b = upper === 'TRUE';
    return op === '<>' ? (v) => v !== b : (v) => v === b;
  }
  if (op === '=' || op === '<>') {
    let test;
    if (operand === '') test = (v) => v === null || v === '';
    else {
      const re = wildcardRegex(operand);
      test = (v) => typeof v === 'string' && re.test(v);
    }
    return op === '=' ? test : (v) => !test(v);
  }
  const s = operand.toLowerCase();
  return (v) => {
    if (typeof v !== 'string') return false;
    const x = v.toLowerCase();
    switch (op) {
      case '<': return x < s;
      case '>': return x > s;
      case '<=': return x <= s;
      case '>=': return x >= s;
    }
    return false;
  };
}

// ───────────────────────── 날짜 ─────────────────────────
export const EPOCH = Date.UTC(1899, 11, 30);
export const DAY_MS = 86400000;
// 엑셀 1900 날짜 체계: 1900-03-01 앞은 없는 날 1900-02-29 때문에 하루씩 당겨짐 (1900-01-01 = 1)
export const dateToSerial = serialOf;
export function serialToDate(serial) {
  if (serial < 0) throw ERR.NUM;
  return dateParts(serial, 1000);
}
export function todaySerial() {
  const n = new Date();
  return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate());
}
/** 날짜 인수: 숫자 또는 날짜 텍스트 */
export function toDate(v) {
  const n = toNum(v);
  if (n < 0) throw ERR.NUM;
  return Math.floor(n);
}

// ───────────────────────── 함수 표시 ─────────────────────────
/** 인수를 계산하지 않고 AST 로 받는 함수 (IF 등): fn(args, ctx, ev) */
export function lazy(fn) { fn.lazy = true; return fn; }

/** 결과를 참조로도 돌려줄 수 있는 함수 표시 */
export function refFn(fn) { fn.ref = true; return fn; }

/** 오류를 던지는 대신 값으로 */
export function attempt(f) {
  try {
    return f();
  } catch (e) {
    if (e instanceof FormulaError) return e;
    throw e;
  }
}

/**
 * 스칼라 함수를 배열 인수에 원소별로 적용 (동적 배열 엑셀의 '리프팅')
 * positions: 원소별로 나눌 인수 번호 (없으면 전부)
 */
export function lift(fn, positions = null) {
  const wrapped = (args, ...rest) => {
    const idx = (positions ?? args.map((_, i) => i)).filter((i) => {
      const a = args[i];
      return a instanceof Range && (a.height > 1 || a.width > 1);
    });
    if (!idx.length) return fn(args, ...rest);
    let h = 1;
    let w = 1;
    for (const i of idx) { h = Math.max(h, args[i].height); w = Math.max(w, args[i].width); }
    // 원소마다 인수 배열 · 함수를 새로 만들지 않음 (백만 행 배열도 가볍게)
    const rows = new Array(h);
    const a = args.slice();
    const liftCall = () => {
      try {
        return fn(a.slice(), ...rest);
      } catch (e) {
        if (e instanceof FormulaError) return e;
        throw e;
      }
    };
    for (let r = 0; r < h; r++) {
      const row = new Array(w);
      for (let c = 0; c < w; c++) {
        let bad = false;
        for (let k = 0; k < idx.length; k++) {
          const i = idx[k];
          const x = args[i];
          const rr = x.height === 1 ? 0 : r;
          const cc = x.width === 1 ? 0 : c;
          if (rr >= x.height || cc >= x.width) { bad = true; break; }
          a[i] = x.rows[rr][cc];
        }
        if (bad) { row[c] = ERR.NA; continue; }
        let v = liftCall();
        if (v instanceof Range) v = v.height === 1 && v.width === 1 ? v.rows[0][0] : ERR.VALUE;
        row[c] = v === undefined ? null : v;
      }
      rows[r] = row;
    }
    return new Range(rows);
  };
  wrapped.lazy = fn.lazy;
  wrapped.liftPos = positions ?? 'all';
  return wrapped;
}

function call2(f, x, y) {
  try {
    return f(x, y);
  } catch (e) {
    if (e instanceof FormulaError) return e;
    throw e;
  }
}
function call1(f, x) {
  try {
    return f(x);
  } catch (e) {
    if (e instanceof FormulaError) return e;
    throw e;
  }
}

/** 두 값의 원소별 연산 (배열 크기 확장 규칙 포함). 백만 행 배열도 원소마다 함수를 만들지 않음 */
export function broadcast2(a, b, f) {
  const A = a instanceof Range ? a : null;
  const B = b instanceof Range ? b : null;
  if (!A && !B) return f(a, b);
  const ha = A ? A.height : 1;
  const wa = A ? A.width : 1;
  const hb = B ? B.height : 1;
  const wb = B ? B.width : 1;
  const h = Math.max(ha, hb);
  const w = Math.max(wa, wb);
  const ar = A ? A.rows : null;
  const br = B ? B.rows : null;
  const NA = ERR.NA;
  const rows = new Array(h);
  for (let r = 0; r < h; r++) {
    const ra = ar ? (ha === 1 ? ar[0] : r < ha ? ar[r] : null) : null;
    const rb = br ? (hb === 1 ? br[0] : r < hb ? br[r] : null) : null;
    const row = new Array(w);
    for (let c = 0; c < w; c++) {
      const x = !ar ? a : ra === null ? NA : wa === 1 ? ra[0] : c < wa ? ra[c] : NA;
      const y = !br ? b : rb === null ? NA : wb === 1 ? rb[0] : c < wb ? rb[c] : NA;
      row[c] = x instanceof FormulaError ? x : y instanceof FormulaError ? y : call2(f, x, y);
    }
    rows[r] = row;
  }
  return new Range(rows);
}

/** 배열이면 원소별로, 아니면 그대로 */
export function mapValue(v, f) {
  if (!(v instanceof Range)) return f(v);
  const src = v.rows;
  const rows = new Array(src.length);
  for (let r = 0; r < src.length; r++) {
    const s = src[r];
    const row = new Array(s.length);
    for (let c = 0; c < s.length; c++) { const x = s[c]; row[c] = x instanceof FormulaError ? x : call1(f, x); }
    rows[r] = row;
  }
  return new Range(rows);
}

/** 결과 배열 → 1x1 이면 스칼라 */
export function simplify(v) {
  if (v instanceof Range && v.height === 1 && v.width === 1) return v.rows[0][0];
  return v;
}

// 큰 배열의 최대 · 최소 (Math.max(...arr) 는 인수가 10만 개를 넘으면 호출 스택이 넘침)
export function maxOf(arr, init = -Infinity) {
  let m = init;
  for (let i = 0; i < arr.length; i++) if (arr[i] > m) m = arr[i];
  return m;
}
export function minOf(arr, init = Infinity) {
  let m = init;
  for (let i = 0; i < arr.length; i++) if (arr[i] < m) m = arr[i];
  return m;
}
/** dst 뒤에 src 를 붙임 (push(...src) 대신 — 큰 배열도 안전) */
export function pushAll(dst, src) {
  for (let i = 0; i < src.length; i++) dst.push(src[i]);
  return dst;
}
