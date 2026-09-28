// 수식 엔진 공통: 오류 값 · 배열(Range) · 참조 · 값 변환 · 조건 · 날짜 · 배열 브로드캐스트 (DOM 없음)
import { formatGeneral, parseInput } from './format.js';

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
  if (typeof v === 'number') return formatGeneral(v);
  return String(v);
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

/** 엑셀 비교 규칙: 숫자 < 문자 < 논리값, 문자는 대소문자 무시 */
export function compareValues(a, b) {
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a);
  const rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'string') {
    const x = a.toLowerCase();
    const y = b.toLowerCase();
    return x < y ? -1 : x > y ? 1 : 0;
  }
  if (typeof a === 'boolean') return (a ? 1 : 0) - (b ? 1 : 0);
  return a < b ? -1 : a > b ? 1 : 0;
}

// ───────────────────────── 인수 모으기 ─────────────────────────
/** SUM 계열 규칙: 범위 안에서는 숫자만, 직접 인수는 숫자로 변환 */
export function collectNums(args, { errors = true } = {}) {
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
    if (ch === '~' && i + 1 < pattern.length) re += escapeRe(pattern[++i]);
    else if (ch === '*') re += '.*';
    else if (ch === '?') re += '.';
    else re += escapeRe(ch);
  }
  return new RegExp(`^${re}$`, 'is');
}

export function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** COUNTIF/SUMIF 조건 → 판별 함수 */
export function makeCriteria(crit) {
  crit = scalar(crit);
  if (typeof crit === 'number') return (v) => (typeof v === 'number' && v === crit) || (typeof v === 'string' && parseNumberText(v) === crit);
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => typeof v === 'number' && v === 0; // 빈 셀 조건은 0 과 같음 (엑셀 규칙)
  const m = /^(<=|>=|<>|<|>|=)?([\s\S]*)$/.exec(String(crit));
  const op = m[1] || '=';
  const operand = m[2];
  const num = parseNumberText(operand);
  if (num !== null) {
    const cmp = (v) => {
      if (typeof v === 'number') return v;
      if (typeof v === 'string') return parseNumberText(v);
      return null;
    };
    switch (op) {
      case '=': return (v) => cmp(v) === num;
      case '<>': return (v) => cmp(v) !== num;
      case '<': return (v) => typeof v === 'number' && v < num;
      case '>': return (v) => typeof v === 'number' && v > num;
      case '<=': return (v) => typeof v === 'number' && v <= num;
      case '>=': return (v) => typeof v === 'number' && v >= num;
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
export const dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
export function serialToDate(serial) {
  if (serial < 0) throw ERR.NUM;
  const d = new Date(EPOCH + Math.round(serial * DAY_MS / 1000) * 1000);
  return {
    y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    hh: d.getUTCHours(), mm: d.getUTCMinutes(), ss: d.getUTCSeconds(), dow: d.getUTCDay(),
  };
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
    const rows = [];
    for (let r = 0; r < h; r++) {
      const row = [];
      for (let c = 0; c < w; c++) {
        const a = args.slice();
        let bad = false;
        for (const i of idx) {
          const x = args[i];
          const rr = x.height === 1 ? 0 : r;
          const cc = x.width === 1 ? 0 : c;
          if (rr >= x.height || cc >= x.width) { bad = true; break; }
          a[i] = x.rows[rr][cc];
        }
        if (bad) { row.push(ERR.NA); continue; }
        let v = attempt(() => fn(a, ...rest));
        if (v instanceof Range) v = v.height === 1 && v.width === 1 ? v.rows[0][0] : ERR.VALUE;
        row.push(v === undefined ? null : v);
      }
      rows.push(row);
    }
    return new Range(rows);
  };
  wrapped.lazy = fn.lazy;
  wrapped.liftPos = positions ?? 'all';
  return wrapped;
}

/** 두 값의 원소별 연산 (배열 크기 확장 규칙 포함) */
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
  const pick = (X, x, hx, wx, r, c) => {
    if (!X) return x;
    const rr = hx === 1 ? 0 : r;
    const cc = wx === 1 ? 0 : c;
    if (rr >= hx || cc >= wx) return ERR.NA;
    return X.rows[rr][cc];
  };
  const rows = [];
  for (let r = 0; r < h; r++) {
    const row = [];
    for (let c = 0; c < w; c++) {
      const x = pick(A, a, ha, wa, r, c);
      const y = pick(B, b, hb, wb, r, c);
      row.push(isError(x) ? x : isError(y) ? y : attempt(() => f(x, y)));
    }
    rows.push(row);
  }
  return new Range(rows);
}

/** 배열이면 원소별로, 아니면 그대로 */
export function mapValue(v, f) {
  if (!(v instanceof Range)) return f(v);
  return new Range(v.rows.map((row) => row.map((x) => (isError(x) ? x : attempt(() => f(x))))));
}

/** 결과 배열 → 1x1 이면 스칼라 */
export function simplify(v) {
  if (v instanceof Range && v.height === 1 && v.width === 1) return v.rows[0][0];
  return v;
}
