// Tabula 수식 엔진: 토크나이저 · 파서 · 평가기 · 참조 재작성
import { formatGeneral, formatWithPattern } from './format.js';

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
};
const ERR_BY_CODE = Object.fromEntries(Object.values(ERR).map((e) => [e.code, e]));
export const isError = (v) => v instanceof FormulaError;

// ───────────────────────── 주소 유틸 ─────────────────────────
// 엑셀(1,048,576행)보다 많은 1,000만 행 지원. .xlsx 로 저장할 때만 엑셀 한도가 적용됨
export const MAX_ROWS = 10_000_000;
export const EXCEL_MAX_ROWS = 1_048_576;
export const MAX_COLS = 16384;

export function colToName(c) {
  let s = '';
  let n = c + 1;
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function nameToCol(name) {
  let c = 0;
  for (const ch of name.toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
  return c - 1;
}

export const cellName = (r, c) => colToName(c) + (r + 1);

export function parseCellName(s) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(s.trim());
  if (!m) return null;
  const r = +m[2] - 1;
  const c = nameToCol(m[1]);
  if (r < 0 || r >= MAX_ROWS || c >= MAX_COLS) return null;
  return { r, c };
}

/** "A1" 또는 "A1:C3" → {r1,c1,r2,c2} */
export function parseRangeName(s) {
  const parts = s.trim().split(':');
  if (parts.length > 2) return null;
  const a = parseCellName(parts[0]);
  const b = parts.length === 2 ? parseCellName(parts[1]) : a;
  if (!a || !b) return null;
  return {
    r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c),
    r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c),
  };
}

export function quoteSheetName(name) {
  return /^[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !parseCellName(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

// ───────────────────────── 토크나이저 ─────────────────────────
const SHEET = String.raw`(?:'((?:[^']|'')+)'|([A-Za-z_À-￿][\w.À-￿]*))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d+)`;
const RANGE_RE = new RegExp(`(?:${SHEET})?${CELL}(?::${CELL})?(?![\\w(!])`, 'y');
const COLS_RE = new RegExp(`(?:${SHEET})?(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})(?![\\w(!])`, 'y');
const NUMBER_RE = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const IDENT_RE = /[A-Za-z_À-￿][\w.À-￿]*/y;
const ERROR_RE = /#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|NULL!|CIRC!)/y;
const OPS = ['<=', '>=', '<>', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%'];

function sticky(re, src, pos) {
  re.lastIndex = pos;
  return re.exec(src);
}

export function tokenize(src) {
  const toks = [];
  let i = 0;
  while (i < src.length) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    const start = i;
    if (ch === '"') {
      let j = i + 1;
      let v = '';
      for (;;) {
        if (j >= src.length) throw new SyntaxError('닫는 따옴표가 없습니다');
        if (src[j] === '"') {
          if (src[j + 1] === '"') { v += '"'; j += 2; continue; }
          break;
        }
        v += src[j++];
      }
      i = j + 1;
      toks.push({ t: 'str', v, s: start, e: i });
      continue;
    }
    if (ch === '#') {
      const m = sticky(ERROR_RE, src, i);
      if (!m) throw new SyntaxError('알 수 없는 오류 값');
      i += m[0].length;
      toks.push({ t: 'err', v: m[0], s: start, e: i });
      continue;
    }
    if (/[\d.]/.test(ch)) {
      const m = sticky(NUMBER_RE, src, i);
      if (!m) throw new SyntaxError('잘못된 숫자');
      i += m[0].length;
      toks.push({ t: 'num', v: parseFloat(m[0]), s: start, e: i });
      continue;
    }
    if (/[A-Za-z_$'À-￿]/.test(ch)) {
      let m = sticky(RANGE_RE, src, i);
      if (m && validRangeMatch(m)) {
        i += m[0].length;
        toks.push({ t: 'ref', ref: refFromRangeMatch(m), s: start, e: i });
        continue;
      }
      m = sticky(COLS_RE, src, i);
      if (m) {
        i += m[0].length;
        toks.push({ t: 'ref', ref: refFromColsMatch(m), s: start, e: i });
        continue;
      }
      m = sticky(IDENT_RE, src, i);
      if (m) {
        i += m[0].length;
        const up = m[0].toUpperCase();
        let j = i;
        while (j < src.length && /\s/.test(src[j])) j++;
        if (src[j] === '(') toks.push({ t: 'func', v: up, s: start, e: i });
        else if (up === 'TRUE' || up === 'FALSE') toks.push({ t: 'bool', v: up === 'TRUE', s: start, e: i });
        else toks.push({ t: 'ident', v: m[0], s: start, e: i });
        continue;
      }
      throw new SyntaxError(`잘못된 문자: ${ch}`);
    }
    if (ch === '(' || ch === ')' || ch === ',' || ch === ':') {
      i++;
      toks.push({ t: ch, s: start, e: i });
      continue;
    }
    const op = OPS.find((o) => src.startsWith(o, i));
    if (op) {
      i += op.length;
      toks.push({ t: 'op', v: op, s: start, e: i });
      continue;
    }
    throw new SyntaxError(`잘못된 문자: ${ch}`);
  }
  return toks;
}

function sheetFromMatch(quoted, plain) {
  if (quoted != null) return quoted.replace(/''/g, "'");
  return plain ?? null;
}

function validRangeMatch(m) {
  const ok = (col, row) => {
    const r = +row - 1;
    return r >= 0 && r < MAX_ROWS && nameToCol(col) < MAX_COLS;
  };
  return ok(m[4], m[6]) && (m[8] == null || ok(m[8], m[10]));
}

function refFromRangeMatch(m) {
  const sheet = sheetFromMatch(m[1], m[2]);
  const r1 = +m[6] - 1;
  const c1 = nameToCol(m[4]);
  const ref = { sheet, r1, c1, ar1: !!m[5], ac1: !!m[3], cols: false };
  if (m[8] != null) {
    Object.assign(ref, { r2: +m[10] - 1, c2: nameToCol(m[8]), ar2: !!m[9], ac2: !!m[7], range: true });
  } else {
    Object.assign(ref, { r2: r1, c2: c1, ar2: ref.ar1, ac2: ref.ac1, range: false });
  }
  return ref;
}

function refFromColsMatch(m) {
  return {
    sheet: sheetFromMatch(m[1], m[2]),
    r1: 0, r2: MAX_ROWS - 1, ar1: true, ar2: true,
    c1: nameToCol(m[4]), c2: nameToCol(m[6]), ac1: !!m[3], ac2: !!m[5],
    range: true, cols: true,
  };
}

// ───────────────────────── 파서 ─────────────────────────
const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

/** 수식 본문(맨 앞 '=' 제외)을 AST로 변환 */
export function parse(src) {
  const toks = tokenize(src);
  let i = 0;
  const peek = () => toks[i];
  const next = () => toks[i++];
  const fail = (msg) => { throw new SyntaxError(msg); };

  function expr(minPrec) {
    let left = unary();
    for (;;) {
      const k = peek();
      if (!k || k.t !== 'op') break;
      if (k.v === '%') { next(); left = { type: 'pct', a: left }; continue; }
      const p = BIN_PREC[k.v];
      if (p === undefined || p < minPrec) break;
      next();
      left = { type: 'bin', op: k.v, a: left, b: expr(p + 1) };
    }
    return left;
  }

  function unary() {
    const k = peek();
    if (k && k.t === 'op' && (k.v === '-' || k.v === '+')) {
      next();
      const a = unary();
      return k.v === '-' ? { type: 'neg', a } : { type: 'pos', a };
    }
    return primary();
  }

  function primary() {
    const k = next();
    if (!k) fail('수식이 완전하지 않습니다');
    switch (k.t) {
      case 'num': return { type: 'num', v: k.v };
      case 'str': return { type: 'str', v: k.v };
      case 'bool': return { type: 'bool', v: k.v };
      case 'err': return { type: 'err', v: k.v };
      case 'ref': return { type: 'ref', ref: k.ref };
      case 'ident': return { type: 'name', v: k.v };
      case '(': {
        const e = expr(0);
        if (next()?.t !== ')') fail('닫는 괄호가 없습니다');
        return e;
      }
      case 'func': {
        if (next()?.t !== '(') fail('여는 괄호가 없습니다');
        const args = [];
        if (peek()?.t === ')') { next(); return { type: 'func', name: k.v, args }; }
        for (;;) {
          const t = peek();
          if (t && (t.t === ',' || t.t === ')')) args.push({ type: 'empty' });
          else args.push(expr(0));
          const sep = next();
          if (!sep) fail('닫는 괄호가 없습니다');
          if (sep.t === ')') break;
          if (sep.t !== ',') fail('인수 구분 기호가 잘못되었습니다');
        }
        return { type: 'func', name: k.v, args };
      }
      default:
        return fail('수식에 문제가 있습니다');
    }
  }

  const ast = expr(0);
  if (i < toks.length) fail('수식에 문제가 있습니다');
  return ast;
}

// ───────────────────────── 값 변환 ─────────────────────────
export class Range {
  constructor(rows) { this.rows = rows; }
  get height() { return this.rows.length; }
  get width() { return this.rows[0]?.length ?? 0; }
  *values() { for (const row of this.rows) yield* row; }
}

function scalar(v) {
  if (v instanceof Range) {
    if (v.height === 1 && v.width === 1) v = v.rows[0][0];
    else throw ERR.VALUE;
  }
  if (isError(v)) throw v;
  return v;
}

export function parseNumberText(s) {
  const t = s.trim();
  if (t === '') return null;
  if (/^[+-]?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(t)) return Number(t);
  if (/^[+-]?\d{1,3}(,\d{3})+(\.\d+)?$/.test(t)) return Number(t.replace(/,/g, ''));
  const pct = /^([+-]?(\d+\.?\d*|\.\d+))%$/.exec(t);
  if (pct) return Number(pct[1]) / 100;
  return null;
}

function toNum(v) {
  v = scalar(v);
  if (typeof v === 'number') return v;
  if (v === null || v === undefined) return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = parseNumberText(String(v));
  if (n === null) throw ERR.VALUE;
  return n;
}

function toStr(v) {
  v = scalar(v);
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return formatGeneral(v);
  return String(v);
}

function toBool(v) {
  v = scalar(v);
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined) return false;
  const u = String(v).toUpperCase();
  if (u === 'TRUE') return true;
  if (u === 'FALSE') return false;
  throw ERR.VALUE;
}

function toInt(v) { return Math.trunc(toNum(v)); }

function checkNum(n) {
  if (!Number.isFinite(n)) throw ERR.NUM;
  return n;
}

const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);

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

// ───────────────────────── 평가기 ─────────────────────────
/**
 * ctx: {
 *   cell(sheet, r, c) → 값 (null = 빈 셀), 오류 값일 수 있음
 *   range(sheet, r1, c1, r2, c2) → 2차원 값 배열
 *   usedRows(sheet) → 사용된 행 수 (열 전체 참조 A:A 용)
 * }
 */
export function evaluate(node, ctx) {
  switch (node.type) {
    case 'num':
    case 'str':
    case 'bool':
      return node.v;
    case 'err':
      throw ERR_BY_CODE[node.v] ?? ERR.VALUE;
    case 'empty':
      return null;
    case 'name':
      throw ERR.NAME;
    case 'ref': {
      const f = node.ref;
      if (!f.range) {
        const v = ctx.cell(f.sheet, f.r1, f.c1);
        if (isError(v)) throw v;
        return v;
      }
      const r2 = f.cols ? Math.max(0, ctx.usedRows(f.sheet) - 1) : f.r2;
      return new Range(ctx.range(f.sheet, f.r1, f.c1, r2, f.c2));
    }
    case 'neg': return -toNum(evaluate(node.a, ctx));
    case 'pos': return scalar(evaluate(node.a, ctx));
    case 'pct': return toNum(evaluate(node.a, ctx)) / 100;
    case 'bin': return binary(node.op, evaluate(node.a, ctx), evaluate(node.b, ctx));
    case 'func': {
      const fn = FUNCS[node.name];
      if (!fn) throw ERR.NAME;
      if (fn.lazy) return fn(node.args, ctx);
      return fn(node.args.map((a) => evaluate(a, ctx)));
    }
    default:
      throw ERR.VALUE;
  }
}

function binary(op, a, b) {
  switch (op) {
    case '+': return checkNum(toNum(a) + toNum(b));
    case '-': return checkNum(toNum(a) - toNum(b));
    case '*': return checkNum(toNum(a) * toNum(b));
    case '/': {
      const x = toNum(a);
      const y = toNum(b);
      if (y === 0) throw ERR.DIV0;
      return checkNum(x / y);
    }
    case '^': {
      const x = toNum(a);
      const y = toNum(b);
      if (x === 0 && y === 0) throw ERR.NUM;
      return checkNum(x ** y);
    }
    case '&': return toStr(a) + toStr(b);
    default: {
      const c = compareValues(scalar(a), scalar(b));
      switch (op) {
        case '=': return c === 0;
        case '<>': return c !== 0;
        case '<': return c < 0;
        case '>': return c > 0;
        case '<=': return c <= 0;
        case '>=': return c >= 0;
      }
    }
  }
  throw ERR.VALUE;
}

/** 셀 수식 평가 결과(스칼라 또는 오류 값) */
export function evaluateFormula(ast, ctx) {
  try {
    let v = evaluate(ast, ctx);
    if (v instanceof Range) v = v.height === 1 && v.width === 1 ? v.rows[0][0] : ERR.VALUE;
    if (v === null || v === undefined) return 0;
    if (typeof v === 'number' && !Number.isFinite(v)) return ERR.NUM;
    return v;
  } catch (e) {
    if (e instanceof FormulaError) return e;
    throw e; // RangeError(너무 깊은 재귀)는 통합 문서에서 처리
  }
}

function tryEval(node, ctx) {
  try {
    return scalarOrError(evaluate(node, ctx));
  } catch (e) {
    if (e instanceof FormulaError) return e;
    throw e;
  }
}

function scalarOrError(v) {
  if (v instanceof Range) return v.height === 1 && v.width === 1 ? v.rows[0][0] : ERR.VALUE;
  return v;
}

// ───────────────────────── 함수 ─────────────────────────
/** SUM 계열 규칙: 범위 안에서는 숫자만, 직접 인수는 숫자로 변환 */
function collectNums(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) {
      for (const v of a.values()) {
        if (isError(v)) throw v;
        if (typeof v === 'number') out.push(v);
      }
    } else if (a !== null) {
      out.push(toNum(a));
    }
  }
  return out;
}

function flat(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) out.push(...a.values());
    else out.push(a);
  }
  return out;
}

function asRange(v) {
  if (v instanceof Range) return v;
  return new Range([[v]]);
}

function roundTo(n, digits, mode) {
  const f = 10 ** digits;
  const x = n * f;
  let r;
  const eps = 1e-9;
  if (mode === 'up') r = Math.sign(x) * Math.ceil(Math.abs(x) - eps);
  else if (mode === 'down') r = Math.sign(x) * Math.floor(Math.abs(x) + eps);
  else r = Math.sign(x) * Math.round(Math.abs(x) + eps);
  return r / f;
}

function wildcardRegex(pattern) {
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

function escapeRe(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

/** COUNTIF/SUMIF 조건 → 판별 함수 */
export function makeCriteria(crit) {
  crit = scalar(crit);
  if (typeof crit === 'number') return (v) => typeof v === 'number' && v === crit;
  if (typeof crit === 'boolean') return (v) => v === crit;
  if (crit === null) return (v) => typeof v === 'number' && v === 0;
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

// 날짜: 엑셀 일련번호 (1899-12-30 기준)
const EPOCH = Date.UTC(1899, 11, 30);
const DAY_MS = 86400000;
export const dateToSerial = (y, m, d) => (Date.UTC(y, m - 1, d) - EPOCH) / DAY_MS;
export function serialToDate(serial) {
  const d = new Date(EPOCH + Math.round(serial * DAY_MS));
  return {
    y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(),
    hh: d.getUTCHours(), mm: d.getUTCMinutes(), ss: d.getUTCSeconds(), dow: d.getUTCDay(),
  };
}
function todaySerial() {
  const n = new Date();
  return dateToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate());
}

function lazy(fn) { fn.lazy = true; return fn; }

function lookupExact(values, key) {
  const k = scalar(key);
  if (typeof k === 'string' && /[*?~]/.test(k)) {
    const re = wildcardRegex(k);
    return values.findIndex((v) => typeof v === 'string' && re.test(v));
  }
  return values.findIndex((v) => v !== null && typeRank(v) === typeRank(k) && compareValues(v, k) === 0);
}

function lookupApprox(values, key, descending = false) {
  const k = scalar(key);
  let found = -1;
  for (let i = 0; i < values.length; i++) {
    const v = values[i];
    if (v === null || typeRank(v) !== typeRank(k)) continue;
    const c = compareValues(v, k);
    if (descending ? c >= 0 : c <= 0) found = i;
    else break;
  }
  return found;
}

function ifsPairs(args) {
  if (args.length % 2 !== 0) throw ERR.VALUE;
  const pairs = [];
  for (let i = 0; i < args.length; i += 2) {
    pairs.push({ range: asRange(args[i]), test: makeCriteria(args[i + 1]) });
  }
  return pairs;
}

function matchesAll(pairs, r, c) {
  return pairs.every(({ range, test }) => {
    const row = range.rows[r];
    return row !== undefined && c < row.length && test(row[c]);
  });
}

export const FUNCS = {
  // 수학
  SUM: (a) => collectNums(a).reduce((s, x) => s + x, 0),
  PRODUCT: (a) => collectNums(a).reduce((s, x) => s * x, 1),
  AVERAGE: (a) => {
    const n = collectNums(a);
    if (!n.length) throw ERR.DIV0;
    return n.reduce((s, x) => s + x, 0) / n.length;
  },
  MIN: (a) => { const n = collectNums(a); return n.length ? Math.min(...n) : 0; },
  MAX: (a) => { const n = collectNums(a); return n.length ? Math.max(...n) : 0; },
  MEDIAN: (a) => {
    const n = collectNums(a).sort((x, y) => x - y);
    if (!n.length) throw ERR.NUM;
    const m = n.length >> 1;
    return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
  },
  COUNT: (a) => {
    let k = 0;
    for (const x of a) {
      if (x instanceof Range) { for (const v of x.values()) if (typeof v === 'number') k++; }
      else if (typeof x === 'number' || typeof x === 'boolean' || (typeof x === 'string' && parseNumberText(x) !== null)) k++;
    }
    return k;
  },
  COUNTA: (a) => flat(a).filter((v) => v !== null).length,
  COUNTBLANK: (a) => flat(a).filter((v) => v === null || v === '').length,
  ROUND: ([n, d]) => roundTo(toNum(n), toInt(d ?? 0)),
  ROUNDUP: ([n, d]) => roundTo(toNum(n), toInt(d ?? 0), 'up'),
  ROUNDDOWN: ([n, d]) => roundTo(toNum(n), toInt(d ?? 0), 'down'),
  TRUNC: ([n, d]) => roundTo(toNum(n), toInt(d ?? 0), 'down'),
  INT: ([n]) => Math.floor(toNum(n)),
  ABS: ([n]) => Math.abs(toNum(n)),
  SIGN: ([n]) => Math.sign(toNum(n)),
  MOD: ([n, d]) => {
    const x = toNum(n);
    const y = toNum(d);
    if (y === 0) throw ERR.DIV0;
    return x - y * Math.floor(x / y);
  },
  POWER: ([a, b]) => binary('^', a, b),
  SQRT: ([n]) => { const x = toNum(n); if (x < 0) throw ERR.NUM; return Math.sqrt(x); },
  EXP: ([n]) => checkNum(Math.exp(toNum(n))),
  LN: ([n]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log(x); },
  LOG10: ([n]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log10(x); },
  LOG: ([n, b]) => {
    const x = toNum(n);
    const base = b == null ? 10 : toNum(b);
    if (x <= 0 || base <= 0 || base === 1) throw ERR.NUM;
    return Math.log(x) / Math.log(base);
  },
  PI: () => Math.PI,
  RAND: () => Math.random(),
  RANDBETWEEN: ([lo, hi]) => {
    const a = Math.ceil(toNum(lo));
    const b = Math.floor(toNum(hi));
    if (a > b) throw ERR.NUM;
    return a + Math.floor(Math.random() * (b - a + 1));
  },
  SUMPRODUCT: (a) => {
    const ranges = a.map(asRange);
    const h = ranges[0].height;
    const w = ranges[0].width;
    if (ranges.some((r) => r.height !== h || r.width !== w)) throw ERR.VALUE;
    let s = 0;
    for (let r = 0; r < h; r++) {
      for (let c = 0; c < w; c++) {
        let p = 1;
        for (const rg of ranges) {
          const v = rg.rows[r][c];
          if (isError(v)) throw v;
          p *= typeof v === 'number' ? v : 0;
        }
        s += p;
      }
    }
    return s;
  },

  // 논리
  IF: lazy((args, ctx) => {
    const cond = toBool(evaluate(args[0], ctx));
    const branch = cond ? args[1] : args[2];
    if (!branch) return cond;
    return scalarOrError(evaluate(branch, ctx)) ?? 0;
  }),
  IFERROR: lazy((args, ctx) => {
    const v = tryEval(args[0], ctx);
    if (isError(v)) return args[1] ? scalar(evaluate(args[1], ctx)) ?? 0 : 0;
    return v;
  }),
  IFNA: lazy((args, ctx) => {
    const v = tryEval(args[0], ctx);
    if (v === ERR.NA) return args[1] ? scalar(evaluate(args[1], ctx)) ?? 0 : 0;
    if (isError(v)) throw v;
    return v;
  }),
  IFS: lazy((args, ctx) => {
    for (let i = 0; i + 1 < args.length; i += 2) {
      if (toBool(evaluate(args[i], ctx))) return scalar(evaluate(args[i + 1], ctx)) ?? 0;
    }
    throw ERR.NA;
  }),
  AND: (a) => {
    const vals = flat(a).filter((v) => v !== null && typeof v !== 'string');
    if (!vals.length) throw ERR.VALUE;
    return vals.every((v) => toBool(v));
  },
  OR: (a) => {
    const vals = flat(a).filter((v) => v !== null && typeof v !== 'string');
    if (!vals.length) throw ERR.VALUE;
    return vals.some((v) => toBool(v));
  },
  NOT: ([v]) => !toBool(v),
  TRUE: () => true,
  FALSE: () => false,
  CHOOSE: ([i, ...opts]) => {
    const k = toInt(i);
    if (k < 1 || k > opts.length) throw ERR.VALUE;
    return scalar(opts[k - 1]);
  },

  // 정보
  ISBLANK: lazy((args, ctx) => tryEval(args[0], ctx) === null),
  ISNUMBER: lazy((args, ctx) => typeof tryEval(args[0], ctx) === 'number'),
  ISTEXT: lazy((args, ctx) => typeof tryEval(args[0], ctx) === 'string'),
  ISLOGICAL: lazy((args, ctx) => typeof tryEval(args[0], ctx) === 'boolean'),
  ISERROR: lazy((args, ctx) => isError(tryEval(args[0], ctx))),
  ISERR: lazy((args, ctx) => { const v = tryEval(args[0], ctx); return isError(v) && v !== ERR.NA; }),
  ISNA: lazy((args, ctx) => tryEval(args[0], ctx) === ERR.NA),
  NA: () => { throw ERR.NA; },

  // 텍스트
  CONCATENATE: (a) => a.map(toStr).join(''),
  CONCAT: (a) => flat(a).map((v) => { if (isError(v)) throw v; return toStr(v); }).join(''),
  TEXTJOIN: ([delim, ignoreEmpty, ...rest]) => {
    const d = toStr(delim);
    const skip = toBool(ignoreEmpty);
    return flat(rest)
      .map((v) => { if (isError(v)) throw v; return toStr(v); })
      .filter((s) => !(skip && s === ''))
      .join(d);
  },
  LEN: ([s]) => [...toStr(s)].length,
  LEFT: ([s, n]) => [...toStr(s)].slice(0, n == null ? 1 : Math.max(0, toInt(n))).join(''),
  RIGHT: ([s, n]) => {
    const chars = [...toStr(s)];
    const k = n == null ? 1 : Math.max(0, toInt(n));
    return k === 0 ? '' : chars.slice(-k).join('');
  },
  MID: ([s, start, n]) => {
    const st = toInt(start);
    const k = toInt(n);
    if (st < 1 || k < 0) throw ERR.VALUE;
    return [...toStr(s)].slice(st - 1, st - 1 + k).join('');
  },
  UPPER: ([s]) => toStr(s).toUpperCase(),
  LOWER: ([s]) => toStr(s).toLowerCase(),
  PROPER: ([s]) => toStr(s).toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (_, p, ch) => p + ch.toUpperCase()),
  TRIM: ([s]) => toStr(s).trim().replace(/ {2,}/g, ' '),
  REPT: ([s, n]) => { const k = toInt(n); if (k < 0) throw ERR.VALUE; return toStr(s).repeat(k); },
  SUBSTITUTE: ([s, oldT, newT, inst]) => {
    const text = toStr(s);
    const o = toStr(oldT);
    const nw = toStr(newT);
    if (o === '') return text;
    if (inst == null) return text.split(o).join(nw);
    const k = toInt(inst);
    let idx = -1;
    for (let i = 0; i < k; i++) {
      idx = text.indexOf(o, idx + 1);
      if (idx < 0) return text;
    }
    return text.slice(0, idx) + nw + text.slice(idx + o.length);
  },
  FIND: ([f, s, start]) => {
    const st = start == null ? 1 : toInt(start);
    const idx = toStr(s).indexOf(toStr(f), st - 1);
    if (st < 1 || idx < 0) throw ERR.VALUE;
    return idx + 1;
  },
  SEARCH: ([f, s, start]) => {
    const st = start == null ? 1 : toInt(start);
    const text = toStr(s);
    const re = wildcardRegex(toStr(f));
    for (let i = st - 1; i < text.length; i++) {
      for (let j = text.length; j > i; j--) if (re.test(text.slice(i, j))) return i + 1;
    }
    throw ERR.VALUE;
  },
  EXACT: ([a, b]) => toStr(a) === toStr(b),
  VALUE: ([s]) => toNum(s),
  TEXT: ([v, fmt]) => {
    const x = scalar(v);
    if (typeof x === 'string' && parseNumberText(x) === null) return x;
    return formatWithPattern(toNum(x), toStr(fmt));
  },

  // 찾기/참조
  VLOOKUP: ([key, table, col, approx]) => {
    const t = asRange(table);
    const ci = toInt(col);
    if (ci < 1) throw ERR.VALUE;
    if (ci > t.width) throw ERR.REF;
    const first = t.rows.map((row) => row[0]);
    const exact = approx != null && !toBool(approx);
    const i = exact ? lookupExact(first, key) : lookupApprox(first, key);
    if (i < 0) throw ERR.NA;
    return t.rows[i][ci - 1] ?? 0;
  },
  HLOOKUP: ([key, table, row, approx]) => {
    const t = asRange(table);
    const ri = toInt(row);
    if (ri < 1) throw ERR.VALUE;
    if (ri > t.height) throw ERR.REF;
    const exact = approx != null && !toBool(approx);
    const i = exact ? lookupExact(t.rows[0], key) : lookupApprox(t.rows[0], key);
    if (i < 0) throw ERR.NA;
    return t.rows[ri - 1][i] ?? 0;
  },
  XLOOKUP: ([key, look, ret, notFound]) => {
    const l = [...asRange(look).values()];
    const rr = [...asRange(ret).values()];
    const i = lookupExact(l, key);
    if (i < 0) {
      if (notFound !== undefined && notFound !== null) return scalar(notFound);
      throw ERR.NA;
    }
    return rr[i] ?? 0;
  },
  INDEX: ([table, row, col]) => {
    const t = asRange(table);
    let r = toInt(row ?? 1);
    let c = col == null ? 1 : toInt(col);
    if (t.height === 1 && col == null) { c = r; r = 1; }
    if (r < 1 || c < 1 || r > t.height || c > t.width) throw ERR.REF;
    return t.rows[r - 1][c - 1] ?? 0;
  },
  MATCH: ([key, look, type]) => {
    const vals = [...asRange(look).values()];
    const mt = type == null ? 1 : toInt(type);
    const i = mt === 0 ? lookupExact(vals, key) : lookupApprox(vals, key, mt < 0);
    if (i < 0) throw ERR.NA;
    return i + 1;
  },
  ROWS: ([r]) => asRange(r).height,
  COLUMNS: ([r]) => asRange(r).width,

  // 조건부 집계
  SUMIF: ([range, crit, sumRange]) => {
    const r = asRange(range);
    const s = sumRange == null ? r : asRange(sumRange);
    const test = makeCriteria(crit);
    let total = 0;
    r.rows.forEach((row, i) => row.forEach((v, j) => {
      if (!test(v)) return;
      const x = s.rows[i]?.[j];
      if (isError(x)) throw x;
      if (typeof x === 'number') total += x;
    }));
    return total;
  },
  COUNTIF: ([range, crit]) => {
    const test = makeCriteria(crit);
    let k = 0;
    for (const v of asRange(range).values()) if (test(v)) k++;
    return k;
  },
  AVERAGEIF: ([range, crit, avgRange]) => {
    const r = asRange(range);
    const s = avgRange == null ? r : asRange(avgRange);
    const test = makeCriteria(crit);
    let total = 0;
    let k = 0;
    r.rows.forEach((row, i) => row.forEach((v, j) => {
      const x = s.rows[i]?.[j];
      if (test(v) && typeof x === 'number') { total += x; k++; }
    }));
    if (!k) throw ERR.DIV0;
    return total / k;
  },
  SUMIFS: ([sumRange, ...rest]) => {
    const s = asRange(sumRange);
    const pairs = ifsPairs(rest);
    let total = 0;
    s.rows.forEach((row, i) => row.forEach((x, j) => {
      if (typeof x === 'number' && matchesAll(pairs, i, j)) total += x;
    }));
    return total;
  },
  COUNTIFS: (args) => {
    const pairs = ifsPairs(args);
    const base = pairs[0].range;
    let k = 0;
    base.rows.forEach((row, i) => row.forEach((_, j) => { if (matchesAll(pairs, i, j)) k++; }));
    return k;
  },

  // 날짜
  TODAY: () => todaySerial(),
  NOW: () => {
    const n = new Date();
    return todaySerial() + (n.getHours() * 3600 + n.getMinutes() * 60 + n.getSeconds()) / 86400;
  },
  DATE: ([y, m, d]) => dateToSerial(toInt(y), toInt(m), toInt(d)),
  YEAR: ([s]) => serialToDate(toNum(s)).y,
  MONTH: ([s]) => serialToDate(toNum(s)).m,
  DAY: ([s]) => serialToDate(toNum(s)).d,
  HOUR: ([s]) => serialToDate(toNum(s)).hh,
  MINUTE: ([s]) => serialToDate(toNum(s)).mm,
  SECOND: ([s]) => serialToDate(toNum(s)).ss,
  WEEKDAY: ([s, type]) => {
    const dow = serialToDate(toNum(s)).dow;
    const t = type == null ? 1 : toInt(type);
    if (t === 2) return ((dow + 6) % 7) + 1;
    if (t === 3) return (dow + 6) % 7;
    return dow + 1;
  },
};

export const FUNCTION_NAMES = Object.keys(FUNCS).sort();

/** 결과에 날짜 서식을 자동 적용해야 하는 함수 */
export function autoFormatFor(ast) {
  if (ast?.type !== 'func') return null;
  if (ast.name === 'NOW') return 'datetime';
  if (ast.name === 'TODAY' || ast.name === 'DATE') return 'date';
  return null;
}

// ───────────────────────── 참조 재작성 ─────────────────────────
function refText(ref, sheetPrefix) {
  const cell = (r, c, ar, ac) => `${ac ? '$' : ''}${colToName(c)}${ar ? '$' : ''}${r + 1}`;
  if (ref.cols) {
    return `${sheetPrefix}${ref.ac1 ? '$' : ''}${colToName(ref.c1)}:${ref.ac2 ? '$' : ''}${colToName(ref.c2)}`;
  }
  const a = cell(ref.r1, ref.c1, ref.ar1, ref.ac1);
  if (!ref.range) return sheetPrefix + a;
  return `${sheetPrefix}${a}:${cell(ref.r2, ref.c2, ref.ar2, ref.ac2)}`;
}

/**
 * 수식 문자열 안의 각 참조에 대해 fn(ref)을 호출.
 * fn은 새 ref 객체(재작성), null(#REF!), undefined(변경 없음)를 반환.
 */
export function rewriteRefs(formula, fn) {
  if (!formula.startsWith('=')) return formula;
  const src = formula.slice(1);
  let toks;
  try { toks = tokenize(src); } catch { return formula; }
  let out = '';
  let pos = 0;
  for (const t of toks) {
    if (t.t !== 'ref') continue;
    const res = fn({ ...t.ref });
    if (res === undefined) continue;
    const text = src.slice(t.s, t.e);
    const bang = text.lastIndexOf('!');
    const prefix = res?.sheetPrefix ?? (bang >= 0 && t.ref.sheet != null ? text.slice(0, bang + 1) : '');
    out += src.slice(pos, t.s) + (res === null ? '#REF!' : refText(res, prefix));
    pos = t.e;
  }
  return '=' + out + src.slice(pos);
}

/** 수식 복사 시 상대 참조 이동 */
export function shiftFormula(formula, dr, dc) {
  if (!dr && !dc) return formula;
  return rewriteRefs(formula, (ref) => {
    if (!ref.cols) {
      if (!ref.ar1) ref.r1 += dr;
      if (!ref.ar2) ref.r2 += dr;
    }
    if (!ref.ac1) ref.c1 += dc;
    if (!ref.ac2) ref.c2 += dc;
    if (ref.r1 < 0 || ref.r2 < 0 || ref.c1 < 0 || ref.c2 < 0) return null;
    if (ref.r1 >= MAX_ROWS || ref.r2 >= MAX_ROWS || ref.c1 >= MAX_COLS || ref.c2 >= MAX_COLS) return null;
    return ref;
  });
}

/**
 * 행/열 삽입(count>0)·삭제(count<0) 시 참조 조정.
 * targetSheet: 구조가 바뀐 시트 이름, hostSheet: 수식이 있는 시트 이름
 */
export function adjustFormulaForStructure(formula, { targetSheet, hostSheet, axis, index, count }) {
  const same = (a, b) => a.toLowerCase() === b.toLowerCase();
  return rewriteRefs(formula, (ref) => {
    if (!same(ref.sheet ?? hostSheet, targetSheet)) return undefined;
    if (axis === 'row' && ref.cols) return undefined;
    const k1 = axis === 'row' ? 'r1' : 'c1';
    const k2 = axis === 'row' ? 'r2' : 'c2';
    let a = ref[k1];
    let b = ref[k2];
    if (count > 0) {
      if (a >= index) a += count;
      if (b >= index) b += count;
    } else {
      const n = -count;
      const end = index + n; // [index, end) 삭제
      const map1 = (x) => (x >= end ? x - n : x >= index ? index : x);
      const map2 = (x) => (x >= end ? x - n : x >= index ? index - 1 : x);
      if (!ref.range && a >= index && a < end) return null;
      if (ref.range && a >= index && b < end) return null;
      a = map1(a);
      b = map2(b);
      if (b < a) return null;
    }
    if (a === ref[k1] && b === ref[k2]) return undefined;
    ref[k1] = a;
    ref[k2] = b;
    return ref;
  });
}

/** 시트 이름 변경 시 참조 갱신 */
export function renameSheetInFormula(formula, oldName, newName) {
  return rewriteRefs(formula, (ref) => {
    if (ref.sheet == null || ref.sheet.toLowerCase() !== oldName.toLowerCase()) return undefined;
    ref.sheetPrefix = quoteSheetName(newName) + '!';
    return ref;
  });
}

/** 입력한 수식 정리: 닫는 괄호/따옴표 자동 추가, 함수·참조 대문자 변환 ("=sum(a1" → "=SUM(A1)") */
export function normalizeFormula(text) {
  let depth = 0;
  let inStr = false;
  for (const ch of text) {
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === '(') depth++;
    else if (!inStr && ch === ')') depth--;
  }
  let out = text;
  if (inStr) out += '"';
  if (depth > 0) out += ')'.repeat(depth);
  const body = out.slice(1);
  let toks;
  try { toks = tokenize(body); } catch { return out; }
  let res = '';
  let pos = 0;
  for (const t of toks) {
    if (t.t !== 'func' && t.t !== 'ref' && t.t !== 'bool') continue;
    let s = body.slice(t.s, t.e);
    if (t.t === 'ref') {
      const bang = s.lastIndexOf('!');
      s = s.slice(0, bang + 1) + s.slice(bang + 1).toUpperCase();
    } else {
      s = s.toUpperCase();
    }
    res += body.slice(pos, t.s) + s;
    pos = t.e;
  }
  return `=${res}${body.slice(pos)}`;
}

/** 수식 편집 중 참조 목록 (색상 강조용) */
export function listRefs(formula) {
  if (!formula.startsWith('=')) return [];
  try {
    return tokenize(formula.slice(1)).filter((t) => t.t === 'ref').map((t) => ({ ...t.ref, s: t.s + 1, e: t.e + 1 }));
  } catch {
    return [];
  }
}
