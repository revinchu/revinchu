// WIXEL 수식 엔진: 토크나이저 · 파서 · 평가기(배열·분산·이름·LET/LAMBDA) · 참조 재작성
import {
  FormulaError as FormulaErrorCore, ERR as ERR_CORE, ERR_BY_CODE, isError as isErrorCore, Range as RangeCore, RefValue as RefValueCore,
  Lambda, scalar, toNum, toStr, checkNum, compareValues as compareCore, makeCriteria as makeCriteriaCore,
  parseNumberText as parseNumberCore, dateToSerial as dateToSerialCore, serialToDate as serialToDateCore,
  lazy, attempt, broadcast2, mapValue,
} from './fxcore.js';
import { MATH } from './fx-math.js';
import { STAT } from './fx-stat.js';
import { TEXT } from './fx-text.js';
import { DATE } from './fx-date.js';
import { LOOKUP } from './fx-lookup.js';
import { LOGIC } from './fx-logic.js';
import { FIN } from './fx-fin.js';
import { WEB, WEB_ARRAY } from './fx-web.js';

// 공통 값/오류 (다른 모듈 호환용 재수출)
export const FormulaError = FormulaErrorCore;
export const ERR = ERR_CORE;
export const isError = isErrorCore;
export const Range = RangeCore;
export const RefValue = RefValueCore;
export const compareValues = compareCore;
export const makeCriteria = makeCriteriaCore;
export const parseNumberText = parseNumberCore;
export const dateToSerial = dateToSerialCore;
export const serialToDate = serialToDateCore;

// ───────────────────────── 주소 유틸 ─────────────────────────
// 엑셀(1,048,576행)보다 많은 1,000만 행 지원. .xlsx 로 저장할 때만 엑셀 한도가 적용됨
export const MAX_ROWS = 20_000_000;
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
  return /^(?:\[\d+\])?[A-Za-z_À-￿][\w.À-￿]*$/.test(name) && !parseCellName(name) && !/^R\d*C\d*$/i.test(name)
    ? name
    : `'${name.replace(/'/g, "''")}'`;
}

// ───────────────────────── 토크나이저 ─────────────────────────
// 시트 이름: '따옴표 이름' 또는 일반 이름 (외부 통합 문서 참조는 [1]시트 처럼 앞에 [번호])
const SHEET = String.raw`(?:'((?:[^']|'')+)'|((?:\[\d+\])?[A-Za-z_À-￿][\w.À-￿]*))!`;
const CELL = String.raw`(\$?)([A-Za-z]{1,3})(\$?)(\d+)`;
const RANGE_RE = new RegExp(`(?:${SHEET})?${CELL}(?::${CELL})?(?![\\w(![])`, 'y');
const COLS_RE = new RegExp(`(?:${SHEET})?(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})(?![\\w(!])`, 'y');
const ROWS_RE = new RegExp(`(?:${SHEET})?(\\$?)(\\d+):(\\$?)(\\d+)(?![\\w(!.])`, 'y');
const SHEET_NAME_RE = new RegExp(`${SHEET}([A-Za-z_\\\\À-￿][\\w.À-￿]*)`, 'y');
const EXT_BOOK_RE = /^\[\d+\][^\]]/; // 외부 통합 문서 참조 [1]시트!A1
const NUMBER_RE = /(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/y;
const IDENT_RE = /[A-Za-z_\\À-￿][\w.?À-￿]*/y;
const SHEET_REF_ERR_RE = /(?:'(?:[^']|'')+'|[^\s!'"(),;:+\-*/^&=<>{}#]+)!#REF!/y;
const ERROR_RE = /#(?:DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|NULL!|CIRC!|SPILL!|CALC!|BUSY!|GETTING_DATA)/y;
const OPS = ['<=', '>=', '<>', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%', '@'];

function sticky(re, src, pos) {
  re.lastIndex = pos;
  return re.exec(src);
}

/** 여는 대괄호 위치 → 짝이 맞는 닫는 대괄호 위치 (' 는 이스케이프) */
function bracketEnd(src, open) {
  let depth = 0;
  for (let j = open; j < src.length; j++) {
    const ch = src[j];
    if (ch === "'") { j++; continue; }
    if (ch === '[') depth++;
    else if (ch === ']' && --depth === 0) return j;
  }
  throw new SyntaxError('닫는 대괄호가 없습니다');
}

/** 함수 이름의 파일 접두사 제거 (_xlfn. / _xlws.) */
export const stripFnPrefix = (name) => name.replace(/^(?:_XLFN\.|_XLWS\.)+/i, '').replace(/^_XLFN\._XLWS\./i, '');

export function tokenize(src) {
  const toks = [];
  let i = 0;
  const prevIsValue = () => {
    const p = toks[toks.length - 1];
    return p && (p.t === 'ref' || p.t === 'ident' || p.t === ')' || p.t === 'sref');
  };
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
      if (m) {
        i += m[0].length;
        toks.push({ t: 'err', v: m[0] === '#GETTING_DATA' ? '#BUSY!' : m[0], s: start, e: i });
        continue;
      }
      if (prevIsValue()) {
        // 분산 범위 참조 A1#
        i++;
        toks.push({ t: '#', s: start, e: i });
        continue;
      }
      throw new SyntaxError('알 수 없는 오류 값');
    }
    if (/\d/.test(ch)) {
      const rm = sticky(ROWS_RE, src, i);
      if (rm && validRows(rm)) {
        i += rm[0].length;
        toks.push({ t: 'ref', ref: refFromRowsMatch(rm), s: start, e: i });
        continue;
      }
    }
    if (/[\d.]/.test(ch)) {
      const m = sticky(NUMBER_RE, src, i);
      if (!m) throw new SyntaxError('잘못된 숫자');
      i += m[0].length;
      toks.push({ t: 'num', v: parseFloat(m[0]), s: start, e: i });
      continue;
    }
    if (/[A-Za-z_$'\\À-￿]/.test(ch) || (ch === '[' && EXT_BOOK_RE.test(src.slice(i, i + 8)))) {
      // 지워진 범위를 가리키는 시트 참조 (시트!#REF!): 엑셀은 #REF! 로 계산
      const dead = sticky(SHEET_REF_ERR_RE, src, i);
      if (dead) {
        i += dead[0].length;
        toks.push({ t: 'err', v: '#REF!', s: start, e: i });
        continue;
      }
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
      m = sticky(ROWS_RE, src, i);
      if (m && validRows(m)) {
        i += m[0].length;
        toks.push({ t: 'ref', ref: refFromRowsMatch(m), s: start, e: i });
        continue;
      }
      m = sticky(SHEET_NAME_RE, src, i);
      if (m) {
        // 시트 범위 이름: Sheet1!이름
        i += m[0].length;
        toks.push({ t: 'ident', v: m[3], sheet: sheetFromMatch(m[1], m[2]), s: start, e: i });
        continue;
      }
      m = sticky(IDENT_RE, src, i);
      if (m && src[i + m[0].length] === '[') {
        // 구조적 참조: 표1[금액], 표1[[#머리글],[금액]]
        const end = bracketEnd(src, i + m[0].length);
        toks.push({ t: 'sref', table: m[0], spec: src.slice(i + m[0].length + 1, end), s: start, e: end + 1 });
        i = end + 1;
        continue;
      }
      if (m) {
        i += m[0].length;
        const up = m[0].toUpperCase();
        let j = i;
        while (j < src.length && /\s/.test(src[j])) j++;
        if (src[j] === '(') toks.push({ t: 'func', v: stripFnPrefix(up), s: start, e: i });
        else if (up === 'TRUE' || up === 'FALSE') toks.push({ t: 'bool', v: up === 'TRUE', s: start, e: i });
        else toks.push({ t: 'ident', v: m[0].replace(/^_xlpm\./i, '').replace(/^_xleta\./i, ''), s: start, e: i });
        continue;
      }
      throw new SyntaxError(`잘못된 문자: ${ch}`);
    }
    if (ch === '[') {
      // 표 안의 구조적 참조: [@금액], [금액]
      const end = bracketEnd(src, i);
      toks.push({ t: 'sref', table: null, spec: src.slice(i + 1, end), s: start, e: end + 1 });
      i = end + 1;
      continue;
    }
    if ('(),:{};'.includes(ch)) {
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

function validRows(m) {
  const a = +m[4];
  const b = +m[6];
  return a >= 1 && b >= 1 && a <= MAX_ROWS && b <= MAX_ROWS;
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

function refFromRowsMatch(m) {
  return {
    sheet: sheetFromMatch(m[1], m[2]),
    c1: 0, c2: MAX_COLS - 1, ac1: true, ac2: true,
    r1: +m[4] - 1, r2: +m[6] - 1, ar1: !!m[3], ar2: !!m[5],
    range: true, cols: false, rows: true,
  };
}

// ───────────────────────── 파서 ─────────────────────────
const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

/** 수식 본문(맨 앞 '=' 제외)을 AST로 변환. 노드에는 원문 위치 s/e 가 있음 */
export function parse(src) {
  const toks = tokenize(src);
  let i = 0;
  const peek = () => toks[i];
  const next = () => toks[i++];
  const fail = (msg) => { throw new SyntaxError(msg); };
  const endOf = () => toks[i - 1]?.e ?? src.length;

  function expr(minPrec) {
    let left = unary();
    for (;;) {
      const k = peek();
      if (!k || k.t !== 'op') break;
      if (k.v === '%') { next(); left = { type: 'pct', a: left, s: left.s, e: k.e }; continue; }
      const p = BIN_PREC[k.v];
      if (p === undefined || p < minPrec) break;
      next();
      // ^ 는 왼쪽 결합 (엑셀과 같음)
      const b = expr(p + 1);
      left = { type: 'bin', op: k.v, a: left, b, s: left.s, e: b.e };
    }
    return left;
  }

  function unary() {
    const k = peek();
    if (k && k.t === 'op' && (k.v === '-' || k.v === '+')) {
      next();
      const a = unary();
      return { type: k.v === '-' ? 'neg' : 'pos', a, s: k.s, e: a.e };
    }
    if (k && k.t === 'op' && k.v === '@') {
      next();
      const a = unary();
      return { type: 'at', a, s: k.s, e: a.e };
    }
    return postfix();
  }

  function postfix() {
    let node = primary();
    for (;;) {
      const k = peek();
      if (!k) break;
      if (k.t === '#') { next(); node = { type: 'spill', a: node, s: node.s, e: k.e }; continue; }
      if (k.t === ':') {
        next();
        const b = primaryWithSpill();
        node = { type: 'range', a: node, b, s: node.s, e: b.e };
        continue;
      }
      if (k.t === '(' && (node.type === 'func' || node.type === 'call' || node.type === 'paren')) {
        next();
        const args = argList();
        node = { type: 'call', fn: node, args, s: node.s, e: endOf() };
        continue;
      }
      break;
    }
    return node;
  }

  function primaryWithSpill() {
    let node = primary();
    if (peek()?.t === '#') { const k = next(); node = { type: 'spill', a: node, s: node.s, e: k.e }; }
    return node;
  }

  function argList() {
    const args = [];
    if (peek()?.t === ')') { next(); return args; }
    for (;;) {
      const t = peek();
      if (t && (t.t === ',' || t.t === ')')) args.push({ type: 'empty', s: t.s, e: t.s });
      else args.push(expr(0));
      const sep = next();
      if (!sep) fail('닫는 괄호가 없습니다');
      if (sep.t === ')') break;
      if (sep.t !== ',') fail('인수 구분 기호가 잘못되었습니다');
    }
    return args;
  }

  function arrayConst(open) {
    const rows = [[]];
    for (;;) {
      let k = next();
      if (!k) fail('닫는 중괄호가 없습니다');
      let sign = 1;
      while (k.t === 'op' && (k.v === '-' || k.v === '+')) {
        if (k.v === '-') sign = -sign;
        k = next();
      }
      let v;
      if (k.t === 'num') v = sign * k.v;
      else if (k.t === 'str') v = k.v;
      else if (k.t === 'bool') v = k.v;
      else if (k.t === 'err') v = ERR_BY_CODE[k.v] ?? ERR.VALUE;
      else fail('배열 상수에는 값만 넣을 수 있습니다');
      rows[rows.length - 1].push(v);
      const sep = next();
      if (!sep) fail('닫는 중괄호가 없습니다');
      if (sep.t === '}') break;
      if (sep.t === ';') rows.push([]);
      else if (sep.t !== ',') fail('배열 상수 구분 기호가 잘못되었습니다');
    }
    const w = rows[0].length;
    if (rows.some((r) => r.length !== w)) fail('배열 상수의 행 길이가 다릅니다');
    return { type: 'array', rows, s: open.s, e: endOf() };
  }

  function primary() {
    const k = next();
    if (!k) fail('수식이 완전하지 않습니다');
    switch (k.t) {
      case 'num': return { type: 'num', v: k.v, s: k.s, e: k.e };
      case 'str': return { type: 'str', v: k.v, s: k.s, e: k.e };
      case 'bool': return { type: 'bool', v: k.v, s: k.s, e: k.e };
      case 'err': return { type: 'err', v: k.v, s: k.s, e: k.e };
      case 'ref': return { type: 'ref', ref: k.ref, s: k.s, e: k.e };
      case 'sref': return { type: 'sref', table: k.table, spec: k.spec, s: k.s, e: k.e };
      case 'ident': return { type: 'name', v: k.v, sheet: k.sheet ?? null, s: k.s, e: k.e };
      case '{': return arrayConst(k);
      case '(': {
        const e = expr(0);
        if (peek()?.t === ',') {
          // 참조 합집합 (A1:B2,C3)
          const items = [e];
          while (peek()?.t === ',') { next(); items.push(expr(0)); }
          if (next()?.t !== ')') fail('닫는 괄호가 없습니다');
          return { type: 'union', items, s: k.s, e: endOf() };
        }
        if (next()?.t !== ')') fail('닫는 괄호가 없습니다');
        return { type: 'paren', a: e, s: k.s, e: endOf() };
      }
      case 'func': {
        if (next()?.t !== '(') fail('여는 괄호가 없습니다');
        const args = argList();
        return { type: 'func', name: k.v, args, s: k.s, e: endOf() };
      }
      default:
        return fail('수식에 문제가 있습니다');
    }
  }

  const ast = expr(0);
  if (i < toks.length) fail('수식에 문제가 있습니다');
  return ast;
}

// ───────────────────────── 평가기 ─────────────────────────
/**
 * ctx: {
 *   cell(sheet, r, c) → 값 (null = 빈 셀), 오류 값일 수 있음
 *   range(sheet, r1, c1, r2, c2) → 2차원 값 배열
 *   usedRows(sheet) / usedCols(sheet) → 사용된 행/열 수 (전체 열/행 참조용)
 *   선택: here {si, sheet, r, c}, structRef, rowHidden, name(이름, 시트), spillRef, formulaText, sheetIndex, sheetCount …
 * }
 */
const OMITTED = Object.freeze({ omitted: true });
let lambdaDepth = 0;

/** 참조 → 값 (여러 셀이면 Range, 한 셀이면 스칼라. 한 셀 오류는 던짐) */
function deref(v, ctx) {
  if (v instanceof RefValue) {
    if (v.single) {
      const x = ctx.cell(v.sheet, v.r1, v.c1);
      if (isError(x)) throw x;
      return x;
    }
    let { r2, c2 } = v;
    // 전체 열/행 참조는 사용 중인 영역까지만
    if (r2 - v.r1 >= 65535 && ctx.usedRows) r2 = Math.max(v.r1, Math.min(r2, ctx.usedRows(v.sheet) - 1));
    if (c2 - v.c1 >= 1023 && ctx.usedCols) c2 = Math.max(v.c1, Math.min(c2, ctx.usedCols(v.sheet) - 1));
    return new Range(ctx.range(v.sheet, v.r1, v.c1, r2, c2), v);
  }
  if (Array.isArray(v)) {
    if (v.length === 1) return deref(v[0], ctx);
    throw ERR.VALUE;
  }
  if (v === OMITTED) return null;
  if (isError(v)) throw v;
  return v;
}

/** deref 와 같지만 오류 값을 던지지 않고 돌려줌 (예외는 브라우저에서 비쌈: 인수 · IFERROR 평가에 씀) */
function derefSoft(v, ctx) {
  if (isError(v)) return v;
  if (v instanceof RefValue && v.single) return ctx.cell(v.sheet, v.r1, v.c1);
  return deref(v, ctx);
}

function makeEv(ctx) {
  return {
    deref: (v) => (isError(v) ? v : attempt(() => derefSoft(v, ctx))),
    evaluate,
    evalRef: evalAny,
    value: (node, c) => attempt(() => derefSoft(evalAny(node, c), c)),
    call: callLambda,
    refFromText: (text, a1, c) => refFromText(text, a1, c ?? ctx),
  };
}
const evOf = (ctx) => {
  if (!Object.prototype.hasOwnProperty.call(ctx, '__ev') && !ctx.__ev) {
    Object.defineProperty(ctx, '__ev', { value: makeEv(ctx), enumerable: false });
  }
  return ctx.__ev;
};

export function evaluate(node, ctx) {
  return deref(evalAny(node, ctx), ctx);
}

/** 참조를 값으로 바꾸지 않고 평가 (RefValue · 값 · 참조 합집합 배열) */
export function evalAny(node, ctx) {
  switch (node.type) {
    case 'num':
    case 'str':
    case 'bool':
      return node.v;
    case 'err':
      throw ERR_BY_CODE[node.v] ?? ERR.VALUE;
    case 'empty':
      return null;
    case 'array':
      return new Range(node.rows);
    case 'paren':
      return evalAny(node.a, ctx);
    case 'ref': {
      const f = node.ref;
      // 공유 수식: 같은 모양의 수식이 AST 하나를 쓰고, 셀 위치만큼 상대 참조를 옮김 (ctx.dr · ctx.dc)
      const dr = ctx.dr;
      const dc = ctx.dc;
      if (dr || dc) {
        // 열 전체(A:A)는 행을, 행 전체(1:1)는 열을 옮기지 않음
        return new RefValue(f.sheet, f.ar1 || f.cols ? f.r1 : f.r1 + dr, f.ac1 || f.rows ? f.c1 : f.c1 + dc,
          f.ar2 || f.cols ? f.r2 : f.r2 + dr, f.ac2 || f.rows ? f.c2 : f.c2 + dc);
      }
      return new RefValue(f.sheet, f.r1, f.c1, f.r2, f.c2);
    }
    case 'sref': {
      const ref = ctx.structRef?.(node.table, node.spec);
      if (!ref) throw ERR.REF;
      if (ref.error) throw ref.error;
      return new RefValue(ref.sheet, ref.r1, ref.c1, ref.r2, ref.c2);
    }
    case 'spill': {
      const a = evalAny(node.a, ctx);
      if (!(a instanceof RefValue) || !ctx.spillRef) throw ERR.REF;
      const res = ctx.spillRef(a.sheet, a.r1, a.c1);
      if (!res) throw ERR.REF;
      return res;
    }
    case 'name': {
      const key = node.v.toLowerCase();
      if (!node.sheet && ctx.local?.has(key)) {
        const v = ctx.local.get(key);
        if (isError(v)) throw v;
        return v;
      }
      const v = ctx.name?.(node.v, node.sheet);
      if (v === undefined) {
        // 함수 이름만 쓰면 LAMBDA 처럼 동작 (GROUPBY(..., SUM))
        const fn = !node.sheet && FUNCS[stripFnPrefix(node.v.toUpperCase().replace(/^_XLETA\./, ''))];
        if (fn && !fn.lazy) {
          const lam = new Lambda(['x'], null, ctx);
          lam.builtin = fn;
          return lam;
        }
        throw ERR.NAME;
      }
      if (isError(v)) throw v;
      return v;
    }
    case 'range': {
      const a = evalAny(node.a, ctx);
      const b = evalAny(node.b, ctx);
      if (!(a instanceof RefValue) || !(b instanceof RefValue)) throw ERR.VALUE;
      if ((a.sheet ?? '').toLowerCase() !== (b.sheet ?? '').toLowerCase() && a.sheet && b.sheet) throw ERR.VALUE;
      return new RefValue(a.sheet ?? b.sheet, Math.min(a.r1, b.r1), Math.min(a.c1, b.c1), Math.max(a.r2, b.r2), Math.max(a.c2, b.c2));
    }
    case 'union':
      return node.items.flatMap((it) => {
        const v = evalAny(it, ctx);
        if (Array.isArray(v)) return v;
        if (!(v instanceof RefValue)) throw ERR.VALUE;
        return [v];
      });
    case 'at':
      return intersect(evalAny(node.a, ctx), ctx);
    case 'neg':
      return mapValue(deref(evalAny(node.a, ctx), ctx), (x) => -toNum(x));
    case 'pos':
      return deref(evalAny(node.a, ctx), ctx);
    case 'pct':
      return mapValue(deref(evalAny(node.a, ctx), ctx), (x) => toNum(x) / 100);
    case 'bin': {
      const a = deref(evalAny(node.a, ctx), ctx);
      const b = deref(evalAny(node.b, ctx), ctx);
      const op = node.op;
      if (!(a instanceof Range) && !(b instanceof Range)) return binary(op, a, b);
      return broadcast2(a, b, (x, y) => binary(op, x, y));
    }
    case 'func':
      return callFunc(node, ctx);
    case 'call': {
      const fn = evalAny(node.fn, ctx);
      if (!(fn instanceof Lambda)) throw ERR.VALUE;
      return callLambda(fn, node.args.map((a) => (a.type === 'empty' ? OMITTED : attempt(() => evalAny(a, ctx)))));
    }
    default:
      throw ERR.VALUE;
  }
}

/** 암시적 교차 '@' */
function intersect(v, ctx) {
  if (Array.isArray(v)) {
    if (v.length !== 1) throw ERR.VALUE;
    v = v[0];
  }
  if (v instanceof RefValue) {
    if (v.single) return v;
    const here = ctx.here;
    if (!here) throw ERR.VALUE;
    // 엑셀: 다른 시트의 범위도 수식 칸의 행 · 열 번호로 교차 (=@Sheet2!C2:D2 를 C 열에 쓰면 Sheet2!C2)
    let r;
    let c;
    if (v.r1 === v.r2) r = v.r1;
    else if (here.r >= v.r1 && here.r <= v.r2) r = here.r;
    else throw ERR.VALUE;
    if (v.c1 === v.c2) c = v.c1;
    else if (here.c >= v.c1 && here.c <= v.c2) c = here.c;
    else throw ERR.VALUE;
    return new RefValue(v.sheet, r, c);
  }
  if (v instanceof Range) {
    const x = v.rows[0]?.[0];
    if (isError(x)) throw x;
    return x ?? null;
  }
  return v;
}

function callLambda(fn, args) {
  if (fn.builtin) {
    const vals = args.map((a) => attempt(() => deref(a, fn.ctx)));
    return fn.builtin(vals, fn.ctx, evOf(fn.ctx));
  }
  if (args.length > fn.params.length) throw ERR.VALUE;
  if (lambdaDepth > 1000) throw ERR.NUM;
  const child = Object.create(fn.ctx);
  child.local = new Map(fn.ctx.local ?? []);
  fn.params.forEach((p, i) => child.local.set(p, i < args.length ? args[i] : OMITTED));
  lambdaDepth++;
  try {
    return evalAny(fn.body, child);
  } finally {
    lambdaDepth--;
  }
}

function lookupLambda(name, ctx) {
  const key = name.toLowerCase();
  let v;
  if (ctx.local?.has(key)) v = ctx.local.get(key);
  else v = ctx.name?.(name, null);
  return v instanceof Lambda ? v : null;
}

// 참조 인수의 글자 · 논리값을 무시하는 함수 (SUM(A1) 에서 A1 이 글자면 0, 직접 인수 "abc" 만 #VALUE!): 칸 하나 참조도 범위로 넘김
const REF_AS_RANGE = new Set(['SUM', 'PRODUCT', 'SUMSQ', 'AVERAGE', 'AVERAGEA', 'COUNT', 'MAX', 'MAXA', 'MIN', 'MINA', 'MEDIAN', 'MODE', 'MODE.SNGL',
  'STDEV', 'STDEV.S', 'STDEV.P', 'STDEVP', 'STDEVA', 'STDEVPA', 'VAR', 'VAR.S', 'VAR.P', 'VARP', 'VARA', 'VARPA', 'GEOMEAN', 'HARMEAN', 'AVEDEV', 'DEVSQ', 'KURT', 'SKEW']);
function callFunc(node, ctx) {
  const fn = FUNCS[node.name];
  const ev = evOf(ctx);
  if (!fn) {
    const lam = lookupLambda(node.name, ctx);
    if (!lam) throw ERR.NAME;
    return callLambda(lam, node.args.map((a) => (a.type === 'empty' ? OMITTED : attempt(() => evalAny(a, ctx)))));
  }
  if (fn.lazy) return fn(node.args, ctx, ev);
  if (fn.ref) {
    const args = node.args.map((a) => (a.type === 'empty' ? null : attempt(() => evalAny(a, ctx))));
    return fn(args, ctx, ev);
  }
  const args = [];
  for (const a of node.args) {
    if (a.type === 'union') {
      // 합집합은 여러 인수로 (SUM((A1:A3,C1:C3)))
      for (const it of attempt(() => evalAny(a, ctx))) args.push(attempt(() => derefSoft(it, ctx)));
      continue;
    }
    if (a.type === 'empty') { args.push(null); continue; }
    if (a.type === 'ref' && REF_AS_RANGE.has(node.name)) {
      const v = attempt(() => evalAny(a, ctx));
      args.push(v instanceof RefValue && v.single ? new Range([[ctx.cell(v.sheet, v.r1, v.c1)]], v) : attempt(() => derefSoft(v, ctx)));
      continue;
    }
    args.push(attempt(() => derefSoft(evalAny(a, ctx), ctx)));
  }
  return fn(args, ctx, ev);
}

// 오류는 던지지 않고 값으로 돌려줌 (IFERROR 로 감싼 x/0 이 수십만 개일 때 예외는 브라우저에서 매우 느림). 앞 피연산자의 오류가 먼저 (엑셀과 같음)
function binary(op, a, b) {
  if (isError(a)) return a;
  if (isError(b)) return b;
  switch (op) {
    case '+': return checkNum(toNum(a) + toNum(b));
    case '-': return checkNum(toNum(a) - toNum(b));
    case '*': return checkNum(toNum(a) * toNum(b));
    case '/': {
      const x = toNum(a);
      const y = toNum(b);
      if (y === 0) return ERR.DIV0;
      return checkNum(x / y);
    }
    case '^': {
      const x = toNum(a);
      const y = toNum(b);
      if (x === 0 && y === 0) throw ERR.NUM;
      if (x === 0 && y < 0) throw ERR.DIV0;
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

/** 수식 결과 정리: 여러 셀이면 Range, 아니면 스칼라/오류 값 */
function finish(v, ctx) {
  if (v instanceof RefValue || Array.isArray(v)) v = derefSoft(v, ctx);
  if (v === OMITTED) return 0;
  if (v instanceof Lambda) return ERR.CALC;
  if (v instanceof Range) {
    if (!v.height || !v.width) return ERR.CALC;
    if (v.height === 1 && v.width === 1 && !v.formats && !v.cellFormats) v = v.rows[0][0];
    else return v;
  }
  if (v === null || v === undefined) return 0;
  if (typeof v === 'number' && !Number.isFinite(v)) return ERR.NUM;
  return v;
}

/** 셀 수식 평가 결과 (배열이면 Range 그대로 — 분산은 통합 문서가 처리) */
export function evaluateArray(ast, ctx) {
  try {
    return finish(evalAny(ast, ctx), ctx);
  } catch (e) {
    if (e instanceof FormulaError) return e;
    throw e; // RangeError(너무 깊은 재귀)는 통합 문서에서 처리
  }
}

/** 스칼라 결과 (배열이면 왼쪽 위 값). 조건부 서식 · 유효성 검사용 */
export function evaluateFormula(ast, ctx) {
  const v = evaluateArray(ast, ctx);
  if (v instanceof Range) {
    const x = v.rows[0][0];
    return x === null || x === undefined ? 0 : x;
  }
  return v;
}

/** 참조 텍스트 → RefValue (INDIRECT). A1 또는 R1C1, 시트 이름, 정의된 이름 · 표 참조 */
function refFromText(text, a1, ctx) {
  const t = String(text).trim();
  if (!t) return null;
  if (a1) {
    try {
      const ast = parse(t);
      if (['ref', 'sref', 'name', 'range', 'spill'].includes(ast.type)) {
        // 글자로 만든 참조는 공유 수식의 위치 이동과 무관
        let c2 = ctx;
        if (ctx.dr || ctx.dc) { c2 = Object.create(ctx); c2.dr = 0; c2.dc = 0; }
        const v = evalAny(ast, c2);
        return v instanceof RefValue ? v : null;
      }
    } catch (e) {
      if (e instanceof FormulaError) return null;
      return null;
    }
    return null;
  }
  const m = /^(?:(?:'((?:[^']|'')+)'|([^!]+))!)?R(\[?-?\d*\]?)C(\[?-?\d*\]?)(?::R(\[?-?\d*\]?)C(\[?-?\d*\]?))?$/i.exec(t);
  if (!m) {
    // R1C1 형식이어도 정의된 이름 · 표 참조(표1[일자])는 그대로 씀 (엑셀과 같음)
    try {
      const ast = parse(t);
      if (ast.type !== 'sref' && ast.type !== 'name') return null;
      const v = evalAny(ast, ctx);
      return v instanceof RefValue ? v : null;
    } catch {
      return null;
    }
  }
  const here = ctx.here ?? { r: 0, c: 0 };
  const part = (s, base) => {
    if (s === '' || s === undefined) return base;
    if (s.startsWith('[')) return base + Number(s.slice(1, -1));
    return Number(s) - 1;
  };
  const sheet = m[1] != null ? m[1].replace(/''/g, "'") : m[2] ?? null;
  const r1 = part(m[3], here.r);
  const c1 = part(m[4], here.c);
  const r2 = m[5] !== undefined ? part(m[5], here.r) : r1;
  const c2 = m[6] !== undefined ? part(m[6], here.c) : c1;
  if ([r1, c1, r2, c2].some((x) => !Number.isFinite(x) || x < 0)) return null;
  return new RefValue(sheet, r1, c1, r2, c2);
}

// ───────────────────────── LET · LAMBDA ─────────────────────────
const CORE = {
  LET: lazy((args, ctx) => {
    if (args.length < 3 || args.length % 2 === 0) throw ERR.VALUE;
    const child = Object.create(ctx);
    child.local = new Map(ctx.local ?? []);
    for (let i = 0; i + 1 < args.length; i += 2) {
      const n = args[i];
      if (n.type !== 'name') throw ERR.NAME;
      child.local.set(n.v.toLowerCase(), attempt(() => evalAny(args[i + 1], child)));
    }
    return evalAny(args[args.length - 1], child);
  }),
  LAMBDA: lazy((args, ctx) => {
    if (!args.length) throw ERR.VALUE;
    const params = args.slice(0, -1).map((p) => {
      if (p.type !== 'name') throw ERR.VALUE;
      return p.v.toLowerCase();
    });
    if (new Set(params).size !== params.length) throw ERR.VALUE;
    return new Lambda(params, args[args.length - 1], ctx);
  }),
  ISOMITTED: lazy((args, ctx) => {
    const n = args[0];
    if (!n || n.type !== 'name') throw ERR.VALUE;
    return ctx.local?.get(n.v.toLowerCase()) === OMITTED;
  }),
};

export const FUNCS = { ...MATH, ...STAT, ...TEXT, ...DATE, ...LOOKUP, ...LOGIC, ...FIN, ...WEB, ...CORE };
FUNCS.POWER = FUNCS.POWER ?? ((args) => binary('^', args[0], args[1]));

export const FUNCTION_NAMES = Object.keys(FUNCS).sort();

/** 결과가 배열일 수 있는 함수 (분산 후보 판단용) */
const ARRAY_FUNCS = new Set([
  'SEQUENCE', 'RANDARRAY', 'FILTER', 'SORT', 'SORTBY', 'UNIQUE', 'TRANSPOSE', 'TAKE', 'DROP', 'EXPAND', 'CHOOSECOLS', 'CHOOSEROWS',
  'VSTACK', 'HSTACK', 'TOCOL', 'TOROW', 'WRAPCOLS', 'WRAPROWS', 'MMULT', 'MINVERSE', 'MUNIT', 'FREQUENCY', 'TREND', 'GROWTH',
  'LINEST', 'LOGEST', 'MODE.MULT', 'TEXTSPLIT', 'REGEXEXTRACT', 'MAP', 'SCAN', 'BYROW', 'BYCOL', 'MAKEARRAY', 'REDUCE', 'LET',
  'LAMBDA', 'OFFSET', 'INDIRECT', 'INDEX', 'XLOOKUP', 'CHOOSE', 'GROUPBY', 'PIVOTBY', ...WEB_ARRAY,
]);
/** 인수 중 어느 하나라도 배열이면 결과도 배열일 수 있는 함수 */
const PASS_FUNCS = new Set(['IF', 'IFS', 'SWITCH', 'IFERROR', 'IFNA', 'ROW', 'COLUMN', 'CONCATENATE']);

/** 수식 결과가 여러 셀(분산)일 수 있는지 정적으로 판단 */
export function mayReturnArray(node) {
  if (!node) return false;
  switch (node.type) {
    case 'ref': return node.ref.range && (node.ref.r1 !== node.ref.r2 || node.ref.c1 !== node.ref.c2);
    case 'sref': return !/^\s*@|#this row|#이 행/i.test(node.spec);
    case 'spill': case 'array': case 'name': case 'range': case 'call': return true;
    case 'paren': case 'neg': case 'pos': case 'pct': return mayReturnArray(node.a);
    case 'bin': return mayReturnArray(node.a) || mayReturnArray(node.b);
    case 'func': {
      // XLOOKUP · INDEX 는 결과가 한 칸인 흔한 모양을 구별 (전체 열 조회 수천 개를 분산 후보로 미리 계산하지 않게)
      if (node.name === 'XLOOKUP' || node.name === 'INDEX') {
        const a = node.args;
        const oneD = (n) => (n?.type === 'ref' ? (!n.ref.range || n.ref.c1 === n.ref.c2 || n.ref.r1 === n.ref.r2 ? 1 : 2)
          : n?.type === 'sref' && !/:|#all|#data|#headers|#totals|#머리글|#데이터|#전체/i.test(n.spec) && !/^\s*$/.test(n.spec) ? 1 : 0);
        if (node.name === 'XLOOKUP') {
          if (a.some((x, i) => i !== 1 && i !== 2 && mayReturnArray(x))) return true;
          return oneD(a[2]) !== 1;
        }
        const zero = (n) => !n || n.type === 'empty' || (n.type === 'num' && n.v === 0);
        if (mayReturnArray(a[1]) || mayReturnArray(a[2])) return true;
        const d = oneD(a[0]);
        if (d === 1) return zero(a[1]) && (a.length < 3 || zero(a[2]));
        if (d === 2) return zero(a[1]) || a.length < 3 || zero(a[2]);
        return true;
      }
      // INDIRECT(ADDRESS(…)) 는 늘 한 칸 · 인수 없는 ROW() · COLUMN() 도 한 칸
      if (node.name === 'INDIRECT' && node.args[0]?.type === 'func' && node.args[0].name === 'ADDRESS') return node.args[0].args.some(mayReturnArray);
      if ((node.name === 'ROW' || node.name === 'COLUMN') && !node.args.length) return false;
      if (ARRAY_FUNCS.has(node.name)) return true;
      const fn = FUNCS[node.name];
      if (!fn) return true;
      if (PASS_FUNCS.has(node.name)) return node.args.some(mayReturnArray);
      if (fn.liftPos === 'all') return node.args.some(mayReturnArray);
      if (Array.isArray(fn.liftPos)) return fn.liftPos.some((i) => mayReturnArray(node.args[i]));
      return false;
    }
    default: return false;
  }
}

/** 결과에 날짜 서식을 자동 적용해야 하는 함수 */
export function autoFormatFor(ast) {
  if (ast?.type !== 'func') return null;
  if (ast.name === 'NOW') return 'datetime';
  if (['TODAY', 'DATE', 'EDATE', 'EOMONTH', 'WORKDAY', 'WORKDAY.INTL', 'DATEVALUE'].includes(ast.name)) return 'date';
  if (['TIME', 'TIMEVALUE'].includes(ast.name)) return 'time';
  if (ast.name === 'EPOCHTODATE') return 'datetime';
  if (ast.name === 'TO_DATE') return 'date';
  if (ast.name === 'TO_PERCENT') return 'percent';
  if (ast.name === 'TO_DOLLARS') return 'currency';
  return null;
}

// ───────────────────────── 참조 재작성 ─────────────────────────
function refText(ref, sheetPrefix) {
  const cell = (r, c, ar, ac) => `${ac ? '$' : ''}${colToName(c)}${ar ? '$' : ''}${r + 1}`;
  if (ref.cols) {
    return `${sheetPrefix}${ref.ac1 ? '$' : ''}${colToName(ref.c1)}:${ref.ac2 ? '$' : ''}${colToName(ref.c2)}`;
  }
  if (ref.rows) {
    return `${sheetPrefix}${ref.ar1 ? '$' : ''}${ref.r1 + 1}:${ref.ar2 ? '$' : ''}${ref.r2 + 1}`;
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
    if (!ref.rows) {
      if (!ref.ac1) ref.c1 += dc;
      if (!ref.ac2) ref.c2 += dc;
    }
    if (ref.r1 < 0 || ref.r2 < 0 || ref.c1 < 0 || ref.c2 < 0) return null;
    if (ref.r1 >= MAX_ROWS || ref.r2 >= MAX_ROWS || ref.c1 >= MAX_COLS || ref.c2 >= MAX_COLS) return null;
    return ref;
  });
}

/**
 * 행/열 삽입(count>0)·삭제(count<0) 시 참조 조정.
 * targetSheet: 구조가 바뀐 시트 이름, hostSheet: 수식이 있는 시트 이름
 */
export function adjustFormulaForStructure(formula, { targetSheet, hostSheet, axis, index, count, band = null }) {
  const same = (a, b) => a.toLowerCase() === b.toLowerCase();
  return rewriteRefs(formula, (ref) => {
    if (!same(ref.sheet ?? hostSheet ?? '', targetSheet)) return undefined;
    if (axis === 'row' && ref.cols) return undefined;
    if (axis === 'col' && ref.rows) return undefined;
    // 셀 삽입 · 삭제(밀기): 다른 축으로 [band] 안에 완전히 들어 있는 참조만 옮김 (엑셀과 같음)
    if (band) {
      if (axis === 'row' ? ref.rows : ref.cols) return undefined;
      const o1 = axis === 'row' ? ref.c1 : ref.r1;
      const o2 = axis === 'row' ? ref.c2 : ref.r2;
      if (o1 < band[0] || o2 > band[1]) return undefined;
    }
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

/** 셀 이동: src 안에 완전히 들어 있는 참조를 (dr, dc) 이동. destinationSheet가 있으면 다른 시트로 이동. */
export function moveRefsInFormula(formula, { targetSheet, hostSheet, src, dr, dc, destinationSheet = targetSheet }) {
  const same = (a, b) => a.toLowerCase() === b.toLowerCase();
  return rewriteRefs(formula, (ref) => {
    if (!same(ref.sheet ?? hostSheet ?? '', targetSheet) || ref.rows || ref.cols) return undefined;
    if (ref.r1 < src.r1 || ref.r2 > src.r2 || ref.c1 < src.c1 || ref.c2 > src.c2) return undefined;
    ref.r1 += dr; ref.r2 += dr; ref.c1 += dc; ref.c2 += dc;
    if (!same(destinationSheet, targetSheet)) {
      ref.sheet = destinationSheet;
      ref.sheetPrefix = `${quoteSheetName(destinationSheet)}!`;
    }
    return ref;
  });
}

/** 시트 이름 변경 시 참조 갱신 */
export function renameSheetInFormula(formula, oldName, newName) {
  if (!formula.startsWith('=')) return formula;
  const src = formula.slice(1);
  let toks;
  try { toks = tokenize(src); } catch { return formula; }
  let out = '';
  let pos = 0;
  for (const t of toks) {
    const sheet = t.t === 'ref' ? t.ref.sheet : t.t === 'ident' ? t.sheet : null;
    if (!sheet || sheet.toLowerCase() !== oldName.toLowerCase()) continue;
    const text = src.slice(t.s, t.e);
    const bang = text.lastIndexOf('!');
    out += src.slice(pos, t.s) + quoteSheetName(newName) + text.slice(bang);
    pos = t.e;
  }
  return '=' + out + src.slice(pos);
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
    } else if (t.t === 'func' && !FUNCS[t.v]) {
      continue; // 사용자 LAMBDA 이름은 그대로
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

/** 수식에 쓰인 함수 이름 중 지원하지 않는 것 */
export function unknownFunctions(body, isName = () => false) {
  let toks;
  try { toks = tokenize(body); } catch { return []; }
  return [...new Set(toks.filter((t) => t.t === 'func' && !FUNCS[t.v] && !isName(t.v)).map((t) => t.v))];
}
