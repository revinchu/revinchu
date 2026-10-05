import { isPivotSnapshot, cubeFromPivotSnapshot } from './pivot-cache-data.js';
// 피벗 테이블 계산 (DOM 없음)
// 정의(def): { name, source, range | table, rows: [필드 이름], cols: [필드 이름], values: [{ field, agg, name, showAs, numFmt }],
//              pages: [필드 이름], filters: { 필드 이름: [보이는 항목 글자] }, layout: 'compact' | 'outline' | 'tabular',
//              subtotals, grandRows, grandCols, top, left, area,
//              calcFields: [{ name, formula }]            계산 필드 (합계에 수식 적용, 엑셀과 같음)
//              sort: { 필드: { dir: 'asc'|'desc', by: 값 번호 | undefined } }   항목 정렬 (by 가 있으면 값 기준)
//              order: { 필드: [항목 글자] }               수동 순서 (파일의 항목 순서)
//              fieldFilters: { 필드: { type: 'top', top, n, mode: 'count'|'percent'|'sum', by }
//                                  | { type: 'label', op, v1, v2 } | { type: 'value', op, by, v1, v2 } }
//              style: 'PivotStyleLight16' 등, rowCaption, colCaption, cellFmt: { 역할: 서식 } (파일에서 가져온 셀 서식) }
// 옛 정의 { rowField, colField, valueField, agg, fieldNames } 도 그대로 읽음
import { formatGeneral, formatValue, parseInput, serialOf, dateParts } from './format.js';
import { findTable, dataTop, dataBottom, columnNames, ACCENTS, tint, shade } from './tables.js';
import { logicalCol } from './block.js';
import { presetStyle, presetSwatch, paintPivotPreset, MODERN_STYLES, styleElementsPreset } from './stylepresets.js';
import {
  EMPTY as EMPTY0, IMG_KEY as IMG_KEY0, keyOf as keyOf0, imageOfKey as imageOfKey0, sortKeys as sortKeys0, itemText as itemText0,
  kk, itemIdentity, itemProperty, cubeFromRows, filterRows, groupAggregate, groupAcc, Cube, Column, blockColumn, groupedColumn, groupRank, groupKey, aggregateQuery, planRollup, GROUP_BY as GROUP_BY0,
} from './cube.js';
import { maxOf, minOf, pushAll, ERR, ERR_BY_CODE } from './fxcore.js';

export const AGGREGATES = [
  { id: 'sum', label: '합계' },
  { id: 'count', label: '개수' },
  { id: 'average', label: '평균' },
  { id: 'max', label: '최대' },
  { id: 'min', label: '최소' },
  { id: 'product', label: '곱' },
  { id: 'countNums', label: '숫자 개수' },
  { id: 'stdDev', label: '표준 편차' },
  { id: 'stdDevp', label: '표준 편차(전체)' },
  { id: 'var', label: '분산' },
  { id: 'varp', label: '분산(전체)' },
];
export const SHOW_AS = [
  { id: 'normal', label: '계산 없음' },
  { id: 'percentOfTotal', label: '총합계 비율' },
  { id: 'percentOfCol', label: '열 합계 비율' },
  { id: 'percentOfRow', label: '행 합계 비율' },
  { id: 'percent', label: '기준값 [%]', base: 'item' },
  { id: 'percentOfParentRow', label: '상위 행 합계 비율' },
  { id: 'percentOfParentCol', label: '상위 열 합계 비율' },
  { id: 'percentOfParent', label: '상위 합계 비율', base: 'field' },
  { id: 'difference', label: '차이', base: 'item' },
  { id: 'percentDiff', label: '[%] 차이', base: 'item' },
  { id: 'runTotal', label: '누계', base: 'field' },
  { id: 'percentOfRunningTotal', label: '누계 비율', base: 'field' },
  { id: 'rankAscending', label: '오름차순 순위 지정', base: 'field' },
  { id: 'rankDescending', label: '내림차순 순위 지정', base: 'field' },
  { id: 'index', label: '인덱스' },
];
/** 값 표시 형식이 비율(%)인지 (차이 · 누계 · 순위는 원래 숫자 서식) */
export const showAsPercent = (as) => !!as && as !== 'normal' && !['difference', 'runTotal', 'rankAscending', 'rankDescending', 'index'].includes(as);
/** 기준 항목: 항목 글자 또는 이전 · 다음 (basePos: 'prev' | 'next') */
export const BASE_POS = [{ id: 'prev', label: '(이전)' }, { id: 'next', label: '(다음)' }];
export const LAYOUTS = [
  { id: 'compact', label: '압축 형식으로 표시' },
  { id: 'outline', label: '개요 형식으로 표시' },
  { id: 'tabular', label: '테이블 형식으로 표시' },
];

// 항목 키 · 글자 · 정렬은 열 기반 엔진(cube.js)과 함께 씀
export const EMPTY = EMPTY0;
export const IMG_KEY = IMG_KEY0;
export const keyOf = keyOf0;
export const imageOfKey = imageOfKey0;
export const sortKeys = sortKeys0;
export const itemText = itemText0;
export const GROUP_BY = GROUP_BY0;
export const TOTAL = '총합계';

/** 머리글 이름 (빈 칸은 열N) */
export const headerNames = (rows) => (rows && !Array.isArray(rows) && rows.cube ? rows.cube.header : (rows[0] ?? []).map((h, i) => (h === null || h === '' ? `열${i + 1}` : String(h))));

export const aggLabel = (agg) => AGGREGATES.find((a) => a.id === agg)?.label ?? '합계';
export const valueName = (v) => v.name || `${aggLabel(v.agg)} : ${v.field}`;

// ───────────── 계산 필드 수식 (필드 이름 · 숫자 · 사칙 연산 · 몇 가지 함수) ─────────────
const CALC_ERR = (code) => ({ code });
const isErr = (v) => v !== null && typeof v === 'object';

/** 계산 필드에서 ROWS() 가 읽는 가상 필드 (그룹의 원본 행 수) */
const ROWS_FIELD = '\u0000rows';
/** 계산 필드에서 쓸 수 있는 함수 (도움말 · 삽입 단추) */
export const CALC_FUNCS = [
  ['IF', 'IF(조건, 참일 때, 거짓일 때)'], ['IFERROR', 'IFERROR(수식, 오류일 때)'], ['DIVIDE', 'DIVIDE(분자, 분모, [0일 때]) — 0으로 나눠도 오류 없음'],
  ['ROWS', 'ROWS() — 그룹의 원본 행 수 (엑셀에 없음)'], ['ROUND', 'ROUND(수, 자릿수)'], ['ROUNDUP', 'ROUNDUP(수, 자릿수)'], ['ROUNDDOWN', 'ROUNDDOWN(수, 자릿수)'],
  ['ABS', 'ABS(수)'], ['SQRT', 'SQRT(수)'], ['INT', 'INT(수)'], ['TRUNC', 'TRUNC(수, [자릿수])'], ['MOD', 'MOD(수, 나누는 수)'], ['POWER', 'POWER(수, 지수)'],
  ['EXP', 'EXP(수)'], ['LN', 'LN(수)'], ['LOG', 'LOG(수, [밑])'], ['LOG10', 'LOG10(수)'], ['SIGN', 'SIGN(수)'], ['PI', 'PI()'],
  ['SUM', 'SUM(수1, 수2, …)'], ['MIN', 'MIN(수1, 수2, …)'], ['MAX', 'MAX(수1, 수2, …)'], ['AVERAGE', 'AVERAGE(수1, 수2, …)'],
  ['AND', 'AND(조건1, …)'], ['OR', 'OR(조건1, …)'], ['NOT', 'NOT(조건)'],
];
const CALC_FUNC_SET = new Set(CALC_FUNCS.map(([n]) => n));
const CALC_ARITY = {
  IF: [1, 3], IFERROR: [2, 2], DIVIDE: [2, 3], ROWS: [0, 0], PI: [0, 0],
  ROUND: [2, 2], ROUNDUP: [2, 2], ROUNDDOWN: [2, 2], TRUNC: [1, 2], LOG: [1, 2],
  ABS: [1, 1], SQRT: [1, 1], INT: [1, 1], MOD: [2, 2], POWER: [2, 2], EXP: [1, 1],
  LN: [1, 1], LOG10: [1, 1], SIGN: [1, 1], NOT: [1, 1], AND: [1, 255], OR: [1, 255],
  SUM: [1, 255], MIN: [1, 255], MAX: [1, 255], AVERAGE: [1, 255],
};
const validCalcArgs = (n) => !CALC_ARITY[n.fn] || (n.args.length >= CALC_ARITY[n.fn][0] && n.args.length <= CALC_ARITY[n.fn][1]);

/**
 * 계산 필드 수식 검사 → { ok, error, unknown: [없는 필드], funcs: [없는 함수], refs: [참조 필드], cycle }
 * fields: 원본 필드 이름, calcFields: 계산 필드 목록, selfName: 검사하는 계산 필드 이름 (순환 확인)
 */
export function checkCalc(formula, fields, calcFields = [], selfName = null) {
  let ast;
  try { ast = parseCalc(formula); } catch { return { ok: false, error: '수식 구문이 올바르지 않습니다 (괄호 · 연산자 · 따옴표를 확인하세요).', unknown: [], funcs: [], refs: [] }; }
  const lower = (x) => String(x).toLowerCase();
  const known = new Set(fields.map(lower));
  const calcs = new Map(calcFields.filter((c) => lower(c.name) !== lower(selfName ?? '')).map((c) => [lower(c.name), c]));
  const refs = [...calcRefs(ast)];
  const unknown = refs.filter((r) => !known.has(lower(r)) && !calcs.has(lower(r)) && lower(r) !== lower(selfName ?? ''));
  const funcs = [];
  const badArgs = [];
  const walk = (n) => { if (!n) return; if (n.fn && !CALC_FUNC_SET.has(n.fn)) funcs.push(n.fn); else if (n.fn && !validCalcArgs(n)) badArgs.push(n.fn); ['a', 'b'].forEach((k) => walk(n[k])); (n.args ?? []).forEach(walk); };
  walk(ast);
  // 다른 계산 필드를 거쳐 자기 자신을 참조하면 순환
  let cycle = false;
  if (selfName) {
    const seen = new Set();
    const visit = (names) => names.forEach((nm) => {
      const k = lower(nm);
      if (k === lower(selfName)) { cycle = true; return; }
      const c = calcs.get(k);
      if (!c || seen.has(k)) return;
      seen.add(k);
      try { visit([...calcRefs(parseCalc(c.formula))]); } catch { /* 무시 */ }
    });
    visit(refs);
  }
  const error = unknown.length ? `없는 필드: ${unknown.join(', ')}` : funcs.length ? `지원하지 않는 함수: ${[...new Set(funcs)].join(', ')}` : badArgs.length ? `인수 개수가 올바르지 않습니다: ${[...new Set(badArgs)].join(', ')}` : cycle ? '순환 참조: 이 계산 필드를 다시 참조합니다.' : null;
  return { ok: !error, error, unknown, funcs, refs, cycle };
}

/**
 * 엑셀 파일에 쓸 계산 필드 수식: 엑셀에 없는 함수를 바꿈 (DIVIDE(a,b,c) → IF(b=0,c,a/b), ROWS() → 0)
 * 바꿀 것이 없으면 원래 글자 그대로
 */
export function excelCalcFormula(formula) {
  let ast;
  try { ast = parseCalc(formula); } catch { return formula; }
  let changed = false;
  const q = (f) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(f) ? f : `'${f.replace(/'/g, "''")}'`);
  const PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };
  const txt = (n, parent = 0) => {
    if ('num' in n) return String(n.num);
    if ('str' in n) return `"${n.str.replace(/"/g, '""')}"`;
    if (n.field !== undefined) return q(n.field);
    if (n.fn) {
      if (n.fn === 'DIVIDE') { changed = true; const [a, b, c] = n.args; return `IF(${txt(b)}=0,${c ? txt(c) : '0'},${txt(a, 4)}/${txt(b, 5)})`; }
      if (n.fn === 'ROWS') { changed = true; return '0'; }
      return `${n.fn}(${n.args.map((x) => txt(x)).join(',')})`;
    }
    if (n.op === 'neg') return `-${txt(n.a, 6)}`;
    if (n.op === 'pct') return `${txt(n.a, 7)}%`;
    const p = PREC[n.op];
    const s = `${txt(n.a, p)}${n.op}${txt(n.b, p + 1)}`;
    return p < parent ? `(${s})` : s;
  };
  const out = txt(ast);
  return changed ? out : formula;
}

/** 계산 필드 수식 안의 필드 이름 바꾸기 (이름을 바꾸면 이를 참조하는 다른 계산 필드도 따라 바뀜) */
export function renameCalcRefs(formula, from, to) {
  const quote = (f) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(f) ? f : `'${f.replace(/'/g, "''")}'`);
  const src = String(formula).replace(/^\s*=/, '');
  let out = '';
  let i = 0;
  const low = from.toLowerCase();
  while (i < src.length) {
    const ch = src[i];
    if (ch === "'" || ch === '"') {
      let j = i + 1;
      let v = '';
      while (j < src.length) { if (src[j] === ch) { if (src[j + 1] === ch) { v += ch; j += 2; continue; } break; } v += src[j++]; }
      out += ch === "'" && v.toLowerCase() === low ? quote(to) : src.slice(i, j + 1);
      i = j + 1;
      continue;
    }
    const nm = /^[^\s\-+*/^&=<>(),;'"]+/.exec(src.slice(i));
    if (nm && !/^(\d+\.?\d*|\.\d+)/.test(nm[0])) {
      const isFn = src.slice(i + nm[0].length).trimStart().startsWith('(');
      out += !isFn && nm[0].toLowerCase() === low ? quote(to) : nm[0];
      i += nm[0].length;
      continue;
    }
    out += ch;
    i++;
  }
  return out;
}

export function parseCalc(src) {
  const text = String(src ?? '').replace(/^\s*=/, '');
  const toks = [];
  let i = 0;
  while (i < text.length) {
    const ch = text[i];
    if (/\s/.test(ch)) { i++; continue; }
    const num = /^(\d+\.?\d*|\.\d+)(e[+-]?\d+)?%?/i.exec(text.slice(i));
    if (num) { const t = num[0]; toks.push({ t: 'num', v: t.endsWith('%') ? Number(t.slice(0, -1)) / 100 : Number(t) }); i += t.length; continue; }
    if (ch === "'" || ch === '"') {
      let j = i + 1;
      let v = '';
      while (j < text.length) {
        if (text[j] === ch) { if (text[j + 1] === ch) { v += ch; j += 2; continue; } break; }
        v += text[j++];
      }
      toks.push({ t: ch === "'" ? 'qname' : 'str', v });
      i = j + 1;
      continue;
    }
    const op = /^(<>|<=|>=|[-+*/^&=<>(),;%])/.exec(text.slice(i));
    if (op) { toks.push({ t: 'op', v: op[0] === ';' ? ',' : op[0] }); i += op[0].length; continue; }
    const nm = /^[^\s\-+*/^&=<>(),;'"]+/.exec(text.slice(i));
    if (!nm) throw new Error('수식 오류');
    toks.push({ t: 'name', v: nm[0] });
    i += nm[0].length;
  }
  let p = 0;
  const peek = () => toks[p];
  const eat = (v) => { if (peek()?.t === 'op' && peek().v === v) { p++; return true; } return false; };
  const expr = () => cmp();
  const cmp = () => { let a = cat(); for (;;) { const t = peek(); if (t?.t === 'op' && ['=', '<>', '<', '>', '<=', '>='].includes(t.v)) { p++; a = { op: t.v, a, b: cat() }; } else return a; } };
  const cat = () => { let a = add(); while (eat('&')) a = { op: '&', a, b: add() }; return a; };
  const add = () => { let a = mul(); for (;;) { if (eat('+')) a = { op: '+', a, b: mul() }; else if (eat('-')) a = { op: '-', a, b: mul() }; else return a; } };
  const mul = () => { let a = pow(); for (;;) { if (eat('*')) a = { op: '*', a, b: pow() }; else if (eat('/')) a = { op: '/', a, b: pow() }; else return a; } };
  const pow = () => { let a = un(); while (eat('^')) a = { op: '^', a, b: un() }; return a; };
  // 백분율 (후위 %): 엑셀처럼 "0"% = 0, 클릭% = 클릭/100
  const pct = () => { let a = prim(); while (eat('%')) a = { op: 'pct', a }; return a; };
  const un = () => (eat('-') ? { op: 'neg', a: un() } : eat('+') ? un() : pct());
  const prim = () => {
    const t = toks[p++];
    if (!t) throw new Error('수식 오류');
    if (t.t === 'num') return { num: t.v };
    if (t.t === 'str') return { str: t.v };
    if (t.t === 'qname') return { field: t.v };
    if (t.t === 'op' && t.v === '(') { const e = expr(); if (!eat(')')) throw new Error('수식 오류'); return e; }
    if (t.t === 'name') {
      if (peek()?.t === 'op' && peek().v === '(') {
        p++;
        const args = [];
        // 빈 인수 (IFERROR(x,) 처럼 쉼표 뒤가 비었음) 는 엑셀처럼 0
        const arg = () => (peek()?.t === 'op' && (peek().v === ',' || peek().v === ')') ? { num: 0 } : expr());
        if (!eat(')')) { do args.push(arg()); while (eat(',')); if (!eat(')')) throw new Error('수식 오류'); }
        return { fn: t.v.toUpperCase(), args };
      }
      return { field: t.v };
    }
    throw new Error('수식 오류');
  };
  const ast = expr();
  if (p < toks.length) throw new Error('수식 오류');
  return ast;
}

/** 수식이 참조하는 필드 이름 */
export function calcRefs(ast, out = new Set()) {
  if (!ast) return out;
  if (ast.field !== undefined) out.add(ast.field);
  for (const k of ['a', 'b']) if (ast[k]) calcRefs(ast[k], out);
  (ast.args ?? []).forEach((x) => calcRefs(x, out));
  return out;
}

function evalCalc(ast, get) {
  const num = (v) => (isErr(v) ? v : typeof v === 'number' ? v : typeof v === 'boolean' ? Number(v) : v === null || v === '' ? 0 : Number.isFinite(Number(v)) ? Number(v) : CALC_ERR('#VALUE!'));
  // 중첩된 IFERROR도 비정상 숫자를 오류로 처리하게 각 계산 단계에서 검사합니다.
  const ev = (n) => { const v = evRaw(n); return typeof v === 'number' && !Number.isFinite(v) ? CALC_ERR('#NUM!') : v; };
  const evRaw = (n) => {
    if (!n) return CALC_ERR('#VALUE!');
    if ('num' in n) return n.num;
    if ('str' in n) return n.str;
    if (n.field !== undefined) return get(n.field);
    if (n.fn) {
      const A = n.args;
      if (!validCalcArgs(n)) return CALC_ERR('#VALUE!');
      switch (n.fn) {
        case 'IF': { const c = num(ev(A[0])); if (isErr(c)) return c; return c ? (A[1] ? ev(A[1]) : true) : (A[2] ? ev(A[2]) : false); }
        case 'IFERROR': { const v = ev(A[0]); return isErr(v) ? ev(A[1]) : v; }
        case 'ROUND': { const v = num(ev(A[0])); const d = num(ev(A[1])); if (isErr(v)) return v; if (isErr(d)) return d; const f = 10 ** Math.trunc(d); return Math.sign(v) * Math.round(Math.abs(v) * f) / f; }
        case 'ABS': { const v = num(ev(A[0])); return isErr(v) ? v : Math.abs(v); }
        case 'SQRT': { const v = num(ev(A[0])); return isErr(v) ? v : v < 0 ? CALC_ERR('#NUM!') : Math.sqrt(v); }
        case 'INT': { const v = num(ev(A[0])); return isErr(v) ? v : Math.floor(v); }
        case 'SUM': case 'MIN': case 'MAX': case 'AVERAGE': {
          const vs = A.map((x) => num(ev(x)));
          const e = vs.find(isErr);
          if (e) return e;
          if (n.fn === 'SUM') return vs.reduce((a, b) => a + b, 0);
          if (n.fn === 'AVERAGE') return vs.length ? vs.reduce((a, b) => a + b, 0) / vs.length : 0;
          return n.fn === 'MIN' ? minOf(vs) : maxOf(vs);
        }
        case 'DIVIDE': {
          // 안전한 나누기 (엑셀에 없음): 0 으로 나누면 세 번째 인수(없으면 빈 값)
          const a = num(ev(A[0]));
          const b = num(ev(A[1]));
          if (isErr(a)) return a;
          if (isErr(b)) return b;
          return b === 0 ? (A[2] ? ev(A[2]) : null) : a / b;
        }
        case 'ROWS': return get(ROWS_FIELD);
        case 'MOD': { const a = num(ev(A[0])); const b = num(ev(A[1])); if (isErr(a)) return a; if (isErr(b)) return b; return b === 0 ? CALC_ERR('#DIV/0!') : a - b * Math.floor(a / b); }
        case 'POWER': { const a = num(ev(A[0])); const b = num(ev(A[1])); if (isErr(a)) return a; if (isErr(b)) return b; return a ** b; }
        case 'EXP': { const v = num(ev(A[0])); return isErr(v) ? v : Math.exp(v); }
        case 'LN': { const v = num(ev(A[0])); return isErr(v) ? v : v <= 0 ? CALC_ERR('#NUM!') : Math.log(v); }
        case 'LOG10': { const v = num(ev(A[0])); return isErr(v) ? v : v <= 0 ? CALC_ERR('#NUM!') : Math.log10(v); }
        case 'LOG': { const v = num(ev(A[0])); const b = A[1] ? num(ev(A[1])) : 10; if (isErr(v)) return v; if (isErr(b)) return b; return v <= 0 || b <= 0 || b === 1 ? CALC_ERR('#NUM!') : Math.log(v) / Math.log(b); }
        case 'SIGN': { const v = num(ev(A[0])); return isErr(v) ? v : Math.sign(v); }
        case 'TRUNC': { const v = num(ev(A[0])); const d = A[1] ? num(ev(A[1])) : 0; if (isErr(v)) return v; if (isErr(d)) return d; const f = 10 ** Math.trunc(d); return Math.trunc(v * f) / f; }
        case 'ROUNDUP': case 'ROUNDDOWN': {
          const v = num(ev(A[0]));
          const d = A[1] ? num(ev(A[1])) : 0;
          if (isErr(v)) return v;
          if (isErr(d)) return d;
          const f = 10 ** Math.trunc(d);
          const x = Math.abs(v) * f;
          const r = n.fn === 'ROUNDUP' ? Math.ceil(x - 1e-9) : Math.floor(x + 1e-9);
          return (Math.sign(v) * r) / f;
        }
        case 'PI': return Math.PI;
        case 'AND': case 'OR': {
          // 논리값이 결정되어도 뒤쪽 인수의 오류를 숨기면 안 됩니다.
          const vs = A.map((x) => num(ev(x)));
          const e = vs.find(isErr);
          return e ?? (n.fn === 'AND' ? vs.every(Boolean) : vs.some(Boolean));
        }
        case 'NOT': { const v = num(ev(A[0])); return isErr(v) ? v : !v; }
        default: return CALC_ERR('#NAME?');
      }
    }
    if (n.op === 'neg') { const v = num(ev(n.a)); return isErr(v) ? v : -v; }
    if (n.op === 'pct') { const v = num(ev(n.a)); return isErr(v) ? v : v / 100; }
    if (n.op === '&') { const a = ev(n.a); const b = ev(n.b); return isErr(a) ? a : isErr(b) ? b : `${a}${b}`; }
    const a = num(ev(n.a));
    const b = num(ev(n.b));
    if (isErr(a)) return a;
    if (isErr(b)) return b;
    switch (n.op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? CALC_ERR('#DIV/0!') : a / b;
      case '^': return a ** b;
      case '=': return a === b;
      case '<>': return a !== b;
      case '<': return a < b;
      case '>': return a > b;
      case '<=': return a <= b;
      case '>=': return a >= b;
      default: return CALC_ERR('#VALUE!');
    }
  };
  return ev(ast);
}

// ───────────── 누적 ─────────────
const newAcc = () => ({ count: 0, nums: 0, sum: 0, sq: 0, min: Infinity, max: -Infinity, prod: 1 });
function add(acc, v) {
  if (v !== null && v !== '' && v !== undefined) acc.count++;
  if (typeof v === 'number' && Number.isFinite(v)) {
    acc.nums++;
    acc.sum += v;
    acc.sq += v * v;
    acc.prod *= v;
    if (v < acc.min) acc.min = v;
    if (v > acc.max) acc.max = v;
  } else if (v && typeof v === 'object' && !acc.err) {
    const e = typeof v.code === 'string' && v.code[0] === '#' ? v : typeof v.error === 'string' ? { code: v.error } : null;
    if (e) acc.err = e;
  }
}
function result(acc, agg) {
  if (!acc) return null;
  // 원본에 오류 값이 있으면 합계 · 평균 등은 그 오류 (개수만 셈) — 엑셀과 같음
  if (acc.err && agg !== 'count' && agg !== 'countNums') return acc.err;
  const n = acc.nums;
  const varOf = (sample) => {
    const d = n - (sample ? 1 : 0);
    if (d <= 0) return null;
    return Math.max(0, (acc.sq - (acc.sum * acc.sum) / n) / d);
  };
  switch (agg) {
    case 'count': return acc.count;
    case 'countNums': return n;
    case 'average': return n ? acc.sum / n : null;
    case 'max': return n ? acc.max : null;
    case 'min': return n ? acc.min : null;
    case 'product': return n ? acc.prod : null;
    case 'var': return varOf(true);
    case 'varp': return varOf(false);
    case 'stdDev': { const x = varOf(true); return x === null ? null : Math.sqrt(x); }
    case 'stdDevp': { const x = varOf(false); return x === null ? null : Math.sqrt(x); }
    default: return n || acc.count ? acc.sum : null; // 값이 모두 빈 칸이면 빈칸 (엑셀)
  }
}

/**
 * 값 필드 계산기: 필요한 원본 열만 누적하고 계산 필드는 합계에 수식을 적용
 * header: 원본 머리글 (+ 계산 필드 이름)
 */
const usesRows = (ast) => !!ast && (ast.fn === 'ROWS' || ['a', 'b'].some((k) => usesRows(ast[k])) || (ast.args ?? []).some(usesRows));

/** 오류 · 빈 셀 표시 글자 → 칸 입력 (엑셀: 숫자 모양이면 숫자 0 처럼 숫자로 씀) */
const captionRaw = (t) => (!t ? '' : /^-?\d+(\.\d+)?$/.test(t) ? t : `'${t}`);
// Excel 16.0.14334 COM + PDF + GETPIVOTDATA로 확인한 오류 대체문구의 값 변환.
// 일반 셀 입력 파서와 다름: 소수·지수·퍼센트는 문자열이며 빈 문구만 빈 셀이다.
function errorCaptionValue(caption) {
  const text = String(caption ?? '');
  if (!text) return null;
  const numberText = text.replace(/[０-９－]/g, (c) => c === '－' ? '-' : String(c.charCodeAt(0) - 0xff10)).replace(/^ +| +$/g, '');
  if (/^-?\d*$/.test(numberText)) {
    const n = numberText === '-' ? 0 : Number(numberText);
    // 음수 끝의 세 값은 Excel이 숫자로 바꾸지 않는다(실제 경계값 검증).
    if (n >= -32765 && n <= 32767) return n === 0 ? 0 : n;
  }
  return text;
}
function errorCaptionRaw(caption) {
  const value = errorCaptionValue(caption);
  return value === null ? '' : typeof value === 'number' ? String(value) : `'${value}`;
}

/** 계산 필드 결과는 숫자 (엑셀: IFERROR(…,"0") 의 "0" 은 숫자 0 으로 보임) */
const calcNumber = (v) => (typeof v === 'string' && /^\s*-?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?\s*$/i.test(v) ? Number(v) : v);
function makeMeasures(header, values, calcFields) {
  const lower = (x) => String(x).toLowerCase();
  const calcByName = new Map((calcFields ?? []).map((c) => {
    let ast = null;
    try { ast = parseCalc(c.formula); } catch { ast = null; }
    return [lower(c.name), { ...c, ast }];
  }));
  const cols = [];
  const slot = new Map();
  const need = (col) => { if (!slot.has(col)) { slot.set(col, cols.length); cols.push(col); } return slot.get(col); };
  // 이름 → 원본 열 번호 (계산 필드를 칸마다 계산하므로 한 번만 찾음)
  const idxMemo = new Map();
  const baseIdx = (name) => {
    let i = idxMemo.get(name);
    if (i === undefined) { i = header.findIndex((h) => lower(h) === lower(name) && !calcByName.has(lower(h))); idxMemo.set(name, i); }
    return i;
  };
  const calcMemo = new Map();
  const calcOf = (name) => { let c = calcMemo.get(name); if (c === undefined) { c = calcByName.get(lower(name)) ?? null; calcMemo.set(name, c); } return c; };
  const specs = values.map((v) => {
    const c = calcByName.get(lower(v.field));
    if (c) {
      const walk = (ast, seen) => calcRefs(ast).forEach((n) => {
        const cc = calcByName.get(lower(n));
        if (cc) { if (!seen.has(lower(n))) walk(cc.ast, new Set([...seen, lower(n)])); } else { const i = baseIdx(n); if (i >= 0) need(i); }
      });
      if (c.ast) walk(c.ast, new Set([lower(c.name)]));
      // ROWS() (그룹의 원본 행 수): 1 을 더하는 슬롯
      for (const cc of calcByName.values()) if (cc.ast && usesRows(cc.ast)) need(-1);
      return { calc: c };
    }
    return { slot: need(baseIdx(v.field)), agg: v.agg };
  });
  const calcValue = (c, list, depth = 0) => {
    if (!c.ast) return CALC_ERR('#NAME?');
    if (depth > 20) return CALC_ERR('#REF!');
    return evalCalc(c.ast, (name) => {
      if (name === ROWS_FIELD) return slot.has(-1) ? list[slot.get(-1)]?.sum ?? 0 : 0;
      const cc = calcOf(name);
      if (cc) return calcValue(cc, list, depth + 1);
      const i = baseIdx(name);
      if (i < 0) return CALC_ERR('#NAME?');
      const a = list[slot.get(i)];
      return a?.err ?? a?.sum ?? 0;
    });
  };
  const n = cols.length;
  // 슬롯마다 필요한 누적값 (제곱합 · 최소/최대 · 곱은 쓰는 집계가 있을 때만)
  const needs = cols.map(() => ({ sq: false, mm: false, prod: false }));
  specs.forEach((sp) => {
    if (sp.calc || sp.slot === undefined) return;
    const nd = needs[sp.slot];
    if (/^(var|std)/.test(sp.agg ?? '')) nd.sq = true;
    if (sp.agg === 'min' || sp.agg === 'max') nd.mm = true;
    if (sp.agg === 'product') nd.prod = true;
  });
  return {
    cols,
    needs,
    value(list, vi) {
      if (!list) return null;
      const sp = specs[vi];
      if (!sp) return null;
      if (sp.calc) return calcNumber(calcValue(sp.calc, list));
      return result(list[sp.slot], sp.agg);
    },
    /** 데이터가 없는 총합계: 엑셀은 계산 필드만 합계 0 으로 계산하고 나머지는 빈칸 */
    emptyValue(vi) {
      const sp = specs[vi];
      return sp?.calc ? calcNumber(calcValue(sp.calc, cols.map(newAcc))) : null;
    },
    accumulate(list, r) {
      for (let k = 0; k < n; k++) add(list[k], cols[k] < 0 ? 1 : r[cols[k]]);
    },
    newList() { return cols.map(newAcc); },
  };
}

/** 누적 목록 합치기 (그룹 → 상위 합계) */
function mergeList(into, from) {
  for (let k = 0; k < into.length; k++) {
    const a = into[k];
    const b = from[k];
    a.count += b.count;
    a.nums += b.nums;
    a.sum += b.sum;
    a.sq += b.sq;
    if (b.min < a.min) a.min = b.min;
    if (b.max > a.max) a.max = b.max;
    a.prod *= b.prod;
    if (b.err && !a.err) a.err = b.err;
  }
}

// ───────────── 정의 정리 ─────────────
/** 원본 머리글 + 계산 필드 이름 */
export function pivotFieldNames(rows, def) {
  const header = headerNames(rows);
  return [...header, ...(def.calcFields ?? []).map((c) => c.name).filter((n) => !header.some((h) => h.toLowerCase() === String(n).toLowerCase()))];
}

/** 오류 대체 글자 사용 여부. 새 피벗 기본은 생성 시 설정하고 파일 옵션은 그대로 따릅니다. */
export function pivotErrorDisplay(def) {
  return typeof def.errorShow === 'boolean' ? def.errorShow : def.errorCaption !== null && def.errorCaption !== undefined;
}

/** 옛 정의까지 포함해 필드 이름 기반 정의로 (header: 원본 머리글) */
function dateDef(def, date1904) {
  if (date1904 === undefined) return def;
  return { ...def, date1904: !!date1904, groups: Object.fromEntries(Object.entries(def.groups ?? {}).map(([k, v]) => [k, { ...v, date1904: !!date1904 }])) };
}

/** 표시 옵션은 독립 보존: 클래식 끌어 놓기는 Excel처럼 값 머리글을 강제로 표시한다. */
export function pivotDisplayOptions(def) {
  const classic = !!def.classic;
  // valuesHeadRow는 이전 WIXEL 문서의 별칭. 명시적인 false를 덮어쓰면 안 된다.
  const showValuesRow = !!(def.showValuesRow ?? def.valuesHeadRow ?? false);
  return { classic, showValuesRow, valuesHeadRow: classic || showValuesRow };
}

/** 보고서 필터 배치. 좌표는 피벗 왼쪽 위 기준이며 값 칸은 c+1, 필터 열 사이는 한 칸 비웁니다. */
export function pivotPageLayout(def) {
  const pages = def.pages ?? [], n = pages.length;
  if (!n) return { fields: [], height: 0, width: 0 };
  const requested = Math.floor(Number(def.pageWrap)), wrap = Number.isFinite(requested) && requested > 0 ? Math.min(n, requested) : n;
  const count = Math.max(1, wrap), over = def.pageOrder === 'over';
  const fields = pages.map((field, i) => ({ field, r: over ? Math.floor(i / count) : i % count, c: 3 * (over ? i % count : Math.floor(i / count)) }));
  return { fields, height: over ? Math.ceil(n / count) : count, width: 3 * (over ? count : Math.ceil(n / count)) - 1 };
}

/** 보고서 필터의 체크 목록 모드. 이전 문서는 선택 수로만 복원한다. */
export function pivotPageMulti(def, field) {
  const explicit = def.pageMulti?.[field];
  return typeof explicit === 'boolean' ? explicit : Array.isArray(def.filters?.[field]) && def.filters[field].length > 1;
}

export function normalizeDef(def, header) {
  const allHeader = [...header, ...(def.calcFields ?? []).map((c) => c.name).filter((n) => !header.some((h) => h.toLowerCase() === String(n).toLowerCase()))];
  const name = (i) => (i === null || i === undefined || i < 0 ? null : allHeader[i] ?? null);
  const byName = (n) => allHeader.find((h) => h.toLowerCase() === String(n).toLowerCase()) ?? null;
  let rows = def.rows;
  let cols = def.cols;
  let values = def.values;
  if (!rows) {
    // 옛 형식: 필드 번호 + 이름
    const r = def.fieldNames?.row ? byName(def.fieldNames.row) ?? name(def.rowField) : name(def.rowField);
    const c = def.colField === null || def.colField === undefined ? null : def.fieldNames?.col ? byName(def.fieldNames.col) ?? name(def.colField) : name(def.colField);
    const v = def.valueField === null || def.valueField === undefined ? null : def.fieldNames?.val ? byName(def.fieldNames.val) ?? name(def.valueField) : name(def.valueField);
    rows = r ? [r] : [];
    cols = c ? [c] : [];
    values = [{ field: v ?? r, agg: v === null ? 'count' : def.agg ?? 'sum' }];
  }
  const ok = (n) => byName(n) !== null;
  const byKey = (obj) => Object.fromEntries(Object.entries(obj ?? {}).filter(([k]) => ok(k)).map(([k, v]) => [byName(k), v]));
  return {
    date1904: !!def.date1904,
    rows: (rows ?? []).filter(ok).map(byName),
    cols: (cols ?? []).filter(ok).map(byName),
    pages: (def.pages ?? []).filter(ok).map(byName),
    pageMulti: Object.fromEntries(Object.entries(byKey(def.pageMulti)).filter(([, value]) => typeof value === 'boolean')),
    pageOrder: def.pageOrder === 'over' ? 'over' : 'down',
    pageWrap: Number.isFinite(Number(def.pageWrap)) && Number(def.pageWrap) > 0 ? Math.min(4294967295, Math.floor(Number(def.pageWrap))) : 0,
    values: (values ?? []).filter((v) => ok(v.field)).map((v) => ({ ...v, field: byName(v.field), agg: v.agg ?? 'sum' })),
    layout: def.layout ?? 'compact',
    subtotals: Array.isArray(def.subtotals) ? def.subtotals : def.subtotals !== false,
    grandRows: def.grandRows !== false,
    grandCols: def.grandCols !== false,
    filters: def.filters ?? {},
    calcFields: def.calcFields ?? [],
    sort: byKey(def.sort),
    customListSort: def.customListSort !== false,
    order: byKey(def.order),
    tieOrder: byKey(def.tieOrder),
    tieByParent: byKey(def.tieByParent),
    tieState: def.tieState ?? null,
    fieldFilters: byKey(def.fieldFilters),
    style: def.style ?? DEFAULT_PIVOT_STYLE,
    groups: byKey(def.groups),
    styleDef: def.styleDef ?? null,
    styleElements: def.styleElements ?? null,
    errorShow: pivotErrorDisplay(def),
    errorCaption: pivotErrorDisplay(def) ? def.errorCaption ?? '' : null,
    showHeaders: def.showHeaders !== false && def.fieldCaptions !== false, // 이전 fieldCaptions=false도 머리글 표시 끄기로 복원
    fieldCaptions: def.fieldCaptions ?? null,
    grandCaption: def.grandCaption ?? null,
    itemCaptions: def.itemCaptions ?? null,
    missingCaption: def.missingCaption ?? null,
    collapsed: byKey(def.collapsed),
    showExpand: def.showExpand !== false,
    repeatLabels: !!def.repeatLabels,
    blankRows: Array.isArray(def.blankRows) ? def.blankRows : !!def.blankRows,
    subtotalTop: def.subtotalTop !== false,
    rowCaption: def.rowCaption ?? null,
    colCaption: def.colCaption ?? null,
    ...(Number.isInteger(def.valuesPos) ? { valuesPos: def.valuesPos } : {}),
    ...(def.valuesOnRows ? { valuesOnRows: true } : {}), // Σ 값을 행 영역에 (엑셀 dataOnRows)
    ...pivotDisplayOptions(def),
    ...(def.dataCaption !== undefined && def.dataCaption !== null ? { dataCaption: def.dataCaption } : {}), // 엑셀 dataCaption ('값' 대신 '데이터' 등)
    styleOpts: { rowHeaders: true, colHeaders: true, bandRows: false, bandCols: false, ...(def.styleOpts ?? {}) },
    header: allHeader,
  };
}

/**
 * 피벗 원본 → { rows (머리글 포함 값), sheet, ref: {r1,c1,r2,c2} | null, table: 표 이름 | null }
 * 표 이름이면 지금의 표 범위(누적된 데이터 포함), 아니면 고정 범위
 */
export function pivotSourceData(wb, def) {
  let si;
  let ref;
  let table = null;
  let names = null;
  if (def.table) {
    const f = findTable(wb, def.table);
    if (!f) return null;
    const t = f.t;
    si = f.si;
    ref = { r1: t.header ? t.r1 : dataTop(t), c1: t.c1, r2: dataBottom(t), c2: t.c2 };
    table = t.name;
    if (!t.header) names = columnNames(wb, f.si, t);
  } else {
    si = wb.sheetIndexByName(def.source);
    if (si < 0 || !def.range) return null;
    ref = def.range;
    // 열 전체(A1:T1048576) 원본: 사용 범위 + 빈 행 하나까지만 읽음 (엑셀의 '(비어 있음)' 항목은 그대로)
    const used = wb.usedRange?.(si);
    if (used && ref.r2 > used.rows) ref = { ...ref, r2: Math.max(ref.r1 + 1, used.rows) };
  }
  // 파일에 저장된 피벗 캐시(엑셀이 마지막으로 새로 고친 원본): 원본 시트가 그대로인 동안은 엑셀과 같은 결과가 되게 이것으로 계산
  const snap = def.snapshotId && wb.pivotSnapshots?.get(def.snapshotId);
  if (snap) {
    // 같은 시트의 결과 작성과 구별하여 실제 원본 범위의 변경 여부를 확인한다.
    const valid = wb.pivotSnapshotCurrent ? wb.pivotSnapshotCurrent(snap, def)
      : (snap.ver ??= wb.sourceVersion?.(si) ?? 0) === (wb.sourceVersion?.(si) ?? 0);
    if (valid) return isPivotSnapshot(snap.rows)
      ? sourceOf(cubeFromPivotSnapshot(snap.rows), si, ref, table, null, wb.date1904)
      : sourceOf(cubeFromRows(snap.rows), si, ref, table, snap.rows, wb.date1904);
    wb.pivotSnapshots.delete(def.snapshotId); // 원본을 고침 → 이제부터 원본에서 계산 (자동 새로 고침)
  }
  // 데이터가 열 블록에 있으면 값을 복사하지 않고 블록의 형식화 배열을 그대로 씀 (천만 행도 즉시)
  const bc = blockCube(wb, si, ref, names);
  if (bc) return sourceOf(bc, si, ref, table, null, wb.date1904);
  let rows = cachedRead(wb, si, ref);
  if (names) {
    let m = headedMemo.get(rows);
    if (!m || m[0] !== names.join('\u0001')) { m = [names.join('\u0001'), [names, ...rows]]; headedMemo.set(rows, m); }
    rows = m[1];
  }
  return sourceOf(cubeFromRows(rows), si, ref, table, rows, wb.date1904);
}
const headedMemo = new WeakMap();

/** 피벗 원본: { cube, rows(필요할 때 만듦, 머리글 포함), si, ref, table } */
function sourceOf(cube, si, ref, table, rows, date1904 = false) {
  let made = rows;
  const src = { cube, si, ref, table, date1904 };
  Object.defineProperty(src, 'rows', {
    enumerable: false,
    get() {
      if (made) return made;
      made = new Array(cube.n + 1);
      made[0] = cube.header;
      for (let i = 0; i < cube.n; i++) made[i + 1] = cube.row(i);
      return made;
    },
  });
  return src;
}

/**
 * 원본의 데이터 행이 열 블록 하나에 모두 들어 있으면 블록 기반 큐브 (없으면 null)
 * 블록에 없는 열이나 수식 셀(일반 셀)이 섞인 열은 그 열만 셀 값을 읽음
 */
const blockCubes = new WeakMap();
const headKey = (wb, si, ref, names) => (names ? names.join('\u0001') : Array.from({ length: ref.c2 - ref.c1 + 1 }, (_, j) => String(wb.getValue(si, ref.r1, ref.c1 + j) ?? '')).join('\u0001'));
function blockCube(wb, si, ref, names) {
  const sheet = wb.sheets?.[si];
  if (!sheet?.blocks?.length) return null;
  const d1 = names ? ref.r1 : ref.r1 + 1;
  const d2 = ref.r2;
  if (d2 < d1) return null;
  const b = sheet.blocks.find((x) => x.r0 <= d1 && x.r0 + x.n - 1 >= d2 && x.c0 <= ref.c2 && x.c0 + x.cols.length - 1 >= ref.c1);
  if (!b) return null;
  const sv = wb.sheetVersion?.(si) ?? wb.version;
  const key = `${ref.r1},${ref.c1},${ref.r2},${ref.c2}:${names ? names.join('\u0001') : ''}`;
  let memo = blockCubes.get(b);
  if (!memo) { memo = new Map(); blockCubes.set(b, memo); }
  const n = d2 - d1 + 1;
  const a = d1 - b.r0;
  // 블록 전체를 쓰면 행 순서와 무관(집계는 순서가 상관없음) → 정렬해도 큐브를 다시 만들지 않음
  const whole = a === 0 && n === b.n;
  const hit = memo.get(key);
  if (hit && hit.dver === (b.dver ?? 0) && (whole || hit.ver === (b.ver ?? 0)) && (!hit.mixed || hit.sv === sv) && hit.head === headKey(wb, si, ref, names)) return hit.cube;
  const width = ref.c2 - ref.c1 + 1;
  const header = names ?? Array.from({ length: width }, (_, j) => {
    const v = wb.getValue(si, ref.r1, ref.c1 + j);
    return v === null || v === '' || v === undefined ? `열${j + 1}` : String(v);
  });
  // 블록 안에 일반 셀(수식 등)이 있는 열
  const mixed = new Set();
  for (const [r,c,,count] of sheet.cells.storageEntries()) {
    // 서식 전용 빈 셀도 블록 값을 덮어쓰므로 겹치는 구간은 반드시 포함한다.
    if (r + count - 1 < d1 || r > d2) continue;
    mixed.add(c);
  }
  const cube = new Cube(n, header, (j) => {
    const c = ref.c1 + j;
    const bc = c >= b.c0 && c < b.c0 + b.cols.length && !mixed.has(c) ? b.cols[c - b.c0] : null;
    if (!bc) return new Column(n, (i) => wb.getValue(si, d1 + i, c));
    if (b.perm && !whole) { const lc = logicalCol(b, c - b.c0, a, n); return blockColumn(lc, lc.a, n); }
    return blockColumn(bc, a, n);
  }, (i) => header.map((_, j) => cube.col(j).get(i)));
  memo.set(key, { cube, sv, ver: b.ver ?? 0, dver: b.dver ?? 0, mixed: mixed.size > 0, head: headKey(wb, si, ref, names) });
  return cube;
}

// 같은 원본을 여러 피벗 · 슬라이서가 읽으므로 통합 문서가 바뀌기 전까지(wb.version 이 같으면) 재사용
const readCache = new WeakMap();
function cachedRead(wb, si, ref) {
  const read = () => {
    // 열 단위 빠른 읽기 (칸마다 getValue 보다 몇 배 빠름: 수백만 칸 원본)
    if (wb.rangeRead) return wb.rangeRead(si, ref.r1, ref.c1, ref.r2, ref.c2);
    const rows = [];
    for (let r = ref.r1; r <= ref.r2; r++) {
      const row = new Array(ref.c2 - ref.c1 + 1);
      for (let c = ref.c1; c <= ref.c2; c++) row[c - ref.c1] = wb.getValue(si, r, c);
      rows.push(row);
    }
    return rows;
  };
  if (wb.version === undefined) return read();
  // 시트 버전은 그 시트나 그 시트가 참조하는 시트가 바뀔 때만 올라감 (수식이 있어도 안전)
  const key = `${si}:${ref.r1},${ref.c1},${ref.r2},${ref.c2}`;
  let map = readCache.get(wb);
  if (!map) { map = new Map(); readCache.set(wb, map); }
  const e = map.get(key);
  const sv = wb.sheetVersion?.(si) ?? wb.version;
  if (e && e.sv === sv) return e.rows;
  const rows = read();
  map.set(key, { rows, sv });
  return rows;
}

// ───────────── 레이블 · 값 · 상위 N 필터 ─────────────
const collator = new Intl.Collator('en', { sensitivity: 'accent' }); // 엑셀 정렬 순서: 숫자 < 영문 < 한글, 대소문자 무시 (수식 비교 순서와 다름)
const wildRe = (p) => new RegExp(`^${String(p).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?')}$`, 'i');
function compareOp(op, a, v1, v2, text) {
  const cmp = (x, y) => (text ? collator.compare(String(x).toLowerCase(), String(y).toLowerCase()) : x - y);
  const s = String(a).toLowerCase();
  const t = String(v1 ?? '').toLowerCase();
  switch (op) {
    case 'equal': return text ? wildRe(v1).test(a) : cmp(a, v1) === 0;
    case 'notEqual': return text ? !wildRe(v1).test(a) : cmp(a, v1) !== 0;
    case 'beginsWith': return s.startsWith(t);
    case 'notBeginsWith': return !s.startsWith(t);
    case 'endsWith': return s.endsWith(t);
    case 'notEndsWith': return !s.endsWith(t);
    case 'contains': return s.includes(t);
    case 'notContains': return !s.includes(t);
    case 'greaterThan': return cmp(a, v1) > 0;
    case 'greaterThanOrEqual': return cmp(a, v1) >= 0;
    case 'lessThan': return cmp(a, v1) < 0;
    case 'lessThanOrEqual': return cmp(a, v1) <= 0;
    case 'between': return cmp(a, v1) >= 0 && cmp(a, v2) <= 0;
    case 'notBetween': return cmp(a, v1) < 0 || cmp(a, v2) > 0;
    default: return true;
  }
}
export const LABEL_OPS = [
  ['equal', '같음'], ['notEqual', '같지 않음'], ['beginsWith', '시작 문자'], ['notBeginsWith', '제외할 시작 문자'], ['endsWith', '끝 문자'], ['notEndsWith', '제외할 끝 문자'],
  ['contains', '포함'], ['notContains', '포함하지 않음'], ['greaterThan', '보다 큼'], ['greaterThanOrEqual', '크거나 같음'], ['lessThan', '보다 작음'], ['lessThanOrEqual', '작거나 같음'],
  ['between', '해당 범위'], ['notBetween', '제외 범위'],
];
export const VALUE_OPS = LABEL_OPS.filter(([id]) => !/With|ontains/.test(id));
/** 날짜 필터 (엑셀 피벗 [날짜 필터]: 동적 기간 · 해당 기간의 모든 날짜 · 사용자 지정) — 엑셀 pivotFilter type 이름 그대로 */
export const DATE_OPS = [
  ['dateEqual', '같음'], ['dateOlderThan', '이전'], ['dateNewerThan', '이후'], ['dateBetween', '해당 범위'], null,
  ['tomorrow', '내일'], ['today', '오늘'], ['yesterday', '어제'], null,
  ['nextWeek', '다음 주'], ['thisWeek', '이번 주'], ['lastWeek', '지난 주'], null,
  ['nextMonth', '다음 달'], ['thisMonth', '이번 달'], ['lastMonth', '지난 달'], null,
  ['nextQuarter', '다음 분기'], ['thisQuarter', '이번 분기'], ['lastQuarter', '지난 분기'], null,
  ['nextYear', '내년'], ['thisYear', '올해'], ['lastYear', '작년'], null,
  ['yearToDate', '연간 누계'],
];
export const PIVOT_DATE_PERIODS = [['Q1', '1분기'], ['Q2', '2분기'], ['Q3', '3분기'], ['Q4', '4분기'], ...Array.from({ length: 12 }, (_, i) => [`M${i + 1}`, `${i + 1}월`])];
export const DATE_OP_TYPES = new Set([...DATE_OPS.filter(Boolean).map(([k]) => k), ...PIVOT_DATE_PERIODS.map(([k]) => k), 'dateNotEqual', 'dateOlderThanOrEqual', 'dateNewerThanOrEqual', 'dateNotBetween']);
/** 오늘 일련번호 (로컬 날짜) */
export function todaySerial(now = new Date(), date1904 = false) { return serialOf(now.getFullYear(), now.getMonth() + 1, now.getDate(), date1904); }
/** 날짜 필터 조건 (serial: 항목 날짜 일련번호, today: 오늘) */
export function dateFilterMatch(op, serial, v1, v2, today, date1904 = false) {
  today ??= todaySerial(new Date(), date1904);
  if (typeof serial !== 'number' || !Number.isFinite(serial)) return false;
  const day = Math.floor(serial);
  const P = dateParts(day, 1, date1904);
  const T = dateParts(today, 1, date1904);
  const q = (m) => Math.floor((m - 1) / 3);
  const ym = (x) => x.y * 12 + x.m - 1;
  const yq = (x) => x.y * 4 + q(x.m);
  const week = (s) => s - dateParts(s, 1, date1904).dow; // 일요일 시작
  const num = (v) => (typeof v === 'number' ? v : Number(v));
  switch (op) {
    case 'today': return day === today;
    case 'yesterday': return day === today - 1;
    case 'tomorrow': return day === today + 1;
    case 'thisWeek': return week(day) === week(today);
    case 'lastWeek': return week(day) === week(today) - 7;
    case 'nextWeek': return week(day) === week(today) + 7;
    case 'thisMonth': return ym(P) === ym(T);
    case 'lastMonth': return ym(P) === ym(T) - 1;
    case 'nextMonth': return ym(P) === ym(T) + 1;
    case 'thisQuarter': return yq(P) === yq(T);
    case 'lastQuarter': return yq(P) === yq(T) - 1;
    case 'nextQuarter': return yq(P) === yq(T) + 1;
    case 'thisYear': return P.y === T.y;
    case 'lastYear': return P.y === T.y - 1;
    case 'nextYear': return P.y === T.y + 1;
    case 'yearToDate': return P.y === T.y && day <= today;
    case 'dateEqual': return day === Math.floor(num(v1));
    case 'dateNotEqual': return day !== Math.floor(num(v1));
    case 'dateOlderThan': return day < Math.floor(num(v1));
    case 'dateOlderThanOrEqual': return day <= Math.floor(num(v1));
    case 'dateNewerThan': return day > Math.floor(num(v1));
    case 'dateNewerThanOrEqual': return day >= Math.floor(num(v1));
    case 'dateBetween': return day >= Math.floor(num(v1)) && day <= Math.floor(num(v2));
    case 'dateNotBetween': return day < Math.floor(num(v1)) || day > Math.floor(num(v2));
    default: {
      const m = /^([QM])(\d+)$/.exec(op ?? '');
      if (m) return m[1] === 'Q' ? q(P.m) + 1 === Number(m[2]) : P.m === Number(m[2]);
      return true;
    }
  }
}

/** 필터 설명 (메뉴 · 필드 창) */
export function describeFieldFilter(f, values, date1904 = false) {
  if (!f) return '';
  const vname = () => { const v = values[valueIndex(values, f.by)]; return v ? valueName(v) : ''; };
  if (f.type === 'top') return `${f.top === false ? '하위' : '상위'} ${f.n}${f.mode === 'percent' ? '%' : f.mode === 'sum' ? ' (합계)' : '개'} · ${vname()}`;
  if (f.type === 'date') {
    const lab = [...DATE_OPS.filter(Boolean), ...PIVOT_DATE_PERIODS].find(([k]) => k === f.op)?.[1] ?? f.op;
    const dt = (v) => { const p = dateParts(Number(v), 1, date1904); return `${p.y}-${String(p.m).padStart(2, '0')}-${String(p.d).padStart(2, '0')}`; };
    return /^date/.test(f.op) ? `날짜 ${lab} ${dt(f.v1)}${f.op === 'dateBetween' || f.op === 'dateNotBetween' ? ` ~ ${dt(f.v2)}` : ''}` : `날짜: ${lab}`;
  }
  const op = LABEL_OPS.find(([id]) => id === f.op)?.[1] ?? f.op;
  const rng = f.op === 'between' || f.op === 'notBetween' ? `${f.v1} ~ ${f.v2}` : f.v1;
  return f.type === 'value' ? `${vname()} ${op} ${rng}` : `레이블 ${op} ${rng}`;
}

/** 값 필드 번호 (숫자 또는 이름) */
function valueIndex(values, by) {
  if (typeof by === 'number') return by >= 0 && by < values.length ? by : 0;
  const i = values.findIndex((v) => valueName(v).toLowerCase() === String(by ?? '').toLowerCase() || v.field.toLowerCase() === String(by ?? '').toLowerCase());
  return i >= 0 ? i : 0;
}

/** 행·열 필드 필터(레이블 · 값 · 상위 N)를 그룹에 적용 (바깥 필드부터, 부모 그룹 안에서 평가) */
/** 필터 상태를 비교할 글자 (필드 · 항목 순서와 무관) */
export function pivotFilterKey(filters) {
  return JSON.stringify(Object.entries(filters ?? {}).map(([k, v]) => [k.toLowerCase(), [...new Set([...v].map(x => itemIdentity(String(x))))].sort()]).sort((a, b) => (a[0] < b[0] ? -1 : 1)));
}
function applyFieldFilters(groups, d, measures) {
  const entries = Object.entries(d.fieldFilters ?? {});
  if (!entries.length) return groups;
  const measureOf = (it, vi) => {
    if (!d.values.length) return it.n;
    const v = measures.value(it.list, vi);
    return typeof v === 'number' ? v : null;
  };
  let out = groups;
  for (const [ax, axis] of [['r', d.rows], ['c', d.cols]]) {
    axis.forEach((field, level) => {
      const flt = d.fieldFilters[field];
      if (!flt) return;
      const pkOf = (keys) => {
        let p = '';
        for (let i = 0; i < level; i++) p += `\u0001${kk(keys[i])}`;
        return p;
      };
      const parents = new Map();
      for (const g of out) {
        const keys = g[ax];
        const pk = pkOf(keys);
        let m = parents.get(pk);
        if (!m) { m = new Map(); parents.set(pk, m); }
        const k = kk(keys[level]);
        let it = m.get(k);
        if (!it) { it = { key: keys[level], list: measures.newList(), n: 0 }; m.set(k, it); }
        mergeList(it.list, g.list);
        it.n += g.n;
      }
      const keep = new Set();
      for (const [pk, m] of parents) {
        const items = [...m.values()];
        let kept;
        if (flt.type === 'label') {
          const numeric = items.every((it) => typeof it.key === 'number') && Number.isFinite(Number(flt.v1));
          kept = items.filter((it) => compareOp(flt.op, numeric ? it.key : itemText(it.key), numeric ? Number(flt.v1) : flt.v1, numeric ? Number(flt.v2) : flt.v2, !numeric));
        } else if (flt.type === 'date') {
          // 날짜 항목만 남김 (글자 · 빈 항목은 날짜 필터에서 빠짐 — 엑셀과 같음)
          const today = d.today ?? todaySerial(new Date(), d.date1904);
          kept = items.filter((it) => dateFilterMatch(flt.op, it.key, flt.v1, flt.v2, today, d.date1904));
        } else if (flt.type === 'value') {
          const vi = valueIndex(d.values, flt.by);
          // 빈 값은 0 으로 비교 (엑셀: '값 = 0' 필터에 빈 항목도 남음)
          kept = items.filter((it) => compareOp(flt.op, measureOf(it, vi) ?? 0, Number(flt.v1), Number(flt.v2), false));
        } else if (flt.type === 'top' && snapshotItems(d, field)) {
          const snap = new Set(d.tieOrder[field].map(itemIdentity));
          kept = items.filter((it) => snap.has(itemIdentity(itemText(it.key))));
        } else if (flt.type === 'top') {
          const vi = valueIndex(d.values, flt.by);
          const scored = items.map((it) => ({ it, v: measureOf(it, vi) ?? -Infinity }));
          const tie = tieRank(d, field);
          // 오류 · 빈 값(-Infinity)끼리 빼면 NaN 이 되므로 같음(0)으로
          const diff = (a, b) => (a === b ? 0 : a - b);
          scored.sort((x, y) => (flt.top === false ? diff(x.v, y.v) : diff(y.v, x.v)) || tie(x.it.key) - tie(y.it.key));
          const n = Number(flt.n) || 10;
          if ((flt.mode ?? 'count') === 'count') {
            // 엑셀: N 번째와 값이 같은 항목도 모두 포함 (상위 10 인데 11개 이상 보일 수 있음)
            let m = Math.max(0, Math.floor(n));
            // 오류 · 빈 값 항목끼리도 같은 값으로 봄 (숫자 항목이 N 개보다 적으면 오류 항목이 모두 보임 — 엑셀과 같음)
            if (m > 0) while (m < scored.length && scored[m].v === scored[m - 1].v) m++;
            kept = scored.slice(0, m).map((x) => x.it);
          }
          else {
            const total = scored.reduce((acc, x) => acc + (Number.isFinite(x.v) ? x.v : 0), 0);
            const limit = flt.mode === 'percent' ? (total * n) / 100 : n;
            kept = [];
            let acc = 0;
            for (const x of scored) { if (acc >= limit && kept.length) break; kept.push(x.it); acc += Number.isFinite(x.v) ? x.v : 0; }
          }
        } else kept = items;
        for (const it of kept) keep.add(`${pk}\u0003${kk(it.key)}`);
      }
      out = out.filter((g) => keep.has(`${pkOf(g[ax])}\u0003${kk(g[ax][level])}`));
    });
  }
  return out;
}

/**
 * 같은 원본을 쓰는 피벗들(defs)과 슬라이서 필드(extraFields)를 보고 요약 캐시를 한 번에 준비.
 * 큰 파일을 열 때 · 피벗을 여러 개 새로 그리기 전에 부르면 피벗마다 원본을 다시 훑지 않음
 */
export function warmPivots(wb, defs, extraFields = []) {
  const bySrc = new Map();
  for (const def of defs) {
    const src = pivotSourceData(wb, def);
    if (!src) continue;
    const e = bySrc.get(src.cube) ?? { cube: src.cube, defs: [] };
    e.defs.push(dateDef(def, wb.date1904));
    bySrc.set(src.cube, e);
  }
  for (const { cube, defs: list } of bySrc.values()) {
    const lower = cube.header.map((h) => h.toLowerCase());
    const idx = (n) => lower.indexOf(String(n).toLowerCase());
    const dims = new Map();
    const measures = new Map();
    const addDim = (j, spec) => { if (j < 0) return; const c = spec ? groupedColumn(cube, j, spec) : j; dims.set(typeof c === 'number' ? `b${c}` : c.rollupKey, c); };
    for (const def of list) {
      const d = normalizeDef(def, cube.header);
      for (const f of [...d.rows, ...d.cols]) addDim(idx(f), d.groups?.[f]);
      for (const f of [...d.pages, ...Object.keys(d.filters)]) addDim(idx(f));
      const m = makeMeasures(d.header, d.values, d.calcFields);
      m.cols.forEach((c, i) => {
        const col = c >= 0 && c < cube.header.length ? c : -1;
        const cur = measures.get(col) ?? { col, need: {} };
        for (const k of ['sq', 'mm', 'prod']) if (m.needs[i][k]) cur.need[k] = true;
        measures.set(col, cur);
      });
    }
    for (const f of extraFields) addDim(idx(f));
    measures.set(-1, measures.get(-1) ?? { col: -1, need: {} });
    planRollup(cube, [...dims.values()], [...measures.values()]);
  }
}

/**
 * 큐브 · 선택 행 → 가장 잘게 나눈 그룹 [{ r: [행 필드 키], c: [열 필드 키], list: [슬롯별 누적], n: 행 수 }]
 * 원본을 한 번만 훑어 형식화 배열에 누적하므로 행이 많아도 빠름. 같은 선택 · 같은 필드면 재사용
 */
const groupMemo = new WeakMap();
function cubeGroups(cube, filters, d, measures) {
  const lower = d.header.map((h) => String(h).toLowerCase());
  const idx = (n) => lower.indexOf(String(n).toLowerCase());
  const rowIdx = d.rows.map(idx);
  const colIdx = d.cols.map(idx);
  const valid = (j) => j >= 0 && j < cube.header.length;
  // 그룹화한 필드(날짜 → 월 등)는 파생 열
  const specOf = (j) => d.groups?.[cube.header[j]] ?? null;
  const dimOf = (j) => (specOf(j) ? groupedColumn(cube, j, specOf(j)) : j);
  const dims = [...rowIdx, ...colIdx].filter(valid).map(dimOf);
  const slots = measures.cols.map((c) => (valid(c) ? c : -1));
  const key = JSON.stringify([rowIdx.length, [...rowIdx, ...colIdx].filter(valid).map((j) => [j, specOf(j)]), slots, measures.needs]);
  const fkey = JSON.stringify(filters.map(([j, set]) => [j, [...set].sort()]));
  let memo = groupMemo.get(cube);
  if (!memo) { memo = new Map(); groupMemo.set(cube, memo); }
  const mkey = `${fkey}\u0002${key}`;
  let groups = memo.get(mkey);
  if (groups) return groups;
  // 큰 원본은 요약 캐시(롤업)에서, 작으면 원본을 바로 (aggregateQuery 가 고름)
  const agg = aggregateQuery(cube, filters, dims, [...slots.map((c, i) => ({ col: c, need: measures.needs[i] })), { col: -1 }]);
  const keysOf = dims.map((j) => (typeof j === 'number' ? cube.col(j) : j).dim().keys);
  const S = slots.length;
  groups = new Array(agg.G);
  for (let g = 0; g < agg.G; g++) {
    let p = 0;
    const r = rowIdx.map((j) => (valid(j) ? keysOf[p][agg.codes[g * agg.D + p++]] : EMPTY));
    const c = colIdx.map((j) => (valid(j) ? keysOf[p][agg.codes[g * agg.D + p++]] : EMPTY));
    const list = new Array(S);
    for (let m = 0; m < S; m++) list[m] = groupAcc(agg, m, g);
    groups[g] = { r, c, list, n: agg.stats[S].count[g] };
  }
  if (memo.size > 80) memo.clear();
  memo.set(mkey, groups);
  return groups;
}

/** 필터(보고서 필터 · 항목 선택 · 슬라이서 · 레이블/값/상위 N)를 적용한 행 → { def: 정리된 정의, rows, header } */
// 같은 원본(행 배열)과 같은 정의면 결과를 재사용 (슬라이서 · 피벗 차트가 같은 피벗을 여러 번 계산하지 않게)
const resolveMemo = new WeakMap();
const nonEmptyMemo = new WeakMap();
const filterMemo = new WeakMap();
/** 결과에 영향을 주는 정의 부분만으로 만든 키 (위치 · 서식 제외) */
export function pivotDefKey(def) {
  const { area, top, left, cellFmt, captureFmt, buttons, styleDef, autofit, name, ...rest } = def;
  return JSON.stringify(rest);
}
/**
 * 파생 그룹 필드 (엑셀의 '월2' = '일' 필드를 월로 묶은 새 필드): 원본 열을 새 이름으로 한 번 더 보여 주는 큐브.
 * 그룹화는 보통 필드처럼 def.groups[새 이름] 으로 적용됨
 */
const derivedMemo = new WeakMap();
export function withDerivedFields(cube, groups) {
  const lower = cube.header.map((h) => h.toLowerCase());
  const add = Object.entries(groups ?? {}).filter(([n, g]) => g?.base && !lower.includes(n.toLowerCase()) && lower.includes(String(g.base).toLowerCase()));
  if (!add.length) return cube;
  const key = JSON.stringify(add.map(([n, g]) => [n, g.base]));
  let m = derivedMemo.get(cube);
  if (!m) { m = new Map(); derivedMemo.set(cube, m); }
  let c = m.get(key);
  if (!c) {
    const H = cube.header.length;
    const bases = add.map(([, g]) => lower.indexOf(String(g.base).toLowerCase()));
    c = new Cube(cube.n, [...cube.header, ...add.map(([n]) => n)], (j) => cube.col(j < H ? j : bases[j - H]), (i) => { const r = cube.row(i); return [...r, ...bases.map((b) => r[b])]; });
    m.set(key, c);
  }
  return c;
}

// ───────────── 계산 항목 (엑셀 피벗 › 필드, 항목 및 집합 › 계산 항목) ─────────────
/**
 * 계산 항목 수식: 같은 필드의 항목에 계수를 곱해 더하고 빼는 식 (예: 서울+부산, '서울 강남'*0.5-인천).
 * [{ item, k }] 를 돌려주고, 곱셈 · 나눗셈이 항목끼리이거나 상수만 있는 항은 오류 (Error, 한국어 메시지)
 */
export function parseCalcItem(formula) {
  const src = String(formula ?? '').trim().replace(/^=/, '');
  if (!src) throw new Error('수식을 입력하세요.');
  const terms = [];
  let i = 0;
  let sign = 1;
  const ws = () => { while (src[i] === ' ') i++; };
  const atom = () => {
    ws();
    if (src[i] === "'") {
      let t = '';
      i++;
      while (i < src.length) { if (src[i] === "'" && src[i + 1] === "'") { t += "'"; i += 2; } else if (src[i] === "'") { i++; break; } else t += src[i++]; }
      return { item: t };
    }
    const m = /^(\d+(?:\.\d+)?(?:e[+-]?\d+)?)/i.exec(src.slice(i));
    if (m) { i += m[1].length; return { num: Number(m[1]) }; }
    let t = '';
    while (i < src.length && !/[+\-*/()]/.test(src[i])) t += src[i++];
    t = t.trim();
    if (!t) throw new Error(`수식에 항목 이름이 필요합니다: ${src}`);
    return { item: t };
  };
  while (i < src.length) {
    ws();
    if (src[i] === '+') { i++; continue; }
    if (src[i] === '-') { sign = -sign; i++; continue; }
    let k = sign;
    let item = null;
    for (;;) {
      const a = atom();
      if (a.item != null) { if (item != null) throw new Error('항목끼리 곱하거나 나눌 수는 없습니다.'); item = a.item; } else k *= a.num;
      ws();
      if (src[i] === '*') { i++; continue; }
      if (src[i] === '/') {
        i++;
        const b = atom();
        if (b.item != null || !b.num) throw new Error('항목은 0 이 아닌 숫자로만 나눌 수 있습니다.');
        k /= b.num;
        ws();
        if (src[i] === '*') { i++; continue; }
      }
      break;
    }
    if (item == null) throw new Error('상수만 있는 항은 쓸 수 없습니다. 항목에 곱하세요 (예: 서울*1.1).');
    terms.push({ item, k });
    sign = 1;
    ws();
    if (i < src.length && src[i] !== '+' && src[i] !== '-') throw new Error(`수식을 이해할 수 없습니다: ${src.slice(i)}`);
  }
  if (!terms.length) throw new Error('수식에 항목이 없습니다.');
  return terms;
}

const calcItemMemo = new WeakMap();
/**
 * 계산 항목을 원본 레코드로 펼친 큐브: 항목 A 의 레코드마다 필드 값을 계산 항목 이름으로 바꾸고 값 필드(measureFields)에 계수를 곱한 레코드를 더함.
 * 다른 필드 · 필터 · 부분합 · 총합계(엑셀처럼 계산 항목 포함)가 그대로 동작. calcItems = { 필드: [{ name, formula }] }
 */
export function withCalcItems(cube, calcItems, measureFields = []) {
  const list = Object.entries(calcItems ?? {}).flatMap(([field, items]) => (items ?? []).map((it) => ({ field, ...it })));
  if (!list.length) return cube;
  const key = JSON.stringify([list, measureFields]);
  let m = calcItemMemo.get(cube);
  if (!m) { m = new Map(); calcItemMemo.set(cube, m); }
  const hit = m.get(key);
  if (hit) return hit;
  const lower = cube.header.map((h) => h.toLowerCase());
  const meas = new Set(measureFields.map((f) => lower.indexOf(String(f).toLowerCase())).filter((j) => j >= 0));
  const extra = []; // [원본 행, 필드 열, 이름, 계수]
  for (const ci of list) {
    const j = lower.indexOf(String(ci.field).toLowerCase());
    if (j < 0) continue;
    let terms;
    try { terms = parseCalcItem(ci.formula); } catch { continue; }
    const { codes, keys } = cube.col(j).dim();
    const coef = new Float64Array(keys.length);
    const want = new Map(terms.map((t) => [t.item.toLowerCase(), 0]));
    for (const t of terms) want.set(t.item.toLowerCase(), (want.get(t.item.toLowerCase()) ?? 0) + t.k);
    keys.forEach((k, c) => { const w = want.get(itemText(k).toLowerCase()); if (w) coef[c] = w; });
    for (let i = 0; i < cube.n; i++) { const w = coef[codes[i]]; if (w) extra.push([i, j, ci.name, w]); }
  }
  const n = cube.n;
  const synth = (x) => {
    const [i, j, name, w] = extra[x];
    const r = [...cube.row(i)];
    r[j] = name;
    for (const q of meas) if (q !== j && typeof r[q] === 'number') r[q] *= w;
    return r;
  };
  const rowAt = (i) => (i < n ? cube.row(i) : synth(i - n));
  const c = new Cube(n + extra.length, cube.header, (jj) => { const base = cube.col(jj); return new Column(n + extra.length, (i) => (i < n ? base.get(i) : synth(i - n)[jj] ?? null)); }, rowAt);
  if (m.size > 20) m.clear();
  m.set(key, c);
  return c;
}

export function resolvePivot(input, def) {
  def = dateDef(def, input.date1904 ?? def.date1904);
  // input: 행 배열(머리글 포함) 또는 pivotSourceData 결과({ cube })
  let cube = withDerivedFields(Array.isArray(input) ? cubeFromRows(input) : input.cube, def.groups);
  if (def.calcItems) cube = withCalcItems(cube, def.calcItems, (def.values ?? []).map((v) => v.field));
  let memo = resolveMemo.get(cube);
  if (!memo) { memo = new Map(); resolveMemo.set(cube, memo); }
  const key = pivotDefKey(def);
  const hit = memo.get(key);
  // 캐시된 결과를 펼치면(…) 게터가 불려 천만 행을 만들 수 있으므로 필드를 하나씩 복사
  if (hit) return withRows({ def: normalizeDef(def, hit.header), header: hit.header, cube: hit.cube, filters: hit.filters, groups: hit.groups, measures: hit.measures });
  const d = normalizeDef(def, cube.header);
  const header = d.header;
  const lower = cube.header.map((h) => h.toLowerCase());
  // 보고서 필터 · 항목 선택 · 슬라이서: 코드 단위로 한 번에 (같은 필터를 쓰는 피벗끼리 결과 공유)
  const filters = Object.entries(d.filters).map(([name, allowed]) => {
    const j = lower.indexOf(name.toLowerCase());
    const spec = j >= 0 ? d.groups?.[cube.header[j]] : null;
    if (!spec) return [j, new Set(allowed)];
    // 그룹화한 필드('1월' · '12월9일' 등)의 필터 → 그 그룹에 드는 원래 값들
    const want = new Set(allowed.map(itemIdentity));
    const raw = new Set();
    for (const k of cube.col(j).dim().keys) if (want.has(itemIdentity(itemText(k === EMPTY ? EMPTY : groupKey(k, spec))))) raw.add(itemText(k));
    return [j, raw];
  }).filter(([i]) => i >= 0);
  const measures = makeMeasures(header, d.values, d.calcFields);
  const groups = applyFieldFilters(cubeGroups(cube, filters, d, measures), d, measures);
  const res = withRows({ def: d, header, cube, filters, groups, measures });
  if (memo.size > 60) memo.clear();
  memo.set(key, res);
  return res;
}

/** 걸러진 원본 행(머리글 포함)을 필요할 때만 만드는 rows 속성 (열거되지 않음 → 펼쳐도 안 만들어짐) */
function withRows(res) {
  let made = null;
  Object.defineProperty(res, 'rows', {
    enumerable: false,
    get() {
      if (made) return made;
      const { cube } = res;
      const sel = filterRows(cube, res.filters ?? []);
      const n = sel ? sel.length : cube.n;
      made = new Array(n + 1);
      made[0] = cube.header;
      for (let i = 0; i < n; i++) made[i + 1] = cube.row(sel ? sel[i] : i);
      return made;
    },
  });
  return res;
}

// ───────────── 피벗 스타일 (엑셀 기본 제공 이름: 밝게 1~28, 보통 1~28, 어둡게 1~28) ─────────────
export const DEFAULT_PIVOT_STYLE = 'PivotStyleLight16';
/** 스타일 이름 → 역할별 서식 { header, sub, grand, body, page } */
export function pivotStyleParts(name, custom = null) {
  // 파일에서 가져온 사용자 지정 스타일
  if (name !== 'None' && custom) return { header: {}, sub: {}, grand: {}, body: {}, page: {}, band: {}, ...custom };
  const p = styleParts(name);
  const m = /^PivotStyle(?:Light|Medium|Dark)(\d+)$/i.exec(name ?? '');
  const a = ACCENTS[m ? (Number(m[1]) - 1) % 7 : 1];
  return { ...p, band: { fill: tint(a.mid ?? a.hex, 0.85) } };
}
function styleParts(name) {
  const m = /^PivotStyle(Light|Medium|Dark)(\d+)$/i.exec(name ?? '');
  if (!m) return name === 'None' ? { header: {}, sub: { bold: true }, grand: { bold: true }, body: {}, page: {} } : styleParts(DEFAULT_PIVOT_STYLE);
  const kind = m[1].toLowerCase();
  const n = Math.max(1, Math.min(28, Number(m[2])));
  const a = ACCENTS[(n - 1) % 7];
  const g = Math.floor((n - 1) / 7);
  const base = a.hex;
  const soft = a.mid ?? a.hex;
  const t = (x) => tint(soft, x);
  const W = '#ffffff';
  if (kind === 'light') {
    return [
      { header: { bold: true, bb: true }, sub: { bold: true }, grand: { bold: true, bt: true }, body: {}, page: { bold: true } },
      { header: { bold: true, fill: base, color: W }, sub: { bold: true }, grand: { bold: true, bt: true }, body: {}, page: { bold: true } },
      { header: { bold: true, fill: t(0.8), bb: true }, sub: { bold: true }, grand: { bold: true, fill: t(0.8), bt: true }, body: {}, page: { fill: t(0.8) } },
      { header: { bold: true, fill: t(0.6) }, sub: { bold: true, fill: t(0.8) }, grand: { bold: true, fill: t(0.6) }, body: {}, page: { fill: t(0.8) } },
    ][g];
  }
  if (kind === 'medium') {
    const dark = (n - 1) % 7 === 0 ? '#404040' : shade(base, 0.25);
    return [
      { header: { bold: true, fill: base, color: W }, sub: { bold: true, fill: t(0.8) }, grand: { bold: true, fill: t(0.6), bt: true }, body: {}, page: { fill: t(0.8) } },
      { header: { bold: true, fill: dark, color: W }, sub: { bold: true, fill: t(0.6) }, grand: { bold: true, fill: dark, color: W }, body: {}, page: { fill: t(0.8) } },
      { header: { bold: true, fill: t(0.4) }, sub: { bold: true, fill: t(0.8) }, grand: { bold: true, fill: t(0.4) }, body: {}, page: { fill: t(0.8) } },
      { header: { bold: true, fill: t(0.6), bb: true }, sub: { bold: true, fill: t(0.8) }, grand: { bold: true, fill: t(0.6), bt: true }, body: { fill: t(0.9) }, page: { fill: t(0.8) } },
    ][g];
  }
  const dk = (n - 1) % 7 === 0 ? '#000000' : shade(base, 0.5);
  return [
    { header: { bold: true, fill: dk, color: W }, sub: { bold: true, fill: shade(soft, 0.25), color: W }, grand: { bold: true, fill: dk, color: W }, body: { fill: t(0.6) }, page: { fill: t(0.6) } },
    { header: { bold: true, fill: base, color: W }, sub: { bold: true, fill: t(0.4) }, grand: { bold: true, fill: base, color: W }, body: { fill: t(0.8) }, page: { fill: t(0.8) } },
    { header: { bold: true, fill: '#404040', color: W }, sub: { bold: true, fill: t(0.6) }, grand: { bold: true, fill: '#404040', color: W }, body: { fill: t(0.8) }, page: { fill: t(0.8) } },
    { header: { bold: true, fill: dk, color: W, bb: true }, sub: { bold: true, fill: t(0.6) }, grand: { bold: true, fill: dk, color: W }, body: {}, page: { fill: t(0.8) } },
  ][g];
}
/** 스타일 갤러리용 목록 */
export const PIVOT_STYLES = [
  ...MODERN_STYLES.filter((s) => s.pivot).map((s) => ({ ...s, get swatch() { return presetSwatch(s.name); } })),
  ...['Light', 'Medium', 'Dark'].flatMap((k, gi) => Array.from({ length: 28 }, (_, i) => {
    const name = `PivotStyle${k}${i + 1}`;
    const grp = ['밝게', '보통', '어둡게'][gi];
    return { name, group: grp, label: `피벗 스타일 ${grp} ${i + 1}`, get swatch() { return presetSwatch(name); } };
  })),
];
export const PIVOT_STYLE_GROUPS = [...new Set(PIVOT_STYLES.map((s) => s.group))];

/** 셀 역할 → 스타일 부분 */
export function roleStyle(parts, role) {
  const kind = role.split(':')[0];
  switch (kind) {
    case 'corner': case 'rowHead': case 'colHead': case 'valueCaption': case 'colItem': case 'valueHead': case 'colSubHead': case 'grandHead': return parts.header;
    case 'rowGroup': case 'groupData': case 'rowSub': case 'subData': return parts.sub;
    case 'colSubData': return { bold: true };
    case 'grandLabel': case 'grandData': return parts.grand;
    case 'pageLabel': case 'pageValue': return parts.page;
    case 'data': case 'rowItem': return parts.body;
    default: return {};
  }
}

// ───────────── 계산 ─────────────

const QUOTE_TEXT = /^(?:['=]|[+-]?\.?\d|(?:TRUE|FALSE)$)/i;
const labelMemo = new Map();
/** 항목 레이블 → 셀 입력 글자: 숫자 글자('2025')는 숫자 그대로, 읽으면 바뀌는 글자('001' · "'소득세율")는 글자로 */
function labelRaw(t) {
  if (!QUOTE_TEXT.test(t)) return t;
  let r = labelMemo.get(t);
  if (r === undefined) {
    const v = t[0] === "'" || t[0] === '=' ? null : parseInput(t).value;
    r = typeof v === 'number' && formatGeneral(v) === t ? t : `'${t}`;
    if (labelMemo.size < 50000) labelMemo.set(t, r);
  }
  return r;
}

/**
 * 엑셀이 저장한 상위 N 항목을 그대로 쓸지: 파일을 연 뒤 아직 다시 그리지 않았고(tieState 가 남아 있음) 필터도 저장할 때와 같을 때.
 * (값이 오류인 항목을 엑셀이 어떤 것은 넣고 어떤 것은 빼는 등 규칙으로 재현할 수 없는 결과도 연 화면은 엑셀과 같게)
 */
function snapshotItems(d, field) {
  const list = d.tieOrder?.[field];
  return !!list?.length && d.tieState !== null && d.tieState === pivotFilterKey(d.filters);
}

/** 상위 항목 글자 경로 (tieByParent 키) */
function parentTextPath(n) {
  const out = [];
  for (let p = n; p && p.depth >= 0; p = p.parent) out.push(itemText(p.key));
  return out.reverse().join('\u0001');
}

/** 값이 같은 항목의 순서 (파일에 저장된 표시 순서, 없으면 모두 같음). parent 가 있으면 그 상위 항목 아래의 저장 순서 먼저 */
function tieRank(d, field, parent = null) {
  const local = parent && d.tieByParent?.[field]?.[parentTextPath(parent)];
  if (local?.length) {
    const pos = new Map(local.map((t, i) => [itemIdentity(t), i]));
    const rest = tieRank(d, field);
    return (key) => { const i = pos.get(itemIdentity(itemText(key))); return i !== undefined ? i : local.length + rest(key); };
  }
  const list = d.tieOrder?.[field];
  if (!list?.length) return () => 0;
  const pos = new Map(list.map((t, i) => [itemIdentity(t), i]));
  return (key) => pos.get(itemIdentity(itemText(key))) ?? list.length;
}

/** 항목 순서: 정렬 설정(글자/값) → 수동 순서 → 기본(오름차순) */
function orderTree(root, fields, d, measureAt) {
  const rec = (n) => {
    if (!n.children.length) return;
    const field = fields[n.depth + 1];
    const s = d.sort?.[field];
    const order = d.order?.[field];
    let kids = n.children;
    if (s && s.by !== undefined && s.by !== null) {
      const vi = valueIndex(d.values, s.by);
      const score = new Map(kids.map((c) => [c, measureAt(c, vi, field)]));
      // 빈 값은 0 과 같은 자리 (엑셀), 오류는 맨 끝
      const num = (v) => (typeof v === 'number' ? v : v === null || v === undefined ? 0 : -Infinity);
      const tie = tieRank(d, field, n);
      kids = [...kids].sort((a, b) => (s.dir === 'desc' ? num(score.get(b)) - num(score.get(a)) : num(score.get(a)) - num(score.get(b))) || tie(a.key) - tie(b.key));
    } else if (d.groups?.[field] && d.groups[field].by !== 'items') {
      // 그룹화한 필드: 월 · 분기 · 구간은 숫자 순서
      const spec = d.groups[field];
      const rank = (k) => groupRank(k, spec);
      kids = [...kids].sort((a, b) => {
        const x = rank(a.key);
        const y = rank(b.key);
        if (x !== null && y !== null) return x - y;
        return collator.compare(String(a.key), String(b.key));
      });
      if (s?.dir === 'desc') kids.reverse();
    } else {
      const keys = sortKeys(kids.map((c) => c.key), { customList: d.customListSort !== false });
      const byKey = new Map(kids.map((c) => [kk(c.key), c]));
      kids = keys.map((k) => byKey.get(kk(k)));
      // 글자 정렬 필드인데 항목이 모두 파일에 저장된 순서에 있으면 그 순서 (엑셀이 정렬해 둔 순서: 날짜와 글자가 섞인 경우 등)
      const filePos = s && order?.length ? new Map(order.map((t, i) => [itemIdentity(t), i])) : null;
      if (filePos && kids.every((c) => filePos.has(itemIdentity(itemText(c.key))))) {
        kids = [...kids].sort((a, b) => filePos.get(itemIdentity(itemText(a.key))) - filePos.get(itemIdentity(itemText(b.key))));
        n.children = kids;
        kids.forEach(rec);
        return;
      }
      if (order?.length && !s) {
        const pos = new Map(order.map((t, i) => [itemIdentity(t), i]));
        const known = kids.filter((c) => pos.has(itemIdentity(itemText(c.key)))).sort((a, b) => pos.get(itemIdentity(itemText(a.key))) - pos.get(itemIdentity(itemText(b.key))));
        kids = [...known, ...kids.filter((c) => !pos.has(itemIdentity(itemText(c.key))))];
      }
      if (s?.dir === 'desc') {
        const blank = kids.filter((c) => c.key === EMPTY);
        kids = [...kids.filter((c) => c.key !== EMPTY).reverse(), ...blank];
      }
    }
    n.children = kids;
    kids.forEach(rec);
  };
  rec(root);
}

/**
 * 피벗 계산 → { grid: [[{raw, style, role}]], meta }
 * rows: 머리글 포함 (필터 적용 후), d: normalizeDef 결과
 */
export function computePivot(input, d) {
  // input: resolvePivot 결과(그룹 포함) 또는 이미 걸러진 행 배열(머리글 포함)
  const resolved = input && !Array.isArray(input) && input.groups ? input : null;
  const header = d.header ?? resolved?.header ?? headerNames(input);
  const idx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const rowIdx = d.rows.map(idx);
  const colIdx = d.cols.map(idx);
  const values = d.values;
  const V = values.length;
  const parts = pivotStyleParts(d.style, d.styleDef);
  const opts = d.styleOpts ?? { rowHeaders: true, colHeaders: true };
  const HEAD_ROLES = /^(corner|rowHead|colHead|valueCaption|colItem|valueHead|colSubHead|grandHead)/;
  // 엑셀 기본 제공 스타일은 정확한 정의로 표 모양대로 칠함 (paintPivotPreset) — 역할별 근사 서식은 쓰지 않음
  const preset = d.style !== 'None' && Array.isArray(d.styleElements) ? styleElementsPreset(d.styleElements) : !d.styleDef && presetStyle(d.style) ? d.style : null;
  const styleFor = (role) => {
    if (preset) return {};
    if (opts.colHeaders === false && HEAD_ROLES.test(role)) return {};
    const s = roleStyle(parts, role);
    if (opts.rowHeaders === false && /^(rowGroup|rowItem)/.test(role)) { const { bold, ...rest } = s; return rest; }
    return s;
  };
  const valIdx = values.map((v) => idx(v.field));
  const calcNames = new Set((d.calcFields ?? []).map((c) => c.name.toLowerCase()));

  // 글자 항목은 그대로: '=' · 작은따옴표로 시작하거나, 입력으로 읽으면 글자가 달라지는 것(001 · 날짜처럼 보이는 글자)은 앞에 '
  // 명시적 빈 캡션은 값이 없는 셀과 다르게 빈 문자열로 보존합니다.
  const text = (s, role, explicitEmpty = false) => ({ raw: explicitEmpty && s === '' ? "'" : typeof s === 'number' ? formatGeneral(s) : labelRaw(String(s ?? '')), style: { align: 'general', ...styleFor(role) }, role });
  // 필터만 있는 경우도 본문과 같은 항목 판정·표준 스타일 경로를 사용합니다.
  const pageLayout = pivotPageLayout(d);
  const pageRows = Array.from({ length: pageLayout.height }, () => []);
  const pageColumns = new Map();
  for (const { field: p, r, c } of pageLayout.fields) {
    const filters = d.filters ?? {}, allowed = filters[p] ?? filters[Object.keys(filters).find((k) => k.toLowerCase() === p.toLowerCase())];
    let shown = allowed, all = !allowed;
    if (allowed && resolved?.cube) {
      const j = idx(p);
      if (j >= 0 && j < resolved.cube.header.length) {
        const have = new Map(resolved.cube.col(j).texts().map(t => [itemIdentity(t), t]));
        shown = [...new Set(allowed.filter(t => have.has(itemIdentity(t))).map(t => have.get(itemIdentity(t))))];
        all = have.size > 1 && shown.length === have.size;
      }
    }
    const caption = all ? '(모두)' : shown.length === 1 ? itemProperty(d.itemCaptions?.[p], shown[0]) ?? shown[0] : '(다중 항목)';
    const pair = [{ ...text(d.fieldCaptions?.[p] ?? p, 'pageLabel'), field: p }, { ...text(caption, 'pageValue'), field: p }];
    pageRows[r][c] = pair[0]; pageRows[r][c + 1] = pair[1];
    if (!pageColumns.has(c)) pageColumns.set(c, []);
    pageColumns.get(c).push(pair);
  }
  if (preset) for (const column of pageColumns.values()) paintPivotPreset(preset, column, { pages: column.length, top: column.length, headerRows: 0, labelCols: 1, width: 2 }, opts);
  // 희소 칸은 사용자 셀: 필터 열 사이와 마지막 줄의 빈 필터 위치는 채우거나 지우지 않습니다.
  if (!d.rows.length && !d.cols.length && !V && pageRows.length) {
    return { grid: pageRows, meta: { header, rowIdx, colIdx, valIdx, values, labelCols: 1, pageRows: pageRows.length, pageFields: pageLayout.fields, pageWidth: pageLayout.width, headerRows: 0, colLeaves: [], rowItems: [], colItems: [], rowTree: null, colTree: null, width: pageLayout.width, bodyRows: 0, empty: true } };
  }
  if (!d.rows.length && !d.cols.length && !V) {
    const grid = Array.from({ length: 18 }, (_, r) => Array.from({ length: 3 }, (_, c) => ({
      raw: r === 0 && c === 0 ? '피벗 테이블 보고서를 작성하려면 [피벗 테이블 필드] 목록에서 필드를 선택하세요.' : '',
      style: { fill: '#f3f6fb', ...(r === 0 ? { bt: true } : {}), ...(r === 17 ? { bb: true } : {}), ...(c === 0 ? { bl: true } : {}), ...(c === 2 ? { br: true } : {}), ...(r === 0 && c === 0 ? { color: '#44546a' } : {}) },
      role: 'empty',
    })));
    return { grid, meta: { header, rowIdx, colIdx, valIdx, values, labelCols: 1, pageRows: 0, headerRows: 1, colLeaves: [], rowItems: [], colItems: [], rowTree: null, colTree: null, width: 3, bodyRows: 18, empty: true } };
  }

  const measures = makeMeasures(header, values, d.calcFields);
  // 가장 잘게 나눈 그룹 (원본 행은 열 기반 엔진이 한 번만 훑음)
  const groups = resolved
    ? resolved.groups
    : cubeGroups(cubeFromRows(input), [], { ...d, header }, measures);
  const newNode = (key, path, depth) => ({ key, path, depth, children: [], map: new Map() });
  const rowTree = newNode(null, '', -1);
  const colTree = newNode(null, '', -1);
  // 누적: 행 경로 접두사 × 열 경로 접두사 (그룹 수만큼만)
  const accs = new Map();
  const Lr0 = rowIdx.length;
  const Lc0 = colIdx.length;
  const rp = new Array(Lr0 + 1);
  const cp = new Array(Lc0 + 1);
  const walkPath = (tree, keys, out) => {
    let node = tree;
    out[0] = '';
    for (let i = 0; i < keys.length; i++) {
      const kkey = kk(keys[i]);
      out[i + 1] = `${out[i]}\u0001${kkey}`;
      let ch = node.map.get(kkey);
      if (!ch) { ch = newNode(keys[i], out[i + 1], i); ch.parent = node.depth >= 0 ? node : null; node.map.set(kkey, ch); node.children.push(ch); }
      node = ch;
    }
  };
  for (const g of groups) {
    walkPath(rowTree, g.r, rp);
    walkPath(colTree, g.c, cp);
    for (let a = 0; a <= Lr0; a++) {
      for (let b = 0; b <= Lc0; b++) {
        const key = `${rp[a]}\u0002${cp[b]}`;
        let list = accs.get(key);
        if (!list) { list = measures.newList(); accs.set(key, list); }
        mergeList(list, g.list);
      }
    }
  }
  const raw = (rpath, cpath, vi) => {
    const list = accs.get(`${rpath}\u0002${cpath}`);
    return list || rpath || cpath ? measures.value(list, vi) : measures.emptyValue(vi);
  };
  // 값 기준 정렬의 기준 칸: 보통은 반대 축 합계, sort.at 이 있으면 반대 축의 그 항목 (엑셀 autoSortScope 의 항목 참조)
  const pathAt = (tree, fields, at) => {
    let node = tree;
    for (const f of fields) {
      const want = at.find(([x]) => x.toLowerCase() === f.toLowerCase());
      if (!want) break;
      const next = node.children.find((ch) => itemIdentity(itemText(ch.key)) === itemIdentity(want[1]));
      if (!next) return null;
      node = next;
    }
    return node.path;
  };
  const otherPath = (field, tree, fields) => {
    const at = d.sort?.[field]?.at;
    return at?.length ? pathAt(tree, fields, at) : '';
  };
  orderTree(rowTree, d.rows, d, (node, vi, field) => { const cp = otherPath(field, colTree, d.cols); return cp === null ? null : raw(node.path, cp, vi); });
  orderTree(colTree, d.cols, d, (node, vi, field) => { const rp = otherPath(field, rowTree, d.rows); return rp === null ? null : raw(rp, node.path, vi); });
  // 축소한 항목 (필드별 항목 글자, 엑셀처럼 필드의 같은 항목은 모두 함께)
  const collSets = { r: d.rows.map((f) => new Set((d.collapsed?.[f] ?? []).map(itemIdentity))), c: d.cols.map((f) => new Set((d.collapsed?.[f] ?? []).map(itemIdentity))) };
  const isColl = (axis, node) => node.depth >= 0 && node.depth < (axis === 'r' ? d.rows.length : d.cols.length) - 1 && collSets[axis][node.depth].has(itemIdentity(itemText(node.key)));
  const toggleOf = (axis, node) => (d.showExpand && node.depth >= 0 && node.depth < (axis === 'r' ? d.rows.length : d.cols.length) - 1
    ? { field: (axis === 'r' ? d.rows : d.cols)[node.depth], item: itemText(node.key), collapsed: isColl(axis, node) } : null);
  const cellValue = (rpath, cpath, vi) => {
    if (vi < 0) return null;
    const v = raw(rpath, cpath, vi);
    const as = values[vi].showAs ?? 'normal';
    if (as === 'normal') return v;
    return showValue(as, values[vi], vi, rpath, cpath, v);
  };
  // 값 표시 형식 (엑셀과 같음): 기준 필드의 형제 항목 · 상위 항목 · 합계와 비교
  const num0 = (x) => (typeof x === 'number' ? x : 0);
  const ratio = (a, b) => (isErr(a) ? a : isErr(b) ? b : typeof b === 'number' && b ? num0(a) / b : CALC_ERR('#DIV/0!'));
  let pathIdx = null;
  const nodeOf = (axis, path) => {
    if (!pathIdx) {
      pathIdx = { r: new Map([['', rowTree]]), c: new Map([['', colTree]]) };
      const add = (m, n) => n.children.forEach((ch) => { ch.parent = n.depth >= 0 ? n : null; m.set(ch.path, ch); add(m, ch); });
      add(pathIdx.r, rowTree);
      add(pathIdx.c, colTree);
    }
    return pathIdx[axis].get(path);
  };
  const sibMemo = new Map();
  const showValue = (as, vdef, vi, rpath, cpath, v) => {
    if (as === 'percentOfTotal') return ratio(v, raw('', '', vi));
    if (as === 'percentOfRow') return ratio(v, raw(rpath, '', vi));
    if (as === 'percentOfCol') return ratio(v, raw('', cpath, vi));
    if (as === 'index') {
      const g = raw('', '', vi); const rt = raw(rpath, '', vi); const ct = raw('', cpath, vi);
      return isErr(v) ? v : isErr(rt) ? rt : isErr(ct) ? ct : isErr(g) ? g : typeof rt === 'number' && typeof ct === 'number' && rt && ct ? (num0(v) * num0(g)) / (rt * ct) : CALC_ERR('#DIV/0!');
    }
    if (as === 'percentOfParentRow' || as === 'percentOfParentCol') {
      const axis = as === 'percentOfParentRow' ? 'r' : 'c';
      if (!(axis === 'r' ? d.rows : d.cols).length) return null;
      const node = nodeOf(axis, axis === 'r' ? rpath : cpath);
      const pp = node && node.depth >= 0 ? node.parent?.path ?? '' : axis === 'r' ? rpath : cpath;
      const parent = axis === 'r' ? raw(pp, cpath, vi) : raw(rpath, pp, vi);
      return isErr(parent) ? null : ratio(v, parent);
    }
    // 기준 필드: 행 또는 열 필드 (없으면 첫 행 필드)
    const bf = vdef.baseField ?? d.rows[0] ?? d.cols[0];
    const lower = String(bf ?? '').toLowerCase();
    let axis = 'r';
    let L = d.rows.findIndex((f) => f.toLowerCase() === lower);
    if (L < 0) { axis = 'c'; L = d.cols.findIndex((f) => f.toLowerCase() === lower); }
    if (L < 0) return CALC_ERR('#N/A');
    const mine = axis === 'r' ? rpath : cpath;
    const at = (p) => (axis === 'r' ? raw(p, cpath, vi) : raw(rpath, p, vi));
    const node = nodeOf(axis, mine);
    if (!node) return null;
    if (node.depth < L) return as === 'percentOfRunningTotal' && isErr(v) ? v : null;
    let a = node;
    while (a.depth > L) a = a.parent;
    const rest = mine.slice(a.path.length);
    if (as === 'percentOfParent') {
      const parent = at(a.path);
      return typeof parent === 'number' && parent ? ratio(v, parent) : null;
    }
    const sibs = (a.parent ?? (axis === 'r' ? rowTree : colTree)).children;
    const pos = sibs.indexOf(a);
    if (as === 'percent' || as === 'difference' || as === 'percentDiff') {
      let b;
      if (vdef.basePos === 'prev' || vdef.basePos === 'next') b = sibs[pos + (vdef.basePos === 'prev' ? -1 : 1)] ?? a;
      else {
        const want = String(vdef.baseItem ?? '');
        b = sibs.find((s) => itemIdentity(itemText(s.key)) === itemIdentity(want));
        if (!b) return CALC_ERR('#N/A');
      }
      const base = at(b.path + rest);
      // Excel leaves the base item's difference blank even when its value is an error.
      if (b === a && as !== 'percent') return null;
      if (isErr(v) && b !== a) return v;
      if (as === 'difference') return isErr(base) ? base : num0(v) - num0(base);
      // An unavailable base is blank; a numeric zero is a division error.
      if (base === null || isErr(base)) return null;
      return as === 'percent' ? ratio(v, base) : ratio(num0(v) - base, base);
    }
    // 누계 · 순위: 같은 상위 항목 아래 형제들의 값 (한 번 계산해 둠)
    const mk = `${as}\u0003${axis}\u0003${(a.parent ?? { path: '' }).path}\u0003${rest}\u0003${axis === 'r' ? cpath : rpath}\u0003${vi}`;
    let arr = sibMemo.get(mk);
    if (!arr) {
      const vals = sibs.map((s) => at(s.path + rest));
      if (as === 'runTotal' || as === 'percentOfRunningTotal') {
        let acc = 0;
        // A later error replaces the preceding error, and subsequent values retain it.
        arr = vals.map((x) => (acc = isErr(x) ? x : isErr(acc) ? acc : acc + num0(x)));
        if (as === 'percentOfRunningTotal') arr = arr.map((x) => (isErr(x) ? x : typeof acc === 'number' && acc ? x / acc : CALC_ERR('#DIV/0!')));
      } else {
        const distinct = [...new Set(vals.filter((x) => typeof x === 'number'))].sort((p, q) => (as === 'rankAscending' ? p - q : q - p));
        const rank = new Map(distinct.map((x, i) => [x, i + 1]));
        arr = vals.map((x) => (typeof x === 'number' ? rank.get(x) : null));
      }
      sibMemo.set(mk, arr);
    }
    return v === null && as !== 'runTotal' && as !== 'percentOfRunningTotal' ? null : arr[pos];
  };

  // 사용자가 바꾼 이름 (엑셀: 필드 이름 · 항목 이름을 셀에서 고쳐 쓴 것 — pivotField@name, item@n)
  const fcap = (f) => d.fieldCaptions?.[f] ?? f;
  const icap = (axis, node) => {
    const t = itemText(node.key);
    return itemProperty(d.itemCaptions?.[(axis === 'r' ? d.rows : d.cols)[node.depth]], t) ?? t;
  };
  // 요약 글자: 숫자(날짜) 항목은 원본 열의 표시 형식으로 (엑셀: '2023-12-06 요약')
  const subCap = (axis, node) => {
    const field = (axis === 'r' ? d.rows : d.cols)[node.depth];
    if (typeof node.key === 'number' && !itemProperty(d.itemCaptions?.[field], itemText(node.key))) {
      const st = resolved?.fieldStyle?.(field);
      if (st?.numFmt && st.numFmt !== 'general') return `${formatValue(node.key, st, d.date1904).text} 요약`;
    }
    return `${icap(axis, node)} 요약`;
  };
  // 부분합은 필드마다 켜고 끔 (엑셀 pivotField@defaultSubtotal): true = 모두, 배열 = 그 필드만
  const subOn = (axis, node) => d.subtotals === true || (Array.isArray(d.subtotals) && d.subtotals.includes((axis === 'r' ? d.rows : d.cols)[node.depth]));
  // 열 머리글 잎 목록: { cp, vi, kind: 'item' | 'sub' | 'grand', labels: [수준별 글자] }
  const Lc = colIdx.length;
  const multiV = V > 1;
  // Σ 값이 행 영역에 있으면 (엑셀 dataOnRows) 값 이름은 행 머리글 · 열 잎은 값 구분 없음
  const onRows = multiV && !!d.valuesOnRows;
  const colMulti = multiV && !onRows;
  const VI = V && !onRows ? [...Array(V).keys()] : [-1];
  const colLeaves = [];
  const setParents = (n) => n.children.forEach((c) => { c.parent = n.depth >= 0 ? n : null; setParents(c); });
  setParents(colTree);
  const walkCols = (node, labels) => {
    if (node.depth === Lc - 1 || !node.children.length || isColl('c', node)) {
      const pad = Array(Math.max(0, Lc - 1 - node.depth)).fill('');
      for (const vi of VI) colLeaves.push({ cp: node.path, vi, kind: 'item', labels: [...labels, ...pad, ...(colMulti ? [valueName(values[vi])] : [])], node });
      return;
    }
    node.children.forEach((ch) => walkCols(ch, [...labels, icap('c', ch)]));
    if (node.depth >= 0 && node.depth < Lc - 1 && subOn('c', node)) {
      for (const vi of VI) colLeaves.push({ cp: node.path, vi, kind: 'sub', labels: [...labels.slice(0, -1), subCap('c', node)], node });
    }
  };
  if (Lc) colTree.children.forEach((ch) => walkCols(ch, [icap('c', ch)]));
  else if (V) for (const vi of VI) colLeaves.push({ cp: '', vi, kind: 'item', labels: colMulti ? [valueName(values[vi])] : [], node: colTree });
  if (Lc && d.grandCols && V) {
    for (const vi of VI) colLeaves.push({ cp: '', vi, kind: 'grand', labels: [colMulti ? `전체 ${valueName(values[vi])}` : d.grandCaption ?? TOTAL], node: colTree });
  }
  // Σ 값 위치 (엑셀의 열 영역에서 '값'을 위아래로 옮긴 것): vp 번째 수준에 값 이름, 그보다 안쪽 항목은 값마다 반복
  // 예) [월, 값] → 4월 지표들 · 5월 지표들 / [값, 월] → 지표마다 4월 · 5월이 나란히
  const vp = colMulti && Lc ? Math.max(0, Math.min(Lc, d.valuesPos ?? Lc)) : Lc;
  if (vp < Lc) {
    const groupOf = (leaf) => {
      if (leaf.kind === 'grand') return '\u0002grand';
      let n = leaf.node;
      if (leaf.kind === 'sub' && n.depth < vp) return `${n.path}\u0001sub`;
      while (n && n.depth > vp - 1) n = n.parent;
      return n && n.depth >= 0 ? n.path : '';
    };
    const order = new Map();
    const keyed = colLeaves.map((leaf, i) => {
      const g = groupOf(leaf);
      if (!order.has(g)) order.set(g, order.size);
      return { leaf, i, g: order.get(g) };
    });
    keyed.sort((a, b) => a.g - b.g || a.leaf.vi - b.leaf.vi || a.i - b.i);
    colLeaves.length = 0;
    for (const { leaf } of keyed) {
      if (leaf.kind === 'item') {
        const vn = leaf.labels.pop();
        leaf.labels.splice(vp, 0, vn);
      } else if (leaf.kind === 'sub' && leaf.labels.length > vp) leaf.labels.splice(vp, 0, valueName(values[leaf.vi]));
      colLeaves.push(leaf);
    }
  }

  const defaultNumberCells=new WeakSet();
  const numberCell=(raw,style,role,vi)=>{const cd={raw,style,role};if(!values[vi]?.numFmt)defaultNumberCells.add(cd);return cd;};
  const numStyle = (vi) => {
    if (vi < 0) return {};
    const v = values[vi];
    if (v.numFmt) return typeof v.numFmt === 'object' ? v.numFmt : { numFmt: v.numFmt };
    if (showAsPercent(v.showAs)) return { numFmt: 'percent', decimals: 2 };
    if (v.showAs === 'rankAscending' || v.showAs === 'rankDescending') return { numFmt: 'general' };
    if (v.showAs === 'index') return { numFmt: 'number', decimals: 2 };
    if (calcNames.has(v.field.toLowerCase())) return { numFmt: 'number', decimals: 2 };
    return v.agg === 'average' || v.agg?.startsWith('std') || v.agg?.startsWith('var') ? { numFmt: 'number', decimals: 2 } : { numFmt: 'comma' };
  };
  const val = (n, vi, role) => {
    const style = { align: 'general', ...styleFor(role), ...numStyle(vi) };
    // 빈 셀 표시 옵션
    if (n === null || n === undefined) return numberCell(captionRaw(d.missingCaption),style,role,vi);
    // 오류 값 표시 옵션: 오류 대신 지정한 글자(빈 칸 포함)
    if (isErr(n)) return numberCell(d.errorCaption !== null && d.errorCaption !== undefined?errorCaptionRaw(d.errorCaption):n.code,style,role,vi);
    if (typeof n === 'boolean') return numberCell(n?'TRUE':'FALSE',style,role,vi);
    if (typeof n === 'string') return numberCell(`'${n}`,style,role,vi);
    return numberCell(String(Number(n.toPrecision(15))),style,role,vi);
  };

  const layout = d.layout;
  const Lr = rowIdx.length;
  // 행 · 열 필드가 없고 값만 있으면 엑셀은 레이블 열 없이 값 이름과 합계만 (총합계 글자 없음)
  const noLabel = Lr === 0 && colIdx.length === 0 && !onRows;
  const labelCols = noLabel ? 0 : layout === 'compact' ? 1 : Math.max(1, Lr + (onRows ? 1 : 0));
  const grid = [];
  const rowItems = [];

  // 보고서 필터 (공통 경로에서 배치·서식 적용한 결과)
  if (pageRows.length) { pushAll(grid, pageRows); grid.push([]); }

  // 열 머리글
  const colLevels = Lc + (colMulti ? 1 : 0);
  const hasColHead = colLevels > 0;
  const valueCaption = V === 1 ? valueName(values[0]) : '';
  const rowHeaderCells = () => {
    if (noLabel) return [];
    if (d.showHeaders === false) return Array.from({ length: labelCols }, (_, i) => text('', `rowHead:${i}`));
    if (layout === 'compact') return [text(Lr || onRows ? d.rowCaption ?? '행 레이블' : '', 'rowHead:0', !!(Lr || onRows) && d.rowCaption === '')];
    return Array.from({ length: labelCols }, (_, i) => {
      // 행 필드가 하나면 테이블·개요 형식에서도 Excel의 사용자 행 캡션을 표시합니다.
      const customRow = Lr === 1 && i === 0 && d.rowCaption !== undefined && d.rowCaption !== null;
      const valueHead = d.rows[i] === undefined && onRows && i === Lr;
      const caption = customRow ? d.rowCaption : d.rows[i] === undefined ? (valueHead ? d.dataCaption ?? '값' : '') : fcap(d.rows[i]);
      return text(caption, `rowHead:${i}`, customRow && caption === '' || valueHead && d.dataCaption === '');
    });
  };
  if (hasColHead) {
    // 열 필드가 있으면 맨 위에 '값 이름 | 열 레이블' 행 (값 필드만 여러 개면 생략)
    // 압축 형식은 '열 레이블' 하나, 개요 · 테이블 형식은 열 필드마다 필드 이름 (값 자리는 '값')
    if (Lc) {
      const caps = d.showHeaders === false ? [] : layout === 'compact' ? [d.colCaption ?? '열 레이블']
        : Array.from({ length: colLevels }, (_, lvl) => { const vLvl = colMulti ? vp : -1; return lvl === vLvl ? d.dataCaption ?? '값' : fcap(d.cols[vLvl >= 0 && lvl > vLvl ? lvl - 1 : lvl]); });
      grid.push([text(valueCaption, 'valueCaption'), ...Array(labelCols - 1).fill(null).map(() => text('', 'corner')), ...(colLeaves.length ? colLeaves.map((_, k) => text(caps[k] ?? '', 'colHead', k < caps.length && caps[k] === '')) : [text(caps[0] ?? '', 'colHead', caps.length > 0 && caps[0] === '')])]);
    }
    // 값 행 표시 또는 클래식 끌어 놓기: 열 필드 없이 값 여러 개면 값 이름 위에 '값' 행.
    if (!Lc && colMulti && d.valuesHeadRow) grid.push([...Array.from({ length: labelCols }, () => text('', 'corner')), ...colLeaves.map((_, k) => text(k ? '' : d.dataCaption ?? '값', 'colHead', k === 0 && d.dataCaption === ''))]);
    for (let lvl = 0; lvl < colLevels; lvl++) {
      const row = lvl === colLevels - 1 ? rowHeaderCells() : Array.from({ length: labelCols }, () => text('', 'corner'));
      let prev = null;
      colLeaves.forEach((leaf) => {
        const lab = leaf.labels[lvl];
        // 같은 그룹의 반복 글자는 첫 칸에만 (엑셀과 같음)
        const groupKey = leaf.labels.slice(0, lvl + 1).join('\u0001');
        const show = lab !== undefined && (lvl === colLevels - 1 || groupKey !== prev);
        prev = groupKey;
        const vLvl = colMulti ? vp : -1; // 값 이름이 있는 머리글 수준
        const cl = vLvl >= 0 && lvl > vLvl ? lvl - 1 : lvl; // 열 필드 수준
        const role = leaf.kind === 'grand' ? `grandHead:${Math.max(0, leaf.vi)}` : leaf.kind === 'sub' ? 'colSubHead' : lvl === vLvl ? `valueHead:${leaf.vi}` : `colItem:${cl}`;
        const cell = text(show ? lab : '', role, show && leaf.kind === 'grand' && !colMulti && d.grandCaption === '' && lab === '');
        // 펼치기 · 축소 단추: 하위 수준이 있는 열 항목
        if (show && lab && leaf.kind === 'item' && lvl !== vLvl && cl < Lc) {
          let n = leaf.node;
          while (n && n.depth > cl) n = n.parent;
          const tg = n && n.depth === cl ? toggleOf('c', n) : null;
          if (tg) cell.toggle = { axis: 'c', ...tg };
        }
        row.push(cell);
      });
      grid.push(row);
    }
  } else {
    grid.push([...rowHeaderCells(), ...(V ? [text(valueCaption, 'valueHead:0')] : [])]);
  }
  const firstDataRowRel = grid.length - pageRows.length - (pageRows.length ? 1 : 0);

  // 본문
  const dataRole = (leaf, rowKind, vi0 = leaf.vi) => {
    const vi = Math.max(0, vi0);
    if (rowKind === 'grand') return `grandData:${vi}`;
    if (leaf.kind === 'grand') return `grandColData:${vi}`;
    if (rowKind === 'sub') return `subData:${vi}`;
    if (rowKind === 'group') return `groupData:${vi}`;
    if (leaf.kind === 'sub') return `colSubData:${vi}`;
    return `data:${vi}`;
  };
  const valueCells = (rpath, rowKind, rvi = -1) => colLeaves.map((leaf) => {
    const vi = leaf.vi < 0 && rvi >= 0 ? rvi : leaf.vi;
    return val(cellValue(rpath, leaf.cp, vi), vi, dataRole(leaf, rowKind, vi));
  });
  // 값이 행 영역에 있을 때: 한 항목 → 값 필드마다 한 행 (값 이름은 '값' 열, 압축 형식은 들여 쓴 하위 행)
  const sigCol = layout === 'compact' ? 0 : labelCols - 1;
  const pushValueRows = (cells, rpath, kind, info, labelRole = `rowItem:${Lr}`, depth = 0) => {
    for (let vi = 0; vi < V; vi++) {
      const row = vi === 0 || d.repeatLabels ? cells.map((c) => ({ ...c, style: { ...c.style } })) : cells.map((c) => text('', c.role));
      const at = kind === 'grand' ? 0 : sigCol;
      row[at] = text(kind === 'grand' ? `전체 ${valueName(values[vi])}` : valueName(values[vi]), labelRole);
      if (layout === 'compact' && depth > 0) row[at].style.indent = depth;
      grid.push([...row, ...valueCells(rpath, kind, vi)]);
      rowItems.push({ ...info, vi });
    }
  };
  const blankData = () => colLeaves.map((leaf) => text('', `groupData:${Math.max(0, leaf.vi)}`));
  const labelsRow = (node, labelText, role, withToggle = false) => {
    const cells = Array.from({ length: labelCols }, () => text('', role));
    const col = layout === 'compact' ? 0 : node.depth;
    cells[col] = text(labelText, role);
    const img = imageOfKey(node.key);
    if (img) { cells[col].raw = ''; cells[col].image = { src: img.src, alt: img.alt ?? '' }; }
    if (layout === 'compact' && node.depth > 0) cells[col].style.indent = node.depth;
    if (withToggle) { const tg = toggleOf('r', node); if (tg) cells[col].toggle = { axis: 'r', ...tg }; }
    // 개요 형식 + 항목 레이블 반복: 상위 항목 글자를 빈 칸에 채움
    if (layout === 'outline' && d.repeatLabels) for (let n = node.parent; n && n.depth >= 0; n = n.parent) cells[n.depth] = text(icap('r', n), `rowItem:${n.depth}`);
    return cells;
  };
  const blankAfter = (node) => {
    // 항목 다음에 빈 줄 삽입 (가장 안쪽 필드 제외, 엑셀과 같음)
    if (!d.blankRows || node.depth >= Lr - 1 || (Array.isArray(d.blankRows) && !d.blankRows.includes(d.rows[node.depth]))) return;
    grid.push(Array.from({ length: labelCols + colLeaves.length }, () => text('', 'blank')));
    rowItems.push({ kind: 'blank', node });
  };
  const walkRows = (node, lastPath) => {
    const coll = isColl('r', node);
    const isLeaf = node.depth === Lr - 1 || coll;
    if (layout === 'tabular') {
      if (isLeaf) {
        // 조상 글자는 그룹의 첫 행에만 (항목 레이블 반복이면 모든 행)
        const cells = Array.from({ length: labelCols }, (_, dd) => text('', `rowItem:${dd}`));
        const chain = [];
        for (let n = node; n && n.depth >= 0; n = n.parent) chain.unshift(n);
        chain.forEach((n, dd) => {
          if (lastPath.shown.has(n.path) && !d.repeatLabels) return;
          cells[dd] = text(icap('r', n), `rowItem:${dd}`);
          const img = imageOfKey(n.key);
          if (img) { cells[dd].raw = ''; cells[dd].image = { src: img.src, alt: img.alt ?? '' }; }
          if (!lastPath.shown.has(n.path)) { const tg = toggleOf('r', n); if (tg) cells[dd].toggle = { axis: 'r', ...tg }; }
          lastPath.shown.add(n.path);
        });
        if (onRows) pushValueRows(cells, node.path, coll ? 'sub' : 'item', { kind: 'item', node, chain, coll });
        else {
          grid.push([...cells, ...valueCells(node.path, coll ? 'sub' : 'item')]);
          rowItems.push({ kind: 'item', node, chain, coll });
        }
        blankAfter(node);
        return;
      }
      node.children.forEach((ch) => walkRows(ch, lastPath));
      if (subOn('r', node)) {
        if (onRows) pushValueRows(labelsRow(node, subCap('r', node), `rowSub:${node.depth}`), node.path, 'sub', { kind: 'sub', node }, `rowSub:${node.depth}`);
        else {
          grid.push([...labelsRow(node, subCap('r', node), `rowSub:${node.depth}`), ...valueCells(node.path, 'sub')]);
          rowItems.push({ kind: 'sub', node });
        }
      }
      blankAfter(node);
      return;
    }
    const group = !isLeaf;
    // 부분합: 그룹 맨 위(기본) 또는 맨 아래 '… 요약' 행, 축소한 항목은 자기 행에 합계
    const valuesHere = coll || !group || (subOn('r', node) && d.subtotalTop && !onRows);
    const cells = labelsRow(node, icap('r', node), group ? `rowGroup:${node.depth}` : coll ? `rowGroup:${node.depth}` : `rowItem:${node.depth}`, true);
    grid.push([...cells, ...(valuesHere && !onRows ? valueCells(node.path, coll ? 'group' : group ? 'group' : 'item') : blankData())]);
    rowItems.push({ kind: 'item', node, coll });
    if (onRows && valuesHere) pushValueRows(cells.map((c) => text('', c.role)), node.path, coll ? 'group' : 'item', { kind: 'item', node, coll, valueRow: true }, `rowItem:${Lr}`, node.depth + 1);
    if (group) {
      node.children.forEach((ch) => walkRows(ch, lastPath));
      if (subOn('r', node) && (!d.subtotalTop || onRows)) {
        if (onRows) pushValueRows(labelsRow(node, subCap('r', node), `rowSub:${node.depth}`), node.path, 'sub', { kind: 'sub', node }, `rowSub:${node.depth}`, node.depth + 1);
        else {
          grid.push([...labelsRow(node, subCap('r', node), `rowSub:${node.depth}`), ...valueCells(node.path, 'sub')]);
          rowItems.push({ kind: 'sub', node });
        }
      }
    }
    blankAfter(node);
  };
  setParents(rowTree);
  if (Lr) rowTree.children.forEach((ch) => walkRows(ch, { shown: new Set() }));
  else if (onRows) pushValueRows(Array.from({ length: labelCols }, () => text('', 'rowItem:0')), '', 'item', { kind: 'item', node: rowTree }, 'rowItem:0');
  else if (V) grid.push([...(noLabel ? [] : [text(d.grandCaption ?? TOTAL, 'grandLabel', d.grandCaption === '')]), ...valueCells('', 'grand')]);
  if (Lr && d.grandRows && V) {
    const cells = Array.from({ length: labelCols }, () => text('', 'grandLabel'));
    if (onRows) pushValueRows(cells, '', 'grand', { kind: 'grand' }, 'grandLabel');
    else {
      cells[0] = text(d.grandCaption ?? TOTAL, 'grandLabel', d.grandCaption === '');
      grid.push([...cells, ...valueCells('', 'grand')]);
      rowItems.push({ kind: 'grand' });
    }
  }

  const width = labelCols + colLeaves.length;
  if (preset) {
    const leafCols = (kind) => colLeaves.map((l, i) => (l.kind === kind ? labelCols + i : -1)).filter((c) => c >= 0);
    paintPivotPreset(preset, grid, {
      pages: 0, top: pageRows.length ? pageRows.length + 1 : 0, headerRows: firstDataRowRel, labelCols, width,
      defaultNumberCells, grandCols: leafCols('grand'), subCols: leafCols('sub'), subColDepths:Object.fromEntries(colLeaves.map((leaf,i)=>[labelCols+i,Math.max(0,leaf.node?.depth??0)])), colLevels, colFields: Lc,
    }, opts);
  }
  // 줄무늬 행 · 열 (피벗 스타일 옵션)
  if (!preset && (opts.bandRows || opts.bandCols)) {
    const start = (pageRows.length ? pageRows.length + 1 : 0) + firstDataRowRel;
    let k = 0;
    for (let r = start; r < grid.length; r++) {
      const row = grid[r];
      if (!row?.length || /^grand/.test(row[0]?.role ?? '')) continue;
      row.forEach((cell, c) => {
        if (!cell || !/^(data|rowItem)/.test(cell.role)) return;
        const band = (opts.bandRows && k % 2 === 1) || (opts.bandCols && c >= labelCols && (c - labelCols) % 2 === 1);
        if (band && !cell.style.fill) cell.style = { ...cell.style, ...parts.band };
      });
      k++;
    }
  }

  return {
    grid,
    meta: {
      header, rowIdx, colIdx, valIdx, values, labelCols, pageRows: pageRows.length ? pageRows.length + 1 : 0, pageFields: pageLayout.fields, pageWidth: pageLayout.width,
      headerRows: firstDataRowRel, colLeaves, rowItems, colItems: colLeaves, rowTree, colTree, width,
      bodyRows: grid.length - (pageRows.length ? pageRows.length + 1 : 0),
    },
  };
}

/**
 * 피벗 차트 데이터: 행 항목(잎) = 항목 축, 열 항목 × 값 필드 = 계열 (부분합 · 총합계 제외, 엑셀과 같음)
 * → { categories: [글자], series: [{ name, values }] }
 */
const chartMemo = new WeakMap();
/** fieldStyle(필드 이름) → 원본 열 서식 (날짜 항목을 날짜로 표시하는 데 씀) */
export function pivotChartData(rows, def, fieldStyle = null) {
  def = dateDef(def, rows.date1904 ?? def.date1904);
  const holder = Array.isArray(rows) ? cubeFromRows(rows) : rows.cube;
  let memo = chartMemo.get(holder);
  if (!memo) { memo = new Map(); chartMemo.set(holder, memo); }
  const fmts = fieldStyle ? (def.rows ?? []).map((f) => fieldStyle(f)?.numFmt ?? null) : [];
  const key = pivotDefKey(def) + JSON.stringify(def.styleDef ?? null) + JSON.stringify(fmts);
  if (!memo.has(key)) {
    if (memo.size > 60) memo.clear();
    memo.set(key, pivotChartDataRaw(rows, def, fieldStyle));
  }
  return memo.get(key);
}
function pivotChartDataRaw(rows, def, fieldStyle) {
  const res = resolvePivot(rows, def);
  const d = res.def;
  const { grid, meta } = computePivot(res, d);
  if (meta.empty) return { categories: [], series: [] };
  const body = meta.pageRows + meta.headerRows;
  const cats = [];
  const rowIdx = [];
  // 값이 행 영역(dataOnRows): 행마다 값 필드 하나 → 항목 = (행 항목 /) 값 이름, 계열 = 열 항목
  const onRows = d.valuesOnRows && d.values.length > 1;
  if (!d.rows.length && !onRows) { cats.push(TOTAL); rowIdx.push(body); }
  meta.rowItems.forEach((it, i) => {
    if (it.kind !== 'item' || (it.node.children.length && !it.coll)) return;
    if (onRows && it.vi === undefined) return;
    const chain = [];
    if (onRows) chain.push(valueName(d.values[it.vi]));
    for (let n = it.node; n && n.depth >= 0; n = n.parent) {
      const st = typeof n.key === 'number' ? fieldStyle?.(d.rows[n.depth]) : null;
      const cap = itemProperty(d.itemCaptions?.[d.rows[n.depth]], itemText(n.key));
      chain.unshift(cap ?? (st?.numFmt && st.numFmt !== 'general' ? formatValue(n.key, st, d.date1904).text : itemText(n.key)));
    }
    cats.push(chain.join(' / '));
    rowIdx.push(body + i);
  });
  const series = [];
  meta.colLeaves.forEach((leaf, j) => {
    if (leaf.kind !== 'item' || (leaf.vi < 0 && !onRows)) return;
    const name = leaf.labels.length ? leaf.labels.join(' - ') : leaf.vi < 0 ? TOTAL : valueName(meta.values[leaf.vi]);
    const col = meta.labelCols + j;
    series.push({
      name,
      measure: leaf.vi < 0 ? null : valueName(meta.values[leaf.vi]),
      vi: leaf.vi,
      values: rowIdx.map((ri) => { const raw = grid[ri]?.[col]?.raw; const n = raw === '' || raw === undefined ? NaN : Number(raw); return Number.isFinite(n) ? n : null; }),
      numFmt: grid[rowIdx[0]]?.[col]?.style ?? null,
      col,
      headRow: body - 1,
    });
  });
  return { categories: cats, series, rows: rowIdx };
}

/** 옛 API: 셀 데이터 2차원 배열 */
export function buildPivot(rows, def) {
  const res = resolvePivot(rows, def);
  return computePivot(res, res.def).grid;
}

/** GETPIVOTDATA: 값 필드 이름과 (필드, 항목) 쌍으로 값 찾기 */
const lookupIdx = new WeakMap(); // res.groups → Map(필드 조합 → Map(항목 글자들 → 누적))
export function pivotLookup(rows, def, dataField, pairs, resolved = null) {
  const res = resolved ?? resolvePivot(rows, def);
  // 같은 값 필드 · 필드 조합으로 묻는 GETPIVOTDATA 가 수만 개여도 이름 풀이와 그룹 훑기는 한 번만 (항목 글자 → 누적 색인)
  let bySig = lookupIdx.get(res.groups);
  if (!bySig) { bySig = new Map(); lookupIdx.set(res.groups, bySig); }
  const sig = `${dataField}\u0002${pairs.map((p) => p[0]).join('\u0001')}`;
  let e = bySig.get(sig);
  if (!e) {
    const { def: d, header } = res;
    const lower = String(dataField).trim().toLowerCase();
    const vi = d.values.findIndex((v) => valueName(v).trim().toLowerCase() === lower || v.field.toLowerCase() === lower);
    // 행·열 필드가 아닌 필드로 묻는 것은 엑셀에서 #REF!
    const pos = pairs.map(([f]) => {
      const name = header.find((h) => h.toLowerCase() === String(f).toLowerCase());
      const ri = d.rows.indexOf(name);
      const ci = d.cols.indexOf(name);
      return ri >= 0 ? ['r', ri] : ci >= 0 ? ['c', ci] : null;
    });
    // 월 · 분기로 묶은 날짜 필드는 엑셀처럼 숫자 항목(12 → '12월', 3 → '3분기')으로도 찾음
    const by = pairs.map(([f]) => { const name = header.find((h) => h.toLowerCase() === String(f).toLowerCase()); return d.groups?.[name]?.by ?? null; });
    e = { vi, pos, idx: null, by };
    if (vi >= 0 && !pos.some((p) => !p)) {
      e.idx = new Map();
      // 엑셀: 축소한 항목 아래(보이지 않는) 항목은 GETPIVOTDATA 가 #REF!
      const deep = { r: Math.max(-1, ...pos.filter((p) => p[0] === 'r').map((p) => p[1])), c: Math.max(-1, ...pos.filter((p) => p[0] === 'c').map((p) => p[1])) };
      const coll = [['r', d.rows], ['c', d.cols]].flatMap(([ax, fields]) => fields.map((f, i) => [ax, i, d.collapsed?.[f]?.length ? new Set(d.collapsed[f].map(itemIdentity)) : null]).filter(([, i, set]) => set && i < deep[ax]));
      for (const g of res.groups) {
        if (coll.some(([ax, i, set]) => set.has(itemIdentity(itemText(g[ax][i]))))) continue;
        const k = pos.map(([ax, i]) => itemIdentity(itemText(g[ax][i]))).join('\u0001');
        let l = e.idx.get(k);
        if (!l) { l = res.measures.newList(); e.idx.set(k, l); }
        mergeList(l, g.list);
      }
    }
    bySig.set(sig, e);
  }
  if (!e.idx) return null;
  const itemKey = (v, i) => {
    const b = e.by[i];
    if (typeof v === 'number' && Number.isInteger(v)) { if (b === 'months' && v >= 1 && v <= 12) return `${v}월`; if (b === 'quarters' && v >= 1 && v <= 4) return `${v}분기`; }
    return itemText(v);
  };
  let list = e.idx.get(pairs.length === 1 ? itemIdentity(itemKey(pairs[0][1], 0)) : pairs.map((p, i) => itemIdentity(itemKey(p[1], i))).join('\u0001'));
  // 원본 행이 하나도 없는 피벗의 총합계: 엑셀은 빈 집계(합계 0, 계산 필드는 그 결과)를 돌려줌
  if (!list && !pairs.length && !res.groups.length) list = res.measures.newList();
  if (!list) return null;
  const v = res.measures.value(list, e.vi);
  if (typeof v === 'number') return v;
  // 오류 값: '오류 값 표시' 글자가 있으면 칸에 보이는 값 (빈 글자는 0), 없으면 그 오류
  if (isErr(v)) {
    const cap = res.def.errorCaption;
    if (cap === null || cap === undefined) return ERR_BY_CODE[v.code] ?? ERR.VALUE;
    return errorCaptionValue(cap) ?? 0;
  }
  // 있는 항목인데 값이 모두 빈 칸이라 빈칸으로 보이는 칸은 0
  return list && v === null ? 0 : null;
}

/**
 * 세부 정보 표시 (값 셀 두 번 클릭): 피벗 결과의 (gr, gc) 칸을 이루는 원본 행
 * → { header, cube, idx: Uint32Array (큐브 행 번호), conds: [[필드, 항목 글자]] } | null (값 칸이 아님)
 */
export function pivotDetail(input, def, gr, gc) {
  const res = resolvePivot(input, def);
  const d = res.def;
  const { meta } = computePivot(res, d);
  if (meta.empty) return null;
  const bi = gr - meta.pageRows - meta.headerRows;
  const ci = gc - meta.labelCols;
  if (bi < 0 || ci < 0 || ci >= meta.colLeaves.length) return null;
  const leaf = meta.colLeaves[ci];
  const vi = leaf.vi < 0 ? meta.rowItems[bi]?.vi : leaf.vi;
  if (!Number.isInteger(vi) || !d.values[vi]) return null;
  let rnode = null;
  if (d.rows.length) {
    const it = meta.rowItems[bi];
    if (!it || it.kind === 'blank') return null;
    rnode = it.kind === 'grand' ? null : it.node;
  } else if (d.valuesOnRows && d.values.length > 1 ? !meta.rowItems[bi] : bi !== 0) return null;
  const cube = res.cube;
  const lower = cube.header.map((h) => String(h).toLowerCase());
  const dimFor = (name) => {
    const j = lower.indexOf(String(name).toLowerCase());
    const spec = d.groups?.[cube.header[j]];
    return spec ? groupedColumn(cube, j, spec) : cube.col(j);
  };
  // 조건: 행 · 열 항목 경로 (그룹화한 필드는 그룹 키로)
  const conds = [];
  const tests = [];
  const addCond = (field, key) => {
    conds.push([field, itemText(key)]);
    const { codes, keys } = dimFor(field).dim();
    const want = kk(key);
    const ok = new Uint8Array(keys.length);
    keys.forEach((k, c) => { ok[c] = kk(k) === want ? 1 : 0; });
    tests.push({ codes, ok });
  };
  for (let n = rnode; n && n.depth >= 0; n = n.parent) addCond(d.rows[n.depth], n.key);
  for (let n = leaf.kind === 'grand' ? null : leaf.node; n && n.depth >= 0; n = n.parent) addCond(d.cols[n.depth], n.key);
  // 값 · 레이블 · 상위 10 필터로 숨긴 항목은 제외 (보이는 그룹만)
  let visible = null;
  const allDims = [...d.rows, ...d.cols].map(dimFor);
  if (Object.keys(d.fieldFilters ?? {}).length && allDims.length) {
    visible = new Set(res.groups.map((g) => [...g.r, ...g.c].map(kk).join('\u0001')));
    allDims.forEach((col) => col.dim());
  }
  const sel = filterRows(cube, res.filters ?? []);
  const n = sel ? sel.length : cube.n;
  const out = new Uint32Array(n);
  let m = 0;
  const T = tests.length;
  outer: for (let x = 0; x < n; x++) {
    const i = sel ? sel[x] : x;
    for (let t = 0; t < T; t++) if (!tests[t].ok[tests[t].codes[i]]) continue outer;
    if (visible) {
      const key = allDims.map((col) => { const dm = col._dim ?? col.dim(); return kk(dm.keys[dm.codes[i]]); }).join('\u0001');
      if (!visible.has(key)) continue;
    }
    out[m++] = i;
  }
  return { header: cube.header, cube, idx: out.slice(0, m), conds, valueField: valueName(d.values[vi]) };
}

/**
 * 추천 피벗 테이블 (엑셀 삽입 › 추천 피벗 테이블): 머리글과 열 값(표본)을 보고 요약 후보 def 조각을 돌려줌
 * [{ title, rows, cols, values }] — 범주 열(항목 2~60개) × 숫자 열 합계, 두 범주 교차, 숫자 열이 없으면 개수
 */
export function recommendPivots(header, columns, max = 8, { dates = [] } = {}) {
  const info = header.map((name, i) => {
    const vals = (columns[i] ?? []).filter((v) => v !== null && v !== undefined && v !== '');
    const nums = vals.filter((v) => typeof v === 'number');
    const distinct = new Set(vals.map((v) => (typeof v === 'string' ? v.toLowerCase() : v))).size;
    const idLike = /(^|[^a-z])(id|no)$|번호|코드|순번/i.test(String(name));
    return { name: String(name), n: vals.length, numeric: !dates[i] && vals.length > 0 && nums.length / vals.length >= 0.8, distinct, idLike, date: !!dates[i] };
  }).filter((f) => f.n > 0 && f.name);
  const measures = info.filter((f) => f.numeric && !f.idLike && f.distinct > 1);
  const dateDim = info.find((f) => f.date && f.distinct >= 2);
  const dims = info.filter((f) => !f.numeric && !f.date && f.distinct >= 2 && f.distinct <= 60).sort((a, b) => a.distinct - b.distinct);
  const out = [];
  const seen = new Set();
  const push = (title, rows, cols, values, groups) => {
    const k = JSON.stringify([rows, cols, values]);
    if (seen.has(k) || out.length >= max) return;
    seen.add(k);
    out.push({ title, rows, cols, values, ...(groups ? { groups } : {}) });
  };
  const sumOf = (m) => ({ field: m.name, agg: 'sum' });
  // 날짜 열: 월로 묶어 추이 (엑셀 자동 날짜 그룹과 같음)
  if (dateDim && measures.length) push(`${measures[0].name} 합계 : ${dateDim.name}(월) 기준`, [dateDim.name], [], [sumOf(measures[0])], { [dateDim.name]: { by: 'months' } });
  for (const d of dims.slice(0, 3)) for (const m of measures.slice(0, 2)) push(`${m.name} 합계 : ${d.name} 기준`, [d.name], [], [sumOf(m)]);
  if (measures.length >= 2) for (const d of dims.slice(0, 2)) push(`${measures.slice(0, 3).map((m) => m.name).join(' · ')} 합계 : ${d.name} 기준`, [d.name], [], measures.slice(0, 3).map(sumOf));
  if (dims.length >= 2 && measures.length) {
    const [a, b] = dims[0].distinct <= 12 ? [dims[1], dims[0]] : [dims[0], dims[1]];
    if (b.distinct <= 12) push(`${measures[0].name} 합계 : ${a.name} × ${b.name}`, [a.name], [b.name], [sumOf(measures[0])]);
    push(`${measures[0].name} 합계 : ${dims[0].name} › ${dims[1].name}`, [dims[0].name, dims[1].name], [], [sumOf(measures[0])]);
  }
  if (dateDim && measures.length && dims.length && dims[0].distinct <= 12) push(`${measures[0].name} 합계 : ${dateDim.name}(월) × ${dims[0].name}`, [dateDim.name], [dims[0].name], [sumOf(measures[0])], { [dateDim.name]: { by: 'months' } });
  for (const d of dims.slice(0, 3)) for (const m of measures.slice(0, 1)) push(`${m.name} 평균 : ${d.name} 기준`, [d.name], [], [{ field: m.name, agg: 'average' }]);
  for (const d of dims.slice(0, 3)) push(`${d.name} 개수`, [d.name], [], [{ field: d.name, agg: 'count' }]);
  return out;
}
