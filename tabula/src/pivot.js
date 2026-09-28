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
import { formatGeneral, formatValue } from './format.js';
import { findTable, dataTop, dataBottom, columnNames, ACCENTS, tint, shade } from './tables.js';

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
];
export const LAYOUTS = [
  { id: 'compact', label: '압축 형식으로 표시' },
  { id: 'outline', label: '개요 형식으로 표시' },
  { id: 'tabular', label: '테이블 형식으로 표시' },
];

export const EMPTY = '(비어 있음)';
export const TOTAL = '총합계';
// 셀 안 그림은 주소로 구별 (피벗 항목 · 슬라이서에서 그림을 그대로 보여 줌)
export const IMG_KEY = '\u0000img:';
const imageByKey = new Map();
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

const collator = new Intl.Collator('ko');
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

/** 슬라이서·필터에서 쓰는 항목 글자 */
export const itemText = (v) => {
  if (v === null || v === undefined || v === '') return EMPTY;
  if (typeof v === 'number') return formatGeneral(v);
  if (typeof v === 'object') return v.type === 'image' ? v.alt || `그림 ${String(v.src).slice(-12)}` : v.code;
  const img = imageOfKey(v);
  return img ? img.alt || `그림 ${String(img.src).slice(-12)}` : String(v);
};

/** 머리글 이름 (빈 칸은 열N) */
export const headerNames = (rows) => (rows[0] ?? []).map((h, i) => (h === null || h === '' ? `열${i + 1}` : String(h)));

export const aggLabel = (agg) => AGGREGATES.find((a) => a.id === agg)?.label ?? '합계';
export const valueName = (v) => v.name || `${aggLabel(v.agg)} : ${v.field}`;

// ───────────── 계산 필드 수식 (필드 이름 · 숫자 · 사칙 연산 · 몇 가지 함수) ─────────────
const CALC_ERR = (code) => ({ code });
const isErr = (v) => v !== null && typeof v === 'object';

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
    const op = /^(<>|<=|>=|[-+*/^&=<>(),;])/.exec(text.slice(i));
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
  const un = () => (eat('-') ? { op: 'neg', a: un() } : eat('+') ? un() : prim());
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
        if (!eat(')')) { do args.push(expr()); while (eat(',')); if (!eat(')')) throw new Error('수식 오류'); }
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
  const ev = (n) => {
    if ('num' in n) return n.num;
    if ('str' in n) return n.str;
    if (n.field !== undefined) return get(n.field);
    if (n.fn) {
      const A = n.args;
      switch (n.fn) {
        case 'IF': { const c = num(ev(A[0])); if (isErr(c)) return c; return c ? (A[1] ? ev(A[1]) : true) : (A[2] ? ev(A[2]) : false); }
        case 'IFERROR': { const v = ev(A[0]); return isErr(v) ? ev(A[1]) : v; }
        case 'ROUND': { const v = num(ev(A[0])); const d = A[1] ? num(ev(A[1])) : 0; if (isErr(v)) return v; const f = 10 ** d; return Math.round(v * f) / f; }
        case 'ABS': { const v = num(ev(A[0])); return isErr(v) ? v : Math.abs(v); }
        case 'SQRT': { const v = num(ev(A[0])); return isErr(v) ? v : v < 0 ? CALC_ERR('#NUM!') : Math.sqrt(v); }
        case 'INT': { const v = num(ev(A[0])); return isErr(v) ? v : Math.floor(v); }
        case 'SUM': case 'MIN': case 'MAX': case 'AVERAGE': {
          const vs = A.map((x) => num(ev(x)));
          const e = vs.find(isErr);
          if (e) return e;
          if (n.fn === 'SUM') return vs.reduce((a, b) => a + b, 0);
          if (n.fn === 'AVERAGE') return vs.length ? vs.reduce((a, b) => a + b, 0) / vs.length : 0;
          return n.fn === 'MIN' ? Math.min(...vs) : Math.max(...vs);
        }
        case 'AND': return A.every((x) => num(ev(x)));
        case 'OR': return A.some((x) => num(ev(x)));
        case 'NOT': return !num(ev(A[0]));
        default: return CALC_ERR('#NAME?');
      }
    }
    if (n.op === 'neg') { const v = num(ev(n.a)); return isErr(v) ? v : -v; }
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
  }
}
function result(acc, agg) {
  if (!acc) return null;
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
    default: return acc.sum;
  }
}

/**
 * 값 필드 계산기: 필요한 원본 열만 누적하고 계산 필드는 합계에 수식을 적용
 * header: 원본 머리글 (+ 계산 필드 이름)
 */
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
  const baseIdx = (name) => header.findIndex((h) => lower(h) === lower(name) && !calcByName.has(lower(h)));
  const specs = values.map((v) => {
    const c = calcByName.get(lower(v.field));
    if (c) {
      const walk = (ast, seen) => calcRefs(ast).forEach((n) => {
        const cc = calcByName.get(lower(n));
        if (cc) { if (!seen.has(lower(n))) walk(cc.ast, new Set([...seen, lower(n)])); } else { const i = baseIdx(n); if (i >= 0) need(i); }
      });
      if (c.ast) walk(c.ast, new Set([lower(c.name)]));
      return { calc: c };
    }
    return { slot: need(baseIdx(v.field)), agg: v.agg };
  });
  const calcValue = (c, list, depth = 0) => {
    if (!c.ast) return CALC_ERR('#NAME?');
    if (depth > 20) return CALC_ERR('#REF!');
    return evalCalc(c.ast, (name) => {
      const cc = calcByName.get(lower(name));
      if (cc) return calcValue(cc, list, depth + 1);
      const i = baseIdx(name);
      if (i < 0) return CALC_ERR('#NAME?');
      return list[slot.get(i)]?.sum ?? 0;
    });
  };
  const n = cols.length;
  return {
    cols,
    value(list, vi) {
      if (!list) return null;
      const sp = specs[vi];
      if (!sp) return null;
      if (sp.calc) return calcValue(sp.calc, list);
      return result(list[sp.slot], sp.agg);
    },
    accumulate(list, r) {
      for (let k = 0; k < n; k++) add(list[k], cols[k] < 0 ? 1 : r[cols[k]]);
    },
    newList() { return cols.map(newAcc); },
  };
}

// ───────────── 정의 정리 ─────────────
/** 원본 머리글 + 계산 필드 이름 */
export function pivotFieldNames(rows, def) {
  const header = headerNames(rows);
  return [...header, ...(def.calcFields ?? []).map((c) => c.name).filter((n) => !header.some((h) => h.toLowerCase() === String(n).toLowerCase()))];
}

/** 옛 정의까지 포함해 필드 이름 기반 정의로 (header: 원본 머리글) */
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
    rows: (rows ?? []).filter(ok).map(byName),
    cols: (cols ?? []).filter(ok).map(byName),
    pages: (def.pages ?? []).filter(ok).map(byName),
    values: (values ?? []).filter((v) => ok(v.field)).map((v) => ({ ...v, field: byName(v.field), agg: v.agg ?? 'sum' })),
    layout: def.layout ?? 'compact',
    subtotals: def.subtotals !== false,
    grandRows: def.grandRows !== false,
    grandCols: def.grandCols !== false,
    filters: def.filters ?? {},
    calcFields: def.calcFields ?? [],
    sort: byKey(def.sort),
    order: byKey(def.order),
    fieldFilters: byKey(def.fieldFilters),
    style: def.style ?? DEFAULT_PIVOT_STYLE,
    styleDef: def.styleDef ?? null,
    errorCaption: def.errorCaption ?? null,
    rowCaption: def.rowCaption ?? null,
    colCaption: def.colCaption ?? null,
    styleOpts: { rowHeaders: true, colHeaders: true, bandRows: false, bandCols: false, ...(def.styleOpts ?? {}) },
    header: allHeader,
  };
}

/**
 * 피벗 원본 → { rows (머리글 포함 값), sheet, ref: {r1,c1,r2,c2} | null, table: 표 이름 | null }
 * 표 이름이면 지금의 표 범위(누적된 데이터 포함), 아니면 고정 범위
 */
export function pivotSourceData(wb, def) {
  if (def.table) {
    const f = findTable(wb, def.table);
    if (!f) return null;
    const t = f.t;
    const ref = { r1: t.header ? t.r1 : dataTop(t), c1: t.c1, r2: dataBottom(t), c2: t.c2 };
    const rows = cachedRead(wb, f.si, ref);
    if (!t.header) return { rows: [columnNames(wb, f.si, t), ...rows], si: f.si, ref, table: t.name };
    return { rows, si: f.si, ref, table: t.name };
  }
  const si = wb.sheetIndexByName(def.source);
  if (si < 0 || !def.range) return null;
  return { rows: cachedRead(wb, si, def.range), si, ref: def.range, table: null };
}

// 같은 원본을 여러 피벗 · 슬라이서가 읽으므로 통합 문서가 바뀌기 전까지(wb.version 이 같으면) 재사용
const readCache = new WeakMap();
function cachedRead(wb, si, ref) {
  const read = () => {
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
  ['between', '해당 범위'], ['notBetween', '해당 범위 제외'],
];
export const VALUE_OPS = LABEL_OPS.filter(([id]) => !/With|ontains/.test(id));

/** 필터 설명 (메뉴 · 필드 창) */
export function describeFieldFilter(f, values) {
  if (!f) return '';
  const vname = () => { const v = values[valueIndex(values, f.by)]; return v ? valueName(v) : ''; };
  if (f.type === 'top') return `${f.top === false ? '하위' : '상위'} ${f.n}${f.mode === 'percent' ? '%' : f.mode === 'sum' ? ' (합계)' : '개'} · ${vname()}`;
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

/** 행·열 필드 필터를 원본 행에 적용 (바깥 필드부터, 부모 그룹 안에서 평가) */
function applyFieldFilters(data, header, d) {
  const entries = Object.entries(d.fieldFilters ?? {});
  if (!entries.length) return data;
  const idx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const measures = d.values.length ? makeMeasures(header, d.values, d.calcFields) : null;
  const measureOf = (rows, vi) => {
    if (!measures) return rows.length;
    const list = measures.newList();
    for (const r of rows) measures.accumulate(list, r);
    const v = measures.value(list, vi);
    return typeof v === 'number' ? v : null;
  };
  let out = data;
  for (const axis of [d.rows, d.cols]) {
    axis.forEach((field, level) => {
      const flt = d.fieldFilters[field];
      const f = idx(field);
      if (!flt || f < 0) return;
      const parents = axis.slice(0, level).map(idx);
      const pkOf = (r) => parents.map((p) => kk(keyOf(r[p]))).join('\u0001');
      const groups = new Map();
      for (const r of out) {
        const pk = pkOf(r);
        let g = groups.get(pk);
        if (!g) { g = new Map(); groups.set(pk, g); }
        const k = kk(keyOf(r[f]));
        let it = g.get(k);
        if (!it) { it = { key: keyOf(r[f]), rows: [] }; g.set(k, it); }
        it.rows.push(r);
      }
      const keep = new Set();
      for (const [pk, g] of groups) {
        const items = [...g.values()];
        let kept;
        if (flt.type === 'label') {
          const numeric = items.every((it) => typeof it.key === 'number') && Number.isFinite(Number(flt.v1));
          kept = items.filter((it) => compareOp(flt.op, numeric ? it.key : itemText(it.key), numeric ? Number(flt.v1) : flt.v1, numeric ? Number(flt.v2) : flt.v2, !numeric));
        } else if (flt.type === 'value') {
          const vi = valueIndex(d.values, flt.by);
          kept = items.filter((it) => { const v = measureOf(it.rows, vi); return v !== null && compareOp(flt.op, v, Number(flt.v1), Number(flt.v2), false); });
        } else if (flt.type === 'top') {
          const vi = valueIndex(d.values, flt.by);
          const scored = items.map((it) => ({ it, v: measureOf(it.rows, vi) ?? -Infinity }));
          scored.sort((a, b) => (flt.top === false ? a.v - b.v : b.v - a.v));
          const n = Number(flt.n) || 10;
          if ((flt.mode ?? 'count') === 'count') kept = scored.slice(0, Math.max(0, Math.floor(n))).map((x) => x.it);
          else {
            const total = scored.reduce((s, x) => s + (Number.isFinite(x.v) ? x.v : 0), 0);
            const limit = flt.mode === 'percent' ? (total * n) / 100 : n;
            kept = [];
            let acc = 0;
            for (const x of scored) { if (acc >= limit && kept.length) break; kept.push(x.it); acc += Number.isFinite(x.v) ? x.v : 0; }
          }
        } else kept = items;
        for (const it of kept) keep.add(`${pk}\u0003${kk(it.key)}`);
      }
      out = out.filter((r) => keep.has(`${pkOf(r)}\u0003${kk(keyOf(r[f]))}`));
    });
  }
  return out;
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
export function resolvePivot(rows, def) {
  let memo = resolveMemo.get(rows);
  if (!memo) { memo = new Map(); resolveMemo.set(rows, memo); }
  const key = pivotDefKey(def);
  const hit = memo.get(key);
  if (hit) return { ...hit, def: normalizeDef(def, hit.header) };
  const d = normalizeDef(def, headerNames(rows));
  const header = d.header;
  const filters = Object.entries(d.filters).map(([name, allowed]) => [header.findIndex((h) => h.toLowerCase() === name.toLowerCase()), new Set(allowed)]).filter(([i]) => i >= 0);
  let data = nonEmptyMemo.get(rows);
  if (!data) { data = rows.slice(1).filter((r) => !r.every((v) => v === null || v === '')); nonEmptyMemo.set(rows, data); }
  // 필터 검사는 값마다 한 번만 (같은 값이 수만 행에 반복됨)
  const tests = filters.map(([i, set]) => {
    const seen = new Map();
    return (r) => {
      const v = r[i];
      let b = seen.get(v);
      if (b === undefined) { b = set.has(itemText(v)); seen.set(v, b); }
      return b;
    };
  });
  // 같은 원본 · 같은 필터를 쓰는 피벗(슬라이서로 묶인 피벗들)은 걸러 낸 행을 함께 씀
  const fkey = JSON.stringify(filters.map(([i, set]) => [i, [...set].sort()]));
  let fmemo = filterMemo.get(rows);
  if (!fmemo) { fmemo = new Map(); filterMemo.set(rows, fmemo); }
  let out = fmemo.get(fkey);
  if (!out) {
    out = tests.length ? data.filter((r) => tests.every((t) => t(r))) : data;
    if (fmemo.size > 30) fmemo.clear();
    fmemo.set(fkey, out);
  }
  out = applyFieldFilters(out, header, d);
  const res = { def: d, rows: [rows[0], ...out], header };
  if (memo.size > 60) memo.clear();
  memo.set(key, res);
  return res;
}

// ───────────── 피벗 스타일 (엑셀 기본 제공 이름: 밝게 1~28, 보통 1~28, 어둡게 1~28) ─────────────
export const DEFAULT_PIVOT_STYLE = 'PivotStyleLight16';
/** 스타일 이름 → 역할별 서식 { header, sub, grand, body, page } */
export function pivotStyleParts(name, custom = null) {
  // 파일에서 가져온 사용자 지정 스타일
  if (custom) return { header: {}, sub: {}, grand: {}, body: {}, page: {}, band: {}, ...custom };
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
export const PIVOT_STYLES = ['Light', 'Medium', 'Dark'].flatMap((k, gi) => Array.from({ length: 28 }, (_, i) => {
  const name = `PivotStyle${k}${i + 1}`;
  const p = pivotStyleParts(name);
  const grp = ['밝게', '보통', '어둡게'][gi];
  return { name, group: grp, label: `피벗 스타일 ${grp} ${i + 1}`, swatch: [p.header.fill ?? '#ffffff', p.sub.fill ?? p.body.fill ?? '#ffffff', p.grand.fill ?? '#ffffff'] };
}));

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
const kk = (k) => `${typeof k}:${k}`;

/** 필드 값 트리 */
function buildTree(data, idxs) {
  const root = { key: null, path: '', depth: -1, children: [], map: new Map() };
  for (const r of data) {
    let node = root;
    let path = '';
    for (let d = 0; d < idxs.length; d++) {
      const k = keyOf(r[idxs[d]]);
      const kkey = kk(k);
      path += `\u0001${kkey}`;
      let ch = node.map.get(kkey);
      if (!ch) { ch = { key: k, path, depth: d, children: [], map: new Map() }; node.map.set(kkey, ch); node.children.push(ch); }
      node = ch;
    }
  }
  return root;
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
      const score = new Map(kids.map((c) => [c, measureAt(c, vi)]));
      const num = (v) => (typeof v === 'number' ? v : -Infinity);
      kids = [...kids].sort((a, b) => (s.dir === 'desc' ? num(score.get(b)) - num(score.get(a)) : num(score.get(a)) - num(score.get(b))));
    } else {
      const keys = sortKeys(kids.map((c) => c.key));
      const byKey = new Map(kids.map((c) => [kk(c.key), c]));
      kids = keys.map((k) => byKey.get(kk(k)));
      if (order?.length && !s) {
        const pos = new Map(order.map((t, i) => [t, i]));
        const known = kids.filter((c) => pos.has(itemText(c.key))).sort((a, b) => pos.get(itemText(a.key)) - pos.get(itemText(b.key)));
        kids = [...known, ...kids.filter((c) => !pos.has(itemText(c.key)))];
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
export function computePivot(rows, d) {
  const header = d.header ?? headerNames(rows);
  const idx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const data = rows.slice(1).filter((r) => !r.every((v) => v === null || v === ''));
  const rowIdx = d.rows.map(idx);
  const colIdx = d.cols.map(idx);
  const values = d.values;
  const V = values.length;
  const parts = pivotStyleParts(d.style, d.styleDef);
  const opts = d.styleOpts ?? { rowHeaders: true, colHeaders: true };
  const HEAD_ROLES = /^(corner|rowHead|colHead|valueCaption|colItem|valueHead|colSubHead|grandHead)/;
  const styleFor = (role) => {
    if (opts.colHeaders === false && HEAD_ROLES.test(role)) return {};
    const s = roleStyle(parts, role);
    if (opts.rowHeaders === false && /^(rowGroup|rowItem)/.test(role)) { const { bold, ...rest } = s; return rest; }
    return s;
  };
  const valIdx = values.map((v) => idx(v.field));
  const calcNames = new Set((d.calcFields ?? []).map((c) => c.name.toLowerCase()));

  const text = (s, role) => ({ raw: typeof s === 'number' ? formatGeneral(s) : String(s ?? '').startsWith('=') ? `'${s}` : String(s ?? ''), style: { ...styleFor(role) }, role });
  if (!d.rows.length && !d.cols.length && !V) {
    const grid = Array.from({ length: 18 }, (_, r) => Array.from({ length: 3 }, (_, c) => ({
      raw: r === 0 && c === 0 ? '피벗 테이블 보고서를 작성하려면 [피벗 테이블 필드] 목록에서 필드를 선택하세요.' : '',
      style: { fill: '#f3f6fb', ...(r === 0 ? { bt: true } : {}), ...(r === 17 ? { bb: true } : {}), ...(c === 0 ? { bl: true } : {}), ...(c === 2 ? { br: true } : {}), ...(r === 0 && c === 0 ? { color: '#44546a' } : {}) },
      role: 'empty',
    })));
    return { grid, meta: { header, rowIdx, colIdx, valIdx, values, labelCols: 1, pageRows: 0, headerRows: 1, colLeaves: [], rowItems: [], colItems: [], rowTree: null, colTree: null, width: 3, bodyRows: 18, empty: true } };
  }

  const measures = makeMeasures(header, values, d.calcFields);
  const rowTree = buildTree(data, rowIdx);
  const colTree = buildTree(data, colIdx);

  // 누적: 행 경로 접두사 × 열 경로 접두사
  const accs = new Map();
  const rp = new Array(rowIdx.length + 1);
  const cp = new Array(colIdx.length + 1);
  for (const r of data) {
    rp[0] = '';
    for (let i = 0; i < rowIdx.length; i++) rp[i + 1] = `${rp[i]}\u0001${kk(keyOf(r[rowIdx[i]]))}`;
    cp[0] = '';
    for (let i = 0; i < colIdx.length; i++) cp[i + 1] = `${cp[i]}\u0001${kk(keyOf(r[colIdx[i]]))}`;
    for (const a of rp) {
      for (const b of cp) {
        const key = `${a}\u0002${b}`;
        let list = accs.get(key);
        if (!list) { list = measures.newList(); accs.set(key, list); }
        measures.accumulate(list, r);
      }
    }
  }
  const raw = (rpath, cpath, vi) => measures.value(accs.get(`${rpath}\u0002${cpath}`), vi);
  orderTree(rowTree, d.rows, d, (node, vi) => raw(node.path, '', vi));
  orderTree(colTree, d.cols, d, (node, vi) => raw('', node.path, vi));
  const cellValue = (rpath, cpath, vi) => {
    if (vi < 0) return null;
    const v = raw(rpath, cpath, vi);
    if (v === null || typeof v !== 'number') return v;
    const as = values[vi].showAs ?? 'normal';
    if (as === 'normal') return v;
    const base = as === 'percentOfTotal' ? raw('', '', vi) : as === 'percentOfRow' ? raw(rpath, '', vi) : raw('', cpath, vi);
    return typeof base === 'number' && base ? v / base : null;
  };

  // 열 머리글 잎 목록: { cp, vi, kind: 'item' | 'sub' | 'grand', labels: [수준별 글자] }
  const Lc = colIdx.length;
  const multiV = V > 1;
  const VI = V ? [...Array(V).keys()] : [-1];
  const colLeaves = [];
  const setParents = (n) => n.children.forEach((c) => { c.parent = n.depth >= 0 ? n : null; setParents(c); });
  setParents(colTree);
  const walkCols = (node, labels) => {
    if (node.depth === Lc - 1 || !node.children.length) {
      for (const vi of VI) colLeaves.push({ cp: node.path, vi, kind: 'item', labels: [...labels, ...(multiV ? [valueName(values[vi])] : [])], node });
      return;
    }
    node.children.forEach((ch) => walkCols(ch, [...labels, itemText(ch.key)]));
    if (d.subtotals && node.depth >= 0 && node.depth < Lc - 1) {
      for (const vi of VI) colLeaves.push({ cp: node.path, vi, kind: 'sub', labels: [...labels.slice(0, -1), `${itemText(node.key)} 요약`], node });
    }
  };
  if (Lc) colTree.children.forEach((ch) => walkCols(ch, [itemText(ch.key)]));
  else if (V) for (const vi of VI) colLeaves.push({ cp: '', vi, kind: 'item', labels: multiV ? [valueName(values[vi])] : [], node: colTree });
  if (Lc && d.grandCols && V) {
    for (const vi of VI) colLeaves.push({ cp: '', vi, kind: 'grand', labels: [multiV ? `전체 ${valueName(values[vi])}` : TOTAL], node: colTree });
  }

  const numStyle = (vi) => {
    if (vi < 0) return {};
    const v = values[vi];
    if (v.numFmt) return typeof v.numFmt === 'object' ? v.numFmt : { numFmt: v.numFmt };
    if ((v.showAs ?? 'normal') !== 'normal') return { numFmt: 'percent', decimals: 2 };
    if (calcNames.has(v.field.toLowerCase())) return { numFmt: 'number', decimals: 2 };
    return v.agg === 'average' || v.agg?.startsWith('std') || v.agg?.startsWith('var') ? { numFmt: 'number', decimals: 2 } : { numFmt: 'comma' };
  };
  const val = (n, vi, role) => {
    const style = { ...styleFor(role), ...numStyle(vi) };
    if (n === null || n === undefined) return { raw: '', style, role };
    // 오류 값 표시 옵션: 오류 대신 지정한 글자(빈 칸 포함)
    if (isErr(n)) return d.errorCaption !== null && d.errorCaption !== undefined ? { raw: d.errorCaption === '' ? '' : `'${d.errorCaption}`, style, role } : { raw: n.code, style, role };
    if (typeof n === 'boolean') return { raw: n ? 'TRUE' : 'FALSE', style, role };
    if (typeof n === 'string') return { raw: `'${n}`, style, role };
    return { raw: String(Number(n.toPrecision(15))), style, role };
  };

  const layout = d.layout;
  const Lr = rowIdx.length;
  const labelCols = layout === 'compact' ? 1 : Math.max(1, Lr);
  const grid = [];
  const rowItems = [];

  // 보고서 필터 (맨 위)
  const pageRows = [];
  for (const p of d.pages) {
    const allowed = d.filters[p] ?? d.filters[Object.keys(d.filters).find((k) => k.toLowerCase() === p.toLowerCase())];
    pageRows.push([text(p, 'pageLabel'), text(!allowed ? '(모두)' : allowed.length === 1 ? allowed[0] : '(다중 항목)', 'pageValue')]);
  }
  if (pageRows.length) { grid.push(...pageRows); grid.push([]); }

  // 열 머리글
  const colLevels = Lc + (multiV ? 1 : 0);
  const hasColHead = colLevels > 0;
  const valueCaption = V === 1 ? valueName(values[0]) : '';
  const rowHeaderCells = () => {
    if (layout === 'compact') return [text(Lr ? d.rowCaption ?? '행 레이블' : '', 'rowHead:0')];
    return Array.from({ length: labelCols }, (_, i) => text(d.rows[i] ?? '', `rowHead:${i}`));
  };
  if (hasColHead) {
    // 열 필드가 있으면 맨 위에 '값 이름 | 열 레이블' 행 (값 필드만 여러 개면 생략)
    if (Lc) grid.push([text(valueCaption, 'valueCaption'), ...Array(labelCols - 1).fill(null).map(() => text('', 'corner')), text(d.colCaption ?? '열 레이블', 'colHead'), ...colLeaves.slice(1).map(() => text('', 'colHead'))]);
    for (let lvl = 0; lvl < colLevels; lvl++) {
      const row = lvl === colLevels - 1 ? rowHeaderCells() : Array.from({ length: labelCols }, () => text('', 'corner'));
      let prev = null;
      colLeaves.forEach((leaf) => {
        const lab = leaf.labels[lvl];
        // 같은 그룹의 반복 글자는 첫 칸에만 (엑셀과 같음)
        const groupKey = leaf.labels.slice(0, lvl + 1).join('\u0001');
        const show = lab !== undefined && (lvl === colLevels - 1 || groupKey !== prev);
        prev = groupKey;
        const role = leaf.kind === 'grand' ? `grandHead:${Math.max(0, leaf.vi)}` : leaf.kind === 'sub' ? 'colSubHead' : multiV && lvl === colLevels - 1 ? `valueHead:${leaf.vi}` : `colItem:${lvl}`;
        row.push(text(show ? lab : '', role));
      });
      grid.push(row);
    }
  } else {
    grid.push([...rowHeaderCells(), ...(V ? [text(valueCaption, 'valueHead:0')] : [])]);
  }
  const firstDataRowRel = grid.length - pageRows.length - (pageRows.length ? 1 : 0);

  // 본문
  const dataRole = (leaf, rowKind) => {
    const vi = Math.max(0, leaf.vi);
    if (rowKind === 'grand') return `grandData:${vi}`;
    if (leaf.kind === 'grand') return `grandColData:${vi}`;
    if (rowKind === 'sub') return `subData:${vi}`;
    if (rowKind === 'group') return `groupData:${vi}`;
    if (leaf.kind === 'sub') return `colSubData:${vi}`;
    return `data:${vi}`;
  };
  const valueCells = (rpath, rowKind) => colLeaves.map((leaf) => val(cellValue(rpath, leaf.cp, leaf.vi), leaf.vi, dataRole(leaf, rowKind)));
  const labelsRow = (node, labelText, role) => {
    const cells = Array.from({ length: labelCols }, () => text('', role));
    const col = layout === 'compact' ? 0 : node.depth;
    cells[col] = text(labelText, role);
    const img = imageOfKey(node.key);
    if (img) { cells[col].raw = ''; cells[col].image = { src: img.src, alt: img.alt ?? '' }; }
    if (layout === 'compact' && node.depth > 0) cells[col].style.indent = node.depth;
    return cells;
  };
  const walkRows = (node, lastPath) => {
    const isLeaf = node.depth === Lr - 1;
    if (layout === 'tabular') {
      if (isLeaf) {
        // 조상 글자는 그룹의 첫 행에만
        const cells = Array.from({ length: labelCols }, (_, dd) => text('', `rowItem:${dd}`));
        const chain = [];
        for (let n = node; n && n.depth >= 0; n = n.parent) chain.unshift(n);
        chain.forEach((n, dd) => {
          if (lastPath.shown.has(n.path)) return;
          cells[dd] = text(itemText(n.key), `rowItem:${dd}`);
          const img = imageOfKey(n.key);
          if (img) { cells[dd].raw = ''; cells[dd].image = { src: img.src, alt: img.alt ?? '' }; }
          lastPath.shown.add(n.path);
        });
        grid.push([...cells, ...valueCells(node.path, 'item')]);
        rowItems.push({ kind: 'item', node, chain });
        return;
      }
      node.children.forEach((ch) => walkRows(ch, lastPath));
      if (d.subtotals) {
        grid.push([...labelsRow(node, `${itemText(node.key)} 요약`, `rowSub:${node.depth}`), ...valueCells(node.path, 'sub')]);
        rowItems.push({ kind: 'sub', node });
      }
      return;
    }
    const group = !isLeaf;
    const cells = labelsRow(node, itemText(node.key), group ? `rowGroup:${node.depth}` : `rowItem:${node.depth}`);
    grid.push([...cells, ...(group && !d.subtotals ? colLeaves.map((leaf) => text('', `groupData:${Math.max(0, leaf.vi)}`)) : valueCells(node.path, group ? 'group' : 'item'))]);
    rowItems.push({ kind: 'item', node });
    node.children.forEach((ch) => walkRows(ch, lastPath));
  };
  setParents(rowTree);
  if (Lr) rowTree.children.forEach((ch) => walkRows(ch, { shown: new Set() }));
  else if (V) grid.push([text(TOTAL, 'grandLabel'), ...valueCells('', 'grand')]);
  if (Lr && d.grandRows && V) {
    const cells = Array.from({ length: labelCols }, () => text('', 'grandLabel'));
    cells[0] = text(TOTAL, 'grandLabel');
    grid.push([...cells, ...valueCells('', 'grand')]);
    rowItems.push({ kind: 'grand' });
  }

  // 줄무늬 행 · 열 (피벗 스타일 옵션)
  if (opts.bandRows || opts.bandCols) {
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

  const width = labelCols + colLeaves.length;
  return {
    grid,
    meta: {
      header, rowIdx, colIdx, valIdx, values, labelCols, pageRows: pageRows.length ? pageRows.length + 1 : 0,
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
  let memo = chartMemo.get(rows);
  if (!memo) { memo = new Map(); chartMemo.set(rows, memo); }
  const fmts = fieldStyle ? (def.rows ?? []).map((f) => fieldStyle(f)?.numFmt ?? null) : [];
  const key = pivotDefKey(def) + JSON.stringify(def.styleDef ?? null) + JSON.stringify(fmts);
  if (!memo.has(key)) {
    if (memo.size > 60) memo.clear();
    memo.set(key, pivotChartDataRaw(rows, def, fieldStyle));
  }
  return memo.get(key);
}
function pivotChartDataRaw(rows, def, fieldStyle) {
  const { def: d, rows: r } = resolvePivot(rows, def);
  const { grid, meta } = computePivot(r, d);
  if (meta.empty) return { categories: [], series: [] };
  const body = meta.pageRows + meta.headerRows;
  const cats = [];
  const rowIdx = [];
  if (!d.rows.length) { cats.push(TOTAL); rowIdx.push(body); }
  meta.rowItems.forEach((it, i) => {
    if (it.kind !== 'item' || it.node.children.length) return;
    const chain = [];
    for (let n = it.node; n && n.depth >= 0; n = n.parent) {
      const st = typeof n.key === 'number' ? fieldStyle?.(d.rows[n.depth]) : null;
      chain.unshift(st?.numFmt && st.numFmt !== 'general' ? formatValue(n.key, st).text : itemText(n.key));
    }
    cats.push(chain.join(' / '));
    rowIdx.push(body + i);
  });
  const series = [];
  meta.colLeaves.forEach((leaf, j) => {
    if (leaf.kind !== 'item' || leaf.vi < 0) return;
    const name = leaf.labels.length ? leaf.labels.join(' - ') : valueName(meta.values[leaf.vi]);
    const col = meta.labelCols + j;
    series.push({
      name,
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
  const { def: d, rows: r } = resolvePivot(rows, def);
  return computePivot(r, d).grid;
}

/** GETPIVOTDATA: 값 필드 이름과 (필드, 항목) 쌍으로 값 찾기 */
export function pivotLookup(rows, def, dataField, pairs) {
  const { def: d, rows: filtered, header } = resolvePivot(rows, def);
  const lower = String(dataField).trim().toLowerCase();
  const vi = d.values.findIndex((v) => valueName(v).trim().toLowerCase() === lower || v.field.toLowerCase() === lower);
  if (vi < 0) return null;
  const want = pairs.map(([f, item]) => [header.findIndex((h) => h.toLowerCase() === String(f).toLowerCase()), itemText(item)]);
  if (want.some(([i]) => i < 0)) return null;
  // 행·열 필드가 아닌 필드로 묻는 것은 엑셀에서 #REF!
  const fields = new Set([...d.rows, ...d.cols].map((n) => header.indexOf(n)));
  if (want.some(([i]) => !fields.has(i))) return null;
  const data = filtered.slice(1).filter((r) => want.every(([i, t]) => itemText(r[i]) === t));
  if (!data.length) return null;
  const m = makeMeasures(header, d.values, d.calcFields);
  const list = m.newList();
  for (const r of data) m.accumulate(list, r);
  const v = m.value(list, vi);
  return typeof v === 'number' ? v : null;
}
