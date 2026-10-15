// 논리 · 정보 · LAMBDA 도우미 · 부분합/집계 함수 (DOM 없음)
// lazy 함수: fn(인수 AST, ctx, ev) — ev.evaluate / ev.value / ev.evalRef / ev.deref / ev.call
import {
  ERR, ERROR_TYPE, Range, RefValue, Lambda, isError, scalar, toNum, toStr, toInt, toBool, asRange,
  lift, lazy, refFn, attempt, broadcast2,
} from './fxcore.js';
import { STAT } from './fx-stat.js';
import { MATH } from './fx-math.js';

const isMulti = (v) => v instanceof Range && (v.height > 1 || v.width > 1);
const isEmptyNode = (n) => !n || n.type === 'empty';

/** 조건 배열에 따라 두 값을 원소별로 고름 */
function pickEach(cond, a, b) {
  return broadcast2(cond, broadcast2(a, b, (x, y) => [x, y]), (c, pair) => {
    if (isError(c)) return c;
    const [x, y] = Array.isArray(pair) ? pair : [pair, pair];
    return (toBool(c ?? false) ? x : y) ?? 0; // 배열 IF 의 빈 셀은 0 (엑셀: MIN(IF(…, 빈 칸)) = 0)
  });
}

/** 결과가 1x1 배열이면 스칼라로 */
const one = (v) => (v instanceof Range && v.height === 1 && v.width === 1 ? v.rows[0][0] : v);

function boolsOf(args) {
  const out = [];
  for (const a of args) {
    if (a instanceof Range) {
      for (const v of a.values()) {
        if (isError(v)) throw v;
        if (typeof v === 'boolean') out.push(v);
        else if (typeof v === 'number') out.push(v !== 0);
      }
    } else if (a !== null && a !== undefined) {
      if (typeof a === 'string') {
        const u = a.toUpperCase();
        if (u !== 'TRUE' && u !== 'FALSE') throw ERR.VALUE;
      }
      out.push(toBool(a));
    }
  }
  if (!out.length) throw ERR.VALUE;
  return out;
}

const typeOf = (v) => {
  if (v instanceof Range) return v.height === 1 && v.width === 1 ? typeOf(v.rows[0][0]) : 64;
  if (v instanceof Lambda) return 128;
  if (isError(v)) return 16;
  if (typeof v === 'boolean') return 4;
  if (typeof v === 'string') return 2;
  return 1;
};

/** SUBTOTAL 번호 → 함수 */
const SUBTOTAL_FN = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEVP', 9: 'SUM', 10: 'VAR', 11: 'VARP' };
const AGGREGATE_FN = {
  1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV.S', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR.S',
  11: 'VAR.P', 12: 'MEDIAN', 13: 'MODE.SNGL', 14: 'LARGE', 15: 'SMALL', 16: 'PERCENTILE.INC', 17: 'QUARTILE.INC',
  18: 'PERCENTILE.EXC', 19: 'QUARTILE.EXC',
};
const callAgg = (name, args) => (STAT[name] ?? MATH[name])(args);
const NESTED_RE = /^=\s*(?:_xlfn\.)?(SUBTOTAL|AGGREGATE)\s*\(/i;

/** 참조 → 값 배열 (숨긴 행 · 중첩 부분합 · 오류 제외 옵션) */
function refRows(ref, ctx, ev, { skipManual, skipFiltered = true, skipNested, skipErrors }) {
  if (!(ref instanceof RefValue)) {
    const v = ref instanceof Range ? ref : asRange(ref);
    if (!skipErrors) return v;
    return new Range(v.rows.map((row) => row.map((x) => (isError(x) ? null : x))));
  }
  const full = ev.deref(ref);
  const rg = asRange(full);
  const rows = [];
  rg.rows.forEach((row, i) => {
    const r = ref.r1 + i;
    if (ctx.rowHidden && ctx.rowHidden(ref.sheet, r, skipManual, skipFiltered)) return;
    rows.push(row.map((x, j) => {
      if (skipErrors && isError(x)) return null;
      if (skipNested && ctx.formulaText) {
        const f = ctx.formulaText(ref.sheet, r, ref.c1 + j);
        if (f && NESTED_RE.test(f)) return null;
      }
      return x;
    }));
  });
  return new Range(rows.length ? rows : [[null]]);
}

export const LOGIC = {
  IF: lazy((args, ctx, ev) => {
    if (!args.length || args.length > 3) throw ERR.VALUE;
    const cond = ev.evaluate(args[0], ctx);
    if (isMulti(cond)) {
      const a = args.length > 1 ? (isEmptyNode(args[1]) ? 0 : ev.value(args[1], ctx)) : true;
      const b = args.length > 2 ? (isEmptyNode(args[2]) ? 0 : ev.value(args[2], ctx)) : false;
      return pickEach(cond, a, b);
    }
    const t = toBool(one(cond) ?? false);
    const branch = t ? args[1] : args[2];
    if (branch === undefined) return t;
    if (isEmptyNode(branch)) return 0;
    return ev.evalRef(branch, ctx);
  }),
  IFERROR: lazy((args, ctx, ev) => {
    if (args.length !== 2) throw ERR.VALUE;
    const v = ev.value(args[0], ctx);
    const alt = () => (isEmptyNode(args[1]) ? 0 : ev.value(args[1], ctx));
    if (isMulti(v)) {
      if (![...v.values()].some(isError)) return v;
      const a = alt();
      return broadcast2(new Range(v.rows.map((row) => row.map((x) => ({ x })))), a, (w, y) => (isError(w.x) ? y : w.x));
    }
    const s = one(v);
    return isError(s) ? alt() : s ?? 0;
  }),
  IFNA: lazy((args, ctx, ev) => {
    if (args.length !== 2) throw ERR.VALUE;
    const v = ev.value(args[0], ctx);
    const alt = () => (isEmptyNode(args[1]) ? 0 : ev.value(args[1], ctx));
    if (isMulti(v)) {
      if (![...v.values()].some((x) => x === ERR.NA)) return v;
      const a = alt();
      return broadcast2(new Range(v.rows.map((row) => row.map((x) => ({ x })))), a, (w, y) => (w.x === ERR.NA ? y : w.x));
    }
    const s = one(v);
    return s === ERR.NA ? alt() : s ?? 0;
  }),
  IFS: lazy((args, ctx, ev) => {
    if (args.length < 2 || args.length % 2) throw ERR.VALUE;
    for (let i = 0; i < args.length; i += 2) {
      const c = one(ev.evaluate(args[i], ctx));
      if (isMulti(c)) throw ERR.VALUE;
      if (toBool(c ?? false)) return isEmptyNode(args[i + 1]) ? 0 : ev.evalRef(args[i + 1], ctx);
    }
    throw ERR.NA;
  }),
  SWITCH: lazy((args, ctx, ev) => {
    if (args.length < 3) throw ERR.VALUE;
    const v = one(ev.evaluate(args[0], ctx));
    const rest = args.slice(1);
    for (let i = 0; i + 1 < rest.length; i += 2) {
      const k = one(ev.evaluate(rest[i], ctx));
      const same = typeof v === 'string' && typeof k === 'string' ? v.toLowerCase() === k.toLowerCase() : v === k;
      if (same) return ev.evalRef(rest[i + 1], ctx);
    }
    if (rest.length % 2) return ev.evalRef(rest[rest.length - 1], ctx);
    throw ERR.NA;
  }),
  AND: (a) => boolsOf(a).every(Boolean),
  OR: (a) => boolsOf(a).some(Boolean),
  XOR: (a) => boolsOf(a).filter(Boolean).length % 2 === 1,
  NOT: lift(([v]) => !toBool(v ?? false)),
  TRUE: () => true,
  FALSE: () => false,

  // ───── LAMBDA 도우미 ─────
  MAP: (args, ctx, ev) => {
    const fn = args[args.length - 1];
    const arrays = args.slice(0, -1).map(asRange);
    if (!(fn instanceof Lambda) || !arrays.length) throw ERR.VALUE;
    const h = Math.max(...arrays.map((a) => a.height));
    const w = Math.max(...arrays.map((a) => a.width));
    const rows = [];
    for (let r = 0; r < h; r++) {
      const row = [];
      for (let c = 0; c < w; c++) {
        const vals = arrays.map((a) => a.rows[a.height === 1 ? 0 : r]?.[a.width === 1 ? 0 : c] ?? ERR.NA);
        row.push(one(attempt(() => ev.call(fn, vals))));
      }
      rows.push(row);
    }
    return new Range(rows);
  },
  REDUCE: (args, ctx, ev) => {
    const [init, arr, fn] = args.length === 2 ? [null, args[0], args[1]] : args;
    if (!(fn instanceof Lambda)) throw ERR.VALUE;
    let acc = init ?? 0;
    for (const v of asRange(arr).values()) acc = ev.call(fn, [acc, v]);
    return acc;
  },
  SCAN: (args, ctx, ev) => {
    const [init, arr, fn] = args.length === 2 ? [null, args[0], args[1]] : args;
    if (!(fn instanceof Lambda)) throw ERR.VALUE;
    let acc = init ?? 0;
    return new Range(asRange(arr).rows.map((row) => row.map((v) => {
      acc = one(attempt(() => ev.call(fn, [acc, v])));
      return acc;
    })));
  },
  BYROW: ([arr, fn], ctx, ev) => {
    if (!(fn instanceof Lambda)) throw ERR.VALUE;
    return new Range(asRange(arr).rows.map((row) => [one(attempt(() => ev.call(fn, [new Range([row])])))]));
  },
  BYCOL: ([arr, fn], ctx, ev) => {
    if (!(fn instanceof Lambda)) throw ERR.VALUE;
    const rows = asRange(arr).rows;
    return new Range([rows[0].map((_, c) => one(attempt(() => ev.call(fn, [new Range(rows.map((r) => [r[c]]))]))))]);
  },
  MAKEARRAY: ([h, w, fn], ctx, ev) => {
    const R = toInt(h);
    const C = toInt(w);
    if (R < 1 || C < 1 || R * C > 1e7) throw ERR.VALUE;
    if (!(fn instanceof Lambda)) throw ERR.VALUE;
    const rows = [];
    for (let r = 1; r <= R; r++) {
      const row = [];
      for (let c = 1; c <= C; c++) row.push(one(attempt(() => ev.call(fn, [r, c]))));
      rows.push(row);
    }
    return new Range(rows);
  },

  // ───── 정보 ─────
  ISBLANK: lift(([v]) => v === null || v === undefined),
  ISNUMBER: lift(([v]) => typeof v === 'number'),
  ISTEXT: lift(([v]) => typeof v === 'string'),
  ISNONTEXT: lift(([v]) => typeof v !== 'string'),
  ISLOGICAL: lift(([v]) => typeof v === 'boolean'),
  ISERROR: lift(([v]) => isError(v)),
  ISERR: lift(([v]) => isError(v) && v !== ERR.NA),
  ISNA: lift(([v]) => v === ERR.NA),
  ISEVEN: lift(([v]) => {
    if (typeof v === 'boolean') throw ERR.VALUE;
    return Math.trunc(Math.abs(toNum(v))) % 2 === 0;
  }),
  ISODD: lift(([v]) => {
    if (typeof v === 'boolean') throw ERR.VALUE;
    return Math.trunc(Math.abs(toNum(v))) % 2 === 1;
  }),
  ISREF: refFn(([v]) => v instanceof RefValue || Array.isArray(v)),
  ISFORMULA: refFn(([v], ctx) => {
    if (!(v instanceof RefValue)) throw ERR.VALUE;
    return !!ctx.formulaText?.(v.sheet, v.r1, v.c1);
  }),
  'ERROR.TYPE': lift(([v]) => {
    if (!isError(v)) throw ERR.NA;
    return ERROR_TYPE[v.code] ?? ERR.NA;
  }),
  N: lift(([v]) => {
    if (isError(v)) throw v;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    return 0;
  }),
  NA: () => { throw ERR.NA; },
  TYPE: ([v]) => typeOf(v),
  SHEET: refFn(([v], ctx) => {
    if (v === undefined || v === null) return (ctx.here?.si ?? 0) + 1;
    if (v instanceof RefValue) return ctx.sheetIndex(v.sheet) + 1;
    if (typeof v === 'string') {
      const i = ctx.sheetIndex(v);
      if (i < 0) throw ERR.NA;
      return i + 1;
    }
    throw ERR.VALUE;
  }),
  SHEETS: refFn(([v], ctx) => {
    if (v === undefined || v === null) return ctx.sheetCount?.() ?? 1;
    if (v instanceof RefValue || Array.isArray(v)) return 1;
    throw ERR.VALUE;
  }),
  CELL: refFn(([info, ref], ctx, ev) => {
    const what = toStr(ev.deref(info)).toLowerCase();
    let r = ref;
    if (r === undefined || r === null) {
      if (!ctx.here) throw ERR.VALUE;
      r = new RefValue(null, ctx.here.r, ctx.here.c);
    }
    if (!(r instanceof RefValue)) throw ERR.VALUE;
    const v = ev.deref(new RefValue(r.sheet, r.r1, r.c1));
    let col = '';
    for (let n = r.c1 + 1; n > 0; n = Math.floor((n - 1) / 26)) col = String.fromCharCode(65 + ((n - 1) % 26)) + col;
    switch (what) {
      case 'address': return (r.sheet ? `${ctx.quoteSheet?.(r.sheet) ?? r.sheet}!` : '') + `$${col}$${r.r1 + 1}`;
      case 'col': return r.c1 + 1;
      case 'row': return r.r1 + 1;
      case 'contents': return v ?? 0;
      case 'type': return v === null || v === undefined ? 'b' : typeof v === 'string' ? 'l' : 'v';
      case 'filename': return ctx.fileName?.() ?? '';
      case 'sheetname': return r.sheet ?? ctx.here?.sheet ?? '';
      case 'format': return ctx.cellFormat?.(r.sheet, r.r1, r.c1) ?? 'G';
      case 'width': return ctx.colWidthChars?.(r.sheet, r.c1) ?? 8;
      case 'protect': return 1;
      case 'prefix': return typeof v === 'string' ? "'" : '';
      case 'parentheses': return 0;
      case 'color': return 0;
      default: throw ERR.VALUE;
    }
  }),
  INFO: ([t]) => {
    switch (toStr(t).toLowerCase()) {
      case 'numfile': return 1;
      case 'osversion': return 'Web';
      case 'recalc': return '자동';
      case 'release': return '16.0';
      case 'system': return 'pcdos';
      case 'origin': return '$A:$A$1';
      case 'directory': return '';
      default: throw ERR.VALUE;
    }
  },

  // ───── 부분합 / 집계 ─────
  SUBTOTAL: refFn((args, ctx, ev) => {
    const code = toInt(ev.deref(args[0]));
    const name = SUBTOTAL_FN[code > 100 ? code - 100 : code];
    if (!name || args.length < 2) throw ERR.VALUE;
    const vals = args.slice(1).map((a) => refRows(a, ctx, ev, { skipManual: code > 100, skipNested: true }));
    return callAgg(name, vals);
  }),
  AGGREGATE: refFn((args, ctx, ev) => {
    const code = toInt(ev.deref(args[0]));
    const opt = args[1] === undefined || args[1] === null ? 0 : toInt(ev.deref(args[1]));
    const name = AGGREGATE_FN[code];
    if (!name || opt < 0 || opt > 7) throw ERR.VALUE;
    const hidden = opt === 1 || opt === 3 || opt === 5 || opt === 7;
    const o = { skipNested: opt <= 3, skipManual: hidden, skipFiltered: hidden, skipErrors: opt === 2 || opt === 3 || opt === 6 || opt === 7 };
    if (code >= 14) return callAgg(name, [refRows(args[2], ctx, ev, o), ev.deref(args[3])]);
    const vals = args.slice(2).map((a) => refRows(a, ctx, ev, o));
    return callAgg(name, vals);
  }),
};


// ───────────── GROUPBY · PIVOTBY (엑셀 365) ─────────────
const blankV = (v) => v === null || v === undefined || v === '';
const groupKey = (vals) => vals.map((v) => (blankV(v) ? '\u0000' : isError(v) ? `e${v.code}` : typeof v === 'string' ? `s${v.toLowerCase()}` : `${typeof v}${v}`)).join('\u0001');
const cmpKey = (a, b) => {
  const ra = blankV(a) ? 3 : typeof a === 'number' ? 0 : typeof a === 'string' ? 1 : 2;
  const rb = blankV(b) ? 3 : typeof b === 'number' ? 0 : typeof b === 'string' ? 1 : 2;
  if (ra !== rb) return ra - rb;
  if (ra === 3) return 0;
  if (ra === 1) return a.toLowerCase() < b.toLowerCase() ? -1 : a.toLowerCase() > b.toLowerCase() ? 1 : 0;
  return a < b ? -1 : a > b ? 1 : 0;
};
const TOTAL_LABEL = '총합계';

/** 머리글 자동 감지: 첫 행이 텍스트이고 둘째 행이 숫자면 머리글 */
function detectHeaders(vals) {
  const r = vals.rows;
  return r.length > 1 && r[0].every((v) => typeof v === 'string') && r[1].some((v) => typeof v === 'number');
}

function aggregate(fn, ev, rows, colIdx) {
  const col = new Range(rows.length ? rows.map((row) => [row[colIdx]]) : [[null]]);
  return one(attempt(() => ev.call(fn, [col])));
}

LOGIC.GROUPBY = (args, ctx, ev) => {
  const [rf, vf, fn, headersArg, totalArg, sortArg, filterArg] = args;
  if (!(fn instanceof Lambda)) throw ERR.VALUE;
  const R = asRange(rf);
  const V = asRange(vf);
  if (R.height !== V.height) throw ERR.VALUE;
  const hmode = headersArg === undefined || headersArg === null ? (detectHeaders(V) ? 3 : 0) : toInt(headersArg);
  const hasHead = hmode === 1 || hmode === 3;
  const showHead = hmode === 2 || hmode === 3;
  const start = hasHead ? 1 : 0;
  const filt = filterArg === undefined || filterArg === null ? null : asRange(filterArg);
  const groups = new Map();
  for (let i = start; i < R.height; i++) {
    if (filt && !toBool(filt.rows[i]?.[0] ?? false)) continue;
    const keys = R.rows[i];
    const k = groupKey(keys);
    if (!groups.has(k)) groups.set(k, { keys, rows: [] });
    groups.get(k).rows.push(V.rows[i]);
  }
  const nk = R.width;
  const nv = V.width;
  let out = [...groups.values()].map((g) => [...g.keys, ...Array.from({ length: nv }, (_, j) => aggregate(fn, ev, g.rows, j))]);
  const sorts = sortArg === undefined || sortArg === null ? Array.from({ length: nk }, (_, i) => i + 1) : [...asRange(sortArg).values()].map(toInt);
  out.sort((a, b) => {
    for (const s of sorts) {
      const i = Math.abs(s) - 1;
      const c = cmpKey(a[i], b[i]);
      if (c) return s < 0 ? -c : c;
    }
    return 0;
  });
  const depth = totalArg === undefined || totalArg === null ? 1 : toInt(totalArg);
  if (depth !== 0) {
    const all = [...groups.values()].flatMap((g) => g.rows);
    const total = [TOTAL_LABEL, ...Array(nk - 1).fill(null), ...Array.from({ length: nv }, (_, j) => aggregate(fn, ev, all, j))];
    if (depth < 0) out.unshift(total); else out.push(total);
  }
  if (showHead) {
    const head = hasHead ? [...R.rows[0], ...V.rows[0]] : [...Array.from({ length: nk }, (_, i) => `행 필드 ${i + 1}`), ...Array.from({ length: nv }, (_, i) => `값 ${i + 1}`)];
    out.unshift(head);
  }
  if (!out.length) throw ERR.CALC;
  return new Range(out.map((row) => row.map((v) => (v === undefined ? null : v))));
};

LOGIC.PIVOTBY = (args, ctx, ev) => {
  const [rf, cf, vf, fn, headersArg, rowTotal, rowSort, colTotal, colSort, filterArg] = args;
  if (!(fn instanceof Lambda)) throw ERR.VALUE;
  const R = asRange(rf);
  const C = asRange(cf);
  const V = asRange(vf);
  if (R.height !== V.height || C.height !== V.height) throw ERR.VALUE;
  const hmode = headersArg === undefined || headersArg === null ? (detectHeaders(V) ? 3 : 0) : toInt(headersArg);
  const hasHead = hmode === 1 || hmode === 3;
  const start = hasHead ? 1 : 0;
  const filt = filterArg === undefined || filterArg === null ? null : asRange(filterArg);
  const rowKeys = new Map();
  const colKeys = new Map();
  const cells = new Map();
  const idx = [];
  for (let i = start; i < V.height; i++) {
    if (filt && !toBool(filt.rows[i]?.[0] ?? false)) continue;
    const rk = groupKey(R.rows[i]);
    const ck = groupKey(C.rows[i]);
    if (!rowKeys.has(rk)) rowKeys.set(rk, R.rows[i]);
    if (!colKeys.has(ck)) colKeys.set(ck, C.rows[i]);
    const k = `${rk}\u0002${ck}`;
    if (!cells.has(k)) cells.set(k, []);
    cells.get(k).push(V.rows[i]);
    idx.push({ rk, ck, row: V.rows[i] });
  }
  const order = (map, sortArg) => {
    const list = [...map.entries()];
    const s = sortArg === undefined || sortArg === null ? 1 : toInt(sortArg);
    list.sort((a, b) => {
      for (let i = 0; i < a[1].length; i++) {
        const c = cmpKey(a[1][i], b[1][i]);
        if (c) return s < 0 ? -c : c;
      }
      return 0;
    });
    return list;
  };
  const rows = order(rowKeys, rowSort);
  const cols = order(colKeys, colSort);
  const nr = R.width;
  const rt = rowTotal === undefined || rowTotal === null ? 1 : toInt(rowTotal);
  const ct = colTotal === undefined || colTotal === null ? 1 : toInt(colTotal);
  const out = [];
  const head = [...Array(nr).fill(null), ...cols.map(([, k]) => k[0] ?? null)];
  if (ct) head.push(TOTAL_LABEL);
  out.push(head);
  for (const [rk, keys] of rows) {
    const line = [...keys];
    for (const [ck] of cols) line.push(cells.has(`${rk}\u0002${ck}`) ? aggregate(fn, ev, cells.get(`${rk}\u0002${ck}`), 0) : null);
    if (ct) line.push(aggregate(fn, ev, idx.filter((x) => x.rk === rk).map((x) => x.row), 0));
    out.push(line);
  }
  if (rt) {
    const line = [TOTAL_LABEL, ...Array(nr - 1).fill(null)];
    for (const [ck] of cols) line.push(aggregate(fn, ev, idx.filter((x) => x.ck === ck).map((x) => x.row), 0));
    if (ct) line.push(aggregate(fn, ev, idx.map((x) => x.row), 0));
    if (rt < 0) out.splice(1, 0, line); else out.push(line);
  }
  return new Range(out);
};

LOGIC.PERCENTOF = (args) => {
  const [sub, all] = args;
  const sum = (v) => [...asRange(v).values()].reduce((s, x) => {
    if (isError(x)) throw x;
    return s + (typeof x === 'number' ? x : 0);
  }, 0);
  const d = sum(all);
  if (d === 0) throw ERR.DIV0;
  return sum(sub) / d;
};
