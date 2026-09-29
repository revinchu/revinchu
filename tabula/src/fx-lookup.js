// 찾기/참조 · 동적 배열 함수 (DOM 없음)
// refs 표시 함수는 참조 인수를 RefValue 그대로 받고(ev.deref 로 값), RefValue 를 돌려줄 수 있음
import {
  ERR, Range, RefValue, isError, scalar, toNum, toStr, toInt, toBool, optInt, optBool, asRange, lift, attempt,
  compareValues, typeRank, wildcardRegex, refFn, CellImage,
} from './fxcore.js';

const MAX_ROW = 10_000_000;
const MAX_COL = 16384;

// ───────────── 찾기 도우미 ─────────────
const blank = (v) => v === null || v === undefined;

function keyOf(key) {
  const k = scalar(key);
  return k === null || k === undefined ? 0 : k;
}

/** 정규식 (엑셀 REGEXTEST 와 같이 부분 일치, 기본은 대소문자 구분) */
export function lookupRegex(pattern, ci = false) {
  try {
    return new RegExp(toStr(pattern), ci ? 'iu' : 'u');
  } catch {
    throw ERR.VALUE;
  }
}
const regexText = (v) => (typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v));

/** 정확히 일치 (문자열은 대소문자 무시, wildcard 면 * ? ~ 지원, regex 면 정규식) */
export function exactIndex(vals, key, { wildcard = true, reverse = false, regex = null } = {}) {
  const k = keyOf(key);
  let test;
  if (regex) {
    test = (v) => !blank(v) && !isError(v) && regex.test(regexText(v));
  } else if (typeof k === 'string' && wildcard && /[*?~]/.test(k)) {
    const re = wildcardRegex(k);
    test = (v) => typeof v === 'string' && re.test(v);
  } else if (typeof k === 'string') {
    const lk = k.toLowerCase();
    test = (v) => typeof v === 'string' && v.toLowerCase() === lk;
  } else {
    test = (v) => !blank(v) && typeRank(v) === typeRank(k) && v === k;
  }
  if (reverse) {
    for (let i = vals.length - 1; i >= 0; i--) if (test(vals[i])) return i;
    return -1;
  }
  for (let i = 0; i < vals.length; i++) if (test(vals[i])) return i;
  return -1;
}

/** 이진 검색 (엑셀과 같이 정렬되어 있다고 가정). desc=false: key 이하 중 마지막, desc=true: key 이상 중 마지막 */
export function binaryIndex(vals, key, desc = false) {
  const k = keyOf(key);
  const rank = typeRank(k);
  let lo = 0;
  let hi = vals.length - 1;
  let found = -1;
  while (lo <= hi) {
    let mid = (lo + hi) >> 1;
    // 형식이 다른 값(빈 셀 포함)은 건너뜀
    let m = mid;
    while (m <= hi && (blank(vals[m]) || isError(vals[m]) || typeRank(vals[m]) !== rank)) m++;
    if (m > hi) {
      m = mid - 1;
      while (m >= lo && (blank(vals[m]) || isError(vals[m]) || typeRank(vals[m]) !== rank)) m--;
      if (m < lo) break;
    }
    mid = m;
    const c = compareValues(vals[mid], k);
    if (desc ? c >= 0 : c <= 0) { found = mid; lo = mid + 1; } else hi = mid - 1;
  }
  return found;
}

/** XLOOKUP/XMATCH 공통: matchMode 0 정확, -1 작거나 같은, 1 크거나 같은, 2 와일드카드, 3 정규식. searchMode 1, -1, 2, -2 */
function xIndex(vals, key, matchMode, searchMode) {
  const k = keyOf(key);
  // 정규식은 정렬 여부와 관계없이 차례로 찾음 (-2/2 는 방향만)
  if (matchMode === 3) return exactIndex(vals, k, { regex: lookupRegex(k), reverse: searchMode < 0 });
  if (searchMode === 2 || searchMode === -2) {
    const desc = searchMode === -2;
    // 정렬된 배열: 이진 검색
    if (matchMode === 0 || matchMode === 2) {
      const i = binaryIndex(vals, k, desc);
      return i >= 0 && compareValues(vals[i], k) === 0 ? i : -1;
    }
    const i = binaryIndex(vals, k, desc);
    if (i >= 0 && compareValues(vals[i], k) === 0) return i;
    if (matchMode === -1) return desc ? (i + 1 < vals.length ? i + 1 : -1) : i;
    return desc ? i : (i + 1 < vals.length ? i + 1 : -1);
  }
  const reverse = searchMode === -1;
  if (matchMode === 0 || matchMode === 2) return exactIndex(vals, k, { wildcard: matchMode === 2, reverse });
  const exact = exactIndex(vals, k, { wildcard: false, reverse });
  if (exact >= 0) return exact;
  let best = -1;
  const rank = typeRank(k);
  const order = reverse ? [...vals.keys()].reverse() : vals.keys();
  for (const i of order) {
    const v = vals[i];
    if (blank(v) || isError(v) || typeRank(v) !== rank) continue;
    const c = compareValues(v, k);
    if (matchMode === -1 && c < 0 && (best < 0 || compareValues(v, vals[best]) > 0)) best = i;
    if (matchMode === 1 && c > 0 && (best < 0 || compareValues(v, vals[best]) < 0)) best = i;
  }
  return best;
}

function vectorOf(v) {
  const r = asRange(v);
  if (r.height !== 1 && r.width !== 1) throw ERR.VALUE;
  return r.height === 1 ? r.rows[0] : r.rows.map((row) => row[0]);
}

// ───────────── 참조 도우미 ─────────────
const refHeight = (ref) => ref.r2 - ref.r1 + 1;
const refWidth = (ref) => ref.c2 - ref.c1 + 1;

function subRef(ref, r1, c1, r2, c2) {
  if (r1 < 0 || c1 < 0 || r2 >= MAX_ROW || c2 >= MAX_COL) throw ERR.REF;
  return new RefValue(ref.sheet, r1, c1, r2, c2);
}

/** 인수 → 값 (RefValue 면 셀 값으로) */
const val = (v, ev) => (v instanceof RefValue ? ev.deref(v) : v);

/** 2차원 배열의 행 목록 */
const rowsOf = (v, ev) => asRange(val(v, ev)).rows;

function fromRows(rows) {
  if (!rows.length || !rows[0].length) throw ERR.CALC;
  return new Range(rows);
}

const transpose = (rows) => (rows[0] ?? []).map((_, c) => rows.map((row) => row[c]));

/** 크기가 다른 배열을 쌓을 때 빈 칸은 #N/A */
const padRow = (row, w) => (row.length >= w ? row : row.concat(Array(w - row.length).fill(ERR.NA)));

function sortKeyCompare(a, b) {
  // 정렬 순서: 숫자 < 텍스트 < 논리값 < 오류 < 빈 칸
  const ra = blank(a) || a === '' ? 5 : isError(a) ? 4 : typeRank(a);
  const rb = blank(b) || b === '' ? 5 : isError(b) ? 4 : typeRank(b);
  if (ra !== rb) return ra - rb;
  if (ra >= 4) return 0;
  return compareValues(a, b);
}

function valueKey(v) {
  if (blank(v)) return '\u0000';
  if (isError(v)) return `e${v.code}`;
  if (typeof v === 'string') return `s${v.toLowerCase()}`;
  if (typeof v === 'boolean') return `b${v}`;
  return `n${v}`;
}

function takeRows(rows, n, fromEnd) {
  if (n === 0) throw ERR.CALC;
  const k = Math.min(Math.abs(n), rows.length);
  return n > 0 !== !!fromEnd ? rows.slice(0, k) : rows.slice(rows.length - k);
}

// ───────────── 함수 ─────────────
export const LOOKUP = {
  VLOOKUP: lift(([key, table, col, approx]) => {
    const t = asRange(table);
    const ci = toInt(col);
    if (ci < 1) throw ERR.VALUE;
    if (ci > t.width) throw ERR.REF;
    const first = t.rows.map((row) => row[0]);
    const exact = approx !== undefined && (approx === null || !toBool(approx)); // 빈 인수(,)는 FALSE (엑셀)
    const i = exact ? exactIndex(first, key) : binaryIndex(first, key);
    if (i < 0) throw ERR.NA;
    return t.rows[i][ci - 1] ?? 0;
  }, [0, 2, 3]),
  HLOOKUP: lift(([key, table, row, approx]) => {
    const t = asRange(table);
    const ri = toInt(row);
    if (ri < 1) throw ERR.VALUE;
    if (ri > t.height) throw ERR.REF;
    const exact = approx !== undefined && (approx === null || !toBool(approx)); // 빈 인수(,)는 FALSE (엑셀)
    const i = exact ? exactIndex(t.rows[0], key) : binaryIndex(t.rows[0], key);
    if (i < 0) throw ERR.NA;
    return t.rows[ri - 1][i] ?? 0;
  }, [0, 2, 3]),
  // 위셀 확장: 정규식으로 첫 열에서 찾기 (VLOOKUP 과 같은 모양, 일치하는 첫 행)
  REGEXVLOOKUP: lift(([pattern, table, col, ci]) => {
    const t = asRange(table);
    const ci1 = toInt(col);
    if (ci1 < 1) throw ERR.VALUE;
    if (ci1 > t.width) throw ERR.REF;
    const i = exactIndex(t.rows.map((row) => row[0]), pattern, { regex: lookupRegex(pattern, optInt(ci, 0) === 1) });
    if (i < 0) throw ERR.NA;
    return t.rows[i][ci1 - 1] ?? 0;
  }, [0, 2, 3]),
  REGEXHLOOKUP: lift(([pattern, table, row, ci]) => {
    const t = asRange(table);
    const ri = toInt(row);
    if (ri < 1) throw ERR.VALUE;
    if (ri > t.height) throw ERR.REF;
    const i = exactIndex(t.rows[0], pattern, { regex: lookupRegex(pattern, optInt(ci, 0) === 1) });
    if (i < 0) throw ERR.NA;
    return t.rows[ri - 1][i] ?? 0;
  }, [0, 2, 3]),
  // 위셀 확장: MATCH 의 정규식판 (1부터 시작하는 위치)
  REGEXMATCHPOS: lift(([pattern, look, ci]) => {
    const i = exactIndex(vectorOf(look), pattern, { regex: lookupRegex(pattern, optInt(ci, 0) === 1) });
    if (i < 0) throw ERR.NA;
    return i + 1;
  }, [0, 2]),
  LOOKUP: lift(([key, look, result]) => {
    const l = asRange(look);
    let keys;
    let res;
    if (result !== undefined && result !== null) {
      keys = vectorOf(l);
      res = vectorOf(result);
    } else if (l.width > l.height) {
      keys = l.rows[0];
      res = l.rows[l.height - 1];
    } else {
      keys = l.rows.map((r) => r[0]);
      res = l.rows.map((r) => r[r.length - 1]);
    }
    const i = binaryIndex(keys, key);
    if (i < 0 || i >= res.length) throw ERR.NA;
    return res[i] ?? 0;
  }, [0]),
  MATCH: lift(([key, look, type]) => {
    const vals = vectorOf(look);
    const mt = type === null ? 0 : optInt(type, 1); // MATCH(x, r,) = 정확히 일치
    const i = mt === 0 ? exactIndex(vals, key) : binaryIndex(vals, key, mt < 0);
    if (i < 0) throw ERR.NA;
    return i + 1;
  }, [0, 2]),
  XMATCH: lift(([key, look, mm, sm]) => {
    const vals = vectorOf(look);
    const matchMode = optInt(mm, 0);
    const searchMode = optInt(sm, 1);
    if (![0, -1, 1, 2, 3].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) throw ERR.VALUE;
    const i = xIndex(vals, key, matchMode, searchMode);
    if (i < 0) throw ERR.NA;
    return i + 1;
  }, [0, 2, 3]),
  XLOOKUP: refFn((args, ctx, ev) => {
    const [key, look, ret, notFound, mm, sm] = args;
    const l = asRange(val(look, ev));
    const retVal = val(ret, ev);
    const r = asRange(retVal);
    const byRow = l.width === 1;
    if (l.height !== 1 && l.width !== 1) throw ERR.VALUE;
    if (byRow ? r.height !== l.height : r.width !== l.width) throw ERR.VALUE;
    const vals = byRow ? l.rows.map((x) => x[0]) : l.rows[0];
    const matchMode = optInt(val(mm, ev), 0);
    const searchMode = optInt(val(sm, ev), 1);
    if (![0, -1, 1, 2, 3].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) throw ERR.VALUE;
    const one = (k) => {
      const i = xIndex(vals, k, matchMode, searchMode);
      if (i < 0) {
        if (notFound !== undefined && notFound !== null) return val(notFound, ev);
        throw ERR.NA;
      }
      if (ret instanceof RefValue) {
        return byRow ? subRef(ret, ret.r1 + i, ret.c1, ret.r1 + i, ret.c2) : subRef(ret, ret.r1, ret.c1 + i, ret.r2, ret.c1 + i);
      }
      return new Range(byRow ? [r.rows[i]] : r.rows.map((row) => [row[i]]));
    };
    const k = val(key, ev);
    if (k instanceof Range && (k.height > 1 || k.width > 1)) {
      return new Range(k.rows.map((row) => row.map((x) => {
        if (isError(x)) return x;
        const res = attempt(() => one(x));
        const v = res instanceof RefValue ? ev.deref(res) : res;
        return v instanceof Range ? v.rows[0][0] : v;
      })));
    }
    return one(k);
  }),
  INDEX: refFn((args, ctx, ev) => {
    const [src, row, col, area] = args;
    let base = src;
    if (Array.isArray(base)) {
      // 참조 합집합 (A1:B2,C3:D4)
      const a = optInt(val(area, ev), 1);
      if (a < 1 || a > base.length) throw ERR.REF;
      base = base[a - 1];
    }
    const rv = val(row, ev);
    const cv = val(col, ev);
    const colOmitted = col === undefined;
    const pick = (rArg, cArg) => {
      const isRef = base instanceof RefValue;
      const h = isRef ? refHeight(base) : asRange(base).height;
      const w = isRef ? refWidth(base) : asRange(base).width;
      let r = rArg === undefined || rArg === null ? 0 : toInt(rArg);
      let c;
      if (colOmitted) {
        if (h === 1 && w > 1) { c = r; r = 1; } else c = w === 1 ? 1 : 0;
      } else c = cArg === undefined || cArg === null ? 0 : toInt(cArg);
      if (r < 0 || c < 0 || r > h || c > w) throw ERR.REF;
      if (isRef) {
        const r1 = r ? base.r1 + r - 1 : base.r1;
        const r2 = r ? r1 : base.r2;
        const c1 = c ? base.c1 + c - 1 : base.c1;
        const c2 = c ? c1 : base.c2;
        return new RefValue(base.sheet, r1, c1, r2, c2);
      }
      const rows = asRange(base).rows;
      if (r && c) return rows[r - 1][c - 1] ?? 0;
      if (r) return new Range([rows[r - 1]]);
      if (c) return new Range(rows.map((x) => [x[c - 1]]));
      return new Range(rows);
    };
    const multi = (x) => x instanceof Range && (x.height > 1 || x.width > 1);
    if (multi(rv) || multi(cv)) {
      const R = asRange(rv ?? null);
      const C = asRange(cv ?? null);
      const h = Math.max(R.height, C.height);
      const w = Math.max(R.width, C.width);
      const out = [];
      for (let i = 0; i < h; i++) {
        const line = [];
        for (let j = 0; j < w; j++) {
          const a = R.rows[R.height === 1 ? 0 : i]?.[R.width === 1 ? 0 : j];
          const b = C.rows[C.height === 1 ? 0 : i]?.[C.width === 1 ? 0 : j];
          const res = attempt(() => pick(a, b));
          const v = res instanceof RefValue ? ev.deref(res) : res;
          line.push(v instanceof Range ? v.rows[0][0] : v);
        }
        out.push(line);
      }
      return new Range(out);
    }
    return pick(scalar(rv ?? null), scalar(cv ?? null));
  }),
  OFFSET: refFn((args, ctx, ev) => {
    const [ref, rows, cols, height, width] = args;
    if (!(ref instanceof RefValue)) throw ERR.VALUE;
    const dr = toInt(val(rows, ev));
    const dc = toInt(val(cols, ev));
    const h = height === undefined || height === null ? refHeight(ref) : toInt(val(height, ev));
    const w = width === undefined || width === null ? refWidth(ref) : toInt(val(width, ev));
    if (h === 0 || w === 0) throw ERR.REF;
    const r1 = ref.r1 + dr;
    const c1 = ref.c1 + dc;
    const r2 = h > 0 ? r1 + h - 1 : r1 + h + 1;
    const c2 = w > 0 ? c1 + w - 1 : c1 + w + 1;
    return subRef(ref, Math.min(r1, r2), Math.min(c1, c2), Math.max(r1, r2), Math.max(c1, c2));
  }),
  INDIRECT: refFn((args, ctx, ev) => {
    const text = toStr(val(args[0], ev));
    const a1 = args[1] === undefined || args[1] === null ? true : toBool(val(args[1], ev));
    const ref = ev.refFromText(text, a1, ctx);
    if (!ref) throw ERR.REF;
    return ref;
  }),
  CHOOSE: refFn((args, ctx, ev) => {
    const [i, ...opts] = args;
    const idx = val(i, ev);
    const choose = (k) => {
      const n = toInt(k);
      if (n < 1 || n > opts.length) throw ERR.VALUE;
      const o = opts[n - 1];
      return o === undefined ? 0 : o;
    };
    if (idx instanceof Range && (idx.height > 1 || idx.width > 1)) {
      const vals = opts.map((o) => val(o, ev));
      return new Range(idx.rows.map((row, r) => row.map((k, c) => {
        if (isError(k)) return k;
        return attempt(() => {
          const n = toInt(k);
          if (n < 1 || n > vals.length) throw ERR.VALUE;
          const v = vals[n - 1];
          if (!(v instanceof Range)) return v ?? 0;
          const rr = v.height === 1 ? 0 : r;
          const cc = v.width === 1 ? 0 : c;
          return v.rows[rr]?.[cc] ?? ERR.NA;
        });
      })));
    }
    return choose(idx);
  }),
  ROW: refFn((args, ctx) => {
    const ref = args[0];
    if (ref === undefined || ref === null) {
      if (!ctx.here) throw ERR.VALUE;
      return ctx.here.r + 1;
    }
    if (!(ref instanceof RefValue)) throw ERR.VALUE;
    if (refHeight(ref) === 1) return ref.r1 + 1;
    const h = Math.min(refHeight(ref), ctx.maxRows?.(ref.sheet) ?? refHeight(ref));
    return new Range(Array.from({ length: Math.max(1, h) }, (_, i) => [ref.r1 + i + 1]));
  }),
  COLUMN: refFn((args, ctx) => {
    const ref = args[0];
    if (ref === undefined || ref === null) {
      if (!ctx.here) throw ERR.VALUE;
      return ctx.here.c + 1;
    }
    if (!(ref instanceof RefValue)) throw ERR.VALUE;
    if (refWidth(ref) === 1) return ref.c1 + 1;
    return new Range([Array.from({ length: refWidth(ref) }, (_, i) => ref.c1 + i + 1)]);
  }),
  ROWS: refFn(([v], ctx, ev) => {
    if (v instanceof RefValue) return refHeight(v);
    if (Array.isArray(v)) throw ERR.REF;
    return asRange(val(v, ev)).height;
  }),
  COLUMNS: refFn(([v], ctx, ev) => {
    if (v instanceof RefValue) return refWidth(v);
    if (Array.isArray(v)) throw ERR.REF;
    return asRange(val(v, ev)).width;
  }),
  AREAS: refFn(([v]) => {
    if (Array.isArray(v)) return v.length;
    if (v instanceof RefValue) return 1;
    throw ERR.VALUE;
  }),
  ADDRESS: lift(([row, col, abs, a1, sheet]) => {
    const r = toInt(row);
    const c = toInt(col);
    const mode = optInt(abs, 1);
    if (r < 1 || c < 1 || c > MAX_COL || r > MAX_ROW || mode < 1 || mode > 4) throw ERR.VALUE;
    const useA1 = optBool(a1, true);
    let out;
    if (useA1) {
      let name = '';
      let n = c;
      while (n > 0) { const m = (n - 1) % 26; name = String.fromCharCode(65 + m) + name; n = Math.floor((n - 1) / 26); }
      const ra = mode === 1 || mode === 2 ? '$' : '';
      const ca = mode === 1 || mode === 3 ? '$' : '';
      out = `${ca}${name}${ra}${r}`;
    } else {
      const rr = mode === 1 || mode === 2 ? `R${r}` : `R[${r}]`;
      const cc = mode === 1 || mode === 3 ? `C${c}` : `C[${c}]`;
      out = rr + cc;
    }
    if (sheet !== undefined && sheet !== null) {
      const s = toStr(sheet);
      if (s !== '') out = `${/^[A-Za-z_][\w.]*$/.test(s) ? s : `'${s.replace(/'/g, "''")}'`}!${out}`;
    }
    return out;
  }),
  FORMULATEXT: refFn(([ref], ctx) => {
    if (!(ref instanceof RefValue)) throw ERR.NA;
    const f = ctx.formulaText?.(ref.sheet, ref.r1, ref.c1);
    if (!f) throw ERR.NA;
    return f;
  }),
  // IMAGE(원본, [대체 텍스트], [크기 조정], [높이], [너비]): 웹 주소(https)의 그림을 셀 안에 표시
  IMAGE: lift(([source, alt, sizing, height, width]) => {
    const src = toStr(source).trim();
    if (!/^(https?:|data:image\/)/i.test(src)) throw ERR.VALUE;
    const mode = sizing === undefined || sizing === null ? 0 : Math.trunc(toNum(sizing));
    if (mode < 0 || mode > 3) throw ERR.VALUE;
    const h = height === undefined || height === null ? null : toNum(height);
    const w = width === undefined || width === null ? null : toNum(width);
    if (mode === 3 && !h && !w) throw ERR.VALUE;
    return new CellImage({ src, alt: alt === undefined || alt === null ? '' : toStr(alt), sizing: mode, h, w });
  }),
  HYPERLINK: lift(([link, friendly]) => (friendly === undefined || friendly === null ? toStr(link) : scalar(friendly) ?? 0)),
  GETPIVOTDATA: refFn((args, ctx, ev) => {
    let [field, pt, ...pairs] = args;
    // 엑셀 2000 이전 형식 GETPIVOTDATA(피벗 범위, "값 필드") — 첫 인수가 참조이고 둘째가 글자
    if (field instanceof RefValue && !(pt instanceof RefValue)) [field, pt] = [pt, field];
    if (!(pt instanceof RefValue)) throw ERR.REF;
    if (!ctx.pivotData) throw ERR.REF;
    const items = [];
    for (let i = 0; i + 1 < pairs.length; i += 2) items.push([toStr(val(pairs[i], ev)), scalar(val(pairs[i + 1], ev))]);
    return ctx.pivotData(pt, toStr(val(field, ev)), items);
  }),

  // ───── 동적 배열 ─────
  TRANSPOSE: ([a]) => fromRows(transpose(asRange(a).rows)),
  FILTER: ([a, include, ifEmpty]) => {
    const rows = asRange(a).rows;
    const inc = asRange(include);
    const h = rows.length;
    const w = rows[0].length;
    let out;
    const truthy = (v) => {
      if (isError(v)) throw v;
      return toBool(v ?? false);
    };
    if (inc.width === 1 && inc.height === h) out = rows.filter((_, i) => truthy(inc.rows[i][0]));
    else if (inc.height === 1 && inc.width === w) {
      const keep = inc.rows[0].map(truthy);
      out = rows.map((row) => row.filter((_, j) => keep[j]));
      if (!out[0].length) out = [];
    } else throw ERR.VALUE;
    if (!out.length) {
      if (ifEmpty !== undefined && ifEmpty !== null) return ifEmpty instanceof Range ? ifEmpty : scalar(ifEmpty);
      throw ERR.CALC;
    }
    return new Range(out);
  },
  SORT: ([a, idx, order, byCol]) => {
    let rows = asRange(a).rows;
    const col = optBool(byCol, false);
    if (col) rows = transpose(rows);
    const idxs = idx === undefined || idx === null ? [1] : [...asRange(idx).values()].map(toInt);
    const ords = order === undefined || order === null ? [1] : [...asRange(order).values()].map(toInt);
    const w = rows[0].length;
    if (idxs.some((i) => i < 1 || i > w)) throw ERR.VALUE;
    if (ords.some((o) => o !== 1 && o !== -1)) throw ERR.VALUE;
    const sorted = rows.map((row, i) => ({ row, i })).sort((x, y) => {
      for (let k = 0; k < idxs.length; k++) {
        const c = sortKeyCompare(x.row[idxs[k] - 1], y.row[idxs[k] - 1]);
        const o = ords[Math.min(k, ords.length - 1)];
        if (c) {
          // 빈 칸은 순서와 관계없이 끝으로
          const bx = blank(x.row[idxs[k] - 1]);
          const by = blank(y.row[idxs[k] - 1]);
          if (bx !== by) return bx ? 1 : -1;
          return c * o;
        }
      }
      return x.i - y.i;
    }).map((x) => x.row);
    return new Range(col ? transpose(sorted) : sorted);
  },
  SORTBY: ([a, ...pairs]) => {
    const rows = asRange(a).rows;
    const keys = [];
    for (let i = 0; i < pairs.length; i += 2) {
      const by = asRange(pairs[i]);
      const o = pairs[i + 1] === undefined || pairs[i + 1] === null ? 1 : toInt(pairs[i + 1]);
      if (o !== 1 && o !== -1) throw ERR.VALUE;
      keys.push({ by, o });
    }
    if (!keys.length) throw ERR.VALUE;
    const vertical = keys[0].by.width === 1 && keys[0].by.height === rows.length;
    const n = vertical ? rows.length : rows[0].length;
    const keyAt = (k, i) => (vertical ? k.by.rows[i]?.[0] : k.by.rows[0]?.[i]);
    for (const k of keys) {
      const len = vertical ? k.by.height : k.by.width;
      if (len !== n || (vertical ? k.by.width : k.by.height) !== 1) throw ERR.VALUE;
    }
    const order = [...Array(n).keys()].sort((x, y) => {
      for (const k of keys) {
        const c = sortKeyCompare(keyAt(k, x), keyAt(k, y));
        if (c) return c * k.o;
      }
      return x - y;
    });
    if (vertical) return new Range(order.map((i) => rows[i]));
    return new Range(rows.map((row) => order.map((i) => row[i])));
  },
  UNIQUE: ([a, byCol, once]) => {
    let rows = asRange(a).rows;
    const col = optBool(byCol, false);
    if (col) rows = transpose(rows);
    const counts = new Map();
    const first = [];
    for (const row of rows) {
      const k = row.map(valueKey).join('\u0001');
      if (!counts.has(k)) { counts.set(k, 0); first.push([k, row]); }
      counts.set(k, counts.get(k) + 1);
    }
    let out = first.filter(([k]) => !optBool(once, false) || counts.get(k) === 1).map(([, row]) => row);
    if (!out.length) throw ERR.CALC;
    if (col) out = transpose(out);
    return new Range(out);
  },
  TAKE: ([a, r, c]) => {
    let rows = asRange(a).rows;
    if (r !== undefined && r !== null) rows = takeRows(rows, toInt(r));
    if (c !== undefined && c !== null) rows = transpose(takeRows(transpose(rows), toInt(c)));
    return fromRows(rows);
  },
  DROP: ([a, r, c]) => {
    let rows = asRange(a).rows;
    const drop = (list, n) => (n >= 0 ? list.slice(n) : list.slice(0, Math.max(0, list.length + n)));
    if (r !== undefined && r !== null) rows = drop(rows, toInt(r));
    if (c !== undefined && c !== null) rows = rows.map((row) => drop(row, toInt(c)));
    return fromRows(rows);
  },
  EXPAND: ([a, r, c, pad]) => {
    const rows = asRange(a).rows;
    const h = r === undefined || r === null ? rows.length : toInt(r);
    const w = c === undefined || c === null ? rows[0].length : toInt(c);
    if (h < rows.length || w < rows[0].length) throw ERR.VALUE;
    const fill = pad === undefined || pad === null ? ERR.NA : scalar(pad);
    const out = [];
    for (let i = 0; i < h; i++) {
      const row = rows[i] ? rows[i].slice() : [];
      while (row.length < w) row.push(fill);
      out.push(row);
    }
    return fromRows(out);
  },
  CHOOSECOLS: ([a, ...cols]) => {
    const rows = asRange(a).rows;
    const w = rows[0].length;
    const idx = cols.flatMap((c) => [...asRange(c).values()].map(toInt)).map((i) => (i < 0 ? w + i + 1 : i));
    if (!idx.length || idx.some((i) => i < 1 || i > w)) throw ERR.VALUE;
    return new Range(rows.map((row) => idx.map((i) => row[i - 1])));
  },
  CHOOSEROWS: ([a, ...rs]) => {
    const rows = asRange(a).rows;
    const h = rows.length;
    const idx = rs.flatMap((r) => [...asRange(r).values()].map(toInt)).map((i) => (i < 0 ? h + i + 1 : i));
    if (!idx.length || idx.some((i) => i < 1 || i > h)) throw ERR.VALUE;
    return new Range(idx.map((i) => rows[i - 1]));
  },
  VSTACK: (args) => {
    const parts = args.map((x) => asRange(x).rows);
    const w = Math.max(...parts.map((p) => p[0].length));
    return new Range(parts.flatMap((p) => p.map((row) => padRow(row, w))));
  },
  HSTACK: (args) => {
    const parts = args.map((x) => asRange(x).rows);
    const h = Math.max(...parts.map((p) => p.length));
    const out = [];
    for (let i = 0; i < h; i++) {
      out.push(parts.flatMap((p) => (p[i] ? p[i] : Array(p[0].length).fill(ERR.NA))));
    }
    return new Range(out);
  },
  TOCOL: ([a, ignore, byCol]) => {
    let rows = asRange(a).rows;
    if (optBool(byCol, false)) rows = transpose(rows);
    const vals = rows.flat().filter((v) => keepFor(v, optInt(ignore, 0)));
    return fromRows(vals.map((v) => [v]));
  },
  TOROW: ([a, ignore, byCol]) => {
    let rows = asRange(a).rows;
    if (optBool(byCol, false)) rows = transpose(rows);
    const vals = rows.flat().filter((v) => keepFor(v, optInt(ignore, 0)));
    return fromRows([vals]);
  },
  WRAPROWS: ([a, n, pad]) => wrap(asRange(a), toInt(n), pad, false),
  WRAPCOLS: ([a, n, pad]) => wrap(asRange(a), toInt(n), pad, true),
};

function keepFor(v, ignore) {
  if ((ignore === 1 || ignore === 3) && blank(v)) return false;
  if ((ignore === 2 || ignore === 3) && isError(v)) return false;
  return true;
}

function wrap(rg, n, pad, byCol) {
  if (rg.height !== 1 && rg.width !== 1) throw ERR.VALUE;
  if (n < 1) throw ERR.NUM;
  const vals = [...rg.values()];
  const fill = pad === undefined || pad === null ? ERR.NA : scalar(pad);
  const out = [];
  for (let i = 0; i < vals.length; i += n) {
    const chunk = vals.slice(i, i + n);
    while (chunk.length < n) chunk.push(fill);
    out.push(chunk);
  }
  return new Range(byCol ? transpose(out) : out);
}

// 값 도우미 (다른 모듈용)
export const derefArg = val;
export const rowsOfArg = rowsOf;
