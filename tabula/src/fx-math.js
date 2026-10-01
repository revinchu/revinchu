// 수학/삼각 · 조건부 집계 함수 (DOM 없음)
import {
  ERR, Range, isError, scalar, toNum, toInt, toBool, optNum, optInt, checkNum, collectNums, flat, asRange, matrix,
  makeCriteria, lift, fastCount, parseNumberText, r15,
} from './fxcore.js';
import { maxOf, minOf, CLOSED_BOOK } from './fxcore.js';

export function roundTo(n, digits, mode) {
  if (digits > 15) digits = 15;
  const f = 10 ** digits;
  const x = n * f;
  const eps = 1e-9 * Math.max(1, Math.abs(x));
  let r;
  if (mode === 'up') r = Math.sign(x) * Math.ceil(Math.abs(x) - eps);
  else if (mode === 'down') r = Math.sign(x) * Math.floor(Math.abs(x) + eps);
  else r = Math.sign(x) * Math.round(Math.abs(x) + eps);
  return Number((r / f).toPrecision(15));
}

const fact = (n) => {
  if (n < 0) throw ERR.NUM;
  let r = 1;
  for (let i = 2; i <= n; i++) r *= i;
  return checkNum(r);
};
const combin = (n, k) => {
  if (n < 0 || k < 0 || k > n) throw ERR.NUM;
  k = Math.min(k, n - k);
  let r = 1;
  for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i;
  return Math.round(r);
};
const gcd2 = (a, b) => { while (b) [a, b] = [b, a % b]; return a; };

function sumIfs(sumRange, pairs) {
  const s = asRange(sumRange);
  let total = 0;
  // 조건에 맞는 칸의 오류는 결과가 됨 (엑셀: 맞지 않는 칸의 오류는 무시)
  let err = null;
  s.rows.forEach((row, i) => row.forEach((x, j) => {
    if (err || (typeof x !== 'number' && !isError(x)) || !matchesAll(pairs, i, j)) return;
    if (isError(x)) err = x; else total += x;
  }));
  return err ?? total;
}

export function ifsPairs(args) {
  if (!args.length || args.length % 2 !== 0) throw ERR.VALUE;
  const pairs = [];
  for (let i = 0; i < args.length; i += 2) pairs.push({ range: asRange(args[i]), test: makeCriteria(args[i + 1]) });
  const h = pairs[0].range.height;
  const w = pairs[0].range.width;
  if (pairs.some((p) => p.range.height !== h || p.range.width !== w)) throw ERR.VALUE;
  return pairs;
}

export function matchesAll(pairs, r, c) {
  return pairs.every(({ range, test }) => {
    const row = range.rows[r];
    return row !== undefined && c < row.length && test(row[c]);
  });
}

// ─── *IFS 해시 색인: 같음 조건만 있으면 (범위 조합마다 한 번 묶어 두고) 조건 값으로 바로 찾음 ───
// =MINIFS(표[date], 표[week], [@week], 표[year], [@year]) 를 10만 행에 채운 열이 O(n²) → O(n)
const rowsIds = new WeakMap();
let nextRowsId = 1;
const rowsId = (rows) => { let i = rowsIds.get(rows); if (!i) { i = nextRowsId++; rowsIds.set(rows, i); } return i; };
const ifsIdxMemo = new WeakMap();
/** 조건 → 같음 비교 키 (크기 비교 · 와일드카드 · <> 는 null) */
function critKey(crit) {
  if (crit instanceof Range && crit.height === 1 && crit.width === 1) crit = crit.rows[0][0];
  if (isError(crit)) return `x${crit.code}`;
  crit = scalar(crit);
  if (typeof crit === 'number') return `n${r15(crit)}`;
  if (typeof crit === 'boolean') return crit ? 'b1' : 'b0';
  if (isError(crit)) return `x${crit.code}`;
  if (typeof crit !== 'string') return null;
  const m = /^(=)?([\s\S]*)$/.exec(crit);
  const operand = m[2];
  if (!m[1] && /^[<>]/.test(operand)) return null;
  if (/[*?~]/.test(operand)) return null;
  if (operand === '') return 'e';
  const num = parseNumberText(operand);
  if (num !== null) return `n${r15(num)}`;
  const up = operand.toUpperCase();
  if (up === 'TRUE' || up === 'FALSE') return up === 'TRUE' ? 'b1' : 'b0';
  return `s${operand.toLowerCase()}`;
}
/** 칸 값 → 같음 비교 키 (critKey 와 같은 규칙: 숫자 모양 글자는 숫자로) */
function cellKey(v) {
  if (typeof v === 'number') return `n${r15(v)}`;
  if (v === null || v === undefined || v === '') return 'e';
  if (typeof v === 'string') { const p = parseNumberText(v); return p !== null ? `n${r15(p)}` : `s${v.toLowerCase()}`; }
  if (typeof v === 'boolean') return v ? 'b1' : 'b0';
  if (isError(v)) return `x${v.code}`;
  return '\u0000';
}
/**
 * target(값 범위, 없으면 개수만) 과 조건 범위들 → 조건 키 조합별 { sum, n (숫자 수), rows (맞는 칸 수), min, max, err }
 * 조건이 모두 같음 비교가 아니거나 범위가 작으면 null
 */
function fastIfs(target, args) {
  if (!args.length || args.length % 2) return null;
  const keys = [];
  for (let i = 1; i < args.length; i += 2) { const k = critKey(args[i]); if (k === null) return null; keys.push(k); }
  const ranges = [];
  for (let i = 0; i < args.length; i += 2) { const r = args[i]; if (!(r instanceof Range)) return null; ranges.push(r); }
  const h = ranges[0].height;
  const w = ranges[0].width;
  if (h * w < 64) return null;
  if (ranges.some((r) => r.height !== h || r.width !== w)) return null;
  const t = target ? asRange(target) : null;
  if (t && (t.height !== h || t.width !== w)) return null;
  const anchor = (t ?? ranges[0]).rows;
  let byAnchor = ifsIdxMemo.get(anchor);
  if (!byAnchor) { byAnchor = new Map(); ifsIdxMemo.set(anchor, byAnchor); }
  const sig = `${t ? 't' : 'c'}${ranges.map((r) => rowsId(r.rows)).join(',')}`;
  let idx = byAnchor.get(sig);
  if (!idx) {
    idx = new Map();
    const rr = ranges.map((r) => r.rows);
    const tr = t?.rows;
    for (let i = 0; i < h; i++) {
      for (let j = 0; j < w; j++) {
        let k = cellKey(rr[0][i]?.[j] ?? null);
        for (let q = 1; q < rr.length; q++) k += `\u0001${cellKey(rr[q][i]?.[j] ?? null)}`;
        let e = idx.get(k);
        if (!e) { e = { sum: 0, n: 0, rows: 0, min: Infinity, max: -Infinity, err: null }; idx.set(k, e); }
        e.rows++;
        if (!tr) continue;
        const x = tr[i]?.[j];
        if (typeof x === 'number') { e.sum += x; e.n++; if (x < e.min) e.min = x; if (x > e.max) e.max = x; } else if (isError(x) && !e.err) e.err = x;
      }
    }
    byAnchor.set(sig, idx);
  }
  return idx.get(keys.join('\u0001')) ?? { sum: 0, n: 0, rows: 0, min: Infinity, max: -Infinity, err: null };
}

/** 범위 자리에 온 오류 값 (=SUMIFS(#REF!, #REF!, …) 는 #REF!) */
// 닫힌 외부 통합 문서의 범위([1]시트!A1:A9)는 엑셀도 *IF 함수에서 #VALUE! (값 캐시로 계산하지 않음)
const closedBook = (x) => x instanceof Range && /^\[\d+\]/.test(x.ref?.sheet ?? '');
// 닫힌 외부 통합 문서의 범위: #VALUE! (엑셀) — 다만 파일에 저장된 값이 있으면 다시 계산하기 전까지 그 값을 보임 (CLOSED_BOOK.hit → workbook)
const rangeError = (list) => list.find((x) => isError(x)) ?? (list.some(closedBook) ? ((CLOSED_BOOK.hit = true), ERR.VALUE) : null);

function ifsValues(target, pairs) {
  const s = asRange(target);
  if (pairs.length && (pairs[0].range.height !== s.height || pairs[0].range.width !== s.width)) throw ERR.VALUE;
  const out = [];
  s.rows.forEach((row, i) => row.forEach((x, j) => { if (typeof x === 'number' && matchesAll(pairs, i, j)) out.push(x); }));
  return out;
}

function det(m) {
  const n = m.length;
  const a = m.map((r) => r.slice());
  let d = 1;
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r;
    if (Math.abs(a[p][i]) < 1e-15) return 0;
    if (p !== i) { [a[p], a[i]] = [a[i], a[p]]; d = -d; }
    d *= a[i][i];
    for (let r = i + 1; r < n; r++) {
      const f = a[r][i] / a[i][i];
      for (let c = i; c < n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return d;
}

export function inverse(m) {
  const n = m.length;
  const a = m.map((r, i) => [...r, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let i = 0; i < n; i++) {
    let p = i;
    for (let r = i + 1; r < n; r++) if (Math.abs(a[r][i]) > Math.abs(a[p][i])) p = r;
    if (Math.abs(a[p][i]) < 1e-15) throw ERR.NUM;
    [a[p], a[i]] = [a[i], a[p]];
    const piv = a[i][i];
    for (let c = 0; c < 2 * n; c++) a[i][c] /= piv;
    for (let r = 0; r < n; r++) {
      if (r === i) continue;
      const f = a[r][i];
      for (let c = 0; c < 2 * n; c++) a[r][c] -= f * a[i][c];
    }
  }
  return a.map((r) => r.slice(n));
}

/** 숫자 행렬 (빈 칸·글자는 #VALUE!) */
export function numMatrix(v) {
  return matrix(v).map((row) => row.map((x) => {
    if (isError(x)) throw x;
    if (typeof x !== 'number') throw ERR.VALUE;
    return x;
  }));
}

export function mmult(a, b) {
  if (a[0].length !== b.length) throw ERR.VALUE;
  return a.map((row) => b[0].map((_, j) => row.reduce((s, x, k) => s + x * b[k][j], 0)));
}

const ROMAN = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];

const SCALAR = {
  ABS: ([n]) => Math.abs(toNum(n)),
  SIGN: ([n]) => Math.sign(toNum(n)),
  INT: ([n]) => Math.floor(toNum(n)),
  ROUND: ([n, d]) => roundTo(toNum(n), optInt(d, 0)),
  ROUNDUP: ([n, d]) => roundTo(toNum(n), optInt(d, 0), 'up'),
  ROUNDDOWN: ([n, d]) => roundTo(toNum(n), optInt(d, 0), 'down'),
  TRUNC: ([n, d]) => roundTo(toNum(n), optInt(d, 0), 'down'),
  MOD: ([n, d]) => {
    const x = toNum(n);
    const y = toNum(d);
    if (y === 0) throw ERR.DIV0;
    return Number((x - y * Math.floor(x / y)).toPrecision(15));
  },
  QUOTIENT: ([n, d]) => { const y = toNum(d); if (y === 0) throw ERR.DIV0; return Math.trunc(toNum(n) / y); },
  POWER: ([a, b]) => {
    const x = toNum(a);
    const y = toNum(b);
    if (x === 0 && y === 0) throw ERR.NUM;
    if (x === 0 && y < 0) throw ERR.DIV0;
    return checkNum(x ** y);
  },
  SQRT: ([n]) => { const x = toNum(n); if (x < 0) throw ERR.NUM; return Math.sqrt(x); },
  SQRTPI: ([n]) => { const x = toNum(n); if (x < 0) throw ERR.NUM; return Math.sqrt(x * Math.PI); },
  EXP: ([n]) => checkNum(Math.exp(toNum(n))),
  LN: ([n]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log(x); },
  LOG10: ([n]) => { const x = toNum(n); if (x <= 0) throw ERR.NUM; return Math.log10(x); },
  LOG: ([n, b]) => {
    const x = toNum(n);
    const base = optNum(b, 10);
    if (x <= 0 || base <= 0) throw ERR.NUM;
    if (base === 1) throw ERR.DIV0;
    return Math.log(x) / Math.log(base);
  },
  CEILING: ([n, s]) => {
    const x = toNum(n);
    const sig = optNum(s, 1);
    if (sig === 0) return 0;
    if (x > 0 && sig < 0) throw ERR.NUM;
    return roundTo(Math.ceil(x / sig - 1e-12) * sig, 12);
  },
  'CEILING.MATH': ([n, s, mode]) => {
    const x = toNum(n);
    const sig = Math.abs(optNum(s, 1));
    if (sig === 0) return 0;
    const away = optNum(mode, 0) !== 0 && x < 0;
    return roundTo((away ? Math.floor(x / sig + 1e-12) : Math.ceil(x / sig - 1e-12)) * sig, 12);
  },
  'CEILING.PRECISE': ([n, s]) => { const sig = Math.abs(optNum(s, 1)); if (!sig) return 0; return roundTo(Math.ceil(toNum(n) / sig - 1e-12) * sig, 12); },
  'ISO.CEILING': ([n, s]) => { const sig = Math.abs(optNum(s, 1)); if (!sig) return 0; return roundTo(Math.ceil(toNum(n) / sig - 1e-12) * sig, 12); },
  FLOOR: ([n, s]) => {
    const x = toNum(n);
    const sig = toNum(s);
    if (sig === 0) { if (x === 0) return 0; throw ERR.DIV0; }
    if (x > 0 && sig < 0) throw ERR.NUM;
    return roundTo(Math.floor(x / sig + 1e-12) * sig, 12);
  },
  'FLOOR.MATH': ([n, s, mode]) => {
    const x = toNum(n);
    const sig = Math.abs(optNum(s, 1));
    if (sig === 0) return 0;
    const toward = optNum(mode, 0) !== 0 && x < 0;
    return roundTo((toward ? Math.ceil(x / sig - 1e-12) : Math.floor(x / sig + 1e-12)) * sig, 12);
  },
  'FLOOR.PRECISE': ([n, s]) => { const sig = Math.abs(optNum(s, 1)); if (!sig) return 0; return roundTo(Math.floor(toNum(n) / sig + 1e-12) * sig, 12); },
  MROUND: ([n, m]) => {
    const x = toNum(n);
    const mul = toNum(m);
    if (mul === 0) return 0;
    if (Math.sign(x) * Math.sign(mul) < 0) throw ERR.NUM;
    return roundTo(Math.sign(x) * Math.round(Math.abs(x / mul) + 1e-12) * mul, 12);
  },
  EVEN: ([n]) => { const x = toNum(n); const r = Math.ceil(Math.abs(x) / 2) * 2; return x < 0 ? -r : r; },
  ODD: ([n]) => {
    const x = toNum(n);
    let r = Math.ceil(Math.abs(x));
    if (r % 2 === 0) r += 1;
    return x < 0 ? -r : r;
  },
  FACT: ([n]) => fact(Math.floor(toNum(n))),
  FACTDOUBLE: ([n]) => {
    const x = Math.floor(toNum(n));
    if (x < -1) throw ERR.NUM;
    let r = 1;
    for (let i = x; i > 1; i -= 2) r *= i;
    return r;
  },
  COMBIN: ([n, k]) => combin(Math.floor(toNum(n)), Math.floor(toNum(k))),
  COMBINA: ([n, k]) => {
    const a = Math.floor(toNum(n));
    const b = Math.floor(toNum(k));
    if (a < 0 || b < 0) throw ERR.NUM;
    return a === 0 && b === 0 ? 1 : combin(a + b - 1, b);
  },
  PERMUT: ([n, k]) => {
    const a = Math.floor(toNum(n));
    const b = Math.floor(toNum(k));
    if (a < 0 || b < 0 || b > a) throw ERR.NUM;
    let r = 1;
    for (let i = 0; i < b; i++) r *= a - i;
    return r;
  },
  PERMUTATIONA: ([n, k]) => { const a = Math.floor(toNum(n)); const b = Math.floor(toNum(k)); if (a < 0 || b < 0) throw ERR.NUM; return a ** b; },
  PI: () => Math.PI,
  RADIANS: ([n]) => (toNum(n) * Math.PI) / 180,
  DEGREES: ([n]) => (toNum(n) * 180) / Math.PI,
  SIN: ([n]) => Math.sin(toNum(n)),
  COS: ([n]) => Math.cos(toNum(n)),
  TAN: ([n]) => Math.tan(toNum(n)),
  ASIN: ([n]) => { const x = toNum(n); if (Math.abs(x) > 1) throw ERR.NUM; return Math.asin(x); },
  ACOS: ([n]) => { const x = toNum(n); if (Math.abs(x) > 1) throw ERR.NUM; return Math.acos(x); },
  ATAN: ([n]) => Math.atan(toNum(n)),
  ATAN2: ([x, y]) => { const a = toNum(x); const b = toNum(y); if (a === 0 && b === 0) throw ERR.DIV0; return Math.atan2(b, a); },
  SINH: ([n]) => Math.sinh(toNum(n)),
  COSH: ([n]) => Math.cosh(toNum(n)),
  TANH: ([n]) => Math.tanh(toNum(n)),
  ASINH: ([n]) => Math.asinh(toNum(n)),
  ACOSH: ([n]) => { const x = toNum(n); if (x < 1) throw ERR.NUM; return Math.acosh(x); },
  ATANH: ([n]) => { const x = toNum(n); if (Math.abs(x) >= 1) throw ERR.NUM; return Math.atanh(x); },
  COT: ([n]) => { const x = toNum(n); if (x === 0) throw ERR.DIV0; return 1 / Math.tan(x); },
  COTH: ([n]) => { const x = toNum(n); if (x === 0) throw ERR.DIV0; return 1 / Math.tanh(x); },
  ACOT: ([n]) => Math.PI / 2 - Math.atan(toNum(n)),
  ACOTH: ([n]) => { const x = toNum(n); if (Math.abs(x) <= 1) throw ERR.NUM; return 0.5 * Math.log((x + 1) / (x - 1)); },
  SEC: ([n]) => 1 / Math.cos(toNum(n)),
  SECH: ([n]) => 1 / Math.cosh(toNum(n)),
  CSC: ([n]) => { const x = toNum(n); if (x === 0) throw ERR.DIV0; return 1 / Math.sin(x); },
  CSCH: ([n]) => { const x = toNum(n); if (x === 0) throw ERR.DIV0; return 1 / Math.sinh(x); },
  ROMAN: ([n, form]) => {
    let x = Math.trunc(toNum(n));
    optNum(form, 0);
    if (x < 0 || x > 3999) throw ERR.VALUE;
    let s = '';
    for (const [v, r] of ROMAN) while (x >= v) { s += r; x -= v; }
    return s;
  },
  ARABIC: ([t]) => {
    const s = String(scalar(t) ?? '').trim().toUpperCase();
    if (!s) return 0;
    const neg = s.startsWith('-');
    const body = neg ? s.slice(1) : s;
    const val = { I: 1, V: 5, X: 10, L: 50, C: 100, D: 500, M: 1000 };
    let total = 0;
    for (let i = 0; i < body.length; i++) {
      const v = val[body[i]];
      if (!v) throw ERR.VALUE;
      const next = val[body[i + 1]] ?? 0;
      total += v < next ? -v : v;
    }
    return neg ? -total : total;
  },
  BASE: ([n, radix, len]) => {
    const x = Math.trunc(toNum(n));
    const r = toInt(radix);
    if (x < 0 || r < 2 || r > 36) throw ERR.NUM;
    return x.toString(r).toUpperCase().padStart(optInt(len, 0), '0');
  },
  DECIMAL: ([t, radix]) => {
    const r = toInt(radix);
    if (r < 2 || r > 36) throw ERR.NUM;
    const s = String(scalar(t) ?? '').trim();
    if (!s) return 0;
    const v = parseInt(s, r);
    if (Number.isNaN(v) || !new RegExp(`^[0-9A-Z]+$`, 'i').test(s) || [...s.toUpperCase()].some((ch) => parseInt(ch, 36) >= r)) throw ERR.NUM;
    return v;
  },
};

export const MATH = {
  SUM: (a) => collectNums(a).reduce((s, x) => s + x, 0),
  PRODUCT: (a) => { const n = collectNums(a); return n.length ? n.reduce((s, x) => s * x, 1) : 0; },
  SUMSQ: (a) => collectNums(a).reduce((s, x) => s + x * x, 0),
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
  SUMX2MY2: ([x, y]) => pairSum(x, y, (a, b) => a * a - b * b),
  SUMX2PY2: ([x, y]) => pairSum(x, y, (a, b) => a * a + b * b),
  SUMXMY2: ([x, y]) => pairSum(x, y, (a, b) => (a - b) ** 2),
  SERIESSUM: ([x, n, m, coef]) => {
    const xv = toNum(x);
    const nv = toNum(n);
    const mv = toNum(m);
    return collectNums([coef]).reduce((s, c, i) => s + c * xv ** (nv + i * mv), 0);
  },
  GCD: (a) => {
    const n = collectNums(a).map(Math.trunc);
    if (n.some((x) => x < 0)) throw ERR.NUM;
    return n.reduce((g, x) => gcd2(g, x), 0);
  },
  LCM: (a) => {
    const n = collectNums(a).map(Math.trunc);
    if (n.some((x) => x < 0)) throw ERR.NUM;
    if (n.some((x) => x === 0)) return 0;
    return n.reduce((l, x) => (l * x) / gcd2(l, x), 1);
  },
  MULTINOMIAL: (a) => {
    const n = collectNums(a).map(Math.trunc);
    if (n.some((x) => x < 0)) throw ERR.NUM;
    return Math.round(fact(n.reduce((s, x) => s + x, 0)) / n.reduce((p, x) => p * fact(x), 1));
  },
  RAND: () => Math.random(),
  RANDBETWEEN: ([lo, hi]) => {
    const a = Math.ceil(toNum(lo));
    const b = Math.floor(toNum(hi));
    if (a > b) throw ERR.NUM;
    return a + Math.floor(Math.random() * (b - a + 1));
  },
  RANDARRAY: ([rows, cols, min, max, whole]) => {
    const h = optInt(rows, 1);
    const w = optInt(cols, 1);
    const lo = optNum(min, 0);
    const hi = optNum(max, 1);
    const int = whole === undefined || whole === null ? false : toBool(whole);
    if (h < 1 || w < 1 || lo > hi) throw ERR.VALUE;
    const f = () => (int ? Math.ceil(lo) + Math.floor(Math.random() * (Math.floor(hi) - Math.ceil(lo) + 1)) : lo + Math.random() * (hi - lo));
    return new Range(Array.from({ length: h }, () => Array.from({ length: w }, f)));
  },
  SEQUENCE: ([rows, cols, start, step]) => {
    const h = optInt(rows, 1);
    const w = optInt(cols, 1);
    const s = optNum(start, 1);
    const d = optNum(step, 1);
    if (h < 1 || w < 1) throw ERR.CALC;
    if (h * w > 5_000_000) throw ERR.NUM;
    return new Range(Array.from({ length: h }, (_, r) => Array.from({ length: w }, (__, c) => s + (r * w + c) * d)));
  },
  MMULT: ([a, b]) => new Range(mmult(numMatrix(a), numMatrix(b))),
  MDETERM: ([a]) => {
    const m = numMatrix(a);
    if (m.length !== m[0].length) throw ERR.VALUE;
    return det(m);
  },
  MINVERSE: ([a]) => {
    const m = numMatrix(a);
    if (m.length !== m[0].length) throw ERR.VALUE;
    return new Range(inverse(m));
  },
  MUNIT: ([n]) => {
    const k = toInt(n);
    if (k < 1) throw ERR.VALUE;
    return new Range(Array.from({ length: k }, (_, i) => Array.from({ length: k }, (__, j) => (i === j ? 1 : 0))));
  },

  // 조건부 집계
  SUMIF: ([range, crit, sumRange]) => {
    { const re = rangeError([range, sumRange]); if (re) return re; }
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
  SUMIFS: ([sumRange, ...rest]) => {
    const re = rangeError([sumRange, ...rest.filter((_, i) => i % 2 === 0)]);
    if (re) return re; const f = fastIfs(sumRange, rest); return f ? f.err ?? f.sum : sumIfs(sumRange, ifsPairs(rest)); },
  COUNTIF: ([range, crit]) => {
    { const re = rangeError([range]); if (re) return re; }
    const fast = fastCount(asRange(range), crit);
    if (fast !== null) return fast;
    const test = makeCriteria(crit);
    let k = 0;
    for (const v of asRange(range).values()) if (test(v)) k++;
    return k;
  },
  COUNTIFS: (args) => {
    const re = rangeError(args.filter((_, i) => i % 2 === 0));
    if (re) return re;
    const f = fastIfs(null, args);
    if (f) return f.rows;
    const pairs = ifsPairs(args);
    let k = 0;
    pairs[0].range.rows.forEach((row, i) => row.forEach((_, j) => { if (matchesAll(pairs, i, j)) k++; }));
    return k;
  },
  AVERAGEIF: ([range, crit, avgRange]) => {
    { const re = rangeError([range, avgRange]); if (re) return re; }
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
  AVERAGEIFS: ([target, ...rest]) => {
    const re = rangeError([target, ...rest.filter((_, i) => i % 2 === 0)]);
    if (re) return re;
    const f = fastIfs(target, rest);
    if (f) { if (!f.n) throw ERR.DIV0; return f.sum / f.n; }
    const v = ifsValues(target, ifsPairs(rest));
    if (!v.length) throw ERR.DIV0;
    return v.reduce((s, x) => s + x, 0) / v.length;
  },
  MAXIFS: ([target, ...rest]) => { const re = rangeError([target, ...rest.filter((_, i) => i % 2 === 0)]); if (re) return re; const f = fastIfs(target, rest); if (f) return f.n ? f.max : 0; const v = ifsValues(target, ifsPairs(rest)); return v.length ? maxOf(v) : 0; },
  MINIFS: ([target, ...rest]) => { const re = rangeError([target, ...rest.filter((_, i) => i % 2 === 0)]); if (re) return re; const f = fastIfs(target, rest); if (f) return f.n ? f.min : 0; const v = ifsValues(target, ifsPairs(rest)); return v.length ? minOf(v) : 0; },
};

function pairSum(x, y, f) {
  const a = flat([x]);
  const b = flat([y]);
  if (a.length !== b.length) throw ERR.NA;
  let s = 0;
  for (let i = 0; i < a.length; i++) {
    if (isError(a[i])) throw a[i];
    if (isError(b[i])) throw b[i];
    if (typeof a[i] === 'number' && typeof b[i] === 'number') s += f(a[i], b[i]);
  }
  return s;
}

for (const [k, fn] of Object.entries(SCALAR)) MATH[k] = lift(fn);
// 조건 자리에 범위를 주면 조건마다 결과를 돌려줌 (엑셀: =SUMPRODUCT(1/COUNTIF(A1:A9,A1:A9)) 고유 개수)
const odd = Array.from({ length: 127 }, (_, i) => i * 2 + 1);
const evenFrom2 = Array.from({ length: 126 }, (_, i) => i * 2 + 2);
for (const k of ['SUMIF', 'COUNTIF', 'AVERAGEIF']) MATH[k] = lift(MATH[k], [1]);
MATH.COUNTIFS = lift(MATH.COUNTIFS, odd);
for (const k of ['SUMIFS', 'AVERAGEIFS', 'MAXIFS', 'MINIFS']) MATH[k] = lift(MATH[k], evenFrom2);
