// 통계 · 분포 함수 (DOM 없음)
import {
  ERR, Range, isError, toNum, toInt, toBool, optNum, optInt, checkNum, collectNums, collectNumsA, flat, asRange, lift,
  parseNumberText,
} from './fxcore.js';
import { mmult, inverse } from './fx-math.js';
import { fitEts, etsForecast, etsConfint } from './ets.js';

// ───────────── 특수 함수 ─────────────
const LANCZOS = [0.99999999999980993, 676.5203681218851, -1259.1392167224028, 771.32342877765313, -176.61502916214059,
  12.507343278686905, -0.13857109526572012, 9.9843695780195716e-6, 1.5056327351493116e-7];
export function lnGamma(x) {
  if (x < 0.5) return Math.log(Math.PI / Math.abs(Math.sin(Math.PI * x))) - lnGamma(1 - x);
  x -= 1;
  let a = LANCZOS[0];
  const t = x + 7.5;
  for (let i = 1; i < 9; i++) a += LANCZOS[i] / (x + i);
  return 0.5 * Math.log(2 * Math.PI) + (x + 0.5) * Math.log(t) - t + Math.log(a);
}
export function gammaFn(x) {
  if (x === Math.floor(x) && x <= 0) throw ERR.NUM;
  if (x < 0.5) return Math.PI / (Math.sin(Math.PI * x) * gammaFn(1 - x));
  return Math.exp(lnGamma(x));
}

/** 정규화된 하부 불완전 감마 P(a, x) */
export function gammaP(a, x) {
  if (x <= 0) return 0;
  if (x < a + 1) {
    let sum = 1 / a;
    let del = sum;
    let ap = a;
    for (let n = 0; n < 1000; n++) {
      ap += 1;
      del *= x / ap;
      sum += del;
      if (Math.abs(del) < Math.abs(sum) * 1e-16) break;
    }
    return sum * Math.exp(-x + a * Math.log(x) - lnGamma(a));
  }
  return 1 - gammaQcf(a, x);
}
function gammaQcf(a, x) {
  let b = x + 1 - a;
  let c = 1 / 1e-300;
  let d = 1 / b;
  let h = d;
  for (let i = 1; i < 1000; i++) {
    const an = -i * (i - a);
    b += 2;
    d = an * d + b;
    if (Math.abs(d) < 1e-300) d = 1e-300;
    c = b + an / c;
    if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return Math.exp(-x + a * Math.log(x) - lnGamma(a)) * h;
}

/** 정규화된 불완전 베타 I_x(a, b) */
export function betaI(x, a, b) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  const bt = Math.exp(lnGamma(a + b) - lnGamma(a) - lnGamma(b) + a * Math.log(x) + b * Math.log(1 - x));
  if (x < (a + 1) / (a + b + 2)) return (bt * betacf(x, a, b)) / a;
  return 1 - (bt * betacf(1 - x, b, a)) / b;
}
function betacf(x, a, b) {
  const qab = a + b;
  const qap = a + 1;
  const qam = a - 1;
  let c = 1;
  let d = 1 - (qab * x) / qap;
  if (Math.abs(d) < 1e-300) d = 1e-300;
  d = 1 / d;
  let h = d;
  for (let m = 1; m <= 1000; m++) {
    const m2 = 2 * m;
    let aa = (m * (b - m) * x) / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < 1e-300) d = 1e-300;
    c = 1 + aa / c; if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    h *= d * c;
    aa = (-(a + m) * (qab + m) * x) / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < 1e-300) d = 1e-300;
    c = 1 + aa / c; if (Math.abs(c) < 1e-300) c = 1e-300;
    d = 1 / d;
    const del = d * c;
    h *= del;
    if (Math.abs(del - 1) < 1e-16) break;
  }
  return h;
}

export const erf = (x) => (x >= 0 ? gammaP(0.5, x * x) : -gammaP(0.5, x * x));
export const erfc = (x) => (x >= 0 ? (x > 26 ? 0 : x < 0.5 ? 1 - erf(x) : gammaQcf(0.5, x * x)) : 2 - erfc(-x));
export const normCdf = (z) => 0.5 * erfc(-z / Math.SQRT2);
export const normPdf = (z) => Math.exp(-0.5 * z * z) / Math.sqrt(2 * Math.PI);

/** 표준 정규 역함수 (Acklam + 뉴턴 보정) */
export function normInv(p) {
  if (p <= 0 || p >= 1) throw ERR.NUM;
  const a = [-3.969683028665376e1, 2.209460984245205e2, -2.759285104469687e2, 1.383577518672690e2, -3.066479806614716e1, 2.506628277459239];
  const b = [-5.447609879822406e1, 1.615858368580409e2, -1.556989798598866e2, 6.680131188771972e1, -1.328068155288572e1];
  const c = [-7.784894002430293e-3, -3.223964580411365e-1, -2.400758277161838, -2.549732539343734, 4.374664141464968, 2.938163982698783];
  const d = [7.784695709041462e-3, 3.224671290700398e-1, 2.445134137142996, 3.754408661907416];
  const pl = 0.02425;
  let x;
  if (p < pl) {
    const q = Math.sqrt(-2 * Math.log(p));
    x = (((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  } else if (p <= 1 - pl) {
    const q = p - 0.5;
    const r = q * q;
    x = ((((((a[0] * r + a[1]) * r + a[2]) * r + a[3]) * r + a[4]) * r + a[5]) * q) / (((((b[0] * r + b[1]) * r + b[2]) * r + b[3]) * r + b[4]) * r + 1);
  } else {
    const q = Math.sqrt(-2 * Math.log(1 - p));
    x = -(((((c[0] * q + c[1]) * q + c[2]) * q + c[3]) * q + c[4]) * q + c[5]) / ((((d[0] * q + d[1]) * q + d[2]) * q + d[3]) * q + 1);
  }
  for (let i = 0; i < 2; i++) {
    const e = normCdf(x) - p;
    const u = e * Math.sqrt(2 * Math.PI) * Math.exp((x * x) / 2);
    x -= u / (1 + (x * u) / 2);
  }
  return x;
}

/** 누적 분포 함수의 역함수 (단조 증가 cdf, 범위 [lo, hi]) */
export function invert(cdf, p, lo, hi) {
  if (!(p > 0 && p < 1)) throw ERR.NUM;
  let a = lo;
  let b = hi;
  while (cdf(b) < p && b < 1e300) { a = b; b *= 2; }
  for (let i = 0; i < 300; i++) {
    const m = (a + b) / 2;
    if (cdf(m) < p) a = m; else b = m;
    if (b - a < 1e-15 * Math.max(1, Math.abs(m))) break;
  }
  return (a + b) / 2;
}

export const tCdf = (t, df) => {
  const x = df / (df + t * t);
  const tail = 0.5 * betaI(x, df / 2, 0.5);
  return t >= 0 ? 1 - tail : tail;
};
const chiCdf = (x, k) => gammaP(k / 2, x / 2);
export const fCdf = (x, d1, d2) => (x <= 0 ? 0 : betaI((d1 * x) / (d1 * x + d2), d1 / 2, d2 / 2));

// ───────────── 지수 평활 예측 (FORECAST.ETS) ─────────────
function etsModel(values, timeline, seas, comp, agg) {
  const cells = (v) => asRange(v).rows.flat().map((x) => (isError(x) ? (() => { throw x; })() : x === '' ? null : x));
  const vals = cells(values);
  const tl = cells(timeline);
  if (vals.length !== tl.length) throw ERR.NA;
  const seasonality = optInt(seas, 1);
  const completion = optInt(comp, 1);
  const aggregation = optInt(agg, 1);
  if (seasonality < 0 || ![0, 1].includes(completion) || aggregation < 1 || aggregation > 7) throw ERR.NUM;
  const m = fitEts(vals.map((x) => (typeof x === 'number' ? x : null)), tl.map((x) => (typeof x === 'number' ? x : NaN)), { seasonality, completion, aggregation });
  if (m.error) throw ERR[m.error];
  return m;
}
const etsTarget = (m, x) => {
  const t = toNum(x);
  if (m.axis.pos(t) < 0) throw ERR.NUM;
  return t;
};

// ───────────── 통계 도우미 ─────────────
const mean = (n) => n.reduce((s, x) => s + x, 0) / n.length;
function variance(n, sample) {
  if (n.length < (sample ? 2 : 1)) throw ERR.DIV0;
  const m = mean(n);
  return n.reduce((s, x) => s + (x - m) ** 2, 0) / (n.length - (sample ? 1 : 0));
}
function sortedNums(v) {
  const n = collectNums([v]).sort((a, b) => a - b);
  if (!n.length) throw ERR.NUM;
  return n;
}
function percentileInc(n, p) {
  if (p < 0 || p > 1) throw ERR.NUM;
  const h = (n.length - 1) * p;
  const lo = Math.floor(h);
  return n[lo] + (h - lo) * ((n[lo + 1] ?? n[lo]) - n[lo]);
}
function percentileExc(n, p) {
  const h = (n.length + 1) * p - 1;
  if (h < 0 || h > n.length - 1) throw ERR.NUM;
  const lo = Math.floor(h);
  return n[lo] + (h - lo) * ((n[lo + 1] ?? n[lo]) - n[lo]);
}
/** 두 배열에서 둘 다 숫자인 쌍만 */
function pairs(x, y) {
  const a = flat([x]);
  const b = flat([y]);
  if (a.length !== b.length) throw ERR.NA;
  const xs = [];
  const ys = [];
  for (let i = 0; i < a.length; i++) {
    if (isError(a[i])) throw a[i];
    if (isError(b[i])) throw b[i];
    if (typeof a[i] === 'number' && typeof b[i] === 'number') { xs.push(a[i]); ys.push(b[i]); }
  }
  return [xs, ys];
}
function linreg(ys, xs) {
  if (xs.length < 2) throw ERR.DIV0;
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  for (let i = 0; i < xs.length; i++) { sxy += (xs[i] - mx) * (ys[i] - my); sxx += (xs[i] - mx) ** 2; }
  if (sxx === 0) throw ERR.DIV0;
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, mx, my };
}
function correl(xs, ys) {
  const mx = mean(xs);
  const my = mean(ys);
  let sxy = 0;
  let sxx = 0;
  let syy = 0;
  for (let i = 0; i < xs.length; i++) { sxy += (xs[i] - mx) * (ys[i] - my); sxx += (xs[i] - mx) ** 2; syy += (ys[i] - my) ** 2; }
  if (!sxx || !syy) throw ERR.DIV0;
  return sxy / Math.sqrt(sxx * syy);
}
function rank(v, ref, order, avg) {
  const x = toNum(v);
  const n = collectNums([ref]);
  const asc = optNum(order, 0) !== 0;
  if (!n.includes(x)) throw ERR.NA;
  const better = n.filter((y) => (asc ? y < x : y > x)).length;
  if (!avg) return better + 1;
  const same = n.filter((y) => y === x).length;
  return better + (same + 1) / 2;
}
function modes(v) {
  const n = collectNums(Array.isArray(v) ? v : [v]);
  const counts = new Map();
  for (const x of n) counts.set(x, (counts.get(x) ?? 0) + 1);
  const max = Math.max(0, ...counts.values());
  if (max < 2) throw ERR.NA;
  return [...counts.entries()].filter(([, k]) => k === max).map(([x]) => x);
}
const lookupPercentRank = (n, x, sig, exc) => {
  if (x < n[0] || x > n[n.length - 1]) throw ERR.NA;
  let i = 0;
  while (i < n.length - 1 && n[i + 1] <= x) i++;
  let pos = i;
  if (n[i] !== x) pos = i + (x - n[i]) / (n[i + 1] - n[i]);
  const r = exc ? (pos + 1) / (n.length + 1) : pos / (n.length - 1);
  const f = 10 ** sig;
  return Math.floor(r * f + 1e-9) / f;
};

const DIST = {
  'NORM.DIST': ([x, m, s, cum]) => {
    const sd = toNum(s);
    if (sd <= 0) throw ERR.NUM;
    const z = (toNum(x) - toNum(m)) / sd;
    return toBool(cum) ? normCdf(z) : normPdf(z) / sd;
  },
  'NORM.INV': ([p, m, s]) => { const sd = toNum(s); if (sd <= 0) throw ERR.NUM; return toNum(m) + sd * normInv(toNum(p)); },
  'NORM.S.DIST': ([z, cum]) => (toBool(cum) ? normCdf(toNum(z)) : normPdf(toNum(z))),
  NORMSDIST: ([z]) => normCdf(toNum(z)),
  'NORM.S.INV': ([p]) => normInv(toNum(p)),
  NORMSINV: ([p]) => normInv(toNum(p)),
  PHI: ([z]) => normPdf(toNum(z)),
  GAUSS: ([z]) => normCdf(toNum(z)) - 0.5,
  STANDARDIZE: ([x, m, s]) => { const sd = toNum(s); if (sd <= 0) throw ERR.NUM; return (toNum(x) - toNum(m)) / sd; },
  'LOGNORM.DIST': ([x, m, s, cum]) => {
    const v = toNum(x);
    const sd = toNum(s);
    if (v <= 0 || sd <= 0) throw ERR.NUM;
    const z = (Math.log(v) - toNum(m)) / sd;
    return toBool(cum) ? normCdf(z) : normPdf(z) / (v * sd);
  },
  LOGNORMDIST: ([x, m, s]) => { const v = toNum(x); const sd = toNum(s); if (v <= 0 || sd <= 0) throw ERR.NUM; return normCdf((Math.log(v) - toNum(m)) / sd); },
  'LOGNORM.INV': ([p, m, s]) => { const sd = toNum(s); if (sd <= 0) throw ERR.NUM; return Math.exp(toNum(m) + sd * normInv(toNum(p))); },
  'EXPON.DIST': ([x, l, cum]) => {
    const v = toNum(x);
    const lam = toNum(l);
    if (v < 0 || lam <= 0) throw ERR.NUM;
    return toBool(cum) ? 1 - Math.exp(-lam * v) : lam * Math.exp(-lam * v);
  },
  'POISSON.DIST': ([x, m, cum]) => {
    const k = Math.floor(toNum(x));
    const mu = toNum(m);
    if (k < 0 || mu < 0) throw ERR.NUM;
    const pmf = (i) => Math.exp(i * Math.log(mu) - mu - lnGamma(i + 1));
    if (!toBool(cum)) return mu === 0 ? (k === 0 ? 1 : 0) : pmf(k);
    return mu === 0 ? 1 : 1 - gammaP(k + 1, mu);
  },
  'BINOM.DIST': ([x, n, p, cum]) => {
    const k = Math.floor(toNum(x));
    const t = Math.floor(toNum(n));
    const pr = toNum(p);
    if (k < 0 || k > t || pr < 0 || pr > 1) throw ERR.NUM;
    const pmf = (i) => Math.exp(lnGamma(t + 1) - lnGamma(i + 1) - lnGamma(t - i + 1) + (i ? i * Math.log(pr) : 0) + (t - i ? (t - i) * Math.log(1 - pr) : 0));
    if (!toBool(cum)) return pmf(k);
    let s = 0;
    for (let i = 0; i <= k; i++) s += pmf(i);
    return Math.min(1, s);
  },
  'BINOM.INV': ([n, p, a]) => {
    const t = Math.floor(toNum(n));
    const pr = toNum(p);
    const alpha = toNum(a);
    if (t < 0 || pr < 0 || pr > 1 || alpha <= 0 || alpha >= 1) throw ERR.NUM;
    let s = 0;
    for (let i = 0; i <= t; i++) {
      s += Math.exp(lnGamma(t + 1) - lnGamma(i + 1) - lnGamma(t - i + 1) + (i ? i * Math.log(pr) : 0) + (t - i ? (t - i) * Math.log(1 - pr) : 0));
      if (s >= alpha - 1e-12) return i;
    }
    return t;
  },
  'NEGBINOM.DIST': ([f, s, p, cum]) => {
    const k = Math.floor(toNum(f));
    const r = Math.floor(toNum(s));
    const pr = toNum(p);
    if (k < 0 || r < 1 || pr < 0 || pr > 1) throw ERR.NUM;
    const pmf = (i) => Math.exp(lnGamma(i + r) - lnGamma(i + 1) - lnGamma(r) + r * Math.log(pr) + i * Math.log(1 - pr));
    if (!toBool(cum)) return pmf(k);
    return betaI(pr, r, k + 1);
  },
  'HYPGEOM.DIST': ([x, n, K, N, cum]) => {
    const k = Math.floor(toNum(x));
    const ns = Math.floor(toNum(n));
    const ks = Math.floor(toNum(K));
    const np = Math.floor(toNum(N));
    if (k < 0 || ns <= 0 || ks <= 0 || np <= 0 || ns > np || ks > np || k > ns || k > ks) throw ERR.NUM;
    const lc = (a, b) => lnGamma(a + 1) - lnGamma(b + 1) - lnGamma(a - b + 1);
    const pmf = (i) => (i < Math.max(0, ns - (np - ks)) ? 0 : Math.exp(lc(ks, i) + lc(np - ks, ns - i) - lc(np, ns)));
    if (!toBool(cum)) return pmf(k);
    let s = 0;
    for (let i = 0; i <= k; i++) s += pmf(i);
    return s;
  },
  GAMMA: ([x]) => checkNum(gammaFn(toNum(x))),
  GAMMALN: ([x]) => { const v = toNum(x); if (v <= 0) throw ERR.NUM; return lnGamma(v); },
  'GAMMALN.PRECISE': ([x]) => { const v = toNum(x); if (v <= 0) throw ERR.NUM; return lnGamma(v); },
  'GAMMA.DIST': ([x, a, b, cum]) => {
    const v = toNum(x);
    const al = toNum(a);
    const be = toNum(b);
    if (v < 0 || al <= 0 || be <= 0) throw ERR.NUM;
    if (toBool(cum)) return gammaP(al, v / be);
    return Math.exp((al - 1) * Math.log(v) - v / be - lnGamma(al) - al * Math.log(be));
  },
  'GAMMA.INV': ([p, a, b]) => {
    const al = toNum(a);
    const be = toNum(b);
    if (al <= 0 || be <= 0) throw ERR.NUM;
    return invert((x) => gammaP(al, x / be), toNum(p), 0, al * be * 2 + 1);
  },
  'BETA.DIST': ([x, a, b, cum, lo, hi]) => {
    const A = optNum(lo, 0);
    const B = optNum(hi, 1);
    const v = (toNum(x) - A) / (B - A);
    const al = toNum(a);
    const be = toNum(b);
    if (al <= 0 || be <= 0 || v < 0 || v > 1) throw ERR.NUM;
    if (toBool(cum)) return betaI(v, al, be);
    return Math.exp((al - 1) * Math.log(v) + (be - 1) * Math.log(1 - v) - (lnGamma(al) + lnGamma(be) - lnGamma(al + be))) / (B - A);
  },
  'BETA.INV': ([p, a, b, lo, hi]) => {
    const A = optNum(lo, 0);
    const B = optNum(hi, 1);
    const al = toNum(a);
    const be = toNum(b);
    if (al <= 0 || be <= 0) throw ERR.NUM;
    return A + (B - A) * invert((x) => betaI(Math.min(1, x), al, be), toNum(p), 0, 1);
  },
  'T.DIST': ([x, df, cum]) => {
    const t = toNum(x);
    const n = toNum(df);
    if (n < 1) throw ERR.NUM;
    if (toBool(cum)) return tCdf(t, n);
    return Math.exp(lnGamma((n + 1) / 2) - lnGamma(n / 2)) / Math.sqrt(n * Math.PI) * (1 + (t * t) / n) ** (-(n + 1) / 2);
  },
  'T.DIST.2T': ([x, df]) => { const t = toNum(x); const n = toNum(df); if (t < 0 || n < 1) throw ERR.NUM; return 2 * (1 - tCdf(t, n)); },
  'T.DIST.RT': ([x, df]) => { const n = toNum(df); if (n < 1) throw ERR.NUM; return 1 - tCdf(toNum(x), n); },
  TDIST: ([x, df, tails]) => {
    const t = toNum(x);
    const n = toNum(df);
    const k = toInt(tails);
    if (t < 0 || n < 1 || (k !== 1 && k !== 2)) throw ERR.NUM;
    return k * (1 - tCdf(t, n));
  },
  'T.INV': ([p, df]) => {
    const pr = toNum(p);
    const n = toNum(df);
    if (n < 1) throw ERR.NUM;
    if (pr === 0.5) return 0;
    const q = pr > 0.5 ? pr : 1 - pr;
    const t = invert((x) => tCdf(x, n), q, 0, 10);
    return pr > 0.5 ? t : -t;
  },
  'T.INV.2T': ([p, df]) => { const pr = toNum(p); const n = toNum(df); if (pr <= 0 || pr > 1 || n < 1) throw ERR.NUM; return invert((x) => tCdf(x, n), 1 - pr / 2, 0, 10); },
  TINV: ([p, df]) => { const pr = toNum(p); const n = toNum(df); if (pr <= 0 || pr > 1 || n < 1) throw ERR.NUM; return invert((x) => tCdf(x, n), 1 - pr / 2, 0, 10); },
  'CHISQ.DIST': ([x, df, cum]) => {
    const v = toNum(x);
    const k = toNum(df);
    if (v < 0 || k < 1) throw ERR.NUM;
    if (toBool(cum)) return chiCdf(v, k);
    return Math.exp((k / 2 - 1) * Math.log(v) - v / 2 - (k / 2) * Math.log(2) - lnGamma(k / 2));
  },
  'CHISQ.DIST.RT': ([x, df]) => { const v = toNum(x); const k = toNum(df); if (v < 0 || k < 1) throw ERR.NUM; return 1 - chiCdf(v, k); },
  CHIDIST: ([x, df]) => { const v = toNum(x); const k = toNum(df); if (v < 0 || k < 1) throw ERR.NUM; return 1 - chiCdf(v, k); },
  'CHISQ.INV': ([p, df]) => { const k = toNum(df); if (k < 1) throw ERR.NUM; return invert((x) => chiCdf(x, k), toNum(p), 0, k + 10); },
  'CHISQ.INV.RT': ([p, df]) => { const k = toNum(df); if (k < 1) throw ERR.NUM; return invert((x) => chiCdf(x, k), 1 - toNum(p), 0, k + 10); },
  CHIINV: ([p, df]) => { const k = toNum(df); if (k < 1) throw ERR.NUM; return invert((x) => chiCdf(x, k), 1 - toNum(p), 0, k + 10); },
  'F.DIST': ([x, a, b, cum]) => {
    const v = toNum(x);
    const d1 = toNum(a);
    const d2 = toNum(b);
    if (v < 0 || d1 < 1 || d2 < 1) throw ERR.NUM;
    if (toBool(cum)) return fCdf(v, d1, d2);
    return Math.exp(0.5 * (d1 * Math.log(d1 * v) + d2 * Math.log(d2) - (d1 + d2) * Math.log(d1 * v + d2)) - Math.log(v) - (lnGamma(d1 / 2) + lnGamma(d2 / 2) - lnGamma((d1 + d2) / 2)));
  },
  'F.DIST.RT': ([x, a, b]) => { const d1 = toNum(a); const d2 = toNum(b); if (toNum(x) < 0 || d1 < 1 || d2 < 1) throw ERR.NUM; return 1 - fCdf(toNum(x), d1, d2); },
  FDIST: ([x, a, b]) => { const d1 = toNum(a); const d2 = toNum(b); if (toNum(x) < 0 || d1 < 1 || d2 < 1) throw ERR.NUM; return 1 - fCdf(toNum(x), d1, d2); },
  'F.INV': ([p, a, b]) => { const d1 = toNum(a); const d2 = toNum(b); if (d1 < 1 || d2 < 1) throw ERR.NUM; return invert((x) => fCdf(x, d1, d2), toNum(p), 0, 10); },
  'F.INV.RT': ([p, a, b]) => { const d1 = toNum(a); const d2 = toNum(b); if (d1 < 1 || d2 < 1) throw ERR.NUM; return invert((x) => fCdf(x, d1, d2), 1 - toNum(p), 0, 10); },
  FINV: ([p, a, b]) => { const d1 = toNum(a); const d2 = toNum(b); if (d1 < 1 || d2 < 1) throw ERR.NUM; return invert((x) => fCdf(x, d1, d2), 1 - toNum(p), 0, 10); },
  'WEIBULL.DIST': ([x, a, b, cum]) => {
    const v = toNum(x);
    const al = toNum(a);
    const be = toNum(b);
    if (v < 0 || al <= 0 || be <= 0) throw ERR.NUM;
    if (toBool(cum)) return 1 - Math.exp(-((v / be) ** al));
    return (al / be ** al) * v ** (al - 1) * Math.exp(-((v / be) ** al));
  },
  'CONFIDENCE.NORM': ([a, s, n]) => {
    const alpha = toNum(a);
    const sd = toNum(s);
    const size = Math.floor(toNum(n));
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 1) throw ERR.NUM;
    return normInv(1 - alpha / 2) * sd / Math.sqrt(size);
  },
  'CONFIDENCE.T': ([a, s, n]) => {
    const alpha = toNum(a);
    const sd = toNum(s);
    const size = Math.floor(toNum(n));
    if (alpha <= 0 || alpha >= 1 || sd <= 0 || size < 2) throw ERR.NUM;
    return invert((x) => tCdf(x, size - 1), 1 - alpha / 2, 0, 10) * sd / Math.sqrt(size);
  },
  FISHER: ([x]) => { const v = toNum(x); if (Math.abs(v) >= 1) throw ERR.NUM; return 0.5 * Math.log((1 + v) / (1 - v)); },
  FISHERINV: ([y]) => { const v = toNum(y); return (Math.exp(2 * v) - 1) / (Math.exp(2 * v) + 1); },
  ERF: ([lo, hi]) => (hi === undefined || hi === null ? erf(toNum(lo)) : erf(toNum(hi)) - erf(toNum(lo))),
  'ERF.PRECISE': ([x]) => erf(toNum(x)),
  ERFC: ([x]) => erfc(toNum(x)),
  'ERFC.PRECISE': ([x]) => erfc(toNum(x)),
};
// 옛 이름
Object.assign(DIST, {
  NORMDIST: DIST['NORM.DIST'], NORMINV: DIST['NORM.INV'], EXPONDIST: DIST['EXPON.DIST'], POISSON: DIST['POISSON.DIST'],
  BINOMDIST: DIST['BINOM.DIST'], CRITBINOM: DIST['BINOM.INV'], NEGBINOMDIST: (a) => DIST['NEGBINOM.DIST']([...a.slice(0, 3), false]),
  HYPGEOMDIST: (a) => DIST['HYPGEOM.DIST']([...a.slice(0, 4), false]), GAMMADIST: DIST['GAMMA.DIST'], GAMMAINV: DIST['GAMMA.INV'],
  BETADIST: (a) => DIST['BETA.DIST']([a[0], a[1], a[2], true, a[3], a[4]]), BETAINV: DIST['BETA.INV'], WEIBULL: DIST['WEIBULL.DIST'],
  CONFIDENCE: DIST['CONFIDENCE.NORM'], LOGINV: DIST['LOGNORM.INV'],
});

export const STAT = {
  AVERAGE: (a) => { const n = collectNums(a); if (!n.length) throw ERR.DIV0; return mean(n); },
  AVERAGEA: (a) => { const n = collectNumsA(a); if (!n.length) throw ERR.DIV0; return mean(n); },
  COUNT: (a) => {
    let k = 0;
    for (const x of a) {
      if (x instanceof Range) { for (const v of x.values()) if (typeof v === 'number') k++; }
      else if (typeof x === 'number' || typeof x === 'boolean' || (typeof x === 'string' && parseNumberText(x) !== null)) k++;
    }
    return k;
  },
  COUNTA: (a) => flat(a).filter((v) => v !== null && v !== undefined).length,
  COUNTBLANK: (a) => flat(a).filter((v) => v === null || v === '').length,
  MAX: (a) => { const n = collectNums(a); return n.length ? Math.max(...n) : 0; },
  MIN: (a) => { const n = collectNums(a); return n.length ? Math.min(...n) : 0; },
  MAXA: (a) => { const n = collectNumsA(a); return n.length ? Math.max(...n) : 0; },
  MINA: (a) => { const n = collectNumsA(a); return n.length ? Math.min(...n) : 0; },
  MEDIAN: (a) => {
    const n = collectNums(a).sort((x, y) => x - y);
    if (!n.length) throw ERR.NUM;
    const m = n.length >> 1;
    return n.length % 2 ? n[m] : (n[m - 1] + n[m]) / 2;
  },
  MODE: (a) => modes(a)[0],
  'MODE.SNGL': (a) => modes(a)[0],
  'MODE.MULT': (a) => new Range(modes(a).map((x) => [x])),
  STDEV: (a) => Math.sqrt(variance(collectNums(a), true)),
  'STDEV.S': (a) => Math.sqrt(variance(collectNums(a), true)),
  STDEVP: (a) => Math.sqrt(variance(collectNums(a), false)),
  'STDEV.P': (a) => Math.sqrt(variance(collectNums(a), false)),
  STDEVA: (a) => Math.sqrt(variance(collectNumsA(a), true)),
  STDEVPA: (a) => Math.sqrt(variance(collectNumsA(a), false)),
  VAR: (a) => variance(collectNums(a), true),
  'VAR.S': (a) => variance(collectNums(a), true),
  VARP: (a) => variance(collectNums(a), false),
  'VAR.P': (a) => variance(collectNums(a), false),
  VARA: (a) => variance(collectNumsA(a), true),
  VARPA: (a) => variance(collectNumsA(a), false),
  AVEDEV: (a) => { const n = collectNums(a); if (!n.length) throw ERR.NUM; const m = mean(n); return mean(n.map((x) => Math.abs(x - m))); },
  DEVSQ: (a) => { const n = collectNums(a); if (!n.length) throw ERR.NUM; const m = mean(n); return n.reduce((s, x) => s + (x - m) ** 2, 0); },
  GEOMEAN: (a) => { const n = collectNums(a); if (!n.length || n.some((x) => x <= 0)) throw ERR.NUM; return Math.exp(mean(n.map(Math.log))); },
  HARMEAN: (a) => { const n = collectNums(a); if (!n.length || n.some((x) => x <= 0)) throw ERR.NUM; return n.length / n.reduce((s, x) => s + 1 / x, 0); },
  KURT: (a) => {
    const n = collectNums(a);
    const k = n.length;
    if (k < 4) throw ERR.DIV0;
    const m = mean(n);
    const s = Math.sqrt(variance(n, true));
    if (!s) throw ERR.DIV0;
    const sum4 = n.reduce((t, x) => t + ((x - m) / s) ** 4, 0);
    return (k * (k + 1)) / ((k - 1) * (k - 2) * (k - 3)) * sum4 - (3 * (k - 1) ** 2) / ((k - 2) * (k - 3));
  },
  SKEW: (a) => {
    const n = collectNums(a);
    const k = n.length;
    if (k < 3) throw ERR.DIV0;
    const m = mean(n);
    const s = Math.sqrt(variance(n, true));
    if (!s) throw ERR.DIV0;
    return (k / ((k - 1) * (k - 2))) * n.reduce((t, x) => t + ((x - m) / s) ** 3, 0);
  },
  'SKEW.P': (a) => {
    const n = collectNums(a);
    if (n.length < 1) throw ERR.DIV0;
    const m = mean(n);
    const s = Math.sqrt(variance(n, false));
    if (!s) throw ERR.DIV0;
    return mean(n.map((x) => ((x - m) / s) ** 3));
  },
  TRIMMEAN: ([v, p]) => {
    const n = sortedNums(v);
    const pc = toNum(p);
    if (pc < 0 || pc >= 1) throw ERR.NUM;
    const cut = Math.floor((n.length * pc) / 2);
    return mean(n.slice(cut, n.length - cut));
  },
  LARGE: lift(([v, k]) => { const n = sortedNums(v); const i = Math.ceil(toNum(k)); if (i < 1 || i > n.length) throw ERR.NUM; return n[n.length - i]; }, [1]),
  SMALL: lift(([v, k]) => { const n = sortedNums(v); const i = Math.ceil(toNum(k)); if (i < 1 || i > n.length) throw ERR.NUM; return n[i - 1]; }, [1]),
  RANK: lift(([v, ref, order]) => rank(v, ref, order, false), [0]),
  'RANK.EQ': lift(([v, ref, order]) => rank(v, ref, order, false), [0]),
  'RANK.AVG': lift(([v, ref, order]) => rank(v, ref, order, true), [0]),
  PERCENTILE: lift(([v, p]) => percentileInc(sortedNums(v), toNum(p)), [1]),
  'PERCENTILE.INC': lift(([v, p]) => percentileInc(sortedNums(v), toNum(p)), [1]),
  'PERCENTILE.EXC': lift(([v, p]) => percentileExc(sortedNums(v), toNum(p)), [1]),
  QUARTILE: lift(([v, q]) => { const k = toInt(q); if (k < 0 || k > 4) throw ERR.NUM; return percentileInc(sortedNums(v), k / 4); }, [1]),
  'QUARTILE.INC': lift(([v, q]) => { const k = toInt(q); if (k < 0 || k > 4) throw ERR.NUM; return percentileInc(sortedNums(v), k / 4); }, [1]),
  'QUARTILE.EXC': lift(([v, q]) => { const k = toInt(q); if (k < 1 || k > 3) throw ERR.NUM; return percentileExc(sortedNums(v), k / 4); }, [1]),
  PERCENTRANK: lift(([v, x, sig]) => lookupPercentRank(sortedNums(v), toNum(x), optInt(sig, 3), false), [1]),
  'PERCENTRANK.INC': lift(([v, x, sig]) => lookupPercentRank(sortedNums(v), toNum(x), optInt(sig, 3), false), [1]),
  'PERCENTRANK.EXC': lift(([v, x, sig]) => lookupPercentRank(sortedNums(v), toNum(x), optInt(sig, 3), true), [1]),
  CORREL: ([x, y]) => correl(...pairs(x, y)),
  PEARSON: ([x, y]) => correl(...pairs(x, y)),
  RSQ: ([y, x]) => correl(...pairs(x, y)) ** 2,
  SLOPE: ([y, x]) => { const [xs, ys] = pairs(x, y); return linreg(ys, xs).slope; },
  INTERCEPT: ([y, x]) => { const [xs, ys] = pairs(x, y); return linreg(ys, xs).intercept; },
  STEYX: ([y, x]) => {
    const [xs, ys] = pairs(x, y);
    if (xs.length < 3) throw ERR.DIV0;
    const { slope, intercept } = linreg(ys, xs);
    return Math.sqrt(ys.reduce((s, yy, i) => s + (yy - (intercept + slope * xs[i])) ** 2, 0) / (xs.length - 2));
  },
  COVAR: ([x, y]) => { const [xs, ys] = pairs(x, y); if (!xs.length) throw ERR.DIV0; const mx = mean(xs); const my = mean(ys); return mean(xs.map((v, i) => (v - mx) * (ys[i] - my))); },
  'COVARIANCE.P': ([x, y]) => { const [xs, ys] = pairs(x, y); if (!xs.length) throw ERR.DIV0; const mx = mean(xs); const my = mean(ys); return mean(xs.map((v, i) => (v - mx) * (ys[i] - my))); },
  'COVARIANCE.S': ([x, y]) => {
    const [xs, ys] = pairs(x, y);
    if (xs.length < 2) throw ERR.DIV0;
    const mx = mean(xs);
    const my = mean(ys);
    return xs.reduce((s, v, i) => s + (v - mx) * (ys[i] - my), 0) / (xs.length - 1);
  },
  FORECAST: lift(([x, y, xk]) => { const [xs, ys] = pairs(xk, y); const { slope, intercept } = linreg(ys, xs); return intercept + slope * toNum(x); }, [0]),
  'FORECAST.LINEAR': lift(([x, y, xk]) => { const [xs, ys] = pairs(xk, y); const { slope, intercept } = linreg(ys, xs); return intercept + slope * toNum(x); }, [0]),
  'FORECAST.ETS': lift(([x, values, timeline, seas, comp, agg]) => {
    const m = etsModel(values, timeline, seas, comp, agg);
    return etsForecast(m, etsTarget(m, x));
  }, [0]),
  'FORECAST.ETS.CONFINT': lift(([x, values, timeline, cl, seas, comp, agg]) => {
    const conf = optNum(cl, 0.95);
    if (!(conf > 0 && conf < 1)) throw ERR.NUM;
    const m = etsModel(values, timeline, seas, comp, agg);
    return etsConfint(m, etsTarget(m, x), normInv((1 + conf) / 2));
  }, [0]),
  'FORECAST.ETS.SEASONALITY': ([values, timeline, comp, agg]) => etsModel(values, timeline, 1, comp, agg).m,
  'FORECAST.ETS.STAT': lift(([values, timeline, type, seas, comp, agg]) => {
    const k = toInt(type);
    const key = [null, 'alpha', 'beta', 'gamma', 'mase', 'smape', 'mae', 'rmse', 'step'][k];
    if (!key) throw ERR.NUM;
    return etsModel(values, timeline, seas, comp, agg).stats[key];
  }, [2]),
  TREND: ([y, x, nx]) => {
    const ys = collectNums([y]);
    const xs = x === undefined || x === null ? ys.map((_, i) => i + 1) : collectNums([x]);
    const { slope, intercept } = linreg(ys, xs);
    const target = nx === undefined || nx === null ? (x === undefined || x === null ? ys.map((_, i) => i + 1) : xs) : collectNums([nx]);
    const shape = asRange(nx ?? x ?? y);
    const vals = target.map((v) => intercept + slope * v);
    return shape.height === 1 ? new Range([vals]) : new Range(vals.map((v) => [v]));
  },
  GROWTH: ([y, x, nx]) => {
    const ys = collectNums([y]);
    if (ys.some((v) => v <= 0)) throw ERR.NUM;
    const xs = x === undefined || x === null ? ys.map((_, i) => i + 1) : collectNums([x]);
    const { slope, intercept } = linreg(ys.map(Math.log), xs);
    const target = nx === undefined || nx === null ? xs : collectNums([nx]);
    const shape = asRange(nx ?? x ?? y);
    const vals = target.map((v) => Math.exp(intercept + slope * v));
    return shape.height === 1 ? new Range([vals]) : new Range(vals.map((v) => [v]));
  },
  LINEST: ([y, x, konst, stats]) => {
    // 단순/다중 선형 회귀의 계수 (stats 가 TRUE 면 엑셀과 같은 5행 통계)
    const Y = asRange(y);
    const ys = collectNums([Y]);
    const useConst = konst === undefined || konst === null ? true : toBool(konst);
    let X;
    if (x === undefined || x === null) X = ys.map((_, i) => [i + 1]);
    else {
      const xr = asRange(x);
      X = xr.width === ys.length && xr.height !== ys.length ? xr.rows[0].map((_, j) => xr.rows.map((r) => r[j])) : xr.rows.map((r) => r.slice());
    }
    const k = X[0].length;
    const A = X.map((r) => (useConst ? [...r, 1] : r.slice()));
    const At = A[0].map((_, j) => A.map((r) => r[j]));
    const beta = mmult(inverse(mmult(At, A)), mmult(At, ys.map((v) => [v]))).map((r) => r[0]);
    const coef = beta.slice(0, k).reverse();
    const first = [...coef, useConst ? beta[k] : 0];
    if (stats === undefined || stats === null || !toBool(stats)) return new Range([first]);
    const n = ys.length;
    const p = k + (useConst ? 1 : 0);
    const df = n - p;
    const fit = A.map((r) => r.reduce((acc, v, j) => acc + v * beta[j], 0));
    const ssresid = ys.reduce((acc, v, i) => acc + (v - fit[i]) ** 2, 0);
    const my = mean(ys);
    const sstot = useConst ? ys.reduce((acc, v) => acc + (v - my) ** 2, 0) : ys.reduce((acc, v) => acc + v * v, 0);
    const ssreg = sstot - ssresid;
    const mse = df > 0 ? ssresid / df : NaN;
    const inv = inverse(mmult(At, A));
    const se = beta.map((_, j) => Math.sqrt(Math.max(0, mse * inv[j][j])));
    const NA = ERR.NA;
    const num = (v) => (Number.isFinite(v) ? v : ERR.NUM);
    const pad = (row) => [...row, ...Array(k + 1 - row.length).fill(NA)];
    return new Range([
      first,
      [...se.slice(0, k).reverse().map(num), useConst ? num(se[k]) : NA],
      pad([num(sstot ? ssreg / sstot : 1), num(Math.sqrt(mse))]),
      pad([num((ssreg / k) / mse), df]),
      pad([ssreg, ssresid]),
    ]);
  },
  FREQUENCY: ([data, bins]) => {
    const d = collectNums([data]);
    const b = collectNums([bins]);
    const order = b.map((v, i) => [v, i]).sort((p, q) => p[0] - q[0]);
    const counts = new Array(b.length + 1).fill(0);
    for (const v of d) {
      const j = order.findIndex(([edge]) => v <= edge);
      counts[j < 0 ? b.length : order[j][1]]++;
    }
    return new Range(counts.map((c) => [c]));
  },
  PROB: ([x, p, lo, hi]) => {
    const [xs, ps] = pairs(x, p);
    if (Math.abs(ps.reduce((s, v) => s + v, 0) - 1) > 1e-9) throw ERR.NUM;
    const a = toNum(lo);
    const b = hi === undefined || hi === null ? a : toNum(hi);
    return xs.reduce((s, v, i) => s + (v >= a && v <= b ? ps[i] : 0), 0);
  },
};
for (const [k, fn] of Object.entries(DIST)) STAT[k] = lift(fn);

// ───────────── 검정 · 지수 회귀 ─────────────
STAT.LOGEST = ([y, x, konst]) => {
  const ys = collectNums([y]);
  if (ys.some((v) => v <= 0)) throw ERR.NUM;
  const logs = new Range(ys.map((v) => [Math.log(v)]));
  const res = STAT.LINEST([logs, x, konst]);
  return new Range([res.rows[0].map(Math.exp)]);
};
STAT['T.TEST'] = ([a, b, tails, type]) => {
  const t = toInt(tails);
  const k = toInt(type);
  if ((t !== 1 && t !== 2) || k < 1 || k > 3) throw ERR.NUM;
  let stat;
  let df;
  if (k === 1) {
    const [xs, ys] = pairs(a, b);
    const d = xs.map((v, i) => v - ys[i]);
    if (d.length < 2) throw ERR.DIV0;
    stat = mean(d) / Math.sqrt(variance(d, true) / d.length);
    df = d.length - 1;
  } else {
    const xs = collectNums([a]);
    const ys = collectNums([b]);
    if (xs.length < 2 || ys.length < 2) throw ERR.DIV0;
    const va = variance(xs, true);
    const vb = variance(ys, true);
    if (k === 2) {
      df = xs.length + ys.length - 2;
      const sp = ((xs.length - 1) * va + (ys.length - 1) * vb) / df;
      stat = (mean(xs) - mean(ys)) / Math.sqrt(sp * (1 / xs.length + 1 / ys.length));
    } else {
      const sa = va / xs.length;
      const sb = vb / ys.length;
      stat = (mean(xs) - mean(ys)) / Math.sqrt(sa + sb);
      df = (sa + sb) ** 2 / (sa ** 2 / (xs.length - 1) + sb ** 2 / (ys.length - 1));
    }
  }
  const p = 1 - tCdf(Math.abs(stat), df);
  return t === 1 ? p : 2 * p;
};
STAT.TTEST = STAT['T.TEST'];
STAT['Z.TEST'] = ([arr, x, sigma]) => {
  const n = collectNums([arr]);
  if (!n.length) throw ERR.NA;
  const s = sigma === undefined || sigma === null ? Math.sqrt(variance(n, true)) : toNum(sigma);
  return 1 - normCdf((mean(n) - toNum(x)) / (s / Math.sqrt(n.length)));
};
STAT.ZTEST = STAT['Z.TEST'];
STAT['CHISQ.TEST'] = ([actual, expected]) => {
  const A = asRange(actual);
  const E = asRange(expected);
  if (A.height !== E.height || A.width !== E.width) throw ERR.NA;
  let chi = 0;
  A.rows.forEach((row, i) => row.forEach((v, j) => {
    const e = E.rows[i][j];
    if (typeof v !== 'number' || typeof e !== 'number') return;
    if (e === 0) throw ERR.DIV0;
    chi += (v - e) ** 2 / e;
  }));
  const df = A.height > 1 && A.width > 1 ? (A.height - 1) * (A.width - 1) : Math.max(A.height, A.width) - 1;
  if (df < 1) throw ERR.NA;
  return 1 - chiCdf(chi, df);
};
STAT.CHITEST = STAT['CHISQ.TEST'];
STAT['F.TEST'] = ([a, b]) => {
  const xs = collectNums([a]);
  const ys = collectNums([b]);
  if (xs.length < 2 || ys.length < 2) throw ERR.DIV0;
  const va = variance(xs, true);
  const vb = variance(ys, true);
  if (va === 0 || vb === 0) throw ERR.DIV0;
  const f = va / vb;
  const p = fCdf(f, xs.length - 1, ys.length - 1);
  return 2 * Math.min(p, 1 - p);
};
STAT.FTEST = STAT['F.TEST'];
