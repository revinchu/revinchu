// 재무 · 공학 · 데이터베이스 함수 (DOM 없음)
import {
  ERR, Range, isError, scalar, toNum, toStr, toInt, optNum, optInt, optBool, checkNum, collectNums, asRange,
  lift, makeCriteria, toDate, serialToDate, dateToSerial,
} from './fxcore.js';
import { STAT } from './fx-stat.js';
import { MATH } from './fx-math.js';
import { yearFrac } from './fx-date.js';

// ───────────── 화폐의 시간 가치 ─────────────
function fv(rate, nper, pmt, pv, type) {
  if (rate === 0) return -(pv + pmt * nper);
  const f = (1 + rate) ** nper;
  return -(pv * f + pmt * (1 + rate * type) * (f - 1) / rate);
}
function pv(rate, nper, pmt, fvv, type) {
  if (rate === 0) return -(fvv + pmt * nper);
  const f = (1 + rate) ** nper;
  return -(fvv + pmt * (1 + rate * type) * (f - 1) / rate) / f;
}
function pmt(rate, nper, pvv, fvv, type) {
  if (nper === 0) throw ERR.NUM;
  if (rate === 0) return -(pvv + fvv) / nper;
  const f = (1 + rate) ** nper;
  return -(rate * (pvv * f + fvv)) / ((1 + rate * type) * (f - 1));
}
function nper(rate, pm, pvv, fvv, type) {
  if (rate === 0) {
    if (pm === 0) throw ERR.NUM;
    return -(pvv + fvv) / pm;
  }
  const a = pm * (1 + rate * type) - fvv * rate;
  const b = pvv * rate + pm * (1 + rate * type);
  if (a / b <= 0) throw ERR.NUM;
  return Math.log(a / b) / Math.log(1 + rate);
}
function ipmt(rate, per, np, pvv, fvv, type) {
  if (per < 1 || per > np) throw ERR.NUM;
  const p = pmt(rate, np, pvv, fvv, type);
  if (per === 1 && type === 1) return 0;
  let interest = fv(rate, per - 1, p, pvv, type) * rate;
  if (type === 1) interest /= 1 + rate;
  return interest;
}
function rate(np, pm, pvv, fvv, type, guess) {
  let r = guess;
  for (let i = 0; i < 100; i++) {
    const f = (x) => (x === 0 ? pvv + pm * np + fvv : pvv * (1 + x) ** np + pm * (1 + x * type) * ((1 + x) ** np - 1) / x + fvv);
    const y = f(r);
    const h = 1e-7;
    const d = (f(r + h) - f(r - h)) / (2 * h);
    if (!Number.isFinite(d) || d === 0) break;
    const nr = r - y / d;
    if (Math.abs(nr - r) < 1e-10) return nr;
    r = nr;
  }
  throw ERR.NUM;
}

/** 뉴턴법 (IRR/XIRR) */
function newton(f, guess) {
  let x = guess;
  for (let i = 0; i < 200; i++) {
    const y = f(x);
    const h = 1e-7 * Math.max(1, Math.abs(x));
    const d = (f(x + h) - f(x - h)) / (2 * h);
    if (!Number.isFinite(y) || !Number.isFinite(d) || d === 0) break;
    const nx = x - y / d;
    if (Math.abs(nx - x) < 1e-12) return nx;
    x = nx;
    if (x <= -1) x = -0.999999;
  }
  // 이분법 보조
  let lo = -0.999999;
  let hi = 10;
  if (Math.sign(f(lo)) === Math.sign(f(hi))) throw ERR.NUM;
  for (let i = 0; i < 300; i++) {
    const mid = (lo + hi) / 2;
    if (Math.sign(f(mid)) === Math.sign(f(lo))) lo = mid; else hi = mid;
  }
  return (lo + hi) / 2;
}

function ddb(cost, salvage, life, period, factor) {
  let value = cost;
  let dep = 0;
  for (let i = 1; i <= Math.ceil(period); i++) {
    dep = Math.min(value * factor / life, Math.max(0, value - salvage));
    value -= dep;
  }
  return dep;
}

function vdb(cost, salvage, life, start, end, factor, noSwitch) {
  let total = 0;
  let value = cost;
  const steps = Math.ceil(end);
  for (let i = 0; i < steps; i++) {
    const ddbDep = Math.min(value * factor / life, Math.max(0, value - salvage));
    const sl = (value - salvage) / (life - i);
    const dep = !noSwitch && sl > ddbDep ? Math.max(0, sl) : ddbDep;
    const from = Math.max(start, i);
    const to = Math.min(end, i + 1);
    if (to > from) total += dep * (to - from);
    value -= dep;
  }
  return total;
}

// 채권: 이자 지급일 계산
function coupDates(settle, maturity, freq) {
  const m = serialToDate(maturity);
  const months = 12 / freq;
  let k = 0;
  let prev;
  let next;
  for (;;) {
    const t = new Date(Date.UTC(m.y, m.m - 1 - months * k, 1));
    const dim = new Date(Date.UTC(t.getUTCFullYear(), t.getUTCMonth() + 1, 0)).getUTCDate();
    const s = dateToSerial(t.getUTCFullYear(), t.getUTCMonth() + 1, Math.min(m.d, dim));
    if (s <= settle) { prev = s; break; }
    next = s;
    k++;
  }
  return { prev, next, n: k };
}
function coupDays(settle, maturity, freq, basis) {
  const { prev, next } = coupDates(settle, maturity, freq);
  if (basis === 1) return next - prev;
  if (basis === 3) return 365 / freq;
  return 360 / freq;
}
function daysBetween(a, b, basis) {
  if (basis === 0 || basis === 4) return yearFrac(a, b, basis) * 360;
  return b - a;
}

function bondPrice(settle, maturity, rateV, yld, redemption, freq, basis) {
  const { prev, next, n } = coupDates(settle, maturity, freq);
  const E = coupDays(settle, maturity, freq, basis);
  const A = daysBetween(prev, settle, basis);
  const DSC = basis === 0 || basis === 4 ? E - A : next - settle;
  const coupon = 100 * rateV / freq;
  if (n === 1) {
    const t = DSC / E;
    return (redemption + coupon) / (1 + t * yld / freq) - A / E * coupon;
  }
  let price = redemption / (1 + yld / freq) ** (n - 1 + DSC / E);
  for (let k = 1; k <= n; k++) price += coupon / (1 + yld / freq) ** (k - 1 + DSC / E);
  return price - A / E * coupon;
}

function checkFreq(f) {
  if (![1, 2, 4].includes(f)) throw ERR.NUM;
  return f;
}
function checkBasis(b) {
  if (b < 0 || b > 4) throw ERR.NUM;
  return b;
}

function cashflows(v) {
  const out = [];
  for (const x of asRange(v).values()) {
    if (isError(x)) throw x;
    if (typeof x === 'number') out.push(x);
  }
  return out;
}

// ───────────── 진법 변환 ─────────────
const BASES = { BIN: 2, OCT: 8, DEC: 10, HEX: 16 };
function fromBase(text, base) {
  const s = toStr(text).trim().toUpperCase();
  if (s.length > 10) throw ERR.NUM;
  if (s === '') return 0;
  const valid = { 2: /^[01]+$/, 8: /^[0-7]+$/, 10: /^-?\d+$/, 16: /^[0-9A-F]+$/ }[base];
  if (!valid.test(s)) throw ERR.NUM;
  let n = parseInt(s, base);
  if (base !== 10 && s.length === 10) {
    const bits = { 2: 10, 8: 30, 16: 40 }[base];
    if (n >= 2 ** (bits - 1)) n -= 2 ** bits;
  }
  return n;
}
function toBase(n, base, places) {
  const limits = { 2: 512, 8: 536870912, 16: 549755813888 };
  n = Math.trunc(n);
  if (base !== 10 && (n < -limits[base] || n >= limits[base])) throw ERR.NUM;
  let s;
  if (n < 0) {
    const bits = { 2: 10, 8: 30, 16: 40 }[base];
    s = (2 ** bits + n).toString(base).toUpperCase();
    return s;
  }
  s = n.toString(base).toUpperCase();
  if (places !== undefined && places !== null) {
    const p = toInt(places);
    if (p < s.length || p > 10) throw ERR.NUM;
    s = s.padStart(p, '0');
  }
  return s;
}

// CONVERT 단위 (기준 단위로의 배율)
const UNITS = {
  mass: { g: 1, kg: 1000, mg: 1e-3, lbm: 453.59237, ozm: 28.349523125, u: 1.66053906660e-24, grain: 0.06479891, stone: 6350.29318, ton: 907184.74, uk_ton: 1016046.9088, sg: 14593.90294 },
  length: { m: 1, km: 1000, cm: 0.01, mm: 0.001, um: 1e-6, nm: 1e-9, mi: 1609.344, Nmi: 1852, in: 0.0254, ft: 0.3048, yd: 0.9144, ang: 1e-10, ell: 1.143, ly: 9.46073047258e15, parsec: 3.08567758149137e16, pc: 3.08567758149137e16, Pica: 0.0254 / 6, pica: 0.0254 / 72, survey_mi: 1609.347219 },
  time: { sec: 1, s: 1, min: 60, mn: 60, hr: 3600, day: 86400, d: 86400, yr: 31557600 },
  pressure: { Pa: 1, p: 1, atm: 101325, at: 101325, mmHg: 133.322, psi: 6894.757293168, Torr: 133.322368421 },
  force: { N: 1, dyn: 1e-5, dy: 1e-5, lbf: 4.4482216152605, pond: 0.00980665 },
  energy: { J: 1, e: 1e-7, c: 4.184, cal: 4.1868, eV: 1.602176634e-19, ev: 1.602176634e-19, HPh: 2684519.537696, hh: 2684519.537696, Wh: 3600, wh: 3600, flb: 1.3558179483314, BTU: 1055.05585262, btu: 1055.05585262 },
  power: { W: 1, w: 1, HP: 745.69987158227, h: 745.69987158227, PS: 735.49875 },
  magnetism: { T: 1, ga: 1e-4 },
  speed: { 'm/s': 1, 'm/sec': 1, 'm/h': 1 / 3600, 'm/hr': 1 / 3600, mph: 0.44704, kn: 1852 / 3600, admkn: 0.514773333 },
  volume: { l: 0.001, L: 0.001, lt: 0.001, ml: 1e-6, m3: 1, tsp: 4.92892159375e-6, tspm: 5e-6, tbs: 1.478676478125e-5, oz: 2.95735295625e-5, cup: 2.365882365e-4, pt: 4.73176473e-4, us_pt: 4.73176473e-4, uk_pt: 5.6826125e-4, qt: 9.46352946e-4, uk_qt: 1.1365225e-3, gal: 3.785411784e-3, uk_gal: 4.54609e-3, ft3: 0.028316846592, in3: 1.6387064e-5, yd3: 0.764554857984, barrel: 0.158987294928 },
  area: { m2: 1, km2: 1e6, cm2: 1e-4, mm2: 1e-6, ha: 1e4, ar: 100, 'uk_acre': 4046.8564224, 'us_acre': 4046.87261, ft2: 0.09290304, in2: 6.4516e-4, yd2: 0.83612736, mi2: 2589988.110336 },
  info: { bit: 1, byte: 8 },
};
const PREFIX = { Y: 1e24, Z: 1e21, E: 1e18, P: 1e15, T: 1e12, G: 1e9, M: 1e6, k: 1e3, h: 1e2, da: 10, d: 0.1, c: 0.01, m: 1e-3, u: 1e-6, n: 1e-9, p: 1e-12, f: 1e-15, a: 1e-18, z: 1e-21, y: 1e-24, Ki: 1024, Mi: 1024 ** 2, Gi: 1024 ** 3, Ti: 1024 ** 4 };
const TEMP = ['C', 'cel', 'F', 'fah', 'K', 'kel', 'Rank', 'Reau'];

function unitInfo(u) {
  for (const [kind, table] of Object.entries(UNITS)) {
    if (u in table) return { kind, f: table[u] };
  }
  for (const [p, pf] of Object.entries(PREFIX)) {
    if (!u.startsWith(p)) continue;
    const base = u.slice(p.length);
    for (const [kind, table] of Object.entries(UNITS)) {
      if (base in table) {
        // 제곱/세제곱 단위는 배율도 거듭제곱
        const pow = /[23]$/.test(base) && kind !== 'info' ? Number(base.slice(-1)) : 1;
        return { kind, f: table[base] * pf ** pow };
      }
    }
  }
  return null;
}

function toKelvin(v, u) {
  switch (u) {
    case 'C': case 'cel': return v + 273.15;
    case 'F': case 'fah': return (v - 32) * 5 / 9 + 273.15;
    case 'K': case 'kel': return v;
    case 'Rank': return v * 5 / 9;
    case 'Reau': return v * 1.25 + 273.15;
  }
  throw ERR.NA;
}
function fromKelvin(k, u) {
  switch (u) {
    case 'C': case 'cel': return k - 273.15;
    case 'F': case 'fah': return (k - 273.15) * 9 / 5 + 32;
    case 'K': case 'kel': return k;
    case 'Rank': return k * 9 / 5;
    case 'Reau': return (k - 273.15) * 0.8;
  }
  throw ERR.NA;
}

// ───────────── 데이터베이스 ─────────────
function dbSelect(database, field, criteria) {
  const db = asRange(database).rows;
  const crit = asRange(criteria).rows;
  const heads = db[0].map((h) => String(h ?? '').toLowerCase());
  let col = null;
  if (field !== undefined && field !== null) {
    const f = scalar(field);
    if (typeof f === 'number') col = Math.trunc(f) - 1;
    else col = heads.indexOf(String(f).toLowerCase());
    if (col < 0 || col >= heads.length) throw ERR.VALUE;
  }
  const critHeads = crit[0].map((h) => heads.indexOf(String(h ?? '').toLowerCase()));
  const conds = crit.slice(1).map((row) => row.map((v, j) => (v === null || v === '' ? null : { j: critHeads[j], test: makeCriteria(v) })).filter(Boolean));
  const out = [];
  for (const row of db.slice(1)) {
    const ok = !conds.length || conds.some((cs) => cs.every(({ j, test }) => j >= 0 && test(row[j])));
    if (ok) out.push(col === null ? row : row[col]);
  }
  return out;
}
const dbFn = (name) => ([db, field, crit]) => {
  const vals = dbSelect(db, field, crit);
  return (STAT[name] ?? MATH[name])([new Range([vals.length ? vals : [null]])]);
};

const SCALAR = {
  FV: ([r, n, p, v, t]) => checkNum(fv(toNum(r), toNum(n), toNum(p), optNum(v, 0), optInt(t, 0) ? 1 : 0)),
  PV: ([r, n, p, v, t]) => checkNum(pv(toNum(r), toNum(n), toNum(p), optNum(v, 0), optInt(t, 0) ? 1 : 0)),
  PMT: ([r, n, v, f, t]) => checkNum(pmt(toNum(r), toNum(n), toNum(v), optNum(f, 0), optInt(t, 0) ? 1 : 0)),
  NPER: ([r, p, v, f, t]) => checkNum(nper(toNum(r), toNum(p), toNum(v), optNum(f, 0), optInt(t, 0) ? 1 : 0)),
  RATE: ([n, p, v, f, t, g]) => checkNum(rate(toNum(n), toNum(p), toNum(v), optNum(f, 0), optInt(t, 0) ? 1 : 0, optNum(g, 0.1))),
  IPMT: ([r, per, n, v, f, t]) => checkNum(ipmt(toNum(r), toNum(per), toNum(n), toNum(v), optNum(f, 0), optInt(t, 0) ? 1 : 0)),
  PPMT: ([r, per, n, v, f, t]) => {
    const rr = toNum(r);
    const nn = toNum(n);
    const pvv = toNum(v);
    const fvv = optNum(f, 0);
    const type = optInt(t, 0) ? 1 : 0;
    return checkNum(pmt(rr, nn, pvv, fvv, type) - ipmt(rr, toNum(per), nn, pvv, fvv, type));
  },
  CUMIPMT: ([r, n, v, s, e, t]) => {
    const rr = toNum(r); const nn = toInt(n); const pvv = toNum(v); const a = toInt(s); const b = toInt(e); const type = toInt(t);
    if (rr <= 0 || nn <= 0 || pvv <= 0 || a < 1 || b < a || b > nn || (type !== 0 && type !== 1)) throw ERR.NUM;
    let sum = 0;
    for (let k = a; k <= b; k++) sum += ipmt(rr, k, nn, pvv, 0, type);
    return sum;
  },
  CUMPRINC: ([r, n, v, s, e, t]) => {
    const rr = toNum(r); const nn = toInt(n); const pvv = toNum(v); const a = toInt(s); const b = toInt(e); const type = toInt(t);
    if (rr <= 0 || nn <= 0 || pvv <= 0 || a < 1 || b < a || b > nn || (type !== 0 && type !== 1)) throw ERR.NUM;
    const p = pmt(rr, nn, pvv, 0, type);
    let sum = 0;
    for (let k = a; k <= b; k++) sum += p - ipmt(rr, k, nn, pvv, 0, type);
    return sum;
  },
  ISPMT: ([r, per, n, v]) => {
    const pvv = toNum(v);
    const nn = toNum(n);
    if (nn === 0) throw ERR.DIV0;
    return -pvv * toNum(r) * (1 - toNum(per) / nn);
  },
  SLN: ([c, s, l]) => {
    const life = toNum(l);
    if (life === 0) throw ERR.DIV0;
    return (toNum(c) - toNum(s)) / life;
  },
  SYD: ([c, s, l, p]) => {
    const life = toNum(l);
    const per = toNum(p);
    if (life <= 0 || per < 1 || per > life) throw ERR.NUM;
    return ((toNum(c) - toNum(s)) * (life - per + 1) * 2) / (life * (life + 1));
  },
  DDB: ([c, s, l, p, f]) => {
    const cost = toNum(c); const salvage = toNum(s); const life = toNum(l); const per = toNum(p); const factor = optNum(f, 2);
    if (cost < 0 || salvage < 0 || life <= 0 || per <= 0 || per > life || factor <= 0) throw ERR.NUM;
    return ddb(cost, salvage, life, per, factor);
  },
  DB: ([c, s, l, p, m]) => {
    const cost = toNum(c); const salvage = toNum(s); const life = toInt(l); const per = toInt(p); const month = optInt(m, 12);
    if (cost < 0 || salvage < 0 || life <= 0 || per <= 0 || month < 1 || month > 12 || per > life + 1) throw ERR.NUM;
    if (cost === 0) return 0;
    const rate0 = Math.round((1 - (salvage / cost) ** (1 / life)) * 1000) / 1000;
    let total = cost * rate0 * month / 12;
    if (per === 1) return total;
    let dep = 0;
    for (let i = 2; i <= per; i++) {
      dep = i === life + 1 ? (cost - total) * rate0 * (12 - month) / 12 : (cost - total) * rate0;
      total += dep;
    }
    return dep;
  },
  VDB: ([c, s, l, a, b, f, ns]) => {
    const cost = toNum(c); const salvage = toNum(s); const life = toNum(l); const start = toNum(a); const end = toNum(b);
    if (cost < 0 || salvage < 0 || life <= 0 || start < 0 || end < start || end > life) throw ERR.NUM;
    return vdb(cost, salvage, life, start, end, optNum(f, 2), optBool(ns, false));
  },
  EFFECT: ([r, n]) => {
    const rr = toNum(r); const np = toInt(n);
    if (rr <= 0 || np < 1) throw ERR.NUM;
    return (1 + rr / np) ** np - 1;
  },
  NOMINAL: ([r, n]) => {
    const rr = toNum(r); const np = toInt(n);
    if (rr <= 0 || np < 1) throw ERR.NUM;
    return np * ((1 + rr) ** (1 / np) - 1);
  },
  PDURATION: ([r, p, f]) => {
    const rr = toNum(r); const a = toNum(p); const b = toNum(f);
    if (rr <= 0 || a <= 0 || b <= 0) throw ERR.NUM;
    return (Math.log(b) - Math.log(a)) / Math.log(1 + rr);
  },
  RRI: ([n, p, f]) => {
    const nn = toNum(n); const a = toNum(p); const b = toNum(f);
    if (nn <= 0 || a === 0) throw ERR.NUM;
    return (b / a) ** (1 / nn) - 1;
  },
  DOLLARDE: ([d, f]) => {
    const x = toNum(d); const frac = toInt(f);
    if (frac < 0) throw ERR.NUM;
    if (frac === 0) throw ERR.DIV0;
    const ip = Math.trunc(x);
    const digits = 10 ** Math.ceil(Math.log10(frac));
    return ip + ((x - ip) * digits) / frac;
  },
  DOLLARFR: ([d, f]) => {
    const x = toNum(d); const frac = toInt(f);
    if (frac < 0) throw ERR.NUM;
    if (frac === 0) throw ERR.DIV0;
    const ip = Math.trunc(x);
    const digits = 10 ** Math.ceil(Math.log10(frac));
    return ip + ((x - ip) * frac) / digits;
  },
  TBILLPRICE: ([s, m, d]) => {
    const a = toDate(s); const b = toDate(m); const disc = toNum(d);
    if (a >= b || b - a > 366 || disc <= 0) throw ERR.NUM;
    return 100 * (1 - disc * (b - a) / 360);
  },
  TBILLYIELD: ([s, m, p]) => {
    const a = toDate(s); const b = toDate(m); const pr = toNum(p);
    if (a >= b || b - a > 366 || pr <= 0) throw ERR.NUM;
    return ((100 - pr) / pr) * (360 / (b - a));
  },
  TBILLEQ: ([s, m, d]) => {
    const a = toDate(s); const b = toDate(m); const disc = toNum(d);
    if (a >= b || b - a > 366 || disc <= 0) throw ERR.NUM;
    return (365 * disc) / (360 - disc * (b - a));
  },
  DISC: ([s, m, p, r, b]) => {
    const a = toDate(s); const e = toDate(m); const pr = toNum(p); const red = toNum(r);
    if (a >= e || pr <= 0 || red <= 0) throw ERR.NUM;
    return ((red - pr) / red) / yearFrac(a, e, checkBasis(optInt(b, 0)));
  },
  INTRATE: ([s, m, inv, r, b]) => {
    const a = toDate(s); const e = toDate(m); const i = toNum(inv); const red = toNum(r);
    if (a >= e || i <= 0 || red <= 0) throw ERR.NUM;
    return ((red - i) / i) / yearFrac(a, e, checkBasis(optInt(b, 0)));
  },
  RECEIVED: ([s, m, inv, d, b]) => {
    const a = toDate(s); const e = toDate(m); const i = toNum(inv); const disc = toNum(d);
    if (a >= e || i <= 0 || disc <= 0) throw ERR.NUM;
    return i / (1 - disc * yearFrac(a, e, checkBasis(optInt(b, 0))));
  },
  PRICEDISC: ([s, m, d, r, b]) => {
    const a = toDate(s); const e = toDate(m); const disc = toNum(d); const red = toNum(r);
    if (a >= e || disc <= 0 || red <= 0) throw ERR.NUM;
    return red - disc * red * yearFrac(a, e, checkBasis(optInt(b, 0)));
  },
  YIELDDISC: ([s, m, p, r, b]) => {
    const a = toDate(s); const e = toDate(m); const pr = toNum(p); const red = toNum(r);
    if (a >= e || pr <= 0 || red <= 0) throw ERR.NUM;
    return ((red - pr) / pr) / yearFrac(a, e, checkBasis(optInt(b, 0)));
  },
  ACCRINTM: ([i, s, r, par, b]) => {
    const a = toDate(i); const e = toDate(s); const rt = toNum(r); const p = optNum(par, 1000);
    if (a >= e || rt <= 0 || p <= 0) throw ERR.NUM;
    return p * rt * yearFrac(a, e, checkBasis(optInt(b, 0)));
  },
  PRICE: ([s, m, r, y, red, f, b]) => {
    const a = toDate(s); const e = toDate(m);
    if (a >= e || toNum(r) < 0 || toNum(y) < 0 || toNum(red) <= 0) throw ERR.NUM;
    return bondPrice(a, e, toNum(r), toNum(y), toNum(red), checkFreq(toInt(f)), checkBasis(optInt(b, 0)));
  },
  YIELD: ([s, m, r, p, red, f, b]) => {
    const a = toDate(s); const e = toDate(m); const pr = toNum(p);
    const fr = checkFreq(toInt(f)); const bs = checkBasis(optInt(b, 0));
    if (a >= e || toNum(r) < 0 || pr <= 0 || toNum(red) <= 0) throw ERR.NUM;
    return newton((y) => bondPrice(a, e, toNum(r), y, toNum(red), fr, bs) - pr, 0.05);
  },
  DURATION: ([s, m, c, y, f, b]) => {
    const a = toDate(s); const e = toDate(m); const cp = toNum(c); const yl = toNum(y);
    const fr = checkFreq(toInt(f)); const bs = checkBasis(optInt(b, 0));
    if (a >= e || cp < 0 || yl < 0) throw ERR.NUM;
    const { n } = coupDates(a, e, fr);
    const E = coupDays(a, e, fr, bs);
    const { prev } = coupDates(a, e, fr);
    const dsc = (E - daysBetween(prev, a, bs)) / E;
    let num = 0;
    let den = 0;
    for (let k = 1; k <= n; k++) {
      const t = k - 1 + dsc;
      const cf = (k === n ? 100 : 0) + 100 * cp / fr;
      const df = (1 + yl / fr) ** t;
      num += (t / fr) * cf / df;
      den += cf / df;
    }
    return num / den;
  },
  MDURATION: ([s, m, c, y, f, b]) => {
    const d = SCALAR.DURATION([s, m, c, y, f, b]);
    return d / (1 + toNum(y) / toInt(f));
  },
  COUPNUM: ([s, m, f]) => coupDates(toDate(s), toDate(m), checkFreq(toInt(f))).n,
  COUPPCD: ([s, m, f]) => coupDates(toDate(s), toDate(m), checkFreq(toInt(f))).prev,
  COUPNCD: ([s, m, f]) => coupDates(toDate(s), toDate(m), checkFreq(toInt(f))).next,
  COUPDAYS: ([s, m, f, b]) => coupDays(toDate(s), toDate(m), checkFreq(toInt(f)), checkBasis(optInt(b, 0))),
  COUPDAYBS: ([s, m, f, b]) => {
    const a = toDate(s);
    const { prev } = coupDates(a, toDate(m), checkFreq(toInt(f)));
    return daysBetween(prev, a, checkBasis(optInt(b, 0)));
  },
  COUPDAYSNC: ([s, m, f, b]) => {
    const a = toDate(s); const e = toDate(m); const fr = checkFreq(toInt(f)); const bs = checkBasis(optInt(b, 0));
    const { prev, next } = coupDates(a, e, fr);
    if (bs === 0 || bs === 4) return coupDays(a, e, fr, bs) - daysBetween(prev, a, bs);
    return next - a;
  },

  // 공학
  DELTA: ([a, b]) => (toNum(a) === optNum(b, 0) ? 1 : 0),
  GESTEP: ([a, b]) => (toNum(a) >= optNum(b, 0) ? 1 : 0),
  BITAND: ([a, b]) => Number(BigInt(bit(a)) & BigInt(bit(b))),
  BITOR: ([a, b]) => Number(BigInt(bit(a)) | BigInt(bit(b))),
  BITXOR: ([a, b]) => Number(BigInt(bit(a)) ^ BigInt(bit(b))),
  BITLSHIFT: ([a, n]) => {
    const k = toInt(n);
    if (Math.abs(k) > 53) throw ERR.NUM;
    const v = k >= 0 ? BigInt(bit(a)) << BigInt(k) : BigInt(bit(a)) >> BigInt(-k);
    if (v >= 2n ** 48n) throw ERR.NUM;
    return Number(v);
  },
  BITRSHIFT: ([a, n]) => {
    const k = toInt(n);
    if (Math.abs(k) > 53) throw ERR.NUM;
    const v = k >= 0 ? BigInt(bit(a)) >> BigInt(k) : BigInt(bit(a)) << BigInt(-k);
    if (v >= 2n ** 48n) throw ERR.NUM;
    return Number(v);
  },
  CONVERT: ([n, from, to]) => {
    const x = toNum(n);
    const f = toStr(from);
    const t = toStr(to);
    if (TEMP.includes(f) || TEMP.includes(t)) {
      if (!TEMP.includes(f) || !TEMP.includes(t)) throw ERR.NA;
      return fromKelvin(toKelvin(x, f), t);
    }
    const a = unitInfo(f);
    const b = unitInfo(t);
    if (!a || !b || a.kind !== b.kind) throw ERR.NA;
    return (x * a.f) / b.f;
  },
};

function bit(v) {
  const n = toNum(v);
  if (n < 0 || n >= 2 ** 48 || n !== Math.trunc(n)) throw ERR.NUM;
  return n;
}

for (const from of Object.keys(BASES)) {
  for (const to of Object.keys(BASES)) {
    if (from === to) continue;
    SCALAR[`${from}2${to}`] = ([v, places]) => {
      const n = from === 'DEC' ? Math.trunc(toNum(v)) : fromBase(v, BASES[from]);
      if (to === 'DEC') return n;
      return toBase(n, BASES[to], places);
    };
  }
}

export const FIN = {
  NPV: ([r, ...vals]) => {
    const rr = toNum(r);
    if (rr === -1) throw ERR.DIV0;
    return collectNums(vals).reduce((s, v, i) => s + v / (1 + rr) ** (i + 1), 0);
  },
  IRR: ([vals, guess]) => {
    const cf = cashflows(vals);
    if (!cf.some((v) => v > 0) || !cf.some((v) => v < 0)) throw ERR.NUM;
    return newton((r) => cf.reduce((s, v, i) => s + v / (1 + r) ** i, 0), optNum(guess, 0.1));
  },
  MIRR: ([vals, fr, rr]) => {
    const cf = cashflows(vals);
    const f = toNum(fr);
    const r = toNum(rr);
    const n = cf.length;
    const pos = cf.reduce((s, v, i) => s + (v > 0 ? v * (1 + r) ** (n - 1 - i) : 0), 0);
    const neg = cf.reduce((s, v, i) => s + (v < 0 ? v / (1 + f) ** i : 0), 0);
    if (pos === 0 || neg === 0) throw ERR.DIV0;
    return (-pos / neg) ** (1 / (n - 1)) - 1;
  },
  XNPV: ([r, vals, dates]) => {
    const rr = toNum(r);
    const cf = [...asRange(vals).values()].map(toNum);
    const ds = [...asRange(dates).values()].map((d) => Math.floor(toNum(d)));
    if (cf.length !== ds.length) throw ERR.NUM;
    return cf.reduce((s, v, i) => s + v / (1 + rr) ** ((ds[i] - ds[0]) / 365), 0);
  },
  XIRR: ([vals, dates, guess]) => {
    const cf = [...asRange(vals).values()].map(toNum);
    const ds = [...asRange(dates).values()].map((d) => Math.floor(toNum(d)));
    if (cf.length !== ds.length || !cf.some((v) => v > 0) || !cf.some((v) => v < 0)) throw ERR.NUM;
    return newton((r) => cf.reduce((s, v, i) => s + v / (1 + r) ** ((ds[i] - ds[0]) / 365), 0), optNum(guess, 0.1));
  },
  FVSCHEDULE: ([p, sched]) => {
    let v = toNum(p);
    for (const r of asRange(sched).values()) {
      if (isError(r)) throw r;
      v *= 1 + (typeof r === 'number' ? r : r === null ? 0 : toNum(r));
    }
    return v;
  },
  DSUM: dbFn('SUM'),
  DAVERAGE: dbFn('AVERAGE'),
  DCOUNT: dbFn('COUNT'),
  DCOUNTA: dbFn('COUNTA'),
  DMAX: dbFn('MAX'),
  DMIN: dbFn('MIN'),
  DPRODUCT: dbFn('PRODUCT'),
  DSTDEV: dbFn('STDEV'),
  DSTDEVP: dbFn('STDEVP'),
  DVAR: dbFn('VAR'),
  DVARP: dbFn('VARP'),
  DGET: ([db, field, crit]) => {
    const vals = dbSelect(db, field, crit);
    if (!vals.length) throw ERR.VALUE;
    if (vals.length > 1) throw ERR.NUM;
    return vals[0] ?? 0;
  },
};
// DCOUNT 는 필드 생략 가능
FIN.DCOUNT = ([db, field, crit]) => {
  const vals = dbSelect(db, field, crit);
  if (field === undefined || field === null) return vals.length;
  return vals.filter((v) => typeof v === 'number').length;
};
FIN.DCOUNTA = ([db, field, crit]) => {
  const vals = dbSelect(db, field, crit);
  if (field === undefined || field === null) return vals.length;
  return vals.filter((v) => v !== null && v !== '').length;
};
for (const [k, fn] of Object.entries(SCALAR)) FIN[k] = lift(fn);

