// 지수 평활 예측 (엑셀 FORECAST.ETS 계열: AAA — 가법 오차 · 가법 추세 · 가법 계절성) — DOM 없음, 의존성 없음
// 시간 표시줄을 일정한 간격으로 맞추고(누락 보간 · 중복 집계), 계절 길이를 찾고, α·β·γ 를 오차 제곱합 최소로 맞춤

/** 중복 시점 집계: 1 AVERAGE, 2 COUNT, 3 COUNTA, 4 MAX, 5 MEDIAN, 6 MIN, 7 SUM */
function aggregate(list, how) {
  const n = list.length;
  switch (how) {
    case 2: case 3: return n;
    case 4: return Math.max(...list);
    case 5: { const s = [...list].sort((a, b) => a - b); return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2; }
    case 6: return Math.min(...list);
    case 7: return list.reduce((a, b) => a + b, 0);
    default: return list.reduce((a, b) => a + b, 0) / n;
  }
}

// 엑셀 날짜 일련번호 ↔ 연 · 월 · 일 (1900 기준, UTC)
const EPOCH = Date.UTC(1899, 11, 30);
const ymd = (t) => { const d = new Date(EPOCH + Math.round(t) * 86400000); return [d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()]; };
const serial = (y, m, d) => (Date.UTC(y, m, d) - EPOCH) / 86400000;
const monthDays = (y, m) => new Date(Date.UTC(y, m + 1, 0)).getUTCDate();

/**
 * 시간 축: 일정한 간격(step)으로 떨어진 시점들. 날짜가 매월 같은 날(또는 월말)이면 '월' 단위 (엑셀처럼 1·3·12개월 간격 인식).
 * 반환 { pos(t) → 시작점에서 몇 단계인지(실수), at(k) → k 단계 뒤 시점, step, months } | null
 */
export function timeAxis(ts) {
  const t0 = ts[0];
  let step = Infinity;
  for (let i = 1; i < ts.length; i++) step = Math.min(step, ts[i] - ts[i - 1]);
  if (!(step > 0)) return null;
  const even = ts.every((t, i) => !i || Math.abs((t - ts[i - 1]) / step - Math.round((t - ts[i - 1]) / step)) < 1e-6);
  if (even) return { step, months: 0, pos: (t) => (t - t0) / step, at: (k) => t0 + k * step };
  // 월 단위: 모두 정수 날짜이고 같은 날(또는 모두 월말)
  if (!ts.every((t) => Number.isInteger(t) && t > 0)) return null;
  const parts = ts.map(ymd);
  const endOfMonth = parts.every(([y, m, d]) => d === monthDays(y, m));
  const day = parts[0][2];
  if (!endOfMonth && !parts.every((p) => p[2] === day)) return null;
  const mi = parts.map(([y, m]) => y * 12 + m);
  let ms = Infinity;
  for (let i = 1; i < mi.length; i++) ms = Math.min(ms, mi[i] - mi[i - 1]);
  if (!(ms > 0) || mi.some((m, i) => i && (m - mi[i - 1]) % ms)) return null;
  const m0 = mi[0];
  const dateAt = (m) => { const y = Math.floor(m / 12); const mm = m - y * 12; return serial(y, mm, endOfMonth ? monthDays(y, mm) : Math.min(day, monthDays(y, mm))); };
  return {
    step: ms, months: ms,
    pos: (t) => {
      // 달 사이의 날짜는 앞뒤 기준일 사이를 직선 보간
      const [y, m] = ymd(t);
      let k = y * 12 + m;
      if (dateAt(k) > t) k--;
      const a = dateAt(k); const b = dateAt(k + 1);
      return (k - m0 + (t - a) / (b - a)) / ms;
    },
    at: (k) => dateAt(m0 + Math.round(k * ms)),
  };
}

/**
 * 값 · 시간 표시줄 → 일정 간격 계열. values[i] 가 null 이면 누락(보간 대상).
 * 반환 { y: number[], axis } | { error: 'NUM' | 'VALUE' }
 */
export function prepareSeries(values, timeline, { completion = 1, aggregation = 1 } = {}) {
  if (values.length !== timeline.length || values.length < 2) return { error: 'VALUE' };
  const byT = new Map();
  for (let i = 0; i < timeline.length; i++) {
    const t = timeline[i];
    if (typeof t !== 'number' || !Number.isFinite(t)) return { error: 'VALUE' };
    const v = values[i];
    if (!byT.has(t)) byT.set(t, []);
    if (typeof v === 'number' && Number.isFinite(v)) byT.get(t).push(v);
  }
  const ts = [...byT.keys()].sort((a, b) => a - b);
  if (ts.length < 2) return { error: 'NUM' };
  const axis = timeAxis(ts);
  if (!axis) return { error: 'NUM' };
  const n = Math.round(axis.pos(ts[ts.length - 1])) + 1;
  if (n > 1e6) return { error: 'NUM' };
  const y = new Array(n).fill(null);
  for (const t of ts) {
    const list = byT.get(t);
    if (list.length) y[Math.round(axis.pos(t))] = aggregate(list, aggregation);
  }
  // 누락: completion 1 이면 앞뒤 값으로 직선 보간, 0 이면 0
  let prev = -1;
  for (let i = 0; i < n; i++) {
    if (y[i] === null) continue;
    if (prev < 0) { for (let j = 0; j < i; j++) y[j] = completion ? y[i] : 0; }
    else if (i - prev > 1) for (let j = prev + 1; j < i; j++) y[j] = completion ? y[prev] + ((y[i] - y[prev]) * (j - prev)) / (i - prev) : 0;
    prev = i;
  }
  if (prev < 0) return { error: 'NUM' };
  for (let j = prev + 1; j < n; j++) y[j] = completion ? y[prev] : 0;
  return { y, axis };
}

/** 계절 길이 자동 찾기: 추세를 뺀 계열의 자기상관이 가장 큰 주기 (없으면 0) */
export function detectSeason(y) {
  const n = y.length;
  if (n < 6) return 0;
  // 선형 추세 제거
  let sx = 0; let sy = 0; let sxx = 0; let sxy = 0;
  for (let i = 0; i < n; i++) { sx += i; sy += y[i]; sxx += i * i; sxy += i * y[i]; }
  const b = (n * sxy - sx * sy) / (n * sxx - sx * sx || 1);
  const a = (sy - b * sx) / n;
  const d = y.map((v, i) => v - (a + b * i));
  let den = 0;
  for (const v of d) den += v * v;
  if (den <= 1e-12) return 0;
  const maxLag = Math.min(Math.floor(n / 2), 8760);
  const acf = [1];
  for (let k = 1; k <= maxLag; k++) {
    let s = 0;
    for (let i = k; i < n; i++) s += d[i] * d[i - k];
    acf[k] = s / den;
  }
  let best = 0;
  let bestV = 0.25; // 이보다 약한 주기는 계절성으로 보지 않음
  for (let k = 2; k <= maxLag; k++) {
    // 자기상관의 봉우리이고, 주기가 두 번 이상 들어가야 함
    if (acf[k] > bestV && acf[k] >= acf[k - 1] && (k === maxLag || acf[k] >= acf[k + 1]) && n >= 2 * k) { best = k; bestV = acf[k]; }
  }
  return best;
}

/** AAA 상태 공간 재귀: 한 단계 예측 오차들과 마지막 상태 */
function run(y, m, alpha, beta, gamma, keep = false) {
  const n = y.length;
  let l;
  let b;
  const s = new Float64Array(Math.max(m, 1));
  let start;
  if (m > 1 && n >= 2 * m) {
    let m1 = 0; let m2 = 0;
    for (let i = 0; i < m; i++) { m1 += y[i]; m2 += y[i + m]; }
    m1 /= m; m2 /= m;
    b = (m2 - m1) / m;
    l = m1 - b * (m - 1) / 2; // 첫 주기 가운데 → 시작점
    for (let i = 0; i < m; i++) s[i] = y[i] - (l + b * i);
    l += b * (m - 1); // 첫 주기 끝의 수준
    start = m;
  } else {
    l = y[0];
    b = n > 1 ? y[1] - y[0] : 0;
    start = 1;
  }
  const seasonal = m > 1 && n >= 2 * m;
  let sse = 0;
  const fitted = keep ? new Array(n).fill(null) : null;
  const errs = keep ? [] : null;
  for (let t = start; t < n; t++) {
    const si = seasonal ? t % m : 0;
    const f = l + b + (seasonal ? s[si] : 0);
    const e = y[t] - f;
    sse += e * e;
    if (keep) { fitted[t] = f; errs.push(e); }
    const nl = l + b + alpha * e;
    b += beta * e;
    if (seasonal) s[si] += gamma * e;
    l = nl;
  }
  return { sse, count: n - start, l, b, s, seasonal, fitted, errs, start };
}

/** Nelder–Mead (작은 차원): f 를 [0,1]^d 에서 최소화 */
function minimize(f, x0) {
  const d = x0.length;
  const clamp = (x) => x.map((v) => Math.min(0.999, Math.max(0.001, v)));
  let pts = [clamp(x0)];
  for (let i = 0; i < d; i++) { const p = [...x0]; p[i] = p[i] + (p[i] < 0.5 ? 0.2 : -0.2); pts.push(clamp(p)); }
  let vals = pts.map(f);
  for (let it = 0; it < 120 * d; it++) {
    const ord = vals.map((v, i) => i).sort((a, b) => vals[a] - vals[b]);
    pts = ord.map((i) => pts[i]); vals = ord.map((i) => vals[i]);
    if (Math.abs(vals[d] - vals[0]) <= 1e-10 * (Math.abs(vals[0]) + 1e-10)) break;
    const c = new Array(d).fill(0);
    for (let i = 0; i < d; i++) for (let j = 0; j < d; j++) c[j] += pts[i][j] / d;
    const at = (k) => clamp(c.map((v, j) => v + k * (pts[d][j] - v)));
    const xr = at(-1); const fr = f(xr);
    if (fr < vals[0]) {
      const xe = at(-2); const fe = f(xe);
      if (fe < fr) { pts[d] = xe; vals[d] = fe; } else { pts[d] = xr; vals[d] = fr; }
    } else if (fr < vals[d - 1]) { pts[d] = xr; vals[d] = fr; }
    else {
      const xc = at(0.5); const fc = f(xc);
      if (fc < vals[d]) { pts[d] = xc; vals[d] = fc; }
      else {
        for (let i = 1; i <= d; i++) { pts[i] = clamp(pts[i].map((v, j) => pts[0][j] + 0.5 * (v - pts[0][j]))); vals[i] = f(pts[i]); }
      }
    }
  }
  let bi = 0;
  for (let i = 1; i <= d; i++) if (vals[i] < vals[bi]) bi = i;
  return pts[bi];
}

const memo = new Map();

/**
 * 모형 맞추기. seasonality: 1 자동, 0 없음, 2 이상 = 주기 길이.
 * 반환 { alpha, beta, gamma, m, l, b, s, sigma, n, t0, step, y, fitted, stats } | { error }
 */
export function fitEts(values, timeline, { seasonality = 1, completion = 1, aggregation = 1 } = {}) {
  const key = `${seasonality}|${completion}|${aggregation}|${values.join(',')}|${timeline.join(',')}`;
  if (memo.has(key)) return memo.get(key);
  const res = fitUncached(values, timeline, { seasonality, completion, aggregation });
  if (memo.size > 16) memo.delete(memo.keys().next().value);
  memo.set(key, res);
  return res;
}

function fitUncached(values, timeline, opts) {
  const ser = prepareSeries(values, timeline, opts);
  if (ser.error) return ser;
  const { y } = ser;
  let m = opts.seasonality === 1 ? detectSeason(y) : opts.seasonality === 0 ? 0 : Math.round(opts.seasonality);
  if (m < 0 || m > 8760) return { error: 'NUM' };
  if (m > 1 && y.length < 2 * m) m = 0;
  const seasonal = m > 1;
  const obj = (p) => run(y, m, p[0], p[1], seasonal ? p[2] : 0).sse;
  // 몇 군데에서 시작해 가장 좋은 해
  let best = null;
  let bestV = Infinity;
  for (const a of [0.2, 0.5, 0.8]) {
    const x0 = seasonal ? [a, 0.05, 0.1] : [a, 0.05];
    const p = minimize(obj, x0);
    const v = obj(p);
    if (v < bestV) { bestV = v; best = p; }
  }
  const [alpha, beta, gamma = 0] = best;
  const r = run(y, m, alpha, beta, gamma, true);
  const sigma = Math.sqrt(r.sse / Math.max(1, r.count));
  // 통계: MASE · SMAPE · MAE · RMSE
  let mae = 0; let smape = 0; let naive = 0;
  r.errs.forEach((e, i) => {
    const t = r.start + i;
    const f = r.fitted[t];
    mae += Math.abs(e);
    const den = Math.abs(y[t]) + Math.abs(f);
    smape += den ? (2 * Math.abs(e)) / den : 0;
  });
  for (let t = 1; t < y.length; t++) naive += Math.abs(y[t] - y[t - 1]);
  const k = Math.max(1, r.errs.length);
  mae /= k;
  const stats = {
    alpha, beta, gamma: seasonal ? gamma : 0, mase: naive ? mae / (naive / (y.length - 1)) : 0,
    smape: smape / k, mae, rmse: sigma, step: ser.axis.step,
  };
  return { alpha, beta, gamma, m: seasonal ? m : 0, l: r.l, b: r.b, s: r.s, sigma, n: y.length, axis: ser.axis, y, fitted: r.fitted, stats };
}

/** h 단계 뒤 예측 (h 는 실수 가능: 앞뒤 정수 단계를 직선 보간) */
export function etsForecast(model, target) {
  const h = model.axis.pos(target) - (model.n - 1);
  const at = (k) => {
    if (k <= 0) {
      // 과거 시점: 맞춘 값 (없으면 실제 값)
      const i = model.n - 1 + k;
      if (i < 0) return NaN;
      return model.fitted[i] ?? model.y[i];
    }
    const t = model.n - 1 + k;
    return model.l + k * model.b + (model.m ? model.s[t % model.m] : 0);
  };
  const lo = Math.floor(h);
  if (lo === h) return at(h);
  return at(lo) + (at(lo + 1) - at(lo)) * (h - lo);
}

/** 예측 구간의 반폭 (z = 신뢰 수준의 정규 분위수) */
export function etsConfint(model, target, z) {
  const h = Math.max(1, Math.ceil(model.axis.pos(target) - (model.n - 1) - 1e-9));
  let v = 1;
  for (let j = 1; j < h; j++) {
    const c = model.alpha + model.beta * j + (model.m && j % model.m === 0 ? model.gamma : 0);
    v += c * c;
  }
  return z * model.sigma * Math.sqrt(v);
}
