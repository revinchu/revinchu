// 데이터 분석 도구 (엑셀 분석 도구 모음 · Analysis ToolPak 과 같은 출력) — DOM 없음
// 각 도구는 입력 값(2차원 배열)과 옵션을 받아 출력 표 { rows, heads, pct } 를 돌려줌.
// 셀: 숫자 · 문자열 · null · { f: '=수식' } · { err: '#N/A' }. heads = 제목 줄(굵게 · 아래 테두리), pct = 백분율 칸 'r,c'
import { tCdf, fCdf, invert, normCdf, normInv } from './fx-stat.js';
import { maxOf, minOf } from './fxcore.js';

// ───────────── 도우미 ─────────────
const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const sum = (a) => a.reduce((s, x) => s + x, 0);
const mean = (a) => sum(a) / a.length;
const devsq = (a) => { const m = mean(a); return a.reduce((s, x) => s + (x - m) ** 2, 0); };
const varS = (a) => (a.length > 1 ? devsq(a) / (a.length - 1) : NaN);
const tInv2 = (p, df) => invert((x) => tCdf(x, df), 1 - p / 2, 0, 10); // 양측 기각치
const tInv1 = (p, df) => invert((x) => tCdf(x, df), 1 - p, 0, 10); // 단측 기각치
const fInvRt = (p, d1, d2) => invert((x) => fCdf(x, d1, d2), 1 - p, 0, 10);
const tTwoTail = (t, df) => 2 * (1 - tCdf(Math.abs(t), df));
const tOneTail = (t, df) => 1 - tCdf(Math.abs(t), df);
const NA = { err: '#N/A' };
const fin = (v) => (Number.isFinite(v) ? v : NA);

/** 입력 범위 → 변수 목록. byRows: 행 단위, labels: 첫 행/열이 이름표 */
export function splitGroups(rows, { byRows = false, labels = false } = {}) {
  const h = rows.length;
  const w = h ? maxOf(rows.map((r) => r.length)) : 0;
  const at = (r, c) => rows[r]?.[c] ?? null;
  const out = [];
  const count = byRows ? h : w;
  for (let k = 0; k < count; k++) {
    const cells = [];
    const len = byRows ? w : h;
    for (let i = labels ? 1 : 0; i < len; i++) cells.push(byRows ? at(k, i) : at(i, k));
    const lab = labels ? at(byRows ? k : 0, byRows ? 0 : k) : null;
    out.push({
      label: lab === null || lab === '' ? `${byRows ? '행' : '열'} ${k + 1}` : String(lab),
      raw: cells,
      vals: cells.filter(isNum),
    });
  }
  return out;
}

/** 표본 첨도 · 왜도 (엑셀 KURT · SKEW) */
function kurt(a) {
  const n = a.length;
  if (n < 4) return NaN;
  const m = mean(a);
  const s = Math.sqrt(varS(a));
  if (!s) return NaN;
  const q = a.reduce((acc, x) => acc + ((x - m) / s) ** 4, 0);
  return ((n * (n + 1)) / ((n - 1) * (n - 2) * (n - 3))) * q - (3 * (n - 1) ** 2) / ((n - 2) * (n - 3));
}
function skew(a) {
  const n = a.length;
  if (n < 3) return NaN;
  const m = mean(a);
  const s = Math.sqrt(varS(a));
  if (!s) return NaN;
  return (n / ((n - 1) * (n - 2))) * a.reduce((acc, x) => acc + ((x - m) / s) ** 3, 0);
}
function modeOf(a) {
  const cnt = new Map();
  let best = null;
  let bc = 1;
  for (const x of a) {
    const c = (cnt.get(x) ?? 0) + 1;
    cnt.set(x, c);
    if (c > bc) { bc = c; best = x; }
  }
  return best === null ? NA : best;
}
function median(a) {
  const s = [...a].sort((x, y) => x - y);
  const n = s.length;
  return n % 2 ? s[(n - 1) / 2] : (s[n / 2 - 1] + s[n / 2]) / 2;
}

/** 작은 연립 방정식: (A)x = b — 가우스 소거 (특이하면 null) */
function inv(A) {
  const n = A.length;
  const M = A.map((row, i) => [...row, ...Array.from({ length: n }, (_, j) => (i === j ? 1 : 0))]);
  for (let c = 0; c < n; c++) {
    let p = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[p][c])) p = r;
    if (Math.abs(M[p][c]) < 1e-12) return null;
    [M[c], M[p]] = [M[p], M[c]];
    const d = M[c][c];
    for (let j = 0; j < 2 * n; j++) M[c][j] /= d;
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const f = M[r][c];
      if (f) for (let j = 0; j < 2 * n; j++) M[r][j] -= f * M[c][j];
    }
  }
  return M.map((row) => row.slice(n));
}

export class AnalysisError extends Error {}
const need = (ok, msg) => { if (!ok) throw new AnalysisError(msg); };

// ───────────── 도구 ─────────────
/** 기술 통계법 */
export function descriptive(groups, { summary = true, conf = 0.95, kth = 0, kthSmall = 0 } = {}) {
  need(groups.length && groups.every((g) => g.vals.length), '입력 범위에 숫자가 없습니다.');
  const rows = [];
  const heads = [];
  const blocks = groups.map((g) => {
    const a = g.vals;
    const n = a.length;
    const sd = Math.sqrt(varS(a));
    const lines = [];
    if (summary) {
      lines.push(['평균', mean(a)], ['표준 오차', fin(sd / Math.sqrt(n))], ['중앙값', median(a)], ['최빈값', modeOf(a)],
        ['표준 편차', fin(sd)], ['분산', fin(varS(a))], ['첨도', fin(kurt(a))], ['왜도', fin(skew(a))],
        ['범위', maxOf(a) - minOf(a)], ['최소값', minOf(a)], ['최대값', maxOf(a)], ['합', sum(a)], ['관측수', n]);
    }
    const sorted = [...a].sort((x, y) => x - y);
    if (kth > 0) lines.push([`최대값(${kth})`, kth <= n ? sorted[n - kth] : NA]);
    if (kthSmall > 0) lines.push([`최소값(${kthSmall})`, kthSmall <= n ? sorted[kthSmall - 1] : NA]);
    if (conf) lines.push([`신뢰 수준(${(conf * 100).toFixed(1)}%)`, n > 1 ? fin(tInv2(1 - conf, n - 1) * (sd / Math.sqrt(n))) : NA]);
    return { label: g.label, lines };
  });
  // 엑셀처럼 변수마다 두 칸씩 옆으로
  const height = blocks[0].lines.length;
  heads.push(0);
  rows.push(blocks.flatMap((b) => [b.label, null]));
  rows.push(blocks.flatMap(() => [null, null]));
  for (let i = 0; i < height; i++) rows.push(blocks.flatMap((b) => b.lines[i]));
  return { rows, heads };
}

/** 상관 분석 / 공분산 분석 (공분산은 엑셀처럼 모집단 기준) */
export function matrixTool(groups, kind = 'correl') {
  need(groups.length >= 2, '변수가 두 개 이상 있어야 합니다.');
  const k = groups.length;
  const pair = (a, b) => {
    const xs = [];
    const ys = [];
    const n = Math.min(a.raw.length, b.raw.length);
    for (let i = 0; i < n; i++) if (isNum(a.raw[i]) && isNum(b.raw[i])) { xs.push(a.raw[i]); ys.push(b.raw[i]); }
    return [xs, ys];
  };
  const rows = [[null, ...groups.map((g) => g.label)]];
  for (let i = 0; i < k; i++) {
    const row = [groups[i].label];
    for (let j = 0; j < k; j++) {
      if (j > i) { row.push(null); continue; }
      const [xs, ys] = pair(groups[i], groups[j]);
      const n = xs.length;
      const mx = mean(xs);
      const my = mean(ys);
      let sxy = 0; let sxx = 0; let syy = 0;
      for (let t = 0; t < n; t++) { sxy += (xs[t] - mx) * (ys[t] - my); sxx += (xs[t] - mx) ** 2; syy += (ys[t] - my) ** 2; }
      if (kind === 'correl') row.push(i === j ? 1 : sxx && syy ? sxy / Math.sqrt(sxx * syy) : { err: '#DIV/0!' });
      else row.push(n ? sxy / n : { err: '#DIV/0!' });
    }
    rows.push(row);
  }
  return { rows, heads: [0], firstColBold: true };
}

/** 회귀 분석 (최소 제곱, 여러 X) */
export function regression(yRaw, xCols, { constant = true, conf = 0.95, residuals = false, stdResiduals = false, xLabels = null, yLabel = null } = {}) {
  const n0 = yRaw.length;
  need(xCols.every((c) => c.length === n0), '입력 Y 범위와 입력 X 범위의 행 수가 같아야 합니다.');
  const obs = [];
  for (let i = 0; i < n0; i++) {
    const xs = xCols.map((c) => c[i]);
    need(isNum(yRaw[i]) && xs.every(isNum), '입력 범위에 숫자가 아닌 값이 있습니다.');
    obs.push({ y: yRaw[i], x: xs });
  }
  const n = obs.length;
  const k = xCols.length;
  const p = k + (constant ? 1 : 0);
  need(n > p, '관측수가 변수 수보다 많아야 합니다.');
  const X = obs.map((o) => (constant ? [1, ...o.x] : o.x));
  const Y = obs.map((o) => o.y);
  const XtX = Array.from({ length: p }, (_, i) => Array.from({ length: p }, (_, j) => X.reduce((s, row) => s + row[i] * row[j], 0)));
  const XtY = Array.from({ length: p }, (_, i) => X.reduce((s, row, t) => s + row[i] * Y[t], 0));
  const XtXi = inv(XtX);
  need(XtXi, '입력 X 변수들이 서로 선형 종속입니다.');
  const beta = XtXi.map((row) => row.reduce((s, v, j) => s + v * XtY[j], 0));
  const fit = X.map((row) => row.reduce((s, v, j) => s + v * beta[j], 0));
  const res = Y.map((y, i) => y - fit[i]);
  const sse = res.reduce((s, e) => s + e * e, 0);
  const my = mean(Y);
  const sst = constant ? Y.reduce((s, y) => s + (y - my) ** 2, 0) : Y.reduce((s, y) => s + y * y, 0);
  const ssr = sst - sse;
  const dfR = k;
  const dfE = n - p;
  const r2 = sst ? ssr / sst : 1;
  const adj = 1 - (1 - r2) * ((n - (constant ? 1 : 0)) / dfE);
  const mse = sse / dfE;
  const F = (ssr / dfR) / mse;
  const rows = [];
  const heads = [];
  const pct = Math.round(conf * 100 * 10) / 10;
  rows.push(['요약 출력']);
  rows.push([]);
  heads.push(rows.length); rows.push(['회귀분석 통계량', null]);
  rows.push(['다중 상관계수', Math.sqrt(Math.max(0, r2))], ['결정계수', r2], ['조정된 결정계수', adj], ['표준 오차', Math.sqrt(mse)], ['관측수', n]);
  rows.push([]);
  rows.push(['분산 분석']);
  heads.push(rows.length); rows.push([null, '자유도', '제곱합', '제곱 평균', 'F 비', '유의한 F']);
  rows.push(['회귀', dfR, ssr, ssr / dfR, fin(F), Number.isFinite(F) ? 1 - fCdf(F, dfR, dfE) : NA]);
  rows.push(['잔차', dfE, sse, mse]);
  rows.push(['계', dfR + dfE, sst]);
  rows.push([]);
  const tc = tInv2(1 - conf, dfE);
  const extra = conf !== 0.95;
  heads.push(rows.length);
  rows.push([null, '계수', '표준 오차', 't 통계량', 'P-값', '하위 95%', '상위 95%', ...(extra ? [`하위 ${pct}%`, `상위 ${pct}%`] : [])]);
  const t95 = tInv2(0.05, dfE);
  const names = [...(constant ? ['Y 절편'] : []), ...xCols.map((_, i) => xLabels?.[i] ?? `X ${i + 1}`)];
  beta.forEach((b, j) => {
    const se = Math.sqrt(Math.max(0, mse * XtXi[j][j]));
    const t = se ? b / se : NaN;
    rows.push([names[j], b, se, fin(t), Number.isFinite(t) ? tTwoTail(t, dfE) : NA, b - t95 * se, b + t95 * se, ...(extra ? [b - tc * se, b + tc * se] : [])]);
  });
  if (!constant) rows.splice(rows.length - k, 0, ['Y 절편', 0, NA, NA, NA, NA, NA, ...(extra ? [NA, NA] : [])]);
  if (residuals || stdResiduals) {
    rows.push([], [], ['잔차 출력']);
    rows.push([]);
    heads.push(rows.length);
    rows.push(['관측수', `예측치 ${yLabel ?? 'Y'}`, '잔차', ...(stdResiduals ? ['표준 잔차'] : [])]);
    const sdRes = Math.sqrt(sse / (n - 1));
    res.forEach((e, i) => rows.push([i + 1, fit[i], e, ...(stdResiduals ? [sdRes ? e / sdRes : NA] : [])]));
  }
  return { rows, heads, coef: beta, fit, res };
}

/** 히스토그램: bins 가 없으면 최소~최대를 고르게 나눔 */
export function histogram(vals, bins = null, { pareto = false, cumulative = false } = {}) {
  need(vals.length, '입력 범위에 숫자가 없습니다.');
  let b = bins?.filter(isNum);
  if (!b?.length) {
    const lo = minOf(vals);
    const hi = maxOf(vals);
    const k = Math.max(1, Math.round(Math.sqrt(vals.length)));
    const w = (hi - lo) / k || 1;
    b = Array.from({ length: k + 1 }, (_, i) => Number((lo + w * i).toPrecision(12)));
  }
  b = [...b].sort((x, y) => x - y);
  const freq = new Array(b.length + 1).fill(0);
  for (const v of vals) {
    let i = 0;
    while (i < b.length && v > b[i]) i++;
    freq[i]++;
  }
  const labels = [...b, '기타'];
  const total = vals.length;
  const make = (order) => {
    let acc = 0;
    return order.map((i) => { acc += freq[i]; return [labels[i], freq[i], ...(cumulative ? [acc / total] : [])]; });
  };
  const idx = labels.map((_, i) => i);
  const rows = [['계급', '빈도수', ...(cumulative ? ['누적 %'] : [])], ...make(idx)];
  const pct = new Set();
  if (cumulative) for (let r = 1; r < rows.length; r++) pct.add(`${r},2`);
  if (pareto) {
    const w = rows[0].length;
    const sorted = make([...idx].sort((x, y) => freq[y] - freq[x] || x - y));
    rows[0].push(null, '계급', '빈도수', ...(cumulative ? ['누적 %'] : []));
    sorted.forEach((row, i) => { rows[i + 1].push(null, ...row); if (cumulative) pct.add(`${i + 1},${w + 3}`); });
  }
  return { rows, heads: [0], pct, bins: b, freq };
}

/** 순위와 백분율 */
export function rankPercentile(groups) {
  const rows = [];
  const pct = new Set();
  const heads = [0];
  const blocks = groups.map((g) => {
    const items = g.raw.map((v, i) => ({ v, i })).filter((x) => isNum(x.v));
    const sorted = [...items].sort((a, b) => b.v - a.v || a.i - b.i);
    const n = items.length;
    return sorted.map((x) => {
      const rank = 1 + items.filter((y) => y.v > x.v).length;
      const below = items.filter((y) => y.v < x.v).length;
      return [x.i + 1, x.v, rank, n > 1 ? Math.floor((below / (n - 1)) * 1000) / 1000 : 1];
    });
  });
  rows.push(groups.flatMap((g) => ['요소', g.label, '순위', '백분율']));
  const h = maxOf(blocks.map((b) => b.length));
  for (let i = 0; i < h; i++) {
    rows.push(blocks.flatMap((b) => b[i] ?? [null, null, null, null]));
    blocks.forEach((b, j) => { if (b[i]) pct.add(`${i + 1},${j * 4 + 3}`); });
  }
  return { rows, heads, pct };
}

/** t-검정: kind 'paired' | 'equal' | 'unequal' */
export function tTest(a, b, { kind = 'equal', hyp = 0, alpha = 0.05, labels = ['변수 1', '변수 2'] } = {}) {
  let x = a;
  let y = b;
  if (kind === 'paired') {
    const xs = []; const ys = [];
    for (let i = 0; i < Math.min(a.length, b.length); i++) if (isNum(a[i]) && isNum(b[i])) { xs.push(a[i]); ys.push(b[i]); }
    x = xs; y = ys;
  } else { x = a.filter(isNum); y = b.filter(isNum); }
  need(x.length > 1 && y.length > 1, '각 변수에 숫자가 두 개 이상 있어야 합니다.');
  const n1 = x.length; const n2 = y.length;
  const m1 = mean(x); const m2 = mean(y);
  const v1 = varS(x); const v2 = varS(y);
  const rows = [[null, labels[0], labels[1]], ['평균', m1, m2], ['분산', v1, v2], ['관측수', n1, n2]];
  let t;
  let df;
  if (kind === 'paired') {
    const d = x.map((v, i) => v - y[i]);
    const sd = Math.sqrt(varS(d));
    let sxy = 0;
    for (let i = 0; i < n1; i++) sxy += (x[i] - m1) * (y[i] - m2);
    rows.push(['피어슨 상관 계수', sxy / Math.sqrt(devsq(x) * devsq(y))]);
    df = n1 - 1;
    t = (mean(d) - hyp) / (sd / Math.sqrt(n1));
  } else if (kind === 'equal') {
    df = n1 + n2 - 2;
    const pooled = ((n1 - 1) * v1 + (n2 - 1) * v2) / df;
    rows.push(['공동(Pooled) 분산', pooled]);
    t = (m1 - m2 - hyp) / Math.sqrt(pooled * (1 / n1 + 1 / n2));
  } else {
    const a1 = v1 / n1; const a2 = v2 / n2;
    df = Math.round((a1 + a2) ** 2 / (a1 ** 2 / (n1 - 1) + a2 ** 2 / (n2 - 1)));
    t = (m1 - m2 - hyp) / Math.sqrt(a1 + a2);
  }
  rows.push(['가설 평균차', hyp], ['자유도', df], ['t 통계량', fin(t)],
    ['P(T<=t) 단측 검정', Number.isFinite(t) ? tOneTail(t, df) : NA], ['t 기각치 단측 검정', tInv1(alpha, df)],
    ['P(T<=t) 양측 검정', Number.isFinite(t) ? tTwoTail(t, df) : NA], ['t 기각치 양측 검정', tInv2(alpha, df)]);
  return { rows, heads: [0], t, df };
}

/** z-검정: 평균에 대한 두 집단 (알려진 분산) */
export function zTest(a, b, { var1, var2, hyp = 0, alpha = 0.05, labels = ['변수 1', '변수 2'] } = {}) {
  const x = a.filter(isNum); const y = b.filter(isNum);
  need(x.length && y.length, '각 변수에 숫자가 있어야 합니다.');
  need(var1 > 0 && var2 > 0, '두 변수의 알려진 분산을 입력하세요.');
  const z = (mean(x) - mean(y) - hyp) / Math.sqrt(var1 / x.length + var2 / y.length);
  const rows = [[null, labels[0], labels[1]], ['평균', mean(x), mean(y)], ['알려진 분산', var1, var2], ['관측수', x.length, y.length],
    ['가설 평균차', hyp], ['z 통계량', z], ['P(Z<=z) 단측 검정', 1 - normCdf(Math.abs(z))], ['z 기각치 단측 검정', normInv(1 - alpha)],
    ['P(Z<=z) 양측 검정', 2 * (1 - normCdf(Math.abs(z)))], ['z 기각치 양측 검정', normInv(1 - alpha / 2)]];
  return { rows, heads: [0], z };
}

/** F-검정: 분산에 대한 두 집단 */
export function fTest(a, b, { alpha = 0.05, labels = ['변수 1', '변수 2'] } = {}) {
  const x = a.filter(isNum); const y = b.filter(isNum);
  need(x.length > 1 && y.length > 1, '각 변수에 숫자가 두 개 이상 있어야 합니다.');
  const v1 = varS(x); const v2 = varS(y);
  const d1 = x.length - 1; const d2 = y.length - 1;
  const F = v1 / v2;
  const left = F < 1;
  const p = left ? fCdf(F, d1, d2) : 1 - fCdf(F, d1, d2);
  const crit = left ? invert((q) => fCdf(q, d1, d2), alpha, 0, 10) : fInvRt(alpha, d1, d2);
  const rows = [[null, labels[0], labels[1]], ['평균', mean(x), mean(y)], ['분산', v1, v2], ['관측수', x.length, y.length], ['자유도', d1, d2],
    ['F 비', fin(F)], ['P(F<=f) 단측 검정', fin(p)], ['F 기각치: 단측 검정', crit]];
  return { rows, heads: [0], F };
}

/** 분산 분석: 일원 배치법 */
export function anova1(groups, { alpha = 0.05 } = {}) {
  const gs = groups.filter((g) => g.vals.length);
  need(gs.length >= 2, '두 개 이상의 집단에 숫자가 있어야 합니다.');
  const all = gs.flatMap((g) => g.vals);
  const gm = mean(all);
  const ssb = gs.reduce((s, g) => s + g.vals.length * (mean(g.vals) - gm) ** 2, 0);
  const ssw = gs.reduce((s, g) => s + devsq(g.vals), 0);
  const dfb = gs.length - 1;
  const dfw = all.length - gs.length;
  const F = (ssb / dfb) / (ssw / dfw);
  const rows = [['분산 분석: 일원 배치법'], [], ['요약표']];
  const heads = [rows.length];
  rows.push(['인자의 수준', '관측수', '합', '평균', '분산']);
  for (const g of gs) rows.push([g.label, g.vals.length, sum(g.vals), mean(g.vals), fin(varS(g.vals))]);
  rows.push([], [], ['분산 분석']);
  heads.push(rows.length);
  rows.push(['변동의 요인', '제곱합', '자유도', '제곱 평균', 'F 비', 'P-값', 'F 기각치']);
  rows.push(['처리', ssb, dfb, ssb / dfb, fin(F), Number.isFinite(F) ? 1 - fCdf(F, dfb, dfw) : NA, fInvRt(alpha, dfb, dfw)]);
  rows.push(['잔차', ssw, dfw, ssw / dfw]);
  rows.push([]);
  rows.push(['계', ssb + ssw, all.length - 1]);
  return { rows, heads, F };
}

/** 분산 분석: 반복 없는 이원 배치법 (행 = 인자 A, 열 = 인자 B) */
export function anova2(matrix, { alpha = 0.05, rowLabels, colLabels } = {}) {
  const r = matrix.length;
  const c = matrix[0]?.length ?? 0;
  need(r >= 2 && c >= 2 && matrix.every((row) => row.length === c && row.every(isNum)), '숫자로 채워진 2×2 이상의 표가 필요합니다.');
  const all = matrix.flat();
  const gm = mean(all);
  const rowM = matrix.map(mean);
  const colM = Array.from({ length: c }, (_, j) => mean(matrix.map((row) => row[j])));
  const ssA = c * rowM.reduce((s, m) => s + (m - gm) ** 2, 0);
  const ssB = r * colM.reduce((s, m) => s + (m - gm) ** 2, 0);
  const sst = all.reduce((s, x) => s + (x - gm) ** 2, 0);
  const sse = sst - ssA - ssB;
  const dfA = r - 1; const dfB = c - 1; const dfE = dfA * dfB;
  const mse = sse / dfE;
  const FA = (ssA / dfA) / mse; const FB = (ssB / dfB) / mse;
  const rows = [['분산 분석: 반복 없는 이원 배치법'], []];
  const heads = [rows.length];
  rows.push(['요약표', '관측수', '합', '평균', '분산']);
  matrix.forEach((row, i) => rows.push([rowLabels?.[i] ?? `행 ${i + 1}`, c, sum(row), mean(row), varS(row)]));
  rows.push([]);
  for (let j = 0; j < c; j++) { const col = matrix.map((row) => row[j]); rows.push([colLabels?.[j] ?? `열 ${j + 1}`, r, sum(col), mean(col), varS(col)]); }
  rows.push([], [], ['분산 분석']);
  heads.push(rows.length);
  rows.push(['변동의 요인', '제곱합', '자유도', '제곱 평균', 'F 비', 'P-값', 'F 기각치']);
  rows.push(['인자 A(행)', ssA, dfA, ssA / dfA, fin(FA), Number.isFinite(FA) ? 1 - fCdf(FA, dfA, dfE) : NA, fInvRt(alpha, dfA, dfE)]);
  rows.push(['인자 B(열)', ssB, dfB, ssB / dfB, fin(FB), Number.isFinite(FB) ? 1 - fCdf(FB, dfB, dfE) : NA, fInvRt(alpha, dfB, dfE)]);
  rows.push(['잔차', sse, dfE, mse]);
  rows.push([]);
  rows.push(['계', sst, r * c - 1]);
  return { rows, heads };
}

/**
 * 이동 평균법: 엑셀처럼 수식으로 출력. inRef(i) = i번째 입력 셀 주소, outRef(i) = i번째 출력 셀 주소
 * interval = 구간, stdErr = 표준 오차 열 추가
 */
export function movingAverage(n, inRef, outRef, { interval = 3, stdErr = false } = {}) {
  need(interval >= 2 && interval <= n, `구간은 2 이상 ${n} 이하여야 합니다.`);
  const rows = [];
  for (let i = 0; i < n; i++) {
    const avg = i < interval - 1 ? NA : { f: `=AVERAGE(${inRef(i - interval + 1)}:${inRef(i)})` };
    const row = [avg];
    if (stdErr) {
      row.push(i < 2 * interval - 2 ? NA : { f: `=SQRT(SUMXMY2(${inRef(i - interval + 1)}:${inRef(i)},${outRef(i - interval + 1)}:${outRef(i)})/${interval})` });
    }
    rows.push(row);
  }
  return { rows, heads: [] };
}

/** 지수 평활법: 감쇠 인수 damping (α = 1 - damping) */
export function expSmoothing(n, inRef, outRef, { damping = 0.3, stdErr = false } = {}) {
  need(damping >= 0 && damping <= 1, '감쇠 인수는 0과 1 사이여야 합니다.');
  const a = Number((1 - damping).toPrecision(12));
  const rows = [];
  for (let i = 0; i < n; i++) {
    const v = i === 0 ? NA : i === 1 ? { f: `=${inRef(0)}` } : { f: `=${a}*${inRef(i - 1)}+${damping}*${outRef(i - 1)}` };
    const row = [v];
    if (stdErr) row.push(i < 4 ? NA : { f: `=SQRT(SUMXMY2(${inRef(i - 3)}:${inRef(i - 1)},${outRef(i - 3)}:${outRef(i - 1)})/3)` });
    rows.push(row);
  }
  return { rows, heads: [] };
}

/** 재현 가능한 난수 (mulberry32) */
export function rng(seed) {
  let s = seed ? seed >>> 0 : (Math.random() * 2 ** 32) >>> 0;
  return () => {
    s = (s + 0x6D2B79F5) >>> 0;
    let t = s;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

/** 난수 생성: dist 'uniform' | 'normal' | 'bernoulli' | 'binomial' | 'poisson' */
export function randomNumbers({ vars = 1, count = 10, dist = 'uniform', p1 = 0, p2 = 1, seed = 0 } = {}) {
  need(vars >= 1 && count >= 1 && vars * count <= 1e6, '변수 수와 난수 개수를 확인하세요.');
  const r = rng(seed);
  const gen = {
    uniform: () => p1 + (p2 - p1) * r(),
    normal: () => p1 + p2 * normInv(Math.min(1 - 1e-12, Math.max(1e-12, r()))),
    bernoulli: () => (r() < p1 ? 1 : 0),
    binomial: () => { let k = 0; for (let i = 0; i < p2; i++) if (r() < p1) k++; return k; },
    poisson: () => { const L = Math.exp(-p1); let k = 0; let q = 1; do { k++; q *= r(); } while (q > L); return k - 1; },
  }[dist];
  need(gen, '분포를 고르세요.');
  return { rows: Array.from({ length: count }, () => Array.from({ length: vars }, gen)), heads: [] };
}

/** 표본 추출: 'periodic' (주기 n 마다) | 'random' (n 개, 복원 추출) */
export function sampling(vals, { method = 'random', n = 1, seed = 0 } = {}) {
  need(vals.length, '입력 범위에 숫자가 없습니다.');
  need(n >= 1, '값을 1 이상으로 입력하세요.');
  if (method === 'periodic') {
    const out = [];
    for (let i = n - 1; i < vals.length; i += n) out.push([vals[i]]);
    return { rows: out, heads: [] };
  }
  const r = rng(seed);
  return { rows: Array.from({ length: n }, () => [vals[Math.floor(r() * vals.length)]]), heads: [] };
}

/** 도구 목록 (엑셀 [데이터 분석] 대화상자와 같은 순서) */
export const ANALYSIS_TOOLS = [
  { id: 'anova1', label: '분산 분석: 일원 배치법' },
  { id: 'anova2', label: '분산 분석: 반복 없는 이원 배치법' },
  { id: 'correl', label: '상관 분석' },
  { id: 'covar', label: '공분산 분석' },
  { id: 'descr', label: '기술 통계법' },
  { id: 'expsmooth', label: '지수 평활법' },
  { id: 'ftest', label: 'F-검정: 분산에 대한 두 집단' },
  { id: 'hist', label: '히스토그램' },
  { id: 'movavg', label: '이동 평균법' },
  { id: 'random', label: '난수 생성' },
  { id: 'rank', label: '순위와 백분율' },
  { id: 'regress', label: '회귀 분석' },
  { id: 'sampling', label: '표본 추출' },
  { id: 'ttestPaired', label: 't-검정: 쌍체 비교' },
  { id: 'ttestEq', label: 't-검정: 등분산 가정 두 집단' },
  { id: 'ttestUneq', label: 't-검정: 이분산 가정 두 집단' },
  { id: 'ztest', label: 'z-검정: 평균에 대한 두 집단' },
];

/**
 * 해 찾기용 최소화: objective(x) → { f, pen } (pen = 제약 위반 합). 벌점 가중치를 키워 가며 넬더–미드,
 * 끝에 좌표 탐색으로 다듬고, ints (정수 변수 번호) 는 반올림해 고정한 뒤 나머지를 다시 맞춤.
 * 반환 { x, f, pen }
 */
export function solveMin(objective, x0, { ints = [], maxEval = 20000 } = {}) {
  let evals = 0;
  const cache = new Map();
  const ev = (x) => {
    const k = x.join(',');
    if (cache.has(k)) return cache.get(k);
    evals++;
    const r = objective(x);
    const o = typeof r === 'number' ? { f: r, pen: 0 } : r;
    if (cache.size > 50000) cache.clear();
    cache.set(k, o);
    return o;
  };
  const d = x0.length;
  const fixed = new Map();
  const full = (y) => { const x = []; let j = 0; for (let i = 0; i < d; i++) x.push(fixed.has(i) ? fixed.get(i) : y[j++]); return x; };
  const nm = (F, y0, scale) => {
    const n = y0.length;
    if (!n) return y0;
    let pts = [y0];
    for (let i = 0; i < n; i++) { const p = [...y0]; p[i] += scale[i]; pts.push(p); }
    let vals = pts.map(F);
    for (let it = 0; it < 400 * n && evals < maxEval; it++) {
      const ord = vals.map((v, i) => i).sort((a, b) => vals[a] - vals[b]);
      pts = ord.map((i) => pts[i]); vals = ord.map((i) => vals[i]);
      if (Math.abs(vals[n] - vals[0]) <= 1e-12 * (Math.abs(vals[0]) + 1e-12) && it > 5) break;
      const c = new Array(n).fill(0);
      for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) c[j] += pts[i][j] / n;
      const at = (k) => c.map((v, j) => v + k * (pts[n][j] - v));
      const xr = at(-1); const fr = F(xr);
      if (fr < vals[0]) { const xe = at(-2); const fe = F(xe); if (fe < fr) { pts[n] = xe; vals[n] = fe; } else { pts[n] = xr; vals[n] = fr; } }
      else if (fr < vals[n - 1]) { pts[n] = xr; vals[n] = fr; }
      else {
        const xc = at(0.5); const fc = F(xc);
        if (fc < vals[n]) { pts[n] = xc; vals[n] = fc; }
        else for (let i = 1; i <= n; i++) { pts[i] = pts[i].map((v, j) => pts[0][j] + 0.5 * (v - pts[0][j])); vals[i] = F(pts[i]); }
      }
    }
    let bi = 0;
    for (let i = 1; i <= n; i++) if (vals[i] < vals[bi]) bi = i;
    return pts[bi];
  };
  // 좌표 탐색 (단계 절반씩)
  const polish = (F, y0) => {
    let y = [...y0];
    let fy = F(y);
    let step = y.map((v) => Math.max(1, Math.abs(v)) * 0.1);
    for (let round = 0; round < 60 && evals < maxEval; round++) {
      let moved = false;
      for (let i = 0; i < y.length; i++) {
        for (const sgn of [1, -1]) {
          const t = [...y]; t[i] += sgn * step[i];
          const ft = F(t);
          if (ft < fy) { y = t; fy = ft; moved = true; break; }
        }
      }
      if (!moved) { step = step.map((s) => s / 2); if (step.every((s, i) => s < 1e-10 * Math.max(1, Math.abs(y[i])))) break; }
    }
    return y;
  };
  const solveFree = (y0) => {
    let y = y0;
    for (const mu of [1e2, 1e4, 1e6, 1e8]) {
      const F = (yy) => { const o = ev(full(yy)); return o.f + mu * o.pen; };
      const scale = y.map((v) => Math.max(1, Math.abs(v)) * 0.25);
      y = nm(F, y, scale);
      y = polish(F, y);
    }
    return y;
  };
  let y = solveFree([...x0]);
  if (ints.length) {
    // 정수 변수: 하나씩 반올림 → 고정 → 나머지 다시
    for (const i of ints) {
      const x = full(y);
      const lo = Math.floor(x[i]);
      const cand = [lo, lo + 1].map((v) => { fixed.set(i, v); const rest = x.filter((_, k) => !fixed.has(k)); const yy = solveFree(rest); const o = ev(full(yy)); return { v, yy, s: o.f + 1e8 * o.pen }; });
      const best = cand[0].s <= cand[1].s ? cand[0] : cand[1];
      fixed.set(i, best.v);
      y = best.yy;
    }
  }
  const x = full(y);
  const o = ev(x);
  return { x, f: o.f, pen: o.pen };
}
