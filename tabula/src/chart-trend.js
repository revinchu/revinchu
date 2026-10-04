// Render-only trend samples. Keep workbook values/options unchanged.
const finite = v => typeof v === 'number' && Number.isFinite(v);

export function categoryTrendPoints(series) {
  if (!['linear', 'exp', 'movingAvg'].includes(series.trend)) return [];
  const pts = [];
  series.values.forEach((v, i) => { if (finite(v)) pts.push([i, v]); });
  if (pts.length < 2) return [];
  if (series.trend === 'movingAvg') {
    const period = finite(series.trendPeriod) ? Math.floor(series.trendPeriod) : 3;
    const k = Math.max(2, Math.min(pts.length, period)), out = [];
    for (let j = k - 1; j < pts.length; j++) {
      let mean = 0;
      for (let q = j - k + 1; q <= j; q++) mean += pts[q][1] / k;
      if (finite(mean)) out.push([pts[j][0], mean]);
    }
    return out;
  }
  const expo = series.trend === 'exp';
  if (expo && pts.some(p => p[1] <= 0)) return [];
  const ys = pts.map(p => expo ? Math.log(p[1]) : p[1]);
  const mx = pts.reduce((sum, p) => sum + p[0] / pts.length, 0);
  const my = ys.reduce((sum, y) => sum + y / pts.length, 0);
  let sxy = 0, sxx = 0;
  pts.forEach((p, i) => { sxy += (p[0] - mx) * (ys[i] - my); sxx += (p[0] - mx) ** 2; });
  const slope = sxx ? sxy / sxx : 0;
  const forward = finite(series.trendForward) ? Math.max(0, series.trendForward) : 0;
  const start = pts[0][0], end = pts.at(-1)[0] + forward;
  const steps = expo ? 64 : 1, out = [];
  for (let q = 0; q <= steps; q++) {
    const x = start + (end - start) * (q / steps);
    const estimate = my + slope * (x - mx), y = expo ? Math.exp(estimate) : estimate;
    if (!finite(x) || !finite(y)) break;
    // Regression round-off must not promote an exact tick boundary to the next one.
    out.push([x, Number(y.toPrecision(15))]);
  }
  return out.length > 1 ? out : [];
}
