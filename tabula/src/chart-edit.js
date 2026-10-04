// 차트 요소 편집: 입력 모델을 변이하지 않고, 원본 셀·참조·값을 보존하는 패치를 생성한다.
const limit = (n, low, high) => Math.max(low, Math.min(high, n));
export function chartSeriesPatch(chart, index, patch) {
  if (!Number.isInteger(index) || index < 0 || index > 10000) return {};
  const list = (chart.seriesFmt ?? []).map(format => ({ ...format }));
  while (list.length <= index) list.push({});
  list[index] = { ...list[index], ...patch };
  for (const key of Object.keys(list[index])) if (list[index][key] === undefined) delete list[index][key];
  return { seriesFmt: list };
}
// 색 선택은 가져온 명시 색보다 우선한다. 값·축·레이블·배치는 변경하지 않는다.
const SERIES_COLOR_KEYS = ['color', 'grad', 'colors', 'pointColors', 'hierarchyColors', 'markerColor', 'trendColor', 'outline'];
function clearSeriesColors(format) {
  const out = { ...format };
  for (const key of SERIES_COLOR_KEYS) delete out[key];
  return out;
}
function chartColorSources(chart, transform) {
  const patch = {};
  if (Array.isArray(chart.series)) patch.series = chart.series.map((s, i) => transform(s, i));
  if (chart.snapshotData?.series) patch.snapshotData = { ...chart.snapshotData, series: chart.snapshotData.series.map((s, i) => transform(s, s._fi ?? i)) };
  return patch;
}
export function chartPalettePatch(chart, palette) {
  const patch = { ...chartColorSources(chart, clearSeriesColors), palette: palette === 'office' ? undefined : palette };
  if (chart.seriesFmt) patch.seriesFmt = chart.seriesFmt.map(clearSeriesColors);
  for (const key of ['upColor', 'downColor', 'totalColor', 'mapLowColor', 'mapMidColor', 'mapHighColor', 'otherColor']) patch[key] = undefined;
  return patch;
}
export function chartSeriesColorPatch(chart, index, color) {
  if (!Number.isInteger(index) || index < 0 || index > 10000) return {};
  const recolor = (s, i) => i === index ? { ...clearSeriesColors(s), ...(color ? { color } : {}) } : s;
  const seriesFmt = (chart.seriesFmt ?? []).map(recolor);
  while (seriesFmt.length <= index) seriesFmt.push({});
  seriesFmt[index] = recolor(seriesFmt[index], index);
  return { ...chartColorSources(chart, recolor), seriesFmt };
}
export function chartStylePatch(chart, preset, index) {
  // 스타일을 바꿀 때 앞서 고른 어두운 배경·제목 색이 다음 스타일에 남지 않는다.
  const patch = Object.fromEntries(['fill', 'plotFill', 'border', 'gridY', 'textColor', 'gridColor', 'titleColor', 'titleBold', 'rounded', 'labels', 'chartAreaFormat', 'plotAreaFormat'].map(k => [k, undefined]));
  if (Object.hasOwn(preset, 'palette')) Object.assign(patch, chartPalettePatch(chart, preset.palette));
  return { ...patch, ...preset, chartStyle: index };
}
export function chartPointColorPatch(chart, series, point, color) {
  if (!Number.isInteger(series) || series < 0 || series > 10000 || !Number.isInteger(point) || point < 0) return {};
  const resetPoint = format => {
    const out = { ...format };
    for (const key of ['pointColors', 'colors']) {
      if (!out[key]) continue;
      const values = Array.isArray(out[key]) ? [...out[key]] : { ...out[key] };
      delete values[point];
      if (Object.keys(values).length) out[key] = values; else delete out[key];
    }
    return out;
  };
  const fmt = resetPoint(chart.seriesFmt?.[series]);
  if (color) fmt.pointColors = { ...fmt.pointColors, [point]: color };
  const seriesFmt = (chart.seriesFmt ?? []).map(f => ({ ...f }));
  while (seriesFmt.length <= series) seriesFmt.push({});
  seriesFmt[series] = fmt;
  return { ...chartColorSources(chart, (s, i) => i === series ? resetPoint(s) : s), seriesFmt };
}
export function chartExplosionPatch(chart, part, value) {
  const n = Number(value); if (!Number.isFinite(n)) return {};
  const explode = limit(n, 0, 400);
  if (part.kind !== 'point') return chartSeriesPatch(chart, part.s, { explode });
  return chartSeriesPatch(chart, part.s, { pointExplosion: { ...(chart.seriesFmt?.[part.s]?.pointExplosion ?? {}), [part.p]: explode } });
}
export function chartPartDeletePatch(chart, part) {
  if (!part) return null;
  if (['axis-x', 'axis-y', 'axis-y2'].includes(part.kind)) { const key = part.kind.slice(5); return { axes: { ...chart.axes, [key]: { ...chart.axes?.[key], hide: true } } }; }
  if (part.kind === 'title') return { title: '', titleLayout: undefined };
  if (part.kind === 'legend') return { legend: 'none', legendLayout: undefined };
  if (part.kind === 'series' || part.kind === 'point') return { hiddenSeries: [...new Set([...(chart.hiddenSeries ?? []), part.s])] };
  if (part.kind === 'label') return chartSeriesPatch(chart, part.s, { labels: false, catName: false, serName: false, pct: false });
  if (part.kind === 'dataTable') return { dataTable: false };
  return null;
}
export function chartLayoutAfterDrag(layout, delta, width, height, bounds = {}) {
  const w = Math.max(1, width), h = Math.max(1, height);
  const bw = limit(Number(bounds.w ?? layout.w ?? 0), 0, 1), bh = limit(Number(bounds.h ?? layout.h ?? 0), 0, 1);
  return { ...layout, x: limit(layout.x + delta[0] / w, 0, 1 - bw), y: limit(layout.y + delta[1] / h, 0, 1 - bh) };
}
export function chartExplosionAfterDrag(current, delta, geometry) {
  const angle = Number(geometry.angle), radius = Math.max(1, Number(geometry.r)), squash = Math.max(.05, Number(geometry.squash) || 1);
  const movement = (delta[0] * Math.cos(angle) + delta[1] / squash * Math.sin(angle)) / radius * 100;
  return Number.isFinite(movement) ? limit(current + movement, 0, 400) : current;
}
