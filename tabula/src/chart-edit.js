// 차트 요소 편집: 원본 셀·계열 모델을 변경하지 않는 패치만 생성한다.
const limit = (n, low, high) => Math.max(low, Math.min(high, n));
export function chartSeriesPatch(chart, index, patch) {
  if (!Number.isInteger(index) || index < 0 || index > 10000) return {};
  const list = (chart.seriesFmt ?? []).map(format => ({ ...format }));
  while (list.length <= index) list.push({});
  list[index] = { ...list[index], ...patch };
  for (const key of Object.keys(list[index])) if (list[index][key] === undefined) delete list[index][key];
  return { seriesFmt: list };
}
export function chartPointColorPatch(chart, series, point, color) {
  if (!Number.isInteger(point) || point < 0) return {};
  const colors = { ...(chart.seriesFmt?.[series]?.pointColors ?? {}) };
  if (color) colors[point] = color; else delete colors[point];
  return chartSeriesPatch(chart, series, { pointColors: Object.keys(colors).length ? colors : undefined });
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
