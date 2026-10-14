// 숨긴 행/열 안에서 접힌 차트는 원래 셀 앵커를 보존해 숨기기 해제와 XLSX 저장에 사용한다.
export function chartAnchorCurrent(chart) {
  const a = chart.collapsedAnchor;
  return a && ['x', 'y', 'w', 'h'].every(k => chart[k] === a.box[k]) ? a : null;
}

export function chartAnchorBox(chart, cols, rows) {
  const anchor = chartAnchorCurrent(chart);
  if (!anchor) return chart;
  const point = p => ({ x: cols.pos(p.col) + Math.min(p.colOff / 9525, cols.size(p.col)), y: rows.pos(p.row) + Math.min(p.rowOff / 9525, rows.size(p.row)) });
  const a = point(anchor.from), b = point(anchor.to);
  return { ...chart, x: Math.round(a.x), y: Math.round(a.y), w: Math.max(0, Math.round(b.x - a.x)), h: Math.max(0, Math.round(b.y - a.y)) };
}

export function chartAnchorPointXml(point) {
  return ['col', 'colOff', 'row', 'rowOff'].map(k => `<xdr:${k}>${Math.max(0, Math.round(Number(point[k]) || 0))}</xdr:${k}>`).join('');
}
