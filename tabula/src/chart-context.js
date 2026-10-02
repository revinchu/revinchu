// 차트 서식만 초기화한다. 원본 범위·값·종류·축의 척도·요소 표시 여부는 보존한다.
export function chartResetFormattingPatch(chart, part = { kind: 'chart' }, style = {}) {
  const kind = part?.kind ?? 'chart';
  if (kind === 'plot') return { plotAreaFormat: undefined, plotFill: style.plotFill };
  if (kind === 'title' || kind === 'legend') return Object.fromEntries(['Color', 'Size', 'Bold', 'Layout'].map(k => [kind + k, style[kind + k]]));
  if (kind === 'series' || kind === 'point' || kind === 'label') {
    const formats = (chart.seriesFmt ?? []).map(x => ({ ...x }));
    if (!formats[part.s]) return {};
    if (kind === 'point') {
      for (const key of ['pointColors', 'pointExplosion']) { const values = { ...formats[part.s][key] }; delete values[part.p]; formats[part.s][key] = Object.keys(values).length ? values : undefined; }
    } else {
      for (const key of ['color', 'outline', 'lineWidth', 'dash', 'markerSize', 'markerColor', 'pointColors', 'labelColor', 'labelSize', 'labelBold', 'grad', 'shadow']) delete formats[part.s][key];
    }
    return { seriesFmt: formats };
  }
  const keys = ['chartAreaFormat', 'plotAreaFormat', 'fill', 'plotFill', 'border', 'textColor', 'gridColor', 'titleColor', 'titleSize', 'titleBold', 'legendColor', 'legendSize', 'legendBold', 'rounded'];
  return Object.fromEntries(keys.map(key => [key, style[key]]));
}

// 템플릿에 문서·시트 이름, 원본 참조나 사용자 값이 섞이지 않도록 허용 목록만 복사한다.
export const CHART_TEMPLATE_KEYS = ['type', 'threeD', 'barShape', 'view3D', 'grouping', 'stacked', 'percent', 'legend', 'legendColor', 'legendSize', 'legendBold', 'legendLayout', 'titleSize', 'titleColor', 'titleBold', 'titleLayout', 'axisSize', 'gridX', 'gridY', 'gridColor', 'textColor', 'fill', 'plotFill', 'border', 'rounded', 'palette', 'chartAreaFormat', 'plotAreaFormat', 'chartStyle', 'seriesFmt', 'labels', 'hole', 'explode', 'firstAngle', 'gap', 'overlap', 'marker', 'scatterStyle', 'radarStyle', 'bubbleScale', 'bubble3D', 'surfaceStyle', 'bandCount', 'ohlc', 'volume', 'splitType', 'splitPos', 'secondSize', 'splitGap', 'comboAxis', 'comboLayout', 'showMean', 'connectors', 'upColor', 'downColor', 'totalColor', 'mapLowColor', 'mapMidColor', 'mapHighColor'];
// resolveChart는 계열에 seriesFmt를 펼쳐 넣는다. 따라서 값·참조·내부 인덱스는
// 서식 파일에서 절대 전달하지 않고, 실제 렌더러/UI/입출력의 서식 속성만 복사한다.
const SERIES_FORMAT_KEYS = ['type', 'axis', 'grouping', 'barShape', 'color', 'outline', 'lineWidth', 'dash', 'marker', 'markerSize', 'markerColor', 'smooth', 'labels', 'catName', 'serName', 'pct', 'numFmt', 'percentFmt', 'labelPos', 'labelSeparator', 'labelColor', 'labelSize', 'labelBold', 'explode', 'trend', 'trendColor', 'trendPeriod', 'trendForward'];
function templateSeriesFormat(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  const primitive = v => typeof v === 'string' || typeof v === 'boolean' || typeof v === 'number' && Number.isFinite(v);
  const out = Object.fromEntries(SERIES_FORMAT_KEYS.filter(k => primitive(value[k])).map(k => [k, value[k]]));
  const color = v => typeof v === 'string';
  const point = k => /^(0|[1-9]\d*)$/.test(k) && Number(k) < 1000000;
  for (const key of ['pointColors', 'pointExplosion', 'hierarchyColors']) {
    const map = value[key];
    if (!map || typeof map !== 'object' || Array.isArray(map)) continue;
    out[key] = Object.fromEntries(Object.entries(map).filter(([k, v]) =>
      (key === 'hierarchyColors' || point(k)) && (key === 'pointExplosion' ? typeof v === 'number' && Number.isFinite(v) : color(v))));
  }
  if (Array.isArray(value.colors)) out.colors = value.colors.map(v => color(v) ? v : null);
  if (value.grad && Array.isArray(value.grad.stops)) {
    const stops = value.grad.stops.filter(s => Array.isArray(s) && Number.isFinite(s[0]) && color(s[1])).map(s => [s[0], s[1]]);
    if (stops.length > 1) out.grad = { stops, ...(Number.isFinite(value.grad.ang) ? { ang: value.grad.ang } : {}) };
  }
  // 현재 계열 그림자는 고정된 효과의 켜기/끄기이다. 임의 객체 내용은 보존하지 않는다.
  if (value.shadow !== undefined) out.shadow = !!value.shadow;
  return out;
}
export function chartTemplateFormat(chart) {
  const out = Object.fromEntries(CHART_TEMPLATE_KEYS.filter(k => k !== 'seriesFmt' && chart[k] !== undefined).map(k => [k, structuredClone(chart[k])]));
  if (Array.isArray(chart.seriesFmt)) out.seriesFmt = chart.seriesFmt.map(templateSeriesFormat);
  return out;
}
export function applyChartTemplatePatch(current, template) {
  const source = chartTemplateFormat(template), patch = Object.fromEntries(CHART_TEMPLATE_KEYS.map(k => [k, source[k]]));
  // 서식 파일에 있는 계열 번호를 현재 계열 참조에 덮어쓰지 않는다.
  if (current.series?.length && source.seriesFmt?.length) patch.seriesFmt = current.series.map((_, i) => structuredClone(source.seriesFmt[i % source.seriesFmt.length]));
  return patch;
}
