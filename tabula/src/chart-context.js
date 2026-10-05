const TEXT_KEYS = ['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike'];
const textKey = (prefix, key) => prefix ? prefix + key[0].toUpperCase() + key.slice(1) : key;
const primitiveFormat = value => typeof value === 'string' || typeof value === 'boolean' || typeof value === 'number' && Number.isFinite(value);
function textFormat(value, prefix = '') {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return {};
  return Object.fromEntries(TEXT_KEYS.map(key => textKey(prefix, key)).filter(key => primitiveFormat(value[key])).map(key => [key, value[key]]));
}
function resetText(value, prefix = '', defaults = {}) {
  const next = { ...value };
  for (const key of TEXT_KEYS.map(key => textKey(prefix, key))) {
    if (defaults[key] === undefined) delete next[key]; else next[key] = defaults[key];
  }
  return next;
}
function resetSeriesText(value) {
  const next = resetText(value, 'label'); delete next.pointLabelStyles; return next;
}

function resetSourceText(chart, part = null) {
  const source = chart.series;
  if (!Array.isArray(source)) return {};
  const affected = (series, index) => (!part || index === part.s) && (series?.pointLabelStyles !== undefined || TEXT_KEYS.some(key => series?.[textKey('label', key)] !== undefined));
  if (!source.some(affected)) return {};
  return { series: source.map((series, index) => {
    if (!affected(series, index)) return series;
    if (part?.kind === 'label' && Number.isInteger(part.p)) { const styles = { ...series.pointLabelStyles }; delete styles[part.p]; return { ...series, pointLabelStyles: Object.keys(styles).length ? styles : undefined }; }
    return resetSeriesText(series);
  }) };
}

// 차트 서식만 초기화한다. 원본 범위·값·종류·축의 척도·요소 표시 여부는 보존한다.
export function chartResetFormattingPatch(chart, part = { kind: 'chart' }, style = {}) {
  const kind = part?.kind ?? 'chart';
  if (kind === 'plot') return { plotAreaFormat: undefined, plotFill: style.plotFill };
  if (kind === 'title' || kind === 'legend') return Object.fromEntries([...TEXT_KEYS.map(key => textKey(kind, key)), kind + 'Layout'].map(key => [key, style[key]]));
  const axis = /^axis-(title-)?(x|y|y2)$/.exec(kind);
  if (axis) return { axes: { ...chart.axes, [axis[2]]: resetText(chart.axes?.[axis[2]], axis[1] ? 'title' : '', style.axes?.[axis[2]]) } };
  if (kind === 'dataTable') return { dataTableText: style.dataTableText };
  if (kind === 'series' || kind === 'point' || kind === 'label') {
    const formats = (chart.seriesFmt ?? []).map(x => ({ ...x }));
    if (!formats[part.s]) return kind === 'point' ? {} : resetSourceText(chart, part);
    if (kind === 'point') {
      for (const key of ['pointColors', 'pointExplosion']) { const values = { ...formats[part.s][key] }; delete values[part.p]; formats[part.s][key] = Object.keys(values).length ? values : undefined; }
    } else if (kind === 'label' && Number.isInteger(part.p)) {
      const styles = { ...formats[part.s].pointLabelStyles }; delete styles[part.p];
      formats[part.s].pointLabelStyles = Object.keys(styles).length ? styles : undefined;
    } else {
      formats[part.s] = resetSeriesText(formats[part.s]);
      if (kind === 'series') for (const key of ['color', 'outline', 'lineWidth', 'dash', 'markerSize', 'markerColor', 'pointColors', 'grad', 'shadow']) delete formats[part.s][key];
    }
    return { seriesFmt: formats, ...(kind === 'point' ? {} : resetSourceText(chart, part)) };
  }
  const keys = ['chartAreaFormat', 'plotAreaFormat', 'fill', 'plotFill', 'border', 'textColor', 'gridColor', 'axisSize', 'rounded', 'dataTableText', ...TEXT_KEYS.filter(key => key !== 'color'), ...['title', 'legend'].flatMap(prefix => TEXT_KEYS.map(key => textKey(prefix, key)))];
  const patch = Object.fromEntries(keys.map(key => [key, style[key]]));
  if (chart.axes) patch.axes = Object.fromEntries(Object.entries(chart.axes).map(([key, value]) => [key, resetText(resetText(value, '', style.axes?.[key]), 'title', style.axes?.[key])]));
  if (chart.seriesFmt) patch.seriesFmt = chart.seriesFmt.map(resetSeriesText);
  return { ...patch, ...resetSourceText(chart) };
}

// 템플릿에 문서·시트 이름, 원본 참조나 사용자 값이 섞이지 않도록 허용 목록만 복사한다.
export const CHART_TEMPLATE_KEYS = ['type', 'threeD', 'barShape', 'view3D', 'grouping', 'stacked', 'percent', 'legend', 'legendColor', 'legendSize', 'legendBold', 'legendFont', 'legendItalic', 'legendUnderline', 'legendStrike', 'legendLayout', 'titleSize', 'titleColor', 'titleBold', 'titleFont', 'titleItalic', 'titleUnderline', 'titleStrike', 'titleLayout', 'font', 'size', 'bold', 'italic', 'underline', 'strike', 'axes', 'dataTableText', 'axisSize', 'gridX', 'gridY', 'gridColor', 'textColor', 'fill', 'plotFill', 'border', 'rounded', 'palette', 'chartAreaFormat', 'plotAreaFormat', 'chartStyle', 'seriesFmt', 'labels', 'hole', 'explode', 'firstAngle', 'gap', 'overlap', 'marker', 'scatterStyle', 'radarStyle', 'bubbleScale', 'bubble3D', 'surfaceStyle', 'bandCount', 'ohlc', 'volume', 'splitType', 'splitPos', 'secondSize', 'splitGap', 'comboAxis', 'comboLayout', 'showMean', 'connectors', 'upColor', 'downColor', 'totalColor', 'mapLowColor', 'mapMidColor', 'mapHighColor'];
// resolveChart는 계열에 seriesFmt를 펼쳐 넣는다. 따라서 값·참조·내부 인덱스는
// 서식 파일에서 절대 전달하지 않고, 실제 렌더러/UI/입출력의 서식 속성만 복사한다.
const SERIES_FORMAT_KEYS = ['type', 'axis', 'grouping', 'barShape', 'color', 'outline', 'lineWidth', 'dash', 'marker', 'markerSize', 'markerColor', 'smooth', 'labels', 'catName', 'serName', 'pct', 'numFmt', 'percentFmt', 'labelPos', 'labelSeparator', 'labelColor', 'labelSize', 'labelBold', 'labelFont', 'labelItalic', 'labelUnderline', 'labelStrike', 'explode', 'trend', 'trendColor', 'trendPeriod', 'trendForward'];
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
  if (value.pointLabelStyles && typeof value.pointLabelStyles === 'object' && !Array.isArray(value.pointLabelStyles)) out.pointLabelStyles = Object.fromEntries(Object.entries(value.pointLabelStyles).filter(([key]) => point(key)).map(([key, style]) => [key, textFormat(style)]));
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
  const textFields = new Set([...TEXT_KEYS, ...['title', 'legend'].flatMap(prefix => TEXT_KEYS.map(key => textKey(prefix, key)))]);
  const out = Object.fromEntries(CHART_TEMPLATE_KEYS.filter(k => !['seriesFmt', 'axes', 'dataTableText'].includes(k) && chart[k] !== undefined && (!textFields.has(k) || primitiveFormat(chart[k]))).map(k => [k, structuredClone(chart[k])]));
  if (chart.axes && typeof chart.axes === 'object' && !Array.isArray(chart.axes)) out.axes = Object.fromEntries(['x', 'y', 'y2'].filter(key => chart.axes[key] && typeof chart.axes[key] === 'object' && !Array.isArray(chart.axes[key])).map(key => [key, { ...textFormat(chart.axes[key]), ...textFormat(chart.axes[key], 'title') }]));
  if (chart.dataTableText && typeof chart.dataTableText === 'object' && !Array.isArray(chart.dataTableText)) out.dataTableText = textFormat(chart.dataTableText);
  if (Array.isArray(chart.seriesFmt)) out.seriesFmt = chart.seriesFmt.map(templateSeriesFormat);
  return out;
}
export function applyChartTemplatePatch(current, template) {
  const source = chartTemplateFormat(template), patch = Object.fromEntries(CHART_TEMPLATE_KEYS.map(k => [k, source[k]]));
  // 축 제목·숫자 척도·필터/배치 값은 현재 차트의 것을 유지하고 글꼴 서식만 교체한다.
  if (current.axes || source.axes) {
    patch.axes = { ...current.axes };
    for (const key of ['x', 'y', 'y2']) if (current.axes?.[key] || source.axes?.[key]) patch.axes[key] = { ...resetText(resetText(current.axes?.[key]), 'title'), ...source.axes?.[key] };
  }
  // 서식 파일에 있는 계열 번호를 현재 계열 참조에 덮어쓰지 않는다.
  if (current.series?.length && source.seriesFmt?.length) patch.seriesFmt = current.series.map((_, i) => structuredClone(source.seriesFmt[i % source.seriesFmt.length]));
  return { ...patch, ...resetSourceText(current) };
}
