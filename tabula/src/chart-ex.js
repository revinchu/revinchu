// Office 2016+ ChartEx adapter. MS-ODRAWXML §2.24 / §5.22.
// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/e2723b0a-9120-42a5-bd11-c252ccb13c1e
// Chart geometry stays in chart.js; this module only handles standard chart XML.
import { child, kids, descendants, allText, esc } from './xml.js';

export const CHARTEX_NS = 'http://schemas.microsoft.com/office/drawing/2014/chartex';
export const CHARTEX_REL = 'http://schemas.microsoft.com/office/2014/relationships/chartEx';
export const CHARTEX_CONTENT = 'application/vnd.ms-office.chartex+xml';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const TYPES = { waterfall: 'waterfall', funnel: 'funnel', histogram: 'clusteredColumn', pareto: 'clusteredColumn', treemap: 'treemap', sunburst: 'sunburst', boxWhisker: 'boxWhisker', map: 'regionMap' };
const OWN_KEYS = ['type', 'byRows', 'axes', 'seriesFmt', 'labels', 'dataTable', 'gap', 'legend', 'palette', 'hiddenSeries', 'hiddenCats', 'titleSize', 'titleColor', 'titleBold', 'legendSize', 'legendColor', 'legendBold', 'axisSize', 'textColor', 'gridColor', 'rounded', 'gridX', 'gridY', 'fill', 'plotFill', 'border', 'totals', 'binCount', 'binWidth', 'upColor', 'downColor', 'totalColor', 'showMean', 'connectors', 'quartileMethod', 'showOutliers', 'showInnerPoints', 'mapLowColor', 'mapMidColor', 'mapHighColor'];
const num = (v) => typeof v === 'number' && Number.isFinite(v);
const on = (v) => v === '1' || v === 'true';
const hex = (v) => /^#?[\da-f]{6}$/i.test(String(v ?? '')) ? String(v).replace('#', '').toUpperCase() : null;
const fill = (v) => hex(v) ? `<a:solidFill><a:srgbClr val="${hex(v)}"/></a:solidFill>` : '';
const shape = (v, border) => v || border ? `<cx:spPr>${fill(v)}${border ? `<a:ln w="9525">${fill(border)}</a:ln>` : ''}</cx:spPr>` : '';
const f = (value, dir) => value ? `<cx:f${dir ? ` dir="${dir}"` : ''}>${esc(value)}</cx:f>` : '';
const refDirection = (r) => { const m = /\$?[A-Z]+\$?(\d+):\$?[A-Z]+\$?(\d+)$/i.exec(r ?? ''); return m && m[1] === m[2] ? 'row' : 'col'; };
const tx = (value, ref) => `<cx:tx><cx:txData>${f(ref)}<cx:v>${esc(value ?? '')}</cx:v></cx:txData></cx:tx>`;
const textPr = (size, color, bold) => `<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${num(size) ? ` sz="${Math.round(Math.max(1, Math.min(400, size)) * 100)}"` : ''}${bold !== undefined ? ` b="${bold ? 1 : 0}"` : ''}>${fill(color)}</a:defRPr></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></cx:txPr>`;
const lvl = (values, numeric, code, name) => `<cx:lvl ptCount="${values.length}"${code ? ` formatCode="${esc(code)}"` : ''}${name ? ` name="${esc(name)}"` : ''}>${values.map((v, i) => (numeric ? num(v) : v !== null && v !== undefined && v !== '') ? `<cx:pt idx="${i}">${esc(v)}</cx:pt>` : '').join('')}</cx:lvl>`;

export function isChartEx(chart) { return Object.hasOwn(TYPES, chart?.type); }

// cx:lvl is ordered leaf → root (Microsoft Open-XML-SDK SunburstChartExample).
function categoryLevels(data) {
  return [data.categories ?? [], ...(data.catLevels ?? []).map((spans) => {
    const values = Array(data.categories.length).fill('');
    for (const s of spans) for (let i = Math.max(0, s.start); i <= Math.min(values.length - 1, s.end); i++) values[i] = s.text;
    return values;
  })];
}

/** chartModelData + source refs [{tx,cat,val}] → native ChartEx XML. */
export function writeChartEx(chart, data, refs = [], palette = ['#4472c4', '#ed7d31', '#a5a5a5']) {
  if (!isChartEx(chart)) throw new Error('지원하지 않는 확장 차트입니다.');
  const type = chart.type, hierarchy = type === 'sunburst' || type === 'treemap';
  const cartesian = ['waterfall', 'histogram', 'pareto', 'boxWhisker'].includes(type);
  const source = data.series.length ? data.series : [{ name: '', values: [] }];
  const levels = categoryLevels(data), catAuto = (data.categories ?? []).every((v, i) => String(v) === String(i + 1));
  const parts = [], series = [];
  const layout = (sr) => {
    let p = '';
    if (type === 'treemap') p += '<cx:parentLabelLayout val="banner"/>';
    if (type === 'waterfall') p += `<cx:visibility connectorLines="${chart.connectors === false ? 0 : 1}"/>`;
    if (type === 'boxWhisker') p += `<cx:visibility meanMarker="${chart.showMean ? 1 : 0}" outliers="${chart.showOutliers === false ? 0 : 1}" nonoutliers="${chart.showInnerPoints ? 1 : 0}"/>`;
    if (type === 'histogram' || type === 'pareto') {
      if (type === 'pareto' && !catAuto) p += '<cx:aggregation/>';
      else p += `<cx:binning intervalClosed="r">${num(chart.binWidth) && chart.binWidth > 0 ? `<cx:binSize>${chart.binWidth}</cx:binSize>` : num(chart.binCount) && chart.binCount > 0 ? `<cx:binCount>${Math.round(chart.binCount)}</cx:binCount>` : ''}</cx:binning>`;
    }
    if (type === 'map') p += '<cx:geography projectionType="mercator" viewedRegionType="world" cultureLanguage="ko-KR" cultureRegion="KR" attribution=""/>';
    if (type === 'boxWhisker') p += `<cx:statistics quartileMethod="${chart.quartileMethod === 'exclusive' ? 'exclusive' : 'inclusive'}"/>`;
    if (type === 'waterfall') {
      const totals = chart.totals ?? (data.categories ?? []).map((v, i) => /^(합계|총계|총합계|소계|전체|total|subtotal|grand total|net)$/i.test(String(v).trim()) ? i : -1).filter((i) => i >= 0);
      p += `<cx:subtotals>${totals.filter((i) => Number.isInteger(i) && i >= 0 && i < sr.values.length).map((i) => `<cx:idx val="${i}"/>`).join('')}</cx:subtotals>`;
    }
    return p ? `<cx:layoutPr>${p}</cx:layoutPr>` : '';
  };
  source.forEach((sr, i) => {
    const r = refs[i] ?? {}, fi = sr._fi ?? i, color = sr.color ?? palette[i % palette.length];
    // Size drives area charts; geographic regions use colorVal. Other layouts use val.
    const dim = hierarchy ? 'size' : type === 'map' ? 'colorVal' : 'val';
    const cats = type === 'histogram' || type === 'pareto' && catAuto ? '' : `<cx:strDim type="cat">${f(r.cat, refDirection(r.val))}${levels.map((v) => lvl(v, false)).join('')}</cx:strDim>`;
    parts.push(`<cx:data id="${i}">${cats}<cx:numDim type="${dim}">${f(r.val, refDirection(r.val))}${lvl(sr.values, true, sr.numFmt ?? 'General', sr.name)}</cx:numDim></cx:data>`);
    const labels = sr.labels ?? chart.labels;
    const dataLabels = labels !== undefined || hierarchy ? `<cx:dataLabels pos="${hierarchy ? 'ctr' : 'bestFit'}"><cx:visibility seriesName="0" categoryName="${hierarchy && labels !== false ? 1 : 0}" value="${labels ? 1 : 0}"/></cx:dataLabels>` : '';
    const ptColors = sr.pointColors ?? sr.colors ?? {};
    const points = Object.entries(ptColors).filter(([k, v]) => /^\d+$/.test(k) && Number(k) < sr.values.length && hex(v)).map(([k, v]) => `<cx:dataPt idx="${k}">${shape(v)}</cx:dataPt>`).join('');
    const mapColors = type === 'map' ? ['min', 'mid', 'max'].map((stop, k) => { const h = hex(chart[['mapLowColor', 'mapMidColor', 'mapHighColor'][k]]); return h ? `<cx:${stop}Color><a:srgbClr val="${h}"/></cx:${stop}Color>` : ''; }).join('') : '';
    series.push(`<cx:series layoutId="${TYPES[type]}" formatIdx="${fi}"${(chart.hiddenSeries ?? []).includes(fi) ? ' hidden="1"' : ''}>${tx(sr.name, r.tx)}${shape(color)}${mapColors ? `<cx:valueColors>${mapColors}</cx:valueColors>` : ''}${points}${dataLabels}<cx:dataId val="${i}"/>${layout(sr)}${cartesian ? '<cx:axisId>0</cx:axisId><cx:axisId>1</cx:axisId>' : ''}</cx:series>`);
  });
  if (type === 'pareto') series.push('<cx:series layoutId="paretoLine" ownerIdx="0" formatIdx="1"><cx:axisId>0</cx:axisId><cx:axisId>2</cx:axisId></cx:series>');
  const axis = (id, key, category = false) => {
    const a = chart.axes?.[key] ?? {};
    const scaling = category ? `<cx:catScaling${num(chart.gap) ? ` gapWidth="${Math.max(0, chart.gap / 100)}"` : ''}/>` : `<cx:valScaling${['min', 'max', 'major'].filter((k) => num(a[k]) && (k !== 'major' || a[k] > 0)).map((k) => ` ${k === 'major' ? 'majorUnit' : k}="${a[k]}"`).join('')}/>`;
    return `<cx:axis id="${id}" hidden="${a.hide ? 1 : 0}">${scaling}${a.title ? `<cx:title>${tx(a.title)}</cx:title>` : ''}${!category && chart.gridY !== false ? '<cx:majorGridlines/>' : ''}<cx:tickLabels/>${a.numFmt ? `<cx:numFmt formatCode="${esc(a.numFmt)}" sourceLinked="0"/>` : ''}${textPr(chart.axisSize, chart.textColor)}</cx:axis>`;
  };
  const axes = cartesian ? axis(0, 'x', true) + axis(1, 'y') + (type === 'pareto' ? '<cx:axis id="2"><cx:valScaling min="0" max="1"/><cx:tickLabels/><cx:numFmt formatCode="0%" sourceLinked="0"/></cx:axis>' : '') : '';
  const title = chart.title ? `<cx:title pos="t" align="ctr" overlay="0">${tx(chart.title)}${textPr(chart.titleSize ?? 14, chart.titleColor, chart.titleBold ?? false)}</cx:title>` : '';
  const lp = chart.legend ?? (source.length > 1 || hierarchy ? 'b' : 'none');
  const legend = ['l', 't', 'r', 'b'].includes(lp) ? `<cx:legend pos="${lp}" align="ctr" overlay="0">${textPr(chart.legendSize, chart.legendColor, chart.legendBold)}</cx:legend>` : '';
  const props = Object.fromEntries(OWN_KEYS.filter((k) => chart[k] !== undefined).map((k) => [k, chart[k]]));
  if (chart.pivot) props.wxPivot = chart.pivot;
  const ext = `<cx:extLst><cx:ext uri="{5E2A6C7B-8F4D-4B1A-9C3E-7D6F1A2B3C4D}"><tb:props xmlns:tb="urn:tabula:chart" json="${esc(JSON.stringify(props))}"/></cx:ext></cx:extLst>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cx:chartSpace xmlns:cx="${CHARTEX_NS}" xmlns:a="${A}"><cx:chartData>${parts.join('')}</cx:chartData><cx:chart>${title}<cx:plotArea><cx:plotAreaRegion>${series.join('')}</cx:plotAreaRegion>${axes}${shape(chart.plotFill)}</cx:plotArea>${legend}</cx:chart>${shape(chart.fill, chart.border)}${ext}</cx:chartSpace>`;
}

const cached = (level, numeric = false) => {
  const count = Math.min(1000000, Math.max(0, Number(level?.attrs.ptCount) || 0));
  const values = Array(count).fill(numeric ? null : '');
  for (const p of kids(level, 'pt')) {
    const i = Number(p.attrs.idx), value = numeric ? Number(p.text) : p.text;
    if (Number.isInteger(i) && i >= 0 && i < count && (!numeric || Number.isFinite(value))) values[i] = value;
  }
  return values;
};
const font = (n, colorOf) => {
  const p = descendants(n, 'defRPr')[0] ?? descendants(n, 'rPr')[0];
  const out = {};
  if (p?.attrs.sz) out.size = Number(p.attrs.sz) / 100;
  if (p?.attrs.b !== undefined) out.bold = on(p.attrs.b);
  const color = colorOf(child(p, 'solidFill')); if (color) out.color = color;
  return out;
};
const text = (n) => child(child(n, 'txData'), 'v')?.text ?? allText(child(n, 'rich'));
// Extensions are optional untrusted JSON; copy only the chart option allowlist,
// ordinary values and bounded nesting. Never restore paths/URLs or object keys.
function safeOption(value, depth = 0) {
  if (depth > 12) return undefined;
  if (value === null || typeof value === 'boolean' || typeof value === 'string') return value;
  if (typeof value === 'number') return Number.isFinite(value) ? value : undefined;
  if (Array.isArray(value)) return value.slice(0, 10000).map((v) => safeOption(v, depth + 1));
  if (value && typeof value === 'object') return Object.fromEntries(Object.entries(value).filter(([k]) => !['__proto__', 'prototype', 'constructor'].includes(k)).slice(0, 10000).map(([k, v]) => [k, safeOption(v, depth + 1)]));
  return undefined;
}

/** A namespace-qualified ChartEx chartSpace → WIXEL model; no extension is needed for type/data. */
export function readChartEx(root, refOf = () => null, colorOf = (n) => { const h = child(n, 'srgbClr')?.attrs.val; return h ? `#${h}` : null; }) {
  if (!Object.values(root.attrs ?? {}).includes(CHARTEX_NS) || !child(root, 'chartData')) return null;
  const ch = child(root, 'chart'), plot = child(ch, 'plotArea'), region = child(plot, 'plotAreaRegion');
  const all = kids(region, 'series'), main = all.filter((s) => s.attrs.layoutId !== 'paretoLine');
  if (!main.length) return null;
  const layout = main[0].attrs.layoutId;
  const type = layout === 'clusteredColumn' ? (all.some((s) => s.attrs.layoutId === 'paretoLine') ? 'pareto' : 'histogram') : layout === 'regionMap' ? 'map' : layout;
  if (!isChartEx({ type })) return null;
  const out = { type, series: [], seriesFmt: [], axes: {}, legend: child(ch, 'legend')?.attrs.pos ?? 'none' };
  const data = new Map(kids(child(root, 'chartData'), 'data').map((d) => [d.attrs.id, d]));
  const hidden = [];
  main.forEach((s, i) => {
    const d = data.get(child(s, 'dataId')?.attrs.val), dims = kids(d, 'numDim');
    const val = dims.find((n) => ['val', 'size', 'colorVal'].includes(n.attrs.type)) ?? dims[0];
    const cat = kids(d, 'strDim').find((n) => n.attrs.type === 'cat');
    const levels = kids(cat, 'lvl').map((n) => cached(n)), values = cached(child(val, 'lvl'), true);
    const nameNode = child(child(s, 'tx'), 'txData');
    const sr = { name: { text: child(nameNode, 'v')?.text ?? allText(child(child(s, 'tx'), 'rich')) ?? `계열${i + 1}` }, cache: values };
    for (const [key, node] of [['val', val], ['cat', cat]]) { const ref = refOf(child(node, 'f')?.text); if (ref) sr[key] = ref; }
    const nr = refOf(child(nameNode, 'f')?.text); if (nr) sr.name.ref = nr;
    if (child(val, 'f')?.attrs.dir === 'row') out.byRows = true;
    if (levels.length) sr.catCache = levels[0];
    if (levels.length > 1) sr.catLevels = levels.slice(1).map((vs, depth) => vs.reduce((spans, v, j) => { const last = spans.at(-1), sameParent = levels.slice(depth + 2).every((level) => level[j] === level[j - 1]); if (last && last.text === v && sameParent) last.end = j; else spans.push({ text: v, start: j, end: j }); return spans; }, []));
    out.series.push(sr);
    const sf = {}, color = colorOf(child(child(s, 'spPr'), 'solidFill'));
    if (color) sf.color = color;
    const code = child(val, 'lvl')?.attrs.formatCode; if (code) sf.numFmt = code;
    const ptColors = {}; for (const pt of kids(s, 'dataPt')) { const c = colorOf(child(child(pt, 'spPr'), 'solidFill')); if (c) ptColors[pt.attrs.idx] = c; }
    if (Object.keys(ptColors).length) sf.pointColors = ptColors;
    const labels = child(s, 'dataLabels'), visible = child(labels, 'visibility');
    if (visible && (!['sunburst', 'treemap'].includes(type) || on(visible.attrs.value) || !on(visible.attrs.categoryName))) sf.labels = on(visible.attrs.value);
    out.seriesFmt[i] = sf;
    if (on(s.attrs.hidden)) hidden.push(i);
    if (i !== 0) return;
    const p = child(s, 'layoutPr'), bin = child(p, 'binning'), vis = child(p, 'visibility');
    for (const [stop, key] of [['minColor', 'mapLowColor'], ['midColor', 'mapMidColor'], ['maxColor', 'mapHighColor']]) { const c = colorOf(child(child(s, 'valueColors'), stop)); if (c) out[key] = c; }
    const width = Number(child(bin, 'binSize')?.text), count = Number(child(bin, 'binCount')?.text);
    if (width > 0) out.binWidth = width; if (count > 0) out.binCount = count;
    if (vis?.attrs.connectorLines !== undefined) out.connectors = on(vis.attrs.connectorLines);
    if (vis?.attrs.meanMarker !== undefined) out.showMean = on(vis.attrs.meanMarker);
    if (vis?.attrs.outliers !== undefined) out.showOutliers = on(vis.attrs.outliers);
    if (vis?.attrs.nonoutliers !== undefined) out.showInnerPoints = on(vis.attrs.nonoutliers);
    const stats = child(p, 'statistics'); if (stats?.attrs.quartileMethod) out.quartileMethod = stats.attrs.quartileMethod;
    const totals = child(p, 'subtotals'); if (totals) out.totals = kids(totals, 'idx').map((n) => Number(n.attrs.val)).filter((n) => Number.isInteger(n) && n >= 0);
  });
  if (hidden.length) out.hiddenSeries = hidden;
  for (const [i, a] of kids(plot, 'axis').entries()) {
    const cat = child(a, 'catScaling'), v = child(a, 'valScaling'), key = cat ? 'x' : i > 1 ? 'y2' : 'y', props = {};
    if (on(a.attrs.hidden)) props.hide = true;
    for (const [attr, k] of [['min', 'min'], ['max', 'max'], ['majorUnit', 'major']]) if (v?.attrs[attr] !== undefined && Number.isFinite(Number(v.attrs[attr]))) props[k] = Number(v.attrs[attr]);
    const title = text(child(child(a, 'title'), 'tx')); if (title) props.title = title;
    const code = child(a, 'numFmt')?.attrs.formatCode; if (code) props.numFmt = code;
    out.axes[key] = props;
    if (cat && Number.isFinite(Number(cat.attrs.gapWidth))) out.gap = Number(cat.attrs.gapWidth) * 100;
  }
  const title = child(ch, 'title'); out.title = text(child(title, 'tx'));
  for (const [node, prefix] of [[title, 'title'], [child(ch, 'legend'), 'legend']]) {
    const ff = font(child(node, 'txPr') ?? node, colorOf);
    for (const [k, v] of Object.entries(ff)) out[`${prefix}${k[0].toUpperCase()}${k.slice(1)}`] = v;
  }
  const bg = colorOf(child(child(root, 'spPr'), 'solidFill')); if (bg) out.fill = bg;
  const border = colorOf(child(child(child(root, 'spPr'), 'ln'), 'solidFill')); if (border) out.border = border;
  const pg = colorOf(child(child(plot, 'spPr'), 'solidFill')); if (pg) out.plotFill = pg;
  const own = descendants(child(root, 'extLst'), 'props').find((p) => p.attrs['xmlns:tb'] === 'urn:tabula:chart');
  if (own?.attrs.json) { try {
    const p = JSON.parse(own.attrs.json);
    if (p && typeof p === 'object' && !Array.isArray(p)) {
      for (const k of OWN_KEYS) if (k !== 'type' && Object.hasOwn(p, k)) {
        const v = safeOption(p[k]);
        if (['axes', 'seriesFmt'].includes(k) && (!v || typeof v !== 'object')) continue;
        if (['labels', 'dataTable', 'byRows', 'rounded', 'gridX', 'gridY', 'showMean', 'connectors', 'showOutliers', 'showInnerPoints'].includes(k) && typeof v !== 'boolean') continue;
        if (['hiddenSeries', 'hiddenCats', 'totals'].includes(k) && (!Array.isArray(v) || v.some((n) => !Number.isInteger(n) || n < 0))) continue;
        if (v !== undefined) out[k] = v;
      }
      if (p.wxPivot && typeof p.wxPivot === 'object' && typeof p.wxPivot.name === 'string') out.pivot = safeOption(p.wxPivot);
    }
  } catch { /* unknown extension does not affect native data */ } }
  return out;
}
