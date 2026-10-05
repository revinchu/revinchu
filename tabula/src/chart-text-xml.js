import { child, descendants, esc } from './xml.js';

const TEXT_KEYS = ['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike'];
const prefixKey = (prefix, key) => prefix ? prefix + key[0].toUpperCase() + key.slice(1) : key;

/** Explicit properties only: omitted OOXML properties inherit from their parent. */
export function chartTextFields(object, prefix = '', root = false) {
  const out = {};
  for (const key of TEXT_KEYS) {
    const value = object?.[root && key === 'color' ? 'textColor' : prefixKey(prefix, key)];
    if (value !== undefined && value !== null) out[key] = value;
  }
  return out;
}

export function assignChartTextFields(object, style, prefix = '', root = false) {
  for (const key of TEXT_KEYS) if (style?.[key] !== undefined) object[root && key === 'color' ? 'textColor' : prefixKey(prefix, key)] = style[key];
}

/** Merge paragraph defaults and the first actual run, preserving explicit false. */
export function readChartTextFont(node, colorOf) {
  const def = descendants(node, 'defRPr')[0], run = descendants(node, 'rPr')[0];
  const out = {};
  for (const props of [def, run]) {
    if (!props) continue;
    const a = props.attrs;
    if (Number.isFinite(Number(a.sz)) && Number(a.sz) > 0) out.size = Number(a.sz) / 100;
    for (const [attr, key] of [['b', 'bold'], ['i', 'italic']]) if (a[attr] !== undefined) out[key] = a[attr] === '1' || a[attr] === 'true';
    if (a.u !== undefined) out.underline = a.u !== 'none';
    if (a.strike !== undefined) out.strike = a.strike !== 'noStrike';
    const color = colorOf(child(props, 'solidFill')); if (color) out.color = color;
    const face = child(props, 'ea')?.attrs.typeface || child(props, 'latin')?.attrs.typeface || child(props, 'cs')?.attrs.typeface;
    // Theme tokens are not installed font names. Preserve the normal inherited face.
    if (face && !/^\+(?:mj|mn)-/.test(face)) out.font = face;
  }
  return out;
}

export function chartTextRunXml(style = {}, tag = 'a:defRPr') {
  let attrs = '';
  if (Number.isFinite(style.size)) attrs += ` sz="${Math.round(Math.max(1, Math.min(409, style.size)) * 100)}"`;
  for (const [key, attr] of [['bold', 'b'], ['italic', 'i']]) if (style[key] !== undefined) attrs += ` ${attr}="${style[key] ? 1 : 0}"`;
  if (style.underline !== undefined) attrs += ` u="${style.underline ? 'sng' : 'none'}"`;
  if (style.strike !== undefined) attrs += ` strike="${style.strike ? 'sngStrike' : 'noStrike'}"`;
  const raw = String(style.color ?? '').replace(/^#/, '');
  const hex = /^[\da-f]{3}$/i.test(raw) ? [...raw].map(c => c + c).join('') : /^[\da-f]{6}$/i.test(raw) ? raw : null;
  const color = hex ? `<a:solidFill><a:srgbClr val="${hex.toUpperCase()}"/></a:solidFill>` : '';
  const font = style.font ? `<a:latin typeface="${esc(style.font)}"/><a:ea typeface="${esc(style.font)}"/><a:cs typeface="${esc(style.font)}"/>` : '';
  return color || font ? `<${tag}${attrs}>${color}${font}</${tag}>` : `<${tag}${attrs}/>`;
}

export function chartTextPropertiesXml(style, prefix = 'c', rotation = null) {
  const rotated = Number.isFinite(rotation) && Math.abs(rotation) <= 90;
  if (!Object.keys(style ?? {}).length && !rotated) return '';
  return `<${prefix}:txPr><a:bodyPr${rotated ? ` rot="${Math.round(rotation * 60000)}"` : ''}/><a:lstStyle/><a:p><a:pPr>${chartTextRunXml(style)}</a:pPr><a:endParaRPr lang="ko-KR"/></a:p></${prefix}:txPr>`;
}

export function chartRichTextXml(text, style = {}) {
  return `<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr>${chartTextRunXml(style)}</a:pPr><a:r>${chartTextRunXml(style, 'a:rPr')}<a:t>${esc(text)}</a:t></a:r></a:p></c:rich></c:tx>`;
}

/** Keep formatting for temporarily hidden chart elements without storing data/references. */
export function chartTextSupplement(chart, series = chart.seriesFmt ?? []) {
  const out = {}, radial = ['pie', 'doughnut', 'pieOfPie', 'barOfPie'].includes(chart.type), hierarchy = ['sunburst', 'treemap'].includes(chart.type);
  const noAxes = radial || hierarchy || ['map', 'funnel'].includes(chart.type);
  const secondary = chart.type === 'pareto' || series.some(sr => sr?.axis === 1 || sr?.axis === 'secondary');
  const legendVisible = (chart.legend ?? (series.length > 1 || radial || hierarchy ? 'b' : 'none')) !== 'none';
  for (const prefix of ['title', 'legend']) {
    if (prefix === 'title' ? !!chart.title : legendVisible) continue;
    const style = chartTextFields(chart, prefix);
    if (Object.keys(style).length) out[prefix] = style;
  }
  const modern = ['sunburst', 'treemap', 'waterfall', 'funnel', 'histogram', 'pareto', 'boxWhisker', 'map'].includes(chart.type);
  if (chart.dataTableText && (!chart.dataTable || noAxes || modern)) out.dataTable = chartTextFields(chart.dataTableText);
  const axes = {};
  for (const axis of ['x', 'y', 'y2']) if (chart.axes?.[axis]) {
    const present = !noAxes && (axis !== 'y2' || secondary);
    const style = present ? {} : chartTextFields(chart.axes[axis]);
    const title = present && chart.axes[axis].title ? {} : chartTextFields(chart.axes[axis], 'title');
    if (Object.keys(style).length || Object.keys(title).length) axes[axis] = { style, title };
  }
  if (Object.keys(axes).length) out.axes = axes;
  const formats = series.map(sr => {
    const native = !['surface'].includes(chart.type) && (radial || hierarchy || (sr?.labels ?? chart.labels) !== undefined || sr?.pct || sr?.catName !== undefined || sr?.serName !== undefined);
    if (native) return null;
    const style = chartTextFields(sr, 'label'), points = {};
    for (const [key, value] of Object.entries(sr?.pointLabelStyles ?? {})) if (/^\d+$/.test(key) && Number.isSafeInteger(Number(key))) {
      const point = chartTextFields(value); if (Object.keys(point).length) points[key] = point;
    }
    return Object.keys(style).length || Object.keys(points).length ? { style, points } : null;
  });
  if (formats.some(Boolean)) out.series = formats;
  return Object.keys(out).length ? out : null;
}

/** Standard native formatting wins; supplements only fill absent text properties. */
export function applyChartTextSupplement(chart, source) {
  if (!source || typeof source !== 'object') return;
  const safe = style => {
    const out = {};
    for (const key of TEXT_KEYS) {
      const value = style?.[key];
      if (key === 'font' && typeof value === 'string' && value.length <= 256) out[key] = value;
      else if (key === 'color' && typeof value === 'string' && /^#?[\da-f]{3}(?:[\da-f]{3})?$/i.test(value)) out[key] = value;
      else if (key === 'size' && Number.isFinite(value) && value >= 1 && value <= 409) out[key] = value;
      else if (['bold', 'italic', 'underline', 'strike'].includes(key) && typeof value === 'boolean') out[key] = value;
    }
    return out;
  };
  const merge = (object, style, prefix = '', root = false) => {
    const native = chartTextFields(object, prefix, root);
    assignChartTextFields(object, { ...safe(style), ...native }, prefix, root);
  };
  for (const prefix of ['', 'title', 'legend']) merge(chart, source[prefix || 'chart'], prefix, !prefix);
  if (source.dataTable) { chart.dataTableText ??= {}; merge(chart.dataTableText, source.dataTable); }
  for (const axis of ['x', 'y', 'y2']) if (source.axes?.[axis]) {
    chart.axes ??= {}; chart.axes[axis] ??= {};
    merge(chart.axes[axis], source.axes[axis].style); merge(chart.axes[axis], source.axes[axis].title, 'title');
  }
  if (Array.isArray(source.series)) for (let i = 0; i < Math.min(10000, source.series.length); i++) {
    const record = source.series[i]; if (!record) continue;
    chart.seriesFmt ??= []; chart.seriesFmt[i] ??= {};
    merge(chart.seriesFmt[i], record.style, 'label');
    for (const [key, style] of Object.entries(record.points ?? {}).slice(0, 10000)) if (/^\d+$/.test(key) && Number.isSafeInteger(Number(key))) {
      chart.seriesFmt[i].pointLabelStyles ??= {}; chart.seriesFmt[i].pointLabelStyles[key] ??= {};
      merge(chart.seriesFmt[i].pointLabelStyles[key], style);
    }
  }
}
