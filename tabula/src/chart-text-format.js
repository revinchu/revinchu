import { fontFamilyCandidates } from './fonts.js';

const KEYS = ['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike'];
const BOOLEAN_KEYS = new Set(['bold', 'italic', 'underline', 'strike']);
const prefixed = (prefix, key) => prefix ? prefix + key[0].toUpperCase() + key.slice(1) : key;
const own = (value, key) => value != null && Object.prototype.hasOwnProperty.call(value, key);
const seriesIndex = part => Number.isInteger(part?.s) && part.s >= 0 ? part.s : null;

/** 차트 안의 텍스트만 선택 대상으로 해석한다. 계열의 선/채우기 선택은 포함하지 않는다. */
export function chartTextTargetSupported(chart, part = null) {
  if (!chart || part?.id && chart.id && part.id !== chart.id) return false;
  const kind = part?.kind ?? 'chart';
  return ['chart', 'title', 'legend', 'dataTable'].includes(kind) || /^axis-(?:title-)?(?:x|y|y2)$/.test(kind) || kind === 'label' && seriesIndex(part) !== null;
}

function source(chart, part) {
  const kind = part?.kind ?? 'chart';
  if (kind === 'title' || kind === 'legend') return { object: chart, prefix: kind };
  const axis = /^axis-(title-)?(x|y|y2)$/.exec(kind);
  if (axis) return { object: chart.axes?.[axis[2]] ?? {}, prefix: axis[1] ? 'title' : '', axis: axis[2] };
  if (kind === 'label') return { object: { ...chart.series?.[part.s], ...chart.seriesFmt?.[part.s] }, prefix: 'label', series: part.s };
  if (kind === 'dataTable') return { object: chart.dataTableText ?? {}, prefix: '' };
  return { object: chart, prefix: '', whole: true };
}

export function chartTextOverrides(chart, part) {
  const target = source(chart, part), result = {};
  for (const key of KEYS) {
    const rootKey = key === 'color' ? 'textColor' : key;
    if (chart[rootKey] != null) result[key] = chart[rootKey];
    if (key === 'size' && part?.kind?.startsWith('axis-') && !part.kind.startsWith('axis-title-') && chart.axisSize != null) result.size = chart.axisSize;
    const targetKey = target.whole ? rootKey : prefixed(target.prefix, key);
    if (target.object[targetKey] != null) result[key] = target.object[targetKey];
  }
  if (target.series !== undefined && Number.isInteger(part?.p) && part.p >= 0) {
    const point = target.object.pointLabelStyles?.[part.p];
    for (const key of KEYS) if (point?.[key] != null) result[key] = point[key];
  }
  return result;
}

/** 홈 리본에 표시할 값. 크기는 셀과 같은 pt 단위이며 SVG px 값과 구분한다. */
export function chartTextStyle(chart, part = null) {
  if (!chartTextTargetSupported(chart, part)) return null;
  return { font: '맑은 고딕', size: part?.kind === 'title' ? 14 : 9, color: '#595959', bold: false, italic: false, underline: false, strike: false, ...chartTextOverrides(chart, part) };
}

function cleanDelta(delta) {
  const out = {};
  for (const key of KEYS) {
    if (!own(delta, key)) continue;
    const value = delta[key];
    if (key === 'size') { const number = Number(value); if (Number.isFinite(number) && number >= 1 && number <= 409) out.size = number; }
    else if (BOOLEAN_KEYS.has(key)) out[key] = !!value;
    else if (key === 'color') { if (typeof value === 'string' && /^(?:#[0-9a-f]{3}|#[0-9a-f]{6}|#[0-9a-f]{8}|[a-z]+)$/i.test(value.trim())) out.color = value.trim(); }
    else if (typeof value === 'string' && value.trim()) out[key] = value.trim();
  }
  return out;
}
function patchObject(object, prefix, delta) {
  const out = { ...object };
  for (const [key, value] of Object.entries(delta)) out[prefixed(prefix, key)] = value;
  return out;
}

/** wb.setSheetProp 거래에 넘길 얕은 차트 패치. 원본/다른 요소의 서식은 바꾸지 않는다. */
export function chartTextFormatPatch(chart, part, values) {
  if (!chartTextTargetSupported(chart, part)) return null;
  const delta = cleanDelta(values), target = source(chart, part), kind = part?.kind ?? 'chart';
  if (!Object.keys(delta).length) return {};
  if (target.whole) {
    const patch = {};
    for (const [key, value] of Object.entries(delta)) {
      patch[key === 'color' ? 'textColor' : key] = value;
      for (const prefix of ['title', 'legend']) patch[prefixed(prefix, key)] = value;
      if (key === 'size') patch.axisSize = value;
    }
    patch.axes = { ...chart.axes };
    for (const axis of ['x', 'y', 'y2']) patch.axes[axis] = patchObject(patchObject(chart.axes?.[axis], '', delta), 'title', delta);
    patch.seriesFmt = Array.from({ length: Math.max(chart.series?.length ?? 0, chart.seriesFmt?.length ?? 0) }, (_, i) => {
      const series = { ...chart.series?.[i], ...chart.seriesFmt?.[i] };
      const out = patchObject(chart.seriesFmt?.[i], 'label', delta);
      if (series?.pointLabelStyles) out.pointLabelStyles = Object.fromEntries(Object.entries(series.pointLabelStyles).map(([index, style]) => [index, { ...style, ...delta }]));
      return out;
    });
    patch.dataTableText = { ...chart.dataTableText, ...delta };
    return patch;
  }
  if (target.axis) return { axes: { ...chart.axes, [target.axis]: patchObject(target.object, target.prefix, delta) } };
  if (target.series !== undefined) {
    const formats = [...(chart.seriesFmt ?? [])], existing = chart.seriesFmt?.[target.series] ?? {};
    while (formats.length <= target.series) formats.push({});
    if (Number.isInteger(part?.p) && part.p >= 0) formats[target.series] = { ...existing, pointLabelStyles: { ...target.object.pointLabelStyles, [part.p]: { ...target.object.pointLabelStyles?.[part.p], ...delta } } };
    else {
      formats[target.series] = patchObject(existing, 'label', delta);
      if (target.object.pointLabelStyles) formats[target.series].pointLabelStyles = Object.fromEntries(Object.entries(target.object.pointLabelStyles).map(([index, style]) => [index, { ...style, ...delta }]));
    }
    return { seriesFmt: formats };
  }
  if (kind === 'dataTable') return { dataTableText: { ...chart.dataTableText, ...delta } };
  return Object.fromEntries(Object.entries(delta).map(([key, value]) => [prefixed(target.prefix, key), value]));
}

const escape = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
function setAttribute(tag, name, value) {
  const pattern = new RegExp('\\s' + name + '="[^"]*"', 'g');
  return tag.replace(pattern, '').replace(/>$/, ` ${name}="${escape(value)}">`);
}
const attribute = (tag, name) => new RegExp('\\s' + name + '="([^"]*)"').exec(tag)?.[1];

/** 모든 차트 종류의 SVG 텍스트에 공통 서식을 적용한다. 도형/범례 표지 색은 건드리지 않는다. */
export function chartTextSvg(chart, svg) {
  const groups = [];
  return svg.replace(/<g\b[^>]*>|<\/g>|<text\b[^>]*>/g, tag => {
    if (tag === '</g>') { groups.pop(); return tag; }
    const kind = attribute(tag, 'data-el') ?? groups.at(-1) ?? 'chart';
    if (tag.startsWith('<g')) { if (!tag.endsWith('/>')) groups.push(kind); return tag; }
    const part = { kind };
    if (kind === 'label') {
      part.s = Number(attribute(tag, 'data-s') ?? 0);
      const point = attribute(tag, 'data-p');
      if (point !== undefined) part.p = Number(point);
    }
    const style = chartTextOverrides(chart, part);
    if (/^axis-(?:title-)?(?:x|y|y2)$/.test(kind)) style.size = chartTextStyle(chart, part).size;
    if (style.font !== undefined) tag = setAttribute(tag, 'font-family', `${fontFamilyCandidates(style.font).map(name => `'${name.replace(/['\\]/g, '')}'`).join(',')},'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif`);
    if (style.size !== undefined) tag = setAttribute(tag, 'font-size', Math.round(Number(style.size) * 4 / 3 * 10) / 10);
    if (style.color !== undefined) tag = setAttribute(tag, 'fill', style.color);
    if (style.bold !== undefined) tag = setAttribute(tag, 'font-weight', style.bold ? 700 : 400);
    if (style.italic !== undefined) tag = setAttribute(tag, 'font-style', style.italic ? 'italic' : 'normal');
    if (style.underline !== undefined || style.strike !== undefined) tag = setAttribute(tag, 'text-decoration', [style.underline ? 'underline' : '', style.strike ? 'line-through' : ''].filter(Boolean).join(' ') || 'none');
    return tag;
  });
}

/** 선택된 차트 및 내보내기에서 필요한 웹 글꼴만 지연 로딩한다. */
export function chartTextFontNames(chart) {
  if (!chart) return [];
  const names = new Set();
  const take = object => {
    if (!object) return;
    for (const key of ['font', 'titleFont', 'legendFont', 'labelFont']) if (typeof object[key] === 'string' && object[key].trim()) names.add(object[key].trim());
  };
  take(chart); take(chart.dataTableText);
  for (const axis of Object.values(chart.axes ?? {})) take(axis);
  for (const series of [...(chart.series ?? []), ...(chart.seriesFmt ?? [])]) {
    take(series);
    for (const style of Object.values(series?.pointLabelStyles ?? {})) take(style);
  }
  return [...names];
}
