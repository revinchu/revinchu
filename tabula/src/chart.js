// 차트: 범위 값 → 데이터 해석, SVG 그리기 (DOM 없이 문자열 생성)
// 차트 모델: { type, title, range, sheet, x, y, w, h, z,
//   series?: [{ name: { ref | text }, cat: ref, val: ref, x: ref }]   파일에서 가져온 계열 참조 (ref = { sheet, r1, c1, r2, c2 })
//   pivot?: { sheet, name }                                           피벗 차트 (피벗 결과를 따라감)
//   seriesFmt?: [{ type, axis, color, labels, marker, smooth, numFmt }] 계열별 서식 (순서대로)
//   labels, legend: 'b'|'t'|'r'|'l'|'none', grouping: 'clustered'|'stacked'|'percentStacked',
//   axes?: { y: { title, numFmt, min, max }, y2: {…}, x: { title } } }
import { formatGeneral, formatValue, formatCode } from './format.js';
import { pivotSourceData, pivotChartData } from './pivot.js';

export const CHART_TYPES = [
  { id: 'column', label: '세로 막대형' },
  { id: 'bar', label: '가로 막대형' },
  { id: 'line', label: '꺾은선형' },
  { id: 'area', label: '영역형' },
  { id: 'pie', label: '원형' },
  { id: 'doughnut', label: '도넛형' },
  { id: 'scatter', label: '분산형' },
  { id: 'bubble', label: '거품형' },
  { id: 'radar', label: '방사형' },
  { id: 'waterfall', label: '폭포' },
  { id: 'funnel', label: '깔때기형' },
  { id: 'histogram', label: '히스토그램' },
  { id: 'pareto', label: '파레토' },
  { id: 'treemap', label: '트리맵' },
  { id: 'boxWhisker', label: '상자 수염' },
  { id: 'stock', label: '주식형' },
  { id: 'combo', label: '콤보 (막대 + 꺾은선 보조 축)' },
];

/** 엑셀 [모든 차트] 대화상자와 같은 분류 · 하위 종류 (차트 모델에 덮어쓸 값) */
export const CHART_GALLERY = [
  ['세로 막대형', [['묶은 세로 막대형', { type: 'column', grouping: 'clustered' }], ['누적 세로 막대형', { type: 'column', grouping: 'stacked' }], ['100% 기준 누적 세로 막대형', { type: 'column', grouping: 'percentStacked' }]]],
  ['꺾은선형', [['꺾은선형', { type: 'line', marker: 'none' }], ['누적 꺾은선형', { type: 'line', grouping: 'stacked', marker: 'none' }], ['100% 기준 누적 꺾은선형', { type: 'line', grouping: 'percentStacked', marker: 'none' }], ['표식이 있는 꺾은선형', { type: 'line' }], ['표식이 있는 누적 꺾은선형', { type: 'line', grouping: 'stacked' }]]],
  ['원형', [['원형', { type: 'pie' }], ['쪼개진 원형', { type: 'pie', explode: 12 }], ['도넛형', { type: 'doughnut' }]]],
  ['가로 막대형', [['묶은 가로 막대형', { type: 'bar', grouping: 'clustered' }], ['누적 가로 막대형', { type: 'bar', grouping: 'stacked' }], ['100% 기준 누적 가로 막대형', { type: 'bar', grouping: 'percentStacked' }]]],
  ['영역형', [['영역형', { type: 'area' }], ['누적 영역형', { type: 'area', grouping: 'stacked' }], ['100% 기준 누적 영역형', { type: 'area', grouping: 'percentStacked' }]]],
  ['분산형', [['분산형', { type: 'scatter' }], ['곡선 및 표식이 있는 분산형', { type: 'scatter', scatterStyle: 'smoothMarker' }], ['곡선이 있는 분산형', { type: 'scatter', scatterStyle: 'smooth' }], ['직선 및 표식이 있는 분산형', { type: 'scatter', scatterStyle: 'lineMarker' }], ['직선이 있는 분산형', { type: 'scatter', scatterStyle: 'line' }], ['거품형', { type: 'bubble' }]]],
  ['주식형', [['고가-저가-종가', { type: 'stock' }], ['시가-고가-저가-종가', { type: 'stock', ohlc: true }]]],
  ['방사형', [['방사형', { type: 'radar' }], ['표식이 있는 방사형', { type: 'radar', radarStyle: 'marker' }], ['채워진 방사형', { type: 'radar', radarStyle: 'filled' }]]],
  ['트리맵', [['트리맵', { type: 'treemap' }]]],
  ['히스토그램', [['히스토그램', { type: 'histogram' }], ['파레토', { type: 'pareto' }]]],
  ['상자 수염', [['상자 수염', { type: 'boxWhisker' }]]],
  ['폭포', [['폭포', { type: 'waterfall' }]]],
  ['깔때기형', [['깔때기형', { type: 'funnel' }]]],
  ['콤보', [['묶은 세로 막대형 - 꺾은선형, 보조 축', { type: 'combo' }]]],
];

export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];
/** 색 변경 (엑셀 [차트 디자인] → [색 변경]) + WIXEL 모던 팔레트 */
export const CHART_PALETTES = {
  office: { label: '다양한 색 1 (Office)', colors: PALETTE },
  colorful2: { label: '다양한 색 2', colors: ['#5B9BD5', '#A5A5A5', '#4472C4', '#264478', '#636363', '#255E91'] },
  colorful3: { label: '다양한 색 3', colors: ['#ED7D31', '#FFC000', '#70AD47', '#9E480E', '#997300', '#43682B'] },
  mono1: { label: '단색 파랑', colors: ['#264478', '#335AA1', '#4472C4', '#698ED0', '#8FAADC', '#B4C7E7'] },
  mono2: { label: '단색 주황', colors: ['#843C0C', '#C55A11', '#ED7D31', '#F4B183', '#F8CBAD', '#FBE5D6'] },
  mono6: { label: '단색 녹색', colors: ['#385723', '#548235', '#70AD47', '#A9D18E', '#C5E0B4', '#E2F0D9'] },
  modern: { label: 'WIXEL 모던', colors: ['#4F46E5', '#0EA5E9', '#10B981', '#F59E0B', '#F43F5E', '#8B5CF6', '#64748B', '#14B8A6'] },
  pastel: { label: 'WIXEL 파스텔', colors: ['#818CF8', '#7DD3FC', '#6EE7B7', '#FCD34D', '#FDA4AF', '#C4B5FD', '#CBD5E1', '#5EEAD4'] },
  slate: { label: 'WIXEL 슬레이트', colors: ['#1E293B', '#475569', '#64748B', '#94A3B8', '#CBD5E1', '#0EA5E9'] },
  vivid: { label: 'WIXEL 비비드', colors: ['#2563EB', '#DC2626', '#16A34A', '#D97706', '#9333EA', '#0891B2', '#DB2777', '#65A30D'] },
};
export const paletteOf = (ch) => CHART_PALETTES[ch?.palette]?.colors ?? PALETTE;

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const label = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? formatGeneral(v) : String(v?.code ?? v));

/**
 * 2차원 값 → { categories, series:[{name, values}] }
 * 엑셀처럼: 첫 행/첫 열이 숫자가 아니면 머리글로, 행 수 ≥ 열 수이면 열 단위 계열
 */
export function chartLayout(rows, type = 'column', flip = false) {
  const R = rows.length;
  const C = rows[0]?.length ?? 0;
  const headRow = R > 1 && rows[0].slice(C > 1 ? 1 : 0).every((v) => !isNum(v)) && rows[0].some((v) => v !== null && v !== '');
  const firstDataRow = headRow ? 1 : 0;
  const xy = type === 'scatter' || type === 'bubble';
  const catCol = C > 1 && (xy || rows.slice(firstDataRow).every((r) => !isNum(r[0])));
  const firstDataCol = catCol ? 1 : 0;
  const auto = xy || type === 'stock' || type === 'boxWhisker' || R - firstDataRow >= C - firstDataCol;
  const byCols = flip && !xy ? !auto : auto;
  return { R, C, headRow, catCol, firstDataRow, firstDataCol, byCols };
}

export function chartData(rows, type = 'column', flip = false) {
  if (!rows.length || !rows[0]?.length) return { categories: [], series: [] };
  const { R, C, headRow, catCol, firstDataRow, firstDataCol, byCols } = chartLayout(rows, type, flip);
  const series = [];
  let categories = [];
  if (byCols) {
    categories = rows.slice(firstDataRow).map((r, i) => (catCol ? label(r[0]) : String(i + 1)));
    for (let c = firstDataCol; c < C; c++) {
      series.push({
        name: headRow ? label(rows[0][c]) : `계열${c - firstDataCol + 1}`,
        values: rows.slice(firstDataRow).map((r) => (isNum(r[c]) ? r[c] : null)),
        x: (type === 'scatter' || type === 'bubble') && catCol && rows.slice(firstDataRow).some((r) => isNum(r[0])) ? rows.slice(firstDataRow).map((r) => (isNum(r[0]) ? r[0] : null)) : null,
      });
    }
    // 거품형: (Y, 크기) 열 쌍마다 계열 하나
    if (type === 'bubble') {
      const paired = [];
      for (let k = 0; k < series.length; k += 2) paired.push({ ...series[k], size: series[k + 1]?.values ?? series[k].values.map(() => 1) });
      series.splice(0, series.length, ...paired);
    }
  } else {
    categories = [];
    for (let c = firstDataCol; c < C; c++) categories.push(headRow ? label(rows[0][c]) : String(c - firstDataCol + 1));
    for (let r = firstDataRow; r < R; r++) {
      series.push({
        name: catCol ? label(rows[r][0]) : `계열${r - firstDataRow + 1}`,
        values: rows[r].slice(firstDataCol).map((v) => (isNum(v) ? v : null)),
        x: null,
      });
    }
  }
  // 숫자가 하나도 없는 계열(텍스트 열)은 제외
  const numeric = series.filter((sr) => sr.values.some((v) => v !== null));
  return { categories, series: numeric.length ? numeric : series };
}

/** 참조 값 → 1차원 (한 열이면 행 순서, 한 행이면 열 순서, 여러 열이면 행마다 글자를 이어 붙임) */
function flatRef(rows, asText) {
  if (!rows?.length) return [];
  if (rows.length === 1) return rows[0];
  if (rows[0].length === 1) return rows.map((r) => r[0]);
  return asText ? rows.map((r) => r.filter((v) => v !== null && v !== '').map(label).join(' / ')) : rows.map((r) => r[r.length - 1]);
}

/**
 * 차트 모델 → 그릴 데이터 { categories, series: [{ name, values, x, type, axis, color, labels, marker, numFmt }] }
 * api: { values(ref) → 2차원 값, pivot({ sheet, name }) → { categories, series } | null, range(ch) → 2차원 값 }
 */
export function resolveChart(ch, api) {
  let base;
  if (ch.pivot) base = api.pivot(ch.pivot) ?? { categories: [], series: [] };
  else if (ch.series?.length) {
    const catRef = ch.series.find((s) => s.cat)?.cat;
    const series = ch.series.map((s, i) => {
      const vals = s.val ? flatRef(api.values(s.val), false) : s.cache ?? [];
      const name = s.name?.ref ? label(flatRef(api.values(s.name.ref), true)[0]) : s.name?.text ?? `계열${i + 1}`;
      const xs = s.x ? flatRef(api.values(s.x), false) : null;
      const size = s.size ? flatRef(api.values(s.size), false).map((v) => (isNum(v) ? v : null)) : null;
      return { name, values: vals.map((v) => (isNum(v) ? v : null)), x: xs ? xs.map((v) => (isNum(v) ? v : null)) : null, ...(size ? { size } : {}) };
    });
    const n = Math.max(0, ...series.map((s) => s.values.length));
    const categories = catRef ? flatRef(api.values(catRef), true).map(label) : Array.from({ length: n }, (_, i) => String(i + 1));
    base = { categories, series };
  } else base = chartData(api.range(ch), ch.type === 'combo' ? 'column' : ch.type, !!ch.byRows);
  const fmt = ch.seriesFmt ?? [];
  const comboDefault = (i) => (ch.type === 'combo' ? (i === base.series.length - 1 && base.series.length > 1 ? { type: 'line', axis: 1 } : { type: 'column' }) : {});
  base.series = base.series.map((s, i) => ({ ...s, ...comboDefault(i), ...(fmt[i] ?? {}) }));
  return base;
}

/** 차트 모델 → 그릴 데이터 (범위 · 계열 참조 · 피벗 차트). hostSi: 차트가 있는 시트 */
export function chartModelData(wb, hostSi, ch) {
  const sheetOf = (name) => { const i = name ? wb.sheetIndexByName(name) : hostSi; return i >= 0 ? i : hostSi; };
  const read = (s, rg) => {
    // 행이 아주 많으면 전체 범위에서 고르게 2,000개를 뽑음 (앞부분만 그리지 않게)
    const total = rg.r2 - rg.r1 + 1;
    const step = total > 2000 ? (total - 1) / 1999 : 1;
    const rows = [];
    for (let k = 0; k < Math.min(total, 2000); k++) {
      const r = rg.r1 + Math.round(k * step);
      const row = [];
      for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) row.push(wb.getValue(s, r, c));
      rows.push(row);
    }
    return rows;
  };
  // 정의된 이름 참조 (OFFSET 등으로 바뀌는 범위): 그릴 때마다 이름을 계산
  const nameRange = (ref) => {
    let v;
    try { v = wb.nameValue(ref.name, hostSi, ref.sheet ?? null, null); } catch { v = null; }
    if (!v || v.r1 === undefined) return null;
    return { sheet: v.sheet ?? ref.sheet ?? null, r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 };
  };
  return resolveChart(ch, {
    range: (c) => (c.range ? read(sheetOf(c.sheet), c.range) : []),
    values: (ref) => {
      if (ref.name) { const rg = nameRange(ref); return rg ? read(sheetOf(rg.sheet), rg) : []; }
      return read(sheetOf(ref.sheet ?? ch.sheet), ref);
    },
    pivot: (p) => {
      const sh = wb.sheets[sheetOf(p.sheet)];
      const defs = [sh?.pivot, ...(sh?.pivotsExtra ?? [])].filter(Boolean);
      const def = defs.find((d) => d.name && d.name === p.name) ?? (p.name ? null : defs[0]);
      if (!def) return null;
      const src = pivotSourceData(wb, def);
      if (!src) return null;
      // 날짜 등 숫자 항목은 원본 열의 표시 형식으로
      const header = src.cube.header.map((h) => String(h ?? '').toLowerCase());
      const fieldStyle = (f) => {
        const i = header.indexOf(String(f).toLowerCase());
        return i >= 0 && src.ref && src.si !== undefined ? wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + i) : null;
      };
      return pivotChartData(src, def, fieldStyle);
    },
  });
}

/** 보기 좋은 눈금 */
export function niceScale(min, max, ticks = 5) {
  if (min === max) {
    if (min === 0) return { min: 0, max: 1, step: 0.2 };
    const d = Math.abs(min) * 0.5;
    min -= d; max += d;
  }
  const range = max - min;
  const rough = range / ticks;
  const mag = 10 ** Math.floor(Math.log10(rough));
  const norm = rough / mag;
  const step = (norm <= 1 ? 1 : norm <= 2 ? 2 : norm <= 2.5 ? 2.5 : norm <= 5 ? 5 : 10) * mag;
  const lo = Math.floor(min / step) * step;
  const hi = Math.ceil(max / step) * step;
  return { min: lo, max: hi === lo ? lo + step : hi, step };
}

const escSvg = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

function axisLabel(n, code) {
  if (code && code !== 'General') { try { return formatCode(n, code).text; } catch { /* 기본 */ } }
  const a = Math.abs(n);
  if (a >= 1e6 && a < 1e15 && !code) return `${formatGeneral(Number((n / 1e6).toPrecision(4)))}M`;
  if (Number.isInteger(n)) return n.toLocaleString('en-US');
  return formatGeneral(Number(n.toPrecision(6)));
}

/** 데이터 레이블 글자 (계열 서식: 서식 코드 문자열 또는 셀 서식) */
function valueLabel(v, fmt) {
  try {
    if (typeof fmt === 'string' && fmt && fmt !== 'General') return formatCode(v, fmt).text;
    if (fmt && typeof fmt === 'object' && (fmt.numFmt || fmt.code)) return formatValue(v, fmt).text;
  } catch { /* 기본 */ }
  if (Number.isInteger(v)) return v.toLocaleString('en-US');
  return formatGeneral(Number(v.toPrecision(6)));
}

function truncate(s, maxChars) {
  return [...s].length > maxChars ? `${[...s].slice(0, Math.max(1, maxChars - 1)).join('')}…` : s;
}

const MARKERS = {
  circle: (x, y, r, c) => `<circle cx="${x}" cy="${y}" r="${r}" fill="${c}"/>`,
  square: (x, y, r, c) => `<rect x="${x - r}" y="${y - r}" width="${2 * r}" height="${2 * r}" fill="${c}"/>`,
  diamond: (x, y, r, c) => `<path d="M${x},${y - r - 1}L${x + r + 1},${y}L${x},${y + r + 1}L${x - r - 1},${y}Z" fill="${c}"/>`,
  triangle: (x, y, r, c) => `<path d="M${x},${y - r - 1}L${x + r + 1},${y + r}L${x - r - 1},${y + r}Z" fill="${c}"/>`,
};

/**
 * chart: 차트 모델 (+ w, h), data: resolveChart 결과 → SVG 문자열
 */
let svgSeq = 0;
export function renderChartSvg(chart, data) {
  const W = Math.max(120, chart.w);
  const H = Math.max(90, chart.h);
  const FONT = "font-family=\"'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif\"";
  // 글꼴 크기(px): 엑셀 기본값 — 제목 14pt, 축 · 범례 · 레이블 9pt (파일에 적힌 크기가 있으면 그것)
  const pt = (v) => Math.round(v * (4 / 3) * 10) / 10;
  const FS = { title: pt(chart.titleSize ?? 14), axis: pt(chart.axisSize ?? 9), legend: pt(chart.legendSize ?? 9) };
  const CW = FS.axis * 0.58; // 글자 하나의 대략 너비
  const LW = FS.legend * 0.58;
  const uid = `c${(++svgSeq).toString(36)}`;
  const defs = [];
  const pal = paletteOf(chart);
  const TXT = chart.textColor ?? '#595959';
  const GRID = chart.gridColor ?? '#d9d9d9';
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ${FONT}>`, '',
    `<rect x="0.5" y="0.5" width="${W - 1}" height="${H - 1}" rx="${chart.rounded ? 8 : 0}" fill="${chart.fill ?? '#fff'}"${chart.border ? ` stroke="${chart.border}"` : ''}/>`];
  let top = 10;
  if (chart.title) {
    parts.push(`<text x="${W / 2}" y="${Math.round(FS.title + 10)}" text-anchor="middle" font-size="${FS.title}"${chart.titleBold ? ' font-weight="700"' : ''} fill="${chart.titleColor ?? TXT}">${escSvg(truncate(chart.title, Math.floor(W / (FS.title * 0.62))))}</text>`);
    top = Math.round(FS.title * 1.5 + 16);
  }
  // 계열 채우기: 단색 또는 그라데이션 (url), 그림자 필터
  const fillOf = (s) => {
    if (!s.grad) return s.color;
    const id = `${uid}g${defs.length}`;
    const a = ((s.grad.ang ?? 90) * Math.PI) / 180;
    const x = Math.cos(a) / 2;
    const y = Math.sin(a) / 2;
    defs.push(`<linearGradient id="${id}" x1="${(0.5 - x).toFixed(3)}" y1="${(0.5 - y).toFixed(3)}" x2="${(0.5 + x).toFixed(3)}" y2="${(0.5 + y).toFixed(3)}">${s.grad.stops.map(([o, c]) => `<stop offset="${o}" stop-color="${c}"/>`).join('')}</linearGradient>`);
    return `url(#${id})`;
  };
  let shadowId = null;
  const shadowAttr = (s) => {
    if (!s.shadow) return '';
    if (!shadowId) { shadowId = `${uid}s`; defs.push(`<filter id="${shadowId}" x="-20%" y="-20%" width="140%" height="140%"><feDropShadow dx="0" dy="1.5" stdDeviation="2" flood-color="#000" flood-opacity="0.45"/></filter>`); }
    return ` filter="url(#${shadowId})"`;
  };
  const finish = () => { if (defs.length) parts[1] = `<defs>${defs.join('')}</defs>`; parts.push('</svg>'); return parts.join(''); };
  const { categories } = data;
  const baseType = chart.type === 'combo' ? 'column' : chart.type;
  const series = data.series.map((s, i) => ({ ...s, type: s.type ?? baseType, axis: s.axis ?? 0, color: s.color ?? pal[i % pal.length] }));
  const pieLike = baseType === 'pie' || baseType === 'doughnut';
  const special = SPECIAL[baseType];
  const legendPos = chart.legend ?? (special?.legend === false ? 'none' : 'b');
  const legendItems = pieLike || baseType === 'treemap' ? categories.map((c, i) => ({ name: c, color: series[0]?.colors?.[i] ?? pal[i % pal.length], line: false }))
    : baseType === 'waterfall' ? [{ name: '증가', color: chart.upColor ?? pal[0] }, { name: '감소', color: chart.downColor ?? pal[1] }, { name: '합계', color: chart.totalColor ?? pal[2] }]
      : baseType === 'pareto' ? [{ name: series[0]?.name ?? '', color: series[0]?.color }, { name: '누적 %', color: pal[1], line: true }]
        : series.map((s) => ({ name: s.name, color: s.color, line: s.type === 'line' || s.type === 'radar' }));
  const showLegend = legendPos !== 'none' && (legendItems.length > 1 || pieLike || baseType === 'treemap');
  const sideLegend = showLegend && (legendPos === 'r' || legendPos === 'l');
  const legendW = sideLegend ? Math.min(180, Math.max(60, ...legendItems.map((it) => [...String(it.name)].length * LW * 1.4 + 22))) : 0;
  const legendH = showLegend && !sideLegend ? Math.round(FS.legend * 2.2) : 0;
  const plot = {
    x: 10 + (legendPos === 'l' ? legendW : 0), y: top + (legendPos === 't' ? legendH : 0),
    w: W - 20 - legendW, h: H - top - 10 - legendH,
  };

  if (!series.length || series.every((s) => s.values.every((v) => v === null))) {
    parts.push(`<text x="${W / 2}" y="${H / 2}" text-anchor="middle" font-size="12" fill="#999">표시할 숫자 데이터가 없습니다</text></svg>`);
    return parts.join('');
  }

  if (showLegend) {
    const items = legendItems.slice(0, 16);
    const key = (it, lx, ly) => (it.line
      ? `<line x1="${lx - 2}" y1="${ly - 4}" x2="${lx + 11}" y2="${ly - 4}" stroke="${it.color}" stroke-width="2.25"/><circle cx="${lx + 4.5}" cy="${ly - 4}" r="2.5" fill="${it.color}"/>`
      : `<rect x="${lx}" y="${ly - 8}" width="9" height="9" fill="${it.color}"/>`);
    if (sideLegend) {
      const lx0 = legendPos === 'r' ? W - legendW - 4 : 8;
      let ly = Math.max(top + 12, H / 2 - (items.length * Math.round(FS.legend * 1.6)) / 2);
      items.forEach((it) => {
        parts.push(key(it, lx0, ly), `<text x="${lx0 + 15}" y="${ly}" font-size="${FS.legend}" fill="${TXT}">${escSvg(truncate(String(it.name), Math.floor((legendW - 20) / (LW * 1.4))))}</text>`);
        ly += Math.round(FS.legend * 1.6);
      });
    } else {
      const itemW = Math.min(160, Math.max(60, (W - 20) / items.length));
      const total = itemW * items.length;
      let lx = Math.max(10, (W - total) / 2);
      const ly = legendPos === 't' ? top + 12 : H - 14;
      items.forEach((it) => {
        parts.push(key(it, lx, ly), `<text x="${lx + 15}" y="${ly}" font-size="${FS.legend}" fill="${TXT}">${escSvg(truncate(String(it.name), Math.floor((itemW - 18) / (LW * 1.4))))}</text>`);
        lx += itemW;
      });
    }
  }

  const wantLabels = (s) => s.labels ?? chart.labels ?? false;
  if (chart.plotFill) parts.push(`<rect x="${plot.x}" y="${plot.y}" width="${plot.w}" height="${plot.h}" fill="${chart.plotFill}"/>`);
  if (special) {
    special.draw({ chart, series, categories, plot, parts, FS, TXT, GRID, pal, defs, uid, wantLabels, W, H });
    return finish();
  }

  if (pieLike) {
    const s0 = series[0];
    const vals = s0.values.map((v) => (v && v > 0 ? v : 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    const cx = plot.x + plot.w / 2;
    const cy = plot.y + plot.h / 2;
    const ex = chart.explode ? Math.min(0.3, chart.explode / 100) : 0;
    // 항목 이름 레이블(엑셀의 원형 바깥 레이블)이 있으면 원을 줄여 둘레에 글자 자리를 둠
    const outside = s0.catName || s0.labelPos === 'out';
    const r = Math.max(10, (Math.min(plot.w, plot.h) / 2 - 6) / (1 + ex) * (outside ? 0.72 : 1));
    const inner = baseType === 'doughnut' ? r * ((chart.hole ?? 50) / 100) : 0;
    let a = -Math.PI / 2 + ((chart.firstAngle ?? 0) * Math.PI) / 180;
    vals.forEach((v, i) => {
      if (!v) return;
      const frac = v / sum;
      const a2 = a + frac * Math.PI * 2;
      const color = s0.colors?.[i] ?? pal[i % pal.length];
      const mid0 = (a + a2) / 2;
      const ox = ex ? Math.cos(mid0) * r * ex : 0;
      const oy = ex ? Math.sin(mid0) * r * ex : 0;
      if (ex) parts.push(`<g transform="translate(${ox.toFixed(2)},${oy.toFixed(2)})">`);
      if (frac >= 0.9999) {
        parts.push(`<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}" stroke="#fff"/>`);
        if (inner) parts.push(`<circle cx="${cx}" cy="${cy}" r="${inner}" fill="#fff"/>`);
      } else {
        const large = a2 - a > Math.PI ? 1 : 0;
        const p = (ang, rad) => `${(cx + Math.cos(ang) * rad).toFixed(2)},${(cy + Math.sin(ang) * rad).toFixed(2)}`;
        const d = inner
          ? `M${p(a, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}L${p(a2, inner)}A${inner},${inner} 0 ${large} 0 ${p(a, inner)}Z`
          : `M${cx},${cy}L${p(a, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}Z`;
        parts.push(`<path d="${d}" fill="${color}" stroke="${chart.fill ?? '#fff'}" stroke-width="1.5"/>`);
      }
      if (ex) parts.push('</g>');
      // 레이블: false = 없음(파일에 없던 원형), pct = 백분율, labels = 값, 정하지 않음 = 백분율
      if (outside && (s0.labels !== false || s0.catName)) {
        // 바깥 레이블: 항목 이름 + (백분율 · 값) 두 줄, 조각 색 글자
        const mid = (a + a2) / 2;
        const lr = r * 1.2 + 6;
        const lx = cx + ox + Math.cos(mid) * lr;
        const ly = cy + oy + Math.sin(mid) * lr;
        const anchor = Math.cos(mid) > 0.25 ? 'start' : Math.cos(mid) < -0.25 ? 'end' : 'middle';
        const second = s0.labels === false ? '' : !s0.pct && wantLabels(s0) ? valueLabel(v, s0.numFmt) : `${Math.round(frac * 100)}%`;
        const fsz = s0.labelSize ? s0.labelSize * (4 / 3) : 11;
        const tc = s0.labelColor ?? color;
        const name = s0.catName ? String(categories[i] ?? '') : '';
        const lines = [name, second].filter(Boolean);
        lines.forEach((t, k) => parts.push(`<text x="${lx.toFixed(1)}" y="${(ly + 4 + (k - (lines.length - 1) / 2) * (fsz + 2)).toFixed(1)}" text-anchor="${anchor}" font-size="${fsz}" fill="${tc}" font-weight="${k === lines.length - 1 && second ? 700 : 400}">${escSvg(t)}</text>`));
      } else if (frac >= 0.04 && s0.labels !== false) {
        const mid = (a + a2) / 2;
        const lr = inner ? (r + inner) / 2 : r * 0.65;
        const txt = !s0.pct && wantLabels(s0) ? valueLabel(v, s0.numFmt) : `${Math.round(frac * 100)}%`;
        parts.push(`<text x="${(cx + ox + Math.cos(mid) * lr).toFixed(1)}" y="${(cy + oy + Math.sin(mid) * lr + 4).toFixed(1)}" text-anchor="middle" font-size="11" fill="#fff" font-weight="700">${escSvg(txt)}</text>`);
      }
      a = a2;
    });
    return finish();
  }

  const horizontal = baseType === 'bar';
  const stacked = chart.grouping === 'stacked' || chart.grouping === 'percentStacked';
  const pct = chart.grouping === 'percentStacked';
  const n = categories.length;
  const barTypes = new Set(['column', 'bar']);

  // 축별 값 범위 (누적이면 합계 기준)
  const scaleFor = (axis) => {
    const ss = series.filter((s) => s.axis === axis && (s.type !== 'scatter' || baseType === 'scatter'));
    if (!ss.length) return null;
    let vals = ss.flatMap((s) => s.values).filter(isNum);
    const bars = ss.filter((s) => barTypes.has(s.type) || s.type === 'area' || (s.type === 'line' && stacked));
    if (stacked && bars.length > 1) {
      vals = [];
      for (let i = 0; i < n; i++) {
        let pos = 0;
        let neg = 0;
        for (const s of bars) { const v = s.values[i]; if (isNum(v)) { if (v >= 0) pos += v; else neg += v; } }
        vals.push(pos, neg);
      }
      if (pct) vals = [0, 1];
      vals.push(...ss.filter((s) => !bars.includes(s)).flatMap((s) => s.values).filter(isNum));
    }
    if (!vals.length) vals = [0, 1];
    const cfg = chart.axes?.[axis ? 'y2' : 'y'] ?? {};
    let min = Math.min(0, ...vals);
    let max = Math.max(0, ...vals);
    if (ss.every((s) => s.type === 'line') && !stacked) {
      min = Math.min(...vals);
      max = Math.max(...vals);
      if (min > 0 && min < max * 0.5) min = 0;
    }
    const sc = niceScale(isNum(cfg.min) ? cfg.min : min, isNum(cfg.max) ? cfg.max : max);
    if (isNum(cfg.min)) sc.min = cfg.min;
    if (isNum(cfg.max)) sc.max = cfg.max;
    if (isNum(cfg.major) && cfg.major > 0 && (sc.max - sc.min) / cfg.major <= 200) sc.step = cfg.major;
    const ticks = [];
    for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) ticks.push(Number(t.toPrecision(12)));
    const srcCode = ss.map((s) => s.numFmt).find((f) => typeof f === 'string' && f !== 'General') ?? null;
    return { ...sc, ticks, code: pct ? '0%' : cfg.numFmt ?? srcCode };
  };
  const scale = scaleFor(0) ?? scaleFor(1);
  const scale2 = series.some((s) => s.axis === 1) && series.some((s) => s.axis === 0) ? scaleFor(1) : null;
  const primaryAxis = scaleFor(0) ? 0 : 1;
  const hideY = !!chart.axes?.y?.hide;
  const hideX = !!chart.axes?.x?.hide;
  const labelW = hideY ? 6 : Math.min(100, Math.max(...scale.ticks.map((t) => axisLabel(t, scale.code).length)) * CW + 8);
  const label2W = scale2 ? Math.min(100, Math.max(...scale2.ticks.map((t) => axisLabel(t, scale2.code).length)) * CW + 8) : 0;

  // 가로축 값 범위 (분산형)
  let xScale = null;
  if (baseType === 'scatter') {
    const xs = series.flatMap((s) => (s.x ?? s.values.map((_, i) => i + 1))).filter(isNum);
    xScale = niceScale(Math.min(...xs), Math.max(...xs));
  }

  const catLabelW = horizontal ? Math.min(140, Math.max(...categories.map((c) => [...c].length)) * CW * 1.4 + 8) : 0;
  const area = horizontal
    ? { x: plot.x + catLabelW, y: plot.y, w: plot.w - catLabelW - 10, h: plot.h - 18 }
    : { x: plot.x + labelW, y: plot.y + 4, w: plot.w - labelW - 6 - label2W, h: plot.h - Math.round(FS.axis * 1.9) };
  if (area.w < 20 || area.h < 20) return finish();
  const posFor = (sc) => (v) => (horizontal
    ? area.x + ((v - sc.min) / (sc.max - sc.min)) * area.w
    : area.y + area.h - ((v - sc.min) / (sc.max - sc.min)) * area.h);
  const vpos = posFor(scale);
  const vpos2 = scale2 ? posFor(scale2) : vpos;
  const posOf = (s) => (s.axis === 1 && scale2 ? vpos2 : s.axis === primaryAxis || !scale2 ? vpos : vpos2);

  // 눈금선 + 값 축 레이블
  for (const t of scale.ticks) {
    const p = vpos(t).toFixed(1);
    if (horizontal) {
      parts.push(chart.gridY === false ? '' : `<line x1="${p}" y1="${area.y}" x2="${p}" y2="${area.y + area.h}" stroke="${GRID}"/>`,
        hideY ? '' : `<text x="${p}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, scale.code))}</text>`);
    } else {
      parts.push(chart.gridY === false ? '' : `<line x1="${area.x}" y1="${p}" x2="${area.x + area.w}" y2="${p}" stroke="${GRID}"/>`,
        hideY ? '' : `<text x="${area.x - 5}" y="${Number(p) + FS.axis * 0.35}" text-anchor="end" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, scale.code))}</text>`);
    }
  }
  if (scale2 && !horizontal) {
    for (const t of scale2.ticks) {
      const p = vpos2(t);
      parts.push(`<text x="${area.x + area.w + 5}" y="${(p + 3.5).toFixed(1)}" text-anchor="start" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(t, scale2.code))}</text>`);
    }
  }
  const axisTitle = (txt, x, y, rot) => (txt ? `<text x="${x}" y="${y}" text-anchor="middle" font-size="11" fill="${TXT}"${rot ? ` transform="rotate(${rot} ${x} ${y})"` : ''}>${escSvg(txt)}</text>` : '');

  const base = vpos(Math.max(scale.min, Math.min(0, scale.max)));
  const labelsOut = [];
  const pushLabel = (x, y, v, s, anchor = 'middle') => {
    labelsOut.push(`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${pt(s.labelSize ?? 9)}"${s.labelBold ? ' font-weight="700"' : ''} fill="${s.labelColor ?? '#404040'}" paint-order="stroke" stroke="#fff" stroke-width="2.5">${escSvg(valueLabel(v, s.numFmt))}</text>`);
  };

  if (baseType === 'scatter') {
    const xpos = (x) => area.x + ((x - xScale.min) / (xScale.max - xScale.min)) * area.w;
    for (let t = xScale.min; t <= xScale.max + xScale.step / 2; t += xScale.step) {
      parts.push(`<text x="${xpos(t).toFixed(1)}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="${TXT}">${escSvg(axisLabel(Number(t.toPrecision(12))))}</text>`);
    }
    const sty = chart.scatterStyle ?? 'marker';
    series.forEach((s) => {
      const pts = [];
      s.values.forEach((v, i) => {
        const x = s.x ? s.x[i] : i + 1;
        if (isNum(v) && isNum(x)) pts.push([xpos(x), vpos(v), v]);
      });
      if (/line|smooth/i.test(sty) && pts.length > 1) {
        const d = /smooth/i.test(sty) ? smoothPath(pts) : pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
        parts.push(`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.lineWidth ?? 2}" stroke-linejoin="round"/>`);
      }
      if (sty === 'marker' || /Marker$/.test(sty)) for (const p of pts) parts.push((MARKERS[s.marker] ?? MARKERS.circle)(p[0].toFixed(1), p[1].toFixed(1), 3.5, s.markerColor ?? s.color));
      if (wantLabels(s)) for (const p of pts) pushLabel(p[0], p[1] - 7, p[2], s);
    });
  } else {
    // 항목 축 레이블
    const band = (horizontal ? area.h : area.w) / Math.max(1, n);
    const maxChars = Math.max(2, Math.floor(band / (CW * 1.4)));
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horizontal ? area.h : area.w) / 28))));
    if (chart.gridX) {
      for (let i = 1; i < n; i++) {
        const q = ((horizontal ? area.y : area.x) + band * i).toFixed(1);
        parts.push(horizontal ? `<line x1="${area.x}" y1="${q}" x2="${area.x + area.w}" y2="${q}" stroke="${GRID}"/>` : `<line x1="${q}" y1="${area.y}" x2="${q}" y2="${area.y + area.h}" stroke="${GRID}"/>`);
      }
    }
    categories.forEach((c, i) => {
      if (i % every || hideX) return;
      const mid = (horizontal ? area.y : area.x) + band * (i + 0.5);
      if (horizontal) parts.push(`<text x="${area.x - 5}" y="${(mid + 3.5).toFixed(1)}" text-anchor="end" font-size="${FS.axis}" fill="${TXT}">${escSvg(truncate(c, 16))}</text>`);
      else parts.push(`<text x="${mid.toFixed(1)}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="${TXT}">${escSvg(truncate(c, maxChars * every))}</text>`);
    });
    // 막대 (묶은 · 누적)
    const bars = series.filter((s) => barTypes.has(s.type));
    if (bars.length) {
      const k = stacked ? 1 : bars.length;
      const groupW = isNum(chart.gap) ? band / (1 + Math.max(0, chart.gap) / 100) : band * 0.65;
      const barW = groupW / k;
      const posAcc = new Array(n).fill(0);
      const negAcc = new Array(n).fill(0);
      const totals = pct ? categories.map((_, i) => bars.reduce((a, s) => a + Math.abs(isNum(s.values[i]) ? s.values[i] : 0), 0) || 1) : null;
      bars.forEach((s, bi) => {
        const vp = posOf(s);
        const bf = fillOf(s);
        const sh = shadowAttr(s);
        const b0 = vp(Math.max(scale.min, Math.min(0, scale.max)));
        s.values.forEach((raw, i) => {
          if (!isNum(raw)) return;
          const v = pct ? raw / totals[i] : raw;
          const start = (horizontal ? area.y : area.x) + band * i + (band - groupW) / 2 + barW * (stacked ? 0 : bi);
          let from = b0;
          let to = vp(v);
          if (stacked) {
            const acc = v >= 0 ? posAcc : negAcc;
            from = vp(acc[i]);
            to = vp(acc[i] + v);
            acc[i] += v;
          }
          const a = Math.min(from, to);
          const len = Math.abs(to - from);
          if (horizontal) parts.push(`<rect x="${a.toFixed(1)}" y="${start.toFixed(1)}" width="${len.toFixed(1)}" height="${Math.max(1, barW - 1).toFixed(1)}" fill="${bf}"${sh}/>`);
          else parts.push(`<rect x="${start.toFixed(1)}" y="${a.toFixed(1)}" width="${Math.max(1, barW - 1).toFixed(1)}" height="${len.toFixed(1)}" fill="${bf}"${sh}/>`);
          if (wantLabels(s)) {
            if (horizontal) pushLabel((stacked ? (from + to) / 2 : Math.max(from, to) + 4), start + barW / 2 + 3.5, raw, s, stacked ? 'middle' : 'start');
            else pushLabel(start + barW / 2, stacked ? (from + to) / 2 + 3.5 : Math.min(from, to) - 4, raw, s);
          }
        });
      });
    }
    // 영역 · 꺾은선 (막대 위에)
    // 누적 꺾은선 · 영역: 앞 계열 값 위에 쌓음 (100% 기준이면 항목 합계로 나눔)
    const stackAcc = { area: new Array(n).fill(0), line: new Array(n).fill(0) };
    const stackTot = (type) => categories.map((_, i) => series.filter((s) => s.type === type).reduce((a, s) => a + Math.abs(isNum(s.values[i]) ? s.values[i] : 0), 0) || 1);
    const tots = { area: pct ? stackTot('area') : null, line: pct ? stackTot('line') : null };
    for (const type of ['area', 'line']) {
      series.filter((s) => s.type === type).forEach((s) => {
        const vp = posOf(s);
        const lower = stacked ? [...stackAcc[type]] : null;
        const pts = s.values.map((v0, i) => {
          if (!isNum(v0)) return null;
          let v = pct ? v0 / tots[type][i] : v0;
          if (stacked) { stackAcc[type][i] += v; v = stackAcc[type][i]; }
          return [area.x + band * (i + 0.5), vp(v), v0, i];
        });
        const segs = [];
        let cur = [];
        for (const p of pts) { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } }
        if (cur.length) segs.push(cur);
        const b0 = vp(Math.max(scale.min, Math.min(0, scale.max)));
        for (const seg of segs) {
          const d = s.smooth ? smoothPath(seg) : seg.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
          if (type === 'area') {
            const back = lower
              ? seg.slice().reverse().map((p) => `L${p[0].toFixed(1)},${vp(lower[p[3]]).toFixed(1)}`).join('')
              : `L${seg.at(-1)[0].toFixed(1)},${b0.toFixed(1)}L${seg[0][0].toFixed(1)},${b0.toFixed(1)}`;
            parts.push(`<path d="${d}${back}Z" fill="${fillOf(s)}" fill-opacity="${stacked ? 0.9 : 0.75}"/>`);
          } else {
            parts.push(`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.lineWidth ?? 2.25}" stroke-linejoin="round" stroke-linecap="round"${shadowAttr(s)}/>`);
            const mkName = s.marker ?? chart.marker;
            if (mkName !== false && mkName !== 'none') {
              const mk = MARKERS[mkName] ?? MARKERS.circle;
              const mr = s.markerSize ? (s.markerSize * 4) / 3 / 2 : 3;
              for (const p of seg) parts.push(mk(p[0].toFixed(1), p[1].toFixed(1), mr, s.markerColor ?? s.color));
            }
          }
          if (wantLabels(s)) for (const p of seg) pushLabel(p[0], p[1] - 7, p[2], s);
        }
      });
    }
  }
  // 기준선
  if (horizontal) parts.push(`<line x1="${base}" y1="${area.y}" x2="${base}" y2="${area.y + area.h}" stroke="#bfbfbf"/>`);
  else parts.push(`<line x1="${area.x}" y1="${base}" x2="${area.x + area.w}" y2="${base}" stroke="#bfbfbf"/>`);
  parts.push(...labelsOut);
  parts.push(axisTitle(chart.axes?.y?.title, plot.x + 6, area.y + area.h / 2, -90));
  if (scale2) parts.push(axisTitle(chart.axes?.y2?.title, area.x + area.w + label2W - 2, area.y + area.h / 2, 90));
  parts.push(axisTitle(chart.axes?.x?.title, area.x + area.w / 2, H - (legendH ? legendH + 2 : 2)));
  return finish();
}

/** 부드러운 곡선 (Catmull-Rom → 베지어) */
function smoothPath(pts) {
  let d = `M${pts[0][0].toFixed(1)},${pts[0][1].toFixed(1)}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] ?? p2;
    const c1 = [p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6];
    const c2 = [p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6];
    d += `C${c1[0].toFixed(1)},${c1[1].toFixed(1)} ${c2[0].toFixed(1)},${c2[1].toFixed(1)} ${p2[0].toFixed(1)},${p2[1].toFixed(1)}`;
  }
  return d;
}

// ───────────── 통계 · 계층 차트 (엑셀 2016 이후 차트) 와 방사형 · 거품형 · 주식형 ─────────────
const T = (x, y, txt, size, fill, anchor = 'middle', extra = '') => `<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${size}" fill="${fill}"${extra}>${escSvg(txt)}</text>`;

/** 값 축이 있는 직교 영역 (세로 값 축 + 항목 축) — 특수 차트 공용 */
function cartesian(ctx, vals, { horizontal = false, code = null, cats = null, zero = true, right = 0 } = {}) {
  const { chart, plot, parts, FS, TXT, GRID } = ctx;
  const cfg = chart.axes?.y ?? {};
  let lo = Math.min(...vals.filter(isNum));
  let hi = Math.max(...vals.filter(isNum));
  if (zero) { lo = Math.min(0, lo); hi = Math.max(0, hi); }
  if (!Number.isFinite(lo)) { lo = 0; hi = 1; }
  const sc = niceScale(isNum(cfg.min) ? cfg.min : lo, isNum(cfg.max) ? cfg.max : hi);
  if (isNum(cfg.min)) sc.min = cfg.min;
  if (isNum(cfg.max)) sc.max = cfg.max;
  if (isNum(cfg.major) && cfg.major > 0) sc.step = cfg.major;
  const ticks = [];
  for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) ticks.push(Number(t.toPrecision(12)));
  const fmt = cfg.numFmt ?? code;
  const CW = FS.axis * 0.58;
  const labelW = horizontal ? Math.min(140, Math.max(...(cats ?? ['']).map((c) => [...String(c)].length)) * CW * 1.4 + 8) : Math.min(100, Math.max(...ticks.map((t) => axisLabel(t, fmt).length)) * CW + 8);
  const area = horizontal
    ? { x: plot.x + labelW, y: plot.y, w: plot.w - labelW - 10, h: plot.h - 18 }
    : { x: plot.x + labelW, y: plot.y + 4, w: plot.w - labelW - 6 - right, h: plot.h - Math.round(FS.axis * 1.9) };
  const pos = (v) => (horizontal ? area.x + ((v - sc.min) / (sc.max - sc.min)) * area.w : area.y + area.h - ((v - sc.min) / (sc.max - sc.min)) * area.h);
  for (const t of ticks) {
    const p = pos(t);
    if (horizontal) parts.push(chart.gridY === false ? '' : `<line x1="${p.toFixed(1)}" y1="${area.y}" x2="${p.toFixed(1)}" y2="${area.y + area.h}" stroke="${GRID}"/>`, T(p, area.y + area.h + FS.axis * 1.35, axisLabel(t, fmt), FS.axis, TXT));
    else parts.push(chart.gridY === false ? '' : `<line x1="${area.x}" y1="${p.toFixed(1)}" x2="${area.x + area.w}" y2="${p.toFixed(1)}" stroke="${GRID}"/>`, T(area.x - 5, p + FS.axis * 0.35, axisLabel(t, fmt), FS.axis, TXT, 'end'));
  }
  const n = cats?.length ?? 0;
  const band = (horizontal ? area.h : area.w) / Math.max(1, n);
  if (cats) {
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horizontal ? area.h : area.w) / 28))));
    cats.forEach((c, i) => {
      if (i % every) return;
      const mid = (horizontal ? area.y : area.x) + band * (i + 0.5);
      if (horizontal) parts.push(T(area.x - 5, mid + 3.5, truncate(String(c), 16), FS.axis, TXT, 'end'));
      else parts.push(T(mid, area.y + area.h + FS.axis * 1.35, truncate(String(c), Math.max(2, Math.floor((band * every) / (CW * 1.4)))), FS.axis, TXT));
    });
  }
  return { area, pos, band, sc, fmt };
}

const labelTxt = (ctx, s, v, x, y, anchor = 'middle') => ctx.parts.push(`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${(s?.labelSize ?? 9) * 4 / 3}" fill="${s?.labelColor ?? '#404040'}" paint-order="stroke" stroke="#fff" stroke-width="2.5">${escSvg(valueLabel(v, s?.numFmt))}</text>`);

/** 사분위수 (엑셀 QUARTILE.EXC 가 아닌 포함 방식: 상자 수염 차트 기본값) */
function quantile(sorted, q) {
  if (!sorted.length) return 0;
  const pos = (sorted.length - 1) * q;
  const lo = Math.floor(pos);
  const hi = Math.ceil(pos);
  return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
}

/** 히스토그램 구간: 엑셀 자동 구간(스콧 규칙)과 비슷하게 */
export function histogramBins(values, count = null, width = null) {
  const v = values.filter(isNum).sort((a, b) => a - b);
  if (!v.length) return [];
  const min = v[0];
  const max = v.at(-1);
  let w = width;
  if (!w) {
    const n = count ?? Math.max(1, Math.min(30, Math.ceil(Math.sqrt(v.length))));
    w = (max - min) / n || 1;
    const mag = 10 ** Math.floor(Math.log10(w));
    w = Math.ceil(w / mag * 2) / 2 * mag;
  }
  const bins = [];
  for (let a = min; a <= max || !bins.length; a += w) {
    const b = a + w;
    bins.push({ from: a, to: b, count: 0 });
    if (bins.length > 200) break;
  }
  for (const x of v) {
    let k = Math.min(bins.length - 1, Math.floor((x - min) / w));
    if (k > 0 && x === bins[k].from) k -= 1; // (a, b] — 엑셀처럼 오른쪽 끝 포함
    bins[k].count++;
  }
  return bins;
}

const binLabel = (b, i) => `${i ? '(' : '['}${axisLabel(Number(b.from.toPrecision(10)))}, ${axisLabel(Number(b.to.toPrecision(10)))}]`;
const TOTAL_RE = /^(합계|총계|총합계|소계|전체|total|subtotal|grand total|net)$/i;

export const SPECIAL = {
  waterfall: {
    draw(ctx) {
      const { chart, series, categories, parts, pal } = ctx;
      const s = series[0];
      if (!s) return;
      const totals = new Set(chart.totals ?? categories.map((c, i) => (TOTAL_RE.test(String(c).trim()) ? i : -1)).filter((i) => i >= 0));
      let run = 0;
      const bars = s.values.map((v0, i) => {
        const v = isNum(v0) ? v0 : 0;
        if (totals.has(i)) { run = i === 0 ? v : run; return { a: 0, b: totals.has(i) && i > 0 ? run : v, v: totals.has(i) && i > 0 ? run : v, kind: 'total' }; }
        const a = run;
        run += v;
        return { a, b: run, v, kind: v >= 0 ? 'up' : 'down' };
      });
      const { area, pos, band } = cartesian(ctx, bars.flatMap((b) => [b.a, b.b]), { cats: categories, code: s.numFmt });
      const w = band * 0.62;
      const col = { up: chart.upColor ?? pal[0], down: chart.downColor ?? pal[1], total: chart.totalColor ?? pal[2] };
      bars.forEach((b, i) => {
        const x = area.x + band * i + (band - w) / 2;
        const y1 = pos(Math.max(b.a, b.b));
        const y2 = pos(Math.min(b.a, b.b));
        parts.push(`<rect x="${x.toFixed(1)}" y="${y1.toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(1, y2 - y1).toFixed(1)}" fill="${col[b.kind]}"/>`);
        // 연결선
        if (i < bars.length - 1 && chart.connectors !== false) {
          const yb = pos(b.kind === 'total' ? b.b : b.b);
          parts.push(`<line x1="${(x + w).toFixed(1)}" y1="${yb.toFixed(1)}" x2="${(x + band).toFixed(1)}" y2="${yb.toFixed(1)}" stroke="#a6a6a6" stroke-dasharray="2,2"/>`);
        }
        if ((s.labels ?? chart.labels) !== false) labelTxt(ctx, s, b.v, x + w / 2, y1 - 4);
      });
      parts.push(`<line x1="${area.x}" y1="${pos(0).toFixed(1)}" x2="${area.x + area.w}" y2="${pos(0).toFixed(1)}" stroke="#bfbfbf"/>`);
    },
  },
  funnel: {
    legend: false,
    draw(ctx) {
      const { series, categories, plot, parts, FS, TXT } = ctx;
      const s = series[0];
      if (!s) return;
      const vals = s.values.map((v) => (isNum(v) && v > 0 ? v : 0));
      const max = Math.max(...vals, 1);
      const CW = FS.axis * 0.58;
      const lw = Math.min(160, Math.max(...categories.map((c) => [...String(c)].length)) * CW * 1.4 + 10);
      const area = { x: plot.x + lw, y: plot.y + 2, w: plot.w - lw - 4, h: plot.h - 4 };
      const band = area.h / Math.max(1, vals.length);
      vals.forEach((v, i) => {
        const bw = (v / max) * area.w;
        const x = area.x + (area.w - bw) / 2;
        const y = area.y + band * i + band * 0.08;
        parts.push(`<rect x="${x.toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(1, bw).toFixed(1)}" height="${(band * 0.84).toFixed(1)}" fill="${s.colors?.[i] ?? s.color}"/>`);
        parts.push(T(area.x - 6, y + band * 0.42 + 4, truncate(String(categories[i] ?? ''), 20), FS.axis, TXT, 'end'));
        parts.push(`<text x="${(area.x + area.w / 2).toFixed(1)}" y="${(y + band * 0.42 + 4).toFixed(1)}" text-anchor="middle" font-size="${FS.axis}" fill="#fff" font-weight="700">${escSvg(valueLabel(s.values[i] ?? 0, s.numFmt))}</text>`);
      });
    },
  },
  histogram: {
    legend: false,
    draw(ctx) {
      const { chart, series, parts } = ctx;
      const vals = series.flatMap((s) => s.values);
      const bins = histogramBins(vals, chart.binCount ?? null, chart.binWidth ?? null);
      const cats = bins.map(binLabel);
      const { area, pos, band } = cartesian(ctx, bins.map((b) => b.count), { cats });
      const s = series[0];
      bins.forEach((b, i) => {
        const x = area.x + band * i;
        const y = pos(b.count);
        parts.push(`<rect x="${(x + 0.5).toFixed(1)}" y="${y.toFixed(1)}" width="${Math.max(1, band - 1).toFixed(1)}" height="${(pos(0) - y).toFixed(1)}" fill="${s?.color}"/>`);
        if (ctx.wantLabels(s ?? {})) labelTxt(ctx, null, b.count, x + band / 2, y - 4);
      });
    },
  },
  pareto: {
    draw(ctx) {
      const { series, categories, parts, pal, FS, TXT } = ctx;
      const s = series[0];
      if (!s) return;
      let items;
      let cats;
      const catsAreNums = categories.every((c, i) => String(c) === String(i + 1));
      if (catsAreNums) {
        const bins = histogramBins(s.values);
        items = bins.map((b, i) => ({ c: binLabel(b, i), v: b.count }));
      } else items = s.values.map((v, i) => ({ c: categories[i], v: isNum(v) ? v : 0 }));
      items.sort((a, b) => b.v - a.v);
      cats = items.map((x) => x.c);
      const total = items.reduce((a, x) => a + Math.max(0, x.v), 0) || 1;
      const { area, pos, band } = cartesian(ctx, items.map((x) => x.v), { cats, code: s.numFmt, right: FS.axis * 2.6 });
      const w = band * 0.9;
      let acc = 0;
      const pts = [];
      items.forEach((x, i) => {
        const bx = area.x + band * i + (band - w) / 2;
        const y = pos(x.v);
        parts.push(`<rect x="${bx.toFixed(1)}" y="${y.toFixed(1)}" width="${w.toFixed(1)}" height="${(pos(0) - y).toFixed(1)}" fill="${s.color}"/>`);
        acc += Math.max(0, x.v);
        pts.push([area.x + band * (i + 0.5), area.y + area.h - (acc / total) * area.h]);
      });
      parts.push(`<path d="${pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('')}" fill="none" stroke="${pal[1]}" stroke-width="2.25"/>`);
      for (const p of pts) parts.push(`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="2.5" fill="${pal[1]}"/>`);
      for (let k = 0; k <= 4; k++) parts.push(T(area.x + area.w + 4, area.y + area.h - (k / 4) * area.h + 3.5, `${k * 25}%`, FS.axis, TXT, 'start'));
    },
  },
  treemap: {
    draw(ctx) {
      const { series, categories, plot, parts, pal } = ctx;
      const s = series[0];
      if (!s) return;
      const items = s.values.map((v, i) => ({ v: isNum(v) && v > 0 ? v : 0, i })).filter((x) => x.v > 0).sort((a, b) => b.v - a.v);
      const rects = squarify(items, { x: plot.x, y: plot.y, w: plot.w, h: plot.h });
      for (const r of rects) {
        const col = s.colors?.[r.i] ?? pal[r.i % pal.length];
        parts.push(`<rect x="${r.x.toFixed(1)}" y="${r.y.toFixed(1)}" width="${Math.max(0, r.w - 1.5).toFixed(1)}" height="${Math.max(0, r.h - 1.5).toFixed(1)}" fill="${col}"/>`);
        if (r.w > 34 && r.h > 18) {
          parts.push(`<text x="${(r.x + 5).toFixed(1)}" y="${(r.y + 14).toFixed(1)}" font-size="11" fill="#fff" font-weight="700">${escSvg(truncate(String(categories[r.i] ?? ''), Math.floor((r.w - 8) / 7)))}</text>`);
          if (r.h > 32) parts.push(`<text x="${(r.x + 5).toFixed(1)}" y="${(r.y + 28).toFixed(1)}" font-size="10" fill="#fff">${escSvg(valueLabel(r.v, s.numFmt))}</text>`);
        }
      }
    },
  },
  boxWhisker: {
    draw(ctx) {
      const { chart, series, parts } = ctx;
      const stats = series.map((s) => {
        const v = s.values.filter(isNum).sort((a, b) => a - b);
        const q1 = quantile(v, 0.25);
        const q3 = quantile(v, 0.75);
        const iqr = q3 - q1;
        const inside = v.filter((x) => x >= q1 - 1.5 * iqr && x <= q3 + 1.5 * iqr);
        return { s, q1, q3, med: quantile(v, 0.5), mean: v.reduce((a, b) => a + b, 0) / (v.length || 1), lo: inside[0] ?? q1, hi: inside.at(-1) ?? q3, out: v.filter((x) => x < q1 - 1.5 * iqr || x > q3 + 1.5 * iqr) };
      });
      const all = stats.flatMap((x) => [x.lo, x.hi, ...x.out]);
      const { area, pos, band } = cartesian(ctx, all, { cats: series.map((s) => s.name), zero: false });
      stats.forEach((st, i) => {
        const cx = area.x + band * (i + 0.5);
        const w = Math.min(60, band * 0.5);
        const c = st.s.color;
        parts.push(`<line x1="${cx}" y1="${pos(st.hi).toFixed(1)}" x2="${cx}" y2="${pos(st.q3).toFixed(1)}" stroke="#595959"/>`, `<line x1="${cx}" y1="${pos(st.q1).toFixed(1)}" x2="${cx}" y2="${pos(st.lo).toFixed(1)}" stroke="#595959"/>`);
        parts.push(`<line x1="${cx - w / 4}" y1="${pos(st.hi).toFixed(1)}" x2="${cx + w / 4}" y2="${pos(st.hi).toFixed(1)}" stroke="#595959"/>`, `<line x1="${cx - w / 4}" y1="${pos(st.lo).toFixed(1)}" x2="${cx + w / 4}" y2="${pos(st.lo).toFixed(1)}" stroke="#595959"/>`);
        parts.push(`<rect x="${(cx - w / 2).toFixed(1)}" y="${pos(st.q3).toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(1, pos(st.q1) - pos(st.q3)).toFixed(1)}" fill="${c}" stroke="#595959" stroke-width="0.75"/>`);
        parts.push(`<line x1="${(cx - w / 2).toFixed(1)}" y1="${pos(st.med).toFixed(1)}" x2="${(cx + w / 2).toFixed(1)}" y2="${pos(st.med).toFixed(1)}" stroke="#fff" stroke-width="1.5"/>`);
        if (chart.showMean !== false) { const my = pos(st.mean); parts.push(`<path d="M${cx - 3.5},${my - 3.5}L${cx + 3.5},${my + 3.5}M${cx + 3.5},${my - 3.5}L${cx - 3.5},${my + 3.5}" stroke="#fff" stroke-width="1.5"/>`); }
        for (const o of st.out) parts.push(`<circle cx="${cx}" cy="${pos(o).toFixed(1)}" r="2.5" fill="none" stroke="${c}"/>`);
      });
    },
  },
  stock: {
    legend: false,
    draw(ctx) {
      const { chart, series, categories, parts } = ctx;
      const ohlc = chart.ohlc || series.length >= 4;
      const [o, h, l, c] = ohlc ? series : [null, ...series];
      if (!h || !l) return;
      const vals = [...h.values, ...l.values, ...(c?.values ?? []), ...(o?.values ?? [])].filter(isNum);
      const { area, pos, band } = cartesian(ctx, vals, { cats: categories, zero: false, code: c?.numFmt });
      categories.forEach((_, i) => {
        const cx = area.x + band * (i + 0.5);
        const hi = h.values[i];
        const lo = l.values[i];
        if (isNum(hi) && isNum(lo)) parts.push(`<line x1="${cx.toFixed(1)}" y1="${pos(hi).toFixed(1)}" x2="${cx.toFixed(1)}" y2="${pos(lo).toFixed(1)}" stroke="#404040"/>`);
        if (ohlc && isNum(o.values[i]) && isNum(c.values[i])) {
          const up = c.values[i] >= o.values[i];
          const y1 = pos(Math.max(o.values[i], c.values[i]));
          const y2 = pos(Math.min(o.values[i], c.values[i]));
          const w = Math.min(14, band * 0.5);
          parts.push(`<rect x="${(cx - w / 2).toFixed(1)}" y="${y1.toFixed(1)}" width="${w.toFixed(1)}" height="${Math.max(1, y2 - y1).toFixed(1)}" fill="${up ? (chart.upColor ?? '#ffffff') : (chart.downColor ?? '#404040')}" stroke="#404040"/>`);
        } else if (c && isNum(c.values[i])) {
          parts.push(`<line x1="${cx.toFixed(1)}" y1="${pos(c.values[i]).toFixed(1)}" x2="${(cx + Math.min(8, band * 0.3)).toFixed(1)}" y2="${pos(c.values[i]).toFixed(1)}" stroke="#404040" stroke-width="2"/>`);
        }
      });
    },
  },
  radar: {
    draw(ctx) {
      const { chart, series, categories, plot, parts, FS, TXT, GRID } = ctx;
      const n = categories.length;
      if (n < 3) return;
      const cx = plot.x + plot.w / 2;
      const cy = plot.y + plot.h / 2 + 4;
      const R = Math.max(10, Math.min(plot.w / 2 - 40, plot.h / 2 - 16));
      const vals = series.flatMap((s) => s.values).filter(isNum);
      const sc = niceScale(Math.min(0, ...vals), Math.max(...vals, 1), 4);
      const ang = (i) => -Math.PI / 2 + (i / n) * Math.PI * 2;
      const at = (i, v) => { const r = ((v - sc.min) / (sc.max - sc.min)) * R; return [cx + Math.cos(ang(i)) * r, cy + Math.sin(ang(i)) * r]; };
      for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) {
        const pts = categories.map((_, i) => at(i, t));
        parts.push(`<path d="${pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('')}Z" fill="none" stroke="${GRID}"/>`);
        parts.push(T(cx - 4, at(0, t)[1] + 3.5, axisLabel(Number(t.toPrecision(12))), FS.axis * 0.9, TXT, 'end'));
      }
      categories.forEach((c, i) => {
        const [x, y] = at(i, sc.max);
        parts.push(`<line x1="${cx}" y1="${cy}" x2="${x.toFixed(1)}" y2="${y.toFixed(1)}" stroke="${GRID}"/>`);
        const [lx, ly] = [cx + Math.cos(ang(i)) * (R + 12), cy + Math.sin(ang(i)) * (R + 12)];
        parts.push(T(lx, ly + 3.5, truncate(String(c), 12), FS.axis, TXT, Math.abs(Math.cos(ang(i))) < 0.2 ? 'middle' : Math.cos(ang(i)) > 0 ? 'start' : 'end'));
      });
      const filled = chart.radarStyle === 'filled';
      series.forEach((s) => {
        const pts = s.values.map((v, i) => at(i, isNum(v) ? v : sc.min));
        parts.push(`<path d="${pts.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('')}Z" fill="${filled ? s.color : 'none'}" fill-opacity="${filled ? 0.55 : 0}" stroke="${s.color}" stroke-width="2"/>`);
        if (chart.radarStyle === 'marker') for (const p of pts) parts.push(`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="${s.color}"/>`);
        if (ctx.wantLabels(s)) s.values.forEach((v, i) => { if (isNum(v)) { const p = at(i, v); labelTxt(ctx, s, v, p[0], p[1] - 6); } });
      });
    },
  },
  bubble: {
    draw(ctx) {
      const { series, plot, parts, FS, TXT, GRID, chart } = ctx;
      const xs = series.flatMap((s) => (s.x ?? s.values.map((_, i) => i + 1))).filter(isNum);
      const ys = series.flatMap((s) => s.values).filter(isNum);
      const sizes = series.flatMap((s) => s.size ?? []).filter(isNum);
      if (!xs.length || !ys.length) return;
      const xsc = niceScale(Math.min(...xs), Math.max(...xs));
      const ysc = niceScale(Math.min(0, ...ys), Math.max(...ys));
      const CW = FS.axis * 0.58;
      const ticksOf = (sc) => { const t = []; for (let v = sc.min; v <= sc.max + sc.step / 2; v += sc.step) t.push(Number(v.toPrecision(12))); return t; };
      const yt = ticksOf(ysc);
      const lw = Math.max(...yt.map((t) => axisLabel(t).length)) * CW + 8;
      const area = { x: plot.x + lw, y: plot.y + 4, w: plot.w - lw - 10, h: plot.h - FS.axis * 1.9 };
      const X = (v) => area.x + ((v - xsc.min) / (xsc.max - xsc.min)) * area.w;
      const Y = (v) => area.y + area.h - ((v - ysc.min) / (ysc.max - ysc.min)) * area.h;
      for (const t of yt) parts.push(chart.gridY === false ? '' : `<line x1="${area.x}" y1="${Y(t).toFixed(1)}" x2="${area.x + area.w}" y2="${Y(t).toFixed(1)}" stroke="${GRID}"/>`, T(area.x - 5, Y(t) + 3.5, axisLabel(t), FS.axis, TXT, 'end'));
      for (const t of ticksOf(xsc)) parts.push(T(X(t), area.y + area.h + FS.axis * 1.35, axisLabel(t), FS.axis, TXT));
      const maxS = Math.max(...sizes.map(Math.abs), 1);
      const maxR = Math.min(area.w, area.h) * 0.12 * ((chart.bubbleScale ?? 100) / 100);
      series.forEach((s) => s.values.forEach((v, i) => {
        const x = s.x ? s.x[i] : i + 1;
        if (!isNum(v) || !isNum(x)) return;
        const sz = Math.abs(s.size?.[i] ?? 1);
        const r = Math.max(2, Math.sqrt(sz / maxS) * maxR);
        parts.push(`<circle cx="${X(x).toFixed(1)}" cy="${Y(v).toFixed(1)}" r="${r.toFixed(1)}" fill="${s.color}" fill-opacity="0.75" stroke="${s.color}"/>`);
        if (ctx.wantLabels(s)) labelTxt(ctx, s, v, X(x), Y(v) + 3.5);
      }));
    },
  },
};

/** 스퀘어리파이 트리맵 배치 (items: [{ v, i }] 큰 값부터) */
function squarify(items, box) {
  const out = [];
  const total = items.reduce((a, x) => a + x.v, 0);
  if (!total) return out;
  const scale = (box.w * box.h) / total;
  let rest = items.map((x) => ({ ...x, a: x.v * scale }));
  let b = { ...box };
  while (rest.length) {
    const short = Math.min(b.w, b.h);
    let row = [rest[0]];
    const worst = (r) => { const s = r.reduce((a, x) => a + x.a, 0); const mx = Math.max(...r.map((x) => x.a)); const mn = Math.min(...r.map((x) => x.a)); return Math.max((short * short * mx) / (s * s), (s * s) / (short * short * mn)); };
    let k = 1;
    while (k < rest.length && worst([...row, rest[k]]) <= worst(row)) { row.push(rest[k]); k++; }
    const sum = row.reduce((a, x) => a + x.a, 0);
    if (b.w >= b.h) {
      const w = sum / b.h;
      let y = b.y;
      for (const x of row) { const h = x.a / w; out.push({ x: b.x, y, w, h, v: x.v, i: x.i }); y += h; }
      b = { x: b.x + w, y: b.y, w: b.w - w, h: b.h };
    } else {
      const h = sum / b.w;
      let x0 = b.x;
      for (const x of row) { const w = x.a / h; out.push({ x: x0, y: b.y, w, h, v: x.v, i: x.i }); x0 += w; }
      b = { x: b.x, y: b.y + h, w: b.w, h: b.h - h };
    }
    rest = rest.slice(k);
  }
  return out;
}
