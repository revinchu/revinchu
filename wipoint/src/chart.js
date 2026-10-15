// 차트 → SVG 문자열 (DOM 없음) — PowerPoint 차트와 같은 요소 · 서식
//
// chart = {
//   kind: col|bar|line|area|pie|doughnut|scatter|radar|bubble|combo,
//   grouping: clustered|stacked|percent (막대) · standard|stacked|percent (꺾은선/영역) — 예전 stacked: true 도 읽음,
//   variant: 'markers' (꺾은선) · 'smooth'|'line' (분산형) · 'marker'|'filled' (방사형) · 'explode' (원형),
//   title, titleSize(pt), cats, series: [{ name, vals, color, type(콤보: col|line|area), axis: 'secondary',
//     sizes(거품), labels(계열별 레이블), marker, markerSize, lineWidth(pt), dash, smooth, pointColors: {i: color},
//     trend: { type: linear|exp|log|poly|power|movingAvg, order, period, forward, eq, r2 }, err: { type: fixed|percent|stdDev|stdErr, val } }],
//   legend: t|b|l|r|tr|null, labels: true | { pos: outEnd|inEnd|ctr|inBase|t|b|l|r|bestFit, val, cat, ser, pct },
//   axes: { x: { title, hide, min, max, major, fmt, reverse }, y: {…}, y2: {…} }, grid: { major, minor, cat },
//   dataTable: false|true|'nokeys', palette: colorful1…4|mono1…6, style: 1…8, font: { size(pt), color },
//   bg, plotBg, fontColor, barOutline, dashGrid, shadow, gap(간격 너비 %), overlap(%), hole(%), firstAngle(°), explode(%) }
import { resolveColor } from './themes.js';

export const CHART_KINDS = [['col', '세로 막대형'], ['bar', '가로 막대형'], ['line', '꺾은선형'], ['area', '영역형'], ['pie', '원형'], ['doughnut', '도넛형'], ['scatter', '분산형'], ['radar', '방사형'], ['bubble', '거품형'], ['combo', '콤보'], ['waterfall', '폭포'], ['funnel', '깔때기형'], ['histogram', '히스토그램'], ['box', '상자 수염'], ['treemap', '트리맵'], ['sunburst', '선버스트']];
export const CHART_LABEL = Object.fromEntries(CHART_KINDS);
/** 종류별 하위 형식 (PowerPoint [차트 종류 변경] 의 위 줄) — [grouping/variant 키, 이름] */
export const CHART_SUBTYPES = {
  col: [['clustered', '묶은 세로 막대형'], ['stacked', '누적 세로 막대형'], ['percent', '100% 기준 누적 세로 막대형'], ['clustered3d', '3차원 묶은 세로 막대형'], ['stacked3d', '3차원 누적 세로 막대형'], ['percent3d', '3차원 100% 기준 누적 세로 막대형'], ['full3d', '3차원 세로 막대형']],
  bar: [['clustered', '묶은 가로 막대형'], ['stacked', '누적 가로 막대형'], ['percent', '100% 기준 누적 가로 막대형'], ['clustered3d', '3차원 묶은 가로 막대형'], ['stacked3d', '3차원 누적 가로 막대형'], ['percent3d', '3차원 100% 기준 누적 가로 막대형']],
  line: [['standard', '꺾은선형'], ['stacked', '누적 꺾은선형'], ['percent', '100% 기준 누적 꺾은선형'], ['markers', '표식이 있는 꺾은선형'], ['smooth', '부드러운 꺾은선형'], ['standard3d', '3차원 꺾은선형']],
  area: [['standard', '영역형'], ['stacked', '누적 영역형'], ['percent', '100% 기준 누적 영역형'], ['standard3d', '3차원 영역형'], ['stacked3d', '3차원 누적 영역형'], ['percent3d', '3차원 100% 기준 누적 영역형']],
  pie: [['pie', '원형'], ['explode', '쪼개진 원형'], ['pie3d', '3차원 원형'], ['explode3d', '3차원 쪼개진 원형']],
  doughnut: [['doughnut', '도넛형']],
  scatter: [['marker', '분산형'], ['smooth', '곡선 및 표식이 있는 분산형'], ['line', '직선 및 표식이 있는 분산형']],
  radar: [['standard', '방사형'], ['marker', '표식이 있는 방사형'], ['filled', '채워진 방사형']],
  bubble: [['bubble', '거품형']],
  waterfall: [['waterfall', '폭포']], funnel: [['funnel', '깔때기형']], histogram: [['histogram', '히스토그램'], ['pareto', '파레토']], box: [['box', '상자 수염']], treemap: [['treemap', '트리맵']], sunburst: [['sunburst', '선버스트']],
  combo: [['colLine', '묶은 세로 막대형 - 꺾은선형'], ['colLine2', '묶은 세로 막대형 - 보조 축의 꺾은선형'], ['areaCol', '누적 영역형 - 묶은 세로 막대형']],
};

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const f = (n) => Math.round(n * 10) / 10;
const ACC = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

// ───────────── 색 ─────────────
/** 색 변경 팔레트 (PowerPoint: 다양한 색상표 1–4, 단색 색상표 1–6) */
export const CHART_PALETTES = [
  ['colorful1', '다양한 색상표 1', [1, 2, 3, 4, 5, 6]],
  ['colorful2', '다양한 색상표 2', [2, 4, 6]],
  ['colorful3', '다양한 색상표 3', [3, 5, 1]],
  ['colorful4', '다양한 색상표 4', [6, 5, 4, 3, 2, 1]],
  ...[1, 2, 3, 4, 5, 6].map((k) => [`mono${k}`, `단색 색상표 ${k}`, k]),
];
const MONO_STEPS = ['', ':s65', ':t65', ':s50', ':t50', ':s85', ':t80'];
/** 팔레트의 i 번째 색 (테마 참조 '@accent…') */
export function paletteRef(palette, i) {
  const p = CHART_PALETTES.find((x) => x[0] === palette) ?? CHART_PALETTES[0];
  if (typeof p[2] === 'number') return `@accent${p[2]}${MONO_STEPS[i % MONO_STEPS.length]}`;
  const list = p[2];
  const base = `@accent${list[i % list.length]}`;
  const round = Math.floor(i / list.length);
  return round === 0 ? base : `${base}${round % 2 ? ':lm60' : ':lm60:lo40'}`;
}
export const paletteColor = (theme, palette, i) => resolveColor(theme, paletteRef(palette, i));
export function seriesColor(theme, chart, i) {
  const s = chart.series?.[i];
  if (s?.color) return resolveColor(theme, s.color);
  return paletteColor(theme, chart.palette, i);
}
/** 점(원형 조각 등) 색 */
function pointColor(theme, chart, s, si, i, vary) {
  if (s?.pointColors?.[i]) return resolveColor(theme, s.pointColors[i]);
  return vary ? paletteColor(theme, chart.palette, i) : seriesColor(theme, chart, si);
}

// ───────────── 숫자 · 눈금 ─────────────
/** 보기 좋은 눈금 (최소, 최대, 간격) — opts: { min, max, major } 사용자 고정, zero: 0 포함 (엑셀: 값이 0에서 멀면 0 을 빼지 않음) */
export function niceScale(min, max, ticks = 5, opts = {}) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) { min = 0; max = 1; }
  const zero = opts.zero ?? true;
  if (zero) {
    if (min > 0 && !(max > 0 && (max - min) / max < 1 / 6)) min = 0;
    if (max < 0 && !(min < 0 && (min - max) / min < 1 / 6)) max = 0;
  }
  if (opts.min != null && Number.isFinite(+opts.min)) min = +opts.min;
  if (opts.max != null && Number.isFinite(+opts.max)) max = +opts.max;
  if (min === max) max = min + 1;
  if (min > max) [min, max] = [max, min];
  let step = Number(opts.major);
  if (!(step > 0)) {
    const raw = (max - min) / ticks;
    const mag = 10 ** Math.floor(Math.log10(raw));
    step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw * 0.999) ?? 10 * mag;
  }
  const lo = opts.min != null && Number.isFinite(+opts.min) ? +opts.min : Math.floor(min / step + 1e-9) * step;
  const hi = opts.max != null && Number.isFinite(+opts.max) ? +opts.max : Math.ceil(max / step - 1e-9) * step;
  return { lo, hi: hi > lo ? hi : lo + step, step };
}
/** 숫자 표시 형식: General | 0 | 0.0 | #,##0 | #,##0.00 | 0% | 0.0% | ₩#,##0 | $#,##0 | #,##0"원" */
export function fmtNum(v, code) {
  if (v == null || v === '') return '';
  const n = Number(v);
  if (!Number.isFinite(n)) return String(v);
  if (!code || code === 'General') {
    const a = Math.abs(n);
    if (a >= 1000) return Number(n.toFixed(2)).toLocaleString('ko-KR', { maximumFractionDigits: 2 });
    if (a > 0 && a < 0.001) return n.toExponential(1);
    return String(Number(n.toPrecision(10)) === Math.round(n * 1e6) / 1e6 ? Math.round(n * 1e6) / 1e6 : Number(n.toPrecision(6)));
  }
  const m = /^([^#0]*)([#,0]*0)(?:\.(0+))?(%?)(.*)$/.exec(code.replace(/"/g, ''));
  if (!m) return String(n);
  const [, pre, int, dec = '', pct, post] = m;
  const x = pct ? n * 100 : n;
  const s = int.includes(',') ? x.toLocaleString('ko-KR', { minimumFractionDigits: dec.length, maximumFractionDigits: dec.length }) : x.toFixed(dec.length);
  return `${pre}${s}${pct}${post}`;
}
/** 글 너비 어림 (한글 · 한자 1em, 그 밖 0.55em) */
export function textWidth(s, fs) {
  let w = 0;
  for (const ch of String(s ?? '')) w += /[ᄀ-ᇿ　-鿿가-힯＀-￯]/.test(ch) ? 1 : /[A-Z0-9%$₩@#&]/.test(ch) ? 0.62 : 0.52;
  return w * fs;
}

// ───────────── 추세선 ─────────────
/** 추세선 식 맞추기 → { f(x), eq, r2 } (엑셀 방식: 지수 · 거듭제곱은 로그 변환한 값으로 R²) */
export function trendFit(xs, ys, type = 'linear', order = 2) {
  const pts = xs.map((x, i) => [Number(x), Number(ys[i])]).filter(([x, y]) => Number.isFinite(x) && Number.isFinite(y));
  const tx = (x) => (type === 'log' || type === 'power' ? Math.log(x) : x);
  const ty = (y) => (type === 'exp' || type === 'power' ? Math.log(y) : y);
  const use = pts.filter(([x, y]) => (type === 'log' || type === 'power' ? x > 0 : true) && (type === 'exp' || type === 'power' ? y > 0 : true)).map(([x, y]) => [tx(x), ty(y)]);
  if (use.length < 2) return null;
  const deg = type === 'poly' ? Math.max(2, Math.min(6, order | 0)) : 1;
  const coef = polyFit(use, deg);
  if (!coef) return null;
  const p = (x) => coef.reduce((s, c, k) => s + c * x ** k, 0);
  const mean = use.reduce((s, [, y]) => s + y, 0) / use.length;
  const ssTot = use.reduce((s, [, y]) => s + (y - mean) ** 2, 0);
  const ssRes = use.reduce((s, [x, y]) => s + (y - p(x)) ** 2, 0);
  const r2 = ssTot > 0 ? 1 - ssRes / ssTot : 1;
  const r = (v) => Number(v.toPrecision(4));
  const sgn = (v) => (v < 0 ? ` - ${r(-v)}` : ` + ${r(v)}`);
  let fn;
  let eq;
  if (type === 'exp') { const a = Math.exp(coef[0]); const b = coef[1]; fn = (x) => a * Math.exp(b * x); eq = `y = ${r(a)}e^${r(b)}x`; }
  else if (type === 'power') { const a = Math.exp(coef[0]); const b = coef[1]; fn = (x) => a * x ** b; eq = `y = ${r(a)}x^${r(b)}`; }
  else if (type === 'log') { fn = (x) => coef[1] * Math.log(x) + coef[0]; eq = `y = ${r(coef[1])}ln(x)${sgn(coef[0])}`; }
  else {
    fn = p;
    const big = Math.max(...coef.map(Math.abs));
    const terms = coef.map((c, k) => [Math.abs(c) < big * 1e-9 ? 0 : c, k]).reverse().filter(([c]) => c !== 0);
    const coefTxt = (c, k) => (k > 0 && Math.abs(c) === 1 ? (c < 0 ? '-' : '') : String(r(c)));
    eq = `y = ${terms.map(([c, k], i) => { const term = k === 0 ? '' : k === 1 ? 'x' : `x${'⁰¹²³⁴⁵⁶'[k]}`; return i === 0 ? `${coefTxt(c, k)}${term}` : `${c < 0 ? ' - ' : ' + '}${k > 0 && Math.abs(c) === 1 ? '' : r(Math.abs(c))}${term}`; }).join('') || '0'}`;
  }
  return { f: fn, eq, r2 };
}
function polyFit(pts, deg) {
  const n = deg + 1;
  const A = Array.from({ length: n }, () => new Array(n + 1).fill(0));
  for (const [x, y] of pts) {
    for (let i = 0; i < n; i++) {
      for (let j = 0; j < n; j++) A[i][j] += x ** (i + j);
      A[i][n] += y * x ** i;
    }
  }
  for (let c = 0; c < n; c++) {
    let piv = c;
    for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[piv][c])) piv = r;
    if (Math.abs(A[piv][c]) < 1e-12) return null;
    [A[c], A[piv]] = [A[piv], A[c]];
    for (let r = 0; r < n; r++) {
      if (r === c) continue;
      const k = A[r][c] / A[c][c];
      for (let j = c; j <= n; j++) A[r][j] -= k * A[c][j];
    }
  }
  return A.map((row, i) => row[n] / row[i]);
}
export const TREND_LABEL = { linear: '선형', exp: '지수', log: '로그', poly: '다항식', power: '거듭제곱', movingAvg: '이동 평균' };

// ───────────── 계열 정리 ─────────────
const MAIN_TYPE = { col: 'col', bar: 'bar', line: 'line', area: 'area', scatter: 'scatter', bubble: 'bubble', radar: 'radar', combo: 'col' };
/** 계열마다 그리는 종류 */
export function seriesType(chart, s, i) {
  if (chart.kind === 'combo') return s.type ?? (i === chart.series.length - 1 && chart.series.length > 1 ? 'line' : 'col');
  return MAIN_TYPE[chart.kind] ?? 'col';
}
/** 묶음 방식 */
export function groupingOf(chart, type) {
  const g = chart.grouping ?? (chart.stacked ? 'stacked' : null);
  if (type === 'col' || type === 'bar') return g === 'stacked' || g === 'percent' ? g : 'clustered';
  if (type === 'line' || type === 'area') {
    if (chart.kind === 'combo') return type === 'area' && (g === 'stacked' || g === 'percent') ? g : 'standard';
    return g === 'stacked' || g === 'percent' ? g : 'standard';
  }
  return 'standard';
}
/** 레이블 설정 (계열 → 차트 순, true = 값) */
export function labelOpts(chart, s, kind) {
  const v = s?.labels ?? chart.labels;
  if (!v) return null;
  const pie = kind === 'pie' || kind === 'doughnut';
  const o = v === true ? {} : v;
  const any = o.val || o.cat || o.ser || o.pct;
  return { pos: o.pos ?? null, val: any ? !!o.val : !pie, cat: !!o.cat, ser: !!o.ser, pct: any ? !!o.pct : pie };
}
const numVal = (v) => { const n = Number(v); return Number.isFinite(n) ? n : null; };

// ───────────── 그리기 ─────────────
export function chartSvg(chart, w, h, theme, { font = '맑은 고딕' } = {}) {
  const c = expandChart(chart ?? {});
  const fc = c.fontColor ? resolveColor(theme, c.fontColor) : c.font?.color ? resolveColor(theme, c.font.color) : resolveColor(theme, '@tx1:lm65:lo35') ?? '#595959';
  const gridC = c.gridColor ? resolveColor(theme, c.gridColor) : resolveColor(theme, '@tx1:lm15:lo85') ?? '#D9D9D9';
  const axisC = resolveColor(theme, '@tx1:lm25:lo75') ?? '#BFBFBF';
  // 파일에서 읽은 글자 크기는 그대로 (데이터 레이블 · 범례 등은 fs × 0.9 로 그리므로 나눠 둠)
  const fs = c.font?.size ? (c.font.size * (4 / 3)) / (c.font.exact ? 0.9 : 1) : Math.max(9, Math.min(16, Math.min(w, h) / 22));
  const parts = [];
  const defs = [];
  const P = (s) => parts.push(s);
  if (c.bg) P(`<rect x="0" y="0" width="${f(w)}" height="${f(h)}" fill="${resolveColor(theme, c.bg)}"${c.border ? ` stroke="${resolveColor(theme, c.border)}"` : ''}/>`);
  else if (c.border) P(`<rect x="0.5" y="0.5" width="${f(w - 1)}" height="${f(h - 1)}" fill="none" stroke="${resolveColor(theme, c.border)}"/>`);
  if (c.shadow) defs.push('<filter id="wpsh" x="-10%" y="-10%" width="130%" height="130%"><feDropShadow dx="1.5" dy="2" stdDeviation="1.5" flood-opacity="0.35"/></filter>');
  const shadow = c.shadow ? ' filter="url(#wpsh)"' : '';
  // 축별 글자 크기 (파일의 축 txPr): 축 눈금 이름에만
  const axSize = (el) => { const k = /^axis-(x|y2|y)$/.exec(el)?.[1]; const pt = k && c.axes?.[k]?.size; return pt ? pt * (4 / 3) : null; };
  const text = (x, y, s, { size = fs, anchor = 'middle', color = fc, bold = false, rot = 0, el = '' } = {}) => `<text x="${f(x)}" y="${f(y)}" font-size="${f(axSize(el) ?? size)}" text-anchor="${anchor}" fill="${color}"${bold ? ' font-weight="bold"' : ''}${rot ? ` transform="rotate(${rot} ${f(x)} ${f(y)})"` : ''}${el ? ` data-el="${el}"` : ''}>${esc(s)}</text>`;

  let top = 6;
  let bottom = h - 6;
  let left = 8;
  let right = w - 8;
  // 제목
  if (c.title) {
    const ts = c.titleSize ? c.titleSize * (4 / 3) : fs * 1.4;
    P(text(w / 2, top + ts * 1.05, c.title, { size: ts, el: 'title', bold: !!c.titleBold }));
    if (!c.titleOverlay) top += ts * 1.6;
  }
  const kind = c.kind ?? 'col';
  const series = (c.series ?? []).filter(Boolean);
  const cats = c.cats ?? [];
  const pie = kind === 'pie' || kind === 'doughnut';
  // 범례 항목 (+ 추세선)
  const types = series.map((s, i) => seriesType(c, s, i));
  const legendItems = pie
    ? cats.map((ct, i) => ({ name: ct, color: pointColor(theme, c, series[0], 0, i, true), kind: 'rect' }))
    : SPECIAL.has(kind) ? specialLegend(c, theme)
    : [...series.map((s, i) => ({ ghost: s.ghost, name: s.name, color: seriesColor(theme, c, i), kind: types[i] === 'line' || types[i] === 'scatter' || (types[i] === 'radar' && c.variant !== 'filled') ? 'line' : 'rect', marker: types[i] === 'line' ? lineMarker(c, s) : null })).filter((it) => !it.ghost),
      ...series.flatMap((s, i) => (s.trend?.type ? [{ name: s.trend.name || `${TREND_LABEL[s.trend.type] ?? ''} (${s.name})`, color: seriesColor(theme, c, i), kind: 'dash' }] : []))];
  const lfs = fs * 0.95;
  const legendKey = (x, y, it) => {
    if (it.kind === 'rect') return `<rect x="${f(x)}" y="${f(y - lfs * 0.62)}" width="${f(lfs * 0.62)}" height="${f(lfs * 0.62)}" fill="${it.color}"/>`;
    const l = `<line x1="${f(x - lfs * 0.2)}" y1="${f(y - lfs * 0.31)}" x2="${f(x + lfs * 0.9)}" y2="${f(y - lfs * 0.31)}" stroke="${it.color}" stroke-width="${f(Math.max(1.5, lfs / 7))}"${it.kind === 'dash' ? ' stroke-dasharray="3 2"' : ''}/>`;
    return it.marker ? l + markerSvg(it.marker, x + lfs * 0.35, y - lfs * 0.31, lfs * 0.28, it.color) : l;
  };
  const legendPos = c.legend ?? null;
  if (legendPos && legendItems.length) {
    const itemW = legendItems.map((it) => lfs * 1.5 + textWidth(it.name, lfs) + lfs * 0.8);
    if (legendPos === 'r' || legendPos === 'l' || legendPos === 'tr') {
      const lw = Math.min(w * 0.35, Math.max(...itemW));
      const lh = legendItems.length * lfs * 1.5;
      const x = legendPos === 'l' ? left : w - 8 - lw;
      const y0 = legendPos === 'tr' ? top + 4 : Math.max(top, (top + bottom) / 2 - lh / 2);
      P(`<g data-el="legend">${legendItems.map((it, i) => { const y = y0 + (i + 0.8) * lfs * 1.5; return legendKey(x + 2, y, it) + text(x + lfs * 1.3, y, it.name, { size: lfs, anchor: 'start' }); }).join('')}</g>`);
      if (legendPos === 'l') left += lw + 6; else right -= lw + 6;
    } else {
      // 위 · 아래: 줄바꿈
      const rows = [[]];
      let rowW = 0;
      itemW.forEach((iw, i) => { if (rowW + iw > w - 16 && rows[rows.length - 1].length) { rows.push([]); rowW = 0; } rows[rows.length - 1].push(i); rowW += iw; });
      const lh = rows.length * lfs * 1.5;
      const yb = legendPos === 't' ? top : bottom - lh;
      const g = [];
      rows.forEach((r, ri) => {
        const tw = r.reduce((s, i) => s + itemW[i], 0);
        let x = (w - tw) / 2;
        const y = yb + (ri + 0.85) * lfs * 1.5;
        for (const i of r) { g.push(legendKey(x, y, legendItems[i]) + text(x + lfs * 1.1, y, legendItems[i].name, { size: lfs, anchor: 'start' })); x += itemW[i]; }
      });
      P(`<g data-el="legend">${g.join('')}</g>`);
      if (legendPos === 't') top += lh + 4; else bottom -= lh + 2;
    }
  }
  const box = { left, right, top, bottom };
  let body;
  if (is3d(c) && kind === 'pie') body = pie3dSvg(c, series, cats, box, theme, { fs, fc, text });
  else if (is3d(c)) body = cartesian3dSvg(c, series, cats, box, theme, { fs, fc, text, gridC, axisC });
  else if (pie) body = pieSvg(c, series, cats, box, theme, { fs, fc, text, shadow });
  else if (kind === 'box') body = boxPlotSvg(c, series, box, theme, { fs, fc, text, gridC, axisC });
  else if (kind === 'treemap') body = treemapSvg(c, box, theme, { fs, text });
  else if (kind === 'sunburst') body = sunburstSvg(c, box, theme, { fs, text });
  else if (kind === 'radar') body = radarSvg(c, series, cats, box, theme, { fs, fc, text, gridC, axisC });
  else body = cartesianSvg(c, series, cats, types, box, theme, { fs, fc, text, gridC, axisC, shadow, legendKey, w, h });
  return wrap(w, h, font, [...parts, body], defs);
}

function lineMarker(c, s) {
  if (s.marker) return s.marker === 'none' ? null : s.marker;
  return c.variant === 'markers' || c.markers ? 'circle' : null;
}
function markerSvg(kind, x, y, r, color, fill = color) {
  if (kind === 'square') return `<rect x="${f(x - r)}" y="${f(y - r)}" width="${f(r * 2)}" height="${f(r * 2)}" fill="${fill}" stroke="${color}"/>`;
  if (kind === 'diamond') return `<path d="M${f(x)} ${f(y - r * 1.2)} L${f(x + r * 1.2)} ${f(y)} L${f(x)} ${f(y + r * 1.2)} L${f(x - r * 1.2)} ${f(y)} Z" fill="${fill}" stroke="${color}"/>`;
  if (kind === 'triangle') return `<path d="M${f(x)} ${f(y - r * 1.2)} L${f(x + r * 1.1)} ${f(y + r * 0.8)} L${f(x - r * 1.1)} ${f(y + r * 0.8)} Z" fill="${fill}" stroke="${color}"/>`;
  if (kind === 'x') return `<path d="M${f(x - r)} ${f(y - r)} L${f(x + r)} ${f(y + r)} M${f(x + r)} ${f(y - r)} L${f(x - r)} ${f(y + r)}" stroke="${color}" stroke-width="1.5"/>`;
  return `<circle cx="${f(x)}" cy="${f(y)}" r="${f(r)}" fill="${fill}" stroke="${color}"/>`;
}
function smoothD(pts) {
  if (pts.length < 3) return pts.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' ');
  let d = `M${f(pts[0][0])} ${f(pts[0][1])}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[i - 1] ?? pts[i];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[i + 2] ?? p2;
    d += ` C${f(p1[0] + (p2[0] - p0[0]) / 6)} ${f(p1[1] + (p2[1] - p0[1]) / 6)} ${f(p2[0] - (p3[0] - p1[0]) / 6)} ${f(p2[1] - (p3[1] - p1[1]) / 6)} ${f(p2[0])} ${f(p2[1])}`;
  }
  return d;
}
const lineD = (pts, smooth) => (smooth ? smoothD(pts) : pts.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' '));
const DASH = { dash: '6 3', dot: '1.5 2.5', dashDot: '6 3 1.5 3', longDash: '10 4' };

function labelText(lo, v, s, cat, pct, fmt) {
  const out = [];
  if (lo.ser) out.push(s.name);
  if (lo.cat) out.push(cat);
  if (lo.val) out.push(fmtNum(v, lo.fmt ?? fmt));
  if (lo.pct && pct != null) out.push(`${Math.round(pct * 100)}%`);
  return out.join(', ');
}

// ───────────── 원형 · 도넛형 ─────────────
function pieSvg(c, series, cats, box, theme, { fs, fc, text, shadow }) {
  const out = [];
  const cx = (box.left + box.right) / 2;
  const cy = (box.top + box.bottom) / 2;
  const lo = labelOpts(c, series[0], c.kind);
  const outside = lo && (lo.pos === 'outEnd' || lo.pos === 'bestFit' || lo.cat);
  const R = Math.max(10, Math.min(box.right - box.left - (outside ? fs * 8 : 8), box.bottom - box.top - (outside ? fs * 3.2 : 8)) / 2);
  const doughnut = c.kind === 'doughnut';
  const rings = doughnut ? series : series.slice(0, 1);
  const hole = doughnut ? Math.max(10, Math.min(90, c.hole ?? 50)) / 100 : 0;
  const explodeAll = (c.variant === 'explode' ? 18 : c.explode ?? 0) / 100;
  const a0start = ((c.firstAngle ?? 0) - 90) * (Math.PI / 180);
  rings.forEach((s, si) => {
    const vals = (s.vals ?? []).map((v) => Math.max(0, numVal(v) ?? 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const rOut = doughnut ? R - (si * (R * (1 - hole))) / rings.length : R;
    const rIn = doughnut ? rOut - (R * (1 - hole)) / rings.length : 0;
    const slo = labelOpts(c, s, c.kind);
    let a0 = a0start;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2;
      const am = (a0 + a1) / 2;
      const ex = (s.explodeAt?.[i] ?? 0) / 100 + explodeAll;
      const ox = Math.cos(am) * rOut * ex;
      const oy = Math.sin(am) * rOut * ex;
      const col = pointColor(theme, c, s, si, i, true);
      const X = (r, a) => f(cx + ox + r * Math.cos(a));
      const Y = (r, a) => f(cy + oy + r * Math.sin(a));
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const stroke = resolveColor(theme, c.bg ?? '@bg1') ?? '#fff';
      if (v / total >= 0.9999) {
        out.push(doughnut ? `<path d="M${X(rOut, 0)} ${Y(rOut, 0)} A${f(rOut)} ${f(rOut)} 0 1 1 ${X(rOut, Math.PI)} ${Y(rOut, Math.PI)} A${f(rOut)} ${f(rOut)} 0 1 1 ${X(rOut, 0)} ${Y(rOut, 0)} M${X(rIn, 0)} ${Y(rIn, 0)} A${f(rIn)} ${f(rIn)} 0 1 0 ${X(rIn, Math.PI)} ${Y(rIn, Math.PI)} A${f(rIn)} ${f(rIn)} 0 1 0 ${X(rIn, 0)} ${Y(rIn, 0)} Z" fill="${col}" fill-rule="evenodd" data-el="ser${si}"${shadow}/>` : `<circle cx="${f(cx)}" cy="${f(cy)}" r="${f(rOut)}" fill="${col}" stroke="${stroke}" data-el="ser${si}"${shadow}/>`);
      } else if (v > 0) {
        const d = doughnut
          ? `M${X(rOut, a0)} ${Y(rOut, a0)} A${f(rOut)} ${f(rOut)} 0 ${large} 1 ${X(rOut, a1)} ${Y(rOut, a1)} L${X(rIn, a1)} ${Y(rIn, a1)} A${f(rIn)} ${f(rIn)} 0 ${large} 0 ${X(rIn, a0)} ${Y(rIn, a0)} Z`
          : `M${f(cx + ox)} ${f(cy + oy)} L${X(rOut, a0)} ${Y(rOut, a0)} A${f(rOut)} ${f(rOut)} 0 ${large} 1 ${X(rOut, a1)} ${Y(rOut, a1)} Z`;
        out.push(`<path d="${d}" fill="${col}" stroke="${stroke}" stroke-width="1.5" data-el="ser${si}"${shadow}/>`);
      }
      if (slo && v > 0) {
        const t = labelText(slo, v, s, cats[i], v / total, c.axes?.y?.fmt);
        const pos = slo.pos ?? (doughnut ? 'ctr' : 'bestFit');
        const out2 = pos === 'outEnd' || (pos === 'bestFit' && (v / total < 0.08 || slo.cat));
        if (out2 && !doughnut) {
          const lx = cx + ox + (rOut + fs * 0.8) * Math.cos(am);
          const ly = cy + oy + (rOut + fs * 0.8) * Math.sin(am);
          out.push(`<line x1="${X(rOut, am)}" y1="${Y(rOut, am)}" x2="${f(lx)}" y2="${f(ly)}" stroke="${fc}" stroke-width="0.6" stroke-opacity="0.6"/>`);
          out.push(text(lx + (Math.cos(am) >= 0 ? 3 : -3), ly + fs / 3, t, { anchor: Math.cos(am) >= 0 ? 'start' : 'end', size: fs * 0.9 }));
        } else {
          const rr = doughnut ? (rOut + rIn) / 2 : pos === 'inEnd' ? rOut * 0.8 : pos === 'ctr' ? rOut * 0.5 : rOut * 0.64;
          out.push(text(cx + ox + rr * Math.cos(am), cy + oy + rr * Math.sin(am) + fs / 3, t, { size: fs * 0.9, color: '#fff' }));
        }
      }
      a0 = a1;
    });
  });
  return out.join('');
}

// ───────────── 방사형 ─────────────
function radarSvg(c, series, cats, box, theme, { fs, fc, text, gridC }) {
  const out = [];
  const n = Math.max(3, cats.length);
  const cx = (box.left + box.right) / 2;
  const cy = (box.top + box.bottom) / 2 + fs * 0.3;
  const R = Math.max(10, Math.min(box.right - box.left - fs * 8, box.bottom - box.top - fs * 3.5) / 2);
  let hi = -Infinity;
  let lo = Infinity;
  for (const s of series) for (const v of s.vals ?? []) { const x = numVal(v); if (x != null) { hi = Math.max(hi, x); lo = Math.min(lo, x); } }
  const sc = niceScale(Math.min(0, lo), hi, 5, c.axes?.y ?? {});
  const ang = (i) => -Math.PI / 2 + (i / n) * Math.PI * 2;
  const pos = (v, i) => { const r = ((v - sc.lo) / (sc.hi - sc.lo)) * R; return [cx + r * Math.cos(ang(i)), cy + r * Math.sin(ang(i))]; };
  for (let v = sc.lo; v <= sc.hi + sc.step / 2; v += sc.step) {
    const pts = Array.from({ length: n }, (_, i) => pos(v, i));
    out.push(`<path d="${pts.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' ')} Z" fill="none" stroke="${gridC}"/>`);
    if (!c.axes?.y?.hide) out.push(text(cx - 4, pos(v, 0)[1] + fs / 3, fmtNum(v, c.axes?.y?.fmt), { anchor: 'end', size: fs * 0.85 }));
  }
  for (let i = 0; i < n; i++) {
    const [x, y] = pos(sc.hi, i);
    out.push(`<line x1="${f(cx)}" y1="${f(cy)}" x2="${f(x)}" y2="${f(y)}" stroke="${gridC}"/>`);
    const a = ang(i);
    out.push(text(cx + (R + fs * 0.9) * Math.cos(a), cy + (R + fs * 0.9) * Math.sin(a) + fs / 3, cats[i] ?? '', { anchor: Math.abs(Math.cos(a)) < 0.2 ? 'middle' : Math.cos(a) > 0 ? 'start' : 'end', size: fs * 0.9 }));
  }
  series.forEach((s, si) => {
    const col = seriesColor(theme, c, si);
    const pts = Array.from({ length: n }, (_, i) => pos(numVal(s.vals?.[i]) ?? 0, i));
    const d = `${pts.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' ')} Z`;
    if (c.variant === 'filled') out.push(`<path d="${d}" fill="${col}" fill-opacity="0.55" stroke="${col}" data-el="ser${si}"/>`);
    else out.push(`<path d="${d}" fill="none" stroke="${col}" stroke-width="${f((s.lineWidth ?? 2.25) * 4 / 3)}" stroke-linejoin="round" data-el="ser${si}"/>`);
    if (c.variant === 'marker') for (const [x, y] of pts) out.push(markerSvg(s.marker && s.marker !== 'none' ? s.marker : 'circle', x, y, fs * 0.28, col));
    const lo2 = labelOpts(c, s, 'radar');
    if (lo2) pts.forEach(([x, y], i) => out.push(text(x, y - 5, labelText(lo2, s.vals?.[i], s, cats[i], null, c.axes?.y?.fmt), { size: fs * 0.85 })));
  });
  return out.join('');
}

// ───────────── 축이 있는 차트 (막대 · 꺾은선 · 영역 · 분산 · 거품 · 콤보) ─────────────
function cartesianSvg(c, series, cats, types, box, theme, ctx) {
  const { fs, fc, text, gridC, axisC, shadow, legendKey } = ctx;
  const out = [];
  const horiz = c.kind === 'bar';
  const xy = c.kind === 'scatter' || c.kind === 'bubble';
  const ax = c.axes ?? {};
  const grid = c.grid ?? { major: true };
  const n = Math.max(1, cats.length);
  // 범주 x 값 (분산형 · 거품형: 숫자)
  const xs = cats.map((v, i) => (xy ? (numVal(v) ?? i + 1) : i + 1));
  // 쌓기 · 100%
  const plotted = series.map(() => []); // [ci] → { a, b, v }
  const groups = new Map();
  series.forEach((s, si) => {
    const t = types[si];
    const g = groupingOf(c, t);
    const key = `${s.axis === 'secondary' ? 2 : 1}|${t === 'bar' ? 'col' : t}|${g}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(si);
  });
  for (const [key, idx] of groups) {
    const g = key.split('|')[2];
    for (let ci = 0; ci < n; ci++) {
      let pos = 0;
      let neg = 0;
      const tot = g === 'percent' ? idx.reduce((t, si) => t + Math.abs(numVal(series[si].vals?.[ci]) ?? 0), 0) || 1 : 1;
      for (const si of idx) {
        const raw = numVal(series[si].vals?.[ci]);
        if (raw == null) { plotted[si][ci] = null; continue; }
        const v = g === 'percent' ? raw / tot : raw;
        if (g === 'stacked' || g === 'percent') {
          const base = v >= 0 || types[si] === 'line' || types[si] === 'area' ? pos : neg;
          plotted[si][ci] = { a: base, b: base + v, v: raw, p: g === 'percent' ? v : null };
          if (v >= 0 || types[si] === 'line' || types[si] === 'area') pos += v; else neg += v;
        } else plotted[si][ci] = { a: 0, b: v, v: raw };
      }
    }
  }
  const percent = (axis) => [...groups.keys()].some((k) => k.startsWith(`${axis}|`) && k.endsWith('|percent'));
  const range = (axis) => {
    let lo = Infinity;
    let hi = -Infinity;
    series.forEach((s, si) => {
      if ((s.axis === 'secondary' ? 2 : 1) !== axis) return;
      for (const p of plotted[si]) if (p) { lo = Math.min(lo, p.a, p.b); hi = Math.max(hi, p.a, p.b); }
      if (s.err) for (const p of plotted[si]) if (p) { const e = errAmount(s, p.v); hi = Math.max(hi, p.b + e); lo = Math.min(lo, p.b - e); }
      if (c.kind === 'bubble' && Number.isFinite(hi)) { const pad = (hi - Math.min(lo, hi) || 1) * 0.15; hi += pad; }
    });
    if (types.every((t) => t === 'line' || t === 'scatter' || t === 'bubble')) { /* 0 포함은 niceScale 이 판단 */ }
    return [lo, hi];
  };
  const has2 = series.some((s) => s.axis === 'secondary');
  const [lo1, hi1] = range(1);
  const yOpts = { ...(ax.y ?? {}) };
  const sc1 = niceScale(Number.isFinite(lo1) ? lo1 : 0, Number.isFinite(hi1) ? hi1 : 1, 5, yOpts);
  if (percent(1) && ax.y?.max == null) { sc1.hi = Math.max(1, Math.ceil(sc1.hi)); sc1.lo = Math.min(0, Math.floor(sc1.lo)); sc1.step = ax.y?.major ?? (sc1.hi - sc1.lo > 1 ? 0.2 : 0.1); }
  const [lo2, hi2] = has2 ? range(2) : [0, 1];
  const sc2 = has2 ? niceScale(Number.isFinite(lo2) ? lo2 : 0, Number.isFinite(hi2) ? hi2 : 1, 5, ax.y2 ?? {}) : null;
  const fmt1 = ax.y?.fmt ?? (percent(1) ? '0%' : null);
  // 데이터 레이블은 값 자체의 서식 (100% 누적이어도 원래 값, 축의 자동 % 서식은 쓰지 않음)
  const lfmt1 = ax.y?.fmt ?? null;
  const fmt2 = ax.y2?.fmt ?? (has2 && percent(2) ? '0%' : null);
  // x 값 축 (분산형)
  const xpad = c.kind === 'bubble' ? (Math.max(...xs) - Math.min(...xs) || 1) * 0.15 : 0;
  const xsc = xy ? niceScale(Math.min(...xs) - xpad, Math.max(...xs) + xpad, 5, ax.x ?? {}) : null;
  // 거품 크기
  const maxSize = c.kind === 'bubble' ? Math.max(1e-9, ...series.flatMap((s) => (s.sizes ?? []).map((v) => Math.abs(numVal(v) ?? 0)))) : 1;
  // 여백 계산
  const valLabels = (sc, fmt) => { const a = []; for (let v = sc.lo; v <= sc.hi + sc.step / 2; v += sc.step) a.push(fmtNum(Math.round(v / sc.step) * sc.step, fmt)); return a; };
  const yHide = ax.y?.hide;
  const xHide = ax.x?.hide;
  const catLabels = xy ? valLabels(xsc, ax.x?.fmt) : cats.map((v) => String(v ?? ''));
  const vl1 = valLabels(sc1, fmt1);
  const titleY = ax.y?.title ? fs * 1.5 : 0;
  const titleX = ax.x?.title ? fs * 1.5 : 0;
  const titleY2 = has2 && ax.y2?.title ? fs * 1.5 : 0;
  const dt = c.dataTable && !xy && !horiz;
  const dtRows = dt ? series.length + 1 : 0;
  const dtNameW = dt ? Math.max(...series.map((s) => textWidth(s.name, fs * 0.9))) + (c.dataTable === 'nokeys' ? 10 : fs * 1.6) : 0;
  let x0 = box.left + titleY + (horiz ? (xHide ? 4 : Math.min((box.right - box.left) * 0.35, Math.max(...catLabels.map((s) => textWidth(s, fs))) + 8)) : yHide ? 4 : Math.max(...vl1.map((s) => textWidth(s, fs))) + 8);
  if (dt) x0 = Math.max(x0, box.left + dtNameW + 4);
  let x1 = box.right - 6 - titleY2 - (has2 && !ax.y2?.hide ? Math.max(...valLabels(sc2, fmt2).map((s) => textWidth(s, fs))) + 8 : 0);
  const y0 = box.top + 4;
  // 범주 레이블 기울이기 (자리가 모자라면)
  let rot = 0;
  let step = 1;
  if (!horiz && !xy && !xHide && !dt) {
    const band0 = (x1 - x0) / n;
    const widest = Math.max(...catLabels.map((s) => textWidth(s, fs)));
    if (widest > band0 * 0.95) { rot = -45; const slots = Math.floor((x1 - x0) / (fs * 1.2)); step = Math.max(1, Math.ceil(n / Math.max(1, slots))); }
  }
  const catH = horiz ? (yHide ? 4 : fs * 1.6) : xHide ? 4 : rot ? Math.min((box.bottom - box.top) * 0.4, Math.max(...catLabels.map((s) => textWidth(s, fs))) * 0.72 + fs) : fs * 1.6;
  const y1 = box.bottom - titleX - (dt ? dtRows * fs * 1.5 : catH);
  if (y1 - y0 < 10 || x1 - x0 < 10) return '';
  // 그림 영역
  if (c.plotBg) out.push(`<rect x="${f(x0)}" y="${f(y0)}" width="${f(x1 - x0)}" height="${f(y1 - y0)}" fill="${resolveColor(theme, c.plotBg)}" data-el="plot"/>`);
  const vpos = (v, sc) => (horiz ? x0 + ((v - sc.lo) / (sc.hi - sc.lo)) * (x1 - x0) : y1 - ((v - sc.lo) / (sc.hi - sc.lo)) * (y1 - y0));
  const catRev = !!ax.x?.reverse;
  const band = (horiz ? y1 - y0 : x1 - x0) / n;
  const catMid = (ci) => { const k = catRev ? n - 1 - ci : ci; return horiz ? y1 - band * (k + 0.5) : x0 + band * (k + 0.5); };
  const xpos = (x) => x0 + ((x - xsc.lo) / (xsc.hi - xsc.lo)) * (x1 - x0);
  const dash = c.dashGrid ? ' stroke-dasharray="4 3"' : '';
  // 눈금선
  const gridLines = (sc, minor) => {
    const st = minor ? sc.step / 5 : sc.step;
    for (let v = sc.lo; v <= sc.hi + st / 2; v += st) {
      const p = vpos(v, sc);
      if (horiz) out.push(`<line x1="${f(p)}" y1="${f(y0)}" x2="${f(p)}" y2="${f(y1)}" stroke="${gridC}" stroke-width="${minor ? 0.5 : 1}"${dash}/>`);
      else out.push(`<line x1="${f(x0)}" y1="${f(p)}" x2="${f(x1)}" y2="${f(p)}" stroke="${gridC}" stroke-width="${minor ? 0.5 : 1}"${dash}/>`);
    }
  };
  if (grid.minor) gridLines(sc1, true);
  if (grid.major !== false) gridLines(sc1, false);
  if (grid.cat) {
    if (xy) for (let v = xsc.lo; v <= xsc.hi + xsc.step / 2; v += xsc.step) out.push(`<line x1="${f(xpos(v))}" y1="${f(y0)}" x2="${f(xpos(v))}" y2="${f(y1)}" stroke="${gridC}"${dash}/>`);
    else for (let k = 0; k <= n; k++) { const p = (horiz ? y0 : x0) + band * k; out.push(horiz ? `<line x1="${f(x0)}" y1="${f(p)}" x2="${f(x1)}" y2="${f(p)}" stroke="${gridC}"${dash}/>` : `<line x1="${f(p)}" y1="${f(y0)}" x2="${f(p)}" y2="${f(y1)}" stroke="${gridC}"${dash}/>`); }
  }
  // 값 축 레이블
  if (!yHide) {
    for (let v = sc1.lo; v <= sc1.hi + sc1.step / 2; v += sc1.step) {
      const p = vpos(v, sc1);
      const t = fmtNum(Math.round(v / sc1.step) * sc1.step, fmt1);
      out.push(horiz ? text(p, y1 + fs * 1.25, t, { el: 'axis-y' }) : text(x0 - 6, p + fs / 3, t, { anchor: 'end', el: 'axis-y' }));
    }
  }
  if (has2 && !ax.y2?.hide) {
    for (let v = sc2.lo; v <= sc2.hi + sc2.step / 2; v += sc2.step) {
      const p = vpos(v, sc2);
      out.push(horiz ? text(p, y0 - 4, fmtNum(Math.round(v / sc2.step) * sc2.step, fmt2), { el: 'axis-y2' }) : text(x1 + 6, p + fs / 3, fmtNum(Math.round(v / sc2.step) * sc2.step, fmt2), { anchor: 'start', el: 'axis-y2' }));
    }
  }
  // 범주 축 레이블
  if (!xHide && !dt) {
    if (xy) for (let v = xsc.lo; v <= xsc.hi + xsc.step / 2; v += xsc.step) out.push(text(xpos(v), y1 + fs * 1.25, fmtNum(Math.round(v / xsc.step) * xsc.step, ax.x?.fmt), { el: 'axis-x' }));
    else cats.forEach((ct, ci) => {
      if (ci % step) return;
      const m = catMid(ci);
      if (horiz) out.push(text(x0 - 6, m + fs / 3, ct, { anchor: 'end', el: 'axis-x' }));
      else if (rot) out.push(text(m, y1 + fs, ct, { anchor: 'end', rot, el: 'axis-x' }));
      else out.push(text(m, y1 + fs * 1.25, ct, { el: 'axis-x' }));
    });
  }
  // 축 제목
  if (ax.y?.title) out.push(horiz ? text((x0 + x1) / 2, box.bottom - fs * 0.3, ax.y.title, { el: 'axis-y-title' }) : text(box.left + fs, (y0 + y1) / 2, ax.y.title, { rot: -90, el: 'axis-y-title' }));
  if (ax.x?.title) out.push(horiz ? text(box.left + fs, (y0 + y1) / 2, ax.x.title, { rot: -90, el: 'axis-x-title' }) : text((x0 + x1) / 2, box.bottom - fs * 0.3, ax.x.title, { el: 'axis-x-title' }));
  if (has2 && ax.y2?.title) out.push(text(box.right - fs * 0.6, (y0 + y1) / 2, ax.y2.title, { rot: 90, el: 'axis-y2-title' }));

  // 막대
  const barIdx = series.map((s, si) => si).filter((si) => types[si] === 'col' || types[si] === 'bar');
  const barGroups = [];
  for (const si of barIdx) { const g = groupingOf(c, types[si]); const key = `${series[si].axis === 'secondary' ? 2 : 1}|${g}`; let gr = barGroups.find((x) => x.key === key); if (!gr) barGroups.push(gr = { key, g, idx: [] }); gr.idx.push(si); }
  const gapPct = (c.gap ?? (barGroups.some((g) => g.g !== 'clustered') ? 150 : 219)) / 100;
  // 묶은 막대: 한 범주 안의 막대 수 (쌓은 묶음은 하나)
  const slots = [];
  for (const gr of barGroups) { if (gr.g === 'clustered') for (const si of gr.idx) slots.push([si]); else slots.push(gr.idx); }
  const overlap = (c.overlap ?? (barGroups.some((g) => g.g !== 'clustered') ? 100 : -27)) / 100;
  const ns = Math.max(1, slots.length);
  const bw = band / (ns - (ns - 1) * overlap + gapPct);
  const labels = [];
  slots.forEach((slot, k) => {
    for (const si of slot) {
      const s = series[si];
      const sc = s.axis === 'secondary' ? sc2 : sc1;
      const lo = labelOpts(c, s, c.kind);
      plotted[si].forEach((p, ci) => {
        if (!p || s.ghost) return;
        const col = pointColor(theme, c, s, si, ci, c.vary);
        const a = vpos(p.a, sc);
        const b = vpos(p.b, sc);
        const kk = horiz ? ns - 1 - k : k; // 가로 막대: 첫 계열이 축(아래)에 가깝게 (PowerPoint)
        const start = catMid(ci) - band / 2 + (band - (bw * (ns - (ns - 1) * overlap))) / 2 + kk * bw * (1 - overlap);
        const ol = c.barOutline ? ` stroke="${resolveColor(theme, c.barOutline)}" stroke-width="1"` : '';
        if (horiz) out.push(`<rect x="${f(Math.min(a, b))}" y="${f(start)}" width="${f(Math.abs(b - a))}" height="${f(bw)}" fill="${col}"${ol} data-el="ser${si}"${shadow}/>`);
        else out.push(`<rect x="${f(start)}" y="${f(Math.min(a, b))}" width="${f(bw)}" height="${f(Math.abs(b - a))}" fill="${col}"${ol} data-el="ser${si}"${shadow}/>`);
        if (s.err) errBar(out, s, p, start + bw / 2, sc, vpos, horiz, fc);
        if (lo) {
          const t = labelText(lo, s.labelVals?.[ci] ?? p.v, s, cats[ci], p.p, s.axis === 'secondary' ? fmt2 : lfmt1);
          const stacked = !c.derived && (slot.length > 1 || groupingOf(c, types[si]) !== 'clustered');
          const pos = lo.pos ?? (stacked ? 'ctr' : 'outEnd');
          const neg = p.b < p.a;
          const end = b;
          const white = pos !== 'outEnd';
          if (horiz) {
            const x = pos === 'ctr' ? (a + b) / 2 : pos === 'inBase' ? a + (neg ? -4 : 4) : pos === 'inEnd' ? end + (neg ? 4 : -4) : end + (neg ? -4 : 4);
            const anchor = pos === 'ctr' ? 'middle' : (pos === 'inEnd') === neg ? 'start' : 'end';
            labels.push(text(x, start + bw / 2 + fs / 3, t, { size: fs * 0.9, anchor: pos === 'outEnd' || pos === 'inBase' ? (neg ? 'end' : 'start') : anchor, color: white ? '#fff' : fc }));
          } else {
            const y = pos === 'ctr' ? (a + b) / 2 + fs / 3 : pos === 'inBase' ? a + (neg ? fs : -4) : pos === 'inEnd' ? end + (neg ? -4 : fs) : end + (neg ? fs : -4);
            labels.push(text(start + bw / 2, y, t, { size: fs * 0.9, color: white ? '#fff' : fc }));
          }
        }
      });
    }
  });
  // 영역 → 꺾은선 → 분산/거품
  const ptsOf = (si) => plotted[si].map((p, ci) => (p ? [xy ? xpos(xs[ci]) : catMid(ci), vpos(p.b, series[si].axis === 'secondary' ? sc2 : sc1)] : null));
  series.forEach((s, si) => {
    if (types[si] !== 'area') return;
    const sc = s.axis === 'secondary' ? sc2 : sc1;
    const col = seriesColor(theme, c, si);
    const top = [];
    const base = [];
    plotted[si].forEach((p, ci) => { if (p) { top.push([catMid(ci), vpos(p.b, sc)]); base.push([catMid(ci), vpos(p.a, sc)]); } });
    if (!top.length) return;
    // 영역은 그림 영역 끝까지 (PowerPoint 기본 '눈금 사이' 에서도 가운데부터 그림)
    const d = `${top.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' ')} ${base.reverse().map(([x, y]) => `L${f(x)} ${f(y)}`).join(' ')} Z`;
    out.push(`<path d="${d}" fill="${col}" fill-opacity="${c.kind === 'combo' ? 0.6 : 0.85}" data-el="ser${si}"${shadow}/>`);
    const lo = labelOpts(c, s, c.kind);
    if (lo) plotted[si].forEach((p, ci) => { if (p) labels.push(text(catMid(ci), (vpos(p.a, sc) + vpos(p.b, sc)) / 2 + fs / 3, labelText(lo, p.v, s, cats[ci], p.p, lfmt1), { size: fs * 0.9, color: '#fff' })); });
  });
  series.forEach((s, si) => {
    const t = types[si];
    if (t !== 'line' && t !== 'scatter' && t !== 'bubble') return;
    const sc = s.axis === 'secondary' ? sc2 : sc1;
    const col = seriesColor(theme, c, si);
    const pts = ptsOf(si);
    const lw = (s.lineWidth ?? 2.25) * (4 / 3);
    const smooth = s.smooth ?? (c.variant === 'smooth' || c.smooth);
    if (t === 'bubble') {
      pts.forEach((p, ci) => {
        if (!p) return;
        const r = Math.max(2, Math.sqrt(Math.abs(numVal(s.sizes?.[ci]) ?? 1) / maxSize) * Math.min(x1 - x0, y1 - y0) * 0.12);
        out.push(`<circle cx="${f(p[0])}" cy="${f(p[1])}" r="${f(r)}" fill="${pointColor(theme, c, s, si, ci, c.vary)}" fill-opacity="0.75" stroke="${col}" data-el="ser${si}"/>`);
      });
    } else {
      const drawLine = t === 'line' || (t === 'scatter' && (c.variant === 'smooth' || c.variant === 'line' || s.lineWidth));
      // 빈 값에서 끊기 (PowerPoint 기본: 간격)
      const runs = [[]];
      for (const p of pts) { if (p) runs[runs.length - 1].push(p); else if (runs[runs.length - 1].length) runs.push([]); }
      if (drawLine) for (const r of runs) if (r.length > 1) out.push(`<path d="${lineD(r, smooth)}" fill="none" stroke="${col}" stroke-width="${f(lw)}" stroke-linejoin="round" stroke-linecap="round"${s.dash && DASH[s.dash] ? ` stroke-dasharray="${DASH[s.dash]}"` : ''} data-el="ser${si}"${shadow}/>`);
      const mk = t === 'scatter' ? (s.marker && s.marker !== 'none' ? s.marker : s.marker === 'none' ? null : 'circle') : lineMarker(c, s);
      const mr = (s.markerSize ?? 5) * 0.6 * (fs / 12) + 1.2;
      if (mk) pts.forEach((p, ci) => { if (p) out.push(markerSvg(mk, p[0], p[1], mr, col, s.pointColors?.[ci] ? resolveColor(theme, s.pointColors[ci]) : col)); });
    }
    if (s.err) plotted[si].forEach((p, ci) => { if (p) errBar(out, s, p, pts[ci][0], sc, vpos, false, fc); });
    const lo = labelOpts(c, s, c.kind);
    if (lo) plotted[si].forEach((p, ci) => {
      if (!p) return;
      const [x, y] = pts[ci];
      const pos = lo.pos ?? (t === 'bubble' ? 'ctr' : 'r');
      const tt = labelText(lo, p.v, s, xy ? String(cats[ci]) : cats[ci], p.p, s.axis === 'secondary' ? fmt2 : lfmt1);
      if (pos === 'ctr') labels.push(text(x, y + fs / 3, tt, { size: fs * 0.9 }));
      else if (pos === 'b') labels.push(text(x, y + fs * 1.2, tt, { size: fs * 0.9 }));
      else if (pos === 'l') labels.push(text(x - 6, y + fs / 3, tt, { size: fs * 0.9, anchor: 'end' }));
      else if (pos === 'r' && t !== 'line') labels.push(text(x + 6, y + fs / 3, tt, { size: fs * 0.9, anchor: 'start' }));
      else labels.push(text(x, y - fs * 0.55, tt, { size: fs * 0.9 }));
    });
  });
  // 추세선
  series.forEach((s, si) => {
    const tr = s.trend;
    if (!tr?.type) return;
    const sc = s.axis === 'secondary' ? sc2 : sc1;
    const col = seriesColor(theme, c, si);
    const vx = xy ? xs : cats.map((_, i) => i + 1);
    const vy = (s.vals ?? []).map(numVal);
    const toX = (x) => (xy ? xpos(x) : horiz ? null : catMid(0) + (x - 1) * band * (catRev ? -1 : 1));
    if (horiz) return;
    const clampY = (y) => Math.max(y0 - 2, Math.min(y1 + 2, y));
    if (tr.type === 'movingAvg') {
      const per = Math.max(2, tr.period ?? 2);
      const pts = [];
      for (let i = per - 1; i < vy.length; i++) { const win = vy.slice(i - per + 1, i + 1); if (win.some((v) => v == null)) continue; pts.push([toX(vx[i]), clampY(vpos(win.reduce((a, b) => a + b, 0) / per, sc))]); }
      if (pts.length > 1) out.push(`<path d="${lineD(pts, false)}" fill="none" stroke="${col}" stroke-width="1.6" stroke-dasharray="4 3"/>`);
      return;
    }
    const fit = trendFit(vx, vy, tr.type, tr.order);
    if (!fit) return;
    const xa = Math.min(...vx) - (tr.backward ?? 0);
    const xb = Math.max(...vx) + (tr.forward ?? 0);
    const pts = [];
    for (let k = 0; k <= 40; k++) { const x = xa + ((xb - xa) * k) / 40; const y = fit.f(x); if (Number.isFinite(y) && (!(tr.type === 'log' || tr.type === 'power') || x > 0)) pts.push([toX(x), clampY(vpos(y, sc))]); }
    if (pts.length > 1) out.push(`<path d="${lineD(pts, false)}" fill="none" stroke="${col}" stroke-width="1.6" stroke-dasharray="4 3" data-el="trend${si}"/>`);
    if (tr.eq || tr.r2) {
      const lines = [tr.eq ? fit.eq : null, tr.r2 ? `R² = ${fit.r2.toFixed(4)}` : null].filter(Boolean);
      const [lx, ly] = pts[pts.length - 1] ?? [x1, y0];
      lines.forEach((l, i) => labels.push(text(Math.min(lx, x1 - 4), ly - fs * (lines.length - i) - 2, l, { size: fs * 0.85, anchor: 'end', color: col })));
    }
  });
  // 축 선
  const z1 = vpos(Math.max(sc1.lo, Math.min(sc1.hi, 0)), sc1);
  if (!xHide) out.push(horiz ? `<line x1="${f(z1)}" y1="${f(y0)}" x2="${f(z1)}" y2="${f(y1)}" stroke="${axisC}" stroke-width="1"/>` : `<line x1="${f(x0)}" y1="${f(xy ? y1 : z1)}" x2="${f(x1)}" y2="${f(xy ? y1 : z1)}" stroke="${axisC}" stroke-width="1"/>`);
  if (xy && !yHide) out.push(`<line x1="${f(xpos(Math.max(xsc.lo, Math.min(xsc.hi, 0))))}" y1="${f(y0)}" x2="${f(xpos(Math.max(xsc.lo, Math.min(xsc.hi, 0))))}" y2="${f(y1)}" stroke="${axisC}"/>`);
  out.push(...labels);
  // 데이터 표
  if (dt) {
    const rh = fs * 1.5;
    const ty = y1;
    const line = (a, b, c2, d2) => `<line x1="${f(a)}" y1="${f(b)}" x2="${f(c2)}" y2="${f(d2)}" stroke="${axisC}" stroke-width="0.8"/>`;
    const g = [];
    for (let r = 0; r <= dtRows; r++) g.push(line(r === 0 ? x0 : box.left, ty + r * rh, x1, ty + r * rh));
    g.push(line(box.left, ty + rh, box.left, ty + dtRows * rh));
    for (let k = 0; k <= n; k++) g.push(line(x0 + band * k, ty, x0 + band * k, ty + dtRows * rh));
    cats.forEach((ct, ci) => g.push(text(catMid(ci), ty + rh * 0.7, ct, { size: fs * 0.9 })));
    series.forEach((s, si) => {
      const y = ty + rh * (si + 1);
      const col = seriesColor(theme, c, si);
      if (c.dataTable !== 'nokeys') g.push(legendKey(box.left + 4, y + rh * 0.72, { kind: types[si] === 'line' ? 'line' : 'rect', color: col }));
      g.push(text(box.left + (c.dataTable === 'nokeys' ? 4 : fs * 1.4), y + rh * 0.7, s.name, { size: fs * 0.9, anchor: 'start' }));
      cats.forEach((_, ci) => g.push(text(catMid(ci), y + rh * 0.7, fmtNum(s.vals?.[ci], s.axis === 'secondary' ? fmt2 : fmt1), { size: fs * 0.9 })));
    });
    out.push(`<g data-el="table">${g.join('')}</g>`);
  }
  return out.join('');
}
function errAmount(s, v) {
  const e = s.err;
  if (!e) return 0;
  const vals = (s.vals ?? []).map(numVal).filter((x) => x != null);
  const mean = vals.reduce((a, b) => a + b, 0) / Math.max(1, vals.length);
  const sd = Math.sqrt(vals.reduce((a, b) => a + (b - mean) ** 2, 0) / Math.max(1, vals.length - 1));
  if (e.type === 'percent') return Math.abs(v) * (e.val ?? 5) / 100;
  if (e.type === 'stdDev') return sd * (e.val ?? 1);
  if (e.type === 'stdErr') return sd / Math.sqrt(Math.max(1, vals.length));
  return Math.abs(e.val ?? 1);
}
function errBar(out, s, p, at, sc, vpos, horiz, color) {
  const e = errAmount(s, p.v);
  const a = vpos(p.b - e, sc);
  const b = vpos(p.b + e, sc);
  const k = 4;
  if (horiz) out.push(`<path d="M${f(a)} ${f(at)} L${f(b)} ${f(at)} M${f(a)} ${f(at - k)} L${f(a)} ${f(at + k)} M${f(b)} ${f(at - k)} L${f(b)} ${f(at + k)}" stroke="${color}" stroke-width="1" fill="none"/>`);
  else out.push(`<path d="M${f(at)} ${f(a)} L${f(at)} ${f(b)} M${f(at - k)} ${f(a)} L${f(at + k)} ${f(a)} M${f(at - k)} ${f(b)} L${f(at + k)} ${f(b)}" stroke="${color}" stroke-width="1" fill="none"/>`);
}

function wrap(w, h, font, parts, defs = []) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${f(w)}" height="${f(h)}" viewBox="0 0 ${f(w)} ${f(h)}" font-family="${esc(font)}, 'Malgun Gothic', sans-serif">${defs.length ? `<defs>${defs.join('')}</defs>` : ''}${parts.join('')}</svg>`;
}

// ───────────── 스타일 · 빠른 레이아웃 (PowerPoint 차트 디자인 탭) ─────────────
/** 차트 스타일 1–8: 고르면 이 값들을 차트에 씀 */
export const CHART_STYLES = [
  { bg: null, plotBg: null, fontColor: null, barOutline: null, dashGrid: false, shadow: false, grid: { major: true }, gap: null, labels: false },
  { bg: null, plotBg: null, fontColor: null, barOutline: null, dashGrid: false, shadow: false, grid: { major: false }, gap: 100, labels: { pos: 'inEnd' }, axes: { y: { hide: true } } },
  { bg: null, plotBg: null, fontColor: null, barOutline: '@bg1', dashGrid: true, shadow: false, grid: { major: true }, gap: 80, labels: false },
  { bg: null, plotBg: null, fontColor: null, barOutline: null, dashGrid: false, shadow: true, grid: { major: true }, gap: 150, labels: false },
  { bg: '@bg1:lm95', plotBg: '@bg1', fontColor: null, barOutline: null, dashGrid: false, shadow: false, grid: { major: true }, gap: 120, labels: false },
  { bg: '@tx1:lm75:lo25', plotBg: null, fontColor: '#FFFFFF', gridColor: '@bg1:lm50', barOutline: null, dashGrid: false, shadow: false, grid: { major: true }, gap: 120, labels: false },
  { bg: '@accent1:lm20:lo80', plotBg: null, fontColor: null, barOutline: '@accent1:lm50', dashGrid: false, shadow: false, grid: { major: true }, gap: 60, labels: { pos: 'outEnd' } },
  { bg: '@accent1:lm50', plotBg: null, fontColor: '#FFFFFF', gridColor: '@bg1:lm75', barOutline: '@bg1', dashGrid: true, shadow: true, grid: { major: true }, gap: 100, labels: false },
];
const STYLE_KEYS = ['bg', 'plotBg', 'fontColor', 'gridColor', 'barOutline', 'dashGrid', 'shadow', 'grid', 'gap', 'labels'];
export function applyChartStyle(chart, n) {
  const st = CHART_STYLES[n - 1] ?? CHART_STYLES[0];
  const out = { ...chart, style: n };
  for (const k of STYLE_KEYS) { if (st[k] == null || st[k] === false) delete out[k]; else out[k] = JSON.parse(JSON.stringify(st[k])); }
  if (st.axes) out.axes = { ...(chart.axes ?? {}), y: { ...(chart.axes?.y ?? {}), ...st.axes.y } };
  else if (chart.axes?.y?.hide && chart.style === 2) out.axes = { ...chart.axes, y: { ...chart.axes.y, hide: undefined } };
  return out;
}
/** 빠른 레이아웃 1–10 */
export const CHART_LAYOUTS = [
  { title: true, legend: 'r', labels: false, xt: false, yt: false, table: false, grid: true },
  { title: true, legend: 't', labels: { pos: 'outEnd' }, xt: false, yt: false, table: false, grid: false, yhide: true },
  { title: true, legend: 'b', labels: false, xt: false, yt: false, table: false, grid: true },
  { title: false, legend: 'b', labels: { pos: 'outEnd' }, xt: false, yt: false, table: false, grid: false },
  { title: true, legend: null, labels: false, xt: false, yt: true, table: true, grid: true },
  { title: true, legend: null, labels: false, xt: false, yt: true, table: false, grid: true },
  { title: false, legend: 'r', labels: false, xt: true, yt: true, table: false, grid: true },
  { title: true, legend: null, labels: false, xt: true, yt: true, table: false, grid: false },
  { title: true, legend: 'r', labels: false, xt: true, yt: true, table: false, grid: true },
  { title: true, legend: 'r', labels: { pos: 'ctr' }, xt: false, yt: false, table: false, grid: true },
];
export function applyChartLayout(chart, n) {
  const L = CHART_LAYOUTS[n - 1] ?? CHART_LAYOUTS[0];
  const out = { ...chart, layout: n, legend: L.legend, labels: L.labels ? { ...L.labels } : false, dataTable: L.table || undefined, grid: { ...(chart.grid ?? {}), major: L.grid } };
  out.title = L.title ? chart.title || '차트 제목' : '';
  const axes = { x: { ...(chart.axes?.x ?? {}) }, y: { ...(chart.axes?.y ?? {}) }, ...(chart.axes?.y2 ? { y2: { ...chart.axes.y2 } } : {}) };
  axes.x.title = L.xt ? axes.x.title || '축 제목' : undefined;
  axes.y.title = L.yt ? axes.y.title || '축 제목' : undefined;
  axes.y.hide = L.yhide || undefined;
  out.axes = axes;
  return out;
}

// ───────────── 데이터 ─────────────
/** 차트 데이터 ↔ 표 텍스트 (탭 구분, 첫 행 = 계열 이름, 첫 열 = 항목; 거품형은 계열마다 '크기' 열이 하나 더) */
export function chartToTsv(chart) {
  const bub = chart.kind === 'bubble';
  const rows = [['', ...chart.series.flatMap((s) => (bub ? [s.name, `${s.name} 크기`] : [s.name]))]];
  chart.cats.forEach((c, i) => rows.push([c, ...chart.series.flatMap((s) => (bub ? [s.vals[i] ?? '', s.sizes?.[i] ?? ''] : [s.vals[i] ?? '']))]));
  return rows.map((r) => r.join('\t')).join('\n');
}
/** 표 (2차원 배열) → 차트 데이터 */
export function gridToChart(chart, rows) {
  rows = rows.map((r) => r.map((v) => String(v ?? '')));
  while (rows.length && rows[rows.length - 1].every((v) => v.trim() === '')) rows.pop();
  if (rows.length < 2) throw new Error('항목 행이 하나 이상 있어야 합니다');
  const width = Math.max(...rows.map((r) => r.length));
  let lastCol = width - 1;
  while (lastCol > 0 && rows.every((r) => String(r[lastCol] ?? '').trim() === '')) lastCol--;
  const bub = chart.kind === 'bubble';
  const head = rows[0].slice(1, lastCol + 1);
  const body = rows.slice(1);
  const num = (v) => { const x = Number(String(v ?? '').replace(/[,\s₩$%]/g, '')); return String(v ?? '').trim() === '' ? null : Number.isFinite(x) ? (/%\s*$/.test(String(v)) ? x / 100 : x) : null; };
  const series = [];
  for (let k = 0; k < head.length; k += bub ? 2 : 1) {
    const si = series.length;
    const prev = chart.series[si] ?? {};
    series.push({ ...prev, name: String(head[k] ?? '').trim() || `계열 ${si + 1}`, vals: body.map((r) => num(r[k + 1]) ?? 0), ...(bub ? { sizes: body.map((r) => num(r[k + 2]) ?? 1) } : {}) });
  }
  if (!series.length) throw new Error('계열(열)이 하나 이상 있어야 합니다');
  return { ...chart, cats: body.map((r) => r[0] ?? ''), series };
}
export function tsvToChart(chart, tsv) {
  const text = String(tsv).replace(/\r/g, '');
  const sep = text.includes('\t') ? '\t' : ',';
  return gridToChart(chart, text.split('\n').filter((l) => l.trim() !== '').map((l) => l.split(sep)));
}
/** 행/열 전환 */
export function switchRowCol(chart) {
  const series = chart.cats.map((cat, ci) => ({ name: String(cat), vals: chart.series.map((s) => s.vals[ci] ?? 0) }));
  return { ...chart, cats: chart.series.map((s) => s.name), series };
}
/** CSV (엑셀용 BOM 포함) */
export function chartCsv(chart) {
  const q = (v) => { const s = String(v ?? ''); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  return `﻿${chartToTsv(chart).split('\n').map((l) => l.split('\t').map(q).join(',')).join('\r\n')}`;
}

// ───────────── 폭포 · 깔때기 · 히스토그램 · 상자 수염 · 트리맵 · 선버스트 (Office 2016 차트) ─────────────
const SPECIAL = new Set(['box', 'treemap', 'sunburst']);
const TOTAL_RE = /^(합계|총계|소계|누계|total|subtotal|sum)$/i;
/** 폭포의 합계 막대 위치 (series.totals 또는 이름이 합계 · 총계 · Total) */
export const waterfallTotals = (ch) => new Set(ch.series?.[0]?.totals ?? (ch.cats ?? []).map((c, i) => (TOTAL_RE.test(String(c).trim()) ? i : -1)).filter((i) => i >= 0));
/** 히스토그램 구간 → { cats, counts } (구간 수 = 데이터 수의 제곱근, 3~20) */
export function histogramBins(values, { count, width } = {}) {
  const v = values.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
  if (!v.length) return { cats: [], counts: [] };
  const lo = v[0];
  const hi = v[v.length - 1];
  let n = count ?? Math.max(3, Math.min(20, Math.ceil(Math.sqrt(v.length))));
  let bw = width ?? (hi - lo) / n;
  if (!(bw > 0)) { bw = 1; n = 1; } else n = Math.max(1, Math.ceil((hi - lo) / bw - 1e-9));
  const counts = new Array(n).fill(0);
  for (const x of v) counts[Math.min(n - 1, Math.max(0, Math.ceil((x - lo) / bw - 1e-9) - 1 + (x === lo ? 1 : 0)))]++;
  const r = (x) => fmtNum(Math.round(x * 100) / 100);
  const cats = counts.map((_, i) => `${i === 0 ? '[' : '('}${r(lo + bw * i)}, ${r(lo + bw * (i + 1))}]`);
  return { cats, counts };
}
/** 사분위수 (엑셀 상자 수염: 포함 중앙값 방식) */
export function boxStats(values) {
  const v = values.map(Number).filter(Number.isFinite).sort((a, b) => a - b);
  if (!v.length) return null;
  const q = (p) => { const k = (v.length - 1) * p; const a = Math.floor(k); return v[a] + (v[Math.min(v.length - 1, a + 1)] - v[a]) * (k - a); };
  const q1 = q(0.25); const med = q(0.5); const q3 = q(0.75); const iqr = q3 - q1;
  const lo = v.find((x) => x >= q1 - 1.5 * iqr); const hi = [...v].reverse().find((x) => x <= q3 + 1.5 * iqr);
  return { q1, med, q3, lo, hi, mean: v.reduce((s, x) => s + x, 0) / v.length, out: v.filter((x) => x < lo || x > hi), min: v[0], max: v[v.length - 1] };
}
/** 화면 그리기용으로 바꾼 차트 (폭포 · 깔때기 · 히스토그램 → 누적/묶은 막대) — 원래 차트는 그대로 */
export function expandChart(ch) {
  const k = ch.kind;
  if (k === 'waterfall') {
    const vals = (ch.series?.[0]?.vals ?? []).map((v) => Number(v) || 0);
    const totals = waterfallTotals(ch);
    const base = []; const up = []; const down = []; const tot = []; const lbl = [];
    let run = 0;
    vals.forEach((v, i) => {
      if (totals.has(i)) { base.push(0); tot.push(run); up.push(null); down.push(null); lbl.push(run); return; }
      if (v >= 0) { base.push(run); up.push(v); down.push(null); } else { base.push(run + v); down.push(-v); up.push(null); }
      tot.push(null); lbl.push(v); run += v;
    });
    const name = ch.series?.[0]?.name ?? '';
    return { ...ch, kind: 'col', derived: 'waterfall', grouping: 'stacked', gap: ch.gap ?? 50, labels: ch.labels ?? { pos: 'inEnd' }, series: [
      { name: '기준', ghost: true, vals: base },
      { name: '증가', vals: up, color: ch.series?.[0]?.upColor ?? '@accent1', labelVals: lbl },
      { name: '감소', vals: down, color: ch.series?.[0]?.downColor ?? '@accent2', labelVals: lbl },
      { name: '합계', vals: tot, color: ch.series?.[0]?.totalColor ?? '@accent3', labelVals: lbl },
    ], title: ch.title ?? name };
  }
  if (k === 'funnel') {
    const vals = (ch.series?.[0]?.vals ?? []).map((v) => Math.max(0, Number(v) || 0));
    const max = Math.max(1, ...vals);
    return { ...ch, kind: 'bar', derived: 'funnel', grouping: 'stacked', gap: ch.gap ?? 10, legend: null, labels: ch.labels ?? { pos: 'ctr' }, grid: { major: false },
      axes: { ...(ch.axes ?? {}), x: { ...(ch.axes?.x ?? {}), reverse: true }, y: { ...(ch.axes?.y ?? {}), hide: true } },
      series: [{ name: 'offset', ghost: true, vals: vals.map((v) => (max - v) / 2) }, { ...(ch.series?.[0] ?? {}), name: ch.series?.[0]?.name ?? '', vals, color: ch.series?.[0]?.color ?? paletteRef(ch.palette, 0) }] };
  }
  if (k === 'histogram') {
    const all = (ch.series ?? []).flatMap((s) => s.vals ?? []);
    const { cats, counts } = histogramBins(all, ch.bins ?? {});
    const out = { ...ch, kind: 'col', derived: 'histogram', cats, gap: ch.gap ?? 0, legend: ch.variant === 'pareto' ? ch.legend : null, series: [{ name: '빈도', vals: counts, color: ch.series?.[0]?.color }] };
    if (ch.variant === 'pareto') {
      // 파레토: 빈도 내림차순 + 누적 백분율 꺾은선 (보조 축)
      const order = counts.map((n, i) => [n, i]).sort((a, b) => b[0] - a[0]);
      const tot = counts.reduce((s, n) => s + n, 0) || 1;
      let acc = 0;
      out.kind = 'combo';
      out.cats = order.map(([, i]) => cats[i]);
      out.series = [{ name: '빈도', type: 'col', vals: order.map(([n]) => n) }, { name: '누적 %', type: 'line', axis: 'secondary', vals: order.map(([n]) => (acc += n) / tot) }];
      out.axes = { ...(ch.axes ?? {}), y2: { min: 0, max: 1, fmt: '0%', ...(ch.axes?.y2 ?? {}) } };
    }
    return out;
  }
  return ch;
}
function specialLegend(c, theme) {
  if (c.kind === 'box') return (c.series ?? []).map((s, i) => ({ name: s.name, color: seriesColor(theme, c, i), kind: 'rect' }));
  return treeData(c).map((g, i) => ({ name: g.name, color: paletteColor(theme, c.palette, i), kind: 'rect' }));
}
/** 트리맵 · 선버스트 자료: 항목 이름 '부모/자식' (또는 '부모 > 자식') → [{ name, value, kids: [{ name, value }] }] */
export function treeData(ch) {
  const vals = ch.series?.[0]?.vals ?? [];
  const groups = new Map();
  (ch.cats ?? []).forEach((cat, i) => {
    const parts = String(cat).split(/\s*[/>]\s*/).filter(Boolean);
    const v = Math.max(0, Number(vals[i]) || 0);
    const [p, ...rest] = parts.length > 1 ? parts : [String(cat)];
    if (!groups.has(p)) groups.set(p, { name: p, value: 0, kids: [] });
    const g = groups.get(p);
    g.value += v;
    if (rest.length) g.kids.push({ name: rest.join(' / '), value: v });
  });
  return [...groups.values()].filter((g) => g.value > 0);
}
/** 정사각형에 가깝게 나누기 (squarified treemap) */
export function squarify(items, x, y, w, h) {
  const total = items.reduce((s, it) => s + it.value, 0) || 1;
  const area = w * h;
  const list = items.map((it) => ({ it, a: (it.value / total) * area })).sort((a, b) => b.a - a.a);
  const out = [];
  let rx = x; let ry = y; let rw = w; let rh = h;
  let row = [];
  const worst = (r, side) => { const s = r.reduce((t, e) => t + e.a, 0); const mx = Math.max(...r.map((e) => e.a)); const mn = Math.min(...r.map((e) => e.a)); return Math.max((side * side * mx) / (s * s), (s * s) / (side * side * mn)); };
  const lay = () => {
    const s = row.reduce((t, e) => t + e.a, 0);
    if (rw >= rh) { const cw = s / rh; let yy = ry; for (const e of row) { const eh = e.a / cw; out.push({ ...e.it, x: rx, y: yy, w: cw, h: eh }); yy += eh; } rx += cw; rw -= cw; }
    else { const ch2 = s / rw; let xx = rx; for (const e of row) { const ew = e.a / ch2; out.push({ ...e.it, x: xx, y: ry, w: ew, h: ch2 }); xx += ew; } ry += ch2; rh -= ch2; }
    row = [];
  };
  for (const e of list) {
    const side = Math.min(rw, rh);
    if (!row.length || worst([...row, e], side) <= worst(row, side)) row.push(e);
    else { lay(); row.push(e); }
  }
  if (row.length) lay();
  return out;
}
function treemapSvg(c, box, theme, { fs, text }) {
  const groups = treeData(c);
  const out = [];
  const W = box.right - box.left; const H = box.bottom - box.top;
  for (const [gi, g] of squarify(groups, box.left, box.top, W, H).map((r) => [groups.indexOf(groups.find((x) => x.name === r.name)), r])) {
    const base = paletteColor(theme, c.palette, gi);
    const kids = g.kids.length ? squarify(g.kids, g.x, g.y, g.w, g.h) : [{ ...g, name: g.name }];
    for (const k of kids) {
      out.push(`<rect x="${f(k.x)}" y="${f(k.y)}" width="${f(Math.max(0, k.w))}" height="${f(Math.max(0, k.h))}" fill="${base}" stroke="#fff" stroke-width="1.5" data-el="ser0"/>`);
      if (k.w > fs * 2.5 && k.h > fs * 1.4) {
        out.push(text(k.x + 4, k.y + fs * 1.05, k.name, { size: fs * 0.85, anchor: 'start', color: '#fff' }));
        if (c.labels !== false && k.h > fs * 2.6) out.push(text(k.x + 4, k.y + fs * 2.15, fmtNum(k.value, c.axes?.y?.fmt), { size: fs * 0.8, anchor: 'start', color: '#fff' }));
      }
    }
    if (g.kids.length && g.w > fs * 3) out.push(`<rect x="${f(g.x)}" y="${f(g.y)}" width="${f(g.w)}" height="${f(g.h)}" fill="none" stroke="#fff" stroke-width="3"/>`);
  }
  return out.join('');
}
function sunburstSvg(c, box, theme, { fs, text }) {
  const groups = treeData(c);
  const total = groups.reduce((s, g) => s + g.value, 0) || 1;
  const cx = (box.left + box.right) / 2; const cy = (box.top + box.bottom) / 2;
  const R = Math.max(10, Math.min(box.right - box.left, box.bottom - box.top) / 2 - 4);
  const hasKids = groups.some((g) => g.kids.length);
  const r0 = R * 0.18; const r1 = hasKids ? R * 0.58 : R; const r2 = R;
  const arc = (ra, rb, a0, a1, color, label, el) => {
    const P = (r, a) => `${f(cx + r * Math.cos(a))} ${f(cy + r * Math.sin(a))}`;
    const large = a1 - a0 > Math.PI ? 1 : 0;
    const d = a1 - a0 >= Math.PI * 2 - 1e-6
      ? `M${P(rb, 0)} A${f(rb)} ${f(rb)} 0 1 1 ${P(rb, Math.PI)} A${f(rb)} ${f(rb)} 0 1 1 ${P(rb, 0)} M${P(ra, 0)} A${f(ra)} ${f(ra)} 0 1 0 ${P(ra, Math.PI)} A${f(ra)} ${f(ra)} 0 1 0 ${P(ra, 0)} Z`
      : `M${P(ra, a0)} L${P(rb, a0)} A${f(rb)} ${f(rb)} 0 ${large} 1 ${P(rb, a1)} L${P(ra, a1)} A${f(ra)} ${f(ra)} 0 ${large} 0 ${P(ra, a0)} Z`;
    const am = (a0 + a1) / 2; const rm = (ra + rb) / 2;
    const fits = (a1 - a0) * rm > fs * 2.2 && rb - ra > fs * 1.1;
    return `<path d="${d}" fill="${color}" stroke="#fff" stroke-width="1.5" fill-rule="evenodd" data-el="${el}"/>${fits ? text(cx + rm * Math.cos(am), cy + rm * Math.sin(am) + fs / 3, label, { size: fs * 0.8, color: '#fff' }) : ''}`;
  };
  const out = [];
  let a = -Math.PI / 2;
  groups.forEach((g, gi) => {
    const sw = (g.value / total) * Math.PI * 2;
    const col = paletteColor(theme, c.palette, gi);
    out.push(arc(r0, r1, a, a + sw, col, g.name, 'ser0'));
    let b = a;
    for (const k of g.kids) { const ks = (k.value / total) * Math.PI * 2; out.push(arc(r1, r2, b, b + ks, resolveColor(theme, `${['@accent1', '@accent2', '@accent3', '@accent4', '@accent5', '@accent6'][gi % 6]}:lm60:lo40`) ?? col, k.name, 'ser0')); b += ks; }
    a += sw;
  });
  return out.join('');
}
function boxPlotSvg(c, series, box, theme, { fs, text, gridC, axisC }) {
  const stats = series.map((s) => boxStats(s.vals ?? []));
  const all = stats.filter(Boolean).flatMap((st) => [st.min, st.max]);
  if (!all.length) return '';
  const sc = niceScale(Math.min(...all), Math.max(...all), 5, c.axes?.y ?? {});
  const labels = []; for (let v = sc.lo; v <= sc.hi + sc.step / 2; v += sc.step) labels.push([v, fmtNum(Math.round(v / sc.step) * sc.step, c.axes?.y?.fmt)]);
  const x0 = box.left + Math.max(...labels.map(([, t]) => textWidth(t, fs))) + 8; const x1 = box.right - 6;
  const y0 = box.top + 4; const y1 = box.bottom - fs * 1.6;
  const Y = (v) => y1 - ((v - sc.lo) / (sc.hi - sc.lo)) * (y1 - y0);
  const out = [];
  for (const [v, t] of labels) { out.push(`<line x1="${f(x0)}" y1="${f(Y(v))}" x2="${f(x1)}" y2="${f(Y(v))}" stroke="${gridC}"/>`); out.push(text(x0 - 6, Y(v) + fs / 3, t, { anchor: 'end' })); }
  const n = Math.max(1, series.length); const band = (x1 - x0) / n; const bw = band * 0.45;
  series.forEach((s, i) => {
    const st = stats[i];
    const cx = x0 + band * (i + 0.5);
    out.push(text(cx, y1 + fs * 1.25, s.name));
    if (!st) return;
    const col = seriesColor(theme, c, i);
    out.push(`<line x1="${f(cx)}" y1="${f(Y(st.hi))}" x2="${f(cx)}" y2="${f(Y(st.q3))}" stroke="${axisC}" stroke-width="1.2"/><line x1="${f(cx)}" y1="${f(Y(st.q1))}" x2="${f(cx)}" y2="${f(Y(st.lo))}" stroke="${axisC}" stroke-width="1.2"/>`);
    out.push(`<line x1="${f(cx - bw / 4)}" y1="${f(Y(st.hi))}" x2="${f(cx + bw / 4)}" y2="${f(Y(st.hi))}" stroke="${axisC}" stroke-width="1.2"/><line x1="${f(cx - bw / 4)}" y1="${f(Y(st.lo))}" x2="${f(cx + bw / 4)}" y2="${f(Y(st.lo))}" stroke="${axisC}" stroke-width="1.2"/>`);
    out.push(`<rect x="${f(cx - bw / 2)}" y="${f(Y(st.q3))}" width="${f(bw)}" height="${f(Math.max(0.5, Y(st.q1) - Y(st.q3)))}" fill="${col}" stroke="${col}" data-el="ser${i}"/>`);
    out.push(`<line x1="${f(cx - bw / 2)}" y1="${f(Y(st.med))}" x2="${f(cx + bw / 2)}" y2="${f(Y(st.med))}" stroke="#fff" stroke-width="1.5"/>`);
    out.push(`<path d="M${f(cx - 4)} ${f(Y(st.mean) - 4)} L${f(cx + 4)} ${f(Y(st.mean) + 4)} M${f(cx + 4)} ${f(Y(st.mean) - 4)} L${f(cx - 4)} ${f(Y(st.mean) + 4)}" stroke="#fff" stroke-width="1.5"/>`);
    for (const o of st.out) out.push(`<circle cx="${f(cx)}" cy="${f(Y(o))}" r="${f(fs * 0.22)}" fill="none" stroke="${col}"/>`);
    if (c.labels) for (const [v, k] of [[st.med, 'm'], [st.q1, 'q'], [st.q3, 'q']]) { void k; out.push(text(cx + bw / 2 + 4, Y(v) + fs / 3, fmtNum(v, c.axes?.y?.fmt), { size: fs * 0.8, anchor: 'start' })); }
  });
  out.push(`<line x1="${f(x0)}" y1="${f(y1)}" x2="${f(x1)}" y2="${f(y1)}" stroke="${axisC}"/>`);
  return out.join('');
}
/** pptx 에 쓸 때: 다른 프로그램이 읽을 수 있는 일반 차트 (원본은 WIPOINT 확장에 그대로) */
export function fallbackChart(ch) {
  if (ch.kind === 'waterfall' || ch.kind === 'funnel' || ch.kind === 'histogram') {
    const e = expandChart(ch);
    return { ...e, series: e.series.map((s) => (s.ghost ? { ...s, color: undefined, noFill: true } : s)) };
  }
  if (ch.kind === 'treemap' || ch.kind === 'sunburst') return { ...ch, kind: ch.kind === 'sunburst' ? 'doughnut' : 'bar', series: ch.series.slice(0, 1) };
  if (ch.kind === 'box') {
    const stats = ch.series.map((s) => boxStats(s.vals ?? []) ?? { lo: 0, q1: 0, med: 0, q3: 0, hi: 0 });
    return { ...ch, kind: 'col', grouping: 'stacked', cats: ch.series.map((s) => s.name), series: [
      { name: '최소', vals: stats.map((st) => st.lo), noFill: true }, { name: '1사분위', vals: stats.map((st) => st.q1 - st.lo), noFill: true },
      { name: '중간값 아래', vals: stats.map((st) => st.med - st.q1) }, { name: '중간값 위', vals: stats.map((st) => st.q3 - st.med) }] };
  }
  return ch;
}

// ───────────── 3차원 차트 (PowerPoint 3-D 세로/가로 막대 · 꺾은선 · 영역 · 원형) ─────────────
/** 3차원 회전 기본값 (PowerPoint: X 회전 15°, Y 회전 20°, 깊이 100%) — 원형은 X 30° */
export const VIEW3D_DEFAULT = { rotX: 15, rotY: 20, depth: 100, persp: 30 };
export const is3d = (c) => !!c?.d3 && ['col', 'bar', 'line', 'area', 'pie'].includes(c.kind);
/** 색 밝게/어둡게 (k>0 흰색 쪽, k<0 검은색 쪽) */
export function shadeHex(hex, k) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex ?? ''));
  if (!m) return hex;
  const n = parseInt(m[1], 16);
  const ch = [(n >> 16) & 255, (n >> 8) & 255, n & 255].map((v) => Math.round(k >= 0 ? v + (255 - v) * k : v * (1 + k)));
  return `#${ch.map((v) => v.toString(16).padStart(2, '0')).join('')}`;
}
const poly = (pts, fill, extra = '') => `<path d="M${pts.map(([x, y]) => `${f(x)} ${f(y)}`).join(' L')} Z" fill="${fill}"${extra}/>`;
/** 상자(직육면체) 하나: 앞면 (x, y, w, h) + 깊이 (dx, dy) — 위 · 옆면은 밝기를 달리 */
function prism(x, y, w, h, dx, dy, color, el) {
  const top = Math.min(y, y + h);
  const bot = Math.max(y, y + h);
  const a = el ? ` data-el="${el}"` : '';
  return poly([[x + w, top], [x + w + dx, top + dy], [x + w + dx, bot + dy], [x + w, bot]], shadeHex(color, -0.28), a)
    + poly([[x, top], [x + dx, top + dy], [x + w + dx, top + dy], [x + w, top]], shadeHex(color, 0.25), a)
    + `<rect x="${f(x)}" y="${f(top)}" width="${f(Math.max(0, w))}" height="${f(bot - top)}" fill="${color}"${a}/>`;
}

/** 3차원 가로 · 세로 막대, 꺾은선(리본), 영역 */
function cartesian3dSvg(c, series, cats, box, theme, { fs, fc, text, gridC, axisC }) {
  const out = [];
  const v3 = { ...VIEW3D_DEFAULT, ...(c.view3d ?? {}) };
  const kind = c.kind;
  const horiz = kind === 'bar';
  const grouping = groupingOf(c, kind === 'col' || kind === 'bar' ? 'col' : 'line');
  const full = c.variant3d === 'full' || ((kind === 'line' || kind === 'area') && grouping === 'standard');
  const n = Math.max(1, cats.length);
  const m = Math.max(1, series.length);
  const vals = series.map((s) => cats.map((_, i) => numVal(s.vals?.[i]) ?? 0));
  // 누적 / 100%
  const stacked = grouping === 'stacked' || grouping === 'percent';
  const totals = cats.map((_, i) => vals.reduce((a, v) => a + Math.abs(v[i]), 0) || 1);
  const val = (si, i) => (grouping === 'percent' ? (vals[si][i] / totals[i]) * 100 : vals[si][i]);
  const base = (si, i) => { let b = 0; for (let k = 0; k < si; k++) b += val(k, i); return stacked ? b : 0; };
  let lo = Infinity; let hi = -Infinity;
  series.forEach((_, si) => cats.forEach((__, i) => { const b = base(si, i); const t = b + val(si, i); lo = Math.min(lo, b, t); hi = Math.max(hi, b, t); }));
  if (!Number.isFinite(lo)) { lo = 0; hi = 1; }
  const sc = niceScale(Math.min(0, lo), Math.max(0, hi), 5, { min: c.axes?.y?.min, max: c.axes?.y?.max, major: c.axes?.y?.major });
  // 깊이: 그림 크기 · Y 회전 · X 회전
  const W = box.right - box.left; const H = box.bottom - box.top;
  const depthPx = Math.min(W, H) * 0.16 * (v3.depth / 100) * (full ? Math.min(2.2, 0.8 + m * 0.35) : 1);
  const dx = depthPx * Math.cos((Math.min(80, Math.max(0, v3.rotY)) * Math.PI) / 180);
  const dy = -depthPx * Math.max(0.15, Math.sin((Math.min(80, Math.max(5, v3.rotX)) * Math.PI) / 180) * 2.2);
  const valLbl = (v) => fmtNum(v, grouping === 'percent' ? '0"%"' : c.axes?.y?.fmt);
  const labW = Math.max(...[sc.lo, sc.hi].map((v) => textWidth(valLbl(v), fs * 0.9))) + 8;
  // 항목 이름이 길면 기울여 쓰므로 아래 자리를 더 둠
  const catMaxW = Math.max(0, ...cats.map((t) => textWidth(String(t), fs * 0.9)));
  const catH = !horiz && catMaxW > ((W - labW - Math.abs(dx)) / Math.max(1, n)) * 0.95 ? Math.min(H * 0.35, catMaxW * 0.72 + fs) : fs * 1.6;
  const x0 = box.left + (horiz ? Math.min(W * 0.3, Math.max(...cats.map((t) => textWidth(String(t), fs * 0.9)), 10) + 10) : labW);
  const x1 = box.right - dx - 4;
  const y0 = box.top - dy + 4;
  const y1 = box.bottom - (horiz ? catH : catH);
  const vpos = (v) => (horiz ? x0 + ((v - sc.lo) / (sc.hi - sc.lo)) * (x1 - x0) : y1 - ((v - sc.lo) / (sc.hi - sc.lo)) * (y1 - y0));
  // 벽 · 바닥
  const wall = resolveColor(theme, c.plotBg ?? '@bg1') ?? '#fff';
  out.push(poly([[x0 + dx, y0 + dy], [x1 + dx, y0 + dy], [x1 + dx, y1 + dy], [x0 + dx, y1 + dy]], wall, ` stroke="${gridC}" stroke-width="0.6" data-el="plot"`));
  out.push(poly([[x0, y0], [x0 + dx, y0 + dy], [x0 + dx, y1 + dy], [x0, y1]], shadeHex(wall.startsWith('#') ? wall : '#ffffff', -0.04), ` stroke="${gridC}" stroke-width="0.6"`));
  out.push(poly([[x0, y1], [x1, y1], [x1 + dx, y1 + dy], [x0 + dx, y1 + dy]], shadeHex(wall.startsWith('#') ? wall : '#ffffff', -0.08), ` stroke="${axisC}" stroke-width="0.6"`));
  // 눈금선 (뒷벽 + 옆벽/바닥)
  const g = [];
  for (let v = sc.lo; v <= sc.hi + sc.step * 1e-6; v += sc.step) {
    const p = vpos(v);
    if (horiz) {
      g.push(`<line x1="${f(p)}" y1="${f(y1)}" x2="${f(p + dx)}" y2="${f(y1 + dy)}" stroke="${gridC}" stroke-width="0.7"/><line x1="${f(p + dx)}" y1="${f(y0 + dy)}" x2="${f(p + dx)}" y2="${f(y1 + dy)}" stroke="${gridC}" stroke-width="0.7"/>`);
      if (!c.axes?.y?.hide) g.push(text(p, box.bottom - fs * 0.35, valLbl(v), { size: fs * 0.9 }));
    } else {
      g.push(`<line x1="${f(x0)}" y1="${f(p)}" x2="${f(x0 + dx)}" y2="${f(p + dy)}" stroke="${gridC}" stroke-width="0.7"/><line x1="${f(x0 + dx)}" y1="${f(p + dy)}" x2="${f(x1 + dx)}" y2="${f(p + dy)}" stroke="${gridC}" stroke-width="0.7"/>`);
      if (!c.axes?.y?.hide) g.push(text(x0 - 6, p + fs * 0.32, valLbl(v), { size: fs * 0.9, anchor: 'end' }));
    }
  }
  out.push(`<g data-el="axisY">${g.join('')}</g>`);
  // 항목 이름
  const slot = (horiz ? y1 - y0 : x1 - x0) / n;
  const tilt = !horiz && Math.max(0, ...cats.map((t) => textWidth(String(t), fs * 0.9))) > slot * 0.95;
  const cl = cats.map((t, i) => (horiz
    ? text(x0 - 6, y0 + slot * (i + 0.5) + fs * 0.32, String(t), { size: fs * 0.9, anchor: 'end' })
    : tilt ? text(x0 + slot * (i + 0.5), y1 + fs * 0.8, String(t), { size: fs * 0.9, anchor: 'end', rot: -45 }) : text(x0 + slot * (i + 0.5), y1 + fs * 1.15, String(t), { size: fs * 0.9 })));
  out.push(`<g data-el="axisX">${cl.join('')}</g>`);
  const gap = (c.gap ?? 150) / 100;
  const lo2 = labelOpts(c, series[0], kind);
  if (kind === 'col' || kind === 'bar') {
    // 막대: 계열을 깊이로 (3차원 세로 막대형) 또는 나란히 (묶은) / 쌓기
    const per = full || stacked ? 1 : m;
    const bw = slot / (per + gap);
    const order = full ? [...series.keys()].reverse() : [...series.keys()];
    for (const si of order) {
      const col = seriesColor(theme, c, si);
      const fz = full ? (si + 0.15) / m : 0.15;
      const ddx = full ? (dx * 0.7) / m : dx * 0.7;
      const ddy = full ? (dy * 0.7) / m : dy * 0.7;
      cats.forEach((_, i) => {
        const b = base(si, i); const t = b + val(si, i);
        const k = full || stacked ? 0 : si;
        const pc = pointColor(theme, c, series[si], si, i, false) ?? col;
        if (horiz) {
          const yy = y0 + slot * i + (slot - bw * per) / 2 + bw * k + dy * fz;
          const xa = vpos(b) + dx * fz; const xb = vpos(t) + dx * fz;
          out.push(prism(Math.min(xa, xb), yy - ddy, Math.abs(xb - xa), bw, ddx, ddy, pc, `ser${si}`));
        } else {
          const xx = x0 + slot * i + (slot - bw * per) / 2 + bw * k + dx * fz;
          const ya = vpos(b) + dy * fz; const yb = vpos(t) + dy * fz;
          out.push(prism(xx, Math.min(ya, yb), bw, Math.abs(yb - ya), ddx, ddy, pc, `ser${si}`));
          if (lo2) out.push(text(xx + bw / 2 + ddx / 2, Math.min(ya, yb) + ddy - 3, fmtNum(vals[si][i], c.axes?.y?.fmt), { size: fs * 0.85 }));
        }
      });
    }
  } else {
    // 꺾은선 = 리본, 영역 = 깊이가 있는 면 (뒤 계열부터)
    const order = [...series.keys()].reverse();
    const xs = (i) => x0 + slot * (i + 0.5);
    for (const si of order) {
      const col = seriesColor(theme, c, si);
      const fz = full ? si / m : 0;
      const ddx = full ? (dx * 0.8) / m : dx * 0.8;
      const ddy = full ? (dy * 0.8) / m : dy * 0.8;
      const ox = dx * fz; const oy = dy * fz;
      const top = cats.map((_, i) => [xs(i) + ox, vpos(base(si, i) + val(si, i)) + oy]);
      const bot = cats.map((_, i) => [xs(i) + ox, vpos(base(si, i)) + oy]);
      if (kind === 'area') {
        // 윗면 (리본) → 앞면
        for (let i = 0; i < top.length - 1; i++) out.push(poly([top[i], top[i + 1], [top[i + 1][0] + ddx, top[i + 1][1] + ddy], [top[i][0] + ddx, top[i][1] + ddy]], shadeHex(col, 0.22), ` data-el="ser${si}"`));
        const last = top.length - 1;
        out.push(poly([top[last], [top[last][0] + ddx, top[last][1] + ddy], [bot[last][0] + ddx, bot[last][1] + ddy], bot[last]], shadeHex(col, -0.3)));
        out.push(poly([...top, ...bot.slice().reverse()], col, ` data-el="ser${si}"`));
      } else {
        const th = Math.max(3, fs * 0.35);
        for (let i = 0; i < top.length - 1; i++) {
          const a = top[i]; const b = top[i + 1];
          // 리본 윗면 + 앞면 (두께)
          out.push(poly([a, b, [b[0] + ddx, b[1] + ddy], [a[0] + ddx, a[1] + ddy]], shadeHex(col, 0.18), ` data-el="ser${si}"`));
          out.push(poly([a, b, [b[0], b[1] + th], [a[0], a[1] + th]], shadeHex(col, -0.2), ` data-el="ser${si}"`));
        }
      }
    }
  }
  return out.join('');
}

/** 3차원 원형: 기울인 타원 + 앞쪽 옆면 */
function pie3dSvg(c, series, cats, box, theme, { fs, text }) {
  const out = [];
  const s = series[0];
  if (!s) return '';
  const v3 = { ...VIEW3D_DEFAULT, rotX: 30, ...(c.view3d ?? {}) };
  const vals = (s.vals ?? []).map((v) => Math.max(0, numVal(v) ?? 0));
  const total = vals.reduce((a, b) => a + b, 0) || 1;
  const cx = (box.left + box.right) / 2;
  const R = Math.max(10, Math.min((box.right - box.left) / 2 - 8, (box.bottom - box.top) / 1.4));
  const tilt = Math.sin((Math.min(80, Math.max(10, v3.rotX)) * Math.PI) / 180);
  const ry = R * Math.max(0.2, tilt);
  const th = R * 0.22 * (v3.depth / 100) * Math.max(0.3, Math.cos((Math.min(80, v3.rotX) * Math.PI) / 180));
  const cy = (box.top + box.bottom) / 2 - th / 2;
  const explodeAll = (c.variant === 'explode' ? 15 : c.explode ?? 0) / 100;
  const P = (a, r = R, dy = 0, ox = 0, oy = 0) => [cx + ox + r * Math.cos(a), cy + oy + dy + (r * ry / R) * Math.sin(a)];
  const start = ((c.firstAngle ?? 0) - 90) * (Math.PI / 180);
  const slices = [];
  let a0 = start;
  vals.forEach((v, i) => { const a1 = a0 + (v / total) * Math.PI * 2; slices.push({ i, v, a0, a1 }); a0 = a1; });
  const stroke = resolveColor(theme, c.bg ?? '@bg1') ?? '#fff';
  const arcD = (a, b, dy, ox, oy) => {
    const steps = Math.max(2, Math.ceil(((b - a) * 24) / Math.PI));
    const pts = [];
    for (let k = 0; k <= steps; k++) pts.push(P(a + ((b - a) * k) / steps, R, dy, ox, oy));
    return pts;
  };
  // 옆면: 화면 앞쪽(사인 > 0) 부분만, 뒤 조각부터
  const sides = [];
  for (const sl of slices) {
    if (sl.v <= 0) continue;
    const am = (sl.a0 + sl.a1) / 2;
    const ex = ((s.explodeAt?.[sl.i] ?? 0) / 100 + explodeAll) * R;
    const ox = Math.cos(am) * ex; const oy = Math.sin(am) * ex * (ry / R);
    sl.ox = ox; sl.oy = oy;
    const col = pointColor(theme, c, s, 0, sl.i, true);
    sl.col = col;
    // [a0, a1] 와 앞쪽 구간 [2kπ, 2kπ+π] 의 겹침
    for (let k = -1; k <= 2; k++) {
      const lo = Math.max(sl.a0, 2 * Math.PI * k); const hi = Math.min(sl.a1, 2 * Math.PI * k + Math.PI);
      if (hi > lo) {
        const topPts = arcD(lo, hi, 0, ox, oy);
        const botPts = arcD(lo, hi, th, ox, oy).reverse();
        sides.push({ y: Math.max(...topPts.map((p) => p[1])), svg: poly([...topPts, ...botPts], shadeHex(col, -0.3), ` stroke="${shadeHex(col, -0.35)}" stroke-width="0.5" data-el="ser0"`) });
      }
    }
    // 자른 면 (조각을 쪼갰을 때 보이는 단면)
    if (ex > 0) for (const a of [sl.a0, sl.a1]) if (Math.cos(a) * (a === sl.a0 ? -1 : 1) > -2) sides.push({ y: cy + oy, svg: poly([[cx + ox, cy + oy], P(a, R, 0, ox, oy), P(a, R, th, ox, oy), [cx + ox, cy + oy + th]], shadeHex(col, -0.4)) });
  }
  sides.sort((a, b) => a.y - b.y);
  for (const sd of sides) out.push(sd.svg);
  // 윗면
  const lo = labelOpts(c, s, 'pie');
  for (const sl of slices) {
    if (sl.v <= 0) continue;
    const pts = [[cx + sl.ox, cy + sl.oy], ...arcD(sl.a0, sl.a1, 0, sl.ox, sl.oy)];
    out.push(poly(pts, sl.col, ` stroke="${stroke}" stroke-width="1" data-el="ser0"`));
    if (lo) {
      const am = (sl.a0 + sl.a1) / 2;
      const [lx, ly] = P(am, R * 0.62, 0, sl.ox, sl.oy);
      out.push(text(lx, ly + fs / 3, labelText(lo, sl.v, s, cats[sl.i], sl.v / total, c.axes?.y?.fmt), { size: fs * 0.9, color: '#fff' }));
    }
  }
  return out.join('');
}
