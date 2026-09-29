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
  { id: 'combo', label: '콤보 (막대 + 꺾은선 보조 축)' },
];

export const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

const isNum = (v) => typeof v === 'number' && Number.isFinite(v);
const label = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? formatGeneral(v) : String(v?.code ?? v));

/**
 * 2차원 값 → { categories, series:[{name, values}] }
 * 엑셀처럼: 첫 행/첫 열이 숫자가 아니면 머리글로, 행 수 ≥ 열 수이면 열 단위 계열
 */
export function chartLayout(rows, type = 'column') {
  const R = rows.length;
  const C = rows[0]?.length ?? 0;
  const headRow = R > 1 && rows[0].slice(C > 1 ? 1 : 0).every((v) => !isNum(v)) && rows[0].some((v) => v !== null && v !== '');
  const firstDataRow = headRow ? 1 : 0;
  const catCol = C > 1 && (type === 'scatter' || rows.slice(firstDataRow).every((r) => !isNum(r[0])));
  const firstDataCol = catCol ? 1 : 0;
  const byCols = type === 'scatter' || R - firstDataRow >= C - firstDataCol;
  return { R, C, headRow, catCol, firstDataRow, firstDataCol, byCols };
}

export function chartData(rows, type = 'column') {
  if (!rows.length || !rows[0]?.length) return { categories: [], series: [] };
  const { R, C, headRow, catCol, firstDataRow, firstDataCol, byCols } = chartLayout(rows, type);
  const series = [];
  let categories = [];
  if (byCols) {
    categories = rows.slice(firstDataRow).map((r, i) => (catCol ? label(r[0]) : String(i + 1)));
    for (let c = firstDataCol; c < C; c++) {
      series.push({
        name: headRow ? label(rows[0][c]) : `계열${c - firstDataCol + 1}`,
        values: rows.slice(firstDataRow).map((r) => (isNum(r[c]) ? r[c] : null)),
        x: type === 'scatter' && catCol ? rows.slice(firstDataRow).map((r) => (isNum(r[0]) ? r[0] : null)) : null,
      });
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
      return { name, values: vals.map((v) => (isNum(v) ? v : null)), x: xs ? xs.map((v) => (isNum(v) ? v : null)) : null };
    });
    const n = Math.max(0, ...series.map((s) => s.values.length));
    const categories = catRef ? flatRef(api.values(catRef), true).map(label) : Array.from({ length: n }, (_, i) => String(i + 1));
    base = { categories, series };
  } else base = chartData(api.range(ch), ch.type === 'combo' ? 'column' : ch.type);
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
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ${FONT}>`, '', 
    `<rect width="${W}" height="${H}" fill="#fff"/>`];
  let top = 10;
  if (chart.title) {
    parts.push(`<text x="${W / 2}" y="${Math.round(FS.title + 10)}" text-anchor="middle" font-size="${FS.title}" fill="#595959">${escSvg(truncate(chart.title, Math.floor(W / (FS.title * 0.62))))}</text>`);
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
  const series = data.series.map((s, i) => ({ ...s, type: s.type ?? baseType, axis: s.axis ?? 0, color: s.color ?? PALETTE[i % PALETTE.length] }));
  const pieLike = baseType === 'pie' || baseType === 'doughnut';
  const legendPos = chart.legend ?? 'b';
  const legendItems = pieLike ? categories.map((c, i) => ({ name: c, color: series[0]?.colors?.[i] ?? PALETTE[i % PALETTE.length], line: false }))
    : series.map((s) => ({ name: s.name, color: s.color, line: s.type === 'line' }));
  const showLegend = legendPos !== 'none' && (legendItems.length > 1 || pieLike);
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
        parts.push(key(it, lx0, ly), `<text x="${lx0 + 15}" y="${ly}" font-size="${FS.legend}" fill="#595959">${escSvg(truncate(String(it.name), Math.floor((legendW - 20) / (LW * 1.4))))}</text>`);
        ly += Math.round(FS.legend * 1.6);
      });
    } else {
      const itemW = Math.min(160, Math.max(60, (W - 20) / items.length));
      const total = itemW * items.length;
      let lx = Math.max(10, (W - total) / 2);
      const ly = legendPos === 't' ? top + 12 : H - 14;
      items.forEach((it) => {
        parts.push(key(it, lx, ly), `<text x="${lx + 15}" y="${ly}" font-size="${FS.legend}" fill="#595959">${escSvg(truncate(String(it.name), Math.floor((itemW - 18) / (LW * 1.4))))}</text>`);
        lx += itemW;
      });
    }
  }

  const wantLabels = (s) => s.labels ?? chart.labels ?? false;

  if (pieLike) {
    const s0 = series[0];
    const vals = s0.values.map((v) => (v && v > 0 ? v : 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    const cx = plot.x + plot.w / 2;
    const cy = plot.y + plot.h / 2;
    const r = Math.max(10, Math.min(plot.w, plot.h) / 2 - 6);
    const inner = baseType === 'doughnut' ? r * 0.5 : 0;
    let a = -Math.PI / 2;
    vals.forEach((v, i) => {
      if (!v) return;
      const frac = v / sum;
      const a2 = a + frac * Math.PI * 2;
      const color = s0.colors?.[i] ?? PALETTE[i % PALETTE.length];
      if (frac >= 0.9999) {
        parts.push(`<circle cx="${cx}" cy="${cy}" r="${r}" fill="${color}" stroke="#fff"/>`);
        if (inner) parts.push(`<circle cx="${cx}" cy="${cy}" r="${inner}" fill="#fff"/>`);
      } else {
        const large = a2 - a > Math.PI ? 1 : 0;
        const p = (ang, rad) => `${(cx + Math.cos(ang) * rad).toFixed(2)},${(cy + Math.sin(ang) * rad).toFixed(2)}`;
        const d = inner
          ? `M${p(a, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}L${p(a2, inner)}A${inner},${inner} 0 ${large} 0 ${p(a, inner)}Z`
          : `M${cx},${cy}L${p(a, r)}A${r},${r} 0 ${large} 1 ${p(a2, r)}Z`;
        parts.push(`<path d="${d}" fill="${color}" stroke="#fff" stroke-width="1.5"/>`);
      }
      // 레이블: false = 없음(파일에 없던 원형), pct = 백분율, labels = 값, 정하지 않음 = 백분율
      if (frac >= 0.04 && s0.labels !== false) {
        const mid = (a + a2) / 2;
        const lr = inner ? (r + inner) / 2 : r * 0.65;
        const txt = !s0.pct && wantLabels(s0) ? valueLabel(v, s0.numFmt) : `${Math.round(frac * 100)}%`;
        parts.push(`<text x="${(cx + Math.cos(mid) * lr).toFixed(1)}" y="${(cy + Math.sin(mid) * lr + 4).toFixed(1)}" text-anchor="middle" font-size="11" fill="#fff" font-weight="700">${escSvg(txt)}</text>`);
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
    const ss = series.filter((s) => s.axis === axis && s.type !== 'scatter');
    if (!ss.length) return null;
    let vals = ss.flatMap((s) => s.values).filter(isNum);
    const bars = ss.filter((s) => barTypes.has(s.type) || s.type === 'area');
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
    const ticks = [];
    for (let t = sc.min; t <= sc.max + sc.step / 2; t += sc.step) ticks.push(Number(t.toPrecision(12)));
    const srcCode = ss.map((s) => s.numFmt).find((f) => typeof f === 'string' && f !== 'General') ?? null;
    return { ...sc, ticks, code: pct ? '0%' : cfg.numFmt ?? srcCode };
  };
  const scale = scaleFor(0) ?? scaleFor(1);
  const scale2 = series.some((s) => s.axis === 1) && series.some((s) => s.axis === 0) ? scaleFor(1) : null;
  const primaryAxis = scaleFor(0) ? 0 : 1;
  const labelW = Math.min(100, Math.max(...scale.ticks.map((t) => axisLabel(t, scale.code).length)) * CW + 8);
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
      parts.push(`<line x1="${p}" y1="${area.y}" x2="${p}" y2="${area.y + area.h}" stroke="#d9d9d9"/>`,
        `<text x="${p}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="#595959">${escSvg(axisLabel(t, scale.code))}</text>`);
    } else {
      parts.push(`<line x1="${area.x}" y1="${p}" x2="${area.x + area.w}" y2="${p}" stroke="#d9d9d9"/>`,
        `<text x="${area.x - 5}" y="${Number(p) + FS.axis * 0.35}" text-anchor="end" font-size="${FS.axis}" fill="#595959">${escSvg(axisLabel(t, scale.code))}</text>`);
    }
  }
  if (scale2 && !horizontal) {
    for (const t of scale2.ticks) {
      const p = vpos2(t);
      parts.push(`<text x="${area.x + area.w + 5}" y="${(p + 3.5).toFixed(1)}" text-anchor="start" font-size="${FS.axis}" fill="#595959">${escSvg(axisLabel(t, scale2.code))}</text>`);
    }
  }
  const axisTitle = (txt, x, y, rot) => (txt ? `<text x="${x}" y="${y}" text-anchor="middle" font-size="11" fill="#595959"${rot ? ` transform="rotate(${rot} ${x} ${y})"` : ''}>${escSvg(txt)}</text>` : '');

  const base = vpos(Math.max(scale.min, Math.min(0, scale.max)));
  const labelsOut = [];
  const pushLabel = (x, y, v, s, anchor = 'middle') => {
    labelsOut.push(`<text x="${x.toFixed(1)}" y="${y.toFixed(1)}" text-anchor="${anchor}" font-size="${pt(s.labelSize ?? 9)}"${s.labelBold ? ' font-weight="700"' : ''} fill="${s.labelColor ?? '#404040'}" paint-order="stroke" stroke="#fff" stroke-width="2.5">${escSvg(valueLabel(v, s.numFmt))}</text>`);
  };

  if (baseType === 'scatter') {
    const xpos = (x) => area.x + ((x - xScale.min) / (xScale.max - xScale.min)) * area.w;
    for (let t = xScale.min; t <= xScale.max + xScale.step / 2; t += xScale.step) {
      parts.push(`<text x="${xpos(t).toFixed(1)}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="#595959">${escSvg(axisLabel(Number(t.toPrecision(12))))}</text>`);
    }
    series.forEach((s) => {
      s.values.forEach((v, i) => {
        const x = s.x ? s.x[i] : i + 1;
        if (!isNum(v) || !isNum(x)) return;
        parts.push(`<circle cx="${xpos(x).toFixed(1)}" cy="${vpos(v).toFixed(1)}" r="3.5" fill="${s.color}"/>`);
        if (wantLabels(s)) pushLabel(xpos(x), vpos(v) - 7, v, s);
      });
    });
  } else {
    // 항목 축 레이블
    const band = (horizontal ? area.h : area.w) / Math.max(1, n);
    const maxChars = Math.max(2, Math.floor(band / (CW * 1.4)));
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horizontal ? area.h : area.w) / 28))));
    categories.forEach((c, i) => {
      if (i % every) return;
      const mid = (horizontal ? area.y : area.x) + band * (i + 0.5);
      if (horizontal) parts.push(`<text x="${area.x - 5}" y="${(mid + 3.5).toFixed(1)}" text-anchor="end" font-size="${FS.axis}" fill="#595959">${escSvg(truncate(c, 16))}</text>`);
      else parts.push(`<text x="${mid.toFixed(1)}" y="${area.y + area.h + Math.round(FS.axis * 1.35)}" text-anchor="middle" font-size="${FS.axis}" fill="#595959">${escSvg(truncate(c, maxChars * every))}</text>`);
    });
    // 막대 (묶은 · 누적)
    const bars = series.filter((s) => barTypes.has(s.type));
    if (bars.length) {
      const k = stacked ? 1 : bars.length;
      const groupW = band * 0.65;
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
    for (const type of ['area', 'line']) {
      series.filter((s) => s.type === type).forEach((s) => {
        const vp = posOf(s);
        const pts = s.values.map((v, i) => (isNum(v) ? [area.x + band * (i + 0.5), vp(v), v] : null));
        const segs = [];
        let cur = [];
        for (const p of pts) { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } }
        if (cur.length) segs.push(cur);
        const b0 = vp(Math.max(scale.min, Math.min(0, scale.max)));
        for (const seg of segs) {
          const d = seg.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
          if (type === 'area') {
            parts.push(`<path d="${d}L${seg.at(-1)[0].toFixed(1)},${b0.toFixed(1)}L${seg[0][0].toFixed(1)},${b0.toFixed(1)}Z" fill="${s.color}" fill-opacity="0.75"/>`);
          } else {
            parts.push(`<path d="${d}" fill="none" stroke="${s.color}" stroke-width="${s.lineWidth ?? 2.25}" stroke-linejoin="round" stroke-linecap="round"${shadowAttr(s)}/>`);
            if (s.marker !== false && s.marker !== 'none') {
              const mk = MARKERS[s.marker] ?? MARKERS.circle;
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
