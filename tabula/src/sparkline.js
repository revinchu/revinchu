// 스파크라인 (DOM 없음): 셀 안의 작은 차트 — 꺾은선형 · 열 · 승패
// 시트 속성 sparklines = [{ id, type: 'line'|'column'|'winloss', color, negColor, markerColor, highColor, lowColor, firstColor, lastColor,
//   markers, high, low, first, last, negative, weight, items: [{ r, c, ref: 'Sheet1!A2:F2' }] }]  (그룹 하나에 여러 셀)
import { parse, isError } from './formula.js';
import { maxOf, minOf } from './fxcore.js';

export const SPARK_TYPES = [
  { id: 'line', label: '꺾은선형' },
  { id: 'column', label: '열' },
  { id: 'winloss', label: '승패' },
];

export const sparkDefaults = (type) => ({
  type, color: '#376092', negColor: '#d00000', markerColor: '#d00000', highColor: '#d00000', lowColor: '#d00000',
  firstColor: '#d00000', lastColor: '#d00000', markers: false, high: false, low: false, first: false, last: false, negative: type !== 'line', weight: 1.25,
});

/** 참조 글자 → { sheet, r1, c1, r2, c2 } (없으면 null) */
export function sparkRef(ref) {
  try {
    const ast = parse(ref);
    if (ast.type !== 'ref') return null;
    const f = ast.ref;
    return { sheet: f.sheet ?? null, r1: f.r1, c1: f.c1, r2: f.r2, c2: f.c2 };
  } catch { return null; }
}

/** 스파크라인 데이터: 범위의 값을 한 줄로 (숫자가 아니면 null = 빈 칸) */
export function sparkValues(wb, si, ref) {
  const rg = sparkRef(ref);
  if (!rg) return [];
  const s = rg.sheet ? wb.sheetIndexByName(rg.sheet) : si;
  if (s < 0) return [];
  const out = [];
  const n = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
  if (n > 5000) return out;
  for (let r = rg.r1; r <= rg.r2; r++) {
    for (let c = rg.c1; c <= rg.c2; c++) {
      const v = wb.getValue(s, r, c);
      out.push(typeof v === 'number' && !isError(v) ? v : null);
    }
  }
  return out;
}

/** 스파크라인 SVG 내용 (w × h 픽셀, 여백 2px) */
export function sparkSvg(values, g, w, h) {
  const pad = 2;
  const W = Math.max(1, w - pad * 2);
  const H = Math.max(1, h - pad * 2);
  const nums = values.filter((v) => v !== null);
  if (!nums.length) return '';
  const n = values.length;
  let min = minOf(nums);
  let max = maxOf(nums);
  const out = [];
  const idxOf = (want) => values.findIndex((v) => v === want);
  const firstI = values.findIndex((v) => v !== null);
  let lastI = -1;
  for (let i = n - 1; i >= 0; i--) if (values[i] !== null) { lastI = i; break; }
  const pointColor = (i, v) => {
    if (g.high && v === max && i === idxOf(max)) return g.highColor;
    if (g.low && v === min && i === idxOf(min)) return g.lowColor;
    if (g.first && i === firstI) return g.firstColor;
    if (g.last && i === lastI) return g.lastColor;
    if (g.negative && v < 0) return g.negColor;
    return null;
  };
  if (g.type === 'line') {
    if (max === min) { max += 1; min -= 1; }
    const x = (i) => pad + (n === 1 ? W / 2 : (i * W) / (n - 1));
    const y = (v) => pad + H - ((v - min) / (max - min)) * H;
    const pts = [];
    values.forEach((v, i) => {
      if (v === null) { if (pts.length) { out.push(`<polyline points="${pts.join(' ')}" fill="none" stroke="${g.color}" stroke-width="${g.weight ?? 1.25}" stroke-linejoin="round"/>`); pts.length = 0; } return; }
      pts.push(`${x(i).toFixed(1)},${y(v).toFixed(1)}`);
    });
    if (pts.length) out.push(`<polyline points="${pts.join(' ')}" fill="none" stroke="${g.color}" stroke-width="${g.weight ?? 1.25}" stroke-linejoin="round"/>`);
    values.forEach((v, i) => {
      if (v === null) return;
      const c = pointColor(i, v) ?? (g.markers ? g.markerColor : null);
      if (c) out.push(`<circle cx="${x(i).toFixed(1)}" cy="${y(v).toFixed(1)}" r="1.8" fill="${c}"/>`);
    });
    return out.join('');
  }
  // 열 · 승패: 0 을 기준으로 위아래 막대
  const bw = W / n;
  const gap = Math.min(2, bw * 0.2);
  if (g.type === 'winloss') {
    const mid = pad + H / 2;
    values.forEach((v, i) => {
      if (v === null || v === 0) return;
      const c = pointColor(i, v) ?? (v < 0 && g.negative !== false ? g.negColor : g.color);
      out.push(`<rect x="${(pad + i * bw + gap / 2).toFixed(1)}" y="${(v > 0 ? pad : mid).toFixed(1)}" width="${Math.max(0.5, bw - gap).toFixed(1)}" height="${(H / 2 - 0.5).toFixed(1)}" fill="${c}"/>`);
    });
    return out.join('');
  }
  const lo = Math.min(0, min);
  const hi = Math.max(0, max);
  const span = hi - lo || 1;
  const yOf = (v) => pad + H - ((v - lo) / span) * H;
  const zero = yOf(0);
  values.forEach((v, i) => {
    if (v === null) return;
    const c = pointColor(i, v) ?? g.color;
    const y1 = Math.min(yOf(v), zero);
    const hh = Math.max(0.8, Math.abs(yOf(v) - zero));
    out.push(`<rect x="${(pad + i * bw + gap / 2).toFixed(1)}" y="${y1.toFixed(1)}" width="${Math.max(0.5, bw - gap).toFixed(1)}" height="${hh.toFixed(1)}" fill="${c}"/>`);
  });
  return out.join('');
}

/** 데이터 범위 · 위치 범위 → 항목 목록 (위치가 한 열이면 행마다, 한 행이면 열마다) */
export function sparkItems(data, loc, sheetPrefix = '') {
  const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };
  const a1 = (r, c) => `${colName(c)}${r + 1}`;
  const rows = data.r2 - data.r1 + 1;
  const cols = data.c2 - data.c1 + 1;
  const locN = (loc.r2 - loc.r1 + 1) * (loc.c2 - loc.c1 + 1);
  if (locN === 1) return [{ r: loc.r1, c: loc.c1, ref: `${sheetPrefix}${a1(data.r1, data.c1)}:${a1(data.r2, data.c2)}` }];
  if (loc.c1 === loc.c2 && loc.r2 - loc.r1 + 1 === rows) {
    return Array.from({ length: rows }, (_, i) => ({ r: loc.r1 + i, c: loc.c1, ref: `${sheetPrefix}${a1(data.r1 + i, data.c1)}:${a1(data.r1 + i, data.c2)}` }));
  }
  if (loc.r1 === loc.r2 && loc.c2 - loc.c1 + 1 === cols) {
    return Array.from({ length: cols }, (_, i) => ({ r: loc.r1, c: loc.c1 + i, ref: `${sheetPrefix}${a1(data.r1, data.c1 + i)}:${a1(data.r2, data.c1 + i)}` }));
  }
  return null;
}
