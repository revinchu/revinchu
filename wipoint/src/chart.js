// 차트 → SVG 문자열 (DOM 없음). chart = { kind, title, cats, series: [{name, vals, color?}], legend, labels, stacked }
import { resolveColor } from './themes.js';

export const CHART_KINDS = [['col', '세로 막대형'], ['bar', '가로 막대형'], ['line', '꺾은선형'], ['area', '영역형'], ['pie', '원형'], ['doughnut', '도넛형'], ['scatter', '분산형']];
export const CHART_LABEL = Object.fromEntries(CHART_KINDS);

const esc = (s) => String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
const f = (n) => Math.round(n * 10) / 10;
const ACC = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];

export function seriesColor(theme, chart, i) {
  const s = chart.series[i];
  if (s?.color) return resolveColor(theme, s.color);
  const base = resolveColor(theme, `@${ACC[i % 6]}`);
  return i < 6 ? base : resolveColor(theme, `@${ACC[i % 6]}:lm60`);
}

/** 보기 좋은 눈금 (최소, 최대, 간격) */
export function niceScale(min, max, ticks = 5) {
  if (!Number.isFinite(min) || !Number.isFinite(max)) return { lo: 0, hi: 1, step: 0.2 };
  if (min > 0) min = 0;
  if (max < 0) max = 0;
  if (min === max) max = min + 1;
  const raw = (max - min) / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => s >= raw) ?? 10 * mag;
  return { lo: Math.floor(min / step) * step, hi: Math.ceil(max / step) * step, step };
}
const fmt = (v) => {
  if (!Number.isFinite(v)) return '';
  const a = Math.abs(v);
  const s = a >= 1000 ? Math.round(v).toLocaleString('ko-KR') : String(Math.round(v * 100) / 100);
  return s;
};

export function chartSvg(chart, w, h, theme, { font = '맑은 고딕' } = {}) {
  const txt = resolveColor(theme, '@tx1:lm65:lo35') ?? '#595959';
  const grid = resolveColor(theme, '@tx1:lm15:lo85') ?? '#D9D9D9';
  const fs = Math.max(10, Math.min(16, Math.min(w, h) / 22));
  const parts = [];
  let top = 8;
  if (chart.title) {
    parts.push(`<text x="${f(w / 2)}" y="${f(top + fs * 1.4)}" text-anchor="middle" font-size="${f(fs * 1.4)}" fill="${txt}">${esc(chart.title)}</text>`);
    top += fs * 2.2;
  }
  const series = chart.series ?? [];
  const cats = chart.cats ?? [];
  const pie = chart.kind === 'pie' || chart.kind === 'doughnut';
  // 범례
  let bottom = h - 6;
  let right = w - 8;
  const legendItems = pie ? cats.map((c, i) => [c, resolveColor(theme, `@${ACC[i % 6]}`)]) : series.map((s, i) => [s.name, seriesColor(theme, chart, i)]);
  if (chart.legend && legendItems.length) {
    if (chart.legend === 'r') {
      const lw = Math.min(w * 0.3, 20 + Math.max(...legendItems.map(([n]) => String(n).length)) * fs * 0.75);
      right = w - lw - 8;
      legendItems.forEach(([n, c], i) => {
        const y = top + 10 + i * fs * 1.6;
        parts.push(`<rect x="${f(right + 12)}" y="${f(y - fs * 0.7)}" width="${f(fs * 0.7)}" height="${f(fs * 0.7)}" fill="${c}"/><text x="${f(right + 16 + fs)}" y="${f(y)}" font-size="${f(fs)}" fill="${txt}">${esc(n)}</text>`);
      });
    } else {
      bottom = h - fs * 2;
      const widths = legendItems.map(([n]) => fs * 1.4 + String(n).length * fs * 0.62 + 14);
      let x = (w - widths.reduce((a, b) => a + b, 0)) / 2;
      legendItems.forEach(([n, c], i) => {
        parts.push(`<rect x="${f(x)}" y="${f(h - fs * 1.35)}" width="${f(fs * 0.7)}" height="${f(fs * 0.7)}" fill="${c}"/><text x="${f(x + fs)}" y="${f(h - fs * 0.7)}" font-size="${f(fs)}" fill="${txt}">${esc(n)}</text>`);
        x += widths[i];
      });
    }
  }
  if (pie) {
    const vals = (series[0]?.vals ?? []).map((v) => Math.max(0, Number(v) || 0));
    const total = vals.reduce((a, b) => a + b, 0) || 1;
    const cx = (right + 8) / 2;
    const cy = (top + bottom) / 2;
    const r = Math.max(10, Math.min(right - 16, bottom - top - 8) / 2);
    let a0 = -Math.PI / 2;
    vals.forEach((v, i) => {
      const a1 = a0 + (v / total) * Math.PI * 2;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const c = resolveColor(theme, `@${ACC[i % 6]}`);
      if (v / total >= 0.9999) parts.push(`<circle cx="${f(cx)}" cy="${f(cy)}" r="${f(r)}" fill="${c}" stroke="#fff" stroke-width="1"/>`);
      else parts.push(`<path d="M${f(cx)} ${f(cy)} L${f(cx + r * Math.cos(a0))} ${f(cy + r * Math.sin(a0))} A${f(r)} ${f(r)} 0 ${large} 1 ${f(cx + r * Math.cos(a1))} ${f(cy + r * Math.sin(a1))} Z" fill="${c}" stroke="#fff" stroke-width="1.5"/>`);
      if (chart.labels && v > 0) {
        const am = (a0 + a1) / 2;
        parts.push(`<text x="${f(cx + r * 0.65 * Math.cos(am))}" y="${f(cy + r * 0.65 * Math.sin(am) + fs / 3)}" text-anchor="middle" font-size="${f(fs)}" fill="#fff">${Math.round((v / total) * 100)}%</text>`);
      }
      a0 = a1;
    });
    if (chart.kind === 'doughnut') parts.push(`<circle cx="${f(cx)}" cy="${f(cy)}" r="${f(r * 0.5)}" fill="${resolveColor(theme, '@bg1')}"/>`);
    return wrap(w, h, font, parts);
  }
  // 축 · 막대 · 선
  const horiz = chart.kind === 'bar';
  const stacked = !!chart.stacked;
  let lo = Infinity;
  let hi = -Infinity;
  if (stacked) {
    cats.forEach((_, ci) => {
      let p = 0; let n = 0;
      for (const s of series) { const v = Number(s.vals[ci]) || 0; if (v >= 0) p += v; else n += v; }
      lo = Math.min(lo, n); hi = Math.max(hi, p);
    });
  } else for (const s of series) for (const v of s.vals) { const x = Number(v); if (Number.isFinite(x)) { lo = Math.min(lo, x); hi = Math.max(hi, x); } }
  if (!Number.isFinite(lo)) { lo = 0; hi = 1; }
  const sc = niceScale(lo, hi);
  const labelW = horiz ? Math.min(w * 0.3, Math.max(...cats.map((c) => String(c).length), 1) * fs * 0.62 + 8) : Math.max(...[sc.lo, sc.hi].map((v) => fmt(v).length)) * fs * 0.62 + 10;
  const x0 = 8 + labelW;
  const x1 = right - 6;
  const y0 = top + 6;
  const y1 = bottom - fs * 1.8;
  const val2pos = (v) => (horiz ? x0 + ((v - sc.lo) / (sc.hi - sc.lo)) * (x1 - x0) : y1 - ((v - sc.lo) / (sc.hi - sc.lo)) * (y1 - y0));
  // 눈금선 + 값 레이블
  for (let v = sc.lo; v <= sc.hi + sc.step / 2; v += sc.step) {
    const p = val2pos(v);
    if (horiz) parts.push(`<line x1="${f(p)}" y1="${f(y0)}" x2="${f(p)}" y2="${f(y1)}" stroke="${grid}" stroke-width="1"/><text x="${f(p)}" y="${f(y1 + fs * 1.3)}" text-anchor="middle" font-size="${f(fs)}" fill="${txt}">${fmt(v)}</text>`);
    else parts.push(`<line x1="${f(x0)}" y1="${f(p)}" x2="${f(x1)}" y2="${f(p)}" stroke="${grid}" stroke-width="1"/><text x="${f(x0 - 6)}" y="${f(p + fs / 3)}" text-anchor="end" font-size="${f(fs)}" fill="${txt}">${fmt(v)}</text>`);
  }
  const n = Math.max(1, cats.length);
  const band = (horiz ? y1 - y0 : x1 - x0) / n;
  const zero = val2pos(0);
  cats.forEach((c, ci) => {
    const mid = (horiz ? y0 : x0) + band * (ci + 0.5);
    if (horiz) parts.push(`<text x="${f(x0 - 6)}" y="${f(mid + fs / 3)}" text-anchor="end" font-size="${f(fs)}" fill="${txt}">${esc(c)}</text>`);
    else parts.push(`<text x="${f(mid)}" y="${f(y1 + fs * 1.3)}" text-anchor="middle" font-size="${f(fs)}" fill="${txt}">${esc(c)}</text>`);
  });
  if (chart.kind === 'col' || chart.kind === 'bar') {
    const gap = band * 0.27;
    const bw = stacked ? band - gap * 2 : (band - gap * 2) / Math.max(1, series.length);
    cats.forEach((_, ci) => {
      let pos = 0;
      let neg = 0;
      series.forEach((s, si) => {
        const v = Number(s.vals[ci]) || 0;
        const c = seriesColor(theme, chart, si);
        let a;
        let b;
        if (stacked) { const base = v >= 0 ? pos : neg; a = val2pos(base); b = val2pos(base + v); if (v >= 0) pos += v; else neg += v; } else { a = zero; b = val2pos(v); }
        const start = (horiz ? y0 : x0) + band * ci + gap + (stacked ? 0 : bw * si);
        if (horiz) parts.push(`<rect x="${f(Math.min(a, b))}" y="${f(start)}" width="${f(Math.abs(b - a))}" height="${f(bw)}" fill="${c}"/>`);
        else parts.push(`<rect x="${f(start)}" y="${f(Math.min(a, b))}" width="${f(bw)}" height="${f(Math.abs(b - a))}" fill="${c}"/>`);
        if (chart.labels) {
          if (horiz) parts.push(`<text x="${f(Math.max(a, b) + 4)}" y="${f(start + bw / 2 + fs / 3)}" font-size="${f(fs * 0.9)}" fill="${txt}">${fmt(v)}</text>`);
          else parts.push(`<text x="${f(start + bw / 2)}" y="${f(Math.min(a, b) - 4)}" text-anchor="middle" font-size="${f(fs * 0.9)}" fill="${txt}">${fmt(v)}</text>`);
        }
      });
    });
  } else {
    // 꺾은선 · 영역 · 분산
    series.forEach((s, si) => {
      const c = seriesColor(theme, chart, si);
      const pts = s.vals.map((v, ci) => [x0 + band * (ci + 0.5), val2pos(Number(v) || 0)]);
      if (!pts.length) return;
      const d = pts.map(([x, y], i) => `${i ? 'L' : 'M'}${f(x)} ${f(y)}`).join(' ');
      if (chart.kind === 'area') parts.push(`<path d="${d} L${f(pts[pts.length - 1][0])} ${f(zero)} L${f(pts[0][0])} ${f(zero)} Z" fill="${c}" fill-opacity="0.75"/>`);
      else if (chart.kind !== 'scatter') parts.push(`<path d="${d}" fill="none" stroke="${c}" stroke-width="${f(Math.max(2, fs / 5))}" stroke-linejoin="round" stroke-linecap="round"/>`);
      if (chart.kind === 'scatter' || chart.markers) for (const [x, y] of pts) parts.push(`<circle cx="${f(x)}" cy="${f(y)}" r="${f(fs / 3)}" fill="${c}"/>`);
      if (chart.labels) s.vals.forEach((v, ci) => parts.push(`<text x="${f(pts[ci][0])}" y="${f(pts[ci][1] - 6)}" text-anchor="middle" font-size="${f(fs * 0.9)}" fill="${txt}">${fmt(Number(v))}</text>`));
    });
  }
  parts.push(horiz ? `<line x1="${f(zero)}" y1="${f(y0)}" x2="${f(zero)}" y2="${f(y1)}" stroke="${grid}" stroke-width="1.2"/>` : `<line x1="${f(x0)}" y1="${f(zero)}" x2="${f(x1)}" y2="${f(zero)}" stroke="${grid}" stroke-width="1.2"/>`);
  return wrap(w, h, font, parts);
}

function wrap(w, h, font, parts) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${f(w)}" height="${f(h)}" viewBox="0 0 ${f(w)} ${f(h)}" font-family="${esc(font)}, 'Malgun Gothic', sans-serif">${parts.join('')}</svg>`;
}

/** 차트 데이터 ↔ 표 텍스트 (탭 구분, 첫 행 = 계열 이름, 첫 열 = 항목) */
export function chartToTsv(chart) {
  const rows = [['', ...chart.series.map((s) => s.name)]];
  chart.cats.forEach((c, i) => rows.push([c, ...chart.series.map((s) => s.vals[i] ?? '')]));
  return rows.map((r) => r.join('\t')).join('\n');
}
export function tsvToChart(chart, tsv) {
  const text = String(tsv).replace(/\r/g, '');
  const sep = text.includes('\t') ? '\t' : ',';
  const rows = text.split('\n').filter((l) => l.trim() !== '').map((l) => l.split(sep));
  if (rows.length < 2) throw new Error('항목 행이 하나 이상 있어야 합니다');
  const head = rows[0].slice(1);
  const out = { ...chart, cats: rows.slice(1).map((r) => r[0] ?? ''), series: head.map((name, si) => ({ ...(chart.series[si] ?? {}), name: name.trim() || `계열 ${si + 1}`, vals: rows.slice(1).map((r) => { const v = Number(String(r[si + 1] ?? '').replace(/,/g, '')); return Number.isFinite(v) ? v : 0; }) })) };
  return out;
}
