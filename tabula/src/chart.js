// 차트: 범위 값 → 데이터 해석, SVG 그리기 (DOM 없이 문자열 생성)
import { formatGeneral } from './format.js';

export const CHART_TYPES = [
  { id: 'column', label: '세로 막대형' },
  { id: 'bar', label: '가로 막대형' },
  { id: 'line', label: '꺾은선형' },
  { id: 'area', label: '영역형' },
  { id: 'pie', label: '원형' },
  { id: 'doughnut', label: '도넛형' },
  { id: 'scatter', label: '분산형' },
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

function axisLabel(n) {
  const a = Math.abs(n);
  if (a >= 1e6 && a < 1e15) return `${formatGeneral(Number((n / 1e6).toPrecision(4)))}M`;
  if (Number.isInteger(n)) return n.toLocaleString('en-US');
  return formatGeneral(Number(n.toPrecision(6)));
}

function truncate(s, maxChars) {
  return [...s].length > maxChars ? `${[...s].slice(0, Math.max(1, maxChars - 1)).join('')}…` : s;
}

/**
 * chart: { type, title, w, h }, data: chartData 결과 → SVG 문자열
 */
export function renderChartSvg(chart, data) {
  const W = Math.max(120, chart.w);
  const H = Math.max(90, chart.h);
  const FONT = "font-family=\"'Malgun Gothic','Apple SD Gothic Neo','Noto Sans KR',sans-serif\"";
  const parts = [`<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}" viewBox="0 0 ${W} ${H}" ${FONT}>`,
    `<rect width="${W}" height="${H}" fill="#fff"/>`];
  let top = 10;
  if (chart.title) {
    parts.push(`<text x="${W / 2}" y="28" text-anchor="middle" font-size="16" fill="#595959">${escSvg(truncate(chart.title, Math.floor(W / 11)))}</text>`);
    top = 42;
  }
  const { categories, series } = data;
  const pieLike = chart.type === 'pie' || chart.type === 'doughnut';
  const legendItems = pieLike ? categories : series.map((s) => s.name);
  const showLegend = legendItems.length > 1 || pieLike;
  const legendH = showLegend ? 26 : 0;
  const plot = { x: 10, y: top, w: W - 20, h: H - top - 10 - legendH };

  if (!series.length || series.every((s) => s.values.every((v) => v === null))) {
    parts.push(`<text x="${W / 2}" y="${H / 2}" text-anchor="middle" font-size="12" fill="#999">표시할 숫자 데이터가 없습니다</text></svg>`);
    return parts.join('');
  }

  if (showLegend) {
    const items = legendItems.slice(0, 12);
    const itemW = Math.min(140, Math.max(60, (W - 20) / items.length));
    const total = itemW * items.length;
    let lx = Math.max(10, (W - total) / 2);
    const ly = H - 14;
    items.forEach((name, i) => {
      const color = PALETTE[i % PALETTE.length];
      parts.push(`<rect x="${lx}" y="${ly - 8}" width="9" height="9" fill="${color}"/>`,
        `<text x="${lx + 13}" y="${ly}" font-size="11" fill="#595959">${escSvg(truncate(name, Math.floor((itemW - 16) / 7)))}</text>`);
      lx += itemW;
    });
  }

  if (pieLike) {
    const vals = series[0].values.map((v) => (v && v > 0 ? v : 0));
    const sum = vals.reduce((a, b) => a + b, 0);
    const cx = plot.x + plot.w / 2;
    const cy = plot.y + plot.h / 2;
    const r = Math.max(10, Math.min(plot.w, plot.h) / 2 - 6);
    const inner = chart.type === 'doughnut' ? r * 0.5 : 0;
    let a = -Math.PI / 2;
    vals.forEach((v, i) => {
      if (!v) return;
      const frac = v / sum;
      const a2 = a + frac * Math.PI * 2;
      const color = PALETTE[i % PALETTE.length];
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
      if (frac >= 0.04) {
        const mid = (a + a2) / 2;
        const lr = inner ? (r + inner) / 2 : r * 0.65;
        parts.push(`<text x="${(cx + Math.cos(mid) * lr).toFixed(1)}" y="${(cy + Math.sin(mid) * lr + 4).toFixed(1)}" text-anchor="middle" font-size="11" fill="#fff" font-weight="700">${Math.round(frac * 100)}%</text>`);
      }
      a = a2;
    });
    parts.push('</svg>');
    return parts.join('');
  }

  // 값 범위
  const all = series.flatMap((s) => s.values).filter(isNum);
  let min = Math.min(0, ...all);
  let max = Math.max(0, ...all);
  if (chart.type === 'scatter' || chart.type === 'line') { min = Math.min(...all); max = Math.max(...all); if (min > 0 && min < max * 0.5) min = 0; }
  const scale = niceScale(min, max);
  const ticks = [];
  for (let t = scale.min; t <= scale.max + scale.step / 2; t += scale.step) ticks.push(Number(t.toPrecision(12)));
  const horizontal = chart.type === 'bar';
  const labelW = Math.min(90, Math.max(...ticks.map((t) => axisLabel(t).length)) * 7 + 8);

  // 가로축 값 범위 (분산형)
  let xScale = null;
  if (chart.type === 'scatter') {
    const xs = series.flatMap((s) => (s.x ?? s.values.map((_, i) => i + 1))).filter(isNum);
    xScale = niceScale(Math.min(...xs), Math.max(...xs));
  }

  const catLabelW = horizontal ? Math.min(100, Math.max(...categories.map((c) => [...c].length)) * 7 + 8) : 0;
  const area = horizontal
    ? { x: plot.x + catLabelW, y: plot.y, w: plot.w - catLabelW - 10, h: plot.h - 18 }
    : { x: plot.x + labelW, y: plot.y + 4, w: plot.w - labelW - 6, h: plot.h - 22 };
  if (area.w < 20 || area.h < 20) { parts.push('</svg>'); return parts.join(''); }
  const vpos = (v) => (horizontal
    ? area.x + ((v - scale.min) / (scale.max - scale.min)) * area.w
    : area.y + area.h - ((v - scale.min) / (scale.max - scale.min)) * area.h);

  // 눈금선 + 값 축 레이블
  for (const t of ticks) {
    const p = vpos(t).toFixed(1);
    if (horizontal) {
      parts.push(`<line x1="${p}" y1="${area.y}" x2="${p}" y2="${area.y + area.h}" stroke="#d9d9d9"/>`,
        `<text x="${p}" y="${area.y + area.h + 14}" text-anchor="middle" font-size="10" fill="#595959">${escSvg(axisLabel(t))}</text>`);
    } else {
      parts.push(`<line x1="${area.x}" y1="${p}" x2="${area.x + area.w}" y2="${p}" stroke="#d9d9d9"/>`,
        `<text x="${area.x - 5}" y="${Number(p) + 3.5}" text-anchor="end" font-size="10" fill="#595959">${escSvg(axisLabel(t))}</text>`);
    }
  }

  const n = categories.length;
  const base = vpos(Math.max(scale.min, Math.min(0, scale.max)));

  if (chart.type === 'scatter') {
    const xpos = (x) => area.x + ((x - xScale.min) / (xScale.max - xScale.min)) * area.w;
    for (let t = xScale.min; t <= xScale.max + xScale.step / 2; t += xScale.step) {
      parts.push(`<text x="${xpos(t).toFixed(1)}" y="${area.y + area.h + 14}" text-anchor="middle" font-size="10" fill="#595959">${escSvg(axisLabel(Number(t.toPrecision(12))))}</text>`);
    }
    series.forEach((s, si) => {
      const color = PALETTE[si % PALETTE.length];
      s.values.forEach((v, i) => {
        const x = s.x ? s.x[i] : i + 1;
        if (!isNum(v) || !isNum(x)) return;
        parts.push(`<circle cx="${xpos(x).toFixed(1)}" cy="${vpos(v).toFixed(1)}" r="3.5" fill="${color}"/>`);
      });
    });
  } else {
    // 항목 축 레이블
    const band = (horizontal ? area.h : area.w) / Math.max(1, n);
    const maxChars = Math.max(2, Math.floor(band / 7));
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horizontal ? area.h : area.w) / 28))));
    categories.forEach((c, i) => {
      if (i % every) return;
      const mid = (horizontal ? area.y : area.x) + band * (i + 0.5);
      if (horizontal) parts.push(`<text x="${area.x - 5}" y="${(mid + 3.5).toFixed(1)}" text-anchor="end" font-size="10" fill="#595959">${escSvg(truncate(c, 14))}</text>`);
      else parts.push(`<text x="${mid.toFixed(1)}" y="${area.y + area.h + 14}" text-anchor="middle" font-size="10" fill="#595959">${escSvg(truncate(c, maxChars * every))}</text>`);
    });
    if (chart.type === 'column' || chart.type === 'bar') {
      const k = series.length;
      const groupW = band * 0.7;
      const barW = groupW / k;
      series.forEach((s, si) => {
        const color = PALETTE[si % PALETTE.length];
        s.values.forEach((v, i) => {
          if (!isNum(v)) return;
          const start = (horizontal ? area.y : area.x) + band * i + (band - groupW) / 2 + barW * si;
          const p = vpos(v);
          const a = Math.min(p, base);
          const len = Math.abs(p - base);
          if (horizontal) parts.push(`<rect x="${a.toFixed(1)}" y="${start.toFixed(1)}" width="${len.toFixed(1)}" height="${Math.max(1, barW - 1).toFixed(1)}" fill="${color}"/>`);
          else parts.push(`<rect x="${start.toFixed(1)}" y="${a.toFixed(1)}" width="${Math.max(1, barW - 1).toFixed(1)}" height="${len.toFixed(1)}" fill="${color}"/>`);
        });
      });
    } else {
      series.forEach((s, si) => {
        const color = PALETTE[si % PALETTE.length];
        const pts = s.values.map((v, i) => (isNum(v) ? [area.x + band * (i + 0.5), vpos(v)] : null));
        const segs = [];
        let cur = [];
        for (const p of pts) { if (p) cur.push(p); else if (cur.length) { segs.push(cur); cur = []; } }
        if (cur.length) segs.push(cur);
        for (const seg of segs) {
          const d = seg.map((p, i) => `${i ? 'L' : 'M'}${p[0].toFixed(1)},${p[1].toFixed(1)}`).join('');
          if (chart.type === 'area') {
            parts.push(`<path d="${d}L${seg.at(-1)[0].toFixed(1)},${base.toFixed(1)}L${seg[0][0].toFixed(1)},${base.toFixed(1)}Z" fill="${color}" fill-opacity="0.75"/>`);
          } else {
            parts.push(`<path d="${d}" fill="none" stroke="${color}" stroke-width="2.25" stroke-linejoin="round"/>`);
            for (const p of seg) parts.push(`<circle cx="${p[0].toFixed(1)}" cy="${p[1].toFixed(1)}" r="3" fill="${color}"/>`);
          }
        }
      });
    }
  }
  // 기준선
  if (horizontal) parts.push(`<line x1="${base}" y1="${area.y}" x2="${base}" y2="${area.y + area.h}" stroke="#bfbfbf"/>`);
  else parts.push(`<line x1="${area.x}" y1="${base}" x2="${area.x + area.w}" y2="${base}" stroke="#bfbfbf"/>`);
  parts.push('</svg>');
  return parts.join('');
}
