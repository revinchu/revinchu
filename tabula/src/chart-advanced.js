// 계층·표면·보조 원형·거래량 주식형. 값의 집계/보간은 DOM 없이 검증합니다.
import { formatGeneral } from './format.js';
import { chartView3D } from './chart-3d.js';

const finite = (v) => typeof v === 'number' && Number.isFinite(v);
const esc = (v) => String(v ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);
const xy = (p) => p.map((v) => Number(v.toFixed(3))).join(',');
const text = (x, y, value, color = '#404040', anchor = 'middle', size = 11) => `<text x="${x.toFixed(2)}" y="${y.toFixed(2)}" text-anchor="${anchor}" font-size="${size}" fill="${esc(color)}">${esc(value)}</text>`;
const tag = (s, i) => ` data-s="${s?._fi ?? 0}"${i === undefined ? '' : ` data-p="${i}"`}`;
const short = (v, n = 14) => [...String(v)].slice(0, n).join('') + ([...String(v)].length > n ? '…' : '');
const message = (ctx, s) => ctx.parts.push(text(ctx.plot.x + ctx.plot.w / 2, ctx.plot.y + ctx.plot.h / 2, s, ctx.TXT));
const domain = (values, cfg = {}, zero = false) => {
  let lo = Infinity, hi = -Infinity;
  for (const v of values) if (finite(v)) { lo = Math.min(lo, v); hi = Math.max(hi, v); }
  if (!finite(lo)) { lo = 0; hi = 1; }
  if (zero) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
  const min = finite(cfg.min) ? cfg.min : lo, max = finite(cfg.max) ? cfg.max : hi;
  if (max > min) return { min, max };
  const pad = Math.max(1, Math.abs(lo) * .05);
  return { min: zero ? Math.min(0, lo) : lo - pad, max: hi + pad };
};
function surfaceScale(data, chart) {
  const sc = domain(data.series.flatMap((s) => s.values), chart.axes?.y);
  let count = Math.max(2, Math.min(12, Math.round(chart.bandCount ?? 6)));
  if (chart.bandCount === undefined && !finite(chart.axes?.y?.min) && !finite(chart.axes?.y?.max)) {
    const raw = (sc.max - sc.min) / count, mag = 10 ** Math.floor(Math.log10(raw));
    const step = ([1, 2, 2.5, 5, 10].find((v) => v * mag >= raw) ?? 10) * mag;
    sc.min = Math.floor(sc.min / step) * step; sc.max = Math.ceil(sc.max / step) * step;
    count = Math.max(1, Math.round((sc.max - sc.min) / step));
  }
  return { ...sc, count, levels: Array.from({ length: count + 1 }, (_, i) => sc.min + (sc.max - sc.min) * i / count) };
}

/** 이름이 같아도 부모가 다르면 별도 항목입니다. 끝의 빈 단계는 짧은 가지로 유지합니다. */
export function sunburstNodes(data) {
  const root = { name: '', value: 0, own: 0, depth: 0, children: [], map: new Map(), points: [] };
  const levels = [...(data.catLevels ?? [])].reverse(), values = data.series[0]?.values ?? [];
  values.forEach((v, i) => {
    if (!finite(v) || v <= 0) return;
    const path = levels.map((spans) => spans.find((g) => i >= g.start && i <= g.end)?.text ?? '');
    path.push(data.categories[i] ?? '');
    const labels = path.filter((label) => String(label).trim() !== '');
    if (!labels.length) labels.push('(비어 있음)');
    let node = root; root.value += v;
    labels.forEach((label) => {
      const name = String(label);
      if (!node.map.has(name)) { const child = { name, value: 0, own: 0, depth: node.depth + 1, children: [], map: new Map(), points: [] }; node.map.set(name, child); node.children.push(child); }
      node = node.map.get(name); node.value += v; node.points.push(i);
    });
    node.own += v;
  });
  const out = [];
  const place = (node, start, span, rootIndex) => {
    node.start = start; node.end = start + span; node.rootIndex = rootIndex;
    if (node.depth) out.push(node);
    let angle = start;
    node.children.forEach((child, i) => { const arc = node.value ? span * child.value / node.value : 0; place(child, angle, arc, node.depth ? rootIndex : i); angle += arc; });
  };
  place(root, -Math.PI / 2, 2 * Math.PI, 0);
  return { nodes: out.map(({ map, children, ...node }) => node), total: root.value, depth: out.reduce((n, node) => Math.max(n, node.depth), 0) };
}

export function annularSector(cx, cy, inner, outer, start, end) {
  const p = (r, a) => xy([cx + r * Math.cos(a), cy + r * Math.sin(a)]), span = end - start;
  if (!(outer > inner) || !(span > 0)) return '';
  if (span >= Math.PI * 2 - 1e-8) {
    const ring = (r, dir) => `M${p(r, start)}A${r},${r} 0 1 ${dir} ${p(r, start + Math.PI)}A${r},${r} 0 1 ${dir} ${p(r, start)}Z`;
    return ring(outer, 1) + (inner ? ring(inner, 0) : '');
  }
  const large = span > Math.PI ? 1 : 0;
  return `M${p(outer, start)}A${outer},${outer} 0 ${large} 1 ${p(outer, end)}` + (inner ? `L${p(inner, end)}A${inner},${inner} 0 ${large} 0 ${p(inner, start)}Z` : `L${cx},${cy}Z`);
}

function drawSunburst(ctx) {
  const { data, series, plot, parts, pal, chart } = ctx, hierarchy = sunburstNodes(data), s = series[0];
  if (!hierarchy.total) return message(ctx, '선버스트에는 양수 값과 계층 항목이 필요합니다');
  const R = Math.max(1, Math.min(plot.w, plot.h) / 2 - 6), cx = plot.x + plot.w / 2, cy = plot.y + plot.h / 2, band = R / hierarchy.depth;
  for (const node of hierarchy.nodes) {
    const inner = (node.depth - 1) * band, outer = node.depth * band;
    const color = s.pointColors?.[node.points[0]] ?? pal[node.rootIndex % pal.length], shift = (chart.firstAngle ?? 0) * Math.PI / 180;
    const attrs = `${tag(s, node.points.length === 1 ? node.points[0] : undefined)} data-depth="${node.depth}" data-value="${node.value}"`;
    parts.push(`<path d="${annularSector(cx, cy, inner, outer, node.start + shift, node.end + shift)}" fill="${esc(color)}" fill-opacity="${Math.max(.5, 1 - (node.depth - 1) * .12)}" stroke="#fff" stroke-width="1"${attrs}><title>${esc(node.name)}: ${esc(formatGeneral(node.value))}</title></path>`);
    const r = inner ? (inner + outer) / 2 : outer * .63, span = node.end - node.start;
    if (chart.labels !== false && span * r > 24 && band > 13) {
      const angle = (node.start + node.end) / 2 + shift;
      parts.push(text(cx + Math.cos(angle) * r, cy + Math.sin(angle) * r + 4, short(node.name, Math.max(2, Math.floor(span * r / 7))), '#fff', 'middle', Math.min(11, band * .6)));
    }
  }
}

/** 위치=N은 마지막 N개. 값·백분율은 임곗값보다 작은 항목, custom은 원본 번호입니다. */
export function splitPieData(values, chart = {}) {
  const items = values.map((v, i) => ({ i, value: finite(v) ? Math.abs(v) : 0 })).filter((p) => p.value > 0);
  const total = items.reduce((sum, p) => sum + p.value, 0), type = chart.splitType ?? 'position', pos = finite(chart.splitPos) ? chart.splitPos : 3;
  const selected = new Set(chart.splitPoints ?? []);
  let secondary = items.filter((p, i) => type === 'custom' ? selected.has(p.i) : type === 'value' ? p.value < pos : type === 'percent' ? p.value / total * 100 < pos : i >= items.length - Math.min(items.length - 1, Math.max(0, Math.floor(pos))));
  // 주 원에는 원래 항목이 하나 이상 있어야 합니다. 모두를 고른 경우 가장 큰 항목을 주 원에 남깁니다.
  if (secondary.length === items.length && items.length > 1) { const largest = items.reduce((a, b) => b.value > a.value ? b : a); secondary = secondary.filter((p) => p !== largest); }
  if (items.length === 1) secondary = [];
  const set = new Set(secondary), primary = items.filter((p) => !set.has(p));
  const other = secondary.reduce((sum, p) => sum + p.value, 0);
  if (other) primary.push({ i: -1, value: other });
  return { primary, secondary, total, other };
}

function drawOfPie(ctx) {
  const { chart, series, categories, parts, plot, pal } = ctx, s = series[0], split = splitPieData(s.values, chart);
  if (!split.total) return message(ctx, '양수 또는 음수의 숫자 항목이 필요합니다');
  const ratio = Math.max(.1, Math.min(2, (chart.secondSize ?? 75) / 100));
  const gapFactor = Math.max(.1, Math.min(5, (chart.splitGap ?? 150) / 100));
  const r = Math.max(2, Math.min(plot.h / 2 - 15, (plot.w - 16) / (2 + 2 * ratio + gapFactor))), r2 = r * ratio;
  const width = 2 * r + 2 * r2 + r * gapFactor, x0 = plot.x + (plot.w - width) / 2;
  const cx = x0 + r, cx2 = x0 + 2 * r + r * gapFactor + r2, cy = plot.y + plot.h / 2;
  const color = (p) => p.i < 0 ? chart.otherColor ?? '#a5a5a5' : s.pointColors?.[p.i] ?? s.colors?.[p.i] ?? pal[p.i % pal.length];
  const slice = (p, x, radius, start, end) => {
    const title = p.i < 0 ? '기타' : categories[p.i] ?? String(p.i + 1);
    parts.push(`<path d="${annularSector(x, cy, 0, radius, start, end)}" fill="${esc(color(p))}" stroke="#fff" data-ofpie="${x === cx ? 'primary' : 'secondary'}" data-value="${p.value}"${tag(s, p.i < 0 ? undefined : p.i)}><title>${esc(title)}: ${formatGeneral(p.value)}</title></path>`);
    if (chart.labels !== false && (end - start) * radius > 32) { const a = (start + end) / 2; parts.push(text(x + Math.cos(a) * radius * .65, cy + Math.sin(a) * radius * .65 + 4, `${Math.round(p.value / split.total * 100)}%`, '#fff')); }
  };
  // 기타 조각의 가운데를 오른쪽으로 향하게 하여 두 차트의 연결 의미를 유지합니다.
  let a = split.other ? -Math.PI * split.other / split.total : -Math.PI / 2;
  const ordered = split.other ? [split.primary.at(-1), ...split.primary.slice(0, -1)] : split.primary;
  for (const p of ordered) { const end = a + p.value / split.total * Math.PI * 2; slice(p, cx, r, a, end); a = end; }
  if (!split.secondary.length) return;
  if (chart.type === 'barOfPie') {
    let y = cy + r2; const w = Math.max(8, r2 * .85);
    for (const p of split.secondary) { const h = 2 * r2 * p.value / split.other; y -= h; parts.push(`<rect x="${cx2 - w / 2}" y="${y}" width="${w}" height="${h}" fill="${esc(color(p))}" stroke="#fff" data-ofpie="secondary" data-value="${p.value}"${tag(s, p.i)}><title>${esc(categories[p.i])}: ${formatGeneral(p.value)}</title></rect>`); if (chart.labels !== false && h > 15) parts.push(text(cx2, y + h / 2 + 4, `${Math.round(p.value / split.total * 100)}%`, '#fff')); }
  } else { a = -Math.PI / 2; for (const p of split.secondary) { const end = a + p.value / split.other * Math.PI * 2; slice(p, cx2, r2, a, end); a = end; } }
  const angle = Math.PI * split.other / split.total, startX = cx + Math.cos(angle) * r;
  const endX = chart.type === 'barOfPie' ? cx2 - r2 * .425 : cx2;
  for (const sign of [-1, 1]) parts.push(`<line x1="${startX}" y1="${cy + sign * Math.sin(angle) * r}" x2="${endX}" y2="${cy + sign * r2}" stroke="#7f7f7f" data-ofpie="connector"/>`);
}

function clipZ(points, z, above) {
  const out = [];
  for (let i = 0; i < points.length; i++) {
    const a = points[i], b = points[(i + 1) % points.length], ai = above ? a.z >= z : a.z <= z, bi = above ? b.z >= z : b.z <= z;
    if (ai) out.push(a);
    if (ai !== bi) { const t = (z - a.z) / (b.z - a.z); out.push({ x: a.x + (b.x - a.x) * t, y: a.y + (b.y - a.y) * t, z }); }
  }
  return out;
}

/** 삼각 분할한 격자에서 값 범위 다각형과 같은 값의 등고선을 선형 보간합니다. */
export function surfaceGeometry(data, chart = {}) {
  const matrix = data.series.map((s) => s.values), rows = matrix.length, cols = data.categories.length;
  const sc = surfaceScale(data, chart), count = sc.count, step = (sc.max - sc.min) / count;
  const levels = sc.levels, faces = [], lines = [], mesh = [];
  const triangles = [];
  for (let y = 0; y + 1 < rows; y++) for (let x = 0; x + 1 < cols; x++) {
    const a = { x, y, z: matrix[y][x] }, b = { x: x + 1, y, z: matrix[y][x + 1] }, c = { x: x + 1, y: y + 1, z: matrix[y + 1][x + 1] }, d = { x, y: y + 1, z: matrix[y + 1][x] };
    if (![a, b, c, d].every((p) => finite(p.z))) continue;
    triangles.push({ points: [a, b, c], x, y }, { points: [a, c, d], x, y });
  }
  for (const t of triangles) {
    if (t.points.every((p) => p.z === t.points[0].z)) {
      const z = t.points[0].z;
      if (z >= sc.min && z <= sc.max) faces.push({ points: t.points, band: Math.min(count - 1, Math.floor((z - sc.min) / step)), x: t.x, y: t.y });
      continue;
    }
    for (let k = 0; k < count; k++) { const points = clipZ(clipZ(t.points, levels[k], true), levels[k + 1], false); if (points.length >= 3) faces.push({ points, band: k, x: t.x, y: t.y }); }
    for (const level of levels.slice(1, -1)) {
      const cuts = [];
      for (let i = 0; i < 3; i++) { const a = t.points[i], b = t.points[(i + 1) % 3]; if ((a.z < level && b.z >= level) || (b.z < level && a.z >= level)) { const ratio = (level - a.z) / (b.z - a.z); cuts.push({ x: a.x + (b.x - a.x) * ratio, y: a.y + (b.y - a.y) * ratio, z: level }); } }
      if (cuts.length === 2) lines.push({ points: cuts, value: level });
    }
  }
  for (let y = 0; y < rows; y++) for (let x = 0; x < cols; x++) {
    const a = { x, y, z: matrix[y][x] };
    for (const [nx, ny] of [[x + 1, y], [x, y + 1]]) if (nx < cols && ny < rows && finite(a.z) && finite(matrix[ny][nx])) mesh.push([a, { x: nx, y: ny, z: matrix[ny][nx] }]);
  }
  return { rows, cols, levels, faces, lines, mesh, min: sc.min, max: sc.max };
}

function surfaceLegend(data, chart, pal) {
  const { levels } = surfaceScale(data, chart);
  return levels.slice(0, -1).map((v, i) => ({ name: `${formatGeneral(Number(v.toPrecision(4)))}–${formatGeneral(Number(levels[i + 1].toPrecision(4)))}`, color: pal[i % pal.length] }));
}

function drawSurface(ctx) {
  const { chart, data, parts, plot, pal, TXT, FS } = ctx, g = surfaceGeometry(data, chart);
  if (g.rows < 2 || g.cols < 2 || !g.mesh.length) return message(ctx, '표면형에는 2행 × 2열 이상의 수치 격자가 필요합니다');
  const style = chart.surfaceStyle ?? (chart.threeD === false ? 'contour' : 'surface'), flat = /contour/i.test(style), wire = /wireframe/i.test(style);
  const area = { x: plot.x + 44, y: plot.y + 10, w: Math.max(10, plot.w - 76), h: Math.max(10, plot.h - 42) };
  const v = chartView3D(chart), angle = v.rotY * Math.PI / 180, tilt = Math.max(.15, Math.sin(Math.abs(v.rotX || 25) * Math.PI / 180));
  const projectRaw = (p) => { const x = p.x / (g.cols - 1) - .5, y = p.y / (g.rows - 1) - .5, z = (p.z - g.min) / (g.max - g.min); return [x * Math.cos(angle) - y * Math.sin(angle), (x * Math.sin(angle) + y * Math.cos(angle)) * tilt - z * .8]; };
  const corners = []; for (const x of [0, g.cols - 1]) for (const y of [0, g.rows - 1]) for (const z of [g.min, g.max]) corners.push(projectRaw({ x, y, z }));
  const xd = domain(corners.map((p) => p[0])), yd = domain(corners.map((p) => p[1]));
  const project = flat ? (p) => [area.x + p.x / (g.cols - 1) * area.w, area.y + area.h - p.y / (g.rows - 1) * area.h] : (p) => { const q = projectRaw(p); return [area.x + (q[0] - xd.min) / (xd.max - xd.min) * area.w, area.y + (q[1] - yd.min) / (yd.max - yd.min) * area.h]; };
  const depth = (p) => (p.x / (g.cols - 1) - .5) * Math.sin(angle) + (p.y / (g.rows - 1) - .5) * Math.cos(angle);
  parts.push(`<g data-surface="${esc(style)}"${flat ? '' : ' data-3d="surface"'}>`);
  const faceDepth = (f) => f.points.reduce((sum, p) => sum + depth(p), 0) / f.points.length;
  if (!wire) for (const face of g.faces.sort((a, b) => faceDepth(a) - faceDepth(b))) parts.push(`<polygon points="${face.points.map((p) => xy(project(p))).join(' ')}" fill="${esc(pal[face.band % pal.length])}" stroke="${esc(pal[face.band % pal.length])}" stroke-width="0.4" data-band="${face.band}" data-s="${data.series[face.y]?._fi ?? face.y}" data-p="${face.x}"/>`);
  const paths = flat ? g.lines.map((l) => l.points) : wire ? g.mesh : [];
  for (const segment of paths) parts.push(`<path d="M${segment.map((p) => xy(project(p))).join('L')}" fill="none" stroke="${wire ? esc(TXT) : '#ffffff'}" stroke-opacity="${wire ? 1 : .4}" stroke-width="${wire ? 1 : .6}" data-surface-line="${flat ? 'contour' : 'mesh'}"/>`);
  parts.push('</g>');
  const everyX = Math.max(1, Math.ceil(g.cols / 8)), everyY = Math.max(1, Math.ceil(g.rows / 6));
  data.categories.forEach((c, i) => { if (i % everyX) return; const p = project({ x: i, y: 0, z: g.min }); parts.push(text(p[0], p[1] + 16, short(c, 9), TXT, 'middle', FS.axis)); });
  data.series.forEach((s, i) => { if (i % everyY) return; const p = project({ x: 0, y: i, z: g.min }); parts.push(text(p[0] - 7, p[1] + 3, short(s.name, 9), TXT, 'end', FS.axis)); });
  if (!flat) for (let i = 0; i <= 4; i++) { const value = g.min + (g.max - g.min) * i / 4, p = project({ x: 0, y: 0, z: value }); parts.push(text(p[0] - 6, p[1] + 3, formatGeneral(Number(value.toPrecision(4))), TXT, 'end', FS.axis)); }
}

export function stockVolumeData(series, chart = {}) {
  const volume = !!chart.volume, offset = volume ? 1 : 0, ohlc = chart.ohlc ?? series.length - offset >= 4;
  const price = series.slice(offset), [open, high, low, close] = ohlc ? price : [null, ...price];
  return { volume: volume ? series[0] : null, open, high, low, close, ohlc, valid: !!high && !!low && !!close && (!ohlc || !!open) };
}

export function drawVolumeStock(ctx) {
  const { chart, series, categories, parts, plot, TXT, GRID, FS } = ctx, s = stockVolumeData(series, chart);
  if (!s.valid) return message(ctx, chart.ohlc ? '거래량·시가·고가·저가·종가 순서의 5개 계열이 필요합니다' : '거래량·고가·저가·종가 순서의 4개 계열이 필요합니다');
  const area = { x: plot.x + 50, y: plot.y + 8, w: Math.max(10, plot.w - 110), h: Math.max(10, plot.h - 34) }, band = area.w / Math.max(1, categories.length);
  const vd = domain(s.volume.values, chart.axes?.y, true), pd = domain([...(s.open?.values ?? []), ...s.high.values, ...s.low.values, ...s.close.values], chart.axes?.y2);
  const Y = (v, d, reverse) => area.y + area.h * (reverse ? (v - d.min) / (d.max - d.min) : 1 - (v - d.min) / (d.max - d.min));
  const V = (v) => Y(v, vd, chart.axes?.y?.reverse), P = (v) => Y(v, pd, chart.axes?.y2?.reverse);
  for (let i = 0; i <= 4; i++) {
    const vv = vd.min + (vd.max - vd.min) * i / 4, pv = pd.min + (pd.max - pd.min) * i / 4, y = V(vv);
    if (chart.gridY !== false) parts.push(`<line x1="${area.x}" y1="${y}" x2="${area.x + area.w}" y2="${y}" stroke="${esc(GRID)}"/>`);
    if (!chart.axes?.y?.hide) parts.push(text(area.x - 5, y + 4, formatGeneral(Number(vv.toPrecision(4))), TXT, 'end', FS.axis));
    if (!chart.axes?.y2?.hide) parts.push(text(area.x + area.w + 5, P(pv) + 4, formatGeneral(Number(pv.toPrecision(4))), TXT, 'start', FS.axis));
  }
  categories.forEach((c, i) => {
    const x = area.x + band * (i + .5), volume = s.volume.values[i], hi = s.high.values[i], lo = s.low.values[i], close = s.close.values[i], open = s.open?.values[i];
    if (finite(volume) && volume >= 0) { const base = V(Math.max(vd.min, Math.min(0, vd.max))), y = V(volume); parts.push(`<rect x="${x - band * .3}" y="${Math.min(y, base)}" width="${band * .6}" height="${Math.abs(base - y)}" fill="${esc(s.volume.color ?? '#4472c4')}" fill-opacity="0.42" data-stock="volume"${tag(s.volume, i)}/>`); }
    if (finite(hi) && finite(lo)) parts.push(`<line x1="${x}" y1="${P(hi)}" x2="${x}" y2="${P(lo)}" stroke="${esc(TXT)}" data-stock="highLow"${tag(s.high, i)}/>`);
    if (s.ohlc && finite(open) && finite(close)) { const w = Math.min(16, band * .45), y1 = P(open), y2 = P(close); parts.push(`<rect x="${x - w / 2}" y="${Math.min(y1, y2)}" width="${w}" height="${Math.max(1, Math.abs(y1 - y2))}" fill="${esc(close >= open ? chart.upColor ?? '#fff' : chart.downColor ?? '#404040')}" stroke="${esc(TXT)}" data-stock="openClose"${tag(s.close, i)}/>`); }
    else if (finite(close)) parts.push(`<line x1="${x}" y1="${P(close)}" x2="${x + Math.min(8, band * .3)}" y2="${P(close)}" stroke="${esc(TXT)}" stroke-width="2" data-stock="close"${tag(s.close, i)}/>`);
    if (i % Math.max(1, Math.ceil(categories.length / 10)) === 0) parts.push(text(x, area.y + area.h + 17, short(c, 8), TXT, 'middle', FS.axis));
  });
  parts.push(text(area.x, area.y - 2, chart.axes?.y?.title ?? s.volume.name ?? '거래량', TXT, 'start', 10), text(area.x + area.w, area.y - 2, chart.axes?.y2?.title ?? '가격', TXT, 'end', 10));
}

export const ADVANCED_CHARTS = {
  sunburst: { legend: false, draw: drawSunburst },
  surface: { legendItems: surfaceLegend, draw: drawSurface },
  pieOfPie: { draw: drawOfPie },
  barOfPie: { draw: drawOfPie },
};
