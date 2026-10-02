// 계층 차트의 원본 항목 번호와 부모 경로를 함께 보존합니다. DOM/차트 모듈에 의존하지 않습니다.
import { formatCode, formatGeneral } from './format.js';

const number = (v) => typeof v === 'number' && Number.isFinite(v);
const name = (v) => v === null || v === undefined ? '' : String(v);
const esc = (v) => name(v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&apos;' })[c]);

/** 여러 열은 바깥→안쪽 순서. 중간 빈 부모는 같은 상위 경로의 직전 부모만 이어받습니다. */
export function hierarchyCategories(rows, count = rows?.length ?? 0) {
  if (!rows?.length) return null;
  let grid = rows;
  if (rows.length !== count && rows[0]?.length === count) grid = rows[0].map((_, c) => rows.map((r) => r[c]));
  const depth = grid[0]?.length ?? 0;
  if (!depth) return null;
  let previous = [];
  const paths = grid.map((row) => {
    const path = Array.from({ length: depth }, (_, d) => name(row[d]));
    let end = depth - 1;
    while (end > 0 && path[end] === '') end--;
    for (let d = 0; d < end; d++) {
      if (path[d] === '' && path.slice(0, d).every((v, k) => v === previous[k])) path[d] = previous[d] ?? '';
    }
    previous = path;
    return path;
  });
  return { categories: paths.map((p) => p[depth - 1]), levels: levelsFromPaths(paths, depth), paths };
}

function levelsFromPaths(paths, depth) {
  const levels = [];
  for (let d = depth - 2; d >= 0; d--) {
    const spans = [];
    paths.forEach((p, i) => {
      const last = spans.at(-1), previous = paths[i - 1];
      if (last && last.text === p[d] && p.slice(0, d).every((v, k) => v === previous[k])) last.end = i;
      else spans.push({ text: p[d] ?? '', start: i, end: i });
    });
    levels.push(spans);
  }
  return levels;
}

/** 캐시의 span도 한 번만 펼칩니다. 항목마다 spans.find를 반복하지 않습니다. */
export function hierarchyPaths(data) {
  const count = data.categories?.length ?? data.series?.[0]?.values?.length ?? 0;
  if (data.hierarchyPaths?.length === count) return data.hierarchyPaths.map((p) => p.map(name));
  const levels = [...(data.catLevels ?? [])].reverse();
  const paths = Array.from({ length: count }, (_, i) => Array.from({ length: levels.length + 1 }, (_, d) => d === levels.length ? name(data.categories?.[i]) : ''));
  levels.forEach((spans, d) => {
    for (const span of spans) {
      const start = Math.max(0, Math.trunc(span.start)), end = Math.min(count - 1, Math.trunc(span.end));
      for (let i = start; i <= end; i++) paths[i][d] = name(span.text);
    }
  });
  return paths;
}

/** 필터 뒤에도 계층과 색의 원래 순서가 유지됩니다. 입력 객체는 바꾸지 않습니다. */
export function filterHierarchyData(data, keep) {
  const all = hierarchyPaths(data), paths = all.filter((_, i) => keep[i] ?? true);
  const roots = data.hierarchyRootOrder ?? [...new Set(all.map((p) => name(p[0])))];
  return { hierarchyPaths: paths, catLevels: levelsFromPaths(paths, all[0]?.length ?? 1), hierarchyRootOrder: [...roots] };
}

/** key는 JSON 경로: 이름에 구분자가 있어도 충돌하지 않으며 필터/정렬 뒤에도 바뀌지 않습니다. */
export function buildChartHierarchy(data) {
  const empty = { roots: [], nodes: [], total: 0, depth: 0, ignored: { zero: 0, negative: 0, missing: 0 } };
  const values = data.series?.[0]?.values ?? [], depth = data.hierarchyPaths?.[0]?.length ?? (data.catLevels?.length ?? 0) + 1;
  if (values.length > 30000 || depth > 64 || values.length * depth > 200000) return { ...empty, error: '계층 차트는 3만 항목·64단계·계층 칸 20만 개까지 표시합니다' };
  const paths = hierarchyPaths(data), indices = data.series?.[0]?._pi;
  const root = { children: [], map: new Map(), value: 0, depth: 0, path: [] }, nodes = [], ignored = { ...empty.ignored };
  const rootOrder = new Map((data.hierarchyRootOrder ?? [...new Set(paths.map((p) => name(p[0])))]).map((n, i) => [n, i]));
  values.forEach((value, index) => {
    if (!number(value)) { ignored.missing++; return; }
    if (value <= 0) { ignored[value ? 'negative' : 'zero']++; return; }
    const path = [...(paths[index] ?? [name(data.categories?.[index])])];
    while (path.length > 1 && path.at(-1) === '') path.pop();
    if (!path.length) path.push('');
    const point = indices?.[index] ?? index;
    let parent = root; root.value += value;
    path.forEach((part) => {
      if (!parent.map.has(part)) {
        const p = [...parent.path, part], key = JSON.stringify(p);
        const child = { name: part || '(비어 있음)', blank: part === '', key, parentKey: parent.key ?? null, path: p, value: 0, own: 0, points: [], ownPoints: [], indices: [], depth: parent.depth + 1, rootIndex: rootOrder.get(p[0]) ?? 0, children: [], map: new Map() };
        parent.map.set(part, child); parent.children.push(child); nodes.push(child);
      }
      parent = parent.map.get(part); parent.value += value; parent.points.push(point); parent.indices.push(index);
    });
    parent.own += value; parent.ownPoints.push(point);
  });
  if (!Number.isFinite(root.value)) return { ...empty, error: '계층 차트 값의 합계가 너무 큽니다' };
  const clean = (node) => { delete node.map; for (const child of node.children) clean(child); };
  for (const node of root.children) clean(node);
  return { roots: root.children, nodes, total: root.value, depth: nodes.reduce((n, node) => Math.max(n, node.depth), 0), ignored };
}

export function hierarchyNodeColor(node, series, chart, palette) {
  const own = series?.hierarchyColors?.[node.key];
  if (own) return own;
  // 부모 조각은 첫 번째 자식의 점 색으로 칠하지 않습니다.
  if (!node.children.length && node.ownPoints.length === 1) {
    const p = node.ownPoints[0], color = series?.pointColors?.[p] ?? series?.colors?.[p];
    if (color) return color;
  }
  for (let d = node.path.length - 1; d > 0; d--) {
    const color = series?.hierarchyColors?.[JSON.stringify(node.path.slice(0, d))];
    if (color) return color;
  }
  return chart.seriesFmt?.[series?._fi ?? 0]?.color ?? palette[node.rootIndex % palette.length];
}

export function hierarchyNodeAttrs(node, series) {
  const point = !node.children.length && node.ownPoints.length === 1 ? ` data-p="${node.ownPoints[0]}"` : '';
  return ` data-s="${series?._fi ?? 0}"${point} data-node="${esc(node.key)}" data-depth="${node.depth}" data-value="${node.value}"`;
}

export function hierarchyLabel(node, series, chart) {
  const value = series.labels ?? chart.labels ?? false;
  const category = series.catName ?? chart.labels !== false;
  const pieces = [];
  if (series.serName && series.name) pieces.push(series.name);
  if (category) pieces.push(node.name);
  if (value) { try { pieces.push(series.numFmt ? formatCode(node.value, series.numFmt, !!chart.date1904).text : formatGeneral(node.value)); } catch { pieces.push(formatGeneral(node.value)); } }
  if (series.pct && chart._hierarchyTotal) pieces.push(`${Math.round(node.value / chart._hierarchyTotal * 100)}%`);
  return pieces.join(' · ');
}

/** 한국어는 전각 폭으로 계산하여 작은 링/사각형 밖으로 긴 이름이 넘치지 않게 합니다. */
export function fitHierarchyLabel(value, width, size) {
  const letters = [...String(value)], unit = (c) => (c.codePointAt(0) > 255 ? 1 : .6) * size;
  let used = 0;
  for (const c of letters) used += unit(c);
  if (used <= width) return String(value);
  let text = ''; used = size;
  for (const c of letters) { if (used + unit(c) > width) break; text += c; used += unit(c); }
  return text ? text + '…' : '';
}

/** 양수 면적만 나누는 squarified treemap. 입력 순서와 객체는 바꾸지 않습니다. */
export function hierarchyRects(items, box) {
  const total = items.reduce((s, item) => s + item.value, 0);
  if (!(total > 0 && box.w > 0 && box.h > 0)) return [];
  const scale = box.w * box.h / total, rest = items.map((node, i) => ({ node, order: i, a: node.value * scale })).sort((a, b) => b.a - a.a || a.order - b.order), out = [];
  let b = { ...box }, at = 0;
  const worst = (sum, lo, hi, side) => Math.max(side * side * hi / (sum * sum), sum * sum / (side * side * lo));
  while (at < rest.length && b.w > 0 && b.h > 0) {
    const side = Math.min(b.w, b.h); let end = at + 1, sum = rest[at].a, lo = sum, hi = sum;
    while (end < rest.length) { const a = rest[end].a, next = sum + a; if (worst(next, Math.min(lo, a), Math.max(hi, a), side) > worst(sum, lo, hi, side)) break; sum = next; lo = Math.min(lo, a); hi = Math.max(hi, a); end++; }
    if (b.w >= b.h) {
      const width = Math.min(b.w, sum / b.h); let y = b.y;
      for (let i = at; i < end; i++) { const height = i === end - 1 ? b.y + b.h - y : rest[i].a / width; out.push({ node: rest[i].node, x: b.x, y, w: width, h: height }); y += height; }
      b.x += width; b.w -= width;
    } else {
      const height = Math.min(b.h, sum / b.w); let x = b.x;
      for (let i = at; i < end; i++) { const width = i === end - 1 ? b.x + b.w - x : rest[i].a / height; out.push({ node: rest[i].node, x, y: b.y, w: width, h: height }); x += width; }
      b.y += height; b.h -= height;
    }
    at = end;
  }
  return out;
}

export function hierarchyLegend(data, chart, palette) {
  const h = buildChartHierarchy(data), s = data.series?.[0];
  return h.roots.map((node) => ({ name: node.name, color: hierarchyNodeColor(node, s, chart, palette), line: false }));
}

export function drawHierarchyTreemap(ctx) {
  const { data, chart, series, plot, parts, pal } = ctx, hierarchy = buildChartHierarchy(data), s = series[0];
  if (hierarchy.error || !hierarchy.total) { parts.push(`<text x="${plot.x + plot.w / 2}" y="${plot.y + plot.h / 2}" text-anchor="middle" fill="${esc(ctx.TXT)}">${esc(hierarchy.error ?? '트리맵에는 양수 값과 계층 항목이 필요합니다')}</text>`); return; }
  const labelChart = { ...chart, _hierarchyTotal: hierarchy.total };
  const label = (node, box, parent = false) => {
    const content = parent ? node.name : hierarchyLabel(node, s, labelChart), size = (s.labelSize ?? 8.25) * 4 / 3;
    if (!content || box.w < size * 2.5 || box.h < size + 4) return;
    const clipped = fitHierarchyLabel(content, box.w - 8, size);
    parts.push(`<text data-el="label"${hierarchyNodeAttrs(node, s)} x="${box.x + 4}" y="${box.y + size + 2}" font-size="${size}" fill="${esc(s.labelColor ?? '#fff')}"${s.labelBold || parent ? ' font-weight="700"' : ''}><title>${esc(content)}</title>${esc(clipped)}</text>`);
  };
  const rectangle = (node, box) => parts.push(`<rect x="${box.x}" y="${box.y}" width="${Math.max(0, box.w - 1)}" height="${Math.max(0, box.h - 1)}" fill="${esc(hierarchyNodeColor(node, s, chart, pal))}" stroke="#fff" stroke-width="1"${hierarchyNodeAttrs(node, s)}><title>${esc(node.path.filter(Boolean).join(' / ') || node.name)}: ${esc(formatGeneral(node.value))}</title></rect>`);
  for (const group of hierarchyRects(hierarchy.roots, plot)) {
    const node = group.node;
    if (!node.children.length) { rectangle(node, group); label(node, group); continue; }
    rectangle(node, group);
    const mode = chart.treemapLabelLayout ?? 'banner', head = mode === 'banner' && group.w > 40 && group.h > 40 ? Math.min(24, (s.labelSize ?? 8.25) * 4 / 3 + 5) : 0;
    // Excel 트리맵과 같이 중간 계층은 경로/툴팁으로 보존하고 최상위 안에 말단을 배치합니다.
    const leaves = [], collect = (n) => { if (n.own > 0) leaves.push(n.children.length ? { ...n, value: n.own, children: [], points: n.ownPoints } : n); for (const c of n.children) collect(c); };
    collect(node);
    for (const leaf of hierarchyRects(leaves, { x: group.x + 1, y: group.y + head + 1, w: Math.max(0, group.w - 2), h: Math.max(0, group.h - head - 2) })) { rectangle(leaf.node, leaf); label(leaf.node, leaf); }
    if (mode !== 'none') label(node, { ...group, h: head || group.h }, true);
  }
}
