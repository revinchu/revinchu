// DrawingML-compatible, DOM-free freeform geometry. Coordinates are relative
// to the object's box; Bezier control points may lie outside 0..1.
import { child, kids, esc } from './xml.js';

export const MAX_SHAPE_PATH_COMMANDS = 10000;
const ARITY = { M: 2, L: 2, C: 6, Q: 4, Z: 0 };
const near = (a, b) => a[0] === b[0] && a[1] === b[1];
const rounded = n => Math.round(n * 1000000) / 1000000;

export function validShapePath(path) {
  if (!path || !Array.isArray(path.paths) || !path.paths.length || path.paths.length > 256) return false;
  let count = 0;
  for (const p of path.paths) {
    if (!Array.isArray(p.commands) || !p.commands.length || p.commands[0]?.[0] !== 'M') return false;
    for (const c of p.commands) {
      if (++count > MAX_SHAPE_PATH_COMMANDS || !Array.isArray(c) || !(c[0] in ARITY) || c.length !== ARITY[c[0]] + 1) return false;
      for (let i = 1; i < c.length; i++) if (typeof c[i] !== 'number' || !Number.isFinite(c[i]) || Math.abs(c[i]) > 1000000) return false;
    }
  }
  return true;
}

export function shapePathParts(path, width, height) {
  if (!validShapePath(path)) return [];
  return path.paths.map(p => ({ d: p.commands.map(c => c[0] + c.slice(1).map((n, i) => rounded(n * (i % 2 ? height : width))).join(',')).join(' '),
    line: p.fill === false, noStroke: p.stroke === false }));
}

function curveValue(p0, p1, p2, p3, t) { const u = 1 - t; return u * u * u * p0 + 3 * u * u * t * p1 + 3 * u * t * t * p2 + t * t * t * p3; }
function roots(p0, p1, p2, p3) {
  const a = -p0 + 3 * p1 - 3 * p2 + p3, b = 2 * (p0 - 2 * p1 + p2), c = p1 - p0;
  if (Math.abs(a) < 1e-12) return Math.abs(b) < 1e-12 ? [] : [-c / b];
  const d = b * b - 4 * a * c;
  return d < 0 ? [] : [(-b + Math.sqrt(d)) / (2 * a), (-b - Math.sqrt(d)) / (2 * a)];
}

export function normalizedPathBounds(path) {
  if (!validShapePath(path)) return null;
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  const include = (x, y) => { minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y); };
  for (const p of path.paths) {
    let current = [0, 0], start = current;
    for (const c of p.commands) {
      if (c[0] === 'M' || c[0] === 'L') { current = c.slice(1); if (c[0] === 'M') start = current; include(...current); }
      else if (c[0] === 'Z') { current = start; include(...current); }
      else {
        const cp = c[0] === 'C' ? c.slice(1) : [current[0] + (c[1] - current[0]) * 2 / 3, current[1] + (c[2] - current[1]) * 2 / 3, c[3] + (c[1] - c[3]) * 2 / 3, c[4] + (c[2] - c[4]) * 2 / 3, c[3], c[4]];
        include(cp[4], cp[5]);
        for (const t of [...roots(current[0], cp[0], cp[2], cp[4]), ...roots(current[1], cp[1], cp[3], cp[5])]) if (t > 0 && t < 1) include(curveValue(current[0], cp[0], cp[2], cp[4], t), curveValue(current[1], cp[1], cp[3], cp[5], t));
        current = cp.slice(4);
      }
    }
  }
  return { minX, minY, maxX, maxY };
}

/** Sheet-space pointer points -> box + normalized path. Does not allocate an id. */
export function shapeFromPoints(kind, points, { closed = false } = {}) {
  if (!['scribble', 'freeform', 'curve'].includes(kind)) throw new Error('지원하지 않는 자유 도형입니다.');
  if (!Array.isArray(points) || points.length < 2 || points.length > MAX_SHAPE_PATH_COMMANDS - 2) throw new Error('도형 점은 2개 이상, 9,998개 이하로 입력하세요.');
  const pts = [];
  for (const point of points) {
    if (!Array.isArray(point) || point.length !== 2 || !point.every(n => typeof n === 'number' && Number.isFinite(n))) throw new Error('도형 점 좌표가 올바르지 않습니다.');
    if (!pts.length || !near(point, pts.at(-1))) pts.push([...point]);
  }
  if (pts.length < 2) throw new Error('서로 다른 도형 점이 필요합니다.');
  if (closed && near(pts[0], pts.at(-1))) pts.pop();
  const commands = [['M', ...pts[0]]];
  const n = pts.length;
  for (let i = 0; i < n - 1 + Number(closed && kind === 'curve'); i++) {
    const p = pts[i % n], q = pts[(i + 1) % n];
    if (kind === 'curve') {
      const prev = closed ? pts[(i - 1 + n) % n] : pts[Math.max(0, i - 1)], next = closed ? pts[(i + 2) % n] : pts[Math.min(n - 1, i + 2)];
      commands.push(['C', p[0] + (q[0] - prev[0]) / 6, p[1] + (q[1] - prev[1]) / 6,
        q[0] - (next[0] - p[0]) / 6, q[1] - (next[1] - p[1]) / 6, ...q]);
    } else commands.push(['L', ...q]);
  }
  if (closed) commands.push(['Z']);
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity, current = pts[0];
  const include = (x, y) => { minX = Math.min(minX, x); minY = Math.min(minY, y); maxX = Math.max(maxX, x); maxY = Math.max(maxY, y); };
  for (const c of commands) {
    if (c[0] === 'M' || c[0] === 'L') { include(c[1], c[2]); current = c.slice(1); }
    else if (c[0] === 'C') {
      include(c[5], c[6]);
      for (const t of [...roots(current[0], c[1], c[3], c[5]), ...roots(current[1], c[2], c[4], c[6])]) {
        if (t > 0 && t < 1) include(curveValue(current[0], c[1], c[3], c[5], t), curveValue(current[1], c[2], c[4], c[6], t));
      }
      current = c.slice(5);
    }
  }
  const w = Math.max(1, maxX - minX), h = Math.max(1, maxY - minY);
  const normalized = commands.map(c => [c[0], ...c.slice(1).map((v, i) => (v - (i % 2 ? minY : minX)) / (i % 2 ? h : w))]);
  return { x: minX, y: minY, w, h, path: { paths: [{ commands: normalized, fill: closed, stroke: true }] } };
}

/** Standard custom geometry, without WIXEL extension metadata. */
export function customGeometryXml(path) {
  if (!validShapePath(path)) throw new Error('도형 경로를 저장할 수 없습니다.');
  const scale = 1000000;
  const pt = (x, y) => `<a:pt x="${Math.round(x * scale)}" y="${Math.round(y * scale)}"/>`;
  const tags = { M: 'moveTo', L: 'lnTo', C: 'cubicBezTo', Q: 'quadBezTo' };
  const paths = path.paths.map(p => `<a:path w="${scale}" h="${scale}" fill="${p.fill === false ? 'none' : 'norm'}" stroke="${p.stroke === false ? 0 : 1}">${p.commands.map(c => {
    if (c[0] === 'Z') return '<a:close/>';
    let body = ''; for (let i = 1; i < c.length; i += 2) body += pt(c[i], c[i + 1]);
    return `<a:${tags[c[0]]}>${body}</a:${tags[c[0]]}>`;
  }).join('')}</a:path>`).join('');
  return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`;
}

/** Literal Office freeforms plus simple val guides. Unsupported formula/arc
 * geometry returns null so the importer can report the compatibility limit. */
export function readCustomGeometry(geom, width, height) {
  if (!geom) return null;
  const paths = [];
  for (const p of kids(child(geom, 'pathLst'), 'path')) {
    const w = Number(p.attrs.w ?? width * 9525), h = Number(p.attrs.h ?? height * 9525);
    if (!(w > 0 && h > 0 && Number.isFinite(w) && Number.isFinite(h))) return null;
    const vars = { w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, wd2: w / 2, hd2: h / 2 };
    const value = s => /^-?(?:\d+(?:\.\d*)?|\.\d+)$/.test(String(s)) ? Number(s) : vars[s];
    for (const g of [...kids(child(geom, 'avLst'), 'gd'), ...kids(child(geom, 'gdLst'), 'gd')]) {
      const m = /^val\s+(\S+)$/.exec(g.attrs.fmla ?? ''); if (m && Number.isFinite(value(m[1]))) vars[g.attrs.name] = value(m[1]);
    }
    const commands = [];
    for (const c of p.children) {
      if (c.name === 'close') { commands.push(['Z']); continue; }
      const type = { moveTo: 'M', lnTo: 'L', cubicBezTo: 'C', quadBezTo: 'Q' }[c.name];
      if (!type) return null;
      const values = kids(c, 'pt').flatMap(pt => [value(pt.attrs.x) / w, value(pt.attrs.y) / h]);
      commands.push([type, ...values]);
    }
    paths.push({ commands, fill: p.attrs.fill !== 'none', stroke: p.attrs.stroke !== '0' && p.attrs.stroke !== 'false' });
  }
  const path = { paths }; return validShapePath(path) ? path : null;
}

/** Preserve unsupported DrawingML guides/arcs as DrawingML, never as SVG/HTML. */
export function storedCustomGeometryXml(geom) {
  const names = new Set(['custGeom', 'avLst', 'gdLst', 'gd', 'ahLst', 'ahXY', 'ahPolar', 'pos', 'cxnLst', 'cxn', 'rect', 'pathLst', 'path', 'moveTo', 'lnTo', 'arcTo', 'cubicBezTo', 'quadBezTo', 'pt', 'close']);
  let count = 0;
  const write = n => {
    if (++count > 50000 || !names.has(n?.name)) throw new Error('사용자 지정 도형 원본을 저장할 수 없습니다.');
    const attrs = Object.entries(n.attrs ?? {}).filter(([key]) => /^[A-Za-z][A-Za-z0-9]*$/.test(key)).map(([key, value]) => ` ${key}="${esc(value)}"`).join('');
    return `<a:${n.name}${attrs}>${(n.children ?? []).map(write).join('')}</a:${n.name}>`;
  };
  if (geom?.name !== 'custGeom') throw new Error('사용자 지정 도형 원본이 올바르지 않습니다.');
  return write(geom);
}
