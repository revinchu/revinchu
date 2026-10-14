// DrawingML 도형 기하 계산 (미리 설정 도형 187종 · 사용자 지정 도형의 안내선 식) — DOM 없음
// presetGeom(name, w, h, adj) → { paths: [{ d, fill: 'norm'|'none'|'darken'|…, stroke }], rect: [l, t, r, b] } | null
import { PRESETS } from './prstgeom.js';

const f = (n) => Math.round(n * 100) / 100;
const RAD = Math.PI / 180 / 60000; // 60000분의 1도 → 라디안

/** 기본 변수 (w, h 에서) */
function baseVars(w, h) {
  const ss = Math.min(w, h); const ls = Math.max(w, h);
  const v = { w, h, l: 0, t: 0, r: w, b: h, hc: w / 2, vc: h / 2, ss, ls, cd2: 10800000, cd4: 5400000, cd8: 2700000, '3cd4': 16200000, '3cd8': 8100000, '5cd8': 13500000, '7cd8': 18900000 };
  for (const n of [2, 3, 4, 5, 6, 8, 10, 12, 32]) { v[`wd${n}`] = w / n; v[`hd${n}`] = h / n; }
  for (const n of [2, 4, 6, 8, 16, 32]) v[`ssd${n}`] = ss / n;
  return v;
}

/** 안내선 식 하나 계산 (ECMA-376 20.1.10.38) */
export function evalFormula(fmla, v) {
  const [op, ...args] = String(fmla).trim().split(/\s+/);
  const a = args.map((x) => (x in v ? v[x] : Number(x) || 0));
  const [x, y, z] = a;
  switch (op) {
    case '*/': return (x * y) / (z || 1);
    case '+-': return x + y - z;
    case '+/': return (x + y) / (z || 1);
    case '?:': return x > 0 ? y : z;
    case 'abs': return Math.abs(x);
    case 'at2': return (Math.atan2(y, x) / RAD);
    case 'cat2': return x * Math.cos(Math.atan2(z, y));
    case 'cos': return x * Math.cos(y * RAD);
    case 'max': return Math.max(x, y);
    case 'min': return Math.min(x, y);
    case 'mod': return Math.sqrt(x * x + y * y + z * z);
    case 'pin': return y < x ? x : y > z ? z : y;
    case 'sat2': return x * Math.sin(Math.atan2(z, y));
    case 'sin': return x * Math.sin(y * RAD);
    case 'sqrt': return Math.sqrt(Math.max(0, x));
    case 'tan': return x * Math.tan(y * RAD);
    case 'val': return x;
    default: return 0;
  }
}

/** 조정값 · 안내선 → 변수 표 */
export function guideVars(w, h, av = [], gd = [], adj = {}) {
  const v = baseVars(w, h);
  for (const [n, fm] of av) v[n] = adj?.[n] != null ? Number(adj[n]) : evalFormula(fm, v);
  // 조정값 이름이 adj 하나뿐인 도형에서 adj1 로 저장된 값, 그 반대도 받아 줌
  if (adj && av.length === 1) { const n = av[0][0]; const alt = n === 'adj' ? adj.adj1 : n === 'adj1' ? adj.adj : null; if (adj[n] == null && alt != null) v[n] = Number(alt); }
  for (const [n, fm] of gd) v[n] = evalFormula(fm, v);
  return v;
}

/** 경로 명령 (이름 또는 숫자 인수) → SVG d. 경로의 w · h 가 있으면 그 좌표계를 도형 크기로 늘림 */
export function pathD(path, v, W, H) {
  const sx = path.w ? W / path.w : 1;
  const sy = path.h ? H / path.h : 1;
  const val = (x) => (typeof x === 'number' ? x : x in v ? v[x] : Number(x) || 0);
  const out = [];
  let cx = 0; let cy = 0; let mx = 0; let my = 0;
  for (const [op, ...a] of path.c) {
    if (op === 'M' || op === 'L') {
      cx = val(a[0]); cy = val(a[1]);
      if (op === 'M') { mx = cx; my = cy; }
      out.push(`${op}${f(cx * sx)} ${f(cy * sy)}`);
    } else if (op === 'C') {
      const p = a.map(val);
      out.push(`C${f(p[0] * sx)} ${f(p[1] * sy)} ${f(p[2] * sx)} ${f(p[3] * sy)} ${f(p[4] * sx)} ${f(p[5] * sy)}`);
      cx = p[4]; cy = p[5];
    } else if (op === 'Q') {
      const p = a.map(val);
      out.push(`Q${f(p[0] * sx)} ${f(p[1] * sy)} ${f(p[2] * sx)} ${f(p[3] * sy)}`);
      cx = p[2]; cy = p[3];
    } else if (op === 'A') {
      const [wr, hr, st, sw] = a.map(val);
      const r = arcSvg(cx, cy, wr, hr, st, sw, sx, sy);
      out.push(r.d); cx = r.x; cy = r.y;
    } else if (op === 'Z') { out.push('Z'); cx = mx; cy = my; }
  }
  return out.join(' ');
}

/** arcTo: 현재 점에서 타원 (wR, hR) 위를 stAng 부터 swAng 만큼 (각도는 보이는 각도, 60000분의 1도) */
export function arcSvg(cx, cy, wr, hr, st, sw, sx = 1, sy = 1) {
  const param = (ang) => Math.atan2(wr * Math.sin(ang * RAD), hr * Math.cos(ang * RAD));
  const t0 = param(st);
  const ox = cx - wr * Math.cos(t0); const oy = cy - hr * Math.sin(t0);
  if (!wr || !hr || !sw) return { d: '', x: cx, y: cy };
  // 한 바퀴 이상이면 두 번에 나눠 그림 (SVG 호는 시작 = 끝이면 그려지지 않음)
  const steps = Math.abs(sw) > 10800000 ? 2 : 1;
  let d = ''; let x = cx; let y = cy;
  for (let i = 1; i <= steps; i++) {
    const ang = st + (sw * i) / steps;
    let t1 = param(ang);
    // 매개 변수 각을 보이는 각과 같은 바퀴에 맞춤
    const turns = Math.round((ang * RAD - t1) / (2 * Math.PI));
    t1 += turns * 2 * Math.PI;
    const ex = ox + wr * Math.cos(t1); const ey = oy + hr * Math.sin(t1);
    const span = Math.abs(sw / steps);
    d += `${d ? ' ' : ''}A${f(wr * sx)} ${f(hr * sy)} 0 ${span > 10800000 ? 1 : 0} ${sw > 0 ? 1 : 0} ${f(ex * sx)} ${f(ey * sy)}`;
    x = ex; y = ey;
  }
  return { d, x, y };
}

const cache = new Map();
/** 미리 설정 도형 → 경로들 + 글 영역 (없는 이름이면 null) */
export function presetGeom(name, w, h, adj) {
  const def = PRESETS[name];
  if (!def) return null;
  const key = `${name}|${f(w)}|${f(h)}|${adj ? JSON.stringify(adj) : ''}`;
  const hit = cache.get(key);
  if (hit) return hit;
  const [av, gd, rect, paths] = def;
  const v = guideVars(w, h, av, gd, adj);
  const out = {
    paths: paths.map((p) => ({ d: pathD(p, v, w, h), fill: p.f ?? 'norm', stroke: p.s !== 0 })),
    rect: rect ? rect.map((x) => (typeof x === 'number' ? x : v[x] ?? 0)) : [0, 0, w, h],
  };
  if (cache.size > 3000) cache.clear();
  cache.set(key, out);
  return out;
}

/** 조정값 기본값 (avLst) */
export function presetAdjDefaults(name) {
  const def = PRESETS[name];
  if (!def) return {};
  const v = baseVars(100, 100);
  return Object.fromEntries(def[0].map(([n, fm]) => [n, evalFormula(fm, v)]));
}
export const PRESET_NAMES = Object.keys(PRESETS);

/** 그룹 안 사각형 (자식 좌표, rot · flipH · flipV) → 바깥 좌표: 그룹마다 chOff/chExt 비율 + 그룹 중심 기준 대칭 · 회전 */
export function placeInGroups(xf, r) {
  if (!xf || !r) return r;
  let { x, y, w, h } = r;
  let rot = r.rot ?? 0; let fH = !!r.flipH; let fV = !!r.flipV;
  for (let g = xf; g; g = g.up) {
    const sx = g.chw ? g.w / g.chw : 1; const sy = g.chh ? g.h / g.chh : 1;
    let cx = g.x + (x + w / 2 - (g.chx ?? g.x)) * sx;
    let cy = g.y + (y + h / 2 - (g.chy ?? g.y)) * sy;
    w *= sx; h *= sy;
    if (g.flipH || g.flipV || g.rot) {
      const gcx = g.x + g.w / 2; const gcy = g.y + g.h / 2;
      let dx = cx - gcx; let dy = cy - gcy;
      if (g.flipH) dx = -dx;
      if (g.flipV) dy = -dy;
      const t = ((g.rot ?? 0) * Math.PI) / 180;
      cx = gcx + dx * Math.cos(t) - dy * Math.sin(t);
      cy = gcy + dx * Math.sin(t) + dy * Math.cos(t);
      if (!!g.flipH !== !!g.flipV) { rot = -rot; if (g.flipH) fH = !fH; else fV = !fV; } else if (g.flipH && g.flipV) rot += 180;
      rot += g.rot ?? 0;
    }
    x = cx - w / 2; y = cy - h / 2;
  }
  const out = { ...r, x, y, w, h, rot: Math.round((((rot % 360) + 360) % 360) * 1000) / 1000 };
  delete out.flipH; delete out.flipV;
  if (fH) out.flipH = true;
  if (fV) out.flipV = true;
  return out;
}
