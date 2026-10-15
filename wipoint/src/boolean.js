// 도형 병합 (PowerPoint 도형 서식 › 도형 병합): 통합 · 결합 · 조각 · 교차 · 빼기 — DOM 없음
//
// 방법: 모든 도형 외곽을 다각형으로 펴고(곡선은 잘게), 서로 만나는 곳에서 변을 나눈 뒤
//   각 조각 변의 양쪽(왼쪽 · 오른쪽 아주 가까운 점)이 결과 영역에 속하는지 보고, 양쪽이 다르면 결과의 경계로 남김.
//   남은 변을 끝점끼리 이어 닫힌 고리로 만들면 구멍 · 떨어진 조각 · 겹친 변까지 한 번에 처리됨 (채우기는 짝수-홀수).
const EPS = 1e-7;

/** 경로 명령 (M/L/C/Q/Z, 도형 좌표) → 닫힌 다각형들 */
export function flatten(cmds, steps = 14) {
  const loops = [];
  let cur = null;
  let x = 0;
  let y = 0;
  const close = () => { if (cur && cur.length > 2) loops.push(cur); cur = null; };
  for (const c of cmds) {
    const op = c[0];
    if (op === 'M') { close(); cur = [[c[1], c[2]]]; x = c[1]; y = c[2]; }
    else if (op === 'L') { (cur ??= [[x, y]]).push([c[1], c[2]]); x = c[1]; y = c[2]; }
    else if (op === 'C') {
      cur ??= [[x, y]];
      for (let k = 1; k <= steps; k++) {
        const t = k / steps; const u = 1 - t;
        cur.push([u * u * u * x + 3 * u * u * t * c[1] + 3 * u * t * t * c[3] + t * t * t * c[5], u * u * u * y + 3 * u * u * t * c[2] + 3 * u * t * t * c[4] + t * t * t * c[6]]);
      }
      x = c[5]; y = c[6];
    } else if (op === 'Q') {
      cur ??= [[x, y]];
      for (let k = 1; k <= steps; k++) { const t = k / steps; const u = 1 - t; cur.push([u * u * x + 2 * u * t * c[1] + t * t * c[3], u * u * y + 2 * u * t * c[2] + t * t * c[4]]); }
      x = c[3]; y = c[4];
    } else if (op === 'Z') { if (cur) { x = cur[0][0]; y = cur[0][1]; } close(); }
  }
  close();
  // 같은 점 겹침 정리
  return loops.map((l) => l.filter((p, i) => { const q = l[(i + l.length - 1) % l.length]; return Math.hypot(p[0] - q[0], p[1] - q[1]) > 1e-6; })).filter((l) => l.length > 2);
}

/** 개체(x, y, w, h, rot, flip) 의 국소 좌표 다각형 → 슬라이드 좌표 */
export function placeLoops(loops, o) {
  const cx = o.w / 2;
  const cy = o.h / 2;
  const a = ((o.rot ?? 0) * Math.PI) / 180;
  const cos = Math.cos(a);
  const sin = Math.sin(a);
  return loops.map((l) => l.map(([px, py]) => {
    let dx = px - cx;
    let dy = py - cy;
    if (o.flipH) dx = -dx;
    if (o.flipV) dy = -dy;
    return [o.x + cx + dx * cos - dy * sin, o.y + cy + dx * sin + dy * cos];
  }));
}

/** 짝수-홀수 규칙으로 점이 다각형들 안에 있는가 */
export function inside(loops, x, y) {
  let c = false;
  for (const l of loops) {
    for (let i = 0, j = l.length - 1; i < l.length; j = i++) {
      const [xi, yi] = l[i];
      const [xj, yj] = l[j];
      if ((yi > y) !== (yj > y) && x < ((xj - xi) * (y - yi)) / (yj - yi) + xi) c = !c;
    }
  }
  return c;
}

function segIntersect(a, b, c, d) {
  const r0 = b[0] - a[0]; const r1 = b[1] - a[1];
  const s0 = d[0] - c[0]; const s1 = d[1] - c[1];
  const den = r0 * s1 - r1 * s0;
  if (Math.abs(den) < EPS) return null;
  const t = ((c[0] - a[0]) * s1 - (c[1] - a[1]) * s0) / den;
  const u = ((c[0] - a[0]) * r1 - (c[1] - a[1]) * r0) / den;
  if (t < -EPS || t > 1 + EPS || u < -EPS || u > 1 + EPS) return null;
  return [t, u];
}
/** 점 p 를 선분 c-d 에 투영한 매개변수 (선분 위에 있을 때만) — 겹치는 변 나누기 */
function onSeg(p, c, d, tol) {
  const dx = d[0] - c[0]; const dy = d[1] - c[1];
  const L = dx * dx + dy * dy;
  if (L < EPS) return null;
  const t = ((p[0] - c[0]) * dx + (p[1] - c[1]) * dy) / L;
  if (t <= EPS || t >= 1 - EPS) return null;
  const qx = c[0] + t * dx; const qy = c[1] + t * dy;
  return Math.hypot(qx - p[0], qy - p[1]) < tol ? t : null;
}

/**
 * shapes: [[loop, …], …] (도형마다 다각형들), keep(memberships: bool[]) → 결과 영역에 속하는가
 * → 결과 경계 고리들
 */
export function booleanLoops(shapes, keep) {
  // 1) 모든 변 모으기
  const segs = [];
  shapes.forEach((loops, si) => loops.forEach((l) => l.forEach((p, i) => segs.push({ a: p, b: l[(i + 1) % l.length], si, ts: [0, 1] }))));
  let span = 1;
  for (const s of segs) span = Math.max(span, Math.abs(s.a[0]), Math.abs(s.a[1]));
  const tol = span * 1e-7 + 1e-6;
  // 2) 다른 도형의 변과 만나는 곳에서 나누기 (같은 도형끼리도 — 자기 교차)
  for (let i = 0; i < segs.length; i++) {
    const s = segs[i];
    const minX = Math.min(s.a[0], s.b[0]); const maxX = Math.max(s.a[0], s.b[0]);
    const minY = Math.min(s.a[1], s.b[1]); const maxY = Math.max(s.a[1], s.b[1]);
    for (let j = i + 1; j < segs.length; j++) {
      const q = segs[j];
      if (Math.max(q.a[0], q.b[0]) < minX - tol || Math.min(q.a[0], q.b[0]) > maxX + tol || Math.max(q.a[1], q.b[1]) < minY - tol || Math.min(q.a[1], q.b[1]) > maxY + tol) continue;
      const hit = segIntersect(s.a, s.b, q.a, q.b);
      if (hit) { s.ts.push(hit[0]); q.ts.push(hit[1]); continue; }
      // 평행하게 겹침: 서로의 끝점에서 나눔
      for (const [p, into] of [[q.a, s], [q.b, s], [s.a, q], [s.b, q]]) { const t = onSeg(p, into.a, into.b, tol); if (t != null) into.ts.push(t); }
    }
  }
  // 3) 조각 변마다 양쪽 소속 비교
  const memb = (x, y) => keep(shapes.map((loops) => inside(loops, x, y)));
  const kept = new Map();
  for (const s of segs) {
    const ts = [...new Set(s.ts.map((t) => Math.round(Math.min(1, Math.max(0, t)) * 1e9) / 1e9))].sort((a, b) => a - b);
    const dx = s.b[0] - s.a[0]; const dy = s.b[1] - s.a[1];
    const len = Math.hypot(dx, dy);
    if (len < EPS) continue;
    const nx = -dy / len; const ny = dx / len;
    const off = Math.max(1e-4, len * 1e-5, tol * 20);
    for (let k = 0; k < ts.length - 1; k++) {
      const t0 = ts[k]; const t1 = ts[k + 1];
      if (t1 - t0 < 1e-9) continue;
      const p0 = [s.a[0] + dx * t0, s.a[1] + dy * t0];
      const p1 = [s.a[0] + dx * t1, s.a[1] + dy * t1];
      const mx = (p0[0] + p1[0]) / 2; const my = (p0[1] + p1[1]) / 2;
      if (memb(mx + nx * off, my + ny * off) === memb(mx - nx * off, my - ny * off)) continue;
      const k0 = key(p0, tol); const k1 = key(p1, tol);
      if (k0 === k1) continue;
      const id = k0 < k1 ? `${k0}|${k1}` : `${k1}|${k0}`;
      if (!kept.has(id)) kept.set(id, [p0, p1, k0, k1]);
    }
  }
  return linkLoops([...kept.values()]);
}
const key = (p, tol) => { const q = Math.max(tol * 10, 1e-4); return `${Math.round(p[0] / q)},${Math.round(p[1] / q)}`; };

/** 끝점끼리 이어 닫힌 고리 만들기 */
function linkLoops(edges) {
  const adj = new Map();
  edges.forEach((e, i) => { for (const k of [e[2], e[3]]) { if (!adj.has(k)) adj.set(k, []); adj.get(k).push(i); } });
  const used = new Uint8Array(edges.length);
  const loops = [];
  for (let i = 0; i < edges.length; i++) {
    if (used[i]) continue;
    used[i] = 1;
    const start = edges[i][2];
    const loop = [edges[i][0]];
    let curKey = edges[i][3];
    let curPt = edges[i][1];
    let guard = edges.length + 2;
    while (curKey !== start && guard-- > 0) {
      loop.push(curPt);
      const next = (adj.get(curKey) ?? []).find((j) => !used[j]);
      if (next == null) break;
      used[next] = 1;
      const e = edges[next];
      if (e[2] === curKey) { curKey = e[3]; curPt = e[1]; } else { curKey = e[2]; curPt = e[0]; }
    }
    if (curKey === start && loop.length > 2) loops.push(simplify(loop));
  }
  return loops.filter((l) => l.length > 2 && Math.abs(area(l)) > 1e-3);
}
/** 한 직선 위의 가운데 점 없애기 */
function simplify(l) {
  const out = [];
  for (let i = 0; i < l.length; i++) {
    const a = l[(i + l.length - 1) % l.length]; const b = l[i]; const c = l[(i + 1) % l.length];
    const cross = (b[0] - a[0]) * (c[1] - b[1]) - (b[1] - a[1]) * (c[0] - b[0]);
    const lab = Math.hypot(b[0] - a[0], b[1] - a[1]) * Math.hypot(c[0] - b[0], c[1] - b[1]);
    if (lab > 0 && Math.abs(cross) / lab < 1e-6) continue;
    out.push(b);
  }
  return out;
}
export const area = (l) => l.reduce((s, p, i) => { const q = l[(i + 1) % l.length]; return s + p[0] * q[1] - q[0] * p[1]; }, 0) / 2;

/** 고리들을 바깥 고리마다 (구멍 포함) 묶기 — [조각 = 고리들] */
export function groupPieces(loops) {
  const info = loops.map((l) => ({ l, a: Math.abs(area(l)), depth: 0, parent: -1 }));
  info.forEach((x, i) => {
    let best = -1;
    info.forEach((y, j) => { if (i !== j && y.a > x.a && inside([y.l], x.l[0][0] + 1e-9, x.l[0][1] + 1e-9) && (best < 0 || y.a < info[best].a)) best = j; });
    x.parent = best;
  });
  const depth = (i) => (info[i].parent < 0 ? 0 : 1 + depth(info[i].parent));
  const pieces = [];
  const owner = new Map();
  info.forEach((x, i) => { if (depth(i) % 2 === 0) { owner.set(i, pieces.length); pieces.push([x.l]); } });
  info.forEach((x, i) => { if (depth(i) % 2 === 1 && owner.has(x.parent)) pieces[owner.get(x.parent)].push(x.l); });
  return pieces;
}

export const MERGE_OPS = [['union', '통합'], ['combine', '결합'], ['fragment', '조각'], ['intersect', '교차'], ['subtract', '빼기']];
/** 슬라이드 좌표 도형 다각형들 → 결과 조각들 (조각마다 고리들). 빼기는 첫 도형에서 나머지를 뺌 */
export function mergeLoops(shapes, op) {
  if (op === 'fragment') {
    // 소속 조합마다 영역 (하나 이상의 도형에 속하는 칸)
    const sigs = new Map();
    // 각 도형 조합: 2^n 이 커지지 않게 n ≤ 8 로 제한
    const n = Math.min(shapes.length, 8);
    for (let mask = 1; mask < 1 << n; mask++) {
      const loops = booleanLoops(shapes.slice(0, n), (m) => m.every((v, i) => v === !!(mask & (1 << i))));
      if (loops.length) sigs.set(mask, loops);
    }
    return [...sigs.values()].flatMap((loops) => groupPieces(loops));
  }
  const keep = {
    union: (m) => m.some(Boolean),
    intersect: (m) => m.every(Boolean),
    subtract: (m) => m[0] && !m.slice(1).some(Boolean),
    combine: (m) => m.filter(Boolean).length % 2 === 1,
  }[op];
  const loops = booleanLoops(shapes, keep);
  return loops.length ? [loops] : [];
}

/** 결과 고리들 → 자유형 도형 위치 · 경로 ({ x, y, w, h, path }) */
export function loopsToShape(loops) {
  let x0 = Infinity; let y0 = Infinity; let x1 = -Infinity; let y1 = -Infinity;
  for (const l of loops) for (const [x, y] of l) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y); }
  const w = Math.max(1, x1 - x0);
  const h = Math.max(1, y1 - y0);
  const r = (v) => Math.round(v * 100) / 100;
  const cmds = [];
  for (const l of loops) { l.forEach(([x, y], i) => cmds.push([i ? 'L' : 'M', r(x - x0), r(y - y0)])); cmds.push(['Z']); }
  return { x: x0, y: y0, w, h, path: [{ w, h, cmds }] };
}
