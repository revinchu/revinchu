// SVG 열기 (DOM 없음): SVG 파일 → 슬라이드 개체 (편집 가능한 자유형 도형 · 글 상자 · 그림)
//   path (호 A 는 베지어로) · rect · circle · ellipse · line · polyline · polygon · text/tspan · image · g/use,
//   transform (matrix/translate/scale/rotate/skew) · fill/stroke/opacity · style 속성 · <style> 의 .class/#id/태그 규칙 · 그라데이션(색 정지점)
import { parseXml } from './xml.js';
import { uid } from './model.js';

const NUM = /[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g;
const nums = (s) => (String(s ?? '').match(NUM) ?? []).map(Number);
const I = [1, 0, 0, 1, 0, 0];
const mul = (a, b) => [a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1], a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3], a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5]];
const ap = (m, x, y) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]];

export function parseTransform(s) {
  let m = I;
  for (const [, fn, args] of String(s ?? '').matchAll(/(\w+)\s*\(([^)]*)\)/g)) {
    const v = nums(args);
    let t = I;
    if (fn === 'matrix' && v.length === 6) t = v;
    else if (fn === 'translate') t = [1, 0, 0, 1, v[0] ?? 0, v[1] ?? 0];
    else if (fn === 'scale') t = [v[0] ?? 1, 0, 0, v[1] ?? v[0] ?? 1, 0, 0];
    else if (fn === 'rotate') {
      const a = ((v[0] ?? 0) * Math.PI) / 180; const c = Math.cos(a); const sn = Math.sin(a);
      t = [c, sn, -sn, c, 0, 0];
      if (v.length >= 3) t = mul(mul([1, 0, 0, 1, v[1], v[2]], t), [1, 0, 0, 1, -v[1], -v[2]]);
    } else if (fn === 'skewX') t = [1, 0, Math.tan(((v[0] ?? 0) * Math.PI) / 180), 1, 0, 0];
    else if (fn === 'skewY') t = [1, Math.tan(((v[0] ?? 0) * Math.PI) / 180), 0, 1, 0, 0];
    m = mul(m, t);
  }
  return m;
}

/** SVG path d → 절대 좌표 [['M',x,y],['L',x,y],['C',…6],['Q',…4],['Z']] */
export function parsePath(d) {
  const toks = String(d ?? '').match(/[a-zA-Z]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g) ?? [];
  const out = [];
  let i = 0; let cmd = ''; let x = 0; let y = 0; let sx = 0; let sy = 0; let lc = null; let lq = null;
  const n = () => Number(toks[i++]);
  const more = () => i < toks.length && !/^[a-zA-Z]$/.test(toks[i]);
  while (i < toks.length) {
    if (/^[a-zA-Z]$/.test(toks[i])) cmd = toks[i++];
    const rel = cmd === cmd.toLowerCase();
    const C = cmd.toUpperCase();
    if (C === 'Z') { out.push(['Z']); x = sx; y = sy; lc = lq = null; if (!more()) continue; }
    if (C === 'M') { let nx = n(); let ny = n(); if (rel) { nx += x; ny += y; } out.push(['M', nx, ny]); x = sx = nx; y = sy = ny; cmd = rel ? 'l' : 'L'; lc = lq = null; continue; }
    if (C === 'L') { let nx = n(); let ny = n(); if (rel) { nx += x; ny += y; } out.push(['L', nx, ny]); x = nx; y = ny; lc = lq = null; continue; }
    if (C === 'H') { let nx = n(); if (rel) nx += x; out.push(['L', nx, y]); x = nx; lc = lq = null; continue; }
    if (C === 'V') { let ny = n(); if (rel) ny += y; out.push(['L', x, ny]); y = ny; lc = lq = null; continue; }
    if (C === 'C' || C === 'S') {
      let x1; let y1;
      if (C === 'C') { x1 = n(); y1 = n(); if (rel) { x1 += x; y1 += y; } } else { [x1, y1] = lc ? [2 * x - lc[0], 2 * y - lc[1]] : [x, y]; }
      let x2 = n(); let y2 = n(); let ex = n(); let ey = n();
      if (rel) { x2 += x; y2 += y; ex += x; ey += y; }
      out.push(['C', x1, y1, x2, y2, ex, ey]); lc = [x2, y2]; lq = null; x = ex; y = ey; continue;
    }
    if (C === 'Q' || C === 'T') {
      let x1; let y1;
      if (C === 'Q') { x1 = n(); y1 = n(); if (rel) { x1 += x; y1 += y; } } else { [x1, y1] = lq ? [2 * x - lq[0], 2 * y - lq[1]] : [x, y]; }
      let ex = n(); let ey = n();
      if (rel) { ex += x; ey += y; }
      out.push(['Q', x1, y1, ex, ey]); lq = [x1, y1]; lc = null; x = ex; y = ey; continue;
    }
    if (C === 'A') {
      const rx = n(); const ry = n(); const rot = n(); const large = n(); const sweep = n();
      let ex = n(); let ey = n();
      if (rel) { ex += x; ey += y; }
      out.push(...arcToCubic(x, y, rx, ry, rot, large, sweep, ex, ey));
      x = ex; y = ey; lc = lq = null; continue;
    }
    i++; // 모르는 것 건너뜀
  }
  return out;
}

/** SVG 호 → 베지어 (W3C 구현 노트의 중심 매개변수화) */
function arcToCubic(x1, y1, rx, ry, phiDeg, fa, fs, x2, y2) {
  if (!rx || !ry) return [['L', x2, y2]];
  rx = Math.abs(rx); ry = Math.abs(ry);
  const phi = (phiDeg * Math.PI) / 180; const cp = Math.cos(phi); const sp = Math.sin(phi);
  const dx = (x1 - x2) / 2; const dy = (y1 - y2) / 2;
  const x1p = cp * dx + sp * dy; const y1p = -sp * dx + cp * dy;
  const lam = (x1p * x1p) / (rx * rx) + (y1p * y1p) / (ry * ry);
  if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam); }
  const sign = fa === fs ? -1 : 1;
  const num = rx * rx * ry * ry - rx * rx * y1p * y1p - ry * ry * x1p * x1p;
  const co = sign * Math.sqrt(Math.max(0, num / (rx * rx * y1p * y1p + ry * ry * x1p * x1p)));
  const cxp = (co * rx * y1p) / ry; const cyp = (-co * ry * x1p) / rx;
  const cx = cp * cxp - sp * cyp + (x1 + x2) / 2; const cy = sp * cxp + cp * cyp + (y1 + y2) / 2;
  const ang = (ux, uy, vx, vy) => { const a = Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy); return a; };
  const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry);
  let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry);
  if (!fs && dt > 0) dt -= 2 * Math.PI; else if (fs && dt < 0) dt += 2 * Math.PI;
  const segs = Math.ceil(Math.abs(dt) / (Math.PI / 2));
  const out = [];
  const d = dt / segs;
  const k = (4 / 3) * Math.tan(d / 4);
  let t = t1;
  const pt = (a) => [cx + rx * Math.cos(a) * cp - ry * Math.sin(a) * sp, cy + rx * Math.cos(a) * sp + ry * Math.sin(a) * cp];
  const dv = (a) => [-rx * Math.sin(a) * cp - ry * Math.cos(a) * sp, -rx * Math.sin(a) * sp + ry * Math.cos(a) * cp];
  for (let s = 0; s < segs; s++) {
    const a2 = t + d;
    const [px, py] = pt(t); const [qx, qy] = pt(a2); const [d1x, d1y] = dv(t); const [d2x, d2y] = dv(a2);
    out.push(['C', px + k * d1x, py + k * d1y, qx - k * d2x, qy - k * d2y, qx, qy]);
    t = a2;
  }
  return out;
}

// ───────────── 색 · 스타일 ─────────────
const NAMED = { black: '#000000', white: '#ffffff', red: '#ff0000', green: '#008000', blue: '#0000ff', yellow: '#ffff00', gray: '#808080', grey: '#808080', orange: '#ffa500', purple: '#800080', navy: '#000080', silver: '#c0c0c0', maroon: '#800000', teal: '#008080', lime: '#00ff00', aqua: '#00ffff', fuchsia: '#ff00ff', olive: '#808000', pink: '#ffc0cb', brown: '#a52a2a', gold: '#ffd700', darkgray: '#a9a9a9', lightgray: '#d3d3d3', crimson: '#dc143c', coral: '#ff7f50', tomato: '#ff6347', skyblue: '#87ceeb', steelblue: '#4682b4', indigo: '#4b0082', violet: '#ee82ee', tan: '#d2b48c', khaki: '#f0e68c', salmon: '#fa8072' };
function color(v) {
  if (!v) return null;
  v = v.trim().toLowerCase();
  if (v === 'none' || v === 'transparent') return 'none';
  if (NAMED[v]) return NAMED[v];
  let m = /^#([0-9a-f]{3})$/.exec(v);
  if (m) return `#${[...m[1]].map((c) => c + c).join('')}`;
  m = /^#([0-9a-f]{6})/.exec(v);
  if (m) return `#${m[1]}`;
  m = /^rgba?\(([^)]+)\)/.exec(v);
  if (m) { const p = m[1].split(/[\s,/]+/).filter(Boolean).slice(0, 3).map((x) => (x.endsWith('%') ? Math.round(parseFloat(x) * 2.55) : Math.round(Number(x)))); return `#${p.map((x) => Math.max(0, Math.min(255, x)).toString(16).padStart(2, '0')).join('')}`; }
  return null;
}
const STYLE_KEYS = ['fill', 'stroke', 'stroke-width', 'opacity', 'fill-opacity', 'stroke-opacity', 'font-size', 'font-family', 'font-weight', 'font-style', 'text-anchor', 'stroke-dasharray', 'display', 'visibility'];
function parseCss(text) {
  const rules = [];
  for (const [, sel, body] of String(text).replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
    const decl = {};
    for (const d of body.split(';')) { const i = d.indexOf(':'); if (i > 0) decl[d.slice(0, i).trim()] = d.slice(i + 1).trim(); }
    for (const s of sel.split(',')) rules.push([s.trim(), decl]);
  }
  return rules;
}
function matches(el, sel) {
  if (sel.startsWith('.')) return (el.attrs.class ?? '').split(/\s+/).includes(sel.slice(1));
  if (sel.startsWith('#')) return el.attrs.id === sel.slice(1);
  if (/^[a-z]+\.[\w-]+$/i.test(sel)) { const [t, c] = sel.split('.'); return el.name === t && (el.attrs.class ?? '').split(/\s+/).includes(c); }
  return el.name === sel;
}

/**
 * SVG 글 → { objects, width, height } (개체 좌표는 슬라이드 px, size 에 맞춰 크기 · 가운데)
 * opts.size = { w, h } (슬라이드 크기), opts.fit = true 이면 슬라이드 안에 맞춤
 */
export function svgToObjects(text, opts = {}) {
  const doc = parseXml(String(text));
  const root = doc?.name === 'svg' ? doc : doc?.children?.find((c) => c.name === 'svg');
  if (!root) throw new Error('SVG 파일이 아닙니다');
  const vb = nums(root.attrs.viewBox);
  const W = parseFloat(root.attrs.width) || vb[2] || 800;
  const H = parseFloat(root.attrs.height) || vb[3] || 600;
  let base = I;
  if (vb.length === 4 && vb[2] && vb[3]) base = [W / vb[2], 0, 0, H / vb[3], -vb[0] * (W / vb[2]), -vb[1] * (H / vb[3])];
  const css = [];
  const byId = new Map();
  const grads = new Map();
  const walkIds = (el) => { if (el.attrs?.id) byId.set(el.attrs.id, el); if (el.name === 'style') css.push(...parseCss(el.children.map((c) => c.text ?? '').join('') + (el.text ?? ''))); for (const c of el.children ?? []) walkIds(c); };
  walkIds(root);
  for (const [id, el] of byId) if (/Gradient$/.test(el.name)) {
    let g = el;
    while (!g.children.some((c) => c.name === 'stop') && (g.attrs['xlink:href'] ?? g.attrs.href)) g = byId.get(String(g.attrs['xlink:href'] ?? g.attrs.href).slice(1)) ?? { children: [], attrs: {} };
    const stops = g.children.filter((c) => c.name === 'stop').map((s) => {
      const st = Object.fromEntries((s.attrs.style ?? '').split(';').map((d) => d.split(':').map((x) => x.trim())).filter((x) => x[1]));
      const off = String(s.attrs.offset ?? '0');
      return [off.endsWith('%') ? parseFloat(off) / 100 : Number(off) || 0, color(st['stop-color'] ?? s.attrs['stop-color'] ?? '#000') ?? '#000000'];
    });
    if (stops.length) grads.set(id, { radial: el.name === 'radialGradient', stops, angle: (() => { const x1 = parseFloat(el.attrs.x1 ?? '0'); const x2 = parseFloat(el.attrs.x2 ?? '1'); const y1 = parseFloat(el.attrs.y1 ?? '0'); const y2 = parseFloat(el.attrs.y2 ?? '0'); return (Math.atan2(y2 - y1, x2 - x1) * 180) / Math.PI; })() });
  }
  const out = [];
  const media = {};
  const styleOf = (el, inherited) => {
    const s = { ...inherited };
    for (const [sel, decl] of css) if (matches(el, sel)) for (const k of STYLE_KEYS) if (decl[k] != null) s[k] = decl[k];
    for (const k of STYLE_KEYS) if (el.attrs[k] != null) s[k] = el.attrs[k];
    for (const d of (el.attrs.style ?? '').split(';')) { const i = d.indexOf(':'); if (i > 0) { const k = d.slice(0, i).trim(); if (STYLE_KEYS.includes(k)) s[k] = d.slice(i + 1).trim(); } }
    // opacity 는 곱해짐
    if (el.attrs.opacity != null || /opacity\s*:/.test(el.attrs.style ?? '')) s._op = (inherited._op ?? 1) * Number(s.opacity ?? 1); else s._op = inherited._op ?? 1;
    return s;
  };
  const scaleOf = (m) => Math.sqrt(Math.abs(m[0] * m[3] - m[1] * m[2])) || 1;
  const fillOf = (s) => {
    const v = s.fill ?? '#000000';
    const ref = /url\(#([^)]+)\)/.exec(v);
    if (ref) { const g = grads.get(ref[1]); return g ? (g.stops.length > 1 ? { type: 'gradient', stops: g.stops, angle: g.angle, ...(g.radial ? { path: true } : {}) } : { type: 'solid', color: g.stops[0][1] }) : null; }
    const c = color(v);
    if (!c || c === 'none') return null;
    const a = (s._op ?? 1) * Number(s['fill-opacity'] ?? 1);
    return { type: 'solid', color: c, ...(a < 1 ? { alpha: Math.round(a * 100) / 100 } : {}) };
  };
  const lineOf = (s, m) => {
    const c = color(s.stroke ?? 'none');
    if (!c || c === 'none') return null;
    const w = parseFloat(s['stroke-width'] ?? '1') * scaleOf(m);
    return { color: c, width: Math.max(0.25, Math.round(w * 100) / 100), dash: s['stroke-dasharray'] && s['stroke-dasharray'] !== 'none' ? 'dash' : 'solid' };
  };
  const pushPath = (cmds, s, m, el) => {
    if (!cmds.length) return;
    const pts = [];
    const tc = cmds.map(([op, ...v]) => { const r = [op]; for (let k = 0; k < v.length; k += 2) { const [px, py] = ap(m, v[k], v[k + 1]); r.push(px, py); pts.push([px, py]); } return r; });
    const xs = pts.map((p) => p[0]); const ys = pts.map((p) => p[1]);
    const x = Math.min(...xs); const y = Math.min(...ys);
    const w = Math.max(0.5, Math.max(...xs) - x); const h = Math.max(0.5, Math.max(...ys) - y);
    const rel = tc.map(([op, ...v]) => [op, ...v.map((n, k) => Math.round((n - (k % 2 ? y : x)) * 100) / 100)]);
    const closed = rel.some((c) => c[0] === 'Z');
    // 열린 선(닫히지 않은 경로)은 채우기를 따로 적지 않았으면 선으로 (그림 도구가 내보낸 선 · 화살표)
    const fill = el.name === 'line' || (!closed && s.fill == null) ? null : fillOf(s);
    const line = lineOf(s, m);
    if (!fill && !line) return;
    out.push({ id: uid(), type: 'shape', shape: 'rect', name: el.attrs.id ?? undefined, x, y, w, h, rot: 0, fill, line, path: [{ w, h, cmds: rel }], ...(closed || fill ? {} : { openPath: true }), text: null });
  };
  const walk = (el, m, inherited) => {
    if (['defs', 'style', 'title', 'desc', 'metadata', 'clipPath', 'mask', 'symbol', 'linearGradient', 'radialGradient', 'pattern', 'filter', 'marker'].includes(el.name)) return;
    const s = styleOf(el, inherited);
    if (s.display === 'none' || s.visibility === 'hidden') return;
    const mm = el.attrs.transform ? mul(m, parseTransform(el.attrs.transform)) : m;
    const A = (k, d = 0) => parseFloat(el.attrs[k] ?? d) || 0;
    switch (el.name) {
      case 'svg': case 'g': case 'a': case 'switch': for (const c of el.children) walk(c, mm, s); break;
      case 'use': {
        const ref = byId.get(String(el.attrs['xlink:href'] ?? el.attrs.href ?? '').slice(1));
        if (ref) { const t = mul(mm, [1, 0, 0, 1, A('x'), A('y')]); if (ref.name === 'symbol') for (const c of ref.children) walk(c, t, s); else walk(ref, t, s); }
        break;
      }
      case 'path': pushPath(parsePath(el.attrs.d), s, mm, el); break;
      case 'rect': {
        const x = A('x'); const y = A('y'); const w = A('width'); const h = A('height');
        let rx = Math.min(A('rx', el.attrs.ry ?? 0), w / 2); let ry = Math.min(A('ry', el.attrs.rx ?? 0), h / 2);
        if (!rx) ry = 0; if (!ry) rx = 0;
        const d = rx ? `M${x + rx} ${y}H${x + w - rx}A${rx} ${ry} 0 0 1 ${x + w} ${y + ry}V${y + h - ry}A${rx} ${ry} 0 0 1 ${x + w - rx} ${y + h}H${x + rx}A${rx} ${ry} 0 0 1 ${x} ${y + h - ry}V${y + ry}A${rx} ${ry} 0 0 1 ${x + rx} ${y}Z` : `M${x} ${y}H${x + w}V${y + h}H${x}Z`;
        pushPath(parsePath(d), s, mm, el);
        break;
      }
      case 'circle': case 'ellipse': {
        const cx = A('cx'); const cy = A('cy'); const rx = el.name === 'circle' ? A('r') : A('rx'); const ry = el.name === 'circle' ? A('r') : A('ry');
        pushPath(parsePath(`M${cx - rx} ${cy}A${rx} ${ry} 0 1 0 ${cx + rx} ${cy}A${rx} ${ry} 0 1 0 ${cx - rx} ${cy}Z`), s, mm, el);
        break;
      }
      case 'line': pushPath([['M', A('x1'), A('y1')], ['L', A('x2'), A('y2')]], s, mm, el); break;
      case 'polyline': case 'polygon': {
        const v = nums(el.attrs.points);
        const cmds = [];
        for (let k = 0; k + 1 < v.length; k += 2) cmds.push([k ? 'L' : 'M', v[k], v[k + 1]]);
        if (el.name === 'polygon') cmds.push(['Z']);
        pushPath(cmds, s, mm, el);
        break;
      }
      case 'text': {
        const parts = [];
        const collect = (n) => { if (n.text) parts.push(n.text); for (const c of n.children ?? []) collect(c); };
        collect(el);
        const t = parts.join('').replace(/\s+/g, ' ').trim();
        const tspan = el.children.find((c) => c.name === 'tspan');
        const x = A('x', tspan?.attrs.x ?? 0); const y = A('y', tspan?.attrs.y ?? 0);
        if (!t) break;
        const sc = scaleOf(mm);
        const size = parseFloat(s['font-size'] ?? '16') * sc;
        const [px, py] = ap(mm, x, y);
        const wEst = Math.max(size, t.length * size * 0.62);
        const anchor = s['text-anchor'] ?? 'start';
        const left = anchor === 'middle' ? px - wEst / 2 : anchor === 'end' ? px - wEst : px;
        const fam = String(s['font-family'] ?? '').split(',')[0].replace(/['"]/g, '').trim();
        const fc = fillOf(s);
        out.push({ id: uid(), type: 'shape', shape: 'rect', txBox: true, x: left, y: py - size * 1.05, w: wEst + 8, h: size * 1.35, rot: 0, fill: null, line: null,
          text: { paras: [{ align: anchor === 'middle' ? 'ctr' : anchor === 'end' ? 'r' : 'l', runs: [{ t, size: Math.round((size * 0.75) * 10) / 10, ...(fc?.color ? { color: fc.color } : {}), ...(fam ? { font: fam } : {}), ...(/bold|[6-9]00/.test(s['font-weight'] ?? '') ? { b: true } : {}), ...(s['font-style'] === 'italic' ? { i: true } : {}) }] }], anchor: 't', insets: [0, 0, 0, 0], wrap: false, autofit: 'none' } });
        break;
      }
      case 'image': {
        const href = el.attrs['xlink:href'] ?? el.attrs.href;
        if (!href || !/^data:image\//.test(href)) break;
        const id = uid('m');
        media[id] = href;
        const [x1, y1] = ap(mm, A('x'), A('y'));
        const [x2, y2] = ap(mm, A('x') + A('width'), A('y') + A('height'));
        out.push({ id: uid(), type: 'image', media: id, x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1), rot: 0 });
        break;
      }
      default: for (const c of el.children ?? []) walk(c, mm, s);
    }
  };
  walk(root, base, {});
  // 슬라이드에 맞추기 (비율 유지, 가운데)
  if (opts.size) {
    const k = opts.fit === false ? 1 : Math.min(opts.size.w / W, opts.size.h / H, opts.maxScale ?? Infinity);
    const ox = (opts.size.w - W * k) / 2; const oy = (opts.size.h - H * k) / 2;
    for (const o of out) {
      o.x = Math.round((o.x * k + ox) * 100) / 100; o.y = Math.round((o.y * k + oy) * 100) / 100; o.w *= k; o.h *= k;
      if (o.line) o.line.width = Math.max(0.25, Math.round(o.line.width * k * 100) / 100);
      if (o.text) for (const p of o.text.paras) for (const r of p.runs) r.size = Math.max(4, Math.round(r.size * k * 10) / 10);
    }
  }
  return { objects: out, media, width: W, height: H };
}
