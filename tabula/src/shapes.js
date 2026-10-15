// 그림 개체(차트·그림·도형) 공통 도우미와 도형 SVG (DOM 없음)

export const OBJECT_PROPS = ['charts', 'images', 'shapes', 'slicers'];
export const OBJECT_LABEL = { charts: '차트', images: '그림', shapes: '도형', slicers: '슬라이서' };

// 선 갤러리 항목 → 선 종류 + 화살표
const LINE_PRESET = {
  line: ['line'], lineArrow: ['line', 'end'], lineDblArrow: ['line', 'both'],
  bentConnector3: ['bentConnector3'], bentArrow3: ['bentConnector3', 'end'], curvedConnector3: ['curvedConnector3'], curvedArrow3: ['curvedConnector3', 'end'],
};

/** 새 도형 기본값 (kind: 갤러리 id = 엑셀 도형 이름) */
export function newShape(kind, box) {
  const id = `sh${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  if (kind === 'textbox') return { id, kind, ...box, fill: '#ffffff', stroke: '#000000', text: '', color: '#000000', size: 11, align: 'left' };
  if (LINE_PRESET[kind]) {
    const [k, arrow] = LINE_PRESET[kind];
    return { id, kind: k, ...box, fill: null, stroke: '#4472c4', strokeWidth: 1.5, text: '', ...(arrow ? { arrow } : {}) };
  }
  const lineOnly = GEOM[kind]?.(10, 10).every((p) => p.line);
  if (lineOnly) return { id, kind, ...box, fill: null, stroke: '#4472c4', strokeWidth: 1.5, text: '' };
  return { id, kind, ...box, fill: '#4472c4', stroke: '#2f528f', text: '', color: '#ffffff', size: 11, align: 'center' };
}

/** 시트에서 id 로 개체 찾기 → { prop, obj } | null */
export function findObject(sheet, id) {
  if (!id) return null;
  for (const prop of OBJECT_PROPS) {
    const obj = (sheet[prop] ?? []).find((o) => o.id === id);
    if (obj) return { prop, obj };
  }
  return null;
}

const attr = (s) => String(s).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');

function shadeHex(hex, t) {
  const m = /^#?([0-9a-f]{6})$/i.exec(hex ?? '');
  if (!m) return hex;
  const n = parseInt(m[1], 16);
  const ch = [n >> 16, (n >> 8) & 255, n & 255].map((x) => Math.round(t < 0 ? x * (1 + t) : x + (255 - x) * t));
  return `#${ch.map((x) => x.toString(16).padStart(2, '0')).join('')}`;
}

/** 도형 모양 SVG (글자는 따로 HTML 로 그림). 선 굵기만큼 안쪽으로 그려서 잘리지 않게 함 */
export function shapeSvg(sh) {
  const w = Math.max(1, sh.w);
  const h = Math.max(1, sh.h);
  const sw = sh.stroke ? (sh.strokeWidth ?? 1) : 0;
  const dash = sh.dash ? ` stroke-dasharray="${sh.dash === 'dot' ? `${sw},${sw * 2}` : `${sw * 4},${sw * 3}`}"` : '';
  let body;
  if (LINE_KINDS.has(sh.kind)) {
    const col = attr(sh.stroke ?? '#000000');
    const mk = `ar${col.replace('#', '')}`;
    const defs = sh.arrow ? `<defs><marker id="${mk}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="5" markerHeight="5" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 Z" fill="${col}"/></marker></defs>` : '';
    body = `${defs}<path d="${linePath(sh.kind, w, h, sh.flip, sh.flipV)}" fill="none" stroke="${col}" stroke-width="${Math.max(1, sw)}"${dash}${sh.arrow ? ` marker-end="url(#${mk})"` : ''}${sh.arrow === 'both' ? ` marker-start="url(#${mk})"` : ''}/>`;
  } else {
    const geom = GEOM[sh.kind] ?? GEOM.rect;
    // 선 굵기만큼 안쪽으로
    const i = sw / 2;
    const parts = geom(Math.max(0, w - sw), Math.max(0, h - sw));
    const stroke = sh.stroke ? `stroke="${attr(sh.stroke)}" stroke-width="${sw}" stroke-linejoin="round"${dash}${sh.strokeOpacity !== undefined ? ` stroke-opacity="${sh.strokeOpacity}"` : ''}` : 'stroke="none"';
    // 채우기: 단색 · 그라데이션 · 투명도, 효과: 그림자 · 네온 · 부드러운 가장자리
    const uid = `s${(sh.id ?? '').replace(/[^\w]/g, '')}${Math.round(w)}`;
    const defs = [];
    let gradRef = null;
    if (sh.grad && sh.fill) {
      const a = ((sh.grad.ang ?? 90) * Math.PI) / 180;
      const gx = Math.cos(a) / 2;
      const gy = Math.sin(a) / 2;
      const stops = sh.grad.stops ?? [[0, shadeHex(sh.fill, 0.35)], [1, shadeHex(sh.fill, -0.15)]];
      defs.push(`<linearGradient id="g${uid}" x1="${0.5 - gx}" y1="${0.5 - gy}" x2="${0.5 + gx}" y2="${0.5 + gy}">${stops.map(([o, c]) => `<stop offset="${o}" stop-color="${attr(c)}"/>`).join('')}</linearGradient>`);
      gradRef = `url(#g${uid})`;
    }
    const fx = [];
    if (sh.shadow) {
      const sd = typeof sh.shadow === 'object' ? sh.shadow : {};
      fx.push(`<feDropShadow dx="${sd.dx ?? 2.5}" dy="${sd.dy ?? 2.5}" stdDeviation="${sd.blur ?? 2.5}" flood-color="${attr(sd.color ?? '#000')}" flood-opacity="${sd.opacity ?? 0.4}"/>`);
    }
    if (sh.glow) fx.push(`<feMorphology operator="dilate" radius="${(sh.glow.size ?? 5) / 2}" in="SourceAlpha" result="gd"/><feGaussianBlur in="gd" stdDeviation="${(sh.glow.size ?? 5) / 2}" result="gb"/><feFlood flood-color="${attr(sh.glow.color ?? '#4472c4')}" flood-opacity="0.6"/><feComposite in2="gb" operator="in" result="gc"/><feMerge><feMergeNode in="gc"/><feMergeNode in="SourceGraphic"/></feMerge>`);
    if (sh.soft) fx.push(`<feGaussianBlur stdDeviation="${sh.soft / 2}"/>`);
    let filt = '';
    if (fx.length) {
      // 효과마다 따로 필터 (겹치면 원본 그래픽 기준이 달라지므로 그림자 → 네온 → 부드러운 가장자리 순서로 중첩)
      const ids = fx.map((f, k) => { defs.push(`<filter id="f${uid}${k}" x="-40%" y="-40%" width="180%" height="180%">${f}</filter>`); return `f${uid}${k}`; });
      filt = ids.reduce((acc, id) => `<g filter="url(#${id})">${acc}</g>`, '§');
    }
    const fillOf = (p) => {
      if (p.line || !sh.fill) return 'none';
      if (p.shade === 'dark') return attr(shadeHex(sh.fill, -0.2));
      if (p.shade === 'light') return attr(shadeHex(sh.fill, 0.2));
      return gradRef ?? attr(sh.fill);
    };
    const op = sh.fillOpacity !== undefined ? ` fill-opacity="${sh.fillOpacity}"` : '';
    const paths = parts.map((p) => `<path d="${p.d}" fill="${fillOf(p)}"${op}${p.evenodd ? ' fill-rule="evenodd"' : ''} ${p.line && !sh.stroke ? `stroke="${attr(sh.fill ?? '#000')}" stroke-width="1.5"` : stroke}/>`).join('');
    const flipT = sh.flip || sh.flipV ? ` translate(${sh.flip ? w - sw : 0},${sh.flipV ? h - sw : 0}) scale(${sh.flip ? -1 : 1},${sh.flipV ? -1 : 1})` : '';
    const g = `<g transform="translate(${i},${i})${flipT}">${paths}</g>`;
    body = `${defs.length ? `<defs>${defs.join('')}</defs>` : ''}${filt ? filt.replace('§', g) : g}`;
  }
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" overflow="visible">${body}</svg>`;
}

// ───────────── 도형 모양 (엑셀 기본 도형 이름 = prstGeom 그대로) ─────────────
// 모양 함수 (w, h) → [{ d, shade?: 'dark'|'light', line?: true(채우기 없이 선만), evenodd?: true }]
const P = (pts) => `M${pts.map(([x, y]) => `${r2(x)},${r2(y)}`).join(' L')} Z`;
const r2 = (v) => Math.round(v * 100) / 100;
const one = (d, more) => [{ d, ...more }];
const ell = (cx, cy, rx, ry) => `M${r2(cx - rx)},${r2(cy)} a${r2(rx)},${r2(ry)} 0 1,0 ${r2(2 * rx)},0 a${r2(rx)},${r2(ry)} 0 1,0 ${r2(-2 * rx)},0 Z`;
const rectD = (x, y, w, h) => P([[x, y], [x + w, y], [x + w, y + h], [x, y + h]]);
const regular = (n, w, h, rot = -90) => P(Array.from({ length: n }, (_, k) => {
  const a = ((rot + (360 / n) * k) * Math.PI) / 180;
  return [w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)];
}));
const starD = (n, ratio, w, h) => P(Array.from({ length: n * 2 }, (_, k) => {
  const a = ((-90 + (180 / n) * k) * Math.PI) / 180;
  const f = k % 2 ? ratio : 1;
  return [w / 2 + (w / 2) * f * Math.cos(a), h / 2 + (h / 2) * f * Math.sin(a)];
}));
const grid = (pts, w, h, n = 21600) => P(pts.map(([x, y]) => [(x / n) * w, (y / n) * h]));
const roundRectD = (w, h, r, corners = [1, 1, 1, 1]) => {
  const [tl, tr, br, bl] = corners.map((c) => (c ? r : 0));
  return `M${tl},0 L${r2(w - tr)},0 ${tr ? `A${r2(tr)},${r2(tr)} 0 0,1 ${r2(w)},${r2(tr)}` : ''} L${r2(w)},${r2(h - br)} ${br ? `A${r2(br)},${r2(br)} 0 0,1 ${r2(w - br)},${r2(h)}` : ''} L${r2(bl)},${r2(h)} ${bl ? `A${r2(bl)},${r2(bl)} 0 0,1 0,${r2(h - bl)}` : ''} L0,${r2(tl)} ${tl ? `A${r2(tl)},${r2(tl)} 0 0,1 ${r2(tl)},0` : ''} Z`;
};
const plusD = (x, y, w, h, t) => {
  const ax = x + (w - t) / 2;
  const ay = y + (h - t) / 2;
  return P([[ax, y], [ax + t, y], [ax + t, ay], [x + w, ay], [x + w, ay + t], [ax + t, ay + t], [ax + t, y + h], [ax, y + h], [ax, ay + t], [x, ay + t], [x, ay], [ax, ay]]);
};
const rotPts = (pts, cx, cy, deg) => {
  const a = (deg * Math.PI) / 180;
  return pts.map(([x, y]) => [cx + (x - cx) * Math.cos(a) - (y - cy) * Math.sin(a), cy + (x - cx) * Math.sin(a) + (y - cy) * Math.cos(a)]);
};
const cloudD = (w, h) => {
  const n = 10;
  const pts = Array.from({ length: n }, (_, k) => {
    const a = (k / n) * Math.PI * 2 + 0.3;
    return [w / 2 + w * 0.4 * Math.cos(a), h / 2 + h * 0.36 * Math.sin(a)];
  });
  let d = `M${r2(pts[0][0])},${r2(pts[0][1])}`;
  for (let k = 1; k <= n; k++) {
    const [x, y] = pts[k % n];
    const [px, py] = pts[k - 1];
    const r = Math.hypot(x - px, y - py) * 0.62;
    d += ` A${r2(r)},${r2(r)} 0 1,1 ${r2(x)},${r2(y)}`;
  }
  return `${d} Z`;
};
const arrowRight = (w, h, hl = Math.min(w, h) * 0.5, t = 0.5) => P([[0, h * (0.5 - t / 2)], [w - hl, h * (0.5 - t / 2)], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * (0.5 + t / 2)], [0, h * (0.5 + t / 2)]]);
const arrowDown = (w, h, hl = Math.min(w, h) * 0.5) => P([[w * 0.25, 0], [w * 0.75, 0], [w * 0.75, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [w * 0.25, h - hl]]);

export const GEOM = {
  // 사각형
  rect: (w, h) => one(rectD(0, 0, w, h)),
  roundRect: (w, h) => one(roundRectD(w, h, Math.min(w, h) * 0.1667)),
  snip1Rect: (w, h) => { const s = Math.min(w, h) * 0.1667; return one(P([[0, 0], [w - s, 0], [w, s], [w, h], [0, h]])); },
  snip2SameRect: (w, h) => { const s = Math.min(w, h) * 0.1667; return one(P([[s, 0], [w - s, 0], [w, s], [w, h], [0, h], [0, s]])); },
  snip2DiagRect: (w, h) => { const s = Math.min(w, h) * 0.1667; return one(P([[0, 0], [w - s, 0], [w, s], [w, h], [s, h], [0, h - s]])); },
  round1Rect: (w, h) => one(roundRectD(w, h, Math.min(w, h) * 0.1667, [0, 1, 0, 0])),
  round2SameRect: (w, h) => one(roundRectD(w, h, Math.min(w, h) * 0.1667, [1, 1, 0, 0])),
  round2DiagRect: (w, h) => one(roundRectD(w, h, Math.min(w, h) * 0.1667, [1, 0, 1, 0])),
  snipRoundRect: (w, h) => { const s = Math.min(w, h) * 0.1667; return one(`M${r2(s)},0 L${r2(w - s)},0 L${r2(w)},${r2(s)} L${r2(w)},${r2(h)} L0,${r2(h)} L0,${r2(s)} A${r2(s)},${r2(s)} 0 0,1 ${r2(s)},0 Z`); },
  // 기본 도형
  ellipse: (w, h) => one(ell(w / 2, h / 2, w / 2, h / 2)),
  triangle: (w, h) => one(P([[w / 2, 0], [w, h], [0, h]])),
  rtTriangle: (w, h) => one(P([[0, 0], [w, h], [0, h]])),
  parallelogram: (w, h) => { const a = Math.min(w, h) * 0.25; return one(P([[a, 0], [w, 0], [w - a, h], [0, h]])); },
  trapezoid: (w, h) => { const a = Math.min(w, h) * 0.25; return one(P([[a, 0], [w - a, 0], [w, h], [0, h]])); },
  diamond: (w, h) => one(P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]])),
  pentagon: (w, h) => one(regular(5, w, h)),
  hexagon: (w, h) => { const a = Math.min(w, h) * 0.25; return one(P([[a, 0], [w - a, 0], [w, h / 2], [w - a, h], [a, h], [0, h / 2]])); },
  heptagon: (w, h) => one(regular(7, w, h)),
  octagon: (w, h) => { const a = Math.min(w, h) * 0.2929; return one(P([[a, 0], [w - a, 0], [w, a], [w, h - a], [w - a, h], [a, h], [0, h - a], [0, a]])); },
  decagon: (w, h) => one(regular(10, w, h, 0)),
  dodecagon: (w, h) => one(regular(12, w, h, -75)),
  pie: (w, h) => one(`M${r2(w / 2)},${r2(h / 2)} L${r2(w)},${r2(h / 2)} A${r2(w / 2)},${r2(h / 2)} 0 1,1 ${r2(w / 2)},0 Z`),
  chord: (w, h) => { const p = (a) => [w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)]; const [a1, a2] = [p((45 * Math.PI) / 180), p((270 * Math.PI) / 180)]; return one(`M${r2(a1[0])},${r2(a1[1])} A${r2(w / 2)},${r2(h / 2)} 0 1,1 ${r2(a2[0])},${r2(a2[1])} Z`); },
  teardrop: (w, h) => one(`M0,${r2(h / 2)} A${r2(w / 2)},${r2(h / 2)} 0 0,1 ${r2(w / 2)},0 L${r2(w)},0 L${r2(w)},${r2(h / 2)} A${r2(w / 2)},${r2(h / 2)} 0 0,1 ${r2(w / 2)},${r2(h)} A${r2(w / 2)},${r2(h / 2)} 0 0,1 0,${r2(h / 2)} Z`),
  frame: (w, h) => { const t = Math.min(w, h) * 0.125; return one(`${rectD(0, 0, w, h)} ${rectD(t, t, w - 2 * t, h - 2 * t)}`, { evenodd: true }); },
  halfFrame: (w, h) => { const t = Math.min(w, h) * 0.3333; return one(P([[0, 0], [w, 0], [w - t, t], [t, t], [t, h - t], [0, h]])); },
  corner: (w, h) => { const a = Math.min(w, h) * 0.5; return one(P([[0, 0], [a, 0], [a, h - a], [w, h - a], [w, h], [0, h]])); },
  diagStripe: (w, h) => one(P([[0, h / 2], [w / 2, 0], [w, 0], [0, h]])),
  plus: (w, h) => { const a = Math.min(w, h) * 0.25; return one(P([[a, 0], [w - a, 0], [w - a, a], [w, a], [w, h - a], [w - a, h - a], [w - a, h], [a, h], [a, h - a], [0, h - a], [0, a], [a, a]])); },
  plaque: (w, h) => { const r = Math.min(w, h) * 0.1667; return one(`M${r2(r)},0 L${r2(w - r)},0 A${r2(r)},${r2(r)} 0 0,0 ${r2(w)},${r2(r)} L${r2(w)},${r2(h - r)} A${r2(r)},${r2(r)} 0 0,0 ${r2(w - r)},${r2(h)} L${r2(r)},${r2(h)} A${r2(r)},${r2(r)} 0 0,0 0,${r2(h - r)} L0,${r2(r)} A${r2(r)},${r2(r)} 0 0,0 ${r2(r)},0 Z`); },
  can: (w, h) => { const ry = Math.min(h * 0.2, w * 0.25); return [{ d: `M0,${r2(ry)} A${r2(w / 2)},${r2(ry)} 0 0,1 ${r2(w)},${r2(ry)} L${r2(w)},${r2(h - ry)} A${r2(w / 2)},${r2(ry)} 0 0,1 0,${r2(h - ry)} Z` }, { d: ell(w / 2, ry, w / 2, ry), shade: 'light' }]; },
  cube: (w, h) => { const d = Math.min(w, h) * 0.25; return [{ d: P([[0, d], [w - d, d], [w - d, h], [0, h]]) }, { d: P([[0, d], [d, 0], [w, 0], [w - d, d]]), shade: 'light' }, { d: P([[w - d, d], [w, 0], [w, h - d], [w - d, h]]), shade: 'dark' }]; },
  bevel: (w, h) => { const d = Math.min(w, h) * 0.125; return [{ d: rectD(d, d, w - 2 * d, h - 2 * d) }, { d: P([[0, 0], [w, 0], [w - d, d], [d, d]]), shade: 'light' }, { d: P([[0, 0], [d, d], [d, h - d], [0, h]]), shade: 'light' }, { d: P([[w, 0], [w, h], [w - d, h - d], [w - d, d]]), shade: 'dark' }, { d: P([[0, h], [d, h - d], [w - d, h - d], [w, h]]), shade: 'dark' }]; },
  donut: (w, h) => { const t = Math.min(w, h) * 0.25; return one(`${ell(w / 2, h / 2, w / 2, h / 2)} ${ell(w / 2, h / 2, w / 2 - t, h / 2 - t)}`, { evenodd: true }); },
  noSmoking: (w, h) => {
    const t = Math.min(w, h) * 0.1875;
    const bar = rotPts([[t * 0.2, h / 2 - t / 2], [w - t * 0.2, h / 2 - t / 2], [w - t * 0.2, h / 2 + t / 2], [t * 0.2, h / 2 + t / 2]], w / 2, h / 2, 45);
    return [{ d: `${ell(w / 2, h / 2, w / 2, h / 2)} ${ell(w / 2, h / 2, w / 2 - t, h / 2 - t)}`, evenodd: true }, { d: P(bar) }];
  },
  blockArc: (w, h) => { const t = Math.min(w, h) * 0.25; return one(`M0,${r2(h / 2)} A${r2(w / 2)},${r2(h / 2)} 0 0,1 ${r2(w)},${r2(h / 2)} L${r2(w - t)},${r2(h / 2)} A${r2(w / 2 - t)},${r2(h / 2 - t)} 0 0,0 ${r2(t)},${r2(h / 2)} Z`); },
  foldedCorner: (w, h) => { const d = Math.min(w, h) * 0.1667; return [{ d: P([[0, 0], [w, 0], [w, h - d], [w - d, h], [0, h]]) }, { d: P([[w, h - d], [w - d * 0.8, h - d * 0.8], [w - d, h]]), shade: 'dark' }]; },
  smileyFace: (w, h) => [
    { d: ell(w / 2, h / 2, w / 2, h / 2) },
    { d: `${ell(w * 0.35, h * 0.37, w * 0.055, h * 0.055)} ${ell(w * 0.65, h * 0.37, w * 0.055, h * 0.055)}`, shade: 'dark' },
    { d: `M${r2(w * 0.28)},${r2(h * 0.62)} Q${r2(w / 2)},${r2(h * 0.84)} ${r2(w * 0.72)},${r2(h * 0.62)}`, line: true },
  ],
  heart: (w, h) => one(`M${r2(w / 2)},${r2(h * 0.24)} C${r2(w * 0.42)},${r2(-h * 0.06)} 0,0 0,${r2(h * 0.3)} C0,${r2(h * 0.62)} ${r2(w * 0.4)},${r2(h * 0.78)} ${r2(w / 2)},${r2(h)} C${r2(w * 0.6)},${r2(h * 0.78)} ${r2(w)},${r2(h * 0.62)} ${r2(w)},${r2(h * 0.3)} C${r2(w)},0 ${r2(w * 0.58)},${r2(-h * 0.06)} ${r2(w / 2)},${r2(h * 0.24)} Z`),
  lightningBolt: (w, h) => one(grid([[8458, 0], [0, 3923], [7564, 8416], [4993, 9720], [12197, 13904], [9987, 14934], [21600, 21600], [14768, 12911], [16558, 12016], [11030, 6490], [12831, 5739]], w, h)),
  sun: (w, h) => {
    const cx = w / 2;
    const cy = h / 2;
    const rays = [];
    for (let k = 0; k < 8; k++) {
      const a = (k * Math.PI) / 4;
      const p = (rr, da) => [cx + (w / 2) * rr * Math.cos(a + da), cy + (h / 2) * rr * Math.sin(a + da)];
      rays.push(P([p(1, 0), p(0.66, 0.16), p(0.66, -0.16)]));
    }
    return [{ d: ell(cx, cy, w * 0.26, h * 0.26) }, { d: rays.join(' ') }];
  },
  moon: (w, h) => one(`M${r2(w)},0 A${r2(w)},${r2(h / 2)} 0 0,0 ${r2(w)},${r2(h)} A${r2(w / 2)},${r2(h / 2)} 0 0,1 ${r2(w)},0 Z`),
  cloud: (w, h) => one(cloudD(w, h)),
  arc: (w, h) => one(`M${r2(w / 2)},0 A${r2(w / 2)},${r2(h / 2)} 0 0,1 ${r2(w)},${r2(h / 2)}`, { line: true }),
  bracketPair: (w, h) => { const r = Math.min(w, h) * 0.1667; return one(`M${r2(r)},${r2(h)} A${r2(r)},${r2(r)} 0 0,1 0,${r2(h - r)} L0,${r2(r)} A${r2(r)},${r2(r)} 0 0,1 ${r2(r)},0 M${r2(w - r)},0 A${r2(r)},${r2(r)} 0 0,1 ${r2(w)},${r2(r)} L${r2(w)},${r2(h - r)} A${r2(r)},${r2(r)} 0 0,1 ${r2(w - r)},${r2(h)}`, { line: true }); },
  bracePair: (w, h) => { const r = Math.min(w * 0.1, h * 0.08); return one(`M${r2(2 * r)},${r2(h)} Q${r2(r)},${r2(h)} ${r2(r)},${r2(h - r)} L${r2(r)},${r2(h / 2 + r)} Q${r2(r)},${r2(h / 2)} 0,${r2(h / 2)} Q${r2(r)},${r2(h / 2)} ${r2(r)},${r2(h / 2 - r)} L${r2(r)},${r2(r)} Q${r2(r)},0 ${r2(2 * r)},0 M${r2(w - 2 * r)},0 Q${r2(w - r)},0 ${r2(w - r)},${r2(r)} L${r2(w - r)},${r2(h / 2 - r)} Q${r2(w - r)},${r2(h / 2)} ${r2(w)},${r2(h / 2)} Q${r2(w - r)},${r2(h / 2)} ${r2(w - r)},${r2(h / 2 + r)} L${r2(w - r)},${r2(h - r)} Q${r2(w - r)},${r2(h)} ${r2(w - 2 * r)},${r2(h)}`, { line: true }); },
  leftBracket: (w, h) => { const r = Math.min(w, h * 0.1); return one(`M${r2(w)},${r2(h)} A${r2(w)},${r2(r)} 0 0,1 0,${r2(h - r)} L0,${r2(r)} A${r2(w)},${r2(r)} 0 0,1 ${r2(w)},0`, { line: true }); },
  rightBracket: (w, h) => { const r = Math.min(w, h * 0.1); return one(`M0,0 A${r2(w)},${r2(r)} 0 0,1 ${r2(w)},${r2(r)} L${r2(w)},${r2(h - r)} A${r2(w)},${r2(r)} 0 0,1 0,${r2(h)}`, { line: true }); },
  leftBrace: (w, h) => one(`M${r2(w)},${r2(h)} Q${r2(w / 2)},${r2(h)} ${r2(w / 2)},${r2(h * 0.9)} L${r2(w / 2)},${r2(h * 0.6)} Q${r2(w / 2)},${r2(h / 2)} 0,${r2(h / 2)} Q${r2(w / 2)},${r2(h / 2)} ${r2(w / 2)},${r2(h * 0.4)} L${r2(w / 2)},${r2(h * 0.1)} Q${r2(w / 2)},0 ${r2(w)},0`, { line: true }),
  rightBrace: (w, h) => one(`M0,0 Q${r2(w / 2)},0 ${r2(w / 2)},${r2(h * 0.1)} L${r2(w / 2)},${r2(h * 0.4)} Q${r2(w / 2)},${r2(h / 2)} ${r2(w)},${r2(h / 2)} Q${r2(w / 2)},${r2(h / 2)} ${r2(w / 2)},${r2(h * 0.6)} L${r2(w / 2)},${r2(h * 0.9)} Q${r2(w / 2)},${r2(h)} 0,${r2(h)}`, { line: true }),
  // 블록 화살표
  rightArrow: (w, h) => one(arrowRight(w, h)),
  leftArrow: (w, h) => { const hl = Math.min(w, h) * 0.5; return one(P([[w, h * 0.25], [hl, h * 0.25], [hl, 0], [0, h / 2], [hl, h], [hl, h * 0.75], [w, h * 0.75]])); },
  upArrow: (w, h) => { const hl = Math.min(w, h) * 0.5; return one(P([[w / 2, 0], [w, hl], [w * 0.75, hl], [w * 0.75, h], [w * 0.25, h], [w * 0.25, hl], [0, hl]])); },
  downArrow: (w, h) => one(arrowDown(w, h)),
  leftRightArrow: (w, h) => { const hl = Math.min(w / 2, h) * 0.5; return one(P([[0, h / 2], [hl, 0], [hl, h / 4], [w - hl, h / 4], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * 0.75], [hl, h * 0.75], [hl, h]])); },
  upDownArrow: (w, h) => { const hl = Math.min(w, h / 2) * 0.5; return one(P([[w / 2, 0], [w, hl], [w * 0.75, hl], [w * 0.75, h - hl], [w, h - hl], [w / 2, h], [0, h - hl], [w * 0.25, h - hl], [w * 0.25, hl], [0, hl]])); },
  quadArrow: (w, h) => {
    const s = Math.min(w, h);
    const aw = s * 0.1125;
    const hw = s * 0.225;
    const hl = s * 0.225;
    const cx = w / 2;
    const cy = h / 2;
    return one(P([[cx, 0], [cx + hw, hl], [cx + aw, hl], [cx + aw, cy - aw], [w - hl, cy - aw], [w - hl, cy - hw], [w, cy], [w - hl, cy + hw], [w - hl, cy + aw], [cx + aw, cy + aw], [cx + aw, h - hl], [cx + hw, h - hl], [cx, h], [cx - hw, h - hl], [cx - aw, h - hl], [cx - aw, cy + aw], [hl, cy + aw], [hl, cy + hw], [0, cy], [hl, cy - hw], [hl, cy - aw], [cx - aw, cy - aw], [cx - aw, hl], [cx - hw, hl]]));
  },
  notchedRightArrow: (w, h) => { const hl = Math.min(w, h) * 0.5; return one(P([[0, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * 0.75], [0, h * 0.75], [hl * 0.5, h / 2]])); },
  stripedRightArrow: (w, h) => { const hl = Math.min(w, h) * 0.5; const u = Math.min(w, h) * 0.0625; return one(`${rectD(0, h * 0.25, u, h * 0.5)} ${rectD(2 * u, h * 0.25, u, h * 0.5)} ${P([[4 * u, h * 0.25], [w - hl, h * 0.25], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h * 0.75], [4 * u, h * 0.75]])}`); },
  chevron: (w, h) => { const a = Math.min(w, h) * 0.5; return one(P([[0, 0], [w - a, 0], [w, h / 2], [w - a, h], [0, h], [a, h / 2]])); },
  homePlate: (w, h) => { const a = Math.min(w, h) * 0.5; return one(P([[0, 0], [w - a, 0], [w, h / 2], [w - a, h], [0, h]])); },
  bentUpArrow: (w, h) => { const s = Math.min(w, h); const t = s * 0.25; const hw = s * 0.25; return one(P([[0, h - t], [w - hw - t / 2, h - t], [w - hw - t / 2, hw], [w - 2 * hw, hw], [w - hw, 0], [w, hw], [w - hw + t / 2, hw], [w - hw + t / 2, h], [0, h]])); },
  uturnArrow: (w, h) => {
    const s = Math.min(w, h);
    const t = s * 0.25;
    const r = w * 0.35;
    return one(`M0,${r2(h)} L0,${r2(r)} A${r2(r)},${r2(r)} 0 0,1 ${r2(2 * r)},${r2(r)} L${r2(2 * r)},${r2(h * 0.55)} L${r2(2 * r + t * 0.75)},${r2(h * 0.55)} L${r2(2 * r - t / 2)},${r2(h * 0.8)} L${r2(2 * r - t * 1.75)},${r2(h * 0.55)} L${r2(2 * r - t)},${r2(h * 0.55)} L${r2(2 * r - t)},${r2(r)} A${r2(r - t)},${r2(r - t)} 0 0,0 ${r2(t)},${r2(r)} L${r2(t)},${r2(h)} Z`);
  },
  // 수식 도형
  mathPlus: (w, h) => { const t = Math.min(w, h) * 0.2352; return one(plusD(w * 0.07, h * 0.07, w * 0.86, h * 0.86, t)); },
  mathMinus: (w, h) => { const t = Math.min(w, h) * 0.2352; return one(rectD(w * 0.07, h / 2 - t / 2, w * 0.86, t)); },
  mathMultiply: (w, h) => { const s = Math.min(w, h); const t = s * 0.2352; const L = s * 0.9; const base = [[w / 2 - t / 2, h / 2 - L / 2], [w / 2 + t / 2, h / 2 - L / 2], [w / 2 + t / 2, h / 2 + L / 2], [w / 2 - t / 2, h / 2 + L / 2]]; const bar2 = [[w / 2 - L / 2, h / 2 - t / 2], [w / 2 + L / 2, h / 2 - t / 2], [w / 2 + L / 2, h / 2 + t / 2], [w / 2 - L / 2, h / 2 + t / 2]]; return [{ d: P(rotPts(base, w / 2, h / 2, 45)) }, { d: P(rotPts(bar2, w / 2, h / 2, 45)) }]; },
  mathDivide: (w, h) => { const t = Math.min(w, h) * 0.2352; return [{ d: rectD(w * 0.07, h / 2 - t / 2, w * 0.86, t) }, { d: `${ell(w / 2, h / 2 - t * 1.5, t * 0.55, t * 0.55)} ${ell(w / 2, h / 2 + t * 1.5, t * 0.55, t * 0.55)}` }]; },
  mathEqual: (w, h) => { const t = Math.min(w, h) * 0.2352; const g = t * 0.5; return one(`${rectD(w * 0.07, h / 2 - g / 2 - t, w * 0.86, t)} ${rectD(w * 0.07, h / 2 + g / 2, w * 0.86, t)}`); },
  mathNotEqual: (w, h) => { const t = Math.min(w, h) * 0.2352; const g = t * 0.5; const s = [[w * 0.55, h * 0.05], [w * 0.55 + t * 0.8, h * 0.05], [w * 0.45, h * 0.95], [w * 0.45 - t * 0.8, h * 0.95]]; return [{ d: `${rectD(w * 0.07, h / 2 - g / 2 - t, w * 0.86, t)} ${rectD(w * 0.07, h / 2 + g / 2, w * 0.86, t)}` }, { d: P(s) }]; },
  // 순서도
  flowChartProcess: (w, h) => one(rectD(0, 0, w, h)),
  flowChartAlternateProcess: (w, h) => one(roundRectD(w, h, Math.min(w, h) * 0.1667)),
  flowChartDecision: (w, h) => one(P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]])),
  flowChartInputOutput: (w, h) => one(P([[w * 0.2, 0], [w, 0], [w * 0.8, h], [0, h]])),
  flowChartPredefinedProcess: (w, h) => [{ d: rectD(0, 0, w, h) }, { d: `M${r2(w * 0.125)},0 L${r2(w * 0.125)},${r2(h)} M${r2(w * 0.875)},0 L${r2(w * 0.875)},${r2(h)}`, line: true }],
  flowChartInternalStorage: (w, h) => [{ d: rectD(0, 0, w, h) }, { d: `M${r2(w * 0.125)},0 L${r2(w * 0.125)},${r2(h)} M0,${r2(h * 0.125)} L${r2(w)},${r2(h * 0.125)}`, line: true }],
  flowChartDocument: (w, h) => one(`M0,0 L${r2(w)},0 L${r2(w)},${r2(h * 0.8)} C${r2(w * 0.72)},${r2(h * 0.62)} ${r2(w * 0.52)},${r2(h * 0.74)} ${r2(w * 0.5)},${r2(h * 0.8)} C${r2(w * 0.3)},${r2(h * 1.02)} ${r2(w * 0.12)},${r2(h * 0.96)} 0,${r2(h * 0.86)} Z`),
  flowChartMultidocument: (w, h) => { const d = (x, y, ww, hh) => `M${r2(x)},${r2(y)} L${r2(x + ww)},${r2(y)} L${r2(x + ww)},${r2(y + hh * 0.8)} C${r2(x + ww * 0.72)},${r2(y + hh * 0.62)} ${r2(x + ww * 0.52)},${r2(y + hh * 0.74)} ${r2(x + ww * 0.5)},${r2(y + hh * 0.8)} C${r2(x + ww * 0.3)},${r2(y + hh * 1.02)} ${r2(x + ww * 0.12)},${r2(y + hh * 0.96)} ${r2(x)},${r2(y + hh * 0.86)} Z`; return [{ d: d(w * 0.12, 0, w * 0.88, h * 0.8) }, { d: d(w * 0.06, h * 0.08, w * 0.88, h * 0.8) }, { d: d(0, h * 0.16, w * 0.88, h * 0.84) }]; },
  flowChartTerminator: (w, h) => { const r = Math.min(h / 2, w / 2); return one(`M${r2(r)},0 L${r2(w - r)},0 A${r2(r)},${r2(h / 2)} 0 0,1 ${r2(w - r)},${r2(h)} L${r2(r)},${r2(h)} A${r2(r)},${r2(h / 2)} 0 0,1 ${r2(r)},0 Z`); },
  flowChartPreparation: (w, h) => one(P([[w * 0.2, 0], [w * 0.8, 0], [w, h / 2], [w * 0.8, h], [w * 0.2, h], [0, h / 2]])),
  flowChartManualInput: (w, h) => one(P([[0, h * 0.2], [w, 0], [w, h], [0, h]])),
  flowChartManualOperation: (w, h) => one(P([[0, 0], [w, 0], [w * 0.8, h], [w * 0.2, h]])),
  flowChartConnector: (w, h) => one(ell(w / 2, h / 2, w / 2, h / 2)),
  flowChartOffpageConnector: (w, h) => one(P([[0, 0], [w, 0], [w, h * 0.8], [w / 2, h], [0, h * 0.8]])),
  flowChartPunchedCard: (w, h) => one(P([[w * 0.2, 0], [w, 0], [w, h], [0, h], [0, h * 0.2]])),
  flowChartPunchedTape: (w, h) => one(`M0,${r2(h * 0.1)} Q${r2(w * 0.25)},${r2(h * 0.3)} ${r2(w * 0.5)},${r2(h * 0.1)} T${r2(w)},${r2(h * 0.1)} L${r2(w)},${r2(h * 0.9)} Q${r2(w * 0.75)},${r2(h * 0.7)} ${r2(w * 0.5)},${r2(h * 0.9)} T0,${r2(h * 0.9)} Z`),
  flowChartSummingJunction: (w, h) => { const k = 0.3536; return [{ d: ell(w / 2, h / 2, w / 2, h / 2) }, { d: `M${r2(w / 2 - w * k)},${r2(h / 2 - h * k)} L${r2(w / 2 + w * k)},${r2(h / 2 + h * k)} M${r2(w / 2 + w * k)},${r2(h / 2 - h * k)} L${r2(w / 2 - w * k)},${r2(h / 2 + h * k)}`, line: true }]; },
  flowChartOr: (w, h) => [{ d: ell(w / 2, h / 2, w / 2, h / 2) }, { d: `M${r2(w / 2)},0 L${r2(w / 2)},${r2(h)} M0,${r2(h / 2)} L${r2(w)},${r2(h / 2)}`, line: true }],
  flowChartCollate: (w, h) => one(P([[0, 0], [w, 0], [0, h], [w, h]])),
  flowChartSort: (w, h) => [{ d: P([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]) }, { d: `M0,${r2(h / 2)} L${r2(w)},${r2(h / 2)}`, line: true }],
  flowChartExtract: (w, h) => one(P([[w / 2, 0], [w, h], [0, h]])),
  flowChartMerge: (w, h) => one(P([[0, 0], [w, 0], [w / 2, h]])),
  flowChartOnlineStorage: (w, h) => one(`M${r2(w * 0.17)},0 L${r2(w)},0 A${r2(w * 0.17)},${r2(h / 2)} 0 0,0 ${r2(w)},${r2(h)} L${r2(w * 0.17)},${r2(h)} A${r2(w * 0.17)},${r2(h / 2)} 0 0,1 ${r2(w * 0.17)},0 Z`),
  flowChartDelay: (w, h) => one(`M0,0 L${r2(w / 2)},0 A${r2(w / 2)},${r2(h / 2)} 0 0,1 ${r2(w / 2)},${r2(h)} L0,${r2(h)} Z`),
  flowChartMagneticDisk: (w, h) => { const ry = h * 0.17; return [{ d: `M0,${r2(ry)} A${r2(w / 2)},${r2(ry)} 0 0,1 ${r2(w)},${r2(ry)} L${r2(w)},${r2(h - ry)} A${r2(w / 2)},${r2(ry)} 0 0,1 0,${r2(h - ry)} Z` }, { d: `M0,${r2(ry)} A${r2(w / 2)},${r2(ry)} 0 0,0 ${r2(w)},${r2(ry)}`, line: true }]; },
  flowChartDisplay: (w, h) => one(`M0,${r2(h / 2)} L${r2(w * 0.17)},0 L${r2(w * 0.83)},0 A${r2(w * 0.17)},${r2(h / 2)} 0 0,1 ${r2(w * 0.83)},${r2(h)} L${r2(w * 0.17)},${r2(h)} Z`),
  // 별 및 현수막
  irregularSeal1: (w, h) => one(grid([[10800, 5800], [14522, 0], [14155, 5325], [18380, 4457], [16702, 7315], [21097, 8137], [17607, 10475], [21600, 13290], [16837, 12942], [18145, 18095], [14020, 14457], [13247, 19737], [10532, 14935], [8485, 21600], [7715, 15627], [4762, 17617], [5667, 13937], [135, 14587], [3722, 11775], [0, 8615], [4627, 7617], [370, 2295], [7312, 6320], [8352, 2295]], w, h)),
  star4: (w, h) => one(starD(4, 0.25, w, h)),
  star5: (w, h) => one(starD(5, 0.382, w, h)),
  star6: (w, h) => one(starD(6, 0.577, w, h)),
  star7: (w, h) => one(starD(7, 0.692, w, h)),
  star8: (w, h) => one(starD(8, 0.765, w, h)),
  star10: (w, h) => one(starD(10, 0.851, w, h)),
  star12: (w, h) => one(starD(12, 0.75, w, h)),
  star16: (w, h) => one(starD(16, 0.75, w, h)),
  star24: (w, h) => one(starD(24, 0.75, w, h)),
  star32: (w, h) => one(starD(32, 0.75, w, h)),
  wave: (w, h) => one(`M0,${r2(h * 0.13)} C${r2(w * 0.33)},${r2(-h * 0.13)} ${r2(w * 0.67)},${r2(h * 0.39)} ${r2(w)},${r2(h * 0.13)} L${r2(w)},${r2(h * 0.87)} C${r2(w * 0.67)},${r2(h * 1.13)} ${r2(w * 0.33)},${r2(h * 0.61)} 0,${r2(h * 0.87)} Z`),
  doubleWave: (w, h) => one(`M0,${r2(h * 0.07)} C${r2(w * 0.17)},${r2(-h * 0.07)} ${r2(w * 0.33)},${r2(h * 0.2)} ${r2(w * 0.5)},${r2(h * 0.07)} C${r2(w * 0.67)},${r2(-h * 0.07)} ${r2(w * 0.83)},${r2(h * 0.2)} ${r2(w)},${r2(h * 0.07)} L${r2(w)},${r2(h * 0.93)} C${r2(w * 0.83)},${r2(h * 1.07)} ${r2(w * 0.67)},${r2(h * 0.8)} ${r2(w * 0.5)},${r2(h * 0.93)} C${r2(w * 0.33)},${r2(h * 1.07)} ${r2(w * 0.17)},${r2(h * 0.8)} 0,${r2(h * 0.93)} Z`),
  ribbon2: (w, h) => { const e = w * 0.125; const t = h * 0.25; return [{ d: P([[0, t], [w * 0.25, t], [w * 0.25, h], [0, h], [e, (t + h) / 2]]), shade: 'dark' }, { d: P([[w, t], [w * 0.75, t], [w * 0.75, h], [w, h], [w - e, (t + h) / 2]]), shade: 'dark' }, { d: rectD(w * 0.18, 0, w * 0.64, h * 0.75) }]; },
  horizontalScroll: (w, h) => { const r = Math.min(w, h) * 0.0625; return [{ d: `M${r2(r * 2)},${r2(r * 2)} L${r2(w - r)},${r2(r * 2)} L${r2(w - r)},${r2(h - r * 2)} L${r2(r * 2)},${r2(h - r * 2)} Z` }, { d: `${ell(r * 2, h / 2, r * 2, h / 2 - r * 2)} ${ell(w - r, r * 2, r, r * 2)}`, shade: 'light' }]; },
  // 설명선
  wedgeRectCallout: (w, h) => one(P([[0, 0], [w, 0], [w, h], [w * 0.4167, h], [w * 0.29, h * 1.125], [w * 0.1667, h], [0, h]])),
  wedgeRoundRectCallout: (w, h) => { const r = Math.min(w, h) * 0.1667; return one(`M${r2(r)},0 L${r2(w - r)},0 A${r2(r)},${r2(r)} 0 0,1 ${r2(w)},${r2(r)} L${r2(w)},${r2(h - r)} A${r2(r)},${r2(r)} 0 0,1 ${r2(w - r)},${r2(h)} L${r2(w * 0.4167)},${r2(h)} L${r2(w * 0.29)},${r2(h * 1.125)} L${r2(w * 0.1667)},${r2(h)} L${r2(r)},${r2(h)} A${r2(r)},${r2(r)} 0 0,1 0,${r2(h - r)} L0,${r2(r)} A${r2(r)},${r2(r)} 0 0,1 ${r2(r)},0 Z`); },
  wedgeEllipseCallout: (w, h) => {
    const p = (deg) => { const a = (deg * Math.PI) / 180; return [w / 2 + (w / 2) * Math.cos(a), h / 2 + (h / 2) * Math.sin(a)]; };
    const [a1, a2] = [p(100), p(125)];
    return one(`M${r2(w * 0.29)},${r2(h * 1.125)} L${r2(a1[0])},${r2(a1[1])} A${r2(w / 2)},${r2(h / 2)} 0 1,0 ${r2(a2[0])},${r2(a2[1])} Z`);
  },
  cloudCallout: (w, h) => one(`${cloudD(w, h)} ${ell(w * 0.3, h * 0.92, w * 0.05, h * 0.05)} ${ell(w * 0.24, h * 1.02, w * 0.03, h * 0.03)} ${ell(w * 0.2, h * 1.09, w * 0.018, h * 0.018)}`),
  borderCallout1: (w, h) => [{ d: rectD(0, 0, w, h) }, { d: `M0,${r2(h * 0.19)} L${r2(-w * 0.38)},${r2(h * 1.12)}`, line: true }],
};
// 옛 이름
GEOM.arrow = GEOM.rightArrow;
GEOM.textbox = GEOM.rect;

// 선 종류 (채우기 없음, 끝 화살표 가능)
export const LINE_KINDS = new Set(['line', 'bentConnector3', 'curvedConnector3']);
export function linePath(kind, w, h, flip, flipV) {
  const [x1, x2] = flip ? [w, 0] : [0, w];
  const [y1, y2] = flipV ? [h, 0] : [0, h];
  if (kind === 'bentConnector3') { const mx = (x1 + x2) / 2; return `M${x1},${y1} L${mx},${y1} L${mx},${y2} L${x2},${y2}`; }
  if (kind === 'curvedConnector3') { const mx = (x1 + x2) / 2; return `M${x1},${y1} C${mx},${y1} ${mx},${y2} ${x2},${y2}`; }
  return `M${x1},${y1} L${x2},${y2}`;
}

/** 도형 갤러리 (엑셀 [삽입] → [도형] 분류) */
export const SHAPE_GROUPS = [
  ['선', [['line', '선'], ['lineArrow', '선 화살표'], ['lineDblArrow', '양쪽 화살표 선'], ['bentConnector3', '꺾인 연결선'], ['bentArrow3', '꺾인 화살표 연결선'], ['curvedConnector3', '곡선 연결선'], ['curvedArrow3', '곡선 화살표 연결선']]],
  ['사각형', [['rect', '사각형'], ['roundRect', '둥근 사각형'], ['snip1Rect', '한쪽 모서리가 잘린 사각형'], ['snip2SameRect', '위쪽 모서리가 잘린 사각형'], ['snip2DiagRect', '대각선 방향 모서리가 잘린 사각형'], ['snipRoundRect', '한쪽 모서리는 잘리고 다른 쪽은 둥근 사각형'], ['round1Rect', '한쪽 모서리가 둥근 사각형'], ['round2SameRect', '위쪽 모서리가 둥근 사각형'], ['round2DiagRect', '대각선 방향 모서리가 둥근 사각형']]],
  ['기본 도형', [['textbox', '텍스트 상자'], ['ellipse', '타원'], ['triangle', '이등변 삼각형'], ['rtTriangle', '직각 삼각형'], ['parallelogram', '평행 사변형'], ['trapezoid', '사다리꼴'], ['diamond', '다이아몬드'], ['pentagon', '오각형'], ['hexagon', '육각형'], ['heptagon', '칠각형'], ['octagon', '팔각형'], ['decagon', '십각형'], ['dodecagon', '십이각형'], ['pie', '원형'], ['chord', '현'], ['teardrop', '눈물 방울'], ['frame', '액자'], ['halfFrame', 'L 도형 액자'], ['corner', 'L 도형'], ['diagStripe', '대각선 줄무늬'], ['plus', '십자형'], ['plaque', '배지'], ['can', '원통형'], ['cube', '정육면체'], ['bevel', '빗면'], ['donut', '도넛'], ['noSmoking', '금지'], ['blockArc', '막힌 원호'], ['foldedCorner', '모서리가 접힌 도형'], ['smileyFace', '웃는 얼굴'], ['heart', '하트'], ['lightningBolt', '번개'], ['sun', '해'], ['moon', '달'], ['cloud', '구름'], ['arc', '원호'], ['bracketPair', '양쪽 대괄호'], ['bracePair', '양쪽 중괄호'], ['leftBracket', '왼쪽 대괄호'], ['rightBracket', '오른쪽 대괄호'], ['leftBrace', '왼쪽 중괄호'], ['rightBrace', '오른쪽 중괄호']]],
  ['블록 화살표', [['rightArrow', '오른쪽 화살표'], ['leftArrow', '왼쪽 화살표'], ['upArrow', '위쪽 화살표'], ['downArrow', '아래쪽 화살표'], ['leftRightArrow', '왼쪽/오른쪽 화살표'], ['upDownArrow', '위쪽/아래쪽 화살표'], ['quadArrow', '왼쪽/오른쪽/위쪽/아래쪽 화살표'], ['bentUpArrow', '위쪽 굽은 화살표'], ['uturnArrow', 'U자형 화살표'], ['stripedRightArrow', '줄무늬가 있는 오른쪽 화살표'], ['notchedRightArrow', '톱니 모양의 오른쪽 화살표'], ['homePlate', '오각형 화살표'], ['chevron', '갈매기형 수장']]],
  ['수식 도형', [['mathPlus', '더하기 기호'], ['mathMinus', '빼기 기호'], ['mathMultiply', '곱하기 기호'], ['mathDivide', '나누기 기호'], ['mathEqual', '등호'], ['mathNotEqual', '부등호']]],
  ['순서도', [['flowChartProcess', '순서도: 프로세스'], ['flowChartAlternateProcess', '순서도: 대체 프로세스'], ['flowChartDecision', '순서도: 판단'], ['flowChartInputOutput', '순서도: 데이터'], ['flowChartPredefinedProcess', '순서도: 미리 정의된 프로세스'], ['flowChartInternalStorage', '순서도: 내부 저장소'], ['flowChartDocument', '순서도: 문서'], ['flowChartMultidocument', '순서도: 다중 문서'], ['flowChartTerminator', '순서도: 수행의 시작/종료'], ['flowChartPreparation', '순서도: 준비'], ['flowChartManualInput', '순서도: 수동 입력'], ['flowChartManualOperation', '순서도: 수동 연산'], ['flowChartConnector', '순서도: 연결자'], ['flowChartOffpageConnector', '순서도: 페이지 연결자'], ['flowChartPunchedCard', '순서도: 카드'], ['flowChartPunchedTape', '순서도: 천공 테이프'], ['flowChartSummingJunction', '순서도: 가산 접합'], ['flowChartOr', '순서도: 논리합'], ['flowChartCollate', '순서도: 대조'], ['flowChartSort', '순서도: 정렬'], ['flowChartExtract', '순서도: 추출'], ['flowChartMerge', '순서도: 병합'], ['flowChartOnlineStorage', '순서도: 저장 데이터'], ['flowChartDelay', '순서도: 지연'], ['flowChartMagneticDisk', '순서도: 자기 디스크'], ['flowChartDisplay', '순서도: 표시']]],
  ['별 및 현수막', [['irregularSeal1', '폭발: 8pt'], ['star4', '별: 꼭짓점 4개'], ['star5', '별: 꼭짓점 5개'], ['star6', '별: 꼭짓점 6개'], ['star7', '별: 꼭짓점 7개'], ['star8', '별: 꼭짓점 8개'], ['star10', '별: 꼭짓점 10개'], ['star12', '별: 꼭짓점 12개'], ['star16', '별: 꼭짓점 16개'], ['star24', '별: 꼭짓점 24개'], ['star32', '별: 꼭짓점 32개'], ['ribbon2', '리본: 위로 기울어짐'], ['horizontalScroll', '두루마리 모양: 가로로 말림'], ['wave', '물결'], ['doubleWave', '이중 물결']]],
  ['설명선', [['wedgeRectCallout', '말풍선: 사각형'], ['wedgeRoundRectCallout', '말풍선: 모서리가 둥근 사각형'], ['wedgeEllipseCallout', '말풍선: 타원형'], ['cloudCallout', '생각 풍선: 구름 모양'], ['borderCallout1', '설명선: 선']]],
];

/** 모든 도형 (갤러리 순서) */
export const SHAPE_KINDS = SHAPE_GROUPS.flatMap(([, list]) => list.map(([id, label]) => ({ id, label })));
export const LINE_SHAPES = LINE_KINDS;
