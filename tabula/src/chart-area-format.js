import { esc } from './xml.js';
import { safeUrl } from './safe-html.js';
import { SHAPE_DASH_OPTIONS, SHAPE_PATTERN_OPTIONS } from './shape-format.js';

// Chart/plot area formatting is independent of data-series formatting. Units are CSS px.
const choice = (v, list, fallback) => list.includes(v) ? v : fallback;
const number = (v, fallback, min, max) => Number.isFinite(Number(v)) && v !== null && v !== '' ? Math.max(min, Math.min(max, Number(v))) : fallback;
export function chartAreaColor(v, fallback = '#4472c4') {
  if (/^#[0-9a-f]{6}$/i.test(v ?? '')) return v.toLowerCase();
  if (/^#[0-9a-f]{3}$/i.test(v ?? '')) return '#' + v.slice(1).split('').map(c => c + c).join('').toLowerCase();
  return fallback;
}
export function normalizeAreaGradient(value, fallback = '#4472c4') {
  const g = value ?? {}, stops = Array.isArray(g.stops) ? g.stops.slice(0, 16).filter(s => Array.isArray(s) && Number.isFinite(Number(s[0]))).map(s => [number(s[0], 0, 0, 1), chartAreaColor(s[1], fallback), number(s[2], 1, 0, 1)]).sort((a, b) => a[0] - b[0]) : [];
  return { type: choice(g.type, ['linear', 'radial'], 'linear'), ang: number(g.ang, 90, 0, 360), stops: stops.length >= 2 ? stops : [[0, fallback, 1], [1, '#ffffff', 1]] };
}
export function chartAreaImageSize(src) {
  // PNG IHDR dimensions are available without DOM/decoding and survive standard XLSX media round trips.
  if (/^data:image\/png;base64,/i.test(src ?? '')) {
    try { const b=atob(src.slice(src.indexOf(',')+1,src.indexOf(',')+1+44));if(b.length>=24&&b.charCodeAt(0)===137&&b.slice(1,4)==='PNG'){const read=i=>((b.charCodeAt(i)*16777216)+(b.charCodeAt(i+1)<<16)+(b.charCodeAt(i+2)<<8)+b.charCodeAt(i+3));return {width:read(16),height:read(20)};} } catch {}
  }
  return {width:128,height:128};
}
export function normalizeChartAreaFormat(value = {}, kind = 'chart') {
  const f = value ?? {}, fill = chartAreaColor(f.fill, kind === 'plot' ? '#ffffff' : '#ffffff');
  const image = safeUrl(f.picture?.src, 'image'), size=chartAreaImageSize(image);
  const bevel = v => ({ type: choice(v?.type, ['none', 'angle', 'circle', 'convex'], 'none'), w: number(v?.w, 6, 0, 72), h: number(v?.h, 6, 0, 72) });
  return {
    fillMode: choice(f.fillMode, ['auto', 'none', 'solid', 'gradient', 'picture', 'pattern'], 'auto'), fill, fillOpacity: number(f.fillOpacity, 1, 0, 1),
    grad: normalizeAreaGradient(f.grad, fill),
    picture: { src: image ?? '', width:number(f.picture?.width,size.width,1,32768),height:number(f.picture?.height,size.height,1,32768), mode: choice(f.picture?.mode, ['stretch', 'tile'], 'stretch'), scale: number(f.picture?.scale, 1, .05, 10), opacity: number(f.picture?.opacity, 1, 0, 1) },
    pattern: { preset: choice(f.pattern?.preset, SHAPE_PATTERN_OPTIONS.map(p => p[0]), 'pct25'), fg: chartAreaColor(f.pattern?.fg), bg: chartAreaColor(f.pattern?.bg, '#ffffff') },
    lineMode: choice(f.lineMode, ['auto', 'none', 'solid', 'gradient'], 'auto'), stroke: chartAreaColor(f.stroke, '#bfbfbf'), strokeWidth: number(f.strokeWidth, 1, 0, 40), strokeOpacity: number(f.strokeOpacity, 1, 0, 1),
    strokeGrad: normalizeAreaGradient(f.strokeGrad, chartAreaColor(f.stroke, '#bfbfbf')),
    compound: choice(f.compound, ['sng', 'dbl', 'tri', 'thickThin', 'thinThick'], 'sng'), dash: choice(f.dash, SHAPE_DASH_OPTIONS.map(p => p[0]), ''),
    lineCap: choice(f.lineCap, ['flat', 'rnd', 'sq'], 'flat'), lineJoin: choice(f.lineJoin, ['round', 'bevel', 'miter'], 'round'),
    shadow: f.shadow ? { dx: number(f.shadow.dx, 3, -100, 100), dy: number(f.shadow.dy, 3, -100, 100), blur: number(f.shadow.blur, 4, 0, 100), color: chartAreaColor(f.shadow.color, '#000000'), opacity: number(f.shadow.opacity, .3, 0, 1), scale: number(f.shadow.scale, 1, .05, 4) } : null,
    glow: f.glow ? { color: chartAreaColor(f.glow.color, '#ffc000'), size: number(f.glow.size, 6, 0, 100), opacity: number(f.glow.opacity, .5, 0, 1) } : null,
    soft: number(f.soft, 0, 0, 50),
    threeD: { bevelTop: bevel(f.threeD?.bevelTop), bevelBottom: bevel(f.threeD?.bevelBottom), depth: number(f.threeD?.depth, 0, 0, 100), contourWidth: number(f.threeD?.contourWidth, 0, 0, 20), contourColor: chartAreaColor(f.threeD?.contourColor, '#4472c4'), material: choice(f.threeD?.material, ['plastic', 'matte', 'metal'], 'plastic'), lightRig: choice(f.threeD?.lightRig, ['threePt', 'balanced', 'soft'], 'threePt'), lightAngle: number(f.threeD?.lightAngle, 45, 0, 360) },
  };
}
export function chartAreaFormat(chart, kind = 'chart') {
  const own = chart?.[kind === 'plot' ? 'plotAreaFormat' : 'chartAreaFormat'];
  if (own) return normalizeChartAreaFormat(own, kind);
  const fill = kind === 'plot' ? chart?.plotFill : chart?.fill;
  return normalizeChartAreaFormat({ fillMode: fill === 'none' || fill === 'transparent' ? 'none' : fill ? 'solid' : 'auto', fill, lineMode: kind !== 'plot' && chart?.border ? 'solid' : 'auto', stroke: chart?.border }, kind);
}
export function chartAreaFormatPatch(chart, kind, patch) {
  const key = kind === 'plot' ? 'plotAreaFormat' : 'chartAreaFormat';
  return { [key]: normalizeChartAreaFormat({ ...chartAreaFormat(chart, kind), ...patch }, kind) };
}

const DASH = { dot: [1, 2], dash: [4, 3], lgDash: [8, 3], dashDot: [4, 2, 1, 2], lgDashDot: [8, 2, 1, 2], lgDashDotDot: [8, 2, 1, 2, 1, 2], sysDash: [3, 1], sysDot: [1, 1], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1] };
function gradientSvg(g, id) {
  const stops = g.stops.map(([o, c, a]) => `<stop offset="${o}" stop-color="${c}" stop-opacity="${a}"/>`).join('');
  if (g.type === 'radial') return `<radialGradient id="${id}">${stops}</radialGradient>`;
  const angle = g.ang * Math.PI / 180, dx = Math.cos(angle) / 2, dy = Math.sin(angle) / 2;
  return `<linearGradient id="${id}" x1="${.5 - dx}" y1="${.5 - dy}" x2="${.5 + dx}" y2="${.5 + dy}">${stops}</linearGradient>`;
}
function patternSvg(p, id) {
  const spacing = p.preset === 'pct5' ? 12 : p.preset === 'pct50' ? 4 : 8;
  const paths = { horz: 'M0 0H8', vert: 'M0 0V8', cross: 'M0 0H8M0 0V8', dnDiag: 'M-1 -1L9 9', upDiag: 'M-1 9L9 -1', diagCross: 'M-1 -1L9 9M-1 9L9 -1' };
  return `<pattern id="${id}" width="${spacing}" height="${spacing}" patternUnits="userSpaceOnUse"><rect width="${spacing}" height="${spacing}" fill="${p.bg}"/>${paths[p.preset] ? `<path d="${paths[p.preset]}" fill="none" stroke="${p.fg}" stroke-width="1.2"/>` : `<circle cx="1" cy="1" r="1" fill="${p.fg}"/>`}</pattern>`;
}
function reliefSvg(f, rect) {
  const d = f.threeD, { x, y, w, h } = rect;
  const points = pts => pts.map(p => p.join(',')).join(' ');
  const polygon = (pts, color, opacity) => `<polygon points="${points(pts)}" fill="${color}" opacity="${opacity}"/>`;
  const a = d.lightAngle * Math.PI / 180, gain = d.lightRig === 'soft' ? .22 : d.lightRig === 'balanced' ? .35 : .55, material = d.material === 'metal' ? 1.4 : d.material === 'matte' ? .55 : 1;
  const amount = side => Math.min(.9, gain * material * (.4 + .6 * Math.abs(Math.cos(a - side))));
  let html = '';
  if (d.depth > 0) {
    const dx = d.depth * .6, dy = d.depth * .6;
    html += polygon([[x, y + h], [x + dx, y + h + dy], [x + w + dx, y + h + dy], [x + w, y + h]], f.fill, 1);
    html += polygon([[x + w, y], [x + w + dx, y + dy], [x + w + dx, y + h + dy], [x + w, y + h]], f.fill, 1);
    html += polygon([[x, y + h], [x + dx, y + h + dy], [x + w + dx, y + h + dy], [x + w, y + h]], '#000000', amount(Math.PI / 2));
    html += polygon([[x + w, y], [x + w + dx, y + dy], [x + w + dx, y + h + dy], [x + w, y + h]], '#000000', amount(0));
  }
  const bevel = (b, bottom) => {
    if (b.type === 'none' || !b.w || !b.h) return '';
    const bw = Math.min(w / 3, b.w), bh = Math.min(h / 3, b.h), steps = b.type === 'angle' ? 1 : 5;
    let s = '';
    for (let i = 0; i < steps; i++) {
      const t = i / steps, next = (i + 1) / steps, o = b.type === 'circle' ? Math.cos(t * Math.PI / 2) : b.type === 'convex' ? Math.sin((t + .1) * Math.PI / 2) : 1;
      const l = x + bw * t, r = x + w - bw * t, u = y + bh * t, v = y + h - bh * t;
      const L = x + bw * next, R = x + w - bw * next, U = y + bh * next, V = y + h - bh * next;
      s += polygon([[l,u],[r,u],[R,U],[L,U]], bottom ? '#000000' : '#ffffff', amount(Math.PI) * o);
      s += polygon([[l,u],[L,U],[L,V],[l,v]], bottom ? '#000000' : '#ffffff', amount(-Math.PI / 2) * o);
      s += polygon([[l,v],[L,V],[R,V],[r,v]], bottom ? '#ffffff' : '#000000', amount(Math.PI / 2) * o);
      s += polygon([[r,u],[r,v],[R,V],[R,U]], bottom ? '#ffffff' : '#000000', amount(0) * o);
    }
    return s;
  };
  return { back: html, front: bevel(d.bevelBottom, true) + bevel(d.bevelTop, false) + (d.contourWidth ? `<rect x="${x}" y="${y}" width="${w}" height="${h}" fill="none" stroke="${d.contourColor}" stroke-width="${d.contourWidth}"/>` : '') };
}

/** Sanitized, isolated SVG group. Effects apply only to the area, never data/labels. */
export function chartAreaSvg(value, rect, { id = 'area', kind = 'chart', rounded = false } = {}) {
  const f = normalizeChartAreaFormat(value, kind), x = number(rect.x, 0, -1e6, 1e6), y = number(rect.y, 0, -1e6, 1e6), w = number(rect.w, 0, 0, 1e6), h = number(rect.h, 0, 0, 1e6);
  if (!w || !h) return '';
  id = String(id).replace(/[^a-zA-Z0-9_-]/g, '');
  const defs = [], url = suffix => `url(#${id}${suffix})`;
  let fill = f.fillMode === 'auto' ? kind === 'plot' ? 'none' : '#ffffff' : f.fillMode === 'none' ? 'none' : f.fill;
  if (f.fillMode === 'gradient') { defs.push(gradientSvg(f.grad, id + 'fill')); fill = url('fill'); }
  if (f.fillMode === 'pattern') { defs.push(patternSvg(f.pattern, id + 'pattern')); fill = url('pattern'); }
  if (f.fillMode === 'picture') fill = 'none';
  let stroke = f.lineMode === 'none' || f.lineMode === 'auto' ? 'none' : f.stroke;
  if (f.lineMode === 'gradient') { defs.push(gradientSvg(f.strokeGrad, id + 'line')); stroke = url('line'); }
  const sw = f.strokeWidth, rx = rounded ? Math.min(8, w / 2, h / 2) : 0;
  const box = (inset = 0) => `x="${x + inset}" y="${y + inset}" width="${Math.max(0, w - inset * 2)}" height="${Math.max(0, h - inset * 2)}" rx="${Math.max(0, rx - inset)}"`;
  let body = `<rect ${box()} fill="${fill}" fill-opacity="${f.fillOpacity}"/>`;
  if (f.fillMode === 'picture' && f.picture.src) {
    defs.push(`<clipPath id="${id}clip"><rect ${box()}/></clipPath>`);
    if (f.picture.mode === 'tile') {
      const tw = f.picture.width * f.picture.scale, th=f.picture.height * f.picture.scale;
      defs.push(`<pattern id="${id}picture" x="${x}" y="${y}" width="${tw}" height="${th}" patternUnits="userSpaceOnUse"><image href="${esc(f.picture.src)}" width="${tw}" height="${th}" preserveAspectRatio="none"/></pattern>`);
      body += `<rect ${box()} fill="${url('picture')}" opacity="${f.picture.opacity * f.fillOpacity}"/>`;
    } else body += `<image href="${esc(f.picture.src)}" x="${x}" y="${y}" width="${w}" height="${h}" preserveAspectRatio="none" opacity="${f.picture.opacity * f.fillOpacity}" clip-path="${url('clip')}"/>`;
  }
  const relief = reliefSvg(f, {x,y,w,h}); body = relief.back + body + relief.front;
  if (stroke !== 'none' && sw > 0) {
    const dash = DASH[f.dash]?.map(n => n * sw).join(' '), cap = {flat:'butt',rnd:'round',sq:'square'}[f.lineCap];
    const lineAttrs = `fill="none" stroke="${stroke}" stroke-opacity="${f.strokeOpacity}" stroke-linecap="${cap}" stroke-linejoin="${f.lineJoin}"${dash ? ` stroke-dasharray="${dash}"` : ''}`;
    // Separate strokes preserve transparent gaps; painting over with the fill would fail on images/no-fill.
    const bands = {sng:[[.5,1]],dbl:[[1/6,1/3],[5/6,1/3]],tri:[[.1,.2],[.5,.2],[.9,.2]],thickThin:[[.25,.5],[.875,.25]],thinThick:[[.125,.25],[.75,.5]]}[f.compound];
    for (const [center, width] of bands) body += `<rect ${box(sw * center)} ${lineAttrs} stroke-width="${sw * width}"/>`;
  }
  const filters = [];
  if (f.shadow) {
    const s=f.shadow,cx=x+w/2,cy=y+h/2;
    defs.push(`<filter id="${id}shadow" x="-100%" y="-100%" width="300%" height="300%" color-interpolation-filters="sRGB"><feGaussianBlur in="SourceAlpha" stdDeviation="${s.blur / 2 / s.scale}" result="blur"/><feFlood flood-color="${s.color}" flood-opacity="${s.opacity}"/><feComposite in2="blur" operator="in"/></filter>`);
    body=`<g transform="translate(${s.dx+cx} ${s.dy+cy}) scale(${s.scale}) translate(${-cx} ${-cy})" pointer-events="none"><g filter="${url('shadow')}">${body}</g></g>${body}`;
  }
  if (f.glow?.size) filters.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${f.glow.size / 2}" result="blur"/><feFlood flood-color="${f.glow.color}" flood-opacity="${f.glow.opacity}"/><feComposite in2="blur" operator="in"/><feMerge><feMergeNode/><feMergeNode in="SourceGraphic"/></feMerge>`);
  if (f.soft) filters.push(`<feGaussianBlur stdDeviation="${f.soft / 3}"/>`);
  // Each filter wraps the previous result so combined effects are not accidentally discarded.
  filters.forEach((filter, i) => { const key = id + 'fx' + i; defs.push(`<filter id="${key}" x="-100%" y="-100%" width="300%" height="300%" color-interpolation-filters="sRGB">${filter}</filter>`); body = `<g filter="url(#${key})">${body}</g>`; });
  return `<g data-el="${kind === 'plot' ? 'plot' : 'chart'}" data-area-format="1"><defs>${defs.join('')}</defs><rect ${box()} fill="transparent"/>${body}</g>`;
}
