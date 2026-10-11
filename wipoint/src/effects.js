// 도형 · 그림 효과 (PowerPoint [도형 효과] / [그림 효과]) — DOM 없음
//   o.fx = { shadow: { inner?, dist, dir, blur, color, alpha } (pt · 도), glow: { size, color, alpha },
//            soft: pt, reflect: { size(0–1), dist, alpha, blur }, bevel: { prst, w, h }, rot3d: { x, y, persp } }
//   예전 o.shadow = true 는 기본 바깥쪽 그림자로 읽음.
import { resolveColor } from './themes.js';

const PX = 4 / 3; // pt → px
const f = (n) => Math.round(n * 100) / 100;
const rad = (d) => (d * Math.PI) / 180;

export const DEFAULT_SHADOW = { dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.4 };
/** 효과 정리 (예전 shadow 포함) */
export function fxOf(o) {
  if (o.fx) return o.fx;
  if (o.shadow) return { shadow: { ...DEFAULT_SHADOW } };
  return null;
}
const rgba = (theme, c, a) => {
  const hex = resolveColor(theme, c ?? '#000000', '#000000');
  const m = /^#?([0-9a-f]{6})/i.exec(hex) ?? [0, '000000'];
  const n = parseInt(m[1], 16);
  return `rgba(${(n >> 16) & 255},${(n >> 8) & 255},${n & 255},${f(a ?? 1)})`;
};

// ───────────── 미리 설정 (메뉴 갤러리) ─────────────
const OUT = (dir, dist = 3, blur = 4) => ({ dist, dir, blur, color: '#000000', alpha: 0.4 });
export const SHADOW_PRESETS = [
  ['바깥쪽', [['오프셋: 오른쪽 아래', OUT(45)], ['오프셋: 아래쪽', OUT(90)], ['오프셋: 왼쪽 아래', OUT(135)], ['오프셋: 오른쪽', OUT(0)], ['오프셋: 가운데', { dist: 0, dir: 0, blur: 5, color: '#000000', alpha: 0.4 }], ['오프셋: 왼쪽', OUT(180)], ['오프셋: 오른쪽 위', OUT(315)], ['오프셋: 위쪽', OUT(270)], ['오프셋: 왼쪽 위', OUT(225)]]],
  ['안쪽', [['안쪽: 왼쪽 위', { inner: true, dist: 3, dir: 45, blur: 4, color: '#000000', alpha: 0.5 }], ['안쪽: 위쪽', { inner: true, dist: 3, dir: 90, blur: 4, color: '#000000', alpha: 0.5 }], ['안쪽: 오른쪽 위', { inner: true, dist: 3, dir: 135, blur: 4, color: '#000000', alpha: 0.5 }], ['안쪽: 가운데', { inner: true, dist: 0, dir: 0, blur: 9, color: '#000000', alpha: 0.5 }]]],
  ['원근감', [['원근감: 아래쪽', { dist: 8, dir: 90, blur: 10, color: '#000000', alpha: 0.25, persp: true }], ['원근감: 오른쪽 아래', { dist: 10, dir: 60, blur: 12, color: '#000000', alpha: 0.25, persp: true }]]],
];
export const GLOW_SIZES = [5, 8, 11, 18];
export const GLOW_COLORS = ['@accent1', '@accent2', '@accent3', '@accent4', '@accent5', '@accent6'];
export const SOFT_PRESETS = [1, 2.5, 5, 10, 25, 50];
export const REFLECT_PRESETS = [
  ['근접 반사: 터치', { size: 0.35, dist: 0, alpha: 0.5, blur: 0.5 }], ['1/2 반사: 터치', { size: 0.55, dist: 0, alpha: 0.5, blur: 0.5 }], ['전체 반사: 터치', { size: 0.9, dist: 0, alpha: 0.5, blur: 0.5 }],
  ['근접 반사: 4pt 오프셋', { size: 0.35, dist: 4, alpha: 0.5, blur: 0.5 }], ['1/2 반사: 4pt 오프셋', { size: 0.55, dist: 4, alpha: 0.5, blur: 0.5 }], ['전체 반사: 4pt 오프셋', { size: 0.9, dist: 4, alpha: 0.5, blur: 0.5 }],
  ['근접 반사: 8pt 오프셋', { size: 0.35, dist: 8, alpha: 0.5, blur: 0.5 }], ['1/2 반사: 8pt 오프셋', { size: 0.55, dist: 8, alpha: 0.5, blur: 0.5 }], ['전체 반사: 8pt 오프셋', { size: 0.9, dist: 8, alpha: 0.5, blur: 0.5 }],
];
export const BEVEL_PRESETS = [['circle', '둥글게'], ['relaxedInset', '느슨하게'], ['cross', '십자'], ['coolSlant', '사선'], ['angle', '각지게'], ['softRound', '부드럽게 둥글게'], ['convex', '볼록하게'], ['slope', '경사'], ['divot', '오목하게'], ['riblet', '리블렛'], ['hardEdge', '딱딱한 가장자리'], ['artDeco', '아르데코']];
const BEVEL_DEPTH = { circle: 6, relaxedInset: 4, cross: 5, coolSlant: 6, angle: 6, softRound: 7, convex: 6, slope: 5, divot: 4, riblet: 3, hardEdge: 7, artDeco: 6 };
export const ROT3D_PRESETS = [
  ['평행', [['등각: 왼쪽 아래', { x: 0, y: 0, z: 0, iso: 'isometricLeftDown' }], ['등각: 오른쪽 위', { x: 0, y: 0, z: 0, iso: 'isometricRightUp' }], ['등각: 위쪽', { x: 0, y: 0, z: 0, iso: 'isometricTopUp' }], ['축 이동: 왼쪽', { x: 20, y: 10, z: 0 }], ['축 이동: 오른쪽', { x: -20, y: 10, z: 0 }]]],
  ['원근감', [['원근감: 앞쪽', { x: 0, y: 20, persp: true }], ['원근감: 왼쪽', { x: 30, y: 0, persp: true }], ['원근감: 오른쪽', { x: -30, y: 0, persp: true }], ['원근감: 아래', { x: 0, y: -25, persp: true }], ['원근감: 대비 왼쪽', { x: 40, y: 10, persp: true }], ['원근감: 대비 오른쪽', { x: -40, y: 10, persp: true }]]],
];
const ISO = { isometricLeftDown: [45, 35], isometricRightUp: [-45, -35], isometricTopUp: [0, 55] };
/** 미리 설정 1–12 (PowerPoint [도형 효과 › 미리 설정]) */
export const FX_PRESETS = [
  { shadow: { ...OUT(45) } },
  { shadow: { ...OUT(90, 4, 8) }, bevel: { prst: 'circle', w: 6, h: 6 } },
  { shadow: { ...OUT(90, 3, 6), alpha: 0.3 }, bevel: { prst: 'softRound', w: 6, h: 6 } },
  { bevel: { prst: 'angle', w: 6, h: 6 } },
  { shadow: { inner: true, dist: 0, dir: 0, blur: 9, color: '#000000', alpha: 0.5 } },
  { glow: { size: 8, color: '@accent1', alpha: 0.4 } },
  { reflect: { size: 0.55, dist: 0, alpha: 0.5, blur: 0.5 } },
  { soft: 5 },
  { bevel: { prst: 'convex', w: 8, h: 8 }, shadow: { ...OUT(45, 4, 6) } },
  { rot3d: { x: 0, y: 20, persp: true }, shadow: { dist: 10, dir: 90, blur: 10, color: '#000000', alpha: 0.25 } },
  { rot3d: { x: 0, y: 0, iso: 'isometricLeftDown' }, bevel: { prst: 'circle', w: 5, h: 5 } },
  { glow: { size: 11, color: '@accent2', alpha: 0.4 }, bevel: { prst: 'softRound', w: 5, h: 5 } },
];

// ───────────── 그리기 (CSS · SVG 필터) ─────────────
/**
 * 효과 → { filter: CSS filter (url + 그림자 · 네온), defs: SVG <filter> 글, wrap: 바깥 div 의 CSS (3차원 회전 · 반사), transform: [변환 함수] }
 * id: 이 개체 고유 (SVG 필터 id)
 */
export function fxStyle(theme, fx, id) {
  const out = { filter: '', defs: '', wrap: '', transform: [] };
  if (!fx) return out;
  const fns = [];
  const svg = [];
  const sh = fx.shadow;
  if (sh && !sh.inner) {
    const d = (sh.dist ?? 3) * PX;
    const dx = Math.cos(rad(sh.dir ?? 45)) * d;
    const dy = Math.sin(rad(sh.dir ?? 45)) * d;
    fns.push(`drop-shadow(${f(dx)}px ${f(dy)}px ${f((sh.blur ?? 4) * PX * 0.5)}px ${rgba(theme, sh.color, sh.alpha ?? 0.4)})`);
  }
  if (fx.glow) {
    const g = rgba(theme, fx.glow.color ?? '@accent1', fx.glow.alpha ?? 0.4);
    const r = (fx.glow.size ?? 8) * PX * 0.5;
    fns.push(`drop-shadow(0 0 ${f(r)}px ${g})`, `drop-shadow(0 0 ${f(r)}px ${g})`);
  }
  // SVG 필터 사슬: 부드러운 가장자리 → 안쪽 그림자 → 입체 효과
  let last = 'SourceGraphic';
  if (fx.soft) {
    const r = Math.max(0.5, fx.soft * PX * 0.5);
    svg.push(`<feMorphology in="SourceAlpha" operator="erode" radius="${f(r)}" result="se1"/><feGaussianBlur in="se1" stdDeviation="${f(r)}" result="se2"/><feComposite in="${last}" in2="se2" operator="in" result="soft"/>`);
    last = 'soft';
  }
  if (sh?.inner) {
    const d = (sh.dist ?? 3) * PX;
    svg.push(`<feComponentTransfer in="SourceAlpha" result="is1"><feFuncA type="table" tableValues="1 0"/></feComponentTransfer><feGaussianBlur in="is1" stdDeviation="${f((sh.blur ?? 4) * PX * 0.5)}" result="is2"/><feOffset in="is2" dx="${f(Math.cos(rad(sh.dir ?? 45)) * d)}" dy="${f(Math.sin(rad(sh.dir ?? 45)) * d)}" result="is3"/><feFlood flood-color="${resolveColor(theme, sh.color ?? '#000000', '#000000')}" flood-opacity="${f(sh.alpha ?? 0.5)}"/><feComposite in2="is3" operator="in" result="is4"/><feComposite in="is4" in2="SourceAlpha" operator="in" result="is5"/><feMerge result="inner"><feMergeNode in="${last}"/><feMergeNode in="is5"/></feMerge>`);
    last = 'inner';
  }
  if (fx.bevel) {
    const w = Math.max(1, (fx.bevel.w ?? BEVEL_DEPTH[fx.bevel.prst] ?? 6) * PX);
    const conc = fx.bevel.prst === 'divot' || fx.bevel.prst === 'relaxedInset';
    svg.push(`<feGaussianBlur in="SourceAlpha" stdDeviation="${f(w / 2)}" result="bv1"/><feSpecularLighting in="bv1" surfaceScale="${f(conc ? -w / 2 : w / 2)}" specularConstant="0.9" specularExponent="${fx.bevel.prst === 'hardEdge' || fx.bevel.prst === 'angle' ? 30 : 16}" lighting-color="#ffffff" result="bv2"><feDistantLight azimuth="225" elevation="40"/></feSpecularLighting><feComposite in="bv2" in2="SourceAlpha" operator="in" result="bv3"/><feComposite in="${last}" in2="bv3" operator="arithmetic" k1="0" k2="1" k3="0.55" k4="0" result="bevel"/>`);
    last = 'bevel';
  }
  if (svg.length) {
    const fid = `fx-${String(id).replace(/[^\w-]/g, '')}`;
    out.defs = `<filter id="${fid}" x="-25%" y="-25%" width="150%" height="150%" color-interpolation-filters="sRGB">${svg.join('')}</filter>`;
    fns.unshift(`url(#${fid})`);
  }
  out.filter = fns.join(' ');
  if (fx.reflect) {
    const r = fx.reflect;
    const end = Math.max(5, Math.min(100, (r.size ?? 0.5) * 100));
    out.wrap += `-webkit-box-reflect:below ${f((r.dist ?? 0) * PX)}px linear-gradient(to bottom, transparent ${f(100 - end)}%, rgba(255,255,255,${f(r.alpha ?? 0.5)}));`;
  }
  if (fx.rot3d) {
    const r = fx.rot3d;
    const [ry, rx] = r.iso ? ISO[r.iso] ?? [0, 0] : [r.x ?? 0, r.y ?? 0];
    out.transform.push(`${r.persp ? 'perspective(900px) ' : ''}rotateY(${f(ry)}deg) rotateX(${f(rx)}deg)${r.z ? ` rotateZ(${f(r.z)}deg)` : ''}`);
  }
  return out;
}

// ───────────── pptx (a:effectLst · a:scene3d · a:sp3d) ─────────────
const EMU = 12700;
const clr = (c, a) => {
  const alpha = a != null && a < 1 ? `<a:alpha val="${Math.round(a * 100000)}"/>` : '';
  if (typeof c === 'string' && c.startsWith('@')) { const [slot] = c.slice(1).split(':'); return `<a:schemeClr val="${slot}">${alpha}</a:schemeClr>`; }
  return `<a:srgbClr val="${String(c ?? '#000000').replace('#', '').slice(0, 6).toUpperCase()}">${alpha}</a:srgbClr>`;
};
/** 효과 → spPr 안에 넣을 XML (ln 다음) */
export function fxXml(fx) {
  if (!fx) return '';
  const lst = [];
  if (fx.glow) lst.push(`<a:glow rad="${Math.round((fx.glow.size ?? 8) * EMU)}">${clr(fx.glow.color ?? '@accent1', fx.glow.alpha ?? 0.4)}</a:glow>`);
  const sh = fx.shadow;
  if (sh?.inner) lst.push(`<a:innerShdw blurRad="${Math.round((sh.blur ?? 4) * EMU)}" dist="${Math.round((sh.dist ?? 3) * EMU)}" dir="${Math.round((((sh.dir ?? 45) % 360) + 360) % 360 * 60000)}">${clr(sh.color, sh.alpha ?? 0.5)}</a:innerShdw>`);
  else if (sh) lst.push(`<a:outerShdw blurRad="${Math.round((sh.blur ?? 4) * EMU)}" dist="${Math.round((sh.dist ?? 3) * EMU)}" dir="${Math.round((((sh.dir ?? 45) % 360) + 360) % 360 * 60000)}"${sh.persp ? ' sy="23000" kx="1200000"' : ''} algn="${sh.persp ? 'bl' : 'tl'}" rotWithShape="0">${clr(sh.color, sh.alpha ?? 0.4)}</a:outerShdw>`);
  if (fx.reflect) { const r = fx.reflect; lst.push(`<a:reflection blurRad="${Math.round((r.blur ?? 0.5) * EMU)}" stA="${Math.round((r.alpha ?? 0.5) * 100000)}" endPos="${Math.round((r.size ?? 0.5) * 100000)}" dist="${Math.round((r.dist ?? 0) * EMU)}" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/>`); }
  if (fx.soft) lst.push(`<a:softEdge rad="${Math.round(fx.soft * EMU)}"/>`);
  let x = lst.length ? `<a:effectLst>${lst.join('')}</a:effectLst>` : '';
  if (fx.rot3d) {
    const r = fx.rot3d;
    const cam = r.iso ? `<a:camera prst="${r.iso}"/>` : `<a:camera prst="${r.persp ? 'perspectiveFront' : 'orthographicFront'}"><a:rot lat="${Math.round(((((r.y ?? 0) % 360) + 360) % 360) * 60000)}" lon="${Math.round(((((r.x ?? 0) % 360) + 360) % 360) * 60000)}" rev="${Math.round(((((r.z ?? 0) % 360) + 360) % 360) * 60000)}"/></a:camera>`;
    x += `<a:scene3d>${cam}<a:lightRig rig="threePt" dir="t"/></a:scene3d>`;
  } else if (fx.bevel) x += '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="threePt" dir="t"/></a:scene3d>';
  if (fx.bevel) x += `<a:sp3d><a:bevelT w="${Math.round((fx.bevel.w ?? 6) * EMU)}" h="${Math.round((fx.bevel.h ?? 6) * EMU)}" prst="${fx.bevel.prst ?? 'circle'}"/></a:sp3d>`;
  return x;
}
/** spPr 요소 → 효과 (readColor(el) 는 색 읽기 함수) */
export function fxFromXml(spPr, { child, attr, readColor }) {
  if (!spPr) return null;
  const fx = {};
  const eff = child(spPr, 'effectLst');
  const n = (v, d = 0) => (v == null ? d : Number(v));
  const colorOf = (el) => {
    const c = (readColor(el) ?? '#000000').replace(/:a[\d.]+/, '');
    const a = el?.children?.[0] ? child(el.children[0], 'alpha') : null;
    return { color: c, alpha: a ? n(attr(a, 'val'), 100000) / 100000 : 1 };
  };
  if (eff) {
    const o = child(eff, 'outerShdw');
    const i = child(eff, 'innerShdw');
    const s = o ?? i;
    if (s) { const c = colorOf(s); fx.shadow = { ...(i && !o ? { inner: true } : {}), dist: n(attr(s, 'dist')) / EMU, dir: n(attr(s, 'dir')) / 60000, blur: n(attr(s, 'blurRad')) / EMU, color: c.color, alpha: c.alpha, ...(attr(s, 'sy') && n(attr(s, 'sy')) < 90000 ? { persp: true } : {}) }; }
    const g = child(eff, 'glow');
    if (g) { const c = colorOf(g); fx.glow = { size: n(attr(g, 'rad')) / EMU, color: c.color, alpha: c.alpha }; }
    const se = child(eff, 'softEdge');
    if (se) fx.soft = n(attr(se, 'rad')) / EMU;
    const rf = child(eff, 'reflection');
    if (rf) fx.reflect = { size: n(attr(rf, 'endPos'), 100000) / 100000, dist: n(attr(rf, 'dist')) / EMU, alpha: n(attr(rf, 'stA'), 50000) / 100000, blur: n(attr(rf, 'blurRad')) / EMU };
  }
  const sc = child(spPr, 'scene3d');
  const cam = child(sc, 'camera');
  if (cam) {
    const prst = attr(cam, 'prst') ?? '';
    const rot = child(cam, 'rot');
    if (ISO[prst]) fx.rot3d = { iso: prst };
    else if (rot && (n(attr(rot, 'lat')) || n(attr(rot, 'lon')) || n(attr(rot, 'rev')))) {
      const deg = (v) => { const d = n(v) / 60000; return d > 180 ? d - 360 : d; };
      fx.rot3d = { x: deg(attr(rot, 'lon')), y: deg(attr(rot, 'lat')), ...(deg(attr(rot, 'rev')) ? { z: deg(attr(rot, 'rev')) } : {}), ...(/^perspective/.test(prst) ? { persp: true } : {}) };
    } else if (/^perspective(?!Front)/.test(prst)) fx.rot3d = { x: 0, y: 15, persp: true };
  }
  const bt = child(child(spPr, 'sp3d'), 'bevelT');
  if (bt) fx.bevel = { prst: attr(bt, 'prst') ?? 'circle', w: n(attr(bt, 'w'), 76200) / EMU, h: n(attr(bt, 'h'), 76200) / EMU };
  return Object.keys(fx).length ? fx : null;
}
