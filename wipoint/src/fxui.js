// 도형 효과 · 그림 효과 메뉴 (PowerPoint: 미리 설정 · 그림자 · 반사 · 네온 · 부드러운 가장자리 · 입체 효과 · 3차원 회전) 와 서식 창 칸
import { S, change, selObjects, run } from './state.js';
import { el, closeMenus } from './ui.js';
import { objectHtml } from './render.js';
import { colorButton } from './colorpick.js';
import { fxOf, FX_PRESETS, SHADOW_PRESETS, GLOW_SIZES, GLOW_COLORS, SOFT_PRESETS, REFLECT_PRESETS, BEVEL_PRESETS, ROT3D_PRESETS, DEFAULT_SHADOW } from './effects.js';

const targets = () => selObjects().filter((o) => o.type === 'shape' || o.type === 'image');
/** fn(fx) 로 고른 개체들의 효과를 바꿈 (null 이면 지움) */
export function setFx(fn) {
  const objs = targets();
  if (!objs.length) return;
  change(() => {
    for (const o of objs) {
      const cur = JSON.parse(JSON.stringify(fxOf(o) ?? {}));
      const next = fn(cur) ?? cur;
      delete o.shadow;
      if (next && Object.keys(next).some((k) => next[k] != null)) { for (const k of Object.keys(next)) if (next[k] == null) delete next[k]; o.fx = next; } else delete o.fx;
    }
  });
}
let pid = 0;
function preview(fx, pic) {
  const o = pic ? { id: `pv${pid++}`, type: 'shape', shape: 'rect', x: 10, y: 8, w: 38, h: 30, fill: { type: 'gradient', angle: 45, stops: [[0, '@accent1:lm40:lo60'], [1, '@accent1']] }, line: null, fx }
    : { id: `pv${pid++}`, type: 'shape', shape: 'rect', x: 10, y: 8, w: 38, h: 30, fill: { type: 'solid', color: '@accent1' }, line: null, fx };
  return `<div class="fx-pv">${objectHtml(S.pres.theme, o, {}, {})}</div>`;
}
function gallery(items, pic, cols = 4) {
  return { node: el('div', { class: 'fx-gal', style: { gridTemplateColumns: `repeat(${cols}, auto)` } }, items.map(([label, fx, apply]) => el('button', { class: 'fx-item', title: label, type: 'button', onmousedown: (e) => e.preventDefault(), onclick: () => { closeMenus(); apply(); }, html: preview(fx, pic) }))) };
}
const head = (t) => ({ header: t });

export function effectsMenuItems(pic = false) {
  const none = (k) => ({ label: `${{ shadow: '그림자', reflect: '반사', glow: '네온', soft: '부드러운 가장자리', bevel: '입체 효과', rot3d: '3차원 회전' }[k]} 없음`, action: () => setFx((x) => { delete x[k]; return x; }) });
  const opts = (t) => ({ label: `${t} 옵션...`, action: () => { S.paneFocus = 'effects'; run('formatPane', 'shape'); } });
  return [
    { label: '미리 설정(P)', submenu: [{ label: '효과 없음', action: () => setFx(() => null) }, gallery(FX_PRESETS.map((p, i) => [`미리 설정 ${i + 1}`, p, () => setFx(() => JSON.parse(JSON.stringify(p)))]), pic)] },
    { label: '그림자(S)', submenu: [none('shadow'), ...SHADOW_PRESETS.flatMap(([g, list]) => [head(g), gallery(list.map(([l, sh]) => [l, { shadow: sh }, () => setFx((x) => { x.shadow = { ...sh }; return x; })]), pic, g === '바깥쪽' ? 3 : 4)]), opts('그림자')] },
    { label: '반사(R)', submenu: [none('reflect'), head('반사 변형'), gallery(REFLECT_PRESETS.map(([l, r]) => [l, { reflect: r }, () => setFx((x) => { x.reflect = { ...r }; return x; })]), pic, 3), opts('반사')] },
    { label: '네온(G)', submenu: [none('glow'), head('네온 변형'), gallery(GLOW_SIZES.flatMap((sz) => GLOW_COLORS.map((c, i) => [`네온: ${sz}pt, 강조 ${i + 1}`, { glow: { size: sz, color: c, alpha: 0.4 } }, () => setFx((x) => { x.glow = { size: sz, color: c, alpha: 0.4 }; return x; })])), pic, 6), opts('네온')] },
    { label: '부드러운 가장자리(E)', submenu: [none('soft'), ...SOFT_PRESETS.map((v) => ({ label: `${v}pt`, action: () => setFx((x) => { x.soft = v; return x; }) })), opts('부드러운 가장자리')] },
    { label: '입체 효과(B)', submenu: [none('bevel'), head('입체 효과'), gallery(BEVEL_PRESETS.map(([k, l]) => [l, { bevel: { prst: k, w: 6, h: 6 } }, () => setFx((x) => { x.bevel = { prst: k, w: 6, h: 6 }; return x; })]), pic, 4), opts('3차원')] },
    { label: '3차원 회전(D)', submenu: [none('rot3d'), ...ROT3D_PRESETS.flatMap(([g, list]) => [head(g), gallery(list.map(([l, r]) => [l, { rot3d: r }, () => setFx((x) => { x.rot3d = { ...r }; return x; })]), pic, 3)]), opts('3차원 회전')] },
    { sep: true },
    { label: pic ? '그림 서식...' : '도형 서식...', action: () => { S.paneFocus = 'effects'; run('formatPane', 'shape'); } },
  ];
}

/** 서식 창 [효과] 칸 */
export function effectsPaneSections(o, { sec, row, num, sel, check }) {
  const fx = fxOf(o) ?? {};
  const sh = fx.shadow;
  const glow = fx.glow;
  const refl = fx.reflect;
  const bev = fx.bevel;
  const r3 = fx.rot3d;
  const upd = (k, patch) => setFx((x) => { x[k] = patch == null ? undefined : { ...(x[k] ?? {}), ...patch }; return x; });
  const out = [];
  out.push(sec('그림자',
    sel(!sh ? 'none' : sh.inner ? 'inner' : sh.persp ? 'persp' : 'outer', [['none', '없음'], ['outer', '바깥쪽'], ['inner', '안쪽'], ['persp', '원근감']], (v) => setFx((x) => { x.shadow = v === 'none' ? undefined : { ...DEFAULT_SHADOW, ...(x.shadow ?? {}), inner: v === 'inner' || undefined, persp: v === 'persp' || undefined }; return x; })),
    sh ? [
      row('색', colorButton(sh.color ?? '#000000', (c) => upd('shadow', { color: c ?? '#000000' }))),
      row('투명도', num(Math.round((1 - (sh.alpha ?? 0.4)) * 100), (v) => upd('shadow', { alpha: Math.max(0, Math.min(100, 100 - v)) / 100 }), { min: 0, max: 100, suffix: '%' })),
      row('흐리게', num(sh.blur ?? 4, (v) => upd('shadow', { blur: Math.max(0, v) }), { min: 0, step: 0.5, suffix: 'pt' })),
      row('각도', num(sh.dir ?? 45, (v) => upd('shadow', { dir: ((v % 360) + 360) % 360 }), { min: 0, max: 359, suffix: '°' })),
      row('간격', num(sh.dist ?? 3, (v) => upd('shadow', { dist: Math.max(0, v) }), { min: 0, step: 0.5, suffix: 'pt' })),
    ] : null));
  out.push(sec('반사', check(!!refl, (v) => upd('reflect', v ? { size: 0.5, dist: 0, alpha: 0.5, blur: 0.5 } : null), '반사'),
    refl ? [row('투명도', num(Math.round((1 - (refl.alpha ?? 0.5)) * 100), (v) => upd('reflect', { alpha: (100 - v) / 100 }), { min: 0, max: 100, suffix: '%' })), row('크기', num(Math.round((refl.size ?? 0.5) * 100), (v) => upd('reflect', { size: Math.max(5, Math.min(100, v)) / 100 }), { min: 5, max: 100, suffix: '%' })), row('간격', num(refl.dist ?? 0, (v) => upd('reflect', { dist: Math.max(0, v) }), { min: 0, suffix: 'pt' }))] : null));
  out.push(sec('네온', check(!!glow, (v) => upd('glow', v ? { size: 8, color: '@accent1', alpha: 0.4 } : null), '네온'),
    glow ? [row('색', colorButton(glow.color ?? '@accent1', (c) => upd('glow', { color: c ?? '@accent1' }))), row('크기', num(glow.size ?? 8, (v) => upd('glow', { size: Math.max(0, v) }), { min: 0, suffix: 'pt' })), row('투명도', num(Math.round((1 - (glow.alpha ?? 0.4)) * 100), (v) => upd('glow', { alpha: (100 - v) / 100 }), { min: 0, max: 100, suffix: '%' }))] : null));
  out.push(sec('부드러운 가장자리', row('크기', num(fx.soft ?? 0, (v) => setFx((x) => { x.soft = v > 0 ? v : undefined; return x; }), { min: 0, step: 0.5, suffix: 'pt' }))));
  out.push(sec('3차원 서식',
    row('위쪽 입체', sel(bev?.prst ?? '', [['', '없음'], ...BEVEL_PRESETS], (v) => upd('bevel', v ? { prst: v, w: bev?.w ?? 6, h: bev?.h ?? 6 } : null))),
    bev ? [row('너비', num(bev.w ?? 6, (v) => upd('bevel', { w: Math.max(0, v) }), { min: 0, suffix: 'pt' })), row('높이', num(bev.h ?? 6, (v) => upd('bevel', { h: Math.max(0, v) }), { min: 0, suffix: 'pt' }))] : null));
  out.push(sec('3차원 회전',
    row('X 회전', num(r3?.x ?? 0, (v) => upd('rot3d', { x: v, iso: undefined }), { suffix: '°' })),
    row('Y 회전', num(r3?.y ?? 0, (v) => upd('rot3d', { y: v, iso: undefined }), { suffix: '°' })),
    row('Z 회전', num(r3?.z ?? 0, (v) => upd('rot3d', { z: v, iso: undefined }), { suffix: '°' })),
    check(!!r3?.persp, (v) => upd('rot3d', { persp: v || undefined }), '원근감'),
    el('div', { class: 'pane-btns' }, el('button', { class: 'btn small', onclick: () => setFx((x) => { delete x.rot3d; return x; }) }, '원래대로'))));
  return out;
}
