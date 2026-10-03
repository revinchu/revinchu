// 같은 그림 모델을 격자(HTML), 그룹/내보내기(SVG), 인쇄에서 표현한다.
import { pictureCropStyle } from './picture.js';
import { pictureVisual } from './picture-filter.js';
import { esc } from './xml.js';

const finite = (v, fallback = 0) => Number.isFinite(Number(v)) ? Number(v) : fallback;
const idOf = value => 'pic' + String(value).replace(/[^A-Za-z0-9_-]/g, '_');
const dashOf = (kind, width) => ({dash:[4,3],dot:[1,2],dashDot:[4,2,1,2]})[kind]?.map(n => n * width).join(' ');

/** Inner markup only: caller owns position, rotation and flips. */
export function pictureMarkup(picture, id = picture.id ?? 'picture') {
  const visual = pictureVisual(picture, idOf(id), 1, 'objectBoundingBox'), e = visual.effects, css = pictureCropStyle(picture.crop);
  const image = `<img src="${esc(picture.src ?? '')}" alt="${esc(picture.alt ?? picture.name ?? '')}" draggable="false" referrerpolicy="no-referrer" style="position:absolute;max-width:none;left:${css.left};top:${css.top};width:${css.width};height:${css.height}">`;
  const clipped = `<div style="position:absolute;inset:0;overflow:hidden;border-radius:${e.radius}px">${image}</div>`;
  const filter = visual.filter ? `filter:${visual.filter};` : '';
  const content = `<div class="picture-pixels" style="position:absolute;inset:0;${filter}">${clipped}</div>`;
  const defs = visual.defs ? `<svg aria-hidden="true" width="0" height="0" style="position:absolute;pointer-events:none"><defs>${visual.defs}</defs></svg>` : '';
  const borderW = Math.max(0, finite(picture.borderW, 2));
  const border = /^#[0-9a-f]{6}$/i.test(picture.border ?? '') ? `<svg class="picture-border" aria-hidden="true" width="100%" height="100%" viewBox="0 0 ${picture.w} ${picture.h}" preserveAspectRatio="none" style="position:absolute;inset:0;overflow:visible;pointer-events:none"><rect x="${borderW / 2}" y="${borderW / 2}" width="${Math.max(0,picture.w-borderW)}" height="${Math.max(0,picture.h-borderW)}" rx="${e.radius}" fill="none" stroke="${picture.border}" stroke-width="${borderW}"${dashOf(e.borderDash,borderW) ? ` stroke-dasharray="${dashOf(e.borderDash,borderW)}"` : ''}/></svg>` : '';
  const ref = e.reflection;
  const reflection = ref && ref.size > 0 ? `<div class="picture-reflection" aria-hidden="true" style="position:absolute;left:0;top:calc(100% + ${ref.gap}px);width:100%;height:${ref.size*100}%;overflow:hidden;opacity:${ref.opacity};mask-image:linear-gradient(to bottom,#000,transparent);pointer-events:none"><div style="position:absolute;left:0;top:0;width:100%;height:${100/ref.size}%;transform:scaleY(-1)">${content}${border}</div></div>` : '';
  return `${defs}<div class="picture-visual" style="position:absolute;inset:0;opacity:${e.opacity}">${content}${border}${reflection}</div>`;
}

/** Complete SVG, with crop/flip/effects; parent owns rotation and sheet position. */
export function pictureSvg(picture, id = picture.id ?? 'picture') {
  const uid = idOf(id), visual = pictureVisual(picture, uid), e = visual.effects;
  const w = Math.max(1,finite(picture.w,1)), h = Math.max(1,finite(picture.h,1));
  const crop = picture.crop ?? {}, l = finite(crop.l), t = finite(crop.t);
  const cw = Math.max(.01,1-l-finite(crop.r)), ch = Math.max(.01,1-t-finite(crop.b));
  const filter = visual.filter ? ` filter="${visual.filter}"` : '';
  const image = `<g${filter}><g clip-path="url(#${uid}clip)"><image href="${esc(picture.src ?? '')}" x="${-l*w/cw}" y="${-t*h/ch}" width="${w/cw}" height="${h/ch}" preserveAspectRatio="none"/></g></g>`;
  const borderW = Math.max(0,finite(picture.borderW,2));
  const border = /^#[0-9a-f]{6}$/i.test(picture.border ?? '') ? `<rect x="${borderW/2}" y="${borderW/2}" width="${Math.max(0,w-borderW)}" height="${Math.max(0,h-borderW)}" rx="${e.radius}" fill="none" stroke="${picture.border}" stroke-width="${borderW}"${dashOf(e.borderDash,borderW) ? ` stroke-dasharray="${dashOf(e.borderDash,borderW)}"` : ''}/>` : '';
  const ref = e.reflection, refH = ref?.size * h;
  const reflection = ref && refH > 0 ? `<g mask="url(#${uid}fade)"><g transform="translate(0,${2*h+ref.gap}) scale(1,-1)">${image}${border}</g></g>` : '';
  const refDefs = ref && refH > 0 ? `<linearGradient id="${uid}gradient" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="white" stop-opacity="${ref.opacity}"/><stop offset="1" stop-color="white" stop-opacity="0"/></linearGradient><mask id="${uid}fade" maskUnits="userSpaceOnUse" x="0" y="${h+ref.gap}" width="${w}" height="${refH}"><rect x="0" y="${h+ref.gap}" width="${w}" height="${refH}" fill="url(#${uid}gradient)"/></mask>` : '';
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}" viewBox="0 0 ${w} ${h}" preserveAspectRatio="none" style="overflow:visible"><defs>${visual.defs}<clipPath id="${uid}clip"><rect width="${w}" height="${h}" rx="${e.radius}"/></clipPath>${refDefs}</defs><g opacity="${e.opacity}" transform="translate(${picture.flip?w:0},${picture.flipV?h:0}) scale(${picture.flip?-1:1},${picture.flipV?-1:1})">${image}${border}${reflection}</g></svg>`;
}
