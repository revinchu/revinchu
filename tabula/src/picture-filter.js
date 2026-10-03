import { pictureEffects } from './picture.js';
const n = (v, fallback, lo, hi) => Math.max(lo, Math.min(hi, Number.isFinite(Number(v)) ? Number(v) : fallback));
const rgb = value => /^#[0-9a-f]{6}$/i.test(value ?? '') ? value : '#5b9bd5';
export const PICTURE_ARTISTIC = [{ id: 'none', name: '효과 없음' }, { id: 'blur', name: '흐리게' }, { id: 'pencil', name: '연필 스케치' }, { id: 'posterize', name: '포스터' }, { id: 'paint', name: '물감' }, { id: 'grain', name: '필름 입자' }];
export const PICTURE_RECOLOR = [{ id: 'none', name: '원래 색' }, { id: 'grayscale', name: '회색조' }, { id: 'sepia', name: '세피아' }, { id: 'washout', name: '희미하게' }, { id: 'blue', name: '파랑조' }, { id: 'green', name: '초록조' }, { id: 'orange', name: '주황조' }];
export const PICTURE_TINTS = { blue: '#266bd9', green: '#2e994d', orange: '#f26b1a' };
export function normalizePictureVisual(pic) {
  const correction = { brightness: n(pic.correction?.brightness, 0, -1, 1), contrast: n(pic.correction?.contrast, 0, -1, 1), sharpness: n(pic.correction?.sharpness, 0, 0, 1) };
  const color = { saturation: n(pic.color?.saturation, 1, 0, 2), temperature: n(pic.color?.temperature, 0, -1, 1), recolor: PICTURE_RECOLOR.some(v => v.id === pic.color?.recolor) ? pic.color.recolor : 'none' };
  if (Array.isArray(pic.color?.duotone) && pic.color.duotone.length === 2 && pic.color.duotone.every(c => /^#[0-9a-f]{6}$/i.test(c))) color.duotone = [...pic.color.duotone];
  const artistic = { type: PICTURE_ARTISTIC.some(v => v.id === pic.artistic?.type) ? pic.artistic.type : 'none', amount: n(pic.artistic?.amount, .5, 0, 1) };
  const glow = pic.glow ? { color: rgb(pic.glow.color), size: n(pic.glow.size, 8, 0, 100), opacity: n(pic.glow.opacity, .6, 0, 1) } : undefined;
  const softEdge = n(pic.softEdge, 0, 0, 100), reflection = pic.reflection ? { opacity: n(pic.reflection.opacity, .4, 0, 1), size: n(pic.reflection.size, .4, 0, 1), gap: n(pic.reflection.gap, 2, 0, 100) } : undefined;
  return { ...pictureEffects(pic), correction, color, artistic, glow, softEdge, reflection, borderDash: ['dash', 'dot', 'dashDot'].includes(pic.borderDash) ? pic.borderDash : 'solid' };
}
export function needsPictureBake(pic) {
  const { correction, color, artistic } = normalizePictureVisual(pic);
  return !!(correction.sharpness || color.saturation !== 1 || color.temperature || artistic.type !== 'none' || ['sepia', 'washout'].includes(color.recolor));
}
export function pictureVisual(pic, id = 'picture', scale = 1, filterUnits = 'userSpaceOnUse') {
  const effects = normalizePictureVisual(pic), { correction, color, artistic, glow, softEdge } = effects;
  const parts = [], k = n(scale, 1, .0001, 1000), fid = 'picfx-' + String(id).replace(/[^a-zA-Z0-9_-]/g, '').slice(0, 120);
  const transfer = (slope, intercept) => `<feComponentTransfer>${['R', 'G', 'B'].map(c => `<feFunc${c} type="linear" slope="${slope}" intercept="${intercept}"/>`).join('')}</feComponentTransfer>`;
  if (correction.brightness || correction.contrast) parts.push(transfer(1 + correction.contrast, correction.brightness - correction.contrast / 2));
  if (color.saturation !== 1) parts.push(`<feColorMatrix type="saturate" values="${color.saturation}"/>`);
  if (color.temperature) { const t = color.temperature * .35; parts.push(`<feColorMatrix type="matrix" values="${1 + t} 0 0 0 0 0 1 0 0 0 0 0 ${1 - t} 0 0 0 0 0 1 0"/>`); }
  if (color.duotone) {
    const channels = hex => [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255), lo = channels(color.duotone[0]), hi = channels(color.duotone[1]);
    parts.push(`<feColorMatrix type="matrix" values="${lo.map((v, i) => `${.2126 * (hi[i] - v)} ${.7152 * (hi[i] - v)} ${.0722 * (hi[i] - v)} 0 ${v}`).join(' ')} 0 0 0 1 0"/>`);
  } else if (color.recolor === 'grayscale') parts.push('<feColorMatrix type="saturate" values="0"/>');
  else if (color.recolor === 'sepia') parts.push('<feColorMatrix type="matrix" values=".393 .769 .189 0 0 .349 .686 .168 0 0 .272 .534 .131 0 0 0 0 0 1 0"/>');
  else if (color.recolor === 'washout') parts.push('<feColorMatrix type="saturate" values=".2"/>', transfer(.5, .5));
  else if (color.recolor !== 'none') { const hex = PICTURE_TINTS[color.recolor], tint = [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255); parts.push(`<feColorMatrix type="matrix" values="${tint.map(t => `${.2126 * t} ${.7152 * t} ${.0722 * t} 0 0`).join(' ')} 0 0 0 1 0"/>`); }
  if (correction.sharpness) { const a = correction.sharpness; parts.push(`<feConvolveMatrix order="3" kernelMatrix="0 ${-a} 0 ${-a} ${1 + 4 * a} ${-a} 0 ${-a} 0" divisor="1" edgeMode="duplicate" preserveAlpha="true"/>`); }
  if (artistic.type === 'blur') parts.push(`<feGaussianBlur stdDeviation="${(.2 + artistic.amount * 5) * k}"/>`);
  if (artistic.type === 'pencil') parts.push('<feColorMatrix type="saturate" values="0"/>', `<feConvolveMatrix order="3" kernelMatrix="-1 -1 -1 -1 8 -1 -1 -1 -1" divisor="1" bias="0" edgeMode="duplicate" preserveAlpha="true"/>`, transfer(-(1 + artistic.amount * 2), 1));
  if (artistic.type === 'posterize' || artistic.type === 'paint') {
    if (artistic.type === 'paint') parts.push(`<feGaussianBlur stdDeviation="${(1 + artistic.amount * 3) * k}"/>`);
    const steps = Math.max(2, Math.round(8 - artistic.amount * 6)), values = Array.from({ length: steps }, (_, i) => i / (steps - 1)).join(' ');
    parts.push(`<feComponentTransfer>${['R', 'G', 'B'].map(c => `<feFunc${c} type="discrete" tableValues="${values}"/>`).join('')}</feComponentTransfer>`);
  }
  if (artistic.type === 'grain') parts.push('<feColorMatrix type="matrix" values="1 0 0 0 0 0 1 0 0 0 0 0 1 0 0 0 0 0 1 0" result="picture-color"/>', '<feTurbulence type="fractalNoise" baseFrequency=".7" numOctaves="2" seed="7" result="noise"/>', `<feColorMatrix in="noise" type="matrix" values="0 0 0 0 .5 0 0 0 0 .5 0 0 0 0 .5 0 0 0 ${artistic.amount * .45} 0"/>`, '<feComposite operator="in" in2="SourceAlpha"/>', '<feBlend in2="picture-color" mode="multiply"/>');
  if (softEdge) parts.push('<feColorMatrix type="matrix" values="1 0 0 0 0 0 1 0 0 0 0 0 1 0 0 0 0 0 1 0" result="before-soft"/>', `<feMorphology in="SourceAlpha" operator="erode" radius="${softEdge * k / 2}"/>`, `<feGaussianBlur stdDeviation="${softEdge * k / 2}"/>`, '<feComposite in="before-soft" operator="in"/>');
  if (glow?.size && glow.opacity) parts.push(`<feDropShadow dx="0" dy="0" stdDeviation="${glow.size * k / 2}" flood-color="${glow.color}" flood-opacity="${glow.opacity}"/>`);
  if (effects.shadow) { const s = effects.shadow; parts.push(`<feDropShadow dx="${s.dx * k}" dy="${s.dy * k}" stdDeviation="${s.blur * k / 2}" flood-color="${s.color}" flood-opacity="${s.opacity}"/>`); }
  const shadowPad = effects.shadow ? Math.max(Math.abs(effects.shadow.dx), Math.abs(effects.shadow.dy)) + effects.shadow.blur * 2 : 0;
  const pad = Math.max(12, (softEdge + (glow?.size ?? 0) + (artistic.type === 'blur' ? 15 : 0)) * k * 3, shadowPad * k), w = Math.max(1, Number(pic.w) || 300) * k, h = Math.max(1, Number(pic.h) || 200) * k;
  // HTML에 참조한 userSpaceOnUse 영역을 WebKit은 문서 위치에 따라 자른다.
  // 격자/미리보기는 그림 자체의 경계 비율, SVG/내보내기는 SVG 좌표를 사용한다.
  // 흐림·그림자 등 효과의 길이는 양쪽 모두 그림 픽셀 단위를 유지한다.
  const relative = filterUnits === 'objectBoundingBox';
  const region = relative
    ? `filterUnits="objectBoundingBox" x="${-pad / w}" y="${-pad / h}" width="${1 + pad * 2 / w}" height="${1 + pad * 2 / h}"`
    : `filterUnits="userSpaceOnUse" x="${-pad}" y="${-pad}" width="${w + pad * 2}" height="${h + pad * 2}"`;
  return { defs: parts.length ? `<filter id="${fid}" ${region} primitiveUnits="userSpaceOnUse" color-interpolation-filters="sRGB">${parts.join('')}</filter>` : '', filter: parts.length ? `url(#${fid})` : '', effects };
}
