// 그림 편집의 수치 규칙. 원본 객체/자르기 데이터는 변경하지 않습니다.
const finite = (n, fallback = 0) => Number.isFinite(Number(n)) ? Number(n) : fallback;
const bound = (n, lo, hi) => Math.max(lo, Math.min(hi, n));
export function resizePicture(picture, axis, value, locked = picture.lockAspect !== false) {
  const w = Math.max(4, finite(picture.w, 4)), h = Math.max(4, finite(picture.h, 4));
  const n = bound(finite(value, axis === 'h' ? h : w), 4, 20000);
  if (!locked) return { w: axis === 'w' ? n : w, h: axis === 'h' ? n : h };
  const k = bound(n / (axis === 'h' ? h : w), Math.max(4 / w, 4 / h), Math.min(20000 / w, 20000 / h));
  return { w: Math.round(w * k * 1000) / 1000, h: Math.round(h * k * 1000) / 1000 };
}
export function setPictureCrop(crop, side, percent) {
  const next = { ...crop }, opposite = { l: 'r', r: 'l', t: 'b', b: 't' }[side];
  if (!opposite) return next;
  next[side] = bound(finite(percent) / 100, -10, Math.max(-10, .99 - finite(next[opposite])));
  return next;
}
export function pictureCropStyle(crop) {
  const c = crop ?? {}, l = finite(c.l), t = finite(c.t), w = Math.max(.01, 1 - l - finite(c.r)), h = Math.max(.01, 1 - t - finite(c.b));
  const pct = v => `${Math.round(v * 100000) / 1000}%`;
  return { left: pct(-l / w), top: pct(-t / h), width: pct(1 / w), height: pct(1 / h) };
}
export function dragPictureCrop(crop, edge, dx, dy) {
  const original = crop ?? {}, width = Math.max(.01, 1 - finite(original.l) - finite(original.r)), height = Math.max(.01, 1 - finite(original.t) - finite(original.b));
  let next = { ...original };
  for (const side of String(edge)) {
    const delta = side === 'l' ? finite(dx) * width : side === 'r' ? -finite(dx) * width : side === 't' ? finite(dy) * height : side === 'b' ? -finite(dy) * height : 0;
    next = setPictureCrop(next, side, (finite(original[side]) + delta) * 100);
  }
  return next;
}
// 키보드는 현재 남은 폭이 아니라 원본 그림의 비율만큼 이동합니다.
export function nudgePictureCrop(crop, edge, dx, dy) {
  const original = crop ?? {};
  let next = { ...original };
  for (const side of String(edge)) {
    const delta = side === 'l' ? finite(dx) : side === 'r' ? -finite(dx) : side === 't' ? finite(dy) : side === 'b' ? -finite(dy) : 0;
    next = setPictureCrop(next, side, (finite(original[side]) + delta) * 100);
  }
  return next;
}
export function pictureTransform(picture) {
  const rot = finite(picture.rot), flipX = picture.flip ? -1 : 1, flipY = picture.flipV ? -1 : 1;
  return `${rot ? `rotate(${rot}deg) ` : ''}${flipX < 0 || flipY < 0 ? `scale(${flipX},${flipY})` : ''}`.trim();
}
export function resetPictureFormatting() {
  return { crop: undefined, rot: undefined, flip: undefined, flipV: undefined, border: undefined, borderW: undefined, borderDash: undefined, radius: undefined, shadow: undefined, opacity: undefined, correction: undefined, color: undefined, artistic: undefined, glow: undefined, softEdge: undefined, reflection: undefined };
}

export const PICTURE_CROP_PRESETS = [
  { id: 'original', name: '원본 비율', ratio: 0 }, { id: 'square', name: '정사각형 1:1', ratio: 1 },
  { id: 'landscape43', name: '가로 4:3', ratio: 4 / 3 }, { id: 'landscape169', name: '가로 16:9', ratio: 16 / 9 },
  { id: 'landscape32', name: '가로 3:2', ratio: 1.5 }, { id: 'portrait34', name: '세로 3:4', ratio: .75 },
  { id: 'portrait916', name: '세로 9:16', ratio: 9 / 16 }, { id: 'portrait23', name: '세로 2:3', ratio: 2 / 3 },
];
export function pictureCropPreset(picture, ratio, natural = {}) {
  const sw = Math.max(1, finite(natural.w, picture.w)), sh = Math.max(1, finite(natural.h, picture.h)), target = finite(ratio, sw / sh) || sw / sh;
  const crop = { l: 0, r: 0, t: 0, b: 0 };
  if (sw / sh > target) crop.l = crop.r = (1 - target * sh / sw) / 2;
  else crop.t = crop.b = (1 - sw / (target * sh)) / 2;
  return { crop, ...resizePicture({ w: picture.w, h: picture.w / target }, 'w', picture.w, true) };
}
export const pictureCm = px => Math.round(finite(px) * 2.54 / 96 * 1000) / 1000;
export const picturePx = cm => finite(cm) * 96 / 2.54;

export function pictureSourcePatch(picture, src, natural = {}) {
  return { src, png: undefined, emf: undefined,
    originalSrc: picture.originalSrc ?? picture.src,
    originalWidth: picture.originalWidth ?? natural.w ?? picture.w, originalHeight: picture.originalHeight ?? natural.h ?? picture.h,
    originalEmf: picture.originalEmf ?? picture.emf, originalPng: picture.originalPng ?? picture.png };
}
export function resetPictureSource(picture, withSize = false) {
  const patch = picture.originalSrc ? { src: picture.originalSrc, emf: picture.originalEmf, png: picture.originalPng,
    originalSrc: undefined, originalWidth: undefined, originalHeight: undefined, originalEmf: undefined, originalPng: undefined } : {};
  if (withSize && picture.originalWidth > 0 && picture.originalHeight > 0) Object.assign(patch, resizePicture({ w: picture.originalWidth, h: picture.originalHeight }, 'w', picture.originalWidth));
  return patch;
}

const styleShadow = (blur, dy, opacity = .35) => ({ dx: 0, dy, blur, opacity, color: '#000000' });
export const PICTURE_STYLES = [
  { id: 'plain', name: '기본 그림', patch: {} },
  { id: 'white-frame', name: '흰색 단순 프레임', patch: { border: '#ffffff', borderW: 6, shadow: styleShadow(5, 2) } },
  { id: 'black-frame', name: '검정 단순 프레임', patch: { border: '#222222', borderW: 5 } },
  { id: 'gray-frame', name: '회색 얇은 프레임', patch: { border: '#808080', borderW: 2 } },
  { id: 'round-white', name: '흰색 둥근 프레임', patch: { border: '#ffffff', borderW: 6, radius: 18, shadow: styleShadow(8, 3) } },
  { id: 'round-black', name: '검정 둥근 프레임', patch: { border: '#222222', borderW: 4, radius: 18 } },
  { id: 'soft-shadow', name: '부드러운 그림자', patch: { shadow: styleShadow(12, 5) } },
  { id: 'deep-shadow', name: '진한 그림자', patch: { shadow: styleShadow(18, 10, .6) } },
  { id: 'offset-shadow', name: '오른쪽 아래 그림자', patch: { shadow: { dx: 8, dy: 8, blur: 5, color: '#000000', opacity: .4 } } },
  { id: 'floating', name: '떠 있는 그림', patch: { border: '#ffffff', borderW: 3, shadow: styleShadow(18, 14, .25) } },
  { id: 'blue-frame', name: '파랑 강조 프레임', patch: { border: '#4472c4', borderW: 5 } },
  { id: 'green-frame', name: '초록 강조 프레임', patch: { border: '#548235', borderW: 5 } },
  { id: 'orange-frame', name: '주황 강조 프레임', patch: { border: '#ed7d31', borderW: 5 } },
  { id: 'gold-frame', name: '금색 강조 프레임', patch: { border: '#bf9000', borderW: 5 } },
  { id: 'dash-frame', name: '파랑 파선 프레임', patch: { border: '#4472c4', borderW: 3, borderDash: 'dash' } },
  { id: 'dot-frame', name: '회색 점선 프레임', patch: { border: '#666666', borderW: 3, borderDash: 'dot' } },
  { id: 'blue-glow', name: '파랑 네온', patch: { glow: { color: '#5b9bd5', size: 10, opacity: .7 } } },
  { id: 'gold-glow', name: '금색 네온', patch: { glow: { color: '#ffc000', size: 10, opacity: .7 } } },
  { id: 'soft-edge', name: '부드러운 가장자리', patch: { softEdge: 10 } },
  { id: 'soft-round', name: '부드러운 둥근 모서리', patch: { radius: 20, softEdge: 5 } },
  { id: 'reflection-small', name: '가까운 작은 반사', patch: { reflection: { opacity: .35, size: .3, gap: 2 } } },
  { id: 'reflection-full', name: '넓은 반사', patch: { reflection: { opacity: .45, size: .65, gap: 5 } } },
  { id: 'round-reflection', name: '둥근 그림과 반사', patch: { radius: 18, reflection: { opacity: .3, size: .4, gap: 4 } } },
  { id: 'glow-shadow', name: '네온과 그림자', patch: { glow: { color: '#70ad47', size: 6, opacity: .55 }, shadow: styleShadow(8, 4) } },
];
export function pictureStylePatch(id, picture) {
  const style = PICTURE_STYLES.find(s => s.id === id);
  if (!style) throw new Error('선택한 그림 스타일을 찾을 수 없습니다.');
  const patch = { border: undefined, borderW: undefined, borderDash: undefined, radius: undefined, shadow: undefined, glow: undefined, softEdge: undefined, reflection: undefined, ...structuredClone(style.patch) };
  if (patch.radius && picture) patch.radius = Math.min(patch.radius, picture.w / 2, picture.h / 2);
  return patch;
}

// DrawingML과 화면이 같은 단위(px, 0..1 투명도)를 사용합니다.
export function pictureEffects(picture) {
  const opacity = bound(finite(picture.opacity, 1), 0, 1);
  const radius = bound(finite(picture.radius), 0, Math.min(Math.max(0, finite(picture.w)), Math.max(0, finite(picture.h))) / 2);
  const raw = picture.shadow;
  const s = raw && typeof raw === 'object' ? raw : {};
  const shadow = raw ? { dx: bound(finite(s.dx, 3), -1000, 1000), dy: bound(finite(s.dy, 3), -1000, 1000), blur: bound(finite(s.blur, 8), 0, 1000), color: /^#[0-9a-f]{6}$/i.test(s.color ?? '') ? s.color : '#000000', opacity: bound(finite(s.opacity, .4), 0, 1) } : undefined;
  return { opacity, radius, shadow };
}
export function pictureShadowStyle(picture, scale = 1) {
  const s = pictureEffects(picture).shadow;
  if (!s) return '';
  const c = s.color.slice(1), rgb = [0, 2, 4].map(i => parseInt(c.slice(i, i + 2), 16)).join(',');
  return `${s.dx * scale}px ${s.dy * scale}px ${s.blur * scale}px rgba(${rgb},${s.opacity})`;
}
