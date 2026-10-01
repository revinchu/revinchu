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
export function pictureTransform(picture) {
  const rot = finite(picture.rot), flipX = picture.flip ? -1 : 1, flipY = picture.flipV ? -1 : 1;
  return `${rot ? `rotate(${rot}deg) ` : ''}${flipX < 0 || flipY < 0 ? `scale(${flipX},${flipY})` : ''}`.trim();
}
export function resetPictureFormatting() {
  return { crop: undefined, rot: undefined, flip: undefined, flipV: undefined, border: undefined, borderW: undefined, radius: undefined, shadow: undefined, opacity: undefined };
}
