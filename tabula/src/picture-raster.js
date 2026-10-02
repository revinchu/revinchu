import { pictureVisual, needsPictureBake } from './picture-filter.js';
import { pictureCompressionSize } from './picture-pixels.js';
export async function loadPictureBitmap(src) {
  if (typeof src !== 'string' || !/^(data:image\/|blob:|https?:\/\/)/i.test(src)) throw new Error('이 그림은 픽셀 편집을 지원하지 않습니다.');
  const image = new Image(); if (!src.startsWith('data:') && !src.startsWith('blob:')) image.crossOrigin = 'anonymous';
  await new Promise((resolve, reject) => { const timeout = setTimeout(() => reject(new Error('그림을 읽는 시간이 초과되었습니다.')), 15000); image.onload = () => { clearTimeout(timeout); resolve(); }; image.onerror = () => { clearTimeout(timeout); reject(new Error('그림 픽셀을 읽지 못했습니다. 외부 그림은 파일로 내려받아 삽입하세요.')); }; image.src = src; });
  if (!image.naturalWidth || !image.naturalHeight || image.naturalWidth * image.naturalHeight > 32000000) throw new Error('픽셀 편집은 3,200만 픽셀 이하 그림에서 지원합니다.');
  return image;
}
export function pictureBitmapCanvas(image, width = image.naturalWidth, height = image.naturalHeight) {
  const canvas = document.createElement('canvas'); canvas.width = width; canvas.height = height;
  const ctx = canvas.getContext('2d', { willReadFrequently: true }); ctx.drawImage(image, 0, 0, width, height);
  try { ctx.getImageData(0, 0, 1, 1); } catch { throw new Error('외부 사이트가 그림 픽셀 읽기를 허용하지 않습니다. 그림 파일을 직접 삽입하세요.'); }
  return canvas;
}
export async function compressPicture(src, { maxWidth = 1920, maxHeight = 1080, quality = .85, type = 'image/jpeg' } = {}) {
  const image = await loadPictureBitmap(src), size = pictureCompressionSize(image.naturalWidth, image.naturalHeight, maxWidth, maxHeight), canvas = document.createElement('canvas'); canvas.width = size.w; canvas.height = size.h;
  const ctx = canvas.getContext('2d');
  const mime = ['image/jpeg', 'image/png', 'image/webp'].includes(type) ? type : 'image/jpeg';
  if (mime === 'image/jpeg') { ctx.fillStyle = '#ffffff'; ctx.fillRect(0, 0, size.w, size.h); }
  ctx.drawImage(image, 0, 0, size.w, size.h);
  try { return { src: canvas.toDataURL(mime, Math.max(.1, Math.min(1, Number(quality) || .85))), w: size.w, h: size.h, originalWidth: image.naturalWidth, originalHeight: image.naturalHeight }; }
  catch { throw new Error('이 외부 그림은 압축할 수 없습니다. 그림 파일을 직접 삽입하세요.'); }
}
/** 내보내기 사본용 픽셀 보정. 원본·자르기·도형 효과·배치는 바꾸지 않습니다. */
export async function bakePictureEffects(pic) {
  if (!needsPictureBake(pic)) return null;
  const image = await loadPictureBitmap(pic.src), canvas = pictureBitmapCanvas(image), w = canvas.width, h = canvas.height;
  if (w * h > 16000000) throw new Error('보정 그림 저장은 1,600만 픽셀 이하에서 지원합니다. 그림을 압축한 뒤 저장하세요.');
  const ratio = w / Math.max(1, pic.w), fx = pictureVisual({ ...pic, w: w / ratio, h: h / ratio, shadow: undefined, glow: undefined, softEdge: 0 }, 'baked-picture', ratio), safe = canvas.toDataURL('image/png');
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w}" height="${h}"><defs>${fx.defs}</defs><image href="${safe}" width="${w}" height="${h}" filter="${fx.filter}"/></svg>`;
  const painted = await loadPictureBitmap('data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg));
  return pictureBitmapCanvas(painted, w, h).toDataURL('image/png');
}
