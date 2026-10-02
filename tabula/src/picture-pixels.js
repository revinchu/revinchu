// 사진 편집은 새 RGBA 버퍼에 적용합니다. 사용자 원본과 입력 배열은 변경하지 않습니다.
export const PICTURE_PIXEL_LIMIT = 4000000;
export function removePictureBackground(image, { color = [255, 255, 255], tolerance = 18, seed = null, marks = [] } = {}) {
  const { width, height, data } = image, count = width * height;
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || count > PICTURE_PIXEL_LIMIT || data.length !== count * 4) throw new Error('배경 제거는 400만 픽셀 이하 그림에서 지원합니다. 먼저 그림을 압축하세요.');
  const output = new Uint8ClampedArray(data), seen = new Uint8Array(count), queue = new Uint32Array(count), threshold = Math.max(0, Math.min(100, Number(tolerance) || 0)) * 2.55;
  let head = 0, tail = 0, removed = 0;
  const visit = n => {
    if (n < 0 || n >= count || seen[n]) return; seen[n] = 1;
    const i = n * 4, dr = data[i] - color[0], dg = data[i + 1] - color[1], db = data[i + 2] - color[2];
    if (data[i + 3] !== 0 && dr * dr + dg * dg + db * db > threshold * threshold * 3) return;
    queue[tail++] = n;
  };
  if (seed && Number.isFinite(seed.x) && Number.isFinite(seed.y)) visit(Math.max(0, Math.min(height - 1, Math.floor(seed.y))) * width + Math.max(0, Math.min(width - 1, Math.floor(seed.x))));
  else { for (let x = 0; x < width; x++) { visit(x); visit((height - 1) * width + x); } for (let y = 1; y < height - 1; y++) { visit(y * width); visit(y * width + width - 1); } }
  while (head < tail) { const n = queue[head++], x = n % width, i = n * 4; if (output[i + 3]) removed++; output[i + 3] = 0; if (x) visit(n - 1); if (x + 1 < width) visit(n + 1); if (n >= width) visit(n - width); if (n + width < count) visit(n + width); }
  if (!Array.isArray(marks) || marks.length > 2000) throw new Error('유지·제거 표시는 2,000개 이내로 지정하세요.');
  for (const mark of marks) {
    if (!Number.isFinite(mark.x) || !Number.isFinite(mark.y)) continue;
    const radius = Math.max(1, Math.min(200, Number(mark.radius) || 8)), x1 = Math.max(0, Math.floor(mark.x - radius)), x2 = Math.min(width - 1, Math.ceil(mark.x + radius)), y1 = Math.max(0, Math.floor(mark.y - radius)), y2 = Math.min(height - 1, Math.ceil(mark.y + radius));
    for (let y = y1; y <= y2; y++) for (let x = x1; x <= x2; x++) if ((x - mark.x) ** 2 + (y - mark.y) ** 2 <= radius ** 2) { const i = (y * width + x) * 4; output[i + 3] = mark.mode === 'keep' ? data[i + 3] : 0; }
  }
  return { width, height, data: output, removed };
}
export function pictureDataBytes(src) {
  if (typeof src !== 'string' || !src.startsWith('data:')) return null;
  const comma = src.indexOf(','); if (comma < 0) return null;
  if (/;base64$/i.test(src.slice(0, comma))) return Math.max(0, Math.floor((src.length - comma - 1) * .75) - (src.endsWith('==') ? 2 : src.endsWith('=') ? 1 : 0));
  try { return new TextEncoder().encode(decodeURIComponent(src.slice(comma + 1))).length; } catch { return null; }
}
export function pictureCompressionSize(width, height, maxWidth = 1920, maxHeight = 1080) {
  const k = Math.min(1, Math.max(1, maxWidth) / width, Math.max(1, maxHeight) / height);
  return { w: Math.max(1, Math.round(width * k)), h: Math.max(1, Math.round(height * k)) };
}
