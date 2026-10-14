// 큰 미디어 (동영상 · 큰 그림): data URL 대신 Blob + object URL (blob:…)
//   파일 조각(File.slice)을 그대로 가리키므로 GB 단위 파일도 메모리에 올리지 않음.
//   pres.media[id] 에는 'blob:…' 글이 들어가고, 실제 Blob 은 BLOBS 에서 찾음.
//   파일로 저장할 때 (pptx: Blob 그대로 zip · json/공유/공동 편집: data URL 로 바꿈) · SVG/PNG 로 그릴 때 (inlineBlobs) 사용.

export const BLOBS = new Map();

export const isBlobUrl = (u) => typeof u === 'string' && u.startsWith('blob:');

/** Blob → blob: URL (등록) */
export function blobUrl(blob, type) {
  const b = type && blob.type !== type ? blob.slice(0, blob.size, type) : blob;
  const url = URL.createObjectURL(b);
  BLOBS.set(url, b);
  return url;
}
export const blobOf = (url) => BLOBS.get(url) ?? null;

/** Blob → data URL */
export function blobToDataUrl(blob) {
  return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(blob); });
}

/** 그림 blob URL → 그릴 수 있는 data URL (maxSide 넘으면 줄임 · 다른 형식은 그대로) */
const inlineCache = new Map();
export async function inlineDataUrl(url, maxSide = 2560) {
  const key = `${url}|${maxSide}`;
  if (inlineCache.has(key)) return inlineCache.get(key);
  const b = blobOf(url);
  let out = null;
  if (b) {
    if (/^image\/(png|jpe?g|gif|webp|bmp)/.test(b.type) && b.size > 2e6) {
      try {
        const bmp = await createImageBitmap(b);
        const k = Math.min(1, maxSide / Math.max(bmp.width, bmp.height));
        const c = document.createElement('canvas');
        c.width = Math.max(1, Math.round(bmp.width * k));
        c.height = Math.max(1, Math.round(bmp.height * k));
        c.getContext('2d').drawImage(bmp, 0, 0, c.width, c.height);
        out = c.toDataURL(b.type === 'image/png' || b.type === 'image/gif' ? 'image/png' : 'image/jpeg', 0.9);
      } catch { out = null; }
    }
    if (!out && b.size < 200e6) out = await blobToDataUrl(b);
  }
  inlineCache.set(key, out);
  return out;
}

/** 글(HTML · SVG) 안의 blob: URL 을 data URL 로 (SVG 그림 · foreignObject 는 바깥 자원을 못 읽음) */
export async function inlineBlobs(text) {
  const urls = [...new Set(text.match(/blob:[^"')\s&]+/g) ?? [])];
  for (const u of urls) {
    const d = await inlineDataUrl(u);
    if (d) text = text.split(u).join(d);
  }
  return text;
}

/** 문서의 media 를 모두 data URL 로 (json 저장 · 공유 링크 · 공동 편집용) */
export async function mediaAsData(media, maxBytes = Infinity) {
  const out = {};
  for (const [id, u] of Object.entries(media ?? {})) out[id] = isBlobUrl(u) && blobOf(u) && blobOf(u).size <= maxBytes ? await blobToDataUrl(blobOf(u)) : u;
  return out;
}

/** IndexedDB 에서 읽은 media → Blob 은 blob URL 로 */
export function mediaFromStore(media) {
  const out = {};
  for (const [id, v] of Object.entries(media ?? {})) out[id] = typeof v === 'string' ? v : v instanceof Blob ? blobUrl(v) : v;
  return out;
}
