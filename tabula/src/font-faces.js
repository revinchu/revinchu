// Shared, inert parser for catalog-owned CSS. It never executes CSS or follows imports.
export function parseFontFaces(css, baseUrl) {
  const faces = [];
  for (const [block] of String(css).replace(/\/\*[\s\S]*?\*\//g, '').matchAll(/@font-face\s*\{[^}]*\}/gi)) {
    const field = name => new RegExp('(?:[;{])\\s*' + name + '\\s*:\\s*([^;}]+)', 'i').exec(block)?.[1]?.trim();
    const urls = [...(field('src') ?? '').matchAll(/url\(\s*(?:"([^"]+)"|'([^']+)'|([^\s)]+))\s*\)\s*(?:format\(['"]?([^)'"\s]+)['"]?\))?/gi)];
    const pick = urls.find(m => /woff2/i.test(m[4] ?? '') || /\.woff2(?:[?#]|$)/i.test(m[1] ?? m[2] ?? m[3])) ?? urls[0];
    if (!pick) continue;
    let url; try { url = new URL(pick[1] ?? pick[2] ?? pick[3], baseUrl).href; } catch { continue; }
    faces.push({ url, weight: field('font-weight') ?? '400', style: field('font-style') ?? 'normal', unicodeRange: field('unicode-range') });
  }
  return faces;
}
export function fontFaceWeightDistance(face, weight) {
  const values = String(face.weight ?? '400').trim().split(/\s+/).map(n => n === 'normal' ? 400 : n === 'bold' ? 700 : Number(n));
  const start = Number.isFinite(values[0]) ? values[0] : 400, end = Number.isFinite(values[1]) ? values[1] : start;
  return weight < start ? start - weight : weight > end ? weight - end : 0;
}
export function fontIdentityFaces(entry, faces) {
  if (!Number.isFinite(entry.fixedWeight) || !faces.length) return faces;
  let distance = Infinity;
  for (const face of faces) distance = Math.min(distance, fontFaceWeightDistance(face, entry.fixedWeight));
  return faces.filter(face => fontFaceWeightDistance(face, entry.fixedWeight) === distance).map(face => ({ ...face, weight: 400 }));
}
