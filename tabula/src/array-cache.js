// Excel 배열 수식의 저장 결과. 실제 저장된 좌표만 보존하며 빈 범위를 펼치지 않습니다.
const checked = new WeakSet();
export function arrayCache(data) {
  if (!data || typeof data !== 'object') return null;
  if (checked.has(data)) return data;
  const { h, w, values } = data;
  if (!Number.isInteger(h) || h < 1 || h > 1048576 || !Number.isInteger(w) || w < 1 || w > 16384
      || !Array.isArray(values) || values.length % 3) return null;
  let previous = -1;
  for (let i = 0; i < values.length; i += 3) {
    const r = values[i], c = values[i + 1], v = values[i + 2], key = r * w + c;
    if (!Number.isInteger(r) || r < 0 || r >= h || !Number.isInteger(c) || c < 0 || c >= w || key <= previous) return null;
    if (!(v === null || typeof v === 'string' || typeof v === 'boolean' || typeof v === 'number' && Number.isFinite(v)
        || v && typeof v === 'object' && typeof v.error === 'string')) return null;
    previous = key;
  }
  // 소유 셀의 dirty 상태만 바뀝니다. 값 목록은 복원·서식 변경·Undo 사이에서 공유합니다.
  Object.freeze(values); Object.freeze(data); checked.add(data);
  return data;
}
export function arrayCacheValue(data, r, c) {
  if (r < 0 || c < 0 || r >= data.h || c >= data.w) return undefined;
  const a = data.values, key = r * data.w + c;
  let lo = 0, hi = a.length / 3;
  while (lo < hi) { const m = Math.floor((lo + hi) / 2), at = m * 3; if (a[at] * data.w + a[at + 1] < key) lo = m + 1; else hi = m; }
  const at = lo * 3;
  return at < a.length && a[at] === r && a[at + 1] === c ? a[at + 2] : undefined;
}
export function* spillOffsets(sp) {
  if (sp.cachedArray) {
    const a = sp.cachedArray.values;
    for (let i = 0; i < a.length; i += 3) if (a[i] || a[i + 1]) yield [a[i], a[i + 1]];
  } else for (let r = 0; r < sp.h; r++) for (let c = 0; c < sp.w; c++) if (r || c) yield [r, c];
}
