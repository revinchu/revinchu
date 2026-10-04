import { createJsonSizer } from './json-size.js';

const PART_BYTES = 1 << 20, PART_ITEMS = 2048;
const TYPES = { Float64Array, Float32Array, Int32Array, Uint32Array, Int16Array, Uint16Array, Int8Array, Uint8Array, Uint8ClampedArray };
const invalid = () => new Error('저장된 배열 조각이 없거나 완전하지 않습니다. 이전 저장본을 확인하세요.');

/** One bounded array copy per completed IDB write. An exceptional single value
 * stays intact; no whole dictionary, numeric column or pivot matrix is cloned. */
export async function storeArrayParts(array, keyForPart, put, checkpoint) {
  const plain = Array.isArray(array), Type = TYPES[array?.constructor?.name];
  if (!plain && (!Type || !(array instanceof Type))) throw invalid();
  const parts = [], per = plain ? PART_ITEMS : Math.floor(PART_BYTES / Type.BYTES_PER_ELEMENT);
  for (let start = 0; start < array.length;) {
    await checkpoint();
    let end = Math.min(array.length, start + per);
    if (plain) { const measure = createJsonSizer(PART_BYTES, { conservativeStrings: true }).size; let size = 0; end = start; while (end < array.length && end - start < per) {
      const next = measure(array[end]) + 16;
      if (end > start && size + next > PART_BYTES) break;
      size += next; end++; if (size > PART_BYTES) break;
    } }
    const key = keyForPart(parts.length);
    await put(key, array.slice(start, end)); parts.push(key); start = end;
    await checkpoint();
  }
  return { parts, len: array.length, kind: plain ? 'Array' : array.constructor.name };
}

export async function restoreArrayParts(info, get, checkpoint) {
  const plain = info?.kind === 'Array', Type = TYPES[info?.kind];
  if ((!plain && !Type) || !Number.isSafeInteger(info?.len) || info.len < 0 || !Array.isArray(info.parts)) throw invalid();
  const out = plain ? new Array(info.len) : new Type(info.len); let at = 0;
  for (const key of info.parts) {
    await checkpoint(); const part = await get(key);
    if (!(plain ? Array.isArray(part) : part instanceof Type) || !part.length || at + part.length > info.len) throw invalid();
    if (plain) { for (let i = 0; i < part.length; i++) out[at + i] = part[i]; }
    else out.set(part, at);
    at += part.length; await checkpoint();
  }
  if (at !== info.len) throw invalid();
  return out;
}
