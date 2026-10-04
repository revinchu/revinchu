// Some WebKit ephemeral stores reject Blob/File backing stores even though
// ordinary structured-clone values work. Keep normal Blob storage elsewhere.
const PART_BYTES = 1 << 20, FORMAT = 'wixel-blob-bytes-v1';
const byteStores = new WeakSet();
const brokenBlob = error => ['UnknownError', 'DataCloneError'].includes(error?.name) && /(?:preparing|stor(?:e|ed|ing)|clon)[\s\S]*Blob|Blob[\s\S]*(?:stor|clon)/i.test(error.message ?? '');
const invalid = () => new Error('브라우저에 저장된 문서 조각이 없거나 손상되었습니다.');

async function blobParts(blob, check) {
  const parts = [];
  for (let start = 0; start < blob.size; start += PART_BYTES) {
    check(); parts.push(new Uint8Array(await blob.slice(start, start + PART_BYTES).arrayBuffer())); check();
  }
  return { format: FORMAT, size: blob.size, parts };
}
/** Retry only the known Blob preparation failure, after write() has confirmed
 * transaction abort. Quota, cancellation, CAS and unrelated errors never retry.
 * Large autosave callers pass one bounded cell chunk; library snapshots become
 * 1 MiB buffers instead of one contiguous ArrayBuffer. Their CAS remains atomic.
 */
export async function storeBlobCompatible(blob, write, check = () => {}) {
  if (!(blob instanceof Blob)) throw invalid();
  const owner = globalThis.indexedDB, known = owner && byteStores.has(owner);
  check();
  if (known) return write(await blobParts(blob, check));
  try { return await write(blob); }
  catch (error) {
    if (!brokenBlob(error)) throw error;
    check(); const result = await write(await blobParts(blob, check));
    if (owner && typeof owner === 'object') byteStores.add(owner);
    return result;
  }
}
/** Read old native Blob records and byte-part records without a whole-body copy. */
export function storedBlobStream(value) {
  if (value instanceof Blob) return value.stream();
  if (value?.format !== FORMAT || !Number.isSafeInteger(value.size) || value.size < 0 || !Array.isArray(value.parts)) throw invalid();
  let total = 0;
  for (const part of value.parts) {
    if (!(part instanceof Uint8Array) || !part.byteLength || part.byteLength > PART_BYTES) throw invalid();
    total += part.byteLength;
  }
  if (total !== value.size) throw invalid();
  let at = 0;
  return new ReadableStream({ pull(controller) {
    if (at === value.parts.length) controller.close();
    else controller.enqueue(value.parts[at++]);
  } });
}
export async function storedBlobText(value, gzip = false) {
  const stream = storedBlobStream(value);
  if (gzip && typeof DecompressionStream !== 'function') throw new Error('압축된 보관 문서를 열려면 최신 브라우저를 사용하세요.');
  return new Response(gzip ? stream.pipeThrough(new DecompressionStream('gzip')) : stream).text();
}
