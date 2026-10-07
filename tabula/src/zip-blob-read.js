// Blob ZIP index: only the end records, central directory and local headers
// are read here. Entry payloads remain lazy ranges of the original File/Blob.
const decoder = new TextDecoder();
function limit(value, name) {
  if (value !== Infinity && (!Number.isSafeInteger(value) || value < 0)) throw new RangeError(name + ' 예산이 올바르지 않습니다.');
  return value;
}
function span(blob, offset, length) {
  if (!Number.isSafeInteger(offset) || !Number.isSafeInteger(length) || offset < 0 || length < 0 || offset + length > blob.size) throw new Error('ZIP 범위가 파일 밖에 있습니다.');
}
async function read(blob, offset, length) {
  span(blob, offset, length);
  const bytes = new Uint8Array(await blob.slice(offset, offset + length).arrayBuffer());
  if (bytes.length !== length) throw new Error('ZIP 파일 읽기가 끝까지 완료되지 않았습니다.');
  return bytes;
}
function uint64(view, offset, end = view.byteLength) {
  if (offset < 0 || offset + 8 > end) throw new Error('ZIP64 정보가 잘렸습니다.');
  const value = view.getUint32(offset, true) + view.getUint32(offset + 4, true) * 0x100000000;
  if (!Number.isSafeInteger(value)) throw new Error('ZIP64 크기를 정확하게 읽을 수 없습니다.');
  return value;
}
function rawLoader(blob, localOffset, length, directoryOffset, method, maxEntryBytes) {
  // Do not capture the central-directory buffer in each payload closure.
  return async () => {
    if (length > maxEntryBytes) throw new RangeError('ZIP 항목이 설정한 읽기 예산을 넘습니다.');
    const header = await read(blob, localOffset, 30), local = new DataView(header.buffer, header.byteOffset, header.byteLength);
    if (local.getUint32(0, true) !== 0x04034b50 || local.getUint16(8, true) !== method || (local.getUint16(6, true) & 1)) throw new Error('ZIP 항목 머리글이 손상되었습니다.');
    const offset = localOffset + 30 + local.getUint16(26, true) + local.getUint16(28, true);
    span(blob, offset, length);
    if (offset + length > directoryOffset) throw new Error('ZIP 항목 본문이 목록과 겹칩니다.');
    return read(blob, offset, length);
  };
}

/** Return directory entries with async loadRaw(), never blob.arrayBuffer(). */
export async function blobZipEntries(blob, { maxDirectoryBytes = Infinity, maxEntryBytes = Infinity } = {}) {
  if (!blob || !Number.isSafeInteger(blob.size) || blob.size < 22 || typeof blob.slice !== 'function') throw new TypeError('ZIP 파일 Blob이 올바르지 않습니다.');
  maxDirectoryBytes = limit(maxDirectoryBytes, 'ZIP 목록'); maxEntryBytes = limit(maxEntryBytes, 'ZIP 항목');
  const tailOffset = Math.max(0, blob.size - 65557), tail = await read(blob, tailOffset, blob.size - tailOffset);
  const tailView = new DataView(tail.buffer, tail.byteOffset, tail.byteLength); let at = -1;
  for (let i = tail.length - 22; i >= 0; i--) {
    if (tailView.getUint32(i, true) === 0x06054b50 && i + 22 + tailView.getUint16(i + 20, true) === tail.length) { at = i; break; }
  }
  if (at < 0) throw new Error('ZIP 파일 끝부분이 없거나 잘렸습니다.');
  const endOffset = tailOffset + at;
  if (tailView.getUint16(at + 4, true) || tailView.getUint16(at + 6, true) || tailView.getUint16(at + 8, true) !== tailView.getUint16(at + 10, true)) throw new Error('분할 ZIP 파일은 지원하지 않습니다.');
  let count = tailView.getUint16(at + 10, true), directorySize = tailView.getUint32(at + 12, true), directoryOffset = tailView.getUint32(at + 16, true);
  let directoryEnd = endOffset;
  let locator = null;
  if (endOffset >= 20 && (count === 0xffff || directorySize === 0xffffffff || directoryOffset === 0xffffffff)) {
    const bytes = await read(blob, endOffset - 20, 20), view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (view.getUint32(0, true) === 0x07064b50) locator = view;
  }
  if (locator) {
    if (locator.getUint32(4, true) || locator.getUint32(16, true) !== 1) throw new Error('분할 ZIP64 파일은 지원하지 않습니다.');
    const offset = uint64(locator, 8), bytes = await read(blob, offset, 56), view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (view.getUint32(0, true) !== 0x06064b50 || view.getUint32(16, true) || view.getUint32(20, true)) throw new Error('ZIP64 파일 끝 정보가 올바르지 않습니다.');
    const size = uint64(view, 4); if (size < 44 || offset + 12 + size > endOffset - 20) throw new Error('ZIP64 파일 끝 정보가 잘렸습니다.');
    count = uint64(view, 32); if (uint64(view, 24) !== count) throw new Error('ZIP64 항목 수가 올바르지 않습니다.');
    directorySize = uint64(view, 40); directoryOffset = uint64(view, 48); directoryEnd = offset;
  } else if (directorySize === 0xffffffff || directoryOffset === 0xffffffff) throw new Error('ZIP64 파일 끝 정보가 없습니다.');
  if (directorySize > maxDirectoryBytes) throw new RangeError('ZIP 목록이 설정한 읽기 예산을 넘습니다.');
  span(blob, directoryOffset, directorySize);
  if (directoryOffset + directorySize > directoryEnd || count > Math.floor(directorySize / 46)) throw new Error('ZIP 목록 크기가 올바르지 않습니다.');
  const central = await read(blob, directoryOffset, directorySize), view = new DataView(central.buffer, central.byteOffset, central.byteLength), entries = [];
  let p = 0;
  for (let i = 0; i < count; i++) {
    if (p + 46 > central.length || view.getUint32(p, true) !== 0x02014b50) throw new Error('ZIP 목록이 손상되었습니다.');
    const flags = view.getUint16(p + 8, true), method = view.getUint16(p + 10, true), crc = view.getUint32(p + 16, true);
    const nameLength = view.getUint16(p + 28, true), extraLength = view.getUint16(p + 30, true), commentLength = view.getUint16(p + 32, true), disk = view.getUint16(p + 34, true);
    const stop = p + 46 + nameLength + extraLength + commentLength;
    if (stop > central.length) throw new Error('ZIP 항목 정보가 잘렸습니다.');
    if (flags & 1) throw new Error('암호화된 ZIP 항목은 지원하지 않습니다.');
    if (disk) throw new Error('분할 ZIP 파일은 지원하지 않습니다.');
    if (method !== 0 && method !== 8) throw new Error('지원하지 않는 ZIP 압축 방식입니다.');
    let size = view.getUint32(p + 24, true), compressedSize = view.getUint32(p + 20, true), localOffset = view.getUint32(p + 42, true);
    if (size === 0xffffffff || compressedSize === 0xffffffff || localOffset === 0xffffffff) {
      let found = false;
      for (let extra = p + 46 + nameLength, end = extra + extraLength; extra + 4 <= end;) {
        const kind = view.getUint16(extra, true), length = view.getUint16(extra + 2, true), next = extra + 4 + length;
        if (next > end) throw new Error('ZIP64 항목 정보가 잘렸습니다.');
        if (kind === 1) {
          let value = extra + 4;
          if (size === 0xffffffff) { size = uint64(view, value, next); value += 8; }
          if (compressedSize === 0xffffffff) { compressedSize = uint64(view, value, next); value += 8; }
          if (localOffset === 0xffffffff) localOffset = uint64(view, value, next);
          found = true; break;
        }
        extra = next;
      }
      if (!found) throw new Error('ZIP64 항목 크기 정보가 없습니다.');
    }
    const name = decoder.decode(central.subarray(p + 46, p + 46 + nameLength)); p = stop;
    if (name.endsWith('/')) continue;
    span(blob, localOffset, 30);
    entries.push({ name, method, size, crc, compressedSize, loadRaw: rawLoader(blob, localOffset, compressedSize, directoryOffset, method, maxEntryBytes) });
  }
  if (p !== central.length) throw new Error('ZIP 목록의 항목 수와 크기가 다릅니다.');
  return entries;
}
