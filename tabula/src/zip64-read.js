// ZIP64-only metadata. Classic ZIP parsing stays in zip.js.
function span(view, at, length) {
  if (!Number.isSafeInteger(at) || at < 0 || length < 0 || at + length > view.byteLength) throw new Error('ZIP64 정보가 잘렸습니다.');
}
function read64(view, at) {
  span(view, at, 8); const value = view.getUint32(at, true) + view.getUint32(at + 4, true) * 0x100000000;
  if (!Number.isSafeInteger(value)) throw new RangeError('ZIP64 크기를 정확하게 읽을 수 없습니다.');
  return value;
}

export function zip64Directory(view, eocd) {
  const locator = eocd - 20; span(view, locator, 20);
  if (view.getUint32(locator, true) !== 0x07064b50 || view.getUint32(locator + 4, true) || view.getUint32(locator + 16, true) !== 1) throw new Error('ZIP64 디렉터리 위치가 올바르지 않습니다.');
  const end = read64(view, locator + 8); span(view, end, 56);
  if (view.getUint32(end, true) !== 0x06064b50 || view.getUint32(end + 16, true) || view.getUint32(end + 20, true)) throw new Error('ZIP64 디렉터리가 손상되었습니다.');
  const recordSize = read64(view, end + 4), count = read64(view, end + 32), size = read64(view, end + 40), offset = read64(view, end + 48);
  if (recordSize < 44 || end + 12 + recordSize > locator || read64(view, end + 24) !== count) throw new Error('ZIP64 디렉터리 크기가 올바르지 않습니다.');
  span(view, offset, size);
  if (offset + size > end || count > Math.floor(size / 46)) throw new Error('ZIP64 디렉터리 크기가 올바르지 않습니다.');
  return { count, size, offset };
}

export function zip64Entry(view, start, length, size, compSize, local) {
  span(view, start, length); const stop = start + length;
  for (let p = start; p + 4 <= stop;) {
    const id = view.getUint16(p, true), length = view.getUint16(p + 2, true), end = p + 4 + length;
    if (end > stop) throw new Error('ZIP64 항목 정보가 잘렸습니다.');
    if (id === 1) {
      let at = p + 4;
      const field = () => { if (at + 8 > end) throw new Error('ZIP64 항목 정보가 잘렸습니다.'); const value = read64(view, at); at += 8; return value; };
      if (size === 0xffffffff) size = field(); if (compSize === 0xffffffff) compSize = field(); if (local === 0xffffffff) local = field();
      span(view, local, 30);
      if (view.getUint32(local, true) !== 0x04034b50) throw new Error('ZIP64 항목 머리글이 손상되었습니다.');
      const body = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      span(view, body, compSize);
      return { size, compSize, local };
    }
    p = end;
  }
  throw new Error('ZIP64 항목 크기 정보가 없습니다.');
}
