// ZIP32/ZIP64 records. Sizes stay exact Numbers until their 64-bit LE encoding.
const ZIP32_SENTINEL = 0xffffffff;
const utf8 = new TextEncoder();

export function zipCount(value) {
  if (!Number.isSafeInteger(value) || value < 0) throw new RangeError('ZIP 크기가 정확하게 저장할 수 있는 범위를 넘었습니다.');
  return value;
}
function uint64(view, at, value) {
  zipCount(value); view.setUint32(at, value % 0x100000000, true); view.setUint32(at + 4, Math.floor(value / 0x100000000), true);
}
function nameBytes(name) {
  const bytes = utf8.encode(name);
  if (!bytes.length || bytes.length > 0xffff) throw new RangeError('ZIP 항목 이름 길이가 올바르지 않습니다.');
  return bytes;
}
function stamp(now = new Date()) {
  return { time: (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1),
    date: ((Math.max(1980, Math.min(2107, now.getFullYear())) - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate() };
}

/** Streaming local header: bit 3 means CRC and sizes follow the payload. */
export function zipStreamLocalRecord(name, method, zip64, now) {
  const bytes = nameBytes(name), { time, date } = stamp(now), extra = zip64 ? 20 : 0;
  const out = new Uint8Array(30 + bytes.length + extra), view = new DataView(out.buffer);
  view.setUint32(0, 0x04034b50, true); view.setUint16(4, zip64 ? 45 : 20, true); view.setUint16(6, 0x0808, true);
  view.setUint16(8, method, true); view.setUint16(10, time, true); view.setUint16(12, date, true);
  view.setUint16(26, bytes.length, true); view.setUint16(28, extra, true); out.set(bytes, 30);
  if (zip64) {
    view.setUint32(18, ZIP32_SENTINEL, true); view.setUint32(22, ZIP32_SENTINEL, true);
    view.setUint16(30 + bytes.length, 1, true); view.setUint16(32 + bytes.length, 16, true);
  }
  return out;
}

/** Final local header for seekable output. The reserved local extra holds real sizes. */
export function zipSeekLocalRecord(name, method, crc, compressedSize, size, now) {
  zipCount(size); zipCount(compressedSize);
  const out = zipStreamLocalRecord(name, method, true, now), view = new DataView(out.buffer), big = size >= ZIP32_SENTINEL || compressedSize >= ZIP32_SENTINEL;
  view.setUint16(4, big ? 45 : 20, true); view.setUint16(6, 0x0800, true); view.setUint32(14, crc, true);
  view.setUint32(18, big ? ZIP32_SENTINEL : compressedSize, true); view.setUint32(22, big ? ZIP32_SENTINEL : size, true);
  const extra = 30 + view.getUint16(26, true); uint64(view, extra + 4, size); uint64(view, extra + 12, compressedSize);
  return out;
}

export function zipDescriptorRecord(crc, compressedSize, size, zip64) {
  zipCount(compressedSize); zipCount(size);
  if (!zip64 && (size >= ZIP32_SENTINEL || compressedSize >= ZIP32_SENTINEL)) throw new RangeError('ZIP32 항목 크기를 넘었습니다. ZIP64로 저장해 주세요.');
  const out = new Uint8Array(zip64 ? 24 : 16), view = new DataView(out.buffer);
  view.setUint32(0, 0x08074b50, true); view.setUint32(4, crc, true);
  if (zip64) { uint64(view, 8, compressedSize); uint64(view, 16, size); }
  else { view.setUint32(8, compressedSize, true); view.setUint32(12, size, true); }
  return out;
}

export function zipCentralRecord({ name, method, crc, size, compressedSize, offset, zip64 = false, descriptor = true }, now) {
  zipCount(size); zipCount(compressedSize); zipCount(offset);
  const bytes = nameBytes(name), { time, date } = stamp(now);
  const bigSize = zip64 || size >= ZIP32_SENTINEL, bigCompressed = zip64 || compressedSize >= ZIP32_SENTINEL, bigOffset = offset >= ZIP32_SENTINEL;
  const fields = []; if (bigSize) fields.push(size); if (bigCompressed) fields.push(compressedSize); if (bigOffset) fields.push(offset);
  const extra = fields.length ? 4 + fields.length * 8 : 0, out = new Uint8Array(46 + bytes.length + extra), view = new DataView(out.buffer);
  view.setUint32(0, 0x02014b50, true); view.setUint16(4, extra ? 45 : 20, true); view.setUint16(6, extra ? 45 : 20, true);
  view.setUint16(8, descriptor ? 0x0808 : 0x0800, true); view.setUint16(10, method, true); view.setUint16(12, time, true); view.setUint16(14, date, true);
  view.setUint32(16, crc, true); view.setUint32(20, bigCompressed ? ZIP32_SENTINEL : compressedSize, true); view.setUint32(24, bigSize ? ZIP32_SENTINEL : size, true);
  view.setUint16(28, bytes.length, true); view.setUint16(30, extra, true); view.setUint32(42, bigOffset ? ZIP32_SENTINEL : offset, true); out.set(bytes, 46);
  if (extra) {
    const start = 46 + bytes.length; view.setUint16(start, 1, true); view.setUint16(start + 2, fields.length * 8, true);
    fields.forEach((value, i) => uint64(view, start + 4 + i * 8, value));
  }
  return out;
}

/** Directory end records, independent of output capacity; usable with disk sinks. */
export function zipEndRecords(count, directorySize, directoryOffset, forceZip64 = false) {
  zipCount(count); zipCount(directorySize); zipCount(directoryOffset);
  const zip64 = forceZip64 || count >= 0xffff || directorySize >= ZIP32_SENTINEL || directoryOffset >= ZIP32_SENTINEL;
  const out = [];
  if (zip64) {
    const end = new Uint8Array(56), view = new DataView(end.buffer);
    view.setUint32(0, 0x06064b50, true); uint64(view, 4, 44); view.setUint16(12, 45, true); view.setUint16(14, 45, true);
    uint64(view, 24, count); uint64(view, 32, count); uint64(view, 40, directorySize); uint64(view, 48, directoryOffset); out.push(end);
    const locator = new Uint8Array(20), lv = new DataView(locator.buffer);
    lv.setUint32(0, 0x07064b50, true); uint64(lv, 8, zipCount(directoryOffset + directorySize)); lv.setUint32(16, 1, true); out.push(locator);
  }
  const end = new Uint8Array(22), view = new DataView(end.buffer);
  view.setUint32(0, 0x06054b50, true); view.setUint16(8, zip64 ? 0xffff : count, true); view.setUint16(10, zip64 ? 0xffff : count, true);
  view.setUint32(12, zip64 ? ZIP32_SENTINEL : directorySize, true); view.setUint32(16, zip64 ? ZIP32_SENTINEL : directoryOffset, true); out.push(end);
  return out;
}
