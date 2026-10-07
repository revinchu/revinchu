// RFC 1951 streaming inflater. Output ownership transfers to the caller;
// retains a 32KiB distance window, one bounded output chunk and bounded
// Huffman lookup tables (at most 64KiB per literal/distance table).
const LEN_BASE = [3, 4, 5, 6, 7, 8, 9, 10, 11, 13, 15, 17, 19, 23, 27, 31, 35, 43, 51, 59, 67, 83, 99, 115, 131, 163, 195, 227, 258];
const LEN_EXTRA = [0, 0, 0, 0, 0, 0, 0, 0, 1, 1, 1, 1, 2, 2, 2, 2, 3, 3, 3, 3, 4, 4, 4, 4, 5, 5, 5, 5, 0];
const DIST_BASE = [1, 2, 3, 4, 5, 7, 9, 13, 17, 25, 33, 49, 65, 97, 129, 193, 257, 385, 513, 769, 1025, 1537, 2049, 3073, 4097, 6145, 8193, 12289, 16385, 24577];
const DIST_EXTRA = [0, 0, 0, 0, 1, 1, 2, 2, 3, 3, 4, 4, 5, 5, 6, 6, 7, 7, 8, 8, 9, 9, 10, 10, 11, 11, 12, 12, 13, 13];
const CL_ORDER = [16, 17, 18, 0, 8, 7, 9, 6, 10, 5, 11, 4, 12, 3, 13, 2, 14, 1, 15];
const MAX_CHUNK = 1 << 20;

export function inflateChunkSize(value = 64 << 10) {
  if (!Number.isSafeInteger(value) || value < 1 || value > MAX_CHUNK) throw new RangeError('압축 해제 청크 크기는 1바이트부터 1MiB까지여야 합니다.');
  return value;
}

function tree(lengths, { empty = false, complete = false } = {}) {
  const counts = new Uint16Array(16); let symbolsCount = 0, max = 0;
  for (const length of lengths) {
    if (length > 15) throw new Error('잘못된 허프만 코드 길이입니다.');
    if (length) { counts[length]++; symbolsCount++; max = Math.max(max, length); }
  }
  if (!symbolsCount) {
    if (empty) return { table: new Uint16Array(1), max: 0, mask: 0 };
    throw new Error('허프만 코드가 비어 있습니다.');
  }
  let left = 1;
  for (let length = 1; length < 16; length++) {
    left = (left << 1) - counts[length];
    if (left < 0) throw new Error('허프만 코드가 겹칩니다.');
  }
  if (left && (complete || max !== 1)) throw new Error('허프만 코드가 완전하지 않습니다.');
  const next = new Uint16Array(16), table = new Uint16Array(1 << max); let code = 0;
  for (let length = 1; length < 16; length++) { code = (code + counts[length - 1]) << 1; next[length] = code; }
  for (let symbol = 0; symbol < lengths.length; symbol++) {
    const length = lengths[symbol]; if (!length) continue;
    let canonical = next[length]++, reversed = 0;
    for (let bit = 0; bit < length; bit++) { reversed = (reversed << 1) | (canonical & 1); canonical >>>= 1; }
    const value = (symbol << 4) | length, step = 1 << length;
    for (let at = reversed; at < table.length; at += step) table[at] = value;
  }
  return { table, max, mask: table.length - 1 };
}
let FIXED = null;
function fixed() {
  if (!FIXED) {
    const lengths = new Uint8Array(288);
    lengths.fill(8, 0, 144); lengths.fill(9, 144, 256); lengths.fill(7, 256, 280); lengths.fill(8, 280, 288);
    FIXED = { lit: tree(lengths), dist: tree(new Uint8Array(32).fill(5)) };
  }
  return FIXED;
}

/** Raw deflate bytes -> independent bounded byte chunks. sizeLimit prevents
 * an incorrect ZIP size from producing unbounded work before the final check.
 * Returning early releases inflater state without touching its input bytes. */
export function* inflateChunks(data, { chunkSize = 64 << 10, sizeLimit = Number.MAX_SAFE_INTEGER } = {}) {
  chunkSize = inflateChunkSize(chunkSize);
  if (!(data instanceof Uint8Array)) throw new TypeError('압축 데이터는 바이트 배열이어야 합니다.');
  if (!Number.isSafeInteger(sizeLimit) || sizeLimit < 0) throw new RangeError('ZIP 항목 크기가 올바르지 않습니다.');
  let position = 0, bitBuffer = 0, bitCount = 0, written = 0, used = 0;
  let output = new Uint8Array(chunkSize), window = new Uint8Array(32768);
  const bits = count => {
    while (bitCount < count) {
      if (position >= data.length) throw new Error('압축 데이터가 잘렸습니다.');
      bitBuffer |= data[position++] << bitCount; bitCount += 8;
    }
    const value = bitBuffer & ((1 << count) - 1);
    bitBuffer >>>= count; bitCount -= count; return value;
  };
  const decode = tree => {
    // Peek only available bytes. A short final code is valid even when fewer
    // than tree.max bits remain; its replicated lookup entries still match.
    while (bitCount < tree.max && position < data.length) { bitBuffer |= data[position++] << bitCount; bitCount += 8; }
    const value = tree.table[bitBuffer & tree.mask], length = value & 15;
    if (!length || length > bitCount) {
      if (bitCount < tree.max) throw new Error('압축 데이터가 잘렸습니다.');
      throw new Error('잘못된 허프만 코드입니다.');
    }
    bitBuffer >>>= length; bitCount -= length; return value >>> 4;
  };
  const put = value => {
    if (written >= sizeLimit) throw new Error('ZIP 항목 크기를 넘는 압축 데이터입니다.');
    window[written & 32767] = value; output[used++] = value; written++;
    return used === chunkSize;
  };
  try {
    let last = false;
    while (!last) {
      last = !!bits(1); const type = bits(2);
      if (type === 0) {
        const skip = bitCount & 7; bitBuffer >>>= skip; bitCount -= skip;
        const length = bits(16), inverse = bits(16);
        if ((length ^ 0xffff) !== inverse) throw new Error('비압축 블록 길이가 올바르지 않습니다.');
        for (let i = 0; i < length; i++) if (put(bits(8))) {
          yield output; output = new Uint8Array(chunkSize); used = 0;
        }
        continue;
      }
      if (type === 3) throw new Error('잘못된 압축 블록 형식입니다.');
      let tables;
      if (type === 1) tables = fixed();
      else {
        const literalCount = bits(5) + 257, distanceCount = bits(5) + 1, codeCount = bits(4) + 4;
        if (literalCount > 286) throw new Error('허프만 리터럴 코드 수가 올바르지 않습니다.');
        const codeLengths = new Uint8Array(19);
        for (let i = 0; i < codeCount; i++) codeLengths[CL_ORDER[i]] = bits(3);
        const codeTree = tree(codeLengths, { complete: true }), lengths = new Uint8Array(literalCount + distanceCount);
        for (let i = 0; i < lengths.length;) {
          const symbol = decode(codeTree);
          if (symbol < 16) { lengths[i++] = symbol; continue; }
          let count, value = 0;
          if (symbol === 16) {
            if (!i) throw new Error('허프만 코드 길이 반복이 올바르지 않습니다.');
            count = bits(2) + 3; value = lengths[i - 1];
          } else if (symbol === 17) count = bits(3) + 3;
          else count = bits(7) + 11;
          if (i + count > lengths.length) throw new Error('허프만 코드 길이 반복이 너무 깁니다.');
          lengths.fill(value, i, i + count); i += count;
        }
        if (!lengths[256]) throw new Error('압축 블록 종료 코드가 없습니다.');
        tables = { lit: tree(lengths.subarray(0, literalCount)), dist: tree(lengths.subarray(literalCount), { empty: true }) };
      }
      for (;;) {
        const symbol = decode(tables.lit);
        if (symbol === 256) break;
        if (symbol < 256) {
          if (put(symbol)) { yield output; output = new Uint8Array(chunkSize); used = 0; }
          continue;
        }
        const lengthIndex = symbol - 257;
        if (lengthIndex >= LEN_BASE.length) throw new Error('잘못된 압축 길이 코드입니다.');
        const length = LEN_BASE[lengthIndex] + bits(LEN_EXTRA[lengthIndex]);
        const distanceIndex = decode(tables.dist);
        if (distanceIndex >= DIST_BASE.length) throw new Error('잘못된 압축 거리 코드입니다.');
        const distance = DIST_BASE[distanceIndex] + bits(DIST_EXTRA[distanceIndex]);
        if (distance > written || distance > 32768) throw new Error('압축 데이터의 참조 거리가 올바르지 않습니다.');
        if (written + length > sizeLimit) throw new Error('ZIP 항목 크기를 넘는 압축 데이터입니다.');
        let remaining = length;
        while (remaining) {
          const count = Math.min(remaining, chunkSize - used), start = written;
          // Inline the match copy so bounds/output-flush checks happen once per
          // segment instead of once per byte. Forward overlap remains valid.
          for (let i = 0; i < count; i++) {
            const value = window[(start + i - distance) & 32767];
            window[(start + i) & 32767] = value; output[used + i] = value;
          }
          written += count; used += count; remaining -= count;
          if (used === chunkSize) { yield output; output = new Uint8Array(chunkSize); used = 0; }
        }
      }
    }
    if (position - Math.floor(bitCount / 8) !== data.length) throw new Error('압축 데이터 끝에 불필요한 바이트가 있습니다.');
    if (used) yield output.slice(0, used);
  } finally { output = null; window = null; }
}
