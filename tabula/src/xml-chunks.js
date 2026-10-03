/** XML 조각을 제한된 문자열로 묶습니다. 전체 문서를 join하지 않습니다. */
export function createXmlChunks(limit = 1 << 20) {
  if (!Number.isInteger(limit) || limit < 2) throw new RangeError('XML 청크 크기가 올바르지 않습니다.');
  const chunks = [];
  let pending = [], length = 0, count = 0;
  const flush = () => { if (length) { chunks.push(pending.join('')); pending = []; length = 0; } };
  return {
    get count() { return count; },
    push(text) {
      if (typeof text !== 'string') throw new TypeError('XML 조각은 문자열이어야 합니다.');
      count++;
      for (let offset = 0; offset < text.length;) {
        const end = Math.min(text.length, offset + limit - length);
        pending.push(text.slice(offset, end)); length += end - offset; offset = end;
        if (length === limit) flush();
      }
    },
    finish() { flush(); return chunks; },
  };
}
