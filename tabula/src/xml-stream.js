// 한 항목만 XML 트리로 만들어 공유 문자열 전체 DOM을 메모리에 보관하지 않습니다.
export function* scanXmlChildren(bytes, itemName, chunkSize = 1 << 20) {
  const decoder = new TextDecoder();
  let pending = '', position = 0, start = -1, depth = 0;
  for (let offset = 0; offset < bytes.length || pending; offset += chunkSize) {
    const end = Math.min(bytes.length, offset + chunkSize), last = end === bytes.length;
    if (offset < bytes.length) pending += decoder.decode(bytes.subarray(offset, end), { stream: !last });
    let needMore = false;
    while (position < pending.length) {
      const at = pending.indexOf('<', position);
      if (at < 0) { position = pending.length; break; }
      let close, skip = false;
      if (pending.startsWith('<!--', at)) { close = pending.indexOf('-->', at + 4); if (close >= 0) close += 2; skip = true; }
      else if (pending.startsWith('<![CDATA[', at)) { close = pending.indexOf(']]>', at + 9); if (close >= 0) close += 2; skip = true; }
      else if (pending.startsWith('<?', at)) { close = pending.indexOf('?>', at + 2); if (close >= 0) close++; skip = true; }
      else {
        let quote = '';
        close = -1;
        for (let i = at + 1; i < pending.length; i++) {
          const char = pending[i];
          if (quote) { if (char === quote) quote = ''; }
          else if (char === '"' || char === "'") quote = char;
          else if (char === '>') { close = i; break; }
        }
      }
      if (close < 0) { position = at; needMore = true; break; }
      if (!skip && pending[at + 1] !== '!') {
        const tag = pending.slice(at, close + 1), match = /^<(\/?)(?:[\w.-]+:)?([\w.-]+)/.exec(tag);
        if (match) {
          const closing = !!match[1], self = /\/\s*>$/.test(tag), name = match[2];
          if (!closing) {
            if (depth === 1 && name === itemName) start = at;
            if (!self) depth++;
          } else depth = Math.max(0, depth - 1);
          if (start >= 0 && name === itemName && ((closing && depth === 1) || (self && depth === 1))) {
            yield pending.slice(start, close + 1);
            start = -1;
          }
        }
      }
      position = close + 1;
    }
    const keep = start >= 0 ? start : position;
    if (keep) { pending = pending.slice(keep); position -= keep; if (start >= 0) start -= keep; }
    if (last) {
      if (start >= 0 || (needMore && pending.trim())) throw new Error('XML 항목이 끝까지 저장되지 않았습니다.');
      break;
    }
  }
}
