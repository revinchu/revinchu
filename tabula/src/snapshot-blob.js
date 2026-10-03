// 기존 JSON 저장 형식을 유지하며 가변 Workbook를 await 전에 불변 Blob으로 고정합니다.
// 문자열 경계를 UTF-16 서로게이트 쌍 사이에 두면 Blob이 대체 문자로 바꾸므로 분리하지 않습니다.
export function jsonPartsBlob(parts, chunkSize = 65536) {
  if (!Number.isSafeInteger(chunkSize) || chunkSize < 2) throw new RangeError('조각 크기는 2 이상의 정수여야 합니다.');
  const blobs = [];
  let pending = '', high = '';
  const flush = () => { if (pending) { blobs.push(new Blob([pending])); pending = ''; } };
  for (const part of parts) {
    const text = high + part; high = '';
    let start = 0, end = text.length;
    if (end && text.charCodeAt(end - 1) >= 0xd800 && text.charCodeAt(end - 1) <= 0xdbff) {
      high = text[end - 1]; end--;
    }
    while (start < end) {
      let cut = Math.min(end, start + chunkSize - pending.length);
      if (cut < end && text.charCodeAt(cut - 1) >= 0xd800 && text.charCodeAt(cut - 1) <= 0xdbff
          && text.charCodeAt(cut) >= 0xdc00 && text.charCodeAt(cut) <= 0xdfff) cut--;
      if (cut === start) { flush(); continue; }
      pending += text.slice(start, cut); start = cut;
      if (pending.length >= chunkSize - 1) flush();
    }
  }
  pending += high;
  flush();
  return new Blob(blobs, { type: 'application/json' });
}

/** 저장 호출 시점의 문서 본문을 고정합니다. 충돌 사본은 같은 본문에 ID만 바꿉니다. */
export function librarySnapshot(book, metadata) {
  const body = book.serializeBlob(), frozen = JSON.parse(JSON.stringify(metadata));
  const wrap = header => new Blob([JSON.stringify(header).slice(0, -1), ',"workbook":', body, '}'], { type: 'application/json' });
  return { blob: wrap(frozen), withDocId: id => wrap({ ...frozen, docId: id }) };
}

/** 읽기 청크 경계와 무관한 정확한 바이트 비교. 전체 JSON/해시 사본은 만들지 않습니다. */
export async function equalByteStreams(a, b) {
  const left = a.getReader(), right = b.getReader();
  let l = new Uint8Array(), r = new Uint8Array(), li = 0, ri = 0, ld = false, rd = false;
  try {
    for (;;) {
      while (li === l.length && !ld) { const next = await left.read(); ld = next.done; l = next.value ?? new Uint8Array(); li = 0; }
      while (ri === r.length && !rd) { const next = await right.read(); rd = next.done; r = next.value ?? new Uint8Array(); ri = 0; }
      if (ld || rd) return ld && rd;
      const n = Math.min(l.length - li, r.length - ri);
      for (let i = 0; i < n; i++) if (l[li + i] !== r[ri + i]) return false;
      li += n; ri += n;
    }
  } finally {
    await Promise.allSettled([left.cancel(), right.cancel()]);
    left.releaseLock(); right.releaseLock();
  }
}
