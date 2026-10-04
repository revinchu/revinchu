// 합성 packed JSON의 전송·저장만 검증합니다. 청크 내용은 실제 사용자 문서가 아닙니다.
import { createHash } from 'node:crypto';

export function packedPayload(bytes) {
  const prefix = Buffer.from('{"format":"wixel-packed","version":1,"compression":"none","chunks":[');
  const suffix = Buffer.from(']}'), piece = 'A'.repeat(256 * 1024);
  function* parts() {
    yield prefix;
    let left = bytes - prefix.length - suffix.length, first = true;
    while (left >= (first ? 6 : 7)) {
      const overhead = first ? 2 : 3, count = Math.min(piece.length, Math.floor((left - overhead) / 4) * 4);
      if (count < 4) break;
      const part = Buffer.from((first ? '"' : ',"') + piece.slice(0, count) + '"');
      yield part; left -= part.length; first = false;
    }
    yield suffix;
    if (left) yield Buffer.from(' '.repeat(left));
  }
  const hash = createHash('sha256'); let length = 0;
  for (const part of parts()) { hash.update(part); length += part.length; }
  if (length !== bytes) throw new Error('Synthetic payload length mismatch');
  return { bytes, sha256: hash.digest('hex'), stream() {
    const it = parts();
    return new ReadableStream({ pull(controller) { const next = it.next(); if (next.done) controller.close(); else controller.enqueue(next.value); }, cancel() { it.return?.(); } });
  } };
}

export async function streamDigest(body) {
  const hash = createHash('sha256'); let bytes = 0;
  for await (const part of body) { hash.update(part); bytes += part.length; }
  return { bytes, sha256: hash.digest('hex') };
}
