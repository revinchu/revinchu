// Text decoding for large local files; never buffers the entire decoded document.
export async function detectTextFileEncoding(file, onProgress) {
  const head = new Uint8Array(await file.slice(0, 3).arrayBuffer());
  if (head[0] === 0xff && head[1] === 0xfe) return 'utf-16le';
  if (head[0] === 0xfe && head[1] === 0xff) return 'utf-16be';
  if (head[0] === 0xef && head[1] === 0xbb && head[2] === 0xbf) return 'utf-8';
  // A short ASCII prefix cannot distinguish UTF-8 from Korean Excel CP949.
  const reader = file.stream().getReader(), decoder = new TextDecoder('utf-8', { fatal: true });
  let scanned = 0, complete = false;
  try {
    for (;;) {
      const { value, done } = await reader.read(); // I/O failures are not encoding failures.
      if (done) { complete = true; break; }
      try { decoder.decode(value, { stream: true }); } catch { return 'euc-kr'; }
      scanned += value.byteLength;
      await onProgress?.(scanned, file.size);
    }
    try { decoder.decode(); } catch { return 'euc-kr'; }
    return 'utf-8';
  } finally {
    if (!complete) { try { await reader.cancel(); } catch {} }
    reader.releaseLock();
  }
}

export async function readDecodedChunks(file, encoding, consume) {
  const reader = file.stream().getReader(), decoder = new TextDecoder(encoding);
  let read = 0, complete = false;
  try {
    for (;;) {
      const { value, done } = await reader.read();
      if (done) break;
      read += value.byteLength;
      const text = decoder.decode(value, { stream: true });
      if (text) await consume(text, read);
    }
    const tail = decoder.decode();
    if (tail) await consume(tail, read);
    complete = true;
  } finally {
    if (!complete) { try { await reader.cancel(); } catch {} }
    reader.releaseLock();
  }
}
