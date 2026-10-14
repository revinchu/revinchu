// OLE 복합 문서 (Compound File Binary: .ppt · .doc · .xls · OLE 개체) 읽기 — DOM 없음
// openCfb(bytes) → { names: [스트림 이름], stream(name) → Uint8Array | null }

export const isCfb = (b) => !!b && b.length > 512 && b[0] === 0xd0 && b[1] === 0xcf && b[2] === 0x11 && b[3] === 0xe0 && b[4] === 0xa1 && b[5] === 0xb1 && b[6] === 0x1a && b[7] === 0xe1;

export function openCfb(buf) {
  if (!isCfb(buf)) throw new Error('OLE 복합 문서가 아닙니다');
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const secSize = 1 << dv.getUint16(30, true);
  const miniSize = 1 << dv.getUint16(32, true);
  const dirStart = dv.getUint32(48, true);
  const miniCutoff = dv.getUint32(56, true);
  const miniFatStart = dv.getUint32(60, true);
  let difatSec = dv.getUint32(68, true);
  const END = 0xfffffffe;
  const off = (s) => (s + 1) * secSize;
  // FAT 섹터 목록 (헤더의 109개 + DIFAT 사슬)
  const fatSecs = [];
  for (let i = 0; i < 109; i++) { const s = dv.getUint32(76 + i * 4, true); if (s < 0xfffffffa) fatSecs.push(s); }
  for (let guard = 0; difatSec < 0xfffffffa && guard < 100000 && off(difatSec) + secSize <= buf.length; guard++) {
    const n = secSize / 4 - 1;
    for (let i = 0; i < n; i++) { const s = dv.getUint32(off(difatSec) + i * 4, true); if (s < 0xfffffffa) fatSecs.push(s); }
    difatSec = dv.getUint32(off(difatSec) + n * 4, true);
  }
  const perSec = secSize / 4;
  const fat = new Uint32Array(fatSecs.length * perSec);
  fatSecs.forEach((s, k) => { for (let i = 0; i < perSec && off(s) + i * 4 + 4 <= buf.length; i++) fat[k * perSec + i] = dv.getUint32(off(s) + i * 4, true); });
  const chain = (start, table) => {
    const out = [];
    const seen = new Set();
    for (let s = start; s < table.length && s !== END && !seen.has(s) && out.length < 1e7; s = table[s]) { seen.add(s); out.push(s); }
    return out;
  };
  const readChain = (start) => {
    const secs = chain(start, fat);
    const out = new Uint8Array(secs.length * secSize);
    secs.forEach((s, i) => { if (off(s) < buf.length) out.set(buf.subarray(off(s), Math.min(buf.length, off(s) + secSize)), i * secSize); });
    return out;
  };
  const dir = readChain(dirStart);
  const ddv = new DataView(dir.buffer);
  const entries = [];
  for (let p = 0; p + 128 <= dir.length; p += 128) {
    const len = ddv.getUint16(p + 64, true);
    let name = '';
    for (let i = 0; i + 2 < len && i < 64; i += 2) name += String.fromCharCode(ddv.getUint16(p + i, true));
    entries.push({ name, type: dir[p + 66], left: ddv.getUint32(p + 68, true), right: ddv.getUint32(p + 72, true), child: ddv.getUint32(p + 76, true), start: ddv.getUint32(p + 116, true), size: ddv.getUint32(p + 120, true) });
  }
  const root = entries.find((e) => e.type === 5);
  let miniStream = null;
  let miniFat = null;
  const read = (e) => {
    if (!e) return null;
    if (e.size >= miniCutoff || !root) return readChain(e.start).subarray(0, e.size);
    // 작은 스트림은 미니 스트림에서
    miniStream ??= readChain(root.start);
    if (!miniFat) {
      const mf = readChain(miniFatStart);
      const mdv = new DataView(mf.buffer);
      miniFat = new Uint32Array(mf.length / 4);
      for (let i = 0; i < miniFat.length; i++) miniFat[i] = mdv.getUint32(i * 4, true);
    }
    const secs = chain(e.start, miniFat);
    const out = new Uint8Array(secs.length * miniSize);
    secs.forEach((s, i) => out.set(miniStream.subarray(s * miniSize, s * miniSize + miniSize), i * miniSize));
    return out.subarray(0, e.size);
  };
  // 경로 (저장소/스트림) — 디렉터리 트리를 따라
  const paths = new Map();
  const walk = (idx, prefix, depth = 0) => {
    if (idx === 0xffffffff || idx >= entries.length || depth > 64) return;
    const e = entries[idx];
    if (!e) return;
    walk(e.left, prefix, depth + 1);
    walk(e.right, prefix, depth + 1);
    const p = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.type === 2) paths.set(p, e);
    if (e.type === 1) walk(e.child, p, depth + 1);
  };
  if (root) walk(root.child, '');
  return {
    names: [...paths.keys()],
    /** 스트림 읽기: 이름(최상위) 또는 '저장소/스트림' 경로 */
    stream: (name) => read(paths.get(name) ?? entries.find((e) => e.type === 2 && e.name === name)),
  };
}
