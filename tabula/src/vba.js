// .xlsm 매크로(vbaProject.bin) 읽기: OLE 복합 문서(CFB) + MS-OVBA 압축 해제
// 매크로는 실행하지 않고, 코드 보기와 .xlsm 저장 시 보존에만 사용.

const ENDOFCHAIN = 0xfffffffe;
const FREESECT = 0xffffffff;

/** CFB 파일 → { '경로/이름': Uint8Array } (경로는 루트 아래 저장소 이름을 '/' 로 연결) */
export function readCfb(bytes) {
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (dv.getUint32(0, true) !== 0xe011cfd0 || dv.getUint32(4, true) !== 0xe11ab1a1) throw new Error('OLE 파일이 아닙니다');
  const secSize = 1 << dv.getUint16(0x1e, true);
  const miniSize = 1 << dv.getUint16(0x20, true);
  const numFat = dv.getUint32(0x2c, true);
  const firstDir = dv.getUint32(0x30, true);
  const cutoff = dv.getUint32(0x38, true);
  const firstMiniFat = dv.getUint32(0x3c, true);
  let difatSec = dv.getUint32(0x44, true);
  const secOff = (n) => (n + 1) * secSize;

  const fatSectors = [];
  for (let i = 0; i < 109 && fatSectors.length < numFat; i++) {
    const s = dv.getUint32(0x4c + i * 4, true);
    if (s !== FREESECT) fatSectors.push(s);
  }
  for (let guard = 0; difatSec !== ENDOFCHAIN && difatSec !== FREESECT && guard < 10000 && fatSectors.length < numFat; guard++) {
    const base = secOff(difatSec);
    const n = secSize / 4 - 1;
    for (let i = 0; i < n && fatSectors.length < numFat; i++) fatSectors.push(dv.getUint32(base + i * 4, true));
    difatSec = dv.getUint32(base + n * 4, true);
  }
  const fat = new Uint32Array(fatSectors.length * (secSize / 4));
  fatSectors.forEach((s, i) => { for (let j = 0; j < secSize / 4; j++) fat[i * (secSize / 4) + j] = dv.getUint32(secOff(s) + j * 4, true); });

  const chain = (start, table) => {
    const out = [];
    for (let s = start, guard = 0; s !== ENDOFCHAIN && s !== FREESECT && s < table.length && guard < 1e6; guard++) { out.push(s); s = table[s]; }
    return out;
  };
  const readChain = (start, size) => {
    const secs = chain(start, fat);
    const out = new Uint8Array(secs.length * secSize);
    secs.forEach((s, i) => out.set(bytes.subarray(secOff(s), secOff(s) + secSize), i * secSize));
    return size === undefined ? out : out.subarray(0, size);
  };

  const dirBytes = readChain(firstDir);
  const ddv = new DataView(dirBytes.buffer, dirBytes.byteOffset, dirBytes.byteLength);
  const entries = [];
  for (let off = 0; off + 128 <= dirBytes.length; off += 128) {
    const nameLen = ddv.getUint16(off + 64, true);
    let name = '';
    for (let i = 0; i + 2 < nameLen; i += 2) name += String.fromCharCode(ddv.getUint16(off + i, true));
    entries.push({
      name, type: dirBytes[off + 66],
      left: ddv.getUint32(off + 68, true), right: ddv.getUint32(off + 72, true), child: ddv.getUint32(off + 76, true),
      start: ddv.getUint32(off + 116, true), size: ddv.getUint32(off + 120, true),
    });
  }
  const root = entries[0];
  const miniStream = readChain(root.start, root.size);
  const miniFatBytes = firstMiniFat === ENDOFCHAIN ? new Uint8Array(0) : readChain(firstMiniFat);
  const miniFat = new Uint32Array(miniFatBytes.length / 4);
  const mdv = new DataView(miniFatBytes.buffer, miniFatBytes.byteOffset, miniFatBytes.byteLength);
  for (let i = 0; i < miniFat.length; i++) miniFat[i] = mdv.getUint32(i * 4, true);
  const readMini = (start, size) => {
    const secs = chain(start, miniFat);
    const out = new Uint8Array(secs.length * miniSize);
    secs.forEach((s, i) => out.set(miniStream.subarray(s * miniSize, (s + 1) * miniSize), i * miniSize));
    return out.subarray(0, size);
  };

  const files = {};
  const walk = (idx, prefix, seen = new Set()) => {
    if (idx === FREESECT || idx >= entries.length || seen.has(idx)) return;
    seen.add(idx);
    const e = entries[idx];
    walk(e.left, prefix, seen);
    walk(e.right, prefix, seen);
    const path = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.type === 2) files[path] = e.size < cutoff ? readMini(e.start, e.size) : readChain(e.start, e.size);
    else if (e.type === 1) walk(e.child, path, seen);
  };
  walk(root.child, '');
  return files;
}

/** MS-OVBA 압축 해제 */
export function ovbaDecompress(buf) {
  if (!buf.length || buf[0] !== 1) throw new Error('VBA 압축 형식이 아닙니다');
  const out = [];
  let pos = 1;
  while (pos + 1 < buf.length) {
    const header = buf[pos] | (buf[pos + 1] << 8);
    pos += 2;
    const size = (header & 0x0fff) + 3;
    const chunkEnd = Math.min(buf.length, pos + size - 2);
    const chunkStart = out.length;
    if (!(header & 0x8000)) {
      for (let i = 0; i < 4096 && pos < buf.length; i++) out.push(buf[pos++]);
      continue;
    }
    while (pos < chunkEnd) {
      const flags = buf[pos++];
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        if (((flags >> bit) & 1) === 0) {
          out.push(buf[pos++]);
        } else {
          const token = buf[pos] | (buf[pos + 1] << 8);
          pos += 2;
          const d = out.length - chunkStart;
          let bitCount = 4;
          while ((1 << bitCount) < d) bitCount++;
          const lengthMask = 0xffff >> bitCount;
          const length = (token & lengthMask) + 3;
          const offset = (token >> (16 - bitCount)) + 1;
          for (let i = 0; i < length; i++) out.push(out[out.length - offset]);
        }
      }
    }
  }
  return Uint8Array.from(out);
}

const CODEPAGES = { 949: 'euc-kr', 1252: 'windows-1252', 932: 'shift_jis', 936: 'gbk', 950: 'big5', 65001: 'utf-8', 1200: 'utf-16le' };
function decoder(cp) {
  try { return new TextDecoder(CODEPAGES[cp] ?? `windows-${cp}`); } catch { return new TextDecoder('windows-1252'); }
}

/** vbaProject.bin → [{ name, type: '표준 모듈'|'클래스 모듈'|'문서', code }] */
export function extractVbaModules(bytes) {
  const files = readCfb(bytes);
  const dirKey = Object.keys(files).find((k) => /(^|\/)VBA\/dir$/i.test(k));
  if (!dirKey) return [];
  const base = dirKey.slice(0, -3);
  const dir = ovbaDecompress(files[dirKey]);
  const dv = new DataView(dir.buffer, dir.byteOffset, dir.byteLength);
  let pos = 0;
  let codepage = 1252;
  const modules = [];
  let cur = null;
  while (pos + 6 <= dir.length) {
    const id = dv.getUint16(pos, true);
    const size = dv.getUint32(pos + 2, true);
    const data = dir.subarray(pos + 6, pos + 6 + size);
    pos += 6 + size;
    if (id === 0x0009) { pos += 2; continue; } // PROJECTVERSION 은 크기 필드와 실제 길이가 다름
    switch (id) {
      case 0x0003: codepage = data[0] | (data[1] << 8); break;
      case 0x0019: cur = { name: decoder(codepage).decode(data), stream: null, offset: 0, type: '표준 모듈' }; break;
      case 0x0047: if (cur) cur.name = new TextDecoder('utf-16le').decode(data); break;
      case 0x001a: if (cur) cur.stream = decoder(codepage).decode(data); break;
      case 0x0032: if (cur) cur.stream = new TextDecoder('utf-16le').decode(data); break;
      case 0x0031: if (cur) cur.offset = new DataView(data.buffer, data.byteOffset, 4).getUint32(0, true); break;
      case 0x0022: if (cur) cur.type = '클래스 모듈'; break;
      case 0x002b: if (cur) { modules.push(cur); cur = null; } break;
      default:
    }
  }
  const dec = decoder(codepage);
  return modules.map((m) => {
    const key = Object.keys(files).find((k) => k.toLowerCase() === `${base}${m.stream ?? m.name}`.toLowerCase());
    let code = '';
    if (key) {
      try { code = dec.decode(ovbaDecompress(files[key].subarray(m.offset))); } catch { code = '(코드를 읽을 수 없습니다)'; }
    }
    const lines = code.split(/\r?\n/).filter((l) => !/^Attribute VB_/.test(l));
    const isDoc = /^Attribute VB_Base = "0\{00020819|^Attribute VB_Base = "0\{00020820/m.test(code);
    return { name: m.name, type: isDoc ? '문서' : m.type, code: lines.join('\n').replace(/\s+$/, '') };
  });
}

// base64 (브라우저/Node 공용)
export function toBase64(bytes) {
  if (typeof bytes.toBase64 === 'function') return bytes.toBase64(); // 최신 브라우저 · Node 내장 (몇 배 빠름)
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(s);
}
export function fromBase64(b64) {
  const s = atob(b64);
  const out = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) out[i] = s.charCodeAt(i);
  return out;
}
