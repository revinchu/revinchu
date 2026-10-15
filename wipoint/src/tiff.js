// TIFF → RGBA (DOM 없음) — 브라우저(Chrome · Edge · Firefox)는 TIFF 를 그리지 못하므로 화면용 PNG 로 바꿈
// 지원: 바이트 순서 II/MM, 무압축 · LZW · PackBits · Deflate, 가로 예측(Predictor 2), 띠(strip) · 타일,
//       흑백/회색조 · RGB(A) · 팔레트, 표본 1/2/4/8/16비트, 첫 번째 그림(IFD0)만.
import { inflate } from './zip.js';

export const isTiff = (b) => b && b.length > 8 && ((b[0] === 0x49 && b[1] === 0x49 && b[2] === 42 && b[3] === 0) || (b[0] === 0x4d && b[1] === 0x4d && b[2] === 0 && b[3] === 42));

function lzw(src, expected) {
  const out = new Uint8Array(Math.max(expected, 1));
  let n = 0;
  let dict = [];
  const reset = () => { dict = []; for (let i = 0; i < 256; i++) dict[i] = [i]; dict[256] = null; dict[257] = null; };
  reset();
  let bits = 9;
  let pos = 0;
  let prev = null;
  const read = () => {
    let v = 0;
    for (let i = 0; i < bits; i++) {
      const byte = src[(pos + i) >> 3];
      if (byte === undefined) return 257;
      v = (v << 1) | ((byte >> (7 - ((pos + i) & 7))) & 1);
    }
    pos += bits;
    return v;
  };
  const put = (arr) => { for (const v of arr) { if (n < out.length) out[n++] = v; } };
  for (;;) {
    const code = read();
    if (code === 257) break;
    if (code === 256) { reset(); bits = 9; prev = null; continue; }
    let entry;
    if (code < dict.length && dict[code]) entry = dict[code];
    else if (prev) entry = [...prev, prev[0]];
    else break;
    put(entry);
    if (prev) dict.push([...prev, entry[0]]);
    prev = entry;
    // TIFF LZW 은 한 단계 일찍 비트 수를 늘림 (early change)
    if (dict.length + 1 >= (1 << bits) && bits < 12) bits++;
    if (n >= out.length) break;
  }
  return out.subarray(0, n);
}

function packbits(src, expected) {
  const out = new Uint8Array(expected);
  let i = 0;
  let n = 0;
  while (i < src.length && n < expected) {
    const c = (src[i++] << 24) >> 24;
    if (c >= 0) { for (let k = 0; k <= c && n < expected; k++) out[n++] = src[i++]; } else if (c !== -128) { const v = src[i++]; for (let k = 0; k < 1 - c && n < expected; k++) out[n++] = v; }
  }
  return out;
}

/** TIFF 바이트 → { w, h, rgba } (못 읽으면 null) */
export function decodeTiff(bytes) {
  if (!isTiff(bytes)) return null;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const le = bytes[0] === 0x49;
  const u16 = (o) => dv.getUint16(o, le);
  const u32 = (o) => dv.getUint32(o, le);
  const ifd = u32(4);
  const n = u16(ifd);
  const tags = {};
  const SIZE = { 1: 1, 2: 1, 3: 2, 4: 4, 5: 8, 6: 1, 7: 1, 8: 2, 9: 4, 10: 8, 11: 4, 12: 8, 16: 8 };
  for (let i = 0; i < n; i++) {
    const e = ifd + 2 + i * 12;
    const tag = u16(e); const type = u16(e + 2); const count = u32(e + 4);
    const sz = (SIZE[type] ?? 1) * count;
    const at = sz <= 4 ? e + 8 : u32(e + 8);
    const vals = [];
    for (let k = 0; k < Math.min(count, 1 << 20); k++) {
      const o = at + k * (SIZE[type] ?? 1);
      if (o + (SIZE[type] ?? 1) > bytes.length) break;
      vals.push(type === 3 || type === 8 ? u16(o) : type === 4 || type === 9 || type === 16 ? u32(o) : type === 5 ? u32(o) / (u32(o + 4) || 1) : bytes[o]);
    }
    tags[tag] = vals;
  }
  const w = tags[256]?.[0]; const h = tags[257]?.[0];
  if (!w || !h || w * h > 64e6) return null;
  const bps = tags[258]?.[0] ?? 1;
  const comp = tags[259]?.[0] ?? 1;
  const photo = tags[262]?.[0] ?? 1;
  const spp = tags[277]?.[0] ?? 1;
  const pred = tags[317]?.[0] ?? 1;
  const planar = tags[284]?.[0] ?? 1;
  if (planar !== 1 || ![1, 2, 4, 8, 16].includes(bps)) return null;
  const map = tags[320];
  const tiled = !!tags[324];
  const tw = tiled ? tags[322][0] : w;
  const th = tiled ? tags[323][0] : tags[278]?.[0] ?? h;
  const offs = tiled ? tags[324] : tags[273] ?? [];
  const cnts = tiled ? tags[325] : tags[279] ?? [];
  const rowBytes = Math.ceil((tw * spp * bps) / 8);
  const rgba = new Uint8ClampedArray(w * h * 4);
  const across = Math.ceil(w / tw);
  const sample = (row, i) => {
    if (bps === 8) return row[i];
    if (bps === 16) return row[i * 2 + (le ? 1 : 0)];
    const bit = i * bps;
    const v = (row[bit >> 3] >> (8 - bps - (bit & 7))) & ((1 << bps) - 1);
    return Math.round((v * 255) / ((1 << bps) - 1));
  };
  for (let s = 0; s < offs.length; s++) {
    let data = bytes.subarray(offs[s], offs[s] + (cnts[s] ?? 0));
    const want = rowBytes * th;
    try {
      if (comp === 5) data = lzw(data, want);
      else if (comp === 8 || comp === 32946) data = inflate(data, want);
      else if (comp === 32773) data = packbits(data, want);
      else if (comp !== 1) return null;
    } catch { continue; }
    const x0 = tiled ? (s % across) * tw : 0;
    const y0 = tiled ? Math.floor(s / across) * th : s * th;
    for (let r = 0; r < th && y0 + r < h; r++) {
      const row = data.subarray(r * rowBytes, (r + 1) * rowBytes);
      if (pred === 2 && bps === 8) for (let i = spp; i < tw * spp; i++) row[i] = (row[i] + row[i - spp]) & 255;
      for (let c = 0; c < tw && x0 + c < w; c++) {
        const o = ((y0 + r) * w + x0 + c) * 4;
        if (photo === 2) { rgba[o] = sample(row, c * spp); rgba[o + 1] = sample(row, c * spp + 1); rgba[o + 2] = sample(row, c * spp + 2); rgba[o + 3] = spp > 3 ? sample(row, c * spp + 3) : 255; } else if (photo === 3 && map) {
          const k = bps === 8 ? row[c] : sample(row, c) * ((1 << bps) - 1) / 255;
          const m = map.length / 3;
          rgba[o] = map[k] >> 8; rgba[o + 1] = map[m + k] >> 8; rgba[o + 2] = map[2 * m + k] >> 8; rgba[o + 3] = 255;
        } else {
          let g = sample(row, c * spp);
          if (photo === 0) g = 255 - g;
          rgba[o] = rgba[o + 1] = rgba[o + 2] = g; rgba[o + 3] = spp > 1 ? sample(row, c * spp + 1) : 255;
        }
      }
    }
  }
  return { w, h, rgba };
}
