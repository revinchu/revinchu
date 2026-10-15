// TrueType(glyf) 글꼴 읽기 · 부분 집합 만들기 (DOM 없음) — PDF 에 글꼴을 넣을 때 쓴 글자만 남김
// 글리프 번호는 그대로 두고 쓰지 않는 글리프만 비움 (PDF 는 CID = 글리프 번호, Identity-H) → 다시 번호 매기기 없이 단순하고 안전
const u16 = (b, o) => (b[o] << 8) | b[o + 1];
const i16 = (b, o) => { const v = u16(b, o); return v > 32767 ? v - 65536 : v; };
const u32 = (b, o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
const tagAt = (b, o) => String.fromCharCode(b[o], b[o + 1], b[o + 2], b[o + 3]);

/** 글꼴 모음(.ttc) 에서 글꼴 하나의 시작 위치들 */
export function ttcOffsets(b) { if (tagAt(b, 0) !== 'ttcf') return [0]; const n = u32(b, 8); return Array.from({ length: n }, (_, i) => u32(b, 12 + i * 4)); }
function tables(b, off) {
  const n = u16(b, off + 4); const t = {};
  for (let i = 0; i < n; i++) { const o = off + 12 + i * 16; t[tagAt(b, o)] = { off: u32(b, o + 8), len: u32(b, o + 12) }; }
  return t;
}
/** name 표의 PostScript 이름 (6) · 전체 이름 (4) */
function names(b, t) {
  const out = {};
  const nt = t.name; if (!nt) return out;
  const base = nt.off; const count = u16(b, base + 2); const strOff = base + u16(b, base + 4);
  for (let i = 0; i < count; i++) {
    const r = base + 6 + i * 12; const pid = u16(b, r); const nid = u16(b, r + 6); const len = u16(b, r + 8); const o = strOff + u16(b, r + 10);
    if (nid !== 4 && nid !== 6 && nid !== 1) continue;
    let s = '';
    if (pid === 3 || pid === 0) for (let k = 0; k < len; k += 2) s += String.fromCharCode(u16(b, o + k));
    else for (let k = 0; k < len; k++) s += String.fromCharCode(b[o + k]);
    (out[nid] ??= []).push(s);
  }
  return out;
}

/** sfnt(.ttf/.ttc) → 글꼴 객체 (psName 이 있으면 모음에서 그 글꼴) — glyf 글꼴이 아니면 null */
export function openTtf(bytes, psName = null) {
  for (const off of ttcOffsets(bytes)) {
    const t = tables(bytes, off);
    if (!t.glyf || !t.loca || !t.head || !t.hmtx || !t.cmap) continue;
    const nm = names(bytes, t);
    if (psName && ttcOffsets(bytes).length > 1 && !(nm[6] ?? []).concat(nm[4] ?? []).some((n) => n.replace(/\s/g, '') === psName.replace(/\s/g, ''))) continue;
    return makeFont(bytes, t, nm);
  }
  return null;
}
function makeFont(b, t, nm) {
  const head = t.head.off;
  const upm = u16(b, head + 18);
  const locFmt = i16(b, head + 50);
  const numGlyphs = u16(b, t.maxp.off + 4);
  const numHM = u16(b, t.hhea.off + 34);
  const loca = (g) => (locFmt ? u32(b, t.loca.off + g * 4) : u16(b, t.loca.off + g * 2) * 2);
  const adv = (g) => u16(b, t.hmtx.off + Math.min(g, numHM - 1) * 4);
  // cmap: (3,10) 형식 12 > (3,1)/(0,x) 형식 4
  let map4 = null; let map12 = null;
  const cm = t.cmap.off; const nsub = u16(b, cm + 2);
  for (let i = 0; i < nsub; i++) {
    const pid = u16(b, cm + 4 + i * 8); const eid = u16(b, cm + 6 + i * 8); const so = cm + u32(b, cm + 8 + i * 8);
    const fmt = u16(b, so);
    if (fmt === 12 && (pid === 3 || pid === 0)) map12 = so;
    else if (fmt === 4 && ((pid === 3 && (eid === 1 || eid === 0)) || pid === 0)) map4 ??= so;
  }
  const cache = new Map();
  const gidOf = (cp) => {
    if (cache.has(cp)) return cache.get(cp);
    let g = 0;
    if (map12) {
      const n = u32(b, map12 + 12);
      let lo = 0; let hi = n - 1;
      while (lo <= hi) { const m = (lo + hi) >> 1; const o = map12 + 16 + m * 12; const s = u32(b, o); const e = u32(b, o + 4); if (cp < s) hi = m - 1; else if (cp > e) lo = m + 1; else { g = u32(b, o + 8) + cp - s; break; } }
    } else if (map4 && cp <= 0xffff) {
      const segX2 = u16(b, map4 + 6); const ends = map4 + 14; const starts = ends + segX2 + 2; const deltas = starts + segX2; const ranges = deltas + segX2;
      for (let k = 0; k < segX2 / 2; k++) {
        if (u16(b, ends + k * 2) < cp) continue;
        const s = u16(b, starts + k * 2); if (cp < s) break;
        const ro = u16(b, ranges + k * 2); const d = u16(b, deltas + k * 2);
        if (!ro) g = (cp + d) & 0xffff;
        else { const gi = u16(b, ranges + k * 2 + ro + (cp - s) * 2); g = gi ? (gi + d) & 0xffff : 0; }
        break;
      }
    }
    if (g >= numGlyphs) g = 0;
    cache.set(cp, g);
    return g;
  };
  const os2 = t['OS/2']?.off;
  return {
    bytes: b, t, upm, numGlyphs, locFmt,
    psName: (nm[6]?.[0] ?? nm[4]?.[0] ?? 'Font').replace(/[^\x21-\x7e]/g, '').replace(/[()<>[\]{}/%]/g, '') || 'Font',
    family: nm[1]?.[0] ?? '',
    gidOf,
    /** 글리프 너비 (1000 단위) */
    width: (g) => Math.round(adv(g) * 1000 / upm),
    bbox: [i16(b, head + 36), i16(b, head + 38), i16(b, head + 40), i16(b, head + 42)].map((v) => Math.round(v * 1000 / upm)),
    ascent: Math.round((os2 ? i16(b, os2 + 68) : i16(b, t.hhea.off + 4)) * 1000 / upm),
    descent: Math.round((os2 ? i16(b, os2 + 70) : i16(b, t.hhea.off + 6)) * 1000 / upm),
    capHeight: os2 && t['OS/2'].len >= 90 ? Math.round(i16(b, os2 + 88) * 1000 / upm) : 700,
    fsType: os2 ? u16(b, os2 + 8) : 0,
    glyphRange: (g) => { const a = loca(g); const z = loca(g + 1); return [t.glyf.off + a, Math.max(0, z - a)]; },
  };
}

/** 쓴 글리프만 남긴 .ttf (복합 글리프의 부품까지) */
export function subsetTtf(font, used) {
  const b = font.bytes;
  const keep = new Set([0, ...used]);
  // 복합 글리프 부품
  const stack = [...keep];
  while (stack.length) {
    const g = stack.pop();
    const [o, len] = font.glyphRange(g);
    if (len < 10 || i16(b, o) >= 0) continue;
    let p = o + 10;
    for (;;) {
      const flags = u16(b, p); const comp = u16(b, p + 2);
      if (!keep.has(comp)) { keep.add(comp); stack.push(comp); }
      p += 4 + (flags & 1 ? 4 : 2) + (flags & 8 ? 2 : flags & 0x40 ? 4 : flags & 0x80 ? 8 : 0);
      if (!(flags & 0x20)) break;
    }
  }
  // 새 glyf · loca (long 형식)
  const n = font.numGlyphs;
  let size = 0;
  for (const g of keep) size += (font.glyphRange(g)[1] + 3) & ~3;
  const glyf = new Uint8Array(size);
  const loca = new Uint8Array((n + 1) * 4); const lv = new DataView(loca.buffer);
  let pos = 0;
  for (let g = 0; g < n; g++) {
    lv.setUint32(g * 4, pos);
    if (keep.has(g)) { const [o, len] = font.glyphRange(g); glyf.set(b.subarray(o, o + len), pos); pos += (len + 3) & ~3; }
  }
  lv.setUint32(n * 4, pos);
  const copy = (tag) => (font.t[tag] ? b.slice(font.t[tag].off, font.t[tag].off + font.t[tag].len) : null);
  const head = copy('head'); head[50] = 0; head[51] = 1; head[8] = head[9] = head[10] = head[11] = 0; // indexToLocFormat = long, checkSumAdjustment = 0
  const out = { head, hhea: copy('hhea'), maxp: copy('maxp'), hmtx: copy('hmtx'), loca, glyf: glyf.subarray(0, pos) };
  for (const tag of ['cvt ', 'fpgm', 'prep']) { const v = copy(tag); if (v) out[tag] = v; }
  // sfnt 조립
  const tags = Object.keys(out).sort();
  const hdr = 12 + tags.length * 16;
  let total = hdr; for (const tg of tags) total += (out[tg].length + 3) & ~3;
  const f = new Uint8Array(total); const dv = new DataView(f.buffer);
  dv.setUint32(0, 0x00010000); dv.setUint16(4, tags.length);
  let sr = 1; let es = 0; while (sr * 2 <= tags.length) { sr *= 2; es++; }
  dv.setUint16(6, sr * 16); dv.setUint16(8, es); dv.setUint16(10, tags.length * 16 - sr * 16);
  let off = hdr;
  tags.forEach((tg, i) => {
    const data = out[tg];
    for (let k = 0; k < 4; k++) f[12 + i * 16 + k] = tg.charCodeAt(k);
    let sum = 0; for (let k = 0; k < data.length; k += 4) sum = (sum + ((data[k] << 24) | ((data[k + 1] ?? 0) << 16) | ((data[k + 2] ?? 0) << 8) | (data[k + 3] ?? 0))) >>> 0;
    dv.setUint32(12 + i * 16 + 4, sum); dv.setUint32(12 + i * 16 + 8, off); dv.setUint32(12 + i * 16 + 12, data.length);
    f.set(data, off); off += (data.length + 3) & ~3;
  });
  return f;
}
