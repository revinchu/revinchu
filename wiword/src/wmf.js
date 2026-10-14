// WMF (Windows 메타파일, 16비트) → SVG 변환 — DOM 없음
// PowerPoint 97–2003 파일 · 예전 클립 아트 · OLE 개체 미리 보기는 WMF 가 많은데 브라우저는 그리지 못함.
// 자주 쓰는 GDI 레코드(선 · 다각형 · 사각형 · 타원 · 호 · 글자 · 비트맵 · 잘라내기 · 창 좌표)만 SVG 로 옮김.
import { symbolTable, symbolText } from './symfont.js';
import { dibToBmp, arcPath } from './emf.js';

const hex2 = (n) => n.toString(16).padStart(2, '0');
const escXml = (s) => s.replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
const num = (v) => (Math.abs(v) < 1e-9 ? '0' : String(Math.round(v * 100) / 100));
function b64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** WMF 인지 (자리 지정 머리말 9AC6CDD7 또는 META 머리말) */
export function isWmf(bytes) {
  if (!bytes || bytes.length < 18) return false;
  if (bytes[0] === 0xd7 && bytes[1] === 0xcd && bytes[2] === 0xc6 && bytes[3] === 0x9a) return true;
  return (bytes[0] === 1 || bytes[0] === 2) && bytes[1] === 0 && bytes[2] === 9 && bytes[3] === 0 && (bytes[4] === 0 || bytes[4] === 3) && bytes[5] === 3;
}

/** 글자 바이트 → 문자열 (한글 글꼴 charset 129 는 CP949 로) */
function decodeText(bytes, charset) {
  if (charset === 2) return String.fromCharCode(...bytes); // SYMBOL_CHARSET: 바이트 = 글자 번호
  const label = { 129: 'euc-kr', 128: 'shift_jis', 134: 'gbk', 136: 'big5', 161: 'windows-1253', 162: 'windows-1254', 177: 'windows-1255', 178: 'windows-1256', 186: 'windows-1257', 204: 'windows-1251', 222: 'windows-874', 238: 'windows-1250' }[charset] ?? 'windows-1252';
  try { return new TextDecoder(label).decode(bytes); } catch { return String.fromCharCode(...bytes); }
}

/** DIB (BITMAPINFO + 비트, 한 덩어리) → [bmi, bits] */
function splitDib(dib) {
  if (dib.length < 12) return null;
  const dv = new DataView(dib.buffer, dib.byteOffset, dib.byteLength);
  const hs = dv.getUint32(0, true);
  if (hs === 12) { // BITMAPCOREHEADER
    const bpp = dv.getUint16(10, true);
    const pal = bpp <= 8 ? (1 << bpp) * 3 : 0;
    return [dib.subarray(0, 12 + pal), dib.subarray(12 + pal)];
  }
  if (hs < 40 || dib.length < hs) return null;
  const bpp = dv.getUint16(14, true);
  const comp = dv.getUint32(16, true);
  const used = dv.getUint32(32, true);
  let pal = (used || (bpp <= 8 ? 1 << bpp : 0)) * 4;
  if (comp === 3 && hs === 40) pal += 12; // BI_BITFIELDS 색 마스크
  const n = Math.min(dib.length, hs + pal);
  return [dib.subarray(0, n), dib.subarray(n)];
}

/** WMF 바이트 → SVG 문자열 (못 그리면 null). 그림 틀에 맞게 늘어남 */
export function wmfToSvg(bytes) {
  if (!isWmf(bytes)) return null;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const i16 = (o) => dv.getInt16(o, true);
  const u16 = (o) => dv.getUint16(o, true);
  const u32 = (o) => dv.getUint32(o, true);
  const color = (o) => `#${hex2(dv.getUint8(o))}${hex2(dv.getUint8(o + 1))}${hex2(dv.getUint8(o + 2))}`;
  let o = 0;
  let base = null; // 출력 좌표계 [x, y, w, h]
  if (u32(0) === 0x9ac6cdd7) {
    const l = i16(6), t = i16(8), r = i16(10), b = i16(12);
    if (r !== l && b !== t) base = [Math.min(l, r), Math.min(t, b), Math.abs(r - l), Math.abs(b - t)];
    o = 22;
  }
  const hdrWords = u16(o + 2);
  o += hdrWords * 2;
  const objs = [];
  const newObj = (ob) => { let i = 0; while (objs[i]) i++; objs[i] = ob; };
  let st = { winOrg: [0, 0], winExt: null, pen: { style: 0, width: 1, color: '#000000' }, brush: { style: 0, color: '#ffffff' }, font: { height: -12, weight: 400, face: '', esc: 0, charset: 0 }, textColor: '#000000', bkColor: '#ffffff', bkMode: 2, align: 0, fillRule: 'evenodd', clip: null, cur: [0, 0] };
  const stack = [];
  const out = [];
  const clips = [];
  let clipKey = '';
  let clipId = null;
  let drawn = 0;
  // 논리 좌표 → 출력 좌표 (처음 정한 창 = 출력 좌표계; 나중에 창이 바뀌면 비율로 옮김)
  const pt = (x, y) => {
    const we = st.winExt;
    if (!we || !base) return [x, y];
    return [base[0] + ((x - st.winOrg[0]) * base[2]) / (we[0] || 1), base[1] + ((y - st.winOrg[1]) * base[3]) / (we[1] || 1)];
  };
  const sx = () => (st.winExt && base ? Math.abs(base[2] / (st.winExt[0] || 1)) : 1);
  const sy = () => (st.winExt && base ? Math.abs(base[3] / (st.winExt[1] || 1)) : 1);
  const penAttrs = () => {
    const p = st.pen;
    if (!p || p.style === 5) return 'stroke="none"';
    const w = Math.max((p.width || 1) * sx(), sx() * 0.5);
    const dash = { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] }[p.style];
    return `stroke="${p.color}" stroke-width="${num(w)}"${dash ? ` stroke-dasharray="${dash.map((d) => num(d * w * 2)).join(' ')}"` : ''} stroke-linejoin="round" stroke-linecap="round"`;
  };
  const brushAttrs = () => (!st.brush || st.brush.style === 1 ? 'fill="none"' : `fill="${st.brush.color}"`);
  const clipAttr = () => {
    const c = st.clip;
    const key = c ? c.map(num).join(',') : '';
    if (key !== clipKey) {
      clipKey = key;
      if (!c) clipId = null;
      else { clipId = `w${clips.length}`; clips.push(`<clipPath id="${clipId}"><rect x="${num(c[0])}" y="${num(c[1])}" width="${num(Math.max(0, c[2] - c[0]))}" height="${num(Math.max(0, c[3] - c[1]))}"/></clipPath>`); }
    }
    return clipId ? ` clip-path="url(#${clipId})"` : '';
  };
  const emit = (el) => { out.push(el.replace(/^<(\w+)/, (m) => m + clipAttr())); drawn++; };
  const shape = (d, fill = true) => {
    const f = fill ? brushAttrs() : 'fill="none"';
    const s = penAttrs();
    if (f === 'fill="none"' && s === 'stroke="none"') return;
    emit(`<path d="${d}" ${f} ${s} fill-rule="${st.fillRule}"/>`);
  };
  const polyD = (pts, close) => pts.map((p, i) => `${i ? 'L' : 'M'}${num(p[0])} ${num(p[1])}`).join('') + (close ? 'Z' : '');
  const pts16 = (q, n) => { const a = []; for (let i = 0; i < n; i++) a.push(pt(i16(q + i * 4), i16(q + i * 4 + 2))); return a; };
  const rect = (q) => { const a = pt(i16(q + 6), i16(q + 4)); const b = pt(i16(q + 2), i16(q)); return [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])]; };
  const image = (x, y, w, h, dib) => {
    const parts = splitDib(dib);
    if (!parts || !parts[1].length) return;
    const p1 = pt(x, y); const p2 = pt(x + w, y + h);
    emit(`<image x="${num(Math.min(p1[0], p2[0]))}" y="${num(Math.min(p1[1], p2[1]))}" width="${num(Math.abs(p2[0] - p1[0]))}" height="${num(Math.abs(p2[1] - p1[1]))}" preserveAspectRatio="none" href="data:image/bmp;base64,${b64(dibToBmp(parts[0], parts[1]))}"/>`);
  };
  const fillRect = (r, c) => emit(`<rect x="${num(r[0])}" y="${num(r[1])}" width="${num(r[2] - r[0])}" height="${num(r[3] - r[1])}" fill="${c}"/>`);
  const text = (x, y, str, dx) => {
    if (st.align & 1) [x, y] = st.cur;
    if (!str.trim()) return;
    const f = st.font;
    const size = Math.max(1, Math.abs(f.height || 12) * (f.height < 0 ? 1 : 0.8) * sy());
    const p0 = pt(x, y);
    const ha = st.align & 6, va = st.align & 24;
    const anchor = ha === 6 ? 'middle' : ha === 2 ? 'end' : 'start';
    const bl = va === 24 ? '' : va === 8 ? ' dominant-baseline="text-after-edge"' : ' dominant-baseline="text-before-edge"';
    const angle = -(f.esc || 0) / 10;
    let xs = '';
    if (dx?.length === [...str].length && anchor === 'start' && !angle) {
      let acc = 0;
      xs = dx.map((d) => { const v = num(p0[0] + acc * sx()); acc += d; return v; }).join(' ');
    }
    const family = `${f.face ? `'${escXml(f.face).replace(/'/g, '')}', ` : ''}'맑은 고딕', 'Malgun Gothic', sans-serif`;
    const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
    const rot = angle ? ` transform="rotate(${num(angle)} ${num(p0[0])} ${num(p0[1])})"` : '';
    // Symbol · Wingdings 글꼴 글자 (수식 편집기 그림의 τ 등) → 유니코드
    const shown = symbolTable(f.face) ? symbolText(f.face, str) : str;
    emit(`<text x="${xs || num(p0[0])}" y="${num(p0[1])}" font-family="${family}" font-size="${num(size)}"${f.weight >= 600 ? ' font-weight="bold"' : ''}${f.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${st.textColor}" text-anchor="${anchor}"${bl}${rot} xml:space="preserve">${escXml(shown)}</text>`);
    if (st.align & 1) st.cur = [x + (dx ? dx.reduce((a, b) => a + b, 0) : 0), y];
  };
  const end = bytes.length;
  while (o + 6 <= end) {
    const size = u32(o) * 2;
    const fn = u16(o + 4);
    if (size < 6 || o + size > end) break;
    const p = o + 6;
    switch (fn) {
      case 0x0000: o = end; continue; // EOF
      case 0x0201: st.bkColor = color(p); break;
      case 0x0102: st.bkMode = u16(p); break;
      case 0x0106: st.fillRule = u16(p) === 2 ? 'nonzero' : 'evenodd'; break;
      case 0x0209: st.textColor = color(p); break;
      case 0x012e: st.align = u16(p); break;
      case 0x020b: st.winOrg = [i16(p + 2), i16(p)]; break; // SETWINDOWORG (y, x)
      case 0x020c: { // SETWINDOWEXT
        st.winExt = [i16(p + 2), i16(p)];
        if (!base) base = [st.winOrg[0] + Math.min(0, st.winExt[0]), st.winOrg[1] + Math.min(0, st.winExt[1]), Math.abs(st.winExt[0]) || 1, Math.abs(st.winExt[1]) || 1];
        break;
      }
      case 0x020f: st.winOrg = [st.winOrg[0] + i16(p + 2), st.winOrg[1] + i16(p)]; break; // OFFSETWINDOWORG
      case 0x0410: if (st.winExt) st.winExt = [st.winExt[0] * i16(p + 6) / (i16(p + 4) || 1), st.winExt[1] * i16(p + 2) / (i16(p) || 1)]; break; // SCALEWINDOWEXT
      case 0x0214: st.cur = [i16(p + 2), i16(p)]; break; // MOVETO
      case 0x0213: { // LINETO
        const a = pt(...st.cur);
        st.cur = [i16(p + 2), i16(p)];
        const b = pt(...st.cur);
        shape(`M${num(a[0])} ${num(a[1])}L${num(b[0])} ${num(b[1])}`, false);
        break;
      }
      case 0x001e: stack.push({ ...st }); break; // SAVEDC
      case 0x0127: { // RESTOREDC
        const k = i16(p);
        const idx = k < 0 ? stack.length + k : k - 1;
        if (idx >= 0 && idx < stack.length) { st = stack[idx]; stack.length = idx; }
        break;
      }
      case 0x0416: { // INTERSECTCLIPRECT
        const r = rect(p);
        st.clip = st.clip ? [Math.max(st.clip[0], r[0]), Math.max(st.clip[1], r[1]), Math.min(st.clip[2], r[2]), Math.min(st.clip[3], r[3])] : r;
        break;
      }
      case 0x012c: if (u16(p) === 0 || !objs[u16(p)]) st.clip = null; break; // SELECTCLIPREGION (0 = 없앰)
      case 0x02fa: newObj({ kind: 'pen', style: u16(p) & 0xf, width: Math.max(1, i16(p + 2)), color: color(p + 6) }); break; // CREATEPENINDIRECT
      case 0x02fc: { const s = u16(p); newObj({ kind: 'brush', style: s === 1 ? 1 : 0, color: color(p + 2) }); break; } // CREATEBRUSHINDIRECT (무늬 = 같은 색)
      case 0x02fb: { // CREATEFONTINDIRECT
        let face = '';
        const fb = [];
        for (let i = p + 18; i < Math.min(o + size, p + 50); i++) { if (!bytes[i]) break; fb.push(bytes[i]); }
        face = decodeText(new Uint8Array(fb), bytes[p + 13]);
        newObj({ kind: 'font', height: i16(p), esc: i16(p + 4), weight: i16(p + 8), italic: !!bytes[p + 10], underline: !!bytes[p + 11], strike: !!bytes[p + 12], charset: bytes[p + 13], face });
        break;
      }
      case 0x00f7: case 0x01f9: case 0x06ff: newObj({ kind: 'other' }); break; // 팔레트 · 무늬 붓 · 영역
      case 0x0142: newObj({ kind: 'brush', style: 0, color: '#808080' }); break; // DIBCREATEPATTERNBRUSH
      case 0x012d: { // SELECTOBJECT
        const ob = objs[u16(p)];
        if (ob?.kind === 'pen') st.pen = ob;
        else if (ob?.kind === 'brush') st.brush = ob;
        else if (ob?.kind === 'font') st.font = ob;
        break;
      }
      case 0x01f0: objs[u16(p)] = undefined; break; // DELETEOBJECT
      case 0x041b: { const r = rect(p); shape(`M${num(r[0])} ${num(r[1])}H${num(r[2])}V${num(r[3])}H${num(r[0])}Z`); break; } // RECTANGLE
      case 0x061c: { // ROUNDRECT (h, w, b, r, t, l)
        const r = rect(p + 4);
        const rx = Math.min((r[2] - r[0]) / 2, (i16(p + 2) * sx()) / 2), ry = Math.min((r[3] - r[1]) / 2, (i16(p) * sy()) / 2);
        shape(`M${num(r[0] + rx)} ${num(r[1])}H${num(r[2] - rx)}A${num(rx)} ${num(ry)} 0 0 1 ${num(r[2])} ${num(r[1] + ry)}V${num(r[3] - ry)}A${num(rx)} ${num(ry)} 0 0 1 ${num(r[2] - rx)} ${num(r[3])}H${num(r[0] + rx)}A${num(rx)} ${num(ry)} 0 0 1 ${num(r[0])} ${num(r[3] - ry)}V${num(r[1] + ry)}A${num(rx)} ${num(ry)} 0 0 1 ${num(r[0] + rx)} ${num(r[1])}Z`);
        break;
      }
      case 0x0418: { // ELLIPSE
        const r = rect(p);
        const rx = (r[2] - r[0]) / 2, ry = (r[3] - r[1]) / 2;
        shape(`M${num(r[0])} ${num(r[1] + ry)}A${num(rx)} ${num(ry)} 0 1 0 ${num(r[2])} ${num(r[1] + ry)}A${num(rx)} ${num(ry)} 0 1 0 ${num(r[0])} ${num(r[1] + ry)}Z`);
        break;
      }
      case 0x0817: case 0x081a: case 0x0830: { // ARC / PIE / CHORD (yEnd, xEnd, yStart, xStart, b, r, t, l)
        const r = rect(p + 8);
        const s0 = pt(i16(p + 6), i16(p + 4)); const e0 = pt(i16(p + 2), i16(p));
        const d = arcPath(r[0], r[1], r[2], r[3], s0[0], s0[1], e0[0], e0[1], fn === 0x0817 ? 'arc' : fn === 0x081a ? 'pie' : 'chord');
        if (d) shape(d, fn !== 0x0817);
        break;
      }
      case 0x0324: case 0x0325: { // POLYGON / POLYLINE
        const n = i16(p);
        if (n > 0) shape(polyD(pts16(p + 2, n), fn === 0x0324), fn === 0x0324);
        break;
      }
      case 0x0538: { // POLYPOLYGON
        const polys = u16(p);
        let q = p + 2 + polys * 2;
        let d = '';
        for (let i = 0; i < polys; i++) { const c = u16(p + 2 + i * 2); d += polyD(pts16(q, c), true); q += c * 4; }
        if (d) shape(d);
        break;
      }
      case 0x041f: { const a = pt(i16(p + 6), i16(p + 4)); emit(`<rect x="${num(a[0])}" y="${num(a[1])}" width="${num(sx())}" height="${num(sy())}" fill="${color(p)}"/>`); break; } // SETPIXEL
      case 0x0521: { // TEXTOUT (len, str, y, x)
        const n = i16(p);
        const sp = p + 2 + ((n + 1) & ~1);
        text(i16(sp + 2), i16(sp), decodeText(bytes.subarray(p + 2, p + 2 + n), st.font.charset), null);
        break;
      }
      case 0x0a32: { // EXTTEXTOUT (y, x, len, opts, [rect], str, [dx])
        const y = i16(p), x = i16(p + 2), n = i16(p + 4), opts = u16(p + 6);
        let q = p + 8;
        if (opts & 0x0006) {
          // 이 사각형만 (왼쪽, 위, 오른쪽, 아래) 순서
          const a = pt(i16(q), i16(q + 2)); const b = pt(i16(q + 4), i16(q + 6));
          if (opts & 0x0002) fillRect([Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])], st.bkColor); // ETO_OPAQUE
          q += 8;
        }
        const str = decodeText(bytes.subarray(q, q + n), st.font.charset);
        q += (n + 1) & ~1;
        let dx = null;
        if (q + n * 2 <= o + size && [...str].length === n) { dx = []; for (let i = 0; i < n; i++) dx.push(i16(q + i * 2)); }
        text(x, y, str, dx);
        break;
      }
      case 0x0f43: { // STRETCHDIB (rop, usage, sh, sw, sy, sx, dh, dw, dy, dx, DIB)
        image(i16(p + 20), i16(p + 18), i16(p + 16), i16(p + 14), bytes.subarray(p + 22, o + size));
        break;
      }
      case 0x0b41: { // DIBSTRETCHBLT (rop, sh, sw, sy, sx, dh, dw, dy, dx, DIB)
        if (size > 26 + 6) image(i16(p + 18), i16(p + 16), i16(p + 14), i16(p + 12), bytes.subarray(p + 20, o + size));
        else if (st.brush?.style !== 1) { const a = pt(i16(p + 20), i16(p + 18)); const b = pt(i16(p + 20) + i16(p + 16), i16(p + 18) + i16(p + 14)); fillRect([Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])], st.brush.color); }
        break;
      }
      case 0x0940: { // DIBBITBLT (rop, sy, sx, h, w, dy, dx, DIB)
        if (size > 22 + 6) image(i16(p + 14), i16(p + 12), i16(p + 10), i16(p + 8), bytes.subarray(p + 16, o + size));
        else {
          const rop = u32(p);
          const a = pt(i16(p + 16), i16(p + 14)); const b = pt(i16(p + 16) + i16(p + 12), i16(p + 14) + i16(p + 10));
          const r = [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])];
          if (rop === 0x00f00021 && st.brush?.style !== 1) fillRect(r, st.brush.color);
          else if (rop === 0x00ff0062) fillRect(r, '#ffffff');
          else if (rop === 0x00000042) fillRect(r, '#000000');
        }
        break;
      }
      case 0x061d: { // PATBLT (rop, h, w, y, x)
        const a = pt(i16(p + 10), i16(p + 8)); const b = pt(i16(p + 10) + i16(p + 6), i16(p + 8) + i16(p + 4));
        const rop = u32(p);
        const c = rop === 0x00ff0062 ? '#ffffff' : rop === 0x00000042 ? '#000000' : st.brush?.style !== 1 ? st.brush.color : null;
        if (c) fillRect([Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])], c);
        break;
      }
      default: break;
    }
    o += size;
  }
  if (!drawn) return null;
  const [vx, vy, vw, vh] = base ?? [0, 0, 1000, 1000];
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${num(vx)} ${num(vy)} ${num(vw)} ${num(vh)}" width="${num(vw)}" height="${num(vh)}" preserveAspectRatio="none">${clips.length ? `<defs>${clips.join('')}</defs>` : ''}${out.join('')}</svg>`;
}
