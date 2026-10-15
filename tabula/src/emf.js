// EMF (Windows 확장 메타파일) → SVG 변환
// 엑셀 파일에 붙여 넣은 차트 · 도식은 EMF 로 들어오는 경우가 많은데 브라우저는 EMF 를 그리지 못함.
// 자주 쓰는 GDI 레코드(선 · 다각형 · 경로 · 글자 · 비트맵 · 잘라내기 · 좌표 변환)만 SVG 로 옮김. EMF+ 주석은 무시(GDI 대체 레코드를 그림).

const STOCK = {
  0: { kind: 'brush', style: 0, color: '#ffffff' },
  1: { kind: 'brush', style: 0, color: '#c0c0c0' },
  2: { kind: 'brush', style: 0, color: '#808080' },
  3: { kind: 'brush', style: 0, color: '#404040' },
  4: { kind: 'brush', style: 0, color: '#000000' },
  5: { kind: 'brush', style: 1 },
  6: { kind: 'pen', style: 0, width: 0, color: '#ffffff' },
  7: { kind: 'pen', style: 0, width: 0, color: '#000000' },
  8: { kind: 'pen', style: 5 },
};
const DEFAULT_FONT = { kind: 'font', height: -12, weight: 400, italic: false, underline: false, strike: false, esc: 0, face: '' };

const hex2 = (n) => n.toString(16).padStart(2, '0');
const colorRef = (dv, o) => `#${hex2(dv.getUint8(o))}${hex2(dv.getUint8(o + 1))}${hex2(dv.getUint8(o + 2))}`;
const escXml = (s) => s.replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[ch]);
const num = (v) => (Math.abs(v) < 1e-9 ? '0' : String(Math.round(v * 100) / 100));
const IDENT = [1, 0, 0, 1, 0, 0];
// 점 [x y 1] × M (M = [a b; c d; e f])
const mul = (A, B) => [
  A[0] * B[0] + A[1] * B[2], A[0] * B[1] + A[1] * B[3],
  A[2] * B[0] + A[3] * B[2], A[2] * B[1] + A[3] * B[3],
  A[4] * B[0] + A[5] * B[2] + B[4], A[4] * B[1] + A[5] * B[3] + B[5],
];

/** DIB(BITMAPINFO + 비트) → BMP 파일 바이트 */
export function dibToBmp(bmi, bits) {
  const size = 14 + bmi.length + bits.length;
  const out = new Uint8Array(size);
  const off = 14 + bmi.length;
  out.set([0x42, 0x4d, size & 255, (size >> 8) & 255, (size >> 16) & 255, (size >>> 24) & 255, 0, 0, 0, 0, off & 255, (off >> 8) & 255, (off >> 16) & 255, 0]);
  out.set(bmi, 14);
  out.set(bits, off);
  return out;
}

function b64(bytes) {
  let s = '';
  for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
  return btoa(s);
}

/** EMF 인지 (첫 레코드가 EMR_HEADER, ' EMF' 서명) */
export function isEmf(bytes) {
  return bytes && bytes.length > 88 && bytes[0] === 1 && bytes[1] === 0 && bytes[2] === 0 && bytes[3] === 0
    && bytes[40] === 0x20 && bytes[41] === 0x45 && bytes[42] === 0x4d && bytes[43] === 0x46;
}

/**
 * EMF 바이트 → SVG 문자열 (그릴 것이 없거나 EMF 가 아니면 null)
 * 반환 SVG 는 그림 틀에 맞게 늘어남 (preserveAspectRatio="none")
 */
export function emfToSvg(bytes) {
  if (!isEmf(bytes)) return null;
  const dv = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const i32 = (o) => dv.getInt32(o, true);
  const u32 = (o) => dv.getUint32(o, true);
  const i16 = (o) => dv.getInt16(o, true);
  const f32 = (o) => dv.getFloat32(o, true);

  // 머리말: rclBounds (장치 픽셀)
  const bx1 = i32(8), by1 = i32(12), bx2 = i32(16), by2 = i32(20);
  const vw = Math.max(1, bx2 - bx1 + 1), vh = Math.max(1, by2 - by1 + 1);

  const objs = [];
  let st = {
    xf: IDENT, mapMode: 1, winOrg: [0, 0], winExt: [1, 1], vpOrg: [0, 0], vpExt: [1, 1],
    pen: STOCK[7], brush: STOCK[0], font: DEFAULT_FONT, textColor: '#000000', bkColor: '#ffffff', bkMode: 2,
    align: 0, fillRule: 'evenodd', clip: null, cur: [0, 0],
  };
  const stack = [];
  const out = [];
  let drawn = 0;
  let path = null; // BEGINPATH ~ ENDPATH 사이의 경로 (장치 좌표 문자열 조각)
  let lastPath = null;
  const clips = [];
  const clipIds = new Map();
  let clipKey = '', clipId = null;

  // 논리 좌표 → 장치 좌표
  const scaleXY = () => {
    if (st.mapMode === 1) return [1, 1];
    return [st.vpExt[0] / (st.winExt[0] || 1), st.vpExt[1] / (st.winExt[1] || 1)];
  };
  const pt = (x, y) => {
    const m = st.xf;
    const px = x * m[0] + y * m[2] + m[4];
    const py = x * m[1] + y * m[3] + m[5];
    const [sx, sy] = scaleXY();
    return [(px - st.winOrg[0]) * sx + st.vpOrg[0], (py - st.winOrg[1]) * sy + st.vpOrg[1]];
  };
  const lenScale = () => {
    const [sx, sy] = scaleXY();
    const m = st.xf;
    return Math.sqrt(Math.abs((m[0] * m[3] - m[1] * m[2]) * sx * sy)) || 1;
  };
  const yScale = () => {
    const [, sy] = scaleXY();
    return Math.hypot(st.xf[2], st.xf[3]) * Math.abs(sy) || 1;
  };

  const penAttrs = () => {
    const p = st.pen;
    if (!p || p.style === 5) return 'stroke="none"';
    const w = p.geometric ? Math.max(p.width * lenScale(), 0.5) : Math.max(p.width || 1, 1);
    let s = `stroke="${p.color}" stroke-width="${num(w)}"`;
    const dash = { 1: [3, 1], 2: [1, 1], 3: [3, 1, 1, 1], 4: [3, 1, 1, 1, 1, 1] }[p.style];
    if (dash) s += ` stroke-dasharray="${dash.map((d) => num(d * Math.max(w, 1))).join(' ')}"`;
    s += ` stroke-linecap="${p.cap ?? 'round'}" stroke-linejoin="${p.join ?? 'round'}"`;
    return s;
  };
  const brushAttrs = () => {
    const b = st.brush;
    if (!b || b.style === 1) return 'fill="none"';
    return `fill="${b.color}"`;
  };
  const clipAttr = () => {
    const c = st.clip;
    const key = c ? c.map(num).join(',') : '';
    if (key !== clipKey) {
      clipKey = key;
      if (!c) clipId = null;
      else if (clipIds.has(key)) clipId = clipIds.get(key);
      else {
        clipId = `c${clips.length}`;
        clipIds.set(key, clipId);
        clips.push(`<clipPath id="${clipId}"><rect x="${num(c[0])}" y="${num(c[1])}" width="${num(Math.max(0, c[2] - c[0]))}" height="${num(Math.max(0, c[3] - c[1]))}"/></clipPath>`);
      }
    }
    return clipId ? ` clip-path="url(#${clipId})"` : '';
  };
  const emit = (el) => { out.push(el.replace(/^<(\w+)/, (m) => m + clipAttr())); drawn++; };
  const shape = (d, { fill = true, stroke = true } = {}) => {
    if (path) { path.push(d); return; }
    const f = fill ? brushAttrs() : 'fill="none"';
    const s = stroke ? penAttrs() : 'stroke="none"';
    if (f === 'fill="none"' && s === 'stroke="none"') return;
    emit(`<path d="${d}" ${f} ${s} fill-rule="${st.fillRule}"/>`);
  };
  const polyD = (pts, close) => {
    let d = '';
    for (let i = 0; i < pts.length; i++) d += `${i ? 'L' : 'M'}${num(pts[i][0])} ${num(pts[i][1])}`;
    return close ? d + 'Z' : d;
  };
  const readPts16 = (o, n) => {
    const a = new Array(n);
    for (let i = 0; i < n; i++) a[i] = pt(i16(o + i * 4), i16(o + i * 4 + 2));
    return a;
  };
  const readPts32 = (o, n) => {
    const a = new Array(n);
    for (let i = 0; i < n; i++) a[i] = pt(i32(o + i * 8), i32(o + i * 8 + 4));
    return a;
  };
  const rectPts = (o) => [pt(i32(o), i32(o + 4)), pt(i32(o + 8), i32(o + 12))];
  const bezD = (start, pts, moveFirst) => {
    let d = moveFirst ? `M${num(start[0])} ${num(start[1])}` : '';
    for (let i = 0; i + 2 < pts.length; i += 3) d += `C${pts.slice(i, i + 3).map((p) => `${num(p[0])} ${num(p[1])}`).join(' ')}`;
    return d;
  };
  const image = (x, y, w, h, bmi, bits) => {
    if (!bmi.length || !bits.length || !(w && h)) return;
    const bmp = dibToBmp(bmi, bits);
    const [p1, p2] = [pt(x, y), pt(x + w, y + h)];
    const X = Math.min(p1[0], p2[0]), Y = Math.min(p1[1], p2[1]);
    emit(`<image x="${num(X)}" y="${num(Y)}" width="${num(Math.abs(p2[0] - p1[0]))}" height="${num(Math.abs(p2[1] - p1[1]))}" preserveAspectRatio="none" href="data:image/bmp;base64,${b64(bmp)}"/>`);
  };
  const fillRect = (x, y, w, h, color) => {
    const [p1, p2] = [pt(x, y), pt(x + w, y + h)];
    emit(`<rect x="${num(Math.min(p1[0], p2[0]))}" y="${num(Math.min(p1[1], p2[1]))}" width="${num(Math.abs(p2[0] - p1[0]))}" height="${num(Math.abs(p2[1] - p1[1]))}" fill="${color}"/>`);
  };
  const text = (o, wide) => {
    // EMREXTTEXTOUT: rclBounds 16, iGraphicsMode, exScale, eyScale, EMRTEXT
    const t = o + 36;
    let x = i32(t), y = i32(t + 4);
    const n = u32(t + 8), offStr = u32(t + 12), opts = u32(t + 16), offDx = u32(t + 36);
    if (!n) return;
    let s = '';
    for (let i = 0; i < n; i++) s += wide ? String.fromCharCode(dv.getUint16(o + offStr + i * 2, true)) : String.fromCharCode(dv.getUint8(o + offStr + i));
    if (st.align & 1) [x, y] = st.cur; // TA_UPDATECP
    const f = st.font;
    // ETO_OPAQUE: 배경색으로 사각형 채우기
    if (opts & 2) {
      const r = [i32(t + 20), i32(t + 24), i32(t + 28), i32(t + 32)];
      if (r[2] > r[0] && r[3] > r[1]) fillRect(r[0], r[1], r[2] - r[0], r[3] - r[1], st.bkColor);
    }
    if (!s.trim()) return;
    const dx = [];
    if (offDx) {
      const step = opts & 0x2000 ? 8 : 4;
      for (let i = 0; i < n; i++) dx.push(i32(o + offDx + i * step));
    }
    const size = Math.max(1, (f.height < 0 ? -f.height : f.height * 0.8) * yScale());
    const p0 = pt(x, y);
    const angle = -(f.esc || 0) / 10;
    const ha = st.align & 6, va = st.align & 24;
    const anchor = ha === 6 ? 'middle' : ha === 2 ? 'end' : 'start';
    const base = va === 24 ? '' : va === 8 ? ' dominant-baseline="text-after-edge"' : ' dominant-baseline="text-before-edge"';
    let xs = '';
    if (dx.length === n && anchor === 'start' && !angle) {
      // 글자마다 위치 (EMF 가 준 간격 그대로)
      const sx = scaleXY()[0] * Math.hypot(st.xf[0], st.xf[1]);
      const pos = [];
      let acc = 0;
      for (let i = 0; i < n; i++) { pos.push(num(p0[0] + acc * sx)); acc += dx[i]; }
      xs = pos.join(' ');
    }
    let len = '';
    if (!xs && dx.length === n) {
      const sx = scaleXY()[0] * Math.hypot(st.xf[0], st.xf[1]);
      const total = dx.reduce((a, b) => a + b, 0) * sx;
      if (total > 0) len = ` textLength="${num(total)}" lengthAdjust="spacingAndGlyphs"`;
    }
    const family = `${f.face ? `'${escXml(f.face).replace(/'/g, '')}', ` : ''}'맑은 고딕', 'Malgun Gothic', sans-serif`;
    const deco = [f.underline && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
    const rot = angle ? ` transform="rotate(${num(angle)} ${num(p0[0])} ${num(p0[1])})"` : '';
    emit(`<text x="${xs || num(p0[0])}" y="${num(p0[1])}" font-family="${family}" font-size="${num(size)}"${f.weight >= 600 ? ' font-weight="bold"' : ''}${f.italic ? ' font-style="italic"' : ''}${deco ? ` text-decoration="${deco}"` : ''} fill="${st.textColor}" text-anchor="${anchor}"${base}${len}${rot} xml:space="preserve">${escXml(s)}</text>`);
  };

  let o = 0;
  const end = bytes.length;
  while (o + 8 <= end) {
    const type = u32(o), size = u32(o + 4);
    if (size < 8 || o + size > end) break;
    const d = o + 8; // 레코드 몸통
    switch (type) {
      case 14: o = end; continue; // EOF
      case 9: st.winExt = [i32(d), i32(d + 4)]; break;
      case 10: st.winOrg = [i32(d), i32(d + 4)]; break;
      case 11: st.vpExt = [i32(d), i32(d + 4)]; break;
      case 12: st.vpOrg = [i32(d), i32(d + 4)]; break;
      case 17: st.mapMode = u32(d); break;
      case 18: st.bkMode = u32(d); break;
      case 19: st.fillRule = u32(d) === 2 ? 'nonzero' : 'evenodd'; break;
      case 22: st.align = u32(d); break;
      case 24: st.textColor = colorRef(dv, d); break;
      case 25: st.bkColor = colorRef(dv, d); break;
      case 27: st.cur = [i32(d), i32(d + 4)]; if (path) path.push(`M${pt(...st.cur).map(num).join(' ')}`); break; // MOVETOEX
      case 54: { // LINETO
        const a = pt(...st.cur);
        st.cur = [i32(d), i32(d + 4)];
        const b = pt(...st.cur);
        if (path) path.push(`L${num(b[0])} ${num(b[1])}`);
        else shape(`M${num(a[0])} ${num(a[1])}L${num(b[0])} ${num(b[1])}`, { fill: false });
        break;
      }
      case 33: stack.push({ ...st }); break; // SAVEDC
      case 34: { // RESTOREDC (음수: 상대 위치)
        const k = i32(d);
        const idx = k < 0 ? stack.length + k : k - 1;
        if (idx >= 0 && idx < stack.length) { st = stack[idx]; stack.length = idx; }
        break;
      }
      case 35: st.xf = [f32(d), f32(d + 4), f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)]; break;
      case 36: { // MODIFYWORLDTRANSFORM
        const m = [f32(d), f32(d + 4), f32(d + 8), f32(d + 12), f32(d + 16), f32(d + 20)];
        const mode = u32(d + 24);
        st.xf = mode === 1 ? IDENT : mode === 2 ? mul(m, st.xf) : mode === 3 ? mul(st.xf, m) : m;
        break;
      }
      case 37: { // SELECTOBJECT
        const ih = u32(d);
        const k = ih & 0x7fffffff;
        const ob = ih & 0x80000000 ? (STOCK[k] ?? (k >= 10 && k <= 17 ? DEFAULT_FONT : null)) : objs[ih];
        if (ob?.kind === 'pen') st.pen = ob;
        else if (ob?.kind === 'brush') st.brush = ob;
        else if (ob?.kind === 'font') st.font = ob;
        break;
      }
      case 38: { // CREATEPEN
        const ps = u32(d + 4);
        objs[u32(d)] = { kind: 'pen', style: ps & 0xf, width: i32(d + 8), color: colorRef(dv, d + 16), geometric: i32(d + 8) > 1, ...penShape(ps) };
        break;
      }
      case 95: { // EXTCREATEPEN
        const ps = u32(d + 20);
        objs[u32(d)] = { kind: 'pen', style: ps & 0xf, width: u32(d + 24), color: colorRef(dv, d + 32), geometric: !!(ps & 0x10000), ...penShape(ps) };
        if (u32(d + 28) === 1) objs[u32(d)].style = 5; // BS_NULL 붓
        break;
      }
      case 39: objs[u32(d)] = { kind: 'brush', style: u32(d + 4) === 1 ? 1 : 0, color: colorRef(dv, d + 8) }; break; // CREATEBRUSHINDIRECT
      case 93: case 94: objs[u32(d)] = { kind: 'brush', style: 0, color: '#808080' }; break; // 무늬 붓: 회색으로 대신
      case 82: { // EXTCREATEFONTINDIRECTW
        const l = d + 4;
        let face = '';
        for (let i = 0; i < 32; i++) {
          const ch = dv.getUint16(l + 28 + i * 2, true);
          if (!ch) break;
          face += String.fromCharCode(ch);
        }
        objs[u32(d)] = { kind: 'font', height: i32(l), esc: i32(l + 8), weight: i32(l + 16), italic: !!dv.getUint8(l + 20), underline: !!dv.getUint8(l + 21), strike: !!dv.getUint8(l + 22), face };
        break;
      }
      case 40: delete objs[u32(d)]; break;
      case 59: path = []; break; // BEGINPATH
      case 60: lastPath = path; path = null; break; // ENDPATH
      case 61: if (path) path.push('Z'); break; // CLOSEFIGURE
      case 68: path = null; lastPath = null; break; // ABORTPATH
      case 62: case 63: case 64: { // FILLPATH / STROKEANDFILLPATH / STROKEPATH
        const p = path ?? lastPath;
        if (p?.length) shape(p.join(''), { fill: type !== 64, stroke: type !== 62 });
        lastPath = null;
        break;
      }
      case 67: { // SELECTCLIPPATH: 경로의 외곽 사각형으로 잘라내기
        const p = lastPath;
        if (p?.length) {
          const nums = p.join(' ').match(/-?\d+(\.\d+)?/g) ?? [];
          const r = [Infinity, Infinity, -Infinity, -Infinity];
          for (let i = 0; i + 1 < nums.length; i += 2) {
            const x = Number(nums[i]), y = Number(nums[i + 1]);
            r[0] = Math.min(r[0], x); r[1] = Math.min(r[1], y); r[2] = Math.max(r[2], x); r[3] = Math.max(r[3], y);
          }
          if (r[0] <= r[2]) st.clip = r;
        }
        lastPath = null;
        break;
      }
      case 30: { // INTERSECTCLIPRECT
        const [a, b] = rectPts(d);
        const r = [Math.min(a[0], b[0]), Math.min(a[1], b[1]), Math.max(a[0], b[0]), Math.max(a[1], b[1])];
        st.clip = st.clip ? [Math.max(st.clip[0], r[0]), Math.max(st.clip[1], r[1]), Math.min(st.clip[2], r[2]), Math.min(st.clip[3], r[3])] : r;
        break;
      }
      case 75: { // EXTSELECTCLIPRGN
        const cb = u32(d), mode = u32(d + 4);
        if (!cb) { if (mode === 5) st.clip = null; break; }
        const rd = d + 8;
        const cnt = u32(rd + 8);
        if (cnt) {
          // 여러 사각형이면 외곽 사각형 (장치 좌표)
          let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
          for (let i = 0; i < cnt; i++) {
            const q = rd + 32 + i * 16;
            x1 = Math.min(x1, i32(q)); y1 = Math.min(y1, i32(q + 4)); x2 = Math.max(x2, i32(q + 8)); y2 = Math.max(y2, i32(q + 12));
          }
          const r = [x1, y1, x2, y2];
          st.clip = mode === 1 && st.clip ? [Math.max(st.clip[0], r[0]), Math.max(st.clip[1], r[1]), Math.min(st.clip[2], r[2]), Math.min(st.clip[3], r[3])] : r;
        }
        break;
      }
      case 42: case 43: case 44: { // ELLIPSE / RECTANGLE / ROUNDRECT
        const [a, b] = rectPts(d);
        const x = Math.min(a[0], b[0]), y = Math.min(a[1], b[1]), w = Math.abs(b[0] - a[0]), h = Math.abs(b[1] - a[1]);
        if (type === 42) {
          const rx = w / 2, ry = h / 2;
          shape(`M${num(x)} ${num(y + ry)}A${num(rx)} ${num(ry)} 0 1 0 ${num(x + w)} ${num(y + ry)}A${num(rx)} ${num(ry)} 0 1 0 ${num(x)} ${num(y + ry)}Z`);
        } else if (type === 44) {
          const s = lenScale();
          const rx = Math.min(w / 2, (i32(d + 16) * s) / 2), ry = Math.min(h / 2, (i32(d + 20) * s) / 2);
          shape(`M${num(x + rx)} ${num(y)}H${num(x + w - rx)}A${num(rx)} ${num(ry)} 0 0 1 ${num(x + w)} ${num(y + ry)}V${num(y + h - ry)}A${num(rx)} ${num(ry)} 0 0 1 ${num(x + w - rx)} ${num(y + h)}H${num(x + rx)}A${num(rx)} ${num(ry)} 0 0 1 ${num(x)} ${num(y + h - ry)}V${num(y + ry)}A${num(rx)} ${num(ry)} 0 0 1 ${num(x + rx)} ${num(y)}Z`);
        } else shape(`M${num(x)} ${num(y)}H${num(x + w)}V${num(y + h)}H${num(x)}Z`);
        break;
      }
      case 3: case 4: case 86: case 87: { // POLYGON / POLYLINE (32 · 16비트)
        const n = u32(d + 16);
        const pts = type >= 86 ? readPts16(d + 20, n) : readPts32(d + 20, n);
        if (!pts.length) break;
        const close = type === 3 || type === 86;
        if (path) path.push(polyD(pts, close));
        else shape(polyD(pts, close), { fill: close });
        break;
      }
      case 6: case 89: { // POLYLINETO
        const n = u32(d + 16);
        const pts = type === 89 ? readPts16(d + 20, n) : readPts32(d + 20, n);
        if (!pts.length) break;
        const d0 = pts.map((p) => `L${num(p[0])} ${num(p[1])}`).join('');
        if (path) path.push(d0);
        else shape(`M${pt(...st.cur).map(num).join(' ')}${d0}`, { fill: false });
        st.cur = type === 89 ? [i16(d + 20 + (n - 1) * 4), i16(d + 22 + (n - 1) * 4)] : [i32(d + 20 + (n - 1) * 8), i32(d + 24 + (n - 1) * 8)];
        break;
      }
      case 2: case 85: case 5: case 88: { // POLYBEZIER / POLYBEZIERTO
        const n = u32(d + 16);
        const pts = type === 85 || type === 88 ? readPts16(d + 20, n) : readPts32(d + 20, n);
        if (!pts.length) break;
        const to = type === 5 || type === 88;
        const dd = to ? bezD(pt(...st.cur), pts, !path) : bezD(pts[0], pts.slice(1), true);
        if (path) path.push(dd);
        else shape(dd, { fill: false });
        const q = d + 20 + (n - 1) * (type === 85 || type === 88 ? 4 : 8);
        st.cur = type === 85 || type === 88 ? [i16(q), i16(q + 2)] : [i32(q), i32(q + 4)];
        break;
      }
      case 7: case 8: case 90: case 91: { // POLYPOLYLINE / POLYPOLYGON
        const polys = u32(d + 16), n = u32(d + 20);
        const small = type >= 90;
        const counts = [];
        for (let i = 0; i < polys; i++) counts.push(u32(d + 24 + i * 4));
        const base = d + 24 + polys * 4;
        const pts = small ? readPts16(base, n) : readPts32(base, n);
        const close = type === 8 || type === 91;
        let k = 0, dd = '';
        for (const c of counts) { dd += polyD(pts.slice(k, k + c), close); k += c; }
        if (path) path.push(dd);
        else shape(dd, { fill: close });
        break;
      }
      case 83: case 84: text(o, type === 84); break; // EXTTEXTOUTA / W
      case 76: { // BITBLT
        const x = i32(d + 16), y = i32(d + 20), w = i32(d + 24), h = i32(d + 28), rop = u32(d + 32);
        const offBmi = u32(d + 76), cbBmi = u32(d + 80), offBits = u32(d + 84), cbBits = u32(d + 88);
        if (cbBmi && cbBits) image(x, y, w, h, bytes.subarray(o + offBmi, o + offBmi + cbBmi), bytes.subarray(o + offBits, o + offBits + cbBits));
        else if (rop === 0x00f00021 && st.brush?.style !== 1) fillRect(x, y, w, h, st.brush.color); // PATCOPY
        else if (rop === 0x00ff0062) fillRect(x, y, w, h, '#ffffff'); // WHITENESS
        else if (rop === 0x00000042) fillRect(x, y, w, h, '#000000'); // BLACKNESS
        break;
      }
      case 77: { // STRETCHBLT
        const x = i32(d + 16), y = i32(d + 20), w = i32(d + 24), h = i32(d + 28);
        const offBmi = u32(d + 76), cbBmi = u32(d + 80), offBits = u32(d + 84), cbBits = u32(d + 88);
        if (cbBmi && cbBits) image(x, y, w, h, bytes.subarray(o + offBmi, o + offBmi + cbBmi), bytes.subarray(o + offBits, o + offBits + cbBits));
        break;
      }
      case 81: { // STRETCHDIBITS
        const x = i32(d + 16), y = i32(d + 20);
        const offBmi = u32(d + 40), cbBmi = u32(d + 44), offBits = u32(d + 48), cbBits = u32(d + 52);
        const w = i32(d + 64), h = i32(d + 68);
        image(x, y, w, h, bytes.subarray(o + offBmi, o + offBmi + cbBmi), bytes.subarray(o + offBits, o + offBits + cbBits));
        break;
      }
      default: break;
    }
    o += size;
  }
  if (!drawn) return null;
  return `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${bx1} ${by1} ${vw} ${vh}" width="${vw}" height="${vh}" preserveAspectRatio="none">${clips.length ? `<defs>${clips.join('')}</defs>` : ''}${out.join('')}</svg>`;
}

function penShape(ps) {
  const cap = { 0: 'round', 0x100: 'square', 0x200: 'butt' }[ps & 0xf00] ?? 'round';
  const join = { 0: 'round', 0x1000: 'bevel', 0x2000: 'miter' }[ps & 0xf000] ?? 'round';
  return { cap, join };
}

/** EMF 바이트 → SVG data URL (못 그리면 null) */
export function emfDataUrl(bytes) {
  const svg = emfToSvg(bytes);
  if (!svg) return null;
  return `data:image/svg+xml;base64,${b64(new TextEncoder().encode(svg))}`;
}
