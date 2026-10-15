// PDF 파일 쓰기 (DOM 없음, 의존성 없음)
// 글자: Adobe-Korea1 미리 정의된 CID 글꼴(HYGoThic-Medium · HYSMyeongJo-Medium, UniKS-UCS2-H) — 글꼴을 넣지 않아 파일이 작고,
//   글자를 고르고 찾을 수 있음. 글자 하나하나의 위치는 TJ 간격으로 맞춤 (보는 프로그램의 대체 글꼴 너비와 상관없이 같은 자리).
// 그림: JPEG 은 그대로(DCTDecode), 나머지는 RGB + 투명도(SMask) 를 deflate. 연결(링크) 주석, 제목 책갈피(Outlines), 문서 정보.
import { subsetTtf } from './ttfsubset.js';

const PT = 0.75; // px → pt
const n2 = (v) => (Math.round(v * 100) / 100).toString();
const enc = new TextEncoder();
const hex4 = (c) => c.toString(16).padStart(4, '0');
const utf16hex = (s) => `<FEFF${Array.from(s, (ch) => { const c = ch.codePointAt(0); if (c > 0xffff) { const v = c - 0x10000; return hex4(0xd800 + (v >> 10)) + hex4(0xdc00 + (v & 0x3ff)); } return hex4(c); }).join('')}>`;
const pdfStr = (s) => `(${String(s).replace(/[\\()]/g, (m) => `\\${m}`).replace(/[^\x20-\x7e]/g, '')})`;
const rgb = (c) => {
  const m = /^#?([0-9a-f]{6})$/i.exec(c ?? '') ?? /^#?([0-9a-f]{3})$/i.exec(c ?? '');
  if (!m) return [0, 0, 0];
  const h = m[1].length === 3 ? m[1].split('').map((x) => x + x).join('') : m[1];
  return [0, 2, 4].map((i) => parseInt(h.slice(i, i + 2), 16) / 255);
};
const col = (c) => rgb(c).map((v) => n2(v)).join(' ');

/** deflate (zlib 형식) — 브라우저 · Node 의 CompressionStream */
export async function deflate(bytes) {
  if (typeof CompressionStream !== 'function') return null;
  const s = new Blob([bytes]).stream().pipeThrough(new CompressionStream('deflate'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}

export class PdfWriter {
  constructor() { this.pages = []; this.images = []; this.fontsUsed = new Set(); this.outline = []; this.info = {}; this.embedded = []; }
  /** TrueType 글꼴 넣기 (ttfsubset.openTtf 결과) → 글꼴 참조 */
  embedFont(font) { const ref = { key: `E${this.embedded.length + 1}`, font, used: new Map() }; this.embedded.push(ref); return ref; }
  /** 쪽 하나 (크기 px) → 그리기 도구 */
  addPage(wPx, hPx) {
    const page = { w: wPx * PT, h: hPx * PT, ops: [], links: [], index: this.pages.length };
    this.pages.push(page);
    const H = page.h;
    const Y = (y) => H - y * PT;
    const op = (s) => page.ops.push(s);
    const self = this;
    return {
      page,
      /** 칠한 사각형 (px) */
      rect(x, y, w, h, color, alpha = 1) { if (w <= 0 || h <= 0) return; op(`q ${alpha < 1 ? `/GA${self.gs(alpha)} gs ` : ''}${col(color)} rg ${n2(x * PT)} ${n2(Y(y + h))} ${n2(w * PT)} ${n2(h * PT)} re f Q`); },
      /** 선 (px) */
      line(x1, y1, x2, y2, width, color, dash = null) { op(`q ${col(color)} RG ${n2(Math.max(0.1, width * PT))} w ${dash ? `[${dash.map((d) => n2(d * PT)).join(' ')}] 0 d ` : ''}${n2(x1 * PT)} ${n2(Y(y1))} m ${n2(x2 * PT)} ${n2(Y(y2))} l S Q`); },
      /** 그림 (images[] 번호, px 상자) */
      image(id, x, y, w, h, clip = null) {
        const c = clip ? `${n2(clip.x * PT)} ${n2(Y(clip.y + clip.h))} ${n2(clip.w * PT)} ${n2(clip.h * PT)} re W n ` : '';
        op(`q ${c}${n2(w * PT)} 0 0 ${n2(h * PT)} ${n2(x * PT)} ${n2(Y(y + h))} cm /Im${id} Do Q`);
      },
      /**
       * 글자 한 줄 조각: chars[i] 를 xs[i](px) 에, 기준선 y(px), 크기 sizePx
       * style: { color, serif, bold, italic, alpha }
       */
      text(chars, xs, y, sizePx, style = {}) {
        if (!chars.length) return;
        const font = style.serif ? 'F2' : 'F1';
        self.fontsUsed.add(font);
        const s = sizePx * PT;
        const codes = chars.map((ch) => { const c = ch.codePointAt(0); return c > 0xffff ? 0x25a1 : c; });
        const wid = (c) => (c < 0x7f ? 500 : 1000);
        let tj = '[';
        let buf = '';
        for (let i = 0; i < codes.length; i++) {
          buf += hex4(codes[i]);
          if (i < codes.length - 1) {
            const adv = (xs[i + 1] - xs[i]) * PT;
            const adj = Math.round((wid(codes[i]) / 1000 * s - adv) * 1000 / s);
            if (adj !== 0) { tj += `<${buf}> ${adj} `; buf = ''; }
          }
        }
        if (buf) tj += `<${buf}>`;
        tj += ']';
        const skew = style.italic ? 0.2 : 0;
        const bold = style.bold ? `2 Tr ${n2(s * 0.03)} w ${col(style.color)} RG ` : '';
        op(`q ${style.alpha && style.alpha < 1 ? `/GA${self.gs(style.alpha)} gs ` : ''}BT ${bold}${col(style.color)} rg /${font} ${n2(s)} Tf 1 0 ${n2(skew)} 1 ${n2(xs[0] * PT)} ${n2(Y(y))} Tm ${tj} TJ ET Q`);
      },
      /** 넣은 글꼴로 글리프 그리기: gids[i] (글리프 번호) 를 xs[i] 에, unis[i] = 글자 (복사 · 찾기용) */
      glyphs(ref, gids, unis, xs, y, sizePx, style = {}) {
        if (!gids.length) return;
        const s = sizePx * PT;
        gids.forEach((g, i) => { if (!ref.used.has(g)) ref.used.set(g, unis[i]); });
        let tj = '['; let buf = '';
        for (let i = 0; i < gids.length; i++) {
          buf += hex4(gids[i]);
          if (i < gids.length - 1) {
            const adj = Math.round((ref.font.width(gids[i]) / 1000 * s - (xs[i + 1] - xs[i]) * PT) * 1000 / s);
            if (adj !== 0) { tj += `<${buf}> ${adj} `; buf = ''; }
          }
        }
        if (buf) tj += `<${buf}>`;
        tj += ']';
        const skew = style.italic ? 0.2 : 0;
        const bold = style.bold ? `2 Tr ${n2(s * 0.03)} w ${col(style.color)} RG ` : '';
        op(`q ${style.alpha && style.alpha < 1 ? `/GA${self.gs(style.alpha)} gs ` : ''}BT ${bold}${col(style.color)} rg /${ref.key} ${n2(s)} Tf 1 0 ${n2(skew)} 1 ${n2(xs[0] * PT)} ${n2(Y(y))} Tm ${tj} TJ ET Q`);
      },
      link(x, y, w, h, url) { page.links.push({ rect: [x * PT, Y(y + h), (x + w) * PT, Y(y)], url }); },
    };
  }
  gs(alpha) { this.alphas ??= new Map(); const k = Math.round(alpha * 100); if (!this.alphas.has(k)) this.alphas.set(k, k); return k; }
  /** 그림 등록: { jpeg: bytes, w, h } 또는 { rgb: bytes, alpha?: bytes, w, h } → 번호 */
  addImage(img) { this.images.push(img); return this.images.length - 1; }
  addOutline(title, pageIndex, yPx, level = 0) { this.outline.push({ title, page: pageIndex, y: yPx, level }); }

  async build() {
    const objs = []; // [번호] = 내용 (문자열 또는 { dict, stream })
    const alloc = () => { objs.push(null); return objs.length; };
    const catalog = alloc(); const pagesId = alloc();
    // 글꼴
    const fontIds = {};
    for (const [key, base] of [['F1', 'HYGoThic-Medium'], ['F2', 'HYSMyeongJo-Medium']]) {
      if (!this.fontsUsed.has(key)) continue;
      const desc = alloc(); const cid = alloc(); const t0 = alloc();
      objs[desc - 1] = `<< /Type /FontDescriptor /FontName /${base} /Flags 6 /FontBBox [-100 -142 1000 880] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 700 /StemV 80 >>`;
      objs[cid - 1] = `<< /Type /Font /Subtype /CIDFontType0 /BaseFont /${base} /CIDSystemInfo << /Registry (Adobe) /Ordering (Korea1) /Supplement 1 >> /FontDescriptor ${desc} 0 R /DW 1000 /W [1 95 500] >>`;
      objs[t0 - 1] = `<< /Type /Font /Subtype /Type0 /BaseFont /${base} /Encoding /UniKS-UCS2-H /DescendantFonts [${cid} 0 R] >>`;
      fontIds[key] = t0;
    }
    // 넣은 글꼴 (부분 집합)
    for (const ref of this.embedded) {
      if (!ref.used.size) continue;
      const f = ref.font;
      const sub = subsetTtf(f, [...ref.used.keys()]);
      const z = await deflate(sub);
      const ff = alloc(); const desc = alloc(); const cid = alloc(); const tu = alloc(); const t0 = alloc();
      objs[ff - 1] = { dict: `/Length1 ${sub.length}${z ? ' /Filter /FlateDecode' : ''}`, stream: z ?? sub };
      const tagName = `${String.fromCharCode(...Array.from({ length: 6 }, (_, i) => 65 + ((ref.key.charCodeAt(1) * 7 + i * 3 + this.embedded.indexOf(ref)) % 26)))}+${f.psName}`;
      objs[desc - 1] = `<< /Type /FontDescriptor /FontName /${tagName} /Flags 4 /FontBBox [${f.bbox.join(' ')}] /ItalicAngle 0 /Ascent ${f.ascent} /Descent ${f.descent} /CapHeight ${f.capHeight} /StemV 80 /FontFile2 ${ff} 0 R >>`;
      const gids = [...ref.used.keys()].sort((a, b) => a - b);
      const W = gids.map((g) => `${g} [${f.width(g)}]`).join(' ');
      objs[cid - 1] = `<< /Type /Font /Subtype /CIDFontType2 /BaseFont /${tagName} /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> /FontDescriptor ${desc} 0 R /DW 1000 /W [${W}] /CIDToGIDMap /Identity >>`;
      const u16be = (str) => Array.from(str, (ch) => { const c = ch.codePointAt(0); if (c > 0xffff) { const v = c - 0x10000; return hex4(0xd800 + (v >> 10)) + hex4(0xdc00 + (v & 0x3ff)); } return hex4(c); }).join('');
      const lines = [];
      for (let i = 0; i < gids.length; i += 100) { const part = gids.slice(i, i + 100); lines.push(`${part.length} beginbfchar`, ...part.map((g) => `<${hex4(g)}> <${u16be(ref.used.get(g) || ' ')}>`), 'endbfchar'); }
      const cmap = enc.encode(`/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${lines.join('\n')}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`);
      const zc = await deflate(cmap);
      objs[tu - 1] = { dict: zc ? '/Filter /FlateDecode' : '', stream: zc ?? cmap };
      objs[t0 - 1] = `<< /Type /Font /Subtype /Type0 /BaseFont /${tagName} /Encoding /Identity-H /DescendantFonts [${cid} 0 R] /ToUnicode ${tu} 0 R >>`;
      fontIds[ref.key] = t0;
    }
    // 그림
    const imgIds = [];
    for (const im of this.images) {
      const id = alloc();
      if (im.jpeg) objs[id - 1] = { dict: `/Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /${im.gray ? 'DeviceGray' : im.cmyk ? 'DeviceCMYK' : 'DeviceRGB'} /BitsPerComponent 8 /Filter /DCTDecode${im.cmyk ? ' /Decode [1 0 1 0 1 0 1 0]' : ''}`, stream: im.jpeg };
      else {
        let smask = '';
        if (im.alpha) { const sid = alloc(); const a = await deflate(im.alpha); objs[sid - 1] = { dict: `/Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceGray /BitsPerComponent 8${a ? ' /Filter /FlateDecode' : ''}`, stream: a ?? im.alpha }; smask = ` /SMask ${sid} 0 R`; }
        const c = await deflate(im.rgb);
        objs[id - 1] = { dict: `/Type /XObject /Subtype /Image /Width ${im.w} /Height ${im.h} /ColorSpace /DeviceRGB /BitsPerComponent 8${c ? ' /Filter /FlateDecode' : ''}${smask}`, stream: c ?? im.rgb };
      }
      imgIds.push(id);
    }
    // 투명도
    const gsDict = this.alphas?.size ? `/ExtGState << ${[...this.alphas.keys()].map((k) => `/GA${k} << /ca ${n2(k / 100)} /CA ${n2(k / 100)} >>`).join(' ')} >>` : '';
    // 쪽
    const pageIds = [];
    for (const p of this.pages) {
      const pid = alloc(); const cid = alloc();
      pageIds.push(pid);
      const body = enc.encode(p.ops.join('\n'));
      const z = await deflate(body);
      objs[cid - 1] = { dict: z ? '/Filter /FlateDecode' : '', stream: z ?? body };
      const annots = p.links.map((l) => { const a = alloc(); objs[a - 1] = `<< /Type /Annot /Subtype /Link /Rect [${l.rect.map(n2).join(' ')}] /Border [0 0 0] /A << /S /URI /URI ${pdfStr(l.url)} >> >>`; return `${a} 0 R`; });
      const fonts = Object.entries(fontIds).map(([k, id]) => `/${k} ${id} 0 R`).join(' ');
      const xobj = imgIds.length ? `/XObject << ${imgIds.map((id, i) => `/Im${i} ${id} 0 R`).join(' ')} >>` : '';
      objs[pid - 1] = `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${n2(p.w)} ${n2(p.h)}] /Resources << /Font << ${fonts} >> ${xobj} ${gsDict} >> /Contents ${cid} 0 R${annots.length ? ` /Annots [${annots.join(' ')}]` : ''} >>`;
    }
    objs[pagesId - 1] = `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(' ')}] /Count ${pageIds.length} >>`;
    // 책갈피 (제목 수준별 나무)
    let outlines = '';
    if (this.outline.length) {
      const root = alloc();
      const items = this.outline.map((o) => ({ ...o, id: alloc(), kids: [] }));
      const top = []; const stack = [];
      for (const it of items) {
        while (stack.length && stack[stack.length - 1].level >= it.level) stack.pop();
        it.parent = stack.length ? stack[stack.length - 1] : null;
        (it.parent ? it.parent.kids : top).push(it);
        stack.push(it);
      }
      const write = (list, parentId) => list.forEach((it, i) => {
        const pg = this.pages[it.page];
        const d = [`/Title ${utf16hex(it.title)}`, `/Parent ${parentId} 0 R`, `/Dest [${pageIds[it.page]} 0 R /XYZ 0 ${n2(pg.h - it.y * PT)} null]`];
        if (i > 0) d.push(`/Prev ${list[i - 1].id} 0 R`);
        if (i < list.length - 1) d.push(`/Next ${list[i + 1].id} 0 R`);
        if (it.kids.length) { d.push(`/First ${it.kids[0].id} 0 R`, `/Last ${it.kids[it.kids.length - 1].id} 0 R`, `/Count ${it.kids.length}`); write(it.kids, it.id); }
        objs[it.id - 1] = `<< ${d.join(' ')} >>`;
      });
      write(top, root);
      objs[root - 1] = `<< /Type /Outlines /First ${top[0].id} 0 R /Last ${top[top.length - 1].id} 0 R /Count ${top.length} >>`;
      outlines = ` /Outlines ${root} 0 R /PageMode /UseOutlines`;
    }
    objs[catalog - 1] = `<< /Type /Catalog /Pages ${pagesId} 0 R${outlines} /Lang (ko-KR) >>`;
    const infoId = alloc();
    const now = new Date(); const pad = (v) => String(v).padStart(2, '0');
    const date = `D:${now.getFullYear()}${pad(now.getMonth() + 1)}${pad(now.getDate())}${pad(now.getHours())}${pad(now.getMinutes())}${pad(now.getSeconds())}`;
    objs[infoId - 1] = `<< /Producer (WIWORD) /Creator (WIWORD) /CreationDate (${date})${this.info.title ? ` /Title ${utf16hex(this.info.title)}` : ''}${this.info.author ? ` /Author ${utf16hex(this.info.author)}` : ''}${this.info.subject ? ` /Subject ${utf16hex(this.info.subject)}` : ''} >>`;
    // 이어 붙이기
    const parts = [enc.encode('%PDF-1.7\n%\xE2\xE3\xCF\xD3\n')];
    let pos = parts[0].length;
    const offsets = [];
    objs.forEach((o, i) => {
      offsets.push(pos);
      let chunk;
      if (typeof o === 'string') chunk = [enc.encode(`${i + 1} 0 obj\n${o}\nendobj\n`)];
      else chunk = [enc.encode(`${i + 1} 0 obj\n<< ${o.dict} /Length ${o.stream.length} >>\nstream\n`), o.stream, enc.encode('\nendstream\nendobj\n')];
      for (const c of chunk) { parts.push(c); pos += c.length; }
    });
    const xref = `xref\n0 ${objs.length + 1}\n0000000000 65535 f \n${offsets.map((o) => `${String(o).padStart(10, '0')} 00000 n \n`).join('')}trailer\n<< /Size ${objs.length + 1} /Root ${catalog} 0 R /Info ${infoId} 0 R >>\nstartxref\n${pos}\n%%EOF\n`;
    parts.push(enc.encode(xref));
    const total = parts.reduce((n, p) => n + p.length, 0);
    const out = new Uint8Array(total); let k = 0;
    for (const p of parts) { out.set(p, k); k += p.length; }
    return out;
  }
}
