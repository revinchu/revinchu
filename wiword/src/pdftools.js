// PDF 도구 모음 (DOM 없음, 의존성 없음) — 이음오피스(IeumOffice)의 PDF 도구 창을 참고해 직접 구현
// 원본 PDF 의 개체를 그대로 옮겨 새 PDF 를 만듦: 병합 · 분할 · 페이지 추출/삭제/순서/회전/역순 · 모아찍기 · 양면 인쇄 맞추기 ·
// 배경 겹치기 · 주석 한꺼번에 지우기 · 빈 페이지 찾기 · 이미지로 PDF 만들기 · 쪽 번호/워터마크 · 암호 해제 · 정리(용량 줄이기) · 문서 정보
// 결과는 항상 새 파일 (원본은 바꾸지 않음). 암호 문서는 열기 암호로 풀어서 읽고, 결과는 암호 없이 씀.
import { PdfDoc, Ref, Name, Stream, Lexer, Cmd } from './pdf.js';
import { inflate } from './zip.js';
import { deflate } from './pdfwrite.js';
import { openTtf, subsetTtf } from './ttfsubset.js';

const latin1 = new TextDecoder('latin1');
const enc = new TextEncoder();
const isDict = (v) => v && typeof v === 'object' && !Array.isArray(v) && !(v instanceof Uint8Array) && !(v instanceof Name) && !(v instanceof Ref) && !(v instanceof Stream) && !(v instanceof Cmd) && !v.stream;
const nm = (n) => new Name(n);

/** 용지 크기 (pt) */
export const PAPERS = { A3: [842, 1191], A4: [595, 842], A5: [420, 595], B4: [729, 1032], B5: [516, 729], Letter: [612, 792], Legal: [612, 1008] };

// ───────────── 행렬 (PDF 순서: a b c d e f, 앞의 것을 먼저 적용) ─────────────
const mul = (a, b) => [a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3], a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3], a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5]];
const inv = (m) => { const d = m[0] * m[3] - m[1] * m[2] || 1; return [m[3] / d, -m[1] / d, -m[2] / d, m[0] / d, (m[2] * m[5] - m[3] * m[4]) / d, (m[1] * m[4] - m[0] * m[5]) / d]; };
const num = (v) => { if (!Number.isFinite(v)) return '0'; const r = Math.round(v * 10000) / 10000; return Object.is(r, -0) ? '0' : String(r); };
const mtx = (m) => m.map(num).join(' ');
/** 쪽의 회전(시계 방향)을 반영: 쪽 좌표(상자 원점 0) → 보이는 좌표 */
const rotM = (r, W, H) => (r === 90 ? [0, -1, 1, 0, 0, W] : r === 180 ? [-1, 0, 0, -1, W, H] : r === 270 ? [0, 1, -1, 0, H, 0] : [1, 0, 0, 1, 0, 0]);

// ───────────── 열기 ─────────────
/** PDF 바이트 → { doc, root, pages: [{ ref, node, res, box, rotate, w, h }] } (w/h = 보이는 크기 pt) */
export function openPdf(bytes, { password = '' } = {}) {
  if (!(bytes?.length > 8) || !latin1.decode(bytes.subarray(0, 1024)).includes('%PDF')) throw new Error('PDF 파일이 아닙니다.');
  const doc = new PdfDoc(bytes);
  if (doc.trailer.Encrypt) doc.unlock(password);
  const root = doc.val(doc.trailer.Root);
  if (!root) throw new Error('PDF 구조를 읽을 수 없습니다 (손상된 파일).');
  const pages = [];
  const seen = new Set();
  const nums = (a) => (Array.isArray(doc.val(a)) ? doc.val(a).map((x) => Number(doc.val(x)) || 0) : null);
  const walk = (ref, inh, depth) => {
    const node = doc.val(ref);
    if (!isDict(node) || depth > 40) return;
    if (ref instanceof Ref) { if (seen.has(ref.num)) return; seen.add(ref.num); }
    const next = { Resources: node.Resources ?? inh.Resources, MediaBox: node.MediaBox ?? inh.MediaBox, CropBox: node.CropBox ?? inh.CropBox, Rotate: node.Rotate ?? inh.Rotate };
    const type = doc.val(node.Type)?.n;
    if (type === 'Page' || (!node.Kids && (node.Contents || node.MediaBox))) {
      const mb = nums(next.MediaBox) ?? [0, 0, 612, 792];
      const cb = nums(next.CropBox);
      let b = cb && cb.length === 4 ? cb : mb;
      b = [Math.min(b[0], b[2]), Math.min(b[1], b[3]), Math.max(b[0], b[2]), Math.max(b[1], b[3])];
      if (b[2] - b[0] < 1 || b[3] - b[1] < 1) b = [0, 0, 612, 792];
      const rotate = ((Math.round((Number(doc.val(next.Rotate)) || 0) / 90) * 90) % 360 + 360) % 360;
      const W = b[2] - b[0]; const H = b[3] - b[1];
      pages.push({ ref: ref instanceof Ref ? ref : null, node, res: next.Resources, mediaBox: next.MediaBox, cropBox: next.CropBox, box: b, rotate, w: rotate % 180 ? H : W, h: rotate % 180 ? W : H, index: pages.length });
      return;
    }
    for (const k of doc.val(node.Kids) ?? []) walk(k, next, depth + 1);
  };
  walk(root.Pages, {}, 0);
  if (!pages.length) throw new Error('PDF 에 페이지가 없습니다.');
  return { doc, root, pages, info: doc.val(doc.trailer.Info) ?? {} };
}

/** PDF 문자열 → JS 문자열 (UTF-16BE BOM · UTF-8 BOM · PDFDocEncoding ≈ latin1) */
export function pdfText(b) {
  if (!(b instanceof Uint8Array)) return b instanceof Name ? b.n : b == null ? '' : String(b);
  if (b[0] === 0xfe && b[1] === 0xff) { let s = ''; for (let i = 2; i + 1 < b.length; i += 2) s += String.fromCharCode((b[i] << 8) | b[i + 1]); return s; }
  if (b[0] === 0xef && b[1] === 0xbb && b[2] === 0xbf) return new TextDecoder().decode(b.subarray(3));
  return latin1.decode(b);
}
/** JS 문자열 → PDF 글 문자열 (ASCII 는 그대로, 아니면 UTF-16BE) */
function textBytes(s) {
  s = String(s ?? '');
  if (/^[\x20-\x7e]*$/.test(s)) return enc.encode(s);
  const out = [0xfe, 0xff];
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); out.push(c >> 8, c & 255); }
  return Uint8Array.from(out);
}

// ───────────── 쓰기 ─────────────
class Builder {
  constructor() { this.objs = [null]; this.maps = new Map(); this.pages = []; this.pagesNum = this.alloc(); this.forms = new Map(); this.fonts = {}; this.gs = new Map(); }
  alloc() { this.objs.push(undefined); return this.objs.length - 1; }
  put(v) { const n = this.alloc(); this.objs[n] = v; return new Ref(n, 0); }
  map(doc) { let m = this.maps.get(doc); if (!m) { m = new Map(); this.maps.set(doc, m); } return m; }
  /** 원본 개체 복사 (참조는 새 번호로). 남기지 않는 페이지 · 페이지 나무를 가리키면 null */
  copy(doc, v) {
    if (v instanceof Ref) {
      const m = this.map(doc);
      if (m.has(v.num)) { const n = m.get(v.num); return n == null ? null : new Ref(n, 0); }
      const o = doc.get(v);
      const t = isDict(o) ? doc.val(o.Type)?.n : null;
      if (t === 'Page' || t === 'Pages') { m.set(v.num, null); return null; }
      const n = this.alloc();
      m.set(v.num, n);
      this.objs[n] = this.direct(doc, o);
      return new Ref(n, 0);
    }
    return this.direct(doc, v);
  }
  direct(doc, v) {
    if (v instanceof Stream) {
      const dict = {};
      for (const k in v.dict) if (k !== 'Length') dict[k] = this.copy(doc, v.dict[k]);
      let data = v.bytes.subarray(v.start, v.end);
      if (v.crypt) data = doc.decryptBytes(data, v.crypt[0], v.crypt[1], doc.crypt.stm);
      return { stream: true, dict, data };
    }
    if (Array.isArray(v)) return v.map((x) => this.copy(doc, x));
    if (v instanceof Cmd) return null;
    if (isDict(v)) { const o = {}; for (const k in v) o[k] = this.copy(doc, v[k]); return o; }
    return v;
  }
  /** 새 PDF 의 개체 값 (참조면 그 개체) */
  deref(v) { return v instanceof Ref ? this.objs[v.num] : v; }
  /** 남길 페이지를 먼저 번호 매김 (링크 · 책갈피가 서로 가리킬 수 있게) */
  reserve(src, pages) { const m = this.map(src.doc); for (const p of pages) if (p.ref && !m.has(p.ref.num)) m.set(p.ref.num, this.alloc()); }
  /** 데이터 스트림 개체 */
  stream(dict, data) { return this.put({ stream: true, dict, data: typeof data === 'string' ? enc.encode(data) : data, compress: true }); }
  /** 원본 페이지 → 새 페이지 */
  addPage(src, p, { rotate = 0, annots = null, under = '', over = '', res = null } = {}) {
    const { doc } = src;
    const m = this.map(doc);
    let n = p.ref ? m.get(p.ref.num) : null;
    if (n == null || this.objs[n] !== undefined) { n = this.alloc(); if (p.ref && !m.has(p.ref.num)) m.set(p.ref.num, n); }
    const d = { Type: nm('Page'), Parent: new Ref(this.pagesNum, 0) };
    const SKIP = new Set(['Type', 'Parent', 'Resources', 'MediaBox', 'CropBox', 'Rotate', 'Annots', 'B', 'StructParents', 'Contents', 'Thumb', 'PieceInfo', 'ID']);
    for (const k in p.node) if (!SKIP.has(k)) d[k] = this.copy(doc, p.node[k]);
    d.MediaBox = this.copy(doc, p.mediaBox) ?? [0, 0, 612, 792];
    if (p.cropBox) d.CropBox = this.copy(doc, p.cropBox);
    const r = (((p.rotate + rotate) % 360) + 360) % 360;
    if (r) d.Rotate = r;
    let resources = this.copy(doc, p.res) ?? {};
    if (res) {
      // 쓰는 자원(글꼴 · 그림 · 투명도)을 더함 — 여러 페이지가 함께 쓰는 원본 자원 사전은 복사본을 고침
      const own = { ...(this.deref(resources) ?? {}) };
      for (const cat of Object.keys(res)) own[cat] = { ...(this.deref(own[cat]) ?? {}), ...res[cat] };
      resources = own;
    }
    d.Resources = resources;
    // 주석
    const list = doc.val(p.node.Annots);
    if (Array.isArray(list) && list.length) {
      const keep = [];
      const dropped = new Set();
      for (const a of list) {
        const ad = doc.val(a);
        if (!isDict(ad)) continue;
        const sub = doc.val(ad.Subtype)?.n ?? '';
        if (annots && !annots(sub, ad)) { if (a instanceof Ref) dropped.add(a.num); continue; }
        keep.push(a);
      }
      const out = [];
      for (const a of keep) {
        const ad = doc.val(a);
        // 지운 주석에 딸린 팝업도 지움
        if (doc.val(ad.Subtype)?.n === 'Popup' && ad.Parent instanceof Ref && dropped.has(ad.Parent.num)) continue;
        const c = this.copy(doc, a);
        if (c) out.push(c);
      }
      if (out.length) d.Annots = out;
    }
    // 내용
    let contents = p.node.Contents;
    const arr = Array.isArray(doc.val(contents)) ? doc.val(contents) : contents ? [contents] : [];
    let copied = arr.map((x) => this.copy(doc, x)).filter(Boolean);
    if (under || over) {
      const parts = [];
      if (under) parts.push(this.stream({}, `${under}\n`));
      parts.push(this.stream({}, 'q\n'), ...copied, this.stream({}, '\nQ\n'));
      if (over) parts.push(this.stream({}, `${over}\n`));
      copied = parts;
    }
    if (copied.length) d.Contents = copied.length === 1 ? copied[0] : copied;
    this.objs[n] = d;
    this.pages.push(n);
    return n;
  }
  /** 빈 페이지 (보이는 크기 pt) + 그릴 내용 */
  addBlank(w, h, content = '', res = null) {
    const n = this.alloc();
    const d = { Type: nm('Page'), Parent: new Ref(this.pagesNum, 0), MediaBox: [0, 0, w, h], Resources: res ?? {} };
    if (content) d.Contents = this.stream({}, content);
    this.objs[n] = d;
    this.pages.push(n);
    return n;
  }
  /** 원본 페이지를 그림처럼 쓰는 Form XObject (모아찍기 · 배경 겹치기) */
  formOf(src, p) {
    const key = `${this.srcId(src)}:${p.index}`;
    if (this.forms.has(key)) return this.forms.get(key);
    const { doc } = src;
    const cs = doc.val(p.node.Contents);
    const parts = (Array.isArray(cs) ? cs : cs ? [p.node.Contents] : []).map((c) => doc.val(c)).filter((s) => s instanceof Stream);
    const chunks = [];
    for (const s of parts) { try { chunks.push(doc.decode(s), enc.encode('\n')); } catch { /* 읽을 수 없는 조각은 건너뜀 */ } }
    const total = chunks.reduce((a, c) => a + c.length, 0);
    const data = new Uint8Array(total); let k = 0; for (const c of chunks) { data.set(c, k); k += c.length; }
    const dict = { Type: nm('XObject'), Subtype: nm('Form'), BBox: p.box.slice(), Resources: this.copy(doc, p.res) ?? {} };
    if (p.node.Group) dict.Group = this.copy(doc, p.node.Group);
    const ref = this.put({ stream: true, dict, data, compress: true });
    this.forms.set(key, ref);
    return ref;
  }
  srcId(src) { this.ids ??= new Map(); if (!this.ids.has(src)) this.ids.set(src, this.ids.size); return this.ids.get(src); }
  /** 글꼴: 'helv' (Helvetica, ASCII) · 'kor' (HYGoThic-Medium, 한글 포함 모든 글자) */
  font(kind) {
    if (this.fonts[kind]) return this.fonts[kind];
    if (kind === 'helv') this.fonts[kind] = this.put({ Type: nm('Font'), Subtype: nm('Type1'), BaseFont: nm('Helvetica'), Encoding: nm('WinAnsiEncoding') });
    else {
      const desc = this.put({ Type: nm('FontDescriptor'), FontName: nm('HYGoThic-Medium'), Flags: 6, FontBBox: [-100, -142, 1000, 880], ItalicAngle: 0, Ascent: 880, Descent: -120, CapHeight: 700, StemV: 80 });
      const cid = this.put({ Type: nm('Font'), Subtype: nm('CIDFontType0'), BaseFont: nm('HYGoThic-Medium'), CIDSystemInfo: { Registry: enc.encode('Adobe'), Ordering: enc.encode('Korea1'), Supplement: 1 }, FontDescriptor: desc, DW: 1000, W: [1, 95, 500] });
      this.fonts[kind] = this.put({ Type: nm('Font'), Subtype: nm('Type0'), BaseFont: nm('HYGoThic-Medium-UniKS-UCS2-H'), Encoding: nm('UniKS-UCS2-H'), DescendantFonts: [cid] });
    }
    return this.fonts[kind];
  }
  /** TrueType 글꼴 넣기 (쓴 글자만, finish 에서 채움) → { ref, font, used } */
  embed(font) {
    this.embedded ??= [];
    const e = { ref: new Ref(this.alloc(), 0), font, used: new Map() };
    this.embedded.push(e);
    return e;
  }
  async fillFonts() {
    const hex4 = (c) => c.toString(16).padStart(4, '0');
    for (const [k, e] of (this.embedded ?? []).entries()) {
      const f = e.font;
      if (!e.used.size) e.used.set(0, ' ');
      const sub = subsetTtf(f, [...e.used.keys()]);
      const tag = `${String.fromCharCode(...Array.from({ length: 6 }, (_, i) => 65 + ((k * 7 + i * 5 + 3) % 26)))}+${f.psName}`;
      const ff = this.put({ stream: true, dict: { Length1: sub.length }, data: sub, compress: true });
      const desc = this.put({ Type: nm('FontDescriptor'), FontName: nm(tag), Flags: 4, FontBBox: f.bbox, ItalicAngle: 0, Ascent: f.ascent, Descent: f.descent, CapHeight: f.capHeight, StemV: 80, FontFile2: ff });
      const gids = [...e.used.keys()].sort((a, b) => a - b);
      const W = []; for (const g of gids) W.push(g, [f.width(g)]);
      const cid = this.put({ Type: nm('Font'), Subtype: nm('CIDFontType2'), BaseFont: nm(tag), CIDSystemInfo: { Registry: enc.encode('Adobe'), Ordering: enc.encode('Identity'), Supplement: 0 }, FontDescriptor: desc, DW: 1000, W, CIDToGIDMap: nm('Identity') });
      const u16 = (str) => Array.from(str, (ch) => { const c = ch.codePointAt(0); if (c > 0xffff) { const v = c - 0x10000; return hex4(0xd800 + (v >> 10)) + hex4(0xdc00 + (v & 0x3ff)); } return hex4(c); }).join('');
      const lines = [];
      for (let i = 0; i < gids.length; i += 100) { const part = gids.slice(i, i + 100); lines.push(`${part.length} beginbfchar`, ...part.map((g) => `<${hex4(g)}> <${u16(e.used.get(g) || ' ')}>`), 'endbfchar'); }
      const tu = this.stream({}, `/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${lines.join('\n')}\nendcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`);
      this.objs[e.ref.num] = { Type: nm('Font'), Subtype: nm('Type0'), BaseFont: nm(tag), Encoding: nm('Identity-H'), DescendantFonts: [cid], ToUnicode: tu };
    }
  }
  gstate(alpha) {
    const k = Math.round(alpha * 100);
    if (!this.gs.has(k)) this.gs.set(k, this.put({ Type: nm('ExtGState'), ca: k / 100, CA: k / 100 }));
    return this.gs.get(k);
  }
  /** 책갈피 [{ title, page (새 페이지 순서), y?, level }] */
  outline(items) {
    if (!items?.length) return null;
    const root = this.alloc();
    const nodes = items.map((it) => ({ it, n: this.alloc(), kids: [] }));
    const top = [];
    const stack = [];
    for (const nd of nodes) {
      while (stack.length && stack[stack.length - 1].it.level >= nd.it.level) stack.pop();
      (stack.length ? stack[stack.length - 1].kids : top).push(nd);
      nd.parent = stack.length ? stack[stack.length - 1].n : root;
      stack.push(nd);
    }
    const link = (list) => list.forEach((nd, i) => {
      const pn = this.pages[nd.it.page];
      const o = { Title: textBytes(nd.it.title), Parent: new Ref(nd.parent, 0) };
      if (pn != null) o.Dest = [new Ref(pn, 0), nm('XYZ'), null, nd.it.y ?? null, null];
      if (i > 0) o.Prev = new Ref(list[i - 1].n, 0);
      if (i < list.length - 1) o.Next = new Ref(list[i + 1].n, 0);
      if (nd.kids.length) { o.First = new Ref(nd.kids[0].n, 0); o.Last = new Ref(nd.kids[nd.kids.length - 1].n, 0); o.Count = -nd.kids.length; link(nd.kids); }
      this.objs[nd.n] = o;
    });
    link(top);
    this.objs[root] = { Type: nm('Outlines'), First: new Ref(top[0].n, 0), Last: new Ref(top[top.length - 1].n, 0), Count: top.length };
    return new Ref(root, 0);
  }
  /** → Uint8Array */
  async finish({ catalog = {}, info = {}, compressAll = false } = {}) {
    await this.fillFonts();
    this.objs[this.pagesNum] = { Type: nm('Pages'), Kids: this.pages.map((n) => new Ref(n, 0)), Count: this.pages.length };
    const cat = this.put({ Type: nm('Catalog'), Pages: new Ref(this.pagesNum, 0), ...catalog });
    const now = new Date();
    const pad = (v) => String(v).padStart(2, '0');
    const date = `D:${now.getUTCFullYear()}${pad(now.getUTCMonth() + 1)}${pad(now.getUTCDate())}${pad(now.getUTCHours())}${pad(now.getUTCMinutes())}${pad(now.getUTCSeconds())}Z`;
    const inf = this.put({ ...info, Producer: textBytes('WIWORD PDF 도구'), ModDate: enc.encode(date) });
    // 압축: 새로 만든 스트림, 그리고 (정리할 때) 압축 안 된 원본 스트림
    for (let i = 1; i < this.objs.length; i++) {
      const o = this.objs[i];
      if (!o?.stream) continue;
      const plain = !o.dict.Filter;
      if ((o.compress || (compressAll && plain)) && plain && o.data.length > 64) {
        const z = await deflate(o.data);
        if (z && z.length < o.data.length) { o.data = z; o.dict.Filter = nm('FlateDecode'); }
      }
    }
    const chunks = [enc.encode('%PDF-1.7\n'), Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a)];
    let pos = chunks[0].length + chunks[1].length;
    const offs = new Array(this.objs.length).fill(0);
    for (let i = 1; i < this.objs.length; i++) {
      let o = this.objs[i];
      if (o === undefined) o = null;
      offs[i] = pos;
      let head; let data = null;
      if (o?.stream) { head = `${i} 0 obj\n${ser({ ...o.dict, Length: o.data.length })}\nstream\n`; data = o.data; } else head = `${i} 0 obj\n${ser(o)}\nendobj\n`;
      const hb = enc.encode(head);
      chunks.push(hb); pos += hb.length;
      if (data) { const tail = enc.encode('\nendstream\nendobj\n'); chunks.push(data, tail); pos += data.length + tail.length; }
    }
    const id = Array.from({ length: 16 }, () => Math.floor(Math.random() * 256));
    const idHex = id.map((b) => b.toString(16).padStart(2, '0')).join('');
    let xref = `xref\n0 ${this.objs.length}\n0000000000 65535 f \n`;
    for (let i = 1; i < this.objs.length; i++) xref += `${String(offs[i]).padStart(10, '0')} 00000 n \n`;
    xref += `trailer\n<</Size ${this.objs.length}/Root ${cat.num} 0 R/Info ${inf.num} 0 R/ID [<${idHex}><${idHex}>]>>\nstartxref\n${pos}\n%%EOF\n`;
    chunks.push(enc.encode(xref));
    const total = chunks.reduce((a, c) => a + c.length, 0);
    const out = new Uint8Array(total); let k = 0;
    for (const c of chunks) { out.set(c, k); k += c.length; }
    return out;
  }
}

/** 값 → PDF 문법 */
function ser(v) {
  if (v == null) return 'null';
  if (typeof v === 'number') return num(v);
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  if (v instanceof Ref) return `${v.num} 0 R`;
  if (v instanceof Name) return `/${[...v.n].map((ch) => { const c = ch.charCodeAt(0); return c > 0x20 && c < 0x7f && !'()<>[]{}/%#'.includes(ch) ? ch : `#${(c & 255).toString(16).padStart(2, '0')}`; }).join('')}`;
  if (v instanceof Uint8Array) {
    if (v.every((c) => c >= 0x20 && c < 0x7f)) return `(${latin1.decode(v).replace(/[\\()]/g, (m) => `\\${m}`)})`;
    let h = '<'; for (const c of v) h += c.toString(16).padStart(2, '0'); return `${h}>`;
  }
  if (Array.isArray(v)) return `[${v.map(ser).join(' ')}]`;
  if (typeof v === 'object') return `<<${Object.keys(v).filter((k) => v[k] !== undefined).map((k) => `${ser(nm(k))} ${ser(v[k])}`).join(' ')}>>`;
  return 'null';
}

// ───────────── 책갈피 읽기 ─────────────
function nameTree(doc, node, key, depth = 0) {
  node = doc.val(node);
  if (!isDict(node) || depth > 30) return null;
  const names = doc.val(node.Names);
  if (Array.isArray(names)) for (let i = 0; i + 1 < names.length; i += 2) if (pdfText(doc.val(names[i])) === key) return doc.val(names[i + 1]);
  for (const k of doc.val(node.Kids) ?? []) { const r = nameTree(doc, k, key, depth + 1); if (r != null) return r; }
  return null;
}
function destOf(doc, root, d) {
  d = doc.val(d);
  if (d instanceof Uint8Array || d instanceof Name) {
    const key = d instanceof Name ? d.n : pdfText(d);
    d = doc.val(doc.val(root.Dests)?.[key]) ?? nameTree(doc, doc.val(root.Names)?.Dests, key);
    d = doc.val(d);
  }
  if (isDict(d)) d = doc.val(d.D);
  return Array.isArray(d) ? d : null;
}
/** 원본 책갈피 → [{ title, page (원본 페이지 번호), y, level }] */
export function readOutline(src) {
  const { doc, root, pages } = src;
  const byNum = new Map(pages.map((p, i) => [p.ref?.num, i]));
  const out = [];
  const seen = new Set();
  const walk = (ref, level) => {
    let guard = 0;
    while (ref && guard++ < 5000) {
      if (ref instanceof Ref) { if (seen.has(ref.num)) return; seen.add(ref.num); }
      const it = doc.val(ref);
      if (!isDict(it)) return;
      let dest = it.Dest ? destOf(doc, root, it.Dest) : null;
      const a = doc.val(it.A);
      if (!dest && isDict(a) && doc.val(a.S)?.n === 'GoTo') dest = destOf(doc, root, a.D);
      const pr = dest?.[0];
      const page = pr instanceof Ref ? byNum.get(pr.num) : typeof doc.val(pr) === 'number' ? doc.val(pr) : undefined;
      const y = dest && doc.val(dest[1])?.n === 'XYZ' && typeof doc.val(dest[3]) === 'number' ? doc.val(dest[3]) : null;
      out.push({ title: pdfText(doc.val(it.Title)).replace(/\0/g, ''), page, y, level });
      if (it.First) walk(it.First, level + 1);
      ref = it.Next;
    }
  };
  const ol = doc.val(root.Outlines);
  if (isDict(ol) && ol.First) walk(ol.First, 0);
  return out;
}
/** 원본 책갈피를 새 페이지 순서로 옮김 (남지 않는 페이지를 가리키는 것은 뺌) */
function remapOutline(items, newIndexOf, levelAdd = 0) {
  return items.filter((it) => it.page != null && newIndexOf.has(it.page)).map((it) => ({ ...it, page: newIndexOf.get(it.page), level: it.level + levelAdd }));
}

/** 원본 문서 사전 중 새 문서에 옮길 것 */
function catalogFrom(b, src, { labels = false } = {}) {
  const { doc, root } = src;
  const out = {};
  for (const k of ['Lang', 'ViewerPreferences', 'PageLayout', 'OCProperties', 'AcroForm', 'Metadata']) if (root[k] != null) out[k] = b.copy(doc, root[k]);
  if (labels && root.PageLabels) out.PageLabels = b.copy(doc, root.PageLabels);
  const names = doc.val(root.Names);
  if (isDict(names)) {
    const n = {};
    for (const k of ['Dests', 'EmbeddedFiles']) if (names[k]) n[k] = b.copy(doc, names[k]);
    if (Object.keys(n).length) out.Names = n;
  }
  if (root.Dests) out.Dests = b.copy(doc, root.Dests);
  return out;
}
function infoFrom(b, src, extra = {}) {
  const out = {};
  const { doc, info } = src;
  if (isDict(info)) for (const k of ['Title', 'Author', 'Subject', 'Keywords', 'Creator', 'CreationDate']) if (info[k] != null) out[k] = b.copy(doc, info[k]);
  for (const [k, v] of Object.entries(extra)) if (v != null) out[k] = v === '' ? undefined : textBytes(v);
  return out;
}

// ───────────── 페이지 지정 ─────────────
/** "1-3, 5, 8-" → [0,1,2,4,7…] (1부터, 범위 밖은 버림). 빈 글 → 모든 페이지 */
export function parsePages(spec, count) {
  const s = String(spec ?? '').trim();
  if (!s || s === '*' || /^(전체|모두|all)$/i.test(s)) return Array.from({ length: count }, (_, i) => i);
  const out = [];
  for (const part of s.split(/[,;\s]+/).filter(Boolean)) {
    const m = /^(\d+|끝|end|z)?\s*(?:[-~]\s*(\d+|끝|end|z)?)?$/i.exec(part);
    if (!m) throw new Error(`페이지 지정을 읽을 수 없습니다: '${part}' (예: 1-3, 5, 8-)`);
    const v = (t, d) => (t == null || t === '' ? d : /^(끝|end|z)$/i.test(t) ? count : Number(t));
    const a = v(m[1], 1); const hasRange = part.includes('-') || part.includes('~');
    const z = hasRange ? v(m[2], count) : a;
    const step = a <= z ? 1 : -1;
    for (let i = a; step > 0 ? i <= z : i >= z; i += step) if (i >= 1 && i <= count) out.push(i - 1);
  }
  return out;
}

// ───────────── 도구 ─────────────
/** 페이지 정보 */
export function pdfInfo(bytes, opts = {}) {
  const src = openPdf(bytes, opts);
  const info = {};
  if (isDict(src.info)) for (const k of ['Title', 'Author', 'Subject', 'Keywords', 'Creator', 'Producer']) if (src.info[k] != null) info[k.toLowerCase()] = pdfText(src.doc.val(src.info[k]));
  const annots = {};
  for (const p of src.pages) for (const a of src.doc.val(p.node.Annots) ?? []) { const t = src.doc.val(src.doc.val(a)?.Subtype)?.n ?? '?'; annots[t] = (annots[t] ?? 0) + 1; }
  return { pages: src.pages.map((p) => ({ w: Math.round(p.w), h: Math.round(p.h), rotate: p.rotate })), encrypted: !!src.doc.trailer.Encrypt, info, outline: readOutline(src).length, annots };
}

/** 고른 페이지로 새 PDF (추출 · 삭제 · 순서 바꾸기 · 역순 · 회전 · 암호 해제 · 정리 공통) */
export async function selectPages(bytes, { pages = null, rotate = null, password, keepOutline = true, compress = false, info = null } = {}) {
  const src = openPdf(bytes, { password });
  const order = pages ?? src.pages.map((_, i) => i);
  if (!order.length) throw new Error('남길 페이지가 없습니다.');
  const b = new Builder();
  const chosen = order.map((i) => src.pages[i]).filter(Boolean);
  b.reserve(src, chosen);
  const newIndex = new Map();
  order.forEach((i, k) => { if (!newIndex.has(i)) newIndex.set(i, k); b.addPage(src, src.pages[i], { rotate: rotate?.[k] ?? rotate?.all ?? 0 }); });
  const identity = order.length === src.pages.length && order.every((v, i) => v === i);
  const catalog = catalogFrom(b, src, { labels: identity });
  const ol = keepOutline ? b.outline(remapOutline(readOutline(src), newIndex)) : null;
  if (ol) { catalog.Outlines = ol; catalog.PageMode = nm('UseOutlines'); }
  return b.finish({ catalog, info: infoFrom(b, src, info ?? {}), compressAll: compress });
}
export const extractPages = (bytes, spec, opts = {}) => { const n = openPdf(bytes, opts).pages.length; return selectPages(bytes, { ...opts, pages: parsePages(spec, n) }); };
export const deletePages = (bytes, spec, opts = {}) => {
  const n = openPdf(bytes, opts).pages.length;
  const del = new Set(parsePages(spec, n));
  const pages = Array.from({ length: n }, (_, i) => i).filter((i) => !del.has(i));
  return selectPages(bytes, { ...opts, pages });
};
export const reversePages = (bytes, opts = {}) => { const n = openPdf(bytes, opts).pages.length; return selectPages(bytes, { ...opts, pages: Array.from({ length: n }, (_, i) => n - 1 - i) }); };
/** 고른 페이지만 deg(90 · 180 · 270) 만큼 시계 방향으로 돌림 */
export const rotatePages = (bytes, deg, spec = '', opts = {}) => {
  const n = openPdf(bytes, opts).pages.length;
  const sel = new Set(parsePages(spec, n));
  return selectPages(bytes, { ...opts, rotate: Array.from({ length: n }, (_, i) => (sel.has(i) ? deg : 0)) });
};

/** 여러 PDF 병합 [{ bytes, name?, pages?(지정 글), password? }] → 파일마다 책갈피 하나 + 그 아래 원래 책갈피 */
export async function mergePdfs(files, { bookmarks = true, padEven = 0 } = {}) {
  if (!files?.length) throw new Error('합칠 PDF 가 없습니다.');
  const srcs = files.map((f) => ({ f, src: openPdf(f.bytes, { password: f.password }) }));
  const b = new Builder();
  const outline = [];
  const fields = [];
  let catalog = {};
  for (const [k, { f, src }] of srcs.entries()) {
    const order = parsePages(f.pages, src.pages.length);
    b.reserve(src, order.map((i) => src.pages[i]));
    const start = b.pages.length;
    const newIndex = new Map();
    for (const i of order) { if (!newIndex.has(i)) newIndex.set(i, b.pages.length); b.addPage(src, src.pages[i]); }
    // 양면 인쇄 맞추기: 파일마다 padEven 의 배수로 빈 페이지를 채움
    if (padEven > 1) {
      const last = src.pages[order[order.length - 1]] ?? src.pages[0];
      while ((b.pages.length - start) % padEven) b.addBlank(last.w, last.h);
    }
    if (k === 0) catalog = catalogFrom(b, src);
    const af = src.doc.val(src.root.AcroForm);
    if (k > 0 && isDict(af)) for (const fl of src.doc.val(af.Fields) ?? []) { const c = b.copy(src.doc, fl); if (c) fields.push(c); }
    if (bookmarks) {
      outline.push({ title: f.name ?? `문서 ${k + 1}`, page: start, level: 0 });
      outline.push(...remapOutline(readOutline(src), newIndex, 1));
    }
  }
  if (fields.length && catalog.AcroForm) { const af = { ...(b.deref(catalog.AcroForm) ?? {}) }; af.Fields = [...(b.deref(af.Fields) ?? []), ...fields]; catalog.AcroForm = af; }
  delete catalog.PageLabels;
  const ol = bookmarks ? b.outline(outline) : null;
  if (ol) { catalog.Outlines = ol; catalog.PageMode = nm('UseOutlines'); }
  return b.finish({ catalog, info: infoFrom(b, srcs[0].src, files.length > 1 ? { Title: '' } : {}) });
}

/** 양면 인쇄 맞추기: 파일마다 앞면에서 시작하도록 빈 페이지를 넣어 합침 (perSheet = 모아찍기 칸 수: 장 단위로 맞춤) */
export async function duplexMerge(files, { perSheet = 1, bookmarks = true } = {}) {
  const merged = await mergePdfs(files, { bookmarks, padEven: 2 * Math.max(1, perSheet) });
  return perSheet > 1 ? nUp(merged, { n: perSheet }) : merged;
}

/** 분할: every = N쪽마다, 또는 ranges = '1-3, 4-6' (쉼표마다 파일 하나) → [{ name, bytes }] */
export async function splitPdf(bytes, { every = 1, ranges = null, password, base = 'part' } = {}) {
  const src = openPdf(bytes, { password });
  const n = src.pages.length;
  const groups = [];
  if (ranges) for (const r of String(ranges).split(/[,;]+/).map((x) => x.trim()).filter(Boolean)) groups.push({ label: r.replace(/\s+/g, ''), pages: parsePages(r, n) });
  else for (let i = 0; i < n; i += Math.max(1, every)) { const pages = []; for (let k = i; k < Math.min(n, i + Math.max(1, every)); k++) pages.push(k); groups.push({ label: pages.length > 1 ? `${i + 1}-${pages[pages.length - 1] + 1}` : `${i + 1}`, pages }); }
  const out = [];
  for (const g of groups.filter((x) => x.pages.length)) out.push({ name: `${base}_${g.label}.pdf`, bytes: await selectPages(bytes, { pages: g.pages, password }) });
  return out;
}

/** 칸 배치: 2 = 2×1, 4 = 2×2, 6 = 3×2, 8 = 4×2, 9 = 3×3, 16 = 4×4 (가로 × 세로, 가로 용지 기준) */
const NUP = { 2: [2, 1], 4: [2, 2], 6: [3, 2], 8: [4, 2], 9: [3, 3], 16: [4, 4] };
/** 원본 페이지를 rect(보이는 좌표) 안에 맞춰 그리는 행렬 */
function placeMatrix(p, [x, y, w, h], stretch = false) {
  const [x0, y0, x1, y1] = p.box; const W = x1 - x0; const H = y1 - y0;
  const VW = p.rotate % 180 ? H : W; const VH = p.rotate % 180 ? W : H;
  const T0 = [1, 0, 0, 1, -x0, -y0];
  const R = rotM(p.rotate, W, H);
  let S;
  if (stretch) S = [w / VW, 0, 0, h / VH, x, y];
  else { const s = Math.min(w / VW, h / VH); S = [s, 0, 0, s, x + (w - VW * s) / 2, y + (h - VH * s) / 2]; }
  return mul(mul(T0, R), S);
}
/** 보이는 좌표 → 그 페이지의 사용자 좌표 */
const visualToUser = (p) => { const [x0, y0, x1, y1] = p.box; return inv(mul([1, 0, 0, 1, -x0, -y0], rotM(p.rotate, x1 - x0, y1 - y0))); };

/** 모아찍기: 한 장에 n 페이지 */
export async function nUp(bytes, { n = 2, paper = 'A4', margin = 18, gap = 8, border = false, order = 'row', password } = {}) {
  const src = openPdf(bytes, { password });
  const g = NUP[n];
  if (!g) throw new Error('모아찍기는 2 · 4 · 6 · 8 · 9 · 16 쪽만 됩니다.');
  const first = src.pages[0];
  const base = PAPERS[paper] ?? [first.w, first.h].sort((a, b) => a - b);
  // 가로/세로 중 칸이 큰 쪽
  const fit = (cols, rows, [pw, ph]) => Math.min((pw - 2 * margin - (cols - 1) * gap) / cols / first.w, (ph - 2 * margin - (rows - 1) * gap) / rows / first.h);
  const land = [base[1], base[0]]; const port = [base[0], base[1]];
  const opts = [[g[0], g[1], land], [g[1], g[0], port]];
  const [cols, rows, [PW, PH]] = opts.sort((a, b) => fit(b[0], b[1], b[2]) - fit(a[0], a[1], a[2]))[0];
  const b = new Builder();
  const cw = (PW - 2 * margin - (cols - 1) * gap) / cols; const ch = (PH - 2 * margin - (rows - 1) * gap) / rows;
  for (let s = 0; s < src.pages.length; s += n) {
    const res = { XObject: {} };
    let content = '';
    for (let k = 0; k < n && s + k < src.pages.length; k++) {
      const p = src.pages[s + k];
      const c = order === 'col' ? Math.floor(k / rows) : k % cols; const r = order === 'col' ? k % rows : Math.floor(k / cols);
      const rect = [margin + c * (cw + gap), PH - margin - (r + 1) * ch - r * gap, cw, ch];
      const xo = b.formOf(src, p);
      const name = `P${k}`;
      res.XObject[name] = xo;
      content += `q ${mtx(placeMatrix(p, rect))} cm /${name} Do Q\n`;
      if (border) { const m = placeMatrix(p, rect); const [x0, y0, x1, y1] = p.box; const W = x1 - x0; const H = y1 - y0; content += `q ${mtx(mul([1, 0, 0, 1, -x0, -y0], m))} cm 0.5 0.5 0.5 RG 0.5 w ${num(x0)} ${num(y0)} ${num(W)} ${num(H)} re S Q\n`; }
    }
    b.addBlank(PW, PH, content, res);
  }
  return b.finish({ catalog: {}, info: infoFrom(b, src) });
}

/** 배경 겹치기: bg PDF 의 한 페이지를 고른 페이지의 뒤(또는 앞)에 그림 — 돌린 페이지에서도 똑바로 */
export async function overlayPdf(bytes, bgBytes, { bgPage = 1, behind = true, pages = '', stretch = false, password, bgPassword } = {}) {
  const src = openPdf(bytes, { password });
  const bg = openPdf(bgBytes, { password: bgPassword });
  const bp = bg.pages[Math.max(0, Math.min(bg.pages.length - 1, bgPage - 1))];
  const sel = new Set(parsePages(pages, src.pages.length));
  const b = new Builder();
  b.reserve(src, src.pages);
  const xo = b.formOf(bg, bp);
  const newIndex = new Map();
  src.pages.forEach((p, i) => {
    newIndex.set(i, i);
    if (!sel.has(i)) { b.addPage(src, p); return; }
    const m = mul(placeMatrix(bp, [0, 0, p.w, p.h], stretch), visualToUser(p));
    const draw = `q ${mtx(m)} cm /WxBg Do Q`;
    b.addPage(src, p, behind ? { under: draw, res: { XObject: { WxBg: xo } } } : { over: draw, res: { XObject: { WxBg: xo } } });
  });
  const catalog = catalogFrom(b, src, { labels: true });
  const ol = b.outline(remapOutline(readOutline(src), newIndex));
  if (ol) catalog.Outlines = ol;
  return b.finish({ catalog, info: infoFrom(b, src) });
}

/** 주석 종류 (지우기 화면의 선택지) */
export const ANNOT_TYPES = [
  ['Text', '메모'], ['FreeText', '글 상자'], ['Highlight', '형광펜'], ['Underline', '밑줄'], ['StrikeOut', '취소선'], ['Squiggly', '물결 밑줄'],
  ['Ink', '자유형(펜)'], ['Line', '선'], ['Square', '사각형'], ['Circle', '원'], ['Polygon', '다각형'], ['PolyLine', '꺾은선'],
  ['Stamp', '스탬프'], ['Caret', '삽입 표시'], ['FileAttachment', '첨부 파일'], ['Sound', '소리'], ['Redact', '가림 표시'],
  ['Link', '링크'], ['Widget', '양식 입력란'],
];
/** 주석 한꺼번에 지우기: types = 지울 종류 (기본: 링크 · 입력란을 뺀 모두), pages = 지정 글 */
export async function removeAnnotations(bytes, { types = null, pages = '', password } = {}) {
  const src = openPdf(bytes, { password });
  const sel = new Set(parsePages(pages, src.pages.length));
  const drop = types ? new Set(types) : null;
  const b = new Builder();
  b.reserve(src, src.pages);
  let removed = 0;
  const newIndex = new Map();
  src.pages.forEach((p, i) => {
    newIndex.set(i, i);
    const test = sel.has(i) ? (sub) => { const del = drop ? drop.has(sub) : sub !== 'Link' && sub !== 'Widget'; if (del && sub !== 'Popup') removed++; return !del; } : null;
    b.addPage(src, p, { annots: test });
  });
  const catalog = catalogFrom(b, src, { labels: true });
  const ol = b.outline(remapOutline(readOutline(src), newIndex));
  if (ol) catalog.Outlines = ol;
  const out = await b.finish({ catalog, info: infoFrom(b, src) });
  out.removed = removed;
  return out;
}

// ───────────── 빈 페이지 찾기 ─────────────
const WHITE = (cs) => cs.every((v) => v >= 0.98);
/** 페이지마다 { index, blank, imageOnly, text, paint, images } — imageOnly 페이지는 그림을 보고 판단해야 함 (스캔한 빈 종이) */
export function analyzePages(bytes, { password } = {}) {
  const src = openPdf(bytes, { password });
  const { doc } = src;
  const scan = (data, res, st, depth) => {
    const lx = new Lexer(data);
    const stack = [];
    let fill = [0]; let stroke = [0]; let tr = 0;
    const gsStack = [];
    for (let guard = 0; guard < 2e6; guard++) {
      const t = lx.next();
      if (t === undefined) break;
      if (!(t instanceof Cmd)) { stack.push(t); continue; }
      const op = t.c;
      const nums = () => stack.filter((x) => typeof x === 'number');
      switch (op) {
        case 'q': gsStack.push([fill, stroke, tr]); break;
        case 'Q': if (gsStack.length) [fill, stroke, tr] = gsStack.pop(); break;
        case 'g': fill = nums(); break;
        case 'G': stroke = nums(); break;
        case 'rg': fill = nums(); break;
        case 'RG': stroke = nums(); break;
        case 'k': { const v = nums(); fill = v.length === 4 && v.every((x) => x <= 0.02) ? [1] : [0]; break; }
        case 'K': { const v = nums(); stroke = v.length === 4 && v.every((x) => x <= 0.02) ? [1] : [0]; break; }
        case 'sc': case 'scn': fill = nums().length ? nums() : [0]; break;
        case 'SC': case 'SCN': stroke = nums().length ? nums() : [0]; break;
        case 'cs': fill = [0]; break;
        case 'CS': stroke = [0]; break;
        case 'Tr': tr = nums()[0] ?? 0; break;
        case 'Tj': case "'": case '"': case 'TJ': {
          if (tr === 3 || tr === 7) break;
          const strs = op === 'TJ' ? (Array.isArray(stack[stack.length - 1]) ? stack[stack.length - 1] : []).filter((x) => x instanceof Uint8Array) : stack.filter((x) => x instanceof Uint8Array);
          const n = strs.reduce((a, s) => a + s.length, 0);
          if (n && !(tr === 0 && WHITE(fill)) ) st.text += n;
          break;
        }
        case 'f': case 'F': case 'f*': if (!WHITE(fill)) st.paint++; break;
        case 'S': case 's': if (!WHITE(stroke)) st.paint++; break;
        case 'B': case 'B*': case 'b': case 'b*': if (!WHITE(fill) || !WHITE(stroke)) st.paint++; break;
        case 'sh': st.paint++; break;
        case 'BI': {
          // 인라인 그림: ID … EI 까지 건너뜀
          st.images++;
          const b = lx.b; let p = lx.pos;
          while (p < lx.end - 2 && !(b[p] === 0x45 && b[p + 1] === 0x49 && (b[p + 2] === 0x0a || b[p + 2] === 0x0d || b[p + 2] === 0x20) && (b[p - 1] === 0x0a || b[p - 1] === 0x0d || b[p - 1] === 0x20))) p++;
          lx.pos = p + 2;
          break;
        }
        case 'Do': {
          const nmv = stack[stack.length - 1];
          const xo = nmv instanceof Name ? doc.val(doc.val(doc.val(res)?.XObject)?.[nmv.n]) : null;
          if (xo instanceof Stream) {
            const sub = doc.val(xo.dict.Subtype)?.n;
            if (sub === 'Image') st.images++;
            else if (sub === 'Form' && depth < 8) { try { scan(doc.decode(xo), xo.dict.Resources ?? res, st, depth + 1); } catch { st.paint++; } }
          }
          break;
        }
        default: break;
      }
      stack.length = 0;
    }
  };
  return src.pages.map((p, index) => {
    const st = { text: 0, paint: 0, images: 0 };
    const cs = doc.val(p.node.Contents);
    for (const c of Array.isArray(cs) ? cs : cs ? [p.node.Contents] : []) {
      const s = doc.val(c);
      if (s instanceof Stream) { try { scan(doc.decode(s), p.res, st, 0); } catch { st.paint++; } }
    }
    // 보이는 주석(메모 · 도장 등)이 있으면 빈 페이지가 아님
    const annots = (doc.val(p.node.Annots) ?? []).map((a) => doc.val(doc.val(a)?.Subtype)?.n).filter((t) => t && t !== 'Link' && t !== 'Popup');
    const blank = !st.text && !st.paint && !st.images && !annots.length;
    return { index, blank, imageOnly: !st.text && !st.paint && st.images > 0 && !annots.length, ...st, annots: annots.length };
  });
}

// ───────────── 이미지로 PDF ─────────────
function jpegInfo(b) {
  if (b[0] !== 0xff || b[1] !== 0xd8) return null;
  let p = 2; let orient = 1; let adobe = false; let w = 0; let h = 0; let comps = 3;
  while (p + 4 < b.length) {
    if (b[p] !== 0xff) { p++; continue; }
    const mk = b[p + 1];
    if (mk === 0xd8 || (mk >= 0xd0 && mk <= 0xd7) || mk === 0x01 || mk === 0xff) { p += mk === 0xff ? 1 : 2; continue; }
    const len = (b[p + 2] << 8) | b[p + 3];
    const seg = p + 4;
    if (mk === 0xe1 && b[seg] === 0x45 && b[seg + 1] === 0x78 && b[seg + 2] === 0x69 && b[seg + 3] === 0x66) {
      // Exif: TIFF 머리 → IFD0 의 Orientation(0x0112)
      const t = seg + 6; const le = b[t] === 0x49;
      const u16 = (o) => (le ? b[o] | (b[o + 1] << 8) : (b[o] << 8) | b[o + 1]);
      const u32 = (o) => (le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0);
      const ifd = t + u32(t + 4); const n = u16(ifd);
      for (let i = 0; i < n && i < 200; i++) { const e = ifd + 2 + i * 12; if (u16(e) === 0x0112) { orient = u16(e + 8) || 1; break; } }
    }
    if (mk === 0xee && b[seg] === 0x41 && b[seg + 1] === 0x64 && b[seg + 2] === 0x6f && b[seg + 3] === 0x62 && b[seg + 4] === 0x65) adobe = true;
    if ((mk >= 0xc0 && mk <= 0xcf) && mk !== 0xc4 && mk !== 0xc8 && mk !== 0xcc) { h = (b[seg + 1] << 8) | b[seg + 2]; w = (b[seg + 3] << 8) | b[seg + 4]; comps = b[seg + 5]; break; }
    p = seg + len - 2;
  }
  return w && h ? { w, h, comps, orient: orient >= 1 && orient <= 8 ? orient : 1, adobe } : null;
}
function pngDecode(b) {
  if (!(b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47)) return null;
  let p = 8; let w = 0; let h = 0; let depth = 8; let ctype = 2; let interlace = 0; let pal = null; let trns = null;
  const idat = [];
  const u32 = (o) => ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3]) >>> 0;
  while (p + 8 <= b.length) {
    const len = u32(p); const type = latin1.decode(b.subarray(p + 4, p + 8)); const d = b.subarray(p + 8, p + 8 + len);
    if (type === 'IHDR') { w = u32(p + 8); h = u32(p + 12); depth = d[8]; ctype = d[9]; interlace = d[12]; }
    else if (type === 'PLTE') pal = d.slice();
    else if (type === 'tRNS') trns = d.slice();
    else if (type === 'IDAT') idat.push(d);
    else if (type === 'IEND') break;
    p += 12 + len;
  }
  if (interlace) throw new Error('인터레이스 PNG 는 바로 넣을 수 없습니다.');
  const z = new Uint8Array(idat.reduce((a, c) => a + c.length, 0)); let k = 0; for (const c of idat) { z.set(c, k); k += c.length; }
  const raw = inflate(z.subarray(2), w * h * 8 + h);
  const ch = { 0: 1, 2: 3, 3: 1, 4: 2, 6: 4 }[ctype];
  const bpp = Math.max(1, (ch * depth) >> 3); const stride = (w * ch * depth + 7) >> 3;
  const px = new Uint8Array(stride * h);
  for (let y = 0; y < h; y++) {
    const f = raw[y * (stride + 1)]; const row = raw.subarray(y * (stride + 1) + 1, (y + 1) * (stride + 1)); const o = y * stride;
    for (let x = 0; x < stride; x++) {
      const a = x >= bpp ? px[o + x - bpp] : 0; const up = y ? px[o - stride + x] : 0; const c = x >= bpp && y ? px[o - stride + x - bpp] : 0;
      let v = row[x];
      if (f === 1) v += a; else if (f === 2) v += up; else if (f === 3) v += (a + up) >> 1;
      else if (f === 4) { const pp = a + up - c; const pa = Math.abs(pp - a); const pb = Math.abs(pp - up); const pc = Math.abs(pp - c); v += pa <= pb && pa <= pc ? a : pb <= pc ? up : c; }
      px[o + x] = v;
    }
  }
  // → 8비트 RGB/회색 + 알파
  const sample = (y, x, i) => {
    if (depth === 8) return px[y * stride + x * ch + i];
    if (depth === 16) return px[y * stride + (x * ch + i) * 2];
    const per = 8 / depth; const byte = px[y * stride + Math.floor(x / per)]; const shift = 8 - depth * ((x % per) + 1);
    const v = (byte >> shift) & ((1 << depth) - 1);
    return ctype === 3 ? v : Math.round((v * 255) / ((1 << depth) - 1));
  };
  const gray = ctype === 0 || ctype === 4;
  const out = new Uint8Array(w * h * (gray ? 1 : 3)); let alpha = null;
  if (ctype === 4 || ctype === 6 || trns) alpha = new Uint8Array(w * h).fill(255);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
    const i = y * w + x;
    if (ctype === 3) { const v = sample(y, x, 0); out[i * 3] = pal?.[v * 3] ?? 0; out[i * 3 + 1] = pal?.[v * 3 + 1] ?? 0; out[i * 3 + 2] = pal?.[v * 3 + 2] ?? 0; if (alpha && trns) alpha[i] = trns[v] ?? 255; }
    else if (gray) { out[i] = sample(y, x, 0); if (ctype === 4) alpha[i] = sample(y, x, 1); else if (trns && sample(y, x, 0) === (depth === 16 ? trns[0] : trns[1])) alpha[i] = 0; }
    else { for (let c = 0; c < 3; c++) out[i * 3 + c] = sample(y, x, c); if (ctype === 6) alpha[i] = sample(y, x, 3); }
  }
  if (alpha && alpha.every((v) => v === 255)) alpha = null;
  return { w, h, gray, data: out, alpha };
}
/** EXIF 방향 → 단위 정사각형 변환 (저장된 그림 → 보이는 그림) */
const ORIENT = { 1: [1, 0, 0, 1, 0, 0], 2: [-1, 0, 0, 1, 1, 0], 3: [-1, 0, 0, -1, 1, 1], 4: [1, 0, 0, -1, 0, 1], 5: [0, -1, -1, 0, 1, 1], 6: [0, -1, 1, 0, 0, 1], 7: [0, 1, 1, 0, 0, 0], 8: [0, 1, -1, 0, 1, 0] };
/**
 * 이미지(JPEG · PNG) 여러 장 → PDF. images = [{ bytes, name? }]
 * paper: 'fit'(그림 크기 그대로) | 'A4' … ; margin pt ; 방향은 그림에 맞춰 자동. 사진의 EXIF 회전 정보를 반영.
 */
export async function imagesToPdf(images, { paper = 'A4', margin = 24, dpi = 96 } = {}) {
  if (!images?.length) throw new Error('이미지가 없습니다.');
  const b = new Builder();
  for (const [k, im] of images.entries()) {
    const bytes = im.bytes;
    let xo; let iw; let ih; let orient = 1;
    const j = jpegInfo(bytes);
    if (j) {
      const cs = j.comps === 1 ? 'DeviceGray' : j.comps === 4 ? 'DeviceCMYK' : 'DeviceRGB';
      const dict = { Type: nm('XObject'), Subtype: nm('Image'), Width: j.w, Height: j.h, ColorSpace: nm(cs), BitsPerComponent: 8, Filter: nm('DCTDecode') };
      if (j.comps === 4 && j.adobe) dict.Decode = [1, 0, 1, 0, 1, 0, 1, 0];
      xo = b.put({ stream: true, dict, data: bytes });
      iw = j.w; ih = j.h; orient = j.orient;
    } else {
      const p = pngDecode(bytes);
      if (!p) throw new Error(`'${im.name ?? `그림 ${k + 1}`}' 은(는) JPEG · PNG 가 아닙니다.`);
      const dict = { Type: nm('XObject'), Subtype: nm('Image'), Width: p.w, Height: p.h, ColorSpace: nm(p.gray ? 'DeviceGray' : 'DeviceRGB'), BitsPerComponent: 8 };
      if (p.alpha) dict.SMask = b.put({ stream: true, dict: { Type: nm('XObject'), Subtype: nm('Image'), Width: p.w, Height: p.h, ColorSpace: nm('DeviceGray'), BitsPerComponent: 8 }, data: p.alpha, compress: true });
      xo = b.put({ stream: true, dict, data: p.data, compress: true });
      iw = p.w; ih = p.h;
    }
    const swap = orient >= 5;
    const dw = (swap ? ih : iw) * (72 / dpi); const dh = (swap ? iw : ih) * (72 / dpi);
    let PW; let PH; let m0 = margin;
    if (paper === 'fit' || !PAPERS[paper]) { PW = dw; PH = dh; m0 = 0; }
    else { const [a, c] = PAPERS[paper]; [PW, PH] = dw > dh ? [c, a] : [a, c]; }
    const s = Math.min((PW - 2 * m0) / dw, (PH - 2 * m0) / dh, paper === 'fit' ? 1 : Infinity);
    const W = dw * s; const H = dh * s;
    const m = mul(ORIENT[orient], [W, 0, 0, H, (PW - W) / 2, (PH - H) / 2]);
    b.addBlank(PW, PH, `q ${mtx(m)} cm /Im0 Do Q`, { XObject: { Im0: xo } });
  }
  return b.finish({ info: { Title: textBytes(images.length === 1 ? (images[0].name ?? '') : '') } });
}

// ───────────── 쪽 번호 · 워터마크 · 머리글/바닥글 ─────────────
// Helvetica 너비 (WinAnsi 32~126, 1000 단위)
const HELV = [278, 278, 355, 556, 556, 889, 667, 191, 333, 333, 389, 584, 278, 333, 278, 278, 556, 556, 556, 556, 556, 556, 556, 556, 556, 556, 278, 278, 584, 584, 584, 556, 1015, 667, 667, 722, 722, 667, 611, 778, 722, 278, 500, 667, 556, 833, 722, 778, 667, 778, 722, 667, 611, 722, 667, 944, 667, 667, 611, 278, 278, 278, 469, 556, 333, 556, 556, 500, 556, 556, 278, 556, 556, 222, 222, 500, 222, 833, 556, 556, 556, 556, 333, 500, 278, 556, 500, 722, 500, 500, 500, 334, 260, 334, 584];
/** 글 → { font: 'helv'|'kor', code, width(1000 단위) } */
function textRun(s, emb = null) {
  if (emb) {
    let hex = '<'; let width = 0;
    for (const ch of s) { const g = emb.font.gidOf(ch.codePointAt(0)); if (!emb.used.has(g)) emb.used.set(g, ch); hex += g.toString(16).padStart(4, '0'); width += emb.font.width(g); }
    return { font: 'emb', code: `${hex}>`, width };
  }
  if (/^[\x20-\x7e]*$/.test(s)) return { font: 'helv', code: `(${s.replace(/[\\()]/g, (m) => `\\${m}`)})`, width: [...s].reduce((a, c) => a + HELV[c.charCodeAt(0) - 32], 0) };
  let hex = '<'; let width = 0;
  for (const ch of s) { const c = ch.codePointAt(0); const v = c > 0xffff ? 0x25a1 : c; hex += v.toString(16).padStart(4, '0'); width += v >= 0x20 && v < 0x7f ? 500 : 1000; }
  return { font: 'kor', code: `${hex}>`, width };
}
const colorOps = (c) => { const m = /^#?([0-9a-f]{6})$/i.exec(c ?? ''); const v = m ? [0, 2, 4].map((i) => parseInt(m[1].slice(i, i + 2), 16) / 255) : [0, 0, 0]; return `${v.map(num).join(' ')} rg`; };
const fmtDate = () => { const d = new Date(); return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`; };
/**
 * 도장 찍기: numbers { format: '{n} / {total}', pos: 'bc'|'br'|'bl'|'tc'|'tr'|'tl', size, start, from (첫 번호를 붙일 페이지, 1부터), margin, color },
 * watermark { text, size, opacity, angle, color }, header/footer { left, center, right, size } ({n} {total} {date} {file} 사용 가능), pages = 지정 글
 */
export async function stampPdf(bytes, { numbers = null, watermark = null, header = null, footer = null, pages = '', password, file = '', font = null, fontName = null } = {}) {
  const src = openPdf(bytes, { password });
  // 글꼴 파일(TTF · TTC)을 주면 쓴 글자만 넣음 — 없으면 ASCII 는 Helvetica, 한글은 미리 정의된 HYGoThic-Medium (보는 프로그램의 글꼴)
  const ttf = font && canEmbedFont(font, fontName) ? openTtf(font, fontName) : null;
  const sel = new Set(parsePages(pages, src.pages.length));
  const b = new Builder();
  const emb = ttf ? b.embed(ttf) : null;
  b.reserve(src, src.pages);
  const total = src.pages.length;
  const newIndex = new Map();
  src.pages.forEach((p, i) => {
    newIndex.set(i, i);
    if (!sel.has(i)) { b.addPage(src, p); return; }
    const res = { Font: {}, ExtGState: {} };
    const ops = [];
    const fill = (t, n) => String(t ?? '').replace(/\{n\}/g, String(n)).replace(/\{total\}/g, String(total + (numbers?.start ?? 1) - 1)).replace(/\{date\}/g, fmtDate()).replace(/\{file\}/g, file);
    const draw = (text, x, y, size, align, color = '#000000', alpha = 1, angle = 0) => {
      if (!text) return;
      const r = textRun(text, /^[\x20-\x7e]*$/.test(text) ? null : emb);
      const fname = r.font === 'helv' ? 'WxH' : r.font === 'emb' ? 'WxE' : 'WxK';
      res.Font[fname] = r.font === 'emb' ? emb.ref : b.font(r.font);
      const w = (r.width / 1000) * size;
      const dx = align === 'c' ? -w / 2 : align === 'r' ? -w : 0;
      let gs = '';
      if (alpha < 1) { res.ExtGState.WxA = b.gstate(alpha); gs = '/WxA gs '; }
      const a = (angle * Math.PI) / 180; const cs = Math.cos(a); const sn = Math.sin(a);
      ops.push(`q ${gs}${colorOps(color)} BT /${fname} ${num(size)} Tf ${num(cs)} ${num(sn)} ${num(-sn)} ${num(cs)} ${num(x + dx * cs - (angle ? -size * 0.35 * sn : 0))} ${num(y + dx * sn - (angle ? size * 0.35 * cs : 0))} Tm ${r.code} Tj ET Q`);
    };
    const W = p.w; const H = p.h;
    if (watermark?.text) draw(watermark.text, W / 2, H / 2, watermark.size ?? 60, 'c', watermark.color ?? '#888888', watermark.opacity ?? 0.25, watermark.angle ?? 45);
    const pageNo = i + (numbers?.start ?? 1) - ((numbers?.from ?? 1) - 1);
    if (numbers && i >= (numbers.from ?? 1) - 1) {
      const pos = numbers.pos ?? 'bc'; const mg = numbers.margin ?? 28; const size = numbers.size ?? 10;
      const x = pos[1] === 'l' ? mg : pos[1] === 'r' ? W - mg : W / 2;
      const y = pos[0] === 't' ? H - mg - size * 0.7 : mg;
      draw(fill(numbers.format ?? '{n}', pageNo), x, y, size, pos[1] === 'l' ? 'l' : pos[1] === 'r' ? 'r' : 'c', numbers.color ?? '#000000');
    }
    for (const [hf, top] of [[header, true], [footer, false]]) {
      if (!hf) continue;
      const size = hf.size ?? 9; const mg = hf.margin ?? 24;
      const y = top ? H - mg - size * 0.7 : mg;
      draw(fill(hf.left, pageNo), mg, y, size, 'l', hf.color);
      draw(fill(hf.center, pageNo), W / 2, y, size, 'c', hf.color);
      draw(fill(hf.right, pageNo), W - mg, y, size, 'r', hf.color);
    }
    if (!ops.length) { b.addPage(src, p); return; }
    // 보이는 좌표로 그린 것을 페이지 좌표로 (돌린 페이지)
    const over = `q ${mtx(visualToUser(p))} cm\n${ops.join('\n')}\nQ`;
    if (!Object.keys(res.ExtGState).length) delete res.ExtGState;
    b.addPage(src, p, { over, res });
  });
  const catalog = catalogFrom(b, src, { labels: true });
  const ol = b.outline(remapOutline(readOutline(src), newIndex));
  if (ol) catalog.Outlines = ol;
  return b.finish({ catalog, info: infoFrom(b, src) });
}

/** 이 글꼴 파일을 PDF 에 넣을 수 있는지 (TrueType glyf, 넣기 금지(fsType 2) 아님) */
export function canEmbedFont(bytes, psName = null) {
  try { const f = openTtf(bytes, psName); return !!f && (f.fsType & 0x0f) !== 2; } catch { return false; }
}
/** 암호 풀어서 저장 (열기 암호 필요) */
export const decryptPdf = (bytes, password = '') => selectPages(bytes, { password });
/** 정리해서 다시 쓰기: 쓰지 않는 개체 버리기 + 압축 안 된 스트림 압축 */
export const compactPdf = (bytes, opts = {}) => selectPages(bytes, { ...opts, compress: true });
/** 문서 정보 바꾸기 { title, author, subject, keywords } (빈 글 = 지움) */
export const setPdfInfo = (bytes, meta, opts = {}) => selectPages(bytes, { ...opts, info: { Title: meta.title, Author: meta.author, Subject: meta.subject, Keywords: meta.keywords } });
