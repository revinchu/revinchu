// 한글 97 이후 바이너리 문서 (.hwp, HWP 5.0) 읽기 — DOM 없음
// 형식: OLE 복합 문서(cfb.js) · FileHeader · DocInfo(글꼴 · 글자/문단 모양 · 테두리 · 번호 · 스타일) · BodyText/SectionN(레코드)
// 레코드 배치는 한컴 공개 문서(HWP 5.0 파일 구조) 와 kordoc (MIT) · hwp-parser 의 구현을 참고함
import { openCfb } from './cfb.js';
import { inflate } from './zip.js';
import { newDoc, newSection, para, uid, normalizeRuns } from './model.js';
import { mediaDataUrl } from './docx.js';
import { hwpEqToLatex } from './hwpx.js';

const HU = 75;
const T = { DOC_PROPS: 0x10, ID_MAPPINGS: 0x11, BIN_DATA: 0x12, FACE_NAME: 0x13, BORDER_FILL: 0x14, CHAR_SHAPE: 0x15, TAB_DEF: 0x16, NUMBERING: 0x17, BULLET: 0x18, PARA_SHAPE: 0x19, STYLE: 0x1a,
  PARA_HEADER: 0x42, PARA_TEXT: 0x43, PARA_CHAR_SHAPE: 0x44, PARA_LINE_SEG: 0x45, CTRL_HEADER: 0x47, LIST_HEADER: 0x48, PAGE_DEF: 0x49, FOOTNOTE_SHAPE: 0x4a, PAGE_BORDER_FILL: 0x4b, SHAPE_COMPONENT: 0x4c, TABLE: 0x4d, SC_LINE: 0x4e, SC_RECT: 0x4f, SC_ELLIPSE: 0x50, SC_PICTURE: 0x55, SC_CONTAINER: 0x56, CTRL_DATA: 0x57, EQEDIT: 0x58 };
const cid = (s) => (s.charCodeAt(0) << 24) | (s.charCodeAt(1) << 16) | (s.charCodeAt(2) << 8) | s.charCodeAt(3);
const C = { tbl: cid('tbl ') >>> 0, gso: cid('gso ') >>> 0, secd: cid('secd') >>> 0, cold: cid('cold') >>> 0, head: cid('head') >>> 0, foot: cid('foot') >>> 0, fn: cid('fn  ') >>> 0, en: cid('en  ') >>> 0, pgnp: cid('pgnp') >>> 0, eqed: cid('eqed') >>> 0, hlk: cid('%hlk') >>> 0, bokm: cid('bokm') >>> 0, atno: cid('atno') >>> 0 };

class R {
  constructor(b) { this.b = b; this.dv = new DataView(b.buffer, b.byteOffset, b.byteLength); this.p = 0; }
  ok(n) { return this.p + n <= this.b.length; }
  u8() { return this.ok(1) ? this.b[this.p++] : 0; }
  i8() { const v = this.ok(1) ? this.dv.getInt8(this.p) : 0; this.p += 1; return v; }
  u16() { const v = this.ok(2) ? this.dv.getUint16(this.p, true) : 0; this.p += 2; return v; }
  i16() { const v = this.ok(2) ? this.dv.getInt16(this.p, true) : 0; this.p += 2; return v; }
  u32() { const v = this.ok(4) ? this.dv.getUint32(this.p, true) : 0; this.p += 4; return v; }
  i32() { const v = this.ok(4) ? this.dv.getInt32(this.p, true) : 0; this.p += 4; return v; }
  str() { const n = this.u16(); let s = ''; for (let i = 0; i < n && this.ok(2); i++) s += String.fromCharCode(this.u16()); return s; }
  skip(n) { this.p += n; return this; }
  at(p) { this.p = p; return this; }
}
const u16 = (b, o) => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);
const u32 = (b, o) => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);
const i32 = (b, o) => (o + 4 <= b.length ? b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24) : 0);
const rgb = (v) => `#${[v & 0xff, (v >> 8) & 0xff, (v >> 16) & 0xff].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`;

function unpack(bytes, compressed) {
  if (!compressed || !bytes?.length) return bytes;
  try { return inflate(bytes[0] === 0x78 && (bytes[1] === 0x9c || bytes[1] === 0xda || bytes[1] === 0x01) ? bytes.subarray(2) : bytes); } catch { return inflate(bytes); }
}
function records(data) {
  const out = [];
  let o = 0;
  while (o + 4 <= data.length && out.length < 2_000_000) {
    const h = u32(data, o); o += 4;
    const tag = h & 0x3ff; const level = (h >>> 10) & 0x3ff; let size = (h >>> 20) & 0xfff;
    if (size === 0xfff) { size = u32(data, o); o += 4; }
    if (o + size > data.length) break;
    out.push({ tag, level, data: data.subarray(o, o + size) });
    o += size;
  }
  return out;
}
/** 레코드 목록 → 나무 (level 기준) */
function tree(recs) {
  const root = { tag: -1, level: -1, children: [] };
  const stack = [root];
  for (const r of recs) {
    const n = { ...r, children: [] };
    while (stack.length > 1 && stack[stack.length - 1].level >= r.level) stack.pop();
    stack[stack.length - 1].children.push(n);
    stack.push(n);
  }
  return root.children;
}

// ───────────── DocInfo ─────────────
const LINE_KIND = ['none', 'single', 'dashed', 'dotted', 'dotDash', 'dotDotDash', 'dashed', 'dotted', 'double', 'thinThickSmallGap', 'thickThinSmallGap', 'triple', 'wave', 'wave', 'single', 'single', 'single'];
const LINE_W = [0.1, 0.12, 0.15, 0.2, 0.25, 0.3, 0.4, 0.5, 0.6, 0.7, 1.0, 1.5, 2.0, 3.0, 4.0, 5.0]; // mm
const NUMFMT = { 0: 'decimal', 1: 'decimalCircle', 2: 'upperRoman', 3: 'lowerRoman', 4: 'upperLetter', 5: 'lowerLetter', 6: 'upperLetter', 7: 'lowerLetter', 8: 'ganada', 9: 'ganada', 10: 'chosung', 11: 'chosung', 12: 'koreanDigital', 13: 'ideographDigital', 14: 'ideographDigital' };
function readDocInfo(data) {
  const D = { faces: [[], [], [], [], [], [], []], bf: [], cs: [], ps: [], styles: [], num: [], bullets: [], bin: [], tabs: [], counts: null };
  let faceLang = 0; let faceLeft = null;
  for (const r of records(data)) {
    const b = r.data;
    switch (r.tag) {
      case T.ID_MAPPINGS: { const rd = new R(b); D.counts = Array.from({ length: Math.floor(b.length / 4) }, () => rd.i32()); faceLeft = D.counts.slice(1, 8); break; }
      case T.BIN_DATA: {
        const rd = new R(b);
        const attr = rd.u16();
        const type = attr & 0xf;
        if (type === 0) { rd.str(); rd.str(); D.bin.push({ link: true }); } else { const id = rd.u16(); const ext = rd.str(); D.bin.push({ id, ext: ext.toLowerCase(), compress: (attr >> 4) & 3 }); }
        break;
      }
      case T.FACE_NAME: {
        const rd = new R(b); rd.u8(); const name = rd.str();
        while (faceLeft && faceLang < 7 && faceLeft[faceLang] <= 0) faceLang++;
        D.faces[Math.min(6, faceLang)].push(name);
        if (faceLeft && faceLang < 7) faceLeft[faceLang]--;
        break;
      }
      case T.BORDER_FILL: {
        const rd = new R(b);
        rd.u16();
        const side = () => { const k = rd.u8(); const w = rd.u8(); const c = rd.u32(); return k === 0 ? { style: 'none' } : { style: LINE_KIND[k] ?? 'single', w: Math.max(0.25, ((LINE_W[w] ?? 0.12) * 72) / 25.4), color: rgb(c) }; };
        const l = side(); const rr = side(); const t = side(); const bt = side();
        rd.skip(6); // 대각선
        let fill = null;
        if (rd.ok(4)) { const ft = rd.u32(); if (ft & 1) { const bg = rd.u32(); rd.u32(); rd.i32(); if ((bg & 0xffffffff) !== 0xffffffff) fill = rgb(bg); } }
        D.bf.push({ l, r: rr, t, b: bt, fill });
        break;
      }
      case T.CHAR_SHAPE: {
        const rd = new R(b);
        const face = Array.from({ length: 7 }, () => rd.u16());
        const ratio = Array.from({ length: 7 }, () => rd.u8());
        const spacing = Array.from({ length: 7 }, () => rd.i8());
        rd.skip(7); // relSize
        const offset = Array.from({ length: 7 }, () => rd.i8());
        const base = rd.i32();
        const attr = rd.u32();
        rd.skip(2);
        const color = rd.u32(); const uColor = rd.u32(); const shade = rd.u32();
        const o = { sz: base / 100 };
        o.fontEa = D.faces[0][face[0]] ?? undefined;
        o.font = D.faces[1][face[1]] ?? o.fontEa;
        if (attr & 1) o.i = true;
        if (attr & 2) o.b = true;
        const ut = (attr >> 2) & 3;
        if (ut === 1) { const shape = (attr >> 4) & 0xf; o.u = shape === 8 ? 'double' : shape === 2 ? 'dotted' : shape === 1 ? 'dash' : shape === 12 ? 'wave' : true; if ((uColor & 0xffffff) !== (color & 0xffffff)) o.uColor = rgb(uColor); }
        if (((attr >> 8) & 7) !== 0) o.outline = true;
        if (((attr >> 11) & 3) !== 0) o.shadow = true;
        if (attr & (1 << 13)) o.emboss = true;
        if (attr & (1 << 14)) o.imprint = true;
        if (attr & (1 << 15)) o.va = 'super';
        else if (attr & (1 << 16)) o.va = 'sub';
        if (((attr >> 18) & 7) !== 0 && ((attr >> 26) & 0xf) <= 12) o.s = true;
        const sym = (attr >> 21) & 0xf;
        if (sym) o.emph = sym === 2 ? 'circle' : 'dot';
        const c = color & 0xffffff;
        if (c) o.color = rgb(c);
        if ((shade & 0xffffffff) !== 0xffffffff && (shade & 0xffffff) !== 0xffffff) o.shd = rgb(shade);
        if (ratio[0] && ratio[0] !== 100) o.scale = ratio[0];
        if (spacing[0]) o.spc = Math.round(((spacing[0] / 100) * o.sz) * 100) / 100;
        if (offset[0]) o.pos = Math.round(((-offset[0] / 100) * o.sz) * 100) / 100;
        D.cs.push(o);
        break;
      }
      case T.TAB_DEF: {
        const rd = new R(b); rd.u32(); const n = rd.i16();
        const list = [];
        for (let k = 0; k < n && rd.ok(8); k++) { const pos = rd.u32(); const kind = rd.u8(); const fill = rd.u8(); rd.u16(); list.push({ pos: pos / HU, align: ['left', 'right', 'center', 'decimal'][kind] ?? 'left', ...(fill ? { leader: fill === 3 ? 'dot' : 'hyphen' } : {}) }); }
        D.tabs.push(list);
        break;
      }
      case T.NUMBERING: {
        const rd = new R(b);
        const levels = [];
        for (let k = 0; k < 7; k++) { if (!rd.ok(12)) break; const attr = rd.u32(); rd.i16(); rd.i16(); rd.u32(); const fmt = rd.str(); levels.push({ fmt: NUMFMT[(attr >> 5) & 0xf] ?? 'decimal', text: fmt.replace(/\^(\d)/g, '%$1') }); }
        const start = rd.ok(2) ? rd.u16() : 1;
        const starts = []; for (let k = 0; k < 7; k++) starts.push(rd.ok(4) ? rd.u32() || start || 1 : start || 1);
        D.num.push({ levels: levels.map((l, k) => ({ ...l, start: starts[k] })) });
        break;
      }
      case T.BULLET: { const ch = u16(b, 12); D.bullets.push(ch && ch !== 0xffff ? String.fromCharCode(ch) : '●'); break; }
      case T.PARA_SHAPE: {
        const rd = new R(b);
        const a1 = rd.u32();
        const left = rd.i32(); const right = rd.i32(); const indent = rd.i32(); const prev = rd.i32(); const next = rd.i32(); const lsOld = rd.i32();
        const tabId = rd.u16(); const numId = rd.u16(); const bfId = rd.u16();
        rd.skip(8);
        const a2 = rd.ok(4) ? rd.u32() : 0;
        const a3 = rd.ok(4) ? rd.u32() : 0;
        const lsNew = rd.ok(4) ? rd.u32() : null;
        // 문단 여백 값은 HWPUNIT × 2 로 저장됨 (한컴 화면 값의 2배)
        const k = 1 / HU / 2;
        const o = { align: ['justify', 'left', 'right', 'center', 'distribute', 'distribute'][(a1 >> 2) & 7] ?? 'justify' };
        const ind = {};
        if (left) ind.l = left * k;
        if (right) ind.r = right * k;
        if (indent > 0) ind.first = indent * k;
        else if (indent < 0) { ind.l = (ind.l ?? 0) - indent * k; ind.hanging = -indent * k; }
        if (Object.keys(ind).length) o.ind = ind;
        const lsType = lsNew != null ? a3 & 0x1f : a1 & 3;
        const lsVal = lsNew != null ? lsNew : lsOld;
        o.spc = { before: prev * k, after: next * k, ...(lsType === 0 ? { line: lsVal / 100, rule: 'pct' } : lsType === 1 ? { line: lsVal / HU / 2, rule: 'exact' } : lsType === 3 ? { line: lsVal / HU / 2, rule: 'atLeast' } : { line: 1.6, rule: 'pct' }) };
        if (a1 & (1 << 16)) o.widow = true; else o.widow = false;
        if (a1 & (1 << 17)) o.keepNext = true;
        if (a1 & (1 << 18)) o.keepLines = true;
        if (a1 & (1 << 19)) o.pageBreakBefore = true;
        const headType = (a1 >>> 23) & 3;
        const level = (a1 >>> 25) & 7;
        if (headType) o._head = { type: headType, id: numId, level };
        if (bfId) o._bf = bfId;
        if (tabId != null) o._tab = tabId;
        void a2;
        D.ps.push(o);
        break;
      }
      case T.STYLE: {
        const rd = new R(b);
        const name = rd.str(); const eng = rd.str();
        const type = rd.u8(); rd.u8(); rd.i16();
        const pid = rd.u16(); const cid0 = rd.u16();
        D.styles.push({ name, eng, type, ps: pid, cs: cid0 });
        break;
      }
      default: break;
    }
  }
  return D;
}

// ───────────── 본문 ─────────────
function readHwpCore(cfb, warnings) {
  const fh = cfb.stream('FileHeader');
  if (!fh || fh.length < 40) throw new Error('한글 문서(.hwp)의 FileHeader 가 없습니다.');
  const sig = new TextDecoder().decode(fh.subarray(0, 17));
  if (!sig.startsWith('HWP Document File')) throw new Error('한글 5.0 문서가 아닙니다. (한글 97 이전 형식은 한글에서 .hwpx 로 저장한 뒤 열어 주세요)');
  const flags = u32(fh, 36);
  if (flags & 2) throw new Error('암호로 보호된 한글 문서입니다. 한글에서 암호를 풀고 저장한 뒤 열어 주세요.');
  if (flags & (1 << 2)) throw new Error('배포용 한글 문서입니다 (복사 · 편집 제한). 한글에서 일반 문서로 저장한 뒤 열어 주세요.');
  if (flags & (1 << 4)) throw new Error('DRM 으로 보호된 한글 문서입니다.');
  const compressed = !!(flags & 1);
  const D = readDocInfo(unpack(cfb.stream('DocInfo'), compressed));
  const doc = newDoc();
  doc.media = {}; doc.lists = {};
  const ctx = { D, doc, compressed, cfb, warnings, binIds: new Map(), hf: null, outlineNum: 0 };
  // 스타일
  const styleMap = {};
  D.styles.forEach((s, i) => {
    if (s.type !== 0) return;
    const outline = /^개요\s*(\d)$/.exec(s.name) ?? /^Outline\s*(\d)$/i.exec(s.eng);
    const key = i === 0 ? 'Normal' : outline ? `Heading${outline[1]}` : `hwp${i}`;
    styleMap[i] = key;
    const pPr = paraPr(s.ps, ctx); delete pPr.num;
    const rPr = { ...(D.cs[s.cs] ?? {}) };
    if (key === 'Normal') { doc.defaults = { rPr: { ...doc.defaults.rPr, ...rPr }, pPr: { ...pPr } }; doc.styles.Normal = { ...doc.styles.Normal, name: s.name || '바탕글', pPr: {}, rPr: {} }; return; }
    doc.styles[key] = { ...(doc.styles[key] ?? {}), id: key, type: 'para', name: s.name, basedOn: 'Normal', quick: !!outline || /^본문|^제목/.test(s.name), order: outline ? 9 + Number(outline[1]) : 100 + i, pPr: diff(pPr, doc.defaults.pPr), rPr: diff(rPr, doc.defaults.rPr) };
  });
  ctx.styleMap = styleMap;
  // 구역
  const names = cfb.names.filter((n) => /^BodyText\/Section\d+$/.test(n)).sort((a, b) => Number(/\d+$/.exec(a)[0]) - Number(/\d+$/.exec(b)[0]));
  if (!names.length) throw new Error('본문(BodyText)이 없습니다.');
  const body = [];
  let lastSec = null;
  for (const n of names) {
    ctx.hf = { headers: {}, footers: {} };
    ctx.sec = newSection();
    ctx.sec.headers = {}; ctx.sec.footers = {};
    const nodes = tree(records(unpack(cfb.stream(n), compressed)));
    const blocks = [];
    for (const nd of nodes) if (nd.tag === T.PARA_HEADER) blocks.push(...readPara(nd, ctx));
    const sec = ctx.sec;
    sec.headers = ctx.hf.headers; sec.footers = ctx.hf.footers;
    if (!blocks.length) blocks.push(para(''));
    if (lastSec) { const prev = body[body.length - 1]; if (prev?.type === 'p') prev.secBreak = { kind: 'nextPage', sec: lastSec }; else body.push({ ...para(''), secBreak: { kind: 'nextPage', sec: lastSec } }); }
    body.push(...blocks);
    lastSec = sec;
  }
  doc.body = body;
  doc.section = lastSec;
  // 요약 정보 (제목 · 지은이)
  return doc;
}
function diff(full, base) { const o = {}; for (const [k, v] of Object.entries(full)) if (JSON.stringify(v) !== JSON.stringify(base?.[k])) o[k] = v; return o; }
function paraPr(id, ctx) {
  const src = ctx.D.ps[id] ?? {};
  const { _head, _bf, _tab, ...p } = src;
  if (_tab != null && ctx.D.tabs[_tab]?.length) p.tabs = ctx.D.tabs[_tab];
  if (_bf) { const bf = ctx.D.bf[_bf - 1]; if (bf) { const b = {}; for (const [k, s] of [['top', 't'], ['bottom', 'b'], ['left', 'l'], ['right', 'r']]) if (bf[s].style !== 'none') b[k] = { ...bf[s], space: 2 }; if (Object.keys(b).length) p.border = b; if (bf.fill && bf.fill !== '#FFFFFF') p.shd = bf.fill; } }
  if (_head) {
    const lvl = Math.min(8, _head.level);
    if (_head.type === 1) {
      p.outline = lvl;
      const nid = ctx.outlineNum || _head.id;
      const def = ctx.D.num[nid - 1];
      if (def && def.levels.some((l) => l.text)) { const key = `hwpOutline${nid}`; if (!ctx.doc.lists[key]) ctx.doc.lists[key] = listOf(def); p.num = { id: key, lvl }; }
    } else if (_head.type === 2) { const def = ctx.D.num[_head.id - 1]; if (def) { const key = `hwpNum${_head.id}`; if (!ctx.doc.lists[key]) ctx.doc.lists[key] = listOf(def); p.num = { id: key, lvl }; } }
    else if (_head.type === 3) { const key = `hwpBullet${_head.id}`; if (!ctx.doc.lists[key]) ctx.doc.lists[key] = { levels: Array.from({ length: 9 }, () => ({ fmt: 'bullet', text: ctx.D.bullets[_head.id - 1] ?? '●', start: 1, ind: { l: 0, hanging: 0 }, suff: 'space' })) }; p.num = { id: key, lvl }; }
  }
  return p;
}
const listOf = (def) => ({ levels: Array.from({ length: 9 }, (_, i) => { const l = def.levels[i] ?? { fmt: 'decimal', text: `%${i + 1}.`, start: 1 }; return { fmt: l.fmt, text: l.text, start: l.start ?? 1, ind: { l: 0, hanging: 0 }, suff: 'space' }; }) });

function readPara(node, ctx) {
  const h = node.data;
  const psId = u16(h, 8);
  const stId = h[10] ?? 0;
  const brk = h[11] ?? 0; // 1 구역 · 2 다단 · 4 쪽 · 8 단
  const style = ctx.styleMap[stId] ?? 'Normal';
  const base = ctx.doc.styles[style];
  const pPr = diff(paraPr(psId, ctx), { ...(ctx.doc.defaults.pPr ?? {}), ...(base?.pPr ?? {}) });
  const textRec = node.children.find((c) => c.tag === T.PARA_TEXT);
  const csRec = node.children.find((c) => c.tag === T.PARA_CHAR_SHAPE);
  const ctrls = node.children.filter((c) => c.tag === T.CTRL_HEADER);
  // 글자 모양 위치표
  const csPos = [];
  if (csRec) for (let o = 0; o + 8 <= csRec.data.length; o += 8) csPos.push([u32(csRec.data, o), u32(csRec.data, o + 4)]);
  if (!csPos.length) csPos.push([0, 0]);
  const rpAt = (pos) => { let id = csPos[0][1]; for (const [p, c] of csPos) { if (p <= pos) id = c; else break; } return diff(ctx.D.cs[id] ?? {}, { ...ctx.doc.defaults.rPr, ...(base?.rPr ?? {}) }); };
  const out = [];
  let cur = { ...para(''), style, pPr: { ...pPr }, runs: [] };
  if (brk & 4 || brk & 1) cur.pPr.pageBreakBefore = brk & 4 ? true : cur.pPr.pageBreakBefore;
  if (brk & 8) cur.runs.push({ obj: 'br', kind: 'column' });
  const flush = (force) => { if (force || cur.runs.length) { cur.runs = normalizeRuns(cur.runs); out.push(cur); } cur = { ...para(''), style, pPr: { ...pPr }, runs: [] }; };
  let ci = 0;
  let link = null;
  const b = textRec?.data;
  let buf = ''; let bufPos = 0;
  const flushText = () => { if (buf) { cur.runs.push({ ...rpAt(bufPos), ...(link ? { link } : {}), text: buf }); buf = ''; } };
  if (b) {
    for (let o = 0; o + 1 < b.length;) {
      const pos = o / 2;
      const ch = u16(b, o);
      if (ch >= 32) {
        const rp = rpAt(pos);
        if (buf && JSON.stringify(rpAt(bufPos)) !== JSON.stringify(rp)) flushText();
        if (!buf) bufPos = pos;
        if (ch >= 0xd800 && ch <= 0xdbff && o + 3 < b.length) { buf += String.fromCharCode(ch, u16(b, o + 2)); o += 4; continue; }
        buf += String.fromCharCode(ch);
        o += 2;
        continue;
      }
      // 제어 문자
      flushText();
      const extended = (ch >= 1 && ch <= 3) || ch === 11 || ch === 12 || (ch >= 14 && ch <= 18) || (ch >= 21 && ch <= 23);
      const inline = (ch >= 4 && ch <= 9) || ch === 19 || ch === 20;
      if (ch === 13) { o += 2; continue; }
      if (ch === 10) { cur.runs.push({ obj: 'br' }); o += 2; continue; }
      if (ch === 24) { o += 2; continue; }
      if (ch === 30) { buf = ' '; bufPos = pos; flushText(); o += 2; continue; }
      if (ch === 31) { buf = ' '; bufPos = pos; flushText(); o += 2; continue; }
      if (ch === 9) { cur.runs.push({ ...rpAt(pos), obj: 'tab' }); o += 16; continue; }
      if (ch === 4) { link = null; o += 16; continue; } // 필드 끝
      if (extended) {
        const ctrl = ctrls[ci++];
        if (ctrl) {
          const r = readCtrl(ctrl, ctx, rpAt(pos));
          if (r?.blocks) { flush(false); out.push(...r.blocks); }
          else if (r?.runs) cur.runs.push(...r.runs);
          else if (r?.link) link = r.link;
        }
        o += 16;
        continue;
      }
      o += inline ? 16 : 2;
    }
    flushText();
  }
  // PARA_TEXT 가 없어도 컨트롤만 있는 문단 (드묾)
  if (!b) for (const c of ctrls) { const r = readCtrl(c, ctx, rpAt(0)); if (r?.blocks) { flush(false); out.push(...r.blocks); } else if (r?.runs) cur.runs.push(...r.runs); }
  if (cur.runs.length || !out.length || out[out.length - 1].type === 'tbl') flush(true);
  if (!cur.runs.length && out.length && out[out.length - 1].type !== 'p') { /* 표 뒤 빈 문단은 위에서 추가됨 */ }
  return out;
}

/** 컨트롤 하위의 문단 목록 (LIST_HEADER 다음 PARA_HEADER 들) */
function listParas(nodes, ctx) {
  const out = [];
  for (const n of nodes) if (n.tag === T.PARA_HEADER) out.push(...readPara(n, ctx));
  return out;
}
function groupLists(children) {
  // [LIST_HEADER, P, P, LIST_HEADER, P …] → [{ lh, paras }]
  const groups = [];
  for (const c of children) {
    if (c.tag === T.LIST_HEADER) groups.push({ lh: c, paras: [] });
    else if (c.tag === T.PARA_HEADER && groups.length) groups[groups.length - 1].paras.push(c);
  }
  return groups;
}
function firstList(node) {
  // 하위 나무에서 첫 LIST_HEADER 와 그 뒤 문단들
  const kids = node.children;
  const i = kids.findIndex((c) => c.tag === T.LIST_HEADER);
  if (i >= 0) return kids.slice(i + 1).filter((c) => c.tag === T.PARA_HEADER);
  for (const c of kids) if (c.tag === T.SHAPE_COMPONENT || c.tag === T.SC_CONTAINER) { const r = firstList(c); if (r.length) return r; }
  return [];
}
function objPos(d) {
  // 개체 공통 속성: ctrlId(4) attr(4) vOff(4) hOff(4) w(4) h(4) z(4) margins(8) instId(4) …
  const attr = u32(d, 4);
  const treat = !!(attr & 1);
  const vRel = (attr >> 3) & 3; const hRel = (attr >> 8) & 3;
  const wrap = (attr >> 21) & 7;
  return { treat, vRel, hRel, wrap, y: i32(d, 8) / HU, x: i32(d, 12) / HU, w: u32(d, 16) / HU, h: u32(d, 20) / HU, hAlign: (attr >> 10) & 7, vAlign: (attr >> 5) & 7 };
}
function floatOf(p) {
  if (p.treat) return null;
  const wrap = ['square', 'tight', 'tight', 'topBottom', 'behind', 'front'][p.wrap] ?? 'square';
  const relH = ['page', 'margin', 'column', 'column'][p.hRel] ?? 'column';
  const relV = ['page', 'margin', 'paragraph'][p.vRel] ?? 'paragraph';
  const f = { wrap, relH, relV, x: p.x, y: p.y };
  if (p.hAlign === 1) f.alignH = 'center'; else if (p.hAlign === 2) f.alignH = 'right';
  if (wrap === 'square' || wrap === 'tight') f.side = p.hAlign === 2 ? 'right' : 'left';
  return f;
}
function binMedia(id, ctx) {
  if (!id) return null;
  if (ctx.binIds.has(id)) return ctx.binIds.get(id);
  const item = ctx.D.bin[id - 1];
  if (!item || item.link) return null;
  const sid = item.id || id;
  const name = `BinData/BIN${sid.toString(16).toUpperCase().padStart(4, '0')}.${item.ext}`;
  let bytes = ctx.cfb.stream(name) ?? ctx.cfb.stream(name.toLowerCase()) ?? ctx.cfb.names.map((n) => (n.toUpperCase() === name.toUpperCase() ? ctx.cfb.stream(n) : null)).find(Boolean);
  if (!bytes) return null;
  const comp = item.compress === 0 ? ctx.compressed : item.compress === 1;
  if (comp) { try { bytes = unpack(bytes, true); } catch { /* 그대로 */ } }
  const mid = uid('m');
  ctx.doc.media[mid] = mediaDataUrl(bytes, `x.${item.ext}`);
  ctx.binIds.set(id, mid);
  return mid;
}
function readCtrl(node, ctx, rp) {
  const d = node.data;
  const id = u32(d, 0);
  switch (id) {
    case C.secd: {
      const pd = node.children.find((c) => c.tag === T.PAGE_DEF);
      if (pd) {
        const b = pd.data;
        const W = u32(b, 0) / HU; const H = u32(b, 4) / HU;
        const L = u32(b, 8) / HU; const Rm = u32(b, 12) / HU; const Tp = u32(b, 16) / HU; const Bt = u32(b, 20) / HU; const Hd = u32(b, 24) / HU; const Ft = u32(b, 28) / HU; const G = u32(b, 32) / HU;
        const land = !!(u32(b, 36) & 1);
        ctx.sec.page = { w: land ? Math.max(W, H) : W, h: land ? Math.min(W, H) : H, orient: land ? 'landscape' : 'portrait' };
        ctx.sec.margin = { t: Tp + Hd, b: Bt + Ft, l: L, r: Rm, header: Tp, footer: Bt, gutter: G };
      }
      const pb = node.children.find((c) => c.tag === T.PAGE_BORDER_FILL);
      if (pb) { const bfId = u16(pb.data, 12); const bf = ctx.D.bf[bfId - 1]; if (bf && bf.t.style !== 'none') ctx.sec.pageBorder = { ...bf.t, space: 12 }; }
      const outlineId = u16(d, 14);
      if (outlineId) ctx.outlineNum = outlineId;
      const pgStart = u16(d, 16);
      if (pgStart) ctx.sec.pgNum = { fmt: 'decimal', start: pgStart };
      return null;
    }
    case C.cold: {
      const attr = u16(d, 4);
      const n = (attr >> 2) & 0xff;
      if (n > 1) ctx.sec.cols = { n, space: u16(d, 6) / HU || 48, sep: false };
      return null;
    }
    case C.head: case C.foot: {
      const attr = u32(d, 4);
      const which = (attr & 3) === 1 ? 'even' : 'default';
      const paras = listParas(firstList(node), ctx);
      for (const p of paras) if (p.type === 'p' && p.style === 'Normal') p.style = id === C.head ? 'Header' : 'Footer';
      const prop = id === C.head ? 'headers' : 'footers';
      ctx.hf[prop][which] = paras.length ? paras : [para('')];
      if ((attr & 3) === 2 && !ctx.hf[prop].default) ctx.hf[prop].default = ctx.hf[prop][which];
      if (which === 'even') ctx.doc.settings.evenOdd = true;
      return null;
    }
    case C.fn: case C.en: {
      const kind = id === C.fn ? 'foot' : 'end';
      const nid = uid('n');
      const paras = listParas(firstList(node), ctx);
      ctx.doc.notes[kind][nid] = paras.length ? paras : [para('')];
      return { runs: [{ obj: 'note', id: nid, kind }] };
    }
    case C.pgnp: {
      const attr = u32(d, 4);
      const pos = (attr >> 8) & 0xf;
      const side = String.fromCharCode(u16(d, 14) || 0).replace('\0', '');
      if (!pos) return null;
      const prop = pos <= 3 ? 'headers' : 'footers';
      const align = [0, 1, 4, 7].includes(pos) ? 'left' : [3, 6, 9].includes(pos) ? 'right' : 'center';
      if (!ctx.hf[prop].default) ctx.hf[prop].default = [{ ...para(''), style: prop === 'headers' ? 'Header' : 'Footer', pPr: { align }, runs: [...(side ? [{ text: `${side} ` }] : []), { obj: 'field', code: 'PAGE', result: '1' }, ...(side ? [{ text: ` ${side}` }] : [])] }];
      return null;
    }
    case C.tbl: { ctx.lastCaption = []; const t = readTable(node, ctx); return { blocks: [t, ...(ctx.lastCaption ?? [])].filter(Boolean) }; }
    case C.gso: {
      const p = objPos(d);
      const pic = findTag(node, T.SC_PICTURE);
      if (pic) {
        const bid = u16(pic.data, 71);
        const media = binMedia(bid, ctx);
        const run = { obj: 'img', id: uid('i'), w: p.w || 100, h: p.h || 100, alt: '', ...(media ? { media } : { src: '' }) };
        const f = floatOf(p);
        if (f) run.float = f;
        return { runs: [run] };
      }
      // 글상자 · 도형
      const paras = listParas(firstList(node), ctx);
      const sc = findTag(node, T.SHAPE_COMPONENT);
      const run = { obj: 'shape', id: uid('s'), w: p.w || 100, h: p.h || 60, prst: findTag(node, T.SC_ELLIPSE) ? 'ellipse' : 'rect', blocks: paras, float: floatOf(p) ?? { wrap: 'topBottom', relH: 'column', relV: 'paragraph', x: 0, y: 0 } };
      if (sc) {
        // 선 · 채우기 (SHAPE_COMPONENT 뒤쪽): 선 색 u32, 굵기 i32 … 단순화: 테두리 있음으로
        run.line = { w: 0.5, color: '#000000' };
      }
      if (!paras.length && !sc) return null;
      return { runs: [run] };
    }
    case C.eqed: {
      const eq = findTag(node, T.EQEDIT);
      if (!eq) return null;
      const rd = new R(eq.data); rd.u32(); const script = rd.str();
      return { runs: [{ obj: 'eq', latex: hwpEqToLatex(script) }] };
    }
    case C.hlk: {
      const rd = new R(d); rd.u32(); rd.u32(); rd.u8(); const cmd = rd.str();
      const url = cmd.split(';')[0].replace(/\\(.)/g, '$1');
      return { link: /^[a-z]+:/i.test(url) ? { url } : { anchor: url.replace(/^#/, '') } };
    }
    case C.bokm: {
      const cd = node.children.find((c) => c.tag === T.CTRL_DATA);
      const name = cd ? (() => { const rd = new R(cd.data); rd.skip(12); return rd.str(); })() : '';
      return name ? { runs: [{ obj: 'mark', kind: 'bm', name, edge: 'start' }, { obj: 'mark', kind: 'bm', name, edge: 'end' }] } : null;
    }
    default: return null;
  }
}
function findTag(node, tag) { for (const c of node.children) { if (c.tag === tag) return c; const r = findTag(c, tag); if (r) return r; } return null; }
function readTable(node, ctx) {
  const tr = node.children.find((c) => c.tag === T.TABLE);
  if (!tr) return null;
  const b = tr.data;
  const rows = u16(b, 4); const cols = u16(b, 6);
  const inMar = { l: u16(b, 10) / HU, r: u16(b, 12) / HU, t: u16(b, 14) / HU, b: u16(b, 16) / HU };
  const repeat = !!(u32(b, 0) & 2);
  const cells = [];
  // TABLE 레코드 앞의 LIST_HEADER 는 캡션 (셀 아님)
  const ti = node.children.indexOf(tr);
  const capGroups = groupLists(node.children.slice(0, ti));
  ctx.lastCaption = capGroups.flatMap((g) => listParas(g.paras, ctx));
  for (const g of groupLists(node.children.slice(ti + 1))) {
    const lh = g.lh.data;
    const cell = { c: u16(lh, 8), r: u16(lh, 10), cs: u16(lh, 12) || 1, rs: u16(lh, 14) || 1, w: u32(lh, 16) / HU, h: u32(lh, 20) / HU, mar: { l: u16(lh, 24) / HU, r: u16(lh, 26) / HU, t: u16(lh, 28) / HU, b: u16(lh, 30) / HU }, bf: ctx.D.bf[u16(lh, 32) - 1], va: (u32(lh, 2) >> 21) & 3, header: !!(u16(lh, 6) & 4) };
    cell.blocks = listParas(g.paras, ctx);
    cells.push(cell);
  }
  if (!cells.length) return null;
  const Rn = Math.max(rows, ...cells.map((c) => c.r + c.rs));
  const Cn = Math.max(cols, ...cells.map((c) => c.c + c.cs));
  const grid = Array(Cn).fill(0);
  for (const c of cells) if (c.cs === 1 && !grid[c.c]) grid[c.c] = c.w;
  for (const c of cells) if (c.cs > 1) { const known = grid.slice(c.c, c.c + c.cs).reduce((a, x) => a + x, 0); const miss = grid.slice(c.c, c.c + c.cs).filter((x) => !x).length; if (miss) for (let k = c.c; k < c.c + c.cs; k++) if (!grid[k]) grid[k] = Math.max(8, (c.w - known) / miss); }
  for (let k = 0; k < Cn; k++) if (!grid[k]) grid[k] = 48;
  const at = new Map(cells.map((c) => [`${c.r},${c.c}`, c]));
  const cover = new Map();
  for (const c of cells) for (let r = c.r + 1; r < c.r + c.rs; r++) cover.set(`${r},${c.c}`, c);
  const t = { type: 'tbl', id: uid('t'), style: 'TableNormal', tblPr: { look: { firstRow: false, firstCol: false, bandRows: false }, cellMar: inMar }, grid, rows: [] };
  const p = objPos(node.data);
  if (!p.treat && p.hAlign === 1) t.tblPr.align = 'center';
  const bordersOf = (bf) => (bf ? { top: bf.t, bottom: bf.b, left: bf.l, right: bf.r } : undefined);
  for (let r = 0; r < Rn; r++) {
    const row = { cells: [] };
    let hgt = 0;
    for (let c = 0; c < Cn;) {
      const cell = at.get(`${r},${c}`);
      const cov = cover.get(`${r},${c}`);
      if (cell) {
        const o = { blocks: cell.blocks.length ? cell.blocks : [para('')] };
        if (cell.cs > 1) o.span = cell.cs;
        if (cell.rs > 1) o.vMerge = 'restart';
        if (cell.bf) { o.borders = bordersOf(cell.bf); if (cell.bf.fill && cell.bf.fill !== '#FFFFFF') o.shd = cell.bf.fill; }
        if (cell.va === 1) o.vAlign = 'center'; else if (cell.va === 2) o.vAlign = 'bottom';
        if (cell.mar && (Math.abs(cell.mar.l - inMar.l) > 0.5 || Math.abs(cell.mar.t - inMar.t) > 0.5)) o.mar = cell.mar;
        if (cell.header && repeat) row.header = true;
        if (cell.rs === 1) hgt = Math.max(hgt, cell.h);
        row.cells.push(o);
        c += cell.cs;
      } else if (cov) {
        const o = { blocks: [para('')], vMerge: 'cont', ...(cov.cs > 1 ? { span: cov.cs } : {}) };
        if (cov.bf) { o.borders = bordersOf(cov.bf); if (cov.bf.fill && cov.bf.fill !== '#FFFFFF') o.shd = cov.bf.fill; }
        row.cells.push(o);
        c += cov.cs;
      } else { row.cells.push({ blocks: [para('')] }); c++; }
    }
    if (hgt) { row.h = hgt; row.hRule = 'atLeast'; }
    t.rows.push(row);
  }
  return t;
}

/** 요약 정보 (\005HwpSummaryInformation): 제목 · 지은이 — OLE 속성 집합 */
function readSummary(cfb, doc) {
  const s = cfb.stream('\u0005HwpSummaryInformation');
  if (!s || s.length < 48) return;
  try {
    const off = u32(s, 44);
    const n = u32(s, off + 4);
    for (let k = 0; k < n; k++) {
      const pid = u32(s, off + 8 + k * 8);
      const po = off + u32(s, off + 12 + k * 8);
      if (u32(s, po) !== 0x1f) continue; // VT_LPWSTR
      const len = u32(s, po + 4);
      let str = '';
      for (let i = 0; i < len - 1; i++) str += String.fromCharCode(u16(s, po + 8 + i * 2));
      if (pid === 2) doc.props.title = str;
      else if (pid === 3) doc.props.subject = str;
      else if (pid === 4) doc.props.author = str;
      else if (pid === 5) doc.props.keywords = str;
      else if (pid === 6) doc.props.comments = str;
    }
  } catch { /* 무시 */ }
}

export function readHwp(bytes) {
  const warnings = [];
  if (new TextDecoder('latin1').decode(bytes.subarray(0, 20)).startsWith('HWP Document File V')) throw new Error('한글 97 이전(HWP 3.0) 형식입니다. 한글에서 [다른 이름으로 저장 › .hwpx 또는 .hwp]로 저장한 뒤 열어 주세요.');
  const cfb = openCfb(bytes);
  const doc = readHwpCore(cfb, warnings);
  readSummary(cfb, doc);
  return { doc, warnings };
}
