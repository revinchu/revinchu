// 한글 3.0 / 한글 97 문서 (.hwp, "HWP Document File V3.00") 읽기 — DOM 없음
// 형식: 30바이트 서명 + 문서 정보 128 + 요약 1008 + 정보 블록 + 본문(raw deflate) — 본문은 글꼴 · 스타일 · 문단 목록 · 추가 정보(그림) 블록
// 바이트 배치와 글자 코드(상용 조합형 + 한컴 사적 영역)는 rhwp (MIT, src/parser/hwp3) 와 kordoc (MIT, src/hwp3) 를 참고함
import { inflate } from './zip.js';
import { newDoc, newSection, para, uid, normalizeRuns } from './model.js';
import { mediaDataUrl } from './docx.js';
import { hwpEqToLatex } from './hwpx.js';
import { gridTable } from './hwp.js';
import { desEcbDecrypt } from './crypto.js';

const PX = 4 / 75; // hunit(1/1800 in) → px(1/96 in)
const px = (v) => Math.round(v * PX * 100) / 100;

// ───────────── 글자 코드 (상용 조합형) ─────────────
const CHO = [-1, -1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1, -1];
const JUNG = [-1, -1, -1, 0, 1, 2, 3, 4, -1, -1, 5, 6, 7, 8, 9, 10, -1, -1, 11, 12, 13, 14, 15, 16, -1, -1, 17, 18, 19, 20, -1, -1];
const JONG = [-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, -1, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, -1, -1];
// 자모 하나만 있는 글자 → 한글 호환 자모
const CHO_COMPAT = [0x3131, 0x3132, 0x3134, 0x3137, 0x3138, 0x3139, 0x3141, 0x3142, 0x3143, 0x3145, 0x3146, 0x3147, 0x3148, 0x3149, 0x314a, 0x314b, 0x314c, 0x314d, 0x314e];
const JONG_COMPAT = [0, 0x3131, 0x3132, 0x3133, 0x3134, 0x3135, 0x3136, 0x3137, 0x3139, 0x313a, 0x313b, 0x313c, 0x313d, 0x313e, 0x313f, 0x3140, 0x3141, 0x3142, 0x3144, 0x3145, 0x3146, 0x3147, 0x3148, 0x314a, 0x314b, 0x314c, 0x314d, 0x314e];
let eucKr;
/** KS X 1001 행 · 칸(0xA1~0xFE) → 글자 */
function ksc(row, cell, hancom = false) {
  if (eucKr === undefined) { try { eucKr = new TextDecoder('euc-kr'); } catch { eucKr = null; } }
  if (!eucKr || row < 0xa1 || row > 0xfe || cell < 0xa1 || cell > 0xfe) return null;
  const s = eucKr.decode(new Uint8Array([row, cell]));
  if (s.length !== 1 || s === '\uFFFD') return row === 0xa2 && cell === 0xe6 ? '\u20AC' : row === 0xa2 && cell === 0xe7 ? '\u00AE' : null; // € ® (1998 추가분)
  if (!hancom) return s;
  return s === '\u223C' ? '\uFF5E' : s === '\u2299' ? '\u25C9' : s; // 한컴 사적 영역 관행 (∼ → ～, ⊙ → ◉)
}
/** 조합형 기호 · 한자 영역(첫 바이트 0xD9~0xDE, 0xE0~0xF9) → KS X 1001 */
function johabKsc(ch) {
  const b1 = ch >> 8; const b2 = ch & 0xff;
  let row0;
  if (b1 >= 0xd9 && b1 <= 0xde) row0 = 0xa1 + (b1 - 0xd9) * 2;
  else if (b1 >= 0xe0 && b1 <= 0xf9) row0 = 0xca + (b1 - 0xe0) * 2;
  else return null;
  let idx;
  if (b2 >= 0x31 && b2 <= 0x7e) idx = b2 - 0x31;
  else if (b2 >= 0x91 && b2 <= 0xfe) idx = b2 - 0x91 + 78;
  else return null;
  return idx < 94 ? ksc(row0, 0xa1 + idx) : ksc(row0 + 1, 0xa1 + idx - 94);
}
const EXTRA = new Map([[0x81, 0x201c], [0x82, 0x201d], [0x83, 0x2018], [0x84, 0x2019], [0x301c, 0x2501], [0x3013, 0x250c], [0x3014, 0x252c], [0x3015, 0x2510], [0x3019, 0x2514], [0x301b, 0x2518], [0x301d, 0x2502], [0x303d, 0x25a0], [0x3366, 0x25a1], [0x3404, 0x2024], [0x3441, 0x25a0], [0x3446, 0x2192], [0x35e1, 0x2500], [0x3479, 0x25b7], [0x347a, 0x25b6], [0x2f67, 0x25b8], [0x2f00, 0x25a1],
  [0x2024, 0x30fb], [0x2058, 0x25b3], [0x205a, 0x25cb], [0x2f08, 0x25aa], [0x2f11, 0x25e6], [0x2f14, 0x25e6], [0x3157, 0x2027], [0x0480, 0x02d0], [0x1f2e, 0x306e], [0x32b0, 0xff70], [0x3067, 0x2018], [0x3068, 0x2019], [0x309b, 0xff62], [0x309d, 0xff63], [0x2f17, 0x2022], [0x2f06, 0x25a0], [0x3062, 0x201c], [0x3063, 0x201d], [0x3066, 0x2027], [0x30bb, 0x300a], [0x30bd, 0x300b], [0x3048, 0x2550], [0x37ed, 0x2550], [0x2e00, 0x24ea]]);
const SAME = new Set([0x2010, 0x2013, 0x2103, 0x2113, 0x2190, 0x2192, 0x2193, 0x2219, 0x22c5, 0x203b, 0x2022, 0x2014, 0x2223, 0x21b3]);
/** 한컴 사적 글자 영역(0x0080~0x7FFF) */
function extra(ch) {
  if (ch >= 0xa0 && ch <= 0xff) return String.fromCharCode(ch);
  if (SAME.has(ch)) return String.fromCharCode(ch);
  const m = EXTRA.get(ch); if (m) return String.fromCharCode(m);
  if (ch >= 0x3590 && ch <= 0x3599) return String.fromCharCode(0x2160 + ch - 0x3590); // Ⅰ~Ⅹ
  if (ch >= 0x36e7 && ch <= 0x36f0) return String.fromCharCode(0x2460 + ch - 0x36e7); // ①~⑩
  if (ch >= 0x37c0 && ch <= 0x37c5) return '한글과컴퓨터'[ch - 0x37c0];
  if (ch >= 0x2e01 && ch <= 0x2e07) return String.fromCharCode(0x2460 + ch - 0x2e01);
  if (ch >= 0x2e0a && ch <= 0x2e12) return String.fromCharCode(0x2460 + ch - 0x2e0a);
  if (ch >= 0x2c21 && ch <= 0x2c26) return String.fromCharCode(0x24d0 + ch - 0x2c21);
  if (ch >= 0x2c40 && ch <= 0x2c42) return String.fromCharCode(0x3260 + ch - 0x2c40);
  if (ch >= 0x3401) { const i = ch - 0x3401; const row = 0xa1 + Math.floor(i / 96); if (row <= 0xac) { const s = ksc(row, 0xa1 + (i % 96), true); if (s) return s; } }
  if (ch >= 0x4000) { const i = ch - 0x4000; const row = 0xca + Math.floor(i / 94); if (row <= 0xfd) return ksc(row, 0xa1 + (i % 94)); }
  return null;
}
/** hchar(u16) → 글자. 모르는 코드는 null */
export function hchar(ch) {
  if (ch < 0x80) return String.fromCharCode(ch);
  if (ch < 0x8000) return extra(ch);
  const ci = (ch >> 10) & 31; const vi = (ch >> 5) & 31; const ti = ch & 31;
  const cho = CHO[ci]; const jung = JUNG[vi]; const jong = JONG[ti];
  if (cho >= 0 && jung >= 0 && jong >= 0) return String.fromCharCode(0xac00 + cho * 588 + jung * 28 + jong);
  // 채움 문자로 된 낱자
  if (ci === 1 && vi === 2 && ti === 1) return '\u3000';
  if (ci === 1 && vi === 2 && jong > 0) return String.fromCharCode(JONG_COMPAT[jong]);
  if (ci === 1 && jung >= 0 && ti === 1) return String.fromCharCode(0x314f + jung);
  if (cho >= 0 && vi === 2 && ti === 1) return String.fromCharCode(CHO_COMPAT[cho]);
  // 아래아 (옛한글) → 첫가끝 자모
  if (vi === 30 && jong >= 0) { const a = '\u119E'; const t = jong ? String.fromCharCode(0x11a7 + jong) : ''; if (cho >= 0) return String.fromCharCode(0x1100 + cho) + a + t; if (ci === 1) return a + t; }
  return johabKsc(ch);
}
/** 바이트 문자열 (1바이트 ASCII, 2바이트 조합형), NUL 에서 끝 */
function bstr(b) { let s = ''; for (let i = 0; i < b.length;) { const c = b[i]; if (!c) break; if (c < 0x80) { s += String.fromCharCode(c); i++; } else { s += hchar((c << 8) | (b[i + 1] ?? 0)) ?? ''; i += 2; } } return s; }
/** hchar 문자열 (u16), 0 에서 끝 */
function hstr(b) { let s = ''; for (let i = 0; i + 1 < b.length; i += 2) { const c = b[i] | (b[i + 1] << 8); if (!c) break; s += hchar(c) ?? ''; } return s; }

// ───────────── 읽기 도구 ─────────────
class Rd {
  constructor(b, p = 0) { this.b = b; this.p = p; }
  need(n) { if (this.p + n > this.b.length) throw new Error('HWP3: 데이터가 잘렸습니다'); }
  u8() { this.need(1); return this.b[this.p++]; }
  i8() { const v = this.u8(); return v > 127 ? v - 256 : v; }
  u16() { this.need(2); const v = this.b[this.p] | (this.b[this.p + 1] << 8); this.p += 2; return v; }
  i16() { const v = this.u16(); return v > 32767 ? v - 65536 : v; }
  u32() { this.need(4); const v = (this.b[this.p] | (this.b[this.p + 1] << 8) | (this.b[this.p + 2] << 16) | (this.b[this.p + 3] << 24)) >>> 0; this.p += 4; return v; }
  bytes(n) { this.need(n); const v = this.b.subarray(this.p, this.p + n); this.p += n; return v; }
  skip(n) { this.need(n); this.p += n; }
  get left() { return this.b.length - this.p; }
}
const u16 = (b, o) => b[o] | (b[o + 1] << 8);
const i16 = (b, o) => { const v = u16(b, o); return v > 32767 ? v - 65536 : v; };
const u32 = (b, o) => (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0;

const PALETTE = ['#000000', '#0000FF', '#00FF00', '#00FFFF', '#FF0000', '#FF00FF', '#FFFF00', '#FFFFFF'];
/** 흰 바탕에 팔레트 색을 비율만큼 */
function tint(idx, pct) {
  const hex = PALETTE[idx] ?? '#000000'; const r = Math.min(100, pct);
  return `#${[1, 3, 5].map((o) => Math.floor((parseInt(hex.slice(o, o + 2), 16) * r + 255 * (100 - r)) / 100).toString(16).padStart(2, '0')).join('').toUpperCase()}`;
}

/** 글자 모양 31바이트 → rPr */
function charShape(b, o, faces) {
  const size = u16(b, o);
  const f = [...b.subarray(o + 2, o + 9)];
  const ratio = b[o + 9]; const spc = b[o + 16] > 127 ? b[o + 16] - 256 : b[o + 16];
  const shadeColor = b[o + 23]; const color = b[o + 24]; const shadeRatio = b[o + 25]; const attr = b[o + 26];
  const r = { sz: Math.round(size / 25 * 10) / 10 || 10 };
  const ea = faces[0][f[0]]; const lat = faces[1][f[1]];
  if (ea) r.fontEa = ea;
  if (lat || ea) r.font = lat || ea;
  if (attr & 1) r.i = true;
  if (attr & 2) r.b = true;
  if (attr & 4) r.u = true;
  if (attr & 8) r.outline = true;
  if (attr & 16) r.shadow = true;
  if (attr & 32) r.va = 'super'; else if (attr & 64) r.va = 'sub';
  if (color) r.color = PALETTE[color] ?? '#000000';
  if (shadeRatio) r.shd = tint(shadeColor, shadeRatio);
  if (ratio && ratio !== 100) r.scale = ratio;
  if (spc) r.spc = Math.round(spc / 100 * r.sz * 100) / 100;
  return r;
}
const ALIGN = ['justify', 'left', 'right', 'center', 'distribute', 'distribute', 'justify', 'distribute'];
/** 문단 모양 187바이트 → pPr */
function paraShape(b, o) {
  const left = u16(b, o); const right = u16(b, o + 2); const indent = i16(b, o + 4); const ls = u16(b, o + 6); const after = u16(b, o + 8);
  const align = b[o + 11];
  const p = { align: ALIGN[align] ?? 'justify' };
  const ind = {};
  if (left) ind.l = px(left);
  if (right) ind.r = px(right);
  if (indent > 0) ind.first = px(indent);
  else if (indent < 0) ind.hanging = px(-indent);
  if (Object.keys(ind).length) p.ind = ind;
  const before = u16(b, o + 183);
  p.spc = { before: px(before), after: px(after), ...(ls & 0x8000 ? { line: px(ls & 0x7fff), rule: 'exact' } : { line: (ls || 160) / 100, rule: 'pct' }) };
  // 탭: 기본(1000 hunit 간격) 이 아닌 것만
  const tabs = [];
  for (let k = 0; k < 40; k++) {
    const q = o + 12 + k * 4; const kind = b[q]; const lead = b[q + 1]; const pos = u16(b, q + 2);
    if (!pos || (kind === 0 && lead === 0 && pos === 1000 * (k + 1))) continue;
    tabs.push({ pos: px(pos), align: ['left', 'right', 'center', 'decimal'][kind] ?? 'left', ...(lead ? { leader: 'dot' } : {}) });
  }
  if (tabs.length) p.tabs = tabs;
  const cols = b[o + 172];
  if (cols > 1) p._cols = { n: cols, gap: px(u16(b, o + 174)), line: !!b[o + 173] };
  const shade = b[o + 180];
  if (shade) p.shd = tint(0, shade);
  return p;
}
const OUTLINE = [['I', '', '.'], ['1', '', '.'], ['가', '', '.'], ['1', '', ')'], ['가', '', ')'], ['1', '(', ')'], ['가', '(', ')']];
const GANADA = '가나다라마바사아자차카타파하';
const roman = (n) => { const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; for (const [v, c] of t) while (n >= v) { s += c; n -= v; } return s; };
function outlineText(level, n) {
  const [kind, before, after] = OUTLINE[level] ?? ['1', '', '.'];
  const v = kind === 'I' ? roman(n) : kind === '가' ? (GANADA[(n - 1) % 14] ?? String(n)) : String(n);
  return `${before}${v}${after} `;
}

// ───────────── 문단 목록 ─────────────
class Ctx {
  constructor(doc, faces, styles) { this.doc = doc; this.faces = faces; this.styles = styles; this.pics = []; this.depth = 0; this.warn = new Set(); this.noteSeq = 0; this.hf = { headers: {}, footers: {} }; this.cols = null; }
}
function paragraphs(rd, ctx) {
  if (++ctx.depth > 16) throw new Error('HWP3: 문단 목록이 너무 깊게 겹쳤습니다');
  const out = [];
  let prevShape = null;
  try {
    for (;;) {
      if (rd.left < 3) break;
      const follow = rd.u8(); const n = rd.u16();
      if (n === 0) { rd.skip(Math.min(40, rd.left)); break; }
      const lines = rd.u16();
      if (n > 60000 || lines > 4096) { ctx.warn.add('문단 머리가 이상하여 나머지를 건너뛰었습니다'); break; }
      const hasCs = rd.u8(); rd.skip(1); rd.skip(4);
      const styleIdx = rd.u8();
      const rep = rd.bytes(31);
      let shape = prevShape;
      if (follow === 0) shape = rd.bytes(187);
      prevShape = shape;
      rd.skip(lines * 14);
      const cs = new Array(n);
      if (hasCs) for (let i = 0; i < n; i++) { const flag = rd.u8(); if (flag !== 1) cs[i] = rd.bytes(31); }
      out.push(...paragraph(rd, ctx, n, { rep, shape, cs, styleIdx }));
    }
  } finally { ctx.depth--; }
  return out;
}
function paragraph(rd, ctx, n, info) {
  const baseR = charShape(info.rep, 0, ctx.faces);
  const pPr = info.shape ? paraShape(info.shape, 0) : {};
  if (pPr._cols && !ctx.cols) ctx.cols = pPr._cols;
  delete pPr._cols;
  const st = ctx.styles[info.styleIdx];
  const style = st?.key ?? 'Normal';
  const out = [];
  let cur = { ...para(''), style, pPr: { ...pPr }, runs: [] };
  const flush = (final) => { if (cur.runs.length || (final && !out.length)) { cur.runs = normalizeRuns(cur.runs); out.push(cur); } cur = { ...para(''), style, pPr: { ...pPr }, runs: [] }; };
  // 글자 모양: 위치 i 에서 바뀐 모양이 다음 바뀜까지 이어짐 (처음은 대표 모양)
  const shapes = new Array(n); let curShape = baseR;
  for (let k = 0; k < n; k++) { if (info.cs[k]) curShape = charShape(info.cs[k], 0, ctx.faces); shapes[k] = curShape; }
  const rAt = (i) => shapes[i] ?? baseR;
  let buf = ''; let bufR = null;
  const flushText = () => { if (buf) { cur.runs.push({ ...bufR, text: buf }); buf = ''; } };
  const addText = (s, r) => { if (!s) return; if (buf && r !== bufR && JSON.stringify(r) !== JSON.stringify(bufR)) flushText(); if (!buf) bufR = r; buf += s; };
  let i = 0;
  while (i < n) {
    const at = i;
    const ch = rd.u16(); i++;
    if (ch >= 32) { addText(hchar(ch) ?? '', rAt(at)); continue; }
    if (ch === 13) break;
    if (ch === 0) continue;
    if (ch === 9) { rd.skip(6); i += 3; flushText(); cur.runs.push({ ...rAt(at), obj: 'tab' }); continue; }
    if (ch === 18) {
      const b = rd.bytes(6); i += 3; const kind = u16(b, 0); const num = u16(b, 2);
      flushText();
      if (kind === 0) cur.runs.push({ ...rAt(at), obj: 'field', code: 'PAGE', result: String(num || 1) });
      else if (kind >= 3) addText(String(num), rAt(at));
      continue;
    }
    if (ch === 28) {
      const b = rd.bytes(62); i += 31;
      if (u16(b, 0) === 1) { const lvl = Math.min(b[3], 6); const num = u16(b, 4 + lvl * 2); if (num > 0) addText(outlineText(lvl, num), rAt(at)); cur.pPr.outline = lvl; }
      continue;
    }
    const simple = { 19: 6, 20: 6, 21: 6, 22: 22, 23: 8, 24: 4, 25: 4, 26: 244, 30: 2, 31: 2 }[ch];
    if (simple != null) {
      rd.skip(simple); i += simple / 2;
      if (ch === 24) addText('-', rAt(at));
      else if (ch === 30 || ch === 31) addText(ch === 30 ? '\u00A0' : ' ', rAt(at));
      continue;
    }
    // 8바이트 머리 + 종류별 내용
    const hv = rd.u32(); rd.u16(); i += 3;
    flushText();
    switch (ch) {
      case 10: {
        const r = tableLike(rd, ctx, rAt(at));
        if (r.runs) cur.runs.push(...r.runs);
        if (r.blocks) { flush(); out.push(...r.blocks); }
        break;
      }
      case 11: {
        const r = picture(rd, ctx);
        if (r.runs) cur.runs.push(...r.runs);
        if (r.blocks?.length) { flush(); out.push(...r.blocks); }
        break;
      }
      case 14: rd.skip(84); break; // 선
      case 15: rd.skip(8); paragraphs(rd, ctx); break; // 숨은 설명 (화면에 안 보임)
      case 16: { // 머리말 · 꼬리말
        const b = rd.bytes(10);
        const blocks = paragraphs(rd, ctx);
        const which = ['default', 'even', 'odd'][b[9]] ?? 'default';
        const key = which === 'odd' ? 'default' : which;
        (b[8] === 1 ? ctx.hf.footers : ctx.hf.headers)[key] = blocks.length ? blocks : [para('')];
        if (which === 'even') ctx.doc.settings.evenOdd = true;
        break;
      }
      case 17: { // 각주 · 미주
        const b = rd.bytes(14);
        const kind = u16(b, 10) === 1 ? 'end' : 'foot';
        const blocks = paragraphs(rd, ctx);
        const id = `n${++ctx.noteSeq}`;
        ctx.doc.notes[kind][id] = blocks.length ? blocks : [para('')];
        cur.runs.push({ ...rAt(at), va: 'super', obj: 'note', id, kind });
        break;
      }
      case 5: if (hv > 0 && hv < 1e6) rd.skip(hv); break; // 필드
      case 6: rd.skip(34); break; // 책갈피
      case 7: rd.skip(76); break; // 날짜 형식
      case 8: { // 날짜 코드: 형식 + 날짜 word[4]
        const b = rd.bytes(88);
        const y = u16(b, 74); const m = u16(b, 76); const d = u16(b, 78);
        if (y > 1900 && m >= 1 && m <= 12) addText(`${y}. ${m}. ${d}.`, rAt(at));
        break;
      }
      case 29: if (hv < 1e6) rd.skip(hv); break; // 상호 참조
      default: break; // 예약 코드: 머리만
    }
  }
  flushText();
  flush(true);
  return out;
}

/** 표 · 글상자 · 수식 · 단추 (ch 10) */
function tableLike(rd, ctx, rPr) {
  const info = rd.bytes(84);
  const count = u16(info, 80) || 1;
  if (count * 27 > rd.left) throw new Error(`HWP3: 표 칸 수가 이상합니다 (${count})`);
  const cellInfo = rd.bytes(27 * count);
  const cellBlocks = [];
  for (let k = 0; k < count; k++) cellBlocks.push(paragraphs(rd, ctx));
  const caption = paragraphs(rd, ctx);
  const kind = u16(info, 78);
  const text = (blocks) => blocks.map((b) => (b.runs ?? []).map((r) => r.text ?? '').join('')).join('\n');
  if (kind === 2) { // 수식: 첫 칸 글이 한글 수식 스크립트
    const script = text(cellBlocks[0] ?? []).trim();
    const runs = script ? [{ ...rPr, obj: 'eq', latex: hwpEqToLatex(script) }] : [];
    return { runs, blocks: caption.length ? caption : undefined };
  }
  if (kind === 3 || (u16(info, 14) & 0x10)) return { runs: [{ ...rPr, text: text(cellBlocks.flat()) }] }; // 단추 · 하이퍼텍스트
  // 칸 기하 → 격자 (왼쪽 · 위쪽 변만 경계 — kordoc hwp3/table 규칙)
  const geo = [];
  let right = 0; let bottom = 0;
  for (let k = 0; k < count; k++) {
    const o = k * 27; const x = u16(cellInfo, o + 4); const y = u16(cellInfo, o + 6);
    const g = { x, y, r: x + u16(cellInfo, o + 8), b: y + u16(cellInfo, o + 10), o };
    geo.push(g); right = Math.max(right, g.r); bottom = Math.max(bottom, g.b);
  }
  const xs = [...new Set([...geo.map((g) => g.x), right])].sort((a, b) => a - b);
  const ys = [...new Set([...geo.map((g) => g.y), bottom])].sort((a, b) => a - b);
  const next = (g, axis) => {
    let best;
    for (const h of geo) {
      if (h === g) continue;
      if (axis === 'x' ? h.x > g.x && h.x >= g.r - 10 && h.y < g.b && h.b > g.y : h.y > g.y && h.y >= g.b - 10 && h.x < g.r && h.r > g.x) {
        const v = axis === 'x' ? h.x : h.y; if (best === undefined || v < best) best = v;
      }
    }
    return best;
  };
  const line = (v) => (v === 1 ? { style: 'single', w: 0.5, color: '#000000' } : v === 2 ? { style: 'single', w: 1.5, color: '#000000' } : v === 3 ? { style: 'dotted', w: 0.5, color: '#000000' } : v === 4 ? { style: 'double', w: 1.5, color: '#000000' } : { style: 'none' });
  const cells = geo.map((g, k) => {
    const c = xs.indexOf(g.x); const r = ys.indexOf(g.y);
    const nx = next(g, 'x'); const ny = next(g, 'y');
    const c2 = nx !== undefined ? xs.indexOf(nx) : xs.length - 1;
    const r2 = ny !== undefined ? ys.indexOf(ny) : ys.length - 1;
    const o = g.o;
    const shade = cellInfo[o + 24];
    const cell = { r, c, rs: Math.max(1, r2 - r), cs: Math.max(1, c2 - c), h: px(g.b - g.y), blocks: cellBlocks[k], va: cellInfo[o + 19] === 1 ? 1 : cellInfo[o + 19] === 2 ? 2 : 0,
      borders: { left: line(cellInfo[o + 20]), right: line(cellInfo[o + 21]), top: line(cellInfo[o + 22]), bottom: line(cellInfo[o + 23]) } };
    if (shade > 0 && shade <= 100) cell.shd = tint(u16(cellInfo, o + 2) & 7, shade);
    return cell;
  });
  // 같은 자리 겹침 정리 (먼저 나온 칸)
  const seen = new Set();
  const uniq = cells.filter((c) => { const k = `${c.r},${c.c}`; if (seen.has(k)) { c.blocks.length && cells.find((d) => `${d.r},${d.c}` === k)?.blocks.push(...c.blocks); return false; } seen.add(k); return true; });
  const grid = xs.slice(1).map((x, k) => Math.max(4, px(x - xs[k])));
  const mar = { l: px(i16(info, 34)), r: px(i16(info, 36)), t: px(i16(info, 38)), b: px(i16(info, 40)) };
  const t = gridTable(uniq, ys.length - 1, xs.length - 1, grid, { look: { firstRow: false, firstCol: false, bandRows: false }, cellMar: mar });
  if (i16(info, 10) === -3) t.tblPr.align = 'center'; else if (i16(info, 10) === -2) t.tblPr.align = 'right';
  if (kind === 1 && count === 1) t.tblPr.textbox = true; // 글상자 → 1×1 표 (한컴 변환과 같음)
  return { blocks: [t, ...caption] };
}

/** 그림 · 그리기 개체 (ch 11) */
function picture(rd, ctx) {
  const info = rd.bytes(348);
  const nExt = u32(info, 0);
  const ext = nExt > 0 && nExt <= rd.left ? rd.bytes(nExt) : null;
  const w = px(u16(info, 42)) || 100; const h = px(u16(info, 44)) || 60;
  const ref = info[8]; const wrapK = info[9];
  const type = info[74];
  const blocks = [];
  const runs = [];
  if (type === 3 && ext) {
    // 그리기 개체: 안의 글상자 글만 꺼냄
    try { for (const list of drawingTextLists(ext)) blocks.push(...paragraphs(new Rd(list), ctx)); } catch { ctx.warn.add('일부 그리기 개체를 읽지 못했습니다'); }
  } else {
    const name = bstr(info.subarray(83, 83 + 256));
    const run = { obj: 'img', id: uid('i'), media: null, w, h, alt: name };
    if (ref !== 0) {
      const hx = i16(info, 10); const vy = i16(info, 12);
      run.float = { wrap: wrapK === 1 ? 'behind' : wrapK === 2 ? 'square' : 'topBottom', relH: ref === 2 ? 'page' : ref === 3 ? 'page' : 'column', relV: ref >= 2 ? 'page' : 'paragraph', x: hx >= 0 ? px(hx) : 0, y: vy >= 0 ? px(vy) : 0 };
      if (hx === -3) run.float.alignH = 'center'; else if (hx === -2) run.float.alignH = 'right';
      if (run.float.wrap === 'square') run.float.side = 'left';
    }
    ctx.pics.push({ name, run });
    runs.push(run);
  }
  blocks.push(...paragraphs(rd, ctx)); // 캡션
  return { runs, blocks };
}
/** 그리기 개체 트리 → 글상자 문단 목록 바이트들 (kordoc hwp3/drawing 규칙) */
function drawingTextLists(ext) {
  const rd = new Rd(ext);
  const out = [];
  const head = rd.u32(); rd.skip(4); const count = rd.u32(); rd.skip(16);
  if (head > 24) rd.skip(621);
  if (!count) return out;
  let budget = 4096;
  const obj = () => {
    const start = rd.p; const len = rd.u32(); const type = rd.u16(); const conn = rd.u16();
    rd.skip(80);
    const opt = rd.u32();
    if (opt & (1 << 17)) rd.skip(32);
    if (opt & (1 << 16)) rd.skip(28);
    if (opt & (1 << 18)) rd.skip(278);
    const end = start + 4 + len;
    if (end > rd.p && end - rd.p <= 65536 && end <= ext.length) {
      const surplus = end - rd.p; let got = false;
      if (type !== 6 && surplus >= 8) { const a = rd.u32(); const b2 = rd.u32(); if (a + b2 === surplus - 8 && b2 > 0) { rd.skip(a); out.push(rd.bytes(b2)); got = true; } }
      if (!got) rd.p = end;
    }
    const pts = (attrs) => { rd.skip(4); const n = rd.u32(); rd.skip(4); rd.skip(n * 8 + (attrs ? n : 0)); };
    switch (type) {
      case 0: case 2: case 3: rd.skip(8); break;
      case 1: case 4: rd.skip(12); break;
      case 5: case 7: pts(false); break;
      case 6: { rd.skip(4); const l2 = rd.u32(); if (l2) out.push(rd.bytes(l2)); break; }
      case 8: rd.skip(24); break;
      case 9: break;
      case 10: case 11: pts(true); break;
      default: { const a = rd.u32(); rd.skip(a); const b2 = rd.u32(); rd.skip(b2); }
    }
    return conn;
  };
  const walk = (depth) => {
    if (depth > 256) throw new Error('too deep');
    for (;;) { if (budget-- <= 0) throw new Error('too many'); const c = obj(); if (c & 2) walk(depth + 1); if (!(c & 1)) return; }
  };
  walk(0);
  return out;
}

function imageExt(b) {
  if (b[0] === 0xff && b[1] === 0xd8) return 'jpg';
  if (b[0] === 0x89 && b[1] === 0x50) return 'png';
  if (b[0] === 0x47 && b[1] === 0x49) return 'gif';
  if (b[0] === 0x42 && b[1] === 0x4d) return 'bmp';
  if ((b[0] === 0xd7 && b[1] === 0xcd) || (b[0] === 1 && b[1] === 0 && b[2] === 9)) return 'wmf';
  if (b[0] === 0x49 && b[1] === 0x49) return 'tif';
  return 'bin';
}

// ───────────── 암호 ─────────────
/** 열기 암호 → DES 키 (UTF-16LE 바이트를 7바이트 고리에 섞음) */
function desKey(password) {
  const ring = new Uint8Array(7); let k = 0;
  for (let i = 0; i < password.length; i++) {
    const u = password.charCodeAt(i);
    for (const byte of [u & 255, u >> 8]) { const s = k % 7; const m = ring[s] ^ byte; ring[s] = ((m << 1) | (m >> 7)) & 255; k++; }
  }
  const key = new Uint8Array(8);
  for (let bit = 0; bit < 56; bit++) key[Math.floor(bit / 7)] |= ((ring[bit >> 3] >> (7 - (bit & 7))) & 1) << (7 - (bit % 7));
  return key;
}
export const isHwp3 = (b) => b?.length > 30 && new TextDecoder('latin1').decode(b.subarray(0, 23)) === 'HWP Document File V3.00';

export function readHwp3(bytes, opts = {}) {
  if (!isHwp3(bytes)) throw new Error('한글 3.0 문서가 아닙니다.');
  const FIX = 30 + 128 + 1008;
  if (bytes.length < FIX) throw new Error('한글 3.0 문서의 머리가 잘렸습니다.');
  const di = 30;
  const encrypted = u16(bytes, di + 96);
  let compressed = bytes[di + 124];
  const infoLen = u16(bytes, di + 126);
  let body = bytes.subarray(FIX + infoLen);
  if (encrypted) {
    if (!opts.password) { const e = new Error('암호로 보호된 한글 97 문서입니다. 열기 암호를 입력하세요.'); e.code = 'PASSWORD'; throw e; }
    if (!compressed) throw new Error('지원하지 않는 한글 97 암호 방식입니다 (압축되지 않은 본문).');
    const plain = desEcbDecrypt(body.subarray(0, body.length - (body.length % 8)), desKey(opts.password));
    let inflated;
    try { inflated = inflate(plain); } catch { inflated = null; }
    if (!inflated || inflated.length <= 256) { const e = new Error('암호가 맞지 않습니다.'); e.code = 'PASSWORD'; e.wrong = true; throw e; }
    body = inflated.subarray(256);
    compressed = 0;
  } else if (compressed) {
    try { body = inflate(body); } catch (e) { throw new Error(`한글 3.0 문서 본문을 풀지 못했습니다: ${e.message}`); }
  }
  const doc = newDoc();
  doc.media = {}; doc.lists = {};
  // 쪽 설정
  const sec = newSection();
  const hgt = u16(bytes, di + 6); const w = u16(bytes, di + 8);
  const land = bytes[di + 5] !== 0; // 용지 방향: 0 세로, 1 가로 (크기는 세로 기준으로 저장)
  if (w && hgt) sec.page = land ? { w: px(Math.max(w, hgt)), h: px(Math.min(w, hgt)), orient: 'landscape' } : { w: px(w), h: px(hgt), orient: 'portrait' };
  const top = u16(bytes, di + 10); const bot = u16(bytes, di + 12); const hdr = u16(bytes, di + 18); const ftr = u16(bytes, di + 20);
  sec.margin = { t: px(top + hdr), b: px(bot + ftr), l: px(u16(bytes, di + 14)), r: px(u16(bytes, di + 16)), header: px(top), footer: px(bot), gutter: px(u16(bytes, di + 22)) };
  const startPage = u16(bytes, di + 96 + 2);
  if (startPage > 1) sec.pgNum = { start: startPage };
  // 요약 정보
  const sum = 30 + 128;
  const props = { title: hstr(bytes.subarray(sum, sum + 112)), subject: hstr(bytes.subarray(sum + 112, sum + 224)), author: hstr(bytes.subarray(sum + 224, sum + 336)) };
  for (const [k, v] of Object.entries(props)) if (v) doc.props[k] = v;
  const rd = new Rd(body);
  const warnings = [];
  // 글꼴 (언어 7개)
  const faces = [];
  for (let l = 0; l < 7; l++) { const n = rd.u16(); const list = []; for (let k = 0; k < n; k++) list.push(bstr(rd.bytes(40)).trim()); faces.push(list); }
  // 스타일
  const nStyles = rd.u16();
  const styles = [];
  for (let k = 0; k < nStyles; k++) {
    const name = bstr(rd.bytes(20)).trim(); const cs = rd.bytes(31); const ps = rd.bytes(187);
    const outline = /^개요\s*(\d)/.exec(name);
    const key = k === 0 ? 'Normal' : outline ? `Heading${outline[1]}` : `hwp${k}`;
    styles.push({ key });
    const rPr = charShape(cs, 0, faces); const pPr = paraShape(ps, 0); delete pPr._cols;
    if (key === 'Normal') { doc.defaults = { rPr: { ...doc.defaults.rPr, ...rPr }, pPr: { ...doc.defaults.pPr, ...pPr } }; doc.styles.Normal = { ...doc.styles.Normal, name: name || '보통' }; }
    else doc.styles[key] = { ...(doc.styles[key] ?? {}), id: key, type: 'para', name: name || key, basedOn: 'Normal', quick: !!outline, order: 100 + k, pPr: {}, rPr: {} };
  }
  const ctx = new Ctx(doc, faces, styles);
  let blocks = [];
  try { blocks = paragraphs(rd, ctx); } catch (e) { warnings.push(`본문 일부를 읽지 못했습니다: ${e.message}`); }
  // 추가 정보 블록: #1 = 문서에 포함된 그림
  const pics = new Map();
  try {
    while (rd.left >= 8) {
      const id = rd.u32(); if (!id) break;
      const len = rd.u32(); if (len > rd.left) break;
      const data = rd.bytes(len);
      if (id === 1 && data.length > 32) pics.set(bstr(data.subarray(0, 16)).trim(), data.subarray(32));
    }
  } catch { /* 끝 */ }
  let missing = 0;
  for (const { name, run } of ctx.pics) {
    const data = pics.get(name.trim()) ?? pics.get(name.split(/[\\/]/).pop().trim());
    if (!data) { missing++; run.media = null; continue; }
    const mid = uid('m');
    doc.media[mid] = mediaDataUrl(data, `x.${imageExt(data)}`);
    run.media = mid;
  }
  if (missing) warnings.push(`그림 ${missing}개는 문서 밖의 파일(연결 그림 · OLE 개체)이라 표시하지 못합니다.`);
  // 그림 자료가 없는 그림 칸은 대체 글로
  const fix = (list) => { for (const b of list) { if (b.runs) b.runs = b.runs.map((r) => (r.obj === 'img' && !r.media ? { text: `[그림: ${r.alt || '연결 그림'}]` } : r)); if (b.rows) for (const row of b.rows) for (const c of row.cells) fix(c.blocks); } };
  fix(blocks);
  for (const k of ['foot', 'end']) for (const v of Object.values(doc.notes[k])) fix(v);
  sec.headers = ctx.hf.headers; sec.footers = ctx.hf.footers;
  if (ctx.cols) sec.cols = { n: ctx.cols.n, gap: ctx.cols.gap || 24, sep: ctx.cols.line };
  doc.section = sec;
  doc.body = blocks.length ? blocks : [para('')];
  for (const w2 of ctx.warn) warnings.push(w2);
  return { doc, warnings };
}
