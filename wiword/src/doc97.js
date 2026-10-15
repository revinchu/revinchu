// Word 97-2003 문서 (.doc, MS-DOC) 읽기 — DOM 없음
// 조각 표(Clx) 로 모든 이야기(본문 · 각주 · 머리글 · 미주 · 글상자)의 글자를 모으고,
// 글자 서식(CHPX FKP) · 문단 서식(PAPX FKP) · 글꼴 표 · 목록(PlfLst/PlfLfo) · 구역(PlcfSed) · 머리글(PlcfHdd) ·
// 각주/미주(Plcffnd/Plcfend) · 표(셀 표시 \x07) · 필드(하이퍼링크) · 그림(글자처럼: PICF + OfficeArt BLIP, 떠 있는: PlcSpaMom + OfficeArt) 을 읽음
import { openCfb } from './cfb.js';
import { inflate } from './zip.js';
import { newDoc, newSection, para, uid, normalizeRuns } from './model.js';
import { mediaDataUrl } from './docx.js';
import { md5, sha1, rc4 } from './crypto.js';

const u16 = (b, o) => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);
const i16 = (b, o) => { const v = u16(b, o); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, o) => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);
const i32 = (b, o) => (o + 4 <= b.length ? b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24) : 0);
const TW = 1 / 15; // twip → px
const CP1252 = '€\x81‚ƒ„…†‡ˆ‰Š‹Œ\x8DŽ\x8F\x90‘’“”•–—˜™š›œ\x9DžŸ';
const ICO = [null, '#000000', '#0000FF', '#00FFFF', '#00FF00', '#FF00FF', '#FF0000', '#FFFF00', '#FFFFFF', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#C0C0C0'];
const HL = [null, 'black', 'blue', 'cyan', 'brightGreen', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];
const NFC = { 0: 'decimal', 1: 'upperRoman', 2: 'lowerRoman', 3: 'upperLetter', 4: 'lowerLetter', 5: 'ordinal', 22: 'decimalZero', 23: 'bullet', 24: 'ganada', 25: 'chosung', 26: 'decimalEnclosedCircle', 47: 'bullet' };

/** sprm 목록 → [{ op, o(인자 시작), size }] */
function sprms(b, o, end) {
  const out = [];
  while (o + 2 <= end) {
    const op = u16(b, o); o += 2;
    const spra = op >> 13;
    let size = [1, 1, 2, 4, 2, 2, -1, 3][spra];
    if (size === -1) { if (op === 0xd608 || op === 0xd606) { size = u16(b, o) + 1; } else { size = b[o] + 1; } }
    out.push({ op, o, size });
    o += size;
  }
  return out;
}
function chpFrom(b, list, fonts) {
  const r = {};
  for (const { op, o } of list) {
    switch (op) {
      case 0x0835: r.b = b[o] === 1 || b[o] === 0x81 ? true : b[o] === 0 ? false : r.b; break;
      case 0x0836: r.i = b[o] === 1 || b[o] === 0x81 ? true : b[o] === 0 ? false : r.i; break;
      case 0x0837: r.s = !!b[o]; break;
      case 0x2a53: r.dstrike = !!b[o]; break;
      case 0x0838: r.outline = !!b[o]; break;
      case 0x0839: r.shadow = !!b[o]; break;
      case 0x083a: r.smallCaps = !!b[o]; break;
      case 0x083b: r.caps = !!b[o]; break;
      case 0x083c: r.hidden = !!b[o]; break;
      case 0x4a43: r.sz = u16(b, o) / 2; break;
      case 0x2a42: if (ICO[b[o]]) r.color = ICO[b[o]]; break;
      case 0x6870: { const c = u32(b, o); if (!(c & 0xff000000)) r.color = `#${[c & 0xff, (c >> 8) & 0xff, (c >> 16) & 0xff].map((x) => x.toString(16).padStart(2, '0')).join('').toUpperCase()}`; break; }
      case 0x2a3e: { const k = b[o]; r.u = k === 0 ? false : k === 3 ? 'double' : k === 4 ? 'dotted' : k === 7 ? 'dash' : k === 11 ? 'wave' : true; break; }
      case 0x2a0c: if (HL[b[o]]) r.hl = HL[b[o]]; break;
      case 0x2a48: r.va = b[o] === 1 ? 'super' : b[o] === 2 ? 'sub' : undefined; break;
      case 0x4a4f: { const f = fonts[u16(b, o)]; if (f) r.font = f; break; }
      case 0x4a50: { const f = fonts[u16(b, o)]; if (f) r.fontEa = f; break; }
      case 0x4a51: break;
      case 0x8840: r.spc = i16(b, o) / 20; break;
      case 0x6a03: r._pic = u32(b, o); break; // 그림 위치 (Data 스트림)
      case 0x0855: r._spec = !!b[o]; break;
      case 0x0806: r._data = !!b[o]; break;
      case 0x080a: r._ole = !!b[o]; break;
      default: break;
    }
  }
  for (const k of Object.keys(r)) if (r[k] === undefined) delete r[k];
  return r;
}
function papFrom(b, list) {
  const p = {};
  for (const { op, o } of list) {
    switch (op) {
      case 0x2403: case 0x2461: p.align = ['left', 'center', 'right', 'justify', 'distribute'][b[o]] ?? 'left'; break;
      case 0x840f: case 0x845e: p.l = i16(b, o) * TW; break;
      case 0x840e: case 0x845d: p.r = i16(b, o) * TW; break;
      case 0x8411: case 0x8460: p.first = i16(b, o) * TW; break;
      case 0xa413: p.before = u16(b, o) * TW; break;
      case 0xa414: p.after = u16(b, o) * TW; break;
      case 0x6412: { const dya = i16(b, o); const mult = u16(b, o + 2); p.line = mult ? { line: Math.abs(dya) / 240, rule: 'auto' } : { line: Math.abs(dya) * TW, rule: dya < 0 ? 'exact' : 'atLeast' }; break; }
      case 0x2416: p.inTable = !!b[o]; break;
      case 0x2417: p.ttp = !!b[o]; break;
      case 0x6649: p.depth = u32(b, o); break;
      case 0x244b: p.inner = !!b[o]; break;
      case 0x244c: p.ttp = p.ttp || !!b[o]; break;
      case 0x2640: p.outline = b[o]; break;
      case 0x260a: p.ilvl = b[o]; break;
      case 0x460b: p.ilfo = u16(b, o); break;
      case 0x2405: p.keepLines = !!b[o]; break;
      case 0x2406: p.keepNext = !!b[o]; break;
      case 0x2407: p.pageBreakBefore = !!b[o]; break;
      case 0xd608: { // 표 정의 (TDefTable): 칸 경계 (twip)
        const cb = u16(b, o); const n = b[o + 2];
        if (n > 0 && n < 64 && 3 + (n + 1) * 2 <= cb + 1) { const xs = []; for (let k = 0; k <= n; k++) xs.push(i16(b, o + 3 + k * 2)); p.cellX = xs; }
        break;
      }
      default: break;
    }
  }
  return p;
}

// ───────────── OfficeArt (그림 자료) ─────────────
const BLIP = { 0xf01a: 'emf', 0xf01b: 'wmf', 0xf01c: 'pict', 0xf01d: 'jpg', 0xf01e: 'png', 0xf01f: 'bmp', 0xf029: 'tif', 0xf02a: 'jpg' };
/** BLIP 레코드 하나 → { ext, data } */
function blipData(b, o) {
  const verInst = u16(b, o); const type = u16(b, o + 2); const len = u32(b, o + 4);
  const ext = BLIP[type];
  if (!ext) return null;
  const inst = verInst >> 4;
  const body = o + 8;
  // 두 번째 UID 가 있는 inst (홀수 짝)
  const twoUid = { emf: 0x3d5, wmf: 0x217, pict: 0x543, jpg: 0x46b, png: 0x6e1, bmp: 0x7a9, tif: 0x6e5 }[ext] === inst || inst === 0x6e3 || inst === 0x46d;
  let p = body + 16 + (twoUid ? 16 : 0);
  if (ext === 'emf' || ext === 'wmf' || ext === 'pict') {
    // 메타파일 머리 34바이트: cbSize · rcBounds · ptSize · cbSave · compression · filter
    const cbSave = u32(b, p + 28); const comp = b[p + 32];
    p += 34;
    let data = b.subarray(p, Math.min(body + len, p + cbSave));
    if (comp === 0) { try { data = inflate(data[0] === 0x78 ? data.subarray(2) : data); } catch { /* 그대로 */ } }
    if (ext === 'wmf' && !(data[0] === 0xd7 && data[1] === 0xcd)) {
      // 위치 지정 머리(placeable)가 없는 WMF → 22바이트 머리를 붙여 줌
      const h = new Uint8Array(22 + data.length); const dv = new DataView(h.buffer);
      dv.setUint32(0, 0x9ac6cdd7, true); dv.setInt16(10, 1000, true); dv.setInt16(12, 1000, true); dv.setUint16(14, 1440, true);
      let chk = 0; for (let k = 0; k < 10; k++) chk ^= dv.getUint16(k * 2, true); dv.setUint16(20, chk, true);
      h.set(data, 22); data = h;
    }
    return { ext, data };
  }
  p += 1; // tag
  return { ext, data: b.subarray(p, body + len) };
}
/** 레코드 범위에서 첫 BLIP 찾기 (FBSE 안 · 컨테이너 안까지) */
function findBlip(b, o, end, delay) {
  while (o + 8 <= end) {
    const verInst = u16(b, o); const type = u16(b, o + 2); const len = u32(b, o + 4);
    if (len > end - o - 8 + 16) break;
    if (BLIP[type]) { const r = blipData(b, o); if (r) return r; }
    if (type === 0xf007) { // FBSE
      const size = u32(b, o + 8 + 20); const foDelay = u32(b, o + 8 + 28); const cbName = b[o + 8 + 33];
      const inner = o + 8 + 36 + cbName;
      if (inner + 8 <= o + 8 + len) { const r = findBlip(b, inner, o + 8 + len, delay); if (r) return r; }
      if (delay && size && foDelay < delay.length) { const r = blipData(delay, foDelay); if (r) return r; }
    } else if ((verInst & 0xf) === 0xf) { const r = findBlip(b, o + 8, o + 8 + len, delay); if (r) return r; }
    o += 8 + len;
  }
  return null;
}
/** OfficeArtContent (표 스트림): BStore 그림 목록 + 도형(spid → 그림 번호 · 글상자) */
function readArt(tbl, fc, lcb, wd) {
  const art = { blips: [], shapes: new Map() };
  if (!lcb) return art;
  const end = fc + lcb;
  const walk = (o, stop, depth) => {
    while (o + 8 <= stop && depth < 20) {
      const verInst = u16(tbl, o); const type = u16(tbl, o + 2); const len = u32(tbl, o + 4);
      const body = o + 8;
      if (body + len > stop + 8) break;
      if (type === 0xf001) { // BStore
        let q = body;
        while (q + 8 <= body + len) { const l = u32(tbl, q + 4); art.blips.push(u16(tbl, q + 2) === 0xf007 ? findBlip(tbl, q, q + 8 + l, wd) : blipData(tbl, q)); q += 8 + l; }
      } else if (type === 0xf004) { // SpContainer
        let spid = 0; const props = {};
        let q = body;
        while (q + 8 <= body + len) {
          const t = u16(tbl, q + 2); const l = u32(tbl, q + 4); const vi = u16(tbl, q);
          if (t === 0xf00a) spid = u32(tbl, q + 8);
          if (t === 0xf00b || t === 0xf121 || t === 0xf122) { // 속성 표
            const n = vi >> 4;
            for (let k = 0; k < n; k++) { const op = u16(tbl, q + 8 + k * 6) & 0x3fff; const v = u32(tbl, q + 8 + k * 6 + 2); props[op] = v; }
          }
          q += 8 + l;
        }
        if (spid) art.shapes.set(spid, { pib: props[0x104], txid: props[0x80], fill: props[0x181], line: props[0x1c0] });
      } else if ((verInst & 0xf) === 0xf) walk(body, body + len, depth + 1);
      o = body + len;
    }
  };
  // DggContainer 다음에 dgglbl(1바이트) + DgContainer 가 이어짐
  let o = fc;
  while (o + 8 <= end) {
    const type = u16(tbl, o + 2); const len = u32(tbl, o + 4);
    if (type >= 0xf000 && type <= 0xf122 && o + 8 + len <= end) { walk(o, o + 8 + len, 0); o += 8 + len; } else o += 1;
  }
  return art;
}

/** PICF (Data 스트림) → { ext, data, w, h } */
function picAt(data, fc, wd) {
  if (!data || fc + 0x44 > data.length) return null;
  const lcb = u32(data, fc); const cbHeader = u16(data, fc + 4);
  if (cbHeader !== 0x44 || lcb < 0x44 || fc + lcb > data.length + 8) return null;
  const mm = u16(data, fc + 6);
  const dxaGoal = i16(data, fc + 28); const dyaGoal = i16(data, fc + 30); const mx = u16(data, fc + 32) || 1000; const my = u16(data, fc + 34) || 1000;
  let o = fc + 0x44;
  if (mm === 0x66) o += 1 + data[o];
  const blip = findBlip(data, o, fc + lcb, wd);
  if (!blip) return null;
  return { ...blip, w: Math.max(4, dxaGoal * mx / 1000 * TW), h: Math.max(4, dyaGoal * my / 1000 * TW) };
}

// ───────────── 목록 ─────────────
function readLists(tbl, fcLst, lcbLst, fcLfo, lcbLfo) {
  const lists = new Map(); // lsid → levels
  if (lcbLst > 2) {
    const n = i16(tbl, fcLst);
    let o = fcLst + 2;
    const lstf = [];
    for (let k = 0; k < n; k++) { lstf.push({ lsid: i32(tbl, o), simple: !!(tbl[o + 26] & 1) }); o += 28; }
    for (const L of lstf) {
      const levels = [];
      for (let lv = 0; lv < (L.simple ? 1 : 9) && o + 28 <= tbl.length; lv++) {
        const start = i32(tbl, o); const nfc = tbl[o + 4]; const jc = tbl[o + 5] & 3; const follow = tbl[o + 15];
        const cbChpx = tbl[o + 24]; const cbPapx = tbl[o + 25];
        let q = o + 28;
        const papx = papFrom(tbl, sprms(tbl, q, q + cbPapx)); q += cbPapx + cbChpx;
        const cch = u16(tbl, q); q += 2;
        let text = '';
        for (let c = 0; c < cch; c++) { const ch = u16(tbl, q + c * 2); text += ch < 9 ? `%${ch + 1}` : String.fromCharCode(ch); }
        q += cch * 2;
        o = q;
        const fmt = NFC[nfc] ?? 'decimal';
        levels.push({ fmt, text: fmt === 'bullet' ? (text.replace(/%\d/g, '') || '●').replace(/[]/g, '●').replace(//g, '■').replace(//g, '➢').replace(/[-]/g, '•') : text, start: start || 1, ind: { l: papx.l ?? 24 * (lv + 1), hanging: papx.first < 0 ? -papx.first : 18 }, suff: follow === 1 ? 'space' : follow === 2 ? 'nothing' : 'tab', ...(jc ? { align: ['left', 'center', 'right'][jc] } : {}) });
      }
      while (levels.length < 9) levels.push({ ...levels[levels.length - 1] ?? { fmt: 'bullet', text: '●', start: 1, ind: { l: 24, hanging: 18 } }, ind: { l: 24 * (levels.length + 1), hanging: 18 } });
      lists.set(L.lsid, levels);
    }
  }
  const lfo = [];
  if (lcbLfo > 4) { const n = u32(tbl, fcLfo); for (let k = 0; k < n && k < 4096; k++) lfo.push(i32(tbl, fcLfo + 4 + k * 16)); }
  return { lists, lfo };
}

// ───────────── 구역 ─────────────
function secFrom(wd, fcSepx) {
  const sec = newSection();
  if (fcSepx === 0xffffffff || fcSepx + 2 > wd.length) return { sec, bkc: 2 };
  const cb = i16(wd, fcSepx);
  let bkc = 2; let land = false;
  const m = { ...sec.margin };
  let w = sec.page.w; let h = sec.page.h; let cols = 1; let space = 48;
  for (const { op, o } of sprms(wd, fcSepx + 2, fcSepx + 2 + cb)) {
    switch (op) {
      case 0x3009: bkc = wd[o]; break;
      case 0xb01f: w = u16(wd, o) * TW; break;
      case 0xb020: h = u16(wd, o) * TW; break;
      case 0xb021: m.l = u16(wd, o) * TW; break;
      case 0xb022: m.r = u16(wd, o) * TW; break;
      case 0x9023: m.t = Math.abs(i16(wd, o)) * TW; break;
      case 0x9024: m.b = Math.abs(i16(wd, o)) * TW; break;
      case 0xb017: m.header = u16(wd, o) * TW; break;
      case 0xb018: m.footer = u16(wd, o) * TW; break;
      case 0xb025: m.gutter = u16(wd, o) * TW; break;
      case 0x500b: cols = u16(wd, o) + 1; break;
      case 0x900c: space = u16(wd, o) * TW; break;
      case 0x301d: land = wd[o] === 2; break;
      case 0x300a: sec.titlePg = !!wd[o]; break;
      default: break;
    }
  }
  sec.page = { w, h, orient: land || w > h ? 'landscape' : 'portrait' };
  sec.margin = m;
  if (cols > 1) sec.cols = { n: cols, space, sep: false };
  return { sec, bkc };
}

/** 암호 문서 풀기 (MS-OFFCRYPTO: RC4 · RC4 CryptoAPI). 512바이트 블록마다 키를 새로 만듦 */
function rc4Key(hdr, password) {
  const major = u16(hdr, 0); const minor = u16(hdr, 2);
  const pw = new Uint8Array(password.length * 2); for (let i = 0; i < password.length; i++) { pw[2 * i] = password.charCodeAt(i) & 255; pw[2 * i + 1] = password.charCodeAt(i) >> 8; }
  const cat = (...a) => { const n = a.reduce((x, y) => x + y.length, 0); const o = new Uint8Array(n); let k = 0; for (const x of a) { o.set(x, k); k += x.length; } return o; };
  const le = (n) => new Uint8Array([n & 255, (n >> 8) & 255, (n >> 16) & 255, (n >>> 24) & 255]);
  if (major === 1 && minor === 1) {
    const salt = hdr.subarray(4, 20); const ev = hdr.subarray(20, 36); const evh = hdr.subarray(36, 52);
    const h0 = md5(pw).subarray(0, 5);
    const buf = new Uint8Array(16 * 21); for (let i = 0; i < 16; i++) { buf.set(h0, i * 21); buf.set(salt, i * 21 + 5); }
    const t = md5(buf).subarray(0, 5);
    const key = (b) => md5(cat(t, le(b)));
    const r = rc4(key(0)); const v = r(ev); const vh = r(evh);
    return md5(v).every((x, i) => x === vh[i]) ? key : null;
  }
  if (minor === 2 && major >= 2 && major <= 4) {
    const hsz = u32(hdr, 8); const h = 12; const keyBits = u32(hdr, h + 16) || 40;
    const vo = h + hsz; const saltSize = u32(hdr, vo); const salt = hdr.subarray(vo + 4, vo + 4 + saltSize);
    const ev = hdr.subarray(vo + 4 + saltSize, vo + 20 + saltSize); const hs = u32(hdr, vo + 20 + saltSize); const evh = hdr.subarray(vo + 24 + saltSize, vo + 24 + saltSize + hs);
    const h0 = sha1(cat(salt, pw));
    const key = (b) => { const k = sha1(cat(h0, le(b))).subarray(0, keyBits / 8); if (keyBits === 40) { const p16 = new Uint8Array(16); p16.set(k); return p16; } return k; };
    const r = rc4(key(0)); const v = r(ev); const vh = r(evh);
    return sha1(v).every((x, i) => x === vh[i]) ? key : null;
  }
  throw new Error(`지원하지 않는 Word 암호 방식입니다 (${major}.${minor}).`);
}
function rc4Stream(data, key, keep) {
  const out = new Uint8Array(data.length);
  for (let o = 0, b = 0; o < data.length; o += 512, b++) out.set(rc4(key(b))(data.subarray(o, o + 512)), o);
  out.set(data.subarray(0, keep), 0);
  return out;
}

export function readDoc97(bytes, opts = {}) {
  const cfb = openCfb(bytes);
  let wd = cfb.stream('WordDocument');
  if (!wd || wd.length < 0x200) throw new Error('Word 97-2003 문서(.doc)의 본문이 없습니다.');
  const ident = u16(wd, 0);
  if (ident !== 0xa5ec && ident !== 0xa5dc) throw new Error('Word 97 이전 형식(.doc, Word 6/95)은 지원하지 않습니다. Word 에서 .docx 로 저장한 뒤 열어 주세요.');
  if (u16(wd, 2) < 0xc1) throw new Error('Word 6.0/95 형식(.doc)은 지원하지 않습니다. Word 에서 .docx 로 저장한 뒤 열어 주세요.');
  const flags = u16(wd, 0x0a);
  let tbl = cfb.stream(flags & 0x0200 ? '1Table' : '0Table');
  if (!tbl) throw new Error('Word 문서의 표 스트림(0Table/1Table)이 없습니다.');
  let dataStream = cfb.stream('Data');
  const warnings = [];
  if (flags & 0x0100) {
    if (flags & 0x8000) throw new Error('XOR 방식으로 잠긴 오래된 Word 문서입니다. Word 에서 암호를 풀고 저장한 뒤 열어 주세요.');
    if (!opts.password) { const e = new Error('암호로 보호된 Word 문서입니다. 열기 암호를 입력하세요.'); e.code = 'PASSWORD'; throw e; }
    const lKey = u32(wd, 0x0e);
    const key = rc4Key(tbl, opts.password);
    if (!key) { const e = new Error('암호가 맞지 않습니다.'); e.code = 'PASSWORD'; e.wrong = true; throw e; }
    wd = rc4Stream(wd, key, 68);
    tbl = rc4Stream(tbl, key, lKey);
    if (dataStream) dataStream = rc4Stream(dataStream, key, 0);
    warnings.push('암호를 풀어 열었습니다. 저장하면 암호가 없는 문서가 됩니다.');
  }
  // FIB: csw · fibRgW · cslw · fibRgLw · cbRgFcLcb · fibRgFcLcb
  let o = 32;
  const csw = u16(wd, o); o += 2 + csw * 2;
  const cslw = u16(wd, o); o += 2;
  const lw = o; o += cslw * 4;
  const ccp = (i) => u32(wd, lw + i * 4);
  const ccpText = ccp(3); const ccpFtn = ccp(4); const ccpHdd = ccp(5); const ccpAtn = ccp(7); const ccpEdn = ccp(8); const ccpTxbx = ccp(9); const ccpHdrTxbx = ccp(10);
  const cbRgFcLcb = u16(wd, o); o += 2;
  const fcl = o;
  const pair = (i) => (i < cbRgFcLcb ? [u32(wd, fcl + i * 8), u32(wd, fcl + i * 8 + 4)] : [0, 0]);
  const [fcClx, lcbClx] = pair(33);
  const [fcBteChpx, lcbBteChpx] = pair(12);
  const [fcBtePapx, lcbBtePapx] = pair(13);
  const [fcFfn, lcbFfn] = pair(15);
  // 이야기 시작 CP
  const ST = { main: 0, ftn: ccpText, hdd: ccpText + ccpFtn, atn: ccpText + ccpFtn + ccpHdd + ccp(6) };
  ST.edn = ST.atn + ccpAtn; ST.txbx = ST.edn + ccpEdn; ST.hdrTxbx = ST.txbx + ccpTxbx;
  const total = ST.hdrTxbx + ccpHdrTxbx;
  // 글꼴 표
  const fonts = [];
  if (lcbFfn) {
    const n = u16(tbl, fcFfn);
    let p = fcFfn + 4;
    for (let k = 0; k < n && p < fcFfn + lcbFfn; k++) {
      const cb = tbl[p] + 1;
      let name = '';
      for (let q = p + 40; q + 1 < p + cb; q += 2) { const c = u16(tbl, q); if (!c) break; name += String.fromCharCode(c); }
      fonts.push(name);
      p += cb;
    }
  }
  // 조각 표 → 글자마다 fc
  let p = fcClx;
  const end = fcClx + lcbClx;
  while (p < end && tbl[p] === 1) p += 3 + u16(tbl, p + 1); // Prc 건너뛰기
  if (tbl[p] !== 2) throw new Error('Word 문서의 조각 표(Clx)를 읽을 수 없습니다.');
  const lcb = u32(tbl, p + 1);
  const plc = p + 5;
  const npcd = (lcb - 4) / 12;
  const chars = [];
  const fcs = [];
  for (let k = 0; k < npcd && chars.length < total + 1; k++) {
    const cp0 = u32(tbl, plc + k * 4); const cp1 = u32(tbl, plc + (k + 1) * 4);
    const pcd = plc + (npcd + 1) * 4 + k * 8;
    let fc = u32(tbl, pcd + 2);
    const compressed = !!(fc & 0x40000000);
    fc &= 0x3fffffff;
    if (compressed) fc /= 2;
    for (let cp = cp0; cp < cp1 && chars.length < total + 1; cp++) {
      if (compressed) { const c = wd[fc]; chars.push(c >= 0x80 && c <= 0x9f ? CP1252[c - 0x80] : String.fromCharCode(c)); fcs.push(fc); fc++; }
      else { chars.push(String.fromCharCode(u16(wd, fc))); fcs.push(fc); fc += 2; }
    }
  }
  // 서식 실행 (FKP)
  const runsOf = (fcB, lcbB, kind) => {
    const out = []; // [fcStart, fcEnd, props]
    if (!lcbB) return out;
    const n = (lcbB - 4) / 8;
    for (let k = 0; k < n; k++) {
      const pn = u32(tbl, fcB + (n + 1) * 4 + k * 4) & 0x3fffff;
      const page = pn * 512;
      if (page + 512 > wd.length) continue;
      const cnt = wd[page + 511];
      for (let j = 0; j < cnt; j++) {
        const f0 = u32(wd, page + j * 4); const f1 = u32(wd, page + (j + 1) * 4);
        if (kind === 'chp') {
          const off = wd[page + (cnt + 1) * 4 + j] * 2;
          if (!off) { out.push([f0, f1, {}]); continue; }
          const cb = wd[page + off];
          out.push([f0, f1, chpFrom(wd, sprms(wd, page + off + 1, page + off + 1 + cb), fonts)]);
        } else {
          const off = wd[page + (cnt + 1) * 4 + j * 13] * 2;
          if (!off) { out.push([f0, f1, { istd: 0 }]); continue; }
          let cb = wd[page + off];
          let q = page + off + 1;
          if (!cb) { cb = wd[q]; q++; }
          const len = cb * 2;
          const istd = u16(wd, q);
          out.push([f0, f1, { istd, ...papFrom(wd, sprms(wd, q + 2, q + len - (cb === wd[page + off] ? 1 : 0))) }]);
        }
      }
    }
    out.sort((a, b) => a[0] - b[0]);
    return out;
  };
  const chpRuns = runsOf(fcBteChpx, lcbBteChpx, 'chp');
  const papRuns = runsOf(fcBtePapx, lcbBtePapx, 'pap');
  const find = (list, fc) => { let lo = 0; let hi = list.length - 1; while (lo <= hi) { const m = (lo + hi) >> 1; if (list[m][1] <= fc) lo = m + 1; else if (list[m][0] > fc) hi = m - 1; else return list[m][2]; } return null; };
  // PLC 도우미: (fc, lcb, 자료 크기) → { cps, data(i) }
  const plcf = (i, cbData) => {
    const [f, l] = pair(i);
    if (!l) return { cps: [], at: () => 0 };
    const n = Math.floor((l - 4) / (4 + cbData));
    const cps = []; for (let k = 0; k <= n; k++) cps.push(u32(tbl, f + k * 4));
    return { cps, n, at: (k) => f + (n + 1) * 4 + k * cbData };
  };
  const doc = newDoc();
  doc.media = {}; doc.lists = {};
  // 목록
  const [fcLst, lcbLst] = pair(73); const [fcLfo, lcbLfo] = pair(74);
  let L = { lists: new Map(), lfo: [] };
  try { L = readLists(tbl, fcLst, lcbLst, fcLfo, lcbLfo); } catch { warnings.push('목록 정의 일부를 읽지 못했습니다.'); }
  // 그림 · 도형
  const [fcDgg, lcbDgg] = pair(50);
  let art = { blips: [], shapes: new Map() };
  try { art = readArt(tbl, fcDgg, lcbDgg, wd); } catch { /* 그림 없이 */ }
  const spa = plcf(40, 26);
  const spaAt = new Map(); for (let k = 0; k < (spa.n ?? 0); k++) spaAt.set(spa.cps[k], spa.at(k));
  const addMedia = (blip) => { if (!blip?.data?.length) return null; const mid = uid('m'); doc.media[mid] = mediaDataUrl(blip.data, `x.${blip.ext === 'pict' ? 'bin' : blip.ext}`); return mid; };
  // 각주 · 미주 참조
  const ftnRef = plcf(2, 2); const ftnTxt = plcf(3, 0); const ednRef = plcf(46, 2); const ednTxt = plcf(47, 0);
  const noteAt = new Map();
  ftnRef.cps.slice(0, ftnRef.n).forEach((cp, k) => noteAt.set(cp, { kind: 'foot', k }));
  ednRef.cps.slice(0, ednRef.n).forEach((cp, k) => noteAt.set(cp, { kind: 'end', k }));
  // 글상자 이야기
  const txbx = plcf(56, 22);

  // ───── 이야기 범위 [cp0, cp1) → 블록 ─────
  let depthGuard = 0;
  const build = (cp0, cp1, story) => {
    if (++depthGuard > 50) return [];
    const blocks = [];
    let runs = [];
    let buf = ''; let bufProps = null;
    const fld = [];
    let row = null; let table = null; let cellBlocks = []; let cellX = null;
    const flushBuf = () => { if (buf) { runs.push({ ...(bufProps ?? {}), text: buf }); buf = ''; } };
    const makePara = (pap) => {
      flushBuf();
      const p = para('');
      p.runs = normalizeRuns(runs);
      runs = [];
      if (pap) {
        const ind = {};
        if (pap.l) ind.l = pap.l;
        if (pap.r) ind.r = pap.r;
        if (pap.first > 0) ind.first = pap.first; else if (pap.first < 0) ind.hanging = -pap.first;
        if (Object.keys(ind).length) p.pPr.ind = ind;
        if (pap.align) p.pPr.align = pap.align;
        if (pap.before != null || pap.after != null || pap.line) p.pPr.spc = { ...(pap.before != null ? { before: pap.before } : {}), ...(pap.after != null ? { after: pap.after } : {}), ...(pap.line ?? {}) };
        if (pap.keepNext) p.pPr.keepNext = true;
        if (pap.keepLines) p.pPr.keepLines = true;
        if (pap.pageBreakBefore) p.pPr.pageBreakBefore = true;
        if (pap.istd >= 1 && pap.istd <= 9) p.style = `Heading${pap.istd}`;
        else if (pap.outline != null && pap.outline < 9) p.pPr.outline = pap.outline;
        if (pap.ilfo > 0 && pap.ilfo < 2047) {
          const lsid = L.lfo[pap.ilfo - 1];
          const key = `doc${lsid ?? pap.ilfo}`;
          if (!doc.lists[key]) doc.lists[key] = { levels: L.lists.get(lsid) ?? Array.from({ length: 9 }, (_, i) => ({ fmt: 'bullet', text: ['●', '○', '■'][i % 3], start: 1, ind: { l: 24 * (i + 1), hanging: 18 } })) };
          p.pPr.num = { id: key, lvl: Math.min(8, pap.ilvl ?? 0) };
          if (p.style === 'Normal') p.style = 'ListParagraph';
        }
      }
      if (story === 'hdr' && p.style === 'Normal') p.style = 'Header';
      return p;
    };
    const closeTable = () => {
      if (!table) return;
      const cols = Math.max(...table.rows.map((r) => r.cells.length));
      const x = table._x;
      table.grid = x && x.length === cols + 1 ? x.slice(1).map((v, k) => Math.max(8, (v - x[k]) * TW)) : Array.from({ length: cols }, () => 600 / cols);
      for (const r of table.rows) while (r.cells.length < cols) r.cells.push({ blocks: [para('')] });
      delete table._x;
      blocks.push(table); table = null;
    };
    for (let cp = cp0; cp < cp1 && cp < chars.length; cp++) {
      const ch = chars[cp];
      const fc = fcs[cp];
      const code = ch.charCodeAt(0);
      if (code === 0x13) { flushBuf(); fld.push({ phase: 'code', code: '' }); continue; }
      if (code === 0x14) { const f = fld[fld.length - 1]; if (f) { f.phase = 'result'; const m = /HYPERLINK\s+(?:\\l\s+)?"([^"]*)"/i.exec(f.code); if (m) f.link = /\\l/.test(f.code) ? { anchor: m[1] } : { url: m[1] }; } flushBuf(); continue; }
      if (code === 0x15) { flushBuf(); fld.pop(); continue; }
      if (fld.length && fld[fld.length - 1].phase === 'code') { fld[fld.length - 1].code += ch; continue; }
      const link = fld.find((f) => f.link)?.link;
      if (code === 0x0d || code === 0x07) {
        const pap = find(papRuns, fc) ?? {};
        if (code === 0x07) {
          if (pap.ttp || (buf === '' && runs.length === 0 && row && row.justClosed)) {
            // 행 끝
            flushBuf(); runs = [];
            if (!table) table = { type: 'tbl', id: uid('t'), style: 'TableGrid', tblPr: { look: { firstRow: false, firstCol: false, bandRows: false } }, grid: [], rows: [] };
            if (pap.cellX && !table._x) table._x = pap.cellX;
            table.rows.push({ cells: row?.cells ?? [] });
            row = null;
            continue;
          }
          const pp = makePara(pap);
          cellBlocks.push(pp);
          if (!row) row = { cells: [] };
          row.cells.push({ blocks: cellBlocks });
          cellBlocks = [];
          row.justClosed = true;
          continue;
        }
        const pp = makePara(pap);
        if (pap.inTable) { cellBlocks.push(pp); if (row) row.justClosed = false; continue; }
        closeTable();
        blocks.push(pp);
        if (story === 'main') { const s = secEnd.get(cp + 1); if (s) pp.secBreak = s; }
        continue;
      }
      if (row) row.justClosed = false;
      if (code === 0x0c) { flushBuf(); if (!(story === 'main' && secEnd.get(cp + 1))) runs.push({ obj: 'br', kind: 'page' }); continue; }
      if (code === 0x0e) { flushBuf(); runs.push({ obj: 'br', kind: 'column' }); continue; }
      if (code === 0x0b) { flushBuf(); runs.push({ obj: 'br' }); continue; }
      if (code === 0x09) { flushBuf(); runs.push({ ...clean(find(chpRuns, fc)), obj: 'tab' }); continue; }
      if (code === 0x02 && story === 'main' && noteAt.has(cp)) {
        // 각주 · 미주 참조
        flushBuf();
        const n = noteAt.get(cp);
        const txt = n.kind === 'foot' ? ftnTxt : ednTxt;
        const base = n.kind === 'foot' ? ST.ftn : ST.edn;
        const id = uid('n');
        const a = txt.cps[n.k]; const b = txt.cps[n.k + 1];
        const nb = a != null && b != null ? build(base + a, base + b, 'note') : [];
        // 각주 글 앞의 번호 표시(0x02) 는 build 에서 빠짐
        doc.notes[n.kind][id] = nb.length ? nb : [para('')];
        runs.push({ obj: 'note', id, kind: n.kind });
        continue;
      }
      if (code === 0x01) {
        // 글자처럼 그림 (PICF)
        const props = find(chpRuns, fc) ?? {};
        if (props._pic != null && !props._ole) {
          const pic = picAt(dataStream, props._pic, wd);
          const mid = pic && addMedia(pic);
          if (mid) { flushBuf(); runs.push({ obj: 'img', id: uid('i'), media: mid, w: Math.round(pic.w), h: Math.round(pic.h), alt: '' }); }
        }
        continue;
      }
      if (code === 0x08) {
        // 떠 있는 개체 (FSPA)
        const q = story === 'main' ? spaAt.get(cp) : null;
        if (q != null) {
          const spid = u32(tbl, q); const x0 = i32(tbl, q + 4); const y0 = i32(tbl, q + 8); const x1 = i32(tbl, q + 12); const y1 = i32(tbl, q + 16); const fl = u16(tbl, q + 20);
          const shp = art.shapes.get(spid);
          const bx = (fl >> 1) & 3; const by = (fl >> 3) & 3; const wr = (fl >> 5) & 15; const below = !!(fl & (1 << 14));
          const float = { wrap: wr === 3 ? (below ? 'behind' : 'front') : wr === 1 ? 'topBottom' : wr === 4 || wr === 5 ? 'tight' : 'square', relH: bx === 1 ? 'page' : bx === 2 ? 'column' : 'margin', relV: by === 1 ? 'page' : by === 2 ? 'paragraph' : 'margin', x: x0 * TW, y: y0 * TW };
          if (float.wrap === 'square' || float.wrap === 'tight') float.side = 'both';
          const w = Math.max(4, (x1 - x0) * TW); const h = Math.max(4, (y1 - y0) * TW);
          if (shp?.pib && art.blips[shp.pib - 1]) {
            const mid = addMedia(art.blips[shp.pib - 1]);
            if (mid) { flushBuf(); runs.push({ obj: 'img', id: uid('i'), media: mid, w, h, alt: '', float }); }
          } else if (shp?.txid != null && txbx.n) {
            const k = (shp.txid >>> 16) - 1;
            if (k >= 0 && k < txbx.n) {
              const tb = build(ST.txbx + txbx.cps[k], ST.txbx + txbx.cps[k + 1], 'txbx');
              flushBuf();
              runs.push({ obj: 'shape', id: uid('s'), w, h, prst: 'rect', line: { w: 0.75, color: '#000000' }, blocks: tb.length ? tb : [para('')], float });
            }
          }
        }
        continue;
      }
      if (code === 0x02 || code === 0x05 || code === 0x03 || code === 0x04) continue; // 번호 표시 · 메모 표시 · 구분선
      if (code === 0x1e) { buf += '‑'; continue; }
      if (code === 0x1f) continue;
      const props = { ...clean(find(chpRuns, fc)), ...(link ? { link } : {}) };
      if (bufProps && JSON.stringify(props) !== JSON.stringify(bufProps)) flushBuf();
      bufProps = props;
      buf += ch;
    }
    flushBuf();
    if (runs.length) blocks.push(makePara(null));
    closeTable();
    depthGuard--;
    return blocks;
  };
  const clean = (r) => { if (!r) return {}; const o2 = {}; for (const [k, v] of Object.entries(r)) if (k[0] !== '_') o2[k] = v; return o2; };

  // 구역
  const sed = plcf(6, 12);
  const secs = [];
  for (let k = 0; k < (sed.n ?? 0); k++) secs.push({ cpEnd: sed.cps[k + 1], ...secFrom(wd, u32(tbl, sed.at(k) + 2)) });
  if (!secs.length) secs.push({ cpEnd: ccpText, sec: newSection(), bkc: 2 });
  const secEnd = new Map();
  const BKC = ['continuous', 'nextColumn', 'nextPage', 'evenPage', 'oddPage'];
  for (let k = 0; k < secs.length - 1; k++) secEnd.set(secs[k].cpEnd, { kind: BKC[secs[k + 1].bkc] ?? 'nextPage', sec: secs[k].sec });
  // 머리글 · 바닥글: 앞 6개는 각주 구분선, 그다음 구역마다 6개 (짝수 머리글 · 홀수 머리글 · 짝수 바닥글 · 홀수 바닥글 · 첫 머리글 · 첫 바닥글)
  const hdd = plcf(11, 0);
  const story = (i) => { const a = hdd.cps[i]; const b = hdd.cps[i + 1]; if (a == null || b == null || b - a <= 1) return null; const bl = build(ST.hdd + a, ST.hdd + b, 'hdr'); while (bl.length > 1 && bl[bl.length - 1].type === 'p' && !bl[bl.length - 1].runs.length) bl.pop(); return bl.length ? bl : null; };
  secs.forEach((s, k) => {
    const base = 6 + k * 6;
    if (base + 6 >= hdd.cps.length) { if (k) { s.sec.headers = secs[k - 1].sec.headers; s.sec.footers = secs[k - 1].sec.footers; } return; }
    const set = (obj, key, i) => { const v = story(i); if (v) obj[key] = v; else if (k && secs[k - 1].sec[obj === s.sec.headers ? 'headers' : 'footers'][key]) obj[key] = secs[k - 1].sec[obj === s.sec.headers ? 'headers' : 'footers'][key]; };
    set(s.sec.headers, 'even', base); set(s.sec.headers, 'default', base + 1); set(s.sec.footers, 'even', base + 2); set(s.sec.footers, 'default', base + 3); set(s.sec.headers, 'first', base + 4); set(s.sec.footers, 'first', base + 5);
    if (s.sec.headers.even || s.sec.footers.even) doc.settings.evenOdd = true;
  });
  const blocks = build(0, ccpText, 'main');
  doc.body = blocks.length ? blocks : [para('')];
  doc.section = secs[secs.length - 1].sec;
  return { doc, warnings };
}
