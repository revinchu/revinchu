// RTF 읽기 (DOM 없음): 글꼴/색 표 · 굵게/기울임/밑줄/취소선 · 크기 · 색 · 정렬 · 들여쓰기 · 표(\cell \row) · 그림(\pngblip \jpegblip) · 한국어(\ansicpg949, \'hh, \uN)
import { newDoc, para, uid, normalizeRuns } from './model.js';

const CP = { 949: 'euc-kr', 1252: 'windows-1252', 932: 'shift_jis', 936: 'gbk', 950: 'big5', 1250: 'windows-1250', 1251: 'windows-1251', 65001: 'utf-8' };
const SKIP = new Set(['fonttbl', 'colortbl', 'stylesheet', 'info', 'header', 'footer', 'headerl', 'headerr', 'headerf', 'footerl', 'footerr', 'footerf', 'footnote', 'listtable', 'listoverridetable', 'rsidtbl', 'generator', 'xmlnstbl', 'themedata', 'colorschememapping', 'latentstyles', 'datastore', 'pgdsctbl', 'revtbl', 'object', 'fldinst', 'bkmkstart', 'bkmkend', 'field-inst', 'nonshppict', 'mmathPr', 'fchars', 'lchars', 'pnseclvl', 'xe', 'tc']);

export function readRtf(text) {
  const doc = newDoc();
  doc.media = {};
  const fonts = {};
  const colors = [null];
  let cp = 'windows-1252';
  let i = 0;
  const n = text.length;
  const stack = [];
  let st = { b: false, i: false, u: false, s: false, sz: 24, color: null, font: null, skip: 0, uc: 1, va: null };
  let pp = { align: null, l: 0, r: 0, first: 0, intbl: false, before: 0, after: 0 };
  const blocks = [];
  let runs = [];
  let buf = '';
  let bytes = [];
  let table = null; let row = []; let cellBlocks = [];
  let pict = null;
  let fldUrl = null;
  const decodeBytes = () => { if (bytes.length) { try { buf += new TextDecoder(cp).decode(new Uint8Array(bytes)); } catch { buf += String.fromCharCode(...bytes); } bytes = []; } };
  const flush = () => {
    decodeBytes();
    if (buf) {
      const r = { text: buf };
      if (st.b) r.b = true; if (st.i) r.i = true; if (st.u) r.u = true; if (st.s) r.s = true;
      if (st.sz && st.sz !== 12) r.sz = st.sz / 2 === 10 ? undefined : st.sz / 2;
      if (st.sz) r.sz = st.sz / 2;
      if (st.color) r.color = st.color;
      if (st.font) { r.font = st.font; r.fontEa = st.font; }
      if (st.va) r.va = st.va;
      if (fldUrl) r.link = { url: fldUrl };
      for (const k of Object.keys(r)) if (r[k] === undefined) delete r[k];
      runs.push(r);
      buf = '';
    }
  };
  const endPara = () => {
    flush();
    const p = para('');
    p.runs = normalizeRuns(runs);
    runs = [];
    if (pp.align) p.pPr.align = pp.align;
    const ind = {};
    if (pp.l) ind.l = pp.l / 15; if (pp.r) ind.r = pp.r / 15;
    if (pp.first > 0) ind.first = pp.first / 15; else if (pp.first < 0) ind.hanging = -pp.first / 15;
    if (Object.keys(ind).length) p.pPr.ind = ind;
    if (pp.before || pp.after) p.pPr.spc = { before: pp.before / 15, after: pp.after / 15 };
    return p;
  };
  const closeTable = () => { if (table) { const cols = Math.max(...table.rows.map((r) => r.cells.length)); table.grid = Array.from({ length: cols }, () => 600 / cols); for (const r of table.rows) while (r.cells.length < cols) r.cells.push({ blocks: [para('')] }); blocks.push(table); table = null; } };
  while (i < n) {
    const c = text[i];
    if (c === '{') { flush(); stack.push({ ...st }); i++; continue; }
    if (c === '}') {
      flush();
      if (pict && st.skip === 0 && stack.length === pict.depth) { /* 아래에서 처리 */ }
      if (pict && stack.length === pict.depth) {
        const hex = pict.hex.replace(/[^0-9a-f]/gi, '');
        if (hex.length > 16) {
          const b = new Uint8Array(hex.length / 2);
          for (let k = 0; k < b.length; k++) b[k] = parseInt(hex.substr(k * 2, 2), 16);
          let s = '';
          for (let k = 0; k < b.length; k += 0x8000) s += String.fromCharCode.apply(null, b.subarray(k, k + 0x8000));
          const id = uid('m');
          doc.media[id] = `data:${pict.type};base64,${typeof btoa === 'function' ? btoa(s) : Buffer.from(b).toString('base64')}`;
          const w = pict.w || pict.pw || 300; const h = pict.h || pict.ph || 200;
          runs.push({ obj: 'img', id: uid('i'), media: id, w: Math.min(600, w), h: (h * Math.min(600, w)) / w });
        }
        pict = null;
      }
      const prevFld = st.fldrslt;
      st = stack.pop() ?? st;
      if (prevFld && !st.fldrslt) fldUrl = null;
      i++;
      continue;
    }
    if (c === '\\') {
      const nx = text[i + 1];
      if (nx === '\\' || nx === '{' || nx === '}') { if (!st.skip) buf += nx; i += 2; continue; }
      if (nx === "'") { if (!st.skip) { bytes.push(parseInt(text.substr(i + 2, 2), 16)); } i += 4; continue; }
      if (nx === '*') { st.skip = 1; i += 2; continue; }
      if (nx === '~') { if (!st.skip) buf += ' '; i += 2; continue; }
      if (nx === '-' || nx === '_') { i += 2; continue; }
      if (nx === '\n' || nx === '\r') { if (!st.skip) { decodeBytes(); const p = endPara(); (pp.intbl ? cellBlocks : (closeTable(), blocks)).push(p); } i += 2; continue; }
      const m = /^([a-zA-Z]+)(-?\d+)? ?/.exec(text.slice(i + 1, i + 40));
      if (!m) { i++; continue; }
      i += 1 + m[0].length;
      const w = m[1]; const arg = m[2] != null ? Number(m[2]) : null;
      if (pict) {
        if (w === 'pngblip') pict.type = 'image/png'; else if (w === 'jpegblip') pict.type = 'image/jpeg'; else if (w === 'emfblip' || w === 'wmetafile') pict.type = null;
        else if (w === 'picwgoal') pict.w = arg / 15; else if (w === 'pichgoal') pict.h = arg / 15; else if (w === 'picw') pict.pw = arg; else if (w === 'pich') pict.ph = arg;
        continue;
      }
      if (SKIP.has(w)) {
        if (w === 'fonttbl') { // 글꼴 표: {\f0 … 이름;}
          let depth = 1; const s0 = i; while (i < n && depth > 0) { if (text[i] === '{') depth++; else if (text[i] === '}') depth--; i++; }
          const hexName = (x) => x.replace(/(?:\\'[0-9a-f]{2})+/gi, (run) => { const bs = run.split("\\'").filter(Boolean).map((h) => parseInt(h, 16)); try { return new TextDecoder(cp).decode(new Uint8Array(bs)); } catch { return String.fromCharCode(...bs); } });
          const body = hexName(text.slice(s0, i - 1));
          for (const fm of body.matchAll(/\\f(\d+)[^;{}]*?(?:\\fcharset(\d+))?[^;{} ]* ([^;{}\\]+);/g)) fonts[fm[1]] = fm[3].trim();
          for (const fm of body.matchAll(/\{\\f(\d+)[^}]*?\s([^\\;{}]+);\}/g)) fonts[fm[1]] = fonts[fm[1]] ?? fm[2].trim();
          st = stack.pop() ?? st;
          continue;
        }
        if (w === 'colortbl') {
          let depth = 1; const s0 = i; while (i < n && depth > 0) { if (text[i] === '{') depth++; else if (text[i] === '}') depth--; i++; }
          const body = text.slice(s0, i - 1);
          for (const e of body.split(';').slice(1)) { const r = /\\red(\d+)/.exec(e); const g = /\\green(\d+)/.exec(e); const b = /\\blue(\d+)/.exec(e); if (r && g && b) colors.push(`#${[r[1], g[1], b[1]].map((x) => Number(x).toString(16).padStart(2, '0')).join('').toUpperCase()}`); }
          st = stack.pop() ?? st;
          continue;
        }
        if (w === 'fldinst') { // HYPERLINK "url"
          let depth = 1; const s0 = i; while (i < n && depth > 0) { if (text[i] === '{') depth++; else if (text[i] === '}') depth--; i++; }
          const mm = /HYPERLINK\s+"([^"]+)"/.exec(text.slice(s0, i));
          if (mm) fldUrl = mm[1];
          st = stack.pop() ?? st;
          continue;
        }
        st.skip = 1;
        continue;
      }
      if (st.skip) continue;
      switch (w) {
        case 'ansicpg': cp = CP[arg] ?? cp; break;
        case 'par': { decodeBytes(); const p = endPara(); if (pp.intbl) cellBlocks.push(p); else { closeTable(); blocks.push(p); } break; }
        case 'line': flush(); runs.push({ obj: 'br' }); break;
        case 'page': flush(); runs.push({ obj: 'br', kind: 'page' }); break;
        case 'tab': flush(); runs.push({ obj: 'tab' }); break;
        case 'cell': { decodeBytes(); if (buf || runs.length || !cellBlocks.length) cellBlocks.push(endPara()); row.push({ blocks: cellBlocks }); cellBlocks = []; break; }
        case 'row': { flush(); if (!table) table = { type: 'tbl', id: uid('t'), style: 'TableGrid', tblPr: {}, grid: [], rows: [] }; table.rows.push({ cells: row }); row = []; break; }
        case 'intbl': pp.intbl = true; break;
        case 'pard': flush(); pp = { align: null, l: 0, r: 0, first: 0, intbl: false, before: 0, after: 0 }; break;
        case 'plain': flush(); st = { ...st, b: false, i: false, u: false, s: false, sz: 24, color: null, font: null, va: null }; break;
        case 'b': flush(); st.b = arg !== 0; break;
        case 'i': flush(); st.i = arg !== 0; break;
        case 'ul': flush(); st.u = arg !== 0; break;
        case 'ulnone': flush(); st.u = false; break;
        case 'strike': flush(); st.s = arg !== 0; break;
        case 'super': flush(); st.va = 'super'; break;
        case 'sub': flush(); st.va = 'sub'; break;
        case 'nosupersub': flush(); st.va = null; break;
        case 'fs': flush(); st.sz = arg; break;
        case 'cf': flush(); st.color = colors[arg] ?? null; break;
        case 'f': flush(); st.font = fonts[arg] ?? null; break;
        case 'qc': pp.align = 'center'; break;
        case 'qr': pp.align = 'right'; break;
        case 'qj': pp.align = 'justify'; break;
        case 'ql': pp.align = 'left'; break;
        case 'li': pp.l = arg; break;
        case 'ri': pp.r = arg; break;
        case 'fi': pp.first = arg; break;
        case 'sb': pp.before = arg; break;
        case 'sa': pp.after = arg; break;
        case 'uc': st.uc = arg; break;
        case 'u': { decodeBytes(); buf += String.fromCharCode(arg < 0 ? arg + 65536 : arg); // 대체 글자 건너뛰기
          for (let k = 0; k < st.uc && i < n; k++) { if (text[i] === '\\' && text[i + 1] === "'") i += 4; else if (text[i] !== '\\' && text[i] !== '{' && text[i] !== '}') i++; }
          break; }
        case 'pict': flush(); pict = { depth: stack.length, hex: '', type: 'image/png' }; break;
        case 'fldrslt': st.fldrslt = true; break;
        case 'emdash': buf += '—'; break;
        case 'endash': buf += '–'; break;
        case 'bullet': buf += '•'; break;
        case 'lquote': buf += '‘'; break;
        case 'rquote': buf += '’'; break;
        case 'ldblquote': buf += '“'; break;
        case 'rdblquote': buf += '”'; break;
        default: break;
      }
      continue;
    }
    if (c === '\r' || c === '\n') { i++; continue; }
    if (pict) { pict.hex += c; i++; continue; }
    if (!st.skip) { if (bytes.length) decodeBytes(); buf += c; }
    i++;
  }
  flush();
  if (runs.length) blocks.push(endPara());
  closeTable();
  doc.body = blocks.length ? blocks : [para('')];
  return { doc, warnings: [] };
}
