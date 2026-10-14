// Word 97-2003 문서 (.doc, MS-DOC) 읽기 — DOM 없음
// 조각 표(Clx) 로 글자를 모으고, 글자 서식(CHPX FKP) · 문단 서식(PAPX FKP) · 글꼴 표 · 표(셀 표시 \x07) · 필드(하이퍼링크) 를 읽음
import { openCfb } from './cfb.js';
import { newDoc, para, uid, normalizeRuns } from './model.js';

const u16 = (b, o) => (o + 2 <= b.length ? b[o] | (b[o + 1] << 8) : 0);
const i16 = (b, o) => { const v = u16(b, o); return v & 0x8000 ? v - 0x10000 : v; };
const u32 = (b, o) => (o + 4 <= b.length ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) >>> 0 : 0);
const CP1252 = '€\x81‚ƒ„…†‡ˆ‰Š‹Œ\x8DŽ\x8F\x90‘’“”•–—˜™š›œ\x9DžŸ';
const ICO = [null, '#000000', '#0000FF', '#00FFFF', '#00FF00', '#FF00FF', '#FF0000', '#FFFF00', '#FFFFFF', '#000080', '#008080', '#008000', '#800080', '#800000', '#808000', '#808080', '#C0C0C0'];
const HL = [null, 'black', 'blue', 'cyan', 'brightGreen', 'magenta', 'red', 'yellow', 'white', 'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkYellow', 'darkGray', 'lightGray'];

/** sprm 목록 → [{ op, arg(바이트 배열 시작), size }] */
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
      case 0x840f: case 0x845e: p.l = i16(b, o) / 15; break;
      case 0x840e: case 0x845d: p.r = i16(b, o) / 15; break;
      case 0x8411: case 0x8460: p.first = i16(b, o) / 15; break;
      case 0xa413: p.before = u16(b, o) / 15; break;
      case 0xa414: p.after = u16(b, o) / 15; break;
      case 0x6412: { const dya = i16(b, o); const mult = u16(b, o + 2); p.line = mult ? { line: Math.abs(dya) / 240, rule: 'auto' } : { line: Math.abs(dya) / 15, rule: dya < 0 ? 'exact' : 'atLeast' }; break; }
      case 0x2416: p.inTable = !!b[o]; break;
      case 0x2417: p.ttp = !!b[o]; break;
      case 0x2640: p.outline = b[o]; break;
      case 0x260a: p.ilvl = b[o]; break;
      case 0x460b: p.ilfo = u16(b, o); break;
      case 0x2405: p.keepLines = !!b[o]; break;
      case 0x2406: p.keepNext = !!b[o]; break;
      case 0x2407: p.pageBreakBefore = !!b[o]; break;
      default: break;
    }
  }
  return p;
}

export function readDoc97(bytes) {
  const cfb = openCfb(bytes);
  const wd = cfb.stream('WordDocument');
  if (!wd || wd.length < 0x200) throw new Error('Word 97-2003 문서(.doc)의 본문이 없습니다.');
  const ident = u16(wd, 0);
  if (ident !== 0xa5ec && ident !== 0xa5dc) throw new Error('Word 97 이전 형식(.doc, Word 6/95)은 지원하지 않습니다. Word 에서 .docx 로 저장한 뒤 열어 주세요.');
  const flags = u16(wd, 0x0a);
  if (flags & 0x0100) throw new Error('암호로 보호된 Word 문서입니다. Word 에서 암호를 풀고 저장한 뒤 열어 주세요.');
  const tbl = cfb.stream(flags & 0x0200 ? '1Table' : '0Table');
  if (!tbl) throw new Error('Word 문서의 표 스트림(0Table/1Table)이 없습니다.');
  // FIB: csw · fibRgW · cslw · fibRgLw · cbRgFcLcb · fibRgFcLcb
  let o = 32;
  const csw = u16(wd, o); o += 2 + csw * 2;
  const cslw = u16(wd, o); o += 2;
  const lw = o; o += cslw * 4;
  const ccpText = u32(wd, lw + 3 * 4);
  o += 2;
  const fcl = o;
  const pair = (i) => [u32(wd, fcl + i * 8), u32(wd, fcl + i * 8 + 4)];
  const [fcClx, lcbClx] = pair(33);
  const [fcBteChpx, lcbBteChpx] = pair(12);
  const [fcBtePapx, lcbBtePapx] = pair(13);
  const [fcFfn, lcbFfn] = pair(15);
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
  for (let k = 0; k < npcd && chars.length < ccpText; k++) {
    const cp0 = u32(tbl, plc + k * 4); const cp1 = u32(tbl, plc + (k + 1) * 4);
    const pcd = plc + (npcd + 1) * 4 + k * 8;
    let fc = u32(tbl, pcd + 2);
    const compressed = !!(fc & 0x40000000);
    fc &= 0x3fffffff;
    if (compressed) fc /= 2;
    for (let cp = cp0; cp < cp1 && chars.length < ccpText; cp++) {
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
  // 문단 · 표 · 필드
  const doc = newDoc();
  const blocks = [];
  let runs = [];
  let buf = ''; let bufProps = null;
  const fld = []; // { phase: 'code'|'result', code, link }
  let row = null; // 진행 중인 표 행 셀들
  let table = null;
  let cellBlocks = [];
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
      if (pap.ilfo) { const key = `doc${pap.ilfo}`; p.pPr.num = { id: key, lvl: pap.ilvl ?? 0 }; p.style = p.style === 'Normal' ? 'ListParagraph' : p.style; }
    }
    return p;
  };
  const closeTable = () => { if (table) { const cols = Math.max(...table.rows.map((r) => r.cells.length)); const W = 600; table.grid = Array.from({ length: cols }, () => W / cols); for (const r of table.rows) while (r.cells.length < cols) r.cells.push({ blocks: [para('')] }); blocks.push(table); table = null; } };
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i];
    const fc = fcs[i];
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
          if (!table) { closeTable(); table = { type: 'tbl', id: uid('t'), style: 'TableGrid', tblPr: { look: { firstRow: false, firstCol: false, bandRows: false } }, grid: [], rows: [] }; }
          table.rows.push({ cells: row?.cells ?? [] });
          row = null;
          continue;
        }
        const p = makePara(pap);
        cellBlocks.push(p);
        if (!row) row = { cells: [] };
        row.cells.push({ blocks: cellBlocks });
        cellBlocks = [];
        row.justClosed = true;
        continue;
      }
      const p = makePara(pap);
      if (pap.inTable) { cellBlocks.push(p); if (row) row.justClosed = false; continue; }
      closeTable();
      blocks.push(p);
      continue;
    }
    if (row) row.justClosed = false;
    if (code === 0x0c) { flushBuf(); runs.push({ obj: 'br', kind: 'page' }); continue; }
    if (code === 0x0e) { flushBuf(); runs.push({ obj: 'br', kind: 'column' }); continue; }
    if (code === 0x0b) { flushBuf(); runs.push({ obj: 'br' }); continue; }
    if (code === 0x09) { flushBuf(); runs.push({ ...(find(chpRuns, fc) ?? {}), obj: 'tab' }); continue; }
    if (code === 0x01 || code === 0x08 || code === 0x05 || code === 0x02 || code === 0x03 || code === 0x04) continue; // 그림 · 메모 · 각주 표시
    if (code === 0x1e) { buf += '‑'; continue; }
    if (code === 0x1f) continue;
    const props = { ...(find(chpRuns, fc) ?? {}), ...(link ? { link } : {}) };
    if (bufProps && JSON.stringify(props) !== JSON.stringify(bufProps)) flushBuf();
    bufProps = props;
    buf += ch;
  }
  flushBuf();
  if (runs.length) blocks.push(makePara(null));
  closeTable();
  doc.body = blocks.length ? blocks : [para('')];
  // 목록: ilfo 마다 단순 번호/글머리 (LST 를 읽지 않음 — 첫 수준 글머리로)
  for (const b of doc.body) if (b.type === 'p' && b.pPr.num && !doc.lists[b.pPr.num.id]) doc.lists[b.pPr.num.id] = { levels: Array.from({ length: 9 }, (_, i) => ({ fmt: 'bullet', text: ['●', '○', '■'][i % 3], start: 1, ind: { l: 24 * (i + 1), hanging: 18 } })) };
  return { doc, warnings: ['Word 97-2003 문서(.doc)를 열었습니다. 글 · 글자 서식 · 문단 서식 · 표를 가져오며, 그림 · 머리글 · 각주는 빠질 수 있습니다. 저장하면 .docx 입니다.'] };
}
