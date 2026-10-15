// 쪽 나누기 (브라우저 DOM 으로 측정). 글자 모양 · 줄 바꿈은 브라우저가, 쪽 · 단 · 머리글 · 각주는 이 모듈이 정함.
//  - 구역마다 단 너비의 측정 상자에 블록을 문서 순서대로 넣고 위치를 읽음 (떠 있는 그림의 글 흐름도 같게)
//  - 쪽 경계에 걸친 문단은 줄 경계(lines)에서 잘라 두 쪽에 '잘린 보기'(clip)로 그림
//  - 측정한 블록 요소는 블록 객체별로 보관 → 바뀐 블록만 다시 그림 (입력할 때 빠름)
import { sections, computeNumbering, effPara, walkParas, paraLen } from './model.js';
import { blockHtml, tableHtml, tableGrid, DOC_CSS } from './render.js';

const f2 = (v) => Math.round(v * 100) / 100;

/** 측정 상자 + 캐시 */
export class Layout {
  constructor(host) {
    this.host = host; // 화면 밖 숨은 상자 (class="wd")
    this.cache = new Map(); // key(블록 객체) → { el, w, sig }
    this.secBoxes = [];
    this.hfBox = null;
    this.version = 0;
  }

  /** 문서 전체 쪽 나누기 → { pages, byId, noteNums, num } */
  run(doc, opts = {}) {
    const t0 = performance.now();
    const num = computeNumbering(doc);
    const noteNums = noteNumbers(doc);
    const ropts = { num, media: doc.media, marks: !!opts.marks, markup: opts.markup ?? 'all', noteNums, fieldShade: opts.fieldShade, authorColor: opts.authorColor, cmtOpenIn: null, merge: opts.merge ?? null, mergeShade: opts.mergeShade };
    const secs = sections(doc);
    const pages = [];
    const byId = new Map(); // 블록 id → [{ page, item }]
    let pageNo = 0;
    const keep = new Set();
    // 구역 상자 준비
    while (this.secBoxes.length < secs.length) { const b = document.createElement('div'); b.className = 'wd measure-sec'; this.host.append(b); this.secBoxes.push(b); }
    while (this.secBoxes.length > secs.length) this.secBoxes.pop().remove();
    const measured = secs.map((s, si) => this.measureSection(doc, s, si, ropts, keep));
    // 쓰지 않는 캐시 정리
    for (const [k, v] of this.cache) if (!keep.has(k)) { v.el.remove(); this.cache.delete(k); }
    // 쪽 나누기
    let prevPage = null;
    secs.forEach((s, si) => {
      const sec = s.sec;
      const m = measured[si];
      const pgStart = sec.pgNum?.start;
      if (pgStart != null) pageNo = pgStart - 1;
      const continuous = si > 0 && (secs[si - 1] ? s.kind === 'continuous' : false) && prevPage && samePageGeom(prevPage.sec, sec);
      const st = new PageFlow(this, doc, sec, si, ropts, m, pages, byId, () => ++pageNo, continuous ? prevPage : null);
      if (si > 0 && !continuous && (s.kind === 'evenPage' || s.kind === 'oddPage')) {
        // 짝수 · 홀수 쪽에서 시작: 필요하면 빈 쪽 하나
        const next = pageNo + 1;
        if ((s.kind === 'evenPage' && next % 2 === 1) || (s.kind === 'oddPage' && next % 2 === 0)) st.blankPage();
      }
      st.flow();
      prevPage = pages[pages.length - 1];
    });
    // 머리글 · 바닥글 측정 (쪽마다 높이가 같으므로 구역 · 종류별 한 번)
    this.layoutHeaders(doc, pages, ropts);
    const total = pages.length;
    for (const p of pages) p.total = total;
    this.version++;
    return { pages, byId, noteNums, num, ms: performance.now() - t0 };
  }

  /** 구역 블록 측정 → { items: [{ block, el, top, h, lines?, tbl? }], colW } */
  measureSection(doc, s, si, ropts, keep) {
    const sec = s.sec;
    const colW = columnWidth(sec);
    const box = this.secBoxes[si];
    box.style.cssText = `position:absolute;left:0;top:0;width:${f2(colW)}px;visibility:hidden;contain:layout style`;
    box.classList.toggle('hyph', !!doc.settings?.hyphenate);
    const blocks = doc.body.slice(s.start, s.end);
    const items = [];
    const want = [];
    let prevStyle = null;
    blocks.forEach((b, k) => {
      // 같은 스타일 사이 간격 없음 (contextualSpacing) · 이웃 테두리 합치기는 이웃에 따라 HTML 이 달라짐
      const next = blocks[k + 1];
      const ctx = b.type === 'p' ? contextOf(doc, b, prevStyle, next) : '';
      prevStyle = b.type === 'p' ? b.style : null;
      const numSig = b.type === 'p' ? (ropts.num.get(b.id)?.text ?? '') : tblNumSig(b, ropts.num);
      const noteSig = b.type === 'p' && b.runs.some((r) => r.obj === 'note') ? b.runs.filter((r) => r.obj === 'note').map((r) => ropts.noteNums.get(r.id)).join(',') : '';
      const mergeSig = ropts.merge && b.type === 'p' && b.runs.some((r) => r.obj === 'field') ? JSON.stringify(ropts.merge) : (ropts.merge ? 'm' : '') + (ropts.mergeShade ? 'h' : '');
      const sig = `${colW}|${ctx}|${numSig}|${noteSig}|${ropts.marks ? 1 : 0}|${ropts.markup}|${mergeSig}|${doc.themeFonts?.minor?.ea}|${doc.themeFonts?.major?.ea}|${JSON.stringify(doc.defaults)}`;
      let c = this.cache.get(b);
      if (!c || c.sig !== sig || c.styles !== doc.styles || c.media !== doc.media) {
        const tmp = document.createElement('div');
        const o = { ...ropts, ...(ctx ? JSON.parse(ctx) : {}) };
        tmp.innerHTML = blockHtml(doc, b, o);
        const el = tmp.firstElementChild;
        if (c) c.el.remove();
        c = { el, sig, styles: doc.styles, media: doc.media, tabsDone: false };
        this.cache.set(b, c);
      }
      keep.add(b);
      want.push(c.el);
      items.push({ block: b, el: c.el, c });
    });
    this.lastStyles = doc.styles;
    // 상자 안 순서 맞추기 (바뀐 곳만 옮김)
    let cur = box.firstElementChild;
    for (const el of want) {
      if (cur === el) { cur = cur.nextElementSibling; continue; }
      box.insertBefore(el, cur);
    }
    while (cur) { const n = cur.nextElementSibling; if (!want.includes(cur)) cur.remove(); cur = n; }
    // 탭 너비 (Word 탭 위치) — 새로 그린 문단만
    const boxRect = box.getBoundingClientRect();
    for (const it of items) if (!it.c.tabsDone) { fixTabs(doc, it.block, it.el, boxRect.left, colW, ropts); it.c.tabsDone = true; }
    // 위치 읽기 (한 번에)
    const base = boxRect.top;
    const tops = items.map((it) => it.el.getBoundingClientRect().top - base);
    const end = box.getBoundingClientRect().height;
    items.forEach((it, k) => {
      it.top = tops[k];
      it.h = (k + 1 < items.length ? tops[k + 1] : Math.max(end, tops[k] + it.el.offsetHeight)) - tops[k];
      it.h = Math.max(it.h, it.el.offsetHeight);
    });
    return { items, colW, box };
  }

  /** 머리글 · 바닥글: 구역 + 종류별로 측정해 쪽에 붙임 */
  layoutHeaders(doc, pages, ropts) {
    if (!this.hfBox) { this.hfBox = document.createElement('div'); this.hfBox.className = 'wd measure-hf'; this.host.append(this.hfBox); }
    const box = this.hfBox;
    box.innerHTML = '';
    const memo = new Map();
    const measure = (blocks, w) => {
      if (!blocks?.length) return null;
      const key = blocks;
      if (memo.has(key)) return memo.get(key);
      const div = document.createElement('div');
      div.style.cssText = `width:${f2(w)}px`;
      div.innerHTML = blocks.map((b) => blockHtml(doc, b, { ...ropts, fields: { PAGE: '9', NUMPAGES: '99' } })).join('');
      box.append(div);
      for (const p of div.querySelectorAll('.p')) { const blk = findBlock(blocks, p.dataset.pid); if (blk) fixTabs(doc, blk, p, div.getBoundingClientRect().left, w, ropts); }
      const out = { html: div.innerHTML, h: div.getBoundingClientRect().height };
      memo.set(key, out);
      return out;
    };
    for (const pg of pages) {
      const sec = pg.sec;
      const w = sec.page.w - sec.margin.l - sec.margin.r - (sec.margin.gutter ?? 0);
      const kind = sec.titlePg && pg.firstOfSec ? 'first' : doc.settings?.evenOdd && pg.num % 2 === 0 ? 'even' : 'default';
      const hb = sec.headers?.[kind] ?? (kind === 'first' ? null : sec.headers?.default);
      const fb = sec.footers?.[kind] ?? (kind === 'first' ? null : sec.footers?.default);
      pg.hfKind = kind;
      pg.header = measure(hb, w);
      pg.footer = measure(fb, w);
    }
  }
}

function findBlock(blocks, id) {
  let out = null;
  walkParas(blocks, (p) => { if (p.id === id) out = p; });
  return out;
}
const samePageGeom = (a, b) => a && b && Math.abs(a.page.w - b.page.w) < 1 && Math.abs(a.page.h - b.page.h) < 1;
export function columnWidth(sec) {
  const W = sec.page.w - sec.margin.l - sec.margin.r - (sec.margin.gutter ?? 0);
  const n = Math.max(1, sec.cols?.n ?? 1);
  if (sec.cols?.widths?.length === n) return sec.cols.widths[0];
  return (W - (n - 1) * (sec.cols?.space ?? 48)) / n;
}
function tblNumSig(t, num) {
  let s = '';
  for (const r of t.rows) for (const c of r.cells) for (const b of c.blocks) if (b.type === 'p') s += num.get(b.id)?.text ?? '';
  return s;
}
/** 이웃에 따른 문단 그리기 옵션 (같은 스타일 사이 간격 · 테두리 합치기) */
function contextOf(doc, b, prevStyle, next) {
  const ep = effPara(doc, b);
  const o = {};
  if (ep.contextual) {
    if (prevStyle === b.style) o.noBefore = true;
    if (next?.type === 'p' && next.style === b.style) o.noAfter = true;
  }
  if (ep.border && next?.type === 'p') {
    const en = effPara(doc, next);
    if (en.border && JSON.stringify(en.border) === JSON.stringify(ep.border)) o.joinBottom = true;
  }
  return Object.keys(o).length ? JSON.stringify(o) : '';
}

/** 각주 · 미주 번호 (나온 순서) */
export function noteNumbers(doc) {
  const out = new Map();
  let fn = 0; let en = 0;
  walkParas(doc.body, (p) => { for (const r of p.runs) if (r.obj === 'note') out.set(r.id, (r.kind ?? 'foot') === 'end' ? toRomanLower(++en) : String(++fn)); });
  return out;
}
const toRomanLower = (n) => { const R = [[10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; for (const [v, r] of R) while (n >= v) { s += r; n -= v; } return s; };

// ───────────── 탭 ─────────────
/** 문단 안 탭 요소의 너비를 Word 탭 위치에 맞춤 (기본 탭 간격 · 사용자 탭 · 내어쓰기 위치 · 맞춤 · 채움 선) */
export function fixTabs(doc, p, el, colLeft, colW, ropts) {
  const tabsEls = el.querySelectorAll(':scope .tab');
  if (!tabsEls.length) return;
  const ep = effPara(doc, p);
  const num = ropts.num?.get(p.id);
  const ind = { ...(ep.ind ?? {}) };
  if (num?.ind && p.pPr?.ind?.l == null) { ind.l = num.ind.l; ind.hanging = num.ind.hanging; }
  const stops = (ep.tabs ?? []).filter((t) => t.align !== 'clear').map((t) => ({ ...t, pos: Math.min(t.pos, colW) })).sort((a, b) => a.pos - b.pos);
  const def = doc.settings?.defaultTab ?? 48;
  const hang = ind.hanging ? ind.l : null;
  for (const t of tabsEls) {
    t.style.width = '1px';
    t.className = t.className.replace(/\blead-\w+/g, '').trim();
  }
  for (const t of tabsEls) {
    const x = t.getBoundingClientRect().left - colLeft;
    let stop = stops.find((s) => s.pos > x + 0.5);
    if (!stop && hang != null && hang > x + 0.5) stop = { pos: hang, align: 'left' };
    if (!stop) stop = { pos: Math.max(x + 1, Math.floor(x / def + 1) * def), align: 'left' };
    let w = stop.pos - x;
    if (stop.align === 'right' || stop.align === 'center' || stop.align === 'decimal') {
      // 다음 탭(또는 문단 끝)까지 글 너비
      let seg = 0;
      for (let n = t.nextSibling; n; n = n.nextSibling) {
        if (n.nodeType === 1 && (n.classList.contains('tab') || n.classList.contains('end'))) break;
        if (n.nodeType === 1) seg += n.getBoundingClientRect().width;
      }
      if (stop.align === 'decimal') {
        const txt = [];
        for (let n = t.nextSibling; n && !(n.nodeType === 1 && n.classList.contains('tab')); n = n.nextSibling) txt.push(n.textContent);
        const s = txt.join('');
        const k = s.search(/[.,]/);
        if (k >= 0 && s.length) seg *= k / s.length;
      }
      w = stop.align === 'center' ? stop.pos - x - seg / 2 : stop.pos - x - seg;
    }
    w = Math.max(1, w);
    t.style.width = `${f2(w)}px`;
    if (stop.leader && stop.leader !== 'none') t.classList.add(`lead-${stop.leader}`);
  }
}

// ───────────── 줄 경계 ─────────────
/** 문단 요소의 줄 위치 [{ top, bottom }] (요소 위 기준) + 강제 쪽 · 단 나누기 줄 번호 */
export function lineBoxes(el) {
  const base = el.getBoundingClientRect().top;
  const rects = [];
  const range = document.createRange();
  const forced = [];
  const walk = (node) => {
    for (let n = node.firstChild; n; n = n.nextSibling) {
      if (n.nodeType === 3) {
        if (!n.data.length) continue;
        range.selectNodeContents(n);
        for (const r of range.getClientRects()) if (r.height > 0) rects.push([r.top - base, r.bottom - base]);
      } else if (n.nodeType === 1) {
        const cl = n.classList;
        if (cl.contains('fl')) continue; // 떠 있는 그림은 줄이 아님
        if (n.tagName === 'IMG' || cl.contains('obj') || cl.contains('lbl') || cl.contains('tab') || n.tagName === 'svg') {
          const r = n.getBoundingClientRect();
          if (r.height > 0) rects.push([r.top - base, r.bottom - base]);
          if (cl.contains('obj') || cl.contains('lbl')) continue;
        }
        if (n.dataset?.br === 'page' || n.dataset?.br === 'column') { const r = n.getBoundingClientRect(); forced.push({ y: r.top - base + 1, kind: n.dataset.br }); }
        walk(n);
      }
    }
  };
  walk(el);
  rects.sort((a, b) => a[0] - b[0]);
  const lines = [];
  for (const [t, b] of rects) {
    const last = lines[lines.length - 1];
    // 같은 줄: 세로로 겹침 (아래 글자의 위가 앞 줄 가운데보다 위)
    if (last && t < (last.top + last.bottom) / 2) { last.top = Math.min(last.top, t); last.bottom = Math.max(last.bottom, b); }
    else lines.push({ top: t, bottom: b });
  }
  const forcedLines = forced.map((f) => ({ line: lines.findIndex((l) => f.y >= l.top - 1 && f.y <= l.bottom + 1), kind: f.kind })).filter((f) => f.line >= 0);
  return { lines, forced: forcedLines };
}
/** 줄 k 앞에서 자를 위치 (요소 위 기준): 앞 줄 아래와 이 줄 위의 가운데 */
const cutBefore = (lines, k, h) => (k <= 0 ? 0 : k >= lines.length ? h : (lines[k - 1].bottom + lines[k].top) / 2);

// ───────────── 쪽 채우기 ─────────────
class PageFlow {
  constructor(layout, doc, sec, si, ropts, m, pages, byId, nextNo, contPage) {
    Object.assign(this, { layout, doc, sec, si, ropts, m, pages, byId, nextNo });
    this.colW = m.colW;
    this.ncols = Math.max(1, sec.cols?.n ?? 1);
    this.page = null;
    this.firstOfSec = true;
    if (contPage) {
      // 이어지는 구역: 같은 쪽에 새 띠 (band)
      this.page = contPage;
      const band = contPage.bands[contPage.bands.length - 1];
      const used = bandHeight(band);
      this.balance(band);
      const top = band.top + bandHeight(band);
      void used;
      this.band = this.newBand(top);
      this.firstOfSec = false;
    }
  }
  pageGeom() {
    const s = this.sec;
    return { top: s.margin.t, bottom: s.page.h - s.margin.b, left: s.margin.l + (s.margin.gutter ?? 0), w: s.page.w - s.margin.l - s.margin.r - (s.margin.gutter ?? 0) };
  }
  newPage() {
    const g = this.pageGeom();
    const no = this.nextNo();
    const page = { index: this.pages.length, num: no, numText: pageNumText(no, this.sec.pgNum?.fmt), sec: this.sec, si: this.si, w: this.sec.page.w, h: this.sec.page.h, bodyTop: g.top, bodyBottom: g.bottom, bodyLeft: g.left, bodyW: g.w, bands: [], notes: [], noteH: 0, firstOfSec: this.firstOfSec };
    this.firstOfSec = false;
    this.pages.push(page);
    this.page = page;
    this.band = this.newBand(0);
    return page;
  }
  blankPage() { this.newPage(); this.page.blank = true; this.page = null; }
  newBand(top) {
    const band = { top, cols: Array.from({ length: this.ncols }, () => ({ items: [], y: 0 })), col: 0, ncols: this.ncols, colW: this.colW, space: this.sec.cols?.space ?? 48, sep: !!this.sec.cols?.sep, widths: this.sec.cols?.widths, si: this.si };
    this.page.bands.push(band);
    return band;
  }
  /** 이 단에 남은 높이 */
  avail() {
    const body = this.page.bodyBottom - this.page.bodyTop - this.page.noteH;
    return body - this.band.top - this.band.cols[this.band.col].y;
  }
  nextColumn() {
    if (this.band.col + 1 < this.ncols) { this.band.col++; return; }
    this.newPage();
  }
  place(item, h) {
    const col = this.band.cols[this.band.col];
    item.y = col.y;
    item.h = h;
    item.page = this.page.index;
    col.items.push(item);
    col.y += h;
    const id = item.block.id;
    if (!this.byId.has(id)) this.byId.set(id, []);
    this.byId.get(id).push(item);
  }
  /** 쪽 위인가 (이 단에 아직 아무것도 없음) */
  atTop() { return this.band.cols[this.band.col].y <= 0.5 && this.band.cols[this.band.col].items.length === 0; }

  flow() {
    if (!this.page) this.newPage();
    const items = this.m.items;
    for (let k = 0; k < items.length; k++) {
      const it = items[k];
      const b = it.block;
      if (b.type === 'p') this.flowPara(it, items, k);
      else this.flowTable(it);
    }
  }

  /** keepNext 사슬: 이 문단부터 keepNext 가 이어지는 동안 높이 합 (다음 블록의 첫 줄까지) */
  keepChainH(items, k) {
    let h = 0;
    let i = k;
    while (i < items.length) {
      const it = items[i];
      if (it.block.type !== 'p') { h += Math.min(it.h, 40); break; }
      const ep = effPara(this.doc, it.block);
      if (i > k && !ep.keepNext) { h += firstLineH(it); break; }
      h += it.h;
      if (!ep.keepNext) break;
      i++;
    }
    return h;
  }

  flowPara(it, items, k) {
    const doc = this.doc;
    const p = it.block;
    const ep = effPara(doc, p);
    if (ep.pageBreakBefore && !this.atTop()) this.newPage();
    // 각주
    const notes = p.runs.some((r) => r.obj === 'note' && (r.kind ?? 'foot') === 'foot') ? this.noteRefs(it) : null;
    let from = 0; // 이미 놓은 높이 (요소 위 기준)
    const H = it.h;
    let lines = null;
    const getLines = () => { if (!lines) lines = lineBoxes(it.el); return lines; };
    // 강제 쪽 나누기 (문단 안)
    const hasBreak = p.runs.some((r) => r.obj === 'br' && (r.kind === 'page' || r.kind === 'column'));
    let guard = 0;
    while (from < H - 0.5 && guard++ < 500) {
      const avail = this.avail();
      const rest = H - from;
      // 문단 안의 강제 나누기
      let forcedCut = null;
      if (hasBreak) {
        const L = getLines();
        for (const f of L.forced) {
          const c = cutBefore(L.lines, f.line + 1, H);
          if (c > from + 0.5 && (forcedCut == null || c < forcedCut.y)) forcedCut = { y: c, kind: f.kind };
        }
      }
      const restEnd = forcedCut ? forcedCut.y : H;
      const need = restEnd - from;
      const noteH = notes ? this.notesHeight(notes, from, restEnd) : 0;
      // keepNext: 다음 블록 첫 줄까지 같은 쪽에
      const chain = from === 0 && ep.keepNext && !forcedCut ? this.keepChainH(items, k) : need;
      const fitsWhole = need + noteH <= avail + 0.5 && (from > 0 || chain <= avail + 0.5 || this.atTop() || chain > this.pageBodyH());
      if (fitsWhole) {
        this.addNotes(notes, from, restEnd);
        this.place({ kind: 'p', block: p, el: it.el, clip: from > 0 || restEnd < H ? { from, to: restEnd } : null }, need);
        from = restEnd;
        if (forcedCut) { if (forcedCut.kind === 'column') this.nextColumn(); else this.newPage(); }
        continue;
      }
      // 줄 단위로 자르기
      if ((ep.keepLines && from === 0 && !this.atTop() && need <= this.pageBodyH()) || (!this.atTop() && from === 0 && ep.keepNext && chain <= this.pageBodyH())) { this.nextColumn(); continue; }
      const L = getLines();
      const lns = L.lines;
      // from 이후 첫 줄 번호
      let first = lns.findIndex((l) => (l.top + l.bottom) / 2 > from);
      if (first < 0) { this.place({ kind: 'p', block: p, el: it.el, clip: { from, to: H } }, H - from); break; }
      let k2 = first;
      while (k2 < lns.length) {
        const cut = cutBefore(lns, k2 + 1, H);
        const nh = notes ? this.notesHeight(notes, from, cut) : 0;
        if (cut - from + nh > avail + 0.5) break;
        if (forcedCut && cut > forcedCut.y + 0.5) break;
        k2++;
      }
      let fitLines = k2 - first;
      // 외톨이 줄 방지: 쪽 끝에 첫 줄 하나만 · 다음 쪽에 마지막 줄 하나만 남기지 않음
      const totalLines = lns.length - first;
      if (ep.widow !== false && from === 0 && fitLines === 1 && totalLines > 1 && !this.atTop()) fitLines = 0;
      if (ep.widow !== false && fitLines > 1 && totalLines - fitLines === 1 && totalLines > 2) fitLines--;
      if (fitLines <= 0) {
        if (this.atTop()) fitLines = 1; // 쪽보다 큰 줄: 그냥 놓음
        else { this.nextColumn(); continue; }
      }
      const cut = cutBefore(lns, first + fitLines, H);
      this.addNotes(notes, from, cut);
      this.place({ kind: 'p', block: p, el: it.el, clip: { from, to: cut } }, cut - from);
      from = cut;
      if (from < H - 0.5) this.nextColumn();
    }
  }
  pageBodyH() { return this.page.bodyBottom - this.page.bodyTop; }

  flowTable(it) {
    const t = it.block;
    const el = it.el;
    const trs = [...el.querySelectorAll(':scope > tbody > tr')];
    const base = el.getBoundingClientRect().top;
    const rowTop = trs.map((tr) => tr.getBoundingClientRect().top - base);
    const rowH = trs.map((tr, i) => (i + 1 < trs.length ? rowTop[i + 1] : el.getBoundingClientRect().height) - rowTop[i]);
    const header = [];
    for (let r = 0; r < t.rows.length && t.rows[r].header; r++) header.push(r);
    const headH = header.reduce((a, r) => a + rowH[r], 0);
    const info = tableGrid(t);
    // 세로 병합으로 묶인 행은 함께 (병합 칸이 쪽을 넘지 않게)
    const groupEnd = t.rows.map((_, r) => { let e = r; info[r].forEach((ci) => { if (ci.rowSpan > 1) e = Math.max(e, r + ci.rowSpan - 1); }); return e; });
    let r = 0;
    let firstFrag = true;
    let guard = 0;
    while (r < t.rows.length && guard++ < 10000) {
      const avail = this.avail();
      const rows = [];
      let h = firstFrag ? 0 : headH;
      const rep = firstFrag ? [] : header;
      let i = r;
      while (i < t.rows.length) {
        let e = i;
        for (let k = i; k <= e; k++) e = Math.max(e, groupEnd[k]);
        let gh = 0;
        for (let k = i; k <= e; k++) gh += rowH[k];
        if (h + gh > avail + 0.5 && (rows.length || !this.atTop())) break;
        for (let k = i; k <= e; k++) rows.push(k);
        h += gh;
        i = e + 1;
        if (t.rows[i - 1]?.cantSplit === false) { /* 행 나누기 허용은 단순화: 행 단위 */ }
      }
      if (!rows.length) { this.nextColumn(); continue; }
      this.place({ kind: 'tbl', block: t, el, rows, repeat: rep, rowTop, rowH }, h);
      r = rows[rows.length - 1] + 1;
      firstFrag = false;
      if (r < t.rows.length) this.nextColumn();
    }
  }

  // ── 각주 ──
  noteRefs(it) {
    const base = it.el.getBoundingClientRect().top;
    return [...it.el.querySelectorAll('.note[data-kind="foot"]')].map((n) => ({ id: n.dataset.note, y: n.getBoundingClientRect().top - base + 1 }));
  }
  notesHeight(refs, from, to) {
    let h = 0;
    for (const r of refs) if (r.y >= from && r.y < to && !this.page.notes.includes(r.id)) h += this.noteH(r.id);
    if (h && !this.page.notes.length) h += 20; // 구분선
    return h;
  }
  addNotes(refs, from, to) {
    if (!refs) return;
    for (const r of refs) if (r.y >= from && r.y < to && !this.page.notes.includes(r.id)) {
      if (!this.page.notes.length) this.page.noteH += 20;
      this.page.notes.push(r.id);
      this.page.noteH += this.noteH(r.id);
    }
  }
  noteH(id) {
    const lay = this.layout;
    lay.noteCache ??= new Map();
    const blocks = this.doc.notes?.foot?.[id] ?? [];
    const key = `${id}|${this.colW}`;
    const c = lay.noteCache.get(key);
    if (c && c.blocks === blocks) return c.h;
    const div = document.createElement('div');
    div.className = 'wd';
    div.style.cssText = `position:absolute;visibility:hidden;width:${f2(this.colW)}px`;
    div.innerHTML = noteBlocksHtml(this.doc, blocks, this.ropts.noteNums.get(id), this.ropts);
    lay.host.append(div);
    const h = div.getBoundingClientRect().height;
    const html = div.innerHTML;
    div.remove();
    lay.noteCache.set(key, { blocks, h, html });
    return h;
  }

  /** 단 균형 (이어지는 구역 나누기 앞): 단 높이를 줄여 다시 채움 — 문단 단위 근사 */
  balance(band) {
    if (band.ncols < 2) return;
    const all = band.cols.flatMap((c) => c.items);
    if (!all.length) return;
    const total = all.reduce((a, x) => a + x.h, 0);
    let target = total / band.ncols;
    for (let tries = 0; tries < 20; tries++) {
      const cols = Array.from({ length: band.ncols }, () => ({ items: [], y: 0 }));
      let c = 0;
      for (const x of all) { if (cols[c].y + x.h > target + 0.5 && cols[c].items.length && c + 1 < band.ncols) c++; x.y = cols[c].y; cols[c].items.push(x); cols[c].y += x.h; }
      if (cols[band.ncols - 1].y <= target + 0.5) { band.cols = cols; return; }
      target *= 1.08;
    }
  }
}
const bandHeight = (band) => Math.max(0, ...band.cols.map((c) => c.y));
function firstLineH(it) {
  if (it.block.type !== 'p') return 40;
  try { const L = lineBoxes(it.el).lines; return L.length ? cutBefore(L, 1, it.h) : it.h; } catch { return Math.min(it.h, 30); }
}
export function pageNumText(n, fmt) {
  if (!fmt || fmt === 'decimal') return String(n);
  if (fmt === 'upperRoman' || fmt === 'lowerRoman') { const R = [[1000, 'm'], [900, 'cm'], [500, 'd'], [400, 'cd'], [100, 'c'], [90, 'xc'], [50, 'l'], [40, 'xl'], [10, 'x'], [9, 'ix'], [5, 'v'], [4, 'iv'], [1, 'i']]; let s = ''; let v = n; for (const [a, r] of R) while (v >= a) { s += r; v -= a; } return fmt === 'upperRoman' ? s.toUpperCase() : s; }
  if (fmt === 'upperLetter') return String.fromCharCode(64 + ((n - 1) % 26) + 1);
  if (fmt === 'lowerLetter') return String.fromCharCode(96 + ((n - 1) % 26) + 1);
  if (fmt === 'dash') return `- ${n} -`;
  if (fmt === 'ganada') return '가나다라마바사아자차카타파하'[(n - 1) % 14];
  return String(n);
}
/** 각주 내용 HTML (번호 + 블록) */
export function noteBlocksHtml(doc, blocks, numText, ropts) {
  const bs = blocks?.length ? blocks : [{ type: 'p', id: 'empty', style: 'FootnoteText', pPr: {}, runs: [] }];
  return bs.map((b, i) => {
    let h = blockHtml(doc, b, { ...ropts, num: ropts.num });
    if (i === 0) h = h.replace(/(<div class="p[^"]*"[^>]*>)/, `$1<span class="note-num" style="vertical-align:super;font-size:7pt">${numText ?? ''}</span> `);
    return h;
  }).join('');
}

// ───────────── 쪽 HTML ─────────────
/** 한 쪽의 본문 · 머리글 · 바닥글 · 각주 HTML (편집 화면 · 인쇄 공용). 문단은 측정 요소를 복제해 씀 */
export function pageNode(doc, page, res, opts = {}) {
  const pg = document.createElement('div');
  pg.className = `page wd${opts.marks ? ' marks' : ''}${doc.settings?.hyphenate ? ' hyph' : ''}`;
  pg.dataset.page = String(page.index);
  pg.style.cssText = `width:${f2(page.w)}px;height:${f2(page.h)}px`;
  if (doc.settings?.pageColor) pg.style.background = doc.settings.pageColor;
  const sec = page.sec;
  const fields = { PAGE: page.numText, NUMPAGES: String(page.total ?? res.pages.length), SECTIONPAGES: String(page.total ?? res.pages.length) };
  // 머리글
  const hdrTop = sec.margin.header ?? 48;
  let bodyTop = page.bodyTop;
  if (page.header) {
    const hd = document.createElement('div');
    hd.className = 'hdr hf';
    hd.style.cssText = `position:absolute;left:${f2(page.bodyLeft)}px;top:${f2(hdrTop)}px;width:${f2(page.bodyW)}px`;
    hd.innerHTML = page.header.html;
    setFields(hd, fields);
    pg.append(hd);
    bodyTop = Math.max(bodyTop, hdrTop + page.header.h);
  }
  if (page.footer) {
    const fd = document.createElement('div');
    fd.className = 'ftr hf';
    fd.style.cssText = `position:absolute;left:${f2(page.bodyLeft)}px;bottom:${f2(sec.margin.footer ?? 48)}px;width:${f2(page.bodyW)}px`;
    fd.innerHTML = page.footer.html;
    setFields(fd, fields);
    pg.append(fd);
  }
  page.bodyTopEff = bodyTop;
  // 페이지 테두리 · 워터마크
  if (sec.pageBorder) {
    const b = sec.pageBorder;
    const sp = ((b.space ?? 24) * 4) / 3;
    const bd = document.createElement('div');
    bd.className = 'pgborder';
    bd.style.cssText = `position:absolute;inset:${f2(sp)}px;border:${f2(Math.max(1, ((b.w ?? 1) * 4) / 3))}px ${b.style === 'double' ? 'double' : b.style === 'dashed' ? 'dashed' : b.style === 'dotted' ? 'dotted' : 'solid'} ${b.color ?? '#000'};pointer-events:none`;
    pg.prepend(bd);
  }
  const wm = doc.settings?.watermark;
  if (wm?.text) {
    const w = document.createElement('div');
    w.className = 'wmark';
    const size = wm.size || Math.min(140, (page.w * 0.9) / Math.max(2, wm.text.length) * 1.4);
    w.style.cssText = `position:absolute;left:50%;top:50%;transform:translate(-50%,-50%)${wm.diagonal !== false ? ' rotate(-45deg)' : ''};font:${f2(size)}px 'Malgun Gothic','맑은 고딕',sans-serif;color:${wm.color ?? '#D9D9D9'};opacity:.6;white-space:nowrap;pointer-events:none;user-select:none;z-index:0`;
    w.textContent = wm.text;
    pg.prepend(w);
  }
  // 본문 띠 · 단
  const body = document.createElement('div');
  body.className = 'body';
  body.style.cssText = `position:absolute;left:${f2(page.bodyLeft)}px;top:${f2(bodyTop)}px;width:${f2(page.bodyW)}px;height:${f2(page.bodyBottom - bodyTop)}px`;
  pg.append(body);
  for (const band of page.bands) {
    let x = 0;
    band.cols.forEach((col, ci) => {
      const w = band.widths?.[ci] ?? band.colW;
      const cd = document.createElement('div');
      cd.className = 'col';
      cd.style.cssText = `position:absolute;left:${f2(x)}px;top:${f2(band.top)}px;width:${f2(w)}px`;
      if (band.sep && ci > 0) cd.style.borderLeft = '1px solid #000';
      for (const item of col.items) {
        cd.append(itemNode(doc, item, res, opts, fields));
        if (item.kind === 'p' && (!item.clip || item.clip.from === 0)) floatObjects(doc, item.block, pg, { page, colX: page.bodyLeft + x, top: bodyTop + band.top + (item.y ?? 0), opts, res });
      }
      body.append(cd);
      x += w + (band.space ?? 48);
    });
  }
  // 각주
  if (page.notes.length) {
    const nd = document.createElement('div');
    nd.className = 'notes';
    nd.style.cssText = `position:absolute;left:${f2(page.bodyLeft)}px;bottom:${f2(sec.margin.b)}px;width:${f2(page.bodyW)}px`;
    let h = '<div class="note-sep"></div>';
    for (const id of page.notes) {
      const c = res.layout?.noteCache?.get(`${id}|${columnWidth(sec)}`);
      h += `<div class="note-body" data-note="${id}">${c?.html ?? noteBlocksHtml(doc, doc.notes?.foot?.[id], res.noteNums.get(id), { num: res.num, noteNums: res.noteNums, media: doc.media })}</div>`;
    }
    nd.innerHTML = h;
    pg.append(nd);
  }
  return pg;
}
/** 글 뒤 · 글 앞 · 배치 없음 그림 · 글상자: 쪽 기준 절대 위치 */
function floatObjects(doc, p, pg, { page, colX, top, opts, res }) {
  for (const r of p.runs ?? []) {
    const isShape = r.obj === 'shape';
    if (!isShape && !(r.obj === 'img' && r.float && ['behind', 'front', 'none'].includes(r.float.wrap))) continue;
    const fl = r.float ?? { relH: 'column', relV: 'paragraph', x: 0, y: 0, wrap: 'front' };
    const relH = fl.relH ?? 'column';
    const relV = fl.relV ?? 'paragraph';
    let x = fl.x ?? 0;
    let y = fl.y ?? 0;
    if (fl.alignH) {
      const [l, w] = relH === 'page' ? [0, page.w] : [page.bodyLeft, page.bodyW];
      x = fl.alignH === 'center' ? l + (w - r.w) / 2 : fl.alignH === 'right' ? l + w - r.w : l;
    } else x += relH === 'page' ? 0 : relH === 'margin' ? page.bodyLeft : colX;
    if (fl.alignV) {
      const [t, h] = relV === 'page' ? [0, page.h] : [page.bodyTop, page.bodyBottom - page.bodyTop];
      y = fl.alignV === 'center' ? t + (h - r.h) / 2 : fl.alignV === 'bottom' ? t + h - r.h : t;
    } else y += relV === 'page' ? 0 : relV === 'margin' ? page.bodyTop : top;
    const box = document.createElement('div');
    box.className = `fobj${fl.wrap === 'behind' ? ' behind' : ''}`;
    box.dataset.anchorPid = p.id;
    box.dataset.img = r.id ?? '';
    const css = [`position:absolute`, `left:${f2(x)}px`, `top:${f2(y)}px`, `width:${f2(r.w ?? 100)}px`, `height:${f2(r.h ?? 60)}px`, `z-index:${fl.wrap === 'behind' ? 0 : 3}`];
    if (r.rot) css.push(`transform:rotate(${r.rot}deg)`);
    if (isShape) {
      if (r.fill) css.push(`background:${r.fill}`);
      if (r.line) css.push(`border:${f2(Math.max(0.75, ((r.line.w ?? 0.75) * 4) / 3))}px solid ${r.line.color ?? '#000'}`);
      if (r.prst === 'ellipse') css.push('border-radius:50%');
      else if (r.prst === 'roundRect') css.push('border-radius:10px');
      const ins = r.inset ?? { l: 9.6, r: 9.6, t: 4.8, b: 4.8 };
      css.push(`padding:${f2(ins.t)}px ${f2(ins.r)}px ${f2(ins.b)}px ${f2(ins.l)}px`, 'box-sizing:border-box', 'overflow:hidden');
      if (r.anchor === 'ctr') css.push('display:flex;flex-direction:column;justify-content:center');
      box.innerHTML = r.svg ?? (r.blocks ?? []).map((b) => blockHtml(doc, b, { num: res.num, media: doc.media, noteNums: res.noteNums, marks: opts.marks })).join('');
    } else {
      const src = doc.media?.[r.media] ?? r.src ?? '';
      box.innerHTML = r.svg ? `<span class="svgobj" style="display:block;width:100%;height:100%">${r.svg}</span>` : `<img src="${src.replace(/"/g, '&quot;')}" alt="" draggable="false" style="width:100%;height:100%;display:block">`;
    }
    box.style.cssText = css.join(';');
    if (fl.wrap === 'behind') pg.prepend(box); else pg.append(box);
  }
}
function setFields(root, fields) {
  for (const f of root.querySelectorAll('[data-field]')) { const v = fields[f.dataset.field]; if (v != null) f.textContent = v; }
}
function itemNode(doc, item, res, opts, fields) {
  if (item.kind === 'p') {
    const el = item.el.cloneNode(true);
    setFields(el, fields);
    if (!item.clip) return el;
    const wrap = document.createElement('div');
    wrap.className = 'frag';
    wrap.dataset.pid = item.block.id;
    wrap.dataset.from = String(item.clip.from);
    wrap.style.cssText = `height:${f2(item.clip.to - item.clip.from)}px;overflow:hidden;position:relative`;
    const inner = document.createElement('div');
    inner.style.marginTop = `${f2(-item.clip.from)}px`;
    inner.append(el);
    wrap.append(inner);
    return wrap;
  }
  // 표 조각: 반복 머리글 + 행
  const t = item.block;
  const full = item.rows.length === t.rows.length && !item.repeat.length;
  if (full) { const el = item.el.cloneNode(true); setFields(el, fields); return el; }
  const tmp = document.createElement('div');
  tmp.innerHTML = tableHtml(doc, t, { num: res.num, media: doc.media, noteNums: res.noteNums, marks: opts.marks, markup: opts.markup, rows: [...item.repeat, ...item.rows] });
  const el = tmp.firstElementChild;
  el.dataset.frag = '1';
  for (const tr of el.querySelectorAll(':scope > tbody > tr')) if (item.repeat.includes(Number(tr.dataset.r))) tr.classList.add('rep');
  setFields(el, fields);
  return el;
}

