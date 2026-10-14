// 편집 화면: 쪽 그리기 (보이는 쪽만) · 확대/축소 · 캐럿 · 선택 · 화면 좌표 ↔ 문서 위치
import { S, emit } from './state.js';
import { Layout, pageNode, lineBoxes } from './layout.js';
import { walkParas, blockAt, paraLen, orderPos, cmpPos, posEq, sections } from './model.js';
import { DOC_CSS } from './render.js';

const f2 = (v) => Math.round(v * 100) / 100;
const GAP = 16; // 쪽 사이 간격 (화면 px)

export const V = {
  scroller: null, pagesEl: null, host: null, layout: null,
  holders: [], rendered: new Map(), // 쪽 번호 → { node, sig }
  pathById: new Map(), orderById: new Map(),
  caretEl: null, caretX: null, // 위아래 이동 때 유지할 x
  blinkT: null,
};

export function initView(scroller) {
  V.scroller = scroller;
  V.pagesEl = document.createElement('div');
  V.pagesEl.className = 'pages';
  scroller.append(V.pagesEl);
  const style = document.createElement('style');
  style.textContent = DOC_CSS;
  document.head.append(style);
  V.host = document.createElement('div');
  V.host.className = 'measure-host';
  V.host.setAttribute('aria-hidden', 'true');
  document.body.append(V.host);
  V.layout = new Layout(V.host);
  let raf = 0;
  scroller.addEventListener('scroll', () => { if (!raf) raf = requestAnimationFrame(() => { raf = 0; renderVisible(); emit('scroll'); }); }, { passive: true });
  new ResizeObserver(() => { if (S.zoomMode) applyZoomMode(); renderVisible(); }).observe(scroller);
}

/** 문서 다시 배치 (변경 후) */
export function relayout() {
  const doc = S.doc;
  V.pathById = new Map();
  V.orderById = new Map();
  let k = 0;
  walkParas(doc.body, (p, path) => { V.pathById.set(p.id, path); V.orderById.set(p.id, k++); });
  // 스토리 (머리글 · 바닥글 · 각주 · 미주)
  sections(doc).forEach((s, si) => {
    for (const [prop, ch] of [['headers', 'h'], ['footers', 'f']]) for (const [which, bs] of Object.entries(s.sec?.[prop] ?? {})) {
      if (!bs) continue;
      let kk = 1e7 + si * 1e5 + (ch === 'f' ? 5e4 : 0);
      walkParas(bs, (p, path) => { V.pathById.set(p.id, path); V.orderById.set(p.id, kk++); }, [`${ch}:${si}:${which}`]);
    }
  });
  for (const kind of ['foot', 'end']) for (const [id, bs] of Object.entries(doc.notes?.[kind] ?? {})) { let kk = 2e7; walkParas(bs ?? [], (p, path) => { V.pathById.set(p.id, path); V.orderById.set(p.id, kk++); }, [`n:${kind}:${id}`]); }
  const mg = S.merge?.preview && S.merge.rows?.length ? Object.fromEntries((S.merge.cols ?? []).map((c, i) => [c, S.merge.rows[S.merge.index ?? 0]?.[i] ?? ''])) : null;
  const res = V.layout.run(doc, { marks: S.marks, markup: S.markup, fieldShade: S.opts.fieldShade !== false, authorColor, merge: mg, mergeShade: !!S.opts.mergeShade });
  res.layout = V.layout;
  S.layout = res;
  buildHolders(res);
  renderVisible(true);
  drawSelection();
  emit('layout', res);
  return res;
}
const AUTHOR_COLORS = ['#C00000', '#2E75B6', '#548235', '#7030A0', '#BF8F00', '#C55A11'];
const authorIdx = new Map();
export function authorColor(name) { if (!authorIdx.has(name)) authorIdx.set(name, authorIdx.size); return AUTHOR_COLORS[authorIdx.get(name) % AUTHOR_COLORS.length]; }

function buildHolders(res) {
  const z = S.zoom;
  const multi = S.zoomMode === 'multi';
  V.pagesEl.classList.toggle('multi', multi);
  while (V.holders.length < res.pages.length) {
    const h = document.createElement('div');
    h.className = 'page-holder';
    V.pagesEl.append(h);
    V.holders.push(h);
  }
  while (V.holders.length > res.pages.length) { const h = V.holders.pop(); h.remove(); V.rendered.delete(V.holders.length); }
  res.pages.forEach((p, i) => {
    const h = V.holders[i];
    h.style.width = `${f2(p.w * z)}px`;
    h.style.height = `${f2(p.h * z)}px`;
    h.dataset.page = String(i);
  });
}

/** 쪽 내용 서명 (바뀐 쪽만 다시 그림) */
const elIds = new WeakMap();
let elSeq = 0;
const eid = (el) => { let v = elIds.get(el); if (!v) { v = ++elSeq; elIds.set(el, v); } return v; };
function pageSig(p) {
  let s = `${p.w}x${p.h}|${p.numText}|${p.total}|${p.header?.html.length ?? 0}:${p.header?.html.slice(0, 64) ?? ''}|${p.footer?.html.length ?? 0}|${p.notes.join(',')}|${S.marks}|${S.zoom}`;
  for (const b of p.bands) for (const c of b.cols) for (const it of c.items) s += `|${eid(it.el)}:${it.clip ? `${f2(it.clip.from)}-${f2(it.clip.to)}` : ''}${it.rows ? `r${it.rows[0]}-${it.rows[it.rows.length - 1]}` : ''}`;
  if (p.header) s += eid(p.header);
  if (p.footer) s += eid(p.footer);
  return s;
}

/** 보이는 쪽 (±1 화면) 만 그림 */
export function renderVisible(force = false) {
  const res = S.layout;
  if (!res) return;
  const top = V.scroller.scrollTop;
  const vh = V.scroller.clientHeight;
  const lo = top - vh;
  const hi = top + vh * 2;
  const keepIdx = new Set();
  // 캐럿이 있는 쪽은 항상 (위아래 이동 · 입력)
  const cp = caretPage();
  res.pages.forEach((p, i) => {
    const h = V.holders[i];
    const y = h.offsetTop;
    const vis = (y + h.offsetHeight >= lo && y <= hi) || i === cp || i === cp - 1 || i === cp + 1;
    if (!vis) return;
    keepIdx.add(i);
    const sig = pageSig(p);
    const cur = V.rendered.get(i);
    if (cur && cur.sig === sig && !force) return;
    if (cur && cur.sig === sig && cur.node.isConnected) return;
    const node = pageNode(S.doc, p, res, { marks: S.marks, markup: S.markup });
    node.style.transform = S.zoom !== 1 ? `scale(${S.zoom})` : '';
    node.style.transformOrigin = '0 0';
    const ov = document.createElement('div');
    ov.className = 'overlay';
    node.append(ov);
    if (S.view === 'print' && S.opts.showBoundaries) node.classList.add('bounds');
    h0(i).replaceChildren(node);
    V.rendered.set(i, { node, sig });
  });
  for (const [i, r] of V.rendered) if (!keepIdx.has(i)) { r.node.remove(); V.rendered.delete(i); }
  drawSelection();
}
const h0 = (i) => V.holders[i];

// ───────────── 확대/축소 ─────────────
export function setZoom(z, mode = null) {
  S.zoom = Math.max(0.1, Math.min(5, z));
  S.zoomMode = mode;
  for (const r of V.rendered.values()) r.node.remove();
  V.rendered.clear();
  if (S.layout) { buildHolders(S.layout); renderVisible(true); }
  emit('zoom', S.zoom);
}
export function applyZoomMode() {
  const p = S.layout?.pages[0];
  if (!p) return;
  const W = V.scroller.clientWidth - 40;
  const H = V.scroller.clientHeight - 40;
  if (S.zoomMode === 'pageWidth') setZoom(W / p.w, 'pageWidth');
  else if (S.zoomMode === 'onePage') setZoom(Math.min(W / p.w, H / p.h), 'onePage');
  else if (S.zoomMode === 'multi') setZoom(Math.min((W - GAP) / (2 * p.w), H / p.h), 'multi');
}

// ───────────── 위치 ↔ DOM ─────────────
/** 표시된 문단 요소들 (쪽에 나뉘었으면 여러 개) */
export function paraEls(pid) {
  const out = [];
  const pref = S.hf?.page;
  const entries = [...V.rendered.entries()].sort((a, b) => (a[0] === pref ? -1 : b[0] === pref ? 1 : a[0] - b[0]));
  for (const [, r] of entries) for (const el of r.node.querySelectorAll(`.p[data-pid="${pid}"]`)) out.push(el);
  return out;
}
const isTextRun = (sp) => sp.classList.contains('r') && !/\b(tab|obj|fld|note|pbr|cbr|lbr|anc|end|bm)\b/.test(sp.className);
function textNodes(node, out = []) {
  for (let n = node.firstChild; n; n = n.nextSibling) {
    if (n.nodeType === 3) out.push(n);
    else if (n.nodeType === 1 && !n.classList.contains('pm') && !n.classList.contains('tabm') && !n.classList.contains('brm') && !n.classList.contains('lbrm')) textNodes(n, out);
  }
  return out;
}
/** 문단 요소 안 위치 o → DOM { node, offset } */
export function domPoint(el, o) {
  const spans = el.querySelectorAll(':scope > .r');
  let best = null;
  for (const sp of spans) {
    const s0 = Number(sp.dataset.o);
    if (sp.classList.contains('end')) { if (o >= s0) best = { span: sp, kind: 'end' }; break; }
    if (isTextRun(sp)) {
      const len = sp.textContent.length;
      if (o >= s0 && o < s0 + len) { best = { span: sp, kind: 'text', off: o - s0 }; break; }
      if (o === s0 + len) best = { span: sp, kind: 'text', off: len };
    } else if (sp.classList.contains('bm')) continue;
    else {
      if (o === s0) { best = { span: sp, kind: 'before' }; break; }
      if (o === s0 + 1) best = { span: sp, kind: 'after' };
    }
  }
  if (!best) { const last = spans[spans.length - 1]; return last ? { node: last, offset: 0 } : { node: el, offset: 0 }; }
  if (best.kind === 'end') { const t = textNodes(best.span)[0]; return t ? { node: t, offset: 0 } : { node: best.span, offset: 0 }; }
  if (best.kind === 'before') return { node: best.span.parentNode, offset: [...best.span.parentNode.childNodes].indexOf(best.span) };
  if (best.kind === 'after') return { node: best.span.parentNode, offset: [...best.span.parentNode.childNodes].indexOf(best.span) + 1 };
  let k = best.off;
  for (const t of textNodes(best.span)) { if (k <= t.data.length) return { node: t, offset: k }; k -= t.data.length; }
  return { node: best.span, offset: best.span.childNodes.length };
}
/** DOM 위치 → 문서 위치 */
export function posFromDom(node, offset) {
  let el = node.nodeType === 1 ? node : node.parentElement;
  if (el?.closest?.('.lbl')) { const p = el.closest('.p'); return p ? { p: V.pathById.get(p.dataset.pid), o: 0 } : null; }
  const pEl = el?.closest?.('.p[data-pid]');
  if (!pEl) return null;
  const path = V.pathById.get(pEl.dataset.pid);
  if (!path) return null;
  if (node.nodeType === 1 && node.classList?.contains('p')) {
    const child = node.childNodes[offset];
    if (child?.dataset?.o != null) return { p: path, o: Number(child.dataset.o) };
    const last = node.querySelector(':scope > .end');
    return { p: path, o: Number(last?.dataset.o ?? 0) };
  }
  const sp = el.closest('.r');
  if (!sp) return { p: path, o: 0 };
  const s0 = Number(sp.dataset.o);
  if (!isTextRun(sp)) {
    if (sp.classList.contains('end')) return { p: path, o: s0 };
    return { p: path, o: s0 + (node.nodeType === 3 && offset > 0 ? 1 : node === sp && offset > 0 ? 1 : 0) };
  }
  let k = 0;
  for (const t of textNodes(sp)) { if (t === node) return { p: path, o: s0 + Math.min(offset, t.data.length) + k }; k += t.data.length; }
  return { p: path, o: s0 + (offset > 0 ? sp.textContent.length : 0) };
}
/** 문단 요소가 보이는 상자 (잘린 조각이면 그 창) */
function visibleBox(el) {
  const frag = el.closest('.frag');
  return frag ? frag.getBoundingClientRect() : null;
}
/** 위치 → 화면 사각형 { left, top, height, page } */
export function caretRect(pos) {
  if (!pos) return null;
  const blk = blockAt(S.doc, pos.p);
  if (!blk) return null;
  for (const el of paraEls(blk.id)) {
    const { node, offset } = domPoint(el, pos.o);
    const r = pointRect(node, offset);
    if (!r) continue;
    const box = visibleBox(el);
    const cy = r.top + r.height / 2;
    if (box && (cy < box.top - 1 || cy > box.bottom + 1)) continue;
    return { left: r.left, top: r.top, height: r.height, el };
  }
  return null;
}
function pointRect(node, offset) {
  const rg = document.createRange();
  try { rg.setStart(node, offset); rg.collapse(true); } catch { return null; }
  let r = [...rg.getClientRects()].find((x) => x.height > 0);
  if (r) return r;
  // 빈 사각형: 다음 글자나 앞 글자 사각형으로
  if (node.nodeType === 3) {
    if (offset < node.data.length) { rg.setEnd(node, offset + 1); r = [...rg.getClientRects()][0]; if (r) return { left: r.left, top: r.top, height: r.height }; }
    if (offset > 0) { rg.setStart(node, offset - 1); rg.setEnd(node, offset); r = [...rg.getClientRects()].pop(); if (r) return { left: r.right, top: r.top, height: r.height }; }
  }
  const tgt = node.nodeType === 1 ? (node.childNodes[offset] ?? node.childNodes[offset - 1] ?? node) : node.parentElement;
  if (tgt?.getBoundingClientRect) { const b = tgt.getBoundingClientRect(); const after = node.nodeType === 1 && !node.childNodes[offset]; return { left: after ? b.right : b.left, top: b.top, height: b.height || 16 }; }
  return null;
}
/** 화면 좌표 → 문서 위치 */
export function posFromPoint(x, y) {
  let node = null; let offset = 0;
  if (document.caretPositionFromPoint) { const c = document.caretPositionFromPoint(x, y); if (c) { node = c.offsetNode; offset = c.offset; } }
  else if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(x, y); if (r) { node = r.startContainer; offset = r.startOffset; } }
  const el = node && (node.nodeType === 1 ? node : node.parentElement);
  const region = S.hf ? '.hf, .notes' : '.body, .notes';
  if (el && el.closest(region) && el.closest('.p[data-pid]')) {
    const p = posFromDom(node, offset);
    if (p) return p;
  }
  // 문단 밖 (여백 · 단락 사이): 가장 가까운 줄
  return nearestPos(x, y);
}
function nearestPos(x, y) {
  let best = null; let bd = Infinity;
  for (const r of V.rendered.values()) {
    const pr = r.node.getBoundingClientRect();
    if (y < pr.top - GAP || y > pr.bottom + GAP) continue;
    for (const el of r.node.querySelectorAll(S.hf ? '.hf .p[data-pid]' : '.body .p[data-pid]')) {
      const b = el.getBoundingClientRect();
      const box = visibleBox(el) ?? b;
      const top = Math.max(b.top, box.top); const bot = Math.min(b.bottom, box.bottom);
      if (bot <= top) continue;
      const d = y < top ? top - y : y > bot ? y - bot : 0;
      if (d < bd) { bd = d; best = { el, top, bot, b }; }
    }
  }
  if (!best) return null;
  const yy = Math.max(best.top + 2, Math.min(best.bot - 2, y));
  const xx = Math.max(best.b.left + 1, Math.min(best.b.right - 1, x));
  let node = null; let offset = 0;
  if (document.caretPositionFromPoint) { const c = document.caretPositionFromPoint(xx, yy); if (c) { node = c.offsetNode; offset = c.offset; } }
  else if (document.caretRangeFromPoint) { const r = document.caretRangeFromPoint(xx, yy); if (r) { node = r.startContainer; offset = r.startOffset; } }
  if (node && best.el.contains(node)) return posFromDom(node, offset);
  return { p: V.pathById.get(best.el.dataset.pid), o: 0 };
}

/** 쪽 번호 (캐럿이 있는 쪽) */
export function caretPage() {
  const res = S.layout;
  if (!res || !S.sel) return 0;
  const top = S.sel.f.p[0];
  const blk = S.doc.body[top];
  const items = blk ? res.byId.get(blk.id) : null;
  if (!items?.length) return 0;
  if (items.length === 1 || blk.type !== 'p') return items[0].page;
  // 나뉜 문단: 위치의 y 로 조각 찾기 (측정 요소 기준)
  const el = items[0].el;
  const { node, offset } = domPoint(el, S.sel.f.o);
  const r = pointRect(node, offset);
  if (!r) return items[0].page;
  const y = r.top - el.getBoundingClientRect().top;
  for (const it of items) if (it.clip && y >= it.clip.from - 1 && y < it.clip.to) return it.page;
  return items[items.length - 1].page;
}

// ───────────── 캐럿 · 선택 그리기 ─────────────
function overlayOf(el) { return el.closest('.page')?.querySelector(':scope > .overlay'); }
function toLocal(pageEl, x, y) { const r = pageEl.getBoundingClientRect(); return { x: (x - r.left) / S.zoom, y: (y - r.top) / S.zoom }; }
export function drawSelection() {
  for (const r of V.rendered.values()) { const ov = r.node.querySelector(':scope > .overlay'); if (ov) ov.replaceChildren(); }
  if (!S.sel || S.view === 'read' && false) return;
  const { a, f } = S.sel;
  if (S.objSel) { markObject(); }
  if (posEq(a, f)) { drawCaret(f); return; }
  hideCaret();
  const [s, e] = orderPos(a, f);
  const sameDepth = s.p.length === e.p.length;
  const si = V.orderById.get(blockAt(S.doc, s.p)?.id) ?? 0;
  const ei = V.orderById.get(blockAt(S.doc, e.p)?.id) ?? 0;
  void sameDepth;
  const cellMode = cellSelection();
  for (const r of V.rendered.values()) {
    const ov = r.node.querySelector(':scope > .overlay');
    for (const el of r.node.querySelectorAll('.body .p[data-pid], .hf .p[data-pid], .notes .p[data-pid]')) {
      if (el.closest('.hf') && !S.hf) continue;
      if (S.hf && !el.closest('.hf')) continue;
      if (S.hf && S.hf.page != null && Number(r.node.dataset.page) !== S.hf.page && el.closest('.hf')) continue;
      const k = V.orderById.get(el.dataset.pid);
      if (k == null || k < si || k > ei) continue;
      if (cellMode && !cellMode.has(el.dataset.pid)) continue;
      const path = V.pathById.get(el.dataset.pid);
      const blk = blockAt(S.doc, path);
      const from = k === si && !cellMode ? s.o : 0;
      const to = k === ei && !cellMode ? e.o : paraLen(blk);
      const p1 = domPoint(el, from);
      const p2 = domPoint(el, to);
      const rg = document.createRange();
      try { rg.setStart(p1.node, p1.offset); rg.setEnd(p2.node, p2.offset); } catch { continue; }
      const box = visibleBox(el);
      const rects = [...rg.getClientRects()];
      // 문단 끝까지 선택: 문단 표시 자리(작은 폭)도 칠함
      if (to === paraLen(blk) && k < ei) { const end = el.querySelector(':scope > .end'); if (end) { const b = end.getBoundingClientRect(); rects.push({ left: b.left, top: b.top, right: b.left + 6, bottom: b.bottom, width: 6, height: b.height }); } }
      for (const q of mergeRects(rects)) {
        if (box && (q.bottom < box.top + 1 || q.top > box.bottom - 1)) continue;
        const lt = toLocal(r.node, q.left, q.top);
        const d = document.createElement('div');
        d.className = 'selrect';
        d.style.cssText = `left:${f2(lt.x)}px;top:${f2(lt.y)}px;width:${f2(Math.max(2, q.width / S.zoom))}px;height:${f2(q.height / S.zoom)}px`;
        ov.append(d);
      }
    }
  }
}
function mergeRects(rects) {
  const out = [];
  for (const r of rects) {
    if (r.width <= 0 && r.height <= 0) continue;
    const last = out[out.length - 1];
    if (last && Math.abs(last.top - r.top) < 2 && Math.abs(last.bottom - r.bottom) < 2 && r.left <= last.right + 1) { last.right = Math.max(last.right, r.right); last.width = last.right - last.left; continue; }
    out.push({ left: r.left, top: r.top, right: r.right, bottom: r.bottom, width: r.width, height: r.height });
  }
  return out;
}
/** 표에서 칸을 걸쳐 선택 → 선택한 칸의 문단 id 집합 */
export function cellSelection() {
  const { a, f } = S.sel ?? {};
  if (!a || a.p.length < 4 || f.p.length < 4) return null;
  if (a.p[0] !== f.p[0] || (a.p[1] === f.p[1] && a.p[2] === f.p[2])) return null;
  const t = S.doc.body[a.p[0]];
  if (t?.type !== 'tbl') return null;
  const r1 = Math.min(a.p[1], f.p[1]); const r2 = Math.max(a.p[1], f.p[1]);
  const c1 = Math.min(a.p[2], f.p[2]); const c2 = Math.max(a.p[2], f.p[2]);
  const ids = new Set();
  for (let r = r1; r <= r2; r++) for (let c = c1; c <= Math.min(c2, t.rows[r].cells.length - 1); c++) walkParas(t.rows[r].cells[c].blocks, (p) => ids.add(p.id));
  ids.r1 = r1; ids.r2 = r2; ids.c1 = c1; ids.c2 = c2; ids.t = a.p[0];
  return ids;
}
function hideCaret() { if (V.caretEl) V.caretEl.remove(); }
export function drawCaret(pos) {
  const cr = caretRect(pos);
  hideCaret();
  if (!cr) return;
  const pageEl = cr.el.closest('.page');
  const ov = overlayOf(cr.el);
  if (!ov) return;
  const lt = toLocal(pageEl, cr.left, cr.top);
  const c = V.caretEl ?? document.createElement('div');
  c.className = 'wcaret';
  c.style.cssText = `left:${f2(lt.x - 0.5)}px;top:${f2(lt.y)}px;height:${f2(cr.height / S.zoom)}px`;
  ov.append(c);
  V.caretEl = c;
  c.style.animation = 'none';
  void c.offsetWidth;
  c.style.animation = '';
  emit('caret', cr);
}
function markObject() {
  for (const r of V.rendered.values()) for (const n of r.node.querySelectorAll('.obj.sel')) n.classList.remove('sel');
  const o = S.objSel;
  if (!o) return;
  const blk = blockAt(S.doc, o.p);
  for (const el of paraEls(blk?.id)) {
    const sp = el.querySelector(`:scope > .r.obj[data-o="${o.o}"]`);
    if (!sp) continue;
    sp.classList.add('sel');
    const ov = overlayOf(sp);
    const b = (sp.querySelector('img, .svgobj, span') ?? sp).getBoundingClientRect();
    const lt = toLocal(sp.closest('.page'), b.left, b.top);
    const box = document.createElement('div');
    box.className = 'objbox';
    box.style.cssText = `left:${f2(lt.x)}px;top:${f2(lt.y)}px;width:${f2(b.width / S.zoom)}px;height:${f2(b.height / S.zoom)}px`;
    for (const h of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) { const d = document.createElement('i'); d.className = `hd hd-${h}`; d.dataset.h = h; box.append(d); }
    ov.append(box);
  }
}

/** 캐럿이 보이게 스크롤 */
export function scrollToCaret() {
  const pg = caretPage();
  const h = V.holders[pg];
  if (!h) return;
  if (!V.rendered.has(pg)) {
    // 그리지 않은 쪽: 쪽 위치로 먼저 이동
    const t = h.offsetTop;
    if (t < V.scroller.scrollTop || t > V.scroller.scrollTop + V.scroller.clientHeight) V.scroller.scrollTop = t - 40;
    renderVisible();
  }
  const cr = caretRect(S.sel?.f);
  if (!cr) return;
  const sr = V.scroller.getBoundingClientRect();
  if (cr.top < sr.top + 20) V.scroller.scrollTop -= sr.top + 20 - cr.top;
  else if (cr.top + cr.height > sr.bottom - 20) V.scroller.scrollTop += cr.top + cr.height - (sr.bottom - 20);
  if (cr.left < sr.left + 10) V.scroller.scrollLeft -= sr.left + 10 - cr.left;
  else if (cr.left > sr.right - 20) V.scroller.scrollLeft += cr.left - (sr.right - 20);
}

/** 위 · 아래 줄의 같은 x 위치 */
export function verticalPos(pos, dir, x = null) {
  const cr = caretRect(pos);
  if (!cr) return null;
  const xx = x ?? cr.left;
  // 같은 문단 안 다음 줄: 줄 상자로 계산
  let y = dir > 0 ? cr.top + cr.height + 3 : cr.top - 3;
  for (let tries = 0; tries < 40; tries++) {
    const el = document.elementFromPoint(xx, y);
    if (el && el.closest('.page')) {
      const p = posFromPoint(xx, y);
      if (p && !posEq(p, pos) && (dir > 0 ? cmpPos(p, pos) > 0 || p.p.length !== pos.p.length : cmpPos(p, pos) < 0 || p.p.length !== pos.p.length)) return p;
    }
    y += dir * 6;
    // 화면 밖이면 스크롤
    const sr = V.scroller.getBoundingClientRect();
    if (y < sr.top || y > sr.bottom) { V.scroller.scrollTop += dir * 60; renderVisible(); y -= dir * 60; }
  }
  return null;
}
/** 줄 처음 · 끝 (Home / End) */
export function lineEdge(pos, dir) {
  const blk = blockAt(S.doc, pos.p);
  const cr = caretRect(pos);
  if (!cr || !blk) return pos;
  const el = cr.el;
  const { lines } = lineBoxes(el);
  const base = el.getBoundingClientRect().top;
  const cy = cr.top - base + cr.height / 2;
  const line = lines.find((l) => cy >= l.top - 1 && cy <= l.bottom + 1) ?? lines[0];
  if (!line) return pos;
  // 줄 안에서 이분 탐색: 위치 o 의 y 가 이 줄인 가장 작은/큰 o
  const len = paraLen(blk);
  const yOf = (o) => { const { node, offset } = domPoint(el, o); const r = pointRect(node, offset); return r ? r.top - base + r.height / 2 : -1; };
  const onLine = (o) => { const y = yOf(o); return y >= line.top - 1 && y <= line.bottom + 1; };
  let lo = 0; let hi = len;
  if (dir < 0) {
    let l = 0; let h = pos.o;
    while (l < h) { const m = (l + h) >> 1; if (onLine(m)) h = m; else l = m + 1; }
    lo = l;
    return { p: pos.p, o: lo };
  }
  let l = pos.o; let h = len;
  while (l < h) { const m = (l + h + 1) >> 1; if (onLine(m)) l = m; else h = m - 1; }
  hi = l;
  // 줄 끝이 공백 뒤 줄 바꿈이면 그 앞 (Word 와 같음)
  return { p: pos.p, o: hi };
}
