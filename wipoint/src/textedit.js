// 글 편집: contenteditable DOM ⇄ 글 모델, 커서 위치(단락, 글자 위치) 저장 · 복원
import { normalizeRuns } from './model.js';

const SKIP = (n) => n.nodeType === 1 && (n.classList?.contains('bu') || n.classList?.contains('ph-prompt'));
const parse = (s) => { try { return s ? JSON.parse(s) : null; } catch { return null; } };

/** 단락 요소 목록 (.txi 의 자식 블록) */
function paraEls(txi) {
  const out = [];
  let loose = null;
  for (const n of [...txi.childNodes]) {
    if (n.nodeType === 1 && (n.tagName === 'DIV' || n.tagName === 'P')) { out.push(n); loose = null; } else if (n.nodeType === 3 && n.textContent !== '') {
      // 블록 밖 글자 (브라우저가 만든 것) → 단락으로 감쌈
      if (!loose) { loose = document.createElement('div'); loose.className = 'p'; n.before(loose); out.push(loose); }
      loose.append(n);
    }
  }
  return out;
}

/** 글자 노드 목록 (글머리 · 안내 글 제외) */
function textNodes(p) {
  const out = [];
  const walk = (n) => {
    for (const c of n.childNodes) {
      if (SKIP(c)) continue;
      if (c.nodeType === 3) out.push(c);
      else if (c.nodeType === 1) { if (c.tagName === 'BR') out.push(c); else walk(c); }
    }
  };
  walk(p);
  return out;
}

function inlineProps(node, p) {
  const props = {};
  let el = node.parentElement;
  const chain = [];
  while (el && el !== p) { chain.push(el); el = el.parentElement; }
  // 바깥 → 안쪽 순서로 적용
  for (const e of chain.reverse()) {
    const r = parse(e.dataset?.r);
    if (r) Object.assign(props, r);
    const t = e.tagName;
    if (t === 'B' || t === 'STRONG') props.b = true;
    if (t === 'I' || t === 'EM') props.i = true;
    if (t === 'U') props.u = true;
    if (t === 'S' || t === 'STRIKE') props.s = true;
  }
  return props;
}

/** 편집 DOM → 글 모델 (단락 서식은 data-p, 글자 서식은 data-r) */
export function domToBody(txi, base) {
  const paras = [];
  let prev = base?.paras?.[0] ?? { align: 'l', lvl: 0 };
  for (const p of paraEls(txi)) {
    const pp = parse(p.dataset?.p) ?? { ...prev, runs: undefined };
    delete pp.runs;
    const runs = [];
    const nodes = textNodes(p);
    nodes.forEach((n, i) => {
      if (n.nodeType === 1) {
        // 줄 바꿈(Shift+Enter) — 단락 끝의 <br> 은 빈 줄 표시용이라 무시
        if (i === nodes.length - 1) return;
        runs.push({ ...(pp.end ?? {}), ...inlineProps(n, p), t: '\v' });
        return;
      }
      const t = n.textContent.replace(/​/g, '');
      if (!t) return;
      runs.push({ ...(pp.end ?? {}), ...inlineProps(n, p), t });
    });
    for (const r of runs) for (const k of Object.keys(r)) if (r[k] === undefined) delete r[k];
    paras.push({ ...pp, runs });
    prev = pp;
  }
  const body = { ...base, paras: paras.length ? paras : [{ ...(base?.paras?.[0] ?? {}), runs: [] }] };
  normalizeRuns(body);
  return body;
}

// ───────────── 커서 위치 ─────────────
function pointToOffset(txi, node, offset) {
  const ps = paraEls(txi);
  let pi = ps.findIndex((p) => p === node || p.contains(node));
  if (pi < 0) {
    // .txi 자체를 가리키면 자식 순서로
    if (node === txi) return offset >= ps.length ? { p: Math.max(0, ps.length - 1), o: Infinity } : { p: offset, o: 0 };
    return { p: 0, o: 0 };
  }
  const p = ps[pi];
  let o = 0;
  for (const n of textNodes(p)) {
    const len = n.nodeType === 3 ? n.textContent.length : 1;
    if (n === node) return { p: pi, o: o + Math.min(offset, len) };
    // 요소를 가리키는 위치 (자식 번호)
    if (node.nodeType === 1 && node.contains(n)) {
      const idx = [...node.childNodes].findIndex((c) => c === n || c.contains(n));
      if (idx >= offset) return { p: pi, o };
    }
    o += len;
  }
  return { p: pi, o };
}

export function getOffsets(txi) {
  const sel = window.getSelection();
  if (!sel.rangeCount) return null;
  const r = sel.getRangeAt(0);
  if (!txi.contains(r.startContainer)) return null;
  const a = pointToOffset(txi, r.startContainer, r.startOffset);
  const b = pointToOffset(txi, r.endContainer, r.endOffset);
  return { a, b, collapsed: r.collapsed };
}

function offsetToPoint(txi, { p, o }) {
  const ps = paraEls(txi);
  const pe = ps[Math.min(p, ps.length - 1)];
  if (!pe) return [txi, 0];
  let left = o;
  const nodes = textNodes(pe);
  for (const n of nodes) {
    const len = n.nodeType === 3 ? n.textContent.length : 1;
    if (left <= len && n.nodeType === 3) return [n, left];
    left -= len;
  }
  const last = nodes.filter((n) => n.nodeType === 3).pop();
  if (last) return [last, last.textContent.length];
  // 빈 단락: 글머리 다음
  const bu = pe.querySelector('.bu');
  return [pe, bu ? [...pe.childNodes].indexOf(bu) + 1 : 0];
}

export function setOffsets(txi, { a, b }) {
  const sel = window.getSelection();
  const r = document.createRange();
  const [sn, so] = offsetToPoint(txi, a);
  const [en, eo] = offsetToPoint(txi, b ?? a);
  try {
    r.setStart(sn, so);
    r.setEnd(en, eo);
    sel.removeAllRanges();
    sel.addRange(r);
  } catch { /* 위치가 사라짐 */ }
}

/** 커서가 단어 안이면 단어 범위로 넓힘 (PowerPoint: 커서만 두고 굵게 → 그 단어) */
export function wordAt(body, { p, o }) {
  const para = body.paras[p];
  if (!para) return null;
  const t = para.runs.map((r) => r.t).join('');
  const isW = (ch) => /[\p{L}\p{N}_]/u.test(ch ?? '');
  if (!isW(t[o]) || !isW(t[o - 1])) return null;
  let s = o;
  let e = o;
  while (s > 0 && isW(t[s - 1])) s--;
  while (e < t.length && isW(t[e])) e++;
  return { a: { p, o: s }, b: { p, o: e } };
}

export function placeCaretAtPoint(x, y) {
  let r = null;
  if (document.caretRangeFromPoint) r = document.caretRangeFromPoint(x, y);
  else if (document.caretPositionFromPoint) {
    const pos = document.caretPositionFromPoint(x, y);
    if (pos) { r = document.createRange(); r.setStart(pos.offsetNode, pos.offset); r.collapse(true); }
  }
  if (!r) return false;
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
  return true;
}

export function selectAllIn(el) {
  const r = document.createRange();
  r.selectNodeContents(el);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(r);
}
