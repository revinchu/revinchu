// 입력: 숨은 입력 상자(IME) · 키보드 · 마우스 · 클립보드 → 문서 편집
import { S, change, emit, run, curSel } from './state.js';
import { V, posFromPoint, drawSelection, scrollToCaret, verticalPos, lineEdge, caretRect, cellSelection, renderVisible } from './view.js';
import {
  blockAt, paraLen, posEq, orderPos, cmpPos, insertText, insertRuns, splitPara, deleteRange, joinWithPrev, wordAt, wordStep, sentenceAt,
  nextParaPath, firstPos, lastPos, propsAt, sliceRange, insertBlocks, paraPlain, walkParas, formatRange, rangeParas, effPara, containerOf, uid, splitRuns, normalizeRuns, runLen,
} from './model.js';
import { htmlToBlocks, blocksToHtml } from './htmlio.js';

export const E = { ta: null, dragging: null, lastClick: 0, clicks: 0, goalX: null };

const author = () => S.opts.userName || '사용자';
const now = () => new Date().toISOString();

export function initEditor(scroller, ta) {
  E.ta = ta;
  scroller.addEventListener('mousedown', onMouseDown);
  window.addEventListener('mousemove', onMouseMove);
  window.addEventListener('mouseup', () => { E.dragging = null; });
  scroller.addEventListener('dblclick', (e) => e.preventDefault());
  ta.addEventListener('keydown', onKeyDown);
  ta.addEventListener('beforeinput', onBeforeInput);
  ta.addEventListener('input', () => { if (!S.compose) ta.value = ''; });
  ta.addEventListener('compositionstart', onCompStart);
  ta.addEventListener('compositionupdate', onCompUpdate);
  ta.addEventListener('compositionend', onCompEnd);
  ta.addEventListener('copy', (e) => { e.preventDefault(); copyTo(e.clipboardData); });
  ta.addEventListener('cut', (e) => { e.preventDefault(); copyTo(e.clipboardData); deleteSelection(); });
  ta.addEventListener('paste', (e) => { e.preventDefault(); pasteFrom(e.clipboardData, { plain: E.pastePlain }); E.pastePlain = false; });
  ta.addEventListener('focus', () => document.body.classList.add('editing'));
  ta.addEventListener('blur', () => document.body.classList.remove('editing'));
}
export function placeTextarea(cr) {
  if (!E.ta || !cr) return;
  E.ta.style.left = `${Math.round(cr.left)}px`;
  E.ta.style.top = `${Math.round(cr.top)}px`;
  E.ta.style.height = `${Math.max(12, Math.round(cr.height))}px`;
}
export function focusEditor() { if (E.ta && document.activeElement !== E.ta) E.ta.focus({ preventScroll: true }); }

// ───────────── 선택 도우미 ─────────────
const sel = () => curSel();
const collapsed = () => posEq(sel().a, sel().f);
export function setSel(a, f = a, { scroll = true, keepGoal = false } = {}) {
  S.sel = { a, f };
  S.objSel = null;
  if (!keepGoal) E.goalX = null;
  S.typing = null;
  drawSelection();
  if (scroll) scrollToCaret();
  emit('selection');
}
const clampPos = (p) => { const b = blockAt(S.doc, p.p); if (!b || b.type !== 'p') return firstPos(S.doc.body); return { p: p.p, o: Math.max(0, Math.min(paraLen(b), p.o)) }; };

// ───────────── 글 넣기 · 지우기 ─────────────
/** 선택한 곳에 글 넣기 (선택은 바꿈) */
export function typeText(text, { key = 'type' } = {}) {
  if (!text) return;
  const props = S.typing;
  change((d) => {
    let pos = collapsed() ? sel().f : deleteSel(d);
    const base = { ...propsAt(blockAt(d.doc, pos.p), pos.o), ...(props ?? {}) };
    if (S.doc.settings.track) { base.ins = { author: author(), date: now() }; delete base.del; }
    else { delete base.ins; delete base.del; }
    pos = insertText(d, pos, text, base);
    return pos;
  }, { key: /\s/.test(text) ? `${key}-sp` : key });
  autoFormat(text);
}
/** 변경 내용 추적: 지울 범위를 '삭제' 표시로 (자기가 넣은 글은 정말 지움) */
function trackDelete(d, a, b) {
  [a, b] = orderPos(a, b);
  for (const path of rangeParas(d.doc, a, b)) {
    const p = d.block(path);
    const s = posEqPath(path, a.p) ? a.o : 0;
    const e = posEqPath(path, b.p) ? b.o : paraLen(p);
    const [x, rest] = splitRuns(p.runs, s);
    const [mid, y] = splitRuns(rest, e - s);
    const kept = [];
    for (const r of mid) {
      if (r.ins && r.ins.author === author()) continue; // 내가 넣은 글은 그냥 지움
      if (r.ins) { const c = { ...r }; delete c.ins; c.del = { author: author(), date: now() }; kept.push(c); continue; }
      kept.push({ ...r, del: r.del ?? { author: author(), date: now() } });
    }
    p.runs = normalizeRuns([...x, ...kept, ...y]);
  }
  return a;
}
const posEqPath = (x, y) => x.length === y.length && x.every((v, i) => v === y[i]);
/** 선택 지우기 (초안 안에서) → 위치 */
export function deleteSel(d) {
  const { a, f } = sel();
  if (posEq(a, f)) return f;
  const cs = cellSelection();
  if (cs) {
    // 칸 선택: 칸 내용만 비움
    const t = d.block([cs.t]);
    void t;
    for (let r = cs.r1; r <= cs.r2; r++) for (let c = cs.c1; c <= cs.c2; c++) {
      const cell = d.cell([cs.t], r, c);
      if (cell) cell.blocks = [{ type: 'p', id: uid('p'), style: cell.blocks[0]?.style ?? 'Normal', pPr: { ...(cell.blocks[0]?.pPr ?? {}) }, runs: [] }];
    }
    return { p: [cs.t, cs.r1, cs.c1, 0], o: 0 };
  }
  if (S.doc.settings.track) return trackDelete(d, a, f);
  return deleteRange(d, a, f);
}
export function deleteSelection() {
  if (collapsed() && !S.objSel) return;
  if (S.objSel) { const o = S.objSel; change((d) => { const p = d.block(o.p); const [x, rest] = splitRuns(p.runs, o.o); p.runs = normalizeRuns([...x, ...splitRuns(rest, 1)[1]]); return { p: o.p, o: o.o }; }); return; }
  change((d) => deleteSel(d));
}
function backspace(word = false) {
  if (S.objSel) { deleteSelection(); return; }
  if (!collapsed()) { deleteSelection(); return; }
  const pos = sel().f;
  const p = blockAt(S.doc, pos.p);
  if (pos.o === 0) {
    // 목록 · 들여쓰기 먼저 해제 (Word)
    const ep = effPara(S.doc, p);
    if (ep.num?.id && p.pPr?.num) { change((d) => { const b = d.block(pos.p); b.pPr = { ...b.pPr }; delete b.pPr.num; if (b.style === 'ListParagraph') b.style = 'Normal'; return pos; }); return; }
    if ((p.pPr?.ind?.first ?? 0) > 0) { change((d) => { const b = d.block(pos.p); b.pPr = { ...b.pPr, ind: { ...b.pPr.ind, first: 0 } }; return pos; }); return; }
    const i = pos.p[pos.p.length - 1];
    if (i === 0) return;
    const prev = containerOf(S.doc, pos.p)[i - 1];
    if (prev.type === 'tbl') { const lp = lastPos(prev.rows[prev.rows.length - 1].cells[prev.rows[prev.rows.length - 1].cells.length - 1].blocks, [...pos.p.slice(0, -1), i - 1, prev.rows.length - 1, prev.rows[prev.rows.length - 1].cells.length - 1]); if (!paraLen(p)) change((d) => { d.list(pos.p).splice(i, 1); return lp; }); else setSel(lp); return; }
    change((d) => joinWithPrev(d, pos) ?? pos, { key: 'del' });
    return;
  }
  const o2 = word ? wordStep(p, pos.o, -1) : prevChar(p, pos.o);
  change((d) => (S.doc.settings.track ? trackDelete(d, { p: pos.p, o: o2 }, pos) : deleteRange(d, { p: pos.p, o: o2 }, pos)), { key: 'del' });
}
function del(word = false) {
  if (S.objSel || !collapsed()) { deleteSelection(); return; }
  const pos = sel().f;
  const p = blockAt(S.doc, pos.p);
  const len = paraLen(p);
  if (pos.o >= len) {
    const lst = containerOf(S.doc, pos.p);
    const i = pos.p[pos.p.length - 1];
    const next = lst[i + 1];
    if (!next) return;
    if (next.type === 'tbl') { if (!len) change((d) => { d.list(pos.p).splice(i, 1); return firstPos(next.rows[0].cells[0].blocks, [...pos.p.slice(0, -1), i, 0, 0]); }); return; }
    change((d) => joinWithPrev(d, { p: [...pos.p.slice(0, -1), i + 1], o: 0 }) ?? pos, { key: 'del' });
    return;
  }
  const o2 = word ? wordStep(p, pos.o, 1) : nextChar(p, pos.o);
  change((d) => (S.doc.settings.track ? trackDelete(d, pos, { p: pos.p, o: o2 }) : deleteRange(d, pos, { p: pos.p, o: o2 })), { key: 'del' });
}
// 서로게이트 쌍 · 결합 문자를 한 글자로
function prevChar(p, o) { const t = paraPlainPos(p); let k = o - 1; if (k > 0 && /[\uDC00-\uDFFF]/.test(t[k]) && /[\uD800-\uDBFF]/.test(t[k - 1])) k--; return Math.max(0, k); }
function nextChar(p, o) { const t = paraPlainPos(p); let k = o + 1; if (/[\uD800-\uDBFF]/.test(t[o]) && /[\uDC00-\uDFFF]/.test(t[o + 1] ?? '')) k++; return Math.min(t.length, k); }
function paraPlainPos(p) { let s = ''; for (const r of p.runs) s += r.obj ? (r.obj === 'mark' ? '' : '￼') : r.text; return s; }

/** Enter */
export function newParagraph() {
  const pos = collapsed() ? sel().f : null;
  const p = pos ? blockAt(S.doc, pos.p) : null;
  // 빈 목록 항목에서 Enter → 목록 끝 (Word)
  if (p && !paraLen(p) && p.pPr?.num) {
    change((d) => { const b = d.block(pos.p); b.pPr = { ...b.pPr }; delete b.pPr.num; if (b.style === 'ListParagraph') b.style = 'Normal'; delete b.pPr.ind; return pos; });
    return;
  }
  change((d) => {
    let at = collapsed() ? sel().f : deleteSel(d);
    at = splitPara(d, at);
    if (S.doc.settings.track) { const nb = d.block(at.p); nb.paraIns = { author: author(), date: now() }; }
    return at;
  }, { key: null });
}
function lineBreak() { change((d) => { const at = collapsed() ? sel().f : deleteSel(d); return insertRuns(d, at, [{ obj: 'br', ...propsAt(blockAt(d.doc, at.p), at.o) }]); }); }
export function insertObject(run) { change((d) => { const at = collapsed() ? sel().f : deleteSel(d); return insertRuns(d, at, [run]); }); }

// ───────────── 자동 서식 (입력하면서) ─────────────
function autoFormat(text) {
  if (S.opts.autoFormat === false) return;
  if (text !== ' ' && text !== '\t') return;
  const pos = sel().f;
  const p = blockAt(S.doc, pos.p);
  if (!p) return;
  const t = paraPlain(p).slice(0, pos.o);
  // 목록: '* ' '- ' '1. ' '가. ' 로 시작 → 글머리 · 번호
  if (pos.o <= 4 && !p.pPr?.num) {
    const m = /^(\*|-|•|>|1[.)]|가[.)]|①|a[.)]|I\.)\s$/.exec(t);
    if (m) {
      const kind = m[1];
      change((d) => {
        const b = d.block(pos.p);
        const [, rest] = splitRuns(b.runs, pos.o);
        b.runs = rest;
        return { p: pos.p, o: 0 };
      }, { key: null });
      if (kind === '>') run('applyStyle', 'Quote');
      else if (['*', '-', '•'].includes(kind)) run('bullets', kind === '-' ? '–' : '●');
      else run('numbering', kind.startsWith('가') ? 'ganada' : kind.startsWith('①') ? 'decimalCircle' : kind.startsWith('a') ? 'lowerLetter' : kind.startsWith('I') ? 'upperRoman' : 'decimal', kind.endsWith(')') ? ')' : '.');
      return;
    }
    // 마크다운 제목: '# ' '## '
    const h = /^(#{1,6})\s$/.exec(t);
    if (h && S.opts.mdAuto !== false) {
      change((d) => { const b = d.block(pos.p); b.runs = splitRuns(b.runs, pos.o)[1]; b.style = `Heading${h[1].length}`; return { p: pos.p, o: 0 }; }, { key: null });
      return;
    }
  }
  // 자동 고침: 낱말 바꾸기 (-- → —, (c) → ©, 따옴표)
  const ac = [[/--$/, '—'], [/\(c\)$/i, '©'], [/\(r\)$/i, '®'], [/\(tm\)$/i, '™'], [/->$/, '→'], [/<-$/, '←'], [/=>$/, '⇒'], [/\.\.\.$/, '…'], [/1\/2$/, '½']];
  const before = t.slice(0, -1);
  for (const [re, rep] of ac) {
    const m = re.exec(before);
    if (m) {
      const s = pos.o - 1 - m[0].length;
      change((d) => { deleteRange(d, { p: pos.p, o: s }, { p: pos.p, o: pos.o - 1 }); insertText(d, { p: pos.p, o: s }, rep, propsAt(blockAt(d.doc, pos.p), s + 1)); return { p: pos.p, o: s + rep.length + 1 }; }, { key: 'type' });
      return;
    }
  }
  // URL → 하이퍼링크
  const um = /(https?:\/\/[^\s]+|www\.[^\s]+)$/.exec(before);
  if (um) {
    const s = pos.o - 1 - um[1].length;
    const url = um[1].startsWith('www.') ? `https://${um[1]}` : um[1];
    change((d) => { formatRange(d, { p: pos.p, o: s }, { p: pos.p, o: pos.o - 1 }, { link: { url } }); return pos; }, { key: 'type' });
  }
  // 문장 첫 글자 대문자 (영어)
  if (S.opts.autoCap !== false && text === ' ') {
    const m2 = /(?:^|[.!?]\s+)([a-z])[a-z]*$/.exec(before);
    if (m2 && /^[a-z]+$/.test(before.slice(before.length - (m2[0].trim().length)).replace(/^[.!?]\s*/, ''))) {
      const word = /([a-z]+)$/.exec(before)[1];
      const s = pos.o - 1 - word.length;
      if (s === 0 || /[.!?]\s+$/.test(before.slice(0, s))) {
        change((d) => { const b = d.block(pos.p); const [x, rest] = splitRuns(b.runs, s); const [mid, y] = splitRuns(rest, 1); b.runs = normalizeRuns([...x, ...mid.map((r) => (r.obj ? r : { ...r, text: r.text.toUpperCase() })), ...y]); return pos; }, { key: 'type' });
      }
    }
  }
}

// ───────────── IME 조합 ─────────────
function onCompStart() {
  const before = S.doc;
  const selBefore = S.sel;
  let start = sel().f;
  if (!collapsed()) start = change((d) => deleteSel(d), { noHistory: true });
  S.compose = { start, len: 0, before, selBefore };
}
function onCompUpdate(e) {
  const c = S.compose;
  if (!c) return;
  const data = e.data ?? '';
  change((d) => {
    if (c.len) deleteRange(d, c.start, { p: c.start.p, o: c.start.o + c.len });
    const base = { ...propsAt(blockAt(d.doc, c.start.p), c.start.o), ...(S.typing ?? {}) };
    delete base.ins; delete base.del;
    if (S.doc.settings.track) base.ins = { author: author(), date: now() };
    const end = data ? insertText(d, c.start, data, base) : c.start;
    c.len = data.length;
    return end;
  }, { noHistory: true });
}
function onCompEnd(e) {
  const c = S.compose;
  if (!c) return;
  S.compose = null;
  const data = e.data ?? '';
  // 마지막 조합 글과 다르면 바꿈 (기록 없이), 조합 전 문서를 실행 취소 한 번으로 기록
  change((d) => {
    if (c.len) deleteRange(d, c.start, { p: c.start.p, o: c.start.o + c.len });
    const base = { ...propsAt(blockAt(d.doc, c.start.p), c.start.o), ...(S.typing ?? {}) };
    delete base.ins; delete base.del;
    if (S.doc.settings.track) base.ins = { author: author(), date: now() };
    return data ? insertText(d, c.start, data, base) : c.start;
  }, { noHistory: true });
  S.history.record(c.before, c.selBefore, 'type');
  E.ta.value = '';
  if (data) autoFormat(data.slice(-1));
}

function onBeforeInput(e) {
  if (S.compose || e.isComposing) return;
  const t = e.inputType;
  if (t === 'insertText' || t === 'insertReplacementText') { e.preventDefault(); if (e.data) typeText(e.data); }
  else if (t === 'insertParagraph') { e.preventDefault(); newParagraph(); }
  else if (t === 'insertLineBreak') { e.preventDefault(); lineBreak(); }
  else if (t === 'deleteContentBackward') { e.preventDefault(); backspace(); }
  else if (t === 'deleteContentForward') { e.preventDefault(); del(); }
  else if (t === 'insertFromPaste' || t === 'insertFromDrop') { /* paste 이벤트에서 */ }
  else if (t === 'historyUndo') { e.preventDefault(); run('undo'); }
  else if (t === 'historyRedo') { e.preventDefault(); run('redo'); }
}

// ───────────── 키보드 ─────────────
export const KEYMAP = []; // [{ test(e) → bool, run }] — app.js 가 단축키를 채움
function move(fn, e) {
  const s = sel();
  let to;
  if (!e.shiftKey && !posEq(s.a, s.f) && (e.key === 'ArrowLeft' || e.key === 'ArrowRight') && !e.ctrlKey && !e.metaKey) {
    const [a, b] = orderPos(s.a, s.f);
    to = e.key === 'ArrowLeft' ? a : b;
  } else to = fn(s.f);
  if (!to) return;
  to = clampPos(to);
  if (e.shiftKey) setSel(s.a, to, { keepGoal: e.key === 'ArrowUp' || e.key === 'ArrowDown' });
  else setSel(to, to, { keepGoal: e.key === 'ArrowUp' || e.key === 'ArrowDown' });
}
function stepChar(pos, dir) {
  const p = blockAt(S.doc, pos.p);
  if (dir < 0) {
    if (pos.o > 0) return { p: pos.p, o: prevChar(p, pos.o) };
    const pp = nextParaPath(S.doc, pos.p, -1);
    return pp ? { p: pp, o: paraLen(blockAt(S.doc, pp)) } : pos;
  }
  if (pos.o < paraLen(p)) return { p: pos.p, o: nextChar(p, pos.o) };
  const np = nextParaPath(S.doc, pos.p, 1);
  return np ? { p: np, o: 0 } : pos;
}
function stepWord(pos, dir) {
  const p = blockAt(S.doc, pos.p);
  if (dir < 0 && pos.o === 0) return stepChar(pos, -1);
  if (dir > 0 && pos.o >= paraLen(p)) return stepChar(pos, 1);
  return { p: pos.p, o: wordStep(p, pos.o, dir) };
}
function paraEdge(pos, dir) {
  const p = blockAt(S.doc, pos.p);
  if (dir < 0) { if (pos.o > 0) return { p: pos.p, o: 0 }; const pp = nextParaPath(S.doc, pos.p, -1); return pp ? { p: pp, o: 0 } : pos; }
  const np = nextParaPath(S.doc, pos.p, 1);
  void p;
  return np ? { p: np, o: 0 } : { p: pos.p, o: paraLen(blockAt(S.doc, pos.p)) };
}
function vertical(pos, dir) {
  const cr = caretRect(pos);
  if (E.goalX == null && cr) E.goalX = cr.left;
  return verticalPos(pos, dir, E.goalX) ?? (dir < 0 ? pos : pos);
}
function page(pos, dir) {
  const cr = caretRect(pos);
  if (!cr) return pos;
  V.scroller.scrollTop += dir * V.scroller.clientHeight * 0.9;
  renderVisible();
  return posFromPoint(E.goalX ?? cr.left, cr.top + cr.height / 2) ?? pos;
}
function onKeyDown(e) {
  if (S.compose || e.isComposing || e.keyCode === 229) return;
  for (const k of KEYMAP) if (k.test(e)) { e.preventDefault(); k.run(e); return; }
  const ctrl = e.ctrlKey || e.metaKey;
  switch (e.key) {
    case 'ArrowLeft': e.preventDefault(); move((p) => (ctrl ? stepWord(p, -1) : stepChar(p, -1)), e); return;
    case 'ArrowRight': e.preventDefault(); move((p) => (ctrl ? stepWord(p, 1) : stepChar(p, 1)), e); return;
    case 'ArrowUp': e.preventDefault(); move((p) => (ctrl ? paraEdge(p, -1) : vertical(p, -1)), e); return;
    case 'ArrowDown': e.preventDefault(); move((p) => (ctrl ? paraEdge(p, 1) : vertical(p, 1)), e); return;
    case 'Home': e.preventDefault(); move((p) => (ctrl ? firstPos(S.doc.body) : lineEdge(p, -1)), e); return;
    case 'End': e.preventDefault(); move((p) => (ctrl ? lastPos(S.doc.body) : lineEdge(p, 1)), e); return;
    case 'PageUp': e.preventDefault(); move((p) => page(p, -1), e); return;
    case 'PageDown': e.preventDefault(); move((p) => page(p, 1), e); return;
    case 'Backspace': e.preventDefault(); backspace(ctrl); return;
    case 'Delete': e.preventDefault(); del(ctrl); return;
    case 'Enter':
      e.preventDefault();
      if (ctrl) run('pageBreak');
      else if (e.shiftKey) lineBreak();
      else newParagraph();
      return;
    case 'Tab': {
      e.preventDefault();
      const pos = sel().f;
      if (pos.p.length >= 4) { run(e.shiftKey ? 'prevCell' : 'nextCell'); return; }
      const p = blockAt(S.doc, pos.p);
      if (p?.pPr?.num && (pos.o === 0 || !collapsed())) { run(e.shiftKey ? 'indentLess' : 'indentMore'); return; }
      if (e.shiftKey) return;
      if (pos.o === 0 && collapsed() && paraLen(p) && S.opts.tabIndent !== false && !p.runs.some((r) => r.obj === 'tab')) { /* Word: 첫 줄 들여쓰기 대신 탭 문자 */ }
      insertObject({ obj: 'tab', ...propsAt(p, pos.o) });
      return;
    }
    case 'Escape':
      if (S.objSel) { S.objSel = null; drawSelection(); }
      emit('escape');
      return;
    default: break;
  }
}

// ───────────── 마우스 ─────────────
function onMouseDown(e) {
  if (e.button !== 0) return;
  const tgt = e.target;
  if (tgt.closest('.objbox .hd')) { startResize(e, tgt.closest('.objbox .hd').dataset.h); return; }
  const pageEl = tgt.closest('.page');
  if (!pageEl) return;
  e.preventDefault();
  focusEditor();
  // 링크 (Ctrl+클릭)
  const lk = tgt.closest('[data-link]');
  if (lk && (e.ctrlKey || e.metaKey || S.opts.linkClick)) { emit('followLink', lk.dataset.link); return; }
  // 머리글 · 바닥글: 위 · 아래 여백을 두 번 클릭 → 편집, 편집 중 본문 클릭 → 끝
  const pgNo = Number(pageEl.dataset.page);
  const pgInfo = S.layout?.pages[pgNo];
  if (pgInfo) {
    const pr = pageEl.getBoundingClientRect();
    const yy = (e.clientY - pr.top) / S.zoom;
    const inTop = yy < (pgInfo.bodyTopEff ?? pgInfo.bodyTop);
    const inBottom = yy > pgInfo.bodyBottom;
    if (!S.hf && (inTop || inBottom) && (e.detail >= 2 || tgt.closest('.hf'))) {
      if (e.detail >= 2) { emit('editHeaderFooter', { kind: inTop ? 'header' : 'footer', page: pgNo, x: e.clientX, y: e.clientY }); return; }
      return;
    }
    if (S.hf && !inTop && !inBottom && !tgt.closest('.hf')) { if (e.detail >= 2 || true) { emit('closeHeaderFooter'); } }
  }
  // 그림
  const obj = tgt.closest('.r.obj');
  if (obj && obj.closest('.body')) {
    const pEl = obj.closest('.p[data-pid]');
    const path = V.pathById.get(pEl.dataset.pid);
    const o = Number(obj.dataset.o);
    S.sel = { a: { p: path, o }, f: { p: path, o: o + 1 } };
    S.objSel = { p: path, o };
    drawSelection();
    emit('selection');
    emit('objectSelected', S.objSel);
    if (e.detail >= 2) emit('objectDblClick', S.objSel);
    return;
  }
  const now2 = Date.now();
  const body = pageEl.querySelector('.body')?.getBoundingClientRect();
  // 왼쪽 여백 (선택 막대): 줄 · 문단 · 전체 선택
  if (body && e.clientX < body.left - 4 && e.clientY > body.top && e.clientY < body.bottom) {
    const p = posFromPoint(body.left + 2, e.clientY);
    if (!p) return;
    if (e.detail >= 3 || e.ctrlKey) { run('selectAll'); return; }
    const blk = blockAt(S.doc, p.p);
    if (e.detail === 2) { setSel({ p: p.p, o: 0 }, { p: p.p, o: paraLen(blk) }, { scroll: false }); return; }
    const a = lineEdge(p, -1); const b = lineEdge(p, 1);
    setSel(a, b, { scroll: false });
    E.dragging = { mode: 'line', anchor: a };
    return;
  }
  const p = posFromPoint(e.clientX, e.clientY);
  if (!p) return;
  const blk = blockAt(S.doc, p.p);
  if (e.detail === 2) {
    const [s, t] = wordAt(blk, p.o);
    setSel({ p: p.p, o: s }, { p: p.p, o: t }, { scroll: false });
    E.dragging = { mode: 'word', anchor: { p: p.p, o: s }, anchorEnd: { p: p.p, o: t } };
    emit('dblWord');
    return;
  }
  if (e.detail >= 3) { setSel({ p: p.p, o: 0 }, { p: p.p, o: paraLen(blk) }, { scroll: false }); return; }
  if (e.ctrlKey && !e.shiftKey) { const [s, t] = sentenceAt(blk, p.o); setSel({ p: p.p, o: s }, { p: p.p, o: t }, { scroll: false }); return; }
  if (e.shiftKey) setSel(sel().a, p, { scroll: false });
  else setSel(p, p, { scroll: false });
  E.dragging = { mode: 'char', anchor: e.shiftKey ? sel().a : p };
  void now2;
  if (S.painter) { E.dragging.painter = true; }
}
function onMouseMove(e) {
  const d = E.dragging;
  if (!d) return;
  if (d.mode === 'resize') { doResize(e); return; }
  if (!(e.buttons & 1)) { E.dragging = null; return; }
  const p = posFromPoint(e.clientX, e.clientY);
  if (!p) return;
  if (d.mode === 'word') {
    const blk = blockAt(S.doc, p.p);
    const [s, t] = wordAt(blk, p.o);
    if (cmpPos(p, d.anchor) < 0) setSel(d.anchorEnd, { p: p.p, o: s }, { scroll: false });
    else setSel(d.anchor, { p: p.p, o: t }, { scroll: false });
  } else setSel(d.anchor, p, { scroll: false });
  // 가장자리에서 자동 스크롤
  const sr = V.scroller.getBoundingClientRect();
  if (e.clientY > sr.bottom - 20) V.scroller.scrollTop += 20;
  else if (e.clientY < sr.top + 20) V.scroller.scrollTop -= 20;
}
window.addEventListener('mouseup', () => {
  if (S.painter && sel() && !posEq(sel().a, sel().f)) run('applyPainter');
});

// ── 그림 크기 바꾸기 ──
function startResize(e, h) {
  e.preventDefault();
  const o = S.objSel;
  if (!o) return;
  const blk = blockAt(S.doc, o.p);
  let pos = 0; let r0 = null;
  for (const r of blk.runs) { if (pos === o.o) { r0 = r; break; } pos += runLen(r); }
  if (!r0) return;
  E.dragging = { mode: 'resize', h, x0: e.clientX, y0: e.clientY, w0: r0.w, h0: r0.h, ratio: r0.w / r0.h, o };
}
function doResize(e) {
  const d = E.dragging;
  let dw = (e.clientX - d.x0) / S.zoom;
  let dh = (e.clientY - d.y0) / S.zoom;
  if (d.h.includes('w')) dw = -dw;
  if (d.h.includes('n')) dh = -dh;
  let w = d.w0 + (d.h === 'n' || d.h === 's' ? 0 : dw);
  let h = d.h0 + (d.h === 'e' || d.h === 'w' ? 0 : dh);
  if (d.h.length === 2 || e.shiftKey) { const s = Math.max(w / d.w0, h / d.h0); w = d.w0 * s; h = d.h0 * s; }
  w = Math.max(8, w); h = Math.max(8, h);
  if (d.raf) return;
  d.raf = requestAnimationFrame(() => {
    d.raf = 0;
    change((dr) => {
      const p = dr.block(d.o.p);
      let pos = 0;
      p.runs = p.runs.map((r) => { const at = pos; pos += runLen(r); return at === d.o.o ? { ...r, w: Math.round(w), h: Math.round(h) } : r; });
      return { a: { p: d.o.p, o: d.o.o }, f: { p: d.o.p, o: d.o.o + 1 } };
    }, { key: 'resize' });
    S.objSel = d.o;
    drawSelection();
  });
}

// ───────────── 클립보드 ─────────────
export function selectedBlocks() {
  const { a, f } = sel();
  if (S.objSel) return sliceRange(S.doc, { p: S.objSel.p, o: S.objSel.o }, { p: S.objSel.p, o: S.objSel.o + 1 });
  if (posEq(a, f)) return [];
  const cs = cellSelection();
  if (cs) {
    const t = S.doc.body[cs.t];
    const rows = t.rows.slice(cs.r1, cs.r2 + 1).map((r) => ({ ...r, cells: r.cells.slice(cs.c1, cs.c2 + 1) }));
    return [{ ...t, rows, grid: t.grid.slice(cs.c1, cs.c2 + 1) }];
  }
  return sliceRange(S.doc, a, f);
}
export function copyTo(dt) {
  const blocks = selectedBlocks();
  if (!blocks.length) return;
  const json = JSON.stringify({ wiword: 1, blocks, media: usedMediaOf(blocks), styles: S.doc.styles, lists: S.doc.lists });
  dt.setData('text/plain', blocks.map((b) => (b.type === 'p' ? paraPlain(b) : b.rows.map((r) => r.cells.map((c) => c.blocks.map(paraPlain).join(' ')).join('\t')).join('\n'))).join('\n'));
  dt.setData('text/html', `<meta charset="utf-8"><!--wiword:${encodeURIComponent(json)}-->${blocksToHtml(S.doc, blocks)}`);
  try { localStorage.setItem('wiword:clip', json); } catch { /* 큰 것 */ }
}
function usedMediaOf(blocks) {
  const out = {};
  const walk = (bs) => walkParas(bs, (p) => { for (const r of p.runs) if (r.obj === 'img' && r.media && S.doc.media[r.media]) out[r.media] = S.doc.media[r.media]; });
  walk(blocks);
  return out;
}
export function pasteFrom(dt, { plain = false } = {}) {
  const html = dt.getData('text/html');
  const text = dt.getData('text/plain');
  const files = [...(dt.files ?? [])];
  if (files.some((f) => f.type.startsWith('image/'))) { emit('pasteImages', files.filter((f) => f.type.startsWith('image/'))); return; }
  if (!plain && html) {
    const m = /<!--wiword:([^>]*)-->/.exec(html);
    if (m) {
      try {
        const j = JSON.parse(decodeURIComponent(m[1]));
        pasteBlocks(j.blocks, j.media, j.styles, j.lists);
        return;
      } catch { /* HTML 로 */ }
    }
    const { blocks, media, lists } = htmlToBlocks(html, S.doc);
    if (blocks.length) { pasteBlocks(blocks, media, null, lists); return; }
  }
  if (text) {
    if (S.opts.mdPaste !== false && looksMarkdown(text)) { emit('pasteMarkdown', text); return; }
    change((d) => {
      const at = collapsed() ? sel().f : deleteSel(d);
      return insertText(d, at, text, { ...propsAt(blockAt(d.doc, at.p), at.o), ...(S.typing ?? {}) });
    });
  }
}
const looksMarkdown = (t) => /^(#{1,6} |[-*+] |\d+\. |> |```|\|.+\|)/m.test(t) && t.split('\n').length > 2;
export function pasteBlocks(blocks, media = {}, styles = null, lists = null) {
  change((d) => {
    if (lists && Object.keys(lists).length) { const l = d.prop('lists'); for (const [k, v] of Object.entries(lists)) if (!l[k]) l[k] = v; }
    if (media && Object.keys(media).length) { const m = d.prop('media'); Object.assign(m, media); }
    if (styles) { const st = d.prop('styles'); for (const [k, v] of Object.entries(styles)) if (!st[k]) st[k] = v; }
    const at = collapsed() ? sel().f : deleteSel(d);
    return insertBlocks(d, at, blocks);
  });
}
