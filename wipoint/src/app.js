// WIPOINT 시작점: 화면 조립 · 단축키 · 클립보드 · 끌어 놓기 · 자동 저장
import { S, on, emit, run, register, curSlide, selObjects, selOne, goSlide, COMMANDS } from './state.js';
import { initEditor, renderCanvas, renderOverlay, fitZoom, startEdit, endEdit, applyTextFormat, focusEditing } from './editor.js';
import { initPanels, renderThumbs, updateThumb, markActiveThumb, renderNotes, renderStatus, renderPane, renderSorter, thumbsFocused, markPeerDots } from './panels.js';
import { collab, startCollab, stopCollab, collabLink } from './collab.js';
import { initMobile } from './mobile.js';
import { canvasMenu, thumbMenu } from './ctxmenu.js';
import { initRibbon, renderRibbon, TABS, setRibbonTab, initQat } from './ribbon.js';
import { initKeytips } from './keytips.js';
import { stopEyedrop } from './eyedrop.js';
import { GALLERIES } from './galleries.js';
import { APP_VERSION } from './commands.js';
import { openBackstage, closeBackstage, backstageOpen } from './backstage.js';
import { scheduleAutosave, autosaveNow, restoreAutosave, loadViewLink, loadFile, setDocument } from './fileio.js';
import { startShow, showActive } from './show.js';
import { SLIDE_CSS } from './render.js';
import { setOffsets } from './textedit.js';
import { el, toast, alertDialog, openMenu, hydrateIcons, isDialogOpen, isMenuOpen, closeMenus } from './ui.js';

const $ = (id) => document.getElementById(id);

function boot() {
  document.head.append(el('style', {}, SLIDE_CSS));
  initRibbon($('ribbonTabs'), $('ribbon'), GALLERIES);
  initQat($('qat'), $('qatTop'));
  initKeytips({ isBlocked: () => showActive() || !!S.eyedrop });
  initEditor($('stage'));
  initPanels({ thumbs: $('thumbs'), notes: $('notes'), status: $('status'), pane: $('pane'), sorter: $('sorter') });
  // 모바일 화면 + 끌어서 스크롤 (리본 탭 · 리본 · 빠른 실행 · 상태 표시줄 · 슬라이드 목록 · 작업 창)
  initMobile([[$('ribbonTabs'), 'x'], [$('ribbon'), 'x'], [$('qat'), 'x'], [$('status'), 'x', { mobileOnly: true }], [$('thumbs'), 'auto', { mobileOnly: true }], [$('pane'), 'y', { mobileOnly: true }], [$('sorter'), 'y', { mobileOnly: true }]]);
  addEventListener('wipoint-mobile', () => { applyView(); renderAll(); });
  register({ backstage: (page) => openBackstage(page ?? 'home') });
  wireTitlebar();
  wireEvents();
  applyView();
  renderAll();
  start();
}

async function start() {
  try {
    if (location.hash.startsWith('#view=')) {
      if (await loadViewLink(location.hash)) {
        document.body.classList.add('view-only');
        toast('읽기 전용으로 열었습니다');
        startShow({ from: 0, windowed: true });
        return;
      }
    }
  } catch (e) { alertDialog('공유 링크', `링크를 열지 못했습니다: ${e.message}`); }
  // ?collab=<방>: 공동 편집 방에 들어가서 문서를 받아 옴
  const room = new URLSearchParams(location.search).get('collab');
  if (room && /^[\w-]{4,64}$/.test(room)) {
    await startCollab(room, { join: true });
    toast('공동 편집 방에 들어왔습니다 — 문서를 받는 중…');
    return;
  }
  const restored = await restoreAutosave();
  if (restored) toast('이전에 작업하던 프레젠테이션을 복원했습니다');
  else if (!localStorage.getItem('wipoint:seen')) setTimeout(() => openBackstage('home'), 50);
  try {
    if (localStorage.getItem('wipoint:version') !== APP_VERSION) { localStorage.setItem('wipoint:version', APP_VERSION); localStorage.setItem('wipoint:seen', '1'); }
  } catch { /* 저장소 없음 */ }
}

// ───────────── 다시 그리기 ─────────────
let thumbTimer = 0;
let ribbonTimer = 0;
function renderAll() {
  renderThumbs();
  renderCanvas();
  renderNotes();
  renderStatus();
  renderRibbon();
  renderPane();
  if (S.view === 'sorter') renderSorter();
  updateTitle();
}
function renderRibbonSoon() { clearTimeout(ribbonTimer); ribbonTimer = setTimeout(renderRibbon, 120); }

on('change', ({ scope = 'slide', keepPick } = {}) => {
  if (scope === 'all') renderAll();
  else if (scope === 'slide') { updateThumb(); if (!S.editing) renderCanvas(); else renderOverlay(); renderRibbonSoon(); if (S.formatPane) renderPane(); renderStatus(); if (S.view === 'sorter') renderSorter(); }
  else if (scope === 'text') { clearTimeout(thumbTimer); thumbTimer = setTimeout(() => updateThumb(), 250); renderOverlay(); renderRibbonSoon(); }
  else if (scope === 'nav') {
    if (!keepPick) S.slideSel = new Set([curSlide()?.id]);
    if (S.pres.slides.length !== document.querySelectorAll('#thumbs .thumb').length) renderThumbs(); else markActiveThumb();
    renderCanvas(); renderNotes(); renderStatus(); renderRibbon(); renderPane();
    if (S.view === 'sorter') renderSorter();
  } else if (scope === 'view') { applyView(); renderStatus(); renderRibbon(); }
  else if (scope === 'none') renderStatus();
  if (scope !== 'nav' && scope !== 'view') { scheduleAutosave(); updateTitle(); }
  if (collab.room) markPeerDots();
});
on('selection', () => {
  renderOverlay();
  renderRibbonSoon();
  if (S.painter && selObjects().length) {
    if (S.painter.anim) {
      const s = curSlide();
      const ids = new Set(selObjects().map((o) => o.id));
      run('removeAnim', { all: true });
      for (const id of ids) for (const a of S.animClip ?? []) { const { id: _i, obj: _o, ...rest } = a; s.anims.push({ ...rest, id: `a${Math.random().toString(36).slice(2)}`, obj: id }); }
      S.painter = null;
      emit('change', { scope: 'slide' });
    } else run('applyPainter');
    emit('painter');
  }
});
on('zoom', () => renderStatus());
on('hydrate', (node) => hydrateIcons(node));
on('error', (e) => { console.error(e); toast(`오류: ${e?.message ?? e}`); });
on('drawMode', () => { $('stage').classList.toggle('drawing', !!S.drawShape); });
on('painter', () => { document.body.classList.toggle('painting', !!S.painter); renderRibbon(); });
on('saved', () => updateTitle());
// 공동 편집: 참가자 표시 (제목 표시줄 동그라미 · 축소판 점 · 개체 테두리)
on('collabPeers', () => {
  renderOverlay();
  markPeerDots();
  const bar = $('collabBar');
  bar.hidden = !collab.room;
  if (!collab.room) { bar.innerHTML = ''; return; }
  bar.replaceChildren(
    ...[...collab.peers.values()].map((p) => el('span', { class: 'peer-av', title: p.name, style: { background: p.color } }, [...p.name][0] ?? '?')),
    el('button', { class: 'collab-chip', title: `공동 편집 중 (${collab.mode === 'server' ? '서버 연결' : '이 브라우저의 탭끼리'}) — 누르면 링크 · 끝내기`, onclick: () => run('collabStart') }, `공동 편집 ${collab.peers.size + 1}명`));
});
on('collabStatus', (t) => toast(t));
on('docLoaded', () => { fitZoom(); updateTitle(); });
on('showEnded', () => { if (S.viewOnly) toast('읽기 전용 문서입니다. 편집하려면 [파일 › 다른 이름으로 저장]으로 내려받으세요.'); });
on('openPane', () => { renderPane(); fitZoom(); });
on('focusCanvas', () => { if (S.editing) focusEditing(); else $('stage').focus(); });

function applyView() {
  const app = $('app');
  app.dataset.view = S.view;
  $('sorter').hidden = S.view !== 'sorter';
  $('stage').hidden = S.view === 'sorter';
  app.classList.toggle('no-notes', !S.showNotes || S.view === 'sorter');
  app.classList.toggle('big-notes', !!S.bigNotes);
  renderThumbs();
  if (S.view === 'sorter') renderSorter();
  else requestAnimationFrame(() => { if (S.fitZoom) fitZoom(); renderCanvas(); });
}

function updateTitle() {
  const t = `${S.docName}${S.dirty ? ' •' : ''} - WIPOINT`;
  document.title = t;
  const d = $('docTitle');
  if (d && !d.querySelector('input')) d.textContent = `${S.docName} - WIPOINT`;
  const st = $('saveState');
  if (st) st.textContent = S.fileHandle ? (S.dirty ? '저장 안 됨' : '저장됨') : '이 브라우저에 자동 보관';
}

// ───────────── 제목 표시줄 · 검색 ─────────────
function wireTitlebar() {
  for (const b of document.querySelectorAll('[data-cmd]')) b.addEventListener('click', () => run(b.dataset.cmd));
  hydrateIcons(document);
  const d = $('docTitle');
  d.addEventListener('click', () => {
    if (d.querySelector('input')) return;
    const inp = el('input', { type: 'text', value: S.docName });
    d.replaceChildren(inp);
    inp.focus();
    inp.select();
    const done = () => { const v = inp.value.trim(); if (v) S.docName = v; d.textContent = ''; updateTitle(); scheduleAutosave(); };
    inp.addEventListener('blur', done);
    inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') inp.blur(); if (e.key === 'Escape') { inp.value = S.docName; inp.blur(); } });
  });
  // 검색 (명령 찾기): 리본의 모든 단추 이름
  const search = $('searchBox');
  const index = [];
  const walk = (it, tab) => {
    if (!it) return;
    if (it.items) for (const x of it.items) walk(x, tab);
    if (Array.isArray(it.rows)) for (const r of it.rows) for (const x of r) walk(x, tab);
    if ((it.label || it.title) && (it.cmd || it.menu)) index.push({ label: it.label ?? it.title, cmd: it.cmd, menu: it.menu, tab });
  };
  for (const t of TABS) for (const g of t.groups) for (const it of g.items) walk(it, t);
  search.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Escape') { search.value = ''; search.blur(); closeMenus(); return; }
    if (e.key !== 'Enter' && e.key !== 'ArrowDown') return;
    e.preventDefault();
    const q = search.value.trim().replace(/\s+/g, '');
    if (!q) return;
    const seen = new Set();
    const hits = index.filter((x) => x.label.replace(/\s+/g, '').includes(q) && !seen.has(x.label + x.cmd) && seen.add(x.label + x.cmd)).slice(0, 12);
    if (!hits.length) { toast('찾는 명령이 없습니다'); return; }
    openMenu(search, hits.map((h) => ({ label: `${h.label}  ·  ${h.tab.label}`, action: () => { search.blur(); if (h.cmd) run(h.cmd); else { setRibbonTab(h.tab.id); toast(`[${h.tab.label}] 탭에 있습니다`); } } })));
  });
}

// ───────────── 이벤트 ─────────────
function inTextInput(t) { return t && (t.tagName === 'INPUT' || t.tagName === 'TEXTAREA' || t.tagName === 'SELECT' || (t.isContentEditable && !t.closest('.slide-layer'))); }

function wireEvents() {
  document.addEventListener('keydown', onKey);
  on('editKey', (e) => onEditCtrl(e));
  $('thumbs').addEventListener('focus', () => { S.focusThumbs = true; });
  $('thumbs').addEventListener('blur', () => { S.focusThumbs = false; });
  // 클립보드
  document.addEventListener('copy', (e) => {
    if (inTextInput(e.target) || S.editing || showActive()) return;
    e.preventDefault();
    run('copy');
  });
  document.addEventListener('cut', (e) => {
    if (inTextInput(e.target) || S.editing || showActive()) return;
    e.preventDefault();
    run('cut');
  });
  document.addEventListener('paste', (e) => {
    if (inTextInput(e.target) || S.editing || showActive() || isDialogOpen()) return;
    e.preventDefault();
    const files = [...(e.clipboardData?.files ?? [])].filter((f) => f.type.startsWith('image/'));
    const text = e.clipboardData?.getData('text/plain') ?? '';
    const internalFresh = S.clipboard && (!text.trim() || selectedText() === text || Date.now() - S.clipboard.at < 60000 && clipText() === text);
    if (files.length && !internalFresh) { run('pasteFiles', files); return; }
    if (S.clipboard && (internalFresh || !text)) { run('pasteInternal'); return; }
    if (text) run('pasteText', text);
  });
  // 파일 끌어 놓기
  addEventListener('dragover', (e) => { if (e.dataTransfer?.types?.includes('Files')) { e.preventDefault(); document.body.classList.add('dropping'); } });
  addEventListener('dragleave', (e) => { if (!e.relatedTarget) document.body.classList.remove('dropping'); });
  addEventListener('drop', async (e) => {
    document.body.classList.remove('dropping');
    const files = [...(e.dataTransfer?.files ?? [])];
    if (!files.length) return;
    e.preventDefault();
    const deck = files.find((f) => /\.(pptx|ppsx|potx|json|ppt|pdf)$/i.test(f.name));
    if (deck) {
      try { const r = await loadFile(deck); setDocument(r.pres, r.name); if (r.pdf) { toast(r.warnings[0]); if (r.warnings.length > 1) alertDialog('일부 내용', r.warnings.slice(1).join('\n')); } else if (r.warnings.length) alertDialog('일부 내용', r.warnings.join('\n')); else toast('프레젠테이션을 열었습니다'); } catch (err) { alertDialog('열기', err.message); }
      return;
    }
    run('pasteFiles', files);
  });
  addEventListener('beforeunload', (e) => { if (S.dirty && !S.viewOnly) { autosaveNow(); if (S.fileHandle) { e.preventDefault(); e.returnValue = ''; } } });
  addEventListener('error', (e) => emit('error', e.error ?? e.message));
  addEventListener('unhandledrejection', (e) => emit('error', e.reason));
  // 편집 화면 오른쪽 클릭
  // 오른쪽 클릭 메뉴 (PowerPoint 와 같은 구성 — ctxmenu.js)
  on('canvasMenu', ({ e, hit }) => canvasMenu(e, hit));
  on('thumbMenu', ({ e, index }) => thumbMenu(e, index));
}
function selectedText() { return selObjects().map((o) => (o.text ? o.text.paras.map((p) => p.runs.map((r) => r.t).join('')).join('\n') : '')).join('\n'); }
function clipText() { return (S.clipboard?.objs ?? []).map((o) => (o.text ? o.text.paras.map((p) => p.runs.map((r) => r.t).join('')).join('\n') : '')).filter(Boolean).join('\n') || ' '; }

/** 글 편집 중 Ctrl 조합 · 기능 키 */
function onEditCtrl(e) {
  const k = e.key.toLowerCase();
  const shift = e.shiftKey;
  if (fnKey(e)) return;
  if (e.altKey && shift && e.key.startsWith('Arrow')) {
    e.preventDefault();
    if (e.key === 'ArrowLeft') run('indentLess'); else if (e.key === 'ArrowRight') run('indentMore'); else run('moveParagraph', e.key === 'ArrowUp' ? -1 : 1);
    return;
  }
  if (k === 'enter') { e.preventDefault(); run('nextPlaceholder'); return; }
  if (shift && (k === 'f' || k === 'p')) { e.preventDefault(); run('fontDialog'); return; }
  const map = {
    t: () => run('fontDialog'), d: () => { endEdit(); run('duplicate'); }, n: () => run('newBlank'), o: () => run('open'), a: () => document.execCommand('selectAll'),
    b: () => run('bold'), i: () => run('italic'), u: () => run('underline'),
    e: () => run('alignCenter'), l: () => run('alignLeft'), r: () => run('alignRight'), j: () => run('alignJustify'),
    z: () => run('undo'), y: () => run('redo'), s: () => run('save'), k: () => run('hyperlink'), f: () => run('find'), h: () => run('replace'),
    ' ': () => run('clearFormat'), ']': () => run('growFont'), '[': () => run('shrinkFont'), '>': () => run('growFont'), '<': () => run('shrinkFont'), '.': () => (shift ? run('growFont') : null), ',': () => (shift ? run('shrinkFont') : null),
    '=': () => run(shift ? 'superscript' : 'subscript'), '+': () => run('superscript'), m: () => { endEdit(); run('newSlide'); }, p: () => run('print'),
  };
  if (map[k]) { e.preventDefault(); map[k](); }
}

/** 기능 키 (편집 중에도): PowerPoint 와 같은 동작 */
function fnKey(e) {
  const k = e.key;
  // Alt+= : 수식 삽입 (PowerPoint 와 같음)
  if (e.altKey && !e.ctrlKey && !e.metaKey && e.code === 'Equal') { e.preventDefault(); run('insertEquation'); return true; }
  if (!/^F\d+$/.test(k) && k !== 'ContextMenu') return false;
  const ctrl = e.ctrlKey || e.metaKey;
  const shift = e.shiftKey;
  const alt = e.altKey;
  const act = {
    F1: () => (ctrl ? run('toggleRibbon') : run('shortcuts')),
    F2: () => (ctrl ? run('print') : !S.editing && selObjects().length === 1 ? run('editText') : null),
    F3: () => (shift ? run('cycleCase') : null),
    F4: () => (ctrl ? null : alt ? null : run('repeatLast')),
    F5: () => run(alt ? 'presenterView' : shift ? 'showFromCurrent' : 'showFromStart'),
    F6: () => cyclePanes(shift ? -1 : 1),
    F7: () => run('spellCheck'),
    F9: () => (alt ? run('toggleGuides') : shift ? run('toggleGrid') : null),
    F10: () => (shift ? contextMenuKey() : alt ? run('selectionPane') : null),
    F12: () => (ctrl ? run('open') : run('saveAs', 'pptx')),
    ContextMenu: () => contextMenuKey(),
  }[k];
  if (!act) return false;
  const r = act();
  if (r === null) return false;
  e.preventDefault();
  e.stopPropagation();
  return true;
}
function cyclePanes(d) {
  const panes = [$('ribbon').querySelector('button:not(:disabled)'), $('thumbs'), $('stage'), $('notes'), $('status').querySelector('button')].filter((x) => x && x.offsetParent);
  const cur = panes.findIndex((p) => p.contains(document.activeElement));
  if (S.editing) endEdit();
  panes[(cur + d + panes.length) % panes.length]?.focus();
  return true;
}
function contextMenuKey() {
  const o = selObjects()[0];
  const r = (o && document.querySelector(`#stage .ob[data-id="${o.id}"]`))?.getBoundingClientRect() ?? $('stage').getBoundingClientRect();
  emit('canvasMenu', { e: { clientX: r.left + r.width / 2, clientY: r.top + r.height / 2 }, hit: o ?? null });
  return true;
}

function onKey(e) {
  if (showActive() || backstageOpen()) return;
  if (isDialogOpen()) return;
  if (inTextInput(e.target)) return;
  // 편집 중 키는 editor.js 가 처리 (초점이 리본 등으로 옮겨 갔을 때 Esc 는 편집 끝내기)
  if (S.editing) { if (e.key === 'Escape') { e.preventDefault(); endEdit(); } return; }
  const k = e.key;
  const ctrl = e.ctrlKey || e.metaKey;
  const shift = e.shiftKey;
  const stop = () => { e.preventDefault(); e.stopPropagation(); };
  if (fnKey(e)) return;
  if (ctrl && e.altKey && !shift && k.toLowerCase() === 'v') { stop(); run('pasteMenu', document.querySelector('#ribbon .rbtn.large') ?? { x: 200, y: 140 }); return; }
  if (ctrl && shift && k === 'Tab') { stop(); if (thumbsFocused()) $('stage').focus(); else $('thumbs').focus(); return; }
  if (ctrl && shift && !e.altKey && (k.toLowerCase() === 'f' || k.toLowerCase() === 'p')) { stop(); run('fontDialog'); return; }
  if (ctrl && !e.altKey && k === 'Enter') { stop(); run('nextPlaceholder'); return; }
  if (ctrl && e.altKey && (k === 'ArrowLeft' || k === 'ArrowRight') && selObjects().length) { stop(); run('rotate', k === 'ArrowLeft' ? -1 : 1); return; }
  if (ctrl && shift && k.startsWith('Arrow') && selObjects().length) { stop(); const d = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, 1], ArrowDown: [0, -1] }[k]; run('resizeSel', d[0], d[1]); return; }
  if (ctrl && !e.altKey) {
    const key = k.toLowerCase();
    const map = {
      n: () => run('newBlank'), o: () => run('open'), s: () => run('save'), p: () => run('print'), z: () => run('undo'), y: () => run('redo'),
      d: () => run('duplicate'), m: () => run('newSlide'), a: () => run('selectAll'), g: () => run(shift ? 'ungroup' : 'group'),
      b: () => run('bold'), i: () => run('italic'), u: () => run('underline'), e: () => run('alignCenter'), l: () => run('alignLeft'), r: () => run('alignRight'), j: () => run('alignJustify'),
      k: () => run('hyperlink'), f: () => run('find'), h: () => run('replace'), ' ': () => run('clearFormat'),
      ']': () => run(shift ? 'bringToFront' : 'growFont'), '[': () => run(shift ? 'sendToBack' : 'shrinkFont'), '}': () => run('bringToFront'), '{': () => run('sendToBack'),
      '>': () => run('growFont'), '<': () => run('shrinkFont'), '.': () => shift && run('growFont'), ',': () => shift && run('shrinkFont'),
      c: () => (shift ? run('copyFormat') : null), v: () => (shift ? run('pasteFormat') : null),
      t: () => run('fontDialog'),
      '=': () => (shift ? run('superscript') : selObjects().some((o) => o.text) ? run('subscript') : run('zoomIn')), '+': () => (selObjects().some((o) => o.text) ? run('superscript') : run('zoomIn')), '-': () => run('zoomOut'), '0': () => fitZoom(),
    };
    if (key in map) {
      const r = (key === 'c' || key === 'v') && !shift ? null : map[key];
      if (r) { stop(); r(); }
    }
    if (k.startsWith('Arrow') && selObjects().length) { stop(); const d = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[k]; run('nudge', d[0], d[1]); }
    return;
  }
  if (thumbsFocused()) return;
  const objs = selObjects();
  if (k === 'Escape' || (k === 'Enter' && S.cropping)) {
    if (S.cropping) { stop(); run('cropPicture'); return; }
    if (S.eyedrop) { stopEyedrop(); return; }
    if (S.drawShape) { S.drawShape = null; emit('drawMode'); return; }
    if (S.painter) { S.painter = null; emit('painter'); return; }
    if (isMenuOpen()) { closeMenus(); return; }
    if (S.formatPane && !objs.length) { S.formatPane = null; renderPane(); fitZoom(); return; }
    S.sel.clear(); emit('selection'); return;
  }
  if (k === 'Delete' || k === 'Backspace') { if (objs.length) { stop(); run('deleteSelection'); } return; }
  if (k.startsWith('Arrow') && objs.length && e.altKey && shift) { stop(); if (k === 'ArrowLeft') run('indentLess'); else if (k === 'ArrowRight') run('indentMore'); return; }
  if (k.startsWith('Arrow') && objs.length && e.altKey && !shift) { stop(); if (k === 'ArrowLeft' || k === 'ArrowRight') run('rotate', k === 'ArrowLeft' ? -15 : 15); return; }
  if (k.startsWith('Arrow') && objs.length && shift) { stop(); const st = 7.5; const d = { ArrowLeft: [-st, 0], ArrowRight: [st, 0], ArrowUp: [0, st], ArrowDown: [0, -st] }[k]; run('resizeSel', d[0], d[1]); return; }
  if (k.startsWith('Arrow')) {
    if (objs.length) { stop(); const step = 7.5; const d = { ArrowLeft: [-step, 0], ArrowRight: [step, 0], ArrowUp: [0, -step], ArrowDown: [0, step] }[k]; run('nudge', d[0], d[1]); }
    else if (k === 'ArrowDown' || k === 'ArrowRight') { stop(); goSlide(S.cur + 1); } else { stop(); goSlide(S.cur - 1); }
    return;
  }
  if (k === 'PageDown') { stop(); goSlide(S.cur + 1); return; }
  if (k === 'PageUp') { stop(); goSlide(S.cur - 1); return; }
  if (k === 'Home' && !objs.length) { stop(); goSlide(0); return; }
  if (k === 'End' && !objs.length) { stop(); goSlide(S.pres.slides.length - 1); return; }
  if (k === 'Tab') {
    stop();
    const list = curSlide().objects.filter((o) => !o.hidden);
    if (!list.length) return;
    const i = objs.length ? list.indexOf(objs[objs.length - 1]) : -1;
    const n = list[(i + (shift ? -1 : 1) + list.length) % list.length];
    S.sel = new Set([n.id]);
    emit('selection');
    return;
  }
  if ((k === 'F2' || k === 'Enter') && objs.length === 1) { stop(); run('editText'); return; }
  // 개체를 고른 채 글자를 치면 그 개체의 글을 바꿈 (PowerPoint 와 같음)
  const one = selOne();
  if (one && one.type === 'shape' && !e.altKey && k.length === 1 && k !== ' ') {
    stop();
    startEdit(one, { all: true });
    requestAnimationFrame(() => document.execCommand('insertText', false, k));
  }
}

// 찾기 등에서 편집 중 글자 범위를 고를 때
on('setTextSel', ({ txi, a, b }) => setOffsets(txi, { a, b }));

// 서식 적용 도우미 노출 (도구 · 테스트)
window.wipoint = {
  S, run, emit, COMMANDS, applyTextFormat, goSlide, renderAll,
  /** 도구 · 테스트용: 바이트 배열로 파일 열기 */
  openBytes: async (name, bytes) => { const r = await loadFile(new File([new Uint8Array(bytes)], name)); setDocument(r.pres, r.name); return r.warnings; },
};

boot();
