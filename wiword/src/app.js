// WIWORD 시작점: 화면 조립 · 리본 · 단축키 · 오른쪽 클릭 메뉴 · 미니 도구 모음 · 끌어 놓기 · 자동 저장 · AI 연결
import { S, on, emit, run, register, setDoc, change, curSel, COMMAND_INFO } from './state.js';
import { initView, relayout, drawSelection, scrollToCaret, renderVisible, V, setZoom, applyZoomMode } from './view.js';
import { initEditor, placeTextarea, focusEditor, KEYMAP, setSel } from './editor.js';
import { } from './commands.js';
import { } from './dialogs.js';
import { } from './features.js';
import { } from './pdftoolsui.js';
import { GALLERIES } from './menus.js';
import { initRibbon, renderRibbon, initQat, TABS, setRibbonTab, leafItems } from './ribbon.js';
import { initKeytips } from './keytips.js';
import { initPanels, renderPanels, renderStatus } from './panels.js';
import { scheduleAutosave, restoreAutosave, initDrop, openFile } from './fileio.js';
import { exportPdfBytes } from './pdfexport.js';
import { openBackstage, closeBackstage, backstageOpen } from './backstage.js';
import { loadOpts, APP_VERSION } from './extras.js';
import { bridgeConnect } from './bridge.js';
import { initMobile } from './mobile.js';
import { sampleDoc } from './sample.js';
import { blockAt, posEq, walkParas, runLen, paraPlain } from './model.js';
import { hydrateIcons, toast, el, openMenu, closeMenus, isDialogOpen, isMenuOpen } from './ui.js';
import { restoreUserFonts, detectLocalFonts, ensureDocFonts } from './fontlib.js';

const $ = (id) => document.getElementById(id);
const scroller = $('docView');

function boot() {
  loadOpts();
  initView(scroller);
  initEditor(scroller, $('ime'));
  initRibbon($('ribbonTabs'), $('ribbon'), GALLERIES);
  initQat($('qat'), $('qatTop'));
  initKeytips({ isBlocked: () => false });
  initPanels();
  initDrop(document.body);
  initMobile([[$('ribbonTabs'), 'x'], [$('ribbon'), 'x'], [$('qat'), 'x'], [$('status'), 'x', { mobileOnly: true }], [$('pane'), 'y', { mobileOnly: true }], [$('navPane'), 'y', { mobileOnly: true }]]);
  register({ backstage: (page) => openBackstage(page ?? 'home') }, { backstage: { label: '파일 탭', args: ['home|new|open|info|saveAs|print|share|export'] } });
  hydrateIcons();
  wireTitlebar();
  wireDoc();
  setTimeout(() => { detectLocalFonts(); restoreUserFonts?.().then?.(() => relayout()).catch?.(() => {}); }, 400);
}

// ───────────── 다시 그리기 ─────────────
let ribbonT = 0;
const renderRibbonSoon = () => { clearTimeout(ribbonT); ribbonT = setTimeout(renderRibbon, 90); };
function updateTitle() {
  const t = `${S.docName}${S.dirty ? ' •' : ''} - WIWORD`;
  document.title = t;
  $('docTitle').textContent = t;
}
on('change', (e) => {
  if (e?.newDoc) { S.merge = null; applyViewClass(); }
  relayout();
  if (e?.scope !== 'view' && !e?.noScroll) scrollToCaret();
  updateTitle();
  renderRibbonSoon();
  if (e?.scope !== 'view' && S.opts.autosave !== false) scheduleAutosave();
  if (e?.newDoc) { renderPanels(); ensureDocFonts?.(S.doc, () => relayout()); }
});
on('selection', () => { renderRibbonSoon(); if (S.opts.miniToolbar !== false) maybeMiniToolbar(); });
on('caret', (cr) => placeTextarea(cr));
on('ui', () => { applyViewClass(); renderPanels(); renderRibbon(); relayout(); });
on('pane', () => renderPanels());
on('error', (e) => { console.error(e); toast(`오류: ${e?.message ?? e}`); });
on('toast', (m) => toast(m));
on('focusCanvas', () => focusEditor());
on('saved', () => updateTitle());
on('objectSelected', () => renderRibbon());
on('editHeaderFooter', ({ kind, page }) => run('editHeaderFooterAt', kind, page));
on('closeHeaderFooter', () => run('closeHeaderFooter'));
on('editNote', ({ id, kind }) => { setTimeout(() => { setSel({ p: [`n:${kind}:${id}`, 0], o: 0 }); focusEditor(); }, 30); });
on('escape', () => { if (S.hf) run('closeHeaderFooter'); else if (S.painter) run('formatPainter'); });
on('followLink', (href) => {
  if (!href) return;
  if (href.startsWith('#')) {
    const name = href.slice(1);
    let hit = null;
    walkParas(S.doc.body, (p, path) => { if (hit) return; if (p.id === name) hit = { p: path, o: 0 }; let o = 0; for (const r of p.runs) { if (r.obj === 'mark' && r.kind === 'bm' && r.name === name && r.edge === 'start') { hit = { p: path, o }; break; } o += runLen(r); } });
    if (hit) setSel(hit); else toast(`책갈피 '${name}' 을(를) 찾을 수 없습니다`);
    return;
  }
  if (/^(https?|mailto|tel):/i.test(href)) window.open(href, '_blank', 'noopener');
});
function applyViewClass() {
  $('app').dataset.view = S.view;
  document.body.classList.toggle('read-mode', S.view === 'read');
  document.body.classList.toggle('hf-editing', !!S.hf);
}

// ───────────── 제목 표시줄 · 검색 ─────────────
function wireTitlebar() {
  for (const b of document.querySelectorAll('.titlebar [data-cmd]')) b.addEventListener('click', () => run(b.dataset.cmd));
  $('docTitle').addEventListener('click', () => {
    const inp = el('input', { type: 'text', value: S.docName, class: 'title-edit' });
    $('docTitle').replaceChildren(inp);
    inp.focus(); inp.select();
    const done = (ok) => { if (ok && inp.value.trim()) S.docName = inp.value.trim(); updateTitle(); };
    inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') done(true); if (e.key === 'Escape') done(false); });
    inp.addEventListener('blur', () => done(true));
  });
  const box = $('searchBox');
  const items = () => {
    const out = [];
    const seen = new Set();
    for (const t of TABS) for (const it of leafItems(t)) { const label = (it.label ?? it.title ?? '').replace(/\s*\(.*\)$/, ''); const cmd = it.cmd ?? it.menu; if (!label || !cmd || seen.has(label)) continue; seen.add(label); out.push({ label: `${label}  ·  ${t.label}`, cmd: it.cmd, arg: it.arg, menu: it.menu }); }
    for (const [name, i] of Object.entries(COMMAND_INFO)) if (i.label && !seen.has(i.label)) { seen.add(i.label); out.push({ label: i.label, cmd: name }); }
    return out;
  };
  box.addEventListener('input', () => {
    const q = box.value.trim().toLowerCase();
    if (!q) { closeMenus(); return; }
    const hits = items().filter((x) => x.label.toLowerCase().includes(q)).slice(0, 14);
    const list = hits.map((h) => ({ label: h.label, action: () => { box.value = ''; if (h.cmd) run(h.cmd, h.arg); else if (h.menu) run(h.menu, box); } }));
    list.push({ sep: true }, { label: `"${box.value}" 문서에서 찾기`, action: () => { run('findPane'); run('find', box.value, {}); } });
    openMenu(box, list, { minWidth: box.offsetWidth });
    box.focus();
  });
  box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') { box.value = ''; closeMenus(); focusEditor(); } if (e.key === 'Enter') { const first = document.querySelector('#menuLayer .menu .menu-item'); first?.click(); } });
}

// ───────────── 문서: 오른쪽 클릭 · 확대 · 미니 도구 모음 ─────────────
function wireDoc() {
  scroller.addEventListener('wheel', (e) => {
    if (!(e.ctrlKey || e.metaKey)) return;
    e.preventDefault();
    setZoom(Math.max(0.1, Math.min(5, S.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1))));
  }, { passive: false });
  scroller.addEventListener('contextmenu', (e) => {
    if (e.target.closest('input,textarea')) return;
    e.preventDefault();
    openMenu({ x: e.clientX, y: e.clientY }, docMenu(e));
  });
  document.addEventListener('mouseup', () => setTimeout(() => { if (S.opts.miniToolbar !== false) maybeMiniToolbar(true); }, 0));
}
function inTable() { return curSel().f.p.length > 1 && !S.hf; }
function docMenu(e) {
  const sel = curSel();
  const hasSel = !posEq(sel.a, sel.f);
  const img = S.objSel || e.target.closest('[data-img]');
  const link = e.target.closest('[data-link]');
  const items = [
    { label: '잘라내기(T)', icon: 'cut', disabled: !hasSel, action: () => run('cut') },
    { label: '복사(C)', icon: 'copy', disabled: !hasSel, action: () => run('copy') },
    { label: '붙여넣기 옵션:', icon: 'paste', submenu: [{ label: '원본 서식 유지(K)', action: () => run('paste') }, { label: '텍스트만 유지(T)', action: () => run('paste', 'text') }] },
    { sep: true },
  ];
  // 맞춤법 제안
  const issue = (S.spellIssues ?? []).find((it) => it.p.join() === sel.f.p.join() && sel.f.o >= it.s && sel.f.o <= it.e);
  if (issue) items.unshift({ label: `→ ${issue.sug}`, action: () => run('applySpell', S.spellIssues.indexOf(issue)) }, { sep: true });
  if (img) items.push({ label: '그림 바꾸기(A)', icon: 'picture', action: () => run('changePicture') }, { label: '대체 텍스트 편집(A)…', action: () => run('altTextDialog') }, { label: '텍스트 줄 바꿈(W)', icon: 'wrap', action: (a) => run('wrapMenu', a) }, { label: '자르기', icon: 'crop', action: () => run('cropPicture') }, { label: '그림 서식(O)…', action: () => run('imageProps', {}) }, { sep: true });
  if (inTable()) items.push(
    { label: '삽입(I)', icon: 'rowInsert', submenu: [{ label: '위에 행 삽입(A)', action: () => run('insertRowAbove') }, { label: '아래에 행 삽입(B)', action: () => run('insertRowBelow') }, { label: '왼쪽에 열 삽입(L)', action: () => run('insertColLeft') }, { label: '오른쪽에 열 삽입(R)', action: () => run('insertColRight') }] },
    { label: '셀 삭제(D)', submenu: [{ label: '행 삭제(R)', action: () => run('deleteRow') }, { label: '열 삭제(C)', action: () => run('deleteCol') }, { label: '표 삭제(T)', action: () => run('deleteTable') }] },
    { label: '셀 병합(M)', action: () => run('mergeCells') }, { label: '셀 분할(P)…', action: () => run('splitCellDialog') },
    { label: '테두리 스타일(B)', icon: 'border', action: (a) => run('cellBordersMenu', a) },
    { label: '표 속성(R)…', action: () => run('tablePropsDialog') }, { sep: true });
  if (link) items.push({ label: '하이퍼링크 편집(H)…', icon: 'link', action: () => run('insertLinkDialog') }, { label: '하이퍼링크 열기(O)', action: () => emit('followLink', link.dataset.link) }, { label: '하이퍼링크 복사(C)', action: () => navigator.clipboard?.writeText(link.dataset.link) }, { label: '하이퍼링크 제거(R)', action: () => run('removeLink') }, { sep: true });
  items.push(
    { label: '글꼴(F)…', icon: 'fontColor', action: () => run('fontDialog') },
    { label: '단락(P)…', icon: 'marks', action: () => run('paragraphDialog') },
    { label: '스마트 조회(L)', icon: 'search', disabled: !hasSel, action: () => { const q = run('__selText'); if (q) window.open(`https://ko.wikipedia.org/w/index.php?search=${encodeURIComponent(q)}`, '_blank', 'noopener'); } },
    { label: '번역(S)', icon: 'translate', action: () => run('translateDialog', 'sel') },
    { label: '링크(I)', icon: 'link', action: () => run('insertLinkDialog') },
    { label: '새 메모(M)', icon: 'comment', action: () => run('newComment') },
    { sep: true },
    { label: '스타일', icon: 'styles', submenu: ['Normal', 'Heading1', 'Heading2', 'Heading3', 'Title', 'Subtitle', 'Quote', 'ListParagraph'].filter((k) => S.doc.styles[k]).map((k) => ({ label: S.doc.styles[k].name, action: () => run('applyStyle', k) })) },
    { label: '글머리 기호', icon: 'bullets', action: (a) => run('bulletsMenu', a) },
    { label: '번호 매기기', icon: 'numbering', action: (a) => run('numberingMenu', a) },
  );
  return items;
}
let mini = null;
function hideMini() { mini?.remove(); mini = null; }
function maybeMiniToolbar(fromMouse = false) {
  const sel = curSel();
  if (posEq(sel.a, sel.f) || isDialogOpen() || isMenuOpen() || backstageOpen()) { hideMini(); return; }
  if (!fromMouse) return;
  const r = document.querySelector('.selrect')?.getBoundingClientRect();
  if (!r) return;
  hideMini();
  const b = (icon, title, cmd, arg) => el('button', { class: 'rbtn small', title, 'data-icon': icon, onmousedown: (e) => e.preventDefault(), onclick: () => { run(cmd, arg); } });
  mini = el('div', { class: 'mini-toolbar' },
    el('div', { class: 'mt-row' }, b('bold', '굵게', 'bold'), b('italic', '기울임꼴', 'italic'), b('underline', '밑줄', 'underline'), b('highlight', '형광펜', 'highlight'), b('fontColor', '글꼴 색', 'fontColor'), b('growFont', '글꼴 크기 크게', 'growFont'), b('shrinkFont', '글꼴 크기 작게', 'shrinkFont')),
    el('div', { class: 'mt-row' }, b('bullets', '글머리 기호', 'bullets'), b('numbering', '번호 매기기', 'numbering'), b('styles', '스타일', 'stylesPane'), b('comment', '새 메모', 'newComment'), b('link', '링크', 'insertLinkDialog'), b('formatPainter', '서식 복사', 'formatPainter')));
  document.body.append(mini);
  hydrateIcons(mini);
  Object.assign(mini.style, { left: `${Math.max(8, Math.min(innerWidth - 300, r.left))}px`, top: `${Math.max(8, r.top - 74)}px`, opacity: '0.35' });
  mini.addEventListener('mouseenter', () => { mini.style.opacity = '1'; });
}
document.addEventListener('mousemove', (e) => { if (!mini) return; const r = mini.getBoundingClientRect(); const d = Math.max(0, r.left - e.clientX, e.clientX - r.right, r.top - e.clientY, e.clientY - r.bottom); mini.style.opacity = d > 140 ? '0' : d > 40 ? String(1 - (d - 40) / 100) : '1'; if (d > 220) hideMini(); });
document.addEventListener('keydown', () => hideMini(), true);

// ───────────── 단축키 (Word) ─────────────
const K = (spec, fn) => {
  const parts = spec.toLowerCase().split('+');
  const key = parts.pop();
  const need = { ctrl: parts.includes('ctrl'), shift: parts.includes('shift'), alt: parts.includes('alt') };
  KEYMAP.push({ test: (e) => (e.ctrlKey || e.metaKey) === need.ctrl && e.shiftKey === need.shift && e.altKey === need.alt && (e.key.toLowerCase() === key || e.code.toLowerCase() === `key${key}` || e.code.toLowerCase() === `digit${key}` || (key === 'enter' && e.key === 'Enter') || (key === 'space' && e.code === 'Space') || (key === 'up' && e.key === 'ArrowUp') || (key === 'down' && e.key === 'ArrowDown')), run: fn });
};
K('ctrl+z', () => run('undo')); K('ctrl+y', () => run(S.lastRepeat && !S.history.canRedo?.() ? 'repeat' : 'redo')); K('ctrl+shift+z', () => run('redo'));
K('ctrl+s', () => run('save')); K('ctrl+o', () => run('backstage', 'open')); K('ctrl+n', () => run('newBlank')); K('ctrl+p', () => run('backstage', 'print')); K('f12', () => run('backstage', 'saveAs')); K('ctrl+w', () => run('backstage', 'close'));
K('ctrl+b', () => run('bold')); K('ctrl+i', () => run('italic')); K('ctrl+u', () => run('underline'));
K('ctrl+e', () => run('alignCenter')); K('ctrl+l', () => run('alignLeft')); K('ctrl+r', () => run('alignRight')); K('ctrl+j', () => run('alignJustify'));
K('ctrl+a', () => run('selectAll'));
K('ctrl+1', () => run('lineSpacing', 1)); K('ctrl+2', () => run('lineSpacing', 2)); K('ctrl+5', () => run('lineSpacing', 1.5)); K('ctrl+0', () => run('toggleSpaceBefore'));
K('ctrl+alt+1', () => run('applyStyle', 'Heading1')); K('ctrl+alt+2', () => run('applyStyle', 'Heading2')); K('ctrl+alt+3', () => run('applyStyle', 'Heading3'));
K('ctrl+shift+n', () => run('applyStyle', 'Normal'));
K('ctrl+]', () => run('growFont1')); K('ctrl+[', () => run('shrinkFont1'));
K('ctrl+shift+>', () => run('growFont')); K('ctrl+shift+<', () => run('shrinkFont'));
K('ctrl+shift+.', () => run('growFont')); K('ctrl+shift+,', () => run('shrinkFont'));
K('ctrl+=', () => run('subscript')); K('ctrl+shift++', () => run('superscript')); K('ctrl+shift+=', () => run('superscript'));
K('ctrl+space', () => run('clearCharFormat')); K('ctrl+q', () => run('paragraphProps', { ind: undefined, spc: undefined, align: undefined }));
K('ctrl+m', () => run('indentMore')); K('ctrl+shift+m', () => run('indentLess')); K('ctrl+t', () => run('hangingIndent'));
K('ctrl+shift+8', () => run('toggleMarks')); K('ctrl+shift+*', () => run('toggleMarks'));
K('f4', () => run('repeat')); K('shift+f3', () => run('cycleCase'));
K('ctrl+shift+d', () => run('dstrike')); K('ctrl+shift+k', () => run('smallCaps')); K('ctrl+shift+a', () => run('allCaps'));
K('ctrl+shift+l', () => run('bullets'));
K('ctrl+d', () => run('fontDialog'));
K('ctrl+f', () => run('findPane')); K('ctrl+h', () => run('replaceDialog', 'replace')); K('ctrl+g', () => run('replaceDialog', 'goto')); K('f5', () => run('replaceDialog', 'goto'));
K('ctrl+k', () => run('insertLinkDialog'));
K('ctrl+alt+m', () => run('newComment')); K('ctrl+shift+e', () => run('trackChanges'));
K('ctrl+alt+f', () => run('insertFootnote', 'foot')); K('ctrl+alt+d', () => run('insertFootnote', 'end'));
K('f7', () => run('spellCheck')); K('ctrl+shift+g', () => run('wordCountDialog')); K('f9', () => run('updateFields'));
K('ctrl+shift+c', () => run('formatPainter')); K('ctrl+shift+v', () => run('pastePlain'));
K('alt+=', () => run('equationDialog')); K('alt+shift+d', () => run('insertDate')); K('alt+shift+t', () => run('dateTimeDialog'));
K('ctrl+alt+shift+s', () => run('stylesPane'));
K('ctrl+alt+p', () => run('viewPrint')); K('ctrl+alt+o', () => run('viewOutline')); K('ctrl+alt+n', () => run('viewDraft'));
K('alt+shift+up', () => run('moveUp')); K('alt+shift+down', () => run('moveDown'));
K('ctrl+f1', () => document.body.classList.toggle('ribbon-collapsed'));
K('ctrl+shift+enter', () => run('columnBreak'));
K('ctrl+enter', () => run('pageBreak'));
K('alt+q', () => $('searchBox').focus());
K('f1', () => run('helpPane'));
// 편집기 밖(리본 · 패널)에서도 파일 단축키
document.addEventListener('keydown', (e) => {
  if (e.target === $('ime') || e.target.closest?.('input,textarea,select,[contenteditable]')) return;
  if (isDialogOpen()) return;
  for (const k of KEYMAP) if (['s', 'o', 'n', 'p', 'f'].includes(e.key.toLowerCase()) && (e.ctrlKey || e.metaKey) && k.test(e)) { e.preventDefault(); k.run(e); return; }
  if (e.key === 'Escape' && backstageOpen()) closeBackstage();
});
addEventListener('beforeunload', (e) => { if (S.dirty && S.opts.autosave === false) { e.preventDefault(); e.returnValue = ''; } });

// ───────────── 시작 ─────────────
boot();
(async () => {
  const params = new URLSearchParams(location.search);
  let started = false;
  if (params.get('new') != null) { run('newBlank'); started = true; }
  else if (!params.has('tour')) started = await restoreAutosave();
  if (!started) setDoc(sampleDoc(), 'WIWORD 둘러보기');
  applyViewClass();
  renderPanels();
  renderRibbon();
  updateTitle();
  if (location.protocol.startsWith('http')) fetch('/api/bridge/status').then((r) => (r.ok ? r.json() : null)).then((j) => { if (j) bridgeConnect(); }).catch(() => {});
  try { if (localStorage.getItem('wiword:version') !== APP_VERSION) { localStorage.setItem('wiword:version', APP_VERSION); setTimeout(() => toast('WIWORD 에 오신 것을 환영합니다 — 파일 탭에서 .docx · .hwp · .hwpx · PDF · 마크다운을 열 수 있습니다'), 800); } } catch { /* 저장소 없음 */ }
  requestAnimationFrame(() => focusEditor());
})();
window.__wiword = { S, run, relayout, V, setDoc, change, setZoom, openFile, exportPdfBytes, renderStatus, applyZoomMode, renderVisible, drawSelection, blockAt, paraPlain };
