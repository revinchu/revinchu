// 슬라이드 미리 보기 창 · 메모 · 상태 표시줄 · 여러 슬라이드 보기 · 오른쪽 작업 창(서식/애니메이션/선택)
import { S, curSlide, selObjects, change, emit, on, run, goSlide } from './state.js';
import { slideHtml, DASH_LABEL } from './render.js';
import { ANIM_EFFECTS, ANIM_CLASS_LABEL, TRANSITION_LABEL, plainText, outlineRows, applyOutline, masterSlide, LAYOUTS } from './model.js';
import { SHAPE_LABEL, objLabel } from './shapes.js';
import { el, openMenu } from './ui.js';
import { colorButton } from './colorpick.js';
import { fitZoom, setZoom } from './editor.js';

let thumbsEl;
let notesEl;
let statusEl;
let paneEl;
let sorterEl;
const THUMB_W = 168;

export function initPanels({ thumbs, notes, status, pane, sorter }) {
  thumbsEl = thumbs;
  notesEl = notes;
  statusEl = status;
  paneEl = pane;
  sorterEl = sorter;
  thumbs.tabIndex = 0;
  thumbs.addEventListener('keydown', onThumbKey);
  thumbs.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const t = e.target.closest('.thumb');
    if (t) { const i = Number(t.dataset.i); if (!S.slideSel.has(S.pres.slides[i].id)) goSlide(i); }
    emit('thumbMenu', { e, index: t ? Number(t.dataset.i) : null });
  });
  notes.addEventListener('input', () => {
    const s = curSlide();
    change(() => { s.notes = notes.value; }, { scope: 'none', key: `notes:${s.id}` });
  });
  notes.addEventListener('keydown', (e) => e.stopPropagation());
}

// ───────────── 미리 보기 창 ─────────────
function thumbHtml(slide, i, width) {
  const sc = width / S.pres.size.w;
  return `<div class="tw" style="width:${width}px;height:${Math.round(S.pres.size.h * sc)}px"><div class="tw-in" style="transform:scale(${sc})">${slideHtml(S.pres, slide, { index: i })}</div></div>`;
}

// ───────────── 개요 보기 (PowerPoint: 왼쪽 창에서 제목 · 본문을 바로 편집) ─────────────
let olBusy = false;
function renderOutline() {
  if (olBusy || thumbsEl.contains(document.activeElement) && thumbsEl.querySelector('.ol-edit')) return;
  thumbsEl.innerHTML = '';
  const box = el('div', { class: 'ol-edit' });
  for (const r of outlineRows(S.pres)) box.append(olRow(r));
  thumbsEl.append(box);
  olNumber(box);
  box.addEventListener('keydown', olKey);
  box.addEventListener('input', () => { clearTimeout(olTimer); olTimer = setTimeout(olApply, 300); });
  box.addEventListener('focusin', (e) => { const i = olSlideIndex(e.target.closest('.ol-row')); if (i >= 0 && i !== S.cur) { olBusy = true; try { goSlide(i); } finally { olBusy = false; } } });
  box.addEventListener('focusout', () => setTimeout(() => { if (!thumbsEl.contains(document.activeElement)) { olApply(); renderOutline(); } }, 0));
}
let olTimer = 0;
function olRow(r) {
  const d = el('div', { class: `ol-row ${r.kind}`, contenteditable: 'plaintext-only', spellcheck: 'false' }, r.text);
  d.dataset.kind = r.kind;
  d.dataset.lvl = String(r.lvl ?? 0);
  if (r.slide) d.dataset.slide = r.slide;
  if (r.obj) d.dataset.obj = r.obj;
  if (r.pi != null) d.dataset.pi = String(r.pi);
  d.style.paddingLeft = r.kind === 'body' ? `${34 + (r.lvl ?? 0) * 18}px` : '';
  return d;
}
function olNumber(box) { let n = 0; for (const d of box.children) if (d.dataset.kind === 'title') d.dataset.n = String(++n); }
function olSlideIndex(row) { if (!row) return -1; let n = -1; for (const d of row.parentElement.children) { if (d.dataset.kind === 'title') n++; if (d === row) break; } return n; }
function olRowsFromDom(box) {
  return [...box.children].map((d) => ({ kind: d.dataset.kind, text: d.textContent.replace(/\n/g, ' '), lvl: Number(d.dataset.lvl) || 0, slide: d.dataset.slide || null, obj: d.dataset.obj || null, pi: d.dataset.pi != null && d.dataset.pi !== '' ? Number(d.dataset.pi) : null, el: d }));
}
function olApply() {
  clearTimeout(olTimer);
  const box = thumbsEl?.querySelector('.ol-edit');
  if (!box) return;
  const rows = olRowsFromDom(box);
  const cur = JSON.stringify(outlineRows(S.pres).map((r) => [r.kind, r.text, r.lvl ?? 0]));
  const next = JSON.stringify(rows.filter((r) => r.kind === 'title' || r.text).map((r) => [r.kind, r.text, r.kind === 'body' ? r.lvl : 0]));
  if (cur === next && rows.filter((r) => r.kind === 'title').length === S.pres.slides.length) return;
  olBusy = true;
  try { change(() => applyOutline(S.pres, rows), { scope: 'all' }); } finally { olBusy = false; }
  // 새로 생긴 슬라이드 · 개체 id 를 줄에 기록 (다음 적용 때 같은 슬라이드로)
  for (const r of rows) { if (r.slide) r.el.dataset.slide = r.slide; if (r.obj) r.el.dataset.obj = r.obj; if (r.pi != null) r.el.dataset.pi = String(r.pi); }
  olNumber(box);
}
function olCaretAtStart(d) { const s = getSelection(); if (!s.rangeCount) return false; const r = s.getRangeAt(0).cloneRange(); r.selectNodeContents(d); r.setEnd(s.getRangeAt(0).startContainer, s.getRangeAt(0).startOffset); return r.toString().length === 0; }
function olFocus(d, atEnd = false) { d.focus(); const r = document.createRange(); r.selectNodeContents(d); r.collapse(!atEnd); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
function olOnlyPlaceholders(row) {
  const s = S.pres.slides.find((x) => x.id === row.dataset.slide);
  return !s || s.objects.every((o) => o.ph && o.text);
}
function olSetKind(d, kind, lvl) {
  d.dataset.kind = kind; d.dataset.lvl = String(lvl);
  d.className = `ol-row ${kind}`;
  d.style.paddingLeft = kind === 'body' ? `${34 + lvl * 18}px` : '';
  if (kind === 'title') { delete d.dataset.obj; delete d.dataset.pi; } else delete d.dataset.slide;
}
function olKey(e) {
  const d = e.target.closest('.ol-row');
  if (!d) return;
  e.stopPropagation();
  const kind = d.dataset.kind;
  const lvl = Number(d.dataset.lvl) || 0;
  if (e.key === 'Enter') {
    e.preventDefault();
    // 커서 뒤 글은 새 줄로 (제목 줄에서 Enter = 새 슬라이드)
    const s = getSelection();
    let tail = '';
    if (s.rangeCount) { const r = s.getRangeAt(0).cloneRange(); r.setEndAfter(d.lastChild ?? d); tail = r.toString(); r.deleteContents(); }
    const nd = olRow({ kind, lvl, text: tail });
    d.after(nd);
    olNumber(d.parentElement);
    olFocus(nd);
    olApply();
  } else if (e.key === 'Tab' && !e.shiftKey) {
    e.preventDefault();
    if (kind === 'title') {
      if (!d.previousElementSibling) return;
      if (!olOnlyPlaceholders(d)) { emit('error', new Error('그림 · 표 같은 다른 개체가 있는 슬라이드는 앞 슬라이드와 합칠 수 없습니다')); return; }
      olSetKind(d, 'body', 0);
    } else olSetKind(d, 'body', Math.min(8, lvl + 1));
    olNumber(d.parentElement); olApply();
  } else if (e.key === 'Tab' && e.shiftKey) {
    e.preventDefault();
    if (kind === 'body' && lvl === 0) olSetKind(d, 'title', 0);
    else if (kind === 'body') olSetKind(d, 'body', lvl - 1);
    olNumber(d.parentElement); olApply();
  } else if (e.key === 'Backspace' && olCaretAtStart(d) && d.previousElementSibling) {
    if (kind === 'title' && !olOnlyPlaceholders(d)) return;
    e.preventDefault();
    const prev = d.previousElementSibling;
    const keep = d.textContent;
    const n = prev.textContent.length;
    prev.textContent += keep;
    d.remove();
    olNumber(prev.parentElement);
    prev.focus();
    const t = prev.firstChild;
    if (t) { const r = document.createRange(); r.setStart(t, n); r.collapse(true); const s = getSelection(); s.removeAllRanges(); s.addRange(r); }
    olApply();
  } else if (e.key === 'ArrowUp' && d.previousElementSibling && (e.altKey || olCaretAtStart(d))) { e.preventDefault(); olFocus(d.previousElementSibling, true); }
  else if (e.key === 'ArrowDown' && d.nextElementSibling) { const s = getSelection(); const r = s.rangeCount ? s.getRangeAt(0).cloneRange() : null; if (r) { r.setEndAfter(d.lastChild ?? d); if (r.toString().length === 0 || e.altKey) { e.preventDefault(); olFocus(d.nextElementSibling); } } }
}

// ───────────── 슬라이드 마스터 보기: 마스터 + 레이아웃 목록 ─────────────
function renderMasterList() {
  thumbsEl.innerHTML = '';
  const keys = ['master', ...LAYOUTS.map(([k]) => k)];
  keys.forEach((key, i) => {
    const m = masterSlide(S.pres, key);
    const name = m.name ?? (key === 'master' ? '슬라이드 마스터' : LAYOUTS.find(([k]) => k === key)?.[1]);
    const w = key === 'master' ? THUMB_W : THUMB_W - 30;
    const t = el('div', { class: `thumb master-thumb${key === 'master' ? ' is-master' : ''}${S.masterKey === key ? ' active' : ''}`, title: name },
      el('div', { class: 'tnum' }, el('span', {}, key === 'master' ? '1' : '')),
      el('div', { class: 'tbox', html: thumbHtml(m, i, w) }),
      el('div', { class: 'mname' }, name));
    t.addEventListener('mousedown', (e) => { if (e.button === 0) run('masterPick', key); });
    thumbsEl.append(t);
  });
}

export function renderThumbs() {
  if (!thumbsEl) return;
  thumbsEl.classList.toggle('outline-mode', S.view === 'outline' && !S.masterKey);
  if (S.masterKey) { renderMasterList(); return; }
  if (S.view === 'outline') { renderOutline(); return; }
  thumbsEl.innerHTML = '';
  S.pres.slides.forEach((s, i) => {
    if (s.section) thumbsEl.append(sectionHead(s, i));
    const t = el('div', { class: `thumb${i === S.cur ? ' active' : ''}${S.slideSel.has(s.id) ? ' picked' : ''}${s.hidden ? ' hidden-slide' : ''}`, dataset: { i: String(i) }, draggable: 'true' },
      el('div', { class: 'tnum' }, el('span', {}, String(i + 1)), (s.transition && s.transition.type !== 'none') || s.anims?.length ? el('span', { class: 'tstar', title: '전환 또는 애니메이션' }, '★') : null),
      el('div', { class: 'tbox', html: thumbHtml(s, i, THUMB_W) }));
    t.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      if (e.ctrlKey || e.metaKey) {
        if (S.slideSel.has(s.id) && S.slideSel.size > 1) S.slideSel.delete(s.id); else S.slideSel.add(s.id);
        S.cur = i;
        emit('change', { scope: 'nav', keepPick: true });
        return;
      }
      if (e.shiftKey) {
        const [a, b] = [Math.min(S.cur, i), Math.max(S.cur, i)];
        S.slideSel = new Set(S.pres.slides.slice(a, b + 1).map((x) => x.id));
        S.cur = i;
        emit('change', { scope: 'nav', keepPick: true });
        return;
      }
      goSlide(i);
    });
    t.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-wipoint-slide', String(i)); e.dataTransfer.effectAllowed = 'move'; });
    t.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('text/x-wipoint-slide')) { e.preventDefault(); const r = t.getBoundingClientRect(); t.classList.toggle('drop-before', e.clientY < r.top + r.height / 2); t.classList.toggle('drop-after', e.clientY >= r.top + r.height / 2); } });
    t.addEventListener('dragleave', () => t.classList.remove('drop-before', 'drop-after'));
    t.addEventListener('drop', (e) => {
      e.preventDefault();
      const after = t.classList.contains('drop-after');
      t.classList.remove('drop-before', 'drop-after');
      const from = Number(e.dataTransfer.getData('text/x-wipoint-slide'));
      run('moveSlidesTo', after ? i + 1 : i, from);
    });
    thumbsEl.append(t);
  });
  thumbsEl.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' });
}

function sectionHead(s, i) {
  const h = el('div', { class: 'section-head', title: '구역 (오른쪽 클릭: 이름 바꾸기 · 제거)' }, el('span', {}, '▾ '), s.section);
  h.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    e.stopPropagation();
    openMenu({ x: e.clientX, y: e.clientY }, [
      { label: '구역 이름 바꾸기', action: () => run('renameSection', i) },
      { label: '구역 제거', action: () => run('removeSection', i) },
    ]);
  });
  return h;
}

/** 현재 슬라이드의 미리 보기만 다시 그림 */
export function updateThumb(i = S.cur) {
  if (S.masterKey) { renderMasterList(); return; }
  if (S.view === 'outline') { renderOutline(); return; }
  const t = thumbsEl?.querySelector(`.thumb[data-i="${i}"] .tbox`);
  const s = S.pres.slides[i];
  if (t && s) t.innerHTML = thumbHtml(s, i, THUMB_W);
}
export function markActiveThumb() {
  if (!thumbsEl) return;
  if (S.masterKey) { renderMasterList(); return; }
  if (S.view === 'outline') { renderOutline(); return; }
  for (const t of thumbsEl.querySelectorAll('.thumb')) {
    const i = Number(t.dataset.i);
    t.classList.toggle('active', i === S.cur);
    t.classList.toggle('picked', S.slideSel.has(S.pres.slides[i]?.id));
  }
  thumbsEl.querySelector('.thumb.active')?.scrollIntoView({ block: 'nearest' });
}

function onThumbKey(e) {
  if (S.editing) return;
  const k = e.key;
  if (k === 'ArrowDown' || k === 'ArrowRight' || k === 'PageDown') { e.preventDefault(); goSlide(S.cur + 1); thumbsEl.focus(); }
  else if (k === 'ArrowUp' || k === 'ArrowLeft' || k === 'PageUp') { e.preventDefault(); goSlide(S.cur - 1); thumbsEl.focus(); }
  else if (k === 'Home') { e.preventDefault(); goSlide(0); }
  else if (k === 'End') { e.preventDefault(); goSlide(S.pres.slides.length - 1); }
  else if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); run('deleteSlide'); thumbsEl.focus(); }
  else if (k === 'Enter') { e.preventDefault(); run('newSlide'); thumbsEl.focus(); }
  else return;
  e.stopPropagation();
}
export const thumbsFocused = () => document.activeElement === thumbsEl;

// ───────────── 여러 슬라이드 보기 ─────────────
export function renderSorter() {
  if (!sorterEl) return;
  sorterEl.innerHTML = '';
  const zoomW = Math.round(220 * Math.max(0.6, Math.min(2, S.sorterZoom ?? 1)));
  const grid = el('div', { class: 'sorter-grid' });
  S.pres.slides.forEach((s, i) => {
    if (s.section) grid.append(el('div', { class: 'sorter-section' }, s.section));
    const c = el('div', { class: `sorter-item${i === S.cur ? ' active' : ''}${S.slideSel.has(s.id) ? ' picked' : ''}${s.hidden ? ' hidden-slide' : ''}`, draggable: 'true', dataset: { i: String(i) } },
      el('div', { class: 'tbox', html: thumbHtml(s, i, zoomW) }),
      el('div', { class: 'sorter-cap' }, el('span', {}, String(i + 1)), s.transition && s.transition.type !== 'none' ? el('span', { class: 'tstar', title: TRANSITION_LABEL[s.transition.type] }, '★') : null));
    c.addEventListener('mousedown', (e) => {
      if (e.ctrlKey || e.metaKey) { if (S.slideSel.has(s.id)) S.slideSel.delete(s.id); else S.slideSel.add(s.id); S.cur = i; } else goSlide(i);
      renderSorter();
    });
    c.addEventListener('dblclick', () => { goSlide(i); run('viewNormal'); });
    c.addEventListener('contextmenu', (e) => { e.preventDefault(); if (!S.slideSel.has(s.id)) goSlide(i); emit('thumbMenu', { e, index: i }); });
    c.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-wipoint-slide', String(i)); });
    c.addEventListener('dragover', (e) => { e.preventDefault(); c.classList.add('drop-before'); });
    c.addEventListener('dragleave', () => c.classList.remove('drop-before'));
    c.addEventListener('drop', (e) => { e.preventDefault(); c.classList.remove('drop-before'); run('moveSlidesTo', i, Number(e.dataTransfer.getData('text/x-wipoint-slide'))); });
    grid.append(c);
  });
  sorterEl.append(grid);
}

// ───────────── 메모 ─────────────
export function renderNotes() {
  if (!notesEl) return;
  const s = curSlide();
  if (document.activeElement !== notesEl) notesEl.value = s?.notes ?? '';
}

// ───────────── 상태 표시줄 ─────────────
export function renderStatus() {
  if (!statusEl) return;
  const n = S.pres.slides.length;
  const z = Math.round(S.zoom * 100);
  statusEl.innerHTML = '';
  const viewBtn = (cmd, icon, title, on) => {
    const b = el('button', { class: `sb-btn${on ? ' on' : ''}`, title, 'data-icon': icon, onclick: () => run(cmd) });
    return b;
  };
  const slider = el('input', { type: 'range', min: '10', max: '400', step: '5', value: String(z), class: 'zoom-slider', 'aria-label': '확대/축소' });
  slider.addEventListener('input', () => setZoom(Number(slider.value) / 100));
  statusEl.append(
    el('span', { class: 'sb-item' }, `슬라이드 ${S.cur + 1}/${n}`),
    el('span', { class: 'sb-item muted' }, '한국어'),
    el('span', { class: 'sb-item muted', id: 'saveState' }, S.dirty ? '' : ''),
    el('span', { class: 'sb-spacer' }),
    el('button', { class: `sb-btn text${S.showNotes ? ' on' : ''}`, onclick: () => run('toggleNotes') }, el('span', { 'data-icon': 'notes' }), ' 슬라이드 노트'),
    viewBtn('viewNormal', 'normalView', '기본', S.view === 'normal'),
    viewBtn('viewSorter', 'sorter', '여러 슬라이드', S.view === 'sorter'),
    viewBtn('viewReading', 'readingView', '읽기용 보기', S.view === 'reading'),
    viewBtn('showFromCurrent', 'slideshow', '슬라이드 쇼 (Shift+F5)', false),
    el('button', { class: 'sb-btn', title: '축소', onclick: () => setZoom(S.zoom / 1.1) }, '−'),
    slider,
    el('button', { class: 'sb-btn', title: '확대', onclick: () => setZoom(S.zoom * 1.1) }, '+'),
    el('button', { class: 'sb-btn text', title: '확대/축소', onclick: () => run('zoomDialog') }, `${z}%`),
    el('button', { class: 'sb-btn', title: '현재 창 크기에 맞춤', 'data-icon': 'fit', onclick: () => fitZoom() }));
  emit('hydrate', statusEl);
}

// ───────────── 오른쪽 작업 창 ─────────────
const sec = (title, ...body) => el('section', { class: 'pane-sec' }, el('h4', {}, title), ...body);
const row = (label, ...ctrl) => el('label', { class: 'pane-row' }, el('span', {}, label), ...ctrl);
function num(value, onSet, { step = 1, min, max, suffix = '' } = {}) {
  const i = el('input', { type: 'number', value: String(Math.round(value * 100) / 100), step: String(step), ...(min != null ? { min: String(min) } : {}), ...(max != null ? { max: String(max) } : {}) });
  i.addEventListener('change', () => { const v = Number(i.value); if (Number.isFinite(v)) onSet(v); });
  i.addEventListener('keydown', (e) => e.stopPropagation());
  return suffix ? el('span', { class: 'num-wrap' }, i, el('small', {}, suffix)) : i;
}
function sel(value, options, onSet) {
  const s = el('select', {}, options.map(([v, l]) => el('option', { value: v, selected: v === value }, l)));
  s.addEventListener('change', () => onSet(s.value));
  s.addEventListener('keydown', (e) => e.stopPropagation());
  return s;
}
function check(value, onSet, label) {
  const c = el('input', { type: 'checkbox', checked: !!value });
  c.addEventListener('change', () => onSet(c.checked));
  return el('label', { class: 'pane-check' }, c, el('span', {}, label));
}
const radio = (name, value, options, onSet) => el('div', { class: 'pane-radios' }, options.map(([v, l]) => {
  const r = el('input', { type: 'radio', name, checked: v === value });
  r.addEventListener('change', () => { if (r.checked) onSet(v); });
  return el('label', { class: 'pane-check' }, r, el('span', {}, l));
}));

const PX_CM = 2.54 / 96;
const PANE_TITLE = { comments: '메모',  shape: '도형 서식', bg: '배경 서식', anim: '애니메이션 창', selection: '선택' };

export function renderPane() {
  if (!paneEl) return;
  const kind = S.formatPane;
  paneEl.hidden = !kind;
  document.body.classList.toggle('pane-open', !!kind);
  if (!kind) return;
  // 창 안의 입력 칸을 쓰는 중이면 다시 그리지 않음 (초점 유지)
  if (paneEl.contains(document.activeElement) && document.activeElement.tagName === 'INPUT' && document.activeElement.type !== 'checkbox' && document.activeElement.type !== 'radio') return;
  paneEl.innerHTML = '';
  const head = el('div', { class: 'pane-head' }, el('b', {}, PANE_TITLE[kind] ?? ''), el('button', { title: '닫기', onclick: () => { S.formatPane = null; renderPane(); fitZoomSoon(); } }, '✕'));
  const body = el('div', { class: 'pane-body' });
  paneEl.append(head, body);
  if (kind === 'shape') shapePane(body);
  else if (kind === 'bg') bgPane(body);
  else if (kind === 'anim') animPane(body);
  else if (kind === 'selection') selectionPane(body);
  else if (kind === 'comments') commentsPane(body);
  emit('hydrate', paneEl);
}
/** 메모 창: 슬라이드의 메모 목록, 답글, 해결, 삭제 */
function commentsPane(body) {
  const s = curSlide();
  const list = s.comments ?? [];
  const when = (iso) => { try { return new Date(iso).toLocaleString('ko-KR', { month: 'numeric', day: 'numeric', hour: '2-digit', minute: '2-digit' }); } catch { return ''; } };
  const top = el('div', { class: 'cm-top' },
    el('button', { class: 'btn primary', onclick: () => run('newComment') }, '+ 새 메모'),
    el('button', { class: 'btn', title: '이전 메모', onclick: () => run('prevComment') }, '◀'),
    el('button', { class: 'btn', title: '다음 메모', onclick: () => run('nextComment') }, '▶'));
  body.append(top);
  if (!list.length) { body.append(el('p', { class: 'muted' }, '이 슬라이드에는 메모가 없습니다.')); return; }
  for (const c of list) {
    const reply = el('input', { type: 'text', placeholder: '답글...' });
    reply.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key !== 'Enter' || !reply.value.trim()) return;
      const t = reply.value.trim();
      change(() => { c.replies = [...(c.replies ?? []), { author: localStorage.getItem('wipoint:user') ?? '사용자', text: t, at: new Date().toISOString() }]; }, { scope: 'none' });
      renderPane();
    });
    body.append(el('div', { class: `cm-card${c.done ? ' done' : ''}`, onmouseenter: () => document.querySelector(`.cm-pin[data-cm="${c.id}"]`)?.classList.add('hot'), onmouseleave: () => document.querySelector(`.cm-pin[data-cm="${c.id}"]`)?.classList.remove('hot') },
      el('div', { class: 'cm-head' }, el('b', {}, c.author ?? '사용자'), el('small', {}, when(c.at))),
      el('div', { class: 'cm-text' }, c.text),
      ...(c.replies ?? []).map((r) => el('div', { class: 'cm-reply' }, el('b', {}, r.author ?? ''), ' ', r.text)),
      reply,
      el('div', { class: 'cm-actions' },
        el('button', { class: 'btn', onclick: () => { change(() => { c.done = !c.done || undefined; }, { scope: 'none' }); renderPane(); } }, c.done ? '다시 열기' : '스레드 해결'),
        el('button', { class: 'btn', onclick: () => { change(() => { s.comments = list.filter((x) => x !== c); }); renderPane(); } }, '삭제'))));
  }
}

const fitZoomSoon = () => setTimeout(() => { if (S.fitZoom) fitZoom(); }, 0);

function fillControls(fill, setFill, name, { allowPicture = true } = {}) {
  const type = !fill ? 'none' : fill.type;
  const out = [radio(`${name}-type`, type, [['none', '채우기 없음'], ['solid', '단색 채우기'], ['gradient', '그라데이션 채우기'], ...(allowPicture ? [['image', '그림 채우기']] : [])], (t) => {
    if (t === 'none') setFill(null);
    else if (t === 'solid') setFill({ type: 'solid', color: fill?.color ?? fill?.stops?.[0]?.[1] ?? '@accent1' });
    else if (t === 'gradient') setFill({ type: 'gradient', angle: 90, stops: [[0, fill?.color ?? '@accent1'], [1, '@accent1:lm60:lo40']] });
    else run('pickFillImage', setFill);
  })];
  if (fill?.type === 'solid') {
    out.push(row('색', colorButton(fill.color, (c) => setFill({ ...fill, color: c ?? '@accent1' }))));
    const a = fill.color?.match(/:a([\d.]+)/);
    out.push(row('투명도', num(a ? 100 - Number(a[1]) : 0, (v) => { const base = String(fill.color).replace(/:a[\d.]+/, ''); setFill({ ...fill, color: v > 0 ? `${base.startsWith('#') || base.startsWith('@') ? base : '#000000'}:a${Math.max(0, 100 - v)}` : base }); }, { min: 0, max: 100, suffix: '%' })));
  }
  if (fill?.type === 'gradient') {
    out.push(row('종류', sel(fill.path ? 'path' : 'linear', [['linear', '선형'], ['path', '방사형']], (v) => setFill({ ...fill, path: v === 'path' || undefined }))));
    if (!fill.path) out.push(row('각도', num(fill.angle ?? 90, (v) => setFill({ ...fill, angle: v }), { suffix: '°' })));
    fill.stops.forEach(([p, c], i) => out.push(row(`중지점 ${i + 1}`, colorButton(c, (nc) => { const stops = fill.stops.map((s) => [...s]); stops[i][1] = nc ?? '#FFFFFF'; setFill({ ...fill, stops }); }), num(Math.round(p * 100), (v) => { const stops = fill.stops.map((s) => [...s]); stops[i][0] = Math.max(0, Math.min(100, v)) / 100; setFill({ ...fill, stops }); }, { suffix: '%' }))));
    out.push(el('div', { class: 'pane-btns' },
      el('button', { class: 'btn small', onclick: () => setFill({ ...fill, stops: [...fill.stops, [1, '@accent2']] }) }, '중지점 추가'),
      fill.stops.length > 2 ? el('button', { class: 'btn small', onclick: () => setFill({ ...fill, stops: fill.stops.slice(0, -1) }) }, '중지점 제거') : null));
  }
  if (fill?.type === 'image') out.push(el('div', { class: 'pane-btns' }, el('button', { class: 'btn small', onclick: () => run('pickFillImage', setFill) }, '그림 바꾸기...'), check(fill.tile, (v) => setFill({ ...fill, tile: v || undefined }), '바둑판식으로 배열')));
  return out;
}

function shapePane(body) {
  const objs = selObjects();
  if (!objs.length) { body.append(el('p', { class: 'muted' }, '개체를 선택하면 서식을 바꿀 수 있습니다.')); return; }
  const o = objs[0];
  const set = (fn, key) => change(() => { for (const x of objs) fn(x); }, { key });
  const shapes = objs.filter((x) => x.type === 'shape');
  if (shapes.length) {
    body.append(sec('채우기', ...fillControls(o.fill, (f) => set((x) => { if (x.type === 'shape') x.fill = f; }), 'fill')));
    const ln = o.line;
    body.append(sec('선',
      radio('line-type', ln ? 'solid' : 'none', [['none', '선 없음'], ['solid', '실선']], (t) => set((x) => { x.line = t === 'none' ? null : { color: '@tx1', width: 1.33, dash: 'solid' }; })),
      ln ? [
        row('색', colorButton(ln.color, (c) => set((x) => { x.line = { ...(x.line ?? {}), color: c ?? '@tx1' }; }))),
        row('너비', num(Math.round(ln.width * 0.75 * 100) / 100, (v) => set((x) => { x.line = { ...(x.line ?? {}), width: Math.max(0.25, v) / 0.75 }; }), { step: 0.25, min: 0, suffix: 'pt' })),
        row('대시 종류', sel(ln.dash ?? 'solid', DASH_LABEL, (v) => set((x) => { x.line = { ...(x.line ?? {}), dash: v }; }))),
        row('시작 화살표', sel(ln.head ?? 'none', ARROWS, (v) => set((x) => { x.line = { ...(x.line ?? {}), head: v === 'none' ? undefined : v }; }))),
        row('끝 화살표', sel(ln.tail ?? 'none', ARROWS, (v) => set((x) => { x.line = { ...(x.line ?? {}), tail: v === 'none' ? undefined : v }; }))),
      ] : null));
  }
  const imgs = objs.filter((x) => x.type === 'image' && x.media);
  if (imgs.length) {
    const im = imgs[0];
    body.append(sec('그림',
      row('밝기', num(Math.round((im.bright ?? 0) * 100), (v) => set((x) => { if (x.type === 'image') x.bright = v / 100 || undefined; }), { min: -100, max: 100, suffix: '%' })),
      row('대비', num(Math.round((im.contrast ?? 0) * 100), (v) => set((x) => { if (x.type === 'image') x.contrast = v / 100 || undefined; }), { min: -100, max: 100, suffix: '%' })),
      row('투명도', num(Math.round((1 - (im.alpha ?? 1)) * 100), (v) => set((x) => { if (x.type === 'image') x.alpha = v ? 1 - v / 100 : undefined; }), { min: 0, max: 100, suffix: '%' })),
      check(im.gray, (v) => set((x) => { if (x.type === 'image') x.gray = v || undefined; }), '회색조'),
      el('div', { class: 'pane-btns' }, el('button', { class: 'btn small', onclick: () => run('cropPicture') }, '자르기...'), el('button', { class: 'btn small', onclick: () => run('resetPicture') }, '그림 원래대로'))));
  }
  body.append(sec('효과', check(o.shadow, (v) => set((x) => { x.shadow = v || undefined; }), '그림자')));
  // 크기 및 속성
  const one = objs.length === 1;
  body.append(sec('크기 및 위치',
    row('높이', num(o.h * PX_CM, (v) => set((x) => { const k = (v / PX_CM) / x.h; x.h = v / PX_CM; if (S.lockRatio) x.w *= k; }), { step: 0.01, min: 0, suffix: 'cm' })),
    row('너비', num(o.w * PX_CM, (v) => set((x) => { const k = (v / PX_CM) / x.w; x.w = v / PX_CM; if (S.lockRatio) x.h *= k; }), { step: 0.01, min: 0, suffix: 'cm' })),
    row('회전', num(o.rot ?? 0, (v) => set((x) => { x.rot = ((v % 360) + 360) % 360; }), { suffix: '°' })),
    check(S.lockRatio, (v) => { S.lockRatio = v; }, '가로 세로 비율 고정'),
    one ? row('가로 위치', num(o.x * PX_CM, (v) => set((x) => { x.x = v / PX_CM; }), { step: 0.01, suffix: 'cm' })) : null,
    one ? row('세로 위치', num(o.y * PX_CM, (v) => set((x) => { x.y = v / PX_CM; }), { step: 0.01, suffix: 'cm' })) : null));
  const texts = objs.filter((x) => x.text);
  if (texts.length) {
    const t = texts[0].text;
    const setT = (fn) => change(() => { for (const x of texts) fn(x.text, x); });
    const ins = t.insets ?? [9.6, 4.8, 9.6, 4.8];
    body.append(sec('텍스트 상자',
      row('세로 맞춤', sel(t.anchor ?? 't', [['t', '위쪽'], ['ctr', '중간'], ['b', '아래쪽']], (v) => setT((b) => { b.anchor = v; }))),
      row('텍스트 방향', sel(t.vert ?? 'horz', [['horz', '가로'], ['eaVert', '세로'], ['vert270', '모든 텍스트 270° 회전']], (v) => setT((b) => { if (v === 'horz') delete b.vert; else b.vert = v; }))),
      radio('autofit', t.autofit ?? 'none', [['none', '자동 맞춤 안 함'], ['shrink', '넘치면 텍스트 크기 조정'], ['resize', '텍스트에 맞게 도형 크기 조정']], (v) => setT((b) => { b.autofit = v; if (v !== 'shrink') { delete b.fontScale; delete b.lnSpcReduction; } })),
      row('왼쪽 여백', num(ins[0] * PX_CM, (v) => setT((b) => { b.insets = [...(b.insets ?? ins)]; b.insets[0] = v / PX_CM; }), { step: 0.05, min: 0, suffix: 'cm' })),
      row('오른쪽 여백', num(ins[2] * PX_CM, (v) => setT((b) => { b.insets = [...(b.insets ?? ins)]; b.insets[2] = v / PX_CM; }), { step: 0.05, min: 0, suffix: 'cm' })),
      row('위쪽 여백', num(ins[1] * PX_CM, (v) => setT((b) => { b.insets = [...(b.insets ?? ins)]; b.insets[1] = v / PX_CM; }), { step: 0.05, min: 0, suffix: 'cm' })),
      row('아래쪽 여백', num(ins[3] * PX_CM, (v) => setT((b) => { b.insets = [...(b.insets ?? ins)]; b.insets[3] = v / PX_CM; }), { step: 0.05, min: 0, suffix: 'cm' })),
      check(t.wrap !== false, (v) => setT((b) => { b.wrap = v; }), '도형의 텍스트 배치')));
  }
  if (one) {
    const alt = el('textarea', { rows: '3', placeholder: '이 개체를 설명하는 글 (화면 읽기 프로그램용)' }, o.alt ?? '');
    alt.addEventListener('change', () => change(() => { o.alt = alt.value || undefined; }));
    alt.addEventListener('keydown', (e) => e.stopPropagation());
    body.append(sec('대체 텍스트', alt));
  }
}
const ARROWS = [['none', '없음'], ['triangle', '삼각형 화살표'], ['arrow', '열린 화살표'], ['oval', '타원 화살표'], ['diamond', '다이아몬드 화살표']];

function bgPane(body) {
  const slide = curSlide();
  const bg = slide.bg ?? null;
  const setBg = (f) => change(() => { slide.bg = f ?? { type: 'solid', color: '@bg1' }; });
  body.append(sec('채우기', ...fillControls(bg ?? S.pres.theme.bg ?? { type: 'solid', color: '@bg1' }, (f) => setBg(f), 'bgfill')));
  body.append(sec('배경 그래픽', check(slide.hideDecor, (v) => change(() => { slide.hideDecor = v || undefined; if (v && slide.bgObjects) { slide._bgObjects = slide.bgObjects; slide.bgObjects = []; } if (!v && slide._bgObjects) { slide.bgObjects = slide._bgObjects; delete slide._bgObjects; } }), '배경 그래픽 숨기기')));
  body.append(el('div', { class: 'pane-btns' },
    el('button', { class: 'btn small', onclick: () => change(() => { for (const s of S.pres.slides) s.bg = slide.bg ? JSON.parse(JSON.stringify(slide.bg)) : null; }, { scope: 'all' }) }, '모두 적용'),
    el('button', { class: 'btn small', onclick: () => change(() => { slide.bg = null; }) }, '배경 원래대로')));
}


function animPane(body) {
  const slide = curSlide();
  const anims = slide.anims ?? [];
  body.append(el('div', { class: 'pane-btns' },
    el('button', { class: 'btn small', onclick: () => run('previewAnim') }, '▶ 모두 재생'),
    el('button', { class: 'btn small', onclick: () => run('addAnimMenu') }, '애니메이션 추가 ▾')));
  if (!anims.length) { body.append(el('p', { class: 'muted' }, '이 슬라이드에는 애니메이션이 없습니다. 개체를 고른 뒤 [애니메이션] 탭에서 효과를 고르세요.')); return; }
  const list = el('div', { class: 'anim-list' });
  // 기본 순서 다음에 트리거별 묶음 (PowerPoint: '트리거: 개체 이름' 머리글)
  const groups = [null, ...new Set(anims.map((a) => a.trigger).filter(Boolean))];
  for (const t of groups) {
  let step = 0;
  if (t) { const to = slide.objects.find((x) => x.id === t); list.append(el('div', { class: 'anim-trig' }, `트리거: ${to ? objLabel(to, slide.objects.indexOf(to)) : '(없음)'}`)); }
  anims.filter((a) => (a.trigger ?? null) === t).forEach((a) => {
    const o = slide.objects.find((x) => x.id === a.obj);
    if (a.start === 'click') step++;
    const label = ANIM_EFFECTS[a.cls]?.find(([k]) => k === a.effect)?.[1] ?? a.effect;
    const r = el('div', { class: `anim-row ${a.cls}${S.animSel === a.id ? ' on' : ''}` },
      el('span', { class: 'anim-step' }, a.start === 'click' ? String(step) : ''),
      el('span', { class: 'anim-ico', title: ANIM_CLASS_LABEL[a.cls] }, a.cls === 'entr' ? '★' : a.cls === 'exit' ? '✦' : '✷'),
      el('span', { class: 'anim-name' }, `${o ? objLabel(o, slide.objects.indexOf(o)) : '(없음)'}: ${label}`));
    r.addEventListener('click', () => { S.animSel = a.id; if (o) S.sel = new Set([o.id]); emit('selection'); renderPane(); });
    list.append(r);
  });
  }
  body.append(list);
  const cur = anims.find((a) => a.id === S.animSel);
  if (cur) {
    const setA = (fn) => change(() => { fn(cur); });
    body.append(sec('타이밍',
      row('효과', sel(cur.effect, ANIM_EFFECTS[cur.cls], (v) => setA((a) => { a.effect = v; }))),
      row('시작', sel(cur.start, [['click', '클릭할 때'], ['with', '이전 효과와 함께'], ['after', '이전 효과 다음에']], (v) => setA((a) => { a.start = v; }))),
      row('재생 시간', num(cur.dur ?? 0.5, (v) => setA((a) => { a.dur = Math.max(0, v); }), { step: 0.25, min: 0, suffix: '초' })),
      row('지연', num(cur.delay ?? 0, (v) => setA((a) => { a.delay = Math.max(0, v); }), { step: 0.25, min: 0, suffix: '초' })),
      ['fly', 'wipe', 'bounce'].includes(cur.effect) ? row('방향', sel(cur.dir ?? 'b', [['b', '아래에서'], ['t', '위에서'], ['l', '왼쪽에서'], ['r', '오른쪽에서']], (v) => setA((a) => { a.dir = v; }))) : null,
      el('div', { class: 'pane-btns' },
        el('button', { class: 'btn small', title: '앞으로 이동', onclick: () => run('moveAnim', -1) }, '▲'),
        el('button', { class: 'btn small', title: '뒤로 이동', onclick: () => run('moveAnim', 1) }, '▼'),
        el('button', { class: 'btn small danger', onclick: () => run('removeAnim') }, '제거'))));
  }
}

function selectionPane(body) {
  const slide = curSlide();
  body.append(el('div', { class: 'pane-btns' },
    el('button', { class: 'btn small', onclick: () => change(() => { for (const o of slide.objects) delete o.hidden; }) }, '모두 표시'),
    el('button', { class: 'btn small', onclick: () => change(() => { for (const o of slide.objects) o.hidden = true; }) }, '모두 숨기기')));
  const list = el('div', { class: 'sel-list' });
  [...slide.objects].reverse().forEach((o) => {
    const i = slide.objects.indexOf(o);
    const r = el('div', { class: `sel-row${S.sel.has(o.id) ? ' on' : ''}${o.hidden ? ' off' : ''}` },
      el('span', { class: 'sel-name' }, objLabel(o, i)),
      el('button', { class: 'eye', title: o.hidden ? '표시' : '숨기기', onclick: (e) => { e.stopPropagation(); change(() => { o.hidden = o.hidden ? undefined : true; }); } }, o.hidden ? '—' : '👁'));
    r.addEventListener('click', (e) => { if (e.ctrlKey) { if (S.sel.has(o.id)) S.sel.delete(o.id); else S.sel.add(o.id); } else S.sel = new Set([o.id]); emit('selection'); renderPane(); });
    r.addEventListener('dblclick', () => {
      const name = prompt('개체 이름', objLabel(o, i));
      if (name != null) change(() => { o.name = name.trim() || undefined; });
    });
    list.append(r);
  });
  body.append(list, el('div', { class: 'pane-btns' },
    el('button', { class: 'btn small', title: '앞으로 가져오기', onclick: () => run('bringForward') }, '▲'),
    el('button', { class: 'btn small', title: '뒤로 보내기', onclick: () => run('sendBackward') }, '▼')));
}

on('selection', () => { if (S.formatPane) renderPane(); });

