// 화면 둘레: 상태 표시줄 · 눈금자 · 탐색 창 · 작업 창 (스타일 · 메모 · 변경 내용 · 마크다운 · 맞춤법 · 접근성 · AI 연결 · 도움말)
import { S, run, register, emit, on, change, curSel, COMMAND_INFO, COMMANDS } from './state.js';
import { V, caretPage, setZoom, relayout } from './view.js';
import { setSel } from './editor.js';
import { pageNode } from './layout.js';
import { el, hydrateIcons, toast, openMenu } from './ui.js';
import { effPara, blockAt, headings, docStats, wordCount, paraPlain, walkParas, runLen, pxToCm, cm, orderPos, posEq, styleChain } from './model.js';
import { docToMd, mdToDoc } from './md.js';
import { selectedText } from './commands.js';

const $ = (id) => document.getElementById(id);
const P = { status: null, hruler: null, vruler: null, nav: null, pane: null, navTab: 'headings', navQuery: '', tabKind: 'left' };

export function initPanels() {
  P.status = $('status'); P.hruler = $('hruler'); P.vruler = $('vruler'); P.nav = $('navPane'); P.pane = $('pane');
  on('layout', () => { renderStatusSoon(); renderRulers(); if (S.showNav) renderNavSoon(); if (S.pane) renderPaneSoon(); });
  on('selection', () => { renderStatusSoon(); renderRulers(); if (S.showNav && P.navTab === 'headings') markNavHeading(); if (S.pane === 'comments') markActiveComment(); });
  on('scroll', () => { renderRulers(); if (S.showNav && P.navTab === 'pages') markNavPage(); });
  on('zoom', () => { renderStatus(); renderRulers(); });
  on('pane', () => renderPane());
  on('findHits', () => { if (S.showNav) { P.navTab = 'results'; renderNav(); } });
  on('wordCount', () => run('wordCountDialog'));
  P.hruler.addEventListener('pointerdown', rulerDown);
  P.hruler.addEventListener('dblclick', () => run('paragraphDialog', 'tabs'));
  P.vruler.addEventListener('dblclick', () => run('pageSetupDialog'));
}

export function renderPanels() {
  const app = $('app');
  app.dataset.view = S.view;
  document.body.classList.toggle('read-mode', S.view === 'read');
  P.nav.hidden = !S.showNav || S.view === 'read';
  P.pane.hidden = !S.pane;
  const rulerOn = S.showRuler && S.view !== 'read' && S.view !== 'web';
  P.hruler.hidden = !rulerOn;
  P.vruler.hidden = !rulerOn || S.view !== 'print';
  renderStatus();
  renderRulers();
  if (S.showNav) renderNav();
  if (S.pane) renderPane();
}

// ───────────── 상태 표시줄 ─────────────
let stT = 0;
let statsCache = { doc: null, words: 0 };
function renderStatusSoon() { clearTimeout(stT); stT = setTimeout(renderStatus, 120); }
const ZMAP = { toSlider: (z) => (z <= 1 ? ((z - 0.1) / 0.9) * 50 : 50 + ((z - 1) / 4) * 50), fromSlider: (v) => (v <= 50 ? 0.1 + (v / 50) * 0.9 : 1 + ((v - 50) / 50) * 4) };
export function renderStatus() {
  const st = P.status;
  if (!st) return;
  const res = S.layout;
  const total = res?.pages.length ?? 1;
  const pg = (caretPage() ?? 0) + 1;
  if (statsCache.doc !== S.doc) {
    // 큰 문서: 단어 수는 잠시 뒤 (입력을 막지 않음)
    const doc = S.doc;
    statsCache = { doc, words: statsCache.words };
    setTimeout(() => { if (statsCache.doc === doc) { statsCache.words = docStats(doc).words; paintStatus(); } }, S.doc.body.length > 3000 ? 600 : 0);
  }
  paintStatus(pg, total);
}
function paintStatus(pg = (caretPage() ?? 0) + 1, total = S.layout?.pages.length ?? 1) {
  const st = P.status;
  const sel = curSel();
  const selWords = posEq(sel.a, sel.f) ? null : wordCount(selectedText());
  const words = selWords != null ? `단어 ${statsCache.words.toLocaleString()}개 중 ${selWords.toLocaleString()}개` : `단어 ${statsCache.words.toLocaleString()}개`;
  const z = Math.round(S.zoom * 100);
  const slider = el('input', { type: 'range', class: 'zoom-slider', min: '0', max: '100', step: '1', value: String(ZMAP.toSlider(S.zoom)), title: '확대/축소' });
  slider.addEventListener('input', () => { let v = ZMAP.fromSlider(Number(slider.value)); if (Math.abs(v - 1) < 0.04) v = 1; setZoom(v); });
  const vb = (view, icon, title) => el('button', { class: `sb-btn${S.view === view ? ' on' : ''}`, title, 'data-icon': icon, onclick: () => run(view === 'read' ? 'viewRead' : view === 'web' ? 'viewWeb' : 'viewPrint') });
  st.replaceChildren(...[
    el('button', { class: 'sb-btn sb-item', title: '이동 (Ctrl+G) — 탐색 창', onclick: () => run('toggleNav') }, `${pg}/${total} 페이지`),
    el('button', { class: 'sb-btn sb-item', title: '단어 개수 (Ctrl+Shift+G)', onclick: () => run('wordCountDialog') }, words),
    el('button', { class: 'sb-btn sb-item', title: '맞춤법 검사 (F7)', onclick: () => run('spellCheck') }, '📖'),
    el('span', { class: 'sb-item' }, '한국어'),
    S.doc.settings?.track ? el('button', { class: 'sb-btn sb-item on', title: '변경 내용 추적 끄기', onclick: () => run('trackChanges') }, '변경 내용 추적: 켜짐') : null,
    S.merge?.preview ? el('span', { class: 'sb-item' }, `편지 병합 미리 보기 ${S.merge.index + 1}/${S.merge.rows.length}`) : null,
    S.hf ? el('span', { class: 'sb-item' }, S.hf.kind === 'header' ? '머리글 편집 중' : '바닥글 편집 중') : null,
    el('button', { class: 'sb-btn sb-item', title: '접근성 검사', onclick: () => run('accessibilityCheck') }, '접근성: 검사'),
    S.bridge?.connected ? el('button', { class: 'sb-btn sb-item', title: 'AI 연결 (MCP) 창', onclick: () => run('togglePane', 'mcp') }, '● AI 연결됨') : null,
    el('span', { class: 'sb-spacer' }),
    el('button', { class: `sb-btn${S.opts.focus ? ' on' : ''}`, title: '포커스 모드', onclick: () => run('focusMode') }, '◎'),
    vb('read', 'readMode', '읽기 모드'), vb('print', 'printLayout', '인쇄 모양'), vb('web', 'webLayout', '웹 모양'),
    el('button', { class: 'sb-btn', title: '축소', onclick: () => setZoom(Math.max(0.1, Math.round((S.zoom - 0.1) * 10) / 10)) }, '−'),
    slider,
    el('button', { class: 'sb-btn', title: '확대', onclick: () => setZoom(Math.min(5, Math.round((S.zoom + 0.1) * 10) / 10)) }, '+'),
    el('button', { class: 'sb-btn sb-item', title: '확대/축소 대화상자', onclick: () => run('zoomDialog') }, `${z}%`)].filter(Boolean));
  hydrateIcons(st);
}

// ───────────── 눈금자 ─────────────
let rulerState = null; // { page, holder, left, para }
function curParaInfo() {
  const pos = curSel().f;
  const b = blockAt(S.doc, pos.p);
  if (!b || b.type !== 'p') return null;
  return { pos, p: b, ep: effPara(S.doc, b), inCell: pos.p.length > 1 };
}
export function renderRulers() {
  if (!P.hruler || P.hruler.hidden) return;
  const res = S.layout;
  if (!res) return;
  const pi = Math.min(caretPage() ?? 0, res.pages.length - 1);
  const page = res.pages[pi];
  const holder = V.holders[pi];
  if (!page || !holder) return;
  const z = S.zoom;
  const hr = holder.getBoundingClientRect();
  const rr = P.hruler.getBoundingClientRect();
  const left = hr.left - rr.left;
  const info = curParaInfo();
  // 단 (여러 단이면 캐럿이 있는 단)
  let colL = page.bodyLeft; let colW = page.bodyW;
  const band = page.bands?.[0];
  if (band?.cols?.length > 1) {
    const cr = document.querySelector('.wcaret')?.getBoundingClientRect();
    const x = cr ? (cr.left - hr.left) / z : 0;
    for (const c of band.cols) if (x >= c.x - 4 && x <= c.x + c.w + 4) { colL = c.x; colW = c.w; }
  }
  // 표 칸 안: 칸 범위
  if (info?.inCell) {
    const ce = document.querySelector('.wcaret') && [...document.querySelectorAll(`.p[data-pid="${info.p.id}"]`)].map((e) => e.closest('td'))[0];
    if (ce) { const r = ce.getBoundingClientRect(); colL = (r.left - hr.left) / z + 7; colW = r.width / z - 14; }
  }
  rulerState = { page, pi, left, colL, colW, info };
  const strip = el('div', { class: 'ruler-strip', style: { left: `${left}px`, width: `${page.w * z}px` } });
  strip.append(el('div', { class: 'margin', style: { left: '0', width: `${colL * z}px` } }), el('div', { class: 'margin', style: { left: `${(colL + colW) * z}px`, right: '0' } }));
  const step = cm(0.25);
  for (let k = -Math.floor(colL / step); k * step + colL <= page.w; k++) {
    const x = (colL + k * step) * z;
    if (x < 0) continue;
    if (k % 4 === 0) { if (k !== 0) strip.append(el('span', { class: 'num', style: { left: `${x}px` } }, String(Math.abs(k / 4)))); } else strip.append(el('span', { class: `tick${k % 2 === 0 ? ' half' : ''}`, style: { left: `${x}px` } }));
  }
  if (info) {
    const ind = info.ep.ind ?? {};
    const l = ind.l ?? 0; const r = ind.r ?? 0;
    const first = l + (ind.first ?? 0) - (ind.hanging ?? 0);
    const mk = (cls, x, title) => el('span', { class: `mk ${cls}`, 'data-mk': cls, title, style: { left: `${(colL + x) * z}px` } });
    strip.append(mk('first', first, '첫 줄 들여쓰기'), mk('hang', l, '내어쓰기'), mk('left', l, '왼쪽 들여쓰기'), mk('right', colW - r, '오른쪽 들여쓰기'));
    for (const t of info.ep.tabs ?? []) if (t.align !== 'clear' && t.pos <= colW) strip.append(el('span', { class: `tabmk ${t.align ?? 'left'}`, 'data-tab': String(t.pos), title: `탭 ${pxToCm(t.pos).toFixed(2)}cm (끌어서 옮기기 · 밖으로 끌어 지우기)`, style: { left: `${(colL + t.pos) * z}px` } }));
  }
  const kindBtn = el('button', { class: 'ruler-kind', title: `탭 종류: ${{ left: '왼쪽', center: '가운데', right: '오른쪽', decimal: '소수점' }[P.tabKind]} (눌러서 바꾸기)`, onclick: (e) => { e.stopPropagation(); P.tabKind = { left: 'center', center: 'right', right: 'decimal', decimal: 'left' }[P.tabKind]; renderRulers(); } }, { left: '⌊', center: '⊥', right: '⌋', decimal: '⊥.' }[P.tabKind]);
  P.hruler.replaceChildren(kindBtn, strip);
  // 세로 눈금자
  if (!P.vruler.hidden) {
    const vr = P.vruler.getBoundingClientRect();
    const top = hr.top - vr.top;
    const vs = el('div', { class: 'ruler-strip v', style: { top: `${top}px`, height: `${page.h * z}px`, left: '3px', width: '16px' } });
    const bt = page.bodyTopEff ?? page.bodyTop;
    vs.append(el('div', { class: 'margin', style: { top: '0', left: '0', right: '0', height: `${bt * z}px` } }), el('div', { class: 'margin', style: { top: `${page.bodyBottom * z}px`, left: '0', right: '0', bottom: '0' } }));
    for (let k = -Math.floor(bt / step); k * step + bt <= page.h; k++) {
      const y = (bt + k * step) * z;
      if (y < 0 || k === 0) continue;
      if (k % 4 === 0) vs.append(el('span', { class: 'num v', style: { top: `${y - 6}px`, left: '3px' } }, String(Math.abs(k / 4))));
      else vs.append(el('span', { class: `tick v${k % 2 === 0 ? ' half' : ''}`, style: { top: `${y}px` } }));
    }
    P.vruler.replaceChildren(vs);
  }
}
function rulerDown(e) {
  const rs = rulerState;
  if (!rs?.info || e.button !== 0) return;
  const strip = P.hruler.querySelector('.ruler-strip');
  if (!strip) return;
  const sr = strip.getBoundingClientRect();
  const z = S.zoom;
  const xOf = (cx) => (cx - sr.left) / z - rs.colL;
  const snap = (v) => (e.altKey ? v : Math.round(v / cm(0.125)) * cm(0.125));
  const ind0 = { l: 0, r: 0, ...(rs.info.ep.ind ?? {}) };
  const tabs0 = (rs.info.ep.tabs ?? []).filter((t) => t.align !== 'clear');
  const mkEl = e.target.closest('[data-mk]');
  const tabEl = e.target.closest('[data-tab]');
  if (!mkEl && !tabEl) {
    // 빈 곳 클릭: 탭 추가 (본문 범위 안에서만)
    const x = snap(xOf(e.clientX));
    if (e.target.closest('.ruler-strip') && x > 0 && x < rs.colW) run('paragraphProps', { tabs: [...tabs0, { pos: x, align: P.tabKind }].sort((a, b) => a.pos - b.pos) });
    return;
  }
  e.preventDefault();
  const guide = el('div', { class: 'ruler-guide' });
  document.body.append(guide);
  const move = (ev) => {
    const x = snap(xOf(ev.clientX));
    Object.assign(guide.style, { left: `${ev.clientX}px` });
    if (mkEl) { mkEl.style.left = `${(rs.colL + x) * z}px`; if (mkEl.dataset.mk === 'left') strip.querySelector('[data-mk=hang]').style.left = mkEl.style.left; }
    else tabEl.style.left = `${(rs.colL + x) * z}px`;
  };
  const up = (ev) => {
    removeEventListener('pointermove', move); removeEventListener('pointerup', up); guide.remove();
    const x = snap(xOf(ev.clientX));
    if (mkEl) {
      const first = ind0.l + (ind0.first ?? 0) - (ind0.hanging ?? 0);
      const kind = mkEl.dataset.mk;
      let l = ind0.l; let f = first; let r = ind0.r;
      if (kind === 'first') f = x;
      else if (kind === 'hang') l = x;
      else if (kind === 'left') { const d = x - ind0.l; l = x; f = first + d; }
      else if (kind === 'right') r = rs.colW - x;
      const ind = { l: Math.max(-rs.colL, l), r: Math.max(-200, r) };
      if (f > l) ind.first = f - l; else if (f < l) ind.hanging = l - f;
      run('paragraphProps', { ind });
    } else {
      const old = Number(tabEl.dataset.tab);
      const off = Math.abs(ev.clientY - sr.top - sr.height / 2) > 24;
      const rest = tabs0.filter((t) => Math.abs(t.pos - old) > 0.5);
      run('paragraphProps', { tabs: off ? rest : [...rest, { ...tabs0.find((t) => Math.abs(t.pos - old) <= 0.5), pos: x }].sort((a, b) => a.pos - b.pos) });
    }
  };
  addEventListener('pointermove', move);
  addEventListener('pointerup', up);
}

// ───────────── 탐색 창 ─────────────
let navT = 0;
function renderNavSoon() { clearTimeout(navT); navT = setTimeout(renderNav, 250); }
export function renderNav() {
  const box = P.nav;
  if (!box || box.hidden) return;
  const q = el('input', { type: 'search', placeholder: '문서 검색', value: P.navQuery });
  q.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') { P.navQuery = q.value; if (q.value) { run('find', q.value, {}); P.navTab = 'results'; } renderNav(); setTimeout(() => box.querySelector('.nav-search input')?.focus(), 0); }
    if (e.key === 'Escape') emit('focusCanvas');
  });
  const tab = (k, l) => el('button', { class: P.navTab === k ? 'on' : '', onclick: () => { P.navTab = k; renderNav(); } }, l);
  const body = el('div', { class: 'nav-body' });
  box.replaceChildren(
    el('div', { class: 'nav-head' }, el('b', {}, '탐색'), el('button', { title: '닫기', onclick: () => run('toggleNav') }, '✕')),
    el('label', { class: 'nav-search' }, q, el('button', { class: 'sb-btn', title: '검색 옵션 · 찾기/바꾸기', onclick: (e) => openMenu(e.currentTarget, [{ label: '고급 찾기…', action: () => run('replaceDialog', 'find') }, { label: '바꾸기…', action: () => run('replaceDialog', 'replace') }, { label: '이동…', action: () => run('replaceDialog', 'goto') }]) }, '▾')),
    el('div', { class: 'nav-tabs' }, tab('headings', '제목'), tab('pages', '페이지'), tab('results', '결과')),
    body);
  if (P.navTab === 'headings') {
    const hs = headings(S.doc);
    if (!hs.length) body.append(el('p', { class: 'muted', style: { padding: '8px' } }, '문서에 제목이 없습니다. 홈 탭의 [제목 1]·[제목 2] 스타일을 적용하면 여기에 목차가 나타납니다.'));
    hs.forEach((h, k) => {
      const b = el('button', { class: `nav-h l${h.level}`, 'data-pid': h.p.id, draggable: 'true', style: { paddingLeft: `${8 + (h.level - 1) * 14}px` }, title: h.text, onclick: () => { setSel({ p: h.path, o: 0 }); emit('focusCanvas'); } }, h.text || '(빈 제목)');
      b.addEventListener('contextmenu', (e) => { e.preventDefault(); openMenu({ x: e.clientX, y: e.clientY }, [{ label: '수준 올리기', action: () => moveHeading(k, 'promote') }, { label: '수준 내리기', action: () => moveHeading(k, 'demote') }, { sep: true }, { label: '위로 이동 (구역째)', action: () => moveHeading(k, 'up') }, { label: '아래로 이동 (구역째)', action: () => moveHeading(k, 'down') }, { sep: true }, { label: '제목과 내용 삭제', action: () => moveHeading(k, 'delete') }, { label: '제목과 내용 선택', action: () => moveHeading(k, 'select') }]); });
      b.addEventListener('dragstart', (e) => { e.dataTransfer.setData('text/x-wiword-heading', String(k)); });
      b.addEventListener('dragover', (e) => { if (e.dataTransfer.types.includes('text/x-wiword-heading')) { e.preventDefault(); b.classList.add('drop'); } });
      b.addEventListener('dragleave', () => b.classList.remove('drop'));
      b.addEventListener('drop', (e) => { e.preventDefault(); b.classList.remove('drop'); const from = Number(e.dataTransfer.getData('text/x-wiword-heading')); if (from !== k) moveHeadingTo(from, k); });
      body.append(b);
    });
    markNavHeading();
  } else if (P.navTab === 'pages') {
    const wrap = el('div', { class: 'nav-pages' });
    const res = S.layout;
    const io = new IntersectionObserver((ents) => {
      for (const en of ents) {
        if (!en.isIntersecting) continue;
        const i = Number(en.target.dataset.i);
        const pg = res.pages[i];
        const sc = 110 / pg.w;
        const n = pageNode(S.doc, pg, res, { marks: false });
        n.style.transform = `scale(${sc})`; n.style.transformOrigin = '0 0';
        en.target.querySelector('.pbox').replaceChildren(n);
        io.unobserve(en.target);
      }
    }, { root: body });
    (res?.pages ?? []).forEach((pg, i) => {
      const t = el('div', { class: 'pthumb', 'data-i': String(i), onclick: () => run('goToPage', i + 1) }, el('div', { class: 'pbox', style: { height: `${(110 * pg.h) / pg.w}px` } }), el('div', {}, String(i + 1)));
      wrap.append(t);
      io.observe(t);
    });
    body.append(wrap);
    markNavPage();
  } else {
    const hits = S.findHits ?? [];
    if (!P.navQuery) body.append(el('p', { class: 'muted', style: { padding: '8px' } }, '위 상자에 찾을 글자를 입력하고 Enter 를 누르세요.'));
    else body.append(el('div', { class: 'muted', style: { padding: '4px 8px' } }, `${hits.length.toLocaleString()}개 결과`));
    hits.slice(0, 500).forEach((h, i) => {
      const p = blockAt(S.doc, h.p);
      const t = p ? paraPlain(p) : '';
      const a = Math.max(0, h.s - 30);
      const row = el('div', { class: 'nav-res', onclick: () => { setSel({ p: h.p, o: h.s }, { p: h.p, o: h.e }); for (const x of body.querySelectorAll('.nav-res.on')) x.classList.remove('on'); row.classList.add('on'); } },
        a > 0 ? '…' : '', t.slice(a, h.s), el('b', {}, t.slice(h.s, h.e)), t.slice(h.e, h.e + 50));
      row.dataset.i = String(i);
      body.append(row);
    });
  }
}
function markNavHeading() {
  const box = P.nav;
  if (!box) return;
  const top = curSel().f.p[0];
  let cur = null;
  for (const h of headings(S.doc)) if (h.path[0] <= top) cur = h.p.id;
  for (const b of box.querySelectorAll('.nav-h')) b.classList.toggle('on', b.dataset.pid === cur);
}
function markNavPage() {
  const pi = caretPage() ?? 0;
  for (const t of P.nav?.querySelectorAll('.pthumb') ?? []) t.classList.toggle('on', Number(t.dataset.i) === pi);
}
/** 제목 + 그 아래 내용 (다음 같은/상위 제목까지) 블록 범위 */
function headingSpan(hs, k) {
  const h = hs[k];
  let end = S.doc.body.length;
  for (let j = k + 1; j < hs.length; j++) if (hs[j].level <= h.level) { end = hs[j].path[0]; break; }
  return [h.path[0], end];
}
function moveHeading(k, op) {
  const hs = headings(S.doc);
  const [a, b] = headingSpan(hs, k);
  if (op === 'promote' || op === 'demote') {
    change((d) => { for (let i = a; i < b; i++) { const blk = S.doc.body[i]; const lv = hs.find((h) => h.path[0] === i)?.level; if (lv) { const nl = Math.max(1, Math.min(9, lv + (op === 'promote' ? -1 : 1))); d.block([i]).style = `Heading${nl}`; } } return S.sel; });
    return;
  }
  if (op === 'select') { const last = S.doc.body[b - 1]; setSel({ p: [a], o: 0 }, { p: [b - 1], o: last.type === 'p' ? paraLenOf(last) : 0 }); return; }
  if (op === 'delete') { change((d) => { const body = d.list([]); body.splice(a, b - a); if (!body.length) body.push({ type: 'p', id: `p${Date.now()}`, style: 'Normal', pPr: {}, runs: [] }); return { p: [Math.min(a, body.length - 1)], o: 0 }; }); return; }
  // 같은 수준 이웃과 자리 바꾸기
  const lv = hs[k].level;
  if (op === 'up') { let j = k - 1; while (j >= 0 && hs[j].level > lv) j--; if (j >= 0 && hs[j].level === lv) moveHeadingTo(k, j); }
  if (op === 'down') { let j = k + 1; while (j < hs.length && hs[j].level > lv) j++; if (j < hs.length && hs[j].level === lv) { const [, e2] = headingSpan(hs, j); moveBlocks(a, b, e2); } }
}
const paraLenOf = (p) => p.runs.reduce((n, r) => n + runLen(r), 0);
function moveHeadingTo(from, to) {
  const hs = headings(S.doc);
  const [a, b] = headingSpan(hs, from);
  const dest = hs[to].path[0];
  if (dest >= a && dest < b) return;
  moveBlocks(a, b, dest);
}
function moveBlocks(a, b, dest) {
  change((d) => {
    const body = d.list([]);
    const part = body.splice(a, b - a);
    const at = dest > a ? dest - (b - a) : dest;
    body.splice(at, 0, ...part);
    return { p: [at], o: 0 };
  });
}

// ───────────── 작업 창 ─────────────
let paneT = 0;
function renderPaneSoon() { clearTimeout(paneT); paneT = setTimeout(renderPane, 200); }
const PANES = {};
export function renderPane() {
  const box = P.pane;
  if (!box) return;
  box.hidden = !S.pane;
  if (!S.pane) { box.replaceChildren(); return; }
  const def = PANES[S.pane];
  if (!def) { box.replaceChildren(); return; }
  // 입력 중인 칸이 있으면 다시 그리지 않음
  if (box.contains(document.activeElement) && /^(TEXTAREA|INPUT|SELECT)$/.test(document.activeElement.tagName) && box.dataset.pane === S.pane && !def.always) return;
  const body = el('div', { class: `pane-body ${def.cls ?? ''}` });
  box.dataset.pane = S.pane;
  box.replaceChildren(el('div', { class: 'pane-head' }, el('b', {}, def.title), el('button', { title: '닫기', onclick: () => run('togglePane', S.pane) }, '✕')), body);
  def.render(body);
  hydrateIcons(box);
}

PANES.styles = {
  title: '스타일',
  render(body) {
    const cur = blockAt(S.doc, curSel().f.p);
    const all = Object.values(S.doc.styles).filter((s) => !s.hidden && s.type !== 'table' && s.type !== 'numbering').sort((a, b) => (a.order ?? 99) - (b.order ?? 99));
    const show = el('select', {}, [['rec', '권장'], ['used', '사용 중'], ['all', '모든 스타일']].map(([v, l]) => el('option', { value: v, selected: v === (P.styleShow ?? 'rec') }, l)));
    show.onchange = () => { P.styleShow = show.value; renderPane(); };
    const used = new Set();
    walkParas(S.doc.body, (p) => { used.add(p.style); for (const r of p.runs) if (r.cstyle) used.add(r.cstyle); });
    const list = all.filter((s) => (P.styleShow === 'all' ? true : P.styleShow === 'used' ? used.has(s.id) : s.quick || used.has(s.id)));
    for (const s of list) {
      const row = el('div', { class: `rev-row${cur?.style === s.id ? ' on' : ''}`, style: { display: 'flex', justifyContent: 'space-between' }, title: describeStyle(s), onclick: () => run('applyStyle', s.id) },
        el('span', {}, s.name), el('span', { class: 'muted' }, s.type === 'char' ? 'a' : '¶'));
      row.addEventListener('contextmenu', (e) => { e.preventDefault(); openMenu({ x: e.clientX, y: e.clientY }, [{ label: '수정…', action: () => run('modifyStyleDialog', s.id) }, { label: '선택 영역과 일치하도록 업데이트', action: () => run('updateStyleFromSel', s.id) }, { label: `모든 인스턴스 선택 (${countStyle(s.id)}개)`, action: () => run('selectStyleInstances', s.id) }, { sep: true }, { label: '삭제', disabled: ['Normal', 'Heading1', 'Heading2', 'Heading3'].includes(s.id), action: () => run('deleteStyle', s.id) }]); });
      body.append(row);
    }
    body.append(el('div', { class: 'pane-sec' }, el('label', { class: 'pane-row' }, el('span', {}, '표시'), show)),
      el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('newStyleDialog') }, '새 스타일'), el('button', { class: 'btn', onclick: () => run('clearFormat') }, '모두 지우기')));
  },
};
function describeStyle(s) {
  const parts = [s.name, s.type === 'char' ? '글자 스타일' : '단락 스타일'];
  if (s.basedOn) parts.push(`기준: ${S.doc.styles[s.basedOn]?.name ?? s.basedOn}`);
  const r = s.rPr ?? {};
  if (r.sz) parts.push(`${r.sz}pt`);
  if (r.b) parts.push('굵게');
  if (r.i) parts.push('기울임꼴');
  if (r.color) parts.push(r.color);
  return parts.join(' · ');
}
function countStyle(id) { let n = 0; walkParas(S.doc.body, (p) => { if (p.style === id || p.runs.some((r) => r.cstyle === id)) n++; }); return n; }

PANES.comments = {
  title: '메모',
  render(body) {
    const list = S.doc.comments ?? [];
    body.append(el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('newComment') }, '+ 새 메모'), el('button', { class: 'btn', onclick: () => run('nextComment', -1) }, '◀ 이전'), el('button', { class: 'btn', onclick: () => run('nextComment', 1) }, '다음 ▶')));
    if (!list.length) body.append(el('p', { class: 'muted' }, '메모가 없습니다. 글을 선택하고 [검토 › 새 메모] (Ctrl+Alt+M) 를 누르세요.'));
    for (const c of list) {
      const ta = el('textarea', { rows: 2, placeholder: '메모 입력…' }, c.text ?? '');
      ta.addEventListener('input', () => run('editComment', c.id, ta.value));
      ta.addEventListener('keydown', (e) => e.stopPropagation());
      const reply = el('input', { type: 'text', placeholder: '회신…', style: { width: '100%', marginTop: '4px' } });
      reply.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter' && reply.value.trim()) { run('replyComment', c.id, reply.value.trim()); reply.value = ''; } });
      const card = el('div', { class: `cmt-card${c.done ? ' done' : ''}${S.activeComment === c.id ? ' on' : ''}`, 'data-cid': c.id, onclick: (e) => { if (e.target.closest('textarea,input,button')) return; gotoComment(c.id); } },
        el('div', { class: 'who' }, el('span', {}, c.author ?? '사용자'), el('span', {},
          el('button', { class: 'sb-btn', title: c.done ? '다시 열기' : '해결됨으로 표시', onclick: () => run('resolveComment', c.id) }, c.done ? '↺' : '✓'),
          el('button', { class: 'sb-btn', title: '삭제', onclick: () => run('deleteComment', c.id) }, '🗑'))),
        el('div', { class: 'when' }, c.date ? new Date(c.date).toLocaleString('ko-KR') : ''),
        el('div', { class: 'muted', style: { fontStyle: 'italic', margin: '2px 0' } }, `“${anchorText(c.id).slice(0, 60)}”`),
        ta,
        ...(c.replies ?? []).map((r) => el('div', { class: 'reply' }, el('b', {}, r.author ?? ''), ' ', el('span', { class: 'when' }, r.date ? new Date(r.date).toLocaleDateString('ko-KR') : ''), el('div', {}, r.text))),
        reply);
      body.append(card);
    }
    if (P.focusComment) { const t = body.querySelector(`[data-cid="${P.focusComment}"] textarea`); P.focusComment = null; setTimeout(() => t?.focus(), 0); }
  },
};
function anchorText(id) {
  let s = ''; let on = false;
  walkParas(S.doc.body, (p) => { for (const r of p.runs) { if (r.obj === 'mark' && r.kind === 'cmt' && r.id === id) on = r.edge === 'start'; else if (on && !r.obj) s += r.text; } if (on) s += ' '; });
  return s.trim();
}
function gotoComment(id) {
  walkParas(S.doc.body, (p, path) => { let o = 0; for (const r of p.runs) { if (r.obj === 'mark' && r.kind === 'cmt' && r.id === id && r.edge === 'start') { S.activeComment = id; setSel({ p: path, o }); } o += runLen(r); } });
  markActiveComment();
}
function markActiveComment() { for (const c of P.pane?.querySelectorAll('.cmt-card') ?? []) c.classList.toggle('on', c.dataset.cid === S.activeComment); }
on('editComment', (id) => { P.focusComment = id; S.activeComment = id; renderPane(); });

PANES.revisions = {
  title: '변경 내용',
  render(body) {
    const items = [];
    walkParas(S.doc.body, (p, path) => { let o = 0; let cur = null; for (const r of p.runs) { const n = runLen(r); const k = r.ins ? 'ins' : r.del ? 'del' : null; if (k) { const who = (r.ins ?? r.del).author ?? ''; if (cur && cur.k === k && cur.who === who && cur.e === o) { cur.e += n; cur.text += r.text ?? ''; } else { cur = { k, who, date: (r.ins ?? r.del).date, p: path, s: o, e: o + n, text: r.text ?? '' }; items.push(cur); } } else cur = null; o += n; } });
    body.append(el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('acceptAll') }, '모두 적용'), el('button', { class: 'btn', onclick: () => run('rejectAll') }, '모두 취소')),
      el('div', { class: 'muted' }, `수정 ${items.length}개 (삽입 ${items.filter((x) => x.k === 'ins').length} · 삭제 ${items.filter((x) => x.k === 'del').length})`));
    for (const it of items) body.append(el('div', { class: 'rev-row', onclick: () => setSel({ p: it.p, o: it.s }, { p: it.p, o: it.e }) }, el('span', { class: `k ${it.k}` }, it.k === 'ins' ? '삽입: ' : '삭제: '), el('span', {}, it.text.slice(0, 80) || '(개체)'), el('div', { class: 'when muted' }, `${it.who} ${it.date ? new Date(it.date).toLocaleString('ko-KR') : ''}`)));
  },
};

// 마크다운 (양방향): 문서 → 마크다운 글, 글을 고치면 잠시 뒤 문서에 반영
PANES.markdown = {
  title: '마크다운', cls: 'md-pane',
  render(body) {
    const ta = el('textarea', { spellcheck: 'false' });
    ta.value = docToMd(S.doc, { images: 'ref' });
    P.mdDoc = S.doc;
    let t = 0;
    ta.addEventListener('keydown', (e) => e.stopPropagation());
    ta.addEventListener('input', () => { clearTimeout(t); t = setTimeout(() => applyMd(ta.value), 500); });
    const live = el('input', { type: 'checkbox', checked: S.opts.mdLive !== false, onchange: (e) => { S.opts.mdLive = e.target.checked; } });
    body.append(el('div', { class: 'md-tools' },
      el('button', { class: 'btn', onclick: () => { navigator.clipboard?.writeText(docToMd(S.doc, { images: 'none' })); toast('마크다운을 복사했습니다'); } }, '복사'),
      el('button', { class: 'btn', onclick: () => run('saveAs', 'md') }, '.md 로 저장'),
      el('button', { class: 'btn', onclick: () => { ta.value = docToMd(S.doc, { images: 'ref' }); } }, '새로 고침'),
      el('label', { class: 'pane-check' }, live, '입력하면 바로 문서에 반영')), ta,
    el('p', { class: 'muted' }, '# 제목 · **굵게** · *기울임* · - 목록 · 1. 번호 · > 인용 · ``` 코드 · | 표 | · ![그림](주소) · [^1] 각주 · $$수식$$'));
  },
};
function applyMd(text) {
  if (S.opts.mdLive === false) return;
  const cur = S.doc;
  const nd = mdToDoc(text, { doc: { ...cur, body: [], lists: { ...cur.lists }, notes: { foot: {}, end: { ...(cur.notes?.end ?? {}) } }, props: { ...cur.props } } });
  change((d) => { d.doc.body = nd.body; d.doc.lists = nd.lists; d.doc.notes = nd.notes; d.doc.styles = nd.styles; d.doc.media = nd.media; return { p: [0], o: 0 }; }, { key: 'md' });
  P.mdDoc = S.doc;
}
on('layout', () => {
  // 문서가 다른 곳에서 바뀌면 마크다운 창 글도 새로 (입력 중이 아니면)
  if (S.pane !== 'markdown' || P.mdDoc === S.doc) return;
  const ta = P.pane?.querySelector('.md-pane textarea');
  if (ta && document.activeElement !== ta) { ta.value = docToMd(S.doc, { images: 'ref' }); P.mdDoc = S.doc; }
});

PANES.spelling = {
  title: '편집기 — 맞춤법',
  render(body) {
    const list = S.spellIssues ?? [];
    body.append(el('div', { class: 'muted' }, list.length ? `제안 ${list.length}개` : '문제가 없습니다 ✓'), el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('spellCheck') }, '다시 검사')));
    list.slice(0, 300).forEach((it, i) => body.append(el('div', { class: 'rev-row' },
      el('div', { onclick: () => setSel({ p: it.p, o: it.s }, { p: it.p, o: it.e }) }, el('span', { class: 'k del' }, it.word), ' → ', el('b', {}, it.sug === ' ' ? '(공백 하나)' : it.sug), el('div', { class: 'muted' }, it.why)),
      el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('applySpell', i) }, '바꾸기'), el('button', { class: 'btn', onclick: () => { S.spellIssues.splice(i, 1); renderPane(); } }, '무시')))));
  },
};
PANES.a11y = {
  title: '접근성',
  render(body) {
    const list = S.a11yIssues ?? [];
    body.append(el('div', { class: 'muted' }, list.length ? `검사 결과 ${list.length}개` : '문제가 없습니다 ✓'), el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('accessibilityCheck') }, '다시 검사')));
    for (const g of ['오류', '경고', '팁']) {
      const xs = list.filter((x) => x.kind === g);
      if (!xs.length) continue;
      body.append(el('div', { class: 'pane-group' }, `${g} (${xs.length})`));
      for (const x of xs) body.append(el('div', { class: 'rev-row', onclick: () => { if (x.path) setSel({ p: x.path, o: 0 }); else if (/속성/.test(x.msg)) run('backstage', 'info'); } }, x.msg));
    }
  },
};

// AI 연결 (MCP)
PANES.mcp = {
  title: 'AI 연결 (MCP)', always: true,
  render(body) {
    const b = S.bridge ?? {};
    const origin = location.origin.startsWith('http') ? location.origin : 'http://127.0.0.1:5180';
    const cfg = JSON.stringify({ mcpServers: { wiword: { command: 'node', args: ['<WIWORD 폴더>/mcp.js'], env: { WIWORD_URL: origin } } } }, null, 2);
    body.append(
      el('div', { class: 'pane-sec' }, el('h4', {}, '상태'),
        el('p', {}, b.connected ? `● 연결됨 — AI 가 이 창의 문서를 읽고 고칠 수 있습니다 (요청 ${b.calls ?? 0}번)` : b.available === false ? '○ 서버 없이 열려 있어 실시간 연결을 쓸 수 없습니다. (npm start 로 실행하면 됩니다)' : '○ 연결 대기 중'),
        el('div', { class: 'pane-btns' }, el('button', { class: 'btn', onclick: () => run('bridgeConnect') }, b.connected ? '다시 연결' : '연결'), b.connected ? el('button', { class: 'btn', onclick: () => run('bridgeDisconnect') }, '끊기') : null),
        el('label', { class: 'pane-check' }, el('input', { type: 'checkbox', checked: S.opts.bridgeAsk === true, onchange: (e) => { S.opts.bridgeAsk = e.target.checked; } }), 'AI 가 문서를 고치기 전에 물어보기')),
      el('div', { class: 'pane-sec' }, el('h4', {}, 'Claude Code 에 연결'),
        el('div', { class: 'code-box' }, 'claude mcp add wiword -- node <WIWORD 폴더>/mcp.js'),
        el('h4', { style: { marginTop: '10px' } }, 'Claude Desktop · 다른 MCP 클라이언트'),
        el('div', { class: 'code-box' }, cfg),
        el('p', { class: 'muted' }, 'MCP 서버 혼자서도 .docx · .hwp · .hwpx · .md · .pdf 파일을 열고 고치고 저장합니다. WIWORD 가 열려 있으면 이 창의 문서를 실시간으로 다룹니다.')),
      el('div', { class: 'pane-sec' }, el('h4', {}, '도구'), el('div', { class: 'muted' }, 'list_commands · execute · open_document · new_document · get_text · get_markdown · set_markdown · inspect · find · replace · insert_text · insert_markdown · apply_style · save_document · convert · comments · track_changes')),
      el('div', { class: 'pane-sec' }, el('h4', {}, '기록'), el('div', { class: 'mcp-log' }, (b.log ?? []).slice(-200).join('\n') || '(아직 없음)')));
  },
};

PANES.help = {
  title: '도움말',
  render(body) {
    const q = el('input', { type: 'search', placeholder: '명령 · 단축키 검색', style: { width: '100%' } });
    const list = el('div', {});
    const draw = () => {
      const s = q.value.trim().toLowerCase();
      list.replaceChildren();
      for (const [k, txt] of SHORTCUTS) if (!s || k.toLowerCase().includes(s) || txt.includes(s)) list.append(el('div', { class: 'rev-row', style: { display: 'flex', justifyContent: 'space-between' } }, el('span', {}, txt), el('b', {}, k)));
      if (s) for (const [name, info] of Object.entries(COMMAND_INFO)) if ((info.label ?? '').toLowerCase().includes(s) || name.toLowerCase().includes(s)) list.append(el('div', { class: 'rev-row', onclick: () => run(name) }, `▶ ${info.label ?? name}`));
    };
    q.addEventListener('input', draw);
    q.addEventListener('keydown', (e) => e.stopPropagation());
    body.append(el('p', {}, 'WIWORD 는 Microsoft Word 와 같은 화면 · 단축키로 쓰는 문서 편집기입니다. .docx · .hwp · .hwpx · .md · .pdf · .html · .txt 를 열고 .docx · .hwpx · .md · .pdf 로 저장합니다.'), q, list);
    draw();
  },
};
export const SHORTCUTS = [
  ['Ctrl+N', '새 문서'], ['Ctrl+O', '열기'], ['Ctrl+S', '저장'], ['F12', '다른 이름으로 저장'], ['Ctrl+P', '인쇄'], ['Ctrl+W', '닫기'],
  ['Ctrl+Z', '실행 취소'], ['Ctrl+Y', '다시 실행 · 반복'], ['F4', '마지막 작업 반복'], ['Ctrl+X / C / V', '잘라내기 · 복사 · 붙여넣기'], ['Ctrl+Shift+V', '텍스트만 붙여넣기'], ['Ctrl+Shift+C / V', '서식 복사 · 붙여넣기'],
  ['Ctrl+A', '모두 선택'], ['Ctrl+F', '탐색 창 검색'], ['Ctrl+H', '바꾸기'], ['Ctrl+G · F5', '이동'], ['Shift+F5', '마지막 편집 위치'],
  ['Ctrl+B / I / U', '굵게 · 기울임꼴 · 밑줄'], ['Ctrl+Shift+D', '이중 밑줄'], ['Ctrl+D', '글꼴 대화상자'], ['Ctrl+Shift+> / <', '글꼴 크기 크게 · 작게'], ['Ctrl+] / [', '글꼴 크기 1pt'], ['Ctrl+= / Ctrl+Shift+=', '아래 · 위 첨자'], ['Shift+F3', '대/소문자 바꾸기'], ['Ctrl+Space', '글자 서식 지우기'],
  ['Ctrl+E / L / R / J', '가운데 · 왼쪽 · 오른쪽 · 양쪽 맞춤'], ['Ctrl+1 / 2 / 5', '줄 간격 1 · 2 · 1.5'], ['Ctrl+0', '단락 앞 간격 12pt 켜기/끄기'], ['Ctrl+M / Ctrl+Shift+M', '들여쓰기 · 내어쓰기'], ['Ctrl+T', '내어쓰기 (첫 줄 밖으로)'], ['Ctrl+Q', '단락 서식 지우기'],
  ['Ctrl+Alt+1 / 2 / 3', '제목 1 · 2 · 3'], ['Ctrl+Shift+N', '표준 스타일'], ['Ctrl+Shift+L', '글머리 기호'], ['Alt+Ctrl+Shift+S', '스타일 창'], ['Ctrl+Shift+8', '편집 기호 표시'],
  ['Ctrl+Enter', '페이지 나누기'], ['Ctrl+Shift+Enter', '단 나누기'], ['Shift+Enter', '줄 바꿈'], ['Ctrl+K', '하이퍼링크'], ['Ctrl+Alt+F', '각주'], ['Ctrl+Alt+D', '미주'], ['Ctrl+Alt+M', '새 메모'], ['Ctrl+Shift+E', '변경 내용 추적'], ['Alt+=', '수식'], ['Alt+Shift+D / T', '날짜 · 시간 넣기'],
  ['F7', '맞춤법 검사'], ['Ctrl+Shift+G', '단어 개수'], ['F9', '필드 업데이트'], ['Alt+Shift+↑ / ↓', '단락 위 · 아래로 옮기기'], ['Ctrl+↑ / ↓', '이전 · 다음 단락'], ['Ctrl+← / →', '이전 · 다음 낱말'], ['Ctrl+Home / End', '문서 처음 · 끝'], ['Ctrl+Backspace / Delete', '낱말 지우기'],
  ['Tab / Shift+Tab (표)', '다음 · 이전 칸'], ['Alt · F10', '키 팁'], ['Ctrl+F1', '리본 접기'], ['Ctrl+휠', '확대/축소'], ['Ctrl+Alt+P / Ctrl+Alt+O / Ctrl+Alt+N', '인쇄 모양 · 개요 · 초안'],
];

// ───────────── 명령 ─────────────
register({
  togglePane: (name) => { S.pane = S.pane === name ? null : name; renderPanels(); if (S.pane === 'mcp' && !S.bridge?.connected) run('bridgeConnect'); },
  stylesPane: () => run('togglePane', 'styles'),
  markdownPane: () => run('togglePane', 'markdown'),
  mcpPane: () => run('togglePane', 'mcp'),
  helpPane: () => run('togglePane', 'help'),
  findPane: () => { S.showNav = true; P.navTab = 'results'; renderPanels(); setTimeout(() => { const i = P.nav.querySelector('.nav-search input'); i?.focus(); i?.select(); }, 0); },
  toggleGridlines: () => { S.opts.gridlines = !S.opts.gridlines; document.body.classList.toggle('gridlines', !!S.opts.gridlines); emit('ui'); },
  focusMode: () => { S.opts.focus = !S.opts.focus; document.body.classList.toggle('focus-mode', !!S.opts.focus); if (S.opts.focus) document.documentElement.requestFullscreen?.().catch(() => {}); else if (document.fullscreenElement) document.exitFullscreen?.(); emit('ui'); },
  updateStyleFromSel: (id) => {
    const b = blockAt(S.doc, curSel().f.p);
    if (!b || b.type !== 'p') return;
    const r = b.runs.find((x) => !x.obj) ?? {};
    const rPr = {}; for (const k of ['font', 'fontEa', 'sz', 'b', 'i', 'u', 'color']) if (r[k] != null) rPr[k] = r[k];
    change((d) => { const st = d.prop('styles'); st[id] = { ...st[id], rPr: { ...st[id].rPr, ...rPr }, pPr: { ...st[id].pPr, ...b.pPr } }; return S.sel; });
  },
  selectStyleInstances: (id) => {
    const paths = [];
    walkParas(S.doc.body, (p, path) => { if (p.style === id) paths.push(path); });
    if (!paths.length) return;
    const last = blockAt(S.doc, paths[paths.length - 1]);
    setSel({ p: paths[0], o: 0 }, { p: paths[paths.length - 1], o: paraLenOf(last) });
    toast(`${paths.length}개 단락 (처음 ~ 끝 범위를 선택했습니다)`);
  },
  deleteStyle: (id) => change((d) => { const st = d.prop('styles'); const base = st[id]?.basedOn ?? 'Normal'; delete st[id]; walkParas(d.doc.body, (p, path) => { if (p.style === id) d.block(path).style = base; }); return S.sel; }),
}, { togglePane: { label: '작업 창 열기/닫기', args: ['name'] }, stylesPane: { label: '스타일 창' }, markdownPane: { label: '마크다운 창' }, focusMode: { label: '포커스 모드' } });

