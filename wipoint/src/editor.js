// 편집 화면(가운데 슬라이드): 선택 · 이동 · 크기 · 회전 · 스마트 가이드 · 도형 그리기 · 글 편집
import { S, curSlide, selObjects, objById, change, emit, on, run } from './state.js';
import { slideHtml, textHtml } from './render.js';
import { groupMembers, bbox, rotatedBox, newShape, newTextBox, formatRange, rangeRunProp, uid } from './model.js';
import { domToBody, getOffsets, setOffsets, wordAt, placeCaretAtPoint, selectAllIn } from './textedit.js';
import { isLineShape } from './shapes.js';
import { el } from './ui.js';

let stage;
let host;
let layer;
let overlay;
let txi = null;           // 편집 중인 글 요소
let editBase = null;      // 편집 시작 때의 글 (data 없는 부분 기본값)
let lastOffs = null;      // 리본 입력 칸으로 초점이 옮겨가도 기억하는 글 선택 위치
document.addEventListener('selectionchange', () => { if (txi?.isConnected) { const o = getOffsets(txi); if (o) lastOffs = o; } });
const offsetsNow = () => (txi ? getOffsets(txi) ?? lastOffs : null);

export function initEditor(stageEl) {
  stage = stageEl;
  host = el('div', { class: 'slide-host' });
  layer = el('div', { class: 'slide-layer' });
  overlay = el('div', { class: 'overlay' });
  host.append(layer, overlay);
  stage.append(el('div', { class: 'stage-inner' }, host));
  stage.addEventListener('pointerdown', onStageDown);
  host.addEventListener('pointerdown', onPointerDown);
  host.addEventListener('dblclick', onDblClick);
  host.addEventListener('contextmenu', (e) => { e.preventDefault(); emit('canvasMenu', { e, hit: hitObject(e) }); });
  stage.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    setZoom(S.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
  }, { passive: false });
  new ResizeObserver(() => { if (S.fitZoom) fitZoom(); }).observe(stage);
  document.execCommand?.('defaultParagraphSeparator', false, 'div');
}

// ───────────── 확대/축소 ─────────────
export function fitZoom() {
  if (!stage) return;
  const { w, h } = S.pres.size;
  const aw = stage.clientWidth - 48;
  const ah = stage.clientHeight - 40;
  if (aw <= 0 || ah <= 0) return;
  S.zoom = Math.max(0.1, Math.min(aw / w, ah / h));
  S.fitZoom = true;
  layout();
  emit('zoom');
}
export function setZoom(z) {
  S.zoom = Math.max(0.1, Math.min(4, z));
  S.fitZoom = false;
  layout();
  emit('zoom');
}
function layout() {
  const { w, h } = S.pres.size;
  host.style.width = `${w * S.zoom}px`;
  host.style.height = `${h * S.zoom}px`;
  layer.style.transform = `scale(${S.zoom})`;
  layer.style.width = `${w}px`;
  layer.style.height = `${h}px`;
  renderOverlay();
}

// ───────────── 그리기 ─────────────
export function renderCanvas() {
  if (!host) return;
  const slide = curSlide();
  if (!slide) { layer.innerHTML = ''; return; }
  let offs = null;
  if (S.editing && txi?.isConnected) offs = getOffsets(txi);
  layer.innerHTML = slideHtml(S.pres, slide, { prompt: true, index: S.cur, editId: S.editing?.id });
  layer.classList.toggle('grid', S.showGrid);
  layer.classList.toggle('guides', S.showGuides);
  layout();
  if (S.editing) attachEditor(offs);
}

const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
const HPOS = { nw: [0, 0], n: [0.5, 0], ne: [1, 0], e: [1, 0.5], se: [1, 1], s: [0.5, 1], sw: [0, 1], w: [0, 0.5] };

/** 선택 상자: 그룹 전체를 골랐으면 그룹 상자 하나, 아니면 개체마다 */
function selectionFrames() {
  const objs = selObjects();
  if (!objs.length) return [];
  const grps = new Set(objs.map((o) => o.grp ?? null));
  if (grps.size === 1 && !grps.has(null)) {
    const all = curSlide().objects.filter((o) => o.grp === objs[0].grp);
    if (all.length === objs.length) return [{ box: bbox(objs), rot: 0, group: true, objs }];
  }
  return objs.map((o) => ({ box: { x: o.x, y: o.y, w: o.w, h: o.h }, rot: o.rot ?? 0, objs: [o], line: o.type === 'shape' && isLineShape(o.shape) }));
}

export function renderOverlay() {
  if (!overlay) return;
  overlay.innerHTML = '';
  const z = S.zoom;
  const frames = S.cropping ? [] : selectionFrames();
  for (const [i, f] of frames.entries()) {
    const { x, y, w, h } = f.box;
    const fr = el('div', { class: `sel-frame${f.group ? ' group' : ''}${S.editing ? ' editing' : ''}`, style: { left: `${x * z}px`, top: `${y * z}px`, width: `${w * z}px`, height: `${h * z}px`, transform: f.rot ? `rotate(${f.rot}deg)` : '' } });
    fr.dataset.frame = String(i);
    if (frames.length <= 12) {
      const handles = f.line ? ['nw', 'se'] : HANDLES;
      for (const hd of handles) {
        const [hx, hy] = HPOS[hd];
        fr.append(el('div', { class: `hd hd-${hd}`, dataset: { h: hd }, style: { left: `${hx * 100}%`, top: `${hy * 100}%` } }));
      }
      if (!f.line && frames.length === 1) fr.append(el('div', { class: 'hd-rot', dataset: { h: 'rot' }, title: '회전' }));
    }
    overlay.append(fr);
  }
  tableOverlay(z);
  cropOverlay(z);
  // 메모 표시 (말풍선)
  if (S.showComments !== false) {
    for (const [k, c] of (curSlide()?.comments ?? []).entries()) {
      const pin = el('div', { class: `cm-pin${c.done ? ' done' : ''}`, title: `${c.author ?? ''}: ${c.text}`, dataset: { cm: c.id }, style: { left: `${c.x * z}px`, top: `${c.y * z}px` } }, String(k + 1));
      pin.addEventListener('pointerdown', (e) => { e.stopPropagation(); e.preventDefault(); S.formatPane = 'comments'; emit('openPane'); });
      overlay.append(pin);
    }
  }
  // 개체 틀 점선 (비어 있는 개체 틀은 render 에서 안내 글로)
}

function showGuides(lines) {
  overlay.querySelectorAll('.guide').forEach((g) => g.remove());
  const z = S.zoom;
  for (const g of lines) {
    overlay.append(el('div', { class: `guide ${g.v ? 'v' : 'h'}`, style: g.v ? { left: `${g.at * z}px`, top: `${g.from * z}px`, height: `${(g.to - g.from) * z}px` } : { top: `${g.at * z}px`, left: `${g.from * z}px`, width: `${(g.to - g.from) * z}px` } }));
  }
}

// ───────────── 좌표 · 맞히기 ─────────────
function toSlide(e) {
  const r = host.getBoundingClientRect();
  return { x: (e.clientX - r.left) / S.zoom, y: (e.clientY - r.top) / S.zoom };
}
function hitObject(e) {
  const ob = e.target.closest?.('.slide-layer .ob:not(.decor)');
  return ob ? objById(ob.dataset.id) : null;
}

/** 클릭한 개체 → 선택할 id 목록 (그룹은 처음엔 그룹 전체, 그룹이 선택돼 있으면 그 안의 개체) */
function selectionFor(o, additive) {
  const slide = curSlide();
  if (!o.grp) return [o.id];
  const members = groupMembers(slide, o);
  const groupSelected = members.every((m) => S.sel.has(m.id)) && S.sel.size === members.length;
  if (groupSelected && !additive) return [o.id];
  if (S.sel.size === 1 && S.sel.has(o.id)) return [o.id];
  return members.map((m) => m.id);
}

// ───────────── 마우스 ─────────────
function onStageDown(e) {
  if (e.target !== stage && !e.target.classList?.contains('stage-inner')) return;
  if (S.editing) endEdit();
  if (S.sel.size) { S.sel.clear(); emit('selection'); renderOverlay(); }
}

let drag = null;

function onPointerDown(e) {
  if (e.button !== 0) return;
  if (S.cropping && !e.target.closest('.crop-win, .crop-ghost')) endCrop();
  const p = toSlide(e);
  // 도형 그리기 모드
  if (S.drawShape) { startDraw(e, p); return; }
  const handle = e.target.dataset?.h;
  if (handle) { e.preventDefault(); startHandleDrag(e, p, handle, Number(e.target.closest('.sel-frame')?.dataset.frame ?? 0)); return; }
  // 표 칸 범위: Shift+클릭 / 칸에서 칸으로 끌기
  const tdHit = e.target.closest?.('[data-cell]');
  let tblObj = tdHit ? hitObject(e) : null;
  // 표 가장자리(테두리 근처)를 누르면 표 옮기기
  if (tblObj?.type === 'table') { const m = 6 / S.zoom; if (p.x - tblObj.x < m || tblObj.x + tblObj.w - p.x < m || p.y - tblObj.y < m || tblObj.y + tblObj.h - p.y < m) tblObj = null; }
  if (tblObj?.type === 'table' && (S.editing?.id === tblObj.id || S.cellRange?.id === tblObj.id || S.sel.has(tblObj.id))) {
    const [r, c] = tdHit.dataset.cell.split(',').map(Number);
    const anchor = S.editing?.id === tblObj.id && S.editing.cell ? S.editing.cell : S.cellRange?.id === tblObj.id ? [S.cellRange.ar, S.cellRange.ac] : null;
    if (e.shiftKey && anchor) { e.preventDefault(); setCellRange(tblObj, anchor[0], anchor[1], r, c); return; }
    const inText = S.editing && txi && txi.contains(e.target);
    startCellDrag(e, tblObj, r, c, !inText);
    // 표 안을 누르면 칸 편집 · 끌면 칸 범위 (표 옮기기는 테두리를 끌어서 — PowerPoint 와 같음)
    return;
  } else if (S.cellRange) { S.cellRange = null; }
  // 편집 중인 글 안 클릭은 브라우저에 맡김
  if (S.editing && txi && txi.contains(e.target)) return;
  const o = hitObject(e);
  const frameHit = e.target.closest('.sel-frame');
  if (S.editing) endEdit();
  if (!o) {
    if (frameHit) { startMove(e, p); return; }
    // 빈 곳: 영역 선택
    if (!e.shiftKey && !e.ctrlKey) { S.sel.clear(); emit('selection'); }
    startBand(e, p);
    return;
  }
  // 그림 개체 틀의 아이콘 → 그림 넣기
  if (o.type === 'image' && o.ph === 'pic' && !o.media) { S.sel = new Set([o.id]); emit('selection'); run('insertPicture', { into: o.id }); return; }
  const additive = e.shiftKey || e.ctrlKey;
  const ids = selectionFor(o, additive);
  if (additive) {
    const allIn = ids.every((id) => S.sel.has(id));
    for (const id of ids) { if (allIn) S.sel.delete(id); else S.sel.add(id); }
  } else if (!ids.every((id) => S.sel.has(id)) || S.sel.size !== ids.length) {
    // 이미 여러 개를 고른 상태에서 그중 하나를 끌면 그대로 함께 옮김
    if (!(S.sel.has(o.id) && S.sel.size > 1 && !o.grp)) S.sel = new Set(ids);
  }
  emit('selection');
  renderOverlay();
  // 글 상자 · 개체 틀은 글자 위를 누르면 바로 편집 (PowerPoint 와 같음)
  const textTarget = (o.txBox || o.ph) && o.type === 'shape' && S.sel.size === 1 && !o.grp;
  startMove(e, p, textTarget ? () => startEdit(o, { x: e.clientX, y: e.clientY }) : null);
}

// ───────────── 표: 칸 범위 · 열 너비/행 높이 끌기 ─────────────
function setCellRange(t, ar, ac, br, bc) {
  if (S.editing) endEdit();
  S.sel = new Set([t.id]);
  S.cellRange = { id: t.id, ar, ac, r1: Math.min(ar, br), c1: Math.min(ac, bc), r2: Math.max(ar, br), c2: Math.max(ac, bc) };
  // 병합된 칸이 걸치면 넓힘
  for (let grow = true; grow;) {
    grow = false;
    const g = S.cellRange;
    t.rows.forEach((row, ri) => row.cells.forEach((cell, ci) => {
      const r2 = ri + (cell.rowSpan ?? 1) - 1;
      const c2 = ci + (cell.span ?? 1) - 1;
      if (cell.hMerge || cell.vMerge) return;
      const inter = ri <= g.r2 && r2 >= g.r1 && ci <= g.c2 && c2 >= g.c1;
      if (inter && (ri < g.r1 || r2 > g.r2 || ci < g.c1 || c2 > g.c2)) { g.r1 = Math.min(g.r1, ri); g.r2 = Math.max(g.r2, r2); g.c1 = Math.min(g.c1, ci); g.c2 = Math.max(g.c2, c2); grow = true; }
    }));
  }
  emit('selection');
  renderOverlay();
}
function startCellDrag(e, t, r0, c0, editOnClick = true) {
  let moved = false;
  const at = { x: e.clientX, y: e.clientY };
  const onMove = (ev) => {
    const td = document.elementFromPoint(ev.clientX, ev.clientY)?.closest?.(`.ob[data-id="${t.id}"] [data-cell]`);
    if (!td) return;
    const [r, c] = td.dataset.cell.split(',').map(Number);
    if (!moved && r === r0 && c === c0) return;
    moved = true;
    window.getSelection()?.removeAllRanges();
    setCellRange(t, r0, c0, r, c);
  };
  const onUp = () => {
    removeEventListener('pointermove', onMove);
    removeEventListener('pointerup', onUp);
    if (!moved && editOnClick) { S.cellRange = null; S.sel = new Set([t.id]); startEdit(t, { ...at, cell: [r0, c0] }); }
  };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}
function tableOverlay(z) {
  const t = selObjects().length === 1 ? selObjects()[0] : S.editing ? objById(S.editing.id) : null;
  if (t?.type !== 'table' || t.rot) return;
  const g = S.cellRange?.id === t.id ? S.cellRange : null;
  if (g) {
    let y1 = t.y; for (let i = 0; i < g.r1; i++) y1 += t.rows[i].h;
    let y2 = y1; for (let i = g.r1; i <= g.r2; i++) y2 += t.rows[i].h;
    let x1 = t.x; for (let i = 0; i < g.c1; i++) x1 += t.cols[i];
    let x2 = x1; for (let i = g.c1; i <= g.c2; i++) x2 += t.cols[i];
    overlay.append(el('div', { class: 'cell-range', style: { left: `${x1 * z}px`, top: `${y1 * z}px`, width: `${(x2 - x1) * z}px`, height: `${(y2 - y1) * z}px` } }));
  }
  // 열 경계 · 행 경계 끌기 손잡이
  let x = t.x;
  t.cols.forEach((w, i) => {
    x += w;
    const hd = el('div', { class: 'col-grip', title: '열 너비 (끌기)', style: { left: `${x * z - 3}px`, top: `${t.y * z}px`, height: `${t.h * z}px` } });
    hd.addEventListener('pointerdown', (e) => gripDrag(e, t, 'col', i));
    overlay.append(hd);
  });
  let y = t.y;
  t.rows.forEach((r, i) => {
    y += r.h;
    const hd = el('div', { class: 'row-grip', title: '행 높이 (끌기)', style: { top: `${y * z - 3}px`, left: `${t.x * z}px`, width: `${t.w * z}px` } });
    hd.addEventListener('pointerdown', (e) => gripDrag(e, t, 'row', i));
    overlay.append(hd);
  });
}
function gripDrag(e, t, kind, i) {
  e.preventDefault();
  e.stopPropagation();
  if (S.editing) endEdit();
  const p0 = toSlide(e);
  const start = kind === 'col' ? t.cols[i] : t.rows[i].h;
  const next = kind === 'col' ? t.cols[i + 1] : null;
  S.history.record(S.pres);
  const onMove = (ev) => {
    const p = toSlide(ev);
    const d = kind === 'col' ? p.x - p0.x : p.y - p0.y;
    if (kind === 'col') {
      // 가운데 경계: 옆 열과 나눔 (전체 너비 유지), 마지막 경계: 표가 넓어짐 — PowerPoint 와 같음
      const v = Math.max(12, start + d);
      if (next != null && !ev.shiftKey) { const n2 = Math.max(12, next - (v - start)); t.cols[i] = start + next - n2; t.cols[i + 1] = n2; } else t.cols[i] = v;
      t.w = t.cols.reduce((a, b) => a + b, 0);
    } else {
      t.rows[i].h = Math.max(10, start + d);
      t.h = t.rows.reduce((a, b) => a + b.h, 0);
    }
    renderCanvas();
  };
  const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); S.dirty = true; emit('change', { scope: 'slide' }); };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}

// ───────────── 그림 자르기 (화면에서) ─────────────
/** 자르기 모드: 원래 그림 전체(흐리게) + 자르기 손잡이. 손잡이를 끌면 자르기, 그림을 끌면 안에서 옮기기 */
export function startCrop(o) {
  if (o?.type !== 'image' || !o.media) return false;
  if (S.editing) endEdit();
  S.cropping = o.id;
  S.sel = new Set([o.id]);
  emit('selection');
  renderOverlay();
  return true;
}
export function endCrop() { if (!S.cropping) return; S.cropping = null; renderOverlay(); emit('change', { scope: 'slide' }); }
function fullRect(o) {
  const c = o.crop ?? { l: 0, t: 0, r: 0, b: 0 };
  const fw = o.w / Math.max(0.01, 1 - c.l - c.r);
  const fh = o.h / Math.max(0.01, 1 - c.t - c.b);
  return { x: o.x - c.l * fw, y: o.y - c.t * fh, w: fw, h: fh };
}
function cropOverlay(z) {
  if (!S.cropping) return;
  const o = objById(S.cropping);
  if (!o || o.type !== 'image') { S.cropping = null; return; }
  const F = fullRect(o);
  const src = S.pres.media[o.media];
  const ghost = el('div', { class: 'crop-ghost', style: { left: `${F.x * z}px`, top: `${F.y * z}px`, width: `${F.w * z}px`, height: `${F.h * z}px`, backgroundImage: `url("${src}")` } });
  ghost.addEventListener('pointerdown', (e) => cropDrag(e, o, 'pan'));
  const win = el('div', { class: 'crop-win', style: { left: `${o.x * z}px`, top: `${o.y * z}px`, width: `${o.w * z}px`, height: `${o.h * z}px` } });
  win.addEventListener('pointerdown', (e) => cropDrag(e, o, 'pan'));
  for (const hd of ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w']) {
    const h = el('div', { class: `crop-hd ch-${hd}` });
    h.addEventListener('pointerdown', (e) => cropDrag(e, o, hd));
    win.append(h);
  }
  overlay.append(ghost, win);
}
function cropDrag(e, o, kind) {
  e.preventDefault();
  e.stopPropagation();
  const p0 = toSlide(e);
  const F0 = fullRect(o);
  const r0 = { x: o.x, y: o.y, w: o.w, h: o.h };
  S.history.record(S.pres);
  const MIN = 8;
  const onMove = (ev) => {
    const p = toSlide(ev);
    const dx = p.x - p0.x;
    const dy = p.y - p0.y;
    let F = { ...F0 };
    let r = { ...r0 };
    if (kind === 'pan') {
      F.x = Math.min(r.x, Math.max(r.x + r.w - F.w, F0.x + dx));
      F.y = Math.min(r.y, Math.max(r.y + r.h - F.h, F0.y + dy));
    } else {
      if (kind.includes('w')) { const nx = Math.max(F.x, Math.min(r0.x + r0.w - MIN, r0.x + dx)); r.w = r0.x + r0.w - nx; r.x = nx; }
      if (kind.includes('e')) r.w = Math.max(MIN, Math.min(F.x + F.w - r0.x, r0.w + dx));
      if (kind.includes('n')) { const ny = Math.max(F.y, Math.min(r0.y + r0.h - MIN, r0.y + dy)); r.h = r0.y + r0.h - ny; r.y = ny; }
      if (kind.includes('s')) r.h = Math.max(MIN, Math.min(F.y + F.h - r0.y, r0.h + dy));
    }
    Object.assign(o, r);
    o.crop = { l: (r.x - F.x) / F.w, t: (r.y - F.y) / F.h, r: (F.x + F.w - r.x - r.w) / F.w, b: (F.y + F.h - r.y - r.h) / F.h };
    for (const k of ['l', 't', 'r', 'b']) if (Math.abs(o.crop[k]) < 1e-4) o.crop[k] = 0;
    renderCanvas();
    renderOverlay();
  };
  const onUp = () => { removeEventListener('pointermove', onMove); removeEventListener('pointerup', onUp); S.dirty = true; emit('change', { scope: 'slide' }); };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}

function onDblClick(e) {
  const o = hitObject(e);
  if (!o) return;
  if (o.type === 'chart') { run('chartData'); return; }
  if (o.type === 'image') { if (o.media) startCrop(o); return; }
  if (o.type === 'table') {
    const td = e.target.closest('[data-cell]');
    if (td) { const [r, c] = td.dataset.cell.split(',').map(Number); startEdit(o, { x: e.clientX, y: e.clientY, cell: [r, c] }); }
    return;
  }
  if (o.text) startEdit(o, { x: e.clientX, y: e.clientY });
}

/** 끌어서 옮기기 (클릭만 하면 onClick) */
function startMove(e, p0, onClick = null) {
  const slide = curSlide();
  const moving = new Set();
  for (const o of selObjects()) for (const m of groupMembers(slide, o)) moving.add(m.id);
  const objs = slide.objects.filter((o) => moving.has(o.id));
  if (!objs.length) return;
  const start = objs.map((o) => ({ o, x: o.x, y: o.y }));
  const box0 = bbox(objs);
  const others = slide.objects.filter((o) => !moving.has(o.id));
  let moved = false;
  let recorded = false;
  let copyMode = false;
  drag = { kind: 'move' };
  const onMove = (ev) => {
    const p = toSlide(ev);
    let dx = p.x - p0.x;
    let dy = p.y - p0.y;
    if (!moved && Math.hypot(dx * S.zoom, dy * S.zoom) < 4) return;
    if (!moved) {
      moved = true;
      // Ctrl + 끌기 = 복사
      if (ev.ctrlKey) { copyMode = true; }
    }
    if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
    const guides = [];
    if (!ev.altKey) {
      const snapped = snapMove({ x: box0.x + dx, y: box0.y + dy, w: box0.w, h: box0.h }, others, guides);
      dx += snapped.dx; dy += snapped.dy;
    }
    if (!recorded) {
      S.history.record(S.pres);
      recorded = true;
      if (copyMode) {
        // 원래 자리에 복사본을 남기고 끌던 것을 옮김
        const copies = start.map(({ o }) => ({ ...JSON.parse(JSON.stringify(o)), id: uid() }));
        const gmap = new Map();
        for (const c of copies) { if (c.grp) { if (!gmap.has(c.grp)) gmap.set(c.grp, uid('g')); c.grp = gmap.get(c.grp); } delete c.ph; }
        const firstIdx = slide.objects.indexOf(start[0].o);
        slide.objects.splice(firstIdx, 0, ...copies);
      }
    }
    for (const s of start) { s.o.x = Math.round((s.x + dx) * 100) / 100; s.o.y = Math.round((s.y + dy) * 100) / 100; }
    moveDom(objs);
    renderOverlay();
    showGuides(guides);
  };
  const onUp = () => {
    removeEventListener('pointermove', onMove);
    removeEventListener('pointerup', onUp);
    drag = null;
    showGuides([]);
    if (moved) { S.dirty = true; emit('change', { scope: 'slide' }); } else onClick?.();
  };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}

/** 끄는 동안은 HTML 을 다시 만들지 않고 위치만 바꿈 */
function moveDom(objs) {
  for (const o of objs) {
    const d = layer.querySelector(`.ob[data-id="${CSS.escape(o.id)}"]`);
    if (!d) continue;
    d.style.left = `${o.x}px`;
    d.style.top = `${o.y}px`;
  }
}

/** 스마트 가이드: 슬라이드 가장자리 · 가운데, 다른 개체 가장자리 · 가운데에 붙임 */
function snapTargets(others) {
  const { w, h } = S.pres.size;
  const xs = [[0, 0, h], [w / 2, 0, h], [w, 0, h]];
  const ys = [[0, 0, w], [h / 2, 0, w], [h, 0, w]];
  for (const o of others) {
    if (o.decor) continue;
    const [x1, y1, x2, y2] = rotatedBox(o);
    for (const x of [x1, (x1 + x2) / 2, x2]) xs.push([x, y1, y2]);
    for (const y of [y1, (y1 + y2) / 2, y2]) ys.push([y, x1, x2]);
  }
  return { xs, ys };
}
function snapMove(b, others, guides) {
  const th = 6 / S.zoom;
  const { xs, ys } = snapTargets(others);
  const best = (edges, targets) => {
    let out = null;
    for (const e of edges) for (const t of targets) { const d = t[0] - e; if (Math.abs(d) <= th && (!out || Math.abs(d) < Math.abs(out.d))) out = { d, t }; }
    return out;
  };
  const bx = best([b.x, b.x + b.w / 2, b.x + b.w], xs);
  const by = best([b.y, b.y + b.h / 2, b.y + b.h], ys);
  const dx = bx?.d ?? 0;
  const dy = by?.d ?? 0;
  if (bx) guides.push({ v: true, at: bx.t[0], from: Math.min(bx.t[1], b.y + dy), to: Math.max(bx.t[2], b.y + dy + b.h) });
  if (by) guides.push({ v: false, at: by.t[0], from: Math.min(by.t[1], b.x + dx), to: Math.max(by.t[2], b.x + dx + b.w) });
  return { dx, dy };
}
function snapValue(v, targets, guides, vertical, span) {
  const th = 6 / S.zoom;
  let out = null;
  for (const t of targets) { const d = t[0] - v; if (Math.abs(d) <= th && (!out || Math.abs(d) < Math.abs(out.d))) out = { d, t }; }
  if (!out) return v;
  guides.push({ v: vertical, at: out.t[0], from: Math.min(out.t[1], span[0]), to: Math.max(out.t[2], span[1]) });
  return v + out.d;
}

/** 크기 조절 · 회전 손잡이 */
function startHandleDrag(e, p0, handle, frameIdx) {
  const frame = selectionFrames()[frameIdx];
  if (!frame) return;
  const slide = curSlide();
  const objs = frame.group ? frame.objs : frame.objs;
  const single = objs.length === 1 && !frame.group ? objs[0] : null;
  const orig = objs.map((o) => ({ o, x: o.x, y: o.y, w: o.w, h: o.h, rot: o.rot ?? 0, flipH: !!o.flipH, flipV: !!o.flipV, cols: o.cols ? [...o.cols] : null, rowsH: o.rows ? o.rows.map((r) => r.h) : null }));
  const others = slide.objects.filter((o) => !objs.includes(o));
  const { xs, ys } = snapTargets(others);
  let recorded = false;
  const rec = () => { if (!recorded) { S.history.record(S.pres); recorded = true; } };
  const onMove = (ev) => {
    const p = toSlide(ev);
    rec();
    const guides = [];
    if (handle === 'rot') {
      const o = single ?? objs[0];
      const cx = orig[0].x + orig[0].w / 2;
      const cy = orig[0].y + orig[0].h / 2;
      let ang = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
      if (ev.shiftKey) ang = Math.round(ang / 15) * 15;
      ang = ((ang % 360) + 360) % 360;
      if (Math.abs(ang) < 2 || Math.abs(ang - 360) < 2) ang = 0;
      o.rot = Math.round(ang * 10) / 10;
    } else if (single) {
      resizeOne(single, orig[0], handle, p.x - p0.x, p.y - p0.y, ev, guides, xs, ys);
    } else {
      // 여러 개 · 그룹: 둘러싼 상자 비율로
      const b0 = frame.box;
      const [hx, hy] = HPOS[handle];
      let x1 = b0.x; let y1 = b0.y; let x2 = b0.x + b0.w; let y2 = b0.y + b0.h;
      const dx = p.x - p0.x;
      const dy = p.y - p0.y;
      if (hx === 0) x1 += dx; if (hx === 1) x2 += dx;
      if (hy === 0) y1 += dy; if (hy === 1) y2 += dy;
      let sx = Math.max(0.02, (x2 - x1) / b0.w);
      let sy = Math.max(0.02, (y2 - y1) / b0.h);
      if (ev.shiftKey) { const k = hx === 0.5 ? sy : hy === 0.5 ? sx : Math.max(sx, sy); sx = k; sy = k; }
      const ox = hx === 0 ? b0.x + b0.w : hx === 1 ? b0.x : b0.x + b0.w / 2;
      const oy = hy === 0 ? b0.y + b0.h : hy === 1 ? b0.y : b0.y + b0.h / 2;
      for (const s of orig) {
        s.o.x = ox + (s.x - ox) * (hx === 0.5 ? 1 : sx);
        s.o.y = oy + (s.y - oy) * (hy === 0.5 ? 1 : sy);
        s.o.w = s.w * (hx === 0.5 ? 1 : sx);
        s.o.h = s.h * (hy === 0.5 ? 1 : sy);
        if (s.cols) s.o.cols = s.cols.map((c) => c * (hx === 0.5 ? 1 : sx));
        if (s.rowsH) s.o.rows.forEach((r, i) => { r.h = s.rowsH[i] * (hy === 0.5 ? 1 : sy); });
      }
    }
    renderCanvas();
    showGuides(guides);
  };
  const onUp = () => {
    removeEventListener('pointermove', onMove);
    removeEventListener('pointerup', onUp);
    showGuides([]);
    if (recorded) { S.dirty = true; emit('change', { scope: 'slide' }); }
  };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}

function resizeOne(o, s, handle, dx, dy, ev, guides, xs, ys) {
  const [hx, hy] = HPOS[handle].map((v) => v * 2 - 1); // -1, 0, 1
  const th = (s.rot * Math.PI) / 180;
  const cos = Math.cos(th);
  const sin = Math.sin(th);
  // 회전하지 않은 개체는 움직이는 가장자리를 가이드에 붙임
  if (!s.rot && !ev.altKey) {
    if (hx) { const edge = hx > 0 ? s.x + s.w + dx : s.x + dx; dx += snapValue(edge, xs, guides, true, [s.y, s.y + s.h]) - edge; }
    if (hy) { const edge = hy > 0 ? s.y + s.h + dy : s.y + dy; dy += snapValue(edge, ys, guides, false, [s.x, s.x + s.w]) - edge; }
  }
  const lx = dx * cos + dy * sin;
  const ly = -dx * sin + dy * cos;
  const line = o.type === 'shape' && isLineShape(o.shape);
  if (line && !s.rot) {
    // 선: 끝점을 옮김 (반대쪽을 지나가면 방향을 뒤집음)
    let x1 = s.x; let x2 = s.x + s.w; let y1 = s.y; let y2 = s.y + s.h;
    if (hx > 0) x2 += dx; if (hx < 0) x1 += dx;
    if (hy > 0) y2 += dy; if (hy < 0) y1 += dy;
    o.x = Math.min(x1, x2); o.y = Math.min(y1, y2); o.w = Math.abs(x2 - x1); o.h = Math.abs(y2 - y1);
    const fh = (x2 < x1) !== s.flipH;
    const fv = (y2 < y1) !== s.flipV;
    if (fh) o.flipH = true; else delete o.flipH;
    if (fv) o.flipV = true; else delete o.flipV;
    return;
  }
  let w = s.w + hx * lx;
  let h = s.h + hy * ly;
  const keep = ev.shiftKey !== (o.type === 'image' && hx !== 0 && hy !== 0);
  if (keep && hx && hy && s.w > 0 && s.h > 0) {
    const k = Math.max(w / s.w, h / s.h);
    w = s.w * k;
    h = s.h * k;
  }
  const nw = Math.max(1, w);
  const nh = Math.max(1, h);
  // 중심 이동 (로컬 좌표 → 전체)
  const cxl = hx * (nw - s.w) / 2;
  const cyl = hy * (nh - s.h) / 2;
  const cx = s.x + s.w / 2 + cxl * cos - cyl * sin;
  const cy = s.y + s.h / 2 + cxl * sin + cyl * cos;
  o.w = Math.round(nw * 100) / 100;
  o.h = Math.round(nh * 100) / 100;
  o.x = Math.round((cx - nw / 2) * 100) / 100;
  o.y = Math.round((cy - nh / 2) * 100) / 100;
  if (s.cols) { const k = nw / s.w; o.cols = s.cols.map((c) => c * k); }
  if (s.rowsH) { const k = nh / s.h; o.rows.forEach((r, i) => { r.h = s.rowsH[i] * k; }); }
}

/** 빈 곳에서 끌어 여러 개 고르기 */
function startBand(e, p0) {
  const band = el('div', { class: 'band' });
  overlay.append(band);
  const base = new Set(S.sel);
  const onMove = (ev) => {
    const p = toSlide(ev);
    const x1 = Math.min(p0.x, p.x); const y1 = Math.min(p0.y, p.y);
    const x2 = Math.max(p0.x, p.x); const y2 = Math.max(p0.y, p.y);
    Object.assign(band.style, { left: `${x1 * S.zoom}px`, top: `${y1 * S.zoom}px`, width: `${(x2 - x1) * S.zoom}px`, height: `${(y2 - y1) * S.zoom}px` });
    const hit = new Set(base);
    for (const o of curSlide().objects) {
      const [a, b, c, d] = rotatedBox(o);
      if (a >= x1 && b >= y1 && c <= x2 && d <= y2) for (const m of groupMembers(curSlide(), o)) hit.add(m.id);
    }
    S.sel = hit;
    renderOverlay();
    overlay.append(band);
  };
  const onUp = () => {
    removeEventListener('pointermove', onMove);
    removeEventListener('pointerup', onUp);
    band.remove();
    emit('selection');
  };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}

/** 삽입 › 도형 / 글 상자: 끌어서 그리기 (클릭만 하면 기본 크기) */
function startDraw(e, p0) {
  const kind = S.drawShape;
  const preview = el('div', { class: 'draw-preview' });
  overlay.append(preview);
  let p1 = p0;
  const onMove = (ev) => {
    p1 = toSlide(ev);
    let w = p1.x - p0.x;
    let h = p1.y - p0.y;
    if (ev.shiftKey) { const m = Math.max(Math.abs(w), Math.abs(h)); w = Math.sign(w || 1) * m; h = Math.sign(h || 1) * m; p1 = { x: p0.x + w, y: p0.y + h }; }
    Object.assign(preview.style, { left: `${Math.min(p0.x, p1.x) * S.zoom}px`, top: `${Math.min(p0.y, p1.y) * S.zoom}px`, width: `${Math.abs(w) * S.zoom}px`, height: `${Math.abs(h) * S.zoom}px` });
  };
  const onUp = () => {
    removeEventListener('pointermove', onMove);
    removeEventListener('pointerup', onUp);
    preview.remove();
    S.drawShape = null;
    emit('drawMode');
    const w = p1.x - p0.x;
    const h = p1.y - p0.y;
    const small = Math.abs(w) < 4 && Math.abs(h) < 4;
    const line = kind === 'line' || kind === 'straightConnector1';
    let rect = small ? { x: p0.x, y: p0.y, w: kind === 'textbox' ? 300 : 96, h: kind === 'textbox' ? 40 : line ? 0 : 96 } : { x: Math.min(p0.x, p1.x), y: Math.min(p0.y, p1.y), w: Math.abs(w), h: Math.abs(h) };
    if (kind === 'textbox' && !small) rect = { ...rect, h: Math.max(rect.h, 40) };
    const o = kind === 'textbox' ? newTextBox(rect) : newShape(kind, rect);
    if (line && !small) { if (w < 0) o.flipH = true; if (h < 0) o.flipV = true; if (o.flipH && o.flipV) { delete o.flipH; delete o.flipV; } }
    change(() => { curSlide().objects.push(o); });
    S.sel = new Set([o.id]);
    emit('selection');
    if (kind === 'textbox') startEdit(o, { end: true });
  };
  addEventListener('pointermove', onMove);
  addEventListener('pointerup', onUp);
}

// ───────────── 글 편집 ─────────────
function editTargetEl(o, cell) {
  const ob = layer.querySelector(`.ob[data-id="${CSS.escape(o.id)}"]`);
  if (!ob) return null;
  if (cell) return ob.querySelector(`.tcell[data-cell="${cell.join(',')}"] .txi`);
  return ob.querySelector(`.tx[data-tx="${CSS.escape(o.id)}"] .txi`);
}
export function editingBody() {
  if (!S.editing) return null;
  const o = objById(S.editing.id);
  if (!o) return null;
  return S.editing.cell ? o.rows[S.editing.cell[0]]?.cells[S.editing.cell[1]]?.text : o.text;
}
function setEditingBody(body) {
  const o = objById(S.editing.id);
  if (S.editing.cell) o.rows[S.editing.cell[0]].cells[S.editing.cell[1]].text = body; else o.text = body;
}

/** 글 편집 시작: at = {x, y} 클릭 위치 | {end} | {all} | {cell} */
export function startEdit(o, at = {}) {
  if (!o) return;
  if (o.type === 'table' && !at.cell) at = { ...at, cell: [0, 0] };
  if (!o.text && o.type === 'shape' && !isLineShape(o.shape)) {
    change(() => { o.text = { paras: [{ runs: [], align: 'ctr', lvl: 0 }], anchor: 'ctr', insets: [9.6, 4.8, 9.6, 4.8], wrap: true, autofit: 'none', defColor: '@lt1' }; });
  }
  if (!o.text && o.type !== 'table') return;
  S.sel = new Set([o.id]);
  S.editing = { id: o.id, ...(at.cell ? { cell: at.cell } : {}) };
  renderCanvas();
  if (!txi) return;
  if (at.x != null && placeCaretAtPoint(at.x, at.y) && txi.contains(window.getSelection().anchorNode)) { /* 클릭한 자리 */ } else if (at.all) selectAllIn(txi);
  else {
    const body = editingBody();
    const last = Math.max(0, body.paras.length - 1);
    setOffsets(txi, { a: { p: last, o: Infinity } });
  }
  emit('selection');
}

function attachEditor(offs) {
  const o = objById(S.editing.id);
  txi = o ? editTargetEl(o, S.editing.cell) : null;
  if (!txi) { S.editing = null; return; }
  editBase = JSON.parse(JSON.stringify(editingBody()));
  txi.contentEditable = 'true';
  txi.spellcheck = false;
  txi.classList.add('editing');
  txi.closest('.ob')?.classList.add('editing');
  txi.addEventListener('input', onInput);
  txi.addEventListener('keydown', onEditKey);
  txi.addEventListener('paste', onPaste);
  txi.focus({ preventScroll: true });
  if (offs) setOffsets(txi, offs);
}

function onInput() {
  const body = domToBody(txi, editBase);
  change(() => setEditingBody(body), { scope: 'text', key: `text:${S.editing.id}` });
  fitTextBox();
}

/** 글 상자(도형 크기를 텍스트에 맞춤): 높이를 글에 맞춤 */
function fitTextBox() {
  const o = objById(S.editing?.id);
  if (!o || S.editing.cell || o.text?.autofit !== 'resize' || o.rot) return;
  const ins = o.text.insets ?? [9.6, 4.8, 9.6, 4.8];
  const need = Math.ceil(txi.scrollHeight + ins[1] + ins[3]);
  if (Math.abs(need - o.h) < 1) return;
  o.h = need;
  const ob = txi.closest('.ob');
  ob.style.height = `${need}px`;
  const tx = txi.parentElement;
  tx.style.height = `${need - ins[1] - ins[3]}px`;
  renderOverlay();
}

function onPaste(e) {
  e.preventDefault();
  const text = e.clipboardData?.getData('text/plain') ?? '';
  if (!text) { emit('pasteInText', e); return; }
  const lines = text.replace(/\r/g, '').split('\n');
  lines.forEach((line, i) => {
    if (i) document.execCommand('insertParagraph');
    if (line) document.execCommand('insertText', false, line);
  });
}

function onEditKey(e) {
  if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); endEdit(); return; }
  if (e.key === 'Tab') {
    e.preventDefault();
    e.stopPropagation();
    if (S.editing.cell) { moveCell(e.shiftKey ? -1 : 1); return; }
    // 목록 수준 높이기 / 낮추기
    applyTextFormat(null, null, e.shiftKey ? -1 : 1);
    return;
  }
  // 그 밖의 단축키는 app.js 의 전역 처리로 (Ctrl+B 등) — 글자 입력은 막지 않음
  e.stopPropagation();
  if (((e.ctrlKey || e.metaKey) && !e.altKey) || /^F\d+$/.test(e.key) || (e.altKey && e.shiftKey && e.key.startsWith('Arrow'))) emit('editKey', e);
}

function moveCell(dir) {
  const o = objById(S.editing.id);
  let [r, c] = S.editing.cell;
  const cols = o.cols.length;
  c += dir;
  if (c >= cols) { c = 0; r++; }
  if (c < 0) { c = cols - 1; r--; }
  if (r < 0) r = 0;
  if (r >= o.rows.length) {
    // 마지막 칸에서 Tab → 새 행 (PowerPoint 와 같음)
    change(() => {
      const last = o.rows[o.rows.length - 1];
      o.rows.push({ h: last.h, cells: last.cells.map((cell) => ({ text: { ...JSON.parse(JSON.stringify(cell.text)), paras: [{ ...cell.text.paras[0], runs: [] }] } })) });
      o.h += last.h;
    });
  }
  S.editing = { id: o.id, cell: [r, c] };
  renderCanvas();
  if (txi) selectAllIn(txi);
}

/** 편집 끝: 줄어들 글꼴 비율 계산 (넘치면 텍스트 크기 줄이기) */
export function endEdit() {
  if (!S.editing) return;
  const o = objById(S.editing.id);
  if (txi?.isConnected) {
    const body = domToBody(txi, editBase);
    setEditingBody(body);
    if (o && !S.editing.cell && body.autofit === 'shrink') shrinkToFit(o, body);
  }
  S.editing = null;
  txi = null;
  editBase = null;
  lastOffs = null;
  window.getSelection()?.removeAllRanges();
  emit('change', { scope: 'slide' });
}

function shrinkToFit(o, body) {
  const tx = layer.querySelector(`.ob[data-id="${CSS.escape(o.id)}"] .tx`);
  if (!tx) return;
  const inner = tx.querySelector('.txi');
  const fits = () => inner.scrollHeight <= tx.clientHeight + 1;
  const theme = S.pres.theme;
  let scale = 1;
  const render = (s) => { inner.innerHTML = textHtml(theme, o, { ...body, fontScale: s, lnSpcReduction: s < 1 ? 0.1 : 0 }); };
  render(1);
  if (fits()) { delete body.fontScale; delete body.lnSpcReduction; return; }
  for (const s of [0.925, 0.85, 0.775, 0.7, 0.625, 0.55, 0.475, 0.4, 0.325, 0.25]) {
    scale = s;
    render(s);
    if (fits()) break;
  }
  body.fontScale = scale;
  body.lnSpcReduction = 0.1;
}

/** 서식 대상: 편집 중이면 선택한 글자, 아니면 선택한 개체 전체 글 */
export function textTargets() {
  if (S.editing) return { mode: 'edit' };
  const out = [];
  for (const o of selObjects()) {
    if (o.text) out.push(o.text);
    if (o.type === 'table') for (const r of o.rows) for (const c of r.cells) out.push(c.text);
  }
  return { mode: 'objects', bodies: out };
}

/**
 * 편집 중 글자 · 단락 서식
 * runProps: {b: true…} | null, paraProps: {align…} | null, lvlDelta: ±1
 */
export function applyTextFormat(runProps, paraProps, lvlDelta = 0, paraFn = null) {
  if (!S.editing || !txi) return false;
  const offs = offsetsNow() ?? { a: { p: 0, o: 0 }, b: { p: 0, o: 0 }, collapsed: true };
  const body = JSON.parse(JSON.stringify(domToBody(txi, editBase)));
  if (runProps) {
    let range = offs;
    if (offs.collapsed) range = wordAt(body, offs.a) ?? offs;
    formatRange(body, range.a, range.b, runProps);
  }
  const [p1, p2] = offs.a.p <= offs.b.p ? [offs.a.p, offs.b.p] : [offs.b.p, offs.a.p];
  for (let i = p1; i <= p2; i++) {
    const p = body.paras[i];
    if (!p) continue;
    if (paraProps) for (const [k, v] of Object.entries(paraProps)) { if (v == null) delete p[k]; else p[k] = v; }
    if (lvlDelta) p.lvl = Math.max(0, Math.min(8, (p.lvl ?? 0) + lvlDelta));
    paraFn?.(p);
  }
  change(() => setEditingBody(body), { scope: 'text' });
  editBase = JSON.parse(JSON.stringify(body));
  const o = objById(S.editing.id);
  const pseudo = S.editing.cell ? { ...o, ph: null, phKind: null } : o;
  txi.innerHTML = textHtml(S.pres.theme, pseudo, body, { editable: true });
  txi.focus({ preventScroll: true });
  setOffsets(txi, offs);
  fitTextBox();
  return true;
}

/** 편집 중 커서 위치 글자의 서식 값 */
export function caretRunProp(key, dflt) {
  if (!S.editing || !txi) return undefined;
  const offs = offsetsNow();
  const body = domToBody(txi, editBase);
  if (!offs) return undefined;
  return rangeRunProp(body, offs.a, offs.b, key, dflt);
}
export function caretPara() {
  if (!S.editing || !txi) return null;
  const offs = offsetsNow();
  const body = domToBody(txi, editBase);
  return body.paras[offs?.a.p ?? 0] ?? null;
}

/** Alt+Shift+↑/↓: 커서가 있는 단락을 위/아래로 옮김 */
export function moveParagraph(d) {
  if (!S.editing || !txi) return false;
  const offs = offsetsNow();
  const o = objById(S.editing.id);
  const cell = S.editing.cell;
  const body = domToBody(txi, editBase);
  const i = offs?.a.p ?? 0;
  const j = i + d;
  if (j < 0 || j >= body.paras.length) return true;
  [body.paras[i], body.paras[j]] = [body.paras[j], body.paras[i]];
  endEdit();
  change(() => { if (cell) o.rows[cell[0]].cells[cell[1]].text = body; else o.text = body; });
  startEdit(o, cell ? { cell } : { end: true });
  requestAnimationFrame(() => { if (txi) setOffsets(txi, { a: { p: j, o: offs?.a.o ?? 0 }, b: { p: j, o: offs?.a.o ?? 0 } }); });
  return true;
}

/** 편집 중 기호 · 글 넣기 */
export function insertTextAtCaret(t) {
  if (!S.editing || !txi) return false;
  txi.focus();
  document.execCommand('insertText', false, t);
  return true;
}

on('endEdit', () => endEdit());
export const focusEditing = () => { if (txi?.isConnected) txi.focus({ preventScroll: true }); };
export const isEditing = () => !!S.editing;
export const editorStage = () => stage;
export const editorLayer = () => layer;
