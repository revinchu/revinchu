// 명령 모음: 리본 · 메뉴 · 단축키가 run('이름') 으로 실행
import { S, curSlide, selObjects, selOne, objById, change, emit, register, run, goSlide, refresh } from './state.js';
import {
  newSlide, nextLayout, duplicateSlide, changeLayout, resetSlideLayout, cloneObjects, newTable, newChart, newImage, newTextBox, newShape,
  applyRunProps, applyParaProps, commonRunProp, alignObjects, reorder, groupMembers, uid, LAYOUTS,
  resizePresentation, addAnim, animSteps, ANIM_EFFECTS, ANIM_CLASS_LABEL, MOTION_PATHS, motionPoints, findAll, replaceAll, replaceInPara, setPlainText, plainText, slideTitle, defaultSize, isEmptyText, para, outline,
} from './model.js';
import { setOffsets } from './textedit.js';
import { THEMES, cloneTheme, themeByName, resolveColor } from './themes.js';
import { SHAPE_GALLERY, SHAPE_LABEL, objLabel } from './shapes.js';
import { CHART_KINDS, chartToTsv, tsvToChart } from './chart.js';
import { slideHtml, NUM_SCHEMES, BULLET_CHARS, DASH_LABEL } from './render.js';
import { SHAPE_STYLES, WORDART, PICTURE_STYLES, TABLE_STYLES, tableStyleProps, DESIGN_IDEAS } from './presets.js';
import { smartArt, SMART_KINDS } from './smartart.js';
import { backgroundMask, applyMask } from './bgremove.js';
import { parseLatex, toMathML } from './math.js';
import { el, openMenu, openDialog, formDialog, alertDialog, toast, closeMenus } from './ui.js';
import { colorMenu } from './colorpick.js';
import { startEyedrop } from './eyedrop.js';
import { moveParagraph, startCrop, endCrop, startEdit, endEdit, textTargets, applyTextFormat, caretRunProp, caretPara, insertTextAtCaret, fitZoom, setZoom, renderCanvas, editorLayer } from './editor.js';
import { startShow, playStepOn } from './show.js';
import { shapeIconSvg, styleChip, wordArtCss } from './galleries.js';
import { save, saveAs, openWithPicker, setDocument, printSlides, exportPng, exportVideo, shareLink, pickFile } from './fileio.js';
import { buildTemplate } from './templates.js';
import { renderPane } from './panels.js';
import { FONTS, SIZES, setRibbonTab } from './ribbon.js';

const lastColor = { font: '#C00000', hl: '#FFFF00', fill: '@accent1', line: '@tx1', cell: '@accent2:lm20:lo80' };
const slide = () => curSlide();
const need = (list, msg = '먼저 개체를 선택하세요') => { if (!list.length) { toast(msg); return false; } return true; };

// ───────────── 글자 · 단락 서식 ─────────────
function applyRun(props) {
  const t = textTargets();
  if (t.mode === 'edit') { applyTextFormat(props, null); return; }
  if (!need(t.bodies, '글자 서식을 바꿀 텍스트나 개체를 선택하세요')) return;
  change(() => { for (const b of t.bodies) applyRunProps(b, props); });
}
function applyPara(props, fn) {
  const t = textTargets();
  if (t.mode === 'edit') { applyTextFormat(null, props, 0, fn); return; }
  if (!need(t.bodies, '단락 서식을 바꿀 텍스트나 개체를 선택하세요')) return;
  change(() => { for (const b of t.bodies) { if (props) applyParaProps(b, props); if (fn) for (const p of b.paras) fn(p); } });
}
function firstObj() { return S.editing ? objById(S.editing.id) : selObjects().find((o) => o.text || o.type === 'table'); }
function currentRunProp(key) {
  if (S.editing) return caretRunProp(key, (p) => (key === 'size' ? defaultSize(firstObj() ?? {}, p.lvl ?? 0) : undefined));
  const o = firstObj();
  const body = o?.text ?? o?.rows?.[0]?.cells?.[0]?.text;
  if (!body) return undefined;
  return commonRunProp(body, key, (p) => (key === 'size' ? defaultSize(o, p.lvl ?? 0) * (body.fontScale ?? 1) : key === 'font' ? (o.ph === 'title' || o.ph === 'ctrTitle' ? '+mj' : '+mn') : undefined));
}
function currentPara() {
  if (S.editing) return caretPara();
  const o = firstObj();
  return o?.text?.paras?.[0] ?? o?.rows?.[0]?.cells?.[0]?.text?.paras?.[0] ?? null;
}
const toggleRun = (key) => applyRun({ [key]: currentRunProp(key) ? undefined : true });
function stepFont(dir) {
  const cur = Number(currentRunProp('size')) || 18;
  const i = SIZES.findIndex((s) => s >= cur);
  const nextSize = dir > 0 ? SIZES.find((s) => s > cur) ?? cur + 4 : [...SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  applyRun({ size: i < 0 && dir > 0 ? cur + 4 : nextSize });
}

// ───────────── 클립보드 ─────────────
function copyObjects(cut = false) {
  // 그룹 안에서 개체 하나만 골랐으면 그것만 (그룹 전체를 골랐으면 선택에 이미 모두 들어 있음)
  const objs = selObjects();
  if (!objs.length) return false;
  const media = {};
  for (const o of objs) if (o.media) media[o.media] = S.pres.media[o.media];
  S.clipboard = { kind: 'objects', objs: JSON.parse(JSON.stringify(objs)), media, at: Date.now(), from: slide().id, n: 0 };
  navigator.clipboard?.writeText(objs.map((o) => plainText(o.text)).filter(Boolean).join('\n') || ' ').catch(() => { /* 권한 없음 → 내부 클립보드만 */ });
  if (cut) change(() => { const ids = new Set(objs.map((o) => o.id)); slide().objects = slide().objects.filter((o) => !ids.has(o.id)); slide().anims = (slide().anims ?? []).filter((a) => !ids.has(a.obj)); S.sel.clear(); });
  return true;
}
function copySlides(cut = false) {
  const picked = S.pres.slides.filter((s) => S.slideSel.has(s.id));
  const list = picked.length ? picked : [slide()];
  const media = {};
  for (const s of list) for (const o of [...s.objects, ...(s.bgObjects ?? [])]) if (o.media) media[o.media] = S.pres.media[o.media];
  S.clipboard = { kind: 'slides', slides: JSON.parse(JSON.stringify(list)), media, at: Date.now() };
  if (cut) run('deleteSlide');
  toast(`슬라이드 ${list.length}장을 복사했습니다`);
}
function pasteInternal() {
  const c = S.clipboard;
  if (!c) return false;
  Object.assign(S.pres.media, c.media);
  if (c.kind === 'slides') {
    change(() => {
      const copies = c.slides.map((s) => duplicateSlide(s));
      S.pres.slides.splice(S.cur + 1, 0, ...copies);
    }, { scope: 'all' });
    goSlide(S.cur + 1);
    return true;
  }
  c.n = c.from === slide().id ? (c.n ?? 0) + 1 : 0;
  const off = c.from === slide().id ? 18 * c.n || 18 : 0;
  const copies = cloneObjects(c.objs, off, off);
  change(() => { slide().objects.push(...copies); });
  S.sel = new Set(copies.map((o) => o.id));
  emit('selection');
  return true;
}
const expandGroups = (objs) => { const s = slide(); const ids = new Set(); for (const o of objs) for (const m of groupMembers(s, o)) ids.add(m.id); return s.objects.filter((o) => ids.has(o.id)); };

async function readImageFile(file, maxSide = 2560) {
  const url = await new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(file); });
  return shrinkImage(url, maxSide, file.size);
}
/** 아주 큰 그림은 줄여서 넣음 (문서 크기 · 메모리) */
function shrinkImage(url, maxSide = 2560, bytes = 0) {
  return new Promise((resolve) => {
    if (url.startsWith('data:image/svg')) { const img = new Image(); img.onload = () => resolve({ url, w: img.naturalWidth || 300, h: img.naturalHeight || 300 }); img.onerror = () => resolve({ url, w: 300, h: 300 }); img.src = url; return; }
    const img = new Image();
    img.onload = () => {
      let w = img.naturalWidth;
      let h = img.naturalHeight;
      if (Math.max(w, h) <= maxSide && bytes < 4e6) { resolve({ url, w, h }); return; }
      const k = Math.min(1, maxSide / Math.max(w, h));
      w = Math.round(w * k); h = Math.round(h * k);
      const cv = document.createElement('canvas');
      cv.width = w; cv.height = h;
      cv.getContext('2d').drawImage(img, 0, 0, w, h);
      const png = url.startsWith('data:image/png') || url.startsWith('data:image/gif');
      resolve({ url: cv.toDataURL(png ? 'image/png' : 'image/jpeg', 0.9), w, h });
    };
    img.onerror = () => resolve(null);
    img.src = url;
  });
}
function addMedia(url) { const id = uid('m'); S.pres.media[id] = url; return id; }

/** 그림 넣기: into = 그림 개체 틀 id */
async function insertImages(files, into) {
  const list = [...files].filter((f) => f.type.startsWith('image/'));
  if (!list.length) return;
  const added = [];
  for (const f of list) {
    const r = await readImageFile(f);
    if (!r) { toast(`그림을 읽지 못했습니다: ${f.name}`); continue; }
    placeImage(r, into && !added.length ? into : null, added);
  }
  S.sel = new Set(added.map((o) => o.id));
  emit('selection');
}
function placeImage(r, into, added = []) {
  const s = slide();
  const ph = into ? s.objects.find((o) => o.id === into) : null;
  const id = addMedia(r.url);
  if (ph) {
    // 개체 틀을 채우도록 자름 (가로세로 비율 유지)
    const ar = r.w / r.h;
    const pr = ph.w / ph.h;
    const crop = ar > pr ? { l: (1 - pr / ar) / 2, r: (1 - pr / ar) / 2, t: 0, b: 0 } : { t: (1 - ar / pr) / 2, b: (1 - ar / pr) / 2, l: 0, r: 0 };
    change(() => { ph.media = id; ph.crop = crop; });
    added.push(ph);
    return ph;
  }
  const { w: W, h: H } = S.pres.size;
  const k = Math.min(1, (W * 0.8) / r.w, (H * 0.8) / r.h);
  const w = r.w * k;
  const h = r.h * k;
  const o = newImage(id, { x: (W - w) / 2 + added.length * 20, y: (H - h) / 2 + added.length * 20, w, h });
  change(() => { s.objects.push(o); });
  added.push(o);
  return o;
}

// ───────────── 비디오 · 오디오 ─────────────
const fileUrl = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(f); });
/** 비디오 첫 화면(1초 지점) → 표지 그림 · 크기 */
function videoInfo(url) {
  return new Promise((resolve) => {
    const v = document.createElement('video');
    v.muted = true;
    v.preload = 'auto';
    v.playsInline = true;
    let done = false;
    const finish = (poster) => { if (done) return; done = true; resolve({ w: v.videoWidth || 1280, h: v.videoHeight || 720, dur: v.duration || 0, poster }); };
    v.addEventListener('loadeddata', () => { v.currentTime = Math.min(1, (v.duration || 2) / 2); });
    v.addEventListener('seeked', () => {
      try {
        const c = document.createElement('canvas');
        c.width = v.videoWidth;
        c.height = v.videoHeight;
        c.getContext('2d').drawImage(v, 0, 0);
        finish(c.toDataURL('image/jpeg', 0.85));
      } catch { finish(null); }
    });
    v.addEventListener('error', () => finish(null));
    setTimeout(() => finish(null), 8000);
    v.src = url;
  });
}
const MEDIA_LIMIT = 80 * 1024 * 1024;
async function insertMedia(kind) {
  const f = await pickFile(kind === 'audio' ? 'audio/*,.mp3,.m4a,.wav,.aac,.ogg' : 'video/*,.mp4,.webm,.mov,.m4v');
  if (!f) return;
  if (f.size > MEDIA_LIMIT) { alertDialog(kind === 'audio' ? '오디오' : '비디오', `파일이 너무 큽니다 (${Math.round(f.size / 1048576)}MB). 80MB 이하로 줄여서 넣어 주세요.`); return; }
  toast('불러오는 중...');
  let url = await fileUrl(f);
  // 브라우저가 형식을 비워 두는 경우 (아이폰 .mov 등)
  if (/^data:;|^data:application\/octet-stream/.test(url)) {
    const ext = f.name.split('.').pop().toLowerCase();
    const mime = { mp4: 'video/mp4', m4v: 'video/x-m4v', mov: 'video/quicktime', webm: 'video/webm', mp3: 'audio/mpeg', m4a: 'audio/mp4', wav: 'audio/wav', aac: 'audio/aac', ogg: 'audio/ogg' }[ext] ?? (kind === 'audio' ? 'audio/mpeg' : 'video/mp4');
    url = url.replace(/^data:[^;,]*/, `data:${mime}`);
  }
  const s = slide();
  const { w: W, h: H } = S.pres.size;
  const id = addMedia(url);
  let o;
  if (kind === 'audio') {
    o = { id: uid(), type: 'media', kind: 'audio', media: id, x: (W - 64) / 2, y: (H - 64) / 2, w: 64, h: 64, rot: 0, name: f.name };
  } else {
    const info = await videoInfo(url);
    const k = Math.min(1, (W * 0.7) / info.w, (H * 0.7) / info.h);
    const w = info.w * k;
    const h = info.h * k;
    o = { id: uid(), type: 'media', kind: 'video', media: id, x: (W - w) / 2, y: (H - h) / 2, w, h, rot: 0, name: f.name };
    if (info.poster) o.poster = addMedia(info.poster);
  }
  change(() => { s.objects.push(o); });
  S.sel = new Set([o.id]);
  emit('selection');
  toast(kind === 'audio' ? '오디오를 넣었습니다. 슬라이드 쇼에서 아이콘을 누르면 재생됩니다' : '비디오를 넣었습니다. 슬라이드 쇼에서 재생됩니다');
}
function mediaOptions() {
  const o = selOne();
  if (o?.type !== 'media') { toast('비디오나 오디오를 선택하세요'); return; }
  formDialog(o.kind === 'audio' ? '오디오 재생' : '비디오 재생', [
    { name: 'autoplay', label: '자동 실행 (슬라이드가 나타날 때)', type: 'checkbox', value: !!o.autoplay },
    { name: 'loop', label: '반복 재생', type: 'checkbox', value: !!o.loop },
    { name: 'st', label: '시작 시간 (초)', type: 'number', value: String(o.trim?.st ?? 0) },
    { name: 'end', label: '끝 시간 (초, 0 = 끝까지)', type: 'number', value: String(o.trim?.end ?? 0) },
  ], (v) => change(() => {
    o.autoplay = v.autoplay || undefined;
    o.loop = v.loop || undefined;
    const st = Math.max(0, Number(v.st) || 0);
    const end = Math.max(0, Number(v.end) || 0);
    if (st || end) o.trim = { st, end }; else delete o.trim;
  }));
}

// ───────────── 표 ─────────────
function curTable() { const o = S.editing ? objById(S.editing.id) : selOne(); return o?.type === 'table' ? o : null; }
function curCell() { if (S.cellRange && S.cellRange.id === curTable()?.id) return [S.cellRange.r1, S.cellRange.c1]; return S.editing?.cell ?? [0, 0]; }
/** 표에서 고른 칸 범위 (칸 범위 → 편집 중인 칸 → 표 전체) */
function cellRange(t) {
  const g = S.cellRange;
  if (g && g.id === t.id) return g;
  if (S.editing?.id === t.id && S.editing.cell) { const [r, c] = S.editing.cell; const cell = t.rows[r].cells[c]; return { r1: r, c1: c, r2: r + (cell.rowSpan ?? 1) - 1, c2: c + (cell.span ?? 1) - 1 }; }
  return { r1: 0, c1: 0, r2: t.rows.length - 1, c2: t.cols.length - 1, all: true };
}
/** 범위 안의 칸들 [r, c, cell] */
function cellsIn(t, g) { const out = []; for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) out.push([r, c, t.rows[r].cells[c]]); return out; }
const PEN = { color: '@tx1', width: 1 };
const NO_LINE = { color: 'rgba(0,0,0,0)', width: 0 };
/** 표 테두리 (PowerPoint 테두리 메뉴): where = all | none | outside | inside | insideH | insideV | top | bottom | left | right | style */
function tableBorders(where) {
  tableOp((t) => {
    const g = cellRange(t);
    for (const [r, c, cell] of cellsIn(t, g)) {
      const b = { ...(cell.borders ?? {}) };
      const set = (side, v) => { b[side] = v; };
      const pen = { ...PEN, ...(S.pen ?? {}) };
      if (where === 'style') { delete cell.borders; continue; }
      if (where === 'all') { for (const k of ['l', 'r', 't', 'b']) set(k, pen); }
      if (where === 'none') { for (const k of ['l', 'r', 't', 'b']) set(k, NO_LINE); }
      if (where === 'outside') { if (r === g.r1) set('t', pen); if (r === g.r2) set('b', pen); if (c === g.c1) set('l', pen); if (c === g.c2) set('r', pen); }
      if (where === 'inside' || where === 'insideH') { if (r > g.r1) set('t', pen); if (r < g.r2) set('b', pen); }
      if (where === 'inside' || where === 'insideV') { if (c > g.c1) set('l', pen); if (c < g.c2) set('r', pen); }
      if (where === 'top' && r === g.r1) set('t', pen);
      if (where === 'bottom' && r === g.r2) set('b', pen);
      if (where === 'left' && c === g.c1) set('l', pen);
      if (where === 'right' && c === g.c2) set('r', pen);
      if (where === 'diagDown') b.dd = pen;
      if (where === 'diagUp') b.du = pen;
      cell.borders = b;
    }
  });
}
/** 칸 병합 (범위 전체, 가로 · 세로) */
function mergeRange(t, g) {
  const tl = t.rows[g.r1].cells[g.c1];
  for (const [r, c, cell] of cellsIn(t, g)) {
    if (r === g.r1 && c === g.c1) continue;
    for (const p of cell.text.paras) if (p.runs.some((x) => x.t)) tl.text.paras.push(p);
    cell.text.paras = [{ ...cell.text.paras[0], runs: [] }];
    delete cell.span; delete cell.rowSpan;
    cell.hMerge = c > g.c1 || undefined;
    cell.vMerge = r > g.r1 || undefined;
    if (!cell.hMerge) delete cell.hMerge;
    if (!cell.vMerge) delete cell.vMerge;
  }
  if (tl.text.paras.length > 1 && !tl.text.paras[0].runs.some((x) => x.t)) tl.text.paras.shift();
  if (g.c2 > g.c1) tl.span = g.c2 - g.c1 + 1; else delete tl.span;
  if (g.r2 > g.r1) tl.rowSpan = g.r2 - g.r1 + 1; else delete tl.rowSpan;
}
/** 병합 풀기 */
function unmergeAt(t, r, c) {
  const cell = t.rows[r].cells[c];
  const g = { r1: r, c1: c, r2: r + (cell.rowSpan ?? 1) - 1, c2: c + (cell.span ?? 1) - 1 };
  for (const [, , x] of cellsIn(t, g)) { delete x.hMerge; delete x.vMerge; }
  delete cell.span; delete cell.rowSpan;
}
/** 칸 나누기 (병합되지 않은 칸): 열/행을 끼워 넣고 다른 행·열은 병합으로 이어 줌 */
function splitCell(t, r, c, nCols, nRows) {
  const blank = (cell) => ({ text: { ...JSON.parse(JSON.stringify(cell.text)), paras: [{ ...cell.text.paras[0], runs: [] }] }, ...(cell.fill !== undefined ? { fill: cell.fill } : {}) });
  if (nCols > 1) {
    const w = t.cols[c] / nCols;
    t.cols.splice(c, 1, ...Array(nCols).fill(w));
    t.rows.forEach((row, ri) => {
      const cell = row.cells[c];
      const extra = Array.from({ length: nCols - 1 }, () => blank(cell));
      if (ri === r) { row.cells.splice(c + 1, 0, ...extra); return; }
      // 다른 행: 원래 칸이 새 열들까지 덮음
      for (const x of extra) { x.hMerge = true; if (cell.vMerge) x.vMerge = true; }
      row.cells.splice(c + 1, 0, ...extra);
      if (!cell.hMerge) cell.span = (cell.span ?? 1) + nCols - 1;
      else { for (let k = c; k >= 0; k--) { const head = row.cells[k]; if (!head.hMerge) { head.span = (head.span ?? 1) + nCols - 1; break; } } }
    });
  }
  if (nRows > 1) {
    const row = t.rows[r];
    const h = row.h / nRows;
    row.h = h;
    const added = Array.from({ length: nRows - 1 }, () => ({ h, cells: row.cells.map((cell, ci) => {
      const x = blank(cell);
      if (ci >= c && ci < c + (nCols > 1 ? nCols : 1)) return x;
      x.vMerge = true;
      if (cell.hMerge) x.hMerge = true;
      return x;
    }) }));
    t.rows.splice(r + 1, 0, ...added);
    row.cells.forEach((cell, ci) => {
      if (ci >= c && ci < c + (nCols > 1 ? nCols : 1)) return;
      if (!cell.vMerge && !cell.hMerge) cell.rowSpan = (cell.rowSpan ?? 1) + nRows - 1;
      else if (cell.vMerge && !cell.hMerge) { for (let k = r; k >= 0; k--) { const head = t.rows[k].cells[ci]; if (!head.vMerge) { head.rowSpan = (head.rowSpan ?? 1) + nRows - 1; break; } } }
    });
  }
}
function tableOp(fn) {
  const t = curTable();
  if (!t) { toast('표를 선택하세요'); return; }
  const wasEditing = !!S.editing;
  if (wasEditing) endEdit();
  change(() => fn(t));
  t.h = t.rows.reduce((s, r) => s + r.h, 0);
  t.w = t.cols.reduce((s, c) => s + c, 0);
  refresh();
}
const blankCell = (like) => ({ text: { ...JSON.parse(JSON.stringify(like?.text ?? { paras: [], anchor: 'ctr', insets: [7, 3.5, 7, 3.5] })), paras: [{ ...(like?.text?.paras?.[0] ?? { align: 'l', lvl: 0 }), runs: [] }] } });

// ───────────── 슬라이드 ─────────────
function addSlide(layout) {
  const s = slide();
  const ns = newSlide(S.pres, layout ?? nextLayout(s?.layout));
  if (s?.hideDecor) ns.hideDecor = true;
  if (s?.bgObjects && !layout) { ns.bgObjects = JSON.parse(JSON.stringify(s.bgObjects)); ns.bg = s.bg ? JSON.parse(JSON.stringify(s.bg)) : null; }
  change(() => { S.pres.slides.splice(S.cur + 1, 0, ns); }, { scope: 'all' });
  goSlide(S.cur + 1);
}
function pickedSlides() { const l = S.pres.slides.filter((s) => S.slideSel.has(s.id)); return l.length ? l : [slide()]; }

function layoutPicker(anchor, onPick) {
  const grid = el('div', { class: 'layout-grid' }, LAYOUTS.map(([k, l]) => {
    const p = { ...S.pres, slides: [] };
    const s = newSlide(S.pres, k);
    for (const o of s.objects) if (o.text) o.text.paras[0].runs = [];
    const sc = 120 / S.pres.size.w;
    const b = el('button', { class: 'layout-item', title: l, onclick: () => { closeMenus(); onPick(k); } },
      el('div', { class: 'tw', style: { width: '120px', height: `${S.pres.size.h * sc}px` }, html: `<div class="tw-in" style="transform:scale(${sc})">${slideHtml(p, s, { prompt: true })}</div>` }),
      el('span', {}, l));
    return b;
  }));
  return openMenu(anchor, [{ title: 'Office 테마' }, { node: grid }], { minWidth: 420 });
}

// ───────────── 메뉴 도우미 ─────────────
const menuAt = (anchor, items, o) => openMenu(anchor, items, o);

function shapesGrid(anchor, onPick) {
  const node = el('div', { class: 'shape-menu' }, SHAPE_GALLERY.map(([cat, list]) => el('div', {},
    el('div', { class: 'menu-title' }, cat),
    el('div', { class: 'shape-grid' }, list.map(([k, l]) => el('button', { class: 'shape-btn', title: l, html: shapeIconSvg(k, 20), onclick: () => { closeMenus(); onPick(k); } }))))));
  return openMenu(anchor, [{ node }], { minWidth: 300, scroll: true });
}

function tableGrid(anchor) {
  const label = el('div', { class: 'menu-title' }, '표 삽입');
  const cells = [];
  const grid = el('div', { class: 'tbl-grid' });
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const b = el('div', { class: 'tg', dataset: { r: String(r), c: String(c) } });
    cells.push(b);
    grid.append(b);
  }
  grid.addEventListener('mousemove', (e) => {
    const t = e.target.closest('.tg');
    if (!t) return;
    const R = Number(t.dataset.r); const C = Number(t.dataset.c);
    for (const x of cells) x.classList.toggle('on', Number(x.dataset.r) <= R && Number(x.dataset.c) <= C);
    label.textContent = `${C + 1}x${R + 1} 표`;
  });
  grid.addEventListener('click', (e) => {
    const t = e.target.closest('.tg');
    if (!t) return;
    closeMenus();
    insertTable(Number(t.dataset.r) + 1, Number(t.dataset.c) + 1);
  });
  return openMenu(anchor, [{ node: label }, { node: grid }, { sep: true }, { label: '표 삽입...', icon: 'table', action: () => formDialog('표 삽입', [{ name: 'c', label: '열 개수', type: 'number', value: '5' }, { name: 'r', label: '행 개수', type: 'number', value: '2' }], (v) => insertTable(Math.max(1, Math.min(75, Number(v.r))), Math.max(1, Math.min(75, Number(v.c))))) }]);
}
function insertTable(rows, cols) {
  const { w: W, h: H } = S.pres.size;
  const w = Math.min(W * 0.8, cols * 160);
  const t = newTable(rows, cols, { x: (W - w) / 2, y: H * 0.28, w, h: rows * 40 });
  const ph = slide().objects.find((o) => o.ph === 'body' && isEmptyText(o.text));
  if (ph) { t.x = ph.x; t.y = ph.y; t.cols = t.cols.map(() => ph.w / cols); t.w = ph.w; }
  change(() => { if (ph) slide().objects = slide().objects.filter((o) => o !== ph); slide().objects.push(t); });
  S.sel = new Set([t.id]);
  emit('selection');
}

function insertChart(kind) {
  const { w: W, h: H } = S.pres.size;
  const ph = slide().objects.find((o) => o.ph === 'body' && isEmptyText(o.text));
  const rect = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: W * 0.15, y: H * 0.2, w: W * 0.7, h: H * 0.65 };
  const c = newChart(rect);
  c.chart.kind = kind;
  if (kind === 'pie' || kind === 'doughnut') { c.chart.series = [c.chart.series[0]]; c.chart.labels = true; c.chart.legend = 'r'; }
  change(() => { if (ph) slide().objects = slide().objects.filter((o) => o !== ph); slide().objects.push(c); });
  S.sel = new Set([c.id]);
  emit('selection');
  chartDataDialog(c);
}

function chartDataDialog(o) {
  const c = o ?? selOne();
  if (c?.type !== 'chart') { toast('차트를 선택하세요'); return; }
  const ch = c.chart;
  const kind = el('select', {}, CHART_KINDS.map(([k, l]) => el('option', { value: k, selected: k === ch.kind }, l)));
  const title = el('input', { type: 'text', value: ch.title ?? '' });
  const legend = el('select', {}, [['b', '아래쪽'], ['r', '오른쪽'], ['', '없음']].map(([k, l]) => el('option', { value: k, selected: (ch.legend ?? '') === k }, l)));
  const labels = el('input', { type: 'checkbox', checked: !!ch.labels });
  const stacked = el('input', { type: 'checkbox', checked: !!ch.stacked });
  const data = el('textarea', { class: 'chart-data', spellcheck: 'false' }, chartToTsv(ch));
  data.addEventListener('keydown', (e) => { if (e.key === 'Tab') { e.preventDefault(); const s = data.selectionStart; data.setRangeText('\t', s, data.selectionEnd, 'end'); } });
  openDialog({
    title: '차트 데이터 편집', width: 560,
    body: el('div', { class: 'form-grid' },
      el('label', {}, el('span', {}, '차트 종류'), kind), el('label', {}, el('span', {}, '차트 제목'), title),
      el('label', {}, el('span', {}, '범례'), legend),
      el('div', { class: 'row-checks' }, el('label', {}, labels, el('span', {}, '데이터 레이블')), el('label', {}, stacked, el('span', {}, '누적'))),
      el('div', { class: 'muted' }, '첫 행은 계열 이름, 첫 열은 항목입니다. 엑셀(WIXEL)에서 복사해 붙여 넣을 수 있습니다 (탭으로 구분).'),
      data),
    buttons: [{ label: '확인', primary: true, action: () => {
      try {
        const next = tsvToChart({ ...ch, kind: kind.value, title: title.value, legend: legend.value || null, labels: labels.checked, stacked: stacked.checked }, data.value);
        change(() => { c.chart = next; });
      } catch (e) { alertDialog('차트', e.message); return false; }
      return true;
    } }, { label: '취소' }],
  });
}

// ───────────── 테마 ─────────────
function applyTheme(name, { colorsOnly = false, fontsFrom = null } = {}) {
  const t = themeByName(name);
  if (!t) return;
  change(() => {
    const cur = S.pres.theme;
    if (colorsOnly) S.pres.theme = { ...cur, colors: { ...t.colors } };
    else S.pres.theme = cloneTheme(t);
    if (fontsFrom) S.pres.theme.fonts = { ...fontsFrom };
  }, { scope: 'all' });
}
const FONT_PAIRS = [['맑은 고딕', '맑은 고딕'], ['나눔고딕', '나눔고딕'], ['나눔스퀘어', '나눔고딕'], ['Noto Sans KR', 'Noto Sans KR'], ['나눔명조', '나눔고딕'], ['Pretendard', 'Pretendard'], ['본고딕', '본고딕'], ['Arial Black', '맑은 고딕'], ['Georgia', '바탕'], ['궁서', '바탕']];
const BG_STYLES = [
  { type: 'solid', color: '@bg1' }, { type: 'solid', color: '@bg2' }, { type: 'solid', color: '@tx2' }, { type: 'solid', color: '@tx1' },
  { type: 'gradient', angle: 90, stops: [[0, '@bg1'], [1, '@bg2']] }, { type: 'gradient', angle: 90, stops: [[0, '@bg2'], [1, '@accent1:lm20:lo80']] },
  { type: 'gradient', angle: 90, stops: [[0, '@tx2:lm75'], [1, '@tx2']] }, { type: 'gradient', path: true, stops: [[0, '@accent1:lm60:lo40'], [1, '@accent1:lm75']] },
  { type: 'solid', color: '@accent1:lm20:lo80' }, { type: 'solid', color: '@accent2:lm20:lo80' }, { type: 'gradient', angle: 45, stops: [[0, '@accent1'], [1, '@accent2']] }, { type: 'solid', color: '#000000' },
];

// ───────────── 정렬 · 그룹 ─────────────
function doAlign(how) {
  const objs = expandGroups(selObjects());
  if (!need(objs)) return;
  change(() => alignObjects(objs, how, S.pres.size, S.alignToSlide || objs.length === 1));
}
function doGroup() {
  const objs = selObjects();
  if (objs.length < 2) { toast('그룹으로 묶을 개체를 두 개 이상 선택하세요'); return; }
  const g = uid('g');
  change(() => {
    // 그룹 구성원은 순서상 붙어 있어야 함 (pptx 그룹) → 가장 위 구성원 자리로 모음
    const s = slide();
    const ids = new Set(objs.map((o) => o.id));
    const members = s.objects.filter((o) => ids.has(o.id));
    for (const o of members) o.grp = g;
    const lastIdx = Math.max(...members.map((o) => s.objects.indexOf(o)));
    const rest = s.objects.filter((o) => !ids.has(o.id));
    const before = s.objects.slice(0, lastIdx + 1).filter((o) => !ids.has(o.id)).length;
    rest.splice(before, 0, ...members);
    s.objects = rest;
  });
}
function doUngroup() {
  const objs = selObjects().filter((o) => o.grp);
  if (!objs.length) { toast('그룹을 선택하세요'); return; }
  change(() => { for (const o of expandGroups(objs)) delete o.grp; });
}

// ───────────── 애니메이션 ─────────────
/** 이동 경로 반대로 (끝점이 원점이 되도록) */
function reversePath(path) {
  const pts = motionPoints(path, 8);
  const [ex, ey] = pts[pts.length - 1];
  return `M 0 0 ${pts.slice(0, -1).reverse().map(([x, y]) => `L ${Math.round((x - ex) * 10000) / 10000} ${Math.round((y - ey) * 10000) / 10000}`).join(' ')}`;
}
function setAnim(cls, effect) {
  const objs = expandGroups(selObjects());
  if (!need(objs, '애니메이션을 적용할 개체를 선택하세요')) return;
  const s = slide();
  change(() => {
    for (const o of objs) {
      const ex = s.anims.find((a) => a.obj === o.id);
      if (ex) { ex.cls = cls; ex.effect = effect; if (effect === 'appear' || effect === 'disappear') ex.dur = 0; else if (!ex.dur) ex.dur = cls === 'path' ? 2 : 0.5; if (cls === 'path') ex.motion = MOTION_PATHS[effect]; else delete ex.motion; } else addAnim(s, o.id, cls, effect, objs.indexOf(o) > 0 ? { start: 'with' } : {});
    }
  });
  previewAnims(objs.map((o) => o.id));
}
function previewAnims(onlyIds = null) {
  const s = slide();
  let steps = animSteps(s);
  if (onlyIds) { const set = new Set(onlyIds); steps = steps.map((st) => st.filter((a) => set.has(a.obj))).filter((st) => st.length).map((st, i) => st.map((a, k) => (i === 0 && k === 0 ? { ...a, start: 'click' } : a))); }
  if (!steps.length) return;
  const layer = editorLayer();
  const first = new Map();
  for (const st of steps) for (const a of st) if (!first.has(a.obj)) first.set(a.obj, a);
  for (const [obj, a] of first) if (a.cls === 'entr') { const e = layer.querySelector(`.ob[data-id="${CSS.escape(obj)}"]`); if (e) e.style.visibility = 'hidden'; }
  (async () => {
    for (const st of steps) await playStepOn(layer, st, s);
    setTimeout(() => renderCanvas(), 400);
  })();
}

function previewTransition() {
  const s = slide();
  const tr = s.transition;
  if (!tr || tr.type === 'none') { toast('전환 효과가 없습니다'); return; }
  const layer = editorLayer();
  const prev = S.pres.slides[S.cur - 1];
  const host = layer.parentElement;
  const old = el('div', { class: 'tr-prev', html: prev ? slideHtml(S.pres, prev, { index: S.cur - 1 }) : `<div class="sl" style="width:${S.pres.size.w}px;height:${S.pres.size.h}px;background:#000"></div>` });
  old.style.cssText = `position:absolute;left:0;top:0;transform:${layer.style.transform};transform-origin:0 0;z-index:0`;
  host.insertBefore(old, layer);
  layer.style.zIndex = '1';
  layer.style.position = 'absolute';
  const dur = (tr.dur ?? 1) * 1000;
  const { w, h } = S.pres.size;
  const off = { b: [0, h], t: [0, -h], l: [-w, 0], r: [w, 0] }[tr.dir ?? 'b'] ?? [0, h];
  const base = layer.style.transform;
  const T = (x, y) => `translate(${x * S.zoom}px, ${y * S.zoom}px) ${base}`;
  const opt = { duration: dur, easing: 'ease-in-out' };
  const map = {
    fade: () => layer.animate([{ opacity: 0 }, { opacity: 1 }], opt),
    dissolve: () => layer.animate([{ opacity: 0, filter: 'blur(4px)' }, { opacity: 1, filter: 'blur(0)' }], opt),
    push: () => { old.animate([{ transform: base }, { transform: T(-off[0], -off[1]) }], opt); return layer.animate([{ transform: T(off[0], off[1]) }, { transform: base }], opt); },
    cover: () => layer.animate([{ transform: T(off[0], off[1]) }, { transform: base }], opt),
    uncover: () => { old.style.zIndex = '2'; return old.animate([{ transform: base }, { transform: T(-off[0], -off[1]) }], { ...opt, fill: 'forwards' }); },
    wipe: () => layer.animate([{ clipPath: { b: 'inset(100% 0 0 0)', t: 'inset(0 0 100% 0)', l: 'inset(0 100% 0 0)', r: 'inset(0 0 0 100%)' }[tr.dir ?? 'r'] ?? 'inset(0 0 0 100%)' }, { clipPath: 'inset(0 0 0 0)' }], opt),
    split: () => layer.animate([{ clipPath: 'inset(0 50% 0 50%)' }, { clipPath: 'inset(0 0% 0 0%)' }], opt),
    zoom: () => layer.animate([{ opacity: 0, transform: `${base} scale(.3)` }, { opacity: 1, transform: base }], opt),
    circle: () => layer.animate([{ clipPath: 'circle(0% at 50% 50%)' }, { clipPath: 'circle(75% at 50% 50%)' }], opt),
    flip: () => layer.animate([{ opacity: 0, transform: `${base} perspective(2000px) rotateY(-90deg)` }, { opacity: 1, transform: base }], opt),
    cut: () => layer.animate([{ opacity: 0 }, { opacity: 0, offset: 0.99 }, { opacity: 1 }], { duration: 120 }),
  };
  const a = (map[tr.type] ?? map.fade)();
  a.finished.catch(() => {}).then(() => { old.remove(); layer.style.zIndex = ''; layer.style.position = ''; });
}

// ───────────── 대화상자 ─────────────
function findDialog(replace = false) {
  const q = el('input', { type: 'text', value: S.lastFind ?? '' });
  const r = el('input', { type: 'text' });
  const mc = el('input', { type: 'checkbox' });
  const ww = el('input', { type: 'checkbox' });
  const status = el('div', { class: 'muted find-status' });
  let hits = [];
  let k = -1;
  const opts = () => ({ matchCase: mc.checked, wholeWord: ww.checked });
  const findNext = () => {
    S.lastFind = q.value;
    hits = findAll(S.pres, q.value, opts());
    if (!hits.length) { status.textContent = '찾는 항목이 없습니다.'; return; }
    k = (k + 1) % hits.length;
    const h = hits[k];
    status.textContent = `${hits.length}개 중 ${k + 1}번째 · 슬라이드 ${h.slide + 1}${h.notes ? ' (메모)' : ''}`;
    if (S.cur !== h.slide) goSlide(h.slide);
    if (h.notes) return;
    const o = slide().objects.find((x) => x.id === h.obj);
    if (o) {
      startEdit(o, h.cell ? { cell: h.cell } : {});
      requestAnimationFrame(() => { const txi = document.querySelector('.txi.editing'); if (txi) setOffsets(txi, { a: { p: h.para, o: h.start }, b: { p: h.para, o: h.start + h.len } }); });
    }
  };
  const doReplaceAll = () => {
    if (S.editing) endEdit();
    let n = 0;
    change(() => { n = replaceAll(S.pres, q.value, r.value, opts()); }, { scope: 'all' });
    status.textContent = `${n}개를 바꿨습니다.`;
  };
  const doReplace = () => {
    const h = hits[k];
    if (!h) { findNext(); return; }
    if (S.editing) endEdit();
    change(() => {
      const s = S.pres.slides[h.slide];
      if (h.notes) { s.notes = s.notes.slice(0, h.start) + r.value + s.notes.slice(h.start + h.len); return; }
      const o = s.objects.find((x) => x.id === h.obj);
      const body = h.cell ? o.rows[h.cell[0]].cells[h.cell[1]].text : o.text;
      replaceInPara(body.paras[h.para], h.start, h.len, r.value);
    });
    k--;
    setTimeout(findNext, 30);
  };
  q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); findNext(); } });
  openDialog({
    title: replace ? '바꾸기' : '찾기', modeless: true, width: 420,
    body: el('div', { class: 'form-grid' },
      el('label', {}, el('span', {}, '찾을 내용'), q),
      replace ? el('label', {}, el('span', {}, '바꿀 내용'), r) : null,
      el('div', { class: 'row-checks' }, el('label', {}, mc, el('span', {}, '대/소문자 구분')), el('label', {}, ww, el('span', {}, '전체 단어만'))),
      status),
    buttons: [{ label: '다음 찾기', primary: true, action: () => { findNext(); return false; } }, ...(replace ? [{ label: '바꾸기', action: () => { doReplace(); return false; } }, { label: '모두 바꾸기', action: () => { doReplaceAll(); return false; } }] : [{ label: '바꾸기...', action: () => { findDialog(true); } }]), { label: '닫기' }],
  });
}

function fontDialog() {
  const fam = currentRunProp('font') ?? '+mn';
  formDialog('글꼴', [
    { name: 'font', label: '한글/영문 글꼴', type: 'combo', value: fam === '+mj' ? `${S.pres.theme.fonts.major} (제목)` : fam === '+mn' ? `${S.pres.theme.fonts.minor} (본문)` : fam, options: FONTS.map((f) => ({ value: f === '+mj' ? `${S.pres.theme.fonts.major} (제목)` : f === '+mn' ? `${S.pres.theme.fonts.minor} (본문)` : f })) },
    { name: 'size', label: '크기 (pt)', type: 'number', value: String(currentRunProp('size') ?? 18) },
    { name: 'b', label: '굵게', type: 'checkbox', value: !!currentRunProp('b') },
    { name: 'i', label: '기울임꼴', type: 'checkbox', value: !!currentRunProp('i') },
    { name: 'u', label: '밑줄', type: 'checkbox', value: !!currentRunProp('u') },
    { name: 's', label: '취소선', type: 'checkbox', value: !!currentRunProp('s') },
    { name: 'base', label: '첨자', type: 'select', value: String(currentRunProp('base') ?? 0), options: [{ value: '0', label: '없음' }, { value: '1', label: '위 첨자' }, { value: '-1', label: '아래 첨자' }] },
    { name: 'cap', label: '대문자', type: 'select', value: currentRunProp('cap') ?? '', options: [{ value: '', label: '없음' }, { value: 'small', label: '작은 대문자' }, { value: 'all', label: '모두 대문자' }] },
    { name: 'spc', label: '문자 간격 (pt)', type: 'number', value: String(currentRunProp('spc') ?? 0) },
  ], (v) => {
    let font = v.font.trim();
    if (font.endsWith('(제목)')) font = '+mj'; else if (font.endsWith('(본문)')) font = '+mn';
    applyRun({ font: font || undefined, size: Number(v.size) || undefined, b: v.b || undefined, i: v.i || undefined, u: v.u || undefined, s: v.s || undefined, base: Number(v.base) || undefined, cap: v.cap || undefined, spc: Number(v.spc) || undefined });
  });
}
function paragraphDialog() {
  const p = currentPara() ?? {};
  const cm = (px) => String(Math.round((px ?? 0) * 2.54 / 96 * 100) / 100);
  formDialog('단락', [
    { name: 'align', label: '맞춤', type: 'select', value: p.align ?? 'l', options: [{ value: 'l', label: '왼쪽' }, { value: 'ctr', label: '가운데' }, { value: 'r', label: '오른쪽' }, { value: 'just', label: '양쪽' }, { value: 'dist', label: '균등 분할' }] },
    { name: 'marL', label: '텍스트 앞 (cm)', type: 'number', value: cm(p.marL) },
    { name: 'indent', label: '첫 줄 (cm, 내어쓰기는 −)', type: 'number', value: cm(p.indent) },
    { name: 'spcBef', label: '단락 앞 (pt)', type: 'number', value: String(p.spcBef ?? 0) },
    { name: 'spcAft', label: '단락 뒤 (pt)', type: 'number', value: String(p.spcAft ?? 0) },
    { name: 'ls', label: '줄 간격 (배수)', type: 'number', value: String(p.lineSpacing ?? 1) },
  ], (v) => {
    const px = (c) => (Number(c) * 96) / 2.54;
    applyPara({ align: v.align, marL: v.marL === '' ? undefined : px(v.marL), indent: v.indent === '' ? undefined : px(v.indent), spcBef: Number(v.spcBef), spcAft: Number(v.spcAft), lineSpacing: Number(v.ls) || 1, lineSpacingPt: undefined });
  });
}

function headerFooterDialog() {
  const f = S.pres.footer ?? {};
  formDialog('머리글/바닥글', [
    { name: 'date', label: '날짜 및 시간', type: 'checkbox', value: !!f.date },
    { name: 'dateText', label: '고정 날짜 (비우면 오늘)', type: 'text', value: f.dateText ?? '' },
    { name: 'slideNum', label: '슬라이드 번호', type: 'checkbox', value: !!f.slideNum },
    { name: 'on', label: '바닥글', type: 'checkbox', value: !!f.text },
    { name: 'text', label: '바닥글 내용', type: 'text', value: f.text ?? '' },
    { name: 'hideOnTitle', label: '제목 슬라이드에는 표시 안 함', type: 'checkbox', value: f.hideOnTitle !== false },
  ], (v) => change(() => { S.pres.footer = { date: v.date, dateText: v.dateText || undefined, slideNum: v.slideNum, text: v.on ? v.text : '', hideOnTitle: v.hideOnTitle }; }, { scope: 'all' }), { okLabel: '모두 적용' });
}

function hyperlinkDialog() {
  const o = S.editing ? null : selOne();
  if (!S.editing && !o) { toast('링크를 걸 텍스트나 개체를 선택하세요'); return; }
  const cur = S.editing ? caretRunProp('link') : o?.link;
  const slides = S.pres.slides.map((s, i) => ({ value: `#slide${i + 1}`, label: `${i + 1}. ${slideTitle(s) || '(제목 없음)'}` }));
  formDialog('하이퍼링크 삽입', [
    { name: 'kind', label: '연결 대상', type: 'select', value: /^#slide/.test(cur ?? '') ? 'slide' : 'web', options: [{ value: 'web', label: '웹 페이지 또는 파일' }, { value: 'slide', label: '현재 문서의 슬라이드' }] },
    { name: 'url', label: '주소', type: 'text', value: /^#slide/.test(cur ?? '') ? '' : cur ?? 'https://' },
    { name: 'slide', label: '슬라이드', type: 'select', value: cur ?? '#slide1', options: slides },
  ], (v) => {
    let link = v.kind === 'slide' ? v.slide : v.url.trim();
    if (link === 'https://' || link === '') link = undefined;
    if (link && v.kind === 'web' && !/^(https?:|mailto:|tel:|#)/i.test(link)) link = `https://${link}`;
    if (S.editing) applyTextFormat({ link, u: link ? true : undefined, color: link ? '@hlink' : undefined }, null);
    else change(() => { o.link = link; });
  }, { onChange: (i) => { i.url.closest('label').style.display = i.kind.value === 'web' ? '' : 'none'; i.slide.closest('label').style.display = i.kind.value === 'slide' ? '' : 'none'; } });
}

const SYMBOLS = '※★☆●○◎◆◇■□▲△▼▽◀▶◁▷→←↑↓↔⇒⇔↗↘♥♡♠♣✓✔✗✘☑☐∑∏√∞±×÷≠≈≤≥°℃℉‰%€£¥₩$¢©®™§¶†‡…·•‥「」『』【】〈〉《》①②③④⑤⑥⑦⑧⑨⑩⑪⑫ⓐⓑⓒⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ㉠㉡㉢㉮㉯㉰㈜㈔℡☎☏✉✂✈☀☁☂☃♪♫αβγδεθλμπσΩ';
function symbolDialog() {
  const grid = el('div', { class: 'sym-grid' }, [...SYMBOLS].map((ch) => el('button', { class: 'sym', title: `U+${ch.codePointAt(0).toString(16).toUpperCase()}`, onclick: () => insertSymbol(ch) }, ch)));
  openDialog({ title: '기호', body: grid, width: 520, buttons: [{ label: '닫기' }] });
}
function insertSymbol(ch) {
  if (insertTextAtCaret(ch)) return;
  const o = selOne();
  if (o?.text) { startEdit(o, { end: true }); insertTextAtCaret(ch); return; }
  const { w, h } = S.pres.size;
  const tb = newTextBox({ x: w / 2 - 40, y: h / 2 - 30, w: 80, h: 60 }, ch);
  tb.text.paras[0].runs[0].size = 40;
  change(() => slide().objects.push(tb));
  S.sel = new Set([tb.id]);
  emit('selection');
}

function smartArtDialog() {
  let kind = 'process';
  const list = el('div', { class: 'smart-kinds' }, SMART_KINDS.map(([k, l]) => {
    const b = el('button', { class: `smart-kind${k === kind ? ' on' : ''}`, onclick: () => { kind = k; list.querySelectorAll('.smart-kind').forEach((x) => x.classList.toggle('on', x === b)); draw(); } }, l);
    return b;
  }));
  const items = el('textarea', { class: 'smart-items' }, '계획\n실행\n점검\n개선');
  const multi = el('input', { type: 'checkbox', checked: true });
  const preview = el('div', { class: 'smart-preview' });
  const draw = () => {
    const objs = smartArt(kind, items.value.split('\n'), { x: 40, y: 40, w: 1200, h: 640 }, { multicolor: multi.checked });
    const s = { id: 'pv', layout: 'blank', objects: objs, anims: [], hideDecor: true };
    const sc = 300 / 1280;
    preview.innerHTML = `<div class="tw" style="width:300px;height:${720 * sc}px"><div class="tw-in" style="transform:scale(${sc})">${slideHtml({ ...S.pres, size: { w: 1280, h: 720 } }, s, {})}</div></div>`;
  };
  items.addEventListener('input', draw);
  multi.addEventListener('change', draw);
  items.addEventListener('keydown', (e) => e.stopPropagation());
  openDialog({
    title: 'SmartArt 그래픽 선택', width: 640,
    body: el('div', { class: 'smart-dlg' }, list, el('div', { class: 'smart-right' }, preview, el('div', { class: 'muted' }, '한 줄에 항목 하나 (조직도는 첫 줄이 맨 위)'), items, el('label', { class: 'pane-check' }, multi, el('span', {}, '여러 색')))),
    buttons: [{ label: '확인', primary: true, action: () => {
      const { w, h } = S.pres.size;
      const ph = slide().objects.find((o) => o.ph === 'body' && isEmptyText(o.text));
      const rect = ph ? { x: ph.x, y: ph.y, w: ph.w, h: ph.h } : { x: w * 0.08, y: h * 0.25, w: w * 0.84, h: h * 0.62 };
      const objs = smartArt(kind, items.value.split('\n'), rect, { multicolor: multi.checked });
      change(() => { if (ph) slide().objects = slide().objects.filter((o) => o !== ph); slide().objects.push(...objs); });
      S.sel = new Set(objs.map((o) => o.id));
      emit('selection');
    } }, { label: '취소' }],
  });
  draw();
}

// ───────────── SmartArt 텍스트 창 · 디자인 ─────────────
/** 선택한 SmartArt 묶음: { members, items, kind, multi, rect, at } */
function currentSmart() {
  const o = selObjects().find((x) => x.smart) ?? (S.editing ? objById(S.editing.id) : null);
  if (!o?.smart || !o.grp) return null;
  const s = slide();
  const members = s.objects.filter((x) => x.grp === o.grp);
  const head = members.find((x) => x.smartItems);
  const items = head?.smartItems ?? members.filter((x) => x.text && !isEmptyText(x.text)).map((x) => plainText(x.text).replace(/\n/g, ' '));
  const x1 = Math.min(...members.map((m) => m.x)); const y1 = Math.min(...members.map((m) => m.y));
  const x2 = Math.max(...members.map((m) => m.x + m.w)); const y2 = Math.max(...members.map((m) => m.y + m.h));
  return { members, items, kind: o.smart, multi: head?.smartMulti ?? true, rect: { x: x1, y: y1, w: x2 - x1, h: y2 - y1 }, at: s.objects.indexOf(members[0]), grp: o.grp };
}
function rebuildSmart(cur, { kind = cur.kind, items = cur.items, multi = cur.multi } = {}) {
  const s = slide();
  const objs = smartArt(kind, items, cur.rect, { multicolor: multi });
  for (const o of objs) o.grp = cur.grp;
  change(() => {
    s.objects = s.objects.filter((o) => o.grp !== cur.grp);
    s.objects.splice(Math.min(cur.at, s.objects.length), 0, ...objs);
    // 애니메이션은 첫 도형으로 옮김
    const ids = new Set(cur.members.map((m) => m.id));
    for (const a of s.anims ?? []) if (ids.has(a.obj)) a.obj = objs[0].id;
  });
  S.sel = new Set(objs.map((o) => o.id));
  emit('selection');
}
function smartTextPane() {
  const cur = currentSmart();
  if (!cur) { toast('SmartArt 그래픽을 선택하세요'); return; }
  const ta = el('textarea', { class: 'smart-items', rows: 10 }, cur.items.join('\n'));
  ta.addEventListener('keydown', (e) => e.stopPropagation());
  let timer = 0;
  let live = cur;
  ta.addEventListener('input', () => { clearTimeout(timer); timer = setTimeout(() => { rebuildSmart(live, { items: ta.value.split('\n') }); live = currentSmart() ?? live; }, 350); });
  openDialog({ title: '텍스트 입력', width: 360, modeless: true, body: el('div', {}, el('p', { class: 'muted' }, '한 줄에 항목 하나 — 입력하는 대로 그래픽이 바뀝니다'), ta), buttons: [{ label: '닫기', primary: true }] });
}

// ───────────── 확대/축소 (PowerPoint [삽입 › 확대/축소]: 요약 · 구역 · 슬라이드) ─────────────
function zoomObject(slideId, rect, extra = {}) {
  return { id: uid(), type: 'zoom', x: rect.x, y: rect.y, w: rect.w, h: rect.h, rot: 0, zoom: { slide: slideId, ret: true, ...extra } };
}
/** 축소판을 격자로 배치 (슬라이드 비율 유지) */
function zoomGrid(n, area) {
  const { w: W, h: H } = S.pres.size;
  const cols = Math.ceil(Math.sqrt(n * (area.w / area.h) / (W / H)));
  const rows = Math.ceil(n / cols);
  const gap = 24;
  const tw = Math.min((area.w - gap * (cols - 1)) / cols, ((area.h - gap * (rows - 1)) / rows) * (W / H));
  const th = tw * (H / W);
  const x0 = area.x + (area.w - (cols * tw + (cols - 1) * gap)) / 2;
  const y0 = area.y + (area.h - (rows * th + (rows - 1) * gap)) / 2;
  return Array.from({ length: n }, (_, i) => ({ x: x0 + (i % cols) * (tw + gap), y: y0 + Math.floor(i / cols) * (th + gap), w: tw, h: th }));
}
function pickSlidesDialog(title, { sections = false } = {}) {
  return new Promise((resolve) => {
    const list = sections
      ? S.pres.slides.map((s, i) => [s, i]).filter(([s]) => s.section).map(([s, i]) => ({ s, i, label: `${s.section} (슬라이드 ${i + 1})` }))
      : S.pres.slides.map((s, i) => ({ s, i, label: `${i + 1}. ${slideTitle(s) || '(제목 없음)'}` }));
    if (!list.length) { toast(sections ? '구역이 없습니다. 먼저 [구역 추가]를 하세요' : '슬라이드가 없습니다'); resolve(null); return; }
    const checks = list.map((it) => { const c = el('input', { type: 'checkbox' }); return [it, c]; });
    const sc = 160 / S.pres.size.w;
    const grid = el('div', { class: 'zoom-pick' }, checks.map(([it, c]) => el('label', { class: 'zoom-item' },
      el('div', { class: 'tw', style: `width:160px;height:${S.pres.size.h * sc}px`, html: `<div class="tw-in" style="transform:scale(${sc})">${slideHtml(S.pres, it.s, { index: it.i })}</div>` }),
      el('span', {}, c, ` ${it.label}`))));
    openDialog({ title, width: 640, body: grid, buttons: [{ label: '삽입', primary: true, action: () => { const picked = checks.filter(([, c]) => c.checked).map(([it]) => it); if (!picked.length) { toast('슬라이드를 고르세요'); return false; } resolve(picked); return true; } }, { label: '취소', action: () => resolve(null) }] });
  });
}
async function insertZoom(kind) {
  if (kind === 'summary') {
    const picked = await pickSlidesDialog('요약 확대/축소 삽입');
    if (!picked) return;
    const { w, h } = S.pres.size;
    const first = Math.min(...picked.map((p) => p.i));
    const sum = newSlide(S.pres, 'titleOnly');
    const titleObj = sum.objects.find((o) => o.ph === 'title');
    if (titleObj) setPlainText(titleObj.text, '요약 확대/축소');
    const rects = zoomGrid(picked.length, { x: w * 0.06, y: h * 0.26, w: w * 0.88, h: h * 0.66 });
    change(() => {
      // 고른 슬라이드마다 구역 시작 (구역 확대/축소는 구역 끝까지 보여 주고 돌아옴)
      for (const p of picked) if (!p.s.section) p.s.section = slideTitle(p.s) || `구역 ${p.i + 1}`;
      sum.objects.push(...picked.map((p, k) => zoomObject(p.s.id, rects[k], { section: true })));
      if (!sum.section && first > 0) sum.section = '요약';
      S.pres.slides.splice(first, 0, sum);
    }, { scope: 'all' });
    goSlide(first);
    return;
  }
  const picked = await pickSlidesDialog(kind === 'section' ? '구역 확대/축소 삽입' : '슬라이드 확대/축소 삽입', { sections: kind === 'section' });
  if (!picked) return;
  const { w, h } = S.pres.size;
  const tw = w * 0.22;
  const objs = picked.map((p, k) => zoomObject(p.s.id, { x: w * 0.1 + k * 30, y: h * 0.3 + k * 30, w: tw, h: tw * (h / w) }, kind === 'section' ? { section: true } : {}));
  change(() => slide().objects.push(...objs));
  S.sel = new Set(objs.map((o) => o.id));
  emit('selection');
}

// ───────────── 수식 (PowerPoint [삽입 › 수식], Alt+=) ─────────────
const EQ_GALLERY = [
  ['근의 공식', 'x=\\frac{-b\\pm \\sqrt{b^2-4ac}}{2a}'],
  ['피타고라스 정리', 'a^2+b^2=c^2'],
  ['원의 넓이', 'A=\\pi r^2'],
  ['이항 정리', '(x+a)^n=\\sum_{k=0}^{n} \\left( \\begin{matrix}n \\\\ k\\end{matrix} \\right) x^k a^{n-k}'],
  ['오일러 공식', 'e^{i\\pi}+1=0'],
  ['테일러 급수', 'e^x=1+\\frac{x}{1!}+\\frac{x^2}{2!}+\\frac{x^3}{3!}+\\cdots'],
  ['삼각 항등식', '\\sin^2\\theta +\\cos^2\\theta =1'],
  ['정규 분포', 'f(x)=\\frac{1}{\\sigma \\sqrt{2\\pi}} e^{-\\frac{(x-\\mu)^2}{2\\sigma^2}}'],
  ['ROAS', '\\text{ROAS}=\\frac{\\text{매출}}{\\text{광고비}}\\times 100'],
];
const EQ_PALETTE = [
  ['a/b', '\\frac{a}{b}'], ['√x', '\\sqrt{x}'], ['ⁿ√x', '\\sqrt[n]{x}'], ['xⁿ', 'x^{n}'], ['xₙ', 'x_{n}'], ['∑', '\\sum_{i=1}^{n} '], ['∫', '\\int_{a}^{b} '], ['∏', '\\prod_{i=1}^{n} '],
  ['( )', '\\left( x \\right)'], ['[ ]', '\\left[ x \\right]'], ['행렬', '\\begin{pmatrix}a & b \\\\ c & d\\end{pmatrix}'], ['x̂', '\\hat{x}'], ['x̄', '\\bar{x}'], ['x⃗', '\\vec{x}'], ['lim', '\\lim_{x\\to 0} '], ['글', '\\text{글}'],
  ...['alpha', 'beta', 'gamma', 'delta', 'theta', 'lambda', 'mu', 'pi', 'sigma', 'omega', 'Delta', 'Sigma', 'times', 'div', 'pm', 'le', 'ge', 'ne', 'approx', 'infty', 'partial', 'to', 'cdot', 'cdots'].map((k) => [({ alpha: 'α', beta: 'β', gamma: 'γ', delta: 'δ', theta: 'θ', lambda: 'λ', mu: 'μ', pi: 'π', sigma: 'σ', omega: 'ω', Delta: 'Δ', Sigma: 'Σ', times: '×', div: '÷', pm: '±', le: '≤', ge: '≥', ne: '≠', approx: '≈', infty: '∞', partial: '∂', to: '→', cdot: '⋅', cdots: '⋯' })[k], `\\${k} `]),
];
function equationDialog(target = null) {
  const ta = el('textarea', { class: 'eq-src', rows: 4, spellcheck: 'false', placeholder: 'LaTeX 로 입력 (예: \\frac{a}{b}, x^2, \\sqrt{x}, \\sum_{i=1}^{n})' }, target?.latex ?? '');
  const pv = el('div', { class: 'eq-preview' });
  const size = el('input', { type: 'number', min: 8, max: 120, value: target?.size ?? 28, style: { width: '64px' } });
  const draw = () => { try { pv.innerHTML = toMathML(parseLatex(ta.value)); } catch { pv.textContent = '수식을 읽지 못했습니다'; } };
  const insert = (t) => { const a = ta.selectionStart; const b = ta.selectionEnd; ta.value = ta.value.slice(0, a) + t + ta.value.slice(b); ta.focus(); ta.selectionStart = ta.selectionEnd = a + t.length; draw(); };
  ta.addEventListener('input', draw);
  ta.addEventListener('keydown', (e) => e.stopPropagation());
  const pal = el('div', { class: 'eq-pal' }, EQ_PALETTE.map(([l, t]) => el('button', { type: 'button', class: 'btn small', title: t, onclick: () => insert(t) }, l)));
  const gal = el('div', { class: 'eq-gal' }, EQ_GALLERY.map(([l, t]) => { const b = el('button', { type: 'button', class: 'eq-card', title: l, onclick: () => { ta.value = t; draw(); } }, el('small', {}, l)); b.insertAdjacentHTML('beforeend', toMathML(parseLatex(t))); return b; }));
  openDialog({
    title: target ? '수식 편집' : '수식 삽입', width: 720,
    body: el('div', { class: 'eq-dlg' }, el('div', { class: 'muted' }, '기본 제공 수식'), gal, el('div', { class: 'muted' }, '구조 · 기호'), pal, ta, el('label', { class: 'pane-check' }, '글꼴 크기 ', size, ' pt'), pv),
    buttons: [{ label: target ? '확인' : '삽입', primary: true, action: () => {
      const latex = ta.value.trim();
      if (!latex) return false;
      const sz = Math.max(8, Math.min(120, Number(size.value) || 28));
      if (target) { change(() => { target.latex = latex; target.size = sz; }); return true; }
      const { w, h } = S.pres.size;
      const o = { id: uid(), type: 'equation', x: w * 0.25, y: h * 0.4, w: w * 0.5, h: Math.max(60, sz * 4), rot: 0, latex, size: sz };
      change(() => slide().objects.push(o));
      S.sel = new Set([o.id]);
      emit('selection');
      return true;
    } }, { label: '취소' }],
  });
  draw();
  setTimeout(() => ta.focus(), 30);
}

// ───────────── 그림 배경 제거 (PowerPoint [그림 서식 › 배경 제거]) ─────────────
async function removeBackground() {
  const o = selObjects().find((x) => x.type === 'image' && S.pres.media[x.media]);
  if (!o) { toast('배경을 제거할 그림을 선택하세요'); return; }
  const img = new Image();
  img.src = S.pres.media[o.bgOrig ?? o.media];
  try { await img.decode(); } catch { alertDialog('배경 제거', '이 그림 형식은 배경 제거를 지원하지 않습니다.'); return; }
  // 큰 그림은 줄여서 계산 (긴 쪽 1600px)
  const k = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
  const w = Math.max(1, Math.round(img.naturalWidth * k)); const h = Math.max(1, Math.round(img.naturalHeight * k));
  const src = document.createElement('canvas'); src.width = w; src.height = h;
  const g = src.getContext('2d', { willReadFrequently: true });
  g.drawImage(img, 0, 0, w, h);
  const data = g.getImageData(0, 0, w, h).data;
  const st = { tolerance: 25, remove: [], keep: [], mode: 'remove' };
  const view = el('canvas', { class: 'bgr-canvas', width: w, height: h });
  const vg = view.getContext('2d');
  let out = null;
  const draw = () => {
    const alpha = backgroundMask(data, w, h, st);
    out = applyMask(data, alpha);
    // 미리 보기: 지워질 곳은 PowerPoint 처럼 자홍색으로 덮어 보임
    const pv = new Uint8ClampedArray(data);
    for (let p = 0; p < alpha.length; p++) {
      const a = alpha[p] / 255;
      pv[p * 4] = Math.round(pv[p * 4] * a + 0xb4 * (1 - a));
      pv[p * 4 + 1] = Math.round(pv[p * 4 + 1] * a + 0x2d * (1 - a));
      pv[p * 4 + 2] = Math.round(pv[p * 4 + 2] * a + 0xb4 * (1 - a));
    }
    vg.putImageData(new ImageData(pv, w, h), 0, 0);
    for (const [x, y] of st.remove) { vg.fillStyle = '#d00'; vg.beginPath(); vg.arc(x, y, Math.max(4, w / 150), 0, 7); vg.fill(); }
    for (const [x, y] of st.keep) { vg.fillStyle = '#0a0'; vg.beginPath(); vg.arc(x, y, Math.max(4, w / 150), 0, 7); vg.fill(); }
  };
  view.addEventListener('click', (e) => {
    const r = view.getBoundingClientRect();
    const pt = [Math.round((e.clientX - r.left) * w / r.width), Math.round((e.clientY - r.top) * h / r.height)];
    (e.shiftKey || st.mode === 'keep' ? st.keep : st.remove).push(pt);
    draw();
  });
  const modeBtn = (m, label) => el('button', { type: 'button', class: `btn small${st.mode === m ? ' on' : ''}`, 'data-mode': m, onclick: (e) => { st.mode = m; e.target.parentElement.querySelectorAll('[data-mode]').forEach((b) => b.classList.toggle('on', b.dataset.mode === m)); } }, label);
  const tol = el('input', { type: 'range', min: 0, max: 100, value: st.tolerance });
  tol.addEventListener('change', () => { st.tolerance = Number(tol.value); draw(); });
  const undoMark = () => { const all = [...st.remove.map((p) => ['remove', p]), ...st.keep.map((p) => ['keep', p])]; if (!all.length) return; if (st.keep.length) st.keep.pop(); else st.remove.pop(); draw(); };
  openDialog({
    title: '배경 제거', width: 720,
    body: el('div', { class: 'bgr' },
      el('div', { class: 'pane-btns' }, modeBtn('keep', '✚ 보관할 영역 표시'), modeBtn('remove', '− 제거할 영역 표시'), el('button', { type: 'button', class: 'btn small', onclick: undoMark }, '표시 취소'),
        el('label', { class: 'bgr-tol' }, '허용 범위 ', tol)),
      el('p', { class: 'muted' }, '자홍색 부분이 지워집니다. 그림을 눌러 영역을 표시하세요 (Shift+클릭 = 보관).'),
      el('div', { class: 'bgr-wrap' }, view)),
    buttons: [{ label: '변경 내용 유지', primary: true, action: () => {
      const c = document.createElement('canvas'); c.width = w; c.height = h;
      c.getContext('2d').putImageData(new ImageData(out, w, h), 0, 0);
      const id = addMedia(c.toDataURL('image/png'));
      change(() => { o.bgOrig = o.bgOrig ?? o.media; o.media = id; });
    } }, { label: '모든 변경 내용 취소' }],
  });
  draw();
}

// ───────────── 아이콘 (WIXEL 과 같은 아이콘 모음) ─────────────
let iconLib = null;
async function loadIconLib() {
  if (iconLib) return iconLib;
  let res = await fetch('assets/iconlib.json.gz').catch(() => null);
  if (!res?.ok) res = await fetch('assets/iconlib.json').catch(() => null);
  if (!res?.ok) throw new Error(`아이콘 모음을 불러오지 못했습니다 (${res?.status ?? '연결 안 됨'}).`);
  let buf = new Uint8Array(await res.arrayBuffer());
  if (buf[0] === 0x1f && buf[1] === 0x8b) buf = new Uint8Array(await new Response(new Blob([buf]).stream().pipeThrough(new DecompressionStream('gzip'))).arrayBuffer());
  iconLib = JSON.parse(new TextDecoder().decode(buf)).map(([cat, items]) => [cat, items.map(([name, body, vb]) => ({ name, body, vb: vb ?? '0 0 96 96', cat }))]);
  return iconLib;
}
export const iconSvg = (ic, fill = '#000000') => `<svg xmlns="http://www.w3.org/2000/svg" viewBox="${ic.vb}" fill="${fill}">${ic.body.replace(/\sfill="(?!none)[^"]*"/g, '').replace(/fill:\s*#[0-9a-fA-F]{3,8};?/g, '')}</svg>`;
const svgUrl = (text) => { const u = new TextEncoder().encode(text); let s = ''; for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return `data:image/svg+xml;base64,${btoa(s)}`; };
async function iconsDialog() {
  let lib;
  try { lib = await loadIconLib(); } catch (e) { alertDialog('아이콘', `${e.message} 배포 파일(assets/iconlib.json.gz)을 확인하세요.`); return; }
  const all = lib.flatMap(([, items]) => items);
  const picked = new Set();
  let cat = '';
  const search = el('input', { type: 'search', placeholder: '아이콘 검색 (예: 사람, 화살표, 돈)', class: 'ic-search' });
  const cats = el('div', { class: 'ic-cats' });
  const grid = el('div', { class: 'ic-grid' });
  let list = [];
  let shown = 0;
  const more = () => {
    for (const ic of list.slice(shown, shown + 200)) {
      const b = el('button', { type: 'button', class: `ic-cell${picked.has(ic) ? ' on' : ''}`, title: ic.name, html: `<svg viewBox="${ic.vb}">${ic.body}</svg>` });
      b.addEventListener('click', () => { if (picked.has(ic)) picked.delete(ic); else picked.add(ic); b.classList.toggle('on', picked.has(ic)); });
      grid.append(b);
    }
    shown += 200;
  };
  const draw = () => {
    const q = search.value.trim().toLowerCase();
    const seen = new Set();
    list = (cat ? lib.find((x) => x[0] === cat)[1] : all).filter((ic) => (!q || ic.name.toLowerCase().includes(q) || ic.cat.toLowerCase().includes(q)) && (seen.has(ic.body) ? false : (seen.add(ic.body), true)));
    grid.replaceChildren(); shown = 0; more();
    cats.querySelectorAll('button').forEach((b) => b.classList.toggle('on', (b.dataset.cat ?? '') === cat));
  };
  grid.addEventListener('scroll', () => { if (grid.scrollTop + grid.clientHeight > grid.scrollHeight - 200 && shown < list.length) more(); });
  cats.append(el('button', { type: 'button', 'data-cat': '', onclick: () => { cat = ''; draw(); } }, '전체'), ...lib.map(([c, items]) => el('button', { type: 'button', 'data-cat': c, onclick: () => { cat = c; draw(); } }, `${c} `, el('small', {}, String(items.length)))));
  let t = 0;
  search.addEventListener('input', () => { clearTimeout(t); t = setTimeout(draw, 150); });
  search.addEventListener('keydown', (e) => e.stopPropagation());
  openDialog({
    title: '아이콘 삽입', width: 760,
    body: el('div', { class: 'ic-dlg' }, search, el('div', { class: 'ic-main' }, cats, grid)),
    buttons: [{ label: '삽입', primary: true, action: () => {
      const { w, h } = S.pres.size;
      const objs = [...picked].map((ic, k) => {
        const id = addMedia(svgUrl(iconSvg(ic)));
        const o = newImage(id, { x: w / 2 - 64 + k * 24, y: h / 2 - 64 + k * 24, w: 128, h: 128 });
        o.icon = { vb: ic.vb, body: ic.body, fill: '#000000' };
        o.name = ic.name;
        return o;
      });
      if (!objs.length) return false;
      change(() => slide().objects.push(...objs));
      S.sel = new Set(objs.map((o) => o.id));
      emit('selection');
      return true;
    } }, { label: '취소' }],
  });
  draw();
}

function designIdeasDialog() {
  const s = slide();
  const title = s.objects.find((o) => o.ph === 'title' || o.ph === 'ctrTitle');
  const body = s.objects.find((o) => o.ph === 'body') ?? s.objects.find((o) => o.type !== 'shape' || (o.text && !o.ph));
  const make = (idea) => {
    const c = duplicateSlide(s);
    const { w: W, h: H } = S.pres.size;
    const R = (r) => ({ x: r.x * W, y: r.y * H, w: r.w * W, h: r.h * H });
    c.bg = idea.bg ?? c.bg;
    c.hideDecor = true;
    c.bgObjects = undefined;
    const extra = [];
    if (idea.band) { const b = newShape(idea.band.shape ?? 'rect', R(idea.band)); b.fill = { type: 'solid', color: idea.band.color }; b.line = null; b.text = null; b.name = '디자인 배경'; extra.push(b); }
    if (idea.card) { const b = newShape('roundRect', R(idea.card)); b.adj = { adj: 4000 }; b.fill = { type: 'solid', color: '@bg1' }; b.line = null; b.shadow = true; b.text = null; b.name = '디자인 카드'; extra.push(b); }
    const t = c.objects.find((o) => o.ph === 'title' || o.ph === 'ctrTitle');
    if (t && idea.title) { Object.assign(t, R(idea.title)); for (const p of t.text.paras) { p.align = idea.title.align; if (idea.title.color) for (const r of p.runs) r.color = idea.title.color; if (idea.title.color) p.end = { ...(p.end ?? {}), color: idea.title.color }; } }
    const b = body ? c.objects[s.objects.indexOf(body)] : null;
    if (b && idea.body) { Object.assign(b, R(idea.body)); if (b.text && idea.body.color) for (const p of b.text.paras) { for (const r of p.runs) r.color = idea.body.color; p.end = { ...(p.end ?? {}), color: idea.body.color }; } }
    c.objects.unshift(...extra);
    return c;
  };
  if (!title && !body) { toast('제목이나 내용이 있는 슬라이드에서 사용할 수 있습니다'); return; }
  const variants = DESIGN_IDEAS.map((d) => ({ d, s: make(d) }));
  const sc = 220 / S.pres.size.w;
  const grid = el('div', { class: 'ideas' }, variants.map(({ d, s: v }) => el('button', { class: 'idea', title: d.label, onclick: () => { closeAll(); change(() => { S.pres.slides[S.cur] = { ...v, id: s.id }; }); } }, el('div', { class: 'tw', style: { width: '220px', height: `${S.pres.size.h * sc}px` }, html: `<div class="tw-in" style="transform:scale(${sc})">${slideHtml(S.pres, v, { index: S.cur })}</div>` }), el('span', {}, d.label))));
  const dlg = openDialog({ title: '디자인 아이디어', body: grid, width: 760, buttons: [{ label: '닫기' }] });
  const closeAll = () => dlg.close();
}

function accessibilityDialog() {
  const issues = [];
  S.pres.slides.forEach((s, i) => {
    if (!slideTitle(s).trim()) issues.push([i, '슬라이드 제목 없음', '화면 읽기 프로그램이 슬라이드를 구분할 수 있도록 제목을 넣으세요.']);
    for (const o of s.objects) {
      if (o.type === 'image' && o.media && !o.alt) issues.push([i, '대체 텍스트 없음', `그림${o.name ? ` (${o.name})` : ''}에 설명을 넣으세요.`]);
      if (o.type === 'table' && !o.style?.firstRow) issues.push([i, '표 머리글 없음', '표의 첫 행을 머리글 행으로 지정하세요.']);
      if (o.type === 'chart' && !o.alt) issues.push([i, '차트 설명 없음', '차트에 대체 텍스트를 넣으세요.']);
    }
  });
  const body = issues.length ? el('table', { class: 'kbd-table' }, issues.map(([i, t, d]) => el('tr', { class: 'clickable', onclick: () => goSlide(i) }, el('td', {}, `슬라이드 ${i + 1}`), el('td', {}, el('b', {}, t), el('div', { class: 'muted' }, d))))) : el('p', {}, '문제가 없습니다. 접근성 검사를 통과했습니다.');
  openDialog({ title: `접근성 검사 (${issues.length}건)`, body, width: 560, buttons: [{ label: '닫기' }] });
}

function statsDialog() {
  let words = 0; let chars = 0; let pics = 0; let notesWords = 0;
  for (const s of S.pres.slides) {
    for (const o of s.objects) {
      const t = plainText(o.text) + (o.type === 'table' ? o.rows.flatMap((r) => r.cells.map((c) => plainText(c.text))).join(' ') : '');
      words += t.split(/\s+/).filter(Boolean).length;
      chars += t.replace(/\s/g, '').length;
      if (o.type === 'image' && o.media) pics++;
    }
    notesWords += (s.notes ?? '').split(/\s+/).filter(Boolean).length;
  }
  const rows = [['슬라이드', S.pres.slides.length], ['숨긴 슬라이드', S.pres.slides.filter((s) => s.hidden).length], ['단어', words], ['글자 (공백 제외)', chars], ['메모 단어', notesWords], ['그림', pics], ['애니메이션', S.pres.slides.reduce((n, s) => n + (s.anims?.length ?? 0), 0)]];
  openDialog({ title: '통계', body: el('table', { class: 'kbd-table' }, rows.map(([k, v]) => el('tr', {}, el('td', {}, k), el('td', {}, String(v))))), buttons: [{ label: '닫기', primary: true }] });
}

function spellDialog() {
  const items = [];
  S.pres.slides.forEach((s, si) => s.objects.forEach((o) => { if (o.text && !isEmptyText(o.text)) items.push({ si, o }); }));
  if (!items.length) { toast('검사할 텍스트가 없습니다'); return; }
  const areas = items.map(({ si, o }) => {
    const ta = el('textarea', { spellcheck: 'true', lang: 'ko', rows: '2' }, plainText(o.text));
    ta.addEventListener('keydown', (e) => e.stopPropagation());
    return { si, o, ta, row: el('label', { class: 'spell-row' }, el('span', {}, `슬라이드 ${si + 1}`), ta) };
  });
  openDialog({
    title: '맞춤법 검사 (브라우저 맞춤법 검사기 사용)', width: 620,
    body: el('div', { class: 'spell-list' }, el('div', { class: 'muted' }, '빨간 밑줄이 있는 단어를 오른쪽 클릭하면 추천 단어가 나옵니다. 고친 뒤 [적용]을 누르세요. (글자 서식은 단락 첫 서식으로 맞춰집니다)'), areas.map((a) => a.row)),
    buttons: [{ label: '적용', primary: true, action: () => {
      change(() => { for (const a of areas) if (a.ta.value !== plainText(a.o.text)) setPlainText(a.o.text, a.ta.value); }, { scope: 'all' });
    } }, { label: '취소' }],
  });
}

function cropDialog() {
  const o = selOne();
  if (o?.type !== 'image' || !o.media) { toast('그림을 선택하세요'); return; }
  const c = { l: 0, t: 0, r: 0, b: 0, ...(o.crop ?? {}) };
  const pct = (v) => String(Math.round(v * 1000) / 10);
  formDialog('그림 자르기 (%)', [
    { name: 'l', label: '왼쪽', type: 'number', value: pct(c.l) }, { name: 'r', label: '오른쪽', type: 'number', value: pct(c.r) },
    { name: 't', label: '위쪽', type: 'number', value: pct(c.t) }, { name: 'b', label: '아래쪽', type: 'number', value: pct(c.b) },
  ], (v) => {
    const n = { l: Number(v.l) / 100, r: Number(v.r) / 100, t: Number(v.t) / 100, b: Number(v.b) / 100 };
    if (n.l + n.r >= 0.98 || n.t + n.b >= 0.98) { alertDialog('자르기', '너무 많이 잘랐습니다.'); return false; }
    // 보이는 크기 유지 비율로 개체 크기 조정
    const fullW = o.w / (1 - c.l - c.r);
    const fullH = o.h / (1 - c.t - c.b);
    change(() => { o.x += (n.l - c.l) * fullW; o.y += (n.t - c.t) * fullH; o.w = fullW * (1 - n.l - n.r); o.h = fullH * (1 - n.t - n.b); o.crop = n; });
    return true;
  }, { note: '각 가장자리에서 잘라 낼 비율' });
}

function slideSizeDialog() {
  const { w, h } = S.pres.size;
  const cm = (px) => String(Math.round(px * 2.54 / 96 * 100) / 100);
  formDialog('슬라이드 크기', [
    { name: 'preset', label: '슬라이드 크기', type: 'select', value: 'custom', options: [{ value: 'custom', label: '사용자 지정' }, { value: 'wide', label: '와이드스크린 (33.867x19.05cm)' }, { value: 'standard', label: '화면 슬라이드 쇼 4:3 (25.4x19.05cm)' }, { value: 'a4', label: 'A4 용지 (27.517x19.05cm)' }, { value: '1610', label: '화면 16:10 (25.4x15.875cm)' }] },
    { name: 'w', label: '너비 (cm)', type: 'number', value: cm(w) },
    { name: 'h', label: '높이 (cm)', type: 'number', value: cm(h) },
    { name: 'mode', label: '내용 크기', type: 'select', value: 'fit', options: [{ value: 'fit', label: '맞춤 확인 (내용을 줄여 모두 보이게)' }, { value: 'max', label: '최대화 (내용을 키움)' }] },
  ], (v) => {
    const map = { wide: [1280, 720], standard: [960, 720], a4: [1040, 720], 1610: [960, 600] };
    const [nw, nh] = map[v.preset] ?? [(Number(v.w) * 96) / 2.54, (Number(v.h) * 96) / 2.54];
    if (!(nw > 50 && nh > 50 && nw < 8000 && nh < 8000)) { alertDialog('슬라이드 크기', '크기가 올바르지 않습니다.'); return false; }
    change(() => resizePresentation(S.pres, nw, nh, v.mode), { scope: 'all' });
    fitZoom();
    return true;
  });
}

/** 메모가 있는 다음/이전 슬라이드로 */
function jumpComment(d) {
  const n = S.pres.slides.length;
  for (let k = 1; k <= n; k++) {
    const i = (S.cur + d * k + n * 2) % n;
    if (S.pres.slides[i].comments?.length) { goSlide(i); S.formatPane = 'comments'; renderPane(); return; }
  }
  toast('메모가 없습니다');
}
function setupShowDialog() {
  const sh = S.pres.show ?? {};
  formDialog('슬라이드 쇼 설정', [
    { name: 'loop', label: 'Esc 키를 누를 때까지 계속 실행', type: 'checkbox', value: !!sh.loop },
    { name: 'useTimings', label: '설정된 시간 사용 (자동 넘기기)', type: 'checkbox', value: sh.useTimings !== false },
  ], (v) => change(() => { S.pres.show = { ...sh, loop: v.loop, useTimings: v.useTimings }; }, { scope: 'none' }));
}

// ───────────── 사용자 지정 쇼 ─────────────
function customShowsDialog() {
  const list = () => S.pres.customShows ?? [];
  const box = el('div', { class: 'cs-list' });
  const draw = () => {
    box.replaceChildren(...(list().length ? list().map((c) => el('div', { class: 'cs-row' },
      el('b', {}, c.name), el('small', {}, ` 슬라이드 ${c.slides.length}장`),
      el('span', { class: 'spacer' }),
      el('button', { class: 'btn', onclick: () => editCustomShow(c, draw) }, '편집'),
      el('button', { class: 'btn', onclick: () => change(() => { S.pres.customShows = list().filter((x) => x !== c); }, { scope: 'none' }) || draw() }, '제거'),
      el('button', { class: 'btn primary', onclick: () => { closeDialogs(); startShow({ custom: c.name }); } }, '쇼 보기'))) : [el('p', { class: 'muted' }, '사용자 지정 쇼가 없습니다. [새로 만들기]로 일부 슬라이드만 골라 순서를 정해 쇼를 만드세요.')]));
  };
  draw();
  openDialog({ title: '사용자 지정 쇼', width: 560, body: box, buttons: [{ label: '새로 만들기...', action: () => { editCustomShow(null, draw); return false; } }, { label: '닫기', primary: true }] });
}
function closeDialogs() { for (const d of document.querySelectorAll('#dialogLayer .dialog-backdrop')) d.remove(); }
function editCustomShow(cs, done) {
  const slides = S.pres.slides;
  const chosen = cs ? [...cs.slides] : [];
  const name = el('input', { type: 'text', value: cs?.name ?? `사용자 지정 쇼 ${(S.pres.customShows?.length ?? 0) + 1}` });
  const left = el('select', { multiple: true, size: 12, class: 'cs-sel' });
  const right = el('select', { multiple: true, size: 12, class: 'cs-sel' });
  const label = (id) => { const i = slides.findIndex((x) => x.id === id); return `${i + 1}. ${slideTitle(slides[i]) || '(제목 없음)'}`; };
  const draw = () => {
    left.replaceChildren(...slides.map((s2) => el('option', { value: s2.id }, label(s2.id))));
    right.replaceChildren(...chosen.map((id, k) => el('option', { value: String(k) }, label(id))));
  };
  draw();
  const add = () => { for (const o of left.selectedOptions) chosen.push(o.value); draw(); };
  const remove = () => { const ks = new Set([...right.selectedOptions].map((o) => Number(o.value))); for (let k = chosen.length - 1; k >= 0; k--) if (ks.has(k)) chosen.splice(k, 1); draw(); };
  const move = (d) => { const k = right.selectedIndex; const j = k + d; if (k < 0 || j < 0 || j >= chosen.length) return; [chosen[k], chosen[j]] = [chosen[j], chosen[k]]; draw(); right.selectedIndex = j; };
  const body = el('div', { class: 'form-grid' },
    el('label', {}, el('span', {}, '쇼 이름'), name),
    el('div', { class: 'cs-cols' },
      el('div', {}, el('small', {}, '프레젠테이션 슬라이드'), left),
      el('div', { class: 'cs-mid' }, el('button', { class: 'btn', onclick: add }, '추가 ›'), el('button', { class: 'btn', onclick: remove }, '‹ 제거'), el('button', { class: 'btn', onclick: () => move(-1) }, '▲'), el('button', { class: 'btn', onclick: () => move(1) }, '▼')),
      el('div', {}, el('small', {}, '사용자 지정 쇼 슬라이드'), right)));
  openDialog({ title: cs ? '사용자 지정 쇼 편집' : '사용자 지정 쇼 정의', width: 640, body, buttons: [{ label: '확인', primary: true, action: () => {
    const nm = name.value.trim();
    if (!nm || !chosen.length) { toast('이름과 슬라이드를 정하세요'); return false; }
    change(() => {
      const all = (S.pres.customShows ?? []).filter((x) => x !== cs && x.name !== nm);
      all.push({ name: nm, slides: chosen });
      S.pres.customShows = all;
    }, { scope: 'none' });
    done?.();
    return true;
  } }, { label: '취소' }] });
}

// ───────────── 메모 (댓글) ─────────────
function commentsList() { return slide().comments ?? []; }
function addComment(text, at = null) {
  const s = slide();
  const o = selOne();
  const pos = at ?? (o ? { x: o.x + o.w, y: o.y } : { x: 20, y: 20 });
  const c = { id: uid('c'), author: S.userName ?? localStorage.getItem('wipoint:user') ?? '사용자', text, at: new Date().toISOString(), x: Math.round(pos.x), y: Math.round(pos.y), replies: [] };
  change(() => { s.comments = [...(s.comments ?? []), c]; });
  return c;
}

// ───────────── 동영상 내보내기 ─────────────
function exportVideoDialog() {
  formDialog('동영상 만들기', [
    { name: 'q', label: '화질', type: 'select', value: '1280', options: [{ value: '1920', label: '전체 HD (1080p)' }, { value: '1280', label: 'HD (720p)' }, { value: '854', label: '표준 (480p)' }] },
    { name: 'sec', label: '슬라이드당 시간 (초)', type: 'number', value: '5' },
    { name: 'timings', label: '설정된 시간 · 전환 사용', type: 'checkbox', value: true },
  ], (v) => {
    const bar = el('div', { class: 'progress' }, el('i'));
    const msg = el('p', {}, '준비 중...');
    const ac = new AbortController();
    const dlg = openDialog({ title: '동영상 만드는 중', width: 420, body: el('div', {}, msg, bar, el('p', { class: 'muted' }, '슬라이드 시간만큼 실시간으로 녹화합니다. 이 탭을 닫거나 숨기지 마세요.')), buttons: [{ label: '취소', action: () => { ac.abort(); return true; } }] });
    exportVideo({ width: Number(v.q), sec: Math.max(0.5, Number(v.sec) || 5), useTimings: v.timings }, (p, t) => { bar.firstChild.style.width = `${Math.round(p * 100)}%`; msg.textContent = t; }, ac.signal)
      .then((r) => { dlg?.close?.(); closeDialogs(); if (r) toast(`동영상(.${r.ext}, ${Math.round(r.size / 1048576 * 10) / 10}MB)을 저장했습니다`); })
      .catch((e) => { closeDialogs(); alertDialog('동영상 만들기', e.message); });
  }, { okLabel: '동영상 만들기' });
}

function zoomDialog() {
  formDialog('확대/축소', [{ name: 'z', label: '배율 (%)', type: 'combo', value: String(Math.round(S.zoom * 100)), options: ['400', '200', '100', '66', '50', '33'].map((v) => ({ value: v })) }], (v) => { const z = Number(v.z); if (z > 0) setZoom(z / 100); });
}

const SHORTCUTS = [
  ['Alt+=', '수식 삽입 (LaTeX 입력, PowerPoint 수식으로 저장)'],
  ['Ctrl+N / Ctrl+O / Ctrl+S', '새로 만들기 / 열기 / 저장'], ['F12', '다른 이름으로 저장'], ['Ctrl+P', '인쇄'], ['Ctrl+M', '새 슬라이드'], ['Ctrl+D', '복제 (개체 · 슬라이드)'],
  ['F5 / Shift+F5', '처음부터 / 현재 슬라이드부터 쇼'], ['Alt+F5', '발표자 보기'], ['Ctrl+Z / Ctrl+Y', '실행 취소 / 다시 실행'], ['Ctrl+C / X / V', '복사 / 잘라내기 / 붙여넣기'],
  ['Ctrl+Shift+C / V', '서식 복사 / 서식 붙여넣기'], ['Ctrl+G / Ctrl+Shift+G', '그룹 / 그룹 해제'], ['Ctrl+B / I / U', '굵게 / 기울임꼴 / 밑줄'], ['Ctrl+Shift+> / <', '글꼴 크기 크게 / 작게'],
  ['Ctrl+L / E / R / J', '왼쪽 / 가운데 / 오른쪽 / 양쪽 맞춤'], ['Ctrl+Space', '글자 서식 지우기'], ['Tab / Shift+Tab', '목록 수준 늘림 / 줄임 (글 편집 중)'], ['Ctrl+F / Ctrl+H', '찾기 / 바꾸기'],
  ['Ctrl+K', '하이퍼링크'], ['Ctrl+A', '모두 선택'], ['방향키 · Ctrl+방향키', '개체 이동 (작게)'], ['Shift+끌기', '가로/세로로만 이동, 비율 유지'], ['Ctrl+끌기', '복사하며 이동'], ['Alt+끌기', '스마트 가이드 끄기'],
  ['Ctrl+]/[ (Shift)', '앞으로 / 뒤로 (맨 앞 / 맨 뒤)'], ['Ctrl+마우스 휠', '확대/축소'], ['F2 / Enter', '선택한 개체 글 편집'], ['Esc', '편집 끝내기 · 선택 해제'],
  ['Alt 또는 F10', '키 팁 표시 (예: Alt, H, F, S = 글꼴 크기)'], ['Alt+1 … Alt+9', '빠른 실행 도구 모음 명령 (기본: 가로 가운데 · 세로 가운데 · 텍스트 상자 · 스포이트)'],
  ['Shift+방향키 / Ctrl+Shift+방향키', '개체 크기 조절 (크게 / 조금씩)'], ['Alt+←/→ / Ctrl+Alt+←/→', '15도 / 1도 회전'], ['Alt+Shift+←/→', '단락 수준 올림 / 내림'], ['Alt+Shift+↑/↓', '단락을 위/아래로 이동 (글 편집 중)'],
  ['Ctrl+T / Ctrl+Shift+F / Ctrl+Shift+P', '글꼴 대화 상자'], ['Shift+F3', '대/소문자 바꾸기 (소문자 → 대문자 → 단어 첫 글자)'], ['Ctrl+= / Ctrl+Shift+=', '아래 첨자 / 위 첨자'],
  ['Ctrl+Enter', '다음 개체 틀 (마지막이면 새 슬라이드)'], ['F4', '마지막 작업 반복'], ['Alt+F9 / Shift+F9', '안내선 / 눈금선 표시'], ['Alt+F10', '선택 창'],
  ['F6 / Shift+F6', '창 사이 이동 (리본 · 슬라이드 목록 · 편집 화면 · 노트 · 상태 표시줄)'], ['Shift+F10 / 메뉴 키', '바로 가기 메뉴'], ['Ctrl+F1', '리본 접기/펴기'], ['Ctrl+Alt+V', '선택하여 붙여넣기'], ['Ctrl+F12 / Ctrl+F2', '열기 / 인쇄'],
  ['쇼: N / P / 숫자+Enter', '다음 / 이전 / 해당 슬라이드로'], ['쇼: B (.) / W (,)', '검정 / 흰 화면'], ['쇼: Ctrl+P / Ctrl+L / Ctrl+A', '펜 / 레이저 포인터 / 화살표'], ['쇼: E / Ctrl+E / Ctrl+M', '잉크 지우기 / 지우개 / 잉크 숨기기'],
  ['쇼: Ctrl+H / Ctrl+U', '포인터 숨기기 / 보이기'], ['쇼: G 또는 Ctrl+S', '모든 슬라이드 목록'], ['쇼: Home / End / Esc (−)', '첫 / 마지막 슬라이드 / 쇼 마치기'],
];

export const APP_VERSION = '1.0';
const WHATS_NEW = [
  'PowerPoint 와 같은 리본 · 슬라이드 미리 보기 · 슬라이드 노트 · 여러 슬라이드 보기',
  '.pptx 열기 · 저장 (마스터/레이아웃 상속, 테마 색, 표, 차트, SmartArt 그림, 그룹, 전환, 애니메이션, 메모)',
  '도형 60여 종 · 스마트 가이드 · 맞춤/배분 · 그룹 · 회전 · 빠른 스타일 · WordArt',
  '전환 효과 12종 · 애니메이션(나타내기/강조/끝내기) · 애니메이션 창',
  '슬라이드 쇼: 발표자 보기 · 청중 창 · 펜/레이저 포인터 · 예행 연습 · 자동 넘기기',
  '서식 파일 9종 (사업 계획서 · 마케팅 성과 보고 · 제안서 · 강의 · 회의 · 프로젝트 · 서비스 소개 · 포토 앨범)',
  'SmartArt 도식 9종 · 아이콘 3,600여 개 · 차트 7종 · 디자인 아이디어',
  'PDF(인쇄) · PNG 내보내기 · 읽기 전용 공유 링크 · 자동 저장 · 최근 문서',
  'PDF 열기: 쪽마다 슬라이드로 — 글은 텍스트 상자, 도형 · 선 · 그림은 그대로 편집',
  'pptx 에 포함된 글꼴(.fntdata, MTX 압축 EOT 포함)을 읽어 그대로 표시 · 다시 저장해도 유지',
  '키 팁: Alt 또는 F10 → 모든 탭 · 명령에 글자 (예: Alt, H, F, S = 글꼴 크기), 메뉴는 화살표 · 글자로 고르기',
  '빠른 실행 도구 모음 (리본 아래): Alt+1 가로 가운데 · Alt+2 세로 가운데 · Alt+3 텍스트 상자 · Alt+4 스포이트, 리본 단추 오른쪽 클릭으로 추가',
  '비디오 · 오디오 넣기 (재생 탭: 자동 실행 · 반복 · 시작/끝 시간), pptx 로 저장해도 PowerPoint 에서 재생',
  '표: 칸 끌어서 범위 선택 → 세로 · 가로 병합, 셀 분할(열/행 개수), 범위 테두리 · 대각선 · 펜 색/두께, 열 너비 · 행 높이 끌기',
  '차트 데이터 통합 문서 포함 → PowerPoint 에서 [데이터 편집] 가능',
  '모핑 전환 · 단락별 애니메이션 · 이동 경로(선 · 호 · 원 · 물결)',
  '메모(댓글 · 답글 · 해결), 사용자 지정 슬라이드 쇼, 그림을 화면에서 자르기(두 번 클릭), 동영상 만들기(MP4/WebM)',
  'PowerPoint 단축키 추가: Shift/Ctrl+Shift+방향키 크기, Alt+방향키 회전, Ctrl+T 글꼴, Shift+F3, F4, F6, Alt+F9, Shift+F9, Alt+F10, Ctrl+Enter, Alt+Shift+방향키 등',
];

// ───────────── 명령 등록 ─────────────
register({
  // 파일
  newPres: () => run('backstage', 'new'),
  newBlank: () => { setDocument(buildTemplate('blank'), '프레젠테이션1'); },
  open: () => openWithPicker().catch((e) => alertDialog('열기', e.message)),
  save: () => save().catch((e) => alertDialog('저장', e.message)),
  saveAs: (kind = 'pptx') => saveAs(kind).catch((e) => alertDialog('저장', e.message)),
  print: () => run('backstage', 'print'),
  printNow: (opts) => printSlides(opts),
  exportPng: (all) => exportPng(all).catch((e) => alertDialog('내보내기', e.message)),
  share: async () => {
    const { url, size } = await shareLink();
    const inp = el('input', { type: 'text', value: url, readonly: true, style: { width: '100%' } });
    openDialog({ title: '읽기 전용 링크로 공유', width: 560, body: el('div', { class: 'form-grid' }, el('p', {}, '이 링크를 받은 사람은 프레젠테이션을 보기(슬라이드 쇼)만 할 수 있습니다. 문서 내용이 링크 안에 압축되어 들어 있어 서버에 올라가지 않습니다.'), inp, size > 60000 ? el('p', { class: 'warn' }, `링크가 깁니다 (${Math.round(size / 1024)}KB). 그림이 많으면 메신저에서 잘릴 수 있습니다 — .pptx 파일로 보내는 것을 권장합니다.`) : null), buttons: [{ label: '링크 복사', primary: true, action: () => { navigator.clipboard?.writeText(url).then(() => toast('링크를 복사했습니다')); } }, { label: '닫기' }], onOpen: () => inp.select() });
  },

  // 실행 취소
  undo: () => { if (S.editing) endEdit(); if (S.history.doUndo(S.pres)) { S.sel = new Set([...S.sel].filter((id) => objById(id))); if (S.cur >= S.pres.slides.length) S.cur = S.pres.slides.length - 1; refresh('all'); } },
  redo: () => { if (S.editing) endEdit(); if (S.history.doRedo(S.pres)) refresh('all'); },

  // 클립보드
  copy: () => { if (S.focusThumbs) copySlides(); else if (!copyObjects()) toast('복사할 개체를 선택하세요'); },
  cut: () => { if (S.focusThumbs) copySlides(true); else if (!copyObjects(true)) toast('잘라낼 개체를 선택하세요'); },
  paste: async () => {
    try {
      if (navigator.clipboard?.read) {
        const items = await navigator.clipboard.read();
        for (const it of items) {
          const type = it.types.find((t) => t.startsWith('image/'));
          if (type && (!S.clipboard || Date.now() - S.clipboard.at > 1500)) { const blob = await it.getType(type); await insertImages([new File([blob], 'clip.png', { type })]); return; }
        }
      }
    } catch { /* 권한 없음 → 내부 클립보드 */ }
    if (!pasteInternal()) {
      const text = await navigator.clipboard?.readText?.().catch(() => '');
      if (text) run('pasteText', text); else toast('붙여 넣을 내용이 없습니다');
    }
  },
  pasteInternal: () => pasteInternal(),
  pasteText: (text) => {
    if (S.editing) { insertTextAtCaret(text); return; }
    const { w, h } = S.pres.size;
    const tb = newTextBox({ x: w * 0.2, y: h * 0.4, w: w * 0.6, h: 60 }, '');
    tb.text.paras = text.replace(/\r/g, '').split('\n').map((l) => para(l));
    change(() => slide().objects.push(tb));
    S.sel = new Set([tb.id]);
    emit('selection');
  },
  pasteFiles: (files) => insertImages(files),
  pasteMenu: (a) => menuAt(a, [
    { label: '붙여넣기', icon: 'paste', key: 'Ctrl+V', action: () => run('paste') },
    { label: '텍스트만 유지', icon: 'textbox', action: async () => { const t = await navigator.clipboard?.readText?.().catch(() => ''); if (t) run('pasteText', t); } },
    { label: '서식 붙여넣기', icon: 'painter', key: 'Ctrl+Shift+V', action: () => run('pasteFormat') },
  ]),
  duplicate: () => {
    if (S.focusThumbs || !selObjects().length) { run('duplicateSlide'); return; }
    const objs = selObjects();
    const copies = cloneObjects(objs, 18, 18);
    change(() => slide().objects.push(...copies));
    S.sel = new Set(copies.map((o) => o.id));
    emit('selection');
  },
  formatPainter: (keep = false) => {
    const o = selObjects()[0];
    if (!o) { toast('서식을 복사할 개체를 선택하세요'); return; }
    if (S.painter) { S.painter = null; emit('painter'); return; }
    S.painter = { style: captureStyle(o), keep };
    emit('painter');
    toast('서식을 적용할 개체를 클릭하세요 (Esc: 취소)');
  },
  copyFormat: () => { const o = selObjects()[0]; if (o) { S.formatClip = captureStyle(o); toast('서식을 복사했습니다'); } },
  pasteFormat: () => { if (!S.formatClip) { toast('먼저 서식을 복사하세요 (Ctrl+Shift+C)'); return; } applyStyle(selObjects(), S.formatClip); },
  applyPainter: () => { if (!S.painter) return; applyStyle(selObjects(), S.painter.style); if (!S.painter.keep) { S.painter = null; emit('painter'); } },

  // 슬라이드
  newSlide: (layout) => addSlide(layout),
  newSlideMenu: (a) => layoutPicker(a, (k) => addSlide(k)),
  layoutMenu: (a) => layoutPicker(a, (k) => change(() => { for (const s of pickedSlides()) changeLayout(S.pres, s, k); }, { scope: 'all' })),
  resetSlide: () => change(() => { for (const s of pickedSlides()) resetSlideLayout(S.pres, s); }),
  duplicateSlide: () => {
    const list = pickedSlides();
    const last = S.pres.slides.indexOf(list[list.length - 1]);
    change(() => { S.pres.slides.splice(last + 1, 0, ...list.map((s) => duplicateSlide(s))); }, { scope: 'all' });
    goSlide(last + 1);
  },
  deleteSlide: () => {
    const list = pickedSlides();
    if (list.length >= S.pres.slides.length) { if (S.pres.slides.length === 1) { change(() => { S.pres.slides = [newSlide(S.pres, 'title')]; }, { scope: 'all' }); goSlide(0); return; } }
    const first = S.pres.slides.indexOf(list[0]);
    const ids = new Set(list.map((s) => s.id));
    change(() => { S.pres.slides = S.pres.slides.filter((s) => !ids.has(s.id)); if (!S.pres.slides.length) S.pres.slides.push(newSlide(S.pres, 'title')); }, { scope: 'all' });
    goSlide(Math.min(first, S.pres.slides.length - 1));
  },
  hideSlide: () => { const list = pickedSlides(); const v = !list[0].hidden; change(() => { for (const s of list) s.hidden = v || undefined; }, { scope: 'all' }); },
  moveSlidesTo: (to, from) => {
    const sl = S.pres.slides;
    const moving = S.slideSel.has(sl[from]?.id) ? sl.filter((s) => S.slideSel.has(s.id)) : [sl[from]];
    if (!moving[0]) return;
    const ids = new Set(moving.map((s) => s.id));
    const before = sl.slice(0, to).filter((s) => !ids.has(s.id)).length;
    change(() => { const rest = sl.filter((s) => !ids.has(s.id)); rest.splice(before, 0, ...moving); S.pres.slides = rest; }, { scope: 'all' });
    S.cur = S.pres.slides.indexOf(moving[0]);
    refresh('all');
  },
  moveSlideUp: () => { if (S.cur > 0) run('moveSlidesTo', S.cur - 1, S.cur); },
  moveSlideDown: () => { if (S.cur < S.pres.slides.length - 1) run('moveSlidesTo', S.cur + 2, S.cur); },
  sectionMenu: (a) => menuAt(a, [
    { label: '구역 추가', icon: 'section', action: () => run('addSection') },
    { label: '구역 이름 바꾸기', action: () => run('renameSection') },
    { label: '구역 제거', action: () => run('removeSection') },
    { label: '모든 구역 제거', action: () => change(() => { for (const s of S.pres.slides) delete s.section; }, { scope: 'all' }) },
  ]),
  addSection: () => formDialog('구역 추가', [{ name: 'n', label: '구역 이름', type: 'text', value: '제목 없는 구역' }], (v) => change(() => { slide().section = v.n || '제목 없는 구역'; if (S.cur > 0 && !S.pres.slides.slice(0, S.cur).some((s) => s.section)) S.pres.slides[0].section = '기본 구역'; }, { scope: 'all' })),
  renameSection: (i = null) => {
    let k = i ?? S.cur;
    while (k > 0 && !S.pres.slides[k].section) k--;
    const s = S.pres.slides[k];
    if (!s?.section) { toast('구역이 없습니다'); return; }
    formDialog('구역 이름 바꾸기', [{ name: 'n', label: '구역 이름', type: 'text', value: s.section }], (v) => change(() => { s.section = v.n || s.section; }, { scope: 'all' }));
  },
  removeSection: (i = null) => { let k = i ?? S.cur; while (k > 0 && !S.pres.slides[k].section) k--; change(() => { delete S.pres.slides[k].section; }, { scope: 'all' }); },

  // 글꼴 · 단락
  fontFamily: (f) => applyRun({ font: f }),
  fontSize: (n) => applyRun({ size: n }),
  growFont: () => stepFont(1),
  shrinkFont: () => stepFont(-1),
  bold: () => toggleRun('b'),
  italic: () => toggleRun('i'),
  underline: () => toggleRun('u'),
  strike: () => toggleRun('s'),
  textShadow: () => toggleRun('shadow'),
  superscript: () => applyRun({ base: currentRunProp('base') === 1 ? undefined : 1 }),
  subscript: () => applyRun({ base: currentRunProp('base') === -1 ? undefined : -1 }),
  clearFormat: () => applyRun({ b: undefined, i: undefined, u: undefined, s: undefined, color: undefined, font: undefined, size: undefined, hl: undefined, shadow: undefined, spc: undefined, base: undefined, cap: undefined, outline: undefined }),
  fontColor: (c = lastColor.font) => { lastColor.font = c; applyRun({ color: c }); },
  fontColorMenu: (a) => colorMenu(a, (c) => run('fontColor', c), { auto: '자동', target: 'font' }),
  highlight: (c = lastColor.hl) => { lastColor.hl = c; applyRun({ hl: c }); },
  highlightMenu: (a) => colorMenu(a, (c) => (c ? run('highlight', c) : applyRun({ hl: undefined })), { none: '강조 표시 없음' }),
  charSpacingMenu: (a) => menuAt(a, [[-3, '매우 좁게'], [-1.5, '좁게'], [0, '표준'], [3, '넓게'], [6, '매우 넓게']].map(([v, l]) => ({ label: l, action: () => applyRun({ spc: v || undefined }) })).concat([{ sep: true }, { label: '기타 간격...', action: fontDialog }])),
  changeCaseMenu: (a) => menuAt(a, [
    { label: '문장의 첫 글자를 대문자로', action: () => caseChange((t) => t.replace(/(^\s*|[.!?]\s+)([a-z])/g, (m, p, c) => p + c.toUpperCase())) },
    { label: '소문자로', action: () => caseChange((t) => t.toLowerCase()) },
    { label: '대문자로', action: () => caseChange((t) => t.toUpperCase()) },
    { label: '단어의 첫 글자를 대문자로', action: () => caseChange((t) => t.replace(/\b[a-z]/g, (c) => c.toUpperCase())) },
  ]),
  bullets: (ch) => {
    const p = currentPara();
    const on = p?.bullet?.type === 'char' && !ch;
    applyPara(null, (q) => { q.bullet = on ? { type: 'none' } : { type: 'char', char: ch ?? '•' }; delete q.marL; delete q.indent; });
  },
  bulletsMenu: (a) => menuAt(a, [{ label: '없음', action: () => applyPara(null, (q) => { q.bullet = { type: 'none' }; delete q.marL; delete q.indent; }) }, ...BULLET_CHARS.map(([c, l]) => ({ label: `${c}  ${l}`, action: () => run('bullets', c) }))]),
  numbering: (scheme) => {
    const p = currentPara();
    const on = p?.bullet?.type === 'num' && !scheme;
    applyPara(null, (q) => { q.bullet = on ? { type: 'none' } : { type: 'num', scheme: scheme ?? 'arabicPeriod', start: 1 }; delete q.marL; delete q.indent; });
  },
  numberingMenu: (a) => menuAt(a, [{ label: '없음', action: () => applyPara(null, (q) => { q.bullet = { type: 'none' }; }) }, ...NUM_SCHEMES.map(([k, l]) => ({ label: l, action: () => run('numbering', k) }))]),
  indentMore: () => { if (S.editing) applyTextFormat(null, null, 1); else applyPara(null, (q) => { q.lvl = Math.min(8, (q.lvl ?? 0) + 1); delete q.marL; delete q.indent; }); },
  indentLess: () => { if (S.editing) applyTextFormat(null, null, -1); else applyPara(null, (q) => { q.lvl = Math.max(0, (q.lvl ?? 0) - 1); delete q.marL; delete q.indent; }); },
  lineSpacingMenu: (a) => menuAt(a, [...[1, 1.5, 2, 2.5, 3].map((v) => ({ label: String(v.toFixed(1)), checked: currentPara()?.lineSpacing === v, action: () => applyPara({ lineSpacing: v, lineSpacingPt: undefined }) })), { sep: true }, { label: '줄 간격 옵션...', action: paragraphDialog }]),
  alignLeft: () => applyPara({ align: 'l' }),
  alignCenter: () => applyPara({ align: 'ctr' }),
  alignRight: () => applyPara({ align: 'r' }),
  alignJustify: () => applyPara({ align: 'just' }),
  textDirMenu: (a) => menuAt(a, [['horz', '가로'], ['eaVert', '세로'], ['vert270', '모든 텍스트 270° 회전']].map(([v, l]) => ({ label: l, action: () => textBodyProp((b) => { if (v === 'horz') delete b.vert; else b.vert = v; }) }))),
  textAnchorMenu: (a) => menuAt(a, [['t', '위쪽', 'alignTop'], ['ctr', '중간', 'alignMiddle'], ['b', '아래쪽', 'alignBottom']].map(([v, l, i]) => ({ label: l, icon: i, action: () => textBodyProp((b) => { b.anchor = v; }) }))),
  anchorTop: () => textBodyProp((b) => { b.anchor = 't'; }),
  anchorMiddle: () => textBodyProp((b) => { b.anchor = 'ctr'; }),
  anchorBottom: () => textBodyProp((b) => { b.anchor = 'b'; }),
  fontDialog,
  paragraphDialog,

  // 그리기
  drawShape: (k) => { if (S.editing) endEdit(); S.drawShape = k; emit('drawMode'); toast(`${SHAPE_LABEL[k] ?? '도형'}: 슬라이드에서 끌어서 그리세요`); },
  objAlign: (how) => doAlign(how),
  // PowerPoint 단축키용
  resizeSel: (dw, dh) => {
    const objs = expandGroups(selObjects());
    if (!objs.length) return;
    change(() => { for (const o of objs) { if (o.w + dw >= 1) o.w += dw; if (o.h + dh >= 1) o.h += dh; } }, { key: 'resize-key' });
  },
  cycleCase: () => {
    const modes = [(t) => t.toLowerCase(), (t) => t.toUpperCase(), (t) => t.replace(/(^|\s)([a-z])/g, (m, p, c) => p + c.toUpperCase())];
    S.caseCycle = ((S.caseCycle ?? -1) + 1) % modes.length;
    caseChange(modes[S.caseCycle]);
  },
  moveParagraph: (d) => { if (!moveParagraph(d)) toast('글 편집 중에 사용할 수 있습니다'); },
  nextPlaceholder: () => {
    const s = slide();
    const cur = S.editing ? objById(S.editing.id) : selOne();
    if (S.editing) endEdit();
    const list = s.objects.filter((o) => o.ph && o.text && o.type === 'shape');
    const i = cur ? list.indexOf(cur) : -1;
    const n = list[i + 1];
    if (n) { S.sel = new Set([n.id]); emit('selection'); startEdit(n, { end: true }); return; }
    run('newSlide');
    requestAnimationFrame(() => { const t = slide().objects.find((o) => o.ph && o.text); if (t) { S.sel = new Set([t.id]); emit('selection'); startEdit(t, { end: true }); } });
  },
  repeatLast: () => { const l = S.lastRepeat; if (!l) { toast('반복할 작업이 없습니다'); return; } run(l.name, ...l.args); },
  toggleRibbon: () => { document.getElementById('app')?.classList.toggle('ribbon-collapsed'); requestAnimationFrame(() => fitZoom()); },
  eyedropFill: () => { if (S.editing) endEdit(); startEyedrop('fill'); },
  eyedropLine: () => { if (S.editing) endEdit(); startEyedrop('line'); },
  eyedropFont: () => startEyedrop('font'),
  drawTextbox: () => { if (S.editing) endEdit(); S.drawShape = 'textbox'; emit('drawMode'); toast('텍스트 상자: 슬라이드를 클릭하거나 끌어서 만드세요'); },
  shapesMenu: (a) => shapesGrid(a, (k) => run('drawShape', k)),
  changeShapeMenu: (a) => shapesGrid(a, (k) => { const objs = selObjects().filter((o) => o.type === 'shape' || o.type === 'image'); if (!need(objs)) return; change(() => { for (const o of objs) { o.shape = k; delete o.adj; delete o.path; } }); }),
  arrangeMenu: (a) => menuAt(a, [
    { title: '개체 순서' },
    { label: '맨 앞으로 가져오기', icon: 'bringForward', action: () => doReorder('front') }, { label: '맨 뒤로 보내기', icon: 'sendBackward', action: () => doReorder('back') },
    { label: '앞으로 가져오기', action: () => doReorder('forward') }, { label: '뒤로 보내기', action: () => doReorder('backward') },
    { title: '개체 그룹화' },
    { label: '그룹', icon: 'group', key: 'Ctrl+G', action: doGroup }, { label: '그룹 해제', icon: 'ungroup', key: 'Ctrl+Shift+G', action: doUngroup },
    { title: '개체 위치' },
    { label: '맞춤', icon: 'align', submenu: alignItems() }, { label: '회전', icon: 'rotate', submenu: rotateItems() },
    { label: '선택 창...', icon: 'selectionPane', action: () => run('selectionPane') },
  ]),
  alignMenu: (a) => menuAt(a, alignItems()),
  rotateMenu: (a) => menuAt(a, rotateItems()),
  groupMenu: (a) => menuAt(a, [{ label: '그룹', icon: 'group', key: 'Ctrl+G', action: doGroup }, { label: '그룹 해제', icon: 'ungroup', key: 'Ctrl+Shift+G', action: doUngroup }]),
  bringMenu: (a) => menuAt(a, [{ label: '앞으로 가져오기', action: () => doReorder('forward') }, { label: '맨 앞으로 가져오기', action: () => doReorder('front') }]),
  sendMenu: (a) => menuAt(a, [{ label: '뒤로 보내기', action: () => doReorder('backward') }, { label: '맨 뒤로 보내기', action: () => doReorder('back') }]),
  bringForward: () => doReorder('forward'),
  sendBackward: () => doReorder('backward'),
  bringToFront: () => doReorder('front'),
  sendToBack: () => doReorder('back'),
  group: doGroup,
  ungroup: doUngroup,
  align: (how) => doAlign(how),
  quickStylesMenu: (a) => menuAt(a, [{ node: el('div', { class: 'style-grid7' }, SHAPE_STYLES.map((st, i) => styleChip(st, i))) }], { minWidth: 360 }),
  applyQuickStyle: (i) => {
    const st = SHAPE_STYLES[i];
    const objs = selObjects().filter((o) => o.type === 'shape');
    if (!need(objs, '도형을 선택하세요')) return;
    change(() => { for (const o of objs) { o.fill = JSON.parse(JSON.stringify(st.fill)); o.line = st.line ? { ...st.line, dash: 'solid' } : null; if (o.text) { o.text.defColor = st.text; for (const p of o.text.paras) for (const r of p.runs) delete r.color; } } });
  },
  shapeFill: (c = lastColor.fill) => { lastColor.fill = c; fillSelected({ type: 'solid', color: c }); },
  shapeFillMenu: (a) => colorMenu(a, (c) => { if (c === null) fillSelected(null); else run('shapeFill', c); }, { none: '채우기 없음', extra: [
    { label: '그림...', icon: 'picture', action: () => run('pickFillImage', (f) => fillSelected(f)) },
    { label: '그라데이션', icon: 'fill', submenu: [['밝은 그라데이션', 'light'], ['어두운 그라데이션', 'dark'], ['다른 그라데이션...', 'pane']].map(([l, k]) => ({ label: l, action: () => { if (k === 'pane') { run('formatPane', 'shape'); return; } const base = selObjects()[0]?.fill?.color ?? '@accent1'; fillSelected({ type: 'gradient', angle: 90, stops: k === 'light' ? [[0, `${String(base).split(':')[0]}:lm20:lo80`], [1, base]] : [[0, base], [1, `${String(base).split(':')[0]}:lm50`]] }); } })) },
  ] }),
  shapeOutline: (c = lastColor.line) => { lastColor.line = c; lineSelected((l) => ({ ...(l ?? { width: 1.33, dash: 'solid' }), color: c })); },
  shapeOutlineMenu: (a) => colorMenu(a, (c) => { if (c === null) lineSelected(() => null); else run('shapeOutline', c); }, { none: '윤곽선 없음', target: 'line', extra: [
    { label: '두께', submenu: [0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((pt) => ({ label: `${pt}pt`, action: () => lineSelected((l) => ({ ...(l ?? { color: '@tx1', dash: 'solid' }), width: pt / 0.75 })) })) },
    { label: '대시', submenu: DASH_LABEL.map(([k, l]) => ({ label: l, action: () => lineSelected((l2) => ({ ...(l2 ?? { color: '@tx1', width: 1.33 }), dash: k })) })) },
    { label: '화살표', submenu: [['none', 'none', '화살표 없음'], ['none', 'triangle', '끝 화살표'], ['triangle', 'none', '시작 화살표'], ['triangle', 'triangle', '양쪽 화살표'], ['oval', 'triangle', '원 · 화살표']].map(([h, t, l]) => ({ label: l, action: () => lineSelected((l2) => ({ ...(l2 ?? { color: '@tx1', width: 1.33 }), head: h === 'none' ? undefined : h, tail: t === 'none' ? undefined : t })) })) },
  ] }),
  shapeEffectsMenu: (a) => menuAt(a, [
    { label: '그림자', checked: !!selObjects()[0]?.shadow, action: () => { const objs = selObjects(); if (!need(objs)) return; const v = !objs[0].shadow; change(() => { for (const o of objs) o.shadow = v || undefined; }); } },
    { label: '효과 없음', action: () => change(() => { for (const o of selObjects()) delete o.shadow; }) },
    { sep: true }, { label: '도형 서식...', action: () => run('formatPane', 'shape') },
  ]),
  formatPaneShape: () => run('formatPane', 'shape'),
  formatPane: (k) => { S.formatPane = S.formatPane === k ? null : k; renderPane(); requestAnimationFrame(() => { if (S.fitZoom) fitZoom(); }); },
  setSize: ({ w, h }) => { const objs = selObjects(); if (!need(objs)) return; change(() => { for (const o of objs) { if (w > 0) o.w = w; if (h > 0) o.h = h; } }); },
  rotate: (deg) => change(() => { for (const o of selObjects()) o.rot = (((o.rot ?? 0) + deg) % 360 + 360) % 360; }),
  flipH: () => change(() => { for (const o of selObjects()) o.flipH = !o.flipH || undefined; }),
  flipV: () => change(() => { for (const o of selObjects()) o.flipV = !o.flipV || undefined; }),
  selectMenu: (a) => menuAt(a, [{ label: '모두 선택', key: 'Ctrl+A', action: () => run('selectAll') }, { label: '선택 창...', icon: 'selectionPane', action: () => run('selectionPane') }]),
  selectAll: () => { S.sel = new Set(slide().objects.filter((o) => !o.hidden).map((o) => o.id)); emit('selection'); },
  selectionPane: () => run('formatPane', 'selection'),
  deleteSelection: () => {
    const objs = expandGroups(selObjects());
    if (!objs.length) return;
    const ids = new Set(objs.map((o) => o.id));
    change(() => {
      const s = slide();
      // 개체 틀은 지우면 빈 개체 틀로 (PowerPoint: 내용 있는 개체 틀을 지우면 빈 틀이 남음)
      s.objects = s.objects.flatMap((o) => {
        if (!ids.has(o.id)) return [o];
        if (o.ph && o.type === 'shape' && !isEmptyText(o.text)) { o.text.paras = [{ ...o.text.paras[0], runs: [] }]; return [o]; }
        if (o.ph === 'pic' && o.media) { o.media = null; delete o.crop; return [o]; }
        return [];
      });
      s.anims = (s.anims ?? []).filter((a) => s.objects.some((o) => o.id === a.obj));
    });
    S.sel.clear();
    emit('selection');
  },
  nudge: (dx, dy) => { const objs = expandGroups(selObjects()); if (!objs.length) return; change(() => { for (const o of objs) { o.x += dx; o.y += dy; } }, { key: 'nudge' }); },
  editText: () => { const o = selOne(); if (o) startEdit(o, o.type === 'table' ? { cell: [0, 0] } : { all: true }); },

  // 찾기
  find: () => findDialog(false),
  replace: () => findDialog(true),

  // 삽입
  tableMenu: (a) => tableGrid(a),
  insertTable: (r, c) => insertTable(r, c),
  insertVideo: () => insertMedia('video').catch((e) => alertDialog('비디오', e.message)),
  insertAudio: () => insertMedia('audio').catch((e) => alertDialog('오디오', e.message)),
  mediaOptions: () => mediaOptions(),
  previewMedia: () => {
    const o = selOne();
    if (o?.type !== 'media') { toast('비디오나 오디오를 선택하세요'); return; }
    const src = S.pres.media[o.media];
    const m = el(o.kind === 'audio' ? 'audio' : 'video', { src, controls: true, autoplay: true, playsInline: true, style: { width: '100%', maxHeight: '60vh', background: '#000' } });
    openDialog({ title: o.name ?? (o.kind === 'audio' ? '오디오' : '비디오'), width: 640, body: m, buttons: [{ label: '닫기', primary: true }], onClose: () => m.pause() });
  },
  insertPicture: async (opt = {}) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = 'image/*';
    inp.multiple = !opt.into;
    inp.addEventListener('change', () => insertImages(inp.files, opt.into).catch((e) => alertDialog('그림', e.message)));
    inp.click();
  },
  changePicture: () => {
    const o = selOne();
    if (o?.type !== 'image') { toast('그림을 선택하세요'); return; }
    pickFile('image/*').then(async (f) => { if (!f) return; const r = await readImageFile(f); if (!r) return; const id = addMedia(r.url); change(() => { o.media = id; delete o.crop; delete o.icon; o.h = o.w * (r.h / r.w); }); });
  },
  pickFillImage: (setFill) => pickFile('image/*').then(async (f) => { if (!f) return; const r = await readImageFile(f, 1920); if (r) setFill({ type: 'image', media: addMedia(r.url) }); }),
  screenshot: async () => {
    if (!navigator.mediaDevices?.getDisplayMedia) { toast('이 브라우저는 화면 캡처를 지원하지 않습니다'); return; }
    try {
      const stream = await navigator.mediaDevices.getDisplayMedia({ video: true });
      const video = document.createElement('video');
      video.srcObject = stream;
      await video.play();
      await new Promise((r) => setTimeout(r, 300));
      const cv = document.createElement('canvas');
      cv.width = video.videoWidth; cv.height = video.videoHeight;
      cv.getContext('2d').drawImage(video, 0, 0);
      for (const t of stream.getTracks()) t.stop();
      const r = await shrinkImage(cv.toDataURL('image/png'), 2560);
      if (r) { const added = []; placeImage(r, null, added); S.sel = new Set(added.map((o) => o.id)); emit('selection'); }
    } catch (e) { if (e.name !== 'NotAllowedError') alertDialog('스크린샷', e.message); }
  },
  insertIcons: () => iconsDialog(),
  insertSmartArt: () => smartArtDialog(),
  smartTextPane: () => smartTextPane(),
  smartAddShape: () => { const c = currentSmart(); if (!c) { toast('SmartArt 그래픽을 선택하세요'); return; } rebuildSmart(c, { items: [...c.items, '텍스트'] }); },
  smartLayoutMenu: (a) => { const c = currentSmart(); if (!c) { toast('SmartArt 그래픽을 선택하세요'); return; } menuAt(a, SMART_KINDS.map(([k, l]) => ({ label: l, checked: k === c.kind, action: () => rebuildSmart(c, { kind: k }) }))); },
  smartColors: () => { const c = currentSmart(); if (!c) return; rebuildSmart(c, { multi: !c.multi }); },
  smartToShapes: () => { const c = currentSmart(); if (!c) return; change(() => { for (const m of c.members) { delete m.smart; delete m.smartItems; delete m.smartMulti; } }); toast('도형으로 변환했습니다 (그룹은 유지)'); },
  chartMenu: (a) => menuAt(a, CHART_KINDS.map(([k, l]) => ({ label: l, icon: { col: 'chartColumn', bar: 'chartBar', line: 'chartLine', area: 'chartArea', pie: 'chartPie', doughnut: 'chartDoughnut', scatter: 'chartScatter' }[k], action: () => insertChart(k) }))),
  chartData: () => chartDataDialog(),
  chartKindMenu: (a) => menuAt(a, CHART_KINDS.map(([k, l]) => ({ label: l, action: () => { const c = selOne(); if (c?.type === 'chart') change(() => { c.chart.kind = k; }); } }))),
  chartElementsMenu: (a) => {
    const c = selOne();
    if (c?.type !== 'chart') { toast('차트를 선택하세요'); return; }
    const ch = c.chart;
    menuAt(a, [
      { label: '차트 제목', checked: !!ch.title, action: () => change(() => { ch.title = ch.title ? '' : '차트 제목'; }) },
      { label: '데이터 레이블', checked: !!ch.labels, action: () => change(() => { ch.labels = !ch.labels; }) },
      { label: '범례: 아래쪽', checked: ch.legend === 'b', action: () => change(() => { ch.legend = 'b'; }) },
      { label: '범례: 오른쪽', checked: ch.legend === 'r', action: () => change(() => { ch.legend = 'r'; }) },
      { label: '범례 없음', checked: !ch.legend, action: () => change(() => { ch.legend = null; }) },
      { label: '누적', checked: !!ch.stacked, action: () => change(() => { ch.stacked = !ch.stacked; }) },
      { label: '표식 (꺾은선)', checked: !!ch.markers, action: () => change(() => { ch.markers = !ch.markers; }) },
    ]);
  },
  chartColorsMenu: (a) => {
    const c = selOne();
    if (c?.type !== 'chart') { toast('차트를 선택하세요'); return; }
    menuAt(a, [
      { label: '다양한 색 (테마)', action: () => change(() => { for (const s of c.chart.series) delete s.color; }) },
      ...['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((acc, i) => ({ label: `단색 ${i + 1}`, swatch: resolveColor(S.pres.theme, `@${acc}`), action: () => change(() => { c.chart.series.forEach((s, k) => { s.color = resolveColor(S.pres.theme, `@${acc}:lm${[100, 75, 50, 40, 60, 80][k % 6]}`); }); }) })),
    ]);
  },
  hyperlink: () => hyperlinkDialog(),
  headerFooter: () => headerFooterDialog(),
  insertDate: () => {
    const t = new Date().toLocaleDateString('ko-KR', { year: 'numeric', month: 'long', day: 'numeric' });
    if (!insertTextAtCaret(t)) run('headerFooter');
  },
  insertSlideNumber: () => { if (!insertTextAtCaret(String(S.cur + 1))) change(() => { S.pres.footer = { ...(S.pres.footer ?? {}), slideNum: true }; }, { scope: 'all' }); },
  wordArtMenu: (a) => menuAt(a, [{ node: el('div', { class: 'wa-grid' }, WORDART.map((w, i) => el('button', { class: 'gal-item wa', title: w.label, html: `<span style="${wordArtCss(w)}">가</span>`, onclick: () => { closeMenus(); insertWordArt(i); } }))) }]),
  applyWordArt: (i) => {
    const w = WORDART[i];
    const t = textTargets();
    if (t.mode === 'objects' && !t.bodies.length) { insertWordArt(i); return; }
    applyRun({ ...w.run, outline: w.run.outline, shadow: w.run.shadow || undefined, b: w.run.b || undefined });
  },
  symbol: () => symbolDialog(),

  // 디자인
  applyTheme: (name) => applyTheme(name),
  themeColorsMenu: (a) => menuAt(a, THEMES.map((t) => ({ label: t.label, node: el('button', { class: 'menu-item theme-colors', onclick: () => { closeMenus(); applyTheme(t.name, { colorsOnly: true }); } }, el('span', { class: 'tc-strip' }, ['dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((k) => el('i', { style: { background: t.colors[k] } }))), el('span', {}, t.label)) })), { scroll: true }),
  themeFontsMenu: (a) => menuAt(a, FONT_PAIRS.map(([mj, mn]) => ({ node: el('button', { class: 'menu-item font-pair', onclick: () => { closeMenus(); change(() => { S.pres.theme = { ...S.pres.theme, fonts: { major: mj, minor: mn } }; }, { scope: 'all' }); } }, el('b', { style: { fontFamily: `'${mj}'` } }, mj), el('span', { style: { fontFamily: `'${mn}'` } }, mn)) }))),
  bgStylesMenu: (a) => {
    const t = S.pres.theme;
    const grid = el('div', { class: 'bg-grid' }, BG_STYLES.map((f) => el('button', { class: 'bg-chip', style: { background: (f.type === 'solid' ? resolveColor(t, f.color) : f.path ? `radial-gradient(circle, ${f.stops.map(([p, c]) => `${resolveColor(t, c)} ${p * 100}%`).join(',')})` : `linear-gradient(${(f.angle ?? 90) + 90}deg, ${f.stops.map(([p, c]) => `${resolveColor(t, c)} ${p * 100}%`).join(',')})`) }, onclick: () => { closeMenus(); change(() => { for (const s of S.pres.slides) s.bg = JSON.parse(JSON.stringify(f)); }, { scope: 'all' }); } })));
    menuAt(a, [{ node: grid }, { sep: true }, { label: '배경 서식...', icon: 'formatBg', action: () => run('formatBg') }, { label: '배경 원래대로', action: () => change(() => { for (const s of S.pres.slides) s.bg = null; }, { scope: 'all' }) }]);
  },
  slideSizeMenu: (a) => menuAt(a, [
    { label: '표준 (4:3)', checked: Math.abs(S.pres.size.w / S.pres.size.h - 4 / 3) < 0.01, action: () => { change(() => resizePresentation(S.pres, 960, 720, 'fit'), { scope: 'all' }); fitZoom(); } },
    { label: '와이드스크린 (16:9)', checked: Math.abs(S.pres.size.w / S.pres.size.h - 16 / 9) < 0.01, action: () => { change(() => resizePresentation(S.pres, 1280, 720, 'fit'), { scope: 'all' }); fitZoom(); } },
    { sep: true }, { label: '사용자 지정 슬라이드 크기...', action: slideSizeDialog },
  ]),
  formatBg: () => run('formatPane', 'bg'),
  designIdeas: () => designIdeasDialog(),

  // 전환
  setTransition: (type) => {
    for (const s of pickedSlides()) {
      change(() => { s.transition = type === 'none' ? (s.transition?.advAfter != null ? { type: 'none', advAfter: s.transition.advAfter } : null) : { dur: 1, ...(s.transition ?? {}), type, ...(type === 'cut' ? { dur: 0 } : {}) }; }, { scope: 'all' });
    }
    if (type !== 'none') previewTransition();
  },
  transitionOptionsMenu: (a) => menuAt(a, [['b', '아래에서'], ['t', '위에서'], ['l', '왼쪽에서'], ['r', '오른쪽에서']].map(([d, l]) => ({ label: l, checked: (slide().transition?.dir ?? 'b') === d, disabled: !['push', 'wipe', 'cover', 'uncover'].includes(slide().transition?.type), action: () => { change(() => { slide().transition = { ...slide().transition, dir: d }; }); previewTransition(); } }))),
  transitionDuration: (v) => change(() => { const s = slide(); s.transition = { type: 'fade', ...(s.transition ?? {}), dur: Math.max(0, Number(v) || 0) }; }),
  transitionAdvance: (o) => change(() => { const s = slide(); s.transition = { type: 'none', ...(s.transition ?? {}), ...('advClick' in o ? { advClick: o.advClick ? undefined : false } : {}), ...('advAfter' in o ? { advAfter: o.advAfter ?? undefined } : {}) }; if (s.transition.advAfter == null) delete s.transition.advAfter; if (s.transition.advClick == null) delete s.transition.advClick; }),
  applyTransitionAll: () => { const t = slide().transition; change(() => { for (const s of S.pres.slides) s.transition = t ? JSON.parse(JSON.stringify(t)) : null; }, { scope: 'all' }); toast('모든 슬라이드에 적용했습니다'); },
  previewTransition: () => previewTransition(),

  // 애니메이션
  setAnim: (cls, effect) => setAnim(cls, effect),
  addAnimMenu: (a) => menuAt(a ?? { x: innerWidth / 2, y: 160 }, Object.entries(ANIM_EFFECTS).map(([cls, list]) => ({ label: ANIM_CLASS_LABEL[cls], submenu: list.map(([k, l]) => ({ label: l, action: () => { const objs = expandGroups(selObjects()); if (!need(objs, '개체를 선택하세요')) return; change(() => { for (const o of objs) addAnim(slide(), o.id, cls, k); }); previewAnims(objs.map((o) => o.id)); } })) }))),
  animOptionsMenu: (a) => {
    const cur = (slide().anims ?? []).find((x) => x.id === S.animSel || S.sel.has(x.obj));
    menuAt(a, [
      { title: '방향' },
      ...[['b', '아래에서'], ['t', '위에서'], ['l', '왼쪽에서'], ['r', '오른쪽에서']].map(([d, l]) => ({ label: l, checked: cur?.dir === d, action: () => run('animTiming', { dir: d }) })),
      { title: '시퀀스' },
      { label: '하나의 개체로', checked: !!cur && !cur.byPara, action: () => run('animTiming', { byPara: undefined }) },
      { label: '단락별', checked: !!cur?.byPara, action: () => run('animTiming', { byPara: true }) },
      ...(cur?.cls === 'path' ? [{ title: '경로' }, { label: '경로 반대로', action: () => run('animTiming', { motion: reversePath(cur.motion) }) }] : []),
    ]);
  },
  // 트리거: 고른 개체의 애니메이션을 다른 개체를 누를 때 재생 (PowerPoint [고급 애니메이션 › 트리거])
  triggerMenu: (a) => {
    const s = slide();
    const ids = new Set(expandGroups(selObjects()).map((o) => o.id));
    const mine = (s.anims ?? []).filter((x) => x.id === S.animSel || ids.has(x.obj));
    if (!mine.length) { toast('애니메이션이 있는 개체를 선택하세요'); return; }
    const cur = mine[0].trigger ?? null;
    const setT = (t) => change(() => { for (const x of mine) { if (t) x.trigger = t; else delete x.trigger; } });
    menuAt(a ?? { x: innerWidth / 2, y: 160 }, [
      { label: '이전 클릭 순서대로 (트리거 없음)', checked: !cur, action: () => setT(null) },
      { label: '클릭할 때', submenu: s.objects.filter((o) => !ids.has(o.id)).map((o) => ({ label: objLabel(o, s.objects.indexOf(o)), checked: cur === o.id, action: () => setT(o.id) })) },
    ]);
  },
  animTiming: (props) => { const objs = expandGroups(selObjects()); const ids = new Set(objs.map((o) => o.id)); change(() => { for (const a of slide().anims ?? []) if (ids.has(a.obj) || a.id === S.animSel) Object.assign(a, props); }); },
  removeAnim: (opt = {}) => {
    const s = slide();
    const ids = new Set(expandGroups(selObjects()).map((o) => o.id));
    change(() => { s.anims = (s.anims ?? []).filter((a) => (opt.all || !S.animSel ? !ids.has(a.obj) : a.id !== S.animSel)); });
    S.animSel = null;
  },
  moveAnim: (dir) => {
    const s = slide();
    const list = s.anims ?? [];
    const i = list.findIndex((a) => a.id === S.animSel || (!S.animSel && S.sel.has(a.obj)));
    const j = i + dir;
    if (i < 0 || j < 0 || j >= list.length) return;
    change(() => { [list[i], list[j]] = [list[j], list[i]]; });
  },
  animPane: () => run('formatPane', 'anim'),
  previewAnim: () => previewAnims(),
  animPainter: () => {
    const o = selOne();
    const a = o ? slide().anims?.filter((x) => x.obj === o.id) : null;
    if (!a?.length) { toast('애니메이션이 있는 개체를 선택하세요'); return; }
    S.animClip = a.map((x) => ({ ...x }));
    S.painter = { anim: true };
    emit('painter');
    toast('애니메이션을 적용할 개체를 클릭하세요');
  },

  // 슬라이드 쇼
  showFromStart: () => { if (S.editing) endEdit(); startShow({ from: 0 }); },
  showFromCurrent: () => { if (S.editing) endEdit(); startShow({ from: S.cur }); },
  presenterView: () => { if (S.editing) endEdit(); startShow({ from: S.cur, presenter: true }); },
  rehearse: () => { if (S.editing) endEdit(); startShow({ from: 0, rehearse: true, presenter: true }); },
  setupShow: () => setupShowDialog(),
  customShows: () => customShowsDialog(),
  customShowMenu: (a) => menuAt(a, [...(S.pres.customShows ?? []).map((c) => ({ label: c.name, action: () => startShow({ custom: c.name }) })), ...(S.pres.customShows?.length ? [{ sep: true }] : []), { label: '쇼 재구성...', action: () => customShowsDialog() }]),
  exportVideo: () => exportVideoDialog(),
  newComment: () => {
    if (S.editing) endEdit();
    const ta = el('textarea', { rows: 4, style: { width: '100%' }, placeholder: '메모를 입력하세요 (@이름 으로 언급)' });
    openDialog({ title: '새 메모', width: 420, body: ta, onOpen: () => ta.focus(), buttons: [{ label: '게시', primary: true, action: () => { const t = ta.value.trim(); if (!t) return false; addComment(t); S.formatPane = 'comments'; renderPane(); emit('change', { scope: 'slide' }); return true; } }, { label: '취소' }] });
  },
  commentsPane: () => { S.formatPane = S.formatPane === 'comments' ? null : 'comments'; renderPane(); fitZoom(); emit('change', { scope: 'view' }); },
  deleteComments: (scope = 'slide') => change(() => { if (scope === 'all') for (const s2 of S.pres.slides) delete s2.comments; else delete slide().comments; }),
  prevComment: () => jumpComment(-1),
  nextComment: () => jumpComment(1),

  // 검토
  spellCheck: () => spellDialog(),
  accessibility: () => accessibilityDialog(),
  wordCount: () => statsDialog(),

  // 보기
  // 슬라이드 마스터 보기
  viewMaster: () => { if (S.editing) endEdit(); S.masterKey = 'master'; S.view = 'normal'; S.sel = new Set(); refresh('all'); refresh('view'); setRibbonTab('slideMaster'); },
  closeMaster: () => { if (S.editing) endEdit(); S.masterKey = null; S.sel = new Set(); refresh('all'); refresh('view'); setRibbonTab('view'); },
  masterPick: (key) => { if (S.editing) endEdit(); S.masterKey = key; S.sel = new Set(); refresh('all'); },
  masterRename: () => { const m = curSlide(); if (!S.masterKey || !m) return; formDialog('레이아웃 이름 바꾸기', [{ name: 'n', label: '이름', value: m.name ?? (LAYOUTS.find(([k]) => k === m.layout)?.[1] ?? '슬라이드 마스터') }], (v) => change(() => { m.name = v.n || undefined; }, { scope: 'all' })); },
  masterReset: () => { if (!S.masterKey) return; change(() => { delete S.pres.masters[S.masterKey]; }, { scope: 'all' }); S.sel = new Set(); },
  masterHideGraphics: () => { const m = curSlide(); if (!S.masterKey || S.masterKey === 'master') { toast('레이아웃을 고르세요 (마스터 아래 목록)'); return; } change(() => { m.hideMaster = !m.hideMaster; }, { scope: 'all' }); },
  viewNormal: () => { S.view = 'normal'; if (S.masterKey) { S.masterKey = null; S.sel = new Set(); refresh('all'); } refresh('view'); },
  viewSorter: () => { if (S.editing) endEdit(); S.view = 'sorter'; refresh('view'); },
  viewReading: () => { if (S.editing) endEdit(); startShow({ from: S.cur, windowed: true }); },
  viewOutline: () => { if (S.editing) endEdit(); S.view = 'outline'; refresh('view'); },
  viewNotesPage: () => { S.showNotes = true; S.bigNotes = !S.bigNotes; refresh('view'); },
  toggleNotes: () => { S.showNotes = !S.showNotes; S.bigNotes = false; refresh('view'); },
  toggleGrid: () => { S.showGrid = !S.showGrid; renderCanvas(); emit('selection'); },
  toggleGuides: () => { S.showGuides = !S.showGuides; renderCanvas(); emit('selection'); },
  addGuide: (dir = 'v') => {
    const { w, h } = S.pres.size;
    const list = [...(S.pres.guides ?? [{ dir: 'v', pos: w / 2 }, { dir: 'h', pos: h / 2 }])];
    list.push({ dir, pos: (dir === 'v' ? w : h) / 2 + 24 * list.filter((g) => g.dir === dir).length });
    S.showGuides = true;
    change(() => { S.pres.guides = list; }, { scope: 'all' });
  },
  toggleRuler: () => { S.showRuler = !S.showRuler; refresh('view'); },
  zoomDialog: () => zoomDialog(),
  fitZoom: () => fitZoom(),
  zoomIn: () => setZoom(S.zoom * 1.1),
  zoomOut: () => setZoom(S.zoom / 1.1),

  // 도움말
  shortcuts: () => openDialog({ title: '바로 가기 키', width: 600, body: el('table', { class: 'kbd-table' }, SHORTCUTS.map(([k, d]) => el('tr', {}, el('td', {}, k), el('td', {}, d)))), buttons: [{ label: '닫기', primary: true }] }),
  whatsNew: () => openDialog({ title: `WIPOINT ${APP_VERSION} 새로운 기능`, width: 560, body: el('ul', { class: 'whats-new' }, WHATS_NEW.map((t) => el('li', {}, t))), buttons: [{ label: '확인', primary: true }] }),
  about: () => openDialog({ title: 'WIPOINT 정보', width: 460, body: el('div', {}, el('p', {}, el('b', {}, `WIPOINT (위포인트) ${APP_VERSION}`), ' — PowerPoint 와 같은 웹 프레젠테이션'), el('p', { class: 'muted' }, '의존성 없는 순수 JavaScript · 문서는 내 컴퓨터(브라우저)에만 저장됩니다. WIXEL(위셀)과 같은 방식으로 만들었습니다.'), el('p', { class: 'muted' }, `슬라이드 ${S.pres.slides.length}장 · ${Math.round(S.pres.size.w * 2.54 / 96 * 10) / 10}×${Math.round(S.pres.size.h * 2.54 / 96 * 10) / 10}cm`)), buttons: [{ label: '확인', primary: true }] }),

  // 그림 서식
  removeBackground: () => removeBackground(),
  insertEquation: () => equationDialog(),
  editEquation: (id) => { const o = (id && objById(id)) || selObjects().find((x) => x.type === 'equation'); if (o) equationDialog(o); else equationDialog(); },
  insertZoom: (kind = 'slide') => insertZoom(kind),
  zoomMenu: (a) => menuAt(a ?? { x: innerWidth / 2, y: 160 }, [
    { label: '요약 확대/축소', desc: '고른 슬라이드의 축소판으로 목차 슬라이드를 만듭니다', action: () => run('insertZoom', 'summary') },
    { label: '구역 확대/축소', desc: '구역으로 가는 축소판 — 구역을 다 본 뒤 돌아옵니다', action: () => run('insertZoom', 'section') },
    { label: '슬라이드 확대/축소', desc: '한 슬라이드로 가는 축소판', action: () => run('insertZoom', 'slide') },
  ]),
  zoomReturnToggle: () => change(() => { for (const o of selObjects()) if (o.type === 'zoom') o.zoom.ret = !o.zoom.ret; }),
  resetPicture: () => change(() => { for (const o of selObjects()) if (o.type === 'image') { if (o.bgOrig) { o.media = o.bgOrig; delete o.bgOrig; } for (const k of ['crop', 'bright', 'contrast', 'alpha', 'gray', 'shadow', 'line', 'shape', 'adj']) delete o[k]; } }),
  applyPictureStyle: (i) => { const st = PICTURE_STYLES[i].set; const objs = selObjects().filter((o) => o.type === 'image'); if (!need(objs, '그림을 선택하세요')) return; change(() => { for (const o of objs) for (const [k, v] of Object.entries(st)) { if (v === undefined) delete o[k]; else o[k] = JSON.parse(JSON.stringify(v)); } }); },
  pictureCorrectionsMenu: (a) => menuAt(a, [[-0.4, '밝기 -40%'], [-0.2, '밝기 -20%'], [0, '밝기 표준'], [0.2, '밝기 +20%'], [0.4, '밝기 +40%']].map(([v, l]) => ({ label: l, action: () => change(() => { for (const o of selObjects()) if (o.type === 'image') o.bright = v || undefined; }) })).concat([{ sep: true }, { label: '대비 +20%', action: () => change(() => { for (const o of selObjects()) if (o.type === 'image') o.contrast = 0.2; }) }, { label: '그림 수정 옵션...', action: () => run('formatPane', 'shape') }])),
  pictureColorMenu: (a) => menuAt(a, [
    { label: '회색조', action: () => change(() => { for (const o of selObjects()) if (o.type === 'image') o.gray = true; }) },
    { label: '원래 색', action: () => change(() => { for (const o of selObjects()) if (o.type === 'image') delete o.gray; }) },
    ...[0.25, 0.5, 0.75].map((v) => ({ label: `투명도 ${v * 100}%`, action: () => change(() => { for (const o of selObjects()) if (o.type === 'image') o.alpha = 1 - v; }) })),
    { label: '그래픽 채우기 (아이콘 색)...', action: () => { const icons = selObjects().filter((o) => o.icon); if (!need(icons, '아이콘을 선택하세요')) return; colorMenu(a, (c) => recolorIcons(icons, resolveColor(S.pres.theme, c ?? '#000000'))); } },
  ]),
  cropPicture: () => { const o = selOne(); if (o?.type === 'image' && o.media) { if (S.cropping) endCrop(); else { startCrop(o); toast('검은 손잡이를 끌어 자르고, 그림을 끌어 위치를 맞춘 뒤 Esc 또는 바깥을 누르세요'); } } else cropDialog(); },
  cropMenu: (a) => menuAt(a, [{ label: '자르기', icon: 'crop', action: () => run('cropPicture') }, { label: '자르기 (수치 입력)...', action: cropDialog }, { label: '자르기 원래대로', action: () => change(() => { for (const o of selObjects()) if (o.type === 'image' && o.crop) { const c = o.crop; const fw = o.w / (1 - c.l - c.r); const fh = o.h / (1 - c.t - c.b); o.x -= c.l * fw; o.y -= c.t * fh; o.w = fw; o.h = fh; delete o.crop; } }) }, { label: '도형에 맞춰 자르기', submenu: ['rect', 'roundRect', 'ellipse', 'triangle', 'diamond', 'hexagon', 'star5', 'heart', 'cloud'].map((k) => ({ label: SHAPE_LABEL[k] ?? k, action: () => change(() => { for (const o of selObjects()) if (o.type === 'image') { o.shape = k === 'rect' ? undefined : k; } }) })) }, { label: '가로 세로 비율', submenu: [['1:1', 1], ['4:3', 4 / 3], ['16:9', 16 / 9], ['3:4', 3 / 4]].map(([l, r]) => ({ label: l, action: () => cropAspect(r) })) }]),
  altText: () => {
    const o = selOne();
    if (!o) { toast('개체를 선택하세요'); return; }
    formDialog('대체 텍스트', [{ name: 'alt', label: '설명', type: 'textarea', value: o.alt ?? '' }], (v) => change(() => { o.alt = v.alt.trim() || undefined; }));
  },

  // 표
  tblOpt: (k) => tableOp((t) => { t.style = { ...(t.style ?? {}), none: undefined, [k]: !t.style?.[k] }; }),
  applyTableStyle: (i) => tableOp((t) => { t.style = { ...(t.style ?? { firstRow: true, banded: true }), ...tableStyleProps(TABLE_STYLES[i]) }; for (const r of t.rows) for (const c of r.cells) delete c.fill; }),
  cellFill: (c = lastColor.cell) => { lastColor.cell = c; const t0 = curTable(); const g = t0 ? cellRange(t0) : null; tableOp((t) => { for (const [, , cell] of cellsIn(t, g)) cell.fill = c ? { type: 'solid', color: c } : null; }); },
  cellFillMenu: (a) => colorMenu(a, (c) => run('cellFill', c), { none: '채우기 없음' }),
  cellBorderMenu: (a) => menuAt(a, [
    { label: '아래쪽 테두리', icon: 'borderBottom', action: () => tableBorders('bottom') },
    { label: '위쪽 테두리', icon: 'borderTop', action: () => tableBorders('top') },
    { label: '왼쪽 테두리', icon: 'borderLeft', action: () => tableBorders('left') },
    { label: '오른쪽 테두리', icon: 'borderRight', action: () => tableBorders('right') },
    { sep: true },
    { label: '테두리 없음', icon: 'borderNone', action: () => tableBorders('none') },
    { label: '모든 테두리', icon: 'borderAll', action: () => tableBorders('all') },
    { label: '바깥쪽 테두리', icon: 'borderOutside', action: () => tableBorders('outside') },
    { label: '안쪽 테두리', action: () => tableBorders('inside') },
    { label: '안쪽 가로 테두리', action: () => tableBorders('insideH') },
    { label: '안쪽 세로 테두리', action: () => tableBorders('insideV') },
    { label: '하향 대각선 테두리', action: () => tableBorders('diagDown') },
    { label: '상향 대각선 테두리', action: () => tableBorders('diagUp') },
    { sep: true },
    { label: '펜 색...', icon: 'shapeOutline', action: () => colorMenu(a, (c) => { S.pen = { ...(S.pen ?? {}), color: c ?? '@tx1' }; toast('펜 색을 바꿨습니다. 테두리를 고르면 적용됩니다'); }) },
    { label: '펜 두께', submenu: [0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({ label: `${w}pt`, checked: (S.pen?.width ?? 1) === w * 96 / 72, action: () => { S.pen = { ...(S.pen ?? {}), width: w * 96 / 72 }; } })) },
    { label: '표 스타일 테두리로', action: () => tableBorders('style') },
  ]),
  tblSelectMenu: (a) => menuAt(a, [{ label: '표 선택', action: () => { const t = curTable(); if (t) { endEdit(); S.sel = new Set([t.id]); emit('selection'); } } }]),
  tblDeleteMenu: (a) => menuAt(a, [{ label: '열 삭제', action: () => run('tblDeleteCol') }, { label: '행 삭제', action: () => run('tblDeleteRow') }, { label: '표 삭제', action: () => { const t = curTable(); if (t) { endEdit(); S.sel = new Set([t.id]); run('deleteSelection'); } } }]),
  tblRowAbove: () => tableOp((t) => { const [r] = curCell(); t.rows.splice(r, 0, { h: t.rows[r].h, cells: t.rows[r].cells.map(blankCell) }); }),
  tblRowBelow: () => tableOp((t) => { const [r] = curCell(); t.rows.splice(r + 1, 0, { h: t.rows[r].h, cells: t.rows[r].cells.map(blankCell) }); }),
  tblColLeft: () => tableOp((t) => { const [, c] = curCell(); const w = t.cols[c]; t.cols.splice(c, 0, w); for (const r of t.rows) r.cells.splice(c, 0, blankCell(r.cells[c])); }),
  tblColRight: () => tableOp((t) => { const [, c] = curCell(); const w = t.cols[c]; t.cols.splice(c + 1, 0, w); for (const r of t.rows) r.cells.splice(c + 1, 0, blankCell(r.cells[c])); }),
  tblDeleteRow: () => tableOp((t) => { if (t.rows.length < 2) return; const [r] = curCell(); t.rows.splice(r, 1); }),
  tblDeleteCol: () => tableOp((t) => { if (t.cols.length < 2) return; const [, c] = curCell(); t.cols.splice(c, 1); for (const r of t.rows) r.cells.splice(c, 1); }),
  tblMerge: () => {
    const t = curTable();
    if (!t) { toast('표를 선택하세요'); return; }
    const g = cellRange(t);
    if (g.all || (g.r1 === g.r2 && g.c1 === g.c2)) { toast('병합할 칸들을 끌어서(또는 Shift+클릭으로) 고르세요'); return; }
    tableOp((tt) => { mergeRange(tt, g); });
    S.cellRange = null;
  },
  tblSplit: () => {
    const t = curTable();
    if (!t) { toast('표를 선택하세요'); return; }
    const [r, c] = curCell();
    const cell = t.rows[r].cells[c];
    if (cell.span || cell.rowSpan) { tableOp((tt) => unmergeAt(tt, r, c)); return; }
    formDialog('셀 분할', [{ name: 'c', label: '열 개수', type: 'number', value: '2' }, { name: 'r', label: '행 개수', type: 'number', value: '1' }], (v) => {
      const nc = Math.max(1, Math.min(20, Math.round(Number(v.c) || 1)));
      const nr = Math.max(1, Math.min(20, Math.round(Number(v.r) || 1)));
      if (nc === 1 && nr === 1) return;
      tableOp((tt) => splitCell(tt, r, c, nc, nr));
      S.cellRange = null;
    });
  },
  tblBorders: (where) => tableBorders(where),
  tblEqualRows: () => tableOp((t) => { const h = t.rows.reduce((s, r) => s + r.h, 0) / t.rows.length; for (const r of t.rows) r.h = h; }),
  tblEqualCols: () => tableOp((t) => { const w = t.cols.reduce((s, c) => s + c, 0) / t.cols.length; t.cols = t.cols.map(() => w); }),

  // 상태 (리본 켜짐 표시)
  ribbonState: () => ribbonState(),
});

function caseChange(fn) {
  const t = textTargets();
  if (t.mode === 'edit') { endEdit(); }
  const bodies = textTargets().bodies ?? [];
  if (!need(bodies, '텍스트를 선택하세요')) return;
  change(() => { for (const b of bodies) for (const p of b.paras) for (const r of p.runs) r.t = fn(r.t); });
}
function textBodyProp(fn) {
  const objs = S.editing ? [objById(S.editing.id)] : selObjects();
  const bodies = [];
  for (const o of objs) {
    if (o?.text) bodies.push(o.text);
    if (o?.type === 'table') { if (S.editing?.cell) bodies.push(o.rows[S.editing.cell[0]].cells[S.editing.cell[1]].text); else for (const r of o.rows) for (const c of r.cells) bodies.push(c.text); }
  }
  if (!need(bodies, '텍스트 상자를 선택하세요')) return;
  const wasEditing = S.editing;
  if (wasEditing) endEdit();
  change(() => { for (const b of bodies) fn(b); });
}
function fillSelected(f) {
  const objs = selObjects().filter((o) => o.type === 'shape' || o.icon);
  if (!need(objs, '도형을 선택하세요')) return;
  const icons = objs.filter((o) => o.icon);
  if (icons.length && f?.type === 'solid') recolorIcons(icons, resolveColor(S.pres.theme, f.color));
  change(() => { for (const o of objs) if (o.type === 'shape') o.fill = f ? JSON.parse(JSON.stringify(f)) : null; });
}
function lineSelected(fn) {
  const objs = selObjects().filter((o) => o.type === 'shape' || o.type === 'image');
  if (!need(objs)) return;
  change(() => { for (const o of objs) { o.line = fn(o.line); if (o.line) o.line.dash ??= 'solid'; } });
}
function recolorIcons(icons, color) {
  change(() => {
    for (const o of icons) {
      const svg = iconSvg({ vb: o.icon.vb, body: o.icon.body }, color);
      const u = new TextEncoder().encode(svg);
      let s = '';
      for (let i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000));
      o.media = addMedia(`data:image/svg+xml;base64,${btoa(s)}`);
      o.icon = { ...o.icon, fill: color };
    }
  });
}
function doReorder(how) {
  const ids = expandGroups(selObjects()).map((o) => o.id);
  if (!ids.length) return;
  change(() => reorder(slide(), ids, how));
}
function alignItems() {
  return [
    { label: '왼쪽 맞춤', icon: 'alignLeft', action: () => doAlign('l') }, { label: '가운데 맞춤', icon: 'alignCenter', action: () => doAlign('c') }, { label: '오른쪽 맞춤', icon: 'alignRight', action: () => doAlign('r') },
    { label: '위쪽 맞춤', icon: 'alignTop', action: () => doAlign('t') }, { label: '중간 맞춤', icon: 'alignMiddle', action: () => doAlign('m') }, { label: '아래쪽 맞춤', icon: 'alignBottom', action: () => doAlign('b') },
    { sep: true }, { label: '가로 간격을 동일하게', action: () => doAlign('dh') }, { label: '세로 간격을 동일하게', action: () => doAlign('dv') },
    { sep: true }, { label: '슬라이드에 맞춤', checked: !!S.alignToSlide, action: () => { S.alignToSlide = true; } }, { label: '선택한 개체 맞춤', checked: !S.alignToSlide, action: () => { S.alignToSlide = false; } },
  ];
}
function rotateItems() {
  return [
    { label: '오른쪽으로 90도 회전', icon: 'rotate', action: () => run('rotate', 90) }, { label: '왼쪽으로 90도 회전', action: () => run('rotate', -90) },
    { label: '상하 대칭', icon: 'flipV', action: () => run('flipV') }, { label: '좌우 대칭', icon: 'flipH', action: () => run('flipH') },
    { sep: true }, { label: '기타 회전 옵션...', action: () => run('formatPane', 'shape') },
  ];
}
function cropAspect(r) {
  const o = selOne();
  if (o?.type !== 'image') return;
  const ar = o.w / o.h;
  change(() => {
    if (ar > r) { const nw = o.h * r; const cut = (o.w - nw) / o.w; o.crop = { l: cut / 2, r: cut / 2, t: 0, b: 0 }; o.x += (o.w - nw) / 2; o.w = nw; } else { const nh = o.w / r; const cut = (o.h - nh) / o.h; o.crop = { t: cut / 2, b: cut / 2, l: 0, r: 0 }; o.y += (o.h - nh) / 2; o.h = nh; }
  });
}
function insertWordArt(i) {
  const w = WORDART[i];
  const { w: W, h: H } = S.pres.size;
  const tb = newTextBox({ x: W * 0.2, y: H * 0.38, w: W * 0.6, h: 110 }, '여기에 텍스트를 입력하십시오');
  tb.text.paras[0].align = 'ctr';
  Object.assign(tb.text.paras[0].runs[0], { size: 54, ...JSON.parse(JSON.stringify(w.run)) });
  change(() => slide().objects.push(tb));
  S.sel = new Set([tb.id]);
  emit('selection');
  startEdit(tb, { all: true });
}

/** 서식 복사: 채우기 · 선 · 효과 · 첫 글자 서식 · 단락 맞춤 */
function captureStyle(o) {
  const r0 = o.text?.paras?.find((p) => p.runs.length)?.runs[0] ?? o.text?.paras?.[0]?.end ?? {};
  const { t: _t, ...run0 } = r0;
  return { fill: o.fill, line: o.line, shadow: o.shadow, run: run0, align: o.text?.paras?.[0]?.align, defColor: o.text?.defColor, type: o.type, img: o.type === 'image' ? { bright: o.bright, contrast: o.contrast, gray: o.gray, alpha: o.alpha, line: o.line, shape: o.shape } : null };
}
function applyStyle(objs, st) {
  if (!objs.length) return;
  change(() => {
    for (const o of objs) {
      if (o.type === 'shape') {
        if (st.type === 'shape') { o.fill = st.fill ? JSON.parse(JSON.stringify(st.fill)) : null; o.line = st.line ? JSON.parse(JSON.stringify(st.line)) : null; }
        o.shadow = st.shadow || undefined;
        if (o.text) { applyRunProps(o.text, Object.fromEntries(['b', 'i', 'u', 's', 'size', 'color', 'font', 'hl', 'shadow', 'spc', 'outline'].map((k) => [k, st.run?.[k]]))); if (st.align) applyParaProps(o.text, { align: st.align }); if (st.defColor) o.text.defColor = st.defColor; }
      } else if (o.type === 'image' && st.img) Object.assign(o, JSON.parse(JSON.stringify(st.img)));
    }
  });
}

function ribbonState() {
  const objs = selObjects();
  const ctx = [];
  if (S.editing || objs.some((o) => o.type === 'shape')) ctx.push('shape');
  if (objs.some((o) => o.type === 'image' && o.media)) ctx.push('picture');
  if (objs.some((o) => o.type === 'table') || (S.editing && objById(S.editing.id)?.type === 'table')) ctx.push('table');
  if (objs.some((o) => o.type === 'chart')) ctx.push('chart');
  if (objs.some((o) => o.type === 'media')) ctx.push('media');
  if (objs.some((o) => o.smart)) ctx.push('smart');
  if (S.masterKey) ctx.unshift('master');
  const p = currentPara();
  const tbl = curTable();
  const st = tbl?.style ?? {};
  const color = currentRunProp('color');
  return {
    context: ctx,
    font: currentRunProp('font') ?? (objs.length || S.editing ? null : '+mn'),
    size: (() => { const v = currentRunProp('size'); return v == null ? null : Math.round(v * 10) / 10; })(),
    b: !!currentRunProp('b'), i: !!currentRunProp('i'), u: !!currentRunProp('u'), s: !!currentRunProp('s'), shadow: !!currentRunProp('shadow'),
    bullets: p?.bullet?.type === 'char', numbering: p?.bullet?.type === 'num',
    al: p?.align === 'l' || (!p?.align && !!p), ac: p?.align === 'ctr', ar: p?.align === 'r', aj: p?.align === 'just',
    painter: !!S.painter, cmPane: S.formatPane === 'comments', animPane: S.formatPane === 'anim', selPane: S.formatPane === 'selection', hidden: !!slide()?.hidden,
    vMaster: !!S.masterKey, masterHide: !!(S.masterKey && curSlide()?.hideMaster), vNormal: S.view === 'normal' && !S.masterKey, vOutline: S.view === 'outline', vSorter: S.view === 'sorter', notes: S.showNotes, grid: S.showGrid, guides: S.showGuides, ruler: S.showRuler,
    tFirstRow: !!st.firstRow, tLastRow: !!st.lastRow, tBanded: !!st.banded, tFirstCol: !!st.firstCol, tLastCol: !!st.lastCol,
    bars: { font: resolveColor(S.pres.theme, lastColor.font), hl: resolveColor(S.pres.theme, lastColor.hl), fill: resolveColor(S.pres.theme, lastColor.fill), line: resolveColor(S.pres.theme, lastColor.line), cell: resolveColor(S.pres.theme, lastColor.cell) },
    curColor: color,
  };
}

