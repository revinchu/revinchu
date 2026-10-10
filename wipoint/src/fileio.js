// 파일: 열기 · 저장 (.pptx / .wpt.json) · 자동 저장(IndexedDB) · 내보내기(PDF 인쇄 · PNG) · 공유 링크
import { S, emit, goSlide } from './state.js';
import { readPptx, pptxEntries } from './pptx.js';
import { readPdf } from './pdf.js';
import { decodeEmbeddedFont } from './fonts.js';
import { zipAsync } from './zip.js';
import { validatePresentation, pruneMedia, snapshot, newPresentation } from './model.js';
import { svgToObjects } from './svgimport.js';
import { slideSvg } from './svgexport.js';
import { slideHtml, SLIDE_CSS } from './render.js';
import { toast, alertDialog } from './ui.js';

const PPTX_MIME = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';

// ───────────── 열기 ─────────────
export function pickFile(accept) {
  return new Promise((resolve) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = accept;
    inp.addEventListener('change', () => resolve(inp.files?.[0] ?? null));
    inp.click();
  });
}

/** 파일 바이트 → 문서 (pptx · WIPOINT JSON) */
export async function loadFile(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  // 아이폰 · 카카오톡 · iCloud 에서 아직 내려받지 않은 파일은 0바이트로 넘어옴
  if (!bytes.length) throw new Error(`'${file.name}' 파일이 비어 있습니다 (0바이트). 아이폰이라면 파일 앱이나 카카오톡에서 파일을 먼저 열어 내려받은 뒤(또는 [파일에 저장] 후) 다시 열어 주세요.`);
  const name = file.name.replace(/\.(pptx|potx|pptm|ppsx|json|wpt|pdf)$/i, '').replace(/\.wpt$/i, '');
  if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46 || /\.pdf$/i.test(file.name)) {
    const { pres, warnings, jobs } = readPdf(bytes, { encodeImage: canvasEncode });
    await Promise.all(jobs.map(async (j) => { const u = await maskJpeg(j); if (u) pres.media[j.media] = u; }));
    return { pres, name, warnings: ['PDF 를 슬라이드로 바꾸어 열었습니다. 글은 텍스트 상자, 도형 · 선은 도형으로 편집할 수 있습니다.', ...warnings], pdf: true };
  }
  // SVG: 편집 가능한 도형으로 바꾼 새 프레젠테이션 (슬라이드 크기는 SVG 비율)
  if (/\.svgz?$/i.test(file.name) || /^\s*(<\?xml[^>]*>\s*)?(<!--[\s\S]*?-->\s*)*(<!DOCTYPE[^>]*>\s*)?<svg[\s>]/i.test(new TextDecoder().decode(bytes.subarray(0, 2048)))) {
    const text = new TextDecoder().decode(bytes);
    const probe = svgToObjects(text);
    const ratio = probe.width / Math.max(1, probe.height);
    const pres = newPresentation({ firstLayout: 'blank' });
    pres.size = ratio >= 1 ? { w: 1280, h: Math.round(1280 / ratio) } : { w: Math.round(720 * ratio), h: 720 };
    const { objects, media } = svgToObjects(text, { size: pres.size });
    Object.assign(pres.media, media);
    pres.slides[0].objects = objects;
    pres.slides[0].hideDecor = true;
    return { pres, name: file.name.replace(/\.svgz?$/i, ''), warnings: [`SVG 를 도형 ${objects.length}개로 바꾸어 열었습니다. 도형 · 글을 바로 편집할 수 있습니다.`], pdf: true };
  }
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf) throw new Error('예전 PowerPoint 97-2003 형식(.ppt)은 열 수 없습니다. PowerPoint 에서 .pptx 로 저장한 뒤 열어 주세요.');
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) {
    const { pres, warnings } = readPptx(bytes);
    return { pres, name, warnings };
  }
  const text = new TextDecoder().decode(bytes);
  if (!/^\s*\{/.test(text)) throw new Error(`'${file.name}' 은(는) PowerPoint 파일(.pptx)이 아닙니다. 파일이 손상되었거나 다른 형식일 수 있습니다.`);
  let data;
  try { data = JSON.parse(text); } catch { throw new Error(`'${file.name}' 을(를) 읽을 수 없습니다 (내려받기가 덜 끝났거나 손상된 파일).`); }
  const pres = validatePresentation(data);
  return { pres, name, warnings: [] };
}

/** PDF 그림 RGBA → data URL (불투명하면 JPEG, 투명이 있으면 PNG) */
function canvasEncode(w, h, rgba, hasAlpha) {
  try {
    const c = document.createElement('canvas');
    c.width = w;
    c.height = h;
    c.getContext('2d').putImageData(new ImageData(rgba, w, h), 0, 0);
    return hasAlpha ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.92);
  } catch { return null; }
}

/** JPEG + 투명 가리개(SMask) → PNG */
async function maskJpeg({ url, alpha, w, h }) {
  try {
    const img = new Image();
    img.src = url;
    await img.decode();
    const c = document.createElement('canvas');
    c.width = w;
    c.height = h;
    const g = c.getContext('2d');
    g.drawImage(img, 0, 0, w, h);
    const d = g.getImageData(0, 0, w, h);
    for (let i = 0; i < w * h; i++) d.data[i * 4 + 3] = alpha[i];
    g.putImageData(d, 0, 0);
    return c.toDataURL('image/png');
  } catch { return null; }
}

// ───────────── 포함된 글꼴 ─────────────
const fontFaces = [];
/** pres.fonts (pptx 에 포함된 .fntdata) → 브라우저 글꼴로 등록 (같은 이름의 컴퓨터 글꼴보다 먼저 쓰임) */
export async function applyEmbeddedFonts(pres) {
  if (typeof FontFace !== 'function' || !document.fonts) return 0;
  for (const f of fontFaces.splice(0)) { try { document.fonts.delete(f); } catch { /* 무시 */ } }
  let n = 0;
  for (const f of pres.fonts ?? []) {
    try {
      const m = /base64,(.*)$/s.exec(f.data ?? '');
      if (!m) continue;
      const bin = atob(m[1]);
      const bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
      const { data } = decodeEmbeddedFont(bytes);
      const face = new FontFace(f.typeface, data, { weight: /bold/i.test(f.style) ? '700' : '400', style: /italic/i.test(f.style) ? 'italic' : 'normal' });
      await face.load();
      document.fonts.add(face);
      fontFaces.push(face);
      n++;
    } catch (e) { console.warn('포함된 글꼴', f.typeface, e); }
  }
  if (n) emit('change', { scope: 'all' });
  return n;
}

export function setDocument(pres, name, { handle = null } = {}) {
  S.pres = pres;
  S.docName = name || '프레젠테이션1';
  S.fileHandle = handle;
  S.history.clear();
  S.sel.clear();
  S.editing = null;
  S.masterKey = null;
  S.dirty = false;
  goSlide(0);
  emit('change', { scope: 'all' });
  emit('docLoaded');
  if (pres.fonts?.length || fontFaces.length) applyEmbeddedFonts(pres).then((n) => { if (n) toast(`파일에 포함된 글꼴 ${n}개를 불러왔습니다`); });
}

export async function openWithPicker() {
  if (window.showOpenFilePicker) {
    try {
      const [h] = await window.showOpenFilePicker({ id: 'wipoint-open', types: [{ description: '프레젠테이션', accept: { [PPTX_MIME]: ['.pptx', '.ppsx', '.potx'], 'application/pdf': ['.pdf'], 'image/svg+xml': ['.svg'], 'application/json': ['.json'] } }] });
      const f = await h.getFile();
      const r = await loadFile(f);
      setDocument(r.pres, r.name, { handle: /\.pptx$/i.test(f.name) ? h : null });
      report(r.warnings, r.pdf);
      return;
    } catch (e) {
      if (e.name === 'AbortError') return;
      if (e.name !== 'SecurityError' && e.name !== 'TypeError') throw e;
    }
  }
  const f = await pickFile('.pptx,.ppsx,.potx,.pdf,.svg,.json,.ppt,application/pdf,image/svg+xml,application/vnd.openxmlformats-officedocument.presentationml.presentation');
  if (!f) return;
  const r = await loadFile(f);
  setDocument(r.pres, r.name);
  report(r.warnings, r.pdf);
}
function report(warnings, pdf) {
  if (pdf) { toast(warnings[0]); warnings = warnings.slice(1); }
  if (warnings?.length) alertDialog('일부 내용', warnings.join('\n'));
  else toast('프레젠테이션을 열었습니다');
}

// ───────────── 저장 ─────────────
/** SVG 그림 → PNG (PowerPoint 2016 이전 호환용 대체 그림) */
function svgToPng(url, w = 512) {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      const h = Math.max(1, Math.round((w * (img.naturalHeight || 1)) / (img.naturalWidth || 1)));
      const cv = document.createElement('canvas');
      cv.width = w; cv.height = h;
      cv.getContext('2d').drawImage(img, 0, 0, w, h);
      resolve(cv.toDataURL('image/png'));
    };
    img.onerror = () => resolve(null);
    img.src = url;
  });
}

export async function pptxBlob() {
  pruneMedia(S.pres);
  const png = {};
  for (const [id, url] of Object.entries(S.pres.media)) if (url.startsWith('data:image/svg')) { const p = await svgToPng(url); if (p) png[id] = p; }
  const entries = pptxEntries(S.pres, { png });
  const bytes = await zipAsync(entries);
  return new Blob([bytes], { type: PPTX_MIME });
}
export const jsonBlob = () => new Blob([JSON.stringify(S.pres)], { type: 'application/json' });

function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}

/** 저장 위치 고르기 (가능한 브라우저) → 핸들, 아니면 내려받기 */
async function pickSave(name, kind) {
  if (!window.showSaveFilePicker) return null;
  const types = { pptx: [{ description: 'PowerPoint 프레젠테이션', accept: { [PPTX_MIME]: ['.pptx'] } }], json: [{ description: 'WIPOINT 문서', accept: { 'application/json': ['.json'] } }], pdf: [{ description: 'PDF', accept: { 'application/pdf': ['.pdf'] } }], png: [{ description: 'PNG 그림', accept: { 'image/png': ['.png'] } }], jpg: [{ description: 'JPEG 그림', accept: { 'image/jpeg': ['.jpg'] } }], svg: [{ description: 'SVG 그림', accept: { 'image/svg+xml': ['.svg'] } }] }[kind];
  try { return await window.showSaveFilePicker({ id: 'wipoint-save', suggestedName: name, types }); } catch (e) { if (e.name === 'AbortError') return false; return null; }
}
async function writeHandle(h, blob) {
  const w = await h.createWritable();
  await w.write(blob);
  await w.close();
}

export async function save() {
  if (S.fileHandle) {
    await writeHandle(S.fileHandle, await pptxBlob());
    S.dirty = false;
    emit('saved');
    toast('저장했습니다');
    return;
  }
  await saveAs('pptx');
}

export async function saveAs(kind = 'pptx') {
  const ext = kind === 'json' ? 'wpt.json' : kind;
  const name = `${S.docName}.${ext}`;
  // 사용자 동작(클릭) 직후에 위치를 먼저 고름
  const h = await pickSave(name, kind);
  if (h === false) return;
  const blob = kind === 'json' ? jsonBlob() : await pptxBlob();
  if (h) {
    await writeHandle(h, blob);
    if (kind === 'pptx') { S.fileHandle = h; S.docName = h.name.replace(/\.pptx$/i, ''); }
  } else download(blob, name);
  S.dirty = false;
  emit('saved');
  toast(kind === 'json' ? 'WIPOINT 문서(.json)로 저장했습니다' : 'PowerPoint 파일(.pptx)로 저장했습니다');
}

// ───────────── 자동 저장 (IndexedDB) ─────────────
const DB = 'wipoint';
function idb() {
  return new Promise((resolve, reject) => {
    const r = indexedDB.open(DB, 1);
    r.onupgradeneeded = () => r.result.createObjectStore('kv');
    r.onsuccess = () => resolve(r.result);
    r.onerror = () => reject(r.error);
  });
}
export async function idbSet(key, value) {
  const db = await idb();
  await new Promise((res, rej) => { const tx = db.transaction('kv', 'readwrite'); tx.objectStore('kv').put(value, key); tx.oncomplete = res; tx.onerror = () => rej(tx.error); });
  db.close();
}
export async function idbGet(key) {
  const db = await idb();
  const v = await new Promise((res, rej) => { const r = db.transaction('kv').objectStore('kv').get(key); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); });
  db.close();
  return v;
}

let autoTimer = 0;
export function scheduleAutosave() {
  if (S.viewOnly) return;
  clearTimeout(autoTimer);
  autoTimer = setTimeout(autosaveNow, 1200);
}
export async function autosaveNow() {
  try {
    await idbSet('autosave', { name: S.docName, json: snapshot(S.pres), media: S.pres.media, fonts: S.pres.fonts ?? null, at: Date.now(), cur: S.cur });
    await addRecent();
    emit('autosaved');
  } catch (e) { console.warn('자동 저장 실패', e); }
}
export async function restoreAutosave() {
  try {
    const v = await idbGet('autosave');
    if (!v?.json) return false;
    const pres = validatePresentation({ ...JSON.parse(v.json), media: v.media ?? {}, ...(v.fonts ? { fonts: v.fonts } : {}) });
    setDocument(pres, v.name);
    if (v.cur) goSlide(Math.min(v.cur, pres.slides.length - 1));
    return true;
  } catch { return false; }
}

/** 최근 문서 (첫 슬라이드 미리 보기 포함, 최대 12개) */
async function addRecent() {
  const list = (await idbGet('recent')) ?? [];
  const id = S.docId ?? (S.docId = `d${Date.now().toString(36)}`);
  const item = { id, name: S.docName, at: Date.now(), slides: S.pres.slides.length };
  const out = [item, ...list.filter((x) => x.id !== id)].slice(0, 12);
  await idbSet('recent', out);
  await idbSet(`doc:${id}`, { name: S.docName, json: snapshot(S.pres), media: S.pres.media, fonts: S.pres.fonts ?? null });
  for (const x of list.slice(12)) await idbSet(`doc:${x.id}`, null);
}
export async function recentDocs() { return (await idbGet('recent')) ?? []; }
export async function openRecent(id) {
  const v = await idbGet(`doc:${id}`);
  if (!v) { toast('문서를 찾을 수 없습니다'); return; }
  setDocument(validatePresentation({ ...JSON.parse(v.json), media: v.media ?? {}, ...(v.fonts ? { fonts: v.fonts } : {}) }), v.name);
  S.docId = id;
}

// ───────────── 인쇄 · PDF ─────────────
/** mode: slides | handout2 | handout3 | handout6 | notes ; range: [from, to] (1부터) */
export function printSlides({ mode = 'slides', range = null, hidden = false } = {}) {
  const area = document.getElementById('printArea');
  const { w, h } = S.pres.size;
  let idx = S.pres.slides.map((s, i) => i).filter((i) => hidden || !S.pres.slides[i].hidden);
  if (range) idx = idx.filter((i) => i + 1 >= range[0] && i + 1 <= range[1]);
  const pageW = mode === 'slides' ? w : 794;
  const scale = mode === 'slides' ? 1 : mode === 'notes' ? 640 / w : mode === 'handout2' ? 560 / w : mode === 'handout3' ? 300 / w : 330 / w;
  const tile = (i) => `<div class="pr-slide" style="width:${w * scale}px;height:${h * scale}px"><div style="transform:scale(${scale});transform-origin:0 0;width:${w}px;height:${h}px">${slideHtml(S.pres, S.pres.slides[i], { index: i })}</div></div>`;
  const pages = [];
  if (mode === 'slides') for (const i of idx) pages.push(`<div class="pr-page slides">${tile(i)}</div>`);
  else if (mode === 'notes') for (const i of idx) pages.push(`<div class="pr-page notes">${tile(i)}<div class="pr-notes">${(S.pres.slides[i].notes ?? '').replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' }[c])).replace(/\n/g, '<br>')}</div><div class="pr-no">${i + 1}</div></div>`);
  else {
    const per = { handout2: 2, handout3: 3, handout6: 6 }[mode];
    for (let k = 0; k < idx.length; k += per) {
      const chunk = idx.slice(k, k + per);
      pages.push(`<div class="pr-page handout h${per}">${chunk.map((i) => `<div class="pr-cell">${tile(i)}${per === 3 ? '<div class="pr-lines"></div>' : ''}</div>`).join('')}</div>`);
    }
  }
  const landscape = mode === 'slides';
  area.innerHTML = `<style>${SLIDE_CSS}@page{size:${landscape ? `${w}px ${h}px` : 'A4 portrait'};margin:${landscape ? 0 : '12mm'}}</style>${pages.join('')}`;
  area.dataset.mode = mode;
  area.style.setProperty('--pw', `${pageW}px`);
  setTimeout(() => { window.print(); }, 50);
}

// ───────────── PNG 내보내기 ─────────────
/** 슬라이드 → PNG data URL (SVG foreignObject → canvas) */
export function slidePng(i, width = 1920) {
  const { w, h } = S.pres.size;
  const sc = width / w;
  const html = slideHtml(S.pres, S.pres.slides[i], { index: i });
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${w * sc}" height="${h * sc}"><foreignObject width="100%" height="100%"><div xmlns="http://www.w3.org/1999/xhtml" style="transform:scale(${sc});transform-origin:0 0;width:${w}px;height:${h}px"><style>${SLIDE_CSS}</style>${html.replace(/<br>/g, '<br/>').replace(/<img([^>]*?)>/g, '<img$1/>').replace(/<col(?!group)([^>]*?)>/g, '<col$1/>').replace(/<(video|audio)\b[^>]*>(<\/\1>)?/g, '')}</div></foreignObject></svg>`;
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const cv = document.createElement('canvas');
      cv.width = Math.round(w * sc);
      cv.height = Math.round(h * sc);
      const ctx = cv.getContext('2d');
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, cv.width, cv.height);
      ctx.drawImage(img, 0, 0);
      try { resolve(cv.toDataURL('image/png')); } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error('슬라이드를 그림으로 바꾸지 못했습니다'));
    img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  });
}
// ───────────── SVG · JPG · 그림으로 저장 ─────────────
let measureCtx = null;
/** 브라우저 글자 너비 (SVG 줄 바꿈 계산용) */
export function measureText(font, text) {
  measureCtx ??= document.createElement('canvas').getContext('2d');
  measureCtx.font = font;
  return measureCtx.measureText(text).width;
}
/** 슬라이드 i (또는 고른 개체 only) → SVG 글 */
export const slideSvgText = (i, only = null) => slideSvg(S.pres, S.pres.slides[i], { measure: measureText, only });
/** SVG 글 → 그림 data URL (png · jpeg), 배경 (jpeg 는 흰색) */
export function svgToRaster(svg, { type = 'image/png', scale = 2, quality = 0.92 } = {}) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const cv = document.createElement('canvas');
      cv.width = Math.max(1, Math.round(img.naturalWidth * scale));
      cv.height = Math.max(1, Math.round(img.naturalHeight * scale));
      const g = cv.getContext('2d');
      if (type === 'image/jpeg') { g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); }
      g.drawImage(img, 0, 0, cv.width, cv.height);
      try { resolve(cv.toDataURL(type, quality)); } catch (e) { reject(e); }
    };
    img.onerror = () => reject(new Error('그림으로 바꾸지 못했습니다'));
    img.src = URL.createObjectURL(new Blob([svg], { type: 'image/svg+xml' }));
  });
}
/** 내보내기: kind = svg | png | jpg, all = 모든 슬라이드 */
export async function exportImages(kind = 'png', all = false) {
  const list = all ? S.pres.slides.map((_, i) => i) : [S.cur];
  for (const i of list) {
    const name = `${S.docName}${list.length > 1 ? `_${i + 1}` : ''}.${kind}`;
    let blob;
    if (kind === 'svg') blob = new Blob([slideSvgText(i)], { type: 'image/svg+xml' });
    else {
      // PNG 는 화면과 같은 HTML 렌더(글꼴 · 효과 정확), JPG 는 흰 배경
      const url = kind === 'png' ? await slidePng(i) : await slidePng(i).then((u) => jpegFrom(u));
      blob = await (await fetch(url)).blob();
    }
    if (list.length === 1) {
      const h = await pickSave(name, kind);
      if (h === false) return;
      if (h) { await writeHandle(h, blob); toast('저장했습니다'); return; }
    }
    download(blob, name);
  }
}
async function jpegFrom(pngUrl) {
  const img = new Image();
  img.src = pngUrl;
  await img.decode();
  const cv = document.createElement('canvas');
  cv.width = img.naturalWidth; cv.height = img.naturalHeight;
  const g = cv.getContext('2d');
  g.fillStyle = '#fff'; g.fillRect(0, 0, cv.width, cv.height); g.drawImage(img, 0, 0);
  return cv.toDataURL('image/jpeg', 0.92);
}
/** 고른 개체를 그림 파일로 (PowerPoint [그림으로 저장]) */
export async function saveObjectsAsPicture(ids, kind = 'png') {
  const svg = slideSvgText(S.cur, new Set(ids));
  const name = `${S.docName}_그림.${kind}`;
  const blob = kind === 'svg' ? new Blob([svg], { type: 'image/svg+xml' }) : await (await fetch(await svgToRaster(svg, { type: kind === 'jpg' ? 'image/jpeg' : 'image/png' }))).blob();
  const h = await pickSave(name, kind);
  if (h === false) return;
  if (h) { await writeHandle(h, blob); toast('저장했습니다'); return; }
  download(blob, name);
}

export async function exportPng(all = false) {
  const list = all ? S.pres.slides.map((_, i) => i) : [S.cur];
  for (const i of list) {
    const url = await slidePng(i);
    const blob = await (await fetch(url)).blob();
    download(blob, `${S.docName}${list.length > 1 ? `_${i + 1}` : ''}.png`);
  }
}

// ───────────── 동영상으로 내보내기 (MediaRecorder · WebM/MP4) ─────────────
/** opts: { sec: 슬라이드당 초, width, fade: 전환 초, hidden: 숨긴 슬라이드 포함 }, onProgress(0..1, 글) — 실제 시간만큼 걸림 */
export async function exportVideo(opts = {}, onProgress = () => {}, signal = null) {
  if (typeof MediaRecorder !== 'function' || !HTMLCanvasElement.prototype.captureStream) throw new Error('이 브라우저는 동영상 만들기를 지원하지 않습니다 (Chrome · Edge · 최신 Safari 에서 해 주세요)');
  const width = opts.width ?? 1280;
  const { w, h } = S.pres.size;
  const W = Math.round(width / 2) * 2;
  const H = Math.round((width * h) / w / 2) * 2;
  const idx = S.pres.slides.map((s, i) => i).filter((i) => opts.hidden || !S.pres.slides[i].hidden);
  if (!idx.length) throw new Error('내보낼 슬라이드가 없습니다');
  // 그림 미리 만들기
  const frames = [];
  for (const [k, i] of idx.entries()) {
    onProgress(k / idx.length * 0.2, `슬라이드 그림 만드는 중 (${k + 1}/${idx.length})`);
    const url = await slidePng(i, W);
    const img = new Image();
    img.src = url;
    await img.decode();
    const t = S.pres.slides[i].transition;
    frames.push({ img, sec: t?.advAfter != null && opts.useTimings !== false ? Math.max(0.5, t.advAfter) : opts.sec ?? 5, fade: t && t.type !== 'none' && t.type !== 'cut' ? Math.min(1.5, t.dur ?? 0.7) : opts.fade ?? 0 });
  }
  const cv = document.createElement('canvas');
  cv.width = W;
  cv.height = H;
  const g = cv.getContext('2d');
  const types = ['video/mp4;codecs=avc1', 'video/mp4', 'video/webm;codecs=vp9', 'video/webm;codecs=vp8', 'video/webm'];
  const mime = types.find((t) => MediaRecorder.isTypeSupported?.(t)) ?? '';
  const stream = cv.captureStream(30);
  const rec = new MediaRecorder(stream, { mimeType: mime || undefined, videoBitsPerSecond: width >= 1920 ? 8e6 : 5e6 });
  const chunks = [];
  rec.ondataavailable = (e) => { if (e.data.size) chunks.push(e.data); };
  const stopped = new Promise((res) => { rec.onstop = res; });
  const draw = (img, a = 1) => { g.globalAlpha = a; g.drawImage(img, 0, 0, W, H); g.globalAlpha = 1; };
  g.fillStyle = '#000';
  g.fillRect(0, 0, W, H);
  draw(frames[0].img);
  rec.start(500);
  const total = frames.reduce((n, f) => n + f.sec, 0);
  let elapsed = 0;
  const wait = (ms) => new Promise((r) => setTimeout(r, ms));
  for (let k = 0; k < frames.length; k++) {
    if (signal?.aborted) break;
    const f = frames[k];
    if (k > 0 && f.fade > 0) {
      const t0 = performance.now();
      for (;;) {
        const u = Math.min(1, (performance.now() - t0) / (f.fade * 1000));
        draw(frames[k - 1].img);
        draw(f.img, u);
        if (u >= 1) break;
        await new Promise((r) => requestAnimationFrame(r));
      }
    } else draw(f.img);
    // 한 장을 머무는 동안에도 몇 번씩 다시 그려야 녹화기가 프레임을 받음
    const end = performance.now() + Math.max(0, f.sec - (k > 0 ? f.fade : 0)) * 1000;
    while (performance.now() < end) {
      if (signal?.aborted) break;
      draw(f.img);
      onProgress(0.2 + 0.8 * Math.min(1, (elapsed + f.sec - (end - performance.now()) / 1000) / total), `녹화 중 ${k + 1}/${frames.length}`);
      await wait(200);
    }
    elapsed += f.sec;
  }
  rec.stop();
  await stopped;
  if (signal?.aborted) return null;
  const blob = new Blob(chunks, { type: (mime || 'video/webm').split(';')[0] });
  const ext = blob.type.includes('mp4') ? 'mp4' : 'webm';
  download(blob, `${S.docName}.${ext}`);
  return { size: blob.size, ext };
}

// ───────────── 읽기 전용 공유 링크 (#view=) ─────────────
async function gzip(bytes) {
  const s = new Blob([bytes]).stream().pipeThrough(new CompressionStream('gzip'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}
async function gunzip(bytes) {
  const s = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}
const b64url = (u8) => { let s = ''; for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000)); return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, ''); };
const unb64url = (t) => { const s = atob(t.replace(/-/g, '+').replace(/_/g, '/')); const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u; };

export async function shareLink() {
  const json = new TextEncoder().encode(JSON.stringify({ ...S.pres, props: { ...S.pres.props, name: S.docName } }));
  const z = await gzip(json);
  const url = `${location.origin}${location.pathname}#view=${b64url(z)}`;
  return { url, size: url.length };
}
export async function loadViewLink(hash) {
  const m = /#view=([A-Za-z0-9_-]+)/.exec(hash);
  if (!m) return false;
  const bytes = await gunzip(unb64url(m[1]));
  const pres = validatePresentation(JSON.parse(new TextDecoder().decode(bytes)));
  setDocument(pres, pres.props?.name ?? '공유된 프레젠테이션');
  S.viewOnly = true;
  return true;
}
