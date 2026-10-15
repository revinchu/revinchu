// 파일: 열기 (docx · hwp · hwpx · doc · pdf · md · html · txt · rtf · wiword) · 저장 (docx · hwpx · md · html · txt · pdf(인쇄) · wiword) · 자동 저장 · 최근 문서 · 인쇄 · 클립보드 · 그림
import { S, emit, on, run, register, setDoc, change, curSel } from './state.js';
import { newDoc, uid, deserialize, serialize, docText } from './model.js';
import { MIME, OPEN_ACCEPT, readBytes, docBytes, decodeText, csvToDoc, detectKind } from './formats.js';
import { mdToDoc } from './md.js';
import { pageNode } from './layout.js';
import { exportPdfBytes } from './pdfexport.js';
import { V, relayout } from './view.js';
import { focusEditor, pasteFrom, copyTo, deleteSelection, E, pasteBlocks } from './editor.js';
import { importFontFiles, askLocalFontsOnce } from './fontlib.js';
import { toast, alertDialog, progressOverlay, formDialog, el, openDialog } from './ui.js';


export function pickFile(accept, multiple = false) {
  return new Promise((resolve) => {
    const inp = document.createElement('input');
    inp.type = 'file';
    inp.accept = accept;
    inp.multiple = multiple;
    inp.addEventListener('change', () => resolve(multiple ? [...(inp.files ?? [])] : inp.files?.[0] ?? null));
    inp.click();
  });
}
export const fmtSize = (n) => (n >= 1 << 30 ? `${(n / (1 << 30)).toFixed(1)}GB` : n >= 1 << 20 ? `${Math.round(n / (1 << 20))}MB` : `${Math.max(1, Math.round(n / 1024))}KB`);

// ───────────── 브라우저: 열기 ─────────────
function canvasEncode(w, h, rgba, hasAlpha) {
  try {
    const c = document.createElement('canvas');
    c.width = w; c.height = h;
    c.getContext('2d').putImageData(new ImageData(rgba, w, h), 0, 0);
    return hasAlpha ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.9);
  } catch { return null; }
}
async function maskJpeg({ url, alpha, w, h }) {
  try {
    const img = new Image(); img.src = url; await img.decode();
    const c = document.createElement('canvas'); c.width = w; c.height = h;
    const g = c.getContext('2d'); g.drawImage(img, 0, 0, w, h);
    const d = g.getImageData(0, 0, w, h);
    for (let i = 0; i < w * h; i++) d.data[i * 4 + 3] = alpha[i];
    g.putImageData(d, 0, 0);
    return c.toDataURL('image/png');
  } catch { return null; }
}
const KIND_LABEL = { docx: 'Word 문서', hwpx: '한글 문서(.hwpx)', hwp: '한글 문서(.hwp)', hwp3: '한글 97 문서(.hwp)', doc: 'Word 97-2003 문서', pdf: 'PDF', md: '마크다운', html: '웹 페이지', txt: '텍스트', rtf: 'RTF', odt: 'OpenDocument 텍스트', csv: 'CSV', wiword: 'WIWORD 문서' };
/** 열기 암호 묻기 → 문자열 | null(취소) */
export function askPassword(name, wrong = false) {
  return new Promise((resolve) => {
    let done = false;
    const pw = el('input', { type: 'password', autocomplete: 'off', style: { width: '100%' } });
    openDialog({
      title: '암호', width: 380,
      body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
        el('div', {}, `'${name}' 은(는) 암호로 보호되어 있습니다.`),
        wrong ? el('div', { style: { color: '#c42b1c' } }, '암호가 맞지 않습니다. 다시 입력하세요.') : null,
        el('label', {}, el('span', {}, '암호: '), pw)),
      onOpen: () => setTimeout(() => pw.focus(), 30),
      buttons: [{ label: '확인', primary: true, action: () => { done = true; resolve(pw.value); } }, { label: '취소' }],
      onClose: () => { if (!done) resolve(null); },
    });
    pw.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); pw.closest('.dialog')?.querySelector('.dialog-foot button.primary, button.primary')?.click(); } });
  });
}
export async function loadFile(file, password) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const bar = bytes.length > 4 << 20 ? progressOverlay(`'${file.name}' 여는 중 (${fmtSize(bytes.length)})`) : null;
  try {
    bar?.set(0.2, '읽는 중…');
    await new Promise((r) => setTimeout(r, 0));
    let r; let pw = password; let wrong = false;
    for (;;) {
      try { r = readBytes(bytes, file.name, { encodeImage: canvasEncode, password: pw }); break; } catch (e) {
        if (e.code !== 'PASSWORD') throw e;
        bar?.close();
        pw = await askPassword(file.name, wrong || !!e.wrong);
        if (pw == null) throw new Error('암호를 입력하지 않아 열지 않았습니다.');
        wrong = true;
        await new Promise((res) => setTimeout(res, 30)); // 대화상자가 닫히고 그려지도록
      }
    }
    if (r.jobs?.length) await Promise.all(r.jobs.map(async (j) => { const u = await maskJpeg(j); if (u) r.doc.media[j.media] = u; }));
    await rasterizeMetafiles(r.doc);
    return { ...r, name: file.name.replace(/\.[^.]+$/, '') };
  } finally { bar?.close(); }
}
/** EMF/WMF 그림 → 화면용 SVG 는 읽을 때 만듦 (docx/hwp 리더가 media 에 data:image/svg 로 넣음). 여기서는 TIFF 등 브라우저가 못 그리는 것만 */
async function rasterizeMetafiles() { /* 리더에서 처리 */ }

const setDocument = (doc, name, { handle = null, kind = null } = {}) => {
  setDoc(doc, name);
  S.fileHandle = handle;
  S.fileKind = kind;
  S.docId = null;
  emit('docLoaded');
};
export async function openFile(f, handle = null) {
  try {
    const r = await loadFile(f);
    setDocument(r.doc, r.name, { handle: handle && ['docx', 'hwpx', 'md', 'txt', 'html'].includes(r.kind) ? handle : null, kind: r.kind });
    if (r.warnings?.length) alertDialog(`${KIND_LABEL[r.kind] ?? ''} 열기`, r.warnings.join('\n'));
    else toast(`${KIND_LABEL[r.kind] ?? '문서'}를 열었습니다`);
    return r;
  } catch (e) {
    alertDialog('열 수 없음', e.message ?? String(e));
    return null;
  }
}
export async function openWithPicker() {
  if (window.showOpenFilePicker) {
    try {
      const [h] = await window.showOpenFilePicker({ id: 'wiword-open', types: [{ description: '문서', accept: { [MIME.docx]: ['.docx', '.docm', '.dotx'], 'application/msword': ['.doc'], 'application/x-hwp': ['.hwp', '.hwpx', '.hwt'], 'application/pdf': ['.pdf'], 'text/markdown': ['.md', '.markdown'], 'text/plain': ['.txt', '.csv', '.tsv'], 'text/html': ['.html', '.htm'], 'application/rtf': ['.rtf'], 'application/vnd.oasis.opendocument.text': ['.odt'], 'application/json': ['.wiword', '.json'] } }] });
      await openFile(await h.getFile(), h);
      return;
    } catch (e) {
      if (e.name === 'AbortError') return;
      if (e.name !== 'SecurityError' && e.name !== 'TypeError') throw e;
    }
  }
  const f = await pickFile(OPEN_ACCEPT);
  if (f) await openFile(f);
}

// ───────────── 저장 ─────────────
function download(blob, name) {
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 2000);
}
const SAVE_TYPES = {
  docx: { description: 'Word 문서', accept: { [MIME.docx]: ['.docx'] } },
  dotx: { description: 'Word 서식 파일', accept: { [MIME.dotx]: ['.dotx'] } },
  hwpx: { description: '한글 문서 (HWPX)', accept: { [MIME.hwpx]: ['.hwpx'] } },
  md: { description: '마크다운', accept: { 'text/markdown': ['.md'] } },
  html: { description: '웹 페이지', accept: { 'text/html': ['.html'] } },
  txt: { description: '텍스트', accept: { 'text/plain': ['.txt'] } },
  wiword: { description: 'WIWORD 문서', accept: { 'application/json': ['.wiword'] } },
  pdf: { description: 'PDF', accept: { 'application/pdf': ['.pdf'] } },
};
async function pickSave(name, kind) {
  if (!window.showSaveFilePicker) return null;
  try { return await window.showSaveFilePicker({ id: 'wiword-save', suggestedName: name, types: [SAVE_TYPES[kind]] }); } catch (e) { if (e.name === 'AbortError') return false; return null; }
}
async function writeHandle(h, blob) { const w = await h.createWritable(); await w.write(blob); await w.close(); }
/** 바이트를 파일로 저장 (위치 고르기 → 안 되면 내려받기). 취소하면 false */
export async function saveBytes(bytes, name, { description = 'PDF', mime = 'application/pdf', ext = '.pdf' } = {}) {
  let h = null;
  if (window.showSaveFilePicker) {
    try { h = await window.showSaveFilePicker({ id: 'wiword-tools', suggestedName: name, types: [{ description, accept: { [mime]: [ext] } }] }); } catch (e) { if (e.name === 'AbortError') return false; h = null; }
  }
  const blob = new Blob([bytes], { type: mime });
  if (h) await writeHandle(h, blob); else download(blob, name);
  return true;
}
export async function docBlob(kind) {
  let out = docBytes(S.doc, kind, { name: S.docName });
  if (out instanceof Promise) out = await out;
  const type = MIME[kind] ?? 'application/octet-stream';
  return new Blob([kind === 'md' || kind === 'txt' || kind === 'html' ? `${kind === 'txt' ? '\uFEFF' : ''}${out}` : out], { type: kind === 'md' || kind === 'txt' || kind === 'html' ? `${type};charset=utf-8` : type });
}
export async function save() {
  if (S.fileHandle && S.fileKind) {
    await writeHandle(S.fileHandle, await docBlob(S.fileKind));
    S.dirty = false;
    emit('saved');
    toast('저장했습니다');
    return;
  }
  await saveAs(S.opts.saveFormat ?? 'docx');
}
/** PDF 로 바로 저장 (인쇄 대화상자 없이) */
export async function savePdf(opts = {}) {
  const name = `${S.docName}.pdf`;
  // 이 컴퓨터의 글꼴을 PDF 에 넣기 (처음 한 번 브라우저가 허용을 물음) — 안 되면 보는 프로그램의 한글 글꼴로
  if (S.opts.pdfEmbedFonts !== false) await askLocalFontsOnce().catch(() => false);
  const h = await pickSave(name, 'pdf');
  if (h === false) return;
  const bar = progressOverlay('PDF 만드는 중…');
  try {
    const bytes = await exportPdfBytes({ ask: true, embed: S.opts.pdfEmbedFonts !== false, ...opts, onProgress: (i, n) => bar.set(i / n, `${i} / ${n} 쪽`) });
    const blob = new Blob([bytes], { type: 'application/pdf' });
    if (h) await writeHandle(h, blob); else download(blob, name);
    toast(`PDF 로 저장했습니다 (${fmtSize(blob.size)})`);
  } finally { bar.close(); }
}
export async function saveAs(kind = 'docx') {
  if (kind === 'pdf') { await savePdf(); return; }
  const name = `${S.docName}.${kind}`;
  const h = await pickSave(name, kind);
  if (h === false) return;
  const blob = await docBlob(kind);
  if (h) {
    await writeHandle(h, blob);
    if (['docx', 'hwpx', 'md'].includes(kind)) { S.fileHandle = h; S.fileKind = kind; S.docName = h.name.replace(/\.[^.]+$/, ''); }
  } else download(blob, name);
  S.dirty = false;
  emit('saved');
  toast(`${KIND_LABEL[kind] ?? kind} 형식으로 저장했습니다 (${fmtSize(blob.size)})`);
}

// ───────────── 인쇄 · PDF ─────────────
export function printDoc({ from = 1, to = null, pdf = false } = {}) {
  const res = S.layout ?? relayout();
  const area = document.getElementById('printArea');
  area.replaceChildren();
  const pages = res.pages.slice(Math.max(0, from - 1), to ?? res.pages.length);
  const first = pages[0] ?? res.pages[0];
  const style = document.createElement('style');
  style.textContent = `@page { size: ${first.w}px ${first.h}px; margin: 0 }`;
  area.append(style);
  for (const p of pages) {
    const n = pageNode(S.doc, p, res, { marks: false, print: true });
    n.style.transform = '';
    area.append(n);
  }
  if (pdf) toast('인쇄 대화상자의 대상에서 [PDF로 저장]을 고르세요');
  setTimeout(() => { window.print(); setTimeout(() => area.replaceChildren(), 500); }, 50);
}

// ───────────── 자동 저장 · 최근 문서 (IndexedDB) ─────────────
const DB = 'wiword';
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
  await new Promise((res, rej) => { const tx = db.transaction('kv', 'readwrite'); if (value == null) tx.objectStore('kv').delete(key); else tx.objectStore('kv').put(value, key); tx.oncomplete = res; tx.onerror = () => rej(tx.error); });
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
  autoTimer = setTimeout(autosaveNow, 1500);
}
export async function autosaveNow() {
  try {
    const json = serialize(S.doc);
    await idbSet('autosave', { name: S.docName, json, at: Date.now(), kind: S.fileKind ?? null });
    const id = S.docId ?? (S.docId = `d${Date.now().toString(36)}`);
    const list = (await idbGet('recent')) ?? [];
    const item = { id, name: S.docName, at: Date.now(), words: docText(S.doc).length, preview: docText(S.doc).slice(0, 120) };
    const out = [item, ...list.filter((x) => x.id !== id)].slice(0, 20);
    await idbSet('recent', out);
    await idbSet(`doc:${id}`, { name: S.docName, json });
    for (const x of list.slice(19)) if (!out.some((y) => y.id === x.id)) await idbSet(`doc:${x.id}`, null);
    emit('autosaved');
  } catch (e) { console.warn('자동 저장 실패', e); }
}
export async function restoreAutosave() {
  try {
    const v = await idbGet('autosave');
    if (!v?.json) return false;
    setDoc(deserialize(v.json), v.name);
    return true;
  } catch { return false; }
}
export async function recentDocs() { return (await idbGet('recent').catch(() => null)) ?? []; }
export async function openRecent(id) {
  const v = await idbGet(`doc:${id}`);
  if (!v) { toast('문서를 찾을 수 없습니다'); return; }
  setDoc(deserialize(v.json), v.name);
  S.docId = id;
}
export async function removeRecent(id) {
  await idbSet('recent', (await recentDocs()).filter((x) => x.id !== id));
  await idbSet(`doc:${id}`, null);
}

// ───────────── 그림 ─────────────
const fileToDataUrl = (f) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsDataURL(f); });
async function imageSize(src) {
  try { const img = new Image(); img.src = src; await img.decode(); return { w: img.naturalWidth || 300, h: img.naturalHeight || 200 }; } catch { return { w: 300, h: 200 }; }
}
/** 그림을 본문 너비 안으로 맞춤 */
export async function insertImageFile(f) {
  const src = await fileToDataUrl(f);
  const { w, h } = await imageSize(src);
  const max = (S.layout?.pages[0]?.bodyW ?? 600);
  const k = Math.min(1, max / w);
  run('insertImage', src, Math.round(w * k), Math.round(h * k), f.name.replace(/\.[^.]+$/, ''));
}
on('pasteImages', async (files) => { for (const f of files) await insertImageFile(f); });
on('pasteMarkdown', (text) => {
  const d = mdToDoc(text, { doc: { ...newDoc(), styles: S.doc.styles } });
  pasteBlocks(d.body, d.media, d.styles, d.lists);
  toast('마크다운을 서식 있는 글로 붙여넣었습니다 (Ctrl+Shift+V = 텍스트만)');
});

// ───────────── 끌어 놓기 ─────────────
export function initDrop(target) {
  target.addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('Files')) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; } });
  target.addEventListener('drop', async (e) => {
    const files = [...(e.dataTransfer?.files ?? [])];
    if (!files.length) return;
    e.preventDefault();
    const imgs = files.filter((f) => f.type.startsWith('image/'));
    if (imgs.length) { for (const f of imgs) await insertImageFile(f); return; }
    const f = files[0];
    if (/\.(ttf|otf|woff2?)$/i.test(f.name)) { const n = await importFontFiles(files); toast(`글꼴 ${n?.length ?? n ?? 0}개를 불러왔습니다`); return; }
    if (S.dirty && !confirm(`'${f.name}' 을(를) 열까요? (현재 문서는 자동 저장되어 [파일 › 열기 › 최근 항목]에 남습니다)`)) {
      return;
    }
    await openFile(f);
  });
}

// ───────────── 명령 ─────────────
register({
  open: () => openWithPicker(),
  save: () => save(),
  saveAs: (kind) => (kind ? saveAs(kind) : run('backstage', 'saveAs')),
  saveHwpx: () => saveAs('hwpx'),
  exportPdf: (o) => savePdf(o ?? {}),
  printPdf: () => printDoc({ pdf: true }),
  print: () => run('backstage', 'print'),
  printNow: (o) => printDoc(o ?? {}),
  newBlank: () => run('newDocument'),
  readAnyFile: async (f) => { try { return (await loadFile(f)).doc; } catch (e) { alertDialog('열 수 없음', e.message); return null; } },
  openRecent: (id) => openRecent(id),
  insertPicture: async () => { const fs = await pickFile('image/*', true); for (const f of fs ?? []) await insertImageFile(f); },
  insertPictureUrl: () => formDialog('온라인 그림', [{ name: 'url', label: '그림 주소 (https://…)', value: '' }, { name: 'alt', label: '대체 텍스트', value: '' }], async (v) => {
    if (!/^https?:|^data:/.test(v.url)) { toast('그림 주소를 입력하세요'); return false; }
    const { w, h } = await imageSize(v.url);
    const k = Math.min(1, (S.layout?.pages[0]?.bodyW ?? 600) / w);
    run('insertImage', v.url, Math.round(w * k), Math.round(h * k), v.alt);
  }),
  changePicture: async () => {
    const o = S.objSel;
    if (!o) { toast('그림을 먼저 선택하세요'); return; }
    const f = await pickFile('image/*');
    if (!f) return;
    const src = await fileToDataUrl(f);
    const id = uid('m');
    change((d) => { const m = d.prop('media'); m[id] = src; const b = d.block(o.p); let k = 0; b.runs = b.runs.map((r) => { const at = k; k += r.obj ? 1 : (r.text?.length ?? 0); return at === o.o && r.obj === 'img' ? { ...r, media: id, src: undefined } : r; }); return S.sel; });
  },
  insertFileContents: async () => {
    const f = await pickFile(OPEN_ACCEPT);
    if (!f) return;
    const r = await loadFile(f).catch((e) => { alertDialog('열 수 없음', e.message); return null; });
    if (!r) return;
    pasteBlocks(r.doc.body, r.doc.media, r.doc.styles, r.doc.lists);
  },
  insertCsvTable: async () => {
    const f = await pickFile('.csv,.tsv,.txt');
    if (!f) return;
    const d = csvToDoc(decodeText(new Uint8Array(await f.arrayBuffer())));
    pasteBlocks(d.body.slice(0, 1));
  },
  importFonts: async (done) => { const fs = await pickFile('.ttf,.otf,.woff,.woff2', true); if (!fs?.length) return; const r = await importFontFiles(fs); toast(`글꼴을 불러왔습니다: ${(r ?? []).map?.((x) => x.name ?? x).join(', ') ?? ''}`); if (typeof done === 'function') done(); relayout(); },
  embedFonts: () => run('docProps', { embedFonts: !S.doc.props?.embedFonts }),
  copy: () => { focusEditor(); if (!document.execCommand('copy')) toast('Ctrl+C 를 누르세요'); },
  cut: () => { focusEditor(); if (!document.execCommand('cut')) toast('Ctrl+X 를 누르세요'); },
  paste: async (mode) => {
    focusEditor();
    try {
      const items = await navigator.clipboard.read();
      const data = {};
      const files = [];
      for (const it of items) for (const t of it.types) {
        const b = await it.getType(t);
        if (t.startsWith('image/')) files.push(new File([b], `그림.${t.split('/')[1]}`, { type: t }));
        else data[t] = await b.text();
      }
      pasteFrom({ getData: (t) => data[t] ?? '', files }, { plain: mode === 'text' });
    } catch {
      try { const t = await navigator.clipboard.readText(); pasteFrom({ getData: (k) => (k === 'text/plain' ? t : ''), files: [] }); } catch { toast('붙여넣기: Ctrl+V 를 누르세요 (브라우저가 클립보드 읽기를 막았습니다)'); }
    }
  },
  pastePlain: () => run('paste', 'text'),
}, {
  open: { label: '열기' }, save: { label: '저장' }, saveAs: { label: '다른 이름으로 저장', args: ['docx|hwpx|md|html|txt|pdf|wiword'] }, exportPdf: { label: 'PDF 로 내보내기' }, printPdf: { label: '인쇄 대화상자로 PDF 만들기' }, printNow: { label: '인쇄' },
  insertPicture: { label: '그림 넣기 (파일)' }, insertPictureUrl: { label: '온라인 그림' }, insertFileContents: { label: '파일의 텍스트 넣기' }, copy: { label: '복사' }, cut: { label: '잘라내기' }, paste: { label: '붙여넣기' },
});

