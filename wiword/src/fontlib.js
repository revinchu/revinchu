// 글꼴 (브라우저): 이 컴퓨터의 글꼴 찾기 + 한글 웹 글꼴 자동 내려받기 + PowerPoint 같은 글꼴 목록
//   이 컴퓨터: Local Font Access API (권한이 이미 있으면 시작할 때 바로, 없으면 글꼴 목록을 처음 열 때 물어봄)
//              + 권한 없이 흔한 글꼴 120여 개를 글자 너비로 확인
import { webFont, WEB_FONTS, WEB_FONT_CATS, fontsUsed } from './webfonts.js';
import { fontKind, woffToSfnt, sfntInfo, embeddable, makeEot } from './fonts.js';

const KO_CANDIDATES = ['맑은 고딕', 'Malgun Gothic', '굴림', '굴림체', '돋움', '돋움체', '바탕', '바탕체', '궁서', '궁서체', '나눔고딕', 'NanumGothic', '나눔명조', '나눔스퀘어', 'NanumSquare', '나눔스퀘어 네오', '나눔스퀘어라운드', '나눔바른고딕', '나눔바른펜', '나눔손글씨 펜', '나눔고딕코딩', '본고딕', 'Source Han Sans K', '본명조', 'Apple SD Gothic Neo', 'AppleMyungjo', 'AppleGothic', '프리텐다드', 'Pretendard', 'SUIT', '스포카 한 산스 Neo', 'Spoqa Han Sans Neo', '에스코어 드림', 'S-Core Dream', 'G마켓 산스', 'Gmarket Sans', 'KoPub돋움체', 'KoPub바탕체', 'KoPubWorld돋움체', 'KoPubWorld바탕체', '한컴 고딕', '한컴산뜻돋움', '함초롬돋움', '함초롬바탕', 'HY견고딕', 'HY헤드라인M', 'HY중고딕', 'HY신명조', 'HY견명조', '휴먼명조', '휴먼고딕', '휴먼엑스포', '양재튼튼체B', '윤고딕', '산돌고딕 Neo', '배달의민족 주아', '배달의민족 도현', '배달의민족 한나는 열한살', '배달의민족 연성', '여기어때 잘난체', '티몬 몬소리', '넥슨Lv1고딕', '넥슨Lv2고딕', '마루 부리', 'MaruBuri', '카카오 Regular', 'Cafe24 Ssurround', 'Cafe24 Ohsquare', '페이퍼로지', 'Paperlogy', '원티드 산스', 'Wanted Sans', 'D2Coding', 'Noto Sans KR', 'Noto Serif KR', 'Noto Sans CJK KR', 'Noto Serif CJK KR', 'IBM Plex Sans KR', 'Gowun Dodum', 'Gowun Batang', 'Hahmlet'];
const EN_CANDIDATES = ['Arial', 'Arial Black', 'Arial Narrow', 'Calibri', 'Calibri Light', 'Cambria', 'Cambria Math', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas', 'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Helvetica', 'Helvetica Neue', 'Impact', 'Lucida Console', 'Lucida Sans Unicode', 'Palatino Linotype', 'Segoe UI', 'Segoe UI Light', 'Segoe UI Semibold', 'Segoe Print', 'Segoe Script', 'Tahoma', 'Times New Roman', 'Trebuchet MS', 'Verdana', 'Aptos', 'Aptos Display', 'Aptos Narrow', 'Bahnschrift', 'Gill Sans MT', 'Book Antiqua', 'Bookman Old Style', 'Century', 'Rockwell', 'Tw Cen MT', 'Roboto', 'Open Sans', 'Lato', 'Montserrat', 'Inter', 'Poppins', 'SF Pro Text', 'SF Pro Display', 'Menlo', 'Monaco', 'Gill Sans', 'Optima', 'Avenir', 'Avenir Next', 'Futura', 'Baskerville', 'Didot', 'Hoefler Text', 'MS Gothic', 'MS Mincho', 'Meiryo', 'Yu Gothic', 'Microsoft YaHei', 'SimSun', 'DengXian', 'Ubuntu', 'DejaVu Sans', 'Liberation Sans', 'Liberation Serif', 'Cantarell'];

let detected = null;
let localAll = null;
const RECENT_KEY = 'wipoint:recentFonts';

function installed(name, ctx) {
  const sample = 'mmmmmmmmmmlliWW가나다라마바사0123';
  for (const base of ['monospace', 'serif', 'sans-serif']) {
    ctx.font = `72px ${base}`;
    const w0 = ctx.measureText(sample).width;
    ctx.font = `72px '${name}', ${base}`;
    if (Math.abs(ctx.measureText(sample).width - w0) > 0.5) return true;
  }
  return false;
}
/** 권한 없이 확인되는 이 컴퓨터의 글꼴 */
export function detectLocalFonts() {
  if (detected) return detected;
  const ctx = document.createElement('canvas').getContext('2d');
  detected = [...new Set([...KO_CANDIDATES, ...EN_CANDIDATES])].filter((n) => installed(n, ctx));
  return detected;
}
export const canQueryLocalFonts = () => typeof globalThis.queryLocalFonts === 'function';
/** 이 컴퓨터의 모든 글꼴 (Local Font Access). ask=false 이면 이미 허용된 경우에만 */
export async function loadAllLocalFonts({ ask = true } = {}) {
  if (localAll) return localAll;
  if (!canQueryLocalFonts()) return null;
  try {
    if (!ask) {
      const st = await navigator.permissions?.query?.({ name: 'local-fonts' }).catch(() => null);
      if (st?.state !== 'granted') return null;
    }
    const fonts = await globalThis.queryLocalFonts();
    localAll = [...new Set(fonts.map((f) => f.family))].sort((a, b) => a.localeCompare(b, 'ko'));
    return localAll;
  } catch { return null; }
}
export const localFontList = () => localAll ?? detectLocalFonts();
const isLocal = (name) => (localAll ?? detected ?? []).some((n) => n.toLowerCase() === String(name).toLowerCase()) || userFonts.has(String(name).toLowerCase());

// ───── 불러온 글꼴 파일 (.ttf · .otf · .woff · .woff2) ─────
// 모든 언어의 이름(한글 · 영문)으로 FontFace 를 등록 → PowerPoint 파일에 어느 이름으로 적혀 있어도 그 글꼴로 그림.
// 브라우저 IndexedDB (wipoint-fonts) 에 보관해 다음에도 그대로. 저장할 때 [글꼴 포함]이면 TrueType 은 pptx 에 넣음.
const userFonts = new Map(); // 소문자 이름 → { name, names, kind, bytes(sfnt 또는 woff2 원본), info }
const FDB = 'wipoint-fonts';
const fdb = () => new Promise((resolve, reject) => {
  const r = indexedDB.open(FDB, 1);
  r.onupgradeneeded = () => r.result.createObjectStore('f');
  r.onsuccess = () => resolve(r.result);
  r.onerror = () => reject(r.error);
});
async function fdbPut(key, val) { const db = await fdb(); await new Promise((res, rej) => { const t = db.transaction('f', 'readwrite'); t.objectStore('f').put(val, key); t.oncomplete = res; t.onerror = () => rej(t.error); }); }
async function fdbAll() { const db = await fdb(); return new Promise((res, rej) => { const t = db.transaction('f'); const q = t.objectStore('f').getAll(); q.onsuccess = () => res(q.result ?? []); q.onerror = () => rej(q.error); }); }
async function fdbDel(key) { const db = await fdb(); await new Promise((res) => { const t = db.transaction('f', 'readwrite'); t.objectStore('f').delete(key); t.oncomplete = res; t.onerror = res; }); }

/** 글꼴 바이트 등록 → 대표 이름 (못 읽으면 Error) */
async function registerFontBytes(raw, fileName = '') {
  const kind = fontKind(raw);
  if (!kind) throw new Error(`${fileName}: 글꼴 파일(.ttf · .otf · .woff · .woff2)이 아닙니다`);
  if (kind === 'ttc') throw new Error(`${fileName}: 글꼴 모음(.ttc)은 지원하지 않습니다. 각각의 .ttf 로 불러와 주세요`);
  const bytes = kind === 'woff' ? woffToSfnt(raw) : raw;
  const info = kind === 'woff2' ? { families: [fileName.replace(/\.[^.]+$/, '')], family: fileName.replace(/\.[^.]+$/, ''), style: 'Regular', weight: 400, italic: false, fsType: 0 } : sfntInfo(bytes);
  const names = info.families.length ? info.families : [fileName.replace(/\.[^.]+$/, '')];
  for (const n of names) {
    const face = new FontFace(n, bytes, { weight: String(info.weight ?? 400), style: info.italic ? 'italic' : 'normal' });
    await face.load();
    document.fonts.add(face);
  }
  const rec = { name: info.family || names[0], names, kind: kind === 'woff' ? 'ttf' : kind, info, bytes };
  for (const n of names) userFonts.set(n.toLowerCase(), rec);
  return rec;
}
/** 파일 선택 상자에서 고른 글꼴 파일들 → 등록 + 보관. 반환: { ok: [이름], fail: [메시지] } */
export async function importFontFiles(files) {
  const ok = []; const fail = [];
  for (const file of files) {
    try {
      const raw = new Uint8Array(await file.arrayBuffer());
      const rec = await registerFontBytes(raw, file.name);
      ok.push(rec.name);
      try { await fdbPut(`${rec.name}|${rec.info.weight}|${rec.info.italic ? 1 : 0}`, { file: file.name, raw }); } catch { /* 보관 못 해도 이번에는 씀 */ }
    } catch (e) { fail.push(e.message); }
  }
  return { ok, fail };
}
/** 시작할 때: 전에 불러온 글꼴 다시 등록 */
export async function restoreUserFonts() {
  if (typeof indexedDB === 'undefined' || typeof FontFace !== 'function') return 0;
  let n = 0;
  try { for (const r of await fdbAll()) { try { await registerFontBytes(r.raw, r.file); n++; } catch { /* 건너뜀 */ } } } catch { /* 없음 */ }
  return n;
}
/** 불러온 글꼴 지우기 (브라우저 보관에서도) */
export async function removeUserFont(name) {
  const rec = userFonts.get(String(name).toLowerCase());
  if (!rec) return;
  for (const n of rec.names) userFonts.delete(n.toLowerCase());
  await fdbDel(`${rec.name}|${rec.info.weight}|${rec.info.italic ? 1 : 0}`);
}
export const userFontList = () => [...new Set([...userFonts.values()].map((r) => r.name))].sort((a, b) => a.localeCompare(b, 'ko'));

/**
 * 파일에 포함할 글꼴 (PowerPoint [파일의 글꼴 포함]): 문서가 쓰는 글꼴 중 불러온 글꼴 · 이 컴퓨터의 글꼴(허용된 경우)에서
 * TrueType 이고 포함이 허용된(fsType) 것 → [{ typeface, style, data(EOT data URL) }]. 이미 포함된 글꼴은 빼고, skipped 에 못 넣은 이유
 */
export async function fontsToEmbed(pres) {
  const have = new Set((pres.fonts ?? []).map((f) => f.typeface.toLowerCase()));
  const out = []; const skipped = [];
  const used = fontsUsed(pres).filter((n) => !have.has(n.toLowerCase()));
  const toUrl = (b) => { let s = ''; for (let i = 0; i < b.length; i += 0x8000) s += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return `data:application/x-fontdata;base64,${btoa(s)}`; };
  let local = null;
  for (const name of used) {
    let rec = userFonts.get(name.toLowerCase());
    if (!rec && canQueryLocalFonts()) {
      try {
        local ??= await globalThis.queryLocalFonts();
        const fd = local.find((f) => f.family.toLowerCase() === name.toLowerCase() && /^(regular|normal|보통|표준)$/i.test(f.style)) ?? local.find((f) => f.family.toLowerCase() === name.toLowerCase());
        if (fd) { const raw = new Uint8Array(await (await fd.blob()).arrayBuffer()); const k = fontKind(raw); if (k === 'ttf' || k === 'otf') rec = { name, kind: k, bytes: raw, info: sfntInfo(raw) }; }
      } catch { /* 권한 없음 */ }
    }
    if (!rec) continue;
    if (rec.kind !== 'ttf') { skipped.push(`${name}: TrueType 글꼴만 포함할 수 있습니다`); continue; }
    if (!embeddable(rec.info)) { skipped.push(`${name}: 글꼴 사용권이 파일 포함을 허용하지 않습니다`); continue; }
    out.push({ typeface: name, style: rec.info.weight >= 600 ? (rec.info.italic ? 'boldItalic' : 'bold') : rec.info.italic ? 'italic' : 'regular', data: toUrl(makeEot(rec.bytes, { ...rec.info, family: name })) });
  }
  return { fonts: out, skipped };
}

// ───── 웹 글꼴 ─────
const loading = new Map();
/** 글꼴이 웹 글꼴이고 이 컴퓨터에 없으면 내려받기 (CSS 연결). 다 받으면 resolve */
export function ensureFont(name) {
  const wf = webFont(name);
  if (!wf || typeof document === 'undefined') return Promise.resolve(false);
  if (loading.has(wf.family)) return loading.get(wf.family);
  if (isLocal(wf.family) || isLocal(name)) { loading.set(wf.family, Promise.resolve(true)); return loading.get(wf.family); }
  const p = new Promise((resolve) => {
    const link = document.createElement('link');
    link.rel = 'stylesheet';
    link.href = wf.css;
    link.crossOrigin = 'anonymous';
    link.onload = () => { document.fonts?.load?.(`16px "${wf.family}"`, '가A').then(() => resolve(true), () => resolve(true)); };
    link.onerror = () => { loading.delete(wf.family); resolve(false); };
    document.head.append(link);
  });
  loading.set(wf.family, p);
  return p;
}
/** 문서가 쓰는 웹 글꼴을 모두 받음 → 받은 것이 있으면 onDone (다시 그리기) */
export async function ensureDocFonts(pres, onDone) {
  const need = fontsUsed(pres).filter((n) => webFont(n));
  if (!need.length) return;
  const r = await Promise.all(need.map(ensureFont));
  if (r.some(Boolean)) onDone?.();
}

export function recentFonts() { try { return JSON.parse(localStorage.getItem(RECENT_KEY) ?? '[]'); } catch { return []; } }
export function touchRecent(name) {
  if (!name || name.startsWith('+')) return;
  const l = [name, ...recentFonts().filter((n) => n !== name)].slice(0, 8);
  try { localStorage.setItem(RECENT_KEY, JSON.stringify(l)); } catch { /* 없음 */ }
}

/** 글꼴 목록 구역들: [[제목, [{name, label, web?}]]] */
export function fontSections(pres, filter = '') {
  const q = filter.trim().toLowerCase();
  const ok = (n, label = '') => !q || n.toLowerCase().includes(q) || label.toLowerCase().includes(q);
  const tf = pres.themeFonts ? { major: pres.themeFonts.major?.ea ?? pres.themeFonts.major?.latin, minor: pres.themeFonts.minor?.ea ?? pres.themeFonts.minor?.latin } : pres.theme.fonts;
  const theme = [['+mj', `${tf.major}`, '제목'], ['+mn', `${tf.minor}`, '본문']].filter(([, n]) => ok(n)).map(([k, n, role]) => ({ name: k, label: n, note: role, face: n }));
  const doc = (pres.fonts ?? []).map((f) => f.typeface).filter((n, i, a) => a.indexOf(n) === i && ok(n)).map((n) => ({ name: n, label: n, face: n, note: '포함됨' }));
  const recent = recentFonts().filter((n) => ok(n)).map((n) => ({ name: n, label: n, face: n }));
  const local = localFontList().filter((n) => ok(n)).map((n) => ({ name: n, label: n, face: n }));
  const loaded = userFontList().filter((n) => ok(n)).map((n) => ({ name: n, label: n, face: n, note: '불러옴' }));
  const out = [['테마 글꼴', theme], ['문서에 포함된 글꼴', doc], ['최근에 사용한 글꼴', recent], ['불러온 글꼴 파일', loaded], [localAll ? `이 컴퓨터의 글꼴 (${localAll.length})` : '이 컴퓨터의 글꼴', local]];
  for (const cat of WEB_FONT_CATS) out.push([`무료 한글 글꼴 · ${cat}`, WEB_FONTS.filter((f) => f.cat === cat && ok(f.family, f.ko)).map((f) => ({ name: f.family, label: f.ko === f.family ? f.family : `${f.ko}`, face: f.family, web: true, note: f.src === 'google' ? 'Google' : '무료' }))]);
  return out.filter(([, l]) => l.length);
}

// ───── PDF 에 넣을 글꼴 파일 찾기 ─────
const FAMILY_ALIAS = { '맑은 고딕': 'Malgun Gothic', 굴림: 'Gulim', 굴림체: 'GulimChe', 돋움: 'Dotum', 돋움체: 'DotumChe', 바탕: 'Batang', 바탕체: 'BatangChe', 궁서: 'Gungsuh', 궁서체: 'GungsuhChe', 나눔고딕: 'NanumGothic', 나눔명조: 'NanumMyeongjo', 나눔바른고딕: 'NanumBarunGothic', 함초롬바탕: 'HCR Batang', 함초롬돋움: 'HCR Dotum', '본고딕': 'Noto Sans CJK KR', '맑은고딕': 'Malgun Gothic' };
const fileCache = new Map();
let localFontData = null;
/**
 * 글꼴 이름 → { bytes, ps(모음 안 글꼴의 PostScript 이름), fakeBold, fakeItalic } | null
 * 불러온 글꼴 → 이 컴퓨터의 글꼴(Local Font Access, ask=true 면 권한을 물음) 순서
 */
export async function fontFileFor(family, { bold = false, italic = false, ask = false } = {}) {
  const fam = String(family).replace(/^['"]|['"]$/g, '').trim();
  const key = `${fam}|${bold ? 1 : 0}|${italic ? 1 : 0}`.toLowerCase();
  if (fileCache.has(key)) return fileCache.get(key);
  const p = (async () => {
    const names = [fam, FAMILY_ALIAS[fam]].filter(Boolean).map((n) => n.toLowerCase());
    const rec = names.map((n) => userFonts.get(n)).find(Boolean);
    if (rec && rec.kind === 'ttf') return { bytes: rec.bytes, ps: null, fakeBold: bold && !(rec.info.weight >= 600), fakeItalic: italic && !rec.info.italic };
    if (!canQueryLocalFonts()) return null;
    try {
      if (!localFontData) {
        if (!ask) { const st = await navigator.permissions?.query?.({ name: 'local-fonts' }).catch(() => null); if (st?.state !== 'granted') return null; }
        localFontData = await globalThis.queryLocalFonts();
      }
      const cands = localFontData.filter((f) => names.includes(f.family.toLowerCase()) || names.includes(f.fullName.toLowerCase()));
      if (!cands.length) return null;
      const score = (f) => (/bold|black|heavy|semibold|demibold/i.test(f.style) === bold ? 2 : 0) + (/italic|oblique/i.test(f.style) === italic ? 1 : 0);
      const fd = cands.sort((a, b) => score(b) - score(a))[0];
      const bytes = new Uint8Array(await (await fd.blob()).arrayBuffer());
      return { bytes, ps: fd.postscriptName, fakeBold: bold && !/bold|black|heavy/i.test(fd.style), fakeItalic: italic && !/italic|oblique/i.test(fd.style) };
    } catch { return null; }
  })();
  fileCache.set(key, p);
  return p;
}
/** 이 컴퓨터의 글꼴을 PDF 에 넣어도 되는지 한 번 묻기 (사용자 동작 안에서) */
export async function askLocalFontsOnce() {
  if (!canQueryLocalFonts() || localFontData) return !!localFontData;
  try { localFontData = await globalThis.queryLocalFonts(); return true; } catch { return false; }
}
