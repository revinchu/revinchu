// 기기 글꼴 + 라이선스가 확인된 무료 웹 글꼴. 목록/이름 조회는 네트워크 요청을 만들지 않는다.
import { WEB_FONT_CATALOG } from './web-font-catalog.js';

// 한글 이름 ↔ 영문 이름 (윈도우는 한글 이름, 맥 · 다른 브라우저는 영문 이름만 아는 경우가 있음)
const ALIASES = [
  ['맑은 고딕', 'Malgun Gothic'], ['굴림', 'Gulim'], ['굴림체', 'GulimChe'], ['돋움', 'Dotum'], ['돋움체', 'DotumChe'],
  ['바탕', 'Batang'], ['바탕체', 'BatangChe'], ['궁서', 'Gungsuh'], ['궁서체', 'GungsuhChe'],
  ['나눔고딕', 'NanumGothic'], ['나눔명조', 'NanumMyeongjo'], ['나눔바른고딕', 'NanumBarunGothic'], ['나눔스퀘어', 'NanumSquare'],
  ['나눔스퀘어라운드', 'NanumSquareRound'], ['나눔손글씨 펜', 'Nanum Pen Script'], ['나눔고딕코딩', 'NanumGothicCoding'],
  ['본고딕', 'Source Han Sans KR'], ['본명조', 'Source Han Serif KR'], ['애플 SD 산돌고딕 Neo', 'Apple SD Gothic Neo'],
  ['프리텐다드', 'Pretendard'], ['스포카 한 산스 Neo', 'Spoqa Han Sans Neo'], ['에스코어 드림', 'S-Core Dream'],
  ['KoPub돋움체', 'KoPubDotum'], ['KoPub바탕체', 'KoPubBatang'], ['한컴 고딕', 'Hancom Gothic'], ['함초롬돋움', 'HCR Dotum'],
  ['함초롬바탕', 'HCR Batang'], ['휴먼명조', 'HYMyeongJo'], ['HY견고딕', 'HYGothic'], ['G마켓 산스', 'Gmarket Sans'],
];
const toEn = new Map(ALIASES.map(([k, e]) => [k.toLowerCase(), e]));
const toKo = new Map(ALIASES.map(([k, e]) => [e.toLowerCase(), k]));

/** 같은 글꼴의 다른 이름 (없으면 null) */
export const fontAlias = (name) => toEn.get(String(name).toLowerCase()) ?? toKo.get(String(name).toLowerCase()) ?? null;

// 확인해 볼 흔한 글꼴 (한글 + 영문)
const CANDIDATES = [
  ...ALIASES.map(([k]) => k),
  'Arial', 'Arial Black', 'Arial Narrow', 'Calibri', 'Calibri Light', 'Cambria', 'Candara', 'Century Gothic', 'Comic Sans MS', 'Consolas',
  'Constantia', 'Corbel', 'Courier New', 'Franklin Gothic Medium', 'Garamond', 'Georgia', 'Helvetica', 'Helvetica Neue', 'Impact',
  'Lucida Console', 'Lucida Sans Unicode', 'Palatino Linotype', 'Segoe UI', 'Segoe UI Light', 'Tahoma', 'Times New Roman',
  'Trebuchet MS', 'Verdana', 'Aptos', 'Aptos Display', 'Roboto', 'Noto Sans KR', 'Noto Serif KR', 'Noto Sans', 'Open Sans',
  'Lato', 'Montserrat', 'Inter', 'SF Pro Text', 'Menlo', 'Monaco', 'Gill Sans', 'Optima', 'Avenir', 'Futura', 'Baskerville',
  'MS Gothic', 'MS Mincho', 'Meiryo', 'Yu Gothic', 'Microsoft YaHei', 'SimSun', 'DengXian',
];
export const DEFAULT_FONTS = ['맑은 고딕', '굴림', '돋움', '바탕', '궁서', 'Arial', 'Calibri', 'Consolas', 'Times New Roman', 'Verdana'];

let detected = null;
let localList = null;

/** 글꼴이 설치되어 있는지: 기본 글꼴과 글자 너비가 다르면 있음 */
function isInstalled(name, ctx) {
  const sample = 'mmmmmmmmmmlliWW가나다라마바사0123';
  for (const base of ['monospace', 'serif', 'sans-serif']) {
    ctx.font = `72px ${base}`;
    const w0 = ctx.measureText(sample).width;
    ctx.font = `72px '${name}', ${base}`;
    if (Math.abs(ctx.measureText(sample).width - w0) > 0.5) return true;
  }
  return false;
}

/** 흔한 글꼴 중 설치된 것 (권한 없이 바로) */
export function detectFonts() {
  if (detected) return detected;
  const ctx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
  if (!ctx) return DEFAULT_FONTS;
  const found = [];
  for (const name of CANDIDATES) {
    const alias = fontAlias(name);
    if (isInstalled(name, ctx) || (alias && isInstalled(alias, ctx))) found.push(name);
  }
  detected = [...new Set([...found])];
  return detected;
}

export const canListLocalFonts = () => typeof globalThis.queryLocalFonts === 'function';

/** 이 PC의 모든 글꼴 (브라우저가 권한을 물어봄 — 크롬 · 엣지) */
export async function loadLocalFonts() {
  if (localList) return localList;
  if (!canListLocalFonts()) return null;
  const fonts = await globalThis.queryLocalFonts();
  const fams = new Set();
  for (const f of fonts) fams.add(f.family);
  localList = [...fams].sort((a, b) => a.localeCompare(b, 'ko'));
  return localList;
}

/** 무료 웹 글꼴은 기기 글꼴 접근 권한과 관계없이 모든 브라우저에 제공한다. */
export function fontList() {
  return [...new Set([...DEFAULT_FONTS, ...(localList ?? detectFonts()), ...WEB_FONT_CATALOG.map((f) => f.family)])];
}


const fontKey = (name) => String(name ?? '').trim().toLocaleLowerCase().replace(/\s+/g, '');
const webFontIndex = new Map();
for (const font of WEB_FONT_CATALOG) {
  for (const name of [font.family, font.label, ...(font.aliases || [])]) if (name) webFontIndex.set(fontKey(name), font);
}
/** CSS 이름/한글 별칭을 같은 무료 웹 글꼴로 해석한다. 임의의 URL은 받지 않는다. */
export function getWebFont(name) { return webFontIndex.get(fontKey(name)) ?? webFontIndex.get(fontKey(fontAlias(name))) ?? null; }
export function fontFamilyCandidates(name) {
  const font = getWebFont(name);
  return [...new Set([String(name || '').trim(), font?.family, fontAlias(name)].filter(Boolean))];
}
export function fontLabel(name) {
  const f = getWebFont(name);
  return f?.label && f.label !== f.family ? f.label + ' · ' + f.family : String(name);
}
export function fontMatches(name, query) {
  const q = fontKey(query), f = getWebFont(name);
  return !q || [name, fontAlias(name), f?.family, f?.label, ...(f?.aliases || [])].some((x) => x && fontKey(x).includes(q));
}
export function fontIsKorean(name) { return !!getWebFont(name)?.scripts?.includes('korean'); }
export function fontInSource(name, source) {
  const f = getWebFont(name);
  return !source || source === 'all' || (source === 'korean' ? fontIsKorean(name) : source === 'local' ? !f : f?.source === source);
}
// Exact origins only. This registry is independent of workbook strings and contains no user-provided URL.
export function isAllowedFontUrl(value) {
  try {
    const u = new URL(value);
    return !u.username && !u.password && ['https://fonts.googleapis.com', 'https://fonts.gstatic.com', 'https://cdn.jsdelivr.net', 'https://hangeul.pstatic.net', 'https://script.gmarket.com', 'https://img.cafe24.com', 'https://framerusercontent.com'].includes(u.origin);
  } catch { return false; }
}
export function webFontCssUrl(nameOrEntry) {
  const f = typeof nameOrEntry === 'string' ? getWebFont(nameOrEntry) : nameOrEntry;
  if (!f) return null;
  if (f.cssUrl) return isAllowedFontUrl(f.cssUrl) ? f.cssUrl : null;
  if (f.source !== 'google') return null;
  const variants = (f.variants || []).map((v) => typeof v === 'string' && /^\d{1,4}i?$/.test(v) ? [v.endsWith('i') ? 1 : 0, Number(v.replace('i', ''))] : null).filter((v) => v && v[1] >= 1 && v[1] <= 1000);
  const weights = (f.weights || [400]).filter((w) => Number.isInteger(w) && w >= 1 && w <= 1000);
  const tuples = variants.length ? variants : (f.styles?.includes('italic') ? [0, 1] : [0]).flatMap((i) => weights.map((w) => [i, w]));
  tuples.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const italic = tuples.some(([i]) => i);
  const axis = tuples.length ? (italic ? ':ital,wght@' + tuples.map(([i,w]) => i + ',' + w).join(';') : ':wght@' + tuples.map(([,w]) => w).join(';')) : '';
  return 'https://fonts.googleapis.com/css2?family=' + encodeURIComponent(f.family).replace(/%20/g, '+') + axis + '&display=swap';
}

const webLoads = new Map();
const fontListeners = new Set();
let fontSetListener = false;
function fontEvent(detail) { for (const listener of fontListeners) { try { listener(detail); } catch { /* A preview cannot break another renderer. */ } } }
export function onWebFontChange(listener) {
  fontListeners.add(listener);
  if (!fontSetListener && typeof document !== 'undefined' && document.fonts?.addEventListener) {
    fontSetListener = true;
    // Dynamic Korean/CJK subsets can finish after the first Latin face.
    document.fonts.addEventListener('loadingdone', () => fontEvent({ status: 'loaded', family: null }));
  }
  return () => fontListeners.delete(listener);
}
export function webFontStatus(name) {
  const f = getWebFont(name);
  return f ? webLoads.get(f.family)?.status || 'idle' : 'local';
}
const cssQuote = (value) => '"' + String(value).replace(/["\\\n\r\f]/g, '') + '"';
/** Download only a used/previewed font. Completion never changes document styles, history or selection. */
export function requestWebFont(name, options = {}) {
  const f = getWebFont(name);
  if (!f || typeof document === 'undefined' || !document.fonts?.load) return Promise.resolve({ status: 'local', family: f?.family || name });
  const old = webLoads.get(f.family);
  if (old && !(options.retry && old.status === 'error')) return old.promise;
  const state = { status: 'loading', promise: null };
  webLoads.set(f.family, state);
  state.promise = new Promise((resolve) => {
    let finished = false, node = null;
    const done = (status) => {
      if (finished) return;
      finished = true;
      clearTimeout(timer);
      if (status === 'error') node?.remove();
      state.status = status;
      const detail = { status, family: f.family };
      resolve(detail);
      fontEvent(detail);
    };
    const timer = setTimeout(() => done('error'), 25000);
    const load = async () => {
      try {
        const weights = f.weights?.length ? f.weights : [400, 700];
        const normal = weights.reduce((a, b) => Math.abs(b - 400) < Math.abs(a - 400) ? b : a, weights[0]);
        const bold = weights.reduce((a, b) => Math.abs(b - 700) < Math.abs(a - 700) ? b : a, weights[0]);
        const sample = fontIsKorean(f.family) ? '가나다Aa0123' : 'Aa0123';
        const italicOnly = f.styles?.length === 1 && f.styles[0] === 'italic';
        const faces = await Promise.all([...new Set([normal, bold])].map((w) => document.fonts.load((italicOnly ? 'italic ' : '') + w + ' 16px ' + cssQuote(f.family), sample)));
        let registered = false;
        document.fonts.forEach?.((face) => { if (fontKey(String(face.family).replace(/["']/g, '')) === fontKey(f.family)) registered = true; });
        done(faces.some((arr) => arr.length) || registered ? 'loaded' : 'error');
      } catch { done('error'); }
    };
    const url = webFontCssUrl(f);
    if (url) {
      node = document.createElement('link');
      node.rel = 'stylesheet'; node.href = url;
      node.crossOrigin = 'anonymous'; node.referrerPolicy = 'no-referrer';
      node.dataset.webFont = f.family;
      node.onload = load; node.onerror = () => done('error');
      document.head.append(node);
    } else if (f.faces?.length && f.faces.every((face) => isAllowedFontUrl(face.url))) {
      node = document.createElement('style');
      node.dataset.webFont = f.family;
      node.textContent = f.faces.map((face) => '@font-face{font-family:' + cssQuote(f.family) + ';src:url(' + cssQuote(face.url) + ');font-weight:' + (/^[\d ]+$/.test(String(face.weight)) ? face.weight : 400) + ';font-style:' + (face.style === 'italic' ? 'italic' : 'normal') + ';font-display:swap;' + (face.unicodeRange && /^[uU+\da-fA-F?,\s-]+$/.test(face.unicodeRange) ? 'unicode-range:' + face.unicodeRange + ';' : '') + '}').join('\n');
      document.head.append(node);
      load();
    } else done('error');
  });
  return state.promise;
}
