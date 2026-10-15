// 글꼴: 이 PC에 설치된 글꼴 찾기 (Local Font Access API → 없으면 너비 비교로 확인), 한글 ↔ 영문 이름
// 셀은 CSS font-family 로 그리므로 설치된 글꼴이면 그대로 보임

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

/** 글꼴 목록: 불러온 전체 목록 > 확인된 흔한 글꼴 */
export function fontList() {
  return localList ?? detectFonts();
}
