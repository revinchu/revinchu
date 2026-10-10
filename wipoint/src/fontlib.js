// 글꼴 (브라우저): 이 컴퓨터의 글꼴 찾기 + 한글 웹 글꼴 자동 내려받기 + PowerPoint 같은 글꼴 목록
//   이 컴퓨터: Local Font Access API (권한이 이미 있으면 시작할 때 바로, 없으면 글꼴 목록을 처음 열 때 물어봄)
//              + 권한 없이 흔한 글꼴 120여 개를 글자 너비로 확인
import { webFont, WEB_FONTS, WEB_FONT_CATS, fontsUsed } from './webfonts.js';

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
const isLocal = (name) => (localAll ?? detected ?? []).some((n) => n.toLowerCase() === String(name).toLowerCase());

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
  const theme = [['+mj', `${pres.theme.fonts.major}`, '제목'], ['+mn', `${pres.theme.fonts.minor}`, '본문']].filter(([, n]) => ok(n)).map(([k, n, role]) => ({ name: k, label: n, note: role, face: n }));
  const doc = (pres.fonts ?? []).map((f) => f.typeface).filter((n, i, a) => a.indexOf(n) === i && ok(n)).map((n) => ({ name: n, label: n, face: n, note: '포함됨' }));
  const recent = recentFonts().filter((n) => ok(n)).map((n) => ({ name: n, label: n, face: n }));
  const local = localFontList().filter((n) => ok(n)).map((n) => ({ name: n, label: n, face: n }));
  const out = [['테마 글꼴', theme], ['문서에 포함된 글꼴', doc], ['최근에 사용한 글꼴', recent], [localAll ? `이 컴퓨터의 글꼴 (${localAll.length})` : '이 컴퓨터의 글꼴', local]];
  for (const cat of WEB_FONT_CATS) out.push([`한글 웹 글꼴 · ${cat}`, WEB_FONTS.filter((f) => f.cat === cat && ok(f.family, f.ko)).map((f) => ({ name: f.family, label: f.ko === f.family ? f.family : `${f.ko}`, face: f.family, web: true, note: f.src === 'google' ? 'Google' : '웹' }))]);
  return out.filter(([, l]) => l.length);
}
