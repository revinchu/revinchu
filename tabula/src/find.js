// 찾기 및 바꾸기 (엑셀과 같은 옵션): 와일드카드(* ? ~), 대/소문자 구분, 전체 셀 내용 일치,
// 찾는 위치(수식 · 값 · 메모), 검색 순서(행 · 열), 범위(시트 · 통합 문서), 서식으로 찾기 — DOM 없음
import { inBlock } from './block.js';

const escRe = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** 엑셀 찾기 문자열 → 정규식 (* = 아무 글자 여러 개, ? = 한 글자, ~* ~? ~~ = 글자 그대로) */
export function findRegex(text, { matchCase = false, whole = false, global = false } = {}) {
  let src = '';
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (ch === '~' && i + 1 < text.length && '*?~'.includes(text[i + 1])) { src += escRe(text[++i]); continue; }
    src += ch === '*' ? '[\\s\\S]*' : ch === '?' ? '[\\s\\S]' : escRe(ch);
  }
  if (whole) src = `^(?:${src})$`;
  return new RegExp(src, `${matchCase ? '' : 'i'}${global ? 'g' : ''}u`);
}

/** 찾은 글자를 바꿈 (전체 셀 내용 일치면 셀 내용 전체) */
export function replaceText(raw, opts, replacement) {
  const re = findRegex(opts.text, { ...opts, global: true });
  if (opts.whole) return re.test(raw) ? replacement : raw;
  return raw.replace(re, (m) => (m === '' ? m : replacement));
}

// 서식 비교에 쓰는 속성 (없음 = 기본값)
export const FIND_FORMAT_KEYS = ['bold', 'italic', 'underline', 'strike', 'color', 'fill', 'font', 'size', 'align', 'valign', 'numFmt', 'code', 'wrap', 'bt', 'bb', 'bl', 'br'];
const norm = (k, v) => {
  if (v === undefined || v === null || v === false || v === '') return null;
  if (typeof v === 'string' && v.startsWith('#')) return v.toLowerCase();
  return k === 'size' ? Number(v) : v;
};

/** 셀 서식이 찾을 서식의 모든 항목과 같은지 */
export function formatMatches(style, fmt) {
  if (!fmt) return true;
  const s = style ?? {};
  for (const k of Object.keys(fmt)) {
    if (!FIND_FORMAT_KEYS.includes(k)) continue;
    if (norm(k, s[k]) !== norm(k, fmt[k])) return false;
  }
  return true;
}

/** 시트의 내용이 있는 칸 위치 (셀 + 열 블록) — 순서 없음 */
function positions(sheet, fn) {
  sheet.cells.forEachRC((cell, r, c) => fn(r, c, cell));
  for (const b of sheet.blocks ?? []) {
    for (let j = 0; j < b.cols.length; j++) {
      const c = b.c0 + j;
      for (let i = 0; i < b.n; i++) {
        const r = b.r0 + i;
        if (!inBlock(b, r, c) || sheet.cells.getRC(r, c)) continue;
        fn(r, c, null);
      }
    }
  }
}

/**
 * 조건에 맞는 칸 목록 [{si, r, c}] — 시트 순서, 그 안에서 행(또는 열) 순서
 * opts: { text, matchCase, whole, lookIn: 'formulas'|'values'|'comments', byCols, sheets: [si], format }
 * io: { display(si, r, c) → 화면 글자, style(si, r, c) → 서식, limit }
 */
export function findMatches(wb, opts, io = {}) {
  const out = [];
  const limit = io.limit ?? 1e6;
  const text = opts.text ?? '';
  if (!text && !opts.format) return out;
  const re = text ? findRegex(text, opts) : null;
  const styleOf = io.style ?? ((si, r, c) => wb.styleAt(si, r, c));
  for (const si of opts.sheets ?? [0]) {
    const sheet = wb.sheets[si];
    if (!sheet) continue;
    const found = [];
    positions(sheet, (r, c, cell) => {
      if (found.length >= limit) return;
      if (re) {
        let hay;
        if (opts.lookIn === 'comments') hay = cell?.comment ?? '';
        else if (opts.lookIn === 'values') hay = io.display ? io.display(si, r, c) : String(wb.getValue(si, r, c) ?? '');
        else {
          hay = cell ? cell.raw ?? '' : wb.getRaw(si, r, c);
          if (hay.startsWith("'")) hay = hay.slice(1);
        }
        if (!hay || !re.test(hay)) return;
      } else if (!cell) return; // 서식만으로 찾을 때 서식 없는 블록 칸은 건너뜀 (블록 서식은 아래에서 비교)
      if (opts.format && !formatMatches(styleOf(si, r, c), opts.format)) return;
      found.push({ si, r, c });
    });
    found.sort(opts.byCols ? (a, b) => a.c - b.c || a.r - b.r : (a, b) => a.r - b.r || a.c - b.c);
    for (const f of found) out.push(f);
    if (out.length >= limit) break;
  }
  return out;
}

/** 현재 칸 다음(또는 이전)의 찾은 칸 — 끝에 닿으면 처음으로 */
export function nextMatch(list, cur, { byCols = false, back = false, order = null } = {}) {
  if (!list.length) return null;
  const sheetRank = (si) => (order ? order.indexOf(si) : si);
  const key = (m) => [sheetRank(m.si), byCols ? m.c : m.r, byCols ? m.r : m.c];
  const cmp = (a, b) => { for (let i = 0; i < 3; i++) if (a[i] !== b[i]) return a[i] - b[i]; return 0; };
  const k0 = key(cur);
  if (back) {
    for (let i = list.length - 1; i >= 0; i--) if (cmp(key(list[i]), k0) < 0) return list[i];
    return list[list.length - 1];
  }
  for (const m of list) if (cmp(key(m), k0) > 0) return m;
  return list[0];
}
