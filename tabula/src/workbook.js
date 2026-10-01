// 통합 문서 모델: 시트 · 셀 · 재계산 · 실행 취소 · 행/열 구조 변경
import { resolveStructRef, findTable } from './tables.js';
import { pivotSourceData, pivotLookup, resolvePivot } from './pivot.js';
import {
  parse, evaluateArray, evalAny, ERR, compareValues, isError, autoFormatFor, mayReturnArray, Range, RefValue,
  adjustFormulaForStructure, renameSheetInFormula, shiftFormula, quoteSheetName, MAX_ROWS, MAX_COLS, moveRefsInFormula,
} from './formula.js';
import { parseInput } from './format.js';
import { inBlock, blockValue, blockSet, blockClone, blockShift, rawOf, sortOrder, blockPermute, logicalCol, reorderRows, setRowOrder, materialize } from './block.js';
import { hid, shiftHidden } from './axis.js';
import { CellImage, compareSortValues } from './fxcore.js';
import { DepGraph, cellNum } from './depgraph.js';
import { CellMap } from './cellmap.js';
import { pushAll, CLOSED_BOOK, ERR_BY_CODE } from './fxcore.js';

export const DEFAULT_COL_WIDTH = 64;
export const DEFAULT_ROW_HEIGHT = 20;

const key = (r, c) => `${r},${c}`;
const unkey = (k) => { const i = k.indexOf(','); return [+k.slice(0, i), +k.slice(i + 1)]; };

// 같은 서식 객체를 여러 셀이 공유할 때(파일 가져오기 등) 한 번만 정리. 셀 서식은 바꿀 때 항상 새 객체로 교체하므로 공유해도 안전
const cleanMemo = new WeakMap();
function cleanStyle(style) {
  if (!style) return undefined;
  const hit = cleanMemo.get(style);
  if (hit !== undefined) return hit || undefined;
  let any = false;
  let holes = false;
  for (const k in style) { any = true; const v = style[k]; if (v === undefined || v === null) { holes = true; break; } }
  if (!any) return undefined;
  // 빈 값이 없으면 그대로 씀 (피벗을 그릴 때 백만 칸의 서식을 복사하지 않게)
  if (!holes) return style;
  const out = {};
  for (const [k, v] of Object.entries(style)) if (v !== undefined && v !== null) out[k] = v;
  const res = Object.keys(out).length ? out : undefined;
  cleanMemo.set(style, res ?? false);
  if (res) cleanMemo.set(res, res);
  return res;
}
const PLAIN_NUMBER = /^-?(?:0|[1-9]\d{0,14})(?:\.\d+)?$/;

// 같은 수식 문자열(표의 계산 열처럼 수십만 셀에 같은 수식)은 한 번만 해석해 AST 를 공유 (AST 는 바꾸지 않음)
const astMemo = new Map();
function parseMemo(text) {
  let e = astMemo.get(text);
  if (!e) {
    try {
      const ast = parse(text);
      e = { ast, maybeArray: mayReturnArray(ast) };
    } catch (err) {
      e = { ast: null, error: err.message };
    }
    if (astMemo.size > 100000) astMemo.clear();
    astMemo.set(text, e);
  }
  return e;
}

// 공유 수식 (엑셀과 같은 방식): 아래로 채운 수식(=H2/G2, =H3/G3 …)은 셀 위치 기준 상대 좌표가 같으므로
// AST 를 하나만 만들고, 셀마다 옮긴 거리(dr · dc)만 기억한다. 백만 행 수식도 해석은 한 번.
// 글자 하나가 이름의 일부인지 (영문 · 숫자 · _ · . · $ · 한글 등 비ASCII)
const wordish = (ch) => (ch >= 48 && ch <= 57) || (ch >= 65 && ch <= 90) || (ch >= 97 && ch <= 122) || ch === 95 || ch === 46 || ch === 36 || ch > 127;
const letter = (ch) => (ch >= 65 && ch <= 90) || (ch >= 97 && ch <= 122);
const digit = (ch) => ch >= 48 && ch <= 57;
/**
 * 수식 글자에서 A1 참조를 찾아 [글자, {cabs, ci, rabs, ri}, 글자, …] 로 나눔.
 * 열 전체(B:B) · 행 전체(1:1)는 { cols | rows, … } 조각. 글자열 · 작은따옴표 시트 이름 · 표 참조 [...] 안은 그대로 둠. 참조가 없으면 null
 */
function scanRefs(text) {
  const n = text.length;
  const segs = [];
  let last = 0;
  let refs = 0;
  let i = 0;
  while (i < n) {
    const ch = text.charCodeAt(i);
    if (ch === 34 || ch === 39) {
      i++;
      while (i < n) {
        if (text.charCodeAt(i) === ch) { if (text.charCodeAt(i + 1) === ch) { i += 2; continue; } i++; break; }
        i++;
      }
      continue;
    }
    if (ch === 91) {
      let depth = 1;
      i++;
      while (i < n && depth) { const x = text.charCodeAt(i); if (x === 91) depth++; else if (x === 93) depth--; i++; }
      continue;
    }
    const prev = i ? text.charCodeAt(i - 1) : 0;
    if (wordish(prev)) { i++; continue; }
    if (ch === 36 || letter(ch)) {
      let j = i;
      const cabs = ch === 36;
      if (cabs) j++;
      const ls = j;
      while (j < n && letter(text.charCodeAt(j))) j++;
      const nl = j - ls;
      if (nl >= 1 && nl <= 3) {
        const rabs = text.charCodeAt(j) === 36;
        if (rabs) j++;
        const ds = j;
        while (j < n && digit(text.charCodeAt(j))) j++;
        const nd = j - ds;
        const nx = j < n ? text.charCodeAt(j) : 0;
        if (nd >= 1 && nd <= 8 && !wordish(nx) && nx !== 40 && nx !== 33) {
          let ci = 0;
          for (let q = ls; q < ls + nl; q++) ci = ci * 26 + (text.charCodeAt(q) & 31);
          ci -= 1;
          const ri = +text.slice(ds, j) - 1;
          if (ci > 16383 || ri < 0) return null;
          segs.push(text.slice(last, i), { cabs, ci, rabs, ri });
          last = j;
          refs++;
          i = j;
          continue;
        }
        if (nd === 0 && !rabs && nx === 58) {
          // A:A · $A:$C (열 전체): 열만 옮겨짐
          let k = j + 1;
          const cabs2 = text.charCodeAt(k) === 36;
          if (cabs2) k++;
          const ls2 = k;
          while (k < n && letter(text.charCodeAt(k))) k++;
          const nx2 = k < n ? text.charCodeAt(k) : 0;
          if (k - ls2 < 1 || k - ls2 > 3 || wordish(nx2) || nx2 === 40 || nx2 === 33) return null;
          const colOf = (a, b) => { let v = 0; for (let q = a; q < b; q++) v = v * 26 + (text.charCodeAt(q) & 31); return v - 1; };
          const ci = colOf(ls, ls + nl);
          const ci2 = colOf(ls2, k);
          if (ci > 16383 || ci2 > 16383) return null;
          segs.push(text.slice(last, i), { cols: true, cabs, ci, cabs2, ci2 });
          last = k;
          refs++;
          i = k;
          continue;
        }
      }
      while (i < n && wordish(text.charCodeAt(i))) i++;
      continue;
    }
    if (digit(ch)) {
      let j = i;
      while (j < n && digit(text.charCodeAt(j))) j++;
      if (text.charCodeAt(j) === 58) {
        // 1:1 · 2:$5 (행 전체): 행만 옮겨짐
        let k = j + 1;
        const rabs2 = text.charCodeAt(k) === 36;
        if (rabs2) k++;
        const ds2 = k;
        while (k < n && digit(text.charCodeAt(k))) k++;
        if (k > ds2) {
          const nx2 = k < n ? text.charCodeAt(k) : 0;
          const ri = +text.slice(i, j) - 1;
          const ri2 = +text.slice(ds2, k) - 1;
          if (wordish(nx2) || nx2 === 40 || nx2 === 33 || ri < 0 || ri2 < 0 || j - i > 8 || k - ds2 > 8) return null;
          segs.push(text.slice(last, i), { rows: true, rabs: false, ri, rabs2, ri2 });
          last = k;
          refs++;
          i = k;
          continue;
        }
        if (rabs2) return null;
      }
      i = j;
      while (i < n && wordish(text.charCodeAt(i))) i++;
      continue;
    }
    i++;
  }
  if (!refs) return null;
  segs.push(text.slice(last));
  return segs;
}

/** 수식 글자 → 위치에 무관한 키 (A1 참조를 (r, c) 기준 상대 좌표로) */
function shareKey(text, r, c) {
  const segs = scanRefs(text);
  if (!segs) return null;
  let out = '';
  for (const x of segs) {
    if (typeof x === 'string') out += x;
    else if (x.cols) out += `\u0002${x.cabs ? `C${x.ci}` : `c${x.ci - c}`}:${x.cabs2 ? `C${x.ci2}` : `c${x.ci2 - c}`}`;
    else if (x.rows) out += `\u0003${x.rabs ? `R${x.ri}` : `r${x.ri - r}`}:${x.rabs2 ? `R${x.ri2}` : `r${x.ri2 - r}`}`;
    else out += `\u0001${x.cabs ? `C${x.ci}` : `c${x.ci - c}`}${x.rabs ? `R${x.ri}` : `r${x.ri - r}`}`;
  }
  return out;
}

const COL_NAMES = [];
const colName = (c) => {
  let s = COL_NAMES[c];
  if (s) return s;
  s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  if (c < 20000) COL_NAMES[c] = s;
  return s;
};
/**
 * 채우기 · 복사용: 수식을 한 번만 나눠 두고 (dr, dc) 만큼 옮긴 글자를 빠르게 만드는 함수 (shiftFormula 와 같은 결과).
 * 나눌 수 없는 수식은 shiftFormula 를 그대로 씀
 */
export function formulaShifter(raw) {
  const plain = (dr, dc) => shiftFormula(raw, dr, dc);
  if (!raw.startsWith('=')) return () => raw;
  const segs = scanRefs(raw.slice(1));
  if (!segs) return plain;
  const gen = (dr, dc) => {
    let out = '=';
    for (const x of segs) {
      if (typeof x === 'string') { out += x; continue; }
      if (x.cols) {
        const a = x.cabs ? x.ci : x.ci + dc;
        const b = x.cabs2 ? x.ci2 : x.ci2 + dc;
        if (a < 0 || b < 0 || a > 16383 || b > 16383) return null;
        out += `${x.cabs ? '$' : ''}${colName(a)}:${x.cabs2 ? '$' : ''}${colName(b)}`;
        continue;
      }
      if (x.rows) {
        const a = x.rabs ? x.ri : x.ri + dr;
        const b = x.rabs2 ? x.ri2 : x.ri2 + dr;
        if (a < 0 || b < 0 || a >= MAX_ROWS || b >= MAX_ROWS) return null;
        out += `${a + 1}:${x.rabs2 ? '$' : ''}${b + 1}`;
        continue;
      }
      const c = x.cabs ? x.ci : x.ci + dc;
      const r = x.rabs ? x.ri : x.ri + dr;
      if (c < 0 || r < 0 || c > 16383 || r >= MAX_ROWS) return null;
      out += `${x.cabs ? '$' : ''}${colName(c)}${x.rabs ? '$' : ''}${r + 1}`;
    }
    return out;
  };
  // 몇 가지 이동으로 shiftFormula 와 같은지 확인 (다르면 느린 길)
  for (const [dr, dc] of [[1, 0], [0, 1], [5, 2]]) {
    const a = gen(dr, dc);
    if (a !== null && a !== shiftFormula(raw, dr, dc)) return plain;
  }
  return (dr, dc) => gen(dr, dc) ?? shiftFormula(raw, dr, dc);
}
const sharedMemo = new Map();
/** (r, c) 칸의 수식 → { ast, maybeArray, error, r0, c0 } (r0 · c0: AST 를 만든 기준 위치) */
let lastAt = null; // 입력 한 번에 두 번 부르는 경우(자동 서식 + 셀 만들기) 재사용
function parseAt(text, r, c) {
  if (lastAt && lastAt.text === text && lastAt.r === r && lastAt.c === c) return lastAt.res;
  const res = parseAt0(text, r, c);
  lastAt = { text, r, c, res };
  return res;
}
// A1 참조가 없는 수식 글자 (표 수식 =[@a]*[@b] 처럼 행마다 글자가 같음): 다시 훑지 않음
const noRefText = new Set();
function parseAt0(text, r, c) {
  const k = r === undefined || noRefText.has(text) ? null : shareKey(text, r, c);
  if (k === null && r !== undefined) { if (noRefText.size > 100000) noRefText.clear(); noRefText.add(text); }
  if (k === null) { const p = parseMemo(text); return p.r0 === undefined ? { ...p, r0: r ?? 0, c0: c ?? 0, fixed: true } : p; }
  let e = sharedMemo.get(k);
  if (!e) {
    e = { ...parseMemo(text), r0: r, c0: c };
    if (sharedMemo.size > 200000) sharedMemo.clear();
    sharedMemo.set(k, e);
  }
  return e;
}

/** 수식이 참조하는 시트 · 표 · 이름 (시트 사이 의존 관계 계산용) */
const depMemo = new WeakMap();
function astDeps(ast) {
  let d = depMemo.get(ast);
  if (d) return d;
  d = { sheets: new Set(), tables: new Set(), names: new Set(), all: false };
  const walk = (n) => {
    if (!n || typeof n !== 'object') return;
    if (Array.isArray(n)) { for (const x of n) walk(x); return; }
    if (n.type === 'ref') {
      const s = n.ref?.sheet;
      if (s) { if (s.includes(':')) d.all = true; else d.sheets.add(s.toLowerCase()); }
    } else if (n.type === 'sref') {
      if (n.table) d.tables.add(n.table.toLowerCase());
    } else if (n.type === 'name') {
      d.names.add(n.v.toLowerCase());
      if (n.sheet) d.sheets.add(n.sheet.toLowerCase());
    } else if (n.type === 'func' && (n.name === 'INDIRECT' || n.name === 'EVALUATE')) {
      d.all = true;
    }
    for (const k in n) if (k !== 'ref') { const v = n[k]; if (v && typeof v === 'object') walk(v); }
  };
  walk(ast);
  depMemo.set(ast, d);
  return d;
}

/** 글자 값이 입력 해석으로 다른 값이 되지 않도록 raw 생성 */
function textRawOf(s) {
  if (s.startsWith('=') || s.startsWith("'")) return `'${s}`;
  const p = parseInput(s);
  return typeof p.value === 'string' && p.value === s && !p.errorLiteral ? s : `'${s}`;
}
/** 블록 값 → 셀 값 (오류는 오류 값) */
function blockCellValue(v) {
  if (v && typeof v === 'object') return ERR[Object.keys(ERR).find((k) => ERR[k].code === v.error)] ?? ERR.NA;
  return v;
}

/** 저장 형태 {raw, style, comment} → 계산용 셀 객체 */
export function makeCell(data, k = null) {
  if (k === null) return makeCellRC(data);
  const i = k.indexOf(',');
  return makeCellRC(data, +k.slice(0, i), +k.slice(i + 1));
}
/** 저장 형태 → 셀 객체 (r · c 를 알면 같은 모양의 수식끼리 AST 공유) */
export function makeCellRC(data, r, c) {
  if (!data) return null;
  const cell = { raw: data.raw ?? '' };
  const style = cleanStyle(data.style);
  if (style) cell.style = style;
  if (data.comment) cell.comment = data.comment;
  if (data.link) cell.link = data.link; // 하이퍼링크: 주소(URL) 또는 '#시트!A1'
  if (data.cached !== undefined) cell.cached = data.cached;
  if (data.image?.src) cell.image = { ...data.image }; // 셀에 배치한 그림
  // 텍스트 서식 칸의 수식: 새로 입력하면 글자지만, 파일에 수식으로 저장된 것(fx)은 엑셀처럼 수식
  if (data.fx) cell.fx = true;
  if (cell.raw.startsWith('=') && cell.raw.length > 1 && (style?.numFmt !== 'text' || data.fx)) {
    cell.formula = true;
    let p;
    if (r !== undefined) {
      p = parseAt(cell.raw.slice(1), r, c);
      if (!p.fixed && p.ast) {
        if (r !== p.r0) cell.dr = r - p.r0;
        if (c !== p.c0) cell.dc = c - p.c0;
      }
    } else p = parseMemo(cell.raw.slice(1));
    cell.ast = p.ast;
    if (p.ast) cell.maybeArray = p.maybeArray;
    else cell.parseError = p.error;
  } else if (cell.image && cell.raw === '') {
    cell.v = new CellImage(cell.image);
  } else if (style?.numFmt === 'text') {
    cell.v = cell.raw === '' ? null : cell.raw;
  } else {
    if (PLAIN_NUMBER.test(cell.raw)) cell.v = Number(cell.raw);
    else {
      const p = parseInput(cell.raw);
      cell.v = p.errorLiteral ? ERR_BY_CODE[p.errorLiteral] : p.value;
    }
  }
  if (!cell.raw && !cell.style && !cell.comment && !cell.link && !cell.image) return null;
  return cell;
}

// 값이 사실상 같음: 숫자는 유효 숫자 15자리까지 (피벗 합계의 마지막 자리 반올림 차이로 재계산하지 않게)
const sameValue = (a, b) => a === b
  || (typeof a === 'number' && typeof b === 'number' && a.toPrecision(15) === b.toPrecision(15))
  || (a !== null && b !== null && typeof a === 'object' && typeof b === 'object' && a.constructor === b.constructor && String(a) === String(b));

// 서식 비교: 같은 객체거나 키 · 값이 같으면 같음 (키 순서는 상관없음)
function sameStyle(a, b) {
  if (a === b) return true;
  if (!a || !b) return !a && !b;
  let n = 0;
  for (const k in a) {
    n++;
    const x = a[k];
    const y = b[k];
    if (x !== y && !(x && y && typeof x === 'object' && typeof y === 'object' && JSON.stringify(x) === JSON.stringify(y))) return false;
  }
  for (const k in b) n--; // eslint-disable-line no-unused-vars
  return n === 0;
}
const savedCached = (c) => (c.cached !== undefined && c.formula && !c.dirty ? c.cached : undefined);
/** 두 셀의 저장 형태(cellData)가 같은지 — cellData 를 만들어 JSON 으로 비교하는 것과 같지만 훨씬 빠름 */
function sameCell(a, b) {
  if (!a || !b) return !a && !b;
  if (a.raw !== b.raw || (a.comment || null) !== (b.comment || null) || (a.link || null) !== (b.link || null) || !!a.fx !== !!b.fx) return false;
  if (!sameStyle(a.style, b.style)) return false;
  if ((a.image || b.image) && JSON.stringify(a.image ?? null) !== JSON.stringify(b.image ?? null)) return false;
  const ca = savedCached(a);
  const cb = savedCached(b);
  return ca === cb || JSON.stringify(ca) === JSON.stringify(cb);
}

export function cellData(cell) {
  if (!cell) return null;
  const d = { raw: cell.raw };
  if (cell.style) d.style = { ...cell.style };
  if (cell.comment) d.comment = cell.comment;
  if (cell.link) d.link = cell.link;
  if (cell.image) d.image = { ...cell.image };
  if (cell.fx) d.fx = true;
  // 입력이 바뀌어 다시 계산한 수식은 파일의 옛 계산 결과를 버림
  if (cell.cached !== undefined && cell.formula && !cell.dirty) d.cached = cell.cached;
  return d;
}

/** 지원하지 않는 함수 수식의 파일 속 계산 결과 (저장 가능한 형태) */
function cachedValue(c) {
  if (c && typeof c === 'object' && c.error) return ERR[Object.keys(ERR).find((k) => ERR[k].code === c.error)] ?? ERR.NAME;
  return c;
}

function newSheet(name) {
  return {
    name, cells: new CellMap(), colWidths: {}, rowHeights: {}, merges: [], cond: [],
    colStyles: {}, rowStyles: {}, allStyle: null, hiddenRows: {}, hiddenCols: {}, rowManual: {},
    freeze: { rows: 0, cols: 0 }, filter: null, charts: [], pivot: null,
    validations: [], images: [], shapes: [], tables: [], slicers: [], pivotsExtra: [], blocks: [],
  };
}

/** 저장 형태의 시트 하나 → 시트 객체 (셀은 객체 또는 Map) */
function sheetFromData(s) {
  const sheet = newSheet(s.name);
  const entries = s.cells instanceof Map || s.cells instanceof CellMap ? s.cells : Object.entries(s.cells || {});
  for (const [k, d] of entries) {
    const cell = makeCell(d, k);
    if (cell) sheet.cells.set(k, cell);
  }
  for (const p of SHEET_PROPS) if (s[p] !== undefined && s[p] !== null) sheet[p] = structuredClone(s[p]);
  sheet.freeze = { rows: 0, cols: 0, ...(s.freeze || {}) };
  if (s._sid) sheet._sid = s._sid;
  sheet.blocks = (s.blocks ?? []).map(blockClone);
  return sheet;
}

/** 시트의 부가 속성 (셀 외) — 저장/복원/복제용 */
const SHEET_PROPS = ['colWidths', 'rowHeights', 'merges', 'cond', 'colStyles', 'rowStyles', 'allStyle',
  'hiddenRows', 'hiddenCols', 'rowManual', 'freeze', 'filter', 'charts', 'pivot', 'validations', 'images', 'shapes', 'tables', 'slicers', 'pivotsExtra', 'state', 'noGrid', 'noZeros', 'outline', 'protect', 'sparklines', 'page', 'defRowH', 'defColW', 'zoom', 'view', 'tabColor', 'scenarios', 'external'];
// 바뀌어도 수식 결과가 달라지지 않는 시트 속성
const CALC_NEUTRAL = new Set(['external', 'scenarios', 'tabColor', 'defRowH', 'defColW', 'zoom', 'view', 'outline', 'protect', 'sparklines', 'page', 'state', 'noGrid', 'noZeros', 'charts', 'images', 'shapes', 'slicers', 'freeze', 'cond', 'validations', 'colStyles', 'rowStyles', 'allStyle', 'merges']);

/** 숫자 키 객체의 키를 삽입/삭제에 맞춰 이동 */
function shiftKeys(obj, index, count) {
  const out = {};
  for (const [k, v] of Object.entries(obj ?? {})) {
    const p = Number(k);
    if (count < 0 && p >= index && p < index - count) continue;
    out[p >= index ? p + count : p] = v;
  }
  return out;
}

/** 한 번의 변경으로 다시 계산할 수식이 이보다 많으면 칸마다 표시하지 않고 시트 단위로 다시 계산 (보이는 칸만 그때 계산) */
const DIRTY_LIMIT = 250000;
const DEEP = new Error('수식 체인이 너무 깊습니다');
const MAX_DEPTH = 300;
const EMPTY_STYLE = Object.freeze({});

/** 행/열 삽입·삭제에 맞춰 범위 {r1,c1,r2,c2} 조정. 완전히 삭제되면 null */
export function adjustRange(rg, axis, index, count) {
  const [k1, k2] = axis === 'row' ? ['r1', 'r2'] : ['c1', 'c2'];
  let a = rg[k1];
  let b = rg[k2];
  if (count > 0) {
    if (a >= index) a += count;
    if (b >= index) b += count;
  } else {
    const end = index - count;
    if (a >= index && b < end) return null;
    a = a >= end ? a + count : a >= index ? index : a;
    b = b >= end ? b + count : b >= index ? index - 1 : b;
  }
  return { ...rg, [k1]: a, [k2]: b };
}

export class Workbook {
  constructor(data) {
    this.caches = []; // 시트별 수식 결과 캐시: Map('r,c' → 값)
    this.deps = null; // 시트별 참조 시트 (지연 계산)
    this.affectMemo = new Map();
    this.arrayList = null; // 분산될 수 있는 수식 셀 목록 (지연 계산)
    this.evaluating = new Set();
    this.usedCache = new Map();
    this.extentCache = new Map();
    this.undoStack = [];
    this.redoStack = [];
    this.tx = null;
    this.listeners = new Set();
    this.depth = 0;
    this.warming = false;
    this.names = [];
    this.spills = new Map();
    this.spillOwner = new Map();
    this.spillState = null;
    this.nameStack = new Set();
    this.version = 0;
    this.graph = null; // 셀 단위 의존 그래프 (처음 필요할 때 만듦)
    this.pending = []; // 트랜잭션 중 바뀐 칸 [시트, 행, 열, …] (읽기 전이나 끝날 때 한 번에 반영)
    this.ctxs = []; // 시트별 계산 문맥 (셀마다 새로 만들지 않음)
    this.colIdx = []; // 시트별 열 → 셀(Map) 행 목록 (큰 범위 빠르게 읽기)
    if (data) this.load(data);
    else this.sheets = [newSheet('Sheet1')];
  }

  // ─────────── 조회 ───────────
  onChange(fn) { this.listeners.add(fn); }
  emit() { for (const fn of this.listeners) fn(); }

  sheetIndexByName(name) {
    // 이름 → 번호 기억 (수식 백만 개의 시트 참조도 빠르게). 시트를 옮기거나 이름을 바꿨으면 확인에서 걸러짐
    const memo = (this.sheetNameMemo ??= new Map());
    const hit = memo.get(name);
    if (hit && this.sheets[hit.i] === hit.sheet && hit.sheet.name === hit.at) return hit.i;
    const n = name.toLowerCase();
    const i = this.sheets.findIndex((s) => s.name.toLowerCase() === n);
    if (i >= 0) memo.set(name, { i, sheet: this.sheets[i], at: this.sheets[i].name });
    else memo.delete(name);
    return i;
  }

  getCell(si, r, c) {
    const sheet = this.sheets[si];
    if (!sheet) return undefined;
    const cell = sheet.cells.getRC(r, c);
    if (cell || !sheet.blocks.length) return cell;
    return this.blockCell(sheet, r, c);
  }

  /** (r, c) 가 들어 있는 열 블록 */
  blockAt(si, r, c) {
    const bl = this.sheets[si]?.blocks;
    if (!bl?.length) return null;
    for (const b of bl) if (inBlock(b, r, c)) return b;
    return null;
  }

  /** 블록 칸 → 가벼운 셀 객체 (그때그때 만듦) */
  blockCell(sheet, r, c) {
    for (const b of sheet.blocks) {
      if (!inBlock(b, r, c)) continue;
      const v = blockValue(b, r, c);
      const fmt = b.cols[c - b.c0].fmt ?? undefined;
      if (v === null) return fmt ? { raw: '', v: null, style: fmt } : undefined;
      return { raw: rawOf(v, fmt, textRawOf), v: blockCellValue(v), style: fmt, block: true };
    }
    return undefined;
  }

  getRaw(si, r, c) { return this.getCell(si, r, c)?.raw ?? ''; }

  getValue(si, r, c) {
    if (this.pending.length) this.flushPending();
    const sheet = this.sheets[si];
    const cell = sheet?.cells.getRC(r, c);
    if (!cell && sheet?.blocks.length) {
      for (const b of sheet.blocks) {
        if (!inBlock(b, r, c)) continue;
        const v = blockValue(b, r, c);
        if (v !== null) return blockCellValue(v);
        break;
      }
    }
    if (!cell || (!cell.formula && cell.raw === '' && !cell.image)) return this.spillValueAt(si, r, c);
    if (!cell.formula) return cell.v ?? null;
    const cache = this.caches[si] ??= new CellMap();
    const hit = cache.getRC(r, c);
    if (hit !== undefined || cache.hasRC(r, c)) return hit;
    // 파일에서 연 뒤 아직 바뀐 것이 없으면 엑셀이 저장해 둔 계산 결과를 그대로 사용
    if (sheet.fileValues && cell.cached !== undefined && !cell.maybeArray && !cell.dirty) return cachedValue(cell.cached);
    const k = cellNum(si, r, c);
    // 읽을 수 없는 수식(외부 통합 문서의 표 참조 등): 파일에 저장된 값 유지
    if (cell.ast === null) return cell.cached !== undefined ? cachedValue(cell.cached) : ERR.NAME;
    // 순환 참조: 파일에 저장된 마지막 값이 있으면 그 값 (엑셀도 반복 계산을 끈 순환 참조는 마지막 값을 유지)
    if (this.evaluating.has(k)) {
      if (cell.cached === undefined) return ERR.CIRC;
      // 고리에 든 칸(계산 중인 칸 중 k 부터 위쪽 전부)은 모두 파일의 마지막 값을 유지 (엑셀은 순환 참조를 계산하지 않음)
      const hits = (this.circHits ??= new Set());
      let on = false;
      for (const x of this.evaluating) { if (x === k) on = true; if (on) hits.add(x); }
      return cachedValue(cell.cached);
    }
    if (this.depth > 0 || this.warming) return this.evalCell(k, cell, si, r, c);
    // 최상위 호출: 긴 참조 사슬(예: 누계 1만 행)은 위에서부터 차례로 미리 계산한 뒤 다시 시도
    try {
      return this.evalCell(k, cell, si, r, c);
    } catch (e) {
      if (e !== DEEP) throw e;
      this.warmup();
      try {
        return this.evalCell(k, cell, si, r, c);
      } catch (e2) {
        if (e2 !== DEEP) throw e2;
        return ERR.CIRC;
      }
    }
  }

  evalCell(k, cell, si, r, c) {
    if (this.depth >= MAX_DEPTH) throw DEEP;
    this.evaluating.add(k);
    this.depth++;
    let v;
    // 시트 문맥 하나를 재사용: 위치(here)와 공유 수식 이동 거리(dr · dc)만 바꾸고 끝나면 되돌림
    const ctx = this.ctxs[si] ?? this.cellCtx(si);
    const ph = ctx.here;
    const pdr = ctx.dr;
    const pdc = ctx.dc;
    ctx.here = { si, r, c, sheet: this.sheets[si]?.name };
    ctx.dr = cell.dr ?? 0;
    ctx.dc = cell.dc ?? 0;
    const closedPrev = CLOSED_BOOK.hit;
    CLOSED_BOOK.hit = false;
    let closed = false;
    try {
      v = evaluateArray(cell.ast, ctx);
      closed = CLOSED_BOOK.hit;
    } catch (e) {
      if (e instanceof RangeError) throw DEEP;
      throw e;
    } finally {
      ctx.here = ph;
      ctx.dr = pdr;
      ctx.dc = pdc;
      CLOSED_BOOK.hit = closedPrev;
      this.depth--;
      this.evaluating.delete(k);
    }
    // 지원하지 않는 함수는 파일에 저장된 계산 결과를 그대로 표시
    if (v === ERR.NAME && cell.cached !== undefined) v = cachedValue(cell.cached);
    // 닫힌 외부 통합 문서를 *IF(S) 로 읽는 수식: 엑셀처럼 파일에 저장된 값 유지 (없으면 #VALUE!)
    if (closed && cell.cached !== undefined) v = cachedValue(cell.cached);
    // 순환이 이 칸에서 닫힘: 엑셀처럼 계산하지 않고 마지막 값 유지 (=G24/J24 를 J24 에)
    if (this.circHits?.has(k)) { this.circHits.delete(k); if (cell.cached !== undefined) v = cachedValue(cell.cached); }
    if (v instanceof Range) v = this.placeSpill(`${si}:${r},${c}`, si, r, c, v);
    (this.caches[si] ??= new CellMap()).setRC(r, c, v);
    return v;
  }

  // ─────────── 동적 배열 (분산) ───────────
  /** 배열 결과를 주변 빈 셀로 분산. 막혀 있으면 #SPILL! */
  placeSpill(k, si, r, c, arr) {
    const h = arr.height;
    const w = arr.width;
    if (r + h > MAX_ROWS || c + w > MAX_COLS) return ERR.SPILL;
    const sheet = this.sheets[si];
    for (let i = 0; i < h; i++) {
      for (let j = 0; j < w; j++) {
        if (!i && !j) continue;
        if (sheet.cells.getRC(r + i, c + j)?.raw) return ERR.SPILL;
        const owner = this.spillOwner.get(`${si}:${r + i},${c + j}`);
        if (owner && owner !== k) return ERR.SPILL;
      }
    }
    if (sheet.merges.some((m) => m.r1 <= r + h - 1 && m.r2 >= r && m.c1 <= c + w - 1 && m.c2 >= c)) return ERR.SPILL;
    this.spills.set(k, { si, r, c, h, w, rows: arr.rows });
    for (let i = 0; i < h; i++) {
      for (let j = 0; j < w; j++) if (i || j) this.spillOwner.set(`${si}:${r + i},${c + j}`, k);
    }
    const v = arr.rows[0][0];
    return v === null || v === undefined ? 0 : v;
  }

  /** 분산될 수 있는 수식을 모두 계산해 분산 영역을 확정 */
  /**
   * 분산 후보 수식을 모두 계산 (빈 칸에 보일 분산 값을 알기 위해).
   * 계산 도중 다시 불리면(수식이 범위를 읽음) 그 범위(area)로 분산될 수 있는 남은 수식만 계산 —
   * 남은 대기열 전체를 겹겹이 계산하면 후보가 수천 개일 때 끝나지 않음
   */
  ensureSpills(area = null) {
    if (this.spillState === 'done') return;
    if (this.spillState === 'running') {
      if (!area) { this.runSpillQueue(); return; }
      const [s, , , r2, c2] = area;
      const q = this.spillQueue;
      for (let i = this.spillPos; i < q.length; i++) {
        const [si, r, c] = q[i];
        // 지금 계산 중인 수식(자기 원본 범위를 읽는 INDIRECT 등)은 건너뜀 — 다시 부르면 순환 참조로 잘못 봄
        if (si === s && r <= r2 && c <= c2 && !this.evaluating.has(cellNum(si, r, c))) this.getValue(si, r, c);
      }
      return;
    }
    this.spillState = 'running';
    if (!this.arrayList) {
      const list = [];
      this.sheets.forEach((sheet, si) => {
        sheet.cells.forEachRC((cell, r, c) => { if (cell.formula && cell.maybeArray) list.push([si, r, c]); });
      });
      list.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]);
      this.arrayList = list;
    }
    this.spillQueue = this.arrayList;
    this.spillPos = 0;
    try {
      this.runSpillQueue();
    } finally {
      this.spillState = this.spillPos >= this.spillQueue.length ? 'done' : null;
    }
  }

  runSpillQueue() {
    while (this.spillPos < this.spillQueue.length) {
      const [si, r, c] = this.spillQueue[this.spillPos++];
      if (!this.evaluating.has(cellNum(si, r, c))) this.getValue(si, r, c);
    }
  }

  spillValueAt(si, r, c) {
    this.ensureSpills([si, r, c, r, c]);
    if (!this.spillOwner.size) return null;
    const owner = this.spillOwner.get(`${si}:${r},${c}`);
    if (!owner) return null;
    const sp = this.spills.get(owner);
    const v = sp.rows[r - sp.r]?.[c - sp.c];
    return v === null || v === undefined ? 0 : v;
  }

  /** 분산된 셀이면 원본(앵커) 셀 {r, c} */
  spillAnchorOf(si, r, c) {
    this.ensureSpills([si, r, c, r, c]);
    const owner = this.spillOwner.get(`${si}:${r},${c}`);
    if (!owner) return null;
    const sp = this.spills.get(owner);
    return { r: sp.r, c: sp.c };
  }

  /** 앵커 셀의 분산 영역 {r1,c1,r2,c2} (없으면 null) */
  spillRange(si, r, c) {
    const cell = this.getCell(si, r, c);
    if (!cell?.formula) return null;
    this.getValue(si, r, c);
    const sp = this.spills.get(`${si}:${r},${c}`);
    return sp ? { r1: sp.r, c1: sp.c, r2: sp.r + sp.h - 1, c2: sp.c + sp.w - 1 } : null;
  }

  /** 시트의 분산 영역 목록 (저장용) */
  spillsOf(si) {
    this.ensureSpills([si, 0, 0, Infinity, Infinity]);
    return [...this.spills.values()].filter((sp) => sp.si === si);
  }

  // ─────────── 이름 정의 ───────────
  /** 이름 찾기: sheet 가 있으면 그 시트 범위 이름, 없으면 현재 시트 범위 → 통합 문서 범위 */
  findName(name, hostSi = null, sheet = null) {
    const n = name.toLowerCase();
    const list = this.names.filter((x) => x.name.toLowerCase() === n);
    if (sheet) return list.find((x) => x.sheet && x.sheet.toLowerCase() === sheet.toLowerCase()) ?? null;
    const hostName = hostSi === null ? null : this.sheets[hostSi]?.name.toLowerCase();
    return list.find((x) => x.sheet && x.sheet.toLowerCase() === hostName) ?? list.find((x) => !x.sheet) ?? null;
  }

  /** 이름 값 (RefValue · 값 · LAMBDA). 없으면 undefined */
  nameValue(name, hostSi, sheet, here) {
    const entry = this.findName(name, hostSi, sheet);
    if (!entry) {
      // 표 이름만 쓰면 데이터 영역
      if (!sheet && findTable(this, name)) {
        const rg = resolveStructRef(this, name, '', null);
        if (rg) return new RefValue(this.sheets[rg.si].name, rg.r1, rg.c1, rg.r2, rg.c2);
      }
      return undefined;
    }
    if (this.nameStack.has(entry)) return ERR.CIRC;
    const text = String(entry.ref ?? '').replace(/^=/, '');
    if (entry._text !== text) {
      entry._text = text;
      try { entry._ast = parse(text); } catch { entry._ast = null; }
    }
    if (!entry._ast) return ERR.NAME;
    const scopeSi = entry.sheet ? this.sheetIndexByName(entry.sheet) : hostSi ?? 0;
    const ctx = this.ctxFor(scopeSi < 0 ? hostSi ?? 0 : scopeSi, here?.r ?? null, here?.c ?? null);
    this.nameStack.add(entry);
    try {
      const v = evalAny(entry._ast, ctx);
      // 시트가 없는 참조는 이름의 시트(또는 수식 시트) 기준으로 고정
      if (v instanceof RefValue && !v.sheet) return new RefValue(this.sheets[scopeSi < 0 ? hostSi ?? 0 : scopeSi].name, v.r1, v.c1, v.r2, v.c2);
      return v;
    } catch (e) {
      if (isError(e)) return e;
      throw e;
    } finally {
      this.nameStack.delete(entry);
    }
  }

  /** 이름 목록 교체 (실행 취소 가능) */
  setNames(list) {
    this.snapshotNames();
    const before = this.sheetDeps();
    this.names = list.map((x) => ({ name: x.name, ref: x.ref, sheet: x.sheet ?? null, ...(x.comment ? { comment: x.comment } : {}), ...(x.hidden ? { hidden: true } : {}) }));
    this.deps = null;
    const after = this.sheetDeps();
    // 이름을 쓰는 시트만 다시 계산
    this.sheets.forEach((sh, i) => { if (before[i]?.all || after[i]?.all) sh.fileValues = false; });
    this.invalidateStructure();
  }

  /** 모든 수식을 행 순서대로 계산해 캐시를 채움 */
  warmup() {
    this.warming = true;
    try {
      const list = [];
      this.sheets.forEach((sheet, si) => {
        sheet.cells.forEachRC((cell, r, c) => { if (cell.formula) list.push([si, r, c]); });
      });
      list.sort((a, b) => a[0] - b[0] || a[1] - b[1] || a[2] - b[2]);
      let pending = list;
      for (let pass = 0; pass < 50 && pending.length; pass++) {
        const next = [];
        for (const [si, r, c] of pending) {
          try { this.getValue(si, r, c); } catch (e) { if (e !== DEEP) throw e; next.push([si, r, c]); }
        }
        if (next.length === pending.length) break;
        pending = next;
      }
    } finally {
      this.warming = false;
      this.depth = 0;
    }
  }

  /** 열/행/시트 전체 서식을 합친 실제 셀 서식 */
  styleAt(si, r, c) {
    const s = this.sheets[si];
    let own = s.cells.getRC(r, c)?.style;
    if (!own && s.blocks.length) { const b = this.blockAt(si, r, c); if (b) own = b.cols[c - b.c0].fmt ?? undefined; }
    const col = s.colStyles[c];
    const row = s.rowStyles[r];
    if (!s.allStyle && !col && !row) return own ?? this.baseStyle ?? EMPTY_STYLE;
    return { ...(own ? null : this.baseStyle), ...s.allStyle, ...col, ...row, ...own };
  }

  hasLineStyle(si, r, c) {
    const s = this.sheets[si];
    return !!(s.allStyle || s.colStyles[c] || s.rowStyles[r]);
  }

  resolveSheet(name, si) {
    if (name == null) return si;
    const i = this.sheetIndexByName(name);
    if (i < 0) throw ERR.REF;
    return i;
  }

  /** 셀 계산용 시트 문맥 (재사용) */
  cellCtx(si) {
    return (this.ctxs[si] ??= this.ctxFor(si));
  }

  ctxFor(si, r = null, c = null) {
    const ctx = { here: r === null ? null : { si, r, c, sheet: this.sheets[si]?.name }, dr: 0, dc: 0 };
    return Object.assign(ctx, {
      name: (n, sheet) => this.nameValue(n, si, sheet, ctx.here),
      spillRef: (sheet, rr, cc) => {
        const s = this.resolveSheet(sheet, si);
        const v = this.getValue(s, rr, cc);
        if (v === ERR.SPILL) throw ERR.REF;
        const sp = this.spills.get(`${s}:${rr},${cc}`);
        const name = this.sheets[s].name;
        if (sp) return new RefValue(name, sp.r, sp.c, sp.r + sp.h - 1, sp.c + sp.w - 1);
        if (!this.getCell(s, rr, cc)?.formula) return null;
        return new RefValue(name, rr, cc);
      },
      formulaText: (sheet, rr, cc) => {
        const cell = this.getCell(this.resolveSheet(sheet, si), rr, cc);
        return cell?.formula ? cell.raw : null;
      },
      sheetIndex: (name) => (name == null ? si : this.sheetIndexByName(name)),
      sheetCount: () => this.sheets.length,
      pivotData: (ref, field, items) => {
        // GETPIVOTDATA: 참조가 들어 있는 피벗 테이블에서 값 찾기
        // 같은 계산 상태(version)에서는 참조 → 피벗 · 피벗 결과를 한 번만 구함 (GETPIVOTDATA 수만 개가 매번 찾지 않게)
        const memo = (this.pivotMemo ??= { v: -1, m: new Map(), at: new Map() });
        if (memo.v !== this.version) { memo.v = this.version; memo.m.clear(); memo.at.clear(); }
        const atKey = `${ref.sheet ?? si}\u0001${ref.r1}\u0001${ref.c1}`;
        let def = memo.at.get(atKey);
        if (def === undefined) {
          const sh = this.sheets[this.resolveSheet(ref.sheet, si)];
          const defs = [sh.pivot, ...(sh.pivotsExtra ?? [])].filter(Boolean);
          def = defs.find((d) => d.area && ref.r1 >= d.area.r1 && ref.r1 <= d.area.r2 && ref.c1 >= d.area.c1 && ref.c1 <= d.area.c2)
            ?? (defs.length && !defs[0].area ? defs[0] : null);
          memo.at.set(atKey, def);
        }
        if (!def) throw ERR.REF;
        let hit = memo.m.get(def);
        if (!hit) {
          const src = pivotSourceData(this, def);
          hit = { src, res: src ? resolvePivot(src, def) : null };
          memo.m.set(def, hit);
        }
        const v = hit.src ? pivotLookup(hit.src, def, field, items, hit.res) : null;
        // 없는 항목은 #REF! 값으로 돌려줌 (IFERROR 로 감싼 수만 개가 예외를 던지면 느림)
        return v === null || v === undefined ? ERR.REF : v;
      },
      quoteSheet: (name) => quoteSheetName(name),
      colWidthChars: (sheet, cc) => Math.round(this.colWidth(this.resolveSheet(sheet, si), cc) / 7.5),
      cellFormat: (sheet, rr, cc) => {
        const st = this.styleAt(this.resolveSheet(sheet, si), rr, cc);
        return { number: 'F2', currency: 'C0', accounting: 'C0', percent: 'P0', date: 'D1', longdate: 'D1', time: 'D9', datetime: 'D4', scientific: 'S2', text: '@' }[st.numFmt] ?? 'G';
      },
      usedCols: (sheet) => this.usedRange(this.resolveSheet(sheet, si)).cols,
      structRef: (table, spec) => {
        const rg = resolveStructRef(this, table, spec, ctx.here);
        // 표의 데이터 행 밖에서 [#This Row]: 엑셀은 #VALUE! (암시적 교차 실패)
        if (!rg) return /^\s*(@|\[?#this row|\[?#현재 행)/i.test(String(spec ?? '')) ? { error: ERR.VALUE } : null;
        const range = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
        return { sheet: this.sheets[rg.si].name, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, range };
      },
      rowHidden: (sheet, row, manualToo, filteredToo = true) => {
        const s = this.sheets[this.resolveSheet(sheet, si)];
        if (!s) return false;
        if (manualToo && s.hiddenRows?.[row]) return true;
        if (!filteredToo) return false;
        if (hid(s.filter?.hidden, row)) return true;
        return (s.tables ?? []).some((t) => hid(t.filter?.hidden, row));
      },
      cell: (sheet, r, c) => this.getValue(this.resolveSheet(sheet, si), r, c),
      range: (sheet, r1, c1, r2, c2) => {
        const s = this.resolveSheet(sheet, si);
        // 256칸이 넘으면 캐시된 읽기 (같은 범위를 읽는 수식 수만 개가 SUMIFS 색인 · 배열을 같이 씀)
        if ((r2 - r1 + 1) * (c2 - c1 + 1) > 256 && this.sheets[s]) return this.rangeFast(s, r1, c1, r2, c2);
        const rows = [];
        for (let r = r1; r <= r2; r++) {
          const row = [];
          for (let c = c1; c <= c2; c++) row.push(this.getValue(s, r, c));
          rows.push(row);
        }
        return rows;
      },
      usedRows: (sheet) => {
        const s = this.resolveSheet(sheet, si);
        let n = this.usedRange(s).rows;
        for (const sp of this.spills.values()) if (sp.si === s) n = Math.max(n, sp.r + sp.h);
        return n;
      },
    });
  }

  /**
   * 큰 범위 값 읽기 (SUM(K2:K1000001) · SUMIFS · XLOOKUP 등): 칸마다 찾지 않고
   * 열 블록은 타입 배열에서 바로, 셀(Map)은 열별 행 목록으로 있는 칸만 읽음
   */
  rangeFast(s, r1, c1, r2, c2) {
    // 같은 큰 범위를 여러 수식이 읽으면(SUMIFS 요약표 등) 시트가 바뀌기 전까지 한 번 읽은 값을 같이 씀
    if (this.pending.length) this.flushPending();
    const mk = `${s}:${r1}:${c1}:${r2}:${c2}`;
    const ver = this.rangeVersion(s, c1, c2);
    const memo = (this.rangeMemo ??= new Map());
    const hit = memo.get(mk);
    if (hit && hit.ver === ver) return hit.rows;
    const rows = this.rangeRead(s, r1, c1, r2, c2);
    if (this.rangeVersion(s, c1, c2) === ver) {
      if (memo.size >= 48) memo.delete(memo.keys().next().value);
      memo.set(mk, { ver, rows });
    }
    return rows;
  }

  /** 범위 캐시 버전: 시트 구조 · 시트 단위 재계산 버전 + 열마다 바뀐 횟수 (다른 열을 고쳐도 유지) */
  rangeVersion(s, c1, c2) {
    let v = `${this.sheetVerAll ?? 0}:${this.baseVer?.[s] ?? 0}`;
    if (c2 - c1 > 64) return `${v}:${this.sheetVer?.[s] ?? 0}`;
    const cv = this.colVer?.[s];
    if (cv) for (let c = c1; c <= c2; c++) v += `:${cv.get(c) ?? 0}`;
    return v;
  }

  /** 열 버전 올림 (칸이 바뀌거나 다시 계산될 수식이 있는 열) */
  bumpCol(s, c) {
    const all = (this.colVer ??= []);
    const m = (all[s] ??= new Map());
    m.set(c, (m.get(c) ?? 0) + 1);
  }

  rangeRead(s, r1, c1, r2, c2) {
    const sh = this.sheets[s];
    const h = r2 - r1 + 1;
    const w = c2 - c1 + 1;
    const rows = new Array(h);
    if (w === 1) for (let i = 0; i < h; i++) rows[i] = [null];
    else for (let i = 0; i < h; i++) { const row = new Array(w); for (let j = 0; j < w; j++) row[j] = null; rows[i] = row; }
    for (const b of sh.blocks) {
      const br1 = Math.max(r1, b.r0);
      const br2 = Math.min(r2, b.r0 + b.n - 1);
      const bc1 = Math.max(c1, b.c0);
      const bc2 = Math.min(c2, b.c0 + b.cols.length - 1);
      if (br1 > br2 || bc1 > bc2) continue;
      const perm = b.perm;
      for (let c = bc1; c <= bc2; c++) {
        const { num, str, dict } = b.cols[c - b.c0];
        const j = c - c1;
        for (let r = br1; r <= br2; r++) {
          const i = perm ? perm[r - b.r0] : r - b.r0;
          let v = null;
          if (str) { const x = str[i]; if (x >= 0) v = dict[x]; }
          if (v === null && num) { const x = num[i]; if (x === x) v = x; }
          if (v !== null) rows[r - r1][j] = typeof v === 'object' ? blockCellValue(v) : v;
        }
      }
    }
    // 셀(Map): 열마다 정렬된 행 목록에서 범위 안의 칸만
    if (this.pending.length) this.flushPending();
    for (let c = c1; c <= c2; c++) {
      const list = this.colRows(s, c);
      if (!list) continue;
      const colMap = sh.cells.col(c);
      const cacheCol = this.caches[s]?.col(c);
      let lo = 0;
      let hi = list.length;
      while (lo < hi) { const m = (lo + hi) >> 1; if (list[m] < r1) lo = m + 1; else hi = m; }
      for (let i = lo; i < list.length && list[i] <= r2; i++) {
        const r = list[i];
        const cell = colMap.get(r);
        let v;
        // 값 셀 · 계산해 둔 수식은 바로, 나머지는 getValue
        if (cell && !cell.formula && (cell.raw !== '' || cell.image)) v = cell.v ?? null;
        else if (cell?.formula && cacheCol !== undefined && (v = cacheCol.get(r)) !== undefined) { /* 캐시 */ } else v = this.getValue(s, r, c);
        rows[r - r1][c - c1] = v;
      }
    }
    // 분산 영역 (빈 칸에 표시되는 동적 배열 값)
    this.ensureSpills([s, r1, c1, r2, c2]);
    for (const sp of this.spills.values()) {
      if (sp.si !== s || sp.r > r2 || sp.r + sp.h - 1 < r1 || sp.c > c2 || sp.c + sp.w - 1 < c1) continue;
      for (let r = Math.max(r1, sp.r); r <= Math.min(r2, sp.r + sp.h - 1); r++) {
        for (let c = Math.max(c1, sp.c); c <= Math.min(c2, sp.c + sp.w - 1); c++) {
          if ((r === sp.r && c === sp.c) || rows[r - r1][c - c1] !== null || sh.cells.getRC(r, c)?.raw) continue;
          const v = sp.rows[r - sp.r]?.[c - sp.c];
          rows[r - r1][c - c1] = v === null || v === undefined ? 0 : v;
        }
      }
    }
    return rows;
  }

  /** 시트 s 의 열 c 에 셀이 있는 행 번호 (정렬됨, 열마다 처음 필요할 때 만들고 putCell 이 고쳐 나감) */
  colRows(s, c) {
    const sh = this.sheets[s];
    let cur = this.colIdx[s];
    if (!cur || cur.cells !== sh.cells) { cur = { cells: sh.cells, map: new Map() }; this.colIdx[s] = cur; }
    let l = cur.map.get(c);
    if (l === undefined) {
      const m = sh.cells.col(c);
      l = m ? [...m.keys()].sort((a, b) => a - b) : null;
      cur.map.set(c, l);
    }
    return l;
  }

  /** 사용 범위 캐시를 다시 훑지 않고 고침: 커지면 넓히고, 끝 칸이 비면 다음에 다시 계산 */
  boundsUpdate(si, r, c, old, cell) {
    const used = this.usedCache.get(si);
    if (used) {
      const had = !!(old && (old.raw || old.image));
      const has = !!(cell && (cell.raw || cell.image));
      if (has && (r >= used.rows || c >= used.cols)) this.usedCache.set(si, { rows: Math.max(used.rows, r + 1), cols: Math.max(used.cols, c + 1) });
      else if (had && !has && (r + 1 >= used.rows || c + 1 >= used.cols)) this.usedCache.delete(si);
    }
    const ext = this.extentCache.get(si);
    if (ext) {
      if (cell && (r >= ext.rows || c >= ext.cols)) this.extentCache.set(si, { rows: Math.max(ext.rows, r + 1), cols: Math.max(ext.cols, c + 1) });
      else if (old && !cell && (r + 1 >= ext.rows || c + 1 >= ext.cols)) this.extentCache.delete(si);
    }
  }

  colIndexUpdate(si, r, c, added) {
    const cur = this.colIdx[si];
    if (!cur || cur.cells !== this.sheets[si].cells) return;
    let l = cur.map.get(c);
    if (l === undefined) return; // 아직 만들지 않은 열
    if (!l) { if (!added) return; l = []; cur.map.set(c, l); }
    let lo = 0;
    let hi = l.length;
    while (lo < hi) { const m = (lo + hi) >> 1; if (l[m] < r) lo = m + 1; else hi = m; }
    if (added) { if (l[lo] !== r) { if (lo === l.length) l.push(r); else l.splice(lo, 0, r); } } else if (l[lo] === r) l.splice(lo, 1);
  }

  // ─────────── 추적 (참조되는 셀 · 참조하는 셀) ───────────
  /** 수식 셀 (si, r, c) 가 참조하는 범위 [{ si, r1, c1, r2, c2 }] (동적 참조가 있으면 dyn: true) */
  precedentsOf(si, r, c) {
    const cell = this.getCell(si, r, c);
    const out = [];
    if (!cell?.formula || !cell.ast) return out;
    const g = this.graph ?? new DepGraph(this, true);
    g.refBoxes(si, r, c, cell, (ts, r1, c1, r2, c2) => {
      if (ts === null) out.dyn = true;
      else out.push({ si: ts, r1, c1, r2, c2 });
    });
    return out;
  }

  /** 칸 (si, r, c) 를 바로 참조하는 수식 셀 [{ si, r, c }] */
  dependentsOf(si, r, c) {
    if (this.pending.length) this.flushPending();
    if (!this.graph) this.graph = new DepGraph(this);
    const out = this.graph.dependentsOf(si, r, c);
    // 동적 수식(INDIRECT 등)은 알 수 없으므로 제외 (엑셀도 추적하지 않음)
    return out;
  }

  // ─────────── 셀 단위 재계산 ───────────
  /**
   * 의존 그래프를 미리 백그라운드로 만듦 (큰 파일을 연 뒤 첫 편집도 바로 반응하도록).
   * 만드는 동안 수식이 바뀌면 버리고 다음 편집 때 다시 만듦
   */
  prepareGraph(budgetMs = 30) {
    if (this.graph) return Promise.resolve(true);
    // 이미 만드는 중이면 같은 작업을 기다림 (두 번째 호출이 첫 작업을 버리지 않게)
    if (this.graphRun) return this.graphRun;
    this.graphRun = this.buildGraphSteps(budgetMs).finally(() => { this.graphRun = null; });
    return this.graphRun;
  }

  async buildGraphSteps(budgetMs) {
    const epoch = (this.graphEpoch = (this.graphEpoch ?? 0) + 1);
    const g = new DepGraph(this, true);
    const it = g.steps();
    let last = performance.now();
    for (;;) {
      if (it.next().done) break;
      if (performance.now() - last > budgetMs) {
        await new Promise((res) => setTimeout(res, 0));
        if (this.graphEpoch !== epoch || this.graph) return false;
        last = performance.now();
      }
    }
    if (this.graphEpoch !== epoch || this.graph) return false;
    this.graph = g;
    return true;
  }

  /** 칸 (si, r, c) 의 값이 바뀜: 트랜잭션 중이면 모아 두고, 아니면 바로 반영 */
  changed(si, r, c) {
    this.pending.push(si, r, c);
    if (!this.tx) this.flushPending();
  }

  flushPending() {
    // 파일을 여는 중 피벗 여러 개를 다시 그림: 바뀐 칸을 모았다가 한 번에 의존 수식을 찾음 (releaseDirty)
    if (this.holdDirty) return;
    const pts = this.pending;
    this.pending = [];
    if (pts.length) this.dirtyPoints(pts);
  }

  /** 바뀐 칸 모으기 시작 · 끝 (끝에서 한 번에 반영) */
  holdDirtyWhile(fn) {
    const prev = this.holdDirty;
    this.holdDirty = true;
    try { return fn(); } finally { this.holdDirty = prev; if (!prev) this.flushPending(); }
  }

  /**
   * 바뀐 칸(평평한 [시트, 행, 열, …])을 참조하는 수식만 다시 계산하게 표시.
   * 의존 그래프로 찾지 못하는 경우(분산 영역 등)는 시트 단위로 다시 계산.
   */
  dirtyPoints(pts) {
    this.version++;
    if (this.manualCalc) {
      // 수동 계산: 바뀐 칸 자신만 다시 계산하고, 참조하는 수식은 F9(지금 계산)까지 이전 값을 유지 (엑셀과 같음)
      pushAll((this.manualPts ??= []), pts);
      this.sheetVer ??= [];
      for (let i = 0; i < pts.length; i += 3) {
        this.caches[pts[i]]?.deleteRC(pts[i + 1], pts[i + 2]);
        this.bumpCol(pts[i], pts[i + 2]);
        this.sheetVer[pts[i]] = (this.sheetVer[pts[i]] ?? 0) + 1;
        const cell = this.sheets[pts[i]]?.cells.getRC(pts[i + 1], pts[i + 2]);
        if (cell) cell.dirty = true;
      }
      return;
    }
    this.sheetVer ??= [];
    const sheetsHit = new Set();
    for (let i = 0; i < pts.length; i += 3) sheetsHit.add(pts[i]);
    if (pts.length > 300000) {
      // 한꺼번에 아주 많이 바뀜(채우기 · 붙여넣기 수십만 칸): 칸마다 찾기보다 시트 단위가 빠름
      for (const s of sheetsHit) this.invalidate(s);
      return;
    }
    // 분산 영역이 걸리면 시트 단위 (분산 크기가 바뀌면 주변 칸 값도 바뀜)
    let fallback = false;
    for (const sp of this.spills.values()) if (sheetsHit.has(sp.si)) { fallback = true; break; }
    let dirty = null;
    if (!fallback) {
      try {
        if (!this.graph || this.graph.stale) this.graph = new DepGraph(this);
        dirty = this.graph.propagate(pts, DIRTY_LIMIT);
      } catch (e) {
        console.warn('의존 그래프 오류, 시트 단위로 다시 계산', e);
        this.graph = null;
        dirty = null;
      }
    }
    if (!dirty) {
      for (const s of sheetsHit) this.invalidate(s);
      return;
    }
    for (let i = 0; i < pts.length; i += 3) { this.caches[pts[i]]?.deleteRC(pts[i + 1], pts[i + 2]); this.bumpCol(pts[i], pts[i + 2]); }
    for (const s of sheetsHit) this.sheetVer[s] = (this.sheetVer[s] ?? 0) + 1;
    const bumped = new Set(sheetsHit);
    let arrays = false;
    for (let i = 0; i < dirty.length; i += 3) {
      const s = dirty[i];
      this.caches[s]?.deleteRC(dirty[i + 1], dirty[i + 2]);
      this.bumpCol(s, dirty[i + 2]);
      const cell = this.sheets[s]?.cells.getRC(dirty[i + 1], dirty[i + 2]);
      if (cell) {
        cell.dirty = true;
        if (cell.maybeArray) arrays = true;
      }
      if (!bumped.has(s)) { bumped.add(s); this.sheetVer[s] = (this.sheetVer[s] ?? 0) + 1; }
    }
    if (arrays) {
      // 분산할 수 있는 수식이 다시 계산되면 분산 영역을 다시 정함
      for (const s of bumped) for (const [k, sp] of this.spills) if (sp.si === s) this.spills.delete(k);
      this.spillOwner.clear();
      for (const [k, sp] of this.spills) for (let i = 0; i < sp.h; i++) for (let j = 0; j < sp.w; j++) if (i || j) this.spillOwner.set(`${sp.si}:${sp.r + i},${sp.c + j}`, k);
      this.spillState = null;
    }
  }

  /** 수동 계산에서 계산하지 않은 변경이 있는지 (상태 표시줄의 [계산]) */
  get needsCalc() { return !!(this.manualCalc && this.manualPts?.length); }

  /** 지금 계산 (F9): 수동 계산 중 쌓인 변경을 의존 수식까지 반영 */
  calculateNow() {
    const pts = this.manualPts ?? [];
    this.manualPts = [];
    const m = this.manualCalc;
    this.manualCalc = false;
    try { if (pts.length) this.dirtyPoints(pts); } finally { this.manualCalc = m; }
  }

  /** 값이 있는 영역 크기 {rows, cols} */
  usedRange(si) {
    if (this.usedCache.has(si)) return this.usedCache.get(si);
    let rows = 0;
    let cols = 0;
    for (const b of this.sheets[si].blocks) { rows = Math.max(rows, b.r0 + b.n); cols = Math.max(cols, b.c0 + b.cols.length); }
    for (const [c, m] of this.sheets[si].cells.cols) {
      for (const [r, cell] of m) {
        if (!cell.raw && !cell.image) continue;
        if (r >= rows) rows = r + 1;
        if (c >= cols) cols = c + 1;
      }
    }
    const res = { rows, cols };
    this.usedCache.set(si, res);
    return res;
  }

  /** 서식만 있는 셀까지 포함한 최대 범위 */
  extent(si) {
    const cached = this.extentCache.get(si);
    if (cached) return cached;
    let rows = 0;
    let cols = 0;
    for (const b of this.sheets[si].blocks) { rows = Math.max(rows, b.r0 + b.n); cols = Math.max(cols, b.c0 + b.cols.length); }
    for (const [c, m] of this.sheets[si].cells.cols) {
      if (c >= cols) cols = c + 1;
      for (const r of m.keys()) if (r >= rows) rows = r + 1;
    }
    const res = { rows, cols };
    this.extentCache.set(si, res);
    return res;
  }

  colWidth(si, c) { const s = this.sheets[si]; return s.colWidths[c] ?? s.defColW ?? DEFAULT_COL_WIDTH; }
  rowHeight(si, r) { const s = this.sheets[si]; return s.rowHeights[r] ?? s.defRowH ?? DEFAULT_ROW_HEIGHT; }

  /** si 를 주면 그 시트의 버전만, 아니면 모든 시트의 버전을 올림 (피벗 원본 캐시가 씀) */
  invalidate(si) {
    this.version++;
    if (si === undefined) this.manualPts = [];
    this.sheetVer ??= [];
    if (si === undefined || !this.sheets[si]) {
      this.sheetVerAll = (this.sheetVerAll ?? 0) + 1;
      this.caches = [];
      this.graph = null;
      this.graphEpoch = (this.graphEpoch ?? 0) + 1;
      this.pending = [];
      this.ctxs = [];
      this.colIdx = [];
      this.deps = null;
      this.affectMemo.clear();
      this.arrayList = null;
      for (const s of this.sheets ?? []) s.fileValues = false;
      this.spills.clear();
      this.spillOwner.clear();
      this.spillState = null;
      this.usedCache.clear();
      this.extentCache.clear();
      return;
    }
    // 바뀐 시트와 그 시트를 (직간접으로) 참조하는 시트만 다시 계산
    this.usedCache.delete(si);
    this.extentCache.delete(si);
    const aff = this.affected(si);
    this.baseVer ??= [];
    for (const s of aff) {
      this.sheetVer[s] = (this.sheetVer[s] ?? 0) + 1;
      this.baseVer[s] = (this.baseVer[s] ?? 0) + 1;
      this.caches[s]?.clear();
      if (this.sheets[s]) this.sheets[s].fileValues = false;
    }
    for (const [k, sp] of this.spills) {
      if (!aff.has(sp.si)) continue;
      this.spills.delete(k);
      for (let i = 0; i < sp.h; i++) for (let j = 0; j < sp.w; j++) if (i || j) this.spillOwner.delete(`${sp.si}:${sp.r + i},${sp.c + j}`);
    }
    this.spillState = null;
  }

  /** 시트별 참조 관계 { sheets: Set(시트 번호), all } — 수식에서 정적으로 찾음 */
  sheetDeps() {
    if (this.deps) return this.deps;
    const byName = new Map(this.sheets.map((s, i) => [s.name.toLowerCase(), i]));
    const tableSheet = new Map();
    this.sheets.forEach((s, i) => { for (const t of s.tables ?? []) tableSheet.set(t.name.toLowerCase(), i); });
    const nameSet = new Set(this.names.map((n) => n.name.toLowerCase()));
    this.depCtx = { byName, tableSheet, nameSet };
    this.deps = this.sheets.map((sheet, si) => {
      const d = { sheets: new Set(), all: false };
      const seen = new Set();
      for (const cell of sheet.cells.values()) {
        if (!cell.formula || !cell.ast || seen.has(cell.ast)) continue;
        seen.add(cell.ast);
        this.addDeps(d, si, cell.ast);
      }
      return d;
    });
    this.affectMemo.clear();
    return this.deps;
  }

  addDeps(d, si, ast) {
    const a = astDeps(ast);
    const { byName, tableSheet, nameSet } = this.depCtx;
    let added = false;
    const add = (s) => { if (s !== undefined && s !== si && !d.sheets.has(s)) { d.sheets.add(s); added = true; } };
    if (a.all && !d.all) { d.all = true; added = true; }
    for (const n of a.sheets) { const s = byName.get(n); if (s === undefined) { if (!d.all) { d.all = true; added = true; } } else add(s); }
    for (const t of a.tables) add(tableSheet.get(t));
    for (const n of a.names) {
      if (nameSet.has(n)) { if (!d.all) { d.all = true; added = true; } } else add(tableSheet.get(n));
    }
    return added;
  }

  /** si 가 바뀌면 다시 계산해야 할 시트들 (si 포함) */
  affected(si) {
    const hit = this.affectMemo.get(si);
    if (hit) return hit;
    const deps = this.sheetDeps();
    const out = new Set([si]);
    const queue = [si];
    while (queue.length) {
      const x = queue.pop();
      deps.forEach((d, s) => {
        if (!out.has(s) && (d.all || d.sheets.has(x))) { out.add(s); queue.push(s); }
      });
    }
    this.affectMemo.set(si, out);
    return out;
  }

  /** 실행 취소 기록 목록이 바꾼 시트만 다시 계산 */
  invalidateEntries(entries, cellsDone = false) {
    if (entries.some((e) => e.t === 'all')) { this.invalidate(); return; }
    if (entries.some((e) => e.t === 'list' || e.t === 'sheet' || e.t === 'names' || e.t === 'rename')) {
      // 구조가 바뀐 시트와 그 시트를 참조하는 시트는 다시 계산
      for (const e of entries) if (e.t === 'sheet') for (const x of this.affected(e.si)) if (this.sheets[x]) this.sheets[x].fileValues = false;
      if (entries.some((e) => e.t === 'names')) this.sheetDeps().forEach((d, i) => { if (d.all) this.sheets[i].fileValues = false; });
      this.invalidateStructure();
    }
    const done = new Set();
    const pts = [];
    for (const e of entries) {
      if (e.t === 'cell') {
        // 셀 변경: 그 칸을 참조하는 수식만 (트랜잭션 안에서 바뀐 칸은 이미 모아 둠)
        if (!cellsDone && !done.has(e.si)) pts.push(e.si, e.r, e.c);
      } else if (e.t === 'perm' || e.t === 'order' || (e.t === 'prop' && !CALC_NEUTRAL.has(e.prop))) {
        if (e.t === 'prop' && e.prop === 'tables') { this.deps = null; this.affectMemo.clear(); this.graph = null; this.graphEpoch = (this.graphEpoch ?? 0) + 1; }
        if (!done.has(e.si)) { done.add(e.si); this.invalidate(e.si); }
      } else if (e.t === 'prop' && e.prop === 'merges') {
        this.spillState = null;
        this.usedCache.delete(e.si);
      }
    }
    if (pts.length) this.dirtyPoints(pts);
    this.version++;
  }

  // ─────────── 트랜잭션 / 실행 취소 ───────────
  /** 여러 변경을 하나의 실행 취소 단위로 묶음. meta는 실행 취소 시 복원할 선택 영역 등 */
  transact(fn, meta) {
    if (this.tx) return fn();
    this.tx = { entries: [], meta };
    let result;
    try {
      result = fn();
    } finally {
      const tx = this.tx;
      this.tx = null;
      for (const e of tx.entries) {
        if (e.t === 'all') e.after = this.serialize();
        else if (e.t === 'prop') e.after = structuredClone(this.sheets[e.si]?.[e.prop]);
        else if (e.t === 'list') e.after = { sheets: [...this.sheets], names: this.copyNames() };
        else if (e.t === 'sheet') e.after = this.serializeSheet(e.si);
        else if (e.t === 'names') e.after = this.copyNames();
      }
      const top = this.undoStack[this.undoStack.length - 1];
      if (tx.entries.length && tx.meta?.joinPrev && top) {
        // 자동으로 따라 바뀐 것(피벗 자동 새로 고침 등)은 직전 실행 취소 단계에 합침
        pushAll(top.entries, tx.entries);
      } else if (tx.entries.length) {
        this.undoStack.push(tx);
        if (this.undoStack.length > 200) this.undoStack.shift();
        this.redoStack = [];
      }
      this.invalidateEntries(tx.entries, true);
      this.flushPending();
      this.emit();
    }
    return result;
  }

  record(entry) {
    if (entry.si !== undefined) this.touch(entry.si);
    if (this.tx) this.tx.entries.push(entry);
  }

  /** 시트 편집 횟수 (자동 저장이 바뀐 시트만 다시 저장하는 데 씀) */
  touch(si) {
    const s = this.sheets[si];
    if (s) s._ev = (s._ev ?? 0) + 1;
  }

  copyNames() {
    return this.names.map(({ _ast, _text, ...n }) => ({ ...n }));
  }

  /** 시트 목록(추가 · 삭제 · 이동)만 실행 취소용으로 기록 — 시트 객체는 그대로 둠 */
  snapshotList() {
    if (this.tx && !this.tx.entries.some((e) => e.t === 'list')) {
      this.tx.entries.push({ t: 'list', before: { sheets: [...this.sheets], names: this.copyNames() } });
    }
  }

  /** 시트 하나를 통째로 기록 (행/열 삽입 · 삭제 등) */
  snapshotSheet(si) {
    this.touch(si);
    if (this.tx && !this.tx.entries.some((e) => e.t === 'sheet' && e.si === si)) {
      this.tx.entries.push({ t: 'sheet', si, before: this.serializeSheet(si) });
    }
  }

  snapshotNames() {
    if (this.tx && !this.tx.entries.some((e) => e.t === 'names')) this.tx.entries.push({ t: 'names', before: this.copyNames() });
  }

  serializeSheet(si) {
    const out = this.sheetMeta(si);
    const cells = {};
    for (const [k, cell] of this.sheets[si].cells) cells[k] = cellData(cell);
    out.cells = cells;
    out._sid = this.sheets[si]._sid;
    if (this.sheets[si].blocks.length) out.blocks = this.sheets[si].blocks.map(blockClone);
    return out;
  }

  /** 시트 하나를 저장 형태로 바꿈 (실행 취소 가능) */
  replaceSheet(si, data) {
    this.snapshotSheet(si);
    this.putSheet(si, data);
    this.invalidateStructure();
  }

  /** 시트 객체는 그대로 두고 내용만 바꿈 (시트 목록 실행 취소 기록이 같은 객체를 가리키므로) */
  putSheet(si, data) {
    const cur = this.sheets[si];
    const next = sheetFromData(data);
    for (const k of Object.keys(cur)) if (!(k in next) && k !== '_sid' && k !== '_ev') delete cur[k];
    Object.assign(cur, next);
    cur.fileValues = !!data.fileValues;
    this.touch(si);
  }

  /** 기록한 셀 교체 (구조 변경 중 다른 시트의 수식 고치기) */
  swapCell(si, k, cell) {
    const [r, c] = unkey(k);
    this.record({ t: 'cell', si, r, c, before: cellData(this.sheets[si].cells.get(k)), after: cellData(cell) });
    this.putCell(si, r, c, cell);
  }

  /** 시트 구성 · 이름이 바뀜: 계산 캐시는 비우되 파일 계산 결과(fileValues)는 유지 */
  /** 의존 그래프만 버림 (다음 변경 때 다시 만듦) — 피벗 영역의 열 범위가 바뀌면 GETPIVOTDATA 의존이 달라짐 */
  dropGraph() {
    this.graph = null;
    this.graphEpoch = (this.graphEpoch ?? 0) + 1;
  }

  invalidateStructure() {
    this.version++;
    this.sheetVerAll = (this.sheetVerAll ?? 0) + 1;
    this.caches = [];
    this.graph = null;
    this.graphEpoch = (this.graphEpoch ?? 0) + 1;
    this.pending = [];
    this.ctxs = [];
    this.colIdx = [];
    this.deps = null;
    this.affectMemo.clear();
    this.arrayList = null;
    this.spills.clear();
    this.spillOwner.clear();
    this.spillState = null;
    this.usedCache.clear();
    this.extentCache.clear();
  }

  snapshotAll() {
    for (let i = 0; i < this.sheets.length; i++) this.touch(i);
    if (this.tx && !this.tx.entries.some((e) => e.t === 'all')) {
      this.tx.entries.push({ t: 'all', before: this.serialize() });
    }
  }

  canUndo() { return this.undoStack.length > 0; }
  canRedo() { return this.redoStack.length > 0; }

  undo() {
    const tx = this.undoStack.pop();
    if (!tx) return null;
    for (const e of [...tx.entries].reverse()) this.applyEntry(e, 'before');
    this.redoStack.push(tx);
    this.invalidateEntries(tx.entries);
    this.emit();
    return tx.meta;
  }

  redo() {
    const tx = this.redoStack.pop();
    if (!tx) return null;
    for (const e of tx.entries) this.applyEntry(e, 'after');
    this.undoStack.push(tx);
    this.invalidateEntries(tx.entries);
    this.emit();
    return tx.meta;
  }

  applyEntry(e, side) {
    if (e.t === 'cell') this.putCell(e.si, e.r, e.c, makeCellRC(e[side], e.r, e.c));
    else if (e.t === 'list') { this.sheets = [...e[side].sheets]; this.names = e[side].names.map((n) => ({ ...n })); }
    else if (e.t === 'sheet') this.putSheet(e.si, e[side]);
    else if (e.t === 'names') this.names = e[side].map((n) => ({ ...n }));
    else if (e.t === 'rename') { if (this.sheets[e.si]) this.sheets[e.si].name = e[side]; }
    else if (e.t === 'perm') { const b = this.sheets[e.si]?.blocks[e.bi]; if (b) { materialize(b); blockPermute(b, e.a, e.n, e.j1, e.j2, e.order, side === 'before'); } }
    else if (e.t === 'order') { const b = this.sheets[e.si]?.blocks[e.bi]; if (b) setRowOrder(b, e.a, e[side]); }
    else if (e.t === 'all') this.restore(e[side]);
    else if (e.t === 'prop') { if (this.sheets[e.si]) this.sheets[e.si][e.prop] = structuredClone(e[side]); }
    else if (e.t === 'colWidth') this.sheets[e.si].colWidths = { ...e[side] };
    else if (e.t === 'rowHeight') {
      this.sheets[e.si].rowHeights = { ...e[side] };
      if (e.beforeManual) this.sheets[e.si].rowManual = { ...(side === 'before' ? e.beforeManual : e.afterManual) };
    }
  }

  putCell(si, r, c, cell) {
    this.touch(si);
    const cells = this.sheets[si].cells;
    const old = cells.getRC(r, c);
    if (old?.maybeArray || cell?.maybeArray) this.arrayList = null;
    if (old?.formula || cell?.formula) {
      if (this.graph) this.graph.set(si, r, c, cell?.formula ? cell : null);
      else this.graphEpoch = (this.graphEpoch ?? 0) + 1; // 백그라운드로 만들던 그래프는 버림
    }
    const b = this.blockAt(si, r, c);
    if (b) {
      // 열 블록 칸: 값만 있고 열 서식과 같으면 블록에, 수식 · 메모 · 다른 서식이면 일반 셀로 (블록 칸은 비움)
      const fmt = b.cols[c - b.c0].fmt ?? undefined;
      const plain = cell && !cell.formula && !cell.comment && !cell.link && !cell.image && cell.cached === undefined && sameStyle(cell.style, fmt);
      if (!cell || plain) {
        const v = cell ? cell.v : null;
        blockSet(b, r, c, isError(v) ? { error: v.code } : v ?? null);
        if (old) { cells.deleteRC(r, c); this.colIndexUpdate(si, r, c, false); }
        if (this.graph && old?.formula) this.graph.set(si, r, c, null);
        return;
      }
      blockSet(b, r, c, null);
    }
    if (cell) { cells.setRC(r, c, cell); if (!old) this.colIndexUpdate(si, r, c, true); } else if (old) { cells.deleteRC(r, c); this.colIndexUpdate(si, r, c, false); }
    this.boundsUpdate(si, r, c, old, cell);
    // 새 수식이 다른 시트를 참조하면 의존 관계에 추가
    if (cell?.formula && cell.ast && this.deps?.[si] && this.addDeps(this.deps[si], si, cell.ast)) this.affectMemo.clear();
  }

  // ─────────── 셀 변경 ───────────
  /** 셀 전체 교체 (data = {raw, style, comment} 또는 null) */
  setCellData(si, r, c, data) {
    const cur = this.getCell(si, r, c);
    const cell = makeCellRC(data, r, c);
    if (sameCell(cur, cell)) return;
    // noUndo: 파일을 여는 중(피벗 다시 그리기) — 실행 취소 기록은 끝나면 비우므로 셀 내용 복사를 만들지 않음
    this.record(this.noUndo ? { t: 'cell', si, r, c } : { t: 'cell', si, r, c, before: cur ? cellData(cur) : null, after: cell ? cellData(cell) : null });
    this.putCell(si, r, c, cell);
    // 서식만 바뀐 값 칸 (피벗 다시 그리기 등): 값이 같으니 참조하는 수식을 다시 계산하지 않음 (엑셀도 서식 변경은 재계산 안 함)
    if (cur && cell && !cur.formula && !cell.formula && !cur.image && !cell.image && sameValue(cur.v, cell.v)) {
      this.version++;
      (this.sheetVer ??= [])[si] = (this.sheetVer[si] ?? 0) + 1;
      if (cur.v !== cell.v) this.bumpCol(si, c);
      return;
    }
    this.changed(si, r, c);
  }

  /** 시트 내용 버전 (그 시트의 셀이 바뀌거나 통합 문서 전체가 바뀌면 달라짐) */
  sheetVersion(si) {
    return `${this.sheetVerAll ?? 0}:${this.sheetVer?.[si] ?? 0}`;
  }

  /** 사용자 입력 → 셀 (서식 유지, 입력에 따른 자동 서식 적용) */
  setInput(si, r, c, raw) {
    const cur = this.getCell(si, r, c);
    const style = { ...(cur?.style || {}) };
    const general = !style.numFmt || style.numFmt === 'general';
    if (raw.startsWith('=') && raw.length > 1) {
      if (general) {
        const ast = parseAt(raw.slice(1), r, c).ast; // 공유 수식 캐시 (백만 행 채우기도 해석 한 번)
        const f = autoFormatFor(ast);
        if (f) style.numFmt = f;
      }
    } else if (general && style.numFmt !== 'text') {
      const p = parseInput(raw);
      if (p.numFmt) {
        style.numFmt = p.numFmt;
        if (p.decimals) style.decimals = p.decimals;
      }
    }
    this.setCellData(si, r, c, { raw, style, comment: cur?.comment, link: cur?.link });
  }

  setStyle(si, r, c, patch) {
    const cur = this.getCell(si, r, c);
    const style = { ...(cur?.style || this.baseStyle || {}), ...patch };
    this.setCellData(si, r, c, { raw: cur?.raw ?? '', style, comment: cur?.comment, link: cur?.link, image: cur?.image, cached: cur?.cached });
  }

  setComment(si, r, c, comment) {
    const cur = this.getCell(si, r, c);
    this.setCellData(si, r, c, { raw: cur?.raw ?? '', style: cur?.style, comment: comment || undefined, link: cur?.link, image: cur?.image });
  }

  clearRange(si, r1, c1, r2, c2, what = 'contents') {
    for (const [k, cell] of [...this.sheets[si].cells]) {
      const [r, c] = unkey(k);
      if (r < r1 || r > r2 || c < c1 || c > c2) continue;
      if (what === 'all') this.setCellData(si, r, c, null);
      else if (what === 'formats') this.setCellData(si, r, c, { raw: cell.raw, comment: cell.comment, link: cell.link, image: cell.image });
      else if (what === 'comments') this.setCellData(si, r, c, { raw: cell.raw, style: cell.style, link: cell.link, image: cell.image });
      else this.setCellData(si, r, c, { raw: '', style: cell.style, comment: cell.comment });
    }
  }

  setColWidth(si, c, w) {
    const sheet = this.sheets[si];
    const before = { ...sheet.colWidths };
    sheet.colWidths[c] = Math.max(0, Math.round(w));
    this.record({ t: 'colWidth', si, before, after: { ...sheet.colWidths } });
  }

  /** manual=false 는 자동 맞춤 높이 (DEFAULT 이면 항목 삭제) */
  setRowHeight(si, r, h, manual = true) {
    const sheet = this.sheets[si];
    const before = { ...sheet.rowHeights };
    const beforeManual = { ...sheet.rowManual };
    const v = Math.max(0, Math.round(h));
    if (!manual && v === (sheet.defRowH ?? DEFAULT_ROW_HEIGHT)) delete sheet.rowHeights[r];
    else sheet.rowHeights[r] = v;
    if (manual) sheet.rowManual[r] = true;
    else delete sheet.rowManual[r];
    this.record({ t: 'rowHeight', si, before, after: { ...sheet.rowHeights }, beforeManual, afterManual: { ...sheet.rowManual } });
  }

  /** 행/열/시트 전체 서식 (kind: 'col' | 'row' | 'all') */
  setLineStyle(si, kind, index, patch) {
    this.propSnap(si, kind === 'all' ? 'allStyle' : kind === 'col' ? 'colStyles' : 'rowStyles');
    const s = this.sheets[si];
    const merged = cleanStyle({ ...(kind === 'all' ? s.allStyle : kind === 'col' ? s.colStyles[index] : s.rowStyles[index]), ...patch });
    if (kind === 'all') s.allStyle = merged ?? null;
    else {
      const map = kind === 'col' ? s.colStyles : s.rowStyles;
      if (merged) map[index] = merged;
      else delete map[index];
    }
  }

  setHidden(si, axis, indices, hidden) {
    this.propSnap(si, axis === 'row' ? 'hiddenRows' : 'hiddenCols');
    this.invalidate(si);
    const map = axis === 'row' ? this.sheets[si].hiddenRows : this.sheets[si].hiddenCols;
    for (const i of indices) { if (hidden) map[i] = true; else delete map[i]; }
  }

  setSheetProp(si, prop, value) {
    this.propSnap(si, prop);
    this.sheets[si][prop] = value;
    if (prop === 'tables') { this.deps = null; this.affectMemo.clear(); }
    // 그림 개체 · 틀 고정 · 조건부 서식 등은 계산 결과에 영향이 없으므로 수식 캐시를 유지
    if (!CALC_NEUTRAL.has(prop)) this.invalidate(si);
    else this.version++;
  }

  /** 시트 속성 하나만 실행 취소용으로 기록 (통합 문서 전체 복사보다 훨씬 가벼움) */
  propSnap(si, prop) {
    this.touch(si);
    const tx = this.tx;
    if (!tx) return;
    const k = `${si}:${prop}`;
    tx.props ??= new Set();
    if (tx.props.has(k)) return;
    tx.props.add(k);
    tx.entries.push({ t: 'prop', si, prop, before: structuredClone(this.sheets[si][prop]) });
  }

  // ─────────── 구조 변경 ───────────
  /** axis: 'row'|'col', count>0 삽입, count<0 삭제 */
  shiftAxis(si, axis, index, count) {
    // 실행 취소: 대상 시트만 통째로, 다른 시트는 바뀐 수식 셀 · 피벗 정의 · 이름만 기록 (큰 통합 문서도 가볍게)
    this.snapshotSheet(si);
    this.snapshotNames();
    const target = this.sheets[si];
    const isRow = axis === 'row';
    const moved = new CellMap();
    target.cells.forEachRC((cell, r, c) => {
      const p = isRow ? r : c;
      if (count < 0 && p >= index && p < index - count) return;
      const np = p >= index ? p + count : p;
      if (isRow) moved.setRC(np, c, cell); else moved.setRC(r, np, cell);
    });
    target.cells = moved;
    target.blocks = target.blocks.map((b) => blockShift(b, axis, index, count)).filter(Boolean);
    const sizes = isRow ? target.rowHeights : target.colWidths;
    const nextSizes = {};
    for (const [k, v] of Object.entries(sizes)) {
      const p = +k;
      if (count < 0 && p >= index && p < index - count) continue;
      nextSizes[p >= index ? p + count : p] = v;
    }
    if (isRow) target.rowHeights = nextSizes; else target.colWidths = nextSizes;
    target.merges = target.merges.map((m) => adjustRange(m, axis, index, count))
      .filter((m) => m && (m.r2 > m.r1 || m.c2 > m.c1));
    target.cond = target.cond.map((rule) => {
      // 추가 범위(more)도 함께 조정, 첫 범위가 사라지면 다음 범위가 대표
      const ranges = [rule, ...(rule.more ?? [])].map((g) => adjustRange({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, axis, index, count)).filter(Boolean);
      if (!ranges.length) return null;
      const { more, ...rest } = rule;
      return { ...rest, ...ranges[0], ...(ranges.length > 1 ? { more: ranges.slice(1) } : {}) };
    }).filter(Boolean);
    target.validations = target.validations.map((v) => adjustRange(v, axis, index, count)).filter(Boolean);
    target.tables = (target.tables ?? []).map((t) => {
      const rg = adjustRange(t, axis, index, count);
      if (!rg) return null;
      const nt = { ...t, ...rg };
      if (isRow) {
        if (t.filter) nt.filter = { ...t.filter, hidden: shiftHidden(t.filter.hidden, index, count) ?? shiftKeys(t.filter.hidden, index, count) };
      } else {
        if (t.filter) nt.filter = { ...t.filter, criteria: shiftKeys(t.filter.criteria, index, count) };
        nt.totalsFns = shiftKeys(t.totalsFns, index, count);
        // 삭제된 열 이름은 빼고, 새 열에는 이름을 채움 (머리글이 없는 표용)
        if (t.columns) {
          const cols = [...t.columns];
          const off = index - t.c1;
          if (count > 0 && off >= 0 && off <= cols.length) cols.splice(off, 0, ...Array(count).fill(''));
          else if (count < 0) cols.splice(Math.max(0, off), Math.min(-count, cols.length - Math.max(0, off)));
          nt.columns = cols;
        }
      }
      return nt;
    }).filter(Boolean);
    // 스파크라인: 이 시트의 위치 이동 + 모든 시트의 데이터 범위 조정
    this.sheets.forEach((sh, i) => {
      if (!sh.sparklines?.length) return;
      if (sh !== target) this.propSnap(i, 'sparklines');
      sh.sparklines = sh.sparklines.map((g) => ({
        ...g,
        items: g.items.map((it) => {
          let { r, c } = it;
          if (sh === target) {
            const p = isRow ? r : c;
            if (count < 0 && p >= index && p < index - count) return null;
            if (p >= index) { if (isRow) r += count; else c += count; }
          }
          const f = adjustFormulaForStructure(`=${it.ref}`, { targetSheet: target.name, hostSheet: sh.name, axis, index, count }).slice(1);
          return f.includes('#REF!') ? null : { ...it, r, c, ref: f };
        }).filter(Boolean),
      })).filter((g) => g.items.length);
    });
    if (target.scenarios?.length) {
      // 시나리오 변경 셀: 삭제된 셀은 빼고 나머지는 이동 (시나리오 요약 결과 셀도 같이)
      const mv = (p) => { const rg = adjustRange({ r1: p.r, c1: p.c, r2: p.r, c2: p.c }, axis, index, count); return rg ? { r: rg.r1, c: rg.c1 } : null; };
      target.scenarios = target.scenarios.map((sc) => {
        const keep = sc.cells.map((p, i) => [mv(p), sc.values[i]]).filter(([p]) => p);
        return keep.length ? { ...sc, cells: keep.map(([p]) => p), values: keep.map(([, v]) => v) } : null;
      }).filter(Boolean);
    }
    if (target.page && (target.page.area || target.page.titleRows || target.page.titleCols)) {
      // 인쇄 영역 · 인쇄 제목도 함께 이동
      const pg = { ...target.page };
      if (pg.area) pg.area = adjustRange(pg.area, axis, index, count);
      const k = isRow ? 'titleRows' : 'titleCols';
      if (pg[k]) {
        const rg = adjustRange(isRow ? { r1: pg[k][0], r2: pg[k][1], c1: 0, c2: 0 } : { c1: pg[k][0], c2: pg[k][1], r1: 0, r2: 0 }, axis, index, count);
        pg[k] = rg ? (isRow ? [rg.r1, rg.r2] : [rg.c1, rg.c2]) : null;
      }
      target.page = pg;
    }
    if (target.outline) {
      // 개요 수준 · 접힘: 가운데에 넣은 행 · 열은 위아래(좌우) 중 낮은 수준을 따름
      const o = { ...target.outline };
      const lk = isRow ? 'rows' : 'cols';
      const ck = isRow ? 'rowsColl' : 'colsColl';
      const before = o[lk] ?? {};
      o[lk] = shiftKeys(before, index, count);
      o[ck] = shiftKeys(o[ck], index, count);
      if (count > 0) {
        const lv = Math.min(before[index - 1] ?? 0, before[index] ?? 0);
        if (lv) for (let i = index; i < index + count; i++) o[lk][i] = lv;
      }
      target.outline = o;
    }
    if (isRow) {
      target.rowStyles = shiftKeys(target.rowStyles, index, count);
      target.hiddenRows = shiftKeys(target.hiddenRows, index, count);
      target.rowManual = shiftKeys(target.rowManual, index, count);
    } else {
      target.colStyles = shiftKeys(target.colStyles, index, count);
      target.hiddenCols = shiftKeys(target.hiddenCols, index, count);
    }
    if (target.filter) {
      const f = adjustRange(target.filter, axis, index, count);
      if (!f) target.filter = null;
      else {
        if (isRow) f.hidden = shiftHidden(f.hidden, index, count) ?? shiftKeys(f.hidden, index, count);
        else f.criteria = shiftKeys(f.criteria, index, count);
        target.filter = f;
      }
    }
    // 차트 범위 · 계열 참조 (모든 시트의 차트 중 대상 시트를 가리키는 것)
    const onTarget = (name, host) => String(name ?? host).toLowerCase() === target.name.toLowerCase();
    const shiftRef = (ref, host) => {
      if (!ref || ref.name || ref.r1 === undefined || !onTarget(ref.sheet, host)) return ref;
      const rg = adjustRange(ref, axis, index, count);
      return rg ? { ...ref, ...rg } : ref;
    };
    this.sheets.forEach((sh, i) => {
      if (!sh.charts?.length) return;
      let changed = false;
      const next = sh.charts.map((ch) => {
        const host = ch.sheet ?? sh.name;
        const out = { ...ch };
        if (ch.range && onTarget(ch.sheet, sh.name)) out.range = adjustRange(ch.range, axis, index, count) ?? ch.range;
        if (ch.series) {
          out.series = ch.series.map((sr) => {
            const n = { ...sr };
            for (const k of ['cat', 'val', 'x', 'size']) if (sr[k]) n[k] = shiftRef(sr[k], host);
            if (sr.name?.ref) n.name = { ...sr.name, ref: shiftRef(sr.name.ref, host) };
            return n;
          });
        }
        if (JSON.stringify(out) === JSON.stringify(ch)) return ch;
        changed = true;
        return out;
      });
      if (!changed) return;
      if (i !== si) this.propSnap(i, 'charts');
      sh.charts = next;
    });
    this.sheets.forEach((sh, i) => {
      const fix = (def) => {
        if (!def || !def.range || String(def.source ?? '').toLowerCase() !== target.name.toLowerCase()) return def;
        const rg = adjustRange(def.range, axis, index, count);
        return rg ? { ...def, range: rg } : def;
      };
      const p = fix(sh.pivot);
      const x = (sh.pivotsExtra ?? []).map(fix);
      if (i !== si && (p !== sh.pivot || x.some((d, j) => d !== sh.pivotsExtra[j]))) { this.propSnap(i, 'pivot'); this.propSnap(i, 'pivotsExtra'); }
      sh.pivot = p;
      sh.pivotsExtra = x;
      // 연결된 그림(카메라)의 원본 범위도 행/열 삽입 · 삭제를 따라감
      const lk = (o) => o.linked && o.linked.sheet.toLowerCase() === target.name.toLowerCase();
      if ((sh.images ?? []).some(lk)) {
        if (i !== si) this.propSnap(i, 'images');
        sh.images = sh.images.map((o) => { if (!lk(o)) return o; const rg = adjustRange(o.linked, axis, index, count); return rg ? { ...o, linked: { ...o.linked, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 } } : o; });
      }
    });

    const deps = this.sheetDeps();
    this.sheets.forEach((sheet, i) => {
      // 다른 시트는 대상 시트를 참조하는 경우만 수식이 바뀜
      if (i !== si && !deps[i]?.sheets.has(si) && !deps[i]?.all) return;
      for (const [k, cell] of [...sheet.cells]) {
        if (!cell.formula) continue;
        const raw = adjustFormulaForStructure(cell.raw, {
          targetSheet: target.name, hostSheet: sheet.name, axis, index, count,
        });
        if (raw === cell.raw) continue;
        const next = makeCell({ ...cellData(cell), raw }, k);
        if (i === si) sheet.cells.set(k, next);
        else this.swapCell(i, k, next);
      }
    });
    this.names = this.names.map((n) => {
      const ref = adjustFormulaForStructure(n.ref, { targetSheet: target.name, hostSheet: n.sheet ?? '', axis, index, count });
      return ref === n.ref ? n : { ...n, ref };
    });
    for (const x of this.affected(si)) this.sheets[x].fileValues = false;
    this.invalidateStructure();
  }

  /**
   * 셀 삽입 · 삭제 (엑셀 [셀을 아래로/오른쪽으로 밀기], [셀을 위로/왼쪽으로 밀기]): rg 칸만 밀고 나머지 열 · 행은 그대로.
   * dir: 'down' | 'right' (삽입), 'up' | 'left' (삭제). 할 수 없으면 이유(문자열), 되면 null.
   */
  shiftCells(si, rg, dir) {
    const target = this.sheets[si];
    const axis = dir === 'down' || dir === 'up' ? 'row' : 'col';
    const isRow = axis === 'row';
    const index = isRow ? rg.r1 : rg.c1;
    const size = isRow ? rg.r2 - rg.r1 + 1 : rg.c2 - rg.c1 + 1;
    const count = dir === 'down' || dir === 'right' ? size : -size;
    const band = isRow ? [rg.c1, rg.c2] : [rg.r1, rg.r2];
    const [o1, o2] = isRow ? ['c1', 'c2'] : ['r1', 'r2'];
    const [p1, p2] = isRow ? ['r1', 'r2'] : ['c1', 'c2'];
    const inBand = (x) => x[o1] >= band[0] && x[o2] <= band[1];
    const crosses = (x) => x[o2] >= band[0] && x[o1] <= band[1] && !inBand(x) && x[p2] >= index;
    // 엑셀처럼 막는 경우: 병합 셀 · 표의 일부만 밀리는 경우, 큰 데이터 블록
    if (target.merges.some(crosses)) return '병합된 셀의 일부를 변경할 수 없습니다.';
    if ((target.tables ?? []).some(crosses)) return '이 작업은 워크시트에 있는 표의 셀을 이동하려고 하므로 사용할 수 없습니다.';
    if (target.blocks.some((b) => (isRow ? b.c0 + b.cols.length - 1 >= band[0] && b.c0 <= band[1] && b.r0 + b.n - 1 >= index : b.r0 + b.n - 1 >= band[0] && b.r0 <= band[1] && b.c0 + b.cols.length - 1 >= index))) {
      return '큰 데이터 영역에서는 셀 밀기를 할 수 없습니다. 행 전체 또는 열 전체를 삽입 · 삭제하세요.';
    }
    this.snapshotSheet(si);
    this.snapshotNames();
    const moved = new CellMap();
    const end = index - count; // 삭제: [index, end)
    target.cells.forEachRC((cell, r, c) => {
      const p = isRow ? r : c;
      const o = isRow ? c : r;
      if (o < band[0] || o > band[1] || p < index) { moved.setRC(r, c, cell); return; }
      if (count < 0 && p < end) return;
      const np = p + count;
      if (np < 0) return;
      if (isRow) moved.setRC(np, c, cell); else moved.setRC(r, np, cell);
    });
    target.cells = moved;
    const adj = (x) => (inBand(x) ? adjustRange(x, axis, index, count) : x);
    target.merges = target.merges.map(adj).filter((m) => m && (m.r2 > m.r1 || m.c2 > m.c1));
    target.cond = target.cond.map((rule) => {
      const ranges = [rule, ...(rule.more ?? [])].map((g) => adj({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 })).filter(Boolean);
      if (!ranges.length) return null;
      const { more, ...rest } = rule;
      return { ...rest, ...ranges[0], ...(ranges.length > 1 ? { more: ranges.slice(1) } : {}) };
    }).filter(Boolean);
    target.validations = target.validations.map(adj).filter(Boolean);
    target.tables = (target.tables ?? []).map((t) => (inBand(t) ? (adjustRange(t, axis, index, count) ? { ...t, ...adjustRange(t, axis, index, count) } : null) : t)).filter(Boolean);
    const onTarget = (name, host) => String(name ?? host).toLowerCase() === target.name.toLowerCase();
    this.sheets.forEach((sh, i) => {
      if (!sh.charts?.length) return;
      const next = sh.charts.map((ch) => {
        const host = ch.sheet ?? sh.name;
        const out = { ...ch };
        if (ch.range && onTarget(ch.sheet, sh.name)) out.range = adj(ch.range) ?? ch.range;
        if (ch.series) out.series = ch.series.map((sr) => { const n = { ...sr }; for (const k of ['cat', 'val', 'x', 'size']) if (sr[k] && !sr[k].name && sr[k].r1 !== undefined && onTarget(sr[k].sheet, host)) n[k] = { ...sr[k], ...(adj(sr[k]) ?? sr[k]) }; return n; });
        return JSON.stringify(out) === JSON.stringify(ch) ? ch : out;
      });
      if (next.some((c, k) => c !== sh.charts[k])) { if (i !== si) this.propSnap(i, 'charts'); sh.charts = next; }
    });
    const deps = this.sheetDeps();
    this.sheets.forEach((sheet, i) => {
      if (i !== si && !deps[i]?.sheets.has(si) && !deps[i]?.all) return;
      for (const [k, cell] of [...sheet.cells]) {
        if (!cell.formula) continue;
        const raw = adjustFormulaForStructure(cell.raw, { targetSheet: target.name, hostSheet: sheet.name, axis, index, count, band });
        if (raw === cell.raw) continue;
        const next = makeCell({ ...cellData(cell), raw }, k);
        if (i === si) sheet.cells.set(k, next);
        else this.swapCell(i, k, next);
      }
    });
    this.names = this.names.map((n) => {
      const ref = adjustFormulaForStructure(n.ref, { targetSheet: target.name, hostSheet: n.sheet ?? '', axis, index, count, band });
      return ref === n.ref ? n : { ...n, ref };
    });
    for (const x of this.affected(si)) this.sheets[x].fileValues = false;
    this.invalidateStructure();
    return null;
  }

  /**
   * 셀 이동 (엑셀: 선택 영역 테두리를 끌어서 옮기기, 잘라내기 → 붙여넣기): src 를 (dr, dc) 만큼 옮기고 대상은 덮어씀.
   * 옮긴 칸을 가리키는 모든 수식 · 이름 · 차트 범위도 따라감. 막히면 한국어 메시지를 돌려줌
   */
  moveRange(si, src, dr, dc) {
    if (!dr && !dc) return null;
    const sh = this.sheets[si];
    const dst = { r1: src.r1 + dr, c1: src.c1 + dc, r2: src.r2 + dr, c2: src.c2 + dc };
    if (dst.r1 < 0 || dst.c1 < 0 || dst.r2 >= MAX_ROWS || dst.c2 >= MAX_COLS) return '시트 밖으로는 옮길 수 없습니다.';
    const over = (a, b) => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
    const inside = (a, b) => a.r1 >= b.r1 && a.r2 <= b.r2 && a.c1 >= b.c1 && a.c2 <= b.c2;
    if (sh.merges.some((m) => (over(m, src) && !inside(m, src)) || (over(m, dst) && !inside(m, dst)))) return '병합된 셀의 일부를 변경할 수 없습니다.';
    if ((sh.tables ?? []).some((t) => over(t, src) || over(t, dst))) return '표 안의 셀은 끌어서 옮길 수 없습니다. 잘라내기 · 붙여넣기를 쓰세요.';
    if (sh.blocks.some((b) => over({ r1: b.r0, c1: b.c0, r2: b.r0 + b.n - 1, c2: b.c0 + b.cols.length - 1 }, { r1: Math.min(src.r1, dst.r1), c1: Math.min(src.c1, dst.c1), r2: Math.max(src.r2, dst.r2), c2: Math.max(src.c2, dst.c2) }))) return '큰 데이터 영역에서는 끌어서 옮길 수 없습니다.';
    this.snapshotSheet(si);
    this.snapshotNames();
    const name = sh.name;
    const fix = (raw, host) => moveRefsInFormula(raw, { targetSheet: name, hostSheet: host, src, dr, dc });
    const moved = [];
    const next = new CellMap();
    sh.cells.forEachRC((cell, r, c) => {
      const inSrc = r >= src.r1 && r <= src.r2 && c >= src.c1 && c <= src.c2;
      if (inSrc) { moved.push([r + dr, c + dc, cell]); return; }
      if (r >= dst.r1 && r <= dst.r2 && c >= dst.c1 && c <= dst.c2) return; // 덮어씀
      next.setRC(r, c, cell);
    });
    for (const [r, c, cell] of moved) next.setRC(r, c, cell);
    sh.cells = next;
    // 수식: 옮긴 칸을 가리키는 참조 이동 (옮긴 수식 자신의 참조는 그대로 — 엑셀과 같음)
    this.sheets.forEach((sheet, i) => {
      for (const [k, cell] of [...sheet.cells]) {
        if (!cell.formula) continue;
        const raw = fix(cell.raw, sheet.name);
        const [rr, cc] = k.split(',').map(Number);
        if (raw === cell.raw && !(i === si && rr >= dst.r1 && rr <= dst.r2 && cc >= dst.c1 && cc <= dst.c2)) continue;
        const nc = makeCell({ ...cellData(cell), raw }, k);
        if (i === si) sheet.cells.set(k, nc); else this.swapCell(i, k, nc);
      }
    });
    const mv = (g) => ({ ...g, r1: g.r1 + dr, r2: g.r2 + dr, c1: g.c1 + dc, c2: g.c2 + dc });
    sh.merges = sh.merges.filter((m) => !inside(m, dst) || inside(m, src)).map((m) => (inside(m, src) ? mv(m) : m));
    sh.cond = sh.cond.map((rule) => {
      const rs = [rule, ...(rule.more ?? [])].map((g) => ({ r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }));
      if (!rs.every((g) => inside(g, src))) return rule;
      const out = rs.map(mv);
      const { more, ...rest } = rule;
      return { ...rest, ...out[0], ...(out.length > 1 ? { more: out.slice(1) } : {}) };
    });
    sh.validations = sh.validations.filter((v) => !inside(v, dst) || inside(v, src)).map((v) => (inside(v, src) ? mv(v) : v));
    this.names = this.names.map((n) => { const ref = fix(n.ref, n.sheet ?? ''); return ref === n.ref ? n : { ...n, ref }; });
    this.sheets.forEach((s2, i) => {
      if (!s2.charts?.length) return;
      const onT = (sheetName, host) => String(sheetName ?? host).toLowerCase() === name.toLowerCase();
      const nextCharts = s2.charts.map((ch) => {
        const host = ch.sheet ?? s2.name;
        const out = { ...ch };
        if (ch.range && onT(ch.sheet, s2.name) && inside(ch.range, src)) out.range = mv(ch.range);
        if (ch.series) out.series = ch.series.map((sr) => { const n = { ...sr }; for (const k of ['cat', 'val', 'x', 'size']) if (sr[k] && sr[k].r1 !== undefined && onT(sr[k].sheet, host) && inside(sr[k], src)) n[k] = mv(sr[k]); return n; });
        return JSON.stringify(out) === JSON.stringify(ch) ? ch : out;
      });
      if (nextCharts.some((c, k) => c !== s2.charts[k])) { if (i !== si) this.propSnap(i, 'charts'); s2.charts = nextCharts; }
    });
    for (const x of this.affected(si)) this.sheets[x].fileValues = false;
    this.invalidateStructure();
    return null;
  }

  insertRows(si, index, count = 1) { this.shiftAxis(si, 'row', index, count); }
  deleteRows(si, index, count = 1) { this.shiftAxis(si, 'row', index, -count); }
  insertCols(si, index, count = 1) { this.shiftAxis(si, 'col', index, count); }
  deleteCols(si, index, count = 1) { this.shiftAxis(si, 'col', index, -count); }

  /** 외부 통합 문서 값 시트(맨 뒤에 모아 둠)의 시작 번호 = 보통 시트 수 */
  ownSheetCount() {
    const i = this.sheets.findIndex((x) => x.external);
    return i < 0 ? this.sheets.length : i;
  }

  addSheet(name, at = this.ownSheetCount()) {
    at = Math.min(at, this.ownSheetCount());
    this.snapshotList();
    let n = this.sheets.length + 1;
    let nm = name;
    while (!nm || this.sheetIndexByName(nm) >= 0) nm = `Sheet${n++}`;
    this.sheets.splice(at, 0, newSheet(nm));
    this.invalidateStructure();
    return at;
  }

  deleteSheet(si) {
    if (this.sheets.length <= 1) return false;
    this.snapshotList();
    const gone = this.sheets[si].name.toLowerCase();
    // 지운 시트를 참조하던 시트는 다시 계산 (#REF!)
    for (const x of this.affected(si)) if (x !== si) this.sheets[x].fileValues = false;
    this.names = this.names.filter((n) => (n.sheet ?? '').toLowerCase() !== gone);
    this.sheets.splice(si, 1);
    this.invalidateStructure();
    return true;
  }

  renameSheet(si, newName) {
    newName = newName.trim();
    if (!newName || /[\\/?*[\]:]/.test(newName) || newName.length > 31) return false;
    const dup = this.sheetIndexByName(newName);
    if (dup >= 0 && dup !== si) return false;
    this.snapshotNames();
    const old = this.sheets[si].name;
    const refs = this.affected(si); // 이 시트를 참조하는 시트 (이름이 바뀐 수식)
    this.sheets[si].name = newName;
    this.record({ t: 'rename', si, before: old, after: newName });
    this.sheets.forEach((sh, i) => {
      const ren = (d) => (d && String(d.source ?? '').toLowerCase() === old.toLowerCase() ? { ...d, source: newName } : d);
      const p = ren(sh.pivot);
      const x = (sh.pivotsExtra ?? []).map(ren);
      if (p !== sh.pivot || x.some((d, j) => d !== sh.pivotsExtra[j])) { this.propSnap(i, 'pivot'); this.propSnap(i, 'pivotsExtra'); }
      sh.pivot = p;
      sh.pivotsExtra = x;
      // 연결된 그림(카메라)의 원본 시트 이름
      if ((sh.images ?? []).some((o) => o.linked && o.linked.sheet.toLowerCase() === old.toLowerCase())) {
        this.propSnap(i, 'images');
        sh.images = sh.images.map((o) => (o.linked && o.linked.sheet.toLowerCase() === old.toLowerCase() ? { ...o, linked: { ...o.linked, sheet: newName } } : o));
      }
    });
    this.sheets.forEach((sheet, i) => {
      if (!refs.has(i)) return;
      for (const [k, cell] of [...sheet.cells]) {
        if (!cell.formula) continue;
        const raw = renameSheetInFormula(cell.raw, old, newName);
        if (raw !== cell.raw) this.swapCell(i, k, makeCell({ ...cellData(cell), raw }, k));
      }
    });
    this.names = this.names.map((n) => ({
      ...n,
      ref: renameSheetInFormula(n.ref, old, newName),
      sheet: n.sheet && n.sheet.toLowerCase() === old.toLowerCase() ? newName : n.sheet,
    }));
    this.invalidateStructure();
    return true;
  }

  /** 범위를 keyCol 기준으로 정렬 (행 단위 이동, 수식 상대 참조 보정) */
  sortRange(si, r1, c1, r2, c2, keyCol, ascending = true) {
    if (this.sortBlock(si, r1, c1, r2, c2, keyCol, ascending)) return;
    const rows = [];
    for (let r = r1; r <= r2; r++) {
      const cells = [];
      for (let c = c1; c <= c2; c++) cells.push(cellData(this.getCell(si, r, c)));
      rows.push({ r, cells, v: this.getValue(si, r, keyCol) });
    }
    const blank = (v) => v === null || v === '';
    rows.sort((a, b) => {
      if (blank(a.v) || blank(b.v)) return blank(a.v) - blank(b.v);
      const ea = isError(a.v);
      const eb = isError(b.v);
      if (ea || eb) return ea - eb;
      const c = compareSortValues(a.v, b.v);
      return ascending ? c : -c;
    });
    rows.forEach((row, i) => {
      const r = r1 + i;
      row.cells.forEach((d, j) => {
        const data = d && d.raw.startsWith('=') ? { ...d, raw: shiftFormula(d.raw, r - row.r, 0) } : d;
        this.setCellData(si, r, c1 + j, data);
      });
    });
  }

  /**
   * 엑셀 정렬 대화상자: 여러 기준 (keys = [{ at: 열(또는 행) 번호, asc, on: 'value'|'fill'|'font', color, list: [항목…] }])
   * 옵션: byCols (왼쪽에서 오른쪽 · 열 순서를 바꿈), caseSensitive (대/소문자 구분, 소문자 먼저), natural (자연 정렬: 글자 속 숫자를 수 크기로)
   */
  sortMulti(si, r1, c1, r2, c2, keys, { byCols = false, caseSensitive = false, natural = false } = {}) {
    if (!keys.length) return;
    const k0 = keys[0];
    if (!byCols && keys.length === 1 && (k0.on ?? 'value') === 'value' && !k0.list && !caseSensitive && !natural) { this.sortRange(si, r1, c1, r2, c2, k0.at, k0.asc !== false); return; }
    const coll = caseSensitive || natural ? new Intl.Collator('en', { sensitivity: caseSensitive ? 'case' : 'accent', caseFirst: 'lower', numeric: natural }) : null;
    const cmpVal = (a, b) => (coll && typeof a === 'string' && typeof b === 'string' ? coll.compare(a, b) : compareSortValues(a, b));
    const blank = (v) => v === null || v === '' || v === undefined;
    const lines = [];
    const n = byCols ? c2 - c1 + 1 : r2 - r1 + 1;
    for (let i = 0; i < n; i++) {
      const cells = [];
      if (byCols) for (let r = r1; r <= r2; r++) cells.push(cellData(this.getCell(si, r, c1 + i)));
      else for (let c = c1; c <= c2; c++) cells.push(cellData(this.getCell(si, r1 + i, c)));
      const at = (k) => (byCols ? [k.at, c1 + i] : [r1 + i, k.at]);
      const key = keys.map((k) => {
        const [r, c] = at(k);
        if (k.on === 'fill' || k.on === 'font') {
          const st = this.styleAt(si, r, c);
          const col = String((k.on === 'fill' ? st.fill : st.color) ?? '').toLowerCase();
          return { hit: col === String(k.color ?? '').toLowerCase() };
        }
        return { v: this.getValue(si, r, c) };
      });
      lines.push({ i, cells, key });
    }
    lines.sort((a, b) => {
      for (let j = 0; j < keys.length; j++) {
        const k = keys[j];
        const x = a.key[j];
        const y = b.key[j];
        if (k.on === 'fill' || k.on === 'font') {
          if (x.hit !== y.hit) return (x.hit ? -1 : 1) * (k.asc === false ? -1 : 1);
          continue;
        }
        if (blank(x.v) || blank(y.v)) { const d = blank(x.v) - blank(y.v); if (d) return d; continue; }
        const ea = isError(x.v);
        const eb = isError(y.v);
        if (ea || eb) { const d = ea - eb; if (d) return d; continue; }
        let c;
        if (k.list) {
          const ia = k.list.findIndex((t) => String(t).toLowerCase() === String(x.v).toLowerCase());
          const ib = k.list.findIndex((t) => String(t).toLowerCase() === String(y.v).toLowerCase());
          c = ia >= 0 && ib >= 0 ? ia - ib : ia >= 0 ? -1 : ib >= 0 ? 1 : cmpVal(x.v, y.v);
        } else c = cmpVal(x.v, y.v);
        if (c) return k.asc === false ? -c : c;
      }
      return a.i - b.i;
    });
    lines.forEach((line, i) => {
      if (line.i === i) return;
      line.cells.forEach((d, j) => {
        const r = byCols ? r1 + j : r1 + i;
        const c = byCols ? c1 + i : c1 + j;
        const data = d && d.raw.startsWith('=') ? { ...d, raw: shiftFormula(d.raw, byCols ? 0 : i - line.i, byCols ? i - line.i : 0) } : d;
        this.setCellData(si, r, c, data);
      });
    });
  }

  /**
   * 정렬 범위가 열 블록 안이고 일반 셀이 섞여 있지 않으면 블록을 바로 재배치 (천만 행도 1~2초).
   * 실행 취소 기록은 순서(순열)만 저장
   */
  sortBlock(si, r1, c1, r2, c2, keyCol, ascending) {
    const sheet = this.sheets[si];
    const bi = sheet.blocks.findIndex((b) => inBlock(b, r1, c1) && inBlock(b, r2, c2));
    if (bi < 0 || r2 - r1 < 1000) return false;
    for (let c = c1; c <= c2; c++) {
      const m = sheet.cells.col(c);
      if (m) for (const r of m.keys()) if (r >= r1 && r <= r2) return false;
    }
    const b = sheet.blocks[bi];
    const a = r1 - b.r0;
    const n = r2 - r1 + 1;
    const key = logicalCol(b, keyCol - b.c0, a, n);
    const order = sortOrder(key, key.a, n, ascending, compareSortValues);
    if (c1 === b.c0 && c2 === b.c0 + b.cols.length - 1) {
      // 블록 전체 너비: 데이터를 옮기지 않고 행 순서만 바꿈 (실행 취소도 즉시)
      const before = reorderRows(b, a, n, order);
      this.record({ t: 'order', si, bi, a, before, after: b.perm.slice(a, a + n) });
    } else {
      materialize(b);
      blockPermute(b, a, n, c1 - b.c0, c2 - b.c0, order);
      this.record({ t: 'perm', si, bi, a, n, j1: c1 - b.c0, j2: c2 - b.c0, order });
    }
    this.invalidate(si);
    return true;
  }

  // ─────────── 셀 병합 / 조건부 서식 ───────────
  mergeAt(si, r, c) {
    return this.sheets[si].merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) ?? null;
  }

  /** 범위와 겹치는 병합 영역 목록 */
  mergesIn(si, r1, c1, r2, c2) {
    return this.sheets[si].merges.filter((m) => m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1);
  }

  /** 병합: 왼쪽 위 셀 값만 남기고 나머지 값 삭제 */
  merge(si, r1, c1, r2, c2) {
    if (r1 === r2 && c1 === c2) return;
    this.propSnap(si, 'merges');
    const sheet = this.sheets[si];
    sheet.merges = sheet.merges.filter((m) => !(m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1));
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) {
        if (r === r1 && c === c1) continue;
        const cell = this.getCell(si, r, c);
        if (cell?.raw) this.setCellData(si, r, c, { raw: '', style: cell.style, comment: cell.comment });
      }
    }
    sheet.merges.push({ r1, c1, r2, c2 });
    this.invalidate(si);
  }

  unmerge(si, r1, c1, r2, c2) {
    const sheet = this.sheets[si];
    const keep = sheet.merges.filter((m) => !(m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1));
    if (keep.length === sheet.merges.length) return;
    this.propSnap(si, 'merges');
    sheet.merges = keep;
  }

  addCondRule(si, rule) {
    this.propSnap(si, 'cond');
    this.sheets[si].cond.unshift(rule); // 새 규칙이 가장 높은 우선순위
    this.version++;
  }

  clearCondRules(si, range) {
    this.propSnap(si, 'cond');
    const sheet = this.sheets[si];
    sheet.cond = range
      ? sheet.cond.filter((c) => !(c.r1 <= range.r2 && c.r2 >= range.r1 && c.c1 <= range.c2 && c.c2 >= range.c1))
      : [];
    this.version++;
  }

  // ─────────── 저장 / 불러오기 ───────────
  /** 시트 밖의 통합 문서 속성 (저장용: 기본 글꼴 · 기본 서식 · 테마 · 매크로 · 이름) */
  bookMeta() {
    return {
      ...(this.vba ? { vba: this.vba } : {}),
      ...(this.externals?.length ? { externals: this.externals } : {}),
      ...(this.defaultFont ? { defaultFont: { ...this.defaultFont } } : {}),
      ...(this.baseStyle ? { baseStyle: { ...this.baseStyle } } : {}),
      ...(this.cellStyles?.length ? { cellStyles: this.cellStyles.map((c) => ({ ...c, style: { ...c.style } })) } : {}),
      ...(this.theme ? { theme: [...this.theme] } : {}),
      ...(this.themeXml ? { themeXml: this.themeXml } : {}),
      ...(this.themeName ? { themeName: this.themeName } : {}),
      ...(this.props && Object.keys(this.props).length ? { props: { ...this.props } } : {}),
      ...(this.names.length ? { names: this.names.map(({ _ast, _text, ...n }) => ({ ...n })) } : {}),
    };
  }

  serialize() {
    return {
      version: 1,
      ...this.bookMeta(),
      sheets: this.sheets.map((s) => {
        const cells = {};
        for (const [k, cell] of s.cells) cells[k] = cellData(cell);
        const out = { name: s.name, cells };
        if (s.blocks.length) out.blocks = s.blocks.map(blockClone);
        for (const p of SHEET_PROPS) out[p] = structuredClone(s[p]);
        if (s.fileValues) out.fileValues = true; // 셀의 파일 계산 결과가 아직 유효함
        return out;
      }),
    };
  }

  /** 시트 하나의 셀 외 속성 (저장용) */
  sheetMeta(si) {
    const s = this.sheets[si];
    const out = { name: s.name };
    for (const p of SHEET_PROPS) out[p] = structuredClone(s[p]);
    if (s.fileValues) out.fileValues = true;
    return out;
  }

  /** 시트의 셀을 size 개씩 [[키, 저장 형태], …] 로 (큰 문서를 나눠 저장) */
  *cellChunks(si, size = 20000) {
    let chunk = [];
    for (const [k, cell] of this.sheets[si].cells) {
      chunk.push([k, cellData(cell)]);
      if (chunk.length >= size) { yield chunk; chunk = []; }
    }
    if (chunk.length) yield chunk;
  }

  restore(data) {
    const it = this.restoreSteps(data);
    while (!it.next().done) { /* 한 번에 */ }
  }

  /** 복원 단계: 셀이 많으면 중간중간 진행률(0~1)을 내보냄. 끝날 때까지 통합 문서는 바뀌지 않음 */
  *restoreSteps(data) {
    const total = data.sheets.reduce((n, s) => n + (s.cells instanceof Map || s.cells instanceof CellMap ? s.cells.size : 0), 0) || 1;
    let done = 0;
    const sheets = [];
    for (const s of data.sheets) {
      const sheet = newSheet(s.name);
      const owned = s.cells instanceof Map || s.cells instanceof CellMap;
      if (owned) {
        // 파일에서 읽은 셀 Map: 새 Map 을 만들지 않고 그 자리에서 셀 객체로 바꿈 (큰 파일에서 훨씬 빠름)
        const src = s.cells;
        const cells = src instanceof CellMap ? src : new CellMap();
        let n = 0;
        if (src instanceof CellMap) {
          for (const [c, m] of src.cols) {
            for (const [r, d] of m) {
              const cell = makeCellRC(d, r, c);
              if (cell) m.set(r, cell);
              else cells.deleteRC(r, c);
              if (++n % 50000 === 0) yield (done + n) / total;
            }
          }
        } else {
          for (const [k, d] of src) {
            const i = k.indexOf(',');
            const r = +k.slice(0, i);
            const c = +k.slice(i + 1);
            const cell = makeCellRC(d, r, c);
            if (cell) cells.setRC(r, c, cell);
            if (++n % 50000 === 0) yield (done + n) / total;
          }
        }
        done += n;
        sheet.cells = cells;
        s.cells = null; // 이제 통합 문서가 소유
      } else {
        for (const [k, d] of Object.entries(s.cells || {})) {
          const cell = makeCell(d, k);
          if (cell) sheet.cells.set(k, cell);
        }
      }
      for (const p of SHEET_PROPS) if (s[p] !== undefined && s[p] !== null) sheet[p] = structuredClone(s[p]);
      sheet.freeze = { rows: 0, cols: 0, ...(s.freeze || {}) };
      // 열 블록: 파일에서 막 읽은 것은 그대로 가져오고, 저장본(실행 취소 기록 등)은 복사
      sheet.blocks = owned || !s.blocks ? (s.blocks ?? []) : s.blocks.map(blockClone);
      if (s._sid) { sheet._sid = s._sid; sheet._ev = s._ev; } // 자동 저장 기록 (바뀐 시트만 다시 저장)
      sheets.push(sheet);
    }
    this.defaultFont = data.defaultFont ?? null; // 통합 문서 기본 글꼴 { name, size } (없으면 맑은 고딕 11)
    this.fitRows = data.fitRows ?? null; // 파일을 열 때 자동 높이로 맞출 행 (저장하지 않음)
    this.theme = data.theme ?? null; // 파일의 테마 색 (없으면 Office 기본)
    this.themeXml = data.themeXml ?? null; // 파일의 테마 XML (글꼴 · 효과 등을 그대로 저장)
    this.themeName = data.themeName ?? null;
    // 문서 속성 · 보호 (엑셀 파일 › 정보): { title, subject, tags, category, comments, creator, lastModifiedBy, created, modified, readOnlyRecommended, lockStructure, markedFinal }
    this.props = data.props ? { ...data.props } : {};
    this.cellStyles = data.cellStyles ?? null; // 이름 있는 셀 스타일 [{ name, style, builtinId? }] (엑셀 [셀 스타일] 사용자 지정)
    this.baseStyle = data.baseStyle ?? null; // 기본 셀 서식 (xlsx 의 xf 0) — 서식이 없는 셀에 적용
    this.vba = data.vba ?? null; // .xlsm 의 매크로(vbaProject.bin, base64) — 실행하지 않고 보존만 함
    this.externals = data.externals ?? null; // 외부 통합 문서 연결 (xlsx externalLink 원본 — 저장할 때 그대로 되돌려 씀)
    this.names = (data.names ?? []).map((n) => ({ ...n }));
    this.sheets = sheets.length ? sheets : [newSheet('Sheet1')];
    this.invalidate();
    data.sheets.forEach((s, i) => { if (s.fileValues && this.sheets[i]) this.sheets[i].fileValues = true; });
  }

  load(data) {
    this.setSnapshots(data);
    this.restore(data);
    this.undoStack = [];
    this.redoStack = [];
  }

  /** 큰 문서: 화면이 멈추지 않도록 나눠서 불러옴. onProgress(0~1) */
  /** 파일의 피벗 캐시 저장본 (엑셀이 마지막으로 새로 고친 원본) — 원본을 고치기 전까지 피벗을 이것으로 계산 */
  setSnapshots(data) {
    this.pivotSnapshots = data?.pivotSnapshots ? new Map(Object.entries(data.pivotSnapshots).map(([k, rows]) => [k, { rows, ver: undefined }])) : null;
  }

  async loadAsync(data, onProgress) {
    this.setSnapshots(data);
    const it = this.restoreSteps(data);
    let last = performance.now();
    for (;;) {
      const s = it.next();
      if (s.done) break;
      if (performance.now() - last > 40) {
        onProgress?.(s.value);
        await new Promise((res) => setTimeout(res, 0));
        last = performance.now();
      }
    }
    this.undoStack = [];
    this.redoStack = [];
  }
}
