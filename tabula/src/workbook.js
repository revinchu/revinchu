// 통합 문서 모델: 시트 · 셀 · 재계산 · 실행 취소 · 행/열 구조 변경
import { resolveStructRef, findTable } from './tables.js';
import { pivotSourceData, pivotLookup } from './pivot.js';
import {
  parse, evaluateArray, evalAny, ERR, compareValues, isError, autoFormatFor, mayReturnArray, Range, RefValue,
  adjustFormulaForStructure, renameSheetInFormula, shiftFormula, quoteSheetName, MAX_ROWS, MAX_COLS,
} from './formula.js';
import { parseInput } from './format.js';
import { CellImage } from './fxcore.js';

export const DEFAULT_COL_WIDTH = 64;
export const DEFAULT_ROW_HEIGHT = 20;

const key = (r, c) => `${r},${c}`;
const unkey = (k) => k.split(',').map(Number);

// 같은 서식 객체를 여러 셀이 공유할 때(파일 가져오기 등) 한 번만 정리. 셀 서식은 바꿀 때 항상 새 객체로 교체하므로 공유해도 안전
const cleanMemo = new WeakMap();
function cleanStyle(style) {
  if (!style) return undefined;
  const hit = cleanMemo.get(style);
  if (hit !== undefined) return hit || undefined;
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

/** 저장 형태 {raw, style, comment} → 계산용 셀 객체 */
export function makeCell(data) {
  if (!data) return null;
  const cell = { raw: data.raw ?? '' };
  const style = cleanStyle(data.style);
  if (style) cell.style = style;
  if (data.comment) cell.comment = data.comment;
  if (data.link) cell.link = data.link; // 하이퍼링크: 주소(URL) 또는 '#시트!A1'
  if (data.cached !== undefined) cell.cached = data.cached;
  if (data.image?.src) cell.image = { ...data.image }; // 셀에 배치한 그림
  if (cell.raw.startsWith('=') && cell.raw.length > 1 && style?.numFmt !== 'text') {
    cell.formula = true;
    const p = parseMemo(cell.raw.slice(1));
    cell.ast = p.ast;
    if (p.ast) cell.maybeArray = p.maybeArray;
    else cell.parseError = p.error;
  } else if (cell.image && cell.raw === '') {
    cell.v = new CellImage(cell.image);
  } else if (style?.numFmt === 'text') {
    cell.v = cell.raw === '' ? null : cell.raw;
  } else {
    cell.v = PLAIN_NUMBER.test(cell.raw) ? Number(cell.raw) : parseInput(cell.raw).value;
  }
  if (!cell.raw && !cell.style && !cell.comment && !cell.link && !cell.image) return null;
  return cell;
}

export function cellData(cell) {
  if (!cell) return null;
  const d = { raw: cell.raw };
  if (cell.style) d.style = { ...cell.style };
  if (cell.comment) d.comment = cell.comment;
  if (cell.link) d.link = cell.link;
  if (cell.image) d.image = { ...cell.image };
  if (cell.cached !== undefined && cell.formula) d.cached = cell.cached;
  return d;
}

/** 지원하지 않는 함수 수식의 파일 속 계산 결과 (저장 가능한 형태) */
function cachedValue(c) {
  if (c && typeof c === 'object' && c.error) return ERR[Object.keys(ERR).find((k) => ERR[k].code === c.error)] ?? ERR.NAME;
  return c;
}

function newSheet(name) {
  return {
    name, cells: new Map(), colWidths: {}, rowHeights: {}, merges: [], cond: [],
    colStyles: {}, rowStyles: {}, allStyle: null, hiddenRows: {}, hiddenCols: {}, rowManual: {},
    freeze: { rows: 0, cols: 0 }, filter: null, charts: [], pivot: null,
    validations: [], images: [], shapes: [], tables: [], slicers: [], pivotsExtra: [],
  };
}

/** 저장 형태의 시트 하나 → 시트 객체 (셀은 객체 또는 Map) */
function sheetFromData(s) {
  const sheet = newSheet(s.name);
  const entries = s.cells instanceof Map ? s.cells : Object.entries(s.cells || {});
  for (const [k, d] of entries) {
    const cell = makeCell(d);
    if (cell) sheet.cells.set(k, cell);
  }
  for (const p of SHEET_PROPS) if (s[p] !== undefined && s[p] !== null) sheet[p] = structuredClone(s[p]);
  sheet.freeze = { rows: 0, cols: 0, ...(s.freeze || {}) };
  if (s._sid) sheet._sid = s._sid;
  return sheet;
}

/** 시트의 부가 속성 (셀 외) — 저장/복원/복제용 */
const SHEET_PROPS = ['colWidths', 'rowHeights', 'merges', 'cond', 'colStyles', 'rowStyles', 'allStyle',
  'hiddenRows', 'hiddenCols', 'rowManual', 'freeze', 'filter', 'charts', 'pivot', 'validations', 'images', 'shapes', 'tables', 'slicers', 'pivotsExtra', 'state', 'noGrid'];
// 바뀌어도 수식 결과가 달라지지 않는 시트 속성
const CALC_NEUTRAL = new Set(['state', 'noGrid', 'charts', 'images', 'shapes', 'slicers', 'freeze', 'cond', 'validations', 'colStyles', 'rowStyles', 'allStyle', 'merges']);

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
    if (data) this.load(data);
    else this.sheets = [newSheet('Sheet1')];
  }

  // ─────────── 조회 ───────────
  onChange(fn) { this.listeners.add(fn); }
  emit() { for (const fn of this.listeners) fn(); }

  sheetIndexByName(name) {
    const n = name.toLowerCase();
    return this.sheets.findIndex((s) => s.name.toLowerCase() === n);
  }

  getCell(si, r, c) { return this.sheets[si]?.cells.get(key(r, c)); }

  getRaw(si, r, c) { return this.getCell(si, r, c)?.raw ?? ''; }

  getValue(si, r, c) {
    const kk = key(r, c);
    const sheet = this.sheets[si];
    const cell = sheet?.cells.get(kk);
    if (!cell || (!cell.formula && cell.raw === '' && !cell.image)) return this.spillValueAt(si, r, c);
    if (!cell.formula) return cell.v ?? null;
    const cache = this.caches[si] ??= new Map();
    const hit = cache.get(kk);
    if (hit !== undefined || cache.has(kk)) return hit;
    // 파일에서 연 뒤 아직 바뀐 것이 없으면 엑셀이 저장해 둔 계산 결과를 그대로 사용
    if (sheet.fileValues && cell.cached !== undefined && !cell.maybeArray) return cachedValue(cell.cached);
    const k = `${si}:${kk}`;
    if (cell.ast === null) return ERR.NAME;
    if (this.evaluating.has(k)) return ERR.CIRC;
    if (this.depth > 0 || this.warming) return this.evalCell(k, cell, si);
    // 최상위 호출: 긴 참조 사슬(예: 누계 1만 행)은 위에서부터 차례로 미리 계산한 뒤 다시 시도
    try {
      return this.evalCell(k, cell, si);
    } catch (e) {
      if (e !== DEEP) throw e;
      this.warmup();
      try {
        return this.evalCell(k, cell, si);
      } catch (e2) {
        if (e2 !== DEEP) throw e2;
        return ERR.CIRC;
      }
    }
  }

  evalCell(k, cell, si) {
    if (this.depth >= MAX_DEPTH) throw DEEP;
    this.evaluating.add(k);
    this.depth++;
    let v;
    const comma = k.indexOf(',');
    const r = Number(k.slice(k.indexOf(':') + 1, comma));
    const c = Number(k.slice(comma + 1));
    try {
      v = evaluateArray(cell.ast, this.ctxFor(si, r, c));
    } catch (e) {
      if (e instanceof RangeError) throw DEEP;
      throw e;
    } finally {
      this.depth--;
      this.evaluating.delete(k);
    }
    // 지원하지 않는 함수는 파일에 저장된 계산 결과를 그대로 표시
    if (v === ERR.NAME && cell.cached !== undefined) v = cachedValue(cell.cached);
    if (v instanceof Range) v = this.placeSpill(k, si, r, c, v);
    (this.caches[si] ??= new Map()).set(key(r, c), v);
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
        if (sheet.cells.get(key(r + i, c + j))?.raw) return ERR.SPILL;
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
  ensureSpills() {
    if (this.spillState === 'done') return;
    if (this.spillState === 'running') { this.runSpillQueue(); return; }
    this.spillState = 'running';
    if (!this.arrayList) {
      const list = [];
      this.sheets.forEach((sheet, si) => {
        for (const [kk, cell] of sheet.cells) if (cell.formula && cell.maybeArray) { const [r, c] = unkey(kk); list.push([si, r, c]); }
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
      this.getValue(si, r, c);
    }
  }

  spillValueAt(si, r, c) {
    this.ensureSpills();
    const owner = this.spillOwner.get(`${si}:${r},${c}`);
    if (!owner) return null;
    const sp = this.spills.get(owner);
    const v = sp.rows[r - sp.r]?.[c - sp.c];
    return v === null || v === undefined ? 0 : v;
  }

  /** 분산된 셀이면 원본(앵커) 셀 {r, c} */
  spillAnchorOf(si, r, c) {
    this.ensureSpills();
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
    this.ensureSpills();
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
        for (const [k, cell] of sheet.cells) if (cell.formula) { const [r, c] = unkey(k); list.push([si, r, c]); }
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
    const own = s.cells.get(key(r, c))?.style;
    const col = s.colStyles[c];
    const row = s.rowStyles[r];
    if (!s.allStyle && !col && !row) return own ?? EMPTY_STYLE;
    return { ...s.allStyle, ...col, ...row, ...own };
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

  ctxFor(si, r = null, c = null) {
    const here = r === null ? null : { si, r, c, sheet: this.sheets[si]?.name };
    return {
      here,
      name: (n, sheet) => this.nameValue(n, si, sheet, here),
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
        const s = this.resolveSheet(ref.sheet, si);
        const sh = this.sheets[s];
        const defs = [sh.pivot, ...(sh.pivotsExtra ?? [])].filter(Boolean);
        const def = defs.find((d) => d.area && ref.r1 >= d.area.r1 && ref.r1 <= d.area.r2 && ref.c1 >= d.area.c1 && ref.c1 <= d.area.c2)
          ?? (defs.length && !defs[0].area ? defs[0] : null);
        if (!def) throw ERR.REF;
        const src = pivotSourceData(this, def);
        const v = src ? pivotLookup(src.rows, def, field, items) : null;
        if (v === null || v === undefined) throw ERR.REF;
        return v;
      },
      quoteSheet: (name) => quoteSheetName(name),
      colWidthChars: (sheet, cc) => Math.round(this.colWidth(this.resolveSheet(sheet, si), cc) / 7.5),
      cellFormat: (sheet, rr, cc) => {
        const st = this.styleAt(this.resolveSheet(sheet, si), rr, cc);
        return { number: 'F2', currency: 'C0', accounting: 'C0', percent: 'P0', date: 'D1', longdate: 'D1', time: 'D9', datetime: 'D4', scientific: 'S2', text: '@' }[st.numFmt] ?? 'G';
      },
      usedCols: (sheet) => this.usedRange(this.resolveSheet(sheet, si)).cols,
      structRef: (table, spec) => {
        const rg = resolveStructRef(this, table, spec, here);
        if (!rg) return null;
        const range = rg.r1 !== rg.r2 || rg.c1 !== rg.c2;
        return { sheet: this.sheets[rg.si].name, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, range };
      },
      rowHidden: (sheet, row, manualToo, filteredToo = true) => {
        const s = this.sheets[this.resolveSheet(sheet, si)];
        if (!s) return false;
        if (manualToo && s.hiddenRows?.[row]) return true;
        if (!filteredToo) return false;
        if (s.filter?.hidden?.[row]) return true;
        return (s.tables ?? []).some((t) => t.filter?.hidden?.[row]);
      },
      cell: (sheet, r, c) => this.getValue(this.resolveSheet(sheet, si), r, c),
      range: (sheet, r1, c1, r2, c2) => {
        const s = this.resolveSheet(sheet, si);
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
    };
  }

  /** 값이 있는 영역 크기 {rows, cols} */
  usedRange(si) {
    if (this.usedCache.has(si)) return this.usedCache.get(si);
    let rows = 0;
    let cols = 0;
    for (const [k, cell] of this.sheets[si].cells) {
      if (!cell.raw && !cell.image) continue;
      const [r, c] = unkey(k);
      rows = Math.max(rows, r + 1);
      cols = Math.max(cols, c + 1);
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
    for (const k of this.sheets[si].cells.keys()) {
      const [r, c] = unkey(k);
      rows = Math.max(rows, r + 1);
      cols = Math.max(cols, c + 1);
    }
    const res = { rows, cols };
    this.extentCache.set(si, res);
    return res;
  }

  colWidth(si, c) { return this.sheets[si].colWidths[c] ?? DEFAULT_COL_WIDTH; }
  rowHeight(si, r) { return this.sheets[si].rowHeights[r] ?? DEFAULT_ROW_HEIGHT; }

  /** si 를 주면 그 시트의 버전만, 아니면 모든 시트의 버전을 올림 (피벗 원본 캐시가 씀) */
  invalidate(si) {
    this.version++;
    this.sheetVer ??= [];
    if (si === undefined || !this.sheets[si]) {
      this.sheetVerAll = (this.sheetVerAll ?? 0) + 1;
      this.caches = [];
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
    for (const s of aff) {
      this.sheetVer[s] = (this.sheetVer[s] ?? 0) + 1;
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
  invalidateEntries(entries) {
    if (entries.some((e) => e.t === 'all')) { this.invalidate(); return; }
    if (entries.some((e) => e.t === 'list' || e.t === 'sheet' || e.t === 'names' || e.t === 'rename')) {
      // 구조가 바뀐 시트와 그 시트를 참조하는 시트는 다시 계산
      for (const e of entries) if (e.t === 'sheet') for (const x of this.affected(e.si)) if (this.sheets[x]) this.sheets[x].fileValues = false;
      if (entries.some((e) => e.t === 'names')) this.sheetDeps().forEach((d, i) => { if (d.all) this.sheets[i].fileValues = false; });
      this.invalidateStructure();
    }
    const done = new Set();
    for (const e of entries) {
      if (e.t === 'cell' || (e.t === 'prop' && !CALC_NEUTRAL.has(e.prop))) {
        if (e.t === 'prop' && e.prop === 'tables') { this.deps = null; this.affectMemo.clear(); }
        if (!done.has(e.si)) { done.add(e.si); this.invalidate(e.si); }
      } else if (e.t === 'prop' && e.prop === 'merges') {
        this.spillState = null;
        this.usedCache.delete(e.si);
      }
    }
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
      if (tx.entries.length) {
        this.undoStack.push(tx);
        if (this.undoStack.length > 200) this.undoStack.shift();
        this.redoStack = [];
      }
      this.invalidateEntries(tx.entries);
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
  invalidateStructure() {
    this.version++;
    this.sheetVerAll = (this.sheetVerAll ?? 0) + 1;
    this.caches = [];
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
    if (e.t === 'cell') this.putCell(e.si, e.r, e.c, makeCell(e[side]));
    else if (e.t === 'list') { this.sheets = [...e[side].sheets]; this.names = e[side].names.map((n) => ({ ...n })); }
    else if (e.t === 'sheet') this.putSheet(e.si, e[side]);
    else if (e.t === 'names') this.names = e[side].map((n) => ({ ...n }));
    else if (e.t === 'rename') { if (this.sheets[e.si]) this.sheets[e.si].name = e[side]; }
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
    const k = key(r, c);
    if (cells.get(k)?.maybeArray || cell?.maybeArray) this.arrayList = null;
    if (cell) cells.set(k, cell);
    else cells.delete(k);
    // 새 수식이 다른 시트를 참조하면 의존 관계에 추가
    if (cell?.formula && cell.ast && this.deps?.[si] && this.addDeps(this.deps[si], si, cell.ast)) this.affectMemo.clear();
  }

  // ─────────── 셀 변경 ───────────
  /** 셀 전체 교체 (data = {raw, style, comment} 또는 null) */
  setCellData(si, r, c, data) {
    const before = cellData(this.getCell(si, r, c));
    const cell = makeCell(data);
    const after = cellData(cell);
    if (JSON.stringify(before) === JSON.stringify(after)) return;
    this.record({ t: 'cell', si, r, c, before, after });
    this.putCell(si, r, c, cell);
    this.invalidate(si);
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
        let ast = null;
        try { ast = parse(raw.slice(1)); } catch { /* 오류 수식 */ }
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
    const style = { ...(cur?.style || {}), ...patch };
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
    if (!manual && v === DEFAULT_ROW_HEIGHT) delete sheet.rowHeights[r];
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
    const moved = new Map();
    for (const [k, cell] of target.cells) {
      let [r, c] = unkey(k);
      const p = isRow ? r : c;
      if (count < 0 && p >= index && p < index - count) continue;
      const np = p >= index ? p + count : p;
      if (isRow) r = np; else c = np;
      moved.set(key(r, c), cell);
    }
    target.cells = moved;
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
        if (t.filter) nt.filter = { ...t.filter, hidden: shiftKeys(t.filter.hidden, index, count) };
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
        if (isRow) f.hidden = shiftKeys(f.hidden, index, count);
        else f.criteria = shiftKeys(f.criteria, index, count);
        target.filter = f;
      }
    }
    target.charts = target.charts.map((ch) => {
      const rg = adjustRange(ch.range, axis, index, count);
      return rg ? { ...ch, range: rg } : ch;
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
        const next = makeCell({ ...cellData(cell), raw });
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

  insertRows(si, index, count = 1) { this.shiftAxis(si, 'row', index, count); }
  deleteRows(si, index, count = 1) { this.shiftAxis(si, 'row', index, -count); }
  insertCols(si, index, count = 1) { this.shiftAxis(si, 'col', index, count); }
  deleteCols(si, index, count = 1) { this.shiftAxis(si, 'col', index, -count); }

  addSheet(name, at = this.sheets.length) {
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
    });
    this.sheets.forEach((sheet, i) => {
      if (!refs.has(i)) return;
      for (const [k, cell] of [...sheet.cells]) {
        if (!cell.formula) continue;
        const raw = renameSheetInFormula(cell.raw, old, newName);
        if (raw !== cell.raw) this.swapCell(i, k, makeCell({ ...cellData(cell), raw }));
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
      const c = compareValues(a.v, b.v);
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
  serialize() {
    return {
      version: 1,
      ...(this.vba ? { vba: this.vba } : {}),
      ...(this.names.length ? { names: this.names.map(({ _ast, _text, ...n }) => ({ ...n })) } : {}),
      sheets: this.sheets.map((s) => {
        const cells = {};
        for (const [k, cell] of s.cells) cells[k] = cellData(cell);
        const out = { name: s.name, cells };
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
    const total = data.sheets.reduce((n, s) => n + (s.cells instanceof Map ? s.cells.size : 0), 0) || 1;
    let done = 0;
    const sheets = [];
    for (const s of data.sheets) {
      const sheet = newSheet(s.name);
      if (s.cells instanceof Map) {
        // 파일에서 읽은 셀 Map: 새 Map 을 만들지 않고 그 자리에서 셀 객체로 바꿈 (큰 파일에서 훨씬 빠름)
        const cells = s.cells;
        let n = 0;
        for (const [k, d] of cells) {
          const cell = makeCell(d);
          if (cell) cells.set(k, cell);
          else cells.delete(k);
          if (++n % 50000 === 0) yield (done + n) / total;
        }
        done += n;
        sheet.cells = cells;
        s.cells = null; // 이제 통합 문서가 소유
      } else {
        for (const [k, d] of Object.entries(s.cells || {})) {
          const cell = makeCell(d);
          if (cell) sheet.cells.set(k, cell);
        }
      }
      for (const p of SHEET_PROPS) if (s[p] !== undefined && s[p] !== null) sheet[p] = structuredClone(s[p]);
      sheet.freeze = { rows: 0, cols: 0, ...(s.freeze || {}) };
      if (s._sid) { sheet._sid = s._sid; sheet._ev = s._ev; } // 자동 저장 기록 (바뀐 시트만 다시 저장)
      sheets.push(sheet);
    }
    this.vba = data.vba ?? null; // .xlsm 의 매크로(vbaProject.bin, base64) — 실행하지 않고 보존만 함
    this.names = (data.names ?? []).map((n) => ({ ...n }));
    this.sheets = sheets.length ? sheets : [newSheet('Sheet1')];
    this.invalidate();
    data.sheets.forEach((s, i) => { if (s.fileValues && this.sheets[i]) this.sheets[i].fileValues = true; });
  }

  load(data) {
    this.restore(data);
    this.undoStack = [];
    this.redoStack = [];
  }

  /** 큰 문서: 화면이 멈추지 않도록 나눠서 불러옴. onProgress(0~1) */
  async loadAsync(data, onProgress) {
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
