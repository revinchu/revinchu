// 통합 문서 모델: 시트 · 셀 · 재계산 · 실행 취소 · 행/열 구조 변경
import {
  parse, evaluateFormula, ERR, compareValues, isError, autoFormatFor,
  adjustFormulaForStructure, renameSheetInFormula, shiftFormula,
} from './formula.js';
import { parseInput } from './format.js';

export const DEFAULT_COL_WIDTH = 64;
export const DEFAULT_ROW_HEIGHT = 20;

const key = (r, c) => `${r},${c}`;
const unkey = (k) => k.split(',').map(Number);

function cleanStyle(style) {
  if (!style) return undefined;
  const out = {};
  for (const [k, v] of Object.entries(style)) if (v !== undefined && v !== null) out[k] = v;
  return Object.keys(out).length ? out : undefined;
}

/** 저장 형태 {raw, style, comment} → 계산용 셀 객체 */
export function makeCell(data) {
  if (!data) return null;
  const cell = { raw: data.raw ?? '' };
  const style = cleanStyle(data.style);
  if (style) cell.style = style;
  if (data.comment) cell.comment = data.comment;
  if (cell.raw.startsWith('=') && cell.raw.length > 1 && style?.numFmt !== 'text') {
    cell.formula = true;
    try {
      cell.ast = parse(cell.raw.slice(1));
    } catch (e) {
      cell.ast = null;
      cell.parseError = e.message;
    }
  } else if (style?.numFmt === 'text') {
    cell.v = cell.raw === '' ? null : cell.raw;
  } else {
    cell.v = parseInput(cell.raw).value;
  }
  if (!cell.raw && !cell.style && !cell.comment) return null;
  return cell;
}

export function cellData(cell) {
  if (!cell) return null;
  const d = { raw: cell.raw };
  if (cell.style) d.style = { ...cell.style };
  if (cell.comment) d.comment = cell.comment;
  return d;
}

function newSheet(name) {
  return {
    name, cells: new Map(), colWidths: {}, rowHeights: {}, merges: [], cond: [],
    colStyles: {}, rowStyles: {}, allStyle: null, hiddenRows: {}, hiddenCols: {}, rowManual: {},
    freeze: { rows: 0, cols: 0 }, filter: null, charts: [], pivot: null,
  };
}

/** 시트의 부가 속성 (셀 외) — 저장/복원/복제용 */
const SHEET_PROPS = ['colWidths', 'rowHeights', 'merges', 'cond', 'colStyles', 'rowStyles', 'allStyle',
  'hiddenRows', 'hiddenCols', 'rowManual', 'freeze', 'filter', 'charts', 'pivot'];

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
    this.cache = new Map();
    this.evaluating = new Set();
    this.usedCache = new Map();
    this.extentCache = new Map();
    this.undoStack = [];
    this.redoStack = [];
    this.tx = null;
    this.listeners = new Set();
    this.depth = 0;
    this.warming = false;
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
    const cell = this.getCell(si, r, c);
    if (!cell) return null;
    if (!cell.formula) return cell.v ?? null;
    const k = `${si}:${r},${c}`;
    if (this.cache.has(k)) return this.cache.get(k);
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
    try {
      v = evaluateFormula(cell.ast, this.ctxFor(si));
    } catch (e) {
      if (e instanceof RangeError) throw DEEP;
      throw e;
    } finally {
      this.depth--;
      this.evaluating.delete(k);
    }
    this.cache.set(k, v);
    return v;
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

  ctxFor(si) {
    return {
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
      usedRows: (sheet) => this.usedRange(this.resolveSheet(sheet, si)).rows,
    };
  }

  /** 값이 있는 영역 크기 {rows, cols} */
  usedRange(si) {
    if (this.usedCache.has(si)) return this.usedCache.get(si);
    let rows = 0;
    let cols = 0;
    for (const [k, cell] of this.sheets[si].cells) {
      if (!cell.raw) continue;
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

  invalidate() {
    this.cache.clear();
    this.usedCache.clear();
    this.extentCache.clear();
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
      for (const e of tx.entries) if (e.t === 'all') e.after = this.serialize();
      if (tx.entries.length) {
        this.undoStack.push(tx);
        if (this.undoStack.length > 200) this.undoStack.shift();
        this.redoStack = [];
      }
      this.invalidate();
      this.emit();
    }
    return result;
  }

  record(entry) {
    if (this.tx) this.tx.entries.push(entry);
  }

  snapshotAll() {
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
    this.invalidate();
    this.emit();
    return tx.meta;
  }

  redo() {
    const tx = this.redoStack.pop();
    if (!tx) return null;
    for (const e of tx.entries) this.applyEntry(e, 'after');
    this.undoStack.push(tx);
    this.invalidate();
    this.emit();
    return tx.meta;
  }

  applyEntry(e, side) {
    if (e.t === 'cell') this.putCell(e.si, e.r, e.c, makeCell(e[side]));
    else if (e.t === 'all') this.restore(e[side]);
    else if (e.t === 'colWidth') this.sheets[e.si].colWidths = { ...e[side] };
    else if (e.t === 'rowHeight') {
      this.sheets[e.si].rowHeights = { ...e[side] };
      if (e.beforeManual) this.sheets[e.si].rowManual = { ...(side === 'before' ? e.beforeManual : e.afterManual) };
    }
  }

  putCell(si, r, c, cell) {
    const cells = this.sheets[si].cells;
    if (cell) cells.set(key(r, c), cell);
    else cells.delete(key(r, c));
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
    this.invalidate();
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
    this.setCellData(si, r, c, { raw, style, comment: cur?.comment });
  }

  setStyle(si, r, c, patch) {
    const cur = this.getCell(si, r, c);
    const style = { ...(cur?.style || {}), ...patch };
    this.setCellData(si, r, c, { raw: cur?.raw ?? '', style, comment: cur?.comment });
  }

  setComment(si, r, c, comment) {
    const cur = this.getCell(si, r, c);
    this.setCellData(si, r, c, { raw: cur?.raw ?? '', style: cur?.style, comment: comment || undefined });
  }

  clearRange(si, r1, c1, r2, c2, what = 'contents') {
    for (const [k, cell] of [...this.sheets[si].cells]) {
      const [r, c] = unkey(k);
      if (r < r1 || r > r2 || c < c1 || c > c2) continue;
      if (what === 'all') this.setCellData(si, r, c, null);
      else if (what === 'formats') this.setCellData(si, r, c, { raw: cell.raw, comment: cell.comment });
      else if (what === 'comments') this.setCellData(si, r, c, { raw: cell.raw, style: cell.style });
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
    this.snapshotAll();
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
    this.snapshotAll();
    const map = axis === 'row' ? this.sheets[si].hiddenRows : this.sheets[si].hiddenCols;
    for (const i of indices) { if (hidden) map[i] = true; else delete map[i]; }
  }

  setSheetProp(si, prop, value) {
    this.snapshotAll();
    this.sheets[si][prop] = value;
    this.invalidate();
  }

  // ─────────── 구조 변경 ───────────
  /** axis: 'row'|'col', count>0 삽입, count<0 삭제 */
  shiftAxis(si, axis, index, count) {
    this.snapshotAll();
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
    target.cond = target.cond.map((rule) => adjustRange(rule, axis, index, count)).filter(Boolean);
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
    for (const sh of this.sheets) {
      if (sh.pivot && sh.pivot.source.toLowerCase() === target.name.toLowerCase()) {
        const rg = adjustRange(sh.pivot.range, axis, index, count);
        if (rg) sh.pivot = { ...sh.pivot, range: rg };
      }
    }

    for (const sheet of this.sheets) {
      for (const [k, cell] of sheet.cells) {
        if (!cell.formula) continue;
        const raw = adjustFormulaForStructure(cell.raw, {
          targetSheet: target.name, hostSheet: sheet.name, axis, index, count,
        });
        if (raw !== cell.raw) sheet.cells.set(k, makeCell({ ...cellData(cell), raw }));
      }
    }
    this.invalidate();
  }

  insertRows(si, index, count = 1) { this.shiftAxis(si, 'row', index, count); }
  deleteRows(si, index, count = 1) { this.shiftAxis(si, 'row', index, -count); }
  insertCols(si, index, count = 1) { this.shiftAxis(si, 'col', index, count); }
  deleteCols(si, index, count = 1) { this.shiftAxis(si, 'col', index, -count); }

  addSheet(name, at = this.sheets.length) {
    this.snapshotAll();
    let n = this.sheets.length + 1;
    let nm = name;
    while (!nm || this.sheetIndexByName(nm) >= 0) nm = `Sheet${n++}`;
    this.sheets.splice(at, 0, newSheet(nm));
    this.invalidate();
    return at;
  }

  deleteSheet(si) {
    if (this.sheets.length <= 1) return false;
    this.snapshotAll();
    this.sheets.splice(si, 1);
    this.invalidate();
    return true;
  }

  renameSheet(si, newName) {
    newName = newName.trim();
    if (!newName || /[\\/?*[\]:]/.test(newName) || newName.length > 31) return false;
    const dup = this.sheetIndexByName(newName);
    if (dup >= 0 && dup !== si) return false;
    this.snapshotAll();
    const old = this.sheets[si].name;
    this.sheets[si].name = newName;
    for (const sh of this.sheets) {
      if (sh.pivot && sh.pivot.source.toLowerCase() === old.toLowerCase()) sh.pivot = { ...sh.pivot, source: newName };
    }
    for (const sheet of this.sheets) {
      for (const [k, cell] of sheet.cells) {
        if (!cell.formula) continue;
        const raw = renameSheetInFormula(cell.raw, old, newName);
        if (raw !== cell.raw) sheet.cells.set(k, makeCell({ ...cellData(cell), raw }));
      }
    }
    this.invalidate();
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
    this.snapshotAll();
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
    this.invalidate();
  }

  unmerge(si, r1, c1, r2, c2) {
    const sheet = this.sheets[si];
    const keep = sheet.merges.filter((m) => !(m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1));
    if (keep.length === sheet.merges.length) return;
    this.snapshotAll();
    sheet.merges = keep;
  }

  addCondRule(si, rule) {
    this.snapshotAll();
    this.sheets[si].cond.push(rule);
    this.invalidate();
  }

  clearCondRules(si, range) {
    this.snapshotAll();
    const sheet = this.sheets[si];
    sheet.cond = range
      ? sheet.cond.filter((c) => !(c.r1 <= range.r2 && c.r2 >= range.r1 && c.c1 <= range.c2 && c.c2 >= range.c1))
      : [];
    this.invalidate();
  }

  // ─────────── 저장 / 불러오기 ───────────
  serialize() {
    return {
      version: 1,
      sheets: this.sheets.map((s) => {
        const cells = {};
        for (const [k, cell] of s.cells) cells[k] = cellData(cell);
        const out = { name: s.name, cells };
        for (const p of SHEET_PROPS) out[p] = structuredClone(s[p]);
        return out;
      }),
    };
  }

  restore(data) {
    this.sheets = data.sheets.map((s) => {
      const sheet = newSheet(s.name);
      for (const [k, d] of Object.entries(s.cells || {})) {
        const cell = makeCell(d);
        if (cell) sheet.cells.set(k, cell);
      }
      for (const p of SHEET_PROPS) if (s[p] !== undefined && s[p] !== null) sheet[p] = structuredClone(s[p]);
      sheet.freeze = { rows: 0, cols: 0, ...(s.freeze || {}) };
      return sheet;
    });
    if (!this.sheets.length) this.sheets = [newSheet('Sheet1')];
    this.invalidate();
  }

  load(data) {
    this.restore(data);
    this.undoStack = [];
    this.redoStack = [];
  }
}
