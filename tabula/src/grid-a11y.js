// One logical accessibility grid for the four visual panes. The existing textarea
// retains DOM focus and IME handling; it controls this grid while navigating.
import { cellName, colToName, MAX_COLS, isError } from './formula.js';
import { normalizeWorksheetRowLimit } from './row-limits.js';
import { formatValue } from './format.js';
import { isProtected, isLockedStyle } from './protect.js';

const MAX_CELLS = 400;
const short = (value, max = 300) => { const text = String(value ?? ''); return text.length > max ? text.slice(0, max) + '… (일부 표시)' : text; };
const inside = (r, c, rg) => r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
const axisLimit = (axis, limit) => Number.isSafeInteger(axis?.max) && axis.max >= 0 ? Math.min(axis.max, limit) : limit;
const boundIndex = (value, count) => Math.max(0, Math.min(Number.isFinite(value) ? Math.floor(value) : 0, count - 1));
const gridAttribute = (node, key, value) => { if (node.getAttribute(key) !== value) node.setAttribute(key, value); };

export function selectionDescription(state) {
  const { selKind, special } = state, rowLimit = normalizeWorksheetRowLimit(state.rowLimit);
  const sel = { r1: boundIndex(state.sel.r1, rowLimit), r2: boundIndex(state.sel.r2, rowLimit), c1: boundIndex(state.sel.c1, MAX_COLS), c2: boundIndex(state.sel.c2, MAX_COLS) };
  if (special) return `비연속 ${special.length}개 셀 선택`;
  if (selKind === 'all' || (sel.r1 === 0 && sel.c1 === 0 && sel.r2 === rowLimit - 1 && sel.c2 === MAX_COLS - 1)) return '워크시트 전체 선택';
  if (selKind === 'rows') return `${sel.r1 + 1}행부터 ${sel.r2 + 1}행까지 선택`;
  if (selKind === 'cols') return `${colToName(sel.c1)}열부터 ${colToName(sel.c2)}열까지 선택`;
  return sel.r1 === sel.r2 && sel.c1 === sel.c2 ? '' : `${cellName(sel.r1, sel.c1)}:${cellName(sel.r2, sel.c2)} 범위 선택`;
}

/** Visits only a bounded window. Hidden million-row regions never cause a scan. */
export function gridCoordinates(active, windows, rows, cols, max = MAX_CELLS, rowLimit = normalizeWorksheetRowLimit(rows?.max)) {
  const points = new Map(), rowCount = axisLimit(rows, normalizeWorksheetRowLimit(rowLimit)), colCount = axisLimit(cols, MAX_COLS);
  if (!rowCount || !colCount) return [];
  const focus = { r: boundIndex(active.r, rowCount), c: boundIndex(active.c, colCount) };
  const add = (r, c, force = false) => {
    if (r < 0 || c < 0 || r >= rowCount || c >= colCount || points.size >= max) return;
    if (!force && (!rows.size(r) || !cols.size(c))) return;
    points.set(`${r},${c}`, { r, c });
  };
  add(focus.r, focus.c, true);
  // Active-cell neighbours come first, including when the visual viewport is elsewhere.
  for (let r = focus.r - 2; r <= focus.r + 2; r++) for (let c = focus.c - 2; c <= focus.c + 2; c++) add(r, c);
  let visits = 0;
  for (const w of windows) {
    if (!w) continue;
    const r2 = Math.min(w.r2, rowCount - 1), c2 = Math.min(w.c2, colCount - 1);
    for (let r = Math.max(0, w.r1); r <= r2 && points.size < max && visits < max * 4; r++) {
      visits++; if (!rows.size(r)) continue;
      for (let c = Math.max(0, w.c1); c <= c2 && points.size < max && visits < max * 4; c++) { visits++; add(r, c); }
    }
  }
  return [...points.values()];
}

/** Index only candidate coordinates, so thousands of merges aren't searched per cell. */
export function gridMerges(points, merges) {
  const byRow = new Map(), found = new Map();
  for (const point of points) { if (!byRow.has(point.r)) byRow.set(point.r, []); byRow.get(point.r).push(point.c); }
  const rowKeys = [...byRow.keys()].sort((a, b) => a - b);
  const lower = (values, target) => { let a = 0, b = values.length; while (a < b) { const mid = (a + b) >> 1; if (values[mid] < target) a = mid + 1; else b = mid; } return a; };
  for (const values of byRow.values()) values.sort((a, b) => a - b);
  for (const merge of merges) {
    for (let i = lower(rowKeys, merge.r1); i < rowKeys.length && rowKeys[i] <= merge.r2; i++) {
      const r = rowKeys[i], values = byRow.get(r);
      for (let j = lower(values, merge.c1); j < values.length && values[j] <= merge.c2; j++) {
        const key = `${r},${values[j]}`; if (!found.has(key)) found.set(key, merge);
      }
    }
  }
  return found;
}

export function gridCellInfo(state, r, c, merge = null) {
  const { wb, si } = state, sheet = wb.sheets[si];
  const rowLimit = normalizeWorksheetRowLimit(state.rowLimit);
  if (merge) { merge = { ...merge, r2: Math.min(merge.r2, rowLimit - 1), c2: Math.min(merge.c2, MAX_COLS - 1) }; r = merge.r1; c = merge.c1; }
  const cell = wb.getCell(si, r, c), style = wb.styleAt(si, r, c), value = wb.getValue(si, r, c);
  const readonly = !!state.readonly || !!wb.props?.markedFinal || (isProtected(sheet) && isLockedStyle(style));
  const hiddenFormula = isProtected(sheet) && style.hideFormula;
  const formula = cell?.formula && !hiddenFormula ? short(cell.raw, 200) : '';
  let text = value == null || value === '' ? '빈 셀' : isError(value) ? `오류 ${value.code}` : short(formatValue(value, style, wb.date1904).text);
  if (style.checkbox) text = `확인란 ${value === true ? '선택됨' : '선택 안 됨'}`;
  if (cell?.image || value?.type === 'image') text = `셀 이미지${value?.alt ? ', ' + short(value.alt) : ''}`;
  if (!text) text = '표시값 없음';
  const label = [cellName(r, c), text];
  if (formula) label.push(`수식 ${formula}`);
  if (merge) label.push(`병합 ${cellName(merge.r1, merge.c1)}:${cellName(merge.r2, merge.c2)}`);
  if (readonly) label.push('읽기 전용');
  if (cell?.comment) label.push('메모 있음');
  if (cell?.link) label.push('하이퍼링크 있음');
  return { r, c, label: label.join(', '), readonly, rowSpan: merge ? merge.r2 - merge.r1 + 1 : 1, colSpan: merge ? merge.c2 - merge.c1 + 1 : 1 };
}

export class GridAccessibility {
  constructor(view, editor) {
    this.view = view; this.editor = editor; this.nodes = new Map(); this.specialRef = null; this.specialSet = null; this.infoCache = new Map();
    const doc = editor.ownerDocument;
    this.grid = doc.createElement('div'); this.grid.id = 'accessibleGrid'; this.grid.className = 'grid-a11y';
    this.grid.setAttribute('role', 'grid'); this.setRowLimit(view.rowLimit); this.grid.setAttribute('aria-colcount', String(MAX_COLS));
    this.grid.setAttribute('aria-multiselectable', 'true');
    this.help = doc.createElement('div'); this.help.id = 'gridKeyboardHelp'; this.help.className = 'grid-a11y-help';
    this.help.textContent = '방향키로 셀 이동, Shift와 방향키로 범위 선택, F2로 편집, Enter로 확정, Escape로 편집 취소. 화면의 일부 셀을 제공하며 이동하면 갱신됩니다.';
    view.wrap.append(this.grid, this.help);
    editor.setAttribute('aria-controls', this.grid.id); editor.setAttribute('aria-describedby', this.help.id);
    this.grid.addEventListener('focusin', (event) => {
      const node = event.target.closest('[data-a11y-cell]');
      if (node) {
        const r = Number(node.dataset.r), c = Number(node.dataset.c), limit = normalizeWorksheetRowLimit(view.host.state().rowLimit ?? view.rowLimit);
        if (Number.isInteger(r) && r >= 0 && r < axisLimit(view.rows, limit) && Number.isInteger(c) && c >= 0 && c < axisLimit(view.cols, MAX_COLS)) view.host.onAccessibleCellFocus?.(r, c);
      }
    });
  }
  setRowLimit(value) { gridAttribute(this.grid, 'aria-rowcount', String(normalizeWorksheetRowLimit(value))); }

  update() {
    const view = this.view, hostState = view.host.state(), rowLimit = normalizeWorksheetRowLimit(hostState.rowLimit ?? view.rowLimit);
    this.setRowLimit(rowLimit);
    const state = { ...hostState, rowLimit }, { wb, si, sel } = state;
    const rowCount = axisLimit(view.rows, rowLimit), colCount = axisLimit(view.cols, MAX_COLS);
    const active = { r: boundIndex(state.active.r, rowCount), c: boundIndex(state.active.c, colCount) };
    if (!wb.sheets[si] || !view.rows || !view.cols) return;
    const sheet = wb.sheets[si], editing = !!state.editing, object = !!state.chartSel;
    const windows = view.panes.map(p => p.win);
    const valueKey = `${si}:${wb.version}:${!!state.readonly}:${!!wb.props?.markedFinal}:${rowLimit}:${rowCount}:${colCount}`;
    const valueChanged = this.workbook !== wb || this.valueKey !== valueKey;
    const signature = `${valueKey}:${active.r},${active.c}:${sel.r1},${sel.c1},${sel.r2},${sel.c2}:${state.selKind}:${editing}:${object}:` + windows.map(w => w ? `${w.r1},${w.c1},${w.r2},${w.c2}` : '').join(';');
    if (!valueChanged && signature === this.signature && this.specialRef === state.special) return;
    this.signature = signature;
    if (valueChanged) { this.workbook = wb; this.valueKey = valueKey; this.infoCache.clear(); }
    const description = selectionDescription(state);
    this.grid.setAttribute('aria-label', `${sheet.name} 워크시트`);
    this.grid.setAttribute('aria-readonly', String(!!state.readonly || !!wb.props?.markedFinal));
    if (description) this.grid.setAttribute('aria-description', description); else this.grid.removeAttribute('aria-description');
    if (this.specialRef !== state.special) {
      this.specialRef = state.special;
      // Do not scan/allocate a million selected coordinates on the keyboard hot path.
      this.specialSet = state.special && state.special.length <= 10000 ? new Set(state.special.map(([r, c]) => `${r},${c}`)) : null;
    }
    const points = gridCoordinates(active, windows, view.rows, view.cols, MAX_CELLS, rowLimit);
    const merges = gridMerges(points, sheet.merges), seen = new Set(), rows = new Map();
    let activeNode = null;
    for (const point of points) {
      const sourceMerge = merges.get(`${point.r},${point.c}`);
      const merge = sourceMerge ? { ...sourceMerge, r2: Math.min(sourceMerge.r2, rowCount - 1), c2: Math.min(sourceMerge.c2, colCount - 1) } : null;
      const r = merge?.r1 ?? point.r, c = merge?.c1 ?? point.c, key = `${r},${c}`;
      if (seen.has(key)) continue; seen.add(key);
      let info = this.infoCache.get(key);
      if (!info) { info = gridCellInfo(state, r, c, merge); this.infoCache.set(key, info); }
      const id = `a11y-cell-${si}-${r}-${c}`;
      let node = this.nodes.get(key);
      if (!node) { node = this.grid.ownerDocument.createElement('div'); node.setAttribute('role', 'gridcell'); node.tabIndex = -1; node.dataset.a11yCell = 'true'; this.nodes.set(key, node); }
      if (node.id !== id) { node.id = id; node.dataset.r = r; node.dataset.c = c; }
      gridAttribute(node, 'aria-colindex', String(c + 1)); gridAttribute(node, 'aria-rowindex', String(r + 1));
      gridAttribute(node, 'aria-readonly', String(info.readonly)); gridAttribute(node, 'aria-rowspan', String(info.rowSpan)); gridAttribute(node, 'aria-colspan', String(info.colSpan));
      const selected = state.special ? (this.specialSet ? this.specialSet.has(key) : null) : inside(r, c, sel);
      if (selected === null) node.removeAttribute('aria-selected'); else gridAttribute(node, 'aria-selected', String(selected));
      const isActive = merge ? inside(active.r, active.c, merge) : r === active.r && c === active.c;
      const label = info.label + (isActive && description ? `, ${description}` : '');
      if (node.textContent !== label) node.textContent = label;
      if (isActive) activeNode = node;
      if (!rows.has(r)) rows.set(r, []); rows.get(r).push(node);
    }
    const ordered = [...rows.keys()].sort((a, b) => a - b).map(r => [r, rows.get(r).sort((a,b) => Number(a.dataset.c)-Number(b.dataset.c))]);
    const structure = ordered.map(([r, cells]) => `${r}:` + cells.map(cell => cell.id).join(',')).join(';');
    // Selection alone must not detach/reinsert 400 unchanged accessible cells.
    if (structure !== this.structure) {
      const rowNodes = [];
      for (const [r, cells] of ordered) {
        const row = this.grid.ownerDocument.createElement('div'); row.setAttribute('role', 'row'); row.setAttribute('aria-rowindex', String(r + 1));
        row.append(...cells); rowNodes.push(row);
      }
      this.grid.replaceChildren(...rowNodes); this.structure = structure;
    }
    for (const key of this.nodes.keys()) if (!seen.has(key)) { this.nodes.delete(key); this.infoCache.delete(key); }
    if (activeNode && !editing && !object) {
      this.editor.setAttribute('aria-activedescendant', activeNode.id); this.grid.setAttribute('aria-activedescendant', activeNode.id);
      this.editor.setAttribute('aria-label', '워크시트 셀 탐색');
    } else {
      this.editor.removeAttribute('aria-activedescendant'); this.grid.removeAttribute('aria-activedescendant');
      this.editor.setAttribute('aria-label', editing ? `${cellName(active.r, active.c)} 셀 편집` : '워크시트 개체 선택');
    }
    this.editor.setAttribute('aria-readonly', String(!!activeNode && activeNode.getAttribute('aria-readonly') === 'true'));
  }
}
