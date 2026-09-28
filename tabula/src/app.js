// Tabula 메인: 상태 · 선택 · 편집 · 키보드/마우스 · 명령 (그리기는 view.js)
import { Workbook, cellData, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import {
  cellName, colToName, parseRangeName, parse, shiftFormula, listRefs, normalizeFormula, tokenize,
  FUNCTION_NAMES, isError, quoteSheetName, MAX_ROWS, MAX_COLS,
} from './formula.js';
import {
  formatValue, displayedDecimals, parseInput, formatCode, styleForCode, codeOfStyle, adjustCodeDecimals, formatGeneral,
} from './format.js';
import { buildRibbon, FONTS, FONT_SIZES, TABS } from './ribbon.js';
import {
  el, hydrateIcons, toast, openMenu, closeMenus, isMenuOpen, openDialog, alertDialog,
  formDialog, setMenuCloseHandler, setDialogCloseHandler, isDialogOpen,
} from './ui.js';
import { FUNC_INFO, CATEGORIES } from './funcinfo.js';
import { makeSeries } from './series.js';
import { parseDelimited, toDelimited, guessDelimiter } from './csv.js';
import { SAMPLES } from './samples.js';
import { GridView, DEFAULT_FONT, DEFAULT_SIZE, measureText, fontStack } from './view.js';
import { readXlsx, writeXlsx, xlsxOverflow } from './xlsx.js';
import { CHART_TYPES, chartData, renderChartSvg } from './chart.js';
import { buildPivot, AGGREGATES, pivotSourceData, resolvePivot, itemText } from './pivot.js';
import { server } from './storage.js';
import { ICONS } from './icons.js';
import {
  CELL_OPS, TEXT_OPS, DATE_PERIODS, ICON_SETS, ICON_SVG, VISUAL_TYPES, iconSetById, describeCond,
} from './condfmt.js';
import {
  TABLE_STYLES, DEFAULT_TABLE_STYLE, TOTAL_FUNCS, tableAt, tableCellStyle, tableFilterRange, dataTop, dataBottom, uniqueNames,
  nextTableName, columnNames, expansionFor, validTableName, findTable, resolveStructRef,
} from './tables.js';
import { splitDelimited, splitFixed, suggestBreaks, parseDateOrder, convertPart, DATE_ORDERS } from './textsplit.js';
import {
  VALIDATION_TYPES, VALIDATION_OPS, validationAt, checkValidation, listItems, describeRule, subtractRange, invalidCells,
} from './validation.js';
import { OBJECT_PROPS, OBJECT_LABEL, SHAPE_KINDS, newShape, findObject, shapeSvg } from './shapes.js';
import { extractVbaModules, fromBase64 } from './vba.js';
import {
  CATEGORIES as FMT_CATEGORIES, CURRENCY_SYMBOLS, NEGATIVE_STYLES, DATE_TYPES, TIME_TYPES, FRACTION_TYPES, SPECIAL_TYPES, CUSTOM_LIST, buildCode, describeCode,
} from './fmtpresets.js';

const $ = (id) => document.getElementById(id);
const dom = {
  view: $('gridView'), editor: $('cellEditor'), ac: $('autocomplete'),
  nameBox: $('nameBox'), formula: $('formulaInput'), formulaRow: $('formulaRow'),
  fxCancel: $('fxCancel'), fxEnter: $('fxEnter'),
  status: $('statusMode'), stats: $('statusStats'), sheetTabs: $('sheetTabs'),
  zoomSlider: $('zoomSlider'), zoomLabel: $('zoomLabel'), title: $('docTitle'), saveState: $('saveState'), search: $('searchBox'),
  autosave: $('autosaveToggle'), autosaveLabel: $('autosaveLabel'), fileInput: $('fileInput'), tip: $('commentTip'),
  undoBtn: $('tbUndo'), redoBtn: $('tbRedo'), printArea: $('printArea'),
};

const STORAGE_KEY = 'tabula.workbook.v1';
const REF_COLORS = ['#2f6fd6', '#d13438', '#8a3fd1', '#0f8a3c', '#c75a00', '#0093b8', '#c2187a'];
const BIG_AREA = 200000;

// ───────────────────────── 상태 ─────────────────────────
const view = { showGrid: true, printGrid: false, showFormulas: false, showHeaders: true, showFormulaBar: true, zoom: 100 };
let wb;
let si = 0;
let docName = '통합 문서1';
let autosave = true;
let dirty = false;
let active = { r: 0, c: 0 };
let anchor = { r: 0, c: 0 };
let focusCell = { r: 0, c: 0 };
let sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
let selKind = 'cells'; // cells | cols | rows | all
const sheetSel = new Map();
let editing = null; // { r, c, mode:'enter'|'edit', original, point, fromBar }
let tabStartCol = null;
let clip = null;
let painter = null;
let drag = null;
let pendingKey = null;
let ac = null;
let editRefs = [];
let fillPreview = null;
let chartSel = null; // 선택한 그림 개체(차트·그림·도형) id
let objClip = null; // 복사한 그림 개체
let drawKind = null; // 그릴 도형 종류 (삽입 → 도형)
let circles = null; // 잘못된 데이터 표시
let dvPromptEl = null;
let lastFill = '#ffff00';
let lastFont = '#ff0000';
let lastBorder = 'bottom';
let ribbon;
let gv;
const serverState = { saving: false, error: null, savedAt: null };

// ───────────────────────── 유틸 ─────────────────────────
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const norm = (a, b) => ({ r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) });
const isSingle = (rg) => rg.r1 === rg.r2 && rg.c1 === rg.c2;
const meta = () => ({ si, sel: { ...sel }, active: { ...active }, selKind });
const escapeHtml = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));
const sheet = () => wb.sheets[si];
const range = (a, b) => Array.from({ length: Math.max(0, b - a + 1) }, (_, i) => a + i);

const styleAt = (r, c) => wb.styleAt(si, r, c);
const valueAt = (r, c) => wb.getValue(si, r, c);
const isEmptyAt = (r, c) => !wb.getCell(si, r, c)?.raw;
const filterHidden = (r) => !!sheet().filter?.hidden?.[r] || (sheet().tables ?? []).some((t) => t.filter?.hidden?.[r]);

function expandMerges(rg) {
  const out = { ...rg };
  for (let changed = true; changed;) {
    changed = false;
    for (const m of wb.mergesIn(si, out.r1, out.c1, out.r2, out.c2)) {
      if (m.r1 < out.r1 || m.c1 < out.c1 || m.r2 > out.r2 || m.c2 > out.c2) {
        out.r1 = Math.min(out.r1, m.r1); out.c1 = Math.min(out.c1, m.c1);
        out.r2 = Math.max(out.r2, m.r2); out.c2 = Math.max(out.c2, m.c2);
        changed = true;
      }
    }
  }
  return out;
}

/** 범위 안 셀 좌표 (필터로 숨겨진 행 제외). 매우 큰 범위는 데이터가 있는 영역으로 제한 */
function* cellsIn(rg, { includeFiltered = false } = {}) {
  let { r2, c2 } = rg;
  const { r1, c1 } = rg;
  if ((r2 - r1 + 1) * (c2 - c1 + 1) > BIG_AREA) {
    const ext = wb.extent(si);
    r2 = Math.min(r2, Math.max(r1, ext.rows - 1));
    c2 = Math.min(c2, Math.max(c1, ext.cols - 1));
  }
  for (let r = r1; r <= r2; r++) {
    if (!includeFiltered && filterHidden(r)) continue;
    for (let c = c1; c <= c2; c++) yield [r, c];
  }
}

/** 데이터가 있는 범위와 교차 */
function usedClip(rg) {
  const u = wb.usedRange(si);
  return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(rg.r1, u.rows - 1)), c2: Math.min(rg.c2, Math.max(rg.c1, u.cols - 1)) };
}

// ───────────────────────── 선택 영역 ─────────────────────────
function setSel(rg, kind = 'cells') {
  sel = expandMerges(rg);
  selKind = kind;
}

function growTo(r, c) {
  if (gv.ensureExtentFor(r, c)) gv.updateSizer();
}

function selectCell(r, c, { keepTab = false, scroll = true } = {}) {
  r = clamp(r, 0, MAX_ROWS - 1);
  c = clamp(c, 0, MAX_COLS - 1);
  const m = wb.mergeAt(si, r, c);
  if (m) { r = m.r1; c = m.c1; }
  active = { r, c };
  anchor = { r, c };
  focusCell = { r, c };
  setSel({ r1: r, c1: c, r2: r, c2: c });
  if (!keepTab) tabStartCol = null;
  growTo(r, c);
  if (scroll) gv.ensureVisible(r, c);
  updateSelectionUI();
}

function extendTo(r, c, { scroll = true } = {}) {
  r = clamp(r, 0, MAX_ROWS - 1);
  c = clamp(c, 0, MAX_COLS - 1);
  focusCell = { r, c };
  const rg = norm(anchor, focusCell);
  const fullRows = rg.r1 === 0 && rg.r2 === MAX_ROWS - 1;
  const fullCols = rg.c1 === 0 && rg.c2 === MAX_COLS - 1;
  setSel(rg, fullRows && fullCols ? 'all' : fullRows ? 'cols' : fullCols ? 'rows' : 'cells');
  tabStartCol = null;
  growTo(r, c);
  if (scroll) gv.ensureVisible(r, c);
  updateSelectionUI();
}

function selectRange(rg, kind = 'cells', act = { r: rg.r1, c: rg.c1 }) {
  active = { ...act };
  anchor = { ...act };
  focusCell = { r: rg.r2, c: rg.c2 };
  setSel(rg, kind);
  updateSelectionUI();
}

function selectAll() {
  selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, 'all', { ...active });
}
function selectCols(c1, c2, act) {
  selectRange({ r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) }, 'cols', act ?? { r: gv.firstVisibleRow(), c: c1 });
}
function selectRows(r1, r2, act) {
  selectRange({ r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 }, 'rows', act ?? { r: r1, c: 0 });
}

const inSel = (r, c) => r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
const selIsActiveOnly = () => {
  const m = wb.mergeAt(si, active.r, active.c);
  return isSingle(sel) || (m && m.r1 === sel.r1 && m.c1 === sel.c1 && m.r2 === sel.r2 && m.c2 === sel.c2);
};

function updateSelectionUI() {
  const selObj = chartSel ? findObject(sheet(), chartSel) : null;
  if (document.activeElement !== dom.nameBox) dom.nameBox.value = selObj ? (selObj.obj.name || OBJECT_LABEL[selObj.prop]) : cellName(active.r, active.c);
  if (!editing) dom.formula.value = chartSel ? '' : wb.getCell(si, active.r, active.c)?.raw ?? '';
  dom.fxCancel.disabled = !editing;
  dom.fxEnter.disabled = !editing;
  gv.renderSelection();
  positionEditor();
  updateStats();
  updateRibbon();
  updateDvPrompt();
}

function updateStats() {
  dom.stats.replaceChildren();
  if (selIsActiveOnly()) return;
  const rg = usedClip(sel);
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let fmtStyle = null;
  for (const [r, c] of cellsIn(rg)) {
    const v = valueAt(r, c);
    if (v === null || v === '') continue;
    count++;
    if (typeof v === 'number') {
      numCount++;
      sum += v;
      fmtStyle ??= styleAt(r, c);
    }
  }
  if (!count) return;
  const fmt = (n) => formatValue(n, fmtStyle?.numFmt && fmtStyle.numFmt !== 'general' ? fmtStyle : {}).text;
  const items = [];
  if (numCount) items.push(`평균: ${fmt(sum / numCount)}`);
  items.push(`개수: ${count}`);
  if (numCount) items.push(`합계: ${fmt(sum)}`);
  for (const t of items) dom.stats.append(el('span', {}, t));
}

// ───────────────────────── 이동 ─────────────────────────
function stepFrom(pos, dr, dc) {
  const m = wb.mergeAt(si, pos.r, pos.c);
  let { r, c } = pos;
  if (m) {
    if (dr > 0) r = m.r2;
    if (dc > 0) c = m.c2;
    if (dr < 0) r = m.r1;
    if (dc < 0) c = m.c1;
  }
  let nr = clamp(r + dr, 0, MAX_ROWS - 1);
  let nc = clamp(c + dc, 0, MAX_COLS - 1);
  // 숨겨진 행/열 건너뛰기
  if (dr) { nr = gv.rows.nextVisible(nr, Math.sign(dr)); if (gv.rows.isHidden(nr)) nr = pos.r; }
  if (dc) { nc = gv.cols.nextVisible(nc, Math.sign(dc)); if (gv.cols.isHidden(nc)) nc = pos.c; }
  return { r: nr, c: nc };
}

/** Ctrl+방향키: 데이터 영역 끝으로 이동 (빈 영역이면 시트 끝까지) */
function jump(pos, dr, dc) {
  const inside = (r, c) => r >= 0 && c >= 0 && r < MAX_ROWS && c < MAX_COLS;
  const hidden = (r, c) => (dr && gv.rows.isHidden(r)) || (dc && gv.cols.isHidden(c));
  const filled = (r, c) => !isEmptyAt(r, c);
  const u = wb.usedRange(si);
  let { r, c } = pos;
  const peek = () => { let rr = r; let cc = c; do { rr += dr; cc += dc; } while (inside(rr, cc) && hidden(rr, cc)); return [rr, cc]; };
  let [pr, pc] = peek();
  if (!inside(pr, pc)) return { r, c };
  if (filled(r, c) && filled(pr, pc)) {
    while (inside(pr, pc) && filled(pr, pc)) { r = pr; c = pc; [pr, pc] = peek(); }
    return { r, c };
  }
  for (;;) {
    [pr, pc] = peek();
    if (!inside(pr, pc)) return { r, c };
    r = pr;
    c = pc;
    if (filled(r, c)) return { r, c };
    if ((dr > 0 && r >= u.rows) || (dc > 0 && c >= u.cols)) {
      return { r: dr > 0 ? gv.rows.nextVisible(MAX_ROWS - 1, -1) : r, c: dc > 0 ? gv.cols.nextVisible(MAX_COLS - 1, -1) : c };
    }
  }
}

function move(dr, dc, { extend = false, ctrl = false } = {}) {
  if (extend) {
    const next = ctrl ? jump(focusCell, dr, dc) : stepFrom(focusCell, dr, dc);
    extendTo(next.r, next.c);
    return;
  }
  const next = ctrl ? jump(active, dr, dc) : stepFrom(active, dr, dc);
  selectCell(next.r, next.c);
}

function moveEnterTab(dir) {
  if (!selIsActiveOnly() && selKind === 'cells') {
    const rows = sel.r2 - sel.r1 + 1;
    const cols = sel.c2 - sel.c1 + 1;
    const total = rows * cols;
    let i = (active.r - sel.r1) * cols + (active.c - sel.c1);
    let j = (active.c - sel.c1) * rows + (active.r - sel.r1);
    let r = active.r;
    let c = active.c;
    for (let guard = 0; guard < Math.min(total, 100000); guard++) {
      if (dir === 'right' || dir === 'left') {
        i = (i + (dir === 'right' ? 1 : -1) + total) % total;
        r = sel.r1 + Math.floor(i / cols);
        c = sel.c1 + (i % cols);
      } else {
        j = (j + (dir === 'down' ? 1 : -1) + total) % total;
        c = sel.c1 + Math.floor(j / rows);
        r = sel.r1 + (j % rows);
      }
      if (!gv.rows.isHidden(r) && !gv.cols.isHidden(c)) break;
    }
    active = { r, c };
    anchor = { r, c };
    gv.ensureVisible(r, c);
    updateSelectionUI();
    return;
  }
  if (dir === 'right') {
    if (tabStartCol === null) tabStartCol = active.c;
    const n = stepFrom(active, 0, 1);
    selectCell(n.r, n.c, { keepTab: true });
  } else if (dir === 'left') {
    const n = stepFrom(active, 0, -1);
    selectCell(n.r, n.c, { keepTab: true });
  } else if (dir === 'down') {
    const n = stepFrom(active, 1, 0);
    selectCell(n.r, tabStartCol ?? n.c);
  } else {
    const n = stepFrom(active, -1, 0);
    selectCell(n.r, tabStartCol ?? n.c);
  }
}

// ───────────────────────── 편집 ─────────────────────────
const edInput = () => (document.activeElement === dom.formula ? dom.formula : dom.editor);

function focusGrid() {
  if (isDialogOpen() || document.querySelector('.backstage')) return;
  if (editing?.fromBar) { dom.formula.focus(); return; }
  if (document.activeElement !== dom.editor) dom.editor.focus({ preventScroll: true });
}

function deselectChart() {
  if (!chartSel) return;
  chartSel = null;
  gv.renderObjectsAll();
}

function startEdit(mode, text = null, { fromBar = false, caret = null } = {}) {
  if (editing) return;
  deselectChart();
  const { r, c } = active;
  const raw = wb.getRaw(si, r, c);
  editing = { r, c, mode, original: raw, point: null, fromBar };
  const value = text ?? raw;
  dom.editor.value = value;
  dom.formula.value = value;
  dom.editor.classList.remove('idle');
  gv.ensureVisible(r, c);
  updateSelectionUI();
  if (!fromBar) {
    dom.editor.focus({ preventScroll: true });
    const pos = caret ?? value.length;
    dom.editor.setSelectionRange(pos, pos);
  }
  afterEditInput();
}

function beginTyping() {
  if (editing) return;
  deselectChart();
  const { r, c } = active;
  editing = { r, c, mode: 'enter', original: wb.getRaw(si, r, c), point: null, fromBar: false };
  dom.editor.classList.remove('idle');
  dom.formula.value = dom.editor.value;
  gv.ensureVisible(r, c);
  updateSelectionUI();
  setMode();
}

function setMode() {
  if (!editing) {
    const f = allFilters().map(([, x]) => x).find((x) => Object.keys(x.hidden ?? {}).length) ?? null;
    const hidden = f ? Object.keys(f.hidden ?? {}).length : 0;
    dom.status.textContent = painter ? '서식 복사' : clip ? '대상을 선택한 후 Enter 키를 누르거나 붙여넣기를 선택하세요.'
      : hidden ? `필터 모드: ${f.r2 - f.r1 - hidden}/${f.r2 - f.r1}개 레코드 표시` : '준비';
  } else if (canPoint()) dom.status.textContent = '참조';
  else dom.status.textContent = editing.mode === 'enter' ? '입력' : '편집';
}

function setEditText(value, caret) {
  dom.editor.value = value;
  dom.formula.value = value;
  const inp = edInput();
  inp.setSelectionRange(caret, caret);
  afterEditInput();
}

function afterEditInput() {
  positionEditor();
  renderRefHighlights();
  updateAutocomplete();
  setMode();
}

function endEditUI() {
  editing = null;
  pendingKey = null;
  dom.editor.value = '';
  dom.editor.classList.add('idle');
  hideAutocomplete();
  editRefs = [];
  setMode();
}

function commitEdit(dir = null, { fillSel = false } = {}) {
  if (!editing) return true;
  let text = dom.editor.value;
  if (text.startsWith('=') && text.length > 1) {
    text = normalizeFormula(text);
    try {
      parse(text.slice(1));
    } catch {
      alertDialog('Tabula', '입력한 수식에 문제가 있습니다. 수식을 확인하세요.').then(() => { if (editing) focusGrid(); });
      return false;
    }
  }
  const { r, c, original } = editing;
  const rule = validationAt(sheet(), r, c);
  if (rule && rule.showError !== false && text !== original && !editing.dvOk && !checkValidation(wb, si, rule, r, c, text)) {
    validationError(rule, dir, fillSel);
    return false;
  }
  endEditUI();
  const multi = fillSel && !selIsActiveOnly();
  if (text !== original || multi) {
    wb.transact(() => {
      if (multi) {
        for (const [rr, cc] of cellsIn(sel)) wb.setInput(si, rr, cc, text.startsWith('=') ? shiftFormula(text, rr - r, cc - c) : text);
        autoFitRows(sel.r1, Math.min(sel.r2, sel.r1 + 500));
      } else {
        wb.setInput(si, r, c, text);
        maybeCalculatedColumn(r, c, text);
        afterDataEntry({ r1: r, c1: c, r2: r, c2: c });
        if (text.includes('\n')) wb.setStyle(si, r, c, { wrap: true });
        else if (!text.startsWith('=') || typeof valueAt(r, c) === 'number') autoWiden({ r1: r, c1: c, r2: r, c2: c });
        autoFitRows(r, r);
      }
    }, meta());
  }
  if (circles) circles = invalidCells(wb, si);
  if (dir && !multi) moveEnterTab(dir);
  else updateSelectionUI();
  focusGrid();
  return true;
}

function commitAndArrow(dr, dc) {
  if (commitEdit()) move(dr, dc);
}

function cancelEdit() {
  if (!editing) return;
  endEditUI();
  updateSelectionUI();
  focusGrid();
}

function positionEditor() {
  if (!gv?.cols) return;
  const ed = dom.editor;
  const target = editing ?? active;
  const m = wb.mergeAt(si, target.r, target.c);
  const rect = gv.screenRect(m ?? { r1: target.r, c1: target.c, r2: target.r, c2: target.c });
  ed.style.left = `${rect.x}px`;
  ed.style.top = `${rect.y}px`;
  if (!editing) return;
  const st = styleAt(target.r, target.c);
  Object.assign(ed.style, {
    fontFamily: fontStack(st.font || DEFAULT_FONT),
    fontSize: `${st.size || DEFAULT_SIZE}pt`,
    fontWeight: st.bold ? '700' : '400',
    fontStyle: st.italic ? 'italic' : 'normal',
    textDecoration: st.underline ? 'underline' : 'none',
    color: st.color || '#000',
    background: st.fill || '#fff',
    textAlign: st.align && st.align !== 'left' && !ed.value.startsWith('=') ? st.align : 'left',
    whiteSpace: st.wrap ? 'pre-wrap' : 'pre',
  });
  const lines = ed.value.split('\n');
  const textW = Math.max(...lines.map((l) => measureText(l, st))) + 12;
  const maxW = Math.max(rect.w, gv.viewW - rect.x - 4);
  const w = st.wrap ? rect.w : clamp(textW, rect.w, maxW);
  ed.style.width = `${w}px`;
  ed.style.height = '0px';
  ed.style.height = `${Math.max(rect.h, ed.scrollHeight)}px`;
  if (ac) placeAutocomplete();
}

// ── 참조 모드 ──
const POINT_CHARS = new Set(['=', '(', ',', '+', '-', '*', '/', '^', '&', '<', '>', ':']);
function canPoint() {
  if (!editing) return false;
  const inp = edInput();
  const text = inp.value;
  if (!text.startsWith('=')) return false;
  if (editing.point) return true;
  const before = text.slice(0, inp.selectionStart).trimEnd();
  return POINT_CHARS.has(before.at(-1));
}

function insertPointRef(rg) {
  const inp = edInput();
  const text = inp.value;
  let start;
  let end;
  if (editing.point) ({ start, end } = editing.point);
  else { start = inp.selectionStart; end = inp.selectionEnd; }
  let ref = cellName(rg.r1, rg.c1);
  if (!isSingle(rg)) ref += `:${cellName(rg.r2, rg.c2)}`;
  const value = text.slice(0, start) + ref + text.slice(end);
  const anchorPt = editing.point?.anchor ?? { r: rg.r1, c: rg.c1 };
  setEditText(value, start + ref.length);
  editing.point = { start, end: start + ref.length, anchor: anchorPt, cur: { r: rg.r2, c: rg.c2 }, rg };
  setMode();
}

function pointMove(dr, dc, extend) {
  const p = editing.point;
  const cur = p?.cur ?? { r: editing.r, c: editing.c };
  const next = { r: clamp(cur.r + dr, 0, MAX_ROWS - 1), c: clamp(cur.c + dc, 0, MAX_COLS - 1) };
  if (extend && p) {
    insertPointRef(norm(p.anchor, next));
    editing.point.cur = next;
  } else {
    insertPointRef({ r1: next.r, c1: next.c, r2: next.r, c2: next.c });
    editing.point.anchor = next;
  }
  growTo(next.r, next.c);
  gv.ensureVisible(next.r, next.c);
}

function toggleAbsolute() {
  const inp = edInput();
  const text = inp.value;
  if (!text.startsWith('=')) return;
  const pos = inp.selectionStart;
  const re = /(\$?)([A-Za-z]{1,3})(\$?)(\d+)/g;
  let m;
  while ((m = re.exec(text))) {
    const s = m.index;
    const e = s + m[0].length;
    if (pos < s || pos > e) continue;
    const state = (m[1] ? 2 : 0) + (m[3] ? 1 : 0);
    const nextState = { 0: 3, 3: 1, 1: 2, 2: 0 }[state];
    const rep = `${nextState & 2 ? '$' : ''}${m[2].toUpperCase()}${nextState & 1 ? '$' : ''}${m[4]}`;
    setEditText(text.slice(0, s) + rep + text.slice(e), s + rep.length);
    editing.point = null;
    return;
  }
}

function renderRefHighlights() {
  editRefs = [];
  if (editing) {
    const colorFor = new Map();
    for (const ref of listRefs(dom.editor.value)) {
      if (ref.sheet && ref.sheet.toLowerCase() !== sheet().name.toLowerCase()) continue;
      const k = `${ref.r1},${ref.c1},${ref.r2},${ref.c2}`;
      if (!colorFor.has(k)) colorFor.set(k, REF_COLORS[colorFor.size % REF_COLORS.length]);
      editRefs.push({ rg: { r1: ref.r1, c1: ref.c1, r2: ref.r2, c2: ref.c2 }, color: colorFor.get(k) });
    }
  }
  gv.renderOverlays();
}

// ── 함수 자동 완성 ──
function updateAutocomplete() {
  if (!editing) return hideAutocomplete();
  const inp = edInput();
  const text = inp.value;
  const pos = inp.selectionStart;
  const before = text.slice(0, pos);
  if (!text.startsWith('=') || (before.split('"').length - 1) % 2 === 1) return hideAutocomplete();
  const m = /(^|[^A-Za-z0-9_."$!:])([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(1));
  if (!m || /^[A-Za-z]{1,3}\d+$/.test(m[2])) return hideAutocomplete();
  const prefix = m[2].toUpperCase();
  const items = FUNCTION_NAMES.filter((n) => n.startsWith(prefix)).slice(0, 12);
  if (!items.length || text[pos] === '(') return hideAutocomplete();
  ac = { items, index: 0, start: pos - m[2].length, end: pos };
  dom.ac.replaceChildren(...items.map((name, i) => el('li', {
    class: i === 0 ? 'active' : '',
    onmousedown: (e) => { e.preventDefault(); e.stopPropagation(); acceptAutocomplete(name); },
    title: FUNC_INFO[name]?.desc ?? '',
  }, el('span', {}, name), el('small', {}, FUNC_INFO[name]?.cat ?? ''))));
  dom.ac.style.display = 'block';
  placeAutocomplete();
  return undefined;
}

function placeAutocomplete() {
  const ed = dom.editor;
  dom.ac.style.left = ed.style.left;
  dom.ac.style.top = `${ed.offsetTop + ed.offsetHeight + 2}px`;
}

function hideAutocomplete() {
  ac = null;
  dom.ac.style.display = 'none';
}

function moveAutocomplete(d) {
  ac.index = (ac.index + d + ac.items.length) % ac.items.length;
  [...dom.ac.children].forEach((li, i) => li.classList.toggle('active', i === ac.index));
  dom.ac.children[ac.index].scrollIntoView({ block: 'nearest' });
}

function acceptAutocomplete(name = ac.items[ac.index]) {
  const inp = edInput();
  const text = inp.value;
  const { start, end } = ac;
  const value = `${text.slice(0, start)}${name}(${text.slice(end)}`;
  hideAutocomplete();
  setEditText(value, start + name.length + 1);
  toast(FUNC_INFO[name]?.sig ?? name);
}

// ───────────────────────── 행 높이 자동 맞춤 · 열 자동 넓힘 ─────────────────────────
function wrappedLines(text, width, st) {
  let lines = 0;
  for (const para of String(text).split('\n')) {
    if (!para) { lines++; continue; }
    let cur = '';
    let count = 1;
    for (const ch of para) {
      if (measureText(cur + ch, st) > width && cur) { count++; cur = ch; } else cur += ch;
    }
    lines += count;
  }
  return Math.max(1, lines);
}

function autoFitRows(r1, r2) {
  const s = sheet();
  const cols = Math.min(wb.usedRange(si).cols, 500);
  for (let r = r1; r <= r2; r++) {
    if (s.rowManual[r]) continue;
    let need = DEFAULT_ROW_HEIGHT;
    for (let c = 0; c < cols; c++) {
      const cell = wb.getCell(si, r, c);
      if (!cell?.raw || wb.mergeAt(si, r, c)) continue;
      const st = styleAt(r, c);
      const text = formatValue(valueAt(r, c), st).text;
      const lineH = ((st.size || DEFAULT_SIZE) * 4 / 3) * 1.2;
      const lines = st.wrap ? wrappedLines(text, wb.colWidth(si, c) - 7, st) : text.split('\n').length;
      need = Math.max(need, Math.ceil(lines * lineH + 3));
    }
    need = Math.min(409, need);
    if ((s.rowHeights[r] ?? DEFAULT_ROW_HEIGHT) !== need) wb.setRowHeight(si, r, need, false);
  }
}

/** 엑셀처럼: 너비를 바꾼 적 없는 열에 숫자가 들어가지 않으면 열을 넓힘 */
function autoWiden(rg, { grow = false } = {}) {
  const s = sheet();
  const u = usedClip(rg);
  for (let c = u.c1; c <= Math.min(u.c2, u.c1 + 200); c++) {
    if (s.colWidths[c] !== undefined && !grow) continue;
    let need = 0;
    for (let r = u.r1; r <= Math.min(u.r2, u.r1 + 2000); r++) {
      const v = valueAt(r, c);
      const st = styleAt(r, c);
      if (typeof v !== 'number' || st.wrap || wb.mergeAt(si, r, c)) continue;
      need = Math.max(need, measureText(formatValue(v, st).text, st) + 10);
    }
    if (need > (s.colWidths[c] ?? DEFAULT_COL_WIDTH) + 1 && need < 400) wb.setColWidth(si, c, Math.ceil(need));
  }
}

// ───────────────────────── 키보드 ─────────────────────────
const NAV_KEYS = new Set(['Enter', 'Tab', 'ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight']);

function onEditorKeyDown(e) {
  if (isDialogOpen()) return;
  if (isMenuOpen()) {
    if (e.key === 'Escape') { closeMenus(); e.preventDefault(); }
    return;
  }
  if (editing) onEditingKey(e);
  else onGridKey(e);
}

function onEditingKey(e) {
  const ctrl = e.ctrlKey || e.metaKey;
  if (e.isComposing || e.keyCode === 229) {
    if (NAV_KEYS.has(e.key)) pendingKey = { key: e.key, shiftKey: e.shiftKey, ctrlKey: e.ctrlKey, metaKey: e.metaKey, altKey: e.altKey };
    return;
  }
  pendingKey = null;
  const fromBar = document.activeElement === dom.formula;
  if (ac) {
    if (e.key === 'ArrowDown') { e.preventDefault(); moveAutocomplete(1); return; }
    if (e.key === 'ArrowUp') { e.preventDefault(); moveAutocomplete(-1); return; }
    if (e.key === 'Tab') { e.preventDefault(); acceptAutocomplete(); return; }
    if (e.key === 'Escape') { e.preventDefault(); hideAutocomplete(); return; }
  }
  switch (e.key) {
    case 'Enter':
      e.preventDefault();
      if (e.altKey) {
        const inp = edInput();
        const pos = inp.selectionStart;
        setEditText(`${inp.value.slice(0, pos)}\n${inp.value.slice(inp.selectionEnd)}`, pos + 1);
        return;
      }
      commitEdit(ctrl ? null : e.shiftKey ? 'up' : 'down', { fillSel: ctrl });
      return;
    case 'Tab':
      e.preventDefault();
      commitEdit(e.shiftKey ? 'left' : 'right');
      return;
    case 'Escape':
      e.preventDefault();
      cancelEdit();
      return;
    case 'F2':
      e.preventDefault();
      editing.mode = editing.mode === 'enter' ? 'edit' : 'enter';
      editing.point = null;
      setMode();
      return;
    case 'F4':
      e.preventDefault();
      toggleAbsolute();
      return;
    case 'ArrowUp': case 'ArrowDown': case 'ArrowLeft': case 'ArrowRight': {
      if (fromBar || editing.mode === 'edit') { editing.point = null; setTimeout(afterCaretMove); return; }
      const [dr, dc] = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      e.preventDefault();
      if (canPoint()) pointMove(dr, dc, e.shiftKey);
      else commitAndArrow(dr, dc);
      return;
    }
    default:
      if (e.key === 'Home' || e.key === 'End') setTimeout(afterCaretMove);
  }
}

function afterCaretMove() {
  if (!editing) return;
  updateAutocomplete();
  setMode();
}

function onGridKey(e) {
  if (handleKeytipKey(e)) return;
  if (e.isComposing || e.keyCode === 229) return;
  const ctrl = e.ctrlKey || e.metaKey;
  const k = e.key;
  const handled = () => e.preventDefault();

  if (drawKind && k === 'Escape') { handled(); endDraw(); return; }
  if (chartSel) {
    if (k === 'Delete' || k === 'Backspace') { handled(); deleteObject(chartSel); return; }
    if (k === 'Escape') { handled(); deselectChart(); updateSelectionUI(); return; }
    const step = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[k];
    if (step) { handled(); nudgeObject(chartSel, step[0] * (ctrl ? 10 : 1), step[1] * (ctrl ? 10 : 1)); return; }
    if (ctrl && (k === 'c' || k === 'C' || k === 'x' || k === 'X')) { handled(); copyObject(chartSel, k.toLowerCase() === 'x'); return; }
    if (ctrl && (k === 'd' || k === 'D')) { handled(); copyObject(chartSel, false); pasteObject(); return; }
    if (k === 'Enter' || k === 'F2') { handled(); editObject(chartSel); return; }
    if (e.altKey && (e.code === 'KeyC' || e.code === 'KeyS') && sheet().slicers?.some((x) => x.id === chartSel)) {
      handled();
      if (e.code === 'KeyC') slicerClear(chartSel); else run('slicerMulti');
      gv.renderObjectsAll();
      return;
    }
    if (!ctrl && !e.altKey && k.length === 1 && sheet().shapes?.some((x) => x.id === chartSel)) { handled(); shapeDialog(chartSel, k); return; }
  }
  if (e.altKey && k === 'ArrowDown' && !ctrl) {
    const tt = tableHere();
    if (tt?.totals && active.r === tt.r2) { handled(); openTotalsMenu(); return; }
    const hdr = tt?.filter && tt.header && active.r === tt.r1;
    if (hdr || (sheet().filter && active.r === sheet().filter.r1)) {
      const btn = document.querySelector(`.fbtn[data-c="${active.c}"][data-t="${hdr ? tt.id : ''}"]`);
      if (btn) { handled(); openFilterMenu(active.c, btn, hdr ? tt.id : ''); return; }
    }
    const rule = validationAt(sheet(), active.r, active.c);
    if (rule?.type === 'list') { handled(); openDvList(); return; }
  }
  if (e.altKey && !ctrl && (k === '=' || e.code === 'Equal')) { handled(); run('autosum'); return; }
  if (e.altKey && k === 'F1') { handled(); run('chartColumn'); return; }
  if (ctrl) {
    if (e.shiftKey) {
      const byCode = {
        Digit1: 'fmtNumber', Digit3: 'fmtDate', Digit4: 'fmtCurrency', Digit5: 'fmtPercent', Digit7: 'borderOutside',
        Backquote: 'fmtGeneral', Semicolon: 'insertTime', Equal: 'insertMenuKey', KeyL: 'toggleFilter',
      };
      if (byCode[e.code]) { handled(); run(byCode[e.code]); return; }
    }
    const lower = k.length === 1 ? k.toLowerCase() : k;
    const map = {
      z: 'undo', y: 'redo', b: 'bold', i: 'italic', u: 'underline', 5: 'strike', d: 'fillDown', r: 'fillRight',
      s: 'save', f: 'find', h: 'replace', g: 'goto', p: 'print', o: 'open', 2: 'bold', 3: 'italic', 4: 'underline',
      ';': 'insertDate', '`': 'toggleFormulas', 1: 'formatCells', '-': 'deleteMenuKey', F1: 'toggleRibbon',
      9: 'hideRows', 0: 'hideCols', t: 'createTable', l: 'createTable',
    };
    if (lower === 'a') { handled(); selectAll(); return; }
    if (lower === ' ' || e.code === 'Space') { handled(); selectCols(sel.c1, sel.c2, active); return; }
    if (k === 'Home') {
      handled();
      const t = { r: sheet().freeze?.rows || 0, c: sheet().freeze?.cols || 0 };
      if (e.shiftKey) extendTo(t.r, t.c); else selectCell(t.r, t.c);
      return;
    }
    if (k === 'End') {
      handled();
      const u = wb.usedRange(si);
      if (e.shiftKey) extendTo(Math.max(0, u.rows - 1), Math.max(0, u.cols - 1));
      else selectCell(Math.max(0, u.rows - 1), Math.max(0, u.cols - 1));
      return;
    }
    if (k === 'PageDown') { handled(); run('nextSheet'); return; }
    if (k === 'PageUp') { handled(); run('prevSheet'); return; }
    const arrows = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] };
    if (arrows[k]) { handled(); move(...arrows[k], { extend: e.shiftKey, ctrl: true }); return; }
    if (map[lower]) { handled(); run(map[lower]); return; }
    return;
  }

  switch (k) {
    case 'ArrowUp': handled(); move(-1, 0, { extend: e.shiftKey }); return;
    case 'ArrowDown': handled(); move(1, 0, { extend: e.shiftKey }); return;
    case 'ArrowLeft': handled(); move(0, -1, { extend: e.shiftKey }); return;
    case 'ArrowRight': handled(); move(0, 1, { extend: e.shiftKey }); return;
    case 'Enter':
      handled();
      if (clip && !e.shiftKey) { pasteInternal('all'); clip = null; updateSelectionUI(); setMode(); return; }
      moveEnterTab(e.shiftKey ? 'up' : 'down');
      return;
    case 'Tab': handled(); moveEnterTab(e.shiftKey ? 'left' : 'right'); return;
    case 'Home': handled(); if (e.shiftKey) extendTo(focusCell.r, 0); else selectCell(active.r, 0); return;
    case 'PageDown': handled(); move(gv.pageRows(), 0, { extend: e.shiftKey }); return;
    case 'PageUp': handled(); move(-gv.pageRows(), 0, { extend: e.shiftKey }); return;
    case 'Delete': handled(); run('clearContents'); return;
    case 'Backspace': handled(); startEdit('enter', ''); return;
    case 'F2':
      handled();
      if (e.shiftKey) run('editComment');
      else startEdit('edit');
      return;
    case 'F9': handled(); run('recalc'); return;
    case 'F11': if (e.shiftKey) { handled(); run('addSheet'); } return;
    case 'F12': handled(); run('saveAs'); return;
    case 'Escape':
      handled();
      if (clip || painter) { clip = null; painter = null; dom.view.classList.remove('painting'); updateSelectionUI(); setMode(); }
      return;
    case 'ContextMenu': {
      handled();
      const rect = gv.clientRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c });
      showContextMenu({ x: rect.left + 10, y: rect.bottom }, 'cell');
      return;
    }
    case ' ':
      if (e.shiftKey) { handled(); selectRows(sel.r1, sel.r2, active); }
      return;
    default:
  }
}

// ───────────────────────── 마우스 ─────────────────────────
let autoScrollTimer = null;
let lastMouse = { x: 0, y: 0 };
let lastShift = false;
function startAutoScroll() {
  stopAutoScroll();
  autoScrollTimer = setInterval(() => {
    if (!drag || drag.type === 'colResize' || drag.type === 'rowResize' || drag.type === 'obj' || drag.type === 'draw') return stopAutoScroll();
    const hit = gv.hitTest(lastMouse.x, lastMouse.y, true);
    let { dx, dy } = hit;
    if (drag.type === 'colSel') dy = 0;
    if (drag.type === 'rowSel') dx = 0;
    if (!dx && !dy) return undefined;
    gv.scrollBy(dx * 40, dy * DEFAULT_ROW_HEIGHT);
    onDragMove(lastMouse.x, lastMouse.y);
    return undefined;
  }, 50);
}
function stopAutoScroll() { clearInterval(autoScrollTimer); autoScrollTimer = null; }

function onViewMouseDown(e) {
  if (e.target === dom.editor || dom.ac.contains(e.target)) return;
  closeMenus();
  const t = e.target;
  if (t.classList.contains('fbtn')) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    openFilterMenu(Number(t.dataset.c), t, t.dataset.t || '');
    return;
  }
  if (t.classList.contains('dv-btn')) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    if (t.dataset.tt) openTotalsMenu(); else openDvList();
    return;
  }
  if (drawKind) {
    e.preventDefault();
    const hit = gv.hitTest(e.clientX, e.clientY);
    const kind = drawKind;
    endDraw();
    if (hit.zone !== 'cell' || e.button !== 0) return;
    if (editing && !commitEdit()) return;
    focusGrid();
    const temp = newShape(kind, { x: Math.round(hit.sheetX), y: Math.round(hit.sheetY), w: 0, h: 0 });
    sheet().shapes.push(temp);
    chartSel = temp.id;
    drag = { type: 'draw', id: temp.id, x0: hit.sheetX, y0: hit.sheetY, start: { x: e.clientX, y: e.clientY } };
    gv.renderObjectsAll();
    return;
  }
  const slBtn = t.closest('.sl-item, .sl-clear, .sl-multi');
  if (slBtn && e.button === 0) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    focusGrid();
    const id = t.closest('.obj').dataset.id;
    chartSel = id;
    if (slBtn.classList.contains('sl-item')) slicerPick(id, slBtn.dataset.k, e.ctrlKey || e.metaKey);
    else if (slBtn.classList.contains('sl-clear')) slicerClear(id);
    else { const sl = sheet().slicers.find((x) => x.id === id); updateObject(id, { multi: !sl.multi }); }
    gv.renderObjectsAll();
    updateSelectionUI();
    return;
  }
  const objEl = t.closest('.obj');
  if (objEl) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    focusGrid();
    const id = objEl.dataset.id;
    if (chartSel !== id) { chartSel = id; gv.renderObjectsAll(); updateSelectionUI(); }
    if (e.button !== 0) return;
    const found = findObject(sheet(), id);
    if (!found) return;
    const o = found.obj;
    const corner = t.classList.contains('ch-h') ? [...t.classList].find((c) => ['nw', 'ne', 'sw', 'se'].includes(c)) : null;
    drag = { type: 'obj', prop: found.prop, id, corner, shift: e.shiftKey, start: { x: e.clientX, y: e.clientY }, orig: { x: o.x, y: o.y, w: o.w, h: o.h } };
    return;
  }
  if (t.classList.contains('fill-handle')) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    drag = { type: 'fill', src: { ...sel }, target: null };
    startAutoScroll();
    return;
  }
  const hit = gv.hitTest(e.clientX, e.clientY);
  e.preventDefault();
  if (hit.zone === 'colHeader' && hit.edgeCol !== null && e.button === 0) {
    drag = { type: 'colResize', c: hit.edgeCol, x: e.clientX, w: wb.colWidth(si, hit.edgeCol), orig: sheet().colWidths[hit.edgeCol] };
    return;
  }
  if (hit.zone === 'rowHeader' && hit.edgeRow !== null && e.button === 0) {
    drag = { type: 'rowResize', r: hit.edgeRow, y: e.clientY, h: gv.rows.size(hit.edgeRow), orig: sheet().rowHeights[hit.edgeRow] };
    return;
  }
  if (chartSel) deselectChart();

  if (editing && hit.zone === 'cell' && canPoint() && e.button === 0) {
    const { r, c } = hit;
    const m = wb.mergeAt(si, r, c) ?? { r1: r, c1: c, r2: r, c2: c };
    if (e.shiftKey && editing.point) insertPointRef(norm(editing.point.anchor, { r, c }));
    else { editing.point = null; insertPointRef(m); }
    editing.point.anchor = { r, c };
    drag = { type: 'point' };
    startAutoScroll();
    return;
  }
  if (editing && !commitEdit()) return;
  focusGrid();

  if (hit.zone === 'corner') { selectAll(); return; }
  if (hit.zone === 'colHeader') {
    const c = hit.c;
    if (e.button === 2 && inSel(0, c) && (selKind === 'cols' || selKind === 'all')) return;
    if (e.shiftKey && selKind === 'cols') selectCols(anchor.c, c, active);
    else selectCols(c, c);
    drag = { type: 'colSel', start: e.shiftKey && selKind === 'cols' ? anchor.c : c };
    startAutoScroll();
    return;
  }
  if (hit.zone === 'rowHeader') {
    const r = hit.r;
    if (e.button === 2 && inSel(r, 0) && (selKind === 'rows' || selKind === 'all')) return;
    if (e.shiftKey && selKind === 'rows') selectRows(anchor.r, r, active);
    else selectRows(r, r);
    drag = { type: 'rowSel', start: e.shiftKey && selKind === 'rows' ? anchor.r : r };
    startAutoScroll();
    return;
  }
  const { r, c } = hit;
  if (e.button === 2) {
    if (!inSel(r, c)) selectCell(r, c, { scroll: false });
    return;
  }
  if (e.shiftKey) extendTo(r, c, { scroll: false });
  else selectCell(r, c, { scroll: false });
  drag = { type: 'select' };
  startAutoScroll();
}

function onViewMouseMove(e) {
  lastMouse = { x: e.clientX, y: e.clientY };
  if (drag) return;
  const t = e.target;
  const cls = dom.view.classList;
  cls.remove('col-resize', 'row-resize', 'col-select', 'row-select', 'default-cursor');
  if (t.closest?.('.obj') || t.classList?.contains('fbtn') || t.classList?.contains('dv-btn')) return;
  const hit = gv.hitTest(e.clientX, e.clientY);
  if (hit.zone === 'colHeader') cls.add(hit.edgeCol !== null ? 'col-resize' : 'col-select');
  else if (hit.zone === 'rowHeader') cls.add(hit.edgeRow !== null ? 'row-resize' : 'row-select');
  else if (hit.zone === 'corner') cls.add('default-cursor');
  const cellEl = t.closest?.('.c');
  if (cellEl?.dataset.cm) {
    const b = cellEl.getBoundingClientRect();
    dom.tip.textContent = cellEl.dataset.cm;
    dom.tip.style.display = 'block';
    dom.tip.style.left = `${Math.min(b.right + 6, innerWidth - 250)}px`;
    dom.tip.style.top = `${b.top}px`;
  } else {
    dom.tip.style.display = 'none';
  }
}

function onDragMove(x, y) {
  if (!drag) return;
  switch (drag.type) {
    case 'select': {
      const { r, c } = gv.hitTest(x, y, true);
      if (r !== focusCell.r || c !== focusCell.c) {
        extendTo(r, c, { scroll: false });
        if (!selIsActiveOnly()) dom.nameBox.value = `${sel.r2 - sel.r1 + 1}R x ${sel.c2 - sel.c1 + 1}C`;
      }
      break;
    }
    case 'point': {
      const { r, c } = gv.hitTest(x, y, true);
      insertPointRef(expandMerges(norm(editing.point.anchor, { r, c })));
      break;
    }
    case 'colSel': {
      const { c } = gv.hitTest(x, y, true);
      selectCols(drag.start, c, { r: gv.firstVisibleRow(), c: drag.start });
      break;
    }
    case 'rowSel': {
      const { r } = gv.hitTest(x, y, true);
      selectRows(drag.start, r, { r: drag.start, c: 0 });
      break;
    }
    case 'fill': {
      const { r, c } = gv.hitTest(x, y, true);
      const s = drag.src;
      const dDown = r - s.r2;
      const dUp = s.r1 - r;
      const dRight = c - s.c2;
      const dLeft = s.c1 - c;
      const vert = Math.max(dDown, dUp);
      const horiz = Math.max(dRight, dLeft);
      let t = null;
      if (vert > 0 || horiz > 0) {
        if (vert >= horiz) t = dDown > 0 ? { ...s, r2: r, dir: 'down' } : { ...s, r1: r, dir: 'up' };
        else t = dRight > 0 ? { ...s, c2: c, dir: 'right' } : { ...s, c1: c, dir: 'left' };
      }
      drag.target = t;
      fillPreview = t;
      gv.renderOverlays();
      break;
    }
    case 'colResize': {
      const w = Math.max(0, drag.w + (x - drag.x) / gv.z);
      drag.newW = w;
      sheet().colWidths[drag.c] = Math.round(w);
      gv.layout();
      positionEditor();
      break;
    }
    case 'rowResize': {
      const h = Math.max(0, drag.h + (y - drag.y) / gv.z);
      drag.newH = h;
      sheet().rowHeights[drag.r] = Math.round(h);
      gv.layout();
      break;
    }
    case 'obj': {
      const ch = (sheet()[drag.prop] ?? []).find((xx) => xx.id === drag.id);
      if (!ch) break;
      const dx = (x - drag.start.x) / gv.z;
      const dy = (y - drag.start.y) / gv.z;
      const o = drag.orig;
      if (!drag.corner) {
        ch.x = Math.max(0, Math.round(o.x + dx));
        ch.y = Math.max(0, Math.round(o.y + dy));
      } else {
        const k = drag.corner;
        const [minW, minH] = drag.prop === 'charts' ? [120, 90] : drag.prop === 'slicers' ? [80, 56] : ch.kind === 'line' ? [0, 0] : [8, 8];
        let { x: nx, y: ny, w: nw, h: nh } = o;
        if (k.includes('e')) nw = o.w + dx;
        if (k.includes('s')) nh = o.h + dy;
        if (k.includes('w')) { nw = o.w - dx; nx = o.x + dx; }
        if (k.includes('n')) { nh = o.h - dy; ny = o.y + dy; }
        // 그림은 모서리를 끌면 가로세로 비율 유지 (Shift 누르면 자유롭게)
        if (drag.prop === 'images' ? !drag.shift : drag.shift && o.w && o.h) {
          const ratio = o.h / o.w;
          nw = Math.max(nw, minW);
          nh = nw * ratio;
          if (k.includes('w')) nx = o.x + o.w - nw;
          if (k.includes('n')) ny = o.y + o.h - nh;
        }
        if (nw >= minW) { ch.w = Math.round(nw); ch.x = Math.max(0, Math.round(nx)); }
        if (nh >= minH) { ch.h = Math.round(nh); ch.y = Math.max(0, Math.round(ny)); }
      }
      drag.moved = true;
      gv.renderObjectsAll();
      break;
    }
    case 'draw': {
      const sh = sheet().shapes.find((xx) => xx.id === drag.id);
      if (!sh) break;
      let dx = (x - drag.start.x) / gv.z;
      let dy = (y - drag.start.y) / gv.z;
      if (lastShift && sh.kind !== 'line') { const m = Math.max(Math.abs(dx), Math.abs(dy)); dx = Math.sign(dx || 1) * m; dy = Math.sign(dy || 1) * m; }
      sh.x = Math.max(0, Math.round(Math.min(drag.x0, drag.x0 + dx)));
      sh.y = Math.max(0, Math.round(Math.min(drag.y0, drag.y0 + dy)));
      sh.w = Math.round(Math.abs(dx));
      sh.h = Math.round(Math.abs(dy));
      if (sh.kind === 'line') sh.flip = (dx < 0) !== (dy < 0);
      drag.moved = true;
      gv.renderObjectsAll();
      break;
    }
    default:
  }
}

function onDragEnd() {
  if (!drag) return;
  const d = drag;
  drag = null;
  stopAutoScroll();
  switch (d.type) {
    case 'fill':
      fillPreview = null;
      gv.renderOverlays();
      if (d.target) doFill(d.src, d.target);
      break;
    case 'colResize': {
      if (d.newW === undefined) break;
      if (d.orig === undefined) delete sheet().colWidths[d.c]; else sheet().colWidths[d.c] = d.orig;
      const cols = selKind === 'cols' && d.c >= sel.c1 && d.c <= sel.c2 ? range(sel.c1, Math.min(sel.c2, sel.c1 + 500)) : [d.c];
      wb.transact(() => cols.forEach((c) => wb.setColWidth(si, c, d.newW)), meta());
      break;
    }
    case 'rowResize': {
      if (d.newH === undefined) break;
      if (d.orig === undefined) delete sheet().rowHeights[d.r]; else sheet().rowHeights[d.r] = d.orig;
      const rows = selKind === 'rows' && d.r >= sel.r1 && d.r <= sel.r2 ? range(sel.r1, Math.min(sel.r2, sel.r1 + 2000)) : [d.r];
      wb.transact(() => rows.forEach((r) => wb.setRowHeight(si, r, d.newH)), meta());
      break;
    }
    case 'obj': {
      if (!d.moved) break;
      const list = sheet()[d.prop];
      const ch = list.find((x) => x.id === d.id);
      if (!ch) break;
      const final = { ...ch };
      Object.assign(ch, d.orig);
      wb.transact(() => wb.setSheetProp(si, d.prop, list.map((x) => (x.id === d.id ? final : { ...x }))), meta());
      break;
    }
    case 'draw': {
      const list = sheet().shapes;
      const i = list.findIndex((x) => x.id === d.id);
      if (i < 0) break;
      const [temp] = list.splice(i, 1);
      if (temp.kind === 'line' ? temp.w + temp.h < 4 : temp.w < 4 || temp.h < 4) {
        Object.assign(temp, temp.kind === 'line' ? { w: 150, h: 0, flip: false } : temp.kind === 'textbox' ? { w: 160, h: 48 } : { w: 150, h: 90 });
      }
      addObject('shapes', temp);
      if (temp.kind === 'textbox') shapeDialog(temp.id);
      break;
    }
    case 'select':
      if (painter) applyPainter();
      updateSelectionUI();
      break;
    case 'point':
      focusGrid();
      break;
    default:
  }
}

function onViewDblClick(e) {
  const t = e.target;
  const objEl = t.closest('.obj');
  if (objEl) { editObject(objEl.dataset.id); return; }
  if (t.classList.contains('fbtn') || t.classList.contains('dv-btn') || t === dom.editor) return;
  const hit = gv.hitTest(e.clientX, e.clientY);
  if (hit.zone === 'colHeader' && hit.edgeCol !== null) {
    const cols = selKind === 'cols' && inSel(0, hit.edgeCol) ? range(sel.c1, Math.min(sel.c2, sel.c1 + 200)) : [hit.edgeCol];
    autofitCols(cols);
    return;
  }
  if (hit.zone === 'rowHeader' && hit.edgeRow !== null) {
    wb.transact(() => { wb.setRowHeight(si, hit.edgeRow, DEFAULT_ROW_HEIGHT, false); autoFitRows(hit.edgeRow, hit.edgeRow); }, meta());
    return;
  }
  if (hit.zone !== 'cell' || editing) return;
  startEdit('edit');
}

function autofitCols(cols) {
  wb.transact(() => {
    const u = wb.usedRange(si);
    for (const c of cols) {
      let w = 0;
      for (let r = 0; r < Math.min(u.rows, 5000); r++) {
        const v = valueAt(r, c);
        if (v === null || wb.mergeAt(si, r, c)) continue;
        const cell = wb.getCell(si, r, c);
        const st = styleAt(r, c);
        const text = view.showFormulas && cell?.formula ? cell.raw : formatValue(v, st).text;
        for (const line of text.split('\n')) w = Math.max(w, measureText(line, st) + (allFilters().some(([, f]) => f.r1 === r && c >= f.c1 && c <= f.c2) ? 28 : 10));
      }
      wb.setColWidth(si, c, w ? Math.min(600, Math.ceil(w)) : DEFAULT_COL_WIDTH);
    }
  }, meta());
}

// ───────────────────────── 채우기 ─────────────────────────
function doFill(src, t) {
  const vertical = t.dir === 'down' || t.dir === 'up';
  const forward = t.dir === 'down' || t.dir === 'right';
  wb.transact(() => {
    if (vertical) {
      const count = forward ? t.r2 - src.r2 : src.r1 - t.r1;
      for (let c = src.c1; c <= src.c2; c++) {
        const seq = [];
        for (let r = src.r1; r <= src.r2; r++) seq.push({ data: cellData(wb.getCell(si, r, c)), value: valueAt(r, c), pos: r });
        const gen = makeSeries(seq, forward, 'row');
        for (let k = 0; k < count; k++) {
          const r = forward ? src.r2 + 1 + k : src.r1 - 1 - k;
          if (!filterHidden(r)) wb.setCellData(si, r, c, gen(k, r));
        }
      }
    } else {
      const count = forward ? t.c2 - src.c2 : src.c1 - t.c1;
      for (let r = src.r1; r <= src.r2; r++) {
        const seq = [];
        for (let c = src.c1; c <= src.c2; c++) seq.push({ data: cellData(wb.getCell(si, r, c)), value: valueAt(r, c), pos: c });
        const gen = makeSeries(seq, forward, 'col');
        for (let k = 0; k < count; k++) {
          const c = forward ? src.c2 + 1 + k : src.c1 - 1 - k;
          wb.setCellData(si, r, c, gen(k, c));
        }
      }
    }
    if (forward) afterDataEntry(vertical ? { r1: src.r2 + 1, c1: src.c1, r2: t.r2, c2: src.c2 } : { r1: src.r1, c1: src.c2 + 1, r2: src.r2, c2: t.c2 });
  }, meta());
  selectRange({ r1: Math.min(src.r1, t.r1), c1: Math.min(src.c1, t.c1), r2: Math.max(src.r2, t.r2), c2: Math.max(src.c2, t.c2) }, 'cells', active);
}

function fillCopy(dir) {
  const rg = usedClip(sel);
  const shift = (d, dr, dc) => d && { ...d, raw: d.raw.startsWith('=') ? shiftFormula(d.raw, dr, dc) : d.raw };
  wb.transact(() => {
    if (dir === 'down') {
      const srcR = rg.r1 === rg.r2 ? rg.r1 - 1 : rg.r1;
      if (srcR < 0) return;
      for (let c = rg.c1; c <= rg.c2; c++) {
        const d = cellData(wb.getCell(si, srcR, c));
        for (let r = srcR + 1; r <= rg.r2; r++) if (!filterHidden(r)) wb.setCellData(si, r, c, shift(d, r - srcR, 0));
      }
    } else if (dir === 'right') {
      const srcC = rg.c1 === rg.c2 ? rg.c1 - 1 : rg.c1;
      if (srcC < 0) return;
      for (let r = rg.r1; r <= rg.r2; r++) {
        if (filterHidden(r)) continue;
        const d = cellData(wb.getCell(si, r, srcC));
        for (let c = srcC + 1; c <= rg.c2; c++) wb.setCellData(si, r, c, shift(d, 0, c - srcC));
      }
    } else if (dir === 'up') {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const d = cellData(wb.getCell(si, rg.r2, c));
        for (let r = rg.r1; r < rg.r2; r++) if (!filterHidden(r)) wb.setCellData(si, r, c, shift(d, r - rg.r2, 0));
      }
    } else if (dir === 'left') {
      for (let r = rg.r1; r <= rg.r2; r++) {
        if (filterHidden(r)) continue;
        const d = cellData(wb.getCell(si, r, rg.c2));
        for (let c = rg.c1; c < rg.c2; c++) wb.setCellData(si, r, c, shift(d, 0, c - rg.c2));
      }
    }
  }, meta());
}

// ───────────────────────── 클립보드 ─────────────────────────
function displayText(r, c, s = si) {
  return formatValue(wb.getValue(s, r, c), wb.styleAt(s, r, c)).text;
}

function copySelection(cut) {
  const full = selKind === 'cells' ? { ...sel } : usedClip(sel);
  if ((full.r2 - full.r1 + 1) * (full.c2 - full.c1 + 1) > 2_000_000) { toast('복사하기에는 선택 영역이 너무 큽니다.'); return { text: '', html: '' }; }
  const data = [];
  const values = [];
  const text = [];
  const rowsIncluded = [];
  for (let r = full.r1; r <= full.r2; r++) {
    if (filterHidden(r)) continue;
    rowsIncluded.push(r);
    const drow = [];
    const vrow = [];
    const trow = [];
    for (let c = full.c1; c <= full.c2; c++) {
      drow.push(cellData(wb.getCell(si, r, c)));
      vrow.push(valueAt(r, c));
      trow.push(displayText(r, c));
    }
    data.push(drow);
    values.push(vrow);
    text.push(trow);
  }
  clip = { si, ...full, r2: full.r1 + rowsIncluded.length - 1, rows: rowsIncluded, data, values, cut, text: toDelimited(text, '\t', '\n') };
  const html = `<table>${text.map((row, i) => `<tr>${row.map((t, j) => {
    const st = data[i][j]?.style ?? {};
    const css = [st.bold && 'font-weight:bold', st.italic && 'font-style:italic', st.color && `color:${st.color}`, st.fill && `background:${st.fill}`].filter(Boolean).join(';');
    return `<td${css ? ` style="${css}"` : ''}>${escapeHtml(t)}</td>`;
  }).join('')}</tr>`).join('')}</table>`;
  updateSelectionUI();
  setMode();
  return { text: clip.text, html };
}

function valueToRaw(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') return String(v);
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (isError(v)) return v.code;
  const p = parseInput(v);
  return typeof p.value === 'string' && !v.startsWith('=') ? v : `'${v}`;
}

function pasteInternal(mode = 'all') {
  if (!clip) return;
  const transpose = mode === 'transpose';
  const h = clip.data.length;
  const w = clip.c2 - clip.c1 + 1;
  const ph = transpose ? w : h;
  const pw = transpose ? h : w;
  const tgt = selKind === 'cells' ? sel : { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
  const sh = tgt.r2 - tgt.r1 + 1;
  const sw = tgt.c2 - tgt.c1 + 1;
  let th = ph;
  let tw = pw;
  if (!clip.cut && sh % ph === 0 && sw % pw === 0 && sh * sw <= BIG_AREA) { th = sh; tw = sw; }
  const cut = clip.cut;
  const src = clip;
  wb.transact(() => {
    if (cut) for (const r of src.rows) for (let c = src.c1; c <= src.c2; c++) wb.setCellData(src.si, r, c, null);
    for (let i = 0; i < th; i++) {
      for (let j = 0; j < tw; j++) {
        const a = i % ph;
        const b = j % pw;
        const [ci, cj] = transpose ? [b, a] : [a, b];
        const d = src.data[ci][cj];
        const tr = tgt.r1 + i;
        const tc = tgt.c1 + j;
        const cur = cellData(wb.getCell(si, tr, tc));
        const srcR = src.rows[ci];
        const shifted = d?.raw.startsWith('=') && !cut ? shiftFormula(d.raw, tr - srcR, tc - (src.c1 + cj)) : d?.raw ?? '';
        let data;
        if (mode === 'values') data = { raw: valueToRaw(src.values[ci][cj]), style: cur?.style, comment: cur?.comment };
        else if (mode === 'formats') data = { raw: cur?.raw ?? '', style: d?.style, comment: cur?.comment };
        else if (mode === 'formulas') data = { raw: shifted, style: cur?.style, comment: cur?.comment };
        else data = d ? { ...d, raw: shifted } : null;
        wb.setCellData(si, tr, tc, data);
      }
    }
    afterDataEntry({ r1: tgt.r1, c1: tgt.c1, r2: tgt.r1 + th - 1, c2: tgt.c1 + tw - 1 });
  }, meta());
  if (cut) clip = null;
  selectRange({ r1: tgt.r1, c1: tgt.c1, r2: tgt.r1 + th - 1, c2: tgt.c1 + tw - 1 }, 'cells', { r: tgt.r1, c: tgt.c1 });
  setMode();
}

function pasteText(text) {
  const rows = parseDelimited(text.replace(/\r\n?/g, '\n').replace(/\n$/, ''), '\t');
  if (!rows.length) return;
  wb.transact(() => {
    if (rows.length === 1 && rows[0].length === 1 && !selIsActiveOnly() && selKind === 'cells') {
      for (const [r, c] of cellsIn(sel)) wb.setInput(si, r, c, rows[0][0]);
    } else {
      rows.forEach((row, i) => row.forEach((v, j) => wb.setInput(si, active.r + i, active.c + j, v)));
      afterDataEntry({ r1: active.r, c1: active.c, r2: active.r + rows.length - 1, c2: active.c + Math.max(...rows.map((x) => x.length)) - 1 });
    }
  }, meta());
  if (rows.length > 1 || rows[0].length > 1) {
    selectRange({ r1: active.r, c1: active.c, r2: active.r + rows.length - 1, c2: active.c + Math.max(...rows.map((r) => r.length)) - 1 });
  }
}

function handlePaste(text) {
  const n = (s) => (s ?? '').replace(/\r\n?/g, '\n').replace(/\n$/, '');
  if (clip && (!text || n(text) === n(clip.text))) { pasteInternal('all'); return; }
  if (text) pasteText(text);
}

async function pasteFromButton(mode = 'all') {
  if (mode !== 'all') { if (clip) pasteInternal(mode); else toast('복사한 셀이 없습니다.'); return; }
  let text = null;
  try { text = await navigator.clipboard.readText(); } catch { /* 권한 없음 */ }
  if (text === null && !clip) { toast('Ctrl+V 를 눌러 붙여넣으세요.'); return; }
  handlePaste(text);
}

// ───────────────────────── 서식 명령 ─────────────────────────
const LINE_KINDS = new Set(['cols', 'rows', 'all']);
const BOOL_KEYS = new Set(['bold', 'italic', 'underline', 'strike', 'wrap', 'bt', 'bb', 'bl', 'br']);

/** 열/행/시트 전체 서식을 상속한 셀에서 서식을 끌 때는 false 로 명시해야 함 */
function explicitOff(patch, r, c) {
  if (!wb.hasLineStyle(si, r, c)) return patch;
  const s = sheet();
  const line = { ...s.allStyle, ...s.colStyles[c], ...s.rowStyles[r] };
  const out = { ...patch };
  for (const [k, v] of Object.entries(patch)) if (v === undefined && BOOL_KEYS.has(k) && line[k]) out[k] = false;
  return out;
}

function applyStyle(patchOrFn, { widen = false } = {}) {
  const rg = sel;
  const patchFor = (cur) => {
    const p = typeof patchOrFn === 'function' ? patchOrFn(cur) : patchOrFn;
    // 기본 표시 형식을 고르면 사용자 지정 코드는 지움
    return p && 'numFmt' in p && p.numFmt !== 'custom' ? { ...p, code: undefined } : p;
  };
  wb.transact(() => {
    if (LINE_KINDS.has(selKind)) {
      const s = sheet();
      if (selKind === 'all') {
        wb.setLineStyle(si, 'all', 0, patchFor(s.allStyle ?? {}) ?? {});
        for (const k of Object.keys(s.colStyles)) wb.setLineStyle(si, 'col', Number(k), patchFor(s.colStyles[k]) ?? {});
        for (const k of Object.keys(s.rowStyles)) wb.setLineStyle(si, 'row', Number(k), patchFor(s.rowStyles[k]) ?? {});
      } else if (selKind === 'cols') {
        for (let c = rg.c1; c <= rg.c2; c++) wb.setLineStyle(si, 'col', c, patchFor(s.colStyles[c] ?? {}) ?? {});
      } else {
        for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 100000); r++) wb.setLineStyle(si, 'row', r, patchFor(s.rowStyles[r] ?? {}) ?? {});
      }
      for (const [k, cell] of [...s.cells]) {
        const [r, c] = k.split(',').map(Number);
        if (!inSel(r, c) || !cell.style) continue;
        const p = patchFor(cell.style);
        if (p) wb.setStyle(si, r, c, p);
      }
    } else {
      for (const [r, c] of cellsIn(rg)) {
        const cur = wb.getCell(si, r, c)?.style ?? {};
        const p = patchFor(cur);
        if (p) wb.setStyle(si, r, c, explicitOff(p, r, c));
      }
    }
    if (widen) autoWiden(rg, { grow: widen === 'grow' });
    const sample = patchFor({}) ?? {};
    if (('wrap' in sample || 'size' in sample || 'font' in sample) && selKind !== 'cols' && selKind !== 'all') {
      const u = usedClip(rg);
      autoFitRows(u.r1, Math.min(u.r2, u.r1 + 1000));
    }
  }, meta());
}

const toggleStyle = (key) => {
  const next = !styleAt(active.r, active.c)[key];
  applyStyle({ [key]: next || undefined });
};

function applyBorder(kind) {
  lastBorder = kind;
  const rg = selKind === 'cells' ? sel : usedClip(sel);
  wb.transact(() => {
    for (const [r, c] of cellsIn(rg)) {
      let p = null;
      switch (kind) {
        case 'none': p = { bt: undefined, bb: undefined, bl: undefined, br: undefined }; break;
        case 'all': p = { bt: true, bb: true, bl: true, br: true }; break;
        case 'outside': p = {
          ...(r === rg.r1 && { bt: true }), ...(r === rg.r2 && { bb: true }),
          ...(c === rg.c1 && { bl: true }), ...(c === rg.c2 && { br: true }),
        }; break;
        case 'bottom': if (r === rg.r2) p = { bb: true }; break;
        case 'top': if (r === rg.r1) p = { bt: true }; break;
        case 'left': if (c === rg.c1) p = { bl: true }; break;
        case 'right': if (c === rg.c2) p = { br: true }; break;
        case 'topBottom': p = { ...(r === rg.r1 && { bt: true }), ...(r === rg.r2 && { bb: true }) }; break;
        default:
      }
      if (p && Object.keys(p).length) wb.setStyle(si, r, c, p);
    }
    if (kind === 'none') {
      // 이웃 셀이 그린 맞닿은 테두리도 제거
      if (rg.r1 > 0) for (let c = rg.c1; c <= rg.c2; c++) if (styleAt(rg.r1 - 1, c).bb) wb.setStyle(si, rg.r1 - 1, c, { bb: undefined });
      if (rg.c1 > 0) for (let r = rg.r1; r <= rg.r2; r++) if (styleAt(r, rg.c1 - 1).br) wb.setStyle(si, r, rg.c1 - 1, { br: undefined });
      for (let c = rg.c1; c <= rg.c2; c++) if (styleAt(rg.r2 + 1, c).bt) wb.setStyle(si, rg.r2 + 1, c, { bt: undefined });
      for (let r = rg.r1; r <= rg.r2; r++) if (styleAt(r, rg.c2 + 1).bl) wb.setStyle(si, r, rg.c2 + 1, { bl: undefined });
    }
  }, meta());
}

function changeDecimals(delta) {
  const st = styleAt(active.r, active.c);
  let v = valueAt(active.r, active.c);
  if (typeof v !== 'number') {
    for (const [r, c] of cellsIn(usedClip(sel))) {
      const x = valueAt(r, c);
      if (typeof x === 'number') { v = x; break; }
    }
  }
  if (typeof v !== 'number') v = 0;
  if (st.numFmt === 'custom' && st.code) {
    applyStyle((s) => (s.numFmt === 'custom' && s.code ? { code: adjustCodeDecimals(s.code, delta) } : null), { widen: true });
    return;
  }
  const next = clamp(displayedDecimals(v, st.numFmt, st.decimals) + delta, 0, 15);
  applyStyle((s) => ({ decimals: next, numFmt: s.numFmt === 'text' ? undefined : s.numFmt ?? st.numFmt }), { widen: true });
}

function changeFontSize(dir) {
  const cur = styleAt(active.r, active.c).size || DEFAULT_SIZE;
  const next = dir > 0 ? FONT_SIZES.find((s) => s > cur) ?? cur + 4 : [...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  applyStyle({ size: next === DEFAULT_SIZE ? undefined : next });
}

function capturePainter(sticky) {
  const full = usedClip(sel);
  const h = Math.min(full.r2 - full.r1 + 1, 500);
  const w = Math.min(full.c2 - full.c1 + 1, 200);
  const styles = [];
  for (let i = 0; i < h; i++) {
    const row = [];
    for (let j = 0; j < w; j++) row.push({ ...styleAt(full.r1 + i, full.c1 + j) });
    styles.push(row);
  }
  painter = { styles, sticky };
  dom.view.classList.add('painting');
  setMode();
  updateRibbon();
}

function applyPainter() {
  const { styles } = painter;
  const h = styles.length;
  const w = styles[0].length;
  const tgt = sel;
  const th = selIsActiveOnly() ? h : Math.min(tgt.r2 - tgt.r1 + 1, 5000);
  const tw = selIsActiveOnly() ? w : Math.min(tgt.c2 - tgt.c1 + 1, 500);
  wb.transact(() => {
    for (let i = 0; i < th; i++) {
      for (let j = 0; j < tw; j++) {
        const cur = wb.getCell(si, tgt.r1 + i, tgt.c1 + j);
        wb.setCellData(si, tgt.r1 + i, tgt.c1 + j, { raw: cur?.raw ?? '', style: styles[i % h][j % w], comment: cur?.comment });
      }
    }
  }, meta());
  if (th !== tgt.r2 - tgt.r1 + 1 || tw !== tgt.c2 - tgt.c1 + 1) {
    selectRange({ r1: tgt.r1, c1: tgt.c1, r2: tgt.r1 + th - 1, c2: tgt.c1 + tw - 1 }, 'cells', active);
  }
  if (!painter.sticky) {
    painter = null;
    dom.view.classList.remove('painting');
  }
  setMode();
  updateRibbon();
}

function toggleMerge(kind = 'center') {
  const rg = sel;
  if (selKind !== 'cells') { toast('셀 범위를 선택한 후 병합하세요.'); return; }
  const same = wb.mergesIn(si, rg.r1, rg.c1, rg.r2, rg.c2)
    .some((m) => m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2);
  if (kind === 'unmerge' || (kind === 'center' && same)) {
    wb.transact(() => wb.unmerge(si, rg.r1, rg.c1, rg.r2, rg.c2), meta());
    return;
  }
  if (isSingle(rg)) return;
  const doMerge = () => {
    wb.transact(() => {
      if (kind === 'across') {
        for (let r = rg.r1; r <= rg.r2; r++) wb.merge(si, r, rg.c1, r, rg.c2);
      } else {
        wb.merge(si, rg.r1, rg.c1, rg.r2, rg.c2);
        if (kind === 'center') wb.setStyle(si, rg.r1, rg.c1, { align: 'center', valign: 'middle' });
      }
    }, meta());
    selectRange(rg, 'cells', { r: rg.r1, c: rg.c1 });
  };
  let filled = 0;
  for (const [r, c] of cellsIn(usedClip(rg))) if (!isEmptyAt(r, c)) filled++;
  if (filled > 1 && kind !== 'across') {
    openDialog({
      title: 'Tabula',
      body: '셀을 병합하면 왼쪽 위의 값만 유지되고 다른 값은 삭제됩니다.',
      buttons: [{ label: '확인', primary: true, action: doMerge }, { label: '취소' }],
    });
  } else {
    doMerge();
  }
}

// ───────────────────────── 숨기기 ─────────────────────────
function hideSel(axis, hide) {
  const [a, b] = axis === 'row' ? [sel.r1, sel.r2] : [sel.c1, sel.c2];
  const limit = axis === 'row' ? MAX_ROWS : MAX_COLS;
  if (hide && a === 0 && b >= limit - 1) { toast('모든 행/열을 숨길 수는 없습니다.'); return; }
  let idx;
  if (hide) idx = range(a, Math.min(b, a + 20000));
  else {
    const map = axis === 'row' ? sheet().hiddenRows : sheet().hiddenCols;
    const pad = a === b ? 1 : 0;
    idx = Object.keys(map).map(Number).filter((i) => i >= a - pad && i <= b + pad);
    if (!idx.length) { toast('숨겨진 항목이 없습니다. 숨겨진 부분 양쪽을 선택한 후 다시 시도하세요.'); return; }
  }
  wb.transact(() => wb.setHidden(si, axis, idx, hide), meta());
  gv.layout();
  if (hide) {
    const n = axis === 'row' ? stepFrom({ r: b, c: active.c }, 1, 0) : stepFrom({ r: active.r, c: b }, 0, 1);
    selectCell(n.r, n.c);
  }
}

// ───────────────────────── 틀 고정 ─────────────────────────
function setFreeze(rows, cols) {
  wb.transact(() => wb.setSheetProp(si, 'freeze', { rows, cols }), meta());
  gv.layout();
  gv.setScroll(0, 0);
  toast(rows || cols ? `틀 고정: ${rows ? `${rows}행` : ''}${rows && cols ? ', ' : ''}${cols ? `${cols}열` : ''}` : '틀 고정을 취소했습니다.');
}

// ───────────────────────── 데이터 명령 ─────────────────────────
function currentRegion(r, c) {
  const rg = { r1: r, c1: c, r2: r, c2: c };
  const filledRow = (row, c1, c2) => { for (let x = Math.max(0, c1); x <= c2; x++) if (!isEmptyAt(row, x)) return true; return false; };
  const filledCol = (col, r1, r2) => { for (let y = Math.max(0, r1); y <= r2; y++) if (!isEmptyAt(y, col)) return true; return false; };
  for (let changed = true, guard = 0; changed && guard < 200000; guard++) {
    changed = false;
    if (rg.r1 > 0 && filledRow(rg.r1 - 1, rg.c1 - 1, rg.c2 + 1)) { rg.r1--; changed = true; }
    if (rg.r2 < MAX_ROWS - 1 && filledRow(rg.r2 + 1, rg.c1 - 1, rg.c2 + 1)) { rg.r2++; changed = true; }
    if (rg.c1 > 0 && filledCol(rg.c1 - 1, rg.r1 - 1, rg.r2 + 1)) { rg.c1--; changed = true; }
    if (rg.c2 < MAX_COLS - 1 && filledCol(rg.c2 + 1, rg.r1 - 1, rg.r2 + 1)) { rg.c2++; changed = true; }
  }
  return rg;
}

function dataRange() {
  if (!selIsActiveOnly()) return usedClip(sel);
  return currentRegion(active.r, active.c);
}

function hasHeader(rg) {
  if (rg.r2 <= rg.r1) return false;
  let textFirst = true;
  let nonTextBelow = false;
  for (let c = rg.c1; c <= rg.c2; c++) {
    const v = valueAt(rg.r1, c);
    if (v !== null && typeof v !== 'string') textFirst = false;
    const below = valueAt(rg.r1 + 1, c);
    if (below !== null && typeof below !== 'string') nonTextBelow = true;
  }
  return textFirst && (nonTextBelow || styleAt(rg.r1, rg.c1).bold === true);
}

function sortData(ascending, keyCol = active.c, header = null, rgIn = null, fkeyIn = null) {
  const tbl = rgIn ? null : tableHere();
  let fkey = fkeyIn ?? (tbl ? (tbl.filter && tbl.header ? tbl.id : null) : sheet().filter ? '' : null);
  const f = fkey === null ? null : getFilter(fkey);
  let rg = rgIn;
  if (!rg && tbl && selIsActiveOnly()) {
    // 표 안: 표의 데이터 행만 정렬 (요약 행 제외)
    rg = { r1: tbl.header ? tbl.r1 : dataTop(tbl), c1: tbl.c1, r2: dataBottom(tbl), c2: tbl.c2 };
    header = tbl.header;
  } else if (!rg && f && fkey === '' && active.r >= f.r1 && active.r <= f.r2 && active.c >= f.c1 && active.c <= f.c2 && selIsActiveOnly()) {
    rg = recomputeFilter(f);
    header = true;
  }
  rg ??= dataRange();
  if (rg.r2 <= rg.r1 && isSingle(rg)) return;
  const h = header ?? hasHeader(rg);
  const key = clamp(keyCol, rg.c1, rg.c2);
  if (wb.mergesIn(si, rg.r1, rg.c1, rg.r2, rg.c2).length) {
    alertDialog('Tabula', '병합된 셀이 있으면 정렬할 수 없습니다.');
    return;
  }
  wb.transact(() => {
    wb.sortRange(si, rg.r1 + (h ? 1 : 0), rg.c1, rg.r2, rg.c2, key, ascending);
    if (f && rg.r1 === f.r1 && rg.c1 === f.c1) putFilter(fkey, recomputeFilter({ ...f, sort: { col: key, asc: ascending } }, fkey));
  }, meta());
  if (!rgIn) selectRange({ r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, 'cells', { r: active.r, c: active.c });
}

function removeDuplicates() {
  const rg = dataRange();
  const h = hasHeader(rg);
  const start = rg.r1 + (h ? 1 : 0);
  const seen = new Set();
  const keep = [];
  for (let r = start; r <= rg.r2; r++) {
    const vals = [];
    for (let c = rg.c1; c <= rg.c2; c++) {
      const v = valueAt(r, c);
      vals.push(typeof v === 'string' ? v.toLowerCase() : isError(v) ? v.code : v);
    }
    const k = JSON.stringify(vals);
    if (seen.has(k)) continue;
    seen.add(k);
    keep.push(r);
  }
  const removed = rg.r2 - start + 1 - keep.length;
  if (removed > 0) {
    wb.transact(() => {
      const rows = keep.map((r) => {
        const cells = [];
        for (let c = rg.c1; c <= rg.c2; c++) cells.push(cellData(wb.getCell(si, r, c)));
        return { r, cells };
      });
      for (let r = start; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) wb.setCellData(si, r, c, null);
      rows.forEach((row, i) => row.cells.forEach((d, j) => {
        const tr = start + i;
        wb.setCellData(si, tr, rg.c1 + j, d && { ...d, raw: d.raw.startsWith('=') ? shiftFormula(d.raw, tr - row.r, 0) : d.raw });
      }));
    }, meta());
  }
  alertDialog('중복된 항목 제거', `중복된 값 ${removed}개를 찾아 제거했습니다. 고유한 값 ${keep.length}개가 남아 있습니다.`);
}

function autoSum(fn = 'SUM') {
  if (selIsActiveOnly()) {
    const { r, c } = active;
    let ref = '';
    let r0 = r - 1;
    while (r0 >= 0 && typeof valueAt(r0, c) === 'number') r0--;
    if (r0 < r - 1) ref = `${cellName(r0 + 1, c)}:${cellName(r - 1, c)}`;
    else {
      let c0 = c - 1;
      while (c0 >= 0 && typeof valueAt(r, c0) === 'number') c0--;
      if (c0 < c - 1) ref = `${cellName(r, c0 + 1)}:${cellName(r, c - 1)}`;
    }
    const text = `=${fn}(${ref})`;
    startEdit('enter', text, { caret: fn.length + 2 + ref.length });
    if (ref) {
      const pa = parseRangeName(ref);
      editing.point = { start: fn.length + 2, end: fn.length + 2 + ref.length, anchor: { r: pa.r1, c: pa.c1 }, cur: { r: pa.r2, c: pa.c2 }, rg: pa };
      dom.editor.setSelectionRange(fn.length + 2, fn.length + 2 + ref.length);
      setMode();
    }
    return;
  }
  const rg = usedClip(sel);
  const lastRowEmpty = (() => { for (let c = rg.c1; c <= rg.c2; c++) if (!isEmptyAt(rg.r2, c)) return false; return true; })();
  const target = lastRowEmpty ? rg.r2 : rg.r2 + 1;
  wb.transact(() => {
    for (let c = rg.c1; c <= rg.c2; c++) {
      let any = false;
      for (let r = rg.r1; r < target; r++) if (typeof valueAt(r, c) === 'number') any = true;
      if (any) wb.setInput(si, target, c, `=${fn}(${cellName(rg.r1, c)}:${cellName(target - 1, c)})`);
    }
  }, meta());
  selectRange({ r1: rg.r1, c1: rg.c1, r2: target, c2: rg.c2 }, 'cells', active);
}

function insertFunctionText(name) {
  if (!editing) {
    startEdit('enter', `=${name}()`, { caret: name.length + 2 });
  } else {
    const inp = edInput();
    const pos = inp.selectionStart;
    const prefix = inp.value.length === 0 ? '=' : '';
    const ins = `${prefix}${name}()`;
    setEditText(inp.value.slice(0, pos) + ins + inp.value.slice(inp.selectionEnd), pos + ins.length - 1);
  }
  toast(FUNC_INFO[name]?.sig ?? name);
  afterEditInput();
}

// ───────────────────────── 필터 ─────────────────────────
/** 필터 조건으로 숨길 행 다시 계산 (데이터 아래로 늘어난 행 포함) */
function recomputeFilter(f, key = '') {
  const r2 = key ? f.r2 : Math.max(f.r2, currentRegion(f.r1, f.c1).r2);
  const crit = Object.entries(f.criteria ?? {}).filter(([, v]) => Array.isArray(v)).map(([c, vals]) => [Number(c), new Set(vals)]);
  const hidden = {};
  if (crit.length) {
    for (let r = f.r1 + 1; r <= r2; r++) {
      for (const [c, allowed] of crit) {
        if (!allowed.has(displayText(r, c))) { hidden[r] = true; break; }
      }
    }
  }
  return { ...f, r2, hidden };
}

function toggleFilter() {
  if (tableHere()) { toggleTableOption('filter'); return; }
  if (sheet().filter) {
    wb.transact(() => wb.setSheetProp(si, 'filter', null), meta());
    toast('필터를 해제했습니다.');
    return;
  }
  const rg = dataRange();
  if (rg.r1 === rg.r2 && isEmptyAt(rg.r1, rg.c1)) { alertDialog('Tabula', '필터를 적용할 데이터 범위를 선택하세요.'); return; }
  wb.transact(() => {
    wb.setSheetProp(si, 'filter', { ...rg, criteria: {}, hidden: {} });
    widenForFilterButtons(rg);
  }, meta());
}

function widenForFilterButtons(rg) {
  // 필터 단추가 머리글을 가리지 않도록 좁은 열은 넓힘
  for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 100); c++) {
    const need = measureText(displayText(rg.r1, c), styleAt(rg.r1, c)) + 28;
    if (need > wb.colWidth(si, c) && need < 300) wb.setColWidth(si, c, Math.ceil(need));
  }
}

function applyFilterCriteria(c, values, key = '', { quiet = false } = {}) {
  const f = getFilter(key);
  if (!f) return;
  const criteria = { ...f.criteria };
  if (values === null) delete criteria[c]; else criteria[c] = values;
  const nf = recomputeFilter({ ...f, criteria }, key);
  wb.transact(() => putFilter(key, nf), meta());
  gv.layout();
  const total = nf.r2 - nf.r1;
  if (quiet) return;
  toast(`${total}개 중 ${total - Object.keys(nf.hidden).length}개의 레코드가 있습니다.`);
  selectCell(nf.r1, c);
  setMode();
}

function openFilterMenu(c, anchorEl, key = '') {
  const f = getFilter(key);
  if (!f) return;
  const full = recomputeFilter(f, key);
  // 다른 열 조건을 통과한 행의 값만 목록에 표시
  const others = Object.entries(f.criteria ?? {}).filter(([k, v]) => Number(k) !== c && Array.isArray(v)).map(([k, v]) => [Number(k), new Set(v)]);
  const values = new Map();
  for (let r = f.r1 + 1; r <= full.r2; r++) {
    if (others.some(([k, allowed]) => !allowed.has(displayText(r, k)))) continue;
    const t = displayText(r, c);
    if (!values.has(t)) values.set(t, valueAt(r, c));
  }
  const items = [...values.entries()].sort((a, b) => {
    if (a[0] === '') return 1;
    if (b[0] === '') return -1;
    const x = a[1];
    const y = b[1];
    if (typeof x === 'number' && typeof y === 'number') return x - y;
    if (typeof x === 'number') return -1;
    if (typeof y === 'number') return 1;
    return String(a[0]).localeCompare(String(b[0]), 'ko');
  }).map(([t]) => t);
  const current = Array.isArray(f.criteria?.[c]) ? new Set(f.criteria[c]) : null;
  const checks = new Map();
  const list = el('div', { class: 'filter-list' });
  const all = el('input', { type: 'checkbox' });
  const syncAll = () => {
    const vis = [...checks.values()].filter((cb) => cb.parentElement.style.display !== 'none');
    all.checked = vis.length > 0 && vis.every((cb) => cb.checked);
    all.indeterminate = !all.checked && vis.some((cb) => cb.checked);
  };
  list.append(el('label', {}, all, '(모두 선택)'));
  for (const t of items) {
    const cb = el('input', { type: 'checkbox', checked: !current || current.has(t), onchange: syncAll });
    checks.set(t, cb);
    list.append(el('label', {}, cb, t === '' ? '(필드 값 없음)' : t));
  }
  all.addEventListener('change', () => {
    for (const cb of checks.values()) if (cb.parentElement.style.display !== 'none') cb.checked = all.checked;
  });
  syncAll();
  const search = el('input', { type: 'search', placeholder: '검색' });
  search.addEventListener('input', () => {
    const q = search.value.trim().toLowerCase();
    for (const [t, cb] of checks) {
      const show = !q || t.toLowerCase().includes(q);
      cb.parentElement.style.display = show ? '' : 'none';
      if (q) cb.checked = show;
    }
    syncAll();
  });
  const ok = () => {
    closeMenus();
    const chosen = [...checks.entries()].filter(([, cb]) => cb.checked).map(([t]) => t);
    applyFilterCriteria(c, chosen.length === items.length ? null : chosen, key);
    focusGrid();
  };
  const header = displayText(f.r1, c) || `${colToName(c)}열`;
  const node = el('div', {
    class: 'filter-menu',
    onkeydown: (e) => { e.stopPropagation(); if (e.key === 'Enter') ok(); if (e.key === 'Escape') { closeMenus(); focusGrid(); } },
  },
  search, list,
  el('div', { class: 'filter-foot' },
    el('button', { class: 'btn primary', onclick: ok }, '확인'),
    el('button', { class: 'btn', onclick: () => { closeMenus(); focusGrid(); } }, '취소')));
  const menu = openMenu(anchorEl, [
    { label: '텍스트 오름차순 정렬', icon: 'sortAsc', action: () => sortData(true, c, true, full, key) },
    { label: '텍스트 내림차순 정렬', icon: 'sortDesc', action: () => sortData(false, c, true, full, key) },
    { sep: true },
    { label: `"${header}"에서 필터 해제`, icon: 'filterClear', disabled: !current, action: () => applyFilterCriteria(c, null, key) },
    { node },
  ]);
  menu.style.minWidth = '280px';
  setTimeout(() => search.focus());
}

// ───────────────────────── 차트 ─────────────────────────
function insertChart(type) {
  let rg = dataRange();
  if (rg.r2 - rg.r1 > 2000) rg = { ...rg, r2: rg.r1 + 2000 };
  let hasNum = false;
  for (const [r, c] of cellsIn(rg)) if (typeof valueAt(r, c) === 'number') { hasNum = true; break; }
  if (!hasNum) { alertDialog('차트 삽입', '차트를 만들려면 숫자가 들어 있는 데이터 범위를 선택하세요.'); return; }
  const box = gv.sheetRect(rg);
  const vis = gv.screenRect(rg);
  let x = box.x + box.w + 24;
  let y = box.y;
  if (vis.x + box.w + 24 + 480 > gv.viewW + 200) { x = box.x; y = box.y + box.h + 16; }
  const id = `ch${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;
  const pieTitle = type === 'pie' || type === 'doughnut' ? displayText(rg.r1, rg.c1 + 1) : '';
  const chart = { id, type, title: pieTitle || '차트 제목', range: { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, x, y, w: 480, h: 288, z: nextZ() };
  wb.transact(() => wb.setSheetProp(si, 'charts', [...sheet().charts.map((c) => ({ ...c })), chart]), meta());
  chartSel = id;
  gv.ensureVisible(gv.rows.indexAt(y + 100), gv.cols.indexAt(x + 200));
  gv.renderObjectsAll();
  updateSelectionUI();
}

function updateChart(id, patch) {
  wb.transact(() => wb.setSheetProp(si, 'charts', sheet().charts.map((c) => (c.id === id ? { ...c, ...patch } : { ...c }))), meta());
}

function deleteChart(id) { deleteObject(id); }

function chartDialog(id) {
  const ch = sheet().charts.find((c) => c.id === id);
  if (!ch) return;
  const preview = el('div', { style: { border: '1px solid #ddd', height: '200px', display: 'grid', placeItems: 'center', overflow: 'hidden' } });
  const typeSel = el('select', {}, CHART_TYPES.map((t) => el('option', { value: t.id, selected: t.id === ch.type }, t.label)));
  const titleIn = el('input', { type: 'text', value: ch.title ?? '' });
  const rangeIn = el('input', { type: 'text', value: `${cellName(ch.range.r1, ch.range.c1)}:${cellName(ch.range.r2, ch.range.c2)}` });
  const draw = () => {
    const rg = parseRangeName(rangeIn.value);
    if (!rg) return;
    const rows = [];
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 300); r++) {
      const row = [];
      for (let c = rg.c1; c <= Math.min(rg.c2, rg.c1 + 50); c++) row.push(valueAt(r, c));
      rows.push(row);
    }
    preview.innerHTML = renderChartSvg({ type: typeSel.value, title: titleIn.value, w: 400, h: 200 }, chartData(rows, typeSel.value));
  };
  [typeSel, titleIn, rangeIn].forEach((i) => i.addEventListener('input', draw));
  draw();
  openDialog({
    title: '차트 편집', width: 460,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
      el('label', {}, el('span', {}, '차트 종류'), typeSel),
      el('label', {}, el('span', {}, '차트 제목'), titleIn),
      el('label', {}, el('span', {}, '데이터 범위'), rangeIn),
      preview),
    buttons: [
      {
        label: '확인', primary: true, action: () => {
          const rg = parseRangeName(rangeIn.value);
          if (!rg) { toast('데이터 범위가 올바르지 않습니다.'); return false; }
          updateChart(id, { type: typeSel.value, title: titleIn.value, range: rg });
          return true;
        },
      },
      { label: '삭제', action: () => deleteChart(id) },
      { label: '취소' },
    ],
  });
}

function chartMenu(id, pos) { objectMenu(id, pos); }

// ───────────────────────── 표 (Ctrl+T) ─────────────────────────
const tableHere = (r = active.r, c = active.c) => tableAt(sheet(), r, c);
const newTableId = () => `tb${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;

/** transact 안에서 호출 */
function setTables(fn) {
  wb.setSheetProp(si, 'tables', fn((sheet().tables ?? []).map((t) => ({ ...t }))));
}

function updateTable(id, patch) {
  wb.transact(() => setTables((list) => list.map((t) => (t.id === id ? { ...t, ...patch } : t))), meta());
}

function tableRangeText(t) {
  return `=$${colToName(t.c1)}$${t.r1 + 1}:$${colToName(t.c2)}$${t.r2 + 1}`;
}

function createTableDialog(styleName = null) {
  if (editing && !commitEdit()) return;
  const existing = tableHere();
  if (existing) {
    if (styleName) updateTable(existing.id, { style: styleName });
    else toast(`이미 '${existing.name}' 표 안에 있습니다.`);
    return;
  }
  let rg = selIsActiveOnly() ? dataRange() : usedClip(sel);
  if (isSingle(rg) && isEmptyAt(rg.r1, rg.c1)) rg = { ...rg };
  const rangeIn = el('input', { type: 'text', value: tableRangeText(rg), style: { width: '100%' } });
  const headIn = el('input', { type: 'checkbox', checked: rg.r2 > rg.r1 ? hasHeader(rg) || typeof valueAt(rg.r1, rg.c1) === 'string' : false });
  openDialog({
    title: '표 만들기', width: 380,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
      el('div', {}, '표에 사용할 데이터를 지정하십시오.'),
      rangeIn,
      el('label', { class: 'fc-check' }, headIn, '머리글 포함')),
    buttons: [
      {
        label: '확인', primary: true,
        action: () => {
          const t = rangeIn.value.replace(/[=$]/g, '').trim();
          const p = parseRangeName(t.includes('!') ? t.slice(t.lastIndexOf('!') + 1) : t);
          if (!p) { alertDialog('표 만들기', '참조가 올바르지 않습니다.'); return false; }
          return createTable(p, headIn.checked, styleName ?? DEFAULT_TABLE_STYLE) ? undefined : false;
        },
      },
      { label: '취소' },
    ],
  });
}

function createTable(rg, header, styleName) {
  const s = sheet();
  if ((s.tables ?? []).some((t) => t.r1 <= rg.r2 && t.r2 >= rg.r1 && t.c1 <= rg.c2 && t.c2 >= rg.c1)) {
    alertDialog('표 만들기', '표는 다른 표와 겹칠 수 없습니다.');
    return false;
  }
  if (wb.mergesIn(si, rg.r1, rg.c1, rg.r2, rg.c2).length) {
    alertDialog('표 만들기', '병합된 셀이 있으면 표를 만들 수 없습니다. 병합을 해제하고 다시 시도하세요.');
    return false;
  }
  if (rg.c2 - rg.c1 > 500) { alertDialog('표 만들기', '표는 500열까지 만들 수 있습니다.'); return false; }
  const id = newTableId();
  const name = nextTableName(wb);
  wb.transact(() => {
    let { r1, r2 } = rg;
    if (!header) {
      // 머리글 행을 새로 넣고 데이터는 한 행 아래로
      wb.insertRows(si, r1, 1);
      r2++;
      for (let c = rg.c1; c <= rg.c2; c++) wb.setInput(si, r1, c, `열${c - rg.c1 + 1}`);
    } else {
      const texts = [];
      for (let c = rg.c1; c <= rg.c2; c++) texts.push(displayText(r1, c));
      uniqueNames(texts).forEach((n, i) => {
        const c = rg.c1 + i;
        if (typeof valueAt(r1, c) !== 'string' || texts[i] !== n) wb.setInput(si, r1, c, `'${n}`);
      });
    }
    if (r2 === r1) r2++; // 머리글만 있으면 빈 데이터 행 하나
    const f = sheet().filter;
    if (f && f.r1 <= r2 && f.r2 >= r1 && f.c1 <= rg.c2 && f.c2 >= rg.c1) wb.setSheetProp(si, 'filter', null);
    const t = {
      id, name, r1, c1: rg.c1, r2, c2: rg.c2, header: true, totals: false, style: styleName,
      banded: true, bandedCols: false, firstCol: false, lastCol: false, filter: { criteria: {}, hidden: {} }, totalsFns: {},
    };
    setTables((list) => [...list, t]);
    widenForFilterButtons({ r1, c1: rg.c1, r2, c2: rg.c2 });
  }, meta());
  const t = sheet().tables.find((x) => x.id === id);
  selectRange({ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 }, 'cells', { r: t.r1 + 1 <= t.r2 ? t.r1 + 1 : t.r1, c: t.c1 });
  ribbon.selectTab('tableDesign');
  toast(`'${name}' 표를 만들었습니다. 아래나 오른쪽에 이어서 입력하면 표가 자동으로 늘어납니다.`);
  return true;
}

/** 값을 넣은 범위(rg) 때문에 표가 늘어나야 하면 늘림 (transact 안에서) */
function afterDataEntry(rg) {
  const s = sheet();
  if (!s.tables?.length) return;
  // 실제로 값이 들어간 셀이 있을 때만
  let any = false;
  for (const [r, c] of cellsIn({ ...rg, r2: Math.min(rg.r2, rg.r1 + 5000), c2: Math.min(rg.c2, rg.c1 + 500) })) if (!isEmptyAt(r, c)) { any = true; break; }
  if (!any) return;
  for (const e of expansionFor(s, rg)) {
    const t = sheet().tables.find((x) => x.id === e.id);
    const nt = { ...t, ...e };
    if (e.c2 !== undefined && t.header) {
      const names = columnNames(wb, si, t).map((n) => n.toLowerCase());
      for (let c = t.c2 + 1; c <= e.c2; c++) {
        if (!isEmptyAt(t.r1, c)) continue;
        let n = c - t.c1 + 1;
        while (names.includes(`열${n}`)) n++;
        names.push(`열${n}`);
        wb.setInput(si, t.r1, c, `열${n}`);
      }
      if (t.totals) nt.r2 = t.r2;
    }
    setTables((list) => list.map((x) => (x.id === t.id ? nt : x)));
    if (e.r2 !== undefined) fillNewTableRows(nt, t.r2 + 1, e.r2);
  }
}

/** 새로 들어온 행: 계산 열 수식과 위 행의 표시 형식을 이어받음 */
function fillNewTableRows(t, from, to) {
  const top = dataTop(t);
  const prev = from - 1;
  if (prev < top) return;
  for (let c = t.c1; c <= t.c2; c++) {
    const above = wb.getCell(si, prev, c);
    const calc = calculatedFormula(t, c, prev);
    for (let r = from; r <= to; r++) {
      const cur = wb.getCell(si, r, c);
      if (calc && !cur?.raw) wb.setInput(si, r, c, shiftFormula(calc.raw, r - calc.r, 0));
      if (above?.style && !cur?.style) wb.setStyle(si, r, c, { ...above.style });
    }
  }
}

/** 열 전체가 같은 수식(상대 위치 기준)이면 그 수식 { raw, r } */
function calculatedFormula(t, c, last) {
  const top = dataTop(t);
  const first = wb.getCell(si, top, c);
  if (!first?.formula) return null;
  for (let r = top + 1; r <= Math.min(last, top + 2000); r++) {
    const cell = wb.getCell(si, r, c);
    if (!cell?.formula || cell.raw !== shiftFormula(first.raw, r - top, 0)) return null;
  }
  return { raw: first.raw, r: top };
}

/** 표 열의 빈 셀에 수식을 넣으면 그 열 전체에 채움 (엑셀의 계산 열) */
function maybeCalculatedColumn(r, c, text) {
  const t = tableHere(r, c);
  if (!t || !text.startsWith('=') || r < dataTop(t) || r > dataBottom(t)) return;
  for (let rr = dataTop(t); rr <= dataBottom(t); rr++) if (rr !== r && !isEmptyAt(rr, c)) return;
  for (let rr = dataTop(t); rr <= dataBottom(t); rr++) if (rr !== r) wb.setInput(si, rr, c, shiftFormula(text, rr - r, 0));
}

// ── 필터 대상 (시트 필터 또는 표) ──
/** key: '' = 시트 필터, 표 id = 표 필터 */
function getFilter(key) {
  if (key) {
    const t = (sheet().tables ?? []).find((x) => x.id === key);
    return t?.filter && t.header ? tableFilterRange(t) : null;
  }
  return sheet().filter;
}

function putFilter(key, nf) {
  if (key) setTables((list) => list.map((t) => (t.id === key ? { ...t, filter: nf ? { criteria: nf.criteria ?? {}, hidden: nf.hidden ?? {}, sort: nf.sort } : null } : t)));
  else wb.setSheetProp(si, 'filter', nf);
}

/** 현재 셀이 속한 필터 대상 */
function filterKeyHere() {
  const t = tableHere();
  if (t?.filter && t.header) return t.id;
  const f = sheet().filter;
  if (f && active.r >= f.r1 && active.r <= f.r2 && active.c >= f.c1 && active.c <= f.c2) return '';
  if (f) return '';
  return null;
}

function allFilters() {
  const s = sheet();
  return [...(s.filter ? [['', s.filter]] : []), ...(s.tables ?? []).filter((t) => t.filter && t.header).map((t) => [t.id, tableFilterRange(t)])];
}

// ── 요약 행 ──
function totalsFormula(t, c, fnId) {
  const f = TOTAL_FUNCS.find((x) => x.id === fnId);
  if (!f?.code) return '';
  const name = columnNames(wb, si, t)[c - t.c1];
  return `=SUBTOTAL(${f.code},${t.name}[${name.replace(/(['[\]#@])/g, "'$1")}])`;
}

function setTotals(t, on) {
  if (on === !!t.totals) return;
  wb.transact(() => {
    if (on) {
      let row = t.r2 + 1;
      let busy = false;
      for (let c = t.c1; c <= t.c2; c++) if (!isEmptyAt(row, c)) busy = true;
      if (busy || (sheet().tables ?? []).some((o) => o !== t && o.r1 <= row && o.r2 >= row && o.c1 <= t.c2 && o.c2 >= t.c1)) wb.insertRows(si, row, 1);
      const cur = sheet().tables.find((x) => x.id === t.id);
      row = cur.r2 + 1;
      const nt = { ...cur, r2: row, totals: true, totalsFns: {} };
      setTables((list) => list.map((x) => (x.id === t.id ? nt : x)));
      // 마지막 숫자 열은 합계, 첫 열은 '요약'
      let lastNum = null;
      for (let c = nt.c2; c >= nt.c1 && lastNum === null; c--) {
        for (let r = dataTop(nt); r <= Math.min(dataBottom(nt), dataTop(nt) + 50); r++) if (typeof valueAt(r, c) === 'number') { lastNum = c; break; }
      }
      const fns = {};
      if (lastNum !== null) { fns[lastNum] = 'sum'; wb.setInput(si, row, lastNum, totalsFormula(nt, lastNum, 'sum')); }
      if (lastNum !== nt.c1) wb.setInput(si, row, nt.c1, '요약');
      setTables((list) => list.map((x) => (x.id === t.id ? { ...x, totalsFns: fns } : x)));
    } else {
      for (let c = t.c1; c <= t.c2; c++) wb.setCellData(si, t.r2, c, null);
      setTables((list) => list.map((x) => (x.id === t.id ? { ...x, r2: t.r2 - 1, totals: false, totalsFns: {} } : x)));
    }
  }, meta());
}

function openTotalsMenu() {
  const t = tableHere();
  if (!t?.totals || active.r !== t.r2) return;
  const c = active.c;
  const b = gv.clientRect({ r1: active.r, c1: c, r2: active.r, c2: c });
  const cur = t.totalsFns?.[c] ?? 'none';
  openMenu({ x: b.left, y: b.bottom }, TOTAL_FUNCS.map((f) => ({
    label: f.label, checked: f.id === cur,
    action: () => {
      wb.transact(() => {
        wb.setInput(si, t.r2, c, totalsFormula(t, c, f.id));
        setTables((list) => list.map((x) => (x.id === t.id ? { ...x, totalsFns: { ...x.totalsFns, [c]: f.id } } : x)));
      }, meta());
      focusGrid();
    },
  })), { minWidth: Math.max(120, b.width + 17) });
}

// ── 테이블 디자인 명령 ──
function renameTable(name) {
  const t = tableHere();
  if (!t) return;
  const n = String(name ?? '').trim();
  if (n === t.name) return;
  if (!validTableName(n)) { alertDialog('표 이름', '표 이름은 글자나 밑줄(_)로 시작해야 하고 공백을 쓸 수 없으며, 셀 주소(A1 등)처럼 보이면 안 됩니다.'); updateRibbon(); return; }
  if (findTable(wb, n)) { alertDialog('표 이름', '이미 사용 중인 이름입니다.'); updateRibbon(); return; }
  const re = new RegExp(`(^|[^\\w.\\u00c0-\\uffff])${t.name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\[`, 'gi');
  wb.transact(() => {
    // 수식 안의 표 이름도 바꿈
    wb.sheets.forEach((s, i) => {
      for (const [k, cell] of s.cells) {
        if (!cell.formula || !re.test(cell.raw)) continue;
        re.lastIndex = 0;
        const [r, c] = k.split(',').map(Number);
        wb.setCellData(i, r, c, { ...cellData(cell), raw: cell.raw.replace(re, (m0, pre) => `${pre}${n}[`) });
      }
    });
    setTables((list) => list.map((x) => (x.id === t.id ? { ...x, name: n } : x)));
    wb.sheets.forEach((s2, i) => {
      if (!(s2.slicers ?? []).some((x) => x.source?.table === t.name)) return;
      wb.setSheetProp(i, 'slicers', s2.slicers.map((x) => (x.source?.table === t.name ? { ...x, source: { ...x.source, table: n } } : { ...x })));
    });
  }, meta());
}

function resizeTableDialog() {
  const t = tableHere();
  if (!t) return;
  formDialog('표 크기 조정', [{ name: 'ref', label: '새 데이터 범위', value: tableRangeText(t) }], ({ ref }) => {
    const p = parseRangeName(ref.replace(/[=$]/g, '').trim());
    if (!p) { toast('참조가 올바르지 않습니다.'); return false; }
    if (p.r1 !== t.r1 || p.c1 > t.c2 || p.c2 < t.c1) { alertDialog('표 크기 조정', '머리글은 같은 행에 있어야 하고, 새 범위는 원래 표와 겹쳐야 합니다.'); return false; }
    if ((sheet().tables ?? []).some((o) => o.id !== t.id && o.r1 <= p.r2 && o.r2 >= p.r1 && o.c1 <= p.c2 && o.c2 >= p.c1)) { alertDialog('표 크기 조정', '표는 다른 표와 겹칠 수 없습니다.'); return false; }
    wb.transact(() => {
      const nt = { ...t, ...p, r2: Math.max(p.r2, p.r1 + 1) };
      if (t.header) {
        const names = uniqueNames([...Array(nt.c2 - nt.c1 + 1)].map((_, i) => displayText(nt.r1, nt.c1 + i)));
        names.forEach((n, i) => { if (displayText(nt.r1, nt.c1 + i) !== n) wb.setInput(si, nt.r1, nt.c1 + i, `'${n}`); });
      }
      setTables((list) => list.map((x) => (x.id === t.id ? nt : x)));
    }, meta());
    return true;
  });
}

/** 표 → 일반 범위: 표 서식은 셀 서식으로, 구조적 참조는 셀 주소로 */
function convertTableToRange() {
  const t = tableHere();
  if (!t) return;
  openDialog({
    title: 'Tabula', body: '표를 정상 범위로 변환하시겠습니까?',
    buttons: [{
      label: '예', primary: true,
      action: () => {
        wb.transact(() => {
          for (let r = t.r1; r <= t.r2; r++) {
            for (let c = t.c1; c <= t.c2; c++) {
              const ts = tableCellStyle(t, r, c);
              if (!ts) continue;
              const own = wb.getCell(si, r, c)?.style ?? {};
              const patch = {};
              for (const [k, v] of Object.entries(ts)) if (own[k] === undefined) patch[k] = v;
              if (Object.keys(patch).length) wb.setStyle(si, r, c, patch);
            }
          }
          derefTableFormulas(t);
          setTables((list) => list.filter((x) => x.id !== t.id));
          wb.sheets.forEach((s2, i) => {
            if ((s2.slicers ?? []).some((x) => x.source?.table === t.name)) wb.setSheetProp(i, 'slicers', s2.slicers.filter((x) => x.source?.table !== t.name).map((x) => ({ ...x })));
          });
        }, meta());
        toast('표를 일반 범위로 바꿨습니다.');
      },
    }, { label: '아니요' }],
  });
}

/** 이 표를 가리키는 구조적 참조를 셀 주소로 바꿈 */
function derefTableFormulas(t) {
  const host = sheet().name;
  wb.sheets.forEach((s, i) => {
    for (const [k, cell] of s.cells) {
      if (!cell.formula || !cell.raw.includes('[')) continue;
      const [r, c] = k.split(',').map(Number);
      let toks;
      try { toks = tokenize(cell.raw.slice(1)); } catch { continue; }
      const body = cell.raw.slice(1);
      let out = '';
      let pos = 0;
      let changed = false;
      for (const tk of toks) {
        if (tk.t !== 'sref') continue;
        const inT = !tk.table && i === si && tableAt(s, r, c)?.id === t.id;
        if (!inT && (tk.table ?? '').toLowerCase() !== t.name.toLowerCase()) continue;
        const rg = resolveStructRef(wb, tk.table, tk.spec, { si: i, r, c });
        if (!rg) continue;
        const prefix = i === si ? '' : `${quoteSheetName(host)}!`;
        const ref = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? cellName(rg.r1, rg.c1) : `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}`;
        const abs = /@|this row|현재 행/i.test(tk.spec) ? ref : ref.replace(/([A-Z]+)(\d+)/g, '$$$1$$$2');
        out += body.slice(pos, tk.s) + prefix + abs;
        pos = tk.e;
        changed = true;
      }
      if (changed) wb.setCellData(i, r, c, { ...cellData(cell), raw: `=${out}${body.slice(pos)}` });
    }
  });
}

function toggleTableOption(key) {
  const t = tableHere();
  if (!t) return;
  if (key === 'totals') { setTotals(t, !t.totals); return; }
  if (key === 'filter') {
    if (!t.header) { toast('필터 단추를 쓰려면 머리글 행이 있어야 합니다.'); return; }
    updateTable(t.id, { filter: t.filter ? null : { criteria: {}, hidden: {} } });
    gv.layout();
    return;
  }
  if (key === 'header') {
    wb.transact(() => {
      if (t.header) {
        // 머리글 행을 표에서 빼고 이름은 기억
        const names = columnNames(wb, si, t);
        for (let c = t.c1; c <= t.c2; c++) wb.setCellData(si, t.r1, c, null);
        setTables((list) => list.map((x) => (x.id === t.id ? { ...x, r1: t.r1 + 1, header: false, columns: names, filter: null } : x)));
      } else {
        let r1 = t.r1 - 1;
        let busy = r1 < 0;
        if (!busy) for (let c = t.c1; c <= t.c2; c++) if (!isEmptyAt(r1, c)) busy = true;
        if (busy || (sheet().tables ?? []).some((o) => o !== t && o.r1 <= r1 && o.r2 >= r1 && o.c1 <= t.c2 && o.c2 >= t.c1)) {
          wb.insertRows(si, t.r1, 1);
          r1 = t.r1;
        }
        const cur = sheet().tables.find((x) => x.id === t.id);
        const names = uniqueNames(cur.columns ?? []);
        names.forEach((n, i) => wb.setInput(si, r1, cur.c1 + i, `'${n}`));
        setTables((list) => list.map((x) => (x.id === t.id ? { ...x, r1, header: true, columns: undefined, filter: { criteria: {}, hidden: {} } } : x)));
      }
    }, meta());
    gv.layout();
    return;
  }
  updateTable(t.id, { [key]: !t[key] });
}

function tableStyleGallery(anchorEl, forCreate = false) {
  const chip = (st) => el('button', {
    class: 'style-chip tstyle', title: st.label, onmousedown: (e) => e.preventDefault(),
    style: { background: `linear-gradient(${st.swatch[0]} 0 28%, ${st.swatch[1]} 28% 52%, ${st.swatch[2] === '#ffffff' ? '#ffffff' : st.swatch[2]} 52% 76%, ${st.swatch[1]} 76%)` },
    onclick: () => {
      closeMenus();
      const t = tableHere();
      if (t && !forCreate) updateTable(t.id, { style: st.name });
      else createTableDialog(st.name);
      focusGrid();
    },
  });
  const groups = ['밝게', '보통', '어둡게'].flatMap((g) => [{ title: g }, { node: el('div', { class: 'style-grid tstyles' }, TABLE_STYLES.filter((s) => s.group === g).map(chip)) }]);
  const t = tableHere();
  openMenu(anchorEl, [
    ...groups,
    ...(t ? [{ sep: true }, { label: '지우기', action: () => updateTable(t.id, { style: 'None' }) }] : []),
  ]);
}

// ───────────────────────── 슬라이서 ─────────────────────────
const SLICER_COLORS = [['blue', '파랑'], ['orange', '주황'], ['gray', '회색'], ['gold', '금색'], ['sky', '하늘색'], ['green', '녹색']];

const sortItems = (entries) => entries.sort((a, b) => {
  if (a.key === '') return 1;
  if (b.key === '') return -1;
  const x = a.v;
  const y = b.v;
  if (typeof x === 'number' && typeof y === 'number') return x - y;
  if (typeof x === 'number') return -1;
  if (typeof y === 'number') return 1;
  return String(a.key).localeCompare(String(b.key), 'ko');
});

/** 슬라이서가 가리키는 대상과 항목 → { caption, items: [{ key, text, selected, hasData }], filtered, broken?, apply(values) } */
function slicerModel(sl) {
  const src = sl.source ?? {};
  if (src.kind === 'table') {
    const f = findTable(wb, src.table);
    if (!f) return { items: [], broken: '연결된 표를 찾을 수 없습니다.' };
    const { t } = f;
    const s = f.si;
    const ci = columnNames(wb, s, t).findIndex((n) => n.toLowerCase() === String(src.column).toLowerCase());
    if (ci < 0) return { items: [], broken: `'${src.column}' 열을 찾을 수 없습니다.` };
    const c = t.c1 + ci;
    const crit = t.filter?.criteria ?? {};
    const sel = Array.isArray(crit[c]) ? new Set(crit[c]) : null;
    const others = Object.entries(crit).filter(([k, v]) => Number(k) !== c && Array.isArray(v)).map(([k, v]) => [Number(k), new Set(v)]);
    const vals = new Map();
    for (let r = dataTop(t); r <= dataBottom(t); r++) {
      const key = displayText(r, c, s);
      const pass = others.every(([k, set]) => set.has(displayText(r, k, s)));
      const e = vals.get(key) ?? { key, v: wb.getValue(s, r, c), hasData: false };
      e.hasData ||= pass;
      vals.set(key, e);
    }
    const items = sortItems([...vals.values()]).map((e) => ({ key: e.key, text: e.key === '' ? '(비어 있음)' : e.key, selected: !sel || sel.has(e.key), hasData: e.hasData }));
    return {
      items, filtered: !!sel,
      apply: (values) => {
        const prev = si;
        if (f.si !== si) si = f.si;
        applyFilterCriteria(c, values, t.id, { quiet: true });
        si = prev;
      },
    };
  }
  if (src.kind === 'pivot') {
    const pi = src.self ? si : wb.sheetIndexByName(src.sheet);
    const def = pi >= 0 ? wb.sheets[pi].pivot : null;
    if (!def) return { items: [], broken: '연결된 피벗 테이블을 찾을 수 없습니다.' };
    const rows = pivotSource(def);
    if (!rows) return { items: [], broken: '피벗 테이블 원본을 찾을 수 없습니다.' };
    const header = rows[0].map((h, i) => (h === null || h === '' ? `열${i + 1}` : String(h)));
    const fi = header.findIndex((h) => h.toLowerCase() === String(src.field).toLowerCase());
    if (fi < 0) return { items: [], broken: `'${src.field}' 필드를 찾을 수 없습니다.` };
    const filters = def.filters ?? {};
    const own = Object.keys(filters).find((k) => k.toLowerCase() === header[fi].toLowerCase());
    const sel = own ? new Set(filters[own]) : null;
    const others = Object.entries(filters).filter(([k]) => k !== own).map(([k, v]) => [header.findIndex((h) => h.toLowerCase() === k.toLowerCase()), new Set(v)]).filter(([i]) => i >= 0);
    const vals = new Map();
    for (const row of rows.slice(1)) {
      if (row.every((v) => v === null || v === '')) continue;
      const key = pivotItemText(row[fi]);
      const pass = others.every(([i, set]) => set.has(pivotItemText(row[i])));
      const e = vals.get(key) ?? { key, v: row[fi], hasData: false };
      e.hasData ||= pass;
      vals.set(key, e);
    }
    const items = sortItems([...vals.values()]).map((e) => ({ key: e.key, text: e.key, selected: !sel || sel.has(e.key), hasData: e.hasData }));
    return {
      items, filtered: !!sel,
      apply: (values) => {
        const nf = { ...filters };
        if (own) delete nf[own];
        if (values) nf[header[fi]] = values;
        const def2 = { ...def, filters: nf };
        wb.transact(() => { writePivot(pi, def2); wb.setSheetProp(pi, 'pivot', def2); }, meta());
      },
    };
  }
  return { items: [], broken: '연결 대상이 없습니다.' };
}

function slicerPick(id, key, additive) {
  const sl = (sheet().slicers ?? []).find((x) => x.id === id);
  if (!sl) return;
  const m = slicerModel(sl);
  if (m.broken) return;
  const all = m.items.map((i) => i.key);
  let cur;
  if (additive || sl.multi) {
    cur = new Set(m.items.filter((i) => i.selected).map((i) => i.key));
    if (cur.has(key)) cur.delete(key); else cur.add(key);
    if (!cur.size) cur = new Set(all);
  } else {
    cur = new Set([key]);
    if (m.filtered && m.items.filter((i) => i.selected).length === 1 && m.items.find((i) => i.key === key)?.selected) cur = new Set(all);
  }
  m.apply(cur.size === all.length ? null : all.filter((k) => cur.has(k)));
  gv.layout();
  setMode();
}

function slicerClear(id) {
  const sl = (sheet().slicers ?? []).find((x) => x.id === id);
  if (!sl) return;
  const m = slicerModel(sl);
  if (m.broken || !m.filtered) return;
  m.apply(null);
  gv.layout();
  setMode();
}

/** 표 또는 피벗 테이블에 슬라이서 넣기 */
function insertSlicerDialog() {
  if (editing && !commitEdit()) return;
  const t = tableHere();
  const pivot = sheet().pivot;
  let fields;
  let makeSource;
  let anchor;
  if (t) {
    fields = columnNames(wb, si, t);
    makeSource = (name) => ({ kind: 'table', table: t.name, column: name });
    anchor = gv.sheetRect({ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
  } else if (pivot) {
    const rows = pivotSource(pivot);
    if (!rows) { alertDialog('슬라이서 삽입', '피벗 테이블 원본을 찾을 수 없습니다.'); return; }
    fields = rows[0].map((h, i) => (h === null || h === '' ? `열${i + 1}` : String(h)));
    makeSource = (name) => ({ kind: 'pivot', self: true, field: name });
    const u = wb.usedRange(si);
    anchor = gv.sheetRect({ r1: 0, c1: 0, r2: Math.max(0, u.rows - 1), c2: Math.max(0, u.cols - 1) });
  } else {
    alertDialog('슬라이서 삽입', '슬라이서는 표나 피벗 테이블에 넣을 수 있습니다. 데이터 안에서 Ctrl+T 로 표를 만들거나 피벗 테이블 시트에서 다시 시도하세요.');
    return;
  }
  const checks = fields.map((n) => [n, el('input', { type: 'checkbox' })]);
  openDialog({
    title: '슬라이서 삽입', width: 320,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } },
      el('div', { class: 'muted' }, t ? `'${t.name}' 표에서 필터할 열을 고르세요.` : '피벗 테이블에서 필터할 필드를 고르세요.'),
      el('div', { class: 'fc-list', style: { maxHeight: '260px', padding: '4px 8px', gap: '4px' } }, checks.map(([n, cb]) => el('label', { class: 'fc-check' }, cb, n)))),
    buttons: [
      {
        label: '확인', primary: true,
        action: () => {
          const chosen = checks.filter(([, cb]) => cb.checked).map(([n]) => n);
          if (!chosen.length) return undefined;
          const list = chosen.map((name, i) => {
            const n = Math.min(12, slicerModel({ source: makeSource(name) }).items.length || 1);
            return {
              id: newObjId('sl'), caption: name, source: makeSource(name), columns: 1, color: 'blue', multi: false,
              x: Math.round(anchor.x + anchor.w + 24 + i * 192), y: Math.round(anchor.y), w: 180, h: Math.min(300, 40 + n * 28),
            };
          });
          wb.transact(() => {
            let z = nextZ();
            wb.setSheetProp(si, 'slicers', [...(sheet().slicers ?? []).map((x) => ({ ...x })), ...list.map((x) => ({ ...x, z: z++ }))]);
            if (t && !t.filter) setTables((l) => l.map((x) => (x.id === t.id ? { ...x, filter: { criteria: {}, hidden: {} } } : x)));
          }, meta());
          chartSel = list[list.length - 1].id;
          gv.ensureVisible(gv.rows.indexAt(list[0].y + 40), gv.cols.indexAt(list[0].x + 170));
          gv.renderObjectsAll();
          updateSelectionUI();
          return undefined;
        },
      },
      { label: '취소' },
    ],
  });
}

function slicerSettings(id) {
  const sl = (sheet().slicers ?? []).find((x) => x.id === id);
  if (!sl) return;
  formDialog('슬라이서 설정', [
    { name: 'caption', label: '캡션', value: sl.caption ?? '' },
    { name: 'columns', label: '열 수', type: 'number', value: sl.columns ?? 1 },
    { name: 'color', label: '스타일', type: 'select', value: sl.color ?? 'blue', options: SLICER_COLORS.map(([v, l]) => ({ value: v, label: l })) },
  ], (v) => {
    updateObject(id, { caption: v.caption, columns: clamp(Number(v.columns) || 1, 1, 20), color: v.color });
  }, { note: sl.source?.kind === 'table' ? `원본: ${sl.source.table}[${sl.source.column}]` : `원본: 피벗 테이블의 '${sl.source?.field}' 필드` });
}

// ───────────────────────── 텍스트 나누기 (Alt+A+E) ─────────────────────────
function textToColumns() {
  if (editing && !commitEdit()) return;
  let rg = selKind === 'cells' ? sel : usedClip(sel);
  if (rg.c1 !== rg.c2) { alertDialog('텍스트 나누기', '한 번에 한 열만 변환할 수 있습니다. 한 열의 셀 범위를 선택하세요.'); return; }
  if (isSingle(rg)) {
    // 셀 하나만 선택했으면 아래로 이어진 데이터까지
    let r2 = rg.r1;
    const last = wb.usedRange(si).rows - 1;
    while (r2 + 1 <= last && !isEmptyAt(r2 + 1, rg.c1)) r2++;
    rg = { ...rg, r2 };
  }
  rg = { ...rg, r2: Math.max(rg.r1, Math.min(rg.r2, wb.usedRange(si).rows - 1)) };
  const lines = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const cell = wb.getCell(si, r, rg.c1);
    lines.push(cell?.raw && !cell.formula ? cell.raw.replace(/^'/, '') : displayText(r, rg.c1));
  }
  if (!lines.some((l) => l)) { alertDialog('텍스트 나누기', '나눌 텍스트가 있는 셀을 선택하세요.'); return; }

  // 상태
  const firstLines = lines.filter((l) => l).slice(0, 50);
  const guess = { tab: firstLines.some((l) => l.includes('\t')), comma: firstLines.some((l) => l.includes(',')), semicolon: false, space: false, other: '' };
  if (!guess.tab && !guess.comma) guess.space = firstLines.some((l) => l.trim().includes(' '));
  const o = { mode: 'delimited', ...guess, consecutive: guess.space, qualifier: '"', breaks: suggestBreaks(firstLines) };
  let colFmts = [];
  let selCol = 0;
  let step = 1;
  const dest = { text: `$${colToName(rg.c1)}$${rg.r1 + 1}` };

  const split = (line) => (o.mode === 'delimited' ? splitDelimited(line, o) : splitFixed(line, o.breaks));
  const previewRows = () => lines.slice(0, 12).map(split);
  const colCount = () => Math.max(1, ...lines.slice(0, 2000).map((l) => split(l).length));

  const body = el('div', { class: 'ttc' });
  const stepTitle = el('div', { class: 'ttc-step' });
  const content = el('div', { class: 'ttc-content' });
  body.append(stepTitle, content);

  const previewTable = (rows, { headers = null, onPick = null } = {}) => {
    const n = Math.max(1, ...rows.map((r) => r.length));
    const t = el('table', { class: 'ttc-table' });
    if (headers) {
      t.append(el('tr', {}, [...Array(n)].map((_, i) => {
        const th = el('th', { class: `${i === selCol ? 'on' : ''}${colFmts[i]?.fmt === 'skip' ? ' skip' : ''}` }, headers(i));
        if (onPick) th.addEventListener('click', () => onPick(i));
        return th;
      })));
    }
    for (const r of rows) {
      t.append(el('tr', {}, [...Array(n)].map((_, i) => {
        const td = el('td', { class: `${headers && i === selCol ? 'on' : ''}${colFmts[i]?.fmt === 'skip' ? ' skip' : ''}` }, r[i] ?? '');
        if (onPick) td.addEventListener('click', () => onPick(i));
        return td;
      })));
    }
    return el('div', { class: 'ttc-preview' }, t);
  };

  const radio = (name, value, cur, label, onChange) => {
    const i = el('input', { type: 'radio', name, value, checked: cur === value });
    i.addEventListener('change', () => { if (i.checked) onChange(value); });
    return el('label', { class: 'fc-check' }, i, label);
  };
  const check = (key, label) => {
    const i = el('input', { type: 'checkbox', checked: !!o[key] });
    i.addEventListener('change', () => { o[key] = i.checked; render(); });
    return el('label', { class: 'fc-check' }, i, label);
  };

  const fixedRuler = () => {
    // 글자를 클릭해 나눌 위치 추가, 선을 다시 클릭하면 제거
    const sample = lines.slice(0, 10);
    const width = Math.max(10, ...sample.map((l) => l.length)) + 2;
    const box = el('div', { class: 'ttc-ruler' });
    const scale = el('div', { class: 'ttc-line scale' });
    for (let i = 0; i < width; i++) scale.append(el('span', { class: `ch${o.breaks.includes(i) ? ' brk' : ''}`, 'data-i': i }, i % 10 === 0 ? String(i / 10 % 10) : i % 5 === 0 ? '·' : ''));
    box.append(scale);
    for (const l of sample) {
      const line = el('div', { class: 'ttc-line' });
      for (let i = 0; i < width; i++) line.append(el('span', { class: `ch${o.breaks.includes(i) ? ' brk' : ''}`, 'data-i': i }, l[i] ?? ''));
      box.append(line);
    }
    box.addEventListener('click', (e) => {
      const ch = e.target.closest('.ch');
      if (!ch) return;
      const i = Number(ch.dataset.i);
      if (i <= 0) return;
      o.breaks = o.breaks.includes(i) ? o.breaks.filter((b) => b !== i) : [...o.breaks, i].sort((a, b) => a - b);
      render();
    });
    return el('div', { class: 'ttc-rulerwrap' }, box);
  };

  const render = () => {
    content.replaceChildren();
    if (step === 1) {
      stepTitle.textContent = '1/3단계 — 원본 데이터 형식';
      content.append(
        el('div', { class: 'muted' }, '데이터를 나눌 방법을 선택하세요.'),
        radio('ttcmode', 'delimited', o.mode, '구분 기호로 분리됨 — 각 필드가 쉼표나 탭, 공백 같은 문자로 나뉘어 있음', (v) => { o.mode = v; render(); }),
        radio('ttcmode', 'fixed', o.mode, '너비가 일정함 — 각 필드가 같은 너비로 정렬되어 있음', (v) => { o.mode = v; render(); }),
        el('div', { class: 'fc-title' }, `선택한 데이터 미리 보기 (${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)})`),
        el('pre', { class: 'ttc-raw' }, lines.slice(0, 10).map((l, i) => `${rg.r1 + i + 1}  ${l}`).join('\n')),
      );
    } else if (step === 2) {
      if (o.mode === 'delimited') {
        stepTitle.textContent = '2/3단계 — 구분 기호';
        const other = el('input', { type: 'text', value: o.other, maxlength: 1, style: { width: '32px' } });
        other.addEventListener('input', () => { o.other = other.value; render(); setTimeout(() => { const x = content.querySelector('input[maxlength="1"]'); x?.focus(); x?.setSelectionRange(1, 1); }); });
        const qual = el('select', {}, [['"', '"'], ["'", "'"], ['', '{없음}']].map(([v, l]) => el('option', { value: v, selected: o.qualifier === v }, l)));
        qual.addEventListener('change', () => { o.qualifier = qual.value; render(); });
        content.append(
          el('div', { class: 'ttc-opts' },
            el('div', { class: 'fc-col' }, el('div', { class: 'fc-title' }, '구분 기호'),
              check('tab', '탭'), check('semicolon', '세미콜론(;)'), check('comma', '쉼표(,)'), check('space', '공백(띄어쓰기)'),
              el('label', { class: 'fc-check' }, el('span', {}, '기타:'), other)),
            el('div', { class: 'fc-col' }, el('div', { class: 'fc-title' }, '옵션'), check('consecutive', '연속된 구분 기호를 하나로 처리'),
              el('label', { class: 'fc-field' }, el('span', {}, '텍스트 한정자'), qual))),
          el('div', { class: 'fc-title' }, '데이터 미리 보기'),
          previewTable(previewRows()),
        );
      } else {
        stepTitle.textContent = '2/3단계 — 열 구분선';
        content.append(
          el('div', { class: 'muted' }, '나눌 위치를 클릭해 구분선을 넣고, 구분선을 다시 클릭하면 지웁니다.'),
          fixedRuler(),
          el('div', { class: 'fc-title' }, '데이터 미리 보기'),
          previewTable(previewRows()),
        );
      }
    } else {
      stepTitle.textContent = '3/3단계 — 열 데이터 서식';
      const n = colCount();
      colFmts = [...Array(n)].map((_, i) => colFmts[i] ?? { fmt: 'general', order: 'YMD' });
      selCol = Math.min(selCol, n - 1);
      const cf = colFmts[selCol];
      const order = el('select', { disabled: cf.fmt !== 'date' }, DATE_ORDERS.map((d) => el('option', { value: d, selected: cf.order === d }, d)));
      order.addEventListener('change', () => { cf.order = order.value; render(); });
      const setFmt = (v) => { cf.fmt = v; render(); };
      const destIn = el('input', { type: 'text', value: dest.text, style: { width: '120px' } });
      destIn.addEventListener('input', () => { dest.text = destIn.value; });
      const label = (i) => ({ general: '일반', text: '텍스트', date: `날짜(${colFmts[i].order})`, skip: '건너뜀' }[colFmts[i].fmt]);
      const shown = previewRows().map((r) => r.map((v, i) => {
        const f = colFmts[i];
        if (!f || f.fmt !== 'date') return v;
        const d = parseDateOrder(v, f.order);
        return d ?? v;
      }));
      content.append(
        el('div', { class: 'ttc-opts' },
          el('div', { class: 'fc-col' }, el('div', { class: 'fc-title' }, `열 데이터 서식 — ${selCol + 1}번째 열`),
            radio('ttcfmt', 'general', cf.fmt, '일반 (숫자는 숫자로, 날짜는 날짜로)', setFmt),
            radio('ttcfmt', 'text', cf.fmt, '텍스트 (앞의 0 유지)', setFmt),
            el('div', { class: 'fc-row' }, radio('ttcfmt', 'date', cf.fmt, '날짜:', setFmt), order),
            radio('ttcfmt', 'skip', cf.fmt, '열 가져오지 않음(건너뜀)', setFmt)),
          el('div', { class: 'fc-col' }, el('div', { class: 'fc-title' }, '대상'), destIn,
            el('div', { class: 'muted fc-note' }, '날짜 순서 예: YMD = 20240315 · 2024.3.15, MDY = 03/15/2024, DMY = 15-03-2024. 날짜로 바뀐 값은 날짜 서식으로 표시됩니다.'))),
        el('div', { class: 'fc-title' }, '데이터 미리 보기 (열 머리글을 눌러 선택)'),
        previewTable(shown, { headers: label, onPick: (i) => { selCol = i; render(); } }),
      );
    }
    btnBack.disabled = step === 1;
    btnNext.disabled = step === 3;
  };

  const finish = () => {
    const d = parseRangeName(dest.text.replace(/\$/g, '').trim());
    if (!d) { alertDialog('텍스트 나누기', '대상 셀 참조가 올바르지 않습니다.'); return false; }
    const n = colCount();
    colFmts = [...Array(n)].map((_, i) => colFmts[i] ?? { fmt: 'general', order: 'YMD' });
    const outCols = colFmts.map((f, i) => (f.fmt === 'skip' ? -1 : i)).filter((i) => i >= 0);
    const r0 = d.r1;
    const c0 = d.c1;
    const rows = lines.map(split);
    // 원본 자리 밖에 이미 데이터가 있으면 확인
    let occupied = false;
    for (let i = 0; i < rows.length && !occupied; i++) {
      for (let k = 0; k < outCols.length; k++) {
        const r = r0 + i;
        const c = c0 + k;
        if (c === rg.c1 && r >= rg.r1 && r <= rg.r2) continue;
        if (!isEmptyAt(r, c)) { occupied = true; break; }
      }
    }
    const apply = () => {
      wb.transact(() => {
        rows.forEach((parts, i) => {
          const r = r0 + i;
          if (c0 !== rg.c1 || r0 !== rg.r1) { /* 원본은 그대로 */ } else if (rg.r1 + i <= rg.r2) wb.setInput(si, r, rg.c1, '');
          outCols.forEach((ci, k) => {
            const f = colFmts[ci];
            const v = convertPart(parts[ci] ?? '', f.fmt, f.order);
            if (v === null) return;
            if (f.fmt === 'text') {
              // 텍스트 서식 셀에는 입력한 그대로 들어감 (앞의 0 유지)
              wb.setStyle(si, r, c0 + k, { numFmt: 'text', code: undefined });
              wb.setInput(si, r, c0 + k, v.replace(/^'/, ''));
              return;
            }
            wb.setInput(si, r, c0 + k, v);
            if (f.fmt === 'date' && typeof wb.getValue(si, r, c0 + k) === 'number') {
              const cur = wb.getCell(si, r, c0 + k)?.style;
              if (!cur?.numFmt || cur.numFmt === 'general') wb.setStyle(si, r, c0 + k, { numFmt: v.includes(':') ? 'datetime' : 'date' });
            }
          });
        });
        autoWiden({ r1: r0, c1: c0, r2: r0 + Math.min(rows.length, 500) - 1, c2: c0 + outCols.length - 1 });
      }, meta());
      selectRange({ r1: r0, c1: c0, r2: r0 + rows.length - 1, c2: c0 + Math.max(0, outCols.length - 1) }, 'cells', { r: r0, c: c0 });
      toast(`${rows.length}행을 ${outCols.length}개 열로 나눴습니다.`);
    };
    if (occupied) {
      openDialog({
        title: 'Tabula', body: '여기에 이미 데이터가 있습니다. 바꾸시겠습니까?',
        buttons: [{ label: '확인', primary: true, action: apply }, { label: '취소' }],
      });
    } else apply();
    return undefined;
  };

  const btnBack = el('button', { class: 'btn', type: 'button' }, '< 뒤로');
  const btnNext = el('button', { class: 'btn', type: 'button' }, '다음 >');
  btnBack.addEventListener('click', () => { if (step > 1) { step--; render(); } });
  btnNext.addEventListener('click', () => {
    if (step === 2 && o.mode === 'fixed' && !o.breaks.length) { toast('구분선을 하나 이상 넣으세요.'); return; }
    if (step < 3) { step++; render(); }
  });
  const nav = el('div', { class: 'ttc-nav' }, btnBack, btnNext);
  body.append(nav);
  render();
  openDialog({
    title: '텍스트 마법사', width: 640, body,
    buttons: [{ label: '마침', primary: true, action: finish }, { label: '취소' }],
  });
}

// ───────────────────────── 바로 가기 키 순서 (Alt+A+E 등) ─────────────────────────
const KEYTIPS = {
  ae: ['textToColumns', '텍스트 나누기'], at: ['toggleFilter', '필터'], am: ['dedupe', '중복된 항목 제거'],
  avv: ['dataValidation', '데이터 유효성 검사'], ass: ['sortDialog', '정렬'], asa: ['sortAsc', '오름차순 정렬'], asd: ['sortDesc', '내림차순 정렬'],
  aa: ['refreshAll', '모두 새로 고침'], ac: ['clearFilter', '필터 지우기'], ay: ['reapplyFilter', '다시 적용'],
  nt: ['createTable', '표'], nv: ['insertPivot', '피벗 테이블'], nsf: ['insertSlicer', '슬라이서'], np: ['insertPicture', '그림'],
  nsh: ['shapesMenu', '도형'], nx: ['insertTextbox', '텍스트 상자'], nc: ['chartColumn', '세로 막대형 차트'],
  hoe: ['formatCells', '셀 서식'], hoi: ['autofitSel', '열 너비 자동 맞춤'], hoa: ['autofitRowsSel', '행 높이 자동 맞춤'],
  hmc: ['mergeCenter', '병합하고 가운데 맞춤'], hw: ['wrap', '텍스트 줄 바꿈'], hfp: ['painter', '서식 복사'], hb: ['borderLast', '테두리'],
  hk: ['fmtComma', '쉼표 스타일'], hp: ['fmtPercent', '백분율'], h0: ['incDecimal', '자릿수 늘림'], h9: ['decDecimal', '자릿수 줄임'],
  hlr: ['condManager', '조건부 서식 규칙 관리'], hln: ['condNewRule', '새 서식 규칙'], hlm: ['condMenuKey', '조건부 서식 메뉴'], ht: ['tableStyleKey', '표 서식'],
  wff: ['freezePanes', '틀 고정'], wfr: ['freezeTop', '첫 행 고정'], wfc: ['freezeFirstCol', '첫 열 고정'], wg: ['toggleGrid', '눈금선'],
  mf: ['insertFunction', '함수 삽입'], mua: ['autosum', '자동 합계'], f: ['backstage', '파일'],
};
const KEYTIP_TABS = { h: 'home', n: 'insert', p: 'layout', m: 'formulas', a: 'data', r: 'review', w: 'view', j: 'tableDesign' };
let keytip = null; // { seq, held }

const TAB_LABELS = Object.fromEntries(TABS.map((t) => [t.id, t.label]));

/** 리본 메뉴를 현재 셀 아래에 열기 (바로 가기 키용) */
function menuAtCell(name) {
  const b = gv.clientRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c });
  const anchor = el('div', { style: { position: 'fixed', left: `${b.left}px`, top: `${b.top}px`, width: `${b.width}px`, height: `${b.height}px` } });
  document.body.append(anchor);
  openNamedMenu(name, anchor);
  anchor.remove();
}

function keytipHint(seq) {
  const next = Object.entries(KEYTIPS).filter(([k]) => k.startsWith(seq) && k !== seq)
    .map(([k, [, label]]) => `${k.slice(seq.length).toUpperCase().split('').join('+')} ${label}`);
  const tabs = seq === '' ? Object.entries(KEYTIP_TABS).map(([k, id]) => `${k.toUpperCase()} ${TAB_LABELS[id] ?? id}`) : [];
  return `Alt${seq ? `+${seq.toUpperCase().split('').join('+')}` : ''}: ${[...tabs, ...next].slice(0, 12).join(' · ')}`;
}

function endKeytip() {
  if (!keytip) return;
  keytip = null;
  document.body.classList.remove('keytips');
  setMode();
}

/** 반환 true 면 키를 처리함 */
function handleKeytipKey(e) {
  // 슬라이서를 선택한 상태의 Alt+C(필터 지우기) · Alt+S(다중 선택)
  if (!keytip && e.altKey && (e.code === 'KeyC' || e.code === 'KeyS') && chartSel && sheet().slicers?.some((x) => x.id === chartSel)) return false;
  if (e.key === 'Alt') {
    if (!keytip) { keytip = { seq: '', held: true, clean: true }; }
    e.preventDefault();
    return true;
  }
  if (!keytip && !(e.altKey && /^Key[A-Z]$|^Digit\d$/.test(e.code) && !e.ctrlKey && !e.metaKey)) return false;
  if (!keytip) keytip = { seq: '', held: true, clean: false };
  if (e.key === 'Escape') { e.preventDefault(); endKeytip(); return true; }
  const m = /^Key([A-Z])$|^Digit(\d)$/.exec(e.code);
  if (!m) { endKeytip(); return false; }
  e.preventDefault();
  keytip.clean = false;
  const seq = keytip.seq + (m[1] ?? m[2]).toLowerCase();
  if (seq.length === 1 && KEYTIP_TABS[seq]) ribbon.selectTab(KEYTIP_TABS[seq]);
  if (KEYTIPS[seq]) {
    endKeytip();
    run(KEYTIPS[seq][0]);
    return true;
  }
  if (Object.keys(KEYTIPS).some((k) => k.startsWith(seq))) {
    keytip.seq = seq;
    document.body.classList.add('keytips');
    dom.status.textContent = keytipHint(seq);
    return true;
  }
  endKeytip();
  return true;
}

function handleKeytipUp(e) {
  if (e.key !== 'Alt' || !keytip) return;
  e.preventDefault();
  keytip.held = false;
  if (keytip.clean) {
    // Alt 만 눌렀다 뗌 → 바로 가기 키 모드
    keytip.clean = false;
    document.body.classList.add('keytips');
    dom.status.textContent = keytipHint('');
  } else if (!keytip.seq) endKeytip();
}

// ───────────────────────── 그림 개체 (차트 · 그림 · 도형) ─────────────────────────
const newObjId = (p) => `${p}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;

function allObjects() { return OBJECT_PROPS.flatMap((p) => sheet()[p] ?? []); }
function nextZ() { return allObjects().reduce((z, o) => Math.max(z, o.z ?? 0), 0) + 1; }
function minZ() { return allObjects().reduce((z, o) => Math.min(z, o.z ?? 0), 0); }

function setObjects(prop, fn) {
  wb.transact(() => wb.setSheetProp(si, prop, fn((sheet()[prop] ?? []).map((o) => ({ ...o })))), meta());
}

function updateObject(id, patch) {
  const f = findObject(sheet(), id);
  if (f) setObjects(f.prop, (list) => list.map((o) => (o.id === id ? { ...o, ...patch } : o)));
}

function addObject(prop, obj) {
  obj.z ??= nextZ();
  setObjects(prop, (list) => [...list, obj]);
  chartSel = obj.id;
  gv.renderObjectsAll();
  updateSelectionUI();
}

function deleteObject(id) {
  const f = findObject(sheet(), id);
  if (f) setObjects(f.prop, (list) => list.filter((o) => o.id !== id));
  chartSel = null;
  updateSelectionUI();
}

function nudgeObject(id, dx, dy) {
  const f = findObject(sheet(), id);
  if (f) updateObject(id, { x: Math.max(0, f.obj.x + dx), y: Math.max(0, f.obj.y + dy) });
}

function copyObject(id, cut) {
  const f = findObject(sheet(), id);
  if (!f) return;
  objClip = { prop: f.prop, obj: structuredClone(f.obj), n: 0, text: '' };
  navigator.clipboard?.writeText('').catch(() => {});
  if (cut) { objClip.n = -1; deleteObject(id); }
}

function pasteObject() {
  if (!objClip) return;
  objClip.n++;
  const d = objClip.n * 12;
  const o = { ...structuredClone(objClip.obj), id: newObjId({ charts: 'ch', images: 'im', slicers: 'sl' }[objClip.prop] ?? 'sh'), z: undefined };
  o.x += d;
  o.y += d;
  addObject(objClip.prop, o);
}

function editObject(id) {
  const f = findObject(sheet(), id);
  if (!f) return;
  if (f.prop === 'charts') chartDialog(id);
  else if (f.prop === 'shapes') shapeDialog(id);
  else if (f.prop === 'slicers') slicerSettings(id);
  else imageDialog(id);
}

/** 새 개체를 둘 위치: 현재 셀 왼쪽 위 */
function objectOrigin() {
  const r = gv.sheetRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c });
  return { x: Math.round(r.x), y: Math.round(r.y) };
}

function objectMenu(id, pos) {
  const f = findObject(sheet(), id);
  if (!f) return;
  if (chartSel !== id) { chartSel = id; gv.renderObjectsAll(); updateSelectionUI(); }
  const items = [
    { label: '잘라내기', icon: 'cut', key: 'Ctrl+X', action: () => copyObject(id, true) },
    { label: '복사', icon: 'copy', key: 'Ctrl+C', action: () => copyObject(id, false) },
    { label: '붙여넣기', icon: 'paste', key: 'Ctrl+V', disabled: !objClip, action: pasteObject },
    { sep: true },
  ];
  if (f.prop === 'charts') {
    items.push(
      { label: '차트 편집...', icon: 'chartColumn', action: () => chartDialog(id) },
      { title: '차트 종류 변경' },
      ...CHART_TYPES.map((t) => ({ label: t.label, checked: f.obj.type === t.id, action: () => updateChart(id, { type: t.id }) })),
    );
  } else if (f.prop === 'shapes') {
    items.push({ label: f.obj.kind === 'line' ? '선 서식...' : '텍스트 편집 및 도형 서식...', icon: 'shapes', action: () => shapeDialog(id) });
    if (f.obj.kind !== 'line') {
      items.push({ title: '도형 모양 변경' }, ...SHAPE_KINDS.filter((k) => k.id !== 'line').map((k) => ({
        label: k.label, checked: f.obj.kind === k.id, action: () => updateObject(id, { kind: k.id }),
      })));
    }
  } else if (f.prop === 'slicers') {
    items.push(
      { label: '슬라이서 설정...', icon: 'slicer', action: () => slicerSettings(id) },
      { label: `"${f.obj.caption}"에서 필터 지우기`, icon: 'filterClear', action: () => slicerClear(id) },
      { label: '다중 선택', checked: !!f.obj.multi, action: () => updateObject(id, { multi: !f.obj.multi }) },
    );
  } else {
    items.push(
      { label: '크기 및 속성...', icon: 'picture', action: () => imageDialog(id) },
      { label: '원래 크기로', action: () => resetImageSize(id) },
    );
  }
  items.push(
    { sep: true },
    { label: '맨 앞으로 가져오기', action: () => updateObject(id, { z: nextZ() }) },
    { label: '맨 뒤로 보내기', action: () => updateObject(id, { z: minZ() - 1 }) },
    { sep: true },
    { label: '삭제', icon: 'delete', key: 'Delete', action: () => deleteObject(id) },
  );
  openMenu(pos, items);
}

// 그림
function insertPicture() {
  const input = el('input', { type: 'file', accept: 'image/*' });
  input.addEventListener('change', () => { if (input.files[0]) addImageFile(input.files[0]); });
  input.click();
}

function addImageFile(file, at = null) {
  if (file.size > 10 * 1024 * 1024) { alertDialog('그림 삽입', '10MB 이하의 그림만 넣을 수 있습니다.'); return; }
  const reader = new FileReader();
  reader.onload = () => {
    const src = reader.result;
    const img = new Image();
    img.onload = () => {
      const w0 = img.naturalWidth || 200;
      const h0 = img.naturalHeight || 150;
      const k = Math.min(1, 480 / w0, 360 / h0);
      const p = at ?? objectOrigin();
      addObject('images', {
        id: newObjId('im'), name: file.name.replace(/\.[^.]+$/, '') || '그림', x: p.x, y: p.y,
        w: Math.max(8, Math.round(w0 * k)), h: Math.max(8, Math.round(h0 * k)), src,
      });
      focusGrid();
    };
    img.onerror = () => alertDialog('그림 삽입', '이 그림 형식은 표시할 수 없습니다. PNG·JPEG·GIF·SVG·WebP 파일을 사용하세요.');
    img.src = src;
  };
  reader.readAsDataURL(file);
}

function resetImageSize(id) {
  const im = sheet().images.find((x) => x.id === id);
  if (!im) return;
  const img = new Image();
  img.onload = () => updateObject(id, { w: img.naturalWidth || im.w, h: img.naturalHeight || im.h });
  img.src = im.src;
}

function imageDialog(id) {
  const im = sheet().images.find((x) => x.id === id);
  if (!im) return;
  formDialog('그림 서식', [
    { name: 'name', label: '이름', value: im.name ?? '' },
    { name: 'w', label: '너비(px)', type: 'number', value: im.w },
    { name: 'h', label: '높이(px)', type: 'number', value: im.h },
    { name: 'lock', label: '가로 세로 비율 고정', type: 'checkbox', value: true },
  ], ({ name, w, h, lock }) => {
    let nw = clamp(Math.round(Number(w) || im.w), 4, 20000);
    let nh = clamp(Math.round(Number(h) || im.h), 4, 20000);
    if (lock) {
      if (nw !== im.w) nh = Math.max(4, Math.round((nw * im.h) / im.w));
      else if (nh !== im.h) nw = Math.max(4, Math.round((nh * im.w) / im.h));
    }
    updateObject(id, { name: name.trim() || im.name, w: nw, h: nh });
  });
}

// 도형
function startDraw(kind) {
  if (editing && !commitEdit()) return;
  drawKind = kind;
  deselectChart();
  dom.view.classList.add('drawing-mode');
  const label = SHAPE_KINDS.find((k) => k.id === kind)?.label ?? '도형';
  toast(`시트를 끌어서 ${label}을(를) 그리세요. 클릭하면 기본 크기로 들어갑니다. (Esc: 취소)`);
}

function endDraw() {
  drawKind = null;
  dom.view.classList.remove('drawing-mode');
}

function shapeDialog(id, typed = null) {
  const sh = sheet().shapes.find((x) => x.id === id);
  if (!sh) return;
  const isLine = sh.kind === 'line';
  const row = (label, ...inputs) => el('label', {}, el('span', {}, label), el('span', { style: { display: 'flex', gap: '8px', alignItems: 'center' } }, ...inputs));
  const text = el('textarea', { rows: 4, style: { width: '100%', minHeight: '80px' } }, typed ?? sh.text ?? '');
  const size = el('input', { type: 'number', min: 6, max: 96, value: sh.size ?? 11, style: { width: '70px' } });
  const bold = el('input', { type: 'checkbox', checked: !!sh.bold });
  const align = el('select', {}, [['left', '왼쪽'], ['center', '가운데'], ['right', '오른쪽']].map(([v, l]) => el('option', { value: v, selected: (sh.align ?? 'center') === v }, l)));
  const color = el('input', { type: 'color', value: sh.color ?? '#000000' });
  const fillOn = el('input', { type: 'checkbox', checked: !!sh.fill });
  const fill = el('input', { type: 'color', value: sh.fill ?? '#4472c4' });
  const lineOn = el('input', { type: 'checkbox', checked: !!sh.stroke });
  const stroke = el('input', { type: 'color', value: sh.stroke ?? '#2f528f' });
  const width = el('input', { type: 'number', min: 0.25, max: 20, step: 0.25, value: sh.strokeWidth ?? 1, style: { width: '70px' } });
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
    isLine ? null : row('텍스트', text),
    isLine ? null : row('글꼴', el('span', {}, '크기'), size, el('label', { style: { display: 'inline-flex', gap: '4px', flexDirection: 'row' } }, bold, '굵게'), color),
    isLine ? null : row('맞춤', align),
    isLine ? null : row('채우기', el('label', { style: { display: 'inline-flex', gap: '4px', flexDirection: 'row' } }, fillOn, '사용'), fill),
    row(isLine ? '선 색' : '윤곽선', isLine ? null : el('label', { style: { display: 'inline-flex', gap: '4px', flexDirection: 'row' } }, lineOn, '사용'), stroke),
    row('선 굵기(pt)', width));
  openDialog({
    title: isLine ? '선 서식' : '도형 서식', body, width: 420,
    onOpen: () => { if (!isLine) { text.focus(); text.setSelectionRange(text.value.length, text.value.length); } },
    buttons: [
      {
        label: '확인', primary: true,
        action: () => updateObject(id, isLine
          ? { stroke: stroke.value, strokeWidth: Number(width.value) || 1 }
          : {
            text: text.value, size: clamp(Number(size.value) || 11, 6, 96), bold: bold.checked || undefined, align: align.value, color: color.value,
            fill: fillOn.checked ? fill.value : null, stroke: lineOn.checked ? stroke.value : null, strokeWidth: Number(width.value) || 1,
          }),
      },
      { label: '취소' },
    ],
  });
}

// ───────────────────────── 데이터 유효성 검사 ─────────────────────────
function validationError(rule, dir, fillSel) {
  const style = rule.errorStyle ?? 'stop';
  const retry = () => { if (editing) setTimeout(() => { dom.editor.focus(); dom.editor.select(); }, 0); };
  const cancel = () => cancelEdit();
  const accept = () => { if (editing) { editing.dvOk = true; setTimeout(() => commitEdit(dir, { fillSel }), 0); } };
  const body = el('div', { class: 'dv-error' },
    el('span', { class: `dv-icon ${style}` }, style === 'stop' ? '✕' : style === 'warning' ? '!' : 'i'),
    el('div', {}, rule.error || describeRule(rule), style === 'warning' ? el('div', { style: { marginTop: '8px' } }, '계속하시겠습니까?') : null));
  const buttons = style === 'stop'
    ? [{ label: '다시 시도', primary: true, action: retry }, { label: '취소', action: cancel }]
    : style === 'warning'
      ? [{ label: '예', primary: true, action: accept }, { label: '아니요', action: retry }, { label: '취소', action: cancel }]
      : [{ label: '확인', primary: true, action: accept }, { label: '취소', action: cancel }];
  openDialog({ title: rule.errorTitle || 'Tabula', body, buttons, width: 400 });
}

function openDvList() {
  const rule = validationAt(sheet(), active.r, active.c);
  if (!rule || rule.type !== 'list') return;
  const items = listItems(wb, si, rule);
  const m = wb.mergeAt(si, active.r, active.c) ?? { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
  const b = gv.clientRect(m);
  const cur = displayText(active.r, active.c);
  const pick = (it) => {
    wb.transact(() => { wb.setInput(si, active.r, active.c, it); autoWiden({ r1: active.r, c1: active.c, r2: active.r, c2: active.c }); }, meta());
    if (circles) circles = invalidCells(wb, si);
    focusGrid();
  };
  const menu = openMenu({ x: b.left, y: b.bottom },
    items.length ? items.slice(0, 1000).map((it) => ({ label: it, checked: it === cur, action: () => pick(it) })) : [{ label: '(목록이 비어 있습니다)', disabled: true }],
    { minWidth: Math.max(90, b.width + 17) });
  menu.classList.add('dv-list');
}

function updateDvPrompt() {
  const rule = chartSel || editing ? null : validationAt(sheet(), active.r, active.c);
  const b = rule ? gv.clientRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c }) : null;
  const v = dom.view.getBoundingClientRect();
  const visible = b && b.bottom > v.top && b.top < v.bottom && b.right > v.left && b.left < v.right;
  if (!rule || rule.showPrompt === false || !(rule.prompt || rule.promptTitle) || !visible) {
    dvPromptEl?.remove();
    dvPromptEl = null;
    return;
  }
  if (!dvPromptEl) { dvPromptEl = el('div', { class: 'dv-prompt' }); document.body.append(dvPromptEl); }
  dvPromptEl.replaceChildren(...(rule.promptTitle ? [el('b', {}, rule.promptTitle)] : []), rule.prompt ?? '');
  dvPromptEl.style.left = `${Math.min(b.left + Math.min(b.width, 48), innerWidth - 250)}px`;
  dvPromptEl.style.top = `${b.bottom + 6}px`;
}

function validationDialog() {
  if (editing && !commitEdit()) return;
  const cur = validationAt(sheet(), active.r, active.c) ?? {};
  const mixed = (sheet().validations ?? []).some((v) => v !== cur && v.r1 <= sel.r2 && v.r2 >= sel.r1 && v.c1 <= sel.c2 && v.c2 >= sel.c1);
  const label = (text, input) => el('label', {}, el('span', {}, text), input);
  const inline = (input, text) => el('label', { style: { display: 'flex', flexDirection: 'row', alignItems: 'center', gap: '6px' } }, input, text);
  const type = el('select', {}, VALIDATION_TYPES.map((t) => el('option', { value: t.id, selected: (cur.type ?? 'any') === t.id }, t.label)));
  const op = el('select', {}, VALIDATION_OPS.map((o) => el('option', { value: o.id, selected: (cur.op ?? 'between') === o.id }, o.label)));
  const strip = (f) => {
    const t = String(f ?? '');
    return t.startsWith('"') && t.endsWith('"') ? t.slice(1, -1) : t && /^[A-Z$]|!/i.test(t) && parseRangeName(t.replace(/\$/g, '').split('!').pop()) ? `=${t}` : t;
  };
  const f1 = el('input', { type: 'text', value: strip(cur.f1) });
  const f2 = el('input', { type: 'text', value: strip(cur.f2) });
  const blank = el('input', { type: 'checkbox', checked: cur.allowBlank !== false });
  const dropdown = el('input', { type: 'checkbox', checked: cur.showDropdown !== false });
  const f1Label = el('span', {});
  const f2Label = el('span', {}, '최대값');
  const opRow = label('제한 방법', op);
  const f1Row = el('label', {}, f1Label, f1);
  const f2Row = el('label', {}, f2Label, f2);
  const ddRow = inline(dropdown, '드롭다운 표시');
  const refresh = () => {
    const t = type.value;
    const ranged = ['whole', 'decimal', 'date', 'time', 'textLength'].includes(t);
    const two = ranged && (op.value === 'between' || op.value === 'notBetween');
    opRow.style.display = ranged ? '' : 'none';
    f1Row.style.display = t === 'any' ? 'none' : '';
    f2Row.style.display = two ? '' : 'none';
    ddRow.style.display = t === 'list' ? '' : 'none';
    f1Label.textContent = t === 'list' ? '원본' : t === 'custom' ? '수식' : two ? '최소값' : { equal: '값', notEqual: '값', greaterThan: '최소값', greaterThanOrEqual: '최소값', lessThan: '최대값', lessThanOrEqual: '최대값' }[op.value] ?? '값';
    f1.placeholder = t === 'list' ? '예: 사과,배,포도 또는 =$A$1:$A$5' : t === 'custom' ? '예: =A1>0' : t === 'date' ? '예: 2024-01-01' : t === 'time' ? '예: 9:00' : '';
    f2.placeholder = f1.placeholder;
  };
  type.addEventListener('change', refresh);
  op.addEventListener('change', refresh);
  refresh();
  const showPrompt = el('input', { type: 'checkbox', checked: cur.showPrompt !== false });
  const promptTitle = el('input', { type: 'text', value: cur.promptTitle ?? '' });
  const prompt = el('textarea', { rows: 4 }, cur.prompt ?? '');
  const showError = el('input', { type: 'checkbox', checked: cur.showError !== false });
  const errorStyle = el('select', {}, [['stop', '중지'], ['warning', '경고'], ['info', '정보']].map(([v, l]) => el('option', { value: v, selected: (cur.errorStyle ?? 'stop') === v }, l)));
  const errorTitle = el('input', { type: 'text', value: cur.errorTitle ?? '' });
  const error = el('textarea', { rows: 4 }, cur.error ?? '');
  const col = (...kids) => el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } }, ...kids);
  const pages = [
    ['설정', col(el('div', { class: 'muted' }, '유효성 조건'), label('제한 대상', type), inline(blank, '공백 무시'), opRow, f1Row, f2Row, ddRow)],
    ['설명 메시지', col(inline(showPrompt, '셀을 선택하면 설명 메시지 표시'), label('제목', promptTitle), label('설명 메시지', prompt))],
    ['오류 메시지', col(inline(showError, '유효하지 않은 데이터를 입력하면 오류 메시지 표시'), label('스타일', errorStyle), label('제목', errorTitle), label('오류 메시지', error))],
  ];
  const tabBar = el('div', { class: 'dlg-tabs' });
  const pageBox = el('div', { style: { minHeight: '260px', paddingTop: '10px' } });
  const show = (i) => {
    [...tabBar.children].forEach((b, j) => b.classList.toggle('on', i === j));
    pageBox.replaceChildren(pages[i][1]);
  };
  pages.forEach(([name], i) => tabBar.append(el('button', { type: 'button', class: 'dlg-tab', onclick: () => show(i) }, name)));
  show(0);
  const target = { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 };
  const apply = (rule) => {
    const rest = (sheet().validations ?? []).flatMap((v) => subtractRange(v, target));
    wb.transact(() => wb.setSheetProp(si, 'validations', rule ? [...rest, { ...target, ...rule }] : rest), meta());
    if (circles) circles = invalidCells(wb, si);
    gv.renderOverlays();
    updateSelectionUI();
  };
  const toFormula = (v, t) => {
    const x = v.trim();
    if (!x) return undefined;
    if (t === 'list') return x.startsWith('=') ? x.slice(1) : x;
    if (x.startsWith('=')) return x.slice(1);
    if (t === 'date' || t === 'time') { const p = parseInput(x).value; return typeof p === 'number' ? String(p) : x; }
    return x;
  };
  openDialog({
    title: '데이터 유효성', width: 460,
    body: el('div', {}, mixed ? el('div', { class: 'muted', style: { marginBottom: '6px' } }, '선택 영역에 다른 유효성 검사 설정이 섞여 있습니다. 확인을 누르면 모두 이 설정으로 바뀝니다.') : null, tabBar, pageBox),
    buttons: [
      { label: '모두 지우기', action: () => apply(null) },
      {
        label: '확인', primary: true,
        action: () => {
          const t = type.value;
          const rule = {
            type: t, op: op.value, allowBlank: blank.checked, showDropdown: dropdown.checked,
            showPrompt: showPrompt.checked, showError: showError.checked, errorStyle: errorStyle.value,
          };
          if (!['whole', 'decimal', 'date', 'time', 'textLength'].includes(t)) delete rule.op;
          if (t !== 'any') {
            rule.f1 = toFormula(f1.value, t);
            if (!rule.f1) { show(0); alertDialog('데이터 유효성', t === 'list' ? '원본을 입력하세요.' : '값을 입력하세요.'); return false; }
            if (rule.op === 'between' || rule.op === 'notBetween') {
              rule.f2 = toFormula(f2.value, t);
              if (!rule.f2) { show(0); alertDialog('데이터 유효성', '최대값을 입력하세요.'); return false; }
            }
          }
          for (const [k, input] of [['promptTitle', promptTitle], ['prompt', prompt], ['errorTitle', errorTitle], ['error', error]]) if (input.value.trim()) rule[k] = input.value;
          apply(t === 'any' && !rule.prompt && !rule.promptTitle ? null : rule);
          return undefined;
        },
      },
      { label: '취소' },
    ],
  });
}

// ───────────────────────── 매크로 ─────────────────────────
let vbaCache = null;
function macroDialog() {
  if (!wb.vba?.bin) {
    alertDialog('매크로', '이 통합 문서에는 매크로가 없습니다. 매크로가 포함된 엑셀 파일(.xlsm)을 열면 여기에서 VBA 코드를 볼 수 있습니다.');
    return;
  }
  if (vbaCache?.bin !== wb.vba.bin) {
    let modules;
    try { modules = extractVbaModules(fromBase64(wb.vba.bin)); } catch (err) { modules = null; toast(`매크로 코드를 읽지 못했습니다: ${err.message}`); }
    vbaCache = { bin: wb.vba.bin, modules: modules ?? [] };
  }
  const modules = vbaCache.modules;
  const withCode = modules.filter((m) => m.code.trim());
  const list = withCode.length ? withCode : modules;
  const procs = (code) => [...code.matchAll(/^\s*(?:Public |Private )?(?:Sub|Function)\s+([\w가-힣]+)/gim)].map((m) => m[1]);
  const pre = el('pre', { class: 'macro-code' });
  const selEl = el('select', { style: { width: '100%' } }, list.map((m, i) => el('option', { value: i }, `${m.name} (${m.type})${procs(m.code).length ? ` — ${procs(m.code).join(', ')}` : ''}`)));
  const showMod = () => { pre.textContent = list[Number(selEl.value)]?.code || '(코드 없음)'; };
  selEl.addEventListener('change', showMod);
  showMod();
  openDialog({
    title: '매크로 (VBA)', width: 640,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
      el('div', { class: 'muted' }, 'Tabula 는 보안을 위해 매크로를 실행하지 않습니다. 코드는 읽기 전용으로 표시되며, 파일을 .xlsm 으로 저장하면 매크로가 그대로 보존되어 엑셀에서 실행할 수 있습니다.'),
      list.length ? selEl : el('div', {}, '표시할 모듈이 없습니다.'),
      pre),
    buttons: [
      { label: '코드 복사', action: () => { navigator.clipboard?.writeText(pre.textContent).then(() => toast('코드를 복사했습니다.')).catch(() => {}); return false; } },
      {
        label: '매크로 제거',
        action: () => {
          openDialog({
            title: '매크로 제거', body: '이 통합 문서의 매크로를 모두 제거합니다. 이후에는 .xlsx 로 저장됩니다. 계속하시겠습니까?',
            buttons: [{ label: '제거', primary: true, action: () => { wb.vba = null; onBookChange(); toast('매크로를 제거했습니다.'); } }, { label: '취소' }],
          });
        },
      },
      { label: '닫기', primary: true },
    ],
  });
}

// ───────────────────────── 피벗 테이블 ─────────────────────────
function pivotSourceRows(src, rg) {
  const s = wb.sheetIndexByName(src);
  if (s < 0) return null;
  const rows = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [];
    for (let c = rg.c1; c <= rg.c2; c++) row.push(wb.getValue(s, r, c));
    rows.push(row);
  }
  return rows;
}

function pivotSource(def) {
  return pivotSourceData(wb, def)?.rows ?? null;
}

const pivotItemText = itemText;

function writePivot(targetSi, def) {
  const src = pivotSource(def);
  if (!src) return false;
  const { def: d, rows } = resolvePivot(src, def);
  const out = buildPivot(rows, d);
  const t = wb.sheets[targetSi];
  const top = def.top ?? 0;
  const left = def.left ?? 0;
  const colsN = out[0]?.length ?? 0;
  if (def.area) {
    // 엑셀에서 가져온 피벗: 그 영역만 지우고 다시 씀 (시트의 다른 내용은 유지)
    const a = def.area;
    for (const k of [...t.cells.keys()]) {
      const [r, c] = k.split(',').map(Number);
      if (r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2) wb.setCellData(targetSi, r, c, null);
    }
    def.area = { r1: top, c1: left, r2: top + out.length - 1, c2: left + colsN - 1 };
  } else {
    for (const k of [...t.cells.keys()]) { const [r, c] = k.split(',').map(Number); wb.setCellData(targetSi, r, c, null); }
  }
  out.forEach((row, r) => row.forEach((cd, c) => { if (cd) wb.setCellData(targetSi, top + r, left + c, cd); }));
  for (let c = 0; c < colsN; c++) {
    let w = 0;
    out.forEach((row, r) => {
      if (row[c]?.raw) w = Math.max(w, measureText(displayText(top + r, left + c, targetSi), wb.styleAt(targetSi, top + r, left + c)) + 12);
    });
    if (w > (wb.sheets[targetSi].colWidths[left + c] ?? DEFAULT_COL_WIDTH)) wb.setColWidth(targetSi, left + c, Math.min(300, Math.ceil(w)));
  }
  return true;
}

function pivotDialog(tableName = null) {
  const inTable = tableName ? findTable(wb, tableName)?.t : tableHere();
  const rg = inTable ? { r1: inTable.r1, c1: inTable.c1, r2: dataBottom(inTable), c2: inTable.c2 } : dataRange();
  if (rg.r2 <= rg.r1) { alertDialog('피벗 테이블', '머리글 행과 데이터가 있는 범위를 선택하세요.'); return; }
  const srcName = sheet().name;
  const headers = [];
  for (let c = rg.c1; c <= rg.c2; c++) headers.push({ value: String(c - rg.c1), label: displayText(rg.r1, c) || `${colToName(c)}열` });
  const firstNum = headers.findIndex((h, i) => typeof valueAt(rg.r1 + 1, rg.c1 + i) === 'number');
  formDialog('피벗 테이블 만들기', [
    { name: 'range', label: '표/범위', value: inTable ? inTable.name : `${quoteSheetName(srcName)}!${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}` },
    { name: 'row', label: '행', type: 'select', value: '0', options: headers },
    { name: 'col', label: '열', type: 'select', value: '', options: [{ value: '', label: '(없음)' }, ...headers] },
    { name: 'val', label: '값', type: 'select', value: firstNum >= 0 ? String(firstNum) : '', options: [{ value: '', label: '(행 개수)' }, ...headers] },
    { name: 'agg', label: '계산', type: 'select', value: firstNum >= 0 ? 'sum' : 'count', options: AGGREGATES.map((a) => ({ value: a.id, label: a.label })) },
  ], (v) => {
    const text = v.range.trim().replace(/^=/, '');
    const tf = findTable(wb, text);
    let def;
    const names = (i) => (i === '' ? null : headers[Number(i)]?.label ?? null);
    const fieldNames = { row: names(v.row), col: names(v.col), val: names(v.val) };
    const fields = {
      rowField: Number(v.row), colField: v.col === '' ? null : Number(v.col), valueField: v.val === '' ? null : Number(v.val),
      agg: v.val === '' ? 'count' : v.agg, fieldNames,
    };
    if (tf) {
      def = { table: tf.t.name, source: wb.sheets[tf.si].name, range: { r1: tf.t.r1, c1: tf.t.c1, r2: dataBottom(tf.t), c2: tf.t.c2 }, ...fields };
    } else {
      const bang = text.lastIndexOf('!');
      const src = bang > 0 ? text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : srcName;
      const prg = parseRangeName((bang > 0 ? text.slice(bang + 1) : text).replace(/\$/g, ''));
      if (!prg || wb.sheetIndexByName(src) < 0) { toast('범위가 올바르지 않습니다. 셀 범위나 표 이름을 입력하세요.'); return false; }
      def = { source: src, range: prg, ...fields };
    }
    let n = 1;
    while (wb.sheetIndexByName(`피벗${n}`) >= 0) n++;
    const at = wb.transact(() => {
      const idx = wb.addSheet(`피벗${n}`, si + 1);
      writePivot(idx, def);
      wb.setSheetProp(idx, 'pivot', def);
      return idx;
    }, meta());
    switchSheet(at, false);
    toast(def.table
      ? `'${def.table}' 표로 피벗 테이블을 만들었습니다. 표에 데이터가 늘어나면 [데이터 → 모두 새로 고침]으로 자동 반영됩니다.`
      : '피벗 테이블을 만들었습니다. 원본이 바뀌면 [데이터 → 모두 새로 고침]을 누르세요.');
    return true;
  }, { note: '표 이름(예: 표1)을 원본으로 쓰면 표에 추가된 데이터까지 새로 고칠 때 자동으로 포함됩니다. 결과는 새 워크시트에 만들어집니다.' });
}

function refreshPivots() {
  let n = 0;
  wb.transact(() => {
    wb.sheets.forEach((s, i) => { if (s.pivot && writePivot(i, s.pivot)) n++; });
  }, meta());
  toast(n ? `피벗 테이블 ${n}개를 새로 고쳤습니다.` : '새로 고칠 피벗 테이블이 없습니다.');
}

// ───────────────────────── 찾기 / 바꾸기 ─────────────────────────
const findState = { text: '', replace: '', matchCase: false, whole: false };

function findNext(text = findState.text, { quiet = false } = {}) {
  if (!text) return false;
  const u = wb.usedRange(si);
  const total = u.rows * Math.max(1, u.cols);
  if (!total) { if (!quiet) toast('찾을 수 없습니다.'); return false; }
  const matches = (r, c) => {
    const raw = wb.getRaw(si, r, c);
    if (!raw) return false;
    return [raw, displayText(r, c)].some((h) => {
      const a = findState.matchCase ? h : h.toLowerCase();
      const b = findState.matchCase ? text : text.toLowerCase();
      return findState.whole ? a === b : a.includes(b);
    });
  };
  const cols = Math.max(1, u.cols);
  const start = Math.min(active.r * cols + active.c, total - 1);
  for (let k = 1; k <= total; k++) {
    const idx = (start + k) % total;
    const r = Math.floor(idx / cols);
    const c = idx % cols;
    if (!gv.rows.isHidden(r) && matches(r, c)) { selectCell(r, c); return true; }
  }
  if (!quiet) toast(`'${text}'을(를) 찾을 수 없습니다.`);
  return false;
}

function replaceIn(r, c) {
  const raw = wb.getRaw(si, r, c);
  const escRe = findState.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(findState.whole ? `^${escRe}$` : escRe, findState.matchCase ? 'g' : 'gi');
  const next = raw.replace(re, () => findState.replace);
  if (next !== raw) { wb.setInput(si, r, c, next); return true; }
  return false;
}

function openFindDialog(tab = 'find') {
  const findInput = el('input', { type: 'text', value: findState.text || displayText(active.r, active.c) || '' });
  const replInput = el('input', { type: 'text', value: findState.replace });
  const caseBox = el('input', { type: 'checkbox', checked: findState.matchCase });
  const wholeBox = el('input', { type: 'checkbox', checked: findState.whole });
  const sync = () => Object.assign(findState, { text: findInput.value, replace: replInput.value, matchCase: caseBox.checked, whole: wholeBox.checked });
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
    el('label', {}, el('span', {}, '찾을 내용'), findInput),
    tab === 'replace' ? el('label', {}, el('span', {}, '바꿀 내용'), replInput) : null,
    el('label', {}, caseBox, '대/소문자 구분'),
    el('label', {}, wholeBox, '전체 셀 내용 일치'));
  const buttons = [];
  if (tab === 'replace') {
    buttons.push({
      label: '모두 바꾸기', action: () => {
        sync();
        if (!findState.text) return false;
        let n = 0;
        const u = wb.usedRange(si);
        wb.transact(() => { for (const [r, c] of cellsIn({ r1: 0, c1: 0, r2: u.rows - 1, c2: u.cols - 1 })) if (replaceIn(r, c)) n++; }, meta());
        toast(`${n}개 항목을 바꾸었습니다.`);
        return false;
      },
    }, {
      label: '바꾸기', action: () => {
        sync();
        wb.transact(() => replaceIn(active.r, active.c), meta());
        findNext(findState.text, { quiet: true });
        return false;
      },
    });
  }
  buttons.push({ label: '다음 찾기', primary: true, action: () => { sync(); findNext(); return false; } }, { label: '닫기' });
  openDialog({ title: '찾기 및 바꾸기', body, buttons, width: 420 });
}

// ───────────────────────── 메모 ─────────────────────────
function editComment() {
  const cur = wb.getCell(si, active.r, active.c)?.comment ?? '';
  formDialog(`메모 - ${cellName(active.r, active.c)}`, [{ name: 'text', label: '내용', type: 'textarea', value: cur }], ({ text }) => {
    wb.transact(() => wb.setComment(si, active.r, active.c, text.trim()), meta());
  }, { okLabel: '저장' });
}

function jumpComment(dir) {
  const list = [];
  for (const [k, cell] of sheet().cells) if (cell.comment) list.push(k.split(',').map(Number));
  if (!list.length) { toast('메모가 없습니다.'); return; }
  list.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const cur = active.r * MAX_COLS + active.c;
  const idx = (p) => p[0] * MAX_COLS + p[1];
  const next = dir > 0 ? list.find((p) => idx(p) > cur) ?? list[0] : [...list].reverse().find((p) => idx(p) < cur) ?? list.at(-1);
  selectCell(next[0], next[1]);
}

// ───────────────────────── 시트 ─────────────────────────
function switchSheet(i, restore = true) {
  if (i === si || i < 0 || i >= wb.sheets.length) return;
  if (editing && !commitEdit()) return;
  if (wb.sheets[si]) sheetSel.set(wb.sheets[si], { active, sel, selKind, scroll: [gv.sx, gv.sy] });
  si = i;
  chartSel = null;
  circles = null;
  gv.resetExtent();
  gv.layout();
  const saved = restore ? sheetSel.get(sheet()) : null;
  if (saved) {
    selectRange(saved.sel, saved.selKind, saved.active);
    gv.setScroll(...saved.scroll);
  } else {
    gv.setScroll(0, 0);
    const f = sheet().freeze;
    selectCell(f?.rows || 0, f?.cols || 0);
  }
  renderSheetTabs();
  setMode();
}

function renderSheetTabs() {
  dom.sheetTabs.replaceChildren(...wb.sheets.map((s, i) => el('button', {
    class: `sheet-tab${i === si ? ' active' : ''}`,
    title: s.pivot ? `피벗 테이블 (원본: ${s.pivot.source})` : undefined,
    onmousedown: (e) => { if (e.button === 0) { e.preventDefault(); switchSheet(i); focusGrid(); } },
    ondblclick: () => renameSheetInline(i),
    oncontextmenu: (e) => {
      e.preventDefault();
      switchSheet(i);
      openMenu({ x: e.clientX, y: e.clientY - 150 }, [
        { label: '삽입', icon: 'sheetInsert', action: () => run('addSheet') },
        { label: '삭제', icon: 'delete', action: () => run('deleteSheet') },
        { label: '이름 바꾸기', action: () => renameSheetInline(i) },
        { label: '복사본 만들기', icon: 'copy', action: () => run('duplicateSheet') },
        { sep: true },
        { label: '왼쪽으로 이동', disabled: i === 0, action: () => moveSheet(i, -1) },
        { label: '오른쪽으로 이동', disabled: i === wb.sheets.length - 1, action: () => moveSheet(i, 1) },
      ]);
    },
  }, s.name)));
  dom.sheetTabs.children[si]?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

function renameSheetInline(i) {
  const tab = dom.sheetTabs.children[i];
  if (!tab) return;
  const input = el('input', { value: wb.sheets[i].name, maxlength: 31 });
  tab.replaceChildren(input);
  input.focus();
  input.select();
  let done = false;
  const finish = (ok) => {
    if (done) return;
    done = true;
    if (ok && input.value.trim() !== wb.sheets[i].name) {
      const res = wb.transact(() => wb.renameSheet(i, input.value), meta());
      if (!res) toast('시트 이름이 올바르지 않거나 이미 사용 중입니다.');
    }
    renderSheetTabs();
    focusGrid();
  };
  input.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') finish(true);
    if (e.key === 'Escape') finish(false);
  });
  input.addEventListener('blur', () => finish(true));
  input.addEventListener('mousedown', (e) => e.stopPropagation());
}

function moveSheet(i, d) {
  const j = i + d;
  if (j < 0 || j >= wb.sheets.length) return;
  wb.transact(() => {
    wb.snapshotAll();
    const [s] = wb.sheets.splice(i, 1);
    wb.sheets.splice(j, 0, s);
  }, meta());
  si = j;
  renderSheetTabs();
}

// ───────────────────────── 파일 ─────────────────────────
function download(name, content, type) {
  const blob = content instanceof Blob ? content : new Blob([content], { type });
  const a = el('a', { href: URL.createObjectURL(blob), download: name });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

const safeFileName = (name) => name.replace(/[\\/:*?"<>|]/g, '_').trim() || '통합 문서';

function sheetToRows(index) {
  const u = wb.usedRange(index);
  const rows = [];
  for (let r = 0; r < u.rows; r++) {
    const row = [];
    for (let c = 0; c < u.cols; c++) row.push(displayText(r, c, index));
    rows.push(row);
  }
  return rows;
}

function exportCsv() {
  download(`${safeFileName(docName)}-${safeFileName(sheet().name)}.csv`, `﻿${toDelimited(sheetToRows(si))}`, 'text/csv;charset=utf-8');
  toast('CSV 파일로 내보냈습니다.');
}

function exportXlsx(name = docName) {
  const over = xlsxOverflow(wb);
  if (over) toast(`엑셀 파일은 1,048,576행까지만 저장할 수 있어 그 아래 셀 ${over.toLocaleString()}개는 빠집니다. 전체는 .tabula 로 저장하세요.`);
  try {
    const bytes = writeXlsx(wb, { activeSheet: si });
    if (wb.vba) {
      download(`${safeFileName(name)}.xlsm`, new Blob([bytes], { type: 'application/vnd.ms-excel.sheet.macroEnabled.12' }));
      toast('Excel 매크로 사용 통합 문서(.xlsm)로 저장했습니다.');
    } else {
      download(`${safeFileName(name)}.xlsx`, new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
      toast('Excel 통합 문서(.xlsx)로 저장했습니다.');
    }
  } catch (err) {
    alertDialog('Tabula', `저장하지 못했습니다: ${err.message}`);
  }
}

function saveAs() {
  formDialog('다른 이름으로 저장', [
    { name: 'name', label: '파일 이름', value: docName },
    {
      name: 'type', label: '파일 형식', type: 'select', value: 'xlsx', options: [
        { value: 'xlsx', label: wb.vba ? 'Excel 매크로 사용 통합 문서 (*.xlsm)' : 'Excel 통합 문서 (*.xlsx)' },
        { value: 'tabula', label: 'Tabula 통합 문서 (*.tabula)' },
        { value: 'csv', label: 'CSV UTF-8 (쉼표로 분리) — 현재 시트' },
        ...(server.available ? [{ value: 'server', label: '서버에 저장 (다른 기기에서 열기)' }] : []),
      ],
    },
  ], ({ name, type }) => {
    const newName = name.trim() || docName;
    if (newName !== docName) renameDoc(newName);
    if (type === 'xlsx') exportXlsx(newName);
    else if (type === 'csv') exportCsv();
    else if (type === 'server') saveNow(true);
    else download(`${safeFileName(newName)}.tabula`, JSON.stringify(snapshot()), 'application/json');
    saveToStorage();
  }, { okLabel: '저장' });
}

let fileMode = 'open';
function pickFile(mode) {
  fileMode = mode;
  dom.fileInput.value = '';
  dom.fileInput.click();
}

async function onFilePicked() {
  const file = dom.fileInput.files[0];
  if (file) await openFileObject(file, fileMode);
}

async function openFileObject(file, mode) {
  fileMode = mode;
  const base = file.name.replace(/\.[^.]+$/, '');
  try {
    if (/\.(xlsx|xlsm)$/i.test(file.name)) {
      const { data, warnings, active: act } = readXlsx(new Uint8Array(await file.arrayBuffer()));
      if (fileMode === 'open') {
        loadWorkbook(data, base, act);
      } else {
        wb.transact(() => {
          for (const s of data.sheets) {
            let name = s.name;
            for (let n = 2; wb.sheetIndexByName(name) >= 0; n++) name = `${s.name} (${n})`.slice(0, 31);
            const at = wb.addSheet(name);
            const all = wb.serialize();
            all.sheets[at] = { ...s, name };
            wb.restore(all);
          }
        }, meta());
        toast(`시트 ${data.sheets.length}개를 가져왔습니다.`);
      }
      if (fileMode === 'open' && data.vba) warnings.push('매크로가 포함된 통합 문서입니다. Tabula 는 매크로를 실행하지 않지만 [보기 → 매크로]에서 코드를 볼 수 있고, .xlsm 으로 저장하면 매크로가 그대로 유지됩니다.');
      if (warnings.length) alertDialog('가져오기', warnings.join('\n'));
      else if (fileMode === 'open') toast(`'${file.name}'을(를) 열었습니다.`);
      return;
    }
    const text = await file.text();
    if (/\.(json|tabula)$/i.test(file.name)) {
      const data = JSON.parse(text);
      loadWorkbook(data.workbook ?? data, data.docName ?? base, data.si ?? 0);
      toast(`'${file.name}'을(를) 열었습니다.`);
      return;
    }
    const rows = parseDelimited(text.replace(/^﻿/, ''), guessDelimiter(text));
    if (fileMode === 'open') {
      loadWorkbook({ sheets: [{ name: base.slice(0, 31) || 'Sheet1', cells: {} }] }, base);
      writeRows(rows, 0, 0);
    } else {
      writeRows(rows, active.r, active.c);
    }
    toast(`${rows.length}개 행을 가져왔습니다.`);
  } catch (err) {
    alertDialog('Tabula', `파일을 열 수 없습니다: ${err.message}`);
  }
}

function writeRows(rows, r0, c0) {
  wb.transact(() => rows.forEach((row, i) => row.forEach((v, j) => { if (v !== '') wb.setInput(si, r0 + i, c0 + j, v); })), meta());
  if (rows.length) selectRange({ r1: r0, c1: c0, r2: r0 + rows.length - 1, c2: c0 + Math.max(1, ...rows.map((r) => r.length)) - 1 }, 'cells', { r: r0, c: c0 });
}

function loadWorkbook(data, name, activeSheet = 0) {
  if (editing) endEditUI();
  wb.load(data);
  docName = name || '통합 문서1';
  si = clamp(activeSheet, 0, wb.sheets.length - 1);
  clip = null;
  painter = null;
  chartSel = null;
  circles = null;
  objClip = null;
  sheetSel.clear();
  gv.resetExtent();
  renderAll();
  gv.setScroll(0, 0);
  const f = sheet().freeze;
  selectCell(f?.rows || 0, f?.cols || 0);
  dirty = true;
  saveToStorage();
  scheduleServerSave(0);
}

async function serverNames() {
  if (!server.available) return [];
  try { return (await server.list()).map((f) => f.name); } catch { return []; }
}

async function newWorkbook(sample) {
  const go = async () => {
    const names = await serverNames();
    if (sample) {
      let name = sample.name;
      for (let n = 2; names.includes(name); n++) name = `${sample.name} ${n}`;
      loadWorkbook(sample.build(), name);
    } else {
      let n = 1;
      while (names.includes(`통합 문서${n}`) || `통합 문서${n}` === docName) n++;
      loadWorkbook({ sheets: [{ name: 'Sheet1', cells: {} }] }, `통합 문서${n}`);
    }
  };
  if (!autosave && dirty) {
    openDialog({
      title: 'Tabula', body: '저장하지 않은 변경 내용이 있습니다. 새 통합 문서를 만드시겠습니까?',
      buttons: [{ label: '새로 만들기', primary: true, action: go }, { label: '취소' }],
    });
  } else go();
}

function renameDoc(name) {
  const old = docName;
  docName = name;
  updateTitle();
  if (server.available && old !== name) {
    server.save(name, snapshot()).then(() => server.remove(old)).catch(() => {});
  }
}

const snapshot = () => ({ app: 'tabula', docName, si, workbook: wb.serialize() });

let storageWarned = false;
function saveToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ docName, si, autosave, workbook: wb.serialize() }));
    if (!server.available) dirty = false;
    updateTitle();
    storageWarned = false;
    return true;
  } catch {
    // 큰 그림 등으로 브라우저 저장 공간(보통 5MB)을 넘은 경우
    if (!server.available && !storageWarned) {
      storageWarned = true;
      toast('브라우저 저장 공간이 부족해 자동 저장하지 못했습니다. [파일 → 다른 이름으로 저장]으로 파일을 내려받으세요.');
    }
    return false;
  }
}

function loadFromStorage() {
  try {
    const data = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null');
    return data?.workbook ? data : null;
  } catch {
    return null;
  }
}

let saveTimer = null;
let serverTimer = null;
function scheduleAutosave() {
  if (!autosave) { updateTitle(); return; }
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveToStorage, 400);
  scheduleServerSave();
}

function scheduleServerSave(delay = 1500) {
  if (!server.available || !autosave) return;
  clearTimeout(serverTimer);
  serverTimer = setTimeout(() => saveNow(false), delay);
}

/** 저장: 브라우저 + (서버가 있으면) 서버 */
async function saveNow(explicit) {
  const stored = saveToStorage();
  if (!server.available) {
    if (explicit && stored) toast('이 브라우저에 저장했습니다. (서버 없이 실행 중)');
    return;
  }
  serverState.saving = true;
  updateTitle();
  try {
    await server.save(docName, snapshot());
    serverState.error = null;
    serverState.savedAt = Date.now();
    dirty = false;
    if (explicit) toast('서버에 저장했습니다. 다른 기기에서도 열 수 있습니다.');
  } catch (err) {
    serverState.error = err.message;
    if (err.status === 401) askServerToken(() => saveNow(explicit));
    else if (explicit) toast(`서버에 저장하지 못했습니다: ${err.message}`);
  } finally {
    serverState.saving = false;
    updateTitle();
  }
}

function askServerToken(then) {
  if (isDialogOpen()) return;
  formDialog('서버 암호', [{ name: 't', label: '암호', type: 'password', value: '' }], ({ t }) => {
    server.setToken(t);
    then?.();
  }, { note: '이 서버는 TABULA_TOKEN 으로 보호되어 있습니다.' });
}

async function openFromServer(name) {
  try {
    const data = await server.load(name);
    loadWorkbook(data.workbook ?? data, data.docName ?? name, data.si ?? 0);
    dirty = false;
    serverState.savedAt = Date.now();
    updateTitle();
    toast(`'${name}'을(를) 열었습니다.`);
  } catch (err) {
    if (err.status === 401) askServerToken(() => openFromServer(name));
    else alertDialog('Tabula', `열 수 없습니다: ${err.message}`);
  }
}

function formatDate(ms) {
  const d = new Date(ms);
  const p = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
}

function openBackstage(panel = 'new') {
  const close = () => { stage.remove(); focusGrid(); };
  const main = el('div', { class: 'backstage-main' });
  const showNew = () => main.replaceChildren(
    el('h2', {}, '새로 만들기'),
    el('div', { class: 'backstage-cards' },
      el('button', { class: 'bcard', onclick: () => { close(); newWorkbook(); } }, el('div', { class: 'thumb' }), el('b', {}, '새 통합 문서'), el('small', {}, '빈 시트로 시작')),
      SAMPLES.map((s) => el('button', { class: 'bcard', onclick: () => { close(); newWorkbook(s); } }, el('div', { class: 'thumb' }), el('b', {}, s.name), el('small', {}, s.desc)))),
    el('h2', { style: { marginTop: '36px' } }, '정보'),
    el('div', { class: 'muted' }, `${docName} · 시트 ${wb.sheets.length}개 · ${server.available ? '서버에 저장 (다른 기기에서 열 수 있음)' : '이 브라우저에 저장'}${autosave ? ' · 자동 저장 켜짐' : ''}`),
  );
  const showOpen = async () => {
    const actions = el('div', { class: 'backstage-actions' },
      el('button', { class: 'btn primary', onclick: () => { close(); pickFile('open'); } }, '이 기기에서 찾아보기 (.xlsx, .csv, .tabula)'));
    main.replaceChildren(el('h2', {}, '열기'), actions);
    if (!server.available) {
      main.append(el('div', { class: 'backstage-note' },
        '서버 없이 실행 중이라 문서가 이 브라우저에만 저장됩니다. ',
        el('b', {}, 'npm start'), '로 실행하면 같은 네트워크의 다른 기기(휴대폰·다른 PC)에서도 같은 문서를 열 수 있습니다. ',
        '또는 [다른 이름으로 저장]에서 .xlsx 파일로 내려받아 옮기세요.'));
      return;
    }
    const loading = el('div', { class: 'muted' }, '불러오는 중...');
    main.append(loading);
    try {
      const files = await server.list();
      loading.remove();
      main.append(el('div', { class: 'backstage-note' }, `서버에 저장된 통합 문서 · 같은 네트워크의 다른 기기에서 접속해도 이 목록이 보입니다.`));
      if (!files.length) { main.append(el('div', { class: 'muted' }, '저장된 문서가 없습니다.')); return; }
      const openF = (f) => { close(); openFromServer(f.name); };
      main.append(el('table', { class: 'backstage-list' },
        el('tr', {}, el('th', {}, '이름'), el('th', {}, '수정한 날짜'), el('th', {}, '크기'), el('th', {}, '')),
        files.map((f) => el('tr', { class: 'file' },
          el('td', { onclick: () => openF(f) }, el('b', {}, f.name), f.name === docName ? el('span', { class: 'muted' }, ' (현재 문서)') : null),
          el('td', { onclick: () => openF(f) }, formatDate(f.modified)),
          el('td', {}, `${Math.max(1, Math.round(f.size / 1024))}KB`),
          el('td', {}, el('button', {
            class: 'btn', onclick: () => openDialog({
              title: '삭제', body: `'${f.name}'을(를) 서버에서 삭제할까요?`,
              buttons: [{ label: '삭제', primary: true, action: async () => { await server.remove(f.name); showOpen(); } }, { label: '취소' }],
            }),
          }, '삭제'))))));
    } catch (err) {
      loading.remove();
      if (err.status === 401) askServerToken(showOpen);
      main.append(el('div', { class: 'muted' }, `목록을 불러오지 못했습니다: ${err.message}`));
    }
  };
  const showShare = () => main.replaceChildren(
    el('h2', {}, '다른 기기에서 열기'),
    el('div', { class: 'backstage-note' }, ...(server.available
      ? ['이 문서는 서버에 자동으로 저장됩니다. 같은 네트워크의 다른 기기에서 아래 주소로 접속한 뒤 ', el('b', {}, '파일 → 열기'), `에서 '${docName}'을(를) 선택하세요.`]
      : ['서버 없이 실행 중입니다. 다른 기기로 옮기려면 .xlsx 파일로 저장해 전달하거나, npm start 로 서버를 실행하세요.'])),
    server.available ? el('div', {}, el('input', { type: 'text', value: location.href.replace(/[?#].*$/, ''), readonly: true, style: { width: '420px', height: '30px', padding: '0 8px' }, onclick: (e) => e.target.select() })) : null,
    el('div', { class: 'backstage-note' }, 'localhost 로 접속 중이면, 서버를 실행한 터미널에 표시된 "다른 기기에서" 주소를 사용하세요.'),
    el('div', { class: 'backstage-actions' }, el('button', { class: 'btn', onclick: () => { close(); exportXlsx(); } }, 'Excel 파일(.xlsx)로 내려받기')),
  );
  const nav = el('div', { class: 'backstage-nav' },
    el('button', { class: 'back', title: '돌아가기', onclick: close }, '←'),
    el('button', { onclick: showNew }, '새로 만들기'),
    el('button', { onclick: showOpen }, '열기'),
    el('button', { onclick: () => { close(); run('save'); } }, '저장'),
    el('button', { onclick: () => { close(); saveAs(); } }, '다른 이름으로 저장'),
    el('button', { onclick: () => { close(); exportXlsx(); } }, 'Excel(.xlsx)로 내보내기'),
    el('button', { onclick: () => { close(); exportCsv(); } }, 'CSV로 내보내기'),
    el('button', { onclick: showShare }, '다른 기기에서 열기'),
    el('button', { onclick: () => { close(); setTimeout(printSheet, 50); } }, '인쇄'),
    el('button', { onclick: close }, '닫기'));
  const stage = el('div', { class: 'backstage' }, nav, main);
  stage.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
  document.body.append(stage);
  if (panel === 'open') showOpen(); else showNew();
  nav.children[panel === 'open' ? 2 : 1].focus();
}

/** 인쇄: 사용한 범위를 표로 만들어 인쇄 */
function printSheet() {
  const s = sheet();
  const u = wb.usedRange(si);
  const rows = Math.min(u.rows, 5000);
  const cols = Math.min(u.cols, 100);
  const parts = [`<h2>${escapeHtml(docName)} — ${escapeHtml(s.name)}</h2>`];
  if (rows && cols) {
    const colgroup = range(0, cols - 1).map((c) => `<col style="width:${gv.cols.size(c)}px">`).join('');
    const body = [];
    for (let r = 0; r < rows; r++) {
      if (gv.rows.isHidden(r)) continue;
      const tds = [];
      for (let c = 0; c < cols; c++) {
        if (gv.cols.isHidden(c)) { tds.push('<td></td>'); continue; }
        const st = styleAt(r, c);
        const { text, align } = formatValue(valueAt(r, c), st);
        const css = [`text-align:${st.align || align}`, st.bold && 'font-weight:700', st.italic && 'font-style:italic', st.color && `color:${st.color}`,
          st.fill && `background:${st.fill}`, st.size && `font-size:${st.size}pt`, st.wrap && 'white-space:pre-wrap',
          st.bb && 'border-bottom:1px solid #000', st.bt && 'border-top:1px solid #000', st.bl && 'border-left:1px solid #000', st.br && 'border-right:1px solid #000'].filter(Boolean).join(';');
        tds.push(`<td style="${css}">${escapeHtml(text)}</td>`);
      }
      body.push(`<tr style="height:${gv.rows.size(r)}px">${tds.join('')}</tr>`);
    }
    parts.push(`<table class="${view.printGrid ? 'grid-lines' : ''}"><colgroup>${colgroup}</colgroup>${body.join('')}</table>`);
  }
  for (const ch of s.charts) parts.push(`<div class="chart-print">${gv.chartSvg(ch)}</div>`);
  for (const im of s.images ?? []) parts.push(`<div class="chart-print"><img src="${escapeHtml(im.src)}" style="width:${im.w}px;height:${im.h}px" alt=""></div>`);
  for (const sh of s.shapes ?? []) parts.push(`<div class="chart-print" style="position:relative;width:${sh.w}px;height:${Math.max(1, sh.h)}px">${shapeSvg(sh)}${sh.text ? `<div style="position:absolute;inset:0;display:flex;flex-direction:column;justify-content:center;padding:4px 8px;white-space:pre-wrap;text-align:${sh.align ?? 'center'};color:${escapeHtml(sh.color ?? '#000')};font-size:${sh.size ?? 11}pt">${escapeHtml(sh.text)}</div>` : ''}</div>`);
  dom.printArea.innerHTML = parts.join('');
  window.print();
}

// ───────────────────────── 대화상자 ─────────────────────────
function insertFunctionDialog() {
  const catSel = el('select', {}, el('option', { value: '' }, '모두'), CATEGORIES.map((c) => el('option', { value: c }, c)));
  const search = el('input', { type: 'text', placeholder: '함수 검색' });
  const list = el('div', { class: 'fn-list' });
  const desc = el('div', { class: 'fn-desc' });
  let chosen = 'SUM';
  const renderList = () => {
    const q = search.value.trim().toUpperCase();
    const names = FUNCTION_NAMES.filter((n) => (!catSel.value || FUNC_INFO[n]?.cat === catSel.value) && (!q || n.includes(q) || FUNC_INFO[n]?.desc.includes(search.value.trim())));
    if (!names.includes(chosen)) chosen = names[0];
    list.replaceChildren(...names.map((n) => el('div', {
      class: n === chosen ? 'active' : '',
      onclick: () => { chosen = n; renderList(); },
      ondblclick: () => { dlg.close(); insertFunctionText(n); },
    }, n)));
    list.querySelector('.active')?.scrollIntoView({ block: 'nearest' });
    const info = FUNC_INFO[chosen];
    desc.replaceChildren(...(chosen ? [el('b', {}, info?.sig ?? chosen), el('div', {}, info?.desc ?? '')] : []));
  };
  catSel.addEventListener('change', renderList);
  search.addEventListener('input', renderList);
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    el('label', {}, el('span', {}, '함수 검색'), search),
    el('label', {}, el('span', {}, '범주 선택'), catSel),
    list, desc);
  renderList();
  const dlg = openDialog({
    title: '함수 삽입', body, width: 460,
    buttons: [{ label: '확인', primary: true, action: () => { if (chosen) setTimeout(() => insertFunctionText(chosen)); } }, { label: '취소' }],
  });
}

const FRIENDLY_CODE = {
  date: () => 'yyyy-mm-dd', longdate: () => 'yyyy"년" m"월" d"일" aaaa', time: () => '[$-412]AM/PM h:mm:ss',
  datetime: () => 'yyyy-mm-dd h:mm', currency: (s) => `"₩"#,##0${s.decimals ? `.${'0'.repeat(s.decimals)}` : ''}`,
  accounting: (s) => buildCode('accounting', { decimals: s.decimals ?? 0, symbol: '₩' }),
  comma: (s) => `#,##0${s.decimals ? `.${'0'.repeat(s.decimals)}` : ''}`,
};

/** 셀 서식 (Ctrl+1): 표시 형식 · 맞춤 · 글꼴 · 테두리 · 채우기 */
function formatCellsDialog(startTab = 0) {
  if (editing && !commitEdit()) return;
  const st = styleAt(active.r, active.c);
  let sample = valueAt(active.r, active.c);
  if (sample === null || sample === '') {
    for (const [r, c] of cellsIn(usedClip(sel))) { const v = valueAt(r, c); if (v !== null && v !== '') { sample = v; break; } }
  }
  const col = (...k) => el('div', { class: 'fc-col' }, ...k);
  const row = (...k) => el('div', { class: 'fc-row' }, ...k);
  const lab = (text, input) => el('label', { class: 'fc-field' }, el('span', {}, text), input);
  const chk = (text, checked) => { const i = el('input', { type: 'checkbox', checked: !!checked }); return [i, el('label', { class: 'fc-check' }, i, text)]; };

  // ── 표시 형식 ──
  const startCode = FRIENDLY_CODE[st.numFmt]?.(st) ?? codeOfStyle(st);
  const init = describeCode(startCode);
  const opts = { decimals: init.decimals ?? 2, thousands: init.thousands ?? true, negative: init.negative ?? 'minus', symbol: init.symbol ?? '₩', type: init.type };
  let cat = init.cat;
  const catList = el('div', { class: 'fc-cats', role: 'listbox' });
  const optBox = el('div', { class: 'fc-opts' });
  const sampleBox = el('div', { class: 'fc-sample' });
  const noteBox = el('div', { class: 'muted fc-note' });
  const codeIn = el('input', { type: 'text', class: 'fc-code', spellcheck: false });
  let customCode = init.cat === 'custom' ? init.type : startCode;
  const currentCode = () => (cat === 'custom' ? (codeIn.value.trim() || 'General') : buildCode(cat, opts));
  const showSample = () => {
    const code = currentCode();
    let text = '';
    let color = null;
    try {
      if (sample !== null && sample !== '' && typeof sample !== 'object') ({ text, color } = cat === 'general' ? formatValue(sample, {}) : formatCode(sample, code === 'G/표준' ? 'General' : code));
      else if (typeof sample === 'object' && sample) text = sample.code;
      sampleBox.classList.remove('bad');
    } catch {
      text = '서식 코드를 해석할 수 없습니다.';
      sampleBox.classList.add('bad');
    }
    sampleBox.replaceChildren(el('span', { class: 'muted' }, '보기'), el('b', { style: { color: color ?? '' } }, text || ' '));
  };
  const decimalsIn = () => {
    const i = el('input', { type: 'number', min: 0, max: 30, value: opts.decimals, style: { width: '64px' } });
    i.addEventListener('input', () => { opts.decimals = Number(i.value) || 0; showSample(); });
    return lab('소수 자릿수', i);
  };
  const typeList = (items, labelOf = (x) => x.label ?? x, codeOf = (x) => x.code ?? x) => {
    const box = el('div', { class: 'fc-list' });
    if (!items.some((x) => codeOf(x) === opts.type)) opts.type = codeOf(items[0]);
    for (const it of items) {
      const code = codeOf(it);
      const b = el('button', { type: 'button', class: `fc-item${code === opts.type ? ' on' : ''}` }, labelOf(it));
      b.addEventListener('click', () => {
        opts.type = code;
        box.querySelectorAll('.on').forEach((x) => x.classList.remove('on'));
        b.classList.add('on');
        showSample();
      });
      box.append(b);
    }
    return box;
  };
  const sampleOf = (code) => {
    const n = cat === 'date' || cat === 'time' ? (typeof sample === 'number' ? sample : 45366.5625) : typeof sample === 'number' ? sample : 1234.5;
    try { return formatCode(n, code).text; } catch { return code; }
  };
  const negList = () => {
    const box = el('div', { class: 'fc-list short' });
    for (const n of NEGATIVE_STYLES) {
      const b = el('button', { type: 'button', class: `fc-item${n.id === opts.negative ? ' on' : ''}`, style: { color: n.red ? '#c00000' : '' } }, n.label);
      b.addEventListener('click', () => { opts.negative = n.id; box.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); b.classList.add('on'); showSample(); });
      box.append(b);
    }
    return lab('음수', box);
  };
  const symbolSel = () => {
    const s = el('select', {}, CURRENCY_SYMBOLS.map((x) => el('option', { value: x.id, selected: x.id === opts.symbol }, x.label)));
    s.addEventListener('change', () => { opts.symbol = s.value; showSample(); });
    return lab('기호', s);
  };
  const renderOpts = () => {
    optBox.replaceChildren();
    const c = FMT_CATEGORIES.find((x) => x.id === cat);
    noteBox.textContent = c.note ?? '';
    if (cat === 'number') {
      const [t, tl] = chk('1000 단위 구분 기호(,) 사용', opts.thousands);
      t.addEventListener('change', () => { opts.thousands = t.checked; showSample(); });
      optBox.append(decimalsIn(), tl, negList());
    } else if (cat === 'currency') optBox.append(decimalsIn(), symbolSel(), negList());
    else if (cat === 'accounting') optBox.append(decimalsIn(), symbolSel());
    else if (cat === 'percent' || cat === 'scientific') optBox.append(decimalsIn());
    else if (cat === 'date') optBox.append(lab('형식', typeList(DATE_TYPES, (x) => sampleOf(x))));
    else if (cat === 'time') optBox.append(lab('형식', typeList(TIME_TYPES, (x) => sampleOf(x))));
    else if (cat === 'fraction') optBox.append(lab('형식', typeList(FRACTION_TYPES)));
    else if (cat === 'special') optBox.append(lab('형식', typeList(SPECIAL_TYPES)));
    else if (cat === 'custom') {
      codeIn.value = customCode;
      const used = new Set();
      for (const s of wb.sheets) for (const cell of s.cells.values()) if (cell.style?.numFmt === 'custom' && cell.style.code) used.add(cell.style.code);
      const list = el('div', { class: 'fc-list tall' });
      for (const code of [...CUSTOM_LIST, ...[...used].filter((x) => !CUSTOM_LIST.includes(x))]) {
        const b = el('button', { type: 'button', class: `fc-item mono${code === codeIn.value ? ' on' : ''}` }, code);
        b.addEventListener('click', () => { codeIn.value = code; customCode = code; list.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); b.classList.add('on'); showSample(); });
        list.append(b);
      }
      codeIn.oninput = () => { customCode = codeIn.value; showSample(); };
      optBox.append(lab('형식', codeIn), list, el('div', { class: 'muted fc-help' },
        '0 # ? 숫자 자리 · , 천 단위 · % 백분율 · "글자" · @ 텍스트 · [빨강] 색 · [>=100] 조건 · 양수;음수;0;텍스트 · yyyy mm dd aaaa h:mm:ss AM/PM [h]'));
      setTimeout(() => codeIn.focus());
    }
    showSample();
  };
  for (const c of FMT_CATEGORIES) {
    const b = el('button', { type: 'button', class: `fc-cat${c.id === cat ? ' on' : ''}` }, c.label);
    b.addEventListener('click', () => {
      if (cat !== 'custom' && c.id === 'custom') customCode = currentCode() === 'General' ? 'G/표준' : currentCode();
      cat = c.id;
      catList.querySelectorAll('.on').forEach((x) => x.classList.remove('on'));
      b.classList.add('on');
      renderOpts();
    });
    catList.append(b);
  }
  renderOpts();
  const numberPage = el('div', { class: 'fc-number' }, catList, col(sampleBox, optBox, noteBox));

  // ── 맞춤 ──
  const hSel = el('select', {}, [['', '일반'], ['left', '왼쪽'], ['center', '가운데'], ['right', '오른쪽']].map(([v, l]) => el('option', { value: v, selected: (st.align ?? '') === v }, l)));
  const vSel = el('select', {}, [['top', '위쪽'], ['middle', '가운데'], ['', '아래쪽']].map(([v, l]) => el('option', { value: v, selected: (st.valign ?? '') === v }, l)));
  const indentIn = el('input', { type: 'number', min: 0, max: 15, value: st.indent ?? 0, style: { width: '64px' } });
  const [wrapIn, wrapL] = chk('텍스트 줄 바꿈', st.wrap);
  const merged = !!wb.mergeAt(si, active.r, active.c);
  const [mergeIn, mergeL] = chk('셀 병합', merged);
  const alignPage = col(el('div', { class: 'fc-title' }, '텍스트 맞춤'), row(lab('가로', hSel), lab('세로', vSel), lab('들여쓰기', indentIn)),
    el('div', { class: 'fc-title' }, '텍스트 조정'), wrapL, mergeL);

  // ── 글꼴 ──
  const fontSel = el('select', {}, FONTS.map((f) => el('option', { value: f, selected: (st.font || DEFAULT_FONT) === f }, f)));
  const sizeIn = el('input', { type: 'number', min: 1, max: 409, value: st.size || DEFAULT_SIZE, style: { width: '64px' } });
  const [bIn, bL] = chk('굵게', st.bold);
  const [iIn, iL] = chk('기울임꼴', st.italic);
  const [uIn, uL] = chk('밑줄', st.underline);
  const [sIn, sL] = chk('취소선', st.strike);
  const colorIn = el('input', { type: 'color', value: st.color || '#000000' });
  const fontPreview = el('div', { class: 'fc-fontprev' }, '가나다 AaBbCc 123');
  const updFont = () => Object.assign(fontPreview.style, {
    fontFamily: fontStack(fontSel.value), fontSize: `${Math.min(Number(sizeIn.value) || 11, 36)}pt`, fontWeight: bIn.checked ? 700 : 400,
    fontStyle: iIn.checked ? 'italic' : 'normal', textDecoration: `${uIn.checked ? 'underline ' : ''}${sIn.checked ? 'line-through' : ''}`, color: colorIn.value,
  });
  [fontSel, sizeIn, bIn, iIn, uIn, sIn, colorIn].forEach((x) => x.addEventListener('input', updFont));
  updFont();
  const fontPage = col(row(lab('글꼴', fontSel), lab('크기', sizeIn)), row(bL, iL, uL, sL), lab('색', colorIn), el('div', { class: 'fc-title' }, '미리 보기'), fontPreview);

  // ── 테두리 ──
  let border = null;
  const borderBtns = el('div', { class: 'fc-row' });
  for (const [k, label, ic] of [['none', '없음', 'borderNone'], ['outside', '윤곽선', 'borderOutside'], ['all', '모든 테두리', 'borderAll'], ['bottom', '아래쪽', 'borderBottom'], ['topBottom', '위쪽/아래쪽', 'borderTop']]) {
    const b = el('button', { type: 'button', class: 'fc-bbtn', title: label, html: `${ICONS[ic] ?? ''}<span>${label}</span>` });
    b.addEventListener('click', () => { border = k; borderBtns.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); b.classList.add('on'); });
    borderBtns.append(b);
  }
  const borderPage = col(el('div', { class: 'fc-title' }, '미리 설정'), borderBtns, el('div', { class: 'muted' }, '선택한 범위에 적용됩니다.'));

  // ── 채우기 ──
  const [noFill, noFillL] = chk('채우기 없음', !st.fill);
  const fillIn = el('input', { type: 'color', value: st.fill || '#ffff00' });
  const swatches = el('div', { class: 'fc-swatches' });
  for (const c of ['#ffffff', '#f2f2f2', '#d9d9d9', '#fff2cc', '#fce4d6', '#e2efda', '#ddebf7', '#ededed', '#ffff00', '#ffc000', '#92d050', '#00b0f0', '#ff0000', '#7030a0', '#4472c4', '#70ad47']) {
    const b = el('button', { type: 'button', class: 'fc-sw', title: c, style: { background: c } });
    b.addEventListener('click', () => { fillIn.value = c; noFill.checked = false; });
    swatches.append(b);
  }
  fillIn.addEventListener('input', () => { noFill.checked = false; });
  const fillPage = col(noFillL, el('div', { class: 'fc-title' }, '배경색'), swatches, lab('다른 색', fillIn));

  const pages = [['표시 형식', numberPage], ['맞춤', alignPage], ['글꼴', fontPage], ['테두리', borderPage], ['채우기', fillPage]];
  const tabBar = el('div', { class: 'dlg-tabs' });
  const pageBox = el('div', { class: 'fc-page' });
  const show = (i) => { [...tabBar.children].forEach((b, j) => b.classList.toggle('on', i === j)); pageBox.replaceChildren(pages[i][1]); };
  pages.forEach(([name], i) => tabBar.append(el('button', { type: 'button', class: 'dlg-tab', onclick: () => show(i) }, name)));
  show(startTab);

  openDialog({
    title: '셀 서식', width: 620, body: el('div', {}, tabBar, pageBox),
    buttons: [
      {
        label: '확인', primary: true,
        action: () => {
          const code = currentCode();
          let fmt;
          try {
            formatCode(1234.5, code);
            fmt = cat === 'general' ? { numFmt: undefined, decimals: undefined, code: undefined } : styleForCode(code);
          } catch {
            show(0);
            alertDialog('셀 서식', '입력한 서식 코드를 사용할 수 없습니다. 코드를 확인하세요.');
            return false;
          }
          const size = Number(sizeIn.value);
          const patch = {
            ...fmt,
            align: hSel.value || undefined, valign: vSel.value || undefined, indent: Number(indentIn.value) || undefined, wrap: wrapIn.checked || undefined,
            font: fontSel.value === DEFAULT_FONT ? undefined : fontSel.value,
            size: !size || size === DEFAULT_SIZE ? undefined : Math.min(409, size),
            bold: bIn.checked || undefined, italic: iIn.checked || undefined, underline: uIn.checked || undefined, strike: sIn.checked || undefined,
            color: colorIn.value === '#000000' ? undefined : colorIn.value,
            fill: noFill.checked ? undefined : fillIn.value,
          };
          wb.transact(() => {
            applyStyle(patch, { widen: fmt.numFmt !== undefined ? 'grow' : false });
            if (border) applyBorder(border);
            if (mergeIn.checked !== merged && selKind === 'cells' && !isSingle(sel)) {
              if (mergeIn.checked) wb.merge(si, sel.r1, sel.c1, sel.r2, sel.c2); else wb.unmerge(si, sel.r1, sel.c1, sel.r2, sel.c2);
            } else if (!mergeIn.checked && merged) {
              const m = wb.mergeAt(si, active.r, active.c);
              if (m) wb.unmerge(si, m.r1, m.c1, m.r2, m.c2);
            }
          }, meta());
          return undefined;
        },
      },
      { label: '취소' },
    ],
  });
}

function sizeDialog(kind) {
  const isCol = kind === 'col';
  const cur = isCol ? wb.colWidth(si, active.c) : wb.rowHeight(si, active.r);
  formDialog(isCol ? '열 너비' : '행 높이', [{ name: 'v', label: isCol ? '열 너비(px)' : '행 높이(px)', type: 'number', value: cur }], ({ v }) => {
    const n = Number(v);
    if (!(n >= 0 && n <= 1000)) { toast('0에서 1000 사이의 값을 입력하세요.'); return false; }
    wb.transact(() => {
      if (isCol) for (let c = sel.c1; c <= Math.min(sel.c2, sel.c1 + 1000); c++) wb.setColWidth(si, c, n);
      else for (let r = sel.r1; r <= Math.min(sel.r2, sel.r1 + 5000); r++) wb.setRowHeight(si, r, n);
    }, meta());
    return true;
  });
}

function sortDialog() {
  const rg = dataRange();
  const header = hasHeader(rg);
  const options = [];
  for (let c = rg.c1; c <= rg.c2; c++) {
    const hv = header ? displayText(rg.r1, c) : '';
    options.push({ value: String(c), label: hv ? `${hv} (${colToName(c)}열)` : `${colToName(c)}열` });
  }
  formDialog('정렬', [
    { name: 'col', label: '정렬 기준', type: 'select', value: String(clamp(active.c, rg.c1, rg.c2)), options },
    { name: 'order', label: '정렬', type: 'select', value: 'asc', options: [{ value: 'asc', label: '오름차순' }, { value: 'desc', label: '내림차순' }] },
    { name: 'header', label: '머리글 포함', type: 'checkbox', value: header },
  ], (v) => {
    sortData(v.order === 'asc', Number(v.col), v.header, rg);
  }, { note: `범위: ${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}` });
}

function gotoDialog() {
  formDialog('이동', [{ name: 'ref', label: '참조', value: '' }], ({ ref }) => gotoRef(ref));
}

function gotoRef(text) {
  let t = text.trim();
  let target = si;
  const bang = t.lastIndexOf('!');
  if (bang > 0) {
    const name = t.slice(0, bang).replace(/^'(.*)'$/, '$1');
    target = wb.sheetIndexByName(name);
    t = t.slice(bang + 1);
    if (target < 0) { toast('시트를 찾을 수 없습니다.'); return false; }
  }
  const colOnly = /^\$?([A-Za-z]{1,3}):\$?([A-Za-z]{1,3})$/.exec(t);
  const rowOnly = /^\$?(\d+):\$?(\d+)$/.exec(t);
  const rg = parseRangeName(t);
  if (!colOnly && !rowOnly && !rg) { toast('참조가 올바르지 않습니다.'); return false; }
  switchSheet(target);
  if (colOnly) {
    const a = parseRangeName(`${colOnly[1]}1`);
    const b = parseRangeName(`${colOnly[2]}1`);
    if (a && b) { growTo(0, Math.max(a.c1, b.c1)); gv.ensureVisible(gv.firstVisibleRow(), a.c1); selectCols(a.c1, b.c1); return true; }
  }
  if (rowOnly) {
    const a = Number(rowOnly[1]) - 1;
    const b = Number(rowOnly[2]) - 1;
    if (a >= 0 && b >= 0 && b < MAX_ROWS) { growTo(Math.max(a, b), 0); gv.ensureVisible(a, 0); selectRows(a, b); return true; }
  }
  if (!rg) return false;
  if (isSingle(rg)) selectCell(rg.r1, rg.c1);
  else { growTo(rg.r2, rg.c2); gv.ensureVisible(rg.r1, rg.c1); selectRange(rg); }
  return true;
}

function condRuleDialog(type) {
  const titles = {
    gt: ['보다 큼', '다음 값보다 큰 셀의 서식 지정:'], lt: ['보다 작음', '다음 값보다 작은 셀의 서식 지정:'],
    between: ['다음 값의 사이에 있음', '다음 값 사이에 있는 셀의 서식 지정:'], eq: ['같음', '다음 값과 같은 셀의 서식 지정:'],
    text: ['텍스트 포함', '다음 텍스트를 포함하는 셀의 서식 지정:'], dup: ['중복 값', '다음 값을 포함하는 셀의 서식 지정:'],
    top: ['상위 10개 항목', '가장 큰 값 순위에 해당하는 셀의 서식 지정:'],
  };
  const presets = [
    { value: '0', label: '진한 빨강 텍스트가 있는 연한 빨강 채우기', style: { fill: '#ffc7ce', color: '#9c0006' } },
    { value: '1', label: '진한 노랑 텍스트가 있는 노랑 채우기', style: { fill: '#ffeb9c', color: '#9c5700' } },
    { value: '2', label: '진한 녹색 텍스트가 있는 녹색 채우기', style: { fill: '#c6efce', color: '#006100' } },
    { value: '3', label: '연한 빨강 채우기', style: { fill: '#ffc7ce' } },
    { value: '4', label: '빨강 텍스트', style: { color: '#9c0006' } },
    { value: '5', label: '굵은 빨강 텍스트', style: { bold: true, color: '#c00000' } },
  ];
  const [title, note] = titles[type];
  const rg = usedClip(sel);
  const cur = displayText(active.r, active.c);
  const fields = [];
  if (type === 'dup') fields.push({ name: 'mode', label: '값', type: 'select', value: 'dup', options: [{ value: 'dup', label: '중복' }, { value: 'unique', label: '고유' }] });
  else if (type === 'top') fields.push({ name: 'v1', label: '상위 개수', type: 'number', value: 10 });
  else {
    fields.push({ name: 'v1', label: type === 'between' ? '최소값' : '값', value: cur });
    if (type === 'between') fields.push({ name: 'v2', label: '최대값', value: '' });
  }
  fields.push({ name: 'preset', label: '적용할 서식', type: 'select', value: '0', options: presets });
  formDialog(title, fields, (v) => {
    const rule = { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2, type: type === 'dup' ? v.mode : type, v1: v.v1, v2: v.v2, style: presets[Number(v.preset)].style };
    wb.transact(() => wb.addCondRule(si, rule), meta());
  }, { note });
}

// ───────────────────────── 조건부 서식 규칙 관리자 ─────────────────────────
const CF_PRESETS = [
  { label: '진한 빨강 텍스트가 있는 연한 빨강 채우기', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { label: '진한 노랑 텍스트가 있는 노랑 채우기', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { label: '진한 녹색 텍스트가 있는 녹색 채우기', style: { fill: '#c6efce', color: '#006100' } },
  { label: '연한 빨강 채우기', style: { fill: '#ffc7ce' } },
  { label: '빨강 텍스트', style: { color: '#9c0006' } },
  { label: '굵은 빨강 텍스트', style: { bold: true, color: '#c00000' } },
  { label: '빨강 테두리', style: { bt: true, bb: true, bl: true, br: true, color: '#c00000' } },
];

const CF_KINDS = [
  { id: 'visual', label: '셀 값을 기준으로 모든 셀의 서식 지정' },
  { id: 'contains', label: '다음을 포함하는 셀만 서식 지정' },
  { id: 'topbottom', label: '상위 또는 하위 값만 서식 지정' },
  { id: 'avg', label: '평균보다 크거나 작은 값만 서식 지정' },
  { id: 'dupuniq', label: '고유 또는 중복 값만 서식 지정' },
  { id: 'formula', label: '수식을 사용하여 서식을 지정할 셀 결정' },
];

function cfKindOf(rule) {
  const t = rule.type;
  if (VISUAL_TYPES.has(t)) return 'visual';
  if (t === 'top' || t === 'bottom') return 'topbottom';
  if (t === 'aboveAvg' || t === 'belowAvg') return 'avg';
  if (t === 'dup' || t === 'unique') return 'dupuniq';
  if (t === 'formula') return 'formula';
  return 'contains';
}

/** 규칙 미리 보기 (목록과 편집 창) */
function cfPreview(rule) {
  if (rule.type === 'bar') return el('span', { class: 'cf-prev', style: { background: `linear-gradient(90deg, ${rule.color ?? '#638ec6'} 70%, transparent 70%)` } });
  if (rule.type === 'scale') return el('span', { class: 'cf-prev', style: { background: `linear-gradient(90deg, ${(rule.colors ?? ['#f8696b', '#63be7b']).join(',')})` } });
  if (rule.type === 'icons') return el('span', { class: 'cf-prev icons', html: iconSetById(rule.icons).icons.map((i) => ICON_SVG[i]).join('') });
  const st = rule.style ?? {};
  return el('span', {
    class: 'cf-prev text',
    style: {
      background: st.fill ?? '#fff', color: st.color ?? '#000', fontWeight: st.bold ? 700 : 400, fontStyle: st.italic ? 'italic' : 'normal',
      textDecoration: `${st.underline ? 'underline ' : ''}${st.strike ? 'line-through' : ''}`, border: st.bt || st.bb ? '1px solid #000' : '1px solid #ddd',
    },
  }, 'AaBbCcYyZz');
}

/** 새 서식 규칙 / 서식 규칙 편집 */
function cfRuleEditor(initial, onSave, { title = '새 서식 규칙' } = {}) {
  const rule = structuredClone(initial ?? { type: 'gt', v1: '', style: { ...CF_PRESETS[0].style } });
  let kind = cfKindOf(rule);
  const body = el('div', { class: 'cf-editor' });
  const kindList = el('div', { class: 'fc-list cf-kinds' });
  const detail = el('div', { class: 'cf-detail' });
  const fmtBox = el('div', { class: 'cf-fmt' });
  const sel = (options, value, onChange) => {
    const s = el('select', {}, options.map(([v, l]) => el('option', { value: v, selected: v === value }, l)));
    s.addEventListener('change', () => onChange(s.value));
    return s;
  };
  const input = (value, onChange, attrs = {}) => {
    const i = el('input', { type: 'text', value: value ?? '', ...attrs });
    i.addEventListener('input', () => onChange(i.value));
    return i;
  };
  const colorIn = (value, onChange) => {
    const i = el('input', { type: 'color', value });
    i.addEventListener('input', () => onChange(i.value));
    return i;
  };
  const chk = (text, checked, onChange) => {
    const i = el('input', { type: 'checkbox', checked: !!checked });
    i.addEventListener('change', () => onChange(i.checked));
    return el('label', { class: 'fc-check' }, i, text);
  };
  const row = (...k) => el('div', { class: 'fc-row' }, ...k);

  const setKind = (k) => {
    kind = k;
    const keep = rule.style ?? { ...CF_PRESETS[0].style };
    const defaults = {
      visual: { type: 'scale', colors: ['#f8696b', '#ffeb84', '#63be7b'] }, contains: { type: 'gt', v1: '' },
      topbottom: { type: 'top', v1: '10' }, avg: { type: 'aboveAvg' }, dupuniq: { type: 'dup' }, formula: { type: 'formula', formula: '=' },
    };
    for (const key of ['type', 'v1', 'v2', 'color', 'colors', 'icons', 'reverse', 'iconOnly', 'percent', 'period', 'formula']) delete rule[key];
    Object.assign(rule, defaults[k], { style: keep });
    render();
  };

  const renderFmt = () => {
    fmtBox.replaceChildren();
    if (kind === 'visual') return;
    const st = rule.style ?? (rule.style = {});
    const prev = el('div', { class: 'cf-fmtprev' }, cfPreview(rule));
    const upd = () => prev.replaceChildren(cfPreview(rule));
    const preset = sel([['', '미리 설정된 서식...'], ...CF_PRESETS.map((p, i) => [String(i), p.label])], '', (v) => {
      if (v === '') return;
      rule.style = { ...CF_PRESETS[Number(v)].style };
      renderFmt();
    });
    const colorOn = !!st.color;
    const fillOn = !!st.fill;
    fmtBox.append(
      el('div', { class: 'fc-title' }, '서식'),
      row(prev, preset),
      row(
        chk('글꼴 색', colorOn, (on) => { st.color = on ? (st.color ?? '#9c0006') : undefined; renderFmt(); }),
        colorOn ? colorIn(st.color, (v) => { st.color = v; upd(); }) : null,
        chk('채우기', fillOn, (on) => { st.fill = on ? (st.fill ?? '#ffc7ce') : undefined; renderFmt(); }),
        fillOn ? colorIn(st.fill, (v) => { st.fill = v; upd(); }) : null,
      ),
      row(
        chk('굵게', st.bold, (on) => { st.bold = on || undefined; upd(); }),
        chk('기울임꼴', st.italic, (on) => { st.italic = on || undefined; upd(); }),
        chk('밑줄', st.underline, (on) => { st.underline = on || undefined; upd(); }),
        chk('취소선', st.strike, (on) => { st.strike = on || undefined; upd(); }),
        chk('테두리', st.bt && st.bb, (on) => { for (const b of ['bt', 'bb', 'bl', 'br']) st[b] = on || undefined; upd(); }),
      ),
    );
  };

  const render = () => {
    kindList.querySelectorAll('.fc-item').forEach((b) => b.classList.toggle('on', b.dataset.k === kind));
    detail.replaceChildren(el('div', { class: 'fc-title' }, '규칙 설명 편집'));
    if (kind === 'visual') {
      const style = sel([['scale2', '2가지 색조'], ['scale3', '3가지 색조'], ['bar', '데이터 막대'], ['icons', '아이콘 집합']],
        rule.type === 'scale' ? (rule.colors?.length === 3 ? 'scale3' : 'scale2') : rule.type, (v) => {
          for (const key of ['color', 'colors', 'icons', 'reverse', 'iconOnly']) delete rule[key];
          if (v === 'scale2') Object.assign(rule, { type: 'scale', colors: ['#fcfcff', '#63be7b'] });
          else if (v === 'scale3') Object.assign(rule, { type: 'scale', colors: ['#f8696b', '#ffeb84', '#63be7b'] });
          else if (v === 'bar') Object.assign(rule, { type: 'bar', color: '#638ec6' });
          else Object.assign(rule, { type: 'icons', icons: '3Arrows' });
          render();
        });
      detail.append(row(el('span', {}, '서식 스타일:'), style));
      if (rule.type === 'scale') {
        const names = rule.colors.length === 3 ? ['최소값', '중간값', '최대값'] : ['최소값', '최대값'];
        detail.append(row(...rule.colors.map((c, i) => el('label', { class: 'fc-field' }, el('span', {}, names[i]), colorIn(c, (v) => { rule.colors[i] = v; prevBox.replaceChildren(cfPreview(rule)); })))));
      } else if (rule.type === 'bar') {
        detail.append(row(el('span', {}, '막대 색:'), colorIn(rule.color ?? '#638ec6', (v) => { rule.color = v; prevBox.replaceChildren(cfPreview(rule)); }), chk('막대만 표시', rule.iconOnly, (on) => { rule.iconOnly = on || undefined; })));
      } else {
        detail.append(
          row(el('span', {}, '아이콘 스타일:'), sel(ICON_SETS.map((s) => [s.id, s.label]), rule.icons, (v) => { rule.icons = v; render(); })),
          row(chk('아이콘 순서 거꾸로', rule.reverse, (on) => { rule.reverse = on || undefined; render(); }), chk('아이콘만 표시', rule.iconOnly, (on) => { rule.iconOnly = on || undefined; })),
          el('div', { class: 'muted' }, `값 범위를 ${iconSetById(rule.icons).icons.length}등분해 낮은 값부터 아이콘을 붙입니다.`),
        );
      }
      const prevBox = el('div', { class: 'cf-fmtprev' }, cfPreview(rule));
      detail.append(el('div', { class: 'fc-title' }, '미리 보기'), prevBox);
    } else if (kind === 'contains') {
      const group = CELL_OPS.some((o) => o.id === rule.type) ? 'cell' : TEXT_OPS.some((o) => o.id === rule.type) ? 'text'
        : rule.type === 'date' ? 'date' : rule.type;
      const first = sel([['cell', '셀 값'], ['text', '특정 텍스트'], ['date', '발생 날짜'], ['blank', '빈 셀'], ['noBlank', '내용 있는 셀'], ['errors', '오류'], ['noErrors', '오류 없음']], group, (v) => {
        delete rule.v1; delete rule.v2; delete rule.period;
        rule.type = { cell: 'gt', text: 'text', date: 'date' }[v] ?? v;
        if (v === 'date') rule.period = 'today';
        render();
      });
      const parts = [first];
      if (group === 'cell') {
        const op = CELL_OPS.find((o) => o.id === rule.type);
        parts.push(sel(CELL_OPS.map((o) => [o.id, o.label]), rule.type, (v) => { rule.type = v; render(); }));
        parts.push(input(rule.v1, (v) => { rule.v1 = v; }, { placeholder: '값 또는 =수식', style: { width: '110px' } }));
        if (op?.two) parts.push(el('span', {}, '및'), input(rule.v2, (v) => { rule.v2 = v; }, { placeholder: '값 또는 =수식', style: { width: '110px' } }));
      } else if (group === 'text') {
        parts.push(sel(TEXT_OPS.map((o) => [o.id, o.label]), rule.type, (v) => { rule.type = v; }));
        parts.push(input(rule.v1, (v) => { rule.v1 = v; }, { style: { width: '140px' } }));
      } else if (group === 'date') {
        parts.push(sel(DATE_PERIODS.map((p) => [p.id, p.label]), rule.period ?? 'today', (v) => { rule.period = v; }));
      }
      detail.append(el('div', { class: 'muted' }, '다음 조건에 맞는 셀만 서식 지정:'), row(...parts));
    } else if (kind === 'topbottom') {
      detail.append(el('div', { class: 'muted' }, '다음 순위에 해당하는 값의 서식 지정:'), row(
        sel([['top', '상위'], ['bottom', '하위']], rule.type, (v) => { rule.type = v; }),
        input(rule.v1 ?? '10', (v) => { rule.v1 = v; }, { type: 'number', min: 1, style: { width: '70px' } }),
        chk('선택한 범위의 %', rule.percent, (on) => { rule.percent = on || undefined; }),
      ));
    } else if (kind === 'avg') {
      detail.append(el('div', { class: 'muted' }, '선택한 범위의 평균 값을 기준으로 다음 값의 서식 지정:'), row(sel([['aboveAvg', '초과'], ['belowAvg', '미만']], rule.type, (v) => { rule.type = v; })));
    } else if (kind === 'dupuniq') {
      detail.append(el('div', { class: 'muted' }, '다음 값의 서식 지정:'), row(sel([['dup', '중복'], ['unique', '고유']], rule.type, (v) => { rule.type = v; })));
    } else {
      detail.append(el('div', { class: 'muted' }, '다음 수식이 참인 값의 서식 지정 (범위의 왼쪽 위 셀 기준으로 입력):'),
        input(rule.formula ?? '=', (v) => { rule.formula = v; }, { class: 'fc-code', placeholder: '=$B2>100' }));
    }
    renderFmt();
  };

  for (const k of CF_KINDS) {
    const b = el('button', { type: 'button', class: 'fc-item', 'data-k': k.id }, `► ${k.label}`);
    b.addEventListener('click', () => { if (k.id !== kind) setKind(k.id); });
    kindList.append(b);
  }
  body.append(el('div', { class: 'fc-title' }, '규칙 유형 선택'), kindList, detail, fmtBox);
  render();
  openDialog({
    title, width: 560, body,
    buttons: [
      {
        label: '확인', primary: true,
        action: () => {
          if (rule.type === 'formula') {
            const f = String(rule.formula ?? '').trim();
            try { if (!f.startsWith('=') || f.length < 2) throw new Error(); parse(f.slice(1)); } catch {
              alertDialog('조건부 서식', '수식이 올바르지 않습니다. = 로 시작하는 수식을 입력하세요.');
              return false;
            }
            rule.formula = normalizeFormula(f);
          }
          const op = CELL_OPS.find((o) => o.id === rule.type);
          if (op && (String(rule.v1 ?? '').trim() === '' || (op.two && String(rule.v2 ?? '').trim() === ''))) {
            alertDialog('조건부 서식', '비교할 값을 입력하세요.');
            return false;
          }
          if (VISUAL_TYPES.has(rule.type)) { delete rule.style; delete rule.stopIfTrue; }
          onSave(rule);
          return undefined;
        },
      },
      { label: '취소' },
    ],
  });
}

const rangeText = (r) => `=$${colToName(r.c1)}$${r.r1 + 1}${r.r1 === r.r2 && r.c1 === r.c2 ? '' : `:$${colToName(r.c2)}$${r.r2 + 1}`}`;

function cfManager() {
  if (editing && !commitEdit()) return;
  const lists = new Map(); // 시트 → 편집 중인 규칙 목록 (확인/적용 때 반영)
  const listOf = (i) => { if (!lists.has(i)) lists.set(i, wb.sheets[i].cond.map((x) => structuredClone(x))); return lists.get(i); };
  const selRange = usedClip(sel);
  let scope = 'selection';
  let scopeSheet = si;
  let current = null;
  const showSel = el('select', {}, [
    el('option', { value: 'selection' }, '현재 선택 영역'),
    el('option', { value: `sheet:${si}` }, '현재 워크시트'),
    ...wb.sheets.map((s, i) => (i === si ? null : el('option', { value: `sheet:${i}` }, `시트: ${s.name}`))).filter(Boolean),
  ]);
  const tbody = el('tbody');
  const visible = () => {
    const list = listOf(scopeSheet);
    return list.filter((rl) => scope !== 'selection' || (rl.r1 <= selRange.r2 && rl.r2 >= selRange.r1 && rl.c1 <= selRange.c2 && rl.c2 >= selRange.c1));
  };
  const render = () => {
    tbody.replaceChildren();
    const vis = visible();
    if (!vis.includes(current)) current = vis[0] ?? null;
    if (!vis.length) tbody.append(el('tr', {}, el('td', { colspan: 4, class: 'muted', style: { padding: '12px' } }, '표시할 규칙이 없습니다. [새 규칙]을 눌러 만드세요.')));
    for (const rl of vis) {
      const rangeIn = el('input', { type: 'text', value: rangeText(rl), class: 'cf-range' });
      rangeIn.addEventListener('change', () => {
        const p = parseRangeName(rangeIn.value.replace(/[=$]/g, '').trim());
        if (!p) { rangeIn.classList.add('bad'); return; }
        rangeIn.classList.remove('bad');
        Object.assign(rl, { r1: p.r1, c1: p.c1, r2: p.r2, c2: p.c2 });
      });
      const stop = el('input', { type: 'checkbox', checked: !!rl.stopIfTrue, disabled: VISUAL_TYPES.has(rl.type) });
      stop.addEventListener('change', () => { rl.stopIfTrue = stop.checked || undefined; });
      const tr = el('tr', { class: rl === current ? 'on' : '' },
        el('td', {}, describeCond(rl)), el('td', {}, cfPreview(rl)), el('td', {}, rangeIn), el('td', { style: { textAlign: 'center' } }, stop));
      tr.addEventListener('mousedown', () => { if (current !== rl) { current = rl; tbody.querySelectorAll('tr.on').forEach((x) => x.classList.remove('on')); tr.classList.add('on'); } });
      tr.addEventListener('dblclick', (e) => { if (e.target.tagName !== 'INPUT') edit(); });
      tbody.append(tr);
    }
    btnEdit.disabled = btnDel.disabled = btnDup.disabled = !current;
    const vis2 = visible();
    btnUp.disabled = !current || vis2.indexOf(current) <= 0;
    btnDown.disabled = !current || vis2.indexOf(current) >= vis2.length - 1;
  };
  const edit = () => {
    if (!current) return;
    const target = current;
    cfRuleEditor(target, (nr) => {
      for (const k of Object.keys(target)) if (!['r1', 'c1', 'r2', 'c2', 'stopIfTrue'].includes(k)) delete target[k];
      Object.assign(target, nr, { r1: target.r1, c1: target.c1, r2: target.r2, c2: target.c2 });
      render();
    }, { title: '서식 규칙 편집' });
  };
  const move = (dir) => {
    const list = listOf(scopeSheet);
    const vis = visible();
    const j = vis.indexOf(current) + dir;
    if (j < 0 || j >= vis.length) return;
    const a = list.indexOf(current);
    const b = list.indexOf(vis[j]);
    [list[a], list[b]] = [list[b], list[a]];
    render();
  };
  const btn = (label, onclick, title) => el('button', { type: 'button', class: 'btn', title, onclick }, label);
  const btnNew = btn('새 규칙...', () => {
    cfRuleEditor(null, (nr) => {
      const rg = scopeSheet === si ? selRange : { r1: 0, c1: 0, r2: 0, c2: 0 };
      const r = { ...rg, ...nr };
      listOf(scopeSheet).unshift(r);
      current = r;
      if (scope === 'selection' && scopeSheet !== si) scope = 'sheet';
      render();
    });
  });
  const btnEdit = btn('규칙 편집...', edit);
  const btnDel = btn('규칙 삭제', () => {
    const list = listOf(scopeSheet);
    list.splice(list.indexOf(current), 1);
    current = null;
    render();
  });
  const btnDup = btn('규칙 복제', () => {
    const list = listOf(scopeSheet);
    const copy = structuredClone(current);
    list.splice(list.indexOf(current), 0, copy);
    current = copy;
    render();
  });
  const btnUp = btn('▲', () => move(-1), '위로 이동 (우선순위 높임)');
  const btnDown = btn('▼', () => move(1), '아래로 이동 (우선순위 낮춤)');
  showSel.addEventListener('change', () => {
    if (showSel.value === 'selection') { scope = 'selection'; scopeSheet = si; } else { scope = 'sheet'; scopeSheet = Number(showSel.value.split(':')[1]); }
    current = null;
    render();
  });
  const apply = () => {
    const changed = [...lists.entries()].filter(([i, list]) => JSON.stringify(list) !== JSON.stringify(wb.sheets[i].cond));
    if (!changed.length) return;
    wb.transact(() => { for (const [i, list] of changed) wb.setSheetProp(i, 'cond', list.map((x) => structuredClone(x))); }, meta());
    lists.clear();
    render();
  };
  const table = el('table', { class: 'cf-table' },
    el('thead', {}, el('tr', {}, el('th', {}, '규칙(표시된 순서대로 적용)'), el('th', {}, '서식'), el('th', {}, '적용 대상'), el('th', {}, 'True일 경우 중지'))),
    tbody);
  render();
  openDialog({
    title: '조건부 서식 규칙 관리자', width: 760,
    body: el('div', { class: 'cf-manager' },
      el('div', { class: 'fc-row' }, el('span', {}, '서식 규칙 표시:'), showSel),
      el('div', { class: 'fc-row' }, btnNew, btnEdit, btnDel, btnDup, el('span', { style: { flex: '1' } }), btnUp, btnDown),
      el('div', { class: 'cf-tablewrap' }, table)),
    buttons: [
      { label: '확인', primary: true, action: () => { apply(); } },
      { label: '적용', action: () => { apply(); return false; } },
      { label: '취소' },
    ],
  });
}

function statsDialog() {
  let cells = 0;
  let formulas = 0;
  let comments = 0;
  let totalCells = 0;
  wb.sheets.forEach((s, i) => {
    for (const cell of s.cells.values()) {
      if (cell.raw) { totalCells++; if (i === si) cells++; }
      if (cell.formula && i === si) formulas++;
      if (cell.comment && i === si) comments++;
    }
  });
  const u = wb.usedRange(si);
  const rows = [
    ['마지막 셀', u.rows ? cellName(u.rows - 1, u.cols - 1) : '-'],
    ['데이터가 있는 셀', cells], ['수식', formulas], ['메모', comments], ['병합된 셀', sheet().merges.length],
    ['차트', sheet().charts.length], ['시트 수', wb.sheets.length], ['통합 문서 전체 데이터 셀', totalCells],
  ];
  openDialog({
    title: '통합 문서 통계', width: 360,
    body: el('table', { class: 'kbd-table' }, el('tr', {}, el('td', { colspan: 2 }, el('b', {}, `현재 시트: ${sheet().name}`))),
      rows.map(([k, v]) => el('tr', {}, el('td', {}, k), el('td', { style: { textAlign: 'right' } }, String(v))))),
    buttons: [{ label: '확인', primary: true }],
  });
}

const SHORTCUTS = [
  ['Enter / Shift+Enter', '입력 후 아래/위로 이동 (Tab으로 이동한 경우 시작 열로 복귀)'],
  ['Tab / Shift+Tab', '입력 후 오른쪽/왼쪽으로 이동'],
  ['F2', '셀 편집 (편집 중에는 입력/편집 모드 전환)'],
  ['Alt+Enter', '셀 안에서 줄 바꿈'],
  ['Ctrl+Enter', '선택한 모든 셀에 같은 내용 입력'],
  ['F4', '수식 편집 중 절대/상대 참조 전환'],
  ['Esc', '편집 취소 / 복사 영역 해제'],
  ['Ctrl+방향키', '데이터 영역의 끝으로 이동 (빈 열에서는 10,000,000행까지)'],
  ['Shift+방향키', '선택 영역 확장'],
  ['Ctrl+Space / Shift+Space', '열 전체 / 행 전체 선택'],
  ['Ctrl+A', '모두 선택'],
  ['Ctrl+Home / Ctrl+End', '처음 셀 / 마지막 셀로 이동'],
  ['Ctrl+C / Ctrl+X / Ctrl+V', '복사 / 잘라내기 / 붙여넣기'],
  ['Ctrl+Z / Ctrl+Y', '실행 취소 / 다시 실행'],
  ['Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5', '굵게 / 기울임꼴 / 밑줄 / 취소선'],
  ['Ctrl+D / Ctrl+R', '아래로 / 오른쪽으로 채우기'],
  ['Alt+=', '자동 합계'],
  ['Alt+F1', '차트 삽입'],
  ['Ctrl+Shift+L', '필터 켜기/끄기'],
  ['Ctrl+9 / Ctrl+0', '행 숨기기 / 열 숨기기'],
  ['Ctrl+;  /  Ctrl+Shift+;', '오늘 날짜 / 현재 시간 입력'],
  ['Ctrl+Shift+1/3/4/5', '숫자 / 날짜 / 통화 / 백분율 서식'],
  ['Ctrl+Shift+~', '일반 서식'],
  ['Ctrl+1', '셀 서식 (표시 형식 · 사용자 지정 서식 · 맞춤 · 글꼴 · 테두리 · 채우기)'],
  ['Ctrl+T / Ctrl+L', '표 만들기 (브라우저가 Ctrl+T 를 가로채면 Ctrl+L)'],
  ['Alt+A+E', '텍스트 나누기'],
  ['Alt → 글자', '리본 바로 가기 키 (예: Alt+N+T 표, Alt+N+V 피벗, Alt+H+O+E 셀 서식)'],
  ['Alt+↓', '목록·필터·요약 함수 펼치기'],
  ['Alt+C / Alt+S', '슬라이서: 필터 지우기 / 다중 선택'],
  ['Ctrl+F / Ctrl+H / Ctrl+G', '찾기 / 바꾸기 / 이동'],
  ['Ctrl+`', '수식 표시'],
  ['Ctrl+Shift+= / Ctrl+-', '행·열 삽입 / 삭제'],
  ['Ctrl+PageUp / PageDown', '이전 / 다음 시트'],
  ['Ctrl+마우스 휠', '확대/축소'],
  ['Shift+F11', '새 시트'],
  ['Shift+F2', '메모 편집'],
  ['Ctrl+S', '저장'],
  ['F9', '다시 계산'],
];

// ───────────────────────── 메뉴 정의 ─────────────────────────
const THEME = ['#FFFFFF', '#000000', '#E7E6E6', '#44546A', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47'];
const STANDARD = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];
function mix(hex, target, t) {
  const p = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const a = p(hex);
  const b = p(target);
  return `#${a.map((x, i) => Math.round(x + (b[i] - x) * t).toString(16).padStart(2, '0')).join('')}`;
}
function themeRows() {
  const rows = [THEME];
  const variants = [[0.8, '#FFFFFF'], [0.6, '#FFFFFF'], [0.4, '#FFFFFF'], [0.25, '#000000'], [0.5, '#000000']];
  for (const [t, target] of variants) {
    rows.push(THEME.map((h, i) => {
      if (i === 0) return mix('#FFFFFF', '#000000', { 0.8: 0.05, 0.6: 0.15, 0.4: 0.25, 0.25: 0.35, 0.5: 0.5 }[t]);
      if (i === 1) return mix('#000000', '#FFFFFF', { 0.8: 0.5, 0.6: 0.35, 0.4: 0.25, 0.25: 0.15, 0.5: 0.05 }[t]);
      return mix(h, target, t);
    }));
  }
  return rows;
}

function colorMenu(anchorEl, kind) {
  const pick = (color) => {
    closeMenus();
    if (kind === 'fill') { if (color) lastFill = color; applyStyle({ fill: color || undefined }); }
    else { if (color) lastFont = color; applyStyle({ color: color || undefined }); }
    focusGrid();
  };
  const swatches = (colors, gap) => el('div', { class: `palette-row${gap ? ' gap' : ''}` }, colors.map((c) => el('button', {
    class: 'swatch', title: c, style: { background: c }, onmousedown: (e) => e.preventDefault(), onclick: () => pick(c),
  })));
  const custom = el('input', { type: 'color', style: { width: '0', height: '0', opacity: '0', position: 'absolute' } });
  custom.addEventListener('change', () => pick(custom.value));
  const rows = themeRows();
  const palette = el('div', { class: 'palette' },
    el('div', { class: 'menu-title', style: { padding: '0 0 4px' } }, '테마 색'),
    rows.map((r, i) => swatches(r, i === 0 || i === rows.length - 1)),
    el('div', { class: 'menu-title', style: { padding: '4px 0' } }, '표준 색'),
    swatches(STANDARD));
  openMenu(anchorEl, [
    kind === 'fill' ? { label: '채우기 없음', action: () => pick(null) } : { label: '자동', action: () => pick(null) },
    { node: palette },
    { sep: true },
    { node: el('div', {}, custom) },
    { label: '다른 색...', action: () => custom.click() },
  ]);
}

const CELL_STYLES = [
  { name: '표준', style: null },
  { name: '좋음', style: { fill: '#c6efce', color: '#006100' } },
  { name: '나쁨', style: { fill: '#ffc7ce', color: '#9c0006' } },
  { name: '보통', style: { fill: '#ffeb9c', color: '#9c5700' } },
  { name: '계산', style: { fill: '#f2f2f2', color: '#fa7d00', bold: true, bt: true, bb: true, bl: true, br: true } },
  { name: '확인할 셀', style: { fill: '#a5a5a5', color: '#ffffff', bold: true } },
  { name: '경고문', style: { color: '#ff0000' } },
  { name: '메모', style: { fill: '#ffffcc', bt: true, bb: true, bl: true, br: true } },
  { name: '입력', style: { fill: '#ffcc99', color: '#3f3f76', bt: true, bb: true, bl: true, br: true } },
  { name: '출력', style: { fill: '#f2f2f2', color: '#3f3f3f', bold: true, bt: true, bb: true, bl: true, br: true } },
  { name: '설명 텍스트', style: { color: '#7f7f7f', italic: true } },
  { name: '연결된 셀', style: { color: '#fa7d00', bb: true } },
  { name: '제목', style: { size: 18, bold: true, color: '#44546a' } },
  { name: '제목 1', style: { size: 15, bold: true, color: '#44546a', bb: true } },
  { name: '제목 2', style: { size: 13, bold: true, color: '#44546a', bb: true } },
  { name: '요약', style: { bold: true, bt: true, bb: true } },
  { name: '강조색1', style: { fill: '#4472c4', color: '#ffffff' } },
  { name: '강조색2', style: { fill: '#ed7d31', color: '#ffffff' } },
  { name: '강조색3', style: { fill: '#a5a5a5', color: '#ffffff' } },
  { name: '강조색6', style: { fill: '#70ad47', color: '#ffffff' } },
];
const RESET_STYLE = { fill: undefined, color: undefined, bold: undefined, italic: undefined, underline: undefined, strike: undefined, size: undefined, font: undefined, bt: undefined, bb: undefined, bl: undefined, br: undefined };

function cellStylesMenu(anchorEl) {
  const chip = (s) => {
    const st = s.style ?? {};
    return el('button', {
      class: 'style-chip', title: s.name, onmousedown: (e) => e.preventDefault(),
      style: {
        background: st.fill ?? '#fff', color: st.color ?? '#000', fontWeight: st.bold ? '700' : '400',
        fontStyle: st.italic ? 'italic' : 'normal', fontSize: st.size ? `${Math.min(st.size, 14)}px` : '11.5px',
        borderBottom: st.bb ? '2px solid #44546a' : undefined,
      },
      onclick: () => { closeMenus(); applyStyle({ ...RESET_STYLE, ...(s.style ?? {}) }); focusGrid(); },
    }, s.name);
  };
  openMenu(anchorEl, [{ title: '셀 스타일' }, { node: el('div', { class: 'style-grid' }, CELL_STYLES.map(chip)) }]);
}

function tableStylesMenu(anchorEl) { tableStyleGallery(anchorEl, !tableHere()); }

const MENUS = {
  tableStylesDesign: (a) => { tableStyleGallery(a, false); },
  shapes: () => [
    { title: '도형' },
    ...SHAPE_KINDS.map((k) => ({
      label: k.label, action: () => startDraw(k.id),
      icon: shapeSvg({ kind: k.id, w: 18, h: k.id === 'line' ? 14 : 13, fill: k.id === 'textbox' ? '#ffffff' : '#dbe5f5', stroke: '#4472c4', flipV: k.id === 'line' }),
    })),
  ],
  validation: () => [
    { label: '데이터 유효성 검사...', icon: 'validation', action: validationDialog },
    { label: '잘못된 데이터 표시', action: () => run('circleInvalid') },
    { label: '유효성 표시 지우기', action: () => run('clearCircles') },
  ],
  paste: () => [
    { label: '붙여넣기', icon: 'paste', key: 'Ctrl+V', action: () => pasteFromButton('all') },
    { label: '값 붙여넣기', action: () => pasteFromButton('values'), disabled: !clip },
    { label: '수식 붙여넣기', icon: 'fx', action: () => pasteFromButton('formulas'), disabled: !clip },
    { label: '서식 붙여넣기', icon: 'painter', action: () => pasteFromButton('formats'), disabled: !clip },
    { label: '행/열 바꿈', action: () => pasteFromButton('transpose'), disabled: !clip },
  ],
  borders: () => [
    { title: '테두리' },
    { label: '아래쪽 테두리', icon: 'borderBottom', action: () => applyBorder('bottom') },
    { label: '위쪽 테두리', icon: 'borderTop', action: () => applyBorder('top') },
    { label: '왼쪽 테두리', icon: 'borderLeft', action: () => applyBorder('left') },
    { label: '오른쪽 테두리', icon: 'borderRight', action: () => applyBorder('right') },
    { sep: true },
    { label: '테두리 없음', icon: 'borderNone', action: () => applyBorder('none') },
    { label: '모든 테두리', icon: 'borderAll', action: () => applyBorder('all') },
    { label: '바깥쪽 테두리', icon: 'borderOutside', action: () => applyBorder('outside') },
    { label: '위쪽/아래쪽 테두리', icon: 'borderBottom', action: () => applyBorder('topBottom') },
  ],
  fillColor: (a) => colorMenu(a, 'fill'),
  fontColor: (a) => colorMenu(a, 'font'),
  merge: () => [
    { label: '병합하고 가운데 맞춤', icon: 'merge', action: () => toggleMerge('center') },
    { label: '전체 병합', icon: 'merge', action: () => toggleMerge('across') },
    { label: '셀 병합', icon: 'merge', action: () => toggleMerge('merge') },
    { label: '셀 분할', icon: 'merge', action: () => toggleMerge('unmerge') },
  ],
  condFormat: () => {
    const add = (rule) => wb.transact(() => wb.addCondRule(si, { ...usedClip(sel), ...rule }), meta());
    return [
      { title: '셀 강조 규칙' },
      { label: '보다 큼...', action: () => condRuleDialog('gt') },
      { label: '보다 작음...', action: () => condRuleDialog('lt') },
      { label: '다음 값의 사이에 있음...', action: () => condRuleDialog('between') },
      { label: '같음...', action: () => condRuleDialog('eq') },
      { label: '텍스트 포함...', action: () => condRuleDialog('text') },
      { label: '발생 날짜...', action: () => cfRuleEditor({ type: 'date', period: 'today', style: { fill: '#ffc7ce', color: '#9c0006' } }, (nr) => wb.transact(() => wb.addCondRule(si, { ...usedClip(sel), ...nr }), meta()), { title: '발생 날짜' }) },
      { label: '중복 값...', action: () => condRuleDialog('dup') },
      { title: '상위/하위 규칙' },
      { label: '상위 10개 항목...', action: () => condRuleDialog('top') },
      { label: '상위 10%...', action: () => cfRuleEditor({ type: 'top', v1: '10', percent: true, style: { fill: '#ffc7ce', color: '#9c0006' } }, (nr) => wb.transact(() => wb.addCondRule(si, { ...usedClip(sel), ...nr }), meta()), { title: '상위 10%' }) },
      { label: '하위 10개 항목...', action: () => cfRuleEditor({ type: 'bottom', v1: '10', style: { fill: '#ffc7ce', color: '#9c0006' } }, (nr) => wb.transact(() => wb.addCondRule(si, { ...usedClip(sel), ...nr }), meta()), { title: '하위 10개 항목' }) },
      { label: '평균 초과', action: () => add({ type: 'aboveAvg', style: { fill: '#ffc7ce', color: '#9c0006' } }) },
      { label: '평균 미만', action: () => add({ type: 'belowAvg', style: { fill: '#ffc7ce', color: '#9c0006' } }) },
      { title: '데이터 막대' },
      ...[['파랑', '#8fb3e8'], ['녹색', '#8fd19e'], ['빨강', '#f19c9c'], ['주황', '#f7c07e']].map(([n, color]) => ({
        label: `${n} 데이터 막대`, icon: `<span style="display:block;width:16px;height:10px;background:${color}"></span>`,
        action: () => add({ type: 'bar', color }),
      })),
      { title: '색조' },
      ...[['녹색 - 노랑 - 빨강', ['#63be7b', '#ffeb84', '#f8696b']], ['빨강 - 노랑 - 녹색', ['#f8696b', '#ffeb84', '#63be7b']], ['흰색 - 녹색', ['#fcfcff', '#63be7b']], ['흰색 - 빨강', ['#fcfcff', '#f8696b']]].map(([n, colors]) => ({
        label: n, icon: `<span style="display:block;width:16px;height:10px;background:linear-gradient(90deg,${colors.join(',')})"></span>`,
        action: () => add({ type: 'scale', colors }),
      })),
      { title: '아이콘 집합' },
      ...ICON_SETS.slice(0, 5).map((set) => ({
        label: set.label, icon: `<span style="display:inline-flex">${set.icons.map((i) => ICON_SVG[i]).join('')}</span>`,
        action: () => add({ type: 'icons', icons: set.id }),
      })),
      { sep: true },
      { label: '새 규칙...', icon: 'condFormat', action: () => cfRuleEditor(null, (nr) => wb.transact(() => wb.addCondRule(si, { ...usedClip(sel), ...nr }), meta())) },
      { label: '규칙 관리...', action: cfManager },
      { label: '규칙 지우기 - 선택한 셀', action: () => wb.transact(() => wb.clearCondRules(si, sel), meta()) },
      { label: '규칙 지우기 - 시트 전체', action: () => wb.transact(() => wb.clearCondRules(si), meta()) },
    ];
  },
  tableStyles: (a) => tableStylesMenu(a),
  cellStyles: (a) => cellStylesMenu(a),
  insert: () => [
    { label: '시트 행 삽입', icon: 'rowInsert', action: () => run('insertRows') },
    { label: '시트 열 삽입', icon: 'colInsert', action: () => run('insertCols') },
    { label: '시트 삽입', icon: 'sheetInsert', action: () => run('addSheet') },
  ],
  delete: () => [
    { label: '시트 행 삭제', icon: 'delete', action: () => run('deleteRows') },
    { label: '시트 열 삭제', icon: 'delete', action: () => run('deleteCols') },
    { label: '시트 삭제', icon: 'delete', action: () => run('deleteSheet'), disabled: wb.sheets.length < 2 },
  ],
  format: () => [
    { title: '셀 크기' },
    { label: '행 높이...', action: () => sizeDialog('row') },
    {
      label: '행 높이 자동 맞춤', action: () => {
        const u = usedClip(sel);
        const r2 = Math.min(u.r2, u.r1 + 5000);
        wb.transact(() => { for (let r = u.r1; r <= r2; r++) wb.setRowHeight(si, r, DEFAULT_ROW_HEIGHT, false); autoFitRows(u.r1, r2); }, meta());
      },
    },
    { label: '열 너비...', action: () => sizeDialog('col') },
    { label: '열 너비 자동 맞춤', action: () => autofitCols(range(sel.c1, Math.min(sel.c2, sel.c1 + 200))) },
    ...MENUS_HIDE(),
    { title: '시트 구성' },
    { label: '시트 이름 바꾸기', action: () => renameSheetInline(si) },
    { label: '시트 복사본 만들기', icon: 'copy', action: () => run('duplicateSheet') },
    { title: '보호' },
    { label: '셀 서식...', key: 'Ctrl+1', action: () => formatCellsDialog() },
  ],
  hideMenu: () => MENUS_HIDE(),
  autosum: () => [
    ['합계', 'SUM'], ['평균', 'AVERAGE'], ['숫자 개수', 'COUNT'], ['최대값', 'MAX'], ['최소값', 'MIN'],
  ].map(([label, fn]) => ({ label, action: () => autoSum(fn) })).concat([{ sep: true }, { label: '기타 함수...', icon: 'fx', action: () => insertFunctionDialog() }]),
  fill: () => [
    { label: '아래쪽', key: 'Ctrl+D', action: () => fillCopy('down') },
    { label: '오른쪽', key: 'Ctrl+R', action: () => fillCopy('right') },
    { label: '위쪽', action: () => fillCopy('up') },
    { label: '왼쪽', action: () => fillCopy('left') },
  ],
  clear: () => [
    { label: '모두 지우기', icon: 'clear', action: () => run('clearAll') },
    { label: '서식 지우기', action: () => run('clearFormats') },
    { label: '내용 지우기', key: 'Delete', action: () => run('clearContents') },
    { label: '메모 지우기', action: () => run('clearComments') },
  ],
  sort: () => [
    { label: '텍스트 오름차순 정렬', icon: 'sortAsc', action: () => sortData(true) },
    { label: '텍스트 내림차순 정렬', icon: 'sortDesc', action: () => sortData(false) },
    { label: '사용자 지정 정렬...', icon: 'sort', action: () => sortDialog() },
    { sep: true },
    { label: '필터', icon: 'filter', key: 'Ctrl+Shift+L', checked: filterKeyHere() !== null && !!getFilter(filterKeyHere()), action: () => toggleFilter() },
    { label: '지우기', icon: 'filterClear', disabled: filterKeyHere() === null, action: () => run('clearFilter') },
    { label: '다시 적용', icon: 'refresh', disabled: filterKeyHere() === null, action: () => run('reapplyFilter') },
    { sep: true },
    { label: '중복된 항목 제거', icon: 'dedupe', action: () => removeDuplicates() },
  ],
  find: () => [
    { label: '찾기...', icon: 'search', key: 'Ctrl+F', action: () => openFindDialog('find') },
    { label: '바꾸기...', key: 'Ctrl+H', action: () => openFindDialog('replace') },
    { label: '이동...', key: 'Ctrl+G', action: () => gotoDialog() },
    { sep: true },
    { label: '다음 메모', icon: 'comment', action: () => jumpComment(1) },
  ],
  freeze: () => {
    const f = sheet().freeze ?? {};
    return [
      f.rows || f.cols
        ? { label: '틀 고정 취소', icon: 'freeze', action: () => setFreeze(0, 0) }
        : { label: `틀 고정 (${cellName(active.r, active.c)}의 위쪽·왼쪽)`, icon: 'freeze', disabled: !active.r && !active.c, action: () => setFreeze(active.r, active.c) },
      { label: '첫 행 고정', action: () => setFreeze(1, 0) },
      { label: '첫 열 고정', action: () => setFreeze(0, 1) },
    ];
  },
  pieCharts: () => [
    { label: '원형', icon: 'chartPie', action: () => insertChart('pie') },
    { label: '도넛형', icon: 'chartDoughnut', action: () => insertChart('doughnut') },
  ],
};

function MENUS_HIDE() {
  return [
    { title: '숨기기 및 숨기기 취소' },
    { label: '행 숨기기', key: 'Ctrl+9', action: () => hideSel('row', true) },
    { label: '열 숨기기', key: 'Ctrl+0', action: () => hideSel('col', true) },
    { label: '행 숨기기 취소', action: () => hideSel('row', false) },
    { label: '열 숨기기 취소', action: () => hideSel('col', false) },
  ];
}

function openNamedMenu(name, anchorEl) {
  if (editing && !commitEdit()) return;
  if (name.startsWith('fn:')) {
    const cat = name.slice(3);
    const names = FUNCTION_NAMES.filter((n) => (cat === 'more' ? ['통계', '정보'].includes(FUNC_INFO[n]?.cat) : FUNC_INFO[n]?.cat === cat));
    openMenu(anchorEl, names.map((n) => ({ label: n, action: () => insertFunctionText(n) })));
    return;
  }
  const items = MENUS[name]?.(anchorEl);
  if (Array.isArray(items)) openMenu(anchorEl, items);
}

function showContextMenu(pos, kind) {
  const items = [
    { label: '잘라내기', icon: 'cut', key: 'Ctrl+X', action: () => { copySelection(true); } },
    { label: '복사', icon: 'copy', key: 'Ctrl+C', action: () => { const { text } = copySelection(false); navigator.clipboard?.writeText(text).catch(() => {}); } },
    { label: '붙여넣기', icon: 'paste', key: 'Ctrl+V', action: () => pasteFromButton('all') },
    { label: '값 붙여넣기', action: () => pasteFromButton('values'), disabled: !clip },
    { sep: true },
  ];
  if (kind === 'col') {
    items.push(
      { label: '삽입', icon: 'colInsert', action: () => run('insertCols') },
      { label: '삭제', icon: 'delete', action: () => run('deleteCols') },
      { label: '내용 지우기', action: () => run('clearContents') },
      { sep: true },
      { label: '열 너비...', action: () => sizeDialog('col') },
      { label: '열 너비 자동 맞춤', action: () => autofitCols(range(sel.c1, Math.min(sel.c2, sel.c1 + 200))) },
      { label: '숨기기', icon: 'hide', action: () => hideSel('col', true) },
      { label: '숨기기 취소', action: () => hideSel('col', false) },
    );
  } else if (kind === 'row') {
    items.push(
      { label: '삽입', icon: 'rowInsert', action: () => run('insertRows') },
      { label: '삭제', icon: 'delete', action: () => run('deleteRows') },
      { label: '내용 지우기', action: () => run('clearContents') },
      { sep: true },
      { label: '행 높이...', action: () => sizeDialog('row') },
      { label: '숨기기', icon: 'hide', action: () => hideSel('row', true) },
      { label: '숨기기 취소', action: () => hideSel('row', false) },
    );
  } else {
    const cm = wb.getCell(si, active.r, active.c)?.comment;
    items.push(
      { label: '행 삽입', icon: 'rowInsert', action: () => run('insertRows') },
      { label: '열 삽입', icon: 'colInsert', action: () => run('insertCols') },
      { label: '행 삭제', icon: 'delete', action: () => run('deleteRows') },
      { label: '열 삭제', icon: 'delete', action: () => run('deleteCols') },
      { label: '내용 지우기', action: () => run('clearContents') },
      { sep: true },
      { label: '필터', icon: 'filter', checked: filterKeyHere() !== null && !!getFilter(filterKeyHere()), action: () => toggleFilter() },
      { label: '오름차순 정렬', icon: 'sortAsc', action: () => sortData(true) },
      { label: '내림차순 정렬', icon: 'sortDesc', action: () => sortData(false) },
      { sep: true },
      cm ? { label: '메모 편집', icon: 'comment', action: () => editComment() } : { label: '새 메모', icon: 'newComment', action: () => editComment() },
      cm ? { label: '메모 삭제', icon: 'deleteComment', action: () => run('deleteComment') } : null,
      { sep: true },
      { label: '셀 서식...', key: 'Ctrl+1', action: () => formatCellsDialog() },
      { label: '병합하고 가운데 맞춤', icon: 'merge', action: () => toggleMerge('center'), disabled: selIsActiveOnly() && !wb.mergeAt(si, active.r, active.c) },
    );
  }
  openMenu(pos, items);
}

// ───────────────────────── 명령 ─────────────────────────
const structural = (fn) => () => { fn(); gv.layout(); updateSelectionUI(); };

const COMMANDS = {
  undo: () => { const m = wb.undo(); if (m) restoreMeta(m); else toast('실행 취소할 작업이 없습니다.'); },
  redo: () => { const m = wb.redo(); if (m) restoreMeta(m); },
  save: () => saveNow(true),
  saveAs,
  open: () => openBackstage('open'),
  backstage: () => openBackstage(),
  print: printSheet,

  cut: () => { const { text } = copySelection(true); navigator.clipboard?.writeText(text).catch(() => {}); },
  copy: () => { const { text } = copySelection(false); navigator.clipboard?.writeText(text).catch(() => {}); },
  paste: () => pasteFromButton('all'),
  painter: () => {
    if (painter) { painter = null; dom.view.classList.remove('painting'); setMode(); updateRibbon(); }
    else capturePainter(false);
  },
  painterSticky: () => capturePainter(true),

  bold: () => toggleStyle('bold'),
  italic: () => toggleStyle('italic'),
  underline: () => toggleStyle('underline'),
  strike: () => toggleStyle('strike'),
  fontFamily: (f) => applyStyle({ font: f === DEFAULT_FONT ? undefined : f }),
  fontSize: (s) => { const n = Number(s); if (n > 0 && n <= 409) applyStyle({ size: n === DEFAULT_SIZE ? undefined : n }); },
  growFont: () => changeFontSize(1),
  shrinkFont: () => changeFontSize(-1),
  borderLast: () => applyBorder(lastBorder),
  borderOutside: () => applyBorder('outside'),
  fillColor: () => applyStyle({ fill: lastFill }),
  fontColor: () => applyStyle({ color: lastFont }),
  fontDialog: formatCellsDialog,
  formatCells: formatCellsDialog,

  alignLeft: () => applyStyle({ align: styleAt(active.r, active.c).align === 'left' ? undefined : 'left' }),
  alignCenter: () => applyStyle({ align: styleAt(active.r, active.c).align === 'center' ? undefined : 'center' }),
  alignRight: () => applyStyle({ align: styleAt(active.r, active.c).align === 'right' ? undefined : 'right' }),
  valignTop: () => applyStyle({ valign: 'top' }),
  valignMiddle: () => applyStyle({ valign: 'middle' }),
  valignBottom: () => applyStyle({ valign: undefined }),
  wrap: () => toggleStyle('wrap'),
  indentInc: () => applyStyle((s) => ({ indent: Math.min(15, (s.indent || 0) + 1), align: s.align === 'right' ? 'right' : 'left' })),
  indentDec: () => applyStyle((s) => ({ indent: Math.max(0, (s.indent || 0) - 1) || undefined })),
  mergeCenter: () => toggleMerge('center'),

  numFmt: (f) => (f === 'custom' || f === 'more' ? formatCellsDialog(0) : applyStyle({ numFmt: f === 'general' ? undefined : f, decimals: undefined }, { widen: true })),
  fmtGeneral: () => applyStyle({ numFmt: undefined, decimals: undefined }),
  fmtNumber: () => applyStyle({ numFmt: 'number', decimals: 2 }, { widen: true }),
  fmtDate: () => applyStyle({ numFmt: 'date', decimals: undefined }, { widen: true }),
  fmtCurrency: () => applyStyle({ numFmt: 'accounting', decimals: undefined }, { widen: true }),
  fmtPercent: () => applyStyle({ numFmt: 'percent', decimals: undefined }, { widen: true }),
  fmtComma: () => applyStyle({ numFmt: 'comma', decimals: undefined }, { widen: true }),
  incDecimal: () => changeDecimals(1),
  decDecimal: () => changeDecimals(-1),

  insertRows: structural(() => {
    const n = selKind === 'cols' || selKind === 'all' ? 1 : Math.min(sel.r2 - sel.r1 + 1, 100000);
    wb.transact(() => wb.insertRows(si, sel.r1, n), meta());
  }),
  insertCols: structural(() => {
    const n = selKind === 'rows' || selKind === 'all' ? 1 : Math.min(sel.c2 - sel.c1 + 1, 5000);
    wb.transact(() => wb.insertCols(si, sel.c1, n), meta());
  }),
  deleteRows: structural(() => {
    const r1 = sel.r1;
    const n = selKind === 'cols' ? 1 : sel.r2 - sel.r1 + 1;
    wb.transact(() => wb.deleteRows(si, r1, n), meta());
    selectCell(r1, active.c);
  }),
  deleteCols: structural(() => {
    const c1 = sel.c1;
    const n = selKind === 'rows' ? 1 : sel.c2 - sel.c1 + 1;
    wb.transact(() => wb.deleteCols(si, c1, n), meta());
    selectCell(active.r, c1);
  }),
  insertMenuKey: () => {
    if (selKind === 'rows') run('insertRows');
    else if (selKind === 'cols') run('insertCols');
    else { const b = gv.clientRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c }); openMenu({ x: b.left, y: b.bottom }, MENUS.insert()); }
  },
  deleteMenuKey: () => {
    if (selKind === 'rows') run('deleteRows');
    else if (selKind === 'cols') run('deleteCols');
    else { const b = gv.clientRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c }); openMenu({ x: b.left, y: b.bottom }, MENUS.delete()); }
  },
  addSheet: () => {
    const i = wb.transact(() => wb.addSheet(null, si + 1), meta());
    switchSheet(i, false);
  },
  deleteSheet: () => {
    if (wb.sheets.length < 2) { toast('통합 문서에는 시트가 하나 이상 있어야 합니다.'); return; }
    const del = () => {
      const i = si;
      wb.transact(() => wb.deleteSheet(i), meta());
      si = -1;
      switchSheet(Math.min(i, wb.sheets.length - 1), false);
    };
    if (sheet().cells.size > 0) {
      openDialog({
        title: 'Tabula', body: `'${sheet().name}' 시트를 영구적으로 삭제합니다. 계속하시겠습니까?`,
        buttons: [{ label: '삭제', primary: true, action: del }, { label: '취소' }],
      });
    } else del();
  },
  duplicateSheet: () => {
    const src = wb.serialize().sheets[si];
    const i = wb.transact(() => {
      const at = wb.addSheet(`${src.name} (2)`.slice(0, 31), si + 1);
      const data = wb.serialize();
      const dup = (list) => (list ?? []).map((c) => ({ ...c, id: `${c.id}d${at}` }));
      data.sheets[at] = { ...src, name: wb.sheets[at].name, charts: dup(src.charts), images: dup(src.images), shapes: dup(src.shapes) };
      wb.restore(data);
      return at;
    }, meta());
    switchSheet(i, false);
  },
  prevSheet: () => switchSheet(si - 1),
  nextSheet: () => switchSheet(si + 1),

  autosum: () => autoSum('SUM'),
  fillDown: () => fillCopy('down'),
  fillRight: () => fillCopy('right'),
  clearContents: () => wb.transact(() => {
    if (allFilters().some(([, f]) => Object.keys(f.hidden ?? {}).length)) {
      for (const [r, c] of cellsIn(usedClip(sel))) {
        if (filterHidden(r)) continue;
        const cell = wb.getCell(si, r, c);
        if (cell?.raw) wb.setCellData(si, r, c, { raw: '', style: cell.style, comment: cell.comment });
      }
    } else wb.clearRange(si, sel.r1, sel.c1, sel.r2, sel.c2, 'contents');
  }, meta()),
  clearAll: () => wb.transact(() => { wb.clearRange(si, sel.r1, sel.c1, sel.r2, sel.c2, 'all'); wb.unmerge(si, sel.r1, sel.c1, sel.r2, sel.c2); }, meta()),
  clearFormats: () => wb.transact(() => wb.clearRange(si, sel.r1, sel.c1, sel.r2, sel.c2, 'formats'), meta()),
  clearComments: () => wb.transact(() => wb.clearRange(si, sel.r1, sel.c1, sel.r2, sel.c2, 'comments'), meta()),
  sortAsc: () => sortData(true),
  sortDesc: () => sortData(false),
  sortDialog,
  dedupe: removeDuplicates,
  find: () => openFindDialog('find'),
  replace: () => openFindDialog('replace'),
  goto: gotoDialog,

  toggleFilter,
  clearFilter: () => {
    const key = filterKeyHere();
    const f = key === null ? null : getFilter(key);
    if (!f) return;
    wb.transact(() => putFilter(key, { ...f, criteria: {}, hidden: {}, sort: undefined }), meta());
    gv.layout();
    setMode();
  },
  reapplyFilter: () => {
    const key = filterKeyHere();
    const f = key === null ? null : getFilter(key);
    if (!f) return;
    wb.transact(() => putFilter(key, recomputeFilter(f, key)), meta());
    gv.layout();
    setMode();
  },
  hideRows: () => hideSel('row', true),
  hideCols: () => hideSel('col', true),

  insertPivot: pivotDialog,
  refreshAll: refreshPivots,
  chartColumn: () => insertChart('column'),
  chartBar: () => insertChart('bar'),
  chartLine: () => insertChart('line'),
  chartPie: () => insertChart('pie'),
  chartArea: () => insertChart('area'),
  chartScatter: () => insertChart('scatter'),
  insertPicture,
  textToColumns,
  condManager: cfManager,
  condNewRule: () => cfRuleEditor(null, (nr) => wb.transact(() => wb.addCondRule(si, { ...usedClip(sel), ...nr }), meta())),
  createTable: () => createTableDialog(),
  insertSlicer: insertSlicerDialog,
  slicerCaption: (v) => { if (chartSel) updateObject(chartSel, { caption: String(v ?? '') }); },
  slicerCols: (v) => { if (chartSel) updateObject(chartSel, { columns: clamp(Number(v) || 1, 1, 20) }); },
  slicerClear: () => { if (chartSel) slicerClear(chartSel); },
  slicerMulti: () => { const sl = (sheet().slicers ?? []).find((x) => x.id === chartSel); if (sl) updateObject(sl.id, { multi: !sl.multi }); },
  slicerSettings: () => { if (chartSel) slicerSettings(chartSel); },
  tblName: (v) => renameTable(v),
  resizeTable: resizeTableDialog,
  convertToRange: convertTableToRange,
  pivotFromTable: () => { const t = tableHere(); if (t) pivotDialog(t.name); },
  tblHeader: () => toggleTableOption('header'),
  tblTotals: () => toggleTableOption('totals'),
  tblBanded: () => toggleTableOption('banded'),
  tblBandedCols: () => toggleTableOption('bandedCols'),
  tblFirstCol: () => toggleTableOption('firstCol'),
  tblLastCol: () => toggleTableOption('lastCol'),
  tblFilter: () => toggleTableOption('filter'),
  autofitSel: () => autofitCols(range(sel.c1, Math.min(sel.c2, sel.c1 + 200))),
  autofitRowsSel: () => wb.transact(() => autoFitRows(sel.r1, Math.min(sel.r2, sel.r1 + 2000)), meta()),
  condMenuKey: () => menuAtCell('condFormat'),
  tableStyleKey: () => menuAtCell('tableStyles'),
  freezePanes: () => { const f = sheet().freeze ?? {}; if (f.rows || f.cols) setFreeze(0, 0); else setFreeze(active.r, active.c); },
  freezeTop: () => setFreeze(1, 0),
  freezeFirstCol: () => setFreeze(0, 1),
  shapesMenu: () => startDraw('rect'),
  insertTextbox: () => startDraw('textbox'),
  dataValidation: validationDialog,
  circleInvalid: () => { circles = invalidCells(wb, si); gv.renderOverlays(); if (!circles.length) toast('잘못된 데이터가 없습니다.'); },
  clearCircles: () => { circles = null; gv.renderOverlays(); },
  macros: macroDialog,

  insertFunction: insertFunctionDialog,
  insertDate: () => {
    const d = new Date();
    const text = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    wb.transact(() => { wb.setInput(si, active.r, active.c, text); autoWiden({ r1: active.r, c1: active.c, r2: active.r, c2: active.c }); }, meta());
  },
  insertTime: () => {
    const d = new Date();
    wb.transact(() => { wb.setInput(si, active.r, active.c, `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`); autoWiden({ r1: active.r, c1: active.c, r2: active.r, c2: active.c }); }, meta());
  },
  editComment,
  deleteComment: () => wb.transact(() => wb.setComment(si, active.r, active.c, ''), meta()),
  prevComment: () => jumpComment(-1),
  nextComment: () => jumpComment(1),
  workbookStats: statsDialog,

  importCsv: () => pickFile('import'),
  exportCsv,

  toggleGrid: (v) => { view.showGrid = v ?? !view.showGrid; applyView(); },
  togglePrintGrid: (v) => { view.printGrid = v ?? !view.printGrid; updateRibbon(); },
  toggleFormulaBar: (v) => { view.showFormulaBar = v ?? !view.showFormulaBar; applyView(); },
  toggleHeaders: (v) => { view.showHeaders = v ?? !view.showHeaders; applyView(); },
  toggleFormulas: () => { view.showFormulas = !view.showFormulas; gv.renderAll(); updateRibbon(); },
  toggleRibbon: () => ribbon.toggleCollapse(),
  zoomIn: () => setZoom(view.zoom + 10),
  zoomOut: () => setZoom(view.zoom - 10),
  zoom100: () => setZoom(100),
  recalc: () => { wb.invalidate(); gv.renderAll(); },

  shortcuts: () => openDialog({
    title: '바로 가기 키', width: 580,
    body: el('table', { class: 'kbd-table' }, SHORTCUTS.map(([k, d]) => el('tr', {}, el('td', {}, k), el('td', {}, d)))),
    buttons: [{ label: '닫기', primary: true }],
  }),
  about: () => openDialog({
    title: 'Tabula 정보', width: 420,
    body: el('div', { style: { lineHeight: '1.7' } },
      el('b', {}, 'Tabula'), ' — 브라우저에서 동작하는 엑셀 스타일 스프레드시트', el('br'),
      el('span', { class: 'muted' }, `시트 크기 10,000,000행 × 16,384열 · 함수 ${FUNCTION_NAMES.length}개 · .xlsx 열기/저장`), el('br'),
      el('span', { class: 'muted' }, server.available ? '서버 저장소에 연결됨 — 다른 기기에서도 열 수 있습니다.' : '서버 없이 실행 중 — 이 브라우저에 저장됩니다.')),
    buttons: [{ label: '확인', primary: true }],
  }),
};

const NO_COMMIT = new Set(['toggleRibbon', 'zoomIn', 'zoomOut', 'zoom100', 'shortcuts', 'about']);

function run(cmd, arg) {
  closeMenus();
  if (editing && !NO_COMMIT.has(cmd)) {
    if (cmd === 'undo') { cancelEdit(); return; }
    if (!commitEdit()) return;
  }
  const fn = COMMANDS[cmd];
  if (!fn) { toast('지원하지 않는 기능입니다.'); return; }
  fn(arg);
  focusGrid();
}

function restoreMeta(m) {
  if (m.si !== si && m.si < wb.sheets.length) switchSheet(m.si);
  gv.layout();
  selectRange(m.sel, m.selKind ?? 'cells', m.active);
  gv.ensureVisible(m.active.r, m.active.c);
}

// ───────────────────────── 보기 ─────────────────────────
function applyView() {
  dom.formulaRow.classList.toggle('hidden', !view.showFormulaBar);
  if (!gv) return;
  gv.layout();
  updateSelectionUI();
}

function setZoom(z) {
  view.zoom = clamp(Math.round(z), 25, 400);
  dom.zoomSlider.value = view.zoom;
  dom.zoomLabel.textContent = `${view.zoom}%`;
  gv.setZoom(view.zoom);
  positionEditor();
}

function ribbonState() {
  const st = styleAt(active.r, active.c);
  const fmt = st.numFmt === 'comma' ? 'number' : st.numFmt === 'datetime' ? 'date' : st.numFmt || 'general';
  const f = sheet().freeze ?? {};
  return {
    bold: st.bold, italic: st.italic, underline: st.underline, strike: st.strike, wrap: st.wrap,
    font: st.font || DEFAULT_FONT, size: String(st.size || DEFAULT_SIZE), numFmt: st.numFmt === 'custom' ? 'custom' : fmt,
    alignLeft: st.align === 'left', alignCenter: st.align === 'center', alignRight: st.align === 'right',
    valignTop: st.valign === 'top', valignMiddle: st.valign === 'middle', valignBottom: !st.valign,
    merged: !!wb.mergeAt(si, active.r, active.c), painter: !!painter, filterOn: (() => { const k = filterKeyHere(); return k !== null && !!getFilter(k); })(),
    frozen: !!(f.rows || f.cols), lastFill, lastFont, ...view,
    ...tableRibbonState(),
  };
}

function tableRibbonState() {
  const t = chartSel ? null : tableHere();
  const sl = chartSel ? (sheet().slicers ?? []).find((x) => x.id === chartSel) : null;
  const context = [...(t ? ['table'] : []), ...(sl ? ['slicer'] : [])];
  if (!t) return { context, slicerCaption: sl?.caption ?? '', slicerCols: String(sl?.columns ?? 1), slicerMultiOn: !!sl?.multi };
  return {
    context, tblName: t.name, tblHeader: t.header, tblTotals: t.totals, tblBanded: t.banded !== false, tblBandedCols: !!t.bandedCols,
    tblFirstCol: !!t.firstCol, tblLastCol: !!t.lastCol, tblFilter: !!t.filter, slicerCaption: sl?.caption ?? '', slicerCols: String(sl?.columns ?? 1), slicerMultiOn: !!sl?.multi,
  };
}

function updateRibbon() {
  ribbon?.update(ribbonState());
  dom.undoBtn.disabled = !wb.canUndo();
  dom.redoBtn.disabled = !wb.canRedo();
}

function updateTitle() {
  const t = `${docName}${!autosave && dirty ? '*' : ''} - Tabula`;
  if (!dom.title.querySelector('input')) dom.title.textContent = t;
  document.title = t;
  dom.autosave.setAttribute('aria-checked', String(autosave));
  dom.autosaveLabel.textContent = autosave ? '켬' : '끔';
  let state;
  if (!server.available) state = dirty && !autosave ? '저장 안 됨' : '이 브라우저에 저장됨';
  else if (serverState.saving) state = '저장 중...';
  else if (serverState.error) state = '서버에 저장하지 못함';
  else if (dirty) state = autosave ? '저장 대기 중' : '저장 안 됨';
  else state = '저장됨';
  dom.saveState.textContent = `· ${state}`;
  dom.saveState.title = server.available ? '서버에 저장되어 다른 기기에서도 열 수 있습니다' : '서버 없이 실행 중 (이 브라우저에만 저장)';
}

function renderAll() {
  if (si >= wb.sheets.length) si = wb.sheets.length - 1;
  if (chartSel && !findObject(sheet(), chartSel) && drag?.type !== 'draw') chartSel = null;
  gv.layout();
  renderSheetTabs();
  updateTitle();
  updateSelectionUI();
  setMode();
}

let renderQueued = false;
function onBookChange() {
  dirty = true;
  if (!renderQueued) {
    renderQueued = true;
    queueMicrotask(() => {
      renderQueued = false;
      renderAll();
    });
  }
  scheduleAutosave();
}

// ───────────────────────── 이벤트 연결 ─────────────────────────
function bindEvents() {
  const ed = dom.editor;
  ed.addEventListener('keydown', onEditorKeyDown);
  ed.addEventListener('keyup', handleKeytipUp);
  document.addEventListener('mousedown', endKeytip, true);
  window.addEventListener('blur', endKeytip);
  // 그림 개체를 선택한 채로 입력하면 셀 대신 도형 글자로 (차트·그림은 무시)
  const objectTyped = () => {
    const v = ed.value;
    ed.value = '';
    if (v && sheet().shapes?.some((x) => x.id === chartSel)) shapeDialog(chartSel, v);
  };
  ed.addEventListener('compositionstart', () => { if (!editing && !chartSel && !keytip) beginTyping(); });
  ed.addEventListener('compositionend', () => {
    if (!editing && keytip) { ed.value = ''; return; }
    if (!editing && chartSel) { objectTyped(); return; }
    const k = pendingKey;
    if (!k) return;
    setTimeout(() => {
      if (pendingKey !== k || !editing) return;
      pendingKey = null;
      onEditingKey({ ...k, preventDefault() {}, isComposing: false, keyCode: 0 });
    }, 30);
  });
  ed.addEventListener('input', (e) => {
    if (!editing && keytip) { if (!e.isComposing) ed.value = ''; return; }
    if (!editing && chartSel) { if (!e.isComposing) objectTyped(); return; }
    if (!editing) beginTyping();
    editing.point = null;
    dom.formula.value = ed.value;
    afterEditInput();
  });
  ed.addEventListener('mousedown', (e) => {
    e.stopPropagation();
    if (editing) { editing.mode = 'edit'; editing.point = null; setTimeout(afterCaretMove); }
  });

  const fb = dom.formula;
  fb.addEventListener('focus', () => {
    if (!editing) startEdit('edit', null, { fromBar: true });
    else editing.fromBar = true;
  });
  fb.addEventListener('keydown', (e) => { if (editing) onEditingKey(e); });
  fb.addEventListener('input', () => {
    if (!editing) startEdit('edit', null, { fromBar: true });
    editing.point = null;
    ed.value = fb.value;
    afterEditInput();
  });
  fb.addEventListener('click', () => { if (editing) { editing.point = null; afterCaretMove(); } });
  fb.addEventListener('blur', () => { if (editing) editing.fromBar = false; });
  dom.fxCancel.addEventListener('mousedown', (e) => { e.preventDefault(); cancelEdit(); });
  dom.fxEnter.addEventListener('mousedown', (e) => { e.preventDefault(); commitEdit(); });
  $('fxInsert').addEventListener('mousedown', (e) => { e.preventDefault(); run('insertFunction'); });

  dom.view.addEventListener('mousedown', onViewMouseDown);
  dom.view.addEventListener('dblclick', onViewDblClick);
  dom.view.addEventListener('mousemove', onViewMouseMove);
  dom.view.addEventListener('mouseleave', () => { dom.tip.style.display = 'none'; });
  dom.view.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const objEl = e.target.closest('.obj');
    if (objEl) { objectMenu(objEl.dataset.id, { x: e.clientX, y: e.clientY }); return; }
    const hit = gv.hitTest(e.clientX, e.clientY);
    const kind = hit.zone === 'colHeader' ? 'col' : hit.zone === 'rowHeader' ? 'row' : 'cell';
    showContextMenu({ x: e.clientX, y: e.clientY }, kind);
  });
  document.addEventListener('mousemove', (e) => {
    lastMouse = { x: e.clientX, y: e.clientY };
    lastShift = e.shiftKey;
    if (drag) onDragMove(e.clientX, e.clientY);
  });
  document.addEventListener('mouseup', onDragEnd);

  // 클립보드
  document.addEventListener('copy', (e) => {
    if (editing || document.activeElement !== dom.editor || chartSel) return;
    objClip = null;
    e.preventDefault();
    const { text, html } = copySelection(false);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('cut', (e) => {
    if (editing || document.activeElement !== dom.editor || chartSel) return;
    objClip = null;
    e.preventDefault();
    const { text, html } = copySelection(true);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('paste', (e) => {
    if (editing || document.activeElement !== dom.editor) return;
    e.preventDefault();
    const text = e.clipboardData.getData('text/plain');
    const image = [...(e.clipboardData.files ?? [])].find((f) => f.type.startsWith('image/'));
    if (image && !text) { addImageFile(image); return; }
    if (objClip && (!text || text === objClip.text)) { pasteObject(); return; }
    handlePaste(text);
  });
  // 그림 파일을 끌어다 놓기
  dom.view.addEventListener('dragover', (e) => { if ([...(e.dataTransfer?.items ?? [])].some((i) => i.kind === 'file')) e.preventDefault(); });
  dom.view.addEventListener('drop', (e) => {
    const file = [...(e.dataTransfer?.files ?? [])][0];
    if (!file) return;
    e.preventDefault();
    if (file.type.startsWith('image/')) {
      const hit = gv.hitTest(e.clientX, e.clientY);
      addImageFile(file, { x: Math.round(hit.sheetX), y: Math.round(hit.sheetY) });
    } else if (/\.(xlsx|xlsm|csv|tsv|txt|tabula|json)$/i.test(file.name)) {
      openFileObject(file, 'open');
    }
  });

  // 이름 상자
  dom.nameBox.addEventListener('focus', () => dom.nameBox.select());
  dom.nameBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); if (gotoRef(dom.nameBox.value)) focusGrid(); }
    if (e.key === 'Escape') { dom.nameBox.value = cellName(active.r, active.c); focusGrid(); }
  });
  dom.nameBox.addEventListener('blur', () => { dom.nameBox.value = cellName(active.r, active.c); });

  dom.search.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') { findState.text = dom.search.value; findNext(dom.search.value); }
    if (e.key === 'Escape') { dom.search.value = ''; focusGrid(); }
  });

  dom.title.addEventListener('click', () => {
    if (dom.title.querySelector('input')) return;
    const input = el('input', { value: docName });
    dom.title.replaceChildren(input);
    input.focus();
    input.select();
    let cancelled = false;
    input.addEventListener('keydown', (e) => {
      e.stopPropagation();
      if (e.key === 'Enter') input.blur();
      if (e.key === 'Escape') { cancelled = true; input.blur(); }
    });
    input.addEventListener('blur', () => {
      const name = input.value.trim();
      dom.title.textContent = '';
      if (!cancelled && name && name !== docName) { renameDoc(name); dirty = true; scheduleAutosave(); }
      updateTitle();
      focusGrid();
    }, { once: true });
  });

  dom.autosave.addEventListener('click', () => {
    autosave = !autosave;
    if (autosave) saveNow(false);
    try {
      const data = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null');
      if (data) { data.autosave = autosave; localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }
    } catch { /* 무시 */ }
    updateTitle();
    toast(autosave ? '자동 저장이 켜졌습니다.' : '자동 저장이 꺼졌습니다. Ctrl+S로 저장하세요.');
    focusGrid();
  });

  document.querySelectorAll('[data-cmd]').forEach((b) => {
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', () => run(b.dataset.cmd));
  });
  dom.zoomSlider.addEventListener('input', () => setZoom(Number(dom.zoomSlider.value)));
  dom.zoomSlider.addEventListener('change', focusGrid);
  dom.fileInput.addEventListener('change', onFilePicked);

  document.addEventListener('keydown', (e) => {
    if (e.target === dom.editor || e.target === dom.formula) return;
    if (e.altKey && (e.key === 'q' || e.key === 'Q')) { e.preventDefault(); dom.search.focus(); return; }
    if (isDialogOpen() || document.querySelector('.backstage')) return;
    if (isMenuOpen() && e.key === 'Escape') { closeMenus(); focusGrid(); return; }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's') { e.preventDefault(); run('save'); }
  });
  document.addEventListener('keydown', (e) => {
    if (e.altKey && (e.key === 'q' || e.key === 'Q') && e.target === dom.editor && !editing) { e.preventDefault(); dom.search.focus(); }
  }, true);

  setMenuCloseHandler(focusGrid);
  setDialogCloseHandler(focusGrid);
  window.addEventListener('beforeunload', (e) => {
    saveToStorage();
    if (!autosave && dirty) { e.preventDefault(); e.returnValue = ''; }
  });
  window.addEventListener('blur', () => { if (drag) onDragEnd(); });
  window.addEventListener('afterprint', () => { dom.printArea.innerHTML = ''; });
}

// ───────────────────────── 시작 ─────────────────────────
async function init() {
  const stored = loadFromStorage();
  wb = new Workbook(stored?.workbook);
  if (stored) {
    docName = stored.docName || docName;
    si = clamp(stored.si || 0, 0, wb.sheets.length - 1);
    autosave = stored.autosave !== false;
  }
  wb.onChange(onBookChange);
  hydrateIcons();
  gv = new GridView({
    state: () => ({
      wb, si, sel, selKind, active, editing: !!editing, clip, fillPreview, refs: editRefs, chartSel, circles,
      showGrid: view.showGrid, showFormulas: view.showFormulas, showHeaders: view.showHeaders,
    }),
    onViewScroll: () => positionEditor(),
    slicerModel,
    onZoomWheel: (d) => setZoom(view.zoom + d),
    isDragging: () => !!drag,
  });
  ribbon = buildRibbon({ run, openMenu: openNamedMenu, focusGrid, refreshRibbon: updateRibbon });
  bindEvents();
  applyView();
  renderAll();
  const f = sheet().freeze;
  selectCell(f?.rows || 0, f?.cols || 0);
  focusGrid();
  window.tabula = {
    wb: () => wb, run, selectCell, selectRange, gv: () => gv,
    get active() { return active; }, get sel() { return sel; }, get si() { return si; }, get chartSel() { return chartSel; },
  };
  // 서버 저장소 (npm start 로 실행한 경우) — 다른 기기와 문서 공유
  if (await server.init()) {
    if (!stored) {
      try {
        const files = await server.list();
        if (files.length && !dirty) await openFromServer(files[0].name);
      } catch { /* 무시 */ }
    } else if (autosave) scheduleServerSave(500);
  }
  updateTitle();
}

init();
