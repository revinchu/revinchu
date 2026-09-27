// Tabula 메인: 상태 · 선택 · 편집 · 키보드/마우스 · 명령 (그리기는 view.js)
import { Workbook, cellData, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import {
  cellName, colToName, parseRangeName, parse, shiftFormula, listRefs, normalizeFormula,
  FUNCTION_NAMES, isError, quoteSheetName, MAX_ROWS, MAX_COLS,
} from './formula.js';
import { formatValue, displayedDecimals, parseInput } from './format.js';
import { buildRibbon, FONTS, FONT_SIZES } from './ribbon.js';
import {
  el, hydrateIcons, toast, openMenu, closeMenus, isMenuOpen, openDialog, alertDialog,
  formDialog, setMenuCloseHandler, setDialogCloseHandler, isDialogOpen,
} from './ui.js';
import { FUNC_INFO, CATEGORIES } from './funcinfo.js';
import { makeSeries } from './series.js';
import { parseDelimited, toDelimited, guessDelimiter } from './csv.js';
import { SAMPLES } from './samples.js';
import { GridView, DEFAULT_FONT, DEFAULT_SIZE, measureText, fontStack } from './view.js';
import { readXlsx, writeXlsx } from './xlsx.js';
import { CHART_TYPES, chartData, renderChartSvg } from './chart.js';
import { buildPivot, AGGREGATES } from './pivot.js';
import { server } from './storage.js';

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
let chartSel = null;
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
const filterHidden = (r) => !!sheet().filter?.hidden?.[r];

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
  if (document.activeElement !== dom.nameBox) dom.nameBox.value = chartSel ? '차트' : cellName(active.r, active.c);
  if (!editing) dom.formula.value = chartSel ? '' : wb.getCell(si, active.r, active.c)?.raw ?? '';
  dom.fxCancel.disabled = !editing;
  dom.fxEnter.disabled = !editing;
  gv.renderSelection();
  positionEditor();
  updateStats();
  updateRibbon();
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
    const f = sheet().filter;
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
  endEditUI();
  const multi = fillSel && !selIsActiveOnly();
  if (text !== original || multi) {
    wb.transact(() => {
      if (multi) {
        for (const [rr, cc] of cellsIn(sel)) wb.setInput(si, rr, cc, text.startsWith('=') ? shiftFormula(text, rr - r, cc - c) : text);
        autoFitRows(sel.r1, Math.min(sel.r2, sel.r1 + 500));
      } else {
        wb.setInput(si, r, c, text);
        if (text.includes('\n')) wb.setStyle(si, r, c, { wrap: true });
        else if (!text.startsWith('=') || typeof valueAt(r, c) === 'number') autoWiden({ r1: r, c1: c, r2: r, c2: c });
        autoFitRows(r, r);
      }
    }, meta());
  }
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
function autoWiden(rg) {
  const s = sheet();
  const u = usedClip(rg);
  for (let c = u.c1; c <= Math.min(u.c2, u.c1 + 200); c++) {
    if (s.colWidths[c] !== undefined) continue;
    let need = 0;
    for (let r = u.r1; r <= Math.min(u.r2, u.r1 + 2000); r++) {
      const v = valueAt(r, c);
      const st = styleAt(r, c);
      if (typeof v !== 'number' || st.wrap || wb.mergeAt(si, r, c)) continue;
      need = Math.max(need, measureText(formatValue(v, st).text, st) + 10);
    }
    if (need > DEFAULT_COL_WIDTH + 1 && need < 400) wb.setColWidth(si, c, Math.ceil(need));
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
  if (e.isComposing || e.keyCode === 229) return;
  const ctrl = e.ctrlKey || e.metaKey;
  const k = e.key;
  const handled = () => e.preventDefault();

  if (chartSel) {
    if (k === 'Delete' || k === 'Backspace') { handled(); deleteChart(chartSel); return; }
    if (k === 'Escape') { handled(); deselectChart(); updateSelectionUI(); return; }
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
      9: 'hideRows', 0: 'hideCols',
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
function startAutoScroll() {
  stopAutoScroll();
  autoScrollTimer = setInterval(() => {
    if (!drag || drag.type === 'colResize' || drag.type === 'rowResize' || drag.type === 'chart') return stopAutoScroll();
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
    openFilterMenu(Number(t.dataset.c), t);
    return;
  }
  const chartEl = t.closest('.chart');
  if (chartEl) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    focusGrid();
    const id = chartEl.dataset.id;
    if (chartSel !== id) { chartSel = id; gv.renderObjectsAll(); updateSelectionUI(); }
    if (e.button !== 0) return;
    const ch = sheet().charts.find((x) => x.id === id);
    const corner = t.classList.contains('ch-h') ? [...t.classList].find((c) => ['nw', 'ne', 'sw', 'se'].includes(c)) : null;
    drag = { type: 'chart', id, corner, start: { x: e.clientX, y: e.clientY }, orig: { x: ch.x, y: ch.y, w: ch.w, h: ch.h } };
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
  if (t.closest?.('.chart') || t.classList?.contains('fbtn')) return;
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
    case 'chart': {
      const ch = sheet().charts.find((xx) => xx.id === drag.id);
      if (!ch) break;
      const dx = (x - drag.start.x) / gv.z;
      const dy = (y - drag.start.y) / gv.z;
      const o = drag.orig;
      if (!drag.corner) {
        ch.x = Math.max(0, Math.round(o.x + dx));
        ch.y = Math.max(0, Math.round(o.y + dy));
      } else {
        const k = drag.corner;
        let { x: nx, y: ny, w: nw, h: nh } = o;
        if (k.includes('e')) nw = o.w + dx;
        if (k.includes('s')) nh = o.h + dy;
        if (k.includes('w')) { nw = o.w - dx; nx = o.x + dx; }
        if (k.includes('n')) { nh = o.h - dy; ny = o.y + dy; }
        if (nw >= 120) { ch.w = Math.round(nw); ch.x = Math.max(0, Math.round(nx)); }
        if (nh >= 90) { ch.h = Math.round(nh); ch.y = Math.max(0, Math.round(ny)); }
      }
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
    case 'chart': {
      if (!d.moved) break;
      const charts = sheet().charts;
      const ch = charts.find((x) => x.id === d.id);
      const final = { ...ch };
      Object.assign(ch, d.orig);
      wb.transact(() => wb.setSheetProp(si, 'charts', charts.map((x) => (x.id === d.id ? final : { ...x }))), meta());
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
  const chartEl = t.closest('.chart');
  if (chartEl) { chartDialog(chartEl.dataset.id); return; }
  if (t.classList.contains('fbtn') || t === dom.editor) return;
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
        for (const line of text.split('\n')) w = Math.max(w, measureText(line, st) + (sheet().filter?.r1 === r ? 28 : 10));
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
  const patchFor = (cur) => (typeof patchOrFn === 'function' ? patchOrFn(cur) : patchOrFn);
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
    if (widen) autoWiden(rg);
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

function sortData(ascending, keyCol = active.c, header = null, rgIn = null) {
  const f = sheet().filter;
  let rg = rgIn;
  if (!rg && f && active.r >= f.r1 && active.r <= f.r2 && active.c >= f.c1 && active.c <= f.c2 && selIsActiveOnly()) {
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
    if (f && rg.r1 === f.r1 && rg.c1 === f.c1) wb.setSheetProp(si, 'filter', recomputeFilter({ ...f, sort: { col: key, asc: ascending } }));
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
function recomputeFilter(f) {
  const region = currentRegion(f.r1, f.c1);
  const r2 = Math.max(f.r2, region.r2);
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

function applyFilterCriteria(c, values) {
  const f = sheet().filter;
  if (!f) return;
  const criteria = { ...f.criteria };
  if (values === null) delete criteria[c]; else criteria[c] = values;
  const nf = recomputeFilter({ ...f, criteria });
  wb.transact(() => wb.setSheetProp(si, 'filter', nf), meta());
  gv.layout();
  const total = nf.r2 - nf.r1;
  toast(`${total}개 중 ${total - Object.keys(nf.hidden).length}개의 레코드가 있습니다.`);
  selectCell(nf.r1, c);
  setMode();
}

function openFilterMenu(c, anchorEl) {
  const f = sheet().filter;
  if (!f) return;
  const full = recomputeFilter(f);
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
    applyFilterCriteria(c, chosen.length === items.length ? null : chosen);
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
    { label: '텍스트 오름차순 정렬', icon: 'sortAsc', action: () => sortData(true, c, true, full) },
    { label: '텍스트 내림차순 정렬', icon: 'sortDesc', action: () => sortData(false, c, true, full) },
    { sep: true },
    { label: `"${header}"에서 필터 해제`, icon: 'filterClear', disabled: !current, action: () => applyFilterCriteria(c, null) },
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
  const chart = { id, type, title: pieTitle || '차트 제목', range: { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, x, y, w: 480, h: 288 };
  wb.transact(() => wb.setSheetProp(si, 'charts', [...sheet().charts.map((c) => ({ ...c })), chart]), meta());
  chartSel = id;
  gv.ensureVisible(gv.rows.indexAt(y + 100), gv.cols.indexAt(x + 200));
  gv.renderObjectsAll();
  updateSelectionUI();
}

function updateChart(id, patch) {
  wb.transact(() => wb.setSheetProp(si, 'charts', sheet().charts.map((c) => (c.id === id ? { ...c, ...patch } : { ...c }))), meta());
}

function deleteChart(id) {
  wb.transact(() => wb.setSheetProp(si, 'charts', sheet().charts.filter((c) => c.id !== id).map((c) => ({ ...c }))), meta());
  chartSel = null;
  updateSelectionUI();
}

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

function chartMenu(id, pos) {
  openMenu(pos, [
    { label: '차트 편집...', icon: 'chartColumn', action: () => chartDialog(id) },
    { title: '차트 종류 변경' },
    ...CHART_TYPES.map((t) => ({ label: t.label, checked: sheet().charts.find((c) => c.id === id)?.type === t.id, action: () => updateChart(id, { type: t.id }) })),
    { sep: true },
    { label: '삭제', icon: 'delete', key: 'Delete', action: () => deleteChart(id) },
  ]);
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

function writePivot(targetSi, def) {
  const rows = pivotSourceRows(def.source, def.range);
  if (!rows) return false;
  const out = buildPivot(rows, def);
  const t = wb.sheets[targetSi];
  for (const k of [...t.cells.keys()]) { const [r, c] = k.split(',').map(Number); wb.setCellData(targetSi, r, c, null); }
  out.forEach((row, r) => row.forEach((d, c) => { if (d) wb.setCellData(targetSi, r, c, d); }));
  const colsN = out[0]?.length ?? 0;
  for (let c = 0; c < colsN; c++) {
    let w = 0;
    out.forEach((row, r) => {
      if (row[c]?.raw) w = Math.max(w, measureText(displayText(r, c, targetSi), wb.styleAt(targetSi, r, c)) + 12);
    });
    if (w > DEFAULT_COL_WIDTH) wb.setColWidth(targetSi, c, Math.min(300, Math.ceil(w)));
  }
  return true;
}

function pivotDialog() {
  const rg = dataRange();
  if (rg.r2 <= rg.r1) { alertDialog('피벗 테이블', '머리글 행과 데이터가 있는 범위를 선택하세요.'); return; }
  const srcName = sheet().name;
  const headers = [];
  for (let c = rg.c1; c <= rg.c2; c++) headers.push({ value: String(c - rg.c1), label: displayText(rg.r1, c) || `${colToName(c)}열` });
  const firstNum = headers.findIndex((h, i) => typeof valueAt(rg.r1 + 1, rg.c1 + i) === 'number');
  formDialog('피벗 테이블 만들기', [
    { name: 'range', label: '표/범위', value: `${quoteSheetName(srcName)}!${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}` },
    { name: 'row', label: '행', type: 'select', value: '0', options: headers },
    { name: 'col', label: '열', type: 'select', value: '', options: [{ value: '', label: '(없음)' }, ...headers] },
    { name: 'val', label: '값', type: 'select', value: firstNum >= 0 ? String(firstNum) : '', options: [{ value: '', label: '(행 개수)' }, ...headers] },
    { name: 'agg', label: '계산', type: 'select', value: firstNum >= 0 ? 'sum' : 'count', options: AGGREGATES.map((a) => ({ value: a.id, label: a.label })) },
  ], (v) => {
    const bang = v.range.lastIndexOf('!');
    const src = bang > 0 ? v.range.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : srcName;
    const prg = parseRangeName(bang > 0 ? v.range.slice(bang + 1) : v.range);
    if (!prg || wb.sheetIndexByName(src) < 0) { toast('범위가 올바르지 않습니다.'); return false; }
    const def = {
      source: src, range: prg, rowField: Number(v.row),
      colField: v.col === '' ? null : Number(v.col), valueField: v.val === '' ? null : Number(v.val), agg: v.val === '' ? 'count' : v.agg,
    };
    let n = 1;
    while (wb.sheetIndexByName(`피벗${n}`) >= 0) n++;
    const at = wb.transact(() => {
      const idx = wb.addSheet(`피벗${n}`, si + 1);
      writePivot(idx, def);
      wb.setSheetProp(idx, 'pivot', def);
      return idx;
    }, meta());
    switchSheet(at, false);
    toast('피벗 테이블을 만들었습니다. 원본이 바뀌면 [데이터 → 모두 새로 고침]을 누르세요.');
    return true;
  }, { note: '원본 데이터의 첫 행은 머리글이어야 합니다. 결과는 새 워크시트에 만들어집니다.' });
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
  try {
    const bytes = writeXlsx(wb, { activeSheet: si });
    download(`${safeFileName(name)}.xlsx`, new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
    toast('Excel 통합 문서(.xlsx)로 저장했습니다.');
  } catch (err) {
    alertDialog('Tabula', `저장하지 못했습니다: ${err.message}`);
  }
}

function saveAs() {
  formDialog('다른 이름으로 저장', [
    { name: 'name', label: '파일 이름', value: docName },
    {
      name: 'type', label: '파일 형식', type: 'select', value: 'xlsx', options: [
        { value: 'xlsx', label: 'Excel 통합 문서 (*.xlsx)' },
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
  if (!file) return;
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

function saveToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ docName, si, autosave, workbook: wb.serialize() }));
    if (!server.available) dirty = false;
    updateTitle();
    return true;
  } catch {
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
  saveToStorage();
  if (!server.available) {
    if (explicit) toast('이 브라우저에 저장했습니다. (서버 없이 실행 중)');
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

function formatCellsDialog() {
  const st = styleAt(active.r, active.c);
  formDialog('셀 서식', [
    { name: 'font', label: '글꼴', type: 'select', value: st.font || DEFAULT_FONT, options: FONTS.map((f) => ({ value: f, label: f })) },
    { name: 'size', label: '크기', type: 'number', value: st.size || DEFAULT_SIZE },
    { name: 'bold', label: '굵게', type: 'checkbox', value: st.bold },
    { name: 'italic', label: '기울임꼴', type: 'checkbox', value: st.italic },
    { name: 'underline', label: '밑줄', type: 'checkbox', value: st.underline },
    { name: 'strike', label: '취소선', type: 'checkbox', value: st.strike },
    { name: 'color', label: '글꼴 색', type: 'color', value: st.color || '#000000' },
    { name: 'fill', label: '채우기 색', type: 'color', value: st.fill || '#ffffff' },
  ], (v) => {
    applyStyle({
      font: v.font === DEFAULT_FONT ? undefined : v.font,
      size: Number(v.size) === DEFAULT_SIZE || !Number(v.size) ? undefined : Number(v.size),
      bold: v.bold || undefined, italic: v.italic || undefined, underline: v.underline || undefined, strike: v.strike || undefined,
      color: v.color === '#000000' ? undefined : v.color,
      fill: v.fill === '#ffffff' ? undefined : v.fill,
    });
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
  ['Ctrl+방향키', '데이터 영역의 끝으로 이동 (빈 열에서는 1,048,576행까지)'],
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
  ['Ctrl+1', '셀 서식'],
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

const TABLE_THEMES = [
  { name: '파랑, 표 스타일 보통 2', head: '#4472c4', band: '#d9e1f2' },
  { name: '주황, 표 스타일 보통 3', head: '#ed7d31', band: '#fce4d6' },
  { name: '회색, 표 스타일 보통 4', head: '#a5a5a5', band: '#ededed' },
  { name: '녹색, 표 스타일 보통 7', head: '#70ad47', band: '#e2efda' },
  { name: '진한 파랑, 표 스타일 진하게 2', head: '#203864', band: '#b4c6e7' },
];

function applyTableStyle(theme) {
  const rg = dataRange();
  wb.transact(() => {
    for (const [r, c] of cellsIn(rg)) {
      if (r === rg.r1) wb.setStyle(si, r, c, { fill: theme.head, color: '#ffffff', bold: true, bb: true });
      else wb.setStyle(si, r, c, { fill: (r - rg.r1) % 2 === 1 ? theme.band : undefined, color: undefined, bold: undefined, ...(r === rg.r2 && { bb: true }) });
    }
    if (!sheet().filter) {
      wb.setSheetProp(si, 'filter', { ...rg, criteria: {}, hidden: {} });
      widenForFilterButtons(rg);
    }
  }, meta());
  selectRange(rg, 'cells', active);
}

function tableStylesMenu(anchorEl) {
  const chip = (t) => el('button', {
    class: 'style-chip', title: t.name, onmousedown: (e) => e.preventDefault(),
    style: { background: `linear-gradient(${t.head} 0 30%, ${t.band} 30% 55%, #fff 55% 78%, ${t.band} 78%)`, height: '40px', width: '60px' },
    onclick: () => { closeMenus(); applyTableStyle(t); focusGrid(); },
  });
  openMenu(anchorEl, [{ title: '밝게 / 보통' }, { node: el('div', { class: 'style-grid', style: { gridTemplateColumns: 'repeat(5, 60px)' } }, TABLE_THEMES.map(chip)) }]);
}

const MENUS = {
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
      { label: '중복 값...', action: () => condRuleDialog('dup') },
      { title: '상위/하위 규칙' },
      { label: '상위 10개 항목...', action: () => condRuleDialog('top') },
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
      { sep: true },
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
    { label: '필터', icon: 'filter', key: 'Ctrl+Shift+L', checked: !!sheet().filter, action: () => toggleFilter() },
    { label: '지우기', icon: 'filterClear', disabled: !sheet().filter, action: () => run('clearFilter') },
    { label: '다시 적용', icon: 'refresh', disabled: !sheet().filter, action: () => run('reapplyFilter') },
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
      { label: '필터', icon: 'filter', checked: !!sheet().filter, action: () => toggleFilter() },
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

  numFmt: (f) => applyStyle({ numFmt: f === 'general' ? undefined : f, decimals: undefined }, { widen: true }),
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
      data.sheets[at] = { ...src, name: wb.sheets[at].name, charts: (src.charts ?? []).map((c) => ({ ...c, id: `${c.id}d` })) };
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
    if (Object.keys(sheet().filter?.hidden ?? {}).length) {
      for (const [r, c] of cellsIn(usedClip(sel))) {
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
    const f = sheet().filter;
    if (!f) return;
    wb.transact(() => wb.setSheetProp(si, 'filter', { ...f, criteria: {}, hidden: {}, sort: undefined }), meta());
    setMode();
  },
  reapplyFilter: () => {
    const f = sheet().filter;
    if (!f) return;
    wb.transact(() => wb.setSheetProp(si, 'filter', recomputeFilter(f)), meta());
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
      el('span', { class: 'muted' }, `시트 크기 1,048,576행 × 16,384열 · 함수 ${FUNCTION_NAMES.length}개 · .xlsx 열기/저장`), el('br'),
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
    font: st.font || DEFAULT_FONT, size: String(st.size || DEFAULT_SIZE), numFmt: fmt,
    alignLeft: st.align === 'left', alignCenter: st.align === 'center', alignRight: st.align === 'right',
    valignTop: st.valign === 'top', valignMiddle: st.valign === 'middle', valignBottom: !st.valign,
    merged: !!wb.mergeAt(si, active.r, active.c), painter: !!painter, filterOn: !!sheet().filter,
    frozen: !!(f.rows || f.cols), lastFill, lastFont, ...view,
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
  if (chartSel && !sheet().charts.some((c) => c.id === chartSel)) chartSel = null;
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
  ed.addEventListener('compositionstart', () => { if (!editing) beginTyping(); });
  ed.addEventListener('compositionend', () => {
    const k = pendingKey;
    if (!k) return;
    setTimeout(() => {
      if (pendingKey !== k || !editing) return;
      pendingKey = null;
      onEditingKey({ ...k, preventDefault() {}, isComposing: false, keyCode: 0 });
    }, 30);
  });
  ed.addEventListener('input', () => {
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
    const chartEl = e.target.closest('.chart');
    if (chartEl) { chartMenu(chartEl.dataset.id, { x: e.clientX, y: e.clientY }); return; }
    const hit = gv.hitTest(e.clientX, e.clientY);
    const kind = hit.zone === 'colHeader' ? 'col' : hit.zone === 'rowHeader' ? 'row' : 'cell';
    showContextMenu({ x: e.clientX, y: e.clientY }, kind);
  });
  document.addEventListener('mousemove', (e) => {
    lastMouse = { x: e.clientX, y: e.clientY };
    if (drag) onDragMove(e.clientX, e.clientY);
  });
  document.addEventListener('mouseup', onDragEnd);

  // 클립보드
  document.addEventListener('copy', (e) => {
    if (editing || document.activeElement !== dom.editor || chartSel) return;
    e.preventDefault();
    const { text, html } = copySelection(false);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('cut', (e) => {
    if (editing || document.activeElement !== dom.editor || chartSel) return;
    e.preventDefault();
    const { text, html } = copySelection(true);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('paste', (e) => {
    if (editing || document.activeElement !== dom.editor) return;
    e.preventDefault();
    handlePaste(e.clipboardData.getData('text/plain'));
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
      wb, si, sel, selKind, active, editing: !!editing, clip, fillPreview, refs: editRefs, chartSel,
      showGrid: view.showGrid, showFormulas: view.showFormulas, showHeaders: view.showHeaders,
    }),
    onViewScroll: () => positionEditor(),
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
