// WIXEL 메인: 상태 · 선택 · 편집 · 키보드/마우스 · 명령 (그리기는 view.js)
import { Workbook, formulaShifter, cellData, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import {
  cellName, colToName, parseRangeName, parse, shiftFormula, listRefs, normalizeFormula, tokenize,
  FUNCTION_NAMES, isError, quoteSheetName, MAX_ROWS, MAX_COLS,
} from './formula.js';
import {
  formatValue, displayedDecimals, parseInput, formatCode, styleForCode, codeOfStyle, adjustCodeDecimals, formatGeneral,
} from './format.js';
import { buildRibbon, FONTS, FONT_SIZES, TABS } from './ribbon.js';
import { flashFill } from './flashfill.js';
import {
  el, hydrateIcons, toast, openMenu, closeMenus, isMenuOpen, openDialog, alertDialog,
  formDialog, setMenuCloseHandler, setDialogCloseHandler, isDialogOpen,
} from './ui.js';
import { FUNC_INFO, CATEGORIES } from './funcinfo.js';
import { makeSeries, CUSTOM_LISTS } from './series.js';
import { parseDelimited, toDelimited, guessDelimiter, CsvBlockReader } from './csv.js';
import { SAMPLES } from './samples.js';
import { TEMPLATES, TEMPLATE_CATS } from './templates.js';
import { GridView, BASE_FONT, setBaseFont, measureText, fontStack, PATTERNS, patternCss, glyphShift, timelinePeriods } from './view.js';
import { setThemeColors } from './stylepresets.js';
import { readXlsxAsync, writeXlsxAsync, xlsxOverflow } from './xlsx.js';
import { readOds, writeOds } from './ods.js';
import { CHART_TYPES, CHART_GALLERY, CHART_PALETTES, PALETTE, paletteOf, renderChartSvg, chartModelData } from './chart.js';
import {
  computePivot, warmPivots, AGGREGATES, SHOW_AS, BASE_POS, LAYOUTS, pivotSourceData, resolvePivot, itemText, headerNames, normalizeDef, valueName,
  pivotFieldNames, parseCalc, PIVOT_STYLES, PIVOT_STYLE_GROUPS, pivotStyleParts, LABEL_OPS, VALUE_OPS, describeFieldFilter, keyOf, sortKeys, pivotDetail, GROUP_BY,
  checkCalc, renameCalcRefs, CALC_FUNCS,
} from './pivot.js';
import { SLICER_STYLES, SLICER_STYLE_GROUPS, slicerStyleName, slicerColors, CUSTOM_KEYS } from './slicerstyle.js';
import { server, idbSet, idbGet, idbDel } from './storage.js';
import { libList, libSave, libLoad, libLoadVersion, libUpdate, libNameVersion, libRemove, newDocId, packText, unpackText, LIB_MAX, VER_MAX } from './library.js';
import { itemStats, blockColumn, EMPTY as PIVOT_EMPTY } from './cube.js';
import { logicalCol, ColBuilder } from './block.js';
import { PROTECT_OPTIONS, defaultAllow, excelHash, isProtected, isLockedStyle, allowed } from './protect.js';
import { PAPERS, MARGINS, normPage, paperOf, printScale, headerParts } from './page.js';
import { SPARK_TYPES, sparkDefaults, sparkItems, sparkRef } from './sparkline.js';
import { evalSteps, goalSeek, dataTable, specialCells, GOTO_KINDS, valueText } from './audit.js';
import { normOutline, outlineEmpty, changeLevels, groupsOf, groupAt, toggleGroup, showLevel, summaryOf, planSubtotals, SUBTOTAL_FNS, maxLevel } from './outline.js';
import { hid, hidCount } from './axis.js';
import { fontList, fontAlias, loadLocalFonts, canListLocalFonts } from './fonts.js';
import { ICONS } from './icons.js';
import {
  CELL_OPS, TEXT_OPS, DATE_PERIODS, ICON_SETS, ICON_SVG, VISUAL_TYPES, iconSetById, describeCond, SCALE_PRESETS, BAR_PRESETS, DEFAULT_SCALE2, DEFAULT_SCALE3,
} from './condfmt.js';
import {
  TABLE_STYLES, TABLE_STYLE_GROUPS, DEFAULT_TABLE_STYLE, TOTAL_FUNCS, tableAt, tableCellStyle, tableFilterRange, dataTop, dataBottom, uniqueNames,
  nextTableName, columnNames, expansionFor, validTableName, findTable, resolveStructRef,
} from './tables.js';
import { splitDelimited, splitFixed, suggestBreaks, parseDateOrder, convertPart, DATE_ORDERS } from './textsplit.js';
import {
  VALIDATION_TYPES, VALIDATION_OPS, validationAt, checkValidation, listItems, describeRule, subtractRange, invalidCells,
} from './validation.js';
import { OBJECT_PROPS, OBJECT_LABEL, SHAPE_KINDS, SHAPE_GROUPS, LINE_SHAPES, newShape, findObject, shapeSvg } from './shapes.js';
import { extractVbaModules, fromBase64 } from './vba.js';
import { findMatches, nextMatch, replaceText } from './find.js';
import {
  ANALYSIS_TOOLS, AnalysisError, splitGroups, descriptive, matrixTool, regression, histogram, rankPercentile, tTest, zTest, fTest, anova1, anova2,
  movingAverage, expSmoothing, randomNumbers, sampling, solveMin,
} from './analysis.js';
import { timeAxis } from './ets.js';
import {
  CATEGORIES as FMT_CATEGORIES, CURRENCY_SYMBOLS, NEGATIVE_STYLES, DATE_TYPES, TIME_TYPES, FRACTION_TYPES, SPECIAL_TYPES, CUSTOM_LIST, buildCode, describeCode,
} from './fmtpresets.js';
import { maxOf, minOf } from './fxcore.js';

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
// 화면 그리기 규칙(피벗 서식 등)이 바뀔 때 올림: 예전 버전의 자동 저장본은 열 때 피벗을 다시 그림
const APP_REV = 3;
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
let special = null; // 이동 옵션으로 고른 칸들 { si, cells: [[r, c]] } (Ctrl+Enter 로 한꺼번에 입력, Delete 로 지우기)
let trace = null; // 추적 화살표 { si, arrows: [{ from: {si, r1, c1, r2, c2}, to: {si, r, c}, err }], frontP, frontD }
let keepSpecial = false;
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
/** 시트 기본 행 높이 · 열 너비 (파일에서 가져온 시트는 다를 수 있음) */
const defRowH = () => sheet().defRowH ?? DEFAULT_ROW_HEIGHT;
const defColW = () => sheet().defColW ?? DEFAULT_COL_WIDTH;
const filterHidden = (r) => hid(sheet().filter?.hidden, r) || (sheet().tables ?? []).some((t) => hid(t.filter?.hidden, r));

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
  if (!keepSpecial) special = null;
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
  if (cfSmartTag && cfSmartTag._at !== `${si}:${sel.r1},${sel.c1},${sel.r2},${sel.c2}`) hideCfSmartTag();
  const selObj = chartSel ? findObject(sheet(), chartSel) : null;
  if (document.activeElement !== dom.nameBox) dom.nameBox.value = selObj ? (selObj.obj.name || OBJECT_LABEL[selObj.prop]) : nameBoxLabel();
  if (!editing) {
    let raw = chartSel ? '' : wb.getCell(si, active.r, active.c)?.raw ?? '';
    // 분산된 셀: 원본 수식을 흐리게 표시 (엑셀과 같음)
    const anchor = !raw && !chartSel ? wb.spillAnchorOf(si, active.r, active.c) : null;
    if (anchor) raw = wb.getRaw(si, anchor.r, anchor.c);
    // 보호된 시트의 '수식 숨기기' 셀은 수식을 보여 주지 않음
    if (raw && isProtected(sheet()) && wb.styleAt(si, active.r, active.c).hideFormula) raw = '';
    dom.formula.value = raw;
    dom.formula.classList.toggle('ghost', !!anchor);
  }
  dom.fxCancel.disabled = !editing;
  dom.fxEnter.disabled = !editing;
  refreshPivotPane();
  gv.renderSelection();
  if (refPick) pickRef();
  positionEditor();
  updateStats();
  updateRibbon();
  updateDvPrompt();
}

function updateStats() {
  statsToken++; // 진행 중인 큰 통계 계산 중단
  dom.stats.replaceChildren();
  if (selIsActiveOnly()) return;
  const rg = usedClip(sel);
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let fmtStyle = null;
  const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
  if (area > 2000000) { bigStats(rg); return; }
  if (area > 200000) {
    // 아주 큰 선택: 열 블록은 형식화 배열을 바로, 나머지는 일반 셀만 (칸마다 값을 묻지 않음)
    const sh = sheet();
    const add = (v) => {
      if (v === null || v === '' || v === undefined) return;
      count++;
      if (typeof v === 'number') { numCount++; sum += v; }
    };
    for (const [k, cell] of sh.cells) {
      const i = k.indexOf(',');
      const r = +k.slice(0, i);
      const c = +k.slice(i + 1);
      if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2 || (!cell.raw && !cell.formula)) continue;
      const v = valueAt(r, c);
      add(v);
      if (typeof v === 'number') fmtStyle ??= styleAt(r, c);
    }
    for (const bk of sh.blocks ?? []) {
      const a1 = Math.max(rg.r1, bk.r0) - bk.r0;
      const a2 = Math.min(rg.r2, bk.r0 + bk.n - 1) - bk.r0;
      for (let c = Math.max(rg.c1, bk.c0); c <= Math.min(rg.c2, bk.c0 + bk.cols.length - 1); c++) {
        const { num, str, fmt } = bk.cols[c - bk.c0];
        const pm = bk.perm && !(a1 === 0 && a2 === bk.n - 1) ? bk.perm : null; // 일부 행만이면 정렬 순서대로
        let s0 = 0;
        for (let q = a1; q <= a2; q++) {
          const i = pm ? pm[q] : q;
          if (str && str[i] >= 0) { count++; continue; }
          if (num) { const v = num[i]; if (v === v) { count++; numCount++; s0 += v; } }
        }
        sum += s0;
        if (s0 && fmt) fmtStyle ??= fmt;
      }
    }
  } else for (const [r, c] of cellsIn(rg)) {
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

/**
 * 아주 큰 선택(수백만 칸)의 상태 표시줄 통계: 열 블록을 조각씩 훑으며 화면을 멈추지 않게 (다른 선택을 하면 중단)
 */
let statsToken = 0;
async function bigStats(rg) {
  const token = ++statsToken;
  dom.stats.append(el('span', { class: 'muted' }, '계산 중…'));
  const sh = sheet();
  let count = 0;
  let numCount = 0;
  let sum = 0;
  let fmtStyle = null;
  for (const [k, cell] of sh.cells) {
    const i = k.indexOf(',');
    const r = +k.slice(0, i);
    const c = +k.slice(i + 1);
    if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2 || (!cell.raw && !cell.formula)) continue;
    const v = valueAt(r, c);
    if (v === null || v === '') continue;
    count++;
    if (typeof v === 'number') { numCount++; sum += v; fmtStyle ??= styleAt(r, c); }
  }
  let last = performance.now();
  for (const bk of sh.blocks ?? []) {
    const a1 = Math.max(rg.r1, bk.r0) - bk.r0;
    const a2 = Math.min(rg.r2, bk.r0 + bk.n - 1) - bk.r0;
    for (let c = Math.max(rg.c1, bk.c0); c <= Math.min(rg.c2, bk.c0 + bk.cols.length - 1); c++) {
      const { num, str, fmt } = bk.cols[c - bk.c0];
      const pm = bk.perm && !(a1 === 0 && a2 === bk.n - 1) ? bk.perm : null;
      for (let q0 = a1; q0 <= a2; q0 += 1 << 20) {
        const q1 = Math.min(a2, q0 + (1 << 20) - 1);
        let s0 = 0;
        for (let q = q0; q <= q1; q++) {
          const i = pm ? pm[q] : q;
          if (str && str[i] >= 0) { count++; continue; }
          if (num) { const v = num[i]; if (v === v) { count++; numCount++; s0 += v; } }
        }
        sum += s0;
        if (s0 && fmt) fmtStyle ??= fmt;
        if (performance.now() - last > 30) {
          await yieldUI();
          if (token !== statsToken || sheet() !== sh) return; // 선택이 바뀜
          last = performance.now();
        }
      }
    }
  }
  if (token !== statsToken) return;
  dom.stats.replaceChildren();
  if (!count) return;
  const fmt = (n) => formatValue(n, fmtStyle?.numFmt && fmtStyle.numFmt !== 'general' ? fmtStyle : {}).text;
  if (numCount) dom.stats.append(el('span', {}, `평균: ${fmt(sum / numCount)}`));
  dom.stats.append(el('span', {}, `개수: ${count.toLocaleString()}`));
  if (numCount) dom.stats.append(el('span', {}, `합계: ${fmt(sum)}`));
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
  objMulti.clear();
  gv.renderObjectsAll();
}

const VIEW_CMDS = new Set(['publish', 'versionHistory', 'exportXlsx', 'saveAs', 'print', 'zoomIn', 'zoomOut', 'zoom100']);
function startEdit(mode, text = null, { fromBar = false, caret = null } = {}) {
  if (viewOnly) { toast('읽기 전용 문서입니다. [편집용 사본 만들기]를 누르면 고칠 수 있습니다.'); return; }
  if (editing) return;
  if (protectBlocked('cells', { r1: active.r, c1: active.c, r2: active.r, c2: active.c })) return;
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
    const f = allFilters().map(([, x]) => x).find((x) => hidCount(x.hidden)) ?? null;
    const hidden = f ? hidCount(f.hidden) : 0;
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
      alertDialog('WIXEL', '입력한 수식에 문제가 있습니다. 수식을 확인하세요.').then(() => { if (editing) focusGrid(); });
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
  const multi = fillSel && (!selIsActiveOnly() || special?.si === si);
  if (text !== original || multi) {
    wb.transact(() => {
      if (multi && special?.si === si) {
        // 이동 옵션으로 고른 칸마다 (수식은 칸 위치만큼 옮김)
        for (const [rr, cc] of special.cells) wb.setInput(si, rr, cc, text.startsWith('=') ? shiftFormula(text, rr - r, cc - c) : text);
      } else if (multi) {
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
    fontFamily: fontStack(st.font || BASE_FONT.name),
    fontSize: `${st.size || BASE_FONT.size}pt`,
    fontWeight: st.bold ? '700' : '400',
    fontStyle: st.italic ? 'italic' : 'normal',
    textDecoration: st.underline ? 'underline' : 'none',
    color: st.color || '#000',
    background: st.fill || '#fff',
    textAlign: st.align && st.align !== 'left' && !ed.value.startsWith('=') ? st.align : 'left',
    whiteSpace: st.wrap ? 'pre-wrap' : 'pre',
  });
  const lines = ed.value.split('\n');
  const textW = maxOf(lines.map((l) => measureText(l, st))) + 12;
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

/** 피벗 값 셀 → GETPIVOTDATA("값 필드", $피벗 왼쪽 위$, "필드", "항목", …) (엑셀의 GetPivotData 옵션) */
function getPivotDataText(r, c) {
  const e = pivotDefs().find(({ def }) => def.area && r >= def.area.r1 && r <= def.area.r2 && c >= def.area.c1 && c <= def.area.c2);
  if (!e) return null;
  const src = pivotSource(e.def);
  if (!src) return null;
  const det = pivotDetail(src, e.def, r - (e.def.top ?? 0), c - (e.def.left ?? 0));
  if (!det?.valueField) return null;
  const q = (x) => `"${String(x).replace(/"/g, '""')}"`;
  const at = `$${colToName(e.def.area.c1)}$${e.def.area.r1 + 1}`;
  return `GETPIVOTDATA(${q(det.valueField)},${at}${det.conds.map(([f, it]) => `,${q(f)},${/^-?\d+(\.\d+)?$/.test(it) ? it : q(it)}`).join('')})`;
}

function insertPointText(ref) {
  const inp = edInput();
  const text = inp.value;
  const start = editing.point ? editing.point.start : inp.selectionStart;
  const end = editing.point ? editing.point.end : inp.selectionEnd;
  setEditText(text.slice(0, start) + ref + text.slice(end), start + ref.length);
  editing.point = null;
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
  if (!editing || opts.formulaAutocomplete === false) return hideAutocomplete();
  const inp = edInput();
  const text = inp.value;
  const pos = inp.selectionStart;
  const before = text.slice(0, pos);
  if (!text.startsWith('=') || (before.split('"').length - 1) % 2 === 1) return hideAutocomplete();
  const m = /(^|[^A-Za-z0-9_."$!:])([A-Za-z][A-Za-z0-9.]*)$/.exec(before.slice(1));
  if (!m || /^[A-Za-z]{1,3}\d+$/.test(m[2])) return hideAutocomplete();
  const prefix = m[2].toUpperCase();
  const nameItems = [...wb.names.filter((n) => !n.hidden).map((n) => n.name), ...wb.sheets.flatMap((sh) => (sh.tables ?? []).map((t) => t.name))]
    .filter((n) => n.toUpperCase().startsWith(prefix));
  const items = [...nameItems, ...FUNCTION_NAMES.filter((n) => n.startsWith(prefix))].slice(0, 12);
  if (!items.length || text[pos] === '(') return hideAutocomplete();
  ac = { items, index: 0, start: pos - m[2].length, end: pos };
  dom.ac.replaceChildren(...items.map((name, i) => el('li', {
    class: i === 0 ? 'active' : '',
    onmousedown: (e) => { e.preventDefault(); e.stopPropagation(); acceptAutocomplete(name); },
    title: FUNC_INFO[name]?.desc ?? '',
  }, el('span', {}, name), el('small', {}, FUNC_INFO[name]?.cat ?? (nameItems.includes(name) ? '이름' : '')))));
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
  const isFn = !!FUNC_INFO[name] && !wb.findName(name, si);
  const value = `${text.slice(0, start)}${name}${isFn ? '(' : ''}${text.slice(end)}`;
  hideAutocomplete();
  setEditText(value, start + name.length + (isFn ? 1 : 0));
  if (isFn) toast(FUNC_INFO[name]?.sig ?? name);
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
    const need = neededRowHeight(si, r, cols);
    if ((s.rowHeights[r] ?? defRowH()) !== need) wb.setRowHeight(si, r, need, false);
  }
}

/** 행에 필요한 높이(px): 큰 글꼴 · 줄 바꿈 · 텍스트 회전(각도 · 세로 쓰기)까지 엑셀처럼 */
function neededRowHeight(sIdx, r, cols) {
  const base = wb.sheets[sIdx].defRowH ?? DEFAULT_ROW_HEIGHT;
  let need = base;
  for (let c = 0; c < cols; c++) {
    const cell = wb.getCell(sIdx, r, c);
    if (!cell?.raw || wb.mergeAt(sIdx, r, c)) continue;
    const st = wb.styleAt(sIdx, r, c);
    const text = formatValue(wb.getValue(sIdx, r, c), st).text;
    const lineH = ((st.size || BASE_FONT.size) * 4 / 3) * 1.2;
    let h;
    if (st.rotate === 255) h = [...text].length * lineH;
    else if (st.rotate) {
      const a = (Math.abs(st.rotate) * Math.PI) / 180;
      h = Math.abs(Math.sin(a)) * measureText(text, st) + Math.abs(Math.cos(a)) * lineH;
    } else {
      const lines = st.wrap ? wrappedLines(text, wb.colWidth(sIdx, c) - 7, st) : text.split('\n').length;
      h = lines * lineH;
    }
    need = Math.max(need, Math.ceil(h + 3));
  }
  return Math.min(545, need);
}

/** 파일을 열 때: 높이가 저장되지 않은 행을 내용에 맞춤 (엑셀은 이런 행을 자동 높이로 그림) — 실행 취소 기록 없음 */
function fitRowsOnOpen() {
  const list = wb.fitRows;
  wb.fitRows = null;
  if (!list) return;
  list.forEach((rows, sIdx) => {
    const s = wb.sheets[sIdx];
    if (!rows || !s) return;
    const cols = Math.min(wb.usedRange(sIdx).cols, 500);
    for (const r of rows.slice(0, 5000)) {
      if (s.rowManual[r]) continue;
      const need = neededRowHeight(sIdx, r, cols);
      if (need !== (s.defRowH ?? DEFAULT_ROW_HEIGHT)) s.rowHeights[r] = need;
    }
  });
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
    if (need > (s.colWidths[c] ?? defColW()) + 1 && need < 400) wb.setColWidth(si, c, Math.ceil(need));
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
      commitEdit(ctrl || opts.enterDir === 'none' ? null : e.shiftKey ? ENTER_OPP[opts.enterDir] ?? 'up' : opts.enterDir ?? 'down', { fillSel: ctrl });
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
    case 'F3':
      e.preventDefault();
      pasteNameDialog();
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
  // 그룹 / 그룹 해제 (Shift+Alt+→ / ←)
  if (e.altKey && e.shiftKey && !ctrl && (k === 'ArrowRight' || k === 'ArrowLeft')) { handled(); run(k === 'ArrowRight' ? 'outlineGroup' : 'outlineUngroup'); return; }
  if (e.altKey && k === 'F1') { handled(); run('chartColumn'); return; }
  // ── 엑셀 바로 가기 키 ──
  const code = e.code;
  if (ctrl && e.altKey && code === 'KeyV') { handled(); run('pasteSpecial'); return; }
  if (ctrl && e.altKey && code === 'KeyL') { handled(); run('reapplyFilter'); return; }
  if (ctrl && e.altKey && k === 'F5') { handled(); run('refreshAll'); return; }
  if (e.altKey && !ctrl && k === 'F5') { handled(); run('pivotRefresh'); return; }
  if (ctrl && e.altKey && k === 'F9') { handled(); run('recalc'); return; }
  if (e.shiftKey && !ctrl && k === 'F9') { handled(); run('calcNowSheet'); return; }
  if (e.altKey && !ctrl && (k === 'F8' || k === 'F11')) { handled(); run('macros'); return; }
  if (e.altKey && !ctrl && (k === 'PageDown' || k === 'PageUp')) { handled(); gv.scrollBy((k === 'PageDown' ? 1 : -1) * gv.viewW * 0.9, 0); return; }
  if (!ctrl && !e.altKey) {
    // Shift+F4: 다음 찾기 (엑셀과 같음)
    if (k === 'F4' && e.shiftKey) { handled(); if (findState.text || findState.format) findNext(); else run('find'); return; }
    if (k === 'F4') { handled(); repeatLast(); return; }
    if (k === 'F5') { handled(); run(e.shiftKey ? 'find' : 'goto'); return; }
    if (k === 'F3' && e.shiftKey) { handled(); run('insertFunction'); return; }
    if (k === 'F11' && !e.shiftKey) { handled(); run('chartColumn'); return; }
    if (k === 'F10' && e.shiftKey) {
      handled();
      const rect = gv.clientRect({ r1: active.r, c1: active.c, r2: active.r, c2: active.c });
      showContextMenu({ x: rect.left + 10, y: rect.bottom }, 'cell');
      return;
    }
    if (k === 'Backspace' && e.shiftKey) { handled(); selectCell(active.r, active.c); return; }
    if (k === 'F8' && !e.shiftKey) { handled(); extendMode = !extendMode; dom.status.textContent = extendMode ? '선택 영역 확장' : '준비'; return; }
    if (extendMode && k === 'Escape') { extendMode = false; setMode(); }
  }
  if (ctrl && !e.altKey) {
    if (k === 'Backspace') { handled(); gv.ensureVisible(active.r, active.c); return; }
    if (k === '.') { handled(); nextCorner(); return; }
    if (code === 'BracketLeft') { handled(); run('selectPrecedents'); return; }
    if (code === 'BracketRight') { handled(); run('selectDependents'); return; }
    if (code === 'Quote') { handled(); copyFromAbove(e.shiftKey); return; }
    if (k === 'F2') { handled(); run('print'); return; }
    if (!e.shiftKey && (k === 'e' || k === 'E')) { handled(); run('flashFill'); return; }
    if (!e.shiftKey && (k === 'k' || k === 'K')) { handled(); run('hyperlink'); return; }
    if (!e.shiftKey && (k === 'n' || k === 'N')) { handled(); run('newWorkbook'); return; }
    if (!e.shiftKey && (k === 'y' || k === 'Y')) { handled(); if (wb.canRedo()) run('redo'); else repeatLast(); return; }
    if (e.shiftKey) {
      const more = {
        Digit2: 'fmtTime', Digit6: 'fmtScientific', Minus: 'borderNone', Digit9: 'unhideRows', Digit0: 'unhideCols', Digit8: 'selectRegion',
        KeyO: 'selectComments', KeyT: 'tblTotals', KeyF: 'fontDialog', KeyP: 'fontDialog', KeyU: 'toggleFormulaBarSize', KeyE: 'flashFill',
      };
      if (more[code]) {
        handled();
        if (more[code] === 'tblTotals') { if (tableHere()) run('tblTotals'); return; }
        if (more[code] === 'toggleFormulaBarSize') { dom.formulaRow.classList.toggle('tall'); return; }
        run(more[code]);
        return;
      }
    }
  }
  if (k === 'F3' && !e.altKey) {
    handled();
    run(ctrl && e.shiftKey ? 'createNamesFromSel' : ctrl ? 'nameManager' : 'pasteName');
    return;
  }
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
    case 'ArrowUp': handled(); move(-1, 0, { extend: e.shiftKey || extendMode }); return;
    case 'ArrowDown': handled(); move(1, 0, { extend: e.shiftKey || extendMode }); return;
    case 'ArrowLeft': handled(); move(0, -1, { extend: e.shiftKey || extendMode }); return;
    case 'ArrowRight': handled(); move(0, 1, { extend: e.shiftKey || extendMode }); return;
    case 'Enter':
      handled();
      if (clip && !e.shiftKey) { pasteInternal('all'); clip = null; updateSelectionUI(); setMode(); return; }
      if (opts.enterDir !== 'none') moveEnterTab(e.shiftKey ? ENTER_OPP[opts.enterDir] ?? 'up' : opts.enterDir ?? 'down');
      return;
    case 'Tab': handled(); moveEnterTab(e.shiftKey ? 'left' : 'right'); return;
    case 'Home': handled(); if (e.shiftKey) extendTo(focusCell.r, 0); else selectCell(active.r, 0); return;
    case 'PageDown': handled(); move(gv.pageRows(), 0, { extend: e.shiftKey }); return;
    case 'PageUp': handled(); move(-gv.pageRows(), 0, { extend: e.shiftKey }); return;
    case 'Delete':
      handled();
      if (special?.si === si && !protectBlocked('cells')) { const cells = special.cells; wb.transact(() => { for (const [rr, cc] of cells) { const cur = wb.getCell(si, rr, cc); if (cur?.raw) wb.setCellData(si, rr, cc, { raw: '', style: cur.style, comment: cur.comment }); } }, meta()); return; }
      run('clearContents');
      return;
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
      if (borderDraw) setBorderDraw(borderDraw);
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


// ───────────────────────── 엑셀 바로 가기 키 도우미 ─────────────────────────
let lastRepeat = null; // F4 / Ctrl+Y: 마지막 작업 반복
let extendMode = false; // F8: 선택 영역 확장 모드
const REPEATABLE = new Set(['insertRows', 'insertCols', 'deleteRows', 'deleteCols', 'mergeCenter', 'wrap', 'clearContents', 'clearFormats', 'clearAll',
  'hideRows', 'hideCols', 'indentInc', 'indentDec', 'incDecimal', 'decDecimal', 'growFont', 'shrinkFont', 'autofitSel', 'autofitRowsSel', 'addSheet', 'fillDown', 'fillRight']);

function repeatLast() {
  if (!lastRepeat) { toast('반복할 작업이 없습니다.'); return; }
  lastRepeat();
}

/** Ctrl+' : 위 셀의 수식을 그대로(참조 이동 없이) 복사해 편집 / Ctrl+Shift+" : 위 셀의 값 */
function copyFromAbove(valueOnly) {
  if (active.r === 0) return;
  const raw = wb.getRaw(si, active.r - 1, active.c);
  const v = wb.getValue(si, active.r - 1, active.c);
  if (valueOnly) {
    const text = v === null ? '' : isError(v) ? v.code : typeof v === 'number' ? formatGeneral(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
    startEdit('enter', text);
  } else startEdit('enter', raw);
}

/** Ctrl+E 빠른 채우기 */
function flashFillCmd() {
  if (editing && !commitEdit()) return;
  const c = active.c;
  const region = currentRegion(active.r, Math.max(0, c - 1));
  const r1 = region.r1;
  const r2 = region.r2;
  const srcCols = [];
  for (let x = region.c1; x <= region.c2; x++) if (x !== c) srcCols.push(x);
  if (!srcCols.length) { toast('빠른 채우기: 옆 열에 원본 데이터가 필요합니다.'); return; }
  const text = (r, x) => { const v = wb.getValue(si, r, x); return v === null ? '' : typeof v === 'number' ? formatGeneral(v) : String(v); };
  const examples = [];
  const todo = [];
  // 머리글 행은 원본과 대상이 모두 글자일 때 건너뜀
  for (let r = r1; r <= r2; r++) {
    const sources = srcCols.map((x) => text(r, x));
    if (sources.every((x) => x === '')) continue;
    const t = wb.getRaw(si, r, c);
    if (t) examples.push({ r, sources, target: text(r, c) });
    else todo.push({ r, sources });
  }
  let ex = examples.filter((e) => e.r >= active.r - 50);
  if (ex.length > 1 && ex[0].r === r1 && todo.length) ex = ex.slice(1).length ? ex.slice(1) : ex; // 머리글 제외
  if (!ex.length || !todo.length) { toast('빠른 채우기: 채울 예시(첫 행에 원하는 결과)를 먼저 입력하세요.'); return; }
  let out = flashFill(ex, todo.map((t) => t.sources));
  if (!out && ex.length > 1) out = flashFill(ex.slice(-1), todo.map((t) => t.sources));
  if (!out || out.every((x) => x === null)) { toast('빠른 채우기로 채울 값을 찾을 수 없습니다.'); return; }
  let n = 0;
  wb.transact(() => todo.forEach((t, i) => { if (out[i] !== null && out[i] !== '') { wb.setInput(si, t.r, c, out[i]); n++; } }), meta());
  toast(`빠른 채우기: ${n}개 셀을 채웠습니다.`);
}

/** Ctrl+[ 참조되는 셀 / Ctrl+] 참조하는 셀 선택 */
function selectPrecedents() {
  const raw = wb.getRaw(si, active.r, active.c);
  const refs = listRefs(raw).filter((x) => !x.sheet || wb.sheetIndexByName(x.sheet) >= 0);
  if (!refs.length) { toast('참조되는 셀이 없습니다.'); return; }
  const target = refs[0].sheet ? wb.sheetIndexByName(refs[0].sheet) : si;
  const same = refs.filter((x) => (x.sheet ? wb.sheetIndexByName(x.sheet) : si) === target);
  const u = wb.usedRange(target);
  const rg = {
    r1: minOf(same.map((x) => x.r1)), c1: minOf(same.map((x) => x.c1)),
    r2: Math.min(maxOf(same.map((x) => x.r2)), Math.max(0, u.rows - 1)), c2: Math.min(maxOf(same.map((x) => x.c2)), Math.max(0, u.cols - 1)),
  };
  if (target !== si) switchSheet(target);
  gv.ensureVisible(rg.r1, rg.c1);
  selectRange(rg, 'cells', { r: rg.r1, c: rg.c1 });
}

function selectDependents() {
  // 의존 그래프로 바로 찾음 (수식 백만 개도 즉시)
  const hits = wb.dependentsOf(si, active.r, active.c).filter((d) => d.si === si).map((d) => [d.r, d.c]);
  if (!hits.length) { toast('참조하는 셀이 없습니다.'); return; }
  const rg = { r1: minOf(hits.map((h) => h[0])), c1: minOf(hits.map((h) => h[1])), r2: maxOf(hits.map((h) => h[0])), c2: maxOf(hits.map((h) => h[1])) };
  gv.ensureVisible(hits[0][0], hits[0][1]);
  selectRange(rg, 'cells', { r: hits[0][0], c: hits[0][1] });
  if (hits.length > 1) toast(`참조하는 셀 ${hits.length}개`);
}

/** Ctrl+. 선택 영역의 다음 모서리로 */
function nextCorner() {
  const corners = [[sel.r1, sel.c1], [sel.r1, sel.c2], [sel.r2, sel.c2], [sel.r2, sel.c1]];
  const i = corners.findIndex(([r, c]) => r === active.r && c === active.c);
  const [r, c] = corners[(i + 1) % 4];
  const keep = { ...sel };
  selectRange(keep, selKind, { r, c });
  gv.ensureVisible(r, c);
}

/** Ctrl+Shift+O 메모가 있는 셀 */
function selectComments() {
  const hits = [...sheet().cells].filter(([, cell]) => cell.comment).map(([k]) => k.split(',').map(Number));
  if (!hits.length) { toast('메모가 있는 셀이 없습니다.'); return; }
  const rg = { r1: minOf(hits.map((h) => h[0])), c1: minOf(hits.map((h) => h[1])), r2: maxOf(hits.map((h) => h[0])), c2: maxOf(hits.map((h) => h[1])) };
  selectRange(rg, 'cells', { r: hits[0][0], c: hits[0][1] });
  toast(`메모가 있는 셀 ${hits.length}개`);
}

/** Ctrl+K 하이퍼링크 */
function hyperlinkDialog() {
  if (editing && !commitEdit()) return;
  const cell = wb.getCell(si, active.r, active.c);
  formDialog('하이퍼링크 삽입', [
    { name: 'text', label: '표시할 텍스트', value: displayText(active.r, active.c) || '' },
    { name: 'url', label: '주소 (웹 주소 또는 #시트!A1)', value: cell?.link ?? '' },
  ], (v) => {
    let url = v.url.trim();
    if (!url) { toast('주소를 입력하세요.'); return false; }
    if (!url.startsWith('#') && !/^[a-z][\w+.-]*:/i.test(url)) url = /^[^\s/]+![$A-Za-z]+\$?\d+/.test(url) ? `#${url}` : `https://${url}`;
    const text = v.text.trim() || url.replace(/^#/, '');
    const cur = wb.getCell(si, active.r, active.c);
    wb.transact(() => {
      const raw = cur?.formula ? cur.raw : text;
      wb.setCellData(si, active.r, active.c, { raw, style: { ...cur?.style, color: cur?.style?.color ?? '#0563c1', underline: true }, comment: cur?.comment, link: url });
    }, meta());
    return true;
  }, { okLabel: '확인' });
}

function removeHyperlink() {
  const cells = cellsIn(sel).filter(([r, c]) => wb.getCell(si, r, c)?.link);
  if (!cells.length) return;
  wb.transact(() => cells.forEach(([r, c]) => {
    const cur = wb.getCell(si, r, c);
    const st = { ...cur.style };
    if (st.color === '#0563c1') delete st.color;
    delete st.underline;
    wb.setCellData(si, r, c, { raw: cur.raw, style: st, comment: cur.comment });
  }), meta());
}

function openLink(url) {
  if (url.startsWith('#')) { gotoRef(url.slice(1)); return; }
  try { window.open(url, '_blank', 'noopener'); } catch { toast(url); }
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
  if (t.classList.contains('olb') || t.classList.contains('olv')) {
    // 개요: 그룹 +/− 단추, 수준 단추
    e.preventDefault();
    e.stopPropagation();
    if (editing && !commitEdit()) return;
    const d = t.dataset;
    if (t.classList.contains('olv')) outlineShowLevel(d.ax, Number(d.l));
    else outlineToggle(d.ax, Number(d.a), Number(d.b), Number(d.l));
    return;
  }
  if (t.classList.contains('pxbtn')) {
    // 피벗 항목 펼치기 · 축소 (+/− 단추)
    e.preventDefault();
    if (editing && !commitEdit()) return;
    const entry = pivotDefs()[Number(t.dataset.p)];
    if (entry) togglePivotItem(entry, t.dataset.f, t.dataset.i);
    return;
  }
  if (t.classList.contains('pc-field') && t.dataset.f) {
    // 피벗 차트 필드 단추 → 연결된 피벗 테이블의 필드 필터
    e.preventDefault();
    e.stopPropagation();
    const ch = sheet().charts.find((c) => c.id === t.closest('.obj')?.dataset.id);
    const entry = ch?.pivot ? findPivotEntry(ch.pivot.sheet ?? null, ch.pivot.name ?? null) : null;
    if (entry) openPivotFilterMenu(entry, t.dataset.k, t.dataset.f, t);
    return;
  }
  if (t.classList.contains('pbtn')) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    const entry = pivotDefs()[Number(t.dataset.p)];
    if (entry) openPivotFilterMenu(entry, t.dataset.k, t.dataset.f || null, t);
    return;
  }
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
  // 시간 표시 막대: 칸을 눌러 끌면 기간 선택, 수준(연도 · 분기 · 월 · 일) 바꾸기
  const tlEl = t.closest('.tl-cell, .tl-level');
  if (tlEl && e.button === 0) {
    e.preventDefault();
    if (editing && !commitEdit()) return;
    focusGrid();
    const objEl = t.closest('.obj');
    const id = objEl.dataset.id;
    chartSel = id;
    const sl = sheet().slicers.find((x) => x.id === id);
    if (tlEl.classList.contains('tl-level')) {
      openMenu(tlEl, [['Y', '연도'], ['Q', '분기'], ['M', '월'], ['D', '일']].map(([k, l]) => ({ label: l, checked: (sl.level ?? 'M') === k, action: () => { updateObject(id, { level: k }); gv.renderObjectsAll(); } })));
      return;
    }
    const a = Number(tlEl.dataset.i);
    tlDrag = { id, a, b: a, root: objEl, shift: e.shiftKey };
    if (e.shiftKey && sl._anchor !== undefined) tlDrag.a = sl._anchor;
    markTimelineDrag();
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
    // Ctrl · Shift + 클릭: 여러 개체 선택 (맞춤 · 배분 · 한꺼번에 서식)
    if ((e.ctrlKey || e.metaKey || e.shiftKey) && e.button === 0 && chartSel && chartSel !== id) {
      if (objMulti.has(id)) objMulti.delete(id); else objMulti.add(id);
      gv.renderObjectsAll(); updateSelectionUI(); selPaneDlg?.redraw?.();
      return;
    }
    if (chartSel !== id) { chartSel = id; objMulti.clear(); gv.renderObjectsAll(); updateSelectionUI(); selPaneDlg?.redraw?.(); }
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
    const gpd = !e.shiftKey && opts.getPivotData ? getPivotDataText(r, c) : null;
    if (gpd) { insertPointText(gpd); drag = null; return; }
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
  const link = !e.shiftKey && !e.ctrlKey ? wb.getCell(si, r, c)?.link : null;
  drag = { type: 'select', ...(link ? { link: { r, c, url: link } } : {}) };
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
        const [minW, minH] = drag.prop === 'charts' ? [120, 90] : drag.prop === 'slicers' ? [80, 56] : LINE_SHAPES.has(ch.kind) ? [0, 0] : [8, 8];
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
      if (lastShift && !LINE_SHAPES.has(sh.kind)) { const m = Math.max(Math.abs(dx), Math.abs(dy)); dx = Math.sign(dx || 1) * m; dy = Math.sign(dy || 1) * m; }
      sh.x = Math.max(0, Math.round(Math.min(drag.x0, drag.x0 + dx)));
      sh.y = Math.max(0, Math.round(Math.min(drag.y0, drag.y0 + dy)));
      sh.w = Math.round(Math.abs(dx));
      sh.h = Math.round(Math.abs(dy));
      if (LINE_SHAPES.has(sh.kind)) { sh.flip = dx < 0; sh.flipV = dy < 0; }
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
  if (d.type === 'select' && d.link && selIsActiveOnly() && active.r === d.link.r && active.c === d.link.c) openLink(d.link.url);
  // 테두리 그리기 모드: 끌어서 고른 범위에 펜으로 바깥쪽(그리기) · 모든(눈금) 테두리, 또는 지우기
  if (d.type === 'select' && borderDraw) applyBorder(borderDraw === 'grid' ? 'all' : borderDraw === 'erase' ? 'none' : 'outside');
  switch (d.type) {
    case 'fill':
      fillPreview = null;
      gv.renderOverlays();
      if (d.target && !protectBlocked('cells', d.target)) doFill(d.src, d.target);
      break;
    case 'colResize': {
      if (d.newW === undefined) break;
      if (d.orig === undefined) delete sheet().colWidths[d.c]; else sheet().colWidths[d.c] = d.orig;
      const cols = selKind === 'cols' && d.c >= sel.c1 && d.c <= sel.c2 ? range(sel.c1, Math.min(sel.c2, sel.c1 + 500)) : [d.c];
      wb.transact(() => anchorObjects(() => cols.forEach((c) => wb.setColWidth(si, c, d.newW))), meta());
      break;
    }
    case 'rowResize': {
      if (d.newH === undefined) break;
      if (d.orig === undefined) delete sheet().rowHeights[d.r]; else sheet().rowHeights[d.r] = d.orig;
      const rows = selKind === 'rows' && d.r >= sel.r1 && d.r <= sel.r2 ? range(sel.r1, Math.min(sel.r2, sel.r1 + 2000)) : [d.r];
      wb.transact(() => anchorObjects(() => rows.forEach((r) => wb.setRowHeight(si, r, d.newH))), meta());
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
      if (LINE_SHAPES.has(temp.kind) ? temp.w + temp.h < 4 : temp.w < 4 || temp.h < 4) {
        Object.assign(temp, LINE_SHAPES.has(temp.kind) ? { w: 150, h: 0, flip: false } : temp.kind === 'textbox' ? { w: 160, h: 48 } : { w: 150, h: 90 });
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
    wb.transact(() => anchorObjects(() => { wb.setRowHeight(si, hit.edgeRow, defRowH(), false); autoFitRows(hit.edgeRow, hit.edgeRow); }), meta());
    return;
  }
  if (hit.zone !== 'cell' || editing) return;
  // 피벗 값 셀: 세부 정보 표시 (원본 행을 새 시트에)
  const pv = pivotHere();
  if (pv?.def.area && showPivotDetail(pv, active.r, active.c)) return;
  if (pv?.def.area) { const b = (pv.def.buttons ?? []).find((x) => x.kind === 'toggle' && x.r === active.r && x.c === active.c); if (b) { togglePivotItem(pv, b.field, b.item); return; } }
  startEdit('edit');
}

/** 세부 정보 표시: 피벗 값 칸을 이루는 원본 행으로 새 시트 + 표 (엑셀과 같음, 많으면 열 블록으로) */
function showPivotDetail(entry, r, c) {
  const def = pivotDefV2(entry.def);
  const src = pivotSource(def);
  if (!src) return false;
  const det = pivotDetail(src, def, r - (def.top ?? 0), c - (def.left ?? 0));
  if (!det) return false;
  const { header, cube, idx } = det;
  const n = idx.length;
  const W = header.length;
  const fmtOf = (j) => {
    if (!src.ref || src.si === undefined) return null;
    const st = wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + j);
    if (!st?.numFmt || st.numFmt === 'general') return null;
    const f = { numFmt: st.numFmt };
    if (st.code !== undefined) f.code = st.code;
    if (st.decimals !== undefined) f.decimals = st.decimals;
    return f;
  };
  const fmts = header.map((_, j) => fmtOf(j));
  const cols = header.map((_, j) => cube.col(j));
  const rawOfV = (v) => (v === null || v === undefined || v === '' ? null : typeof v === 'number' ? String(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'object' ? v.code ?? v.error ?? null : v.startsWith('=') || parseInput(v).value !== v ? `'${v}` : v);
  const t0 = performance.now();
  let at = -1;
  wb.transact(() => {
    at = wb.addSheet(null, entry.si);
    const sh = wb.sheets[at];
    header.forEach((h, j) => wb.setCellData(at, 0, j, { raw: rawOfV(String(h)), style: {} }));
    if (n > 5000) {
      // 많은 행: 열 블록 (값을 그대로 복사, 셀 객체를 만들지 않음)
      const bs = header.map(() => new ColBuilder(n));
      for (let k = 0; k < n; k++) { const i = idx[k]; for (let j = 0; j < W; j++) bs[j].set(k, cols[j].get(i)); }
      sh.blocks = [{ r0: 1, c0: 0, n, ver: 0, dver: 0, cols: bs.map((b, j) => b.finish(n, fmts[j])) }];
    } else {
      for (let k = 0; k < n; k++) {
        const i = idx[k];
        for (let j = 0; j < W; j++) {
          const raw = rawOfV(cols[j].get(i));
          if (raw !== null) wb.setCellData(at, k + 1, j, { raw, style: fmts[j] ?? {} });
        }
      }
    }
    const tname = nextTableName(wb);
    wb.setSheetProp(at, 'tables', [{
      id: newTableId(), name: tname, r1: 0, c1: 0, r2: Math.max(1, n), c2: Math.max(0, W - 1), header: true, totals: false, style: 'TableStyleMedium2',
      banded: true, bandedCols: false, firstCol: false, lastCol: false, filter: { criteria: {}, hidden: {} }, totalsFns: {},
    }]);
    wb.setSheetProp(at, 'freeze', { rows: 1, cols: 0 });
    // 열 너비: 머리글 + 앞쪽 200행으로 맞춤 (같은 실행 취소 단위)
    for (let j = 0; j < W; j++) {
      let w = measureText(String(header[j]), { bold: true }) + 28;
      for (let r = 1; r <= Math.min(n, 200); r++) {
        const v = wb.getValue(at, r, j);
        if (v !== null) w = Math.max(w, measureText(formatValue(v, wb.styleAt(at, r, j)).text, {}) + 12);
      }
      if (w > DEFAULT_COL_WIDTH) wb.setColWidth(at, j, Math.min(300, Math.ceil(w)));
    }
    wb.touch(at);
  }, meta());
  switchSheet(at);
  selectRange({ r1: 1, c1: 0, r2: 1, c2: 0 }, 'cells', { r: 1, c: 0 });
  const cond = det.conds.map(([f, v]) => `${f}=${v}`).join(', ');
  toast(`세부 정보 ${n.toLocaleString()}행을 새 시트 '${wb.sheets[at].name}'에 표시했습니다${cond ? ` (${cond})` : ''}. ${Math.round(performance.now() - t0)}ms`);
  return true;
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
      wb.setColWidth(si, c, w ? Math.min(600, Math.ceil(w)) : defColW());
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
  // 수식은 한 번만 나눠 두고 옮긴 글자만 만듦 (백만 행 채우기도 빠르게)
  const shifters = new Map();
  const shift = (d, dr, dc) => {
    if (!d || !d.raw.startsWith('=')) return d;
    let f = shifters.get(d.raw);
    if (!f) { f = formulaShifter(d.raw); shifters.set(d.raw, f); }
    return { ...d, raw: f(dr, dc) };
  };
  const anyHidden = !!(sheet().filter?.hidden || (sheet().tables ?? []).some((t) => t.filter?.hidden));
  wb.transact(() => {
    if (dir === 'down') {
      const srcR = rg.r1 === rg.r2 ? rg.r1 - 1 : rg.r1;
      if (srcR < 0) return;
      for (let c = rg.c1; c <= rg.c2; c++) {
        const d = cellData(wb.getCell(si, srcR, c));
        for (let r = srcR + 1; r <= rg.r2; r++) if (!anyHidden || !filterHidden(r)) wb.setCellData(si, r, c, shift(d, r - srcR, 0));
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

/**
 * 복사한 셀 붙여넣기. mode: 'all' | 'values' | 'formulas' | 'formats' | 'transpose'
 * opts (선택하여 붙여넣기): { what, op: 'add'|'sub'|'mul'|'div', skipBlanks, transpose }
 */
function pasteInternal(mode = 'all', opts = {}) {
  if (!clip) return;
  const transpose = mode === 'transpose' || !!opts.transpose;
  const what = opts.what ?? (mode === 'transpose' ? 'all' : mode);
  const op = opts.op ?? null;
  if (what === 'colWidths') {
    // 열 너비만
    const w = clip.c2 - clip.c1 + 1;
    wb.transact(() => { for (let j = 0; j < w; j++) wb.setColWidth(si, active.c + j, wb.colWidth(clip.si, clip.c1 + j)); }, meta());
    gv.layout();
    return;
  }
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
        if (opts.skipBlanks && !d?.raw) continue;
        const numFmtOf = (st) => (st ? Object.fromEntries(Object.entries(st).filter(([k2]) => ['numFmt', 'decimals', 'code'].includes(k2))) : {});
        let data;
        if (what === 'values') data = { raw: valueToRaw(src.values[ci][cj]), style: cur?.style, comment: cur?.comment };
        else if (what === 'valuesNum') data = { raw: valueToRaw(src.values[ci][cj]), style: { ...cur?.style, ...numFmtOf(d?.style) }, comment: cur?.comment };
        else if (what === 'formats') data = { raw: cur?.raw ?? '', style: d?.style, comment: cur?.comment };
        else if (what === 'formulas') data = { raw: shifted, style: cur?.style, comment: cur?.comment };
        else if (what === 'formulasNum') data = { raw: shifted, style: { ...cur?.style, ...numFmtOf(d?.style) }, comment: cur?.comment };
        else if (what === 'comments') data = { raw: cur?.raw ?? '', style: cur?.style, comment: d?.comment };
        else if (what === 'noBorders') {
          const st = { ...d?.style };
          for (const k2 of ['bt', 'bb', 'bl', 'br']) delete st[k2];
          data = d ? { ...d, raw: shifted, style: st } : null;
        } else if (what === 'validation') {
          data = cur;
        } else data = d ? { ...d, raw: shifted } : null;
        if (op && data && ['all', 'values', 'valuesNum', 'formulas', 'formulasNum', 'noBorders'].includes(what)) {
          // 연산: 대상 값 (연산) 복사한 값
          const sv = src.values[ci][cj];
          const sym = { add: '+', sub: '-', mul: '*', div: '/' }[op];
          const tv = wb.getValue(si, tr, tc);
          if (typeof sv === 'number' || sv === null) {
            const x = typeof sv === 'number' ? sv : 0;
            if (cur?.raw?.startsWith('=')) data = { ...cur, raw: `=(${cur.raw.slice(1)})${sym}${x}` };
            else if (typeof tv === 'number' || tv === null) {
              const a = typeof tv === 'number' ? tv : 0;
              const res = op === 'add' ? a + x : op === 'sub' ? a - x : op === 'mul' ? a * x : x === 0 ? null : a / x;
              data = { raw: res === null ? '#DIV/0!' : String(Number(res.toPrecision(15))), style: cur?.style ?? data.style, comment: cur?.comment };
            }
          }
        }
        wb.setCellData(si, tr, tc, data);
      }
    }
    if (what === 'validation') {
      // 유효성 검사 규칙을 대상 위치로 옮겨 복사
      const srcRules = (wb.sheets[src.si].validations ?? []).filter((v) => v.r1 <= src.r2 && v.r2 >= src.r1 && v.c1 <= src.c2 && v.c2 >= src.c1);
      const add = srcRules.map((v) => {
        const r1 = Math.max(v.r1, src.r1) - src.r1 + tgt.r1;
        const c1 = Math.max(v.c1, src.c1) - src.c1 + tgt.c1;
        return { ...structuredClone(v), r1, c1, r2: r1 + (Math.min(v.r2, src.r2) - Math.max(v.r1, src.r1)), c2: c1 + (Math.min(v.c2, src.c2) - Math.max(v.c1, src.c1)) };
      });
      if (add.length) wb.setSheetProp(si, 'validations', [...(sheet().validations ?? []), ...add]);
    }
    if (what !== 'formats' && what !== 'comments' && what !== 'validation') afterDataEntry({ r1: tgt.r1, c1: tgt.c1, r2: tgt.r1 + th - 1, c2: tgt.c1 + tw - 1 });
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
      afterDataEntry({ r1: active.r, c1: active.c, r2: active.r + rows.length - 1, c2: active.c + maxOf(rows.map((x) => x.length)) - 1 });
    }
  }, meta());
  if (rows.length > 1 || rows[0].length > 1) {
    selectRange({ r1: active.r, c1: active.c, r2: active.r + rows.length - 1, c2: active.c + maxOf(rows.map((r) => r.length)) - 1 });
  }
}

function handlePaste(text) {
  if (protectBlocked('cells')) return;
  const n = (s) => (s ?? '').replace(/\r\n?/g, '\n').replace(/\n$/, '');
  if (clip && (!text || n(text) === n(clip.text))) { pasteInternal('all'); return; }
  if (text) pasteText(text);
}

/** 선택하여 붙여넣기 (Ctrl+Alt+V) */
function pasteSpecialDialog() {
  if (editing && !commitEdit()) return;
  if (!clip) { toast('먼저 셀을 복사하세요. (선택하여 붙여넣기는 이 문서 안에서 복사한 셀에만 쓸 수 있습니다)'); return; }
  const WHAT = [
    ['all', '모두'], ['formulas', '수식'], ['values', '값'], ['formats', '서식'], ['comments', '메모'], ['validation', '유효성 검사'],
    ['noBorders', '테두리만 제외'], ['colWidths', '열 너비'], ['formulasNum', '수식 및 숫자 서식'], ['valuesNum', '값 및 숫자 서식'],
  ];
  const OPS = [['', '없음'], ['add', '더하기'], ['sub', '빼기'], ['mul', '곱하기'], ['div', '나누기']];
  const radio = (name, list, def) => el('div', { class: 'ps-grid' }, list.map(([v, label], i) => el('label', { class: 'fc-check' },
    el('input', { type: 'radio', name, value: v, checked: v === def }), `${label}(${'ASFTCNHWRU'[i] ?? ''})`)));
  const whatBox = radio('psWhat', WHAT, 'all');
  const opBox = radio('psOp', OPS, '');
  const skip = el('input', { type: 'checkbox' });
  const trans = el('input', { type: 'checkbox' });
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    el('fieldset', {}, el('legend', {}, '붙여넣기'), whatBox),
    el('fieldset', {}, el('legend', {}, '연산'), opBox),
    el('div', { style: { display: 'flex', gap: '16px' } }, el('label', { class: 'fc-check' }, skip, '내용 있는 셀만 붙여넣기(B)'), el('label', { class: 'fc-check' }, trans, '행/열 바꿈(E)')));
  body.addEventListener('keydown', (e) => {
    // 엑셀과 같은 글자 키로 항목 고르기 (V = 값, T = 서식 …)
    const map = { a: 'all', f: 'formulas', v: 'values', t: 'formats', c: 'comments', n: 'validation', x: 'noBorders', w: 'colWidths', r: 'formulasNum', u: 'valuesNum' };
    const k = e.key.toLowerCase();
    if (e.altKey || e.ctrlKey) return;
    if (map[k]) { whatBox.querySelector(`input[value="${map[k]}"]`).checked = true; e.preventDefault(); }
    if (k === 'b') { skip.checked = !skip.checked; e.preventDefault(); }
    if (k === 'e') { trans.checked = !trans.checked; e.preventDefault(); }
  });
  openDialog({
    title: '선택하여 붙여넣기', width: 420, body,
    buttons: [{
      label: '확인', primary: true, action: () => {
        const what = whatBox.querySelector('input:checked').value;
        const op = opBox.querySelector('input:checked').value || null;
        pasteInternal('all', { what, op, skipBlanks: skip.checked, transpose: trans.checked });
      },
    }, { label: '취소' }],
  });
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
  lastRepeat = () => applyStyle(patchOrFn, { widen });
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
        const cur = wb.getCell(si, r, c)?.style ?? wb.baseStyle ?? {};
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

// 테두리 펜: 선 스타일 · 선 색 (엑셀의 [테두리] → [선 색] · [선 스타일])
const borderPen = { style: 'thin', color: null };
let borderDraw = null; // 'outline' | 'grid' | 'erase' — 끌어서 테두리 그리기
const BORDER_STYLES = [['thin', '가는 실선'], ['hair', '아주 가는 선'], ['dotted', '점선'], ['dashDotDot', '이점쇄선'], ['dashDot', '일점쇄선'], ['dashed', '파선'],
  ['medium', '보통 실선'], ['mediumDashDotDot', '보통 이점쇄선'], ['slantDashDot', '기울어진 일점쇄선'], ['mediumDashDot', '보통 일점쇄선'], ['mediumDashed', '보통 파선'], ['thick', '굵은 실선'], ['double', '이중 실선']];

function lineStyleMenu(anchor) {
  openMenu(anchor, BORDER_STYLES.map(([v, l]) => {
    const [w, css] = { thin: [1, 'solid'], hair: [1, 'dotted'], dotted: [1, 'dotted'], dashed: [1, 'dashed'], dashDot: [1, 'dashed'], dashDotDot: [1, 'dashed'], medium: [2, 'solid'], mediumDashed: [2, 'dashed'], mediumDashDot: [2, 'dashed'], mediumDashDotDot: [2, 'dashed'], slantDashDot: [2, 'dashed'], thick: [3, 'solid'], double: [3, 'double'] }[v];
    return { label: l, checked: borderPen.style === v, icon: `<span style="display:block;width:22px;border-top:${w}px ${css} #333;margin-top:6px"></span>`, action: () => { borderPen.style = v; if (!borderDraw) setBorderDraw('outline'); } };
  }));
}

function setBorderDraw(mode) {
  borderDraw = borderDraw === mode ? null : mode;
  dom.view.classList.toggle('border-draw', !!borderDraw);
  if (borderDraw) toast(`${{ outline: '테두리 그리기', grid: '테두리 눈금 그리기', erase: '테두리 지우기' }[borderDraw]}: 셀을 끌어서 적용하세요. (Esc: 끝내기)`);
}

/** [셀 서식] → [테두리]: 가장자리마다 켜기/끄기 (바뀐 것만), pen = { style('none' = 지우기), color } */
function applyEdges(edges, edges0, pen) {
  const rg = selKind === 'cells' ? sel : usedClip(sel);
  const erase = pen.style === 'none';
  const put = (on) => (on && !erase ? { v: true, s: pen.style === 'thin' ? undefined : pen.style, c: pen.color === '#000000' ? undefined : pen.color } : { v: undefined, s: undefined, c: undefined });
  const side = (p, k, spec) => { p[k] = spec.v; p[`${k}s`] = spec.s; p[`${k}c`] = spec.c; };
  const changed = (k) => edges[k] !== edges0[k] || (edges[k] && !erase);
  for (const [r, c] of cellsIn(rg)) {
    const p = {};
    if (r === rg.r1 && changed('top')) side(p, 'bt', put(edges.top));
    if (r === rg.r2 && changed('bottom')) side(p, 'bb', put(edges.bottom));
    if (c === rg.c1 && changed('left')) side(p, 'bl', put(edges.left));
    if (c === rg.c2 && changed('right')) side(p, 'br', put(edges.right));
    if (changed('insideH')) { if (r > rg.r1) side(p, 'bt', put(edges.insideH)); if (r < rg.r2) side(p, 'bb', put(edges.insideH)); }
    if (changed('insideV')) { if (c > rg.c1) side(p, 'bl', put(edges.insideV)); if (c < rg.c2) side(p, 'br', put(edges.insideV)); }
    if (changed('diagDown')) side(p, 'dd', put(edges.diagDown));
    if (changed('diagUp')) side(p, 'du', put(edges.diagUp));
    if (Object.keys(p).length) wb.setStyle(si, r, c, p);
  }
}

function applyBorder(kind, penOverride = null) {
  lastBorder = kind;
  const pen = penOverride ?? borderPen;
  const rg = selKind === 'cells' ? sel : usedClip(sel);
  const thick = { thickOutside: 'medium', thickBottom: 'medium', topThickBottom: 'medium', doubleBottom: 'double', topDoubleBottom: 'double' }[kind];
  wb.transact(() => {
    for (const [r, c] of cellsIn(rg)) {
      let p = null;
      switch (kind) {
        case 'thickOutside': p = { ...(r === rg.r1 && { bt: true }), ...(r === rg.r2 && { bb: true }), ...(c === rg.c1 && { bl: true }), ...(c === rg.c2 && { br: true }) }; break;
        case 'inside': p = { ...(r > rg.r1 && { bt: true }), ...(r < rg.r2 && { bb: true }), ...(c > rg.c1 && { bl: true }), ...(c < rg.c2 && { br: true }) }; break;
        case 'insideH': p = { ...(r > rg.r1 && { bt: true }), ...(r < rg.r2 && { bb: true }) }; break;
        case 'insideV': p = { ...(c > rg.c1 && { bl: true }), ...(c < rg.c2 && { br: true }) }; break;
        case 'thickBottom': case 'doubleBottom': if (r === rg.r2) p = { bb: true }; break;
        case 'topThickBottom': case 'topDoubleBottom': p = { ...(r === rg.r1 && { bt: true }), ...(r === rg.r2 && { bb: true }) }; break;
        case 'diagDown': p = { dd: true, ddc: pen.color ?? undefined, dds: pen.style === 'thin' ? undefined : pen.style }; break;
        case 'diagUp': p = { du: true, duc: pen.color ?? undefined, dus: pen.style === 'thin' ? undefined : pen.style }; break;
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
      // 펜의 선 색 · 선 스타일 (굵은/이중 테두리 메뉴는 그 선 종류; 위쪽 선은 펜 그대로)
      if (p) {
        for (const k of ['bt', 'bb', 'bl', 'br']) {
          if (!(k in p) || !p[k]) continue;
          const st = thick && !(k === 'bt' && kind.startsWith('top')) ? thick : pen.style;
          p[`${k}c`] = pen.color ?? undefined;
          p[`${k}s`] = st === 'thin' ? undefined : st;
        }
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
  const cur = styleAt(active.r, active.c).size || BASE_FONT.size;
  const next = dir > 0 ? FONT_SIZES.find((s) => s > cur) ?? cur + 4 : [...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  applyStyle({ size: next === BASE_FONT.size ? undefined : next });
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
      title: 'WIXEL',
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
  // 엑셀처럼 개체도 셀을 따라 이동 · 크기 변경 (숨긴 행 안의 개체는 높이 0)
  wb.transact(() => anchorObjects(() => wb.setHidden(si, axis, idx, hide), null, true), meta());
  gv.layout();
  if (hide) {
    const n = axis === 'row' ? stepFrom({ r: b, c: active.c }, 1, 0) : stepFrom({ r: active.r, c: b }, 0, 1);
    selectCell(n.r, n.c);
  }
}

// ───────────────────────── 스파크라인 ─────────────────────────
const sparkGroupAt = (r, c) => (sheet().sparklines ?? []).find((g) => g.items.some((it) => it.r === r && it.c === c)) ?? null;

/** 삽입 → 스파크라인: 데이터 범위 + 위치 범위 (위치가 한 열이면 행마다) */
function insertSparkline(type) {
  const data = usedClip(sel);
  const guessLoc = data.c2 > data.c1 ? { r1: data.r1, c1: data.c2 + 1, r2: data.r2, c2: data.c2 + 1 } : { r1: data.r2 + 1, c1: data.c1, r2: data.r2 + 1, c2: data.c2 };
  const a1 = (rg) => `${cellName(rg.r1, rg.c1)}${rg.r1 === rg.r2 && rg.c1 === rg.c2 ? '' : `:${cellName(rg.r2, rg.c2)}`}`;
  formDialog(`스파크라인 만들기 (${SPARK_TYPES.find((t) => t.id === type).label})`, [
    { name: 'data', label: '데이터 범위', value: a1(data) },
    { name: 'loc', label: '위치 범위', value: a1(guessLoc) },
  ], (v) => {
    const d = parseRangeName(v.data.replace(/^.*!/, ''));
    const l = parseRangeName(v.loc.replace(/^.*!/, ''));
    const prefix = v.data.includes('!') ? `${v.data.slice(0, v.data.lastIndexOf('!') + 1)}` : '';
    const items = d && l ? sparkItems(d, l, prefix) : null;
    if (!items) { alertDialog('스파크라인', '위치 범위는 데이터 범위의 행 수(한 열)나 열 수(한 행)와 같아야 합니다.'); return false; }
    // 같은 칸의 기존 스파크라인은 바꿈
    const taken = new Set(items.map((it) => `${it.r},${it.c}`));
    const rest = (sheet().sparklines ?? []).map((g) => ({ ...g, items: g.items.filter((it) => !taken.has(`${it.r},${it.c}`)) })).filter((g) => g.items.length);
    const g = { id: `sp${Date.now().toString(36)}`, ...sparkDefaults(type), items };
    wb.transact(() => wb.setSheetProp(si, 'sparklines', [...rest, g]), meta());
    selectCell(items[0].r, items[0].c);
    gv.renderAll();
    return undefined;
  });
}

/** 활성 셀의 스파크라인 그룹 고치기 */
function updateSparkGroup(patch) {
  const g = sparkGroupAt(active.r, active.c);
  if (!g) { toast('스파크라인이 있는 셀을 선택하세요.'); return; }
  const list = (sheet().sparklines ?? []).map((x) => (x.id === g.id ? { ...x, ...(typeof patch === 'function' ? patch(x) : patch) } : x)).filter((x) => x.items.length);
  wb.transact(() => wb.setSheetProp(si, 'sparklines', list), meta());
  gv.renderAll();
  updateSelectionUI();
}

function sparkEditDialog() {
  const g = sparkGroupAt(active.r, active.c);
  if (!g) { toast('스파크라인이 있는 셀을 선택하세요.'); return; }
  formDialog('스파크라인 편집', [
    { name: 'color', label: '스파크라인 색', type: 'color', value: g.color },
    { name: 'negColor', label: '음수 점 색', type: 'color', value: g.negColor },
    { name: 'markerColor', label: '표식 · 높은/낮은/첫/마지막 점 색', type: 'color', value: g.markerColor },
    { name: 'weight', label: '두께 (꺾은선형, pt)', type: 'number', value: g.weight ?? 1.25 },
    { name: 'refs', label: '데이터 범위 (셀마다, 줄 바꿈으로 구분)', type: 'textarea', value: g.items.map((it) => `${cellName(it.r, it.c)} ← ${it.ref}`).join('\n') },
  ], (v) => {
    const refs = v.refs.split('\n').map((line) => line.split('←').map((x) => x.trim())).filter((x) => x.length === 2);
    const items = g.items.map((it) => { const hit = refs.find(([cell]) => cell.toUpperCase() === cellName(it.r, it.c)); return hit && sparkRef(hit[1]) ? { ...it, ref: hit[1] } : it; });
    const mc = v.markerColor;
    updateSparkGroup({ color: v.color, negColor: v.negColor, markerColor: mc, highColor: mc, lowColor: mc, firstColor: mc, lastColor: mc, weight: Number(v.weight) || 1.25, items });
  });
}

// ───────────────────────── 시트 보호 ─────────────────────────
// 보호된 시트에서 명령마다 필요한 권한 (없는 명령은 선택한 셀이 모두 잠기지 않았을 때만)
const PROTECT_FREE = new Set(['publish', 'versionHistory', 'recentFiles', 'dataAnalysis', 'forecastSheet', 'scenarioManager', 'solver', 'undo', 'redo', 'save', 'open', 'backstage', 'print', 'copy', 'find', 'goto', 'prevSheet', 'nextSheet', 'selectRegion',
  'newWorkbook', 'pivotFieldList', 'tracePrecedents', 'traceDependents', 'removeArrows', 'evaluateFormula', 'errorCheck', 'watchWindow', 'gotoSpecial',
  'outlineShow', 'outlineHide', 'freezePanes', 'freezeTop', 'freezeFirstCol', 'circleInvalid', 'clearCircles', 'macros', 'prevComment', 'nextComment',
  'workbookStats', 'toggleGrid', 'togglePrintGrid', 'toggleFormulaBar', 'toggleHeaders', 'toggleFormulas', 'toggleRibbon', 'zoomIn', 'zoomOut', 'zoom100',
  'recalc', 'shortcuts', 'about', 'protectSheet', 'unprotectSheet', 'insertMenuKey', 'deleteMenuKey', 'addSheet', 'deleteSheet', 'duplicateSheet',
  'hideSheet', 'unhideSheet', 'importCsv', 'exportCsv', 'selectPrecedents', 'selectDependents', 'selectComments', 'pageSetup', 'printArea', 'clearPrintArea',
  'orientPortrait', 'orientLandscape', 'insertFunction']);
const PROTECT_BLOCK = new Set(['mergeCenter', 'createTable', 'condManager', 'condNewRule', 'condMenuKey', 'tableStyleKey', 'dataValidation', 'insertPivot', 'outlineGroup',
  'outlineUngroup', 'outlineClear', 'subtotal', 'resizeTable', 'convertToRange', 'tblName', 'tblHeader', 'tblTotals', 'tblBanded', 'tblBandedCols', 'tblFirstCol',
  'tblLastCol', 'tblFilter', 'textToColumns', 'dedupe', 'sparkLine', 'sparkColumn', 'sparkWinLoss', 'sparkClear', 'sparkEdit']);
const PROTECT_MAP = {
  insertRows: 'insertRows', insertCols: 'insertColumns', deleteRows: 'deleteRows', deleteCols: 'deleteColumns', sortAsc: 'sort', sortDesc: 'sort', sortDialog: 'sort',
  clearFilter: 'autoFilter', reapplyFilter: 'autoFilter', toggleFilter: 'autoFilter', hideRows: 'formatRows', unhideRows: 'formatRows', autofitRowsSel: 'formatRows',
  hideCols: 'formatColumns', autofitSel: 'formatColumns', refreshAll: 'pivotTables', calcField: 'pivotTables', slicerConnections: 'pivotTables',
  chartColumn: 'objects', chartBar: 'objects', chartLine: 'objects', chartPie: 'objects', chartArea: 'objects', chartScatter: 'objects', shapesMenu: 'objects',
  insertTextbox: 'objects', insertPicture: 'objects', insertSlicer: 'objects', insertTimeline: 'objects',
};
const FORMAT_CMDS = /^(painter|painterSticky|bold|italic|underline|strike|fontFamily|fontSize|growFont|shrinkFont|border|fillColor|fontColor|fontDialog|formatCells|align|valign|wrap|indent|numFmt|fmt|incDecimal|decDecimal|clearFormats|cellStyle)/;
function protectAction(cmd) {
  if (PROTECT_FREE.has(cmd)) return 'free';
  if (PROTECT_BLOCK.has(cmd)) return 'block';
  if (PROTECT_MAP[cmd]) return PROTECT_MAP[cmd];
  if (/^(pivot|pv|slicer)/.test(cmd)) return 'pivotTables';
  if (FORMAT_CMDS.test(cmd)) return 'formatCells';
  return 'cells';
}
/** 범위 안에 잠긴 셀이 있는지 (너무 크면 앞쪽만 보고, 시트 기본 서식으로 판단) */
function anyLocked(rg) {
  const r = usedClip(rg);
  const n = (r.r2 - r.r1 + 1) * (r.c2 - r.c1 + 1);
  if (n > 200000) return isLockedStyle(sheet().allStyle ?? {});
  for (let rr = r.r1; rr <= r.r2; rr++) for (let cc = r.c1; cc <= r.c2; cc++) if (isLockedStyle(wb.styleAt(si, rr, cc))) return true;
  return false;
}
/** 보호 때문에 막히면 알리고 true */
/** 피벗 테이블 영역과 겹치면 그 피벗 (엑셀처럼 피벗 결과는 직접 고칠 수 없음 — 옵션으로 허용) */
function pivotAreaHit(rg) {
  for (const { def } of pivotDefs(si)) {
    const a = def.area;
    if (a && rg.r1 <= a.r2 && rg.r2 >= a.r1 && rg.c1 <= a.c2 && rg.c2 >= a.c1) return def;
  }
  return null;
}

// 피벗 테이블 안에서 막는 명령: 셀 값을 바꾸는 것만 (차트 · 슬라이서 · 서식 삽입 등은 엑셀처럼 허용)
const PIVOT_LOCKED = /^(clear(Contents|All)?|delete(Rows|Cols|Cells|MenuKey)?|paste(ValuesKey|Values|Special)?|cut|fill(Down|Right|Up|Left|Series)?|flashFill|dedupe|textToColumns|insert(Rows|Cols|Cells|MenuKey|Date|Time)|merge(Center|Across|Cells)?|unmerge|autosum|dataTable|goalSeek|subtotal)$/;
function protectBlocked(action = 'cells', rg = sel, cmd = null) {
  if (action === 'cells' && !opts.pivotEdit && (!cmd || PIVOT_LOCKED.test(cmd)) && pivotAreaHit(rg)) {
    alertDialog('WIXEL', '피벗 테이블의 일부는 변경할 수 없습니다. 피벗 테이블은 원본 데이터를 계산한 결과이므로 값을 바꾸려면 원본 데이터를 고친 뒤 새로 고치세요.\n(파일 → 옵션 → 피벗 테이블 값 영역의 셀 편집 허용을 켜면 가상 분석용으로 편집할 수 있습니다.)');
    return true;
  }
  const sh = sheet();
  if (!isProtected(sh) || action === 'free') return false;
  const blocked = action === 'cells' ? anyLocked(special?.si === si ? { r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 } : rg) : action === 'block' || !allowed(sh, action);
  if (blocked) alertDialog('WIXEL', '변경하려는 셀이나 차트가 보호된 시트에 있습니다. 변경하려면 [검토] 탭에서 [시트 보호 해제]를 누르세요. 암호를 입력해야 할 수도 있습니다.');
  return blocked;
}

/** 엑셀 최신 방식(SHA-512 + salt + 반복) 암호 확인 */
async function verifyModernHash(m, pw) {
  const enc = new Uint8Array(pw.length * 2);
  for (let i = 0; i < pw.length; i++) { enc[i * 2] = pw.charCodeAt(i) & 255; enc[i * 2 + 1] = pw.charCodeAt(i) >> 8; }
  const salt = Uint8Array.from(atob(m.saltValue), (ch) => ch.charCodeAt(0));
  let h = new Uint8Array(await crypto.subtle.digest(m.algorithmName === 'SHA-256' ? 'SHA-256' : 'SHA-512', new Uint8Array([...salt, ...enc])));
  const n = Number(m.spinCount) || 0;
  const buf = new Uint8Array(h.length + 4);
  for (let i = 0; i < n; i++) {
    buf.set(h);
    buf[h.length] = i & 255; buf[h.length + 1] = (i >> 8) & 255; buf[h.length + 2] = (i >> 16) & 255; buf[h.length + 3] = (i >>> 24) & 255;
    h = new Uint8Array(await crypto.subtle.digest(m.algorithmName === 'SHA-256' ? 'SHA-256' : 'SHA-512', buf.slice(0, h.length + 4)));
  }
  return btoa(String.fromCharCode(...h)) === m.hashValue;
}

function protectSheetDialog() {
  const sh = sheet();
  if (isProtected(sh)) { unprotectSheet(); return; }
  const allow = { ...defaultAllow(), ...(sh.protect?.allow ?? {}) };
  formDialog('시트 보호', [
    { name: 'pw', label: '시트 보호 해제 암호 (선택)', type: 'password', value: '' },
    { name: 'pw2', label: '암호 확인', type: 'password', value: '' },
    ...PROTECT_OPTIONS.map((o) => ({ name: o.id, label: `허용: ${o.label}`, type: 'checkbox', value: allow[o.id] })),
  ], (v) => {
    if (v.pw !== v.pw2) { alertDialog('시트 보호', '확인 암호가 일치하지 않습니다.'); return false; }
    const next = { on: true, hash: excelHash(v.pw), allow: Object.fromEntries(PROTECT_OPTIONS.map((o) => [o.id, !!v[o.id]])) };
    wb.transact(() => wb.setSheetProp(si, 'protect', next), meta());
    ribbon.update?.(ribbonState());
    toast(`'${sh.name}' 시트를 보호했습니다. 잠기지 않은 셀만 편집할 수 있습니다 (셀 서식 → 보호에서 잠금 해제).`);
    return undefined;
  }, { note: '보호하면 잠긴 셀(기본)은 바꿀 수 없습니다. 입력을 허용할 셀은 먼저 [셀 서식 → 보호]에서 [잠금]을 해제하세요.' });
}

function unprotectSheet() {
  const sh = sheet();
  const p = sh.protect;
  const done = () => { wb.transact(() => wb.setSheetProp(si, 'protect', null), meta()); ribbon.update?.(ribbonState()); updateSelectionUI(); toast('시트 보호를 해제했습니다.'); };
  if (!p?.hash && !p?.modern?.hashValue) { done(); return; }
  formDialog('시트 보호 해제', [{ name: 'pw', label: '암호', type: 'password', value: '' }], (v) => {
    if (p.hash) {
      if (excelHash(v.pw) !== String(p.hash).toUpperCase().padStart(4, '0')) { alertDialog('시트 보호 해제', '암호가 잘못되었습니다. Caps Lock 키가 켜져 있는지 확인하고 대/소문자를 정확히 입력하세요.'); return false; }
      done();
      return undefined;
    }
    verifyModernHash(p.modern, v.pw).then((ok) => { if (ok) done(); else alertDialog('시트 보호 해제', '암호가 잘못되었습니다.'); }).catch(() => alertDialog('시트 보호 해제', '이 암호 방식은 확인할 수 없습니다.'));
    return undefined;
  });
}

/** 셀 잠금 · 수식 숨기기 (셀 서식 → 보호) */
function cellProtectionDialog() {
  const st = wb.styleAt(si, active.r, active.c);
  formDialog('셀 보호', [
    { name: 'locked', label: '잠금', type: 'checkbox', value: st.locked !== false },
    { name: 'hide', label: '숨김 (수식 숨기기)', type: 'checkbox', value: !!st.hideFormula },
  ], (v) => {
    applyStyle({ locked: v.locked ? undefined : false, hideFormula: v.hide || undefined });
  }, { note: '셀 잠금이나 수식 숨기기는 시트를 보호해야 적용됩니다 ([검토] → [시트 보호]).' });
}

// ───────────────────────── 수식 분석 · 가상 분석 ─────────────────────────
const cellKey = (x) => `${x.si}:${x.r},${x.c}`;
/** 참조되는 셀 추적: 누를 때마다 한 단계 더 (엑셀과 같음) */
function tracePrecedents() {
  if (!trace || trace.si !== si) trace = { si, arrows: [], frontP: null, frontD: null, seen: new Set() };
  const front = trace.frontP ?? [{ si, r: active.r, c: active.c }];
  const next = [];
  let added = 0;
  for (const cell of front) {
    const k = `p${cellKey(cell)}`;
    if (trace.seen.has(k)) continue;
    trace.seen.add(k);
    for (const b of wb.precedentsOf(cell.si, cell.r, cell.c)) {
      const err = [b].some(() => { for (let r = b.r1; r <= Math.min(b.r2, b.r1 + 50); r++) for (let c = b.c1; c <= Math.min(b.c2, b.c1 + 20); c++) if (isError(wb.getValue(b.si, r, c))) return true; return false; });
      trace.arrows.push({ from: b, to: cell, err });
      added++;
      // 다음 단계: 참조한 칸 중 수식 (범위는 앞쪽 일부만)
      for (let r = b.r1; r <= Math.min(b.r2, b.r1 + 200) && next.length < 300; r++) {
        for (let c = b.c1; c <= Math.min(b.c2, b.c1 + 20); c++) if (wb.getCell(b.si, r, c)?.formula) next.push({ si: b.si, r, c });
      }
    }
  }
  trace.frontP = next;
  if (!added) toast(front.length === 1 && front[0].r === active.r ? '이 셀은 다른 셀을 참조하지 않습니다.' : '더 추적할 참조가 없습니다.');
  gv.renderAll();
}
function traceDependents() {
  if (!trace || trace.si !== si) trace = { si, arrows: [], frontP: null, frontD: null, seen: new Set() };
  const front = trace.frontD ?? [{ si, r: active.r, c: active.c }];
  const next = [];
  let added = 0;
  for (const cell of front) {
    const k = `d${cellKey(cell)}`;
    if (trace.seen.has(k)) continue;
    trace.seen.add(k);
    for (const d of wb.dependentsOf(cell.si, cell.r, cell.c).slice(0, 500)) {
      trace.arrows.push({ from: { si: cell.si, r1: cell.r, c1: cell.c, r2: cell.r, c2: cell.c }, to: d, err: isError(wb.getValue(d.si, d.r, d.c)) });
      added++;
      next.push(d);
    }
  }
  trace.frontD = next;
  if (!added) toast('이 셀을 참조하는 수식이 없습니다.');
  gv.renderAll();
}
function removeArrows() { trace = null; gv.renderAll(); }

/** 수식 계산: 안쪽 부분식부터 한 단계씩 값으로 바꿔 보여 줌 */
function evaluateFormulaDialog() {
  const cell = wb.getCell(si, active.r, active.c);
  if (!cell?.formula || !cell.ast) { alertDialog('수식 계산', '선택한 셀에 수식이 없습니다.'); return; }
  const ctx = wb.ctxFor(si, active.r, active.c);
  ctx.dr = cell.dr ?? 0;
  ctx.dc = cell.dc ?? 0;
  const src = cell.raw.slice(1);
  const steps = evalSteps(src, cell.ast, ctx);
  let i = 0;
  const view = el('div', { class: 'eval-box' });
  const show = () => {
    const st = steps[i];
    view.replaceChildren();
    if (st.final) { view.append(el('div', {}, `${cellName(active.r, active.c)} = `), el('b', {}, st.text)); return; }
    const [a, b] = st.mark ?? [0, 0];
    view.append(el('span', {}, st.text.slice(0, a)), el('u', { class: 'eval-next' }, st.text.slice(a, b)), el('span', {}, st.text.slice(b)));
  };
  show();
  const body = el('div', {}, el('div', { class: 'muted' }, `참조: ${sheet().name}!${cellName(active.r, active.c)}  ·  밑줄 친 식을 계산하려면 [계산]을 누르세요.`), view);
  openDialog({
    title: '수식 계산', body, width: 560,
    buttons: [
      { label: '계산', primary: true, action: () => { if (i < steps.length - 1) i++; else i = 0; show(); return false; } },
      { label: '처음부터', action: () => { i = 0; show(); return false; } },
      { label: '닫기' },
    ],
  });
}

/** 오류 검사: 시트에서 오류 값이 있는 다음 칸으로 */
function errorCheck() {
  const list = [];
  const ext = wb.extent(si);
  sheet().cells.forEachRC((cell, r, c) => { if (cell.formula && isError(wb.getValue(si, r, c))) list.push([r, c]); });
  if (!list.length || ext.rows === 0) { alertDialog('오류 검사', '시트 전체에서 오류를 검사했습니다. 오류가 없습니다.'); return; }
  list.sort((a, b) => a[0] - b[0] || a[1] - b[1]);
  const at = list.find(([r, c]) => r > active.r || (r === active.r && c > active.c)) ?? list[0];
  selectCell(at[0], at[1]);
  const v = wb.getValue(si, at[0], at[1]);
  const why = { '#DIV/0!': '0으로 나누었습니다.', '#N/A': '찾는 값을 사용할 수 없습니다.', '#NAME?': '알 수 없는 이름이나 함수가 있습니다.', '#REF!': '잘못된 셀 참조입니다.', '#VALUE!': '값의 형식이 잘못되었습니다.', '#NUM!': '숫자가 잘못되었습니다.', '#SPILL!': '결과를 펼칠 칸이 비어 있지 않습니다.', '#CALC!': '계산할 수 없습니다.', '#NULL!': '교집합이 없습니다.' }[v.code] ?? '';
  toast(`${cellName(at[0], at[1])}: ${v.code} — ${why} (오류 ${list.length}개 중 ${list.indexOf(at) + 1}번째, 다시 누르면 다음 오류)`);
}

// 조사식 창: 지켜볼 셀의 값 · 수식을 계속 보여 줌
let watches = [];
let watchPane = null;
function watchWindow(show = true) {
  if (!show) { watchPane?.remove(); watchPane = null; return; }
  if (!watchPane) {
    watchPane = el('div', { class: 'watch-pane' });
    document.body.append(watchPane);
    let raf = 0;
    wb.onChange(() => { if (watchPane && !raf) raf = requestAnimationFrame(() => { raf = 0; renderWatches(); }); });
  }
  renderWatches();
}
function renderWatches() {
  if (!watchPane) return;
  const rows = watches.filter((w) => wb.sheets[w.si]).map((w, i) => {
    const cell = wb.getCell(w.si, w.r, w.c);
    const v = wb.getValue(w.si, w.r, w.c);
    return el('tr', {}, el('td', {}, wb.sheets[w.si].name), el('td', {}, cellName(w.r, w.c)),
      el('td', { class: 'num' }, formatValue(v, wb.styleAt(w.si, w.r, w.c)).text), el('td', { class: 'f' }, cell?.formula ? cell.raw : ''),
      el('td', {}, el('button', { class: 'lnk', title: '조사식 삭제', onclick: () => { watches.splice(i, 1); renderWatches(); } }, '✕')));
  });
  watchPane.replaceChildren(
    el('div', { class: 'watch-head' }, '조사식 창',
      el('button', { class: 'lnk', onclick: () => { for (const [r, c] of [...cellsIn(usedClip(sel))].slice(0, 50)) if (!watches.some((w) => w.si === si && w.r === r && w.c === c)) watches.push({ si, r, c }); renderWatches(); } }, '+ 조사식 추가'),
      el('button', { class: 'lnk', title: '닫기', onclick: () => watchWindow(false) }, '✕')),
    el('table', {}, el('tr', {}, ...['시트', '셀', '값', '수식', ''].map((h) => el('th', {}, h))), ...rows),
    rows.length ? null : el('div', { class: 'muted' }, '셀을 선택하고 [+ 조사식 추가]를 누르세요.'),
  );
}

/** 목표값 찾기 */
function goalSeekDialog() {
  formDialog('목표값 찾기', [
    { name: 'set', label: '수식 셀', value: cellName(active.r, active.c) },
    { name: 'to', label: '찾는 값', value: '' },
    { name: 'by', label: '값을 바꿀 셀', value: '' },
  ], (v) => {
    const a = parseRangeName(v.set);
    const b = parseRangeName(v.by);
    const target = Number(v.to);
    if (!a || !b || !Number.isFinite(target)) { alertDialog('목표값 찾기', '셀 주소와 숫자를 올바르게 입력하세요.'); return false; }
    if (!wb.getCell(si, a.r1, a.c1)?.formula) { alertDialog('목표값 찾기', '수식 셀에는 수식이 있어야 합니다.'); return false; }
    if (wb.getCell(si, b.r1, b.c1)?.formula) { alertDialog('목표값 찾기', '값을 바꿀 셀에는 수식이 아닌 값이 있어야 합니다.'); return false; }
    const orig = cellData(wb.getCell(si, b.r1, b.c1));
    const x0 = Number(valueAt(b.r1, b.c1)) || 0;
    // 계산 중에는 실행 취소 기록 없이 값만 바꿔 보고, 끝나면 원래대로 돌린 뒤 결과를 한 번에 기록
    const put = (x) => wb.setCellData(si, b.r1, b.c1, { ...(orig ?? {}), raw: String(x) });
    const res = goalSeek((x) => { put(x); return wb.getValue(si, a.r1, a.c1); }, x0, target);
    wb.setCellData(si, b.r1, b.c1, orig);
    const x = Number(res.x.toPrecision(15));
    const msg = `${cellName(a.r1, a.c1)} 셀로 목표값 찾기: ${res.ok ? '해를 찾았습니다.' : '해를 찾지 못했습니다 (가장 가까운 값).'}\n목표값: ${formatGeneral(target)}\n현재값: ${formatGeneral(Number(res.value.toPrecision(12)))}\n${cellName(b.r1, b.c1)} = ${formatGeneral(x)}`;
    openDialog({
      title: '목표값 찾기 상태', body: el('pre', { class: 'plain' }), width: 380,
      onOpen: (d) => { d.querySelector('pre').textContent = msg; },
      buttons: [
        { label: '확인', primary: true, action: () => { wb.transact(() => wb.setCellData(si, b.r1, b.c1, { ...(orig ?? {}), raw: String(x) }), meta()); } },
        { label: '취소' },
      ],
    });
    return undefined;
  });
}

/** 데이터 표 (가상 분석): 선택 범위의 첫 행 · 첫 열에 입력 값, 모서리 · 첫 행/열에 수식 */
function dataTableDialog() {
  const rg = usedClip(sel);
  if (rg.r2 <= rg.r1 && rg.c2 <= rg.c1) { alertDialog('데이터 표', '입력 값과 수식을 포함한 표 범위를 선택하세요.'); return; }
  formDialog('데이터 표', [
    { name: 'row', label: '행 입력 셀 (첫 행의 값을 넣을 셀)', value: '' },
    { name: 'col', label: '열 입력 셀 (첫 열의 값을 넣을 셀)', value: '' },
  ], (v) => {
    const ri = v.row.trim() ? parseRangeName(v.row.trim()) : null;
    const ci = v.col.trim() ? parseRangeName(v.col.trim()) : null;
    if (!ri && !ci) { alertDialog('데이터 표', '입력 셀을 하나 이상 지정하세요.'); return false; }
    const origs = [ri, ci].filter(Boolean).map((x) => [x, cellData(wb.getCell(si, x.r1, x.c1))]);
    const setIn = (x, val) => wb.setCellData(si, x.r1, x.c1, { ...(cellData(wb.getCell(si, x.r1, x.c1)) ?? {}), raw: val === null || val === undefined ? '' : isError(val) ? val.code : String(val) });
    const out = dataTable(rg, ri, ci, setIn, (r, c) => wb.getValue(si, r, c));
    for (const [x, d] of origs) wb.setCellData(si, x.r1, x.c1, d);
    wb.transact(() => { for (const [r, c, val] of out) wb.setCellData(si, r, c, { ...(cellData(wb.getCell(si, r, c)) ?? {}), raw: val === null ? '' : isError(val) ? val.code : typeof val === 'string' ? `'${val}` : typeof val === 'boolean' ? (val ? 'TRUE' : 'FALSE') : String(val) }); }, meta());
    toast(`데이터 표: ${out.length}개 칸을 계산했습니다 (값으로 입력됨, 입력 값을 바꾸면 다시 실행하세요).`);
    return undefined;
  }, { note: '행 입력 셀만 지정하면 첫 열에 수식, 열 입력 셀만 지정하면 첫 행에 수식, 둘 다 지정하면 왼쪽 위 모서리에 수식을 두세요.' });
}

// ───────────────────────── 데이터 분석 · 예측 시트 · 시나리오 관리자 ─────────────────────────
// 범위 입력 칸: 모덜리스 대화상자에서 칸을 누른 뒤 시트에서 끌어 선택하면 주소가 들어감 (엑셀 RefEdit)
let refPick = null;
const selKey = () => `${si}:${sel.r1},${sel.c1},${sel.r2},${sel.c2}`;
function refText(rg, s = si, withSheet = false) {
  const a = (r, c) => `$${colToName(c)}$${r + 1}`;
  const t = rg.r1 === rg.r2 && rg.c1 === rg.c2 ? a(rg.r1, rg.c1) : `${a(rg.r1, rg.c1)}:${a(rg.r2, rg.c2)}`;
  return withSheet ? `${quoteSheetName(wb.sheets[s].name)}!${t}` : t;
}
function pickRef() {
  if (!refPick?.input.isConnected) { refPick = null; return; }
  const k = selKey();
  if (k === refPick.key) return;
  refPick.key = k;
  refPick.input.value = refText(usedClip(sel), si, si !== refPick.home);
}
/** 'Sheet 2'!$A$1:$B$9 · A1:B9 → { si, rg } (열 전체 · 행 전체는 사용 범위로 자름) */
function parseRefInput(text) {
  let body = String(text ?? '').trim();
  if (!body) return null;
  let s = si;
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!(.+)$/.exec(body);
  if (m) {
    s = wb.sheetIndexByName((m[1] ?? m[2]).replace(/''/g, "'"));
    if (s < 0) return null;
    body = m[3];
  }
  const rg = parseRangeName(body.replace(/\$/g, ''));
  if (!rg) return null;
  const u = wb.usedRange(s);
  return { si: s, rg: { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, Math.max(rg.r1, u.rows - 1)), c2: Math.min(rg.c2, Math.max(rg.c1, u.cols - 1)) } };
}
function refInput(value = '', home = si) {
  const inp = el('input', { type: 'text', class: 'ref-input', value, spellcheck: 'false' });
  inp.addEventListener('focus', () => { refPick = { input: inp, key: selKey(), home }; });
  return inp;
}
const readBlock = (p) => {
  const out = [];
  for (let r = p.rg.r1; r <= p.rg.r2; r++) {
    const row = [];
    for (let c = p.rg.c1; c <= p.rg.c2; c++) { const v = wb.getValue(p.si, r, c); row.push(v === '' || v === undefined ? null : v); }
    out.push(row);
  }
  return out;
};
function confirmBox(title, msg) {
  return new Promise((resolve) => {
    openDialog({ title, body: el('div', { style: { whiteSpace: 'pre-wrap' } }, msg), width: 380, onClose: () => resolve(false), buttons: [{ label: '확인', primary: true, action: () => { resolve(true); } }, { label: '취소' }] });
  });
}
const nextSheetName = (base) => { let n = 1; let name = base; while (wb.sheetIndexByName(name) >= 0) name = `${base} (${++n})`; return name.slice(0, 31); };

/** 분석 결과 표를 시트에 쓰기 → { si, r, c, h, w } */
function writeAnalysis(res, dest) {
  const rows = res.rows;
  const h = rows.length;
  const w = Math.max(1, ...rows.map((r) => r.length));
  return wb.transact(() => {
    let at = dest.si;
    if (dest.newSheet) at = wb.addSheet(nextSheetName(dest.newSheet), si + 1);
    const r0 = dest.newSheet ? 0 : dest.r;
    const c0 = dest.newSheet ? 0 : dest.c;
    const heads = new Set(res.heads ?? []);
    rows.forEach((row, i) => {
      for (let j = 0; j < w; j++) {
        const v = row[j];
        const style = {};
        if (heads.has(i)) { style.italic = true; style.bb = true; if (i === 0 || rows[i - 1]?.length === 0 || rows[i - 1]?.every((x) => x === null || x === undefined)) style.bt = true; }
        if (res.firstColBold && j === 0 && i > 0) style.italic = true;
        if (res.pct?.has(`${i},${j}`)) { style.numFmt = 'percent'; style.decimals = 2; }
        if (v === null || v === undefined) {
          if (Object.keys(style).length && row.length) wb.setCellData(at, r0 + i, c0 + j, { raw: '', style });
          else if (!dest.newSheet && wb.getCell(at, r0 + i, c0 + j)) wb.setCellData(at, r0 + i, c0 + j, null);
          continue;
        }
        const raw = typeof v === 'number' ? String(v) : typeof v === 'string' ? `'${v}` : v.f ?? v.err ?? '';
        wb.setCellData(at, r0 + i, c0 + j, { raw, ...(Object.keys(style).length ? { style } : {}) });
      }
    });
    if (dest.newSheet) for (let j = 0; j < w; j++) wb.setColWidth(at, j, j === 0 ? 150 : 100);
    return { si: at, r: r0, c: c0, h, w };
  }, meta());
}

/** 분석 도구 대화상자 공통: fields 는 입력 칸 설명, run(v, get) 은 결과 또는 { write(dest) } */
let anSeq = 0;
function toolDialog(title, fields, run, { output = true, note } = {}) {
  const home = si;
  const inputs = {};
  const gid = `an${++anSeq}`; // 라디오 묶음 이름 (대화상자마다 따로)
  const cur = usedClip(sel);
  const selRef = isSingle(cur) ? '' : refText(cur);
  const line = (label, input, extra) => el('label', { class: 'an-row' }, el('span', {}, label), input, extra ?? null);
  const body = el('div', { class: 'an-dlg' }, note ? el('div', { class: 'muted' }, note) : null, fields.map((f) => {
    if (f.k === 'ref') { inputs[f.name] = refInput(f.value ?? (f.fill ? selRef : ''), home); return line(f.label, inputs[f.name]); }
    if (f.k === 'num') { inputs[f.name] = el('input', { type: 'number', value: f.value ?? '', step: f.step ?? 'any' }); return line(f.label, inputs[f.name]); }
    if (f.k === 'text') { inputs[f.name] = el('input', { type: 'text', value: f.value ?? '' }); return line(f.label, inputs[f.name]); }
    if (f.k === 'check') { inputs[f.name] = el('input', { type: 'checkbox', checked: !!f.value }); return el('label', { class: 'an-check' }, inputs[f.name], f.label); }
    if (f.k === 'checknum') {
      inputs[f.name] = el('input', { type: 'checkbox', checked: !!f.value });
      inputs[`${f.name}N`] = el('input', { type: 'number', value: f.num, step: 'any', style: { width: '70px' } });
      return el('label', { class: 'an-check' }, inputs[f.name], f.label, inputs[`${f.name}N`], f.unit ?? '');
    }
    if (f.k === 'select') { inputs[f.name] = el('select', {}, f.options.map(([v, l]) => el('option', { value: v, selected: v === f.value }, l))); return line(f.label, inputs[f.name]); }
    if (f.k === 'group') {
      inputs.byRows = el('input', { type: 'radio', name: `${gid}-grp` });
      const cols = el('input', { type: 'radio', name: `${gid}-grp`, checked: true });
      return el('div', { class: 'an-row' }, el('span', {}, '데이터 방향:'), el('label', { class: 'an-check' }, cols, '열'), el('label', { class: 'an-check' }, inputs.byRows, '행'));
    }
    if (f.k === 'head') return el('div', { class: 'an-head' }, f.label);
    return null;
  }), output ? [
    el('div', { class: 'an-head' }, '출력 옵션'),
    el('label', { class: 'an-check' }, inputs.outRange = el('input', { type: 'radio', name: `${gid}-out` }), '출력 범위:', inputs.outRef = refInput('', home)),
    el('label', { class: 'an-check' }, inputs.outSheet = el('input', { type: 'radio', name: `${gid}-out`, checked: true }), '새로운 워크시트:', inputs.sheetName = el('input', { type: 'text', value: '', placeholder: title.split(':')[0] })),
  ] : null);
  if (output) inputs.outRef.addEventListener('focus', () => { inputs.outRange.checked = true; });
  const read = () => Object.fromEntries(Object.entries(inputs).map(([k, i]) => [k, i.type === 'checkbox' || i.type === 'radio' ? i.checked : i.type === 'number' ? (i.value === '' ? null : Number(i.value)) : i.value]));
  const get = (name, label) => {
    const p = parseRefInput(inputs[name].value);
    if (!p) throw new AnalysisError(`${label}가 올바른 범위가 아닙니다.`);
    return p;
  };
  const submit = async () => {
    const v = read();
    let dest = null;
    try {
      if (output) {
        if (v.outRange) {
          const p = parseRefInput(v.outRef);
          if (!p) throw new AnalysisError('출력 범위를 지정하세요.');
          dest = { si: p.si, r: p.rg.r1, c: p.rg.c1 };
        } else dest = { newSheet: v.sheetName.trim() || title.split(':')[0].replace(/[\\/?*[\]]/g, '') };
      }
      const res = run(v, get, dest);
      if (!res) return;
      if (dest && !dest.newSheet) {
        const h = res.rows.length;
        const w = Math.max(1, ...res.rows.map((r) => r.length));
        let busy = false;
        for (let r = dest.r; r < dest.r + h && !busy; r++) for (let c = dest.c; c < dest.c + w; c++) if (wb.getCell(dest.si, r, c)?.raw) { busy = true; break; }
        if (busy && !(await confirmBox(title, '출력 범위에 이미 데이터가 있습니다. 덮어쓰시겠습니까?'))) return;
      }
      dlg.close();
      const out = dest ? writeAnalysis(res, dest) : null;
      if (out) {
        if (out.si !== si) switchSheet(out.si, false);
        selectRange({ r1: out.r, c1: out.c, r2: out.r + out.h - 1, c2: out.c + out.w - 1 }, 'cells', { r: out.r, c: out.c });
        res.after?.(out);
      }
    } catch (e) {
      if (e instanceof AnalysisError) alertDialog(title, e.message);
      else throw e;
    }
  };
  const dlg = openDialog({
    title, body, width: 470, modeless: true, onClose: () => { refPick = null; },
    buttons: [{ label: '확인', primary: true, action: () => { submit(); return false; } }, { label: '취소' }],
  });
  return dlg;
}

/** 입력 범위 → 변수 목록 (이름표 · 방향 옵션) */
function anGroups(v, get, name = 'input', label = '입력 범위') {
  const p = get(name, label);
  const g = splitGroups(readBlock(p), { byRows: !!v.byRows, labels: !!v.labels });
  if (!g.some((x) => x.vals.length)) throw new AnalysisError(`${label}에 숫자가 없습니다.`);
  return { p, g };
}
/** 한 줄(열 또는 행) 입력 → 값 목록과 이름표 */
function vectorOf2(v, get, name, label) {
  const p = get(name, label);
  const rows = readBlock(p);
  const byRow = p.rg.r1 === p.rg.r2 && p.rg.c1 !== p.rg.c2;
  if (!byRow && p.rg.c1 !== p.rg.c2) throw new AnalysisError(`${label}는 한 열 또는 한 행이어야 합니다.`);
  let vals = byRow ? rows[0] : rows.map((r) => r[0]);
  let lab = null;
  if (v.labels) { lab = vals[0] === null ? null : String(vals[0]); vals = vals.slice(1); }
  return { p, vals, label: lab, byRow };
}

/** 이동 평균 · 지수 평활: 입력 셀 주소 / 출력 셀 주소 → 수식 + 차트 */
function smoothingOutput(v, get, dest, kind) {
  const { p, vals, byRow } = vectorOf2(v, get, 'input', '입력 범위');
  if (vals.length < 3) throw new AnalysisError('입력 범위에 값이 세 개 이상 있어야 합니다.');
  const skip = v.labels ? 1 : 0;
  const outSheet = dest.newSheet ? null : dest.si;
  const inCell = (i) => (byRow ? { r: p.rg.r1, c: p.rg.c1 + skip + i } : { r: p.rg.r1 + skip + i, c: p.rg.c1 });
  const sameSheet = outSheet === p.si;
  const inRef = (i) => { const q = inCell(i); return `${sameSheet ? '' : `${quoteSheetName(wb.sheets[p.si].name)}!`}${cellName(q.r, q.c)}`; };
  const r0 = dest.newSheet ? 0 : dest.r;
  const c0 = dest.newSheet ? 0 : dest.c;
  const outRef = (i) => cellName(r0 + i, c0);
  const res = kind === 'movavg'
    ? movingAverage(vals.length, inRef, outRef, { interval: Math.round(v.interval ?? 3), stdErr: v.stdErr })
    : expSmoothing(vals.length, inRef, outRef, { damping: v.damping ?? 0.3, stdErr: v.stdErr });
  if (v.chart) {
    res.after = (out) => {
      const n = vals.length;
      const inS = wb.sheets[p.si].name;
      const cat = null;
      const actual = byRow ? { sheet: inS, r1: p.rg.r1, c1: p.rg.c1 + skip, r2: p.rg.r1, c2: p.rg.c1 + skip + n - 1 } : { sheet: inS, r1: p.rg.r1 + skip, c1: p.rg.c1, r2: p.rg.r1 + skip + n - 1, c2: p.rg.c1 };
      const fc = { sheet: wb.sheets[out.si].name, r1: out.r, c1: out.c, r2: out.r + n - 1, c2: out.c };
      addAnalysisChart(out, kind === 'movavg' ? '이동 평균' : '지수 평활', [
        { name: { text: '실제값' }, cat, val: actual },
        { name: { text: '예측값' }, cat, val: fc },
      ], 'line');
    };
  }
  return res;
}

function addAnalysisChart(out, title, series, type = 'line', place = null, extra = null) {
  const box = gv.sheetRect({ r1: out.r, c1: out.c, r2: out.r + out.h - 1, c2: out.c + out.w - 1 });
  const id = `ch${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;
  const chart = { id, type, title, series, x: place?.x ?? box.x + box.w + 24, y: place?.y ?? box.y, w: 480, h: 288, z: nextZ(), legend: 'b', ...(extra ?? {}) };
  wb.transact(() => wb.setSheetProp(out.si, 'charts', [...(wb.sheets[out.si].charts ?? []).map((c) => ({ ...c })), chart]), meta());
  gv.renderObjectsAll();
}

const ALPHA_FIELD = { k: 'num', name: 'alpha', label: '유의 수준(α):', value: 0.05 };
const TOOL_UI = {
  anova1: () => toolDialog('분산 분석: 일원 배치법', [{ k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'group' }, { k: 'check', name: 'labels', label: '첫째 행 이름표 사용' }, ALPHA_FIELD],
    (v, get) => anova1(anGroups(v, get).g, { alpha: v.alpha ?? 0.05 })),
  anova2: () => toolDialog('분산 분석: 반복 없는 이원 배치법', [{ k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'check', name: 'labels', label: '이름표' }, ALPHA_FIELD],
    (v, get) => {
      const rows = readBlock(get('input', '입력 범위'));
      let body = rows;
      let rowLabels; let colLabels;
      if (v.labels) { colLabels = rows[0].slice(1).map(String); rowLabels = rows.slice(1).map((r) => String(r[0] ?? '')); body = rows.slice(1).map((r) => r.slice(1)); }
      return anova2(body, { alpha: v.alpha ?? 0.05, rowLabels, colLabels });
    }),
  correl: () => toolDialog('상관 분석', [{ k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'group' }, { k: 'check', name: 'labels', label: '첫째 행 이름표 사용' }],
    (v, get) => matrixTool(anGroups(v, get).g, 'correl')),
  covar: () => toolDialog('공분산 분석', [{ k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'group' }, { k: 'check', name: 'labels', label: '첫째 행 이름표 사용' }],
    (v, get) => matrixTool(anGroups(v, get).g, 'covar')),
  descr: () => toolDialog('기술 통계법', [
    { k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'group' }, { k: 'check', name: 'labels', label: '첫째 행 이름표 사용' },
    { k: 'head', label: '출력' }, { k: 'check', name: 'summary', label: '요약 통계량', value: true },
    { k: 'checknum', name: 'conf', label: '평균에 대한 신뢰 수준', num: 95, unit: '%' },
    { k: 'checknum', name: 'kth', label: 'K번째 큰 값', num: 1 }, { k: 'checknum', name: 'kthS', label: 'K번째 작은 값', num: 1 },
  ], (v, get) => {
    if (!v.summary && !v.conf && !v.kth && !v.kthS) throw new AnalysisError('출력할 통계를 하나 이상 고르세요.');
    return descriptive(anGroups(v, get).g, { summary: v.summary, conf: v.conf ? (v.confN ?? 95) / 100 : 0, kth: v.kth ? v.kthN : 0, kthSmall: v.kthS ? v.kthSN : 0 });
  }),
  expsmooth: () => toolDialog('지수 평활법', [
    { k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'num', name: 'damping', label: '감쇠 인수:', value: 0.3 }, { k: 'check', name: 'labels', label: '이름표' },
    { k: 'check', name: 'chart', label: '차트 출력', value: true }, { k: 'check', name: 'stdErr', label: '표준 오차' },
  ], (v, get, dest) => smoothingOutput(v, get, dest, 'expsmooth')),
  ftest: () => toolDialog('F-검정: 분산에 대한 두 집단', [{ k: 'ref', name: 'in1', label: '변수 1 입력 범위:' }, { k: 'ref', name: 'in2', label: '변수 2 입력 범위:' }, { k: 'check', name: 'labels', label: '이름표' }, ALPHA_FIELD],
    (v, get) => { const a = vectorOf2(v, get, 'in1', '변수 1 입력 범위'); const b = vectorOf2(v, get, 'in2', '변수 2 입력 범위'); return fTest(a.vals, b.vals, { alpha: v.alpha ?? 0.05, labels: [a.label ?? '변수 1', b.label ?? '변수 2'] }); }),
  hist: () => toolDialog('히스토그램', [
    { k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'ref', name: 'bins', label: '계급 구간:' }, { k: 'check', name: 'labels', label: '이름표' },
    { k: 'check', name: 'pareto', label: '파레토(정렬된 히스토그램)' }, { k: 'check', name: 'cumulative', label: '누적 백분율' }, { k: 'check', name: 'chart', label: '차트 출력', value: true },
  ], (v, get) => {
    const { g } = anGroups({ ...v, byRows: false }, get);
    const vals = g.flatMap((x) => x.vals);
    let bins = null;
    if (v.bins.trim()) { const p = get('bins', '계급 구간'); bins = readBlock(p).flat().filter((x) => typeof x === 'number'); if (!bins.length) throw new AnalysisError('계급 구간에 숫자가 없습니다.'); }
    const res = histogram(vals, bins, { pareto: v.pareto, cumulative: v.cumulative });
    if (v.chart) {
      res.after = (out) => {
        const S = wb.sheets[out.si].name;
        const n = res.rows.length - 1;
        const cat = { sheet: S, r1: out.r + 1, c1: out.c, r2: out.r + n, c2: out.c };
        const series = [{ name: { text: '빈도수' }, cat, val: { sheet: S, r1: out.r + 1, c1: out.c + 1, r2: out.r + n, c2: out.c + 1 } }];
        if (v.cumulative) series.push({ name: { text: '누적 %' }, cat, val: { sheet: S, r1: out.r + 1, c1: out.c + 2, r2: out.r + n, c2: out.c + 2 } });
        addAnalysisChart(out, '히스토그램', series, v.cumulative ? 'combo' : 'column', null, v.cumulative ? { seriesFmt: [{ type: 'column', axis: 0 }, { type: 'line', axis: 1, numFmt: '0%' }] } : null);
      };
    }
    return res;
  }),
  movavg: () => toolDialog('이동 평균법', [
    { k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'check', name: 'labels', label: '첫째 행 이름표 사용' }, { k: 'num', name: 'interval', label: '구간:', value: 3, step: 1 },
    { k: 'check', name: 'chart', label: '차트 출력', value: true }, { k: 'check', name: 'stdErr', label: '표준 오차' },
  ], (v, get, dest) => smoothingOutput(v, get, dest, 'movavg')),
  random: () => toolDialog('난수 생성', [
    { k: 'num', name: 'vars', label: '변수의 개수:', value: 1, step: 1 }, { k: 'num', name: 'count', label: '난수의 개수:', value: 10, step: 1 },
    { k: 'select', name: 'dist', label: '분포:', value: 'uniform', options: [['uniform', '균일 (시작 ~ 끝)'], ['normal', '정규 (평균, 표준 편차)'], ['bernoulli', '베르누이 (p)'], ['binomial', '이항 (p, 시행 횟수)'], ['poisson', '포아송 (λ)']] },
    { k: 'num', name: 'p1', label: '매개 변수 1:', value: 0 }, { k: 'num', name: 'p2', label: '매개 변수 2:', value: 1 }, { k: 'num', name: 'seed', label: '난수 시드 (선택):', value: '', step: 1 },
  ], (v) => randomNumbers({ vars: Math.round(v.vars ?? 1), count: Math.round(v.count ?? 10), dist: v.dist, p1: v.p1 ?? 0, p2: v.p2 ?? 1, seed: v.seed ?? 0 }),
  { note: '균일: 매개 변수 1~2 사이 · 정규: 평균, 표준 편차 · 베르누이: 성공 확률 · 이항: 성공 확률, 시행 횟수 · 포아송: 평균(λ)' }),
  rank: () => toolDialog('순위와 백분율', [{ k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'group' }, { k: 'check', name: 'labels', label: '첫째 행 이름표 사용' }],
    (v, get) => rankPercentile(anGroups(v, get).g)),
  regress: () => toolDialog('회귀 분석', [
    { k: 'ref', name: 'inY', label: 'Y축 입력 범위:' }, { k: 'ref', name: 'inX', label: 'X축 입력 범위:' },
    { k: 'check', name: 'labels', label: '이름표' }, { k: 'check', name: 'zero', label: '상수에 0을 사용' }, { k: 'checknum', name: 'conf', label: '신뢰 수준', num: 95, unit: '%' },
    { k: 'head', label: '잔차' }, { k: 'check', name: 'residuals', label: '잔차' }, { k: 'check', name: 'stdRes', label: '표준 잔차' }, { k: 'check', name: 'fitChart', label: '선 적합도 차트' },
  ], (v, get) => {
    const y = vectorOf2(v, get, 'inY', 'Y축 입력 범위');
    const px = get('inX', 'X축 입력 범위');
    const xg = splitGroups(readBlock(px), { labels: !!v.labels });
    if (xg.length > 16) throw new AnalysisError('X축 입력 범위는 16열 이하여야 합니다.');
    const res = regression(y.vals, xg.map((g) => g.raw), {
      constant: !v.zero, conf: v.conf ? (v.confN ?? 95) / 100 : 0.95, residuals: v.residuals || v.fitChart, stdResiduals: v.stdRes,
      xLabels: v.labels ? xg.map((g) => g.label) : null, yLabel: y.label,
    });
    if (v.fitChart) {
      res.after = (out) => {
        const S = wb.sheets[out.si].name;
        const start = res.rows.findIndex((r) => r[0] === '관측수' && typeof r[1] === 'string');
        const n = y.vals.length;
        const xs = { sheet: wb.sheets[px.si].name, r1: px.rg.r1 + (v.labels ? 1 : 0), c1: px.rg.c1, r2: px.rg.r1 + (v.labels ? 1 : 0) + n - 1, c2: px.rg.c1 };
        const ys = { sheet: wb.sheets[y.p.si].name, ...(y.byRow ? { r1: y.p.rg.r1, c1: y.p.rg.c1 + (v.labels ? 1 : 0), r2: y.p.rg.r1, c2: y.p.rg.c1 + (v.labels ? 1 : 0) + n - 1 } : { r1: y.p.rg.r1 + (v.labels ? 1 : 0), c1: y.p.rg.c1, r2: y.p.rg.r1 + (v.labels ? 1 : 0) + n - 1, c2: y.p.rg.c1 }) };
        const fit = { sheet: S, r1: out.r + start + 1, c1: out.c + 1, r2: out.r + start + n, c2: out.c + 1 };
        addAnalysisChart(out, `${xg[0].label} 선 적합도 차트`, [{ name: { text: y.label ?? 'Y' }, x: xs, val: ys }, { name: { text: '예측치 Y' }, x: xs, val: fit }], 'scatter', { x: gv.sheetRect({ r1: 0, c1: out.c + 10, r2: 0, c2: out.c + 10 }).x, y: 0 });
      };
    }
    return res;
  }),
  sampling: () => toolDialog('표본 추출', [
    { k: 'ref', name: 'input', label: '입력 범위:', fill: true }, { k: 'check', name: 'labels', label: '이름표' },
    { k: 'select', name: 'method', label: '표본 추출 방법:', value: 'random', options: [['periodic', '주기 (N 번째마다)'], ['random', '무작위 (N 개)']] }, { k: 'num', name: 'n', label: 'N:', value: 10, step: 1 },
  ], (v, get) => { const { g } = anGroups({ ...v, byRows: false }, get); return sampling(g.flatMap((x) => x.vals), { method: v.method, n: Math.round(v.n ?? 1) }); }),
  ttestPaired: () => ttestUi('t-검정: 쌍체 비교', 'paired'),
  ttestEq: () => ttestUi('t-검정: 등분산 가정 두 집단', 'equal'),
  ttestUneq: () => ttestUi('t-검정: 이분산 가정 두 집단', 'unequal'),
  ztest: () => toolDialog('z-검정: 평균에 대한 두 집단', [
    { k: 'ref', name: 'in1', label: '변수 1 입력 범위:' }, { k: 'ref', name: 'in2', label: '변수 2 입력 범위:' }, { k: 'num', name: 'hyp', label: '가설 평균차:', value: 0 },
    { k: 'num', name: 'var1', label: '변수 1의 분산(알려진 값):', value: '' }, { k: 'num', name: 'var2', label: '변수 2의 분산(알려진 값):', value: '' }, { k: 'check', name: 'labels', label: '이름표' }, ALPHA_FIELD,
  ], (v, get) => { const a = vectorOf2(v, get, 'in1', '변수 1 입력 범위'); const b = vectorOf2(v, get, 'in2', '변수 2 입력 범위'); return zTest(a.vals, b.vals, { var1: v.var1, var2: v.var2, hyp: v.hyp ?? 0, alpha: v.alpha ?? 0.05, labels: [a.label ?? '변수 1', b.label ?? '변수 2'] }); }),
};
function ttestUi(title, kind) {
  return toolDialog(title, [
    { k: 'ref', name: 'in1', label: '변수 1 입력 범위:' }, { k: 'ref', name: 'in2', label: '변수 2 입력 범위:' }, { k: 'num', name: 'hyp', label: '가설 평균차:', value: 0 },
    { k: 'check', name: 'labels', label: '이름표' }, ALPHA_FIELD,
  ], (v, get) => { const a = vectorOf2(v, get, 'in1', '변수 1 입력 범위'); const b = vectorOf2(v, get, 'in2', '변수 2 입력 범위'); return tTest(a.vals, b.vals, { kind, hyp: v.hyp ?? 0, alpha: v.alpha ?? 0.05, labels: [a.label ?? '변수 1', b.label ?? '변수 2'] }); });
}

/** [데이터] → [데이터 분석]: 도구 고르기 */
function dataAnalysisDialog() {
  const list = el('select', { size: 12, class: 'an-list' }, ANALYSIS_TOOLS.map((t, i) => el('option', { value: t.id, selected: i === 0 }, t.label)));
  const go = () => { const id = list.value; if (!id) return false; setTimeout(() => TOOL_UI[id]()); return true; };
  list.addEventListener('dblclick', () => { if (go()) dlg.close(); });
  const dlg = openDialog({
    title: '통계 데이터 분석', width: 380,
    body: el('div', {}, el('div', { class: 'muted', style: { marginBottom: '6px' } }, '분석 도구'), list,
      el('div', { class: 'muted', style: { marginTop: '8px' } }, '예측 시트 · 시나리오 관리자는 [데이터] 탭의 [예측] 그룹에 있습니다.')),
    buttons: [{ label: '확인', primary: true, action: go }, { label: '취소' }],
  });
  list.focus();
}

/** [예측 시트]: 시간 표시줄 + 값 → 새 시트에 FORECAST.ETS 수식 표와 차트 */
function forecastSheetDialog() {
  let rg = usedClip(sel);
  if (isSingle(rg)) rg = dataRange();
  if (rg.c2 - rg.c1 < 1) { alertDialog('예측 시트', '시간 표시줄과 값이 들어 있는 두 열을 선택하세요.'); return; }
  const tc = rg.c1; const vc = rg.c1 + 1;
  const header = typeof valueAt(rg.r1, tc) !== 'number';
  const r1 = rg.r1 + (header ? 1 : 0);
  const times = []; const vals = [];
  for (let r = r1; r <= rg.r2; r++) { const t = valueAt(r, tc); if (t === null || t === '' || t === undefined) continue; times.push(t); vals.push(valueAt(r, vc)); }
  if (times.length < 3 || times.some((t) => typeof t !== 'number')) { alertDialog('예측 시트', '시간 표시줄(첫 열)은 날짜나 숫자여야 하고, 값이 세 개 이상 있어야 합니다.'); return; }
  const r2 = r1 + times.length - 1;
  const sorted = [...new Set(times)].sort((a, b) => a - b);
  const axis = sorted.length > 1 ? timeAxis(sorted) : null;
  if (!axis) { alertDialog('예측 시트', '시간 표시줄의 간격이 일정하지 않습니다 (날짜 · 숫자가 같은 간격이거나 매월 같은 날이어야 합니다).'); return; }
  const kLast = Math.round(axis.pos(sorted[sorted.length - 1]));
  const tStyle = wb.styleAt(si, r1, tc);
  const fmtT = (x) => formatValue(x, tStyle).text;
  const last = sorted[sorted.length - 1];
  const parseT = (s) => { const t = String(s).trim(); if (/^-?\d+(\.\d+)?$/.test(t)) return Number(t); const p = parseInput(t); return typeof p.value === 'number' ? p.value : NaN; };
  const tName = header ? displayText(rg.r1, tc) || '타임라인' : '타임라인';
  const vName = header ? displayText(rg.r1, vc) || '값' : '값';
  formDialog('예측 워크시트 만들기', [
    { name: 'end', label: '예측 종료', value: fmtT(axis.at(kLast + Math.max(3, Math.round(times.length / 3)))) },
    { name: 'type', label: '차트 종류', type: 'select', value: 'line', options: [{ value: 'line', label: '꺾은선형' }, { value: 'column', label: '세로 막대형' }] },
    { name: 'start', label: '예측 시작', value: fmtT(last) },
    { name: 'conf', label: '신뢰 구간 (%)', type: 'number', value: 95 },
    { name: 'seasonMode', label: '계절성', type: 'select', value: 'auto', options: [{ value: 'auto', label: '자동 검색' }, { value: 'manual', label: '수동 설정' }, { value: 'none', label: '없음' }] },
    { name: 'season', label: '계절 주기 (수동)', type: 'number', value: 12 },
    { name: 'fill', label: '누락된 요소 채우기', type: 'select', value: '1', options: [{ value: '1', label: '보간' }, { value: '0', label: '0' }] },
    { name: 'agg', label: '중복 집계', type: 'select', value: '1', options: [['1', '평균'], ['2', '개수'], ['4', '최대값'], ['5', '중앙값'], ['6', '최소값'], ['7', '합계']].map(([value, label]) => ({ value, label })) },
    { name: 'stats', label: '예측 통계 포함', type: 'checkbox', value: false },
  ], (x) => {
    const end = parseT(x.end);
    const start = parseT(x.start);
    const conf = Number(x.conf) / 100;
    if (!(end > last)) { alertDialog('예측 시트', `예측 종료는 마지막 시점(${fmtT(last)})보다 뒤여야 합니다.`); return false; }
    if (!(start >= sorted[0] && start <= last)) { alertDialog('예측 시트', '예측 시작은 시간 표시줄 범위 안이어야 합니다.'); return false; }
    if (!(conf > 0 && conf < 1)) { alertDialog('예측 시트', '신뢰 구간은 0보다 크고 100보다 작아야 합니다.'); return false; }
    const seas = x.seasonMode === 'auto' ? 1 : x.seasonMode === 'none' ? 0 : Math.max(2, Math.round(Number(x.season) || 2));
    const src = quoteSheetName(sheet().name);
    const home = si;
    const n = times.length;
    const future = [];
    for (let k = kLast + 1; future.length < 10000; k++) { const t = Number(axis.at(k).toPrecision(15)); if (t > end + 1e-9) break; future.push(t); }
    const name = nextSheetName(`${sheet().name} 예측`);
    const H = 1 + n + future.length;
    const idx = wb.transact(() => {
      const at = wb.addSheet(name, si + 1);
      const put = (r, c, raw, style) => wb.setCellData(at, r, c, { raw, ...(style ? { style } : {}) });
      const TL = `$A$2:$A$${n + 1}`;
      const VL = `$B$2:$B$${n + 1}`;
      ['타임라인', vName, `예측(${vName})`, `낮은 신뢰 한계(${vName})`, `높은 신뢰 한계(${vName})`].forEach((h, c) => put(0, c, `'${h}`));
      if (tName !== '타임라인') put(0, 0, `'${tName}`);
      for (let i = 0; i < n; i++) {
        // 원본 값은 참조로 연결 (원본이 바뀌면 예측도 바뀜)
        put(1 + i, 0, `=${src}!${cellName(r1 + i, tc)}`, tStyle);
        put(1 + i, 1, `=${src}!${cellName(r1 + i, vc)}`, wb.styleAt(home, r1 + i, vc));
      }
      // 예측 시작 시점 행: 예측 = 실제 값 (선이 이어지도록)
      const si0 = times.findIndex((t) => t >= start);
      const startRow = 1 + (si0 < 0 ? n - 1 : si0);
      for (let r = startRow; r <= n; r++) {
        const own = r === n; // 마지막 실제 값
        if (own) { put(r, 2, `=B${r + 1}`); put(r, 3, `=B${r + 1}`); put(r, 4, `=B${r + 1}`); continue; }
        put(r, 2, `=FORECAST.ETS(A${r + 1},${VL},${TL},${seas},${x.fill},${x.agg})`);
      }
      future.forEach((t, i) => {
        const r = 1 + n + i;
        put(r, 0, String(t), tStyle);
        put(r, 2, `=FORECAST.ETS(A${r + 1},${VL},${TL},${seas},${x.fill},${x.agg})`);
        put(r, 3, `=C${r + 1}-FORECAST.ETS.CONFINT(A${r + 1},${VL},${TL},${conf},${seas},${x.fill},${x.agg})`);
        put(r, 4, `=C${r + 1}+FORECAST.ETS.CONFINT(A${r + 1},${VL},${TL},${conf},${seas},${x.fill},${x.agg})`);
      });
      for (let c = 0; c < 5; c++) wb.setColWidth(at, c, c === 0 ? 110 : 150);
      const t = {
        id: newTableId(), name: nextTableName(wb), r1: 0, c1: 0, r2: H - 1, c2: 4, header: true, totals: false, style: DEFAULT_TABLE_STYLE,
        banded: true, bandedCols: false, firstCol: false, lastCol: false, filter: { criteria: {}, hidden: {} }, totalsFns: {},
      };
      wb.setSheetProp(at, 'tables', [t]);
      if (x.stats) {
        const lab = [['알파', 1], ['베타', 2], ['감마', 3], ['MASE', 4], ['SMAPE', 5], ['MAE', 6], ['RMSE', 7]];
        put(0, 7, "'통계"); put(0, 8, "'값");
        lab.forEach(([l, k], i) => { put(1 + i, 7, `'${l}`); put(1 + i, 8, `=FORECAST.ETS.STAT(${VL},${TL},${k},${seas},${x.fill},${x.agg})`, { numFmt: 'number', decimals: 3 }); });
        put(8, 7, "'계절성"); put(8, 8, `=FORECAST.ETS.SEASONALITY(${VL},${TL},${x.fill},${x.agg})`);
        wb.setColWidth(at, 7, 90);
        wb.setColWidth(at, 5, 20); wb.setColWidth(at, 6, 20);
      }
      const S = name;
      const cat = { sheet: S, r1: 1, c1: 0, r2: H - 1, c2: 0 };
      const ser = (c, nm) => ({ name: { text: nm }, cat, val: { sheet: S, r1: 1, c1: c, r2: H - 1, c2: c } });
      const x0 = 110 + 4 * 150 + 24 + (x.stats ? 3 * 90 : 0);
      wb.setSheetProp(at, 'charts', [{
        id: `ch${Date.now().toString(36)}`, type: x.type, title: `${vName} 예측`, series: [ser(1, vName), ser(2, `예측(${vName})`), ser(3, '낮은 신뢰 한계'), ser(4, '높은 신뢰 한계')],
        seriesFmt: [{}, {}, { color: '#ED7D31' }, { color: '#ED7D31' }], x: x0, y: 10, w: 600, h: 340, z: 1, legend: 'b',
      }]);
      return at;
    }, meta());
    switchSheet(idx, false);
    toast(`예측 시트 '${name}'를 만들었습니다 (FORECAST.ETS · 신뢰 구간 ${Math.round(conf * 100)}%).`);
    return undefined;
  }, { note: `시간 표시줄: ${cellName(r1, tc)}:${cellName(r2, tc)}, 값: ${cellName(r1, vc)}:${cellName(r2, vc)} · 간격 ${axis.months ? `${axis.months}개월` : formatGeneral(axis.step)} · AAA 지수 평활(엑셀 FORECAST.ETS 와 같은 방식)` });
}

// ───── 시나리오 관리자 (가상 분석) ─────
const scenarioList = () => sheet().scenarios ?? [];
function setScenarios(list) { wb.transact(() => wb.setSheetProp(si, 'scenarios', list), meta()); }
const cellsText = (cells) => cells.map((p) => cellName(p.r, p.c)).join(',');
function parseCells(text) {
  const out = [];
  for (const part of String(text).split(/[,;]\s*/)) {
    if (!part.trim()) continue;
    const p = parseRangeName(part.trim().replace(/\$/g, ''));
    if (!p) return null;
    for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) out.push({ r, c });
  }
  return out.length && out.length <= 32 ? out : null;
}
function scenarioManager() {
  const body = el('div', { class: 'an-dlg' });
  let pick = 0;
  const draw = () => {
    const list = scenarioList();
    pick = Math.min(pick, Math.max(0, list.length - 1));
    const lb = el('select', { size: 8, class: 'an-list' }, list.map((sc, i) => el('option', { value: String(i), selected: i === pick }, sc.name)));
    lb.addEventListener('change', () => { pick = Number(lb.value); draw(); });
    lb.addEventListener('dblclick', () => showScenario(pick));
    const cur = list[pick];
    body.replaceChildren(
      el('div', { class: 'muted' }, '시나리오:'),
      el('div', { style: { display: 'flex', gap: '8px' } }, list.length ? lb : el('div', { class: 'an-list muted', style: { padding: '12px', flex: 1 } }, '정의된 시나리오가 없습니다. [추가]를 눌러 시나리오를 만드세요.'),
        el('div', { style: { display: 'flex', flexDirection: 'column', gap: '4px' } },
          el('button', { class: 'btn', onclick: () => editScenario(null, draw) }, '추가...'),
          el('button', { class: 'btn', disabled: !cur, onclick: () => { setScenarios(list.filter((_, i) => i !== pick)); draw(); } }, '삭제'),
          el('button', { class: 'btn', disabled: !cur, onclick: () => editScenario(pick, draw) }, '편집...'),
          el('button', { class: 'btn', disabled: !list.length, onclick: () => scenarioSummary() }, '요약...'))),
      el('div', { class: 'an-row' }, el('span', {}, '변경 셀:'), el('span', {}, cur ? cellsText(cur.cells) : '')),
      el('div', { class: 'an-row' }, el('span', {}, '설명:'), el('span', {}, cur?.comment ?? '')),
    );
  };
  draw();
  const showScenario = (i) => {
    const sc = scenarioList()[i];
    if (!sc) return;
    wb.transact(() => sc.cells.forEach((p, k) => wb.setInput(si, p.r, p.c, String(sc.values[k] ?? ''))), meta());
    toast(`시나리오 '${sc.name}'의 값을 표시했습니다 (Ctrl+Z 로 되돌리기).`);
  };
  openDialog({
    title: '시나리오 관리자', body, width: 460, modeless: true,
    buttons: [{ label: '표시', primary: true, action: () => { showScenario(pick); return false; } }, { label: '닫기' }],
  });
}
function editScenario(idx, done) {
  const list = scenarioList();
  const cur = idx === null ? null : list[idx];
  const selCells = [];
  const rg = usedClip(sel);
  for (let r = rg.r1; r <= rg.r2 && selCells.length < 32; r++) for (let c = rg.c1; c <= rg.c2 && selCells.length < 32; c++) selCells.push({ r, c });
  formDialog(cur ? '시나리오 편집' : '시나리오 추가', [
    { name: 'name', label: '시나리오 이름', value: cur?.name ?? '' },
    { name: 'cells', label: '변경 셀 (예: B1,B3:B5, 최대 32개)', value: cellsText(cur?.cells ?? selCells) },
    { name: 'comment', label: '설명', type: 'textarea', value: cur?.comment ?? `만든 날짜: ${new Date().toLocaleDateString('ko-KR')}` },
    { name: 'locked', label: '변경 금지', type: 'checkbox', value: cur ? cur.locked !== false : true },
  ], (v) => {
    const name = v.name.trim();
    if (!name) { alertDialog('시나리오', '시나리오 이름을 입력하세요.'); return false; }
    if (list.some((s, i) => i !== idx && s.name === name)) { alertDialog('시나리오', '같은 이름의 시나리오가 있습니다.'); return false; }
    const cells = parseCells(v.cells);
    if (!cells) { alertDialog('시나리오', '변경 셀을 올바르게 입력하세요 (최대 32개).'); return false; }
    if (cells.some((p) => wb.getCell(si, p.r, p.c)?.formula)) { alertDialog('시나리오', '변경 셀에는 수식이 아닌 값이 있어야 합니다.'); return false; }
    // 시나리오 값: 현재 값(또는 이전 값) 을 기본으로, 셀마다 입력
    const prevOf = (p) => { const k = cur?.cells.findIndex((q) => q.r === p.r && q.c === p.c) ?? -1; return k >= 0 ? cur.values[k] : wb.getRaw(si, p.r, p.c) ?? ''; };
    setTimeout(() => formDialog('시나리오 값', cells.map((p, k) => ({ name: `v${k}`, label: `${cellName(p.r, p.c)}`, value: String(prevOf(p)) })), (vals) => {
      const sc = { name, comment: v.comment, locked: v.locked, cells, values: cells.map((_, k) => vals[`v${k}`]) };
      setScenarios(idx === null ? [...list, sc] : list.map((s, i) => (i === idx ? sc : s)));
      done?.();
      return true;
    }, { note: '각 변경 셀에 들어갈 값을 입력하세요.' }));
    return true;
  });
}
/** 시나리오 요약: 결과 셀들을 시나리오마다 계산해 새 시트에 표로 */
function scenarioSummary() {
  const list = scenarioList();
  if (!list.length) { alertDialog('시나리오 요약', '시나리오가 없습니다.'); return; }
  const cand = usedClip(sel);
  formDialog('시나리오 요약', [
    { name: 'kind', label: '보고서 종류', type: 'select', value: 'summary', options: [{ value: 'summary', label: '시나리오 요약' }] },
    { name: 'result', label: '결과 셀 (예: C10,C12)', value: isSingle(cand) ? cellName(active.r, active.c) : `${cellName(cand.r1, cand.c1)}:${cellName(cand.r2, cand.c2)}` },
  ], (v) => {
    const results = parseCells(v.result);
    if (!results) { alertDialog('시나리오 요약', '결과 셀을 올바르게 입력하세요.'); return false; }
    const home = si;
    const changing = [];
    for (const sc of list) for (const p of sc.cells) if (!changing.some((q) => q.r === p.r && q.c === p.c)) changing.push(p);
    // 각 시나리오를 잠시 적용해 결과 계산 → 원래 값 복원 (실행 취소 기록 없음)
    const orig = changing.map((p) => cellData(wb.getCell(home, p.r, p.c)));
    const apply = (sc) => changing.forEach((p, k) => {
      const j = sc ? sc.cells.findIndex((q) => q.r === p.r && q.c === p.c) : -1;
      const d = orig[k];
      wb.setCellData(home, p.r, p.c, j >= 0 ? { ...(d ?? {}), raw: String(sc.values[j] ?? '') } : d);
    });
    const cols = [{ name: '현재 값:', sc: null }, ...list.map((sc) => ({ name: sc.name, sc }))];
    const table = cols.map((col) => {
      apply(col.sc);
      return { ch: changing.map((p) => wb.getValue(home, p.r, p.c)), rs: results.map((p) => wb.getValue(home, p.r, p.c)) };
    });
    apply(null);
    const S = sheet().name;
    const name = nextSheetName('시나리오 요약');
    const idx = wb.transact(() => {
      const at = wb.addSheet(name, si + 1);
      const head = { bold: true, fill: '#44546a', color: '#ffffff' };
      const sub = { fill: '#d9d9d9' };
      const put = (r, c, val, style) => wb.setCellData(at, r, c, { raw: val === null || val === undefined ? '' : typeof val === 'number' ? String(val) : isError(val) ? val.code : `'${val}`, ...(style ? { style } : {}) });
      put(1, 1, '시나리오 요약', head);
      cols.forEach((col, j) => put(1, 3 + j, col.name, head));
      put(1, 2, '', head);
      put(2, 1, '변경 셀:', sub); for (let j = 2; j < 3 + cols.length; j++) put(2, j, '', sub);
      changing.forEach((p, i) => { put(3 + i, 2, `$${colToName(p.c)}$${p.r + 1}`); table.forEach((col, j) => put(3 + i, 3 + j, col.ch[i], { ...wb.styleAt(home, p.r, p.c), ...(j ? { fill: '#f2f2f2' } : {}) })); });
      const rr = 3 + changing.length;
      put(rr, 1, '결과 셀:', sub); for (let j = 2; j < 3 + cols.length; j++) put(rr, j, '', sub);
      results.forEach((p, i) => { put(rr + 1 + i, 2, `$${colToName(p.c)}$${p.r + 1}`); table.forEach((col, j) => put(rr + 1 + i, 3 + j, col.rs[i], { ...wb.styleAt(home, p.r, p.c), ...(j ? { fill: '#f2f2f2' } : {}) })); });
      const nr = rr + 2 + results.length;
      put(nr, 1, `참고: 현재 값 열은 요약 보고서를 만들 때의 변경 셀 값입니다 ('${S}' 시트). 각 시나리오의 변경 셀은 회색으로 표시했습니다.`);
      wb.setColWidth(at, 0, 20); wb.setColWidth(at, 1, 110); wb.setColWidth(at, 2, 80);
      for (let j = 0; j < cols.length; j++) wb.setColWidth(at, 3 + j, 110);
      return at;
    }, meta());
    switchSheet(idx, false);
    return true;
  });
}

/** 해 찾기 (Solver 부가 기능): 목표 셀을 최대 · 최소 · 지정 값으로, 변수 셀을 제약 조건 안에서 바꿈 */
function solverDialog() {
  const home = si;
  const body = el('div', { class: 'an-dlg' });
  const target = refInput(cellName(active.r, active.c), home);
  const mode = el('select', {}, el('option', { value: 'max' }, '최대값'), el('option', { value: 'min' }, '최소'), el('option', { value: 'value' }, '지정값'));
  const goal = el('input', { type: 'number', value: 0, step: 'any', style: { width: '90px' } });
  const vars = refInput('', home);
  const cons = el('textarea', { rows: 4, placeholder: '한 줄에 하나씩: B2<=100, B3>=0, B2+B3=50, B4 int' });
  const nonneg = el('input', { type: 'checkbox', checked: true });
  body.append(
    el('label', { class: 'an-row' }, el('span', {}, '목표 설정:'), target),
    el('div', { class: 'an-row' }, el('span', {}, '대상:'), mode, goal),
    el('label', { class: 'an-row' }, el('span', {}, '변수 셀 변경:'), vars),
    el('label', { class: 'an-row', style: { alignItems: 'flex-start' } }, el('span', {}, '제한 조건:'), cons),
    el('label', { class: 'an-check' }, nonneg, '제한되지 않는 변수를 음이 아닌 수로 설정'),
    el('div', { class: 'muted' }, '해법: 비선형 GRG 대신 제약 벌점 + 넬더-미드 · 좌표 탐색 (정수 제약은 반올림 후 재탐색). 수식 셀의 값으로 계산합니다.'),
  );
  const run = () => {
    const t = parseRefInput(target.value);
    const vv = parseRefInput(vars.value);
    if (!t || !vv || t.si !== home || vv.si !== home) { alertDialog('해 찾기', '목표 셀과 변수 셀을 이 시트에서 지정하세요.'); return false; }
    if (!wb.getCell(home, t.rg.r1, t.rg.c1)?.formula) { alertDialog('해 찾기', '목표 셀에는 수식이 있어야 합니다.'); return false; }
    const vcells = [];
    for (let r = vv.rg.r1; r <= vv.rg.r2; r++) for (let c = vv.rg.c1; c <= vv.rg.c2; c++) vcells.push({ r, c });
    if (vcells.length > 50) { alertDialog('해 찾기', '변수 셀은 50개까지 지정할 수 있습니다.'); return false; }
    if (vcells.some((p) => wb.getCell(home, p.r, p.c)?.formula)) { alertDialog('해 찾기', '변수 셀에는 수식이 아닌 값이 있어야 합니다.'); return false; }
    // 제약: "식 연산자 식" → 수식으로 평가 (좌변 - 우변)
    const rules = [];
    const ints = new Set();
    for (const line of cons.value.split(/\n/)) {
      const s = line.trim();
      if (!s) continue;
      const mi = /^(.+?)\s+(int|정수|bin|이진)$/i.exec(s);
      if (mi) {
        const p = parseRangeName(mi[1].replace(/\$/g, ''));
        if (!p) { alertDialog('해 찾기', `제한 조건을 이해할 수 없습니다: ${s}`); return false; }
        for (let r = p.r1; r <= p.r2; r++) for (let c = p.c1; c <= p.c2; c++) {
          const k = vcells.findIndex((q) => q.r === r && q.c === c);
          if (k >= 0) { ints.add(k); if (/bin|이진/i.test(mi[2])) rules.push({ lhs: `=${cellName(r, c)}`, op: '<=', rhs: '=1' }, { lhs: `=${cellName(r, c)}`, op: '>=', rhs: '=0' }); }
        }
        continue;
      }
      const m = /^(.+?)(<=|>=|=)(.+)$/.exec(s);
      if (!m) { alertDialog('해 찾기', `제한 조건을 이해할 수 없습니다: ${s}`); return false; }
      rules.push({ lhs: `=${m[1].trim()}`, op: m[2], rhs: `=${m[3].trim()}` });
    }
    const scratch = { r: wb.usedRange(home).rows + 5, c: 0 };
    const orig = vcells.map((p) => cellData(wb.getCell(home, p.r, p.c)));
    const put = (x) => vcells.forEach((p, k) => wb.setCellData(home, p.r, p.c, { ...(orig[k] ?? {}), raw: String(x[k]) }));
    // 제약식은 빈 칸에 잠시 수식으로 넣어 평가
    const evalExpr = (f) => { wb.setCellData(home, scratch.r, scratch.c, { raw: f }); const v = wb.getValue(home, scratch.r, scratch.c); return typeof v === 'number' ? v : NaN; };
    const objective = (x) => {
      put(x);
      const v = wb.getValue(home, t.rg.r1, t.rg.c1);
      let f = typeof v === 'number' ? v : NaN;
      if (!Number.isFinite(f)) return 1e30;
      f = mode.value === 'max' ? -f : mode.value === 'min' ? f : Math.abs(f - Number(goal.value));
      let pen = 0;
      for (const rl of rules) {
        const d = evalExpr(`=(${rl.lhs.slice(1)})-(${rl.rhs.slice(1)})`);
        if (!Number.isFinite(d)) { pen += 1e6; continue; }
        const viol = rl.op === '<=' ? Math.max(0, d) : rl.op === '>=' ? Math.max(0, -d) : Math.abs(d);
        pen += viol;
      }
      if (nonneg.checked) for (const v2 of x) if (v2 < 0) pen += -v2;
      return { f, pen };
    };
    const x0 = vcells.map((p) => Number(wb.getValue(home, p.r, p.c)) || 0);
    let res;
    try {
      res = solveMin(objective, x0, { ints: [...ints] });
    } finally {
      vcells.forEach((p, k) => wb.setCellData(home, p.r, p.c, orig[k]));
      wb.setCellData(home, scratch.r, scratch.c, null);
    }
    const feasible = res.pen < 1e-6;
    const x = res.x.map((v) => Number(v.toPrecision(12)));
    openDialog({
      title: '해 찾기 결과', width: 420,
      body: el('div', {}, el('p', {}, feasible ? '해 찾기가 해를 찾았습니다. 모든 제한 조건이 만족되었습니다.' : '해 찾기가 모든 제한 조건을 만족하는 해를 찾지 못했습니다 (가장 가까운 해).'),
        el('pre', { class: 'plain' }, vcells.map((p, k) => `${cellName(p.r, p.c)} = ${formatGeneral(x[k])}`).join('\n'))),
      buttons: [
        { label: '해 찾기 해 보관', primary: true, action: () => { wb.transact(() => vcells.forEach((p, k) => wb.setCellData(home, p.r, p.c, { ...(orig[k] ?? {}), raw: String(x[k]) })), meta()); } },
        { label: '원래 값 복원' },
      ],
    });
    return undefined;
  };
  openDialog({ title: '해 찾기 매개 변수', body, width: 520, modeless: true, onClose: () => { refPick = null; }, buttons: [{ label: '해 찾기', primary: true, action: run }, { label: '닫기' }] });
}

/** 이동 옵션 (Ctrl+G → 옵션) */
function gotoSpecialDialog() {
  formDialog('이동 옵션', [
    { name: 'kind', label: '종류', type: 'select', value: 'blanks', options: GOTO_KINDS.map((k) => ({ value: k.id, label: k.label })) },
    { name: 'numbers', label: '숫자 (상수 · 수식)', type: 'checkbox', value: true },
    { name: 'text', label: '텍스트', type: 'checkbox', value: true },
    { name: 'logical', label: '논리값', type: 'checkbox', value: true },
    { name: 'errors', label: '오류', type: 'checkbox', value: true },
  ], (v) => {
    const u = wb.usedRange(si);
    if (v.kind === 'lastCell') { const e = wb.extent(si); selectCell(Math.max(0, e.rows - 1), Math.max(0, e.cols - 1)); return; }
    const rg = selIsActiveOnly() ? { r1: 0, c1: 0, r2: Math.max(0, u.rows - 1), c2: Math.max(0, u.cols - 1) } : usedClip(sel);
    const cond = sheet().cond;
    const cells = specialCells(v.kind, rg, {
      cellAt: (r, c) => wb.getCell(si, r, c), valueAt: (r, c) => wb.getValue(si, r, c), hidden: (r) => gv.rows.size(r) === 0,
      inCond: (r, c) => cond.some((rule) => [rule, ...(rule.more ?? [])].some((g) => r >= g.r1 && r <= g.r2 && c >= g.c1 && c <= g.c2)),
      inValidation: (r, c) => !!validationAt(sheet(), r, c),
      types: { numbers: v.numbers, text: v.text, logical: v.logical, errors: v.errors }, limit: 500000,
    });
    if (!cells.length) { alertDialog('이동 옵션', '해당하는 셀이 없습니다.'); return; }
    const box = { r1: minOf(cells.slice(0, 100000).map((x) => x[0])), c1: minOf(cells.slice(0, 100000).map((x) => x[1])), r2: maxOf(cells.slice(-100000).map((x) => x[0])), c2: maxOf(cells.slice(0, 100000).map((x) => x[1])) };
    keepSpecial = true;
    try { selectRange(box, 'cells', { r: cells[0][0], c: cells[0][1] }); } finally { keepSpecial = false; }
    special = { si, cells };
    gv.ensureVisible(cells[0][0], cells[0][1]);
    gv.renderAll();
    toast(`${cells.length.toLocaleString()}개 칸을 골랐습니다. 입력 후 Ctrl+Enter 로 모두 채우거나 Delete 로 지울 수 있습니다.`);
  });
}

// ───────────────────────── 개요 (그룹 · 부분합) ─────────────────────────
const outlineOf = () => normOutline(sheet().outline);

/** 개요 · 숨김 변경을 한 번에 (실행 취소 한 단위) */
function putOutline(o, hidden = null, ax = 'r') {
  wb.transact(() => {
    wb.setSheetProp(si, 'outline', outlineEmpty(o) ? null : o);
    if (hidden) wb.setSheetProp(si, ax === 'r' ? 'hiddenRows' : 'hiddenCols', hidden);
  }, meta());
  gv.layout();
  renderAll();
}

/** 그룹(delta=1) · 그룹 해제(delta=-1): 행 전체를 골랐으면 행, 열 전체면 열, 아니면 물어봄 */
function outlineGroup(delta) {
  const go = (ax) => {
    const o = outlineOf();
    const k = ax === 'r' ? 'rows' : 'cols';
    const ext = wb.extent(si);
    const [a, b] = ax === 'r' ? [sel.r1, Math.min(sel.r2, Math.max(sel.r1, ext.rows + 1000))] : [sel.c1, Math.min(sel.c2, Math.max(sel.c1, ext.cols + 100))];
    if (delta < 0 && !Object.keys(o[k]).some((i) => +i >= a && +i <= b)) { toast('그룹 해제할 수 없습니다.'); return; }
    if (delta > 0 && maxLevel(o[k]) >= 7 && Object.entries(o[k]).some(([i, v]) => +i >= a && +i <= b && v >= 7)) { toast('개요는 7수준까지 만들 수 있습니다.'); return; }
    o[k] = changeLevels(o[k], a, b, delta);
    // 그룹 해제로 수준이 0 이 된 행 · 열은 다시 보이게
    let hidden = null;
    if (delta < 0) {
      const hk = ax === 'r' ? 'hiddenRows' : 'hiddenCols';
      hidden = { ...sheet()[hk] };
      for (let i = a; i <= b; i++) if (!o[k][i]) delete hidden[i];
    }
    putOutline(o, hidden, ax);
  };
  if (selKind === 'rows') go('r');
  else if (selKind === 'cols') go('c');
  else formDialog(delta > 0 ? '그룹' : '그룹 해제', [{ name: 'ax', label: '대상', type: 'select', value: 'r', options: [{ value: 'r', label: '행' }, { value: 'c', label: '열' }] }], (v) => go(v.ax));
}

function outlineToggle(ax, a, b, level, force = null) {
  const o = outlineOf();
  const isRow = ax === 'r';
  const after = isRow ? o.below : o.right;
  const ck = isRow ? 'rowsColl' : 'colsColl';
  const hk = isRow ? 'hiddenRows' : 'hiddenCols';
  const g = { a, b, level };
  const collapse = force ?? !o[ck][summaryOf(g, after)];
  const res = toggleGroup(o[isRow ? 'rows' : 'cols'], sheet()[hk], o[ck], g, after, collapse);
  putOutline({ ...o, [ck]: res.coll }, res.hidden, ax);
}

function outlineShowLevel(ax, n) {
  const o = outlineOf();
  const isRow = ax === 'r';
  const res = showLevel(o[isRow ? 'rows' : 'cols'], sheet()[isRow ? 'hiddenRows' : 'hiddenCols'], n, isRow ? o.below : o.right);
  putOutline({ ...o, [isRow ? 'rowsColl' : 'colsColl']: res.coll }, res.hidden, ax);
}

/** 세부 정보 표시 · 숨기기: 활성 셀이 들어 있는 가장 안쪽 그룹 (행 우선) */
function outlineDetail(show) {
  const o = outlineOf();
  const gr = groupAt(o.rows, active.r, o.below);
  const gc = gr ? null : groupAt(o.cols, active.c, o.right);
  if (!gr && !gc) { toast('선택한 셀이 그룹 안에 있지 않습니다.'); return; }
  if (gr) outlineToggle('r', gr.a, gr.b, gr.level, !show);
  else outlineToggle('c', gc.a, gc.b, gc.level, !show);
}

function outlineClear() {
  const o = outlineOf();
  if (outlineEmpty(o)) { toast('지울 개요가 없습니다.'); return; }
  const hr = { ...sheet().hiddenRows };
  const hc = { ...sheet().hiddenCols };
  for (const i of Object.keys(o.rows)) delete hr[i];
  for (const i of Object.keys(o.cols)) delete hc[i];
  wb.transact(() => { wb.setSheetProp(si, 'outline', null); wb.setSheetProp(si, 'hiddenRows', hr); wb.setSheetProp(si, 'hiddenCols', hc); }, meta());
  gv.layout();
  renderAll();
}

/** 부분합 행을 모두 제거 (SUBTOTAL 수식이 있는 행 삭제 + 개요 지우기). 범위는 머리글 포함 */
function removeSubtotalRows(rg) {
  let r2 = rg.r2;
  for (let r = rg.r2; r > rg.r1; r--) {
    let isSub = false;
    for (let c = rg.c1; c <= rg.c2 && !isSub; c++) if (/^=SUBTOTAL\(/i.test(wb.getRaw(si, r, c))) isSub = true;
    if (isSub) { wb.deleteRows(si, r, 1); r2--; }
  }
  const o = outlineOf();
  o.rows = changeLevels(o.rows, rg.r1, r2 + 1, -7);
  o.rowsColl = {};
  wb.setSheetProp(si, 'outline', outlineEmpty(o) ? null : o);
  const hr = { ...sheet().hiddenRows };
  for (let r = rg.r1; r <= r2 + 1; r++) delete hr[r];
  wb.setSheetProp(si, 'hiddenRows', hr);
  return { ...rg, r2 };
}

/** 데이터 → 부분합: 그룹 열 값이 바뀔 때마다 요약 행(SUBTOTAL) + 총합계, 3수준 개요 */
function subtotalDialog() {
  if (editing && !commitEdit()) return;
  const rg0 = dataRange();
  if (rg0.r2 <= rg0.r1) { alertDialog('부분합', '머리글 행과 데이터가 있는 범위를 선택하세요.'); return; }
  const heads = [];
  for (let c = rg0.c1; c <= rg0.c2; c++) heads.push({ c, name: displayText(rg0.r1, c) || `${colToName(c)}열` });
  const numeric = (c) => { for (let r = rg0.r1 + 1; r <= Math.min(rg0.r2, rg0.r1 + 50); r++) if (typeof valueAt(r, c) === 'number') return true; return false; };
  const lastNum = [...heads].reverse().find((h) => numeric(h.c));
  const hasSubs = (() => { for (let r = rg0.r1; r <= Math.min(rg0.r2, rg0.r1 + 5000); r++) for (let c = rg0.c1; c <= rg0.c2; c++) if (/^=SUBTOTAL\(/i.test(wb.getRaw(si, r, c))) return true; return false; })();
  const fields = [
    { name: 'key', label: '그룹화할 항목', type: 'select', value: String(heads[0].c), options: heads.map((h) => ({ value: String(h.c), label: h.name })) },
    { name: 'fn', label: '사용할 함수', type: 'select', value: '9', options: SUBTOTAL_FNS.map((f) => ({ value: String(f.id), label: f.label })) },
    ...heads.map((h) => ({ name: `col${h.c}`, label: `부분합 계산 항목: ${h.name}`, type: 'checkbox', value: h === lastNum || (numeric(h.c) && h.c !== heads[0].c && heads.length <= 4) })),
    { name: 'replace', label: '새로운 값으로 대치', type: 'checkbox', value: true },
    { name: 'below', label: '데이터 아래에 요약 표시', type: 'checkbox', value: true },
    ...(hasSubs ? [{ name: 'removeAll', label: '모두 제거 (부분합 행과 개요만 지우기)', type: 'checkbox', value: false }] : []),
  ];
  formDialog('부분합', fields, (v) => {
    const t0 = performance.now();
    let groupsN = 0;
    wb.transact(() => {
      let rg = rg0;
      if (v.replace || v.removeAll) rg = removeSubtotalRows(rg);
      if (v.removeAll) return;
      const keyCol = Number(v.key);
      const fn = Number(v.fn);
      const cols = heads.filter((h) => v[`col${h.c}`]).map((h) => h.c);
      if (!cols.length) return;
      const { groups } = planSubtotals(rg.r1, rg.r2, (r) => displayText(r, keyCol));
      groupsN = groups.length;
      if (groups.length > 20000) { toast('그룹이 너무 많습니다 (2만 개 이하).'); return; }
      const below = !!v.below;
      // 아래에서 위로 행 삽입 (위쪽 행 번호는 그대로)
      for (let i = groups.length - 1; i >= 0; i--) wb.insertRows(si, below ? groups[i].b + 1 : groups[i].a, 1);
      const fnName = SUBTOTAL_FNS.find((f) => f.id === fn)?.label ?? '';
      const bold = { bold: true };
      let levels = { ...(outlineOf().rows) };
      let lastRow = rg.r1;
      groups.forEach((g, i) => {
        const shift = below ? i : i + 1;
        const a = g.a + shift;
        const b = g.b + shift;
        const s = below ? b + 1 : g.a + i;
        lastRow = Math.max(lastRow, b, s);
        wb.setCellData(si, s, keyCol, { raw: `'${g.key} 요약`, style: bold });
        for (const c of cols) if (c !== keyCol) wb.setCellData(si, s, c, { raw: `=SUBTOTAL(${fn},${colToName(c)}${a + 1}:${colToName(c)}${b + 1})`, style: { ...wb.styleAt(si, a, c), ...bold } });
        for (let r = a; r <= b; r++) levels[r] = 2;
        levels[s] = 1;
      });
      // 총합계
      const gRow = lastRow + 1;
      wb.insertRows(si, gRow, 1);
      wb.setCellData(si, gRow, keyCol, { raw: '총합계', style: bold });
      for (const c of cols) if (c !== keyCol) wb.setCellData(si, gRow, c, { raw: `=SUBTOTAL(${fn},${colToName(c)}${rg.r1 + 2}:${colToName(c)}${lastRow + 1})`, style: { ...wb.styleAt(si, rg.r1 + 1, c), ...bold } });
      const o = outlineOf();
      o.rows = levels;
      o.below = below;
      wb.setSheetProp(si, 'outline', o);
      toast(`부분합: 그룹 ${groups.length.toLocaleString()}개 (${fnName}) — ${Math.round(performance.now() - t0)}ms. 왼쪽 1 2 3 단추로 수준을 바꿀 수 있습니다.`);
    }, meta());
    gv.layout();
    renderAll();
  }, { note: `범위 ${colToName(rg0.c1)}${rg0.r1 + 1}:${colToName(rg0.c2)}${rg0.r2 + 1} (첫 행은 머리글). 그룹화할 항목 순서로 미리 정렬해 두세요.` });
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
    alertDialog('WIXEL', '병합된 셀이 있으면 정렬할 수 없습니다.');
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
/**
 * 필터 열 c 의 행 r1~r2 가 열 블록에 있으면 { codes, texts, over } — 표시 글자를 값마다 한 번만 만들어 코드로 비교.
 * over: 블록 위에 일반 셀(수식 · 다른 서식)이 있는 행 (그 행만 따로 확인)
 */
const blockFilterMemo = new WeakMap();
function blockFilterCol(r1, r2, c) {
  const b = wb.blockAt(si, r1, c);
  if (!b || r2 >= b.r0 + b.n) return null;
  const key = `${r1},${r2},${c},${b.ver ?? 0},${wb.sheetVersion(si)}`;
  let memo = blockFilterMemo.get(b);
  if (!memo) { memo = new Map(); blockFilterMemo.set(b, memo); }
  const hit = memo.get(key);
  if (hit) return hit;
  const bc = b.cols[c - b.c0];
  const lc = logicalCol(b, c - b.c0, r1 - b.r0, r2 - r1 + 1); // 정렬 순서가 있으면 보이는 순서로
  const d = blockColumn(lc, lc.a, r2 - r1 + 1).dim();
  const fmt = bc.fmt ?? {};
  const texts = d.keys.map((k) => (k === PIVOT_EMPTY ? '' : formatValue(k, fmt).text));
  const over = [];
  for (const k of sheet().cells.keys()) {
    const i = k.indexOf(',');
    const r = +k.slice(0, i);
    if (r >= r1 && r <= r2 && +k.slice(i + 1) === c) over.push(r);
  }
  const res = { codes: d.codes, texts, over };
  if (memo.size > 20) memo.clear();
  memo.set(key, res);
  return res;
}

// 텍스트 · 숫자 조건 (엑셀의 사용자 지정 자동 필터 + 정규식)
const FILTER_OPS = [
  ['eq', '='], ['ne', '<>'], ['gt', '>'], ['ge', '>='], ['lt', '<'], ['le', '<='],
  ['begins', '시작 문자'], ['notBegins', '시작 문자 아님'], ['ends', '끝 문자'], ['notEnds', '끝 문자 아님'],
  ['contains', '포함'], ['notContains', '포함하지 않음'], ['regex', '정규식과 일치'], ['notRegex', '정규식과 일치하지 않음'],
];
const wildRe = (p) => new RegExp(`^${String(p).replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?')}$`, 'i');
function opTest(op, v, text, arg) {
  if (!op || arg === undefined || arg === null || arg === '') return true;
  const a = String(arg);
  const t = String(text ?? '').toLowerCase();
  const an = Number(a.replace(/,/g, ''));
  const num = typeof v === 'number' && a.trim() !== '' && Number.isFinite(an);
  switch (op) {
    case 'eq': return num ? v === an : wildRe(a).test(String(text ?? ''));
    case 'ne': return num ? v !== an : !wildRe(a).test(String(text ?? ''));
    case 'gt': return num ? v > an : t > a.toLowerCase();
    case 'ge': return num ? v >= an : t >= a.toLowerCase();
    case 'lt': return num ? v < an : t < a.toLowerCase();
    case 'le': return num ? v <= an : t <= a.toLowerCase();
    case 'begins': return t.startsWith(a.toLowerCase());
    case 'notBegins': return !t.startsWith(a.toLowerCase());
    case 'ends': return t.endsWith(a.toLowerCase());
    case 'notEnds': return !t.endsWith(a.toLowerCase());
    case 'contains': return wildRe(`*${a}*`).test(String(text ?? ''));
    case 'notContains': return !wildRe(`*${a}*`).test(String(text ?? ''));
    case 'regex': case 'notRegex': {
      let re;
      try { re = new RegExp(a, 'iu'); } catch { return true; }
      return re.test(String(text ?? '')) === (op === 'regex');
    }
    default: return true;
  }
}
/** 조건 객체(색 · 정규식 · 사용자 지정 · 상위 N · 평균) → 행 판정 함수 */
function critPredicate(c, cr, r1, r2) {
  if (cr.type === 'fill') return (r) => (styleAt(r, c)?.fill ?? '') === (cr.value ?? '');
  if (cr.type === 'font') return (r) => (styleAt(r, c)?.color ?? '') === (cr.value ?? '');
  if (cr.type === 'custom') {
    return (r) => {
      const v = valueAt(r, c);
      const t = displayText(r, c);
      const a = opTest(cr.op1, v, t, cr.v1);
      if (!cr.op2 || cr.v2 === undefined || cr.v2 === '') return a;
      const b = opTest(cr.op2, v, t, cr.v2);
      return cr.join === 'or' ? a || b : a && b;
    };
  }
  if (cr.type === 'top' || cr.type === 'avg') {
    const nums = [];
    for (let r = r1; r <= r2; r++) { const v = valueAt(r, c); if (typeof v === 'number') nums.push(v); }
    if (!nums.length) return () => true;
    if (cr.type === 'avg') {
      const avg = nums.reduce((x, y) => x + y, 0) / nums.length;
      return (r) => { const v = valueAt(r, c); return typeof v === 'number' && (cr.above ? v > avg : v < avg); };
    }
    nums.sort((x, y) => (cr.bottom ? x - y : y - x));
    const k = Math.max(1, Math.min(nums.length, cr.percent ? Math.ceil((nums.length * cr.n) / 100) : cr.n));
    const th = nums[k - 1];
    return (r) => { const v = valueAt(r, c); return typeof v === 'number' && (cr.bottom ? v <= th : v >= th); };
  }
  return () => true;
}

function recomputeFilter(f, key = '') {
  const r2 = key ? f.r2 : Math.max(f.r2, currentRegion(f.r1, f.c1).r2);
  const crit = Object.entries(f.criteria ?? {}).filter(([, v]) => Array.isArray(v)).map(([c, vals]) => [Number(c), new Set(vals)]);
  const preds = Object.entries(f.criteria ?? {}).filter(([, v]) => v && !Array.isArray(v) && typeof v === 'object').map(([c, cr]) => critPredicate(Number(c), cr, f.r1 + 1, r2));
  const n = r2 - f.r1;
  if (preds.length) {
    // 색 · 조건 필터가 있으면 행마다 판정
    const hidden = {};
    for (let r = f.r1 + 1; r <= r2; r++) {
      if (crit.some(([c, allowed]) => !allowed.has(displayText(r, c))) || preds.some((p) => !p(r))) hidden[r] = true;
    }
    return { ...f, r2, hidden };
  }
  if (crit.length && n > 50000) {
    // 행이 아주 많으면: 숨긴 행을 비트맵으로, 열 블록은 값(코드)마다 한 번만 판정
    const start = f.r1 + 1;
    const bits = new Uint8Array(n);
    const overRows = new Set();
    for (const [c, allowed] of crit) {
      const bf = blockFilterCol(start, r2, c);
      if (bf) {
        const ok = Uint8Array.from(bf.texts, (t) => (allowed.has(t) ? 1 : 0));
        const codes = bf.codes;
        for (let i = 0; i < n; i++) if (!ok[codes[i]]) bits[i] = 1;
        for (const r of bf.over) overRows.add(r);
      } else {
        for (let i = 0; i < n; i++) if (!bits[i] && !allowed.has(displayText(start + i, c))) bits[i] = 1;
      }
    }
    // 블록 위의 일반 셀이 있는 행은 모든 조건을 표시 글자로 다시 확인
    for (const r of overRows) bits[r - start] = crit.every(([c, allowed]) => allowed.has(displayText(r, c))) ? 0 : 1;
    let count = 0;
    for (let i = 0; i < n; i++) count += bits[i];
    return { ...f, r2, hidden: count ? { __bits: bits, start, count } : {} };
  }
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
  if (rg.r1 === rg.r2 && isEmptyAt(rg.r1, rg.c1)) { alertDialog('WIXEL', '필터를 적용할 데이터 범위를 선택하세요.'); return; }
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
  toast(`${total.toLocaleString()}개 중 ${(total - hidCount(nf.hidden)).toLocaleString()}개의 레코드가 있습니다.`);
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
  const bfc = full.r2 - f.r1 > 50000 ? blockFilterCol(f.r1 + 1, full.r2, c) : null;
  const bfo = bfc ? others.map(([k, allowed]) => [blockFilterCol(f.r1 + 1, full.r2, k), allowed]) : [];
  if (bfc && !bfc.over.length && bfo.every(([x]) => x && !x.over.length)) {
    // 행이 아주 많은 열 블록: 코드 단위로 (다른 열 조건을 통과한 행의 값만)
    const n = full.r2 - f.r1;
    const oks = bfo.map(([x, allowed]) => [x.codes, Uint8Array.from(x.texts, (t) => (allowed.has(t) ? 1 : 0))]);
    const seen = new Uint8Array(bfc.texts.length);
    let left = seen.length;
    for (let i = 0; i < n && left; i++) {
      let pass = true;
      for (const [codes, ok] of oks) if (!ok[codes[i]]) { pass = false; break; }
      if (!pass) continue;
      const code = bfc.codes[i];
      if (!seen[code]) { seen[code] = 1; left--; values.set(bfc.texts[code], valueAt(f.r1 + 1 + i, c)); }
    }
  } else {
    for (let r = f.r1 + 1; r <= full.r2; r++) {
      if (others.some(([k, allowed]) => !allowed.has(displayText(r, k)))) continue;
      const t = displayText(r, c);
      if (!values.has(t)) values.set(t, valueAt(r, c));
    }
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
  const hasCrit = !!f.criteria?.[c];
  // 색 기준 필터: 이 열에 쓰인 채우기 색 · 글꼴 색
  const fills = new Map();
  const fonts = new Map();
  for (let r = f.r1 + 1; r <= Math.min(full.r2, f.r1 + 20000); r++) {
    const st = styleAt(r, c);
    fills.set(st?.fill ?? '', (fills.get(st?.fill ?? '') ?? 0) + 1);
    if (st?.color) fonts.set(st.color, (fonts.get(st.color) ?? 0) + 1);
  }
  const numeric = items.length && items.filter((t) => t !== '').every((t) => typeof values.get(t) === 'number');
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
  // 엑셀: 검색하면 [필터에 현재 선택 내용 추가] — 켜면 지금 필터에 검색 결과를 더함 (여러 번 검색해 누적 선택)
  const addCb = el('input', { type: 'checkbox' });
  const addRow = el('label', { class: 'filter-add', style: { display: 'none' } }, addCb, '필터에 현재 선택 내용 추가');
  const allLabel = list.firstChild;
  search.addEventListener('input', () => {
    const q = search.value.trim().toLowerCase();
    for (const [t, cb] of checks) {
      const show = !q || t.toLowerCase().includes(q);
      cb.parentElement.style.display = show ? '' : 'none';
      if (q) cb.checked = show;
      else cb.checked = !current || current.has(t);
    }
    addRow.style.display = q ? '' : 'none';
    allLabel.lastChild.textContent = q ? '(검색 결과 모두 선택)' : '(모두 선택)';
    syncAll();
  });
  const ok = () => {
    closeMenus();
    const q = search.value.trim();
    let chosen = [...checks.entries()].filter(([, cb]) => cb.checked && (!q || cb.parentElement.style.display !== 'none')).map(([t]) => t);
    if (q && addCb.checked) chosen = [...new Set([...(current ? [...current] : items), ...chosen])];
    applyFilterCriteria(c, chosen.length === items.length ? null : chosen, key);
    focusGrid();
  };
  const header = displayText(f.r1, c) || `${colToName(c)}열`;
  const node = el('div', {
    class: 'filter-menu',
    onkeydown: (e) => { e.stopPropagation(); if (e.key === 'Enter') ok(); if (e.key === 'Escape') { closeMenus(); focusGrid(); } },
  },
  search, list, addRow,
  el('div', { class: 'filter-foot' },
    el('button', { class: 'btn primary', onclick: ok }, '확인'),
    el('button', { class: 'btn', onclick: () => { closeMenus(); focusGrid(); } }, '취소')));
  const menu = openMenu(anchorEl, [
    { label: '텍스트 오름차순 정렬', icon: 'sortAsc', action: () => sortData(true, c, true, full, key) },
    { label: '텍스트 내림차순 정렬', icon: 'sortDesc', action: () => sortData(false, c, true, full, key) },
    { sep: true },
    { label: `"${header}"에서 필터 해제`, icon: 'filterClear', disabled: !hasCrit, action: () => applyFilterCriteria(c, null, key) },
    {
      label: '색 기준 필터', icon: 'fill', disabled: fills.size <= 1 && !fonts.size, submenu: [
        { label: '셀 색 기준 필터', header: true, disabled: true },
        ...[...fills.keys()].map((col) => ({ label: col ? `■ ${col}` : '채우기 없음', swatch: col || null, action: () => applyFilterCriteria(c, { type: 'fill', value: col }, key) })),
        ...(fonts.size ? [{ sep: true }, { label: '글꼴 색 기준 필터', header: true, disabled: true }, ...[...fonts.keys()].map((col) => ({ label: `A ${col}`, swatch: col, action: () => applyFilterCriteria(c, { type: 'font', value: col }, key) }))] : []),
      ],
    },
    { label: numeric ? '숫자 필터...' : '텍스트 필터...', icon: 'filter', action: () => customFilterDialog(c, key, numeric) },
    { label: '정규식 필터...', icon: 'find', action: () => customFilterDialog(c, key, numeric, 'regex') },
    ...(numeric ? [
      { label: '상위 10...', action: () => topFilterDialog(c, key) },
      { label: '평균 초과', action: () => applyFilterCriteria(c, { type: 'avg', above: true }, key) },
      { label: '평균 미만', action: () => applyFilterCriteria(c, { type: 'avg', above: false }, key) },
    ] : []),
    { sep: true },
    { node },
  ]);
  menu.style.minWidth = '280px';
  setTimeout(() => search.focus());
}

/** 사용자 지정 자동 필터 (엑셀과 같은 두 조건 + 그리고/또는, 정규식 포함) */
function customFilterDialog(c, key, numeric, op = null) {
  const f = getFilter(key);
  const cur = f?.criteria?.[c]?.type === 'custom' ? f.criteria[c] : {};
  const opts2 = [{ value: '', label: '' }, ...FILTER_OPS.map(([v, l]) => ({ value: v, label: l }))];
  formDialog(`사용자 지정 자동 필터 — ${displayText(f.r1, c) || `${colToName(c)}열`}`, [
    { name: 'op1', label: '조건 1', type: 'select', value: op ?? cur.op1 ?? (numeric ? 'gt' : 'contains'), options: opts2 },
    { name: 'v1', label: '값 1', value: cur.v1 ?? '' },
    { name: 'join', label: '', type: 'select', value: cur.join ?? 'and', options: [{ value: 'and', label: '그리고' }, { value: 'or', label: '또는' }] },
    { name: 'op2', label: '조건 2', type: 'select', value: cur.op2 ?? '', options: opts2 },
    { name: 'v2', label: '값 2', value: cur.v2 ?? '' },
  ], (x) => {
    for (const [o, v] of [[x.op1, x.v1], [x.op2, x.v2]]) {
      if ((o === 'regex' || o === 'notRegex') && v) { try { new RegExp(v, 'u'); } catch (e) { alertDialog('정규식 필터', `정규식이 올바르지 않습니다: ${e.message}`); return false; } }
    }
    applyFilterCriteria(c, { type: 'custom', op1: x.op1, v1: x.v1, join: x.join, op2: x.op2, v2: x.v2 }, key);
    return true;
  }, { note: '? 는 한 글자, * 는 여러 글자를 나타냅니다. 정규식 예) ^브랜드|MO$ , \\d{4}-\\d{2}' });
}

function topFilterDialog(c, key) {
  formDialog('상위 10 자동 필터', [
    { name: 'which', label: '표시', type: 'select', value: 'top', options: [{ value: 'top', label: '상위' }, { value: 'bottom', label: '하위' }] },
    { name: 'n', label: '개수', type: 'number', value: 10 },
    { name: 'unit', label: '', type: 'select', value: 'items', options: [{ value: 'items', label: '항목' }, { value: 'percent', label: '%' }] },
  ], (x) => { applyFilterCriteria(c, { type: 'top', n: Math.max(1, Number(x.n) || 10), bottom: x.which === 'bottom', percent: x.unit === 'percent' }, key); return true; });
}

// ───────────────────────── 차트 ─────────────────────────
function insertChart(type, patch = null) {
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
  const chart = { id, type, title: pieTitle || '차트 제목', range: { r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 }, x, y, w: 480, h: 288, z: nextZ(), ...(patch ?? {}) };
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
  const preview = el('div', { style: { border: '1px solid #ddd', height: '220px', display: 'grid', placeItems: 'center', overflow: 'hidden' } });
  const typeSel = el('select', {}, CHART_TYPES.map((t) => el('option', { value: t.id, selected: t.id === ch.type }, t.label)));
  const titleIn = el('input', { type: 'text', value: ch.title ?? '' });
  const fixedSrc = !!(ch.pivot || ch.series?.length);
  const rangeText0 = ch.range ? `${cellName(ch.range.r1, ch.range.c1)}:${cellName(ch.range.r2, ch.range.c2)}` : '';
  const rangeIn = el('input', { type: 'text', value: rangeText0, disabled: !!ch.pivot });
  const legendSel = el('select', {}, [['b', '아래쪽'], ['t', '위쪽'], ['r', '오른쪽'], ['l', '왼쪽'], ['none', '없음']].map(([v, l]) => el('option', { value: v, selected: (ch.legend ?? 'b') === v }, l)));
  const groupSel = el('select', {}, [['clustered', '묶은 막대'], ['stacked', '누적 막대'], ['percentStacked', '100% 기준 누적']].map(([v, l]) => el('option', { value: v, selected: (ch.grouping ?? 'clustered') === v }, l)));
  const labelsIn = el('input', { type: 'checkbox', checked: !!ch.labels });
  let fmt = (ch.seriesFmt ?? []).map((f) => ({ ...f }));
  const seriesBox = el('div', { class: 'ch-series' });
  const draft = () => {
    const rg = parseRangeName(rangeIn.value.replace(/\$/g, ''));
    const src = fixedSrc && rangeIn.value === rangeText0 ? {} : { range: rg ?? ch.range, series: undefined, pivot: ch.pivot };
    return { ...ch, ...src, type: typeSel.value, title: titleIn.value, legend: legendSel.value, grouping: groupSel.value, labels: labelsIn.checked || undefined, seriesFmt: fmt };
  };
  const renderSeries = (data) => {
    seriesBox.replaceChildren(el('div', { class: 'fc-title' }, '계열 서식 (콤보 차트: 계열마다 종류 · 축 선택)'));
    data.series.forEach((s, i) => {
      const f = fmt[i] ?? (fmt[i] = {});
      const t = el('select', {}, [['', '기본'], ['column', '막대'], ['line', '꺾은선'], ['area', '영역']].map(([v, l]) => el('option', { value: v, selected: (f.type ?? '') === v }, l)));
      const ax = el('select', {}, [['0', '기본 축'], ['1', '보조 축']].map(([v, l]) => el('option', { value: v, selected: String(f.axis ?? s.axis ?? 0) === v }, l)));
      const col = el('input', { type: 'color', value: s.color ?? PALETTE[i % PALETTE.length] });
      const lab = el('input', { type: 'checkbox', checked: !!(f.labels ?? false), title: '데이터 레이블' });
      t.addEventListener('change', () => { f.type = t.value || undefined; draw(false); });
      ax.addEventListener('change', () => { f.axis = Number(ax.value) || undefined; draw(false); });
      col.addEventListener('input', () => { f.color = col.value; draw(false); });
      lab.addEventListener('change', () => { f.labels = lab.checked || undefined; draw(false); });
      seriesBox.append(el('div', { class: 'fc-row', style: { gap: '6px', alignItems: 'center' } }, el('span', { style: { flex: '1', overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap' } }, s.name), t, ax, col, el('label', { class: 'fc-check' }, lab, '레이블')));
    });
  };
  const draw = (withSeries = true) => {
    const d = draft();
    const data = chartModelData(wb, si, d);
    preview.innerHTML = renderChartSvg({ ...d, w: 480, h: 220 }, data);
    if (withSeries) renderSeries(data);
  };
  [typeSel, titleIn, rangeIn, legendSel, groupSel, labelsIn].forEach((i) => i.addEventListener('input', () => draw(i === rangeIn || i === typeSel)));
  draw();
  openDialog({
    title: '차트 편집', width: 560,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
      el('div', { class: 'fc-row', style: { gap: '10px' } }, el('label', {}, el('span', {}, '차트 종류'), typeSel), el('label', {}, el('span', {}, '범례'), legendSel), el('label', {}, el('span', {}, '막대 배치'), groupSel)),
      el('label', {}, el('span', {}, '차트 제목'), titleIn),
      el('label', {}, el('span', {}, ch.pivot ? `데이터: 피벗 테이블 '${ch.pivot.name}' (피벗 차트)` : '데이터 범위'), rangeIn),
      el('label', { class: 'fc-check' }, labelsIn, '모든 계열에 데이터 레이블(값) 표시'),
      seriesBox, preview),
    buttons: [
      {
        label: '확인', primary: true, action: () => {
          const d = draft();
          if (!d.pivot && !d.series?.length && !d.range) { toast('데이터 범위가 올바르지 않습니다.'); return false; }
          const { w, h, ...rest } = d;
          updateChart(id, { ...rest, seriesFmt: fmt.map((f) => Object.fromEntries(Object.entries(f).filter(([, v]) => v !== undefined))) });
          return true;
        },
      },
      { label: '삭제', action: () => deleteChart(id) },
      { label: '취소' },
    ],
  });
}

function chartMenu(id, pos) { objectMenu(id, pos); }

// ───────────────────────── 차트 디자인 · 서식 ─────────────────────────
const chartHere = () => (chartSel ? sheet().charts.find((c) => c.id === chartSel) : null);

/** 차트 삽입 (엑셀 [모든 차트]): 왼쪽 분류 · 위 하위 종류 · 미리 보기. changeId 가 있으면 [차트 종류 변경] */
function insertChartAllDialog(changeId = null) {
  const base = changeId ? sheet().charts.find((c) => c.id === changeId) : null;
  let rg = null;
  if (!base) {
    rg = dataRange();
    if (rg.r2 - rg.r1 > 2000) rg = { ...rg, r2: rg.r1 + 2000 };
  }
  const draftOf = (patch) => (base ? { ...base, grouping: undefined, marker: undefined, scatterStyle: undefined, radarStyle: undefined, explode: undefined, ohlc: undefined, ...patch } : { type: 'column', range: rg, title: '차트 제목', ...patch });
  const dataFor = (d) => chartModelData(wb, si, d);
  const cats = el('div', { class: 'cg-cats' });
  const subs = el('div', { class: 'cg-subs' });
  const prev = el('div', { class: 'cg-prev' });
  const label = el('div', { class: 'cg-label' });
  let pick = null;
  const show = (gi) => {
    [...cats.children].forEach((b, i) => b.classList.toggle('on', i === gi));
    const [, list] = CHART_GALLERY[gi];
    subs.replaceChildren(...list.map(([name, patch], k) => {
      const d = draftOf(patch);
      const thumb = el('button', { class: 'cg-sub', title: name, html: renderChartSvg({ ...d, title: '', legend: 'none', w: 120, h: 80, axisSize: 6 }, dataFor(d)) });
      thumb.addEventListener('click', () => sel(thumb, name, patch));
      if (k === 0) setTimeout(() => sel(thumb, name, patch), 0);
      return thumb;
    }));
  };
  const sel = (thumb, name, patch) => {
    subs.querySelectorAll('.on').forEach((x) => x.classList.remove('on'));
    thumb.classList.add('on');
    pick = patch;
    label.textContent = name;
    const d = draftOf(patch);
    prev.innerHTML = renderChartSvg({ ...d, w: 460, h: 260 }, dataFor(d));
  };
  CHART_GALLERY.forEach(([g], i) => cats.append(el('button', { class: 'cg-cat', onclick: () => show(i) }, g)));
  const cur = base ? CHART_GALLERY.findIndex(([, list]) => list.some(([, p]) => p.type === base.type)) : 0;
  show(Math.max(0, cur));
  openDialog({
    title: base ? '차트 종류 변경' : '차트 삽입', width: 760,
    body: el('div', { class: 'cg-wrap' }, cats, el('div', { class: 'cg-main' }, subs, label, prev)),
    buttons: [{
      label: '확인', primary: true, action: () => {
        if (!pick) return false;
        if (base) { updateChart(base.id, { grouping: undefined, marker: undefined, scatterStyle: undefined, radarStyle: undefined, explode: undefined, ohlc: undefined, ...pick }); gv.renderObjectsAll(); return undefined; }
        insertChart(pick.type, pick);
        return undefined;
      },
    }, { label: '취소' }],
  });
}

/**
 * 위셀 피벗 차트 지표 고르기: 엑셀은 피벗 차트에 모든 값 필드가 강제로 나와서 지표별 차트를 만들려면 피벗 테이블을 따로 만들어야 하지만,
 * 여기서는 같은 피벗 테이블에서 보고 싶은 지표만 골라 차트를 여러 개 만들 수 있음 (슬라이서 · 필터를 누르면 함께 바뀜)
 * onDone({ values, type, secondary }) — values 는 값 필드 표시 이름(고른 순서)
 */
function pivotMetricDialog(entry, cur, onDone, { title = '피벗 차트 — 표시할 지표 선택' } = {}) {
  const d = pivotDefV2(entry.def);
  const names = (d.values ?? []).map((v) => v.name ?? valueName(v));
  if (!names.length) { toast('피벗 테이블에 값 필드가 없습니다.'); return; }
  let order = [...(cur?.values?.length ? cur.values.filter((n) => names.includes(n)) : names), ...names.filter((n) => !(cur?.values ?? names).includes(n))];
  const checked = new Set(cur?.values?.length ? cur.values : names);
  const listEl = el('div', { class: 'pm-list' });
  const type = el('select', {}, [['column', '묶은 세로 막대형'], ['bar', '묶은 가로 막대형'], ['line', '꺾은선형'], ['area', '영역형'], ['combo', '콤보 (막대 + 꺾은선)'], ['pie', '원형 (첫 지표)'], ['doughnut', '도넛형 (첫 지표)']]
    .map(([v, l]) => el('option', { value: v, selected: v === (cur?.type ?? 'column') }, l)));
  const second = el('input', { type: 'checkbox', checked: cur?.secondary ?? true });
  const draw = () => {
    listEl.replaceChildren(...order.map((n, i) => {
      const cb = el('input', { type: 'checkbox', checked: checked.has(n) });
      cb.addEventListener('change', () => { if (cb.checked) checked.add(n); else checked.delete(n); });
      const mv = (dir) => { const j = i + dir; if (j < 0 || j >= order.length) return; [order[i], order[j]] = [order[j], order[i]]; draw(); };
      return el('div', { class: 'pm-row' }, el('label', { class: 'an-check' }, cb, n),
        el('span', { class: 'pm-move' }, el('button', { type: 'button', class: 'lnk', disabled: i === 0, onclick: () => mv(-1) }, '▲'), el('button', { type: 'button', class: 'lnk', disabled: i === order.length - 1, onclick: () => mv(1) }, '▼')));
    }));
  };
  draw();
  openDialog({
    title, width: 460,
    body: el('div', { class: 'an-dlg' },
      el('div', { class: 'muted' }, `'${pivotNameOf(entry)}'의 값 필드 중 차트에 보일 지표만 고르세요. 같은 피벗 테이블로 지표별 차트를 여러 개 만들 수 있고, 슬라이서 · 필터를 누르면 피벗 테이블과 함께 바뀝니다. (엑셀 파일로 저장하면 피벗 테이블 범위를 참조하는 일반 차트가 됩니다)`),
      el('div', { class: 'an-head' }, '지표 (위아래로 순서 바꾸기)'), listEl,
      el('label', { class: 'an-row' }, el('span', {}, '차트 종류'), type),
      el('label', { class: 'an-check' }, second, '콤보: 마지막 지표를 꺾은선 · 보조 축으로 (단위가 다른 지표용, 예: 비용 + ROAS)')),
    buttons: [
      { label: '확인', primary: true, action: () => {
        const values = order.filter((n) => checked.has(n));
        if (!values.length) { toast('지표를 하나 이상 고르세요.'); return false; }
        onDone({ values: values.length === names.length && values.every((n, i) => n === names[i]) ? null : values, type: type.value, secondary: second.checked, count: values.length });
        return undefined;
      } },
      { label: '취소' },
    ],
  });
}
const pivotChartPatch = ({ values, type, secondary, count }, base) => {
  const combo = type === 'combo';
  return {
    type: combo ? 'combo' : type,
    pivot: { ...base, ...(values ? { values } : { values: undefined }) },
    seriesFmt: combo ? Array.from({ length: count }, (_, i) => (i === count - 1 && count > 1 ? { type: 'line', ...(secondary ? { axis: 1 } : {}) } : { type: 'column' })) : undefined,
  };
};

/** 피벗 차트: 피벗 테이블 안이면 그 피벗으로 (값 필드가 둘 이상이면 표시할 지표를 고름), 아니면 새 피벗 테이블 + 피벗 차트 */
function insertPivotChart() {
  const here = pivotHere();
  if (here) {
    const def = here.def;
    if (!def.name) putPivotDef(here, { ...def, name: pivotNameOf(here) });
    if ((pivotDefV2(def).values ?? []).length > 1) {
      pivotMetricDialog(here, null, (pick) => makePivotChart(here, pick));
      return;
    }
    makePivotChart(here, null);
    return;
  }
  toast('먼저 피벗 테이블을 만든 다음, 피벗 테이블 안의 셀을 고르고 [피벗 차트]를 누르세요.');
  pivotDialog();
}
function makePivotChart(here, pick) {
  {
    const def = here.def;
    const a = def.area ?? { r1: def.top ?? 0, c1: def.left ?? 0, r2: (def.top ?? 0) + 10, c2: (def.left ?? 0) + 3 };
    gv.refreshAxes();
    let x = gv.cols.pos(a.c2 + 2);
    let y = gv.rows.pos(a.r1);
    // 피벗이 넓으면 오른쪽 대신 아래에 (화면 밖에 만들지 않게)
    if (x - (gv.sx ?? 0) + 480 > gv.viewW) { x = gv.cols.pos(a.c1); y = gv.rows.pos(a.r2 + 2); }
    const id = `ch${Date.now().toString(36)}${Math.random().toString(36).slice(2, 5)}`;
    const base = { sheet: wb.sheets[here.si].name, name: pivotNameOf(here) };
    const patch = pick ? pivotChartPatch(pick, base) : { type: 'column', pivot: base };
    if (!patch.seriesFmt) delete patch.seriesFmt;
    if (!patch.pivot.values) delete patch.pivot.values;
    const title = pick?.values?.length ? pick.values.join(' · ') : pivotNameOf(here);
    // 같은 피벗으로 여러 차트를 만들면 겹치지 않게 아래로
    const sameCount = sheet().charts.filter((c) => c.pivot?.name === base.name).length;
    const onRight = y === gv.rows.pos(a.r1);
    const chart = { id, title, fieldButtons: true, x: onRight ? x : x + sameCount * 500, y: onRight ? y + sameCount * 300 : y, w: 480, h: 288, z: nextZ(), ...patch };
    wb.transact(() => wb.setSheetProp(si, 'charts', [...sheet().charts.map((c) => ({ ...c })), chart]), meta());
    chartSel = id;
    gv.ensureVisible(gv.rows.indexAt(chart.y + 100), gv.cols.indexAt(chart.x + 200));
    gv.renderObjectsAll();
    updateSelectionUI();
    toast(pick?.values ? `지표 ${pick.values.length}개만 보이는 피벗 차트를 만들었습니다. 슬라이서를 누르면 피벗 테이블과 함께 바뀝니다.` : '피벗 차트를 만들었습니다. 피벗 테이블 필드를 바꾸거나 슬라이서를 누르면 차트도 함께 바뀝니다.');
  }
}
/** 피벗 차트의 지표 다시 고르기 ([데이터 선택] · 오른쪽 클릭) */
function pivotChartMetrics(chId = chartSel) {
  const ch = sheet().charts.find((c) => c.id === chId);
  const e = ch?.pivot ? findPivotEntry(ch.pivot.sheet ?? null, ch.pivot.name ?? null) : null;
  if (!e) { toast('연결된 피벗 테이블을 찾을 수 없습니다.'); return; }
  pivotMetricDialog(e, { values: ch.pivot.values, type: ch.type, secondary: ch.seriesFmt?.some((f) => f?.axis === 1) ?? true }, (pick) => {
    const patch = pivotChartPatch(pick, { sheet: ch.pivot.sheet, name: ch.pivot.name });
    updateChart(ch.id, { ...patch, ...(pick.type === 'combo' ? {} : { seriesFmt: undefined }) });
    gv.renderObjectsAll();
  }, { title: '피벗 차트 — 표시할 지표' });
}

function chartElementsMenu() {
  const ch = chartHere();
  if (!ch) return [];
  const up = (patch) => { updateChart(ch.id, patch); gv.renderObjectsAll(); };
  return [
    { title: '축' },
    { label: '기본 가로 축', checked: !ch.axes?.x?.hide, action: () => up({ axes: { ...ch.axes, x: { ...ch.axes?.x, hide: !ch.axes?.x?.hide || undefined } } }) },
    { label: '기본 세로 축', checked: !ch.axes?.y?.hide, action: () => up({ axes: { ...ch.axes, y: { ...ch.axes?.y, hide: !ch.axes?.y?.hide || undefined } } }) },
    { title: '축 제목' },
    { label: '가로 축 제목...', action: () => formDialog('가로 축 제목', [{ name: 't', label: '제목', value: ch.axes?.x?.title ?? '' }], ({ t }) => up({ axes: { ...ch.axes, x: { ...ch.axes?.x, title: t || undefined } } })) },
    { label: '세로 축 제목...', action: () => formDialog('세로 축 제목', [{ name: 't', label: '제목', value: ch.axes?.y?.title ?? '' }], ({ t }) => up({ axes: { ...ch.axes, y: { ...ch.axes?.y, title: t || undefined } } })) },
    { title: '차트 제목' },
    { label: '없음', checked: !ch.title, action: () => up({ title: '' }) },
    { label: '차트 위', checked: !!ch.title, action: () => up({ title: ch.title || '차트 제목' }) },
    { title: '데이터 레이블' },
    { label: '없음', checked: !ch.labels, action: () => up({ labels: undefined }) },
    { label: '표시', checked: !!ch.labels, action: () => up({ labels: true }) },
    { title: '눈금선' },
    { label: '기본 주 가로', checked: ch.gridY !== false, action: () => up({ gridY: ch.gridY === false ? undefined : false }) },
    { label: '기본 주 세로', checked: !!ch.gridX, action: () => up({ gridX: !ch.gridX || undefined }) },
    { title: '범례' },
    ...[['none', '없음'], ['r', '오른쪽'], ['t', '위쪽'], ['l', '왼쪽'], ['b', '아래쪽']].map(([v, l]) => ({ label: l, checked: (ch.legend ?? 'b') === v, action: () => up({ legend: v }) })),
  ];
}

function chartLayoutsMenu() {
  const ch = chartHere();
  if (!ch) return [];
  const L = [
    ['레이아웃 1 · 제목 + 범례 오른쪽', { legend: 'r', labels: undefined, gridY: undefined, axes: { ...ch.axes, x: { ...ch.axes?.x, title: undefined }, y: { ...ch.axes?.y, title: undefined } } }],
    ['레이아웃 2 · 레이블 + 범례 위', { legend: 't', labels: true, gridY: false }],
    ['레이아웃 3 · 범례 아래', { legend: 'b', labels: undefined, gridY: undefined }],
    ['레이아웃 4 · 레이블만', { legend: 'none', labels: true, gridY: undefined }],
    ['레이아웃 5 · 축 제목', { legend: 'b', axes: { ...ch.axes, x: { ...ch.axes?.x, title: '항목' }, y: { ...ch.axes?.y, title: '값' } } }],
    ['레이아웃 6 · 깔끔하게 (눈금선 · 범례 없음)', { legend: 'none', gridY: false, labels: true }],
  ];
  return L.map(([n, p]) => ({ label: n, action: () => { updateChart(ch.id, p); gv.renderObjectsAll(); } }));
}

function chartColorsMenu(a) {
  const ch = chartHere();
  if (!ch) return undefined;
  openMenu(a, Object.entries(CHART_PALETTES).map(([k, p]) => ({
    label: p.label, checked: (ch.palette ?? 'office') === k,
    icon: `<span style="display:inline-flex">${p.colors.slice(0, 6).map((c) => `<i style="display:block;width:6px;height:12px;background:${c}"></i>`).join('')}</span>`,
    action: () => { updateChart(ch.id, { palette: k === 'office' ? undefined : k, seriesFmt: (ch.seriesFmt ?? []).map(({ color, ...f }) => f) }); gv.renderObjectsAll(); },
  })));
  return undefined;
}

// 차트 스타일: 엑셀 스타일 갤러리처럼 배경 · 눈금선 · 글꼴 조합 (WIXEL 모던 포함)
const CHART_STYLES = [
  ['스타일 1 (기본)', { fill: undefined, plotFill: undefined, border: undefined, gridY: undefined, textColor: undefined, gridColor: undefined, titleBold: undefined, rounded: undefined }],
  ['스타일 2 (레이블 강조)', { labels: true, gridY: false, titleBold: true }],
  ['스타일 3 (연한 배경)', { plotFill: '#f5f7fb', gridColor: '#ffffff', border: '#d9d9d9' }],
  ['스타일 4 (어두운 배경)', { fill: '#1f2937', plotFill: '#111827', textColor: '#e5e7eb', gridColor: '#374151', titleColor: '#ffffff' }],
  ['스타일 5 (테두리)', { border: '#595959', gridY: undefined }],
  ['WIXEL 카드', { fill: '#ffffff', border: '#e2e8f0', rounded: true, titleBold: true, titleColor: '#0f172a', textColor: '#64748b', gridColor: '#eef2f7', palette: 'modern' }],
  ['WIXEL 대시보드 다크', { fill: '#0f172a', plotFill: '#0f172a', border: '#1e293b', rounded: true, titleColor: '#f8fafc', textColor: '#94a3b8', gridColor: '#1e293b', palette: 'vivid', titleBold: true }],
  ['WIXEL 미니멀', { fill: '#ffffff', gridY: false, textColor: '#475569', titleColor: '#0f172a', palette: 'slate', border: undefined }],
];
function chartStylesMenu(a) {
  const ch = chartHere();
  if (!ch) return undefined;
  const data = chartModelData(wb, si, ch);
  const grid = el('div', { class: 'cs-grid' }, CHART_STYLES.map(([n, p]) => el('button', {
    class: 'cs-chip', title: n, html: renderChartSvg({ ...ch, ...p, title: '', legend: 'none', w: 150, h: 96, axisSize: 6 }, data),
    onmousedown: (e) => e.preventDefault(), onclick: () => { closeMenus(); updateChart(ch.id, p); gv.renderObjectsAll(); },
  })));
  openMenu(a, [{ title: '차트 스타일' }, { node: grid }]);
  return undefined;
}

/** 행/열 전환: 범위 차트는 계열 방향을 바꿈 */
function chartSwitchRowCol() {
  const ch = chartHere();
  if (!ch) return;
  if (!ch.range) { toast('이 차트는 행/열 전환을 할 수 없습니다 (피벗 차트나 계열을 직접 지정한 차트).'); return; }
  updateChart(ch.id, { byRows: !ch.byRows || undefined });
  gv.renderObjectsAll();
}

/** 차트 서식 창 (엑셀의 [차트 영역 서식] 작업 창): 영역 · 제목 · 축 · 계열 · 레이블 */
let chartPaneDlg = null;
/** 데이터 요소(항목 하나)의 색 (엑셀: 요소 하나를 골라 채우기) */
function pointColorRow(s, f, setF) {
  const cats = s.categories ?? null;
  const idx = el('select', {}, (s.values ?? []).map((_, k) => el('option', { value: String(k) }, `${k + 1}. ${String(cats?.[k] ?? '').slice(0, 16)}`)));
  const inp = el('input', { type: 'color', value: '#ed7d31' });
  const apply = () => setF({ pointColors: { ...(f.pointColors ?? {}), [idx.value]: inp.value } });
  inp.addEventListener('input', apply);
  return el('label', { class: 'cfp-row' }, el('span', {}, '요소 색'), el('span', { class: 'cfp-color' }, idx, inp,
    el('button', { class: 'btn small', onclick: () => { const pc = { ...(f.pointColors ?? {}) }; delete pc[idx.value]; setF({ pointColors: Object.keys(pc).length ? pc : undefined }); } }, '되돌리기')));
}
function chartFormatPane(id = chartSel) {
  const ch0 = sheet().charts.find((c) => c.id === id);
  if (!ch0) { toast('차트를 선택하세요.'); return; }
  if (chartPaneDlg) chartPaneDlg.close();
  const body = el('div', { class: 'cfp' });
  const get = () => sheet().charts.find((c) => c.id === id) ?? ch0;
  const up = (patch) => { updateChart(id, patch); gv.renderObjectsAll(); };
  const sec = (title, ...rows) => el('details', { class: 'cfp-sec', open: true }, el('summary', {}, title), ...rows);
  const row = (label, input) => el('label', { class: 'cfp-row' }, el('span', {}, label), input);
  const color = (v, fn, allowNone = true) => {
    const inp = el('input', { type: 'color', value: v ?? '#ffffff' });
    inp.addEventListener('input', () => fn(inp.value));
    return el('span', { class: 'cfp-color' }, inp, allowNone ? el('button', { class: 'btn small', onclick: () => fn(undefined) }, '없음') : null);
  };
  const num = (v, fn, attrs = {}) => { const i = el('input', { type: 'number', value: v ?? '', ...attrs }); i.addEventListener('change', () => fn(i.value === '' ? undefined : Number(i.value))); return i; };
  const txt = (v, fn) => { const i = el('input', { type: 'text', value: v ?? '' }); i.addEventListener('change', () => fn(i.value)); return i; };
  const chk = (v, fn) => { const i = el('input', { type: 'checkbox', checked: !!v }); i.addEventListener('change', () => fn(i.checked)); return i; };
  const sel2 = (v, opts, fn) => { const s = el('select', {}, opts.map(([k, l]) => el('option', { value: k, selected: String(v ?? '') === String(k) }, l))); s.addEventListener('change', () => fn(s.value)); return s; };
  const draw = () => {
    const ch = get();
    const axis = (k) => ch.axes?.[k] ?? {};
    const setAx = (k, patch) => up({ axes: { ...ch.axes, [k]: { ...axis(k), ...patch } } });
    const data = chartModelData(wb, si, ch);
    body.replaceChildren(
      sec('차트 영역',
        row('너비(px)', num(Math.round(ch.w), (v) => v && up({ w: clamp(v, 120, 4000) }), { min: 120, max: 4000, step: 1 })),
        row('높이(px)', num(Math.round(ch.h), (v) => v && up({ h: clamp(v, 90, 4000) }), { min: 90, max: 4000, step: 1 })),
        row('채우기', color(ch.fill, (v) => up({ fill: v }))),
        row('테두리', color(ch.border, (v) => up({ border: v }))),
        row('둥근 모서리', chk(ch.rounded, (v) => up({ rounded: v || undefined }))),
        row('글자 색', color(ch.textColor, (v) => up({ textColor: v }))),
        row('그림 영역 채우기', color(ch.plotFill, (v) => up({ plotFill: v }))),
        row('색 구성', sel2(ch.palette ?? 'office', Object.entries(CHART_PALETTES).map(([k, p]) => [k, p.label]), (v) => up({ palette: v === 'office' ? undefined : v })))),
      sec('차트 제목',
        row('제목', txt(ch.title, (v) => up({ title: v }))),
        row('글꼴 크기(pt)', num(ch.titleSize, (v) => up({ titleSize: v }), { min: 6, max: 40 })),
        row('굵게', chk(ch.titleBold, (v) => up({ titleBold: v || undefined }))),
        row('색', color(ch.titleColor, (v) => up({ titleColor: v })))),
      sec('세로(값) 축',
        row('표시', chk(!axis('y').hide, (v) => setAx('y', { hide: !v || undefined }))),
        row('최소값', num(axis('y').min, (v) => setAx('y', { min: v }))),
        row('최대값', num(axis('y').max, (v) => setAx('y', { max: v }))),
        row('주 단위', num(axis('y').major, (v) => setAx('y', { major: v }), { min: 0 })),
        row('값을 거꾸로', chk(axis('y').reverse, (v) => setAx('y', { reverse: v || undefined }))),
        row('표시 형식', txt(axis('y').numFmt, (v) => setAx('y', { numFmt: v || undefined }))),
        row('제목', txt(axis('y').title, (v) => setAx('y', { title: v || undefined }))),
        row('주 눈금선', chk(ch.gridY !== false, (v) => up({ gridY: v ? undefined : false }))),
        row('눈금선 색', color(ch.gridColor, (v) => up({ gridColor: v }))),
        row('글꼴 크기(pt)', num(ch.axisSize, (v) => up({ axisSize: v }), { min: 6, max: 24 }))),
      sec('보조 세로 축',
        row('최소값', num(axis('y2').min, (v) => setAx('y2', { min: v }))),
        row('최대값', num(axis('y2').max, (v) => setAx('y2', { max: v }))),
        row('표시 형식', txt(axis('y2').numFmt, (v) => setAx('y2', { numFmt: v || undefined }))),
        row('제목', txt(axis('y2').title, (v) => setAx('y2', { title: v || undefined })))),
      sec('가로(항목) 축',
        row('표시', chk(!axis('x').hide, (v) => setAx('x', { hide: !v || undefined }))),
        row('제목', txt(axis('x').title, (v) => setAx('x', { title: v || undefined }))),
        row('세로 눈금선', chk(ch.gridX, (v) => up({ gridX: v || undefined })))),
      sec('범례 · 레이블',
        row('범례 위치', sel2(ch.legend ?? 'b', [['b', '아래쪽'], ['t', '위쪽'], ['r', '오른쪽'], ['l', '왼쪽'], ['none', '없음']], (v) => up({ legend: v }))),
        row('범례 글꼴(pt)', num(ch.legendSize, (v) => up({ legendSize: v }), { min: 6, max: 24 })),
        row('데이터 레이블', chk(ch.labels, (v) => up({ labels: v || undefined }))),
        row('데이터 표', chk(ch.dataTable, (v) => up({ dataTable: v || undefined })))),
      sec('계열 옵션',
        row('간격 너비(%)', num(ch.gap, (v) => up({ gap: v }), { min: 0, max: 500 })),
        row('계열 겹치기(%)', num(ch.overlap, (v) => up({ overlap: v }), { min: -100, max: 100 })),
        row('요소마다 다른 색', chk(ch.varyColors, (v) => up({ varyColors: v || undefined }))),
        row('배치', sel2(ch.grouping ?? 'clustered', [['clustered', '묶은'], ['stacked', '누적'], ['percentStacked', '100% 기준 누적']], (v) => up({ grouping: v === 'clustered' ? undefined : v }))),
        ch.type === 'doughnut' ? row('도넛 구멍 크기(%)', num(ch.hole ?? 50, (v) => up({ hole: v }), { min: 10, max: 90 })) : null,
        ch.type === 'pie' || ch.type === 'doughnut' ? row('첫째 조각 각(°)', num(ch.firstAngle ?? 0, (v) => up({ firstAngle: v }), { min: 0, max: 360 })) : null,
        ch.type === 'pie' ? row('쪼개기(%)', num(ch.explode ?? 0, (v) => up({ explode: v || undefined }), { min: 0, max: 40 })) : null,
        ch.type === 'histogram' ? row('구간 수', num(ch.binCount, (v) => up({ binCount: v }), { min: 1, max: 100 })) : null,
        ch.type === 'histogram' ? row('구간 너비', num(ch.binWidth, (v) => up({ binWidth: v }), { min: 0 })) : null,
        ...data.series.map((s, i) => {
          const f = (ch.seriesFmt ?? [])[i] ?? {};
          const setF = (patch) => { const list = [...(ch.seriesFmt ?? [])]; while (list.length <= i) list.push({}); list[i] = { ...list[i], ...patch }; up({ seriesFmt: list }); };
          return el('div', { class: 'cfp-series' },
            el('b', {}, s.name || `계열${i + 1}`),
            row('색', color(s.color ?? paletteOf(ch)[i % paletteOf(ch).length], (v) => setF({ color: v }), false)),
            row('종류', sel2(f.type ?? '', [['', '기본'], ['column', '막대'], ['line', '꺾은선'], ['area', '영역']], (v) => setF({ type: v || undefined }))),
            row('축', sel2(f.axis ?? 0, [[0, '기본 축'], [1, '보조 축']], (v) => setF({ axis: Number(v) || undefined }))),
            row('선 굵기(px)', num(f.lineWidth, (v) => setF({ lineWidth: v }), { min: 0.5, max: 12, step: 0.25 })),
            row('표식', sel2(f.marker ?? '', [['', '자동'], ['none', '없음'], ['circle', '원'], ['square', '사각형'], ['diamond', '마름모'], ['triangle', '삼각형']], (v) => setF({ marker: v || undefined }))),
            row('부드러운 선', chk(f.smooth, (v) => setF({ smooth: v || undefined }))),
            row('레이블', chk(f.labels, (v) => setF({ labels: v || undefined }))),
            row('레이블 형식', txt(f.numFmt, (v) => setF({ numFmt: v || undefined }))),
            row('레이블 위치', sel2(f.labelPos ?? '', [['', '자동'], ['outEnd', '바깥쪽 끝에'], ['insideEnd', '안쪽 끝에'], ['center', '가운데'], ['insideBase', '축 쪽에']], (v) => setF({ labelPos: v || undefined }))),
            row('선 종류', sel2(f.dash ?? '', [['', '실선'], ['dash', '파선'], ['dot', '점선'], ['dashDot', '일점 쇄선'], ['longDash', '긴 파선']], (v) => setF({ dash: v || undefined }))),
            row('표식 크기', num(f.markerSize, (v) => setF({ markerSize: v }), { min: 2, max: 30 })),
            row('테두리 색', color(f.outline, (v) => setF({ outline: v }))),
            row('추세선', sel2(f.trend ?? '', [['', '없음'], ['linear', '선형'], ['exp', '지수'], ['movingAvg', '이동 평균']], (v) => setF({ trend: v || undefined }))),
            f.trend === 'movingAvg' ? row('이동 평균 구간', num(f.trendPeriod ?? 3, (v) => setF({ trendPeriod: v }), { min: 2, max: 50 })) : null,
            f.trend && f.trend !== 'movingAvg' ? row('앞으로 예측(구간)', num(f.trendForward, (v) => setF({ trendForward: v || undefined }), { min: 0, max: 100 })) : null,
            pointColorRow({ ...s, categories: data.categories }, f, setF));
        })),
    );
  };
  draw();
  chartPaneDlg = openDialog({ title: '차트 서식', width: 340, modeless: true, body, onClose: () => { chartPaneDlg = null; } });
  chartPaneDlg.root.classList.add('pane-dlg');
}


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
    title: 'WIXEL', body: '표를 정상 범위로 변환하시겠습니까?',
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
  const groups = TABLE_STYLE_GROUPS.flatMap((g) => [{ title: g }, { node: el('div', { class: 'style-grid tstyles' }, TABLE_STYLES.filter((s) => s.group === g).map(chip)) }]);
  const t = tableHere();
  openMenu(anchorEl, [
    ...groups,
    ...(t ? [{ sep: true }, { label: '지우기', action: () => updateTable(t.id, { style: 'None' }) }] : []),
  ], { scroll: true });
}

// ───────────────────────── 슬라이서 ─────────────────────────
const SLICER_COLORS = [['blue', '파랑'], ['orange', '주황'], ['gray', '회색'], ['gold', '금색'], ['sky', '하늘색'], ['green', '녹색']];

const koCollator = new Intl.Collator('ko');
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

/**
 * 슬라이서 항목 + 슬라이서 설정 (엑셀 [슬라이서 설정]): 정렬(오름차순/내림차순), 사용자 지정 목록 순서,
 * 데이터 없는 항목 숨기기 / 시각적으로 표시 / 마지막에 표시
 */
function slicerModel(sl) {
  const m = slicerModelRaw(sl);
  if (!m.items?.length) return m;
  let items = m.items;
  // 사용자 지정 목록(요일 · 월 · 분기 …): 모든 항목이 한 목록에 들어 있으면 그 순서
  if (sl.customList !== false) {
    const list = CUSTOM_LISTS.find((L) => items.every((it) => it.key === '' || L.includes(it.text)));
    if (list) items = [...items].sort((a, b) => (a.key === '' ? 1 : b.key === '' ? -1 : list.indexOf(a.text) - list.indexOf(b.text)));
  }
  if (sl.sort === 'desc') {
    const blank = items.filter((it) => it.key === '');
    items = [...items.filter((it) => it.key !== '').reverse(), ...blank];
  }
  if (sl.hideNoData) items = items.filter((it) => it.hasData || it.selected && m.filtered);
  else if (sl.noDataLast !== false) items = [...items.filter((it) => it.hasData), ...items.filter((it) => !it.hasData)];
  if (sl.markNoData === false) items = items.map((it) => ({ ...it, hasData: true }));
  return { ...m, items };
}

/** 슬라이서가 가리키는 대상과 항목 → { caption, items: [{ key, text, selected, hasData }], filtered, broken?, apply(values) } */
function slicerModelRaw(sl) {
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
    const items = sortItems([...vals.values()]).map((e) => ({ key: e.key, v: e.v, text: e.key === '' ? '(비어 있음)' : e.key, selected: !sel || sel.has(e.key), hasData: e.hasData }));
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
    const targets = slicerPivotTargets(src);
    if (!targets.length) return { items: [], broken: '연결된 피벗 테이블을 찾을 수 없습니다.' };
    const def = targets[0].def;
    const rows = pivotSource(def);
    if (!rows) return { items: [], broken: '피벗 테이블 원본을 찾을 수 없습니다.' };
    const header = headerNames(rows);
    const fi = header.findIndex((h) => h.toLowerCase() === String(src.field).toLowerCase());
    if (fi < 0) return { items: [], broken: `'${src.field}' 필드를 찾을 수 없습니다.` };
    const filters = def.filters ?? {};
    const own = Object.keys(filters).find((k) => k.toLowerCase() === header[fi].toLowerCase());
    const sel = own ? new Set(filters[own]) : null;
    // 항목 목록은 원본과 다른 필터가 같으면 다시 계산하지 않음 (슬라이서를 그릴 때마다 원본 전체를 도는 것 방지)
    const memoKey = `${fi}\u0001${JSON.stringify(Object.entries(filters).filter(([k]) => k !== own))}`;
    let memo = slicerMemo.get(rows.cube);
    if (!memo) { memo = new Map(); slicerMemo.set(rows.cube, memo); }
    let items = memo.get(memoKey);
    if (!items) {
      // 열 기반 엔진: 항목 코드별로 다른 필터를 통과한 행이 있는지 한 번에 (수백만 행도 한 번 훑기)
      const cube = rows.cube;
      const lower = cube.header.map((h) => h.toLowerCase());
      const others = Object.entries(filters).filter(([k]) => k !== own).map(([k, v]) => [lower.indexOf(k.toLowerCase()), new Set(v)]).filter(([i]) => i >= 0);
      const st = itemStats(cube, fi, others);
      items = st.keys.map((k, c) => ({ key: st.texts[c], v: k, hasData: !!st.has[c] }))
        .sort((x, y) => {
          if ((x.v === PIVOT_EMPTY) !== (y.v === PIVOT_EMPTY)) return x.v === PIVOT_EMPTY ? 1 : -1;
          if (typeof x.v === 'number' && typeof y.v === 'number') return x.v - y.v;
          if (typeof x.v === 'number') return -1;
          if (typeof y.v === 'number') return 1;
          return koCollator.compare(x.key, y.key);
        });
      if (memo.size > 200) memo.clear();
      memo.set(memoKey, items);
    }
    // 숫자 항목은 원본 열의 표시 형식으로 (날짜 46204 → 2026-07-01)
    const sd = rows;
    const colStyle = sd?.ref ? wb.styleAt(sd.si, Math.min(sd.ref.r1 + 1, sd.ref.r2), sd.ref.c1 + fi) : null;
    const shown = (e) => (typeof e.v === 'number' && colStyle?.numFmt && colStyle.numFmt !== 'general' ? formatValue(e.v, colStyle).text : e.key);
    return {
      items: items.map((e) => ({ key: e.key, v: e.v, text: shown(e), selected: !sel || sel.has(e.key), hasData: e.hasData })),
      filtered: !!sel,
      targets,
      apply: (values) => {
        // 이 슬라이서에 연결된 모든 피벗 테이블(다른 시트 포함)에 같은 필터
        wb.transact(() => {
          for (const tg of slicerPivotTargets(src)) {
            const f0 = tg.def.filters ?? {};
            const nf = { ...f0 };
            const k0 = Object.keys(f0).find((k) => k.toLowerCase() === header[fi].toLowerCase());
            if (k0) delete nf[k0];
            if (values) nf[header[fi]] = values;
            putPivotDef(tg, { ...tg.def, filters: nf });
          }
        }, meta());
      },
    };
  }
  return { items: [], broken: '연결 대상이 없습니다.' };
}

// 원본 행 배열(피벗 원본 캐시) → 슬라이서 항목 목록. 원본이 바뀌면 배열이 새로 만들어져 자동으로 버려짐
const slicerMemo = new WeakMap();

/** 피벗 슬라이서가 연결된 피벗 목록 (source.pivots = [{ sheet, name }], 옛 형식 self/sheet 도 읽음) */
function slicerPivotTargets(src, hostSi = si) {
  if (src.pivots?.length) return src.pivots.map((p) => findPivotEntry(p.sheet ?? null, p.name ?? null, hostSi)).filter(Boolean);
  const s = src.self || !src.sheet ? hostSi : wb.sheetIndexByName(src.sheet);
  const e = s >= 0 ? pivotDefs(s)[0] : null;
  return e ? [e] : [];
}

// ───── 시간 표시 막대 (엑셀 Timeline) ─────
let tlDrag = null;
function markTimelineDrag() {
  const lo = Math.min(tlDrag.a, tlDrag.b);
  const hi = Math.max(tlDrag.a, tlDrag.b);
  for (const c of tlDrag.root.querySelectorAll('.tl-cell')) { const i = Number(c.dataset.i); c.classList.toggle('drag', i >= lo && i <= hi); }
}
/** 기간 칸 a~b 를 골라 연결된 피벗 · 표를 거름 (전부면 필터 해제, 이미 그 기간만이면 해제) */
function timelineApply(id, a, b) {
  const sl = (sheet().slicers ?? []).find((x) => x.id === id);
  if (!sl) return;
  const m = slicerModel(sl);
  if (m.broken) return;
  const periods = timelinePeriods(m.items, sl.level ?? 'M');
  const lo = Math.min(a, b);
  const hi = Math.max(a, b);
  const keys = new Set(periods.slice(lo, hi + 1).flatMap((p) => p.items.map((it) => it.key)));
  const all = m.items.map((it) => it.key);
  const curSel = m.items.filter((it) => it.selected).map((it) => it.key);
  const same = m.filtered && curSel.length === keys.size && curSel.every((k) => keys.has(k));
  sl._anchor = lo;
  if (!keys.size) { toast('그 기간에는 데이터가 없습니다.'); gv.renderObjectsAll(); return; }
  m.apply(same || keys.size === all.length ? null : all.filter((k) => keys.has(k)));
  gv.layout();
  gv.renderObjectsAll();
  setMode();
}
/** 삽입 → 시간 표시 막대: 날짜 필드만 고름 */
function insertTimelineDialog() {
  if (editing && !commitEdit()) return;
  const t = tableHere();
  const pe = t ? null : pivotHere() ?? pivotDefs()[0] ?? null;
  let fields = [];
  let makeSource;
  let anchor;
  const isDateStyle = (st) => /date/.test(st?.numFmt ?? '') || (st?.numFmt === 'custom' && /[yd]/i.test(st.code ?? '') && !/[#0]/.test(st.code ?? ''));
  if (t) {
    const names = columnNames(wb, si, t);
    fields = names.filter((_, i) => isDateStyle(wb.styleAt(si, dataTop(t), t.c1 + i)) && typeof wb.getValue(si, dataTop(t), t.c1 + i) === 'number');
    makeSource = (name) => ({ kind: 'table', table: t.name, column: name });
    anchor = gv.sheetRect({ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
  } else if (pe) {
    const src = pivotSource(pe.def);
    if (!src) { alertDialog('시간 표시 막대', '피벗 테이블 원본을 찾을 수 없습니다.'); return; }
    const hdr = headerNames(src);
    fields = hdr.filter((_, i) => src.ref && isDateStyle(wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + i)));
    if (!pe.def.name) setPivotDef(pe, { ...pe.def, name: pivotNameOf(pe) });
    makeSource = (name) => ({ kind: 'pivot', field: name, pivots: [{ sheet: sheet().name, name: pivotNameOf(pe) }] });
    anchor = gv.sheetRect(pe.def.area ?? { r1: 0, c1: 0, r2: 0, c2: 0 });
  } else {
    alertDialog('시간 표시 막대', '시간 표시 막대는 날짜 열이 있는 표나 피벗 테이블에 넣을 수 있습니다.');
    return;
  }
  if (!fields.length) { alertDialog('시간 표시 막대', '날짜 형식의 필드가 없습니다. 원본 열에 날짜 표시 형식을 적용하세요.'); return; }
  formDialog('시간 표시 막대 삽입', [
    { name: 'f', label: '날짜 필드', type: 'select', value: fields[0], options: fields.map((f) => ({ value: f, label: f })) },
    { name: 'lv', label: '시간 수준', type: 'select', value: 'M', options: [{ value: 'Y', label: '연도' }, { value: 'Q', label: '분기' }, { value: 'M', label: '월' }, { value: 'D', label: '일' }] },
  ], ({ f, lv }) => {
    const tl = { id: newObjId('sl'), caption: f, source: makeSource(f), timeline: true, level: lv, columns: 1, color: 'blue', multi: false, style: 'SlicerStyleLight1', x: Math.round(anchor.x), y: Math.round(anchor.y + anchor.h + 16), w: 520, h: 130, z: nextZ() };
    wb.transact(() => {
      wb.setSheetProp(si, 'slicers', [...(sheet().slicers ?? []).map((x) => ({ ...x })), tl]);
      if (t && !t.filter) setTables((l) => l.map((x) => (x.id === t.id ? { ...x, filter: { criteria: {}, hidden: {} } } : x)));
    }, meta());
    chartSel = tl.id;
    gv.ensureVisible(gv.rows.indexAt(tl.y + 40), gv.cols.indexAt(tl.x + 100));
    gv.renderObjectsAll();
    updateSelectionUI();
    return true;
  }, { note: '칸을 누르거나 끌어서 기간을 고르면 연결된 피벗 테이블 · 표가 그 기간으로 걸러집니다. Shift+클릭으로 기간을 늘릴 수 있습니다.' });
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
  const pe = t ? null : pivotHere() ?? pivotDefs()[0] ?? null;
  let fields;
  let makeSource;
  let anchor;
  let nameFix = null;
  if (t) {
    fields = columnNames(wb, si, t);
    makeSource = (name) => ({ kind: 'table', table: t.name, column: name });
    anchor = gv.sheetRect({ r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 });
  } else if (pe) {
    const rows = pivotSource(pe.def);
    if (!rows) { alertDialog('슬라이서 삽입', '피벗 테이블 원본을 찾을 수 없습니다.'); return; }
    fields = headerNames(rows);
    const pname = pe.def.name ?? nextPivotName();
    if (!pe.def.name) nameFix = pname;
    makeSource = (name) => ({ kind: 'pivot', field: name, pivots: [{ sheet: sheet().name, name: pname }] });
    const a = pe.def.area ?? { r1: 0, c1: 0, r2: 0, c2: 0 };
    anchor = gv.sheetRect(a);
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
          if (nameFix) { setPivotDef(pe, { ...pe.def, name: nameFix }); nameFix = null; }
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
  const src = sl.source?.kind === 'table' ? `원본 이름: ${sl.source.table}[${sl.source.column}]`
    : `원본 이름: ${sl.source?.field} (피벗 테이블 ${slicerPivotTargets(sl.source ?? {}).map((e) => `'${pivotNameOf(e)}'`).join(', ')})`;
  const inp = (v, attrs = {}) => el('input', { type: 'text', value: v ?? '', ...attrs });
  const chk = (on, label) => { const c = el('input', { type: 'checkbox', checked: !!on }); return [c, el('label', { class: 'fc-check' }, c, label)]; };
  const radio = (name, v, on, label) => { const r = el('input', { type: 'radio', name, value: v, checked: !!on }); return [r, el('label', { class: 'fc-check' }, r, label)]; };
  const name = inp(sl.name ?? sl.caption);
  const cap = inp(sl.caption);
  const [hdr, hdrL] = chk(sl.showHeader !== false, '머리글 표시');
  const [asc, ascL] = radio('slsort', 'asc', sl.sort !== 'desc', '오름차순 (A-Z, ㄱ-ㅎ)');
  const [desc, descL] = radio('slsort', 'desc', sl.sort === 'desc', '내림차순 (Z-A, ㅎ-ㄱ)');
  const [cl, clL] = chk(sl.customList !== false, '정렬할 때 사용자 지정 목록 사용');
  const [hide, hideL] = chk(!!sl.hideNoData, '데이터가 없는 항목 숨기기');
  const [mark, markL] = chk(sl.markNoData !== false, '데이터가 없는 항목을 시각적으로 표시');
  const [last, lastL] = chk(sl.noDataLast !== false, '데이터가 없는 항목을 마지막에 표시');
  const [del, delL] = chk(!!sl.showDeleted, '데이터 원본에서 삭제된 항목 표시');
  const sync = () => { mark.disabled = hide.checked; last.disabled = hide.checked; };
  hide.addEventListener('change', sync);
  sync();
  openDialog({
    title: '슬라이서 설정', width: 520,
    body: el('div', { class: 'sl-settings' },
      el('div', { class: 'muted' }, src),
      el('label', {}, el('span', {}, '이름'), name),
      el('div', { class: 'menu-title', style: { padding: '6px 0 2px' } }, '머리글'),
      hdrL, el('label', {}, el('span', {}, '캡션'), cap),
      el('div', { class: 'sl-set-cols' },
        el('div', {}, el('div', { class: 'menu-title', style: { padding: '6px 0 2px' } }, '항목 정렬 및 필터링'), ascL, descL, clL),
        el('div', {}, el('div', { class: 'menu-title', style: { padding: '6px 0 2px' } }, '데이터가 없는 항목'), hideL, markL, lastL, delL))),
    buttons: [
      {
        label: '확인', primary: true, action: () => {
          updateObject(id, {
            name: name.value.trim() || undefined, caption: cap.value, showHeader: hdr.checked ? undefined : false,
            sort: desc.checked ? 'desc' : undefined, customList: cl.checked ? undefined : false,
            hideNoData: hide.checked || undefined, markNoData: mark.checked ? undefined : false, noDataLast: last.checked ? undefined : false,
            showDeleted: del.checked || undefined,
          });
          gv.renderObjectsAll();
        },
      },
      { label: '취소' },
    ],
  });
}

/** 슬라이서 스타일 갤러리 (+ 사용자 지정 색) */
function slicerStyleGallery(anchorEl) {
  const sl = (sheet().slicers ?? []).find((x) => x.id === chartSel);
  if (!sl) { toast('슬라이서를 선택하세요.'); return; }
  const chip = (s) => el('button', {
    class: 'style-chip slstyle', title: s.label, onmousedown: (e) => e.preventDefault(),
    style: { background: s.colors.frame, borderColor: s.colors.border },
    onclick: () => { closeMenus(); updateObject(sl.id, { style: s.name, color: undefined, custom: undefined }); gv.renderObjectsAll(); },
  }, el('i', { style: { background: s.colors.selFill, borderColor: s.colors.selBorder } }), el('i', { style: { background: s.colors.item, borderColor: s.colors.itemBorder } }));
  const groups = SLICER_STYLE_GROUPS.flatMap((g) => [{ title: g }, { node: el('div', { class: 'style-grid slstyles' }, SLICER_STYLES.filter((s) => s.group === g).map(chip)) }]);
  openMenu(anchorEl ?? { x: 200, y: 160 }, [...groups, { sep: true }, { label: '새 슬라이서 스타일 (색 · 선 사용자 지정)...', action: () => slicerCustomDialog(sl.id) }], { scroll: true });
}

function slicerCustomDialog(id) {
  const sl = (sheet().slicers ?? []).find((x) => x.id === id);
  if (!sl) return;
  const cur = slicerColors(sl);
  const inputs = CUSTOM_KEYS.map(([k, label]) => [k, label, el('input', { type: 'color', value: cur[k] })]);
  openDialog({
    title: '슬라이서 스타일 수정', width: 380,
    body: el('div', { class: 'fc-list', style: { gap: '6px', padding: '4px 2px' } },
      inputs.map(([, label, inp]) => el('label', { class: 'fc-field', style: { display: 'flex', justifyContent: 'space-between' } }, el('span', {}, label), inp))),
    buttons: [
      { label: '확인', primary: true, action: () => { updateObject(id, { custom: Object.fromEntries(inputs.map(([k, , inp]) => [k, inp.value])) }); gv.renderObjectsAll(); } },
      { label: '기본 스타일로', action: () => { updateObject(id, { custom: undefined }); gv.renderObjectsAll(); } },
      { label: '취소' },
    ],
  });
}

/** 슬라이서 ↔ 피벗 테이블 연결 (보고서 연결 / 필터 연결) */
function slicerConnectionsDialog() {
  const sl = (sheet().slicers ?? []).find((x) => x.id === chartSel);
  const pe = sl ? null : pivotHere();
  if (!sl && !pe) { toast('슬라이서나 피벗 테이블을 선택하세요.'); return; }
  const pivots = allPivots();
  if (sl) {
    if (sl.source?.kind !== 'pivot') { alertDialog('보고서 연결', '표 슬라이서는 그 표에만 연결됩니다.'); return; }
    const cur = new Set(slicerPivotTargets(sl.source).map((e) => `${e.si}:${pivotNameOf(e)}`));
    const checks = pivots.map((e) => [e, el('input', { type: 'checkbox', checked: cur.has(`${e.si}:${pivotNameOf(e)}`) })]);
    openDialog({
      title: `보고서 연결 (${sl.caption})`, width: 380,
      body: el('div', { class: 'fc-list', style: { gap: '4px', padding: '4px 8px', maxHeight: '300px' } },
        checks.map(([e, cb]) => el('label', { class: 'fc-check' }, cb, `${pivotNameOf(e)}  (${wb.sheets[e.si].name})`))),
      buttons: [{
        label: '확인', primary: true, action: () => {
          const chosen = checks.filter(([, cb]) => cb.checked).map(([e]) => e);
          if (!chosen.length) { toast('피벗 테이블을 하나 이상 고르세요.'); return false; }
          wb.transact(() => {
            chosen.forEach((e) => { if (!e.def.name) putPivotDef(e, { ...e.def, name: pivotNameOf(e) }); });
            const source = { kind: 'pivot', field: sl.source.field, pivots: chosen.map((e) => ({ sheet: wb.sheets[e.si].name, name: pivotNameOf(e) })) };
            wb.setSheetProp(si, 'slicers', sheet().slicers.map((x) => (x.id === sl.id ? { ...x, source } : { ...x })));
            // 새로 연결한 피벗도 슬라이서의 현재 선택을 따름
            const sel = (slicerPivotTargets(sl.source)[0]?.def.filters ?? {})[Object.keys(slicerPivotTargets(sl.source)[0]?.def.filters ?? {}).find((k) => k.toLowerCase() === String(sl.source.field).toLowerCase())];
            for (const e of chosen) {
              const nf = { ...(e.def.filters ?? {}) };
              for (const k of Object.keys(nf)) if (k.toLowerCase() === String(sl.source.field).toLowerCase()) delete nf[k];
              if (sel) nf[sl.source.field] = sel;
              putPivotDef(e, { ...e.def, filters: nf });
            }
          }, meta());
          gv.renderObjectsAll();
          return undefined;
        },
      }, { label: '취소' }],
    });
    return;
  }
  // 피벗 테이블 → 이 통합 문서의 피벗 슬라이서 중 연결할 것
  const pname = pivotNameOf(pe);
  const hostName = sheet().name;
  const list = [];
  wb.sheets.forEach((s, i) => (s.slicers ?? []).forEach((x) => { if (x.source?.kind === 'pivot') list.push([i, x]); }));
  if (!list.length) { alertDialog('필터 연결', '이 통합 문서에 피벗 테이블 슬라이서가 없습니다. [슬라이서 삽입]으로 먼저 만드세요.'); return; }
  const isLinked = (i, x) => slicerPivotTargets(x.source, i).some((e) => e.si === si && pivotNameOf(e) === pname);
  const checks = list.map(([i, x]) => [i, x, el('input', { type: 'checkbox', checked: isLinked(i, x) })]);
  openDialog({
    title: `필터 연결 (${pname})`, width: 380,
    body: el('div', { class: 'fc-list', style: { gap: '4px', padding: '4px 8px', maxHeight: '300px' } },
      checks.map(([i, x, cb]) => el('label', { class: 'fc-check' }, cb, `${x.caption}  (${wb.sheets[i].name})`))),
    buttons: [{
      label: '확인', primary: true, action: () => {
        wb.transact(() => {
          if (!pe.def.name) putPivotDef(pe, { ...pe.def, name: pname });
          const bySheet = new Map();
          for (const [i, x, cb] of checks) {
            const targets = slicerPivotTargets(x.source, i).map((e) => ({ sheet: wb.sheets[e.si].name, name: pivotNameOf(e) }))
              .filter((t) => !(t.sheet === hostName && t.name === pname));
            if (cb.checked) targets.push({ sheet: hostName, name: pname });
            if (!targets.length) continue;
            if (!bySheet.has(i)) bySheet.set(i, new Map());
            bySheet.get(i).set(x.id, { kind: 'pivot', field: x.source.field, pivots: targets });
          }
          for (const [i, m] of bySheet) wb.setSheetProp(i, 'slicers', wb.sheets[i].slicers.map((x) => (m.has(x.id) ? { ...x, source: m.get(x.id) } : { ...x })));
        }, meta());
        return undefined;
      },
    }, { label: '취소' }],
  });
}

/** 피벗 스타일 갤러리 */
function pivotStyleGallery(anchorEl, entry = pivotHere()) {
  if (!entry) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const cur = entry.def.style ?? 'PivotStyleLight16';
  const chip = (st) => el('button', {
    class: `style-chip tstyle${st.name === cur ? ' on' : ''}`, title: st.label, onmousedown: (e) => e.preventDefault(),
    style: { width: '42px', background: `linear-gradient(${st.swatch[0]} 0 30%, #fff 30% 40%, ${st.swatch[1]} 40% 70%, ${st.swatch[2]} 70%)` },
    onclick: () => { closeMenus(); setPivotDef(entry, { ...pivotDefV2(entry.def), style: st.name, styleDef: st.def ?? undefined }); focusGrid(); },
  });
  const groups = PIVOT_STYLE_GROUPS.flatMap((g) => [{ title: g }, { node: el('div', { class: 'style-grid pstyles' }, PIVOT_STYLES.filter((s) => s.group === g).map(chip)) }]);
  // 파일에서 가져온 사용자 지정 스타일
  const cd = entry.def.styleDef;
  if (cd) {
    const p = pivotStyleParts(entry.def.style, cd);
    groups.unshift({ title: '사용자 지정' }, { node: el('div', { class: 'style-grid pstyles' }, [chip({ name: entry.def.style, label: entry.def.style, def: cd, swatch: [p.header.fill ?? '#ffffff', p.sub.fill ?? p.body.fill ?? '#ffffff', p.grand.fill ?? '#ffffff'] })]) });
  }
  openMenu(anchorEl ?? { x: 240, y: 160 }, [...groups, { sep: true }, { label: '지우기 (스타일 없음)', action: () => setPivotDef(entry, { ...pivotDefV2(entry.def), style: 'None' }) }], { scroll: true });
}

/** 데이터 원본 변경 (범위 또는 표 이름) */
function pivotChangeSourceDialog(entry = pivotHere()) {
  if (!entry) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const def = entry.def;
  const cur = def.table ?? `${quoteSheetName(def.source)}!${cellName(def.range.r1, def.range.c1)}:${cellName(def.range.r2, def.range.c2)}`;
  formDialog('피벗 테이블 데이터 원본 변경', [{ name: 'src', label: '표/범위', value: cur }], ({ src }) => {
    const text = src.trim().replace(/^=/, '');
    const tf = findTable(wb, text);
    let next;
    if (tf) next = { ...def, table: tf.t.name, source: wb.sheets[tf.si].name, range: { r1: tf.t.r1, c1: tf.t.c1, r2: dataBottom(tf.t), c2: tf.t.c2 } };
    else {
      const bang = text.lastIndexOf('!');
      const sname = bang > 0 ? text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : sheet().name;
      const prg = parseRangeName((bang > 0 ? text.slice(bang + 1) : text).replace(/\$/g, ''));
      if (!prg || wb.sheetIndexByName(sname) < 0 || prg.r2 <= prg.r1) { toast('범위가 올바르지 않습니다. 머리글을 포함한 셀 범위나 표 이름을 입력하세요.'); return false; }
      const { table, ...rest } = def;
      next = { ...rest, source: sname, range: prg };
    }
    // 새 원본에 없는 필드는 영역에서 뺌
    const rows = pivotSource(next);
    if (!rows) { toast('원본을 읽을 수 없습니다.'); return false; }
    const names = new Set(pivotFieldNames(rows, next).map((n) => n.toLowerCase()));
    const keep = (n) => names.has(String(n).toLowerCase());
    next = { ...pivotDefV2(next), rows: (next.rows ?? []).filter(keep), cols: (next.cols ?? []).filter(keep), pages: (next.pages ?? []).filter(keep), values: (next.values ?? []).filter((v) => keep(v.field)) };
    setPivotDef(entry, next);
    refreshPivotPane(true);
    return true;
  }, { note: '표 이름(예: 표1)을 쓰면 표에 데이터가 늘어날 때 새로 고침으로 자동 포함됩니다.' });
}

/** 피벗 테이블 이름 바꾸기 (슬라이서 연결도 함께) */
function renamePivot(newName, entry = pivotHere()) {
  if (!entry) return;
  const name = String(newName ?? '').trim();
  const old = pivotNameOf(entry);
  if (!name || name === old) { updateRibbon(); return; }
  if (allPivots().some((e) => e !== entry && e.si === entry.si && pivotNameOf(e).toLowerCase() === name.toLowerCase())) {
    alertDialog('피벗 테이블 이름', '이 시트에 같은 이름의 피벗 테이블이 있습니다.');
    updateRibbon();
    return;
  }
  const host = wb.sheets[entry.si].name;
  wb.transact(() => {
    putPivotDef(entry, { ...entry.def, name });
    wb.sheets.forEach((s, i) => {
      if (!(s.slicers ?? []).some((x) => x.source?.pivots?.some((p) => (p.sheet ?? s.name) === host && p.name === old))) return;
      wb.setSheetProp(i, 'slicers', s.slicers.map((x) => (x.source?.pivots ? { ...x, source: { ...x.source, pivots: x.source.pivots.map((p) => ((p.sheet ?? s.name) === host && p.name === old ? { ...p, name } : p)) } } : { ...x })));
    });
  }, meta());
  toast(`피벗 테이블 이름을 '${name}'(으)로 바꿨습니다.`);
}


// ───────────────────────── WIXEL 옵션 (엑셀의 [파일] → [옵션]) ─────────────────────────
const OPTIONS_KEY = 'wixel.options';
const PIVOT_DEFAULTS = {
  autoRefresh: true, layout: 'tabular', repeatLabels: false, blankRows: false, subtotals: 'bottom', grand: 'both', mergeLabels: false,
  errorShow: true, errorText: '', emptyShow: true, emptyText: '', autofit: false, preserveFormat: true, style: 'PivotStyleLight16',
};
const OPTION_DEFAULTS = {
  calcMode: 'auto', getPivotData: false, pivotEdit: false, autoDateGroup: false, focusCell: false, focusColor: '#fff4b8', pivot: PIVOT_DEFAULTS,
  // 고급 · 저장 (엑셀 옵션의 '고급' + 위셀 보관함)
  enterDir: 'down', formulaAutocomplete: true, browserMenu: false, libMax: 30, verMinutes: 10,
};
const ENTER_OPP = { down: 'up', up: 'down', right: 'left', left: 'right' };
const opts = (() => {
  try {
    const o = JSON.parse(localStorage.getItem(OPTIONS_KEY) ?? 'null') ?? {};
    return { ...OPTION_DEFAULTS, ...o, pivot: { ...PIVOT_DEFAULTS, ...(o.pivot ?? {}) } };
  } catch {
    return { ...OPTION_DEFAULTS, pivot: { ...PIVOT_DEFAULTS } };
  }
})();
function saveOptions() {
  try { localStorage.setItem(OPTIONS_KEY, JSON.stringify(opts)); } catch { /* 저장 못 함 */ }
}
/** 옵션을 통합 문서 · 화면에 적용 */
function applyOptions() {
  if (wb) wb.manualCalc = opts.calcMode === 'manual';
  document.documentElement.style.setProperty('--focus-cell', opts.focusColor);
  gv?.renderOverlays?.();
  updateStatusCalc();
}
function updateStatusCalc() {
  const elc = document.getElementById('calcState');
  if (!elc) return;
  elc.textContent = wb?.needsCalc ? '계산' : opts.calcMode === 'manual' ? '수동 계산' : '';
  elc.title = wb?.needsCalc ? 'F9를 누르면 계산합니다.' : '';
}

/** 새 피벗 테이블에 기본 레이아웃 적용 (엑셀의 [기본 레이아웃 편집]) */
function pivotDefaultsDef() {
  const p = opts.pivot;
  return {
    style: p.style, layout: p.layout, repeatLabels: p.repeatLabels || undefined, blankRows: p.blankRows || undefined,
    subtotals: p.subtotals !== 'none', subtotalTop: p.subtotals === 'top' ? undefined : false,
    grandRows: p.grand === 'both' || p.grand === 'rows', grandCols: p.grand === 'both' || p.grand === 'cols',
    mergeLabels: p.mergeLabels || undefined, errorCaption: p.errorShow ? p.errorText : undefined, missingCaption: p.emptyShow ? p.emptyText || undefined : undefined,
    autofit: p.autofit ? undefined : false, preserveFormat: p.preserveFormat ? undefined : false, autoRefresh: p.autoRefresh ? true : undefined,
  };
}

function optionsDialog(startTab = 0) {
  const o = JSON.parse(JSON.stringify(opts));
  const radio = (name, v, cur, label, fn) => { const r = el('input', { type: 'radio', name, value: v, checked: cur === v }); r.addEventListener('change', () => { if (r.checked) fn(v); }); return el('label', { class: 'fc-check' }, r, label); };
  const check = (cur, label, fn) => { const c = el('input', { type: 'checkbox', checked: !!cur }); c.addEventListener('change', () => fn(c.checked)); return el('label', { class: 'fc-check' }, c, label); };
  const title = (t) => el('div', { class: 'opt-title' }, t);
  const pages = [
    ['수식', el('div', { class: 'opt-page' },
      title('계산 옵션'),
      el('div', { class: 'muted' }, '통합 문서 계산'),
      radio('calc', 'auto', o.calcMode, '자동', (v) => { o.calcMode = v; }),
      radio('calc', 'manual', o.calcMode, '수동 (F9를 누를 때 계산 — 큰 통합 문서에서 입력이 빠름)', (v) => { o.calcMode = v; }),
      title('수식 작업'),
      check(o.getPivotData, '피벗 테이블 참조에 GetPivotData 함수 사용', (v) => { o.getPivotData = v; }),
      check(o.pivotEdit, '피벗 테이블 값 영역의 셀 편집 허용 (가상 분석 — 원본 데이터와 달라질 수 있음)', (v) => { o.pivotEdit = v; }))],
    ['데이터', el('div', { class: 'opt-page' },
      title('데이터 옵션'),
      check(o.autoDateGroup, '피벗 테이블에서 날짜/시간 열의 자동 그룹화 사용', (v) => { o.autoDateGroup = v; }),
      el('button', { class: 'btn', onclick: () => pivotDefaultsDialog(o) }, '기본 레이아웃 편집...'),
      el('div', { class: 'muted' }, '새 피벗 테이블의 보고서 레이아웃 · 부분합 · 총합계 · 옵션 기본값을 정합니다.'))],
    ['고급', el('div', { class: 'opt-page' },
      title('편집 옵션'),
      el('label', {}, el('span', {}, 'Enter 키를 누른 후 다음 셀로 이동 — 방향 '), (() => {
        const sel = el('select', {}, [['down', '아래쪽'], ['right', '오른쪽'], ['up', '위쪽'], ['left', '왼쪽'], ['none', '이동 안 함']].map(([v, l]) => el('option', { value: v, selected: o.enterDir === v }, l)));
        sel.addEventListener('change', () => { o.enterDir = sel.value; });
        return sel;
      })()),
      check(o.formulaAutocomplete !== false, '수식 자동 완성 (함수 · 이름 목록)', (v) => { o.formulaAutocomplete = v; }),
      title('표시'),
      check(o.browserMenu, '셀에서 브라우저 기본 오른쪽 클릭 메뉴도 허용 (기본: 끔 — 구글 스프레드시트처럼 위셀 메뉴만)', (v) => { o.browserMenu = v; }))],
    ['저장', el('div', { class: 'opt-page' },
      title('이 브라우저의 보관함'),
      el('label', {}, el('span', {}, '최근 문서 보관 개수 '), (() => { const i = el('input', { type: 'number', min: 5, max: 200, value: o.libMax }); i.addEventListener('change', () => { o.libMax = Number(i.value) || 30; }); return i; })()),
      el('label', {}, el('span', {}, '편집 중 버전 기록 간격(분) '), (() => { const i = el('input', { type: 'number', min: 1, max: 240, value: o.verMinutes }); i.addEventListener('change', () => { o.verMinutes = Number(i.value) || 10; }); return i; })()),
      el('div', { class: 'muted' }, `문서마다 버전은 최근 ${VER_MAX}개까지 (이름 붙인 버전은 오래 보관). 셀이 30만 개를 넘는 문서는 큰 문서 자동 저장만 합니다.`),
      el('button', { class: 'btn', onclick: () => { openBackstage('open'); } }, '보관함 열기...'))],
    ['접근성', el('div', { class: 'opt-page' },
      title('포커스 셀'),
      check(o.focusCell, '포커스 셀 사용 (활성 셀의 행과 열을 강조)', (v) => { o.focusCell = v; }),
      el('label', {}, el('span', {}, '강조 색'), (() => { const i = el('input', { type: 'color', value: o.focusColor }); i.addEventListener('input', () => { o.focusColor = i.value; }); return i; })()))],
  ];
  const tabBar = el('div', { class: 'opt-tabs' });
  const box = el('div', { class: 'opt-box' });
  const show = (i) => { [...tabBar.children].forEach((b, j) => b.classList.toggle('on', i === j)); box.replaceChildren(pages[i][1]); };
  pages.forEach(([n], i) => tabBar.append(el('button', { class: 'opt-tab', onclick: () => show(i) }, n)));
  show(startTab);
  openDialog({
    title: 'WIXEL 옵션', width: 640, body: el('div', { class: 'opt-wrap' }, tabBar, box),
    buttons: [{
      label: '확인', primary: true, action: () => {
        const wasManual = opts.calcMode === 'manual';
        Object.assign(opts, o);
        saveOptions();
        applyOptions();
        if (wasManual && opts.calcMode === 'auto') { wb.calculateNow(); gv.renderAll(); }
        gv.renderAll();
      },
    }, { label: '취소' }],
  });
}

/** 피벗 테이블 기본 레이아웃 편집 */
function pivotDefaultsDialog(o) {
  const p = o.pivot;
  formDialog('기본 레이아웃 편집', [
    { name: 'layout', label: '보고서 레이아웃', type: 'select', value: p.layout, options: [{ value: 'compact', label: '압축 형식으로 표시' }, { value: 'outline', label: '개요 형식으로 표시' }, { value: 'tabular', label: '테이블 형식으로 표시' }] },
    { name: 'repeatLabels', label: '모든 항목 레이블 반복', type: 'checkbox', value: p.repeatLabels },
    { name: 'blankRows', label: '각 항목 다음에 빈 줄 삽입', type: 'checkbox', value: p.blankRows },
    { name: 'subtotals', label: '부분합', type: 'select', value: p.subtotals, options: [{ value: 'none', label: '부분합 표시 안 함' }, { value: 'bottom', label: '그룹 하단에 모든 부분합 표시' }, { value: 'top', label: '그룹 상단에 모든 부분합 표시' }] },
    { name: 'grand', label: '총합계', type: 'select', value: p.grand, options: [{ value: 'none', label: '행 및 열의 총합계 해제' }, { value: 'both', label: '행 및 열의 총합계 설정' }, { value: 'rows', label: '행의 총합계만 설정' }, { value: 'cols', label: '열의 총합계만 설정' }] },
    { name: 'mergeLabels', label: '레이블이 있는 셀 병합 및 가운데 맞춤', type: 'checkbox', value: p.mergeLabels },
    { name: 'errorShow', label: '오류 값 표시', type: 'checkbox', value: p.errorShow },
    { name: 'errorText', label: '오류 값 표시 글자', value: p.errorText },
    { name: 'emptyShow', label: '빈 셀 표시', type: 'checkbox', value: p.emptyShow },
    { name: 'emptyText', label: '빈 셀 표시 글자', value: p.emptyText },
    { name: 'autofit', label: '업데이트 시 열 자동 맞춤', type: 'checkbox', value: p.autofit },
    { name: 'preserveFormat', label: '업데이트 시 셀 서식 유지', type: 'checkbox', value: p.preserveFormat },
    { name: 'autoRefresh', label: '원본 데이터가 바뀌면 자동 새로 고침 (WIXEL)', type: 'checkbox', value: p.autoRefresh },
    { name: 'style', label: '기본 스타일', type: 'select', value: p.style, options: PIVOT_STYLES.map((s) => ({ value: s.name, label: s.label })) },
  ], (v) => { Object.assign(p, v); }, { note: '새로 만드는 피벗 테이블에 적용됩니다.' });
}

/** 피벗 테이블 옵션 (엑셀과 같은 탭: 레이아웃 및 서식 · 요약 및 필터 · 표시 · 인쇄 · 데이터 · 대체 텍스트) */
function pivotOptionsDialog(entry = pivotHere(), startTab = 0) {
  if (!entry) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const def = pivotDefV2(entry.def);
  const v = {
    name: pivotNameOf(entry), mergeLabels: !!def.mergeLabels, indent: def.indent ?? 1, pageOrder: def.pageOrder ?? 'down', pageWrap: def.pageWrap ?? 0,
    errorShow: def.errorShow !== false, errorText: def.errorCaption ?? '',
    emptyShow: def.emptyShow !== false, emptyText: def.missingCaption ?? '',
    autofit: def.autofit !== false, preserveFormat: def.preserveFormat !== false,
    grandRows: def.grandRows !== false, grandCols: def.grandCols !== false, subtotalHidden: !!def.subtotalHidden, multiFilters: !!def.multiFilters, customListSort: def.customListSort !== false,
    showExpand: def.showExpand !== false, tooltips: def.tooltips !== false, fieldCaptions: def.fieldCaptions !== false, classic: !!def.classic, emptyRowsItems: !!def.showEmptyRows, emptyColsItems: !!def.showEmptyCols,
    showValuesRow: !!def.showValuesRow, sortAZ: def.fieldListSort === 'az',
    printExpand: !!def.printExpand, printTitles: !!def.printTitles,
    autoRefresh: !!def.autoRefresh, saveData: def.saveData !== false, refreshOnOpen: !!def.refreshOnOpen, missingItems: def.missingItems ?? 'auto', enableDrill: def.enableDrill !== false,
    altTitle: def.altTitle ?? '', altDesc: def.altDesc ?? '',
    rowCaption: def.rowCaption ?? '행 레이블', colCaption: def.colCaption ?? '열 레이블',
  };
  const chk = (k, label) => { const c = el('input', { type: 'checkbox', checked: !!v[k] }); c.addEventListener('change', () => { v[k] = c.checked; }); return el('label', { class: 'fc-check' }, c, label); };
  const txt = (k, label, w = 160) => { const i = el('input', { type: 'text', value: v[k] ?? '', style: { width: `${w}px` } }); i.addEventListener('input', () => { v[k] = i.value; }); return el('label', {}, el('span', {}, label), i); };
  const num = (k, label) => { const i = el('input', { type: 'number', value: v[k], min: 0, max: 127, style: { width: '70px' } }); i.addEventListener('input', () => { v[k] = Number(i.value) || 0; }); return el('label', {}, el('span', {}, label), i); };
  const sel = (k, label, list) => { const s2 = el('select', {}, list.map(([a, b]) => el('option', { value: a, selected: String(v[k]) === String(a) }, b))); s2.addEventListener('change', () => { v[k] = s2.value; }); return el('label', {}, el('span', {}, label), s2); };
  const t = (x) => el('div', { class: 'opt-title' }, x);
  const pages = [
    ['레이아웃 및 서식', el('div', { class: 'opt-page' },
      t('레이아웃'), chk('mergeLabels', '레이블이 있는 셀 병합 및 가운데 맞춤'), num('indent', '압축 형식일 때 행 레이블 들여쓰기(문자)'),
      sel('pageOrder', '보고서 필터 영역에 필드 표시', [['down', '행 우선'], ['over', '열 우선']]), num('pageWrap', '보고서 필터 열/행당 필드 수'),
      t('서식'), chk('errorShow', '오류 값 표시'), txt('errorText', '　표시 글자'), chk('emptyShow', '빈 셀 표시'), txt('emptyText', '　표시 글자'),
      chk('autofit', '업데이트 시 열 자동 맞춤'), chk('preserveFormat', '업데이트 시 셀 서식 유지'))],
    ['요약 및 필터', el('div', { class: 'opt-page' },
      t('총합계'), chk('grandRows', '행 총합계 표시'), chk('grandCols', '열 총합계 표시'),
      t('필터'), chk('subtotalHidden', '필터링된 페이지 항목 부분합에 포함'), chk('multiFilters', '필드당 여러 필터 허용'),
      t('정렬'), chk('customListSort', '정렬할 때 사용자 지정 목록 사용'))],
    ['표시', el('div', { class: 'opt-page' },
      t('표시'), chk('showExpand', '확장/축소 단추 표시'), chk('tooltips', '상황에 맞는 도구 설명 표시'), chk('fieldCaptions', '필드 캡션 및 필터 드롭다운 표시'),
      chk('classic', '클래식 피벗 테이블 레이아웃 (표 안으로 필드 끌어 놓기)'), chk('showValuesRow', '값 행 표시'), chk('emptyRowsItems', '행에 데이터가 없는 항목 표시'), chk('emptyColsItems', '열에 데이터가 없는 항목 표시'),
      txt('rowCaption', '행 레이블 캡션'), txt('colCaption', '열 레이블 캡션'),
      t('필드 목록'), chk('sortAZ', '오름차순 정렬 (해제하면 데이터 원본 순서)'))],
    ['인쇄', el('div', { class: 'opt-page' },
      chk('printExpand', '피벗 테이블에 확장/축소 단추가 표시될 때 인쇄'), chk('printTitles', '인쇄 제목 설정 (각 페이지에 행 · 열 레이블 반복)'))],
    ['데이터', el('div', { class: 'opt-page' },
      t('피벗 테이블 데이터'), chk('autoRefresh', '원본 데이터가 바뀌면 자동 새로 고침 (WIXEL)'), chk('saveData', '파일에 원본 데이터 저장'), chk('enableDrill', '세부 정보 표시 사용 (값 셀 두 번 클릭)'), chk('refreshOnOpen', '파일을 열 때 데이터 새로 고침'),
      t('데이터 원본에서 삭제된 항목 보존'), sel('missingItems', '필드당 반환할 항목 수', [['auto', '자동'], ['none', '없음'], ['max', '최대']]))],
    ['대체 텍스트', el('div', { class: 'opt-page' }, txt('altTitle', '제목', 320), txt('altDesc', '설명', 320))],
  ];
  const tabBar = el('div', { class: 'opt-tabs' });
  const box = el('div', { class: 'opt-box' });
  const show = (i) => { [...tabBar.children].forEach((b, j) => b.classList.toggle('on', i === j)); box.replaceChildren(pages[i][1]); };
  pages.forEach(([n], i) => tabBar.append(el('button', { class: 'opt-tab', onclick: () => show(i) }, n)));
  show(startTab);
  const nameIn = el('input', { type: 'text', value: v.name, style: { width: '240px' } });
  openDialog({
    title: '피벗 테이블 옵션', width: 620, body: el('div', {}, el('label', {}, el('span', {}, '피벗 테이블 이름'), nameIn), el('div', { class: 'opt-wrap' }, tabBar, box)),
    buttons: [{
      label: '확인', primary: true, action: () => {
        const next = {
          ...def, mergeLabels: v.mergeLabels || undefined, indent: v.indent === 1 ? undefined : v.indent, pageOrder: v.pageOrder === 'down' ? undefined : v.pageOrder, pageWrap: v.pageWrap || undefined,
          errorShow: v.errorShow ? undefined : false, errorCaption: v.errorShow ? v.errorText : undefined, emptyShow: v.emptyShow ? undefined : false, missingCaption: v.emptyShow && v.emptyText ? v.emptyText : undefined,
          autofit: v.autofit ? undefined : false, preserveFormat: v.preserveFormat ? undefined : false,
          grandRows: v.grandRows, grandCols: v.grandCols, subtotalHidden: v.subtotalHidden || undefined, multiFilters: v.multiFilters || undefined, customListSort: v.customListSort ? undefined : false,
          showExpand: v.showExpand ? undefined : false, tooltips: v.tooltips ? undefined : false, fieldCaptions: v.fieldCaptions ? undefined : false, classic: v.classic || undefined,
          showEmptyRows: v.emptyRowsItems || undefined, showEmptyCols: v.emptyColsItems || undefined, showValuesRow: v.showValuesRow || undefined, fieldListSort: v.sortAZ ? 'az' : undefined,
          printExpand: v.printExpand || undefined, printTitles: v.printTitles || undefined, autoRefresh: v.autoRefresh || undefined, saveData: v.saveData ? undefined : false, refreshOnOpen: v.refreshOnOpen || undefined,
          missingItems: v.missingItems === 'auto' ? undefined : v.missingItems, enableDrill: v.enableDrill ? undefined : false,
          altTitle: v.altTitle || undefined, altDesc: v.altDesc || undefined,
          rowCaption: v.rowCaption === '행 레이블' ? undefined : v.rowCaption, colCaption: v.colCaption === '열 레이블' ? undefined : v.colCaption,
        };
        setPivotDef(entry, next);
        if (nameIn.value.trim() && nameIn.value.trim() !== pivotNameOf(entry)) renamePivot(nameIn.value, entry);
        refreshPivotPane(true);
      },
    }, { label: '취소' }],
  });
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
        title: 'WIXEL', body: '여기에 이미 데이터가 있습니다. 바꾸시겠습니까?',
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
  mf: ['insertFunction', '함수 삽입'], mua: ['autosum', '자동 합계'], mn: ['nameManager', '이름 관리자'], mmd: ['defineName', '이름 정의'],
  ms: ['pasteName', '수식에서 사용'], mc: ['createNamesFromSel', '선택 영역에서 만들기'],
  hvv: ['pasteValuesKey', '값 붙여넣기'], hvs: ['pasteSpecial', '선택하여 붙여넣기'], es: ['pasteSpecial', '선택하여 붙여넣기'],
  hac: ['alignCenter', '가운데 맞춤'], hal: ['alignLeft', '왼쪽 맞춤'], har: ['alignRight', '오른쪽 맞춤'],
  hir: ['insertRows', '시트 행 삽입'], hic: ['insertCols', '시트 열 삽입'], hdr: ['deleteRows', '시트 행 삭제'], hdc: ['deleteCols', '시트 열 삭제'],
  hfn: ['fontDialog', '글꼴'], hef: ['clearFormats', '서식 지우기'], hea: ['clearAll', '모두 지우기'], ni: ['hyperlink', '링크'],
  ase: ['flashFill', '빠른 채우기'], mv: ['pivotRefresh', '피벗 새로 고침'], f: ['backstage', '파일'],
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
  if (e.ctrlKey || e.metaKey) {
    // Ctrl+Alt+키 조합은 리본 키 팁이 아님 (Ctrl+Alt+V 등)
    if (keytip) endKeytip();
    return false;
  }
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

/** 개체 위치 속성: twoCell = 셀에 맞춰 위치와 크기 변경(기본), oneCell = 위치만 변경, absolute = 변경 안 함 */
const placementOf = (prop, o) => o.placement ?? (prop === 'slicers' ? 'oneCell' : 'twoCell');

/**
 * 행/열 크기 변경 · 삽입 · 삭제 뒤에도 개체가 셀을 따라가게 (엑셀의 개체 위치 속성). transact 안에서 부름.
 * shift: 삽입/삭제일 때 { axis: 'row'|'col', index, count(음수 = 삭제) }
 */
function anchorObjects(fn, shift = null, moveOnly = false) {
  const s = sheet();
  const props = OBJECT_PROPS.filter((p) => (s[p] ?? []).length);
  if (!props.length) return fn();
  gv.refreshAxes();
  const cols0 = gv.cols;
  const rows0 = gv.rows;
  const at = (ax, p) => { const i = ax.indexAt(Math.max(0, p)); return [i, Math.max(0, p - ax.pos(i))]; };
  const anchors = new Map();
  for (const p of props) {
    for (const o of s[p]) {
      const place = placementOf(p, o);
      if (place === 'absolute') continue;
      anchors.set(o.id, { place, c1: at(cols0, o.x), r1: at(rows0, o.y), c2: at(cols0, o.x + o.w), r2: at(rows0, o.y + o.h) });
    }
  }
  const res = fn();
  gv.refreshAxes();
  const move = (ax, [i, off], kind) => {
    let j = i;
    let o = off;
    if (shift && shift.axis === kind) {
      if (shift.count > 0 && i >= shift.index) j = i + shift.count;
      else if (shift.count < 0) {
        const end = shift.index - shift.count;
        if (i >= end) j = i + shift.count;
        else if (i >= shift.index) { j = shift.index; o = 0; }
      }
    }
    return ax.pos(j) + Math.min(o, ax.size(j) || o);
  };
  for (const p of props) {
    const list = s[p];
    let changed = false;
    const next = list.map((o) => {
      const a = anchors.get(o.id);
      if (!a) return o;
      const x = Math.round(move(gv.cols, a.c1, 'col'));
      const y = Math.round(move(gv.rows, a.r1, 'row'));
      const w = a.place === 'twoCell' && !moveOnly ? Math.max(1, Math.round(move(gv.cols, a.c2, 'col')) - x) : o.w;
      const h = a.place === 'twoCell' && !moveOnly ? Math.max(LINE_SHAPES.has(o.kind) ? 0 : 1, Math.round(move(gv.rows, a.r2, 'row')) - y) : o.h;
      if (x === o.x && y === o.y && w === o.w && h === o.h) return o;
      changed = true;
      return { ...o, x, y, w, h };
    });
    if (changed) wb.setSheetProp(si, p, next);
  }
  return res;
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
      ...(f.obj.pivot ? [{ label: '피벗 차트 지표 선택...', icon: 'pivot', action: () => pivotChartMetrics(id) }] : []),
      { title: '차트 종류 변경' },
      ...CHART_TYPES.map((t) => ({ label: t.label, checked: f.obj.type === t.id, action: () => updateChart(id, { type: t.id }) })),
    );
  } else if (f.prop === 'shapes') {
    items.push({ label: LINE_SHAPES.has(f.obj.kind) ? '선 서식...' : '텍스트 편집 및 도형 서식...', icon: 'shapes', action: () => shapeDialog(id) });
    if (!LINE_SHAPES.has(f.obj.kind)) {
      items.push({ label: '도형 모양 변경...', action: () => setTimeout(() => openMenu({ x: 260, y: 140 }, [{ node: shapeGallery((k) => updateObject(id, { kind: k }), true) }], { scroll: true }), 0) });
    }
    items.push({ sep: true }, { label: '맨 앞으로 가져오기', action: () => arrangeObject(id, 'front') }, { label: '앞으로 가져오기', action: () => arrangeObject(id, 'forward') },
      { label: '뒤로 보내기', action: () => arrangeObject(id, 'backward') }, { label: '맨 뒤로 보내기', action: () => arrangeObject(id, 'back') });
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
      { label: '셀에 배치', icon: 'picture', action: () => imageToCell(id) },
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

/** 그림 파일을 셀에 배치 (Excel '셀에 배치'): 셀 값이 그림이 됨 */
function insertPictureInCell() {
  const input = el('input', { type: 'file', accept: 'image/*' });
  input.addEventListener('change', () => { if (input.files[0]) placeImageFileInCell(input.files[0]); });
  input.click();
}

/** 큰 그림은 셀용으로 줄여서 data URL 로 (저장 크기 절약) */
function shrinkImage(src, max, done) {
  const img = new Image();
  img.onload = () => {
    const w0 = img.naturalWidth || 1;
    const h0 = img.naturalHeight || 1;
    const k = Math.min(1, max / w0, max / h0);
    if (k >= 1 || /^data:image\/(svg|gif)/.test(src)) { done(src); return; }
    const cv = document.createElement('canvas');
    cv.width = Math.max(1, Math.round(w0 * k));
    cv.height = Math.max(1, Math.round(h0 * k));
    cv.getContext('2d').drawImage(img, 0, 0, cv.width, cv.height);
    done(cv.toDataURL(/^data:image\/jpe?g/.test(src) ? 'image/jpeg' : 'image/png', 0.9));
  };
  img.onerror = () => alertDialog('그림 삽입', '이 그림 형식은 표시할 수 없습니다. PNG·JPEG·GIF·SVG·WebP 파일을 사용하세요.');
  img.src = src;
}

function placeImageFileInCell(file) {
  if (file.size > 10 * 1024 * 1024) { alertDialog('그림 삽입', '10MB 이하의 그림만 넣을 수 있습니다.'); return; }
  const reader = new FileReader();
  reader.onload = () => shrinkImage(reader.result, 800, (src) => {
    putCellImage(active.r, active.c, { src, alt: file.name.replace(/\.[^.]+$/, '') });
  });
  reader.readAsDataURL(file);
}

function putCellImage(r, c, image) {
  wb.transact(() => {
    const cur = wb.getCell(si, r, c);
    wb.setCellData(si, r, c, { raw: '', style: cur?.style, comment: cur?.comment, link: cur?.link, image });
  }, meta());
  focusGrid();
}

/** 셀 위에 떠 있는 그림 → 왼쪽 위 모서리가 있는 셀에 배치 */
function imageToCell(id) {
  const im = sheet().images.find((x) => x.id === id);
  if (!im) return;
  const r = gv.rows.indexAt(im.y + 1);
  const c = gv.cols.indexAt(im.x + 1);
  shrinkImage(im.src, 800, (src) => wb.transact(() => {
    const cur = wb.getCell(si, r, c);
    wb.setCellData(si, r, c, { raw: '', style: cur?.style, comment: cur?.comment, link: cur?.link, image: { src, alt: im.name ?? '' } });
    wb.setSheetProp(si, 'images', sheet().images.filter((x) => x.id !== id));
    chartSel = null;
    selectCell(r, c);
  }, meta()));
}

/** 셀에 배치한 그림 → 셀 위에 떠 있는 그림 */
function cellImageToFloating(r, c) {
  const cell = wb.getCell(si, r, c);
  if (!cell?.image) return;
  const img = new Image();
  img.onload = () => {
    const w0 = img.naturalWidth || 200;
    const h0 = img.naturalHeight || 150;
    const k = Math.min(1, 480 / w0, 360 / h0);
    const rc = gv.sheetRect({ r1: r, c1: c, r2: r, c2: c });
    wb.transact(() => {
      wb.setCellData(si, r, c, { raw: '', style: cell.style, comment: cell.comment, link: cell.link });
      const obj = { id: newObjId('im'), name: cell.image.alt || '그림', x: Math.round(rc.x), y: Math.round(rc.y),
        w: Math.max(8, Math.round(w0 * k)), h: Math.max(8, Math.round(h0 * k)), src: cell.image.src, z: nextZ() };
      wb.setSheetProp(si, 'images', [...(sheet().images ?? []), obj]);
      chartSel = obj.id;
    }, meta());
    gv.renderObjectsAll();
    updateSelectionUI();
  };
  img.src = cell.image.src;
}

function cellImageAltDialog(r, c) {
  const cell = wb.getCell(si, r, c);
  if (!cell?.image) return;
  formDialog('대체 텍스트', [{ name: 'alt', label: '설명 (그림을 볼 수 없을 때 표시)', value: cell.image.alt ?? '' }], ({ alt }) => {
    wb.transact(() => wb.setCellData(si, r, c, { raw: '', style: cell.style, comment: cell.comment, link: cell.link, image: { ...cell.image, alt: alt.trim() } }), meta());
  });
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

/** 그림 주소(data URL 또는 웹 주소) → 시트에 그림 개체 추가 */
function addImageSrc(src, name, at = null) {
  const img = new Image();
  img.onload = () => {
    const w0 = img.naturalWidth || 200;
    const h0 = img.naturalHeight || 150;
    const k = Math.min(1, 480 / w0, 360 / h0);
    const p = at ?? objectOrigin();
    addObject('images', { id: newObjId('im'), name: name || '그림', x: p.x, y: p.y, w: Math.max(8, Math.round(w0 * k)), h: Math.max(8, Math.round(h0 * k)), src });
    focusGrid();
  };
  img.onerror = () => alertDialog('그림 삽입', '그림을 불러올 수 없습니다. 주소가 그림 파일을 가리키는지, 인터넷에 연결되어 있는지 확인하세요.');
  img.src = src;
}

/** 웹 그림 → data URL (파일에 함께 저장되게). 사이트가 허용하지 않으면(CORS) null */
async function fetchImageData(url) {
  try {
    const res = await fetch(url, { mode: 'cors' });
    if (!res.ok) return null;
    const blob = await res.blob();
    if (!/^image\//.test(blob.type) || blob.size > 10 * 1024 * 1024) return null;
    return await new Promise((ok) => { const r = new FileReader(); r.onload = () => ok(r.result); r.onerror = () => ok(null); r.readAsDataURL(blob); });
  } catch {
    return null;
  }
}

// 온라인 그림 검색: 크리에이티브 커먼즈 (Openverse → 안 되면 Wikimedia Commons)
async function searchOnlineImages(q, page = 1) {
  try {
    const res = await fetch(`https://api.openverse.org/v1/images/?q=${encodeURIComponent(q)}&page_size=30&page=${page}`);
    if (res.ok) {
      const j = await res.json();
      return (j.results ?? []).map((x) => ({ thumb: x.thumbnail ?? x.url, full: x.thumbnail ?? x.url, title: x.title ?? q, credit: [x.creator, x.license ? `CC ${String(x.license).toUpperCase()}` : ''].filter(Boolean).join(' · '), page: x.foreign_landing_url }));
    }
  } catch { /* 다음 방법 */ }
  const url = `https://commons.wikimedia.org/w/api.php?action=query&format=json&origin=*&generator=search&gsrnamespace=6&gsrlimit=30&gsroffset=${(page - 1) * 30}&gsrsearch=${encodeURIComponent(q)}&prop=imageinfo&iiprop=url|extmetadata&iiurlwidth=480`;
  const res = await fetch(url);
  if (!res.ok) throw new Error(`HTTP ${res.status}`);
  const j = await res.json();
  return Object.values(j.query?.pages ?? {}).filter((p) => p.imageinfo?.[0]?.thumburl).map((p) => {
    const ii = p.imageinfo[0];
    const md = ii.extmetadata ?? {};
    const strip = (h) => String(h ?? '').replace(/<[^>]*>/g, '').trim();
    return { thumb: ii.thumburl, full: ii.thumburl, title: p.title.replace(/^File:/, '').replace(/\.[^.]+$/, ''), credit: [strip(md.Artist?.value), strip(md.LicenseShortName?.value)].filter(Boolean).join(' · '), page: ii.descriptionurl };
  });
}

/** 온라인 그림 (엑셀의 [삽입] → [그림] → [온라인 그림]): 검색 또는 웹 주소 */
function onlinePictureDialog(inCell) {
  const q = el('input', { type: 'text', placeholder: '검색어 (예: 커피, 그래프, 사무실)', style: { flex: '1' } });
  const urlIn = el('input', { type: 'url', placeholder: 'https://… 그림 주소', style: { flex: '1' } });
  const grid = el('div', { class: 'online-grid' });
  const status = el('div', { class: 'muted', style: { fontSize: '12px', minHeight: '16px' } }, '크리에이티브 커먼즈 그림을 검색합니다. 사용 조건(라이선스)을 확인하고 쓰세요.');
  const chosen = new Map();
  let page = 1;
  let lastQ = '';
  const run = async (more = false) => {
    const text = q.value.trim();
    if (!text) return;
    if (!more) { page = 1; lastQ = text; grid.replaceChildren(); chosen.clear(); } else page++;
    status.textContent = '검색 중…';
    try {
      const list = await searchOnlineImages(lastQ, page);
      if (!list.length && !more) { status.textContent = '결과가 없습니다. 다른 검색어를 써 보세요.'; return; }
      for (const it of list) {
        const card = el('button', { class: 'online-item', title: `${it.title}${it.credit ? `\n${it.credit}` : ''}` },
          el('img', { src: it.thumb, alt: it.title, loading: 'lazy', referrerpolicy: 'no-referrer' }),
          el('span', {}, it.credit || it.title));
        card.addEventListener('click', () => {
          if (chosen.has(it.thumb)) { chosen.delete(it.thumb); card.classList.remove('on'); } else { chosen.set(it.thumb, it); card.classList.add('on'); }
          status.textContent = chosen.size ? `${chosen.size}개 선택됨` : '';
        });
        grid.append(card);
      }
      status.textContent = `${grid.children.length}개 그림 — 클릭해서 고르고 [삽입]을 누르세요.`;
    } catch (e) {
      status.textContent = `검색할 수 없습니다 (${e.message}). 인터넷 연결 또는 이 페이지의 외부 접속 허용 여부를 확인하세요. 웹 주소로 넣을 수도 있습니다.`;
    }
  };
  q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); run(); } });
  const more = el('button', { class: 'btn', onclick: () => run(true) }, '더 보기');
  const insertOne = async (src, name, k) => {
    const data = (await fetchImageData(src)) ?? src;
    if (inCell) {
      shrinkImage(data, 800, (s2) => putCellImage(active.r + k, active.c, { src: s2, alt: name }));
    } else {
      const p = objectOrigin();
      addImageSrc(data, name, { x: p.x + k * 24, y: p.y + k * 24 });
    }
  };
  openDialog({
    title: '온라인 그림', width: 640,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
      el('div', { style: { display: 'flex', gap: '6px' } }, q, el('button', { class: 'btn primary', onclick: () => run() }, '검색')),
      grid, el('div', { style: { display: 'flex', justifyContent: 'space-between', alignItems: 'center' } }, status, more),
      el('div', { class: 'menu-title', style: { padding: '6px 0 0' } }, '웹 주소로 삽입'),
      el('div', { style: { display: 'flex', gap: '6px' } }, urlIn)),
    buttons: [
      {
        label: '삽입', primary: true, action: () => {
          const list = [...chosen.values()];
          const u = urlIn.value.trim();
          if (u) list.push({ full: u, title: u.split('/').pop().replace(/\.[^.]+$/, '') || '그림' });
          if (!list.length) { toast('그림을 고르거나 웹 주소를 입력하세요.'); return false; }
          list.forEach((it, k) => insertOne(it.full, it.title, k));
          return undefined;
        },
      },
      { label: '취소' },
    ],
  });
  setTimeout(() => q.focus(), 0);
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

/** 도형 갤러리 (엑셀처럼 분류별 견본 격자) — noLines: 도형 모양 변경용 */
function shapeGallery(pick, noLines = false) {
  const icon = (id) => {
    const sh = newShape(id, { x: 0, y: 0, w: 20, h: 16 });
    return shapeSvg({ ...sh, fill: sh.fill ? (id === 'textbox' ? '#ffffff' : '#dbe5f5') : null, stroke: '#44546a', strokeWidth: 1, flipV: LINE_SHAPES.has(sh.kind) ? true : undefined });
  };
  return el('div', { class: 'shape-gallery' }, SHAPE_GROUPS.filter(([g]) => !(noLines && g === '선')).map(([g, list]) => [
    el('div', { class: 'menu-title' }, g),
    el('div', { class: 'shape-grid' }, list.filter(([id]) => !(noLines && id === 'textbox')).map(([id, label]) => el('button', {
      class: 'shape-btn', title: label, html: icon(id), onmousedown: (e) => e.preventDefault(),
      onclick: () => { closeMenus(); pick(id); },
    }))),
  ]));
}

/** 개체 겹치는 순서: 맨 앞 · 앞으로 · 뒤로 · 맨 뒤 (차트 · 그림 · 도형 · 슬라이서 공통) */
function arrangeObject(id, how) {
  const s = sheet();
  const all = OBJECT_PROPS.flatMap((p) => (s[p] ?? []).map((o) => ({ p, o })));
  all.sort((a, b) => (a.o.z ?? 0) - (b.o.z ?? 0));
  const i = all.findIndex((x) => x.o.id === id);
  if (i < 0) return;
  const [it] = all.splice(i, 1);
  const at = how === 'front' ? all.length : how === 'back' ? 0 : how === 'forward' ? Math.min(all.length, i + 1) : Math.max(0, i - 1);
  all.splice(at, 0, it);
  const zOf = new Map(all.map((x, k) => [x.o.id, k + 1]));
  wb.transact(() => {
    for (const p of OBJECT_PROPS) if ((s[p] ?? []).length) wb.setSheetProp(si, p, s[p].map((o) => ({ ...o, z: zOf.get(o.id) ?? o.z })));
  }, meta());
  gv.renderObjectsAll();
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
  const isLine = LINE_SHAPES.has(sh.kind);
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
  const dash = el('select', {}, [['', '실선'], ['dash', '파선'], ['dot', '점선']].map(([v, l]) => el('option', { value: v, selected: (sh.dash ?? '') === v }, l)));
  const arrow = el('select', {}, [['', '없음'], ['end', '끝 화살표'], ['both', '양쪽 화살표']].map(([v, l]) => el('option', { value: v, selected: (sh.arrow ?? '') === v }, l)));
  const rot = el('input', { type: 'number', min: -360, max: 360, value: sh.rot ?? 0, style: { width: '70px' } });
  const valign = el('select', {}, [['top', '위쪽'], ['middle', '가운데'], ['bottom', '아래쪽']].map(([v, l]) => el('option', { value: v, selected: (sh.valign ?? (sh.kind === 'textbox' ? 'top' : 'middle')) === v }, l)));
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
    isLine ? null : row('텍스트', text),
    isLine ? null : row('글꼴', el('span', {}, '크기'), size, el('label', { style: { display: 'inline-flex', gap: '4px', flexDirection: 'row' } }, bold, '굵게'), color),
    isLine ? null : row('맞춤', align, el('span', {}, '세로'), valign),
    isLine ? null : row('채우기', el('label', { style: { display: 'inline-flex', gap: '4px', flexDirection: 'row' } }, fillOn, '사용'), fill),
    row(isLine ? '선 색' : '윤곽선', isLine ? null : el('label', { style: { display: 'inline-flex', gap: '4px', flexDirection: 'row' } }, lineOn, '사용'), stroke),
    row('선 굵기(pt)', width, dash),
    isLine ? row('화살표', arrow) : row('회전(°)', rot));
  openDialog({
    title: isLine ? '선 서식' : '도형 서식', body, width: 420,
    onOpen: () => { if (!isLine) { text.focus(); text.setSelectionRange(text.value.length, text.value.length); } },
    buttons: [
      {
        label: '확인', primary: true,
        action: () => updateObject(id, isLine
          ? { stroke: stroke.value, strokeWidth: Number(width.value) || 1, dash: dash.value || undefined, arrow: arrow.value || undefined }
          : {
            text: text.value, size: clamp(Number(size.value) || 11, 6, 96), bold: bold.checked || undefined, align: align.value, color: color.value,
            // 파일에서 가져온 조각별 서식은 글자 · 글꼴 설정을 바꾸지 않았을 때만 유지
            ...(sh.paras && (text.value !== sh.text || align.value !== (sh.align ?? 'left') || bold.checked !== !!sh.bold) ? { paras: undefined } : {}),
            fill: fillOn.checked ? fill.value : null, stroke: lineOn.checked ? stroke.value : null, strokeWidth: Number(width.value) || 1,
            dash: dash.value || undefined, rot: (Number(rot.value) % 360) || undefined, valign: valign.value,
          }),
      },
      { label: '취소' },
    ],
  });
}


// ───────────────────────── 셰이프 형식 (도형 · 그림 · 차트 · 슬라이서 공통) ─────────────────────────
const objMulti = new Set(); // Ctrl/Shift+클릭으로 함께 고른 개체 (chartSel 이 기준 개체)

/** 고른 개체들 → [{ prop, obj }] (기준 개체가 처음) */
function selectedObjects() {
  const s = sheet();
  return [chartSel, ...objMulti].filter(Boolean).map((id) => findObject(s, id)).filter(Boolean);
}

/** 고른 개체 모두 바꾸기 (patch 또는 (obj, prop) => patch) — 한 번의 실행 취소 */
function patchObjects(patch, kinds = null) {
  const list = selectedObjects().filter((f) => !kinds || kinds.includes(f.prop));
  if (!list.length) { toast('개체를 선택하세요.'); return; }
  const byProp = new Map();
  for (const f of list) { if (!byProp.has(f.prop)) byProp.set(f.prop, new Map()); byProp.get(f.prop).set(f.obj.id, typeof patch === 'function' ? patch(f.obj, f.prop) : patch); }
  wb.transact(() => {
    for (const [prop, m] of byProp) {
      wb.setSheetProp(si, prop, sheet()[prop].map((o) => (m.has(o.id) && m.get(o.id) ? { ...o, ...m.get(o.id) } : o)));
    }
  }, meta());
  gv.renderObjectsAll();
  updateSelectionUI();
}

/** 맞춤 · 배분 (여러 개체: 선택 영역 기준, 하나: 보이는 화면 기준이 아니라 격자(셀)에 맞춤) */
function alignObjects(how) {
  const list = selectedObjects();
  if (!list.length) return;
  if (how === 'grid') {
    gv.refreshAxes();
    patchObjects((o) => {
      const c = gv.cols.indexAt(o.x); const r = gv.rows.indexAt(o.y);
      const x = gv.cols.pos(o.x - gv.cols.pos(c) > gv.cols.size(c) / 2 ? c + 1 : c);
      const y = gv.rows.pos(o.y - gv.rows.pos(r) > gv.rows.size(r) / 2 ? r + 1 : r);
      return { x, y };
    });
    return;
  }
  if (list.length < 2 && !how.startsWith('dist')) { toast('맞춤: 개체를 두 개 이상 고르세요 (Ctrl 또는 Shift+클릭).'); return; }
  const L = minOf(list.map((f) => f.obj.x));
  const T = minOf(list.map((f) => f.obj.y));
  const R = maxOf(list.map((f) => f.obj.x + f.obj.w));
  const B = maxOf(list.map((f) => f.obj.y + f.obj.h));
  if (how === 'distH' || how === 'distV') {
    if (list.length < 3) { toast('배분: 개체를 세 개 이상 고르세요.'); return; }
    const h = how === 'distH';
    const sorted = [...list].sort((a, b) => (h ? a.obj.x - b.obj.x : a.obj.y - b.obj.y));
    const total = sorted.reduce((a, f) => a + (h ? f.obj.w : f.obj.h), 0);
    const gap = ((h ? R - L : B - T) - total) / (sorted.length - 1);
    let p = h ? L : T;
    const pos = new Map();
    for (const f of sorted) { pos.set(f.obj.id, Math.round(p)); p += (h ? f.obj.w : f.obj.h) + gap; }
    patchObjects((o) => (h ? { x: pos.get(o.id) } : { y: pos.get(o.id) }));
    return;
  }
  patchObjects((o) => ({
    left: { x: L }, center: { x: Math.round((L + R) / 2 - o.w / 2) }, right: { x: R - o.w },
    top: { y: T }, middle: { y: Math.round((T + B) / 2 - o.h / 2) }, bottom: { y: B - o.h },
  }[how]));
}

function rotateObjects(how) {
  patchObjects((o) => {
    if (how === 'r90' || how === 'l90') return { rot: ((((o.rot ?? 0) + (how === 'r90' ? 90 : -90)) % 360) + 360) % 360 || undefined };
    if (how === 'flipH') return { flip: !o.flip || undefined };
    if (how === 'flipV') return { flipV: !o.flipV || undefined };
    return null;
  }, ['shapes', 'images']);
}

/** 선택 창: 시트의 개체 목록 (이름 바꾸기 · 숨기기/표시 · 순서 · 여러 개 선택) */
let selPaneDlg = null;
function selectionPaneDialog() {
  if (selPaneDlg) { selPaneDlg.close(); selPaneDlg = null; return; }
  const list = el('div', { class: 'selpane-list' });
  const draw = () => {
    const s = sheet();
    const all = OBJECT_PROPS.flatMap((p) => (s[p] ?? []).map((o) => ({ p, o }))).sort((a, b) => (b.o.z ?? 0) - (a.o.z ?? 0));
    list.replaceChildren(...(all.length ? all.map(({ p, o }) => {
      const name = o.name || o.caption || o.title || `${OBJECT_LABEL[p]} ${o.id.slice(-3)}`;
      const row = el('div', { class: `selpane-row${chartSel === o.id || objMulti.has(o.id) ? ' on' : ''}` },
        el('span', { class: 'selpane-kind' }, OBJECT_LABEL[p]),
        el('span', { class: 'selpane-name', title: '두 번 클릭해서 이름 바꾸기' }, name),
        el('button', { class: 'selpane-eye', title: o.hidden ? '표시' : '숨기기', onclick: (e) => { e.stopPropagation(); updateObject(o.id, { hidden: !o.hidden || undefined }); gv.renderObjectsAll(); draw(); } }, o.hidden ? '─' : '👁'));
      row.addEventListener('click', (e) => {
        if ((e.ctrlKey || e.metaKey || e.shiftKey) && chartSel && chartSel !== o.id) { if (objMulti.has(o.id)) objMulti.delete(o.id); else objMulti.add(o.id); } else { objMulti.clear(); chartSel = o.id; }
        gv.renderObjectsAll(); updateSelectionUI(); draw();
      });
      row.querySelector('.selpane-name').addEventListener('dblclick', () => {
        const inp = el('input', { value: name });
        row.querySelector('.selpane-name').replaceWith(inp);
        inp.focus(); inp.select();
        const done = () => { const v = inp.value.trim(); if (v && v !== name) updateObject(o.id, p === 'slicers' ? { name: v } : { name: v }); draw(); };
        inp.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') done(); if (e.key === 'Escape') draw(); });
        inp.addEventListener('blur', done);
      });
      return row;
    }) : [el('div', { class: 'muted', style: { padding: '10px' } }, '이 시트에 개체가 없습니다.')]));
  };
  draw();
  const setAll = (hidden) => { wb.transact(() => { for (const p of OBJECT_PROPS) if ((sheet()[p] ?? []).length) wb.setSheetProp(si, p, sheet()[p].map((o) => ({ ...o, hidden: hidden || undefined }))); }, meta()); gv.renderObjectsAll(); draw(); };
  selPaneDlg = openDialog({
    title: '선택', width: 300, modeless: true, onClose: () => { selPaneDlg = null; },
    body: el('div', { class: 'selpane' },
      el('div', { class: 'selpane-tools' },
        el('button', { class: 'btn', onclick: () => setAll(false) }, '모두 표시'), el('button', { class: 'btn', onclick: () => setAll(true) }, '모두 숨기기'),
        el('button', { class: 'btn', title: '앞으로', onclick: () => { if (chartSel) { arrangeObject(chartSel, 'forward'); draw(); } } }, '▲'),
        el('button', { class: 'btn', title: '뒤로', onclick: () => { if (chartSel) { arrangeObject(chartSel, 'backward'); draw(); } } }, '▼')),
      list, el('div', { class: 'muted', style: { fontSize: '11px' } }, 'Ctrl+클릭: 여러 개 선택 · 두 번 클릭: 이름 바꾸기')),
  });
  selPaneDlg.redraw = draw;
}

/** 채우기 · 윤곽선 · 효과 메뉴 (도형 · 그림 · 텍스트) */
function shapeFillMenu(a) {
  paletteMenu(a, '채우기 없음', (c) => patchObjects({ fill: c, grad: undefined }, ['shapes']));
  setTimeout(() => {
    const m = document.querySelector('#menuLayer .menu:last-child');
    if (!m) return;
    m.append(el('div', { class: 'menu-sep' }),
      el('button', { class: 'menu-item', onclick: () => { closeMenus(); patchObjects((o) => ({ grad: o.grad ? undefined : { ang: 90 } }), ['shapes']); } }, el('span', { class: 'mi-icon' }), el('span', {}, '그라데이션 (켜기/끄기)')),
      el('button', { class: 'menu-item', onclick: () => { closeMenus(); patchObjects((o) => ({ fillOpacity: o.fillOpacity === 0.5 ? undefined : 0.5 }), ['shapes']); } }, el('span', { class: 'mi-icon' }), el('span', {}, '반투명 50% (켜기/끄기)')));
  }, 0);
}
function shapeOutlineMenu(a) {
  const items = [
    { title: '두께' }, ...[0.25, 0.5, 0.75, 1, 1.5, 2.25, 3, 4.5, 6].map((w) => ({ label: `${w}pt`, icon: `<span style="display:block;width:22px;border-top:${Math.max(1, w * 1.33)}px solid #333;margin-top:6px"></span>`, action: () => patchObjects({ strokeWidth: w }) })),
    { title: '대시' }, ...[['', '실선'], ['dash', '파선'], ['dot', '점선']].map(([v, l]) => ({ label: l, action: () => patchObjects({ dash: v || undefined }) })),
    { title: '화살표 (선)' }, ...[['', '없음'], ['end', '끝 화살표'], ['both', '양쪽 화살표']].map(([v, l]) => ({ label: l, action: () => patchObjects({ arrow: v || undefined }, ['shapes']) })),
    { sep: true }, { label: '윤곽선 색...', icon: 'border', action: () => setTimeout(() => paletteMenu(a, '윤곽선 없음', (c) => patchObjects((o, p) => (p === 'images' ? { border: c || undefined } : { stroke: c }))), 0) },
  ];
  openMenu(a, items, { scroll: true });
}
function shapeEffectsMenu(a) {
  openMenu(a, [
    { title: '그림자' },
    { label: '그림자 없음', action: () => patchObjects({ shadow: undefined }) },
    { label: '바깥쪽 (오른쪽 아래)', action: () => patchObjects({ shadow: { dx: 3, dy: 3, blur: 3 } }) },
    { label: '바깥쪽 (가운데)', action: () => patchObjects({ shadow: { dx: 0, dy: 0, blur: 5, opacity: 0.5 } }) },
    { label: '원근감 (아래)', action: () => patchObjects({ shadow: { dx: 0, dy: 6, blur: 6, opacity: 0.3 } }) },
    { title: '네온' },
    { label: '네온 없음', action: () => patchObjects({ glow: undefined }) },
    ...[['파랑', '#4472c4'], ['주황', '#ed7d31'], ['금색', '#ffc000'], ['녹색', '#70ad47'], ['회색', '#a5a5a5']].map(([n, c]) => ({ label: `네온: ${n}`, icon: `<span style="display:block;width:14px;height:14px;border-radius:50%;box-shadow:0 0 4px 2px ${c};background:#fff"></span>`, action: () => patchObjects({ glow: { color: c, size: 6 } }, ['shapes']) })),
    { title: '부드러운 가장자리' },
    ...[0, 2.5, 5, 10].map((v) => ({ label: v ? `${v}pt` : '없음', action: () => patchObjects({ soft: v || undefined }, ['shapes']) })),
    { title: '그림 스타일' },
    { label: '둥근 모서리', action: () => patchObjects((o) => ({ radius: o.radius ? undefined : 12 }), ['images']) },
    { label: '흰색 테두리 + 그림자', action: () => patchObjects({ border: '#ffffff', borderW: 6, shadow: true }, ['images']) },
    { label: '스타일 없음', action: () => patchObjects({ border: undefined, radius: undefined, shadow: undefined }, ['images']) },
  ], { scroll: true });
}
const SHAPE_QUICK = ['#4472c4', '#ed7d31', '#a5a5a5', '#ffc000', '#5b9bd5', '#70ad47', '#1e293b', '#4f46e5', '#0ea5e9', '#10b981', '#f43f5e', '#8b5cf6'];
function shapeStylesMenu(a) {
  const shade = (hex, t) => { const n = parseInt(hex.slice(1), 16); const ch = [n >> 16, (n >> 8) & 255, n & 255].map((x) => Math.round(t < 0 ? x * (1 + t) : x + (255 - x) * t)); return `#${ch.map((x) => x.toString(16).padStart(2, '0')).join('')}`; };
  const rows = [
    ['채우기', (c) => ({ fill: c, stroke: shade(c, -0.3), color: '#ffffff', grad: undefined, shadow: undefined })],
    ['밝은 채우기', (c) => ({ fill: shade(c, 0.8), stroke: c, color: shade(c, -0.4), grad: undefined, shadow: undefined })],
    ['윤곽선', (c) => ({ fill: '#ffffff', stroke: c, strokeWidth: 1.5, color: c, grad: undefined, shadow: undefined })],
    ['그라데이션', (c) => ({ fill: c, stroke: undefined, color: '#ffffff', grad: { ang: 90 }, shadow: undefined })],
    ['강한 효과', (c) => ({ fill: c, stroke: undefined, color: '#ffffff', grad: { ang: 90 }, shadow: { dx: 0, dy: 4, blur: 5 } })],
  ];
  const grid = el('div', { class: 'qs-grid' }, rows.flatMap(([n, fn]) => SHAPE_QUICK.map((c) => {
    const p = fn(c);
    return el('button', {
      class: 'qs-chip', title: n, style: { background: p.grad ? `linear-gradient(${shade(c, 0.35)}, ${shade(c, -0.15)})` : p.fill, borderColor: p.stroke ?? 'transparent', color: p.color, boxShadow: p.shadow ? '0 2px 4px rgba(0,0,0,.35)' : 'none' },
      onmousedown: (e) => e.preventDefault(), onclick: () => { closeMenus(); patchObjects(p, ['shapes']); },
    }, 'Abc');
  })));
  openMenu(a, [{ title: '테마 스타일' }, { node: grid }]);
}
function textFillMenu(a) { paletteMenu(a, '자동', (c) => patchObjects({ color: c ?? '#000000', paras: undefined }, ['shapes'])); }
function textOutlineMenu(a) {
  openMenu(a, [
    { label: '윤곽선 없음', action: () => patchObjects({ textOutline: undefined }, ['shapes']) },
    ...[['검정', '#000000'], ['흰색', '#ffffff'], ['파랑', '#4472c4'], ['주황', '#ed7d31']].map(([n, c]) => ({ label: `윤곽선: ${n}`, icon: `<span style="font-weight:900;-webkit-text-stroke:1px ${c};color:transparent">A</span>`, action: () => patchObjects({ textOutline: { color: c, w: 0.75 } }, ['shapes']) })),
  ]);
}
function textEffectsMenu(a) {
  openMenu(a, [
    { label: '그림자', action: () => patchObjects((o) => ({ textShadow: !o.textShadow || undefined }), ['shapes']) },
    { title: '네온' }, { label: '네온 없음', action: () => patchObjects({ textGlow: undefined }, ['shapes']) },
    ...[['파랑', '#4472c4'], ['주황', '#ed7d31'], ['금색', '#ffc000'], ['녹색', '#70ad47']].map(([n, c]) => ({ label: `네온: ${n}`, action: () => patchObjects({ textGlow: c }, ['shapes']) })),
  ]);
}
function wordArtMenu(a) {
  const presets = [
    ['채우기: 검정', { color: '#000000', textOutline: undefined, textShadow: undefined, textGlow: undefined }],
    ['채우기: 파랑, 그림자', { color: '#4472c4', textShadow: true, textOutline: undefined, textGlow: undefined }],
    ['윤곽선: 파랑', { color: '#ffffff', textOutline: { color: '#4472c4', w: 1 }, textShadow: undefined, textGlow: undefined }],
    ['채우기: 흰색, 윤곽선: 주황', { color: '#ffffff', textOutline: { color: '#ed7d31', w: 1 }, textShadow: true, textGlow: undefined }],
    ['채우기: 금색, 네온', { color: '#ffc000', textGlow: '#ffc000', textOutline: undefined, textShadow: undefined }],
    ['채우기: 슬레이트, 굵게', { color: '#1e293b', bold: true, textOutline: undefined, textShadow: undefined, textGlow: undefined }],
  ];
  openMenu(a, presets.map(([n, p]) => ({
    label: n, icon: `<span style="font-weight:800;color:${p.color};${p.textOutline ? `-webkit-text-stroke:1px ${p.textOutline.color};` : ''}${p.textShadow ? 'text-shadow:1px 1px 2px rgba(0,0,0,.5);' : ''}${p.textGlow ? `text-shadow:0 0 4px ${p.textGlow};` : ''}">A</span>`,
    action: () => patchObjects({ ...p, paras: undefined, size: undefined }, ['shapes']),
  })));
}
function setObjSize(key, v) {
  const n = Number(String(v).replace(/[^\d.-]/g, ''));
  if (!Number.isFinite(n)) return;
  if (key === 'rot') { patchObjects({ rot: ((n % 360) + 360) % 360 || undefined }, ['shapes', 'images']); return; }
  if (n < 1 || n > 5000) return;
  patchObjects((o) => (key === 'h' ? { h: Math.round(n) } : { w: Math.round(n) }));
}
function objPlacementMenu(a) {
  const f = selectedObjects()[0];
  const cur = f ? placementOf(f.prop, f.obj) : 'twoCell';
  openMenu(a, [
    { title: '개체 위치 (엑셀 [속성])' },
    { label: '위치와 크기 변함', checked: cur === 'twoCell', action: () => patchObjects({ placement: 'twoCell' }) },
    { label: '위치만 변함', checked: cur === 'oneCell', action: () => patchObjects({ placement: 'oneCell' }) },
    { label: '변하지 않음 (위치 고정)', checked: cur === 'absolute', action: () => patchObjects({ placement: 'absolute' }) },
    { sep: true },
    { label: '개체 인쇄', checked: f ? f.obj.noPrint !== true : true, action: () => patchObjects((o) => ({ noPrint: !o.noPrint || undefined })) },
    { label: '잠금 (시트 보호 시 편집 안 됨)', checked: f ? f.obj.locked !== false : true, action: () => patchObjects((o) => ({ locked: o.locked === false ? undefined : false })) },
  ]);
}

/** 선택 영역 확대/축소: 고른 범위가 창에 맞도록 */
function zoomToSelection() {
  const rg = selKind === 'cells' ? sel : usedClip(sel);
  gv.refreshAxes();
  const w = gv.cols.pos(rg.c2 + 1) - gv.cols.pos(rg.c1);
  const h = gv.rows.pos(rg.r2 + 1) - gv.rows.pos(rg.r1);
  const z = Math.max(10, Math.min(400, Math.floor(Math.min((gv.viewW * view.zoom) / 100 / Math.max(1, w), (gv.viewH * view.zoom) / 100 / Math.max(1, h)) * 100)));
  setZoom(z);
  gv.ensureVisible(rg.r1, rg.c1);
}

/** 기호 삽입 (엑셀 [삽입] → [기호]) — 편집 중이면 커서 위치에, 아니면 셀 끝에 */
const SYMBOL_SETS = [
  ['자주 쓰는 기호', '※ ☆ ★ ○ ● ◎ ◇ ◆ □ ■ △ ▲ ▽ ▼ → ← ↑ ↓ ↔ ⇒ ⇔ ∴ ∵ ♠ ♣ ♥ ♦ ✓ ✔ ✗ ✘ ☎ ☏ ♨ ① ② ③ ④ ⑤ ⑥ ⑦ ⑧ ⑨ ⑩'],
  ['통화 · 단위', '₩ $ € £ ¥ ¢ ₹ ₽ % ‰ ℃ ℉ ㎜ ㎝ ㎞ ㎡ ㎢ ㎏ ㎎ ㎖ ℓ ㎾ ㏊ № ㈜ ™ © ®'],
  ['수학', '± × ÷ ≠ ≤ ≥ ≒ ≈ ∞ √ ∑ ∏ ∫ ∂ ∇ ∈ ∉ ⊂ ⊃ ∪ ∩ ∧ ∨ ¬ ∀ ∃ ° ′ ″ π θ α β γ δ ε λ μ σ φ ω Ω Δ Σ'],
  ['괄호 · 문장 부호', '「 」 『 』 【 】 〈 〉 《 》 〔 〕 ‘ ’ “ ” · ‥ … ¶ § † ‡ ♪ ♬'],
  ['원 · 괄호 문자', '㉠ ㉡ ㉢ ㉣ ㉤ ㉥ ㉦ ㉧ ㉮ ㉯ ㉰ ㉱ ⓐ ⓑ ⓒ ⓓ ⓔ ⑴ ⑵ ⑶ ⑷ ⑸ ㈀ ㈁ ㈂ Ⅰ Ⅱ Ⅲ Ⅳ Ⅴ Ⅵ Ⅶ Ⅷ Ⅸ Ⅹ'],
  ['마케팅 · 이모지', '📈 📉 📊 💰 💸 🛒 🎯 🔥 ⭐ ✅ ❌ ⚠️ 🔔 📌 📍 🚀 👍 👎 💡 🆕 🔝 ⬆️ ⬇️ ➡️ 🟢 🟡 🔴 🔵'],
];
function insertSymbolDialog() {
  let picked = null;
  const prev = el('div', { class: 'sym-prev' }, '');
  const body = el('div', { class: 'sym-body' }, SYMBOL_SETS.map(([name, chars]) => [el('div', { class: 'menu-title' }, name),
    el('div', { class: 'sym-grid' }, [...new Intl.Segmenter('ko', { granularity: 'grapheme' }).segment(chars.replace(/ /g, ''))].map(({ segment }) => el('button', {
      class: 'sym-btn', title: segment, onclick: () => { picked = segment; prev.textContent = segment; },
      ondblclick: () => { picked = segment; ins(); },
    }, segment)))]), prev);
  const ins = () => {
    if (!picked) return false;
    if (editing) { const ed = dom.editor; const a = ed.selectionStart ?? ed.value.length; ed.value = ed.value.slice(0, a) + picked + ed.value.slice(ed.selectionEnd ?? a); ed.setSelectionRange(a + picked.length, a + picked.length); return undefined; }
    const cur = wb.getRaw(si, active.r, active.c);
    if (cur.startsWith('=')) { toast('수식 셀에는 편집 중에 넣으세요.'); return false; }
    wb.transact(() => wb.setInput(si, active.r, active.c, cur + picked), meta());
    return undefined;
  };
  openDialog({ title: '기호', width: 520, body, buttons: [{ label: '삽입', primary: true, action: ins }, { label: '닫기' }] });
}

/** 수식(방정식) 삽입: 자주 쓰는 수식을 텍스트 상자로 (엑셀 [삽입] → [수식]) */
const EQUATIONS = [
  ['근의 공식', 'x = (−b ± √(b² − 4ac)) / 2a'], ['피타고라스 정리', 'a² + b² = c²'], ['원의 넓이', 'A = πr²'],
  ['이항 정리', '(x + a)ⁿ = Σₖ₌₀ⁿ (ⁿₖ) xᵏ aⁿ⁻ᵏ'], ['ROAS', 'ROAS = 전환 매출 ÷ 광고비 × 100%'], ['CPA', 'CPA = 광고비 ÷ 전환 수'],
  ['CTR', 'CTR = 클릭 수 ÷ 노출 수 × 100%'], ['CVR', 'CVR = 전환 수 ÷ 클릭 수 × 100%'], ['LTV', 'LTV = 평균 객단가 × 구매 빈도 × 고객 유지 기간'],
];
function equationMenu() {
  return [...EQUATIONS.map(([n, eq]) => ({ label: n, action: () => insertEquation(eq) })), { sep: true }, { label: '새 수식 입력...', icon: 'equation', action: () => formDialog('수식', [{ name: 'eq', label: '수식', value: '' }], ({ eq }) => { if (eq.trim()) insertEquation(eq.trim()); }) }];
}
function insertEquation(eq) {
  const p = objectOrigin();
  addObject('shapes', { ...newShape('textbox', { x: p.x, y: p.y, w: Math.max(160, [...eq].length * 11), h: 36 }), text: eq, size: 14, font: 'Cambria Math', stroke: null, italic: true, name: '수식' });
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
  openDialog({ title: rule.errorTitle || 'WIXEL', body, buttons, width: 400 });
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
      el('div', { class: 'muted' }, 'WIXEL 는 보안을 위해 매크로를 실행하지 않습니다. 코드는 읽기 전용으로 표시되며, 파일을 .xlsm 으로 저장하면 매크로가 그대로 보존되어 엑셀에서 실행할 수 있습니다.'),
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

/** 피벗 원본 { cube, rows(필요할 때), si, ref, table } */
function pivotSource(def) {
  return pivotSourceData(wb, def);
}

const pivotItemText = itemText;

/**
 * 피벗을 시트에 씀. def.area(이전 결과 영역)만 지우고 다시 씀
 * def.captureFmt(파일에서 가져온 피벗): 처음 한 번 지금 셀의 서식을 역할별로 기억해 두고(def.cellFmt) 다시 그릴 때도 유지
 */
/** 피벗 기본 서식 위에 파일에서 가져온 서식을 덮음 — 표시 형식은 통째로 바꿈 (소수 자릿수 · 사용자 코드가 섞이지 않게) */
function mergeFmt(base, extra) {
  const out = { ...base };
  if ('numFmt' in extra) { delete out.decimals; delete out.code; }
  return Object.assign(out, extra);
}

const pivotWritten = new Map(); // 피벗마다 마지막으로 그린 칸 서식 (업데이트 시 셀 서식 유지)
// ── 피벗 테이블 조건부 서식 범위 (엑셀: 선택한 셀 / "값" 을 표시하는 모든 셀 / "행 필드"에 대해 "값"을 표시하는 모든 셀) ──
// 규칙에 pivot: { name, scope: 'selection' | 'data' | 'field', value, rowField, colField } 를 두고, 피벗을 다시 그릴 때마다 범위를 새로 구함
const pivotLayouts = new Map();
function pivotLayoutFrom(grid, pm, d, top, left) {
  const base = pm.pageRows + pm.headerRows;
  return {
    top, left, roles: grid.map((row) => row.map((cd) => cd?.role ?? '')),
    rowDepth: grid.map((_, r) => pm.rowItems?.[r - base]?.node?.depth ?? -1),
    values: (d.values ?? []).map((v) => valueName(v)), rows: [...(d.rows ?? [])], cols: [...(d.cols ?? [])],
  };
}
function pivotLayoutOf(tsi, def) {
  const k = `${tsi}:${def.name ?? ''}`;
  if (pivotLayouts.has(k)) return pivotLayouts.get(k);
  const src = pivotSource(def);
  if (!src) return null;
  const res = resolvePivot(src, def);
  const { grid, meta } = computePivot(res, res.def);
  const L = pivotLayoutFrom(grid, meta, res.def, def.top ?? 0, def.left ?? 0);
  pivotLayouts.set(k, L);
  return L;
}
const PIVOT_DATA_ROLE = /^(data|groupData|subData|colSubData|grandColData|grandData):(\d+)$/;
function pivotScopeCells(L, pv) {
  const vi = Math.max(0, L.values.indexOf(pv.value));
  const want = pv.rowField ? L.rows.indexOf(pv.rowField) : L.rows.length - 1;
  const inner = L.rows.length - 1;
  const out = [];
  L.roles.forEach((row, r) => row.forEach((role, c) => {
    const m = PIVOT_DATA_ROLE.exec(role);
    if (!m || +m[2] !== vi) return;
    if (pv.scope === 'field') {
      // 행 필드 수준의 칸만 (부분합 · 총합계 · 열 부분합 제외)
      if (!/^(data|groupData|subData)$/.test(m[1])) return;
      if (want < 0 || want >= inner) { if (m[1] !== 'data') return; } else if (m[1] === 'data' || L.rowDepth[r] !== want) return;
    }
    out.push([L.top + r, L.left + c]);
  }));
  return out;
}
/** 칸 목록 → 직사각형 범위들 (열마다 연속 구간 → 옆 열과 같은 구간이면 합침) */
function cellsToRanges(cells) {
  const byCol = new Map();
  for (const [r, c] of cells) { if (!byCol.has(c)) byCol.set(c, []); byCol.get(c).push(r); }
  const runs = [];
  for (const [c, rs] of [...byCol].sort((a, b) => a[0] - b[0])) {
    rs.sort((a, b) => a - b);
    let s0 = rs[0];
    let p = rs[0];
    for (let i = 1; i <= rs.length; i++) {
      if (i < rs.length && rs[i] === p + 1) { p = rs[i]; continue; }
      runs.push({ r1: s0, r2: p, c1: c, c2: c });
      if (i < rs.length) { s0 = rs[i]; p = rs[i]; }
    }
  }
  const out = [];
  for (const g of runs) {
    const prev = out.find((o) => o.r1 === g.r1 && o.r2 === g.r2 && o.c2 === g.c1 - 1);
    if (prev) prev.c2 = g.c2; else out.push({ ...g });
  }
  return out;
}
function applyPivotScope(rule, L) {
  const rgs = cellsToRanges(pivotScopeCells(L, rule.pivot));
  if (!rgs.length) return false;
  delete rule.more;
  Object.assign(rule, rgs[0], rgs.length > 1 ? { more: rgs.slice(1) } : {});
  return true;
}
function refreshPivotCond(tsi, def) {
  const list = wb.sheets[tsi].cond ?? [];
  if (!list.some((rl) => rl.pivot?.name === def.name && rl.pivot.scope !== 'selection')) return;
  const L = pivotLayouts.get(`${tsi}:${def.name ?? ''}`);
  if (!L) return;
  let changed = false;
  const next = list.map((rl) => {
    if (rl.pivot?.name !== def.name || rl.pivot.scope === 'selection') return rl;
    const nr = structuredClone(rl);
    if (!applyPivotScope(nr, L)) return rl;
    if (JSON.stringify(nr) !== JSON.stringify(rl)) changed = true;
    return nr;
  });
  if (changed) wb.setSheetProp(tsi, 'cond', next);
}
/** 규칙이 걸린 피벗 (규칙에 적힌 피벗 → 규칙 범위 → 지금 선택 순) */
function pivotOfRule(rule, tsi = si) {
  if (rule.pivot) {
    const e = allPivots().find((x) => x.si === tsi && (x.def.name ?? '') === rule.pivot.name);
    if (e) return e.def;
  }
  if (rule.r1 !== undefined) {
    for (const { def } of pivotDefs(tsi)) {
      const a = def.area;
      if (a && rule.r1 <= a.r2 && rule.r2 >= a.r1 && rule.c1 <= a.c2 && rule.c2 >= a.c1) return def;
    }
    return null;
  }
  return pivotAreaHit(usedClip(sel));
}
/** 규칙의 적용 대상 기본값: 값 필드 · 행 필드 (규칙에 있으면 그것, 없으면 규칙 범위 첫 칸 / 지금 칸 기준) */
function pivotScopeDefaults(rule, L, at = null) {
  let value = rule.pivot?.value ?? L.values[0];
  let rowField = rule.pivot?.rowField ?? L.rows[L.rows.length - 1] ?? null;
  if (!rule.pivot) {
    const p = at ?? (rule.r1 !== undefined ? { r: rule.r1, c: rule.c1 } : { r: active.r, c: active.c });
    const m = PIVOT_DATA_ROLE.exec(L.roles[p.r - L.top]?.[p.c - L.left] ?? '');
    if (m) value = L.values[+m[2]] ?? value;
    const dep = L.rowDepth[p.r - L.top];
    if (dep >= 0 && L.rows[dep]) rowField = L.rows[dep];
  }
  return { value, rowField, colField: L.cols[L.cols.length - 1] ?? null };
}
const pivotScopeLabels = (d) => ({
  selection: '선택한 셀',
  data: `"${d.value}" 값을 표시하는 모든 셀`,
  field: d.rowField ? `"${d.rowField}"${d.colField ? ` 및 "${d.colField}"` : ''}에 대해 "${d.value}" 값을 표시하는 모든 셀` : `"${d.value}" 값을 표시하는 모든 셀 (행 필드 없음)`,
});
/** 규칙 객체에 적용 대상을 정하고 범위를 다시 구함 (scope: selection · data · field) */
function setRulePivotScope(rule, scope, tsi = si) {
  const def = pivotOfRule(rule, tsi);
  if (!def) return false;
  const L = pivotLayoutOf(tsi, def);
  if (!L?.values.length) return false;
  const d = pivotScopeDefaults(rule, L);
  rule.pivot = { name: def.name ?? '', scope, value: d.value, ...(d.rowField ? { rowField: d.rowField } : {}), ...(d.colField ? { colField: d.colField } : {}) };
  if (scope !== 'selection') applyPivotScope(rule, L);
  return true;
}

// 엑셀의 '서식 옵션' 단추: 피벗 안에 규칙을 넣은 직후 선택 영역 오른쪽 아래에 떠서 적용 대상 3가지를 고름
let cfSmartTag = null;
function hideCfSmartTag() { cfSmartTag?.remove(); cfSmartTag = null; }
function showCfSmartTag(rule) {
  hideCfSmartTag();
  const shownAt = `${si}:${sel.r1},${sel.c1},${sel.r2},${sel.c2}`;
  const def = pivotOfRule(rule);
  const L = def ? pivotLayoutOf(si, def) : null;
  if (!L?.values.length) return;
  const rect = dom.view.getBoundingClientRect();
  const sr = gv.screenRect(usedClip(sel));
  const labels = pivotScopeLabels(pivotScopeDefaults(rule, L));
  const btn = el('button', { class: 'cf-smarttag', title: '서식 옵션 — 규칙 적용 대상 (피벗 테이블)' }, el('span', { html: ICONS.condFormat ?? '' }), '▾');
  btn.style.left = `${Math.min(rect.right - 40, rect.left + sr.x + sr.w + 2)}px`;
  btn.style.top = `${Math.min(rect.bottom - 30, rect.top + sr.y + sr.h + 2)}px`;
  btn.addEventListener('mousedown', (e) => e.preventDefault());
  btn.addEventListener('click', () => {
    const cur = rule.pivot?.scope ?? 'selection';
    openMenu(btn, [
      { title: '서식 규칙 적용 대상' },
      ...['selection', 'data', 'field'].map((k) => ({
        label: labels[k], checked: cur === k,
        action: () => {
          const list = sheet().cond;
          const i = list.indexOf(rule);
          if (i < 0) { hideCfSmartTag(); return; }
          const next = list.map((x) => structuredClone(x));
          setRulePivotScope(next[i], k);
          wb.transact(() => wb.setSheetProp(si, 'cond', next), meta());
          rule = sheet().cond[i];
          toast(`규칙 적용 대상: ${labels[k]}`);
          gv.renderAll();
        },
      })),
    ]);
  });
  document.body.append(btn);
  btn._at = shownAt;
  cfSmartTag = btn;
}
/** 규칙 관리자의 '적용 대상' 칸: 피벗 규칙이면 3가지 중 고르기 */
function pivotScopeSelect(rl, tsi, onChange) {
  const def = pivotOfRule(rl, tsi);
  const L = def ? pivotLayoutOf(tsi, def) : null;
  if (!L?.values.length) return null;
  const labels = pivotScopeLabels(pivotScopeDefaults(rl, L));
  const sel2 = el('select', { class: 'cf-pscope', title: '피벗 테이블 규칙 적용 대상' },
    ['selection', 'data', 'field'].map((k) => el('option', { value: k, selected: (rl.pivot?.scope ?? 'selection') === k }, labels[k])));
  sel2.addEventListener('mousedown', (e) => e.stopPropagation());
  sel2.addEventListener('change', () => { setRulePivotScope(rl, sel2.value, tsi); onChange?.(); });
  return el('div', { class: 'cf-pscope-row' }, el('span', { class: 'muted' }, '피벗: '), sel2);
}

/** 조건부 서식 규칙 추가 (피벗 안이면 피벗에 묶고 '서식 옵션' 단추) */
function addCondRuleHere(rule) {
  const r = { ...usedClip(sel), ...rule };
  const inPivot = !r.pivot && pivotOfRule(r);
  if (inPivot) setRulePivotScope(r, 'selection');
  wb.transact(() => wb.addCondRule(si, r), meta());
  if (inPivot || r.pivot) setTimeout(() => showCfSmartTag(sheet().cond[0]), 0);
}

/** 규칙 편집기의 '규칙 적용 대상' (선택 영역이 피벗 안이거나 규칙이 피벗에 묶여 있을 때) */
function pivotScopeUi(rule) {
  const def = pivotOfRule(rule);
  if (!def) return null;
  const L = pivotLayoutOf(si, def);
  if (!L || !L.values.length) return null;
  // 값 필드 · 행 필드 (규칙에 적힌 것 → 규칙 범위 첫 칸 → 지금 칸)
  const dflt = pivotScopeDefaults(rule, L, rule.r1 === undefined ? { r: active.r, c: active.c } : null);
  let value = dflt.value;
  const rowField = dflt.rowField;
  const colField = dflt.colField;
  const cur = rule.pivot?.scope ?? 'selection';
  const name = `pvs${Date.now()}`;
  const opt = (v, label) => { const r = el('input', { type: 'radio', name, value: v, checked: cur === v }); return [r, el('label', { class: 'fc-check' }, r, label)]; };
  const vSel = el('select', {}, L.values.map((v) => el('option', { value: v, selected: v === value }, v)));
  const [r1, l1] = opt('selection', '선택한 셀');
  const [r2, l2] = opt('data', `"${value}" 값을 표시하는 모든 셀`);
  const [r3, l3] = opt('field', rowField ? `"${rowField}"${colField ? ` 및 "${colField}"` : ''}에 대해 "${value}" 값을 표시하는 모든 셀` : `"${value}" 값을 표시하는 모든 셀 (행 필드 없음)`);
  vSel.addEventListener('change', () => {
    value = vSel.value;
    l2.lastChild.textContent = `"${value}" 값을 표시하는 모든 셀`;
    l3.lastChild.textContent = `"${rowField ?? ''}"${colField ? ` 및 "${colField}"` : ''}에 대해 "${value}" 값을 표시하는 모든 셀`;
  });
  const node = el('div', { class: 'cf-pivot-scope' },
    el('div', { class: 'fc-title' }, '규칙 적용 대상'),
    el('label', { class: 'fc-check' }, '값 필드: ', vSel), l1, l2, l3,
    el('div', { class: 'muted' }, '값 필드 전체를 고르면 피벗 테이블의 행이 늘거나 줄어도 규칙 범위가 따라갑니다.'));
  return {
    node,
    apply: (out) => {
      const scope = r2.checked ? 'data' : r3.checked ? 'field' : 'selection';
      out.pivot = { name: def.name ?? '', scope, value, ...(rowField ? { rowField } : {}), ...(colField ? { colField } : {}) };
      if (scope !== 'selection') applyPivotScope(out, L);
    },
  };
}

function writePivot(targetSi, def, { autofit = true } = {}) {
  delete def.needsRender; // 예제 등에서 처음 한 번 그리라는 표시
  const src = pivotSource(def);
  if (!src) return false;
  const res = resolvePivot(src, def);
  const d = res.def;
  const { grid, meta: pm } = computePivot(res, d);
  const t = wb.sheets[targetSi];
  const top = def.top ?? 0;
  const left = def.left ?? 0;
  const colsN = Math.max(0, ...grid.map((row) => row.length));
  // 날짜 · 시간 등 숫자 항목 레이블은 원본 열의 표시 형식으로 (엑셀과 같음)
  if (src.ref && src.si !== undefined) {
    const hdr = (src.cube?.header ?? []).map((h) => String(h ?? '').toLowerCase());
    const fmtOf = new Map();
    const styleFor = (field) => {
      if (!field || d.groups?.[field]) return null;
      if (fmtOf.has(field)) return fmtOf.get(field);
      const i = hdr.indexOf(String(field).toLowerCase());
      const st = i >= 0 ? wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + i) : null;
      const f = st && st.numFmt && st.numFmt !== 'general' && st.numFmt !== 'text' ? { numFmt: st.numFmt, ...(st.code ? { code: st.code } : {}), ...(st.decimals !== undefined ? { decimals: st.decimals } : {}) } : null;
      fmtOf.set(field, f);
      return f;
    };
    for (const row of grid) for (const cd of row) {
      const m = /^(rowItem|rowGroup|colItem):(\d+)$/.exec(cd?.role ?? '');
      if (!m || cd.raw === '' || !/^-?\d+(\.\d+)?$/.test(String(cd.raw))) continue;
      const f = styleFor(m[1] === 'colItem' ? d.cols[+m[2]] : d.rows[+m[2]]);
      if (f) cd.style = { ...(cd.style ?? {}), ...f };
    }
  }
  // 파일의 서식 기억 (엑셀이 셀에 저장한 서식: 표시 형식 · 맞춤 · 사용자가 바꾼 색 등)
  if (def.captureFmt) {
    const fmt = { ...(def.cellFmt ?? {}) };
    // 역할마다 가장 많이 쓰인 서식 (첫 칸만 보면 강조한 한 행의 굵게 등이 본문 전체로 번짐)
    const votes = new Map();
    grid.forEach((row, r) => row.forEach((cd, c) => {
      if (!cd?.role || cd.role === 'empty' || def.cellFmt?.[cd.role]) return;
      const fc = t.cells.getRC(top + r, left + c);
      const own = fc?.style;
      let f = null;
      if (own && Object.keys(own).length) f = { ...(cd.style?.numFmt ? { numFmt: 'general' } : {}), ...own };
      // 파일 셀이 "일반" 형식이면 피벗 기본 표시 형식을 쓰지 않음 (엑셀 화면과 같게)
      else if (fc && fc.raw !== '' && cd.style?.numFmt) f = { numFmt: 'general' };
      else if (!fc) return;
      const key = JSON.stringify(f);
      let m = votes.get(cd.role);
      if (!m) votes.set(cd.role, (m = new Map()));
      const e = m.get(key);
      if (e) e.n++; else m.set(key, { n: 1, f });
    }));
    for (const [role, m] of votes) {
      let best = null;
      for (const e of m.values()) if (!best || e.n > best.n) best = e;
      if (best?.f) fmt[role] = best.f;
    }
    def.cellFmt = fmt;
    delete def.captureFmt;
    autofit = false;
  }
  if (def.autofit === false) autofit = false;
  // 업데이트 시 셀 서식 유지: 지난번에 그린 서식과 다른 칸(사용자가 바꾼 서식)은 역할별로 기억해 다시 적용
  const wkey = `${targetSi}:${def.name ?? ''}:${def.top ?? 0},${def.left ?? 0}`;
  const written = pivotWritten.get(wkey);
  if (def.preserveFormat !== false && written) {
    const fmt = { ...(def.cellFmt ?? {}) };
    let changedFmt = false;
    grid.forEach((row, r) => row.forEach((cd, c) => {
      if (!cd?.role || cd.role === 'empty') return;
      const k = `${top + r},${left + c}`;
      const was = written.get(k);
      const now = t.cells.getRC(top + r, left + c)?.style;
      if (was === undefined || !now) return;
      const prev = JSON.parse(was || '{}');
      const diff = {};
      for (const key of Object.keys(now)) if (JSON.stringify(now[key]) !== JSON.stringify(prev[key])) diff[key] = now[key];
      if (Object.keys(diff).length) { fmt[cd.role] = { ...(fmt[cd.role] ?? {}), ...diff }; changedFmt = true; }
    }));
    if (changedFmt) def.cellFmt = fmt;
  }
  const cellFmt = def.preserveFormat === false ? {} : def.cellFmt ?? {};
  const a = def.area;
  // 레이블 셀 병합: 이전 영역의 병합은 먼저 풂
  if (a) for (const m of wb.mergesIn(targetSi, a.r1, a.c1, a.r2, a.c2)) if (m.r1 >= a.r1 && m.c1 >= a.c1 && m.r2 <= a.r2 && m.c2 <= a.c2) wb.unmerge(targetSi, m.r1, m.c1, m.r2, m.c2);
  // 처음 그리는 피벗(이전 영역 없음)은 아무것도 지우지 않음 — 같은 시트의 다른 피벗 · 내용을 보존
  const inOld = (r, c) => (a ? r >= a.r1 && r <= a.r2 && c >= a.c1 && c <= a.c2 : false);
  const nr2 = top + grid.length - 1;
  const nc2 = left + Math.max(0, colsN - 1);
  if (a) for (const k of [...t.cells.keys()]) {
    const i = k.indexOf(',');
    const r = +k.slice(0, i);
    const c = +k.slice(i + 1);
    // 새 결과 영역 안의 셀은 아래에서 덮어쓰므로 지우지 않음 (실행 취소 기록이 두 번 생기지 않게)
    if (inOld(r, c) && !(r >= top && r <= nr2 && c >= left && c <= nc2)) wb.setCellData(targetSi, r, c, null);
  }
  def.area = { r1: top, c1: left, r2: nr2, c2: nc2 };
  pivotLayouts.set(`${targetSi}:${def.name ?? ''}`, pivotLayoutFrom(grid, pm, d, top, left));
  refreshPivotCond(targetSi, def);
  const btns = [];
  grid.forEach((row, r) => {
    for (let c = 0; c < colsN; c++) {
      const cd = row[c];
      const rr = top + r;
      const cc = left + c;
      if (!cd || (!cd.raw && !cd.style && !cd.image)) { if (t.cells.has(`${rr},${cc}`)) wb.setCellData(targetSi, rr, cc, null); continue; }
      const extra = cellFmt[cd.role];
      wb.setCellData(targetSi, rr, cc, { raw: cd.raw, style: extra ? mergeFmt(cd.style, extra) : cd.style, ...(cd.image ? { image: cd.image } : {}) });
      // 필터 단추: 행 레이블 머리글, 열 레이블 머리글, 보고서 필터 값
      if (cd.role === 'rowHead:0' && d.rows.length) btns.push({ r: rr, c: cc, kind: 'rows' });
      else if (/^rowHead:\d+$/.test(cd.role) && d.layout !== 'compact' && d.rows[+cd.role.split(':')[1]]) btns.push({ r: rr, c: cc, kind: 'rows', field: d.rows[+cd.role.split(':')[1]] });
      else if (cd.role === 'colHead' && cd.raw) btns.push({ r: rr, c: cc, kind: 'cols' });
      else if (cd.role === 'pageValue') btns.push({ r: rr, c: cc, kind: 'page', field: d.pages[r] });
      if (cd.toggle) btns.push({ r: rr, c: cc, kind: 'toggle', field: cd.toggle.field, item: cd.toggle.item, collapsed: cd.toggle.collapsed });
    }
  });
  def.buttons = btns;
  // 레이블이 있는 셀 병합 및 가운데 맞춤 (테이블 · 개요 형식): 바깥 행 필드 항목을 그 그룹 행만큼 세로 병합
  if (def.mergeLabels && d.layout !== 'compact' && d.rows.length > 1 && !pm.empty) {
    for (let c = 0; c < d.rows.length - 1; c++) {
      let start = -1;
      const flush = (end) => {
        if (start >= 0 && end > start) {
          wb.merge(targetSi, top + start, left + c, top + end, left + c);
          wb.setStyle(targetSi, top + start, left + c, { align: 'center', valign: 'middle' });
        }
        start = -1;
      };
      for (let r = 0; r < grid.length; r++) {
        const cd = grid[r][c];
        const role = cd?.role ?? '';
        const label = /^(rowItem|rowGroup)/.test(role) && cd.raw !== '' && cd.raw !== undefined;
        const blankCont = start >= 0 && (!cd || cd.raw === '' || cd.raw === undefined) && !/^(grand|rowSub)/.test(role) && !/^(grand|rowSub)/.test(grid[r][c + 1]?.role ?? '');
        if (label) { flush(r - 1); start = r; } else if (!blankCont) flush(r - 1);
      }
      flush(grid.length - 1);
    }
  }
  // 이번에 그린 서식 기억 (다음 업데이트에서 사용자가 바꾼 칸을 찾음)
  const wm = new Map();
  grid.forEach((row, r) => row.forEach((cd, c) => { if (cd?.role) wm.set(`${top + r},${left + c}`, JSON.stringify(t.cells.getRC(top + r, left + c)?.style ?? {})); }));
  pivotWritten.set(wkey, wm);
  if (autofit && !pm.empty) {
    for (let c = 0; c < colsN; c++) {
      let w = 0;
      grid.forEach((row, r) => {
        if (row[c]?.raw) w = Math.max(w, measureText(displayText(top + r, left + c, targetSi), wb.styleAt(targetSi, top + r, left + c)) + 12 + (row[c].style?.indent ?? 0) * 12 + (btns.some((b) => b.r === top + r && b.c === left + c) ? 18 : 0));
      });
      if (w > (wb.sheets[targetSi].colWidths[left + c] ?? wb.sheets[targetSi].defColW ?? DEFAULT_COL_WIDTH)) wb.setColWidth(targetSi, left + c, Math.min(300, Math.ceil(w)));
    }
  }
  return true;
}

/** 시트의 모든 피벗 정의 [{ def, prop, index, si }] */
function pivotDefs(s = si) {
  const sh = wb.sheets[s];
  const list = [];
  if (!sh) return list;
  if (sh.pivot) list.push({ def: sh.pivot, prop: 'pivot', index: -1, si: s });
  (sh.pivotsExtra ?? []).forEach((d, i) => { if (d) list.push({ def: d, prop: 'pivotsExtra', index: i, si: s }); });
  return list;
}

/** 통합 문서의 모든 피벗 */
function allPivots() {
  return wb.sheets.flatMap((_, i) => pivotDefs(i));
}

/** 피벗 이름 (없으면 시트 안 순서로) */
const pivotNameOf = (entry) => entry.def.name ?? `피벗 테이블${entry.index + 2}`;

/** 새 피벗 이름: 피벗 테이블1, 2, … */
function nextPivotName() {
  const used = new Set(allPivots().map((e) => String(e.def.name ?? '').toLowerCase()));
  let n = 1;
  while (used.has(`피벗 테이블${n}`)) n++;
  return `피벗 테이블${n}`;
}

/** 시트 이름(없으면 host)과 피벗 이름으로 찾기 */
function findPivotEntry(sheetName, name, hostSi = si) {
  const s = sheetName ? wb.sheetIndexByName(sheetName) : hostSi;
  const list = s >= 0 ? pivotDefs(s) : [];
  if (name) return list.find((e) => String(e.def.name ?? '').toLowerCase() === String(name).toLowerCase()) ?? null;
  return list[0] ?? null;
}

/** 활성 셀이 들어 있는 피벗 (없으면 null) */
function pivotHere() {
  const list = pivotDefs();
  return list.find(({ def: d }) => d.area && active.r >= d.area.r1 && active.r <= d.area.r2 && active.c >= d.area.c1 && active.c <= d.area.c2)
    ?? (list.length === 1 && !list[0].def.area ? list[0] : null);
}

/** 피벗 정의 저장 (transact 안에서) */
function putPivotDef(entry, def) {
  const s = entry.si ?? si;
  writePivot(s, def);
  if (entry.prop === 'pivot') wb.setSheetProp(s, 'pivot', def);
  else {
    const list = [...(wb.sheets[s].pivotsExtra ?? [])];
    list[entry.index] = def;
    wb.setSheetProp(s, 'pivotsExtra', list);
  }
  entry.def = def;
}

/** 피벗 정의 바꾸기 (다시 계산 + 실행 취소 가능) */
function setPivotDef(entry, def, s = entry.si ?? si) {
  wb.transact(() => putPivotDef({ ...entry, si: s }, def), meta());
  entry.def = def;
}

/** 파일에서 연 피벗: 서식을 기억하고 피벗 스타일로 다시 그림 (실행 취소 기록 없이) */
/** 모든 피벗 · 슬라이서가 쓰는 필드로 요약 캐시 준비 (천만 행도 이후 클릭은 즉시) */
function warmAll() {
  const fields = wb.sheets.flatMap((s) => (s.slicers ?? []).filter((x) => x.source?.kind === 'pivot').map((x) => x.source.field));
  try { warmPivots(wb, allPivots().map((e) => e.def), fields); } catch (err) { console.warn('요약 캐시 준비 실패', err); }
}

function renderImportedPivots() {
  warmAll();
  for (const e of allPivots()) {
    if (!e.def.captureFmt && !e.def.needsRender) continue;
    try { wb.transact(() => writePivot(e.si, e.def)); } catch (err) { console.warn('피벗 다시 그리기 실패', err); }
  }
}

/** 피벗 정의를 새 형식({rows, cols, values …})으로 */
function pivotDefV2(def) {
  if (def.rows) return def;
  const src = pivotSource(def);
  if (!src) return def;
  const n = normalizeDef(def, headerNames(src));
  const { rowField, colField, valueField, agg, fieldNames, ...rest } = def;
  return { ...rest, rows: n.rows, cols: n.cols, values: n.values };
}

function pivotDialog(tableName = null) {
  const inTable = tableName ? findTable(wb, tableName)?.t : tableHere();
  const rg = inTable ? { r1: inTable.r1, c1: inTable.c1, r2: dataBottom(inTable), c2: inTable.c2 } : dataRange();
  if (rg.r2 <= rg.r1) { alertDialog('피벗 테이블', '머리글 행과 데이터가 있는 범위를 선택하세요.'); return; }
  const srcName = sheet().name;
  const headers = [];
  for (let c = rg.c1; c <= rg.c2; c++) headers.push(displayText(rg.r1, c) || `열${c - rg.c1 + 1}`);
  const where = el('select', {}, el('option', { value: 'new' }, '새 워크시트'), el('option', { value: 'here' }, '기존 워크시트'));
  const loc = el('input', { type: 'text', value: '', placeholder: '예: H3', disabled: true });
  where.addEventListener('change', () => { loc.disabled = where.value !== 'here'; if (!loc.disabled && !loc.value) loc.value = cellName(rg.r1, rg.c2 + 2); });
  const rangeIn = el('input', { type: 'text', value: inTable ? inTable.name : `${quoteSheetName(srcName)}!${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}` });
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '10px' } },
    el('div', { class: 'muted' }, '표 이름(예: 표1)을 원본으로 쓰면 표에 추가된 데이터까지 새로 고칠 때 자동으로 포함됩니다.'),
    el('label', {}, el('span', {}, '표/범위'), rangeIn),
    el('label', {}, el('span', {}, '넣을 위치'), where),
    el('label', {}, el('span', {}, '위치'), loc));
  openDialog({
    title: '피벗 테이블 만들기', width: 420, body,
    buttons: [{
      label: '확인', primary: true, action: () => {
        const text = rangeIn.value.trim().replace(/^=/, '');
        const tf = findTable(wb, text);
        let def;
        // 엑셀처럼 빈 피벗 테이블을 만들고 [피벗 테이블 필드] 창에서 필터 · 열 · 행 · 값을 고름
        const fields = { name: nextPivotName(), rows: [], cols: [], values: [], ...pivotDefaultsDef() };
        if (tf) def = { table: tf.t.name, source: wb.sheets[tf.si].name, range: { r1: tf.t.r1, c1: tf.t.c1, r2: dataBottom(tf.t), c2: tf.t.c2 }, ...fields };
        else {
          const bang = text.lastIndexOf('!');
          const src = bang > 0 ? text.slice(0, bang).replace(/^'(.*)'$/, '$1').replace(/''/g, "'") : srcName;
          const prg = parseRangeName((bang > 0 ? text.slice(bang + 1) : text).replace(/\$/g, ''));
          if (!prg || wb.sheetIndexByName(src) < 0) { toast('범위가 올바르지 않습니다. 셀 범위나 표 이름을 입력하세요.'); return false; }
          def = { source: src, range: prg, ...fields };
        }
        let target;
        if (where.value === 'here') {
          const p = parseRangeName(loc.value.replace(/\$/g, ''));
          if (!p) { toast('위치가 올바르지 않습니다.'); return false; }
          def.top = p.r1;
          def.left = p.c1;
          def.area = { r1: p.r1, c1: p.c1, r2: p.r1, c2: p.c1 };
          target = si;
          wb.transact(() => {
            writePivot(target, def);
            if (!sheet().pivot) wb.setSheetProp(target, 'pivot', def);
            else wb.setSheetProp(target, 'pivotsExtra', [...(sheet().pivotsExtra ?? []), def]);
          }, meta());
          selectCell(p.r1, p.c1);
        } else {
          let n = 1;
          while (wb.sheetIndexByName(`피벗${n}`) >= 0) n++;
          def.top = 2;
          def.left = 0;
          target = wb.transact(() => {
            const idx = wb.addSheet(`피벗${n}`, si + 1);
            def.area = null;
            writePivot(idx, def);
            wb.setSheetProp(idx, 'pivot', def);
            return idx;
          }, meta());
          switchSheet(target, false);
          selectCell(def.top, 0);
        }
        pivotPaneOpen = true;
        refreshPivotPane(true);
        return true;
      },
    }, { label: '취소' }],
  });
}

function refreshPivots() {
  let n = 0;
  wb.transact(() => {
    wb.sheets.forEach((s, i) => {
      for (const { def } of pivotDefs(i)) if (writePivot(i, def)) n++;
      if (s.pivot) wb.setSheetProp(i, 'pivot', { ...s.pivot });
    });
  }, meta());
  refreshPivotPane(true);
  toast(n ? `피벗 테이블 ${n}개를 새로 고쳤습니다.` : '새로 고칠 피벗 테이블이 없습니다.');
}

// ── 피벗 테이블 필드 창 ──
let pivotPaneOpen = true;
let pivotPaneKey = '';
let pivotDrag = null;

function pivotPaneEl() {
  let pane = document.getElementById('pivotPane');
  if (!pane) {
    pane = el('aside', { id: 'pivotPane', class: 'pivot-pane' });
    document.getElementById('gridWrap').append(pane);
  }
  return pane;
}

function showPivotPane(on) {
  const wrap = document.getElementById('gridWrap');
  const was = wrap.classList.contains('with-pane');
  wrap.classList.toggle('with-pane', on);
  pivotPaneEl().style.display = on ? '' : 'none';
  if (was !== on) gv.layout();
}

function refreshPivotPane(force = false) {
  const here = chartSel || editing ? null : pivotHere();
  if (!here || !pivotPaneOpen) { showPivotPane(false); pivotPaneKey = ''; return; }
  const key = `${si}:${here.prop}:${here.index}:${JSON.stringify(here.def)}`;
  showPivotPane(true);
  if (!force && key === pivotPaneKey) return;
  pivotPaneKey = key;
  renderPivotPane(here);
}

function renderPivotPane(entry) {
  const pane = pivotPaneEl();
  const src = pivotSource(entry.def);
  if (!src) { pane.replaceChildren(el('div', { class: 'pp-head' }, '피벗 테이블 필드'), el('div', { class: 'muted', style: { padding: '10px' } }, '원본 데이터를 찾을 수 없습니다.')); return; }
  const header = pivotFieldNames(src, entry.def);
  const def = pivotDefV2(entry.def);
  const calcSet = new Set((def.calcFields ?? []).map((c) => c.name.toLowerCase()));
  const areas = { pages: def.pages ?? [], cols: def.cols ?? [], rows: def.rows ?? [], values: def.values ?? [] };
  const used = new Set([...areas.pages, ...areas.cols, ...areas.rows, ...areas.values.map((v) => v.field)].map((x) => x.toLowerCase()));
  const isNum = (f) => {
    if (calcSet.has(f.toLowerCase())) return true;
    const i = header.indexOf(f);
    const col = src.cube.col(i);
    if (!col) return false;
    // 처음 200행에 숫자가 있으면 값 필드로
    for (let r = 0; r < Math.min(col.n, 200); r++) if (typeof col.get(r) === 'number') return true;
    return false;
  };
  const apply = (patch) => {
    const next = { ...def, ...patch };
    // 열 필드가 바뀌면 'Σ 값'의 자리를 그 앞에 있던 필드 기준으로 유지
    if (patch.cols && patch.valuesPos === undefined && def.valuesPos != null) {
      const before = new Set(areas.cols.slice(0, def.valuesPos));
      next.valuesPos = patch.cols.filter((c) => before.has(c)).length;
    }
    if (next.valuesPos != null && next.valuesPos >= (next.cols ?? []).length) delete next.valuesPos;
    if (next.pages && !next.pages.length) delete next.pages;
    setPivotDef(entry, next);
    refreshPivotPane(true);
  };
  const removeField = (name) => ({
    pages: areas.pages.filter((x) => x !== name), cols: areas.cols.filter((x) => x !== name),
    rows: areas.rows.filter((x) => x !== name), values: areas.values.filter((v) => v.field !== name),
  });
  const moveTo = (name, area, at = null, fromValue = null, extra = {}) => {
    if (area !== 'values' && calcSet.has(name.toLowerCase())) { toast('계산 필드는 값 영역에만 넣을 수 있습니다.'); refreshPivotPane(true); return; }
    const base = fromValue !== null ? { ...areas, values: areas.values.filter((_, i) => i !== fromValue) } : area === 'values' ? areas : removeField(name);
    if (area === 'values') {
      const list = [...base.values];
      const v = fromValue !== null ? areas.values[fromValue] : { field: name, agg: isNum(name) ? 'sum' : 'count' };
      list.splice(at ?? list.length, 0, v);
      apply({ ...base, values: list });
      return;
    }
    const list = base[area].filter((x) => x !== name);
    list.splice(at ?? list.length, 0, name);
    // 날짜/시간 열 자동 그룹화 (옵션): 행 · 열에 날짜 필드를 넣으면 연-월로 묶음
    let groups = def.groups;
    if (opts.autoDateGroup && (area === 'rows' || area === 'cols') && !def.groups?.[name] && isDateField(name)) groups = { ...(def.groups ?? {}), [name]: { by: 'yearMonth' } };
    apply({ ...base, [area]: list, ...(groups !== def.groups ? { groups } : {}), ...extra });
  };
  const isDateField = (f) => {
    const i = header.indexOf(f);
    if (i < 0 || !src.ref) return false;
    const st = wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + i);
    return /date/.test(st?.numFmt ?? '') || (st?.numFmt === 'custom' && /[yd]/i.test(st.code ?? '') && !/[#0]/.test(st.code ?? ''));
  };
  // 필터가 걸린 필드 (항목 선택 · 슬라이서 · 레이블/값/상위 10 필터) — 엑셀처럼 깔때기 표시, ▾ 로 바로 수정 · 해제
  const filterKey = (f) => Object.keys(def.filters ?? {}).find((k) => k.toLowerCase() === String(f).toLowerCase());
  const fieldFilterOf = (f) => Object.entries(def.fieldFilters ?? {}).find(([k]) => k.toLowerCase() === String(f).toLowerCase())?.[1];
  const filterInfo = (f) => {
    const k = filterKey(f);
    const ff = fieldFilterOf(f);
    if (!k && !ff) return null;
    const parts = [];
    if (k) {
      const label = pivotItemLabeler(def, f);
      const sel = def.filters[k];
      parts.push(`선택한 항목 ${sel.length}개: ${sel.slice(0, 8).map((x) => label(x) || '(비어 있음)').join(', ')}${sel.length > 8 ? ' …' : ''}`);
    }
    if (ff) parts.push(describeFieldFilter(ff, def.values ?? []));
    return parts.join('\n');
  };
  const areaOf = (f) => (areas.rows.some((x) => x.toLowerCase() === f.toLowerCase()) ? 'rows' : areas.cols.some((x) => x.toLowerCase() === f.toLowerCase()) ? 'cols' : 'page');
  const funnel = () => el('span', { class: 'pp-funnel', html: '<svg viewBox="0 0 16 16" width="13" height="13"><path d="M1.5 2h13l-5 6v5.5l-3 1.5V8z" fill="#217346"/></svg>' });
  const filterBtn = (f) => {
    const b = el('button', { type: 'button', class: 'pp-drop', title: `'${f}' 필터 및 정렬` }, '▾');
    b.addEventListener('mousedown', (e) => e.stopPropagation());
    b.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); openPivotFilterMenu(entry, areaOf(f), f, b); });
    return b;
  };
  // 필드 목록
  const search = el('input', { type: 'search', placeholder: '검색', class: 'pp-search' });
  const fieldList = el('div', { class: 'pp-fields' });
  const renderFields = () => {
    const q = search.value.trim().toLowerCase();
    fieldList.replaceChildren(...header.filter((h) => !q || h.toLowerCase().includes(q)).map((h) => {
      const cb = el('input', { type: 'checkbox', checked: used.has(h.toLowerCase()) });
      cb.addEventListener('change', () => {
        if (cb.checked) moveTo(h, isNum(h) ? 'values' : 'rows');
        else apply(removeField(h));
      });
      const calc = calcSet.has(h.toLowerCase()) ? (def.calcFields ?? []).find((c) => c.name.toLowerCase() === h.toLowerCase()) : null;
      if (!calc) {
        const fi = filterInfo(h);
        const item = el('label', { class: `pp-field${fi ? ' filtered' : ''}`, draggable: 'true', title: fi ? `필터 적용됨\n${fi}` : '' }, cb, el('span', { class: 'pp-fname' }, h), fi ? funnel() : null, filterBtn(h));
        item.addEventListener('dragstart', (e) => { pivotDrag = { name: h }; e.dataTransfer.setData('text/plain', h); });
        return item;
      }
      // 계산 필드: ƒx 표시 · 수식 풍선 도움말 · 수정/삭제 메뉴 (오른쪽 클릭 또는 ▾)
      const menu = el('button', { type: 'button', class: 'pp-menu', title: '계산 필드 메뉴' }, '▾');
      const item = el('label', { class: 'pp-field calc', draggable: 'true', title: `계산 필드\n${h} = ${calc.formula}` },
        cb, el('span', { class: 'fx-badge' }, 'ƒx'), el('span', { class: 'pp-fname' }, h), el('span', { class: 'pp-formula' }, `=${calc.formula}`), menu);
      item.addEventListener('dragstart', (e) => { pivotDrag = { name: h }; e.dataTransfer.setData('text/plain', h); });
      const openCalcMenu = (anchor) => openMenu(anchor, [
        { title: `ƒx ${h}` },
        { label: '계산 필드 수정...', action: () => calcFieldDialog(entry, h) },
        { label: used.has(h.toLowerCase()) ? '값 영역에서 제거' : '값 영역에 추가', action: () => (used.has(h.toLowerCase()) ? apply(removeField(h)) : moveTo(h, 'values')) },
        { sep: true },
        { label: '계산 필드 삭제', action: () => {
          if (!confirm(`계산 필드 '${h}'을(를) 삭제할까요?\n같은 원본을 쓰는 모든 피벗에서 삭제됩니다 (엑셀과 같음).`)) return;
          saveCalcFields(entry, (def.calcFields ?? []).filter((c) => c.name.toLowerCase() !== h.toLowerCase()), { remove: h });
          toast(`계산 필드 '${h}'을(를) 삭제했습니다.`);
        } },
        { label: '새 계산 필드...', action: () => calcFieldDialog(entry, '\u0000new') },
      ]);
      menu.addEventListener('click', (e) => { e.preventDefault(); e.stopPropagation(); openCalcMenu(menu); });
      item.addEventListener('contextmenu', (e) => { e.preventDefault(); openCalcMenu({ x: e.clientX, y: e.clientY }); });
      item.addEventListener('dblclick', (e) => { e.preventDefault(); calcFieldDialog(entry, h); });
      return item;
    }));
  };
  search.addEventListener('input', renderFields);
  renderFields();
  // 영역 4개
  const AREA_LABEL = { pages: '필터', cols: '열', rows: '행', values: 'Σ 값' };
  const areaBox = (area) => {
    const box = el('div', { class: 'pp-area' });
    const items = area === 'values' ? areas.values.map((v, i) => ({ name: v.field, label: valueName(v), i })) : areas[area].map((n, i) => ({ name: n, label: n, i }));
    // 값이 둘 이상이면 열 영역에 'Σ 값' — 엑셀처럼 위아래로 옮겨 지표별(4월·5월 나란히) / 월별 배치를 고름
    const sigmaAt = area === 'cols' && areas.values.length > 1 ? Math.min(areas.cols.length, def.valuesPos ?? areas.cols.length) : -1;
    if (sigmaAt >= 0) {
      const row = el('div', { class: 'pp-item sigma', draggable: 'true', title: '값 필드들의 위치 — 위로 올리면 지표마다 열 항목이 나란히 붙습니다' }, el('span', { class: 'pp-label' }, 'Σ 값'));
      const menuBtn = el('button', { type: 'button', class: 'pp-menu', title: '이동' }, '▾');
      row.append(menuBtn);
      row.addEventListener('dragstart', (e) => { pivotDrag = { sigma: true, from: 'cols', index: sigmaAt }; e.dataTransfer.setData('text/plain', 'Σ'); });
      menuBtn.addEventListener('click', () => openMenu(menuBtn, [
        { label: '위로 이동', disabled: sigmaAt === 0, action: () => apply({ valuesPos: sigmaAt - 1 }) },
        { label: '아래로 이동', disabled: sigmaAt >= areas.cols.length, action: () => apply({ valuesPos: sigmaAt + 1 }) },
        { label: '처음으로 이동', disabled: sigmaAt === 0, action: () => apply({ valuesPos: 0 }) },
        { label: '끝으로 이동', disabled: sigmaAt >= areas.cols.length, action: () => apply({ valuesPos: areas.cols.length }) },
      ]));
      box.append(row);
    }
    items.forEach((it) => {
      const menuBtn = el('button', { type: 'button', class: 'pp-menu', title: '필드 설정' }, '▾');
      const isCalc = calcSet.has(String(it.name).toLowerCase());
      const row = el('div', { class: `pp-item${isCalc ? ' calc' : ''}`, draggable: 'true', title: isCalc ? `계산 필드: =${(def.calcFields ?? []).find((c) => c.name.toLowerCase() === String(it.name).toLowerCase())?.formula ?? ''}` : '' },
        isCalc ? el('span', { class: 'fx-badge' }, 'ƒx') : '', el('span', { class: 'pp-label' }, it.label),
        area !== 'values' && filterInfo(it.name) ? funnel() : null, menuBtn);
      if (area !== 'values' && filterInfo(it.name)) { row.classList.add('filtered'); row.title = `필터 적용됨\n${filterInfo(it.name)}`; }
      row.addEventListener('dragstart', (e) => { pivotDrag = { name: it.name, from: area, index: it.i }; e.dataTransfer.setData('text/plain', it.name); });
      menuBtn.addEventListener('click', () => {
        const list = area === 'values' ? areas.values : areas[area];
        const moves = [
          { label: '위로 이동', disabled: it.i === 0, action: () => { const l = [...list]; [l[it.i - 1], l[it.i]] = [l[it.i], l[it.i - 1]]; apply({ [area]: l }); } },
          { label: '아래로 이동', disabled: it.i === list.length - 1, action: () => { const l = [...list]; [l[it.i + 1], l[it.i]] = [l[it.i], l[it.i + 1]]; apply({ [area]: l }); } },
          { sep: true },
          ...Object.entries(AREA_LABEL).filter(([a]) => a !== area).map(([a, lab]) => ({ label: `${lab}(으)로 이동`, action: () => moveTo(it.name, a, null, area === 'values' ? it.i : null) })),
          { sep: true },
          ...(area !== 'values' ? [{ label: area === 'pages' ? '항목 선택...' : '필터 및 정렬...', action: () => openPivotFilterMenu(entry, area === 'pages' ? 'page' : area, it.name, menuBtn) }] : []),
          { label: '필드 제거', action: () => apply(area === 'values' ? { values: areas.values.filter((_, i) => i !== it.i) } : { [area]: list.filter((_, i) => i !== it.i) }) },
          ...(area === 'values' ? [{ sep: true }, { label: '값 필드 설정...', action: () => valueFieldDialog(it.i) }] : []),
          ...(calcSet.has(String(it.name).toLowerCase()) ? [{ label: '계산 필드 수정...', action: () => calcFieldDialog(entry, it.name) }] : []),
        ];
        openMenu(menuBtn, moves);
      });
      if (sigmaAt >= 0 && it.i >= sigmaAt) box.append(row); else box.insertBefore(row, box.querySelector('.pp-item.sigma'));
    });
    box.addEventListener('dragover', (e) => { e.preventDefault(); box.classList.add('over'); });
    box.addEventListener('dragleave', () => box.classList.remove('over'));
    box.addEventListener('drop', (e) => {
      e.preventDefault();
      box.classList.remove('over');
      const dr = pivotDrag;
      pivotDrag = null;
      if (!dr) return;
      const rows = [...box.querySelectorAll('.pp-item')];
      let at = rows.findIndex((r) => e.clientY < r.getBoundingClientRect().top + r.offsetHeight / 2);
      if (at < 0) at = rows.length;
      // 'Σ 값'이 있는 열 영역: 화면 위치 → 필드 위치 · 값 위치
      const sig = rows.findIndex((r) => r.classList.contains('sigma'));
      if (dr.sigma) {
        if (area !== 'cols') { toast('Σ 값은 열 영역 안에서만 옮길 수 있습니다.'); return; }
        apply({ valuesPos: rows.slice(0, at).filter((r) => !r.classList.contains('sigma')).length });
        return;
      }
      if (area === 'cols' && sig >= 0) {
        const fieldsBefore = rows.slice(0, at).filter((r) => !r.classList.contains('sigma')).length;
        const others = areas.cols.filter((x) => x !== dr.name);
        const oldIdx = areas.cols.indexOf(dr.name);
        const fieldAt = oldIdx >= 0 && oldIdx < fieldsBefore ? fieldsBefore - 1 : fieldsBefore;
        const next = [...others];
        next.splice(fieldAt, 0, dr.name);
        const beforeSig = new Set(areas.cols.slice(0, sig).filter((x) => x !== dr.name));
        if (at <= sig) beforeSig.add(dr.name);
        moveTo(dr.name, 'cols', fieldAt, dr.from === 'values' ? dr.index : null, { valuesPos: next.filter((x) => beforeSig.has(x)).length });
        return;
      }
      if (dr.from === area && area === 'values') {
        const l = [...areas.values];
        const [v] = l.splice(dr.index, 1);
        l.splice(at > dr.index ? at - 1 : at, 0, v);
        apply({ values: l });
      } else moveTo(dr.name, area, at, dr.from === 'values' ? dr.index : null);
    });
    return el('div', { class: 'pp-areawrap' }, el('div', { class: 'pp-areahead' }, AREA_LABEL[area]), box);
  };
  const valueFieldDialog = (i) => {
    const v = areas.values[i];
    const baseFields = [...(areas.rows ?? []), ...(areas.cols ?? [])];
    const curItem = v.basePos ? `\u0000${v.basePos}` : v.baseItem ?? `\u0000prev`;
    let itemsFor = null;
    // 값 표시 형식에 따라 기준 필드 · 항목 칸을 보이고, 기준 필드의 항목으로 목록을 채움
    const onChange = (inp) => {
      const meta = SHOW_AS.find((a) => a.id === inp.showAs.value);
      inp.baseField.parentElement.style.display = meta?.base ? '' : 'none';
      inp.baseItem.parentElement.style.display = meta?.base === 'item' ? '' : 'none';
      if (!meta?.base || itemsFor === inp.baseField.value) return;
      itemsFor = inp.baseField.value;
      const want = inp.baseItem.value || curItem;
      inp.baseItem.replaceChildren(...BASE_POS.map((p) => el('option', { value: `\u0000${p.id}` }, p.label)),
        ...pivotFieldItems(def, itemsFor).map((t) => el('option', { value: t }, t)));
      inp.baseItem.value = [...inp.baseItem.options].some((o) => o.value === want) ? want : '\u0000prev';
    };
    formDialog('값 필드 설정', [
      { name: 'name', label: '사용자 지정 이름', value: valueName(v) },
      { name: 'agg', label: '값 요약 기준', type: 'select', value: v.agg, options: AGGREGATES.map((a) => ({ value: a.id, label: a.label })) },
      { name: 'showAs', label: '값 표시 형식', type: 'select', value: v.showAs ?? 'normal', options: SHOW_AS.map((a) => ({ value: a.id, label: a.label })) },
      { name: 'baseField', label: '기준 필드', type: 'select', value: v.baseField ?? baseFields[0] ?? '', options: baseFields.map((f) => ({ value: f, label: f })) },
      { name: 'baseItem', label: '기준 항목', type: 'select', value: '', options: [] },
      {
        name: 'fmt', label: '표시 형식 (서식 코드)', type: 'select', value: v.numFmt?.code ?? '',
        options: [['', '기본'], ['#,##0', '#,##0 (천 단위)'], ['#,##0.00', '#,##0.00'], ['0.00%', '0.00%'], ['0.0%', '0.0%'], ['"₩"#,##0', '₩ 통화'], ['#,##0"원"', '#,##0원'], ['0.00', '0.00']]
          .concat(v.numFmt?.code && !['#,##0', '#,##0.00', '0.00%', '0.0%', '"₩"#,##0', '#,##0"원"', '0.00'].includes(v.numFmt.code) ? [[v.numFmt.code, v.numFmt.code]] : [])
          .map(([value, label]) => ({ value, label })),
      },
    ], (x) => {
      const nv = { field: v.field, agg: x.agg };
      if (x.showAs !== 'normal') nv.showAs = x.showAs;
      const base = SHOW_AS.find((a) => a.id === x.showAs)?.base;
      if (base && x.baseField) nv.baseField = x.baseField;
      if (base === 'item') {
        if (x.baseItem.startsWith('\u0000')) nv.basePos = x.baseItem.slice(1);
        else nv.baseItem = x.baseItem;
      }
      const auto = valueName({ field: v.field, agg: x.agg });
      if (x.name.trim() && x.name.trim() !== auto && x.name.trim() !== valueName(v)) nv.name = x.name.trim();
      else if (v.name && x.name.trim() === v.name) nv.name = v.name;
      if (x.fmt) nv.numFmt = { ...styleForCode(x.fmt), code: x.fmt };
      const l = [...areas.values];
      l[i] = nv;
      // 새로 고른 서식이 파일에서 가져온 셀 서식보다 우선
      const fmtChanged = (x.fmt || '') !== (v.numFmt?.code ?? '');
      let cellFmt = def.cellFmt;
      if (fmtChanged && cellFmt) {
        cellFmt = Object.fromEntries(Object.entries(cellFmt).map(([role, st]) => {
          if (!new RegExp(`^(data|subData|groupData|grandData|grandColData|colSubData):${i}$`).test(role)) return [role, st];
          const { numFmt, code, decimals, ...rest } = st;
          return [role, rest];
        }));
      }
      apply({ values: l, ...(cellFmt ? { cellFmt } : {}) });
    }, { onChange });
  };
  const layoutSel = el('select', {}, LAYOUTS.map((l) => el('option', { value: l.id, selected: (def.layout ?? 'compact') === l.id }, l.label)));
  layoutSel.addEventListener('change', () => apply({ layout: layoutSel.value }));
  const chk = (label, key) => {
    const cb = el('input', { type: 'checkbox', checked: def[key] !== false });
    cb.addEventListener('change', () => apply({ [key]: cb.checked }));
    return el('label', { class: 'pp-opt' }, cb, label);
  };
  pane.replaceChildren(
    el('div', { class: 'pp-head' }, el('span', {}, '피벗 테이블 필드'), el('button', { type: 'button', title: '닫기', onclick: () => { pivotPaneOpen = false; refreshPivotPane(); } }, '✕')),
    el('div', { class: 'muted pp-hint' }, '보고서에 추가할 필드 선택 (영역으로 끌어 놓기):'),
    search, fieldList,
    el('div', { class: 'pp-tools' },
      el('button', { type: 'button', class: 'btn', onclick: () => calcFieldDialog(entry) }, 'ƒx 계산 필드...'),
      el('button', { type: 'button', class: 'btn', onclick: () => pivotStyleGallery(null, entry) }, '스타일...')),
    el('div', { class: 'pp-grid' }, areaBox('pages'), areaBox('cols'), areaBox('rows'), areaBox('values')),
    el('div', { class: 'pp-opts' }, el('label', { class: 'pp-opt' }, '레이아웃 ', layoutSel), chk('부분합', 'subtotals'), chk('행 총합계', 'grandRows'), chk('열 총합계', 'grandCols')),
  );
}

// ── 피벗 필터 · 정렬 메뉴 (행 레이블 / 열 레이블 / 보고서 필터 단추, 필드 창) ──
function pivotFieldItems(def, field) {
  const src = pivotSource(def);
  if (!src) return [];
  const header = headerNames(src);
  const i = header.findIndex((h) => h.toLowerCase() === String(field).toLowerCase());
  if (i < 0) return [];
  // 열 기반 엔진의 항목 사전 (행을 다시 훑지 않음)
  return sortKeys([...src.cube.col(i).dim().keys]).map((k) => itemText(k));
}
/** 피벗 항목 표시 글자: 숫자 항목은 원본 열의 표시 형식으로 (날짜 46279 → 2026-09-14) — 필터 키는 그대로 */
function pivotItemLabeler(def, field) {
  const src = pivotSource(def);
  if (!src?.ref) return (t) => t;
  const header = headerNames(src);
  const i = header.findIndex((h) => h.toLowerCase() === String(field).toLowerCase());
  if (i < 0) return (t) => t;
  const st = wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + i);
  if (!st?.numFmt || st.numFmt === 'general') return (t) => t;
  return (t) => (t !== '' && Number.isFinite(Number(t)) ? formatValue(Number(t), st).text : t);
}

function openPivotFilterMenu(entry, kind, field, anchorEl) {
  const def0 = pivotDefV2(entry.def);
  const choices = field ? [field] : kind === 'rows' ? def0.rows ?? [] : kind === 'cols' ? def0.cols ?? [] : [];
  if (!choices.length) return;
  let cur = choices[0];
  const box = el('div', { class: 'filter-menu' });
  const upd = (patch) => { closeMenus(); setPivotDef(entry, { ...pivotDefV2(entry.def), ...patch }); refreshPivotPane(true); focusGrid(); };
  const render = () => {
    const def = pivotDefV2(entry.def);
    const items = pivotFieldItems(def, cur);
    const label = pivotItemLabeler(def, cur);
    const sel = def.filters?.[cur] ? new Set(def.filters[cur]) : null;
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
      const cb = el('input', { type: 'checkbox', checked: !sel || sel.has(t), onchange: syncAll });
      checks.set(t, cb);
      cb.dataset.label = label(t);
      list.append(el('label', {}, cb, cb.dataset.label));
    }
    all.addEventListener('change', () => { for (const cb of checks.values()) if (cb.parentElement.style.display !== 'none') cb.checked = all.checked; });
    syncAll();
    const search = el('input', { type: 'search', placeholder: '검색' });
    search.addEventListener('input', () => {
      const q = search.value.trim().toLowerCase();
      for (const [t, cb] of checks) { const show = !q || t.toLowerCase().includes(q) || cb.dataset.label.toLowerCase().includes(q); cb.parentElement.style.display = show ? '' : 'none'; if (q) cb.checked = show; }
      syncAll();
    });
    const ok = () => {
      const chosen = [...checks.entries()].filter(([, cb]) => cb.checked).map(([t]) => t);
      if (!chosen.length) { toast('항목을 하나 이상 선택하세요.'); return; }
      const nf = { ...(def.filters ?? {}) };
      if (chosen.length === items.length) delete nf[cur]; else nf[cur] = chosen;
      upd({ filters: nf });
    };
    const ff = def.fieldFilters?.[cur];
    const fieldSel = choices.length > 1 ? el('select', {}, choices.map((f) => el('option', { value: f, selected: f === cur }, f))) : null;
    fieldSel?.addEventListener('change', () => { cur = fieldSel.value; render(); });
    const act = (label, fn, disabled = false) => el('button', { type: 'button', class: `pf-act${disabled ? ' off' : ''}`, disabled, onclick: () => { closeMenus(); fn(); } }, label);
    const withSort = (s) => ({ sort: { ...(def.sort ?? {}), [cur]: s } });
    const noFilters = () => { const nf = { ...(def.filters ?? {}) }; delete nf[cur]; const nff = { ...(def.fieldFilters ?? {}) }; delete nff[cur]; return { filters: nf, fieldFilters: nff }; };
    box.replaceChildren(
      ...(fieldSel ? [el('label', { class: 'pf-field' }, el('span', {}, '필드 선택:'), fieldSel)] : []),
      ...(kind !== 'page' ? [
        act('텍스트 오름차순 정렬', () => upd(withSort({ dir: 'asc' }))),
        act('텍스트 내림차순 정렬', () => upd(withSort({ dir: 'desc' }))),
        act('기타 정렬 옵션...', () => pivotSortDialog(entry, cur)),
        el('div', { class: 'pf-sep' }),
      ] : []),
      act(`"${cur}"에서 필터 해제`, () => upd(noFilters()), !def.filters?.[cur] && !ff),
      ...(kind !== 'page' ? [
        act(`레이블 필터...${ff?.type === 'label' ? ' ✔' : ''}`, () => pivotFilterDialog(entry, cur, 'label')),
        act(`값 필터...${ff?.type === 'value' ? ' ✔' : ''}`, () => pivotFilterDialog(entry, cur, 'value')),
        act(`상위 10...${ff?.type === 'top' ? ' ✔' : ''}`, () => pivotFilterDialog(entry, cur, 'top')),
        ...(ff ? [el('div', { class: 'muted pf-desc' }, `적용된 필터: ${describeFieldFilter(ff, def.values ?? [])}`)] : []),
      ] : []),
      el('div', { class: 'pf-sep' }),
      search, list,
      el('div', { class: 'filter-foot' },
        el('button', { class: 'btn primary', onclick: ok }, '확인'),
        el('button', { class: 'btn', onclick: () => { closeMenus(); focusGrid(); } }, '취소')),
    );
    setTimeout(() => search.focus());
  };
  render();
  box.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Escape') { closeMenus(); focusGrid(); } });
  const menu = openMenu(anchorEl, [{ node: box }]);
  menu.style.minWidth = '290px';
}

function pivotSortDialog(entry, field) {
  const def = pivotDefV2(entry.def);
  const s = def.sort?.[field] ?? {};
  const values = def.values ?? [];
  formDialog(`정렬 (${field})`, [
    { name: 'dir', label: '정렬 옵션', type: 'select', value: s.dir ? `${s.dir}` : 'manual', options: [{ value: 'manual', label: '수동 (원래 순서)' }, { value: 'asc', label: '오름차순' }, { value: 'desc', label: '내림차순' }] },
    { name: 'by', label: '기준', type: 'select', value: s.by === undefined || s.by === null ? '' : String(s.by), options: [{ value: '', label: field }, ...values.map((v, i) => ({ value: String(i), label: valueName(v) }))] },
  ], (v) => {
    const sort = { ...(def.sort ?? {}) };
    if (v.dir === 'manual') delete sort[field];
    else sort[field] = { dir: v.dir, ...(v.by !== '' ? { by: Number(v.by) } : {}) };
    setPivotDef(entry, { ...def, sort });
  });
}

function pivotFilterDialog(entry, field, type) {
  const def = pivotDefV2(entry.def);
  const values = def.values ?? [];
  const cur = def.fieldFilters?.[field]?.type === type ? def.fieldFilters[field] : {};
  const save = (f) => setPivotDef(entry, { ...def, fieldFilters: { ...(def.fieldFilters ?? {}), [field]: f } });
  const valueOpts = values.map((v, i) => ({ value: String(i), label: valueName(v) }));
  if (type !== 'label' && !values.length) { alertDialog('값 필터', '값 영역에 필드를 먼저 추가하세요.'); return; }
  if (type === 'top') {
    formDialog(`상위 10 필터 (${field})`, [
      { name: 'top', label: '표시', type: 'select', value: cur.top === false ? 'bottom' : 'top', options: [{ value: 'top', label: '상위' }, { value: 'bottom', label: '하위' }] },
      { name: 'n', label: '개수', type: 'number', value: cur.n ?? 10 },
      { name: 'mode', label: '기준', type: 'select', value: cur.mode ?? 'count', options: [{ value: 'count', label: '항목' }, { value: 'percent', label: '%' }, { value: 'sum', label: '합계' }] },
      { name: 'by', label: '값 필드', type: 'select', value: String(cur.by ?? 0), options: valueOpts },
    ], (v) => save({ type: 'top', top: v.top === 'top', n: Number(v.n) || 10, mode: v.mode, by: Number(v.by) }));
    return;
  }
  const ops = type === 'label' ? LABEL_OPS : VALUE_OPS;
  formDialog(`${type === 'label' ? '레이블' : '값'} 필터 (${field})`, [
    ...(type === 'value' ? [{ name: 'by', label: '값 필드', type: 'select', value: String(cur.by ?? 0), options: valueOpts }] : []),
    { name: 'op', label: '조건', type: 'select', value: cur.op ?? (type === 'label' ? 'contains' : 'greaterThan'), options: ops.map(([id, label]) => ({ value: id, label })) },
    { name: 'v1', label: '값', value: cur.v1 ?? '' },
    { name: 'v2', label: '그리고 (범위일 때)', value: cur.v2 ?? '' },
  ], (v) => {
    if (String(v.v1).trim() === '') { toast('값을 입력하세요.'); return false; }
    save({ type, op: v.op, v1: v.v1, v2: v.v2, ...(type === 'value' ? { by: Number(v.by) } : {}) });
    return true;
  }, { note: type === 'label' ? '? 는 한 글자, * 는 여러 글자를 나타냅니다.' : '' });
}

/** 원본이 같은 피벗들 (엑셀은 계산 필드를 피벗 캐시에 두므로 같은 원본의 피벗이 함께 씀) */
function sameSourcePivots(entry) {
  const key = (d) => (d.table ? `t:${String(d.table).toLowerCase()}` : `r:${String(d.source ?? '').toLowerCase()}:${JSON.stringify(d.range ?? null)}`);
  const k = key(entry.def);
  return allPivots().filter((e) => key(e.def) === k);
}

/** 계산 필드의 수식을 따옴표로 (공백 · 기호가 있는 필드 이름) */
const calcQuote = (f) => (/^[\p{L}_][\p{L}\p{N}_.]*$/u.test(f) ? f : `'${f.replace(/'/g, "''")}'`);

/**
 * 계산 필드 목록 저장. shared: 같은 원본의 모든 피벗에 적용 (엑셀과 같음)
 * rename: { from, to } — 값 필드 · 다른 계산 필드 수식의 참조도 바꿈, remove: 지운 이름 (값 영역에서도 뺌)
 * addTo: 이 피벗의 값 영역에 넣을 계산 필드 이름, numFmt: { name, style } — 이 피벗의 그 값 필드 표시 형식
 */
function saveCalcFields(entry, calcs, { shared = true, rename = null, remove = null, addTo = null, numFmt = null } = {}) {
  const targets = shared ? sameSourcePivots(entry) : [entry];
  if (!targets.some((t) => t.si === entry.si && t.prop === entry.prop && t.index === entry.index)) targets.push(entry);
  const low = (x) => String(x).toLowerCase();
  wb.transact(() => {
    for (const t of targets) {
      const cur = pivotDefV2(t.def);
      let values = [...(cur.values ?? [])];
      if (rename) values = values.map((v) => (low(v.field) === low(rename.from) ? { ...v, field: rename.to, ...(v.name && low(v.name.trim()) === low(rename.from) ? { name: `${rename.to} ` } : {}) } : v));
      if (remove) values = values.filter((v) => low(v.field) !== low(remove));
      const mine = t.si === entry.si && t.prop === entry.prop && t.index === entry.index;
      if (mine && addTo && !values.some((v) => low(v.field) === low(addTo))) values.push({ field: addTo, agg: 'sum', name: `${addTo} ` });
      if (mine && numFmt) values = values.map((v) => (low(v.field) === low(numFmt.name) ? { ...v, numFmt: numFmt.style ?? undefined } : v));
      let list = calcs;
      if (!shared && !mine) list = cur.calcFields ?? [];
      putPivotDef(t, { ...cur, calcFields: list.map((c) => ({ name: c.name, formula: c.formula })), values });
    }
  }, meta());
  refreshPivotPane(true);
  return targets.length;
}

/** 계산 필드 미리 보기: 첫 행 필드의 항목별 값 + 총합계 [[이름, 값]] */
function calcPreview(def, calcs, name, maxItems = 7) {
  const src = pivotSource(def);
  if (!src) return [];
  const d2 = {
    ...def, rows: (def.rows ?? []).slice(0, 1), cols: [], pages: def.pages ?? [], values: [{ field: name, agg: 'sum' }], calcFields: calcs,
    fieldFilters: {}, sort: {}, layout: 'compact', subtotals: false, grandRows: true, grandCols: false, style: 'None', cellFmt: undefined, collapsed: {},
  };
  const res = resolvePivot(src, d2);
  const { grid } = computePivot(res, res.def);
  const out = [];
  for (const row of grid) {
    const role = row[0]?.role ?? '';
    if (!/^(rowItem|grandLabel)/.test(role)) continue;
    const isTotal = role.startsWith('grandLabel');
    if (!isTotal && out.length >= maxItems) continue;
    const raw = row[row.length - 1]?.raw ?? '';
    const v = raw === '' ? null : /^-?[\d.]+(e[+-]?\d+)?$/i.test(raw) ? Number(raw) : raw.replace(/^'/, '');
    out.push([isTotal ? '총합계' : String(row[0].raw ?? '').replace(/^'/, ''), v, isTotal]);
  }
  return out;
}

/** 계산 필드를 참조하는 피벗 수 (같은 원본 중 값 영역에 넣은 것) */
const calcUsage = (entry, name) => sameSourcePivots(entry).filter((e) => (pivotDefV2(e.def).values ?? []).some((v) => String(v.field).toLowerCase() === name.toLowerCase())).length;

/** 수식 나열: 계산 필드 목록을 새 시트에 (엑셀의 [수식 나열]) */
function listCalcFormulas(entry = pivotHere()) {
  if (!entry) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const def = pivotDefV2(entry.def);
  const calcs = def.calcFields ?? [];
  let n = 1;
  let name = '계산 필드 목록';
  while (wb.sheetIndexByName(name) >= 0) name = `계산 필드 목록 (${++n})`;
  const idx = wb.transact(() => {
    const at = wb.addSheet(name, si + 1);
    const put = (r, c, v, style) => wb.setCellData(at, r, c, { raw: v, ...(style ? { style } : {}) });
    const head = { bold: true, fill: '#1f3864', color: '#ffffff' };
    put(0, 0, `'${pivotNameOf(entry)} — 계산 필드`, { bold: true, size: 14 });
    put(2, 0, '계산 필드', head); put(2, 1, '수식', head); put(2, 2, '사용하는 피벗 수', head); put(2, 3, '총합계', head);
    calcs.forEach((c, i) => {
      put(3 + i, 0, `'${c.name}`);
      put(3 + i, 1, `'=${c.formula}`);
      put(3 + i, 2, String(calcUsage(entry, c.name)));
      const pv = calcPreview(def, calcs, c.name, 0).find((x) => x[2]);
      if (pv && pv[1] !== null) put(3 + i, 3, typeof pv[1] === 'number' ? String(pv[1]) : `'${pv[1]}`, { numFmt: 'number', decimals: 4 });
    });
    if (!calcs.length) put(3, 0, "'(계산 필드가 없습니다)");
    wb.setColWidth(at, 0, 160); wb.setColWidth(at, 1, 360); wb.setColWidth(at, 2, 120); wb.setColWidth(at, 3, 120);
    return at;
  }, meta());
  switchSheet(idx, false);
  toast(`계산 필드 ${calcs.length}개를 '${name}' 시트에 나열했습니다.`);
}

/**
 * 계산 필드 관리자 (엑셀의 [계산 필드 삽입] + 목록 · 편집 · 복제 · 삭제 · 미리 보기 · 검사)
 * startName: 처음 선택할 계산 필드 (없으면 새로 만들기)
 */
function calcFieldDialog(entry = pivotHere(), startName = null) {
  if (!entry) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const src = pivotSource(pivotDefV2(entry.def));
  if (!src) return;
  const baseFields = headerNames(src);
  const low = (x) => String(x).toLowerCase();
  let calcs = [...(pivotDefV2(entry.def).calcFields ?? [])].map((c) => ({ ...c }));
  let sel = startName ? calcs.findIndex((c) => low(c.name) === low(startName)) : (calcs.length ? 0 : -1);
  let dirty = false;
  const nextName = () => { let k = calcs.length + 1; while (calcs.some((c) => low(c.name) === low(`필드${k}`))) k++; return `필드${k}`; };

  // ── 왼쪽: 목록
  const list = el('div', { class: 'cf-list', role: 'listbox' });
  const countEl = el('span', { class: 'muted' });
  // ── 오른쪽: 편집
  const nameIn = el('input', { type: 'text', class: 'cf-name', spellcheck: 'false' });
  const formulaIn = el('textarea', { class: 'cf-formula fc-code', rows: 3, spellcheck: 'false', placeholder: '예: 비용/클릭수' });
  const status = el('div', { class: 'cf-status' });
  const usage = el('div', { class: 'muted cf-usage' });
  const fieldSearch = el('input', { type: 'search', placeholder: '필드 검색', class: 'cf-fsearch' });
  const fieldBox = el('div', { class: 'cf-fields' });
  const funcBox = el('div', { class: 'cf-funcs' });
  const preview = el('table', { class: 'cf-preview' });
  const curVal = () => (pivotDefV2(entry.def).values ?? []).find((v) => low(v.field) === low(nameIn.value.trim()));
  const FMTS = [['', '기본 (쉼표 스타일)'], ['#,##0', '#,##0'], ['#,##0.00', '#,##0.00'], ['0.00%', '0.00%'], ['0.0%', '0.0%'], ['0%', '0%'], ['"₩"#,##0', '₩ 통화'], ['#,##0"원"', '#,##0원'], ['0.00', '0.00']];
  const fmtSel = el('select', { class: 'cf-fmt' }, FMTS.map(([v, l]) => el('option', { value: v }, l)));
  const shared = el('input', { type: 'checkbox', checked: true });
  const addVals = el('input', { type: 'checkbox', checked: true });
  const btnSave = el('button', { type: 'button', class: 'btn primary' }, '저장');

  const insertText = (t, caretBack = 0) => {
    const a = formulaIn.selectionStart ?? formulaIn.value.length;
    const b = formulaIn.selectionEnd ?? a;
    formulaIn.value = formulaIn.value.slice(0, a) + t + formulaIn.value.slice(b);
    formulaIn.focus();
    const p = a + t.length - caretBack;
    formulaIn.setSelectionRange(p, p);
    check();
  };
  const renderFields = () => {
    const q = low(fieldSearch.value.trim());
    const others = calcs.filter((c, i) => i !== sel);
    const items = [...baseFields.map((f) => ({ f, calc: false })), ...others.map((c) => ({ f: c.name, calc: true, formula: c.formula }))].filter((x) => !q || low(x.f).includes(q));
    fieldBox.replaceChildren(...items.map((x) => el('button', {
      type: 'button', class: `cf-field${x.calc ? ' calc' : ''}`, title: x.calc ? `계산 필드: =${x.formula}` : '두 번 클릭하거나 눌러서 수식에 넣기',
      onclick: () => insertText(calcQuote(x.f)),
    }, x.calc ? el('span', { class: 'fx-badge' }, 'ƒx') : '', x.f)));
  };
  funcBox.replaceChildren(...CALC_FUNCS.map(([fn, help]) => el('button', {
    type: 'button', class: 'cf-func', title: help, onclick: () => insertText(fn === 'PI' || fn === 'ROWS' ? `${fn}()` : `${fn}()`, fn === 'PI' || fn === 'ROWS' ? 0 : 1),
  }, fn)));
  const renderList = () => {
    countEl.textContent = `${calcs.length}개`;
    list.replaceChildren(...calcs.map((c, i) => {
      const n = calcUsage(entry, c.name);
      const item = el('div', { class: `cf-item${i === sel ? ' on' : ''}`, role: 'option', tabindex: 0, title: `=${c.formula}` },
        el('span', { class: 'fx-badge' }, 'ƒx'),
        el('div', { class: 'cf-item-main' }, el('div', { class: 'cf-item-name' }, c.name), el('div', { class: 'cf-item-f' }, `=${c.formula}`)),
        n ? el('span', { class: 'cf-used', title: `같은 원본의 피벗 ${n}개가 값 영역에 사용` }, String(n)) : '');
      item.addEventListener('click', () => select(i));
      item.addEventListener('keydown', (e) => { if (e.key === 'ArrowDown' && i < calcs.length - 1) { select(i + 1); list.children[i + 1]?.focus(); } if (e.key === 'ArrowUp' && i > 0) { select(i - 1); list.children[i - 1]?.focus(); } });
      return item;
    }), ...(sel === -1 ? [el('div', { class: 'cf-item on new' }, el('span', { class: 'fx-badge' }, '+'), el('div', { class: 'cf-item-main' }, el('div', { class: 'cf-item-name' }, nameIn.value || '새 계산 필드'), el('div', { class: 'cf-item-f muted' }, '저장하면 목록에 추가')))] : []));
  };
  const select = (i) => {
    if (dirty && !confirm('저장하지 않은 변경 내용이 있습니다. 버리고 이동할까요?')) return;
    sel = i;
    dirty = false;
    const c = calcs[i];
    nameIn.value = c ? c.name : nextName();
    formulaIn.value = c ? `=${c.formula}` : '=';
    const v = c ? curVal() : null;
    fmtSel.value = v?.numFmt?.code && FMTS.some(([x]) => x === v.numFmt.code) ? v.numFmt.code : '';
    addVals.parentElement.style.display = c ? 'none' : '';
    renderList();
    renderFields();
    check();
  };
  let timer = 0;
  const check = () => {
    clearTimeout(timer);
    timer = setTimeout(checkNow, 120);
  };
  const checkNow = () => {
    const nm = nameIn.value.trim();
    const f = formulaIn.value.trim().replace(/^=/, '').trim();
    const others = calcs.filter((_, i) => i !== sel);
    let err = null;
    if (!nm) err = '이름을 입력하세요.';
    else if (baseFields.some((h) => low(h) === low(nm))) err = '원본에 같은 이름의 필드가 있습니다. 다른 이름을 쓰세요.';
    else if (others.some((c) => low(c.name) === low(nm))) err = '같은 이름의 계산 필드가 있습니다.';
    else if (!f) err = '수식을 입력하세요.';
    const res = err ? null : checkCalc(f, baseFields, others, nm);
    if (res && !res.ok) err = res.error;
    status.className = `cf-status ${err ? 'bad' : 'ok'}`;
    status.textContent = err ? `✕ ${err}` : `✓ 올바른 수식 — 참조: ${res.refs.join(', ') || '(없음)'}`;
    btnSave.disabled = !!err;
    const c = calcs[sel];
    const n = c ? calcUsage(entry, c.name) : 0;
    usage.textContent = c ? (n ? `같은 원본의 피벗 ${n}개가 이 계산 필드를 값 영역에 쓰고 있습니다.` : '아직 어느 피벗의 값 영역에도 없습니다.') : '';
    // 미리 보기 (첫 행 필드의 항목별 값 · 총합계)
    preview.replaceChildren();
    if (err) return;
    let rows = [];
    try { rows = calcPreview(pivotDefV2(entry.def), [...others, { name: nm, formula: f }], nm); } catch { rows = []; }
    const code = fmtSel.value;
    const style = code ? styleForCode(code) : { numFmt: 'number', decimals: 2 };
    const fmt = (v) => (v === null ? '' : typeof v === 'number' ? formatValue(v, style).text : String(v));
    const field = (pivotDefV2(entry.def).rows ?? [])[0] ?? '';
    preview.append(el('tr', {}, el('th', {}, field || '항목'), el('th', {}, nm)),
      ...rows.map(([k, v, tot]) => el('tr', { class: tot ? 'tot' : '' }, el('td', {}, k), el('td', { class: typeof v === 'number' ? 'num' : 'err' }, fmt(v)))));
  };
  const save = () => {
    const nm = nameIn.value.trim();
    const f = formulaIn.value.trim().replace(/^=/, '').trim();
    if (btnSave.disabled) return false;
    const old = calcs[sel];
    const next = calcs.map((c) => ({ ...c }));
    let rename = null;
    if (old) {
      if (low(old.name) !== low(nm) || old.name !== nm) {
        rename = { from: old.name, to: nm };
        next.forEach((c, i) => { if (i !== sel) c.formula = renameCalcRefs(c.formula, old.name, nm); });
      }
      next[sel] = { name: nm, formula: f };
    } else next.push({ name: nm, formula: f });
    const code = fmtSel.value;
    const numFmt = { name: nm, style: code ? { ...styleForCode(code), code } : null };
    const n = saveCalcFields(entry, next, { shared: shared.checked, rename, addTo: !old && addVals.checked ? nm : null, numFmt: old || addVals.checked ? numFmt : null });
    calcs = next;
    sel = calcs.findIndex((c) => c.name === nm);
    dirty = false;
    renderList();
    renderFields();
    check();
    toast(old ? `계산 필드 '${nm}'을(를) 수정했습니다${shared.checked && n > 1 ? ` (피벗 ${n}개)` : ''}.` : `계산 필드 '${nm}'을(를) 추가했습니다.`);
    return true;
  };
  btnSave.addEventListener('click', save);
  const btnNew = el('button', { type: 'button', class: 'btn', onclick: () => select(-1) }, '+ 새로 만들기');
  const btnDup = el('button', {
    type: 'button', class: 'btn', title: '선택한 계산 필드를 복사해 새로 만들기',
    onclick: () => { const c = calcs[sel]; if (!c) return; select(-1); nameIn.value = `${c.name} 복사`; formulaIn.value = `=${c.formula}`; dirty = true; renderList(); check(); },
  }, '복제');
  const btnDel = el('button', {
    type: 'button', class: 'btn danger',
    onclick: () => {
      const c = calcs[sel];
      if (!c) { select(calcs.length ? 0 : -1); return; }
      const users = calcs.filter((x, i) => i !== sel && (() => { try { return checkCalc(x.formula, baseFields, calcs).refs.some((r) => low(r) === low(c.name)); } catch { return false; } })());
      const msg = users.length ? `'${c.name}'을(를) 참조하는 계산 필드가 있습니다: ${users.map((u) => u.name).join(', ')}\n그래도 삭제할까요? (참조하는 필드는 #NAME? 이 됩니다)` : `계산 필드 '${c.name}'을(를) 삭제할까요?${shared.checked ? '\n같은 원본을 쓰는 모든 피벗에서 삭제됩니다 (엑셀과 같음).' : ''}`;
      if (!confirm(msg)) return;
      const next = calcs.filter((_, i) => i !== sel);
      saveCalcFields(entry, next, { shared: shared.checked, remove: c.name });
      calcs = next;
      dirty = false;
      sel = calcs.length ? Math.min(sel, calcs.length - 1) : -1;
      select(sel);
      toast(`계산 필드 '${c.name}'을(를) 삭제했습니다.`);
    },
  }, '삭제');
  let dlg = null;
  const btnList = el('button', { type: 'button', class: 'btn', title: '계산 필드와 수식을 새 시트에 나열 (엑셀의 수식 나열)', onclick: () => { dlg?.close(); listCalcFormulas(entry); } }, '수식 나열');
  nameIn.addEventListener('input', () => { dirty = true; renderList(); check(); });
  formulaIn.addEventListener('input', () => { dirty = true; check(); });
  formulaIn.addEventListener('keydown', (e) => { if (e.key === 'Enter' && (e.ctrlKey || e.metaKey)) { e.preventDefault(); save(); } e.stopPropagation(); });
  fmtSel.addEventListener('change', () => { dirty = true; check(); });
  fieldSearch.addEventListener('input', renderFields);

  const body = el('div', { class: 'cf-dialog' },
    el('div', { class: 'cf-left' },
      el('div', { class: 'cf-head' }, el('b', {}, '계산 필드 '), countEl),
      list,
      el('div', { class: 'cf-actions' }, btnNew, btnDup, btnDel),
      el('div', { class: 'cf-actions' }, btnList)),
    el('div', { class: 'cf-right' },
      el('label', { class: 'cf-row' }, el('span', {}, '이름'), nameIn),
      el('div', { class: 'cf-row top' }, el('span', {}, '수식'), el('div', { class: 'cf-fwrap' }, formulaIn, status)),
      el('div', { class: 'cf-row top' }, el('span', {}, '필드'), el('div', { class: 'cf-fwrap' }, fieldSearch, fieldBox)),
      el('div', { class: 'cf-row top' }, el('span', {}, '함수'), funcBox),
      el('div', { class: 'cf-row' }, el('span', {}, '표시 형식'), fmtSel),
      el('div', { class: 'cf-opts' },
        el('label', {}, shared, ' 같은 원본을 쓰는 모든 피벗에 적용 (엑셀과 같음)'),
        el('label', {}, addVals, ' 이 피벗의 값 영역에 추가'),
        btnSave),
      el('div', { class: 'cf-row top' }, el('span', {}, '미리 보기'), el('div', { class: 'cf-pwrap' }, preview, usage)),
      el('div', { class: 'muted cf-help' }, '계산 필드는 각 필드의 합계에 수식을 적용합니다 (엑셀과 같음). DIVIDE(분자, 분모) 는 0으로 나눠도 오류가 없고, ROWS() 는 그룹의 원본 행 수입니다. Ctrl+Enter 로 저장.')));
  dlg = openDialog({
    title: '계산 필드', width: 860, body,
    buttons: [
      { label: '확인', primary: true, action: () => { if (dirty) return save() ? undefined : false; return undefined; } },
      { label: '닫기' },
    ],
  });
  select(sel);
  setTimeout(() => (sel === -1 ? nameIn : formulaIn).focus(), 30);
}

/** 글꼴 목록 메뉴: 통합 문서에서 쓴 글꼴 + 이 PC의 글꼴 (각 글꼴 모양으로 표시), 검색, 전체 목록 불러오기 */
function fontMenu(anchorEl) {
  const used = new Set();
  for (const s of wb.sheets) for (const cell of s.cells.values()) if (cell.style?.font) used.add(cell.style.font);
  const cur = styleAt(active.r, active.c).font || BASE_FONT.name;
  const search = el('input', { type: 'search', placeholder: '글꼴 검색', class: 'font-search' });
  const list = el('div', { class: 'font-list' });
  const pick = (f) => { closeMenus(); run('fontFamily', f); };
  const item = (f) => el('button', {
    type: 'button', class: `font-item${f === cur ? ' on' : ''}`, title: f, onmousedown: (e) => e.preventDefault(),
    style: { fontFamily: fontStack(f) }, onclick: () => pick(f),
  }, f);
  const render = () => {
    const q = search.value.trim().toLowerCase();
    const match = (f) => !q || f.toLowerCase().includes(q) || (fontAlias(f) ?? '').toLowerCase().includes(q);
    const all = fontList();
    const theme = [BASE_FONT.name].filter(match);
    const usedList = [...used].filter((f) => f !== BASE_FONT.name && match(f));
    list.replaceChildren(
      ...(theme.length ? [el('div', { class: 'menu-title' }, '테마 글꼴'), ...theme.map(item)] : []),
      ...(usedList.length ? [el('div', { class: 'menu-title' }, '이 통합 문서에서 쓴 글꼴'), ...usedList.map(item)] : []),
      el('div', { class: 'menu-title' }, `모든 글꼴 (${all.length})`),
      ...all.filter(match).slice(0, 400).map(item),
      ...(q && !all.some((f) => f.toLowerCase() === q) ? [el('button', { type: 'button', class: 'font-item', onclick: () => pick(search.value.trim()) }, `'${search.value.trim()}' 글꼴 사용`)] : []),
    );
  };
  search.addEventListener('input', render);
  search.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter' && search.value.trim()) pick(search.value.trim());
    if (e.key === 'Escape') { closeMenus(); focusGrid(); }
  });
  const loadBtn = canListLocalFonts()
    ? el('button', {
      type: 'button', class: 'btn font-load',
      onclick: async () => {
        try {
          const l = await loadLocalFonts();
          if (l) { toast(`이 PC의 글꼴 ${l.length}개를 불러왔습니다.`); loadBtn.remove(); render(); }
        } catch { toast('글꼴 목록을 볼 권한이 없습니다. 글꼴 이름을 직접 입력할 수도 있습니다.'); }
      },
    }, '이 PC의 모든 글꼴 불러오기')
    : null;
  render();
  const node = el('div', { class: 'font-menu' }, search, loadBtn, list);
  openMenu(anchorEl, [{ node }]);
  setTimeout(() => search.focus());
}

function pivotStyleOpt(key) {
  const e = pivotHere();
  if (!e) return;
  const d = pivotDefV2(e.def);
  const so = { rowHeaders: true, colHeaders: true, bandRows: false, bandCols: false, ...(d.styleOpts ?? {}) };
  setPivotDef(e, { ...d, styleOpts: { ...so, [key]: !so[key] } });
}

/** 항목 펼치기 · 축소: 필드의 같은 항목은 모두 함께 (엑셀과 같음) */
function togglePivotItem(entry, field, item) {
  const def = pivotDefV2(entry.def);
  const list = new Set(def.collapsed?.[field] ?? []);
  if (list.has(item)) list.delete(item); else list.add(item);
  const collapsed = { ...(def.collapsed ?? {}) };
  if (list.size) collapsed[field] = [...list]; else delete collapsed[field];
  setPivotDef(entry, { ...def, collapsed: Object.keys(collapsed).length ? collapsed : undefined });
}
/** 전체 필드 확장 · 축소: 활성 필드(선택한 셀의 필드, 없으면 첫 행 필드)의 모든 항목 */
function pivotExpandField(expand) {
  const here = pivotHere();
  if (!here) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const def = pivotDefV2(here.def);
  const field = pivotActiveField(here) ?? def.rows?.[0] ?? def.cols?.[0];
  const all = [...(def.rows ?? []), ...(def.cols ?? [])];
  if (!field || all.indexOf(field) === (def.rows ?? []).length - 1 || field === def.cols?.[def.cols.length - 1]) {
    toast('이 필드에는 확장하거나 축소할 하위 수준이 없습니다.');
    return;
  }
  const collapsed = { ...(def.collapsed ?? {}) };
  if (expand) delete collapsed[field]; else collapsed[field] = pivotFieldItems(def, field);
  setPivotDef(here, { ...def, collapsed: Object.keys(collapsed).length ? collapsed : undefined });
}
/** 선택한 셀이 속한 행 · 열 필드 (피벗 결과의 역할로 판단) */
function pivotActiveField(entry) {
  const def = pivotDefV2(entry.def);
  const b = (def.buttons ?? []).find((x) => x.kind === 'toggle' && x.r === active.r && x.c === active.c);
  if (b) return b.field;
  const src = pivotSource(def);
  if (!src) return null;
  const res = resolvePivot(src, def);
  const { grid } = computePivot(res, res.def);
  const cd = grid[active.r - (def.top ?? 0)]?.[active.c - (def.left ?? 0)];
  const m = /^(rowGroup|rowItem|rowSub|colItem):(\d+)$/.exec(cd?.role ?? '');
  if (!m) return null;
  return (m[1] === 'colItem' ? res.def.cols : res.def.rows)[Number(m[2])] ?? null;
}

/** 필드 그룹화: 날짜 → 연 · 분기 · 월 · 연-월 · 일, 숫자 → 구간 (시작 · 간격) */
function pivotGroupDialog() {
  const here = pivotHere();
  if (!here) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const def = pivotDefV2(here.def);
  const field = pivotActiveField(here);
  if (!field) { toast('그룹화할 행 또는 열 필드의 항목을 선택하세요.'); return; }
  const src = pivotSource(def);
  const j = src ? src.cube.header.findIndex((h) => h.toLowerCase() === field.toLowerCase()) : -1;
  if (j < 0) return;
  const col = src.cube.col(j);
  const keys = col.dim().keys.filter((k) => typeof k === 'number');
  if (!keys.length) { alertDialog('그룹화', '선택 항목을 그룹화할 수 없습니다. 날짜나 숫자 필드만 그룹화할 수 있습니다.'); return; }
  const lo = minOf(keys.slice(0, 200000));
  const hi = maxOf(keys.slice(0, 200000));
  const st = src.ref && src.si !== undefined ? wb.styleAt(src.si, Math.min(src.ref.r1 + 1, src.ref.r2), src.ref.c1 + j) : null;
  const isDate = /date|time/.test(st?.numFmt ?? '') || /[yd]/i.test(st?.code ?? '');
  const cur = def.groups?.[field] ?? null;
  const span = hi - lo;
  const niceSize = span > 0 ? 10 ** Math.max(0, Math.floor(Math.log10(span / 10))) : 1;
  formDialog('그룹화', [
    { name: 'by', label: '단위', type: 'select', value: cur?.by ?? (isDate ? 'months' : 'number'), options: GROUP_BY.map((g) => ({ value: g.id, label: g.id === 'number' ? '숫자 구간' : g.label })) },
    { name: 'start', label: '시작', type: 'number', value: cur?.start ?? Math.floor(lo) },
    { name: 'size', label: '단위 (구간 크기)', type: 'number', value: cur?.size ?? niceSize },
  ], (v) => {
    const spec = v.by === 'number' ? { by: 'number', start: Number(v.start) || 0, size: Math.max(1e-9, Number(v.size) || 1) } : { by: v.by };
    // 항목 글자가 바뀌므로 이 필드의 선택 · 순서 · 축소 상태는 지움
    const drop = (obj) => { if (!obj?.[field]) return obj; const o = { ...obj }; delete o[field]; return Object.keys(o).length ? o : undefined; };
    setPivotDef(here, { ...def, groups: { ...(def.groups ?? {}), [field]: spec }, filters: drop(def.filters) ?? {}, order: drop(def.order), collapsed: drop(def.collapsed) });
    refreshPivotPane(true);
  }, {
    note: isDate ? `날짜 필드 '${field}'` : `숫자 필드 '${field}' (${formatGeneral(lo)} ~ ${formatGeneral(hi)})`,
    onChange: (inp) => {
      const num = inp.by.value === 'number';
      inp.start.parentElement.style.display = num ? '' : 'none';
      inp.size.parentElement.style.display = num ? '' : 'none';
    },
  });
}
function pivotUngroup() {
  const here = pivotHere();
  if (!here) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  const def = pivotDefV2(here.def);
  const field = pivotActiveField(here) ?? Object.keys(def.groups ?? {})[0];
  if (!field || !def.groups?.[field]) { toast('그룹화된 필드의 항목을 선택하세요.'); return; }
  const groups = { ...def.groups };
  delete groups[field];
  const drop = (obj) => { if (!obj?.[field]) return obj; const o = { ...obj }; delete o[field]; return Object.keys(o).length ? o : undefined; };
  setPivotDef(here, { ...def, groups: Object.keys(groups).length ? groups : undefined, filters: drop(def.filters) ?? {}, order: drop(def.order), collapsed: drop(def.collapsed) });
  refreshPivotPane(true);
}

function pivotLayoutCmd(patch) {
  const here = pivotHere();
  if (!here) { toast('피벗 테이블 안의 셀을 선택하세요.'); return; }
  setPivotDef(here, { ...pivotDefV2(here.def), ...patch });
  refreshPivotPane(true);
}

// ───────────────────────── 찾기 / 바꾸기 ─────────────────────────
const findState = {
  text: '', replace: '', matchCase: false, whole: false, regex: false, lookIn: 'formulas', byCols: false, scope: 'sheet',
  format: null, replaceFormat: null, options: false,
};
let findCache = null;

const findOpts = (lookIn = findState.lookIn) => ({
  text: findState.text, matchCase: findState.matchCase, whole: findState.whole, regex: findState.regex, lookIn, byCols: findState.byCols, format: findState.format,
  sheets: findState.scope === 'book' ? wb.sheets.map((_, i) => i).filter((i) => !isHiddenSheet(i)) : [si],
});

/** 조건에 맞는 모든 칸 (같은 조건 · 같은 내용이면 다시 계산하지 않음) */
function findAllMatches(lookIn) {
  const opts = findOpts(lookIn);
  const key = `${JSON.stringify(opts)}|${wb.version}|${opts.sheets.map((i) => wb.sheets[i]._ev ?? 0).join(',')}`;
  if (findCache?.key === key) return findCache.list;
  const list = findMatches(wb, opts, { display: (s, r, c) => displayText(r, c, s) });
  findCache = { key, list };
  return list;
}

function goToMatch(m) {
  if (m.si !== si) switchSheet(m.si);
  selectCell(m.r, m.c);
}

/** 다음(또는 이전) 찾기 — 범위가 통합 문서면 다른 시트까지 */
function findNext(text = findState.text, { quiet = false, back = false } = {}) {
  findState.text = text;
  if (!text && !findState.format) return false;
  const list = findAllMatches();
  const order = findOpts().sheets;
  const m = nextMatch(list, { si, r: active.r, c: active.c }, { byCols: findState.byCols, back, order });
  if (!m) {
    if (!quiet) alertDialog('WIXEL', '찾는 항목이 없습니다. 검색 조건을 확인하세요. 범위를 통합 문서로 넓히거나 찾는 위치를 바꿔 볼 수 있습니다.');
    return false;
  }
  goToMatch(m);
  return true;
}

/** 칸 하나 바꾸기 (수식 안의 글자 · 바꿀 서식) */
function replaceIn(r, c, s = si) {
  const cell = wb.getCell(s, r, c);
  let raw = cell?.raw ?? '';
  const quoted = raw.startsWith("'");
  if (quoted) raw = raw.slice(1);
  let done = false;
  if (findState.text) {
    const next = replaceText(raw, findState, findState.replace);
    if (next !== raw) { wb.setInput(s, r, c, next); done = true; }
  }
  if (findState.replaceFormat) {
    wb.setStyle(s, r, c, findState.replaceFormat);
    done = true;
  }
  return done;
}

const FMT_LABEL = { bold: '굵게', italic: '기울임', underline: '밑줄', strike: '취소선', color: '글꼴 색', fill: '채우기', font: '글꼴', size: '크기', align: '가로 맞춤', numFmt: '표시 형식', wrap: '줄 바꿈' };
function fmtPreview(fmt) {
  if (!fmt) return el('div', { class: 'fmt-preview' }, '서식 설정 안 함');
  return el('div', {
    class: 'fmt-preview', title: Object.keys(fmt).map((k) => FMT_LABEL[k] ?? k).join(', '),
    style: { background: fmt.fill ?? '#fff', color: fmt.color ?? '#000', fontWeight: fmt.bold ? '700' : '400', fontStyle: fmt.italic ? 'italic' : 'normal', textDecoration: fmt.underline ? 'underline' : 'none' },
  }, '미리 보기*');
}

/** 찾을(바꿀) 서식 고르기 — 항목마다 "상관없음"이 기본 */
function findFormatDialog(title, cur, done) {
  const tri = (v) => (v === undefined ? '' : v ? '1' : '0');
  const colorField = (name, label) => ({ name, label, type: 'select', value: cur?.[name] ? 'set' : '', options: [{ value: '', label: '상관없음' }, { value: 'set', label: '지정한 색' }] });
  const nf = [['', '상관없음'], ['general', '일반'], ['number', '숫자'], ['comma', '쉼표 스타일'], ['currency', '통화'], ['accounting', '회계'], ['percent', '백분율'], ['date', '날짜'], ['time', '시간'], ['text', '텍스트']];
  const pickers = {};
  const dlg = formDialog(title, [
    { name: 'bold', label: '굵게', type: 'select', value: tri(cur?.bold), options: [{ value: '', label: '상관없음' }, { value: '1', label: '굵게' }, { value: '0', label: '굵게 아님' }] },
    { name: 'italic', label: '기울임꼴', type: 'select', value: tri(cur?.italic), options: [{ value: '', label: '상관없음' }, { value: '1', label: '기울임꼴' }, { value: '0', label: '기울임꼴 아님' }] },
    { name: 'underline', label: '밑줄', type: 'select', value: tri(cur?.underline), options: [{ value: '', label: '상관없음' }, { value: '1', label: '밑줄' }, { value: '0', label: '밑줄 없음' }] },
    colorField('color', '글꼴 색'),
    colorField('fill', '채우기 색'),
    { name: 'size', label: '글꼴 크기', type: 'number', value: cur?.size ?? '' },
    { name: 'align', label: '가로 맞춤', type: 'select', value: cur?.align ?? '', options: [{ value: '', label: '상관없음' }, { value: 'left', label: '왼쪽' }, { value: 'center', label: '가운데' }, { value: 'right', label: '오른쪽' }] },
    { name: 'numFmt', label: '표시 형식', type: 'select', value: cur?.numFmt ?? '', options: nf.map(([value, label]) => ({ value, label })) },
  ], (v) => {
    const f = {};
    if (v.bold) f.bold = v.bold === '1';
    if (v.italic) f.italic = v.italic === '1';
    if (v.underline) f.underline = v.underline === '1';
    if (v.color) f.color = pickers.color.value;
    if (v.fill) f.fill = pickers.fill.value;
    if (v.size) f.size = Number(v.size);
    if (v.align) f.align = v.align;
    if (v.numFmt) f.numFmt = v.numFmt;
    done(Object.keys(f).length ? f : null);
  }, {
    onChange: (inputs) => {
      for (const k of ['color', 'fill']) {
        if (!pickers[k]) {
          pickers[k] = el('input', { type: 'color', value: cur?.[k] ?? (k === 'fill' ? '#ffff00' : '#ff0000'), style: { width: '44px', height: '24px', padding: '0' } });
          inputs[k].after(pickers[k]);
        }
        pickers[k].style.visibility = inputs[k].value ? 'visible' : 'hidden';
      }
    },
  });
  // 셀에서 서식 선택: 지금 칸의 서식을 그대로 (엑셀의 [셀에서 서식 선택])
  const foot = dlg.root.querySelector('.dialog-foot');
  foot.prepend(el('button', {
    class: 'btn', style: { marginRight: 'auto' },
    onclick: () => {
      const st = wb.styleAt(si, active.r, active.c) ?? {};
      const f = {};
      for (const k of ['bold', 'italic', 'underline', 'color', 'fill', 'font', 'size', 'align', 'numFmt', 'code']) if (st[k] !== undefined && st[k] !== false) f[k] = st[k];
      for (const k of ['bold', 'italic', 'underline']) f[k] = !!st[k];
      dlg.close();
      done(f);
    },
  }, '셀에서 서식 선택'));
  foot.prepend(el('button', { class: 'btn', onclick: () => { dlg.close(); done(null); } }, '지우기'));
}

let findDlg = null;
function openFindDialog(tab = 'find') {
  if (findDlg) findDlg.close();
  const sel0 = displayText(active.r, active.c);
  const findInput = el('input', { type: 'text', value: findState.text || (sel0 && sel0.length < 60 ? sel0 : '') });
  const replInput = el('input', { type: 'text', value: findState.replace });
  const caseBox = el('input', { type: 'checkbox', checked: findState.matchCase });
  const wholeBox = el('input', { type: 'checkbox', checked: findState.whole });
  const regexBox = el('input', { type: 'checkbox', checked: findState.regex });
  const scopeSel = el('select', {}, el('option', { value: 'sheet' }, '시트'), el('option', { value: 'book' }, '통합 문서'));
  const orderSel = el('select', {}, el('option', { value: 'rows' }, '행'), el('option', { value: 'cols' }, '열'));
  const lookSel = el('select', {});
  scopeSel.value = findState.scope;
  orderSel.value = findState.byCols ? 'cols' : 'rows';
  const status = el('div', { class: 'find-status' });
  const results = el('div', { class: 'find-results', style: { display: 'none' } });
  const fmtBox = el('span', {});
  const rfmtBox = el('span', {});
  const optsBox = el('div', { class: 'find-opts' });
  const replaceRow = el('div', { class: 'find-row' }, el('span', {}, '바꿀 내용:'), replInput, rfmtBox);
  const tabs = el('div', { class: 'find-tabs' });
  let mode = tab;
  const sync = () => Object.assign(findState, {
    text: findInput.value, replace: replInput.value, matchCase: caseBox.checked, whole: wholeBox.checked, regex: regexBox.checked,
    scope: scopeSel.value, byCols: orderSel.value === 'cols', lookIn: mode === 'replace' ? 'formulas' : lookSel.value,
  });
  const drawFmt = () => {
    const fmtBtn = (kind) => el('button', {
      class: 'btn fmt-btn', title: '서식으로 찾기 / 바꾸기',
      onclick: () => findFormatDialog(kind === 'find' ? '찾을 서식' : '바꿀 서식', kind === 'find' ? findState.format : findState.replaceFormat, (f) => {
        if (kind === 'find') findState.format = f; else findState.replaceFormat = f;
        drawFmt();
      }),
    }, '서식...');
    const show = findState.options;
    fmtBox.replaceChildren(...(show ? [fmtPreview(findState.format), fmtBtn('find')] : []));
    rfmtBox.replaceChildren(...(show ? [fmtPreview(findState.replaceFormat), fmtBtn('replace')] : []));
  };
  const drawOpts = () => {
    lookSel.replaceChildren(...(mode === 'replace' ? [['formulas', '수식']] : [['formulas', '수식'], ['values', '값'], ['comments', '메모']]).map(([v, l]) => el('option', { value: v }, l)));
    lookSel.value = mode === 'replace' ? 'formulas' : findState.lookIn;
    optsBox.style.display = findState.options ? '' : 'none';
    optsBox.replaceChildren(
      el('label', {}, el('span', {}, '범위:'), scopeSel), el('label', {}, caseBox, '대/소문자 구분'), el('span', {}),
      el('label', {}, el('span', {}, '검색:'), orderSel), el('label', {}, wholeBox, '전체 셀 내용 일치'), el('span', {}),
      el('label', {}, el('span', {}, '찾는 위치:'), lookSel), el('label', {}, regexBox, '정규식 사용 (바꿀 내용에 $1 등 그룹 참조 가능)'), el('span', {}),
      el('span', {}), el('span', { class: 'muted', style: { fontSize: '11px' } }, '와일드카드: * (여러 글자) ? (한 글자) ~ (글자 그대로)'), el('span', {}),
    );
    optBtn.textContent = findState.options ? '옵션 <<' : '옵션 >>';
    drawFmt();
  };
  const optBtn = el('button', { class: 'btn', onclick: () => { findState.options = !findState.options; drawOpts(); } }, '옵션 >>');
  const drawTabs = () => {
    tabs.replaceChildren(...[['find', '찾기'], ['replace', '바꾸기']].map(([k, l]) => el('button', { class: mode === k ? 'on' : '', onclick: () => { mode = k; draw(); } }, l)));
  };
  const listAll = () => {
    sync();
    const list = findAllMatches();
    if (!list.length) { results.style.display = 'none'; status.textContent = '찾는 항목이 없습니다.'; return list; }
    const shown = list.slice(0, 5000);
    const rows = shown.map((m) => {
      const tr = el('tr', { class: 'row', onclick: () => { results.querySelectorAll('tr.cur').forEach((x) => x.classList.remove('cur')); tr.classList.add('cur'); goToMatch(m); } },
        el('td', {}, wb.sheets[m.si].name), el('td', {}, `$${colToName(m.c)}$${m.r + 1}`),
        el('td', {}, displayText(m.r, m.c, m.si)), el('td', {}, (wb.getRaw(m.si, m.r, m.c) ?? '').startsWith('=') ? wb.getRaw(m.si, m.r, m.c) : ''),
        el('td', {}, wb.getCell(m.si, m.r, m.c)?.comment ?? ''));
      return tr;
    });
    results.replaceChildren(el('table', {}, el('thead', {}, el('tr', {}, ['시트', '셀', '값', '수식', '메모'].map((h) => el('th', {}, h)))), el('tbody', {}, rows)));
    results.style.display = '';
    status.textContent = `${list.length.toLocaleString()}개 셀을 찾았습니다.${list.length > shown.length ? ` (처음 ${shown.length.toLocaleString()}개 표시)` : ''}`;
    return list;
  };
  const body = el('div', { class: 'find-dlg', style: { display: 'flex', flexDirection: 'column', gap: '8px' } });
  const btnRow = el('div', { style: { display: 'flex', gap: '6px', justifyContent: 'flex-end', flexWrap: 'wrap' } });
  const draw = () => {
    drawTabs();
    drawOpts();
    replaceRow.style.display = mode === 'replace' ? '' : 'none';
    btnRow.replaceChildren(
      ...(mode === 'replace' ? [
        el('button', {
          class: 'btn', onclick: () => {
            sync();
            if (!findState.text && !findState.format) return;
            const list = findAllMatches('formulas');
            let n = 0;
            wb.transact(() => { for (const m of list) if (replaceIn(m.r, m.c, m.si)) n++; }, meta());
            findCache = null;
            results.style.display = 'none';
            status.textContent = n ? `모두 바꾸었습니다. ${n.toLocaleString()}개 항목을 바꾸었습니다.` : '바꿀 항목이 없습니다.';
            renderAll();
          },
        }, '모두 바꾸기'),
        el('button', {
          class: 'btn', onclick: () => {
            sync();
            const list = findAllMatches('formulas');
            const here = list.some((m) => m.si === si && m.r === active.r && m.c === active.c);
            if (here) { wb.transact(() => replaceIn(active.r, active.c), meta()); findCache = null; }
            if (!findNext(findState.text, { quiet: true })) status.textContent = '더 이상 바꿀 항목이 없습니다.';
            renderAll();
          },
        }, '바꾸기'),
      ] : []),
      el('button', { class: 'btn', onclick: () => listAll() }, '모두 찾기'),
      el('button', { class: 'btn', title: 'Shift+Enter', onclick: () => { sync(); status.textContent = ''; findNext(findState.text, { back: true }); } }, '이전 찾기'),
      el('button', { class: 'btn primary', title: 'Enter', onclick: () => { sync(); status.textContent = ''; findNext(); } }, '다음 찾기'),
      el('button', { class: 'btn', onclick: () => findDlg?.close() }, '닫기'),
    );
  };
  body.append(tabs,
    el('div', { class: 'find-row' }, el('span', {}, '찾을 내용:'), findInput, fmtBox),
    replaceRow, optsBox, el('div', { style: { display: 'flex', justifyContent: 'flex-end' } }, optBtn), btnRow, status, results);
  body.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' && e.target.tagName === 'INPUT') {
      e.preventDefault();
      e.stopPropagation();
      sync();
      status.textContent = '';
      findNext(findState.text, { back: e.shiftKey });
    }
  });
  body.addEventListener('change', () => { sync(); findCache = null; });
  draw();
  findDlg = openDialog({ title: '찾기 및 바꾸기', body, width: 560, modeless: true, onClose: () => { findDlg = null; } });
  findInput.focus();
  findInput.select();
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
  applySheetZoom();
  const saved = restore ? sheetSel.get(sheet()) : null;
  if (saved) {
    selectRange(saved.sel, saved.selKind, saved.active);
    gv.setScroll(...saved.scroll);
  } else showSheetStart();
  renderSheetTabs();
  setMode();
}

// ── 시트 숨기기 / 숨기기 취소 ──
const isHiddenSheet = (i) => { const st = wb.sheets[i]?.state; return st === 'hidden' || st === 'veryHidden'; };
const visibleSheetCount = () => wb.sheets.filter((_, i) => !isHiddenSheet(i)).length;

function hideSheet(i = si) {
  if (visibleSheetCount() <= 1) { alertDialog('WIXEL', '통합 문서에는 보이는 시트가 하나 이상 있어야 합니다.'); return; }
  wb.transact(() => wb.setSheetProp(i, 'state', 'hidden'), meta());
  if (i === si) {
    let j = i + 1;
    while (j < wb.sheets.length && isHiddenSheet(j)) j++;
    if (j >= wb.sheets.length) { j = i - 1; while (j >= 0 && isHiddenSheet(j)) j--; }
    switchSheet(j, true);
  }
  renderSheetTabs();
}

function unhideSheetDialog() {
  const hidden = wb.sheets.map((s, i) => [s, i]).filter(([, i]) => isHiddenSheet(i));
  if (!hidden.length) { toast('숨겨진 시트가 없습니다.'); return; }
  const list = el('select', { size: Math.min(10, Math.max(4, hidden.length)), multiple: true, style: { width: '100%' } },
    hidden.map(([s, i], k) => el('option', { value: String(i), selected: k === 0 }, s.name)));
  openDialog({
    title: '숨기기 취소', width: 340,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } }, el('div', {}, '숨기기 취소할 시트:'), list),
    buttons: [{
      label: '확인', primary: true, action: () => {
        const chosen = [...list.selectedOptions].map((o) => Number(o.value));
        if (!chosen.length) return false;
        wb.transact(() => chosen.forEach((i) => wb.setSheetProp(i, 'state', undefined)), meta());
        switchSheet(chosen[0]);
        renderSheetTabs();
        return undefined;
      },
    }, { label: '취소' }],
  });
}

function renderSheetTabs() {
  dom.sheetTabs.replaceChildren(...wb.sheets.map((s, i) => el('button', {
    class: `sheet-tab${i === si ? ' active' : ''}${validColor(s.tabColor) ? ' colored' : ''}`,
    style: isHiddenSheet(i) ? { display: 'none' } : validColor(s.tabColor) ? { '--tab-c': s.tabColor, '--tab-t': contrastText(s.tabColor) } : undefined,
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
        { sep: true },
        { label: '탭 색', icon: 'fill', action: () => setTimeout(() => tabColorMenu({ x: e.clientX, y: e.clientY - 330 }, i), 0) },
        { sep: true },
        { label: '숨기기', action: () => hideSheet(i) },
        { label: '숨기기 취소...', disabled: !wb.sheets.some((_, j) => isHiddenSheet(j)), action: () => unhideSheetDialog() },
      ]);
    },
  }, s.name)));
  dom.sheetTabs.children[si]?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
}

/** 글자색: 배경이 어두우면 흰색 */
/** 탭 색 등으로 쓸 수 있는 색인지 (#rgb · #rrggbb) — 잘못된 값이면 색 없이 그림 */
const validColor = (c) => typeof c === 'string' && /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(c.trim());
function contrastText(hex) {
  const h = /^#?([0-9a-f]{6})$/i.exec(String(hex).trim())?.[1] ?? /^#?([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(String(hex).trim())?.slice(1).map((x) => x + x).join('');
  if (!h) return 'var(--text)';
  const n = parseInt(h, 16);
  const lum = (0.299 * (n >> 16) + 0.587 * ((n >> 8) & 255) + 0.114 * (n & 255)) / 255;
  return lum < 0.6 ? '#ffffff' : '#1f1f1f';
}

/** 시트 탭 색 */
function tabColorMenu(anchor, i) {
  const targets = [i];
  paletteMenu(anchor, '색 없음', (color) => {
    wb.transact(() => { for (const t of targets) wb.setSheetProp(t, 'tabColor', color || null); }, meta());
    renderSheetTabs();
  });
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
    wb.snapshotList();
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

function exportCsv(kind = 'csv', name = docName) {
  const tab = kind !== 'csv';
  download(`${safeFileName(name)}-${safeFileName(sheet().name)}.${kind}`, `﻿${toDelimited(sheetToRows(si), tab ? '\t' : ',')}`, `${tab ? 'text/tab-separated-values' : 'text/csv'};charset=utf-8`);
  toast(`${kind.toUpperCase()} 파일로 내보냈습니다 (현재 시트, UTF-8).`);
}

/** OpenDocument 스프레드시트(.ods)로 저장 — 값 · 수식 · 서식 · 병합 · 열 너비 */
function exportOds(name = docName) {
  const info = wb.sheets.map((s, i) => ({ si: i, name: s.name, merges: s.merges ?? [], hidden: s.state === 'hidden' || s.state === 'veryHidden' }));
  const bytes = writeOds(info, {
    raw: (s, r, c) => wb.getRaw(s, r, c), value: (s, r, c) => wb.getValue(s, r, c), style: (s, r, c) => wb.styleAt(s, r, c),
    used: (s) => wb.usedRange(s), colWidth: (s, c) => wb.colWidth(s, c), rowHeight: (s, r) => wb.sheets[s].rowHeights?.[r] ?? null,
    display: (s, r, c) => displayText(r, c, s),
  });
  download(`${safeFileName(name)}.ods`, new Blob([bytes], { type: 'application/vnd.oasis.opendocument.spreadsheet' }));
  toast('OpenDocument 스프레드시트(.ods)로 저장했습니다.');
}

const XLSX_KINDS = {
  xlsx: { mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', label: 'Excel 통합 문서' },
  xlsm: { mime: 'application/vnd.ms-excel.sheet.macroEnabled.12', label: 'Excel 매크로 사용 통합 문서' },
  xltx: { mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template', label: 'Excel 서식 파일' },
  xltm: { mime: 'application/vnd.ms-excel.template.macroEnabled.12', label: 'Excel 매크로 사용 서식 파일' },
};

async function exportXlsx(name = docName, kind = null) {
  const over = xlsxOverflow(wb);
  if (over) toast(`엑셀 파일은 1,048,576행까지만 저장할 수 있어 그 아래 셀 ${over.toLocaleString()}개는 빠집니다. 전체는 .wixel 로 저장하세요.`);
  const k = kind ?? (wb.vba ? 'xlsm' : 'xlsx');
  // 큰 문서는 나눠서 만들고 진행 표시 (압축도 함께 해서 파일이 작아짐)
  const prog = progressOverlay(`'${safeFileName(name)}' 저장 중`);
  try {
    const bytes = await writeXlsxAsync(wb, { activeSheet: si, fileName: `${safeFileName(name)}.${k}`, kind: k }, (st) => prog.set(st.p, st.msg));
    prog.close();
    download(`${safeFileName(name)}.${k}`, new Blob([bytes], { type: XLSX_KINDS[k].mime }));
    toast(`${XLSX_KINDS[k].label}(.${k})로 저장했습니다.`);
  } catch (err) {
    prog.close();
    alertDialog('WIXEL', `저장하지 못했습니다: ${err.message}`);
  }
}

function saveAs() {
  formDialog('다른 이름으로 저장', [
    { name: 'name', label: '파일 이름', value: docName },
    {
      name: 'type', label: '파일 형식', type: 'select', value: 'xlsx', options: [
        { value: wb.vba ? 'xlsm' : 'xlsx', label: wb.vba ? 'Excel 매크로 사용 통합 문서 (*.xlsm)' : 'Excel 통합 문서 (*.xlsx)' },
        ...(wb.vba ? [{ value: 'xlsx', label: 'Excel 통합 문서 (*.xlsx) — 매크로 제외' }] : [{ value: 'xlsm', label: 'Excel 매크로 사용 통합 문서 (*.xlsm)' }]),
        { value: 'xltx', label: 'Excel 서식 파일 (*.xltx)' },
        { value: 'xltm', label: 'Excel 매크로 사용 서식 파일 (*.xltm)' },
        { value: 'ods', label: 'OpenDocument 스프레드시트 (*.ods)' },
        { value: 'wixel', label: 'WIXEL 통합 문서 (*.wixel) — 행 제한 없음' },
        { value: 'csv', label: 'CSV UTF-8 (쉼표로 분리) (*.csv) — 현재 시트' },
        { value: 'tsv', label: '텍스트 (탭으로 분리) (*.tsv) — 현재 시트' },
        { value: 'txt', label: '유니코드 텍스트 (*.txt) — 현재 시트' },
        ...(server.available ? [{ value: 'server', label: '서버에 저장 (다른 기기에서 열기)' }] : []),
      ],
    },
  ], ({ name, type }) => {
    const newName = name.trim() || docName;
    if (newName !== docName) renameDoc(newName);
    if (XLSX_KINDS[type]) exportXlsx(newName, type === 'xlsx' && wb.vba ? 'xlsx' : type);
    else if (type === 'csv' || type === 'tsv' || type === 'txt') exportCsv(type, newName);
    else if (type === 'ods') exportOds(newName);
    else if (type === 'server') saveNow(true);
    else download(`${safeFileName(newName)}.wixel`, JSON.stringify(snapshot()), 'application/json');
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
    if (/\.(xlsx|xlsm|xlsb|xltx|xltm)$/i.test(file.name)) {
      // 큰 파일도 화면이 멈추지 않도록 나눠서 읽고, 진행 상황을 보여 줌
      const prog = progressOverlay(`'${file.name}' 여는 중`);
      let res;
      try {
        res = await readXlsxAsync(new Uint8Array(await file.arrayBuffer()), (st) => prog.set(st.p * 0.6, st.msg));
        if (fileMode === 'open') await loadWorkbookAsync(res.data, base, res.active, prog);
      } finally {
        prog.close();
      }
      const { data, warnings } = res;
      if (fileMode === 'open') {
        // 이미 불러옴
      } else {
        wb.transact(() => {
          for (const s of data.sheets) {
            let name = s.name;
            for (let n = 2; wb.sheetIndexByName(name) >= 0; n++) name = `${s.name} (${n})`.slice(0, 31);
            const at = wb.addSheet(name);
            wb.replaceSheet(at, { ...s, name });
          }
        }, meta());
        toast(`시트 ${data.sheets.length}개를 가져왔습니다.`);
      }
      if (fileMode === 'open' && data.vba) warnings.push('매크로가 포함된 통합 문서입니다. WIXEL 는 매크로를 실행하지 않지만 [보기 → 매크로]에서 코드를 볼 수 있고, .xlsm 으로 저장하면 매크로가 그대로 유지됩니다.');
      if (warnings.length) alertDialog('가져오기', warnings.join('\n'));
      else if (fileMode === 'open') toast(`'${file.name}'을(를) 열었습니다.`);
      return;
    }
    // OpenDocument 스프레드시트 (.ods 압축 · .fods 단일 XML)
    if (/\.(ods|fods)$/i.test(file.name)) {
      const res = readOds(/\.fods$/i.test(file.name) ? await file.text() : new Uint8Array(await file.arrayBuffer()));
      if (fileMode === 'open') loadWorkbook(res.data, base, 0);
      else wb.transact(() => { for (const s of res.data.sheets) { let nm = s.name; for (let n = 2; wb.sheetIndexByName(nm) >= 0; n++) nm = `${s.name} (${n})`.slice(0, 31); const at = wb.addSheet(nm); wb.replaceSheet(at, { ...s, name: nm }); } }, meta());
      toast(`'${file.name}'을(를) 열었습니다.`);
      return;
    }
    // 큰 CSV: 조각씩 읽어 바로 열 블록으로 (수백만 행도 화면이 멈추지 않음)
    if (/\.(csv|tsv|txt)$/i.test(file.name) && file.size > 8 * 1024 * 1024 && fileMode === 'open') {
      await openBigCsv(file, base);
      return;
    }
    const text = await readTextSmart(file);
    if (/\.(json|tabula|wixel)$/i.test(file.name)) {
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
    alertDialog('WIXEL', `파일을 열 수 없습니다: ${err.message}`);
  }
}

/** 글자 파일: UTF-8(BOM 포함)이 기본, UTF-8 이 아니면 한글 엑셀 CSV 의 EUC-KR(CP949), UTF-16 BOM 도 인식 */
async function readTextSmart(file) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch { /* UTF-8 아님 */ }
  try { return new TextDecoder('euc-kr').decode(bytes); } catch { return new TextDecoder().decode(bytes); }
}

async function openBigCsv(file, base) {
  const prog = progressOverlay(`'${file.name}' 여는 중`);
  try {
    const head = await file.slice(0, 65536).text();
    const reader = new CsvBlockReader(guessDelimiter(head), Math.min(1 << 24, Math.max(1024, Math.round(file.size / 50))));
    const rd = file.stream().pipeThrough(new TextDecoderStream()).getReader();
    let read = 0;
    let last = performance.now();
    for (;;) {
      const { value, done } = await rd.read();
      if (done) break;
      reader.push(value);
      read += value.length;
      if (performance.now() - last > 80) {
        prog.set(0.85 * Math.min(1, read / file.size), `${reader.n.toLocaleString()}행 읽는 중`);
        await yieldUI();
        last = performance.now();
      }
    }
    const { header, block } = reader.finish();
    const cells = new Map();
    header.forEach((h, j) => { if (h !== '') cells.set(`0,${j}`, { raw: h, style: { bold: true } }); });
    await loadWorkbookAsync({ sheets: [{ name: base.slice(0, 31) || 'Sheet1', cells, blocks: [block], freeze: { rows: 1, cols: 0 } }] }, base, 0, prog);
    toast(`'${file.name}' — ${block.n.toLocaleString()}행을 열었습니다.`);
  } finally {
    prog.close();
  }
}

function writeRows(rows, r0, c0) {
  wb.transact(() => rows.forEach((row, i) => row.forEach((v, j) => { if (v !== '') wb.setInput(si, r0 + i, c0 + j, v); })), meta());
  if (rows.length) selectRange({ r1: r0, c1: c0, r2: r0 + rows.length - 1, c2: c0 + Math.max(1, ...rows.map((r) => r.length)) - 1 }, 'cells', { r: r0, c: c0 });
}

function loadWorkbook(data, name, activeSheet = 0) {
  if (editing) endEditUI();
  libraryFlush();
  wb.load(data);
  renderImportedPivots();
  afterLoad(name, activeSheet);
}

/** 큰 파일: 셀 준비와 피벗 다시 그리기를 나눠서 (진행 표시 prog: {set(p, 메시지)}) */
async function loadWorkbookAsync(data, name, activeSheet, prog) {
  if (editing) endEditUI();
  libraryFlush();
  const next = new Workbook();
  await next.loadAsync(data, (p) => prog?.set(0.6 + 0.25 * p, '셀 준비 중'));
  next.listeners = wb.listeners; // 화면 갱신 연결 유지
  wb = next;
  si = clamp(activeSheet, 0, wb.sheets.length - 1);
  if (isHiddenSheet(si)) si = Math.max(0, wb.sheets.findIndex((_, i) => !isHiddenSheet(i)));
  sheetSel.clear();
  chartSel = null;
  const list = allPivots().filter((e) => e.def.captureFmt || e.def.needsRender);
  // 큰 원본: 모든 피벗 · 슬라이서 필드로 요약 캐시를 한 번에 준비
  prog?.set(0.85, '요약 캐시 준비 중');
  await yieldUI();
  warmAll();
  for (let i = 0; i < list.length; i++) {
    prog?.set(0.85 + 0.15 * (i / Math.max(1, list.length)), `피벗 테이블 계산 중 (${i + 1}/${list.length})`);
    await new Promise((res) => setTimeout(res, 0));
    const e = list[i];
    // 한 피벗의 셀 수만 개를 한 번에 반영 (칸마다 의존 수식을 찾지 않게) — 실행 취소 기록은 afterLoad 에서 비움
    try { wb.transact(() => writePivot(e.si, e.def)); } catch (err) { console.warn('피벗 다시 그리기 실패', err); }
  }
  afterLoad(name, activeSheet);
}

/** 진행 표시 창 */
function progressOverlay(title) {
  const bar = el('div', { class: 'lp-bar' });
  const msg = el('div', { class: 'lp-msg' }, '');
  const box = el('div', { class: 'load-progress', role: 'progressbar' }, el('div', { class: 'lp-box' }, el('div', { class: 'lp-title' }, title), msg, el('div', { class: 'lp-track' }, bar)));
  document.body.append(box);
  return {
    set(p, m) { bar.style.width = `${Math.round(Math.max(0, Math.min(1, p)) * 100)}%`; if (m) msg.textContent = m; },
    close() { box.remove(); },
  };
}

/** 통합 문서 기본 글꼴 (엑셀의 표준 스타일: 셀 기본 크기 · 열 너비 기준) · 테마 색 */
function applyBookLook() {
  setBaseFont(wb.defaultFont);
  setThemeColors(wb.theme);
  document.documentElement.style.setProperty('--cell-fs', `${BASE_FONT.size}pt`);
  document.documentElement.style.setProperty('--cell-ff', fontStack(BASE_FONT.name));
  // 설치된 글꼴에 맞춰 글자를 세로 가운데로 (글꼴이 늦게 로드되면 다시 잼)
  const shift = () => document.documentElement.style.setProperty('--glyph-dy', `${glyphShift(fontStack(BASE_FONT.name))}em`);
  shift();
  document.fonts?.ready?.then(shift);
}

/** 예전 버전이 자동 저장한 문서: 피벗 결과 칸을 지금 버전으로 다시 그림 (서식 · 색 개선이 반영되게) — 실행 취소 기록 없음 */
function redrawPivotsQuiet() {
  try {
    wb.transact(() => {
      wb.sheets.forEach((s, i) => { for (const { def } of pivotDefs(i)) writePivot(i, def, { autofit: false }); });
    }, meta());
  } catch (e) {
    console.warn('피벗 다시 그리기 실패', e);
  }
  wb.undoStack = [];
  wb.redoStack = [];
}

function afterLoad(name, activeSheet) {
  // 새로 연 문서는 보관함의 새 항목 (보관함에서 연 문서는 그 항목)
  docId = pendingDocId ?? newDocId();
  pendingDocId = null;
  lastVersionAt = 0;
  libDirty = true;
  scheduleLibrarySave();
  wb.undoStack = [];
  wb.redoStack = [];
  docName = name || '통합 문서1';
  si = clamp(activeSheet, 0, wb.sheets.length - 1);
  if (isHiddenSheet(si)) si = Math.max(0, wb.sheets.findIndex((_, i) => !isHiddenSheet(i)));
  clip = null;
  painter = null;
  chartSel = null;
  circles = null;
  objClip = null;
  sheetSel.clear();
  applyBookLook();
  applyOptions();
  fitRowsOnOpen();
  gv.resetExtent();
  applySheetZoom();
  renderAll();
  showSheetStart();
  // 수식 의존 그래프를 쉬는 동안 미리 만듦 (첫 편집도 바로 다시 계산)
  setTimeout(() => { wb.prepareGraph().catch((e) => console.warn('의존 그래프 준비 실패', e)); }, 1200);
  dirty = true;
  if (bigBook()) scheduleAutosave();
  else { saveToStorage(); scheduleServerSave(0); }
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
      if (sample.big) {
        // 빅데이터 예제: 데이터 만들기 · 피벗 계산을 진행 표시와 함께
        const prog = progressOverlay(`'${sample.name}' 만드는 중`);
        try {
          prog.set(0.05, '데이터 만드는 중');
          await yieldUI();
          const data = sample.build();
          await loadWorkbookAsync(data, name, 0, prog);
        } finally {
          prog.close();
        }
      } else loadWorkbook(sample.build(), name);
    } else {
      let n = 1;
      while (names.includes(`통합 문서${n}`) || `통합 문서${n}` === docName) n++;
      loadWorkbook({ sheets: [{ name: 'Sheet1', cells: {} }] }, `통합 문서${n}`);
    }
  };
  if (!autosave && dirty) {
    openDialog({
      title: 'WIXEL', body: '저장하지 않은 변경 내용이 있습니다. 새 통합 문서를 만드시겠습니까?',
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

const snapshot = () => ({ app: 'wixel', docName, docId, si, workbook: wb.serialize() });

// ───── 이 브라우저의 문서 보관함 (최근 문서 30개 · 버전 기록) ─────
let docId = null;
let pendingDocId = null; // 보관함에서 연 문서의 id (불러온 뒤 docId 로)
let libDirty = false;
let libTimer = null;
let lastVersionAt = 0;
const LIB_CELL_LIMIT = 300000; // 이보다 큰 문서는 보관함 대신 큰 문서 자동 저장만
const VERSION_EVERY = 10 * 60 * 1000; // 편집 중에는 10분마다 버전 하나
function scheduleLibrarySave() {
  clearTimeout(libTimer);
  libTimer = setTimeout(() => whenIdle(() => libraryFlush()), 5000);
}
/** 지금 문서를 보관함에 (직렬화는 바로 — 다른 문서를 열기 직전에 불러도 안전), version: { label } 이면 버전 기록에도 */
function libraryFlush({ version = null, force = false } = {}) {
  clearTimeout(libTimer);
  if (viewOnly) return Promise.resolve(null);
  if (!force && !libDirty && !version) return Promise.resolve(null);
  if (!version && cellCount() === 0 && !wb.sheets.some((x) => x.charts?.length || x.shapes?.length)) return Promise.resolve(null);
  if (!wb?.sheets?.length || cellCount() > LIB_CELL_LIMIT) return Promise.resolve(null);
  docId ??= newDocId();
  let json;
  try { json = JSON.stringify(snapshot()); } catch { return Promise.resolve(null); }
  libDirty = false;
  const now = Date.now();
  const ver = version ?? (now - lastVersionAt > (Number(opts.verMinutes) || 10) * 60000 ? { label: '' } : null);
  if (ver) lastVersionAt = now;
  const info = { sheets: wb.sheets.length, cells: cellCount(), sheetNames: wb.sheets.slice(0, 6).map((x) => x.name) };
  return libSave(docId, docName, json, { version: ver, info, max: clamp(Number(opts.libMax) || LIB_MAX, 5, 200) }).catch((err) => { console.warn('보관함 저장 실패', err); return null; });
}
async function openFromLibrary(id, { ts = null, copy = false } = {}) {
  await libraryFlush();
  const prog = progressOverlay('보관함에서 여는 중');
  try {
    const data = ts ? await libLoadVersion(id, ts) : await libLoad(id);
    if (!data?.workbook) { alertDialog('WIXEL', '보관함에서 문서를 찾을 수 없습니다.'); return; }
    pendingDocId = copy ? newDocId() : id;
    const name = copy ? `${data.docName ?? '통합 문서'} (${formatDate(ts ?? Date.now())} 버전)` : data.docName;
    await loadWorkbookAsync(data.workbook, name, data.si ?? 0, prog);
  } finally {
    prog.close();
  }
}

/** 보관함 목록 표 (파일 → 열기 · 최근 항목) */
async function recentTable(close) {
  const box = el('div', { class: 'recent-box' });
  let list = [];
  try { list = await libList(); } catch { /* IndexedDB 없음 */ }
  const kb = (n) => (n > 1048576 ? `${(n / 1048576).toFixed(1)}MB` : `${Math.max(1, Math.round(n / 1024))}KB`);
  const draw = () => {
    box.replaceChildren(
      el('div', { class: 'backstage-note' }, `이 브라우저에 최근 문서 ${LIB_MAX}개까지 자동으로 보관합니다 (문서마다 버전 기록 ${VER_MAX}개). 창을 닫거나 다른 문서를 열어도 사라지지 않습니다.`),
      list.length ? el('table', { class: 'backstage-list' },
        el('tr', {}, el('th', {}, ''), el('th', {}, '이름'), el('th', {}, '수정한 날짜'), el('th', {}, '크기'), el('th', {}, '버전'), el('th', {}, '')),
        list.map((f) => {
          const open = () => { close(); openFromLibrary(f.id); };
          return el('tr', { class: 'file' },
            el('td', { class: 'pin-cell' }, el('button', {
              class: `lnk pin${f.pinned ? ' on' : ''}`, title: f.pinned ? '고정 해제' : '목록에 고정 (자동 정리하지 않음)',
              onclick: async () => { await libUpdate(f.id, { pinned: !f.pinned }); f.pinned = !f.pinned; draw(); },
            }, f.pinned ? '★' : '☆')),
            el('td', { onclick: open }, el('b', {}, f.name ?? '(이름 없음)'), f.id === docId ? el('span', { class: 'muted' }, ' (현재 문서)') : null,
              el('div', { class: 'muted small' }, `${f.sheets ?? 0}개 시트 · ${(f.sheetNames ?? []).join(', ')}`)),
            el('td', { onclick: open }, formatDate(f.updated)),
            el('td', {}, kb(f.size ?? 0)),
            el('td', {}, el('button', { class: 'lnk', onclick: () => { close(); versionHistory(f.id); } }, `${f.versions?.length ?? 0}개`)),
            el('td', {}, el('button', {
              class: 'btn', onclick: () => openDialog({
                title: '삭제', body: `'${f.name}'을(를) 이 브라우저의 보관함에서 삭제할까요? (버전 기록도 함께 삭제)`,
                buttons: [{ label: '삭제', primary: true, action: async () => { await libRemove(f.id); list = list.filter((x) => x !== f); draw(); } }, { label: '취소' }],
              }),
            }, '삭제')));
        })) : el('div', { class: 'muted' }, '보관된 문서가 없습니다.'),
    );
  };
  draw();
  return box;
}

/** 버전 기록 (구글 스프레드시트처럼 시각별 버전 · 이름 붙이기 · 복원 · 사본으로 열기) */
async function versionHistory(id = docId) {
  if (id === docId) await libraryFlush({ force: true });
  const list = await libList().catch(() => []);
  const entry = list.find((x) => x.id === id);
  if (!entry) { alertDialog('버전 기록', cellCount() > LIB_CELL_LIMIT ? '셀이 아주 많은 문서는 버전 기록 대신 큰 문서 자동 저장만 합니다. [다른 이름으로 저장]으로 파일을 보관하세요.' : '아직 보관된 버전이 없습니다.'); return; }
  const body = el('div', { class: 'ver-list' });
  const draw = () => {
    const vers = [...(entry.versions ?? [])].sort((a, b) => b.ts - a.ts);
    body.replaceChildren(
      el('div', { class: 'muted' }, `'${entry.name}' — 편집 중 10분마다, [저장](Ctrl+S)할 때마다 버전을 남깁니다. 최근 ${VER_MAX}개 (이름 붙인 버전은 오래 보관).`),
      el('div', { class: 'ver-rows' },
        el('div', { class: 'ver-row cur' }, el('b', {}, '현재 버전'), el('span', { class: 'muted' }, formatDate(entry.updated))),
        vers.map((v) => el('div', { class: 'ver-row' },
          el('div', {}, el('b', {}, formatDate(v.ts)), v.label ? el('span', { class: 'ver-label' }, v.label) : null),
          el('div', { class: 'ver-acts' },
            el('button', { class: 'lnk', onclick: () => formDialog('버전 이름', [{ name: 'n', label: '이름', value: v.label ?? '' }], async ({ n }) => { await libNameVersion(id, v.ts, n.trim()); v.label = n.trim(); v.named = !!n.trim(); draw(); }) }, '이름 지정'),
            el('button', { class: 'lnk', onclick: () => { dlg.close(); openFromLibrary(id, { ts: v.ts, copy: true }); } }, '사본으로 열기'),
            el('button', {
              class: 'btn', onclick: async () => {
                dlg.close();
                // 복원 전 지금 상태도 버전으로 남김
                if (id === docId) await libraryFlush({ version: { label: '복원 전' } });
                const data = await libLoadVersion(id, v.ts);
                if (!data?.workbook) { alertDialog('버전 기록', '버전을 찾을 수 없습니다.'); return; }
                pendingDocId = id;
                const prog = progressOverlay('버전 복원 중');
                try { await loadWorkbookAsync(data.workbook, data.docName ?? entry.name, data.si ?? 0, prog); } finally { prog.close(); }
                libraryFlush({ version: { label: `${formatDate(v.ts)} 버전으로 복원` } });
                toast(`${formatDate(v.ts)} 버전으로 복원했습니다.`);
              },
            }, '이 버전 복원')))),
      ),
    );
  };
  draw();
  const dlg = openDialog({
    title: '버전 기록', body, width: 560,
    buttons: [
      ...(id === docId ? [{ label: '지금 버전에 이름 지정...', action: () => { formDialog('버전 이름', [{ name: 'n', label: '이름', value: '' }], async ({ n }) => { const e = await libraryFlush({ version: { label: n.trim() || '이름 없는 버전' } }); if (e) { const last = e.versions[e.versions.length - 1]; await libNameVersion(id, last.ts, n.trim() || '이름 없는 버전'); } toast('버전을 저장했습니다.'); }); return true; } }] : []),
      { label: '닫기', primary: true },
    ],
  });
}

// ───── 공유 · 웹에 게시 (읽기 전용 대시보드 보기) ─────
const b64url = (u8) => {
  let s = '';
  for (let i = 0; i < u8.length; i += 0x8000) s += String.fromCharCode.apply(null, u8.subarray(i, i + 0x8000));
  return btoa(s).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '');
};
const fromB64url = (t) => {
  const s = atob(t.replace(/-/g, '+').replace(/_/g, '/'));
  const u8 = new Uint8Array(s.length);
  for (let i = 0; i < s.length; i++) u8[i] = s.charCodeAt(i);
  return u8;
};
const LINK_MAX = 1_500_000; // 주소에 담을 수 있는 크기 (브라우저 한계 2MB 아래)
const pubKey = () => `wixel.pub.${docId}`;
const pubInfo = () => { try { return JSON.parse(localStorage.getItem(pubKey()) ?? 'null'); } catch { return null; } };
const setPubInfo = (v) => { try { if (v) localStorage.setItem(pubKey(), JSON.stringify(v)); else localStorage.removeItem(pubKey()); } catch { /* 무시 */ } };
const appBase = () => location.href.replace(/[?#].*$/, '');

/** 게시용 스냅샷: 선택한 시트만 · 보기 옵션 */
function publishSnapshot(opt) {
  const snap = snapshot();
  const book = snap.workbook;
  if (opt.sheet !== 'all') {
    const i = Number(opt.sheet);
    book.sheets = book.sheets.map((s, k) => (k === i ? s : { ...s, state: s.state ?? 'hidden' }));
    snap.si = i;
  }
  return { ...snap, view: { headers: !!opt.headers, grid: !!opt.grid, title: opt.title ?? docName, published: Date.now() } };
}

async function publishDialog() {
  const info = pubInfo();
  const sheetSel = el('select', {}, el('option', { value: 'all' }, '전체 통합 문서'), wb.sheets.map((s, i) => (isHiddenSheet(i) ? null : el('option', { value: String(i), selected: i === si && false }, s.name))));
  const headers = el('input', { type: 'checkbox' });
  const grid = el('input', { type: 'checkbox', checked: true });
  const auto = el('input', { type: 'checkbox', checked: info?.auto ?? true });
  const out = el('div', { class: 'pub-out' });
  const show = (url, note) => {
    const inp = el('input', { type: 'text', readonly: true, value: url, onclick: (e) => e.target.select() });
    out.replaceChildren(el('div', { class: 'pub-link' }, inp,
      el('button', { class: 'btn primary', onclick: async () => { try { await navigator.clipboard.writeText(url); toast('링크를 복사했습니다.'); } catch { inp.select(); document.execCommand('copy'); toast('링크를 복사했습니다.'); } } }, '복사'),
      el('button', { class: 'btn', onclick: () => window.open(url, '_blank') }, '미리 보기')),
      note ? el('div', { class: 'muted' }, note) : null);
    inp.select();
  };
  const opt = () => ({ sheet: sheetSel.value, headers: headers.checked, grid: grid.checked });
  const linkOnly = async () => {
    const { blob } = await packText(JSON.stringify(publishSnapshot(opt())));
    const code = b64url(new Uint8Array(await blob.arrayBuffer()));
    if (code.length > LINK_MAX) { out.replaceChildren(el('div', { class: 'warn' }, `문서가 커서(${Math.round(code.length / 1024)}KB) 링크에 담을 수 없습니다. 서버(npm start)로 게시하거나 시트 하나만 게시하세요.`)); return; }
    show(`${appBase()}#view=${code}`, '링크 안에 문서 내용이 압축되어 들어 있어 서버 없이도 누구나 볼 수 있습니다 (읽기 전용 · 그 시점의 내용).');
  };
  const serverPub = async () => {
    try {
      const res = await server.publish(publishSnapshot(opt()), info?.id ?? null);
      setPubInfo({ id: res.id, auto: auto.checked, opt: opt() });
      show(`${appBase()}?view=${res.id}`, `서버에 게시했습니다. ${auto.checked ? '문서를 고치면 게시본도 자동으로 새로 고쳐지고, 보는 사람 화면도 30초마다 갱신됩니다.' : '[다시 게시]를 눌러야 게시본이 바뀝니다.'}`);
    } catch (err) {
      if (err.status === 401) askServerToken(serverPub);
      else out.replaceChildren(el('div', { class: 'warn' }, `게시하지 못했습니다: ${err.message}`));
    }
  };
  const body = el('div', { class: 'an-dlg' },
    el('div', { class: 'muted' }, '구글 스프레드시트의 [웹에 게시]처럼 읽기 전용 대시보드 화면(리본 · 수식 입력줄 없이)으로 공유합니다. 보는 사람은 슬라이서 · 필터를 눌러 볼 수 있지만 원본은 바뀌지 않습니다.'),
    el('label', { class: 'an-row' }, el('span', {}, '게시할 내용'), sheetSel),
    el('label', { class: 'an-check' }, grid, '눈금선 표시'),
    el('label', { class: 'an-check' }, headers, '행 · 열 머리글 표시'),
    server.available ? el('label', { class: 'an-check' }, auto, '변경할 때마다 자동으로 다시 게시 (서버)') : null,
    el('div', { class: 'backstage-actions' },
      server.available ? el('button', { class: 'btn primary', onclick: serverPub }, info?.id ? '다시 게시 (같은 링크)' : '서버에 게시 (짧은 링크)') : null,
      el('button', { class: `btn${server.available ? '' : ' primary'}`, onclick: linkOnly }, '링크 만들기 (서버 없이)'),
      info?.id && server.available ? el('button', { class: 'btn', onclick: async () => { await server.unpublish(info.id).catch(() => {}); setPubInfo(null); out.replaceChildren(el('div', { class: 'muted' }, '게시를 중지했습니다. 이전 링크로는 더 이상 볼 수 없습니다.')); } }, '게시 중지') : null),
    out,
    server.available ? el('div', { class: 'muted' }, `공동 작업: 같은 서버에 접속한 사람에게 ${appBase()}?doc=${encodeURIComponent(docName)} 를 보내면 이 문서를 함께 편집할 수 있습니다 (다른 사람이 저장한 내용은 자동으로 불러옴).`) : null,
  );
  openDialog({ title: '공유 · 웹에 게시', body, width: 620, buttons: [{ label: '닫기', primary: true }] });
  if (info?.id && server.available) show(`${appBase()}?view=${info.id}`, '이미 게시된 문서입니다.');
}
/** 자동 다시 게시 (서버 저장과 함께) */
function autoRepublish() {
  const info = pubInfo();
  if (!info?.id || !info.auto || !server.available || viewOnly) return;
  server.publish(publishSnapshot(info.opt ?? { sheet: 'all', grid: true }), info.id).catch(() => {});
}

// 읽기 전용 보기 (게시된 문서 · 링크)
let viewOnly = false;
function enterViewMode(data, { pubId = null } = {}) {
  viewOnly = true;
  autosave = false;
  document.body.classList.add('view-mode');
  const v = data.view ?? {};
  view.showHeaders = !!v.headers;
  if (v.grid === false) view.showGrid = false;
  const bar = el('div', { class: 'view-bar' },
    el('b', {}, v.title ?? data.docName ?? 'WIXEL'),
    el('span', { class: 'muted' }, ` · 읽기 전용${v.published ? ` · 게시: ${formatDate(v.published)}` : ''}`),
    el('span', { class: 'view-live', hidden: !pubId }, '● 자동 새로 고침'),
    el('button', { class: 'btn', onclick: () => { exitViewMode(data); } }, '편집용 사본 만들기'));
  document.getElementById('app').prepend(bar);
  loadWorkbook(data.workbook, data.docName ?? 'WIXEL', data.si ?? 0);
  applyView();
  renderAll();
  if (pubId) {
    let seen = 0;
    setInterval(async () => {
      try {
        const res = await server.published(pubId);
        if (seen && res.modified > seen) {
          const keep = si;
          loadWorkbook(res.data.workbook, res.data.docName ?? docName, keep);
          toast('게시된 문서가 업데이트되었습니다.');
        }
        seen = res.modified;
      } catch { /* 게시 중지 등 */ }
    }, 30000);
  }
}
function exitViewMode(data) {
  viewOnly = false;
  autosave = true;
  document.body.classList.remove('view-mode');
  document.querySelector('.view-bar')?.remove();
  view.showHeaders = true;
  view.showGrid = true;
  history.replaceState(null, '', appBase());
  loadWorkbook(data.workbook, `${data.docName ?? '통합 문서'} 사본`, data.si ?? 0);
  applyView();
  toast('편집할 수 있는 사본을 만들었습니다 (이 브라우저의 보관함에 저장).');
}
/** 주소로 연 경우: #view=압축 · ?view=게시id · ?doc=서버 문서 */
async function openFromUrl() {
  const hash = location.hash;
  const q = new URLSearchParams(location.search);
  try {
    if (hash.startsWith('#view=')) {
      const text = await unpackText({ gz: true, blob: new Blob([fromB64url(hash.slice(6))]) });
      enterViewMode(JSON.parse(text));
      return true;
    }
    if (q.get('view') && server.available) {
      const res = await server.published(q.get('view'));
      enterViewMode(res.data, { pubId: q.get('view') });
      return true;
    }
    if (q.get('doc') && server.available) {
      await openFromServer(q.get('doc'));
      return true;
    }
  } catch (err) {
    alertDialog('WIXEL', `링크의 문서를 열 수 없습니다: ${err.message}`);
  }
  return false;
}

// 공동 작업(서버): 다른 사람이 저장한 새 내용이 있고 내 변경이 없으면 자동으로 불러옴
let collabTimer = null;
function startCollabWatch() {
  clearInterval(collabTimer);
  if (!server.available) return;
  collabTimer = setInterval(async () => {
    if (viewOnly || dirty || editing || !serverState.savedAt) return;
    try {
      const files = await server.list();
      const f = files.find((x) => x.name === docName);
      if (f && f.modified > serverState.savedAt + 1500) {
        const keep = { si, r: active.r, c: active.c };
        const data = await server.load(docName);
        pendingDocId = docId;
        loadWorkbook(data.workbook ?? data, data.docName ?? docName, keep.si);
        dirty = false;
        serverState.savedAt = f.modified;
        selectCell(keep.r, keep.c);
        updateTitle();
        toast('다른 사용자가 저장한 내용을 불러왔습니다.');
      }
    } catch { /* 네트워크 오류 무시 */ }
  }, 15000);
}

let storageWarned = false;
/** 셀이 많은 통합 문서 (자동 저장을 IndexedDB 로, 더 드물게) */
function cellCount() {
  let n = 0;
  for (const s of wb.sheets) {
    n += s.cells.size;
    for (const b of s.blocks ?? []) n += b.n * b.cols.length;
  }
  return n;
}
/** 서버 자동 저장 대상인지 (셀이 아주 많으면 브라우저에만 자동 저장) */
function serverAutosave() {
  return server.available && cellCount() <= 300000;
}
function bigBook() {
  return cellCount() > 50000;
}
let idbSaving = null;
function saveToStorage() {
  if (viewOnly) return false;
  try {
    if (bigBook()) {
      // 큰 문서: IndexedDB 에 시트별로, 바뀐 시트만, 조금씩 나눠 저장 (화면이 멈추지 않게)
      localStorage.setItem(STORAGE_KEY, JSON.stringify({ docName, docId, si, autosave, idb: true }));
      idbSaving = saveBigToIdb().then(() => { if (!serverAutosave()) { dirty = false; updateTitle(); } }).catch((err) => {
        if (err === SAVE_ABORT) return;
        if (!storageWarned) { storageWarned = true; toast('브라우저 저장 공간이 부족해 자동 저장하지 못했습니다. [파일 → 다른 이름으로 저장]으로 파일을 내려받으세요.'); }
      });
      return true;
    }
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ rev: APP_REV, docName, docId, si, autosave, workbook: wb.serialize() }));
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

const SAVE_ABORT = new Error('저장 중단');
const yieldUI = () => new Promise((res) => setTimeout(res, 0));
let bigSaveRun = null;
let bigSaveAgain = false;
/** 큰 문서 저장: 목록 { v: 2, sheets: [{id, ev}] } + 시트마다 { meta, chunks: [JSON 문자열] } */
async function saveBigToIdb() {
  if (bigSaveRun) { bigSaveAgain = true; return bigSaveRun; }
  const book = wb;
  bigSaveRun = (async () => {
    const prev = await idbGet(STORAGE_KEY).catch(() => null);
    const saved = new Map((prev?.v === 2 ? prev.sheets : []).map((x) => [x.id, x.ev]));
    const list = [];
    for (let i = 0; i < book.sheets.length; i++) {
      const s = book.sheets[i];
      s._sid ??= `s${Date.now().toString(36)}${Math.random().toString(36).slice(2, 8)}`;
      const ev = s._ev ?? 0;
      if (saved.get(s._sid) !== ev) {
        const chunks = [];
        let t = performance.now();
        for (const ch of book.cellChunks(i)) {
          chunks.push(await packChunk(JSON.stringify(ch)));
          if (performance.now() - t > 30) {
            await yieldUI();
            if (wb !== book || book.sheets[i] !== s) throw SAVE_ABORT; // 그사이 다른 문서를 열었음
            t = performance.now();
          }
        }
        // 열 블록: 형식화 배열을 16MB 조각으로 나눠 따로 저장 (천만 행 640MB 를 한 번에 복제하면 메모리가 두 배로 튀어 탭이 죽음)
        const prevRec = await idbGet(`${STORAGE_KEY}#${s._sid}`).catch(() => null);
        const gen = Date.now().toString(36);
        const partKeys = [];
        const blocks = [];
        for (let bi = 0; bi < (s.blocks ?? []).length; bi++) {
          const b = s.blocks[bi];
          const splitArr = async (arr, tag) => {
            if (!arr || arr.length * arr.BYTES_PER_ELEMENT <= PART_BYTES) return { inline: arr };
            const per = Math.floor(PART_BYTES / arr.BYTES_PER_ELEMENT);
            const keys = [];
            for (let p = 0; p * per < arr.length; p++) {
              const key = `${STORAGE_KEY}#${s._sid}#${gen}.${bi}.${tag}.${p}`;
              await idbSet(key, arr.slice(p * per, Math.min(arr.length, (p + 1) * per)));
              keys.push(key);
              partKeys.push(key);
              await yieldUI();
              if (wb !== book || book.sheets[i] !== s) throw SAVE_ABORT;
            }
            return { parts: keys, len: arr.length, kind: arr.constructor.name };
          };
          const cols = [];
          for (let ci = 0; ci < b.cols.length; ci++) {
            const c = b.cols[ci];
            const num = await splitArr(c.num, `${ci}n`);
            const str = await splitArr(c.str, `${ci}s`);
            cols.push({ ...c, num: num.inline ?? null, str: str.inline ?? null, ...(num.parts ? { numParts: num } : {}), ...(str.parts ? { strParts: str } : {}) });
          }
          const perm = await splitArr(b.perm ?? null, 'perm');
          blocks.push({ ...b, cols, perm: perm.inline ?? undefined, ...(perm.parts ? { permParts: perm } : {}) });
        }
        await idbSet(`${STORAGE_KEY}#${s._sid}`, { meta: book.sheetMeta(i), chunks, gz: GZ, blocks, partKeys });
        for (const k of prevRec?.partKeys ?? []) idbDel(k).catch(() => {});
      }
      list.push({ id: s._sid, ev });
    }
    if (wb !== book) throw SAVE_ABORT;
    await idbSet(STORAGE_KEY, {
      v: 2, rev: APP_REV, docName, si, autosave, book: book.bookMeta(), vba: book.vba ?? null,
      names: book.names.map(({ _ast, _text, ...n }) => ({ ...n })), sheets: list,
    });
    for (const id of saved.keys()) if (!list.some((x) => x.id === id)) idbDel(`${STORAGE_KEY}#${id}`).catch(() => {});
    // 저장하는 동안 바뀐 시트는 다음 저장에서 다시 (ev 가 달라짐)
    if (book.sheets.some((s, i) => (s._ev ?? 0) !== list[i]?.ev)) bigSaveAgain = true;
  })();
  try {
    await bigSaveRun;
  } finally {
    bigSaveRun = null;
    if (bigSaveAgain) { bigSaveAgain = false; scheduleAutosave(); }
  }
}

const PART_BYTES = 16 * 1024 * 1024; // 열 블록 저장 조각 크기
// 조각은 gzip 으로 압축한 Blob 으로 저장 (문자열 그대로 넣으면 IndexedDB 가 복사하는 동안 화면이 멈춤)
const GZ = typeof CompressionStream === 'function';
async function packChunk(text) {
  const blob = new Blob([text], { type: 'application/json' });
  return GZ ? new Response(blob.stream().pipeThrough(new CompressionStream('gzip'))).blob() : blob;
}
async function unpackChunk(x, gz) {
  if (typeof x === 'string') return x;
  return gz ? new Response(x.stream().pipeThrough(new DecompressionStream('gzip'))).text() : x.text();
}

/** IndexedDB 의 큰 문서 → 불러오기용 데이터 (예전 형식도 읽음) */
async function loadBigFromIdb(onProgress) {
  const idx = await idbGet(STORAGE_KEY);
  if (!idx) return null;
  if (idx.v !== 2) return idx.workbook ? idx : null;
  const sheets = [];
  for (let i = 0; i < idx.sheets.length; i++) {
    const x = idx.sheets[i];
    const rec = await idbGet(`${STORAGE_KEY}#${x.id}`);
    if (!rec) return null;
    const cells = new Map();
    for (const ch of rec.chunks) {
      for (const [k, d] of JSON.parse(await unpackChunk(ch, rec.gz))) cells.set(k, d);
      onProgress?.((i + 0.5) / idx.sheets.length);
      await yieldUI();
    }
    // 조각으로 나눠 저장한 열 블록을 다시 이어 붙임 (배열 하나를 한 번에 만들고 조각을 차례로 채움)
    const joinArr = async (info) => {
      const Ctor = { Float64Array, Int32Array, Uint32Array, Float32Array, Uint8Array, Int16Array, Uint16Array }[info.kind] ?? Float64Array;
      const out = new Ctor(info.len);
      let at = 0;
      for (const key of info.parts) {
        const part = await idbGet(key);
        if (!part) throw new Error('저장된 열 블록 조각이 없습니다');
        out.set(part, at);
        at += part.length;
        await yieldUI();
      }
      return out;
    };
    const blocks = [];
    for (const b of rec.blocks ?? []) {
      const cols = [];
      for (const c of b.cols) {
        const { numParts, strParts, ...rest } = c;
        cols.push({ ...rest, num: numParts ? await joinArr(numParts) : c.num, str: strParts ? await joinArr(strParts) : c.str });
      }
      const { permParts, ...rb } = b;
      blocks.push({ ...rb, cols, perm: permParts ? await joinArr(permParts) : b.perm ?? undefined });
      onProgress?.((i + 0.9) / idx.sheets.length);
    }
    sheets.push({ ...rec.meta, cells, blocks, _sid: x.id, _ev: x.ev });
  }
  return { rev: idx.rev, docName: idx.docName, si: idx.si, autosave: idx.autosave, workbook: { ...idx.book, names: idx.names, vba: idx.vba, sheets } };
}

function loadFromStorage() {
  try {
    const data = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null');
    return data?.workbook || data?.idb ? data : null;
  } catch {
    return null;
  }
}

let saveTimer = null;
let serverTimer = null;
/** 사용자가 쉬는 틈에 실행 (큰 문서 저장이 입력 · 슬라이서 클릭을 막지 않게) */
const whenIdle = (fn) => (globalThis.requestIdleCallback ? requestIdleCallback(fn, { timeout: 5000 }) : setTimeout(fn, 0));
function scheduleAutosave() {
  if (!autosave) { updateTitle(); return; }
  clearTimeout(saveTimer);
  const big = bigBook();
  saveTimer = setTimeout(() => (big ? whenIdle(saveToStorage) : saveToStorage()), big ? 4000 : 400);
  scheduleServerSave();
}

function scheduleServerSave(delay = 1500) {
  if (!server.available || !autosave || viewOnly) return;
  // 셀이 아주 많은 문서는 서버 자동 저장을 하지 않음 (전체를 보내야 해서 느림) — [저장]을 누르면 저장
  if (!serverAutosave()) { updateTitle(); return; }
  clearTimeout(serverTimer);
  const big = bigBook();
  serverTimer = setTimeout(() => (big ? whenIdle(() => saveNow(false)) : saveNow(false)), big ? Math.max(delay, 8000) : delay);
}

/** 저장: 브라우저 + (서버가 있으면) 서버 */
async function saveNow(explicit) {
  if (viewOnly) { if (explicit) toast('읽기 전용 보기입니다. [편집용 사본 만들기]를 누르세요.'); return; }
  const stored = saveToStorage();
  if (explicit) libraryFlush({ version: { label: '저장' } });
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
    autoRepublish();
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
    else alertDialog('WIXEL', `열 수 없습니다: ${err.message}`);
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
  const tplCard = (t) => el('button', { class: `bcard tpl${t.featured ? ' featured' : ''}`, title: t.desc, onclick: () => { close(); newWorkbook(t); } },
    el('div', { class: 'thumb tpl-thumb', style: { '--tc': t.color } }, el('i', {}), el('i', {}), el('i', {}), el('span', {}, t.name.slice(0, 1))),
    el('b', {}, t.name), el('small', {}, t.desc));
  const tplSearch = el('input', { type: 'search', class: 'tpl-search', placeholder: '온라인 템플릿 검색 (예: 예산, 리포트, 간트)' });
  const tplBox = el('div', {});
  const drawTpl = () => {
    const q = tplSearch.value.trim().toLowerCase();
    const hit = (t) => !q || `${t.name} ${t.desc} ${t.cat}`.toLowerCase().includes(q);
    tplBox.replaceChildren(...TEMPLATE_CATS.flatMap((cat) => {
      const list = TEMPLATES.filter((t) => t.cat === cat && hit(t));
      return list.length ? [el('h3', { class: 'tpl-cat' }, `${cat} 템플릿 (${list.length})`), el('div', { class: 'backstage-cards' }, list.map(tplCard))] : [];
    }));
  };
  tplSearch.addEventListener('input', drawTpl);
  drawTpl();
  const showNew = () => main.replaceChildren(
    el('h2', {}, '새로 만들기'),
    el('div', { class: 'backstage-cards' },
      el('button', { class: 'bcard', onclick: () => { close(); newWorkbook(); } }, el('div', { class: 'thumb' }), el('b', {}, '새 통합 문서'), el('small', {}, '빈 시트로 시작')),
      SAMPLES.map((s) => el('button', { class: 'bcard', onclick: () => { close(); newWorkbook(s); } }, el('div', { class: 'thumb' }), el('b', {}, s.name), el('small', {}, s.desc)))),
    el('div', { style: { margin: '26px 0 6px' } }, tplSearch), tplBox,
    el('h2', { style: { marginTop: '36px' } }, '정보'),
    el('div', { class: 'muted' }, `${docName} · 시트 ${wb.sheets.length}개 · ${server.available ? '서버에 저장 (다른 기기에서 열 수 있음)' : '이 브라우저에 저장'}${autosave ? ' · 자동 저장 켜짐' : ''}`),
  );
  const showOpen = async () => {
    const actions = el('div', { class: 'backstage-actions' },
      el('button', { class: 'btn primary', onclick: () => { close(); pickFile('open'); } }, '이 기기에서 찾아보기 (.xlsx · .xlsb · .xlsm · .ods · .csv · .wixel 등)'));
    main.replaceChildren(el('h2', {}, '열기'), actions, el('h3', { class: 'tpl-cat' }, '최근 항목 (이 브라우저)'), await recentTable(close));
    if (server.available) main.append(el('h3', { class: 'tpl-cat' }, '서버 (다른 기기와 공유)'));
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
    el('button', { onclick: () => { close(); versionHistory(); } }, '버전 기록'),
    el('button', { onclick: () => { close(); publishDialog(); } }, '공유 · 웹에 게시'),
    el('button', { onclick: showShare }, '다른 기기에서 열기'),
    el('button', { onclick: () => { close(); setTimeout(printSheet, 50); } }, '인쇄'),
    el('button', { onclick: () => { close(); optionsDialog(); } }, '옵션'),
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
  const pg = normPage(s.page);
  const u = wb.usedRange(si);
  // 인쇄 영역이 있으면 그 범위만
  const area = pg.area ?? { r1: 0, c1: 0, r2: Math.min(u.rows, 20000) - 1, c2: Math.min(u.cols, 200) - 1 };
  const r1 = area.r1;
  const c1 = area.c1;
  const r2 = Math.min(area.r2, r1 + 20000);
  const c2 = Math.min(area.c2, c1 + 200);
  const hf = (code) => headerParts(code, { file: docName, sheet: s.name });
  const hdr = pg.header ? hf(pg.header) : { left: '', center: `${docName} — ${s.name}`, right: '' };
  const parts = [`<div class="print-hf">${['left', 'center', 'right'].map((k) => `<span>${escapeHtml(hdr[k])}</span>`).join('')}</div>`];
  const colsList = range(c1, c2).filter((c) => !gv.cols.isHidden(c));
  const gridOn = pg.gridlines || view.printGrid;
  let tableW = 0;
  let tableH = 0;
  if (r2 >= r1 && c2 >= c1) {
    const colgroup = (pg.headings ? '<col style="width:34px">' : '') + colsList.map((c) => `<col style="width:${gv.cols.size(c)}px">`).join('');
    tableW = colsList.reduce((w, c) => w + gv.cols.size(c), pg.headings ? 34 : 0);
    const rowHtml = (r) => {
      const tds = [];
      if (pg.headings) tds.push(`<td class="ph">${r + 1}</td>`);
      for (const c of colsList) {
        const st = styleAt(r, c);
        const { text, align } = formatValue(valueAt(r, c), st);
        const css = [`text-align:${st.align || align}`, st.bold && 'font-weight:700', st.italic && 'font-style:italic', st.color && `color:${st.color}`,
          st.fill && `background:${st.fill}`, st.size && `font-size:${st.size}pt`, st.wrap && 'white-space:pre-wrap',
          st.bb && 'border-bottom:1px solid #000', st.bt && 'border-top:1px solid #000', st.bl && 'border-left:1px solid #000', st.br && 'border-right:1px solid #000'].filter(Boolean).join(';');
        tds.push(`<td style="${css}">${escapeHtml(text)}</td>`);
      }
      tableH += gv.rows.size(r);
      return `<tr style="height:${gv.rows.size(r)}px">${tds.join('')}</tr>`;
    };
    // 인쇄 제목(반복할 행)은 thead 로 — 브라우저가 쪽마다 반복
    const head = [];
    if (pg.headings) head.push(`<tr><td class="ph"></td>${colsList.map((c) => `<td class="ph">${colToName(c)}</td>`).join('')}</tr>`);
    const titleSet = new Set();
    if (pg.titleRows) for (let r = pg.titleRows[0]; r <= pg.titleRows[1]; r++) { titleSet.add(r); if (!gv.rows.isHidden(r)) head.push(rowHtml(r)); }
    const body = [];
    for (let r = r1; r <= r2; r++) if (!gv.rows.isHidden(r) && !titleSet.has(r)) body.push(rowHtml(r));
    const scale = printScale(pg, tableW, tableH);
    parts.push(`<table class="${gridOn ? 'grid-lines' : ''}" style="zoom:${scale};${pg.hCenter ? 'margin:0 auto;' : ''}"><colgroup>${colgroup}</colgroup>${head.length ? `<thead>${head.join('')}</thead>` : ''}<tbody>${body.join('')}</tbody></table>`);
  }
  if (!pg.area) {
    for (const ch of s.charts) parts.push(`<div class="chart-print">${gv.chartSvg(ch)}</div>`);
    for (const im of s.images ?? []) parts.push(`<div class="chart-print"><img src="${escapeHtml(im.src)}" style="width:${im.w}px;height:${im.h}px" alt=""></div>`);
    for (const sh of s.shapes ?? []) parts.push(`<div class="chart-print" style="position:relative;width:${sh.w}px;height:${Math.max(1, sh.h)}px">${shapeSvg(sh)}${sh.text ? `<div style="position:absolute;inset:0;display:flex;flex-direction:column;justify-content:center;padding:4px 8px;white-space:pre-wrap;text-align:${sh.align ?? 'center'};color:${escapeHtml(sh.color ?? '#000')};font-size:${sh.size ?? 11}pt">${escapeHtml(sh.text)}</div>` : ''}</div>`);
  }
  if (pg.footer) { const ft = hf(pg.footer); parts.push(`<div class="print-hf foot">${['left', 'center', 'right'].map((k) => `<span>${escapeHtml(ft[k])}</span>`).join('')}</div>`); }
  // 용지 · 방향 · 여백
  const m = pg.margins;
  let st = document.getElementById('pageStyle');
  if (!st) { st = document.createElement('style'); st.id = 'pageStyle'; document.head.append(st); }
  st.textContent = `@page { size: ${paperOf(pg.paper).css} ${pg.orientation}; margin: ${m.top}in ${m.right}in ${m.bottom}in ${m.left}in; }`;
  dom.printArea.innerHTML = parts.join('');
  window.print();
}

/** 페이지 설정 대화 상자 (용지 · 방향 · 여백 · 배율 · 인쇄 영역 · 인쇄 제목 · 머리글/바닥글) */
function pageSetupDialog() {
  const pg = normPage(sheet().page);
  const a1 = (rg) => (rg ? `${cellName(rg.r1, rg.c1)}:${cellName(rg.r2, rg.c2)}` : '');
  formDialog('페이지 설정', [
    { name: 'orientation', label: '용지 방향', type: 'select', value: pg.orientation, options: [{ value: 'portrait', label: '세로' }, { value: 'landscape', label: '가로' }] },
    { name: 'paper', label: '용지 크기', type: 'select', value: String(pg.paper), options: PAPERS.map((p) => ({ value: String(p.id), label: p.label })) },
    { name: 'scale', label: '확대/축소 배율 (%)', type: 'number', value: pg.scale },
    { name: 'fitW', label: '자동 맞춤: 용지 너비 (쪽, 0 = 사용 안 함)', type: 'number', value: pg.fitW },
    { name: 'fitH', label: '자동 맞춤: 용지 높이 (쪽, 0 = 제한 없음)', type: 'number', value: pg.fitH },
    { name: 'margins', label: '여백 (인치: 위, 아래, 왼쪽, 오른쪽)', value: [pg.margins.top, pg.margins.bottom, pg.margins.left, pg.margins.right].join(', ') },
    { name: 'hCenter', label: '페이지 가운데 맞춤: 가로', type: 'checkbox', value: pg.hCenter },
    { name: 'header', label: '머리글 (&L 왼쪽 &C 가운데 &R 오른쪽, &P 쪽 &N 전체 &D 날짜 &A 시트 &F 파일)', value: pg.header },
    { name: 'footer', label: '바닥글', value: pg.footer },
    { name: 'area', label: '인쇄 영역 (예: A1:H40, 비우면 전체)', value: a1(pg.area) },
    { name: 'titleRows', label: '반복할 행 (예: 1:2)', value: pg.titleRows ? `${pg.titleRows[0] + 1}:${pg.titleRows[1] + 1}` : '' },
    { name: 'gridlines', label: '눈금선 인쇄', type: 'checkbox', value: pg.gridlines },
    { name: 'headings', label: '행/열 머리글 인쇄', type: 'checkbox', value: pg.headings },
  ], (v) => {
    const mm = v.margins.split(/[,\s]+/).map(Number).filter((x) => Number.isFinite(x));
    const tr = /^\$?(\d+):\$?(\d+)$/.exec(v.titleRows.trim());
    const area = v.area.trim() ? parseRangeName(v.area.trim().replace(/\$/g, '')) : null;
    if (v.area.trim() && !area) { alertDialog('페이지 설정', '인쇄 영역 참조가 올바르지 않습니다.'); return false; }
    const next = {
      ...pg, orientation: v.orientation, paper: Number(v.paper), scale: Math.max(10, Math.min(400, Number(v.scale) || 100)),
      fitW: Math.max(0, Number(v.fitW) || 0), fitH: Math.max(0, Number(v.fitH) || 0), hCenter: v.hCenter, header: v.header, footer: v.footer,
      margins: mm.length === 4 ? { ...pg.margins, top: mm[0], bottom: mm[1], left: mm[2], right: mm[3] } : pg.margins,
      area, titleRows: tr ? [Number(tr[1]) - 1, Number(tr[2]) - 1] : null, gridlines: v.gridlines, headings: v.headings,
    };
    setPage(next);
    return undefined;
  });
}
function setPage(next) {
  wb.transact(() => wb.setSheetProp(si, 'page', next), meta());
  gv.renderAll();
}
const patchPage = (patch) => setPage({ ...normPage(sheet().page), ...patch });

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
  const [shrinkIn, shrinkL] = chk('셀에 맞춤 (글자 크기 자동 축소)', st.shrink);
  wrapIn.addEventListener('change', () => { if (wrapIn.checked) shrinkIn.checked = false; });
  shrinkIn.addEventListener('change', () => { if (shrinkIn.checked) wrapIn.checked = false; });
  // 방향: 각도(-90~90) 또는 세로 쓰기
  const [vertIn, vertL] = chk('세로 쓰기', st.rotate === 255);
  const rotIn = el('input', { type: 'number', min: -90, max: 90, value: st.rotate && st.rotate !== 255 ? st.rotate : 0, style: { width: '64px' } });
  const alignPage = col(el('div', { class: 'fc-title' }, '텍스트 맞춤'), row(lab('가로', hSel), lab('세로', vSel), lab('들여쓰기', indentIn)),
    el('div', { class: 'fc-title' }, '텍스트 조정'), wrapL, shrinkL, mergeL,
    el('div', { class: 'fc-title' }, '방향'), row(lab('각도(도)', rotIn), vertL));

  // ── 글꼴 ──
  // 글꼴: 이 PC의 글꼴 목록에서 고르거나 이름을 직접 입력
  // 엑셀처럼 입력 칸 + 전체 글꼴 목록 (입력하면 목록에서 찾아 줌)
  const fontSel = el('input', { type: 'text', value: st.font || BASE_FONT.name, spellcheck: false });
  const allFonts = [...new Set([BASE_FONT.name, '맑은 고딕', '나눔고딕', '돋움', '굴림', '바탕', '궁서', 'Calibri', 'Arial', 'Times New Roman', 'Segoe UI', 'Verdana', 'Tahoma', 'Consolas', ...fontList()])];
  const fontBox = el('select', { size: 7, class: 'fc-fontlist' }, allFonts.map((f) => el('option', { value: f, selected: f === fontSel.value, style: { fontFamily: fontStack(f) } }, f)));
  fontBox.addEventListener('change', () => { fontSel.value = fontBox.value; updFont(); });
  fontSel.addEventListener('input', () => {
    const q = fontSel.value.toLowerCase();
    const hit = allFonts.find((f) => f.toLowerCase().startsWith(q)) ?? allFonts.find((f) => f.toLowerCase().includes(q));
    if (hit) fontBox.value = hit;
  });
  const fontDl = canListLocalFonts() ? el('button', {
    type: 'button', class: 'btn', onclick: async () => {
      const list = await loadLocalFonts().catch(() => null);
      if (!list) return;
      for (const f of list) if (!allFonts.includes(f)) { allFonts.push(f); fontBox.append(el('option', { value: f, style: { fontFamily: fontStack(f) } }, f)); }
      toast(`이 PC의 글꼴 ${list.length}개를 불러왔습니다.`);
    },
  }, '이 PC의 글꼴 모두 보기') : null;
  const sizeIn = el('input', { type: 'number', min: 1, max: 409, value: st.size || BASE_FONT.size, style: { width: '64px' } });
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
  const fontPage = col(row(el('div', { class: 'fc-fontcol' }, lab('글꼴', fontSel), fontBox, fontDl), lab('크기', sizeIn)), row(bL, iL, uL, sL), lab('색', colorIn), el('div', { class: 'fc-title' }, '미리 보기'), fontPreview);

  // ── 테두리 (엑셀과 같은 구성: 선 스타일 · 색 · 미리 설정 · 가장자리별 단추 · 미리 보기) ──
  let border = null;
  const pen = { style: st.bbs ?? st.bts ?? 'thin', color: st.bbc ?? st.btc ?? '#000000' };
  const multiR = sel.r2 > sel.r1;
  const multiC = sel.c2 > sel.c1;
  const edges = { top: !!st.bt, bottom: !!st.bb, left: !!st.bl, right: !!st.br, insideH: false, insideV: false, diagUp: !!st.du, diagDown: !!st.dd };
  const edges0 = { ...edges };
  const styleList = el('div', { class: 'fc-linestyles' });
  const LS = [['none', '없음'], ...BORDER_STYLES];
  const lsCss = { thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', thick: '3px solid', double: '3px double' };
  for (const [v, label] of LS) {
    const b = el('button', { type: 'button', class: `fc-ls${pen.style === v ? ' on' : ''}`, title: label }, v === 'none' ? '없음' : el('i', { style: { borderTop: `${lsCss[v]} #333` } }));
    b.addEventListener('click', () => { pen.style = v; styleList.querySelectorAll('.on').forEach((x) => x.classList.remove('on')); b.classList.add('on'); });
    styleList.append(b);
  }
  const penColor = el('input', { type: 'color', value: pen.color });
  penColor.addEventListener('input', () => { pen.color = penColor.value; drawPrev(); });
  const prev = el('div', { class: 'fc-bprev' });
  const drawPrev = () => {
    const line = (on) => (on ? `${lsCss[pen.style] ?? '1px solid'} ${pen.color}` : '1px dashed #e0e0e0');
    prev.replaceChildren(el('div', {
      class: 'fc-bbox', style: { borderTop: line(edges.top), borderBottom: line(edges.bottom), borderLeft: line(edges.left), borderRight: line(edges.right) },
    },
    multiR ? el('i', { class: 'fc-mid-h', style: { borderTop: line(edges.insideH) } }) : null,
    multiC ? el('i', { class: 'fc-mid-v', style: { borderLeft: line(edges.insideV) } }) : null,
    edges.diagDown ? el('i', { class: 'fc-diag down', style: { background: `linear-gradient(to top right, transparent calc(50% - 1px), ${pen.color} 50%, transparent calc(50% + 1px))` } }) : null,
    edges.diagUp ? el('i', { class: 'fc-diag up', style: { background: `linear-gradient(to bottom right, transparent calc(50% - 1px), ${pen.color} 50%, transparent calc(50% + 1px))` } }) : null,
    el('span', {}, '텍스트')));
  };
  const edgeBtn = (k, label, ic) => {
    const b = el('button', { type: 'button', class: `fc-bbtn small${edges[k] ? ' on' : ''}`, title: label, html: `${ICONS[ic] ?? ''}` });
    b.addEventListener('click', () => { edges[k] = !edges[k]; b.classList.toggle('on', edges[k]); border = 'edges'; drawPrev(); });
    return b;
  };
  const presetBtn = (k, label, ic) => {
    const b = el('button', { type: 'button', class: 'fc-bbtn', title: label, html: `${ICONS[ic] ?? ''}<span>${label}</span>` });
    b.addEventListener('click', () => {
      if (k === 'none') Object.keys(edges).forEach((e) => { edges[e] = false; });
      if (k === 'outside') Object.assign(edges, { top: true, bottom: true, left: true, right: true });
      if (k === 'inside') Object.assign(edges, { insideH: multiR, insideV: multiC });
      border = 'edges';
      edgeRow.querySelectorAll('button').forEach((x) => x.classList.toggle('on', !!edges[x.dataset.k]));
      drawPrev();
    });
    return b;
  };
  const edgeRow = el('div', { class: 'fc-edges' });
  for (const [k, label, ic] of [['top', '위쪽', 'borderTop'], ['insideH', '안쪽 가로', 'borderBottom'], ['bottom', '아래쪽', 'borderBottom'], ['diagUp', '대각선 ↗', 'borderAll'], ['left', '왼쪽', 'borderLeft'], ['insideV', '안쪽 세로', 'borderLeft'], ['right', '오른쪽', 'borderRight'], ['diagDown', '대각선 ↘', 'borderAll']]) {
    const b = edgeBtn(k, label, ic);
    b.dataset.k = k;
    edgeRow.append(b);
  }
  drawPrev();
  const borderPage = el('div', { class: 'fc-border-page' },
    el('div', {}, el('div', { class: 'fc-title' }, '선 스타일'), styleList, el('div', { class: 'fc-title' }, '색'), penColor),
    el('div', {}, el('div', { class: 'fc-title' }, '미리 설정'), el('div', { class: 'fc-row' }, presetBtn('none', '없음', 'borderNone'), presetBtn('outside', '윤곽선', 'borderOutside'), presetBtn('inside', '안쪽', 'borderAll')),
      el('div', { class: 'fc-title' }, '테두리'), el('div', { class: 'fc-bwrap' }, edgeRow, prev),
      el('div', { class: 'muted' }, '선 스타일과 색을 고른 다음 미리 설정이나 가장자리 단추를 누르세요.')));

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
  // 무늬 스타일 · 무늬 색 (엑셀 채우기 탭)
  const patSel = el('select', {}, [el('option', { value: '' }, '(무늬 없음)'), ...PATTERNS.map(([v, l]) => el('option', { value: v, selected: st.pattern === v }, l))]);
  const patColor = el('input', { type: 'color', value: st.patternColor || '#000000' });
  const fillPrev = el('div', { class: 'fc-fillprev' });
  const updFill = () => { fillPrev.style.background = patSel.value ? patternCss(patSel.value, patColor.value, noFill.checked ? '#ffffff' : fillIn.value) : noFill.checked ? '#ffffff' : fillIn.value; };
  [patSel, patColor, fillIn, noFill].forEach((x) => x.addEventListener('input', updFill));
  swatches.addEventListener('click', () => setTimeout(updFill, 0));
  updFill();
  const fillPage = col(noFillL, el('div', { class: 'fc-title' }, '배경색'), swatches, lab('다른 색', fillIn),
    row(lab('무늬 스타일', patSel), lab('무늬 색', patColor)), el('div', { class: 'fc-title' }, '보기'), fillPrev);

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
            font: fontSel.value === BASE_FONT.name ? undefined : fontSel.value,
            size: !size || size === BASE_FONT.size ? undefined : Math.min(409, size),
            bold: bIn.checked || undefined, italic: iIn.checked || undefined, underline: uIn.checked || undefined, strike: sIn.checked || undefined,
            color: colorIn.value === '#000000' ? undefined : colorIn.value,
            fill: noFill.checked ? undefined : fillIn.value,
            pattern: patSel.value || undefined, patternColor: patSel.value && patColor.value !== '#000000' ? patColor.value : undefined,
            shrink: shrinkIn.checked || undefined,
            rotate: vertIn.checked ? 255 : Number(rotIn.value) ? Math.max(-90, Math.min(90, Number(rotIn.value))) : undefined,
          };
          wb.transact(() => {
            applyStyle(patch, { widen: fmt.numFmt !== undefined ? 'grow' : false });
            if (border === 'edges') applyEdges(edges, edges0, pen);
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
    wb.transact(() => anchorObjects(() => {
      if (isCol) for (let c = sel.c1; c <= Math.min(sel.c2, sel.c1 + 1000); c++) wb.setColWidth(si, c, n);
      else for (let r = sel.r1; r <= Math.min(sel.r2, sel.r1 + 5000); r++) wb.setRowHeight(si, r, n);
    }), meta());
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
  formDialog('이동', [{ name: 'ref', label: '참조 (비워 두고 확인하면 이동 옵션)', value: '' }], ({ ref }) => (ref.trim() ? gotoRef(ref) : setTimeout(gotoSpecialDialog, 0)));
}


// ───────────────────────── 이름 정의 ─────────────────────────
const absRef = (rg) => {
  const a = `$${colToName(rg.c1)}$${rg.r1 + 1}`;
  return rg.r1 === rg.r2 && rg.c1 === rg.c2 ? a : `${a}:$${colToName(rg.c2)}$${rg.r2 + 1}`;
};
const selRefText = (rg = sel, s = si) => `=${quoteSheetName(wb.sheets[s].name)}!${absRef(rg)}`;

/** 엑셀 이름 규칙: 문자/밑줄/\ 로 시작, 셀 주소·R1C1 모양 불가 */
function nameError(name, except = null) {
  const n = name.trim();
  if (!n) return '이름을 입력하세요.';
  if (n.length > 255) return '이름이 너무 깁니다.';
  if (!/^[A-Za-z_\\À-￿][\w.?\\À-￿]*$/.test(n) || /^[A-Za-z]{1,3}\d+$/.test(n) || /^(R|C|RC|R\d+C\d+|TRUE|FALSE)$/i.test(n)) {
    return '입력한 이름이 올바르지 않습니다. 이름은 문자나 밑줄로 시작해야 하며 공백이나 셀 주소 모양은 쓸 수 없습니다.';
  }
  return null;
}

function nameValueText(n) {
  let v;
  try { v = wb.nameValue(n.name, n.sheet ? wb.sheetIndexByName(n.sheet) : si, n.sheet, null); } catch { v = null; }
  if (v === undefined || v === null) return '';
  if (isError(v)) return v.code;
  if (v.constructor?.name === 'RefValue') {
    const s = wb.sheetIndexByName(v.sheet ?? sheet().name);
    const vals = [];
    for (let r = v.r1; r <= Math.min(v.r2, v.r1 + 3); r++) {
      for (let c = v.c1; c <= Math.min(v.c2, v.c1 + 3); c++) {
        const x = s >= 0 ? wb.getValue(s, r, c) : null;
        vals.push(x === null ? '' : isError(x) ? x.code : typeof x === 'number' ? formatGeneral(x) : String(x));
      }
    }
    const one = v.r1 === v.r2 && v.c1 === v.c2;
    return one ? vals[0] : `{${vals.map((x) => `"${x}"`).join(';')}${(v.r2 - v.r1 + 1) * (v.c2 - v.c1 + 1) > vals.length ? ';…' : ''}}`;
  }
  if (v.params) return 'LAMBDA';
  if (v.rows) return `{${v.rows.slice(0, 3).map((r) => r.slice(0, 4).join(',')).join(';')}}`;
  return typeof v === 'number' ? formatGeneral(v) : String(v);
}

/** 이름 편집기 (새 이름 · 편집) */
function nameEditor(entry, list, onSave) {
  const scopes = [{ value: '', label: '통합 문서' }, ...wb.sheets.map((sh) => ({ value: sh.name, label: sh.name }))];
  formDialog(entry ? '이름 편집' : '새 이름', [
    { name: 'name', label: '이름', value: entry?.name ?? '' },
    { name: 'sheet', label: '범위', type: 'select', value: entry?.sheet ?? '', options: scopes },
    { name: 'comment', label: '설명', type: 'textarea', value: entry?.comment ?? '' },
    { name: 'ref', label: '참조 대상', value: entry?.ref ?? selRefText() },
  ], (v) => {
    const err = nameError(v.name);
    if (err) { toast(err); return false; }
    const dup = list.find((x) => x !== entry && x.name.toLowerCase() === v.name.trim().toLowerCase() && (x.sheet ?? '') === v.sheet);
    if (dup) { toast('같은 범위에 이미 있는 이름입니다.'); return false; }
    let ref = v.ref.trim();
    if (!ref.startsWith('=')) ref = `=${ref}`;
    try { parse(ref.slice(1)); } catch { toast('참조 대상 수식에 문제가 있습니다.'); return false; }
    onSave({ name: v.name.trim(), sheet: v.sheet || null, comment: v.comment || undefined, ref: normalizeFormula(ref), hidden: entry?.hidden });
    return true;
  });
}

function defineName() {
  if (editing && !commitEdit()) return;
  nameEditor(null, wb.names, (n) => wb.transact(() => wb.setNames([...wb.names, n]), meta()));
}

function nameManager() {
  if (editing && !commitEdit()) return;
  let list = wb.names.map((x) => ({ ...x }));
  let current = null;
  const tbody = el('tbody');
  const refIn = el('input', { type: 'text', style: { flex: '1' } });
  const filterIn = el('input', { type: 'text', placeholder: '이름 검색', style: { width: '160px' } });
  const commit = () => {
    const out = list.filter((x) => !x.hidden || true).map(({ _ast, _text, ...x }) => x);
    if (JSON.stringify(out) === JSON.stringify(wb.names.map(({ _ast, _text, ...x }) => x))) return;
    wb.transact(() => wb.setNames(out), meta());
  };
  const render = () => {
    tbody.replaceChildren();
    const q = filterIn.value.trim().toLowerCase();
    const vis = list.filter((x) => !x.hidden && (!q || x.name.toLowerCase().includes(q)));
    if (!vis.includes(current)) current = vis[0] ?? null;
    if (!vis.length) tbody.append(el('tr', {}, el('td', { colspan: 5, class: 'muted', style: { padding: '12px' } }, '정의된 이름이 없습니다. [새로 만들기]를 누르세요.')));
    for (const n of vis) {
      const tr = el('tr', { class: n === current ? 'on' : '' },
        el('td', {}, n.name), el('td', {}, nameValueText(n)), el('td', {}, n.ref), el('td', {}, n.sheet ?? '통합 문서'), el('td', {}, n.comment ?? ''));
      tr.addEventListener('mousedown', () => { current = n; render(); });
      tr.addEventListener('dblclick', () => edit());
      tbody.append(tr);
    }
    refIn.value = current?.ref ?? '';
    refIn.disabled = !current;
    btnEdit.disabled = btnDel.disabled = !current;
  };
  const edit = () => {
    if (!current) return;
    const target = current;
    nameEditor(target, list, (n) => { Object.assign(target, n); render(); });
  };
  const btn = (label, onclick) => el('button', { type: 'button', class: 'btn', onclick }, label);
  const btnNew = btn('새로 만들기...', () => nameEditor(null, list, (n) => { list.push(n); current = n; render(); }));
  const btnEdit = btn('편집...', edit);
  const btnDel = btn('삭제', () => { list = list.filter((x) => x !== current); current = null; render(); });
  refIn.addEventListener('change', () => {
    if (!current) return;
    let t = refIn.value.trim();
    if (!t.startsWith('=')) t = `=${t}`;
    try { parse(t.slice(1)); current.ref = normalizeFormula(t); refIn.classList.remove('bad'); } catch { refIn.classList.add('bad'); }
    render();
  });
  filterIn.addEventListener('input', render);
  const table = el('table', { class: 'cf-table' },
    el('thead', {}, el('tr', {}, el('th', {}, '이름'), el('th', {}, '값'), el('th', {}, '참조 대상'), el('th', {}, '범위'), el('th', {}, '설명'))), tbody);
  const body = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } },
    el('div', { style: { display: 'flex', gap: '6px', alignItems: 'center' } }, btnNew, btnEdit, btnDel, el('span', { style: { flex: '1' } }), filterIn),
    el('div', { class: 'cf-tablewrap' }, table),
    el('label', { style: { display: 'flex', gap: '6px', alignItems: 'center' } }, el('span', {}, '참조 대상:'), refIn));
  render();
  openDialog({ title: '이름 관리자', body, width: 760, buttons: [{ label: '닫기', primary: true, action: commit }] });
}

/** 수식에 이름 붙여넣기 (F3) */
function pasteNameDialog() {
  const names = wb.names.filter((n) => !n.hidden && (!n.sheet || n.sheet.toLowerCase() === sheet().name.toLowerCase()));
  if (!names.length) { toast('정의된 이름이 없습니다.'); return; }
  const listEl = el('select', { size: 10, style: { width: '100%' } }, names.map((n, i) => el('option', { value: n.name, selected: i === 0 }, n.name)));
  const insert = () => {
    const name = listEl.value;
    if (!name) return;
    if (editing) {
      const inp = edInput();
      const pos = inp.selectionStart;
      setEditText(`${inp.value.slice(0, pos)}${name}${inp.value.slice(inp.selectionEnd)}`, pos + name.length);
    } else setTimeout(() => startEdit('enter', `=${name}`)); // Enter 키 입력이 편집기로 들어가지 않도록
  };
  const pasteList = () => {
    // 목록 붙여넣기: 이름과 참조 대상을 현재 셀부터 두 열로
    wb.transact(() => names.forEach((n, i) => {
      wb.setInput(si, active.r + i, active.c, n.name);
      wb.setInput(si, active.r + i, active.c + 1, `'${n.ref}`);
    }), meta());
  };
  listEl.addEventListener('dblclick', () => { insert(); document.querySelector('.dialog-backdrop:last-child .dialog-head button')?.click(); });
  openDialog({
    title: '이름 붙여넣기', width: 320, body: el('div', {}, el('div', { class: 'muted', style: { marginBottom: '6px' } }, '이름 붙여넣기(N)'), listEl),
    buttons: [{ label: '확인', primary: true, action: insert }, ...(editing ? [] : [{ label: '목록 붙여넣기', action: pasteList }]), { label: '취소' }],
  });
}

/** 선택 영역에서 이름 만들기 (Ctrl+Shift+F3) */
function createNamesFromSel() {
  if (editing && !commitEdit()) return;
  const rg = usedClip(sel);
  const top = el('input', { type: 'checkbox', checked: rg.r2 > rg.r1 && typeof wb.getValue(si, rg.r1, rg.c1 + (rg.c2 > rg.c1 ? 1 : 0)) === 'string' });
  const left = el('input', { type: 'checkbox', checked: rg.c2 > rg.c1 && !top.checked });
  const bottom = el('input', { type: 'checkbox' });
  const right = el('input', { type: 'checkbox' });
  const row = (box, label) => el('label', { style: { display: 'flex', gap: '6px' } }, box, label);
  openDialog({
    title: '선택 영역에서 이름 만들기', width: 320,
    body: el('div', { style: { display: 'flex', flexDirection: 'column', gap: '6px' } }, el('div', {}, '다음 위치의 값을 이름으로 사용:'),
      row(top, '첫 행(T)'), row(left, '왼쪽 열(L)'), row(bottom, '마지막 행(B)'), row(right, '오른쪽 열(R)')),
    buttons: [{
      label: '확인', primary: true, action: () => {
        const made = [];
        const clean = (v) => String(v ?? '').trim().replace(/[^\w.\\À-￿]+/g, '_').replace(/^(\d)/, '_$1');
        const add = (label, area) => {
          const nm = clean(label);
          if (!nm || nameError(nm)) return;
          made.push({ name: nm, sheet: null, ref: selRefText(area) });
        };
        const r1 = rg.r1 + (top.checked ? 1 : 0);
        const r2 = rg.r2 - (bottom.checked ? 1 : 0);
        const c1 = rg.c1 + (left.checked ? 1 : 0);
        const c2 = rg.c2 - (right.checked ? 1 : 0);
        if (r1 > r2 || c1 > c2) return;
        for (let c = c1; c <= c2; c++) {
          if (top.checked) add(wb.getValue(si, rg.r1, c), { r1, c1: c, r2, c2: c });
          if (bottom.checked) add(wb.getValue(si, rg.r2, c), { r1, c1: c, r2, c2: c });
        }
        for (let r = r1; r <= r2; r++) {
          if (left.checked) add(wb.getValue(si, r, rg.c1), { r1: r, c1, r2: r, c2 });
          if (right.checked) add(wb.getValue(si, r, rg.c2), { r1: r, c1, r2: r, c2 });
        }
        if (!made.length) { toast('만들 이름이 없습니다.'); return; }
        const keep = wb.names.filter((x) => !made.some((m) => m.name.toLowerCase() === x.name.toLowerCase() && !x.sheet));
        wb.transact(() => wb.setNames([...keep, ...made]), meta());
        toast(`이름 ${made.length}개를 만들었습니다.`);
      },
    }, { label: '취소' }],
  });
}

/** 이름 상자 드롭다운 목록 갱신 */
function refreshNameList() {
  let dl = document.getElementById('nameList');
  if (!dl) { dl = el('datalist', { id: 'nameList' }); document.body.append(dl); dom.nameBox.setAttribute('list', 'nameList'); }
  const items = [...wb.names.filter((n) => !n.hidden).map((n) => n.name), ...wb.sheets.flatMap((s) => (s.tables ?? []).map((t) => t.name))];
  const key = items.join('\u0001');
  if (dl.dataset.key === key) return;
  dl.dataset.key = key;
  dl.replaceChildren(...items.map((n) => el('option', { value: n })));
}

/** 이름 상자 Enter: 주소 → 이동, 정의된 이름 → 그 범위 선택, 새 이름 → 선택 영역에 이름 정의 */
/** 선택 영역이 표 전체(또는 데이터 영역)와 같으면 그 표 */
function selectedTable() {
  if (selKind !== 'cells') return null;
  return (sheet().tables ?? []).find((t) => sel.c1 === t.c1 && sel.c2 === t.c2 && sel.r2 === t.r2 && (sel.r1 === t.r1 || sel.r1 === dataTop(t))) ?? null;
}

/** 이름 상자에 보일 글자: 표 전체를 고르면 표 이름, 이름 정의와 같은 범위면 그 이름 (엑셀과 같음) */
function nameBoxLabel() {
  const t = selectedTable();
  if (t) return t.name;
  if (!isSingle(sel) && selKind === 'cells') {
    const here = selRefText().replace(/\$/g, '').toLowerCase();
    const n = wb.names.find((x) => !x.hidden && (!x.sheet || x.sheet.toLowerCase() === sheet().name.toLowerCase()) && String(x.ref).replace(/[=$]/g, '').toLowerCase() === here);
    if (n) return n.name;
  }
  return cellName(active.r, active.c);
}

function nameBoxEnter(text) {
  const t = text.trim();
  if (!t) return false;
  const existing = wb.findName(t, si);
  // 표 전체를 고른 상태에서 새 이름을 입력하면 표 이름 바꾸기 (엑셀과 같음)
  const selT = selectedTable();
  if (selT && !existing && t.toLowerCase() !== selT.name.toLowerCase() && !findTable(wb, t) && validTableName(t)) {
    renameTable(t);
    toast(`표 이름을 '${t}'(으)로 바꿨습니다.`);
    return true;
  }
  const isTable = !existing && wb.sheets.some((s) => (s.tables ?? []).some((x) => x.name.toLowerCase() === t.toLowerCase()));
  if (existing || isTable) {
    const v = wb.nameValue(t, si, null, null);
    if (v && v.constructor?.name === 'RefValue') {
      const s = v.sheet ? wb.sheetIndexByName(v.sheet) : si;
      if (s >= 0) {
        switchSheet(s);
        const rg = { r1: v.r1, c1: v.c1, r2: v.r2, c2: v.c2 };
        if (isSingle(rg)) selectCell(rg.r1, rg.c1); else { growTo(rg.r2, rg.c2); gv.ensureVisible(rg.r1, rg.c1); selectRange(rg); }
        return true;
      }
    }
    toast('이 이름은 셀 범위를 가리키지 않습니다.');
    return false;
  }
  const looksRef = /!|^\$?[A-Za-z]{1,3}\$?\d+(:\$?[A-Za-z]{1,3}\$?\d+)?$|^\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}$|^\$?\d+:\$?\d+$/.test(t);
  if (!looksRef && !nameError(t)) {
    wb.transact(() => wb.setNames([...wb.names, { name: t, sheet: null, ref: selRefText() }]), meta());
    toast(`이름 '${t}'을(를) 정의했습니다.`);
    return true;
  }
  return gotoRef(t);
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
    addCondRuleHere(rule);
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
      visual: { type: 'scale', colors: [...DEFAULT_SCALE3] }, contains: { type: 'gt', v1: '' },
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
          if (v === 'scale2') Object.assign(rule, { type: 'scale', colors: [...DEFAULT_SCALE2] });
          else if (v === 'scale3') Object.assign(rule, { type: 'scale', colors: [...DEFAULT_SCALE3] });
          else if (v === 'bar') Object.assign(rule, { type: 'bar', color: '#638ec6' });
          else Object.assign(rule, { type: 'icons', icons: '3Arrows' });
          render();
        });
      detail.append(row(el('span', {}, '서식 스타일:'), style));
      // 최소 · 중간 · 최대 기준 (엑셀의 [종류] · [값]): 최소값/최대값 · 숫자 · 백분율 · 백분위수 · 수식
      const CFVO_TYPES = [['min', '최소값'], ['num', '숫자'], ['percent', '백분율'], ['percentile', '백분위수'], ['formula', '수식'], ['max', '최대값']];
      const cfvoEditor = (i, n, label) => {
        const dflt = i === 0 ? { type: 'min' } : i === n - 1 ? { type: 'max' } : { type: 'percentile', v: 50 };
        rule.cfvo ??= Array.from({ length: n }, (_, k) => (k === 0 ? { type: 'min' } : k === n - 1 ? { type: 'max' } : { type: 'percentile', v: 50 }));
        const p = rule.cfvo[i] ?? (rule.cfvo[i] = dflt);
        const valIn = el('input', { type: 'text', value: p.v ?? '', style: { width: '80px' }, disabled: p.type === 'min' || p.type === 'max' });
        valIn.addEventListener('input', () => { p.v = valIn.value; prevBox.replaceChildren(cfPreview(rule)); });
        const typeSel = sel(CFVO_TYPES, p.type, (v) => { p.type = v; valIn.disabled = v === 'min' || v === 'max'; if (!valIn.disabled && valIn.value === '') { valIn.value = v === 'percentile' || v === 'percent' ? '50' : '0'; p.v = valIn.value; } });
        return el('div', { class: 'cf-cfvo' }, el('b', {}, label), typeSel, valIn);
      };
      if (rule.type === 'scale') {
        const names = rule.colors.length === 3 ? ['최소값', '중간값', '최대값'] : ['최소값', '최대값'];
        if (rule.cfvo && rule.cfvo.length !== rule.colors.length) delete rule.cfvo;
        detail.append(el('div', { class: 'cf-cfvo-grid' }, ...rule.colors.map((c, i) => el('div', {}, cfvoEditor(i, rule.colors.length, names[i]),
          el('label', { class: 'fc-field' }, el('span', {}, '색'), colorIn(c, (v) => { rule.colors[i] = v; prevBox.replaceChildren(cfPreview(rule)); }))))));
      } else if (rule.type === 'bar') {
        if (rule.cfvo && rule.cfvo.length !== 2) delete rule.cfvo;
        detail.append(
          el('div', { class: 'cf-cfvo-grid' }, cfvoEditor(0, 2, '최소값'), cfvoEditor(1, 2, '최대값')),
          row(el('span', {}, '막대 색:'), colorIn(rule.color ?? '#638ec6', (v) => { rule.color = v; prevBox.replaceChildren(cfPreview(rule)); }),
            sel([['grad', '그라데이션 채우기'], ['solid', '단색 채우기']], rule.gradient === false ? 'solid' : 'grad', (v) => { rule.gradient = v === 'solid' ? false : undefined; prevBox.replaceChildren(cfPreview(rule)); })),
          row(el('span', {}, '음수 막대 색:'), colorIn(rule.negColor ?? '#ff0000', (v) => { rule.negColor = v; }), chk('막대만 표시', rule.iconOnly, (on) => { rule.iconOnly = on || undefined; })),
        );
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
  // 피벗 테이블: 엑셀의 '규칙 적용 대상' 3가지 (선택한 셀 · 값 필드의 모든 셀 · 행 필드 수준의 모든 셀)
  const pscope = pivotScopeUi(rule);
  body.append(...(pscope ? [pscope.node] : []), el('div', { class: 'fc-title' }, '규칙 유형 선택'), kindList, detail, fmtBox);
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
          if (pscope) pscope.apply(rule);
          onSave(rule);
          return undefined;
        },
      },
      { label: '취소' },
    ],
  });
}

const rangeText1 = (r) => `$${colToName(r.c1)}$${r.r1 + 1}${r.r1 === r.r2 && r.c1 === r.c2 ? '' : `:$${colToName(r.c2)}$${r.r2 + 1}`}`;
const rangeText = (r) => `=${[r, ...(r.more ?? [])].map(rangeText1).join(',')}`;

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
  const hits = (rl) => [rl, ...(rl.more ?? [])].some((g) => g.r1 <= selRange.r2 && g.r2 >= selRange.r1 && g.c1 <= selRange.c2 && g.c2 >= selRange.c1);
  const visible = () => {
    const list = listOf(scopeSheet);
    return list.filter((rl) => scope !== 'selection' || hits(rl));
  };
  // 선택 영역에 규칙이 없고 시트에는 있으면 엑셀 파일의 규칙이 바로 보이도록 '현재 워크시트'로 시작
  if (!listOf(si).some(hits) && listOf(si).length) { scope = 'sheet'; showSel.value = `sheet:${si}`; }
  const render = () => {
    tbody.replaceChildren();
    const vis = visible();
    if (!vis.includes(current)) current = vis[0] ?? null;
    if (!vis.length) tbody.append(el('tr', {}, el('td', { colspan: 4, class: 'muted', style: { padding: '12px' } }, '표시할 규칙이 없습니다. [새 규칙]을 눌러 만드세요.')));
    for (const rl of vis) {
      const rangeIn = el('input', { type: 'text', value: rangeText(rl), class: 'cf-range' });
      rangeIn.addEventListener('change', () => {
        // 여러 범위는 쉼표로 구분 (=$A$1:$A$9,$C$1:$C$9)
        const ps = rangeIn.value.replace(/[=$]/g, '').split(/[,\s]+/).filter(Boolean).map((x) => parseRangeName(x));
        if (!ps.length || ps.some((x) => !x)) { rangeIn.classList.add('bad'); return; }
        rangeIn.classList.remove('bad');
        Object.assign(rl, { r1: ps[0].r1, c1: ps[0].c1, r2: ps[0].r2, c2: ps[0].c2 });
        if (ps.length > 1) rl.more = ps.slice(1); else delete rl.more;
      });
      const stop = el('input', { type: 'checkbox', checked: !!rl.stopIfTrue, disabled: VISUAL_TYPES.has(rl.type) });
      stop.addEventListener('change', () => { rl.stopIfTrue = stop.checked || undefined; });
      const tr = el('tr', { class: rl === current ? 'on' : '' },
        el('td', {}, describeCond(rl)), el('td', {}, cfPreview(rl)), el('td', {}, rangeIn, pivotScopeSelect(rl, scopeSheet, () => { rangeIn.value = rangeText(rl); })), el('td', { style: { textAlign: 'center' } }, stop));
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
      const keepRange = !(nr.pivot && nr.pivot.scope !== 'selection');
      const was = { r1: target.r1, c1: target.c1, r2: target.r2, c2: target.c2, more: target.more };
      for (const k of Object.keys(target)) if (!['r1', 'c1', 'r2', 'c2', 'stopIfTrue'].includes(k)) delete target[k];
      Object.assign(target, nr, keepRange ? was : {});
      if (keepRange && !was.more) delete target.more;
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
  ['F4', '수식 편집 중: 절대/상대 참조 전환 · 그 밖에는 마지막 작업 반복'],
  ['Esc', '편집 취소 / 복사 영역 해제'],
  ['Ctrl+방향키', '데이터 영역의 끝으로 이동 (빈 열에서는 20,000,000행까지)'],
  ['Shift+방향키 / F8', '선택 영역 확장 / 확장 모드 켜기·끄기'],
  ['Ctrl+Space / Shift+Space', '열 전체 / 행 전체 선택'],
  ['Ctrl+A / Ctrl+Shift+*', '모두 선택 / 현재 영역 선택'],
  ['Ctrl+Home / Ctrl+End', '처음 셀 / 마지막 셀로 이동'],
  ['Ctrl+. / Shift+Backspace / Ctrl+Backspace', '선택 영역의 다음 모서리 / 활성 셀만 선택 / 활성 셀로 화면 이동'],
  ['Alt+PageDown / Alt+PageUp', '한 화면 오른쪽 / 왼쪽으로'],
  ['Ctrl+C / Ctrl+X / Ctrl+V', '복사 / 잘라내기 / 붙여넣기'],
  ['Ctrl+Alt+V', '선택하여 붙여넣기 (값·서식·연산·행/열 바꿈 …)'],
  ['Ctrl+Z / Ctrl+Y', '실행 취소 / 다시 실행 (다시 실행할 것이 없으면 마지막 작업 반복)'],
  ['Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5', '굵게 / 기울임꼴 / 밑줄 / 취소선'],
  ['Ctrl+D / Ctrl+R', '아래로 / 오른쪽으로 채우기'],
  ['Ctrl+E', '빠른 채우기 (예시를 보고 나머지 행 자동 채우기)'],
  ["Ctrl+' / Ctrl+Shift+\"", '위 셀의 수식 / 값 복사'],
  ['Alt+=', '자동 합계'],
  ['Alt+F1 / F11', '차트 삽입'],
  ['Ctrl+Shift+L / Ctrl+Alt+L', '필터 켜기/끄기 / 필터 다시 적용'],
  ['Ctrl+9 / Ctrl+0', '행 숨기기 / 열 숨기기'],
  ['Ctrl+Shift+9 / Ctrl+Shift+0', '행 숨기기 취소 / 열 숨기기 취소'],
  ['Ctrl+;  /  Ctrl+Shift+;', '오늘 날짜 / 현재 시간 입력'],
  ['Ctrl+Shift+1/2/3/4/5/6', '숫자 / 시간 / 날짜 / 통화 / 백분율 / 지수 서식'],
  ['Ctrl+Shift+~', '일반 서식'],
  ['Ctrl+Shift+& / Ctrl+Shift+_', '바깥쪽 테두리 / 테두리 없음'],
  ['Ctrl+1', '셀 서식 (표시 형식 · 사용자 지정 서식 · 맞춤 · 글꼴 · 테두리 · 채우기)'],
  ['Ctrl+Shift+F / Ctrl+Shift+P', '글꼴 서식'],
  ['Ctrl+K', '하이퍼링크 삽입'],
  ['Ctrl+T / Ctrl+L', '표 만들기 (브라우저가 Ctrl+T 를 가로채면 Ctrl+L)'],
  ['Ctrl+Shift+T', '표 요약 행 켜기/끄기'],
  ['Alt+A+E', '텍스트 나누기'],
  ['Alt → 글자', '리본 바로 가기 키 (예: Alt+N+T 표, Alt+N+V 피벗, Alt+H+O+E 셀 서식, Alt+H+V+V 값 붙여넣기, Alt+E+S 선택하여 붙여넣기)'],
  ['Alt+↓', '목록·필터·요약 함수 펼치기'],
  ['Alt+C / Alt+S', '슬라이서: 필터 지우기 / 다중 선택'],
  ['Alt+F5 / Ctrl+Alt+F5', '피벗 테이블 새로 고침 / 모두 새로 고침'],
  ['Ctrl+F / Ctrl+H / Ctrl+G · F5', '찾기 / 바꾸기 / 이동'],
  ['Shift+F5', '찾기'],
  ['Ctrl+F3 / F3 / Ctrl+Shift+F3', '이름 관리자 / 이름 붙여넣기 / 선택 영역에서 이름 만들기'],
  ['Shift+F3', '함수 삽입'],
  ['Ctrl+[ / Ctrl+]', '참조되는 셀 / 참조하는 셀 선택'],
  ['Ctrl+Shift+O', '메모가 있는 셀 선택'],
  ['Ctrl+`', '수식 표시'],
  ['Ctrl+Shift+U', '수식 입력줄 펼치기/접기'],
  ['Ctrl+Shift+= / Ctrl+-', '행·열 삽입 / 삭제'],
  ['Ctrl+PageUp / PageDown', '이전 / 다음 시트'],
  ['Ctrl+마우스 휠', '확대/축소'],
  ['Shift+F11', '새 시트'],
  ['Shift+F2', '메모 편집'],
  ['Shift+F10', '바로 가기 메뉴'],
  ['Ctrl+N / Ctrl+O / Ctrl+S / F12', '새 통합 문서 / 열기 / 저장 / 다른 이름으로 저장'],
  ['Ctrl+P / Ctrl+F2', '인쇄 / 인쇄 미리 보기'],
  ['Alt+F8 / Alt+F11', '매크로 (VBA 코드 보기)'],
  ['F9 / Shift+F9 / Ctrl+Alt+F9', '다시 계산'],
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

let lastPatternColor = '#000000';
function colorMenu(anchorEl, kind) {
  // 채우기: 엑셀처럼 무늬(패턴) 채우기 · 무늬 색도 바로 고를 수 있게
  const extra = kind !== 'fill' ? [] : (() => {
    const cur = wb.styleAt(si, active.r, active.c);
    const pc = cur.patternColor ?? lastPatternColor;
    const grid = el('div', { class: 'pattern-grid' }, PATTERNS.map(([v, l]) => el('button', {
      class: `pattern-swatch${cur.pattern === v ? ' on' : ''}`, title: l, onmousedown: (e) => e.preventDefault(),
      style: { background: patternCss(v, pc, cur.fill ?? '#ffffff') },
      onclick: () => { closeMenus(); applyStyle({ pattern: v, patternColor: pc }); focusGrid(); },
    })));
    return [
      { sep: true },
      {
        label: '무늬 스타일', icon: 'fill', submenu: [
          { label: '무늬 없음', action: () => applyStyle({ pattern: undefined, patternColor: undefined }) },
          { node: grid },
        ],
      },
      {
        label: '무늬 색', submenu: [
          ...['#000000', '#7f7f7f', '#c00000', '#ff0000', '#ffc000', '#ffff00', '#92d050', '#00b050', '#00b0f0', '#0070c0', '#002060', '#7030a0'].map((c) => ({
            label: c, swatch: c, action: () => { lastPatternColor = c; applyStyle({ patternColor: c, ...(cur.pattern ? {} : { pattern: 'lightGray' }) }); },
          })),
        ],
      },
      { label: '셀 서식 채우기...', action: () => run('formatCells') },
    ];
  })();
  paletteMenu(anchorEl, kind === 'fill' ? '채우기 없음' : '자동', (color) => {
    if (kind === 'fill') { if (color) lastFill = color; applyStyle(color ? { fill: color } : { fill: undefined, pattern: undefined, patternColor: undefined }); }
    else { if (color) lastFont = color; applyStyle({ color: color || undefined }); }
  }, extra);
}

/** 테마 색 · 표준 색 · 다른 색 팔레트 (onPick(색 | null)) */
function paletteMenu(anchorEl, noneLabel, onPick, extra = []) {
  const pick = (color) => {
    closeMenus();
    onPick(color);
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
    { label: noneLabel, action: () => pick(null) },
    { node: palette },
    { sep: true },
    { node: el('div', {}, custom) },
    { label: '다른 색...', action: () => custom.click() },
    ...extra,
  ]);
}

/** 조건부 서식 견본 모음 (엑셀처럼 작은 견본 격자, 마우스를 올리면 이름) */
function cfGallery(items, add) {
  return el('div', { class: 'cf-gallery' }, items.map((it) => el('button', {
    class: 'cf-swatch', title: it.n, style: { background: it.bg }, onmousedown: (e) => e.preventDefault(),
    onclick: () => { closeMenus(); add(it.rule); focusGrid(); },
  })));
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
const RESET_STYLE = { fill: undefined, color: undefined, bold: undefined, italic: undefined, underline: undefined, strike: undefined, size: undefined, font: undefined, bt: undefined, bb: undefined, bl: undefined, br: undefined,
  btc: undefined, bbc: undefined, blc: undefined, brc: undefined, bts: undefined, bbs: undefined, bls: undefined, brs: undefined };

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
  calcOptions: () => [{ label: '자동', checked: opts.calcMode === 'auto', action: () => run('calcAuto') }, { label: '수동', checked: opts.calcMode === 'manual', action: () => run('calcManual') }],
  shapeChange: () => [{ node: shapeGallery((k) => patchObjects({ kind: k }, ['shapes']), true) }],
  shapeFill: (a) => { shapeFillMenu(a); },
  shapeOutline: (a) => { shapeOutlineMenu(a); },
  shapeEffects: (a) => { shapeEffectsMenu(a); },
  shapeStyles: (a) => { shapeStylesMenu(a); },
  wordArt: (a) => { wordArtMenu(a); },
  textFill: (a) => { textFillMenu(a); },
  textOutline: (a) => { textOutlineMenu(a); },
  textEffects: (a) => { textEffectsMenu(a); },
  objForward: () => [{ label: '앞으로 가져오기', icon: 'bringForward', action: () => chartSel && arrangeObject(chartSel, 'forward') }, { label: '맨 앞으로 가져오기', icon: 'bringForward', action: () => chartSel && arrangeObject(chartSel, 'front') }],
  objBackward: () => [{ label: '뒤로 보내기', icon: 'sendBackward', action: () => chartSel && arrangeObject(chartSel, 'backward') }, { label: '맨 뒤로 보내기', icon: 'sendBackward', action: () => chartSel && arrangeObject(chartSel, 'back') }],
  objAlign: () => [
    { label: '왼쪽 맞춤', action: () => alignObjects('left') }, { label: '가운데 맞춤', action: () => alignObjects('center') }, { label: '오른쪽 맞춤', action: () => alignObjects('right') },
    { sep: true }, { label: '위쪽 맞춤', action: () => alignObjects('top') }, { label: '중간 맞춤', action: () => alignObjects('middle') }, { label: '아래쪽 맞춤', action: () => alignObjects('bottom') },
    { sep: true }, { label: '가로 간격을 동일하게', action: () => alignObjects('distH') }, { label: '세로 간격을 동일하게', action: () => alignObjects('distV') },
    { sep: true }, { label: '눈금(셀)에 맞춤', action: () => alignObjects('grid') },
  ],
  objRotate: () => [
    { label: '오른쪽으로 90도 회전', icon: 'rotate', action: () => rotateObjects('r90') }, { label: '왼쪽으로 90도 회전', action: () => rotateObjects('l90') },
    { label: '상하 대칭', action: () => rotateObjects('flipV') }, { label: '좌우 대칭', action: () => rotateObjects('flipH') },
    { sep: true }, { label: '기타 회전 옵션...', action: () => formDialog('회전', [{ name: 'd', label: '회전 각도(°)', type: 'number', value: selectedObjects()[0]?.obj.rot ?? 0 }], ({ d }) => setObjSize('rot', d)) },
  ],
  objPlacement: (a) => { objPlacementMenu(a); },
  equations: () => equationMenu(),
  scatterCharts: () => CHART_GALLERY.find(([g]) => g === '분산형')[1].map(([n, p]) => ({ label: n, icon: 'chartScatter', action: () => insertChart(p.type, p) })),
  chartElements: () => chartElementsMenu(),
  chartLayouts: () => chartLayoutsMenu(),
  chartColors: (a) => { chartColorsMenu(a); },
  chartStyles: (a) => { chartStylesMenu(a); },
  calcFields: () => {
    const e = pivotHere();
    const calcs = e ? pivotDefV2(e.def).calcFields ?? [] : [];
    return [
      { label: '계산 필드...', icon: 'fx', action: () => calcFieldDialog(e, '\u0000new') },
      ...(calcs.length ? [{ title: '계산 필드 수정' }, ...calcs.slice(0, 20).map((c) => ({ label: `ƒx ${c.name}  =${c.formula}`, action: () => calcFieldDialog(e, c.name) }))] : []),
      { sep: true },
      { label: '수식 나열', action: () => listCalcFormulas(e) },
    ];
  },
  picture: () => [
    { title: '이 디바이스' },
    { label: '셀에 배치...', icon: 'picture', action: () => insertPictureInCell() },
    { label: '셀 위에 배치...', icon: 'picture', action: () => insertPicture() },
    { title: '온라인' },
    { label: '온라인 그림...', icon: 'search', action: () => onlinePictureDialog(false) },
    { label: '온라인 그림을 셀에 배치...', icon: 'search', action: () => onlinePictureDialog(true) },
  ],
  pivotLayout: () => [
    { label: '압축 형식으로 표시', action: () => run('pivotCompact') },
    { label: '개요 형식으로 표시', action: () => run('pivotOutline') },
    { label: '테이블 형식으로 표시', action: () => run('pivotTabular') },
    { sep: true },
    { label: '모든 항목 레이블 반복', action: () => pivotLayoutCmd({ repeatLabels: true }) },
    { label: '항목 레이블을 반복하지 않음', action: () => pivotLayoutCmd({ repeatLabels: false }) },
  ],
  pivotSubtotals: () => [
    { label: '부분합 표시 안 함', action: () => run('pivotSubOff') },
    { label: '그룹 하단에 모든 부분합 표시', action: () => pivotLayoutCmd({ subtotals: true, subtotalTop: false }) },
    { label: '그룹 상단에 모든 부분합 표시', action: () => pivotLayoutCmd({ subtotals: true, subtotalTop: true }) },
  ],
  outlineGroupMenu: () => [
    { label: '그룹...', action: () => run('outlineGroup') },
    { label: '자동 개요', action: () => toast('수식이 있는 요약 행을 기준으로 한 자동 개요는 [부분합]을 사용하세요.') },
  ],
  outlineUngroupMenu: () => [
    { label: '그룹 해제...', action: () => run('outlineUngroup') },
    { label: '개요 지우기', action: () => run('outlineClear') },
  ],
  pivotBlank: () => [
    { label: '각 항목 다음에 빈 줄 삽입', action: () => pivotLayoutCmd({ blankRows: true }) },
    { label: '각 항목 다음에 빈 줄 제거', action: () => pivotLayoutCmd({ blankRows: false }) },
  ],
  pivotGrand: () => [
    { label: '행 및 열의 총합계 해제', action: () => run('pivotGrandOff') },
    { label: '행 및 열의 총합계 설정', action: () => run('pivotGrandOn') },
    { label: '행의 총합계만 설정', action: () => run('pivotGrandRows') },
    { label: '열의 총합계만 설정', action: () => run('pivotGrandCols') },
  ],
  fontList: (a) => { fontMenu(a); },
  pivotStylesDesign: (a) => { pivotStyleGallery(a); },
  slicerStyles: (a) => { slicerStyleGallery(a); },
  marginsMenu: () => [
    ...MARGINS.map((m) => ({ label: `${m.label} (위 ${m.m.top}" 아래 ${m.m.bottom}" 왼쪽 ${m.m.left}" 오른쪽 ${m.m.right}")`, action: () => patchPage({ margins: { ...m.m } }) })),
    { sep: true },
    { label: '사용자 지정 여백...', action: () => pageSetupDialog() },
  ],
  orientMenu: () => [
    { label: '세로', action: () => run('orientPortrait') },
    { label: '가로', action: () => run('orientLandscape') },
  ],
  paperMenu: () => PAPERS.map((p) => ({ label: p.label, action: () => patchPage({ paper: p.id }) })),
  printAreaMenu: () => [
    { label: '인쇄 영역 설정', action: () => run('printArea') },
    { label: '인쇄 영역 해제', action: () => run('clearPrintArea') },
  ],
  fitMenu: () => [
    { label: '현재 크기 (100%)', action: () => patchPage({ scale: 100, fitW: 0, fitH: 0 }) },
    { label: '한 페이지에 시트 맞추기', action: () => patchPage({ fitW: 1, fitH: 1 }) },
    { label: '한 페이지에 모든 열 맞추기', action: () => patchPage({ fitW: 1, fitH: 0 }) },
    { label: '한 페이지에 모든 행 맞추기', action: () => patchPage({ fitW: 0, fitH: 1 }) },
  ],
  whatIf: () => [
    { label: '시나리오 관리자...', action: () => run('scenarioManager') },
    { label: '목표값 찾기...', action: () => run('goalSeek') },
    { label: '데이터 표...', action: () => run('dataTable') },
  ],
  useInFormula: () => {
    const names = wb.names.filter((n) => !n.hidden && (!n.sheet || n.sheet.toLowerCase() === sheet().name.toLowerCase()));
    return [
      ...names.map((n) => ({ label: n.name, action: () => { if (editing) { const inp = edInput(); const pos = inp.selectionStart; setEditText(`${inp.value.slice(0, pos)}${n.name}${inp.value.slice(inp.selectionEnd)}`, pos + n.name.length); } else startEdit('enter', `=${n.name}`); } })),
      ...(names.length ? [{ sep: true }] : []),
      { label: '이름 붙여넣기...', action: pasteNameDialog },
    ];
  },
  tableStylesDesign: (a) => { tableStyleGallery(a, false); },
  shapes: () => [{ node: shapeGallery((k) => startDraw(k)) }],
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
    { sep: true },
    { label: '선택하여 붙여넣기...', key: 'Ctrl+Alt+V', action: pasteSpecialDialog, disabled: !clip },
  ],
  borders: (a) => [
    { title: '테두리' },
    { label: '아래쪽 테두리', icon: 'borderBottom', action: () => applyBorder('bottom') },
    { label: '위쪽 테두리', icon: 'borderTop', action: () => applyBorder('top') },
    { label: '왼쪽 테두리', icon: 'borderLeft', action: () => applyBorder('left') },
    { label: '오른쪽 테두리', icon: 'borderRight', action: () => applyBorder('right') },
    { sep: true },
    { label: '테두리 없음', icon: 'borderNone', action: () => applyBorder('none') },
    { label: '모든 테두리', icon: 'borderAll', action: () => applyBorder('all') },
    { label: '바깥쪽 테두리', icon: 'borderOutside', action: () => applyBorder('outside') },
    { label: '굵은 바깥쪽 테두리', icon: 'borderOutside', action: () => applyBorder('thickOutside') },
    { label: '안쪽 테두리', icon: 'borderAll', action: () => applyBorder('inside') },
    { label: '안쪽 가로 테두리', icon: 'borderBottom', action: () => applyBorder('insideH') },
    { label: '안쪽 세로 테두리', icon: 'borderLeft', action: () => applyBorder('insideV') },
    { sep: true },
    { label: '아래쪽 이중 테두리', icon: 'borderBottom', action: () => applyBorder('doubleBottom') },
    { label: '굵은 아래쪽 테두리', icon: 'borderBottom', action: () => applyBorder('thickBottom') },
    { label: '위쪽/아래쪽 테두리', icon: 'borderBottom', action: () => applyBorder('topBottom') },
    { label: '위쪽/굵은 아래쪽 테두리', icon: 'borderBottom', action: () => applyBorder('topThickBottom') },
    { label: '위쪽/아래쪽 이중 테두리', icon: 'borderBottom', action: () => applyBorder('topDoubleBottom') },
    { label: '대각선 테두리 (↘)', action: () => applyBorder('diagDown') },
    { label: '대각선 테두리 (↗)', action: () => applyBorder('diagUp') },
    { title: '테두리 그리기' },
    { label: '테두리 그리기', icon: 'border', checked: borderDraw === 'outline', action: () => setBorderDraw('outline') },
    { label: '테두리 눈금 그리기', icon: 'borderAll', checked: borderDraw === 'grid', action: () => setBorderDraw('grid') },
    { label: '테두리 지우기', icon: 'clear', checked: borderDraw === 'erase', action: () => setBorderDraw('erase') },
    { node: el('div', { class: 'pen-row' }, el('span', {}, '선 색'), el('span', { class: 'pen-swatch', style: { background: borderPen.color ?? '#000000' } })), },
    { label: '선 색...', icon: 'fontColor', action: () => setTimeout(() => paletteMenu(a ?? { x: 200, y: 160 }, '자동', (c) => { borderPen.color = c || null; }), 0) },
    { label: '선 스타일...', action: () => setTimeout(() => lineStyleMenu(a ?? { x: 200, y: 160 }), 0) },
    { sep: true },
    { label: '다른 테두리...', action: () => formatCellsDialog(3) },
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
    const add = (rule) => addCondRuleHere(rule);
    return [
      { title: '셀 강조 규칙' },
      { label: '보다 큼...', icon: '<span class="cfi">&gt;</span>', action: () => condRuleDialog('gt') },
      { label: '보다 작음...', icon: '<span class="cfi">&lt;</span>', action: () => condRuleDialog('lt') },
      { label: '다음 값의 사이에 있음...', icon: '<span class="cfi">↔</span>', action: () => condRuleDialog('between') },
      { label: '같음...', icon: '<span class="cfi">=</span>', action: () => condRuleDialog('eq') },
      { label: '텍스트 포함...', icon: '<span class="cfi">ab</span>', action: () => condRuleDialog('text') },
      { label: '발생 날짜...', icon: '<span class="cfi">31</span>', action: () => cfRuleEditor({ type: 'date', period: 'today', style: { fill: '#ffc7ce', color: '#9c0006' } }, (nr) => addCondRuleHere(nr), { title: '발생 날짜' }) },
      { label: '중복 값...', icon: '<span class="cfi">≡</span>', action: () => condRuleDialog('dup') },
      { title: '상위/하위 규칙' },
      { label: '상위 10개 항목...', icon: '<span class="cfi g">10↑</span>', action: () => condRuleDialog('top') },
      { label: '상위 10%...', icon: '<span class="cfi g">%↑</span>', action: () => cfRuleEditor({ type: 'top', v1: '10', percent: true, style: { fill: '#ffc7ce', color: '#9c0006' } }, (nr) => addCondRuleHere(nr), { title: '상위 10%' }) },
      { label: '하위 10개 항목...', icon: '<span class="cfi">10↓</span>', action: () => cfRuleEditor({ type: 'bottom', v1: '10', style: { fill: '#ffc7ce', color: '#9c0006' } }, (nr) => addCondRuleHere(nr), { title: '하위 10개 항목' }) },
      { label: '평균 초과', icon: '<span class="cfi b">x̄↑</span>', action: () => add({ type: 'aboveAvg', style: { fill: '#ffc7ce', color: '#9c0006' } }) },
      { label: '평균 미만', icon: '<span class="cfi b">x̄↓</span>', action: () => add({ type: 'belowAvg', style: { fill: '#ffc7ce', color: '#9c0006' } }) },
      { title: '데이터 막대' },
      { node: cfGallery([...BAR_PRESETS.modern.map(([n, color]) => ({ n: `${n} (모던)`, bg: `linear-gradient(90deg, ${color} 0 62%, transparent 62%)`, rule: { type: 'bar', color } })),
        ...BAR_PRESETS.excel.map(([n, color]) => ({ n: `${n} 그라데이션 채우기`, bg: `linear-gradient(90deg, ${color}, #fff 62%, transparent 62%)`, rule: { type: 'bar', color } }))], add) },
      { title: '색조 — WIXEL 모던' },
      { node: cfGallery(SCALE_PRESETS.modern.map(([n, colors]) => ({ n, bg: `linear-gradient(180deg,${colors.join(',')})`, rule: { type: 'scale', colors } })), add) },
      { title: '색조 — 엑셀 기본' },
      { node: cfGallery(SCALE_PRESETS.excel.map(([n, colors]) => ({ n, bg: `linear-gradient(180deg,${colors.join(',')})`, rule: { type: 'scale', colors } })), add) },
      { title: '아이콘 집합' },
      ...ICON_SETS.slice(0, 5).map((set) => ({
        label: set.label, icon: `<span style="display:inline-flex">${set.icons.map((i) => ICON_SVG[i]).join('')}</span>`,
        action: () => add({ type: 'icons', icons: set.id }),
      })),
      { sep: true },
      { label: '새 규칙...', icon: 'condFormat', action: () => cfRuleEditor(null, (nr) => addCondRuleHere(nr)) },
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
        wb.transact(() => { for (let r = u.r1; r <= r2; r++) wb.setRowHeight(si, r, defRowH(), false); autoFitRows(u.r1, r2); }, meta());
      },
    },
    { label: '열 너비...', action: () => sizeDialog('col') },
    { label: '열 너비 자동 맞춤', action: () => autofitCols(range(sel.c1, Math.min(sel.c2, sel.c1 + 200))) },
    ...MENUS_HIDE(),
    { title: '시트 구성' },
    { label: '시트 이름 바꾸기', action: () => renameSheetInline(si) },
    { label: '시트 복사본 만들기', icon: 'copy', action: () => run('duplicateSheet') },
    { label: '탭 색', icon: 'fill', action: () => { const b = document.querySelector('.sheet-tab.active')?.getBoundingClientRect(); setTimeout(() => tabColorMenu({ x: b?.left ?? 200, y: (b?.top ?? 600) - 330 }, si), 0); } },
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
    { label: '시트 숨기기', action: () => hideSheet() },
    { label: '시트 숨기기 취소...', disabled: !wb.sheets.some((_, j) => isHiddenSheet(j)), action: () => unhideSheetDialog() },
  ];
}

function openNamedMenu(name, anchorEl) {
  if (editing && !commitEdit()) return;
  if (name.startsWith('fn:')) {
    const cat = name.slice(3);
    if (cat === 'more') {
      openMenu(anchorEl, ['통계', '공학', '정보', '데이터베이스', '호환성'].map((c) => ({ label: c, action: () => openNamedMenu(`fn:${c}`, anchorEl) })));
      return;
    }
    const names = FUNCTION_NAMES.filter((n) => FUNC_INFO[n]?.cat === cat);
    openMenu(anchorEl, names.map((n) => ({ label: n, action: () => insertFunctionText(n) })), { scroll: true });
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
    const lk = wb.getCell(si, active.r, active.c)?.link;
    const ci = wb.getCell(si, active.r, active.c)?.image;
    if (ci) {
      const { r, c } = active;
      items.push(
        { label: '셀 위에 그림 배치', icon: 'picture', action: () => cellImageToFloating(r, c) },
        { label: '대체 텍스트...', action: () => cellImageAltDialog(r, c) },
        { sep: true },
      );
    }
    items.push(
      { label: lk ? '하이퍼링크 편집...' : '링크', key: 'Ctrl+K', action: hyperlinkDialog },
      ...(lk ? [{ label: '하이퍼링크 열기', action: () => openLink(lk) }, { label: '하이퍼링크 제거', action: removeHyperlink }] : []),
      { label: '선택하여 붙여넣기...', key: 'Ctrl+Alt+V', action: pasteSpecialDialog, disabled: !clip },
      { sep: true },
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
  selectionPane: () => selectionPaneDialog(),
  objH: (v) => setObjSize('h', v),
  objW: (v) => setObjSize('w', v),
  objRot: (v) => setObjSize('rot', v),
  zoomSel: () => zoomToSelection(),
  insertSymbol: () => insertSymbolDialog(),
  insertChartAll: () => insertChartAllDialog(),
  insertPivotChart: () => insertPivotChart(),
  chartSwitch: () => chartSwitchRowCol(),
  chartSelectData: () => { if (!chartSel) return; if (chartHere()?.pivot) pivotChartMetrics(chartSel); else chartDialog(chartSel); },
  pivotChartMetrics: () => pivotChartMetrics(),
  chartChangeType: () => { if (chartSel) insertChartAllDialog(chartSel); },
  chartFormat: () => chartFormatPane(),
  chartPivotFields: () => { const ch = chartHere(); if (ch?.pivot) { updateChart(ch.id, { fieldButtons: ch.fieldButtons === false ? undefined : false }); gv.renderObjectsAll(); } else toast('피벗 차트에서 쓸 수 있습니다.'); },
  slicerFontSize: (v) => { if (chartSel) { updateObject(chartSel, { fontSize: Number(v) > 0 ? clamp(Number(v), 5, 72) : undefined }); gv.renderObjectsAll(); } },
  slicerHeadSize: (v) => { if (chartSel) { updateObject(chartSel, { headSize: Number(v) > 0 ? clamp(Number(v), 5, 72) : undefined }); gv.renderObjectsAll(); } },
  slicerBold: () => { const sl = (sheet().slicers ?? []).find((x) => x.id === chartSel); if (sl) { updateObject(sl.id, { bold: !sl.bold || undefined }); gv.renderObjectsAll(); } },
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
  fontFamily: (f) => applyStyle({ font: f === BASE_FONT.name ? undefined : f }),
  fontSize: (s) => { const n = Number(s); if (n > 0 && n <= 409) applyStyle({ size: n === BASE_FONT.size ? undefined : n }); },
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
  fmtNumber: () => applyStyle({ numFmt: 'comma', decimals: 0 }, { widen: true }), // 한글 엑셀처럼 #,##0
  fmtDate: () => applyStyle({ numFmt: 'date', decimals: undefined }, { widen: true }),
  fmtCurrency: () => applyStyle({ numFmt: 'accounting', decimals: undefined }, { widen: true }),
  fmtPercent: () => applyStyle({ numFmt: 'percent', decimals: undefined }, { widen: true }),
  fmtComma: () => applyStyle({ numFmt: 'comma', decimals: undefined }, { widen: true }),
  incDecimal: () => changeDecimals(1),
  decDecimal: () => changeDecimals(-1),

  insertRows: structural(() => {
    const n = selKind === 'cols' || selKind === 'all' ? 1 : Math.min(sel.r2 - sel.r1 + 1, 100000);
    wb.transact(() => anchorObjects(() => wb.insertRows(si, sel.r1, n), { axis: 'row', index: sel.r1, count: n }), meta());
  }),
  insertCols: structural(() => {
    const n = selKind === 'rows' || selKind === 'all' ? 1 : Math.min(sel.c2 - sel.c1 + 1, 5000);
    wb.transact(() => anchorObjects(() => wb.insertCols(si, sel.c1, n), { axis: 'col', index: sel.c1, count: n }), meta());
  }),
  deleteRows: structural(() => {
    const r1 = sel.r1;
    const n = selKind === 'cols' ? 1 : sel.r2 - sel.r1 + 1;
    wb.transact(() => anchorObjects(() => wb.deleteRows(si, r1, n), { axis: 'row', index: r1, count: -n }), meta());
    selectCell(r1, active.c);
  }),
  deleteCols: structural(() => {
    const c1 = sel.c1;
    const n = selKind === 'rows' ? 1 : sel.c2 - sel.c1 + 1;
    wb.transact(() => anchorObjects(() => wb.deleteCols(si, c1, n), { axis: 'col', index: c1, count: -n }), meta());
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
        title: 'WIXEL', body: `'${sheet().name}' 시트를 영구적으로 삭제합니다. 계속하시겠습니까?`,
        buttons: [{ label: '삭제', primary: true, action: del }, { label: '취소' }],
      });
    } else del();
  },
  duplicateSheet: () => {
    const { _sid, ...src } = wb.serializeSheet(si);
    const i = wb.transact(() => {
      const at = wb.addSheet(`${src.name} (2)`.slice(0, 31), si + 1);
      const dup = (list) => (list ?? []).map((c) => ({ ...c, id: `${c.id}d${at}` }));
      wb.replaceSheet(at, { ...src, name: wb.sheets[at].name, charts: dup(src.charts), images: dup(src.images), shapes: dup(src.shapes) });
      return at;
    }, meta());
    switchSheet(i, false);
  },
  prevSheet: () => { let j = si - 1; while (j >= 0 && isHiddenSheet(j)) j--; switchSheet(j); },
  nextSheet: () => { let j = si + 1; while (j < wb.sheets.length && isHiddenSheet(j)) j++; switchSheet(j); },
  hideSheet: () => hideSheet(),
  unhideSheet: () => unhideSheetDialog(),

  autosum: () => autoSum('SUM'),
  fillDown: () => fillCopy('down'),
  fillRight: () => fillCopy('right'),
  clearContents: () => wb.transact(() => {
    if (allFilters().some(([, f]) => hidCount(f.hidden))) {
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
  pasteValuesKey: () => pasteFromButton('values'),
  repeatLast, flashFill: flashFillCmd, pasteSpecial: pasteSpecialDialog, hyperlink: hyperlinkDialog, removeHyperlink,
  selectPrecedents, selectDependents, selectComments,
  unhideRows: () => hideSel('row', false), unhideCols: () => hideSel('col', false),
  fmtTime: () => applyStyle({ numFmt: 'time', decimals: undefined }, { widen: true }),
  fmtScientific: () => applyStyle({ numFmt: 'scientific', decimals: undefined }, { widen: true }),
  borderNone: () => applyBorder('none'),
  selectRegion: () => { const rg = currentRegion(active.r, active.c); selectRange(rg, 'cells', { r: active.r, c: active.c }); },
  newWorkbook: () => newWorkbook(),
  nameManager, defineName, useInFormula: pasteNameDialog, pasteName: pasteNameDialog, createNamesFromSel,

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
  pivotFieldList: () => { pivotPaneOpen = !pivotPaneOpen; refreshPivotPane(true); },
  pivotName: (v) => renamePivot(v),
  pivotOptions: () => pivotOptionsDialog(),
  pageSetup: () => pageSetupDialog(),
  orientPortrait: () => patchPage({ orientation: 'portrait' }),
  orientLandscape: () => patchPage({ orientation: 'landscape' }),
  printArea: () => { patchPage({ area: usedClip(sel) }); toast(`인쇄 영역: ${cellName(sel.r1, sel.c1)}:${cellName(usedClip(sel).r2, usedClip(sel).c2)}`); },
  clearPrintArea: () => patchPage({ area: null }),
  printTitles: () => pageSetupDialog(),
  sparkLine: () => insertSparkline('line'),
  sparkColumn: () => insertSparkline('column'),
  sparkWinLoss: () => insertSparkline('winloss'),
  sparkTypeLine: () => updateSparkGroup({ type: 'line' }),
  sparkTypeColumn: () => updateSparkGroup({ type: 'column' }),
  sparkTypeWinLoss: () => updateSparkGroup({ type: 'winloss' }),
  sparkHigh: () => updateSparkGroup((g) => ({ high: !g.high })),
  sparkLow: () => updateSparkGroup((g) => ({ low: !g.low })),
  sparkNegative: () => updateSparkGroup((g) => ({ negative: !g.negative })),
  sparkFirst: () => updateSparkGroup((g) => ({ first: !g.first })),
  sparkLast: () => updateSparkGroup((g) => ({ last: !g.last })),
  sparkMarkers: () => updateSparkGroup((g) => ({ markers: !g.markers })),
  sparkClear: () => updateSparkGroup(() => ({ items: [] })),
  sparkEdit: () => sparkEditDialog(),
  protectSheet: () => protectSheetDialog(),
  unprotectSheet: () => unprotectSheet(),
  cellProtection: () => cellProtectionDialog(),
  tracePrecedents: () => tracePrecedents(),
  traceDependents: () => traceDependents(),
  removeArrows: () => removeArrows(),
  evaluateFormula: () => evaluateFormulaDialog(),
  errorCheck: () => errorCheck(),
  watchWindow: () => watchWindow(!watchPane),
  goalSeek: () => goalSeekDialog(),
  scenarioManager: () => scenarioManager(),
  publish: () => publishDialog(),
  versionHistory: () => versionHistory(),
  recentFiles: () => openBackstage('open'),
  forecastSheet: () => forecastSheetDialog(),
  dataAnalysis: () => dataAnalysisDialog(),
  solver: () => solverDialog(),
  dataTable: () => dataTableDialog(),
  gotoSpecial: () => gotoSpecialDialog(),
  outlineGroup: () => outlineGroup(1),
  outlineUngroup: () => outlineGroup(-1),
  outlineClear: () => outlineClear(),
  outlineShow: () => outlineDetail(true),
  outlineHide: () => outlineDetail(false),
  subtotal: () => subtotalDialog(),
  pivotExpandField: () => pivotExpandField(true),
  pivotCollapseField: () => pivotExpandField(false),
  pivotGroupField: () => pivotGroupDialog(),
  pivotUngroup: () => pivotUngroup(),
  pivotDetail: () => { const pv = pivotHere(); if (!pv || !showPivotDetail(pv, active.r, active.c)) toast('피벗 테이블의 값 셀을 선택하세요.'); },
  pivotShowExpand: () => { const pv = pivotHere(); if (pv) pivotLayoutCmd({ showExpand: pv.def.showExpand === false }); },
  pivotChangeSource: () => pivotChangeSourceDialog(),
  pivotClear: () => { const e = pivotHere(); if (e) { setPivotDef(e, { ...pivotDefV2(e.def), rows: [], cols: [], values: [], pages: [], filters: {}, fieldFilters: {}, sort: {} }); refreshPivotPane(true); } },
  calcField: () => calcFieldDialog(),
  slicerConnections: () => slicerConnectionsDialog(),
  pvRowHeaders: () => pivotStyleOpt('rowHeaders'),
  pvColHeaders: () => pivotStyleOpt('colHeaders'),
  pvBandRows: () => pivotStyleOpt('bandRows'),
  pvBandCols: () => pivotStyleOpt('bandCols'),
  slicerBtnH: (v) => { if (chartSel) updateObject(chartSel, { buttonHeight: clamp(Number(v) || 24, 10, 120) }); },
  slicerBtnW: (v) => { if (chartSel) { updateObject(chartSel, { buttonWidth: Number(v) > 0 ? clamp(Number(v), 10, 600) : undefined }); gv.renderObjectsAll(); } },
  slicerGap: (v) => { if (chartSel && v !== '') { updateObject(chartSel, { gap: clamp(Number(v) || 0, 0, 30) }); gv.renderObjectsAll(); } },
  slicerHeader: () => { const sl = (sheet().slicers ?? []).find((x) => x.id === chartSel); if (sl) updateObject(sl.id, { showHeader: sl.showHeader === false ? undefined : false }); },
  pivotRefresh: () => refreshPivots(),
  pivotCompact: () => pivotLayoutCmd({ layout: 'compact' }),
  pivotOutline: () => pivotLayoutCmd({ layout: 'outline' }),
  pivotTabular: () => pivotLayoutCmd({ layout: 'tabular' }),
  pivotSubOff: () => pivotLayoutCmd({ subtotals: false }),
  pivotSubOn: () => pivotLayoutCmd({ subtotals: true }),
  pivotGrandOff: () => pivotLayoutCmd({ grandRows: false, grandCols: false }),
  pivotGrandOn: () => pivotLayoutCmd({ grandRows: true, grandCols: true }),
  pivotGrandRows: () => pivotLayoutCmd({ grandRows: true, grandCols: false }),
  pivotGrandCols: () => pivotLayoutCmd({ grandRows: false, grandCols: true }),
  refreshAll: refreshPivots,
  chartColumn: () => insertChart('column'),
  chartBar: () => insertChart('bar'),
  chartLine: () => insertChart('line'),
  chartPie: () => insertChart('pie'),
  chartArea: () => insertChart('area'),
  chartScatter: () => insertChart('scatter'),
  insertPicture,
  insertPictureInCell,
  textToColumns,
  condManager: cfManager,
  condNewRule: () => cfRuleEditor(null, (nr) => addCondRuleHere(nr)),
  createTable: () => createTableDialog(),
  insertSlicer: insertSlicerDialog,
  insertTimeline: insertTimelineDialog,
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

  toggleGrid: (v) => { const on = v ?? !!sheet().noGrid; wb.transact(() => wb.setSheetProp(si, 'noGrid', on ? undefined : true), meta()); applyView(); },
  togglePrintGrid: (v) => { view.printGrid = v ?? !view.printGrid; updateRibbon(); },
  toggleFormulaBar: (v) => { view.showFormulaBar = v ?? !view.showFormulaBar; applyView(); },
  toggleHeaders: (v) => { view.showHeaders = v ?? !view.showHeaders; applyView(); },
  toggleFormulas: () => { view.showFormulas = !view.showFormulas; gv.renderAll(); updateRibbon(); },
  toggleRibbon: () => ribbon.toggleCollapse(),
  zoomIn: () => setZoom(view.zoom + 10),
  zoomOut: () => setZoom(view.zoom - 10),
  zoom100: () => setZoom(100),
  recalc: () => { wb.calculateNow(); wb.invalidate(); gv.renderAll(); updateStatusCalc(); },
  calcNowSheet: () => { wb.calculateNow(); wb.invalidate(si); gv.renderAll(); updateStatusCalc(); },
  options: () => optionsDialog(),
  calcAuto: () => { opts.calcMode = 'auto'; saveOptions(); applyOptions(); wb.calculateNow(); gv.renderAll(); },
  calcManual: () => { opts.calcMode = 'manual'; saveOptions(); applyOptions(); },

  shortcuts: () => openDialog({
    title: '바로 가기 키', width: 580,
    body: el('table', { class: 'kbd-table' }, SHORTCUTS.map(([k, d]) => el('tr', {}, el('td', {}, k), el('td', {}, d)))),
    buttons: [{ label: '닫기', primary: true }],
  }),
  about: () => openDialog({
    title: 'WIXEL 정보', width: 420,
    body: el('div', { style: { lineHeight: '1.7' } },
      el('b', {}, 'WIXEL'), ' — 브라우저에서 동작하는 엑셀 스타일 스프레드시트', el('br'),
      el('span', { class: 'muted' }, `시트 크기 20,000,000행 × 16,384열 · 함수 ${FUNCTION_NAMES.length}개 · .xlsx 열기/저장`), el('br'),
      el('span', { class: 'muted' }, server.available ? '서버 저장소에 연결됨 — 다른 기기에서도 열 수 있습니다.' : '서버 없이 실행 중 — 이 브라우저에 저장됩니다.')),
    buttons: [{ label: '확인', primary: true }],
  }),
};

const NO_COMMIT = new Set(['toggleRibbon', 'zoomIn', 'zoomOut', 'zoom100', 'shortcuts', 'about']);

function run(cmd, arg) {
  closeMenus();
  if (viewOnly && protectAction(cmd) !== 'free' && !VIEW_CMDS.has(cmd)) { toast('읽기 전용으로 게시된 문서입니다. [편집용 사본 만들기]를 누르면 고칠 수 있습니다.'); return; }
  if (editing && !NO_COMMIT.has(cmd)) {
    if (cmd === 'undo') { cancelEdit(); return; }
    if (!commitEdit()) return;
  }
  const fn = COMMANDS[cmd];
  if (!fn) { toast('지원하지 않는 기능입니다.'); return; }
  if (protectBlocked(protectAction(cmd), sel, cmd)) return;
  fn(arg);
  if (REPEATABLE.has(cmd)) lastRepeat = () => COMMANDS[cmd](arg);
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

/** 시트를 처음 볼 때: 엑셀에서 저장한 첫 화면(처음 보이는 칸 · 활성 셀), 없으면 맨 위 */
function showSheetStart() {
  const v = sheet().view;
  const f = sheet().freeze;
  if (v) {
    gv.setScroll(Math.max(0, gv.cols.pos(v.left ?? 0) - gv.cols.pos(f?.cols || 0)), Math.max(0, gv.rows.pos(v.top ?? 0) - gv.rows.pos(f?.rows || 0)));
    selectCell(v.r ?? v.top ?? 0, v.c ?? v.left ?? 0, { scroll: false });
  } else {
    gv.setScroll(0, 0);
    selectCell(f?.rows || 0, f?.cols || 0);
  }
}

/** 시트마다 저장된 확대/축소 (엑셀과 같이 시트별) */
function applySheetZoom() {
  const z = clamp(Math.round(sheet().zoom ?? 100), 25, 400);
  if (z === view.zoom) return;
  view.zoom = z;
  dom.zoomSlider.value = z;
  dom.zoomLabel.textContent = `${z}%`;
  gv.setZoom(z);
}

function setZoom(z) {
  view.zoom = clamp(Math.round(z), 25, 400);
  // 확대/축소는 시트 속성 (파일에 저장, 실행 취소 기록은 남기지 않음)
  if (view.zoom === 100) delete sheet().zoom; else sheet().zoom = view.zoom;
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
    font: st.font || BASE_FONT.name, size: String(st.size || BASE_FONT.size), numFmt: st.numFmt === 'custom' ? 'custom' : fmt,
    alignLeft: st.align === 'left', alignCenter: st.align === 'center', alignRight: st.align === 'right',
    valignTop: st.valign === 'top', valignMiddle: st.valign === 'middle', valignBottom: !st.valign,
    merged: !!wb.mergeAt(si, active.r, active.c), painter: !!painter, filterOn: (() => { const k = filterKeyHere(); return k !== null && !!getFilter(k); })(),
    frozen: !!(f.rows || f.cols), lastFill, lastFont, ...view, showGrid: !sheet().noGrid,
    ...tableRibbonState(),
  };
}

function tableRibbonState() {
  const t = chartSel ? null : tableHere();
  const sl = chartSel ? (sheet().slicers ?? []).find((x) => x.id === chartSel) : null;
  const pv = chartSel ? null : pivotHere();
  const sg = chartSel ? null : sparkGroupAt(active.r, active.c);
  const obj = chartSel ? findObject(sheet(), chartSel) : null;
  const context = [...(t ? ['table'] : []), ...(sl ? ['slicer'] : []), ...(pv ? ['pivot'] : []), ...(sg ? ['spark'] : []),
    ...(obj && obj.prop === 'charts' ? ['chart'] : []), ...(obj && obj.prop !== 'slicers' ? ['object'] : [])];
  const so = { rowHeaders: true, colHeaders: true, bandRows: false, bandCols: false, ...(pv?.def.styleOpts ?? {}) };
  const base = {
    context, slicerCaption: sl?.caption ?? '', slicerCols: String(sl?.columns ?? 1), slicerMultiOn: !!sl?.multi,
    objH: obj ? String(Math.round(obj.obj.h)) : '', objW: obj ? String(Math.round(obj.obj.w)) : '', objRot: obj ? String(obj.obj.rot ?? 0) : '',
    slicerFontSize: sl?.fontSize ? String(sl.fontSize) : '', slicerHeadSize: sl?.headSize ? String(sl.headSize) : '', slicerBtnW: String(sl?.buttonWidth ?? 0), slicerGap: String(sl?.gap ?? 3), slicerBoldOn: !!sl?.bold,
    chartFieldButtons: obj?.prop === 'charts' && !!obj.obj.pivot && obj.obj.fieldButtons !== false,
    slicerBtnH: String(sl?.buttonHeight ?? 24), slicerHeaderOn: sl ? sl.showHeader !== false : false,
    sparkIsLine: sg?.type === 'line', sparkIsColumn: sg?.type === 'column', sparkIsWinLoss: sg?.type === 'winloss',
    sparkHigh: !!sg?.high, sparkLow: !!sg?.low, sparkNegative: !!sg?.negative, sparkFirst: !!sg?.first, sparkLast: !!sg?.last, sparkMarkers: !!sg?.markers,
    sheetProtected: isProtected(sheet()), pivotName: pv ? pivotNameOf(pv) : '', pvShowExpand: pv ? pv.def.showExpand !== false : false, pvRowHeaders: so.rowHeaders, pvColHeaders: so.colHeaders, pvBandRows: so.bandRows, pvBandCols: so.bandCols,
  };
  if (!t) return base;
  return {
    ...base, tblName: t.name, tblHeader: t.header, tblTotals: t.totals, tblBanded: t.banded !== false, tblBandedCols: !!t.bandedCols,
    tblFirstCol: !!t.firstCol, tblLastCol: !!t.lastCol, tblFilter: !!t.filter,
  };
}

function updateRibbon() {
  ribbon?.update(ribbonState());
  dom.undoBtn.disabled = !wb.canUndo();
  dom.redoBtn.disabled = !wb.canRedo();
}

function updateTitle() {
  const t = `${docName}${!autosave && dirty ? '*' : ''} - WIXEL`;
  if (!dom.title.querySelector('input')) dom.title.textContent = t;
  document.title = t;
  dom.autosave.setAttribute('aria-checked', String(autosave));
  dom.autosaveLabel.textContent = autosave ? '켬' : '끔';
  let state;
  if (!server.available) state = dirty && !autosave ? '저장 안 됨' : '이 브라우저에 저장됨';
  else if (!serverAutosave()) state = dirty ? (autosave ? '브라우저에 저장 중' : '저장 안 됨') : '이 브라우저에 저장됨 (서버는 [저장])';
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
/** 원본이 바뀐 피벗 자동 새로 고침 (WIXEL: def.autoRefresh — 템플릿 · 새 피벗 기본) */
const pivotSrcVer = new Map();
let autoPivotTimer = null;
function scheduleAutoPivots() {
  clearTimeout(autoPivotTimer);
  autoPivotTimer = setTimeout(autoRefreshPivots, 350);
}
const pivotSrcSi = (def) => (def.table ? findTable(wb, def.table)?.si ?? -1 : wb.sheetIndexByName(def.source ?? ''));
function autoRefreshPivots() {
  const due = [];
  for (const e of allPivots()) {
    if (!e.def.autoRefresh) continue;
    const srcSi = pivotSrcSi(e.def);
    if (srcSi < 0) continue;
    const key = `${e.si}:${pivotNameOf(e)}`;
    const ver = `${wb.sheetVersion(srcSi)}`;
    const was = pivotSrcVer.get(key);
    pivotSrcVer.set(key, ver);
    if (was !== undefined && was !== ver) due.push(e);
  }
  if (!due.length) return;
  wb.transact(() => { for (const e of due) writePivot(e.si, e.def, { autofit: false }); }, { ...meta(), joinPrev: true });
  for (const e of due) pivotSrcVer.set(`${e.si}:${pivotNameOf(e)}`, `${wb.sheetVersion(pivotSrcSi(e.def))}`);
  gv.renderObjectsAll();
}

function onBookChange() {
  dirty = true;
  libDirty = true;
  scheduleLibrarySave();
  scheduleAutoPivots();
  if (opts.calcMode === 'manual') updateStatusCalc();
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
  // 구글 스프레드시트처럼 브라우저 기본 메뉴는 띄우지 않음 (글 입력 칸 · 링크 제외).
  // Windows 는 contextmenu 가 버튼을 뗄 때 오므로, 이미 열린 위셀 메뉴 위에서 받는 경우도 막음
  document.addEventListener('contextmenu', (e) => {
    const t = e.target;
    const typing = t instanceof Element && t.matches('input:not([type=checkbox]):not([type=radio]), textarea, [contenteditable=""], [contenteditable=true]') && t !== dom.editor;
    if (!typing && !opts.browserMenu) e.preventDefault();
  });
  document.addEventListener('mousemove', (e) => {
    if (tlDrag) {
      const c = document.elementFromPoint(e.clientX, e.clientY)?.closest?.('.tl-cell');
      if (c && tlDrag.root.contains(c)) { tlDrag.b = Number(c.dataset.i); markTimelineDrag(); }
    }
    lastMouse = { x: e.clientX, y: e.clientY };
    lastShift = e.shiftKey;
    if (drag) onDragMove(e.clientX, e.clientY);
  });
  document.addEventListener('mouseup', (e) => { if (tlDrag) { const d = tlDrag; tlDrag = null; timelineApply(d.id, d.a, d.b); } onDragEnd(e); });

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
    } else if (/\.(xlsx|xlsm|xlsb|xltx|xltm|ods|fods|csv|tsv|tab|prn|txt|tabula|wixel|json)$/i.test(file.name)) {
      openFileObject(file, 'open');
    }
  });

  // 이름 상자
  dom.nameBox.addEventListener('focus', () => { refreshNameList(); dom.nameBox.select(); });
  dom.nameBox.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); if (nameBoxEnter(dom.nameBox.value)) focusGrid(); }
    if (e.key === 'Escape') { dom.nameBox.value = nameBoxLabel(); focusGrid(); }
  });
  dom.nameBox.addEventListener('blur', () => { dom.nameBox.value = nameBoxLabel(); });

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
    if (!bigBook() || dirty) saveToStorage();
    libraryFlush();
    if (!autosave && dirty) { e.preventDefault(); e.returnValue = ''; }
  });
  window.addEventListener('blur', () => { if (drag) onDragEnd(); });
  window.addEventListener('afterprint', () => { dom.printArea.innerHTML = ''; });
}

// ───────────────────────── 시작 ─────────────────────────
async function init() {
  let stored = loadFromStorage();
  if (stored?.idb) {
    // 큰 문서는 IndexedDB 에 시트별로 저장되어 있음 — 나눠서 불러옴
    const prog = progressOverlay('저장된 통합 문서를 여는 중');
    try {
      const metaId = stored.docId;
      stored = await loadBigFromIdb((p) => prog.set(p * 0.6, '불러오는 중'));
      if (stored) stored.docId = metaId;
      wb = new Workbook();
      if (stored?.workbook) await wb.loadAsync(stored.workbook, (p) => prog.set(0.6 + 0.4 * p, '셀 준비 중'));
    } catch {
      stored = null;
    } finally {
      prog.close();
    }
  }
  if (!wb || !stored) wb = new Workbook(stored?.workbook);
  applyBookLook();
  wb.manualCalc = opts.calcMode === 'manual';
  if (stored) {
    docName = stored.docName || docName;
    docId = stored.docId ?? null;
    si = clamp(stored.si || 0, 0, wb.sheets.length - 1);
    autosave = stored.autosave !== false;
  }
  wb.onChange(onBookChange);
  hydrateIcons();
  gv = new GridView({
    state: () => ({
      wb, si, sel, selKind, active, editing: !!editing, clip, fillPreview, refs: editRefs, chartSel, objMulti, circles, focusCell: opts.focusCell,
      special: special?.si === si ? special.cells : null, arrows: trace?.arrows ?? null,
      showGrid: view.showGrid && !sheet().noGrid, showFormulas: view.showFormulas, showHeaders: view.showHeaders,
    }),
    onViewScroll: () => positionEditor(),
    pivotChartFields: (ch) => {
      const e = findPivotEntry(ch.pivot.sheet ?? null, ch.pivot.name ?? null);
      if (!e) return null;
      const d = pivotDefV2(e.def);
      const vals = (d.values ?? []).map((v) => v.name ?? valueName(v));
      return { rows: d.rows ?? [], cols: d.cols ?? [], pages: d.pages ?? [], values: ch.pivot.values?.length ? ch.pivot.values.filter((n) => vals.includes(n)) : vals };
    },
    slicerModel,
    onZoomWheel: (d) => setZoom(view.zoom + d),
    isDragging: () => !!drag,
  });
  ribbon = buildRibbon({ run, openMenu: openNamedMenu, focusGrid, refreshRibbon: updateRibbon });
  bindEvents();
  applyView();
  if (stored && stored.rev !== APP_REV) redrawPivotsQuiet();
  renderAll();
  const f = sheet().freeze;
  selectCell(f?.rows || 0, f?.cols || 0);
  focusGrid();
  window.tabula = {
    wb: () => wb, run, selectCell, selectRange, newWorkbook, templates: TEMPLATES, exportXlsx, gv: () => gv, sample: (i) => newWorkbook(SAMPLES[i]), switchSheet: (i) => { switchSheet(i); },
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
  if (location.hash.startsWith('#view=') || /[?&](view|doc)=/.test(location.search)) await openFromUrl();
  startCollabWatch();
  updateTitle();
}

init();
