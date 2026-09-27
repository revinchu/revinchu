// Tabula 메인: 그리드 렌더링 · 선택 · 편집 · 키보드/마우스 · 명령
import { Workbook, cellData, DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import {
  cellName, colToName, parseRangeName, parse, shiftFormula, listRefs, normalizeFormula,
  FUNCTION_NAMES, isError, quoteSheetName, compareValues, MAX_ROWS, MAX_COLS,
} from './formula.js';
import { formatValue, formatGeneral, displayedDecimals, parseInput } from './format.js';
import { buildRibbon, FONTS, FONT_SIZES } from './ribbon.js';
import {
  el, hydrateIcons, toast, openMenu, closeMenus, isMenuOpen, openDialog, alertDialog,
  formDialog, setMenuCloseHandler, setDialogCloseHandler, isDialogOpen,
} from './ui.js';
import { FUNC_INFO, CATEGORIES } from './funcinfo.js';
import { makeSeries } from './series.js';
import { parseDelimited, toDelimited, guessDelimiter } from './csv.js';
import { SAMPLES } from './samples.js';

const $ = (id) => document.getElementById(id);
const dom = {
  scroll: $('gridScroll'), inner: $('gridInner'), grid: $('grid'),
  colgroup: $('grid').querySelector('colgroup'), thead: $('grid').querySelector('thead'), tbody: $('grid').querySelector('tbody'),
  selBorder: $('selBorder'), fillHandle: $('fillHandle'), fillPreview: $('fillPreview'), marquee: $('copyMarquee'),
  refs: $('refHighlights'), editor: $('cellEditor'), ac: $('autocomplete'),
  nameBox: $('nameBox'), formula: $('formulaInput'), formulaRow: $('formulaRow'),
  fxCancel: $('fxCancel'), fxEnter: $('fxEnter'), fxInsert: $('fxInsert'),
  status: $('statusMode'), stats: $('statusStats'), sheetTabs: $('sheetTabs'),
  zoomSlider: $('zoomSlider'), zoomLabel: $('zoomLabel'), title: $('docTitle'), search: $('searchBox'),
  autosave: $('autosaveToggle'), autosaveLabel: $('autosaveLabel'), fileInput: $('fileInput'), tip: $('commentTip'),
  undoBtn: $('tbUndo'), redoBtn: $('tbRedo'),
};

const STORAGE_KEY = 'tabula.workbook.v1';
const DEFAULT_FONT = '맑은 고딕';
const DEFAULT_SIZE = 11;
const ROW_HEADER_MIN = 34;
// 화면에 만드는 그리드 크기 상한 (DOM 기반 렌더링이므로 엑셀 전체 크기 대신 제한)
const LIMIT_R = Math.min(MAX_ROWS, 10000);
const LIMIT_C = Math.min(MAX_COLS, 702);
const REF_COLORS = ['#2f6fd6', '#d13438', '#8a3fd1', '#0f8a3c', '#c75a00', '#0093b8', '#c2187a'];
const fontStack = (f) => `'${f}', 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif`;

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
let nRows = 0;
let nCols = 0;
let clip = null;
let painter = null;
let drag = null;
let pendingKey = null;
let ac = null; // 자동 완성 상태
let lastFill = '#ffff00';
let lastFont = '#ff0000';
let lastBorder = 'bottom';
let ribbon;
let tinted = [];
let highlighted = [];

// ───────────────────────── 유틸 ─────────────────────────
const clamp = (v, lo, hi) => Math.max(lo, Math.min(hi, v));
const norm = (a, b) => ({ r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) });
const isSingle = (rg) => rg.r1 === rg.r2 && rg.c1 === rg.c2;
const meta = () => ({ si, sel: { ...sel }, active: { ...active }, selKind });
const measureCtx = document.createElement('canvas').getContext('2d');
const escapeHtml = (s) => s.replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

function fontCss(st = {}) {
  return `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}${st.size || DEFAULT_SIZE}pt ${fontStack(st.font || DEFAULT_FONT)}`;
}
function measureText(text, st) {
  measureCtx.font = fontCss(st);
  return measureCtx.measureText(text).width;
}
const styleAt = (r, c) => wb.getCell(si, r, c)?.style ?? {};
const valueAt = (r, c) => wb.getValue(si, r, c);
const isEmptyAt = (r, c) => !wb.getCell(si, r, c)?.raw;

/** 병합 영역을 포함하도록 범위 확장 */
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

/** 서식을 적용할 범위 (전체 행/열 선택 시 화면에 있는 범위로 제한) */
function styleRange() {
  return { r1: sel.r1, c1: sel.c1, r2: Math.min(sel.r2, nRows - 1), c2: Math.min(sel.c2, nCols - 1) };
}
function* cellsIn(rg) {
  for (let r = rg.r1; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) yield [r, c];
}
/** 데이터가 있는 범위와 교차 (통계 등 계산량 제한용) */
function usedClip(rg) {
  const u = wb.usedRange(si);
  return { r1: rg.r1, c1: rg.c1, r2: Math.min(rg.r2, u.rows - 1), c2: Math.min(rg.c2, u.cols - 1) };
}

// ───────────────────────── 그리드 구성 ─────────────────────────
function rowHeaderWidth() {
  if (!view.showHeaders) return 0;
  return Math.max(ROW_HEADER_MIN, String(nRows).length * 8 + 12);
}

function buildGrid() {
  const sheet = wb.sheets[si];
  const cols = [`<col style="width:${rowHeaderWidth()}px">`];
  const heads = ['<th class="corner" title="모두 선택"></th>'];
  let total = rowHeaderWidth();
  for (let c = 0; c < nCols; c++) {
    const w = sheet.colWidths[c] ?? DEFAULT_COL_WIDTH;
    total += w;
    cols.push(`<col style="width:${w}px">`);
    heads.push(`<th>${c ? '<span class="cwl"></span>' : ''}${colToName(c)}<span class="cw"></span></th>`);
  }
  dom.colgroup.innerHTML = cols.join('');
  dom.thead.innerHTML = `<tr>${heads.join('')}</tr>`;
  dom.tbody.innerHTML = rowsHtml(0, nRows);
  dom.grid.style.width = `${total}px`;
  tinted = [];
  highlighted = [];
}

function rowsHtml(from, to) {
  const sheet = wb.sheets[si];
  const tds = '<td></td>'.repeat(nCols);
  let html = '';
  for (let r = from; r < to; r++) {
    const h = sheet.rowHeights[r] ?? DEFAULT_ROW_HEIGHT;
    html += `<tr style="height:${h}px"><th>${r ? '<span class="rht"></span>' : ''}${r + 1}<span class="rh"></span></th>${tds}</tr>`;
  }
  return html;
}

function ensureGridSize(minR = 0, minC = 0) {
  const ext = wb.extent(si);
  const needR = Math.min(LIMIT_R, Math.max(100, ext.rows + 20, minR + 20, active.r + 20, nRows));
  const needC = Math.min(LIMIT_C, Math.max(26, ext.cols + 5, minC + 5, active.c + 5, nCols));
  if (needC > nCols || String(needR).length !== String(nRows).length) {
    nRows = needR;
    nCols = needC;
    buildGrid();
    return true;
  }
  if (needR > nRows) {
    dom.tbody.insertAdjacentHTML('beforeend', rowsHtml(nRows, needR));
    nRows = needR;
    return true;
  }
  return false;
}

function resetGrid() {
  nRows = 0;
  nCols = 0;
  ensureGridSize();
}

function applySizes() {
  if (!nRows) return;
  const sheet = wb.sheets[si];
  const colEls = dom.colgroup.children;
  let total = rowHeaderWidth();
  colEls[0].style.width = `${rowHeaderWidth()}px`;
  for (let c = 0; c < nCols; c++) {
    const w = sheet.colWidths[c] ?? DEFAULT_COL_WIDTH;
    total += w;
    colEls[c + 1].style.width = `${w}px`;
  }
  dom.grid.style.width = `${total}px`;
  const rows = dom.tbody.rows;
  for (let r = 0; r < nRows; r++) {
    const h = `${sheet.rowHeights[r] ?? DEFAULT_ROW_HEIGHT}px`;
    if (rows[r].style.height !== h) rows[r].style.height = h;
  }
}

const tdAt = (r, c) => dom.tbody.rows[r]?.cells[c + 1];
function tdPos(td) {
  return { r: td.parentElement.sectionRowIndex, c: td.cellIndex - 1 };
}

/** 범위의 위치 (grid-inner 좌표) */
function rangeRect(rg) {
  const r1 = clamp(rg.r1, 0, nRows - 1);
  const r2 = clamp(rg.r2, 0, nRows - 1);
  const c1 = clamp(rg.c1, 0, nCols - 1);
  const c2 = clamp(rg.c2, 0, nCols - 1);
  const ths = dom.thead.rows[0].cells;
  const trs = dom.tbody.rows;
  const a = ths[c1 + 1];
  const b = ths[c2 + 1];
  const top = trs[r1].offsetTop + dom.grid.offsetTop;
  return {
    x: a.offsetLeft,
    y: top,
    w: b.offsetLeft + b.offsetWidth - a.offsetLeft,
    h: trs[r2].offsetTop + trs[r2].offsetHeight + dom.grid.offsetTop - top,
  };
}

function placeBox(node, rect, show = true) {
  if (!show) { node.style.display = 'none'; return; }
  Object.assign(node.style, {
    display: 'block', left: `${rect.x}px`, top: `${rect.y}px`, width: `${rect.w + 1}px`, height: `${rect.h + 1}px`,
  });
}

// ───────────────────────── 조건부 서식 ─────────────────────────
function scaleColor(colors, t) {
  const hex = (h) => [1, 3, 5].map((i) => parseInt(h.slice(i, i + 2), 16));
  const seg = colors.length - 1;
  const pos = clamp(t, 0, 1) * seg;
  const i = Math.min(seg - 1, Math.floor(pos));
  const f = pos - i;
  const a = hex(colors[i]);
  const b = hex(colors[i + 1]);
  return `rgb(${a.map((x, k) => Math.round(x + (b[k] - x) * f)).join(',')})`;
}

function prepareCond() {
  const sheet = wb.sheets[si];
  return sheet.cond.map((rule) => {
    const rg = usedClip(rule);
    const nums = [];
    const counts = new Map();
    for (const [r, c] of cellsIn(rg)) {
      const v = valueAt(r, c);
      if (typeof v === 'number') nums.push(v);
      if (v !== null && v !== '' && !isError(v)) {
        const k = typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${v}`;
        counts.set(k, (counts.get(k) ?? 0) + 1);
      }
    }
    const sorted = [...nums].sort((a, b) => b - a);
    return {
      rule, counts,
      min: nums.length ? Math.min(...nums) : 0,
      max: nums.length ? Math.max(...nums) : 0,
      avg: nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : 0,
      topCut: sorted[Math.min(sorted.length, Number(rule.v1) || 10) - 1],
    };
  });
}

const condValue = (s) => parseInput(String(s ?? '')).value;

function condMatch(prep, v) {
  const { rule } = prep;
  if (v === null || v === '' || isError(v)) return false;
  const a = condValue(rule.v1);
  const b = condValue(rule.v2);
  switch (rule.type) {
    case 'gt': return typeof v === 'number' && typeof a === 'number' && v > a;
    case 'lt': return typeof v === 'number' && typeof a === 'number' && v < a;
    case 'between': return typeof v === 'number' && v >= Math.min(a, b) && v <= Math.max(a, b);
    case 'eq': return compareValues(v, a) === 0 && typeof v === typeof a;
    case 'text': return String(typeof v === 'number' ? formatGeneral(v) : v).toLowerCase().includes(String(rule.v1).toLowerCase());
    case 'dup': case 'unique': {
      const k = typeof v === 'string' ? `s:${v.toLowerCase()}` : `${typeof v}:${v}`;
      const n = prep.counts.get(k) ?? 0;
      return rule.type === 'dup' ? n > 1 : n === 1;
    }
    case 'top': return typeof v === 'number' && prep.topCut !== undefined && v >= prep.topCut;
    case 'aboveAvg': return typeof v === 'number' && v > prep.avg;
    case 'belowAvg': return typeof v === 'number' && v < prep.avg;
    default: return false;
  }
}

// ───────────────────────── 셀 렌더링 ─────────────────────────
function renderCells(from = 0, to = nRows) {
  const sheet = wb.sheets[si];
  const spans = new Map();
  const covered = new Set();
  for (const m of sheet.merges) {
    spans.set(`${m.r1},${m.c1}`, m);
    for (const [r, c] of cellsIn(m)) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
  }
  const cond = sheet.cond.length ? prepareCond() : null;
  const rows = dom.tbody.rows;
  for (let r = from; r < to; r++) {
    const cellsEls = rows[r].cells;
    for (let c = 0; c < nCols; c++) {
      paintCell(cellsEls[c + 1], r, c, spans, covered, cond, sheet);
    }
  }
}

function paintCell(td, r, c, spans, covered, cond, sheet) {
  const key = `${r},${c}`;
  if (covered.has(key)) {
    if (!td._cov) { td.style.cssText = 'display:none'; td._cov = true; td._sig = null; }
    return;
  }
  if (td._cov) td._cov = false;
  const span = spans.get(key);
  const rs = span ? span.r2 - span.r1 + 1 : 1;
  const cs = span ? span.c2 - span.c1 + 1 : 1;
  if (td.rowSpan !== rs) td.rowSpan = rs;
  if (td.colSpan !== cs) td.colSpan = cs;

  const cell = wb.getCell(si, r, c);
  const below = wb.getCell(si, r + rs, c)?.style;
  const right = wb.getCell(si, r, c + cs)?.style;
  if (!cell && !cond && !below?.bt && !right?.bl) {
    if (td._sig !== '') {
      td.textContent = '';
      td.style.cssText = '';
      td.className = '';
      td.removeAttribute('title');
      td._sig = '';
    }
    return;
  }

  let st = cell?.style ?? {};
  const v = valueAt(r, c);
  let bar = null;
  if (cond) {
    for (const prep of cond) {
      const rule = prep.rule;
      if (r < rule.r1 || r > rule.r2 || c < rule.c1 || c > rule.c2) continue;
      if (rule.type === 'bar') {
        if (typeof v === 'number' && prep.max !== prep.min) bar = { pct: ((v - Math.min(0, prep.min)) / (prep.max - Math.min(0, prep.min))) * 100, color: rule.color };
        else if (typeof v === 'number') bar = { pct: 100, color: rule.color };
      } else if (rule.type === 'scale') {
        if (typeof v === 'number') st = { ...st, fill: scaleColor(rule.colors, prep.max === prep.min ? 0.5 : (v - prep.min) / (prep.max - prep.min)) };
      } else if (condMatch(prep, v)) {
        st = { ...st, ...rule.style };
      }
    }
  }

  let text;
  let align;
  if (view.showFormulas && cell?.formula) {
    text = cell.raw;
    align = 'left';
  } else {
    ({ text, align } = formatValue(v, st));
  }
  const effAlign = st.align || align;
  const css = [];
  if (st.bold) css.push('font-weight:700');
  if (st.italic) css.push('font-style:italic');
  if (st.underline || st.strike) css.push(`text-decoration:${st.underline ? 'underline ' : ''}${st.strike ? 'line-through' : ''}`);
  if (st.color) css.push(`color:${st.color}`);
  if (st.font) css.push(`font-family:${fontStack(st.font)}`);
  if (st.size) css.push(`font-size:${st.size}pt`);
  if (effAlign !== 'left') css.push(`text-align:${effAlign}`);
  if (st.valign === 'top' || st.valign === 'middle') css.push(`vertical-align:${st.valign}`);
  if (st.indent) css.push(`padding-${effAlign === 'right' ? 'right' : 'left'}:${3 + st.indent * 9}px`);
  if (bar) {
    css.push(`background:linear-gradient(90deg, ${bar.color} ${bar.pct}%, transparent ${bar.pct}%) no-repeat 0 50% / 100% 72%${st.fill ? `, ${st.fill}` : ''}`);
  } else if (st.fill) {
    css.push(`background:${st.fill}`);
  }
  if (st.bb || below?.bt) css.push('border-bottom-color:#000');
  if (st.br || right?.bl) css.push('border-right-color:#000');
  const shadows = [];
  if (r === 0 && st.bt) shadows.push('inset 0 1px 0 #000');
  if (c === 0 && st.bl) shadows.push('inset 1px 0 0 #000');
  if (shadows.length) css.push(`box-shadow:${shadows.join(',')}`);

  const cls = [];
  if (st.wrap) cls.push('wrap');
  else if (text && effAlign === 'left' && typeof v !== 'number' && isEmptyAt(r, c + cs) && !covered.has(`${r},${c + cs}`)) cls.push('ovf');
  if (cell?.comment) cls.push('has-comment');

  // 열 너비보다 긴 숫자는 ### 표시
  if (typeof v === 'number' && text && !st.wrap && !view.showFormulas) {
    let w = -6;
    for (let k = 0; k < cs; k++) w += sheet.colWidths[c + k] ?? DEFAULT_COL_WIDTH;
    if (measureText(text, st) > w) {
      const hw = measureText('#', st);
      text = '#'.repeat(Math.max(1, Math.floor(w / hw)));
    }
  }

  const cssText = css.join(';');
  const className = cls.join(' ');
  const sig = `${text}\u0001${cssText}\u0001${className}\u0001${cell?.comment ?? ''}`;
  if (td._sig === sig) return;
  td._sig = sig;
  td.textContent = text;
  td.style.cssText = cssText;
  td.className = className;
  if (cell?.comment) td.dataset.comment = cell.comment;
  else delete td.dataset.comment;
}

// ───────────────────────── 선택 영역 ─────────────────────────
function setSel(rg, kind = 'cells') {
  sel = expandMerges(rg);
  selKind = kind;
}

function selectCell(r, c, { keepTab = false, scroll = true } = {}) {
  r = clamp(r, 0, LIMIT_R - 1);
  c = clamp(c, 0, LIMIT_C - 1);
  const m = wb.mergeAt(si, r, c);
  if (m) { r = m.r1; c = m.c1; }
  active = { r, c };
  anchor = { r, c };
  focusCell = { r, c };
  setSel({ r1: r, c1: c, r2: r, c2: c });
  if (!keepTab) tabStartCol = null;
  if (ensureGridSize(r, c)) renderCells();
  renderSelection();
  if (scroll) ensureVisible(r, c);
}

function extendTo(r, c, { scroll = true } = {}) {
  r = clamp(r, 0, LIMIT_R - 1);
  c = clamp(c, 0, LIMIT_C - 1);
  focusCell = { r, c };
  setSel(norm(anchor, focusCell));
  tabStartCol = null;
  if (ensureGridSize(r, c)) renderCells();
  renderSelection();
  if (scroll) ensureVisible(r, c);
}

function selectRange(rg, kind = 'cells', act = { r: rg.r1, c: rg.c1 }) {
  active = { ...act };
  anchor = { ...act };
  focusCell = { r: rg.r2, c: rg.c2 };
  setSel(rg, kind);
  if (ensureGridSize(Math.min(rg.r2, nRows), Math.min(rg.c2, nCols))) renderCells();
  renderSelection();
}

function selectAll() {
  selectRange({ r1: 0, c1: 0, r2: MAX_ROWS - 1, c2: MAX_COLS - 1 }, 'all', { r: active.r, c: active.c });
}

function selectCols(c1, c2, act) {
  selectRange({ r1: 0, c1: Math.min(c1, c2), r2: MAX_ROWS - 1, c2: Math.max(c1, c2) }, 'cols', act ?? { r: 0, c: c1 });
}

function selectRows(r1, r2, act) {
  selectRange({ r1: Math.min(r1, r2), c1: 0, r2: Math.max(r1, r2), c2: MAX_COLS - 1 }, 'rows', act ?? { r: r1, c: 0 });
}

const inSel = (r, c) => r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
const selIsActiveOnly = () => {
  const m = wb.mergeAt(si, active.r, active.c);
  return isSingle(sel) || (m && m.r1 === sel.r1 && m.c1 === sel.c1 && m.r2 === sel.r2 && m.c2 === sel.c2);
};

function renderSelection() {
  if (!nRows) return;
  const rect = rangeRect(sel);
  placeBox(dom.selBorder, rect);
  const showHandle = !editing && selKind === 'cells';
  if (showHandle) {
    dom.fillHandle.style.display = 'block';
    dom.fillHandle.style.left = `${rect.x + rect.w - 3}px`;
    dom.fillHandle.style.top = `${rect.y + rect.h - 3}px`;
  } else {
    dom.fillHandle.style.display = 'none';
  }

  // 선택 영역 음영
  for (const td of tinted) td.classList.remove('insel');
  tinted = [];
  if (!selIsActiveOnly()) {
    const rg = styleRange();
    const am = wb.mergeAt(si, active.r, active.c) ?? { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
    for (const [r, c] of cellsIn(rg)) {
      if (r >= am.r1 && r <= am.r2 && c >= am.c1 && c <= am.c2) continue;
      const td = tdAt(r, c);
      if (td) { td.classList.add('insel'); tinted.push(td); }
    }
  }

  // 머리글 강조
  for (const th of highlighted) th.classList.remove('hl', 'full');
  highlighted = [];
  const ths = dom.thead.rows[0].cells;
  const colCls = selKind === 'cols' || selKind === 'all' ? 'full' : 'hl';
  const rowCls = selKind === 'rows' || selKind === 'all' ? 'full' : 'hl';
  for (let c = sel.c1; c <= Math.min(sel.c2, nCols - 1); c++) { ths[c + 1].classList.add(colCls); highlighted.push(ths[c + 1]); }
  for (let r = sel.r1; r <= Math.min(sel.r2, nRows - 1); r++) {
    const th = dom.tbody.rows[r].cells[0];
    th.classList.add(rowCls);
    highlighted.push(th);
  }

  // 복사 테두리
  if (clip && clip.si === si) placeBox(dom.marquee, rangeRect(clip));
  else placeBox(dom.marquee, null, false);

  if (document.activeElement !== dom.nameBox) dom.nameBox.value = cellName(active.r, active.c);
  if (!editing) {
    const cell = wb.getCell(si, active.r, active.c);
    dom.formula.value = cell?.raw ?? '';
  }
  dom.fxCancel.disabled = !editing;
  dom.fxEnter.disabled = !editing;
  positionEditor();
  renderRefHighlights();
  updateStats();
  updateRibbon();
}

function ensureVisible(r, c) {
  const sc = dom.scroll.getBoundingClientRect();
  const tr = dom.tbody.rows[clamp(r, 0, nRows - 1)];
  const th = dom.thead.rows[0].cells[clamp(c, 0, nCols - 1) + 1];
  if (!tr || !th) return;
  const headH = dom.thead.getBoundingClientRect().height;
  const headW = dom.thead.rows[0].cells[0].getBoundingClientRect().width;
  const rr = tr.getBoundingClientRect();
  const cr = th.getBoundingClientRect();
  const viewBottom = sc.top + dom.scroll.clientHeight;
  const viewRight = sc.left + dom.scroll.clientWidth;
  if (rr.top < sc.top + headH) dom.scroll.scrollTop -= sc.top + headH - rr.top;
  else if (rr.bottom > viewBottom) dom.scroll.scrollTop += rr.bottom - viewBottom;
  if (cr.left < sc.left + headW) dom.scroll.scrollLeft -= sc.left + headW - cr.left;
  else if (cr.right > viewRight) dom.scroll.scrollLeft += Math.min(cr.right - viewRight, cr.left - sc.left - headW);
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
  const fmt = (n) => formatValue(n, fmtStyle?.numFmt && fmtStyle.numFmt !== 'general' ? fmtStyle : { numFmt: 'general' }).text;
  const items = [];
  if (numCount) items.push(`평균: ${fmt(sum / numCount)}`);
  items.push(`개수: ${count}`);
  if (numCount) items.push(`합계: ${fmt(sum)}`);
  for (const t of items) dom.stats.append(el('span', {}, t));
}

// ───────────────────────── 이동 ─────────────────────────
function maxNavRow() { return Math.max(nRows - 1, wb.usedRange(si).rows - 1); }
function maxNavCol() { return Math.max(nCols - 1, wb.usedRange(si).cols - 1); }

function stepFrom(pos, dr, dc) {
  const m = wb.mergeAt(si, pos.r, pos.c);
  let { r, c } = pos;
  if (m) {
    if (dr > 0) r = m.r2;
    if (dc > 0) c = m.c2;
    if (dr < 0) r = m.r1;
    if (dc < 0) c = m.c1;
  }
  return { r: clamp(r + dr, 0, LIMIT_R - 1), c: clamp(c + dc, 0, LIMIT_C - 1) };
}

/** Ctrl+방향키: 데이터 영역 끝으로 이동 */
function jump(pos, dr, dc) {
  const limR = maxNavRow();
  const limC = maxNavCol();
  const inside = (r, c) => r >= 0 && c >= 0 && r <= limR && c <= limC;
  let { r, c } = pos;
  const filled = (rr, cc) => !isEmptyAt(rr, cc);
  if (!inside(r + dr, c + dc)) return { r, c };
  if (filled(r, c) && filled(r + dr, c + dc)) {
    while (inside(r + dr, c + dc) && filled(r + dr, c + dc)) { r += dr; c += dc; }
    return { r, c };
  }
  r += dr;
  c += dc;
  while (inside(r, c) && !filled(r, c)) {
    if (!inside(r + dr, c + dc)) return { r, c };
    r += dr;
    c += dc;
  }
  return { r, c };
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

/** Enter/Tab: 여러 셀이 선택되어 있으면 선택 영역 안에서 이동 */
function moveEnterTab(dir) {
  if (!selIsActiveOnly() && selKind === 'cells') {
    const rows = sel.r2 - sel.r1 + 1;
    const cols = sel.c2 - sel.c1 + 1;
    let i = (active.r - sel.r1) * cols + (active.c - sel.c1);
    let j = (active.c - sel.c1) * rows + (active.r - sel.r1);
    const total = rows * cols;
    let r;
    let c;
    if (dir === 'right' || dir === 'left') {
      i = (i + (dir === 'right' ? 1 : -1) + total) % total;
      r = sel.r1 + Math.floor(i / cols);
      c = sel.c1 + (i % cols);
    } else {
      j = (j + (dir === 'down' ? 1 : -1) + total) % total;
      c = sel.c1 + Math.floor(j / rows);
      r = sel.r1 + (j % rows);
    }
    active = { r, c };
    anchor = { r, c };
    renderSelection();
    ensureVisible(r, c);
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

function pageRows() {
  return Math.max(1, Math.floor(dom.scroll.clientHeight / (DEFAULT_ROW_HEIGHT * (view.zoom / 100))) - 2);
}

// ───────────────────────── 편집 ─────────────────────────
const edInput = () => (document.activeElement === dom.formula ? dom.formula : dom.editor);

function focusGrid() {
  if (isDialogOpen()) return;
  if (editing?.fromBar) { dom.formula.focus(); return; }
  if (document.activeElement !== dom.editor) dom.editor.focus({ preventScroll: true });
}

function startEdit(mode, text = null, { fromBar = false, caret = null } = {}) {
  if (editing) return;
  const { r, c } = active;
  const raw = wb.getRaw(si, r, c);
  editing = { r, c, mode, original: raw, point: null, fromBar };
  const value = text ?? raw;
  dom.editor.value = value;
  dom.formula.value = value;
  dom.editor.classList.remove('idle');
  renderSelection();
  const inp = fromBar ? dom.formula : dom.editor;
  if (!fromBar) dom.editor.focus({ preventScroll: true });
  const pos = caret ?? value.length;
  if (!fromBar) inp.setSelectionRange(pos, pos);
  setMode();
  ensureVisible(r, c);
  afterEditInput();
}

/** 입력 대기 중인 에디터에 글자가 입력되면 편집 시작 */
function beginTyping() {
  if (editing) return;
  const { r, c } = active;
  editing = { r, c, mode: 'enter', original: wb.getRaw(si, r, c), point: null, fromBar: false };
  dom.editor.classList.remove('idle');
  dom.formula.value = dom.editor.value;
  renderSelection();
  setMode();
  ensureVisible(r, c);
}

function setMode() {
  if (!editing) dom.status.textContent = painter ? '서식 복사' : clip ? '대상을 선택한 후 Enter 키를 누르거나 붙여넣기를 선택하세요.' : '준비';
  else if (canPoint()) dom.status.textContent = '참조';
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
  dom.refs.replaceChildren();
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
      alertDialog('Tabula', '입력한 수식에 문제가 있습니다. 수식을 확인하세요.').then(() => {
        if (editing) { focusGrid(); }
      });
      return false;
    }
  }
  const { r, c, original } = editing;
  endEditUI();
  const multi = fillSel && !selIsActiveOnly();
  if (text !== original || multi) {
    wb.transact(() => {
      if (multi) {
        for (const [rr, cc] of cellsIn(styleRange())) {
          wb.setInput(si, rr, cc, text.startsWith('=') ? shiftFormula(text, rr - r, cc - c) : text);
        }
      } else {
        wb.setInput(si, r, c, text);
        if (text.includes('\n')) wb.setStyle(si, r, c, { wrap: true });
        else if (!text.startsWith('=') || typeof valueAt(r, c) === 'number') autoWiden({ r1: r, c1: c, r2: r, c2: c });
      }
    }, meta());
  }
  renderAll();
  if (dir && !multi) moveEnterTab(dir);
  focusGrid();
  return true;
}

function commitAndArrow(dr, dc) {
  if (commitEdit()) move(dr, dc);
}

function cancelEdit() {
  if (!editing) return;
  endEditUI();
  renderSelection();
  focusGrid();
}

function positionEditor() {
  const ed = dom.editor;
  const target = editing ?? active;
  if (!nRows) return;
  const m = wb.mergeAt(si, target.r, target.c);
  const rect = rangeRect(m ?? { r1: target.r, c1: target.c, r2: target.r, c2: target.c });
  ed.style.left = `${rect.x + 1}px`;
  ed.style.top = `${rect.y + 1}px`;
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
  const maxW = Math.max(rect.w, dom.inner.offsetWidth - rect.x - 4);
  const w = st.wrap ? rect.w - 1 : clamp(textW, rect.w - 1, maxW);
  ed.style.width = `${w}px`;
  ed.style.height = '0px';
  ed.style.height = `${Math.max(rect.h - 1, ed.scrollHeight)}px`;
}

// ── 참조 모드 (수식 입력 중 셀 클릭/방향키로 참조 삽입) ──
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
  if (editing.si !== undefined && editing.si !== si) ref = `${quoteSheetName(wb.sheets[si].name)}!${ref}`;
  const value = text.slice(0, start) + ref + text.slice(end);
  const anchorPt = editing.point?.anchor ?? { r: rg.r1, c: rg.c1 };
  setEditText(value, start + ref.length);
  editing.point = { start, end: start + ref.length, anchor: anchorPt, cur: { r: rg.r2, c: rg.c2 }, rg };
  setMode();
  renderRefHighlights();
}

function pointMove(dr, dc, extend) {
  const p = editing.point;
  const cur = p?.cur ?? { r: editing.r, c: editing.c };
  const next = { r: clamp(cur.r + dr, 0, LIMIT_R - 1), c: clamp(cur.c + dc, 0, LIMIT_C - 1) };
  if (extend && p) {
    insertPointRef(norm(p.anchor, next));
    editing.point.cur = next;
  } else {
    if (p) editing.point.anchor = next;
    insertPointRef({ r1: next.r, c1: next.c, r2: next.r, c2: next.c });
    editing.point.anchor = next;
  }
  if (ensureGridSize(next.r, next.c)) renderCells();
  ensureVisible(next.r, next.c);
}

/** F4: 커서 위치 참조의 절대/상대 전환 */
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
  dom.refs.replaceChildren();
  if (!editing) return;
  const refs = listRefs(dom.editor.value);
  const colorFor = new Map();
  for (const ref of refs) {
    if (ref.sheet && ref.sheet.toLowerCase() !== wb.sheets[si].name.toLowerCase()) continue;
    const k = `${ref.r1},${ref.c1},${ref.r2},${ref.c2}`;
    if (!colorFor.has(k)) colorFor.set(k, REF_COLORS[colorFor.size % REF_COLORS.length]);
    if (ref.r1 >= nRows || ref.c1 >= nCols) continue;
    const box = el('div', { class: 'ref-box' });
    box.style.borderColor = colorFor.get(k);
    box.style.background = `${colorFor.get(k)}14`;
    placeBox(box, rangeRect({ r1: ref.r1, c1: ref.c1, r2: Math.min(ref.r2, nRows - 1), c2: Math.min(ref.c2, nCols - 1) }));
    dom.refs.append(box);
  }
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
    onmousedown: (e) => { e.preventDefault(); acceptAutocomplete(name); },
    title: FUNC_INFO[name]?.desc ?? '',
  }, el('span', {}, name), el('small', {}, FUNC_INFO[name]?.cat ?? ''))));
  const ed = dom.editor;
  dom.ac.style.left = ed.style.left;
  dom.ac.style.top = `${ed.offsetTop + ed.offsetHeight + 2}px`;
  dom.ac.style.display = 'block';
  return undefined;
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

  if (e.altKey && !ctrl && (k === '=' || e.code === 'Equal')) { handled(); run('autosum'); return; }
  if (ctrl) {
    if (e.shiftKey) {
      const byCode = { Digit1: 'fmtNumber', Digit3: 'fmtDate', Digit4: 'fmtCurrency', Digit5: 'fmtPercent', Digit7: 'borderOutside', Backquote: 'fmtGeneral', Semicolon: 'insertTime', Equal: 'insertMenuKey' };
      if (byCode[e.code]) { handled(); run(byCode[e.code]); return; }
    }
    const lower = k.length === 1 ? k.toLowerCase() : k;
    const map = {
      z: 'undo', y: 'redo', b: 'bold', i: 'italic', u: 'underline', 5: 'strike', d: 'fillDown', r: 'fillRight',
      s: 'save', f: 'find', h: 'replace', g: 'goto', p: 'print', o: 'open', 2: 'bold', 3: 'italic', 4: 'underline',
      ';': 'insertDate', '`': 'toggleFormulas', 1: 'formatCells', '-': 'deleteMenuKey', F1: 'toggleRibbon',
    };
    if (lower === 'a') { handled(); selectAll(); return; }
    if (lower === ' ' || e.code === 'Space') { handled(); selectCols(sel.c1, sel.c2, active); return; }
    if (k === 'Home') { handled(); if (e.shiftKey) extendTo(0, 0); else selectCell(0, 0); return; }
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
    return; // Ctrl+C/X/V 는 기본 동작(복사/붙여넣기 이벤트)
  }

  switch (k) {
    case 'ArrowUp': handled(); move(-1, 0, { extend: e.shiftKey }); return;
    case 'ArrowDown': handled(); move(1, 0, { extend: e.shiftKey }); return;
    case 'ArrowLeft': handled(); move(0, -1, { extend: e.shiftKey }); return;
    case 'ArrowRight': handled(); move(0, 1, { extend: e.shiftKey }); return;
    case 'Enter':
      handled();
      if (clip && !e.shiftKey) { pasteInternal('all'); clip = null; renderSelection(); setMode(); return; }
      moveEnterTab(e.shiftKey ? 'up' : 'down');
      return;
    case 'Tab': handled(); moveEnterTab(e.shiftKey ? 'left' : 'right'); return;
    case 'Home': handled(); if (e.shiftKey) extendTo(focusCell.r, 0); else selectCell(active.r, 0); return;
    case 'PageDown': handled(); move(pageRows(), 0, { extend: e.shiftKey }); return;
    case 'PageUp': handled(); move(-pageRows(), 0, { extend: e.shiftKey }); return;
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
      if (clip || painter) { clip = null; painter = null; dom.grid.classList.remove('painting'); renderSelection(); setMode(); }
      return;
    case 'ContextMenu': {
      handled();
      const rect = tdAt(active.r, active.c)?.getBoundingClientRect();
      if (rect) showContextMenu({ x: rect.left + 10, y: rect.bottom }, 'cell');
      return;
    }
    case ' ':
      if (e.shiftKey) { handled(); selectRows(sel.r1, sel.r2, active); }
      return;
    default:
  }
}

// ───────────────────────── 마우스 ─────────────────────────
function cellFromPoint(x, y) {
  const sc = dom.scroll.getBoundingClientRect();
  const headH = dom.thead.getBoundingClientRect().height;
  const headW = dom.thead.rows[0].cells[0].getBoundingClientRect().width;
  const cx = clamp(x, sc.left + headW + 2, sc.left + dom.scroll.clientWidth - 2);
  const cy = clamp(y, sc.top + headH + 2, sc.top + dom.scroll.clientHeight - 2);
  const target = document.elementFromPoint(cx, cy);
  const td = target?.closest?.('#grid td');
  if (td) return tdPos(td);
  // 병합되어 숨겨진 셀 위: 열/행 머리글 좌표로 계산
  const ths = [...dom.thead.rows[0].cells].slice(1);
  let c = ths.findIndex((th) => { const b = th.getBoundingClientRect(); return cx >= b.left && cx < b.right; });
  let r = [...dom.tbody.rows].findIndex((tr) => { const b = tr.getBoundingClientRect(); return cy >= b.top && cy < b.bottom; });
  if (c < 0) c = nCols - 1;
  if (r < 0) r = nRows - 1;
  return { r, c };
}

let autoScrollTimer = null;
let lastMouse = { x: 0, y: 0 };
function startAutoScroll() {
  stopAutoScroll();
  autoScrollTimer = setInterval(() => {
    if (!drag) return stopAutoScroll();
    const sc = dom.scroll.getBoundingClientRect();
    let dx = 0;
    let dy = 0;
    if (lastMouse.y > sc.top + dom.scroll.clientHeight) dy = 20;
    else if (lastMouse.y < sc.top + 20 && drag.type !== 'colSel') dy = -20;
    if (lastMouse.x > sc.left + dom.scroll.clientWidth) dx = 30;
    else if (lastMouse.x < sc.left + 30 && drag.type !== 'rowSel') dx = -30;
    if (!dx && !dy) return undefined;
    if (dy > 0 && nRows < LIMIT_R && dom.scroll.scrollTop + dom.scroll.clientHeight >= dom.scroll.scrollHeight - 30) growRows(Math.min(20, LIMIT_R - nRows));
    dom.scroll.scrollTop += dy;
    dom.scroll.scrollLeft += dx;
    onDragMove(lastMouse.x, lastMouse.y);
    return undefined;
  }, 50);
}
function stopAutoScroll() { clearInterval(autoScrollTimer); autoScrollTimer = null; }

function onGridMouseDown(e) {
  closeMenus();
  const t = e.target;
  if (t.classList.contains('cw') || t.classList.contains('cwl')) {
    e.preventDefault();
    const c = t.parentElement.cellIndex - (t.classList.contains('cwl') ? 2 : 1);
    drag = { type: 'colResize', c, x: e.clientX, w: wb.colWidth(si, c) };
    return;
  }
  if (t.classList.contains('rh') || t.classList.contains('rht')) {
    e.preventDefault();
    const r = t.parentElement.parentElement.sectionRowIndex - (t.classList.contains('rht') ? 1 : 0);
    drag = { type: 'rowResize', r, y: e.clientY, h: dom.tbody.rows[r].offsetHeight };
    return;
  }
  const th = t.closest('th');
  const td = t.closest('td');
  if (!th && !td) return;
  e.preventDefault();

  if (editing && td && canPoint() && e.button === 0) {
    const { r, c } = tdPos(td);
    const m = wb.mergeAt(si, r, c) ?? { r1: r, c1: c, r2: r, c2: c };
    editing.point = e.shiftKey ? editing.point : null;
    if (e.shiftKey && editing.point) insertPointRef(norm(editing.point.anchor, { r, c }));
    else insertPointRef(m);
    editing.point.anchor = { r, c };
    drag = { type: 'point' };
    startAutoScroll();
    return;
  }
  if (editing && !commitEdit()) return;
  focusGrid();

  if (th) {
    if (th.classList.contains('corner')) { selectAll(); return; }
    if (th.parentElement.parentElement === dom.thead) {
      const c = th.cellIndex - 1;
      if (e.button === 2 && sel.c1 <= c && c <= sel.c2 && (selKind === 'cols' || selKind === 'all')) return;
      if (e.shiftKey && selKind === 'cols') selectCols(anchor.c, c, active);
      else selectCols(c, c);
      drag = { type: 'colSel', start: selKind === 'cols' && e.shiftKey ? anchor.c : c };
    } else {
      const r = th.parentElement.sectionRowIndex;
      if (e.button === 2 && sel.r1 <= r && r <= sel.r2 && (selKind === 'rows' || selKind === 'all')) return;
      if (e.shiftKey && selKind === 'rows') selectRows(anchor.r, r, active);
      else selectRows(r, r);
      drag = { type: 'rowSel', start: selKind === 'rows' && e.shiftKey ? anchor.r : r };
    }
    startAutoScroll();
    return;
  }

  const { r, c } = tdPos(td);
  if (e.button === 2) {
    if (!inSel(r, c)) selectCell(r, c, { scroll: false });
    return;
  }
  if (e.shiftKey) extendTo(r, c, { scroll: false });
  else selectCell(r, c, { scroll: false });
  drag = { type: 'select' };
  startAutoScroll();
}

function onDragMove(x, y) {
  if (!drag) return;
  switch (drag.type) {
    case 'select': {
      const { r, c } = cellFromPoint(x, y);
      if (r !== focusCell.r || c !== focusCell.c) {
        extendTo(r, c, { scroll: false });
        if (!selIsActiveOnly()) dom.nameBox.value = `${sel.r2 - sel.r1 + 1}R x ${sel.c2 - sel.c1 + 1}C`;
      }
      break;
    }
    case 'point': {
      const { r, c } = cellFromPoint(x, y);
      insertPointRef(expandMerges(norm(editing.point.anchor, { r, c })));
      break;
    }
    case 'colSel': {
      const { c } = cellFromPoint(x, y);
      selectCols(drag.start, c, { r: 0, c: drag.start });
      break;
    }
    case 'rowSel': {
      const { r } = cellFromPoint(x, y);
      selectRows(drag.start, r, { r: drag.start, c: 0 });
      break;
    }
    case 'fill': {
      const { r, c } = cellFromPoint(x, y);
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
      if (t) placeBox(dom.fillPreview, rangeRect(t));
      else placeBox(dom.fillPreview, null, false);
      break;
    }
    case 'colResize': {
      const w = Math.max(4, drag.w + (x - drag.x) / (view.zoom / 100));
      drag.newW = w;
      dom.colgroup.children[drag.c + 1].style.width = `${w}px`;
      renderSelection();
      break;
    }
    case 'rowResize': {
      const h = Math.max(6, drag.h + (y - drag.y) / (view.zoom / 100));
      drag.newH = h;
      dom.tbody.rows[drag.r].style.height = `${h}px`;
      renderSelection();
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
      placeBox(dom.fillPreview, null, false);
      if (d.target) doFill(d.src, d.target);
      break;
    case 'colResize': {
      if (d.newW === undefined) break;
      const cols = selKind === 'cols' && d.c >= sel.c1 && d.c <= sel.c2 ? [...Array(sel.c2 - sel.c1 + 1)].map((_, i) => sel.c1 + i) : [d.c];
      wb.transact(() => cols.forEach((c) => wb.setColWidth(si, c, d.newW)), meta());
      break;
    }
    case 'rowResize': {
      if (d.newH === undefined) break;
      const rows = selKind === 'rows' && d.r >= sel.r1 && d.r <= sel.r2 ? [...Array(sel.r2 - sel.r1 + 1)].map((_, i) => sel.r1 + i) : [d.r];
      wb.transact(() => rows.forEach((r) => wb.setRowHeight(si, r, d.newH)), meta());
      break;
    }
    case 'select':
      if (painter) applyPainter();
      renderSelection();
      break;
    case 'point':
      focusGrid();
      break;
    default:
  }
}

function onGridDblClick(e) {
  const t = e.target;
  if (t.classList.contains('cw') || t.classList.contains('cwl')) {
    autofitCols([t.parentElement.cellIndex - (t.classList.contains('cwl') ? 2 : 1)]);
    return;
  }
  if (t.classList.contains('rh') || t.classList.contains('rht')) {
    const r = t.parentElement.parentElement.sectionRowIndex - (t.classList.contains('rht') ? 1 : 0);
    wb.transact(() => wb.setRowHeight(si, r, DEFAULT_ROW_HEIGHT), meta());
    return;
  }
  const td = t.closest('td');
  if (!td || editing) return;
  startEdit('edit');
}

function autofitCols(cols) {
  wb.transact(() => {
    for (const c of cols) {
      let w = 0;
      const u = wb.usedRange(si);
      for (let r = 0; r < u.rows; r++) {
        const v = valueAt(r, c);
        if (v === null) continue;
        const cell = wb.getCell(si, r, c);
        if (wb.mergeAt(si, r, c)) continue;
        const text = view.showFormulas && cell?.formula ? cell.raw : formatValue(v, cell?.style).text;
        w = Math.max(w, measureText(text, cell?.style) + 10);
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
          wb.setCellData(si, r, c, gen(k, r));
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

/** Ctrl+D / Ctrl+R */
function fillCopy(dir) {
  const rg = styleRange();
  wb.transact(() => {
    if (dir === 'down') {
      const srcR = rg.r1 === rg.r2 ? rg.r1 - 1 : rg.r1;
      if (srcR < 0) return;
      for (let c = rg.c1; c <= rg.c2; c++) {
        const d = cellData(wb.getCell(si, srcR, c));
        for (let r = srcR + 1; r <= rg.r2; r++) wb.setCellData(si, r, c, d && { ...d, raw: d.raw.startsWith('=') ? shiftFormula(d.raw, r - srcR, 0) : d.raw });
      }
    } else if (dir === 'right') {
      const srcC = rg.c1 === rg.c2 ? rg.c1 - 1 : rg.c1;
      if (srcC < 0) return;
      for (let r = rg.r1; r <= rg.r2; r++) {
        const d = cellData(wb.getCell(si, r, srcC));
        for (let c = srcC + 1; c <= rg.c2; c++) wb.setCellData(si, r, c, d && { ...d, raw: d.raw.startsWith('=') ? shiftFormula(d.raw, 0, c - srcC) : d.raw });
      }
    } else if (dir === 'up') {
      for (let c = rg.c1; c <= rg.c2; c++) {
        const d = cellData(wb.getCell(si, rg.r2, c));
        for (let r = rg.r1; r < rg.r2; r++) wb.setCellData(si, r, c, d && { ...d, raw: d.raw.startsWith('=') ? shiftFormula(d.raw, r - rg.r2, 0) : d.raw });
      }
    } else if (dir === 'left') {
      for (let r = rg.r1; r <= rg.r2; r++) {
        const d = cellData(wb.getCell(si, r, rg.c2));
        for (let c = rg.c1; c < rg.c2; c++) wb.setCellData(si, r, c, d && { ...d, raw: d.raw.startsWith('=') ? shiftFormula(d.raw, 0, c - rg.c2) : d.raw });
      }
    }
  }, meta());
}

// ───────────────────────── 클립보드 ─────────────────────────
function displayText(r, c) {
  const cell = wb.getCell(si, r, c);
  return formatValue(valueAt(r, c), cell?.style).text;
}

function copySelection(cut) {
  const rg = usedClip(styleRange());
  const full = styleRange();
  if (selKind !== 'cells') { full.r2 = Math.max(rg.r2, full.r1); full.c2 = Math.max(rg.c2, full.c1); }
  const data = [];
  const values = [];
  const text = [];
  for (let r = full.r1; r <= full.r2; r++) {
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
  clip = { si, ...full, data, values, cut, text: toDelimited(text, '\t', '\n') };
  const html = `<table>${text.map((row, i) => `<tr>${row.map((t, j) => {
    const st = data[i][j]?.style ?? {};
    const css = [st.bold && 'font-weight:bold', st.italic && 'font-style:italic', st.color && `color:${st.color}`, st.fill && `background:${st.fill}`].filter(Boolean).join(';');
    return `<td${css ? ` style="${css}"` : ''}>${escapeHtml(t)}</td>`;
  }).join('')}</tr>`).join('')}</table>`;
  renderSelection();
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
  const h = clip.r2 - clip.r1 + 1;
  const w = clip.c2 - clip.c1 + 1;
  const ph = transpose ? w : h;
  const pw = transpose ? h : w;
  const tgt = styleRange();
  const sh = tgt.r2 - tgt.r1 + 1;
  const sw = tgt.c2 - tgt.c1 + 1;
  let th = ph;
  let tw = pw;
  if (!clip.cut && selKind === 'cells' && sh % ph === 0 && sw % pw === 0) { th = sh; tw = sw; }
  const cut = clip.cut;
  const src = clip;
  wb.transact(() => {
    if (cut) for (const [r, c] of cellsIn(src)) wb.setCellData(src.si, r, c, null);
    for (let i = 0; i < th; i++) {
      for (let j = 0; j < tw; j++) {
        const a = i % ph;
        const b = j % pw;
        const [ci, cj] = transpose ? [b, a] : [a, b];
        const d = src.data[ci][cj];
        const tr = tgt.r1 + i;
        const tc = tgt.c1 + j;
        const cur = cellData(wb.getCell(si, tr, tc));
        const shifted = d?.raw.startsWith('=') && !cut ? shiftFormula(d.raw, tr - (src.r1 + ci), tc - (src.c1 + cj)) : d?.raw ?? '';
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
  const tgt = styleRange();
  wb.transact(() => {
    if (rows.length === 1 && rows[0].length === 1 && !selIsActiveOnly()) {
      for (const [r, c] of cellsIn(tgt)) wb.setInput(si, r, c, rows[0][0]);
    } else {
      rows.forEach((row, i) => row.forEach((v, j) => wb.setInput(si, tgt.r1 + i, tgt.c1 + j, v)));
    }
  }, meta());
  if (rows.length > 1 || rows[0].length > 1) {
    selectRange({ r1: tgt.r1, c1: tgt.c1, r2: tgt.r1 + rows.length - 1, c2: tgt.c1 + Math.max(...rows.map((r) => r.length)) - 1 });
  }
}

function handlePaste(text) {
  const n = (s) => (s ?? '').replace(/\r\n?/g, '\n').replace(/\n$/, '');
  if (clip && (!text || n(text) === n(clip.text))) { pasteInternal('all'); return; }
  if (text) pasteText(text);
}

async function pasteFromButton(mode = 'all') {
  if (mode !== 'all' && mode !== 'text') { if (clip) pasteInternal(mode); else toast('복사한 셀이 없습니다.'); return; }
  let text = null;
  try { text = await navigator.clipboard.readText(); } catch { /* 권한 없음 */ }
  if (text === null && !clip) { toast('Ctrl+V 를 눌러 붙여넣으세요.'); return; }
  handlePaste(text);
}

// ───────────────────────── 서식 명령 ─────────────────────────
/** 엑셀처럼: 너비를 바꾼 적 없는 열에 숫자가 들어가지 않으면 열을 넓힘 */
function autoWiden(rg) {
  const sheet = wb.sheets[si];
  const u = usedClip(rg);
  for (let c = u.c1; c <= u.c2; c++) {
    if (sheet.colWidths[c] !== undefined) continue;
    let need = 0;
    for (let r = u.r1; r <= u.r2; r++) {
      const v = valueAt(r, c);
      const st = styleAt(r, c);
      if (typeof v !== 'number' || st.wrap || wb.mergeAt(si, r, c)) continue;
      need = Math.max(need, measureText(formatValue(v, st).text, st) + 10);
    }
    if (need > DEFAULT_COL_WIDTH && need < 400 && (need > DEFAULT_COL_WIDTH + 1)) wb.setColWidth(si, c, Math.ceil(need));
  }
}

function applyStyle(patchOrFn, { widen = false } = {}) {
  const rg = styleRange();
  wb.transact(() => {
    for (const [r, c] of cellsIn(rg)) {
      const cur = wb.getCell(si, r, c)?.style ?? {};
      const patch = typeof patchOrFn === 'function' ? patchOrFn(cur, r, c) : patchOrFn;
      if (patch) wb.setStyle(si, r, c, patch);
    }
    if (widen) autoWiden(rg);
  }, meta());
}

const toggleStyle = (key) => {
  const next = !styleAt(active.r, active.c)[key];
  applyStyle({ [key]: next || undefined });
};

function applyBorder(kind) {
  lastBorder = kind;
  const rg = styleRange();
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
  const cur = displayedDecimals(v, st.numFmt, st.decimals);
  const next = clamp(cur + delta, 0, 15);
  applyStyle((s) => ({ decimals: next, numFmt: s.numFmt === 'text' ? 'general' : s.numFmt }), { widen: true });
}

function changeFontSize(dir) {
  const cur = styleAt(active.r, active.c).size || DEFAULT_SIZE;
  const next = dir > 0 ? FONT_SIZES.find((s) => s > cur) ?? cur + 4 : [...FONT_SIZES].reverse().find((s) => s < cur) ?? Math.max(1, cur - 1);
  applyStyle({ size: next === DEFAULT_SIZE ? undefined : next });
}

function capturePainter(sticky) {
  const rg = usedClip(styleRange());
  const full = styleRange();
  const h = Math.max(1, Math.min(full.r2, Math.max(rg.r2, full.r1)) - full.r1 + 1);
  const w = Math.max(1, Math.min(full.c2, Math.max(rg.c2, full.c1)) - full.c1 + 1);
  const styles = [];
  for (let i = 0; i < h; i++) {
    const row = [];
    for (let j = 0; j < w; j++) row.push(styleAt(full.r1 + i, full.c1 + j));
    styles.push(row);
  }
  painter = { styles, sticky };
  dom.grid.classList.add('painting');
  setMode();
  updateRibbon();
}

function applyPainter() {
  const { styles } = painter;
  const h = styles.length;
  const w = styles[0].length;
  const tgt = styleRange();
  const th = selIsActiveOnly() ? h : tgt.r2 - tgt.r1 + 1;
  const tw = selIsActiveOnly() ? w : tgt.c2 - tgt.c1 + 1;
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
    dom.grid.classList.remove('painting');
  }
  setMode();
  updateRibbon();
}

function toggleMerge(kind = 'center') {
  const rg = styleRange();
  const same = wb.mergesIn(si, rg.r1, rg.c1, rg.r2, rg.c2)
    .some((m) => m.r1 === rg.r1 && m.c1 === rg.c1 && m.r2 === rg.r2 && m.c2 === rg.c2);
  if (kind === 'unmerge' || (kind === 'center' && same)) {
    wb.transact(() => wb.unmerge(si, rg.r1, rg.c1, rg.r2, rg.c2), meta());
    renderAll();
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

// ───────────────────────── 데이터 명령 ─────────────────────────
/** 현재 영역: 활성 셀 주위의 연속된 데이터 영역 */
function currentRegion(r, c) {
  const rg = { r1: r, c1: c, r2: r, c2: c };
  const filledRow = (row, c1, c2) => { for (let x = Math.max(0, c1); x <= c2; x++) if (!isEmptyAt(row, x)) return true; return false; };
  const filledCol = (col, r1, r2) => { for (let y = Math.max(0, r1); y <= r2; y++) if (!isEmptyAt(y, col)) return true; return false; };
  for (let changed = true; changed;) {
    changed = false;
    if (rg.r1 > 0 && filledRow(rg.r1 - 1, rg.c1 - 1, rg.c2 + 1)) { rg.r1--; changed = true; }
    if (filledRow(rg.r2 + 1, rg.c1 - 1, rg.c2 + 1)) { rg.r2++; changed = true; }
    if (rg.c1 > 0 && filledCol(rg.c1 - 1, rg.r1 - 1, rg.r2 + 1)) { rg.c1--; changed = true; }
    if (filledCol(rg.c2 + 1, rg.r1 - 1, rg.r2 + 1)) { rg.c2++; changed = true; }
  }
  return rg;
}

function dataRange() {
  if (!selIsActiveOnly()) return usedClip(styleRange());
  return currentRegion(active.r, active.c);
}

/** 첫 행이 머리글인지 추정 */
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

function sortData(ascending, keyCol = active.c, header = null) {
  const rg = dataRange();
  if (rg.r2 <= rg.r1 && isSingle(rg)) return;
  const h = header ?? hasHeader(rg);
  const key = clamp(keyCol, rg.c1, rg.c2);
  if (wb.mergesIn(si, rg.r1, rg.c1, rg.r2, rg.c2).length) {
    alertDialog('Tabula', '병합된 셀이 있으면 정렬할 수 없습니다.');
    return;
  }
  wb.transact(() => wb.sortRange(si, rg.r1 + (h ? 1 : 0), rg.c1, rg.r2, rg.c2, key, ascending), meta());
  selectRange(rg, 'cells', { r: active.r, c: active.c });
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
      const [a, b] = ref.split(':');
      const pa = parseRangeName(`${a}:${b}`);
      editing.point = { start: fn.length + 2, end: fn.length + 2 + ref.length, anchor: { r: pa.r1, c: pa.c1 }, cur: { r: pa.r2, c: pa.c2 }, rg: pa };
      dom.editor.setSelectionRange(fn.length + 2, fn.length + 2 + ref.length);
      setMode();
    }
    return;
  }
  const rg = usedClip(styleRange());
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
    const shown = displayText(r, c);
    const hay = [raw, shown];
    return hay.some((h) => {
      const a = findState.matchCase ? h : h.toLowerCase();
      const b = findState.matchCase ? text : text.toLowerCase();
      return findState.whole ? a === b : a.includes(b);
    });
  };
  const cols = Math.max(1, u.cols);
  const start = active.r * cols + active.c;
  for (let k = 1; k <= total; k++) {
    const idx = (start + k) % total;
    const r = Math.floor(idx / cols);
    const c = idx % cols;
    if (matches(r, c)) { selectCell(r, c); return true; }
  }
  if (!quiet) toast(`'${text}'을(를) 찾을 수 없습니다.`);
  return false;
}

function replaceIn(r, c) {
  const raw = wb.getRaw(si, r, c);
  const flags = findState.matchCase ? 'g' : 'gi';
  const re = findState.whole
    ? new RegExp(`^${findState.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}$`, flags)
    : new RegExp(findState.text.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), flags);
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
  for (const [k, cell] of wb.sheets[si].cells) if (cell.comment) list.push(k.split(',').map(Number));
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
  sheetSel.set(wb.sheets[si], { active, sel, selKind, scroll: [dom.scroll.scrollLeft, dom.scroll.scrollTop] });
  si = i;
  resetGrid();
  renderCells();
  const saved = restore ? sheetSel.get(wb.sheets[si]) : null;
  if (saved) {
    selectRange(saved.sel, saved.selKind, saved.active);
    [dom.scroll.scrollLeft, dom.scroll.scrollTop] = saved.scroll;
  } else {
    dom.scroll.scrollTo(0, 0);
    selectCell(0, 0);
  }
  renderSheetTabs();
}

function renderSheetTabs() {
  dom.sheetTabs.replaceChildren(...wb.sheets.map((s, i) => el('button', {
    class: `sheet-tab${i === si ? ' active' : ''}`,
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
  const blob = new Blob([content], { type });
  const a = el('a', { href: URL.createObjectURL(blob), download: name });
  document.body.append(a);
  a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
}

function sheetToRows(index) {
  const saved = si;
  si = index;
  const u = wb.usedRange(si);
  const rows = [];
  for (let r = 0; r < u.rows; r++) {
    const row = [];
    for (let c = 0; c < u.cols; c++) row.push(displayText(r, c));
    rows.push(row);
  }
  si = saved;
  return rows;
}

function exportCsv() {
  const csv = `﻿${toDelimited(sheetToRows(si))}`;
  download(`${docName}-${wb.sheets[si].name}.csv`, csv, 'text/csv;charset=utf-8');
  toast('CSV 파일로 내보냈습니다.');
}

function saveAs() {
  formDialog('다른 이름으로 저장', [{ name: 'name', label: '파일 이름', value: docName }], ({ name }) => {
    docName = name.trim() || docName;
    updateTitle();
    download(`${docName}.tabula`, JSON.stringify({ app: 'tabula', docName, workbook: wb.serialize() }), 'application/json');
    saveToStorage();
  }, { okLabel: '다운로드' });
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
  const text = await file.text();
  const base = file.name.replace(/\.[^.]+$/, '');
  try {
    if (/\.(json|tabula)$/i.test(file.name)) {
      const data = JSON.parse(text);
      loadWorkbook(data.workbook ?? data, data.docName ?? base);
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

function loadWorkbook(data, name) {
  if (editing) endEditUI();
  wb.load(data);
  docName = name || '통합 문서1';
  si = 0;
  clip = null;
  painter = null;
  sheetSel.clear();
  resetGrid();
  dom.scroll.scrollTo(0, 0);
  renderAll();
  selectCell(0, 0);
  saveToStorage();
}

function newWorkbook(sample) {
  const go = () => {
    if (sample) loadWorkbook(sample.build(), sample.name);
    else {
      const n = Number(/통합 문서(\d+)/.exec(docName)?.[1] ?? 0) + 1;
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

function saveToStorage() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify({ docName, si, autosave, workbook: wb.serialize() }));
    dirty = false;
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
function scheduleAutosave() {
  if (!autosave) return;
  clearTimeout(saveTimer);
  saveTimer = setTimeout(saveToStorage, 500);
}

function openBackstage() {
  const close = () => { stage.remove(); focusGrid(); };
  const main = el('div', { class: 'backstage-main' });
  const showNew = () => main.replaceChildren(
    el('h2', {}, '새로 만들기'),
    el('div', { class: 'backstage-cards' },
      el('button', { class: 'bcard', onclick: () => { close(); newWorkbook(); } }, el('div', { class: 'thumb' }), el('b', {}, '새 통합 문서'), el('small', {}, '빈 시트로 시작')),
      SAMPLES.map((s) => el('button', { class: 'bcard', onclick: () => { close(); newWorkbook(s); } }, el('div', { class: 'thumb' }), el('b', {}, s.name), el('small', {}, s.desc)))),
    el('h2', { style: { marginTop: '36px' } }, '정보'),
    el('div', { class: 'muted' }, `${docName} · 시트 ${wb.sheets.length}개 · ${autosave ? '자동 저장 켜짐 (브라우저에 보관)' : '자동 저장 꺼짐'}`),
  );
  const nav = el('div', { class: 'backstage-nav' },
    el('button', { class: 'back', title: '돌아가기', onclick: close }, '←'),
    el('button', { onclick: showNew }, '새로 만들기'),
    el('button', { onclick: () => { close(); pickFile('open'); } }, '열기'),
    el('button', { onclick: () => { close(); run('save'); } }, '저장'),
    el('button', { onclick: () => { close(); saveAs(); } }, '다른 이름으로 저장'),
    el('button', { onclick: () => { close(); exportCsv(); } }, 'CSV로 내보내기'),
    el('button', { onclick: () => { close(); setTimeout(() => window.print(), 50); } }, '인쇄'),
    el('button', { onclick: close }, '닫기'));
  const stage = el('div', { class: 'backstage' }, nav, main);
  stage.addEventListener('keydown', (e) => { if (e.key === 'Escape') close(); });
  document.body.append(stage);
  showNew();
  nav.children[1].focus();
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
    const rg = styleRange();
    wb.transact(() => {
      if (isCol) for (let c = rg.c1; c <= rg.c2; c++) wb.setColWidth(si, c, n);
      else for (let r = rg.r1; r <= rg.r2; r++) wb.setRowHeight(si, r, n);
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
    sortData(v.order === 'asc', Number(v.col), v.header);
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
  const rg = parseRangeName(t);
  if (!rg) { toast('참조가 올바르지 않습니다.'); return false; }
  if (rg.r2 >= LIMIT_R || rg.c2 >= LIMIT_C) { toast(`이 버전은 ${cellName(LIMIT_R - 1, LIMIT_C - 1)}까지 지원합니다.`); return false; }
  switchSheet(target);
  if (isSingle(rg)) selectCell(rg.r1, rg.c1);
  else { selectRange(rg); ensureVisible(rg.r1, rg.c1); }
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
    { value: '5', label: '빨강 테두리', style: { bold: true, color: '#c00000' } },
  ];
  const [title, note] = titles[type];
  const rg = styleRange();
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
    const rule = { ...rg, type: type === 'dup' ? v.mode : type, v1: v.v1, v2: v.v2, style: presets[Number(v.preset)].style };
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
      if (cell.raw) {
        totalCells++;
        if (i === si) cells++;
      }
      if (cell.formula && i === si) formulas++;
      if (cell.comment && i === si) comments++;
    }
  });
  const u = wb.usedRange(si);
  const rows = [
    ['마지막 셀', u.rows ? cellName(u.rows - 1, u.cols - 1) : '-'],
    ['데이터가 있는 셀', cells], ['수식', formulas], ['메모', comments], ['병합된 셀', wb.sheets[si].merges.length],
    ['시트 수', wb.sheets.length], ['통합 문서 전체 데이터 셀', totalCells],
  ];
  openDialog({
    title: '통합 문서 통계', width: 360,
    body: el('table', { class: 'kbd-table' }, el('tr', {}, el('td', { colspan: 2 }, el('b', {}, `현재 시트: ${wb.sheets[si].name}`))),
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
  ['Ctrl+방향키', '데이터 영역의 끝으로 이동'],
  ['Shift+방향키', '선택 영역 확장'],
  ['Ctrl+Space / Shift+Space', '열 전체 / 행 전체 선택'],
  ['Ctrl+A', '모두 선택'],
  ['Ctrl+Home / Ctrl+End', '처음 셀 / 마지막 셀로 이동'],
  ['Ctrl+C / Ctrl+X / Ctrl+V', '복사 / 잘라내기 / 붙여넣기'],
  ['Ctrl+Z / Ctrl+Y', '실행 취소 / 다시 실행'],
  ['Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5', '굵게 / 기울임꼴 / 밑줄 / 취소선'],
  ['Ctrl+D / Ctrl+R', '아래로 / 오른쪽으로 채우기'],
  ['Alt+=', '자동 합계'],
  ['Ctrl+;  /  Ctrl+Shift+;', '오늘 날짜 / 현재 시간 입력'],
  ['Ctrl+Shift+1/3/4/5', '숫자 / 날짜 / 통화 / 백분율 서식'],
  ['Ctrl+Shift+~', '일반 서식'],
  ['Ctrl+1', '셀 서식'],
  ['Ctrl+F / Ctrl+H / Ctrl+G', '찾기 / 바꾸기 / 이동'],
  ['Ctrl+`', '수식 표시'],
  ['Ctrl+Shift+= / Ctrl+-', '행·열 삽입 / 삭제'],
  ['Ctrl+PageUp / PageDown', '이전 / 다음 시트'],
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
  condFormat: () => [
    { title: '셀 강조 규칙' },
    { label: '보다 큼...', action: () => condRuleDialog('gt') },
    { label: '보다 작음...', action: () => condRuleDialog('lt') },
    { label: '다음 값의 사이에 있음...', action: () => condRuleDialog('between') },
    { label: '같음...', action: () => condRuleDialog('eq') },
    { label: '텍스트 포함...', action: () => condRuleDialog('text') },
    { label: '중복 값...', action: () => condRuleDialog('dup') },
    { title: '상위/하위 규칙' },
    { label: '상위 10개 항목...', action: () => condRuleDialog('top') },
    { label: '평균 초과', action: () => wb.transact(() => wb.addCondRule(si, { ...styleRange(), type: 'aboveAvg', style: { fill: '#ffc7ce', color: '#9c0006' } }), meta()) },
    { label: '평균 미만', action: () => wb.transact(() => wb.addCondRule(si, { ...styleRange(), type: 'belowAvg', style: { fill: '#ffc7ce', color: '#9c0006' } }), meta()) },
    { title: '데이터 막대' },
    ...[['파랑', '#8fb3e8'], ['녹색', '#8fd19e'], ['빨강', '#f19c9c'], ['주황', '#f7c07e']].map(([n, color]) => ({
      label: `${n} 데이터 막대`, icon: `<span style="display:block;width:16px;height:10px;background:${color}"></span>`,
      action: () => wb.transact(() => wb.addCondRule(si, { ...styleRange(), type: 'bar', color }), meta()),
    })),
    { title: '색조' },
    ...[['녹색 - 노랑 - 빨강', ['#63be7b', '#ffeb84', '#f8696b']], ['빨강 - 노랑 - 녹색', ['#f8696b', '#ffeb84', '#63be7b']], ['흰색 - 녹색', ['#fcfcff', '#63be7b']], ['흰색 - 빨강', ['#fcfcff', '#f8696b']]].map(([n, colors]) => ({
      label: n, icon: `<span style="display:block;width:16px;height:10px;background:linear-gradient(90deg,${colors.join(',')})"></span>`,
      action: () => wb.transact(() => wb.addCondRule(si, { ...styleRange(), type: 'scale', colors }), meta()),
    })),
    { sep: true },
    { label: '규칙 지우기 - 선택한 셀', action: () => wb.transact(() => wb.clearCondRules(si, styleRange()), meta()) },
    { label: '규칙 지우기 - 시트 전체', action: () => wb.transact(() => wb.clearCondRules(si), meta()) },
  ],
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
    { label: '행 높이 자동 맞춤', action: () => { const rg = styleRange(); wb.transact(() => { for (let r = rg.r1; r <= rg.r2; r++) wb.setRowHeight(si, r, DEFAULT_ROW_HEIGHT); }, meta()); } },
    { label: '열 너비...', action: () => sizeDialog('col') },
    { label: '열 너비 자동 맞춤', action: () => { const rg = styleRange(); autofitCols([...Array(rg.c2 - rg.c1 + 1)].map((_, i) => rg.c1 + i)); } },
    { title: '시트 구성' },
    { label: '시트 이름 바꾸기', action: () => renameSheetInline(si) },
    { label: '시트 복사본 만들기', icon: 'copy', action: () => run('duplicateSheet') },
    { title: '보호' },
    { label: '셀 서식...', key: 'Ctrl+1', action: () => formatCellsDialog() },
  ],
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
    { label: '중복된 항목 제거', icon: 'dedupe', action: () => removeDuplicates() },
  ],
  find: () => [
    { label: '찾기...', icon: 'search', key: 'Ctrl+F', action: () => openFindDialog('find') },
    { label: '바꾸기...', key: 'Ctrl+H', action: () => openFindDialog('replace') },
    { label: '이동...', key: 'Ctrl+G', action: () => gotoDialog() },
    { sep: true },
    { label: '다음 메모', icon: 'comment', action: () => jumpComment(1) },
  ],
};

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
      { label: '열 너비 자동 맞춤', action: () => autofitCols([...Array(sel.c2 - sel.c1 + 1)].map((_, i) => sel.c1 + i)) },
    );
  } else if (kind === 'row') {
    items.push(
      { label: '삽입', icon: 'rowInsert', action: () => run('insertRows') },
      { label: '삭제', icon: 'delete', action: () => run('deleteRows') },
      { label: '내용 지우기', action: () => run('clearContents') },
      { sep: true },
      { label: '행 높이...', action: () => sizeDialog('row') },
    );
  } else {
    items.push(
      { label: '행 삽입', icon: 'rowInsert', action: () => run('insertRows') },
      { label: '열 삽입', icon: 'colInsert', action: () => run('insertCols') },
      { label: '행 삭제', icon: 'delete', action: () => run('deleteRows') },
      { label: '열 삭제', icon: 'delete', action: () => run('deleteCols') },
      { label: '내용 지우기', action: () => run('clearContents') },
      { sep: true },
      { label: '오름차순 정렬', icon: 'sortAsc', action: () => sortData(true) },
      { label: '내림차순 정렬', icon: 'sortDesc', action: () => sortData(false) },
      { sep: true },
      wb.getCell(si, active.r, active.c)?.comment
        ? { label: '메모 편집', icon: 'comment', action: () => editComment() }
        : { label: '새 메모', icon: 'newComment', action: () => editComment() },
      wb.getCell(si, active.r, active.c)?.comment ? { label: '메모 삭제', icon: 'deleteComment', action: () => run('deleteComment') } : null,
      { sep: true },
      { label: '셀 서식...', key: 'Ctrl+1', action: () => formatCellsDialog() },
      { label: '병합하고 가운데 맞춤', icon: 'merge', action: () => toggleMerge('center'), disabled: selIsActiveOnly() && !wb.mergeAt(si, active.r, active.c) },
    );
  }
  openMenu(pos, items);
}

// ───────────────────────── 명령 ─────────────────────────
const structural = (fn) => () => {
  fn();
  resetGrid();
  renderAll();
};

const COMMANDS = {
  undo: () => { const m = wb.undo(); if (m) restoreMeta(m); else toast('실행 취소할 작업이 없습니다.'); },
  redo: () => { const m = wb.redo(); if (m) restoreMeta(m); },
  save: () => { if (saveToStorage()) toast('저장되었습니다. (이 브라우저에 보관됨)'); else toast('저장하지 못했습니다.'); },
  saveAs,
  open: () => pickFile('open'),
  backstage: openBackstage,
  print: () => window.print(),

  cut: () => { const { text } = copySelection(true); navigator.clipboard?.writeText(text).catch(() => {}); },
  copy: () => { const { text } = copySelection(false); navigator.clipboard?.writeText(text).catch(() => {}); },
  paste: () => pasteFromButton('all'),
  painter: () => {
    if (painter) { painter = null; dom.grid.classList.remove('painting'); setMode(); updateRibbon(); }
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
    const rg = styleRange();
    const n = selKind === 'cols' || selKind === 'all' ? 1 : rg.r2 - rg.r1 + 1;
    wb.transact(() => wb.insertRows(si, rg.r1, n), meta());
  }),
  insertCols: structural(() => {
    const rg = styleRange();
    const n = selKind === 'rows' || selKind === 'all' ? 1 : rg.c2 - rg.c1 + 1;
    wb.transact(() => wb.insertCols(si, rg.c1, n), meta());
  }),
  deleteRows: structural(() => {
    const rg = styleRange();
    wb.transact(() => wb.deleteRows(si, rg.r1, selKind === 'rows' || selKind === 'all' ? Math.min(sel.r2, nRows - 1) - rg.r1 + 1 : rg.r2 - rg.r1 + 1), meta());
    selectCell(rg.r1, active.c);
  }),
  deleteCols: structural(() => {
    const rg = styleRange();
    wb.transact(() => wb.deleteCols(si, rg.c1, rg.c2 - rg.c1 + 1), meta());
    selectCell(active.r, rg.c1);
  }),
  insertMenuKey: () => {
    if (selKind === 'rows') run('insertRows');
    else if (selKind === 'cols') run('insertCols');
    else openMenu(tdAt(active.r, active.c).getBoundingClientRect(), MENUS.insert());
  },
  deleteMenuKey: () => {
    if (selKind === 'rows') run('deleteRows');
    else if (selKind === 'cols') run('deleteCols');
    else {
      const b = tdAt(active.r, active.c).getBoundingClientRect();
      openMenu({ x: b.left, y: b.bottom }, MENUS.delete());
    }
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
    const hasData = wb.sheets[si].cells.size > 0;
    if (hasData) {
      openDialog({
        title: 'Tabula', body: `'${wb.sheets[si].name}' 시트를 영구적으로 삭제합니다. 계속하시겠습니까?`,
        buttons: [{ label: '삭제', primary: true, action: del }, { label: '취소' }],
      });
    } else del();
  },
  duplicateSheet: () => {
    const src = wb.serialize().sheets[si];
    const i = wb.transact(() => {
      const at = wb.addSheet(`${src.name} (2)`.slice(0, 31), si + 1);
      const data = wb.serialize();
      data.sheets[at] = { ...src, name: wb.sheets[at].name };
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
  clearContents: () => { const rg = styleRange(); wb.transact(() => wb.clearRange(si, rg.r1, rg.c1, rg.r2, rg.c2, 'contents'), meta()); },
  clearAll: () => {
    const rg = styleRange();
    wb.transact(() => { wb.clearRange(si, rg.r1, rg.c1, rg.r2, rg.c2, 'all'); wb.unmerge(si, rg.r1, rg.c1, rg.r2, rg.c2); }, meta());
  },
  clearFormats: () => { const rg = styleRange(); wb.transact(() => wb.clearRange(si, rg.r1, rg.c1, rg.r2, rg.c2, 'formats'), meta()); },
  clearComments: () => { const rg = styleRange(); wb.transact(() => wb.clearRange(si, rg.r1, rg.c1, rg.r2, rg.c2, 'comments'), meta()); },
  sortAsc: () => sortData(true),
  sortDesc: () => sortData(false),
  sortDialog,
  dedupe: removeDuplicates,
  find: () => openFindDialog('find'),
  replace: () => openFindDialog('replace'),
  goto: gotoDialog,

  insertFunction: insertFunctionDialog,
  insertDate: () => {
    const d = new Date();
    const text = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
    wb.transact(() => wb.setInput(si, active.r, active.c, text), meta());
  },
  insertTime: () => {
    const d = new Date();
    wb.transact(() => wb.setInput(si, active.r, active.c, `${d.getHours()}:${String(d.getMinutes()).padStart(2, '0')}`), meta());
  },
  editComment,
  deleteComment: () => wb.transact(() => wb.setComment(si, active.r, active.c, ''), meta()),
  prevComment: () => jumpComment(-1),
  nextComment: () => jumpComment(1),
  workbookStats: statsDialog,

  importCsv: () => pickFile('import'),
  exportCsv,

  toggleGrid: (v) => { view.showGrid = v ?? !view.showGrid; applyView(); },
  togglePrintGrid: (v) => { view.printGrid = v ?? !view.printGrid; applyView(); },
  toggleFormulaBar: (v) => { view.showFormulaBar = v ?? !view.showFormulaBar; applyView(); },
  toggleHeaders: (v) => { view.showHeaders = v ?? !view.showHeaders; applyView(); },
  toggleFormulas: () => { view.showFormulas = !view.showFormulas; renderAll(); },
  toggleRibbon: () => ribbon.toggleCollapse(),
  zoomIn: () => setZoom(view.zoom + 10),
  zoomOut: () => setZoom(view.zoom - 10),
  zoom100: () => setZoom(100),
  recalc: () => { wb.invalidate(); renderAll(); },

  shortcuts: () => openDialog({
    title: '바로 가기 키', width: 560,
    body: el('table', { class: 'kbd-table' }, SHORTCUTS.map(([k, d]) => el('tr', {}, el('td', {}, k), el('td', {}, d)))),
    buttons: [{ label: '닫기', primary: true }],
  }),
  about: () => openDialog({
    title: 'Tabula 정보', width: 400,
    body: el('div', { style: { lineHeight: '1.7' } },
      el('b', {}, 'Tabula'), ' — 브라우저에서 동작하는 엑셀 스타일 스프레드시트', el('br'),
      el('span', { class: 'muted' }, `지원 함수 ${FUNCTION_NAMES.length}개 · 데이터는 이 브라우저에만 저장됩니다.`)),
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
  if (!isDialogOpen() && !document.querySelector('.backstage')) focusGrid();
}

function restoreMeta(m) {
  if (m.si !== si && m.si < wb.sheets.length) switchSheet(m.si);
  resetGrid();
  renderCells();
  selectRange(m.sel, m.selKind ?? 'cells', m.active);
  ensureVisible(m.active.r, m.active.c);
}

// ───────────────────────── 보기 ─────────────────────────
function applyView() {
  dom.grid.classList.toggle('nogrid', !view.showGrid);
  dom.grid.classList.toggle('nogrid-print', !view.printGrid);
  dom.grid.classList.toggle('noheaders', !view.showHeaders);
  dom.formulaRow.classList.toggle('hidden', !view.showFormulaBar);
  applySizes();
  renderSelection();
}

function setZoom(z) {
  view.zoom = clamp(Math.round(z), 25, 400);
  dom.inner.style.zoom = view.zoom / 100;
  dom.zoomSlider.value = view.zoom;
  dom.zoomLabel.textContent = `${view.zoom}%`;
  renderSelection();
}

function ribbonState() {
  const st = styleAt(active.r, active.c);
  const fmt = st.numFmt === 'comma' ? 'number' : st.numFmt === 'datetime' ? 'date' : st.numFmt || 'general';
  return {
    bold: st.bold, italic: st.italic, underline: st.underline, strike: st.strike, wrap: st.wrap,
    font: st.font || DEFAULT_FONT, size: String(st.size || DEFAULT_SIZE), numFmt: fmt,
    alignLeft: st.align === 'left', alignCenter: st.align === 'center', alignRight: st.align === 'right',
    valignTop: st.valign === 'top', valignMiddle: st.valign === 'middle', valignBottom: !st.valign,
    merged: !!wb.mergeAt(si, active.r, active.c), painter: !!painter,
    lastFill, lastFont, ...view,
  };
}

function updateRibbon() {
  ribbon?.update(ribbonState());
  dom.undoBtn.disabled = !wb.canUndo();
  dom.redoBtn.disabled = !wb.canRedo();
}

function updateTitle() {
  const t = `${docName}${!autosave && dirty ? '*' : ''} - Tabula`;
  dom.title.textContent = t;
  document.title = t;
  dom.autosave.setAttribute('aria-checked', String(autosave));
  dom.autosaveLabel.textContent = autosave ? '켬' : '끔';
}

function growRows(n) {
  dom.tbody.insertAdjacentHTML('beforeend', rowsHtml(nRows, nRows + n));
  const from = nRows;
  nRows += n;
  if (String(from).length !== String(nRows).length) applySizes();
  renderCells(from, nRows);
}

function renderAll() {
  if (!nRows) resetGrid();
  else ensureGridSize();
  applySizes();
  renderCells();
  renderSelection();
  renderSheetTabs();
  updateTitle();
}

let renderQueued = false;
function onBookChange() {
  dirty = true;
  if (!renderQueued) {
    renderQueued = true;
    queueMicrotask(() => {
      renderQueued = false;
      if (si >= wb.sheets.length) si = wb.sheets.length - 1;
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
    if (!editing) { beginTyping(); }
    editing.point = null;
    dom.formula.value = ed.value;
    afterEditInput();
  });
  ed.addEventListener('mousedown', () => { if (editing) { editing.mode = 'edit'; editing.point = null; setTimeout(afterCaretMove); } });

  const fb = dom.formula;
  fb.addEventListener('focus', () => {
    if (!editing) {
      startEdit('edit', null, { fromBar: true });
    } else {
      editing.fromBar = true;
    }
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
  dom.fxInsert.addEventListener('mousedown', (e) => { e.preventDefault(); run('insertFunction'); });

  dom.grid.addEventListener('mousedown', onGridMouseDown);
  dom.grid.addEventListener('dblclick', onGridDblClick);
  dom.grid.addEventListener('contextmenu', (e) => {
    e.preventDefault();
    const th = e.target.closest('th');
    const kind = th && !th.classList.contains('corner') ? (th.parentElement.parentElement === dom.thead ? 'col' : 'row') : 'cell';
    showContextMenu({ x: e.clientX, y: e.clientY }, kind);
  });
  dom.fillHandle.addEventListener('mousedown', (e) => {
    e.preventDefault();
    e.stopPropagation();
    if (editing && !commitEdit()) return;
    drag = { type: 'fill', src: { ...sel }, target: null };
    startAutoScroll();
  });
  document.addEventListener('mousemove', (e) => {
    lastMouse = { x: e.clientX, y: e.clientY };
    if (drag) onDragMove(e.clientX, e.clientY);
  });
  document.addEventListener('mouseup', onDragEnd);

  // 메모 표시
  dom.grid.addEventListener('mouseover', (e) => {
    const td = e.target.closest?.('td');
    if (td?.dataset.comment) {
      const b = td.getBoundingClientRect();
      dom.tip.textContent = td.dataset.comment;
      dom.tip.style.display = 'block';
      dom.tip.style.left = `${Math.min(b.right + 6, innerWidth - 250)}px`;
      dom.tip.style.top = `${b.top}px`;
    } else {
      dom.tip.style.display = 'none';
    }
  });
  dom.grid.addEventListener('mouseleave', () => { dom.tip.style.display = 'none'; });

  dom.scroll.addEventListener('scroll', () => {
    const s = dom.scroll;
    if (s.scrollTop + s.clientHeight > s.scrollHeight - 300 && nRows < LIMIT_R) growRows(Math.min(50, LIMIT_R - nRows));
    if (s.scrollLeft + s.clientWidth > s.scrollWidth - 200 && nCols < LIMIT_C) {
      nCols = Math.min(LIMIT_C, nCols + 10);
      buildGrid();
      renderCells();
      renderSelection();
    }
  });

  // 클립보드
  document.addEventListener('copy', (e) => {
    if (editing || document.activeElement !== dom.editor) return;
    e.preventDefault();
    const { text, html } = copySelection(false);
    e.clipboardData.setData('text/plain', text);
    e.clipboardData.setData('text/html', html);
  });
  document.addEventListener('cut', (e) => {
    if (editing || document.activeElement !== dom.editor) return;
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

  // 검색 상자 = 빠른 찾기
  dom.search.addEventListener('keydown', (e) => {
    e.stopPropagation();
    if (e.key === 'Enter') { findState.text = dom.search.value; findNext(dom.search.value); }
    if (e.key === 'Escape') { dom.search.value = ''; focusGrid(); }
  });

  // 제목 이름 바꾸기
  dom.title.addEventListener('click', () => {
    const input = el('input', { value: docName });
    dom.title.replaceChildren(input);
    input.focus();
    input.select();
    const done = (ok) => {
      if (ok && input.value.trim()) { docName = input.value.trim(); scheduleAutosave(); dirty = true; }
      updateTitle();
      focusGrid();
    };
    input.addEventListener('keydown', (e) => { e.stopPropagation(); if (e.key === 'Enter') input.blur(); if (e.key === 'Escape') { input.value = ''; input.blur(); } });
    input.addEventListener('blur', () => done(true), { once: true });
  });

  dom.autosave.addEventListener('click', () => {
    autosave = !autosave;
    updateTitle();
    if (autosave) saveToStorage();
    else {
      try {
        const data = JSON.parse(localStorage.getItem(STORAGE_KEY) ?? 'null');
        if (data) { data.autosave = false; localStorage.setItem(STORAGE_KEY, JSON.stringify(data)); }
      } catch { /* 무시 */ }
    }
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

  // 전역 단축키 (포커스가 다른 곳에 있을 때)
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
    if (autosave) saveToStorage();
    else if (dirty) { e.preventDefault(); e.returnValue = ''; }
  });
  window.addEventListener('blur', () => { if (drag) onDragEnd(); });
}

// ───────────────────────── 시작 ─────────────────────────
function init() {
  const stored = loadFromStorage();
  wb = new Workbook(stored?.workbook);
  if (stored) {
    docName = stored.docName || docName;
    si = clamp(stored.si || 0, 0, wb.sheets.length - 1);
    autosave = stored.autosave !== false;
  }
  wb.onChange(onBookChange);
  hydrateIcons();
  ribbon = buildRibbon({ run, openMenu: openNamedMenu, focusGrid, refreshRibbon: updateRibbon });
  bindEvents();
  applyView();
  resetGrid();
  renderAll();
  selectCell(0, 0);
  focusGrid();
  // 테스트/디버깅용
  window.tabula = { wb: () => wb, run, selectCell, get active() { return active; }, get sel() { return sel; }, get si() { return si; } };
}

init();
