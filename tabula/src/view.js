// 가상 스크롤 그리드: 화면에 보이는 행/열만 그림 (20,000,000행 × 16,384열 지원)
// 틀 고정은 4개 창(TL/TR/BL/BR)으로, 각 창은 시트 좌표계 콘텐츠를 transform 으로 이동시켜 표시.
import { Axis } from './axis.js';
import { colToName, MAX_ROWS, MAX_COLS } from './formula.js';
import { formatValue, formatGeneral } from './format.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { renderChartSvg, chartModelData } from './chart.js';
import { shapeSvg } from './shapes.js';
import { validationAt } from './validation.js';
import { prepareCond, condFormatAt, ICON_SVG, EMPTY_MATCH_TYPES, ruleRanges, inRule } from './condfmt.js';
import { tableAt, tableCellStyle, tableFilterRange, styleByName } from './tables.js';
import { slicerCssVars } from './slicerstyle.js';
import { fontAlias } from './fonts.js';

export const DEFAULT_FONT = '맑은 고딕';
export const DEFAULT_SIZE = 11;
export const fontStack = (f) => {
  const name = String(f).replace(/'/g, '');
  const alias = fontAlias(name);
  return `'${name}'${alias ? `, '${alias}'` : ''}, 'Malgun Gothic', '맑은 고딕', 'Apple SD Gothic Neo', 'Noto Sans KR', sans-serif`;
};
const HEAD_H = 20;
const MAX_PX = 15_000_000; // 스크롤 영역 최대 픽셀 (브라우저 한계 회피)
const OVER_R = 12;
const OVER_C = 4;

const measureCtx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
const measureCache = new Map();
export function fontCss(st = {}) {
  return `${st.italic ? 'italic ' : ''}${st.bold ? '700 ' : ''}${st.size || DEFAULT_SIZE}pt ${fontStack(st.font || DEFAULT_FONT)}`;
}
export function measureText(text, st) {
  const font = fontCss(st);
  const k = `${font}\u0001${text}`;
  let w = measureCache.get(k);
  if (w === undefined) {
    measureCtx.font = font;
    w = measureCtx.measureText(text).width;
    if (measureCache.size > 20000) measureCache.clear();
    measureCache.set(k, w);
  }
  return w;
}

/** 열이 좁을 때: 일반 서식은 소수 자릿수를 줄이거나 지수로, 그 밖에는 ### (엑셀과 동일) */
function fitNumber(v, maxW, style) {
  const general = (!style.numFmt || style.numFmt === 'general') && style.decimals === undefined;
  if (general) {
    const abs = Math.abs(v);
    if (abs >= 1e-4 && abs < 1e11 && !Number.isInteger(v)) {
      for (let d = 9; d >= 0; d--) {
        const t = formatGeneral(Number(v.toFixed(d)));
        if (measureText(t, style) <= maxW && t !== '0' && t !== '-0') return t;
      }
    }
    for (let d = 4; d >= 0; d--) {
      const [m, e] = v.toExponential(d).split('e');
      const t = `${m.includes('.') ? m.replace(/\.?0+$/, '') : m}E${e[0] === '-' ? '-' : '+'}${e.replace(/^[+-]/, '').padStart(2, '0')}`;
      if (measureText(t, style) <= maxW) return t;
    }
  }
  return '#'.repeat(Math.max(1, Math.floor(maxW / Math.max(1, measureText('#', style)))));
}

const esc = (s) => String(s).replace(/[&<>"]/g, (ch) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch]));

/**
 * host.state() → { wb, si, sel, selKind, active, editing, clip, fillPreview, refs, chartSel,
 *                  showGrid, showFormulas, showHeaders }
 * host.onViewScroll() : 스크롤 후 호출 (편집기 위치 갱신 등)
 */
export class GridView {
  constructor(host) {
    this.host = host;
    this.wrap = document.getElementById('gridWrap');
    this.scroll = document.getElementById('gridScroll');
    this.sizer = document.getElementById('gridSizer');
    this.viewEl = document.getElementById('gridView');
    this.z = 1;
    this.sx = 0;
    this.sy = 0;
    this.extR = 100;
    this.extC = 26;
    this.hw = 34;
    this.hh = HEAD_H;
    this.cond = null;

    const mk = (cls, parent = this.viewEl) => { const d = document.createElement('div'); d.className = cls; parent.append(d); return d; };
    this.paneBox = mk('panes');
    this.panes = ['tl', 'tr', 'bl', 'br'].map((id) => {
      const el = mk(`pane pane-${id}`, this.paneBox);
      const content = mk('pane-content', el);
      return {
        id, el, content,
        grid: mk('gl', content), cells: mk('cells', content), objects: mk('objects', content), overlay: mk('overlay', content),
        scrollX: id === 'tr' || id === 'br', scrollY: id === 'bl' || id === 'br', win: null, ox: 0, oy: 0,
      };
    });
    this.colHead = mk('col-head');
    this.colHeadFrozen = mk('head-clip', this.colHead);
    this.colHeadScroll = mk('head-clip', this.colHead);
    this.rowHead = mk('row-head');
    this.rowHeadFrozen = mk('head-clip', this.rowHead);
    this.rowHeadScroll = mk('head-clip', this.rowHead);
    this.corner = mk('corner');
    this.corner.title = '모두 선택';
    this.freezeV = mk('freeze-line v');
    this.freezeH = mk('freeze-line h');

    this.scroll.addEventListener('scroll', () => this.onScroll());
    this.viewEl.addEventListener('wheel', (e) => this.onWheel(e), { passive: false });
    this.bindTouch();
    new ResizeObserver(() => this.layout()).observe(this.wrap);
  }

  // ───────────── 좌표 ─────────────
  refreshAxes() {
    const { wb, si } = this.host.state();
    const s = wb.sheets[si];
    this.cols = new Axis(DEFAULT_COL_WIDTH, s.colWidths, [s.hiddenCols], MAX_COLS);
    this.rows = new Axis(DEFAULT_ROW_HEIGHT, s.rowHeights, [s.hiddenRows, s.filter?.hidden, ...(s.tables ?? []).map((t) => t.filter?.hidden)], MAX_ROWS);
    this.fr = Math.min(s.freeze?.rows || 0, MAX_ROWS - 1);
    this.fc = Math.min(s.freeze?.cols || 0, MAX_COLS - 1);
    this.frozenW = this.cols.pos(this.fc);
    this.frozenH = this.rows.pos(this.fr);
  }

  get viewW() { return this.scroll.clientWidth / this.z; }
  get viewH() { return this.scroll.clientHeight / this.z; }

  sheetRect(rg) {
    const c2 = Math.min(rg.c2, MAX_COLS - 1);
    const r2 = Math.min(rg.r2, MAX_ROWS - 1);
    const x = this.cols.pos(rg.c1);
    const y = this.rows.pos(rg.r1);
    return { x, y, w: this.cols.pos(c2 + 1) - x, h: this.rows.pos(r2 + 1) - y };
  }

  /** 범위의 화면 좌표 (보기 영역 기준, 확대 전 단위) */
  screenRect(rg) {
    const r = this.sheetRect(rg);
    const fx = rg.c1 < this.fc;
    const fy = rg.r1 < this.fr;
    return { x: this.hw + (fx ? r.x : r.x - this.sx), y: this.hh + (fy ? r.y : r.y - this.sy), w: r.w, h: r.h };
  }

  /** 범위의 브라우저 좌표 */
  clientRect(rg) {
    const s = this.screenRect(rg);
    const v = this.viewEl.getBoundingClientRect();
    return { left: v.left + s.x * this.z, top: v.top + s.y * this.z, right: v.left + (s.x + s.w) * this.z, bottom: v.top + (s.y + s.h) * this.z, width: s.w * this.z, height: s.h * this.z };
  }

  /** 브라우저 좌표 → 영역/셀 */
  hitTest(clientX, clientY, clampToCells = false) {
    const v = this.viewEl.getBoundingClientRect();
    let x = (clientX - v.left) / this.z;
    let y = (clientY - v.top) / this.z;
    const out = { dx: 0, dy: 0 };
    if (clampToCells) {
      const minX = this.hw + 1;
      const minY = this.hh + 1;
      const maxX = this.viewW - 2;
      const maxY = this.viewH - 2;
      if (x < minX) { out.dx = -1; x = minX; }
      if (y < minY) { out.dy = -1; y = minY; }
      if (x > maxX) { out.dx = 1; x = maxX; }
      if (y > maxY) { out.dy = 1; y = maxY; }
    }
    const zone = y < this.hh && x < this.hw ? 'corner' : y < this.hh ? 'colHeader' : x < this.hw ? 'rowHeader' : 'cell';
    const sheetX = x - this.hw < this.frozenW ? Math.max(0, x - this.hw) : x - this.hw + this.sx;
    const sheetY = y - this.hh < this.frozenH ? Math.max(0, y - this.hh) : y - this.hh + this.sy;
    const c = this.cols.indexAt(sheetX);
    const r = this.rows.indexAt(sheetY);
    let edgeCol = null;
    let edgeRow = null;
    if (zone === 'colHeader') {
      if (this.cols.pos(c + 1) - sheetX < 5) edgeCol = c;
      else if (sheetX - this.cols.pos(c) < 4 && c > 0) edgeCol = this.cols.nextVisible(c - 1, -1);
    }
    if (zone === 'rowHeader') {
      if (this.rows.pos(r + 1) - sheetY < 4) edgeRow = r;
      else if (sheetY - this.rows.pos(r) < 3 && r > 0) edgeRow = this.rows.nextVisible(r - 1, -1);
    }
    return { zone, r, c, x, y, sheetX, sheetY, edgeCol, edgeRow, ...out };
  }

  // ───────────── 스크롤 ─────────────
  maxScroll() {
    const totalW = this.cols.pos(this.extC);
    const totalH = this.rows.pos(this.extR);
    return {
      x: Math.max(0, this.hw + totalW - this.viewW),
      y: Math.max(0, this.hh + totalH - this.viewH),
      pxW: Math.min(MAX_PX, (this.hw + totalW) * this.z),
      pxH: Math.min(MAX_PX, (this.hh + totalH) * this.z),
    };
  }

  updateSizer() {
    const m = this.maxScroll();
    this.sizer.style.width = `${Math.ceil(m.pxW)}px`;
    this.sizer.style.height = `${Math.ceil(m.pxH)}px`;
    return m;
  }

  readScroll() {
    const m = this.maxScroll();
    const maxPxX = Math.max(1, this.scroll.scrollWidth - this.scroll.clientWidth);
    const maxPxY = Math.max(1, this.scroll.scrollHeight - this.scroll.clientHeight);
    this.sx = m.pxW >= MAX_PX ? (this.scroll.scrollLeft / maxPxX) * m.x : this.scroll.scrollLeft / this.z;
    this.sy = m.pxH >= MAX_PX ? (this.scroll.scrollTop / maxPxY) * m.y : this.scroll.scrollTop / this.z;
    this.sx = Math.max(0, Math.min(this.sx, m.x));
    this.sy = Math.max(0, Math.min(this.sy, m.y));
  }

  setScroll(sx, sy) {
    this.ensureExtentFor(this.rows.indexAt(sy + this.frozenH + this.viewH), this.cols.indexAt(sx + this.frozenW + this.viewW));
    const m = this.updateSizer();
    sx = Math.max(0, Math.min(sx, m.x));
    sy = Math.max(0, Math.min(sy, m.y));
    const maxPxX = Math.max(1, this.scroll.scrollWidth - this.scroll.clientWidth);
    const maxPxY = Math.max(1, this.scroll.scrollHeight - this.scroll.clientHeight);
    this.scroll.scrollLeft = m.pxW >= MAX_PX ? (sx / Math.max(1, m.x)) * maxPxX : sx * this.z;
    this.scroll.scrollTop = m.pxH >= MAX_PX ? (sy / Math.max(1, m.y)) * maxPxY : sy * this.z;
    this.sx = sx;
    this.sy = sy;
    this.update();
  }

  scrollBy(dx, dy) { this.setScroll(this.sx + dx, this.sy + dy); }

  /** 시트 확장 범위 (스크롤바가 닿는 곳) */
  ensureExtentFor(r, c) {
    const { wb, si } = this.host.state();
    const used = wb.extent(si);
    const needR = Math.min(MAX_ROWS, Math.max(100, used.rows + 50, r + 50));
    const needC = Math.min(MAX_COLS, Math.max(26, used.cols + 10, c + 10));
    let changed = false;
    if (needR > this.extR) { this.extR = needR; changed = true; }
    if (needC > this.extC) { this.extC = needC; changed = true; }
    return changed;
  }

  resetExtent() {
    this.extR = 100;
    this.extC = 26;
  }

  onScroll() {
    this.readScroll();
    // 끝 가까이 가면 영역 확장 (무한 스크롤)
    const m = this.maxScroll();
    let grown = false;
    if (this.sy > m.y - this.viewH && this.extR < MAX_ROWS) { this.extR = Math.min(MAX_ROWS, this.extR + 500); grown = true; }
    if (this.sx > m.x - this.viewW && this.extC < MAX_COLS) { this.extC = Math.min(MAX_COLS, this.extC + 30); grown = true; }
    if (grown) this.updateSizer();
    this.update();
  }

  onWheel(e) {
    // 슬라이서 항목 목록은 그 안에서 스크롤
    const list = e.target.closest?.('.sl-items');
    if (list && !e.ctrlKey && list.scrollHeight > list.clientHeight) {
      const atTop = list.scrollTop <= 0 && e.deltaY < 0;
      const atEnd = list.scrollTop + list.clientHeight >= list.scrollHeight - 1 && e.deltaY > 0;
      if (!atTop && !atEnd) return;
    }
    if (e.ctrlKey) {
      e.preventDefault();
      this.host.onZoomWheel?.(e.deltaY < 0 ? 10 : -10);
      return;
    }
    e.preventDefault();
    const unit = e.deltaMode === 1 ? 20 : e.deltaMode === 2 ? this.viewH : 1;
    let dx = e.deltaX * unit;
    let dy = e.deltaY * unit;
    if (e.shiftKey && !dx) { dx = dy; dy = 0; }
    if (e.deltaMode === 1) dy = Math.sign(dy) * Math.max(Math.abs(dy), DEFAULT_ROW_HEIGHT * 3);
    this.scroll.scrollLeft += dx * this.z;
    this.scroll.scrollTop += dy * this.z;
  }

  bindTouch() {
    let last = null;
    this.viewEl.addEventListener('touchstart', (e) => {
      if (e.touches.length === 1) last = { x: e.touches[0].clientX, y: e.touches[0].clientY };
    }, { passive: true });
    this.viewEl.addEventListener('touchmove', (e) => {
      if (!last || e.touches.length !== 1 || this.host.isDragging?.()) return;
      const t = e.touches[0];
      this.scroll.scrollLeft -= t.clientX - last.x;
      this.scroll.scrollTop -= t.clientY - last.y;
      last = { x: t.clientX, y: t.clientY };
      e.preventDefault();
    }, { passive: false });
    this.viewEl.addEventListener('touchend', () => { last = null; });
  }

  setZoom(pct) {
    const keep = { r: this.rows.indexAt(this.sy + this.frozenH), c: this.cols.indexAt(this.sx + this.frozenW) };
    this.z = pct / 100;
    this.viewEl.style.zoom = this.z;
    this.layout(false);
    this.setScroll(this.cols.pos(keep.c) - this.frozenW, this.rows.pos(keep.r) - this.frozenH);
  }

  /** (r,c)가 보이도록 스크롤 */
  ensureVisible(r, c) {
    let { sx, sy } = this;
    const paneW = this.viewW - this.hw - this.frozenW;
    const paneH = this.viewH - this.hh - this.frozenH;
    if (r >= this.fr) {
      const top = this.rows.pos(r) - this.frozenH;
      const bottom = this.rows.pos(r + 1) - this.frozenH;
      if (top < sy) sy = top;
      else if (bottom > sy + paneH) sy = Math.min(top, bottom - paneH);
    }
    if (c >= this.fc) {
      const left = this.cols.pos(c) - this.frozenW;
      const right = this.cols.pos(c + 1) - this.frozenW;
      if (left < sx) sx = left;
      else if (right > sx + paneW) sx = Math.min(left, right - paneW);
    }
    if (sx !== this.sx || sy !== this.sy) {
      this.ensureExtentFor(r + 50, c + 10);
      this.setScroll(sx, sy);
    }
  }

  /** 한 화면에 보이는 행 수 */
  pageRows() {
    return Math.max(1, Math.floor((this.viewH - this.hh - this.frozenH) / DEFAULT_ROW_HEIGHT) - 1);
  }

  firstVisibleRow() { return this.rows.indexAt(this.sy + this.frozenH); }

  // ───────────── 배치 ─────────────
  layout(render = true) {
    if (!this.host.state().wb) return;
    this.viewEl.style.width = `${this.viewW}px`;
    this.viewEl.style.height = `${this.viewH}px`;
    this.refreshAxes();
    this.updateSizer();
    this.readScroll();
    if (render) this.renderAll();
  }

  computeHeaderSize() {
    const { showHeaders } = this.host.state();
    if (!showHeaders) { this.hw = 0; this.hh = 0; return; }
    this.hh = HEAD_H;
    const bottom = this.rows.indexAt(this.sy + this.frozenH + this.viewH);
    this.hw = Math.max(34, String(bottom + 1).length * 8 + 12);
  }

  paneRects() {
    const { hw, hh, frozenW, frozenH } = this;
    const W = this.viewW;
    const H = this.viewH;
    const fw = Math.min(frozenW, Math.max(0, W - hw));
    const fh = Math.min(frozenH, Math.max(0, H - hh));
    return {
      tl: { x: hw, y: hh, w: fw, h: fh },
      tr: { x: hw + fw, y: hh, w: Math.max(0, W - hw - fw), h: fh },
      bl: { x: hw, y: hh + fh, w: fw, h: Math.max(0, H - hh - fh) },
      br: { x: hw + fw, y: hh + fh, w: Math.max(0, W - hw - fw), h: Math.max(0, H - hh - fh) },
    };
  }

  /** 창에 보이는 시트 영역 */
  visibleRange(p, rect) {
    const x0 = p.scrollX ? this.frozenW + this.sx : 0;
    const y0 = p.scrollY ? this.frozenH + this.sy : 0;
    let c1 = this.cols.indexAt(x0);
    let c2 = this.cols.indexAt(x0 + Math.max(0, rect.w - 0.5));
    let r1 = this.rows.indexAt(y0);
    let r2 = this.rows.indexAt(y0 + Math.max(0, rect.h - 0.5));
    if (p.scrollX) c1 = Math.max(c1, this.fc); else { c1 = 0; c2 = Math.max(0, this.fc - 1); }
    if (p.scrollY) r1 = Math.max(r1, this.fr); else { r1 = 0; r2 = Math.max(0, this.fr - 1); }
    return { r1, r2: Math.max(r1, r2), c1, c2: Math.max(c1, c2), x0, y0 };
  }

  /** 스크롤 후 창 이동 · 필요하면 다시 그림 */
  update(force = false) {
    const prevHw = this.hw;
    this.computeHeaderSize();
    if (this.hw !== prevHw) force = true;
    const rects = this.paneRects();
    for (const p of this.panes) {
      const rect = rects[p.id];
      const visible = rect.w > 0 && rect.h > 0;
      p.el.style.display = visible ? 'block' : 'none';
      if (!visible) { p.win = null; continue; }
      Object.assign(p.el.style, { left: `${rect.x}px`, top: `${rect.y}px`, width: `${rect.w}px`, height: `${rect.h}px` });
      const need = this.visibleRange(p, rect);
      const w = p.win;
      if (force || !w || need.r1 < w.r1 || need.r2 > w.r2 || need.c1 < w.c1 || need.c2 > w.c2) {
        p.win = {
          r1: p.scrollY ? Math.max(this.fr, need.r1 - OVER_R) : need.r1,
          r2: p.scrollY ? Math.min(MAX_ROWS - 1, need.r2 + OVER_R) : need.r2,
          c1: p.scrollX ? Math.max(this.fc, need.c1 - OVER_C) : need.c1,
          c2: p.scrollX ? Math.min(MAX_COLS - 1, need.c2 + OVER_C) : need.c2,
        };
        this.renderPane(p);
        this.renderPaneOverlay(p);
      }
      p.content.style.transform = `translate(${p.ox - need.x0}px, ${p.oy - need.y0}px)`;
    }
    this.renderHeaders(rects);
    this.freezeV.style.display = this.fc ? 'block' : 'none';
    this.freezeH.style.display = this.fr ? 'block' : 'none';
    this.freezeV.style.left = `${this.hw + this.frozenW - 1}px`;
    this.freezeH.style.top = `${this.hh + this.frozenH - 1}px`;
    this.host.onViewScroll?.();
  }

  renderAll() {
    const { wb, si } = this.host.state();
    this.cond = wb.sheets[si].cond.length ? prepareCond(wb, si) : null;
    this.update(true);
  }

  // ───────────── 셀 그리기 ─────────────
  renderPane(p) {
    const st = this.host.state();
    const { wb, si } = st;
    const sheet = wb.sheets[si];
    const { r1, r2, c1, c2 } = p.win;
    const cols = this.cols;
    const rows = this.rows;
    p.ox = cols.pos(c1);
    p.oy = rows.pos(r1);
    const W = cols.pos(c2 + 1) - p.ox;
    const H = rows.pos(r2 + 1) - p.oy;
    const visRows = [];
    for (let r = r1; r <= r2; r++) if (rows.size(r)) visRows.push(r);
    const visCols = [];
    for (let c = c1; c <= c2; c++) if (cols.size(c)) visCols.push(c);

    // 눈금선
    const g = [];
    if (st.showGrid) {
      for (const c of visCols) g.push(`<i class="gv" style="left:${cols.pos(c + 1) - 1 - p.ox}px;height:${H}px"></i>`);
      for (const r of visRows) g.push(`<i class="gh" style="top:${rows.pos(r + 1) - 1 - p.oy}px;width:${W}px"></i>`);
    }
    p.grid.innerHTML = g.join('');

    const merges = sheet.merges.filter((m) => m.r1 <= r2 && m.r2 >= r1 && m.c1 <= c2 && m.c2 >= c1);
    const inMerge = (r, c) => merges.some((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2);
    const html = [];
    const hasLine = !!(sheet.allStyle || Object.keys(sheet.colStyles).length || Object.keys(sheet.rowStyles).length);
    const tables = (sheet.tables ?? []).filter((t) => t.r1 <= r2 && t.r2 >= r1 && t.c1 <= c2 && t.c2 >= c1 && styleByName(t.style));
    const inTable = (r, c) => tables.some((t) => r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2);
    const emptyRules = (this.cond ?? []).map((pr) => pr.rule).filter((rl) => EMPTY_MATCH_TYPES.has(rl.type) && ruleRanges(rl).some((g) => g.r1 <= r2 && g.r2 >= r1 && g.c1 <= c2 && g.c2 >= c1));
    const emptyCond = (r, c) => emptyRules.some((rl) => inRule(rl, r, c));
    const spills = wb.spillsOf(si).filter((sp) => sp.r <= r2 && sp.r + sp.h - 1 >= r1 && sp.c <= c2 && sp.c + sp.w - 1 >= c1);
    const inSpill = (r, c) => spills.some((sp) => r >= sp.r && r < sp.r + sp.h && c >= sp.c && c < sp.c + sp.w);
    for (const r of visRows) {
      // 창 왼쪽 밖에서 넘쳐 들어오는 텍스트
      if (c1 > 0) {
        for (let k = c1 - 1; k >= Math.max(0, c1 - 40); k--) {
          const cell = wb.getCell(si, r, k);
          if (!cell?.raw) continue;
          if (!inMerge(r, k) && cols.size(k)) html.push(this.cellHtml(r, k, p, sheet, merges));
          break;
        }
      }
      for (const c of visCols) {
        if (inMerge(r, c)) continue;
        if (!hasLine && !sheet.cells.has(`${r},${c}`) && !(sheet.blocks?.length && wb.blockAt(si, r, c)) && !inTable(r, c) && !emptyCond(r, c) && !inSpill(r, c)) continue;
        html.push(this.cellHtml(r, c, p, sheet, merges));
      }
    }
    for (const m of merges) html.push(this.cellHtml(m.r1, m.c1, p, sheet, merges, m));

    // 필터 단추 (시트 필터 + 표마다)
    const targets = [
      ...(sheet.filter ? [['', sheet.filter]] : []),
      ...(sheet.tables ?? []).filter((t) => t.filter && t.header).map((t) => [t.id, tableFilterRange(t)]),
    ];
    for (const [tid, f] of targets) {
      if (!(f.r1 >= r1 && f.r1 <= r2)) continue;
      for (let c = Math.max(f.c1, c1); c <= Math.min(f.c2, c2); c++) {
        if (!cols.size(c) || !rows.size(f.r1)) continue;
        const active = Array.isArray(f.criteria?.[c]);
        const sort = f.sort?.col === c ? (f.sort.asc ? ' asc' : ' desc') : '';
        html.push(`<div class="fbtn${active ? ' on' : ''}${sort}" data-c="${c}" data-t="${esc(tid)}" title="${active ? '필터 적용됨' : '필터'}" style="left:${cols.pos(c + 1) - 18 - p.ox}px;top:${rows.pos(f.r1 + 1) - 18 - p.oy}px"></div>`);
      }
    }
    // 피벗 테이블 필터 단추 (행 레이블 · 열 레이블 · 보고서 필터)
    [sheet.pivot, ...(sheet.pivotsExtra ?? [])].filter(Boolean).forEach((pd, pi) => {
      if (!pd.buttons) return;
      const filtered = (field) => !!(field && (pd.filters?.[field] || pd.fieldFilters?.[field]));
      for (const b of pd.buttons) {
        if (b.r < r1 || b.r > r2 || b.c < c1 || b.c > c2 || !cols.size(b.c) || !rows.size(b.r)) continue;
        const fields = b.field ? [b.field] : b.kind === 'rows' ? pd.rows ?? [] : b.kind === 'cols' ? pd.cols ?? [] : [];
        const on = fields.some(filtered) || (b.kind !== 'page' && fields.some((f) => pd.sort?.[f]));
        html.push(`<div class="fbtn pbtn${on ? ' on' : ''}" data-p="${pi}" data-k="${b.kind}" data-f="${esc(b.field ?? '')}" title="${on ? '필터 적용됨' : '필터'}" style="left:${cols.pos(b.c + 1) - 18 - p.ox}px;top:${rows.pos(b.r + 1) - 18 - p.oy}px"></div>`);
      }
    });
    p.cells.innerHTML = html.join('');
    this.renderObjects(p);
  }

  cellHtml(r, c, p, sheet, merges, merge = null) {
    const st = this.host.state();
    const { wb, si } = st;
    const cell = wb.getCell(si, r, c);
    let style = wb.styleAt(si, r, c);
    const tbl = tableAt(sheet, r, c);
    if (tbl) {
      // 표 서식은 셀에 직접 지정한 서식 아래에 깔림
      const ts = tableCellStyle(tbl, r, c);
      if (ts) {
        const own = {};
        for (const [k, val] of Object.entries(style)) if (val !== undefined) own[k] = val;
        style = { ...ts, ...own };
      }
    }
    const v = wb.getValue(si, r, c);
    const x = this.cols.pos(c);
    const y = this.rows.pos(r);
    const w = merge ? this.cols.pos(merge.c2 + 1) - x : this.cols.size(c);
    const h = merge ? this.rows.pos(merge.r2 + 1) - y : this.rows.size(r);
    if (!w || !h) return '';
    let bar = null;
    let icon = null;
    let hideValue = false;
    if (this.cond) {
      const cf = condFormatAt(this.cond, wb, si, r, c, v);
      if (cf.style) style = { ...style, ...cf.style };
      ({ bar, icon, hideValue } = cf);
    }
    let text;
    let align;
    let fmtColor = null;
    let image = null;
    if (st.showFormulas && cell?.formula) { text = cell.raw; align = 'left'; } else ({ text, align, color: fmtColor, image } = formatValue(v, style));
    const eff = style.align || align;
    const css = [`left:${x - 1 - p.ox}px`, `top:${y - 1 - p.oy}px`, `width:${w + 1}px`, `height:${h + 1}px`];
    if (style.bold) css.push('font-weight:700');
    if (style.italic) css.push('font-style:italic');
    if (style.underline || style.strike) css.push(`text-decoration:${style.underline ? 'underline ' : ''}${style.strike ? 'line-through' : ''}`);
    if (fmtColor || style.color) css.push(`color:${fmtColor || style.color}`);
    if (style.font) css.push(`font-family:${fontStack(style.font)}`);
    if (style.size) css.push(`font-size:${style.size}pt`);
    if (eff !== 'left') css.push(`justify-content:${eff === 'center' ? 'center' : 'flex-end'};text-align:${eff}`);
    if (style.valign === 'top') css.push('align-items:flex-start');
    else if (style.valign === 'middle') css.push('align-items:center');
    if (style.indent) css.push(`padding-${eff === 'right' ? 'right' : 'left'}:${3 + style.indent * 9}px`);
    const bg = style.fill || (merge ? '#fff' : null);
    if (bar) {
      // 그라데이션(엑셀 기본) 또는 단색, 음수 막대는 오른쪽에서 왼쪽으로
      const img = bar.gradient
        ? `linear-gradient(90deg, ${bar.color}, ${bar.color}33)`
        : `linear-gradient(${bar.color}, ${bar.color})`;
      css.push(`background:${img} no-repeat ${bar.neg ? '100%' : '0'} 50% / ${bar.pct}% 72%${bg ? `, ${bg}` : ''};background-clip:padding-box`);
    }
    else if (bg) css.push(`background-color:${bg}`);
    if (style.bt) css.push('border-top-color:#000');
    if (style.bb) css.push('border-bottom-color:#000');
    if (style.bl) css.push('border-left-color:#000');
    if (style.br) css.push('border-right-color:#000');
    const cls = [];
    if (style.wrap) cls.push('wrap');
    else if (text && eff === 'left' && typeof v !== 'number' && !merge) {
      const next = wb.getCell(si, r, c + 1);
      if (!next?.raw && !merges.some((m) => r >= m.r1 && r <= m.r2 && c + 1 >= m.c1 && c + 1 <= m.c2)) cls.push('ovf');
    }
    if (cell?.comment) cls.push('cm');
    // 아이콘 집합: 아이콘은 왼쪽 끝에 고정하고 글자는 남은 너비 안에 (엑셀과 같음)
    const ICON_W = 18;
    if (icon) {
      css.push(`padding-left:${(style.indent && eff !== 'right' ? 3 + style.indent * 9 : 3) + ICON_W}px`);
      const i = cls.indexOf('ovf');
      if (i >= 0) cls.splice(i, 1);
    }
    const room = w - 6 - (icon ? ICON_W : 0);
    if (typeof v === 'number' && text && !style.wrap && !st.showFormulas && measureText(text, style) > room) {
      text = fitNumber(v, room, style);
    }
    const comment = cell?.comment ? ` data-cm="${esc(cell.comment)}"` : '';
    const iconHtml = icon ? `<i class="cf-icon">${ICON_SVG[icon] ?? ''}</i>` : '';
    if (icon) cls.push('has-icon');
    if (image) {
      // 셀 안 그림: 0 셀에 맞춤(비율 유지), 1 셀 채우기, 2 원래 크기, 3 높이 · 너비 지정
      const fit = image.sizing === 1 ? 'fill' : image.sizing === 2 ? 'none' : 'contain';
      const size = image.sizing === 3 ? `width:${image.w ? `${image.w}px` : 'auto'};height:${image.h ? `${image.h}px` : 'auto'};` : 'width:100%;height:100%;';
      const img = `<img class="cimg" src="${esc(image.src)}" alt="${esc(image.alt ?? '')}" title="${esc(image.alt ?? '')}" draggable="false" loading="lazy" style="${size}object-fit:${fit}">`;
      const i = cls.indexOf('ovf');
      if (i >= 0) cls.splice(i, 1);
      return `<div class="c cimg-cell${cls.length ? ` ${cls.join(' ')}` : ''}" data-r="${r}" data-c="${c}" style="${css.join(';')}"${comment}>${img}</div>`;
    }
    return `<div class="c${cls.length ? ` ${cls.join(' ')}` : ''}" data-r="${r}" data-c="${c}" style="${css.join(';')}"${comment}>${iconHtml}<span>${hideValue ? '' : esc(text)}</span></div>`;
  }

  /** 그림 개체: 차트 · 그림 · 도형 */
  renderObjects(p) {
    const st = this.host.state();
    const { wb, si } = st;
    const sheet = wb.sheets[si];
    const images = sheet.images ?? [];
    const shapes = sheet.shapes ?? [];
    const slicers = sheet.slicers ?? [];
    if (!sheet.charts.length && !images.length && !shapes.length && !slicers.length) { p.objects.innerHTML = ''; return; }
    const winX1 = p.scrollX ? this.frozenW : 0;
    const winY1 = p.scrollY ? this.frozenH : 0;
    const winX2 = p.scrollX ? Infinity : this.frozenW;
    const winY2 = p.scrollY ? Infinity : this.frozenH;
    const html = [];
    const handles = '<i class="ch-h nw"></i><i class="ch-h ne"></i><i class="ch-h sw"></i><i class="ch-h se"></i>';
    const box = (o, cls, inner, extraCss = '') => {
      const h = Math.max(o.h, cls.includes('line') ? 1 : 0);
      if (o.x + o.w < winX1 || o.x > winX2 || o.y + h < winY1 || o.y > winY2) return;
      const selected = st.chartSel === o.id;
      html.push(`<div class="obj ${cls}${selected ? ' sel' : ''}" data-id="${esc(o.id)}" style="left:${o.x - p.ox}px;top:${o.y - p.oy}px;width:${o.w}px;height:${h}px;${extraCss}">${inner}${selected ? handles : ''}</div>`);
    };
    // 엑셀처럼 그림 → 도형 → 차트 순서가 아니라 저장된 순서(z)대로 겹침
    const all = [
      ...sheet.charts.map((o) => ['charts', o]), ...images.map((o) => ['images', o]), ...shapes.map((o) => ['shapes', o]),
      ...slicers.map((o) => ['slicers', o]),
    ].sort((a, b) => (a[1].z ?? 0) - (b[1].z ?? 0));
    for (const [prop, o] of all) {
      if (prop === 'charts') box(o, 'chart', this.chartSvg(o));
      else if (prop === 'slicers') box(o, 'slicer', this.slicerHtml(o), slicerCssVars(o));
      else if (prop === 'images') box(o, 'pic', `<img src="${esc(o.src)}" alt="${esc(o.name ?? '')}" draggable="false">`);
      else {
        const text = o.text && o.kind !== 'line'
          ? `<div class="sh-text" style="justify-content:${o.kind === 'textbox' ? 'flex-start' : 'center'};text-align:${o.align ?? 'center'};color:${esc(o.color ?? '#000')};font-size:${o.size ?? 11}pt;${o.bold ? 'font-weight:700;' : ''}">${esc(o.text)}</div>`
          : '';
        box(o, `shape ${o.kind === 'line' ? 'line' : ''}`, shapeSvg(o) + text);
      }
    }
    p.objects.innerHTML = html.join('');
  }

  /** 슬라이서: 머리글(캡션·다중 선택·필터 지우기) + 항목 단추 */
  slicerHtml(sl) {
    const m = this.host.slicerModel?.(sl) ?? { items: [], filtered: false };
    const items = m.broken
      ? `<div class="sl-broken">${esc(m.broken)}</div>`
      : m.items.map((it) => `<button type="button" class="sl-item${it.selected ? ' on' : ''}${it.hasData ? '' : ' nodata'}" data-k="${esc(it.key)}" title="${esc(it.text)}">${esc(it.text)}</button>`).join('');
    const head = sl.showHeader === false ? '' : `<div class="sl-head"><span class="sl-cap">${esc(sl.caption ?? '')}</span>`
      + `<button type="button" class="sl-multi${sl.multi ? ' on' : ''}" title="다중 선택 (Alt+S)">☰</button>`
      + `<button type="button" class="sl-clear${m.filtered ? '' : ' off'}" title="필터 지우기 (Alt+C)">✕</button></div>`;
    return `${head}<div class="sl-items" style="grid-template-columns:repeat(${Math.max(1, sl.columns ?? 1)}, minmax(0, 1fr))">${items}</div>`;
  }

  chartSvg(ch) {
    const { wb, si } = this.host.state();
    return renderChartSvg(ch, chartModelData(wb, si, ch));
  }

  /** 선택 영역이 바뀌었을 때 (셀은 그대로) */
  renderSelection() {
    this.renderOverlays();
    this.renderHeaders(this.paneRects());
  }

  renderObjectsAll() { for (const p of this.panes) if (p.win) this.renderObjects(p); }

  // ───────────── 선택 영역 등 겹쳐 그리기 ─────────────
  renderOverlays() { for (const p of this.panes) if (p.win) this.renderPaneOverlay(p); }

  renderPaneOverlay(p) {
    const st = this.host.state();
    const { wb, si, sel, active, selKind } = st;
    const win = p.win;
    const bx1 = this.cols.pos(win.c1) - 3;
    const by1 = this.rows.pos(win.r1) - 3;
    const bx2 = this.cols.pos(win.c2 + 1) + 3;
    const by2 = this.rows.pos(win.r2 + 1) + 3;
    const clip = (r) => {
      const x1 = Math.max(r.x, bx1);
      const y1 = Math.max(r.y, by1);
      const x2 = Math.min(r.x + r.w, bx2);
      const y2 = Math.min(r.y + r.h, by2);
      return x2 <= x1 || y2 <= y1 ? null : { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
    };
    const box = (cls, r, extra = '') => {
      const c = clip(r);
      return c ? `<div class="${cls}" style="left:${c.x - p.ox}px;top:${c.y - p.oy}px;width:${c.w}px;height:${c.h}px;${extra}"></div>` : '';
    };
    const html = [];
    const selRect = this.sheetRect(sel);
    const am = wb.mergeAt(si, active.r, active.c) ?? { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
    const single = sel.r1 === am.r1 && sel.c1 === am.c1 && sel.r2 === am.r2 && sel.c2 === am.c2;
    if (!single) {
      const a = this.sheetRect(am);
      const t = selRect;
      html.push(box('tint', { x: t.x, y: t.y, w: t.w, h: a.y - t.y }));
      html.push(box('tint', { x: t.x, y: a.y + a.h, w: t.w, h: t.y + t.h - a.y - a.h }));
      html.push(box('tint', { x: t.x, y: a.y, w: a.x - t.x, h: a.h }));
      html.push(box('tint', { x: a.x + a.w, y: a.y, w: t.x + t.w - a.x - a.w, h: a.h }));
    }
    html.push(box('sel-border', { x: selRect.x - 1, y: selRect.y - 1, w: selRect.w + 1, h: selRect.h + 1 }));
    if (!st.editing && selKind === 'cells' && !st.chartSel) {
      const hx = selRect.x + selRect.w - 4;
      const hy = selRect.y + selRect.h - 4;
      if (hx >= bx1 && hx <= bx2 && hy >= by1 && hy <= by2) html.push(`<div class="fill-handle" style="left:${hx - p.ox}px;top:${hy - p.oy}px"></div>`);
    }
    // 데이터 유효성 검사: 목록 단추 · 잘못된 데이터 동그라미
    if (!st.editing && !st.chartSel) {
      const rule = validationAt(wb.sheets[si], active.r, active.c);
      const tt = tableAt(wb.sheets[si], active.r, active.c);
      if (tt?.totals && active.r === tt.r2) {
        const a = this.sheetRect(am);
        const x = a.x + a.w + 1;
        const y = a.y + a.h - 18;
        if (x >= bx1 && x <= bx2 && y + 18 >= by1 && y <= by2) html.push(`<div class="dv-btn" data-tt="1" title="요약 함수 선택" style="left:${x - p.ox}px;top:${y - p.oy}px"></div>`);
      } else if (rule?.type === 'list' && rule.showDropdown !== false) {
        const a = this.sheetRect(am);
        const x = a.x + a.w + 1;
        const y = a.y + a.h - 18;
        if (x >= bx1 && x <= bx2 && y + 18 >= by1 && y <= by2) html.push(`<div class="dv-btn" title="목록에서 선택" style="left:${x - p.ox}px;top:${y - p.oy}px"></div>`);
      }
    }
    // 동적 배열 분산 범위 (파란 점선)
    if (!st.editing && !st.chartSel) {
      const own = wb.getCell(si, active.r, active.c);
      const anc = own?.formula ? { r: active.r, c: active.c } : !own?.raw ? wb.spillAnchorOf(si, active.r, active.c) : null;
      const sp = anc ? wb.spillRange(si, anc.r, anc.c) : null;
      if (sp) {
        const r = this.sheetRect(sp);
        html.push(box('spill-border', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }));
      }
    }
    for (const cc of st.circles ?? []) {
      const r = this.sheetRect({ r1: cc.r, c1: cc.c, r2: cc.r, c2: cc.c });
      html.push(box('dv-circle', { x: r.x - 4, y: r.y - 3, w: r.w + 8, h: r.h + 6 }));
    }
    if (st.clip && st.clip.si === si) {
      const r = this.sheetRect(st.clip);
      html.push(box('marquee', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }));
    }
    if (st.fillPreview) {
      const r = this.sheetRect(st.fillPreview);
      html.push(box('fill-preview', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }));
    }
    for (const ref of st.refs ?? []) {
      const r = this.sheetRect(ref.rg);
      html.push(box('ref-box', { x: r.x - 1, y: r.y - 1, w: r.w + 1, h: r.h + 1 }, `border-color:${ref.color};background:${ref.color}14`));
    }
    p.overlay.innerHTML = html.join('');
  }

  // ───────────── 머리글 ─────────────
  renderHeaders(rects) {
    const st = this.host.state();
    const { sel, selKind } = st;
    const show = st.showHeaders;
    this.colHead.style.display = show ? 'block' : 'none';
    this.rowHead.style.display = show ? 'block' : 'none';
    this.corner.style.display = show ? 'block' : 'none';
    if (!show) return;
    const { hw, hh } = this;
    this.corner.style.width = `${hw}px`;
    this.corner.style.height = `${hh}px`;
    Object.assign(this.colHead.style, { left: '0px', top: '0px', width: `${this.viewW}px`, height: `${hh}px` });
    Object.assign(this.rowHead.style, { left: '0px', top: '0px', width: `${hw}px`, height: `${this.viewH}px` });
    const colFull = selKind === 'cols' || selKind === 'all';
    const rowFull = selKind === 'rows' || selKind === 'all';
    const colCls = (c) => (c >= sel.c1 && c <= sel.c2 ? (colFull ? ' full' : ' hl') : '');
    const rowCls = (r) => (r >= sel.r1 && r <= sel.r2 ? (rowFull ? ' full' : ' hl') : '');

    const colPart = (clipEl, rect, c1, c2, offset) => {
      Object.assign(clipEl.style, { left: `${rect.x}px`, top: '0px', width: `${rect.w}px`, height: `${hh}px`, display: rect.w > 0 ? 'block' : 'none' });
      const out = [];
      for (let c = c1; c <= c2; c++) {
        const w = this.cols.size(c);
        if (!w) continue;
        out.push(`<div class="hc${colCls(c)}" style="left:${this.cols.pos(c) - offset}px;width:${w}px">${colToName(c)}</div>`);
      }
      clipEl.innerHTML = out.join('');
    };
    const rowPart = (clipEl, rect, r1, r2, offset) => {
      Object.assign(clipEl.style, { left: '0px', top: `${rect.y}px`, width: `${hw}px`, height: `${rect.h}px`, display: rect.h > 0 ? 'block' : 'none' });
      const out = [];
      for (let r = r1; r <= r2; r++) {
        const h = this.rows.size(r);
        if (!h) continue;
        out.push(`<div class="hr${rowCls(r)}" style="top:${this.rows.pos(r) - offset}px;height:${h}px;line-height:${h - 1}px">${r + 1}</div>`);
      }
      clipEl.innerHTML = out.join('');
    };
    const tl = rects.tl;
    const br = rects.br;
    colPart(this.colHeadFrozen, { x: tl.x, w: tl.w }, 0, this.fc - 1, 0);
    const vx = this.visibleRange(this.panes[3], br);
    colPart(this.colHeadScroll, { x: br.x, w: br.w }, vx.c1, vx.c2, vx.x0);
    rowPart(this.rowHeadFrozen, { y: tl.y, h: tl.h }, 0, this.fr - 1, 0);
    rowPart(this.rowHeadScroll, { y: br.y, h: br.h }, vx.r1, vx.r2, vx.y0);
  }
}
