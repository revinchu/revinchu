// 시트의 페이지 나누기와 실제 인쇄가 공유하는 DOM 없는 좌표 계산.
import { normPage, paperOf, printScale } from './page.js';

export const PRINT_LAYOUT_LIMITS = { areas: 128, pages: 2000, rows: 20000000, cols: 16384 };
const finite = (v) => Number.isFinite(Number(v));
const clamp = (n, a, b) => Math.max(a, Math.min(b, n));

export function normalizePrintAreas(page, fallbackArea, { maxRows = PRINT_LAYOUT_LIMITS.rows, maxCols = PRINT_LAYOUT_LIMITS.cols } = {}) {
  const input = page?.areas?.length ? page.areas : page?.area ? [page.area] : fallbackArea ? [fallbackArea] : [];
  if (!Array.isArray(input) || input.length > PRINT_LAYOUT_LIMITS.areas) throw new Error('인쇄 영역은 128개 이내로 나누어 지정하세요.');
  const seen = new Set(), result = [];
  for (const area of input) {
    if (!area || !['r1', 'r2', 'c1', 'c2'].every((k) => finite(area[k]))) throw new Error('인쇄 영역의 행·열 범위가 올바르지 않습니다.');
    if (Number(area.r2) < 0 || Number(area.c2) < 0) continue;
    const r1 = clamp(Math.floor(Math.min(+area.r1, +area.r2)), 0, maxRows - 1), c1 = clamp(Math.floor(Math.min(+area.c1, +area.c2)), 0, maxCols - 1);
    let r2 = clamp(Math.floor(Math.max(+area.r1, +area.r2)), 0, maxRows - 1), c2 = clamp(Math.floor(Math.max(+area.c1, +area.c2)), 0, maxCols - 1);
    // 전체 행/열 선택은 실제 데이터·개체의 끝까지만 출력합니다. 작은 명시 영역은 자르지 않습니다.
    if (fallbackArea && r1 === 0 && r2 >= Math.min(1048575, maxRows - 1)) r2 = Math.min(r2, Math.max(0, fallbackArea.r2));
    if (fallbackArea && c1 === 0 && c2 >= maxCols - 1) c2 = Math.min(c2, Math.max(0, fallbackArea.c2));
    const key = `${r1},${c1},${r2},${c2}`;
    if (!seen.has(key)) { seen.add(key); result.push({ r1, c1, r2, c2 }); }
  }
  return result;
}

export function normalizePageBreaks(breaks, max = PRINT_LAYOUT_LIMITS.rows) {
  return [...new Set((Array.isArray(breaks) ? breaks : []).filter((n) => Number.isInteger(n) && n > 0 && n < max))].sort((a, b) => a - b);
}

export function printPageGeometry(page) {
  const pg = normPage(page), paper = paperOf(pg.paper), landscape = pg.orientation === 'landscape';
  const w = (landscape ? paper.h : paper.w) * 96, h = (landscape ? paper.w : paper.h) * 96;
  return { paper: { w, h }, content: { w: w - (pg.margins.left + pg.margins.right) * 96, h: h - (pg.margins.top + pg.margins.bottom) * 96 }, headerHeight: 24, footerHeight: 24 };
}

const titleRange = (range, axis) => Array.isArray(range) && range.length >= 2 && finite(range[0]) && finite(range[1]) ? [clamp(Math.floor(Math.min(range[0], range[1])), 0, axis.max - 1), clamp(Math.floor(Math.max(range[0], range[1])), 0, axis.max - 1)] : null;
const span = (axis, start, end) => end >= start ? Math.max(0, axis.pos(end + 1) - axis.pos(start)) : 0;
const overlap = (axis, a, b, repeat) => repeat ? span(axis, Math.max(a, repeat[0]), Math.min(b, repeat[1])) : 0;

// nextVisible / pos와 이분 탐색을 사용해 백만 행 범위에서도 전체 행 배열을 만들지 않습니다.
function splitAxis(axis, first, last, capacity, repeat, manual, maxBands) {
  const bands = [], repeatSize = repeat ? span(axis, repeat[0], repeat[1]) : 0;
  const available = Math.max(.001, capacity - repeatSize);
  let oversized = repeatSize >= capacity && repeatSize > 0, cursor = first;
  const nextBody = (index) => {
    if (repeat && index >= repeat[0] && index <= repeat[1]) index = repeat[1] + 1;
    if (index > last) return last + 1;
    const next = axis.nextVisible(index, 1);
    if (next < index || next > last || axis.size(next) <= 0) return last + 1;
    if (repeat && next >= repeat[0] && next <= repeat[1]) return nextBody(repeat[1] + 1);
    return next;
  };
  let breakAt = 0;
  while (cursor <= last) {
    const start = nextBody(cursor);
    if (start > last) break;
    while (breakAt < manual.length && manual[breakAt] <= start) breakAt++;
    const stop = Math.min(last, (manual[breakAt] ?? last + 1) - 1);
    let lo = start, hi = stop;
    while (lo < hi) { const mid = Math.ceil((lo + hi) / 2), size = span(axis, start, mid) - overlap(axis, start, mid, repeat); if (size <= available + .001) lo = mid; else hi = mid - 1; }
    oversized ||= axis.size(start) > available + .001;
    const next = lo === stop && stop < last ? stop + 1 : nextBody(lo + 1), end = next > last ? last : next - 1;
    bands.push({ start: cursor, end, bodyStart: start, bodyEnd: lo, manual: bands.length > 0 && manual.includes(cursor) });
    if (bands.length > maxBands) throw new Error(`페이지가 ${maxBands.toLocaleString('ko-KR')}쪽을 초과합니다. 인쇄 영역이나 배율을 조정하세요.`);
    cursor = next;
  }
  // 제목 행/열만 있는 작은 인쇄 영역도 한 번 출력합니다.
  if (!bands.length && span(axis, first, last) > 0) bands.push({ start: first, end: last, bodyStart: last + 1, bodyEnd: last, manual: false });
  return { bands, oversized, repeatSize };
}

/** Axis 객체 두 개를 받아 전체 셀 순회 없이 인쇄 페이지와 경계선을 계산합니다. */
export function computePrintLayout({ page, areas: requestedAreas, fallbackArea, rows, cols, maxPages = PRINT_LAYOUT_LIMITS.pages }) {
  const pg = normPage(page), geometry = printPageGeometry(pg), result = { ...geometry, areas: [], pages: [], breaks: [], scale: pg.scale / 100, oversized: false, manualIgnored: false };
  try {
    if (!(geometry.content.w >= 20 && geometry.content.h >= 20)) throw new Error('용지에 비해 여백이 너무 큽니다. 페이지 설정에서 여백을 줄이세요.');
    if (!rows?.pos || !cols?.pos) throw new Error('인쇄 행·열 크기를 확인할 수 없습니다.');
    maxPages = clamp(Math.floor(Number(maxPages) || PRINT_LAYOUT_LIMITS.pages), 1, PRINT_LAYOUT_LIMITS.pages);
    const areas = normalizePrintAreas(requestedAreas ? { areas: requestedAreas } : pg, fallbackArea, { maxRows: rows.max, maxCols: cols.max });
    const titleRows = titleRange(pg.titleRows, rows), titleCols = titleRange(pg.titleCols, cols);
    const repeatH = titleRows ? span(rows, titleRows[0], titleRows[1]) : 0, repeatW = titleCols ? span(cols, titleCols[0], titleCols[1]) : 0;
    const headingW = pg.headings ? 34 : 0, headingH = pg.headings ? 20 : 0;
    result.areas = areas.map((area, index) => ({ area, index, x: cols.pos(area.c1), y: rows.pos(area.r1), w: span(cols, area.c1, area.c2), h: span(rows, area.r1, area.r2), rowBands: [], colBands: [], titleRows, titleCols }));
    const fitting = !!(pg.fitW || pg.fitH), rowBreaks = normalizePageBreaks(pg.rowBreaks, rows.max), colBreaks = normalizePageBreaks(pg.colBreaks, cols.max);
    result.manualIgnored = fitting && !!(rowBreaks.length || colBreaks.length);
    let scale = fitting ? 1 : pg.scale / 100;
    for (const item of result.areas) if (fitting) {
      const a = item.area, w = item.w + repeatW - overlap(cols, a.c1, a.c2, titleCols) + headingW, h = item.h + repeatH - overlap(rows, a.r1, a.r2, titleRows) + headingH;
      scale = Math.min(scale, printScale(pg, w, h));
    }
    const bandsFor = (item, atScale, cap = maxPages) => ({
      rows: splitAxis(rows, item.area.r1, item.area.r2, geometry.content.h / atScale - headingH, titleRows, fitting ? [] : rowBreaks, cap),
      cols: splitAxis(cols, item.area.c1, item.area.c2, geometry.content.w / atScale - headingW, titleCols, fitting ? [] : colBreaks, cap),
    });
    // 반복 제목 때문에 두 번째 쪽에서 폭/높이가 늘어나는 경우도 맞춤 쪽 수에 포함합니다.
    if (fitting) {
      const fits = (atScale) => result.areas.every((item) => { try { const b = bandsFor(item, atScale); return (!pg.fitW || b.cols.bands.length <= pg.fitW) && (!pg.fitH || b.rows.bands.length <= pg.fitH); } catch { return false; } });
      if (!fits(scale) && fits(.1)) { let lo = .1, hi = scale; for (let i = 0; i < 24; i++) { const mid = (lo + hi) / 2; if (fits(mid)) lo = mid; else hi = mid; } scale = lo; }
    }
    result.scale = scale;
    for (const item of result.areas) {
      const b = bandsFor(item, scale); item.scale = scale; item.rowBands = b.rows.bands; item.colBands = b.cols.bands;
      result.oversized ||= b.rows.oversized || b.cols.oversized;
      if (result.pages.length + item.rowBands.length * item.colBands.length > maxPages) throw new Error(`페이지가 ${maxPages.toLocaleString('ko-KR')}쪽을 초과합니다. 인쇄 영역이나 배율을 조정하세요.`);
      const add = (row, col) => result.pages.push({ number: result.pages.length + 1, areaIndex: item.index, r1: row.start, r2: row.end, c1: col.start, c2: col.end, bodyR1: row.bodyStart, bodyR2: row.bodyEnd, bodyC1: col.bodyStart, bodyC2: col.bodyEnd, x: cols.pos(col.start), y: rows.pos(row.start), w: span(cols, col.start, col.end), h: span(rows, row.start, row.end), scale });
      if (pg.order === 'overThenDown') for (const row of item.rowBands) for (const col of item.colBands) add(row, col);
      else for (const col of item.colBands) for (const row of item.rowBands) add(row, col);
      for (const [axis, list, dim] of [['row', item.rowBands, rows], ['col', item.colBands, cols]]) for (const band of list.slice(1)) result.breaks.push({ axis, index: band.start, position: dim.pos(band.start), areaIndex: item.index, manual: band.manual });
    }
    return result;
  } catch (error) { return { ...result, pages: [], breaks: [], error: error.message }; }
}
