import { chartModelData } from './chart.js';
import { parseInput } from './format.js';
import { tableAt, tableCellStyle } from './tables.js';
import { prepareCond, condFormatAt, ruleRanges } from './condfmt.js';

export const PUBLISH_CELL_LIMIT = 250000;
const STYLE_KEYS = 'font size bold italic underline strike color fill align valign wrap shrink rotate indent numFmt code decimals queryFormat checkbox pattern patternColor gradient bt btc bts bb bbc bbs bl blc bls br brc brs du duc dus dd ddc dds'.split(' ');
const SHEET_KEYS = 'name colWidths rowHeights merges hiddenRows hiddenCols rowManual freeze noGrid noZeros outline defRowH defColW zoom view tabColor'.split(' ');
const CHART_KEYS = 'id name x y w h type title titleSize titleColor titleBold titleOverlay textColor axisSize legend legendColor legendSize legendBold labels axes palette colors style fill border plotFill gridColor gridX gridY threeD view3D grouping gap overlap marker rounded dataTable hole firstAngle explode varyColors bubbleScale scatterStyle radarStyle ohlc upColor downColor totalColor totals connectors showMean binCount binWidth'.split(' ');
const SHAPE_KEYS = 'id name kind x y w h rotation rot flip flipV fill fillOpacity stroke strokeOpacity strokeWidth dash arrow grad shadow glow soft text color size font bold italic underline align valign margins wrap'.split(' ');
const IMAGE_KEYS = 'id name src alt x y w h sizing crop rotation rot flip flipV opacity'.split(' ');
const SERIES_KEYS = 'name values x size type axis color labels marker numFmt lineWidth dash smooth pointColors points _fi'.split(' ');
function pick(source, keys) {
  const out = {};
  for (const key of keys) if (source?.[key] !== undefined) out[key] = structuredClone(source[key]);
  return out;
}
function styleCopy(source) { return pick(source, STYLE_KEYS); }
function styleMap(source) {
  const out = {}; for (const [key, value] of Object.entries(source ?? {})) out[key] = styleCopy(value); return out;
}
function literalCell(value, style) {
  let raw = '', image;
  if (value?.type === 'image') image = pick(value, ['src', 'alt', 'sizing', 'w', 'h']);
  else if (value && typeof value === 'object') raw = String(value.code ?? value.error ?? '#VALUE!');
  else if (typeof value === 'boolean') raw = value ? 'TRUE' : 'FALSE';
  else if (typeof value === 'number') raw = Number.isFinite(value) ? String(value) : '#NUM!';
  else if (value != null) {
    raw = String(value);
    if (style.numFmt !== 'text') {
      const parsed = parseInput(raw);
      if (raw.startsWith('=') || raw.startsWith("'") || parsed.value !== raw || parsed.errorLiteral) raw = "'" + raw;
    }
  }
  // 텍스트 서식 수식이 숫자/논리/오류를 반환한 경우에도 고정된 값의 타입을 유지합니다.
  if (value != null && typeof value !== 'string' && style.numFmt === 'text') style.numFmt = 'general';
  return { raw, ...(Object.keys(style).length ? { style } : {}), ...(image ? { image } : {}) };
}
function chartSnapshot(wb, si, chart) {
  const data = chartModelData(wb, si, chart);
  const snapshotData = {
    date1904: !!wb.date1904,
    categories: structuredClone(data.categories ?? []),
    series: (data.series ?? []).map(series => pick(series, SERIES_KEYS)),
    ...(data.catLevels ? { catLevels: structuredClone(data.catLevels) } : {}),
  };
  // 원래 range/series/pivot/sheet 및 비표시 cache·확장 필드는 복사하지 않습니다.
  return { ...pick(chart, CHART_KEYS), snapshotData };
}

/** 전체 공유는 원본 전체, 개별 시트 공유는 표시 값만 가진 독립 스냅숏. 원본 셀/시트 속성을 수정하지 않습니다. */
export function publishedWorkbook(wb, selection = 'all', { maxCells = PUBLISH_CELL_LIMIT } = {}) {
  if (selection === 'all') return wb.serialize();
  const si = Number(selection), source = wb.sheets[si];
  if (!Number.isInteger(si) || !source) throw new Error('게시할 시트를 찾지 못했습니다. 시트를 다시 선택하세요.');
  let estimate = source.cells.size;
  for (const block of source.blocks ?? []) estimate += block.n * block.cols.length;
  if (estimate > maxCells) throw new Error('시트 하나 게시하기는 계산 셀 25만 개까지 지원합니다. 공개할 범위를 새 시트로 복사한 뒤 게시하세요.');
  const positions = new Set();
  const add = (r, c) => {
    positions.add(r + ',' + c);
    if (positions.size > maxCells) throw new Error('분산 결과를 포함한 게시 셀이 25만 개를 초과했습니다. 공개할 범위를 줄여 주세요.');
  };
  source.cells.forEachRC((cell, r, c) => add(r, c));
  for (const block of source.blocks ?? []) for (let r = block.r0; r < block.r0 + block.n; r++) for (let c = block.c0; c < block.c0 + block.cols.length; c++) add(r, c);
  // 계산은 읽기 API로 수행합니다. 수식/블록/셀 메타데이터를 게시본에 복사하지 않습니다.
  for (const position of positions) { const split = position.indexOf(','); wb.getValue(si, +position.slice(0, split), +position.slice(split + 1)); }
  for (const spill of wb.spillsOf(si)) {
    if (spill.h * spill.w > maxCells) throw new Error('분산 결과가 너무 커서 시트 하나를 게시할 수 없습니다. 공개할 범위를 줄여 주세요.');
    for (let r = spill.r; r < spill.r + spill.h; r++) for (let c = spill.c; c < spill.c + spill.w; c++) add(r, c);
  }
  // 빈 셀까지 대규모로 순회하는 조건부 서식은 미리 차단합니다.
  let condWork = 0;
  const used = wb.usedRange(si);
  for (const rule of source.cond ?? []) for (const range of ruleRanges(rule))
    condWork += Math.max(0, Math.min(range.r2, used.rows - 1) - range.r1 + 1) * Math.max(0, Math.min(range.c2, used.cols - 1) - range.c1 + 1);
  if (condWork > maxCells) throw new Error('조건부 서식의 계산 범위가 너무 큽니다. 공개할 범위를 새 시트로 복사한 뒤 게시하세요.');
  const conditions = source.cond?.length ? prepareCond(wb, si) : [];
  const sheet = { ...pick(source, SHEET_KEYS), cells: {}, state: 'visible', charts: [], images: [], shapes: [],
    colStyles: styleMap(source.colStyles), rowStyles: styleMap(source.rowStyles), allStyle: styleCopy(source.allStyle) };
  for (const position of positions) {
    const split = position.indexOf(','), r = +position.slice(0, split), c = +position.slice(split + 1);
    const value = wb.getValue(si, r, c), table = tableAt(source, r, c);
    let style = { ...(table ? tableCellStyle(table, r, c) : null), ...wb.styleAt(si, r, c) };
    if (conditions.length) {
      const result = condFormatAt(conditions, wb, si, r, c, value);
      if (result.style) style = { ...style, ...result.style };
    }
    const cell = literalCell(value, styleCopy(style));
    if (cell.raw || cell.image || cell.style) sheet.cells[position] = cell;
  }
  sheet.charts = (source.charts ?? []).map(chart => chartSnapshot(wb, si, chart));
  sheet.images = (source.images ?? []).map(image => pick(image, IMAGE_KEYS));
  sheet.shapes = (source.shapes ?? []).map(shape => pick(shape, SHAPE_KEYS));
  // names/props/VBA/themeXml/externals/pivotSnapshots는 허용 목록에 포함하지 않습니다.
  const book = { version: 1, sheets: [sheet], ...(wb.date1904 ? { date1904: true } : {}) };
  if (wb.defaultFont) book.defaultFont = pick(wb.defaultFont, ['name', 'size']);
  if (wb.baseStyle) book.baseStyle = styleCopy(wb.baseStyle);
  if (wb.theme) book.theme = structuredClone(wb.theme);
  return book;
}
