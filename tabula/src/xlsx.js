import { createZipStreamWriter } from './zip-stream.js';
import { slicerStartItem } from './slicer-window.js';
import { bindImportedPivotPresentation, importedPivotPresentationCurrent, pivotItemOrderSignature } from './pivot-import-presentation.js';
import { importedPivotToggleButtons } from './pivot-import-toggles.js';
import { makeImportedPivotSourceReference, importedPivotSourceReference, pivotSourceReferenceCurrent, pivotSourceBindingCurrent, pivotCacheLayoutBinding, pivotCacheLayoutCurrent } from './pivot-source-reference.js';
import { fontDesktopStyle } from './font-identity.js';
import { arrayCacheValue } from './array-cache.js';
import { hasSavedNameError } from './calculation-state.js';
import { createXmlChunks } from './xml-chunks.js';
import { pivotExportData, pivotValueStats } from './pivot-export-data.js';
import { xlsxRowPlan } from './xlsx-row-stream.js';
import { xlsxBoundedRowPlan } from './xlsx-compact-rows.js';
import { storedCellEntries } from './cell-storage.js';
import { readAutoFilter, readFilterSort, autoFilterXml, filterSortXml } from './xlsx-filter.js';
import { TABLE_VISUAL_KEYS } from './table-format.js';
import { normalizeObjectStyles, findObjectStyle, objectStylePatch, TABLE_STYLE_ELEMENTS, SLICER_STYLE_ELEMENTS } from './object-styles.js';
import { relocateValidation, VALIDATION_IME_MODES } from './validation.js';
import { chartAreaFormatXml, readChartAreaFormat } from './chart-area-drawingml.js';
import { readPhonetic, normalizePhonetic, phoneticXml } from './phonetic.js';
import { normalizeVideo, MEDIA_OBJECT_URI } from './media-object.js';
import { noteVisible } from './review-state.js';
import { isDrawingGroup, drawingGroupXml, readDrawingGroup, readDrawingHyperlink, drawingHyperlinkXml } from './smartart-xlsx.js';
// .xlsx 읽기/쓰기 (Office Open XML). DOM 없이 동작하므로 Node 에서도 테스트 가능.
import { isXlsb, convertXlsb, readPivotSnapshotBinary } from './xlsb.js';
import { isPivotSnapshot, readPivotSnapshotXml } from './pivot-cache-data.js';
import { unzip, unzipBlob, prepareZipEntry, prepareZipEntryRaw, deleteZipEntry, zipEntryInfo, zipEntryChunks, zip, zipAsync, createZipAsyncWriter, textOf } from './zip.js';
import { sheetXmlStream } from './xlsx-sheet-stream.js';
import { createImportedBlankRuns } from './import-blank-runs.js';
import { readSharedStrings } from './shared-strings.js';
import { CellMap } from './cellmap.js';
import { createImportedLiteralMemo } from './import-cell-memo.js';
import { protectFromAttrs, protectXml } from './protect.js';
import { pageXml, pageFromXml, normPage } from './page.js';
import { readThemeOptions, applyThemeOptionsXml } from './theme-options.js';
import { parseXml, child, kids, descendants, allText, esc, decodeEntities, unx } from './xml.js';
import {
  parse, tokenize, colToName, nameToCol, cellName, parseRangeName, FUNCS, isError,
  quoteSheetName, MAX_ROWS, MAX_COLS, EXCEL_MAX_ROWS, mayReturnArray, unknownFunctions,
} from './formula.js';
import { toFileFormula, fromFileFormula } from './xlfn.js';
import { parseInput, formatGeneral, fmtCode as fmtCodeRaw, fileCode, styleForCode, dateParts, serialOf } from './format.js';
const fmtCode = (style) => fileCode(fmtCodeRaw(style));
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, formulaShifter } from './workbook.js';
import { chartLayout, PALETTE, chartModelData, paletteOf } from './chart.js';
import { inferPivotCategorySeries } from './chart-source.js';
import { chartView3D } from './chart-3d.js';
import { chartTextFields, assignChartTextFields, readChartTextFont, chartTextPropertiesXml, chartRichTextXml, chartTextSupplement, applyChartTextSupplement } from './chart-text-xml.js';
import { isChartEx, writeChartEx, readChartEx, chartExStyleXml, chartExColorsXml, chartExDrawingProps, applyChartExOptions, CHARTEX_NS, CHARTEX_REL, CHARTEX_CONTENT, CHARTEX_STYLE_CONTENT, CHARTEX_COLOR_CONTENT } from './chart-ex.js';
import { Axis, hid, hidKeys } from './axis.js';
import { toBase64, fromBase64 } from './vba.js';
import { CellImage } from './fxcore.js';
import { pictureMediaSource, pictureBlipXml, pictureGeometryXml, pictureBorderXml, pictureEffectXml, readPictureAdjustments, pictureMetadataXml, restorePictureMetadata } from './picture-drawingml.js';
import { emfDataUrl, isEmf } from './emf.js';
import { GEOM, LINE_KINDS, shapeLineEnds } from './shapes.js';
import { customGeometryXml, readCustomGeometry, storedCustomGeometryXml } from './shape-path.js';
import { BLOCK_MIN_ROWS, ColBuilder, inBlock, blockValue, rawOf } from './block.js';
import { normalizeStyleName, DEFAULT_TABLE_STYLE, dataTop, dataBottom, canonicalRef, tableAt, columnNames, findTable, TOTAL_FUNCS } from './tables.js';
import { pivotSourceData, resolvePivot, itemText, keyOf, sortKeys, EMPTY, headerNames, normalizeDef, computePivot, valueName, showAsPercent, excelCalcFormula, pivotFilterKey, pivotPageMulti, pivotPageLayout, DATE_OP_TYPES } from './pivot.js';
import { groupKey, EMPTY_TEXT, itemIdentity } from './cube.js';
import { slicerStyleName, slicerColors, isModernSlicer, slicerStyleElements } from './slicerstyle.js';
import { SLICER_DEFAULT_BUTTON_HEIGHT } from './slicer-properties.js';
import { applyTint, DEFAULT_THEME, PRESET_STYLES, presetStyle, isModernStyle, ELEMENT_TYPES, elementDxfStyle } from './stylepresets.js';
import { maxOf, minOf, pushAll, DAY_MS } from './fxcore.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU = 9525; // 1px
const TABLE_STYLE_INHERIT_URI = '{3720B142-2CF8-4B6E-95A7-574958454C54}';
const TABLE_STYLE_INHERIT_NS = 'https://wixel.app/table-style/1';
// Native XF stores a neutral component; this extension remembers why it is neutral.
// Only visual channels can inherit from a table. File metadata cannot alter values or number formats.
function tableStyleInheritance(value) {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const out = {};
  for (const [key, neutral] of Object.entries(value)) {
    if (!TABLE_VISUAL_KEYS.includes(key)) continue;
    if (neutral === '' || neutral === false || neutral === null) out[key] = neutral;
  }
  return Object.keys(out).length ? out : null;
}

// 변환 이력만 보관한다. 재계산/원본 XML 보존을 뜻하지 않으며, 파일의 임의 문구는 표시하지 않는다.
const XLSX_IMPORT_WARNINGS = {
  olapDataModelValuesOnly: '이 문서의 데이터 모델·OLAP 피벗 테이블은 셀에 저장된 결과만 표시합니다. 관련 피벗·슬라이서 조작, DAX 계산 및 데이터 새로 고침은 지원하지 않습니다. XLSX로 다시 저장하면 데이터 모델과 해당 피벗·슬라이서 구성이 보존되지 않습니다. 원본 Excel 파일을 유지하세요.',
  dataTableValuesOnly: '이 문서는 가져올 때 가상 분석 데이터 표(TABLE)의 저장된 결과만 보존했습니다. WIXEL에서 입력을 바꾸어도 데이터 표는 재계산되지 않으며, XLSX에 저장하면 데이터 표 수식 없이 값으로 저장됩니다. 원래 데이터 표가 필요하면 원본 파일을 유지하세요.',
};
const LOG_AXIS_WARNING = '로그 축 설정은 XLSX에 보존하지만 WIXEL 화면에서는 선형 축으로 표시합니다. 차트의 간격과 눈금이 Excel과 다르므로 Excel에서 확인하세요.';
const importWarningCodes = (props) => Array.isArray(props?.xlsxImportWarnings) ? [...new Set(props.xlsxImportWarnings.filter(k => typeof k === 'string' && Object.hasOwn(XLSX_IMPORT_WARNINGS, k)))] : [];
/** 여러 문서의 알려진 변환 경고만 병합합니다. 파일이 넣은 임의 문구는 전달하지 않습니다. */
export const mergeXlsxImportWarnings = (...props) => [...new Set(props.flatMap(importWarningCodes))];
function warnOlapImport(ctx) { (ctx.importWarningCodes ??= new Set()).add('olapDataModelValuesOnly'); ctx.warnings.add(XLSX_IMPORT_WARNINGS.olapDataModelValuesOnly); }
const chartHasLogAxis = (chart) => ['x', 'y', 'y2'].some(k => chart.axes?.[k]?.logBase !== undefined);

/** XLSX 저장 전 안내. 값의 계산 또는 통합 문서 변경 없이 현재 보존 한계를 알린다. */
export function xlsxExportWarnings(wb) {
  const warnings = importWarningCodes(wb.props).map(k => XLSX_IMPORT_WARNINGS[k]);
  if (wb.sheets.some(sh => (sh.charts ?? []).some(chartHasLogAxis))) warnings.push(LOG_AXIS_WARNING);
  if (wb.calculation?.iterate) warnings.push('반복 계산 설정은 파일에 보존하지만 WIXEL은 반복 계산을 실행하지 않습니다. 저장된 계산 결과를 Excel에서 확인하세요.');
  if (wb.calculation?.fullPrecision === false) warnings.push('표시된 정밀도로 계산 설정은 파일에 보존하지만 WIXEL은 원래 숫자의 정밀도로 계산합니다. 저장된 계산 결과를 Excel에서 확인하세요.');
  return warnings;
}

// SpreadsheetML ST_Xstring: 이스케이프처럼 생긴 원문과 XML 금지 문자/CR도 보존.
// https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/bd0aa042-434a-4ca7-b25f-4e1fd25a954d
function xesc(value) {
  let s = String(value);
  if (s.includes('_x')) s = s.replace(/_(?=x[0-9A-Fa-f]{4}_)/g, '_x005F_');
  return esc(s.replace(/[\u0000-\u0008\u000b-\u001f\ufffe\uffff]/g, (c) => `_x${c.charCodeAt(0).toString(16).toUpperCase().padStart(4, '0')}_`));
}

const DEFAULT_FONT = '맑은 고딕';

// ───────────────────────── 공통 ─────────────────────────
// 엑셀 열 너비: 파일의 너비(w) × 기본 글꼴의 숫자 너비(MDW, 픽셀). 기본 글꼴이 맑은 고딕 11pt 면 MDW 8, Calibri 11pt 면 7
const DIGIT_EM = { calibri: 0.507, 'calibri light': 0.49, '맑은 고딕': 0.55, 'malgun gothic': 0.55, arial: 0.556, '굴림': 0.5, gulim: 0.5, '굴림체': 0.5, '돋움': 0.5, dotum: 0.5, '돋움체': 0.5, '바탕': 0.5, batang: 0.5, '나눔고딕': 0.6, nanumgothic: 0.6, 'nanum gothic': 0.6, 'times new roman': 0.5, cambria: 0.556, 'segoe ui': 0.55, verdana: 0.636, tahoma: 0.546, 'meiryo ui': 0.55, 'ms gothic': 0.5, simsun: 0.5 };
export const digitWidth = (font) => Math.max(4, Math.round(((font?.size || 11) * 96) / 72 * (DIGIT_EM[String(font?.name ?? '맑은 고딕').toLowerCase()] ?? 0.53)));
export const width2pxM = (w, mdw) => Math.max(0, Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw));
const px2widthM = (px, mdw) => Math.max(0, Math.round((px / mdw) * 256) / 256);
/** 기본 열 너비(글자 수, baseColWidth) → 픽셀: 8 픽셀 단위로 올림 (엑셀과 같음) */
export const baseColPx = (base, mdw) => Math.ceil((base * mdw + 5) / 8) * 8;
const WRITE_FONT = { name: '맑은 고딕', size: 11 };
const px2width = (px, mdw = digitWidth(WRITE_FONT)) => px2widthM(px, mdw);
const width2px = (w, mdw = 7) => width2pxM(w, mdw);
export const pt2px = (pt) => Math.round((pt * 4) / 3);
const px2pt = (px) => Math.round(px * 0.75 * 100) / 100;

function refToRange(ref) {
  const rg = parseRangeName(ref.replace(/\$/g, ''));
  return rg;
}

function rangeRef(rg, abs = false) {
  const a = (r, c) => (abs ? `$${colToName(c)}$${r + 1}` : cellName(r, c));
  return rg.r1 === rg.r2 && rg.c1 === rg.c2 ? a(rg.r1, rg.c1) : `${a(rg.r1, rg.c1)}:${a(rg.r2, rg.c2)}`;
}

/** Print_Area / Print_Titles unions. A comma inside a quoted sheet name is not a separator. */
export function parsePrintAreas(text, sheetName, { strict = true } = {}) {
  const parts = []; let start = 0, quoted = false;
  const source = String(text ?? '').trim();
  for (let i = 0; i < source.length; i++) {
    if (source[i] === "'") { if (quoted && source[i + 1] === "'") i++; else quoted = !quoted; }
    else if (source[i] === ',' && !quoted) { parts.push(source.slice(start, i)); start = i + 1; }
  }
  if (quoted) return [];
  parts.push(source.slice(start));
  const out = [];
  for (let part of parts) {
    part = part.trim().replace(/^[=(\s]+|[)\s]+$/g, '');
    const bang = part.lastIndexOf('!');
    if (bang >= 0) {
      const raw = part.slice(0, bang), sheet = raw.startsWith("'") && raw.endsWith("'") ? raw.slice(1, -1).replace(/''/g, "'") : raw;
      if (sheet.includes('[') || sheetName !== undefined && sheet.toLowerCase() !== String(sheetName).toLowerCase()) { if (strict) return []; continue; }
      part = part.slice(bang + 1);
    }
    part = part.replace(/\$/g, '');
    const rows = /^(\d+):(\d+)$/.exec(part), cols = /^([A-Z]+):([A-Z]+)$/i.exec(part);
    const rg = rows ? { r1: Number(rows[1]) - 1, c1: 0, r2: Number(rows[2]) - 1, c2: MAX_COLS - 1 }
      : cols ? { r1: 0, c1: nameToCol(cols[1].toUpperCase()), r2: EXCEL_MAX_ROWS - 1, c2: nameToCol(cols[2].toUpperCase()) }
        : /^[A-Z]+[1-9]\d*(?::[A-Z]+[1-9]\d*)?$/i.test(part) ? parseRangeName(part) : null;
    if (!rg) { if (strict) return []; continue; }
    const r = { r1: Math.min(rg.r1, rg.r2), c1: Math.min(rg.c1, rg.c2), r2: Math.max(rg.r1, rg.r2), c2: Math.max(rg.c1, rg.c2) };
    if (!Object.values(r).every(Number.isInteger) || r.r1 < 0 || r.c1 < 0 || r.r2 >= EXCEL_MAX_ROWS || r.c2 >= MAX_COLS) { if (strict) return []; continue; }
    if (!out.some(v => v.r1 === r.r1 && v.r2 === r.r2 && v.c1 === r.c1 && v.c2 === r.c2)) out.push(r);
  }
  return out;
}

function printAreaRef(rg) {
  if (rg.r1 === 0 && rg.r2 === EXCEL_MAX_ROWS - 1) return `$${colToName(rg.c1)}:$${colToName(rg.c2)}`;
  if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) return `$${rg.r1 + 1}:$${rg.r2 + 1}`;
  return rangeRef(rg, true);
}

const validBreaks = (values, max) => [...new Set((Array.isArray(values) ? values : []).filter(n => Number.isInteger(n) && n > 0 && n < max))].sort((a, b) => a - b);
function pageBreakXml(page) {
  return [['rowBreaks', EXCEL_MAX_ROWS, MAX_COLS], ['colBreaks', MAX_COLS, EXCEL_MAX_ROWS]].map(([tag, max, span]) => {
    const breaks = validBreaks(page?.[tag], max);
    return breaks.length ? `<${tag} count="${breaks.length}" manualBreakCount="${breaks.length}">${breaks.map(id => `<brk id="${id}" min="0" max="${span - 1}" man="1"/>`).join('')}</${tag}>` : '';
  }).join('');
}

/** 문자열 값이 입력 해석으로 다른 값이 되지 않도록 raw 생성 */
export function textRaw(s) {
  if (s === '') return "'"; // 빈 글자("") 셀: 빈 칸과 달리 COUNTA · 피벗 개수에 셈 (엑셀과 같음)
  if (s.startsWith('=') || s.startsWith("'")) return `'${s}`;
  const p = parseInput(s);
  return typeof p.value === 'string' && p.value === s && !p.errorLiteral ? s : `'${s}`;
}

// ───────────────────────── 색상 ─────────────────────────
export const INDEXED = ('000000,FFFFFF,FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,000000,FFFFFF,FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,'
  + '800000,008000,000080,808000,800080,008080,C0C0C0,808080,9999FF,993366,FFFFCC,CCFFFF,660066,FF8080,0066CC,CCCCFF,'
  + '000080,FF00FF,FFFF00,00FFFF,800080,800000,008080,0000FF,00CCFF,CCFFFF,CCFFCC,FFFF99,99CCFF,FF99CC,CC99FF,FFCC99,'
  + '3366FF,33CCCC,99CC00,FFCC00,FF9900,FF6600,666699,969696,003366,339966,003300,333300,993300,993366,333399,333333,000000,FFFFFF').split(',');

function colorOf(el, theme) {
  if (!el) return null;
  const a = el.attrs;
  let hex = null;
  if (a.rgb) hex = a.rgb.length === 8 ? a.rgb.slice(2) : a.rgb;
  else if (a.theme !== undefined) hex = theme[Number(a.theme)] ?? null;
  else if (a.indexed !== undefined) hex = (theme.indexedColors ?? INDEXED)[Number(a.indexed)] ?? null;
  if (!hex) return null;
  hex = applyTint(hex.toUpperCase(), Number(a.tint || 0));
  return `#${hex.toLowerCase()}`;
}

const argb = (color) => `FF${String(color).replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6)}`;

// 계산/구조 보호 설정은 표준 속성만 보존한다. 엔진 미지원 옵션을 지원한다고 해석하지 않는다.
const CALC_BOOLS = ['fullCalcOnLoad', 'iterate', 'fullPrecision', 'calcCompleted', 'calcOnSave', 'concurrentCalc', 'forceFullCalc'];
const CALC_INTS = ['calcId', 'iterateCount', 'concurrentManualCount'];
const WORKBOOK_PROTECTION_KEYS = ['workbookPassword', 'revisionsPassword', 'lockStructure', 'lockWindows', 'lockRevision', 'workbookAlgorithmName', 'workbookHashValue', 'workbookSaltValue', 'workbookSpinCount', 'revisionsAlgorithmName', 'revisionsHashValue', 'revisionsSaltValue', 'revisionsSpinCount'];
function calculationFromAttrs(a) {
  const out = { mode: ['auto', 'manual', 'autoNoTable'].includes(a.calcMode) ? a.calcMode : 'auto' };
  for (const k of CALC_BOOLS) if (a[k] !== undefined) out[k] = !falseAttr(a[k]);
  for (const k of CALC_INTS) if (/^\d+$/.test(String(a[k] ?? ''))) out[k] = Number(a[k]);
  if (a.iterateDelta !== undefined && Number.isFinite(Number(a.iterateDelta)) && Number(a.iterateDelta) >= 0) out.iterateDelta = Number(a.iterateDelta);
  if (['A1', 'R1C1'].includes(a.refMode)) out.refMode = a.refMode;
  return out;
}
function calculationXml(calc) {
  if (!calc) return '<calcPr calcId="191029" fullCalcOnLoad="1"/>';
  const a = { calcMode: ['auto', 'manual', 'autoNoTable'].includes(calc.mode) ? calc.mode : 'auto' };
  for (const k of CALC_BOOLS) if (typeof calc[k] === 'boolean') a[k] = calc[k] ? 1 : 0;
  for (const k of CALC_INTS) if (Number.isSafeInteger(calc[k]) && calc[k] >= 0) a[k] = calc[k];
  if (Number.isFinite(calc.iterateDelta) && calc.iterateDelta >= 0) a.iterateDelta = calc.iterateDelta;
  if (['A1', 'R1C1'].includes(calc.refMode)) a.refMode = calc.refMode;
  return `<calcPr${Object.entries(a).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>`;
}
function workbookProtectionXml(props) {
  const a = Object.fromEntries(WORKBOOK_PROTECTION_KEYS.filter(k => props?.workbookProtection?.[k] !== undefined).map(k => [k, props.workbookProtection[k]]));
  if (typeof props?.lockStructure === 'boolean') a.lockStructure = props.lockStructure ? '1' : '0';
  return Object.keys(a).length ? `<workbookProtection${Object.entries(a).map(([k, v]) => ` ${k}="${esc(v)}"`).join('')}/>` : '';
}

// ───────────────────────── 표시 형식 ─────────────────────────
export const BUILTIN_FMT = {
  0: {}, 1: { decimals: 0 }, 2: { decimals: 2 }, 3: { numFmt: 'comma' }, 4: { numFmt: 'number', decimals: 2 },
  9: { numFmt: 'percent' }, 10: { numFmt: 'percent', decimals: 2 }, 11: { numFmt: 'scientific', decimals: 2 },
  12: { numFmt: 'fraction' }, 13: { numFmt: 'fraction' }, 14: { numFmt: 'date' }, 15: { numFmt: 'date' },
  16: { numFmt: 'date' }, 17: { numFmt: 'date' }, 18: { numFmt: 'time' }, 19: { numFmt: 'time' },
  20: { numFmt: 'time' }, 21: { numFmt: 'time' }, 22: { numFmt: 'datetime' }, 37: { numFmt: 'comma' },
  38: { numFmt: 'comma' }, 39: { numFmt: 'number', decimals: 2 }, 40: { numFmt: 'number', decimals: 2 },
  44: { numFmt: 'accounting' }, 45: { numFmt: 'time' }, 46: { numFmt: 'time' }, 47: { numFmt: 'time' },
  48: { numFmt: 'scientific', decimals: 1 }, 49: { numFmt: 'text' },
};
for (const id of [27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58]) BUILTIN_FMT[id] = { numFmt: 'date' };
// 시간 기본 형식: 코드 그대로 (모두 'time'(오전/오후)으로 묶으면 24시간 h:mm:ss 가 '오후 3:45:00' 으로 보임)
const TIME_CODES = {
  18: '[$-412]AM/PM h:mm', 19: '[$-412]AM/PM h:mm:ss', 20: 'h:mm', 21: 'h:mm:ss', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0',
  32: 'h"시" mm"분"', 33: 'h"시" mm"분" ss"초"', 55: '[$-412]AM/PM h"시" mm"분"', 56: '[$-412]AM/PM h"시" mm"분" ss"초"',
};
for (const [id, code] of Object.entries(TIME_CODES)) BUILTIN_FMT[id] = styleForCode(code);
const BUILTIN_CODE_ID = { '# ?/?': 12, '0.00E+00': 11, '@': 49, 'h:mm': 20, 'h:mm:ss': 21, 'mm:ss': 45, '[h]:mm:ss': 46 };
// 기본 제공 번호 중 음수 괄호 · 빨강 · 통화 · 회계 형식은 실제 서식 코드로 (한국어 엑셀 기준)
export const BUILTIN_CODE = {
  5: '"₩"#,##0;"₩"\\-#,##0', 6: '"₩"#,##0;[Red]"₩"\\-#,##0', 7: '"₩"#,##0.00;"₩"\\-#,##0.00', 8: '"₩"#,##0.00;[Red]"₩"\\-#,##0.00',
  37: '#,##0_);(#,##0)', 38: '#,##0_);[Red](#,##0)', 39: '#,##0.00_);(#,##0.00)', 40: '#,##0.00_);[Red](#,##0.00)',
  41: '_-* #,##0_-;\\-* #,##0_-;_-* "-"_-;_-@_-', 42: '_-"₩"* #,##0_-;\\-"₩"* #,##0_-;_-"₩"* "-"_-;_-@_-',
  43: '_-* #,##0.00_-;\\-* #,##0.00_-;_-* "-"??_-;_-@_-', 44: '_-"₩"* #,##0.00_-;\\-"₩"* #,##0.00_-;_-"₩"* "-"??_-;_-@_-',
};

// ───────────────────────── 읽기 ─────────────────────────
function relsOf(files, path) {
  const dir = path.slice(0, path.lastIndexOf('/') + 1);
  const relPath = `${dir}_rels/${path.slice(dir.length)}.rels`;
  const xml = textOf(files[relPath]);
  const out = {};
  if (!xml) return out;
  for (const r of kids(parseXml(xml), 'Relationship')) {
    let target = r.attrs.Target;
    if (r.attrs.TargetMode === 'External') {
      out[r.attrs.Id] = { target, rawTarget: target, type: r.attrs.Type.split('/').pop(), external: true, targetMode: 'External' };
      continue;
    }
    if (target.startsWith('/')) target = target.slice(1);
    else {
      const parts = (dir + target).split('/');
      const res = [];
      for (const p of parts) { if (p === '..') res.pop(); else if (p !== '.') res.push(p); }
      target = res.join('/');
    }
    out[r.attrs.Id] = { target, rawTarget: r.attrs.Target, type: r.attrs.Type.split('/').pop(), ...(r.attrs.TargetMode === 'Internal' ? { targetMode: 'Internal' } : {}) };
  }
  return out;
}

const rid = (el) => {
  const k = Object.keys(el.attrs).find((a) => a === 'r:id' || a.endsWith(':id') || a === 'id');
  return k ? el.attrs[k] : null;
};

function readTheme(files, wbRels) {
  const rel = Object.values(wbRels).find((r) => r.type === 'theme');
  const xml = rel && textOf(files[rel.target]);
  if (!xml) return DEFAULT_THEME;
  const scheme = descendants(parseXml(xml), 'clrScheme')[0];
  if (!scheme) return DEFAULT_THEME;
  const get = (name) => {
    const el = child(scheme, name);
    const c = el?.children[0];
    return (c?.attrs.lastClr ?? c?.attrs.val ?? '').toUpperCase() || null;
  };
  const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  return order.map((n, i) => get(n) ?? DEFAULT_THEME[i] ?? '000000');
}

const STYLE_APPLY = { number: 'applyNumberFormat', alignment: 'applyAlignment', font: 'applyFont', border: 'applyBorder', fill: 'applyFill', protection: 'applyProtection' };
const styleNameKey = (name) => String(name ?? '').trim().toLowerCase();
const falseAttr = (value) => value === '0' || value === 'false';
const trueAttr = (value) => value === '1' || value === 'true';

function readStyles(files, wbRels, theme) {
  const rel = Object.values(wbRels).find((r) => r.type === 'styles');
  const xml = rel && textOf(files[rel.target]);
  if (!xml) return { xfs: [], dxfs: [], defaultFont: null };
  const root = parseXml(xml);
  const indexed = kids(child(child(root, 'colors'), 'indexedColors'), 'rgbColor');
  if (indexed.length) theme.indexedColors = indexed.map((c, i) => /^[0-9a-f]{6}([0-9a-f]{2})?$/i.test(c.attrs.rgb ?? '') ? c.attrs.rgb.slice(-6).toUpperCase() : INDEXED[i]);
  const numFmts = {};
  for (const f of kids(child(root, 'numFmts'), 'numFmt')) numFmts[f.attrs.numFmtId] = f.attrs.formatCode;
  const fontOf = (f) => {
    const st = {};
    if (child(f, 'b')) st.bold = !falseAttr(child(f, 'b').attrs.val);
    if (child(f, 'i')) st.italic = !falseAttr(child(f, 'i').attrs.val);
    if (child(f, 'u')) st.underline = child(f, 'u').attrs.val !== 'none';
    if (child(f, 'strike')) st.strike = !falseAttr(child(f, 'strike').attrs.val);
    const sz = Number(child(f, 'sz')?.attrs.val);
    if (sz) st.size = sz;
    const color = colorOf(child(f, 'color'), theme);
    // 검정은 기본값으로 보지만, RGB 로 직접 정한 검정은 셀 서식 (엑셀: 표 스타일의 흰 머리글 글자보다 우선)
    if (color && (color !== '#000000' || child(f, 'color')?.attrs.rgb)) st.color = color;
    const name = child(f, 'name')?.attrs.val;
    if (name) st.font = name;
    return st;
  };
  const fonts = kids(child(root, 'fonts'), 'font').map(fontOf);
  const defaultFont = fonts[0]?.font ?? null;
  const wbFont = { name: fonts[0]?.font ?? '맑은 고딕', size: fonts[0]?.size ?? 11 };
  const defaultSize = wbFont.size;
  const fillOf = (f, dxf = false) => {
    const pf = child(f, 'patternFill');
    if (!pf) {
      const gf = child(f, 'gradientFill');
      const stops = descendants(f, 'stop').map((s) => [Number(s.attrs.position) || 0, colorOf(child(s, 'color'), theme) ?? '#ffffff']);
      if (!stops.length) return null;
      if (dxf || !gf) return stops[0][1];
      // 그라데이션 채우기 (엑셀 [채우기 효과]): 선형(degree) 또는 사각 경로(type="path")
      const a = gf.attrs;
      const gradient = a.type === 'path'
        ? { path: true, l: Number(a.left) || 0, r: Number(a.right) || 0, t: Number(a.top) || 0, b: Number(a.bottom) || 0, stops }
        : { deg: Number(a.degree) || 0, stops };
      return { fill: stops[0][1], gradient };
    }
    if (!dxf && (pf.attrs.patternType === 'none' || !pf.attrs.patternType)) return null;
    // 무늬 채우기 (solid 가 아님): 배경색 = bgColor, 무늬 색 = fgColor
    const pt = pf.attrs.patternType;
    if (!dxf && pt && pt !== 'solid') {
      const out = { pattern: pt, patternColor: colorOf(child(pf, 'fgColor'), theme) ?? '#000000' };
      const bg = colorOf(child(pf, 'bgColor'), theme);
      if (bg && child(pf, 'bgColor')?.attrs.indexed !== '64') out.fill = bg;
      return out;
    }
    return colorOf(child(pf, dxf ? 'bgColor' : 'fgColor') ?? child(pf, 'fgColor'), theme);
  };
  const fills = kids(child(root, 'fills'), 'fill').map((f) => fillOf(f));
  // 테두리: 선 종류(thin 이 아니면 bts…)와 색(검정이 아니면 btc…)까지 (엑셀과 같은 모양)
  const borderOf = (b) => {
    const st = {};
    const side = (key, ...names) => {
      const e = names.map((n) => child(b, n)).find((x) => x && x.attrs.style && x.attrs.style !== 'none');
      if (!e) return;
      st[key] = true;
      if (e.attrs.style !== 'thin') st[`${key}s`] = e.attrs.style;
      const c = colorOf(child(e, 'color'), theme);
      if (c && c !== '#000000') st[`${key}c`] = c;
    };
    side('bt', 'top');
    side('bb', 'bottom');
    side('bl', 'left', 'start');
    side('br', 'right', 'end');
    // 대각선: diagonalDown(↘) · diagonalUp(↗) 가 같은 선을 씀
    if (b?.attrs.diagonalDown === '1' || b?.attrs.diagonalUp === '1') {
      const tmp = {};
      const e = child(b, 'diagonal');
      if (e && e.attrs.style && e.attrs.style !== 'none') {
        const c = colorOf(child(e, 'color'), theme);
        for (const k of [b.attrs.diagonalDown === '1' ? 'dd' : null, b.attrs.diagonalUp === '1' ? 'du' : null].filter(Boolean)) {
          tmp[k] = true;
          if (e.attrs.style !== 'thin') tmp[`${k}s`] = e.attrs.style;
          if (c && c !== '#000000') tmp[`${k}c`] = c;
        }
      }
      Object.assign(st, tmp);
    }
    return st;
  };
  const borders = kids(child(root, 'borders'), 'border').map(borderOf);
  const numFmtOf = (id) => {
    const n = Number(id);
    if (numFmts[id]) return styleForCode(numFmts[id]);
    if (BUILTIN_CODE[n]) return styleForCode(BUILTIN_CODE[n]);
    if (BUILTIN_FMT[n]) return BUILTIN_FMT[n];
    return {};
  };
  // 셀 스타일(cellStyleXfs): 셀 서식에 맞춤이 없으면 부모 스타일의 맞춤을 물려받음
  // (한국어 엑셀의 '표준' 스타일은 세로 가운데 → 거의 모든 셀이 세로 가운데로 보임)
  const styleXfs = kids(child(root, 'cellStyleXfs'), 'xf');
  const xfStyle = (xf, parent) => {
    const a = xf.attrs;
    // 명시적 apply*=0 또는 생략된 구성요소는 이름 있는 부모 스타일에서 가져옵니다.
    // 부모가 그 구성요소를 적용하지 않도록 한 경우는 통합 문서 기본 구성요소를 씁니다.
    const inherited = (flag) => parent && !falseAttr(parent.attrs[flag]);
    const componentId = (key, flag) => Number(parent && (falseAttr(a[flag]) || a[key] === undefined)
      ? inherited(flag) ? parent.attrs[key] ?? 0 : 0 : a[key] ?? 0);
    const componentNode = (name, flag) => {
      const own = child(xf, name);
      if (parent && falseAttr(a[flag])) return inherited(flag) ? child(parent, name) : null;
      return own ?? (trueAttr(a[flag]) ? null : inherited(flag) ? child(parent, name) : null);
    };
    const st = { ...fonts[componentId('fontId', 'applyFont')] };
    // 확인란: 엑셀 365 (xfpb:xfComplement → featurePropertyBag 의 CellControl) · WIXEL 확장
    if (descendants(xf, 'xfComplement').length || descendants(xf, 'wx:checkbox').length || descendants(xf, 'checkbox').length) st.checkbox = true;
    if (st.font && st.font === defaultFont) delete st.font;
    if (st.size === defaultSize) delete st.size; // 통합 문서 기본 크기는 적지 않음 (기본 글꼴로 표시)
    const fill = fills[componentId('fillId', 'applyFill')];
    if (fill && typeof fill === 'object') Object.assign(st, fill);
    else if (fill) st.fill = fill;
    Object.assign(st, borders[componentId('borderId', 'applyBorder')] ?? {});
    Object.assign(st, numFmtOf(componentId('numFmtId', 'applyNumberFormat')));
    if (componentId('numFmtId', 'applyNumberFormat') === 0 && (trueAttr(a.applyNumberFormat)
      || Number(a.xfId) > 0 && falseAttr(a.applyNumberFormat) && trueAttr(parent?.attrs.applyNumberFormat))) st.numFmt = 'general';
    const al = componentNode('alignment', 'applyAlignment');
    if (al) {
      const h = al.attrs.horizontal;
      if (h === 'general' || h === 'left' || h === 'center' || h === 'right') st.align = h;
      else if (h === 'centerContinuous') st.align = 'centerContinuous';
      else if (h === 'distributed') st.align = 'center';
      else if (h === 'justify') { st.align = 'left'; st.wrap = true; }
      const v = al.attrs.vertical;
      if (v === 'top' || v === 'bottom') st.valign = v;
      else if (v === 'center' || v === 'justify' || v === 'distributed') st.valign = 'middle';
      if (al.attrs.wrapText === '1' || al.attrs.wrapText === 'true') st.wrap = true;
      if (Number(al.attrs.indent)) st.indent = Number(al.attrs.indent);
      const rot = Number(al.attrs.textRotation ?? 0);
      if (rot) st.rotate = rot === 255 ? 255 : rot > 90 ? -(rot - 90) : rot; // 255 = 세로 쓰기
      if (al.attrs.shrinkToFit === '1' || al.attrs.shrinkToFit === 'true') st.shrink = true;
      else if (falseAttr(al.attrs.shrinkToFit)) st.shrink = false;
    }
    // 셀 보호: 잠금 해제 · 수식 숨기기
    const pr = componentNode('protection', 'applyProtection');
    if (pr) {
      if (pr.attrs.locked === '0' || pr.attrs.locked === 'false') st.locked = false;
      if (pr.attrs.hidden === '1' || pr.attrs.hidden === 'true') st.hideFormula = true;
    }
    const extension = kids(child(xf, 'extLst'), 'ext').find(e => e.attrs.uri?.toUpperCase() === TABLE_STYLE_INHERIT_URI);
    const inheritance = child(extension, 'tableStyleInherit');
    if (inheritance?.attrs['xmlns:wx'] === TABLE_STYLE_INHERIT_NS) {
      try {
        const marker = tableStyleInheritance(JSON.parse(inheritance.attrs.json ?? 'null'));
        if (marker) {
          for (const [key, neutral] of Object.entries(marker)) {
            // Native formatting wins if another editor changed the component after export.
            if (st[key] === undefined || st[key] === neutral) st[key] = neutral;
            else delete marker[key];
          }
          if (Object.keys(marker).length) st.tableStyleInherit = marker;
        }
      } catch { /* Ignore malformed optional metadata; native Excel formatting remains usable. */ }
    }
    return st;
  };
  // 표준은 baseStyle로 보존하고 다른 이름/구성요소/숨김 상태는 갤러리 모델로 전달합니다.
  // 중복 이름은 대소문자와 양끝 공백을 무시해 첫 정의를 유지합니다.
  const cellStyles = [], namedByKey = new Map(), nameByXf = new Map();
  for (const item of kids(child(root, 'cellStyles'), 'cellStyle')) {
    const a = item.attrs, id = Number(a.xfId), name = unx(a.name ?? '').trim(), key = styleNameKey(name);
    if (!Number.isInteger(id) || id < 0 || !styleXfs[id] || !name) continue;
    if (a.builtinId === '0' || key === '표준' || key === 'normal') continue;
    const previous = namedByKey.get(key);
    if (previous) { if (!nameByXf.has(id)) nameByXf.set(id, previous.name); continue; }
    const include = {};
    for (const [part, flag] of Object.entries(STYLE_APPLY)) include[part] = !falseAttr(styleXfs[id].attrs[flag]);
    const itemStyle = { name, style: xfStyle(styleXfs[id], null), include };
    if (/^\d+$/.test(a.builtinId ?? '')) itemStyle.builtinId = Number(a.builtinId);
    if (a.customBuiltin !== undefined) itemStyle.customBuiltin = trueAttr(a.customBuiltin);
    if (trueAttr(a.hidden)) itemStyle.hidden = true;
    if (/^\d+$/.test(a.iLevel ?? '')) itemStyle.iLevel = Number(a.iLevel);
    cellStyles.push(itemStyle); namedByKey.set(key, itemStyle);
    if (!nameByXf.has(id)) nameByXf.set(id, name);
  }
  const xfs = kids(child(root, 'cellXfs'), 'xf').map((xf) => {
    const id = Number(xf.attrs.xfId ?? 0), style = xfStyle(xf, styleXfs[id]);
    const name = nameByXf.get(id);
    if (name) style.cellStyleName = name;
    return style;
  });
  const dxfOf = (d) => {
    const st = {}, f = child(d, 'font');
    if (f) {
      Object.assign(st, fontOf(f));
      const c = colorOf(child(f, 'color'), theme);
      if (c) st.color = c;
    }
    const fill = child(d, 'fill');
    if (fill) { const c = fillOf(fill, true); if (c) st.fill = c; }
    const bd = child(d, 'border');
    if (bd) Object.assign(st, borderOf(bd));
    const nf = child(d, 'numFmt');
    if (nf?.attrs.formatCode) Object.assign(st, styleForCode(nf.attrs.formatCode));
    return st;
  };
  const dxfNodes = kids(child(root, 'dxfs'), 'dxf');
  const dxfs = dxfNodes.map(dxfOf);
  // Color filters (including font-color filters) use solid foreground fills.
  const filterDxfs = dxfNodes.map(d => {
    const st = dxfOf(d), fill = child(d, 'fill'), pf = child(fill, 'patternFill');
    if (pf) st.fill = pf.attrs.patternType === 'none' ? null : colorOf(child(pf, 'fgColor'), theme) ?? colorOf(child(pf, 'bgColor'), theme) ?? st.fill;
    return st;
  });
  const objectDxfOf = d => {
    const st=dxfOf(d),f=child(d,'font');
    Object.assign(st,fontOf(f));
    for(const [tag,key]of [['b','bold'],['i','italic'],['strike','strike']])if(child(f,tag))st[key]=!falseAttr(child(f,tag).attrs.val);
    if(child(f,'u'))st.underline=child(f,'u').attrs.val!=='none';
    const fc=colorOf(child(f,'color'),theme);if(fc)st.color=fc;
    const fill=child(d,'fill'),pf=child(fill,'patternFill');
    if(fill){const val=fillOf(fill,!(child(fill,'gradientFill')||pf?.attrs.patternType&&pf.attrs.patternType!=='solid'));if(val&&typeof val==='object')Object.assign(st,val);else if(val)st.fill=val;}
    if(pf?.attrs.patternType==='none')st.fill=null;
    const b=child(d,'border');
    for(const [tag,key]of [['left','bl'],['right','br'],['top','bt'],['bottom','bb'],['vertical','bv'],['horizontal','bh']]){
      const e=child(b,tag);if(!e)continue;st[key]=!!e.attrs.style&&e.attrs.style!=='none';
      if(st[key]){if(e.attrs.style!=='thin')st[key+'s']=e.attrs.style;const c=colorOf(child(e,'color'),theme);if(c)st[key+'c']=c;}
    }
    for(const [attr,key]of [['diagonalDown','dd'],['diagonalUp','du']])if(b?.attrs[attr]!==undefined&&!st[key])st[key]=false;
    return st;
  };
  const tableDxfs = dxfNodes.map(objectDxfOf);
  const x14DxfRoot=kids(child(root,'extLst'),'ext').find(e=>e.attrs.uri?.toUpperCase()==='{46F421CA-312F-682F-3DD2-61675219B42D}');
  const slicerDxfNodes=x14DxfRoot?kids(child(x14DxfRoot,'dxfs'),'dxf'):dxfNodes;
  const readElement=(e,nodes=dxfNodes)=>{const d=nodes[Number(e.attrs.dxfId)],style=d?objectDxfOf(d):{};return {type:e.attrs.type,...(/Stripe$/.test(e.attrs.type)?{size:Number(e.attrs.size)>0?Number(e.attrs.size):1}:{}),style,...(d?{sourceDxf:structuredClone(d),sourceStyle:structuredClone(style)}:{})};};
  const tableStylesNode=child(root,'tableStyles'),slicerStylesNode=descendants(root,'slicerStyles')[0];
  const slicerStyleNames=new Set(kids(slicerStylesNode,'slicerStyle').map(ss=>ss.attrs.name));
  const objectStyles=normalizeObjectStyles({
    // 동명 false/false 보조 정의는 아래 슬라이서에 흡수한다. 미참조 정의는 보존한다.
    tables:kids(tableStylesNode,'tableStyle').filter(ts=>!(falseAttr(ts.attrs.table)&&falseAttr(ts.attrs.pivot)&&slicerStyleNames.has(ts.attrs.name))).map(ts=>({name:ts.attrs.name,table:!falseAttr(ts.attrs.table),pivot:!falseAttr(ts.attrs.pivot),elements:kids(ts,'tableStyleElement').map(e=>readElement(e))})),
    slicers:kids(slicerStylesNode,'slicerStyle').map(ss=>({name:ss.attrs.name,elements:[...kids(kids(tableStylesNode,'tableStyle').find(t=>t.attrs.name===ss.attrs.name),'tableStyleElement').filter(e=>['wholeTable','headerRow'].includes(e.attrs.type)).map(e=>readElement(e)),...descendants(ss,'slicerStyleElement').map(e=>readElement(e,slicerDxfNodes))]})),
    ...(tableStylesNode?.attrs.defaultTableStyle?{defaultTableStyle:tableStylesNode.attrs.defaultTableStyle}:{}),
    ...(tableStylesNode?.attrs.defaultPivotStyle?{defaultPivotStyle:tableStylesNode.attrs.defaultPivotStyle}:{}),
    ...(slicerStylesNode?.attrs.defaultSlicerStyle?{defaultSlicerStyle:slicerStylesNode.attrs.defaultSlicerStyle}:{})
  });
  // 사용자 지정 피벗 스타일 (<tableStyles>) → 역할별 서식 { header, sub, grand, body, page, band }
  const tableStyles = {};
  for (const ts of kids(child(root, 'tableStyles'), 'tableStyle')) {
    if (ts.attrs.pivot === '0') continue;
    const el = {};
    for (const e of kids(ts, 'tableStyleElement')) el[e.attrs.type] = dxfs[Number(e.attrs.dxfId)] ?? {};
    const pick = (...types) => Object.assign({}, ...types.map((t) => el[t] ?? {}));
    const body = pick('wholeTable');
    const clean = (o) => { const x = { ...o }; delete x.numFmt; delete x.decimals; delete x.code; return x; };
    tableStyles[ts.attrs.name] = {
      header: clean({ ...pick('headerRow', 'firstHeaderCell') }),
      sub: clean(pick('firstSubtotalRow', 'firstRowSubheading')),
      grand: clean(pick('totalRow')),
      body: clean({ ...(body.fill ? { fill: body.fill } : {}), ...(body.color ? { color: body.color } : {}) }),
      page: clean(pick('pageFieldLabels')),
      band: clean(pick('firstRowStripe')),
    };
  }
  // 사용자 지정 슬라이서 스타일 (x14:slicerStyles + 같은 이름의 tableStyle) → 슬라이서 색
  const slicerStyles = {};
  const tsByName = new Map(kids(child(root, 'tableStyles'), 'tableStyle').map((ts) => [ts.attrs.name, ts]));
  for (const ss of descendants(root, 'slicerStyle')) {
    const el = {};
    for (const e of descendants(ss, 'slicerStyleElement')) el[e.attrs.type] = dxfs[Number(e.attrs.dxfId)] ?? {};
    for (const e of kids(tsByName.get(ss.attrs.name), 'tableStyleElement')) el[e.attrs.type] = dxfs[Number(e.attrs.dxfId)] ?? {};
    const edge = (st) => st?.bbc ?? st?.btc ?? st?.blc ?? st?.brc;
    const c = {
      frame: el.wholeTable?.fill, border: edge(el.wholeTable), head: el.headerRow?.color,
      selFill: el.selectedItemWithData?.fill, selText: el.selectedItemWithData?.color, selBorder: edge(el.selectedItemWithData),
      item: el.unselectedItemWithData?.fill, itemText: el.unselectedItemWithData?.color, itemBorder: edge(el.unselectedItemWithData),
      noData: el.unselectedItemWithNoData?.color ?? el.selectedItemWithNoData?.color,
    };
    slicerStyles[ss.attrs.name] = Object.fromEntries(Object.entries(c).filter(([, v]) => v));
  }
  return { xfs, dxfs, filterDxfs, tableDxfs, dxfOf, defaultFont, tableStyles, wbFont, slicerStyles, cellStyles, fonts, objectStyles };
}

const CHUNK_MIN = 48 << 20;
const CHUNK = 8 << 20;
const asciiBytes = (t) => Uint8Array.from(t, (ch) => ch.charCodeAt(0));
const B_SD = asciiBytes('<sheetData');
const B_SD_END = asciiBytes('</sheetData>');
const B_ROW_END = asciiBytes('</row>');
/** 바이트 배열에서 패턴 찾기 */
function bytesIndexOf(u8, pat, from) {
  const first = pat[0];
  for (let i = u8.indexOf(first, from); i >= 0 && i <= u8.length - pat.length; i = u8.indexOf(first, i + 1)) {
    let k = 1;
    while (k < pat.length && u8[i + k] === pat[k]) k++;
    if (k === pat.length) return i;
  }
  return -1;
}
function bytesLastIndexOf(u8, pat) {
  for (let i = u8.lastIndexOf(pat[0]); i >= 0; i = i > 0 ? u8.lastIndexOf(pat[0], i - 1) : -1) {
    let k = 1;
    while (k < pat.length && u8[i + k] === pat[k]) k++;
    if (k === pat.length) return i;
  }
  return -1;
}
/**
 * 큰 시트 XML(바이트): sheetData 앞뒤만 글자로 바꾼 rest 와, 행 경계에서 자른 조각마다 scanRows 하는 rows()
 * 접두사 있는 형식 · 작은 파일은 null (일반 경로)
 */
function chunkedSheet(u8) {
  if (!u8 || u8.length < CHUNK_MIN) return null;
  const a = bytesIndexOf(u8, B_SD, 0);
  if (a < 0) return null;
  const head = textOf(u8.subarray(0, a));
  if (/<[A-Za-z_][\w.-]*:worksheet[\s>]/.test(head.slice(0, 2000))) return null;
  const tagEnd = u8.indexOf(62, a); // '>'
  if (u8[tagEnd - 1] === 47) return null; // '<sheetData/>'
  const b = bytesLastIndexOf(u8, B_SD_END);
  if (b < 0 || b < tagEnd) return null;
  const rest = `${head}<sheetData/>${textOf(u8.subarray(b + B_SD_END.length))}`;
  return {
    rest,
    *rows() {
      let p = tagEnd + 1;
      while (p < b) {
        let e = Math.min(b, p + CHUNK);
        if (e < b) {
          const re = bytesIndexOf(u8, B_ROW_END, e);
          e = re < 0 || re > b ? b : re + B_ROW_END.length;
        }
        const text = textOf(u8.subarray(p, e));
        yield* scanRows(text, 0, text.length);
        p = e;
      }
    },
  };
}

/** sheetData 부분을 떼어 냄 (접두사 없는 일반 형식일 때만) */
function splitSheetData(xml) {
  const a = xml.indexOf('<sheetData');
  if (a < 0 || /<[A-Za-z_][\w.-]*:worksheet[\s>]/.test(xml.slice(0, 2000))) return null;
  const tagEnd = xml.indexOf('>', a);
  if (xml[tagEnd - 1] === '/') return null;
  const b = xml.indexOf('</sheetData>', tagEnd);
  if (b < 0) return null;
  return { start: tagEnd + 1, end: b, rest: `${xml.slice(0, a)}<sheetData/>${xml.slice(b + 12)}` };
}

/** 태그 속성 문자열 → 객체 (xml[from..to) 범위) */
function scanAttrs(xml, from, to) {
  const attrs = {};
  let i = from;
  while (i < to) {
    const eq = xml.indexOf('=', i);
    if (eq < 0 || eq >= to) break;
    const name = xml.slice(i, eq).trim();
    let q = eq + 1;
    while (xml[q] === ' ') q++;
    const quote = xml[q];
    const close = xml.indexOf(quote, q + 1);
    if (close < 0) break;
    const v = xml.slice(q + 1, close);
    attrs[name] = v.includes('&') ? decodeEntities(v) : v;
    i = close + 1;
  }
  return attrs;
}

/** 셀 내용 문자열에서 <tag …>본문</tag> 찾기 */
function innerText(body, tag) {
  const open = body.indexOf(`<${tag}`);
  if (open < 0) return undefined;
  const next = body.charCodeAt(open + tag.length + 1);
  if (next !== 62 && next !== 32 && next !== 47) return undefined; // '>' ' ' '/'
  const gt = body.indexOf('>', open);
  if (body[gt - 1] === '/') return { attrs: scanAttrs(body, open + tag.length + 1, gt - 1), text: '' };
  const close = body.indexOf(`</${tag}>`, gt);
  const t = body.slice(gt + 1, close);
  return { attrs: scanAttrs(body, open + tag.length + 1, gt), text: unx(t.includes('&') || t.includes('\r') ? decodeEntities(t) : t) };
}

// 저장할 짧은 문자열이 수 MB의 원본 XML 조각을 계속 붙잡지 않게 분리합니다.
// JSON 왕복은 고립 서로게이트를 포함한 UTF-16 값을 그대로 보존합니다.
const detachedXmlText = value => value ? JSON.parse(JSON.stringify(value)) : value;

/** sheetData 의 행/셀을 DOM 없이 읽음 → { attrs, cells: [{ attrs, v, f, fa, is }] } */
function* scanRows(xml, start, end) {
  let p = start;
  while (p < end) {
    const ro = xml.indexOf('<row', p);
    if (ro < 0 || ro >= end) return;
    const rgt = xml.indexOf('>', ro);
    const selfClose = xml[rgt - 1] === '/';
    const attrs = scanAttrs(xml, ro + 4, selfClose ? rgt - 1 : rgt);
    const cells = [];
    if (selfClose) { p = rgt + 1; yield { attrs, cells }; continue; }
    const rEnd = xml.indexOf('</row>', rgt);
    if (rEnd < 0) return;
    const rowXml = xml.slice(rgt + 1, rEnd);
    let q = 0;
    const n = rowXml.length;
    while (q < n) {
      const co = rowXml.indexOf('<c', q);
      if (co < 0) break;
      const ch = rowXml.charCodeAt(co + 2);
      if (ch !== 32 && ch !== 62 && ch !== 47) { q = co + 2; continue; }
      const cgt = rowXml.indexOf('>', co);
      const cSelf = rowXml[cgt - 1] === '/';
      const cell = { attrs: scanAttrs(rowXml, co + 2, cSelf ? cgt - 1 : cgt), v: null, f: null, fa: null, is: null };
      if (cSelf) { cells.push(cell); q = cgt + 1; continue; }
      let cEnd = rowXml.indexOf('</c>', cgt);
      if (cEnd < 0) cEnd = n;
      const body = rowXml.slice(cgt + 1, cEnd);
      if (body) {
        const v = innerText(body, 'v');
        if (v) cell.v = v.text;
        const f = innerText(body, 'f');
        if (f) { cell.f = detachedXmlText(f.text); cell.fa = f.attrs; }
        const is = body.indexOf('<is>');
        if (is >= 0) cell.is = body.slice(is + 4, body.lastIndexOf('</is>'));
      }
      cells.push(cell);
      q = cEnd + 4;
    }
    p = rEnd + 6;
    yield { attrs, cells };
  }
}

/** 일반 XML 파서로 읽은 sheetData → scanRows 와 같은 모양 */
function* domRows(data) {
  for (const row of kids(data, 'row')) {
    yield {
      attrs: row.attrs,
      cells: kids(row, 'c').map((c) => {
        const v = child(c, 'v');
        const f = child(c, 'f');
        const is = child(c, 'is');
        return { attrs: c.attrs, v: v ? v.text : null, f: f ? f.text : null, fa: f ? f.attrs : null, is: is ? `<t>${esc(allText(is))}</t>` : null };
      }),
    };
  }
}

/** 수식 모양: 셀 참조의 행 번호를 # 로 (함수 이름 LOG10( 등은 그대로). 길이 제한 */
const SHAPE_RE = /(\$?\b[A-Za-z]{1,3}\$?)\d+(?![\d(.])/g;
function formulaShape(f) {
  return f.length > 2000 ? null : f.replace(SHAPE_RE, '$1#');
}
/** 모양의 # 자리에 수식 f 의 행 번호를 차례로 채움 (변환 결과 모양 → 이 칸의 변환 결과) */
function fillShape(shape, f) {
  const nums = [];
  f.replace(SHAPE_RE, (m, p) => { nums.push(m.slice(p.length)); return m; });
  let i = 0;
  const out = shape.replace(/(\$?\b[A-Za-z]{1,3}\$?)#/g, (m, p) => `${p}${nums[i++]}`);
  return i === nums.length ? out : null;
}

/** 파일 수식 → 앱 수식 본문 (_xlfn. 등 접두사 제거, SINGLE → @, ANCHORARRAY → #) */
function cleanFormula(f, opt = {}) {
  return fromFileFormula(f, opt);
}

function* streamedRows(stream) {
  for (const xml of stream.rows()) {
    if (/^<[\w.-]+:row[\s/>]|<\?|<!|<\/row\s+>/.test(xml)) yield* domRows(parseXml('<sheetData>'+xml+'</sheetData>'));
    else yield* scanRows(xml, 0, xml.length);
  }
}

function* readSheet(files, path, ctx) {
  // 셀 데이터(sheetData)는 빠른 전용 스캐너로, 나머지는 일반 XML 파서로 읽음
  const binRows = files.__xlsb?.rows.get(path); // xlsb: 셀은 바이너리에서 바로
  // 아주 큰 시트(수백 MB XML)는 한 번에 글자로 바꾸지 않고 행 경계에서 잘라 조금씩 (시간 · 메모리)
  const stream = !binRows && ctx.streamParts?.has(path) ? sheetXmlStream(zipEntryChunks(files, path)) : null;
  const big = !binRows && !stream ? chunkedSheet(files[path]) : null;
  const xmlText = big || binRows || stream ? null : textOf(files[path]);
  const sd = binRows || big || stream ? null : splitSheetData(xmlText);
  let root = parseXml(stream ? stream.head : big ? big.rest : binRows ? textOf(files[path]) : sd ? sd.rest : xmlText);
  const sheetRows = binRows ? binRows() : stream ? streamedRows(stream) : big ? big.rows() : sd ? scanRows(xmlText, sd.start, sd.end) : domRows(child(root, 'sheetData'));
  const sheet = {
    cells: new CellMap(), colWidths: {}, rowHeights: {}, merges: [], cond: [], colStyles: {}, rowStyles: {},
    hiddenRows: {}, hiddenCols: {}, rowManual: {}, freeze: { rows: 0, cols: 0 }, filter: null, charts: [], images: [], shapes: [], validations: [], slicers: [],
  };
  const { xfs, dxfs, strings } = ctx;
  // 시트 기본 행 높이 · 열 너비 (엑셀은 시트마다 다름: 기본 글꼴 9pt 면 행 12pt = 16px)
  const mdw = ctx.mdw ?? 7;
  const sfp = child(root, 'sheetFormatPr');
  const defRowH = sfp?.attrs.defaultRowHeight ? pt2px(Number(sfp.attrs.defaultRowHeight)) : DEFAULT_ROW_HEIGHT;
  const defColW = sfp?.attrs.defaultColWidth ? width2pxM(Number(sfp.attrs.defaultColWidth), mdw) : baseColPx(Number(sfp?.attrs.baseColWidth ?? 8), mdw);
  if (defRowH !== DEFAULT_ROW_HEIGHT) sheet.defRowH = defRowH;
  if (defColW !== DEFAULT_COL_WIDTH) sheet.defColW = defColW;
  // 서식 객체는 xf 번호마다 하나를 공유 (셀마다 복사하지 않음)
  const styleMemo = ctx.styleMemo ??= new Map();
  const styleOf = (s) => {
    const k = s || '0';
    let st = styleMemo.get(k);
    if (st === undefined) {
      const x = xfs[Number(k)];
      st = x && Object.keys(x).length ? x : null;
      styleMemo.set(k, st);
    }
    return st ?? undefined;
  };
  // 엑셀: <c> 요소가 있는 셀은 자기 xf 만 씀 (행 · 열 서식은 셀 요소가 없는 빈 칸에만).
  // 앱은 행 · 열 서식 위에 셀 서식을 겹치므로, 행 · 열에만 있는 속성(줄 바꿈 · 굵게 · 채우기 · 테두리…)은 기본값으로 막음
  const ownMemo = new Map();
  const lineSource = new WeakMap(); // default masks are not newly applied direct formatting
  const NONE = {};
  const neutral = (k, v) => (k === 'locked' ? true : k === 'decimals' ? null : k === 'size' ? ctx.wbFont.size : k === 'font' ? ctx.wbFont.name : typeof v === 'boolean' ? false : typeof v === 'number' ? 0 : '');
  const ownStyle = (st, r, cc) => {
    const row = sheet.rowStyles[r];
    const col = sheet.colStyles[cc];
    if (!row && !col) return st;
    const k1 = st ?? NONE;
    let m = ownMemo.get(k1);
    if (!m) { m = new Map(); ownMemo.set(k1, m); }
    let m2 = m.get(row ?? NONE);
    if (!m2) { m2 = new Map(); m.set(row ?? NONE, m2); }
    let out = m2.get(col ?? NONE);
    if (out === undefined) {
      out = st;
      for (const src of [row, col]) {
        if (!src) continue;
        for (const [key, v] of Object.entries(src)) {
          if (v === undefined || (out && out[key] !== undefined)) continue;
          if (out === st) out = { ...st };
          out[key] = neutral(key, v);
          // A row mask may originate in an unrelated column. Keep that new
          // neutral out of the direct-format layer of an already-cleared table.
          // Original row/column properties and explicit cell properties still win.
          if (st?.tableStyleInherit && TABLE_VISUAL_KEYS.includes(key)
            && (lineSource.get(row) ?? row)?.[key] === undefined
            && (lineSource.get(col) ?? col)?.[key] === undefined) {
            out.tableStyleInherit = { ...out.tableStyleInherit, [key]: out[key] };
          }
        }
      }
      m2.set(col ?? NONE, out ?? null);
    }
    return out ?? undefined;
  };

  // 행/열 xf는 완전한 서식이다. 저장을 줄이려고 생략한 기본값이 아래 서식에서
  // 다시 나타나지 않도록, 실제 아래 계층에 있는 속성만 기본값으로 막는다.
  const lineStyle = (st, under) => {
    let out=st;
    for(const [key,value] of Object.entries(under??{})) {
      if(value===undefined || st?.[key]!==undefined)continue;
      if(out===st)out={...st};out[key]=neutral(key,value);
    }
    if(out && out!==st)lineSource.set(out,st??NONE);
    return out;
  };
  const colStyleMemo=new Map(),rowStyleMemo=new Map(),baseStyle=styleOf('0');
  const colStyleOf=id=>{const k=String(id);if(!colStyleMemo.has(k))colStyleMemo.set(k,lineStyle(styleOf(k),baseStyle));return colStyleMemo.get(k);};
  for (const col of kids(child(root, 'cols'), 'col')) {
    const min = Number(col.attrs.min) - 1;
    const max = Math.min(Number(col.attrs.max) - 1, min + 16384);
    const w = col.attrs.width !== undefined ? width2pxM(Number(col.attrs.width), mdw) : null;
    const st = col.attrs.style !== undefined ? colStyleOf(col.attrs.style) : undefined;
    for (let c = min; c <= max && c < MAX_COLS; c++) {
      // width is authoritative even without customWidth; dropping 1px differences drifts drawing anchors.
      if (w !== null && w !== defColW) sheet.colWidths[c] = w;
      if (col.attrs.hidden === '1' || col.attrs.hidden === 'true') sheet.hiddenCols[c] = true;
      if (st) sheet.colStyles[c] = st;
      const ol = Number(col.attrs.outlineLevel ?? 0);
      if (ol > 0) ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).cols[c] = Math.min(7, ol));
      if (col.attrs.collapsed === '1' || col.attrs.collapsed === 'true') ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).colsColl[c] = true);
    }
  }

  const rowUnderlay={...baseStyle};
  for(const st of Object.values(sheet.colStyles))Object.assign(rowUnderlay,st);
  const rowStyleOf=id=>{const k=String(id);if(!rowStyleMemo.has(k))rowStyleMemo.set(k,lineStyle(styleOf(k),rowUnderlay));return rowStyleMemo.get(k);};

  const shared = {};
  const arrays = []; // 배열 수식 영역: 앵커 밖의 셀 값은 가져오지 않음 (다시 분산됨)
  let unsupported = 0;
  let rowIdx = -1;
  // 행이 아주 많은 시트: 값 셀은 열 블록(형식화 배열)으로 — 셀 객체 수백만 개를 만들지 않음
  const dimRef = refToRange(child(root, 'dimension')?.attrs.ref ?? '');
  const dimRows = dimRef ? dimRef.r2 - dimRef.r1 + 1 : 0;
  const dimCols = dimRef ? dimRef.c2 - dimRef.c1 + 1 : 0;
  // Wide data sheets can hold nearly a million cells below the row threshold.
  // Keep small/sparse-height sheets on their existing path; formulas and local
  // formatting overrides still use ordinary cells inside an eligible block.
  let blockMode = dimRows > BLOCK_MIN_ROWS || (dimRows >= 8192 && dimRows * dimCols >= 262144);
  let blockStart = -1;
  let blockLast = -1;
  let lateEmpty = null; // 열 → 연속 [시작 행, 끝 행, 서식]: 블록 끝이 정해질 때까지 보존
  const builders = [];
  const colFmt = [];
  const formulaMemo = ctx.formulaMemo ??= { legacy: new Map(), modern: new Map(), shape: new Map() }; // 같은 수식 문자열(표의 계산 열 등)은 한 번만 변환
  const textMemo = ctx.textMemo ??= new Map();
  let rowCount = 0, parsedCellsSinceYield = 0;
  let plainFormulaCells = 0, plainValueCells = 0, plainBlankCells = 0;
  const shareLiteral = createImportedLiteralMemo();
  const blankRuns = createImportedBlankRuns(sheet.cells);
  let previousRow = -1;
  // 검수 도구만 요청하는 수치입니다. 셀·행을 펼치거나 원문을 복사하지 않습니다.
  const reportStorage = () => {
    if (typeof ctx.onDiagnostics !== 'function') return;
    let plainPoints = 0, compressedPoints = 0, blankRuns = 0, blankOrderRuns = 0;
    let uniqueBlankRuns = 0, uniqueBlankOrderRuns = 0;
    const patterns = new WeakSet();
    for (const column of sheet.cells.cols.values()) {
      if (column.plain) plainPoints += column.plain.size;
      else {
        compressedPoints += column.points.size; blankRuns += column.runs.length; blankOrderRuns += column.order.length;
        if (!patterns.has(column.dataKey)) {
          patterns.add(column.dataKey); uniqueBlankRuns += column.runs.length; uniqueBlankOrderRuns += column.order.length;
        }
      }
    }
    ctx.onDiagnostics({ part: path, rows: rowCount, row: rowIdx, dimension: dimRef, blockMode, blockStart, blockLast,
      logicalCells: sheet.cells.size, plainPoints, compressedPoints, blankRuns, blankOrderRuns, uniqueBlankRuns, uniqueBlankOrderRuns, plainFormulaCells, plainValueCells, plainBlankCells,
      builders: builders.map((b, column) => b ? { column, capacity: b.cap, rows: b.n, numBytes: b.num?.byteLength ?? 0, strBytes: b.str?.byteLength ?? 0, dictionary: b.dict.length } : null).filter(Boolean) });
  };
  for (const row of sheetRows) {
    if (++rowCount % 1000 === 0 || parsedCellsSinceYield >= 32768) {
      if (rowCount % 10000 === 0) reportStorage();
      parsedCellsSinceYield = 0; yield rowCount;
    }
    parsedCellsSinceYield += row.cells.length;
    rowIdx = row.attrs.r ? Number(row.attrs.r) - 1 : rowIdx + 1;
    const r = rowIdx;
    if (r <= previousRow) blankRuns.flush();
    previousRow = r;
    if (row.attrs.ht && (row.attrs.customHeight === '1' || row.attrs.customHeight === 'true')) {
      const h = pt2px(Number(row.attrs.ht));
      sheet.rowHeights[r] = h;
      sheet.rowManual[r] = true; // 기본 높이와 같아도 사용자가 고정한 행
    } else if (row.attrs.ht) {
      const h = pt2px(Number(row.attrs.ht));
      if (h !== defRowH) sheet.rowHeights[r] = h;
    }
    const noHt = !row.attrs.ht; // 높이가 저장되지 않은 행: 엑셀은 내용(큰 글꼴 · 줄 바꿈 · 회전)에 맞춰 자동 높이
    if (row.attrs.hidden === '1' || row.attrs.hidden === 'true') sheet.hiddenRows[r] = true;
    // 개요 (행 그룹)
    if (row.attrs.outlineLevel && row.attrs.outlineLevel !== '0') ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).rows[r] = Math.min(7, Number(row.attrs.outlineLevel)));
    if (row.attrs.collapsed === '1' || row.attrs.collapsed === 'true') ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).rowsColl[r] = true);
    if (trueAttr(row.attrs.customFormat) && row.attrs.s !== undefined) { const st = rowStyleOf(row.attrs.s); if (st) sheet.rowStyles[r] = st; }
    let colIdx = -1;
    const rowKey = `${r},`;
    for (const c of row.cells) {
      let cc;
      if (c.cc !== undefined) cc = c.cc;
      else if (c.attrs.r) {
        const ref = c.attrs.r;
        let n = 0;
        for (let i = 0; i < ref.length; i++) {
          const ch = ref.charCodeAt(i) & ~32;
          if (ch < 65 || ch > 90) break;
          n = n * 26 + ch - 64;
        }
        cc = n - 1;
      } else cc = colIdx + 1;
      if (cc <= colIdx) blankRuns.flush();
      colIdx = cc;
      const t = c.attrs.t ?? 'n';
      const vText = c.v;
      const style = ownStyle(styleOf(c.attrs.s), r, cc);
      let raw = '';
      let value = null, phonetic;
      if (t === 's') phonetic = ctx.phonetics?.[Number(vText)];
      if (t === 's') value = strings[Number(vText)] ?? '';
      else if (t === 'inlineStr') {
        const inline = c.is !== null ? parseXml(`<is>${c.is}</is>`) : null;
        value = detachedXmlText(allText(inline)); phonetic = readPhonetic(inline, value, ctx.fonts);
      }
      else if (t === 'str') value = detachedXmlText(vText ?? '');
      else if (t === 'b') value = vText === '1';
      else if (t === 'e') value = { error: vText ?? '#N/A' };
      else if (vText !== null && vText !== '') value = Number(vText);
      if (phonetic || c.attrs.ph !== undefined) phonetic = normalizePhonetic({ ...phonetic, visible: c.attrs.ph === '1' || c.attrs.ph === 'true' }, value ?? '');

      let formula = null;
      const fa = c.fa;
      if (fa) {
        if (fa.t === 'dataTable') {
          (ctx.importWarningCodes ??= new Set()).add('dataTableValuesOnly');
          ctx.warnings.add(XLSX_IMPORT_WARNINGS.dataTableValuesOnly);
        } else if (fa.t === 'shared' && fa.si !== undefined) {
          if (c.f) shared[fa.si] = { text: c.f, r, c: cc };
          const m = shared[fa.si];
          // 공유 수식: 기준 수식을 한 번만 나눠 두고 칸마다 행 · 열만 옮김
          if (m) formula = m.r === r && m.c === cc ? m.text : (m.shift ??= formulaShifter(`=${m.text}`))(r - m.r, cc - m.c).slice(1);
        } else if (c.f) formula = c.f;
      }
      let cached, cachedArray;
      if (formula !== null) {
        const isArray = fa.t === 'array';
        if (isArray && fa.ref) {
          const rg = refToRange(fa.ref);
          if (rg && rg.r1 === r && rg.c1 === cc && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) {
            cachedArray = { h: rg.r2 - r + 1, w: rg.c2 - cc + 1, values: [] };
            if (vText !== null || t === 'inlineStr') cachedArray.values.push(0, 0, value);
            arrays.push({ ...rg, r, c: cc, cache: cachedArray });
          }
        }
        // 동적 배열(cm) 또는 배열 수식이 아니면 옛 형식 → 암시적 교차 '@'
        const legacy = !isArray && !c.attrs.cm;
        const memo = legacy ? formulaMemo.legacy : formulaMemo.modern;
        let conv = memo.get(formula);
        if (!conv) {
          // 채우기로 만든 수식은 행 번호만 다름: 같은 모양(행 번호를 뺀 글자)이 바꿀 것 없이 그대로였으면 해석하지 않음
          const shape = legacy ? formulaShape(formula) : null;
          const known = shape !== null ? formulaMemo.shape.get(shape) : undefined;
          // 변환이 행 번호와 무관하면(Sheet!#REF! → #REF! 등) 변환된 모양에 이 칸의 행 번호를 채움
          const filled = known?.out !== undefined ? fillShape(known.out, formula) : null;
          if (known !== undefined && known.out === undefined) conv = { raw: `=${formula}`, unknown: known.unknown };
          else if (filled !== null) conv = { raw: `=${filled}`, unknown: known.unknown };
          else {
            const f = cleanFormula(formula, { legacy, isName: ctx.isName, nameMulti: ctx.nameMulti });
            conv = { raw: `=${f}`, unknown: unknownFunctions(f, ctx.isName).length > 0 };
            if (shape !== null && !formulaMemo.shape.has(shape)) {
              if (f === formula) formulaMemo.shape.set(shape, { unknown: conv.unknown });
              else {
                const out = formulaShape(f);
                if (out !== null && fillShape(out, formula) === f) formulaMemo.shape.set(shape, { unknown: conv.unknown, out });
              }
            }
          }
          if (memo.size < 200000) memo.set(formula, conv);
        }
        raw = conv.raw;
        // 파일 저장값은 처음 열 때 보존합니다. 입력 변경 이후 캐시 신뢰성은 Workbook에서 판정합니다.
        cached = value;
        // 원본에서도 #NAME?였던 수식은 새 미지원 계산 결과로 오인하지 않는다.
        // 원문/오류는 그대로 보존하며 계산 상태 목록에서 '원본 오류'로 확인한다.
        if (conv.unknown && !hasSavedNameError(value)) unsupported++;
      } else if (arrays.length) {
        const owner = arrays.find(a => r >= a.r1 && r <= a.r2 && cc >= a.c1 && cc <= a.c2 && (r !== a.r || cc !== a.c));
        if (owner) {
          if (vText !== null || t === 'inlineStr') owner.cache.values.push(r - owner.r, cc - owner.c, value);
          value = null; // 결과는 앵커에 속하며 분산을 막는 입력 셀로 만들지 않습니다.
        }
      }
      if (!raw) {
        if (value === null) raw = '';
        else if (typeof value === 'boolean') raw = value ? 'TRUE' : 'FALSE';
        else if (typeof value === 'object') raw = value.error;
        else if (typeof value === 'string') {
          if (style?.numFmt === 'text') raw = value;
          else {
            raw = textMemo.get(value);
            if (raw === undefined) { raw = textRaw(value); if (textMemo.size < 200000) textMemo.set(value, raw); }
          }
        } else raw = numberRaw(value, style, ctx.date1904);
      }
      if (noHt && raw !== '' && style && ((style.size && style.size > ctx.wbFont.size) || style.wrap || style.rotate)) (sheet.fitRows ??= new Set()).add(r);
      if (noHt && phonetic?.visible && phonetic.runs.length) (sheet.fitRows ??= new Set()).add(r);
      if (blockMode && blockStart < 0) blockStart = r + 1; // 첫 행(머리글) 다음부터 블록
      // The newly enabled wide-sheet path keeps @ cells in their original
      // input representation; legacy blocks add an apostrophe to numeric text.
      const preserveTextInput = dimRows <= BLOCK_MIN_ROWS && style?.numFmt === 'text';
      if (blockMode && r >= blockStart && formula === null && !c.attrs.vm && !phonetic && !preserveTextInput) {
        if (colFmt[cc] === undefined) colFmt[cc] = style ?? null;
        if ((style ?? null) === colFmt[cc]) {
          if (value !== null && value !== '') {
            (builders[cc] ??= new ColBuilder(dimRef.r2 - blockStart + 1)).set(r - blockStart, value);
            if (r > blockLast) blockLast = r;
            continue;
          }
          // 빈 칸의 서식은 블록 열 서식으로 (블록 마지막 행보다 아래면 나중에 보통 셀로 되살림)
          if (value !== '') {
            if (style) {
              lateEmpty ??= new Map();
              let pending = lateEmpty.get(cc);
              if (!pending) { pending = { runs: [], ordered: true }; lateEmpty.set(cc, pending); }
              const a = pending.runs, end = a.length - 3;
              if (end >= 0 && r === a[end + 1] + 1 && style === a[end + 2]) a[end + 1] = r;
              else { if (end >= 0 && r <= a[end + 1]) pending.ordered = false; a.push(r, r, style); }
            }
            continue;
          } // 빈 글자 셀은 블록 대신 보통 셀로 (빈 칸과 구별)
        }
      }
      // 셀에 배치한 그림 (richData 값 메타데이터 vm)
      const cellImg = !formula && c.attrs.vm ? ctx.richImages?.[Number(c.attrs.vm)] : null;
      if (cellImg) raw = '';
      if (!raw && !style && !cellImg) continue;
      const d = { raw };
      if (phonetic && formula === null && typeof value === 'string') d.phonetic = phonetic;
      if (cellImg) d.image = { ...cellImg };
      if (style) d.style = style;
      if (cached !== undefined && cached !== null) d.cached = cached;
      if (cachedArray) d.cachedArray = cachedArray;
      if (formula !== null && style?.numFmt === 'text') d.fx = true; // 텍스트 서식 칸에 저장된 수식
      if (formula === null && style?.numFmt === 'text' && value !== null && typeof value !== 'string') d.inputType = 'value'; // 표시 형식 @인 숫자·논리·오류 값도 원래 자료형 유지
      if (raw === '' && style && !cellImg && !phonetic && formula === null) blankRuns.add(r, cc, style);
      else sheet.cells.setRC(r, cc, formula === null ? shareLiteral(d) : d);
      if (ctx.onDiagnostics) { if (formula !== null) plainFormulaCells++; else if (raw === '') plainBlankCells++; else plainValueCells++; }
    }
  }
  blankRuns.finish();
  if (stream) root = parseXml(stream.rest);
  // The area threshold also catches sparse dashboards with an inflated
  // dimension. A column block applies its format to every position, including
  // omitted cells. Only keep the new path if every formatted position existed
  // in the source; inspect allocated value arrays and stored cell/run spans,
  // never the full dimension rectangle.
  if (blockMode && dimRows <= BLOCK_MIN_ROWS && blockLast >= blockStart) {
    const n = blockLast - blockStart + 1;
    for (let c = 0; c < colFmt.length; c++) {
      if (!colFmt[c]) continue;
      const spans = [], b = builders[c];
      if (b) {
        let first = -1;
        for (let i = 0, end = Math.min(n, b.n); i <= end; i++) {
          const present = i < end && ((b.str && b.str[i] >= 0) || (b.num && Number.isFinite(b.num[i])));
          if (present && first < 0) first = i;
          else if (!present && first >= 0) { spans.push([blockStart + first, blockStart + i - 1]); first = -1; }
        }
      }
      for (const [r, cell, count] of sheet.cells.col(c)?.storageEntries() ?? []) {
        // Without an explicit style, styleAt would still inherit the block's
        // format. Such a cell therefore cannot cover a formatting gap.
        if (cell.style && r <= blockLast && r + count > blockStart) spans.push([Math.max(r, blockStart), Math.min(r + count - 1, blockLast)]);
      }
      const pending = lateEmpty?.get(c)?.runs ?? [];
      for (let i = 0; i < pending.length; i += 3) if (pending[i] <= blockLast && pending[i + 1] >= blockStart) spans.push([Math.max(pending[i], blockStart), Math.min(pending[i + 1], blockLast)]);
      spans.sort((a, b) => a[0] - b[0]);
      let next = blockStart;
      for (const [a, z] of spans) { if (a > next) break; next = Math.max(next, z + 1); }
      if (next <= blockLast) { blockMode = false; break; }
    }
    if (!blockMode) {
      // Restore only staged source values. The existing lateEmpty restoration
      // below restores explicit blank/style runs; absent cells remain absent.
      for (let c = 0; c < builders.length; c++) {
        const b = builders[c]; if (!b) continue;
        const style = colFmt[c] ?? undefined;
        for (let i = 0; i < b.n; i++) {
          const code = b.str?.[i], value = code >= 0 ? b.dict[code] : b.num?.[i];
          if (value === undefined || (typeof value === 'number' && !Number.isFinite(value))) continue;
          const raw = typeof value === 'number' ? numberRaw(value, style, ctx.date1904) : typeof value === 'string' ? textRaw(value) : typeof value === 'boolean' ? (value ? 'TRUE' : 'FALSE') : value.error;
          if (sheet.cells.hasRC(blockStart + i, c)) continue;
          sheet.cells.setRC(blockStart + i, c, shareLiteral({ raw, ...(style ? { style } : {}) }));
          if (ctx.onDiagnostics) plainValueCells++;
        }
      }
      builders.length = 0;
    }
  }
  reportStorage();
  sheet.unsupported = unsupported;
  if (lateEmpty) {
    const last = blockMode && blockLast >= blockStart ? blockLast : -1;
    for (const [column, pending] of lateEmpty) {
      const a = pending.runs;
      if (!pending.ordered) {
        // 정렬되지 않은/중복 행도 원래의 '먼저 기록된 빈 셀' 의미를 유지합니다.
        for (let i = 0; i < a.length; i += 3) for (let r = Math.max(last + 1, a[i]); r <= a[i + 1]; r++) {
          if (!sheet.cells.hasRC(r, column)) sheet.cells.setRC(r, column, { raw: '', style: a[i + 2] });
        }
        continue;
      }
      const occupied = [];
      for (const [r, , count] of sheet.cells.col(column)?.storageEntries() ?? []) if (r + count - 1 > last) occupied.push([r, r + count - 1]);
      occupied.sort((x, y) => x[0] - y[0]);
      let at = 0;
      for (let i = 0; i < a.length; i += 3) {
        let start = Math.max(last + 1, a[i]); const end = a[i + 1], value = { raw: '', style: a[i + 2] };
        while (at < occupied.length && occupied[at][1] < start) at++;
        for (let j = at; j < occupied.length && occupied[j][0] <= end; j++) {
          if (start < occupied[j][0]) sheet.cells.setRunRC(start, column, occupied[j][0] - start, value);
          start = Math.max(start, occupied[j][1] + 1);
        }
        if (start <= end) sheet.cells.setRunRC(start, column, end - start + 1, value);
      }
    }
    lateEmpty = null;
  }
  if (blockMode && blockLast >= blockStart) {
    const n = blockLast - blockStart + 1;
    const width = Math.max(builders.length, colFmt.length);
    sheet.blocks = [{
      r0: blockStart, c0: 0, n, ver: 0,
      cols: Array.from({ length: width }, (_, j) => (builders[j] ? builders[j].finish(n, colFmt[j] ?? null) : { num: null, str: null, dict: [], fmt: colFmt[j] ?? null })),
    }];
  }
  // 블록 칸에 메모 · 링크를 붙일 때 값을 잃지 않게
  const cellAt = (k) => {
    const hit = sheet.cells.get(k);
    if (hit || !sheet.blocks?.length) return hit;
    const [rr, cc2] = k.split(',').map(Number);
    const b = sheet.blocks[0];
    if (!inBlock(b, rr, cc2)) return undefined;
    const v = blockValue(b, rr, cc2);
    return v === null ? undefined : { raw: typeof v === 'number' ? numberRaw(v, b.cols[cc2].fmt, ctx.date1904) : typeof v === 'string' ? textRaw(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v.error, ...(b.cols[cc2].fmt ? { style: b.cols[cc2].fmt } : {}), ...(b.cols[cc2].fmt?.numFmt === 'text' ? { inputType: 'value' } : {}) };
  };

  for (const m of kids(child(root, 'mergeCells'), 'mergeCell')) {
    const rg = refToRange(m.attrs.ref);
    if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) sheet.merges.push(rg);
  }

  const sheetViews = kids(child(root, 'sheetViews'), 'sheetView');
  const sv0 = sheetViews.find((x) => x.attrs.workbookViewId === '0') ?? sheetViews[0];
  if (sv0 && (sv0.attrs.showGridLines === '0' || sv0.attrs.showGridLines === 'false')) sheet.noGrid = true;
  if (sv0 && (sv0.attrs.showZeros === '0' || sv0.attrs.showZeros === 'false')) sheet.noZeros = true;
  // 확대/축소 · 처음 보이는 칸 · 활성 셀 (엑셀에서 저장한 화면 그대로 열기)
  if (sv0?.attrs.zoomScale && Number(sv0.attrs.zoomScale) !== 100) sheet.zoom = Math.max(10, Math.min(400, Number(sv0.attrs.zoomScale)));
  const tlc = sv0?.attrs.topLeftCell ? refToRange(sv0.attrs.topLeftCell) : null;
  const pane = child(sv0, 'pane');
  const frozen = pane && (pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit');
  const split = (v, max) => Number.isFinite(Number(v)) ? Math.max(0, Math.min(max - 1, Math.floor(Number(v)))) : 0;
  const fr = frozen ? split(pane.attrs.ySplit, EXCEL_MAX_ROWS) : 0;
  const fc = frozen ? split(pane.attrs.xSplit, MAX_COLS) : 0;
  const ft = fr ? Math.min(tlc?.r1 ?? 0, EXCEL_MAX_ROWS - 1 - fr) : 0;
  const fl = fc ? Math.min(tlc?.c1 ?? 0, MAX_COLS - 1 - fc) : 0;
  const paneTlc = frozen && pane.attrs.topLeftCell ? refToRange(pane.attrs.topLeftCell) : null;
  const selections = kids(sv0, 'selection');
  const selEl = selections.find((x) => x.attrs.pane === pane?.attrs.activePane)
    ?? selections.find((x) => !x.attrs.pane) ?? selections[0];
  const act = selEl?.attrs.activeCell ? refToRange(selEl.attrs.activeCell) : null;
  if (fr || fc) {
    // sheetView는 고정창 원점, pane은 본문의 스크롤 위치다. 숨긴 행도 split 개수에 포함된다.
    sheet.freeze = { rows: fr, cols: fc, ...(ft ? { top: ft } : {}), ...(fl ? { left: fl } : {}) };
    const activePane = ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'].includes(pane.attrs.activePane) ? pane.attrs.activePane : fr && fc ? 'bottomRight' : fr ? 'bottomLeft' : 'topRight';
    sheet.view = {
      top: fr ? Math.max(ft + fr, paneTlc?.r1 ?? ft + fr) : paneTlc?.r1 ?? tlc?.r1 ?? 0,
      left: fc ? Math.max(fl + fc, paneTlc?.c1 ?? fl + fc) : paneTlc?.c1 ?? tlc?.c1 ?? 0,
      ...(act ? { r: act.r1, c: act.c1 } : {}), activePane,
    };
  } else if ((tlc && (tlc.r1 || tlc.c1)) || act) {
    sheet.view = { top: tlc?.r1 ?? 0, left: tlc?.c1 ?? 0, ...(act ? { r: act.r1, c: act.c1 } : {}) };
  }
  if (['normal', 'pageBreakPreview', 'pageLayout'].includes(sv0?.attrs.view)) sheet.view = { ...(sheet.view ?? {}), mode: sv0.attrs.view };
  for (const [attr, key] of [['showRowColHeaders', 'headers'], ['showFormulas', 'showFormulas']]) {
    if (sv0?.attrs[attr] !== undefined) sheet.view = { ...(sheet.view ?? {}), [key]: !falseAttr(sv0.attrs[attr]) };
  }
  if (sv0 && falseAttr(sv0.attrs.defaultGridColor) && sv0.attrs.colorId !== undefined) {
    // Excel은 눈금선에 사용자 지정 셀 팔레트가 아닌 고정 ICV 색을 사용한다.
    const index = Number(sv0.attrs.colorId), own = descendants(child(root, 'extLst'), 'gridColor').find(n => n.attrs['xmlns:wx'] === 'https://wixel.app/view/1' && Number(n.attrs.nativeId) === index);
    const gridColor = /^#[0-9a-f]{6}$/i.test(own?.attrs.rgb ?? '') ? own.attrs.rgb.toLowerCase() : INDEXED[index] ? `#${INDEXED[index].toLowerCase()}` : null;
    if (gridColor) sheet.view = { ...(sheet.view ?? {}), gridColor };
  }
  sheet.fileValues = true; // 셀의 파일 계산 결과를 그대로 씀 (바뀌기 전까지)

  const af = child(root, 'autoFilter');
  if (af?.attrs.ref) {
    const rg = refToRange(af.attrs.ref);
    if (rg) {
      const parsed = readAutoFilter(af, rg, { dxfs: ctx.filterDxfs ?? dxfs, sortNode: child(root, 'sortState') });
      const criteria = parsed.criteria;
      const hidden = {};
      for (const k of Object.keys(sheet.hiddenRows)) {
        const r = Number(k);
        if (r > rg.r1 && r <= rg.r2 && Object.keys(criteria).length) { hidden[r] = true; delete sheet.hiddenRows[r]; }
      }
      sheet.filter = { ...rg, ...parsed, hidden };
    }
  }

  // 조건부 서식 (priority 가 작을수록 먼저 적용 → 배열 앞쪽)
  const condList = [];
  const OPS = { greaterThan: 'gt', lessThan: 'lt', greaterThanOrEqual: 'ge', lessThanOrEqual: 'le', equal: 'eq', notEqual: 'ne', between: 'between', notBetween: 'notBetween' };
  const fval = (t) => {
    const x = String(t ?? '');
    if (/^".*"$/.test(x)) return x.slice(1, -1).replace(/""/g, '"');
    if (/^-?[\d.]+(E[+-]?\d+)?$/i.test(x)) return x;
    return x ? `=${cleanFormula(x)}` : '';
  };
  // 엑셀 2010 확장 규칙 (x14): 데이터 막대 세부 설정 · 새 아이콘 집합 · 확장 전용 규칙
  const x14Rules = new Map();
  const x14Only = [];
  const extLst = child(root, 'extLst');
  for (const x14cf of descendants(extLst, 'conditionalFormatting')) {
    const sqref = child(x14cf, 'sqref')?.text ?? x14cf.attrs.sqref ?? '';
    for (const xr of kids(x14cf, 'cfRule')) {
      if (xr.attrs.id) x14Rules.set(xr.attrs.id.toUpperCase(), xr);
      x14Only.push({ sqref, rule: xr });
    }
  }
  const readCfvo = (el) => kids(el, 'cfvo').map((v) => {
    const t = v.attrs.type;
    const val = v.attrs.val ?? child(v, 'f')?.text;
    const o = { type: t === 'num' ? 'num' : t };
    if (val !== undefined) o.v = t === 'formula' || (val && !/^-?[\d.]+(E[+-]?\d+)?$/i.test(val)) ? `=${cleanFormula(val)}` : val;
    if (v.attrs.gte === '0') o.gte = false;
    if (o.v && o.type === 'num' && String(o.v).startsWith('=')) o.type = 'formula';
    return o;
  });
  const ICONS_KNOWN = ['3Arrows', '3ArrowsGray', '3TrafficLights1', '3TrafficLights2', '3Symbols', '3Symbols2', '3Signs', '3Flags', '3Stars', '3Triangles',
    '4Arrows', '4ArrowsGray', '4RedToBlack', '4Rating', '4TrafficLights', '5Arrows', '5ArrowsGray', '5Rating', '5Quarters', '5Boxes'];
  const iconRule = (is) => {
    const name = is?.attrs.iconSet ?? '3TrafficLights1';
    const out = { type: 'icons', icons: ICONS_KNOWN.includes(name) ? name : name[0] === '5' ? '5Arrows' : name[0] === '4' ? '4Arrows' : '3TrafficLights1' };
    if (is?.attrs.reverse === '1') out.reverse = true;
    if (is?.attrs.showValue === '0') out.iconOnly = true;
    const cfvo = readCfvo(is);
    if (cfvo.length && !(cfvo.every((c, i) => c.type === 'percent' && Math.abs(Number(c.v) - Math.round((i * 100) / cfvo.length)) <= 1 && c.gte !== false))) out.cfvo = cfvo;
    return out;
  };
  // 서식(dxf)을 쓰는 규칙 — 기본 규칙과 확장(x14) 규칙이 함께 씀 (x14 는 텍스트를 수식에서 찾음)
  const simpleRule = (a, formulas, style) => {
    // 확장 규칙에서 텍스트가 셀 참조면(예: $A$1) 엑셀이 함께 적어 둔 판정 수식으로 계산
    if (a.text === undefined && formulas[0] && formulas[1] !== undefined && !/^".*"$/.test(formulas[1]) && ['containsText', 'notContainsText', 'beginsWith', 'endsWith'].includes(a.type)) {
      return { type: 'formula', formula: `=${cleanFormula(formulas[0])}`, style };
    }
    const txt = a.text ?? (formulas[1] !== undefined ? fval(formulas[1]) : '');
    switch (a.type) {
          case 'cellIs': return OPS[a.operator] ?  { type: OPS[a.operator], v1: fval(formulas[0]), ...(formulas[1] !== undefined ? { v2: fval(formulas[1]) } : {}), style } : null;
          case 'containsText': return { type: 'text', v1: txt, style };
          case 'notContainsText': return { type: 'notText', v1: txt, style };
          case 'beginsWith': return { type: 'begins', v1: txt, style };
          case 'endsWith': return { type: 'ends', v1: txt, style };
          case 'containsBlanks': return { type: 'blank', style };
          case 'notContainsBlanks': return { type: 'noBlank', style };
          case 'containsErrors': return { type: 'errors', style };
          case 'notContainsErrors': return { type: 'noErrors', style };
          case 'timePeriod': return { type: 'date', period: a.timePeriod ?? 'today', style };
          case 'expression': return formulas[0] ? { type: 'formula', formula: `=${cleanFormula(formulas[0])}`, style } : null;
          case 'duplicateValues': return { type: 'dup', style };
          case 'uniqueValues': return { type: 'unique', style };
          case 'top10': return { type: a.bottom === '1' ? 'bottom' : 'top', v1: a.rank ?? '10', ...(a.percent === '1' ? { percent: true } : {}), style };
          case 'aboveAverage': return { type: a.aboveAverage === '0' ? 'belowAvg' : 'aboveAvg', style, ...(a.equalAverage === '1' ? { equal: true } : {}), ...(a.stdDev ? { stdDev: Number(a.stdDev) } : {}) };
          default: return null;
    }
  };
  for (const cf of kids(root, 'conditionalFormatting')) {
    const ranges = (cf.attrs.sqref ?? '').split(/\s+/).map(refToRange).filter(Boolean);
    if (!ranges.length) continue;
    const rg = ranges[0];
    const more = ranges.length > 1 ? { more: ranges.slice(1) } : {};
    {
      for (const rule of kids(cf, 'cfRule')) {
        const a = rule.attrs;
        const style = dxfs[Number(a.dxfId)] ?? (a.dxfId === undefined ? {} : { fill: '#ffc7ce', color: '#9c0006' });
        const formulas = kids(rule, 'formula').map((f) => f.text);
        let out = simpleRule(a, formulas, style);
        switch (out ? '' : a.type) {
          case 'dataBar': {
            const db = child(rule, 'dataBar');
            const color = colorOf(child(db, 'color') ?? descendants(rule, 'color')[0], ctx.theme) ?? '#638ec6';
            out = { type: 'bar', color: color.toLowerCase(), ...(db?.attrs.showValue === '0' ? { iconOnly: true } : {}) };
            const cfvo = readCfvo(db);
            // 2010 확장: 단색 · 음수 막대 색 · 자동 최소/최대
            const extId = descendants(rule, 'id')[0]?.text?.toUpperCase();
            const xr = extId ? x14Rules.get(extId) : null;
            const xdb = xr ? child(xr, 'dataBar') : null;
            if (xdb) {
              if (xdb.attrs.gradient === '0') out.gradient = false;
              const neg = colorOf(child(xdb, 'negativeFillColor'), ctx.theme);
              if (neg) out.negColor = neg.toLowerCase();
              const xc = readCfvo(xdb);
              if (xc.length === 2) cfvo.splice(0, 2, ...xc);
            }
            if (cfvo.length === 2 && !(cfvo[0].type === 'min' && cfvo[1].type === 'max') && !(cfvo[0].type === 'autoMin' && cfvo[1].type === 'autoMax')) out.cfvo = cfvo;
            break;
          }
          case 'colorScale': {
            const cs = child(rule, 'colorScale');
            const colors = kids(cs, 'color').map((c) => colorOf(c, ctx.theme)).filter(Boolean);
            if (colors.length >= 2) {
              out = { type: 'scale', colors };
              const cfvo = readCfvo(cs);
              const std = colors.length === 3 ? ['min', 'percentile', 'max'] : ['min', 'max'];
              if (cfvo.length === colors.length && !(cfvo.every((c, i) => c.type === std[i]) && (colors.length === 2 || Number(cfvo[1].v) === 50))) out.cfvo = cfvo;
            }
            break;
          }
          case 'iconSet': {
            const extId = descendants(rule, 'id')[0]?.text?.toUpperCase();
            const xr = extId ? x14Rules.get(extId) : null;
            out = iconRule(child(xr, 'iconSet') ?? child(rule, 'iconSet'));
            break;
          }
          default:
        }
        if (out) {
          if (a.stopIfTrue === '1') out.stopIfTrue = true;
          condList.push({ p: Number(a.priority ?? 1e9), i: condList.length, rule: { ...rg, ...more, ...out } });
        }
      }
    }
  }
  // 확장(x14)에만 있는 규칙: 별·삼각형·상자 아이콘 집합 등
  const linked = new Set(descendants(root, 'cfRule').filter((r) => r.attrs.type === 'dataBar' || r.attrs.type === 'iconSet').flatMap((r) => descendants(r, 'id').map((x) => x.text?.toUpperCase())));
  for (const { sqref, rule } of x14Only) {
    if (linked.has(String(rule.attrs.id ?? '').toUpperCase())) continue;
    const ranges = sqref.split(/\s+/).map(refToRange).filter(Boolean);
    if (!ranges.length) continue;
    let out = null;
    if (rule.attrs.type === 'iconSet') out = iconRule(child(rule, 'iconSet'));
    else if (rule.attrs.type !== 'dataBar') {
      const dx = child(rule, 'dxf');
      const style = (dx && ctx.dxfOf ? ctx.dxfOf(dx) : null) ?? {};
      out = simpleRule(rule.attrs, kids(rule, 'f').map((f) => f.text), style);
      if (out && rule.attrs.stopIfTrue === '1') out.stopIfTrue = true;
    }
    else if (rule.attrs.type === 'dataBar') {
      const xdb = child(rule, 'dataBar');
      const color = colorOf(child(xdb, 'fillColor'), ctx.theme) ?? '#638ec6';
      out = { type: 'bar', color: color.toLowerCase(), ...(xdb?.attrs.gradient === '0' ? { gradient: false } : {}) };
    }
    if (out) condList.push({ p: Number(rule.attrs.priority ?? 1e9), i: condList.length, rule: { ...ranges[0], ...(ranges.length > 1 ? { more: ranges.slice(1) } : {}), ...out } });
  }
  condList.sort((x, y) => x.p - y.p || x.i - y.i);
  sheet.cond = condList.map((x) => x.rule);
  // 피벗 테이블 조건부 서식(<conditionalFormats priority>)과 짝지을 때 쓰는 파일의 priority (저장하지 않음)
  sheet._condPrio = condList.map((x) => x.p);

  // 메모 · 그림(차트)
  const rels = relsOf(files, path);
  for (const rel of Object.values(rels)) {
    if (rel.type === 'vmlDrawing' && files[rel.target]) {
      const vml = parseXml(textOf(files[rel.target]));
      for (const shape of descendants(vml, 'shape')) {
        const data = child(shape, 'ClientData');
        if (data?.attrs.ObjectType !== 'Note') continue;
        const row = child(data, 'Row')?.text.trim() ?? '', col = child(data, 'Column')?.text.trim() ?? '';
        if (!/^\d+$/.test(row) || !/^\d+$/.test(col)) continue;
        const r = Number(row), c = Number(col);
        if (!Number.isSafeInteger(r) || !Number.isSafeInteger(c) || r >= 1048576 || c >= 16384) continue;
        const visible = child(data, 'Visible');
        const shown = visible ? !/^(?:false|0)$/i.test(visible.text.trim()) : /visibility\s*:\s*visible/i.test(shape.attrs.style ?? '');
        if (shown) {
          sheet.noteVisibility ??= { states: {} }; sheet.noteVisibility.states[`${r},${c}`] = true;
        }
      }
    }
    if (rel.type === 'comments' && files[rel.target]) {
      const croot = parseXml(textOf(files[rel.target]));
      for (const cm of descendants(croot, 'comment')) {
        const m = /^([A-Z]+)(\d+)$/i.exec(cm.attrs.ref ?? '');
        if (!m) continue;
        const k = `${Number(m[2]) - 1},${nameToCol(m[1])}`;
        const text = allText(child(cm, 'text')).trim();
        if (!text) continue;
        sheet.cells.set(k, { raw: '', ...cellAt(k), comment: text });
      }
    }
  }
  // 하이퍼링크 (외부 주소는 관계 파일, 문서 안 위치는 location)
  for (const h of kids(child(root, 'hyperlinks'), 'hyperlink')) {
    const rg = refToRange(h.attrs.ref ?? '');
    if (!rg) continue;
    const rel = rid(h) ? rels[rid(h)] : null;
    const url = rel ? (rel.rawTarget ?? rel.target) : h.attrs.location ? `#${h.attrs.location}` : null;
    if (!url) continue;
    for (let r = rg.r1; r <= Math.min(rg.r2, rg.r1 + 999); r++) {
      for (let c = rg.c1; c <= rg.c2; c++) sheet.cells.set(`${r},${c}`, { raw: '', ...cellAt(`${r},${c}`), link: url });
    }
  }
  const drawing = child(root, 'drawing');
  if (drawing && rels[rid(drawing)]) Object.assign(sheet, readDrawing(files, rels[rid(drawing)].target, sheet, ctx));
  // 슬라이서 목록 · 피벗 테이블 (통합 문서 전체를 읽은 뒤 연결)
  sheet._slicers = [];
  sheet._pivots = [];
  for (const rel of Object.values(rels)) {
    if (rel.type === 'slicer' && files[rel.target]) {
      for (const sl of descendants(parseXml(textOf(files[rel.target])), 'slicer')) {
        if (sl.attrs.cache) {
          sheet._slicers.push({
            name: sl.attrs.name, cache: sl.attrs.cache, caption: sl.attrs.caption, columns: Number(sl.attrs.columnCount ?? 1), style: sl.attrs.style, startItem: slicerStartItem(sl.attrs.startItem),
            showCaption: !['0', 'false'].includes(sl.attrs.showCaption), lockedPosition: ['1', 'true'].includes(sl.attrs.lockedPosition), rowHeight: Number(sl.attrs.rowHeight ?? SLICER_DEFAULT_BUTTON_HEIGHT * EMU),
          });
        }
      }
    } else if (rel.type === 'pivotTable' && files[rel.target]) {
      const cacheRel = Object.values(relsOf(files, rel.target)).find((r) => r.type === 'pivotCacheDefinition');
      sheet._pivots.push({ root: parseXml(textOf(files[rel.target])), cachePath: cacheRel?.target });
    }
  }
  // 표 (ListObject)
  sheet.tables = [];
  for (const tp of kids(child(root, 'tableParts'), 'tablePart')) {
    const target = rels[rid(tp)]?.target;
    const tx = target && textOf(files[target]);
    if (!tx) continue;
    const t = readTable(parseXml(tx), sheet, ctx.objectStyles, ctx.filterDxfs ?? ctx.dxfs, ctx.tableDxfs ?? ctx.dxfs);
    if (t) sheet.tables.push(t);
  }
  sheet.validations = readValidations(root);
  sheet.codeName = child(root, 'sheetPr')?.attrs.codeName;
  // 시트 탭 색
  const tabC = colorOf(child(child(root, 'sheetPr'), 'tabColor'), ctx.theme);
  if (tabC) sheet.tabColor = tabC;
  // 스파크라인
  const sgs = descendants(root, 'sparklineGroup');
  if (sgs.length) {
    const col = (g, tag, dflt) => colorOf(child(g, tag), ctx.theme) ?? dflt;
    sheet.sparklines = sgs.map((g, gi) => {
      const a = g.attrs;
      const flag = (k) => a[k] === '1' || a[k] === 'true';
      const items = descendants(g, 'sparkline').map((spk) => {
        const f = (child(spk, 'f')?.text ?? '').trim();
        const at = refToRange((child(spk, 'sqref')?.text ?? '').trim());
        if (!f || !at) return null;
        const own = f.replace(/^'?(.*?)'?!/, (m, n) => (n.replace(/''/g, "'") === sheet.name ? '' : m));
        return { r: at.r1, c: at.c1, ref: own };
      }).filter(Boolean);
      return {
        id: `sp${gi}`, type: a.type === 'column' ? 'column' : a.type === 'stacked' ? 'winloss' : 'line',
        color: col(g, 'colorSeries', '#376092'), negColor: col(g, 'colorNegative', '#d00000'), markerColor: col(g, 'colorMarkers', '#d00000'),
        highColor: col(g, 'colorHigh', '#d00000'), lowColor: col(g, 'colorLow', '#d00000'), firstColor: col(g, 'colorFirst', '#d00000'), lastColor: col(g, 'colorLast', '#d00000'),
        markers: flag('markers'), high: flag('high'), low: flag('low'), first: flag('first'), last: flag('last'), negative: flag('negative'),
        weight: Number(a.lineWeight ?? 0.75), items,
      };
    }).filter((g) => g.items.length);
  }
  const pg = pageFromXml({ printOptions: child(root, 'printOptions'), pageMargins: child(root, 'pageMargins'), pageSetup: child(root, 'pageSetup'), headerFooter: child(root, 'headerFooter'), fitToPage: ['1', 'true'].includes(child(child(root, 'sheetPr'), 'pageSetUpPr')?.attrs.fitToPage) });
  if (pg) sheet.page = pg;
  for (const [tag, max] of [['rowBreaks', EXCEL_MAX_ROWS], ['colBreaks', MAX_COLS]]) {
    const breaks = validBreaks(kids(child(root, tag), 'brk').filter(b => ['1', 'true'].includes(b.attrs.man)).map(b => Number(b.attrs.id)), max);
    if (breaks.length) sheet.page = { ...(sheet.page ?? {}), [tag]: breaks };
  }
  const scn = child(root, 'scenarios');
  if (scn) {
    // 시나리오 관리자: <scenario name> + <inputCells r val>
    const list = kids(scn, 'scenario').map((x) => {
      const cells = kids(x, 'inputCells').map((ic) => { const p = parseRangeName(ic.attrs.r ?? ''); return p ? { r: p.r1, c: p.c1, v: ic.attrs.val ?? '' } : null; }).filter(Boolean);
      return cells.length ? { name: x.attrs.name ?? '', comment: x.attrs.comment ?? '', user: x.attrs.user ?? '', locked: x.attrs.locked !== '0', hidden: x.attrs.hidden === '1', cells: cells.map(({ r, c }) => ({ r, c })), values: cells.map((c) => c.v) } : null;
    }).filter(Boolean);
    if (list.length) sheet.scenarios = list;
  }
  const sp = child(root, 'sheetProtection');
  if (sp) { const p = protectFromAttrs(sp.attrs); if (p) sheet.protect = p; }
  const editable = kids(child(root, 'protectedRanges'), 'protectedRange').map(p => {
    const a = p.attrs, ranges = String(a.sqref ?? '').split(/\s+/).map(refToRange).filter(Boolean);
    return { name: a.name || '범위', ranges, ...(a.password ? { hash: a.password } : {}), ...(a.securityDescriptor ? { securityDescriptor: a.securityDescriptor } : {}),
      ...(a.algorithmName ? { modern: { algorithmName: a.algorithmName, hashValue: a.hashValue, saltValue: a.saltValue, spinCount: a.spinCount } } : {}) };
  }).filter(p => p.ranges.length);
  if (editable.length) sheet.protectedRanges = editable;
  const olp = child(child(root, 'sheetPr'), 'outlinePr');
  if (sheet.outline && olp) {
    if (olp.attrs.summaryBelow === '0' || olp.attrs.summaryBelow === 'false') sheet.outline.below = false;
    if (olp.attrs.summaryRight === '0' || olp.attrs.summaryRight === 'false') sheet.outline.right = false;
  }
  return sheet;
}

/** 시나리오 관리자 → <scenarios> */
function scenariosXml(list) {
  if (!list?.length) return '';
  const cells = [...new Set(list.flatMap((sc) => sc.cells.map((p) => cellName(p.r, p.c))))];
  return `<scenarios current="0" show="0" sqref="${cells.join(' ')}">${list.map((sc) => `<scenario name="${esc(sc.name)}"${sc.locked === false ? ' locked="0"' : ' locked="1"'}${sc.hidden ? ' hidden="1"' : ''} count="${sc.cells.length}"${sc.user ? ` user="${esc(sc.user)}"` : ''}${sc.comment ? ` comment="${esc(sc.comment)}"` : ''}>${sc.cells.map((p, i) => `<inputCells r="${cellName(p.r, p.c)}" val="${esc(String(sc.values[i] ?? ''))}"/>`).join('')}</scenario>`).join('')}</scenarios>`;
}

const DV_TYPES = new Set(['whole', 'decimal', 'list', 'date', 'time', 'textLength', 'custom']);

/** <dataValidations> (+ 다른 시트를 참조하는 x14 확장) → 규칙 목록 */
function readValidations(root) {
  const out = [];
  const add = (dv, sqref, f1, f2) => {
    const a = dv.attrs;
    const type = DV_TYPES.has(a.type) ? a.type : 'any';
    const base = {
      type, op: a.operator ?? 'between',
      allowBlank: a.allowBlank === '1' || a.allowBlank === 'true',
      showDropdown: !(a.showDropDown === '1' || a.showDropDown === 'true'), // 엑셀 속성은 '숨기기' 의미
      showError: a.showErrorMessage === '1' || a.showErrorMessage === 'true',
      showPrompt: a.showInputMessage === '1' || a.showInputMessage === 'true',
      errorStyle: a.errorStyle === 'warning' ? 'warning' : a.errorStyle === 'information' ? 'info' : 'stop',
    };
    if (f1 !== undefined && f1 !== '') {
      base.f1 = f1;
      // OOXML 목록 수식은 = 없이 저장된다. 이름·함수는 리터럴 목록과 구별해 둔다.
      if(type==='list'&&!f1.startsWith('"')) {try {if(!['ref','range','sref','spill'].includes(parse(f1.replace(/^=/,'')).type))base.f1='='+f1.replace(/^=/,'');}catch{base.f1='='+f1.replace(/^=/,'');}}
    }
    if (f2 !== undefined && f2 !== '') base.f2 = f2;
    for (const k of ['errorTitle', 'error', 'promptTitle', 'prompt']) if (a[k]) base[k] = a[k];
    if(VALIDATION_IME_MODES.includes(a.imeMode))base.imeMode=a.imeMode;
    let origin=null;
    for (const part of String(sqref ?? '').trim().split(/\s+/)) {
      if (!part) continue;
      const rg = /^[A-Z]+:[A-Z]+$/i.test(part) ? parseRangeName(`${part.split(':')[0]}1:${part.split(':')[1]}${MAX_ROWS}`) : refToRange(part);
      if (rg) {const range={...rg,r2:Math.min(rg.r2,MAX_ROWS-1)};origin??=range;out.push(relocateValidation({...origin,...base},range));}
    }
  };
  for (const dv of kids(child(root, 'dataValidations'), 'dataValidation')) {
    add(dv, dv.attrs.sqref, child(dv, 'formula1')?.text, child(dv, 'formula2')?.text);
  }
  const ext = descendants(child(root, 'extLst'), 'dataValidations');
  for (const group of ext) {
    for (const dv of kids(group, 'dataValidation')) {
      const f = (n) => { const el = child(dv, n); return el ? (child(el, 'f')?.text ?? el.text) : undefined; };
      add(dv, child(dv, 'sqref')?.text ?? dv.attrs.sqref, f('formula1'), f('formula2'));
    }
  }
  return out;
}

/** tables/tableN.xml → 표 모델 */
function readTable(root, sheet, styles, dxfs, tableDxfs) {
  const rg = refToRange(root.attrs.ref ?? '');
  if (!rg) return null;
  const header = root.attrs.headerRowCount !== '0';
  const totals = Number(root.attrs.totalsRowCount ?? 0) > 0;
  const info = child(root, 'tableStyleInfo');
  const cols = kids(child(root, 'tableColumns'), 'tableColumn');
  const name = (root.attrs.displayName || root.attrs.name || 'Table').replace(/\s/g, '_');
  const totalsFns = {}, totalsCells = {};
  cols.forEach((tc, i) => {
    const c = rg.c1 + i, f = tc.attrs.totalsRowFunction;
    if (f && f !== 'none' && f !== 'custom') totalsFns[c] = f;
    const formula = child(tc, 'totalsRowFormula');
    if (formula?.text) totalsCells[c] = { raw: `=${cleanFormula(formula.text)}` };
    else if (tc.attrs.totalsRowLabel !== undefined) totalsCells[c] = { raw: tc.attrs.totalsRowLabel, inputType: 'text' };
    const st = tc.attrs.totalsRowDxfId !== undefined ? tableDxfs?.[Number(tc.attrs.totalsRowDxfId)] : null;
    if (st && Object.keys(st).length) {
      const fn = TOTAL_FUNCS.find(x => x.id === f && x.code);
      const col = (tc.attrs.name ?? '').replace(/(['[\]#@])/g, "'$1");
      totalsCells[c] ??= { raw: fn ? `=SUBTOTAL(${fn.code},${name}[${col}])` : '' };
      totalsCells[c].style = structuredClone(st);
      if (st.numFmt === 'text' && totalsCells[c].raw.startsWith('=') && totalsCells[c].inputType !== 'text') totalsCells[c].fx = true;
    }
  });
  const af = child(root, 'autoFilter');
  let filter = null;
  if (af && header) {
    const parsed = readAutoFilter(af, rg, { dxfs, sortNode: child(root, 'sortState') });
    const criteria = parsed.criteria;
    const hidden = {};
    if (Object.keys(criteria).length) {
      for (const k of Object.keys(sheet.hiddenRows)) {
        const r = Number(k);
        if (r > rg.r1 && r <= rg.r2 - (totals ? 1 : 0)) { hidden[r] = true; delete sheet.hiddenRows[r]; }
      }
    }
    filter = { ...parsed, hidden };
  }
  return {
    id: `tb${Math.random().toString(36).slice(2, 9)}`, name, ...rg, r2: Math.max(rg.r2, rg.r1 + (header ? 1 : 0)),
    header, totals, style: info ? info.attrs.name ?? 'None' : DEFAULT_TABLE_STYLE,
    ...(findObjectStyle(styles,'table',info?.attrs.name) ? objectStylePatch('table',findObjectStyle(styles,'table',info.attrs.name)) : {}),
    banded: info ? info.attrs.showRowStripes !== '0' : true, bandedCols: info?.attrs.showColumnStripes === '1',
    firstCol: info?.attrs.showFirstColumn === '1', lastCol: info?.attrs.showLastColumn === '1',
    filter, ...(!filter && child(root, 'sortState') ? { sort: readFilterSort(child(root, 'sortState'), dxfs) } : {}), totalsFns, ...(!totals && (Object.keys(totalsCells).length || !falseAttr(root.attrs.totalsRowShown)) ? { totalsCells } : {}), ...(header ? {} : { columns: cols.map((c) => c.attrs.name ?? '') }),
    _xmlId: Number(root.attrs.id), _colNames: cols.map((c) => c.attrs.name ?? ''),
  };
}

export function numberRaw(v, style, date1904 = false) {
  const fmt = style?.numFmt;
  // 엑셀 1900 날짜 체계 (60 = 없는 날 1900-02-29 는 숫자 그대로)
  if ((fmt === 'date' || fmt === 'longdate') && Number.isInteger(v) && (date1904 ? v >= 0 : v > 0 && v !== 60)) {
    const d = dateParts(v, 1, date1904);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  if (fmt === 'percent') {
    const s = `${Number((v * 100).toPrecision(15))}%`;
    const p = parseInput(s);
    // Display-friendly input must preserve the exact stored scalar on reparse.
    if (typeof p.value === 'number' && Object.is(p.value, v)) return s;
  }
  return String(v);
}

const SCHEME_INDEX = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9 };
// 엑셀 도형 이름 → 도형 종류 (모양을 아는 도형은 이름 그대로)
const prstKind = (prst) => {
  if (GEOM[prst]) return prst;
  if (prst === 'line' || prst === 'straightConnector1') return 'line';
  if (/^bentConnector/.test(prst)) return 'bentConnector3';
  if (/^curvedConnector/.test(prst)) return 'curvedConnector3';
  if (/^flowChart/.test(prst)) return 'flowChartProcess';
  if (/Callout/.test(prst)) return 'wedgeRectCallout';
  return 'rect';
};
// 연결된 그림 (카메라): WIXEL 확장 uri · 1×1 투명 PNG (스냅숏이 없을 때 자리만)
const LINKED_PIC_URI = '{2C7E4B19-5A3D-4F6E-8B21-57495845434C}';
const BLANK_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';
const linkedRef = (L) => `${/^[A-Za-z_][\w.]*$/.test(L.sheet) ? L.sheet : `'${L.sheet.replace(/'/g, "''")}'`}!${colToName(L.c1)}${L.r1 + 1}:${colToName(L.c2)}${L.r2 + 1}`;
function parseLinkedRef(t) {
  const m = /^(?:'((?:[^']|'')+)'|([^'!]+))!\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(String(t).trim());
  if (!m) return null;
  const col = (x) => x.toUpperCase().split('').reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
  const r1 = Number(m[4]) - 1;
  const c1 = col(m[3]);
  return { sheet: (m[1] ?? m[2]).replace(/''/g, "'"), r1, c1, r2: m[6] ? Number(m[6]) - 1 : r1, c2: m[5] ? col(m[5]) : c1 };
}
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', emf: 'image/x-emf' };

/** DrawingML 색 (srgbClr / schemeClr / sysClr) */
function dmlColor(el, theme) {
  if (!el) return null;
  const c = el.children.find((x) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(x.name));
  if (!c) return null;
  let hex = null;
  if (c.name === 'srgbClr') hex = c.attrs.val;
  else if (c.name === 'sysClr') hex = c.attrs.lastClr ?? (c.attrs.val === 'window' ? 'FFFFFF' : '000000');
  else if (c.name === 'schemeClr') hex = theme[SCHEME_INDEX[c.attrs.val] ?? 4];
  else if (c.name === 'prstClr') hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', blue: '0000FF', green: '00FF00', yellow: 'FFFF00' }[c.attrs.val] ?? '000000';
  if (!hex) return null;
  const lum = child(c, 'lumMod');
  const off = child(c, 'lumOff');
  const shade = child(c, 'shade');
  const tint = child(c, 'tint');
  let t = 0;
  if (lum && off) t = Number(off.attrs.val) / 100000;
  else if (lum) t = Number(lum.attrs.val) / 100000 - 1;
  else if (shade) t = Number(shade.attrs.val) / 100000 - 1;
  else if (tint) t = 1 - Number(tint.attrs.val) / 100000;
  return `#${applyTint(hex.toUpperCase(), t).toLowerCase()}`;
}

/** DrawingML 색의 알파 변환 (100000=불투명). 색상 해석과 별도로 유지. */
function dmlOpacity(el) {
  const color = el?.children.find((x) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(x.name));
  let opacity = 1;
  for (const t of color?.children ?? []) {
    const n = Number(t.attrs.val) / 100000;
    if (!Number.isFinite(n)) continue;
    if (t.name === 'alpha') opacity = n;
    else if (t.name === 'alphaMod') opacity *= n;
    else if (t.name === 'alphaOff') opacity += n;
  }
  return Math.max(0, Math.min(1, opacity));
}

/** 개체 좌표용 행 · 열 축 (시트 기본 크기 + 숨긴 행 · 열, 필터로 숨긴 행 포함 — 화면과 같음) */
function objectAxes(sheet) {
  const hiddenRows = [sheet.hiddenRows, sheet.filter?.hidden, ...(sheet.tables ?? []).map((t) => t.filter?.hidden)];
  return {
    colAxis: new Axis(sheet.defColW ?? DEFAULT_COL_WIDTH, sheet.colWidths, [sheet.hiddenCols], MAX_COLS),
    rowAxis: new Axis(sheet.defRowH ?? DEFAULT_ROW_HEIGHT, sheet.rowHeights, hiddenRows, MAX_ROWS),
  };
}

function readDrawing(files, path, sheet, ctx) {
  const out = { charts: [], images: [], shapes: [], _slicerBoxes: {} };
  let z = 0; // 겹치는 순서
  const xml = textOf(files[path]);
  if (!xml) return out;
  const root = parseXml(xml);
  const rels = relsOf(files, path);
  // 개체 위치는 화면 좌표(숨긴 행 · 열은 높이 0) — 엑셀처럼 숨긴 행 아래의 개체도 위로 올라옴
  const { colAxis, rowAxis } = objectAxes(sheet);
  const point = (el) => {
    const n = (name) => Number(child(el, name)?.text ?? 0);
    // 엑셀: 칸 안의 오프셋은 그 칸의 너비 · 높이를 넘지 않음 (좁은 열에 큰 colOff 가 있어도 다음 열로 넘어가지 않음)
    const c = n('col');
    const r = n('row');
    const cw = colAxis.pos(c + 1) - colAxis.pos(c);
    const rh = rowAxis.pos(r + 1) - rowAxis.pos(r);
    return { x: colAxis.pos(c) + Math.min(n('colOff') / EMU, cw), y: rowAxis.pos(r) + Math.min(n('rowOff') / EMU, rh) };
  };
  const uid = (p) => `${p}${Date.now().toString(36)}${Math.random().toString(36).slice(2, 6)}`;
  const round = (b) => ({ x: Math.round(b.x), y: Math.round(b.y), w: Math.max(1, Math.round(b.w)), h: Math.max(1, Math.round(b.h)) });

  // 도형 한 개 → 모델 (box: 시트 좌표 px)
  const readShape = (el, box) => {
    const spPr = child(el, 'spPr');
    const prst = descendants(spPr, 'prstGeom')[0]?.attrs.prst ?? 'rect';
    const isText = descendants(child(el, 'nvSpPr'), 'cNvSpPr')[0]?.attrs.txBox === '1';
    const style = child(el, 'style');
    let fill = null;
    if (child(spPr, 'solidFill')) fill = dmlColor(child(spPr, 'solidFill'), ctx.theme);
    else if (!child(spPr, 'noFill') && style && !isText && el.name === 'sp') fill = dmlColor(child(style, 'fillRef'), ctx.theme);
    const ln = child(spPr, 'ln');
    let stroke = null;
    if (ln && child(ln, 'solidFill')) stroke = dmlColor(child(ln, 'solidFill'), ctx.theme);
    else if (!(ln && child(ln, 'noFill')) && style) stroke = dmlColor(child(style, 'lnRef'), ctx.theme);
    if (isText && !ln) stroke = null;
    const tx = child(el, 'txBody');
    const paras = kids(tx, 'p');
    // 목록 기본/단계 → 문단 기본 → 조각 서식을 속성별로 상속한다.
    const listStyle = child(tx, 'lstStyle');
    const defaultPara = child(listStyle, 'defPPr');
    const lvl = child(listStyle, 'lvl1pPr');
    const runStyle = (pr, pPr, level) => {
      const props = [child(defaultPara, 'defRPr'), child(level, 'defRPr'), child(pPr, 'defRPr'), pr].filter(Boolean);
      const attrs = Object.assign({}, ...props.map(p => p.attrs));
      const out = { b: ['1', 'true'].includes(attrs.b), i: ['1', 'true'].includes(attrs.i),
        u: !!attrs.u && attrs.u !== 'none', s: !!attrs.strike && attrs.strike !== 'noStrike' };
      if (attrs.sz) out.sz = Number(attrs.sz) / 100;
      for (let i = props.length - 1; i >= 0; i--) {
        if (!out.color) out.color = dmlColor(child(props[i], 'solidFill'), ctx.theme) || undefined;
        if (!out.font) {
          const face = child(props[i], 'ea')?.attrs.typeface || child(props[i], 'latin')?.attrs.typeface;
          if (face && !face.startsWith('+')) out.font = face;
        }
      }
      if (!out.color) delete out.color;
      return out;
    };
    const rich = paras.map((p) => {
      const pPr = child(p, 'pPr');
      const level = child(listStyle, `lvl${Math.max(0, Math.min(8, Number(pPr?.attrs.lvl) || 0)) + 1}pPr`);
      const al = pPr?.attrs.algn ?? level?.attrs.algn ?? defaultPara?.attrs.algn;
      const runs = [];
      for (const r of p.children) {
        if (r.name === 'r' || r.name === 'fld') runs.push({ t: child(r, 't')?.text ?? '', ...runStyle(child(r, 'rPr'), pPr, level) });
        else if (r.name === 'br') runs.push({ t: '\n', ...runStyle(child(r, 'rPr'), pPr, level) });
      }
      const para = { runs };
      if (al) para.align = al === 'ctr' ? 'center' : al === 'r' ? 'right' : al === 'just' || al === 'dist' ? 'justify' : 'left';
      if (!runs.length) { const endStyle = runStyle(child(p, 'endParaRPr'), pPr, level); if (endStyle.sz) para.sz = endStyle.sz; }
      return para;
    });
    // a:br 은 문단 안 줄바꿈이다. 편집용 평문에도 보존하여 편집 시 줄이 합쳐지지 않게 한다.
    const text = rich.map(p => p.runs.map(r => r.t).join('')).join('\n');
    const rPr = descendants(child(el, 'txBody'), 'rPr')[0] ?? descendants(child(el, 'txBody'), 'defRPr')[0] ?? descendants(child(el, 'txBody'), 'endParaRPr')[0];
    // 문단에 algn 이 없으면 목록 스타일, 그것도 없으면 DrawingML 기본값(왼쪽)
    const algn = descendants(child(el, 'txBody'), 'pPr')[0]?.attrs.algn ?? lvl?.attrs.algn;
    const shape = {
      id: uid('sh'), kind: isText ? 'textbox' : prstKind(prst), ...round(box), z: ++z,
      fill, stroke, text,
    };
    if (prst === 'roundRect') {
      const guide = kids(child(child(spPr, 'prstGeom'), 'avLst'), 'gd').find(g => g.attrs.name === 'adj');
      const match = /^val\s+(-?\d+)$/.exec(guide?.attrs.fmla ?? '');
      if (match && Number.isSafeInteger(Number(match[1]))) shape.adjustments = { adj: Number(match[1]) };
    }
    const custom = child(spPr, 'custGeom');
    if (custom) {
      const path = readCustomGeometry(custom, box.w, box.h);
      if (path) { shape.path = path; shape.kind = path.paths.some(p => p.commands.some(c => c[0] === 'C' || c[0] === 'Q')) ? 'curve' : path.paths.every(p => p.fill === false) ? 'scribble' : 'freeform'; }
      else { shape.customGeometry = custom; ctx.warnings.add('일부 사용자 지정 도형의 경로 수식은 표시하지 못합니다. 원래 경로는 XLSX 저장 시 보존합니다.'); }
    }
    const nonVisual = child(el, 'nvSpPr') ?? child(el, 'nvCxnSpPr');
    const hyperlink = readDrawingHyperlink(child(nonVisual, 'cNvPr'), rels);
    if (hyperlink) shape.hyperlink = hyperlink;
    const shapeName = descendants(nonVisual, 'cNvPr')[0]?.attrs.name;
    if (['1','true'].includes(descendants(nonVisual, 'cNvPr')[0]?.attrs.hidden)) shape.hidden = true;
    if (shapeName) shape.name = shapeName;
    const alt = descendants(nonVisual, 'cNvPr')[0]?.attrs.descr;
    if (alt !== undefined) shape.alt = alt;
    const lock = descendants(nonVisual, 'spLocks')[0] ?? descendants(nonVisual, 'cxnSpLocks')[0];
    if (lock?.attrs.noChangeAspect !== undefined) shape.lockAspect = lock.attrs.noChangeAspect === '1' || lock.attrs.noChangeAspect === 'true';
    const pattern = child(spPr, 'pattFill');
    if (pattern) {
      shape.pattern = { preset: pattern.attrs.prst ?? 'pct5', fg: dmlColor(child(pattern, 'fgClr'), ctx.theme) ?? '#000000', bg: dmlColor(child(pattern, 'bgClr'), ctx.theme) ?? '#ffffff' };
      const alpha = dmlOpacity(child(pattern, 'fgClr')); if (alpha !== 1 && alpha === dmlOpacity(child(pattern, 'bgClr'))) shape.fillOpacity = alpha;
    }
    const grad = child(spPr, 'gradFill');
    if (grad) {
      const stops = kids(child(grad, 'gsLst'), 'gs').map((gs) => [Number(gs.attrs.pos ?? 0) / 100000, dmlColor(gs, ctx.theme), dmlOpacity(gs)]).filter((s) => s[1]);
      if (stops.length) {
        const equalAlpha = stops.every((s) => s[2] === stops[0][2]);
        shape.grad = { ang: Number(child(grad, 'lin')?.attrs.ang ?? 5400000) / 60000, stops: stops.map((s) => equalAlpha ? s.slice(0, 2) : s) };
        shape.fill = stops[0][1];
        if (equalAlpha && stops[0][2] !== 1) shape.fillOpacity = stops[0][2];
        if (child(grad, 'path')) ctx.warnings.add('일부 도형의 경로 그라데이션은 선형 그라데이션으로 표시합니다.');
      }
    } else if (fill) {
      const opacity = dmlOpacity(child(spPr, 'solidFill'));
      if (opacity !== 1) shape.fillOpacity = opacity;
    }
    if (stroke) { const opacity = dmlOpacity(child(ln, 'solidFill')); if (opacity !== 1) shape.strokeOpacity = opacity; }
    const effects = child(spPr, 'effectLst');
    const shadow = child(effects, 'outerShdw');
    if (shadow) {
      const angle = Number(shadow.attrs.dir ?? 0) / 60000 * Math.PI / 180;
      const dist = Number(shadow.attrs.dist ?? 0) / EMU;
      const rounded = (n) => Math.round(n * 10000) / 10000;
      shape.shadow = { dx: rounded(Math.cos(angle) * dist), dy: rounded(Math.sin(angle) * dist), blur: Number(shadow.attrs.blurRad ?? 0) / EMU, color: dmlColor(shadow, ctx.theme) ?? '#000000', opacity: dmlOpacity(shadow) };
    }
    const glow = child(effects, 'glow');
    if (glow) shape.glow = { size: Number(glow.attrs.rad ?? 0) / EMU, color: dmlColor(glow, ctx.theme) ?? '#4472c4', opacity: dmlOpacity(glow) };
    const soft = child(effects, 'softEdge');
    if (soft) shape.soft = Number(soft.attrs.rad ?? 0) / EMU;
    const xf = descendants(spPr, 'xfrm')[0];
    if (xf?.attrs.flipH === '1') shape.flip = true;
    if (xf?.attrs.flipV === '1') shape.flipV = true;
    if (Number(xf?.attrs.rot)) shape.rot = Math.round(Number(xf.attrs.rot) / 60000);
    // 선 끝 화살표 · 대시
    for (const name of ['headEnd', 'tailEnd']) { const end = child(ln, name); if (end) shape[name] = { type: end.attrs.type ?? 'none', w: end.attrs.w ?? 'med', len: end.attrs.len ?? 'med' }; }
    const endOn = n => shape[n]?.type && shape[n].type !== 'none';
    if (endOn('tailEnd') || endOn('headEnd')) shape.arrow = endOn('tailEnd') && endOn('headEnd') ? 'both' : endOn('headEnd') ? 'start' : 'end';
    if (ln?.attrs.cap) shape.lineCap = ln.attrs.cap;
    if (ln?.attrs.cmpd) shape.compound = ln.attrs.cmpd;
    if (child(ln, 'round')) shape.lineJoin = 'round';
    if (child(ln, 'bevel')) shape.lineJoin = 'bevel';
    if (child(ln, 'miter')) shape.lineJoin = 'miter';
    const dashV = child(ln, 'prstDash')?.attrs.val;
    if (dashV && dashV !== 'solid') shape.dash = dashV === 'sysDot' ? 'dot' : dashV;
    const lw = Number(ln?.attrs.w);
    if (lw && stroke) shape.strokeWidth = Math.round((lw / EMU) * 4) / 4;
    const firstRun = rich.find(p => p.runs.length)?.runs[0] ?? runStyle(rPr, child(paras[0], 'pPr'), lvl);
    if (firstRun.sz) shape.size = firstRun.sz;
    if (firstRun.b) shape.bold = true;
    if (firstRun.i) shape.italic = true;
    if (firstRun.u) shape.underline = true;
    if (firstRun.s) shape.strike = true;
    if (firstRun.font) shape.font = firstRun.font;
    if (firstRun.color) shape.color = firstRun.color;
    else if (!isText && fill) shape.color = '#ffffff';
    shape.align = rich[0]?.align ?? (algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' || algn === 'dist' ? 'justify' : 'left');
    // 줄바꿈 및 빈 문단 크기도 조각의 일부다. 생략하면 편집/재저장 때 구조가 달라진다.
    const flat = rich.flatMap((p) => p.runs);
    const same = (k) => flat.every((r) => r[k] === flat[0]?.[k]);
    if (rich.length && (!['b', 'i', 'u', 's', 'sz', 'color', 'font'].every(same) ||
      rich.some(p => (p.align ?? shape.align) !== shape.align || p.sz !== undefined && p.sz !== shape.size || p.runs.some(r => r.t.includes('\n'))))) {
      shape.paras = rich;
      for (const k of ['bold', 'italic', 'underline', 'strike', 'font', 'size', 'color']) delete shape[k];
    }
    const body = child(tx, 'bodyPr');
    const anc = body?.attrs.anchor;
    shape.valign = anc === 'ctr' ? 'middle' : anc === 'b' ? 'bottom' : anc === 't' ? 'top' : isText ? 'top' : 'middle';
    const ins = (k, d) => (body?.attrs[k] !== undefined ? Number(body.attrs[k]) / EMU : d);
    const pad = [ins('tIns', 4.8), ins('rIns', 9.6), ins('bIns', 4.8), ins('lIns', 9.6)].map((v) => Math.round(v * 10) / 10);
    if (pad.join() !== '4.8,9.6,4.8,9.6') shape.pad = pad;
    if (body?.attrs.wrap === 'none') shape.nowrap = true;
    if (body?.attrs.rot !== undefined) shape.textRot = ((Number(body.attrs.rot) / 60000 % 360) + 360) % 360;
    if (child(body, 'normAutofit')) shape.textFit = 'shrink';
    else if (child(body, 'noAutofit')) shape.textFit = 'none';
    // 도형에 연결된 매크로 ([0]!이름) — WIXEL 은 VBA 를 실행하지 않고, 같은 이름의 내장 동작이 있으면 그것을 실행
    if (el.attrs.macro) shape.macro = el.attrs.macro.replace(/^\[\d+\]!/, '');
    out.shapes.push(shape);
  };

  const picSrc = new Map();
  const readPic = (el, box) => {
    const blip = descendants(el, 'blip')[0];
    const rel = blip && rels[rid(blip) ?? blip.attrs['r:embed']] ;
    const embed = blip && Object.keys(blip.attrs).find((k) => k.endsWith(':embed') || k === 'embed');
    const target = rels[blip?.attrs[embed]]?.target ?? rel?.target;
    const bytes = target && files[target];
    if (!bytes) return;
    const ext = target.split('.').pop().toLowerCase();
    const mime = MIME[ext];
    // EMF: 브라우저가 못 그리므로 SVG 로 바꿔 보여 주고, 저장할 때는 원본 EMF 를 씀
    const emf = (ext === 'emf' || !mime && isEmf(bytes)) && `data:image/x-emf;base64,${toBase64(bytes)}`;
    const src = picSrc.get(target) ?? (emf ? emfDataUrl(bytes) : mime && `data:${mime};base64,${toBase64(bytes)}`);
    if (!src) { ctx.warnings.add(`지원하지 않는 그림 형식(${ext})은 가져오지 않았습니다.`); return; }
    const name = descendants(child(el, 'nvPicPr'), 'cNvPr')[0]?.attrs.name ?? '그림';
    picSrc.set(target, src); // 같은 그림을 여러 번 쓰면 한 번만 변환
    const im = { id: uid('im'), name, ...round(box), z: ++z, src };
    const nv = descendants(child(el, 'nvPicPr'), 'cNvPr')[0];
    const hyperlink = readDrawingHyperlink(nv, rels);
    if (hyperlink) im.hyperlink = hyperlink;
    if (nv?.attrs.descr !== undefined) im.alt = nv.attrs.descr;
    if (['1','true'].includes(nv?.attrs.hidden)) im.hidden = true;
    const locks = descendants(child(el, 'nvPicPr'), 'picLocks')[0];
    im.lockAspect = locks?.attrs.noChangeAspect === '1' || locks?.attrs.noChangeAspect === 'true';
    const xf = child(child(el, 'spPr'), 'xfrm');
    if (Number(xf?.attrs.rot)) im.rot = Number(xf.attrs.rot) / 60000;
    if (xf?.attrs.flipH === '1' || xf?.attrs.flipH === 'true') im.flip = true;
    if (xf?.attrs.flipV === '1' || xf?.attrs.flipV === 'true') im.flipV = true;
    // 연결된 그림 (WIXEL 확장: 원본 범위) — 엑셀에서는 저장할 때의 모습(PNG)으로 보임
    const lk = descendants(child(el, 'nvPicPr'), 'linked')[0]?.attrs.ref;
    const lr = lk && parseLinkedRef(lk);
    if (lr) im.linked = lr;
    try { const media = normalizeVideo(JSON.parse(descendants(nv, 'video')[0]?.attrs.json || 'null')); if(media)im.media=media; } catch { /* Invalid optional metadata does not affect the picture. */ }
    if (emf) im.emf = emf;
    // SVG 그림(아이콘): svgBlip 원본을 화면에 쓰고 PNG 는 저장용 대체 그림으로 둠
    const svgBlip = descendants(blip, 'svgBlip')[0];
    const svgKey = svgBlip && Object.keys(svgBlip.attrs).find((k) => k.endsWith(':embed') || k === 'embed');
    const svgBytes = svgKey && files[rels[svgBlip.attrs[svgKey]]?.target];
    if (svgBytes && !emf) {
      im.png = src;
      im.src = `data:image/svg+xml;base64,${toBase64(svgBytes)}`;
      const text = new TextDecoder().decode(svgBytes);
      const m = /<svg\b([^>]*)>([\s\S]*)<\/svg>\s*$/.exec(text);
      const vb = m && /viewBox="([^"]+)"/.exec(m[1]);
      if (m && vb && m[2].length < 200000) im.icon = { vb: vb[1], body: m[2], fill: /\bfill="(#[0-9a-fA-F]{6})"/.exec(m[1])?.[1] ?? '#000000' };
    }
    // 그림 자르기 (a:srcRect, 1/1000 %)
    const sr = descendants(child(el, 'blipFill'), 'srcRect')[0];
    if (sr && ['l', 't', 'r', 'b'].some((k) => Number(sr.attrs[k]))) im.crop = Object.fromEntries(['l', 't', 'r', 'b'].filter((k) => Number(sr.attrs[k])).map((k) => [k, Number(sr.attrs[k]) / 100000]));
    // 그림 윤곽선 (a:ln 단색 채우기)
    const ln = child(child(el, 'spPr'), 'ln');
    const lc = ln && !child(ln, 'noFill') && dmlColor(child(ln, 'solidFill'), ctx.theme);
    if (lc) { im.border = lc; im.borderW = Math.max(0, Number(ln.attrs.w ?? 9525) / EMU); }
    // 표준 DrawingML 그림 투명도·둥근 모서리·그림자. 확장 메타데이터 없이 Excel과 교환합니다.
    const alpha = child(blip, 'alphaModFix');
    if (alpha) im.opacity = Math.max(0, Math.min(1, Number(alpha.attrs.amt ?? 100000) / 100000));
    const sp = child(el, 'spPr'), geom = child(sp, 'prstGeom');
    if (geom?.attrs.prst === 'roundRect') {
      const adj = kids(child(geom, 'avLst'), 'gd').find(g => g.attrs.name === 'adj');
      const match = /^val\s+(-?\d+)$/.exec(adj?.attrs.fmla ?? 'val 16667');
      im.radius = Math.min(im.w, im.h) * Math.max(0, Math.min(50000, Number(match?.[1] ?? 16667))) / 100000;
    }
    const shadow = child(child(sp, 'effectLst'), 'outerShdw');
    if (shadow) {
      const angle = Number(shadow.attrs.dir ?? 0) / 60000 * Math.PI / 180, dist = Number(shadow.attrs.dist ?? 0) / EMU;
      const rounded = n => Math.round(n * 1000) / 1000;
      im.shadow = { dx: rounded(Math.cos(angle) * dist), dy: rounded(Math.sin(angle) * dist), blur: Number(shadow.attrs.blurRad ?? 0) / EMU, color: dmlColor(shadow, ctx.theme) ?? '#000000', opacity: dmlOpacity(shadow) };
    }
    readPictureAdjustments(im, blip, sp, n => dmlColor(n, ctx.theme), dmlOpacity);
    restorePictureMetadata(im, nv, bytes, rel => {
      const path = rels[rel]?.target, data = files[path], mime = path && MIME[path.split('.').pop().toLowerCase()];
      return data && mime ? `data:${mime};base64,${toBase64(data)}` : null;
    });
    if (el.attrs.macro) im.macro = el.attrs.macro.replace(/^\[\d+\]!/, '');
    out.images.push(im);
  };

  // 그룹 도형 안의 좌표 변환
  const slicerOnlyGroup = (el) => {
    const parts = el.children.filter(k => ['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'AlternateContent'].includes(k.name));
    return parts.length > 0 && parts.every(k => k.name === 'grpSp' ? slicerOnlyGroup(k)
      : k.name === 'graphicFrame' ? !!descendants(k, 'slicer')[0]?.attrs.name
        : k.name === 'AlternateContent' && !!descendants(child(k, 'Choice'), 'slicer')[0]?.attrs.name);
  };
  const walk = (el, box, map, slicerGroup = null, client = null, editAs = 'twoCell') => {
    const place = (node) => {
      const xfrm = child(node, 'xfrm') ?? descendants(child(node, 'spPr') ?? child(node, 'grpSpPr'), 'xfrm')[0];
      if (!map || !xfrm) return box;
      const off = child(xfrm, 'off');
      const ext = child(xfrm, 'ext');
      return map(Number(off?.attrs.x ?? 0), Number(off?.attrs.y ?? 0), Number(ext?.attrs.cx ?? 0), Number(ext?.attrs.cy ?? 0));
    };
    switch (el.name) {
      case 'AlternateContent': {
        const choice = child(el, 'Choice');
        for (const node of choice?.children ?? []) if (['graphicFrame', 'grpSp'].includes(node.name)) walk(node, box, map, slicerGroup, client, editAs);
        break;
      }
      case 'sp': case 'cxnSp': readShape(el, place(el)); break;
      case 'pic': readPic(el, place(el)); break;
      case 'graphicFrame': {
        const slicer = descendants(el, 'slicer')[0];
        if (slicer?.attrs.name) {
          const b = place(el), nv = child(child(el, 'nvGraphicFramePr'), 'cNvPr');
          out._slicerBoxes[slicer.attrs.name] = { ...b, w: Math.max(1, b.w), h: Math.max(1, b.h), z: ++z,
            ...(slicerGroup ? { objectGroup: slicerGroup } : {}), ...(editAs !== 'oneCell' ? { placement: editAs } : {}),
            ...(nv?.attrs.descr !== undefined ? { alt: nv.attrs.descr } : {}), ...(['1','true'].includes(nv?.attrs.hidden) ? { hidden: true } : {}),
            ...(el.attrs.macro ? { macro: el.attrs.macro.replace(/^\[0\]!/, '') } : {}),
            ...(client?.fPrintsWithSheet !== undefined ? { noPrint: ['0','false'].includes(client.fPrintsWithSheet) } : {}),
            ...(client?.fLocksWithSheet !== undefined ? { locked: !['0','false'].includes(client.fLocksWithSheet) } : {}) };
          break;
        }
        const chartRef = descendants(el, 'chart')[0];
        const target = chartRef && rels[rid(chartRef)]?.target;
        const chart = target && readChart(files, target, ctx.theme);
        if (chart) {
          if (chartHasLogAxis(chart)) ctx.warnings.add(LOG_AXIS_WARNING);
          if (['1','true'].includes(child(child(el, 'nvGraphicFramePr'), 'cNvPr')?.attrs.hidden)) chart.hidden = true;
          if (isChartEx(chart)) {
            const props = descendants(child(child(el, 'nvGraphicFramePr'), 'cNvPr'), 'props').find((p) => p.attrs['xmlns:tb'] === 'urn:tabula:chart');
            if (props?.attrs.json) applyChartExOptions(chart, props.attrs.json);
          }
          const b = round(place(el));
          out.charts.push({ id: uid('c'), ...chart, ...b, w: Math.max(120, b.w), h: Math.max(90, b.h), z: ++z });
        }
        break;
      }
      case 'grpSp': {
        const gb = place(el);
        // Excel can collapse a group to zero width/height without ungrouping it.
        const groupBox = { ...round(gb), ...(gb.w === 0 ? { w: 0 } : {}), ...(gb.h === 0 ? { h: 0 } : {}) };
        const hyperlink = readDrawingHyperlink(child(child(el, 'nvGrpSpPr'), 'cNvPr'), rels);
        const ownGroup = readDrawingGroup(el, groupBox, uid('sh'), { files, rels });
        if (ownGroup) { out.shapes.push({ ...ownGroup, z: ++z }); break; }
        const xfrm = descendants(child(el, 'grpSpPr'), 'xfrm')[0];
        const chOff = child(xfrm, 'chOff');
        const chExt = child(xfrm, 'chExt');
        const cx = Number(chExt?.attrs.cx) || 1;
        const cy = Number(chExt?.attrs.cy) || 1;
        const ox = Number(chOff?.attrs.x ?? 0);
        const oy = Number(chOff?.attrs.y ?? 0);
        const group = slicerOnlyGroup(el) ? slicerGroup ?? uid('slg') : null;
        // 일반 그룹도 원본 자식 좌표계를 유지해야 글꼴과 선 굵기가 부모 배율을 한 번만 따른다.
        // 차트·슬라이서는 별도 개체 모델이므로 기존 시트 좌표 경로를 사용한다.
        const preserveGroup = !group && !descendants(el, 'graphicFrame').length && chExt
          && Number(chExt.attrs.cx) > 0 && Number(chExt.attrs.cy) > 0
          && Number.isFinite(cx) && cx > 0 && Number.isFinite(cy) && cy > 0
          && Number.isFinite(ox) && Number.isFinite(oy) && Number.isFinite(gb.w) && gb.w >= 0 && Number.isFinite(gb.h) && gb.h >= 0;
        const inner = preserveGroup
          ? (x, y, w, h) => ({ x: (x - ox) / EMU, y: (y - oy) / EMU, w: w / EMU, h: h / EMU })
          : (x, y, w, h) => ({ x: gb.x + ((x - ox) / cx) * gb.w, y: gb.y + ((y - oy) / cy) * gb.h, w: (w / cx) * gb.w, h: (h / cy) * gb.h });
        const ownMetadata = descendants(child(child(el, 'nvGrpSpPr'), 'cNvPr'), 'group').some(n => n.attrs['xmlns:wx'] === 'https://wixel.app/drawing/group/1');
        const beforeShapes = out.shapes.length, beforeImages = out.images.length, beforeCharts = out.charts.length;
        for (const k of el.children) if (['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame', 'AlternateContent'].includes(k.name)) walk(k, gb, inner, group, client, editAs);
        if (group && ['1','true'].includes(child(child(el, 'nvGrpSpPr'), 'cNvPr')?.attrs.hidden)) {
          for (const sl of Object.values(out._slicerBoxes)) if (sl.objectGroup === group) sl.hidden = true;
        }
        if (!preserveGroup && !ownMetadata && ['1','true'].includes(child(child(el, 'nvGrpSpPr'), 'cNvPr')?.attrs.hidden)) {
          for (const [items,start] of [[out.shapes,beforeShapes],[out.images,beforeImages],[out.charts,beforeCharts]]) for (let i=start;i<items.length;i++) items[i].hidden = true;
        }
        // 그룹으로 묶을 수 없는 차트·슬라이서 혼합 개체는 부모 링크를 자식에 전달한다.
        if (!preserveGroup && !ownMetadata && hyperlink) {
          for (const [items,start] of [[out.shapes,beforeShapes],[out.images,beforeImages],[out.charts,beforeCharts]]) for (let i=start;i<items.length;i++) if (!items[i].hyperlink) items[i].hyperlink = { ...hyperlink };
        }
        // 표준 자식 도형으로 일반 그룹을 구성하고 회전·대칭·비균등 배율을 유지한다.
        // 오래된 WIXEL 메타데이터가 있다면 현재 XML의 내용을 우선한다.
        if (preserveGroup && beforeCharts === out.charts.length) {
          const children = [...out.shapes.splice(beforeShapes), ...out.images.splice(beforeImages).map(p => ({ ...p, kind: 'picture' }))].sort((a,b) => (a.z ?? 0) - (b.z ?? 0));
          if (children.length) {
            out.shapes.push({ id: uid('sh'), kind: 'group', name: child(child(el, 'nvGrpSpPr'), 'cNvPr')?.attrs.name || '그룹', ...groupBox, z: ++z,
              ...(hyperlink ? { hyperlink } : {}),
              ...(['1','true'].includes(child(child(el, 'nvGrpSpPr'), 'cNvPr')?.attrs.hidden) ? { hidden: true } : {}),
              groupSize: { w: cx / EMU, h: cy / EMU }, groupItems: children,
              ...(child(child(el, 'nvGrpSpPr'), 'cNvPr')?.attrs.descr !== undefined ? { alt: child(child(el, 'nvGrpSpPr'), 'cNvPr').attrs.descr } : {}),
              ...(Number(xfrm?.attrs.rot) ? { rot: Number(xfrm.attrs.rot) / 60000 } : {}),
              ...(['1','true'].includes(xfrm?.attrs.flipH) ? { flip: true } : {}), ...(['1','true'].includes(xfrm?.attrs.flipV) ? { flipV: true } : {}) });
            if (ownMetadata) ctx.warnings.add('외부에서 수정된 SmartArt/그룹은 최신 도형 내용을 보존하는 일반 그룹으로 가져왔습니다.');
          }
        }
        break;
      }
      default:
    }
  };

  for (const anchor of root.children) {
    let box;
    if (anchor.name === 'twoCellAnchor') {
      const a = point(child(anchor, 'from'));
      const b = point(child(anchor, 'to'));
      box = { x: a.x, y: a.y, w: b.x - a.x, h: b.y - a.y };
    } else {
      const from = child(anchor, 'from');
      const pos = child(anchor, 'pos');
      const ext = child(anchor, 'ext');
      const p0 = from ? point(from) : { x: Number(pos?.attrs.x ?? 0) / EMU, y: Number(pos?.attrs.y ?? 0) / EMU };
      box = { ...p0, w: Number(ext?.attrs.cx ?? 480 * EMU) / EMU, h: Number(ext?.attrs.cy ?? 288 * EMU) / EMU };
    }
    const alt = child(anchor, 'AlternateContent'), choice = child(alt, 'Choice');
    const frame = child(choice, 'graphicFrame');
    // A Choice can contain a group of slicers. Only direct frames use this shortcut.
    const sl = frame && descendants(frame, 'slicer')[0];
    // 개체 위치 속성: twoCell(셀에 맞춰 위치와 크기 변경) · oneCell(위치만) · absolute(변경 안 함)
    const editAs = anchor.name === 'absoluteAnchor' ? 'absolute' : anchor.name === 'oneCellAnchor' ? 'oneCell' : anchor.attrs.editAs ?? 'twoCell';
    if (sl?.attrs.name) {
      const nv = descendants(frame, 'cNvPr')[0], client = child(anchor, 'clientData')?.attrs;
      out._slicerBoxes[sl.attrs.name] = { ...box, w: Math.max(1, box.w), h: Math.max(1, box.h), z: ++z, ...(editAs !== 'oneCell' ? { placement: editAs } : {}),
        ...(nv?.attrs.descr !== undefined ? { alt: nv.attrs.descr } : {}), ...(nv?.attrs.hidden === '1' ? { hidden: true } : {}),
        ...(frame?.attrs.macro ? { macro: frame.attrs.macro.replace(/^\[0\]!/, '') } : {}),
        ...(client?.fPrintsWithSheet !== undefined ? { noPrint: ['0','false'].includes(client.fPrintsWithSheet) } : {}),
        ...(client?.fLocksWithSheet !== undefined ? { locked: !['0','false'].includes(client.fLocksWithSheet) } : {}) };
      continue;
    }
    const content = anchor.children.find((k) => ['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame'].includes(k.name))
      ?? choice?.children.find((k) => ['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame'].includes(k.name));
    const before = [out.charts.length, out.images.length, out.shapes.length];
    if (content) walk(content, box, null, null, child(anchor, 'clientData')?.attrs, editAs);
    const client = child(anchor, 'clientData')?.attrs;
    [out.charts, out.images, out.shapes].forEach((list, k) => {
      for (let i = before[k]; i < list.length; i++) {
        if (editAs !== 'twoCell') list[i].placement = editAs;
        // 두 속성의 기본값은 true. 생략된 경우 기존 모델의 기본값을 그대로 쓴다.
        if (client?.fPrintsWithSheet !== undefined) list[i].noPrint = ['0', 'false'].includes(client.fPrintsWithSheet.trim());
        if (client?.fLocksWithSheet !== undefined) list[i].locked = !['0', 'false'].includes(client.fLocksWithSheet.trim());
      }
    });
  }
  return out;
}

/** 차트 참조 글자 'Sheet'!$A$1:$B$5 → { sheet, r1, c1, r2, c2 } */
function chartRef(text) {
  const m = /^(?:'?(.*?)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(String(text ?? '').trim().replace(/^\(|\)$/g, ''));
  if (!m) return null;
  const rg = refToRange(m[2]);
  if (!rg) return null;
  return { ...(m[1] ? { sheet: m[1].replace(/''/g, "'") } : {}), ...rg };
}

/**
 * 차트 계열 참조: 셀 범위, 또는 정의된 이름 ([0]!이름 · 시트!이름) — 이름은 그릴 때 계산 (OFFSET 등 동적 범위도 슬라이서를 따라감)
 */
function seriesRef(text) {
  const r = chartRef(text);
  if (r) return r;
  const m = /^(?:\[\d+\]!|'?([^'!]*)'?!)?([\p{L}_\\][\p{L}\p{N}_.\\?]*)$/u.exec(String(text ?? '').trim());
  if (!m || /^[A-Z]{1,3}\d+$/i.test(m[2])) return null;
  return { name: m[2], ...(m[1] ? { sheet: m[1].replace(/''/g, "'") } : {}) };
}

/** 차트 XML → WIXEL 차트 모델 (콤보 · 보조 축 · 데이터 레이블 · 계열 색 · 피벗 차트) */
// 수동 배치는 차트 왼쪽 위를 원점으로 하는 edge 좌표만 모델로 복원한다.
// factor x/y는 Excel 자동 위치에 대한 상대 이동이므로 절대 좌표로 추측하지 않는다.
function readChartLayout(element) {
  const layout = child(child(element, 'layout'), 'manualLayout');
  if (!layout || child(layout, 'xMode')?.attrs.val !== 'edge' || child(layout, 'yMode')?.attrs.val !== 'edge') return null;
  const value = key => { const v = child(layout, key)?.attrs.val; return v === undefined ? NaN : Number(v); };
  const x = value('x'), y = value('y');
  if (!Number.isFinite(x) || !Number.isFinite(y) || x < 0 || x > 1 || y < 0 || y > 1) return null;
  const out = { x, y };
  for (const [key, origin] of [['w', x], ['h', y]]) {
    const n = value(key) - (child(layout, `${key}Mode`)?.attrs.val === 'edge' ? origin : 0);
    if (Number.isFinite(n) && n > 0 && n <= 1) out[key] = n;
  }
  return out;
}

/** 표준 .crtx OPC 패키지의 차트만 읽는다. 외부 관계는 열거나 가져오지 않는다. */
export function readChartTemplatePackage(bytes) {
  const files = unzip(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const rel = descendants(parseXml(textOf(files['_rels/.rels'])), 'Relationship').find(r => r.attrs.Type?.endsWith('/officeDocument') && r.attrs.TargetMode !== 'External');
  const path = rel?.attrs.Target?.replace(/^\//, '');
  if (!path || !files[path]) throw new Error('올바른 차트 서식 파일이 아닙니다.');
  const relationships = relsOf(files, path);
  const themeRelationships = Object.fromEntries(Object.entries(relationships).map(([k, r]) => [k, r.type === 'themeOverride' ? { ...r, type: 'theme' } : r]));
  const chart = readChart(files, path, readTheme(files, themeRelationships));
  if (!chart) throw new Error('이 차트 서식의 종류를 읽을 수 없습니다.');
  return chart;
}

function chartLayoutXml(layout) {
  if (!layout || !Number.isFinite(layout.x) || !Number.isFinite(layout.y)) return '';
  const clamp = n => Math.max(0, Math.min(1, n));
  const size = ['w', 'h'].filter(k => Number.isFinite(layout[k]) && layout[k] > 0);
  return `<c:layout><c:manualLayout><c:xMode val="edge"/><c:yMode val="edge"/>${size.map(k => `<c:${k}Mode val="factor"/>`).join('')}<c:x val="${clamp(layout.x)}"/><c:y val="${clamp(layout.y)}"/>${size.map(k => `<c:${k} val="${clamp(layout[k])}"/>`).join('')}</c:manualLayout></c:layout>`;
}

function chartAreaReader(files, path, theme) {
  const rels = relsOf(files, path);
  return (spPr, kind) => readChartAreaFormat(spPr, { kind, readColor: n => dmlColor(n, theme), imageSource: id => {
    const rel = rels[id]; if (!rel || rel.external || !files[rel.target]) return null;
    const mime = MIME[rel.target.split('.').pop().toLowerCase()];
    return mime && `data:${mime};base64,${toBase64(files[rel.target])}`;
  } });
}
function readChart(files, path, theme = {}) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  if (child(root, 'chartData') && Object.values(root.attrs).includes(CHARTEX_NS)) {
    const chart = readChartEx(root, seriesRef, (n) => dmlColor(n, theme));
    const colors = Object.values(relsOf(files, path)).find((r) => r.type === 'chartColorStyle' && !r.external);
    if (chart && colors && files[colors.target]) {
      const palette = parseXml(textOf(files[colors.target])).children.filter((n) => ['srgbClr', 'schemeClr', 'sysClr', 'prstClr'].includes(n.name)).map((n) => dmlColor({ children: [n] }, theme)).filter(Boolean);
      if (palette.length) chart.palette = palette;
    }
    if (chart) { const readArea = chartAreaReader(files, path, theme); const a = readArea(child(root, 'spPr'), 'chart'), p = readArea(child(descendants(root, 'plotArea')[0], 'spPr'), 'plot'); if (a) chart.chartAreaFormat = a; if (p) chart.plotAreaFormat = p; }
    return chart;
  }
  const plot = descendants(root, 'plotArea')[0];
  if (!plot) return null;
  const kinds = { barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', radarChart: 'radar', bubbleChart: 'bubble', stockChart: 'stock', ofPieChart: 'pieOfPie', surfaceChart: 'surface', surface3DChart: 'surface' };
  const groups = plot.children.filter((c) => kinds[c.name]);
  if (!groups.length) return null;
  // 축: 오른쪽(axPos r) 또는 crosses=max 인 값 축 → 보조 축
  const valAxes = new Map(kids(plot, 'valAx').map((ax) => [child(ax, 'axId')?.attrs.val, ax]));
  const isSecondary = (group) => kids(group, 'axId').some((a) => {
    const ax = valAxes.get(a.attrs.val);
    return ax && (child(ax, 'axPos')?.attrs.val === 'r' || child(ax, 'crosses')?.attrs.val === 'max');
  });
  const typeOf = (g) => {
    let t = kinds[g.name];
    if (t === 'column' && child(g, 'barDir')?.attrs.val === 'bar') t = 'bar';
    if (g.name === 'ofPieChart' && child(g, 'ofPieType')?.attrs.val === 'bar') t = 'barOfPie';
    return t;
  };
  const dl = (el) => {
    const d = child(el, 'dLbls');
    if (!d) return undefined;
    if (child(d, 'delete')?.attrs.val === '1') return false;
    return child(d, 'showVal')?.attrs.val === '1' || child(d, 'showPercent')?.attrs.val === '1' ? true : undefined;
  };
  const fmtCode = (el) => child(el, 'numFmt')?.attrs.formatCode;
  const runFont = el => readChartTextFont(el, n => dmlColor(n, theme));
  const series = [];
  const serNames = []; // 계열 이름(파일에 저장된 값) — 자동 제목용
  const serOrder = [];
  const fmts = [];
  let sheetName = null;
  const rows = [];
  for (const g of groups) {
    const gType = typeOf(g);
    const secondary = isSecondary(g);
    const gLabels = dl(g);
    for (const ser of kids(g, 'ser')) {
      serOrder.push(Number(child(ser, 'order')?.attrs.val ?? serOrder.length));
      const order = Number(child(ser, 'order')?.attrs.val ?? series.length);
      const txF = descendants(child(ser, 'tx'), 'f')[0]?.text;
      const txV = descendants(child(ser, 'tx'), 'v')[0]?.text;
      serNames.push(txV ?? null);
      const catEl = child(ser, 'cat') ?? child(ser, 'xVal');
      const valEl = child(ser, 'val') ?? child(ser, 'yVal');
      const catF = descendants(catEl, 'f')[0]?.text;
      const valF = descendants(valEl, 'f')[0]?.text;
      const cache = [];
      for (const pt of descendants(valEl, 'pt')) { const index = Number(pt.attrs.idx); if (Number.isInteger(index) && index >= 0 && index < 1000000) cache[index] = Number(descendants(pt, 'v')[0]?.text); }
      const catCache = [];
      const catCacheEl = child(catEl, 'strLit') ?? descendants(catEl, 'strCache')[0] ?? descendants(catEl, 'numCache')[0];
      for (const pt of kids(catCacheEl, 'pt')) { const index = Number(pt.attrs.idx); if (Number.isInteger(index) && index >= 0 && index < 1000000) catCache[index] = descendants(pt, 'v')[0]?.text ?? ''; }
      const s = {
        name: txF && seriesRef(txF) ? { ref: seriesRef(txF) } : { text: txV ?? `계열${series.length + 1}` },
        ...(catF && seriesRef(catF) ? { cat: seriesRef(catF) } : {}),
        ...(valF && seriesRef(valF) ? { val: seriesRef(valF) } : { cache }),
        ...(catCache.length ? { catCache } : {}),
      };
      if ((gType === 'scatter' || gType === 'bubble') && catF && seriesRef(catF)) { s.x = s.cat; delete s.cat; }
      const szF = descendants(child(ser, 'bubbleSize'), 'f')[0]?.text;
      if (szF && seriesRef(szF)) s.size = seriesRef(szF);
      const numericCache = (el) => { const result = []; for (const pt of descendants(el, 'pt')) { const i = Number(pt.attrs.idx), v = Number(descendants(pt, 'v')[0]?.text); if (Number.isInteger(i) && i >= 0 && i < 1000000) result[i] = Number.isFinite(v) ? v : null; } return result; };
      if ((gType === 'scatter' || gType === 'bubble') && !s.x) s.xCache = numericCache(catEl);
      if (gType === 'bubble' && !s.size) s.sizeCache = numericCache(child(ser, 'bubbleSize'));
      for (const r of [s.name.ref, s.cat, s.val, s.x]) if (r?.sheet) sheetName ??= r.sheet;
      // 서식: 막대는 채우기 색, 꺾은선은 선 색
      const spPr = child(ser, 'spPr');
      const fill = dmlColor(child(spPr, 'solidFill'), theme);
      const line = dmlColor(child(child(spPr, 'ln'), 'solidFill'), theme);
      const marker = child(child(ser, 'marker'), 'symbol')?.attrs.val;
      const f = {};
      const shape = child(ser, 'shape')?.attrs.val ?? child(g, 'shape')?.attrs.val;
      if (g.name === 'bar3DChart' && ['box', 'cylinder', 'cone', 'pyramid'].includes(shape)) f.barShape = shape;
      const groupMode = child(g, 'grouping')?.attrs.val;
      if (groups.length > 1 && ['clustered', 'stacked', 'percentStacked', 'standard'].includes(groupMode)) f.grouping = groupMode === 'standard' && g.name !== 'bar3DChart' ? 'clustered' : groupMode;
      const seriesExplosion = child(ser, 'explosion')?.attrs.val;
      if (seriesExplosion !== undefined && Number.isFinite(Number(seriesExplosion))) f.explode = Math.max(0, Math.min(400, Number(seriesExplosion)));
      const pointExplosion = {};
      for (const dp of kids(ser, 'dPt')) {
        const index = Number(child(dp, 'idx')?.attrs.val), value = child(dp, 'explosion')?.attrs.val;
        if (Number.isInteger(index) && index >= 0 && index < 1000000 && value !== undefined && Number.isFinite(Number(value))) pointExplosion[index] = Math.max(0, Math.min(400, Number(value)));
      }
      if (Object.keys(pointExplosion).length) f.pointExplosion = pointExplosion;
      if (groups.length > 1 || gType !== typeOf(groups[0])) f.type = gType;
      if (secondary) f.axis = 1;
      const color = gType === 'line' || gType === 'scatter' ? line ?? fill : fill ?? line;
      if (color) f.color = color;
      // 그라데이션 채우기 · 그림자 · 선 굵기 · 표식 크기 · 레이블 글꼴 (엑셀 차트 스타일 그대로)
      const grad = child(spPr, 'gradFill');
      if (grad && gType !== 'line' && gType !== 'scatter') {
        const stops = kids(child(grad, 'gsLst'), 'gs').map((gs) => [Number(gs.attrs.pos ?? 0) / 100000, dmlColor(gs, theme)]).filter(([, c]) => c);
        if (stops.length > 1) f.grad = { stops, ang: Number(child(grad, 'lin')?.attrs.ang ?? 5400000) / 60000 };
        if (!f.color && stops.length) f.color = stops[Math.floor(stops.length / 2)][1];
      }
      if (descendants(child(spPr, 'effectLst'), 'outerShdw').length) f.shadow = true;
      const lnW = Number(child(spPr, 'ln')?.attrs.w ?? 0);
      if (lnW && ['line', 'scatter', 'radar'].includes(gType)) f.lineWidth = Math.round((lnW / 12700) * (4 / 3) * 100) / 100;
      const mSize = Number(child(child(ser, 'marker'), 'size')?.attrs.val ?? 0);
      if (mSize) f.markerSize = mSize;
      const dsh = DASH_FROM[child(child(spPr, 'ln'), 'prstDash')?.attrs.val];
      if (dsh && ['line', 'scatter', 'radar'].includes(gType)) f.dash = dsh;
      const smoothValue = child(ser, 'smooth')?.attrs.val;
      if (smoothValue !== undefined) f.smooth = smoothValue === '1' || smoothValue === 'true';
      if (['bar', 'column', 'pie', 'doughnut'].includes(gType) && child(spPr, 'ln') && child(child(spPr, 'ln'), 'solidFill')) f.outline = dmlColor(child(child(spPr, 'ln'), 'solidFill'), theme) ?? undefined;
      const tl = child(ser, 'trendline');
      if (tl) {
        const tt = child(tl, 'trendlineType')?.attrs.val;
        if (tt === 'linear' || tt === 'exp' || tt === 'movingAvg') {
          f.trend = tt;
          if (tt === 'movingAvg') f.trendPeriod = Number(child(tl, 'period')?.attrs.val ?? 2);
          const fw = Number(child(tl, 'forward')?.attrs.val ?? 0);
          if (fw) f.trendForward = fw;
        }
      }
      const lposV = child(child(ser, 'dLbls'), 'dLblPos')?.attrs.val;
      const LPOS = { ctr: 'center', inEnd: 'insideEnd', inBase: 'insideBase', outEnd: 'outEnd', t: 'above', b: 'below', l: 'left', r: 'right' };
      if (LPOS[lposV] && gType !== 'pie' && gType !== 'doughnut') f.labelPos = LPOS[lposV];
      {
        const pc = {};
        for (const dp of kids(ser, 'dPt')) { const c = dmlColor(child(child(dp, 'spPr'), 'solidFill'), theme); if (c) pc[Number(child(dp, 'idx')?.attrs.val)] = c; }
        if (Object.keys(pc).length) f.pointColors = pc;
      }
      const lf = { ...runFont(child(child(g, 'dLbls'), 'txPr')), ...runFont(child(child(ser, 'dLbls'), 'txPr')) };
      assignChartTextFields(f, lf, 'label');
      const pointLabels = {};
      for (const labels of [child(g, 'dLbls'), child(ser, 'dLbls')]) for (const label of kids(labels, 'dLbl')) {
        const p = Number(child(label, 'idx')?.attrs.val);
        const font = { ...runFont(child(label, 'txPr')), ...runFont(child(label, 'tx')) };
        if (Number.isInteger(p) && p >= 0 && Object.keys(font).length) pointLabels[p] = { ...pointLabels[p], ...font };
      }
      if (Object.keys(pointLabels).length) f.pointLabelStyles = pointLabels;
      const lab = dl(ser) ?? gLabels;
      if (gType === 'pie' || gType === 'doughnut') {
        // 원형 레이블: 파일에 적힌 대로 (없으면 표시하지 않음 — 엑셀과 같음)
        const d = child(ser, 'dLbls') ?? child(g, 'dLbls');
        const on = (k) => child(d, k)?.attrs.val === '1';
        if (d && child(d, 'delete')?.attrs.val !== '1' && on('showPercent') && !on('showVal')) f.pct = true;
        else if (d && child(d, 'delete')?.attrs.val !== '1' && on('showVal')) f.labels = true;
        else if (!(d && child(d, 'delete')?.attrs.val !== '1' && on('showCatName'))) f.labels = false;
        // 항목 이름 (엑셀: 이름 + 값/백분율 두 줄, 조각 바깥)
        if (d && child(d, 'delete')?.attrs.val !== '1' && on('showCatName')) f.catName = true;
        const pos = child(d, 'dLblPos')?.attrs.val;
        if (pos === 'outEnd' || pos === 'bestFit') f.labelPos = 'out';
      } else if (lab) f.labels = true;
      const labelNode = child(ser, 'dLbls') ?? child(g, 'dLbls');
      if (labelNode) {
        const deleted = child(labelNode, 'delete')?.attrs.val === '1';
        for (const [xmlKey, key] of [['showVal', 'labels'], ['showCatName', 'catName'], ['showSerName', 'serName'], ['showPercent', 'pct']]) {
          const v = child(labelNode, xmlKey)?.attrs.val;
          if (deleted || v !== undefined && (key === 'labels' || ['1', 'true'].includes(v))) f[key] = !deleted && ['1', 'true'].includes(v);
        }
      }
      if (marker) f.marker = marker === 'none' ? 'none' : marker;
      const code = fmtCode(child(ser, 'dLbls')) ?? descendants(valEl, 'formatCode')[0]?.text;
      if (code && code !== 'General') f.numFmt = code;
      if (gType === 'pie' || gType === 'doughnut') {
        const colors = kids(ser, 'dPt').map((dp) => [Number(child(dp, 'idx')?.attrs.val), dmlColor(child(child(dp, 'spPr'), 'solidFill'), theme)]).filter(([, c]) => c);
        if (colors.length) { f.colors = []; for (const [i, c] of colors) f.colors[i] = c; }
      }
      rows.push({ order, s, f });
    }
  }
  rows.sort((a, b) => a.order - b.order);
  for (const r of rows) { series.push(r.s); fmts.push(r.f); }
  // 마지막 계열을 지운 차트도 제목/범례/빈 그림 영역을 유지한다.
  const refs = series.flatMap((s) => [s.name.ref, s.cat, s.val, s.x]).filter((r) => r && r.r1 !== undefined);
  const range = refs.length ? {
    r1: minOf(refs.map((r) => r.r1)), c1: minOf(refs.map((r) => r.c1)),
    r2: maxOf(refs.map((r) => r.r2)), c2: maxOf(refs.map((r) => r.c2)),
  } : null;
  const chartEl = descendants(root, 'chart')[0] ?? root;
  const titleEl = child(chartEl, 'title');
  let title = titleEl ? descendants(titleEl, 't').map((t) => t.text).join('') : '';
  // 셀에 연결된 제목(strRef)은 저장된 값, 제목 요소만 있고 글자가 없으면 엑셀처럼 계열이 하나일 때 계열 이름(자동 제목)
  if (titleEl && !title) title = descendants(child(titleEl, 'tx'), 'v')[0]?.text ?? '';
  if (titleEl && !title && serNames.length === 1) title = serNames[0] ?? '';
  if (titleEl && !title && (groups.some((g) => g.name === 'pieChart' || g.name === 'doughnutChart' || g.name === 'pie3DChart'))) title = serNames[0] ?? '';
  if (!titleEl && child(chartEl, 'autoTitleDeleted')?.attrs.val === '0' && serNames.length === 1) title = serNames[0] ?? '';
  const types = new Set(fmts.map((f, i) => f.type ?? typeOf(groups[0])));
  const out = { type: types.size > 1 ? 'combo' : typeOf(groups[0]), title, series };
  assignChartTextFields(out, runFont(child(root, 'txPr')), '', true);
  if (groups.some((g) => /3DChart$/.test(g.name))) {
    out.threeD = true;
    const view = child(chartEl, 'view3D');
    const config = {};
    for (const key of ['rotX', 'rotY', 'depthPercent', 'perspective']) {
      const value = child(view, key)?.attrs.val;
      if (value !== undefined && Number.isFinite(Number(value))) config[key] = Number(value);
    }
    if (child(view, 'rAngAx')) config.rAngAx = child(view, 'rAngAx').attrs.val !== '0';
    out.view3D = chartView3D({ type: out.type, view3D: config });
  }
  // 글꼴 크기(pt): 제목 · 축 · 범례 (파일에 있을 때만)
  const tf = runFont(child(titleEl, 'tx')) ;
  const tf2 = { ...runFont(child(titleEl, 'txPr')), ...tf };
  assignChartTextFields(out, tf2, 'title');
  if (title && child(titleEl, 'overlay')?.attrs.val === '1') out.titleOverlay = true;
  const lg = runFont(child(child(chartEl, 'legend'), 'txPr'));
  assignChartTextFields(out, lg, 'legend');
  if (range) out.range = range;
  if (sheetName) out.sheet = sheetName;
  if (fmts.some((f) => Object.keys(f).length)) out.seriesFmt = fmts;
  const bar = groups.find((g) => /^bar(?:3D)?Chart$/.test(g.name)) ?? groups.find((g) => /^(line|area)(?:3D)?Chart$/.test(g.name));
  const grouping = child(bar, 'grouping')?.attrs.val;
  if (grouping === 'stacked' || grouping === 'percentStacked' || grouping === 'standard' && bar?.name === 'bar3DChart') out.grouping = grouping;
  const barShape = child(bar, 'shape')?.attrs.val;
  if (bar?.name === 'bar3DChart' && ['box', 'cylinder', 'cone', 'pyramid'].includes(barShape)) out.barShape = barShape;
  const g0 = groups[0];
  const scatterStyle = child(g0, 'scatterStyle')?.attrs.val;
  if (scatterStyle) {
    const noLines = kids(g0, 'ser').every((s) => child(child(child(s, 'spPr'), 'ln'), 'noFill'));
    out.scatterStyle = noLines && /line|smooth/i.test(scatterStyle) ? 'marker' : scatterStyle;
  }
  const rs = child(g0, 'radarStyle')?.attrs.val;
  if (rs === 'filled' || rs === 'marker' && !kids(g0, 'ser').every((s) => child(child(s, 'marker'), 'symbol')?.attrs.val === 'none')) out.radarStyle = rs;
  if (['pieChart', 'pie3DChart', 'doughnutChart'].includes(g0.name)) {
    const explosions = kids(g0, 'ser').map((s) => Number(child(s, 'explosion')?.attrs.val ?? 0));
    if (explosions[0] > 0 && explosions.every((v) => v === explosions[0])) out.explode = explosions[0];
  }
  const stockGroup = groups.find((g) => g.name === 'stockChart');
  if (stockGroup) {
    out.type = 'stock';
    out.ohlc = !!child(stockGroup, 'upDownBars');
    out.volume = groups.some((g) => g.name === 'barChart');
    for (const f of fmts) { delete f.type; delete f.axis; }
  }
  if (g0.name === 'bubbleChart') { out.threeD = descendants(g0, 'bubble3D').some((n) => ['1', 'true'].includes(n.attrs.val)); out.showNegBubbles = ['1', 'true'].includes(child(g0, 'showNegBubbles')?.attrs.val); }
  if (g0.name === 'ofPieChart') {
    out.splitType = ({ cust: 'custom', pos: 'position', val: 'value', percent: 'percent' })[child(g0, 'splitType')?.attrs.val] ?? 'position';
    out.splitPos = Number(child(g0, 'splitPos')?.attrs.val ?? 3);
    out.secondSize = Number(child(g0, 'secondPieSize')?.attrs.val ?? 75);
    out.splitGap = Number(child(g0, 'gapWidth')?.attrs.val ?? 150);
    const points = kids(child(g0, 'custSplit'), 'secondPiePt').map((pt) => Number(pt.attrs.val)).filter(Number.isInteger);
    if (points.length) out.splitPoints = points;
  }
  if (out.type === 'surface') {
    const wire = ['1','true'].includes(child(g0, 'wireframe')?.attrs.val);
    out.threeD = g0.name === 'surface3DChart';
    out.surfaceStyle = out.threeD ? (wire ? 'wireframe' : 'surface') : (wire ? 'wireframeContour' : 'contour');
  }
  const hole = Number(child(g0, 'holeSize')?.attrs.val);
  if (hole && hole !== 50) out.hole = hole;
  const gapW = Number(child(bar, 'gapWidth')?.attrs.val);
  if (/^bar(?:3D)?Chart$/.test(bar?.name) && Number.isFinite(gapW) && gapW !== 150 && gapW !== 182) out.gap = gapW;
  const ovl = Number(child(bar, 'overlap')?.attrs.val);
  if (bar?.name === 'barChart' && Number.isFinite(ovl) && ovl !== 0 && ovl !== 100) out.overlap = ovl;
  if (/^bar(?:3D)?Chart$/.test(bar?.name) && child(bar, 'varyColors')?.attrs.val === '1') out.varyColors = true;
  if (child(plot, 'dTable')) {
    out.dataTable = true;
    const font = runFont(child(child(plot, 'dTable'), 'txPr'));
    if (Object.keys(font).length) out.dataTableText = font;
  }
  if (!descendants(plot, 'majorGridlines').length) out.gridY = false;
  // WIXEL 전용 설정 (원래 차트 종류 · 팔레트 · 서식)
  const nativeText = {};
  for (const prefix of ['', 'title', 'legend']) assignChartTextFields(nativeText, chartTextFields(out, prefix, !prefix), prefix, !prefix);
  if (out.dataTableText) nativeText.dataTableText = out.dataTableText;
  const tbEl = descendants(root, 'props').find((x) => x.attrs.json);
  if (tbEl) { try { Object.assign(out, JSON.parse(tbEl.attrs.json), nativeText); } catch { /* 무시 */ } }
  if (out.wxPivot) { out.pivot = out.wxPivot; delete out.wxPivot; }
  const legend = child(chartEl, 'legend');
  const titleLayout = readChartLayout(titleEl), legendLayout = readChartLayout(legend);
  if (titleLayout) out.titleLayout = titleLayout;
  if (legendLayout) out.legendLayout = legendLayout;
  out.legend = legend ? (child(legend, 'legendPos')?.attrs.val ?? 'r') : 'none';
  if (out.legend === 'tr') out.legend = 'r';
  // 축 서식 · 제목 · 최소/최대
  const axInfo = (ax) => {
    if (!ax) return null;
    const o = {};
    const code = fmtCode(ax);
    if (code && code !== 'General' && child(ax, 'numFmt')?.attrs.sourceLinked !== '1') o.numFmt = code;
    if (child(child(ax, 'scaling'), 'orientation')?.attrs.val === 'maxMin' && ax.name === 'valAx') o.reverse = true;
    const t = child(ax, 'title');
    if (t) o.title = descendants(t, 't').map((x) => x.text).join('');
    assignChartTextFields(o, runFont(child(ax, 'txPr')));
    assignChartTextFields(o, { ...runFont(child(t, 'txPr')), ...runFont(child(t, 'tx')) }, 'title');
    const sc = child(ax, 'scaling');
    if (child(sc, 'logBase')) o.logBase = Number(child(sc, 'logBase').attrs.val);
    if (child(sc, 'min')) o.min = Number(child(sc, 'min').attrs.val);
    if (child(sc, 'max')) o.max = Number(child(sc, 'max').attrs.val);
    if (child(ax, 'majorUnit')) o.major = Number(child(ax, 'majorUnit').attrs.val);
    if (ax.name === 'catAx' || ax.name === 'dateAx') {
      const interval = Number(child(ax, 'tickLblSkip')?.attrs.val);
      if (Number.isSafeInteger(interval) && interval > 0) o.labelInterval = interval;
      const rawRotation = child(child(ax, 'txPr'), 'bodyPr')?.attrs.rot;
      const rotation = rawRotation === undefined ? NaN : Number(rawRotation) / 60000;
      if (Number.isFinite(rotation) && Math.abs(rotation) <= 90) o.labelRotation = rotation;
    }
    if (child(ax, 'delete')?.attrs.val === '1' || child(ax, 'delete')?.attrs.val === 'true') o.hide = true;
    return Object.keys(o).length ? o : null;
  };
  const vals = [...valAxes.values()];
  const secondaryValueAxis = (ax) => child(ax, 'axPos')?.attrs.val === 'r' || child(ax, 'crosses')?.attrs.val === 'max';
  const primaryVals = vals.filter((ax) => !secondaryValueAxis(ax));
  const primaryAx = primaryVals.find((ax) => child(ax, 'axPos')?.attrs.val === (out.type === 'bar' ? 'b' : 'l')) ?? primaryVals[0];
  const secondaryAx = vals.find(secondaryValueAxis);
  const axes = {};
  if (axInfo(primaryAx)) axes.y = axInfo(primaryAx);
  if (axInfo(secondaryAx)) axes.y2 = axInfo(secondaryAx);
  const catAx = ['scatter', 'bubble'].includes(out.type)
    ? primaryVals.find(ax => ['b', 't'].includes(child(ax, 'axPos')?.attrs.val))
    : kids(plot, 'catAx')[0] ?? kids(plot, 'dateAx')[0];
  if (axInfo(catAx)) axes.x = axInfo(catAx);
  if (Object.keys(axes).length) out.axes = axes;
  const axisFonts = [catAx, primaryAx, secondaryAx].filter(Boolean).map(ax => runFont(child(ax, 'txPr')));
  if (axisFonts.length && axisFonts[0].size && axisFonts.every(font => font.size === axisFonts[0].size)) out.axisSize = axisFonts[0].size;
  applyChartTextSupplement(out, out.wxText); delete out.wxText;
  // 피벗 차트: [파일]시트!피벗 이름
  const ps = descendants(child(root, 'pivotSource'), 'name')[0]?.text;
  if (ps) {
    const m = /^(?:\[[^\]]*\])?'?(.*?)'?!(.+)$/.exec(ps.trim());
    if (m) out.pivot = { sheet: m[1].replace(/''/g, "'"), name: m[2] };
  }
  const readArea = chartAreaReader(files, path, theme);
  const chartArea = readArea(child(root, 'spPr'), 'chart'), plotArea = readArea(child(plot, 'spPr'), 'plot');
  if (chartArea) out.chartAreaFormat = chartArea; if (plotArea) out.plotAreaFormat = plotArea;
  return out;
}

// ───────────────────────── 피벗 테이블 · 슬라이서 (읽기) ─────────────────────────
/** ISO 날짜 · 시각 (시간대 없음) → 엑셀 날짜 일련번호 */
export function isoSerial(v, date1904 = false) {
  const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(v ?? '');
  if (!m) return v ?? '';
  const ms = Date.UTC(1970, 0, 1, +(m[4] ?? 0), +(m[5] ?? 0)) + Math.round(+(m[6] ?? 0) * 1000);
  return serialOf(+m[1], +m[2], +m[3], date1904) + ms / DAY_MS;
}
/** pivotCacheDefinition → { source: { ref, sheet, name }, fields: [{ name, items: [값] }] } */
function readPivotCache(files, path, date1904 = false) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const ws = descendants(child(root, 'cacheSource'), 'worksheetSource')[0];
  const olap = !!(child(root, 'cacheHierarchies') || child(root, 'dimensions') || child(root, 'measureGroups'));
  const fields = kids(child(root, 'cacheFields'), 'cacheField').map((cf) => {
    const sharedNode = child(cf, 'sharedItems'), sharedNodes = sharedNode?.children ?? [];
    // An unlisted date field can store its values only in cacheRecords.
    const dateOnly = !sharedNodes.length && ['1', 'true'].includes(sharedNode?.attrs.containsDate) && ['0', 'false'].includes(sharedNode?.attrs.containsNonDate);
    const sharedTypes = sharedNodes.some(it => it.name === 'd') ? sharedNodes.map(it => it.name).join('') : null;
    const items = sharedNodes.map((it) => {
      if (it.name === 'n') return Number(it.attrs.v);
      // 날짜 항목 (<d v="2026-05-01T00:00:00"/>): 원본 셀처럼 날짜 일련번호로 (필터 · 슬라이서 선택이 원본과 맞게)
      if (it.name === 'd') return isoSerial(it.attrs.v, date1904);
      if (it.name === 'b') return it.attrs.v === '1' || it.attrs.v === 'true';
      if (it.name === 'm') return null;
      if (it.name === 'e') return { error: it.attrs.v ?? '#N/A' }; // 오류 항목: 피벗 합계가 오류가 됨
      return unx(it.attrs.v ?? '');
    });
    // tb:formula = WIXEL 가 쓴 원래 수식 (DIVIDE · ROWS 등 엑셀에 없는 함수)
    const tbKey = Object.keys(cf.attrs).find((k) => k.endsWith(':formula'));
    const formula = tbKey ? cf.attrs[tbKey] : cf.attrs.formula;
    // 그룹화 (날짜 → 연 · 분기 · 월 · 일, 숫자 구간): 피벗 필드의 항목 번호는 groupItems 를 가리킴
    const fg = child(cf, 'fieldGroup');
    let group = null;
    let gitems = null;
    if (fg) {
      const rp = child(fg, 'rangePr');
      const gi = child(fg, 'groupItems');
      if (gi) gitems = (gi.children ?? []).map((it) => (it.name === 'n' ? Number(it.attrs.v) : it.name === 'm' ? null : it.attrs.v ?? ''));
      const dp = child(fg, 'discretePr');
      // 선택 항목 그룹화: discretePr 의 x = 원본 필드 항목마다 groupItems 번호
      if (dp && !rp) group = { base: fg.attrs.base !== undefined ? Number(fg.attrs.base) : null, derived: true, by: 'discrete', disc: kids(dp, 'x').map((x) => Number(x.attrs.v)) };
      if (rp) {
        group = {
          base: fg.attrs.base !== undefined ? Number(fg.attrs.base) : null,
          derived: cf.attrs.databaseField === '0',
          by: rp.attrs.groupBy ?? 'range',
          ...(rp.attrs.startDate ? { start: isoSerial(rp.attrs.startDate, date1904), end: isoSerial(rp.attrs.endDate, date1904) } : {}),
          ...(rp.attrs.startNum !== undefined ? { startNum: Number(rp.attrs.startNum), endNum: Number(rp.attrs.endNum), size: Number(rp.attrs.groupInterval ?? 1) } : {}),
        };
      }
    }
    return { name: unx(cf.attrs.name ?? ''), numFmtId: Number(cf.attrs.numFmtId ?? 0), items: gitems && (group || !items.length) ? gitems : items, shared: items, ...(sharedTypes ? { sharedTypes } : {}), ...(dateOnly ? { dateOnly: true } : {}), db: formula === undefined && cf.attrs.databaseField !== '0', ...(group ? { group } : {}), ...(formula !== undefined ? { formula } : {}) };
  });
  const recRel = Object.values(relsOf(files, path)).find((r) => r.type === 'pivotCacheRecords');
  const limit = /^\d+$/.test(root.attrs.missingItemsLimit ?? '') ? Number(root.attrs.missingItemsLimit) : NaN;
  const missingItems = limit === 0 ? 'none' : limit === 1048576 ? 'max' : Number.isInteger(limit) && limit <= 0xffffffff ? limit : undefined;
  const saveData = !['0', 'false'].includes(root.attrs.saveData);
  const refreshOnOpen = ['1', 'true'].includes(root.attrs.refreshOnLoad);
  const recordCount = /^\d+$/.test(root.attrs.recordCount ?? '') ? Number(root.attrs.recordCount) : 0;
  const sourceRel = ws && rid(ws) ? relsOf(files, path)[rid(ws)] : null;
  const external = sourceRel?.external || sourceRel?.targetMode === 'External' || sourceRel?.type === 'externalLinkPath';
  // 외부 원본을 가져오지 않으므로 열 때 새로 고침 요청이 있어도 실제 저장 레코드는 보존합니다.
  const snapshot = !saveData || refreshOnOpen && !external || !recRel ? null : { path: recRel.target, count: Number.isSafeInteger(recordCount) ? recordCount : 0 };
  const slicerId = descendants(child(root, 'extLst'), 'pivotCacheDefinition').find(x => x.attrs.pivotCacheId !== undefined)?.attrs.pivotCacheId;
  return { date1904, slicerId, olap, saveData, refreshOnOpen, missingItems, source: { ref: ws?.attrs.ref ?? null, sheet: ws?.attrs.sheet ?? null, name: ws?.attrs.name ?? null, ...(external ? { external: sourceRel.rawTarget ?? sourceRel.target } : {}) }, fields, snapshot };
}

/**
 * 피벗 캐시 레코드 (엑셀이 마지막으로 새로 고칠 때 저장한 원본) → [머리글, ...행]
 * 엑셀은 열 때 이 저장본으로 피벗을 보여 줌 (원본을 고친 뒤 새로 고치지 않았으면 원본과 다를 수 있음)
 */
function* cacheRecordsSteps(files, cache, streamParts = null) {
  const bytes = streamParts?.has(cache.snapshot.path) ? zipEntryChunks(files, cache.snapshot.path) : files[cache.snapshot.path];
  if (!bytes) return null;
  if (isPivotSnapshot(bytes)) return bytes;
  const binary = files.__xlsb?.cacheRecords.get(cache.snapshot.path);
  return yield* (binary
    ? readPivotSnapshotBinary(bytes, cache.fields, binary, cache.date1904)
    : readPivotSnapshotXml(bytes, cache.fields, cache.date1904, cache.snapshot.count));
}
function readCacheRecords(files, cache) {
  const steps = cacheRecordsSteps(files, cache);
  for (;;) { const next = steps.next(); if (next.done) return next.value; }
}

const AGG_FROM_XLSX = { sum: 'sum', count: 'count', countNums: 'countNums', average: 'average', max: 'max', min: 'min', product: 'product', stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp' };
// 값 표시 형식: 기본 이름 (showDataAs) + x14 확장 (pivotShowAs)
const SHOW_FROM_XLSX = {
  percentOfTotal: 'percentOfTotal', percentOfRow: 'percentOfRow', percentOfCol: 'percentOfCol', percent: 'percent', difference: 'difference',
  percentDiff: 'percentDiff', runTotal: 'runTotal', index: 'index', percentOfParentRow: 'percentOfParentRow', percentOfParentCol: 'percentOfParentCol',
  percentOfParent: 'percentOfParent', percentOfRunningTotal: 'percentOfRunningTotal', rankAscending: 'rankAscending', rankDescending: 'rankDescending',
};
const SHOW_BASE_ONLY = new Set(['normal', 'difference', 'percent', 'percentDiff', 'runTotal', 'percentOfRow', 'percentOfCol', 'percentOfTotal', 'index']);
const SHOW_NEEDS_ITEM = new Set(['percent', 'difference', 'percentDiff']);
const SHOW_NEEDS_FIELD = new Set([...SHOW_NEEDS_ITEM, 'percentOfParent', 'runTotal', 'percentOfRunningTotal', 'rankAscending', 'rankDescending']);
const BASE_PREV = 1048828;
const BASE_NEXT = 1048829;

/**
 * pivotTableDefinition + 캐시 → WIXEL 피벗 정의 (행·열·값·보고서 필터 여러 개, 레이아웃, 부분합, 총합계, 값 표시 형식)
 * tables: 가져온 모든 표 (원본이 표 이름인지 확인)
 */
function pivotDefFrom(root, cache, tables, sheetName) {
  const names = cache.fields.map((f, i) => f.name || `열${i + 1}`);
  const fieldsOf = (tag) => kids(child(root, tag), 'field').map((f) => Number(f.attrs.x)).filter((x) => x >= 0 && x < names.length);
  const rowF = fieldsOf('rowFields');
  const colF = fieldsOf('colFields');
  const pageEls = kids(child(root, 'pageFields'), 'pageField').filter((p) => Number(p.attrs.fld) < names.length);
  const dataF = kids(child(root, 'dataFields'), 'dataField');
  const values = dataF.map((df) => {
    const v = { field: names[Number(df.attrs.fld)] ?? names[0], agg: AGG_FROM_XLSX[df.attrs.subtotal ?? 'sum'] ?? 'sum' };
    if (df.attrs.name && unx(df.attrs.name) !== valueName(v)) v.name = unx(df.attrs.name);
    const x14 = descendants(df, 'dataField')[0];
    const as = SHOW_FROM_XLSX[x14?.attrs.pivotShowAs] ?? SHOW_FROM_XLSX[df.attrs.showDataAs];
    if (as) {
      v.showAs = as;
      const bf = Number(df.attrs.baseField ?? 0);
      if (SHOW_NEEDS_FIELD.has(as) && names[bf] !== undefined) v.baseField = names[bf];
      const bi = Number(df.attrs.baseItem ?? 0);
      if (!SHOW_NEEDS_ITEM.has(as)) { /* 기준 항목 없음 */ } else if (bi === BASE_PREV) v.basePos = 'prev';
      else if (bi === BASE_NEXT) v.basePos = 'next';
      else {
        const its = kids(child(kids(child(root, 'pivotFields'), 'pivotField')[bf], 'items'), 'item').filter((it) => it.attrs.x !== undefined);
        const it = its[bi];
        if (it) v.baseItem = itemText(keyOf(cache.fields[bf]?.items[Number(it.attrs.x)] ?? null));
      }
    }
    return v;
  });
  const pfs = kids(child(root, 'pivotFields'), 'pivotField');
  const first = pfs[rowF[0] ?? colF[0]];
  const tblCompact = root.attrs.compact !== '0';
  const tblOutline = root.attrs.outline !== '0';
  // 필드의 compact · outline 기본값은 켜짐 (표의 compact="0" 은 새로 넣는 필드의 기본값일 뿐)
  const fCompact = first ? first.attrs.compact !== '0' : tblCompact;
  const fOutline = first ? first.attrs.outline !== '0' : tblOutline;
  // 열 영역에서 '값'(x=-2)의 위치 → valuesPos (맨 안쪽이면 생략)
  const colAll = kids(child(root, 'colFields'), 'field').map((f) => Number(f.attrs.x)).filter((x) => x === -2 || (x >= 0 && x < names.length));
  const vIdx = colAll.indexOf(-2);
  const def = {
    rows: rowF.map((f) => names[f]), cols: colF.map((f) => names[f]), values, pages: pageEls.map((p) => names[Number(p.attrs.fld)]),
    pageMulti: Object.fromEntries(pfs.flatMap((pf, f) => pf.attrs.multipleItemSelectionAllowed !== undefined || pageEls.some((p) => Number(p.attrs.fld) === f)
      ? [[names[f], pf.attrs.multipleItemSelectionAllowed === '1' || pf.attrs.multipleItemSelectionAllowed === 'true']] : [])),
    ...(vIdx >= 0 && vIdx < colF.length ? { valuesPos: vIdx } : {}),
    ...(root.attrs.dataOnRows === '1' && values.length > 1 ? { valuesOnRows: true } : {}),
    layout: !fOutline ? 'tabular' : !fCompact ? 'outline' : 'compact',
    classic: root.attrs.gridDropZones === '1' || root.attrs.gridDropZones === 'true',
    ...(cache.saveData === false ? { saveData: false } : {}),
    ...(cache.refreshOnOpen ? { refreshOnOpen: true } : {}),
    ...(cache.missingItems !== undefined ? { missingItems: cache.missingItems } : {}),
  };
  // 부분합: 필드마다 (바깥 필드만 켜 둔 보고서가 많음). 안쪽 끝 필드는 부분합이 없으므로 셈에서 뺌
  {
    const inner = [...rowF.slice(0, -1), ...colF.filter((x) => x >= 0).slice(0, -1)];
    const on = inner.filter((f) => pfs[f]?.attrs.defaultSubtotal !== '0');
    if (!inner.length) { if ([...rowF, ...colF].some((f) => pfs[f]?.attrs.defaultSubtotal === '0')) def.subtotals = false; }
    else if (!on.length) def.subtotals = false;
    else if (on.length < inner.length) def.subtotals = on.map((f) => names[f]);
  }
  if (root.attrs.rowGrandTotals === '0') def.grandRows = false;
  if (root.attrs.colGrandTotals === '0') def.grandCols = false;
  if (!def.pages.length) delete def.pages;
  else {
    def.pageOrder = root.attrs.pageOverThenDown === '1' || root.attrs.pageOverThenDown === 'true' ? 'over' : 'down';
    const wrap = Number(root.attrs.pageWrap);
    def.pageWrap = Number.isInteger(wrap) && wrap > 0 && wrap <= 4294967295 ? wrap : 0;
  }
  const src = cache.source;
  const tbl = !src.external && src.name ? tables.find((t) => t.name.toLowerCase() === src.name.toLowerCase()) : null;
  if (tbl) Object.assign(def, { table: tbl.name, source: tbl.sheetName, range: { r1: tbl.r1, c1: tbl.c1, r2: tbl.r2, c2: tbl.c2 } });
  else if (src.ref && refToRange(src.ref)) Object.assign(def, { source: src.sheet ?? sheetName, range: refToRange(src.ref) });
  else if (src.name) Object.assign(def, { table: src.name, ...(src.sheet ? { source: src.sheet } : {}) });
  else return null;
  // 숨긴 항목 · 보고서 필터에서 고른 항목 → 필터 (슬라이서 선택 상태)
  const filters = {};
  pfs.forEach((pf, f) => {
    const its = kids(child(pf, 'items'), 'item').filter((it) => it.attrs.x !== undefined);
    if (!its.some((it) => it.attrs.h === '1')) return;
    filters[names[f]] = its.filter((it) => it.attrs.h !== '1').map((it) => itemText(keyOf(cache.fields[f]?.items[Number(it.attrs.x)] ?? null)));
  });
  for (const p of pageEls) {
    if (p.attrs.item === undefined || def.pageMulti[names[Number(p.attrs.fld)]]) continue;
    const f = Number(p.attrs.fld);
    const its = kids(child(pfs[f], 'items'), 'item').filter((it) => it.attrs.x !== undefined);
    const it = its[Number(p.attrs.item)];
    if (it) filters[names[f]] = [itemText(keyOf(cache.fields[f]?.items[Number(it.attrs.x)] ?? null))];
  }
  if (Object.keys(filters).length) def.filters = filters;
  // 셀에서 바꿔 쓴 필드 이름 (pivotField@name) · 항목 이름 (item@n)
  pfs.forEach((pf, f) => {
    if (pf.attrs.name && unx(pf.attrs.name) !== names[f]) (def.fieldCaptions ??= {})[names[f]] = unx(pf.attrs.name);
    for (const it of kids(child(pf, 'items'), 'item')) {
      if (it.attrs.n === undefined || it.attrs.x === undefined) continue;
      const t = itemText(keyOf(cache.fields[f]?.items[Number(it.attrs.x)] ?? null));
      if (unx(it.attrs.n) !== t) ((def.itemCaptions ??= {})[names[f]] ??= {})[t] = unx(it.attrs.n);
    }
  });
  // 행 · 열 · 필터 영역에 없는 필드의 숨긴 항목은 엑셀이 적용하지 않음 (필드를 빼면 필터도 풀림).
  // 단, 그 필드의 슬라이서가 이 피벗을 거르면 적용됨 → 슬라이서를 연결한 뒤에 정함 (linkPivotsAndSlicers)
  const onAxis = new Set([...def.rows, ...def.cols, ...(def.pages ?? [])]);
  const offAxis = Object.keys(filters).filter((f) => !onAxis.has(f));
  if (offAxis.length) Object.defineProperty(def, '_offAxis', { value: offAxis, enumerable: false, configurable: true });
  // 이름 · 스타일 · 캡션 · 계산 필드
  if (root.attrs.name) def.name = root.attrs.name;
  const wxCalc = descendants(child(root, 'extLst'), 'calcItems')[0]?.attrs.json;
  if (wxCalc) { try { def.calcItems = JSON.parse(wxCalc); } catch { /* 잘못된 확장 무시 */ } }
  const si0 = child(root, 'pivotTableStyleInfo');
  def.style = si0?.attrs.name ?? 'None';
  const so = {
    rowHeaders: si0?.attrs.showRowHeaders !== '0', colHeaders: si0?.attrs.showColHeaders !== '0',
    bandRows: si0?.attrs.showRowStripes === '1', bandCols: si0?.attrs.showColStripes === '1',
  };
  if (!so.rowHeaders || !so.colHeaders || so.bandRows || so.bandCols) def.styleOpts = so;
  if (root.attrs.rowHeaderCaption !== undefined) def.rowCaption = root.attrs.rowHeaderCaption;
  if (root.attrs.grandTotalCaption !== undefined) def.grandCaption = root.attrs.grandTotalCaption;
  def.errorShow = root.attrs.showError === '1' || root.attrs.showError === 'true';
  if (root.attrs.errorCaption !== undefined || def.errorShow) def.errorCaption = root.attrs.errorCaption ?? '';
  if (root.attrs.colHeaderCaption !== undefined) def.colCaption = root.attrs.colHeaderCaption;
  if (root.attrs.missingCaption && root.attrs.showMissing !== '0') def.missingCaption = root.attrs.missingCaption;
  if (falseAttr(root.attrs.showHeaders)) def.showHeaders = false;
  if (falseAttr(root.attrs.showDrill)) def.showExpand = false;
  if (trueAttr(root.attrs.mergeItem)) def.mergeLabels = true;
  if (falseAttr(root.attrs.useAutoFormatting)) def.autofit = false;
  if (falseAttr(root.attrs.preserveFormatting)) def.preserveFormat = false;
  if (falseAttr(root.attrs.enableDrill)) def.enableDrill = false;
  if (falseAttr(root.attrs.customListSort)) def.customListSort = false;
  if (trueAttr(root.attrs.multipleFieldFilters)) def.multiFilters = true;
  // 축소한 항목 · 부분합 위치 · 빈 줄 · 레이블 반복
  const collapsed = {};
  pfs.forEach((pf, f) => {
    if (![...rowF, ...colF].includes(f)) return;
    const its = kids(child(pf, 'items'), 'item').filter((it) => it.attrs.sd === '0' && it.attrs.x !== undefined);
    if (its.length) collapsed[names[f]] = its.map((it) => itemText(keyOf(cache.fields[f]?.items[Number(it.attrs.x)] ?? null)));
    if (pf.attrs.subtotalTop === '0' && def.layout !== 'tabular') def.subtotalTop = false;
    // 빈 줄 삽입은 필드마다 (엑셀: 바깥 필드에만 켜 두는 경우가 많음)
    if (rowF.includes(f) && pf.attrs.insertBlankRow === '1') (def._blankFields ??= []).push(names[f]);
    if (rowF.includes(f) && descendants(pf, 'pivotField').some((x) => x.attrs.fillDownLabels === '1')) def.repeatLabels = true;
  });
  if (Object.keys(collapsed).length) def.collapsed = collapsed;
  if (def._blankFields) {
    const inner = rowF.slice(0, -1).map((f) => names[f]);
    def.blankRows = inner.every((n) => def._blankFields.includes(n)) ? true : def._blankFields;
    delete def._blankFields;
  }
  // 날짜 · 숫자 그룹 (파생 필드 '월2' 등은 base 필드에서 만듦)
  const groups = {};
  // 그룹은 캐시에 있어 같은 캐시를 쓰는 피벗 모두에 적용됨: 축에 없어도 그 필드로 거르면(슬라이서 · 숨긴 항목 '1월'…) 필요
  const filtered = new Set(Object.keys(def.filters ?? {}).map((n) => n.toLowerCase()));
  const groupFields = new Set([...rowF, ...colF, ...pageEls.map((p) => Number(p.attrs.fld))]);
  names.forEach((n, f) => { if (filtered.has(String(n).toLowerCase())) groupFields.add(f); });
  for (const f of groupFields) {
    const g = cache.fields[f]?.group;
    if (!g) continue;
    if (g.by === 'discrete') {
      const bf = cache.fields[g.base];
      if (!bf) continue;
      const gi = cache.fields[f].items ?? [];
      const map = {};
      (bf.shared ?? []).forEach((k, i) => { const to = gi[g.disc[i]]; if (to !== undefined && to !== null && itemText(to) !== itemText(k)) map[itemText(k)] = itemText(to); });
      groups[names[f]] = { by: 'items', base: bf.name, map };
      continue;
    }
    const by = { years: 'years', quarters: 'quarters', months: 'months', days: 'mdays' }[g.by] ?? (g.by === 'range' && g.size ? 'number' : null);
    if (!by) continue;
    const spec = by === 'number' ? { by, start: g.startNum, size: g.size } : { by, ...(g.start !== undefined ? { start: g.start, end: g.end } : {}) };
    if (g.derived && g.base !== null && cache.fields[g.base]) spec.base = cache.fields[g.base].name;
    groups[names[f]] = spec;
  }
  if (Object.keys(groups).length) def.groups = groups;
  const calc = cache.fields.filter((f) => f.formula !== undefined).map((f) => ({ name: f.name, formula: f.formula }));
  if (calc.length) def.calcFields = calc;
  const ff0 = new Set(kids(child(root, 'filters'), 'filter').map((flt) => Number(flt.attrs.fld)));
  // 정렬 · 항목 순서
  const sort = {};
  const order = {};
  pfs.forEach((pf, f) => {
    if (![...rowF, ...colF].includes(f)) return;
    const st = pf.attrs.sortType;
    if (st === 'ascending' || st === 'descending') {
      const ref = descendants(child(pf, 'autoSortScope'), 'reference').find((r) => r.attrs.field === '4294967294' || r.attrs.field === '-2');
      const by = ref ? Number(child(ref, 'x')?.attrs.v ?? 0) : undefined;
      // 다른 축의 특정 항목 값으로 정렬 (예: 열 '월' 의 12 열 값 기준) → at: [[필드, 항목 글자]]
      const at = [];
      for (const r of descendants(child(pf, 'autoSortScope'), 'reference')) {
        const f2 = Number(r.attrs.field);
        if (!(f2 >= 0) || f2 >= pfs.length || r.attrs.field === '4294967294') continue;
        const pit = kids(child(pfs[f2], 'items'), 'item')[Number(child(r, 'x')?.attrs.v ?? -1)];
        if (pit?.attrs.x !== undefined) at.push([names[f2], itemText(keyOf(cache.fields[f2]?.items[Number(pit.attrs.x)] ?? null))]);
      }
      sort[names[f]] = { dir: st === 'ascending' ? 'asc' : 'desc', ...(by !== undefined ? { by } : {}), ...(at.length ? { at } : {}) };
    }
    const its = kids(child(pf, 'items'), 'item').filter((it) => it.attrs.x !== undefined && !it.attrs.t);
    if (its.length > 1) order[names[f]] = its.map((it) => itemText(keyOf(cache.fields[f]?.items[Number(it.attrs.x)] ?? null)));
  });
  if (Object.keys(sort).length) def.sort = sort;
  if (Object.keys(order).length) def.order = order;
  // 값 기준 정렬 · 상위 N 에서 값이 같은 항목의 순서: 엑셀이 저장한 표시 순서(rowItems · colItems)를 그대로 따름
  const tie = {};
  const tieSeen = {};
  // 안쪽 필드는 상위 항목마다 순서가 다를 수 있음 (같은 검색어가 광고그룹마다 다른 자리) → 상위 항목 글자 경로별 목록도
  const tieByParent = {};
  const itemsMemo = new Map();
  const fieldItems = (f) => { let a = itemsMemo.get(f); if (!a) { a = kids(child(pfs[f], 'items'), 'item'); itemsMemo.set(f, a); } return a; };
  for (const [tag, fl] of [['rowItems', rowF], ['colItems', colAll]]) {
    const fieldsAt = fl.filter((x) => x >= 0 || x === -2);
    const last = [];
    const lastText = [];
    for (const it of kids(child(root, tag), 'i')) {
      if (it.attrs.t) continue;
      const base = Number(it.attrs.r ?? 0);
      kids(it, 'x').forEach((x, j) => {
        const lv = base + j;
        last[lv] = Number(x.attrs.v ?? 0);
        const f = fieldsAt[lv];
        const pit = f !== undefined && f >= 0 ? fieldItems(f)[last[lv]] : null;
        const text = pit && pit.attrs.x !== undefined ? itemText(keyOf(cache.fields[f]?.items[Number(pit.attrs.x)] ?? null)) : null;
        lastText[lv] = text;
        if (text === null || !(sort[names[f]] || ff0.has(f))) return;
        const seen = (tieSeen[names[f]] ??= new Set());
        if (!seen.has(text)) { seen.add(text); (tie[names[f]] ??= []).push(text); }
        if (lv > 0) {
          const pk = lastText.slice(0, lv).filter((t) => t !== null).join('\u0001');
          ((tieByParent[names[f]] ??= {})[pk] ??= []).push(text);
        }
      });
    }
  }
  if (Object.keys(tieByParent).length) def.tieByParent = tieByParent;
  if (Object.keys(tie).length) {
    def.tieOrder = tie;
    // 저장할 때의 필터 상태: 같은 선택이면 상위 N 항목을 엑셀이 저장한 그대로 (오류 값 항목 선택 등 엑셀 고유 동작)
    def.tieState = pivotFilterKey(def.filters ?? {});
  }
  // 레이블 · 값 · 상위 10 필터
  const ff = {};
  for (const flt of kids(child(root, 'filters'), 'filter')) {
    const f = Number(flt.attrs.fld);
    const name = names[f];
    if (!name) continue;
    const type = flt.attrs.type ?? '';
    const by = Number(flt.attrs.iMeasureFld ?? 0);
    const t10 = descendants(flt, 'top10')[0];
    if (['count', 'percent', 'sum'].includes(type)) {
      ff[name] = { type: 'top', top: t10?.attrs.top !== '0', n: Number(t10?.attrs.val ?? flt.attrs.intValue ?? 10), mode: type, by };
    } else if (/^caption/.test(type)) {
      const op = type.replace(/^caption/, '');
      ff[name] = { type: 'label', op: op[0].toLowerCase() + op.slice(1), v1: flt.attrs.stringValue1 ?? '', v2: flt.attrs.stringValue2 ?? '' };
    } else if (DATE_OP_TYPES.has(type)) {
      // 날짜 필터: 값은 일련번호 (customFilter val) 또는 ISO 날짜 (stringValue)
      const cf = descendants(flt, 'customFilter').map((x) => Number(x.attrs.val)).filter(Number.isFinite);
      const toSerial = (v) => (v === undefined || v === '' ? undefined : Number.isFinite(Number(v)) ? Number(v) : isoSerial(v, cache.date1904));
      ff[name] = { type: 'date', op: type, v1: cf[0] ?? toSerial(flt.attrs.stringValue1), v2: cf[1] ?? toSerial(flt.attrs.stringValue2) };
      for (const k of ['v1', 'v2']) if (ff[name][k] === undefined) delete ff[name][k];
    } else if (/^value/.test(type)) {
      const op = type.replace(/^value/, '');
      ff[name] = { type: 'value', op: op[0].toLowerCase() + op.slice(1), v1: flt.attrs.stringValue1 ?? '', v2: flt.attrs.stringValue2 ?? '', by };
    }
  }
  if (Object.keys(ff).length) def.fieldFilters = ff;
  def.captureFmt = true;
  const dc = root.attrs.dataCaption;
  if (dc !== undefined && dc !== '값' && dc !== 'Values') def.dataCaption = unx(dc);
  const locEl = child(root, 'location');
  const loc = refToRange(locEl?.attrs.ref ?? '');
  // Excel 2010+의 값 행 옵션은 위치가 아니라 x14 확장에 저장된다.
  // classic=true는 hideValuesRow=true여도 헤더를 표시하므로 위치만 보면 체크 상태를 잃는다.
  const displayExt = kids(child(root, 'extLst'), 'ext').find((e) => e.attrs.uri?.toUpperCase() === '{962EF5D1-5CA2-4C93-8EF4-DBF5C05439D2}');
  const hideValuesRow = child(displayExt, 'pivotTableDefinition')?.attrs.hideValuesRow;
  def.showValuesRow = hideValuesRow !== '1' && hideValuesRow !== 'true'; // OOXML 기본값: 표시
  if (loc) {
    const page = pivotPageLayout(def);
    const pageRows = page.height ? page.height + 1 : 0;
    const top = Math.max(0, loc.r1 - pageRows);
    const reportOnly = !def.rows.length && !def.cols.length && !def.values.length && page.height > 0;
    // location은 필터를 제외한 본문이다. 필터 전용 피벗도 빈 본문 한 칸을 앵커로 저장한다.
    const area = reportOnly
      ? { r1: top, c1: loc.c1, r2: top + page.height - 1, c2: loc.c1 + page.width - 1 }
      : { ...loc, r1: top, c2: Math.max(loc.c2, loc.c1 + page.width - 1) };
    Object.assign(def, { top, left: loc.c1, area });
  }
  return def;
}

/** 레이아웃 밖 필드의 숨긴 항목 필터를 지움 — 그 필드의 슬라이서가 거르는 피벗은 남김 */
function dropOffAxisFilters(sheets) {
  const sliced = new Set();
  for (const sh of sheets) {
    for (const sl of sh.slicers ?? []) {
      if (sl.source?.kind !== 'pivot') continue;
      for (const p of sl.source.pivots ?? []) sliced.add(`${p.sheet}\u0001${p.name}\u0001${sl.source.field}`);
    }
  }
  for (const sh of sheets) {
    for (const def of [sh.pivot, ...(sh.pivotsExtra ?? [])]) {
      if (!def?._offAxis) continue;
      for (const f of def._offAxis) if (!sliced.has(`${sh.name}\u0001${def.name}\u0001${f}`)) delete def.filters[f];
      if (def.filters && !Object.keys(def.filters).length) delete def.filters;
      delete def._offAxis;
    }
  }
}

/** slicerCacheDefinition → { name, sourceName, table: { tableId, column } | null, pivot: { tabId, name } | null } */
function readSlicerCache(files, path) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const tsc = descendants(root, 'tableSlicerCache')[0];
  const pts = descendants(child(root, 'pivotTables'), 'pivotTable').map((pt) => ({ tabId: Number(pt.attrs.tabId), name: pt.attrs.name }));
  // 슬라이서 설정: 정렬 · 사용자 지정 목록 · 데이터 없는 항목
  const tabular = descendants(root, 'tabular')[0];
  const a = { ...(tabular?.attrs ?? {}), ...(tsc?.attrs ?? {}) };
  const opts = {};
  if (a.sortOrder === 'descending') opts.sort = 'desc';
  if (a.customListSort === '0' || a.customListSort === 'false') opts.customList = false;
  // MS-XLSX CT_TabularSlicerCache: showMissing의 생략 기본값은 true입니다.
  if (tabular && a.showMissing !== '0' && a.showMissing !== 'false') opts.showDeleted = true;
  if (a.crossFilter === 'none') opts.markNoData = false;
  if (a.crossFilter === 'showItemsWithNoData') opts.noDataLast = false;
  if (descendants(root, 'slicerCacheHideItemsWithNoData').length) opts.hideNoData = true;
  return {
    opts, olap: !!descendants(root, 'olap').length,
    cacheId: tabular?.attrs.pivotCacheId,
    items: kids(child(tabular, 'items'), 'i').map(it => ({ index: Number(it.attrs.x), selected: ['1','true'].includes(it.attrs.s), hasData: !['1','true'].includes(it.attrs.nd) })),
    name: root.attrs.name, sourceName: root.attrs.sourceName,
    table: tsc ? { tableId: Number(tsc.attrs.tableId), column: Number(tsc.attrs.column) } : null,
    pivot: pts[0] ?? null,
    pivots: pts,
  };
}

/** 피벗 테이블 정의와 슬라이서를 시트에 연결하고 임시 속성은 지움 */
/**
 * 엑셀 피벗 조건부 서식 (<conditionalFormats><conditionalFormat scope="data|field|selection" priority>) →
 * 시트 규칙에 pivot: { name, scope, value, rowField } — 피벗을 다시 그리면 범위가 따라감
 */
function linkPivotCond(sheet, def, root, cache) {
  const cfs = kids(child(root, 'conditionalFormats'), 'conditionalFormat');
  if (!cfs.length || !sheet._condPrio) return;
  const names = cache.fields.map((f, i) => f.name || `열${i + 1}`);
  for (const cf of cfs) {
    const k = sheet._condPrio.indexOf(Number(cf.attrs.priority));
    const rule = k >= 0 ? sheet.cond[k] : null;
    if (!rule) continue;
    const scope = cf.attrs.scope === 'data' ? 'data' : cf.attrs.scope === 'field' ? 'field' : 'selection';
    let vi = 0;
    let rowField = null;
    for (const ref of descendants(cf, 'reference')) {
      const f = Number(ref.attrs.field);
      if (f === 4294967294 || f === -2) vi = Number(child(ref, 'x')?.attrs.v ?? 0);
      else if (names[f] && (def.rows ?? []).includes(names[f])) rowField = names[f];
    }
    const v = def.values?.[vi];
    rule.pivot = { name: def.name ?? '', scope, ...(v ? { value: valueName(v) } : {}), ...(rowField ? { rowField } : {}) };
  }
}

function linkPivotsAndSlicers(files, wbRels, sheets, ctx) {
  const tables = sheets.flatMap((s) => s.tables.map((t) => ({ ...t, sheetName: s.name })));
  const caches = new Map();
  for (const rel of Object.values(wbRels)) {
    if (rel.type !== 'slicerCache') continue;
    const c = readSlicerCache(files, rel.target);
    if (c?.olap) warnOlapImport(ctx);
    if (c?.name) caches.set(c.name, c);
  }
  const cacheFormat = numFmtId => {
    if (!numFmtId) return null;
    if (!ctx.cacheNumberFormats) {
      const path = Object.values(wbRels).find(r => r.type === 'styles')?.target;
      ctx.cacheNumberFormats = Object.fromEntries(kids(child(parseXml(textOf(files[path]) || ''), 'numFmts'), 'numFmt').map(f => [f.attrs.numFmtId, f.attrs.formatCode]));
    }
    const code = ctx.cacheNumberFormats[numFmtId] ?? BUILTIN_CODE[numFmtId];
    return code ? styleForCode(code) : BUILTIN_FMT[numFmtId] ?? null;
  };
  const cacheFiles = new Map();
  sheets.forEach((s) => {
    for (const p of s._pivots) {
      if (!p.cachePath) continue;
      if (!cacheFiles.has(p.cachePath)) cacheFiles.set(p.cachePath, readPivotCache(files, p.cachePath, ctx.date1904));
      const cache = cacheFiles.get(p.cachePath);
      if (cache?.olap) warnOlapImport(ctx);
      const def = cache && pivotDefFrom(p.root, cache, tables, s.name);
      if (def) {
        // 레코드에 없는 과거 항목도 필터·슬라이서 선택의 원본 번호를 유지합니다.
        def.cacheItemsId = p.cachePath;
        const localSource = def.table ? tables.some(t => t.name.toLowerCase() === String(def.table).toLowerCase())
          : sheets.some(sheet => sheet.name.toLowerCase() === String(def.source ?? '').toLowerCase());
        if (cache.source.external || !localSource) def.sourceReference = makeImportedPivotSourceReference(cache.source, def);
        // A pivot field's number format is independent of its cache field and
        // other pivots using that cache. Keep explicit General and native IDs.
        const fieldFormats = kids(child(p.root, 'pivotFields'), 'pivotField').flatMap((field, i) => {
          if (field.attrs.numFmtId === undefined || !cache.fields[i]) return [];
          const id = Number(field.attrs.numFmtId);
          if (!Number.isInteger(id) || id < 0) return [];
          const known = cacheFormat(id); // Also initializes the custom format map.
          const code = ctx.cacheNumberFormats?.[id];
          const format = code !== undefined ? { numFmt: 'custom', code }
            : id < 164 ? { ...(known ?? { numFmt: 'general' }), xlsxBuiltinId: id } : null;
          return format ? [[cache.fields[i].name || `열${i + 1}`, format]] : [];
        });
        if (fieldFormats.length) def.fieldNumberFormats = Object.fromEntries(fieldFormats);
        const retained = (ctx.pivotCacheItems ??= {});
        retained[p.cachePath] ??= { ...(cache.missingItems !== undefined ? { missingItemsLimit: pivotMissingItemsLimit(cache) } : {}), fields: cache.fields.filter(f => f.db).map(f => ({ name: f.name, shared: f.shared, ...(f.sharedTypes ? { sharedTypes: f.sharedTypes } : {}), ...(f.dateOnly ? { dateOnly: true } : {}), ...(cacheFormat(f.numFmtId) ? { format: cacheFormat(f.numFmtId) } : {}) })) };
      }
      // 엑셀의 저장본(캐시 레코드)으로 처음 화면을 그림 — 원본을 고치거나 새로 고치면 원본에서 다시 계산
      if (def && cache.snapshot && !cache.snapshot.failed) {
        if (!(ctx.pivotSnapshots ??= {})[p.cachePath]) {
          const rows = readCacheRecords(files, cache);
          if (rows) ctx.pivotSnapshots[p.cachePath] = rows; else cache.snapshot.failed = true;
        }
        if (ctx.pivotSnapshots[p.cachePath]) def.snapshotId = p.cachePath;
      }
      if (def?.snapshotId || def?.sourceReference && def.saveData === false) {
        const nodeXml = node => {
          if (!node) return '';
          let body = node.text ? esc(node.text) : '';
          for (const child of node.children ?? []) body += nodeXml(child);
          const attrs = Object.entries(node.attrs ?? {}).map(([key, value]) => ` ${key}="${esc(value)}"`).join('');
          return body ? `<${node.name}${attrs}>${body}</${node.name}>` : `<${node.name}${attrs}/>`;
        };
        const layout = {
          location: { ...child(p.root, 'location')?.attrs }, rowItems: nodeXml(child(p.root, 'rowItems')), colItems: nodeXml(child(p.root, 'colItems')),
          top: def.top ?? 0, left: def.left ?? 0, binding: pivotCacheLayoutBinding(def),
        };
        if (!def.snapshotId) def.sourceReference.pivotLayout = layout;
        else {
          const fieldName = f => f === -2 ? null : cache.fields[f]?.name || `열${f + 1}`;
          const axes = tag => kids(child(p.root, tag), 'field').map(field => fieldName(Number(field.attrs.x)));
          const rows = axes('rowFields'), cols = axes('colFields'), pfs = kids(child(p.root, 'pivotFields'), 'pivotField');
          const itemOrder = [], fieldCounts = new Map();
          for (let f=0;f<cache.fields.length;f++) { const name=fieldName(f);fieldCounts.set(name,(fieldCounts.get(name)??0)+1); }
          for (let f = 0; f < cache.fields.length; f++) {
            const name = fieldName(f);
            if (!rows.includes(name) && !cols.includes(name)) continue;
            const entries = child(pfs[f], 'items')?.children ?? [];
            const tokens = function* () {
              for (const entry of entries) if (entry.name === 'item') {
                if (entry.attrs.x !== undefined) {
                  const x=Number(entry.attrs.x),values=cache.fields[f]?.items;
                  yield Number.isSafeInteger(x)&&x>=0&&x<(values?.length??0) ? 'v:' + itemIdentity(keyOf(values[x])) : '!invalid:' + entry.attrs.x;
                }
                else yield 't:' + (entry.attrs.t ?? 'data');
              }
            };
            itemOrder.push({ field:name, signature:fieldCounts.get(name)===1?pivotItemOrderSignature(tokens()):'!ambiguous-field' });
          }
          const body=refToRange(layout.location.ref ?? '');
          const cellAt=(r,c)=>{
            const cell=s.cells.getRC?s.cells.getRC(r,c):s.cells[r+','+c];
            if(cell)return cell;
            for(const b of s.blocks??[])if(inBlock(b,r,c)){
              const v=blockValue(b,r,c),fmt=b.cols[c-b.c0].fmt??undefined;
              return v===null?undefined:{raw:rawOf(v,fmt,textRaw),v,style:fmt,...(fmt?.numFmt==='text'?{inputType:'value'}:{})};
            }
          };
          const toggles=importedPivotToggleButtons({root:p.root,cache,def,body,cellAt}).map(button=>({...button,raw:cellAt(button.r,button.c).raw}));
          def.importedPresentation = { ...layout, kind:'xlsx-pivot-output', body, axes:{ rows, cols }, itemOrder, ...(toggles.length?{toggles}:{}) };
        }
      }
      if (def && findObjectStyle(ctx.objectStyles,'pivot',def.style)) Object.assign(def,objectStylePatch('pivot',findObjectStyle(ctx.objectStyles,'pivot',def.style)));
      else if (def && ctx.tableStyles?.[def.style] && !PRESET_STYLES[def.style]) def.styleDef = ctx.tableStyles[def.style]; // 파일에 정의된 사용자 지정 스타일 (WIXEL 모던 스타일은 이름으로 앎)
      if (def) linkPivotCond(s, def, p.root, cache);
      if (!def) ctx.warnings.add('외부 데이터 원본을 쓰는 피벗 테이블은 값으로만 가져왔습니다.');
      else if (!s.pivot) { s.pivot = def; s._pivotName = p.root.attrs.name; } else (s.pivotsExtra ??= []).push(def);
    }
  });
  sheets.forEach((s) => {
    const boxes = s._slicerBoxes ?? {};
    let i = 0;
    for (const sl of s._slicers) {
      const c = caches.get(sl.cache);
      let source = null;
      if (c?.table) {
        const t = tables.find((x) => x._xmlId === c.table.tableId);
        const col = t?._colNames[c.table.column - 1];
        if (t && col) source = { kind: 'table', table: t.name, column: col };
      } else if (c?.pivots?.length) {
        // 슬라이서 하나가 여러 피벗 테이블(다른 시트 포함)을 함께 거름
        const hasPivot = (x, name) => [x.pivot, ...(x.pivotsExtra ?? [])].some((d) => d && d.name === name);
        const pivots = c.pivots.map((p) => {
          const ps = sheets.find((x) => x._sheetId === p.tabId && hasPivot(x, p.name)) ?? sheets.find((x) => hasPivot(x, p.name));
          return ps ? { sheet: ps.name, name: p.name } : null;
        }).filter(Boolean);
        if (pivots.length) source = { kind: 'pivot', field: c.sourceName, pivots };
      }
      let cacheSelection;
      if (!source && c?.cacheId !== undefined) {
        // 피벗 연결이 0개여도 x14 캐시가 남아 있는 독립 슬라이서는 항목·선택을 보존한다.
        // 원본은 외부 파일일 수 있으므로 읽거나 같은 이름의 다른 피벗에 추정 연결하지 않는다.
        let standalone;
        for (const rel of Object.values(wbRels)) {
          if (rel.type !== 'pivotCacheDefinition') continue;
          if (!cacheFiles.has(rel.target)) cacheFiles.set(rel.target, readPivotCache(files, rel.target, ctx.date1904));
          const candidate = cacheFiles.get(rel.target);
          if (String(candidate?.slicerId) === String(c.cacheId)) { standalone = { cache: candidate, path: rel.target }; break; }
        }
        const field = standalone?.cache.fields.find(f => f.name.toLowerCase() === String(c.sourceName).toLowerCase());
        if (field) {
          source = { kind: 'cache', field: field.name, cacheKey: c.name, cacheSource: standalone.cache.source, values: field.items, numFmtId: field.numFmtId, items: c.items.filter(it => Number.isInteger(it.index) && it.index >= 0 && it.index < field.items.length).map(({index,hasData}) => ({index,hasData})) };
          const validItems = new Set(source.items.map(it => it.index));
          cacheSelection = c.items.filter(it => it.selected && validItems.has(it.index)).map(it => String(it.index));
          if (!ctx.cacheNumberFormats) {
            const path = Object.values(wbRels).find(r => r.type === 'styles')?.target;
            ctx.cacheNumberFormats = Object.fromEntries(kids(child(parseXml(textOf(files[path]) || ''), 'numFmts'), 'numFmt').map(f => [f.attrs.numFmtId, f.attrs.formatCode]));
          }
          const code = ctx.cacheNumberFormats[field.numFmtId] ?? BUILTIN_CODE[field.numFmtId];
          source.format = code ? styleForCode(code) : BUILTIN_FMT[field.numFmtId] ?? {};
          if (cacheSelection.length === source.items.length) cacheSelection = null;
        }
      }
      if (!source) { ctx.warnings.add('연결 대상을 찾지 못한 슬라이서는 가져오지 않았습니다.'); continue; }
      const box = boxes[sl.name] ?? { x: 20 + i * 190, y: 20, w: 180, h: 200 };
      const styleName=sl.style??ctx.objectStyles?.defaultSlicerStyle??'SlicerStyleLight1';
      s.slicers.push({
        id: `sl${Math.random().toString(36).slice(2, 9)}`, caption: sl.caption ?? sl.name, source, columns: Math.max(1, sl.columns || 1),
        style: styleName === 'WIXEL Slicer None' ? 'None' : styleName, multi: false, ...box, ...(sl.startItem ? { startItem: sl.startItem } : {}),
        ...(source.kind === 'cache' ? { cacheSelection } : {}),
        ...(styleName === 'WIXEL Slicer None' ? {} : findObjectStyle(ctx.objectStyles,'slicer',styleName) ? objectStylePatch('slicer',findObjectStyle(ctx.objectStyles,'slicer',styleName)) : !isModernSlicer(styleName) && ctx.slicerStyles?.[styleName] && Object.keys(ctx.slicerStyles[styleName]).length ? { custom: ctx.slicerStyles[styleName] } : {}),
        ...(sl.showCaption ? {} : { showHeader: false }),
        ...(sl.lockedPosition ? { noMove: true } : {}),
        ...(c.opts ?? {}),
        ...(Number.isFinite(sl.rowHeight) && sl.rowHeight > 0 ? { buttonHeight: sl.rowHeight / EMU } : {}),
      });
      i++;
    }
  });
  for (const s of sheets) {
    delete s._slicers; delete s._pivots; delete s._slicerBoxes; delete s._sheetId; delete s._pivotName; delete s._condPrio;
    for (const t of s.tables) { delete t._xmlId; delete t._colNames; }
  }
}

/** xlsx 바이트 → 통합 문서 데이터 ({ sheets: [...] }, 경고 목록) */
/** .xlsx 바이트 → { data, active, warnings } (한 번에) */
export function readXlsx(bytes) {
  const it = readXlsxSteps(unzip(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes)));
  for (;;) {
    const s = it.next();
    if (s.done) return s.value;
  }
}

/** 큰 파일용: 중간중간 브라우저에 제어를 돌려주며 읽음. onProgress(0~1, 설명) */
export async function readXlsxAsync(bytes, onProgress, { onDiagnostics, streamThreshold = 4 << 20 } = {}) {
  if (!Number.isSafeInteger(streamThreshold) || streamThreshold < 0) throw new RangeError('시트 스트림 기준이 올바르지 않습니다.');
  onProgress?.({ p: 0, msg: '압축 푸는 중' });
  // 필요한 항목을 읽기 직전에 하나씩 해제하여 모든 시트 원문을 동시에 보유하지 않습니다.
  const files = typeof Blob !== 'undefined' && bytes instanceof Blob
    ? await unzipBlob(bytes, { onProgress: status => onProgress?.({ p: 0.04 * status.p, msg: '파일 목록 읽는 중' }) })
    : unzip(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const streamParts = new Set(Object.keys(files).filter(path => /(?:worksheets\/[^/]+\.xml|pivotCacheRecords[^/]+\.xml|sharedStrings\.xml)$/i.test(path) && (zipEntryInfo(files, path)?.size ?? 0) >= streamThreshold));
  const it = readXlsxSteps(files, onDiagnostics, streamParts);
  let last = performance.now();
  for (;;) {
    const s = it.next();
    if (s.done) return s.value;
    if (s.value.preparePart) {
      onProgress?.({ p: s.value.p, msg: s.value.msg });
      if (streamParts.has(s.value.preparePart)) await prepareZipEntryRaw(files, s.value.preparePart);
      else await prepareZipEntry(files, s.value.preparePart);
      onProgress?.({ p: s.value.p, msg: s.value.msg });
      continue;
    }
    const now = performance.now();
    if (now - last > 40) {
      onProgress?.(s.value);
      await new Promise((res) => setTimeout(res, 0));
      last = performance.now();
    }
  }
}

function* readXlsxSteps(files, onDiagnostics, streamParts = null) {
  yield { p: 0.05, msg: '압축 푸는 중' };
  // xlsb: 바이너리 부분을 같은 경로의 xlsx XML 로 바꾼 뒤 그대로 읽음 (큰 시트의 셀은 행 생성기로 바로)
  if (isXlsb(files)) yield* convertXlsb(files, { lazySheets: true });
  const wbPath = Object.keys(files).find((f) => /^xl\/workbook\.xml$/i.test(f))
    ?? relsTarget(files, '', 'officeDocument');
  if (!wbPath || !files[wbPath]) throw new Error('엑셀 통합 문서(.xlsx)가 아닙니다');
  const wbRoot = parseXml(textOf(files[wbPath]));
  const wbRels = relsOf(files, wbPath);
  const theme = readTheme(files, wbRels);
  const { xfs, dxfs, filterDxfs, tableDxfs, dxfOf, tableStyles, wbFont, slicerStyles, cellStyles, fonts, objectStyles } = readStyles(files, wbRels, theme);
  const mdw = digitWidth(wbFont);
  const ssRel = Object.values(wbRels).find((r) => r.type === 'sharedStrings');
  let strings = files.__xlsb?.strings ?? [], phonetics = [];
  if (!files.__xlsb && ssRel && Object.hasOwn(files, ssRel.target)) {
    yield { preparePart: ssRel.target, p: 0.05, msg: '공유 문자열 읽는 중' };
    ({ strings, phonetics } = yield* progressOf(readSharedStrings(streamParts?.has(ssRel.target) ? zipEntryChunks(files, ssRel.target) : files[ssRel.target], fonts), (count) => ({ p: 0.05, msg: `공유 문자열 읽는 중 (${count.toLocaleString()}개)` })));
    deleteZipEntry(files, ssRel.target);
  }
  // 이름 정의 (시트 범위 이름은 localSheetId → 시트 이름)
  const allSheetNames = kids(child(wbRoot, 'sheets'), 'sheet').map((sh) => sh.attrs.name.slice(0, 31));
  const names = [];
  const printNames = []; // 인쇄 영역 · 인쇄 제목 (시트를 읽은 뒤 page 에 넣음)
  for (const dn of kids(child(wbRoot, 'definedNames'), 'definedName')) {
    const name = dn.attrs.name;
    const text = (dn.text ?? '').trim();
    if (/^_xlnm\.(Print_Area|Print_Titles)$/i.test(name) && dn.attrs.localSheetId !== undefined) printNames.push({ kind: /Area/i.test(name) ? 'area' : 'titles', sheet: allSheetNames[Number(dn.attrs.localSheetId)], text });
    if (!name || /^_xlnm\.|^_xlfn\./i.test(name) || !text) continue;
    if (text === '#N/A' && /^(Slicer_|슬라이서_)/i.test(name)) continue;
    const sheet = dn.attrs.localSheetId !== undefined ? allSheetNames[Number(dn.attrs.localSheetId)] ?? null : null;
    names.push({ name, ref: `=${fromFileFormula(text)}`, sheet, ...(dn.attrs.comment ? { comment: dn.attrs.comment } : {}), ...(dn.attrs.hidden === '1' || dn.attrs.hidden === 'true' ? { hidden: true } : {}) });
  }
  const nameSet = new Set(names.map((n) => n.name.toUpperCase()));
  const isName = (n) => nameSet.has(String(n).toUpperCase());
  const nameMulti = (n) => {
    const e = names.find((x) => x.name.toUpperCase() === String(n).toUpperCase());
    if (!e) return false;
    try { return mayReturnArray(parse(e.ref.slice(1))); } catch { return false; }
  };
  const date1904 = ['1', 'true'].includes(child(wbRoot, 'workbookPr')?.attrs.date1904);
  const ctx = { onDiagnostics, streamParts, mdw, wbFont, xfs, dxfs, filterDxfs, tableDxfs, dxfOf, tableStyles, slicerStyles, objectStyles, strings, phonetics, fonts, theme, date1904, warnings: new Set(), isName, nameMulti, richImages: readRichImages(files, wbRels) };
  // 모델 본체는 거대한 바이너리입니다. 경고를 위한 존재 검사만 하고 압축을 해제하지 않습니다.
  if (Object.keys(files).some(path => /^xl\/model\/[^/]+\.(?:data|bin)$/i.test(path)) || Object.values(wbRels).some(rel => rel.type === 'model')) warnOlapImport(ctx);
  const sheets = [];
  const warnings = [];
  const sheetCodes = {};
  let unsupported = 0;
  for (const sh of kids(child(wbRoot, 'sheets'), 'sheet')) {
    const rel = wbRels[rid(sh)];
    if (!rel || rel.type !== 'worksheet' || !Object.hasOwn(files, rel.target)) {
      if (rel) warnings.push(`'${sh.attrs.name}' 시트(차트 시트 등)는 가져오지 않았습니다.`);
      continue;
    }
    const si = sheets.length;
    const total = kids(child(wbRoot, 'sheets'), 'sheet').length;
    yield { preparePart: rel.target, p: 0.05 + 0.85 * (si / total), msg: `'${sh.attrs.name}' 시트 준비 중` };
    files.__xlsb?.prepareSheet(rel.target);
    const sheet = yield* progressOf(readSheet(files, rel.target, ctx), (rows) => ({ p: 0.05 + 0.85 * (si / total), msg: `'${sh.attrs.name}' 시트 읽는 중 (${rows.toLocaleString()}행)` }));
    deleteZipEntry(files, rel.target);
    files.__xlsb?.rows.delete(rel.target);
    if (sheet.codeName) sheetCodes[sh.attrs.name.slice(0, 31)] = sheet.codeName;
    delete sheet.codeName;
    unsupported += sheet.unsupported;
    delete sheet.unsupported;
    sheets.push({ name: sh.attrs.name.slice(0, 31), ...sheet, _sheetId: Number(sh.attrs.sheetId), ...(sh.attrs.state === 'hidden' || sh.attrs.state === 'veryHidden' ? { state: sh.attrs.state } : {}) });
  }
  for (const pn of printNames) {
    const sh = sheets.find((x) => x.name === pn.sheet);
    if (!sh) continue;
    const page = normPage(sh.page);
    const areas = parsePrintAreas(pn.text, sh.name);
    if (pn.kind === 'area' && areas.length) { page.area = areas[0]; if (areas.length > 1) page.areas = areas; }
    else if (pn.kind === 'titles') for (const rg of areas) {
      if (rg.c1 === 0 && rg.c2 === MAX_COLS - 1) page.titleRows = [rg.r1, rg.r2];
      else if (rg.r1 === 0 && rg.r2 === EXCEL_MAX_ROWS - 1) page.titleCols = [rg.c1, rg.c2];
    }
    sh.page = page;
  }
  // 외부 통합 문서 참조 ([1]시트!A1): 파일에 저장된 외부 값을 숨긴 읽기 전용 시트로 (엑셀도 연결을 새로 고치기 전에는 이 값을 보여 줌)
  const extSheets = [];
  const externals = readExternalLinks(files, wbRoot, wbRels, extSheets);
  yield { p: 0.92, msg: "피벗 테이블 · 슬라이서 연결 중" };
  // 같은 캐시를 쓰는 피벗은 한 번만 읽고, 레코드 원문은 압축된 열 배열로 대체합니다.
  const preparedCaches = new Set();
  for (const sheet of sheets) for (const pivot of sheet._pivots ?? []) {
    const path = pivot.cachePath;
    if (!path || preparedCaches.has(path)) continue;
    preparedCaches.add(path);
    yield { preparePart: path, p: 0.92, msg: '피벗 캐시 정의 읽는 중' };
    const cache = readPivotCache(files, path, ctx.date1904);
    if (!cache?.snapshot) continue;
    yield { preparePart: cache.snapshot.path, p: 0.92, msg: '저장된 피벗 데이터 읽는 중' };
    const snapshot = yield* progressOf(cacheRecordsSteps(files, cache, streamParts), count => ({ p: 0.92, msg: `저장된 피벗 데이터 읽는 중 (${count.toLocaleString()}행)` }));
    if (snapshot) files[cache.snapshot.path] = snapshot;
    files.__xlsb?.cacheRecords.delete(cache.snapshot.path);
  }
  linkPivotsAndSlicers(files, wbRels, sheets, ctx);
  dropOffAxisFilters(sheets);
  // 공유 스타일·슬라이서·미배치 필터 정규화가 끝난 실제 정의를 기록합니다.
  for (const sheet of sheets) for (const def of [sheet.pivot, ...(sheet.pivotsExtra ?? [])]) {
    const layout = def?.sourceReference?.pivotLayout;
    if (layout) layout.binding = pivotCacheLayoutBinding(def);
    if (def?.importedPresentation) bindImportedPivotPresentation(def, def.importedPresentation);
  }
  pushAll(sheets, extSheets);
  if (unsupported) warnings.push(`지원하지 않는 함수가 쓰인 수식 ${unsupported}개는 원문과 파일에 저장된 계산 결과를 보존합니다. 관련 입력이 바뀌면 오래된 값을 오류로 표시하므로 [계산 상태 확인]을 확인하세요.`);
  if (files.__xlsb?.unsupported) warnings.push(`바이너리 통합 문서(.xlsb)에서 해석하지 못한 수식 ${files.__xlsb.unsupported}개는 저장된 계산 결과(값)로 가져왔습니다.`);
  for (const warning of new Set(files.__xlsb?.warnings ?? [])) warnings.push(warning);
  pushAll(warnings, [...ctx.warnings]);
  if (!sheets.length) throw new Error('가져올 시트가 없습니다');
  const data = { sheets, ...(date1904 ? { date1904: true } : {}) };
  const cp = child(wbRoot, 'calcPr');
  if (cp) data.calculation = calculationFromAttrs(cp.attrs);
  if (data.calculation?.iterate) warnings.push('반복 계산 설정은 파일에 보존하지만 현재 계산 엔진에서 반복 계산을 실행하지 않습니다. 순환 참조 결과를 확인하세요.');
  if (data.calculation?.fullPrecision === false) warnings.push('표시된 정밀도로 계산 설정은 파일에 보존하지만 현재 계산은 원래 숫자의 정밀도를 사용합니다.');
  if (ctx.pivotSnapshots) data.pivotSnapshots = ctx.pivotSnapshots;
  if (ctx.pivotCacheItems) data.pivotCacheItems = ctx.pivotCacheItems;
  // 자동 높이로 맞출 행 (화면에서 글자 크기를 재어 정함 — 앱이 열 때 한 번 계산)
  if (sheets.some((sh) => sh.fitRows)) data.fitRows = sheets.map((sh) => { const f = sh.fitRows ? [...sh.fitRows] : null; delete sh.fitRows; return f; });
  if (names.length) data.names = names;
  if (externals.length) data.externals = externals;
  // 매크로(.xlsm): vbaProject.bin 을 그대로 보존 (실행하지 않음)
  const vbaRel = Object.values(wbRels).find((r) => r.type === 'vbaProject');
  if (vbaRel && files[vbaRel.target]) {
    data.vba = {
      bin: toBase64(files[vbaRel.target]),
      codeName: child(wbRoot, 'workbookPr')?.attrs.codeName ?? null,
      sheetCodes,
    };
  }
  if (theme.join() !== DEFAULT_THEME.join()) data.theme = [...theme]; // 테마 색 (표 · 피벗 스타일 색 계산)
  { const trel = Object.values(wbRels).find((r) => r.type === 'theme'); const tx = trel && textOf(files[trel.target]); if (tx && tx.length < 400000) { data.themeXml = tx; Object.assign(data, readThemeOptions(tx)); } }
  {
    // 문서 속성 (docProps/core.xml) · 보호 (통합 문서 구조 · 읽기 전용 권장 · 최종본)
    const props = {};
    const core = files['docProps/core.xml'] && parseXml(textOf(files['docProps/core.xml']));
    if (core) {
      const pick = (tag) => descendants(core, tag)[0];
      for (const [k, tag] of [['title', 'title'], ['subject', 'subject'], ['creator', 'creator'], ['tags', 'keywords'], ['comments', 'description'], ['lastModifiedBy', 'lastModifiedBy'], ['category', 'category'], ['created', 'created'], ['modified', 'modified']]) {
        const e = pick(tag);
        const t = e ? String(e.text ?? '').trim() : '';
        if (t) props[k] = t;
      }
    }
    const wp = child(wbRoot, 'workbookProtection');
    if (wp && (wp.attrs.lockStructure === '1' || wp.attrs.lockStructure === 'true')) props.lockStructure = true;
    if (wp) props.workbookProtection = Object.fromEntries(WORKBOOK_PROTECTION_KEYS.filter(k => wp.attrs[k] !== undefined).map(k => [k, wp.attrs[k]]));
    const fs = child(wbRoot, 'fileSharing');
    if (fs && (fs.attrs.readOnlyRecommended === '1' || fs.attrs.readOnlyRecommended === 'true')) props.readOnlyRecommended = true;
    const custom = files['docProps/custom.xml'] && textOf(files['docProps/custom.xml']);
    if (custom && /name="_MarkAsFinal"[^>]*>\s*<vt:bool>(true|1)<\/vt:bool>/.test(custom)) props.markedFinal = true;
    const history = custom && kids(parseXml(custom), 'property').find(p => p.attrs.name === '_WixelXlsxImportWarnings');
    let savedCodes = [];
    if (history) { try { savedCodes = importWarningCodes({ xlsxImportWarnings: JSON.parse(child(history, 'lpwstr')?.text ?? '') }); } catch { /* 잘못된 이력은 무시 */ } }
    const codes = [...new Set([...savedCodes, ...(ctx.importWarningCodes ?? [])])];
    if (codes.length) {
      props.xlsxImportWarnings = codes;
      for (const code of codes) if (!warnings.includes(XLSX_IMPORT_WARNINGS[code])) warnings.push(XLSX_IMPORT_WARNINGS[code]);
    }
    if (Object.keys(props).length) data.props = props;
  }
  data.defaultFont = wbFont; // 통합 문서 기본 글꼴 (표준 스타일) — 셀 기본 크기 · 열 너비 변환에 씀
  // 기본 셀 서식(xf 0): s 속성이 없는 셀에 적용됨 (한국어 엑셀은 보통 세로 가운데 맞춤)
  if (xfs[0] && Object.keys(xfs[0]).length) data.baseStyle = { ...xfs[0] };
  if (cellStyles?.length) data.cellStyles = cellStyles;
  if (objectStyles) data.objectStyles = objectStyles;
  const active = Number(descendants(child(wbRoot, 'bookViews'), 'workbookView')[0]?.attrs.activeTab ?? 0);
  let act = Math.min(active, sheets.length - 1);
  if (sheets[act]?.state) act = Math.max(0, sheets.findIndex((x) => !x.state));
  return { data, active: act, warnings };
}

/**
 * 외부 통합 문서 연결 (xl/externalLinks/externalLinkN.xml): 연결 순서(workbook.xml externalReferences)가 수식의 [N].
 * 저장된 값(sheetDataSet)은 이름 '[N]시트' 인 veryHidden 시트(external: N)로 sheets 끝에 붙이고,
 * 원본 XML · 관계는 되돌려 쓰도록 [{ index, xml, rels, target }] 로 돌려줌
 */
function readExternalLinks(files, wbRoot, wbRels, sheets) {
  const out = [];
  kids(child(wbRoot, 'externalReferences'), 'externalReference').forEach((er, k) => {
    const rel = wbRels[rid(er)];
    if (!rel || !files[rel.target]) return;
    const index = k + 1;
    const xml = textOf(files[rel.target]);
    const relsPath = rel.target.replace(/([^/]+)$/, '_rels/$1.rels');
    const rels = files[relsPath] ? textOf(files[relsPath]) : null;
    const target = rels ? decodeEntities(/Target="([^"]*)"/.exec(rels)?.[1] ?? '') : '';
    out.push({ index, xml, rels, target });
    let root;
    try { root = parseXml(xml); } catch { return; }
    const book = child(root, 'externalBook');
    if (!book) return;
    const names = kids(child(book, 'sheetNames'), 'sheetName').map((x) => x.attrs.val ?? '');
    for (const sd of kids(child(book, 'sheetDataSet'), 'sheetData')) {
      const nm = names[Number(sd.attrs.sheetId)];
      if (nm === undefined) continue;
      const cells = new Map();
      for (const row of kids(sd, 'row')) {
        for (const c of kids(row, 'cell')) {
          const m = /^([A-Z]+)(\d+)$/.exec(c.attrs.r ?? '');
          if (!m) continue;
          const v = child(c, 'v')?.text ?? '';
          const t = c.attrs.t;
          const raw = t === 's' || t === 'str' ? `'${v}` : t === 'b' ? (v === '1' ? 'TRUE' : 'FALSE') : v;
          if (raw !== '') cells.set(`${Number(m[2]) - 1},${nameToCol(m[1])}`, { raw });
        }
      }
      sheets.push({ name: `[${index}]${nm}`, cells, state: 'veryHidden', external: index });
    }
  });
  return out;
}

/** 셀 그림: metadata.xml valueMetadata(vm, 1부터) → 리치 값 → richValueRel → 그림 파일. [vm] = {src, alt} */
function readRichImages(files, wbRels) {
  const byType = (t) => Object.values(wbRels).find((r) => r.type === t && files[r.target]);
  const metaRel = byType('sheetMetadata');
  const rvRel = byType('rdRichValue');
  const stRel = byType('rdRichValueStructure');
  const relRel = byType('richValueRel');
  if (!metaRel || !rvRel || !stRel || !relRel) return null;
  try {
    const meta = parseXml(textOf(files[metaRel.target]));
    const types = kids(child(meta, 'metadataTypes'), 'metadataType').map((m) => m.attrs.name);
    const rich = kids(meta, 'futureMetadata').find((f) => f.attrs.name === 'XLRICHVALUE');
    const rvbs = kids(rich, 'bk').map((bk) => Number(descendants(bk, 'rvb')[0]?.attrs.i));
    const structs = kids(parseXml(textOf(files[stRel.target])), 's').map((st) => ({ t: st.attrs.t, keys: kids(st, 'k').map((k) => k.attrs.n) }));
    const rvs = kids(parseXml(textOf(files[rvRel.target])), 'rv').map((rv) => ({ s: Number(rv.attrs.s), vals: kids(rv, 'v').map((v) => v.text ?? '') }));
    const relIds = kids(parseXml(textOf(files[relRel.target])), 'rel').map((r) => rid(r) ?? Object.entries(r.attrs).find(([k]) => k.endsWith(':id'))?.[1]);
    const rels = relsOf(files, relRel.target);
    const out = [null];
    for (const bk of kids(child(meta, 'valueMetadata'), 'bk')) {
      const rc = child(bk, 'rc');
      let img = null;
      if (rc && types[Number(rc.attrs.t) - 1] === 'XLRICHVALUE') {
        const rv = rvs[rvbs[Number(rc.attrs.v)]];
        const st = rv && structs[rv.s];
        const ki = st ? st.keys.indexOf('_rvRel:LocalImageIdentifier') : -1;
        if (ki >= 0) {
          const target = rels[relIds[Number(rv.vals[ki])]]?.target;
          const bytes = target && files[target];
          const mime = bytes && MIME[target.split('.').pop().toLowerCase()];
          if (mime) {
            const ti = st.keys.indexOf('Text');
            img = { src: `data:${mime};base64,${toBase64(bytes)}`, alt: ti >= 0 ? rv.vals[ti] : '' };
          }
        }
      }
      out.push(img);
    }
    return out;
  } catch {
    return null;
  }
}

/** 하위 생성기의 중간 값을 진행 정보로 바꿔 전달 */
function* progressOf(gen, map) {
  for (;;) {
    const s = gen.next();
    if (s.done) return s.value;
    yield map(s.value);
  }
}

function relsTarget(files, path, type) {
  const rels = relsOf(files, path);
  return Object.values(rels).find((r) => r.type === type)?.target;
}

// ───────────────────────── 쓰기 ─────────────────────────
const THEME_SLOTS = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
/**
 * 테마(theme1.xml): 파일에서 읽은 테마가 있으면 그대로(색만 지금 테마로 바꿈), 없으면 새로 만듦
 * — 엑셀은 표 · 피벗 · 차트 스타일 색을 테마에서 계산하므로 빠지면 다른 색이 됨
 */
function themeXml(wb) {
  const colors = (wb.theme?.length ? wb.theme : DEFAULT_THEME).map((c) => String(c).replace('#', '').toUpperCase().slice(-6));
  if (wb.themeXml && /<a:clrScheme/.test(wb.themeXml)) {
    let x = wb.themeXml;
    THEME_SLOTS.forEach((slot, i) => {
      x = x.replace(new RegExp(`(<a:${slot}>)[\\s\\S]*?(</a:${slot}>)`), (m0, a1, a2) => {
        const cur = /(?:val|lastClr)="([0-9A-Fa-f]{6})"/.exec(m0.includes('lastClr') ? m0.replace(/val="[^"]*"/, '') : m0)?.[1]?.toUpperCase();
        return cur === colors[i] ? m0 : `${a1}<a:srgbClr val="${colors[i]}"/>${a2}`;
      });
    });
    if (wb.themeName) x = x.replace(/(<a:clrScheme name=")[^"]*"/, `$1${esc(wb.themeName)}"`);
    return applyThemeOptionsXml(x, wb);
  }
  const clr = THEME_SLOTS.map((slot, i) => (i === 0 ? `<a:lt1><a:sysClr val="window" lastClr="${colors[0]}"/></a:lt1>` : i === 1 ? `<a:dk1><a:sysClr val="windowText" lastClr="${colors[1]}"/></a:dk1>` : `<a:${slot}><a:srgbClr val="${colors[i]}"/></a:${slot}>`));
  const order = [clr[1], clr[0], clr[3], clr[2], ...clr.slice(4)].join('');
  const font = (latin, ea) => `<a:latin typeface="${latin}" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hang" typeface="${ea}"/>`;
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  const name = esc(wb.themeName ?? 'Office 테마');
  return applyThemeOptionsXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}"><a:themeElements><a:clrScheme name="${name}">${order}</a:clrScheme>`
    + `<a:fontScheme name="Office"><a:majorFont>${font('맑은 고딕', '맑은 고딕')}</a:majorFont><a:minorFont>${font('맑은 고딕', '맑은 고딕')}</a:minorFont></a:fontScheme>`
    + `<a:fmtScheme name="Office"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme>`
    + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>', wb);
}

function styleNodeXml(node) {
  if(!node||!/^[A-Za-z_][\w.-]*$/.test(node.name))return '';
  // 미편집 DXF의 원본 속성은 보존하고 웹 별칭만 설치 글꼴 이름으로 기록한다.
  if (node.name === 'font') {
    const name = child(node, 'name'), bold = child(node, 'b');
    if (name?.attrs.val) {
      const mapped = fontDesktopStyle(name.attrs.val, bold ? { bold: !['0', 'false'].includes(bold.attrs.val) } : {});
      if (mapped.font !== name.attrs.val || mapped.bold !== undefined) {
        const children = (node.children ?? []).map(item => item === name ? { ...item, attrs: { ...item.attrs, val: mapped.font } } : item);
        if (mapped.bold !== undefined) {
          const at = children.indexOf(bold), value = { name: 'b', attrs: { ...(bold?.attrs ?? {}), val: mapped.bold ? '1' : '0' }, children: [] };
          if (at < 0) children.unshift(value); else children[at] = value;
        }
        node = { ...node, children };
      }
    }
  }
  const attrs=Object.entries(node.attrs??{}).filter(([k])=>/^[A-Za-z_][\w.:-]*$/.test(k)).map(([k,v])=>` ${k}="${esc(v)}"`).join('');
  const body=esc(node.text??'')+(node.children??[]).map(styleNodeXml).join('');
  return `<${node.name}${attrs}${body?`>${body}</${node.name}>`:'/>'}`;
}

class StylePool {
  constructor(baseFont = WRITE_FONT, baseStyle = null) {
    this.baseFont = { name: baseFont.name || DEFAULT_FONT, size: baseFont.size || 11 };
    this.baseStyle = baseStyle && Object.keys(baseStyle).length ? baseStyle : null;
    const desktop = fontDesktopStyle(this.baseFont.name);
    this.fonts = [`<font>${desktop.bold !== undefined ? `<b val="${desktop.bold ? 1 : 0}"/>` : ''}<sz val="${this.baseFont.size}"/><color theme="1"/><name val="${esc(desktop.font)}"/><family val="3"/><charset val="129"/></font>`];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = [];
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.dxfs = [];
    this.tableStyles = new Map();
    this.slicerStyles = new Map();
    this.slicerDxfs = [];
    this.byObj = new WeakMap();
    this.maps = { font: new Map([[this.fonts[0], 0]]), fill: new Map(this.fills.map((f, i) => [f, i])), border: new Map([[this.borders[0], 0]]), fmt: new Map(), xf: new Map([['{}', 0]]) };
    // 기본 셀 서식(xf 0)을 통합 문서의 기본 서식으로 씀 — 서식 없는 셀도 엑셀에서 같은 모양
    if (this.baseStyle) {
      this.xfs = [];
      this.maps.xf.clear();
      this.xfOf(this.baseStyle);
      this.maps.xf.set('{}', 0);
    }
  }

  intern(kind, list, xml) {
    const map = this.maps[kind];
    if (!map.has(xml)) { map.set(xml, list.length); list.push(xml); }
    return map.get(xml);
  }

  fmtId(style) {
    const code = fmtCode(style);
    if (code === null) return 0;
    // 엑셀 기본 형식 번호가 있는 코드는 그 번호로 (다시 열 때 같은 형식)
    const builtin = BUILTIN_CODE_ID[code];
    if (builtin !== undefined) return builtin;
    if (!this.maps.fmt.has(code)) {
      const id = 164 + this.numFmts.length;
      this.maps.fmt.set(code, id);
      this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    }
    return this.maps.fmt.get(code);
  }

  xf(style) {
    if (!style) return 0;
    // 셀 서식 객체는 여러 셀이 공유하므로 객체마다 한 번만 계산
    const hit = this.byObj.get(style);
    if (hit !== undefined) return hit;
    const id = this.xfOf(style);
    this.byObj.set(style, id);
    return id;
  }

  xfOf(style) {
    if (!Object.keys(style).length) return 0;
    const k = JSON.stringify(Object.keys(style).sort().map((key) => [key, style[key]]));
    if (!this.styleXfMode && this.maps.xf.has(k)) return this.maps.xf.get(k);
    const desktop = { ...style, ...fontDesktopStyle(style.font || this.baseFont.name, style) };
    const fontFlag = (key, tag) => desktop[key] ? `<${tag}/>` : desktop[key] === false ? `<${tag} val="${tag === 'u' ? 'none' : '0'}"/>` : '';
    const font = `<font>${fontFlag('bold', 'b')}${fontFlag('italic', 'i')}${fontFlag('strike', 'strike')}${fontFlag('underline', 'u')}<sz val="${style.size || this.baseFont.size}"/>${style.color ? `<color rgb="${argb(style.color)}"/>` : '<color theme="1"/>'}<name val="${esc(desktop.font)}"/><family val="3"/><charset val="129"/></font>`;
    const fontId = this.intern('font', this.fonts, font);
    const g = style.gradient;
    const gStops = g?.stops?.length ? g.stops.map(([p, c]) => `<stop position="${Number(p) || 0}"><color rgb="${argb(c)}"/></stop>`).join('') : '';
    const fillId = gStops
      ? this.intern('fill', this.fills, `<fill><gradientFill${g.path ? ` type="path" left="${g.l ?? 0}" right="${g.r ?? 0}" top="${g.t ?? 0}" bottom="${g.b ?? 0}"` : g.deg ? ` degree="${g.deg}"` : ''}>${gStops}</gradientFill></fill>`)
      : style.pattern
      ? this.intern('fill', this.fills, `<fill><patternFill patternType="${esc(style.pattern)}"><fgColor rgb="${argb(style.patternColor ?? '#000000')}"/>${style.fill ? `<bgColor rgb="${argb(style.fill)}"/>` : '<bgColor indexed="64"/>'}</patternFill></fill>`)
      : style.fill ? this.intern('fill', this.fills, `<fill><patternFill patternType="solid"><fgColor rgb="${argb(style.fill)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const side = (n, k) => (style[k] ? `<${n} style="${style[`${k}s`] ?? 'thin'}">${style[`${k}c`] ? `<color rgb="${argb(style[`${k}c`])}"/>` : '<color indexed="64"/>'}</${n}>` : `<${n}/>`);
    const diag = style.dd || style.du ? side('diagonal', style.dd ? 'dd' : 'du') : '<diagonal/>';
    const borderId = style.bt || style.bb || style.bl || style.br || style.dd || style.du
      ? this.intern('border', this.borders, `<border${style.dd ? ' diagonalDown="1"' : ''}${style.du ? ' diagonalUp="1"' : ''}>${side('left', 'bl')}${side('right', 'br')}${side('top', 'bt')}${side('bottom', 'bb')}${diag}</border>`)
      : 0;
    const numFmtId = this.fmtId(style);
    const align = [];
    if (style.align) align.push(`horizontal="${style.align}"`);
    if (style.valign) align.push(`vertical="${style.valign === 'middle' ? 'center' : style.valign === 'bottom' ? 'bottom' : 'top'}"`);
    if (style.wrap) align.push('wrapText="1"');
    if (style.indent) align.push(`indent="${style.indent}"`);
    if (style.rotate) align.push(`textRotation="${style.rotate === 255 ? 255 : style.rotate < 0 ? 90 - style.rotate : style.rotate}"`);
    if (style.shrink != null) align.push(`shrinkToFit="${style.shrink ? 1 : 0}"`);
    const prot = style.locked === false || style.hideFormula ? `<protection${style.locked === false ? ' locked="0"' : ''}${style.hideFormula ? ' hidden="1"' : ''}/>` : '';
    // 확인란: 엑셀 365 방식 (xf 의 xfComplement → featurePropertyBag 의 Checkbox 셀 컨트롤). 옛 엑셀은 TRUE/FALSE 로 표시
    if (style.checkbox) this.hasCheckbox = true;
    const extensions = [];
    if (style.checkbox) extensions.push('<ext uri="{C7286773-470A-42A8-94C5-96B5CB345126}" xmlns:xfpb="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><xfpb:xfComplement i="0"/></ext>');
    const inheritance = tableStyleInheritance(style.tableStyleInherit);
    if (inheritance) {
      for (const [key, neutral] of Object.entries(inheritance)) if (style[key] !== neutral) delete inheritance[key];
      if (Object.keys(inheritance).length) extensions.push(`<ext uri="${TABLE_STYLE_INHERIT_URI}"><wx:tableStyleInherit xmlns:wx="${TABLE_STYLE_INHERIT_NS}" json="${esc(JSON.stringify(inheritance))}"/></ext>`);
    }
    const ext = extensions.length ? `<extLst>${extensions.join('')}</extLst>` : '';
    const inner = (align.length ? `<alignment ${align.join(' ')}/>` : '') + prot + ext;
    const parts = { number: numFmtId, font: fontId, fill: fillId, border: borderId, alignment: align.join(' '), protection: prot };
    // 글꼴/채우기만 지정한 XF에 일반 표시 형식을 명시하면 다시 읽을 때 사용자 숫자 서식으로 굳어진다.
    const implicitGeneral = numFmtId === 0 && !style.numFmt && style.decimals == null && !style.code;
    if (this.styleXfMode) {
      this.lastStyleParts = parts;
      const flags = Object.entries(STYLE_APPLY).map(([part, flag]) => part === 'number' && implicitGeneral && this.styleXfInclude?.[part] !== false ? '' : ` ${flag}="${this.styleXfInclude?.[part] === false ? 0 : 1}"`).join('');
      return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"${flags}${inner ? `>${inner}</xf>` : '/>'}`;
    }
    const named = this.namedByKey?.get(styleNameKey(style.cellStyleName));
    // 이름 연결은 xfId로 보존합니다. 부모와 같은 요소는 상속하고 직접 고친 요소는 셀에 기록합니다.
    const flags = Object.entries(STYLE_APPLY).map(([part, flag]) => {
      const inherit = named && named.include[part] && named.parts[part] === parts[part];
      if (part === 'number' && implicitGeneral && !inherit) return '';
      const value = named && inherit ? 0 : 1;
      return ` ${flag}="${value}"`;
    }).join('');
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="${named?.id ?? 0}"${flags}${inner ? `>${inner}</xf>` : '/>'}`;
    const id = this.xfs.length;
    this.xfs.push(xml);
    this.maps.xf.set(k, id);
    return id;
  }

  /** 이름 있는 셀 스타일 → cellStyleXfs + cellStyles (엑셀 [셀 스타일] 갤러리의 사용자 지정) */
  namedStyles(list) {
    this.namedByKey = new Map();
    this.cellStyleXfs = [];
    this.cellStyleList = ['<cellStyle name="표준" xfId="0" builtinId="0"/>'];
    const compile = (style, include) => {
      this.styleXfMode = true; this.styleXfInclude = include;
      try { return this.xfOf(Object.keys(style ?? {}).length ? style : { numFmt: 'general' }); }
      finally { this.styleXfMode = false; this.styleXfInclude = null; }
    };
    this.cellStyleXfs.push(compile(this.baseStyle, null));
    const seen = new Set(['표준', 'normal']);
    for (const cs of list ?? []) {
      const name = String(cs?.name ?? '').trim(), key = styleNameKey(name);
      if (!name || seen.has(key) || cs.builtinId === 0) continue;
      seen.add(key);
      const include = {};
      for (const part of Object.keys(STYLE_APPLY)) include[part] = cs.include?.[part] !== false;
      const xml = compile(cs.style, include), id = this.cellStyleXfs.length;
      this.namedByKey.set(key, { id, include, parts: this.lastStyleParts });
      const customBuiltin = typeof cs.customBuiltin === 'boolean' ? ` customBuiltin="${cs.customBuiltin ? 1 : 0}"` : '';
      const builtIn = Number.isInteger(cs.builtinId) && cs.builtinId >= 0 ? ` builtinId="${cs.builtinId}"${customBuiltin}` : '';
      const level = Number.isInteger(cs.iLevel) && cs.iLevel >= 0 ? ` iLevel="${cs.iLevel}"` : '';
      this.cellStyleList.push(`<cellStyle name="${xesc(name)}" xfId="${id}"${builtIn}${cs.hidden ? ' hidden="1"' : ''}${level}/>`);
      this.cellStyleXfs.push(xml);
    }
  }

  dxf(style) {
    const desktop = { ...style, ...fontDesktopStyle(style.font, style) };
    const flag = (key, tag) => desktop[key] !== undefined ? `<${tag} val="${desktop[key] ? (tag === 'u' ? 'single' : '1') : (tag === 'u' ? 'none' : '0')}"/>` : '';
    const font = ['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike'].some(key => desktop[key] !== undefined)
      ? `<font>${flag('bold', 'b')}${flag('italic', 'i')}${flag('strike', 'strike')}${flag('underline', 'u')}${desktop.font ? `<name val="${esc(desktop.font)}"/>` : ''}${desktop.size ? `<sz val="${Number(desktop.size)}"/>` : ''}${desktop.color ? `<color rgb="${argb(desktop.color)}"/>` : ''}</font>` : '';
    const fill = style.fill ? `<fill><patternFill><bgColor rgb="${argb(style.fill)}"/></patternFill></fill>` : '';
    const side = (n, k) => (style[k] ? `<${n} style="${style[`${k}s`] ?? 'thin'}">${style[`${k}c`] ? `<color rgb="${argb(style[`${k}c`])}"/>` : '<color auto="1"/>'}</${n}>` : '');
    const border = style.bt || style.bb || style.bl || style.br || style.bv || style.bh ? `<border>${side('left', 'bl')}${side('right', 'br')}${side('top', 'bt')}${side('bottom', 'bb')}${side('vertical', 'bv')}${side('horizontal', 'bh')}</border>` : '';
    // 표시 형식도 조건부 서식으로 바꿀 수 있음 (dxf 안의 numFmt)
    const code = style.numFmt ? fmtCode(style) : null;
    const nf = code !== null ? `<numFmt numFmtId="${this.fmtId(style)}" formatCode="${esc(code)}"/>` : '';
    this.dxfs.push(`<dxf>${font}${nf}${fill}${border}</dxf>`);
    return this.dxfs.length - 1;
  }

  objectDxf(element) {
    if(element.sourceDxf?.name==='dxf' && JSON.stringify(element.style??{})===JSON.stringify(element.sourceStyle)) {
      this.dxfs.push(styleNodeXml(element.sourceDxf));return this.dxfs.length-1;
    }
    const original=element.style??{},st={...original,...fontDesktopStyle(original.font,original)};
    const bools=[['bold','b'],['italic','i'],['strike','strike']].filter(([k])=>st[k]!==undefined).map(([k,t])=>`<${t} val="${st[k]?1:0}"/>`).join('');
    const f=bools+(st.underline!==undefined?`<u val="${st.underline?'single':'none'}"/>`:'')+(st.size?`<sz val="${Number(st.size)}"/>`:'')+(st.font?`<name val="${esc(st.font)}"/>`:'')+(st.color?`<color rgb="${argb(st.color)}"/>`:'');
    let fill='';const g=st.gradient;
    if(g?.stops?.length)fill=`<fill><gradientFill${g.path?` type="path" left="${g.l??0}" right="${g.r??0}" top="${g.t??0}" bottom="${g.b??0}"`:` degree="${Number(g.deg)||0}"`}>${g.stops.map(([p,c])=>`<stop position="${Number(p)||0}"><color rgb="${argb(c)}"/></stop>`).join('')}</gradientFill></fill>`;
    else if(st.pattern)fill=`<fill><patternFill patternType="${esc(st.pattern)}"><fgColor rgb="${argb(st.patternColor??'#000000')}"/>${st.fill?`<bgColor rgb="${argb(st.fill)}"/>`:''}</patternFill></fill>`;
    else if(st.fill===null)fill='<fill><patternFill patternType="none"/></fill>';
    else if(st.fill)fill=`<fill><patternFill patternType="solid"><fgColor rgb="${argb(st.fill)}"/><bgColor rgb="${argb(st.fill)}"/></patternFill></fill>`;
    const side=(n,k)=>st[k]?`<${n} style="${esc(st[k+'s']??'thin')}"><color rgb="${argb(st[k+'c']??'#000000')}"/></${n}>`:`<${n}/>`;
    const diag=st.dd!==undefined||st.du!==undefined?side('diagonal',st.dd?'dd':'du'):'';
    const bd=[['left','bl'],['right','br'],['top','bt'],['bottom','bb']].filter(([,k])=>st[k]!==undefined).map(([n,k])=>side(n,k)).join('')+diag+[['vertical','bv'],['horizontal','bh']].filter(([,k])=>st[k]!==undefined).map(([n,k])=>side(n,k)).join('');
    const code=st.numFmt||st.code||st.decimals!==undefined?fmtCode(st):null;
    const nf=code!==null?`<numFmt numFmtId="${this.fmtId(st)}" formatCode="${esc(code)}"/>`:'';
    this.dxfs.push(`<dxf>${f?`<font>${f}</font>`:''}${nf}${fill}${bd?`<border${st.dd!==undefined?` diagonalDown="${st.dd?1:0}"`:''}${st.du!==undefined?` diagonalUp="${st.du?1:0}"`:''}>${bd}</border>`:''}</dxf>`);return this.dxfs.length-1;
  }
  slicerDxf(element) {
    this.objectDxf(element);this.slicerDxfs.push(this.dxfs.pop());return this.slicerDxfs.length-1;
  }
  objectTableStyle(def) {
    if(this.tableStyles.has(def.name))return;
    const list=(def.elements??[]).filter(e=>TABLE_STYLE_ELEMENTS.includes(e.type));
    const els=list.map(e=>`<tableStyleElement type="${e.type}"${Number.isInteger(e.size)&&e.size>0?` size="${e.size}"`:''} dxfId="${this.objectDxf(e)}"/>`);
    this.tableStyles.set(def.name,`<tableStyle name="${esc(def.name)}" table="${def.table!==false?1:0}" pivot="${def.pivot!==false?1:0}" count="${def._slicerCount??els.length}">${els.join('')}</tableStyle>`);
  }
  objectSlicerStyle(def) {
    if(this.slicerStyles.has(def.name))return;
    const list=(def.elements??[]).filter(e=>SLICER_STYLE_ELEMENTS.includes(e.type));
    // 슬라이서의 whole/header 정의는 읽은 보조 tableStyle보다 최신이다.
    this.tableStyles.delete(def.name);
    this.objectTableStyle({name:def.name,table:false,pivot:false,_slicerCount:list.length,elements:list.filter(e=>['wholeTable','headerRow'].includes(e.type))});
    const els=list.filter(e=>!['wholeTable','headerRow'].includes(e.type)).map(e=>`<x14:slicerStyleElement type="${e.type}" dxfId="${this.slicerDxf(e)}"/>`);
    this.slicerStyles.set(def.name,`<x14:slicerStyle name="${esc(def.name)}">${els.length?`<x14:slicerStyleElements>${els.join('')}</x14:slicerStyleElements>`:''}</x14:slicerStyle>`);
  }
  objectStyles(registry) {
    this.objectRegistry=normalizeObjectStyles(registry);
    for(const d of this.objectRegistry.tables)this.objectTableStyle(d);
    for(const d of this.objectRegistry.slicers)this.objectSlicerStyle(d);
  }

  /** WIXEL 모던 스타일 → 사용자 지정 표/피벗 스타일 (<tableStyles>) — 엑셀에서도 같은 모양 */
  presetTableStyle(name, pivot) {
    if (this.tableStyles.has(name)) return;
    const p = presetStyle(name);
    if (!p) return;
    const els = Object.entries(ELEMENT_TYPES).filter(([k]) => p[k]).map(([k, type]) => `<tableStyleElement type="${type}" dxfId="${this.dxf(elementDxfStyle(p[k]))}"/>`);
    this.tableStyles.set(name, `<tableStyle name="${esc(name)}"${pivot ? ' table="0"' : ' pivot="0"'} count="${els.length}">${els.join('')}</tableStyle>`);
  }

  /** 슬라이서 스타일 이름 (모던 · 사용자 지정 색이면 파일에 사용자 지정 슬라이서 스타일을 만듦) */
  slicerStyleFor(sl) {
    const base = slicerStyleName(sl);
    if (base === 'None') {this.objectSlicerStyle({name:'WIXEL Slicer None',elements:slicerStyleElements({style:'None'})});return 'WIXEL Slicer None';}
    if (Array.isArray(sl.styleElements) && !sl.custom) {this.objectSlicerStyle({name:base,elements:sl.styleElements});return base;}
    if (!sl.custom && !isModernSlicer(base)) return base;
    const c = slicerColors(sl);
    const name = sl.custom ? `WIXEL 사용자 지정 ${[...JSON.stringify(c)].reduce((h, ch) => (h * 31 + ch.charCodeAt(0)) >>> 0, 7).toString(36)}` : base;
    if (this.slicerStyles.has(name)) return name;
    const bd = (color) => ({ bt: true, bb: true, bl: true, br: true, btc: color, bbc: color, blc: color, brc: color });
    const whole = this.dxf({ fill: c.frame, ...bd(c.border) });
    const head = this.dxf({ color: c.head, bold: true });
    this.tableStyles.set(name, `<tableStyle name="${esc(name)}" pivot="0" table="0" count="8"><tableStyleElement type="wholeTable" dxfId="${whole}"/><tableStyleElement type="headerRow" dxfId="${head}"/></tableStyle>`);
    const els = [
      ['selectedItemWithData', { fill: c.selFill, color: c.selText, ...bd(c.selBorder) }],
      ['selectedItemWithNoData', { fill: c.selFill, color: c.noData, ...bd(c.selBorder) }],
      ['unselectedItemWithData', { fill: c.item, color: c.itemText, ...bd(c.itemBorder) }],
      ['unselectedItemWithNoData', { fill: c.item, color: c.noData, ...bd(c.itemBorder) }],
      ['hoveredSelectedItemWithData', { fill: c.selFill, color: c.selText, ...bd(c.selBorder) }],
      ['hoveredUnselectedItemWithData', { fill: c.item, color: c.itemText, ...bd(c.selBorder) }],
    ].map(([type, st]) => `<x14:slicerStyleElement type="${type}" dxfId="${this.slicerDxf({style:st})}"/>`);
    this.slicerStyles.set(name, `<x14:slicerStyle name="${esc(name)}"><x14:slicerStyleElements>${els.join('')}</x14:slicerStyleElements></x14:slicerStyle>`);
    return name;
  }

  /** 사용자 지정 피벗 스타일을 <tableStyles> 에 추가 */
  pivotStyle(name, parts) {
    if (this.tableStyles.has(name)) return;
    const els = [['wholeTable', parts.body], ['headerRow', parts.header], ['totalRow', parts.grand], ['firstRowStripe', parts.band],
      ['firstSubtotalRow', parts.sub], ['pageFieldLabels', parts.page]]
      .filter(([, st]) => st && Object.keys(st).length)
      .map(([type, st]) => `<tableStyleElement type="${type}" dxfId="${this.dxf(st)}"/>`);
    this.tableStyles.set(name, `<tableStyle name="${esc(name)}" table="0" count="${els.length}">${els.join('')}</tableStyle>`);
  }

  xml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS_MAIN}">`
      + (this.numFmts.length ? `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>` : '')
      + `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`
      + `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`
      + `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      + (this.cellStyleXfs ? `<cellStyleXfs count="${this.cellStyleXfs.length}">${this.cellStyleXfs.join('')}</cellStyleXfs>` : '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>')
      + `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`
      + (this.cellStyleList ? `<cellStyles count="${this.cellStyleList.length}">${this.cellStyleList.join('')}</cellStyles>` : '<cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles>')
      + `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>`
      + (this.tableStyles.size || this.objectRegistry?.defaultTableStyle || this.objectRegistry?.defaultPivotStyle ? `<tableStyles count="${this.tableStyles.size}" defaultTableStyle="${esc(this.objectRegistry?.defaultTableStyle??'TableStyleMedium2')}" defaultPivotStyle="${esc(this.objectRegistry?.defaultPivotStyle??'PivotStyleLight16')}">${[...this.tableStyles.values()].join('')}</tableStyles>` : '')
      + (this.indexedColors ? `<colors><indexedColors>${this.indexedColors.map(c => `<rgbColor rgb="FF${c}"/>`).join('')}</indexedColors></colors>` : '')
      + (this.slicerStyles.size || this.objectRegistry?.defaultSlicerStyle ? `<extLst>${this.slicerDxfs.length?`<ext uri="{46F421CA-312F-682f-3DD2-61675219B42D}" xmlns:x14="${NS_X14}"><x14:dxfs count="${this.slicerDxfs.length}">${this.slicerDxfs.join('')}</x14:dxfs></ext>`:''}<ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="${NS_X14}"><x14:slicerStyles defaultSlicerStyle="${esc(this.objectRegistry?.defaultSlicerStyle??'SlicerStyleLight1')}">${[...this.slicerStyles.values()].join('')}</x14:slicerStyles></ext></extLst>` : '')
      + '</styleSheet>';
  }
}

/** 앱 수식 → 파일 수식 (_xlfn. 접두사, 표 참조 표준 형식, '@'·'#' 변환) */
let fileNameCheck = null;
function exportFormula(raw, hereTable = null, dynamic = false) {
  return toFileFormula(raw, { hereTable, dynamic, isName: fileNameCheck });
}

// A remembered menu function cannot overwrite a subsequently edited totals cell.
function matchingTotalFunction(raw, fn, table, column) {
  const f = TOTAL_FUNCS.find(x => x.id === fn && x.code);
  if (!f || !raw?.startsWith('=')) return false;
  try {
    const a = parse(raw.slice(1)), arg = a.args?.[1];
    if (a.type !== 'func' || a.name !== 'SUBTOTAL' || a.args.length !== 2 || a.args[0]?.type !== 'num' || a.args[0].v !== f.code || arg?.type !== 'sref') return false;
    const expected = column.replace(/(['[\]#@])/g, "'$1");
    return canonicalRef(arg.table, arg.spec, table.name).toLowerCase() === canonicalRef(table.name, expected, table.name).toLowerCase();
  } catch { return false; }
}

const CELL_IS = { gt: 'greaterThan', lt: 'lessThan', ge: 'greaterThanOrEqual', le: 'lessThanOrEqual', eq: 'equal', ne: 'notEqual', between: 'between', notBetween: 'notBetween' };
const TIME_FORMULA = {
  today: (c) => `FLOOR(${c},1)=TODAY()`, yesterday: (c) => `FLOOR(${c},1)=TODAY()-1`, tomorrow: (c) => `FLOOR(${c},1)=TODAY()+1`,
  last7Days: (c) => `AND(TODAY()-FLOOR(${c},1)<=6,FLOOR(${c},1)<=TODAY())`,
  thisWeek: (c) => `AND(TODAY()-ROUNDDOWN(${c},0)<=WEEKDAY(TODAY())-1,ROUNDDOWN(${c},0)-TODAY()<=7-WEEKDAY(TODAY()))`,
  lastWeek: (c) => `AND(TODAY()-ROUNDDOWN(${c},0)>=(WEEKDAY(TODAY())),TODAY()-ROUNDDOWN(${c},0)<(WEEKDAY(TODAY())+7))`,
  nextWeek: (c) => `AND(ROUNDDOWN(${c},0)-TODAY()>(7-WEEKDAY(TODAY())),ROUNDDOWN(${c},0)-TODAY()<(15-WEEKDAY(TODAY())))`,
  thisMonth: (c) => `AND(MONTH(${c})=MONTH(TODAY()),YEAR(${c})=YEAR(TODAY()))`,
  lastMonth: (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0-1)),YEAR(${c})=YEAR(EDATE(TODAY(),0-1)))`,
  nextMonth: (c) => `AND(MONTH(${c})=MONTH(EDATE(TODAY(),0+1)),YEAR(${c})=YEAR(EDATE(TODAY(),0+1)))`,
};

const X14_ICONS = new Set(['3Stars', '3Triangles', '5Boxes']);
const cfvoXml = (list, x14 = false) => list.map((c) => {
  const t = c.type === 'num' ? 'num' : c.type;
  const v = c.v === undefined || c.v === null || c.v === '' ? null : String(c.v).startsWith('=') ? exportFormula(String(c.v)) : String(c.v);
  const type = String(c.v ?? '').startsWith('=') ? 'formula' : t;
  const gte = c.gte === false ? ' gte="0"' : '';
  if (x14) return `<x14:cfvo type="${type}"${gte}>${v !== null ? `<xm:f>${esc(v)}</xm:f>` : ''}</x14:cfvo>`;
  return `<cfvo type="${type === 'autoMin' ? 'min' : type === 'autoMax' ? 'max' : type}"${v !== null && !/^(auto)?(min|max)$/i.test(type) ? ` val="${esc(v)}"` : ''}${gte}/>`;
}).join('');

/** 데이터 막대 · 새 아이콘 집합의 엑셀 2010 확장(x14) 규칙 */
function cfX14(rule, id) {
  const sqref = [rule, ...(rule.more ?? [])].map((g) => rangeRef(g)).join(' ');
  if (rule.type === 'bar') {
    const cf = rule.cfvo ?? [{ type: 'autoMin' }, { type: 'autoMax' }];
    return `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="dataBar" id="${id}"><x14:dataBar minLength="0" maxLength="100"${rule.gradient === false ? ' gradient="0"' : ''}${rule.iconOnly ? ' showValue="0"' : ''}>${cfvoXml(cf, true)}<x14:fillColor rgb="${argb(rule.color ?? '#638ec6')}"/><x14:negativeFillColor rgb="${argb(rule.negColor ?? '#ff0000')}"/><x14:axisColor rgb="FF000000"/></x14:dataBar></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;
  }
  const n = Number(String(rule.icons)[0]) || 3;
  const cf = rule.cfvo ?? [...Array(n)].map((_, i) => ({ type: 'percent', v: Math.round((i * 100) / n) }));
  return `<x14:conditionalFormatting xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main"><x14:cfRule type="iconSet" priority="1" id="${id}"><x14:iconSet iconSet="${rule.icons}"${rule.reverse ? ' reverse="1"' : ''}${rule.iconOnly ? ' showValue="0"' : ''}>${cfvoXml(cf, true)}</x14:iconSet></x14:cfRule><xm:sqref>${sqref}</xm:sqref></x14:conditionalFormatting>`;
}

function cfXml(rule, pool, priority, x14 = null, date1904 = false) {
  const ref = [rule, ...(rule.more ?? [])].map((g) => rangeRef(g)).join(' ');
  const top = cellName(rule.r1, rule.c1);
  const lit = (v) => {
    const t = String(v ?? '');
    if (t.startsWith('=')) return exportFormula(t);
    const p = parseInput(t, date1904);
    return typeof p.value === 'number' ? String(p.value) : `"${t.replace(/"/g, '""')}"`;
  };
  const dx = () => pool.dxf(rule.style ?? {});
  const stop = rule.stopIfTrue ? ' stopIfTrue="1"' : '';
  const head = (type, extra = '') => `<cfRule type="${type}" dxfId="${dx()}" priority="${priority}"${stop}${extra}`;
  const f = (x) => `<formula>${esc(x)}</formula>`;
  const text = String(rule.v1 ?? '');
  const q = `"${text.replace(/"/g, '""')}"`;
  let body;
  switch (rule.type) {
    case 'gt': case 'lt': case 'ge': case 'le': case 'eq': case 'ne':
      body = `${head('cellIs', ` operator="${CELL_IS[rule.type]}"`)}>${f(lit(rule.v1))}</cfRule>`;
      break;
    case 'between': case 'notBetween':
      body = `${head('cellIs', ` operator="${CELL_IS[rule.type]}"`)}>${f(lit(rule.v1))}${f(lit(rule.v2))}</cfRule>`;
      break;
    case 'text': body = `${head('containsText', ` operator="containsText" text="${esc(text)}"`)}>${f(`NOT(ISERROR(SEARCH(${q},${top})))`)}</cfRule>`; break;
    case 'notText': body = `${head('notContainsText', ` operator="notContains" text="${esc(text)}"`)}>${f(`ISERROR(SEARCH(${q},${top}))`)}</cfRule>`; break;
    case 'begins': body = `${head('beginsWith', ` operator="beginsWith" text="${esc(text)}"`)}>${f(`LEFT(${top},LEN(${q}))=${q}`)}</cfRule>`; break;
    case 'ends': body = `${head('endsWith', ` operator="endsWith" text="${esc(text)}"`)}>${f(`RIGHT(${top},LEN(${q}))=${q}`)}</cfRule>`; break;
    case 'blank': body = `${head('containsBlanks')}>${f(`LEN(TRIM(${top}))=0`)}</cfRule>`; break;
    case 'noBlank': body = `${head('notContainsBlanks')}>${f(`LEN(TRIM(${top}))>0`)}</cfRule>`; break;
    case 'errors': body = `${head('containsErrors')}>${f(`ISERROR(${top})`)}</cfRule>`; break;
    case 'noErrors': body = `${head('notContainsErrors')}>${f(`NOT(ISERROR(${top}))`)}</cfRule>`; break;
    case 'date': {
      const p = TIME_FORMULA[rule.period] ? rule.period : 'today';
      body = `${head('timePeriod', ` timePeriod="${p}"`)}>${f(TIME_FORMULA[p](top))}</cfRule>`;
      break;
    }
    case 'formula': body = `${head('expression')}>${f(exportFormula(String(rule.formula ?? '=FALSE')))}</cfRule>`; break;
    case 'dup': body = `${head('duplicateValues')}/>`; break;
    case 'unique': body = `${head('uniqueValues')}/>`; break;
    case 'top': case 'bottom':
      body = `${head('top10', `${rule.percent ? ' percent="1"' : ''}${rule.type === 'bottom' ? ' bottom="1"' : ''} rank="${Number(rule.v1) || 10}"`)}/>`;
      break;
    case 'aboveAvg': case 'belowAvg': {
      const ex = `${rule.type === 'belowAvg' ? ' aboveAverage="0"' : ''}${rule.equal ? ' equalAverage="1"' : ''}${rule.stdDev ? ` stdDev="${Number(rule.stdDev)}"` : ''}`;
      body = `${head('aboveAverage', ex)}/>`;
      break;
    }
    case 'bar': {
      const cf = rule.cfvo ?? [{ type: 'min' }, { type: 'max' }];
      const id = `{${(0x10000000 + priority).toString(16).toUpperCase()}-0000-4000-8000-${String(priority).padStart(12, '0')}}`;
      x14?.push(cfX14(rule, id));
      body = `<cfRule type="dataBar" priority="${priority}"><dataBar${rule.iconOnly ? ' showValue="0"' : ''}>${cfvoXml(cf)}<color rgb="${argb(rule.color ?? '#638ec6')}"/></dataBar>${x14 ? `<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${NS_X14}"><x14:id>${id}</x14:id></ext></extLst>` : ''}</cfRule>`;
      break;
    }
    case 'scale': {
      const cs = rule.colors;
      const cf = rule.cfvo?.length === cs.length ? cfvoXml(rule.cfvo)
        : cs.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
      body = `<cfRule type="colorScale" priority="${priority}"><colorScale>${cf}${cs.map((c) => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`;
      break;
    }
    case 'icons': {
      const name = rule.icons ?? '3Arrows';
      if (X14_ICONS.has(name)) {
        // 별·삼각형·상자는 엑셀 2010 확장에만 저장됨
        if (!x14) return '';
        x14.push(cfX14(rule, `{${(0x20000000 + priority).toString(16).toUpperCase()}-0000-4000-8000-${String(priority).padStart(12, '0')}}`));
        return '';
      }
      const n = Number(String(name)[0]) || 3;
      const cf = rule.cfvo?.length === n ? cfvoXml(rule.cfvo) : [...Array(n)].map((_, i) => `<cfvo type="percent" val="${Math.round((i * 100) / n)}"/>`).join('');
      body = `<cfRule type="iconSet" priority="${priority}"><iconSet iconSet="${esc(name)}"${rule.reverse ? ' reverse="1"' : ''}${rule.iconOnly ? ' showValue="0"' : ''}>${cf}</iconSet></cfRule>`;
      break;
    }
    default: return '';
  }
  return `<conditionalFormatting${rule.pivot ? ' pivot="1"' : ''} sqref="${ref}">${body}</conditionalFormatting>`;
}

function chartXml(wb, si, chart, fileName = 'Book1.xlsx', imageRel) {
  const inferred = inferPivotCategorySeries(wb, si, chart);
  if (inferred) chart = { ...chart, series: inferred };
  const srcIndex = chart.sheet ? wb.sheetIndexByName(chart.sheet) : si;
  const s = srcIndex >= 0 ? srcIndex : si;
  const refText = (sheetIdx, r1, c1, r2, c2) => `${quoteSheetName(wb.sheets[sheetIdx].name)}!${rangeRef({ r1, c1, r2, c2 }, true)}`;
  const label = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? formatGeneral(v) : isError(v) ? v.code : String(v));
  const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(label(v))}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (vals, code = 'General') => `<c:numCache><c:formatCode>${esc(code)}</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
  // 참조는 원래 범위를 유지하므로 범주 캐시와 dPt 번호도 원본 순서로 쓴다.
  // hiddenCats는 WIXEL 보조 옵션으로만 복원하며 Excel의 범주 필터를 흉내내지 않는다.
  let sourceRows = null;
  const data = chartModelData(wb, si, { ...chart, hiddenCats: undefined, ...(isChartEx(chart) ? { hiddenSeries: undefined } : {}) }, { sample: false, onRange: rows => { sourceRows = rows; } });
  // 엑셀 2016 차트(폭포 · 깔때기 · 히스토그램 · 파레토 · 트리맵 · 상자 수염)는 호환 차트로 저장하고 원래 종류는 확장 정보로 보관
  const FALLBACK = { waterfall: 'column', histogram: 'column', pareto: 'column', treemap: 'column', boxWhisker: 'column', funnel: 'bar', sunburst: 'column' };
  const baseType = chart.type === 'combo' ? 'column' : FALLBACK[chart.type] ?? chart.type;
  const threeD = !!chart.threeD && ['column', 'bar', 'pie', 'area', 'line', 'surface'].includes(chart.type);
  // 계열마다 { tx, cat, val } 참조 (+ 값 캐시)
  const refs = [];
  if (chart.pivot) {
    const ps = chart.pivot.sheet ? wb.sheetIndexByName(chart.pivot.sheet) : si;
    const sh = wb.sheets[ps];
    const def = [sh?.pivot, ...(sh?.pivotsExtra ?? [])].filter(Boolean).find((d) => d.name === chart.pivot.name);
    const rowsIdx = data.rows ?? [];
    const contiguous = rowsIdx.length && rowsIdx.every((r, i) => !i || r === rowsIdx[i - 1] + 1);
    const top = def?.top ?? 0;
    const left = def?.left ?? 0;
    data.series.forEach((sr) => {
      refs.push(def && contiguous ? {
        tx: refText(ps, top + sr.headRow, left + sr.col, top + sr.headRow, left + sr.col),
        cat: refText(ps, top + rowsIdx[0], left, top + rowsIdx.at(-1), left),
        val: refText(ps, top + rowsIdx[0], left + sr.col, top + rowsIdx.at(-1), left + sr.col),
      } : {});
    });
  } else if (chart.series?.length) {
    for (const sr of chart.series) {
      // 정의된 이름 참조는 그대로 ([0]!이름 = 이 통합 문서의 이름, 시트 범위 이름은 시트!이름)
      const ref = (r) => (!r ? null : r.name ? (r.sheet ? `${quoteSheetName(r.sheet)}!${r.name}` : `[0]!${r.name}`) : refText(r.sheet ? Math.max(0, wb.sheetIndexByName(r.sheet)) : s, r.r1, r.c1, r.r2, r.c2));
      refs.push({ tx: ref(sr.name?.ref), cat: ref(sr.cat ?? sr.x), val: ref(sr.val), sz: ref(sr.size) });
    }
  } else if (chart.range) {
    const rg = chart.range;
    // 레이아웃 판정은 이미 읽은 원본 데이터로 수행해 같은 범위를 두 번 적재하지 않는다.
    const rows = sourceRows ?? [];
    const L = chartLayout(rows, chart.type, !!chart.byRows);
    if (chart.type === 'surface') {
      const header = typeof rows[0]?.[0] !== 'number', off = header ? 1 : 0;
      if (chart.byRows) {
        L.byCols = true; L.firstDataRow = off; L.firstDataCol = off; L.catCol = header; L.headRow = header;
      } else {
        L.byCols = false; L.firstDataRow = off; L.firstDataCol = off; L.catCol = header; L.headRow = header;
      }
    }
    if (chart.type === 'sunburst' || chart.type === 'treemap') {
      L.byCols = !chart.byRows;
      if (L.byCols) { L.firstDataCol = rows[0].length - 1; L.firstDataRow = typeof rows[0].at(-1) === 'number' ? 0 : 1; L.catCol = true; L.headRow = L.firstDataRow === 1; }
      else { L.firstDataRow = rows.length - 1; L.firstDataCol = typeof rows.at(-1)[0] === 'number' ? 0 : 1; L.catCol = L.firstDataCol === 1; L.headRow = true; }
    }
    const dr0 = rg.r1 + L.firstDataRow;
    const dc0 = rg.c1 + L.firstDataCol;
    const all = [];
    if (L.byCols) {
      for (let c = dc0; c <= Math.min(rg.c2, rg.c1 + (rows[0]?.length ?? 0) - 1); c++) {
        all.push({ tx: L.headRow ? refText(s, rg.r1, c, rg.r1, c) : null, cat: L.catCol ? refText(s, dr0, rg.c1, rg.r2, rg.c1 + Math.max(0, L.firstDataCol - 1)) : null, val: refText(s, dr0, c, rg.r2, c), nums: rows.slice(L.firstDataRow).some((r) => typeof r[c - rg.c1] === 'number') });
      }
    } else {
      for (let r = dr0; r <= Math.min(rg.r2, rg.r1 + rows.length - 1); r++) {
        all.push({ tx: L.catCol ? refText(s, r, rg.c1, r, rg.c1) : null, cat: L.headRow ? refText(s, rg.r1, dc0, ['sunburst', 'treemap'].includes(chart.type) ? rg.r2 - 1 : rg.r1, rg.c2) : null, val: refText(s, r, dc0, r, rg.c2), nums: rows[r - rg.r1].slice(L.firstDataCol).some((v) => typeof v === 'number') });
      }
    }
    const sourceRefs = chart.type === 'bubble' ? all.filter((_, i) => i % 2 === 0).map((r, i) => ({ ...r, sz: all[2 * i + 1]?.val })) : all;
    const withNums = sourceRefs.filter((x) => x.nums);
    pushAll(refs, (withNums.length ? withNums : sourceRefs));
  }
  sourceRows = null;
  const pal = paletteOf(chart);
  if (isChartEx(chart)) return writeChartEx(chart, data, refs, pal, imageRel);
  const series = data.series.map((sr, i) => ({
    ...sr, ...refs[chart.pivot ? i : sr._fi ?? i], type: chart.type === 'stock' && chart.volume ? (i === 0 ? 'column' : 'stock') : FALLBACK[sr.type] ?? sr.type ?? baseType,
    axis: chart.type === 'stock' && chart.volume ? (i === 0 ? 0 : 1) : sr.axis ?? 0, color: sr.color ?? pal[(sr._fi ?? i) % pal.length],
  }));
  const LBL_POS = { center: 'ctr', insideEnd: 'inEnd', insideBase: 'inBase', outEnd: 'outEnd', above: 't', below: 'b', left: 'l', right: 'r' };
  const dLbls = (on, code, pct = false, pos = null, sf = {}) => {
    if (on === undefined && !pct && sf.catName === undefined && sf.serName === undefined) return '';
    const font = chartTextPropertiesXml({ ...chartTextFields(chart, '', true), ...chartTextFields(sf, 'label') });
    const pointLabels = (sf.values ?? []).map((_, i) => {
      const p = sf._pi?.[i] ?? i, style = sf.pointLabelStyles?.[p];
      const text = style && chartTextPropertiesXml({ ...chartTextFields(chart, '', true), ...chartTextFields(sf, 'label'), ...chartTextFields(style) });
      return text ? `<c:dLbl><c:idx val="${i}"/>${text}</c:dLbl>` : '';
    }).join('');
    return `<c:dLbls>${pointLabels}${code && typeof code === 'string' ? `<c:numFmt formatCode="${esc(code)}" sourceLinked="0"/>` : ''}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${font}${pos && (LBL_POS[pos] || pos === 'out') ? `<c:dLblPos val="${LBL_POS[pos] ?? 'outEnd'}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="${on ? 1 : 0}"/><c:showCatName val="${sf.catName ? 1 : 0}"/><c:showSerName val="${sf.serName ? 1 : 0}"/><c:showPercent val="${pct ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`;
  };
  const serXml = (sr, i) => {
    const type = sr.type;
    const scatter = type === 'scatter' || type === 'bubble';
    const pie = ['pie', 'doughnut', 'pieOfPie', 'barOfPie'].includes(type);
    const pieOutline = ['pie', 'doughnut'].includes(type) ? hex6(sr.outline ?? '#ffffff') : 'FFFFFF';
    const hex = String(sr.color).replace('#', '').toUpperCase().slice(0, 6);
    const fill = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
    const tx = sr.tx ? `<c:tx><c:strRef><c:f>${esc(sr.tx)}</c:f>${strCache([sr.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${esc(sr.name)}</c:v></c:tx>`;
    let spPr;
    if (type === 'bubble') spPr = `<c:spPr>${fill}</c:spPr>`;
    else if (type === 'stock') spPr = '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    else if (scatter) spPr = /line|smooth/i.test(chart.scatterStyle ?? '') ? `<c:spPr><a:ln w="${sr.lineWidth ? Math.round(sr.lineWidth * 9525) : 19050}">${fill}${DASH_XML[sr.dash] ?? ''}</a:ln></c:spPr>` : '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    else if (type === 'radar') spPr = `<c:spPr>${chart.radarStyle === 'filled' ? fill : ''}<a:ln w="${sr.lineWidth ? Math.round(sr.lineWidth * 9525) : 28575}">${fill}${DASH_XML[sr.dash] ?? ''}</a:ln></c:spPr>`;
    else if (type === 'line') spPr = `<c:spPr><a:ln w="${sr.lineWidth ? Math.round((sr.lineWidth * 3 / 4) * 12700) : 28575}" cap="rnd">${fill}${DASH_XML[sr.dash] ?? ''}<a:round/></a:ln></c:spPr>`;
    else if (pie) spPr = `<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="${pieOutline}"/></a:solidFill></a:ln></c:spPr>`;
    else spPr = `<c:spPr>${fill}${sr.outline ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${String(sr.outline).replace('#', '').toUpperCase().slice(0, 6)}"/></a:solidFill></a:ln>` : ''}</c:spPr>`;
    const markerSym = sr.marker === 'none' || sr.marker === false ? 'none' : typeof sr.marker === 'string' ? sr.marker : 'circle';
    const markerSym2 = sr.marker === undefined && (chart.marker === 'none' || (scatter && /^(line|smooth)$/.test(chart.scatterStyle ?? '')) || (type === 'radar' && chart.radarStyle !== 'marker') || type === 'stock') ? 'none' : markerSym;
    const marker = (type === 'line' || type === 'radar' || type === 'stock' || (scatter && type !== 'bubble')) ? (markerSym2 === 'none' ? '<c:marker><c:symbol val="none"/></c:marker>' : `<c:marker><c:symbol val="${markerSym2}"/><c:size val="${Math.round(Math.max(2,Math.min(72,sr.markerSize ?? 5)))}"/><c:spPr>${fill}<a:ln w="9525">${fill}</a:ln></c:spPr></c:marker>`) : '';
    const explosion = value => Number.isFinite(value) ? Math.round(Math.max(0, Math.min(400, value))) : null;
    // 한 idx에 색/분리 옵션을 한 dPt로 합친다. 0도 계열 분리의 명시적 덮어쓰기다.
    const dPt = pie ? sr.values.map((_, k) => {
      const p = sr._pi?.[k] ?? k, offset = explosion(sr.pointExplosion?.[p]);
      return `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/>${offset !== null ? `<c:explosion val="${offset}"/>` : ''}<c:spPr><a:solidFill><a:srgbClr val="${(sr.pointColors?.[p] ?? sr.colors?.[p] ?? chart.seriesFmt?.[sr._fi ?? i]?.color ?? pal[p % pal.length]).replace('#', '')}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="${pieOutline}"/></a:solidFill></a:ln></c:spPr></c:dPt>`;
    }).join('') : '';
    const invert = type === 'column' || type === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
    // 막대 · 꺾은선의 데이터 요소별 색 · '요소마다 다른 색'
    const ptColor = (k) => { const p = sr._pi?.[k] ?? k; return sr.pointColors?.[p] ?? (chart.varyColors && !chart.seriesFmt?.[sr._fi ?? i]?.color && (type === 'column' || type === 'bar') ? pal[p % pal.length] : null); };
    const dPtBar = !pie && (sr.pointColors || chart.varyColors) && (type === 'column' || type === 'bar') ? sr.values.map((_, k) => (ptColor(k) ? `<c:dPt><c:idx val="${k}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${String(ptColor(k)).replace('#', '').toUpperCase().slice(0, 6)}"/></a:solidFill></c:spPr></c:dPt>` : '')).join('') : '';
    // 추세선
    const TREND = { linear: 'linear', exp: 'exp', movingAvg: 'movingAvg' };
    const trend = sr.trend && TREND[sr.trend] && !pie ? `<c:trendline><c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="${String(sr.trendColor ?? sr.color).replace('#', '').toUpperCase().slice(0, 6)}"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr><c:trendlineType val="${TREND[sr.trend]}"/>${sr.trend === 'movingAvg' ? `<c:period val="${Math.max(2, sr.trendPeriod ?? 3)}"/>` : sr.trendForward ? `<c:forward val="${sr.trendForward}"/>` : ''}<c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>` : '';
    const code = typeof sr.numFmt === 'string' ? sr.numFmt : null;
    // 원형: 레이블을 따로 정하지 않았으면 백분율 (WIXEL 화면과 같게)
    const pieP = !!sr.pct || pie && (sr.labels ?? chart.labels) === undefined && !sr.catName && !sr.serName;
    const labels = type === 'surface' ? '' : dLbls(sr.labels ?? chart.labels, code, pieP, sr.labelPos, sr);
    const cats = data.categories;
    const catTag = scatter ? 'xVal' : 'cat';
    // 다단계 항목 (여러 열의 항목 범위): 안쪽 → 바깥 순서의 lvl
    const lvls = !scatter && sr.cat && data.catLevels?.length
      ? [cats.map((v, k) => ({ k, v })), ...data.catLevels.map((spans) => spans.filter((sp) => sp.text !== '').map((sp) => ({ k: sp.start, v: sp.text })))]
      : null;
    const cat = scatter
      ? `<c:xVal>${sr.cat ? `<c:numRef><c:f>${esc(sr.cat)}</c:f>${numCache(sr.x ?? [])}</c:numRef>` : `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${sr.values.length}"/>${(sr.x ?? sr.values.map((_, k) => k + 1)).map((v, k) => typeof v === 'number' ? `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>` : '').join('')}</c:numLit>`}</c:xVal>`
      : lvls
      ? `<c:cat><c:multiLvlStrRef><c:f>${esc(sr.cat)}</c:f><c:multiLvlStrCache><c:ptCount val="${cats.length}"/>${lvls.map((pts) => `<c:lvl>${pts.map((x) => `<c:pt idx="${x.k}"><c:v>${esc(x.v)}</c:v></c:pt>`).join('')}</c:lvl>`).join('')}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`
      : sr.cat
      ? `<c:${catTag}><c:strRef><c:f>${esc(sr.cat)}</c:f>${strCache(cats)}</c:strRef></c:${catTag}>`
      : cats.length ? `<c:${catTag}><c:strLit><c:ptCount val="${cats.length}"/>${cats.map((v, k) => `<c:pt idx="${k}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strLit></c:${catTag}>` : '';
    const valTag = scatter ? 'yVal' : 'val';
    const val = sr.val
      ? `<c:${valTag}><c:numRef><c:f>${esc(sr.val)}</c:f>${numCache(sr.values, code ?? 'General')}</c:numRef></c:${valTag}>`
      : `<c:${valTag}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${sr.values.length}"/>${sr.values.map((v, k) => (typeof v === 'number' ? `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numLit></c:${valTag}>`;
    const bsz = type === 'bubble' ? `<c:bubbleSize>${sr.sz ? `<c:numRef><c:f>${esc(sr.sz)}</c:f>${numCache(sr.size ?? [])}</c:numRef>` : `<c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${sr.values.length}"/>${(sr.size ?? []).map((v, k) => (typeof v === 'number' ? `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numLit>`}</c:bubbleSize><c:bubble3D val="${chart.threeD ? 1 : 0}"/>` : '';
    const smooth = type === 'line' || (scatter && type !== 'bubble') ? `<c:smooth val="${(sr.smooth ?? /smooth/i.test(chart.scatterStyle ?? '')) ? 1 : 0}"/>` : '';
    const seriesExplosion = explosion(sr.explode ?? chart.explode);
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${pie && seriesExplosion !== null ? `<c:explosion val="${seriesExplosion}"/>` : ''}${invert}${marker}${dPt}${dPtBar}${labels}${trend}${cat}${val}${bsz}${smooth}${threeD && (type === 'column' || type === 'bar') && ['box', 'cylinder', 'cone', 'pyramid'].includes(sr.barShape) ? `<c:shape val="${sr.barShape}"/>` : ''}</c:ser>`;
  };
  const pieLike = ['pie', 'doughnut', 'pieOfPie', 'barOfPie'].includes(baseType);
  const grouping = chart.grouping ?? 'clustered';
  // 계열 종류 · 축별로 차트 그룹 (콤보 차트)
  const groups = [];
  series.forEach((sr, i) => {
    const kind = pieLike ? baseType : sr.type === 'bar' || sr.type === 'column' ? (baseType === 'bar' ? 'bar' : 'column') : sr.type;
    const axis = pieLike ? 0 : sr.axis === 1 ? 1 : 0;
    const mode = sr.grouping ?? grouping;
    let g = groups.find((x) => x.kind === kind && x.axis === axis && x.grouping === mode);
    if (!g) { g = { kind, axis, grouping: mode, xml: [] }; groups.push(g); }
    g.xml.push(serXml(sr, i));
  });
  if (!groups.length) groups.push({ kind: baseType, axis: 0, xml: [] });
  const hasSecondary = groups.some((g) => g.axis === 1);
  const ax = (axis) => (axis === 1 && hasSecondary ? '<c:axId val="333333333"/><c:axId val="444444444"/>' : '<c:axId val="111111111"/><c:axId val="222222222"/>');
  const groupXml = groups.map((g) => {
    const body = g.xml.join('');
    const grouping = g.grouping ?? chart.grouping ?? 'clustered';
    const a = ax(g.axis);
    switch (g.kind) {
      case 'bar': case 'column': {
        const stackedG = grouping === 'stacked' || grouping === 'percentStacked';
        if (threeD) return `<c:bar3DChart><c:barDir val="${g.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="${chart.varyColors ? 1 : 0}"/>${body}<c:gapWidth val="${Math.round(Math.max(0, Math.min(500, chart.gap ?? 150)))}"/><c:gapDepth val="150"/><c:shape val="${['box', 'cylinder', 'cone', 'pyramid'].includes(chart.barShape) ? chart.barShape : 'box'}"/>${a}${grouping === 'standard' ? '<c:axId val="555555555"/>' : ''}</c:bar3DChart>`;
        return `<c:barChart><c:barDir val="${g.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="${chart.varyColors ? 1 : 0}"/>${body}<c:gapWidth val="${typeof chart.gap === 'number' ? Math.round(chart.gap) : g.kind === 'bar' ? 182 : 150}"/>${stackedG ? '<c:overlap val="100"/>' : typeof chart.overlap === 'number' ? `<c:overlap val="${Math.round(Math.max(-100, Math.min(100, chart.overlap)))}"/>` : ''}${a}</c:barChart>`;
      }
      case 'line':
        if (threeD) return `<c:line3DChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${body}<c:gapDepth val="150"/>${a}<c:axId val="555555555"/></c:line3DChart>`;
        return `<c:lineChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${body}<c:marker val="1"/>${a}</c:lineChart>`;
      case 'area':
        if (threeD) return `<c:area3DChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${body}<c:gapDepth val="150"/>${a}</c:area3DChart>`;
        return `<c:areaChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${body}${a}</c:areaChart>`;
      case 'pie':
        if (threeD) return `<c:pie3DChart><c:varyColors val="1"/>${body}</c:pie3DChart>`;
        return `<c:pieChart><c:varyColors val="1"/>${body}<c:firstSliceAng val="${Math.round(chart.firstAngle ?? 0)}"/></c:pieChart>`;
      case 'pieOfPie': case 'barOfPie': {
        const split = ['position', 'value', 'percent', 'custom'].includes(chart.splitType) ? chart.splitType : 'position';
        const points = split === 'custom' ? `<c:custSplit>${(chart.splitPoints ?? []).filter((i) => Number.isInteger(i) && i >= 0 && i < data.categories.length).map((i) => `<c:secondPiePt val="${i}"/>`).join('')}</c:custSplit>` : '';
        return `<c:ofPieChart><c:ofPieType val="${g.kind === 'barOfPie' ? 'bar' : 'pie'}"/><c:varyColors val="1"/>${body}<c:gapWidth val="${Math.round(Math.max(0,Math.min(500,chart.splitGap ?? 150)))}"/><c:splitType val="${({ position: 'pos', value: 'val', percent: 'percent', custom: 'cust' })[split]}"/><c:splitPos val="${Number.isFinite(chart.splitPos) ? chart.splitPos : 3}"/>${points}<c:secondPieSize val="${Math.round(Math.max(5,Math.min(200,chart.secondSize ?? 75)))}"/><c:serLines/></c:ofPieChart>`;
      }
      case 'surface': {
        const tag = threeD ? 'surface3DChart' : 'surfaceChart';
        return `<c:${tag}><c:wireframe val="${/wireframe/i.test(chart.surfaceStyle ?? '') ? 1 : 0}"/>${body}${a}<c:axId val="555555555"/></c:${tag}>`;
      }
      case 'doughnut': return `<c:doughnutChart><c:varyColors val="1"/>${body}<c:firstSliceAng val="${Math.round(chart.firstAngle ?? 0)}"/><c:holeSize val="${Math.round(chart.hole ?? 50)}"/></c:doughnutChart>`;
      case 'scatter': return `<c:scatterChart><c:scatterStyle val="${['marker','line','lineMarker','smooth','smoothMarker'].includes(chart.scatterStyle) ? chart.scatterStyle : 'marker'}"/><c:varyColors val="0"/>${body}${a}</c:scatterChart>`;
      case 'bubble': return `<c:bubbleChart><c:varyColors val="0"/>${body}<c:bubbleScale val="${Math.round(chart.bubbleScale ?? 100)}"/><c:showNegBubbles val="${chart.showNegBubbles ? 1 : 0}"/>${a}</c:bubbleChart>`;
      case 'radar': return `<c:radarChart><c:radarStyle val="${['filled','marker'].includes(chart.radarStyle) ? chart.radarStyle : 'standard'}"/><c:varyColors val="0"/>${body}${a}</c:radarChart>`;
      case 'stock': return `<c:stockChart>${body}<c:hiLowLines/>${chart.ohlc ? '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>' : ''}${a}</c:stockChart>`;
      default: return '';
    }
  }).join('');
  const horizontal = baseType === 'bar';
  const axTitle = cfg => cfg?.title ? `<c:title>${chartRichTextXml(cfg.title, { ...chartTextFields(chart, '', true), ...chartTextFields(cfg, 'title') })}<c:overlay val="0"/></c:title>` : '';
  const axisText = (cfg, rotation = null) => chartTextPropertiesXml({ ...chartTextFields(chart, '', true), ...(Number.isFinite(chart.axisSize) ? { size: chart.axisSize } : {}), ...chartTextFields(cfg) }, 'c', rotation);
  const scaling = (cfg) => `<c:scaling>${Number.isFinite(cfg?.logBase) && cfg.logBase >= 2 && cfg.logBase <= 1000 ? `<c:logBase val="${cfg.logBase}"/>` : ''}<c:orientation val="${cfg?.reverse ? 'maxMin' : 'minMax'}"/>${typeof cfg?.max === 'number' ? `<c:max val="${cfg.max}"/>` : ''}${typeof cfg?.min === 'number' ? `<c:min val="${cfg.min}"/>` : ''}</c:scaling>`;
  const numFmt = (cfg) => (cfg?.numFmt ? `<c:numFmt formatCode="${esc(cfg.numFmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>');
  // Keep explicit category-label settings in standard chart XML, including 0°.
  // Missing values retain Excel's automatic interval and text direction.
  const categoryLabels = chart.axes?.x ?? {};
  const labelIntervalXml = Number.isSafeInteger(categoryLabels.labelInterval) && categoryLabels.labelInterval > 0 ? `<c:tickLblSkip val="${categoryLabels.labelInterval}"/>` : '';
  const labelRotationXml = axisText(categoryLabels, categoryLabels.labelRotation);
  const catAxis = (id, cross, pos, del) => (baseType === 'scatter' || baseType === 'bubble'
    ? `<c:valAx><c:axId val="${id}"/>${scaling(chart.axes?.x)}<c:delete val="${del || chart.axes?.x?.hide ? 1 : 0}"/><c:axPos val="${pos}"/>${del ? '' : axTitle(chart.axes?.x)}${numFmt(chart.axes?.x)}<c:tickLblPos val="nextTo"/>${axisText(chart.axes?.x)}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${del || chart.axes?.x?.hide ? 1 : 0}"/><c:axPos val="${pos}"/>${chart.gridX ? '<c:majorGridlines/>' : ''}${del ? '' : axTitle(chart.axes?.x)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${labelRotationXml}<c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/>${labelIntervalXml}<c:noMultiLvlLbl val="0"/></c:catAx>`);
  const valAxis = (id, cross, pos, cfg, grid, crosses = 'autoZero') => `<c:valAx><c:axId val="${id}"/>${scaling(cfg)}<c:delete val="${cfg?.hide ? 1 : 0}"/><c:axPos val="${pos}"/>${grid && chart.gridY !== false ? '<c:majorGridlines/>' : ''}${axTitle(cfg)}${numFmt(cfg)}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/>${axisText(cfg)}<c:crossAx val="${cross}"/><c:crosses val="${crosses}"/><c:crossBetween val="${baseType === 'area' || baseType === 'scatter' || baseType === 'bubble' ? 'midCat' : 'between'}"/>${typeof cfg?.major === 'number' ? `<c:majorUnit val="${cfg.major}"/>` : ''}</c:valAx>`;
  let axesXml = '';
  if (!pieLike) {
    axesXml = catAxis(111111111, 222222222, horizontal ? 'l' : 'b', false) + valAxis(222222222, 111111111, horizontal ? 'b' : 'l', chart.axes?.y, true);
    if (hasSecondary) axesXml += catAxis(333333333, 444444444, horizontal ? 'l' : 'b', true) + valAxis(444444444, 333333333, horizontal ? 't' : 'r', chart.axes?.y2, false, 'max');
  }
  if (baseType === 'surface' || threeD && (baseType === 'line' || ['column', 'bar'].includes(baseType) && groups.some(g => g.grouping === 'standard'))) axesXml += `<c:serAx><c:axId val="555555555"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:tickLblSkip val="1"/><c:tickMarkSkip val="1"/></c:serAx>`;
  const titleStyle = { size: chart.size ?? 14, bold: chart.bold ?? false, ...chartTextFields(chart, '', true), ...chartTextFields(chart, 'title') };
  const title = chart.title
    ? `<c:title>${chartRichTextXml(chart.title, titleStyle)}${chartLayoutXml(chart.titleLayout)}<c:overlay val="${chart.titleOverlay ? 1 : 0}"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const lp = chart.legend ?? (series.length > 1 || pieLike ? 'b' : 'none');
  // 범례 글꼴 (색 · 크기 · 굵게)
  const legTx = chartTextPropertiesXml({ ...chartTextFields(chart, '', true), ...chartTextFields(chart, 'legend') });
  const legend = lp !== 'none' ? `<c:legend><c:legendPos val="${lp}"/>${chartLayoutXml(chart.legendLayout)}<c:overlay val="0"/>${legTx}</c:legend>` : '';
  // 위셀 지표 선택 피벗 차트: 엑셀에는 피벗 테이블 범위를 참조하는 일반 차트로 (엑셀 피벗 차트는 모든 값 필드를 강제로 보이므로)
  const subsetPivot = !!chart.pivot?.values?.length;
  const pivotSrc = chart.pivot && !subsetPivot ? `<c:pivotSource><c:name>${esc(`[${fileName}]${quoteSheetName(chart.pivot.sheet ?? wb.sheets[si].name)}!${chart.pivot.name}`)}</c:name><c:fmtId val="0"/></c:pivotSource>` : '';
  const pivotFmts = chart.pivot && !subsetPivot ? `<c:pivotFmts>${series.map((_, i) => `<c:pivotFmt><c:idx val="${i}"/></c:pivotFmt>`).join('')}</c:pivotFmts>` : '';
  // 차트 영역 · 그림 영역 채우기와 테두리
  const hexOf = (c) => String(c).replace('#', '').toUpperCase().slice(0, 6);
  const areaSpPr = chart.chartAreaFormat ? chartAreaFormatXml(chart.chartAreaFormat, { imageRel, kind: 'chart' }) : chart.fill || chart.border ? `<c:spPr>${chart.fill ? `<a:solidFill><a:srgbClr val="${hexOf(chart.fill)}"/></a:solidFill>` : ''}${chart.border ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${hexOf(chart.border)}"/></a:solidFill></a:ln>` : ''}</c:spPr>` : '';
  const plotSpPr = chart.plotAreaFormat ? chartAreaFormatXml(chart.plotAreaFormat, { imageRel, kind: 'plot' }) : chart.plotFill ? `<c:spPr><a:solidFill><a:srgbClr val="${hexOf(chart.plotFill)}"/></a:solidFill></c:spPr>` : '';
  // WIXEL 전용 설정 (엑셀은 무시): 원래 차트 종류 · 팔레트 · 서식
  const TB_KEYS = ['chartStyle', 'mapLowColor', 'mapMidColor', 'mapHighColor', 'bandCount', 'showNegBubbles', 'surfaceStyle', 'volume', 'splitType', 'splitPos', 'splitPoints', 'secondSize', 'splitGap', 'bubble3D', 'comboAxis', 'comboLayout', 'threeD', 'view3D', 'titleSize', 'axisSize', 'hiddenSeries', 'hiddenCats', 'legendBold', 'type', 'byRows', 'fieldButtons', 'palette', 'scatterStyle', 'radarStyle', 'ohlc', 'explode', 'hole', 'gap', 'marker', 'gridX', 'gridY', 'fill', 'plotFill', 'border', 'titleColor', 'titleBold', 'textColor', 'gridColor', 'rounded', 'totals', 'binCount', 'binWidth', 'upColor', 'downColor', 'totalColor', 'bubbleScale', 'firstAngle', 'showMean', 'connectors'];
  const tb = Object.fromEntries(TB_KEYS.filter((k) => chart[k] !== undefined && chart[k] !== null).map((k) => [k, chart[k]]));
  // 계열은 위에서 실제로 제거했다. 압축된 ser 목록에 원래 번호를 다시 적용하면
  // 남은 계열까지 숨겨지므로 확장에 hiddenSeries를 중복 저장하지 않는다.
  delete tb.hiddenSeries;
  const textSupplement = chartTextSupplement(chart, series);
  if (textSupplement) tb.wxText = textSupplement;
  if (subsetPivot) tb.wxPivot = chart.pivot; // 위셀로 다시 열면 슬라이서와 연동되는 피벗 차트로 복원
  const extLst = Object.keys(tb).length > 1 || chart.type === 'combo' || FALLBACK[chart.type] ? `<c:extLst><c:ext uri="{5E2A6C7B-8F4D-4B1A-9C3E-7D6F1A2B3C4D}" xmlns:tb="urn:tabula:chart"><tb:props json="${esc(JSON.stringify(tb))}"/></c:ext></c:extLst>` : '';
  const v3 = chartView3D(baseType === 'surface' && !threeD ? { type: 'surface', view3D: { rotX: 90, rotY: 0, depthPercent: 100, rAngAx: true, perspective: 0 } } : chart);
  const viewXml = threeD || baseType === 'surface' ? `<c:view3D><c:rotX val="${Math.round(v3.rotX)}"/><c:rotY val="${Math.round((v3.rotY + (baseType === 'pie' ? chart.firstAngle ?? 0 : 0)) % 360)}"/><c:depthPercent val="${Math.round(v3.depthPercent)}"/><c:rAngAx val="${v3.rAngAx ? 1 : 0}"/><c:perspective val="${Math.round(v3.perspective)}"/></c:view3D>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}"><c:date1904 val="${wb.date1904 ? 1 : 0}"/><c:roundedCorners val="${chart.rounded ? 1 : 0}"/>${pivotSrc}<c:chart>${title}${pivotFmts}${viewXml}<c:plotArea><c:layout/>${groupXml}${axesXml}${chart.dataTable && !pieLike ? `<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/>${chartTextPropertiesXml({ ...chartTextFields(chart, '', true), ...chartTextFields(chart.dataTableText) })}</c:dTable>` : ''}${plotSpPr}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${areaSpPr}${chartTextPropertiesXml(chartTextFields(chart, '', true))}${extLst}</c:chartSpace>`;
}

const DASH_XML = { dash: '<a:prstDash val="dash"/>', dot: '<a:prstDash val="sysDot"/>', dashDot: '<a:prstDash val="dashDot"/>', longDash: '<a:prstDash val="lgDash"/>', sysDash: '<a:prstDash val="sysDash"/>' };
const DASH_FROM = { dash: 'dash', sysDot: 'dot', dot: 'dot', dashDot: 'dashDot', sysDashDot: 'dashDot', lgDash: 'longDash', sysDash: 'sysDash' };
const KIND_PRST = { arrow: 'rightArrow', textbox: 'rect', line: 'straightConnector1' };
const prstOf = (kind) => KIND_PRST[kind] ?? (GEOM[kind] || LINE_KINDS.has(kind) ? kind : 'rect');
const hex6 = (c) => { const h = String(c ?? '#000000').replace('#', ''); return (h.length === 3 ? [...h].map((x) => x + x).join('') : h).toUpperCase().padStart(6, '0').slice(0, 6); };
const shapeColorXml = (color, opacity = 1) => `<a:srgbClr val="${hex6(color)}">${opacity !== 1 ? `<a:alpha val="${Math.round(Math.max(0, Math.min(1, Number(opacity))) * 100000)}"/>` : ''}</a:srgbClr>`;
function shapeEffectsXml(sh) {
  let body = '';
  if (sh.glow) body += `<a:glow rad="${Math.round(Math.max(0, sh.glow.size ?? 5) * EMU)}">${shapeColorXml(sh.glow.color ?? '#4472c4', sh.glow.opacity ?? 0.6)}</a:glow>`;
  if (sh.shadow) {
    const sd = typeof sh.shadow === 'object' ? sh.shadow : {};
    const dx = sd.dx ?? 2.5, dy = sd.dy ?? 2.5;
    const angle = (Math.atan2(dy, dx) * 180 / Math.PI + 360) % 360;
    body += `<a:outerShdw blurRad="${Math.round(Math.max(0, sd.blur ?? 2.5) * EMU)}" dist="${Math.round(Math.hypot(dx, dy) * EMU)}" dir="${Math.round(angle * 60000) % 21600000}" algn="ctr" rotWithShape="1">${shapeColorXml(sd.color ?? '#000000', sd.opacity ?? 0.4)}</a:outerShdw>`;
  }
  if (sh.soft) body += `<a:softEdge rad="${Math.round(Math.max(0, sh.soft) * EMU)}"/>`;
  return body ? `<a:effectLst>${body}</a:effectLst>` : '';
}

/** 도형 → <xdr:sp> / <xdr:cxnSp> */
function shapeXml(sh, id, xfrm, hyperlinkXml = () => '') {
  const name = esc(sh.name || `${sh.kind === 'textbox' ? 'TextBox' : '도형'} ${id - 1}`);
  const description = (sh.alt !== undefined ? ` descr="${esc(sh.alt)}"` : '') + (sh.hidden ? ' hidden="1"' : '');
  let fill = sh.fill ? `<a:solidFill>${shapeColorXml(sh.fill, sh.fillOpacity ?? 1)}</a:solidFill>` : '<a:noFill/>';
  if (sh.grad && sh.fill) {
    const tint = (n) => `#${hex6(sh.fill).match(/../g).map((x) => { const v = parseInt(x, 16); return Math.round(n > 0 ? v + (255 - v) * n : v * (1 + n)).toString(16).padStart(2, '0'); }).join('')}`;
    const stops = sh.grad.stops?.length ? sh.grad.stops : [[0, tint(0.35)], [1, tint(-0.15)]];
    const angle = ((Number(sh.grad.ang ?? 90) % 360) + 360) % 360;
    fill = `<a:gradFill rotWithShape="1"><a:gsLst>${stops.map(([pos, color, opacity]) => `<a:gs pos="${Math.round(Math.max(0, Math.min(1, pos)) * 100000)}">${shapeColorXml(color, (opacity ?? 1) * (sh.fillOpacity ?? 1))}</a:gs>`).join('')}</a:gsLst><a:lin ang="${Math.round(angle * 60000) % 21600000}" scaled="1"/></a:gradFill>`;
  }
  if (sh.pattern && /^[A-Za-z][A-Za-z0-9]*$/.test(sh.pattern.preset)) fill = `<a:pattFill prst="${esc(sh.pattern.preset)}"><a:fgClr>${shapeColorXml(sh.pattern.fg ?? '#000000', sh.fillOpacity ?? 1)}</a:fgClr><a:bgClr>${shapeColorXml(sh.pattern.bg ?? '#ffffff', sh.fillOpacity ?? 1)}</a:bgClr></a:pattFill>`;
  const isLine = LINE_KINDS.has(sh.kind) && !sh.path && !sh.customGeometry;
  const dashes = new Set(['dot', 'dash', 'lgDash', 'dashDot', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot', 'sysDashDot', 'sysDashDotDot']);
  const dashName = sh.dash === 'longDash' ? 'lgDash' : sh.dash === 'dot' ? 'sysDot' : sh.dash;
  const dashXml = dashes.has(dashName) ? `<a:prstDash val="${dashName}"/>` : '';
  const ends = Object.entries(shapeLineEnds(sh)).filter(([, end]) => end.type !== 'none').map(([name, end]) => `<a:${name} type="${end.type}" w="${end.w}" len="${end.len}"/>`).join('');
  const cap = ['flat', 'rnd', 'sq'].includes(sh.lineCap) ? ` cap="${sh.lineCap}"` : '';
  const cmpd = ['sng', 'dbl', 'thickThin', 'thinThick', 'tri'].includes(sh.compound) ? ` cmpd="${sh.compound}"` : '';
  const join = ['round', 'bevel', 'miter'].includes(sh.lineJoin) ? `<a:${sh.lineJoin}/>` : '';
  const ln = sh.stroke ? `<a:ln w="${Math.round((sh.strokeWidth ?? 1) * EMU)}"${cap}${cmpd}><a:solidFill>${shapeColorXml(sh.stroke, sh.strokeOpacity ?? 1)}</a:solidFill>${dashXml}${join}${ends}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
  const effects = shapeEffectsXml(sh);
  if (isLine) {
    const locks = sh.lockAspect !== undefined ? `<a:cxnSpLocks noChangeAspect="${sh.lockAspect ? 1 : 0}"/>` : '';
    return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"${description}>${hyperlinkXml(sh)}</xdr:cNvPr><xdr:cNvCxnSpPr>${locks}</xdr:cNvCxnSpPr></xdr:nvCxnSpPr><xdr:spPr>${xfrm(sh)}<a:prstGeom prst="${sh.kind === 'line' ? 'straightConnector1' : sh.kind}"><a:avLst/></a:prstGeom>${ln}${effects}</xdr:spPr></xdr:cxnSp>`;
  }
  const algnOf = (a) => (a === 'center' ? 'ctr' : a === 'right' ? 'r' : a === 'justify' ? 'just' : 'l');
  const defaultAlign = sh.align ?? (sh.kind === 'textbox' ? 'left' : 'center');
  const algn = algnOf(defaultAlign);
  const faceXml = (font) => font ? `<a:latin typeface="${esc(font)}"/><a:ea typeface="${esc(font)}"/>` : '';
  const desktop = { ...sh, ...fontDesktopStyle(sh.font, sh) };
  const textPr = `lang="ko-KR" sz="${Math.round((sh.size ?? 11) * 100)}" b="${desktop.bold ? 1 : 0}" i="${sh.italic ? 1 : 0}" u="${sh.underline ? 'sng' : 'none'}" strike="${sh.strike ? 'sngStrike' : 'noStrike'}"`;
  const textChildren = `<a:solidFill>${shapeColorXml(sh.color ?? '#000000')}</a:solidFill>${faceXml(desktop.font)}`;
  const rPr = `<a:rPr ${textPr}>${textChildren}</a:rPr>`;
  // false/none도 명시하여 Excel의 목록·문단 기본값이 다시 적용되지 않게 한다.
  const runXml = (r) => {
    const source = { font: r.font ?? sh.font, bold: r.b ?? sh.bold }, face = { ...source, ...fontDesktopStyle(source.font, source) };
    const a = `lang="ko-KR" sz="${Math.round((r.sz ?? sh.size ?? 11) * 100)}" b="${face.bold ? 1 : 0}" i="${(r.i ?? sh.italic) ? 1 : 0}" u="${(r.u ?? sh.underline) ? 'sng' : 'none'}" strike="${(r.s ?? sh.strike) ? 'sngStrike' : 'noStrike'}"`;
    const pr = `<a:rPr ${a}><a:solidFill>${shapeColorXml(r.color ?? sh.color ?? '#000000')}</a:solidFill>${faceXml(face.font)}</a:rPr>`;
    const parts = String(r.t ?? '').split('\n');
    return parts.map((text, i) => `${i ? `<a:br>${pr}</a:br>` : ''}${text || parts.length === 1 ? `<a:r>${pr}<a:t xml:space="preserve">${esc(text)}</a:t></a:r>` : ''}`).join('');
  };
  const paras = sh.paras
    ? sh.paras.map((p) => `<a:p><a:pPr algn="${algnOf(p.align ?? defaultAlign)}"/>${p.runs.length ? p.runs.map(runXml).join('') : `<a:endParaRPr ${textPr.replace(/sz="[^"]*"/, `sz="${Math.round((p.sz ?? sh.size ?? 11) * 100)}"`)}>${textChildren}</a:endParaRPr>`}</a:p>`).join('')
    : String(sh.text ?? '').split('\n').map((line) => `<a:p><a:pPr algn="${algn}"/>${line ? `<a:r>${rPr}<a:t${/^\s|\s$/.test(line) ? ' xml:space="preserve"' : ''}>${esc(line)}</a:t></a:r>` : `<a:endParaRPr ${textPr}>${textChildren}</a:endParaRPr>`}</a:p>`).join('');
  const anchor = sh.valign ? { top: 't', middle: 'ctr', bottom: 'b' }[sh.valign] : sh.kind === 'textbox' ? 't' : 'ctr';
  const insets = sh.pad ? ['tIns', 'rIns', 'bIns', 'lIns'].map((k, i) => ` ${k}="${Math.round((sh.pad[i] ?? (i % 2 ? 9.6 : 4.8)) * EMU)}"`).join('') : '';
  const adjustment = sh.kind === 'roundRect' && Number.isSafeInteger(sh.adjustments?.adj) ? `<a:avLst><a:gd name="adj" fmla="val ${sh.adjustments.adj}"/></a:avLst>` : '<a:avLst/>';
  const geometry = sh.path ? customGeometryXml(sh.path) : sh.customGeometry ? storedCustomGeometryXml(sh.customGeometry) : `<a:prstGeom prst="${prstOf(sh.kind)}">${adjustment}</a:prstGeom>`;
  const locks = sh.lockAspect !== undefined ? `<a:spLocks noChangeAspect="${sh.lockAspect ? 1 : 0}"/>` : '';
  const textRot = [90, 270].includes(sh.textRot) ? ` rot="${(sh.textRot === 270 ? -90 : 90) * 60000}"` : '';
  const autofit = sh.textFit === 'shrink' ? '<a:normAutofit/>' : sh.textFit === 'none' ? '<a:noAutofit/>' : '';
  return `<xdr:sp macro="${sh.macro ? `[0]!${esc(sh.macro)}` : ''}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"${description}>${hyperlinkXml(sh)}</xdr:cNvPr><xdr:cNvSpPr${sh.kind === 'textbox' ? ' txBox="1"' : ''}>${locks}</xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr>${xfrm(sh)}${geometry}${fill}${ln}${effects}</xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="${sh.nowrap ? 'none' : 'square'}" rtlCol="0" anchor="${anchor}"${insets}${textRot}>${autofit}</a:bodyPr><a:lstStyle/>${paras}</xdr:txBody></xdr:sp>`;
}

/** 목록 원본: 범위 참조가 아니면 "a,b" 로 감싸기 */
function dvFormula(rule, f) {
  if (f === undefined || f === null || f === '') return '';
  const source=String(f).trim();
  let t=source.replace(/^=/,'');
  if(rule.type==='list'&&!source.startsWith('=')&&!t.startsWith('"')) {
    let reference=false;try{reference=['ref','range','sref','spill'].includes(parse(t).type);}catch{}
    if(!reference)t='"'+t.replace(/"/g,'""')+'"';
  }
  return t;
}

function validationXml(v) {
  const attrs = [`type="${v.type === 'any' ? 'none' : v.type}"`];
  if (v.errorStyle === 'warning' || v.errorStyle === 'info') attrs.push(`errorStyle="${v.errorStyle === 'info' ? 'information' : 'warning'}"`);
  if (!['list', 'custom', 'any'].includes(v.type) && v.op && v.op !== 'between') attrs.push(`operator="${v.op}"`);
  if (v.allowBlank !== false) attrs.push('allowBlank="1"');
  if (v.type === 'list' && v.showDropdown === false) attrs.push('showDropDown="1"');
  if(VALIDATION_IME_MODES.includes(v.imeMode))attrs.push('imeMode="'+v.imeMode+'"');
  if (v.showPrompt !== false) attrs.push('showInputMessage="1"');
  if (v.showError !== false) attrs.push('showErrorMessage="1"');
  for (const k of ['errorTitle', 'error', 'promptTitle', 'prompt']) if (v[k]) attrs.push(`${k}="${esc(v[k])}"`);
  attrs.push(`sqref="${rangeRef(v)}"`);
  const f1 = dvFormula(v, v.f1);
  const f2 = ['between', 'notBetween'].includes(v.op ?? 'between') && !['list', 'custom'].includes(v.type) ? dvFormula(v, v.f2) : '';
  return `<dataValidation ${attrs.join(' ')}>${f1 ? `<formula1>${esc(f1)}</formula1>` : ''}${f2 ? `<formula2>${esc(f2)}</formula2>` : ''}</dataValidation>`;
}

// ───────────────────────── 피벗 테이블 · 슬라이서 (쓰기) ─────────────────────────
const NS_X14 = 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main';
const NS_X15 = 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const NS_TB = 'https://tabula.local/spreadsheet/2026'; // WIXEL 전용 속성 (엑셀은 mc:Ignorable 로 무시)
const REL_MS = 'http://schemas.microsoft.com/office/2007/relationships';
const PIVOT_SUBTOTAL = { count: 'count', average: 'average', max: 'max', min: 'min', product: 'product', countNums: 'countNums', stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp' };

/** 필드의 고유 값 (피벗 결과와 같은 순서) → { keys, index: Map(key → 순번) } */
function fieldItems(data, f, order = null, retained = null) {
  const seen = new Map();
  for (const value of data.values(f)) { const k = keyOf(value); if (!seen.has(itemIdentity(k))) seen.set(itemIdentity(k), k); }
  let keys = sortKeys([...seen.values()]);
  if (order?.length) {
    // 수동 순서(파일에서 가져온 항목 순서)를 유지하고 새 항목은 뒤에
    const pos = new Map(order.map((t, i) => [itemIdentity(t), i]));
    keys = [...keys.filter((k) => pos.has(itemIdentity(itemText(k)))).sort((a, b) => pos.get(itemIdentity(itemText(a))) - pos.get(itemIdentity(itemText(b)))), ...keys.filter((k) => !pos.has(itemIdentity(itemText(k))))];
  }
  if (retained?.length) {
    // 원래 sharedItems 순서가 캐시 항목 번호이므로 새 값만 뒤에 추가합니다.
    const prior = retained.map(keyOf), identities = new Set(prior.map(itemIdentity));
    keys = prior.concat(keys.filter(k => !identities.has(itemIdentity(k))));
  }
  const index = new Map(keys.map((k, i) => [itemIdentity(k), i]));
  return { keys, index };
}

// ─── 피벗 그룹 (엑셀 fieldGroup) ───
const XL_GROUP_BY = { years: 'years', quarters: 'quarters', months: 'months', mdays: 'days', number: 'range' };
const XL_DATE_GROUP = new Set(['years', 'quarters', 'months', 'mdays']);
const serialIso = (v, date1904 = false) => { const d = dateParts(Math.floor(v), 1, date1904); return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`; };
/** 그룹 설정 → { keys: 엑셀 groupItems 순서의 항목 키(앱의 그룹 키와 같은 글자), rangePr } */
function excelGroupItems(spec, min, max, date1904 = false) {
  if (spec.by === 'number') {
    const size = Number(spec.size) || 10;
    const start = Number.isFinite(Number(spec.start)) ? Number(spec.start) : Math.floor(min);
    const keys = [`<${start}`];
    for (let lo = start; lo <= max && keys.length < 10000; lo += size) keys.push(`${lo}-${lo + size - 1}`);
    keys.push(`>${Math.max(max, start)}`);
    return { keys, rangePr: `<rangePr startNum="${start}" endNum="${Math.max(max, start)}" groupInterval="${size}"/>` };
  }
  const start = spec.start ?? Math.floor(min);
  const end = spec.end ?? Math.floor(max);
  const keys = [`<${serialIso(start, date1904)}`];
  if (spec.by === 'years') {
    const y1 = Number(serialIso(start, date1904).slice(0, 4));
    const y2 = Number(serialIso(end, date1904).slice(0, 4));
    for (let y = y1; y <= y2; y++) keys.push(y);
  } else if (spec.by === 'quarters') for (let q = 1; q <= 4; q++) keys.push(`${q}분기`);
  else if (spec.by === 'months') for (let m = 1; m <= 12; m++) keys.push(`${m}월`);
  else for (let m = 1; m <= 12; m++) for (let d = 1; d <= [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]; d++) keys.push(`${m}월${d}일`);
  keys.push(`>${serialIso(end, date1904)}`);
  return { keys, rangePr: `<rangePr groupBy="${XL_GROUP_BY[spec.by]}" startDate="${serialIso(start, date1904)}T00:00:00" endDate="${serialIso(end, date1904)}T00:00:00"/>` };
}
/** 날짜 필드의 sharedItems (항목은 쓰지 않음: 그룹 항목이 대신함) */
function dateSharedItemsXml(values, date1904 = false) {
  const stat = pivotValueStats(values);
  if (!stat.numbers || stat.numbers + stat.missing !== stat.count) return sharedItemsXml(values, null);
  return `<sharedItems${stat.missing ? '' : ' containsSemiMixedTypes="0"'} containsNonDate="0" containsDate="1" containsString="0"${stat.missing ? ' containsBlank="1"' : ''}/>`;
}

// 숫자로 계산하더라도 OOXML 날짜(<d>)와 일반 숫자(<n>)는 다른 캐시 형식입니다.
function pivotDateItems(field) {
  if (field?.dateOnly === true) return value => typeof value === 'number';
  if (!field?.sharedTypes?.includes('d')) return null;
  const values = new Set(); let onlyDates = true;
  for (let i = 0; i < field.shared.length; i++) if (typeof field.shared[i] === 'number') {
    if (field.sharedTypes[i] === 'd') values.add(field.shared[i]); else onlyDates = false;
  }
  return value => typeof value === 'number' && (onlyDates || values.has(value));
}
function pivotDateXml(value, date1904) {
  const d = dateParts(value, 1, date1904), pad = n => String(n).padStart(2, '0');
  const ms = (Math.round(value * 86400000) % 1000 + 1000) % 1000;
  return '<d v="' + d.y + '-' + pad(d.m) + '-' + pad(d.d) + 'T' + pad(d.hh) + ':' + pad(d.mm) + ':' + pad(d.ss) + (ms ? '.' + String(ms).padStart(3, '0') : '') + '"/>';
}
function sharedItemsXml(values, keys, retained = null, isDate = null, date1904 = false) {
  const allValues = retained?.length ? { *[Symbol.iterator]() { yield* retained; yield* values; } } : values;
  const numericValues = isDate ? { *[Symbol.iterator]() { for (const value of allValues) if (!isDate(value)) yield value; } } : allValues;
  const stat = pivotValueStats(numericValues);
  let hasDate = false;
  if (isDate) for (const value of allValues) if (isDate(value)) { hasDate = true; break; }
  const hasStr = stat.hasString || stat.emptyStrings > 0, hasBlank = stat.missing > 0, hasNum = stat.numbers > 0;
  const attrs = [];
  // Office의 빈 문자열(<s/>)과 실제 빈 항목(<m/>)은 서로 다른 형식입니다.
  if (!hasStr && !hasBlank) attrs.push('containsSemiMixedTypes="0"');
  if (!hasStr) attrs.push('containsString="0"');
  if (hasNum) {
    attrs.push('containsNumber="1"');
    if (stat.integers) attrs.push('containsInteger="1"');
    // MS-OI29500 sharedItems (k,l): 경계값은 실제 숫자 자식이 있을 때만 유효합니다.
    // https://learn.microsoft.com/openspecs/office_standards/ms-oi29500/e6bf5edd-aa4c-40b5-8f43-b3da6b664823
    if (keys?.some(k => typeof k === 'number')) attrs.push(`minValue="${stat.min}"`, `maxValue="${stat.max}"`);
  }
  if (hasBlank) attrs.push('containsBlank="1"');
  if (hasDate) {
    attrs.push('containsDate="1"');
    if (!hasStr && !hasNum) attrs.push('containsNonDate="0"');
  }
  if (hasStr && hasNum || hasDate && (hasStr || hasNum)) attrs.push('containsMixedTypes="1"');
  if (keys?.some(k => typeof k === 'string' && k.length > 255)) attrs.push('longText="1"');
  if (!keys) return `<sharedItems${attrs.length ? ` ${attrs.join(' ')}` : ''}/>`;
  const original = new Map();
  for (const value of allValues) { const k = keyOf(value), tag = itemIdentity(k); if (!original.has(tag)) original.set(tag, value); }
  const items = keys.map(k => {
    if (k === EMPTY) return '<m/>';
    if (k === EMPTY_TEXT) return '<s v=""/>';
    if (typeof k === 'number') return isDate?.(k) ? pivotDateXml(k, date1904) : `<n v="${k}"/>`;
    if (typeof k === 'boolean') return `<b v="${k ? 1 : 0}"/>`;
    const value = original.get(itemIdentity(k));
    return value && typeof value === 'object' && (value.error || value.code)
      ? `<e v="${esc(value.error ?? value.code)}"/>` : `<s v="${xesc(String(k))}"/>`;
  }).join('');
  return `<sharedItems${attrs.length ? ` ${attrs.join(' ')}` : ''} count="${keys.length}">${items}</sharedItems>`;
}

/** 피벗 원본을 구별하는 키 (같은 원본의 피벗은 캐시 하나를 함께 씀 — 슬라이서 연결에 필요) */
function pivotWorksheetKey(def) {
  const source = def.table ? `t:${String(def.table).toLowerCase()}` : `r:${String(def.source ?? '').toLowerCase()}!${def.range ? rangeRef(def.range) : ''}`;
  const reference = pivotSourceBindingCurrent(def) && importedPivotSourceReference(def);
  return reference?.external ? JSON.stringify([source, reference.external]) : source;
}
function pivotSourceKey(def) {
  const source = pivotWorksheetKey(def);
  // 독립 캐시는 분리하고, 명시적으로 연결한 슬라이서만 아래에서 합칩니다.
  return def.cacheItemsId ? JSON.stringify([source, def.cacheItemsId]) : source;
}
// MS Excel: auto is omitted; none=0, Excel 2007+ max=1048576.
// Other unsigned OOXML thresholds (including legacy 32500) remain exact.
function pivotMissingItemsLimit(def) {
  const value = def.missingItems;
  if (value === 'none') return 0;
  if (value === 'max') return 1048576;
  if (value === undefined || value === 'auto' || value === null || value === '') return undefined;
  const n = Number(value); return Number.isInteger(n) && n >= 0 && n <= 0xffffffff ? n : undefined;
}
function matchingPivotCacheData(wb, a, b) {
  const left = pivotSourceData(wb, a), right = pivotSourceData(wb, b);
  if (!left || !right) return false;
  if (left.cube === right.cube) return true;
  const headers = headerNames(left), other = headerNames(right);
  if (left.cube.n !== right.cube.n || headers.length !== other.length || headers.some((v,i) => v !== other[i])) return false;
  for (let c = 0; c < headers.length; c++) {
    const l = left.cube.col(c), r = right.cube.col(c);
    for (let i = 0; i < left.cube.n; i++) if (itemIdentity(keyOf(l.get(i))) !== itemIdentity(keyOf(r.get(i)))) return false;
  }
  return true;
}

/**
 * 같은 원본을 쓰는 피벗들의 공유 캐시 → { cacheXml, header, nBase, items: Map(필드 → {keys, index}), data, src }
 * extraFields: 슬라이서가 쓰는 필드 이름(소문자)
 */
function buildPivotCache(wb, defs, cacheId, extraFields, pool) {
  const src = pivotSourceData(wb, defs[0], { allowCacheMetadata: true });
  if (!src || src.cube.n < 1 && !src.cacheOnly) return null;
  const baseHeader = headerNames(src);
  const nBase = baseHeader.length;
  const limits = defs.map(pivotMissingItemsLimit).filter(n => n !== undefined);
  const missingItemsLimit = limits.length ? minOf(limits) : undefined;
  const validSnapshots = wb.snapshotData?.() ?? {}, retainedByName = new Map(), retainedMeta = new Map(), retainedIds = new Set();
  for (const def of defs) {
    if (!def.cacheItemsId || retainedIds.has(def.cacheItemsId)) continue;
    retainedIds.add(def.cacheItemsId);
    const original = wb.pivotCacheItems?.[def.cacheItemsId];
    // A loaded 'none' cache can still contain selected historical items until
    // refresh. Style-only saves retain it; a new none policy or refresh prunes it.
    const prune = missingItemsLimit === 0 && (original?.missingItemsLimit !== 0 || def.snapshotId && !validSnapshots[def.snapshotId]);
    for (const field of original?.fields ?? []) {
      const name = String(field.name).toLowerCase();
      if (!retainedMeta.has(name)) retainedMeta.set(name, field);
    }
    if (prune) continue;
    for (const field of original?.fields ?? []) {
      const name = String(field.name).toLowerCase(), prior = retainedByName.get(name);
      if (!prior) retainedByName.set(name, field);
      else if (prior.shared !== field.shared) {
        const merged = (prior.shared ?? []).slice(), seen = new Set(merged.map(v => itemIdentity(keyOf(v))));
        const types = [(prior.sharedTypes ?? '').padEnd(merged.length, ' ')];
        for (let i = 0; i < (field.shared?.length ?? 0); i++) {
          const value = field.shared[i], id = itemIdentity(keyOf(value));
          if (!seen.has(id)) { seen.add(id); merged.push(value); types.push(field.sharedTypes?.[i] ?? ' '); }
        }
        retainedByName.set(name, { ...prior, shared: merged, sharedTypes: types.join('') });
      }
    }
  }
  const retainedFor = f => retainedByName.get(String(baseHeader[f]).toLowerCase())?.shared;
  const fieldMetaFor = f => retainedByName.get(String(baseHeader[f]).toLowerCase()) ?? retainedMeta.get(String(baseHeader[f]).toLowerCase());
  const dateItems = new Map();
  for (let f = 0; f < nBase; f++) { const isDate = pivotDateItems(fieldMetaFor(f)); if (isDate) dateItems.set(f, isDate); }
  const calcs = [];
  for (const d of defs) for (const c of d.calcFields ?? []) {
    if (!baseHeader.some((h) => h.toLowerCase() === c.name.toLowerCase()) && !calcs.some((x) => x.name.toLowerCase() === c.name.toLowerCase())) calcs.push(c);
  }
  // 그룹화 (엑셀 형식으로 쓸 수 있는 것만): 제자리 그룹(필드 자체) · 파생 그룹 필드('월2' = '일'을 월로)
  const bx = (n) => baseHeader.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const grouped = new Map(); // 필드 이름(소문자) → { name, spec, base: 원본 필드 번호, derived }
  for (const d of defs) for (const [n, g] of Object.entries(d.groups ?? {})) {
    if (!g || (!XL_GROUP_BY[g.by] && g.by !== 'items') || grouped.has(n.toLowerCase())) continue;
    if (g.by === 'items' && (!g.base || bx(n) >= 0)) continue; // 선택 항목 그룹은 파생 필드('상품명2')만
    const derived = !!g.base && bx(n) < 0;
    const base = derived ? bx(g.base) : bx(n);
    if (base >= 0) grouped.set(n.toLowerCase(), { name: n, spec: g, base, derived });
  }
  const derivedList = [...grouped.values()].filter((g) => g.derived);
  const header = [...baseHeader, ...calcs.map((c) => c.name), ...derivedList.map((g) => g.name)];
  const nCalc = calcs.length;
  const fx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  // 현재 원본도 저장 캐시처럼 모든 행을 기록합니다. 열 접근만 사용하여 행 배열 복사는 만들지 않습니다.
  const data = pivotExportData(src.cube, true);
  // 같은 캐시의 충돌 옵션은 저장 제외 요청을 우선하고, 새로 고침 요청은 함께 적용합니다.
  const saveData = !defs.some(def => def.saveData === false);
  const refreshOnOpen = defs.some(def => def.refreshOnOpen === true);
  const listed = new Set();
  for (const def of defs) {
    const d = normalizeDef(def, header);
    [...d.rows, ...d.cols, ...d.pages, ...Object.keys(d.filters)].forEach((n) => listed.add(fx(n)));
  }
  header.forEach((h, i) => { if (extraFields?.has(h.toLowerCase()) || retainedFor(i)?.length) listed.add(i); });
  for (const g of grouped.values()) if (g.spec.by === 'items') listed.add(g.base); // discretePr 가 원본 항목 번호를 씀
  const items = new Map();
  for (const f of listed) if (f >= 0 && f < nBase && !grouped.has(header[f].toLowerCase())) items.set(f, fieldItems(data, f, null, retainedFor(f)));
  // 그룹 필드의 항목 = 엑셀 groupItems 전체 (피벗 필드의 x 가 이 목록을 가리킴)
  const groupXml = new Map();
  for (const g of grouped.values()) {
    const f = fx(g.name);
    if (g.spec.by === 'items') {
      // 선택 항목 그룹화: 원본 항목마다 그룹 번호(discretePr) + 그룹 항목(groupItems, 그룹에 안 든 항목은 자기 자신)
      const bk = items.get(g.base)?.keys ?? [];
      const keys = [];
      const pos = new Map();
      const disc = bk.map((k) => {
        const to = k === EMPTY ? EMPTY : groupKey(k, g.spec);
        const id = itemIdentity(to);
        if (!pos.has(id)) { pos.set(id, keys.length); keys.push(to); }
        return pos.get(id);
      });
      items.set(f, { keys, index: new Map(keys.map((k, i) => [itemIdentity(k), i])) });
      groupXml.set(f, `<fieldGroup base="${g.base}"><discretePr count="${disc.length}">${disc.map((i) => `<x v="${i}"/>`).join('')}</discretePr><groupItems count="${keys.length}">${keys.map((k) => (k === EMPTY ? '<m/>' : `<s v="${esc(itemText(k))}"/>`)).join('')}</groupItems></fieldGroup>`);
      continue;
    }
    const stat = pivotValueStats(data.values(g.base));
    const gi = excelGroupItems(g.spec, stat.numbers ? stat.min : 0, stat.numbers ? stat.max : 0, wb.date1904);
    items.set(f, { keys: gi.keys, index: new Map(gi.keys.map((k, i) => [itemIdentity(k), i])) });
    const par = !g.derived ? derivedList.find((x) => x.base === g.base && x.spec.by !== 'items') : null;
    groupXml.set(f, `<fieldGroup${par ? ` par="${fx(par.name)}"` : ''} base="${g.base}">${gi.rangePr}<groupItems count="${gi.keys.length}">${gi.keys.map((k) => `<s v="${esc(itemText(k))}"/>`).join('')}</groupItems></fieldGroup>`);
  }
  // 엑셀에 없는 함수(DIVIDE · ROWS)는 엑셀 수식으로 바꿔 쓰고, 원래 수식은 엑셀이 무시하는 tb:formula 에 (다시 열면 그대로)
  const calcAttr = (c) => {
    const xl = excelCalcFormula(c.formula);
    return `formula="${esc(xl)}"${xl !== c.formula ? ` tb:formula="${esc(c.formula)}"` : ''}`;
  };
  const cacheFields = header.map((h, f) => {
    if (f >= nBase + nCalc) return `<cacheField name="${esc(h)}" numFmtId="0" databaseField="0">${groupXml.get(f) ?? ''}</cacheField>`;
    if (f >= nBase) return `<cacheField name="${esc(h)}" numFmtId="0" ${calcAttr(calcs[f - nBase])} databaseField="0"/>`;
    const g = groupXml.get(f);
    const vals = data.values(f);
    // 날짜로 묶는 필드는 날짜 필드로 표시 (엑셀이 그룹을 다시 만들 수 있게)
    const shared = g && XL_DATE_GROUP.has(grouped.get(h.toLowerCase())?.spec.by) ? dateSharedItemsXml(vals, wb.date1904) : sharedItemsXml(vals, g ? null : items.get(f)?.keys ?? null, g ? null : retainedFor(f), dateItems.get(f), wb.date1904);
    return `<cacheField name="${esc(h)}" numFmtId="${g && XL_DATE_GROUP.has(grouped.get(h.toLowerCase())?.spec.by) ? 14 : pool.fmtId(fieldMetaFor(f)?.format ?? {})}">${shared}${g ?? ''}</cacheField>`;
  }).join('');
  const originalSource = src.cacheOnly ? src.sourceReference : null;
  const sourceXml = originalSource
    ? `<worksheetSource${originalSource.ref ? ` ref="${esc(originalSource.ref)}"` : ''}${originalSource.sheet ? ` sheet="${esc(originalSource.sheet)}"` : ''}${originalSource.name ? ` name="${esc(originalSource.name)}"` : ''}${originalSource.external ? ' r:id="rExternalSource"' : ''}/>`
    : src.table ? `<worksheetSource name="${esc(src.table)}"/>`
      : `<worksheetSource ref="${rangeRef(src.ref)}" sheet="${esc(wb.sheets[src.si].name)}"/>`;
  const cacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" xmlns:tb="${NS_TB}" mc:Ignorable="tb"${missingItemsLimit !== undefined ? ` missingItemsLimit="${missingItemsLimit}"` : ''} saveData="${saveData ? 1 : 0}" refreshOnLoad="${refreshOnOpen ? 1 : 0}"${saveData ? ' r:id="rCacheRecords"' : ''} createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${saveData ? data.length : 0}">`
    + `<cacheSource type="worksheet">${sourceXml}</cacheSource><cacheFields count="${header.length}">${cacheFields}</cacheFields>`
    + `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${cacheId}"/></ext></extLst></pivotCacheDefinition>`;
  const dateFields = new Set([...grouped.values()].filter(g => !g.derived && XL_DATE_GROUP.has(g.spec.by)).map(g => g.base));
  return { cacheXml, header, nBase, nCalc, items, data, src, cacheId, saveData, dateFields, dateItems, externalSource: originalSource?.external ?? null };
}

/** 저장 캐시를 행 단위로 인코딩합니다. 큰 단일 XML 문자열을 만들지 않습니다. */
function savedPivotRecords(cache, date1904, streaming = false) {
  const valueXml = (value, column) => {
    if (value == null) return '<m/>';
    if (typeof value === 'number') {
      if (cache.dateFields.has(column) || cache.dateItems?.get(column)?.(value)) return pivotDateXml(value, date1904);
      return `<n v="${value}"/>`;
    }
    if (typeof value === 'boolean') return `<b v="${value ? 1 : 0}"/>`;
    if (typeof value === 'object' && (value.error || value.code)) return `<e v="${esc(value.error ?? value.code)}"/>`;
    return `<s v="${xesc(String(value))}"/>`;
  };
  function* parts() {
    yield `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="${NS_MAIN}" count="${cache.data.length}">`;
    for (let r = 0; r < cache.data.length; r++) {
      yield '<r>'; for (let c = 0; c < cache.nBase; c++) yield valueXml(cache.data.value(r,c), c); yield '</r>';
    }
    yield '</pivotCacheRecords>';
  }
  if (streaming) return parts();
  const chunks = createXmlChunks(); for (const text of parts()) chunks.push(text);
  return chunks.finish();
}

/** 피벗 정의 + 공유 캐시 → 피벗 테이블 XML */
/** 피벗 조건부 서식 (선택 영역 · 값 필드 전체 · 행 필드 수준) → <conditionalFormats>. priority 는 시트 규칙 순서와 같음 */
function pivotCondXml(wb, si, def, header, values) {
  const conds = (wb.sheets[si].cond ?? []).filter((r) => r.r1 < EXCEL_MAX_ROWS);
  const out = [];
  conds.forEach((rl, i) => {
    const pv = rl.pivot;
    if (!pv || pv.name !== (def.name ?? '')) return;
    const vi = Math.max(0, values.findIndex((v) => valueName(v) === pv.value));
    const refs = [`<reference field="4294967294" count="1" selected="0"><x v="${vi}"/></reference>`];
    const rf = (pv.scope === 'field' || pv.scope === 'selection') && pv.rowField ? header.findIndex((h) => String(h).toLowerCase() === String(pv.rowField).toLowerCase()) : -1;
    // 엑셀과 같은 형식: 행 필드 참조는 항목 없이(count="0") — 그 필드의 모든 항목
    if (rf >= 0) refs.push(`<reference field="${rf}" count="0" selected="0"/>`);
    out.push(`<conditionalFormat scope="${pv.scope === 'selection' ? 'selection' : pv.scope === 'data' ? 'data' : 'field'}" priority="${i + 1}"><pivotAreas count="1"><pivotArea outline="0" collapsedLevelsAreSubtotals="1" fieldPosition="0"><references count="${refs.length}">${refs.join('')}</references></pivotArea></pivotAreas></conditionalFormat>`);
  });
  return out.length ? `<conditionalFormats count="${out.length}">${out.join('')}</conditionalFormats>` : '';
}

function pivotParts(wb, si, def, cache, name, pool) {
  const { src, header, nBase, data } = cache;
  const nCalcEnd = nBase + (cache.nCalc ?? 0);
  const isCalc = (f) => f >= nBase && f < nCalcEnd;
  const cacheId = cache.cacheId;
  const d = { ...normalizeDef(def, header), header };
  let cachedLayout = src.metadataOnly ? importedPivotSourceReference(def)?.pivotLayout : null;
  if (src.metadataOnly && (!pivotSourceReferenceCurrent(def) || !pivotCacheLayoutCurrent(def, cachedLayout))) throw new Error('피벗 테이블의 원본과 저장된 레코드가 없어 변경된 필드 배치를 저장할 수 없습니다. 원본 연결을 확인하세요.');
  if (!d.rows.length && !d.cols.length && !d.values.length && !d.pages.length) return null;
  const reportOnly = !d.rows.length && !d.cols.length && !d.values.length;
  const fx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const rowF = d.rows.map(fx);
  const colF = d.cols.map(fx);
  const pageF = d.pages.map(fx);
  const filters = Object.entries(d.filters).map(([n, allowed]) => [fx(n), new Set(allowed.map(itemIdentity))]).filter(([i]) => i >= 0);
  // 이 피벗의 항목 순서 (캐시 순서를 수동 순서대로 재배열) → x 는 이 목록의 위치
  const items = new Map();
  for (const [f, it] of cache.items) {
    const sorting = d.sort[header[f]], group = d.groups?.[header[f]];
    const captionSort = (sorting?.by === undefined || sorting?.by === null) && (!group || group.by === 'items');
    const savedOrder = d.order[header[f]];
    const orderSet = sorting && captionSort && savedOrder?.length ? new Set(savedOrder.map(itemIdentity)) : null;
    const ord = !sorting || captionSort && orderSet && it.keys.every(k => orderSet.has(itemIdentity(itemText(k)))) ? savedOrder : null;
    // Item order belongs to this pivot, while cacheIndex keeps the shared cache
    // untouched. Write the same caption order that a fresh read will restore.
    let keys = captionSort && (sorting || d.customListSort === false) ? sortKeys([...it.keys], { customList: d.customListSort !== false }) : it.keys;
    if (ord?.length) {
      const pos = new Map(ord.map((t, i) => [itemIdentity(t), i]));
      keys = [...keys.filter((k) => pos.has(itemIdentity(itemText(k)))).sort((a, b) => pos.get(itemIdentity(itemText(a))) - pos.get(itemIdentity(itemText(b)))), ...keys.filter((k) => !pos.has(itemIdentity(itemText(k))))];
    } else if (captionSort && sorting?.dir === 'desc') keys = [...keys.filter(k => k !== EMPTY).reverse(), ...keys.filter(k => k === EMPTY)];
    items.set(f, { keys, index: new Map(keys.map((k, i) => [itemIdentity(k), i])), cacheIndex: it.index });
  }

  const hiddenKey = (f, k) => {
    const set = filters.find(([i]) => i === f)?.[1];
    return set ? !set.has(itemIdentity(itemText(k))) : false;
  };
  // 원문 필드 항목 순번과 Σ/값 순서까지 일치할 때만 저장된 배치를 원자적으로 재사용합니다.
  const savedPresentation = def.importedPresentation;
  const snap = def.snapshotId && wb.pivotSnapshots?.get(def.snapshotId);
  if (!cachedLayout && def.captureFmt && !def.needsRender && def.refreshOnOpen !== true && snap
    && wb.pivotSnapshotCurrent({ ...snap }, def) && importedPivotPresentationCurrent(def, savedPresentation)) {
    const onRows0 = d.values.length > 1 && !!d.valuesOnRows, multi0 = d.values.length > 1 && !onRows0;
    const vp0 = multi0 ? Math.max(0, Math.min(colF.length, d.valuesPos ?? colF.length)) : colF.length;
    const rows0 = [...d.rows, ...(onRows0 ? [null] : [])];
    const cols0 = multi0 ? [...d.cols.slice(0, vp0), null, ...d.cols.slice(vp0)] : d.cols;
    const axisSame = JSON.stringify(savedPresentation.axes?.rows) === JSON.stringify(rows0) && JSON.stringify(savedPresentation.axes?.cols) === JSON.stringify(cols0);
    let orderSame = axisSame;
    for (const field of savedPresentation.itemOrder ?? []) {
      const f = fx(field.field), it = items.get(f);
      if (!it) { orderSame = false; break; }
      const onAxis = rowF.includes(f) || colF.includes(f);
      const subOn = d.subtotals === true || Array.isArray(d.subtotals) && d.subtotals.includes(header[f]);
      const tokens = function* () { for (const key of it.keys) yield 'v:' + itemIdentity(key); if (onAxis && subOn) yield 't:default'; };
      if (pivotItemOrderSignature(tokens()) !== field.signature) { orderSame = false; break; }
    }
    if (orderSame) cachedLayout = savedPresentation;
  }
  // 검증되지 않은 원문 배치는 location/rowItems/colItems 모두 기존 계산 경로로 처리합니다.
  const meta = cachedLayout && !src.metadataOnly
    ? { values:d.values, rowItems:[], colLeaves:[], pageRows:cachedLayout.body.r1-cachedLayout.top,
      bodyRows:cachedLayout.body.r2-cachedLayout.body.r1+1, width:cachedLayout.body.c2-cachedLayout.body.c1+1,
      headerRows:Number(cachedLayout.location.firstDataRow), labelCols:Number(cachedLayout.location.firstDataCol) }
    : computePivot(resolvePivot(src, def), d).meta;
  const values = meta.values;
  const V = values.length;
  const multiV = V > 1;
  const onRows = multiV && !!d.valuesOnRows; // Σ 값이 행 영역 (dataOnRows)
  const colMulti = multiV && !onRows;
  const xOf = (f, key) => items.get(f)?.index.get(itemIdentity(key)) ?? 0;
  const xTag = (v) => (v ? `<x v="${v}"/>` : '<x/>');

  // 행 항목
  const rowXml = [];
  const iv = (vi) => (vi > 0 ? ` i="${vi}"` : '');
  if (!rowF.length) {
    if (onRows) for (let vi = 0; vi < V; vi++) rowXml.push(`<i${iv(vi)}>${xTag(vi)}</i>`);
    else rowXml.push('<i/>');
  } else {
    let prev = [];
    for (const it of meta.rowItems) {
      if (it.kind === 'grand') { rowXml.push(`<i t="grand"${iv(it.vi)}><x/></i>`); continue; }
      if (it.kind === 'blank') {
        const n = it.node;
        rowXml.push(`<i t="blank"${n.depth ? ` r="${n.depth}"` : ''}>${xTag(xOf(rowF[n.depth], n.key))}</i>`);
        prev = [];
        continue;
      }
      if (it.kind === 'sub') {
        const n = it.node;
        rowXml.push(`<i t="default"${n.depth ? ` r="${n.depth}"` : ''}${iv(it.vi)}>${xTag(xOf(rowF[n.depth], n.key))}</i>`);
        continue;
      }
      // 값 행 (값이 행 영역): 첫 값은 항목 경로와 함께, 나머지는 값 번호만
      if (onRows && it.vi !== undefined && (it.valueRow || it.vi > 0)) {
        rowXml.push(`<i r="${it.valueRow ? it.node.depth + 1 : (it.chain ?? [it.node]).length}"${iv(it.vi)}>${xTag(it.vi)}</i>`);
        continue;
      }
      const chain = it.chain ?? [it.node];
      const xs = (it.chain ? chain : [it.node]).map((n) => xOf(rowF[n.depth], n.key));
      const start = it.chain ? 0 : it.node.depth;
      if (it.chain) {
        let r = 0;
        while (r < xs.length && r < prev.length && prev[r] === xs[r]) r++;
        if (r >= xs.length) r = xs.length - 1;
        rowXml.push(`<i${r ? ` r="${r}"` : ''}>${[...xs.slice(r), ...(onRows && it.vi === 0 ? [0] : [])].map(xTag).join('')}</i>`);
        prev = xs;
      } else rowXml.push(`<i${start ? ` r="${start}"` : ''}>${xTag(xs[0])}</i>`);
    }
  }
  // 열 항목
  const colXml = [];
  const vpos = colMulti ? Math.max(0, Math.min(colF.length, d.valuesPos ?? colF.length)) : colF.length;
  const colFieldsAll = colMulti ? [...colF.slice(0, vpos), -2, ...colF.slice(vpos)] : [...colF];
  if (!colFieldsAll.length) colXml.push('<i/>');
  else {
    let prev = [];
    for (const leaf of meta.colLeaves) {
      const ia = leaf.vi > 0 ? ` i="${leaf.vi}"` : '';
      if (leaf.kind === 'grand') { colXml.push(`<i t="grand"${ia}><x/></i>`); continue; }
      const chain = [];
      for (let n = leaf.node; n && n.depth >= 0; n = n.parent) chain.unshift(n);
      if (leaf.kind === 'sub') {
        const n = leaf.node;
        colXml.push(`<i t="default"${n.depth ? ` r="${n.depth}"` : ''}${ia}>${xTag(xOf(colF[n.depth], n.key))}</i>`);
        prev = [];
        continue;
      }
      const cx = chain.map((n) => xOf(colF[n.depth], n.key));
      const xs = colMulti ? [...cx.slice(0, vpos), leaf.vi, ...cx.slice(vpos)] : cx;
      let r = 0;
      while (r < xs.length - 1 && r < prev.length && prev[r] === xs[r]) r++;
      colXml.push(`<i${r ? ` r="${r}"` : ''}${ia}>${xs.slice(r).map(xTag).join('')}</i>`);
      prev = xs;
    }
  }

  const tabular = d.layout === 'tabular';
  const outline = d.layout === 'outline';
  const valueFieldIdx = new Set(values.map((v) => fx(v.field)).filter((i) => i >= 0));
  const fieldNumberFormats = new Map(Object.entries(def.fieldNumberFormats ?? {}).map(([field, format]) => [field.toLowerCase(), format]));
  const pivotFields = header.map((h, f) => {
    const format = fieldNumberFormats.get(h.toLowerCase());
    const builtin = format?.xlsxBuiltinId;
    const formatId = format ? Number.isInteger(builtin) && builtin >= 0 && builtin < 164 ? builtin : pool.fmtId(format) : null;
    const formatAttr = formatId === null ? '' : ` numFmtId="${formatId}"`;
    if (isCalc(f)) return `<pivotField${formatAttr}${valueFieldIdx.has(f) ? ' dataField="1"' : ''} dragToRow="0" dragToCol="0" dragToPage="0" showAll="0" defaultSubtotal="0"/>`;
    const attrs = formatId === null ? [] : [`numFmtId="${formatId}"`];
    if (d.fieldCaptions?.[h]) attrs.push(`name="${esc(d.fieldCaptions[h])}"`);
    if (rowF.includes(f)) attrs.push('axis="axisRow"');
    else if (colF.includes(f)) attrs.push('axis="axisCol"');
    else if (pageF.includes(f)) attrs.push('axis="axisPage"');
    if (pageF.includes(f) || typeof d.pageMulti[h] === 'boolean') attrs.push(`multipleItemSelectionAllowed="${pivotPageMulti(d, h) ? '1' : '0'}"`);
    if (valueFieldIdx.has(f)) attrs.push('dataField="1"');
    if (tabular || outline) attrs.push('compact="0"');
    if (tabular) attrs.push('outline="0"');
    attrs.push('showAll="0"');
    const onAxis = rowF.includes(f) || colF.includes(f);
    const srt = onAxis ? d.sort[h] : null;
    if (srt) attrs.push(`sortType="${srt.dir === 'desc' ? 'descending' : 'ascending'}"`);
    const subOn = d.subtotals === true || (Array.isArray(d.subtotals) && d.subtotals.includes(h));
    if (onAxis && !subOn) attrs.push('defaultSubtotal="0"');
    if (onAxis && !d.subtotalTop) attrs.push('subtotalTop="0"');
    if (rowF.includes(f) && (d.blankRows === true || (Array.isArray(d.blankRows) && d.blankRows.includes(h))) && f !== rowF[rowF.length - 1]) attrs.push('insertBlankRow="1"');
    const coll = onAxis ? new Set((d.collapsed?.[h] ?? []).map(itemIdentity)) : null;
    // 항목 레이블 반복 (x14 확장)
    const fill = rowF.includes(f) && d.repeatLabels ? `<extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}" xmlns:x14="${NS_X14}"><x14:pivotField fillDownLabels="1"/></ext></extLst>` : '';
    let scope = '';
    if (srt && srt.by !== undefined && srt.by !== null) {
      const byName = String(srt.by).toLowerCase();
      const vi = typeof srt.by === 'number' ? srt.by : values.findIndex(v => valueName(v).toLowerCase() === byName || v.field.toLowerCase() === byName);
      const refs = [`<reference field="4294967294" count="1" selected="0"><x v="${vi >= 0 && vi < V ? vi : 0}"/></reference>`];
      // autoSortScope item indexes address this pivot's field items, not shared
      // cache items: caption/manual sorts can give the two different orders.
      for (const pair of Array.isArray(srt.at) ? srt.at : []) {
        if (!Array.isArray(pair) || pair.length < 2) continue;
        const other = fx(pair[0]), keys = items.get(other)?.keys;
        const at = keys?.findIndex(k => itemIdentity(itemText(k)) === itemIdentity(String(pair[1]))) ?? -1;
        if (at >= 0) refs.push(`<reference field="${other}" count="1" selected="0"><x v="${at}"/></reference>`);
      }
      scope = `<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="${refs.length}">${refs.join('')}</references></pivotArea></autoSortScope>`;
    }
    const it = items.get(f);
    if (!it) return scope || fill ? `<pivotField ${attrs.join(' ')}>${scope}${fill}</pivotField>` : `<pivotField ${attrs.join(' ')}/>`;
    const caps = new Map(Object.entries(d.itemCaptions?.[h] ?? {}).map(([key, value]) => [itemIdentity(key), value]));
    const list = it.keys.map((k) => `<item${caps.has(itemIdentity(itemText(k))) ? ` n="${esc(caps.get(itemIdentity(itemText(k))))}"` : ''}${hiddenKey(f, k) ? ' h="1"' : ''}${coll?.has(itemIdentity(itemText(k))) ? ' sd="0"' : ''} x="${it.cacheIndex.get(itemIdentity(k))}"/>`).join('');
    const def0 = onAxis && !subOn ? '' : '<item t="default"/>';
    return `<pivotField ${attrs.join(' ')}><items count="${it.keys.length + (def0 ? 1 : 0)}">${list}${def0}</items>${scope}${fill}</pivotField>`;
  }).join('');

  const pageLayout = pivotPageLayout(d);
  const top = (def.top ?? 0) + (reportOnly ? pageLayout.height + 1 : meta.pageRows);
  const left = def.left ?? 0;
  let loc = { r1: top, c1: left, r2: top + (reportOnly ? 0 : meta.bodyRows - 1), c2: left + (reportOnly ? 0 : meta.width - 1) };
  const originalLoc = cachedLayout && refToRange(cachedLayout.location.ref ?? '');
  if (cachedLayout && !originalLoc) throw new Error('피벗 테이블의 원본과 저장된 레코드가 없어 배치 위치를 확인할 수 없습니다.');
  if (originalLoc) {
    const dr = (def.top ?? 0) - cachedLayout.top, dc = (def.left ?? 0) - cachedLayout.left;
    loc = { r1: originalLoc.r1 + dr, c1: originalLoc.c1 + dc, r2: originalLoc.r2 + dr, c2: originalLoc.c2 + dc };
  }
  const firstHeaderRow = reportOnly ? 0 : colF.length ? 1 : multiV ? (d.valuesHeadRow && !d.valuesOnRows ? 1 : 0) : 1;
  const pageXml = pageF.length ? `<pageFields count="${pageF.length}">${pageF.map((f) => {
    const allowed = filters.find(([i]) => i === f)?.[1];
    const one = !pivotPageMulti(d, header[f]) && allowed && allowed.size === 1 ? items.get(f)?.keys.findIndex((k) => allowed.has(itemIdentity(itemText(k)))) : -1;
    return `<pageField fld="${f}"${one >= 0 ? ` item="${one}"` : ''} hier="-1"/>`;
  }).join('')}</pageFields>` : '';
  const dataXml = values.map((v, vi) => {
    const f = Math.max(0, fx(v.field));
    const sub = PIVOT_SUBTOTAL[v.agg] && !isCalc(f) ? ` subtotal="${PIVOT_SUBTOTAL[v.agg]}"` : '';
    const as = SHOW_FROM_XLSX[v.showAs] ? v.showAs : null;
    const show = as && SHOW_BASE_ONLY.has(as) ? ` showDataAs="${as}"` : '';
    // 기준 필드 · 항목 (이전 · 다음은 특수 번호)
    const bf = as ? fx(v.baseField ?? d.rows[0] ?? d.cols[0] ?? '') : -1;
    let bi = 0;
    if (v.basePos === 'prev') bi = BASE_PREV;
    else if (v.basePos === 'next') bi = BASE_NEXT;
    else if (bf >= 0 && v.baseItem !== undefined && items.get(bf)) bi = Math.max(0, items.get(bf).keys.findIndex((k) => itemIdentity(itemText(k)) === itemIdentity(String(v.baseItem))));
    const ext = as && !SHOW_BASE_ONLY.has(as)
      ? `<extLst><ext uri="{E15A36E0-9728-4e99-A89B-3F7291B0FE68}" xmlns:x14="${NS_X14}"><x14:dataField pivotShowAs="${as}"/></ext></extLst>`
      : '';
    // 표시 형식: 파일에서 가져온 셀 서식 > 값 필드 서식 > 기본
    const fmtStyle = def.cellFmt?.[`data:${vi}`] ?? (v.numFmt ? (typeof v.numFmt === 'object' ? v.numFmt : { numFmt: v.numFmt }) : null);
    let numFmt = fmtStyle && pool ? pool.fmtId(fmtStyle) : 0;
    if (!numFmt) numFmt = showAsPercent(as) ? 10 : as === 'rankAscending' || as === 'rankDescending' ? 1 : isCalc(f) ? 0 : ['average', 'stdDev', 'stdDevp', 'var', 'varp'].includes(v.agg) ? 4 : 3;
    const head = `<dataField name="${esc(valueName(v))}" fld="${f}"${sub}${show} baseField="${Math.max(0, bf)}" baseItem="${bi}" numFmtId="${numFmt}"`;
    return ext ? `${head}>${ext}</dataField>` : `${head}/>`;
  }).join('');
  // 레이블 · 값 · 상위 10 필터
  const filterXml = [];
  let filterId = 1;
  for (const [field, flt] of Object.entries(d.fieldFilters ?? {})) {
    const f = fx(field);
    if (f < 0 || isCalc(f) || !flt) continue;
    const cap = (s) => s[0].toUpperCase() + s.slice(1);
    const two = flt.op === 'between' || flt.op === 'notBetween';
    if (flt.type === 'top') {
      const mode = flt.mode ?? 'count';
      filterXml.push(`<filter fld="${f}" type="${mode}" evalOrder="-1" id="${filterId++}" iMeasureFld="${Number(flt.by) || 0}"><autoFilter ref="A1"><filterColumn colId="0"><top10${flt.top === false ? ' top="0"' : ''}${mode === 'percent' ? ' percent="1"' : ''} val="${Number(flt.n) || 10}" filterVal="${Number(flt.n) || 10}"/></filterColumn></autoFilter></filter>`);
    } else if (flt.type === 'date') {
      const dyn = !/^date/.test(flt.op);
      const pair = flt.op === 'dateBetween' || flt.op === 'dateNotBetween';
      const cmpOp = { dateEqual: 'equal', dateNotEqual: 'notEqual', dateOlderThan: 'lessThan', dateOlderThanOrEqual: 'lessThanOrEqual', dateNewerThan: 'greaterThan', dateNewerThanOrEqual: 'greaterThanOrEqual' }[flt.op];
      const inner = dyn ? `<dynamicFilter type="${flt.op}"/>`
        : pair ? `<customFilters${flt.op === 'dateBetween' ? ' and="1"' : ''}><customFilter operator="${flt.op === 'dateBetween' ? 'greaterThanOrEqual' : 'lessThan'}" val="${Number(flt.v1)}"/><customFilter operator="${flt.op === 'dateBetween' ? 'lessThanOrEqual' : 'greaterThan'}" val="${Number(flt.v2)}"/></customFilters>`
          : `<customFilters><customFilter${cmpOp && cmpOp !== 'equal' ? ` operator="${cmpOp}"` : ''} val="${Number(flt.v1)}"/></customFilters>`;
      const sv = (v) => (v === undefined ? '' : serialIso(Number(v), wb.date1904).slice(0, 10));
      filterXml.push(`<filter fld="${f}" type="${flt.op}" evalOrder="-1" id="${filterId++}"${dyn ? '' : ` stringValue1="${sv(flt.v1)}"${pair ? ` stringValue2="${sv(flt.v2)}"` : ''}`}><autoFilter ref="A1"><filterColumn colId="0">${inner}</filterColumn></autoFilter></filter>`);
    } else if (flt.type === 'label' || flt.type === 'value') {
      const type = `${flt.type === 'label' ? 'caption' : 'value'}${cap(flt.op)}`;
      const OPS = { equal: 'equal', notEqual: 'notEqual', greaterThan: 'greaterThan', greaterThanOrEqual: 'greaterThanOrEqual', lessThan: 'lessThan', lessThanOrEqual: 'lessThanOrEqual' };
      const wild = { beginsWith: (x) => `${x}*`, notBeginsWith: (x) => `${x}*`, endsWith: (x) => `*${x}`, notEndsWith: (x) => `*${x}`, contains: (x) => `*${x}*`, notContains: (x) => `*${x}*` };
      let custom;
      if (two) custom = `<customFilters${flt.op === 'between' ? ' and="1"' : ''}><customFilter operator="${flt.op === 'between' ? 'greaterThanOrEqual' : 'lessThan'}" val="${esc(String(flt.v1))}"/><customFilter operator="${flt.op === 'between' ? 'lessThanOrEqual' : 'greaterThan'}" val="${esc(String(flt.v2))}"/></customFilters>`;
      else if (wild[flt.op]) custom = `<customFilters><customFilter${/^not/.test(flt.op) ? ' operator="notEqual"' : ''} val="${esc(wild[flt.op](String(flt.v1)))}"/></customFilters>`;
      else custom = `<customFilters><customFilter${OPS[flt.op] && OPS[flt.op] !== 'equal' ? ` operator="${OPS[flt.op]}"` : ''} val="${esc(String(flt.v1))}"/></customFilters>`;
      filterXml.push(`<filter fld="${f}" type="${type}" evalOrder="-1" id="${filterId++}"${flt.type === 'value' ? ` iMeasureFld="${Number(flt.by) || 0}"` : ''} stringValue1="${esc(String(flt.v1 ?? ''))}"${two ? ` stringValue2="${esc(String(flt.v2 ?? ''))}"` : ''}><autoFilter ref="A1"><filterColumn colId="0">${custom}</filterColumn></autoFilter></filter>`);
    }
  }
  const so = d.styleOpts ?? {};
  const styleName = d.style === 'None' ? '' : d.style ?? 'PivotStyleLight16';
  if (styleName && Array.isArray(def.styleElements)) pool.objectTableStyle({name:styleName,table:false,pivot:true,elements:def.styleElements});
  else if (styleName && isModernStyle(styleName)) pool.presetTableStyle(styleName, true);
  else if (styleName && def.styleDef && !/^PivotStyle(Light|Medium|Dark)\d+$/i.test(styleName)) pool.pivotStyle(styleName, def.styleDef);
  const tableAttrs = [
    `name="${esc(name)}"`, `cacheId="${cacheId}"`, 'applyNumberFormats="0"', 'applyBorderFormats="0"', 'applyFontFormats="0"', 'applyPatternFormats="0"',
    'applyAlignmentFormats="0"', 'applyWidthHeightFormats="1"', `dataCaption="${esc(d.dataCaption ?? '값')}"`, ...(onRows ? ['dataOnRows="1"'] : []), 'updatedVersion="6"', 'minRefreshableVersion="3"', `useAutoFormatting="${def.autofit === false ? 0 : 1}"`,
    ...(def.mergeLabels ? ['mergeItem="1"'] : []), ...(def.showHeaders === false ? ['showHeaders="0"'] : []), ...(def.preserveFormat === false ? ['preserveFormatting="0"'] : []), ...(def.enableDrill === false ? ['enableDrill="0"'] : []),
    ...(d.rowCaption !== null && d.rowCaption !== undefined ? [`rowHeaderCaption="${esc(d.rowCaption)}"`] : []),
    ...(d.grandCaption !== null && d.grandCaption !== undefined ? [`grandTotalCaption="${esc(d.grandCaption)}"`] : []),
    `showError="${d.errorShow ? 1 : 0}"`, ...(def.errorCaption !== null && def.errorCaption !== undefined ? [`errorCaption="${esc(def.errorCaption)}"`] : []), ...(d.colCaption !== null && d.colCaption !== undefined ? [`colHeaderCaption="${esc(d.colCaption)}"`] : []),
    ...(d.grandRows ? [] : ['rowGrandTotals="0"']), ...(d.grandCols ? [] : ['colGrandTotals="0"']),
    ...(d.customListSort === false ? ['customListSort="0"'] : []),
    ...(d.missingCaption ? [`missingCaption="${esc(d.missingCaption)}"`] : []), ...(d.showExpand ? [] : ['showDrill="0"']),
    'itemPrintTitles="1"', 'createdVersion="6"', 'indent="0"', ...(tabular || outline ? ['compact="0"', 'compactData="0"'] : []),
    `outline="${tabular ? 0 : 1}"`, `outlineData="${tabular ? 0 : 1}"`, ...(d.classic ? ['gridDropZones="1"'] : []), `multipleFieldFilters="${def.multiFilters ? 1 : 0}"`,
    ...(pageF.length && d.pageOrder === 'over' ? ['pageOverThenDown="1"'] : []), ...(pageF.length && d.pageWrap ? [`pageWrap="${d.pageWrap}"`] : []),
  ];
  const locationXml = cachedLayout ? `<location${Object.entries({ ...cachedLayout.location, ref: rangeRef(loc) }).map(([key, value]) => ` ${key}="${esc(value)}"`).join('')}/>`
    : `<location ref="${rangeRef(loc)}" firstHeaderRow="${firstHeaderRow}" firstDataRow="${meta.headerRows}" firstDataCol="${reportOnly ? 0 : meta.labelCols}"${pageF.length ? ` rowPageCount="${pageLayout.height}" colPageCount="${(pageLayout.width + 1) / 3}"` : ''}/>`;
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="${NS_MAIN}" ${tableAttrs.join(' ')}>`
    + locationXml
    + `<pivotFields count="${header.length}">${pivotFields}</pivotFields>`
    + (rowF.length || onRows ? `<rowFields count="${rowF.length + (onRows ? 1 : 0)}">${[...rowF, ...(onRows ? [-2] : [])].map((f) => `<field x="${f}"/>`).join('')}</rowFields>` : '')
    + (cachedLayout ? cachedLayout.rowItems : reportOnly || !rowXml.length ? '' : `<rowItems count="${rowXml.length}">${rowXml.join('')}</rowItems>`)
    + (colFieldsAll.length ? `<colFields count="${colFieldsAll.length}">${colFieldsAll.map((f) => `<field x="${f}"/>`).join('')}</colFields>` : '')
    + (cachedLayout ? cachedLayout.colItems : reportOnly || !colXml.length ? '' : `<colItems count="${colXml.length}">${colXml.join('')}</colItems>`)
    + pageXml
    + (values.length ? `<dataFields count="${values.length}">${dataXml}</dataFields>` : '')
    + pivotCondXml(wb, si, def, header, values)
    + `<pivotTableStyleInfo${styleName ? ` name="${esc(styleName)}"` : ''} showRowHeaders="${so.rowHeaders === false ? 0 : 1}" showColHeaders="${so.colHeaders === false ? 0 : 1}" showRowStripes="${so.bandRows ? 1 : 0}" showColStripes="${so.bandCols ? 1 : 0}" showLastColumn="1"/>`
    + (filterXml.length ? `<filters count="${filterXml.length}">${filterXml.join('')}</filters>` : '')
    + `<extLst><ext uri="{962EF5D1-5CA2-4c93-8EF4-DBF5C05439D2}" xmlns:x14="${NS_X14}"><x14:pivotTableDefinition hideValuesRow="${d.showValuesRow ? 0 : 1}"/></ext>`
    // 계산 항목: WIXEL 확장 (엑셀은 모르는 ext 를 무시하고 원래 항목만 보여 줌)
    + (def.calcItems && Object.keys(def.calcItems).length ? `<ext uri="{6B1E4C27-3D5A-4F80-9C12-57495845434D}" xmlns:wx="https://wixel.app/x"><wx:calcItems json="${esc(JSON.stringify(def.calcItems))}"/></ext>` : '')
    + '</extLst>'
    + '</pivotTableDefinition>';

  // 슬라이서 캐시용: 필드 이름 → 항목 선택 상태 (x 는 캐시의 항목 번호)
  const slicerItems = (fieldName) => {
    const f = header.findIndex((h) => h.toLowerCase() === String(fieldName).toLowerCase());
    if (f < 0 || !cache.items.has(f)) return null;
    const others = filters.filter(([i]) => i !== f);
    const present = new Set();
    for (const value of data.values(f, r => others.every(([i,set]) => set.has(itemIdentity(itemText(keyOf(data.value(r,i)))))))) { const k=keyOf(value);present.add(itemIdentity(k)); }
    const keys = cache.items.get(f).keys;
    return {
      field: header[f],
      xml: keys.map((k, i) => `<i x="${i}"${hiddenKey(f, k) ? '' : ' s="1"'}${present.has(itemIdentity(k)) ? '' : ' nd="1"'}/>`).join(''),
      count: keys.length,
    };
  };
  return { tableXml, slicerItems };
}

/** 슬라이서가 연결된 피벗 [{ si, name }] (name 이 null 이면 그 시트의 첫 피벗) */
function slicerPivotRefs(wb, si, sl) {
  if (sl.source?.kind !== 'pivot') return [];
  if (sl.source.pivots?.length) return sl.source.pivots.map((p) => ({ si: p.sheet ? wb.sheetIndexByName(p.sheet) : si, name: p.name ?? null })).filter((p) => p.si >= 0);
  const s = sl.source.self || !sl.source.sheet ? si : wb.sheetIndexByName(sl.source.sheet);
  return s >= 0 ? [{ si: s, name: null }] : [];
}

/** 정의된 이름으로 쓸 수 있는 캐시 이름 */
function cacheNameFor(base, used) {
  let b = `슬라이서_${String(base).replace(/[^\wÀ-￿]/g, '_')}`;
  if (/^\d/.test(b)) b = `_${b}`;
  let n = b;
  for (let i = 1; used.has(n.toLowerCase()); i++) n = `${b}${i}`;
  used.add(n.toLowerCase());
  return n;
}

function slicerContentXml(sl, name, id, kind, grouped = false) {
  const EMUv = (px) => Math.round(px * EMU);
  const choice = kind === 'table'
    ? `<mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">`
    : `<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">`;
  const frame = `<xdr:graphicFrame macro="${sl.macro ? `[0]!${esc(sl.macro)}` : ''}"><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"${sl.alt !== undefined ? ` descr="${esc(sl.alt)}"` : ''}${sl.hidden ? ' hidden="1"' : ''}/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="${grouped ? EMUv(sl.x) : 0}" y="${grouped ? EMUv(sl.y) : 0}"/><a:ext cx="${grouped ? EMUv(sl.w) : 0}" cy="${grouped ? EMUv(sl.h) : 0}"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${esc(name)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const fallback = `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${EMUv(sl.x)}" y="${EMUv(sl.y)}"/><a:ext cx="${EMUv(sl.w)}" cy="${EMUv(sl.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="1100"/><a:t>이 도형은 ${kind === 'table' ? '표' : '피벗 테이블'} 슬라이서를 나타냅니다. 슬라이서는 Excel 2010 이상에서 지원됩니다.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>`;
  return `<mc:AlternateContent xmlns:mc="${NS_MC}">${choice}${frame}</mc:Choice>${fallback}</mc:AlternateContent>`;
}

function slicerAnchorXml(sl, name, id, anchorAt, kind) {
  return `<xdr:twoCellAnchor editAs="${sl.placement ?? 'oneCell'}"><xdr:from>${anchorAt(sl.x, sl.y)}</xdr:from><xdr:to>${anchorAt(sl.x + sl.w, sl.y + sl.h)}</xdr:to>${slicerContentXml(sl, name, id, kind)}<xdr:clientData${sl.noPrint ? ' fPrintsWithSheet="0"' : ''}${sl.locked === false ? ' fLocksWithSheet="0"' : ''}/></xdr:twoCellAnchor>`;
}

function slicerGroupAnchorXml(members, id, nextId, anchorAt) {
  let x = Infinity, y = Infinity, right = -Infinity, bottom = -Infinity;
  for (const [, o] of members) { const sl = o.sl; x = Math.min(x, sl.x); y = Math.min(y, sl.y); right = Math.max(right, sl.x + sl.w); bottom = Math.max(bottom, sl.y + sl.h); }
  const w = right - x, h = bottom - y, emu = px => Math.round(px * EMU);
  const children = members.map(([kind, o]) => slicerContentXml({ ...o.sl, x: o.sl.x - x, y: o.sl.y - y }, o.name, nextId(), kind === 'slicerTable' ? 'table' : 'pivot', true)).join('');
  const first = members[0][1].sl;
  return `<xdr:twoCellAnchor editAs="${first.placement ?? 'oneCell'}"><xdr:from>${anchorAt(x, y)}</xdr:from><xdr:to>${anchorAt(right, bottom)}</xdr:to><xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="${id}" name="슬라이서 그룹 ${id}"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="${emu(x)}" y="${emu(y)}"/><a:ext cx="${emu(w)}" cy="${emu(h)}"/><a:chOff x="0" y="0"/><a:chExt cx="${emu(w)}" cy="${emu(h)}"/></a:xfrm></xdr:grpSpPr>${children}</xdr:grpSp><xdr:clientData${members.every(([,o]) => o.sl.noPrint) ? ' fPrintsWithSheet="0"' : ''}${members.every(([,o]) => o.sl.locked === false) ? ' fLocksWithSheet="0"' : ''}/></xdr:twoCellAnchor>`;
}

/** .xlsx 한도 밖의 값/수식/명시적 셀 서식·메타데이터 수. 순수 빈칸과 외부 참조 캐시는 제외한다. */
export function xlsxOverflow(wb) {
  let n = 0;
  const own = wb.ownSheetCount ? wb.ownSheetCount() : wb.sheets.length;
  for (let si = 0; si < own; si++) {
    const sheet = wb.sheets[si], cells = sheet.cells;
    for (const [r,c,cell,count] of storedCellEntries(cells)) {
      if (!((cell.raw != null && cell.raw !== '') || cell.formula || cell.image?.src || cell.comment || cell.link || cell.phonetic || cell.style && Object.keys(cell.style).length)) continue;
      n += c >= MAX_COLS ? count : Math.max(0,r+count-Math.max(r,EXCEL_MAX_ROWS));
    }
    const blocks = sheet.blocks ?? [];
    for (let bi = 0; bi < blocks.length; bi++) {
      const b = blocks[bi], end = b.r0 + b.n;
      if (end <= EXCEL_MAX_ROWS && b.c0 + b.cols.length <= MAX_COLS) continue;
      for (let j = 0; j < b.cols.length; j++) {
        const c = b.c0 + j, start = c >= MAX_COLS ? b.r0 : Math.max(b.r0, EXCEL_MAX_ROWS);
        const col = b.cols[j];
        const formatted = !!col.fmt && Object.keys(col.fmt).length > 0;
        if (start >= end || !col.num && !col.str && !formatted) continue;
        // Workbook의 첫 블록 우선 규칙: 앞 블록의 빈칸도 뒤 블록의 값을 가린다.
        // 좌표 Set 대신 겹치는 행 구간만 보관하여 큰 블록의 추가 메모리를 제한한다.
        const covered = [];
        for (let k = 0; k < bi; k++) {
          const prev = blocks[k];
          if (c >= prev.c0 && c < prev.c0 + prev.cols.length && prev.r0 < end && prev.r0 + prev.n > start) covered.push([prev.r0, prev.r0 + prev.n]);
        }
        covered.sort((a, b) => a[0] - b[0]);
        const overrides = cells.col ? cells.col(c) : null;
        let cover = 0;
        for (let r = start; r < end; r++) {
          while (cover < covered.length && covered[cover][1] <= r) cover++;
          if (cover < covered.length && covered[cover][0] <= r) { r = covered[cover][1] - 1; continue; }
          if (overrides ? overrides.has(r) : !cells.col && cells.has(`${r},${c}`)) continue;
          const value = blockValue(b, r, c);
          if (formatted || value !== null && value !== undefined && value !== '') n++;
        }
      }
    }
  }
  return n;
}

/** Workbook → xlsx 바이트 (매크로가 있으면 .xlsm 형식) */
/** 통합 문서 → .xlsx 바이트 (한 번에) */
export function writeXlsx(wb, opts) {
  const it = writeXlsxSteps(wb, opts);
  for (;;) {
    const s = it.next();
    if (s.done) return zip(s.value, { consume: true });
  }
}

/** 큰 문서용: 중간중간 브라우저에 제어를 돌려주고, 내장 압축으로 파일 크기도 줄임. onProgress({p, msg}) */
export async function writeXlsxAsync(wb, opts, onProgress) {
  return writeXlsxArchive(wb, opts, onProgress, false);
}

/** 다운로드용: 최종 ZIP 전체를 다시 복사하지 않고 Blob으로 반환합니다. */
export async function writeXlsxBlobAsync(wb, opts, onProgress) {
  return writeXlsxStreamArchive(wb, opts, onProgress);
}

/** File System Access 저장용: Excel 호환 ZIP 바이트를 디스크 sink에 순차 기록합니다. */
export async function writeXlsxToSink(wb, opts, sink, onProgress) {
  return writeXlsxStreamArchive(wb, opts, onProgress, createZipStreamWriter(sink, { zip64: 'auto', signal: opts?.signal }));
}

// 다운로드는 Blob, 파일 선택기로 얻은 대상은 순차 ZIP sink를 사용합니다.
async function writeXlsxStreamArchive(wb, opts, onProgress, diskWriter = null) {
  const writer = diskWriter ?? createZipAsyncWriter();
  const check = () => { opts?.signal?.throwIfAborted(); opts?.assertCurrent?.(); };
  check();
  let last = performance.now();
  const rowProgress = value => { check(); if (performance.now() - last > 40) { onProgress?.(value); last = performance.now(); } };
  const it = writeXlsxSteps(wb, opts, true, rowProgress);
  try {
    for (;;) {
      check();
      const step = it.next();
      if (step.done) {
        const names = Object.keys(step.value);
        for (let i = 0; i < names.length; i++) {
          const name = names[i];
          check();
          await writer.add(name, step.value[name]); delete step.value[name];
          check();
          onProgress?.({ p: 0.85 + 0.15 * (i + 1) / names.length, msg: '파일 마무리 중' });
        }
        check();
        const result = await (diskWriter ? writer.finish() : writer.finish({ blob: true, first: '[Content_Types].xml' }));
        check();
        return result;
      }
      if (step.value.entry) await writer.add(step.value.entry[0], step.value.entry[1]);
      else if (performance.now() - last > 40) {
        onProgress?.(step.value); await new Promise(resolve => setTimeout(resolve, 0)); last = performance.now();
      }
    }
  } finally { it.return?.(); }
}

async function writeXlsxArchive(wb, opts, onProgress, blob) {
  const it = writeXlsxSteps(wb, opts);
  let last = performance.now();
  for (;;) {
    const s = it.next();
    if (s.done) return zipAsync(s.value, (p) => onProgress?.({ p: 0.85 + 0.15 * p, msg: '압축 중' }), { consume: true, blob });
    if (performance.now() - last > 40) {
      onProgress?.(s.value);
      await new Promise((res) => setTimeout(res, 0));
      last = performance.now();
    }
  }
}

// 통합 문서 본문 형식: 일반 · 매크로 사용 · 서식 파일
const MAIN_TYPES = {
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml',
  xlsm: 'application/vnd.ms-excel.sheet.macroEnabled.main+xml',
  xltx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml',
  xltm: 'application/vnd.ms-excel.template.macroEnabled.main+xml',
};

function* writeXlsxSteps(wb, { activeSheet = 0, fileName = 'Book1.xlsx', kind = null } = {}, streaming = false, onRows) {
  const files = {};
  const pool = new StylePool(wb.defaultFont ?? WRITE_FONT, wb.baseStyle);
  const gridIndex = color => {
    const rgb = [1, 3, 5].map(i => parseInt(color.slice(i, i + 2), 16)); let best = 0, distance = Infinity;
    for (let i = 0; i < 64; i++) { const d = rgb.reduce((sum, n, j) => sum + (n - parseInt(INDEXED[i].slice(j * 2, j * 2 + 2), 16)) ** 2, 0); if (d < distance) { best = i; distance = d; } }
    return best;
  };
  pool.namedStyles(wb.cellStyles);
  pool.objectStyles(wb.objectStyles);
  const wmdw = digitWidth(pool.baseFont); // 파일의 열 너비 = 픽셀 ÷ 기본 글꼴 숫자 너비
  const strings = [];
  const stringIndex = new Map();
  const sst = (s) => {
    if (!stringIndex.has(s)) { stringIndex.set(s, strings.length); strings.push(s); }
    return stringIndex.get(s);
  };
  const contentOverrides = [];
  let chartNo = 0;
  let drawingNo = 0;
  let commentNo = 0;
  let tableNo = 0;
  let mediaNo = 0;
  const drawingMediaTargets = new Map();
  const mediaExts = new Set();
  // 셀에 배치한 그림 → 리치 값 (같은 그림·설명은 하나로)
  const richList = [];
  const richMedia = new Map(); // src → rel 순번
  const richKey = new Map(); // src + alt → vm (1부터)
  const richImage = (image) => {
    const k = `${image.alt ?? ''}\u0000${image.src}`;
    if (richKey.has(k)) return richKey.get(k);
    const m = /^data:image\/([a-z+]+);base64,(.*)$/i.exec(image.src);
    if (!m) return 0;
    const ext = { jpeg: 'jpeg', jpg: 'jpeg', png: 'png', gif: 'gif', 'svg+xml': 'svg', webp: 'webp', bmp: 'bmp' }[m[1].toLowerCase()];
    if (!ext || !MIME[ext]) return 0;
    if (!richMedia.has(image.src)) {
      mediaNo++;
      mediaExts.add(ext);
      files[`xl/media/image${mediaNo}.${ext}`] = fromBase64(m[2]);
      richMedia.set(image.src, { i: richMedia.size, path: `../media/image${mediaNo}.${ext}` });
    }
    richList.push({ rel: richMedia.get(image.src).i, alt: image.alt ?? '' });
    richKey.set(k, richList.length);
    return richList.length;
  };
  const definedNames = [];
  const vba = wb.vba?.bin && kind !== 'xlsx' && kind !== 'xltx' ? wb.vba : null; // 매크로 없는 형식으로 저장하면 VBA 제외
  const nameSet = new Set((wb.names ?? []).map((n) => n.name.toUpperCase()));
  fileNameCheck = (n) => nameSet.has(String(n).toUpperCase());
  let dynamicCells = 0;
  // 표 번호(표 슬라이서가 참조)와 피벗(슬라이서 필드) 미리 계산
  const tableIds = new Map();
  let tableSeq = 0;
  wb.sheets.forEach((s) => (s.tables ?? []).forEach((t) => { if (t.r1 < EXCEL_MAX_ROWS) tableIds.set(t.name.toLowerCase(), ++tableSeq); }));
  // 피벗: 같은 원본끼리 캐시 하나를 공유 (슬라이서가 여러 피벗을 함께 거르려면 필요)
  const allDefs = [];
  wb.sheets.forEach((s, i) => [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean).forEach((def, j) => allDefs.push({ si: i, j, def })));
  const exportSnapshots = wb.snapshotData?.() ?? {};
  const originalCacheKeys = new Map(allDefs.map(e => [e.def, pivotSourceKey(e.def) + (exportSnapshots[e.def.snapshotId] ? `!saved:${e.def.snapshotId}` : '')]));
  const parents = new Map([...originalCacheKeys.values()].map(k => [k,k]));
  const rootKey = key => { let root = key; while (parents.get(root) !== root) root = parents.get(root); while (key !== root) { const next = parents.get(key); parents.set(key,root); key = next; } return root; };
  const entryFor = ref => allDefs.find(e => e.si === ref.si && (ref.name === null ? e.j === 0 : String(e.def.name ?? '') === ref.name));
  // A user-created report connection is explicit permission to share a cache.
  // Keep independent caches separate; never replace different saved values.
  wb.sheets.forEach((s,i) => (s.slicers ?? []).forEach(sl => {
    const linked = slicerPivotRefs(wb,i,sl).map(entryFor).filter(Boolean), first = linked[0];
    for (const other of linked.slice(1)) {
      const a = rootKey(originalCacheKeys.get(first.def)), b = rootKey(originalCacheKeys.get(other.def));
      if (a === b) continue;
      if (pivotWorksheetKey(first.def) !== pivotWorksheetKey(other.def) || !matchingPivotCacheData(wb,first.def,other.def)) throw new Error('연결된 피벗 테이블의 원본 또는 저장된 데이터가 다릅니다. 같은 원본으로 새로 고친 뒤 다시 저장하거나 보고서 연결을 해제하세요.');
      parents.set(b,a);
    }
  }));
  const exportCacheKey = def => rootKey(originalCacheKeys.get(def));
  const slicerFields = new Map(); // 원본 키 → 슬라이서 필드 이름(소문자)
  wb.sheets.forEach((s, i) => (s.slicers ?? []).forEach((sl) => {
    for (const ref of slicerPivotRefs(wb, i, sl)) {
      const e = allDefs.find((x) => x.si === ref.si && (ref.name === null ? x.j === 0 : String(x.def.name ?? '') === ref.name));
      if (!e) continue;
      const k = exportCacheKey(e.def);
      if (!slicerFields.has(k)) slicerFields.set(k, new Set());
      slicerFields.get(k).add(String(sl.source.field).toLowerCase());
    }
  }));
  const caches = new Map();
  let cacheSeq = 0;
  for (const e of allDefs) {
    const k = exportCacheKey(e.def);
    if (!caches.has(k)) caches.set(k, { defs: [] });
    caches.get(k).defs.push(e.def);
  }
  for (const [k, g] of caches) {
    g.cache = buildPivotCache(wb, g.defs, ++cacheSeq, slicerFields.get(k), pool);
    if (!g.cache && g.defs.some(importedPivotSourceReference)) throw new Error('피벗 테이블의 원본을 이 문서에서 찾을 수 없고 사용할 수 있는 저장 캐시가 없어 저장할 수 없습니다. 원본 연결을 확인하세요.');
    yield { p: 0.01, msg: '피벗 캐시 만드는 중' };
  }
  const pivotList = new Map(); // 시트 번호 → [{ name, cacheId, tableNo, parts, cache, j }]
  let pivotNo = 0;
  for (const e of allDefs) {
    const cache = caches.get(exportCacheKey(e.def)).cache;
    if (!cache) continue;
    const used = new Set((pivotList.get(e.si) ?? []).map((x) => x.name.toLowerCase()));
    let name = e.def.name || `피벗 테이블${pivotNo + 1}`;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${e.def.name || '피벗 테이블'}${n}`;
    const parts = pivotParts(wb, e.si, e.def, cache, name, pool);
    if (!parts) {
      if (importedPivotSourceReference(e.def) && [...(e.def.rows ?? []), ...(e.def.cols ?? []), ...(e.def.pages ?? []), ...(e.def.values ?? [])].length) throw new Error('피벗 테이블의 저장 캐시 필드와 배치 필드가 일치하지 않아 저장할 수 없습니다. 원본 연결과 필드를 확인하세요.');
      continue;
    }
    pivotNo++;
    const info = { name, origName: e.def.name ?? null, cacheId: cache.cacheId, tableNo: pivotNo, parts, cache, j: e.j };
    if (!pivotList.has(e.si)) pivotList.set(e.si, []);
    pivotList.get(e.si).push(info);
  }
  const findPivotInfo = (ref) => (pivotList.get(ref.si) ?? []).find((x) => (ref.name === null ? x.j === 0 : x.origName === ref.name || x.name === ref.name));
  const pivotCaches = [];
  const standaloneSlicerCaches = new Map();
  const sharedPivotSlicerCaches = new Map();
  const slicerCachesPivot = [];
  const slicerCachesTable = [];
  const usedCacheNames = new Set();
  const usedSlicerNames = new Set();
  let slicerPartNo = 0;
  let slicerCacheNo = 0;

  // 셀 XML (가장 큰 부분): 시트마다 미리 만들며 중간중간 멈춤 — 큰 문서를 저장해도 화면이 멈추지 않게
  const expMemo = new Map(); // 같은 수식(표의 계산 열 등)은 한 번만 변환
  const shapeSame = new Map(); // 수식 모양 → 파일 형식이 앱 형식과 같은지 (채우기로 만든 수식 백만 개도 해석은 한 번)
  const astSame = new WeakMap(); // 같은 AST(같은 모양으로 채운 수식) → 파일 형식이 앱 형식과 같은지
  const exportF = (raw, t, dyn, ast = null) => {
    if (ast && !t && !dyn) {
      const same = astSame.get(ast);
      if (same === true) return xesc(raw.slice(1));
      if (same === undefined) {
        const out = exportFormula(raw, t, dyn);
        astSame.set(ast, out === raw.slice(1));
        return xesc(out);
      }
    }
    const k = `${t ?? ''}\u0001${dyn ? 1 : 0}\u0001${raw}`;
    let v = expMemo.get(k);
    if (v !== undefined) return v;
    const body = raw.startsWith('=') ? raw.slice(1) : raw;
    const shape = !t && !dyn ? formulaShape(body) : null;
    if (shape !== null && shapeSame.get(shape) === true) return xesc(body);
    const out = exportFormula(raw, t, dyn);
    if (shape !== null && !shapeSame.has(shape)) shapeSame.set(shape, out === body);
    v = xesc(out);
    if (expMemo.size < 200000) expMemo.set(k, v);
    return v;
  };
  const colLetter = [];
  const refOf = (r, c) => (colLetter[c] ??= cellName(0, c).replace(/\d+$/, '')) + (r + 1);
  const sheetRows = [];
  let rowsDone = 0;
  // 외부 통합 문서 값 시트(맨 뒤)는 시트로 쓰지 않음 — externalLink 원본으로 되돌려 씀
  const nOwn = wb.ownSheetCount ? wb.ownSheetCount() : wb.sheets.length;
  const rowsTotal = wb.sheets.slice(0, nOwn).reduce((n, sh) => n + sh.cells.size, 0) || 1;
  function* buildSheetRows(si) {
    const sheet = wb.sheets[si];
    const plainStyle = !sheet.allStyle && !Object.keys(sheet.colStyles ?? {}).length && !Object.keys(sheet.rowStyles ?? {}).length;
    // 열 · 행 · 셀 서식을 합친 서식 번호: 같은 조합(서식 객체 셋)은 한 번만 합치고 찾음 (셀마다 새 객체를 만들지 않게)
    const xfMemo = new Map();
    const noKey = {};
    const xfAt = (r, c, cell) => {
      const own = cell.style;
      if (!own && sheet.blocks.length) return pool.xf(wb.styleAt(si, r, c));
      const col = sheet.colStyles[c] ?? noKey;
      const row = sheet.rowStyles[r] ?? noKey;
      let m1 = xfMemo.get(col);
      if (!m1) { m1 = new Map(); xfMemo.set(col, m1); }
      let m2 = m1.get(row);
      if (!m2) { m2 = new Map(); m1.set(row, m2); }
      const k = own ?? noKey;
      let id = m2.get(k);
      if (id === undefined) { id = pool.xf(wb.styleAt(si, r, c)); m2.set(k, id); }
      return id;
    };
    const blocks = sheet.blocks ?? [];
    const plan = (streaming ? xlsxBoundedRowPlan : xlsxRowPlan)(sheet.cells, {
      rowLimit: EXCEL_MAX_ROWS, colLimit: MAX_COLS, blocks, spills: wb.spillsOf(si),
      extraRows: [...Object.keys(sheet.rowHeights), ...Object.keys(sheet.hiddenRows), ...Object.keys(sheet.rowStyles), ...Object.keys(sheet.outline?.rows ?? {}), ...Object.keys(sheet.outline?.rowsColl ?? {}), ...hidKeys(sheet.filter?.hidden, EXCEL_MAX_ROWS), ...(sheet.tables ?? []).flatMap(t => hidKeys(t.filter?.hidden, EXCEL_MAX_ROWS))],
    });
    const { maxR, maxC } = plan;
    function* rowParts() {
    let rowCount = 0;
    for (const [r, rowCells] of plan.rows) {
      let cells = rowCells;
      for (const b of blocks) {
        if (r < b.r0 || r >= b.r0 + b.n) continue;
        const i = b.perm ? b.perm[r - b.r0] : r - b.r0;
        const taken = cells.length ? new Set(cells.map((x) => x[0])) : null;
        for (let j = 0; j < Math.min(b.cols.length, MAX_COLS - b.c0); j++) {
          const col = b.cols[j];
          const has = (col.str && col.str[i] >= 0) || (col.num && col.num[i] === col.num[i]);
          if (!has && !col.fmt) continue;
          if (taken?.has(b.c0 + j)) continue;
          cells.push([b.c0 + j, has ? { raw: '1', style: col.fmt ?? undefined } : { raw: '', style: col.fmt }]);
        }
      }
      cells = cells.sort((a, b) => a[0] - b[0]);
      const attrs = [`r="${r + 1}"`];
      if (sheet.rowHeights[r] !== undefined) {
        attrs.push(`ht="${px2pt(sheet.rowHeights[r])}"`);
        if (sheet.rowManual[r]) attrs.push('customHeight="1"');
      }
      if (sheet.hiddenRows[r] || hid(sheet.filter?.hidden, r) || (sheet.tables ?? []).some((t) => hid(t.filter?.hidden, r))) attrs.push('hidden="1"');
      if (sheet.rowStyles[r]) attrs.push(`s="${pool.xf({ ...sheet.allStyle, ...sheet.rowStyles[r] })}"`, 'customFormat="1"');
      if (sheet.outline?.rows?.[r]) attrs.push(`outlineLevel="${sheet.outline.rows[r]}"`);
      if (sheet.outline?.rowsColl?.[r]) attrs.push('collapsed="1"');
      const cellXml = ([c, cell]) => {
        const ref = refOf(r, c);
        const spillKey = `${si}:${r},${c}`;
        const querySpill = wb.spills.get(wb.spillOwner.get(spillKey) ?? spillKey);
        const queryFormat = querySpill && r - querySpill.r >= querySpill.formatStart ? querySpill.formats?.[c - querySpill.c] : null;
        const s = queryFormat != null || querySpill?.cellFormats?.[r - querySpill.r]?.[c - querySpill.c] != null ? pool.xf(wb.styleAt(si, r, c)) : plainStyle ? pool.xf(cell.style ?? {}) : xfAt(r, c, cell);
        const sAttr = s ? ` s="${s}"` : '';
        const v = wb.getValue(si, r, c);
        const cachedBlank = querySpill?.cachedArray && arrayCacheValue(querySpill.cachedArray, r - querySpill.r, c - querySpill.c) === null;
        if (!cell.raw && cell.image?.src) {
          const vm = richImage(cell.image);
          return vm ? `<c r="${ref}"${sAttr} t="e" vm="${vm}"><v>#VALUE!</v></c>` : (s ? `<c r="${ref}"${sAttr}/>` : '');
        }
        // 빈 칸: 기본 서식(xf 0)이어도 행 · 열 서식이 있으면 적어 둠 (없으면 다시 열 때 행 서식을 물려받음)
        if (!cell.raw && (v === null && !cachedBlank || !cell.spilled && !wb.spillAnchorOf(si, r, c))) return s || (cell.style && (sheet.rowStyles[r] || sheet.colStyles[c])) ? `<c r="${ref}" s="${s}"/>` : '';
        if (cell.formula) {
          // 배열을 돌려줄 수 있는 수식은 동적 배열 수식으로 (cm="1" + t="array")
          const dyn = !!cell.maybeArray;
          const sp = dyn ? wb.spillRange(si, r, c) : null;
          if (dyn) dynamicCells++;
          const fAttr = dyn ? ` t="array" ref="${sp ? rangeRef({ ...sp, r2: Math.min(sp.r2, EXCEL_MAX_ROWS - 1) }) : ref}" aca="false"` : '';
          const cm = dyn ? ' cm="1"' : '';
          const f = `<f${fAttr}>${exportF(cell.raw, cell.raw.includes('[') ? tableAt(sheet, r, c)?.name : null, dyn, cell.ast)}</f>`;
          if (v === null && cachedBlank) return `<c r="${ref}"${sAttr}${cm}>${f}<v/></c>`;
          if (typeof v === 'number') return `<c r="${ref}"${sAttr}${cm}>${f}<v>${v}</v></c>`;
          if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"${cm}>${f}<v>${v ? 1 : 0}</v></c>`;
          if (isError(v)) return `<c r="${ref}"${sAttr} t="e"${cm}>${f}<v>${esc(['#CIRC!', '#SPILL!', '#CALC!', '#BUSY!'].includes(v.code) && !dyn ? '#REF!' : v.code === '#CIRC!' ? '#REF!' : v.code)}</v></c>`;
          if (v instanceof CellImage) return `<c r="${ref}"${sAttr} t="e"${cm}>${f}<v>#VALUE!</v></c>`; // IMAGE: 엑셀이 다시 계산
          return `<c r="${ref}"${sAttr} t="str"${cm}>${f}<v>${xesc(v ?? '')}</v></c>`;
        }
        if (isError(v)) return `<c r="${ref}"${sAttr} t="e"><v>${esc(v.code)}</v></c>`;
        if (v instanceof CellImage) return s ? `<c r="${ref}"${sAttr}/>` : '';
        if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
        if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
        if (v === null) return cachedBlank ? `<c r="${ref}"${sAttr}><v/></c>` : s ? `<c r="${ref}"${sAttr}/>` : '';
        // 배열의 후속 셀도 수식 결과입니다. SST 문자열로 쓰면 Excel이 #VALUE!로 읽습니다.
        if (!cell.raw && querySpill && (r !== querySpill.r || c !== querySpill.c)) return `<c r="${ref}"${sAttr} t="str"><v>${xesc(v)}</v></c>`;
        if (cell.phonetic) {
          const p = normalizePhonetic(cell.phonetic, String(v));
          const f = p.font ?? { size: Math.max(6, (wb.defaultFont?.size ?? 11) * .55) };
          const xf = pool.xf(f), fontId = Number(/fontId="(\d+)"/.exec(pool.xfs[xf])?.[1] ?? 0);
          return `<c r="${ref}"${sAttr} t="inlineStr" ph="${p.visible ? 1 : 0}"><is><t xml:space="preserve">${xesc(v)}</t>${phoneticXml(String(v), p, fontId, xesc)}</is></c>`;
        }
        return `<c r="${ref}"${sAttr} t="s"><v>${sst(String(v))}</v></c>`;
      };
      if (streaming) {
        yield `<row ${attrs.join(' ')}>`;
        for (const cell of cells) yield cellXml(cell);
        yield '</row>';
      } else yield `<row ${attrs.join(' ')}>${cells.map(cellXml).join('')}</row>`;
      rowsDone += cells.length;
      if (++rowCount % 1000 === 0) onRows?.({ p: Math.min(0.85, 0.85 * rowsDone / rowsTotal), msg: `'${sheet.name}' 시트 저장 중` });
    }
    }
    if (streaming) return { rowXml: rowParts(), maxR, maxC };
    const rowXml = createXmlChunks();
    for (const text of rowParts()) {
      rowXml.push(text);
      if (rowXml.count % 1000 === 0) yield { p: 0.85 * rowsDone / rowsTotal, msg: `'${sheet.name}' 시트 저장 중` };
    }
    return { rowXml: rowXml.finish(), maxR, maxC };
  }
  if (!streaming) for (let si = 0; si < nOwn; si++) {
    const rows = yield* buildSheetRows(si);
    sheetRows.push(rows);
  }

  for (let si = 0; si < nOwn; si++) {
    const sheet = wb.sheets[si];
    const sheetRels = [];
    const addRel = (type, target) => { const id = `rId${sheetRels.length + 1}`; sheetRels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`); return id; };

    const { rowXml, maxR, maxC } = streaming ? yield* buildSheetRows(si) : sheetRows[si];

    // 열
    const olc = sheet.outline?.cols ?? {};
    const olcc = sheet.outline?.colsColl ?? {};
    const colKeys = new Set([...Object.keys(sheet.colWidths), ...Object.keys(sheet.hiddenCols), ...Object.keys(sheet.colStyles), ...Object.keys(olc), ...Object.keys(olcc)].map(Number));
    const colsXml = [...colKeys].filter(c => c < MAX_COLS).sort((a, b) => a - b).map((c) => {
      const w = sheet.colWidths[c] ?? sheet.defColW ?? DEFAULT_COL_WIDTH;
      const st = sheet.colStyles[c] ? ` style="${pool.xf({ ...sheet.allStyle, ...sheet.colStyles[c] })}"` : '';
      const ol = (olc[c] ? ` outlineLevel="${olc[c]}"` : '') + (olcc[c] ? ' collapsed="1"' : '');
      return `<col min="${c + 1}" max="${c + 1}" width="${px2widthM(w, wmdw)}"${sheet.colWidths[c] !== undefined ? ' customWidth="1"' : ''}${sheet.hiddenCols[c] ? ' hidden="1"' : ''}${st}${ol}/>`;
    }).join('');
    const olMax = (o) => Object.values(o ?? {}).reduce((m, v) => Math.max(m, v), 0);
    const olRowMax = olMax(sheet.outline?.rows);
    const olColMax = olMax(olc);
    const pgx = pageXml(sheet.page, esc);
    const olPr = (sheet.outline && (sheet.outline.below === false || sheet.outline.right === false)
      ? `<outlinePr${sheet.outline.below === false ? ' summaryBelow="0"' : ''}${sheet.outline.right === false ? ' summaryRight="0"' : ''}/>` : '')
      + (pgx.fitToPage ? '<pageSetUpPr fitToPage="1"/>' : '');
    // 인쇄 영역 · 인쇄 제목 (이름 정의)
    const printAreas = sheet.page?.areas?.length ? sheet.page.areas : sheet.page?.area ? [sheet.page.area] : [];
    if (printAreas.length) definedNames.push(`<definedName name="_xlnm.Print_Area" localSheetId="${si}">${esc(printAreas.map(rg => `${quoteSheetName(sheet.name)}!${printAreaRef(rg)}`).join(','))}</definedName>`);
    if (sheet.page?.titleRows || sheet.page?.titleCols) {
      const parts = [];
      if (sheet.page.titleCols) parts.push(`${quoteSheetName(sheet.name)}!$${colToName(sheet.page.titleCols[0])}:$${colToName(sheet.page.titleCols[1])}`);
      if (sheet.page.titleRows) parts.push(`${quoteSheetName(sheet.name)}!$${sheet.page.titleRows[0] + 1}:$${sheet.page.titleRows[1] + 1}`);
      definedNames.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${si}">${esc(parts.join(','))}</definedName>`);
    }

    // 틀 고정
    const viewIndex = (v, max) => Number.isFinite(Number(v)) ? Math.max(0, Math.min(max - 1, Math.floor(Number(v)))) : 0;
    const fr = viewIndex(sheet.freeze?.rows, EXCEL_MAX_ROWS);
    const fc = viewIndex(sheet.freeze?.cols, MAX_COLS);
    const ft = fr ? Math.min(viewIndex(sheet.freeze?.top, EXCEL_MAX_ROWS), EXCEL_MAX_ROWS - 1 - fr) : 0;
    const fl = fc ? Math.min(viewIndex(sheet.freeze?.left, MAX_COLS), MAX_COLS - 1 - fc) : 0;
    const vt = Math.max(fr ? ft + fr : 0, viewIndex(sheet.view?.top, EXCEL_MAX_ROWS));
    const vl = Math.max(fc ? fl + fc : 0, viewIndex(sheet.view?.left, MAX_COLS));
    const viewTopLeft = cellName(fr ? ft : vt, fc ? fl : vl);
    const activeCell = Number.isFinite(sheet.view?.r) && Number.isFinite(sheet.view?.c)
      ? cellName(viewIndex(sheet.view.r, EXCEL_MAX_ROWS), viewIndex(sheet.view.c, MAX_COLS)) : null;
    const selectionAttrs = activeCell ? ` activeCell="${activeCell}" sqref="${activeCell}"` : '';
    let pane = '';
    if (fr || fc) {
      const available = fr && fc ? ['topLeft', 'topRight', 'bottomLeft', 'bottomRight'] : fr ? ['topLeft', 'bottomLeft'] : ['topLeft', 'topRight'];
      const activePane = available.includes(sheet.view?.activePane) ? sheet.view.activePane : available[available.length - 1];
      pane = `<pane${fc ? ` xSplit="${fc}"` : ''}${fr ? ` ySplit="${fr}"` : ''} topLeftCell="${cellName(vt, vl)}" activePane="${activePane}" state="frozen"/><selection pane="${activePane}"${selectionAttrs}/>`;
    } else if (activeCell) pane = `<selection${selectionAttrs}/>`;
    const dim = sheet.cells.size || sheet.blocks?.length ? `A1:${cellName(maxR, maxC)}` : 'A1';

    // 필터
    let autoFilter = '';
    if (sheet.filter && sheet.filter.r1 < EXCEL_MAX_ROWS) {
      const f = { ...sheet.filter, r2: Math.min(sheet.filter.r2, EXCEL_MAX_ROWS - 1) };
      autoFilter = autoFilterXml(f, f, { dxf: style => pool.objectDxf({style}) });
      definedNames.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${si}" hidden="1">${esc(`${quoteSheetName(sheet.name)}!${rangeRef(f, true)}`)}</definedName>`);
    }
    const fit = (rg) => (rg.r1 >= EXCEL_MAX_ROWS ? null : { ...rg, r2: Math.min(rg.r2, EXCEL_MAX_ROWS - 1) });
    const mergeList = sheet.merges.map(fit).filter(Boolean);
    const merges = mergeList.length ? `<mergeCells count="${mergeList.length}">${mergeList.map((m) => `<mergeCell ref="${rangeRef(m)}"/>`).join('')}</mergeCells>` : '';
    const cfX14 = [];
    const cf = sheet.cond.map((rule) => {
      const g = fit(rule);
      return g && { ...g, ...(rule.more ? { more: rule.more.map(fit).filter(Boolean) } : {}) };
    }).filter(Boolean).map((rule, i) => cfXml(rule, pool, i + 1, cfX14, wb.date1904)).join('');

    // 그림 개체 (차트 · 그림 · 도형)
    let drawing = '';
    const images = sheet.images ?? [];
    const shapes = sheet.shapes ?? [];
    // 슬라이서: 이름·캐시를 정하고 그림 개체와 함께 그림
    const sheetSlicers = { table: [], pivot: [] };
    for (const sl of sheet.slicers ?? []) {
      let cacheXml = null;
      let kind = null;
      let standaloneCache = null;
      let sharedCache = null;
      let cacheName = null;
      // 슬라이서 설정 → 캐시 속성
      const setAttrs = `${sl.sort === 'desc' ? ' sortOrder="descending"' : ''}${sl.customList === false ? ' customListSort="0"' : ''}${sl.markNoData === false ? ' crossFilter="none"' : sl.noDataLast === false ? ' crossFilter="showItemsWithNoData"' : ''}`;
      const hideExt = sl.hideNoData ? `<x:ext uri="{470722E0-AACD-4C17-9CDC-17EF765DBC7E}" xmlns:x15="${NS_X15}"><x15:slicerCacheHideItemsWithNoData/></x:ext>` : '';
      if (sl.source?.kind === 'table') {
        const f = findTable(wb, sl.source.table);
        const tid = f && tableIds.get(f.t.name.toLowerCase());
        if (!tid) continue;
        const names = columnNames(wb, f.si, f.t).map((n) => n.toLowerCase());
        const col = names.indexOf(String(sl.source.column).toLowerCase());
        if (col < 0) continue;
        kind = 'table';
        cacheName = cacheNameFor(sl.source.column, usedCacheNames);
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(cacheName)}" sourceName="${esc(columnNames(wb, f.si, f.t)[col])}"><extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${NS_X15}"><x15:tableSlicerCache tableId="${tid}" column="${col + 1}"${setAttrs}/></x:ext>${hideExt}</extLst></slicerCacheDefinition>`;
      } else if (sl.source?.kind === 'cache') {
        const source = sl.source, values = source.values ?? [];
        const cacheKey = JSON.stringify([source.cacheKey, source.field]);
        standaloneCache = standaloneSlicerCaches.get(cacheKey);
        if (!standaloneCache) {
          standaloneCache = { id: ++cacheSeq, name: cacheNameFor(source.field, usedCacheNames) };
          standaloneSlicerCaches.set(cacheKey, standaloneCache);
          const cacheId = standaloneCache.id;
          const original = source.cacheSource ?? {};
          const ref = original.ref ? ` ref="${esc(original.ref)}"` : '';
          const name = original.name ? ` name="${esc(original.name)}"` : '';
          const sheetName = original.sheet ? ` sheet="${esc(original.sheet)}"` : '';
          const shared = values.map(v => v === null || v === undefined ? '<m/>' : typeof v === 'number' ? `<n v="${v}"/>` : typeof v === 'boolean' ? `<b v="${v ? 1 : 0}"/>` : v?.error ? `<e v="${esc(v.error)}"/>` : `<s v="${esc(v)}"/>`).join('');
          const nums = values.filter(v => typeof v === 'number');
          const nonNumbers = values.some(v => v != null && typeof v !== 'number');
          const sharedAttrs = `${values.some(v => v == null) ? ' containsBlank="1"' : ''} containsString="${values.some(v => typeof v === 'string') ? 1 : 0}" containsNumber="${nums.length ? 1 : 0}" containsMixedTypes="${nums.length && nonNumbers ? 1 : 0}" containsSemiMixedTypes="${nonNumbers ? 1 : 0}"${nums.length ? ` containsInteger="${nums.every(Number.isInteger) ? 1 : 0}" minValue="${minOf(nums)}" maxValue="${maxOf(nums)}"` : ''}`;
          files[`xl/pivotCache/pivotCacheDefinition${cacheId}.xml`] = `<?xml version="1.0" encoding="UTF-8"?><pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" saveData="0" refreshOnLoad="0" recordCount="0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3"><cacheSource type="worksheet"><worksheetSource${ref}${name}${sheetName}${original.external ? ' r:id="rId1"' : ''}/></cacheSource><cacheFields count="1"><cacheField name="${esc(source.field)}" numFmtId="${pool.fmtId(source.format ?? {})}"><sharedItems count="${values.length}"${sharedAttrs}>${shared}</sharedItems></cacheField></cacheFields><extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${cacheId}"/></ext></extLst></pivotCacheDefinition>`;
          // The external source is metadata only. Never resolve or fetch it while displaying saved items.
          if (original.external) files[`xl/pivotCache/_rels/pivotCacheDefinition${cacheId}.xml.rels`] = `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/externalLinkPath" Target="${esc(original.external)}" TargetMode="External"/></Relationships>`;
          contentOverrides.push(`<Override PartName="/xl/pivotCache/pivotCacheDefinition${cacheId}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`);
          pivotCaches.push({ cacheId, target: `pivotCache/pivotCacheDefinition${cacheId}.xml` });
        }
        const items = (source.items ?? values.map((_, index) => ({ index, hasData: true }))).filter(it => Number.isInteger(it.index) && it.index >= 0 && it.index < values.length);
        const selected = Array.isArray(sl.cacheSelection) ? new Set(sl.cacheSelection) : null;
        kind = 'pivot'; cacheName = standaloneCache.name; sharedCache = standaloneCache;
        const cacheId = standaloneCache.id;
        const itemXml = items.map(it => `<i x="${it.index}"${!selected || selected.has(String(it.index)) ? ' s="1"' : ''}${it.hasData === false ? ' nd="1"' : ''}/>`).join('');
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(cacheName)}" sourceName="${esc(source.field)}"><data><tabular pivotCacheId="${cacheId}"${setAttrs}${sl.showDeleted ? '' : ' showMissing="0"'}><items count="${items.length}">${itemXml}</items></tabular></data>${hideExt ? `<extLst>${hideExt}</extLst>` : ''}</slicerCacheDefinition>`;
      } else {
        // 연결된 피벗 중 첫 피벗과 같은 캐시를 쓰는 것만 (엑셀 규칙)
        const refs = slicerPivotRefs(wb, si, sl).map((ref) => ({ ref, info: findPivotInfo(ref) })).filter((x) => x.info);
        const first = refs[0]?.info;
        const it = first?.parts.slicerItems(sl.source.field);
        if (!it) continue;
        const linked = [...new Map(refs.filter(x => x.info.cacheId === first.cacheId).map(x => [JSON.stringify([x.ref.si, x.info.name.toLowerCase()]), x])).values()];
        const connectionKey = linked.map(x => [x.ref.si, x.info.name.toLowerCase()]).sort((a,b) => a[0]-b[0] || a[1].localeCompare(b[1]));
        const sharedKey = JSON.stringify([first.cacheId, it.field.toLowerCase(), connectionKey]);
        // Excel은 같은 필드·피벗 연결을 여러 물리 슬라이서가 하나의 캐시로 공유합니다.
        // 캐시 설정은 첫 개체가 정하고, 이름·스타일·위치는 각 슬라이서에 보존합니다.
        sharedCache = sharedPivotSlicerCaches.get(sharedKey);
        if (!sharedCache) { sharedCache = { name: cacheNameFor(it.field, usedCacheNames) }; sharedPivotSlicerCaches.set(sharedKey, sharedCache); }
        kind = 'pivot';
        cacheName = sharedCache.name;
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(cacheName)}" sourceName="${esc(it.field)}"><pivotTables>${linked.map((x) => `<pivotTable tabId="${x.ref.si + 1}" name="${esc(x.info.name)}"/>`).join('')}</pivotTables><data><tabular pivotCacheId="${first.cacheId}"${setAttrs}${sl.showDeleted ? '' : ' showMissing="0"'}><items count="${it.count}">${it.xml}</items></tabular></data>${hideExt ? `<extLst>${hideExt}</extLst>` : ''}</slicerCacheDefinition>`;
      }
      if (!sharedCache?.written) {
      slicerCacheNo++;
      files[`xl/slicerCaches/slicerCache${slicerCacheNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${cacheXml}`;
      contentOverrides.push(`<Override PartName="/xl/slicerCaches/slicerCache${slicerCacheNo}.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>`);
      (kind === 'table' ? slicerCachesTable : slicerCachesPivot).push(`slicerCaches/slicerCache${slicerCacheNo}.xml`);
      definedNames.push(`<definedName name="${esc(cacheName)}">#N/A</definedName>`);
      if (sharedCache) sharedCache.written = true;
      }
      let nm = sl.caption || sl.source.column || sl.source.field || '슬라이서';
      for (let i = 1; usedSlicerNames.has(nm.toLowerCase()); i++) nm = `${sl.caption || '슬라이서'} ${i}`;
      usedSlicerNames.add(nm.toLowerCase());
      sheetSlicers[kind].push({ sl, name: nm, cache: cacheName });
    }
    const hasSlicers = sheetSlicers.table.length + sheetSlicers.pivot.length > 0;
    if (sheet.charts.length || images.length || shapes.length || hasSlicers) {
      drawingNo++;
      const drawingRels = [];
      const drel = (type, target, mode) => { const id = `rId${drawingRels.length + 1}`; drawingRels.push(`<Relationship Id="${id}" Type="${type.startsWith('http:') ? type : `${REL}/${type}`}" Target="${esc(target)}"${mode ? ` TargetMode="${mode}"` : ''}/>`); return id; };
      const hyperlinkXml = object => drawingHyperlinkXml(object?.hyperlink, (target, mode) => drel('hyperlink', target, mode));
      const { colAxis, rowAxis } = objectAxes(sheet);
      const anchorAt = (x, y) => {
        const c = colAxis.indexAt(Math.max(0, x));
        const r = rowAxis.indexAt(Math.max(0, y));
        return `<xdr:col>${c}</xdr:col><xdr:colOff>${Math.max(0, Math.round((x - colAxis.pos(c)) * EMU))}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${Math.max(0, Math.round((y - rowAxis.pos(r)) * EMU))}</xdr:rowOff>`;
      };
      const anchor = (o, body) => `<xdr:twoCellAnchor editAs="${o.placement ?? 'twoCell'}"><xdr:from>${anchorAt(o.x, o.y)}</xdr:from><xdr:to>${anchorAt(o.x + o.w, o.y + o.h)}</xdr:to>${body}<xdr:clientData${o.noPrint !== undefined ? ` fPrintsWithSheet="${o.noPrint ? 0 : 1}"` : ''}${o.locked !== undefined ? ` fLocksWithSheet="${o.locked === false ? 0 : 1}"` : ''}/></xdr:twoCellAnchor>`;
      const xfrm = (o) => `<a:xfrm${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ''}${o.flip ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}><a:off x="${Math.round(o.x * EMU)}" y="${Math.round(o.y * EMU)}"/><a:ext cx="${Math.round(o.w * EMU)}" cy="${Math.round(o.h * EMU)}"/></a:xfrm>`;
      const pictureRelationships = new Map();
      const embedPictureSource = source => {
        if (pictureRelationships.has(source)) return pictureRelationships.get(source);
        let target = drawingMediaTargets.get(source);
        if (!target) {
          const m = /^data:([^;,]+);base64,(.*)$/s.exec(source ?? ''); if (!m) return null;
          const ext = Object.keys(MIME).find(k => MIME[k] === m[1]); if (!ext) return null;
          mediaNo++; mediaExts.add(ext); target = `../media/image${mediaNo}.${ext}`;
          files[`xl/media/image${mediaNo}.${ext}`] = fromBase64(m[2]);
          drawingMediaTargets.set(source, target);
        }
        const rel = drel('image', target); pictureRelationships.set(source, rel); return rel;
      };
      let objId = 1;
      const parts = [];
      const ordered = [
        ...sheet.charts.map((o) => ['chart', o]), ...images.map((o) => ['image', o]), ...shapes.map((o) => ['shape', o]),
        ...sheetSlicers.table.map((o) => ['slicerTable', o]), ...sheetSlicers.pivot.map((o) => ['slicerPivot', o]),
      ].sort((a, b) => ((a[1].sl ?? a[1]).z ?? 0) - ((b[1].sl ?? b[1]).z ?? 0));
      const slicerGroups = new Map(), writtenSlicerGroups = new Set();
      for (const entry of ordered) {
        const group = entry[1].sl?.objectGroup;
        if (typeof group !== 'string' || !group) continue;
        if (!slicerGroups.has(group)) slicerGroups.set(group, []);
        slicerGroups.get(group).push(entry);
      }
      for (const [kind, o] of ordered) {
        if (kind === 'chart') {
          const ch = o;
          chartNo++;
          objId++;
          const chartImages = [], chartImageIds = new Map();
          const chartImageRel = source => {
            if (chartImageIds.has(source)) return chartImageIds.get(source);
            const m = /^data:([^;,]+);base64,(.*)$/s.exec(source ?? ''); if (!m) return null;
            const ext = Object.keys(MIME).find(k => MIME[k] === m[1]); if (!ext) return null;
            mediaNo++; mediaExts.add(ext); files[`xl/media/image${mediaNo}.${ext}`] = fromBase64(m[2]);
            const id = `rIdAreaImage${chartImages.length + 1}`;
            chartImages.push(`<Relationship Id="${id}" Type="${NS_R}/image" Target="../media/image${mediaNo}.${ext}"/>`); chartImageIds.set(source, id); return id;
          };
          files[`xl/charts/chart${chartNo}.xml`] = chartXml(wb, si, ch, fileName, chartImageRel);
          contentOverrides.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="${isChartEx(ch) ? CHARTEX_CONTENT : 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'}"/>`);
          if (isChartEx(ch)) {
            files[`xl/charts/style${chartNo}.xml`] = chartExStyleXml();
            files[`xl/charts/colors${chartNo}.xml`] = chartExColorsXml(paletteOf(ch));
            files[`xl/charts/_rels/chart${chartNo}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.microsoft.com/office/2011/relationships/chartStyle" Target="style${chartNo}.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/chartColorStyle" Target="colors${chartNo}.xml"/></Relationships>`;
            contentOverrides.push(`<Override PartName="/xl/charts/style${chartNo}.xml" ContentType="${CHARTEX_STYLE_CONTENT}"/>`, `<Override PartName="/xl/charts/colors${chartNo}.xml" ContentType="${CHARTEX_COLOR_CONTENT}"/>`);
          }
          if (chartImages.length) {
            const path = `xl/charts/_rels/chart${chartNo}.xml.rels`;
            const rel = chartImages.join('');
            files[path] = files[path] ? files[path].replace('</Relationships>', rel + '</Relationships>') : `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${rel}</Relationships>`;
          }
          const id = drel(isChartEx(ch) ? CHARTEX_REL : 'chart', `../charts/chart${chartNo}.xml`);
          const chartNamespace = isChartEx(ch) ? CHARTEX_NS : 'http://schemas.openxmlformats.org/drawingml/2006/chart';
          parts.push(anchor(ch, `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${objId}" name="차트 ${objId - 1}"${ch.hidden ? ' hidden="1"' : ''}>${isChartEx(ch) ? chartExDrawingProps(ch, paletteOf(ch)) : ''}</xdr:cNvPr><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="${chartNamespace}"><c:chart xmlns:c="${chartNamespace}" r:id="${id}"/></a:graphicData></a:graphic></xdr:graphicFrame>`));
        } else if (kind === 'image') {
          const im = o, source = pictureMediaSource(im) ?? (im.linked ? 'data:image/png;base64,' + BLANK_PNG : null);
          const id = embedPictureSource(source); if (!id) continue;
          const svgSrc = !im.effectPng && im.png && /^data:image\/svg\+xml;base64,/.test(im.src ?? '') ? im.src : null;
          let svgExt = '';
          if (svgSrc) {
            const sid = embedPictureSource(svgSrc);
            svgExt = `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${sid}"/></a:ext></a:extLst>`;
          }
          const extensions = (im.linked ? `<a:ext uri="${LINKED_PIC_URI}"><wx:linked xmlns:wx="https://wixel.app/x" ref="${esc(linkedRef(im.linked))}"/></a:ext>` : '')
            + (normalizeVideo(im.media) ? `<a:ext uri="${MEDIA_OBJECT_URI}"><wx:video xmlns:wx="https://wixel.app/x" json="${esc(JSON.stringify(normalizeVideo(im.media)))}"/></a:ext>` : '')
            + pictureMetadataXml(im, embedPictureSource);
          objId++;
          const native = `<xdr:pic${im.macro ? ` macro="[0]!${esc(im.macro)}"` : ''}><xdr:nvPicPr><xdr:cNvPr id="${objId}" name="${esc(im.name || `그림 ${objId - 1}`)}"${im.alt !== undefined ? ` descr="${esc(im.alt)}"` : ''}${im.hidden ? ' hidden="1"' : ''}>${hyperlinkXml(im)}${extensions ? `<a:extLst>${extensions}</a:extLst>` : ''}</xdr:cNvPr><xdr:cNvPicPr><a:picLocks noChangeAspect="${im.lockAspect === false ? 0 : 1}"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${id}">${pictureBlipXml(im)}${svgExt}</a:blip>${im.crop ? `<a:srcRect${['l','t','r','b'].map(k => ` ${k}="${Math.round((im.crop[k] ?? 0) * 100000)}"`).join('')}/>` : ''}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(im)}${pictureGeometryXml(im)}${pictureBorderXml(im)}${pictureEffectXml(im)}</xdr:spPr></xdr:pic>`;
          parts.push(anchor(im, native));
        } else if (kind === 'slicerTable' || kind === 'slicerPivot') {
          const group = o.sl.objectGroup, members = slicerGroups.get(group);
          if (members?.length > 1) {
            if (writtenSlicerGroups.has(group)) continue;
            writtenSlicerGroups.add(group); objId++;
            parts.push(slicerGroupAnchorXml(members, objId, () => ++objId, anchorAt));
            continue;
          }
          objId++;
          parts.push(slicerAnchorXml(o.sl, o.name, objId, anchorAt, kind === 'slicerTable' ? 'table' : 'pivot'));
        } else if (isDrawingGroup(o)) {
          objId++;
          const groupXml = drawingGroupXml(o, objId, { shapeXml, xfrm, hyperlinkXml, nextId: () => ++objId,
            embedImage: im => embedPictureSource(pictureMediaSource(im)), embedSource: embedPictureSource });
          parts.push(anchor(o, groupXml));
        } else {
          objId++;
          parts.push(anchor(o, shapeXml(o, objId, xfrm, hyperlinkXml)));
        }
      }
      files[`xl/drawings/drawing${drawingNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">${parts.join('')}</xdr:wsDr>`;
      if (drawingRels.length) files[`xl/drawings/_rels/drawing${drawingNo}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${drawingRels.join('')}</Relationships>`;
      contentOverrides.push(`<Override PartName="/xl/drawings/drawing${drawingNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
      drawing = `<drawing r:id="${addRel('drawing', `../drawings/drawing${drawingNo}.xml`)}"/>`;
    }

    // 데이터 유효성 검사
    const dvList = (sheet.validations ?? []).map(fit).filter(Boolean);
    const dataValidations = dvList.length ? `<dataValidations count="${dvList.length}">${dvList.map(validationXml).join('')}</dataValidations>` : '';
    const linkXml = [];
    for (const [r,c,cell] of storedCellEntries(sheet.cells)) {
      if (!cell.link) continue;
      if (r >= EXCEL_MAX_ROWS) continue;
      if (cell.link.startsWith('#')) linkXml.push(`<hyperlink ref="${cellName(r, c)}" location="${esc(cell.link.slice(1))}" display="${esc(String(wb.getValue(si, r, c) ?? ''))}"/>`);
      else {
        const id = `rId${sheetRels.length + 1}`;
        sheetRels.push(`<Relationship Id="${id}" Type="${REL}/hyperlink" Target="${esc(cell.link)}" TargetMode="External"/>`);
        linkXml.push(`<hyperlink ref="${cellName(r, c)}" r:id="${id}"/>`);
      }
    }
    const hyperlinks = linkXml.length ? `<hyperlinks>${linkXml.join('')}</hyperlinks>` : '';

    // 메모
    let legacy = '';
    const comments = [];
    for (const [r,c,cell] of storedCellEntries(sheet.cells)) if (cell.comment) comments.push([[r,c],cell.comment]);
    if (comments.length) {
      commentNo++;
      files[`xl/comments${commentNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${NS_MAIN}"><authors><author>WIXEL</author></authors><commentList>${comments.map(([[r, c], text]) => `<comment ref="${cellName(r, c)}" authorId="0"><text><r><t xml:space="preserve">${xesc(text)}</t></r></text></comment>`).join('')}</commentList></comments>`;
      files[`xl/drawings/vmlDrawing${commentNo}.vml`] = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${commentNo}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${comments.map(([[r, c]], i) => `<v:shape id="_x0000_s${commentNo * 1024 + i + 1}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:59pt;z-index:${i + 1};visibility:${noteVisible(sheet.noteVisibility, r, c) ? 'visible' : 'hidden'}" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 10, ${c + 3}, 15, ${r + 3}, 4</x:Anchor>${noteVisible(sheet.noteVisibility, r, c) ? '<x:Visible/>' : ''}<x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`).join('')}</xml>`;
      addRel('comments', `../comments${commentNo}.xml`);
      legacy = `<legacyDrawing r:id="${addRel('vmlDrawing', `../drawings/vmlDrawing${commentNo}.vml`)}"/>`;
      contentOverrides.push(`<Override PartName="/xl/comments${commentNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/>`);
    }

    // 표
    const tableRids = [];
    for (const t of (sheet.tables ?? []).map(fit).filter(Boolean)) {
      tableNo = tableIds.get(t.name.toLowerCase());
      const names = columnNames(wb, si, t);
      const cols = names.map((n, i) => {
        const c = t.c1 + i;
        const fn = t.totalsFns?.[c];
        let extra = '';
        let inner = '';
        const tot = t.totals ? wb.getCell(si, t.r2, c) : t.totalsCells?.[c];
        const textInput = tot?.inputType === 'text' || (tot?.inputType !== 'value' && tot?.style?.numFmt === 'text');
        const isFormula = tot?.formula || (tot?.raw?.startsWith('=') && (!textInput || tot.fx));
        if (isFormula) {
          if (matchingTotalFunction(tot.raw, fn, t, n)) extra = ` totalsRowFunction="${fn}"`;
          else { extra = ' totalsRowFunction="custom"'; inner += `<totalsRowFormula>${esc(exportFormula(tot.raw, t.name))}</totalsRowFormula>`; }
        } else if (tot?.raw) extra = ` totalsRowLabel="${esc(tot.inputType === 'text' || tot.style?.numFmt === 'text' ? tot.raw : tot.raw.replace(/^'/, ''))}"`;
        else if (!t.totals && !Object.hasOwn(t.totalsCells ?? {}, c) && TOTAL_FUNCS.some(x => x.id === fn && x.code)) extra = ` totalsRowFunction="${fn}"`;
        else if (!t.totals && Object.hasOwn(t.totalsCells ?? {}, c)) extra = ' totalsRowLabel=""';
        if (tot?.style && Object.keys(tot.style).length) extra += ` totalsRowDxfId="${pool.objectDxf({style:tot.style})}"`;
        // 계산된 열: 데이터 행이 모두 같은 수식이면 새 행에도 자동으로 채워지도록
        const top = dataTop(t);
        const bottom = dataBottom(t);
        const first = wb.getCell(si, top, c);
        if (first?.formula && bottom >= top && first.raw.includes('[')) {
          let same = true;
          for (let r = top + 1; r <= bottom && same; r++) same = wb.getCell(si, r, c)?.raw === first.raw;
          if (same) inner = `<calculatedColumnFormula>${esc(exportFormula(first.raw, t.name))}</calculatedColumnFormula>${inner}`;
        }
        return inner ? `<tableColumn id="${i + 1}" name="${esc(n)}"${extra}>${inner}</tableColumn>` : `<tableColumn id="${i + 1}" name="${esc(n)}"${extra}/>`;
      }).join('');
      const filterRange = { r1: t.r1, c1: t.c1, r2: dataBottom(t), c2: t.c2 };
      const filterOptions = { dxf: style => pool.objectDxf({style}), includeSort: false, header: t.header };
      const af = t.filter && t.header ? autoFilterXml(t.filter, filterRange, filterOptions) : '';
      const tableSort = filterSortXml(t.filter?.sort ?? t.sort, filterRange, filterOptions);
      const style = t.style && t.style !== 'None' ? t.style : null;
      if (style && Array.isArray(t.styleElements)) pool.objectTableStyle({name:style,table:true,pivot:false,elements:t.styleElements});
      else if (style && isModernStyle(style)) pool.presetTableStyle(style, false);
      files[`xl/tables/table${tableNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${NS_MAIN}" id="${tableNo}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${rangeRef(t)}"${t.header ? '' : ' headerRowCount="0"'}${t.totals ? ' totalsRowCount="1"' : ` totalsRowShown="${t.totalsCells != null || Object.keys(t.totalsFns ?? {}).length ? 1 : 0}"`}>${af}${tableSort}<tableColumns count="${names.length}">${cols}</tableColumns><tableStyleInfo${style ? ` name="${esc(style)}"` : ''} showFirstColumn="${t.firstCol ? 1 : 0}" showLastColumn="${t.lastCol ? 1 : 0}" showRowStripes="${t.banded !== false ? 1 : 0}" showColumnStripes="${t.bandedCols ? 1 : 0}"/></table>`;
      contentOverrides.push(`<Override PartName="/xl/tables/table${tableNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
      tableRids.push(addRel('table', `../tables/table${tableNo}.xml`));
    }
    const tableParts = tableRids.length ? `<tableParts count="${tableRids.length}">${tableRids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>` : '';

    // 피벗 테이블
    for (const pinfo of pivotList.get(si) ?? []) {
      const n = pinfo.tableNo;
      const c = pinfo.cacheId;
      files[`xl/pivotTables/pivotTable${n}.xml`] = pinfo.parts.tableXml;
      files[`xl/pivotTables/_rels/pivotTable${n}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${c}.xml"/></Relationships>`;
      contentOverrides.push(`<Override PartName="/xl/pivotTables/pivotTable${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`);
      if (!pivotCaches.some((p) => p.cacheId === c)) {
        files[`xl/pivotCache/pivotCacheDefinition${c}.xml`] = pinfo.cache.cacheXml;
        contentOverrides.push(`<Override PartName="/xl/pivotCache/pivotCacheDefinition${c}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`);
        if (pinfo.cache.saveData) {
          files[`xl/pivotCache/pivotCacheRecords${c}.xml`] = savedPivotRecords(pinfo.cache, wb.date1904, streaming);
          contentOverrides.push(`<Override PartName="/xl/pivotCache/pivotCacheRecords${c}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>`);
        }
        const cacheRelationships = [];
        if (pinfo.cache.saveData) cacheRelationships.push(`<Relationship Id="rCacheRecords" Type="${REL}/pivotCacheRecords" Target="pivotCacheRecords${c}.xml"/>`);
        if (pinfo.cache.externalSource) cacheRelationships.push(`<Relationship Id="rExternalSource" Type="${REL}/externalLinkPath" Target="${esc(pinfo.cache.externalSource)}" TargetMode="External"/>`);
        if (cacheRelationships.length) files[`xl/pivotCache/_rels/pivotCacheDefinition${c}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${NS_PKG}">${cacheRelationships.join('')}</Relationships>`;
        pivotCaches.push({ cacheId: c, target: `pivotCache/pivotCacheDefinition${c}.xml` });
      }
      addRel('pivotTable', `../pivotTables/pivotTable${n}.xml`);
    }

    // 슬라이서 목록 (피벗용 x14, 표용 x15)
    const exts = [];
    for (const kind of ['pivot', 'table']) {
      const list = sheetSlicers[kind];
      if (!list.length) continue;
      slicerPartNo++;
      files[`xl/slicers/slicer${slicerPartNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">${list.map(({ sl, name, cache }) => `<slicer name="${esc(name)}" cache="${esc(cache)}" caption="${esc(sl.caption ?? name)}"${slicerStartItem(sl.startItem) ? ` startItem="${slicerStartItem(sl.startItem)}"` : ''}${(sl.columns ?? 1) > 1 ? ` columnCount="${sl.columns}"` : ''}${((st) => (st !== 'SlicerStyleLight1' ? ` style="${esc(st)}"` : ''))(pool.slicerStyleFor(sl))}${sl.showHeader === false ? ' showCaption="0"' : ''}${sl.noMove ? ' lockedPosition="1"' : ''} rowHeight="${Math.round((sl.buttonHeight ?? SLICER_DEFAULT_BUTTON_HEIGHT) * EMU)}"/>`).join('')}</slicers>`;
      contentOverrides.push(`<Override PartName="/xl/slicers/slicer${slicerPartNo}.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>`);
      const id = `rId${sheetRels.length + 1}`;
      sheetRels.push(`<Relationship Id="${id}" Type="${REL_MS}/slicer" Target="../slicers/slicer${slicerPartNo}.xml"/>`);
      exts.push(kind === 'pivot'
        ? `<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${NS_X14}"><x14:slicerList><x14:slicer r:id="${id}"/></x14:slicerList></ext>`
        : `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="${NS_X15}"><x14:slicerList xmlns:x14="${NS_X14}"><x14:slicer r:id="${id}"/></x14:slicerList></ext>`);
    }
    // 스파크라인 (x14:sparklineGroups)
    if (sheet.sparklines?.length) {
      const c = (tag, col) => `<x14:${tag} rgb="${argb(col ?? '#d00000')}"/>`;
      const groups = sheet.sparklines.filter((g) => g.items?.length).map((g) => {
        const attrs = [g.type !== 'line' ? `type="${g.type === 'winloss' ? 'stacked' : 'column'}"` : '', 'displayEmptyCellsAs="gap"',
          g.weight && g.weight !== 0.75 ? `lineWeight="${g.weight}"` : '',
          ...['markers', 'high', 'low', 'first', 'last', 'negative'].filter((k) => g[k]).map((k) => `${k}="1"`)].filter(Boolean).join(' ');
        const items = g.items.map((it) => `<x14:sparkline><xm:f>${esc(it.ref.includes('!') ? it.ref : `${quoteSheetName(sheet.name)}!${it.ref}`)}</xm:f><xm:sqref>${refOf(it.r, it.c)}</xm:sqref></x14:sparkline>`).join('');
        return `<x14:sparklineGroup ${attrs}>${c('colorSeries', g.color)}${c('colorNegative', g.negColor)}${c('colorAxis', '#000000')}${c('colorMarkers', g.markerColor)}${c('colorFirst', g.firstColor)}${c('colorLast', g.lastColor)}${c('colorHigh', g.highColor)}${c('colorLow', g.lowColor)}<x14:sparklines>${items}</x14:sparklines></x14:sparklineGroup>`;
      });
      if (groups.length) exts.push(`<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${NS_X14}"><x14:sparklineGroups xmlns:xm="http://schemas.microsoft.com/office/excel/2006/main">${groups.join('')}</x14:sparklineGroups></ext>`);
    }
    if (cfX14.length) exts.unshift(`<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${NS_X14}"><x14:conditionalFormattings>${cfX14.join('')}</x14:conditionalFormattings></ext>`);
    if (/^#[0-9a-f]{6}$/i.test(sheet.view?.gridColor ?? '') && INDEXED[gridIndex(sheet.view.gridColor)] !== sheet.view.gridColor.slice(1).toUpperCase()) exts.push(`<ext uri="{82745B15-A53C-4F1D-A901-574958454C56}"><wx:gridColor xmlns:wx="https://wixel.app/view/1" rgb="${sheet.view.gridColor.toLowerCase()}" nativeId="${gridIndex(sheet.view.gridColor)}"/></ext>`);
    const extLst = exts.length ? `<extLst>${exts.join('')}</extLst>` : '';

    const tabOk = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(String(sheet.tabColor ?? ''));
    files[`xl/worksheets/sheet${si + 1}.xml`] = [`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`
      + (vba || olPr || tabOk ? `<sheetPr${vba ? ` codeName="${esc(vba.sheetCodes?.[sheet.name] ?? `Sheet${si + 1}`)}"` : ''}>${tabOk ? `<tabColor rgb="${argb(sheet.tabColor)}"/>` : ''}${olPr}</sheetPr>` : '')
      + `<dimension ref="${dim}"/>`
      + `<sheetViews><sheetView${['normal', 'pageBreakPreview', 'pageLayout'].includes(sheet.view?.mode) ? ` view="${sheet.view.mode}"` : ''}${sheet.noGrid ? ' showGridLines="0"' : ''}${sheet.noZeros ? ' showZeros="0"' : ''}${typeof sheet.view?.headers === 'boolean' ? ` showRowColHeaders="${sheet.view.headers ? 1 : 0}"` : ''}${typeof sheet.view?.showFormulas === 'boolean' ? ` showFormulas="${sheet.view.showFormulas ? 1 : 0}"` : ''}${/^#[0-9a-f]{6}$/i.test(sheet.view?.gridColor ?? '') ? ` defaultGridColor="0" colorId="${gridIndex(sheet.view.gridColor)}"` : ''}${sheet.zoom && sheet.zoom !== 100 ? ` zoomScale="${sheet.zoom}" zoomScaleNormal="${sheet.zoom}"` : ''}${viewTopLeft !== 'A1' ? ` topLeftCell="${viewTopLeft}"` : ''} workbookViewId="0"${si === (wb.sheets[activeSheet]?.state && wb.sheets[activeSheet].state !== 'visible' ? Math.max(0, wb.sheets.findIndex((x) => !x.state || x.state === 'visible')) : activeSheet) ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>`
      + `<sheetFormatPr defaultColWidth="${px2widthM(sheet.defColW ?? DEFAULT_COL_WIDTH, wmdw)}" defaultRowHeight="${px2pt(sheet.defRowH ?? DEFAULT_ROW_HEIGHT)}" customHeight="1"${olRowMax ? ` outlineLevelRow="${olRowMax}"` : ''}${olColMax ? ` outlineLevelCol="${olColMax}"` : ''}/>`
      + (colsXml ? `<cols>${colsXml}</cols>` : '')
      + '<sheetData>', rowXml, '</sheetData>'
      + protectXml(sheet.protect)
      + (sheet.protectedRanges?.length ? `<protectedRanges>${sheet.protectedRanges.filter(p => p.ranges?.length).map(p => `<protectedRange name="${esc(p.name)}" sqref="${esc(p.ranges.map(rangeRef).join(' '))}"${p.hash ? ` password="${esc(p.hash)}"` : ''}${p.securityDescriptor ? ` securityDescriptor="${esc(p.securityDescriptor)}"` : ''}${p.modern ? Object.entries(p.modern).filter(([,v]) => v !== undefined).map(([k,v]) => ` ${k}="${esc(v)}"`).join('') : ''}/>`).join('')}</protectedRanges>` : '')
      + scenariosXml(sheet.scenarios)
      + autoFilter + merges + cf + dataValidations + hyperlinks
      + pgx.printOptions + pgx.margins + pgx.setup + pgx.headerFooter + pageBreakXml(sheet.page)
      + drawing + legacy + tableParts + extLst
      + '</worksheet>'];
    if (sheetRels.length) {
      files[`xl/worksheets/_rels/sheet${si + 1}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${sheetRels.join('')}</Relationships>`;
    }
    if (streaming) for (const name of Object.keys(files)) {
      yield { entry: [name, files[name]] }; delete files[name];
    }
    yield { p: streaming ? Math.min(0.85, 0.85 * rowsDone / rowsTotal) : 0.85, msg: '파일 구성 중' };
  }

  // 통합 문서: 피벗 캐시, 슬라이서 캐시 (확장)
  const wbRels = [
    ...wb.sheets.slice(0, nOwn).map((sh, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`),
    `<Relationship Id="rId${nOwn + 1}" Type="${REL}/styles" Target="styles.xml"/>`,
    `<Relationship Id="rId${nOwn + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`,
    `<Relationship Id="rId${nOwn + 3}" Type="${REL}/theme" Target="theme/theme1.xml"/>`,
  ];
  if (vba) wbRels.push(`<Relationship Id="rId${wbRels.length + 1}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`);
  if (dynamicCells || richList.length) {
    // 셀 메타데이터: 동적 배열 수식(cm="1" → XLDAPR), 셀 그림(vm="N" → XLRICHVALUE)
    wbRels.push(`<Relationship Id="rId${wbRels.length + 1}" Type="${REL}/sheetMetadata" Target="metadata.xml"/>`);
    const types = [];
    let future = '';
    if (dynamicCells) {
      types.push('<metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/>');
      future += '<futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata>';
    }
    if (richList.length) {
      types.push('<metadataType name="XLRICHVALUE" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1"/>');
      future += `<futureMetadata name="XLRICHVALUE" count="${richList.length}">${richList.map((_, i) => `<bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="${i}"/></ext></extLst></bk>`).join('')}</futureMetadata>`;
    }
    const cellMeta = dynamicCells ? '<cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata>' : '';
    const valueMeta = richList.length ? `<valueMetadata count="${richList.length}">${richList.map((_, i) => `<bk><rc t="${types.length}" v="${i}"/></bk>`).join('')}</valueMetadata>` : '';
    files['xl/metadata.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="${types.length}">${types.join('')}</metadataTypes>${future}${cellMeta}${valueMeta}</metadata>`;
    contentOverrides.push('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>');
  }
  if (richList.length) {
    const RD = 'http://schemas.microsoft.com/office/spreadsheetml/2017/richdata';
    const withAlt = richList.some((x) => x.alt);
    files['xl/richData/rdrichvaluestructure.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<rvStructures xmlns="${RD}" count="${withAlt ? 2 : 1}"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s>${withAlt ? '<s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/><k n="Text" t="s"/></s>' : ''}</rvStructures>`;
    files['xl/richData/rdrichvalue.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<rvData xmlns="${RD}" count="${richList.length}">${richList.map((x) => (x.alt ? `<rv s="1"><v>${x.rel}</v><v>5</v><v>${esc(x.alt)}</v></rv>` : `<rv s="0"><v>${x.rel}</v><v>5</v></rv>`)).join('')}</rvData>`;
    const media = [...richMedia.values()];
    files['xl/richData/richValueRel.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="${NS_R}">${media.map((m) => `<rel r:id="rId${m.i + 1}"/>`).join('')}</richValueRels>`;
    files['xl/richData/_rels/richValueRel.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${media.map((m) => `<Relationship Id="rId${m.i + 1}" Type="${REL}/image" Target="${m.path}"/>`).join('')}</Relationships>`;
    files['xl/richData/rdRichValueTypes.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<rvTypesInfo xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x" xmlns:x="${NS_MAIN}"><global><keyFlags><key name="_Self"><flag name="ExcludeFromFile" value="1"/><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_DisplayString"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Flags"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Format"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_SubLabel"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Attribution"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Icon"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_Display"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_CanonicalPropertyNames"><flag name="ExcludeFromCalcComparison" value="1"/></key><key name="_ClassificationId"><flag name="ExcludeFromCalcComparison" value="1"/></key></keyFlags></global></rvTypesInfo>`;
    const add = (type, target) => wbRels.push(`<Relationship Id="rId${wbRels.length + 1}" Type="${type}" Target="${target}"/>`);
    add('http://schemas.microsoft.com/office/2022/10/relationships/richValueRel', 'richData/richValueRel.xml');
    add('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue', 'richData/rdrichvalue.xml');
    add('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure', 'richData/rdrichvaluestructure.xml');
    add('http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueTypes', 'richData/rdRichValueTypes.xml');
    contentOverrides.push(
      '<Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>',
      '<Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>',
      '<Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>',
      '<Override PartName="/xl/richData/rdRichValueTypes.xml" ContentType="application/vnd.ms-excel.rdrichvaluetypes+xml"/>',
    );
  }
  // 사용자 이름 정의
  for (const n of wb.names ?? []) {
    const local = n.sheet ? wb.sheetIndexByName(n.sheet) : -1;
    const attrs = [`name="${esc(n.name)}"`];
    if (local >= 0) attrs.push(`localSheetId="${local}"`);
    if (n.hidden) attrs.push('hidden="1"');
    if (n.comment) attrs.push(`comment="${esc(n.comment)}"`);
    definedNames.push(`<definedName ${attrs.join(' ')}>${esc(exportFormula(String(n.ref).startsWith('=') ? n.ref : `=${n.ref}`))}</definedName>`);
  }
  const wbRel = (type, target) => { const id = `rId${wbRels.length + 1}`; wbRels.push(`<Relationship Id="${id}" Type="${type}" Target="${target}"/>`); return id; };
  const pivotCachesXml = pivotCaches.length ? `<pivotCaches>${pivotCaches.map((p) => `<pivotCache cacheId="${p.cacheId}" r:id="${wbRel(`${REL}/pivotCacheDefinition`, p.target)}"/>`).join('')}</pivotCaches>` : '';
  const wbExts = [];
  if (slicerCachesPivot.length) wbExts.push(`<ext uri="{BBE1A952-AA13-448e-AADC-164F8A28A991}" xmlns:x14="${NS_X14}"><x14:slicerCaches>${slicerCachesPivot.map((t) => `<x14:slicerCache r:id="${wbRel(`${REL_MS}/slicerCache`, t)}"/>`).join('')}</x14:slicerCaches></ext>`);
  if (slicerCachesTable.length) wbExts.push(`<ext uri="{46BE6895-7355-4a93-B00E-2C351335B9C9}" xmlns:x15="${NS_X15}"><x15:slicerCaches xmlns:x14="${NS_X14}">${slicerCachesTable.map((t) => `<x14:slicerCache r:id="${wbRel(`${REL_MS}/slicerCache`, t)}"/>`).join('')}</x15:slicerCaches></ext>`);
  // 외부 통합 문서 연결: 읽을 때 보관한 externalLink 원본을 그대로 (수식의 [N] 순서 유지)
  let extRefsXml = '';
  const exts = wb.externals ?? [];
  if (exts.length) {
    const ids = exts.map((e, i) => {
      const n = i + 1;
      files[`xl/externalLinks/externalLink${n}.xml`] = e.xml;
      if (e.rels) files[`xl/externalLinks/_rels/externalLink${n}.xml.rels`] = e.rels;
      contentOverrides.push(`<Override PartName="/xl/externalLinks/externalLink${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.externalLink+xml"/>`);
      return wbRel(`${REL}/externalLink`, `externalLinks/externalLink${n}.xml`);
    });
    extRefsXml = `<externalReferences>${ids.map((id) => `<externalReference r:id="${id}"/>`).join('')}</externalReferences>`;
  }
  // 숨긴 시트는 활성 시트가 될 수 없음
  const isShown = (i) => wb.sheets[i] && wb.sheets[i].state !== 'hidden' && wb.sheets[i].state !== 'veryHidden';
  const firstVisible = Math.max(0, wb.sheets.findIndex((_, i) => isShown(i)));
  const activeTab = isShown(activeSheet) ? activeSheet : firstVisible;
  files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${wb.props?.readOnlyRecommended ? '<fileSharing readOnlyRecommended="1"/>' : ''}${vba || wb.date1904 ? `<workbookPr${vba ? ` codeName="${esc(vba.codeName || 'ThisWorkbook')}"` : ''}${wb.date1904 ? ' date1904="1"' : ''}/>` : ''}${workbookProtectionXml(wb.props)}<bookViews><workbookView${firstVisible ? ` firstSheet="${firstVisible}"` : ''} activeTab="${activeTab}"/></bookViews><sheets>${wb.sheets.slice(0, nOwn).map((sh, i) => `<sheet name="${esc(sh.name)}" sheetId="${i + 1}"${sh.state === 'hidden' || sh.state === 'veryHidden' ? ` state="${sh.state}"` : ''} r:id="rId${i + 1}"/>`).join('')}</sheets>${extRefsXml}${definedNames.length ? `<definedNames>${definedNames.join('')}</definedNames>` : ''}${calculationXml(wb.calculation)}${pivotCachesXml}${wbExts.length ? `<extLst>${wbExts.join('')}</extLst>` : ''}</workbook>`;
  if (pool.hasCheckbox) {
    files['xl/featurePropertyBag/featurePropertyBag.xml'] = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<FeaturePropertyBags xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/featurepropertybag"><bag type="Checkbox"/><bag type="XFControls"><bagId k="CellControl">0</bagId></bag><bag type="XFComplement"><bagId k="XFControls">1</bagId></bag><bag type="XFComplements" extRef="XFComplementsMapperExtRef"><a k="MappedFeaturePropertyBags"><bagId>2</bagId></a></bag></FeaturePropertyBags>';
    wbRel('http://schemas.microsoft.com/office/2022/11/relationships/FeaturePropertyBag', 'featurePropertyBag/featurePropertyBag.xml');
    contentOverrides.push('<Override PartName="/xl/featurePropertyBag/featurePropertyBag.xml" ContentType="application/vnd.ms-excel.featurepropertybag+xml"/>');
  }
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${wbRels.join('')}</Relationships>`;
  if (vba) files['xl/vbaProject.bin'] = fromBase64(vba.bin);
  function* sharedParts() {
    yield `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${strings.length}" uniqueCount="${strings.length}">`;
    for (const value of strings) yield `<si><t xml:space="preserve">${xesc(value)}</t></si>`;
    yield '</sst>';
  }
  if (streaming) files['xl/sharedStrings.xml'] = sharedParts();
  else {
    const sharedXml = createXmlChunks(); for (const text of sharedParts()) sharedXml.push(text);
    files['xl/sharedStrings.xml'] = sharedXml.finish();
  }
  files['xl/styles.xml'] = pool.xml();
  files['xl/theme/theme1.xml'] = themeXml(wb);
  contentOverrides.push('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const pr = wb.props ?? {};
  const tagX = (tag, v) => (v ? `<${tag}>${esc(String(v))}</${tag}>` : '');
  files['docProps/core.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${tagX('dc:title', pr.title)}${tagX('dc:subject', pr.subject)}${tagX('dc:creator', pr.creator || 'WIXEL')}${tagX('cp:keywords', pr.tags)}${tagX('dc:description', pr.comments)}${tagX('cp:lastModifiedBy', pr.lastModifiedBy)}${tagX('cp:category', pr.category)}<dcterms:created xsi:type="dcterms:W3CDTF">${/^\d{4}-\d\d-\d\dT/.test(pr.created ?? '') ? esc(pr.created) : now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  const importCodes = importWarningCodes(pr);
  if (pr.markedFinal || importCodes.length) {
    const customProps = (pr.markedFinal ? '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true</vt:bool></property>' : '')
      + (importCodes.length ? `<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="3" name="_WixelXlsxImportWarnings"><vt:lpwstr>${esc(JSON.stringify(importCodes))}</vt:lpwstr></property>` : '');
    files['docProps/custom.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">${customProps}</Properties>`;
    files['_rels/.rels'] = files['_rels/.rels'].replace('</Relationships>', `<Relationship Id="rId4" Type="${REL}/custom-properties" Target="docProps/custom.xml"/></Relationships>`);
    contentOverrides.push('<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>');
  }
  files['docProps/app.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>WIXEL</Application></Properties>`;
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>${[...mediaExts].map((e) => `<Default Extension="${e}" ContentType="${MIME[e]}"/>`).join('')}${vba ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ''}<Override PartName="/xl/workbook.xml" ContentType="${MAIN_TYPES[kind ?? (vba ? 'xlsm' : 'xlsx')] ?? MAIN_TYPES.xlsx}"/>${wb.sheets.slice(0, nOwn).map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${contentOverrides.join('')}</Types>`;

  // [Content_Types].xml 을 맨 앞에 두는 것이 관례
  const ordered = { '[Content_Types].xml': files['[Content_Types].xml'] };
  for (const [k, v] of Object.entries(files)) if (k !== '[Content_Types].xml') ordered[k] = v;
  return ordered;
}
