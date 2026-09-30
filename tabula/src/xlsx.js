// .xlsx 읽기/쓰기 (Office Open XML). DOM 없이 동작하므로 Node 에서도 테스트 가능.
import { isXlsb, convertXlsb } from './xlsb.js';
import { unzip, unzipAsync, zip, zipAsync, textOf } from './zip.js';
import { CellMap } from './cellmap.js';
import { protectFromAttrs, protectXml } from './protect.js';
import { pageXml, pageFromXml, normPage } from './page.js';
import { parseXml, child, kids, descendants, allText, esc, decodeEntities, unx } from './xml.js';
import {
  parse, tokenize, colToName, nameToCol, cellName, parseRangeName, FUNCS, isError,
  quoteSheetName, MAX_ROWS, MAX_COLS, EXCEL_MAX_ROWS, mayReturnArray, unknownFunctions,
} from './formula.js';
import { toFileFormula, fromFileFormula } from './xlfn.js';
import { parseInput, formatGeneral, fmtCode, styleForCode, dateParts, serialOf } from './format.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT, formulaShifter } from './workbook.js';
import { chartLayout, PALETTE, chartModelData, paletteOf } from './chart.js';
import { Axis, hid, hidKeys } from './axis.js';
import { toBase64, fromBase64 } from './vba.js';
import { CellImage } from './fxcore.js';
import { emfDataUrl } from './emf.js';
import { GEOM, LINE_KINDS } from './shapes.js';
import { BLOCK_MIN_ROWS, ColBuilder, inBlock, blockValue } from './block.js';
import { normalizeStyleName, DEFAULT_TABLE_STYLE, dataTop, dataBottom, canonicalRef, tableAt, columnNames, findTable } from './tables.js';
import { pivotSourceData, resolvePivot, itemText, keyOf, sortKeys, EMPTY, headerNames, normalizeDef, computePivot, valueName, showAsPercent, excelCalcFormula, pivotFilterKey, DATE_OP_TYPES } from './pivot.js';
import { slicerStyleName, slicerColors, isModernSlicer } from './slicerstyle.js';
import { applyTint, DEFAULT_THEME, PRESET_STYLES, presetStyle, isModernStyle, ELEMENT_TYPES, elementDxfStyle } from './stylepresets.js';
import { maxOf, minOf, pushAll, DAY_MS } from './fxcore.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU = 9525; // 1px

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
  else if (a.indexed !== undefined) hex = INDEXED[Number(a.indexed)] ?? null;
  if (!hex) return null;
  hex = applyTint(hex.toUpperCase(), Number(a.tint || 0));
  return `#${hex.toLowerCase()}`;
}

const argb = (color) => `FF${String(color).replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6)}`;

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
      out[r.attrs.Id] = { target, rawTarget: target, type: r.attrs.Type.split('/').pop(), external: true };
      continue;
    }
    if (target.startsWith('/')) target = target.slice(1);
    else {
      const parts = (dir + target).split('/');
      const res = [];
      for (const p of parts) { if (p === '..') res.pop(); else if (p !== '.') res.push(p); }
      target = res.join('/');
    }
    out[r.attrs.Id] = { target, type: r.attrs.Type.split('/').pop() };
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

function readStyles(files, wbRels, theme) {
  const rel = Object.values(wbRels).find((r) => r.type === 'styles');
  const xml = rel && textOf(files[rel.target]);
  if (!xml) return { xfs: [], dxfs: [], defaultFont: null };
  const root = parseXml(xml);
  const numFmts = {};
  for (const f of kids(child(root, 'numFmts'), 'numFmt')) numFmts[f.attrs.numFmtId] = f.attrs.formatCode;
  const fontOf = (f) => {
    const st = {};
    if (child(f, 'b') && child(f, 'b').attrs.val !== '0') st.bold = true;
    if (child(f, 'i') && child(f, 'i').attrs.val !== '0') st.italic = true;
    if (child(f, 'u') && child(f, 'u').attrs.val !== 'none') st.underline = true;
    if (child(f, 'strike') && child(f, 'strike').attrs.val !== '0') st.strike = true;
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
      const stop = descendants(f, 'stop')[0];
      return stop ? colorOf(child(stop, 'color'), theme) : null;
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
    const st = { ...fonts[Number(a.fontId || 0)] };
    if (st.font && st.font === defaultFont) delete st.font;
    if (st.size === defaultSize) delete st.size; // 통합 문서 기본 크기는 적지 않음 (기본 글꼴로 표시)
    const fill = fills[Number(a.fillId || 0)];
    if (fill && typeof fill === 'object') Object.assign(st, fill);
    else if (fill) st.fill = fill;
    Object.assign(st, borders[Number(a.borderId || 0)] ?? {});
    Object.assign(st, numFmtOf(a.numFmtId || 0));
    const al = child(xf, 'alignment') ?? (a.applyAlignment === '1' ? null : child(parent, 'alignment'));
    if (al) {
      const h = al.attrs.horizontal;
      if (h === 'left' || h === 'center' || h === 'right') st.align = h;
      else if (h === 'centerContinuous' || h === 'distributed') st.align = 'center';
      else if (h === 'justify') { st.align = 'left'; st.wrap = true; }
      const v = al.attrs.vertical;
      if (v === 'top') st.valign = 'top';
      else if (v === 'center' || v === 'justify' || v === 'distributed') st.valign = 'middle';
      if (al.attrs.wrapText === '1' || al.attrs.wrapText === 'true') st.wrap = true;
      if (Number(al.attrs.indent)) st.indent = Number(al.attrs.indent);
      const rot = Number(al.attrs.textRotation ?? 0);
      if (rot) st.rotate = rot === 255 ? 255 : rot > 90 ? -(rot - 90) : rot; // 255 = 세로 쓰기
      if (al.attrs.shrinkToFit === '1' || al.attrs.shrinkToFit === 'true') st.shrink = true;
    }
    // 셀 보호: 잠금 해제 · 수식 숨기기
    const pr = child(xf, 'protection');
    if (pr) {
      if (pr.attrs.locked === '0' || pr.attrs.locked === 'false') st.locked = false;
      if (pr.attrs.hidden === '1' || pr.attrs.hidden === 'true') st.hideFormula = true;
    }
    return st;
  };
  const xfs = kids(child(root, 'cellXfs'), 'xf').map((xf) => xfStyle(xf, styleXfs[Number(xf.attrs.xfId ?? 0)]));
  // 이름 있는 셀 스타일 (엑셀 [셀 스타일]의 사용자 지정 · 기본 제공 스타일을 파일에서 고친 것) — '표준' 은 제외
  const cellStyles = kids(child(root, 'cellStyles'), 'cellStyle')
    .filter((c) => c.attrs.builtinId !== '0' && c.attrs.hidden !== '1' && styleXfs[Number(c.attrs.xfId)])
    .map((c) => ({ name: unx(c.attrs.name ?? ''), style: xfStyle(styleXfs[Number(c.attrs.xfId)], null), ...(c.attrs.builtinId !== undefined ? { builtinId: Number(c.attrs.builtinId) } : {}) }))
    .filter((c) => c.name);
  const dxfOf = (d) => {
    const st = {};
    const f = child(d, 'font');
    if (f) {
      const c = colorOf(child(f, 'color'), theme);
      if (c) st.color = c;
      if (child(f, 'b') && child(f, 'b').attrs.val !== '0') st.bold = true;
      if (child(f, 'i') && child(f, 'i').attrs.val !== '0') st.italic = true;
      if (child(f, 'strike') && child(f, 'strike').attrs.val !== '0') st.strike = true;
      if (child(f, 'u') && child(f, 'u').attrs.val !== 'none') st.underline = true;
    }
    const fill = child(d, 'fill');
    if (fill) { const c = fillOf(fill, true); if (c) st.fill = c; }
    const bd = child(d, 'border');
    if (bd) Object.assign(st, borderOf(bd));
    const nf = child(d, 'numFmt');
    if (nf?.attrs.formatCode) Object.assign(st, styleForCode(nf.attrs.formatCode));
    return st;
  };
  const dxfs = kids(child(root, 'dxfs'), 'dxf').map(dxfOf);
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
  return { xfs, dxfs, dxfOf, defaultFont, tableStyles, wbFont, slicerStyles, cellStyles };
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
        if (f) { cell.f = f.text; cell.fa = f.attrs; }
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

function* readSheet(files, path, ctx) {
  // 셀 데이터(sheetData)는 빠른 전용 스캐너로, 나머지는 일반 XML 파서로 읽음
  const binRows = files.__xlsb?.rows.get(path); // xlsb: 셀은 바이너리에서 바로
  // 아주 큰 시트(수백 MB XML)는 한 번에 글자로 바꾸지 않고 행 경계에서 잘라 조금씩 (시간 · 메모리)
  const big = !binRows ? chunkedSheet(files[path]) : null;
  const xmlText = big || binRows ? null : textOf(files[path]);
  const sd = binRows || big ? null : splitSheetData(xmlText);
  const root = parseXml(big ? big.rest : binRows ? textOf(files[path]) : sd ? sd.rest : xmlText);
  const sheetRows = binRows ? binRows() : big ? big.rows() : sd ? scanRows(xmlText, sd.start, sd.end) : domRows(child(root, 'sheetData'));
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
  const NONE = {};
  const neutral = (k, v) => (k === 'locked' ? true : k === 'size' ? ctx.wbFont.size : k === 'font' ? ctx.wbFont.name : typeof v === 'boolean' ? false : typeof v === 'number' ? 0 : '');
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
        }
      }
      m2.set(col ?? NONE, out ?? null);
    }
    return out ?? undefined;
  };

  for (const col of kids(child(root, 'cols'), 'col')) {
    const min = Number(col.attrs.min) - 1;
    const max = Math.min(Number(col.attrs.max) - 1, min + 16384);
    const w = col.attrs.width !== undefined ? width2pxM(Number(col.attrs.width), mdw) : null;
    const st = col.attrs.style ? styleOf(col.attrs.style) : undefined;
    for (let c = min; c <= max && c < MAX_COLS; c++) {
      if (w !== null && w !== defColW && (col.attrs.customWidth === '1' || Math.abs(w - defColW) > 1)) sheet.colWidths[c] = w;
      if (col.attrs.hidden === '1' || col.attrs.hidden === 'true') sheet.hiddenCols[c] = true;
      if (st && max - min < 1000) sheet.colStyles[c] = st;
      const ol = Number(col.attrs.outlineLevel ?? 0);
      if (ol > 0) ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).cols[c] = Math.min(7, ol));
      if (col.attrs.collapsed === '1' || col.attrs.collapsed === 'true') ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).colsColl[c] = true);
    }
  }

  const shared = {};
  const arrays = []; // 배열 수식 영역: 앵커 밖의 셀 값은 가져오지 않음 (다시 분산됨)
  let unsupported = 0;
  let rowIdx = -1;
  // 행이 아주 많은 시트: 값 셀은 열 블록(형식화 배열)으로 — 셀 객체 수백만 개를 만들지 않음
  const dimRef = refToRange(child(root, 'dimension')?.attrs.ref ?? '');
  const blockMode = !!dimRef && dimRef.r2 - dimRef.r1 + 1 > BLOCK_MIN_ROWS;
  let blockStart = -1;
  let blockLast = -1;
  let lateEmpty = null; // [행, 열, 서식] … 블록 열 서식과 같아 건너뛴 빈 칸
  const builders = [];
  const colFmt = [];
  const formulaMemo = ctx.formulaMemo ??= { legacy: new Map(), modern: new Map(), shape: new Map() }; // 같은 수식 문자열(표의 계산 열 등)은 한 번만 변환
  const textMemo = ctx.textMemo ??= new Map();
  let rowCount = 0;
  for (const row of sheetRows) {
    if (++rowCount % 1000 === 0) yield rowCount;
    rowIdx = row.attrs.r ? Number(row.attrs.r) - 1 : rowIdx + 1;
    const r = rowIdx;
    if (row.attrs.ht && (row.attrs.customHeight === '1' || row.attrs.customHeight === 'true')) {
      const h = pt2px(Number(row.attrs.ht));
      if (h !== defRowH) { sheet.rowHeights[r] = h; sheet.rowManual[r] = true; }
    } else if (row.attrs.ht) {
      const h = pt2px(Number(row.attrs.ht));
      if (h !== defRowH) sheet.rowHeights[r] = h;
    }
    const noHt = !row.attrs.ht; // 높이가 저장되지 않은 행: 엑셀은 내용(큰 글꼴 · 줄 바꿈 · 회전)에 맞춰 자동 높이
    if (row.attrs.hidden === '1' || row.attrs.hidden === 'true') sheet.hiddenRows[r] = true;
    // 개요 (행 그룹)
    if (row.attrs.outlineLevel && row.attrs.outlineLevel !== '0') ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).rows[r] = Math.min(7, Number(row.attrs.outlineLevel)));
    if (row.attrs.collapsed === '1' || row.attrs.collapsed === 'true') ((sheet.outline ??= { rows: {}, cols: {}, rowsColl: {}, colsColl: {}, below: true, right: true }).rowsColl[r] = true);
    if (row.attrs.customFormat === '1' && row.attrs.s) { const st = styleOf(row.attrs.s); if (st) sheet.rowStyles[r] = st; }
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
      colIdx = cc;
      const t = c.attrs.t ?? 'n';
      const vText = c.v;
      const style = ownStyle(styleOf(c.attrs.s), r, cc);
      let raw = '';
      let value = null;
      if (t === 's') value = strings[Number(vText)] ?? '';
      else if (t === 'inlineStr') value = c.is !== null ? allText(parseXml(`<is>${c.is}</is>`)) : '';
      else if (t === 'str') value = vText ?? '';
      else if (t === 'b') value = vText === '1';
      else if (t === 'e') value = { error: vText ?? '#N/A' };
      else if (vText !== null && vText !== '') value = Number(vText);

      let formula = null;
      const fa = c.fa;
      if (fa) {
        if (fa.t === 'shared' && fa.si !== undefined) {
          if (c.f) shared[fa.si] = { text: c.f, r, c: cc };
          const m = shared[fa.si];
          // 공유 수식: 기준 수식을 한 번만 나눠 두고 칸마다 행 · 열만 옮김
          if (m) formula = m.r === r && m.c === cc ? m.text : (m.shift ??= formulaShifter(`=${m.text}`))(r - m.r, cc - m.c).slice(1);
        } else if (c.f) formula = c.f;
      }
      let cached;
      if (formula !== null) {
        const isArray = fa.t === 'array';
        if (isArray && fa.ref) {
          const rg = refToRange(fa.ref);
          if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) arrays.push({ ...rg, r, c: cc });
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
        // 파일에 저장된 계산 결과: 열 때는 이 값을 그대로 씀 (지원하지 않는 함수는 계속 이 값을 표시)
        cached = value;
        if (conv.unknown) unsupported++;
      } else if (arrays.length && arrays.some((a) => r >= a.r1 && r <= a.r2 && cc >= a.c1 && cc <= a.c2 && (r !== a.r || cc !== a.c))) {
        value = null;
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
        } else raw = numberRaw(value, style);
      }
      if (noHt && raw !== '' && style && ((style.size && style.size > ctx.wbFont.size) || style.wrap || style.rotate)) (sheet.fitRows ??= new Set()).add(r);
      if (blockMode && blockStart < 0) blockStart = r + 1; // 첫 행(머리글) 다음부터 블록
      if (blockMode && r >= blockStart && formula === null && !c.attrs.vm) {
        if (colFmt[cc] === undefined) colFmt[cc] = style ?? null;
        if ((style ?? null) === colFmt[cc]) {
          if (value !== null && value !== '') {
            (builders[cc] ??= new ColBuilder(dimRef.r2 - blockStart + 1)).set(r - blockStart, value);
            if (r > blockLast) blockLast = r;
            continue;
          }
          // 빈 칸의 서식은 블록 열 서식으로 (블록 마지막 행보다 아래면 나중에 보통 셀로 되살림)
          if (value !== '') { if (style) (lateEmpty ??= []).push(r, cc, style); continue; } // 빈 글자 셀은 블록 대신 보통 셀로 (빈 칸과 구별)
        }
      }
      // 셀에 배치한 그림 (richData 값 메타데이터 vm)
      const cellImg = !formula && c.attrs.vm ? ctx.richImages?.[Number(c.attrs.vm)] : null;
      if (cellImg) raw = '';
      if (!raw && !style && !cellImg) continue;
      const d = { raw };
      if (cellImg) d.image = { ...cellImg };
      if (style) d.style = style;
      if (cached !== undefined && cached !== null) d.cached = cached;
      if (formula !== null && style?.numFmt === 'text') d.fx = true; // 텍스트 서식 칸에 저장된 수식
      sheet.cells.setRC(r, cc, d);
    }
  }
  sheet.unsupported = unsupported;
  if (lateEmpty) {
    const last = blockMode && blockLast >= blockStart ? blockLast : -1;
    for (let i = 0; i < lateEmpty.length; i += 3) if (lateEmpty[i] > last && !sheet.cells.getRC(lateEmpty[i], lateEmpty[i + 1])) sheet.cells.setRC(lateEmpty[i], lateEmpty[i + 1], { raw: '', style: lateEmpty[i + 2] });
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
    return v === null ? undefined : { raw: typeof v === 'number' ? numberRaw(v, b.cols[cc2].fmt) : typeof v === 'string' ? textRaw(v) : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : v.error, ...(b.cols[cc2].fmt ? { style: b.cols[cc2].fmt } : {}) };
  };

  for (const m of kids(child(root, 'mergeCells'), 'mergeCell')) {
    const rg = refToRange(m.attrs.ref);
    if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) sheet.merges.push(rg);
  }

  const sv0 = descendants(child(root, 'sheetViews'), 'sheetView')[0];
  if (sv0 && (sv0.attrs.showGridLines === '0' || sv0.attrs.showGridLines === 'false')) sheet.noGrid = true;
  // 확대/축소 · 처음 보이는 칸 · 활성 셀 (엑셀에서 저장한 화면 그대로 열기)
  if (sv0?.attrs.zoomScale && Number(sv0.attrs.zoomScale) !== 100) sheet.zoom = Math.max(10, Math.min(400, Number(sv0.attrs.zoomScale)));
  const tlc = sv0?.attrs.topLeftCell ? refToRange(sv0.attrs.topLeftCell) : null;
  const selEl = descendants(sv0, 'selection').find((x) => !x.attrs.pane || x.attrs.pane === 'bottomRight') ?? descendants(sv0, 'selection')[0];
  const act = selEl?.attrs.activeCell ? refToRange(selEl.attrs.activeCell) : null;
  if ((tlc && (tlc.r1 || tlc.c1)) || (act && (act.r1 || act.c1))) sheet.view = { top: tlc?.r1 ?? 0, left: tlc?.c1 ?? 0, ...(act ? { r: act.r1, c: act.c1 } : {}) };
  sheet.fileValues = true; // 셀의 파일 계산 결과를 그대로 씀 (바뀌기 전까지)
  const pane = descendants(child(root, 'sheetViews'), 'pane')[0];
  if (pane && (pane.attrs.state === 'frozen' || pane.attrs.state === 'frozenSplit')) {
    sheet.freeze = { rows: Number(pane.attrs.ySplit || 0), cols: Number(pane.attrs.xSplit || 0) };
  }

  const af = child(root, 'autoFilter');
  if (af?.attrs.ref) {
    const rg = refToRange(af.attrs.ref);
    if (rg) {
      const criteria = {};
      for (const fc of kids(af, 'filterColumn')) {
        const filters = child(fc, 'filters');
        if (!filters) continue;
        const vals = kids(filters, 'filter').map((f) => f.attrs.val);
        if (filters.attrs.blank === '1') vals.push('');
        criteria[rg.c1 + Number(fc.attrs.colId)] = vals;
      }
      const hidden = {};
      for (const k of Object.keys(sheet.hiddenRows)) {
        const r = Number(k);
        if (r > rg.r1 && r <= rg.r2 && Object.keys(criteria).length) { hidden[r] = true; delete sheet.hiddenRows[r]; }
      }
      sheet.filter = { ...rg, criteria, hidden };
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
            name: sl.attrs.name, cache: sl.attrs.cache, caption: sl.attrs.caption, columns: Number(sl.attrs.columnCount ?? 1), style: sl.attrs.style,
            showCaption: sl.attrs.showCaption !== '0', lockedPosition: sl.attrs.lockedPosition === '1', rowHeight: Number(sl.attrs.rowHeight ?? 241300),
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
    const t = readTable(parseXml(tx), sheet);
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
    if (f1 !== undefined && f1 !== '') base.f1 = f1;
    if (f2 !== undefined && f2 !== '') base.f2 = f2;
    for (const k of ['errorTitle', 'error', 'promptTitle', 'prompt']) if (a[k]) base[k] = a[k];
    if (type === 'any' && !base.prompt) return;
    for (const part of String(sqref ?? '').trim().split(/\s+/)) {
      if (!part) continue;
      const rg = /^[A-Z]+:[A-Z]+$/i.test(part) ? parseRangeName(`${part.split(':')[0]}1:${part.split(':')[1]}${MAX_ROWS}`) : refToRange(part);
      if (rg) out.push({ ...rg, r2: Math.min(rg.r2, MAX_ROWS - 1), ...base });
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
function readTable(root, sheet) {
  const rg = refToRange(root.attrs.ref ?? '');
  if (!rg) return null;
  const header = root.attrs.headerRowCount !== '0';
  const totals = Number(root.attrs.totalsRowCount ?? 0) > 0;
  const info = child(root, 'tableStyleInfo');
  const cols = kids(child(root, 'tableColumns'), 'tableColumn');
  const totalsFns = {};
  cols.forEach((tc, i) => {
    const f = tc.attrs.totalsRowFunction;
    if (f && f !== 'none' && f !== 'custom') totalsFns[rg.c1 + i] = f;
  });
  const af = child(root, 'autoFilter');
  let filter = null;
  if (af && header) {
    const criteria = {};
    for (const fc of kids(af, 'filterColumn')) {
      const filters = child(fc, 'filters');
      if (!filters) continue;
      const vals = kids(filters, 'filter').map((f) => f.attrs.val);
      if (filters.attrs.blank === '1') vals.push('');
      criteria[rg.c1 + Number(fc.attrs.colId)] = vals;
    }
    const hidden = {};
    if (Object.keys(criteria).length) {
      for (const k of Object.keys(sheet.hiddenRows)) {
        const r = Number(k);
        if (r > rg.r1 && r <= rg.r2 - (totals ? 1 : 0)) { hidden[r] = true; delete sheet.hiddenRows[r]; }
      }
    }
    filter = { criteria, hidden };
  }
  const name = (root.attrs.displayName || root.attrs.name || 'Table').replace(/\s/g, '_');
  return {
    id: `tb${Math.random().toString(36).slice(2, 9)}`, name, ...rg, r2: Math.max(rg.r2, rg.r1 + (header ? 1 : 0)),
    header, totals, style: normalizeStyleName(info?.attrs.name ?? DEFAULT_TABLE_STYLE),
    banded: info ? info.attrs.showRowStripes !== '0' : true, bandedCols: info?.attrs.showColumnStripes === '1',
    firstCol: info?.attrs.showFirstColumn === '1', lastCol: info?.attrs.showLastColumn === '1',
    filter, totalsFns, ...(header ? {} : { columns: cols.map((c) => c.attrs.name ?? '') }),
    _xmlId: Number(root.attrs.id), _colNames: cols.map((c) => c.attrs.name ?? ''),
  };
}

export function numberRaw(v, style) {
  const fmt = style?.numFmt;
  // 엑셀 1900 날짜 체계 (60 = 없는 날 1900-02-29 는 숫자 그대로)
  if ((fmt === 'date' || fmt === 'longdate') && Number.isInteger(v) && v > 0 && v !== 60) {
    const d = dateParts(v);
    return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`;
  }
  if (fmt === 'percent') {
    const s = `${Number((v * 100).toPrecision(15))}%`;
    const p = parseInput(s);
    if (typeof p.value === 'number' && Math.abs(p.value - v) < 1e-12) return s;
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
    const text = paras.map((p) => descendants(p, 't').map((t) => t.text).join('')).join('\n');
    // 글자 서식을 문단 · 글자 조각(run)마다 그대로 (엑셀과 같은 모양)
    const lvl = child(child(tx, 'lstStyle'), 'lvl1pPr');
    const baseR = child(lvl, 'defRPr');
    const runOf = (r, t) => {
      const pr = child(r, 'rPr');
      const a = { ...(baseR?.attrs ?? {}), ...(pr?.attrs ?? {}) };
      const out = { t };
      if (a.b === '1') out.b = true;
      if (a.i === '1') out.i = true;
      if (a.u && a.u !== 'none') out.u = true;
      if (a.strike && a.strike !== 'noStrike') out.s = true;
      if (a.sz) out.sz = Number(a.sz) / 100;
      const c = dmlColor(child(pr, 'solidFill') ?? child(baseR, 'solidFill'), ctx.theme);
      if (c) out.color = c;
      const face = child(pr, 'ea')?.attrs.typeface ?? child(pr, 'latin')?.attrs.typeface;
      if (face && !face.startsWith('+')) out.font = face;
      return out;
    };
    const rich = paras.map((p) => {
      const pPr = child(p, 'pPr');
      const al = pPr?.attrs.algn ?? lvl?.attrs.algn;
      const runs = [];
      for (const r of p.children) {
        if (r.name === 'r' || r.name === 'fld') runs.push(runOf(r, child(r, 't')?.text ?? ''));
        else if (r.name === 'br') runs.push({ t: '\n' });
      }
      const end = child(p, 'endParaRPr');
      const para = { runs };
      if (al) para.align = al === 'ctr' ? 'center' : al === 'r' ? 'right' : al === 'just' || al === 'dist' ? 'justify' : 'left';
      if (!runs.length && end?.attrs.sz) para.sz = Number(end.attrs.sz) / 100;
      return para;
    });
    const rPr = descendants(child(el, 'txBody'), 'rPr')[0] ?? descendants(child(el, 'txBody'), 'defRPr')[0];
    // 문단에 algn 이 없으면 목록 스타일, 그것도 없으면 DrawingML 기본값(왼쪽)
    const algn = descendants(child(el, 'txBody'), 'pPr')[0]?.attrs.algn ?? lvl?.attrs.algn;
    const shape = {
      id: uid('sh'), kind: isText ? 'textbox' : prstKind(prst), ...round(box), z: ++z,
      fill, stroke, text,
    };
    const xf = descendants(spPr, 'xfrm')[0];
    if (xf?.attrs.flipH === '1') shape.flip = true;
    if (xf?.attrs.flipV === '1') shape.flipV = true;
    if (Number(xf?.attrs.rot)) shape.rot = Math.round(Number(xf.attrs.rot) / 60000);
    // 선 끝 화살표 · 대시
    const endOn = (n) => { const t = child(ln, n)?.attrs.type; return t && t !== 'none'; };
    if (LINE_KINDS.has(shape.kind) && (endOn('tailEnd') || endOn('headEnd'))) shape.arrow = endOn('tailEnd') && endOn('headEnd') ? 'both' : 'end';
    if (LINE_KINDS.has(shape.kind) && !endOn('tailEnd') && endOn('headEnd')) { shape.flip = !shape.flip; shape.flipV = !shape.flipV; }
    const dashV = child(ln, 'prstDash')?.attrs.val;
    if (dashV && dashV !== 'solid') shape.dash = /dot/i.test(dashV) && !/dash/i.test(dashV) ? 'dot' : 'dash';
    const lw = Number(ln?.attrs.w);
    if (lw && stroke) shape.strokeWidth = Math.round((lw / EMU) * 4) / 4;
    if (rPr?.attrs.sz) shape.size = Number(rPr.attrs.sz) / 100;
    if (rPr?.attrs.b === '1') shape.bold = true;
    const tc = rPr && dmlColor(child(rPr, 'solidFill'), ctx.theme);
    if (tc) shape.color = tc;
    else if (!isText && fill) shape.color = '#ffffff';
    shape.align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' || algn === 'dist' ? 'justify' : 'left';
    // 서식이 섞여 있으면 문단 · 조각 그대로 보관 (한 가지 서식이면 단순 글자로 충분)
    const flat = rich.flatMap((p) => p.runs);
    const same = (k) => flat.every((r) => r[k] === flat[0]?.[k]);
    if (rich.length && (!['b', 'i', 'u', 'sz', 'color', 'font'].every(same) || rich.some((p) => (p.align ?? shape.align) !== shape.align))) {
      shape.paras = rich;
      delete shape.bold;
    }
    const body = child(tx, 'bodyPr');
    const anc = body?.attrs.anchor;
    shape.valign = anc === 'ctr' ? 'middle' : anc === 'b' ? 'bottom' : anc === 't' ? 'top' : isText ? 'top' : 'middle';
    const ins = (k, d) => (body?.attrs[k] !== undefined ? Number(body.attrs[k]) / EMU : d);
    const pad = [ins('tIns', 4.8), ins('rIns', 9.6), ins('bIns', 4.8), ins('lIns', 9.6)].map((v) => Math.round(v * 10) / 10);
    if (pad.join() !== '4.8,9.6,4.8,9.6') shape.pad = pad;
    if (body?.attrs.wrap === 'none') shape.nowrap = true;
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
    const emf = ext === 'emf' && `data:image/x-emf;base64,${toBase64(bytes)}`;
    const src = picSrc.get(target) ?? (emf ? emfDataUrl(bytes) : mime && `data:${mime};base64,${toBase64(bytes)}`);
    if (!src) { ctx.warnings.add(`지원하지 않는 그림 형식(${ext})은 가져오지 않았습니다.`); return; }
    const name = descendants(child(el, 'nvPicPr'), 'cNvPr')[0]?.attrs.name ?? '그림';
    picSrc.set(target, src); // 같은 그림을 여러 번 쓰면 한 번만 변환
    const im = { id: uid('im'), name, ...round(box), z: ++z, src };
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
    if (lc) { im.border = lc; im.borderW = Math.max(1, Math.round(Number(ln.attrs.w ?? 9525) / EMU)); }
    if (el.attrs.macro) im.macro = el.attrs.macro.replace(/^\[\d+\]!/, '');
    out.images.push(im);
  };

  // 그룹 도형 안의 좌표 변환
  const walk = (el, box, map) => {
    const place = (node) => {
      const xfrm = descendants(child(node, 'spPr') ?? child(node, 'grpSpPr'), 'xfrm')[0];
      if (!map || !xfrm) return box;
      const off = child(xfrm, 'off');
      const ext = child(xfrm, 'ext');
      return map(Number(off?.attrs.x ?? 0), Number(off?.attrs.y ?? 0), Number(ext?.attrs.cx ?? 0), Number(ext?.attrs.cy ?? 0));
    };
    switch (el.name) {
      case 'sp': case 'cxnSp': readShape(el, place(el)); break;
      case 'pic': readPic(el, place(el)); break;
      case 'graphicFrame': {
        const chartRef = descendants(el, 'chart')[0];
        const target = chartRef && rels[rid(chartRef)]?.target;
        const chart = target && readChart(files, target, ctx.theme);
        if (chart) {
          const b = round(place(el));
          out.charts.push({ id: uid('c'), ...chart, ...b, w: Math.max(120, b.w), h: Math.max(90, b.h), z: ++z });
        }
        break;
      }
      case 'grpSp': {
        const gb = place(el);
        const xfrm = descendants(child(el, 'grpSpPr'), 'xfrm')[0];
        const chOff = child(xfrm, 'chOff');
        const chExt = child(xfrm, 'chExt');
        const cx = Number(chExt?.attrs.cx) || 1;
        const cy = Number(chExt?.attrs.cy) || 1;
        const ox = Number(chOff?.attrs.x ?? 0);
        const oy = Number(chOff?.attrs.y ?? 0);
        const inner = (x, y, w, h) => ({
          x: gb.x + ((x - ox) / cx) * gb.w, y: gb.y + ((y - oy) / cy) * gb.h, w: (w / cx) * gb.w, h: (h / cy) * gb.h,
        });
        for (const k of el.children) if (['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame'].includes(k.name)) walk(k, gb, inner);
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
    const alt = child(anchor, 'AlternateContent');
    const sl = alt && descendants(child(alt, 'Choice'), 'slicer')[0];
    // 개체 위치 속성: twoCell(셀에 맞춰 위치와 크기 변경) · oneCell(위치만) · absolute(변경 안 함)
    const editAs = anchor.name === 'absoluteAnchor' ? 'absolute' : anchor.name === 'oneCellAnchor' ? 'oneCell' : anchor.attrs.editAs ?? 'twoCell';
    if (sl?.attrs.name) { out._slicerBoxes[sl.attrs.name] = { ...round(box), z: ++z, ...(editAs !== 'oneCell' ? { placement: editAs } : {}) }; continue; }
    const content = anchor.children.find((k) => ['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame'].includes(k.name));
    const before = [out.charts.length, out.images.length, out.shapes.length];
    if (content) walk(content, box, null);
    if (editAs !== 'twoCell') [out.charts, out.images, out.shapes].forEach((list, k) => { for (let i = before[k]; i < list.length; i++) list[i].placement = editAs; });
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
function readChart(files, path, theme = {}) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const plot = descendants(root, 'plotArea')[0];
  if (!plot) return null;
  const kinds = { barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter', radarChart: 'radar', bubbleChart: 'bubble', stockChart: 'stock' };
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
    return t;
  };
  const dl = (el) => {
    const d = child(el, 'dLbls');
    if (!d) return undefined;
    if (child(d, 'delete')?.attrs.val === '1') return false;
    return child(d, 'showVal')?.attrs.val === '1' || child(d, 'showPercent')?.attrs.val === '1' ? true : undefined;
  };
  const fmtCode = (el) => child(el, 'numFmt')?.attrs.formatCode;
  // 글자 서식 (txPr · rich 의 defRPr / rPr): 크기(pt) · 색 · 굵게
  const runFont = (el) => {
    const r = el && (descendants(el, 'defRPr')[0] ?? descendants(el, 'rPr')[0]);
    if (!r) return {};
    const out = {};
    if (r.attrs.sz) out.size = Number(r.attrs.sz) / 100;
    if (r.attrs.b === '1') out.bold = true;
    const c = dmlColor(child(r, 'solidFill'), theme);
    if (c) out.color = c;
    return out;
  };
  const series = [];
  const serNames = []; // 계열 이름(파일에 저장된 값) — 자동 제목용
  const serOrder = [];
  const fmts = [];
  let sheetName = null;
  const rows = [];
  for (const g of groups) {
    const gType = typeOf(g);
    const secondary = isSecondary(g) && groups.length > 1;
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
      const cache = descendants(valEl, 'pt').map((pt) => Number(descendants(pt, 'v')[0]?.text));
      const s = {
        name: txF && seriesRef(txF) ? { ref: seriesRef(txF) } : { text: txV ?? `계열${series.length + 1}` },
        ...(catF && seriesRef(catF) ? { cat: seriesRef(catF) } : {}),
        ...(valF && seriesRef(valF) ? { val: seriesRef(valF) } : { cache }),
      };
      if ((gType === 'scatter' || gType === 'bubble') && catF && seriesRef(catF)) { s.x = s.cat; delete s.cat; }
      const szF = descendants(child(ser, 'bubbleSize'), 'f')[0]?.text;
      if (szF && seriesRef(szF)) s.size = seriesRef(szF);
      for (const r of [s.name.ref, s.cat, s.val, s.x]) if (r?.sheet) sheetName ??= r.sheet;
      // 서식: 막대는 채우기 색, 꺾은선은 선 색
      const spPr = child(ser, 'spPr');
      const fill = dmlColor(child(spPr, 'solidFill'), theme);
      const line = dmlColor(child(child(spPr, 'ln'), 'solidFill'), theme);
      const marker = child(child(ser, 'marker'), 'symbol')?.attrs.val;
      const f = {};
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
      if (lnW && (gType === 'line' || gType === 'scatter')) f.lineWidth = Math.round((lnW / 12700) * (4 / 3) * 100) / 100;
      const mSize = Number(child(child(ser, 'marker'), 'size')?.attrs.val ?? 0);
      if (mSize) f.markerSize = mSize;
      const dsh = DASH_FROM[child(child(spPr, 'ln'), 'prstDash')?.attrs.val];
      if (dsh && gType === 'line') f.dash = dsh;
      if ((gType === 'bar' || gType === 'column') && child(spPr, 'ln') && child(child(spPr, 'ln'), 'solidFill')) f.outline = dmlColor(child(child(spPr, 'ln'), 'solidFill'), theme) ?? undefined;
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
      const LPOS = { ctr: 'center', inEnd: 'insideEnd', inBase: 'insideBase', outEnd: 'outEnd' };
      if (LPOS[lposV] && gType !== 'pie' && gType !== 'doughnut') f.labelPos = LPOS[lposV];
      if (gType !== 'pie' && gType !== 'doughnut') {
        const pc = {};
        for (const dp of kids(ser, 'dPt')) { const c = dmlColor(child(child(dp, 'spPr'), 'solidFill'), theme); if (c) pc[Number(child(dp, 'idx')?.attrs.val)] = c; }
        if (Object.keys(pc).length) f.pointColors = pc;
      }
      const lf = runFont(child(child(ser, 'dLbls') ?? child(g, 'dLbls'), 'txPr'));
      if (lf.size) f.labelSize = lf.size;
      if (lf.color) f.labelColor = lf.color;
      if (lf.bold) f.labelBold = true;
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
  if (!series.length) return null;
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
  // 글꼴 크기(pt): 제목 · 축 · 범례 (파일에 있을 때만)
  const tf = runFont(child(titleEl, 'tx')) ;
  const tf2 = tf.size ? tf : runFont(child(titleEl, 'txPr'));
  if (tf2.size) out.titleSize = tf2.size;
  if (title && child(titleEl, 'overlay')?.attrs.val === '1') out.titleOverlay = true;
  const axEl = kids(plot, 'catAx')[0] ?? kids(plot, 'valAx')[0];
  const af = runFont(child(axEl, 'txPr'));
  if (af.size) out.axisSize = af.size;
  const lg = runFont(child(child(chartEl, 'legend'), 'txPr'));
  if (lg.size) out.legendSize = lg.size;
  if (lg.color && lg.color !== '#000000' && lg.color !== '#595959') out.legendColor = lg.color;
  if (lg.bold) out.legendBold = true;
  if (range) out.range = range;
  if (sheetName) out.sheet = sheetName;
  if (fmts.some((f) => Object.keys(f).length)) out.seriesFmt = fmts;
  const bar = groups.find((g) => g.name === 'barChart') ?? groups.find((g) => g.name === 'lineChart' || g.name === 'areaChart');
  const grouping = child(bar, 'grouping')?.attrs.val;
  if (grouping === 'stacked' || grouping === 'percentStacked') out.grouping = grouping;
  const g0 = groups[0];
  const rs = child(g0, 'radarStyle')?.attrs.val;
  if (rs === 'filled' || rs === 'marker') out.radarStyle = rs;
  if (g0.name === 'stockChart' && child(g0, 'upDownBars')) out.ohlc = true;
  const hole = Number(child(g0, 'holeSize')?.attrs.val);
  if (hole && hole !== 50) out.hole = hole;
  const gapW = Number(child(bar, 'gapWidth')?.attrs.val);
  if (bar?.name === 'barChart' && gapW && gapW !== 150 && gapW !== 182) out.gap = gapW;
  const ovl = Number(child(bar, 'overlap')?.attrs.val);
  if (bar?.name === 'barChart' && Number.isFinite(ovl) && ovl !== 0 && ovl !== 100) out.overlap = ovl;
  if (bar?.name === 'barChart' && child(bar, 'varyColors')?.attrs.val === '1') out.varyColors = true;
  if (child(plot, 'dTable')) out.dataTable = true;
  if (!descendants(plot, 'majorGridlines').length) out.gridY = false;
  // WIXEL 전용 설정 (원래 차트 종류 · 팔레트 · 서식)
  const tbEl = descendants(root, 'props').find((x) => x.attrs.json);
  if (tbEl) { try { Object.assign(out, JSON.parse(tbEl.attrs.json)); } catch { /* 무시 */ } }
  if (out.wxPivot) { out.pivot = out.wxPivot; delete out.wxPivot; }
  const legend = child(chartEl, 'legend');
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
    const sc = child(ax, 'scaling');
    if (child(sc, 'min')) o.min = Number(child(sc, 'min').attrs.val);
    if (child(sc, 'max')) o.max = Number(child(sc, 'max').attrs.val);
    return Object.keys(o).length ? o : null;
  };
  const vals = [...valAxes.values()];
  const primaryAx = vals.find((ax) => child(ax, 'axPos')?.attrs.val !== 'r' && child(ax, 'crosses')?.attrs.val !== 'max') ?? vals[0];
  const secondaryAx = vals.find((ax) => ax !== primaryAx);
  const axes = {};
  if (axInfo(primaryAx)) axes.y = axInfo(primaryAx);
  if (axInfo(secondaryAx)) axes.y2 = axInfo(secondaryAx);
  const catT = child(kids(plot, 'catAx')[0], 'title');
  if (catT) axes.x = { title: descendants(catT, 't').map((x) => x.text).join('') };
  if (Object.keys(axes).length) out.axes = axes;
  // 피벗 차트: [파일]시트!피벗 이름
  const ps = descendants(child(root, 'pivotSource'), 'name')[0]?.text;
  if (ps) {
    const m = /^(?:\[[^\]]*\])?'?(.*?)'?!(.+)$/.exec(ps.trim());
    if (m) out.pivot = { sheet: m[1].replace(/''/g, "'"), name: m[2] };
  }
  return out;
}

// ───────────────────────── 피벗 테이블 · 슬라이서 (읽기) ─────────────────────────
/** ISO 날짜 · 시각 (시간대 없음) → 엑셀 날짜 일련번호 */
export function isoSerial(v) {
  const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d)(?::(\d\d(?:\.\d+)?))?)?/.exec(v ?? '');
  if (!m) return v ?? '';
  const ms = Date.UTC(1970, 0, 1, +(m[4] ?? 0), +(m[5] ?? 0)) + Math.round(+(m[6] ?? 0) * 1000);
  return serialOf(+m[1], +m[2], +m[3]) + ms / DAY_MS;
}
/** pivotCacheDefinition → { source: { ref, sheet, name }, fields: [{ name, items: [값] }] } */
function readPivotCache(files, path) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const ws = descendants(child(root, 'cacheSource'), 'worksheetSource')[0];
  const fields = kids(child(root, 'cacheFields'), 'cacheField').map((cf) => {
    const items = (child(cf, 'sharedItems')?.children ?? []).map((it) => {
      if (it.name === 'n') return Number(it.attrs.v);
      // 날짜 항목 (<d v="2026-05-01T00:00:00"/>): 원본 셀처럼 날짜 일련번호로 (필터 · 슬라이서 선택이 원본과 맞게)
      if (it.name === 'd') return isoSerial(it.attrs.v);
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
      if (rp) {
        group = {
          base: fg.attrs.base !== undefined ? Number(fg.attrs.base) : null,
          derived: cf.attrs.databaseField === '0',
          by: rp.attrs.groupBy ?? 'range',
          ...(rp.attrs.startDate ? { start: isoSerial(rp.attrs.startDate), end: isoSerial(rp.attrs.endDate) } : {}),
          ...(rp.attrs.startNum !== undefined ? { startNum: Number(rp.attrs.startNum), endNum: Number(rp.attrs.endNum), size: Number(rp.attrs.groupInterval ?? 1) } : {}),
        };
      }
    }
    return { name: unx(cf.attrs.name ?? ''), items: gitems && (group || !items.length) ? gitems : items, shared: items, db: formula === undefined && cf.attrs.databaseField !== '0', ...(group ? { group } : {}), ...(formula !== undefined ? { formula } : {}) };
  });
  const recRel = Object.values(relsOf(files, path)).find((r) => r.type === 'pivotCacheRecords');
  const snapshot = root.attrs.refreshOnLoad === '1' || !recRel ? null : { path: recRel.target };
  return { source: { ref: ws?.attrs.ref ?? null, sheet: ws?.attrs.sheet ?? null, name: ws?.attrs.name ?? null }, fields, snapshot };
}

/**
 * 피벗 캐시 레코드 (엑셀이 마지막으로 새로 고칠 때 저장한 원본) → [머리글, ...행]
 * 엑셀은 열 때 이 저장본으로 피벗을 보여 줌 (원본을 고친 뒤 새로 고치지 않았으면 원본과 다를 수 있음)
 */
const SNAPSHOT_MAX_BYTES = 40 << 20;
function readCacheRecords(files, cache) {
  const bytes = files[cache.snapshot.path];
  if (!bytes || bytes.length > SNAPSHOT_MAX_BYTES) return null;
  const xml = textOf(bytes);
  const fields = cache.fields.filter((f) => f.db);
  const rows = [fields.map((f) => f.name)];
  const recRe = /<r>([\s\S]*?)<\/r>|<r\/>/g;
  const cellRe = /<(x|n|s|d|b|m|e)(?:\s+v="([^"]*)")?[^>]*\/>/g;
  let m;
  while ((m = recRe.exec(xml))) {
    const row = new Array(fields.length).fill(null);
    if (m[1]) {
      let j = 0;
      let c;
      cellRe.lastIndex = 0;
      while ((c = cellRe.exec(m[1])) && j < fields.length) {
        const [, t, v] = c;
        row[j] = t === 'x' ? fields[j].shared[Number(v)] ?? null
          : t === 'n' ? Number(v)
            : t === 'd' ? isoSerial(v)
              : t === 'b' ? v === '1' || v === 'true'
                : t === 'm' ? null
                  : t === 'e' ? { error: v ?? '#N/A' }
                    : unx(decodeEntities(v ?? ''));
        j++;
      }
    }
    rows.push(row);
  }
  return rows.length > 1 ? rows : null;
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
        if (it) v.baseItem = itemText(cache.fields[bf]?.items[Number(it.attrs.x)] ?? null);
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
    ...(vIdx >= 0 && vIdx < colF.length ? { valuesPos: vIdx } : {}),
    ...(root.attrs.dataOnRows === '1' && values.length > 1 ? { valuesOnRows: true } : {}),
    layout: !fOutline ? 'tabular' : !fCompact ? 'outline' : 'compact',
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
  const src = cache.source;
  const tbl = src.name ? tables.find((t) => t.name.toLowerCase() === src.name.toLowerCase()) : null;
  if (tbl) Object.assign(def, { table: tbl.name, source: tbl.sheetName, range: { r1: tbl.r1, c1: tbl.c1, r2: tbl.r2, c2: tbl.c2 } });
  else if (src.ref && refToRange(src.ref)) Object.assign(def, { source: src.sheet ?? sheetName, range: refToRange(src.ref) });
  else return null;
  // 숨긴 항목 · 보고서 필터에서 고른 항목 → 필터 (슬라이서 선택 상태)
  const filters = {};
  pfs.forEach((pf, f) => {
    const its = kids(child(pf, 'items'), 'item').filter((it) => it.attrs.x !== undefined);
    if (!its.some((it) => it.attrs.h === '1')) return;
    filters[names[f]] = its.filter((it) => it.attrs.h !== '1').map((it) => itemText(cache.fields[f]?.items[Number(it.attrs.x)] ?? null));
  });
  for (const p of pageEls) {
    if (p.attrs.item === undefined) continue;
    const f = Number(p.attrs.fld);
    const its = kids(child(pfs[f], 'items'), 'item').filter((it) => it.attrs.x !== undefined);
    const it = its[Number(p.attrs.item)];
    if (it) filters[names[f]] = [itemText(cache.fields[f]?.items[Number(it.attrs.x)] ?? null)];
  }
  if (Object.keys(filters).length) def.filters = filters;
  // 셀에서 바꿔 쓴 필드 이름 (pivotField@name) · 항목 이름 (item@n)
  pfs.forEach((pf, f) => {
    if (pf.attrs.name && unx(pf.attrs.name) !== names[f]) (def.fieldCaptions ??= {})[names[f]] = unx(pf.attrs.name);
    for (const it of kids(child(pf, 'items'), 'item')) {
      if (it.attrs.n === undefined || it.attrs.x === undefined) continue;
      const t = itemText(cache.fields[f]?.items[Number(it.attrs.x)] ?? null);
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
  const si0 = child(root, 'pivotTableStyleInfo');
  def.style = si0?.attrs.name ?? 'None';
  const so = {
    rowHeaders: si0?.attrs.showRowHeaders !== '0', colHeaders: si0?.attrs.showColHeaders !== '0',
    bandRows: si0?.attrs.showRowStripes === '1', bandCols: si0?.attrs.showColStripes === '1',
  };
  if (!so.rowHeaders || !so.colHeaders || so.bandRows || so.bandCols) def.styleOpts = so;
  if (root.attrs.rowHeaderCaption) def.rowCaption = root.attrs.rowHeaderCaption;
  if (root.attrs.grandTotalCaption) def.grandCaption = root.attrs.grandTotalCaption;
  if (root.attrs.showError === '1' || root.attrs.showError === 'true') def.errorCaption = root.attrs.errorCaption ?? '';
  if (root.attrs.colHeaderCaption) def.colCaption = root.attrs.colHeaderCaption;
  if (root.attrs.missingCaption && root.attrs.showMissing !== '0') def.missingCaption = root.attrs.missingCaption;
  if (root.attrs.showHeaders === '0') def.showHeaders = false;
  if (root.attrs.showDrill === '0') def.showExpand = false;
  if (root.attrs.mergeItem === '1') def.mergeLabels = true;
  if (root.attrs.useAutoFormatting === '0') def.autofit = false;
  if (root.attrs.preserveFormatting === '0') def.preserveFormat = false;
  if (root.attrs.enableDrill === '0') def.enableDrill = false;
  // 축소한 항목 · 부분합 위치 · 빈 줄 · 레이블 반복
  const collapsed = {};
  pfs.forEach((pf, f) => {
    if (![...rowF, ...colF].includes(f)) return;
    const its = kids(child(pf, 'items'), 'item').filter((it) => it.attrs.sd === '0' && it.attrs.x !== undefined);
    if (its.length) collapsed[names[f]] = its.map((it) => itemText(cache.fields[f]?.items[Number(it.attrs.x)] ?? null));
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
  for (const f of [...rowF, ...colF, ...pageEls.map((p) => Number(p.attrs.fld))]) {
    const g = cache.fields[f]?.group;
    if (!g) continue;
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
        if (pit?.attrs.x !== undefined) at.push([names[f2], itemText(cache.fields[f2]?.items[Number(pit.attrs.x)] ?? null)]);
      }
      sort[names[f]] = { dir: st === 'ascending' ? 'asc' : 'desc', ...(by !== undefined ? { by } : {}), ...(at.length ? { at } : {}) };
    }
    const its = kids(child(pf, 'items'), 'item').filter((it) => it.attrs.x !== undefined && !it.attrs.t);
    if (its.length > 1) order[names[f]] = its.map((it) => itemText(cache.fields[f]?.items[Number(it.attrs.x)] ?? null));
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
        const text = pit && pit.attrs.x !== undefined ? itemText(cache.fields[f]?.items[Number(pit.attrs.x)] ?? null) : null;
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
      const toSerial = (v) => (v === undefined || v === '' ? undefined : Number.isFinite(Number(v)) ? Number(v) : isoSerial(v));
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
  if (dc && dc !== '값' && dc !== 'Values') def.dataCaption = unx(dc);
  const locEl = child(root, 'location');
  const loc = refToRange(locEl?.attrs.ref ?? '');
  // 클래식 레이아웃(열 필드 없이 값 여러 개): 값 이름 위에 '값' 단추 행 (firstHeaderRow 1 · firstDataRow 2)
  if (def.values.length > 1 && !def.valuesOnRows && !colF.length && Number(locEl?.attrs.firstHeaderRow) === 1 && Number(locEl?.attrs.firstDataRow) === 2) def.valuesHeadRow = true;
  if (loc) {
    const pageRows = def.pages?.length ? def.pages.length + 1 : 0;
    const top = Math.max(0, loc.r1 - pageRows);
    Object.assign(def, { top, left: loc.c1, area: { ...loc, r1: top } });
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
  const a = { ...(descendants(root, 'tabular')[0]?.attrs ?? {}), ...(tsc?.attrs ?? {}) };
  const opts = {};
  if (a.sortOrder === 'descending') opts.sort = 'desc';
  if (a.customListSort === '0' || a.customListSort === 'false') opts.customList = false;
  if (a.crossFilter === 'none') opts.markNoData = false;
  if (a.crossFilter === 'showItemsWithNoData') opts.noDataLast = false;
  if (descendants(root, 'slicerCacheHideItemsWithNoData').length) opts.hideNoData = true;
  return {
    opts,
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
    if (c?.name) caches.set(c.name, c);
  }
  const cacheFiles = new Map();
  sheets.forEach((s) => {
    for (const p of s._pivots) {
      if (!p.cachePath) continue;
      if (!cacheFiles.has(p.cachePath)) cacheFiles.set(p.cachePath, readPivotCache(files, p.cachePath));
      const cache = cacheFiles.get(p.cachePath);
      const def = cache && pivotDefFrom(p.root, cache, tables, s.name);
      // 엑셀의 저장본(캐시 레코드)으로 처음 화면을 그림 — 원본을 고치거나 새로 고치면 원본에서 다시 계산
      if (def && cache.snapshot && !cache.snapshot.failed) {
        if (!(ctx.pivotSnapshots ??= {})[p.cachePath]) {
          const rows = readCacheRecords(files, cache);
          if (rows) ctx.pivotSnapshots[p.cachePath] = rows; else cache.snapshot.failed = true;
        }
        if (ctx.pivotSnapshots[p.cachePath]) def.snapshotId = p.cachePath;
      }
      if (def && ctx.tableStyles?.[def.style] && !PRESET_STYLES[def.style]) def.styleDef = ctx.tableStyles[def.style]; // 파일에 정의된 사용자 지정 스타일 (WIXEL 모던 스타일은 이름으로 앎)
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
      if (!source) { ctx.warnings.add('연결 대상을 찾지 못한 슬라이서는 가져오지 않았습니다.'); continue; }
      const box = boxes[sl.name] ?? { x: 20 + i * 190, y: 20, w: 180, h: 200 };
      s.slicers.push({
        id: `sl${Math.random().toString(36).slice(2, 9)}`, caption: sl.caption ?? sl.name, source, columns: Math.max(1, sl.columns || 1),
        style: /^SlicerStyle(Light|Other|Dark)\d$/i.test(sl.style ?? '') || isModernSlicer(sl.style) ? sl.style : 'SlicerStyleLight1', multi: false, ...box,
        ...(!isModernSlicer(sl.style) && ctx.slicerStyles?.[sl.style] && Object.keys(ctx.slicerStyles[sl.style]).length ? { custom: ctx.slicerStyles[sl.style] } : {}),
        ...(sl.showCaption ? {} : { showHeader: false }),
        ...(sl.lockedPosition ? { noMove: true } : {}),
        ...(c.opts ?? {}),
        ...(Math.abs(sl.rowHeight - 241300) > 20000 ? { buttonHeight: Math.max(14, Math.round(sl.rowHeight / EMU)) } : {}),
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
export async function readXlsxAsync(bytes, onProgress) {
  onProgress?.({ p: 0, msg: '압축 푸는 중' });
  // 시트 · 공유 문자열처럼 큰 부분은 브라우저 내장 압축 해제로 (계산 체인 · 피벗 캐시 레코드는 읽지 않으므로 풀지 않음)
  const files = await unzipAsync(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes),
    (name, size) => size > 1 << 20 && !/calcChain|pivotCacheRecords/i.test(name));
  const it = readXlsxSteps(files);
  let last = performance.now();
  for (;;) {
    const s = it.next();
    if (s.done) return s.value;
    const now = performance.now();
    if (now - last > 40) {
      onProgress?.(s.value);
      await new Promise((res) => setTimeout(res, 0));
      last = performance.now();
    }
  }
}

function* readXlsxSteps(files) {
  yield { p: 0.05, msg: '압축 푸는 중' };
  // xlsb: 바이너리 부분을 같은 경로의 xlsx XML 로 바꾼 뒤 그대로 읽음 (큰 시트의 셀은 행 생성기로 바로)
  if (isXlsb(files)) yield* convertXlsb(files);
  const wbPath = Object.keys(files).find((f) => /^xl\/workbook\.xml$/i.test(f))
    ?? relsTarget(files, '', 'officeDocument');
  if (!wbPath || !files[wbPath]) throw new Error('엑셀 통합 문서(.xlsx)가 아닙니다');
  const wbRoot = parseXml(textOf(files[wbPath]));
  const wbRels = relsOf(files, wbPath);
  const theme = readTheme(files, wbRels);
  const { xfs, dxfs, dxfOf, tableStyles, wbFont, slicerStyles, cellStyles } = readStyles(files, wbRels, theme);
  const mdw = digitWidth(wbFont);
  const ssRel = Object.values(wbRels).find((r) => r.type === 'sharedStrings');
  const strings = files.__xlsb ? files.__xlsb.strings : ssRel && files[ssRel.target] ? kids(parseXml(textOf(files[ssRel.target])), 'si').map(allText) : [];
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
  const ctx = { mdw, wbFont, xfs, dxfs, dxfOf, tableStyles, slicerStyles, strings, theme, warnings: new Set(), isName, nameMulti, richImages: readRichImages(files, wbRels) };
  const sheets = [];
  const warnings = [];
  const sheetCodes = {};
  let unsupported = 0;
  for (const sh of kids(child(wbRoot, 'sheets'), 'sheet')) {
    const rel = wbRels[rid(sh)];
    if (!rel || rel.type !== 'worksheet' || !files[rel.target]) {
      if (rel) warnings.push(`'${sh.attrs.name}' 시트(차트 시트 등)는 가져오지 않았습니다.`);
      continue;
    }
    const si = sheets.length;
    const total = kids(child(wbRoot, 'sheets'), 'sheet').length;
    const sheet = yield* progressOf(readSheet(files, rel.target, ctx), (rows) => ({ p: 0.05 + 0.85 * (si / total), msg: `'${sh.attrs.name}' 시트 읽는 중 (${rows.toLocaleString()}행)` }));
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
    for (const part of pn.text.split(',')) {
      const ref = part.slice(part.lastIndexOf('!') + 1).replace(/\$/g, '');
      const rows = /^(\d+):(\d+)$/.exec(ref);
      const cols = /^([A-Z]+):([A-Z]+)$/i.exec(ref);
      if (pn.kind === 'area') { const rg = parseRangeName(ref); if (rg) page.area = rg; } else if (rows) page.titleRows = [Number(rows[1]) - 1, Number(rows[2]) - 1];
      else if (cols) page.titleCols = [nameToCol(cols[1].toUpperCase()), nameToCol(cols[2].toUpperCase())];
    }
    sh.page = page;
  }
  // 외부 통합 문서 참조 ([1]시트!A1): 파일에 저장된 외부 값을 숨긴 읽기 전용 시트로 (엑셀도 연결을 새로 고치기 전에는 이 값을 보여 줌)
  const extSheets = [];
  const externals = readExternalLinks(files, wbRoot, wbRels, extSheets);
  yield { p: 0.92, msg: "피벗 테이블 · 슬라이서 연결 중" };
  linkPivotsAndSlicers(files, wbRels, sheets, ctx);
  dropOffAxisFilters(sheets);
  pushAll(sheets, extSheets);
  if (unsupported) warnings.push(`지원하지 않는 함수가 쓰인 수식 ${unsupported}개는 수식을 유지하고 파일에 저장된 계산 결과를 표시합니다.`);
  if (files.__xlsb?.unsupported) warnings.push(`바이너리 통합 문서(.xlsb)에서 해석하지 못한 수식 ${files.__xlsb.unsupported}개는 저장된 계산 결과(값)로 가져왔습니다.`);
  pushAll(warnings, ctx.warnings);
  if (!sheets.length) throw new Error('가져올 시트가 없습니다');
  const data = { sheets };
  if (ctx.pivotSnapshots) data.pivotSnapshots = ctx.pivotSnapshots;
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
  { const trel = Object.values(wbRels).find((r) => r.type === 'theme'); const tx = trel && textOf(files[trel.target]); if (tx && tx.length < 400000) data.themeXml = tx; }
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
    const fs = child(wbRoot, 'fileSharing');
    if (fs && (fs.attrs.readOnlyRecommended === '1' || fs.attrs.readOnlyRecommended === 'true')) props.readOnlyRecommended = true;
    const custom = files['docProps/custom.xml'] && textOf(files['docProps/custom.xml']);
    if (custom && /name="_MarkAsFinal"[^>]*>\s*<vt:bool>(true|1)<\/vt:bool>/.test(custom)) props.markedFinal = true;
    if (Object.keys(props).length) data.props = props;
  }
  data.defaultFont = wbFont; // 통합 문서 기본 글꼴 (표준 스타일) — 셀 기본 크기 · 열 너비 변환에 씀
  // 기본 셀 서식(xf 0): s 속성이 없는 셀에 적용됨 (한국어 엑셀은 보통 세로 가운데 맞춤)
  if (xfs[0] && Object.keys(xfs[0]).length) data.baseStyle = { ...xfs[0] };
  if (cellStyles?.length) data.cellStyles = cellStyles;
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
    return x;
  }
  const clr = THEME_SLOTS.map((slot, i) => (i === 0 ? `<a:lt1><a:sysClr val="window" lastClr="${colors[0]}"/></a:lt1>` : i === 1 ? `<a:dk1><a:sysClr val="windowText" lastClr="${colors[1]}"/></a:dk1>` : `<a:${slot}><a:srgbClr val="${colors[i]}"/></a:${slot}>`));
  const order = [clr[1], clr[0], clr[3], clr[2], ...clr.slice(4)].join('');
  const font = (latin, ea) => `<a:latin typeface="${latin}" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hang" typeface="${ea}"/>`;
  const solid = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
  const ln = (w) => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr">${solid}<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`;
  const name = esc(wb.themeName ?? 'Office 테마');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${name}"><a:themeElements><a:clrScheme name="${name}">${order}</a:clrScheme>`
    + `<a:fontScheme name="Office"><a:majorFont>${font('맑은 고딕', '맑은 고딕')}</a:majorFont><a:minorFont>${font('맑은 고딕', '맑은 고딕')}</a:minorFont></a:fontScheme>`
    + `<a:fmtScheme name="Office"><a:fillStyleLst>${solid}${solid}${solid}</a:fillStyleLst><a:lnStyleLst>${ln(6350)}${ln(12700)}${ln(19050)}</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${solid}${solid}${solid}</a:bgFillStyleLst></a:fmtScheme>`
    + '</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>';
}

class StylePool {
  constructor(baseFont = WRITE_FONT, baseStyle = null) {
    this.baseFont = { name: baseFont.name || DEFAULT_FONT, size: baseFont.size || 11 };
    this.baseStyle = baseStyle && Object.keys(baseStyle).length ? baseStyle : null;
    this.fonts = [`<font><sz val="${this.baseFont.size}"/><color theme="1"/><name val="${esc(this.baseFont.name)}"/><family val="3"/><charset val="129"/></font>`];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = [];
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.dxfs = [];
    this.tableStyles = new Map();
    this.slicerStyles = new Map();
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
    const font = `<font>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.strike ? '<strike/>' : ''}${style.underline ? '<u/>' : ''}<sz val="${style.size || this.baseFont.size}"/>${style.color ? `<color rgb="${argb(style.color)}"/>` : '<color theme="1"/>'}<name val="${esc(style.font || this.baseFont.name)}"/><family val="3"/><charset val="129"/></font>`;
    const fontId = this.intern('font', this.fonts, font);
    const fillId = style.pattern
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
    if (style.valign) align.push(`vertical="${style.valign === 'middle' ? 'center' : 'top'}"`);
    if (style.wrap) align.push('wrapText="1"');
    if (style.indent) align.push(`indent="${style.indent}"`);
    if (style.rotate) align.push(`textRotation="${style.rotate === 255 ? 255 : style.rotate < 0 ? 90 - style.rotate : style.rotate}"`);
    if (style.shrink) align.push('shrinkToFit="1"');
    const prot = style.locked === false || style.hideFormula ? `<protection${style.locked === false ? ' locked="0"' : ''}${style.hideFormula ? ' hidden="1"' : ''}/>` : '';
    const inner = (align.length ? `<alignment ${align.join(' ')}/>` : '') + prot;
    if (this.styleXfMode) return `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align.length ? ' applyAlignment="1"' : ''}${prot ? ' applyProtection="1"' : ''}${inner ? `>${inner}</xf>` : '/>'}`;
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align.length ? ' applyAlignment="1"' : ''}${prot ? ' applyProtection="1"' : ''}${inner ? `>${inner}</xf>` : '/>'}`;
    const id = this.xfs.length;
    this.xfs.push(xml);
    this.maps.xf.set(k, id);
    return id;
  }

  /** 이름 있는 셀 스타일 → cellStyleXfs + cellStyles (엑셀 [셀 스타일] 갤러리의 사용자 지정) */
  namedStyles(list) {
    this.cellStyleXfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>'];
    this.cellStyleList = ['<cellStyle name="표준" xfId="0" builtinId="0"/>'];
    const seen = new Set(['표준']);
    for (const cs of list ?? []) {
      if (!cs?.name || seen.has(cs.name)) continue;
      seen.add(cs.name);
      this.styleXfMode = true;
      const xml = Object.keys(cs.style ?? {}).length ? this.xfOf(cs.style) : '<xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>';
      this.styleXfMode = false;
      this.cellStyleList.push(`<cellStyle name="${esc(cs.name)}" xfId="${this.cellStyleXfs.length}"${cs.builtinId !== undefined ? ` builtinId="${cs.builtinId}" customBuiltin="1"` : ''}/>`);
      this.cellStyleXfs.push(xml);
    }
  }

  dxf(style) {
    const font = style.color || style.bold || style.italic || style.underline || style.strike
      ? `<font>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.strike ? '<strike/>' : ''}${style.underline ? '<u/>' : ''}${style.color ? `<color rgb="${argb(style.color)}"/>` : ''}</font>` : '';
    const fill = style.fill ? `<fill><patternFill><bgColor rgb="${argb(style.fill)}"/></patternFill></fill>` : '';
    const side = (n, k) => (style[k] ? `<${n} style="${style[`${k}s`] ?? 'thin'}">${style[`${k}c`] ? `<color rgb="${argb(style[`${k}c`])}"/>` : '<color auto="1"/>'}</${n}>` : '');
    const border = style.bt || style.bb || style.bl || style.br || style.bv || style.bh ? `<border>${side('left', 'bl')}${side('right', 'br')}${side('top', 'bt')}${side('bottom', 'bb')}${side('vertical', 'bv')}${side('horizontal', 'bh')}</border>` : '';
    // 표시 형식도 조건부 서식으로 바꿀 수 있음 (dxf 안의 numFmt)
    const code = style.numFmt ? fmtCode(style) : null;
    const nf = code !== null ? `<numFmt numFmtId="${this.fmtId(style)}" formatCode="${esc(code)}"/>` : '';
    this.dxfs.push(`<dxf>${font}${nf}${fill}${border}</dxf>`);
    return this.dxfs.length - 1;
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
    if (!sl.custom && !isModernSlicer(base)) return base;
    const c = slicerColors(sl);
    const name = sl.custom ? `WIXEL 사용자 지정 ${[...JSON.stringify(c)].reduce((h, ch) => (h * 31 + ch.charCodeAt(0)) >>> 0, 7).toString(36)}` : base;
    if (this.slicerStyles.has(name)) return name;
    const bd = (color) => ({ bt: true, bb: true, bl: true, br: true, btc: color, bbc: color, blc: color, brc: color });
    const whole = this.dxf({ fill: c.frame, ...bd(c.border) });
    const head = this.dxf({ color: c.head, bold: true });
    this.tableStyles.set(name, `<tableStyle name="${esc(name)}" pivot="0" table="0" count="2"><tableStyleElement type="wholeTable" dxfId="${whole}"/><tableStyleElement type="headerRow" dxfId="${head}"/></tableStyle>`);
    const els = [
      ['selectedItemWithData', { fill: c.selFill, color: c.selText, ...bd(c.selBorder) }],
      ['selectedItemWithNoData', { fill: c.selFill, color: c.noData, ...bd(c.selBorder) }],
      ['unselectedItemWithData', { fill: c.item, color: c.itemText, ...bd(c.itemBorder) }],
      ['unselectedItemWithNoData', { fill: c.item, color: c.noData, ...bd(c.itemBorder) }],
      ['hoveredSelectedItemWithData', { fill: c.selFill, color: c.selText, ...bd(c.selBorder) }],
      ['hoveredUnselectedItemWithData', { fill: c.item, color: c.itemText, ...bd(c.selBorder) }],
    ].map(([type, st]) => `<x14:slicerStyleElement type="${type}" dxfId="${this.dxf(st)}"/>`);
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
      + (this.tableStyles.size ? `<tableStyles count="${this.tableStyles.size}" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16">${[...this.tableStyles.values()].join('')}</tableStyles>` : '')
      + (this.slicerStyles.size ? `<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="${NS_X14}"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1">${[...this.slicerStyles.values()].join('')}</x14:slicerStyles></ext></extLst>` : '')
      + '</styleSheet>';
  }
}

/** 앱 수식 → 파일 수식 (_xlfn. 접두사, 표 참조 표준 형식, '@'·'#' 변환) */
let fileNameCheck = null;
function exportFormula(raw, hereTable = null, dynamic = false) {
  return toFileFormula(raw, { hereTable, dynamic, isName: fileNameCheck });
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

function cfXml(rule, pool, priority, x14 = null) {
  const ref = [rule, ...(rule.more ?? [])].map((g) => rangeRef(g)).join(' ');
  const top = cellName(rule.r1, rule.c1);
  const lit = (v) => {
    const t = String(v ?? '');
    if (t.startsWith('=')) return exportFormula(t);
    const p = parseInput(t);
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
  return `<conditionalFormatting${rule.pivot && rule.pivot.scope !== 'selection' ? ' pivot="1"' : ''} sqref="${ref}">${body}</conditionalFormatting>`;
}

function chartXml(wb, si, chart, fileName = 'Book1.xlsx') {
  const srcIndex = chart.sheet ? wb.sheetIndexByName(chart.sheet) : si;
  const s = srcIndex >= 0 ? srcIndex : si;
  const refText = (sheetIdx, r1, c1, r2, c2) => `${quoteSheetName(wb.sheets[sheetIdx].name)}!${rangeRef({ r1, c1, r2, c2 }, true)}`;
  const label = (v) => (v === null || v === undefined ? '' : typeof v === 'number' ? formatGeneral(v) : isError(v) ? v.code : String(v));
  const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(label(v))}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (vals, code = 'General') => `<c:numCache><c:formatCode>${esc(code)}</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
  const data = chartModelData(wb, si, chart);
  // 엑셀 2016 차트(폭포 · 깔때기 · 히스토그램 · 파레토 · 트리맵 · 상자 수염)는 호환 차트로 저장하고 원래 종류는 확장 정보로 보관
  const FALLBACK = { waterfall: 'column', histogram: 'column', pareto: 'column', treemap: 'column', boxWhisker: 'column', funnel: 'bar' };
  const baseType = chart.type === 'combo' ? 'column' : FALLBACK[chart.type] ?? chart.type;
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
      refs.push({ tx: ref(sr.name?.ref), cat: ref(sr.cat ?? sr.x), val: ref(sr.val) });
    }
  } else if (chart.range) {
    const rg = chart.range;
    const rows = [];
    for (let r = rg.r1; r <= rg.r2; r++) { const row = []; for (let c = rg.c1; c <= rg.c2; c++) row.push(wb.getValue(s, r, c)); rows.push(row); }
    const L = chartLayout(rows, baseType, !!chart.byRows);
    const dr0 = rg.r1 + L.firstDataRow;
    const dc0 = rg.c1 + L.firstDataCol;
    const all = [];
    if (L.byCols) {
      for (let c = dc0; c <= rg.c2; c++) {
        all.push({ tx: L.headRow ? refText(s, rg.r1, c, rg.r1, c) : null, cat: L.catCol ? refText(s, dr0, rg.c1, rg.r2, rg.c1 + Math.max(0, L.firstDataCol - 1)) : null, val: refText(s, dr0, c, rg.r2, c), nums: rows.slice(L.firstDataRow).some((r) => typeof r[c - rg.c1] === 'number') });
      }
    } else {
      for (let r = dr0; r <= rg.r2; r++) {
        all.push({ tx: L.catCol ? refText(s, r, rg.c1, r, rg.c1) : null, cat: L.headRow ? refText(s, rg.r1, dc0, rg.r1, rg.c2) : null, val: refText(s, r, dc0, r, rg.c2), nums: rows[r - rg.r1].slice(L.firstDataCol).some((v) => typeof v === 'number') });
      }
    }
    const withNums = all.filter((x) => x.nums);
    pushAll(refs, (withNums.length ? withNums : all));
  }
  const pal = paletteOf(chart);
  const series = data.series.map((sr, i) => ({
    ...sr, ...refs[i], type: FALLBACK[sr.type] ?? sr.type ?? baseType, axis: sr.axis ?? 0, color: sr.color ?? pal[i % pal.length],
  }));
  const LBL_POS = { center: 'ctr', insideEnd: 'inEnd', insideBase: 'inBase', outEnd: 'outEnd', above: 't', below: 'b', left: 'l', right: 'r' };
  const dLbls = (on, code, pct = false, pos = null) => (on || pct
    ? `<c:dLbls>${code && typeof code === 'string' && !pct ? `<c:numFmt formatCode="${esc(code)}" sourceLinked="0"/>` : ''}<c:spPr><a:noFill/><a:ln><a:noFill/></a:ln></c:spPr>${pos && LBL_POS[pos] ? `<c:dLblPos val="${LBL_POS[pos]}"/>` : ''}<c:showLegendKey val="0"/><c:showVal val="${pct ? 0 : 1}"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="${pct ? 1 : 0}"/><c:showBubbleSize val="0"/></c:dLbls>`
    : '');
  const serXml = (sr, i) => {
    const type = sr.type;
    const scatter = type === 'scatter' || type === 'bubble';
    const pie = type === 'pie' || type === 'doughnut';
    const hex = String(sr.color).replace('#', '').toUpperCase().slice(0, 6);
    const fill = `<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`;
    const tx = sr.tx ? `<c:tx><c:strRef><c:f>${esc(sr.tx)}</c:f>${strCache([sr.name])}</c:strRef></c:tx>` : `<c:tx><c:v>${esc(sr.name)}</c:v></c:tx>`;
    let spPr;
    if (type === 'bubble') spPr = `<c:spPr>${fill}</c:spPr>`;
    else if (type === 'stock') spPr = '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    else if (scatter) spPr = /line|smooth/i.test(chart.scatterStyle ?? '') ? `<c:spPr><a:ln w="19050">${fill}</a:ln></c:spPr>` : '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    else if (type === 'radar') spPr = `<c:spPr>${chart.radarStyle === 'filled' ? fill : ''}<a:ln w="28575">${fill}</a:ln></c:spPr>`;
    else if (type === 'line') spPr = `<c:spPr><a:ln w="${sr.lineWidth ? Math.round((sr.lineWidth * 3 / 4) * 12700) : 28575}" cap="rnd">${fill}${DASH_XML[sr.dash] ?? ''}<a:round/></a:ln></c:spPr>`;
    else if (pie) spPr = '<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>';
    else spPr = `<c:spPr>${fill}${sr.outline ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${String(sr.outline).replace('#', '').toUpperCase().slice(0, 6)}"/></a:solidFill></a:ln>` : ''}</c:spPr>`;
    const markerSym = sr.marker === 'none' || sr.marker === false ? 'none' : typeof sr.marker === 'string' ? sr.marker : 'circle';
    const markerSym2 = sr.marker === undefined && (chart.marker === 'none' || (scatter && /^(line|smooth)$/.test(chart.scatterStyle ?? '')) || (type === 'radar' && chart.radarStyle !== 'marker') || type === 'stock') ? 'none' : markerSym;
    const marker = (type === 'line' || type === 'radar' || type === 'stock' || (scatter && type !== 'bubble')) ? (markerSym2 === 'none' ? '<c:marker><c:symbol val="none"/></c:marker>' : `<c:marker><c:symbol val="${markerSym2}"/><c:size val="5"/><c:spPr>${fill}<a:ln w="9525">${fill}</a:ln></c:spPr></c:marker>`) : '';
    const dPt = pie ? sr.values.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/>${chart.explode ? `<c:explosion val="${Math.round(chart.explode)}"/>` : ''}<c:spPr><a:solidFill><a:srgbClr val="${(sr.colors?.[k] ?? pal[k % pal.length]).slice(1)}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
    const invert = type === 'column' || type === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
    // 막대 · 꺾은선의 데이터 요소별 색 · '요소마다 다른 색'
    const ptColor = (k) => sr.pointColors?.[k] ?? (chart.varyColors && (type === 'column' || type === 'bar') ? pal[k % pal.length] : null);
    const dPtBar = !pie && (sr.pointColors || chart.varyColors) && (type === 'column' || type === 'bar') ? sr.values.map((_, k) => (ptColor(k) ? `<c:dPt><c:idx val="${k}"/><c:invertIfNegative val="0"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${String(ptColor(k)).replace('#', '').toUpperCase().slice(0, 6)}"/></a:solidFill></c:spPr></c:dPt>` : '')).join('') : '';
    // 추세선
    const TREND = { linear: 'linear', exp: 'exp', movingAvg: 'movingAvg' };
    const trend = sr.trend && TREND[sr.trend] && !pie ? `<c:trendline><c:spPr><a:ln w="19050" cap="rnd"><a:solidFill><a:srgbClr val="${String(sr.trendColor ?? sr.color).replace('#', '').toUpperCase().slice(0, 6)}"/></a:solidFill><a:prstDash val="sysDash"/></a:ln></c:spPr><c:trendlineType val="${TREND[sr.trend]}"/>${sr.trend === 'movingAvg' ? `<c:period val="${Math.max(2, sr.trendPeriod ?? 3)}"/>` : sr.trendForward ? `<c:forward val="${sr.trendForward}"/>` : ''}<c:dispRSqr val="0"/><c:dispEq val="0"/></c:trendline>` : '';
    const code = typeof sr.numFmt === 'string' ? sr.numFmt : null;
    // 원형: 레이블을 따로 정하지 않았으면 백분율 (WIXEL 화면과 같게)
    const pieP = pie && sr.labels !== false && (sr.pct || (sr.labels ?? chart.labels) === undefined);
    const labels = dLbls(pie && sr.labels === false ? false : sr.labels ?? chart.labels, code, pieP, pie ? null : sr.labelPos);
    const cats = data.categories;
    const catTag = scatter ? 'xVal' : 'cat';
    // 다단계 항목 (여러 열의 항목 범위): 안쪽 → 바깥 순서의 lvl
    const lvls = !scatter && sr.cat && data.catLevels?.length
      ? [cats.map((v, k) => ({ k, v })), ...data.catLevels.map((spans) => spans.filter((sp) => sp.text !== '').map((sp) => ({ k: sp.start, v: sp.text })))]
      : null;
    const cat = lvls
      ? `<c:cat><c:multiLvlStrRef><c:f>${esc(sr.cat)}</c:f><c:multiLvlStrCache><c:ptCount val="${cats.length}"/>${lvls.map((pts) => `<c:lvl>${pts.map((x) => `<c:pt idx="${x.k}"><c:v>${esc(x.v)}</c:v></c:pt>`).join('')}</c:lvl>`).join('')}</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>`
      : sr.cat
      ? `<c:${catTag}><c:strRef><c:f>${esc(sr.cat)}</c:f>${strCache(cats)}</c:strRef></c:${catTag}>`
      : cats.length ? `<c:${catTag}><c:strLit><c:ptCount val="${cats.length}"/>${cats.map((v, k) => `<c:pt idx="${k}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strLit></c:${catTag}>` : '';
    const valTag = scatter ? 'yVal' : 'val';
    const val = sr.val
      ? `<c:${valTag}><c:numRef><c:f>${esc(sr.val)}</c:f>${numCache(sr.values, code ?? 'General')}</c:numRef></c:${valTag}>`
      : `<c:${valTag}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${sr.values.length}"/>${sr.values.map((v, k) => (typeof v === 'number' ? `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numLit></c:${valTag}>`;
    const bsz = type === 'bubble' ? `<c:bubbleSize><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${sr.values.length}"/>${(sr.size ?? []).map((v, k) => (typeof v === 'number' ? `<c:pt idx="${k}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numLit></c:bubbleSize><c:bubble3D val="0"/>` : '';
    const smooth = type === 'line' || (scatter && type !== 'bubble') ? `<c:smooth val="${sr.smooth || /smooth/i.test(chart.scatterStyle ?? '') ? 1 : 0}"/>` : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${invert}${marker}${dPt}${dPtBar}${labels}${trend}${cat}${val}${bsz}${smooth}</c:ser>`;
  };
  const pieLike = baseType === 'pie' || baseType === 'doughnut';
  const grouping = chart.grouping ?? 'clustered';
  // 계열 종류 · 축별로 차트 그룹 (콤보 차트)
  const groups = [];
  series.forEach((sr, i) => {
    const kind = pieLike ? baseType : sr.type === 'bar' || sr.type === 'column' ? (baseType === 'bar' ? 'bar' : 'column') : sr.type;
    const axis = pieLike ? 0 : sr.axis === 1 ? 1 : 0;
    let g = groups.find((x) => x.kind === kind && x.axis === axis);
    if (!g) { g = { kind, axis, xml: [] }; groups.push(g); }
    g.xml.push(serXml(sr, i));
  });
  const hasSecondary = groups.some((g) => g.axis === 1) && groups.some((g) => g.axis === 0);
  const ax = (axis) => (axis === 1 && hasSecondary ? '<c:axId val="333333333"/><c:axId val="444444444"/>' : '<c:axId val="111111111"/><c:axId val="222222222"/>');
  const groupXml = groups.map((g) => {
    const body = g.xml.join('');
    const a = ax(g.axis);
    switch (g.kind) {
      case 'bar': case 'column': {
        const stackedG = grouping === 'stacked' || grouping === 'percentStacked';
        return `<c:barChart><c:barDir val="${g.kind === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="${chart.varyColors ? 1 : 0}"/>${body}<c:gapWidth val="${typeof chart.gap === 'number' ? Math.round(chart.gap) : g.kind === 'bar' ? 182 : 150}"/>${stackedG ? '<c:overlap val="100"/>' : typeof chart.overlap === 'number' ? `<c:overlap val="${Math.round(Math.max(-100, Math.min(100, chart.overlap)))}"/>` : ''}${a}</c:barChart>`;
      }
      case 'line': return `<c:lineChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${body}<c:marker val="1"/>${a}</c:lineChart>`;
      case 'area': return `<c:areaChart><c:grouping val="${grouping === 'clustered' ? 'standard' : grouping}"/><c:varyColors val="0"/>${body}${a}</c:areaChart>`;
      case 'pie': return `<c:pieChart><c:varyColors val="1"/>${body}<c:firstSliceAng val="${Math.round(chart.firstAngle ?? 0)}"/></c:pieChart>`;
      case 'doughnut': return `<c:doughnutChart><c:varyColors val="1"/>${body}<c:firstSliceAng val="${Math.round(chart.firstAngle ?? 0)}"/><c:holeSize val="${Math.round(chart.hole ?? 50)}"/></c:doughnutChart>`;
      case 'scatter': return `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${body}${a}</c:scatterChart>`;
      case 'bubble': return `<c:bubbleChart><c:varyColors val="0"/>${body}<c:bubbleScale val="${Math.round(chart.bubbleScale ?? 100)}"/><c:showNegBubbles val="0"/>${a}</c:bubbleChart>`;
      case 'radar': return `<c:radarChart><c:radarStyle val="${chart.radarStyle === 'filled' ? 'filled' : 'marker'}"/><c:varyColors val="0"/>${body}${a}</c:radarChart>`;
      case 'stock': return `<c:stockChart>${body}<c:hiLowLines/>${chart.ohlc ? '<c:upDownBars><c:gapWidth val="150"/><c:upBars/><c:downBars/></c:upDownBars>' : ''}${a}</c:stockChart>`;
      default: return '';
    }
  }).join('');
  const horizontal = baseType === 'bar';
  const axTitle = (t) => (t ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="1000" b="0"/><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '');
  const scaling = (cfg) => `<c:scaling><c:orientation val="${cfg?.reverse ? 'maxMin' : 'minMax'}"/>${typeof cfg?.max === 'number' ? `<c:max val="${cfg.max}"/>` : ''}${typeof cfg?.min === 'number' ? `<c:min val="${cfg.min}"/>` : ''}</c:scaling>`;
  const numFmt = (cfg) => (cfg?.numFmt ? `<c:numFmt formatCode="${esc(cfg.numFmt)}" sourceLinked="0"/>` : '<c:numFmt formatCode="General" sourceLinked="1"/>');
  const catAxis = (id, cross, pos, del) => (baseType === 'scatter' || baseType === 'bubble'
    ? `<c:valAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${del ? 1 : 0}"/><c:axPos val="${pos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx><c:axId val="${id}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="${del || chart.axes?.x?.hide ? 1 : 0}"/><c:axPos val="${pos}"/>${chart.gridX ? '<c:majorGridlines/>' : ''}${del ? '' : axTitle(chart.axes?.x?.title)}<c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`);
  const valAxis = (id, cross, pos, cfg, grid, crosses = 'autoZero') => `<c:valAx><c:axId val="${id}"/>${scaling(cfg)}<c:delete val="${cfg?.hide ? 1 : 0}"/><c:axPos val="${pos}"/>${grid && chart.gridY !== false ? '<c:majorGridlines/>' : ''}${axTitle(cfg?.title)}${numFmt(cfg)}<c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="${cross}"/><c:crosses val="${crosses}"/><c:crossBetween val="${baseType === 'area' || baseType === 'scatter' || baseType === 'bubble' ? 'midCat' : 'between'}"/>${typeof cfg?.major === 'number' ? `<c:majorUnit val="${cfg.major}"/>` : ''}</c:valAx>`;
  let axesXml = '';
  if (!pieLike) {
    axesXml = catAxis(111111111, 222222222, horizontal ? 'l' : 'b', false) + valAxis(222222222, 111111111, horizontal ? 'b' : 'l', chart.axes?.y, true);
    if (hasSecondary) axesXml += catAxis(333333333, 444444444, horizontal ? 'l' : 'b', true) + valAxis(444444444, 333333333, horizontal ? 't' : 'r', chart.axes?.y2, false, 'max');
  }
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="ko-KR" sz="1400" b="0"/><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="${chart.titleOverlay ? 1 : 0}"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const lp = chart.legend ?? (series.length > 1 || pieLike ? 'b' : 'none');
  // 범례 글꼴 (색 · 크기 · 굵게)
  const legTx = chart.legendColor || chart.legendSize || chart.legendBold ? `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr${chart.legendSize ? ` sz="${Math.round(chart.legendSize * 100)}"` : ''}${chart.legendBold ? ' b="1"' : ''}>${chart.legendColor ? `<a:solidFill><a:srgbClr val="${String(chart.legendColor).replace('#', '').toUpperCase()}"/></a:solidFill>` : ''}</a:defRPr></a:pPr><a:endParaRPr lang="ko-KR"/></a:p></c:txPr>` : '';
  const legend = lp !== 'none' ? `<c:legend><c:legendPos val="${lp}"/><c:overlay val="0"/>${legTx}</c:legend>` : '';
  // 위셀 지표 선택 피벗 차트: 엑셀에는 피벗 테이블 범위를 참조하는 일반 차트로 (엑셀 피벗 차트는 모든 값 필드를 강제로 보이므로)
  const subsetPivot = !!chart.pivot?.values?.length;
  const pivotSrc = chart.pivot && !subsetPivot ? `<c:pivotSource><c:name>${esc(`[${fileName}]${quoteSheetName(chart.pivot.sheet ?? wb.sheets[si].name)}!${chart.pivot.name}`)}</c:name><c:fmtId val="0"/></c:pivotSource>` : '';
  const pivotFmts = chart.pivot && !subsetPivot ? `<c:pivotFmts>${series.map((_, i) => `<c:pivotFmt><c:idx val="${i}"/></c:pivotFmt>`).join('')}</c:pivotFmts>` : '';
  // 차트 영역 · 그림 영역 채우기와 테두리
  const hexOf = (c) => String(c).replace('#', '').toUpperCase().slice(0, 6);
  const areaSpPr = chart.fill || chart.border ? `<c:spPr>${chart.fill ? `<a:solidFill><a:srgbClr val="${hexOf(chart.fill)}"/></a:solidFill>` : ''}${chart.border ? `<a:ln w="9525"><a:solidFill><a:srgbClr val="${hexOf(chart.border)}"/></a:solidFill></a:ln>` : ''}</c:spPr>` : '';
  const plotSpPr = chart.plotFill ? `<c:spPr><a:solidFill><a:srgbClr val="${hexOf(chart.plotFill)}"/></a:solidFill></c:spPr>` : '';
  // WIXEL 전용 설정 (엑셀은 무시): 원래 차트 종류 · 팔레트 · 서식
  const TB_KEYS = ['hiddenSeries', 'hiddenCats', 'legendBold', 'type', 'byRows', 'fieldButtons', 'palette', 'scatterStyle', 'radarStyle', 'ohlc', 'explode', 'hole', 'gap', 'marker', 'gridX', 'gridY', 'fill', 'plotFill', 'border', 'titleColor', 'titleBold', 'textColor', 'gridColor', 'rounded', 'totals', 'binCount', 'binWidth', 'upColor', 'downColor', 'totalColor', 'bubbleScale', 'firstAngle', 'showMean', 'connectors'];
  const tb = Object.fromEntries(TB_KEYS.filter((k) => chart[k] !== undefined && chart[k] !== null).map((k) => [k, chart[k]]));
  if (subsetPivot) tb.wxPivot = chart.pivot; // 위셀로 다시 열면 슬라이서와 연동되는 피벗 차트로 복원
  const extLst = Object.keys(tb).length > 1 || FALLBACK[chart.type] ? `<c:extLst><c:ext uri="{5E2A6C7B-8F4D-4B1A-9C3E-7D6F1A2B3C4D}" xmlns:tb="urn:tabula:chart"><tb:props json="${esc(JSON.stringify(tb))}"/></c:ext></c:extLst>` : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}"><c:roundedCorners val="${chart.rounded ? 1 : 0}"/>${pivotSrc}<c:chart>${title}${pivotFmts}<c:plotArea><c:layout/>${groupXml}${axesXml}${chart.dataTable && !pieLike ? '<c:dTable><c:showHorzBorder val="1"/><c:showVertBorder val="1"/><c:showOutline val="1"/><c:showKeys val="1"/></c:dTable>' : ''}${plotSpPr}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>${areaSpPr}${extLst}</c:chartSpace>`;
}

const DASH_XML = { dash: '<a:prstDash val="dash"/>', dot: '<a:prstDash val="sysDot"/>', dashDot: '<a:prstDash val="dashDot"/>', longDash: '<a:prstDash val="lgDash"/>', sysDash: '<a:prstDash val="sysDash"/>' };
const DASH_FROM = { dash: 'dash', sysDot: 'dot', dot: 'dot', dashDot: 'dashDot', sysDashDot: 'dashDot', lgDash: 'longDash', sysDash: 'sysDash' };
const KIND_PRST = { arrow: 'rightArrow', textbox: 'rect', line: 'straightConnector1' };
const prstOf = (kind) => KIND_PRST[kind] ?? (GEOM[kind] || LINE_KINDS.has(kind) ? kind : 'rect');
const hex6 = (c) => (c ?? '#000000').replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6);

/** 도형 → <xdr:sp> / <xdr:cxnSp> */
function shapeXml(sh, id, xfrm) {
  const name = esc(sh.name || `${sh.kind === 'textbox' ? 'TextBox' : '도형'} ${id - 1}`);
  const fill = sh.fill ? `<a:solidFill><a:srgbClr val="${hex6(sh.fill)}"/></a:solidFill>` : '<a:noFill/>';
  const isLine = LINE_KINDS.has(sh.kind);
  const dashXml = sh.dash ? `<a:prstDash val="${sh.dash === 'dot' ? 'sysDot' : 'dash'}"/>` : '';
  const ends = isLine && sh.arrow ? `${sh.arrow === 'both' ? '<a:headEnd type="triangle"/>' : ''}<a:tailEnd type="triangle"/>` : '';
  const ln = sh.stroke ? `<a:ln w="${Math.round((sh.strokeWidth ?? 1) * EMU)}"><a:solidFill><a:srgbClr val="${hex6(sh.stroke)}"/></a:solidFill>${dashXml}${ends}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
  if (isLine) {
    return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>${xfrm(sh)}<a:prstGeom prst="${sh.kind === 'line' ? 'straightConnector1' : sh.kind}"><a:avLst/></a:prstGeom>${ln}</xdr:spPr></xdr:cxnSp>`;
  }
  const algnOf = (a) => (a === 'center' ? 'ctr' : a === 'right' ? 'r' : a === 'justify' ? 'just' : 'l');
  const algn = algnOf(sh.align);
  const rPr = `<a:rPr lang="ko-KR" sz="${Math.round((sh.size ?? 11) * 100)}"${sh.bold ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="${hex6(sh.color ?? '#000000')}"/></a:solidFill></a:rPr>`;
  // 문단 · 조각별 서식이 있으면 그대로 저장
  const runXml = (r) => {
    if (r.t === '\n') return '<a:br/>';
    const a = ['lang="ko-KR"', `sz="${Math.round((r.sz ?? sh.size ?? 11) * 100)}"`, r.b ?? sh.bold ? 'b="1"' : '', r.i ? 'i="1"' : '', r.u ? 'u="sng"' : '', r.s ? 'strike="sngStrike"' : ''].filter(Boolean).join(' ');
    return `<a:r><a:rPr ${a}><a:solidFill><a:srgbClr val="${hex6(r.color ?? sh.color ?? '#000000')}"/></a:solidFill>${r.font ? `<a:latin typeface="${esc(r.font)}"/><a:ea typeface="${esc(r.font)}"/>` : ''}</a:rPr><a:t>${esc(r.t)}</a:t></a:r>`;
  };
  const paras = sh.paras
    ? sh.paras.map((p) => `<a:p><a:pPr algn="${algnOf(p.align ?? sh.align)}"/>${p.runs.length ? p.runs.map(runXml).join('') : `<a:endParaRPr lang="ko-KR" sz="${Math.round((p.sz ?? sh.size ?? 11) * 100)}"/>`}</a:p>`).join('')
    : String(sh.text ?? '').split('\n').map((line) => `<a:p><a:pPr algn="${algn}"/>${line ? `<a:r>${rPr}<a:t>${esc(line)}</a:t></a:r>` : `<a:endParaRPr lang="ko-KR" sz="${Math.round((sh.size ?? 11) * 100)}"/>`}</a:p>`).join('');
  const anchor = sh.valign ? { top: 't', middle: 'ctr', bottom: 'b' }[sh.valign] : sh.kind === 'textbox' ? 't' : 'ctr';
  return `<xdr:sp macro="${sh.macro ? `[0]!${esc(sh.macro)}` : ''}" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr${sh.kind === 'textbox' ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr>${xfrm(sh)}<a:prstGeom prst="${prstOf(sh.kind)}"><a:avLst/></a:prstGeom>${fill}${ln}</xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="${anchor}"/><a:lstStyle/>${paras}</xdr:txBody></xdr:sp>`;
}

/** 목록 원본: 범위 참조가 아니면 "a,b" 로 감싸기 */
function dvFormula(rule, f) {
  if (f === undefined || f === null || f === '') return '';
  let t = String(f).trim().replace(/^=/, '');
  if (rule.type === 'list' && !t.startsWith('"')) {
    const bang = t.lastIndexOf('!');
    const ref = (bang > 0 ? t.slice(bang + 1) : t).replace(/\$/g, '');
    if (!parseRangeName(ref)) t = `"${t.replace(/"/g, '')}"`;
  }
  return t;
}

function validationXml(v) {
  const attrs = [`type="${v.type === 'any' ? 'none' : v.type}"`];
  if (v.errorStyle === 'warning' || v.errorStyle === 'info') attrs.push(`errorStyle="${v.errorStyle === 'info' ? 'information' : 'warning'}"`);
  if (!['list', 'custom', 'any'].includes(v.type) && v.op && v.op !== 'between') attrs.push(`operator="${v.op}"`);
  if (v.allowBlank !== false) attrs.push('allowBlank="1"');
  if (v.type === 'list' && v.showDropdown === false) attrs.push('showDropDown="1"');
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
function fieldItems(data, f, order = null) {
  const seen = new Map();
  for (const r of data) { const k = keyOf(r[f]); if (!seen.has(`${typeof k}:${k}`)) seen.set(`${typeof k}:${k}`, k); }
  let keys = sortKeys([...seen.values()]);
  if (order?.length) {
    // 수동 순서(파일에서 가져온 항목 순서)를 유지하고 새 항목은 뒤에
    const pos = new Map(order.map((t, i) => [t, i]));
    keys = [...keys.filter((k) => pos.has(itemText(k))).sort((a, b) => pos.get(itemText(a)) - pos.get(itemText(b))), ...keys.filter((k) => !pos.has(itemText(k)))];
  }
  const index = new Map(keys.map((k, i) => [`${typeof k}:${k}`, i]));
  return { keys, index };
}

// ─── 피벗 그룹 (엑셀 fieldGroup) ───
const XL_GROUP_BY = { years: 'years', quarters: 'quarters', months: 'months', mdays: 'days', number: 'range' };
const XL_DATE_GROUP = new Set(['years', 'quarters', 'months', 'mdays']);
const serialIso = (v) => { const d = dateParts(Math.floor(v)); return `${d.y}-${String(d.m).padStart(2, '0')}-${String(d.d).padStart(2, '0')}`; };
/** 그룹 설정 → { keys: 엑셀 groupItems 순서의 항목 키(앱의 그룹 키와 같은 글자), rangePr } */
function excelGroupItems(spec, min, max) {
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
  const keys = [`<${serialIso(start)}`];
  if (spec.by === 'years') {
    const y1 = Number(serialIso(start).slice(0, 4));
    const y2 = Number(serialIso(end).slice(0, 4));
    for (let y = y1; y <= y2; y++) keys.push(y);
  } else if (spec.by === 'quarters') for (let q = 1; q <= 4; q++) keys.push(`${q}분기`);
  else if (spec.by === 'months') for (let m = 1; m <= 12; m++) keys.push(`${m}월`);
  else for (let m = 1; m <= 12; m++) for (let d = 1; d <= [31, 29, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][m - 1]; d++) keys.push(`${m}월${d}일`);
  keys.push(`>${serialIso(end)}`);
  return { keys, rangePr: `<rangePr groupBy="${XL_GROUP_BY[spec.by]}" startDate="${serialIso(start)}T00:00:00" endDate="${serialIso(end)}T00:00:00"/>` };
}
/** 날짜 필드의 sharedItems (항목은 쓰지 않음: 그룹 항목이 대신함) */
function dateSharedItemsXml(values) {
  const nums = values.filter((v) => typeof v === 'number');
  const blank = values.some((v) => v === null || v === '');
  if (!nums.length || nums.length + (blank ? values.filter((v) => v === null || v === '').length : 0) !== values.length) return sharedItemsXml(values, null);
  return `<sharedItems containsSemiMixedTypes="0" containsNonDate="0" containsDate="1" containsString="0"${blank ? ' containsBlank="1"' : ''} minDate="${serialIso(minOf(nums))}T00:00:00" maxDate="${serialIso(maxOf(nums) + 1)}T00:00:00"/>`;
}

function sharedItemsXml(values, keys) {
  const nums = values.filter((v) => typeof v === 'number');
  const hasStr = values.some((v) => v !== null && v !== '' && typeof v !== 'number');
  const hasBlank = values.some((v) => v === null || v === '');
  const hasNum = nums.length > 0;
  const attrs = [];
  if (!hasStr) attrs.push('containsSemiMixedTypes="0"', 'containsString="0"');
  if (hasNum) {
    attrs.push('containsNumber="1"');
    if (nums.every((n) => Number.isInteger(n))) attrs.push('containsInteger="1"');
    attrs.push(`minValue="${minOf(nums)}"`, `maxValue="${maxOf(nums)}"`);
  }
  if (hasBlank) attrs.push('containsBlank="1"');
  if (hasStr && hasNum) attrs.push('containsMixedTypes="1"');
  if (!keys) return `<sharedItems${attrs.length ? ` ${attrs.join(' ')}` : ''}/>`;
  const items = keys.map((k) => (k === EMPTY ? '<m/>' : typeof k === 'number' ? `<n v="${k}"/>` : `<s v="${esc(String(k))}"/>`)).join('');
  return `<sharedItems${attrs.length ? ` ${attrs.join(' ')}` : ''} count="${keys.length}">${items}</sharedItems>`;
}

/** 피벗 원본을 구별하는 키 (같은 원본의 피벗은 캐시 하나를 함께 씀 — 슬라이서 연결에 필요) */
function pivotSourceKey(def) {
  return def.table ? `t:${String(def.table).toLowerCase()}` : `r:${String(def.source ?? '').toLowerCase()}!${def.range ? rangeRef(def.range) : ''}`;
}

/**
 * 같은 원본을 쓰는 피벗들의 공유 캐시 → { cacheXml, header, nBase, items: Map(필드 → {keys, index}), data, src }
 * extraFields: 슬라이서가 쓰는 필드 이름(소문자)
 */
function buildPivotCache(wb, defs, cacheId, extraFields) {
  const src = pivotSourceData(wb, defs[0]);
  if (!src || src.rows.length < 2) return null;
  const baseHeader = headerNames(src.rows);
  const nBase = baseHeader.length;
  const calcs = [];
  for (const d of defs) for (const c of d.calcFields ?? []) {
    if (!baseHeader.some((h) => h.toLowerCase() === c.name.toLowerCase()) && !calcs.some((x) => x.name.toLowerCase() === c.name.toLowerCase())) calcs.push(c);
  }
  // 그룹화 (엑셀 형식으로 쓸 수 있는 것만): 제자리 그룹(필드 자체) · 파생 그룹 필드('월2' = '일'을 월로)
  const bx = (n) => baseHeader.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const grouped = new Map(); // 필드 이름(소문자) → { name, spec, base: 원본 필드 번호, derived }
  for (const d of defs) for (const [n, g] of Object.entries(d.groups ?? {})) {
    if (!g || !XL_GROUP_BY[g.by] || grouped.has(n.toLowerCase())) continue;
    const derived = !!g.base && bx(n) < 0;
    const base = derived ? bx(g.base) : bx(n);
    if (base >= 0) grouped.set(n.toLowerCase(), { name: n, spec: g, base, derived });
  }
  const derivedList = [...grouped.values()].filter((g) => g.derived);
  const header = [...baseHeader, ...calcs.map((c) => c.name), ...derivedList.map((g) => g.name)];
  const nCalc = calcs.length;
  const fx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const data = src.rows.slice(1).filter((r) => !r.every((v) => v === null || v === ''));
  const listed = new Set();
  for (const def of defs) {
    const d = normalizeDef(def, header);
    [...d.rows, ...d.cols, ...d.pages, ...Object.keys(d.filters)].forEach((n) => listed.add(fx(n)));
  }
  header.forEach((h, i) => { if (extraFields?.has(h.toLowerCase())) listed.add(i); });
  const items = new Map();
  for (const f of listed) if (f >= 0 && f < nBase && !grouped.has(header[f].toLowerCase())) items.set(f, fieldItems(data, f));
  // 그룹 필드의 항목 = 엑셀 groupItems 전체 (피벗 필드의 x 가 이 목록을 가리킴)
  const groupXml = new Map();
  for (const g of grouped.values()) {
    const f = fx(g.name);
    const vals = data.map((r) => r[g.base]).filter((v) => typeof v === 'number');
    const gi = excelGroupItems(g.spec, vals.length ? minOf(vals) : 0, vals.length ? maxOf(vals) : 0);
    items.set(f, { keys: gi.keys, index: new Map(gi.keys.map((k, i) => [`${typeof k}:${k}`, i])) });
    const par = !g.derived ? derivedList.find((x) => x.base === g.base) : null;
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
    const vals = data.map((r) => r[f]);
    // 날짜로 묶는 필드는 날짜 필드로 표시 (엑셀이 그룹을 다시 만들 수 있게)
    const shared = g && XL_DATE_GROUP.has(grouped.get(h.toLowerCase())?.spec.by) ? dateSharedItemsXml(vals) : sharedItemsXml(vals, g ? null : items.get(f)?.keys ?? null);
    return `<cacheField name="${esc(h)}" numFmtId="${g && XL_DATE_GROUP.has(grouped.get(h.toLowerCase())?.spec.by) ? 14 : 0}">${shared}${g ?? ''}</cacheField>`;
  }).join('');
  const sourceXml = src.table
    ? `<worksheetSource name="${esc(src.table)}"/>`
    : `<worksheetSource ref="${rangeRef(src.ref)}" sheet="${esc(wb.sheets[src.si].name)}"/>`;
  const cacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" xmlns:tb="${NS_TB}" mc:Ignorable="tb" saveData="0" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${data.length}">`
    + `<cacheSource type="worksheet">${sourceXml}</cacheSource><cacheFields count="${header.length}">${cacheFields}</cacheFields>`
    + `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${cacheId}"/></ext></extLst></pivotCacheDefinition>`;
  return { cacheXml, header, nBase, nCalc, items, data, src, cacheId };
}

/** 피벗 정의 + 공유 캐시 → 피벗 테이블 XML */
/** 피벗 조건부 서식 (엑셀 '규칙 적용 대상': 값 필드 전체 · 행 필드 수준) → <conditionalFormats>. priority 는 시트 규칙 순서와 같음 */
function pivotCondXml(wb, si, def, header, values) {
  const conds = (wb.sheets[si].cond ?? []).filter((r) => r.r1 < EXCEL_MAX_ROWS);
  const out = [];
  conds.forEach((rl, i) => {
    const pv = rl.pivot;
    if (!pv || pv.name !== (def.name ?? '') || pv.scope === 'selection') return;
    const vi = Math.max(0, values.findIndex((v) => valueName(v) === pv.value));
    const refs = [`<reference field="4294967294" count="1" selected="0"><x v="${vi}"/></reference>`];
    const rf = pv.scope === 'field' && pv.rowField ? header.findIndex((h) => String(h).toLowerCase() === String(pv.rowField).toLowerCase()) : -1;
    // 엑셀과 같은 형식: 행 필드 참조는 항목 없이(count="0") — 그 필드의 모든 항목
    if (rf >= 0) refs.push(`<reference field="${rf}" count="0" selected="0"/>`);
    out.push(`<conditionalFormat scope="${pv.scope === 'data' ? 'data' : 'field'}" priority="${i + 1}"><pivotAreas count="1"><pivotArea outline="0" collapsedLevelsAreSubtotals="1" fieldPosition="0"><references count="${refs.length}">${refs.join('')}</references></pivotArea></pivotAreas></conditionalFormat>`);
  });
  return out.length ? `<conditionalFormats count="${out.length}">${out.join('')}</conditionalFormats>` : '';
}

function pivotParts(wb, si, def, cache, name, pool) {
  const { src, header, nBase, data } = cache;
  const nCalcEnd = nBase + (cache.nCalc ?? 0);
  const isCalc = (f) => f >= nBase && f < nCalcEnd;
  const cacheId = cache.cacheId;
  const d = { ...normalizeDef(def, header), header };
  if (!d.rows.length && !d.cols.length && !d.values.length) return null;
  const fx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const rowF = d.rows.map(fx);
  const colF = d.cols.map(fx);
  const pageF = d.pages.map(fx);
  const filters = Object.entries(d.filters).map(([n, allowed]) => [fx(n), new Set(allowed)]).filter(([i]) => i >= 0);
  // 이 피벗의 항목 순서 (캐시 순서를 수동 순서대로 재배열) → x 는 이 목록의 위치
  const items = new Map();
  for (const [f, it] of cache.items) {
    const ord = d.sort[header[f]] ? null : d.order[header[f]];
    let keys = it.keys;
    if (ord?.length) {
      const pos = new Map(ord.map((t, i) => [t, i]));
      keys = [...keys.filter((k) => pos.has(itemText(k))).sort((a, b) => pos.get(itemText(a)) - pos.get(itemText(b))), ...keys.filter((k) => !pos.has(itemText(k)))];
    }
    items.set(f, { keys, index: new Map(keys.map((k, i) => [`${typeof k}:${k}`, i])), cacheIndex: it.index });
  }

  const hiddenKey = (f, k) => {
    const set = filters.find(([i]) => i === f)?.[1];
    return set ? !set.has(itemText(k)) : false;
  };
  // 보이는 항목 기준으로 결과 배치 계산 (앱 화면과 같은 배치)
  const { meta } = computePivot(resolvePivot(src.rows, def), d);
  const values = meta.values;
  const V = values.length;
  const multiV = V > 1;
  const onRows = multiV && !!d.valuesOnRows; // Σ 값이 행 영역 (dataOnRows)
  const colMulti = multiV && !onRows;
  const xOf = (f, key) => items.get(f)?.index.get(`${typeof key}:${key}`) ?? 0;
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
  const pivotFields = header.map((h, f) => {
    if (isCalc(f)) return `<pivotField${valueFieldIdx.has(f) ? ' dataField="1"' : ''} dragToRow="0" dragToCol="0" dragToPage="0" showAll="0" defaultSubtotal="0"/>`;
    const attrs = [];
    if (d.fieldCaptions?.[h]) attrs.push(`name="${esc(d.fieldCaptions[h])}"`);
    if (rowF.includes(f)) attrs.push('axis="axisRow"');
    else if (colF.includes(f)) attrs.push('axis="axisCol"');
    else if (pageF.includes(f)) attrs.push('axis="axisPage"');
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
    const coll = onAxis ? new Set(d.collapsed?.[h] ?? []) : null;
    // 항목 레이블 반복 (x14 확장)
    const fill = rowF.includes(f) && d.repeatLabels ? `<extLst><ext uri="{2946ED86-A175-432a-8AC1-64E0C546D7DE}" xmlns:x14="${NS_X14}"><x14:pivotField fillDownLabels="1"/></ext></extLst>` : '';
    const scope = srt && srt.by !== undefined && srt.by !== null
      ? `<autoSortScope><pivotArea dataOnly="0" outline="0" fieldPosition="0"><references count="1"><reference field="4294967294" count="1" selected="0"><x v="${Number(srt.by) || 0}"/></reference></references></pivotArea></autoSortScope>`
      : '';
    const it = items.get(f);
    if (!it) return scope || fill ? `<pivotField ${attrs.join(' ')}>${scope}${fill}</pivotField>` : `<pivotField ${attrs.join(' ')}/>`;
    const caps = d.itemCaptions?.[h];
    const list = it.keys.map((k) => `<item${caps?.[itemText(k)] !== undefined ? ` n="${esc(caps[itemText(k)])}"` : ''}${hiddenKey(f, k) ? ' h="1"' : ''}${coll?.has(itemText(k)) ? ' sd="0"' : ''} x="${it.cacheIndex.get(`${typeof k}:${k}`)}"/>`).join('');
    const def0 = onAxis && !subOn ? '' : '<item t="default"/>';
    return `<pivotField ${attrs.join(' ')}><items count="${it.keys.length + (def0 ? 1 : 0)}">${list}${def0}</items>${scope}${fill}</pivotField>`;
  }).join('');

  const top = (def.top ?? 0) + meta.pageRows;
  const left = def.left ?? 0;
  const loc = { r1: top, c1: left, r2: top + meta.bodyRows - 1, c2: left + meta.width - 1 };
  const firstHeaderRow = colF.length ? 1 : multiV ? (d.valuesHeadRow && !d.valuesOnRows ? 1 : 0) : 1;
  const pageXml = pageF.length ? `<pageFields count="${pageF.length}">${pageF.map((f) => {
    const allowed = filters.find(([i]) => i === f)?.[1];
    const one = allowed && allowed.size === 1 ? items.get(f)?.keys.findIndex((k) => itemText(k) === [...allowed][0]) : -1;
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
    else if (bf >= 0 && v.baseItem !== undefined && items.get(bf)) bi = Math.max(0, items.get(bf).keys.findIndex((k) => itemText(k) === String(v.baseItem)));
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
      const sv = (v) => (v === undefined ? '' : serialIso(Number(v)).slice(0, 10));
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
  if (styleName && isModernStyle(styleName)) pool.presetTableStyle(styleName, true);
  else if (styleName && def.styleDef && !/^PivotStyle(Light|Medium|Dark)\d+$/i.test(styleName)) pool.pivotStyle(styleName, def.styleDef);
  const tableAttrs = [
    `name="${esc(name)}"`, `cacheId="${cacheId}"`, 'applyNumberFormats="0"', 'applyBorderFormats="0"', 'applyFontFormats="0"', 'applyPatternFormats="0"',
    'applyAlignmentFormats="0"', 'applyWidthHeightFormats="1"', `dataCaption="${esc(d.dataCaption ?? '값')}"`, ...(onRows ? ['dataOnRows="1"'] : []), 'updatedVersion="6"', 'minRefreshableVersion="3"', `useAutoFormatting="${def.autofit === false ? 0 : 1}"`,
    ...(def.mergeLabels ? ['mergeItem="1"'] : []), ...(def.showHeaders === false ? ['showHeaders="0"'] : []), ...(def.preserveFormat === false ? ['preserveFormatting="0"'] : []), ...(def.multiFilters ? [] : []), ...(def.enableDrill === false ? ['enableDrill="0"'] : []),
    ...(d.rowCaption ? [`rowHeaderCaption="${esc(d.rowCaption)}"`] : []),
    ...(d.grandCaption ? [`grandTotalCaption="${esc(d.grandCaption)}"`] : []),
    ...(d.errorCaption !== null && d.errorCaption !== undefined ? ['showError="1"', ...(d.errorCaption ? [`errorCaption="${esc(d.errorCaption)}"`] : [])] : []), ...(d.colCaption ? [`colHeaderCaption="${esc(d.colCaption)}"`] : []),
    ...(d.grandRows ? [] : ['rowGrandTotals="0"']), ...(d.grandCols ? [] : ['colGrandTotals="0"']),
    ...(d.missingCaption ? [`missingCaption="${esc(d.missingCaption)}"`] : []), ...(d.showExpand ? [] : ['showDrill="0"']),
    'itemPrintTitles="1"', 'createdVersion="6"', 'indent="0"', ...(tabular || outline ? ['compact="0"', 'compactData="0"'] : []),
    `outline="${tabular ? 0 : 1}"`, `outlineData="${tabular ? 0 : 1}"`, ...(tabular ? ['gridDropZones="1"'] : []), 'multipleFieldFilters="0"',
  ];
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="${NS_MAIN}" ${tableAttrs.join(' ')}>`
    + `<location ref="${rangeRef(loc)}" firstHeaderRow="${firstHeaderRow}" firstDataRow="${meta.headerRows}" firstDataCol="${meta.labelCols}"${pageF.length ? ` rowPageCount="${pageF.length}" colPageCount="1"` : ''}/>`
    + `<pivotFields count="${header.length}">${pivotFields}</pivotFields>`
    + (rowF.length || onRows ? `<rowFields count="${rowF.length + (onRows ? 1 : 0)}">${[...rowF, ...(onRows ? [-2] : [])].map((f) => `<field x="${f}"/>`).join('')}</rowFields>` : '')
    + `<rowItems count="${rowXml.length}">${rowXml.join('')}</rowItems>`
    + (colFieldsAll.length ? `<colFields count="${colFieldsAll.length}">${colFieldsAll.map((f) => `<field x="${f}"/>`).join('')}</colFields>` : '')
    + `<colItems count="${colXml.length}">${colXml.join('')}</colItems>`
    + pageXml
    + `<dataFields count="${values.length}">${dataXml}</dataFields>`
    + pivotCondXml(wb, si, def, header, values)
    + `<pivotTableStyleInfo${styleName ? ` name="${esc(styleName)}"` : ''} showRowHeaders="${so.rowHeaders === false ? 0 : 1}" showColHeaders="${so.colHeaders === false ? 0 : 1}" showRowStripes="${so.bandRows ? 1 : 0}" showColStripes="${so.bandCols ? 1 : 0}" showLastColumn="1"/>`
    + (filterXml.length ? `<filters count="${filterXml.length}">${filterXml.join('')}</filters>` : '')
    + '</pivotTableDefinition>';

  // 슬라이서 캐시용: 필드 이름 → 항목 선택 상태 (x 는 캐시의 항목 번호)
  const slicerItems = (fieldName) => {
    const f = header.findIndex((h) => h.toLowerCase() === String(fieldName).toLowerCase());
    if (f < 0 || !cache.items.has(f)) return null;
    const others = filters.filter(([i]) => i !== f);
    const rowsOk = data.filter((r) => others.every(([i, set]) => set.has(itemText(r[i]))));
    const present = new Set(rowsOk.map((r) => `${typeof keyOf(r[f])}:${keyOf(r[f])}`));
    const keys = cache.items.get(f).keys;
    return {
      field: header[f],
      xml: keys.map((k, i) => `<i x="${i}"${hiddenKey(f, k) ? '' : ' s="1"'}${present.has(`${typeof k}:${k}`) ? '' : ' nd="1"'}/>`).join(''),
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

function slicerAnchorXml(sl, name, id, anchorAt, kind) {
  const EMUv = (px) => Math.round(px * EMU);
  const choice = kind === 'table'
    ? `<mc:Choice xmlns:sle15="http://schemas.microsoft.com/office/drawing/2012/slicer" Requires="sle15">`
    : `<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">`;
  const frame = `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="${esc(name)}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2010/slicer"><sle:slicer xmlns:sle="http://schemas.microsoft.com/office/drawing/2010/slicer" name="${esc(name)}"/></a:graphicData></a:graphic></xdr:graphicFrame>`;
  const fallback = `<mc:Fallback><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="0" name=""/><xdr:cNvSpPr><a:spLocks noTextEdit="1"/></xdr:cNvSpPr></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="${EMUv(sl.x)}" y="${EMUv(sl.y)}"/><a:ext cx="${EMUv(sl.w)}" cy="${EMUv(sl.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:prstClr val="white"/></a:solidFill><a:ln w="1"><a:solidFill><a:prstClr val="green"/></a:solidFill></a:ln></xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip"/><a:lstStyle/><a:p><a:r><a:rPr lang="ko-KR" sz="1100"/><a:t>이 도형은 ${kind === 'table' ? '표' : '피벗 테이블'} 슬라이서를 나타냅니다. 슬라이서는 Excel 2010 이상에서 지원됩니다.</a:t></a:r></a:p></xdr:txBody></xdr:sp></mc:Fallback>`;
  return `<xdr:twoCellAnchor editAs="${sl.placement ?? 'oneCell'}"><xdr:from>${anchorAt(sl.x, sl.y)}</xdr:from><xdr:to>${anchorAt(sl.x + sl.w, sl.y + sl.h)}</xdr:to><mc:AlternateContent xmlns:mc="${NS_MC}">${choice}${frame}</mc:Choice>${fallback}</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
}

/** .xlsx 로 저장할 때 엑셀 한도(1,048,576행)를 넘어 빠지는 셀 수 */
export function xlsxOverflow(wb) {
  let n = 0;
  for (const sheet of wb.sheets) {
    for (const [k, cell] of sheet.cells) if (cell.raw && Number(k.slice(0, k.indexOf(','))) >= EXCEL_MAX_ROWS) n++;
  }
  return n;
}

/** Workbook → xlsx 바이트 (매크로가 있으면 .xlsm 형식) */
/** 통합 문서 → .xlsx 바이트 (한 번에) */
export function writeXlsx(wb, opts) {
  const it = writeXlsxSteps(wb, opts);
  for (;;) {
    const s = it.next();
    if (s.done) return zip(s.value);
  }
}

/** 큰 문서용: 중간중간 브라우저에 제어를 돌려주고, 내장 압축으로 파일 크기도 줄임. onProgress({p, msg}) */
export async function writeXlsxAsync(wb, opts, onProgress) {
  const it = writeXlsxSteps(wb, opts);
  let last = performance.now();
  for (;;) {
    const s = it.next();
    if (s.done) return zipAsync(s.value, (p) => onProgress?.({ p: 0.85 + 0.15 * p, msg: '압축 중' }));
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

function* writeXlsxSteps(wb, { activeSheet = 0, fileName = 'Book1.xlsx', kind = null } = {}) {
  const files = {};
  const pool = new StylePool(wb.defaultFont ?? WRITE_FONT, wb.baseStyle);
  if (wb.cellStyles?.length) pool.namedStyles(wb.cellStyles);
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
  const slicerFields = new Map(); // 원본 키 → 슬라이서 필드 이름(소문자)
  wb.sheets.forEach((s, i) => (s.slicers ?? []).forEach((sl) => {
    for (const ref of slicerPivotRefs(wb, i, sl)) {
      const e = allDefs.find((x) => x.si === ref.si && (ref.name === null ? x.j === 0 : String(x.def.name ?? '') === ref.name));
      if (!e) continue;
      const k = pivotSourceKey(e.def);
      if (!slicerFields.has(k)) slicerFields.set(k, new Set());
      slicerFields.get(k).add(String(sl.source.field).toLowerCase());
    }
  }));
  const caches = new Map();
  let cacheSeq = 0;
  for (const e of allDefs) {
    const k = pivotSourceKey(e.def);
    if (!caches.has(k)) caches.set(k, { defs: [] });
    caches.get(k).defs.push(e.def);
  }
  for (const [k, g] of caches) {
    g.cache = buildPivotCache(wb, g.defs, ++cacheSeq, slicerFields.get(k));
    yield { p: 0.01, msg: '피벗 캐시 만드는 중' };
  }
  const pivotList = new Map(); // 시트 번호 → [{ name, cacheId, tableNo, parts, cache, j }]
  let pivotNo = 0;
  for (const e of allDefs) {
    const cache = caches.get(pivotSourceKey(e.def)).cache;
    if (!cache) continue;
    const used = new Set((pivotList.get(e.si) ?? []).map((x) => x.name.toLowerCase()));
    let name = e.def.name || `피벗 테이블${pivotNo + 1}`;
    for (let n = 2; used.has(name.toLowerCase()); n++) name = `${e.def.name || '피벗 테이블'}${n}`;
    const parts = pivotParts(wb, e.si, e.def, cache, name, pool);
    if (!parts) continue;
    pivotNo++;
    const info = { name, origName: e.def.name ?? null, cacheId: cache.cacheId, tableNo: pivotNo, parts, cache, j: e.j };
    if (!pivotList.has(e.si)) pivotList.set(e.si, []);
    pivotList.get(e.si).push(info);
  }
  const findPivotInfo = (ref) => (pivotList.get(ref.si) ?? []).find((x) => (ref.name === null ? x.j === 0 : x.origName === ref.name || x.name === ref.name));
  const pivotCaches = [];
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
      if (same === true) return esc(raw.slice(1));
      if (same === undefined) {
        const out = exportFormula(raw, t, dyn);
        astSame.set(ast, out === raw.slice(1));
        return esc(out);
      }
    }
    const k = `${t ?? ''}\u0001${dyn ? 1 : 0}\u0001${raw}`;
    let v = expMemo.get(k);
    if (v !== undefined) return v;
    const body = raw.startsWith('=') ? raw.slice(1) : raw;
    const shape = !t && !dyn ? formulaShape(body) : null;
    if (shape !== null && shapeSame.get(shape) === true) return esc(body);
    const out = exportFormula(raw, t, dyn);
    if (shape !== null && !shapeSame.has(shape)) shapeSame.set(shape, out === body);
    v = esc(out);
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
  for (let si = 0; si < nOwn; si++) {
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
    // 셀을 행별로 정리
    const rows = new Map();
    let maxR = 0;
    let maxC = 0;
    const eachCell = (fn) => (sheet.cells.forEachRC ? sheet.cells.forEachRC(fn) : sheet.cells.forEach((cell, k) => { const i = k.indexOf(','); fn(cell, +k.slice(0, i), +k.slice(i + 1)); }));
    eachCell((cell, r, c) => {
      if (r >= EXCEL_MAX_ROWS) return; // 엑셀 파일에는 1,048,576행까지만 저장 가능
      let list = rows.get(r);
      if (!list) { list = []; rows.set(r, list); }
      list.push([c, cell]);
      if (r > maxR) maxR = r;
      if (c > maxC) maxC = c;
    });
    // 동적 배열이 분산된 셀: 값만 저장 (엑셀도 같은 방식)
    for (const sp of wb.spillsOf(si)) {
      for (let i = 0; i < sp.h; i++) {
        const r = sp.r + i;
        if (r >= EXCEL_MAX_ROWS) break;
        for (let j = 0; j < sp.w; j++) {
          if (!i && !j) continue;
          const c = sp.c + j;
          if (sheet.cells.has(`${r},${c}`)) continue;
          if (!rows.has(r)) rows.set(r, []);
          rows.get(r).push([c, { raw: '', spilled: true }]);
          maxR = Math.max(maxR, r);
          maxC = Math.max(maxC, c);
        }
      }
    }
    const rowKeys = new Set([...rows.keys()]);
    for (const k of [...Object.keys(sheet.rowHeights), ...Object.keys(sheet.hiddenRows), ...Object.keys(sheet.rowStyles), ...Object.keys(sheet.outline?.rows ?? {}), ...Object.keys(sheet.outline?.rowsColl ?? {}), ...hidKeys(sheet.filter?.hidden, EXCEL_MAX_ROWS), ...(sheet.tables ?? []).flatMap((t) => hidKeys(t.filter?.hidden, EXCEL_MAX_ROWS))]) rowKeys.add(Number(k));
    // 열 블록의 행 (셀 객체 없이 형식화 배열에서 바로 씀)
    const blocks = sheet.blocks ?? [];
    for (const b of blocks) {
      for (let r = b.r0; r < Math.min(b.r0 + b.n, EXCEL_MAX_ROWS); r++) rowKeys.add(r);
      maxR = Math.max(maxR, Math.min(b.r0 + b.n, EXCEL_MAX_ROWS) - 1);
      maxC = Math.max(maxC, b.c0 + b.cols.length - 1);
    }
    const sortedRows = [...rowKeys].filter((r) => r < EXCEL_MAX_ROWS).sort((a, b) => a - b);

    const rowXml = [];
    for (const r of sortedRows) {
      let cells = rows.get(r) ?? [];
      for (const b of blocks) {
        if (r < b.r0 || r >= b.r0 + b.n) continue;
        const i = b.perm ? b.perm[r - b.r0] : r - b.r0;
        const taken = cells.length ? new Set(cells.map((x) => x[0])) : null;
        for (let j = 0; j < b.cols.length; j++) {
          const col = b.cols[j];
          const has = (col.str && col.str[i] >= 0) || (col.num && col.num[i] === col.num[i]);
          if (!has && !col.fmt) continue;
          if (taken?.has(b.c0 + j)) continue;
          cells.push([b.c0 + j, has ? { raw: '1', style: col.fmt ?? undefined } : { raw: '', style: col.fmt }]);
        }
      }
      cells = cells.sort((a, b) => a[0] - b[0]);
      const attrs = [`r="${r + 1}"`];
      if (sheet.rowHeights[r] !== undefined) attrs.push(`ht="${px2pt(sheet.rowHeights[r])}"`, 'customHeight="1"');
      if (sheet.hiddenRows[r] || hid(sheet.filter?.hidden, r) || (sheet.tables ?? []).some((t) => hid(t.filter?.hidden, r))) attrs.push('hidden="1"');
      if (sheet.rowStyles[r]) attrs.push(`s="${pool.xf({ ...sheet.allStyle, ...sheet.rowStyles[r] })}"`, 'customFormat="1"');
      if (sheet.outline?.rows?.[r]) attrs.push(`outlineLevel="${sheet.outline.rows[r]}"`);
      if (sheet.outline?.rowsColl?.[r]) attrs.push('collapsed="1"');
      const cx = cells.map(([c, cell]) => {
        const ref = refOf(r, c);
        const s = plainStyle ? pool.xf(cell.style ?? {}) : xfAt(r, c, cell);
        const sAttr = s ? ` s="${s}"` : '';
        const v = wb.getValue(si, r, c);
        if (!cell.raw && cell.image?.src) {
          const vm = richImage(cell.image);
          return vm ? `<c r="${ref}"${sAttr} t="e" vm="${vm}"><v>#VALUE!</v></c>` : (s ? `<c r="${ref}"${sAttr}/>` : '');
        }
        // 빈 칸: 기본 서식(xf 0)이어도 행 · 열 서식이 있으면 적어 둠 (없으면 다시 열 때 행 서식을 물려받음)
        if (!cell.raw && (v === null || !cell.spilled && !wb.spillAnchorOf(si, r, c))) return s || (cell.style && (sheet.rowStyles[r] || sheet.colStyles[c])) ? `<c r="${ref}" s="${s}"/>` : '';
        if (cell.formula) {
          // 배열을 돌려줄 수 있는 수식은 동적 배열 수식으로 (cm="1" + t="array")
          const dyn = !!cell.maybeArray;
          const sp = dyn ? wb.spillRange(si, r, c) : null;
          if (dyn) dynamicCells++;
          const fAttr = dyn ? ` t="array" ref="${sp ? rangeRef({ ...sp, r2: Math.min(sp.r2, EXCEL_MAX_ROWS - 1) }) : ref}" aca="false"` : '';
          const cm = dyn ? ' cm="1"' : '';
          const f = `<f${fAttr}>${exportF(cell.raw, cell.raw.includes('[') ? tableAt(sheet, r, c)?.name : null, dyn, cell.ast)}</f>`;
          if (typeof v === 'number') return `<c r="${ref}"${sAttr}${cm}>${f}<v>${v}</v></c>`;
          if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"${cm}>${f}<v>${v ? 1 : 0}</v></c>`;
          if (isError(v)) return `<c r="${ref}"${sAttr} t="e"${cm}>${f}<v>${esc(['#CIRC!', '#SPILL!', '#CALC!', '#BUSY!'].includes(v.code) && !dyn ? '#REF!' : v.code === '#CIRC!' ? '#REF!' : v.code)}</v></c>`;
          if (v instanceof CellImage) return `<c r="${ref}"${sAttr} t="e"${cm}>${f}<v>#VALUE!</v></c>`; // IMAGE: 엑셀이 다시 계산
          return `<c r="${ref}"${sAttr} t="str"${cm}>${f}<v>${esc(v ?? '')}</v></c>`;
        }
        if (isError(v)) return `<c r="${ref}"${sAttr} t="e"><v>${esc(v.code)}</v></c>`;
        if (v instanceof CellImage) return s ? `<c r="${ref}"${sAttr}/>` : '';
        if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
        if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
        if (v === null) return s ? `<c r="${ref}"${sAttr}/>` : '';
        return `<c r="${ref}"${sAttr} t="s"><v>${sst(String(v))}</v></c>`;
      }).join('');
      rowXml.push(`<row ${attrs.join(' ')}>${cx}</row>`);
      rowsDone += cells.length;
      if (rowXml.length % 1000 === 0) yield { p: 0.85 * (rowsDone / rowsTotal), msg: `'${sheet.name}' 시트 저장 중` };
    }
    sheetRows.push({ rowXml: rowXml.join(''), maxR, maxC });
  }

  for (let si = 0; si < nOwn; si++) {
    const sheet = wb.sheets[si];
    const sheetRels = [];
    const addRel = (type, target) => { const id = `rId${sheetRels.length + 1}`; sheetRels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`); return id; };

    const { rowXml, maxR, maxC } = sheetRows[si];

    // 열
    const olc = sheet.outline?.cols ?? {};
    const olcc = sheet.outline?.colsColl ?? {};
    const colKeys = new Set([...Object.keys(sheet.colWidths), ...Object.keys(sheet.hiddenCols), ...Object.keys(sheet.colStyles), ...Object.keys(olc), ...Object.keys(olcc)].map(Number));
    const colsXml = [...colKeys].sort((a, b) => a - b).map((c) => {
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
    if (sheet.page?.area) definedNames.push(`<definedName name="_xlnm.Print_Area" localSheetId="${si}">${esc(`${quoteSheetName(sheet.name)}!${rangeRef(sheet.page.area, true)}`)}</definedName>`);
    if (sheet.page?.titleRows || sheet.page?.titleCols) {
      const parts = [];
      if (sheet.page.titleCols) parts.push(`${quoteSheetName(sheet.name)}!$${colToName(sheet.page.titleCols[0])}:$${colToName(sheet.page.titleCols[1])}`);
      if (sheet.page.titleRows) parts.push(`${quoteSheetName(sheet.name)}!$${sheet.page.titleRows[0] + 1}:$${sheet.page.titleRows[1] + 1}`);
      definedNames.push(`<definedName name="_xlnm.Print_Titles" localSheetId="${si}">${esc(parts.join(','))}</definedName>`);
    }

    // 틀 고정
    const fr = sheet.freeze?.rows || 0;
    const fc = sheet.freeze?.cols || 0;
    let pane = '';
    if (fr || fc) {
      const activePane = fr && fc ? 'bottomRight' : fr ? 'bottomLeft' : 'topRight';
      pane = `<pane${fc ? ` xSplit="${fc}"` : ''}${fr ? ` ySplit="${fr}"` : ''} topLeftCell="${cellName(fr, fc)}" activePane="${activePane}" state="frozen"/><selection pane="${activePane}"/>`;
    }
    const dim = sheet.cells.size ? `A1:${cellName(maxR, maxC)}` : 'A1';

    // 필터
    let autoFilter = '';
    if (sheet.filter && sheet.filter.r1 < EXCEL_MAX_ROWS) {
      const f = { ...sheet.filter, r2: Math.min(sheet.filter.r2, EXCEL_MAX_ROWS - 1) };
      const cols = Object.entries(f.criteria ?? {}).filter(([, vals]) => Array.isArray(vals)).map(([c, vals]) => {
        const blank = vals.includes('');
        return `<filterColumn colId="${Number(c) - f.c1}"><filters${blank ? ' blank="1"' : ''}>${vals.filter((v) => v !== '').map((v) => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`;
      }).join('');
      autoFilter = `<autoFilter ref="${rangeRef(f)}">${cols}</autoFilter>`;
      definedNames.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${si}" hidden="1">${esc(`${quoteSheetName(sheet.name)}!${rangeRef(f, true)}`)}</definedName>`);
    }
    const fit = (rg) => (rg.r1 >= EXCEL_MAX_ROWS ? null : { ...rg, r2: Math.min(rg.r2, EXCEL_MAX_ROWS - 1) });
    const mergeList = sheet.merges.map(fit).filter(Boolean);
    const merges = mergeList.length ? `<mergeCells count="${mergeList.length}">${mergeList.map((m) => `<mergeCell ref="${rangeRef(m)}"/>`).join('')}</mergeCells>` : '';
    const cfX14 = [];
    const cf = sheet.cond.map((rule) => {
      const g = fit(rule);
      return g && { ...g, ...(rule.more ? { more: rule.more.map(fit).filter(Boolean) } : {}) };
    }).filter(Boolean).map((rule, i) => cfXml(rule, pool, i + 1, cfX14)).join('');

    // 그림 개체 (차트 · 그림 · 도형)
    let drawing = '';
    const images = sheet.images ?? [];
    const shapes = sheet.shapes ?? [];
    // 슬라이서: 이름·캐시를 정하고 그림 개체와 함께 그림
    const sheetSlicers = { table: [], pivot: [] };
    for (const sl of sheet.slicers ?? []) {
      let cacheXml = null;
      let kind = null;
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
        sl._cache = cacheNameFor(sl.source.column, usedCacheNames);
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(sl._cache)}" sourceName="${esc(columnNames(wb, f.si, f.t)[col])}"><extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${NS_X15}"><x15:tableSlicerCache tableId="${tid}" column="${col + 1}"${setAttrs}/></x:ext>${hideExt}</extLst></slicerCacheDefinition>`;
      } else {
        // 연결된 피벗 중 첫 피벗과 같은 캐시를 쓰는 것만 (엑셀 규칙)
        const refs = slicerPivotRefs(wb, si, sl).map((ref) => ({ ref, info: findPivotInfo(ref) })).filter((x) => x.info);
        const first = refs[0]?.info;
        const it = first?.parts.slicerItems(sl.source.field);
        if (!it) continue;
        const linked = refs.filter((x) => x.info.cacheId === first.cacheId);
        kind = 'pivot';
        sl._cache = cacheNameFor(it.field, usedCacheNames);
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(sl._cache)}" sourceName="${esc(it.field)}"><pivotTables>${linked.map((x) => `<pivotTable tabId="${x.ref.si + 1}" name="${esc(x.info.name)}"/>`).join('')}</pivotTables><data><tabular pivotCacheId="${first.cacheId}"${setAttrs}${sl.showDeleted ? '' : ' showMissing="0"'}><items count="${it.count}">${it.xml}</items></tabular></data>${hideExt ? `<extLst>${hideExt}</extLst>` : ''}</slicerCacheDefinition>`;
      }
      slicerCacheNo++;
      files[`xl/slicerCaches/slicerCache${slicerCacheNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${cacheXml}`;
      contentOverrides.push(`<Override PartName="/xl/slicerCaches/slicerCache${slicerCacheNo}.xml" ContentType="application/vnd.ms-excel.slicerCache+xml"/>`);
      (kind === 'table' ? slicerCachesTable : slicerCachesPivot).push(`slicerCaches/slicerCache${slicerCacheNo}.xml`);
      definedNames.push(`<definedName name="${esc(sl._cache)}">#N/A</definedName>`);
      let nm = sl.caption || sl.source.column || sl.source.field || '슬라이서';
      for (let i = 1; usedSlicerNames.has(nm.toLowerCase()); i++) nm = `${sl.caption || '슬라이서'} ${i}`;
      usedSlicerNames.add(nm.toLowerCase());
      sheetSlicers[kind].push({ sl, name: nm, cache: sl._cache });
      delete sl._cache;
    }
    const hasSlicers = sheetSlicers.table.length + sheetSlicers.pivot.length > 0;
    if (sheet.charts.length || images.length || shapes.length || hasSlicers) {
      drawingNo++;
      const drawingRels = [];
      const drel = (type, target) => { const id = `rId${drawingRels.length + 1}`; drawingRels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`); return id; };
      const { colAxis, rowAxis } = objectAxes(sheet);
      const anchorAt = (x, y) => {
        const c = colAxis.indexAt(Math.max(0, x));
        const r = rowAxis.indexAt(Math.max(0, y));
        return `<xdr:col>${c}</xdr:col><xdr:colOff>${Math.max(0, Math.round((x - colAxis.pos(c)) * EMU))}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${Math.max(0, Math.round((y - rowAxis.pos(r)) * EMU))}</xdr:rowOff>`;
      };
      const anchor = (o, body) => `<xdr:twoCellAnchor editAs="${o.placement ?? 'twoCell'}"><xdr:from>${anchorAt(o.x, o.y)}</xdr:from><xdr:to>${anchorAt(o.x + o.w, o.y + o.h)}</xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`;
      const xfrm = (o) => `<a:xfrm${o.rot ? ` rot="${Math.round(o.rot * 60000)}"` : ''}${o.flip ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}><a:off x="${Math.round(o.x * EMU)}" y="${Math.round(o.y * EMU)}"/><a:ext cx="${Math.round(o.w * EMU)}" cy="${Math.round(o.h * EMU)}"/></a:xfrm>`;
      let objId = 1;
      const parts = [];
      const ordered = [
        ...sheet.charts.map((o) => ['chart', o]), ...images.map((o) => ['image', o]), ...shapes.map((o) => ['shape', o]),
        ...sheetSlicers.table.map((o) => ['slicerTable', o]), ...sheetSlicers.pivot.map((o) => ['slicerPivot', o]),
      ].sort((a, b) => ((a[1].sl ?? a[1]).z ?? 0) - ((b[1].sl ?? b[1]).z ?? 0));
      for (const [kind, o] of ordered) {
        if (kind === 'chart') {
          const ch = o;
          chartNo++;
          objId++;
          files[`xl/charts/chart${chartNo}.xml`] = chartXml(wb, si, ch, fileName);
          contentOverrides.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          const id = drel('chart', `../charts/chart${chartNo}.xml`);
          parts.push(anchor(ch, `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${objId}" name="차트 ${objId - 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${id}"/></a:graphicData></a:graphic></xdr:graphicFrame>`));
        } else if (kind === 'image') {
          const im = o;
          // SVG 그림(아이콘): PNG 대체 그림 + svgBlip 으로 원본 SVG (엑셀과 같은 방식)
          const svgSrc = im.png && /^data:image\/svg\+xml;base64,/.test(im.src ?? '') ? im.src : null;
          const m = /^data:([^;,]+);base64,(.*)$/s.exec(im.emf ?? (svgSrc ? im.png : im.src) ?? '');
          if (!m) continue;
          const ext = Object.keys(MIME).find((k) => MIME[k] === m[1]) ?? 'png';
          mediaNo++;
          mediaExts.add(ext);
          files[`xl/media/image${mediaNo}.${ext}`] = fromBase64(m[2]);
          const id = drel('image', `../media/image${mediaNo}.${ext}`);
          let svgExt = '';
          if (svgSrc) {
            mediaNo++;
            mediaExts.add('svg');
            files[`xl/media/image${mediaNo}.svg`] = fromBase64(svgSrc.slice(svgSrc.indexOf(',') + 1));
            const sid = drel('image', `../media/image${mediaNo}.svg`);
            svgExt = `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}"><asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="${sid}"/></a:ext></a:extLst>`;
          }
          objId++;
          parts.push(anchor(im, `<xdr:pic${im.macro ? ` macro="[0]!${esc(im.macro)}"` : ''}><xdr:nvPicPr><xdr:cNvPr id="${objId}" name="${esc(im.name || `그림 ${objId - 1}`)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill>${svgExt ? `<a:blip r:embed="${id}">${svgExt}</a:blip>` : `<a:blip r:embed="${id}"/>`}${im.crop ? `<a:srcRect${['l', 't', 'r', 'b'].map((k) => (im.crop[k] ? ` ${k}="${Math.round(im.crop[k] * 100000)}"` : '')).join('')}/>` : ''}<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(im)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${/^#[0-9a-f]{6}$/i.test(im.border ?? "") ? `<a:ln w="${Math.round((im.borderW ?? 2) * EMU)}"><a:solidFill><a:srgbClr val="${im.border.replace('#', '').toUpperCase()}"/></a:solidFill></a:ln>` : ''}</xdr:spPr></xdr:pic>`));
        } else if (kind === 'slicerTable' || kind === 'slicerPivot') {
          objId++;
          parts.push(slicerAnchorXml(o.sl, o.name, objId, anchorAt, kind === 'slicerTable' ? 'table' : 'pivot'));
        } else {
          objId++;
          parts.push(anchor(o, shapeXml(o, objId, xfrm)));
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
    for (const [k, cell] of sheet.cells) {
      if (!cell.link) continue;
      const [r, c] = k.split(',').map(Number);
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
    const comments = [...sheet.cells].filter(([, c]) => c.comment).map(([k, c]) => [k.split(',').map(Number), c.comment]);
    if (comments.length) {
      commentNo++;
      files[`xl/comments${commentNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${NS_MAIN}"><authors><author>WIXEL</author></authors><commentList>${comments.map(([[r, c], text]) => `<comment ref="${cellName(r, c)}" authorId="0"><text><r><t xml:space="preserve">${esc(text)}</t></r></text></comment>`).join('')}</commentList></comments>`;
      files[`xl/drawings/vmlDrawing${commentNo}.vml`] = `<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"><o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${commentNo}"/></o:shapelayout><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>${comments.map(([[r, c]], i) => `<v:shape id="_x0000_s${commentNo * 1024 + i + 1}" type="#_x0000_t202" style="position:absolute;margin-left:80pt;margin-top:2pt;width:108pt;height:59pt;z-index:${i + 1};visibility:hidden" fillcolor="#ffffe1" o:insetmode="auto"><v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox><x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>${c + 1}, 15, ${Math.max(0, r - 1)}, 10, ${c + 3}, 15, ${r + 3}, 4</x:Anchor><x:AutoFill>False</x:AutoFill><x:Row>${r}</x:Row><x:Column>${c}</x:Column></x:ClientData></v:shape>`).join('')}</xml>`;
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
        const fn = t.totals ? t.totalsFns?.[c] : null;
        let extra = '';
        let inner = '';
        const tot = t.totals ? wb.getCell(si, t.r2, c) : null;
        if (fn && fn !== 'none') extra = ` totalsRowFunction="${fn}"`;
        else if (tot?.formula) {
          // 요약 행의 사용자 수식
          extra = ' totalsRowFunction="custom"';
          inner += `<totalsRowFormula>${esc(exportFormula(tot.raw, t.name))}</totalsRowFormula>`;
        } else if (tot?.raw) extra = ` totalsRowLabel="${esc(tot.raw.replace(/^'/, ''))}"`;
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
      let af = '';
      if (t.filter && t.header) {
        const fr = { r1: t.r1, c1: t.c1, r2: dataBottom(t), c2: t.c2 };
        const fcs = Object.entries(t.filter.criteria ?? {}).filter(([, vals]) => Array.isArray(vals)).map(([c, vals]) => {
          const blank = vals.includes('');
          return `<filterColumn colId="${Number(c) - t.c1}"><filters${blank ? ' blank="1"' : ''}>${vals.filter((v) => v !== '').map((v) => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`;
        }).join('');
        af = `<autoFilter ref="${rangeRef(fr)}">${fcs}</autoFilter>`;
      }
      const style = t.style && t.style !== 'None' ? t.style : null;
      if (style && isModernStyle(style)) pool.presetTableStyle(style, false);
      files[`xl/tables/table${tableNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${NS_MAIN}" id="${tableNo}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${rangeRef(t)}"${t.header ? '' : ' headerRowCount="0"'}${t.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>${af}<tableColumns count="${names.length}">${cols}</tableColumns><tableStyleInfo${style ? ` name="${style}"` : ''} showFirstColumn="${t.firstCol ? 1 : 0}" showLastColumn="${t.lastCol ? 1 : 0}" showRowStripes="${t.banded !== false ? 1 : 0}" showColumnStripes="${t.bandedCols ? 1 : 0}"/></table>`;
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
      files[`xl/slicers/slicer${slicerPartNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">${list.map(({ sl, name, cache }) => `<slicer name="${esc(name)}" cache="${esc(cache)}" caption="${esc(sl.caption ?? name)}"${(sl.columns ?? 1) > 1 ? ` columnCount="${sl.columns}"` : ''}${((st) => (st !== 'SlicerStyleLight1' ? ` style="${esc(st)}"` : ''))(pool.slicerStyleFor(sl))}${sl.showHeader === false ? ' showCaption="0"' : ''}${sl.noMove ? ' lockedPosition="1"' : ''} rowHeight="${Math.round((sl.buttonHeight ?? 25.3) * EMU)}"/>`).join('')}</slicers>`;
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
    const extLst = exts.length ? `<extLst>${exts.join('')}</extLst>` : '';

    const tabOk = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.test(String(sheet.tabColor ?? ''));
    files[`xl/worksheets/sheet${si + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`
      + (vba || olPr || tabOk ? `<sheetPr${vba ? ` codeName="${esc(vba.sheetCodes?.[sheet.name] ?? `Sheet${si + 1}`)}"` : ''}>${tabOk ? `<tabColor rgb="${argb(sheet.tabColor)}"/>` : ''}${olPr}</sheetPr>` : '')
      + `<dimension ref="${dim}"/>`
      + `<sheetViews><sheetView${sheet.noGrid ? ' showGridLines="0"' : ''}${sheet.zoom && sheet.zoom !== 100 ? ` zoomScale="${sheet.zoom}" zoomScaleNormal="${sheet.zoom}"` : ''}${sheet.view && (sheet.view.top || sheet.view.left) ? ` topLeftCell="${cellName(sheet.view.top, sheet.view.left)}"` : ''} workbookViewId="0"${si === (wb.sheets[activeSheet]?.state && wb.sheets[activeSheet].state !== 'visible' ? Math.max(0, wb.sheets.findIndex((x) => !x.state || x.state === 'visible')) : activeSheet) ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>`
      + `<sheetFormatPr defaultColWidth="${px2widthM(sheet.defColW ?? DEFAULT_COL_WIDTH, wmdw)}" defaultRowHeight="${px2pt(sheet.defRowH ?? DEFAULT_ROW_HEIGHT)}"${sheet.defRowH ? ' customHeight="1"' : ''}${olRowMax ? ` outlineLevelRow="${olRowMax}"` : ''}${olColMax ? ` outlineLevelCol="${olColMax}"` : ''}/>`
      + (colsXml ? `<cols>${colsXml}</cols>` : '')
      + `<sheetData>${rowXml}</sheetData>`
      + protectXml(sheet.protect)
      + scenariosXml(sheet.scenarios)
      + autoFilter + merges + cf + dataValidations + hyperlinks
      + pgx.printOptions + pgx.margins + pgx.setup + pgx.headerFooter
      + drawing + legacy + tableParts + extLst
      + '</worksheet>';
    if (sheetRels.length) {
      files[`xl/worksheets/_rels/sheet${si + 1}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${sheetRels.join('')}</Relationships>`;
    }
    yield { p: 0.85, msg: '파일 구성 중' };
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
  files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${wb.props?.readOnlyRecommended ? '<fileSharing readOnlyRecommended="1"/>' : ''}${vba ? `<workbookPr codeName="${esc(vba.codeName || 'ThisWorkbook')}"/>` : ''}${wb.props?.lockStructure ? '<workbookProtection lockStructure="1"/>' : ''}<bookViews><workbookView${firstVisible ? ` firstSheet="${firstVisible}"` : ''} activeTab="${activeTab}"/></bookViews><sheets>${wb.sheets.slice(0, nOwn).map((sh, i) => `<sheet name="${esc(sh.name)}" sheetId="${i + 1}"${sh.state === 'hidden' || sh.state === 'veryHidden' ? ` state="${sh.state}"` : ''} r:id="rId${i + 1}"/>`).join('')}</sheets>${extRefsXml}${definedNames.length ? `<definedNames>${definedNames.join('')}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/>${pivotCachesXml}${wbExts.length ? `<extLst>${wbExts.join('')}</extLst>` : ''}</workbook>`;
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${wbRels.join('')}</Relationships>`;
  if (vba) files['xl/vbaProject.bin'] = fromBase64(vba.bin);
  files['xl/sharedStrings.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`;
  files['xl/styles.xml'] = pool.xml();
  files['xl/theme/theme1.xml'] = themeXml(wb);
  contentOverrides.push('<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>');
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  const pr = wb.props ?? {};
  const tagX = (tag, v) => (v ? `<${tag}>${esc(String(v))}</${tag}>` : '');
  files['docProps/core.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">${tagX('dc:title', pr.title)}${tagX('dc:subject', pr.subject)}${tagX('dc:creator', pr.creator || 'WIXEL')}${tagX('cp:keywords', pr.tags)}${tagX('dc:description', pr.comments)}${tagX('cp:lastModifiedBy', pr.lastModifiedBy)}${tagX('cp:category', pr.category)}<dcterms:created xsi:type="dcterms:W3CDTF">${/^\d{4}-\d\d-\d\dT/.test(pr.created ?? '') ? esc(pr.created) : now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  if (pr.markedFinal) {
    files['docProps/custom.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="_MarkAsFinal"><vt:bool>true</vt:bool></property></Properties>`;
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
