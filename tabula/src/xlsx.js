// .xlsx 읽기/쓰기 (Office Open XML). DOM 없이 동작하므로 Node 에서도 테스트 가능.
import { unzip, zip, textOf } from './zip.js';
import { parseXml, child, kids, descendants, allText, esc } from './xml.js';
import {
  parse, tokenize, shiftFormula, colToName, nameToCol, cellName, parseRangeName, FUNCS, isError,
  quoteSheetName, MAX_ROWS, MAX_COLS, EXCEL_MAX_ROWS, mayReturnArray, unknownFunctions,
} from './formula.js';
import { toFileFormula, fromFileFormula } from './xlfn.js';
import { parseInput, formatGeneral, fmtCode, styleForCode } from './format.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { chartLayout, PALETTE } from './chart.js';
import { Axis } from './axis.js';
import { toBase64, fromBase64 } from './vba.js';
import { normalizeStyleName, DEFAULT_TABLE_STYLE, dataTop, dataBottom, canonicalRef, tableAt, columnNames, findTable } from './tables.js';
import { pivotSourceData, resolvePivot, itemText, keyOf, sortKeys, EMPTY, headerNames, normalizeDef, computePivot, valueName } from './pivot.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU = 9525; // 1px

const DEFAULT_FONT = '맑은 고딕';

// ───────────────────────── 공통 ─────────────────────────
const px2width = (px) => Math.max(0, Math.round(((px - 5) / 7) * 256) / 256);
const width2px = (w) => Math.max(0, Math.round(w * 7 + 5));
const pt2px = (pt) => Math.round((pt * 4) / 3);
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
function textRaw(s) {
  if (s === '') return '';
  if (s.startsWith('=') || s.startsWith("'")) return `'${s}`;
  const p = parseInput(s);
  return typeof p.value === 'string' && p.value === s ? s : `'${s}`;
}

// ───────────────────────── 색상 ─────────────────────────
const INDEXED = ('000000,FFFFFF,FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,000000,FFFFFF,FF0000,00FF00,0000FF,FFFF00,FF00FF,00FFFF,'
  + '800000,008000,000080,808000,800080,008080,C0C0C0,808080,9999FF,993366,FFFFCC,CCFFFF,660066,FF8080,0066CC,CCCCFF,'
  + '000080,FF00FF,FFFF00,00FFFF,800080,800000,008080,0000FF,00CCFF,CCFFFF,CCFFCC,FFFF99,99CCFF,FF99CC,CC99FF,FFCC99,'
  + '3366FF,33CCCC,99CC00,FFCC00,FF9900,FF6600,666699,969696,003366,339966,003300,333300,993300,993366,333399,333333,000000,FFFFFF').split(',');
const DEFAULT_THEME = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47'];

function applyTint(hex, tint) {
  if (!tint) return hex;
  let [r, g, b] = [0, 2, 4].map((i) => parseInt(hex.slice(i, i + 2), 16) / 255);
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let h = 0;
  let s = 0;
  let l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const hue = (p, q, t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  if (s === 0) { r = l; g = l; b = l; } else {
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue(p, q, h + 1 / 3); g = hue(p, q, h); b = hue(p, q, h - 1 / 3);
  }
  return [r, g, b].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('').toUpperCase();
}

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
const BUILTIN_FMT = {
  0: {}, 1: { decimals: 0 }, 2: { decimals: 2 }, 3: { numFmt: 'comma' }, 4: { numFmt: 'number', decimals: 2 },
  9: { numFmt: 'percent' }, 10: { numFmt: 'percent', decimals: 2 }, 11: { numFmt: 'scientific' },
  12: { numFmt: 'fraction' }, 13: { numFmt: 'fraction' }, 14: { numFmt: 'date' }, 15: { numFmt: 'date' },
  16: { numFmt: 'date' }, 17: { numFmt: 'date' }, 18: { numFmt: 'time' }, 19: { numFmt: 'time' },
  20: { numFmt: 'time' }, 21: { numFmt: 'time' }, 22: { numFmt: 'datetime' }, 37: { numFmt: 'comma' },
  38: { numFmt: 'comma' }, 39: { numFmt: 'number', decimals: 2 }, 40: { numFmt: 'number', decimals: 2 },
  44: { numFmt: 'accounting' }, 45: { numFmt: 'time' }, 46: { numFmt: 'time' }, 47: { numFmt: 'time' },
  48: { numFmt: 'scientific', decimals: 1 }, 49: { numFmt: 'text' },
};
for (const id of [27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 57, 58]) BUILTIN_FMT[id] = { numFmt: 'date' };
for (const id of [32, 33, 55, 56]) BUILTIN_FMT[id] = { numFmt: 'time' };

// ───────────────────────── 읽기 ─────────────────────────
function relsOf(files, path) {
  const dir = path.slice(0, path.lastIndexOf('/') + 1);
  const relPath = `${dir}_rels/${path.slice(dir.length)}.rels`;
  const xml = textOf(files[relPath]);
  const out = {};
  if (!xml) return out;
  for (const r of kids(parseXml(xml), 'Relationship')) {
    let target = r.attrs.Target;
    if (r.attrs.TargetMode === 'External') continue;
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
    if (color && color !== '#000000') st.color = color;
    const name = child(f, 'name')?.attrs.val;
    if (name) st.font = name;
    return st;
  };
  const fonts = kids(child(root, 'fonts'), 'font').map(fontOf);
  const defaultFont = fonts[0]?.font ?? null;
  const fillOf = (f, dxf = false) => {
    const pf = child(f, 'patternFill');
    if (!pf) {
      const stop = descendants(f, 'stop')[0];
      return stop ? colorOf(child(stop, 'color'), theme) : null;
    }
    if (!dxf && (pf.attrs.patternType === 'none' || !pf.attrs.patternType)) return null;
    return colorOf(child(pf, dxf ? 'bgColor' : 'fgColor') ?? child(pf, 'fgColor'), theme);
  };
  const fills = kids(child(root, 'fills'), 'fill').map((f) => fillOf(f));
  const borderOf = (b) => {
    const st = {};
    const side = (n) => { const e = child(b, n); return e && e.attrs.style && e.attrs.style !== 'none'; };
    if (side('top')) st.bt = true;
    if (side('bottom')) st.bb = true;
    if (side('left') || side('start')) st.bl = true;
    if (side('right') || side('end')) st.br = true;
    return st;
  };
  const borders = kids(child(root, 'borders'), 'border').map(borderOf);
  const numFmtOf = (id) => {
    const n = Number(id);
    if (BUILTIN_FMT[n]) return BUILTIN_FMT[n];
    if (numFmts[id]) return styleForCode(numFmts[id]);
    return {};
  };
  const xfs = kids(child(root, 'cellXfs'), 'xf').map((xf) => {
    const a = xf.attrs;
    const st = { ...fonts[Number(a.fontId || 0)] };
    if (st.font && st.font === defaultFont) delete st.font;
    if (st.size === 11) delete st.size;
    const fill = fills[Number(a.fillId || 0)];
    if (fill) st.fill = fill;
    Object.assign(st, borders[Number(a.borderId || 0)] ?? {});
    Object.assign(st, numFmtOf(a.numFmtId || 0));
    const al = child(xf, 'alignment');
    if (al) {
      const h = al.attrs.horizontal;
      if (h === 'left' || h === 'center' || h === 'right') st.align = h;
      else if (h === 'centerContinuous' || h === 'distributed') st.align = 'center';
      const v = al.attrs.vertical;
      if (v === 'top' || v === 'center') st.valign = v === 'center' ? 'middle' : 'top';
      if (al.attrs.wrapText === '1' || al.attrs.wrapText === 'true') st.wrap = true;
      if (Number(al.attrs.indent)) st.indent = Number(al.attrs.indent);
    }
    return st;
  });
  const dxfs = kids(child(root, 'dxfs'), 'dxf').map((d) => {
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
  });
  return { xfs, dxfs, defaultFont };
}

/** 파일 수식 → 앱 수식 본문 (_xlfn. 등 접두사 제거, SINGLE → @, ANCHORARRAY → #) */
function cleanFormula(f, opt = {}) {
  return fromFileFormula(f, opt);
}

function readSheet(files, path, ctx) {
  const root = parseXml(textOf(files[path]));
  const sheet = {
    cells: {}, colWidths: {}, rowHeights: {}, merges: [], cond: [], colStyles: {}, rowStyles: {},
    hiddenRows: {}, hiddenCols: {}, rowManual: {}, freeze: { rows: 0, cols: 0 }, filter: null, charts: [], images: [], shapes: [], validations: [], slicers: [],
  };
  const { xfs, dxfs, strings } = ctx;
  const styleOf = (s) => {
    const st = xfs[Number(s || 0)];
    return st && Object.keys(st).length ? { ...st } : undefined;
  };

  for (const col of kids(child(root, 'cols'), 'col')) {
    const min = Number(col.attrs.min) - 1;
    const max = Math.min(Number(col.attrs.max) - 1, min + 16384);
    const w = col.attrs.width !== undefined ? width2px(Number(col.attrs.width)) : null;
    const st = col.attrs.style ? styleOf(col.attrs.style) : undefined;
    for (let c = min; c <= max && c < MAX_COLS; c++) {
      if (w !== null && w !== DEFAULT_COL_WIDTH && (col.attrs.customWidth === '1' || Math.abs(w - DEFAULT_COL_WIDTH) > 1)) sheet.colWidths[c] = w;
      if (col.attrs.hidden === '1' || col.attrs.hidden === 'true') sheet.hiddenCols[c] = true;
      if (st && max - min < 1000) sheet.colStyles[c] = st;
    }
  }

  const shared = {};
  const arrays = []; // 배열 수식 영역: 앵커 밖의 셀 값은 가져오지 않음 (다시 분산됨)
  let unsupported = 0;
  const data = child(root, 'sheetData');
  let rowIdx = -1;
  for (const row of kids(data, 'row')) {
    rowIdx = row.attrs.r ? Number(row.attrs.r) - 1 : rowIdx + 1;
    const r = rowIdx;
    if (row.attrs.ht && (row.attrs.customHeight === '1' || row.attrs.customHeight === 'true')) {
      const h = pt2px(Number(row.attrs.ht));
      if (h !== DEFAULT_ROW_HEIGHT) { sheet.rowHeights[r] = h; sheet.rowManual[r] = true; }
    } else if (row.attrs.ht) {
      const h = pt2px(Number(row.attrs.ht));
      if (Math.abs(h - DEFAULT_ROW_HEIGHT) > 2) sheet.rowHeights[r] = h;
    }
    if (row.attrs.hidden === '1' || row.attrs.hidden === 'true') sheet.hiddenRows[r] = true;
    if (row.attrs.customFormat === '1' && row.attrs.s) { const st = styleOf(row.attrs.s); if (st) sheet.rowStyles[r] = st; }
    let colIdx = -1;
    for (const c of kids(row, 'c')) {
      let cc;
      if (c.attrs.r) {
        const m = /^([A-Z]+)(\d+)$/i.exec(c.attrs.r);
        cc = nameToCol(m[1]);
      } else cc = colIdx + 1;
      colIdx = cc;
      const t = c.attrs.t ?? 'n';
      const vEl = child(c, 'v');
      const fEl = child(c, 'f');
      const style = styleOf(c.attrs.s);
      let raw = '';
      let value = null;
      if (t === 's') value = strings[Number(vEl?.text)] ?? '';
      else if (t === 'inlineStr') value = allText(child(c, 'is'));
      else if (t === 'str') value = vEl?.text ?? '';
      else if (t === 'b') value = vEl?.text === '1';
      else if (t === 'e') value = { error: vEl?.text ?? '#N/A' };
      else if (vEl && vEl.text !== '') value = Number(vEl.text);

      let formula = null;
      if (fEl) {
        if (fEl.attrs.t === 'shared' && fEl.attrs.si !== undefined) {
          if (fEl.text) shared[fEl.attrs.si] = { text: fEl.text, r, c: cc };
          const m = shared[fEl.attrs.si];
          if (m) formula = m.r === r && m.c === cc ? m.text : shiftFormula(`=${m.text}`, r - m.r, cc - m.c).slice(1);
        } else if (fEl.text) formula = fEl.text;
      }
      let cached;
      if (formula !== null) {
        const isArray = fEl.attrs.t === 'array';
        if (isArray && fEl.attrs.ref) {
          const rg = refToRange(fEl.attrs.ref);
          if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) arrays.push({ ...rg, r, c: cc });
        }
        // 동적 배열(cm) 또는 배열 수식이 아니면 옛 형식 → 암시적 교차 '@'
        const f = cleanFormula(formula, { legacy: !isArray && !c.attrs.cm, isName: ctx.isName, nameMulti: ctx.nameMulti });
        raw = `=${f}`;
        if (unknownFunctions(f, ctx.isName).length) {
          // 지원하지 않는 함수: 수식은 그대로 두고 파일의 계산 결과를 표시
          unsupported++;
          cached = value;
        }
      } else if (arrays.length && arrays.some((a) => r >= a.r1 && r <= a.r2 && cc >= a.c1 && cc <= a.c2 && (r !== a.r || cc !== a.c))) {
        value = null;
      }
      if (!raw) {
        if (value === null) raw = '';
        else if (typeof value === 'boolean') raw = value ? 'TRUE' : 'FALSE';
        else if (typeof value === 'object') raw = value.error;
        else if (typeof value === 'string') raw = style?.numFmt === 'text' ? value : textRaw(value);
        else raw = numberRaw(value, style);
      }
      if (!raw && !style) continue;
      const d = { raw };
      if (style) d.style = style;
      if (cached !== undefined && cached !== null) d.cached = cached;
      sheet.cells[`${r},${cc}`] = d;
    }
  }
  sheet.unsupported = unsupported;

  for (const m of kids(child(root, 'mergeCells'), 'mergeCell')) {
    const rg = refToRange(m.attrs.ref);
    if (rg && (rg.r2 > rg.r1 || rg.c2 > rg.c1)) sheet.merges.push(rg);
  }

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
        let out = null;
        switch (a.type) {
          case 'cellIs': if (OPS[a.operator]) out = { type: OPS[a.operator], v1: fval(formulas[0]), ...(formulas[1] !== undefined ? { v2: fval(formulas[1]) } : {}), style }; break;
          case 'containsText': out = { type: 'text', v1: a.text ?? '', style }; break;
          case 'notContainsText': out = { type: 'notText', v1: a.text ?? '', style }; break;
          case 'beginsWith': out = { type: 'begins', v1: a.text ?? '', style }; break;
          case 'endsWith': out = { type: 'ends', v1: a.text ?? '', style }; break;
          case 'containsBlanks': out = { type: 'blank', style }; break;
          case 'notContainsBlanks': out = { type: 'noBlank', style }; break;
          case 'containsErrors': out = { type: 'errors', style }; break;
          case 'notContainsErrors': out = { type: 'noErrors', style }; break;
          case 'timePeriod': out = { type: 'date', period: a.timePeriod ?? 'today', style }; break;
          case 'expression': if (formulas[0]) out = { type: 'formula', formula: `=${cleanFormula(formulas[0])}`, style }; break;
          case 'duplicateValues': out = { type: 'dup', style }; break;
          case 'uniqueValues': out = { type: 'unique', style }; break;
          case 'top10': out = { type: a.bottom === '1' ? 'bottom' : 'top', v1: a.rank ?? '10', ...(a.percent === '1' ? { percent: true } : {}), style }; break;
          case 'aboveAverage': {
            out = { type: a.aboveAverage === '0' ? 'belowAvg' : 'aboveAvg', style };
            if (a.equalAverage === '1') out.equal = true;
            if (a.stdDev) out.stdDev = Number(a.stdDev);
            break;
          }
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
    else if (rule.attrs.type === 'dataBar') {
      const xdb = child(rule, 'dataBar');
      const color = colorOf(child(xdb, 'fillColor'), ctx.theme) ?? '#638ec6';
      out = { type: 'bar', color: color.toLowerCase(), ...(xdb?.attrs.gradient === '0' ? { gradient: false } : {}) };
    }
    if (out) condList.push({ p: Number(rule.attrs.priority ?? 1e9), i: condList.length, rule: { ...ranges[0], ...(ranges.length > 1 ? { more: ranges.slice(1) } : {}), ...out } });
  }
  condList.sort((x, y) => x.p - y.p || x.i - y.i);
  sheet.cond = condList.map((x) => x.rule);

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
        sheet.cells[k] = { raw: '', ...sheet.cells[k], comment: text };
      }
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
        if (sl.attrs.cache) sheet._slicers.push({ name: sl.attrs.name, cache: sl.attrs.cache, caption: sl.attrs.caption, columns: Number(sl.attrs.columnCount ?? 1), style: sl.attrs.style });
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
  return sheet;
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

function numberRaw(v, style) {
  const fmt = style?.numFmt;
  if ((fmt === 'date' || fmt === 'longdate') && Number.isInteger(v) && v > 0) {
    const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, '0')}-${String(d.getUTCDate()).padStart(2, '0')}`;
  }
  if (fmt === 'percent') {
    const s = `${Number((v * 100).toPrecision(15))}%`;
    const p = parseInput(s);
    if (typeof p.value === 'number' && Math.abs(p.value - v) < 1e-12) return s;
  }
  return String(v);
}

const SCHEME_INDEX = { lt1: 0, bg1: 0, dk1: 1, tx1: 1, lt2: 2, bg2: 2, dk2: 3, tx2: 3, accent1: 4, accent2: 5, accent3: 6, accent4: 7, accent5: 8, accent6: 9 };
const PRST_KIND = {
  rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', rtTriangle: 'triangle',
  rightArrow: 'arrow', leftArrow: 'arrow', line: 'line', straightConnector1: 'line', bentConnector3: 'line',
  flowChartProcess: 'rect', flowChartAlternateProcess: 'roundRect', flowChartConnector: 'ellipse', wedgeRectCallout: 'rect',
};
const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp' };

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

function readDrawing(files, path, sheet, ctx) {
  const out = { charts: [], images: [], shapes: [], _slicerBoxes: {} };
  let z = 0; // 겹치는 순서
  const xml = textOf(files[path]);
  if (!xml) return out;
  const root = parseXml(xml);
  const rels = relsOf(files, path);
  const colAxis = new Axis(DEFAULT_COL_WIDTH, sheet.colWidths, [], MAX_COLS);
  const rowAxis = new Axis(DEFAULT_ROW_HEIGHT, sheet.rowHeights, [], MAX_ROWS);
  const point = (el) => {
    const n = (name) => Number(child(el, name)?.text ?? 0);
    return { x: colAxis.pos(n('col')) + n('colOff') / EMU, y: rowAxis.pos(n('row')) + n('rowOff') / EMU };
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
    const paras = descendants(child(el, 'txBody'), 'p');
    const text = paras.map((p) => descendants(p, 't').map((t) => t.text).join('')).join('\n');
    const rPr = descendants(child(el, 'txBody'), 'rPr')[0] ?? descendants(child(el, 'txBody'), 'defRPr')[0];
    const algn = descendants(child(el, 'txBody'), 'pPr')[0]?.attrs.algn;
    const shape = {
      id: uid('sh'), kind: isText ? 'textbox' : PRST_KIND[prst] ?? 'rect', ...round(box), z: ++z,
      fill, stroke, text,
    };
    const xf = descendants(spPr, 'xfrm')[0];
    if ((prst === 'leftArrow') !== (xf?.attrs.flipH === '1')) shape.flip = true;
    if (xf?.attrs.flipV === '1') shape.flipV = true;
    const lw = Number(ln?.attrs.w);
    if (lw && stroke) shape.strokeWidth = Math.round((lw / EMU) * 4) / 4;
    if (rPr?.attrs.sz) shape.size = Number(rPr.attrs.sz) / 100;
    if (rPr?.attrs.b === '1') shape.bold = true;
    const tc = rPr && dmlColor(child(rPr, 'solidFill'), ctx.theme);
    if (tc) shape.color = tc;
    else if (!isText && fill) shape.color = '#ffffff';
    if (algn) shape.align = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : 'left';
    else if (!isText) shape.align = 'center';
    out.shapes.push(shape);
  };

  const readPic = (el, box) => {
    const blip = descendants(el, 'blip')[0];
    const rel = blip && rels[rid(blip) ?? blip.attrs['r:embed']] ;
    const embed = blip && Object.keys(blip.attrs).find((k) => k.endsWith(':embed') || k === 'embed');
    const target = rels[blip?.attrs[embed]]?.target ?? rel?.target;
    const bytes = target && files[target];
    if (!bytes) return;
    const ext = target.split('.').pop().toLowerCase();
    const mime = MIME[ext];
    if (!mime) { ctx.warnings.add(`지원하지 않는 그림 형식(${ext})은 가져오지 않았습니다.`); return; }
    const name = descendants(child(el, 'nvPicPr'), 'cNvPr')[0]?.attrs.name ?? '그림';
    out.images.push({ id: uid('im'), name, ...round(box), z: ++z, src: `data:${mime};base64,${toBase64(bytes)}` });
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
        const chart = target && readChart(files, target);
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
    if (sl?.attrs.name) { out._slicerBoxes[sl.attrs.name] = { ...round(box), z: ++z }; continue; }
    const content = anchor.children.find((k) => ['sp', 'cxnSp', 'pic', 'grpSp', 'graphicFrame'].includes(k.name));
    if (content) walk(content, box, null);
  }
  return out;
}

function readChart(files, path) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const plot = descendants(root, 'plotArea')[0];
  if (!plot) return null;
  const kinds = { barChart: 'column', bar3DChart: 'column', lineChart: 'line', line3DChart: 'line', pieChart: 'pie', pie3DChart: 'pie', doughnutChart: 'doughnut', areaChart: 'area', area3DChart: 'area', scatterChart: 'scatter' };
  const group = plot.children.find((c) => kinds[c.name]);
  if (!group) return null;
  let type = kinds[group.name];
  if (type === 'column' && child(group, 'barDir')?.attrs.val === 'bar') type = 'bar';
  const refs = [];
  let sheetName = null;
  for (const ser of kids(group, 'ser')) {
    for (const f of descendants(ser, 'f')) {
      const m = /^(?:'?(.*?)'?!)?(\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?)$/i.exec(f.text.trim());
      if (!m) continue;
      if (m[1]) sheetName = m[1].replace(/''/g, "'");
      const rg = refToRange(m[2]);
      if (rg) refs.push(rg);
    }
  }
  if (!refs.length) return null;
  const range = {
    r1: Math.min(...refs.map((r) => r.r1)), c1: Math.min(...refs.map((r) => r.c1)),
    r2: Math.max(...refs.map((r) => r.r2)), c2: Math.max(...refs.map((r) => r.c2)),
  };
  const titleEl = child(descendants(root, 'chart')[0] ?? root, 'title');
  const title = titleEl ? descendants(titleEl, 't').map((t) => t.text).join('') : '';
  const out = { type, title, range };
  if (sheetName) out.sheet = sheetName;
  return out;
}

// ───────────────────────── 피벗 테이블 · 슬라이서 (읽기) ─────────────────────────
/** pivotCacheDefinition → { source: { ref, sheet, name }, fields: [{ name, items: [값] }] } */
function readPivotCache(files, path) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const ws = descendants(child(root, 'cacheSource'), 'worksheetSource')[0];
  const fields = kids(child(root, 'cacheFields'), 'cacheField').map((cf) => {
    const items = (child(cf, 'sharedItems')?.children ?? []).map((it) => {
      if (it.name === 'n') return Number(it.attrs.v);
      if (it.name === 'b') return it.attrs.v === '1' || it.attrs.v === 'true';
      if (it.name === 'm') return null;
      return it.attrs.v ?? '';
    });
    return { name: cf.attrs.name ?? '', items };
  });
  return { source: { ref: ws?.attrs.ref ?? null, sheet: ws?.attrs.sheet ?? null, name: ws?.attrs.name ?? null }, fields };
}

const AGG_FROM_XLSX = { sum: 'sum', count: 'count', countNums: 'countNums', average: 'average', max: 'max', min: 'min', product: 'product', stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp' };
const SHOW_FROM_XLSX = { percentOfTotal: 'percentOfTotal', percentOfRow: 'percentOfRow', percentOfCol: 'percentOfCol' };

/**
 * pivotTableDefinition + 캐시 → Tabula 피벗 정의 (행·열·값·보고서 필터 여러 개, 레이아웃, 부분합, 총합계, 값 표시 형식)
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
    if (df.attrs.name && df.attrs.name !== valueName(v)) v.name = df.attrs.name;
    if (SHOW_FROM_XLSX[df.attrs.showDataAs]) v.showAs = SHOW_FROM_XLSX[df.attrs.showDataAs];
    return v;
  });
  const pfs = kids(child(root, 'pivotFields'), 'pivotField');
  const first = pfs[rowF[0] ?? colF[0]];
  const tblCompact = root.attrs.compact !== '0';
  const tblOutline = root.attrs.outline !== '0';
  const fCompact = first ? first.attrs.compact !== '0' && tblCompact : tblCompact;
  const fOutline = first ? first.attrs.outline !== '0' && tblOutline : tblOutline;
  const def = {
    rows: rowF.map((f) => names[f]), cols: colF.map((f) => names[f]), values, pages: pageEls.map((p) => names[Number(p.attrs.fld)]),
    layout: !fOutline ? 'tabular' : !fCompact ? 'outline' : 'compact',
  };
  if ([...rowF, ...colF].some((f) => pfs[f]?.attrs.defaultSubtotal === '0')) def.subtotals = false;
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
  const loc = refToRange(child(root, 'location')?.attrs.ref ?? '');
  if (loc) {
    const pageRows = def.pages?.length ? def.pages.length + 1 : 0;
    const top = Math.max(0, loc.r1 - pageRows);
    Object.assign(def, { top, left: loc.c1, area: { ...loc, r1: top } });
  }
  return def;
}

/** slicerCacheDefinition → { name, sourceName, table: { tableId, column } | null, pivot: { tabId, name } | null } */
function readSlicerCache(files, path) {
  const xml = textOf(files[path]);
  if (!xml) return null;
  const root = parseXml(xml);
  const tsc = descendants(root, 'tableSlicerCache')[0];
  const pt = descendants(child(root, 'pivotTables'), 'pivotTable')[0];
  return {
    name: root.attrs.name, sourceName: root.attrs.sourceName,
    table: tsc ? { tableId: Number(tsc.attrs.tableId), column: Number(tsc.attrs.column) } : null,
    pivot: pt ? { tabId: Number(pt.attrs.tabId), name: pt.attrs.name } : null,
  };
}

/** 피벗 테이블 정의와 슬라이서를 시트에 연결하고 임시 속성은 지움 */
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
      } else if (c?.pivot) {
        const ps = sheets.find((x) => x._sheetId === c.pivot.tabId && x.pivot) ?? sheets.find((x) => x._pivotName === c.pivot.name && x.pivot);
        if (ps) source = ps === s ? { kind: 'pivot', self: true, field: c.sourceName } : { kind: 'pivot', sheet: ps.name, field: c.sourceName };
      }
      if (!source) { ctx.warnings.add('연결 대상을 찾지 못한 슬라이서는 가져오지 않았습니다.'); continue; }
      const box = boxes[sl.name] ?? { x: 20 + i * 190, y: 20, w: 180, h: 200 };
      s.slicers.push({
        id: `sl${Math.random().toString(36).slice(2, 9)}`, caption: sl.caption ?? sl.name, source, columns: Math.max(1, sl.columns || 1),
        color: STYLE_SLICER[String(sl.style ?? '').toLowerCase()] ?? 'blue', multi: false, ...box,
      });
      i++;
    }
  });
  for (const s of sheets) {
    delete s._slicers; delete s._pivots; delete s._slicerBoxes; delete s._sheetId; delete s._pivotName;
    for (const t of s.tables) { delete t._xmlId; delete t._colNames; }
  }
}

/** xlsx 바이트 → 통합 문서 데이터 ({ sheets: [...] }, 경고 목록) */
export function readXlsx(bytes) {
  const files = unzip(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  const wbPath = Object.keys(files).find((f) => /^xl\/workbook\.xml$/i.test(f))
    ?? relsTarget(files, '', 'officeDocument');
  if (!wbPath || !files[wbPath]) throw new Error('엑셀 통합 문서(.xlsx)가 아닙니다');
  const wbRoot = parseXml(textOf(files[wbPath]));
  const wbRels = relsOf(files, wbPath);
  const theme = readTheme(files, wbRels);
  const { xfs, dxfs } = readStyles(files, wbRels, theme);
  const ssRel = Object.values(wbRels).find((r) => r.type === 'sharedStrings');
  const strings = ssRel && files[ssRel.target] ? kids(parseXml(textOf(files[ssRel.target])), 'si').map(allText) : [];
  // 이름 정의 (시트 범위 이름은 localSheetId → 시트 이름)
  const allSheetNames = kids(child(wbRoot, 'sheets'), 'sheet').map((sh) => sh.attrs.name.slice(0, 31));
  const names = [];
  for (const dn of kids(child(wbRoot, 'definedNames'), 'definedName')) {
    const name = dn.attrs.name;
    const text = (dn.text ?? '').trim();
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
  const ctx = { xfs, dxfs, strings, theme, warnings: new Set(), isName, nameMulti };
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
    const sheet = readSheet(files, rel.target, ctx);
    if (sheet.codeName) sheetCodes[sh.attrs.name.slice(0, 31)] = sheet.codeName;
    delete sheet.codeName;
    unsupported += sheet.unsupported;
    delete sheet.unsupported;
    sheets.push({ name: sh.attrs.name.slice(0, 31), ...sheet, _sheetId: Number(sh.attrs.sheetId) });
  }
  linkPivotsAndSlicers(files, wbRels, sheets, ctx);
  if (unsupported) warnings.push(`지원하지 않는 함수가 쓰인 수식 ${unsupported}개는 수식을 유지하고 파일에 저장된 계산 결과를 표시합니다.`);
  warnings.push(...ctx.warnings);
  if (!sheets.length) throw new Error('가져올 시트가 없습니다');
  const data = { sheets };
  if (names.length) data.names = names;
  // 매크로(.xlsm): vbaProject.bin 을 그대로 보존 (실행하지 않음)
  const vbaRel = Object.values(wbRels).find((r) => r.type === 'vbaProject');
  if (vbaRel && files[vbaRel.target]) {
    data.vba = {
      bin: toBase64(files[vbaRel.target]),
      codeName: child(wbRoot, 'workbookPr')?.attrs.codeName ?? null,
      sheetCodes,
    };
  }
  const active = Number(descendants(child(wbRoot, 'bookViews'), 'workbookView')[0]?.attrs.activeTab ?? 0);
  return { data, active: Math.min(active, sheets.length - 1), warnings };
}

function relsTarget(files, path, type) {
  const rels = relsOf(files, path);
  return Object.values(rels).find((r) => r.type === type)?.target;
}

// ───────────────────────── 쓰기 ─────────────────────────
class StylePool {
  constructor() {
    this.fonts = [`<font><sz val="11"/><color theme="1"/><name val="${DEFAULT_FONT}"/><family val="3"/><charset val="129"/></font>`];
    this.fills = ['<fill><patternFill patternType="none"/></fill>', '<fill><patternFill patternType="gray125"/></fill>'];
    this.borders = ['<border><left/><right/><top/><bottom/><diagonal/></border>'];
    this.numFmts = [];
    this.xfs = ['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'];
    this.dxfs = [];
    this.maps = { font: new Map([[this.fonts[0], 0]]), fill: new Map(this.fills.map((f, i) => [f, i])), border: new Map([[this.borders[0], 0]]), fmt: new Map(), xf: new Map([['{}', 0]]) };
  }

  intern(kind, list, xml) {
    const map = this.maps[kind];
    if (!map.has(xml)) { map.set(xml, list.length); list.push(xml); }
    return map.get(xml);
  }

  fmtId(style) {
    const code = fmtCode(style);
    if (code === null) return 0;
    if (!this.maps.fmt.has(code)) {
      const id = 164 + this.numFmts.length;
      this.maps.fmt.set(code, id);
      this.numFmts.push(`<numFmt numFmtId="${id}" formatCode="${esc(code)}"/>`);
    }
    return this.maps.fmt.get(code);
  }

  xf(style) {
    if (!style || !Object.keys(style).length) return 0;
    const k = JSON.stringify(Object.keys(style).sort().map((key) => [key, style[key]]));
    if (this.maps.xf.has(k)) return this.maps.xf.get(k);
    const font = `<font>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.strike ? '<strike/>' : ''}${style.underline ? '<u/>' : ''}<sz val="${style.size || 11}"/>${style.color ? `<color rgb="${argb(style.color)}"/>` : '<color theme="1"/>'}<name val="${esc(style.font || DEFAULT_FONT)}"/><family val="3"/><charset val="129"/></font>`;
    const fontId = this.intern('font', this.fonts, font);
    const fillId = style.fill ? this.intern('fill', this.fills, `<fill><patternFill patternType="solid"><fgColor rgb="${argb(style.fill)}"/><bgColor indexed="64"/></patternFill></fill>`) : 0;
    const side = (n, on) => (on ? `<${n} style="thin"><color indexed="64"/></${n}>` : `<${n}/>`);
    const borderId = style.bt || style.bb || style.bl || style.br
      ? this.intern('border', this.borders, `<border>${side('left', style.bl)}${side('right', style.br)}${side('top', style.bt)}${side('bottom', style.bb)}<diagonal/></border>`)
      : 0;
    const numFmtId = this.fmtId(style);
    const align = [];
    if (style.align) align.push(`horizontal="${style.align}"`);
    if (style.valign) align.push(`vertical="${style.valign === 'middle' ? 'center' : 'top'}"`);
    if (style.wrap) align.push('wrapText="1"');
    if (style.indent) align.push(`indent="${style.indent}"`);
    const xml = `<xf numFmtId="${numFmtId}" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ' applyNumberFormat="1"' : ''}${fontId ? ' applyFont="1"' : ''}${fillId ? ' applyFill="1"' : ''}${borderId ? ' applyBorder="1"' : ''}${align.length ? ` applyAlignment="1"><alignment ${align.join(' ')}/></xf>` : '/>'}`;
    const id = this.xfs.length;
    this.xfs.push(xml);
    this.maps.xf.set(k, id);
    return id;
  }

  dxf(style) {
    const font = style.color || style.bold || style.italic || style.underline || style.strike
      ? `<font>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.strike ? '<strike/>' : ''}${style.underline ? '<u/>' : ''}${style.color ? `<color rgb="${argb(style.color)}"/>` : ''}</font>` : '';
    const fill = style.fill ? `<fill><patternFill><bgColor rgb="${argb(style.fill)}"/></patternFill></fill>` : '';
    const side = (n, on) => (on ? `<${n} style="thin"><color auto="1"/></${n}>` : '');
    const border = style.bt || style.bb || style.bl || style.br ? `<border>${side('left', style.bl)}${side('right', style.br)}${side('top', style.bt)}${side('bottom', style.bb)}</border>` : '';
    // 표시 형식도 조건부 서식으로 바꿀 수 있음 (dxf 안의 numFmt)
    const code = style.numFmt ? fmtCode(style) : null;
    const nf = code !== null ? `<numFmt numFmtId="${this.fmtId(style)}" formatCode="${esc(code)}"/>` : '';
    this.dxfs.push(`<dxf>${font}${nf}${fill}${border}</dxf>`);
    return this.dxfs.length - 1;
  }

  xml() {
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="${NS_MAIN}">`
      + (this.numFmts.length ? `<numFmts count="${this.numFmts.length}">${this.numFmts.join('')}</numFmts>` : '')
      + `<fonts count="${this.fonts.length}">${this.fonts.join('')}</fonts>`
      + `<fills count="${this.fills.length}">${this.fills.join('')}</fills>`
      + `<borders count="${this.borders.length}">${this.borders.join('')}</borders>`
      + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      + `<cellXfs count="${this.xfs.length}">${this.xfs.join('')}</cellXfs>`
      + '<cellStyles count="1"><cellStyle name="표준" xfId="0" builtinId="0"/></cellStyles>'
      + `<dxfs count="${this.dxfs.length}">${this.dxfs.join('')}</dxfs>`
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
  return `<conditionalFormatting sqref="${ref}">${body}</conditionalFormatting>`;
}

function chartXml(wb, si, chart) {
  const srcIndex = chart.sheet ? wb.sheetIndexByName(chart.sheet) : si;
  const s = srcIndex >= 0 ? srcIndex : si;
  const sheetName = quoteSheetName(wb.sheets[s].name);
  const rg = chart.range;
  const rows = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [];
    for (let c = rg.c1; c <= rg.c2; c++) row.push(wb.getValue(s, r, c));
    rows.push(row);
  }
  const L = chartLayout(rows, chart.type);
  const ref = (r1, c1, r2, c2) => `${sheetName}!${rangeRef({ r1, c1, r2, c2 }, true)}`;
  const label = (v) => (v === null ? '' : typeof v === 'number' ? formatGeneral(v) : isError(v) ? v.code : String(v));
  const strCache = (vals) => `<c:strCache><c:ptCount val="${vals.length}"/>${vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(label(v))}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (vals) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${vals.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('')}</c:numCache>`;
  const series = [];
  const dr0 = rg.r1 + L.firstDataRow;
  const dc0 = rg.c1 + L.firstDataCol;
  if (L.byCols) {
    for (let c = dc0; c <= rg.c2; c++) {
      const vals = rows.slice(L.firstDataRow).map((r) => r[c - rg.c1]);
      series.push({
        tx: L.headRow ? { ref: ref(rg.r1, c, rg.r1, c), v: rows[0][c - rg.c1] } : null,
        cat: L.catCol ? { ref: ref(dr0, rg.c1, rg.r2, rg.c1), v: rows.slice(L.firstDataRow).map((r) => r[0]) } : null,
        val: { ref: ref(dr0, c, rg.r2, c), v: vals },
      });
    }
  } else {
    for (let r = dr0; r <= rg.r2; r++) {
      series.push({
        tx: L.catCol ? { ref: ref(r, rg.c1, r, rg.c1), v: rows[r - rg.r1][0] } : null,
        cat: L.headRow ? { ref: ref(rg.r1, dc0, rg.r1, rg.c2), v: rows[0].slice(L.firstDataCol) } : null,
        val: { ref: ref(r, dc0, r, rg.c2), v: rows[r - rg.r1].slice(L.firstDataCol) },
      });
    }
  }
  const scatter = chart.type === 'scatter';
  const withNums = series.filter((sr) => sr.val.v.some((v) => typeof v === 'number'));
  if (withNums.length) series.splice(0, series.length, ...withNums);
  const serXml = series.map((sr, i) => {
    const tx = sr.tx ? `<c:tx><c:strRef><c:f>${esc(sr.tx.ref)}</c:f>${strCache([sr.tx.v])}</c:strRef></c:tx>` : '';
    const color = PALETTE[i % PALETTE.length].slice(1);
    const fill = `<a:solidFill><a:srgbClr val="${color}"/></a:solidFill>`;
    const pie = chart.type === 'pie' || chart.type === 'doughnut';
    let spPr;
    if (scatter) spPr = '<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>';
    else if (chart.type === 'line') spPr = `<c:spPr><a:ln w="28575" cap="rnd">${fill}<a:round/></a:ln></c:spPr>`;
    else if (pie) spPr = '<c:spPr><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr>';
    else spPr = `<c:spPr>${fill}</c:spPr>`;
    const markerFill = `<c:spPr>${fill}<a:ln w="9525">${fill}</a:ln></c:spPr>`;
    const marker = chart.type === 'line' || scatter ? `<c:marker><c:symbol val="circle"/><c:size val="5"/>${markerFill}</c:marker>` : '';
    const dPt = pie ? sr.val.v.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${PALETTE[k % PALETTE.length].slice(1)}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '';
    const invert = chart.type === 'column' || chart.type === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
    const catNumeric = sr.cat && sr.cat.v.every((v) => typeof v === 'number');
    const cat = sr.cat
      ? (catNumeric
        ? `<c:${scatter ? 'xVal' : 'cat'}><c:numRef><c:f>${esc(sr.cat.ref)}</c:f>${numCache(sr.cat.v)}</c:numRef></c:${scatter ? 'xVal' : 'cat'}>`
        : `<c:${scatter ? 'xVal' : 'cat'}><c:strRef><c:f>${esc(sr.cat.ref)}</c:f>${strCache(sr.cat.v)}</c:strRef></c:${scatter ? 'xVal' : 'cat'}>`)
      : '';
    const val = `<c:${scatter ? 'yVal' : 'val'}><c:numRef><c:f>${esc(sr.val.ref)}</c:f>${numCache(sr.val.v)}</c:numRef></c:${scatter ? 'yVal' : 'val'}>`;
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${invert}${marker}${dPt}${cat}${val}${chart.type === 'line' || scatter ? '<c:smooth val="0"/>' : ''}</c:ser>`;
  }).join('');
  const axes = '<c:axId val="111111111"/><c:axId val="222222222"/>';
  let group;
  switch (chart.type) {
    case 'bar': group = `<c:barChart><c:barDir val="bar"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXml}<c:gapWidth val="182"/>${axes}</c:barChart>`; break;
    case 'line': group = `<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXml}<c:marker val="1"/>${axes}</c:lineChart>`; break;
    case 'area': group = `<c:areaChart><c:grouping val="standard"/><c:varyColors val="0"/>${serXml}${axes}</c:areaChart>`; break;
    case 'pie': group = `<c:pieChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/></c:pieChart>`; break;
    case 'doughnut': group = `<c:doughnutChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/><c:holeSize val="50"/></c:doughnutChart>`; break;
    case 'scatter': group = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${serXml}${axes}</c:scatterChart>`; break;
    default: group = `<c:barChart><c:barDir val="col"/><c:grouping val="clustered"/><c:varyColors val="0"/>${serXml}<c:gapWidth val="219"/>${axes}</c:barChart>`;
  }
  const pie = chart.type === 'pie' || chart.type === 'doughnut';
  const catPos = chart.type === 'bar' ? 'l' : 'b';
  const valPos = chart.type === 'bar' ? 'b' : 'l';
  const catAx = scatter
    ? `<c:valAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx><c:axId val="111111111"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${catPos}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="222222222"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAx = `<c:valAx><c:axId val="222222222"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/><c:majorGridlines/><c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="111111111"/><c:crosses val="autoZero"/><c:crossBetween val="${chart.type === 'area' ? 'midCat' : 'between'}"/></c:valAx>`;
  const title = chart.title
    ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="ko-KR" sz="1400" b="0"/><a:t>${esc(chart.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
    : '<c:autoTitleDeleted val="1"/>';
  const legend = series.length > 1 || pie ? '<c:legend><c:legendPos val="b"/><c:overlay val="0"/></c:legend>' : '';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${group}${pie ? '' : catAx + valAx}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart></c:chartSpace>`;
}

const KIND_PRST = { rect: 'rect', roundRect: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', arrow: 'rightArrow', textbox: 'rect', line: 'straightConnector1' };
const hex6 = (c) => (c ?? '#000000').replace('#', '').toUpperCase().padStart(6, '0').slice(0, 6);

/** 도형 → <xdr:sp> / <xdr:cxnSp> */
function shapeXml(sh, id, xfrm) {
  const name = esc(sh.name || `${sh.kind === 'textbox' ? 'TextBox' : '도형'} ${id - 1}`);
  const fill = sh.fill ? `<a:solidFill><a:srgbClr val="${hex6(sh.fill)}"/></a:solidFill>` : '<a:noFill/>';
  const ln = sh.stroke ? `<a:ln w="${Math.round((sh.strokeWidth ?? 1) * EMU)}"><a:solidFill><a:srgbClr val="${hex6(sh.stroke)}"/></a:solidFill>${sh.kind === 'line' && sh.arrow ? '<a:tailEnd type="triangle"/>' : ''}</a:ln>` : '<a:ln><a:noFill/></a:ln>';
  if (sh.kind === 'line') {
    return `<xdr:cxnSp macro=""><xdr:nvCxnSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvCxnSpPr/></xdr:nvCxnSpPr><xdr:spPr>${xfrm(sh)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${ln}</xdr:spPr></xdr:cxnSp>`;
  }
  const algn = sh.align === 'center' ? 'ctr' : sh.align === 'right' ? 'r' : 'l';
  const rPr = `<a:rPr lang="ko-KR" sz="${Math.round((sh.size ?? 11) * 100)}"${sh.bold ? ' b="1"' : ''}><a:solidFill><a:srgbClr val="${hex6(sh.color ?? '#000000')}"/></a:solidFill></a:rPr>`;
  const paras = String(sh.text ?? '').split('\n').map((line) => `<a:p><a:pPr algn="${algn}"/>${line ? `<a:r>${rPr}<a:t>${esc(line)}</a:t></a:r>` : `<a:endParaRPr lang="ko-KR" sz="${Math.round((sh.size ?? 11) * 100)}"/>`}</a:p>`).join('');
  const anchor = sh.kind === 'textbox' ? 't' : 'ctr';
  return `<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="${id}" name="${name}"/><xdr:cNvSpPr${sh.kind === 'textbox' ? ' txBox="1"' : ''}/></xdr:nvSpPr><xdr:spPr>${xfrm(sh)}<a:prstGeom prst="${KIND_PRST[sh.kind] ?? 'rect'}"><a:avLst/></a:prstGeom>${fill}${ln}</xdr:spPr><xdr:txBody><a:bodyPr vertOverflow="clip" horzOverflow="clip" wrap="square" rtlCol="0" anchor="${anchor}"/><a:lstStyle/>${paras}</xdr:txBody></xdr:sp>`;
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
const REL_MS = 'http://schemas.microsoft.com/office/2007/relationships';
const SLICER_STYLE = { blue: 'SlicerStyleLight1', orange: 'SlicerStyleLight2', gray: 'SlicerStyleLight3', gold: 'SlicerStyleLight4', sky: 'SlicerStyleLight5', green: 'SlicerStyleLight6' };
const STYLE_SLICER = Object.fromEntries(Object.entries(SLICER_STYLE).map(([k, v]) => [v.toLowerCase(), k]));
const PIVOT_SUBTOTAL = { count: 'count', average: 'average', max: 'max', min: 'min', product: 'product', countNums: 'countNums', stdDev: 'stdDev', stdDevp: 'stdDevp', var: 'var', varp: 'varp' };

/** 필드의 고유 값 (피벗 결과와 같은 순서) → { keys, index: Map(key → 순번) } */
function fieldItems(data, f) {
  const seen = new Map();
  for (const r of data) { const k = keyOf(r[f]); if (!seen.has(`${typeof k}:${k}`)) seen.set(`${typeof k}:${k}`, k); }
  const keys = sortKeys([...seen.values()]);
  const index = new Map(keys.map((k, i) => [`${typeof k}:${k}`, i]));
  return { keys, index };
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
    attrs.push(`minValue="${Math.min(...nums)}"`, `maxValue="${Math.max(...nums)}"`);
  }
  if (hasBlank) attrs.push('containsBlank="1"');
  if (hasStr && hasNum) attrs.push('containsMixedTypes="1"');
  if (!keys) return `<sharedItems${attrs.length ? ` ${attrs.join(' ')}` : ''}/>`;
  const items = keys.map((k) => (k === EMPTY ? '<m/>' : typeof k === 'number' ? `<n v="${k}"/>` : `<s v="${esc(String(k))}"/>`)).join('');
  return `<sharedItems${attrs.length ? ` ${attrs.join(' ')}` : ''} count="${keys.length}">${items}</sharedItems>`;
}

/**
 * 시트의 피벗 정의 → 엑셀 피벗 캐시·피벗 테이블 XML
 * extraFields: 슬라이서가 쓰는 필드 이름(소문자)
 */
function pivotParts(wb, si, def, cacheId, name, extraFields) {
  const src = pivotSourceData(wb, def);
  if (!src || src.rows.length < 2) return null;
  const header = headerNames(src.rows);
  const d = normalizeDef(def, header);
  if (!d.rows.length && !d.cols.length && !d.values.length) return null;
  const fx = (n) => header.indexOf(n);
  const rowF = d.rows.map(fx);
  const colF = d.cols.map(fx);
  const pageF = d.pages.map(fx);
  const data = src.rows.slice(1).filter((r) => !r.every((v) => v === null || v === ''));
  const filters = Object.entries(d.filters).map(([n, allowed]) => [header.findIndex((h) => h.toLowerCase() === n.toLowerCase()), new Set(allowed)]).filter(([i]) => i >= 0);
  const listed = new Set([...rowF, ...colF, ...pageF, ...filters.map(([i]) => i)]);
  header.forEach((h, i) => { if (extraFields?.has(h.toLowerCase())) listed.add(i); });
  const items = new Map();
  for (const f of listed) items.set(f, fieldItems(data, f));
  const cacheFields = header.map((h, f) => `<cacheField name="${esc(h)}" numFmtId="0">${sharedItemsXml(data.map((r) => r[f]), items.get(f)?.keys ?? null)}</cacheField>`).join('');
  const sourceXml = src.table
    ? `<worksheetSource name="${esc(src.table)}"/>`
    : `<worksheetSource ref="${rangeRef(src.ref)}" sheet="${esc(wb.sheets[src.si].name)}"/>`;
  const cacheXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotCacheDefinition xmlns="${NS_MAIN}" xmlns:r="${NS_R}" saveData="0" refreshOnLoad="1" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" recordCount="${data.length}">`
    + `<cacheSource type="worksheet">${sourceXml}</cacheSource><cacheFields count="${header.length}">${cacheFields}</cacheFields>`
    + `<extLst><ext uri="{725AE2AE-9491-48be-B2B4-4EB974FC3084}" xmlns:x14="${NS_X14}"><x14:pivotCacheDefinition pivotCacheId="${cacheId}"/></ext></extLst></pivotCacheDefinition>`;

  const hiddenKey = (f, k) => {
    const set = filters.find(([i]) => i === f)?.[1];
    return set ? !set.has(itemText(k)) : false;
  };
  // 보이는 항목 기준으로 결과 배치 계산 (앱 화면과 같은 배치)
  const { rows: visible } = resolvePivot(src.rows, def);
  const { meta } = computePivot(visible, d);
  const values = meta.values;
  const V = values.length;
  const multiV = V > 1;
  const xOf = (f, key) => items.get(f)?.index.get(`${typeof key}:${key}`) ?? 0;
  const xTag = (v) => (v ? `<x v="${v}"/>` : '<x/>');

  // 행 항목
  const rowXml = [];
  if (!rowF.length) rowXml.push('<i/>');
  else {
    let prev = [];
    for (const it of meta.rowItems) {
      if (it.kind === 'grand') { rowXml.push('<i t="grand"><x/></i>'); continue; }
      if (it.kind === 'sub') {
        const n = it.node;
        rowXml.push(`<i t="default"${n.depth ? ` r="${n.depth}"` : ''}>${xTag(xOf(rowF[n.depth], n.key))}</i>`);
        continue;
      }
      const chain = it.chain ?? [it.node];
      const xs = (it.chain ? chain : [it.node]).map((n) => xOf(rowF[n.depth], n.key));
      const start = it.chain ? 0 : it.node.depth;
      if (it.chain) {
        let r = 0;
        while (r < xs.length && r < prev.length && prev[r] === xs[r]) r++;
        if (r >= xs.length) r = xs.length - 1;
        rowXml.push(`<i${r ? ` r="${r}"` : ''}>${xs.slice(r).map(xTag).join('')}</i>`);
        prev = xs;
      } else rowXml.push(`<i${start ? ` r="${start}"` : ''}>${xTag(xs[0])}</i>`);
    }
  }
  // 열 항목
  const colXml = [];
  const colFieldsAll = [...colF, ...(multiV ? [-2] : [])];
  if (!colFieldsAll.length) colXml.push('<i/>');
  else {
    let prev = [];
    for (const leaf of meta.colLeaves) {
      const ia = leaf.vi ? ` i="${leaf.vi}"` : '';
      if (leaf.kind === 'grand') { colXml.push(`<i t="grand"${ia}><x/></i>`); continue; }
      const chain = [];
      for (let n = leaf.node; n && n.depth >= 0; n = n.parent) chain.unshift(n);
      if (leaf.kind === 'sub') {
        const n = leaf.node;
        colXml.push(`<i t="default"${n.depth ? ` r="${n.depth}"` : ''}${ia}>${xTag(xOf(colF[n.depth], n.key))}</i>`);
        prev = [];
        continue;
      }
      const xs = [...chain.map((n) => xOf(colF[n.depth], n.key)), ...(multiV ? [leaf.vi] : [])];
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
    const attrs = [];
    if (rowF.includes(f)) attrs.push('axis="axisRow"');
    else if (colF.includes(f)) attrs.push('axis="axisCol"');
    else if (pageF.includes(f)) attrs.push('axis="axisPage"');
    if (valueFieldIdx.has(f)) attrs.push('dataField="1"');
    if (tabular || outline) attrs.push('compact="0"');
    if (tabular) attrs.push('outline="0"');
    attrs.push('showAll="0"');
    const onAxis = rowF.includes(f) || colF.includes(f);
    if (onAxis && !d.subtotals) attrs.push('defaultSubtotal="0"');
    const it = items.get(f);
    if (!it) return `<pivotField ${attrs.join(' ')}/>`;
    const list = it.keys.map((k, i) => `<item${hiddenKey(f, k) ? ' h="1"' : ''} x="${i}"/>`).join('');
    const def0 = onAxis && !d.subtotals ? '' : '<item t="default"/>';
    return `<pivotField ${attrs.join(' ')}><items count="${it.keys.length + (def0 ? 1 : 0)}">${list}${def0}</items></pivotField>`;
  }).join('');

  const top = (def.top ?? 0) + meta.pageRows;
  const left = def.left ?? 0;
  const loc = { r1: top, c1: left, r2: top + meta.bodyRows - 1, c2: left + meta.width - 1 };
  const firstHeaderRow = colF.length ? 1 : multiV ? 0 : 1;
  const pageXml = pageF.length ? `<pageFields count="${pageF.length}">${pageF.map((f) => {
    const allowed = filters.find(([i]) => i === f)?.[1];
    const one = allowed && allowed.size === 1 ? items.get(f)?.keys.findIndex((k) => itemText(k) === [...allowed][0]) : -1;
    return `<pageField fld="${f}"${one >= 0 ? ` item="${one}"` : ''} hier="-1"/>`;
  }).join('')}</pageFields>` : '';
  const SHOW = { percentOfTotal: 'percentOfTotal', percentOfRow: 'percentOfRow', percentOfCol: 'percentOfCol' };
  const dataXml = values.map((v) => {
    const f = Math.max(0, fx(v.field));
    const sub = PIVOT_SUBTOTAL[v.agg] ? ` subtotal="${PIVOT_SUBTOTAL[v.agg]}"` : '';
    const show = SHOW[v.showAs] ? ` showDataAs="${SHOW[v.showAs]}"` : '';
    const numFmt = SHOW[v.showAs] ? 10 : ['average', 'stdDev', 'stdDevp', 'var', 'varp'].includes(v.agg) ? 4 : 3;
    return `<dataField name="${esc(valueName(v))}" fld="${f}"${sub}${show} baseField="0" baseItem="0" numFmtId="${numFmt}"/>`;
  }).join('');
  const tableAttrs = [
    `name="${esc(name)}"`, `cacheId="${cacheId}"`, 'applyNumberFormats="0"', 'applyBorderFormats="0"', 'applyFontFormats="0"', 'applyPatternFormats="0"',
    'applyAlignmentFormats="0"', 'applyWidthHeightFormats="1"', 'dataCaption="값"', 'updatedVersion="6"', 'minRefreshableVersion="3"', 'useAutoFormatting="1"',
    ...(d.grandRows ? [] : ['rowGrandTotals="0"']), ...(d.grandCols ? [] : ['colGrandTotals="0"']),
    'itemPrintTitles="1"', 'createdVersion="6"', 'indent="0"', ...(tabular || outline ? ['compact="0"', 'compactData="0"'] : []),
    `outline="${tabular ? 0 : 1}"`, `outlineData="${tabular ? 0 : 1}"`, ...(tabular ? ['gridDropZones="1"'] : []), 'multipleFieldFilters="0"',
  ];
  const tableXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<pivotTableDefinition xmlns="${NS_MAIN}" ${tableAttrs.join(' ')}>`
    + `<location ref="${rangeRef(loc)}" firstHeaderRow="${firstHeaderRow}" firstDataRow="${meta.headerRows}" firstDataCol="${meta.labelCols}"${pageF.length ? ` rowPageCount="${pageF.length}" colPageCount="1"` : ''}/>`
    + `<pivotFields count="${header.length}">${pivotFields}</pivotFields>`
    + (rowF.length ? `<rowFields count="${rowF.length}">${rowF.map((f) => `<field x="${f}"/>`).join('')}</rowFields>` : '')
    + `<rowItems count="${rowXml.length}">${rowXml.join('')}</rowItems>`
    + (colFieldsAll.length ? `<colFields count="${colFieldsAll.length}">${colFieldsAll.map((f) => `<field x="${f}"/>`).join('')}</colFields>` : '')
    + `<colItems count="${colXml.length}">${colXml.join('')}</colItems>`
    + pageXml
    + `<dataFields count="${values.length}">${dataXml}</dataFields>`
    + '<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0" showLastColumn="1"/></pivotTableDefinition>';

  // 슬라이서 캐시용: 필드 이름 → 항목 선택 상태
  const slicerItems = (fieldName) => {
    const f = header.findIndex((h) => h.toLowerCase() === String(fieldName).toLowerCase());
    if (f < 0 || !items.has(f)) return null;
    const others = filters.filter(([i]) => i !== f);
    const rowsOk = data.filter((r) => others.every(([i, set]) => set.has(itemText(r[i]))));
    const present = new Set(rowsOk.map((r) => `${typeof keyOf(r[f])}:${keyOf(r[f])}`));
    return {
      field: header[f],
      xml: items.get(f).keys.map((k, i) => `<i x="${i}"${hiddenKey(f, k) ? '' : ' s="1"'}${present.has(`${typeof k}:${k}`) ? '' : ' nd="1"'}/>`).join(''),
      count: items.get(f).keys.length,
    };
  };
  return { cacheXml, tableXml, slicerItems };
}

/** 슬라이서가 가리키는 피벗 시트 번호 */
function slicerPivotSheet(wb, si, sl) {
  if (sl.source?.kind !== 'pivot') return -1;
  return sl.source.self || !sl.source.sheet ? si : wb.sheetIndexByName(sl.source.sheet);
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
  return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${anchorAt(sl.x, sl.y)}</xdr:from><xdr:to>${anchorAt(sl.x + sl.w, sl.y + sl.h)}</xdr:to><mc:AlternateContent xmlns:mc="${NS_MC}">${choice}${frame}</mc:Choice>${fallback}</mc:AlternateContent><xdr:clientData/></xdr:twoCellAnchor>`;
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
export function writeXlsx(wb, { activeSheet = 0 } = {}) {
  const files = {};
  const pool = new StylePool();
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
  const definedNames = [];
  const vba = wb.vba?.bin ? wb.vba : null;
  const nameSet = new Set((wb.names ?? []).map((n) => n.name.toUpperCase()));
  fileNameCheck = (n) => nameSet.has(String(n).toUpperCase());
  let dynamicCells = 0;
  // 표 번호(표 슬라이서가 참조)와 피벗(슬라이서 필드) 미리 계산
  const tableIds = new Map();
  let tableSeq = 0;
  wb.sheets.forEach((s) => (s.tables ?? []).forEach((t) => { if (t.r1 < EXCEL_MAX_ROWS) tableIds.set(t.name.toLowerCase(), ++tableSeq); }));
  const pivotSlicerFields = new Map();
  wb.sheets.forEach((s, i) => (s.slicers ?? []).forEach((sl) => {
    const pi = slicerPivotSheet(wb, i, sl);
    if (pi < 0) return;
    if (!pivotSlicerFields.has(pi)) pivotSlicerFields.set(pi, new Set());
    pivotSlicerFields.get(pi).add(String(sl.source.field).toLowerCase());
  }));
  const pivotInfo = new Map(); // 시트 번호 → { name, cacheId, parts } (시트의 첫 피벗, 슬라이서 연결용)
  const pivotList = new Map(); // 시트 번호 → 모든 피벗
  let pivotNo = 0;
  wb.sheets.forEach((s, i) => {
    for (const [j, def] of [s.pivot, ...(s.pivotsExtra ?? [])].entries()) {
      if (!def) continue;
      const name = `피벗 테이블${pivotNo + 1}`;
      const parts = pivotParts(wb, i, def, pivotNo + 1, name, j === 0 ? pivotSlicerFields.get(i) : null);
      if (!parts) continue;
      pivotNo++;
      const info = { name, cacheId: pivotNo, parts };
      if (j === 0) pivotInfo.set(i, info);
      if (!pivotList.has(i)) pivotList.set(i, []);
      pivotList.get(i).push(info);
    }
  });
  const pivotCaches = [];
  const slicerCachesPivot = [];
  const slicerCachesTable = [];
  const usedCacheNames = new Set();
  const usedSlicerNames = new Set();
  let slicerPartNo = 0;
  let slicerCacheNo = 0;

  wb.sheets.forEach((sheet, si) => {
    const sheetRels = [];
    const addRel = (type, target) => { const id = `rId${sheetRels.length + 1}`; sheetRels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`); return id; };

    // 셀을 행별로 정리
    const rows = new Map();
    let maxR = 0;
    let maxC = 0;
    for (const [k, cell] of sheet.cells) {
      const [r, c] = k.split(',').map(Number);
      if (r >= EXCEL_MAX_ROWS) continue; // 엑셀 파일에는 1,048,576행까지만 저장 가능
      if (!rows.has(r)) rows.set(r, []);
      rows.get(r).push([c, cell]);
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
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
    for (const k of [...Object.keys(sheet.rowHeights), ...Object.keys(sheet.hiddenRows), ...Object.keys(sheet.rowStyles), ...Object.keys(sheet.filter?.hidden ?? {}), ...(sheet.tables ?? []).flatMap((t) => Object.keys(t.filter?.hidden ?? {}))]) rowKeys.add(Number(k));
    const sortedRows = [...rowKeys].filter((r) => r < EXCEL_MAX_ROWS).sort((a, b) => a - b);

    const rowXml = sortedRows.map((r) => {
      const cells = (rows.get(r) ?? []).sort((a, b) => a[0] - b[0]);
      const attrs = [`r="${r + 1}"`];
      if (sheet.rowHeights[r] !== undefined) attrs.push(`ht="${px2pt(sheet.rowHeights[r])}"`, 'customHeight="1"');
      if (sheet.hiddenRows[r] || sheet.filter?.hidden?.[r] || (sheet.tables ?? []).some((t) => t.filter?.hidden?.[r])) attrs.push('hidden="1"');
      if (sheet.rowStyles[r]) attrs.push(`s="${pool.xf({ ...sheet.allStyle, ...sheet.rowStyles[r] })}"`, 'customFormat="1"');
      const cx = cells.map(([c, cell]) => {
        const ref = cellName(r, c);
        const st = wb.styleAt(si, r, c);
        const s = pool.xf(st);
        const sAttr = s ? ` s="${s}"` : '';
        const v = wb.getValue(si, r, c);
        if (!cell.raw && (v === null || !cell.spilled && !wb.spillAnchorOf(si, r, c))) return s ? `<c r="${ref}"${sAttr}/>` : '';
        if (cell.formula) {
          // 배열을 돌려줄 수 있는 수식은 동적 배열 수식으로 (cm="1" + t="array")
          const dyn = !!cell.maybeArray;
          const sp = dyn ? wb.spillRange(si, r, c) : null;
          if (dyn) dynamicCells++;
          const fAttr = dyn ? ` t="array" ref="${sp ? rangeRef({ ...sp, r2: Math.min(sp.r2, EXCEL_MAX_ROWS - 1) }) : ref}" aca="false"` : '';
          const cm = dyn ? ' cm="1"' : '';
          const f = `<f${fAttr}>${esc(exportFormula(cell.raw, cell.raw.includes('[') ? tableAt(sheet, r, c)?.name : null, dyn))}</f>`;
          if (typeof v === 'number') return `<c r="${ref}"${sAttr}${cm}>${f}<v>${v}</v></c>`;
          if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"${cm}>${f}<v>${v ? 1 : 0}</v></c>`;
          if (isError(v)) return `<c r="${ref}"${sAttr} t="e"${cm}>${f}<v>${esc(['#CIRC!', '#SPILL!', '#CALC!', '#BUSY!'].includes(v.code) && !dyn ? '#REF!' : v.code === '#CIRC!' ? '#REF!' : v.code)}</v></c>`;
          return `<c r="${ref}"${sAttr} t="str"${cm}>${f}<v>${esc(v ?? '')}</v></c>`;
        }
        if (isError(v)) return `<c r="${ref}"${sAttr} t="e"><v>${esc(v.code)}</v></c>`;
        if (typeof v === 'number') return `<c r="${ref}"${sAttr}><v>${v}</v></c>`;
        if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b"><v>${v ? 1 : 0}</v></c>`;
        if (v === null) return s ? `<c r="${ref}"${sAttr}/>` : '';
        return `<c r="${ref}"${sAttr} t="s"><v>${sst(String(v))}</v></c>`;
      }).join('');
      return `<row ${attrs.join(' ')}>${cx}</row>`;
    }).join('');

    // 열
    const colKeys = new Set([...Object.keys(sheet.colWidths), ...Object.keys(sheet.hiddenCols), ...Object.keys(sheet.colStyles)].map(Number));
    const colsXml = [...colKeys].sort((a, b) => a - b).map((c) => {
      const w = sheet.colWidths[c] ?? DEFAULT_COL_WIDTH;
      const st = sheet.colStyles[c] ? ` style="${pool.xf({ ...sheet.allStyle, ...sheet.colStyles[c] })}"` : '';
      return `<col min="${c + 1}" max="${c + 1}" width="${px2width(w)}"${sheet.colWidths[c] !== undefined ? ' customWidth="1"' : ''}${sheet.hiddenCols[c] ? ' hidden="1"' : ''}${st}/>`;
    }).join('');

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
      if (sl.source?.kind === 'table') {
        const f = findTable(wb, sl.source.table);
        const tid = f && tableIds.get(f.t.name.toLowerCase());
        if (!tid) continue;
        const names = columnNames(wb, f.si, f.t).map((n) => n.toLowerCase());
        const col = names.indexOf(String(sl.source.column).toLowerCase());
        if (col < 0) continue;
        kind = 'table';
        sl._cache = cacheNameFor(sl.source.column, usedCacheNames);
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(sl._cache)}" sourceName="${esc(columnNames(wb, f.si, f.t)[col])}"><extLst><x:ext uri="{2F2917AC-EB37-4324-AD4E-5DD8C200BD13}" xmlns:x15="${NS_X15}"><x15:tableSlicerCache tableId="${tid}" column="${col + 1}"/></x:ext></extLst></slicerCacheDefinition>`;
      } else {
        const pi = slicerPivotSheet(wb, si, sl);
        const info = pivotInfo.get(pi);
        const it = info?.parts.slicerItems(sl.source.field);
        if (!it) continue;
        kind = 'pivot';
        sl._cache = cacheNameFor(it.field, usedCacheNames);
        cacheXml = `<slicerCacheDefinition xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}" name="${esc(sl._cache)}" sourceName="${esc(it.field)}"><pivotTables><pivotTable tabId="${pi + 1}" name="${esc(info.name)}"/></pivotTables><data><tabular pivotCacheId="${info.cacheId}"><items count="${it.count}">${it.xml}</items></tabular></data></slicerCacheDefinition>`;
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
      const colAxis = new Axis(DEFAULT_COL_WIDTH, sheet.colWidths, [], MAX_COLS);
      const rowAxis = new Axis(DEFAULT_ROW_HEIGHT, sheet.rowHeights, [], MAX_ROWS);
      const anchorAt = (x, y) => {
        const c = colAxis.indexAt(Math.max(0, x));
        const r = rowAxis.indexAt(Math.max(0, y));
        return `<xdr:col>${c}</xdr:col><xdr:colOff>${Math.max(0, Math.round((x - colAxis.pos(c)) * EMU))}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${Math.max(0, Math.round((y - rowAxis.pos(r)) * EMU))}</xdr:rowOff>`;
      };
      const anchor = (o, body, editAs = 'oneCell') => `<xdr:twoCellAnchor editAs="${editAs}"><xdr:from>${anchorAt(o.x, o.y)}</xdr:from><xdr:to>${anchorAt(o.x + o.w, o.y + o.h)}</xdr:to>${body}<xdr:clientData/></xdr:twoCellAnchor>`;
      const xfrm = (o) => `<a:xfrm${o.flip ? ' flipH="1"' : ''}${o.flipV ? ' flipV="1"' : ''}><a:off x="${Math.round(o.x * EMU)}" y="${Math.round(o.y * EMU)}"/><a:ext cx="${Math.round(o.w * EMU)}" cy="${Math.round(o.h * EMU)}"/></a:xfrm>`;
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
          files[`xl/charts/chart${chartNo}.xml`] = chartXml(wb, si, ch);
          contentOverrides.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
          const id = drel('chart', `../charts/chart${chartNo}.xml`);
          parts.push(anchor(ch, `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${objId}" name="차트 ${objId - 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="${id}"/></a:graphicData></a:graphic></xdr:graphicFrame>`));
        } else if (kind === 'image') {
          const im = o;
          const m = /^data:([^;,]+);base64,(.*)$/s.exec(im.src ?? '');
          if (!m) continue;
          const ext = Object.keys(MIME).find((k) => MIME[k] === m[1]) ?? 'png';
          mediaNo++;
          mediaExts.add(ext);
          files[`xl/media/image${mediaNo}.${ext}`] = fromBase64(m[2]);
          const id = drel('image', `../media/image${mediaNo}.${ext}`);
          objId++;
          parts.push(anchor(im, `<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${objId}" name="${esc(im.name || `그림 ${objId - 1}`)}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${id}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>${xfrm(im)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>`));
        } else if (kind === 'slicerTable' || kind === 'slicerPivot') {
          objId++;
          parts.push(slicerAnchorXml(o.sl, o.name, objId, anchorAt, kind === 'slicerTable' ? 'table' : 'pivot'));
        } else {
          objId++;
          parts.push(anchor(o, shapeXml(o, objId, xfrm), 'twoCell'));
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

    // 메모
    let legacy = '';
    const comments = [...sheet.cells].filter(([, c]) => c.comment).map(([k, c]) => [k.split(',').map(Number), c.comment]);
    if (comments.length) {
      commentNo++;
      files[`xl/comments${commentNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${NS_MAIN}"><authors><author>Tabula</author></authors><commentList>${comments.map(([[r, c], text]) => `<comment ref="${cellName(r, c)}" authorId="0"><text><r><t xml:space="preserve">${esc(text)}</t></r></text></comment>`).join('')}</commentList></comments>`;
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
      files[`xl/tables/table${tableNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${NS_MAIN}" id="${tableNo}" name="${esc(t.name)}" displayName="${esc(t.name)}" ref="${rangeRef(t)}"${t.header ? '' : ' headerRowCount="0"'}${t.totals ? ' totalsRowCount="1"' : ' totalsRowShown="0"'}>${af}<tableColumns count="${names.length}">${cols}</tableColumns><tableStyleInfo${style ? ` name="${style}"` : ''} showFirstColumn="${t.firstCol ? 1 : 0}" showLastColumn="${t.lastCol ? 1 : 0}" showRowStripes="${t.banded !== false ? 1 : 0}" showColumnStripes="${t.bandedCols ? 1 : 0}"/></table>`;
      contentOverrides.push(`<Override PartName="/xl/tables/table${tableNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>`);
      tableRids.push(addRel('table', `../tables/table${tableNo}.xml`));
    }
    const tableParts = tableRids.length ? `<tableParts count="${tableRids.length}">${tableRids.map((id) => `<tablePart r:id="${id}"/>`).join('')}</tableParts>` : '';

    // 피벗 테이블
    for (const pinfo of pivotList.get(si) ?? []) {
      const n = pinfo.cacheId;
      files[`xl/pivotCache/pivotCacheDefinition${n}.xml`] = pinfo.parts.cacheXml;
      files[`xl/pivotTables/pivotTable${n}.xml`] = pinfo.parts.tableXml;
      files[`xl/pivotTables/_rels/pivotTable${n}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition${n}.xml"/></Relationships>`;
      contentOverrides.push(`<Override PartName="/xl/pivotCache/pivotCacheDefinition${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>`);
      contentOverrides.push(`<Override PartName="/xl/pivotTables/pivotTable${n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>`);
      pivotCaches.push({ cacheId: n, target: `pivotCache/pivotCacheDefinition${n}.xml` });
      addRel('pivotTable', `../pivotTables/pivotTable${n}.xml`);
    }

    // 슬라이서 목록 (피벗용 x14, 표용 x15)
    const exts = [];
    for (const kind of ['pivot', 'table']) {
      const list = sheetSlicers[kind];
      if (!list.length) continue;
      slicerPartNo++;
      files[`xl/slicers/slicer${slicerPartNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<slicers xmlns="${NS_X14}" xmlns:mc="${NS_MC}" mc:Ignorable="x" xmlns:x="${NS_MAIN}">${list.map(({ sl, name, cache }) => `<slicer name="${esc(name)}" cache="${esc(cache)}" caption="${esc(sl.caption ?? name)}"${(sl.columns ?? 1) > 1 ? ` columnCount="${sl.columns}"` : ''}${sl.color && sl.color !== 'blue' ? ` style="${SLICER_STYLE[sl.color] ?? 'SlicerStyleLight1'}"` : ''} rowHeight="241300"/>`).join('')}</slicers>`;
      contentOverrides.push(`<Override PartName="/xl/slicers/slicer${slicerPartNo}.xml" ContentType="application/vnd.ms-excel.slicer+xml"/>`);
      const id = `rId${sheetRels.length + 1}`;
      sheetRels.push(`<Relationship Id="${id}" Type="${REL_MS}/slicer" Target="../slicers/slicer${slicerPartNo}.xml"/>`);
      exts.push(kind === 'pivot'
        ? `<ext uri="{A8765BA9-456A-4dab-B4F3-ACF838C121DE}" xmlns:x14="${NS_X14}"><x14:slicerList><x14:slicer r:id="${id}"/></x14:slicerList></ext>`
        : `<ext uri="{3A4CF648-6AED-40f4-86FF-DC5316D8AED3}" xmlns:x15="${NS_X15}"><x14:slicerList xmlns:x14="${NS_X14}"><x14:slicer r:id="${id}"/></x14:slicerList></ext>`);
    }
    if (cfX14.length) exts.unshift(`<ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${NS_X14}"><x14:conditionalFormattings>${cfX14.join('')}</x14:conditionalFormattings></ext>`);
    const extLst = exts.length ? `<extLst>${exts.join('')}</extLst>` : '';

    files[`xl/worksheets/sheet${si + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`
      + (vba ? `<sheetPr codeName="${esc(vba.sheetCodes?.[sheet.name] ?? `Sheet${si + 1}`)}"/>` : '')
      + `<dimension ref="${dim}"/>`
      + `<sheetViews><sheetView workbookViewId="0"${si === activeSheet ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>`
      + `<sheetFormatPr defaultRowHeight="${px2pt(DEFAULT_ROW_HEIGHT)}"/>`
      + (colsXml ? `<cols>${colsXml}</cols>` : '')
      + `<sheetData>${rowXml}</sheetData>`
      + autoFilter + merges + cf + dataValidations
      + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      + drawing + legacy + tableParts + extLst
      + '</worksheet>';
    if (sheetRels.length) {
      files[`xl/worksheets/_rels/sheet${si + 1}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${sheetRels.join('')}</Relationships>`;
    }
  });

  // 통합 문서: 피벗 캐시, 슬라이서 캐시 (확장)
  const wbRels = [
    ...wb.sheets.map((sh, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`),
    `<Relationship Id="rId${wb.sheets.length + 1}" Type="${REL}/styles" Target="styles.xml"/>`,
    `<Relationship Id="rId${wb.sheets.length + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/>`,
  ];
  if (vba) wbRels.push(`<Relationship Id="rId${wbRels.length + 1}" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`);
  if (dynamicCells) {
    // 동적 배열 수식 표시 (cm="1" 이 가리키는 셀 메타데이터)
    wbRels.push(`<Relationship Id="rId${wbRels.length + 1}" Type="${REL}/sheetMetadata" Target="metadata.xml"/>`);
    files['xl/metadata.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<metadata xmlns="${NS_MAIN}" xmlns:xda="http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray"><metadataTypes count="1"><metadataType name="XLDAPR" minSupportedVersion="120000" copy="1" pasteAll="1" pasteValues="1" merge="1" splitFirst="1" rowColShift="1" clearFormats="1" clearComments="1" assign="1" coerce="1" cellMeta="1"/></metadataTypes><futureMetadata name="XLDAPR" count="1"><bk><extLst><ext uri="{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}"><xda:dynamicArrayProperties fDynamic="1" fCollapsed="0"/></ext></extLst></bk></futureMetadata><cellMetadata count="1"><bk><rc t="1" v="0"/></bk></cellMetadata></metadata>`;
    contentOverrides.push('<Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>');
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
  files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}">${vba ? `<workbookPr codeName="${esc(vba.codeName || 'ThisWorkbook')}"/>` : ''}<bookViews><workbookView activeTab="${activeSheet}"/></bookViews><sheets>${wb.sheets.map((sh, i) => `<sheet name="${esc(sh.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${definedNames.length ? `<definedNames>${definedNames.join('')}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/>${pivotCachesXml}${wbExts.length ? `<extLst>${wbExts.join('')}</extLst>` : ''}</workbook>`;
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${wbRels.join('')}</Relationships>`;
  if (vba) files['xl/vbaProject.bin'] = fromBase64(vba.bin);
  files['xl/sharedStrings.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`;
  files['xl/styles.xml'] = pool.xml();
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  files['docProps/core.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Tabula</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  files['docProps/app.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Tabula</Application></Properties>`;
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>${[...mediaExts].map((e) => `<Default Extension="${e}" ContentType="${MIME[e]}"/>`).join('')}${vba ? '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/>' : ''}<Override PartName="/xl/workbook.xml" ContentType="${vba ? 'application/vnd.ms-excel.sheet.macroEnabled.main+xml' : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'}"/>${wb.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${contentOverrides.join('')}</Types>`;

  // [Content_Types].xml 을 맨 앞에 두는 것이 관례
  const ordered = { '[Content_Types].xml': files['[Content_Types].xml'] };
  for (const [k, v] of Object.entries(files)) if (k !== '[Content_Types].xml') ordered[k] = v;
  return zip(ordered);
}
