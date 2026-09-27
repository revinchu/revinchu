// .xlsx 읽기/쓰기 (Office Open XML). DOM 없이 동작하므로 Node 에서도 테스트 가능.
import { unzip, zip, textOf } from './zip.js';
import { parseXml, child, kids, descendants, allText, esc } from './xml.js';
import {
  parse, tokenize, shiftFormula, colToName, nameToCol, cellName, parseRangeName, FUNCS, isError,
  quoteSheetName, MAX_ROWS, MAX_COLS,
} from './formula.js';
import { parseInput, formatGeneral } from './format.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { chartLayout, PALETTE } from './chart.js';
import { Axis } from './axis.js';

const NS_MAIN = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const NS_PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const EMU = 9525; // 1px

const DEFAULT_FONT = '맑은 고딕';
const XLFN = new Set(['IFS', 'XLOOKUP', 'CONCAT', 'TEXTJOIN', 'IFNA']);

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

export function fmtFromCode(code) {
  const first = code.split(';')[0];
  const plain = first.replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[[^\]]*\]/g, (m) => (/\$[₩$€¥£]/.test(m) ? '₩' : ''));
  const decimals = (plain.match(/\.(0+)/)?.[1].length) || undefined;
  if (plain.trim() === '@') return { numFmt: 'text' };
  if (/general/i.test(plain) && !/[0#]/.test(plain)) return {};
  if (/E[+-]/i.test(plain)) return { numFmt: 'scientific', decimals: decimals ?? 0 };
  const hasDate = /[yd]/i.test(plain) || /(^|[^a-z])m{1,4}([^a-z]|$)/i.test(plain) && !/h/i.test(plain);
  const hasTime = /[hs]/i.test(plain);
  if (hasDate || hasTime) {
    if (hasDate && hasTime) return { numFmt: 'datetime' };
    if (hasTime) return { numFmt: 'time' };
    return { numFmt: /dddd|aaaa/i.test(plain) || /년/.test(code) ? 'longdate' : 'date' };
  }
  if (plain.includes('%')) return { numFmt: 'percent', decimals };
  if (/[₩$€¥£]/.test(plain) || /[₩$€¥£]/.test(code.split(';')[0].replace(/"/g, ''))) return { numFmt: 'currency', decimals };
  if (/\?\/\?/.test(plain)) return { numFmt: 'fraction' };
  if (plain.includes('#,##0') || plain.includes('#,###')) return decimals ? { numFmt: 'number', decimals } : { numFmt: 'comma' };
  if (/0/.test(plain)) return { decimals: decimals ?? 0 };
  return {};
}

function fmtCode(style) {
  const d = style.decimals;
  const dec = (n) => (n ? `.${'0'.repeat(n)}` : '');
  switch (style.numFmt) {
    case 'number': return `#,##0${dec(d ?? 2)}`;
    case 'comma': return `#,##0${dec(d ?? 0)}`;
    case 'currency': case 'accounting': return `"₩"#,##0${dec(d ?? 0)};\\-"₩"#,##0${dec(d ?? 0)}`;
    case 'percent': return `0${dec(d ?? 0)}%`;
    case 'scientific': return `0${dec(d ?? 2)}E+00`;
    case 'fraction': return '# ?/?';
    case 'date': return 'yyyy\\-mm\\-dd';
    case 'longdate': return 'yyyy"년" m"월" d"일" dddd';
    case 'time': return '[$-412]AM/PM h:mm:ss';
    case 'datetime': return 'yyyy\\-mm\\-dd h:mm';
    case 'text': return '@';
    default: return d !== undefined ? `0${dec(d)}` : null;
  }
}

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
    if (numFmts[id]) return fmtFromCode(numFmts[id]);
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
      if (child(f, 'b')) st.bold = true;
      if (child(f, 'i')) st.italic = true;
    }
    const fill = child(d, 'fill');
    if (fill) { const c = fillOf(fill, true); if (c) st.fill = c; }
    return st;
  });
  return { xfs, dxfs, defaultFont };
}

/** 수식 텍스트 정리: _xlfn. 접두사 제거 */
function cleanFormula(f) {
  return f.replace(/_xlfn\.|_xlws\.|_xlpm\./gi, '');
}

function formulaSupported(text) {
  try {
    const ast = parse(text);
    const walk = (n) => {
      if (!n) return true;
      if (n.type === 'func' && !FUNCS[n.name]) return false;
      if (n.type === 'name') return false;
      return [n.a, n.b, ...(n.args ?? [])].every((x) => x === undefined || walk(x));
    };
    return walk(ast);
  } catch {
    return false;
  }
}

function readSheet(files, path, ctx) {
  const root = parseXml(textOf(files[path]));
  const sheet = {
    cells: {}, colWidths: {}, rowHeights: {}, merges: [], cond: [], colStyles: {}, rowStyles: {},
    hiddenRows: {}, hiddenCols: {}, rowManual: {}, freeze: { rows: 0, cols: 0 }, filter: null, charts: [],
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
      if (formula !== null) {
        const f = cleanFormula(formula);
        if (formulaSupported(f)) raw = `=${f}`;
        else { unsupported++; formula = null; }
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

  // 조건부 서식
  for (const cf of kids(root, 'conditionalFormatting')) {
    for (const sq of (cf.attrs.sqref ?? '').split(/\s+/)) {
      const rg = refToRange(sq);
      if (!rg) continue;
      for (const rule of kids(cf, 'cfRule')) {
        const a = rule.attrs;
        const style = dxfs[Number(a.dxfId)] ?? { fill: '#ffc7ce', color: '#9c0006' };
        const formulas = kids(rule, 'formula').map((f) => f.text.replace(/^"(.*)"$/, '$1'));
        let out = null;
        if (a.type === 'cellIs') {
          const op = { greaterThan: 'gt', lessThan: 'lt', between: 'between', equal: 'eq' }[a.operator];
          if (op) out = { type: op, v1: formulas[0], v2: formulas[1], style };
        } else if (a.type === 'containsText') out = { type: 'text', v1: a.text ?? '', style };
        else if (a.type === 'duplicateValues') out = { type: 'dup', style };
        else if (a.type === 'uniqueValues') out = { type: 'unique', style };
        else if (a.type === 'top10' && a.bottom !== '1') out = { type: 'top', v1: a.rank ?? '10', style };
        else if (a.type === 'aboveAverage') out = { type: a.aboveAverage === '0' ? 'belowAvg' : 'aboveAvg', style };
        else if (a.type === 'dataBar') {
          const color = colorOf(descendants(rule, 'color')[0], ctx.theme) ?? '#638ec6';
          out = { type: 'bar', color: color.toLowerCase() };
        } else if (a.type === 'colorScale') {
          const colors = descendants(child(rule, 'colorScale'), 'color').map((c) => colorOf(c, ctx.theme)).filter(Boolean);
          if (colors.length >= 2) out = { type: 'scale', colors };
        }
        if (out) sheet.cond.push({ ...rg, ...out });
      }
    }
  }

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
  if (drawing && rels[rid(drawing)]) {
    sheet.charts = readDrawing(files, rels[rid(drawing)].target, sheet, ctx);
  }
  return sheet;
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

function readDrawing(files, path, sheet, ctx) {
  const xml = textOf(files[path]);
  if (!xml) return [];
  const root = parseXml(xml);
  const rels = relsOf(files, path);
  const colAxis = new Axis(DEFAULT_COL_WIDTH, sheet.colWidths, [], MAX_COLS);
  const rowAxis = new Axis(DEFAULT_ROW_HEIGHT, sheet.rowHeights, [], MAX_ROWS);
  const point = (el) => {
    const n = (name) => Number(child(el, name)?.text ?? 0);
    return { x: colAxis.pos(n('col')) + n('colOff') / EMU, y: rowAxis.pos(n('row')) + n('rowOff') / EMU };
  };
  const charts = [];
  for (const anchor of root.children) {
    const chartRef = descendants(anchor, 'chart')[0];
    if (!chartRef || !rels[rid(chartRef)]) continue;
    let x;
    let y;
    let w;
    let h;
    if (anchor.name === 'twoCellAnchor') {
      const a = point(child(anchor, 'from'));
      const b = point(child(anchor, 'to'));
      ({ x, y } = a);
      w = b.x - a.x;
      h = b.y - a.y;
    } else {
      const from = child(anchor, 'from');
      const pos = child(anchor, 'pos');
      const ext = child(anchor, 'ext');
      ({ x, y } = from ? point(from) : { x: Number(pos?.attrs.x ?? 0) / EMU, y: Number(pos?.attrs.y ?? 0) / EMU });
      w = Number(ext?.attrs.cx ?? 480 * EMU) / EMU;
      h = Number(ext?.attrs.cy ?? 288 * EMU) / EMU;
    }
    const chart = readChart(files, rels[rid(chartRef)].target, ctx);
    if (chart) charts.push({ id: `c${charts.length + 1}${Date.now().toString(36)}`, ...chart, x: Math.round(x), y: Math.round(y), w: Math.round(Math.max(120, w)), h: Math.round(Math.max(90, h)) });
  }
  return charts;
}

function readChart(files, path, ctx) {
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
  void ctx;
  return out;
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
  const ctx = { xfs, dxfs, strings, theme };
  const sheets = [];
  const warnings = [];
  let unsupported = 0;
  for (const sh of kids(child(wbRoot, 'sheets'), 'sheet')) {
    const rel = wbRels[rid(sh)];
    if (!rel || rel.type !== 'worksheet' || !files[rel.target]) {
      if (rel) warnings.push(`'${sh.attrs.name}' 시트(차트 시트 등)는 가져오지 않았습니다.`);
      continue;
    }
    const sheet = readSheet(files, rel.target, ctx);
    unsupported += sheet.unsupported;
    delete sheet.unsupported;
    sheets.push({ name: sh.attrs.name.slice(0, 31), ...sheet });
  }
  if (unsupported) warnings.push(`지원하지 않는 수식 ${unsupported}개는 계산된 값으로 가져왔습니다.`);
  if (!sheets.length) throw new Error('가져올 시트가 없습니다');
  const active = Number(descendants(child(wbRoot, 'bookViews'), 'workbookView')[0]?.attrs.activeTab ?? 0);
  return { data: { sheets }, active: Math.min(active, sheets.length - 1), warnings };
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
    const font = style.color || style.bold || style.italic ? `<font>${style.bold ? '<b/>' : ''}${style.italic ? '<i/>' : ''}${style.color ? `<color rgb="${argb(style.color)}"/>` : ''}</font>` : '';
    const fill = style.fill ? `<fill><patternFill><bgColor rgb="${argb(style.fill)}"/></patternFill></fill>` : '';
    this.dxfs.push(`<dxf>${font}${fill}</dxf>`);
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

/** 새 함수는 엑셀 파일에서 _xlfn. 접두사가 필요 */
function exportFormula(raw) {
  const body = raw.slice(1);
  let toks;
  try { toks = tokenize(body); } catch { return body; }
  let out = '';
  let pos = 0;
  for (const t of toks) {
    if (t.t === 'func' && XLFN.has(t.v)) {
      out += `${body.slice(pos, t.s)}_xlfn.${body.slice(t.s, t.e)}`;
      pos = t.e;
    }
  }
  return out + body.slice(pos);
}

function cfXml(rule, pool, priority) {
  const ref = rangeRef(rule);
  const lit = (v) => {
    const p = parseInput(String(v ?? ''));
    return typeof p.value === 'number' ? String(p.value) : `"${String(v ?? '').replace(/"/g, '""')}"`;
  };
  const dx = () => pool.dxf(rule.style ?? {});
  let body;
  switch (rule.type) {
    case 'gt': case 'lt': case 'eq':
      body = `<cfRule type="cellIs" dxfId="${dx()}" priority="${priority}" operator="${{ gt: 'greaterThan', lt: 'lessThan', eq: 'equal' }[rule.type]}"><formula>${esc(lit(rule.v1))}</formula></cfRule>`;
      break;
    case 'between':
      body = `<cfRule type="cellIs" dxfId="${dx()}" priority="${priority}" operator="between"><formula>${esc(lit(rule.v1))}</formula><formula>${esc(lit(rule.v2))}</formula></cfRule>`;
      break;
    case 'text':
      body = `<cfRule type="containsText" dxfId="${dx()}" priority="${priority}" operator="containsText" text="${esc(rule.v1 ?? '')}"><formula>${esc(`NOT(ISERROR(SEARCH(${lit(String(rule.v1 ?? ''))},${cellName(rule.r1, rule.c1)})))`)}</formula></cfRule>`;
      break;
    case 'dup': body = `<cfRule type="duplicateValues" dxfId="${dx()}" priority="${priority}"/>`; break;
    case 'unique': body = `<cfRule type="uniqueValues" dxfId="${dx()}" priority="${priority}"/>`; break;
    case 'top': body = `<cfRule type="top10" dxfId="${dx()}" priority="${priority}" rank="${Number(rule.v1) || 10}"/>`; break;
    case 'aboveAvg': body = `<cfRule type="aboveAverage" dxfId="${dx()}" priority="${priority}"/>`; break;
    case 'belowAvg': body = `<cfRule type="aboveAverage" dxfId="${dx()}" priority="${priority}" aboveAverage="0"/>`; break;
    case 'bar':
      body = `<cfRule type="dataBar" priority="${priority}"><dataBar><cfvo type="min"/><cfvo type="max"/><color rgb="${argb(rule.color)}"/></dataBar></cfRule>`;
      break;
    case 'scale': {
      const cs = rule.colors;
      const cfvo = cs.length === 3 ? '<cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' : '<cfvo type="min"/><cfvo type="max"/>';
      body = `<cfRule type="colorScale" priority="${priority}"><colorScale>${cfvo}${cs.map((c) => `<color rgb="${argb(c)}"/>`).join('')}</colorScale></cfRule>`;
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

/**
 * Workbook → xlsx 바이트
 */
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
  const definedNames = [];

  wb.sheets.forEach((sheet, si) => {
    const sheetRels = [];
    const addRel = (type, target) => { const id = `rId${sheetRels.length + 1}`; sheetRels.push(`<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`); return id; };

    // 셀을 행별로 정리
    const rows = new Map();
    let maxR = 0;
    let maxC = 0;
    for (const [k, cell] of sheet.cells) {
      const [r, c] = k.split(',').map(Number);
      if (!rows.has(r)) rows.set(r, []);
      rows.get(r).push([c, cell]);
      maxR = Math.max(maxR, r);
      maxC = Math.max(maxC, c);
    }
    const rowKeys = new Set([...rows.keys()]);
    for (const k of [...Object.keys(sheet.rowHeights), ...Object.keys(sheet.hiddenRows), ...Object.keys(sheet.rowStyles), ...Object.keys(sheet.filter?.hidden ?? {})]) rowKeys.add(Number(k));
    const sortedRows = [...rowKeys].sort((a, b) => a - b);

    const rowXml = sortedRows.map((r) => {
      const cells = (rows.get(r) ?? []).sort((a, b) => a[0] - b[0]);
      const attrs = [`r="${r + 1}"`];
      if (sheet.rowHeights[r] !== undefined) attrs.push(`ht="${px2pt(sheet.rowHeights[r])}"`, 'customHeight="1"');
      if (sheet.hiddenRows[r] || sheet.filter?.hidden?.[r]) attrs.push('hidden="1"');
      if (sheet.rowStyles[r]) attrs.push(`s="${pool.xf({ ...sheet.allStyle, ...sheet.rowStyles[r] })}"`, 'customFormat="1"');
      const cx = cells.map(([c, cell]) => {
        const ref = cellName(r, c);
        const st = wb.styleAt(si, r, c);
        const s = pool.xf(st);
        const sAttr = s ? ` s="${s}"` : '';
        if (!cell.raw) return s ? `<c r="${ref}"${sAttr}/>` : '';
        const v = wb.getValue(si, r, c);
        if (cell.formula) {
          const f = `<f>${esc(exportFormula(cell.raw))}</f>`;
          if (typeof v === 'number') return `<c r="${ref}"${sAttr}>${f}<v>${v}</v></c>`;
          if (typeof v === 'boolean') return `<c r="${ref}"${sAttr} t="b">${f}<v>${v ? 1 : 0}</v></c>`;
          if (isError(v)) return `<c r="${ref}"${sAttr} t="e">${f}<v>${esc(v.code === '#CIRC!' ? '#REF!' : v.code)}</v></c>`;
          return `<c r="${ref}"${sAttr} t="str">${f}<v>${esc(v ?? '')}</v></c>`;
        }
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
    if (sheet.filter) {
      const f = sheet.filter;
      const cols = Object.entries(f.criteria ?? {}).filter(([, vals]) => Array.isArray(vals)).map(([c, vals]) => {
        const blank = vals.includes('');
        return `<filterColumn colId="${Number(c) - f.c1}"><filters${blank ? ' blank="1"' : ''}>${vals.filter((v) => v !== '').map((v) => `<filter val="${esc(v)}"/>`).join('')}</filters></filterColumn>`;
      }).join('');
      autoFilter = `<autoFilter ref="${rangeRef(f)}">${cols}</autoFilter>`;
      definedNames.push(`<definedName name="_xlnm._FilterDatabase" localSheetId="${si}" hidden="1">${esc(`${quoteSheetName(sheet.name)}!${rangeRef(f, true)}`)}</definedName>`);
    }
    const merges = sheet.merges.length ? `<mergeCells count="${sheet.merges.length}">${sheet.merges.map((m) => `<mergeCell ref="${rangeRef(m)}"/>`).join('')}</mergeCells>` : '';
    const cf = sheet.cond.map((rule, i) => cfXml(rule, pool, i + 1)).join('');

    // 차트
    let drawing = '';
    if (sheet.charts.length) {
      drawingNo++;
      const drawingRels = [];
      const colAxis = new Axis(DEFAULT_COL_WIDTH, sheet.colWidths, [], MAX_COLS);
      const rowAxis = new Axis(DEFAULT_ROW_HEIGHT, sheet.rowHeights, [], MAX_ROWS);
      const anchorAt = (x, y) => {
        const c = colAxis.indexAt(x);
        const r = rowAxis.indexAt(y);
        return `<xdr:col>${c}</xdr:col><xdr:colOff>${Math.round((x - colAxis.pos(c)) * EMU)}</xdr:colOff><xdr:row>${r}</xdr:row><xdr:rowOff>${Math.round((y - rowAxis.pos(r)) * EMU)}</xdr:rowOff>`;
      };
      const anchors = sheet.charts.map((ch, i) => {
        chartNo++;
        files[`xl/charts/chart${chartNo}.xml`] = chartXml(wb, si, ch);
        contentOverrides.push(`<Override PartName="/xl/charts/chart${chartNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>`);
        drawingRels.push(`<Relationship Id="rId${i + 1}" Type="${REL}/chart" Target="../charts/chart${chartNo}.xml"/>`);
        return `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${anchorAt(ch.x, ch.y)}</xdr:from><xdr:to>${anchorAt(ch.x + ch.w, ch.y + ch.h)}</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${i + 2}" name="차트 ${i + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId${i + 1}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
      }).join('');
      files[`xl/drawings/drawing${drawingNo}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">${anchors}</xdr:wsDr>`;
      files[`xl/drawings/_rels/drawing${drawingNo}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${drawingRels.join('')}</Relationships>`;
      contentOverrides.push(`<Override PartName="/xl/drawings/drawing${drawingNo}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>`);
      drawing = `<drawing r:id="${addRel('drawing', `../drawings/drawing${drawingNo}.xml`)}"/>`;
    }

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

    files[`xl/worksheets/sheet${si + 1}.xml`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<worksheet xmlns="${NS_MAIN}" xmlns:r="${NS_R}">`
      + `<dimension ref="${dim}"/>`
      + `<sheetViews><sheetView workbookViewId="0"${si === activeSheet ? ' tabSelected="1"' : ''}>${pane}</sheetView></sheetViews>`
      + `<sheetFormatPr defaultRowHeight="${px2pt(DEFAULT_ROW_HEIGHT)}"/>`
      + (colsXml ? `<cols>${colsXml}</cols>` : '')
      + `<sheetData>${rowXml}</sheetData>`
      + autoFilter + merges + cf
      + '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
      + drawing + legacy
      + '</worksheet>';
    if (sheetRels.length) {
      files[`xl/worksheets/_rels/sheet${si + 1}.xml.rels`] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${sheetRels.join('')}</Relationships>`;
    }
  });

  files['xl/workbook.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<workbook xmlns="${NS_MAIN}" xmlns:r="${NS_R}"><bookViews><workbookView activeTab="${activeSheet}"/></bookViews><sheets>${wb.sheets.map((s, i) => `<sheet name="${esc(s.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')}</sheets>${definedNames.length ? `<definedNames>${definedNames.join('')}</definedNames>` : ''}<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>`;
  files['xl/_rels/workbook.xml.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}">${wb.sheets.map((s, i) => `<Relationship Id="rId${i + 1}" Type="${REL}/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')}<Relationship Id="rId${wb.sheets.length + 1}" Type="${REL}/styles" Target="styles.xml"/><Relationship Id="rId${wb.sheets.length + 2}" Type="${REL}/sharedStrings" Target="sharedStrings.xml"/></Relationships>`;
  files['xl/sharedStrings.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<sst xmlns="${NS_MAIN}" count="${strings.length}" uniqueCount="${strings.length}">${strings.map((s) => `<si><t xml:space="preserve">${esc(s)}</t></si>`).join('')}</sst>`;
  files['xl/styles.xml'] = pool.xml();
  files['_rels/.rels'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${REL}/officeDocument" Target="xl/workbook.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${REL}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z');
  files['docProps/core.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Tabula</dc:creator><dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`;
  files['docProps/app.xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>Tabula</Application></Properties>`;
  files['[Content_Types].xml'] = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>${wb.sheets.map((s, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')}<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>${contentOverrides.join('')}</Types>`;

  // [Content_Types].xml 을 맨 앞에 두는 것이 관례
  const ordered = { '[Content_Types].xml': files['[Content_Types].xml'] };
  for (const [k, v] of Object.entries(files)) if (k !== '[Content_Types].xml') ordered[k] = v;
  return zip(ordered);
}
