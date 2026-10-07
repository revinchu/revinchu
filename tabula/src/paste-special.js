// 선택하여 붙여넣기: 복사 시점의 데이터/서식을 보존하고 단일 트랜잭션에서 적용한다.
import { cellData, makeCellRC, adjustRange } from './workbook.js';
import { snapshotCopiedRows, copiedRowTailColumns, applyCopiedRowMetadata } from './copied-lines.js';
import { rewriteWorkbookLink } from './hyperlink.js';
import { tableCellDisplayStyle } from './table-format.js';
import { shiftFormula, cellName, quoteSheetName, MAX_ROWS, MAX_COLS, EXCEL_MAX_ROWS, isError, rewriteRefs, adjustFormulaForStructure } from './formula.js';
import { parseDelimited } from './csv.js';
import { parseInput } from './format.js';
import { cellStyleDefaults, CELL_STYLE_PARTS } from './cell-style.js';

export const PASTE_SPECIAL_TYPES = [
  ['all', '모두', 'a'], ['formulas', '수식', 'f'], ['values', '값', 'v'], ['formats', '서식', 't'], ['comments', '메모', 'c'], ['validation', '유효성 검사', 'n'],
  ['sourceTheme', '원본 테마 사용', 'h'], ['noBorders', '테두리만 제외', 'x'], ['colWidths', '열 너비', 'w'], ['formulasNum', '수식 및 숫자 서식', 'r'], ['valuesNum', '값 및 숫자 서식', 'u'], ['mergeCond', '조건부 서식 병합', 'g'],
];
export const PASTE_SPECIAL_OPERATIONS = [['', '없음', 'o'], ['add', '더하기', 'd'], ['sub', '빼기', 's'], ['mul', '곱하기', 'm'], ['div', '나누기', 'i']];
const ALL_TYPES = new Set(['all', 'sourceTheme', 'noBorders', 'mergeCond']);
const NUMBER_KEYS = CELL_STYLE_PARTS.find(x => x[0] === 'number')[2];
const BORDER_KEYS = CELL_STYLE_PARTS.find(x => x[0] === 'border')[2];
const clone = x => x == null ? x : structuredClone(x);
const intersect = (a, b) => { const n = { r1: Math.max(a.r1, b.r1), c1: Math.max(a.c1, b.c1), r2: Math.min(a.r2, b.r2), c2: Math.min(a.c2, b.c2) }; return n.r1 <= n.r2 && n.c1 <= n.c2 ? n : null; };
const rangesOf = rule => [rule, ...(rule.more ?? [])].map(({ r1, c1, r2, c2 }) => ({ r1, c1, r2, c2 }));
const hasFormula = d => d?.raw?.startsWith('=') && d.inputType !== 'text' && (d.style?.numFmt !== 'text' || d.fx);
const blank = d => !d || d.raw === '' && !d.image;

export function snapshotPasteSource(wb, source) {
  const sh = wb.sheets[source.si], rows = source.rows ?? source.data.map((_, i) => source.r1 + i);
  const range = { r1: rows[0], r2: rows.at(-1), c1: source.c1, c2: source.kind === 'rows' ? MAX_COLS - 1 : source.c2 };
  const copiedRows = snapshotCopiedRows(wb, { ...source, rows });
  const styles = new Map();
  const styleAt = (r, c) => { const effective = tableCellDisplayStyle(wb, source.si, r, c), key = JSON.stringify(effective); if (!styles.has(key)) styles.set(key, cellStyleDefaults(clone(effective))); return styles.get(key); };
  return { ...source, ...copiedRows, rows: [...rows], data: source.data.map((row, i) => row.map((d, j) => { const style = styleAt(rows[i], source.c1 + j), copied = cellData(wb.getCell(source.si, rows[i], source.c1 + j), style) ?? { raw: '' }; copied.style = style; return copied; })), values: source.values.map(row => row.map(value => isError(value) ? value : clone(value))),
    sourceSheet: sh, sourceName: sh.name, widths: Array.from({ length: source.c2 - source.c1 + 1 }, (_, j) => wb.colWidth(source.si, source.c1 + j)),
    cond: clone((sh.cond ?? []).filter(r => rangesOf(r).some(g => intersect(g, range)))), validations: clone((sh.validations ?? []).filter(r => intersect(r, range))), merges: clone((sh.merges ?? []).filter(r => intersect(r, range))), ready: true };
}
export function pasteSourceFromText(text, date1904 = false) {
  const rows = parseDelimited(String(text).replace(/\r\n?/g, '\n').replace(/\n$/, ''), '\t');
  let width = 0; for (const row of rows) width = Math.max(width, row.length);
  if (!rows.length || !width || rows.length * width > 2000000) throw Error('붙여넣을 데이터가 없거나 200만 셀을 초과합니다.');
  const data = rows.map(row => Array.from({ length: width }, (_, c) => { const raw = row[c] ?? '', parsed = parseInput(raw, date1904); return { raw, ...(parsed.numFmt ? { style: { numFmt: parsed.numFmt, decimals: parsed.decimals } } : {}) }; }));
  const values = data.map((row, r) => row.map((d, c) => { const cell = makeCellRC(d, r, c, date1904); return cell?.formula ? d.raw : cell?.v ?? null; }));
  return { external: true, ready: true, r1: 0, c1: 0, r2: rows.length - 1, c2: width - 1, rows: rows.map((_, r) => r), data, values, text, cond: [], validations: [], merges: [], widths: [] };
}
export function pasteSpecialRange(src, target, opts = {}) {
  const height = src.data.length, width = src.data[0]?.length ?? 0;
  if (!height || !width) throw Error('복사한 셀이 없습니다.');
  const wholeRows = src.kind === 'rows' && !opts.transpose && !src.external && !['colWidths', 'link'].includes(opts.what);
  const ph = opts.transpose ? width : height, pw = wholeRows ? MAX_COLS : opts.transpose ? height : width;
  const sh = target.r2 - target.r1 + 1, sw = target.c2 - target.c1 + 1;
  const repeat = sh % ph === 0 && (wholeRows || sw % pw === 0);
  const h = repeat ? sh : ph, w = repeat ? sw : pw;
  if (h * (wholeRows ? width : w) > 2000000) throw Error('한 번에 200만 셀까지 붙여넣을 수 있습니다.');
  if (wholeRows) {
    if (target.r1 < 0 || target.r1 + h > MAX_ROWS) throw Error('붙여넣을 범위가 시트의 마지막 행 또는 열을 넘습니다.');
    return { r1: target.r1, c1: 0, r2: target.r1 + h - 1, c2: MAX_COLS - 1, ph, pw, wholeRows: true };
  }
  if (target.r1 < 0 || target.c1 < 0 || target.r1 + h > MAX_ROWS || target.c1 + w > MAX_COLS) throw Error('붙여넣을 범위가 시트의 마지막 행 또는 열을 넘습니다.');
  return { r1: target.r1, c1: target.c1, r2: target.r1 + h - 1, c2: target.c1 + w - 1, ph, pw };
}
function valueData(value) {
  if (value == null) return { raw: '' };
  if (isError(value)) return { raw: value.code };
  if (typeof value === 'string') return { raw: value, inputType: 'text' };
  return { raw: typeof value === 'boolean' ? value ? 'TRUE' : 'FALSE' : String(value), inputType: 'value' };
}
function contentData(d, value, shifted, valuesOnly) {
  if (d?.image) return { raw: '', image: clone(d.image) };
  if (valuesOnly) return valueData(value);
  const data = { raw: shifted, ...(d?.inputType ? { inputType: d.inputType } : {}), ...(hasFormula(d) ? { fx: true } : {}) };
  if (!hasFormula(d)) Object.assign(data, valueData(value));
  if (hasFormula(d) && d.cached !== undefined && shifted === d.raw) data.cached = d.cached;
  if (d?.image) data.image = clone(d.image);
  return data;
}
function numeric(v) { if (v == null || v === '') return 0; if (typeof v === 'number') return v; if (typeof v === 'string' && v.trim() && Number.isFinite(Number(v))) return Number(v); return null; }
function arithmetic(cur, incoming, tv, sv, op, sourceFormula, valuesOnly) {
  const symbol = { add: '+', sub: '-', mul: '*', div: '/' }[op]; if (!symbol) return incoming;
  if (isError(sv)) return { ...incoming, ...valueData(sv) };
  const b = numeric(sv), a = numeric(tv);
  if (b === null) { const result = { ...incoming }; for (const key of ['raw', 'cached', 'fx', 'inputType', 'image']) delete result[key]; return { ...result, ...contentData(cur, tv, cur?.raw ?? '', false) }; }
  if (a === null) return { ...incoming, ...valueData(tv), ...(cur?.image ? { image: clone(cur.image) } : {}) };
  if (hasFormula(cur) || !valuesOnly && sourceFormula) return { ...incoming, raw: `=(${hasFormula(cur) ? cur.raw.slice(1) : a})${symbol}${!valuesOnly && sourceFormula ? '(' + sourceFormula.slice(1) + ')' : b}`, inputType: undefined, cached: undefined, fx: true };
  const result = op === 'add' ? a + b : op === 'sub' ? a - b : op === 'mul' ? a * b : b === 0 ? null : a / b;
  return { ...incoming, ...(result === null ? { raw: '#DIV/0!', inputType: undefined } : Number.isFinite(result) ? valueData(Number(result.toPrecision(15))) : { raw: '#NUM!', inputType: undefined }), fx: undefined, cached: undefined };
}
/** 영역에서 잘라낸 나머지 사각형(영역 밖 규칙을 지우지 않음). */
function subtract(range, cut) {
  const hit = intersect(range, cut); if (!hit) return [range]; const out = [];
  if (range.r1 < hit.r1) out.push({ ...range, r2: hit.r1 - 1 });
  if (range.r2 > hit.r2) out.push({ ...range, r1: hit.r2 + 1 });
  if (range.c1 < hit.c1) out.push({ r1: hit.r1, r2: hit.r2, c1: range.c1, c2: hit.c1 - 1 });
  if (range.c2 > hit.c2) out.push({ r1: hit.r1, r2: hit.r2, c1: hit.c2 + 1, c2: range.c2 });
  return out;
}
function shiftRule(rule, range, oldAnchor, conditional) {
  const result = { ...clone(rule), ...range }; delete result.more; delete result.pivot;
  for (const k of conditional ? ['formula', 'v1', 'v2'] : ['formula1', 'formula2', 'f1', 'f2']) {
    const f = result[k]; if (typeof f !== 'string' || !f || f.startsWith('"')) continue;
    const prefixed = f.startsWith('='); const shifted = shiftFormula(prefixed ? f : '=' + f, range.r1 - oldAnchor.r1, range.c1 - oldAnchor.c1);
    result[k] = prefixed ? shifted : shifted.slice(1);
  }
  return result;
}
function mappedRanges(range, src, area, opts) {
  const out = [], columns = intersect(range, { r1: range.r1, r2: range.r2, c1: src.c1, c2: src.kind === 'rows' && !opts.transpose ? MAX_COLS - 1 : src.c2 }); if (!columns) return out;
  const runs = []; let run;
  for (let i = 0; i < src.rows.length; i++) {
    const r = src.rows[i]; if (r < range.r1 || r > range.r2) { run = null; continue; }
    if (run && run.end === i - 1 && src.rows[i - 1] === r - 1) run.end = i;
    else { run = { start: i, end: i }; runs.push(run); }
  }
  for (let tr = area.r1; tr <= area.r2; tr += area.ph) for (let tc = area.c1; tc <= area.c2; tc += area.pw) for (const run of runs) {
    if (opts.skipBlanks) {
      // 빈 셀은 유효성/조건부서식도 덮지 않는다. 비어 있지 않은 셀만 작은 범위로 매핑한다.
      for (let i = run.start; i <= run.end; i++) {
        const add = (c, data) => {
          if (c < columns.c1 || c > columns.c2 || blank(data)) return;
          const rr = tr + (opts.transpose ? c - src.c1 : i), cc = tc + (opts.transpose ? i : c - src.c1);
          out.push({ range: { r1: rr, r2: rr, c1: cc, c2: cc }, source: { r1: src.rows[i], c1: c } });
        };
        const first = Math.max(columns.c1, src.c1), last = Math.min(columns.c2, src.c1 + src.data[i].length - 1);
        for (let c = first; c <= last; c++) add(c, src.data[i][c - src.c1]);
        if (area.wholeRows) for (const tail of src.tailCells?.[i] ?? []) add(tail.c, tail.data);
      }
    } else {
      const r = opts.transpose ? { r1: tr + columns.c1 - src.c1, r2: tr + columns.c2 - src.c1, c1: tc + run.start, c2: tc + run.end } : { r1: tr + run.start, r2: tr + run.end, c1: tc + columns.c1 - src.c1, c2: tc + columns.c2 - src.c1 };
      out.push({ range: r, source: { r1: src.rows[run.start], c1: columns.c1 } });
    }
  }
  return out;
}
function pasteRules(wb, si, src, area, opts, prop) {
  const conditional = prop === 'cond', rules = src[prop] ?? [], incoming = [];
  for (const rule of rules) for (const range of rangesOf(rule)) for (const m of mappedRanges(range, src, area, opts)) {
    // 식은 원래 규칙의 기준 셀에서 새 범위의 기준 셀까지 이동한다.
    incoming.push(shiftRule(rule, m.range, rule, conditional));
  }
  const current = wb.sheets[si][prop] ?? [], rest = [];
  const cuts = opts.skipBlanks ? mappedRanges({ r1: src.rows[0], r2: src.rows.at(-1), c1: src.c1, c2: area.wholeRows ? MAX_COLS - 1 : src.c2 }, src, area, opts).map(x => x.range) : [area];
  for (const rule of current) {
    if (conditional && opts.what === 'mergeCond') { rest.push(clone(rule)); continue; }
    let ranges = rangesOf(rule); for (const cut of cuts) ranges = ranges.flatMap(r => subtract(r, cut));
    for (const range of ranges) rest.push(shiftRule(rule, range, rule, conditional));
  }
  if (current.length || incoming.length) wb.setSheetProp(si, prop, [...incoming, ...rest]);
}
/** Excel 전치: 상대 축/고정 표시를 바꾸고, 복사 범위 밖 고정 참조는 그대로 둔다. */
export function transposePasteFormula(formula, src, sourceRow, sourceCol, targetRow, targetCol) {
  const copied = { r1: src.rows[0], r2: src.rows.at(-1), c1: src.c1, c2: src.c2 };
  return rewriteRefs(formula, ref => {
    if ((ref.ar1 || ref.ar2 || ref.ac1 || ref.ac2) && !intersect(ref, copied)) return undefined;
    const next = { ...ref, r1: targetRow + ref.c1 - sourceCol, r2: targetRow + ref.c2 - sourceCol, c1: targetCol + ref.r1 - sourceRow, c2: targetCol + ref.r2 - sourceRow,
      ar1: ref.ac1, ar2: ref.ac2, ac1: ref.ar1, ac2: ref.ar2, rows: false, cols: false };
    if (ref.cols) { next.range = true; next.ac1 = true; next.ac2 = true; next.c2 = MAX_COLS - 1; }
    if (ref.rows) { next.range = true; next.ar1 = true; next.ar2 = true; next.r2 = EXCEL_MAX_ROWS - 1; }
    if (ref.range && ref.c2 === MAX_COLS - 1) next.r2 = EXCEL_MAX_ROWS - 1;
    if (ref.range && (ref.r2 === EXCEL_MAX_ROWS - 1 || ref.r2 === MAX_ROWS - 1)) next.c2 = MAX_COLS - 1;
    if (next.r1 < 0 || next.c1 < 0 || next.r2 >= MAX_ROWS || next.c2 >= MAX_COLS) return null;
    return next;
  });
}
/** 구조 삽입 전에 실패할 수 있는 옵션·범위·희소 셀 수를 검사한다. */
export function preflightPasteSpecial(wb, si, src, target, opts = {}) {
  const what = opts.what ?? 'all', area = pasteSpecialRange(src, target, opts);
  if (!PASTE_SPECIAL_TYPES.some(x => x[0] === what) && what !== 'link') throw Error('지원하지 않는 붙여넣기 옵션입니다.');
  if (opts.op && !PASTE_SPECIAL_OPERATIONS.some(x => x[0] === opts.op)) throw Error('지원하지 않는 붙여넣기 연산입니다.');
  if (what === 'colWidths' && src.external) throw Error('텍스트 클립보드에는 열 너비 정보가 없습니다.');
  if (what === 'link' && (src.external || wb.sheets[src.si] !== src.sourceSheet)) throw Error('연결할 원본 시트를 찾을 수 없습니다. 셀을 다시 복사하세요.');
  if (area.wholeRows && !opts.insertRows && wb.sheets[si].merges.some(m => intersect(m, area) && (m.r1 < area.r1 || m.r2 > area.r2))) {
    throw Error('병합된 셀의 일부에는 붙여넣을 수 없습니다. 병합 영역 전체를 선택하세요.');
  }
  return { area, tailColumns: area.wholeRows ? copiedRowTailColumns(wb, si, src, area, opts) : null };
}
export function applyPasteSpecial(wb, si, src, target, opts = {}, plan) {
  const { area, tailColumns } = plan ?? preflightPasteSpecial(wb, si, src, target, opts);
  const what = opts.what ?? 'all', sh = wb.sheets[si];
  if (what === 'colWidths') { for (let c = area.c1; c <= area.c2; c++) wb.setColWidth(si, c, src.widths[(c - area.c1) % src.widths.length]); return area; }
  const srcStyle = (d, cur) => src.external ? clone(cur?.style) : clone(d?.style ?? {});
  const tails = area.wholeRows ? src.data.map((_, i) => new Map((src.tailCells?.[i] ?? []).map(cell => [cell.c, cell]))) : null;
  const pasteCell = (r, c) => {
    const a = (r - area.r1) % area.ph, b = (c - area.c1) % area.pw, i = opts.transpose ? b : a, j = opts.transpose ? a : b;
    const tail = tails?.[i]?.get(c);
    const d = src.data[i][j] ?? tail?.data ?? (area.wholeRows ? { raw: '', style: src.rowMeta?.[i]?.style } : undefined), val = src.values[i][j] ?? tail?.value ?? null; if (opts.skipBlanks && blank(d)) return;
    const targetCell = wb.getCell(si, r, c), cur = cellData(targetCell), currentStyle = wb.styleAt(si, r, c), shifted = hasFormula(d) && !src.external ? opts.transpose ? transposePasteFormula(d.raw, src, src.rows[i], src.c1 + j, r, c) : shiftFormula(d.raw, r - src.rows[i], c - (src.c1 + j)) : d?.raw ?? '';
    const contents = contentData(d, val, shifted, what === 'values' || what === 'valuesNum'); let data;
    if (what === 'validation') return;
    if (what === 'comments') data = { ...(cur ?? { raw: '' }), comment: clone(d?.comment) };
    else if (what === 'formats') data = cellData(targetCell, srcStyle(d, cur)) ?? { raw: '', style: srcStyle(d, cur) };
    else if (what === 'link') data = { ...cur, raw: `=${src.si === si ? '' : quoteSheetName(wb.sheets[src.si].name) + '!'}$${cellName(src.rows[i], src.c1 + j).replace(/(\d+)$/, '$$$1')}`, inputType: undefined, fx: true, cached: undefined, image: undefined };
    else if (ALL_TYPES.has(what)) {
      const style = srcStyle(d, cur); if (what === 'noBorders') for (const k of BORDER_KEYS) style[k] = currentStyle[k] ?? cellStyleDefaults()[k];
      data = { ...clone(d), ...contents, style };
    } else {
      let style = clone(cur?.style);
      if (what === 'valuesNum' || what === 'formulasNum') { style = { ...style }; const ns = cellStyleDefaults(d?.style); for (const key of NUMBER_KEYS) style[key] = key === 'queryFormat' ? undefined : ns[key]; }
      data = { ...cur, ...contents, style, image: contents.image, cached: contents.cached, fx: contents.fx, inputType: contents.inputType };
    }
    if (opts.op && (ALL_TYPES.has(what) || ['values', 'valuesNum', 'formulas', 'formulasNum'].includes(what))) data = arithmetic(cur, data, wb.getValue(si, r, c), val, opts.op, hasFormula(d) ? shifted : '', what === 'values' || what === 'valuesNum');
    // 보호 시트의 서식 변경 권한으로 셀 잠금 정의를 바꿀 수 없다.
    if (sh.protect?.on && data?.style) { data.style.locked = currentStyle.locked !== false; data.style.hideFormula = !!currentStyle.hideFormula; }
    wb.setCellData(si, r, c, data);
  };
  for (let r = area.r1; r <= area.r2; r++) {
    const end = area.wholeRows ? (src.data[0]?.length ?? 0) - 1 : area.c2;
    for (let c = area.c1; c <= end; c++) pasteCell(r, c);
    if (tailColumns) for (const c of tailColumns) {
      const i = (r - area.r1) % area.ph;
      if (tails[i].has(c) || wb.getCell(si, r, c)) pasteCell(r, c);
    }
  }
  if (area.wholeRows && !opts.skipBlanks && (ALL_TYPES.has(what) || what === 'formats')) applyCopiedRowMetadata(wb, si, src, area, what);
  if (!src.external && (ALL_TYPES.has(what) || what === 'formats')) {
    pasteRules(wb, si, src, area, { ...opts, what }, 'cond');
    const current = sh.merges ?? [], kept = current.filter(m => !intersect(m, area)), incoming = [];
    for (const m of src.merges ?? []) { if (m.r1 < src.rows[0] || m.r2 > src.rows.at(-1) || m.c1 < src.c1 || m.c2 > (area.wholeRows ? MAX_COLS - 1 : src.c2)) continue; for (const n of mappedRanges(m, src, area, opts)) if (n.range.r1 !== n.range.r2 || n.range.c1 !== n.range.c2) incoming.push(n.range); }
    if (!opts.skipBlanks && (incoming.length || kept.length !== current.length)) wb.setSheetProp(si, 'merges', [...kept, ...incoming]);
  }
  if (!src.external && (ALL_TYPES.has(what) || what === 'validation')) pasteRules(wb, si, src, area, { ...opts, what }, 'validations');
  return area;
}

/** 구조 변경 뒤에도 복사 시점의 값·서식은 유지하고 수식/원본 좌표만 옮긴다. 원본 clip은 불변. */
export function remapPasteSource(wb, src, { si, axis, index, count }) {
  if (!src || src.external || axis !== 'row' || !count) return src;
  const target = wb.sheets[si];
  if (!target || !Number.isInteger(index) || !Number.isInteger(count)) throw Error('행 삽입 위치와 개수를 확인하세요.');
  let sourceIndex = src.sourceSheet ? wb.sheets.indexOf(src.sourceSheet) : src.si;
  // Undo가 시트 객체를 복원해도 복사 원본의 이름으로 같은 시트에 다시 연결한다.
  if (sourceIndex < 0) sourceIndex = wb.sheetIndexByName(src.sourceName ?? src.sourceSheet?.name ?? '');
  const sourceSheet = wb.sheets[sourceIndex] ?? src.sourceSheet;
  const hostSheet = sourceSheet?.name ?? src.sourceName ?? '';
  const own = sourceIndex === si;
  const move = r => count > 0 ? r >= index ? r + count : r : r >= index - count ? r + count : r >= index ? null : r;
  const shift = raw => adjustFormulaForStructure(raw, { targetSheet: target.name, hostSheet, axis, index, count });
  const fixCell = data => {
    if (!data) return data;
    let next = data;
    if (hasFormula(data)) { const raw = shift(data.raw); if (raw !== data.raw) next = { ...next, raw }; }
    if (data.link) { const link = rewriteWorkbookLink(data.link, shift, name => wb.sheetIndexByName(name) >= 0); if (link !== data.link) next = { ...next, link }; }
    return next;
  };
  const originalRows = src.rows ?? src.data.map((_, i) => src.r1 + i);
  const kept = originalRows.map((r, i) => ({ r: own ? move(r) : r, i })).filter(x => x.r !== null);
  const mapRule = (rule, conditional) => {
    const range = own ? adjustRange(rule, axis, index, count) : rule;
    if (!range) return null;
    const next = { ...rule, r1: range.r1, r2: range.r2, c1: range.c1, c2: range.c2 };
    if (rule.more) next.more = rule.more.map(r => own ? adjustRange(r, axis, index, count) : r).filter(Boolean);
    for (const key of conditional ? ['formula', 'v1', 'v2'] : ['formula1', 'formula2', 'f1', 'f2']) {
      const raw = rule[key]; if (typeof raw !== 'string' || !raw || raw.startsWith('"')) continue;
      const prefixed = raw.startsWith('='), fixed = shift(prefixed ? raw : '=' + raw); next[key] = prefixed ? fixed : fixed.slice(1);
    }
    return next;
  };
  const r1 = own ? move(src.r1) ?? kept[0]?.r ?? index : src.r1;
  return { ...src, si: sourceIndex >= 0 ? sourceIndex : src.si, sourceSheet, sourceName: sourceSheet?.name ?? hostSheet,
    r1, r2: own ? r1 + kept.length - 1 : src.r2, rows: kept.map(x => x.r),
    data: kept.map(({ i }) => src.data[i].map(fixCell)), values: kept.map(({ i }) => src.values[i]),
    rowMeta: src.rowMeta && kept.map(({ i }) => src.rowMeta[i]),
    tailCells: src.tailCells && kept.map(({ i }) => src.tailCells[i].map(cell => ({ ...cell, data: fixCell(cell.data) }))),
    cond: (src.cond ?? []).map(rule => mapRule(rule, true)).filter(Boolean),
    validations: (src.validations ?? []).map(rule => mapRule(rule, false)).filter(Boolean),
    merges: (src.merges ?? []).map(rule => own ? adjustRange(rule, axis, index, count) : rule).filter(Boolean) };
}
