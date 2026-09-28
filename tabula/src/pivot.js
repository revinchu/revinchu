// 피벗 테이블 계산 (DOM 없음)
// 정의(def): { source, range | table, rows: [필드 이름], cols: [필드 이름], values: [{ field, agg, name, showAs, numFmt }],
//              pages: [필드 이름], filters: { 필드 이름: [보이는 항목 글자] }, layout: 'compact' | 'outline' | 'tabular',
//              subtotals, grandRows, grandCols, top, left, area }
// 옛 정의 { rowField, colField, valueField, agg, fieldNames } 도 그대로 읽음
import { formatGeneral } from './format.js';
import { findTable, dataTop, dataBottom, columnNames } from './tables.js';

export const AGGREGATES = [
  { id: 'sum', label: '합계' },
  { id: 'count', label: '개수' },
  { id: 'average', label: '평균' },
  { id: 'max', label: '최대' },
  { id: 'min', label: '최소' },
  { id: 'product', label: '곱' },
  { id: 'countNums', label: '숫자 개수' },
  { id: 'stdDev', label: '표준 편차' },
  { id: 'stdDevp', label: '표준 편차(전체)' },
  { id: 'var', label: '분산' },
  { id: 'varp', label: '분산(전체)' },
];
export const SHOW_AS = [
  { id: 'normal', label: '계산 없음' },
  { id: 'percentOfTotal', label: '총합계 비율' },
  { id: 'percentOfCol', label: '열 합계 비율' },
  { id: 'percentOfRow', label: '행 합계 비율' },
];
export const LAYOUTS = [
  { id: 'compact', label: '압축 형식으로 표시' },
  { id: 'outline', label: '개요 형식으로 표시' },
  { id: 'tabular', label: '테이블 형식으로 표시' },
];

export const EMPTY = '(비어 있음)';
export const TOTAL = '총합계';
export const keyOf = (v) => (v === null || v === undefined || v === '' ? EMPTY : typeof v === 'object' ? String(v.code) : v);

export function sortKeys(keys) {
  return keys.sort((a, b) => {
    if (a === EMPTY) return 1;
    if (b === EMPTY) return -1;
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'number') return -1;
    if (typeof b === 'number') return 1;
    return String(a).localeCompare(String(b), 'ko');
  });
}

/** 슬라이서·필터에서 쓰는 항목 글자 */
export const itemText = (v) => (v === null || v === undefined || v === '' ? EMPTY : typeof v === 'number' ? formatGeneral(v) : typeof v === 'object' ? v.code : String(v));

/** 머리글 이름 (빈 칸은 열N) */
export const headerNames = (rows) => (rows[0] ?? []).map((h, i) => (h === null || h === '' ? `열${i + 1}` : String(h)));

export const aggLabel = (agg) => AGGREGATES.find((a) => a.id === agg)?.label ?? '합계';
export const valueName = (v) => v.name || `${aggLabel(v.agg)} : ${v.field}`;

// ───────────── 누적 ─────────────
const newAcc = () => ({ count: 0, nums: 0, sum: 0, sq: 0, min: Infinity, max: -Infinity, prod: 1 });
function add(acc, v) {
  if (v !== null && v !== '' && v !== undefined) acc.count++;
  if (typeof v === 'number' && Number.isFinite(v)) {
    acc.nums++;
    acc.sum += v;
    acc.sq += v * v;
    acc.prod *= v;
    if (v < acc.min) acc.min = v;
    if (v > acc.max) acc.max = v;
  }
}
function result(acc, agg) {
  if (!acc) return null;
  const n = acc.nums;
  const varOf = (sample) => {
    const d = n - (sample ? 1 : 0);
    if (d <= 0) return null;
    return Math.max(0, (acc.sq - (acc.sum * acc.sum) / n) / d);
  };
  switch (agg) {
    case 'count': return acc.count;
    case 'countNums': return n;
    case 'average': return n ? acc.sum / n : null;
    case 'max': return n ? acc.max : null;
    case 'min': return n ? acc.min : null;
    case 'product': return n ? acc.prod : null;
    case 'var': return varOf(true);
    case 'varp': return varOf(false);
    case 'stdDev': { const x = varOf(true); return x === null ? null : Math.sqrt(x); }
    case 'stdDevp': { const x = varOf(false); return x === null ? null : Math.sqrt(x); }
    default: return acc.sum;
  }
}

// ───────────── 정의 정리 ─────────────
/** 옛 정의까지 포함해 필드 이름 기반 정의로 */
export function normalizeDef(def, header) {
  const name = (i) => (i === null || i === undefined || i < 0 ? null : header[i] ?? null);
  const byName = (n) => header.find((h) => h.toLowerCase() === String(n).toLowerCase()) ?? null;
  let rows = def.rows;
  let cols = def.cols;
  let values = def.values;
  if (!rows) {
    // 옛 형식: 필드 번호 + 이름
    const r = def.fieldNames?.row ? byName(def.fieldNames.row) ?? name(def.rowField) : name(def.rowField);
    const c = def.colField === null || def.colField === undefined ? null : def.fieldNames?.col ? byName(def.fieldNames.col) ?? name(def.colField) : name(def.colField);
    const v = def.valueField === null || def.valueField === undefined ? null : def.fieldNames?.val ? byName(def.fieldNames.val) ?? name(def.valueField) : name(def.valueField);
    rows = r ? [r] : [];
    cols = c ? [c] : [];
    values = [{ field: v ?? r, agg: v === null ? 'count' : def.agg ?? 'sum' }];
  }
  const ok = (n) => byName(n) !== null;
  return {
    rows: (rows ?? []).filter(ok).map(byName),
    cols: (cols ?? []).filter(ok).map(byName),
    pages: (def.pages ?? []).filter(ok).map(byName),
    values: (values ?? []).filter((v) => ok(v.field)).map((v) => ({ ...v, field: byName(v.field), agg: v.agg ?? 'sum' })),
    layout: def.layout ?? 'compact',
    subtotals: def.subtotals !== false,
    grandRows: def.grandRows !== false,
    grandCols: def.grandCols !== false,
    filters: def.filters ?? {},
  };
}

/**
 * 피벗 원본 → { rows (머리글 포함 값), sheet, ref: {r1,c1,r2,c2} | null, table: 표 이름 | null }
 * 표 이름이면 지금의 표 범위(누적된 데이터 포함), 아니면 고정 범위
 */
export function pivotSourceData(wb, def) {
  const read = (si, rg) => {
    const rows = [];
    for (let r = rg.r1; r <= rg.r2; r++) {
      const row = [];
      for (let c = rg.c1; c <= rg.c2; c++) row.push(wb.getValue(si, r, c));
      rows.push(row);
    }
    return rows;
  };
  if (def.table) {
    const f = findTable(wb, def.table);
    if (!f) return null;
    const t = f.t;
    const ref = { r1: t.header ? t.r1 : dataTop(t), c1: t.c1, r2: dataBottom(t), c2: t.c2 };
    const rows = read(f.si, ref);
    if (!t.header) rows.unshift(columnNames(wb, f.si, t));
    return { rows, si: f.si, ref, table: t.name };
  }
  const si = wb.sheetIndexByName(def.source);
  if (si < 0 || !def.range) return null;
  return { rows: read(si, def.range), si, ref: def.range, table: null };
}

/** 필터(보고서 필터 · 항목 선택 · 슬라이서)를 적용한 행 → { def: 정리된 정의, rows, header } */
export function resolvePivot(rows, def) {
  const header = headerNames(rows);
  const d = normalizeDef(def, header);
  const filters = Object.entries(d.filters).map(([name, allowed]) => [header.findIndex((h) => h.toLowerCase() === name.toLowerCase()), new Set(allowed)]).filter(([i]) => i >= 0);
  const data = rows.slice(1).filter((r) => !r.every((v) => v === null || v === ''));
  const out = filters.length ? data.filter((r) => filters.every(([i, set]) => set.has(itemText(r[i])))) : data;
  return { def: d, rows: [rows[0], ...out], header };
}

// ───────────── 계산 ─────────────
const kk = (k) => `${typeof k}:${k}`;

/** 필드 값 트리 (정렬된 자식) */
function buildTree(data, idxs) {
  const root = { key: null, path: '', depth: -1, children: [], map: new Map() };
  for (const r of data) {
    let node = root;
    let path = '';
    idxs.forEach((f, d) => {
      const k = keyOf(r[f]);
      path += `\u0001${kk(k)}`;
      let ch = node.map.get(kk(k));
      if (!ch) { ch = { key: k, path, depth: d, children: [], map: new Map() }; node.map.set(kk(k), ch); node.children.push(ch); }
      node = ch;
    });
  }
  const sortRec = (n) => {
    const keys = sortKeys(n.children.map((c) => c.key));
    const byKey = new Map(n.children.map((c) => [kk(c.key), c]));
    n.children = keys.map((k) => byKey.get(kk(k)));
    n.children.forEach(sortRec);
  };
  sortRec(root);
  return root;
}

/**
 * 피벗 계산 → { grid: [[{raw, style}]], meta }
 * rows: 머리글 포함 (필터 적용 후), d: normalizeDef 결과
 */
export function computePivot(rows, d) {
  const header = headerNames(rows);
  const idx = (n) => header.findIndex((h) => h.toLowerCase() === String(n).toLowerCase());
  const data = rows.slice(1).filter((r) => !r.every((v) => v === null || v === ''));
  const rowIdx = d.rows.map(idx);
  const colIdx = d.cols.map(idx);
  const values = d.values.length ? d.values : [{ field: d.rows[0] ?? header[0], agg: 'count' }];
  const valIdx = values.map((v) => idx(v.field));
  const V = values.length;
  const rowTree = buildTree(data, rowIdx);
  const colTree = buildTree(data, colIdx);

  // 누적: 행 경로 접두사 × 열 경로 접두사 × 값
  const accs = new Map();
  for (const r of data) {
    const rp = [''];
    let p = '';
    for (const f of rowIdx) { p += `\u0001${kk(keyOf(r[f]))}`; rp.push(p); }
    const cp = [''];
    p = '';
    for (const f of colIdx) { p += `\u0001${kk(keyOf(r[f]))}`; cp.push(p); }
    for (const a of rp) {
      for (const b of cp) {
        const key = `${a}\u0002${b}`;
        let list = accs.get(key);
        if (!list) { list = values.map(newAcc); accs.set(key, list); }
        valIdx.forEach((f, i) => add(list[i], f < 0 ? 1 : r[f]));
      }
    }
  }
  const raw = (rp, cp, vi) => result(accs.get(`${rp}\u0002${cp}`)?.[vi], values[vi].agg);
  const cellValue = (rp, cp, vi) => {
    const v = raw(rp, cp, vi);
    if (v === null) return null;
    const as = values[vi].showAs ?? 'normal';
    if (as === 'normal') return v;
    const base = as === 'percentOfTotal' ? raw('', '', vi) : as === 'percentOfRow' ? raw(rp, '', vi) : raw('', cp, vi);
    return base ? v / base : null;
  };

  // 열 머리글 잎 목록: { cp, vi, kind: 'item' | 'sub' | 'grand', labels: [수준별 글자] }
  const Lc = colIdx.length;
  const multiV = V > 1;
  const colLeaves = [];
  const walkCols = (node, labels) => {
    if (node.depth === Lc - 1 || !node.children.length) {
      for (let vi = 0; vi < V; vi++) colLeaves.push({ cp: node.path, vi, kind: 'item', labels: [...labels, ...(multiV ? [valueName(values[vi])] : [])], node });
      return;
    }
    node.children.forEach((ch) => walkCols(ch, [...labels, itemText(ch.key)]));
    if (d.subtotals && node.depth >= 0 && node.depth < Lc - 1) {
      for (let vi = 0; vi < V; vi++) colLeaves.push({ cp: node.path, vi, kind: 'sub', labels: [...labels.slice(0, -1), `${itemText(node.key)} 요약`], node });
    }
  };
  if (Lc) colTree.children.forEach((ch) => walkCols(ch, [itemText(ch.key)]));
  else for (let vi = 0; vi < V; vi++) colLeaves.push({ cp: '', vi, kind: 'item', labels: multiV ? [valueName(values[vi])] : [], node: colTree });
  if (Lc && d.grandCols) {
    for (let vi = 0; vi < V; vi++) colLeaves.push({ cp: '', vi, kind: 'grand', labels: [multiV ? `전체 ${valueName(values[vi])}` : TOTAL], node: colTree });
  }

  const HEAD = { bold: true, fill: '#d9e1f2', bb: true };
  const SUB = { bold: true };
  const GRAND = { bold: true, fill: '#d9e1f2', bt: true };
  const numStyle = (vi) => {
    const v = values[vi];
    if (v.numFmt) return typeof v.numFmt === 'object' ? v.numFmt : { numFmt: v.numFmt };
    if ((v.showAs ?? 'normal') !== 'normal') return { numFmt: 'percent', decimals: 2 };
    return v.agg === 'average' || v.agg?.startsWith('std') || v.agg?.startsWith('var') ? { numFmt: 'number', decimals: 2 } : { numFmt: 'comma' };
  };
  const val = (n, vi, extra = {}) => (n === null ? { raw: '', style: { ...extra } } : { raw: String(Number(n.toPrecision(15))), style: { ...numStyle(vi), ...extra } });
  const text = (s, style = {}) => ({ raw: typeof s === 'number' ? formatGeneral(s) : String(s).startsWith('=') ? `'${s}` : String(s ?? ''), style });

  const layout = d.layout;
  const Lr = rowIdx.length;
  const labelCols = layout === 'compact' ? 1 : Math.max(1, Lr);
  const grid = [];
  const rowItems = [];
  const colItems = [];

  // 보고서 필터 (맨 위)
  const pageRows = [];
  for (const p of d.pages) {
    const allowed = d.filters[p] ?? d.filters[Object.keys(d.filters).find((k) => k.toLowerCase() === p.toLowerCase())];
    pageRows.push([text(p, { bold: true }), text(!allowed ? '(모두)' : allowed.length === 1 ? allowed[0] : '(다중 항목)')]);
  }
  if (pageRows.length) { grid.push(...pageRows); grid.push([]); }

  // 열 머리글
  const colLevels = Lc + (multiV ? 1 : 0);
  const hasColHead = colLevels > 0;
  const valueCaption = V === 1 ? valueName(values[0]) : '';
  const rowHeaderCells = () => {
    if (layout === 'compact') return [text(Lr ? '행 레이블' : '', HEAD)];
    return Array.from({ length: labelCols }, (_, i) => text(d.rows[i] ?? '', HEAD));
  };
  if (hasColHead) {
    // 열 필드가 있으면 맨 위에 '값 이름 | 열 레이블' 행 (값 필드만 여러 개면 생략)
    if (Lc) grid.push([text(valueCaption, HEAD), ...Array(labelCols - 1).fill(null).map(() => text('', HEAD)), text('열 레이블', HEAD), ...colLeaves.slice(1).map(() => text('', HEAD))]);
    for (let lvl = 0; lvl < colLevels; lvl++) {
      const row = lvl === colLevels - 1 ? rowHeaderCells() : Array.from({ length: labelCols }, () => text('', HEAD));
      let prev = null;
      colLeaves.forEach((leaf) => {
        const lab = leaf.labels[lvl];
        // 같은 그룹의 반복 글자는 첫 칸에만 (엑셀과 같음)
        const groupKey = leaf.labels.slice(0, lvl + 1).join('\u0001');
        const show = lab !== undefined && (lvl === colLevels - 1 || groupKey !== prev);
        prev = groupKey;
        row.push(text(show ? lab : '', HEAD));
      });
      grid.push(row);
    }
  } else {
    grid.push([...rowHeaderCells(), text(valueCaption, HEAD)]);
  }
  const firstDataRowRel = grid.length - pageRows.length - (pageRows.length ? 1 : 0);

  // 본문
  const valueCells = (rp, style) => colLeaves.map((leaf) => val(cellValue(rp, leaf.cp, leaf.vi), leaf.vi, leaf.kind === 'item' ? style : { ...style, bold: true }));
  const labelsRow = (node, labelText, style) => {
    const cells = Array.from({ length: labelCols }, () => text('', style));
    const col = layout === 'compact' ? 0 : node.depth;
    cells[col] = text(labelText, layout === 'compact' && node.depth > 0 ? { ...style, indent: node.depth } : style);
    return cells;
  };
  const walkRows = (node, lastPath) => {
    const isLeaf = node.depth === Lr - 1;
    if (layout === 'tabular') {
      if (isLeaf) {
        // 조상 글자는 그룹의 첫 행에만
        const cells = Array.from({ length: labelCols }, () => text(''));
        const chain = [];
        for (let n = node; n && n.depth >= 0; n = n.parent) chain.unshift(n);
        chain.forEach((n, dd) => { if (!lastPath.shown.has(n.path)) { cells[dd] = text(itemText(n.key)); lastPath.shown.add(n.path); } });
        grid.push([...cells, ...valueCells(node.path, {})]);
        rowItems.push({ kind: 'item', node, chain });
        return;
      }
      node.children.forEach((ch) => walkRows(ch, lastPath));
      if (d.subtotals) {
        const cells = labelsRow(node, `${itemText(node.key)} 요약`, SUB);
        grid.push([...cells, ...valueCells(node.path, SUB)]);
        rowItems.push({ kind: 'sub', node });
      }
      return;
    }
    const group = !isLeaf;
    const style = group ? SUB : {};
    const cells = labelsRow(node, itemText(node.key), style);
    grid.push([...cells, ...(group && !d.subtotals ? colLeaves.map(() => text('', style)) : valueCells(node.path, style))]);
    rowItems.push({ kind: 'item', node });
    node.children.forEach((ch) => walkRows(ch, lastPath));
  };
  const setParents = (n) => n.children.forEach((c) => { c.parent = n.depth >= 0 ? n : null; setParents(c); });
  setParents(rowTree);
  if (Lr) rowTree.children.forEach((ch) => walkRows(ch, { shown: new Set() }));
  else { grid.push([text(TOTAL, {}), ...valueCells('', {})]); }
  if (Lr && d.grandRows) {
    const cells = Array.from({ length: labelCols }, () => text('', GRAND));
    cells[0] = text(TOTAL, GRAND);
    grid.push([...cells, ...valueCells('', GRAND)]);
    rowItems.push({ kind: 'grand' });
  }
  for (const leaf of colLeaves) colItems.push(leaf);

  const width = labelCols + colLeaves.length;
  return {
    grid,
    meta: {
      header, rowIdx, colIdx, valIdx, values, labelCols, pageRows: pageRows.length ? pageRows.length + 1 : 0,
      headerRows: firstDataRowRel, colLeaves, rowItems, colItems, rowTree, colTree, width,
      bodyRows: grid.length - (pageRows.length ? pageRows.length + 1 : 0),
    },
  };
}

/** 옛 API: 셀 데이터 2차원 배열 */
export function buildPivot(rows, def) {
  const header = headerNames(rows);
  return computePivot(rows, normalizeDef(def, header)).grid;
}

/** GETPIVOTDATA: 값 필드 이름과 (필드, 항목) 쌍으로 값 찾기 */
export function pivotLookup(rows, def, dataField, pairs) {
  const header = headerNames(rows);
  const d = normalizeDef(def, header);
  const lower = String(dataField).toLowerCase();
  const vi = d.values.findIndex((v) => valueName(v).toLowerCase() === lower || v.field.toLowerCase() === lower || (v.name ?? '').toLowerCase() === lower);
  if (vi < 0) return null;
  const want = pairs.map(([f, item]) => [header.findIndex((h) => h.toLowerCase() === String(f).toLowerCase()), itemText(item)]);
  if (want.some(([i]) => i < 0)) return null;
  // 행·열 필드가 아닌 필드로 묻는 것은 엑셀에서 #REF!
  const fields = new Set([...d.rows, ...d.cols].map((n) => header.indexOf(n)));
  if (want.some(([i]) => !fields.has(i))) return null;
  const { rows: filtered } = resolvePivot(rows, def);
  const data = filtered.slice(1).filter((r) => want.every(([i, t]) => itemText(r[i]) === t));
  if (!data.length) return null;
  const acc = newAcc();
  const f = header.indexOf(d.values[vi].field);
  for (const r of data) add(acc, r[f]);
  return result(acc, d.values[vi].agg);
}
