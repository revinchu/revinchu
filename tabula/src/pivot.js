// 피벗 테이블 계산 (DOM 없음)
import { formatGeneral } from './format.js';

export const AGGREGATES = [
  { id: 'sum', label: '합계' },
  { id: 'count', label: '개수' },
  { id: 'average', label: '평균' },
  { id: 'max', label: '최대' },
  { id: 'min', label: '최소' },
];

const EMPTY = '(비어 있음)';
const keyOf = (v) => (v === null || v === undefined || v === '' ? EMPTY : typeof v === 'object' ? String(v.code) : v);

function sortKeys(keys) {
  return keys.sort((a, b) => {
    if (a === EMPTY) return 1;
    if (b === EMPTY) return -1;
    if (typeof a === 'number' && typeof b === 'number') return a - b;
    if (typeof a === 'number') return -1;
    if (typeof b === 'number') return 1;
    return String(a).localeCompare(String(b), 'ko');
  });
}

const newAcc = () => ({ sum: 0, count: 0, nums: 0, min: Infinity, max: -Infinity });
function add(acc, v, countOnly) {
  if (countOnly) { if (v !== null && v !== '' && v !== undefined) acc.count++; return; }
  if (v !== null && v !== '' && v !== undefined) acc.count++;
  if (typeof v === 'number' && Number.isFinite(v)) {
    acc.sum += v;
    acc.nums++;
    acc.min = Math.min(acc.min, v);
    acc.max = Math.max(acc.max, v);
  }
}
function result(acc, agg) {
  if (!acc) return null;
  switch (agg) {
    case 'count': return acc.count;
    case 'average': return acc.nums ? acc.sum / acc.nums : null;
    case 'max': return acc.nums ? acc.max : null;
    case 'min': return acc.nums ? acc.min : null;
    default: return acc.sum;
  }
}

/**
 * rows: 머리글 행을 포함한 2차원 값
 * def: { rowField, colField (선택), valueField, agg }  — 필드는 열 번호(0부터)
 * 반환: 셀 데이터 2차원 배열 [{raw, style}|null]
 */
export function buildPivot(rows, def) {
  const [header, ...data] = rows;
  const name = (i) => (i === null || i === undefined ? '' : String(header[i] ?? `열${i + 1}`));
  const agg = def.agg ?? 'sum';
  const countOnly = agg === 'count';
  const rowKeys = new Map();
  const colKeys = new Map();
  const cells = new Map();
  const rowTot = new Map();
  const colTot = new Map();
  const grand = newAcc();
  for (const r of data) {
    if (r.every((v) => v === null || v === '')) continue;
    const rk = keyOf(r[def.rowField]);
    const ck = def.colField === null || def.colField === undefined ? null : keyOf(r[def.colField]);
    const v = def.valueField === null || def.valueField === undefined ? 1 : r[def.valueField];
    rowKeys.set(rk, true);
    if (ck !== null) colKeys.set(ck, true);
    const k = `${typeof rk}:${rk}\u0001${typeof ck}:${ck}`;
    if (!cells.has(k)) cells.set(k, newAcc());
    add(cells.get(k), v, countOnly);
    const rkk = `${typeof rk}:${rk}`;
    if (!rowTot.has(rkk)) rowTot.set(rkk, newAcc());
    add(rowTot.get(rkk), v, countOnly);
    if (ck !== null) {
      const ckk = `${typeof ck}:${ck}`;
      if (!colTot.has(ckk)) colTot.set(ckk, newAcc());
      add(colTot.get(ckk), v, countOnly);
    }
    add(grand, v, countOnly);
  }
  const rks = sortKeys([...rowKeys.keys()]);
  const cks = sortKeys([...colKeys.keys()]);
  const aggLabel = AGGREGATES.find((a) => a.id === agg)?.label ?? '합계';
  const title = `${aggLabel} : ${def.valueField === null || def.valueField === undefined ? name(def.rowField) : name(def.valueField)}`;
  const HEAD = { bold: true, fill: '#d9e1f2', bb: true };
  const TOTAL = { bold: true, fill: '#d9e1f2', bt: true };
  const numStyle = agg === 'average' ? { numFmt: 'number', decimals: 2 } : { numFmt: 'comma' };
  const val = (n, extra = {}) => (n === null ? { raw: '', style: { ...extra } } : { raw: String(Number(n.toPrecision(15))), style: { ...numStyle, ...extra } });
  const text = (s, style) => ({ raw: typeof s === 'number' ? formatGeneral(s) : String(s).startsWith('=') ? `'${s}` : String(s), style });
  const out = [];
  if (cks.length) {
    out.push([text(title, HEAD), text('열 레이블', HEAD), ...cks.slice(1).map(() => ({ raw: '', style: HEAD })), { raw: '', style: HEAD }]);
    out.push([text('행 레이블', HEAD), ...cks.map((c) => text(c, HEAD)), text('총합계', HEAD)]);
    for (const rk of rks) {
      out.push([
        text(rk, {}),
        ...cks.map((ck) => val(result(cells.get(`${typeof rk}:${rk}\u0001${typeof ck}:${ck}`), agg))),
        val(result(rowTot.get(`${typeof rk}:${rk}`), agg), { bold: true }),
      ]);
    }
    out.push([text('총합계', TOTAL), ...cks.map((ck) => val(result(colTot.get(`${typeof ck}:${ck}`), agg), TOTAL)), val(result(grand, agg), TOTAL)]);
  } else {
    out.push([text('행 레이블', HEAD), text(title, HEAD)]);
    for (const rk of rks) out.push([text(rk, {}), val(result(rowTot.get(`${typeof rk}:${rk}`), agg))]);
    out.push([text('총합계', TOTAL), val(result(grand, agg), TOTAL)]);
  }
  return out;
}
