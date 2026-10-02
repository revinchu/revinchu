import { serialOf } from './format.js';

// 값 스냅숏에만 적용하는 로컬 변환. 통합 문서·수식·원본 표를 수정하지 않습니다.
export const RANGE_QUERY_LIMITS = { maxCells: 500000, maxRows: 100000, maxColumns: 1000 };
export const RANGE_QUERY_TYPES = [['keep', '원본 유지'], ['text', '텍스트'], ['number', '숫자'], ['integer', '정수'], ['date', '날짜'], ['boolean', '논리값']];
export const RANGE_QUERY_FILTERS = [['contains', '포함'], ['notContains', '포함하지 않음'], ['eq', '같음'], ['ne', '같지 않음'], ['starts', '시작 문자'], ['ends', '끝 문자'], ['gt', '보다 큼'], ['ge', '이상'], ['lt', '보다 작음'], ['le', '이하'], ['blank', '비어 있음'], ['notBlank', '비어 있지 않음']];
const blank = v => v === null || v === undefined || v === '';
const typeKeys = new Set(RANGE_QUERY_TYPES.map(x => x[0]));
const filterKeys = new Set(RANGE_QUERY_FILTERS.map(x => x[0]));
const bad = Symbol('conversion-error');
const textOf = v => blank(v) ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : typeof v === 'object' ? String(v.code ?? '#VALUE!') : String(v);
const scalar = v => blank(v) ? null : typeof v === 'object' ? { code: String(v.code ?? '#VALUE!') } : typeof v === 'number' && !Number.isFinite(v) ? { code: '#NUM!' } : v;

export function readRangeQuerySource(wb, si, range, options = {}) {
  const sh = wb.sheets[si];
  if (!sh || !range) throw new Error('가져올 워크시트와 범위를 선택하세요.');
  const limits = { ...RANGE_QUERY_LIMITS, ...options }, table = options.table;
  const area = { r1: table?.r1 ?? range.r1, c1: table?.c1 ?? range.c1, r2: table?.r2 ?? range.r2, c2: table?.c2 ?? range.c2 };
  if (table?.totals) area.r2--;
  if (Object.values(area).some(n => !Number.isInteger(n) || n < 0) || area.r1 > area.r2 || area.c1 > area.c2 || area.r2 >= 1048576 || area.c2 >= 16384) throw new Error('가져올 범위가 올바르지 않습니다.');
  let clipped = false;
  if (area.r2 === 1048575 || area.c2 === 16383) {
    const used = wb.usedRange(si);
    if (area.r2 === 1048575) { area.r2 = used.rows - 1; clipped = true; }
    if (area.c2 === 16383) { area.c2 = used.cols - 1; clipped = true; }
  }
  if (area.r2 < area.r1 || area.c2 < area.c1) throw new Error('선택한 범위에 가져올 데이터가 없습니다.');
  const height = area.r2 - area.r1 + 1, width = area.c2 - area.c1 + 1;
  if (height > limits.maxRows || width > limits.maxColumns || height * width > limits.maxCells) throw new Error(`한 번에 ${limits.maxRows.toLocaleString('ko-KR')}행, ${limits.maxColumns.toLocaleString('ko-KR')}열, ${limits.maxCells.toLocaleString('ko-KR')}셀까지 가져올 수 있습니다. 범위를 줄여 주세요.`);
  const matrix = wb.rangeRead(si, area.r1, area.c1, area.r2, area.c2).map(row => Array.from({ length: width }, (_, c) => scalar(row[c])));
  return { matrix, width, range: area, header: table ? table.header !== false : options.header !== false, label: table?.name ? `${sh.name} · ${table.name}` : sh.name, date1904: !!wb.date1904, clipped };
}

export function rangeQueryHeaders(source, header = source.header !== false) {
  const used = new Set();
  return Array.from({ length: source.width }, (_, c) => {
    const base = (header ? textOf(source.matrix[0]?.[c]).trim() : '') || `열${c + 1}`;
    let name = base, n = 2;
    while (used.has(name.toLocaleLowerCase('ko-KR'))) name = `${base}_${n++}`;
    used.add(name.toLocaleLowerCase('ko-KR')); return name;
  });
}

function numberValue(v) {
  if (typeof v === 'number') return Number.isFinite(v) ? v : bad;
  if (typeof v !== 'string') return bad;
  const s = v.trim();
  if (!/^[+-]?(?:(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?%?$/.test(s)) return bad;
  const percent = s.endsWith('%'), n = Number(s.replace(/,/g, '').replace(/%$/, '')) / (percent ? 100 : 1);
  if (!Number.isFinite(n) || (Number.isInteger(n) && !Number.isSafeInteger(n))) return bad;
  return n;
}

function convert(v, type, date1904) {
  if (blank(v)) return null;
  if (type === 'keep') return v;
  if (type === 'text') return textOf(v);
  if (type === 'number' || type === 'integer') { const n = numberValue(v); return n === bad || (type === 'integer' && !Number.isInteger(n)) ? bad : n; }
  if (type === 'boolean') {
    if (typeof v === 'boolean') return v;
    if (v === 1 || v === 0) return !!v;
    if (typeof v === 'string' && /^(true|false|참|거짓|1|0)$/i.test(v.trim())) return /^(true|참|1)$/i.test(v.trim());
    return bad;
  }
  if (type === 'date') {
    if (typeof v === 'number') return Number.isFinite(v) && v >= 0 && v < (date1904 ? 2957004 : 2958466) ? Math.floor(v) : bad;
    if (typeof v !== 'string') return bad;
    const m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(v.trim());
    if (!m) return bad;
    const y = +m[1], month = +m[2], d = +m[3], dt = new Date(Date.UTC(y, month - 1, d));
    if (y < (date1904 ? 1904 : 1900) || y > 9999 || dt.getUTCFullYear() !== y || dt.getUTCMonth() !== month - 1 || dt.getUTCDate() !== d) return bad;
    return serialOf(y, month, d, date1904);
  }
  return bad;
}

function compare(a, b) {
  if (blank(a)) return blank(b) ? 0 : 1;
  if (blank(b)) return -1;
  if (typeof a === 'number' && typeof b === 'number') return a - b;
  if (typeof a === 'boolean' && typeof b === 'boolean') return Number(a) - Number(b);
  // 혼합형도 일관된 전순서: 숫자, 논리값, 텍스트, 오류, 빈 값.
  const rank = v => typeof v === 'number' ? 0 : typeof v === 'boolean' ? 1 : typeof v === 'string' ? 2 : 3;
  if (rank(a) !== rank(b)) return rank(a) - rank(b);
  return textOf(a).localeCompare(textOf(b), 'ko-KR');
}

function filterRow(v, f) {
  if (f.op === 'blank') return blank(v);
  if (f.op === 'notBlank') return !blank(v);
  if (blank(v)) return false;
  const a = textOf(v).toLocaleLowerCase('ko-KR'), b = textOf(f.value).toLocaleLowerCase('ko-KR');
  if (f.op === 'contains') return a.includes(b);
  if (f.op === 'notContains') return !a.includes(b);
  if (f.op === 'starts') return a.startsWith(b);
  if (f.op === 'ends') return a.endsWith(b);
  if (f.op === 'eq' || f.op === 'ne') { const eq = typeof v === 'number' ? numberValue(f.value) === v : a === b; return f.op === 'eq' ? eq : !eq; }
  const n = numberValue(v), limit = numberValue(f.value);
  if (n === bad || limit === bad) return false;
  return f.op === 'gt' ? n > limit : f.op === 'ge' ? n >= limit : f.op === 'lt' ? n < limit : n <= limit;
}

export function transformRangeQuery(source, spec = {}) {
  const header = spec.header ?? source.header !== false, allHeaders = rangeQueryHeaders(source, header);
  const columns = spec.columns ?? Array.from({ length: source.width }, (_, i) => i);
  const validCol = c => Number.isInteger(c) && c >= 0 && c < source.width;
  if (!columns.length || columns.some(c => !validCol(c)) || new Set(columns).size !== columns.length) throw new Error('가져올 열을 하나 이상 선택하세요.');
  const types = Array.from({ length: source.width }, (_, c) => spec.types?.[c] ?? 'keep');
  if (types.some(t => !typeKeys.has(t))) throw new Error('지원하지 않는 열 유형입니다.');
  let filters = spec.filters ?? []; const sorts = spec.sorts ?? [];
  if (filters.some(f => !validCol(f.column) || !filterKeys.has(f.op))) throw new Error('필터의 열과 조건을 확인하세요.');
  if (sorts.some(s => !validCol(s.column) || !['asc', 'desc'].includes(s.direction))) throw new Error('정렬의 열과 방향을 확인하세요.');
  filters = filters.map(f => {
    if (types[f.column] === 'date' && ['eq', 'ne', 'gt', 'ge', 'lt', 'le'].includes(f.op)) {
      const v = convert(f.value, 'date', source.date1904);
      if (v === bad || blank(v)) throw new Error('날짜 필터에는 YYYY-MM-DD 형식의 날짜를 입력하세요.');
      return { ...f, value: v };
    }
    if (types[f.column] === 'boolean' && ['eq', 'ne'].includes(f.op)) {
      const v = convert(f.value, 'boolean', source.date1904);
      if (v === bad || blank(v)) throw new Error('논리값 필터에는 TRUE 또는 FALSE를 입력하세요.');
      return { ...f, value: v };
    }
    return f;
  });
  if (filters.some(f => ['gt', 'ge', 'lt', 'le'].includes(f.op) && numberValue(f.value) === bad)) throw new Error('크기 비교 필터에는 올바른 숫자를 입력하세요.');
  const needed = new Set(columns); for (const f of filters) needed.add(f.column); for (const s of sorts) needed.add(s.column);
  const rows = [], errors = [], seen = new Set(), start = header ? 1 : 0;
  const stats = { inputRows: Math.max(0, source.matrix.length - start), outputRows: 0, filteredRows: 0, blankRows: 0, duplicateRows: 0, errorCount: 0 };
  for (let r = start; r < source.matrix.length; r++) {
    const original = source.matrix[r], values = [];
    for (const c of needed) {
      const v = convert(original[c] ?? null, types[c], source.date1904);
      if (v === bad) {
        stats.errorCount++; if (errors.length < 20) errors.push({ row: (source.range?.r1 ?? 0) + r + 1, column: c, header: allHeaders[c] });
        values[c] = null;
      } else values[c] = v;
    }
    if (!filters.every(f => filterRow(values[f.column], f))) { stats.filteredRows++; continue; }
    const out = columns.map(c => values[c]);
    if (spec.removeBlankRows && out.every(blank)) { stats.blankRows++; continue; }
    if (spec.removeDuplicates) {
      const key = JSON.stringify(out); if (seen.has(key)) { stats.duplicateRows++; continue; } seen.add(key);
    }
    rows.push({ values, out, index: r });
  }
  if (sorts.length) rows.sort((a, b) => { for (const s of sorts) { const n = compare(a.values[s.column], b.values[s.column]); if (n) return s.direction === 'desc' ? -n : n; } return a.index - b.index; });
  const headers = columns.map(c => allHeaders[c]), matrix = [headers];
  for (const row of rows) matrix.push(row.out);
  stats.outputRows = rows.length;
  const columnFormats = columns.map(c => types[c] === 'text' ? { numFmt: 'text' } : types[c] === 'date' ? { numFmt: 'date' } : types[c] === 'integer' ? { numFmt: 'number', decimals: 0 } : types[c] === 'number' ? { numFmt: 'general' } : null);
  return { matrix, headers, columnFormats, stats, errors };
}
