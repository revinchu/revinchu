import { pivotSourceData, pivotChartData, resolvePivot, computePivot } from './pivot.js';

const CATEGORY_TYPES = new Set(['column', 'bar', 'line', 'area', 'combo', 'pie', 'doughnut', 'radar', 'pieOfPie', 'barOfPie', 'waterfall']);
const pivotLabelColumns = new WeakMap();
const validRect = (r) => r && ['r1', 'c1', 'r2', 'c2'].every((k) => Number.isInteger(r[k]) && r[k] >= 0) && r.r1 <= r.r2 && r.c1 <= r.c2;
const contains = (outer, inner) => outer.r1 <= inner.r1 && outer.c1 <= inner.c1 && outer.r2 >= inner.r2 && outer.c2 >= inner.c2;

/**
 * 일반 범위 차트가 피벗의 값 열만 선택했을 때 해당 피벗의 행 레이블 참조를 연결한다.
 * 값 범위는 선택 안의 연속된 잎 행으로만 한정한다. 부분합 사이의 불연속 범위,
 * 일반 셀, 명시 계열은 추측하지 않는다. 차트와 피벗 정의는 변경하지 않는다.
 */
export function inferPivotCategorySeries(wb, hostSi, chart) {
  if (!chart || !CATEGORY_TYPES.has(chart.type) || chart.series != null || chart.pivot || chart.snapshotData || chart.byRows || !validRect(chart.range)) return null;
  const si = chart.sheet ? wb.sheetIndexByName(chart.sheet) : hostSi;
  const sh = wb.sheets[si];
  if (!sh) return null;
  const range = chart.range;
  const defs = [sh.pivot, ...(sh.pivotsExtra ?? [])].filter((d) => d && validRect(d.area) && contains(d.area, range));
  // 겹치는 피벗 정의가 있으면 어느 피벗의 레이블인지 단정하지 않는다.
  if (defs.length !== 1) return null;
  const def = defs[0];
  const top = def.top ?? 0, left = def.left ?? 0;
  if (!Number.isInteger(top) || !Number.isInteger(left) || top < 0 || left < 0) return null;
  let data, labelCols;
  try {
    const source = pivotSourceData(wb, def);
    if (!source) return null;
    const resolved = resolvePivot(source, def);
    // Σ 값을 행에 놓으면 마지막 레이블은 항목이 아니라 지표 이름일 수 있다.
    if (resolved.def.valuesOnRows && resolved.def.values.length > 1) return null;
    data = pivotChartData(source, def);
    // 그릴 때마다 큰 피벗 격자를 다시 만들지 않는다. 피벗 차트 데이터가 바뀌면
    // 원본/정의별 기존 캐시도 새 객체를 반환하므로 레이블 위치도 다시 계산된다.
    if (!pivotLabelColumns.has(data)) pivotLabelColumns.set(data, computePivot(resolved, resolved.def).meta.labelCols);
    labelCols = pivotLabelColumns.get(data);
  } catch {
    // 손상되었거나 지원하지 않는 피벗 정의는 기존 범위 차트 해석에 맡긴다.
    return null;
  }
  if (!labelCols || !data.rows?.length || !data.series?.length) return null;
  const selected = data.series.filter((s) => left + s.col >= range.c1 && left + s.col <= range.c2);
  // 총합계/부분합 열이나 레이블 열이 하나라도 끼면 일반 범위로 유지한다.
  if (selected.length !== range.c2 - range.c1 + 1) return null;
  selected.sort((a, b) => a.col - b.col);
  if (selected.some((s, i) => left + s.col !== range.c1 + i)) return null;
  const rows = data.rows.map((r) => top + r).filter((r) => r >= range.r1 && r <= range.r2);
  if (!rows.length || rows.some((r, i) => i > 0 && r !== rows[i - 1] + 1)) return null;
  const r1 = rows[0], r2 = rows[rows.length - 1], catCol = left + labelCols - 1;
  if (catCol >= range.c1 || catCol < def.area.c1) return null;
  const ref = (col, first = r1, last = r2) => ({ sheet: sh.name, r1: first, c1: col, r2: last, c2: col });
  return selected.map((s) => {
    const col = left + s.col, head = top + s.headRow;
    // 단일 머리글은 참조로 보존한다. 열 그룹의 복합 이름은 메타데이터 이름을 쓴다.
    const name = head >= def.area.r1 && String(wb.getValue(si, head, col) ?? '') === s.name
      ? { ref: ref(col, head, head) } : { text: s.name };
    return { name, cat: ref(catCol), val: ref(col) };
  });
}
