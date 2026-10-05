// Pivot output is a structural result, never an ordinary sortable cell range.
const valid = a => a && ['r1', 'c1', 'r2', 'c2'].every(k => Number.isInteger(a[k]) && a[k] >= 0) && a.r1 <= a.r2 && a.c1 <= a.c2;
const overlaps = (a, b) => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;
const inside = (a, b) => a.r1 >= b.r1 && a.r2 <= b.r2 && a.c1 >= b.c1 && a.c2 <= b.c2;
export const PIVOT_SORT_RANGE_MESSAGE = '피벗 테이블을 포함한 셀 범위를 직접 정렬할 수 없습니다. 피벗 안의 항목이나 값을 선택하여 정렬하세요.';

export function pivotSortIntersections(sheet, range) {
  if (!valid(range)) return [];
  return [sheet?.pivot, ...(sheet?.pivotsExtra ?? [])].filter(def => valid(def?.area) && overlaps(range, def.area));
}

// Only a selection contained in one pivot can become a semantic pivot sort.
// A mixed range or an ordinary filter touching a pivot must not permute its cells.
export function pivotSortScope(sheet, range, active) {
  const hits = pivotSortIntersections(sheet, range);
  if (!hits.length) return { kind: 'none' };
  const point = active && { r1: active.r, r2: active.r, c1: active.c, c2: active.c };
  if (hits.length === 1 && inside(range, hits[0].area) && valid(point) && inside(point, range)) return { kind: 'pivot', def: hits[0] };
  return { kind: 'mixed' };
}

export function assertOrdinarySortRange(sheet, range) {
  if (pivotSortIntersections(sheet, range).length) {
    const error = new Error(PIVOT_SORT_RANGE_MESSAGE);
    error.code = 'PIVOT_SORT_RANGE';
    throw error;
  }
}
