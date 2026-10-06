// 피벗의 결과 범위는 머리글·필터·총합계와 빈 구분 행을 포함한다.
const valid = area => area && ['r1', 'c1', 'r2', 'c2'].every(key => Number.isInteger(area[key]) && area[key] >= 0) && area.r1 <= area.r2 && area.c1 <= area.c2;
const overlaps = (a, b) => a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1;

/** 계산 결과가 차지할 셀 범위. 빈 피벗도 시작 셀 한 칸은 예약한다. */
export function pivotOutputArea(grid, def) {
  const r1 = def.top ?? 0, c1 = def.left ?? 0;
  let width = 0;
  for (const row of grid) width = Math.max(width, row.length);
  return { r1, c1, r2: r1 + Math.max(0, grid.length - 1), c2: c1 + Math.max(0, width - 1) };
}

/** 교체하는 정의 자체만 제외하고, 처음 겹치는 다른 피벗 정의를 반환한다. */
export function pivotAreaConflict(area, otherDefs, replacing = null) {
  if (!valid(area)) return null;
  for (const def of otherDefs) {
    if (!def || def === replacing) continue;
    let other = def.area;
    if (!valid(other)) {
      // 아직 결과가 없더라도 명시된 시작 셀은 다른 피벗이 덮어쓸 수 없다.
      if (!Number.isInteger(def.top) || def.top < 0 || !Number.isInteger(def.left) || def.left < 0) continue;
      other = { r1: def.top, c1: def.left, r2: def.top, c2: def.left };
    }
    if (overlaps(area, other)) return def;
  }
  return null;
}
