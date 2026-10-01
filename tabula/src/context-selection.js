// 우클릭·행/열 작업은 희소 선택의 바깥 사각형이 아닌 실제 선택을 사용한다.
export function contextMenuKind(selKind, hitKind = 'cell') {
  if (hitKind === 'row' || hitKind === 'col') return hitKind;
  return selKind === 'rows' ? 'row' : selKind === 'cols' ? 'col' : 'cell';
}

export function contextContains(sel, specialCells, r, c) {
  if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) return false;
  if (Array.isArray(specialCells)) return specialCells.some((cell) => cell[0] === r && cell[1] === c);
  return !!sel && r >= sel.r1 && r <= sel.r2 && c >= sel.c1 && c <= sel.c2;
}

function selectionError() {
  const error = new RangeError('선택 범위가 올바르지 않습니다.');
  error.code = 'INVALID_SELECTION';
  return error;
}

/** 양 끝을 포함하는 정렬된 구간. 전체 행/열도 한 구간으로 반환해 전수 순회를 피한다. */
export function selectionAxisRanges(sel, axis, { cells } = {}) {
  if (axis !== 'row' && axis !== 'col') throw selectionError();
  const index = axis === 'row' ? 0 : 1;
  if (cells != null) {
    if (!Array.isArray(cells)) throw selectionError();
    const unique = new Set();
    for (const cell of cells) {
      const n = cell?.[index];
      if (!Number.isSafeInteger(n) || n < 0) throw selectionError();
      unique.add(n);
    }
    const sorted = Array.from(unique).sort((a, b) => a - b), result = [];
    for (const n of sorted) {
      const last = result[result.length - 1];
      if (last && last[1] + 1 === n) last[1] = n;
      else result.push([n, n]);
    }
    return result;
  }
  const a = sel?.[axis === 'row' ? 'r1' : 'c1'], b = sel?.[axis === 'row' ? 'r2' : 'c2'];
  if (!Number.isSafeInteger(a) || !Number.isSafeInteger(b) || a < 0 || b < a) throw selectionError();
  return [[a, b]];
}

/** 제한 초과는 일부 처리하지 않고 거부한다. 큰 전체 선택은 ranges로 별도 처리한다. */
export function selectionAxisTargets(sel, axis, { cells, limit = 20000 } = {}) {
  if (!Number.isSafeInteger(limit) || limit < 0) throw selectionError();
  const ranges = selectionAxisRanges(sel, axis, { cells });
  let count = 0;
  for (const [a, b] of ranges) count += b - a + 1;
  if (count > limit) {
    const error = new RangeError(`선택한 ${axis === 'row' ? '행' : '열'} ${count.toLocaleString('ko-KR')}개는 한 번에 처리할 수 없습니다. 범위를 줄여 주세요.`);
    error.code = 'AXIS_SELECTION_LIMIT';
    error.count = count;
    error.limit = limit;
    throw error;
  }
  const result = [];
  for (const [a, b] of ranges) for (let n = a; n <= b; n++) result.push(n);
  return result;
}
