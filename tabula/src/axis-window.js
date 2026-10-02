// 화면의 보이는 행·열만 열거한다. Axis의 비트맵 누적 개수와 희소 크기 키를 재사용한다.
// 숨김 원본은 Axis 생성 후 불변이다(크기/숨김 변경 시 GridView.layout이 Axis를 새로 만든다).
const sparseCache = new WeakMap();

function sparseKeys(axis, visible = false) {
  let cache = sparseCache.get(axis);
  if (!cache) { cache = {}; sparseCache.set(axis, cache); }
  const key = visible ? 'visible' : 'hidden';
  if (!cache[key]) cache[key] = axis.keys.filter(i => visible ? axis.size(i) !== 0 : axis.size(i) === 0);
  return cache[key];
}

function lowerBound(a, value) {
  let lo = 0, hi = a.length;
  while (lo < hi) { const mid = (lo + hi) >>> 1; if (a[mid] < value) lo = mid + 1; else hi = mid; }
  return lo;
}

function before(bitmap, i) {
  const end = Math.max(0, Math.min(bitmap.data.length, i - bitmap.start));
  const q = Math.floor(end / 256);
  let n = bitmap.pref[q];
  for (let k = q * 256; k < end; k++) n += bitmap.data[k];
  return n;
}

// 한 비트맵이 보장하는 연속 숨김 구간만 건너뛴다. 여러 숨김 원본이 겹쳐도
// 위치(pos) 차이를 숨김 개수로 추정하지 않으므로 보이는 행을 건너뛰지 않는다.
function bitmapBoundary(bitmap, i, dir) {
  if (dir > 0) {
    const count = before(bitmap, i);
    let lo = i + 1, hi = bitmap.end;
    while (lo < hi) {
      const mid = Math.floor((lo + hi + 1) / 2);
      if (before(bitmap, mid) - count === mid - i) lo = mid; else hi = mid - 1;
    }
    return lo;
  }
  const count = before(bitmap, i + 1);
  let lo = bitmap.start, hi = i;
  while (lo < hi) {
    const mid = Math.floor((lo + hi) / 2);
    if (count - before(bitmap, mid) === i + 1 - mid) hi = mid; else lo = mid + 1;
  }
  return lo - 1;
}

function sparseBoundary(keys, i, dir) {
  const at = lowerBound(keys, i);
  if (keys[at] !== i) return i + dir;
  if (dir > 0) {
    let lo = at, hi = keys.length - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >>> 1;
      if (keys[mid] - i === mid - at) lo = mid; else hi = mid - 1;
    }
    return keys[lo] + 1;
  }
  let lo = 0, hi = at;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (i - keys[mid] === at - mid) hi = mid; else lo = mid + 1;
  }
  return keys[lo] - 1;
}

/** Axis.nextVisible와 같은 계약: 경계 밖/남은 표시 항목 없음은 원래 i를 반환한다. */
export function nextVisibleAxisIndex(axis, i, dir) {
  if (i < 0 || i >= axis.max || axis.size(i) !== 0) return i;
  const step = dir < 0 ? -1 : 1;
  if (axis.def === 0) {
    const keys = sparseKeys(axis, true), at = lowerBound(keys, i);
    const next = step > 0 ? keys[at] : keys[at - 1];
    return next === undefined ? i : next;
  }
  let j = i;
  while (j >= 0 && j < axis.max && axis.size(j) === 0) {
    let next = j + step;
    for (const b of axis.bits) {
      if (j < b.start || j >= b.end || !b.data[j - b.start]) continue;
      const edge = bitmapBoundary(b, j, step);
      next = step > 0 ? Math.max(next, edge) : Math.min(next, edge);
    }
    if (axis.hidden.has(j) || axis.sizes[j] === 0) {
      const edge = sparseBoundary(sparseKeys(axis), j, step);
      next = step > 0 ? Math.max(next, edge) : Math.min(next, edge);
    }
    j = next;
  }
  return j < 0 || j >= axis.max ? i : j;
}

/** 양 끝 포함, Axis 범위 안의 보이는 항목. 빈/역방향 범위는 빈 배열. */
export function visibleAxisIndices(axis, first, last) {
  const out = [], end = Math.min(axis.max - 1, Math.floor(last));
  for (let i = Math.max(0, Math.ceil(first)); i <= end; i++) {
    if (axis.size(i) === 0) {
      const next = nextVisibleAxisIndex(axis, i, 1);
      if (next === i || next > end) break;
      i = next;
    }
    out.push(i);
  }
  return out;
}
