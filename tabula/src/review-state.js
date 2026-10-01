// 보호 범위·메모 표시 상태. 인증된 암호는 파일이 아닌 앱의 현재 세션에서만 보관한다.
export function rangeContains(range, r, c) { return r >= range.r1 && r <= range.r2 && c >= range.c1 && c <= range.c2; }
export function rangeIntersects(a, b) { return a.r1 <= b.r2 && a.r2 >= b.r1 && a.c1 <= b.c2 && a.c2 >= b.c1; }
export function protectedRangeKey(range) { return JSON.stringify([range.name, range.ranges, range.hash ?? null, range.modern ?? null, range.securityDescriptor ?? null]); }
export function rangeIsUnlocked(range, grants) { return (!range.hash && !range.modern?.hashValue && !range.securityDescriptor) || !!grants?.has(protectedRangeKey(range)); }
export function cellInEditRange(ranges, grants, r, c) { return (ranges ?? []).some(x => rangeIsUnlocked(x, grants) && (x.ranges ?? []).some(a => rangeContains(a, r, c))); }
/** 사각형의 합집합으로 덮였는지 셀마다 순회하지 않고 판정한다. */
export function rangesCover(target, covers) {
  let pending = [{ ...target }];
  for (const cover of covers) {
    const next = [];
    for (const rect of pending) {
      if (!rangeIntersects(rect, cover)) { next.push(rect); continue; }
      const a = { r1: Math.max(rect.r1, cover.r1), c1: Math.max(rect.c1, cover.c1), r2: Math.min(rect.r2, cover.r2), c2: Math.min(rect.c2, cover.c2) };
      if (rect.r1 < a.r1) next.push({ ...rect, r2: a.r1 - 1 });
      if (rect.r2 > a.r2) next.push({ ...rect, r1: a.r2 + 1 });
      if (rect.c1 < a.c1) next.push({ r1: a.r1, r2: a.r2, c1: rect.c1, c2: a.c1 - 1 });
      if (rect.c2 > a.c2) next.push({ r1: a.r1, r2: a.r2, c1: a.c2 + 1, c2: rect.c2 });
    }
    pending = next;
    if (!pending.length) return true;
    if (pending.length > 10000) return false;
  }
  return !pending.length;
}
export function noteVisible(state, r, c) { const k = `${r},${c}`; return state?.states?.[k] ?? !!state?.all; }
export function setNoteVisibility(state, r, c, show) { return { ...state, states: { ...(state?.states ?? {}), [`${r},${c}`]: !!show } }; }
export function shiftNoteVisibility(state, axis, index, count) {
  if (!state) return state;
  const states = {};
  for (const [key, value] of Object.entries(state.states ?? {})) {
    const p = key.split(',').map(Number), at = axis === 'row' ? 0 : 1, n = p[at];
    if (count < 0 && n >= index && n < index - count) continue;
    if (n >= index) p[at] += count;
    states[p.join(',')] = value;
  }
  return { ...state, states };
}
