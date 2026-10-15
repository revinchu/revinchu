// 시트 셀 저장소 (DOM 없음): Map('행,열' → 셀) 과 같은 방법으로 쓸 수 있지만, 안에서는 열마다 Map(행 번호 → 셀) 로 둔다.
// 문자열 키('123456,10')는 백만 개가 넘으면 찾을 때마다 해시 계산 · 캐시 미스로 느리다(칸당 약 0.8µs).
// 행 번호(작은 정수) 키는 약 0.1µs — 수식 계산 · 큰 범위 읽기처럼 칸을 많이 찾는 곳은 getRC/setRC 를 쓴다.
export class CellMap {
  constructor(src = null) {
    this.cols = new Map(); // 열 → Map(행 → 셀)
    this.n = 0;
    if (src) for (const [k, v] of (typeof src.entries === 'function' ? src.entries() : Object.entries(src))) this.set(k, v);
  }

  get size() { return this.n; }

  getRC(r, c) {
    const m = this.cols.get(c);
    return m === undefined ? undefined : m.get(r);
  }

  hasRC(r, c) {
    const m = this.cols.get(c);
    return m !== undefined && m.has(r);
  }

  setRC(r, c, v) {
    let m = this.cols.get(c);
    if (m === undefined) { m = new Map(); this.cols.set(c, m); }
    const before = m.size;
    m.set(r, v);
    if (m.size !== before) this.n++;
    return this;
  }

  deleteRC(r, c) {
    const m = this.cols.get(c);
    if (m === undefined || !m.delete(r)) return false;
    this.n--;
    if (!m.size) this.cols.delete(c);
    return true;
  }

  /** 열 c 의 Map(행 → 셀) (없으면 undefined) */
  col(c) { return this.cols.get(c); }

  get(k) { const i = k.indexOf(','); return this.getRC(+k.slice(0, i), +k.slice(i + 1)); }
  has(k) { const i = k.indexOf(','); return this.hasRC(+k.slice(0, i), +k.slice(i + 1)); }
  set(k, v) { const i = k.indexOf(','); return this.setRC(+k.slice(0, i), +k.slice(i + 1), v); }
  delete(k) { const i = k.indexOf(','); return this.deleteRC(+k.slice(0, i), +k.slice(i + 1)); }
  clear() { this.cols.clear(); this.n = 0; }

  *entries() { for (const [c, m] of this.cols) for (const [r, v] of m) yield [`${r},${c}`, v]; }
  *keys() { for (const [c, m] of this.cols) for (const r of m.keys()) yield `${r},${c}`; }
  *values() { for (const m of this.cols.values()) yield* m.values(); }
  [Symbol.iterator]() { return this.entries(); }
  forEach(fn) { for (const [c, m] of this.cols) for (const [r, v] of m) fn(v, `${r},${c}`, this); }
  /** 문자열 키 없이 훑기: fn(셀, 행, 열) */
  forEachRC(fn) { for (const [c, m] of this.cols) for (const [r, v] of m) fn(v, r, c); }
}

/** Map · 객체 · CellMap → CellMap (이미 CellMap 이면 그대로) */
export function toCellMap(src) {
  return src instanceof CellMap ? src : new CellMap(src);
}
