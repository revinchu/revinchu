// 행/열 위치 계산 (가상 스크롤용). 기본 크기 + 일부 사용자 지정 크기/숨김만 저장하므로
// 1,048,576행 × 16,384열 전체를 다뤄도 계산량은 사용자 지정 항목 수에만 비례.
export class Axis {
  /**
   * def: 기본 크기, sizes: {index: px}, hidden: {index: true} (여러 개 전달 가능), max: 전체 개수
   */
  constructor(def, sizes = {}, hiddenSets = [], max = 1048576) {
    this.def = def;
    this.max = max;
    this.sizes = sizes;
    this.hidden = new Set();
    for (const h of hiddenSets) for (const k of Object.keys(h ?? {})) if (h[k]) this.hidden.add(Number(k));
    const keys = new Set([...Object.keys(sizes).map(Number), ...this.hidden]);
    this.keys = [...keys].filter((k) => k >= 0 && k < max).sort((a, b) => a - b);
    this.cum = new Float64Array(this.keys.length + 1);
    this.keys.forEach((k, i) => { this.cum[i + 1] = this.cum[i] + (this.size(k) - def); });
  }

  size(i) {
    if (this.hidden.has(i)) return 0;
    return this.sizes[i] ?? this.def;
  }

  isHidden(i) { return this.size(i) === 0; }

  /** i 번째 항목의 시작 위치 */
  pos(i) {
    let lo = 0;
    let hi = this.keys.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (this.keys[mid] < i) lo = mid + 1;
      else hi = mid;
    }
    return i * this.def + this.cum[lo];
  }

  /** 위치 p 를 포함하는 (보이는) 항목 번호 */
  indexAt(p) {
    if (p <= 0) return this.nextVisible(0, 1);
    let lo = 0;
    let hi = this.max - 1;
    while (lo < hi) {
      const mid = (lo + hi + 1) >> 1;
      if (this.pos(mid) <= p) lo = mid;
      else hi = mid - 1;
    }
    return lo;
  }

  /** i 부터 dir 방향으로 처음 보이는 항목 (없으면 i) */
  nextVisible(i, dir) {
    let j = i;
    while (j >= 0 && j < this.max && this.size(j) === 0) j += dir;
    return j < 0 || j >= this.max ? i : j;
  }
}
