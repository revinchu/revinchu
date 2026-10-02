import { nextVisibleAxisIndex } from './axis-window.js';

// ───────────── 숨긴 행 모음 ─────────────
// 작은 필터 결과는 { 행: true } 객체, 행이 아주 많으면 비트맵 { __bits: Uint8Array, start, count } (평범한 객체라 복사 · 저장 가능)
/** 행 r 이 숨겨졌는지 */
export const hid = (h, r) => (h ? (h.__bits ? r >= h.start && r - h.start < h.__bits.length && h.__bits[r - h.start] === 1 : !!h[r]) : false);
/** 숨긴 행 수 */
export const hidCount = (h) => (h ? (h.__bits ? h.count : Object.keys(h).length) : 0);
/** 숨긴 행 번호 목록 (비트맵은 limit 까지만) */
export function hidKeys(h, limit = Infinity) {
  if (!h) return [];
  if (!h.__bits) return Object.keys(h).filter((k) => h[k]).map(Number);
  const out = [];
  const b = h.__bits;
  for (let i = 0; i < b.length && out.length < limit; i++) if (b[i]) out.push(h.start + i);
  return out;
}
/** 행 삽입 · 삭제에 맞춰 숨긴 행 옮기기 */
export function shiftHidden(h, index, count) {
  if (!h?.__bits) return null;
  const b = h.__bits;
  const end = h.start + b.length;
  if (count > 0) {
    if (index <= h.start) return { ...h, start: h.start + count };
    if (index >= end) return h;
    const at = index - h.start;
    const nb = new Uint8Array(b.length + count);
    nb.set(b.subarray(0, at));
    nb.set(b.subarray(at), at + count);
    return { __bits: nb, start: h.start, count: h.count };
  }
  const d1 = index;
  const d2 = index - count;
  if (d2 <= h.start) return { ...h, start: h.start + count };
  if (d1 >= end) return h;
  const a1 = Math.max(d1, h.start) - h.start;
  const a2 = Math.min(d2, end) - h.start;
  const nb = new Uint8Array(b.length - (a2 - a1));
  nb.set(b.subarray(0, a1));
  nb.set(b.subarray(a2), a1);
  let cnt = 0;
  for (let i = 0; i < nb.length; i++) cnt += nb[i];
  return { __bits: nb, start: Math.min(h.start, d1), count: cnt };
}

// 행/열 위치 계산 (가상 스크롤용). 기본 크기 + 일부 사용자 지정 크기/숨김만 저장하므로
// 1,000만 행 × 16,384열 전체를 다뤄도 계산량은 사용자 지정 항목 수에만 비례.
export class Axis {
  /**
   * def: 기본 크기, sizes: {index: px}, hidden: {index: true} (여러 개 전달 가능), max: 전체 개수
   */
  constructor(def, sizes = {}, hiddenSets = [], max = 20_000_000) {
    this.def = def;
    this.max = max;
    this.sizes = sizes;
    this.hidden = new Set();
    // 비트맵(수백만 행 필터 결과)은 따로: 구간마다 누적 개수를 두어 위치를 빠르게 계산
    const bitmaps = [];
    for (const h of hiddenSets) {
      if (!h) continue;
      if (h.__bits) { if (h.count && h.__bits.length) bitmaps.push(h); continue; }
      for (const k of Object.keys(h)) if (h[k]) this.hidden.add(Number(k));
    }
    this.bits = prepareBitmaps(bitmaps);
    const keys = new Set([...Object.keys(sizes).map(Number), ...this.hidden]);
    this.keys = [...keys].filter((k) => k >= 0 && k < max && !this.bitHidden(k)).sort((a, b) => a - b);
    this.cum = new Float64Array(this.keys.length + 1);
    this.keys.forEach((k, i) => { this.cum[i + 1] = this.cum[i] + (this.size(k) - def); });
  }

  bitHidden(i) {
    for (const b of this.bits) if (i >= b.start && i < b.end && b.data[i - b.start]) return true;
    return false;
  }

  /** i 앞(i 제외)에 비트맵으로 숨긴 행 수 */
  bitsBefore(i) {
    let n = 0;
    for (const b of this.bits) {
      if (i <= b.start) continue;
      const k = Math.min(i, b.end) - b.start;
      const q = k >> 8;
      n += b.pref[q];
      for (let j = q << 8; j < k; j++) n += b.data[j];
    }
    return n;
  }

  size(i) {
    if (this.hidden.has(i)) return 0;
    if (this.bits.length && this.bitHidden(i)) return 0;
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
    return i * this.def + this.cum[lo] - (this.bits.length ? this.bitsBefore(i) * this.def : 0);
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
    return nextVisibleAxisIndex(this, i, dir);
  }
}

// 여러 표의 필터가 같은 행을 숨겨도 위치에서 한 번만 차감한다.
// 겹치는 구간끼리만 합치므로 떨어진 두 작은 비트맵 사이의 거대한 공백은 할당하지 않는다.
// 단일/비중첩 비트맵은 원래 배열을 읽기 전용으로 재사용하며 입력 배열을 정렬·수정하지 않는다.
function prepareBitmaps(bitmaps) {
  if (bitmaps.length < 2) return bitmaps.map(prepBits);
  const sorted = bitmaps.slice().sort((a, b) => a.start - b.start), prepared = [];
  for (let i = 0; i < sorted.length;) {
    const start = sorted[i].start;
    let end = start + sorted[i].__bits.length, next = i + 1;
    while (next < sorted.length && sorted[next].start < end) {
      end = Math.max(end, sorted[next].start + sorted[next].__bits.length);
      next++;
    }
    if (next === i + 1) prepared.push(prepBits(sorted[i]));
    else {
      const data = new Uint8Array(end - start);
      for (let k = i; k < next; k++) {
        const source = sorted[k].__bits, offset = sorted[k].start - start;
        for (let j = 0; j < source.length; j++) if (source[j]) data[offset + j] = 1;
      }
      prepared.push(prepBits({ start, __bits: data }));
    }
    i = next;
  }
  return prepared;
}

/** 비트맵 → 256행마다 앞쪽 누적 개수 */
function prepBits(h) {
  const data = h.__bits;
  const pref = new Uint32Array((data.length >> 8) + 2);
  let n = 0;
  for (let q = 0; q < pref.length; q++) {
    pref[q] = n;
    const end = Math.min(data.length, (q + 1) << 8);
    for (let j = q << 8; j < end; j++) n += data[j];
  }
  return { start: h.start, end: h.start + data.length, data, pref };
}
