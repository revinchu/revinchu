import { RunColumn } from './run-column.js';
// 시트 셀 저장소 (DOM 없음): Map('행,열' → 셀) 과 같은 방법으로 쓸 수 있지만, 안에서는 열마다 Map(행 번호 → 셀) 로 둔다.
// 문자열 키('123456,10')는 백만 개가 넘으면 찾을 때마다 해시 계산 · 캐시 미스로 느리다(칸당 약 0.8µs).
// 행 번호(작은 정수) 키는 약 0.1µs — 수식 계산 · 큰 범위 읽기처럼 칸을 많이 찾는 곳은 getRC/setRC 를 쓴다.
// 서식만 있는 빈 셀은 위치별 가변 상태가 없습니다. 가져온 수백만 빈 셀의
// 값 객체만 공유하고 좌표·셀 수·서식은 그대로 보존합니다. 편집은 새 셀로 교체합니다.
const blankData = new WeakMap();
const blankPrepared = new WeakMap();
const sharedBlanks = new WeakSet();
export function shareBlankCell(value) {
  if (!value || typeof value !== 'object' || value.raw !== '' || !value.style || typeof value.style !== 'object') return value;
  if (sharedBlanks.has(value)) return value;
  let prepared = false;
  for (const key in value) {
    if (key === 'v') { if (value.v !== null) return value; prepared = true; }
    else if (key !== 'raw' && key !== 'style') return value;
  }
  return getSharedBlankCell(value.style, prepared);
}

/** 이미 확인한 빈 셀은 임시 객체를 만들지 않고 공유합니다. */
export function getSharedBlankCell(style, prepared = false) {
  const memo = prepared ? blankPrepared : blankData;
  let cell = memo.get(style);
  if (!cell) {
    cell = Object.freeze(prepared ? { raw: '', style, v: null } : { raw: '', style });
    memo.set(style, cell); sharedBlanks.add(cell);
  }
  return cell;
}

export class CellMap {
  constructor(src = null) {
    this.cols = new Map(); // 열 → Map 호환 행 저장소
    this.n = 0;
    this.candidates = new Map(); // 수식/링크 후보만: 일반 값과 압축 빈 셀은 구조 검사에서 건너뛴다.
    this.formulaAsts = new Map(); // 공유 AST → 참조 수. null은 해석하지 못한 수식이다.
    if (src instanceof CellMap) {
      for (const [c, column] of src.cols) {
        if (column.blankOnly && column.size) { this.cols.set(c, column.shareData()); this.n += column.size; }
        else for (const [r, cell, count] of column.storageEntries()) this.setRunRC(r, c, count, cell);
      }
    } else if (src) for (const [k, v] of (typeof src.entries === 'function' ? src.entries() : Object.entries(src))) this.set(k, v);
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

  /** 가져오기 중 최초 열 순서를 예약합니다. 값이 없어도 cols 위치는 유지됩니다. */
  ensureColumn(c) {
    let column = this.cols.get(c);
    if (!column) { column = new RunColumn((value) => sharedBlanks.has(value)); this.cols.set(c, column); }
    return column;
  }
  /** 가져오기 collector 전용: 기존 point와 후보 셀을 유지한 열 자료를 받아들입니다. */
  replaceImportedColumn(c, column) {
    const old = this.cols.get(c); this.cols.set(c, column);
    this.n += column.size - (old?.size ?? 0);
  }

  setRC(r, c, v) {
    let m = this.cols.get(c);
    if (m === undefined) { m = new RunColumn((value) => sharedBlanks.has(value)); this.cols.set(c, m); }
    const before = m.size;
    const value = shareBlankCell(v);
    m.set(r, value);
    this.trackCandidate(r, c, value);
    if (m.size !== before) this.n++;
    return this;
  }

  /** 동일한 빈 셀 범위를 모든 좌표와 논리 크기를 보존해 복원합니다. */
  setRunRC(r, c, count, value) {
    value = shareBlankCell(value);
    if (count === 1) return this.setRC(r, c, value);
    let m = this.cols.get(c);
    if (!m) { m = new RunColumn((v) => sharedBlanks.has(v)); this.cols.set(c, m); }
    const before = m.size;
    m.setRun(r, count, value); this.n += m.size - before;
    // 빈 셀 구간으로 덮어쓴 수식/링크도 제거하며 논리 행 전체를 펼치지 않는다.
    const candidates = this.candidates.get(c);
    if (candidates) {
      for (const row of candidates.keys()) if (row >= r && row < r + count) this.trackCandidate(row, c, null);
    }
    return this;
  }

  trackCandidate(r, c, value) {
    const keep = value?.formula || value?.link || typeof value?.raw === 'string' && value.raw.startsWith('=');
    let col = this.candidates.get(c);
    const previous = col?.get(r);
    const ast = value?.formula ? value.ast || null : undefined;
    if (previous !== ast) {
      if (previous !== undefined) {
        const count = this.formulaAsts.get(previous) - 1;
        if (count) this.formulaAsts.set(previous, count); else this.formulaAsts.delete(previous);
      }
      if (ast !== undefined) this.formulaAsts.set(ast, (this.formulaAsts.get(ast) ?? 0) + 1);
    }
    if (keep) {
      if (!col) { col = new Map(); this.candidates.set(c, col); }
      // 이전 AST 자체를 보관합니다. 수식마다 별도 객체·셀 참조를 만들지 않고,
      // 같은 셀 객체의 AST를 바꿔 다시 넣어도 이전 요약을 정확히 차감합니다.
      col.set(r, ast);
    } else if (col) {
      col.delete(r);
      if (!col.size) this.candidates.delete(c);
    }
  }

  /** 가져온 raw 후보도 정규화된 formula 플래그로 구분한다 (텍스트 수식 제외). */
  *formulaEntries() {
    for (const [c, col] of this.candidates) {
      const stored = this.cols.get(c);
      for (const r of col.keys()) { const cell = stored?.get(r); if (cell?.formula) yield [r, c, cell]; }
    }
  }
  forEachFormulaRC(fn) {
    for (const [c, col] of this.candidates) {
      const stored = this.cols.get(c);
      for (const r of col.keys()) { const cell = stored?.get(r); if (cell?.formula) fn(cell, r, c); }
    }
  }
  forEachLinkRC(fn) {
    for (const [c, col] of this.candidates) {
      const stored = this.cols.get(c);
      for (const r of col.keys()) { const cell = stored?.get(r); if (cell?.link) fn(cell, r, c); }
    }
  }

  /** 시트 간 의존성은 같은 불변 AST를 공유하는 셀마다 반복해서 검사하지 않는다. */
  forEachFormulaAST(fn) { for (const ast of this.formulaAsts.keys()) fn(ast); }

  /** [시작 행, 열, 셀, 연속 개수]. 일반 반복자는 여전히 셀마다 한 번 반환합니다. */
  *storageEntries(options) {
    for (const [c, m] of this.cols) for (const [r, value, count] of m.storageEntries(options)) yield [r, c, value, count];
  }

  /** 읽기 전용 메타데이터 조회: 범위 하나당 한 번만 호출합니다. */
  forEachStoredRC(fn) { for (const [r, c, value, count] of this.storageEntries()) fn(value, r, c, count); }

  /** 소유권을 넘겨받은 셀을 정규화하며 처리한 논리 셀 수를 내보냅니다. */
  *mapValues(mapper, sharedMapper) {
    const normalizedBlanks = new WeakMap();
    for (const [c, m] of this.cols) {
      const sharedKey = sharedMapper && m.sharedData && m.blankOnly ? m.dataKey : null;
      const hit = sharedKey && normalizedBlanks.get(sharedKey);
      if (hit) {
        const next = hit.shareData(); this.cols.set(c, next); this.n += next.size - m.size;
        if (!next.size) this.cols.delete(c);
        yield m.size; continue;
      }
      let before = m.size;
      for (const count of m.mapValues((value, r) => {
        const next = mapper(value, r, c);
        this.trackCandidate(r, c, next);
        return next;
      }, sharedMapper)) {
        this.n += m.size - before; before = m.size;
        yield count;
      }
      if (sharedKey) normalizedBlanks.set(sharedKey, m);
      if (!m.size) this.cols.delete(c);
    }
  }

  deleteRC(r, c) {
    const m = this.cols.get(c);
    if (m === undefined || !m.delete(r)) return false;
    this.n--;
    this.trackCandidate(r, c, null);
    if (!m.size) this.cols.delete(c);
    return true;
  }

  /** 열 c 의 Map(행 → 셀) (없으면 undefined) */
  col(c) { return this.cols.get(c); }

  get(k) { const i = k.indexOf(','); return this.getRC(+k.slice(0, i), +k.slice(i + 1)); }
  has(k) { const i = k.indexOf(','); return this.hasRC(+k.slice(0, i), +k.slice(i + 1)); }
  set(k, v) { const i = k.indexOf(','); return this.setRC(+k.slice(0, i), +k.slice(i + 1), v); }
  delete(k) { const i = k.indexOf(','); return this.deleteRC(+k.slice(0, i), +k.slice(i + 1)); }
  clear() { this.cols.clear(); this.candidates.clear(); this.formulaAsts.clear(); this.n = 0; }

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
