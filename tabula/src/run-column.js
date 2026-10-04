// Map 순서와 살아 있는 반복자를 유지하면서 같은 빈 셀의 연속 행만 압축합니다.
export class RunColumn {
  constructor(isShared) { this.isShared = isShared; this.plain = new Map(); this.readers = 0; this.streak = 0; }
  get size() { return this.plain ? this.plain.size : this.n; }
  _locate(row) {
    const a = this.runs; let lo = 0, hi = a.length;
    while (lo < hi) { const mid = (lo + hi) >>> 1; if (a[mid].end < row) lo = mid + 1; else hi = mid; }
    return lo;
  }
  _record(row) {
    const p = this.points.get(row); if (p) return p;
    const run = this.runs[this._locate(row)];
    return run && run.start <= row ? { value: run.value, seq: run.seq + row - run.start } : undefined;
  }
  get(row) {
    if (this.plain) return this.plain.get(row);
    const p = this.points.get(row); if (p) return p.value;
    const run = this.runs[this._locate(row)]; return run && run.start <= row ? run.value : undefined;
  }
  has(row) {
    if (this.plain) return this.plain.has(row);
    if (this.points.has(row)) return true;
    const run = this.runs[this._locate(row)]; return !!run && run.start <= row;
  }
  _orderAdd(row, count, seq) {
    const last = this.order[this.order.length - 1];
    if (last && last.end + 1 === row && last.seq + last.end - last.start + 1 === seq) last.end += count;
    else this.order.push({ start: row, end: row + count - 1, seq });
  }
  _runAdd(start, end, value, seq) {
    const a = this.runs, i = this._locate(start), run = { start, end, value, seq };
    const left = a[i - 1], right = a[i];
    if (left && left.end + 1 === start && left.value === value && left.seq + left.end - left.start + 1 === seq) {
      left.end = end;
      if (right && end + 1 === right.start && right.value === value && seq + end - start + 1 === right.seq) { left.end = right.end; a.splice(i, 1); }
    } else if (right && end + 1 === right.start && right.value === value && seq + end - start + 1 === right.seq) {
      right.start = start; right.seq = seq;
    } else a.splice(i, 0, run);
  }
  _removeRunRow(row, index) {
    const run = this.runs[index], parts = [];
    if (run.start < row) parts.push({ ...run, end: row - 1 });
    if (run.end > row) parts.push({ ...run, start: row + 1, seq: run.seq + row + 1 - run.start });
    this.runs.splice(index, 1, ...parts);
  }
  _convert() {
    if (!this.plain || this.readers) return;
    const plain = this.plain;
    this.plain = null; this.points = new Map(); this.runs = []; this.order = []; this.n = 0; this.seq = 0;
    for (const [row, value] of plain) this.set(row, value);
  }
  set(row, value) {
    if (this.plain) {
      const added = !this.plain.has(row); this.plain.set(row, value);
      if (added && Number.isSafeInteger(row) && this.isShared(value)) {
        this.streak = row === this.lastRow + 1 && value === this.lastValue ? this.streak + 1 : 1;
        this.lastRow = row; this.lastValue = value;
      } else if (added) this.streak = 0;
      if (this.streak >= 128) this._convert(); return this;
    }
    let seq; const point = this.points.get(row);
    if (point) {
      if (!this.isShared(value) || !Number.isSafeInteger(row)) { point.value = value; return this; }
      seq = point.seq; this.points.delete(row);
    } else {
      const i = this._locate(row), run = this.runs[i];
      if (run && run.start <= row) {
        if (run.value === value) return this;
        seq = run.seq + row - run.start; this._removeRunRow(row, i);
      } else { seq = this.seq++; this.n++; this._orderAdd(row, 1, seq); }
    }
    if (Number.isSafeInteger(row) && this.isShared(value)) this._runAdd(row, row, value, seq);
    else this.points.set(row, { value, seq });
    return this;
  }
  /** 새 좌표 구간을 한 번에 복원합니다. 겹치는 구간은 일반 set 의미를 따릅니다. */
  setRun(row, count, value) {
    if (!Number.isSafeInteger(row) || !Number.isSafeInteger(count) || count < 1 || !Number.isSafeInteger(row + count - 1) || !this.isShared(value)) throw new Error('빈 셀 범위 형식이 올바르지 않습니다.');
    if (count < 128 && this.plain) { for (let r = row; r < row + count; r++) this.set(r, value); return this; }
    this._convert();
    // 열린 native Map 반복자가 있으면 그 반복자에서 새 셀을 볼 수 있게 그대로 추가합니다.
    if (this.plain) { for (let r = row; r < row + count; r++) this.set(r, value); return this; }
    const end = row + count - 1, i = this._locate(row); let overlap = this.runs[i]?.start <= end;
    if (!overlap) for (const r of this.points.keys()) if (r >= row && r <= end) { overlap = true; break; }
    if (overlap) { for (let r = row; r <= end; r++) this.set(r, value); return this; }
    this._runAdd(row, end, value, this.seq); this._orderAdd(row, count, this.seq);
    this.seq += count; this.n += count; return this;
  }
  delete(row) {
    if (this.plain) return this.plain.delete(row);
    if (!this.points.delete(row)) {
      const i = this._locate(row), run = this.runs[i]; if (!run || run.start > row) return false;
      this._removeRunRow(row, i);
    }
    this.n--; return true;
  }
  clear() {
    if (this.plain) { this.plain.clear(); this.streak = 0; return; }
    this.points.clear(); this.runs = []; this.n = 0;
    // 진행 중인 반복자는 삭제된 이전 좌표를 건너뛰고 새 삽입을 방문합니다.
    if (!this.readers) this.order = [];
  }
  *entries() {
    this.readers++;
    try {
      if (this.plain) { yield* this.plain; return; }
      for (let i = 0; i < this.order.length; i++) {
        const block = this.order[i];
        for (let row = block.start; row <= block.end; row++) {
          const p = this._record(row); if (p && p.seq === block.seq + row - block.start) yield [row, p.value];
        }
      }
    } finally { this.readers--; if (this.plain && this.streak >= 128) this._convert(); }
  }
  *keys() { for (const [row] of this.entries()) yield row; }
  *values() { for (const [, value] of this.entries()) yield value; }
  [Symbol.iterator]() { return this.entries(); }
  forEach(fn, thisArg) { for (const [row, value] of this) fn.call(thisArg, value, row, this); }
  /** 저장/복원 전용. 한 레코드가 같은 빈 셀의 모든 연속 좌표를 보존합니다. */
  *storageEntries({ bounded = false } = {}) {
    if (this.plain) { for (const [row, value] of this.plain) yield [row, value, 1]; return; }
    // Autosave opts into a live, constant-space traversal. Its version guard
    // rejects edits between chunks. Other callers retain the existing snapshot.
    if (bounded) { yield* this._orderedStorageEntries(); return; }
    const blocks = [...this.runs];
    for (const [row, p] of this.points) blocks.push({ start: row, end: row, ...p });
    blocks.sort((a, b) => a.seq - b.seq);
    for (const block of blocks) yield [block.start, block.value, block.end - block.start + 1];
  }
  *_orderedStorageEntries() {
    this.readers++;
    try {
      for (const block of this.order) {
        let row=block.start;
        while(row<=block.end) {
          const seq=block.seq+row-block.start, point=this.points.get(row);
          if(point) {
            if(point.seq===seq)yield[row,point.value,1];
            row++;continue;
          }
          const run=this.runs[this._locate(row)];
          if(run && run.start<=row && run.seq+row-run.start===seq) {
            const end=Math.min(run.end,block.end);
            yield[row,run.value,end-row+1];row=end+1;
          } else row++;
        }
      }
    } finally {this.readers--;}
  }
  /** sharedMapper는 좌표에 의존하지 않는 빈 셀 정규화에만 사용합니다. */
  *mapValues(mapper, sharedMapper) {
    if (this.plain) {
      for (const [row, value] of this) { const next = mapper(value, row); if (next == null) this.delete(row); else this.set(row, next); yield 1; }
      return;
    }
    const runs = this.runs; this.runs = [];
    for (const run of runs) {
      const next = sharedMapper(run.value), count = run.end - run.start + 1;
      if (next == null) this.n -= count;
      else if (this.isShared(next)) this._runAdd(run.start, run.end, next, run.seq);
      else throw new Error('공유 빈 셀을 일반 셀로 변환할 수 없습니다.');
      yield count;
    }
    for (const [row, point] of this.points) {
      const next = mapper(point.value, row); if (next == null) this.delete(row); else this.set(row, next); yield 1;
    }
  }
}
