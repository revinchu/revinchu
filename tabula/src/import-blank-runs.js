import { getSharedBlankCell } from './cellmap.js';
import { RunColumn } from './run-column.js';

// 파일의 단조 행/열 순서를 확인한 동안만 서식 빈 칸을 가로 범위로 모읍니다.
// XML의 모든 좌표는 유지하고, 같은 빈 행 패턴을 가진 열은 COW 자료를 공유합니다.
export function createImportedBlankRuns(cells) {
  let spans = [], rowSpans = [], previousRow = new Map(), last = null, lastRow = -1, lastCol = -1, pending = 0, finished = false, disabled = false;
  const stats = { collectedBlanks: 0, horizontalSpans: 0, rectangles: 0, sharedColumns: 0, mixedColumns: 0, uniquePatterns: 0 };
  const styles = new WeakMap(); let nextStyle = 1;
  const styleId = value => { let id = styles.get(value); if (!id) { id = nextStyle++; styles.set(value, id); } return id; };

  function* mergedStorage(column, blanks) {
    const original = Array.from(column.storageEntries()).sort((a, b) => a[0] - b[0]);
    const readers = [original[Symbol.iterator](), blanks[Symbol.iterator]()];
    const next = i => { const entry = readers[i].next(); return entry.done ? null : { row: entry.value[0], value: entry.value[1], end: entry.value[0] + entry.value[2] - 1 }; };
    const current = [next(0), next(1)];
    while (current[0] || current[1]) {
      const start = Math.min(current[0]?.row ?? Infinity, current[1]?.row ?? Infinity);
      let end = Infinity, selected = -1;
      for (let i = 0; i < 2; i++) {
        const item = current[i]; if (!item) continue;
        if (item.row <= start) { if (selected < 0) selected = i; end = Math.min(end, item.end); }
        else end = Math.min(end, item.row - 1);
      }
      yield [start, current[selected].value, end - start + 1];
      for (let i = 0; i < 2; i++) if (current[i] && current[i].row <= end) {
        if (current[i].end <= end) current[i] = next(i); else current[i].row = end + 1;
      }
    }
  }

  function completeRow() {
    const next = new Map();
    for (const span of rowSpans) {
      const key = span.c1 + ':' + span.c2 + ':' + styleId(span.value), previous = previousRow.get(key);
      let rectangle;
      if (previous && previous.r2 + 1 === span.r) { previous.r2 = span.r2; rectangle = previous; }
      else { rectangle = span; spans.push(rectangle); }
      next.set(key, rectangle);
    }
    previousRow = next; rowSpans = []; last = null;
  }

  function finish() {
    if (finished) return { ...stats, disabled };
    completeRow(); finished = true; stats.rectangles = spans.length;
    // add()에 반영했던 논리 수를 빼고 실제 저장 열의 변화량으로 다시 확정합니다.
    cells.n -= pending; pending = 0;
    const events = new Map();
    const event = (col, span, add) => { let list = events.get(col); if (!list) { list = []; events.set(col, list); } list.push([span, add]); };
    for (const span of spans) { event(span.c1, span, true); event(span.c2 + 1, span, false); }
    const boundaries = Array.from(events.keys()).sort((a, b) => a - b), active = new Map(), patterns = new Map();
    const sameRuns = (column, runs) => column.runs.length === runs.length && runs.every(([row, value, count], i) => {
      const run = column.runs[i]; return run.start === row && run.end === row + count - 1 && run.value === value;
    });
    let memoSize = 0;
    for (let i = 0; i < boundaries.length - 1; i++) {
      const col = boundaries[i], changes = events.get(col);
      for (const [span, add] of changes) if (!add && active.get(span.r) === span) active.delete(span.r);
      for (const [span, add] of changes) if (add) active.set(span.r, span);
      if (!active.size) continue;
      const runs = [];
      for (const span of Array.from(active.values()).sort((a, b) => a.r - b.r)) {
        const previous = runs[runs.length - 1];
        const length = span.r2 - span.r + 1;
        if (previous && previous[0] + previous[2] === span.r && previous[1] === span.value) previous[2] += length;
        else runs.push([span.r, span.value, length]);
      }
      let hash = 2166136261, count = 0;
      for (const [row, value, length] of runs) { hash = Math.imul(hash ^ row, 16777619); hash = Math.imul(hash ^ length, 16777619); hash = Math.imul(hash ^ styleId(value), 16777619); count += length; }
      const key = `${hash >>> 0}:${runs.length}:${count}`, matches = patterns.get(key);
      let base = matches?.find(column => sameRuns(column, runs));
      if (!base) {
        base = RunColumn.fromSortedStorage(cells.ensureColumn(col).isShared, runs); stats.uniquePatterns++;
        if (memoSize < 512) { if (matches) matches.push(base); else patterns.set(key, [base]); memoSize++; }
      }
      for (let c = col; c < boundaries[i + 1]; c++) {
        const original = cells.ensureColumn(c);
        if (!original.size) { cells.replaceImportedColumn(c, base.shareData()); stats.sharedColumns++; }
        else { cells.replaceImportedColumn(c, RunColumn.fromSortedStorage(original.isShared, mergedStorage(original, runs))); stats.mixedColumns++; }
      }
    }
    spans = []; rowSpans = []; previousRow.clear(); last = null;
    return { ...stats, disabled };
  }

  function flush() { finish(); disabled = true; return { ...stats, disabled }; }
  function add(r, c, style) {
    if (!Number.isSafeInteger(r) || !Number.isSafeInteger(c) || r < 0 || c < 0 || !style || typeof style !== 'object') throw new Error('가져온 빈 셀 위치/서식이 올바르지 않습니다.');
    const value = getSharedBlankCell(style);
    if (disabled || finished) { cells.setRC(r, c, value); return false; }
    // 비정렬·중복 XML은 기존 Map의 덮어쓰기/삽입 순서로 돌아갑니다.
    if (r < lastRow || r === lastRow && c <= lastCol || cells.hasRC(r, c)) { flush(); cells.setRC(r, c, value); return false; }
    if (r !== lastRow && rowSpans.length) completeRow();
    cells.ensureColumn(c);
    if (last && last.r === r && last.c2 + 1 === c && last.value === value) last.c2 = c;
    else { last = { r, r2: r, c1: c, c2: c, value }; rowSpans.push(last); stats.horizontalSpans++; }
    cells.n++; pending++; stats.collectedBlanks++; lastRow = r; lastCol = c;
    return true;
  }
  return { add, finish, flush, get disabled() { return disabled; }, get stats() { return { ...stats, disabled }; } };
}