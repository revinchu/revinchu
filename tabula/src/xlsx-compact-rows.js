import { xlsxRowPlan } from './xlsx-row-stream.js';
import { spillOffsets } from './array-cache.js';

/** Ordinary cells use numeric row keys; styled blank runs remain intervals.
 * This planner does not copy cell records into a second row Map. */
export function xlsxBoundedRowPlan(cells, options) {
  if (!(cells.cols instanceof Map)) return xlsxRowPlan(cells, options);
  const { rowLimit, colLimit, blocks = [], spills = [], extraRows = [] } = options;
  const heap = [], spillRows = new Map(); let maxR = 0, maxC = 0;
  const push = cursor => {
    if (!Number.isFinite(cursor.r)) return;
    let at = heap.length; heap.push(cursor);
    while (at) { const parent = (at - 1) >>> 1; if (heap[parent].r <= cursor.r) break; heap[at] = heap[parent]; at = parent; }
    heap[at] = cursor;
  };
  const pop = () => {
    const first = heap[0], last = heap.pop();
    if (heap.length) {
      let at = 0;
      while (at * 2 + 1 < heap.length) {
        let child = at * 2 + 1; if (child + 1 < heap.length && heap[child + 1].r < heap[child].r) child++;
        if (heap[child].r >= last.r) break; heap[at] = heap[child]; at = child;
      }
      heap[at] = last;
    }
    return first;
  };
  for (const [c, column] of cells.cols) {
    if (c < 0 || c >= colLimit) continue;
    let points = []; const runs = [];
    const entries = column.storageEntries ? column.storageEntries() : column.entries();
    for (const [r, , count = 1] of entries) {
      if (r < 0 || r >= rowLimit) continue;
      const end = Math.min(rowLimit, r + count) - 1;
      maxR = Math.max(maxR, end); maxC = Math.max(maxC, c);
      if (count === 1) points.push(r); else runs.push([r, end]);
    }
    const keys = Uint32Array.from(points); points = null; keys.sort(); runs.sort((a,b) => a[0] - b[0]);
    let point = 0, run = 0, runRow = runs[0]?.[0] ?? Infinity;
    const cursor = { c, r: Math.min(keys[0] ?? Infinity, runRow), column,
      advance() {
        if (keys[point] === this.r) point++;
        if (runRow === this.r) { if (++runRow > runs[run][1]) runRow = runs[++run]?.[0] ?? Infinity; }
        this.r = Math.min(keys[point] ?? Infinity, runRow);
      },
    };
    push(cursor);
  }
  for (const sp of spills) for (const [i,j] of spillOffsets(sp)) {
    const r = sp.r + i, c = sp.c + j;
    if (r < 0 || c < 0 || r >= rowLimit || c >= colLimit || cells.hasRC(r,c)) continue;
    let row = spillRows.get(r); if (!row) spillRows.set(r,row=[]);
    row.push([c,{raw:'',spilled:true}]); maxR = Math.max(maxR,r); maxC = Math.max(maxC,c);
  }
  const meta = new Set();
  for (const value of extraRows) { const r=Number(value); if(Number.isInteger(r)&&r>=0&&r<rowLimit)meta.add(r); }
  for (const r of spillRows.keys()) meta.add(r);
  const extra = Uint32Array.from(meta); extra.sort(); let extraIndex=0;
  if(extra.length)push({c:-1,r:extra[0],advance(){this.r=extra[++extraIndex]??Infinity;}});
  for (const b of blocks) {
    if (b.r0 >= rowLimit || b.c0 >= colLimit || !b.n) continue;
    const end = Math.min(rowLimit,b.r0+b.n)-1;
    maxR=Math.max(maxR,end);maxC=Math.max(maxC,Math.min(colLimit,b.c0+b.cols.length)-1);
    push({c:-1,r:b.r0,advance(){this.r=this.r<end?this.r+1:Infinity;}});
  }
  function* rows() {
    try {
      while (heap.length) {
        const r = heap[0].r, row = spillRows.get(r) ?? []; spillRows.delete(r);
        while (heap.length && heap[0].r === r) {
          const cursor = pop();
          if (cursor.c >= 0) row.push([cursor.c,cursor.column.get(r)]);
          cursor.advance(); push(cursor);
        }
        yield [r,row];
      }
    } finally { heap.length=0; spillRows.clear(); }
  }
  return {maxR,maxC,rows:rows()};
}
