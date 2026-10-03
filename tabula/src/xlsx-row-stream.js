import { spillOffsets } from './array-cache.js';
import { storedCellEntries } from './cell-storage.js';

/** XML 행 순서를 만들되 압축된 서식 빈 셀을 행별 Map으로 펼치지 않습니다.
 * 추가 메모리는 저장 레코드·행 메타데이터와 현재 행 너비에 비례합니다.
 * rows는 한 번 순회하는 반복자입니다. 원본 셀과 블록은 변경하지 않습니다. */
export function xlsxRowPlan(cells, { rowLimit, colLimit, blocks = [], spills = [], extraRows = [] }) {
  const points = new Map(), runs = [], ranges = [];
  let maxR = 0, maxC = 0;
  const addPoint = (r, c, cell) => {
    let list = points.get(r); if (!list) points.set(r, list = []);
    list.push([c, cell]); maxR = Math.max(maxR, r); maxC = Math.max(maxC, c);
  };
  for (const [r,c,cell,count] of storedCellEntries(cells)) {
    if (r >= rowLimit || c >= colLimit) continue;
    const end = Math.min(rowLimit, r + count) - 1;
    if (count === 1) addPoint(r,c,cell);
    else { runs.push({ start:r, end, c, cell }); ranges.push([r,end]); maxR=Math.max(maxR,end);maxC=Math.max(maxC,c); }
  }
  for (const sp of spills) for (const [i, j] of spillOffsets(sp)) {
    const r=sp.r+i, c=sp.c+j;
    if(r>=rowLimit||c>=colLimit)continue;
    if (cells.hasRC ? cells.hasRC(r,c) : cells.has(`${r},${c}`)) continue;
    addPoint(r,c,{raw:'',spilled:true});
  }
  for (const r of points.keys()) ranges.push([r,r]);
  for (const value of extraRows) { const r=Number(value); if(Number.isInteger(r)&&r>=0&&r<rowLimit)ranges.push([r,r]); }
  for (const b of blocks) {
    if (b.r0>=rowLimit || b.c0>=colLimit || !b.n) continue;
    const end=Math.min(b.r0+b.n,rowLimit)-1;ranges.push([b.r0,end]);
    maxR=Math.max(maxR,end);maxC=Math.max(maxC,Math.min(colLimit,b.c0+b.cols.length)-1);
  }
  ranges.sort((a,b)=>a[0]-b[0]||a[1]-b[1]);runs.sort((a,b)=>a.start-b.start||a.c-b.c);
  function* rows() {
    let runIndex=0,last=-1;const active=new Map();
    for (const [start,end] of ranges) for(let r=Math.max(start,last+1);r<=end;r++) {
      for(const [c,run] of active)if(run.end<r)active.delete(c);
      while(runIndex<runs.length&&runs[runIndex].start<=r){const run=runs[runIndex++];if(run.end>=r)active.set(run.c,run);}
      const list=points.get(r)??[];points.delete(r);
      for(const run of active.values())list.push([run.c,run.cell]);
      last=r;yield [r,list];
    }
  }
  return { maxR,maxC,rows:rows() };
}
