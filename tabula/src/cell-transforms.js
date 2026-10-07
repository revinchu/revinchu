import { CellMap } from './cellmap.js';
import { RunColumn } from './run-column.js';
import { storedCellEntries } from './cell-storage.js';

// 서식 빈 셀의 행 구간을 유지하여 행/열 삽입·삭제가 논리 셀 수만큼 메모리를 쓰지 않게 한다.
export function shiftStoredCells(cells, axis, index, count, band = null) {
  const out = new CellMap(), isRow = axis === 'row';
  const put = (r,c,n,cell) => { if(n>0) out.setRunRC(r,c,n,cell); };
  const shift = (r,c,cell,n) => {
    const end=r+n;
    if(isRow) {
      if(band && (c<band[0] || c>band[1])) { put(r,c,n,cell); return; }
      const before=Math.min(end,index); put(r,c,before-r,cell);
      const after=Math.max(r,count<0?index-count:index);
      put(after+count,c,end-after,cell);
    } else {
      const a=band?Math.max(r,band[0]):r,b=band?Math.min(end,band[1]+1):end;
      if(a>=b) {put(r,c,n,cell);return;}
      put(r,c,a-r,cell);put(b,c,end-b,cell);
      if(count<0 && c>=index && c<index-count)return;
      put(a,c>=index?c+count:c,b-a,cell);
    }
  };
  if (cells instanceof CellMap && !band) {
    const shifted = new WeakMap();
    for (const [c,column] of cells.cols) {
      if (!column.blankOnly) { for(const [r,cell,n] of column.storageEntries()) shift(r,c,cell,n); continue; }
      if (!isRow) {
        if(count<0 && c>=index && c<index-count)continue;
        if(column.size)out.replaceImportedColumn(c>=index?c+count:c,column.shareData());
        continue;
      }
      let copy=shifted.get(column.dataKey);
      if(!copy) {
        copy=RunColumn.fromSortedStorage(column.isShared,[]);
        for(const [r,cell,n] of column.storageEntries()) {
          const end=r+n,before=Math.min(end,index),after=Math.max(r,count<0?index-count:index);
          if(before>r)copy.setRun(r,before-r,cell);
          if(end>after)copy.setRun(after+count,end-after,cell);
        }
        shifted.set(column.dataKey,copy);
      }
      if(copy.size)out.replaceImportedColumn(c,copy.shareData());
    }
  } else for(const [r,c,cell,n] of storedCellEntries(cells)) shift(r,c,cell,n);
  return out;
}

export function moveStoredCells(cells, src, dr, dc) {
  const out=new CellMap(),moved=[];
  const dst={r1:src.r1+dr,r2:src.r2+dr,c1:src.c1+dc,c2:src.c2+dc};
  for(const [r,c,cell,n] of storedCellEntries(cells)) {
    const end=r+n;
    const cuts=[r,...[src.r1,src.r2+1,dst.r1,dst.r2+1].filter(x=>x>r&&x<end),end].sort((a,b)=>a-b);
    for(let i=0;i<cuts.length-1;i++) {
      const a=cuts[i],len=cuts[i+1]-a;if(!len)continue;
      if(c>=src.c1&&c<=src.c2&&a>=src.r1&&a<=src.r2)moved.push([a+dr,c+dc,len,cell]);
      else if(!(c>=dst.c1&&c<=dst.c2&&a>=dst.r1&&a<=dst.r2))out.setRunRC(a,c,len,cell);
    }
  }
  // 대상과 원본이 겹칠 때도 원본 값을 마지막에 덮어써 잘라내기 의미를 유지한다.
  for(const [r,c,n,cell] of moved)out.setRunRC(r,c,n,cell);
  return out;
}
