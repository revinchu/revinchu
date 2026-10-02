/** Literal text suggestions from the contiguous column above/below a cell. */
export function cellPickList(book, si, r, c, {maxScan=50000,maxItems=5000} = {}) {
  const seen = new Map(); let scanned=0, limited=false;
  const take = row => {
    const cell=book.getCell(si,row,c), value=book.getValue(si,row,c);
    if (value === null || value === undefined || value === '') return false;
    if (!cell?.formula && typeof value==='string') { const key=value.toLocaleLowerCase('ko'); if(!seen.has(key)) seen.set(key,value); }
    return true;
  };
  const end=book.usedRange(si).rows-1;
  for(const direction of [-1,1]) for(let row=r+direction;row>=0&&row<=end;row+=direction) {
    if(++scanned>maxScan||seen.size>=maxItems){limited=true;break;}
    if(!take(row))break;
  }
  return {items:[...seen.values()].sort((a,b)=>a.localeCompare(b,'ko')),limited};
}
