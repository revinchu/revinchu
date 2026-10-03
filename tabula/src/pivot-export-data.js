// XLSX 피벗 쓰기는 열 큐브를 직접 읽는다. 전체 원본의 row 배열 사본을 만들지 않는다.
export function pivotExportData(cube, keepBlank = false) {
 const cols=cube.header.map((_,c)=>cube.col(c));
 let indices=null,length=cube.n;
 if(!keepBlank)for(let r=0;r<cube.n;r++){
  const blank=cols.every(col=>{const v=col.get(r);return v===null||v==='';});
  if(blank&&!indices){indices=new Uint32Array(cube.n);for(let i=0;i<r;i++)indices[i]=i;length=r;}
  else if(!blank&&indices)indices[length++]=r;
 }
 const value=(r,c)=>cols[c]?.get(indices?indices[r]:r)??null;
 return{length,value,values(c,predicate){return{*[Symbol.iterator](){for(let r=0;r<length;r++)if(!predicate||predicate(r))yield value(r,c);}};}};
}

/** 배열 없이 반복 가능한 열 값을 조사한다. 숫자 경계·빈 항목 의미를 기존 writer와 유지한다. */
export function pivotValueStats(values) {
 let count=0,numbers=0,blanks=0,missing=0,emptyStrings=0,hasString=false,integers=true,min=Infinity,max=-Infinity;
 for(const value of values){count++;if(typeof value==='number'){numbers++;min=Math.min(min,value);max=Math.max(max,value);integers&&=Number.isInteger(value);}else if(value===null||value===''){blanks++;if(value===null)missing++;else emptyStrings++;}else hasString=true;}
 return{count,numbers,blanks,missing,emptyStrings,hasString,integers,min,max};
}
