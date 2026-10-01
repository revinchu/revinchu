import { MAX_ROWS, MAX_COLS } from './formula.js';

/** Excel-style neighbouring data region, preserving the existing diagonal-edge rule.
 * Empty border intervals only expand. Remember their scanned extent instead of
 * rescanning the entire growing empty right column for every new data row (O(n²)).
 */
export function currentDataRegion(r, c, isEmpty, maxRows = MAX_ROWS, maxCols = MAX_COLS) {
  const rg = { r1:r, c1:c, r2:r, c2:c }, emptyRows = new Map(), emptyCols = new Map();
  const filled = (fixed, lo, hi, cache, row) => {
    lo = Math.max(0,lo); hi = Math.min((row ? maxCols : maxRows)-1,hi);
    const old = cache.get(fixed);
    const scan = (a,b) => { for(let i=a;i<=b;i++) if(!(row ? isEmpty(fixed,i) : isEmpty(i,fixed))) return true; return false; };
    if (!old) { if(scan(lo,hi))return true; }
    else if (scan(lo,Math.min(hi,old.lo-1)) || scan(Math.max(lo,old.hi+1),hi)) return true;
    cache.set(fixed,{lo:old ? Math.min(lo,old.lo) : lo,hi:old ? Math.max(hi,old.hi) : hi});
    return false;
  };
  for(let changed=true,guard=0;changed&&guard<200000;guard++) {
    changed=false;
    if(rg.r1>0&&filled(rg.r1-1,rg.c1-1,rg.c2+1,emptyRows,true)){rg.r1--;changed=true;}
    if(rg.r2<maxRows-1&&filled(rg.r2+1,rg.c1-1,rg.c2+1,emptyRows,true)){rg.r2++;changed=true;}
    if(rg.c1>0&&filled(rg.c1-1,rg.r1-1,rg.r2+1,emptyCols,false)){rg.c1--;changed=true;}
    if(rg.c2<maxCols-1&&filled(rg.c2+1,rg.r1-1,rg.r2+1,emptyCols,false)){rg.c2++;changed=true;}
  }
  return rg;
}
