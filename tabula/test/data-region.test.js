import { test } from 'node:test';
import assert from 'node:assert/strict';
import { currentDataRegion } from '../src/data-region.js';

// Original implementation, retained only as an oracle on small synthetic grids.
function original(r,c,isEmpty,maxRows,maxCols){
  const rg={r1:r,c1:c,r2:r,c2:c};
  const row=(r,a,b)=>{for(let c=Math.max(0,a);c<=b;c++)if(!isEmpty(r,c))return true;return false;};
  const col=(c,a,b)=>{for(let r=Math.max(0,a);r<=b;r++)if(!isEmpty(r,c))return true;return false;};
  for(let changed=true,guard=0;changed&&guard<200000;guard++){
    changed=false;
    if(rg.r1>0&&row(rg.r1-1,rg.c1-1,rg.c2+1)){rg.r1--;changed=true;}
    if(rg.r2<maxRows-1&&row(rg.r2+1,rg.c1-1,rg.c2+1)){rg.r2++;changed=true;}
    if(rg.c1>0&&col(rg.c1-1,rg.r1-1,rg.r2+1)){rg.c1--;changed=true;}
    if(rg.c2<maxCols-1&&col(rg.c2+1,rg.r1-1,rg.r2+1)){rg.c2++;changed=true;}
  }return rg;
}
test('영역 탐색은 빈칸·대각 연결·분리된 표·가장자리의 기존 결과를 유지한다',()=>{
  let seed=7821;const random=()=>{seed=(Math.imul(seed,1664525)+1013904223)>>>0;return seed/4294967296;};
  for(let sample=0;sample<120;sample++){
    const rows=18,cols=13,filled=new Set();for(let r=0;r<rows;r++)for(let c=0;c<cols;c++)if(random()<sample%5/10)filled.add(`${r},${c}`);
    const empty=(r,c)=>!filled.has(`${r},${c}`);
    for(const [r,c] of [[0,0],[17,12],[9,6],[3,10]])assert.deepEqual(currentDataRegion(r,c,empty,rows,cols),original(r,c,empty,rows,cols));
  }
});
test('2만 행의 빈 경계열 재검사는 셀 수에 비례하며 20만 셀 범위를 그대로 찾는다',()=>{
  let reads=0;const empty=(r,c)=>{reads++;return r>=20000||c>=10;};
  assert.deepEqual(currentDataRegion(0,0,empty,20002,12),{r1:0,c1:0,r2:19999,c2:9});
  assert.ok(reads<100000,`실제 조회 ${reads}`);
});
test('넓고 얕은 표와 중간에서 시작하는 영역도 빈 경계의 재검사를 제한한다',()=>{
  let reads=0;const empty=(r,c)=>{reads++;return r<2||r>=12||c<2||c>=12002;};
  assert.deepEqual(currentDataRegion(6,6000,empty,15,12005),{r1:2,c1:2,r2:11,c2:12001});
  assert.ok(reads<150000,`실제 조회 ${reads}`);
});
