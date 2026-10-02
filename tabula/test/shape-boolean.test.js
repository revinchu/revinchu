import { test } from 'node:test';
import assert from 'node:assert/strict';
import { mergeShapes, ringsContain } from '../src/shape-boolean.js';
import { parseSvgPath, flattenCommands } from '../src/geometry-path.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip,zip,textOf } from '../src/zip.js';
const rect=(x,y,w,h,extra={})=>({id:'a'+x+y,kind:'rect',x,y,w,h,fill:'#123456',stroke:'#000000',...extra});
const ringsOf=sh=>sh.path.paths.flatMap(p=>flattenCommands(p.commands).map(r=>r.map(([x,y])=>[sh.x+x*sh.w,sh.y+y*sh.h])));
const has=(shapes,x,y)=>shapes.some(s=>ringsContain(ringsOf(s),[x,y]));
const area=shapes=>shapes.reduce((sum,s)=>sum+ringsOf(s).reduce((a,r)=>a+r.reduce((v,p,i)=>v+p[0]*r[(i+1)%r.length][1]-p[1]*r[(i+1)%r.length][0],0)/2,0),0);
const close=(a,b,eps=1e-5)=>assert.ok(Math.abs(a-b)<=eps,`${a} != ${b}`);

test('겹친 사각형 5연산은 면적과 실제 포함영역이 다르고 입력을 보존한다',()=>{
  const shapes=[rect(0,0,100,100),rect(50,0,100,100)],before=structuredClone(shapes);
  for(const [op,expected,samples] of [['union',15000,[true,true,true]],['combine',10000,[true,false,true]],['intersect',5000,[false,true,false]],['subtract',5000,[true,false,false]],['fragment',15000,[true,true,true]]]){
    const out=mergeShapes(shapes,op);close(area(out),expected);assert.deepEqual([has(out,25,50),has(out,75,50),has(out,125,50)],samples);
    assert.ok(out.every(s=>s.fill==='#123456'));assert.equal(out.length,op==='fragment'?3:1);
  }assert.deepEqual(shapes,before);
});
test('동일·접함·분리·완전포함 경계와 빈 결과를 처리한다',()=>{
  const a=rect(0,0,20,20);
  assert.deepEqual(mergeShapes([a,a],'subtract'),[]);assert.deepEqual(mergeShapes([a,a],'combine'),[]);
  close(area(mergeShapes([a,a],'union')),400);
  close(area(mergeShapes([a,rect(20,0,20,20)],'union')),800);
  close(area(mergeShapes([a,rect(20,20,20,20)],'union')),800);
  assert.deepEqual(mergeShapes([a,rect(30,0,20,20)],'intersect'),[]);
  const hole=mergeShapes([rect(0,0,100,100),rect(20,20,60,60)],'subtract');close(area(hole),6400);
  assert.equal(has(hole,50,50),false);assert.equal(has(hole,10,50),true);
  const fragments=mergeShapes([rect(0,0,100,100),rect(20,20,60,60)],'fragment');assert.equal(fragments.length,2);close(area(fragments),10000);
});
test('회전·뒤집기와 오목한 도형·3개 겹침을 실제 좌표로 계산한다',()=>{
  const rotated=mergeShapes([rect(0,0,100,40,{rot:90}),rect(40,-20,20,80)],'intersect');close(area(rotated),1600);
  const t={...rect(0,0,100,100),kind:'rtTriangle',flip:true};
  const out=mergeShapes([t,rect(-1,-1,102,102)],'intersect');assert.equal(has(out,20,90),true);assert.equal(has(out,10,10),false);close(area(out),5000);
  const three=[rect(0,0,60,60),rect(20,0,60,60),rect(40,0,60,60)];close(area(mergeShapes(three,'combine')),3600);close(area(mergeShapes(three,'intersect')),1200);
});
test('원·도넛·cubic 곡선 조합은 구멍과 0.2px 근사 정밀도를 유지한다',()=>{
  const ellipse={...rect(0,0,100,100),kind:'ellipse'};
  const out=mergeShapes([ellipse,rect(-1,-1,102,102)],'intersect');close(area(out),Math.PI*2500,30);
  const donut={...ellipse,kind:'donut'},ring=mergeShapes([donut,rect(-1,-1,102,102)],'intersect');assert.equal(has(ring,50,50),false);assert.equal(has(ring,10,50),true);
  const heart=mergeShapes([{...ellipse,kind:'heart'},rect(0,50,100,50)],'intersect');assert.equal(has(heart,50,75),true);assert.equal(has(heart,50,25),false);
});
test('합성 무작위 사각형은 독립 격자 membership oracle과 일치한다',()=>{
  let seed=281;const rnd=()=>{seed=(seed*1664525+1013904223)>>>0;return seed/4294967296;};
  for(let n=0;n<25;n++){
    const shapes=Array.from({length:3},()=>rect(Math.floor(rnd()*30),Math.floor(rnd()*30),5+Math.floor(rnd()*30),5+Math.floor(rnd()*30)));
    for(const op of ['union','combine','intersect','subtract','fragment']){
      const out=mergeShapes(shapes,op);
      for(let y=.37;y<60;y+=7)for(let x=.41;x<60;x+=7){const flags=shapes.map(s=>x>s.x&&x<s.x+s.w&&y>s.y&&y<s.y+s.h),count=flags.filter(Boolean).length;
        const expected=op==='intersect'?count===3:op==='subtract'?flags[0]&&!flags[1]&&!flags[2]:op==='combine'?count%2===1:count>0;
        assert.equal(has(out,x,y),expected,`${n} ${op} (${x},${y})`);
        if(op==='fragment')assert.equal(out.filter(s=>has([s],x,y)).length,expected?1:0,'조각은 겹치면 안 된다');
      }
    }
  }
});
test('멀리 떨어진 시트 좌표에서도 작은 도형의 면적을 잃지 않는다',()=>{
  const shapes=[rect(10000000,10000000,10,10),rect(10000005,10000000,10,10)];
  const out=mergeShapes(shapes,'intersect');assert.equal(has(out,10000007,10000005),true);close(out[0].w,5);close(out[0].h,10);
});
test('선·해석불가경로·복잡도는 원본을 변경하지 않고 명확히 거절한다',()=>{
  assert.throws(()=>mergeShapes([rect(0,0,10,10),{...rect(0,0,20,20),kind:'line'}],'union'),/지원|열린/);
  assert.throws(()=>mergeShapes([rect(0,0,10,10),{...rect(0,0,20,20),customGeometry:{}}],'union'),/해석/);
  assert.throws(()=>mergeShapes(Array.from({length:17},()=>rect(0,0,10,10)),'union'),/16/);
  const commands=[['M',0,0],...Array.from({length:2100},(_,i)=>['L',i/2100,i%2]),['Z']];
  assert.throws(()=>mergeShapes([rect(0,0,10,10,{path:{paths:[{commands}]}}),rect(0,0,20,20)],'union'),/복잡/);
});
test('SVG path 상대·반사·원호 구문은 정규 명령으로 바뀌며 비정상 입력을 거절한다',()=>{
  assert.deepEqual(parseSvgPath('M1 2h3v4l-1-2z'),[['M',1,2],['L',4,2],['L',4,6],['L',3,4],['Z']]);
  const p=parseSvgPath('M0 0c1 0 1 2 2 2s1-2 2-2t2 0');assert.deepEqual(p[2],['C',3,2,3,0,4,0]);assert.deepEqual(p[3],['Q',4,0,6,0]);
  assert.ok(parseSvgPath('M0 10A10 10 0 0110 0').some(c=>c[0]==='C'));
  assert.throws(()=>parseSvgPath('M0 0L1'),/경로/);assert.throws(()=>parseSvgPath('M0 0 R1 2'),/경로/);
});
test('구멍/회전 조합 결과는 자체 메타데이터 없이 표준 XLSX 경로로 왕복한다',()=>{
  const shapes=mergeShapes([rect(0,0,100,100),rect(20,20,60,60)],'subtract');const wb=new Workbook();wb.setSheetProp(0,'shapes',shapes);
  const files=unzip(writeXlsx(wb));for(const name of Object.keys(files))if(name.endsWith('.xml'))files[name]=textOf(files[name]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'');
  const back=readXlsx(zip(files)).data.sheets[0].shapes;assert.equal(back.length,1);close(area(back),6400,.1);assert.equal(has(back,50,50),false);
});
