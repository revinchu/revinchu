import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { tableCellDisplayStyle } from '../src/table-format.js';
import { rangeImageSvg } from '../src/range-image-export.js';
import { safeObjectSvg, prepareImageSvg } from '../src/object-image-export.js';
import { parseXml, descendants, child } from '../src/xml.js';

const measureText = text => [...String(text)].length * 7;
const area = (r2=1,c2=2) => ({r1:0,c1:0,r2,c2});
function book(cells={}) {
  const wb=new Workbook();
  wb.sheets[0].defRowH=24;wb.sheets[0].defColW=80;wb.sheets[0].allStyle={align:'general'};
  wb.transact(()=>{for(const [key,raw] of Object.entries(cells)){const [r,c]=key.split(',').map(Number);wb.setInput(0,r,c,raw);}});
  return wb;
}
function render(wb,range=area(),options={}) { const out=rangeImageSvg(wb,0,range,{measureText,...options}); return {...out,xml:parseXml(out.svg)}; }
const nodes=(out,name)=>descendants(out.xml,name);
const contentOf=(out,r,c)=>nodes(out,'g').find(n=>n.attrs['data-cell']===`${r},${c}`);
const textOf=node=>node.text+node.children.map(textOf).join('');
function clipOf(out,r,c) {const id=contentOf(out,r,c).attrs['clip-path'].slice(5,-1);return child(nodes(out,'clipPath').find(n=>n.attrs.id===id),'rect').attrs;}
function state(wb) {
  return JSON.stringify({version:wb.version,sheets:wb.sheets.map(s=>({...s,cells:[...s.cells]})),names:wb.names,
    caches:wb.caches.map(c=>c?[...c]:null),spills:[...wb.spills],spillOwner:[...wb.spillOwner],spillState:wb.spillState,arrayList:wb.arrayList,
    undo:wb.undoStack,redo:wb.redoStack,usedCache:[...wb.usedCache],ctxs:wb.ctxs.length,colIdx:wb.colIdx.length});
}

test('범위 SVG: 정확한 행열 크기·표시 값·선택 범위·배율 독립',()=>{
  const wb=book({'0,0':'1234.5','0,1':'0.125','0,2':'outside','1,0':'<script>&"','2,0':'excluded'});
  wb.transact(()=>{wb.setStyle(0,0,0,{numFmt:'custom',code:'#,##0.00'});wb.setStyle(0,0,1,{numFmt:'percent',decimals:1});});
  const out=render(wb,{r1:0,c1:0,r2:1,c2:1});
  assert.equal(out.width,164);assert.equal(out.height,52);assert.equal(out.rows,2);assert.equal(out.cols,2);
  assert.match(out.svg,/>1,234\.50</);assert.match(out.svg,/>12\.5%</);assert.match(out.svg,/&lt;script&gt;&amp;&quot;/);
  assert.doesNotMatch(out.svg,/foreignObject|excluded|outside/);
  wb.sheets[0].zoom=330;assert.equal(render(wb,{r1:0,c1:0,r2:1,c2:1}).svg,out.svg);
  assert.doesNotThrow(()=>safeObjectSvg(out.svg));
});

test('수동·일반 필터·표 필터의 숨김 행열 합집합을 한 번만 제외',()=>{
  const wb=book({'0,0':'first','4,2':'last'}),s=wb.sheets[0];
  s.hiddenRows={1:true};s.hiddenCols={1:true};s.filter={hidden:{__bits:new Uint8Array([1,1]),start:1,count:2}};
  s.tables=[{r1:0,c1:0,r2:4,c2:2,filter:{hidden:{2:true,3:true}}}];
  const out=render(wb,area(4,2));assert.equal(out.rows,2);assert.equal(out.cols,2);assert.equal(out.width,164);assert.equal(out.height,52);
  assert.deepEqual(nodes(out,'g').filter(n=>n.attrs['data-cell']).map(n=>n.attrs['data-cell']),['0,0','0,2','4,0','4,2']);
});

test('병합은 안쪽 눈금선 없이 한 번만 표시하고 부분 선택에서는 앵커 텍스트를 범위에 자른다',()=>{
  const wb=book({'0,0':'병합 제목'});wb.sheets[0].merges=[area(1,1)];
  let out=render(wb,area(1,1));
  assert.equal(nodes(out,'g').filter(n=>n.attrs['data-cell']).length,1);
  const grid=nodes(out,'path').find(n=>n.attrs.stroke==='#d9d9d9');assert.equal(grid.attrs.d,'M0 0H160V48H0Z');
  assert.equal(textOf(contentOf(out,0,0)),'병합 제목');
  out=render(wb,{r1:1,c1:1,r2:1,c2:1});assert.equal(out.width,84);assert.equal(out.height,28);
  assert.equal(clipOf(out,0,0).x,'-80');assert.equal(clipOf(out,0,0).y,'-24');
  assert.equal(textOf(contentOf(out,0,0)),'병합 제목');
});

test('공유 테두리는 강한 선 한 번, 이중선 외곽은 여백 안의 두 획',()=>{
  const wb=book();wb.sheets[0].noGrid=true;
  wb.transact(()=>{wb.setStyle(0,0,0,{br:true,brs:'thin',brc:'#aaaaaa',bt:true,bts:'double',btc:'#123456'});wb.setStyle(0,0,1,{bl:true,bls:'thick',blc:'#000000'});});
  const out=render(wb,area(0,1)),paths=nodes(out,'path');
  assert.equal(paths.filter(p=>p.attrs.stroke==='#aaaaaa').length,0);
  assert.equal(paths.filter(p=>p.attrs.d==='M80 0V24'&&p.attrs['stroke-width']==='3').length,1);
  assert.deepEqual(paths.filter(p=>p.attrs.stroke==='#123456').map(p=>p.attrs.d),['M0 -1H80','M0 1H80']);
  assert.equal(out.height,28);
});

test('문자열 넘침은 선택 범위의 빈 셀까지, 수식 결과·병합·숫자 0은 장애물',()=>{
  const wb=book({'0,0':'긴 문자열이 다음 빈 칸으로 넘어갑니다','0,3':'0','1,0':'긴 문자열','1,1':'=SEQUENCE(1,2)'});
  let out=render(wb,area(1,3));assert.equal(clipOf(out,0,0).width,'240');assert.equal(clipOf(out,1,0).width,'80');
  out=render(wb,area(0,1));assert.equal(clipOf(out,0,0).width,'160');
  wb.sheets[0].merges=[{r1:0,c1:1,r2:0,c2:2}];assert.equal(clipOf(render(wb,area(0,3)),0,0).width,'80');
});

test('선택 영역 가운데·줄 바꿈·축소·회전은 SVG 글자 배치에 적용',()=>{
  const wb=book({'0,0':'제목','1,0':'abcdefghijklmno','1,1':'abcdefghijklmno','1,2':'회전'});
  wb.transact(()=>{wb.setStyle(0,0,0,{align:'centerContinuous'});wb.setStyle(0,0,1,{align:'centerContinuous'});wb.setStyle(0,1,0,{wrap:true});wb.setStyle(0,1,1,{shrink:true});wb.setStyle(0,1,2,{rotate:45});});
  const out=render(wb);assert.equal(clipOf(out,0,0).width,'160');
  assert.equal(descendants(contentOf(out,0,0),'tspan')[0].attrs.x,'79.5');
  assert.equal(descendants(contentOf(out,1,0),'tspan').length,2);
  assert.ok(Number(child(contentOf(out,1,1),'text').attrs['font-size'])<11*4/3);
  assert.match(child(contentOf(out,1,2),'text').attrs.transform,/rotate\(-45\)/);
});

test('조건부서식 수치·색·데이터 막대·아이콘 및 표 머리글을 적용',()=>{
  const wb=book({'0,0':'5','1,0':'10','0,1':'head','1,1':'body'}),s=wb.sheets[0];
  s.tables=[{r1:0,c1:1,r2:1,c2:1,style:'TableStyleMedium2',header:true,bandRows:true}];
  s.cond=[{...area(1,0),type:'ge',v1:'10',style:{fill:'#ffee00',numFmt:'custom',code:'0.0"점"'}},{...area(1,0),type:'bar',color:'#638ec6'},{...area(1,0),type:'icons',icons:'3Stars'}];
  const out=render(wb,area(1,1));assert.match(out.svg,/>10\.0점</);assert.match(out.svg,/fill="#ffee00"/);assert.match(out.svg,/stop-color="#638ec6"/);
  assert.equal(nodes(out,'svg').length,2);assert.match(out.svg,/font-weight="700"/);assert.doesNotThrow(()=>safeObjectSvg(out.svg));
});

test('출력 계산은 원본 캐시·배열·이름·셀·실행 취소·버전을 바꾸지 않는다',()=>{
  const wb=book({'0,0':'=1+1','0,1':'=SEQUENCE(2,1)','0,2':'=MyNum','1,2':'=NOTSUPPORTED(1)'});
  wb.names=[{name:'MyNum',ref:'=21*2'}];wb.sheets[0].cells.getRC(1,2).cached=99;delete wb.sheets[0].cells.getRC(1,2).dirty;
  const before=state(wb),out=render(wb);assert.match(out.svg,/>2</);assert.match(out.svg,/>42</);assert.match(out.svg,/>99</);
  assert.equal(state(wb),before);
});

test('셀 그림은 안전한 데이터·웹주소만 전달하고 명시 resolver로 완전히 포함한다',async()=>{
  const wb=book();const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg==';
  wb.transact(()=>wb.setCellData(0,0,0,{raw:'',image:{src:'https://example.test/p.png',alt:'샘플'}}));
  const out=render(wb,area(0,0));assert.match(out.svg,/href="https:\/\/example.test\/p.png"/);
  await assert.rejects(prepareImageSvg(out.svg),/그림/);
  let calls=0;const svg=await prepareImageSvg(out.svg,{resolveImage:async url=>{calls++;assert.equal(url,'https://example.test/p.png');return png;}});
  assert.equal(calls,1);assert.match(svg,/data:image\/png/);assert.doesNotMatch(svg,/https:\/\/example/);
  wb.transact(()=>wb.setCellData(0,0,0,{raw:'',image:{src:'javascript:alert(1)'}}));
  assert.throws(()=>render(wb,area(0,0)),/안전/);
});

test('대형 범위·모든 행 숨김은 일부만 저장하지 않고 명시적으로 거절',()=>{
  const wb=book();assert.throws(()=>render(wb,area(1048575,0)),/너무 큽니다/);
  assert.throws(()=>render(wb,area(99,1999)),/100,000셀/);
  wb.sheets[0].defColW=65533;assert.throws(()=>render(wb,area(0,0)),/65,536픽셀/);
  wb.sheets[0].hiddenRows={0:true};assert.throws(()=>render(wb,area(0,0)),/모든 행/);
  assert.throws(()=>render(wb,{r1:-1,c1:0,r2:0,c2:0}),/셀 범위/);
});


test('빠른 표 스타일은 범위 그림 저장에도 기본 흰색·상속 중립값에 가려지지 않는다', () => {
  const wb = book({ '0,0': '제목', '1,0': '자료' });
  wb.baseStyle = { fill: '#ffffff', color: '#000000' };
  wb.sheets[0].allStyle = { fill: '#abcdef' };
  wb.sheets[0].tables = [{ id: 't', name: '표1', r1: 0, c1: 0, r2: 1, c2: 0, header: true, banded: true, style: 'TableStyleMedium4' }];
  wb.transact(() => wb.clearTableVisualFormatting(0, 't'));
  const shown = tableCellDisplayStyle(wb, 0, 0, 0), out = render(wb, area(1, 0));
  assert.ok(nodes(out, 'rect').some(node => node.attrs.fill === shown.fill));
  assert.ok(nodes(out, 'text').some(node => node.attrs.fill === shown.color));
  assert.doesNotMatch(out.svg, /#abcdef|tableStyleInherit/);
});
