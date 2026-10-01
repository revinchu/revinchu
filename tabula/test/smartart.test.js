import { test } from 'node:test';
import assert from 'node:assert/strict';
import { SMARTART_LAYOUTS, SMARTART_CATEGORIES, newSmartArt, normalizeSmartArt, smartArtParts, editSmartArt } from '../src/smartart.js';
import { smartArtSvg } from '../src/smartart-render.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';

const PNG = 'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
const model = () => ({ version: 1, layout: 'hierarchy', palette: ['#4472c4'], nodes: [{ id:'a',text:'팀',level:0 },{ id:'b',text:'부서',level:1 },{ id:'c',text:'담당',level:2 },{ id:'d',text:'다른 부서',level:1 }] });
test('SmartArt 8범주·20배치의 모든 항목과 유한한 양수 크기를 작은/큰 도형에 보존한다', () => {
  assert.equal(SMARTART_LAYOUTS.length,20);assert.equal(new Set(SMARTART_LAYOUTS.map(x=>x.category)).size,SMARTART_CATEGORIES.length);
  for(const layout of SMARTART_LAYOUTS) for(const count of [1,3,60]) for(const [w,h] of [[180,120],[600,340]]) {
    const shape=newSmartArt(layout.id,{w,h}); shape.smartArt.nodes=Array.from({length:count},(_,i)=>({id:`n${i}`,text:`값 ${i}`,level:layout.category==='계층'&&i?1:0}));
    const parts=smartArtParts(shape); assert.deepEqual(new Set(parts.filter(p=>p.smartArtNode).map(p=>p.smartArtNode)),new Set(shape.smartArt.nodes.map(n=>n.id)));
    for(const p of parts){for(const k of ['x','y','w','h'])assert.ok(Number.isFinite(p[k]),`${layout.id}/${k}`);assert.ok(p.w>0&&p.h>0);assert.ok(p.x>=-.01&&p.y>=-.01);assert.ok(p.x+p.w<=w+1&&p.y+p.h<=h+1,`${layout.id} ${JSON.stringify(p)}`);}
  }
});
test('SmartArt 텍스트·승격·강등·순서는 원본을 변경하지 않고 하위 항목을 함께 이동한다',()=>{
  const original=model(), before=structuredClone(original), moved=editSmartArt(original,'b','down');
  assert.deepEqual(moved.nodes.map(n=>n.id),['a','d','b','c']);assert.deepEqual(original,before);
  assert.deepEqual(editSmartArt(moved,'b','up').nodes.map(n=>n.id),['a','b','c','d']);
  const promoted=editSmartArt(original,'b','promote');assert.deepEqual(promoted.nodes.map(n=>n.level),[0,0,1,1]);
  const demoted=editSmartArt(original,'d','demote');assert.equal(demoted.nodes[3].level,2);
  assert.equal(editSmartArt(original,'a','text','안전 <제목>').nodes[0].text,'안전 <제목>');
  const removed=editSmartArt(original,'b','delete');assert.deepEqual(removed.nodes.map(n=>[n.id,n.level]),[['a',0],['c',1],['d',1]]);
});
test('SmartArt 한도·그림·수준검사 및 스타일/색 배열 별칭을 통제한다',()=>{
  const value=model(), clean=normalizeSmartArt(value);clean.nodes[0].text='다름';clean.palette[0]='#000000';assert.equal(value.nodes[0].text,'팀');assert.equal(value.palette[0],'#4472c4');
  assert.throws(()=>normalizeSmartArt({...value,nodes:[]}),/1~60/);assert.throws(()=>normalizeSmartArt({...value,nodes:Array.from({length:61},()=>value.nodes[0])}),/1~60/);
  assert.throws(()=>normalizeSmartArt({...value,nodes:[{text:'x',picture:'javascript:alert(1)'}]}),/PNG/);
  const n=normalizeSmartArt({...value,nodes:[{id:'a',text:'x',level:5},{id:'a',text:'y',level:5}]});assert.deepEqual(n.nodes.map(x=>x.level),[0,1]);assert.notEqual(n.nodes[0].id,n.nodes[1].id);
});
test('SmartArt SVG는 텍스트를 escape하고 선택ID·그림·스타일을 렌더한다',()=>{
  const shape=newSmartArt('pictureCards');shape.smartArt.nodes[0].text='<script>alert(1)</script>';shape.smartArt.nodes[0].picture=PNG;shape.smartArt.style='outline';
  const svg=smartArtSvg(shape);assert.ok(!svg.includes('<script>'));assert.match(svg,/&lt;script&gt;/);assert.match(svg,/<image href="data:image\/png/);assert.match(svg,/data-smartart-node=/);assert.match(svg,/fill="#ffffff"/);
  assert.match(smartArtSvg({...shape,flip:true,flipV:true}),/translate\(600,340\) scale\(-1,-1\)/);
  assert.match(smartArtSvg({...shape,flip:true}),/translate\(600,0\) scale\(-1,1\)/);
});
test('SmartArt 한 개 수정은 Undo/Redo 및 WIXEL JSON 복원에서 모델을 보존한다',()=>{
  const wb=new Workbook(),shape=newSmartArt('process');wb.transact(()=>wb.setSheetProp(0,'shapes',[shape]));
  const edited={...shape,smartArt:editSmartArt(shape.smartArt,shape.smartArt.nodes[0].id,'text','수정')};wb.transact(()=>wb.setSheetProp(0,'shapes',[edited]));
  wb.undo();assert.notEqual(wb.sheets[0].shapes[0].smartArt.nodes[0].text,'수정');wb.redo();assert.equal(wb.sheets[0].shapes[0].smartArt.nodes[0].text,'수정');
  const back=new Workbook(wb.serialize());assert.deepEqual(back.sheets[0].shapes,wb.sheets[0].shapes);
});
test('SmartArt 20배치 표준 그룹 저장과 편집모델 XLSX 왕복',()=>{
  const wb=new Workbook();wb.sheets[0].shapes=SMARTART_LAYOUTS.map((x,i)=>newSmartArt(x.id,{x:10,y:i*360,w:600,h:340}));
  wb.sheets[0].shapes.at(-1).smartArt.nodes[0].picture=PNG;
  const bytes=writeXlsx(wb),files=unzip(bytes),xml=textOf(files['xl/drawings/drawing1.xml']),root=parseXml(xml),back=readXlsx(bytes).data.sheets[0].shapes;
  assert.equal(descendants(root,'grpSp').length,20);assert.equal(back.length,20);assert.equal(descendants(root,'pic').length,1);
  for(let i=0;i<20;i++){assert.deepEqual(back[i].smartArt,wb.sheets[0].shapes[i].smartArt);assert.equal(back[i].w,600);assert.equal(back[i].h,340);}
  const ids=descendants(root,'cNvPr').map(n=>n.attrs.id);assert.equal(new Set(ids).size,ids.length);assert.match(xml,/a:chExt/);assert.ok(!xml.includes('<dgm:'));
});
test('WIXEL 확장을 제거해도 Excel 표준 도형과 그림 및 항목 텍스트가 남는다',()=>{
  const wb=new Workbook(),shape=newSmartArt('pictureCards',{x:50,y:70,w:600,h:340});shape.smartArt.nodes[0].picture=PNG;shape.smartArt.nodes[1].text='표준 텍스트 유지';wb.sheets[0].shapes=[shape];
  const files=unzip(writeXlsx(wb)),path='xl/drawings/drawing1.xml';files[path]=new TextEncoder().encode(textOf(files[path]).replace(/<a:extLst>.*?<\/a:extLst>/gs,''));
  const back=readXlsx(zip(files)).data.sheets[0];assert.equal(back.images.length,1);assert.ok(back.shapes.some(s=>s.text==='표준 텍스트 유지'));assert.ok(back.shapes.every(s=>s.kind!=='smartart'));assert.ok(back.shapes.every(s=>s.x>=50&&s.y>=70));
});
test('일반 중첩그룹의 원래 좌표계와 현재 크기·회전·그림을 XLSX 보존한다',()=>{
  const group={id:'g',kind:'group',x:20,y:40,w:400,h:200,rot:15,groupSize:{w:200,h:100},groupItems:[{id:'box',kind:'rect',x:10,y:10,w:80,h:60,text:'사각형',fill:'#4472c4'},{id:'pic',kind:'picture',src:PNG,x:100,y:10,w:80,h:60},{id:'child',kind:'group',x:2,y:2,w:10,h:10,groupSize:{w:10,h:10},groupItems:[{id:'inner',kind:'ellipse',x:0,y:0,w:8,h:8}]}]};
  const wb=new Workbook();wb.sheets[0].shapes=[group];const out=readXlsx(writeXlsx(wb)).data.sheets[0].shapes[0];assert.equal(out.rot,15);assert.deepEqual(out.groupSize,group.groupSize);assert.deepEqual(out.groupItems,group.groupItems);assert.equal(out.w,400);
});
test('Excel 표준 자식도형 텍스트·색·좌표 변경은 이전 SmartArt 메타보다 우선한다',()=>{
  const wb=new Workbook(),shape=newSmartArt('list');wb.sheets[0].shapes=[shape];const bytes=writeXlsx(wb),path='xl/drawings/drawing1.xml';
  for(const [from,to] of [['<a:t>항목 1</a:t>','<a:t>Excel에서 수정</a:t>'],['val="4472C4"','val="FF0000"'],['x="152400"','x="162400"']]){
    const files=unzip(bytes),xml=textOf(files[path]);assert.ok(xml.includes(from));files[path]=new TextEncoder().encode(xml.replace(from,to));const result=readXlsx(zip(files)),back=result.data.sheets[0];assert.equal(back.shapes[0].kind,'group');assert.ok(back.shapes[0].groupItems.length>=3);assert.match(result.warnings.join(' '),/외부에서 수정/);
    if(to.includes('Excel'))assert.ok(back.shapes[0].groupItems.some(s=>s.text==='Excel에서 수정'));
  }
});
test('외부에서 수정한 그룹은 부모 회전·대칭과 표준 자식 좌표를 잃지 않는다',()=>{
  const wb=new Workbook(),shape=newSmartArt('list',{x:70,y:90,rot:27,flip:true,flipV:true});wb.sheets[0].shapes=[shape];
  const files=unzip(writeXlsx(wb)),path='xl/drawings/drawing1.xml';files[path]=new TextEncoder().encode(textOf(files[path]).replace('<a:t>항목 1</a:t>','<a:t>외부 수정</a:t>'));
  const back=readXlsx(zip(files)).data.sheets[0].shapes[0];assert.equal(back.kind,'group');assert.equal(back.rot,27);assert.equal(back.flip,true);assert.equal(back.flipV,true);assert.deepEqual([back.x,back.y,back.w,back.h],[70,90,600,340]);assert.ok(back.groupItems.every(p=>p.x>=0&&p.y>=0));
});
test('그룹 그림의 투명도·윤곽·둥근모서리·그림자·자르기는 메타 없이 표준 XML에 저장한다',()=>{
  const wb=new Workbook();wb.sheets[0].shapes=[{id:'g',kind:'group',x:20,y:30,w:400,h:200,groupSize:{w:400,h:200},groupItems:[{id:'p',kind:'picture',src:PNG,x:20,y:30,w:200,h:100,opacity:.4,radius:12,border:'#f12345',borderW:2.5,shadow:{dx:3,dy:4,blur:5,color:'#112233',opacity:.6},crop:{l:.1},alt:'그룹 그림'}]}];
  const files=unzip(writeXlsx(wb)),path='xl/drawings/drawing1.xml',xml=textOf(files[path]);assert.match(xml,/alphaModFix amt="40000"/);assert.match(xml,/roundRect/);assert.match(xml,/F12345/);assert.match(xml,/outerShdw/);assert.match(xml,/l="10000"/);
  files[path]=new TextEncoder().encode(xml.replace(/<a:extLst>.*?<\/a:extLst>/gs,''));const back=readXlsx(zip(files)).data.sheets[0].images[0];assert.equal(back.opacity,.4);assert.equal(back.border,'#f12345');assert.ok(Math.abs(back.borderW-2.5)<1/9525);assert.equal(back.radius,12);assert.equal(back.shadow.opacity,.6);assert.equal(back.crop.l,.1);assert.equal(back.alt,'그룹 그림');
});


function hiddenGroupFixture() {
  const wb = new Workbook();
  wb.sheets[0].shapes = [{ id:'g',name:'숨김 부모',kind:'group',x:20,y:30,w:400,h:200,hidden:true,groupSize:{w:400,h:200},groupItems:[
    {id:'box',name:'자식 도형',kind:'rect',x:10,y:10,w:80,h:60,text:'숨긴 도형',hidden:true},
    {id:'pic',name:'자식 그림',kind:'picture',src:PNG,x:100,y:10,w:80,h:60,hidden:true},
    {id:'nested',name:'자식 그룹',kind:'group',x:200,y:10,w:80,h:60,hidden:true,groupSize:{w:80,h:60},groupItems:[{id:'inner',name:'안쪽 도형',kind:'ellipse',x:0,y:0,w:40,h:40}]},
  ]}];
  return wb;
}
function changeDrawingNv(bytes, name, hidden) {
  const files=unzip(bytes),path='xl/drawings/drawing1.xml';
  let changed=false;
  const xml=textOf(files[path]).replace(/<xdr:cNvPr\b[^>]*>/g,tag=>{
    if(!tag.includes('name="'+name+'"'))return tag;
    changed=true;tag=tag.replace(/ hidden="[^"]*"/,'');
    return hidden==null?tag:tag.replace(/(\/?>)$/,' hidden="'+hidden+'"$1');
  });
  assert.ok(changed,'검사 대상 cNvPr 존재');files[path]=new TextEncoder().encode(xml);return zip(files);
}
test('그룹·자식 도형·그림의 숨김은 표준 cNvPr와 메타 왕복에 보존한다',()=>{
  const wb=hiddenGroupFixture(),bytes=writeXlsx(wb),files=unzip(bytes),root=parseXml(textOf(files['xl/drawings/drawing1.xml']));
  const props=new Map(descendants(root,'cNvPr').map(x=>[x.attrs.name,x.attrs.hidden]));
  for(const name of ['숨김 부모','자식 도형','자식 그림','자식 그룹'])assert.equal(props.get(name),'1',name);
  const out=readXlsx(bytes).data.sheets[0].shapes[0];assert.equal(out.hidden,true);assert.deepEqual(out.groupItems,wb.sheets[0].shapes[0].groupItems);
  const shown=readXlsx(changeDrawingNv(bytes,'숨김 부모','0')).data.sheets[0].shapes[0];assert.equal(!!shown.hidden,false);assert.deepEqual(shown.groupItems,out.groupItems);
});
test('Excel에서 바꾼 그룹 자식의 숨김은 옛 메타를 무효화하고 표준 상태를 읽는다',()=>{
  const wb=hiddenGroupFixture(),bytes=writeXlsx(wb);
  for(const name of ['자식 도형','자식 그림','자식 그룹']){
    const result=readXlsx(changeDrawingNv(bytes,name,'false')),out=result.data.sheets[0].shapes[0];
    assert.match(result.warnings.join(' '),/외부에서 수정/);assert.equal(out.kind,'group');
    const found=out.groupItems.find(x=>x.name===name);assert.ok(found,name);assert.equal(!!found.hidden,false,name);
  }
  const inner=readXlsx(changeDrawingNv(bytes,'안쪽 도형','true'));
  assert.match(inner.warnings.join(' '),/외부에서 수정/);
  const walk=o=>[o,...(o.groupItems??[]).flatMap(walk)];
  assert.equal(walk(inner.data.sheets[0].shapes[0]).find(x=>x.name==='안쪽 도형').hidden,true);
});
test('숨김 표기의 1/true 및 보임의 생략/0/false는 같은 표준 상태로 처리한다',()=>{
  const bytes=writeXlsx(hiddenGroupFixture());
  for(const [name,value]of [['자식 도형','true'],['자식 그림','true'],['자식 그룹','true'],['안쪽 도형','0'],['안쪽 도형','false']]){
    const result=readXlsx(changeDrawingNv(bytes,name,value));
    assert.equal(result.warnings.some(x=>x.includes('외부에서 수정')),false,name+'/'+value);
    assert.equal(result.data.sheets[0].shapes[0].groupItems[0].hidden,true);
  }
});

test('메타 없는 표준 그룹을 펼쳐 읽을 때 부모 숨김은 모든 자식에 상속된다',()=>{
  const wb=hiddenGroupFixture(),group=wb.sheets[0].shapes[0];
  for(const item of group.groupItems)delete item.hidden;
  const files=unzip(writeXlsx(wb)),path='xl/drawings/drawing1.xml';
  files[path]=new TextEncoder().encode(textOf(files[path]).replace(/<a:extLst>.*?<\/a:extLst>/gs,''));
  const sh=readXlsx(zip(files)).data.sheets[0],items=[...sh.shapes,...sh.images];
  assert.equal(items.length,3);assert.ok(items.every(x=>x.hidden===true),'숨긴 부모의 도형·그림·중첩그룹 자식 모두 숨김');
});
