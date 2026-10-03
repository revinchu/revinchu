import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { pictureSvg, pictureMarkup } from '../src/picture-render.js';
import { groupSvg } from '../src/object-group.js';
import { shapeSvg } from '../src/shapes.js';
const src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/a9sAAAAASUVORK5CYII=';
const second='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8/5+hHgAHggJ/PchI7wAAAABJRU5ErkJggg==';
const pic=patch=>({id:'picture',src,x:30,y:40,w:180,h:120,...patch});
const book=patch=>new Workbook({sheets:[{name:'그림',cells:{},images:[pic(patch)]}]});
const native={correction:{brightness:.23,contrast:-.4,sharpness:0},color:{recolor:'grayscale'},border:'#123456',borderW:1.75,borderDash:'dashDot',glow:{color:'#2255aa',size:6.5,opacity:.32},softEdge:3.5,reflection:{size:.6,opacity:.4,gap:7}};
const close=(a,b)=>assert.ok(Math.abs(a-b)<.001,`${a} != ${b}`);
test('밝기·대비·회색조·대시·광선·부드러운 가장자리·반사는 확장 없이 표준으로 왕복한다',()=>{
 const bytes=writeXlsx(book(native)),f=unzip(bytes),xml=textOf(f['xl/drawings/drawing1.xml']);
 for(const tag of ['lum','grayscl','prstDash','glow','softEdge','reflection'])assert.match(xml,new RegExp('<a:'+tag+'\\b'));
 assert.doesNotMatch(xml,/wx:picture/);
 const p=readXlsx(bytes).data.sheets[0].images[0];
 assert.deepEqual(p.correction,native.correction);assert.deepEqual(p.color,native.color);assert.equal(p.borderDash,'dashDot');
 for(const key of ['size','opacity'])close(p.glow[key],native.glow[key]);assert.equal(p.glow.color,native.glow.color);close(p.softEdge,3.5);assert.deepEqual(p.reflection,native.reflection);
});
test('재색칠 듀오톤과 임의 두 색은 표준 값으로 보존한다',()=>{
 for(const color of [{recolor:'blue'},{recolor:'green'},{recolor:'orange'},{duotone:['#112233','#fedcba']}]){
  const back=readXlsx(writeXlsx(book({color}))).data.sheets[0].images[0];
  assert.deepEqual(back.color,color);
 }
});
test('표준에 없는 효과는 준비된 PNG를 저장하고 원본 관계·편집 설정으로 복원한다',()=>{
 const patch={...native,effectPng:second,correction:{brightness:.23,contrast:-.4,sharpness:.7},color:{saturation:1.4,temperature:.2,recolor:'sepia'},artistic:{type:'posterize',amount:.6},originalSrc:second,originalWidth:640,originalHeight:480};
 const original=book(patch),before=original.serialize(),bytes=writeXlsx(original),f=unzip(bytes),xml=textOf(f['xl/drawings/drawing1.xml']);
 assert.doesNotMatch(xml,/<a:lum\b|<a:grayscl\b/);assert.match(xml,/wx:picture/);assert.doesNotMatch(xml,/data:image/);
 const p=readXlsx(bytes).data.sheets[0].images[0];assert.equal(p.src,src);assert.equal(p.originalSrc,second);assert.equal(p.originalWidth,640);assert.deepEqual(p.correction,patch.correction);assert.deepEqual(p.color,patch.color);assert.deepEqual(p.artistic,patch.artistic);assert.equal(p.effectPng,undefined);assert.deepEqual(original.serialize(),before);
});
test('미준비 효과 PNG는 저장 실패를 명확히 알리며 원본을 바꾸지 않는다',()=>{
 const wb=book({artistic:{type:'pencil',amount:.8}}),before=wb.serialize();assert.throws(()=>writeXlsx(wb),/저장용 이미지/);assert.deepEqual(wb.serialize(),before);
});
test('외부 편집기에서 이미지나 보정을 바꾸면 오래된 원본 복원 메타가 새 그림을 덮지 않는다',()=>{
 const f=unzip(writeXlsx(book({effectPng:second,artistic:{type:'pencil',amount:.5}}))),path='xl/drawings/drawing1.xml';
 f[path]=textOf(f[path]).replace(/(<a:blip r:embed="[^"]+">)/,'$1<a:lum bright="12000" contrast="0"/>');
 const p=readXlsx(zip(f)).data.sheets[0].images[0];assert.equal(p.src,second);assert.equal(p.artistic,undefined);assert.equal(p.correction.brightness,.12);
});
test('압축·배경 제거의 원래 매체와 크기를 별도 관계로 보존한다',()=>{
 const p=readXlsx(writeXlsx(book({src:second,originalSrc:src,originalPng:src,originalEmf:'data:image/x-emf;base64,AQID',originalWidth:1024,originalHeight:768}))).data.sheets[0].images[0];
 assert.equal(p.src,second);assert.equal(p.originalSrc,src);assert.equal(p.originalPng,src);assert.equal(p.originalEmf,'data:image/x-emf;base64,AQID');assert.equal(p.originalHeight,768);
});
test('그룹 그림의 네이티브 효과와 원본 편집 메타는 단독 그림과 같다',()=>{
 const image={...pic(native),kind:'picture'},group={id:'g',kind:'group',x:10,y:20,w:220,h:200,groupSize:{w:220,h:200},groupItems:[image]};
 const wb=new Workbook({sheets:[{name:'그룹',cells:{},shapes:[group]}]}),f=unzip(writeXlsx(wb)),path='xl/drawings/drawing1.xml';
 f[path]=textOf(f[path]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'');
 const nativeBack=readXlsx(zip(f)).data.sheets[0].images[0];
 assert.ok(nativeBack.src);assert.deepEqual(nativeBack.correction,native.correction);assert.equal(nativeBack.borderDash,'dashDot');assert.deepEqual(nativeBack.reflection,native.reflection);
});
test('Excel이 생략하는 그림자·반사 기본 속성 때문에 정상 그룹을 외부 수정으로 오인하지 않는다',()=>{
 const group={id:'g',kind:'group',x:10,y:20,w:220,h:200,groupSize:{w:220,h:200},groupItems:[{...pic({...native,shadow:true}),kind:'picture'}]};
 const f=unzip(writeXlsx(new Workbook({sheets:[{name:'그룹',cells:{},shapes:[group]}]}))),path='xl/drawings/drawing1.xml';
 f[path]=textOf(f[path]).replace(/<a:(reflection|outerShdw)\b[^>]*>/g,s=>s.replace(/\s(?:rotWithShape="1"|blurRad="0"|stPos="0"|endA="0"|fadeDir="5400000"|sx="100000")/g,''));
 const back=readXlsx(zip(f));assert.deepEqual(back.warnings,[]);assert.equal(back.data.sheets[0].shapes[0].groupItems[0].id,'picture');
});
test('원본 이미지 관계 번호를 바꿔도 XML 관계 속성으로 원본을 찾아 복원한다',()=>{
 const f=unzip(writeXlsx(book({effectPng:second,artistic:{type:'pencil',amount:.5}})));
 for(const path of ['xl/drawings/drawing1.xml','xl/drawings/_rels/drawing1.xml.rels'])f[path]=textOf(f[path]).replace(/rId(\d+)/g,'originalRelationship$1');
 const p=readXlsx(zip(f)).data.sheets[0].images[0];assert.equal(p.src,src);assert.equal(p.artistic.type,'pencil');
 assert.match(f['xl/drawings/drawing1.xml'],/<wx:source key="src" r:embed="originalRelationship\d+"/);
});
test('격자·그룹 SVG는 같은 보정·자르기·반사·테두리 표현을 사용한다',()=>{
 const p=pic({...native,shadow:{dx:4,dy:5,blur:6,color:'#333333',opacity:.4},crop:{l:.2},flip:true});
 const html=pictureMarkup(p,'grid'),svg=pictureSvg(p,'standalone'),group=groupSvg({id:'group',w:180,h:120,groupItems:[{...p,kind:'picture'}]},shapeSvg);
 for(const output of [html,svg,group]){assert.match(output,/feDropShadow/);assert.match(output,/stroke-dasharray=/);assert.match(output,/feMorphology/);assert.match(output,/feComponentTransfer/);}
 assert.match(html,/class="picture-reflection"/);assert.match(svg,/<mask/);assert.match(group,/<mask/);assert.match(html,/width:125%/);assert.match(svg,/scale\(-1,1\)/);
});

// HTML 필터의 위치 의존 잘림은 tools/picture-visibility.mjs의 실제 픽셀로 검사한다.
test('HTML 효과 경계는 그림 기준이고 SVG 효과 경계와 픽셀 여유가 같다', () => {
 const p=pic({correction:{brightness:.2},shadow:{dx:16,dy:-8,blur:12}}), before=structuredClone(p);
 const html=pictureMarkup(p,'grid'), svg=pictureSvg(p,'saved');
 const region=markup=>Object.fromEntries([...markup.match(/<filter\b[^>]*>/)[0].matchAll(/(\w+)="([^"]*)"/g)].map(([,key,value])=>[key,value]));
 const a=region(html),b=region(svg);
 assert.equal(a.filterUnits,'objectBoundingBox');assert.equal(b.filterUnits,'userSpaceOnUse');
 assert.equal(a.primitiveUnits,'userSpaceOnUse');assert.equal(b.primitiveUnits,'userSpaceOnUse');
 for(const [key,scale] of [['x',p.w],['y',p.h],['width',p.w],['height',p.h]])assert.ok(Math.abs(Number(a[key])*scale-Number(b[key]))<1e-9,key);
 assert.deepEqual(p,before);
});
test('HTML 그림의 필터 경계는 이동 좌표에 의존하지 않고 효과 없는 그림은 필터를 만들지 않는다', () => {
 const p=pic({color:{recolor:'grayscale'},crop:{l:.2}});
 assert.equal(pictureMarkup(p,'same'),pictureMarkup({...p,x:4000,y:9000},'same'));
 assert.doesNotMatch(pictureMarkup(pic({}),'plain'),/<filter\b|filter:url/);
});
