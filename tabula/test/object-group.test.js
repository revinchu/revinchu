import test from 'node:test';
import assert from 'node:assert/strict';
import {makeObjectGroup,ungroupObjects} from '../src/object-group.js';
import {shapeSvg} from '../src/shapes.js';
const objects=[{id:'a',kind:'rect',x:20,y:30,w:60,h:40,fill:'#4472c4',text:'한글'},{id:'b',kind:'ellipse',x:100,y:60,w:40,h:40,fill:'#ff8800'}];
test('그룹화와 해제는 좌표·텍스트를 유지하고 원본을 변경하지 않는다',()=>{
 const before=structuredClone(objects),g=makeObjectGroup(objects,'g');assert.deepEqual(g.groupSize,{w:120,h:70});assert.equal(g.groupItems[1].x,80);const back=ungroupObjects(g);for(let i=0;i<2;i++)for(const k of Object.keys(objects[i]))assert.deepEqual(back[i][k],objects[i][k]);assert.deepEqual(objects,before);
});
test('그룹 이동·배율·회전·대칭을 해제할 때 자식 중심에 적용한다',()=>{
 const g=makeObjectGroup(objects,'g');g.x=200;g.y=100;g.w*=2;g.h*=2;g.rot=90;g.flip=true;
 const a=ungroupObjects(g)[0];assert.equal(a.w,120);assert.equal(a.h,80);assert.equal(a.rot,90);assert.equal(a.flip,true);assert.ok(Math.abs(a.x-290)<1e-8);assert.ok(Math.abs(a.y-190)<1e-8);
});
test('비균등 배율과 자식 회전은 전단 변환을 잃지 않게 보존한다',()=>{
 const g=makeObjectGroup([{...objects[0],rot:30},objects[1]],'g');g.w*=2;const parts=ungroupObjects(g);assert.equal(parts[0].kind,'group');assert.equal(parts[0].groupItems[0].rot,30);assert.equal(parts[0].w,g.w);
});
test('중첩 그룹과 그림의 자르기·색·효과·텍스트가 SVG에 유지된다',()=>{
 const g=makeObjectGroup(objects,'g'),pic={id:'p',kind:'picture',x:10,y:20,w:100,h:60,src:'data:image/png;base64,a',opacity:.7,radius:4,shadow:true,border:'#123456',borderW:2,crop:{l:.2},flip:true};
 const outer=makeObjectGroup([g,pic],'outer');outer.flip=true;const svg=shapeSvg(outer);assert.match(svg,/한글/);assert.match(svg,/aria-label="그룹"/);assert.match(svg,/opacity="0.7"/);assert.match(svg,/stroke="#123456"/);assert.match(svg,/feDropShadow/);assert.match(svg,/scale\(-1,1\)/);
});
