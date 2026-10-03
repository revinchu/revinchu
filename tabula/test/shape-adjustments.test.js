import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { GEOM, shapeSvg } from '../src/shapes.js';
import { mergeShapes } from '../src/shape-boolean.js';
const shape={id:'card',kind:'roundRect',x:10,y:20,w:300,h:100,fill:'#ffffff',stroke:null};
const make=s=>new Workbook({sheets:[{name:'도형',cells:{},shapes:[s]}]});
const drawing='xl/drawings/drawing1.xml';
const native=(adj)=>{const files=unzip(writeXlsx(make(shape)));files[drawing]=textOf(files[drawing]).replace('<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>',`<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val ${adj}"/></a:avLst></a:prstGeom>`);return readXlsx(zip(files)).data.sheets[0].shapes[0];};

test('Excel 카드의 1746·9661 조정값은 기본 곡률로 바뀌지 않고 표준 XML로 왕복한다',()=>{
  for(const adj of [1746,9661,0,50000]){
    const imported=native(adj);assert.deepEqual(imported.adjustments,{adj});
    const wb=make(imported),before=wb.serialize(),bytes=writeXlsx(wb),xml=textOf(unzip(bytes)[drawing]);
    assert.match(xml,new RegExp(`<a:gd name="adj" fmla="val ${adj}"/>`));
    assert.equal(readXlsx(bytes).data.sheets[0].shapes[0].adjustments.adj,adj);
    assert.deepEqual(wb.serialize(),before);
  }
  assert.match(shapeSvg(native(1746)),/A1\.75,1\.75/);
  assert.match(shapeSvg(native(9661)),/A9\.66,9\.66/);
});

test('모서리 반지름은 짧은 변에 비례하며 0·반원 한계와 기본값을 처리한다',()=>{
  assert.match(GEOM.roundRect(500,120,{adjustments:{adj:9661}})[0].d,/A11\.59,11\.59/);
  assert.match(GEOM.roundRect(240,1000,{adjustments:{adj:9661}})[0].d,/A23\.19,23\.19/);
  assert.match(GEOM.roundRect(300,100)[0].d,/A16\.67,16\.67/);
  assert.doesNotMatch(shapeSvg(native(-500)),/\bA/);
  assert.match(shapeSvg(native(75000)),/A50,50/);
  for(const adj of [NaN,Infinity,'1746']){
    const s={...shape,adjustments:{adj}},xml=textOf(unzip(writeXlsx(make(s)))[drawing]);
    assert.doesNotMatch(xml,/<a:gd name="adj"/);assert.match(shapeSvg(s),/A16\.67,16\.67/);
  }
});

test('둥근 모서리 도형 조합도 화면과 같은 조정값의 면적을 사용한다',()=>{
  const card={...shape,x:0,y:0,w:100,h:100,adjustments:{adj:1746}},corner={id:'corner',kind:'rect',x:.75,y:.75,w:.5,h:.5,fill:'#ffffff'};
  const before=structuredClone(card);
  assert.equal(mergeShapes([card,corner],'intersect').length,1);
  assert.equal(mergeShapes([{...card,adjustments:undefined},corner],'intersect').length,0);
  assert.deepEqual(card,before);
});

test('그룹 안 도형의 조정값과 실행 취소·다시 실행 상태를 보존한다',()=>{
  const card=native(1746),group={id:'group',kind:'group',x:0,y:0,w:400,h:200,groupSize:{w:400,h:200},groupItems:[card]};
  const back=readXlsx(writeXlsx(make(group))).data.sheets[0].shapes[0];
  assert.equal(back.groupItems[0].adjustments.adj,1746);
  const wb=make(card);wb.transact(()=>wb.setSheetProp(0,'shapes',[{...card,adjustments:{adj:9661}}]));
  wb.undo();assert.equal(wb.sheets[0].shapes[0].adjustments.adj,1746);
  wb.redo();assert.equal(wb.sheets[0].shapes[0].adjustments.adj,9661);
});
