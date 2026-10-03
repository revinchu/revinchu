import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';

test('named object styles keep workbook metadata and formula caches through undo, redo and JSON restore', () => {
  const w=new Workbook({sheets:[{name:'원본',fileValues:true,cells:{'0,0':{raw:'=2+2',cached:42}}}]});
  const styles={tables:[{name:'보고서 표',table:true,pivot:false,elements:[{type:'headerRow',style:{fill:'#14734c',bold:true}}]}],slicers:[],defaultTableStyle:'보고서 표'};
  w.transact(()=>w.setObjectStyles(styles));
  assert.equal(w.undoStack.length,1);assert.equal(w.getValue(0,0,0),42);
  styles.tables[0].elements[0].style.fill='#ff0000';
  assert.equal(w.objectStyles.tables[0].elements[0].style.fill,'#14734c');
  const copy=new Workbook(w.serialize());assert.deepEqual(copy.objectStyles,w.objectStyles);assert.equal(copy.getValue(0,0,0),42);
  w.undo();assert.equal(w.objectStyles,null);assert.equal(w.getValue(0,0,0),42);
  w.redo();assert.equal(w.objectStyles.defaultTableStyle,'보고서 표');assert.equal(w.getValue(0,0,0),42);
  w.transact(()=>w.setObjectStyles(w.objectStyles));assert.equal(w.undoStack.length,1);
});

test('object style definition and uses are restored in the same undo transaction', () => {
  const w=new Workbook({sheets:[{name:'시트',cells:{},slicers:[{id:'s',style:'SlicerStyleLight1'}]}]});
  const def={name:'내 슬라이서',elements:[{type:'selectedItemWithData',style:{fill:'#001122'}}]};
  w.transact(()=>{w.setObjectStyles({tables:[],slicers:[def]});w.setSheetProp(0,'slicers',[{id:'s',style:def.name,styleElements:def.elements}]);});
  assert.equal(w.undoStack.length,1);w.undo();assert.equal(w.objectStyles,null);assert.equal(w.sheets[0].slicers[0].style,'SlicerStyleLight1');
  w.redo();assert.equal(w.sheets[0].slicers[0].style,w.objectStyles.slicers[0].name);
});


test('table style-only updates preserve imported formula values, while later structural changes still recalculate',()=>{
 const w=new Workbook({sheets:[{name:'자료',fileValues:true,cells:{'0,0':{raw:'제목'},'1,0':{raw:'=2+2',cached:42}},tables:[{id:'t',name:'표1',r1:0,c1:0,r2:1,c2:0,header:true,style:'TableStyleMedium2'}]}]});
 w.transact(()=>w.setTableStyle(0,'t',{style:'None'}));assert.equal(w.getValue(0,1,0),42);
 w.undo();assert.equal(w.getValue(0,1,0),42);assert.equal(w.sheets[0].tables[0].style,'TableStyleMedium2');
 w.redo();assert.equal(w.getValue(0,1,0),42);
 assert.throws(()=>w.setTableStyle(0,'t',{r2:9}),/서식 외/);
 w.transact(()=>{w.setTableStyle(0,'t',{style:'TableStyleLight1'});w.setSheetProp(0,'tables',w.sheets[0].tables.map(t=>({...t,r2:3})));});
 assert.equal(w.getValue(0,1,0),4);assert.equal(w.undoStack.at(-1).entries.find(e=>e.prop==='tables').calcNeutral,undefined);
});
