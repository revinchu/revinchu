import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { CellMap } from '../src/cellmap.js';
import { readXlsx, readXlsxAsync, writeXlsx } from '../src/xlsx.js';

function saturationBook(groups) {
  const first=new CellMap(),second=new CellMap(),styles=Array.from({length:groups},(_,i)=>({fill:'#'+(0xaabb00+i).toString(16)}));
  for(let c=0;c<groups;c++)for(let r=0;r<513;r++)first.setRC(r,c,{raw:`항목 ${c} ${r}`,style:styles[c]});
  second.setRC(0,0,{raw:'후속 반복',style:styles[0]});second.setRC(1,0,{raw:'후속 반복',style:styles[0]});
  return {sheets:[{name:'앞 시트',cells:first},{name:'뒤 시트',cells:second}]};
}
for(const groups of [1,21])test(`restore sharing budgets restart per sheet after ${groups===1?'512 style':'10000 global'} entries`,()=>{
  const w=new Workbook(saturationBook(groups));
  assert.equal(w.getCell(1,0,0),w.getCell(1,1,0));assert.ok(Object.isFrozen(w.getCell(1,0,0)));
  w.transact(()=>w.setComment(1,0,0,'이 셀만'));assert.equal(w.getCell(1,1,0).comment,undefined);w.undo();
  assert.equal(w.getRaw(1,0,0),'후속 반복');assert.equal(w.getRaw(0,0,0),'항목 0 0');
});
const saturatedBytes=writeXlsx(new Workbook(saturationBook(21)));
for(const async of [false,true])test(`XLSX literal sharing budget restarts on the next sheet: async=${async}`,async()=>{
  const {data}=await (async?readXlsxAsync(saturatedBytes):readXlsx(saturatedBytes));
  const cells=data.sheets[1].cells;
  assert.equal(cells.getRC(0,0),cells.getRC(1,0));assert.ok(Object.isFrozen(cells.getRC(0,0)));
  const w=new Workbook(data);assert.equal(w.getCell(1,0,0),w.getCell(1,1,0));
  w.transact(()=>w.setInput(1,0,0,'바뀜'));assert.equal(w.getRaw(1,1,0),'후속 반복');w.undo();assert.equal(w.getRaw(1,0,0),'후속 반복');
});
test('the same frozen imported literal is prepared once per sheet',()=>{
  const style={fill:'#aabbcc'};let preparations=0;
  const value=Object.freeze({raw:'반복',get style(){preparations++;return style;}});
  const sheets=[];for(let si=0;si<2;si++){const cells=new CellMap();for(let r=0;r<100;r++)cells.setRC(r,0,value);sheets.push({name:`시트${si}`,cells});}
  const w=new Workbook({sheets});assert.equal(preparations,2);
  assert.equal(w.getCell(0,0,0),w.getCell(0,99,0));assert.notEqual(w.getCell(0,0,0),w.getCell(1,0,0));
});
test('frozen date, formula and annotated inputs never reuse a prepared mutable cell',()=>{
  const cells=new CellMap();
  for(const [c,value] of [
    [0,Object.freeze({raw:'2024-01-01'})],
    [1,Object.freeze({raw:'=ROW()'})],
    [2,Object.freeze({raw:'메모',comment:'설명'})],
    [3,Object.freeze({raw:'링크',link:'#A1'})],
    [4,Object.freeze({raw:'한자',phonetic:{text:'윗주'}})],
  ])for(let r=0;r<2;r++)cells.setRC(r,c,value);
  const w=new Workbook({sheets:[{name:'독립',cells}]});
  for(let c=0;c<5;c++){assert.notEqual(w.getCell(0,0,c),w.getCell(0,1,c));assert.ok(!Object.isFrozen(w.getCell(0,0,c)));}
  assert.equal(w.getValue(0,0,1),1);assert.equal(w.getValue(0,1,1),2);
  const serial=w.getValue(0,0,0);w.transact(()=>w.setDate1904(true));assert.equal(w.getValue(0,0,0),serial);assert.equal(w.getValue(0,1,0),serial);w.undo();
  assert.equal(w.getRaw(0,0,0),'2024-01-01');assert.equal(w.getRaw(0,1,0),'2024-01-01');
  w.transact(()=>w.setComment(0,0,2,'다른 메모'));assert.equal(w.getCell(0,1,2).comment,'설명');w.undo();
});
test('ordinary object restore keeps independent cells and cloned metadata',()=>{
  const value=Object.freeze({raw:'반복',style:{fill:'#aabbcc'}}),rowStyles={0:{bold:true}};
  const input={sheets:[{name:'일반',cells:{'0,0':value,'1,0':value},rowStyles}]};
  const w=new Workbook(input);assert.notEqual(w.getCell(0,0,0),w.getCell(0,1,0));
  assert.equal(input.sheets[0].cells['0,0'],value);assert.notEqual(w.sheets[0].rowStyles,rowStyles);
  w.transact(()=>w.setLineStyle(0,'row',0,{bold:false}));assert.equal(rowStyles[0].bold,true);
});
