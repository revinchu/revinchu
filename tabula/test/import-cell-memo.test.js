import test from 'node:test';
import assert from 'node:assert/strict';
import { createImportedLiteralMemo } from '../src/import-cell-memo.js';
import { Workbook } from '../src/workbook.js';
import { writeXlsx,readXlsx,readXlsxAsync } from '../src/xlsx.js';

test('import literal sharing is exact, bounded and excludes coordinate-specific data',()=>{
 const memo=createImportedLiteralMemo({limit:3,perStyleLimit:2}),style={size:12},a=memo({raw:'반복',style});
 assert.equal(memo({raw:'반복',style}),a);assert.ok(Object.isFrozen(a));
 assert.notEqual(memo({raw:'반복',style:{size:12}}),a);
 for(const extra of [{comment:'note'},{link:'#A1'},{image:{src:'data:'}},{cached:1},{fx:true},{phonetic:{}}]){
  const cell={raw:'반복',style,...extra};assert.equal(memo(cell),cell);
 }
 const formula={raw:'=1+1',style};assert.equal(memo(formula),formula);
 const b=memo({raw:'다른 값',style});assert.equal(memo({raw:'다른 값',style}),b);
 const over={raw:'초과',style};assert.equal(memo(over),over);assert.notEqual(memo({...over}),over);
 const other=createImportedLiteralMemo();assert.notEqual(other({raw:'반복',style}),a);
});

for(const async of [false,true])test(`read literals remain independent when annotated, restored, edited and date system changes: async=${async}`,async()=>{
 const cells={};for(let r=0;r<100;r++)cells[r+',0']={raw:'반복',style:{fill:'#123456'}};
 cells['0,0'].comment='첫 셀 메모';cells['1,0'].link='#B1';
 cells['0,1']={raw:'3',style:{numFmt:'text'},inputType:'value'};cells['1,1']={raw:'3',style:{numFmt:'text'}};
 cells['0,2']={raw:'2024-01-01'};cells['1,2']={raw:'2024-01-01'};
 cells['0,3']={raw:'=1+1'};cells['1,3']={raw:'=1+1'};
 const bytes=writeXlsx(new Workbook({sheets:[{name:'Data',cells}]}));
 const {data}=await (async?readXlsxAsync(bytes):readXlsx(bytes));const imported=data.sheets[0].cells;
 assert.equal(imported.getRC(2,0),imported.getRC(99,0));assert.ok(Object.isFrozen(imported.getRC(2,0)));
 assert.equal(imported.getRC(0,0).comment,'첫 셀 메모');assert.equal(imported.getRC(2,0).comment,undefined);
 assert.equal(imported.getRC(1,0).link,'#B1');assert.equal(imported.getRC(2,0).link,undefined);
 assert.notEqual(imported.getRC(0,3),imported.getRC(1,3));
 const wb=new Workbook(data);assert.equal(wb.getValue(0,0,1),3);assert.equal(wb.getValue(0,1,1),'3');
 wb.transact(()=>wb.setInput(0,2,0,'수정'));assert.equal(wb.getRaw(0,99,0),'반복');wb.undo();assert.equal(wb.getRaw(0,2,0),'반복');
 wb.transact(()=>wb.setDate1904(true));assert.equal(wb.getValue(0,0,2),wb.getValue(0,1,2));
 const back=new Workbook(readXlsx(writeXlsx(wb)).data);assert.equal(back.getRaw(0,99,0),'반복');assert.equal(back.getCell(0,0,0).comment,'첫 셀 메모');
});
