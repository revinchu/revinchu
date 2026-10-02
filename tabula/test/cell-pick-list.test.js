import {test} from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../src/workbook.js';
import {cellPickList} from '../src/cell-pick-list.js';
const make=rows=>{const b=new Workbook();b.transact(()=>rows.forEach((v,r)=>{if(v!==null)b.setInput(0,r,0,v);}));return b;};
test('column pick lists include literal neighbours once, exclude numeric/formula values, and stop at blank boundary',()=>{
 const b=make(['outside',null,'배','사과','사과','123',null,'감','=1+1','포도',null,'far']);
 assert.deepEqual(cellPickList(b,0,6,0).items,['감','배','사과','포도']);
});
test('column suggestions preserve literal leading equals and case-insensitive uniqueness without edits',()=>{
 const b=make(["'=SUM(A1)",'Apple','apple',null]);const before=JSON.stringify(b.serialize()),undo=b.undoStack.length;
 assert.deepEqual(cellPickList(b,0,3,0).items,['=SUM(A1)','apple']);assert.equal(JSON.stringify(b.serialize()),before);assert.equal(b.undoStack.length,undo);
});
test('column suggestion work and output are bounded for long lists',()=>{
 const b=make(Array.from({length:100},(_,r)=>'값'+r));const x=cellPickList(b,0,100,0,{maxScan:8,maxItems:3});assert.equal(x.items.length,3);assert.equal(x.limited,true);
});
