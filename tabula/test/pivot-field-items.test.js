import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { cubeFromRows, itemIdentity, EMPTY, EMPTY_TEXT } from '../src/cube.js';
import { filterSelection } from '../src/filter-selection.js';
import { pivotFieldItemModel, pivotItemSelection, slicerPickValues } from '../src/pivot-field-items.js';
const ids = n => Array.from({length:n},(_,i)=>'item'+String(i).padStart(3,'0'));
function fixture({history=275,selected=153}={}) {
 const wb=new Workbook({sheets:[{name:'Source',cells:{}}]});
 const source={cube:cubeFromRows([['Item','Value'],...ids(31).map((x,i)=>[x,i])]),si:0,ref:{r1:0,c1:0,r2:31,c2:1}};
 const def={source:'Source',range:source.ref,rows:['Item'],filters:{Item:ids(selected)},cacheItemsId:'cache'};
 if(history)wb.pivotCacheItems={cache:{fields:[{name:'Other',shared:['unrelated']},{name:'ITEM',shared:ids(history)}]}};
 return{wb,def,source};
}
const checklist=model=>filterSelection(model.items.map(i=>i.v),new Set(model.items.filter(i=>i.selected).map(i=>i.v)),v=>model.items.find(i=>i.v===v)?.text??v);
const choice=(model,state)=>pivotItemSelection(model,state.result().map(itemIdentity));
for(const history of[275,0])test('31 current / 153 selected: no-change confirm preserves original with history '+history,()=>{
 const{wb,def,source}=fixture({history}),before=JSON.stringify(wb.pivotCacheItems),model=pivotFieldItemModel(wb,def,'item',{source});
 assert.equal(model.items.length,history||153);assert.equal(model.selected.size,153);
 const result=choice(model,checklist(model));assert.equal(result.unchanged,true);assert.equal(result.values,def.filters.Item);
 assert.equal(JSON.stringify(wb.pivotCacheItems),before);assert.ok(!model.items.some(i=>i.key==='unrelated'));
});

test('search add selection retains absent history; explicit replace and clear remain intentional',()=>{
 const{wb,def,source}=fixture(),model=pivotFieldItemModel(wb,def,'Item',{source}),state=checklist(model);
 state.search('item200');const add=pivotItemSelection(model,state.result(true).map(itemIdentity));
 assert.equal(add.values.length,154);assert.ok(add.values.includes('item152'));assert.ok(add.values.includes('item200'));
 const replace=choice(model,state);assert.deepEqual(replace.values,['item200']);assert.equal(replace.unchanged,false);
 assert.deepEqual(pivotItemSelection(model,null),{unchanged:false,values:null});
});

test('Ctrl selection preserves 122 hidden selected items; direct single pick replaces them',()=>{
 const{wb,def,source}=fixture(),model=pivotFieldItemModel(wb,def,'Item',{source});
 const all=model.items.map(i=>({...i,key:i.id})),slicer={allItems:all,items:all.filter(i=>!i.deleted),filtered:true,preserveSelection:true};
 const result=pivotItemSelection(model,slicerPickValues(slicer,'string:item000',true));
 assert.equal(result.values.length,152);assert.ok(result.values.includes('item152'));assert.ok(!result.values.includes('item000'));
 assert.deepEqual(pivotItemSelection(model,slicerPickValues(slicer,'string:item000',false)).values,['item000']);
 assert.equal(model.items.filter(i=>i.deleted).length,244);assert.equal(all.filter(i=>i.selected&&i.deleted).length,122);
});

test('typed UI identities do not merge numeric/text/blank values or mutate filter text semantics',()=>{
 const wb=new Workbook(),source={cube:cubeFromRows([['Mixed'],[1],['1'],[null],[''],[true],['TRUE']])};
 const def={cacheItemsId:'x',filters:{Mixed:['1','(비어 있음)','', 'true']}};
 wb.pivotCacheItems={x:{fields:[{name:'Mixed',shared:[1,'1',null,'',true,'TRUE']}]}};
 const m=pivotFieldItemModel(wb,def,'Mixed',{source});assert.equal(m.items.length,6);
 assert.ok(m.items.some(i=>i.id==='number:1'));assert.ok(m.items.some(i=>i.id==='string:1'));
 assert.ok(m.items.some(i=>i.v===EMPTY));assert.ok(m.items.some(i=>i.v===EMPTY_TEXT));
 assert.equal(pivotItemSelection(m,m.selected).values,def.filters.Mixed);
});

for(const date1904 of[false,true])test('date metadata labels retained and current values with date1904='+date1904,()=>{
 const wb=new Workbook();wb.date1904=date1904;
 const serial=date1904?43890:45352,source={cube:cubeFromRows([['Date'],[serial]])};
 wb.pivotCacheItems={x:{fields:[{name:'Date',shared:[serial,serial+1],sharedTypes:'dd',format:{numFmt:'custom',code:'yyyy-mm-dd'}}]}};
 const m=pivotFieldItemModel(wb,{cacheItemsId:'x',filters:{Date:[String(serial),String(serial+1)]}},'Date',{source});
 assert.equal(m.items[0].text,'2024-03-01');assert.equal(m.items[1].text,'2024-03-02');assert.equal(m.items[1].deleted,true);
});

test('derived field uses matching grouped history or projects base history, never raw serial labels',()=>{
 const wb=new Workbook(),source={cube:cubeFromRows([['Date'],[45352],[45383]])};
 const def={cacheItemsId:'x',groups:{Month:{base:'Date',by:'months'}},filters:{Month:['3월','2월']}};
 wb.pivotCacheItems={x:{fields:[{name:'Date',shared:[45323,45352,45383]}]}};
 let m=pivotFieldItemModel(wb,def,'Month',{source});assert.deepEqual(new Set(m.items.map(i=>i.key)),new Set(['2월','3월','4월']));assert.equal(m.items.find(i=>i.key==='2월').deleted,true);
 wb.pivotCacheItems={x:{fields:[{name:'Date',shared:[45000]},{name:'Month',shared:['1월','3월','4월']}]}};
 m=pivotFieldItemModel(wb,def,'Month',{source});assert.deepEqual(new Set(m.items.map(i=>i.key)),new Set(['1월','2월','3월','4월']));
 assert.ok(!m.items.some(i=>/45000/.test(i.key)));assert.equal(m.items.find(i=>i.key==='2월').unknown,true);
});

test('none retention policy keeps saved none history until refresh, then preserves only current and selected',()=>{
 const{wb,def,source}=fixture();wb.pivotCacheItems.cache.missingItemsLimit=0;def.missingItems='none';def.snapshotId='saved';
 wb.pivotSnapshots=new Map([['saved',{rows:[]}]]);wb.pivotSnapshotCurrent=()=>true;
 assert.equal(pivotFieldItemModel(wb,def,'Item',{source}).items.length,275);
 wb.pivotSnapshots=null;const m=pivotFieldItemModel(wb,def,'Item',{source});assert.equal(m.items.length,153);assert.equal(pivotItemSelection(m,m.selected).values,def.filters.Item);
});

test('date-typed cache without numFmtId has a date label, including after none policy',()=>{
 const wb=new Workbook(),source={cube:cubeFromRows([['Date'],[45352]])};
 wb.pivotCacheItems={x:{fields:[{name:'Date',shared:[45352,45353],sharedTypes:'dd'}]}};
 for(const missingItems of[undefined,'none']){const m=pivotFieldItemModel(wb,{cacheItemsId:'x',missingItems},'Date',{source});assert.equal(m.items[0].text,'2024-03-01');}
});


test('mixed dates retain date labels after none policy without converting ordinary numbers',()=>{
 const wb=new Workbook(),source={cube:cubeFromRows([['Mixed'],[45352],[7],['']])};
 wb.pivotCacheItems={x:{fields:[{name:'Mixed',shared:[45352,7,''],sharedTypes:'dns'}]}};
 const m=pivotFieldItemModel(wb,{cacheItemsId:'x',missingItems:'none'},'Mixed',{source});
 assert.equal(m.items.find(i=>i.v===45352).text,'2024-03-01');assert.equal(m.items.find(i=>i.v===7).text,'7');
});

test('renamed and empty item captions change labels without changing persisted selection',()=>{
 const{wb,def,source}=fixture();def.itemCaptions={Item:{item000:'Renamed',item200:''}};
 const m=pivotFieldItemModel(wb,def,'Item',{source});
 assert.equal(m.items.find(i=>i.key==='item000').text,'Renamed');assert.equal(m.items.find(i=>i.key==='item200').text,'');
 assert.equal(pivotItemSelection(m,m.selected).values,def.filters.Item);
});
