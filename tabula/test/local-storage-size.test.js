import test from 'node:test';
import assert from 'node:assert/strict';
import { isLargeLocalWorkbook } from '../src/local-storage-size.js';
const book = (rows, sheet = {}) => ({ sheets: [{ cells: { size: 1 }, ...sheet }], pivotSnapshots: rows ? new Map([['cache', { rows }]]) : null });
const compact = columns => ({ kind: 'pivot-cache', version: 1, header: columns.map((_,i) => 'C' + i), n: columns[0]?.num?.length || 1, columns });
test('small books remain small while ordinary cell and block limits are preserved', () => {
 assert.equal(isLargeLocalWorkbook(book()),false);
 assert.equal(isLargeLocalWorkbook(book(null,{cells:{size:50000}})),false);
 assert.equal(isLargeLocalWorkbook(book(null,{cells:{size:50001}})),true);
 assert.equal(isLargeLocalWorkbook(book(null,{blocks:[{n:25001,cols:[{},{}]}]})),true);
 assert.equal(isLargeLocalWorkbook(book(compact([{num:new Float64Array([1,2]),str:null,dict:[]}]))),false);
});
test('large cached typed arrays route even absent source cells to IDB without JSON serialization', () => {
 const rows=compact([{num:new Float64Array(300000),str:null,dict:[]}]);rows.toJSON=()=>{throw Error('must not stringify');};
 assert.equal(isLargeLocalWorkbook(book(rows,{cells:{size:0}})),true);
 // Cache hits must not scan the same immutable columns again.
 Object.defineProperty(rows,'columns',{get(){throw Error('must use memo');}});
 assert.equal(isLargeLocalWorkbook(book(rows)),true);
});
test('typed-array JSON growth margin and long dictionaries are both included', () => {
 assert.equal(isLargeLocalWorkbook(book(compact([{num:new Float64Array(65000),str:null,dict:[]}]))),true);
 assert.equal(isLargeLocalWorkbook(book(compact([{num:null,str:new Int32Array(1),dict:['x'.repeat(2000001)]}]))),true);
 assert.equal(isLargeLocalWorkbook(book(compact([{num:null,str:new Int32Array([0,1]),dict:[true,{error:'#N/A'}]}]))),false);
});
test('legacy cached rows use an early bounded scan and do not materialize JSON', () => {
 const rows=[['C'],['x'.repeat(2000001)]];Object.defineProperty(rows,2,{get(){throw Error('must stop early');}});
 assert.equal(isLargeLocalWorkbook(book(rows)),true);
 assert.equal(isLargeLocalWorkbook(book([['C'],[123],[true],[null],[{error:'#N/A'}]])),false);
});
test('several small snapshots combine and replacing a snapshot reclassifies the book', () => {
 const a=[['C'],['x'.repeat(1100000)]],b=[['C'],['y'.repeat(1100000)]],w=book(a);
 assert.equal(isLargeLocalWorkbook(w),false);w.pivotSnapshots.set('other',{rows:b});assert.equal(isLargeLocalWorkbook(w),true);
 w.pivotSnapshots=new Map([['cache',{rows:[['C'],['small']]}]]);assert.equal(isLargeLocalWorkbook(w),false);
});
test('grouped pictures, original assets and chart picture fills combine across sheets', () => {
 const w=book(null,{shapes:[{kind:'group',groupItems:[{kind:'group',groupItems:[{kind:'picture',src:'x'.repeat(1100000),originalSrc:'y'.repeat(1100000)}]}]}]});
 assert.equal(isLargeLocalWorkbook(w),true);
 assert.equal(isLargeLocalWorkbook(book(null,{charts:[{chartAreaFormat:{picture:{src:'x'.repeat(2000001)}}}]})),true);
 assert.equal(isLargeLocalWorkbook(book(null,{images:[{src:'https://example.test/image.png'}]})),false);
 const image={src:'x'.repeat(1100000)};
 assert.equal(isLargeLocalWorkbook(book(null,{shapes:[{groupItems:[image,image]}]})),true,'serialized duplicate references count twice');
});
test('mutable picture edits are rechecked and cached standalone slicers are counted', () => {
 const image={src:'tiny'},w=book(null,{images:[image]});assert.equal(isLargeLocalWorkbook(w),false);
 image.src='x'.repeat(2000001);assert.equal(isLargeLocalWorkbook(w),true);image.src='tiny';assert.equal(isLargeLocalWorkbook(w),false);
 assert.equal(isLargeLocalWorkbook(book(null,{slicers:[{source:{kind:'cache',values:['x'.repeat(2000001)]}}]})),true);
});


test('valid long cell text uses IDB despite a small cell count and scanning stops at the threshold', async () => {
 const {Workbook}=await import('../src/workbook.js');
 const cells={};for(let r=0;r<200;r++)cells[`${r},0`]={raw:'x'.repeat(32767)};
 const wb=new Workbook({sheets:[{name:'Data',cells}]});
 assert.ok(JSON.stringify(wb.serialize()).length>6e6);
 assert.equal(isLargeLocalWorkbook(wb),true);
 const payload='x'.repeat(1100000);let visited=0;
 const w=book(null,{cells:{size:3,*storageEntries(){visited++;yield[0,0,{raw:payload},1];visited++;yield[1,0,{raw:payload},1];throw Error('must stop before third cell');}}});
 assert.equal(isLargeLocalWorkbook(w),true);assert.equal(visited,2);
});

test('cell notes, hyperlinks, pictures and formula cached results are serialized payload',()=>{
 for(const cell of [{raw:'x',comment:'n'.repeat(2000001)},{raw:'x',link:'h'.repeat(2000001)},
  {raw:'',image:{src:'p'.repeat(2000001)}},{raw:'=X()',formula:true,cached:'c'.repeat(2000001)}]) {
  const w=book(null,{cells:{size:1,*storageEntries(){yield[0,0,cell,1];}}});assert.equal(isLargeLocalWorkbook(w),true);
 }
 const small=book(null,{cells:{size:1,*storageEntries(){yield[0,0,{raw:'=1',formula:true,ast:{text:'x'.repeat(2000001)},v:'y'.repeat(2000001)},1];}}});
 assert.equal(isLargeLocalWorkbook(small),false,'runtime AST and derived value are not saved');
});

test('compressed blank ranges are inspected physically and shared styles count at every stored position',()=>{
 let visits=0;const w=book(null,{cells:{size:40000,*storageEntries(){visits++;yield[0,0,{raw:'',style:{bold:true}},40000];},*[Symbol.iterator](){throw Error('must not expand a run');}}});
 assert.equal(isLargeLocalWorkbook(w),false);assert.equal(visits,1);
 const style={font:'x'.repeat(1100000)};
 const repeated=book(null,{cells:{size:2,*storageEntries(){yield[0,0,{raw:'a',style},1];yield[1,0,{raw:'b',style},1];}}});
 assert.equal(isLargeLocalWorkbook(repeated),true,'shared references are written at both coordinates');
});

test('book and sheet metadata, names and block dictionaries count without copying or parsing caches',()=>{
 for(const extra of [{vba:'v'.repeat(2000001)},{props:{comments:'c'.repeat(2000001)}},{names:[{name:'N',ref:'r'.repeat(2000001)}]},
  {cellStyles:[{name:'S',style:{font:'f'.repeat(2000001)}}]}])assert.equal(isLargeLocalWorkbook({...book(),...extra}),true);
 assert.equal(isLargeLocalWorkbook(book(null,{name:'n'.repeat(2000001)})),true);
 assert.equal(isLargeLocalWorkbook(book(null,{blocks:[{n:1,cols:[{num:null,str:new Int32Array([0]),dict:['d'.repeat(2000001)],fmt:null}]}]})),true);
 assert.equal(isLargeLocalWorkbook({...book(),names:[{name:'N',ref:'=1',_ast:{text:'a'.repeat(2000001)},_text:'b'.repeat(2000001)}],graph:{text:'c'.repeat(2000001)}}),false);
 const props={comments:'short'},mutable={...book(),props};assert.equal(isLargeLocalWorkbook(mutable),false);
 props.comments='x'.repeat(2000001);assert.equal(isLargeLocalWorkbook(mutable),true);
});

test('JSON escape expansion is included in literal cells and saved pivot cache strings',()=>{
 const escaped='\u0001'.repeat(350000);
 assert.ok(JSON.stringify(escaped).length>2e6);
 const w=book(null,{cells:{size:1,*storageEntries(){yield[0,0,{raw:escaped},1];}}});
 assert.equal(isLargeLocalWorkbook(w),true);
 assert.equal(isLargeLocalWorkbook(book([['C'],[escaped]])),true);
 assert.equal(isLargeLocalWorkbook(book([['C'],['한글😀']])),false);
});

test('one array anchor with large cached spill values routes to IDB before JSON encoding',()=>{
 const cachedArray={h:2,w:1,values:[0,0,'small',1,0,'x'.repeat(2000001)]};
 cachedArray.toJSON=()=>{throw Error('must not stringify array cache');};
 const cell={raw:'=NOSUCHARRAY()',formula:true,cachedArray};
 const w=book(null,{cells:new Map([['0,0',cell]])});
 assert.equal(isLargeLocalWorkbook(w),true);
 cell.dirty=true;assert.equal(isLargeLocalWorkbook(w),true);
 cell.cachedArray={h:2,w:1,values:[0,0,10,1,0,20]};assert.equal(isLargeLocalWorkbook(w),false);
});


test('historical pivot items alone route to IDB and stop at the payload threshold',()=>{
 const shared=['한글😀','x'.repeat(2000001)];
 Object.defineProperty(shared,2,{get(){throw Error('must stop before later cache items');}});
 const pivotCacheItems={cache:{fields:[{name:'분류',shared}]}};
 pivotCacheItems.toJSON=()=>{throw Error('must not serialize pivot metadata');};
 assert.equal(isLargeLocalWorkbook({...book(),pivotCacheItems}),true);
 const small={cache:{fields:[{name:'분류',shared:['현재','과거',null,true,{error:'#N/A'}]}]}};
 assert.equal(isLargeLocalWorkbook({...book(),pivotCacheItems:small}),false);
 const combined={first:{fields:[{name:'A',shared:['x'.repeat(1100000)]}]},second:{fields:[{name:'B',shared:['y'.repeat(1100000)]}]}};
 assert.equal(isLargeLocalWorkbook({...book(),pivotCacheItems:combined}),true);
});
