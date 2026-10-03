import test from 'node:test';
import assert from 'node:assert/strict';
import { pivotExportData,pivotValueStats } from '../src/pivot-export-data.js';
import { Cube } from '../src/cube.js';
import { Workbook } from '../src/workbook.js';
import { PivotSnapshotBuilder } from '../src/pivot-cache-data.js';
import { pivotSourceData } from '../src/pivot.js';
import { writeXlsx,writeXlsxAsync,readXlsx } from '../src/xlsx.js';
import { unzip,textOf } from '../src/zip.js';

test('saved pivot records use indexed columns without reading row arrays or allocating a million records',()=>{
 let gets=0;const cube={n:1000000,header:['a','b'],col(c){return{get(r){gets++;return c?r:r%3;}}},row(){throw Error('row materialization')}};
 const data=pivotExportData(cube,true);assert.equal(gets,0);assert.equal(data.length,1000000);assert.equal(data.value(999999,1),999999);
 const values=data.values(0)[Symbol.iterator]();assert.deepEqual([values.next().value,values.next().value],[0,1]);assert.equal(gets,3);
});

test('non-saved cache drops all-blank rows while preserving false, zero, order and reusable value iterators',()=>{
 const rows=[[null,''],[0,false],[null,null],['',2],[3,'x']];const cube={n:rows.length,header:['a','b'],col(c){return{get:r=>rows[r][c]}}};
 const data=pivotExportData(cube);assert.equal(data.length,3);assert.deepEqual([...data.values(0)],[0,'',3]);assert.deepEqual([...data.values(1)],[false,2,'x']);assert.deepEqual([...data.values(1,r=>data.value(r,0)!=='')],[false,'x']);
 assert.deepEqual(pivotValueStats(data.values(0)),{count:3,numbers:2,blanks:1,missing:0,emptyStrings:1,hasString:false,integers:true,min:0,max:3});
 assert.equal(pivotExportData(cube,true).length,5);assert.equal(pivotValueStats([1,2.5,false,{error:'#N/A'}]).integers,false);
});

for(const async of [false,true])test(`XLSX saved pivot export never materializes source rows, async=${async}`,async()=>{
 const b=new PivotSnapshotBuilder(['분류','채널','값'],5000);
 for(let r=0;r<5000;r++)b.add([r===13?'':r===14?null:'항목'+r%3,r%2?'A':'B',r]);
 const wb=new Workbook({pivotSnapshots:{saved:b.finish()},sheets:[{name:'Source',cells:{'0,0':{raw:'분류'},'0,1':{raw:'채널'},'0,2':{raw:'값'},'1,0':{raw:'different'},'1,1':{raw:'A'},'1,2':{raw:'99999'}}},{name:'Report',cells:{},pivot:{name:'Pivot',source:'Source',range:{r1:0,c1:0,r2:1,c2:2},snapshotId:'saved',rows:['분류'],pages:['채널'],values:[{field:'값',agg:'sum'}],filters:{채널:['A']},top:3,left:0},slicers:[{id:'s',x:20,y:30,w:140,h:200,source:{kind:'pivot',field:'분류',pivots:[{sheet:'Report',name:'Pivot'}]}}]}]});
 const row=Cube.prototype.row;let bytes;
 try{Cube.prototype.row=()=>{throw Error('full pivot row materialization')};bytes=async?await writeXlsxAsync(wb):writeXlsx(wb);}finally{Cube.prototype.row=row;}
 const files=unzip(bytes),cache=textOf(files['xl/pivotCache/pivotCacheDefinition1.xml']);assert.match(cache,/recordCount="5000"/);assert.match(cache,/saveData="1"/);assert.match(cache,/refreshOnLoad="0"/);
 const after=new Workbook(readXlsx(bytes).data),src=pivotSourceData(after,after.sheets[1].pivot);assert.equal(src.cube.n,5000);assert.equal(src.cube.col(2).get(4999),4999);assert.equal(src.cube.col(0).get(13),'');assert.equal(src.cube.col(0).get(14),null);assert.equal(after.getValue(0,1,2),99999);
 assert.deepEqual(after.sheets[1].pivot.filters.채널,['A']);assert.equal(after.sheets[1].slicers.length,1);assert.equal(after.sheets[1].slicers[0].source.field,'분류');
});
