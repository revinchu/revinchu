import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, writeXlsxBlobAsync, readXlsx } from '../src/xlsx.js';
import { unzip,textOf } from '../src/zip.js';
import { parseXml, child, kids } from '../src/xml.js';
import { pivotSourceData,resolvePivot,computePivot } from '../src/pivot.js';
function fixture(axis){
 const wb=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'Group'},'0,1':{raw:'Value'},'1,0':{raw:'Current'},'1,1':{raw:'10'}}},{name:'Report',cells:{},pivot:{name:'Empty',source:'Source',range:{r1:0,c1:0,r2:1,c2:1},rows:[],cols:[],pages:[],[axis]:['Group'],values:[{field:'Value',agg:'sum'}],filters:{Group:['Historical']},grandRows:false,grandCols:false,layout:'tabular',top:2,left:1}}]});
 wb.pivotCacheItems={retained:{fields:[{name:'Group',shared:['Historical','Current']},{name:'Value',shared:[]}]}};wb.sheets[1].pivot.cacheItemsId='retained';
 return wb;
}
for(const axis of ['rows','cols'])for(const mode of['sync','blob'])test('필터 결과가 없는 피벗은 필드·선택을 보존하고 빈 필수목록을 쓰지 않는다: '+axis+' '+mode,async()=>{
 const wb=fixture(axis),def=wb.sheets[1].pivot,original=JSON.stringify(def);
 const bytes=mode==='sync'?writeXlsx(wb):new Uint8Array(await(await writeXlsxBlobAsync(wb)).arrayBuffer());
 const root=parseXml(textOf(unzip(bytes)['xl/pivotTables/pivotTable1.xml']));
 assert.ok(child(root,axis==='rows'?'rowFields':'colFields'));
 assert.equal(child(root,axis==='rows'?'rowItems':'colItems'),null);
 for(const tag of['rowItems','colItems']){const node=child(root,tag);if(node){assert.ok(kids(node,'i').length>0);assert.equal(Number(node.attrs.count),kids(node,'i').length);}}
 assert.equal(JSON.stringify(def),original);
 const back=new Workbook(readXlsx(bytes).data),after=back.sheets[1].pivot;
 assert.deepEqual(after[axis],['Group']);assert.deepEqual(after.filters,def.filters);
 assert.equal(after.grandRows,false);assert.equal(after.grandCols,false);
 const src=pivotSourceData(back,after),result=computePivot(resolvePivot(src,after),after);
 if(axis==='rows')assert.equal(result.meta.rowItems.length,0);else assert.equal(result.meta.colLeaves.length,0);
});
