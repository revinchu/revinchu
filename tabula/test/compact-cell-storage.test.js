import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';

function fixture(count=1_000_000) {
 return new Workbook({sheets:[{name:'서식',cells:{'0,0':{raw:'제목'},'2,2':{raw:'42'}},cellRuns:[[10,0,count,{raw:'',style:{fill:'#aabbcc',bold:true,bb:true,bbc:'#001122'}}]]}]});
}
test('백만 빈 셀은 JSON과 Blob에서 같은 짧은 범위로 보존하고 편집/구조 Undo가 분리된다',async()=>{
 const w=fixture(),encoded=w.serialize(),text=JSON.stringify(encoded),blob=await w.serializeBlob().text();
 assert.equal(text,blob);assert.ok(text.length<10000);assert.equal(encoded.sheets[0].cellRuns[0][2],1_000_000);
 const loaded=new Workbook(JSON.parse(blob));assert.equal(loaded.sheets[0].cells.size,1_000_002);assert.deepEqual(loaded.usedRange(0),{rows:3,cols:3});assert.deepEqual(loaded.extent(0),{rows:1_000_010,cols:3});
 for(const r of [10,500000,1000009])assert.equal(loaded.styleAt(0,r,0).fill,'#aabbcc');
 loaded.transact(()=>loaded.setInput(0,500000,0,'테스트'));assert.equal(loaded.getRaw(0,500000,0),'테스트');assert.equal(loaded.getRaw(0,500001,0),'');loaded.undo();assert.equal(loaded.getRaw(0,500000,0),'');
 const small=fixture(1000);small.transact(()=>small.insertRows(0,30,2));small.undo();assert.equal(small.sheets[0].cells.size,1002);assert.equal(small.styleAt(0,1009,0).fill,'#aabbcc');
});
for(const runCount of [511,512,513,1024])test(`Blob cellRuns 청크 경계 ${runCount}개에서 유효한 JSON 유지`,async()=>{
 const w=new Workbook();const cells=w.sheets[0].cells;
 for(let i=0;i<runCount;i++)cells.setRunRC(i*256,0,128,{raw:'',style:{fill:i%2?'#ffffff':'#abcdef'}});
 const text=await w.serializeBlob().text();assert.equal(text,JSON.stringify(w.serialize()));
 const loaded=new Workbook(JSON.parse(text));assert.equal(loaded.sheets[0].cells.size,runCount*128);assert.equal(loaded.getCell(0,(runCount-1)*256+127,0).style.fill,(runCount-1)%2?'#ffffff':'#abcdef');
});
