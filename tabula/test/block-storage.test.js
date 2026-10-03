import test from 'node:test';
import assert from 'node:assert/strict';
import { blockJsonParts } from '../src/block-storage.js';
import { blockClone } from '../src/block.js';
import { Workbook } from '../src/workbook.js';

test('블록 청크는 기존 저장 JSON과 같고 정렬/NaN/문자/논리/오류를 유지한다', async()=>{
 const block={r0:3,c0:2,n:3,perm:new Int32Array([2,0,1]),cols:[{num:new Float64Array([0,NaN,4.5]),str:new Int32Array([-1,0,-1]),dict:['한글😀'],fmt:{bold:true}},{num:null,str:new Int32Array([0,1,2]),dict:[true,false,{error:'#N/A'}],fmt:null}]};
 assert.equal([...blockJsonParts(block)].join(''),JSON.stringify(blockClone(block)));
 const w=new Workbook({sheets:[{name:'data',cells:{},blocks:[block]}]});
 const restored=new Workbook(JSON.parse(await w.serializeBlob().text()));
 for(let r=3;r<6;r++)for(let c=2;c<4;c++)assert.deepEqual(restored.getValue(0,r,c),w.getValue(0,r,c));
});

test('백만 행 블록 JSON 생성은 큰 단일 문자열이나 배열 복제를 만들지 않는다',()=>{
 const block={r0:0,c0:0,n:1000000,cols:[{num:new Float64Array(1000000).fill(NaN),str:null,dict:[],fmt:null}]};
 block.cols[0].num[999999]=42;let count=0,max=0,total=0;
 for(const part of blockJsonParts(block)){count++;max=Math.max(max,part.length);total+=part.length;}
 assert.ok(count>480);assert.ok(max<40000);assert.ok(total>13000000);
 assert.equal(block.cols[0].num[999999],42);assert.ok(Number.isNaN(block.cols[0].num[0]));
});
