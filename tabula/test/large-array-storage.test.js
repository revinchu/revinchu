import test from 'node:test';
import assert from 'node:assert/strict';
import { storeArrayParts, restoreArrayParts } from '../src/large-array-storage.js';

test('형식 배열은 1MiB 조각으로 정확히 복원하며 NaN과 마지막 값을 보존한다', async()=>{
 const data=new Float64Array(300000).fill(NaN);data[0]=-0;data[299999]=1.25;const stored=new Map();let active=0;
 const info=await storeArrayParts(data,i=>'p'+i,async(key,part)=>{assert.equal(active++,0);await Promise.resolve();assert.ok(part.byteLength<=1<<20);stored.set(key,structuredClone(part));active--;},async()=>{});
 assert.equal(info.parts.length,3);const copy=await restoreArrayParts(info,key=>stored.get(key),async()=>{});assert.deepEqual(copy,data);assert.notEqual(copy,data);
 stored.delete(info.parts[1]);await assert.rejects(restoreArrayParts(info,key=>stored.get(key),async()=>{}),/배열 조각/);
});

test('텍스트 사전은 개수와 크기를 제한하고 자료형·유니코드·순서를 보존한다',async()=>{
 const values=['😀한글'.repeat(20000),null,true,false,23,{error:'#N/A'},...Array.from({length:4200},(_,i)=>'값'+i)],parts=new Map();
 const info=await storeArrayParts(values,i=>'p'+i,async(key,part)=>{assert.ok(part.length<=2048);parts.set(key,structuredClone(part));},async()=>{});
 assert.ok(info.parts.length>=3);assert.deepEqual(await restoreArrayParts(info,key=>parts.get(key),async()=>{}),values);
 for(const bad of [{...info,len:values.length-1},{...info,kind:'MissingType'},{...info,parts:[]}])await assert.rejects(restoreArrayParts(bad,key=>parts.get(key),async()=>{}),/배열 조각/);
});

test('쓰기 실패·취소는 후속 배열을 복제하거나 저장하지 않는다',async()=>{
 const failure=new Error('quota');let writes=0;
 await assert.rejects(storeArrayParts(new Uint8Array(4<<20),i=>'p'+i,async()=>{writes++;throw failure;},async()=>{}),error=>error===failure);assert.equal(writes,1);
 let checks=0;await assert.rejects(storeArrayParts(new Uint8Array(4<<20),i=>'p'+i,async()=>{},async()=>{if(++checks===3)throw failure;}),error=>error===failure);
 for(const values of [[],new Float64Array()]){const info=await storeArrayParts(values,()=>assert.fail(),()=>assert.fail(),async()=>{});assert.deepEqual(await restoreArrayParts(info,()=>assert.fail(),async()=>{}),values);}
});
