import test from 'node:test';
import assert from 'node:assert/strict';
import { storeBlobCompatible, storedBlobStream, storedBlobText } from '../src/stored-blob.js';

const unavailable = () => new DOMException('Error preparing Blob/File data to be stored in object store', 'UnknownError');
test('Blob 호환 경로는 1MiB 조각으로 한글과 바이트를 보존한다', async () => {
  const text = '가😀'.repeat(500000), blob = new Blob([text]); let calls = 0, stored;
  blob.arrayBuffer = () => { throw Error('한 번에 전체 바이트 변환 금지'); };
  await storeBlobCompatible(blob, async value => { calls++; if (value instanceof Blob) throw unavailable(); stored = value; });
  assert.equal(calls,2); assert.equal(stored.size,blob.size); assert.ok(stored.parts.length>2);
  assert.ok(stored.parts.every(part=>part.length<=1<<20)); assert.equal(await storedBlobText(stored),text);
  assert.equal(await storedBlobText(new Blob([text])),text);
});
test('공간 부족·취소·세대 충돌·관련 없는 UnknownError는 다시 쓰지 않는다', async () => {
  for (const cause of [new DOMException('Blob quota','QuotaExceededError'),new DOMException('Blob aborted','AbortError'),
    new DOMException('Database closed','UnknownError'),Object.assign(new Error('conflict'),{code:'IDB_CONFLICT'})]) {
    let calls=0;
    await assert.rejects(storeBlobCompatible(new Blob(['data']),async()=>{calls++;throw cause;}),error=>error===cause);
    assert.equal(calls,1);
  }
});
test('바이트 변환 중 문서 전환은 재시도 쓰기 전에 중단한다',async()=>{
  let checks=0,writes=0;const failure=new Error('document changed');
  await assert.rejects(storeBlobCompatible(new Blob(['a'.repeat(2<<20)]),async()=>{writes++;throw unavailable();},()=>{if(++checks===4)throw failure;}),error=>error===failure);
  assert.equal(writes,1);
});
test('손상된 바이트 조각과 합계·형식은 복원을 거부한다',()=>{
  for(const data of [null,{}, {format:'wixel-blob-bytes-v1',size:1,parts:[]},{format:'wixel-blob-bytes-v1',size:1,parts:[Uint8Array.of(1,2)]},
    {format:'wixel-blob-bytes-v1',size:1,parts:[[1]]},{format:'wixel-blob-bytes-v1',size:0,parts:[new Uint8Array()]},
    {format:'wixel-blob-bytes-v1',size:(1<<20)+1,parts:[new Uint8Array((1<<20)+1)]}])assert.throws(()=>storedBlobStream(data),/조각/);
});
test('이전 gzip 저장본은 해제 API가 없을 때 읽기 오류를 명시하고 내용을 성공으로 반환하지 않는다',async()=>{
  const native=globalThis.DecompressionStream;
  try{globalThis.DecompressionStream=undefined;await assert.rejects(storedBlobText(new Blob(['compressed']),true),/최신 브라우저/);}
  finally{globalThis.DecompressionStream=native;}
});
