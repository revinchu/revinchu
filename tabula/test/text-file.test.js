import { test } from 'node:test';
import assert from 'node:assert/strict';
import { detectTextFileEncoding, readDecodedChunks } from '../src/text-file.js';
import { delimitedRowWidth } from '../src/csv.js';
const example = '서울,매출\n부산,100\n';
const cp949 = Buffer.from('bcadbfef2cb8c5c3e20abacebbea2c3130300a','hex');
function streamed(bytes, chunk = 1) {
 const blob = new Blob([bytes]);
 return {size:blob.size,slice:(a,b)=>blob.slice(a,b),stream:()=>new ReadableStream({start(c){for(let p=0;p<bytes.length;p+=chunk)c.enqueue(bytes.subarray(p,p+chunk));c.close();}})};
}
async function decoded(file, encoding) { let text=''; let last=0; await readDecodedChunks(file,encoding,(chunk,bytes)=>{assert.ok(bytes>=last);last=bytes;text+=chunk;});assert.equal(last,file.size);return text; }
test('streaming decoding preserves Korean text, BOM and characters split across chunks',async()=>{
 const le=Buffer.from(example,'utf16le'),be=Buffer.from(le);be.swap16();
 const cases=[['utf-8',Buffer.from(example)],['utf-8',Buffer.concat([Buffer.from([239,187,191]),Buffer.from(example)])],['utf-16le',Buffer.concat([Buffer.from([255,254]),le])],['utf-16be',Buffer.concat([Buffer.from([254,255]),be])],['euc-kr',cp949]];
 for(const[encoding,bytes]of cases){const file=streamed(bytes);assert.equal(await detectTextFileEncoding(file),encoding);assert.equal(await decoded(file,encoding),example);}
});
test('legacy Korean encoding is detected even after a long ASCII-only prefix',async()=>{
 const prefix='header,value\n'+'ascii,42\n'.repeat(9000),file=streamed(Buffer.concat([Buffer.from(prefix),cp949]),4093);
 assert.equal(await detectTextFileEncoding(file),'euc-kr');assert.equal(await decoded(file,'euc-kr'),prefix+example);
});
test('file read failures are propagated instead of misreported as an encoding fallback',async()=>{
 const error=new Error('synthetic disk failure'),blob=new Blob(['abc']);
 const file={size:3,slice:(a,b)=>blob.slice(a,b),stream:()=>new ReadableStream({start(c){c.error(error);}})};
 await assert.rejects(detectTextFileEncoding(file),e=>e===error);await assert.rejects(readDecodedChunks(file,'utf-8',()=>{}),e=>e===error);
});
test('consumer failure cancels and releases the stream',async()=>{
 let cancelled=false;const error=new Error('stop parsing');const file={stream:()=>new ReadableStream({start(c){c.enqueue(new Uint8Array([65]));},cancel(){cancelled=true;}})};
 await assert.rejects(readDecodedChunks(file,'utf-8',()=>{throw error;}),e=>e===error);assert.equal(cancelled,true);
});
test('delimited width handles 200,000 rows without spreading function arguments',()=>{
 const rows=Array(200000).fill(['x']);rows[199999]=Array(19).fill('last');assert.equal(delimitedRowWidth(rows),19);assert.equal(delimitedRowWidth([]),1);assert.equal(delimitedRowWidth([[],['a','b']]),2);
});
