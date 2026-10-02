// 소형 UTF-8와 8MiB 초과 UTF-16 파일의 전체 셀 행렬을 비교한다. 합성 파일만 사용한다.
import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import { parseDelimited } from '../src/csv.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
if (!['localhost','127.0.0.1'].includes(new URL(url).hostname)) throw Error('로컬 합성 검사만 허용');
const folder='D:/Codex/Temp/wixel-csv-boundaries'; mkdirSync(folder,{recursive:true});
const fixtures=[], filler='x'.repeat(900);
for(const [eolName,eol]of [['lf','\n'],['cr','\r'],['crlf','\r\n']]){
 const text=['h,v','','"first\r\nsecond\rthird",2','','Z,9',...Array.from({length:5000},(_,i)=>`${filler},${i}`),'',''].join(eol);
 const rows=parseDelimited(text).map(r=>Array.from({length:2},(_,i)=>{const v=r[i]??'';return v===''?null:/^\d+$/.test(v)?Number(v):v;}));
 const hash=createHash('sha256').update(JSON.stringify(rows)).digest('hex');
 for(const big of [false,true]){const bytes=big?Buffer.concat([Buffer.from([255,254]),Buffer.from(text,'utf16le')]):Buffer.from(text);assert.equal(bytes.length>8*1024*1024,big);const name=`${eolName}-${big?'stream':'small'}`,file=`${folder}/${name}.csv`;writeFileSync(file,bytes);fixtures.push({name,file,big,rows:rows.length,hash,bytes:bytes.length});}
}
const browser=await chromium.launch(),results=[];
try{for(const f of fixtures){const context=await browser.newContext(),errors=[],writes=[];let started=Date.now();try{
 await context.route('**/*',route=>{const r=route.request(),u=new URL(r.url());if(!['GET','HEAD'].includes(r.method())){writes.push(r.method());return route.abort();}return u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?route.continue():route.abort();});
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');});
 const page=await context.newPage();page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(30000);await page.goto(url,{waitUntil:'domcontentloaded'});await page.waitForFunction(()=>window.tabula?.wb());if(await page.locator('#autosaveToggle').getAttribute('aria-checked')==='true')await page.locator('#autosaveToggle').click();
 started=Date.now();await page.locator('#fileInput').setInputFiles(f.file);await page.waitForFunction(name=>tabula.wb().sheets[0].name===name&&!document.querySelector('.load-progress'),f.name,{timeout:60000});
 const actual=await page.evaluate(async n=>{const w=tabula.wb(),rows=Array.from({length:n},(_,r)=>[w.getValue(0,r,0),w.getValue(0,r,1)]);const bytes=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(JSON.stringify(rows)));return{hash:Array.from(new Uint8Array(bytes),b=>b.toString(16).padStart(2,'0')).join(''),blocks:w.sheets[0].blocks.length,blockRows:w.sheets[0].blocks[0]?.n,selectedLast:tabula.sel.r2};},f.rows);
 assert.equal(actual.hash,f.hash,'모든 셀 값 및 빈 행 위치');assert.equal(actual.blocks,f.big?1:0);if(f.big)assert.equal(actual.blockRows,f.rows-1);else assert.equal(actual.selectedLast,f.rows-1);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);results.push({name:f.name,ok:true,rows:f.rows,bytes:f.bytes,elapsedMs:Date.now()-started});console.log('OK '+f.name);
 }catch(e){results.push({name:f.name,ok:false,error:e.message,errors,writes});console.error('NG '+f.name+': '+e.message);}finally{await context.close();}}}finally{await browser.close();}
console.log(JSON.stringify({url,total:results.length,passed:results.filter(r=>r.ok).length,results},null,2));if(results.some(r=>!r.ok))process.exitCode=1;
