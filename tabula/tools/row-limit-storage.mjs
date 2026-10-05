// 기본 화면 한도와 실제 파일 저장 범위를 분리하는 합성 저장 회귀.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
const url=process.env.WIXEL_URL||'http://127.0.0.1:5196/',origin=new URL(url).origin;
assert.ok(['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname),'로컬 합성 검사 전용');
const out=process.env.WIXEL_ROW_STORAGE_OUT||'D:/Codex/Temp/wixel-row-limit-20261005/storage';assert.match(out,/^D:[/\\]/i);await mkdir(out,{recursive:true});
const mod=process.env.PLAYWRIGHT_MODULE||'playwright';const {chromium}=await import(/^[A-Za-z]:[/\\]/.test(mod)?pathToFileURL(mod).href:mod);
const browser=await chromium.launch(),results=[];
async function test(name,fn){
 const context=await browser.newContext({acceptDownloads:true,serviceWorkers:'block'}),p=await context.newPage(),errors=[],writes=[];
 p.on('pageerror',e=>errors.push(e.message));p.setDefaultTimeout(15000);
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.url());return r.abort();}return u.origin===origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.showSaveFilePicker=undefined;window.showOpenFilePicker=undefined;});
 try{await p.goto(url);await p.waitForFunction(()=>!!window.tabula?.gv());await fn(p);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);results.push({name,ok:true});console.log('OK '+name);}
 catch(e){results.push({name,ok:false,error:e.stack,errors,writes});console.error('NG '+name+': '+e.message);}
 finally{await context.close();}
}
try{
 await test('small-csv-open-retains-extended-row',async p=>{
  const csv=Buffer.from('first\n'+'\n'.repeat(1048575)+'last');assert.ok(csv.length<8*1024*1024);
  await p.locator('#fileInput').setInputFiles({name:'행수 CSV.csv',mimeType:'text/csv',buffer:csv});
  await p.waitForFunction(()=>tabula.wb().getRaw(0,1048576,0)==='last',null,{timeout:60000});
  assert.equal(await p.evaluate(()=>tabula.gv().rows.max),1048576);
  assert.equal(await p.evaluate(()=>tabula.wb().getRaw(0,0,0)),'first');
  assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].cells.size),2,'빈 백만 행을 실제 셀로 만들지 않음');
 });
 await test('wixel-download-and-reopen-retains-hidden-values-formulas-and-metadata',async p=>{
  const fixture={sheets:[{name:'확장 보관',cells:{'0,0':{raw:'=A1048577+A20000000'},'1048576,0':{raw:'7',style:{bold:true,fill:'#abcdef'}},'19999999,0':{raw:'11'}},rowHeights:{1048576:99},hiddenRows:{1048576:true}}]};
  await p.locator('#fileInput').setInputFiles({name:'확장 보관.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(fixture))});
  await p.waitForFunction(()=>tabula.wb().sheets[0]?.name==='확장 보관');
  await p.evaluate(()=>tabula.run('saveAs'));
  const d=p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true});await d.locator('select').selectOption('wixel');await d.getByRole('button',{name:/^저장/}).click();
  const picker=p.getByRole('dialog',{name:'파일로 저장',exact:true});await picker.waitFor();
  const event=p.waitForEvent('download');await picker.getByRole('button',{name:'다운로드',exact:true}).click();const download=await event;
  const stream=await download.createReadStream(),parts=[];for await(const part of stream)parts.push(part);const buffer=Buffer.concat(parts);
  assert.ok(buffer.length>0);await writeFile(out+'/synthetic-extended.wixel',buffer);await picker.waitFor({state:'hidden'});
  await p.locator('#fileInput').setInputFiles({name:'다른 파일.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({sheets:[{name:'다른 문서',cells:{}}]}))});await p.waitForFunction(()=>tabula.wb().sheets[0]?.name==='다른 문서');
  await p.locator('#fileInput').setInputFiles({name:'재열기.wixel',mimeType:'application/octet-stream',buffer});await p.waitForFunction(()=>tabula.wb().sheets[0]?.name==='확장 보관');
  assert.deepEqual(await p.evaluate(()=>{const w=tabula.wb(),s=w.sheets[0];return [tabula.gv().rows.max,w.getRaw(0,1048576,0),w.getRaw(0,19999999,0),w.getValue(0,0,0),w.styleAt(0,1048576,0).fill,s.rowHeights[1048576],s.hiddenRows[1048576]];}),[1048576,'7','11',18,'#abcdef',99,true]);
 });
}finally{await browser.close();}
await writeFile(out+'/result.json',JSON.stringify({url,results},null,2));if(results.some(x=>!x.ok))process.exitCode=1;
