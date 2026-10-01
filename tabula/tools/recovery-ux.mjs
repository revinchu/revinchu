// 로컬 workerd 전용 합성 보관함. 실제 사용자 문서와 인증키를 사용하지 않습니다.
import assert from 'node:assert/strict';
import { randomBytes } from 'node:crypto';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_WORKER_URL || 'http://127.0.0.1:8787';
assert.ok(['127.0.0.1', 'localhost'].includes(new URL(base).hostname));
const key = randomBytes(32).toString('base64url'), errors = [], checks = [];
const snap = v => ({docName:'복구 테스트',si:0,workbook:{sheets:[{name:'합성',cells:{'0,0':{raw:String(v)}}}]}});
async function api(path, {method='GET',data,revision}={}) {
  const res=await fetch(base+'/api/'+path,{method,headers:{'X-Wixel-Vault':key,'Content-Type':'application/json',...(revision!==undefined?{'If-Match':`"${revision}"`}:{})},body:data===undefined?undefined:JSON.stringify(data)});
  assert.ok(res.ok, `${method} ${path}: ${res.status} ${await res.clone().text()}`);return res.json();
}
const browser=await chromium.launch(), page=await browser.newPage({viewport:{width:1440,height:1000}});
page.setDefaultTimeout(15000);page.on('pageerror',e=>errors.push(e.message));
const dialog=name=>page.getByRole('dialog',{name,exact:true});
async function openFiles(){ await page.evaluate(()=>window.tabula.run('open')); await page.getByRole('button',{name:'복구 테스트',exact:true}).waitFor(); }
async function history(){await openFiles();await page.locator('.backstage-list tr.file').filter({has:page.getByRole('button',{name:'복구 테스트',exact:true})}).getByRole('button',{name:'버전 기록',exact:true}).click();}
async function openOnline(name, value) {
 await page.evaluate(()=>window.tabula.run('open'));
 await page.locator('.backstage-list tr.file').filter({has:page.getByRole('button',{name,exact:true})}).getByRole('button',{name,exact:true}).click();
 await page.waitForFunction(value=>window.tabula.wb().getValue(0,0,0)===value,value);
}
try {
 const one=await api('files/'+encodeURIComponent('복구 테스트'),{method:'PUT',revision:0,data:snap(10)});
 await api('files/'+encodeURIComponent('복구 테스트'),{method:'PUT',revision:one.revision,data:snap(20)});
 await page.addInitScript(key=>{window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.connection.v3',JSON.stringify({kind:'vault',key}));},key);
 const healthReady=page.waitForResponse(r=>r.url().endsWith('/api/health'));
 await page.goto(base,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForFunction(()=>!!window.tabula?.wb());await (await healthReady).finished();
 await history();await dialog('온라인 버전 기록').waitFor();assert.equal(await dialog('온라인 버전 기록').getByRole('button',{name:'이 버전 복원',exact:true}).count(),1);
 await dialog('온라인 버전 기록').getByRole('button',{name:'이 버전 복원',exact:true}).click();
 await dialog('온라인 버전 복원').getByRole('button',{name:'복원',exact:true}).click();
 await dialog('온라인 버전 복원').waitFor({state:'hidden'});
 assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,0,0)),10);assert.equal((await api('files/'+encodeURIComponent('복구 테스트'))).workbook.sheets[0].cells['0,0'].raw,'10');checks.push('버전 복원 후 온라인과 화면 값 일치');
 await history();const cur=(await api('files')).find(f=>f.name==='복구 테스트');
 await api('files/'+encodeURIComponent('복구 테스트'),{method:'PUT',revision:cur.revision,data:snap(99)});
 await dialog('온라인 버전 기록').getByRole('button',{name:'이 버전 복원',exact:true}).first().click();
 await dialog('온라인 버전 복원').getByRole('button',{name:'복원',exact:true}).click();
 await dialog('온라인 버전 복원').getByText('다른 저장이 먼저 완료되었습니다.',{exact:false}).waitFor();
 assert.equal((await api('files/'+encodeURIComponent('복구 테스트'))).workbook.sheets[0].cells['0,0'].raw,'99');checks.push('복원 CAS 충돌 안내와 새 서버 값 보존');
 await page.keyboard.press('Escape');await page.keyboard.press('Escape');
 await page.locator('.backstage-nav [data-page="storage"]').click();
 await page.getByRole('button',{name:'보관함 백업 복원',exact:true}).click();
 const backup={format:'wixel-vault-backup',version:1,documents:[{name:'복구 테스트',data:snap(50)},{name:'추가 문서',data:snap(70)}]};
 await dialog('온라인 보관함 백업 복원').getByLabel('보관함 백업 파일').setInputFiles({name:'synthetic-backup.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify(backup))});
 await dialog('온라인 보관함 백업 복원').getByLabel('복원할 문서: 추가 문서',{exact:true}).waitFor();
 await dialog('온라인 보관함 백업 복원').getByRole('button',{name:'선택한 문서 복원',exact:true}).click();
 await dialog('온라인 보관함 백업 복원').getByText('이미 있는 문서 이름입니다.',{exact:false}).waitFor();
 assert.equal((await api('files')).length,1);checks.push('기본 덮어쓰기 차단');
 await dialog('온라인 보관함 백업 복원').getByLabel('같은 이름의 온라인 문서 덮어쓰기').check();
 await dialog('온라인 보관함 백업 복원').getByRole('button',{name:'선택한 문서 복원',exact:true}).click();
 await dialog('온라인 보관함 백업 복원').waitFor({state:'hidden'});
 assert.equal((await api('files')).length,2);assert.equal((await api('files/'+encodeURIComponent('복구 테스트'))).workbook.sheets[0].cells['0,0'].raw,'50');
 assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,0,0)),10);
 const local=await page.evaluate(()=>JSON.parse(localStorage.getItem('tabula.v1')??'null'));
 // Interface confirms that the still-open older workbook is kept as a browser copy.
 assert.match(await page.locator('.backstage-main').innerText(),/이 브라우저/);checks.push('백업 원자 복원과 현재 화면 사본 보존');

 // A real IndexedDB quota failure must stop before restoring either the server or the visible workbook.
 await openOnline('복구 테스트',50);
 if(await page.locator('#autosaveToggle').getAttribute('aria-checked')==='true')await page.locator('#autosaveToggle').click();
 await page.evaluate(()=>{
  const put=IDBObjectStore.prototype.put;
  window.restoreSyntheticLibraryPut=()=>{IDBObjectStore.prototype.put=put;delete window.restoreSyntheticLibraryPut;};
  IDBObjectStore.prototype.put=function(value,key){if(typeof key==='string'&&key.startsWith('lib:'))throw new DOMException('합성 보관함 용량 오류','QuotaExceededError');return put.call(this,value,key);};
  window.tabula.wb().transact(()=>window.tabula.wb().setInput(0,0,0,'12345'));
 });
 let restoreRequests=0;
 const countRestore=async route=>{if(route.request().method()==='POST')restoreRequests++;await route.continue();};
 await page.route('**/api/version?*',countRestore);
 try {
  await page.evaluate(()=>window.tabula.run('versionHistory'));
  await dialog('온라인 버전 기록').getByRole('button',{name:'이 버전 복원',exact:true}).first().click();
  await dialog('온라인 버전 복원').getByRole('button',{name:'복원',exact:true}).click();
  await dialog('온라인 버전 복원').getByRole('alert').filter({hasText:'복구 사본을 보관하지 못해'}).waitFor();
  assert.equal(restoreRequests,0,'브라우저 사본 저장 실패 시 온라인 복원 요청도 보내지 않음');
  assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,0,0)),12345,'미저장 현재 값 보존');
  assert.equal((await api('files/'+encodeURIComponent('복구 테스트'))).workbook.sheets[0].cells['0,0'].raw,'50','온라인 현재 값 보존');
  checks.push('복구 사본 저장 실패 시 복원 중단·현재 값과 온라인 값 보존');
 } finally {
  await page.evaluate(()=>window.restoreSyntheticLibraryPut?.());await page.unroute('**/api/version?*',countRestore);
  await page.keyboard.press('Escape');await page.keyboard.press('Escape');
 }

 // Delay A's actual PUT, switch to B, let B's autosave timer fire, then release A.
 const aName='저장 지연 A',bName='저장 전환 B';
 await api('files/'+encodeURIComponent(aName),{method:'PUT',revision:0,data:{...snap(100),docName:aName}});
 await api('files/'+encodeURIComponent(bName),{method:'PUT',revision:0,data:{...snap(200),docName:bName}});
 await openOnline(aName,100);
 await page.evaluate(()=>window.tabula.wb().transact(()=>window.tabula.wb().setInput(0,0,0,'111')));
 let releaseA,startedA,bWrites=0;
 const waitA=new Promise(resolve=>{startedA=resolve;}),holdA=new Promise(resolve=>{releaseA=resolve;});
 const aUrl=base+'/api/files/'+encodeURIComponent(aName),bUrl=base+'/api/files/'+encodeURIComponent(bName);
 const delayA=async route=>{if(route.request().method()==='PUT'){startedA();await holdA;}await route.continue();};
 const countB=async route=>{if(route.request().method()==='PUT')bWrites++;await route.continue();};
 await page.route(aUrl,delayA);await page.route(bUrl,countB);
 try {
  await page.locator('#autosaveToggle').click();
  await Promise.race([waitA,new Promise((_,reject)=>setTimeout(()=>reject(new Error('A 온라인 저장 시작 실패')),15000))]);
  await openOnline(bName,200);
  await page.evaluate(()=>window.tabula.wb().transact(()=>window.tabula.wb().setInput(0,0,0,'222')));
  await page.waitForTimeout(2200);assert.equal(bWrites,0,'이전 저장 진행 중 B 타이머가 대기 경로에 들어감');
  const bSaved=page.waitForResponse(r=>r.url()===bUrl&&r.request().method()==='PUT'&&r.ok(),{timeout:15000});
  releaseA();await (await bSaved).finished();
  assert.equal((await api('files/'+encodeURIComponent(bName))).workbook.sheets[0].cells['0,0'].raw,'222');
  assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,0,0)),222);
  checks.push('이전 문서 저장 완료 뒤 현재 문서의 보류된 자동 저장 재시도');
 } finally {releaseA();await page.unroute(aUrl,delayA);await page.unroute(bUrl,countB);}
 assert.deepEqual(errors,[]);console.log(JSON.stringify({ok:true,checks,pageErrors:errors}));
}catch(err){console.error(JSON.stringify({pageErrors:errors,body:(await page.locator('body').innerText()).slice(-7000)}));throw err;}finally{
 await browser.close();const files=await api('files');for(const f of files) await api('files/'+encodeURIComponent(f.name),{method:'DELETE',revision:f.revision});
}
