// 새 격리 문서와 지연 가상 GET만 사용한다. 사용자 보관함/문서 API에 접속하지 않는다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
assert.ok(['localhost','127.0.0.1'].includes(new URL(url).hostname), '소스 로컬 서버 전용');
const browser = await chromium.launch(), results=[];
const result = (name, value) => ({docName:name,workbook:{sheets:[{name:'합성',cells:{'0,0':{raw:String(value)}}}]}});
async function test(name, fn, { vault = false } = {}) {
  const ctx=await browser.newContext(),p=await ctx.newPage(), pending=[], writes=[], errors=[];p.setDefaultTimeout(8000);
  p.on('pageerror',e=>errors.push(e.message));
  await ctx.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD'].includes(q.method())){writes.push(q.method());return r.abort();}if(u.origin!==new URL(url).origin)return r.abort();if(u.pathname==='/api/files')return r.fulfill({json:[{name:'합성 A',modified:1,size:1},{name:'합성 B',modified:1,size:1}]});if(u.pathname.startsWith('/api/files/')){pending.push({route:r,name:decodeURIComponent(u.pathname.slice('/api/files/'.length))});return;}if(u.pathname.startsWith('/api/'))return r.abort();return r.continue();});
  await ctx.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');});
  try {
    await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());
    await p.evaluate(async(vault)=>{const {server}=await import('/src/storage.js');server.available=true;server.vault=vault;server.connect(vault ? 'A'.repeat(43) : null);},vault);
    const open=async name=>{const n=pending.length;await p.evaluate(()=>window.tabula.run('backstage'));await p.locator('[data-page=open]').click();await p.getByRole('button',{name,exact:true}).click();for(let i=0;i<80&&pending.length===n;i++)await p.waitForTimeout(25);assert.equal(pending.length,n+1);return pending[n];};
    const release=async(req,value=7,status=200)=>{await req.route.fulfill({status,json:status===200?result(req.name,value):{error:'합성 권한 오류'}});await p.waitForTimeout(100);};
    const fresh=async()=>{await p.evaluate(()=>{window.__before=window.tabula.wb().sheets;window.tabula.run('newWorkbook');});await p.waitForFunction(()=>window.tabula.wb().sheets!==window.__before);};
    const waitRequest=async n=>{for(let i=0;i<80&&pending.length<n;i++)await p.waitForTimeout(25);assert.equal(pending.length,n,'예상한 인증 재시도 요청 수');return pending[n-1];};
    await fn(p,{open,release,fresh,waitRequest,pending});assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);results.push({name,ok:true});console.log('OK '+name);
  }catch(e){results.push({name,ok:false,error:e.message,errors,writes});console.error('NG '+name+': '+e.message);}finally{await ctx.close();}
}
const value=p=>p.evaluate(()=>window.tabula.wb().getValue(0,0,0));
const edit=p=>p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setInput(0,0,0,'900'));});
try {
await test('늦은 온라인 열기 응답은 새 문서의 입력을 덮지 않음',async(p,{open,release,fresh})=>{const req=await open('합성 A');await fresh();await edit(p);await release(req);assert.equal(await value(p),900);});
await test('같은 문서의 추가 편집 뒤 오래된 열기 응답 폐기',async(p,{open,release})=>{const req=await open('합성 A');await edit(p);await release(req);assert.equal(await value(p),900);});
await test('셀 편집 중인 입력도 늦은 응답이 취소하지 않음',async(p,{open,release})=>{const req=await open('합성 A');await p.locator('#cellEditor').focus();await p.keyboard.press('F2');await p.keyboard.type('900');await release(req);assert.equal(await p.locator('#cellEditor').inputValue(),'900');await p.keyboard.press('Enter');assert.equal(await value(p),900);});
await test('A 뒤 B 열기를 선택하면 응답 순서와 무관하게 B 유지',async(p,{open,release})=>{const a=await open('합성 A'),b=await open('합성 B');await release(b,20);assert.equal(await value(p),20);await release(a,10);assert.equal(await value(p),20);});
await test('문서 교체 뒤 이전 열기의 401이 암호 창을 열지 않음',async(p,{open,release,fresh})=>{const req=await open('합성 A');await fresh();await release(req,0,401);assert.equal(await p.getByRole('dialog',{name:'서버 암호',exact:true}).count(),0);});
await test('일반 온라인 문서 열기는 정상 적용',async(p,{open,release})=>{const req=await open('합성 A');await release(req,7);assert.equal(await value(p),7);});
await test('서버 암호 재입력 뒤 정상 문서 열기 재시도',async(p,{open,release,waitRequest})=>{const req=await open('합성 A');await release(req,0,401);const dlg=p.getByRole('dialog',{name:'서버 암호',exact:true});await dlg.getByLabel('암호',{exact:true}).fill('synthetic-token');await dlg.getByRole('button',{name:'확인',exact:true}).click();const retry=await waitRequest(2);assert.equal(retry.route.request().headers()['x-tabula-token'],'synthetic-token');await release(retry,17);assert.equal(await value(p),17);});
await test('개인 보관함 재연결 뒤 정상 문서 열기 재시도',async(p,{open,release,waitRequest})=>{const req=await open('합성 A');await release(req,0,401);const dlg=p.getByRole('dialog',{name:'개인 보관함 연결',exact:true}), key='B'.repeat(42)+'A';await dlg.getByLabel('개인 보관함 복구키',{exact:true}).fill(key);await dlg.getByRole('button',{name:'기존 복구키로 연결',exact:true}).click();const retry=await waitRequest(2);assert.equal(retry.route.request().headers()['x-wixel-vault'],key);await release(retry,27);assert.equal(await value(p),27);},{vault:true});
await test('인증 창을 연 뒤 문서가 바뀌면 재연결 완료로 열기 재개하지 않음',async(p,{open,release,pending})=>{const req=await open('합성 A');await release(req,0,401);const dlg=p.getByRole('dialog',{name:'개인 보관함 연결',exact:true});await edit(p);await dlg.getByLabel('개인 보관함 복구키',{exact:true}).fill('B'.repeat(42)+'A');await dlg.getByRole('button',{name:'기존 복구키로 연결',exact:true}).click();await p.waitForTimeout(250);assert.equal(pending.length,1);assert.equal(await value(p),900);},{vault:true});
await test('자동 저장 꺼짐의 오류 안내가 저장 완료를 보장하지 않음',async(p)=>{await p.locator('#autosaveToggle').click();assert.equal(await p.locator('#autosaveToggle').getAttribute('aria-checked'),'false');await edit(p);await p.evaluate(()=>window.dispatchEvent(new ErrorEvent('error',{message:'합성 처리 오류'})));const text=await p.locator('#toast').allTextContents();assert.ok(text.some(x=>x.includes('합성 처리 오류')));assert.ok(text.every(x=>!x.includes('자동 저장되어 있습니다')));});
}finally{await browser.close();console.log(JSON.stringify({url,total:results.length,good:results.filter(x=>x.ok).length,bad:results.filter(x=>!x.ok).length,results},null,2));if(results.some(x=>!x.ok))process.exitCode=1;}
