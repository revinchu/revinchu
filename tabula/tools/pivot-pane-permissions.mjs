// 합성 문서만 사용. 실제 UI에서 계산 필드 권한·이전 초안 차단·반복 저장과 Undo를 검사한다.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {gzipSync} from 'node:zlib';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/',out=process.env.WIXEL_PIVOT_PERMISSIONS_OUT||'D:/Codex/Temp/wixel-pivot-pane-permissions/source',only=process.env.WIXEL_PIVOT_PERMISSIONS_FILTER||'';
await mkdir(out,{recursive:true});
const browser=await chromium.launch(),results=[],errors=[],writes=[],assets=new Set();let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m)},ok=(a,m)=>{checks++;assert.ok(a,m)};
const snapshot=p=>p.evaluate(()=>{const w=window.tabula.wb();return{book:JSON.stringify(w.serialize()),undo:w.undoStack.length,redo:w.redoStack.length}});
const pivot=p=>p.evaluate(()=>structuredClone(window.tabula.wb().sheets[window.tabula.si].pivot));
const total=p=>p.evaluate(()=>{const t=window.tabula,w=t.wb(),d=w.sheets[t.si].pivot;return w.getValue(t.si,d.area.r2,d.area.c2)});
const dialog=p=>p.getByRole('dialog',{name:'계산 필드',exact:true});
const run=(p,cmd)=>p.evaluate(cmd=>window.tabula.run(cmd),cmd);
const calcButton=p=>p.locator('#pivotPane .pp-tools').getByRole('button',{name:'ƒx 계산 필드...',exact:true});
async function setup(p){
 await p.evaluate(()=>{const t=window.tabula,cells={};[['지역','상품','매출'],['서울','A',100],['부산','B',200]].forEach((row,r)=>row.forEach((v,c)=>cells[r+','+c]={raw:String(v)}));t.wb().restore({sheets:[{name:'합성 원본',cells}]});t.selectRange({r1:0,c1:0,r2:2,c2:2});t.run('insertPivot')});
 await p.getByRole('dialog',{name:'피벗 테이블 만들기',exact:true}).getByRole('button',{name:'확인',exact:true}).click();
 await p.locator('#pivotPane [data-pivot-field="지역"]').check();await p.locator('#pivotPane [data-pivot-field="매출"]').check();
}
async function open(p){await calcButton(p).click();await dialog(p).waitFor();await dialog(p).locator('.cf-name').fill('두배');await dialog(p).locator('.cf-formula').fill('=매출*2');await p.waitForFunction(()=>!document.querySelector('.cf-opts .primary')?.disabled)}
async function save(p){await dialog(p).getByRole('button',{name:'저장',exact:true}).click()}
async function close(p){await dialog(p).locator('.dialog-foot').getByRole('button',{name:'닫기',exact:true}).click()}
async function protect(p,permit){await p.evaluate(permit=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setSheetProp(t.si,'protect',{on:true,allow:{selectLocked:true,selectUnlocked:true,pivotTables:permit}}))},permit)}
// 합성 addSheet가 예약한 350ms 자동 갱신까지 완료한 뒤 사용자 편집을 시작한다.
// 제품 타이머/guard는 변경하지 않는다. 준비 후 발생하는 외부 변경은 아래 stale 검사에서 거절해야 한다.
async function settleFixture(p){const ready=await p.evaluate(()=>new Promise((resolve,reject)=>{const w=window.tabula.wb(),start=performance.now(),versions=[w.version];let v=w.version,at=start;const poll=()=>{const now=performance.now();if(w.version!==v){v=w.version;at=now;versions.push(v)}if(now-at>=500)return resolve({versions,elapsedMs:Math.round(now-start)});if(now-start>8000)return reject(new Error('합성 피벗 준비 중 문서 변경이 안정되지 않았습니다.'));setTimeout(poll,50)};poll()}));console.log('FIXTURE_READY '+JSON.stringify(ready))}
async function sibling(p,permit){await p.evaluate(permit=>{const t=window.tabula,w=t.wb(),source=structuredClone(w.sheets[t.si].pivot);w.transact(()=>{const n=w.addSheet('연결 보고서');w.setSheetProp(n,'pivot',{...source,name:'연결 피벗'});if(permit!==null)w.setSheetProp(n,'protect',{on:true,allow:{selectLocked:true,selectUnlocked:true,pivotTables:permit}})})},permit);await settleFixture(p)}
const styleButton=p=>p.locator('#pivotPane .pp-tools').getByRole('button',{name:'스타일...',exact:true});
const styleChip=p=>p.locator('.style-grid.pstyles .style-chip:not(.on)').first();
async function openStyle(p){await styleButton(p).click();await styleChip(p).waitFor()}

async function test(name,fn){
 if(only&&!name.includes(only))return;
 const c=await browser.newContext({viewport:{width:1366,height:950}}),p=await c.newPage(),pe=[],pw=[];p.setDefaultTimeout(10000);
 p.on('pageerror',e=>{pe.push(e.message);errors.push({name,error:e.message})});p.on('dialog',d=>d.type()==='confirm'||d.type()==='beforeunload'?d.accept():d.dismiss());
 await c.addInitScript(()=>{window.WIXEL_SKIP_START=true;window.TABULA_STATIC=true;localStorage.setItem('wixel:version','3.0.0')});
 await c.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method()+' '+u.pathname);pw.push(u.pathname);return r.abort()}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return r.abort();return r.continue()});
 try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());for(const src of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(src);await setup(p);await fn(p);eq(pe,[],'페이지 오류');eq(pw,[],'원격 쓰기');results.push({name,ok:true});console.log('OK '+name)}
 catch(e){results.push({name,ok:false,error:e.message});console.error('NG '+name+': '+e.stack);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{})}
 finally{await c.close()}
}
try{
 await test('같은 창 생성·수정·삭제와 세 번 Undo',async p=>{
  const before=await snapshot(p);await open(p);await save(p);eq((await pivot(p)).calcFields,[{name:'두배',formula:'매출*2'}]);eq(await total(p),600,'실제 집계 결과');
  await dialog(p).locator('.cf-formula').fill('=매출*3');await dialog(p).locator('.cf-formula').press('Control+Enter');eq((await pivot(p)).calcFields,[{name:'두배',formula:'매출*3'}],'같은 창 재저장');eq(await total(p),900,'수정된 실제 집계 결과');
  await dialog(p).getByRole('button',{name:'삭제',exact:true}).click();eq((await pivot(p)).calcFields,[]);await close(p);eq((await snapshot(p)).undo,before.undo+3);
  for(let i=0;i<3;i++)await run(p,'undo');eq((await snapshot(p)).book,before.book,'원본 복원');
 });
 await test('피벗 사용 허용 보호시트 생성·삭제·Undo',async p=>{
  await protect(p,true);const before=await snapshot(p);await open(p);await save(p);eq((await pivot(p)).calcFields?.[0]?.name,'두배');await dialog(p).getByRole('button',{name:'삭제',exact:true}).click();eq((await pivot(p)).calcFields,[]);await close(p);await run(p,'undo');await run(p,'undo');eq((await snapshot(p)).book,before.book);
 });
 for(const mode of ['protected','markedFinal'])await test(mode+' 필드 목록과 계산 필드 진입 차단',async p=>{
  if(mode==='protected')await protect(p,false);else await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setBookProp('props',{...w.props,markedFinal:true}))});
  const before=await snapshot(p);await p.locator('#pivotPane [data-pivot-field="상품"]').evaluate(n=>{n.checked=true;n.dispatchEvent(new Event('change',{bubbles:true}))});await calcButton(p).evaluate(n=>n.dispatchEvent(new MouseEvent('click',{bubbles:true})));
  eq(await dialog(p).count(),0,'직접 이벤트도 차단');eq(await snapshot(p),before,'원본·Undo 불변');
 });
 await test('공개 읽기 전용 합성 문서 계산 필드 진입 차단',async p=>{
  const workbook=await p.evaluate(()=>window.tabula.wb().serialize()),hash=gzipSync(JSON.stringify({docName:'읽기 전용 합성',workbook,view:{grid:true,headers:true}})).toString('base64url');
  await p.goto(url+'#view='+hash,{waitUntil:'domcontentloaded'});await p.reload({waitUntil:'domcontentloaded'});await p.waitForSelector('body.view-mode');await p.evaluate(()=>{const t=window.tabula;t.switchSheet(1);const d=t.wb().sheets[1].pivot;t.selectCell(d.area.r1,d.area.c1)});
  const before=await snapshot(p);await calcButton(p).evaluate(n=>n.dispatchEvent(new MouseEvent('click',{bubbles:true})));eq(await dialog(p).count(),0);eq(await snapshot(p),before);
 });
 await test('공유 대상 보호시 부분 저장 금지·현재 피벗만 적용 허용',async p=>{
  await sibling(p,false);await open(p);const before=await snapshot(p);await save(p);eq(await snapshot(p),before,'공유 대상 원자적 차단');
  await dialog(p).locator('.cf-opts input[type=checkbox]').first().uncheck();await save(p);eq((await pivot(p)).calcFields?.[0]?.name,'두배');eq(await p.evaluate(()=>window.tabula.wb().sheets.find(s=>s.name==='연결 보고서').pivot.calcFields??[]),[]);
  await close(p);await run(p,'undo');eq((await snapshot(p)).book,before.book);
 });
 await test('허용된 공유 대상 저장과 한 번 Undo',async p=>{
  await sibling(p,true);const before=await snapshot(p);await open(p);await save(p);eq(await p.evaluate(()=>window.tabula.wb().sheets.filter(s=>s.pivot).map(s=>s.pivot.calcFields?.[0]?.formula)),['매출*2','매출*2']);eq((await snapshot(p)).undo,before.undo+1);await close(p);await run(p,'undo');eq((await snapshot(p)).book,before.book);
 });
 for(const mode of ['protected','markedFinal','source','definition','workbook','sheet','sharedDefinition'])await test('열린 창 이후 '+mode+' 변경시 이전 초안 저장 차단',async p=>{
  if(mode==='sharedDefinition')await sibling(p,null);await open(p);
  await p.evaluate(mode=>{const t=window.tabula,w=t.wb(),s=w.sheets[t.si];if(mode==='protected')w.transact(()=>w.setSheetProp(t.si,'protect',{on:true,allow:{pivotTables:false}}));if(mode==='markedFinal')w.transact(()=>w.setBookProp('props',{...w.props,markedFinal:true}));if(mode==='source')w.transact(()=>w.setInput(0,1,2,'999'));if(mode==='definition')w.transact(()=>w.setSheetProp(t.si,'pivot',{...s.pivot,layout:'outline'}));if(mode==='workbook')w.restore(w.serialize());if(mode==='sheet')t.switchSheet(0);if(mode==='sharedDefinition')w.sheets.find(s=>s.name==='연결 보고서').pivot.calcFields=[{name:'외부 변경',formula:'매출+1'}]},mode);
  const before=await snapshot(p);await dialog(p).locator('.cf-formula').press('Control+Enter');eq(await snapshot(p),before,'이전 초안은 원본·Undo 불변');ok(await dialog(p).isVisible(),'초안 창 유지');
 });
 await test('계산 필드 삭제도 외부 변경 후 차단',async p=>{
  await open(p);await save(p);await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setInput(0,1,2,'777'))});const before=await snapshot(p);await dialog(p).getByRole('button',{name:'삭제',exact:true}).click();eq(await snapshot(p),before);
 });
 await test('구조 보호 문서에서 수식 나열의 새 시트 생성 차단',async p=>{
  await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setBookProp('props',{...w.props,lockStructure:true}))});await open(p);const before=await snapshot(p);await dialog(p).getByRole('button',{name:'수식 나열',exact:true}).click();eq(await snapshot(p),before);ok(await dialog(p).isVisible());
 });
 await test('스타일 갤러리 보호·최종본 차단과 허용 보호시트 적용·Undo',async p=>{
  await settleFixture(p);await protect(p,false);let before=await snapshot(p);await styleButton(p).click();eq(await p.locator('.style-grid.pstyles').count(),0,'보호된 스타일 메뉴 진입 거절');eq(await snapshot(p),before);
  await protect(p,true);await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setBookProp('props',{...w.props,markedFinal:true}))});before=await snapshot(p);await styleButton(p).click();eq(await p.locator('.style-grid.pstyles').count(),0,'최종본 스타일 메뉴 진입 거절');eq(await snapshot(p),before);
  await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setBookProp('props',{...w.props,markedFinal:false}))});await settleFixture(p);before=await snapshot(p);const original=(await pivot(p)).style;
  await openStyle(p);await styleChip(p).click();ok((await pivot(p)).style!==original,'허용 보호시트에서 실제 스타일 변경');eq((await snapshot(p)).undo,before.undo+1);await run(p,'undo');eq((await snapshot(p)).book,before.book,'스타일 Undo');
  before=await snapshot(p);await openStyle(p);await p.getByRole('menuitem',{name:'지우기 (스타일 없음)',exact:true}).click();eq((await pivot(p)).style,'None');eq((await snapshot(p)).undo,before.undo+1);await run(p,'undo');eq((await snapshot(p)).book,before.book,'지우기 Undo');
 });
 await test('열린 스타일 갤러리의 오래된 칩·지우기 실행 차단',async p=>{
  await settleFixture(p);
  for(const clear of [false,true]){
   await openStyle(p);await p.evaluate(clear=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setInput(t.si,30,10,clear?'2':'1'))},clear);const before=await snapshot(p);
   if(clear)await p.getByRole('menuitem',{name:'지우기 (스타일 없음)',exact:true}).click();else await styleChip(p).click();
   eq(await snapshot(p),before,clear?'오래된 지우기 실행은 원본·Undo 불변':'오래된 칩 실행은 원본·Undo 불변');await p.keyboard.press('Escape');
  }
 });
}finally{await browser.close()}
const summary={url,assets:[...assets],cases:results.length,passed:results.filter(x=>x.ok).length,checks,pageErrors:errors,remoteWrites:writes,results};await writeFile(out+'/result.json',JSON.stringify(summary,null,2));console.log(JSON.stringify(summary));if(summary.passed!==summary.cases||errors.length||writes.length)process.exitCode=1;

