// 합성 문서 전용. 원격 API·쓰기 및 외부 요청을 차단합니다.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
const engine=process.env.WIXEL_BROWSER||'chromium',url=process.env.WIXEL_URL||'http://localhost:5195/',origin=new URL(url).origin;
assert.ok(['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname)||process.env.WIXEL_ROW_LIMIT_PUBLIC==='1'&&origin==='https://wixel-3.wizx.workers.dev','로컬 또는 승인된 합성 공개 검사만 허용');
const out=process.env.WIXEL_ROW_LIMIT_OUT||'D:/Codex/Temp/wixel-row-limit-20261005/'+engine;
assert.match(out,/^D:[/\\]/i); await mkdir(out,{recursive:true});
const filter=process.env.WIXEL_ROW_LIMIT_FILTER||'',modulePath=process.env.PLAYWRIGHT_MODULE;
const pw=await import(modulePath?(/^[A-Za-z]:[/\\]/.test(modulePath)?pathToFileURL(modulePath).href:modulePath):'playwright');
const browser=await pw[engine].launch(),EXCEL=1048576,EXTENDED=20000000,results=[]; let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(a,m)=>{checks++;assert.ok(a,m);};
const run=(p,cmd)=>p.evaluate(cmd=>tabula.run(cmd),cmd);
const state=p=>p.evaluate(()=>({active:{...tabula.active},sel:{...tabula.sel},sheet:tabula.si,limit:tabula.gv().rowLimit,axis:tabula.gv().rows.max,extent:tabula.gv().extR,aria:document.getElementById('accessibleGrid')?.getAttribute('aria-rowcount'),count:tabula.wb().sheets[0].cells.size,option:JSON.parse(localStorage.getItem('wixel.options')||'{}').extendedRows,editing:!document.getElementById('cellEditor').classList.contains('idle')}));
async function fixture(p,cells={},extra={}){
 await p.evaluate(({cells,extra})=>{tabula.wb().restore({sheets:[{name:'행수 검사',cells,...extra},{name:'두 번째',cells:{'0,0':{raw:'second'}}}]});tabula.wb().undoStack=[];tabula.wb().redoStack=[];tabula.switchSheet(0);tabula.gv().layout();tabula.gv().renderAll();tabula.selectCell(0,0);},{cells,extra});
 await p.locator('#cellEditor').focus();
}
async function options(p){await run(p,'options');const d=p.getByRole('dialog',{name:'WIXEL 옵션',exact:true});await d.waitFor();return d;}
async function extended(p,on){const d=await options(p);await d.getByRole('checkbox',{name:'위셀 행수 확장',exact:true}).setChecked(on);await d.getByRole('button',{name:/^확인/}).click();await d.waitFor({state:'hidden'});await p.waitForFunction(n=>window.tabula.gv().rowLimit===n,on?EXTENDED:EXCEL);}
async function go(p,ref){const n=p.locator('#nameBox');await n.fill(ref);await n.press('Enter');}
async function paste(p,text){await p.locator('#cellEditor').focus();await p.locator('#cellEditor').evaluate((node,text)=>{const e=new Event('paste',{bubbles:true,cancelable:true});Object.defineProperty(e,'clipboardData',{value:{getData:type=>type==='text/plain'?text:'',files:[],types:['text/plain']}});node.dispatchEvent(e);},text);}
async function openSynthetic(p,name,cells){await p.locator('#fileInput').setInputFiles({name:name+'.json',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({sheets:[{name,cells}]}))});await p.waitForFunction(name=>tabula.wb().sheets[0]?.name===name&&!document.querySelector('.load-progress'),name,{timeout:15000});}
async function test(name,body){
 if(filter&&!filter.split('|').some(s=>name.includes(s)))return;
 const start=checks,context=await browser.newContext({viewport:{width:1440,height:960},serviceWorkers:'block'}),p=await context.newPage(),errors=[],writes=[];
 p.setDefaultTimeout(6000);p.on('pageerror',e=>errors.push(e.message));
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.url());return r.abort();}return u.origin===origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
 try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>!!window.tabula?.gv(),null,{timeout:15000});await fixture(p);await body(p);eq(errors,[],'브라우저 오류 없음');eq(writes,[],'원격 쓰기 없음');results.push({name,ok:true,checks:checks-start});console.log('OK '+name);}
 catch(error){results.push({name,ok:false,checks:checks-start,error:error.stack,state:await state(p).catch(()=>null),errors,writes});console.error('NG '+name+': '+error.message);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});}
 finally{await context.close();}
}
try{
 await test('default-excel-limit-and-ctrl-down',async p=>{
  const s=await state(p);eq(s.limit,EXCEL,'기본 행수');eq(s.axis,EXCEL,'표시 축');eq(s.aria,String(EXCEL),'접근성 행수');
  await p.keyboard.press('Control+ArrowDown');eq((await state(p)).active.r,EXCEL-1,'빈 열 Ctrl↓');eq(await p.locator('#nameBox').inputValue(),'A1048576','마지막 주소');
  await p.keyboard.press('ArrowDown');eq((await state(p)).active.r,EXCEL-1,'아래 이동 제한');ok((await state(p)).extent<=EXCEL,'가상 스크롤 한도');
 });
 await test('general-option-cancel-and-keyboard',async p=>{
  const d=await options(p),c=d.getByRole('checkbox',{name:'위셀 행수 확장',exact:true});eq(await c.isChecked(),false,'기본 해제');await c.focus();await p.keyboard.press('Space');eq(await c.isChecked(),true,'키보드 체크');
  eq((await state(p)).limit,EXCEL,'확인 전 보존');await d.getByRole('button',{name:/^취소/}).click();eq((await state(p)).limit,EXCEL,'취소 보존');
  const again=await options(p);eq(await again.getByRole('checkbox',{name:'위셀 행수 확장',exact:true}).isChecked(),false,'취소 복원');await p.keyboard.press('Escape');eq((await state(p)).limit,EXCEL,'Escape 보존');
 });
 await test('enable-extended-rows-and-edit',async p=>{
  await extended(p,true);eq((await state(p)).axis,EXTENDED,'확장 축');eq((await state(p)).aria,String(EXTENDED),'확장 접근성');await go(p,'A1048577');eq((await state(p)).active.r,EXCEL,'다음 행 이동');
  await p.locator('#cellEditor').focus();await p.keyboard.type('extended-value');await p.keyboard.press('Enter');eq(await p.evaluate(r=>tabula.wb().getRaw(0,r,0),EXCEL),'extended-value','확장 입력');
  await go(p,'A20000000');eq((await state(p)).active.r,EXTENDED-1,'확장 마지막 행');await p.locator('#cellEditor').focus();await p.keyboard.type('last-value');await p.keyboard.press('Enter');
  eq(await p.evaluate(r=>tabula.wb().getRaw(0,r,0),EXTENDED-1),'last-value','마지막 행 입력');eq((await state(p)).active.r,EXTENDED-1,'마지막 Enter');
  const d=await options(p);await p.screenshot({path:out+'/extended-option.png'});await d.getByRole('button',{name:/^취소/}).click();
 });
 await test('namebox-outside-limit-and-api-clamp',async p=>{
  await go(p,'B2');const a=(await state(p)).active;for(const ref of ['A1048577','A1:A1048577','1048577:1048578']){await go(p,ref);eq((await state(p)).active,a,ref+' 선택 보존');}
  eq(await p.evaluate(()=>tabula.wb().names.length),0,'잘못된 이름 생성 없음');await p.evaluate(()=>tabula.selectCell(19999999,0));eq((await state(p)).active.r,EXCEL-1,'선택 API 한도');
 });
 await test('disable-retains-high-data-and-formulas',async p=>{
  await extended(p,true);await fixture(p,{'0,0':{raw:'normal'},'0,1':{raw:'=A1048577+A20000000'},'1048576,0':{raw:'7'},'19999999,0':{raw:'11'}});
  await go(p,'A20000000');const before=await p.evaluate(()=>({count:tabula.wb().sheets[0].cells.size,value:tabula.wb().getValue(0,0,1)}));eq(before.value,18,'고행 참조 계산');await extended(p,false);
  const s=await state(p);eq(s.limit,EXCEL,'확장 해제');ok(s.active.r<EXCEL,'선택 안전한 행');eq(s.count,before.count,'셀 삭제 없음');
  eq(await p.evaluate(()=>[tabula.wb().getRaw(0,1048576,0),tabula.wb().getRaw(0,19999999,0),tabula.wb().getValue(0,0,1)]),['7','11',18],'고행·수식 보존');
  await extended(p,true);await go(p,'A20000000');eq((await state(p)).active.r,EXTENDED-1,'재활성화 접근');
 });
 await test('option-reload-persistence',async p=>{
  await extended(p,true);eq((await state(p)).option,true,'옵션 저장');await p.reload();await p.waitForFunction(()=>!!window.tabula?.gv(),null,{timeout:15000});eq((await state(p)).limit,EXTENDED,'다시 열기 확장 유지');
  await extended(p,false);await p.reload();await p.waitForFunction(()=>!!window.tabula?.gv(),null,{timeout:15000});eq((await state(p)).limit,EXCEL,'다시 열기 기본 유지');eq((await state(p)).option,false,'기본 옵션 저장');
 });
 await test('file-and-sheet-switch-preserves-mode',async p=>{
  await extended(p,true);await openSynthetic(p,'첫 파일',{'0,0':{raw:'first'},'1048576,0':{raw:'kept'}});eq((await state(p)).limit,EXTENDED,'파일 열기 옵션 보존');
  await go(p,'A1048577');await openSynthetic(p,'다음 파일',{'0,0':{raw:'next'}});eq(await p.evaluate(()=>tabula.wb().getRaw(0,0,0)),'next','다음 파일');eq((await state(p)).limit,EXTENDED,'다음 파일 옵션');
  await extended(p,false);await fixture(p);await p.evaluate(()=>tabula.switchSheet(1));eq((await state(p)).axis,EXCEL,'시트 전환 축');eq(await p.evaluate(()=>tabula.wb().getRaw(1,0,0)),'second','시트 내용');
  await openSynthetic(p,'마지막 파일',{'0,0':{raw:'final'},'1048576,0':{raw:'hidden-kept'}});eq((await state(p)).limit,EXCEL,'고행 파일도 기본 유지');eq(await p.evaluate(()=>tabula.wb().getRaw(0,1048576,0)),'hidden-kept','가려진 고행 보존');
 });
 await test('whole-column-selection-and-end',async p=>{
  await p.keyboard.press('Control+Space');eq((await state(p)).sel.r2,EXCEL-1,'전체 열 기본');await extended(p,true);await p.locator('#cellEditor').focus();await p.keyboard.press('Control+Space');eq((await state(p)).sel.r2,EXTENDED-1,'전체 열 확장');
  await extended(p,false);await fixture(p,{'0,0':{raw:'first'},'1048575,2':{raw:'edge'},'19999999,3':{raw:'beyond'}});await p.keyboard.press('Control+End');ok((await state(p)).active.r<EXCEL,'CtrlEnd 고행 차단');
  await go(p,'A1');await p.locator('#cellEditor').focus();await p.keyboard.press('Control+Shift+ArrowDown');eq((await state(p)).sel.r2,EXCEL-1,'선택 확장 한도');
 });
 await test('external-paste-boundary-is-atomic',async p=>{
  await fixture(p,{'1048575,0':{raw:'bottom'},'1048576,0':{raw:'preserved'}});await go(p,'A1048576');const before=await p.evaluate(()=>tabula.wb().undoStack.length);await paste(p,'one\ntwo');
  eq(await p.evaluate(()=>[tabula.wb().getRaw(0,1048575,0),tabula.wb().getRaw(0,1048576,0)]),['bottom','preserved'],'초과 붙여넣기 부분 적용 없음');eq(await p.evaluate(()=>tabula.wb().undoStack.length),before,'거부 Undo 없음');
  await extended(p,true);await go(p,'A1048576');await paste(p,'one\ntwo');eq(await p.evaluate(()=>[tabula.wb().getRaw(0,1048575,0),tabula.wb().getRaw(0,1048576,0)]),['one','two'],'확장 시 동일 붙여넣기 성공');
 });
 await test('internal-copy-boundary-is-atomic',async p=>{
  await fixture(p,{'0,0':{raw:'one'},'1,0':{raw:'two'},'1048575,0':{raw:'bottom'}});await go(p,'A1:A2');await run(p,'copy');await go(p,'A1048576');await run(p,'paste');
  eq(await p.evaluate(()=>[tabula.wb().getRaw(0,1048575,0),tabula.wb().getRaw(0,1048576,0)]),['bottom',''],'내부 복사 초과 거부');eq(await p.evaluate(()=>tabula.wb().getRaw(0,0,0)),'one','원본 보존');
 });
 await test('bottom-row-keyboard-edit-boundary',async p=>{
  await go(p,'A1048576');await p.locator('#cellEditor').focus();await p.keyboard.type('bottom');await p.keyboard.press('Enter');eq((await state(p)).active.r,EXCEL-1,'Enter 한도');eq(await p.evaluate(()=>tabula.wb().getRaw(0,1048575,0)),'bottom','마지막 행 편집');
  await p.keyboard.press('Shift+ArrowDown');eq((await state(p)).sel.r2,EXCEL-1,'ShiftDown 한도');await p.keyboard.press('Tab');eq((await state(p)).active.r,EXCEL-1,'Tab 행 보존');eq((await state(p)).active.c,1,'Tab 다음 열');
 });
 await test('virtual-grid-accessibility-and-frozen-boundary',async p=>{
  await fixture(p,{'0,0':{raw:'frozen'},'1048575,0':{raw:'visible-last'},'1048576,0':{raw:'not-visible'}},{freeze:{rows:2,cols:1}});await go(p,'B1048576');
  const v=await p.evaluate(()=>({rows:[...document.querySelectorAll('[data-a11y-cell]')].map(n=>Number(n.dataset.r)),aria:document.getElementById('accessibleGrid').getAttribute('aria-rowcount'),extent:tabula.gv().extR,frozen:tabula.gv().fr}));
  eq(v.aria,String(EXCEL),'틀 고정 ARIA');ok(v.rows.length>0&&v.rows.every(r=>r<EXCEL),'접근성 셀 한도');ok(v.extent<=EXCEL,'스크롤 한도');eq(v.frozen,2,'틀 고정 보존');
  await extended(p,true);await go(p,'B1048577');eq((await state(p)).active.r,EXCEL,'틀 고정 확장 이동');await extended(p,false);eq((await state(p)).axis,EXCEL,'틀 고정 해제 축');
 });
}finally{await browser.close();}
const summary={engine,url,cases:results.length,passed:results.filter(r=>r.ok).length,checks,results};
await writeFile(out+'/result.json',JSON.stringify(summary,null,2));console.log(JSON.stringify({...summary,results:undefined,out}));
if(!results.length||results.some(r=>!r.ok))process.exitCode=1;
