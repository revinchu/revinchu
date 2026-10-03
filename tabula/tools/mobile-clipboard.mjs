// 모바일 내부 셀 복사: OS 클립보드는 모의하고 실제 키 입력과 native 이벤트 부재/중복을 분리한다.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {gzipSync} from 'node:zlib';
const engines=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const engine=process.env.WIXEL_BROWSER||'chromium',url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
assert.ok(['chromium','webkit'].includes(engine));
const out=process.env.WIXEL_CLIPBOARD_OUT||'D:/Codex/Temp/wixel-mobile-clipboard/source-'+engine,filter=process.env.WIXEL_CLIPBOARD_FILTER||'';
await mkdir(out,{recursive:true});
const browser=await engines[engine].launch(),results=[],errors=[],writes=[];let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(v,m)=>{checks++;assert.ok(v,m);};
const frame=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const fixture={sheets:[{name:'합성 복사',cells:{'0,0':{raw:'12',style:{bold:true,fill:'#ffeedd'},comment:'합성 메모'},'0,1':{raw:'=A1*2'},'1,0':{raw:'한글🙂'},'1,1':{raw:'0'},'3,0':{raw:'대상'},'3,1':{raw:'보존'}}},{name:'다른 시트',cells:{'3,0':{raw:'다른 대상'}}}]};
const run=(p,c)=>p.evaluate(c=>tabula.run(c),c),select=(p,r=3,c=0)=>p.evaluate(({r,c})=>tabula.selectCell(r,c),{r,c});
const state=p=>p.evaluate(()=>({raw:tabula.wb().getRaw(tabula.si,3,0),other:tabula.wb().getRaw(tabula.si,3,1),source:tabula.wb().getRaw(0,0,0),undo:tabula.wb().undoStack.length,focus:document.activeElement.id,reads:window.__cb.reads,writes:window.__cb.writes,trace:window.__cb.keys}));
async function test(name,fn,{preference='hardware',mobile=true,readonly=false}={}){
 if(filter&&!name.includes(filter))return;
 const ctx=await browser.newContext({viewport:{width:390,height:844},isMobile:true,hasTouch:true,serviceWorkers:'block'}),p=await ctx.newPage(),pe=[],pw=[];p.setDefaultTimeout(10000);
 p.on('pageerror',e=>{pe.push(e.message);errors.push({name,message:e.message});});
 await ctx.addInitScript(({preference,mobile})=>{
  window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1',mobile?'on':'off');localStorage.setItem('wixel.mobile-keyboard.v1',preference);
  window.__cb={reads:0,writes:0,denied:true,text:'',keys:[]};
  Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>{__cb.reads++;if(__cb.denied)throw new DOMException('합성 권한 거절','NotAllowedError');return __cb.text;},writeText:async t=>{__cb.writes++;if(__cb.denied)throw new DOMException('합성 권한 거절','NotAllowedError');__cb.text=t;}}});
  // 앱 핸들러 이후 OS 기본 동작만 막는다. 사용자 OS 클립보드는 읽거나 쓰지 않는다.
  window.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&/^[cvx]$/i.test(e.key)){__cb.keys.push({key:e.key,target:e.target.id,focused:document.activeElement.id,handled:e.defaultPrevented});e.preventDefault();}});
  window.__nativeClipboard=(type,text='')=>{const data={},e=new Event(type,{bubbles:true,cancelable:true});Object.defineProperty(e,'clipboardData',{value:{getData:t=>t==='text/plain'?text:'',setData:(t,v)=>{data[t]=v;},files:[]}});document.activeElement.dispatchEvent(e);return{handled:e.defaultPrevented,data};};
 },{preference,mobile});
 await ctx.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD'].includes(q.method())){writes.push(u.pathname);pw.push(u.pathname);return r.abort();}return u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
 try{
  const target=readonly?url+'#view='+gzipSync(JSON.stringify({docName:'읽기 전용 합성',workbook:fixture})).toString('base64url'):url;
  await p.goto(target,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());
  if(!readonly)await p.evaluate(data=>{tabula.wb().restore(data);tabula.switchSheet(1);tabula.switchSheet(0);},fixture);
  await p.evaluate(()=>{tabula.selectCell(0,0);tabula.wb().undoStack=[];tabula.wb().redoStack=[];});await frame(p);
  await fn(p);eq(pe,[],'페이지 오류 없음');eq(pw,[],'원격 쓰기 없음');results.push({name,ok:true,state:await state(p)});console.log('OK '+name);
 }catch(e){results.push({name,ok:false,error:e.message,state:await state(p).catch(()=>null)});console.error('NG '+name+': '+e.stack);}
 finally{await ctx.close();}
}
try{
 for(const preference of ['screen','auto','hardware'])for(const mod of ['Control','Meta'])await test(preference+' '+mod+' 셀 복사·붙여넣기·Undo: native 이벤트와 권한 없이',async p=>{
  await p.evaluate(()=>tabula.selectRange({r1:0,c1:0,r2:1,c2:1}));await p.keyboard.press(mod+'+c');eq((await state(p)).undo,0,'복사는 문서를 변경하지 않음');await select(p);await p.keyboard.press(mod+'+v');await frame(p);
  const s=await state(p);eq(s.raw,'12');eq(s.other,'=A4*2');eq(s.source,'12');eq(s.reads,0,'내부 붙여넣기에 OS 읽기/권한창 없음');eq(s.writes,1);eq(s.undo,1);ok(s.trace.every(x=>x.handled),'앱에서 단축키 소비');
  const cell=await p.evaluate(()=>{const w=tabula.wb();return{bold:w.styleAt(0,3,0).bold,fill:w.styleAt(0,3,0).fill,comment:w.getCell(0,3,0).comment,text:w.getRaw(0,4,0)};});eq(cell,{bold:true,fill:'#ffeedd',comment:'합성 메모',text:'한글🙂'});await run(p,'undo');eq((await state(p)).raw,'대상');eq((await state(p)).source,'12');
 },{preference});
 await test('화면 모드 native 복사·외부 붙여넣기 이벤트도 수신',async p=>{const copied=await p.evaluate(()=>__nativeClipboard('copy'));eq(copied.handled,true);eq(copied.data['text/plain'],'12');await select(p);eq((await p.evaluate(()=>__nativeClipboard('paste','외부 합성'))).handled,true);eq((await state(p)).raw,'외부 합성');eq((await state(p)).undo,1);},{preference:'screen'});
 await test('복사·붙여넣기 trailing native 이벤트가 중복 적용되지 않음',async p=>{await p.keyboard.press('Control+c');const copied=await p.evaluate(()=>__nativeClipboard('copy'));eq(copied.data['text/plain'],'12');eq((await state(p)).writes,1);await select(p);await p.keyboard.press('Control+v');const before=await state(p);eq((await p.evaluate(()=>__nativeClipboard('paste','중복 금지'))).handled,true);eq((await state(p)).raw,before.raw);eq((await state(p)).undo,1);await p.keyboard.press('Control+v');eq((await state(p)).raw,'12');});
 await test('잘라내기 한 번 이동·native 중복 무시·Undo 복구',async p=>{await p.keyboard.press('Control+x');eq((await state(p)).source,'12');await select(p);await p.keyboard.press('Control+v');await p.evaluate(()=>__nativeClipboard('paste','중복 금지'));eq((await state(p)).raw,'12');eq((await state(p)).source,'');eq((await state(p)).undo,1);await run(p,'undo');eq((await state(p)).source,'12');eq((await state(p)).raw,'대상');});
 await test('같은 앱의 메뉴 열기·시트 이동은 내부 사본 유지',async p=>{await p.keyboard.press('Control+c');await p.locator('#mobileTools').click();await p.keyboard.press('Escape');await p.evaluate(()=>{tabula.switchSheet(1);tabula.selectCell(3,0);});await p.keyboard.press('Control+v');eq((await state(p)).raw,'12');eq((await state(p)).reads,0);eq((await state(p)).undo,1);});
 for(const kind of ['blur','hidden'])await test(kind+' 이후 외부 자료를 오래된 내부 사본보다 우선',async p=>{await p.keyboard.press('Control+c');await select(p);await p.evaluate(kind=>{if(kind==='blur')window.dispatchEvent(new Event('blur'));else{Object.defineProperty(document,'hidden',{configurable:true,value:true});document.dispatchEvent(new Event('visibilitychange'));Object.defineProperty(document,'hidden',{configurable:true,value:false});}},kind);await p.keyboard.press('Control+v');eq((await state(p)).raw,'대상');eq((await state(p)).trace.at(-1).handled,false);await p.evaluate(()=>__nativeClipboard('paste','외부 새 값'));eq((await state(p)).raw,'외부 새 값');eq((await state(p)).undo,1);});
 for(const text of ['', '12'])await test('blur 뒤 native 외부 입력 '+(text?'같은 텍스트는 값만':'빈 이벤트는 무동작'),async p=>{await p.keyboard.press('Control+c');await select(p);await p.evaluate(()=>window.dispatchEvent(new Event('blur')));await p.evaluate(text=>__nativeClipboard('paste',text),text);eq((await state(p)).raw,text||'대상');eq((await state(p)).undo,text?1:0);const style=await p.evaluate(()=>tabula.wb().styleAt(0,3,0));eq(!!style.bold,false,'이전 사본 굵게 서식을 재사용하지 않음');eq(style.fill??null,null,'이전 사본 채우기를 재사용하지 않음');});
 await test('문서 교체 후 이전 사본으로 새 문서에 붙이지 않음',async p=>{await p.keyboard.press('Control+c');await p.evaluate(()=>{tabula.wb().restore({sheets:[{name:'새 합성',cells:{'3,0':{raw:'새 문서'}}}]});tabula.selectCell(3,0);tabula.wb().undoStack=[];});await p.keyboard.press('Control+v');eq((await state(p)).raw,'새 문서');eq((await state(p)).undo,0);});
 for(const guard of ['보호','최종본','읽기 전용'])await test(guard+'에서 복사는 허용·붙여넣기는 차단',async p=>{await p.keyboard.press('Control+c');await select(p);if(guard==='보호')await p.evaluate(()=>tabula.wb().sheets[0].protect={on:true});if(guard==='최종본')await p.evaluate(()=>tabula.wb().props={markedFinal:true});await p.keyboard.press('Control+v');eq((await state(p)).raw,'대상');eq((await state(p)).undo,0);},{readonly:guard==='읽기 전용'});
 await test('F2 편집 중 Ctrl+C/V는 텍스트 native 경로 유지',async p=>{await p.keyboard.press('F2');await p.keyboard.press('Control+c');await p.keyboard.press('Control+v');ok((await state(p)).trace.every(x=>!x.handled));eq((await state(p)).writes,0);eq((await p.evaluate(()=>__nativeClipboard('paste','편집 문자'))).handled,false);await p.keyboard.press('Escape');eq((await state(p)).source,'12');eq((await state(p)).undo,0);});
 await test('셀 복사 뒤 편집 텍스트 native 복사는 이전 셀 사본 우선권 해제',async p=>{await p.keyboard.press('Control+c');await p.keyboard.press('F2');eq((await p.evaluate(()=>__nativeClipboard('copy'))).handled,false);await p.keyboard.press('Escape');await select(p);await p.keyboard.press('Control+v');eq((await state(p)).raw,'대상');eq((await state(p)).trace.at(-1).handled,false);await p.evaluate(()=>__nativeClipboard('paste','편집 복사 문자'));eq((await state(p)).raw,'편집 복사 문자');eq((await state(p)).undo,1);});
 await test('한글 조합 Ctrl+C/V는 셀 사본을 만들지 않음',async p=>{const r=await p.evaluate(()=>{const e=document.activeElement;const result=[];for(const key of ['c','v']){const evt=new KeyboardEvent('keydown',{key,code:'Key'+key.toUpperCase(),ctrlKey:true,isComposing:true,bubbles:true,cancelable:true});e.dispatchEvent(evt);result.push(window.__cb.keys.at(-1).handled);}return result;});eq(r,[false,false]);eq((await state(p)).writes,0);eq((await state(p)).undo,0);});
 await test('데스크톱 Ctrl+C/V는 기존 native 경로 유지',async p=>{await p.keyboard.press('Control+c');eq((await state(p)).writes,0);eq((await state(p)).trace.at(-1).handled,false);eq((await p.evaluate(()=>__nativeClipboard('copy'))).data['text/plain'],'12');await select(p);await p.evaluate(()=>__nativeClipboard('paste','12'));eq((await state(p)).raw,'12');},{mobile:false});
}finally{await browser.close();const summary={url,engine,cases:results.length,passed:results.filter(r=>r.ok).length,checks,pageErrors:errors,remoteWrites:writes,scope:'Synthetic workbook and mock OS clipboard. Trusted Playwright modifier keys; browser default clipboard blocked after app handlers. Synthetic copy/paste tests native delivery and duplicates. Physical iPhone/Bluetooth and OS process exits are not reproduced.',results};await writeFile(out+'/result.json',JSON.stringify(summary,null,2));console.log(JSON.stringify({...summary,results:results.map(({name,ok,error})=>({name,ok,error}))}));if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;}
