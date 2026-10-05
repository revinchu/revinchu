// Local synthetic tests for the common popup module. Actual app/bundle shortcuts: excel-shortcut-audit.mjs.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import path from 'node:path';
const engine=process.env.WIXEL_BROWSER||'chromium',url=process.env.WIXEL_URL||'http://localhost:5178/';
const output=process.env.WIXEL_OUTPUT||`D:/Codex/Temp/wixel-popup-key-conflicts-${engine}`;
const filter=process.env.WIXEL_FILTER||'';
assert.ok(['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname),'Local synthetic source server only');
const playwright=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
assert.ok(['chromium','firefox','webkit'].includes(engine));
await mkdir(output,{recursive:true});
const browser=await playwright[engine].launch(),results=[];let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(v,m)=>{checks++;assert.ok(v,m);};
const ui=(p,fn)=>p.evaluate(async source=>{const module=await import('/src/ui.js');return new Function('ui',`return (${source})(ui)`)(module);},fn.toString());
const raf=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const calls=p=>p.evaluate(()=>window.popupCalls);
const key=(p,id)=>p.locator('#'+id).getAttribute('data-resolved-access-key');
async function run(name,action){
 if(filter&&!name.includes(filter))return;
 const context=await browser.newContext({viewport:{width:1280,height:900}}),p=await context.newPage(),pageErrors=[],blocked=[],writes=[];const begin=checks;
 p.setDefaultTimeout(12000);p.on('pageerror',e=>pageErrors.push(e.message));
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method()+' '+u.pathname);return r.abort();}if(u.origin!==new URL(url).origin){blocked.push(u.origin);return r.abort();}if(u.pathname.startsWith('/api/'))return r.fulfill({status:404,body:'{}',contentType:'application/json'});return r.continue();});
 try{
  await p.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.popupCalls=[];});
  await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());
  await action(p);eq(pageErrors,[],'page errors');eq(writes,[],'remote writes');eq(blocked,[],'external requests');
  await p.screenshot({path:path.join(output,name+'.png')});results.push({name,ok:true,checks:checks-begin,pageErrors,writes,blocked});console.log('PASS '+name);
 }catch(e){await p.screenshot({path:path.join(output,name+'-failed.png')}).catch(()=>{});results.push({name,ok:false,checks:checks-begin,error:e.stack,pageErrors,writes,blocked});console.error('FAIL '+name+' '+e.stack);}finally{await context.close();}
 await writeFile(path.join(output,'result.json'),JSON.stringify({engine,url,sourceOnly:true,checks,passed:results.filter(r=>r.ok).length,total:results.length,results},null,2));
}
try{
 await run('duplicate-caption-alias',async p=>{
  await ui(p,({openMenu})=>openMenu({x:80,y:240},[{label:'복사(C)',accessKey:'c',accessAliases:'x',action:()=>popupCalls.push('copy')},{label:'내용 지우기(C)',accessKey:'c',accessAliases:'x c',action:()=>popupCalls.push('clear')} ]));
  const rows=p.locator('#menuLayer .menu-item'),keys=await rows.evaluateAll(ns=>ns.map(n=>({key:n.dataset.resolvedAccessKey,text:n.textContent,aria:n.getAttribute('aria-keyshortcuts')})));
  eq(keys[0].key,'c');ok(keys[1].key!=='c');ok(keys[1].text.includes('('+keys[1].key.toUpperCase()+')'));ok(!keys[1].text.includes('(C)'));eq(keys[0].aria,'Alt+C Alt+X');ok(!keys[1].aria.includes('Alt+C')&&!keys[1].aria.includes('Alt+X'));
  await p.keyboard.press('c');eq(await calls(p),['copy']);
  await ui(p,({openMenu})=>openMenu({x:80,y:240},[{label:'복사(C)',accessKey:'c',action:()=>popupCalls.push('copy')},{label:'내용 지우기(C)',accessKey:'c',action:()=>popupCalls.push('clear')} ]));
  const reassigned=await rows.nth(1).getAttribute('data-resolved-access-key');await p.keyboard.press(reassigned);eq(await calls(p),['copy','clear']);
 });
 await run('dialog-explicit-over-default-cancel',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'합성 선택 옵션',body:el('label',{},'메모(C)',el('input',{id:'memo',type:'checkbox'})),buttons:[{label:'확인',action:()=>{popupCalls.push('ok');return false;}},{label:'취소',action:()=>{popupCalls.push('cancel');return false;}}]}));
  eq(await key(p,'memo'),'c');const cancel=p.locator('.dialog-foot button').last(),cancelKey=await cancel.getAttribute('data-resolved-access-key');ok(cancelKey!=='c');eq(await p.locator('.dialog-foot button').first().getAttribute('data-resolved-access-key'),'o');
  await p.keyboard.press('Alt+c');eq(await p.locator('#memo').isChecked(),true);eq(await calls(p),[]);await p.keyboard.press('Alt+'+cancelKey);eq(await calls(p),['cancel']);await p.keyboard.press('Escape');eq(await p.locator('.dialog').count(),0);
  await ui(p,({openDialog,el})=>openDialog({title:'명시 취소 우선',body:el('div'),buttons:[{label:'취소(C)',action:()=>{popupCalls.push('explicit');return false;}},{label:'메모(C)',action:()=>{popupCalls.push('memo');return false;}}]}));
  await p.keyboard.press('Alt+c');eq(await calls(p),['cancel','explicit']);
 });
 await run('dynamic-label-stability',async p=>{
  await ui(p,({openDialog,el})=>{openDialog({title:'동적 표시 검사',body:el('div',{id:'dynamicBody'},el('button',{id:'first',accessKey:'c'},'첫 명령(C)'),el('button',{id:'second',accessKey:'c','aria-label':'두 번째(C)',title:'두 번째(C)',onclick:()=>popupCalls.push('second')},'두 번째(C)'),el('button',{id:'english',accessKey:'c'},'&Close'))});window.mutations=0;new MutationObserver(rs=>window.mutations+=rs.length).observe(document.querySelector('.dialog'),{subtree:true,characterData:true,childList:true});});
  const second=await key(p,'second'),english=await key(p,'english');ok(second&&second!=='c');ok(english!==second&&english!=='c');
  await raf(p);await raf(p);const mutations=await p.evaluate(()=>window.mutations);await raf(p);await raf(p);eq(await p.evaluate(()=>window.mutations),mutations,'no mutation loop');
  eq(await p.locator('#second').getAttribute('aria-label'),`두 번째(${second.toUpperCase()})`);eq(await p.locator('#second').getAttribute('title'),`두 번째(${second.toUpperCase()})`);ok((await p.locator('#english').textContent()).includes('Close'),'English word preserved');
  await ui(p,({el})=>document.querySelector('#dynamicBody').append(el('button',{accessKey:'z'},'추가(Z)')));await raf(p);eq(await key(p,'second'),second);eq(await key(p,'english'),english);
  await p.keyboard.press('Alt+'+second);eq(await calls(p),['second']);
 });
 await run('hidden-disabled-reveal',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'표시 상태',body:el('div',{},el('button',{id:'hidden',hidden:true,accessKey:'c',onclick:()=>popupCalls.push('hidden')},'숨김(C)'),el('button',{id:'disabled',disabled:true,accessKey:'c',onclick:()=>popupCalls.push('disabled')},'비활성(C)'),el('button',{id:'active',accessKey:'c',onclick:()=>popupCalls.push('active')},'활성(C)'))}));
  eq(await key(p,'hidden'),null);eq(await key(p,'disabled'),null);eq(await key(p,'active'),'c');await p.keyboard.press('Alt+c');eq(await calls(p),['active']);
  await p.locator('#hidden').evaluate(n=>n.hidden=false);await raf(p);eq(await key(p,'hidden'),'c');const active=await key(p,'active');ok(active!=='c');await p.keyboard.press('Alt+c');await p.keyboard.press('Alt+'+active);eq(await calls(p),['active','hidden','active']);
  await p.locator('#hidden').evaluate(n=>n.hidden=true);await raf(p);eq(await key(p,'active'),'c');eq(await key(p,'hidden'),null);
 });
 await run('owned-nested-scope',async p=>{
  await ui(p,({openDialog,el,registerAccessKeyScope})=>{const child=el('div',{id:'childScope',style:{border:'1px solid',padding:'8px'}},el('button',{id:'childAction',accessKey:'c',onclick:()=>popupCalls.push('child')},'자식(C)'));openDialog({title:'부모 범위',body:el('div',{},el('button',{id:'parentAction',accessKey:'c',onclick:()=>popupCalls.push('parent')},'부모(C)'),child)});let unregister;unregister=registerAccessKeyScope(child,{owner:document.querySelector('.dialog'),onClose:()=>{unregister();child.remove();}});});
  await raf(p);eq(await key(p,'parentAction'),'c');eq(await key(p,'childAction'),'c');await p.keyboard.press('Alt+c');eq(await calls(p),['child']);await p.keyboard.press('Escape');eq(await p.locator('#childScope').count(),0);await p.keyboard.press('Alt+c');eq(await calls(p),['child','parent']);
 });
 await run('typing-ime-modifiers',async p=>{
  await ui(p,({openMenu,el})=>openMenu({x:90,y:250},[{node:el('div',{},el('input',{id:'typing',value:''}),el('button',{id:'command',accessKey:'c',onclick:()=>popupCalls.push('command')},'명령(C)'))}]));
  await p.locator('#typing').focus();await p.keyboard.type('cccc');eq(await p.locator('#typing').inputValue(),'cccc');eq(await calls(p),[]);
  await p.keyboard.press('Alt');const ime=await p.locator('#typing').evaluate(n=>{const e=new KeyboardEvent('keydown',{key:'ㅊ',code:'KeyC',isComposing:true,keyCode:229,bubbles:true,cancelable:true});n.dispatchEvent(e);return e.defaultPrevented;});eq(ime,false);eq(await calls(p),[]);
  const physical=await p.locator('#typing').evaluate(n=>{const e=new KeyboardEvent('keydown',{key:'ㅊ',code:'KeyC',altKey:true,isComposing:true,keyCode:229,bubbles:true,cancelable:true});n.dispatchEvent(e);return e.defaultPrevented;});eq(physical,true);eq(await calls(p),['command']);
  for(const modifiers of [{ctrlKey:true},{metaKey:true},{ctrlKey:true,altKey:true}]){const prevented=await p.locator('#typing').evaluate((n,mods)=>{const e=new KeyboardEvent('keydown',{key:'c',code:'KeyC',bubbles:true,cancelable:true,...mods});n.dispatchEvent(e);return e.defaultPrevented;},modifiers);eq(prevented,false);}eq(await calls(p),['command']);
 });
 await run('offscreen-reassigned-scroll',async p=>{
  await ui(p,({openMenu,el})=>openMenu({x:90,y:250},[{node:el('div',{id:'scrollbox',style:{height:'100px',overflow:'auto'}},el('button',{id:'top',accessKey:'c',style:{display:'block'}},'위(C)'),el('div',{style:{height:'380px'}}),el('button',{id:'bottom',accessKey:'c',style:{display:'block'},onclick:()=>popupCalls.push('bottom')},'아래(C)'))}]));
  const bottom=await key(p,'bottom');ok(bottom&&bottom!=='c');eq(await p.locator('#scrollbox').evaluate(n=>n.scrollTop),0);await p.keyboard.press(bottom);eq(await calls(p),['bottom']);ok(await p.locator('#scrollbox').evaluate(n=>n.scrollTop>0));
  const visible=await p.locator('#bottom').evaluate(n=>{const b=n.getBoundingClientRect(),a=n.parentElement.getBoundingClientRect();return b.top>=a.top&&b.bottom<=a.bottom;});eq(visible,true);
 });
 await run('overflow-tab-access',async p=>{
  await ui(p,({openDialog,el})=>openDialog({title:'많은 옵션',body:el('div',{},Array.from({length:42},(_,i)=>el('button',{id:'opt'+i,style:{display:'block'},onclick:()=>popupCalls.push(i)},'합성 선택 '+i)))}));
  const assigned=await p.locator('.dialog [data-resolved-access-key]').evaluateAll(ns=>ns.map(n=>n.dataset.resolvedAccessKey));eq(assigned.length,36);eq(new Set(assigned).size,36);eq(await key(p,'opt41'),null);
  await p.locator('#opt40').focus();await p.keyboard.press('Tab');eq(await p.evaluate(()=>document.activeElement.id),'opt41');await p.keyboard.press('Enter');eq(await calls(p),[41]);
 });
}finally{await browser.close();}
const report={engine,url,sourceOnly:true,checks,passed:results.filter(r=>r.ok).length,total:results.length,results};await writeFile(path.join(output,'result.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({engine,checks,passed:report.passed,total:report.total,output}));if(results.some(r=>!r.ok))process.exitCode=1;
