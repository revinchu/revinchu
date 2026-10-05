// Focused integration regressions for shortcut review findings; synthetic local documents only.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import path from 'node:path';
const url=process.env.WIXEL_URL||'http://localhost:5178/',engine=process.env.WIXEL_BROWSER||'chromium';
const out=process.env.WIXEL_OUTPUT||`D:/Codex/Temp/wixel-shortcut-review-${engine}`,filter=process.env.WIXEL_FILTER||'';
assert.ok(['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname),'Local synthetic verification only');
assert.ok(['chromium','firefox','webkit'].includes(engine));
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');await mkdir(out,{recursive:true});const browser=await pw[engine].launch(),results=[];let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(v,m)=>{checks++;assert.ok(v,m);};
const seq=async(p,keys)=>{await p.keyboard.press('Alt');for(const k of keys)await p.keyboard.press(k);};
const doc=()=>({sheets:[{name:'Synthetic',cells:Object.fromEntries([['A','B','C','D'],['1','2','3','4']].flatMap((row,r)=>row.map((raw,c)=>[r+','+c,{raw}])) )},{name:'Second',cells:{'0,0':{raw:'Keep'}}}]});
const snapshot=p=>p.evaluate(()=>({sheets:tabula.wb().sheets.map(s=>({name:s.name,state:s.state,tabColor:s.tabColor})),raw:tabula.wb().getRaw(0,0,0),undo:tabula.wb().undoStack.length}));
const editState=p=>p.evaluate(()=>({raw:tabula.wb().getRaw(0,0,0),editing:!document.querySelector('#cellEditor').classList.contains('idle'),editor:document.querySelector('#cellEditor').value,active:tabula.active}));
async function run(name,fn,{readonly=false}={}){
 if(filter&&!filter.split('|').some(f=>name.includes(f)))return;
 const c=await browser.newContext({viewport:{width:1280,height:900},serviceWorkers:'block'}),p=await c.newPage(),errors=[],writes=[],external=[];const start=checks;
 p.setDefaultTimeout(10000);p.on('pageerror',e=>errors.push(e.message));await c.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 await c.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method()+' '+u.pathname);return r.abort();}if(u.origin!==new URL(url).origin){external.push(u.origin);return r.abort();}if(u.pathname.startsWith('/api/'))return r.fulfill({status:404,body:'{}',contentType:'application/json'});return r.continue();});
 try{
  const address=readonly?url+'#view='+Buffer.from(JSON.stringify({docName:'Synthetic view',workbook:doc(),view:{grid:true,headers:true}})).toString('base64url'):url;
  await p.goto(address,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());
  if(readonly)await p.waitForSelector('body.view-mode');else await p.evaluate(document=>{tabula.wb().restore(document);tabula.wb().undoStack=[];tabula.wb().redoStack=[];tabula.gv().layout();tabula.gv().renderAll();tabula.selectCell(0,0);},doc());
  await p.locator('#cellEditor').focus();await fn(p);eq(errors,[],'page errors');eq(writes,[],'remote writes');eq(external,[],'external requests');await p.screenshot({path:path.join(out,name+'.png')});results.push({name,ok:true,checks:checks-start,errors,writes,external});console.log('PASS '+name);
 }catch(e){await p.screenshot({path:path.join(out,name+'-failed.png')}).catch(()=>{});results.push({name,ok:false,checks:checks-start,error:e.stack,state:await editState(p).catch(()=>null),errors,writes,external});console.error('FAIL '+name+' '+e.stack);}finally{await c.close();}
 await writeFile(path.join(out,'result.json'),JSON.stringify({engine,url,passed:results.filter(r=>r.ok).length,total:results.length,checks,results},null,2));
}
try{
 await run('protected-sheet-save-as',async p=>{
  await p.evaluate(()=>tabula.wb().sheets[0].protect={on:true});const before=await snapshot(p);
  for(const chord of ['Control+Shift+s','F12']){await p.keyboard.press(chord);await p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true}).waitFor();eq(await p.getByRole('dialog').count(),1);await p.keyboard.press('Escape');eq(await snapshot(p),before,'save dialog cancellation preserves protected document');}
 });
 await run('end-mode-lifetime',async p=>{
  await p.keyboard.press('End');await p.keyboard.press('ArrowRight');eq((await editState(p)).active,{r:0,c:3},'End Arrow still moves to boundary');
  await p.evaluate(()=>tabula.selectCell(0,0));await p.keyboard.press('End');await seq(p,'dff');await p.keyboard.press('ArrowRight');eq((await editState(p)).active,{r:0,c:1},'Alt command ends End mode');
  await p.evaluate(()=>tabula.selectCell(0,0));await p.keyboard.press('End');await p.keyboard.press('F10');await p.keyboard.press('Escape');await p.keyboard.press('ArrowRight');eq((await editState(p)).active,{r:0,c:1},'F10 cancels End mode');
  await p.evaluate(()=>tabula.selectCell(0,0));await p.keyboard.press('End');await p.keyboard.press('Shift+ArrowRight');eq(await p.evaluate(()=>tabula.sel),{r1:0,c1:0,r2:0,c2:3},'End Shift Arrow selection remains available');
 });
 await run('legacy-menu-displayed-next-key',async p=>{
  await seq(p,'d');eq(await p.locator('.keytip-command-menu .access-key-hint').count(),0,'no unrelated automatic key captions');await p.locator('[data-keytip-path="df"]').click();eq(await p.locator('.keytip-command-menu').getAttribute('data-keytip-menu'),'df');eq(await p.locator('.keytip-command-menu .access-key-hint').count(),0);
  for(const [entry,k] of [['dff','F'],['dfs','S'],['dfa','A']])eq((await p.locator(`[data-keytip-path="${entry}"] .mi-key`).textContent()).trim(),k);
  await p.keyboard.press('f');ok(await p.evaluate(()=>!!tabula.wb().sheets[0].filter),'displayed F executes filter');eq(await p.locator('.keytip-command-menu').count(),0);
  await seq(p,'d');await p.locator('[data-keytip-path="df"]').click();await p.locator('[data-keytip-path="dff"]').click();eq(await p.evaluate(()=>!!tabula.wb().sheets[0].filter),false,'pointer next steps execute same action');
 });
 await run('rename-enter-escape-ime-preserves-cell',async p=>{
  const before=(await editState(p)).raw;await seq(p,'ohr');const input=p.locator('#sheetTabs input');await input.waitFor();ok(await input.evaluate(n=>n===document.activeElement),'inline input keeps focus');await p.keyboard.type('Renamed');await p.keyboard.press('Enter');eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'Renamed');eq(await editState(p),{raw:before,editing:false,editor:'',active:{r:0,c:0}},'Enter must not become a cell newline');await p.keyboard.press('ArrowRight');eq((await editState(p)).raw,before,'next movement cannot overwrite source cell');
  await seq(p,'ohr');await input.waitFor();await p.keyboard.type('Discard');await p.keyboard.press('Escape');eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'Renamed');eq((await editState(p)).editing,false);eq((await editState(p)).raw,before);
  await seq(p,'ohr');await input.waitFor();await input.fill('한글');await input.dispatchEvent('compositionstart',{data:'글'});await input.dispatchEvent('keydown',{key:'Enter',code:'Enter',isComposing:true,keyCode:229,bubbles:true,cancelable:true});eq(await input.count(),1,'IME Enter does not commit name');eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'Renamed');await input.dispatchEvent('compositionend',{data:'글'});await p.keyboard.press('Enter');eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'한글');eq((await editState(p)).editing,false);eq((await editState(p)).raw,before);
 });
 for(const readonly of [true,false])await run(readonly?'readonly-structure-guards':'marked-final-structure-guards',async p=>{
  if(!readonly)await p.evaluate(()=>tabula.wb().props.markedFinal=true);const before=await snapshot(p);
  await seq(p,'ohr');eq(await p.locator('#sheetTabs input').count(),0);eq(await snapshot(p),before);
  for(const cmd of ['renameSheet','moveCopySheet','sheetTabColor','addSheet','deleteSheet','hideSheet']){await p.evaluate(cmd=>tabula.run(cmd),cmd);eq(await snapshot(p),before,cmd+' remains blocked');eq(await p.getByRole('dialog').count(),0);eq(await p.locator('#menuLayer .menu').count(),0);}
 },{readonly});
 await run('formula-and-global-save-as',async p=>{
  await p.keyboard.press('F2');await p.keyboard.press('Control+a');await p.keyboard.type('=1+2');await p.keyboard.press('Control+Shift+s');await p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true}).waitFor();eq((await editState(p)).raw,'=1+2');eq((await editState(p)).editing,false);await p.keyboard.press('Escape');
  await p.keyboard.press('Alt+q');eq(await p.evaluate(()=>document.activeElement.id),'searchBox','global shortcut is tested with search input focused');await p.locator('#searchBox').fill('synthetic');await p.keyboard.press('Control+Shift+s');await p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true}).waitFor();eq(await p.getByRole('dialog').count(),1,'global save-as opens once outside cell editor');eq((await editState(p)).raw,'=1+2');await p.keyboard.press('Escape');
 });
 await run('direct-sheet-ui-permission-guards',async p=>{
  const tab=()=>p.locator('#sheetTabs .sheet-tab').first(),settle=()=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
  const openColor=async()=>{await tab().click({button:'right'});await p.getByRole('menuitem',{name:/^탭 색/}).click();await settle();};
  await tab().dblclick();const input=p.locator('#sheetTabs input');await input.waitFor();await input.fill('Direct rename');await p.keyboard.press('Enter');eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'Direct rename');eq((await editState(p)).raw,'A');eq((await editState(p)).editing,false);
  await openColor();const swatch=p.locator('#menuLayer button.swatch').first();await swatch.waitFor();const color=await swatch.getAttribute('data-color');await swatch.click();eq(await p.evaluate(()=>tabula.wb().sheets[0].tabColor),color,'ordinary direct tab color remains functional');
  await p.evaluate(()=>tabula.wb().props.markedFinal=true);let before=await snapshot(p);await tab().dblclick();eq(await input.count(),0,'final document direct double-click rename blocked');await openColor();eq(await p.locator('#menuLayer button.swatch').count(),0,'final document direct color menu blocked');eq(await snapshot(p),before);
  await p.evaluate(()=>tabula.wb().props.markedFinal=false);await openColor();await swatch.waitFor();await p.evaluate(()=>tabula.wb().props.markedFinal=true);before=await snapshot(p);await swatch.click();eq(await snapshot(p),before,'permission is checked again when choosing a color');
  const address=url+'#view='+Buffer.from(JSON.stringify({docName:'Synthetic view',workbook:doc(),view:{grid:true,headers:true}})).toString('base64url');await p.goto('about:blank');await p.goto(address,{waitUntil:'domcontentloaded'});await p.waitForSelector('body.view-mode');before=await snapshot(p);
  await tab().dblclick();eq(await input.count(),0,'readonly direct double-click rename blocked');await openColor();eq(await p.locator('#menuLayer button.swatch').count(),0,'readonly direct color menu blocked');eq(await snapshot(p),before);
 });
}finally{await browser.close();}
const report={engine,url,checks,passed:results.filter(r=>r.ok).length,total:results.length,results};await writeFile(path.join(out,'result.json'),JSON.stringify(report,null,2));console.log(JSON.stringify({engine,checks,passed:report.passed,total:report.total,out}));if(results.some(r=>!r.ok))process.exitCode=1;
