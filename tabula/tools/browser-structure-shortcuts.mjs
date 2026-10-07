// Synthetic browser structure shortcuts. User files and the OS clipboard are never used.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const repo=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5193/',origin=new URL(url).origin;
const local=['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname);
if(!local&&process.env.WIXEL_ALLOWED_TEST_URL!==url)throw Error('Public synthetic tests require the exact WIXEL_ALLOWED_TEST_URL.');
const out=path.resolve(repo,process.env.WIXEL_STRUCTURE_OUT||'.local/browser-structure-shortcuts/source');
if(!/^[dD]:/.test(out))throw Error('All outputs must stay on D:.');
const only=process.env.WIXEL_STRUCTURE_FILTER||'',baseline=process.env.WIXEL_STRUCTURE_BASELINE||'',overrides=new Map();
if(baseline){if(!local)throw Error('Baseline routes are local only.');for(const name of ['app.js','keyboard-shortcuts.js'])overrides.set('/src/'+name,execFileSync('git',['-c','safe.directory='+path.resolve(repo,'..').replaceAll(String.fromCharCode(92),'/'),'-C',repo,'show',baseline+':tabula/src/'+name],{encoding:'utf8',maxBuffer:32*1024*1024}));}
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright'),engine=process.env.WIXEL_BROWSER||'chromium';
assert.ok(['chromium','firefox','webkit'].includes(engine));await mkdir(out,{recursive:true});
const browser=await pw[engine].launch(),results=[],errors=[],writes=[],blocked=[],assets=new Set();let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(a,m)=>{checks++;assert.ok(a,m);};
const raf=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
const state=p=>p.evaluate(()=>__structureState()),book=p=>p.evaluate(()=>__structureBook());
const run=async(p,c)=>{await p.evaluate(c=>tabula.run(c),c);await raf(p);};
const raw=(p,r,c)=>p.evaluate(({r,c})=>tabula.wb().getRaw(tabula.si,r,c),{r,c});
async function fixture(p){
  await p.evaluate(()=>{const t=tabula,w=t.wb(),cells={};for(let r=0;r<40;r++)for(let c=0;c<20;c++)cells[r+','+c]={raw:'R'+r+'C'+c,...(r===6||c===6?{style:{bt:true,bb:true,bl:true,br:true,fill:'#e0e7ff'}}:{})};
    cells['35,19']={raw:'=A21',comment:'formula sentinel'};
    w.restore({sheets:[{name:'Structure',cells,rowHeights:{6:26,7:28,8:30},rowManual:{6:true,7:true,8:true},colWidths:{6:81,7:84,8:87},colManual:{6:true,7:true,8:true}}]});
    t.switchSheet(0);t.gv().setZoom(100);t.gv().renderAll();t.selectCell(0,0);w.undoStack=[];w.redoStack=[];window.__structureTrace=[];
  });await raf(p);await p.locator('#cellEditor').focus();
}
async function selectAxis(p,axis,start=6,n=1,{ui=false}={}){
  if(ui){
    for(const at of [start,...(n>1?[start+n-1]:[])]){const pt=await p.evaluate(({axis,at})=>{const g=tabula.gv();g.ensureVisible(axis==='row'?at:0,axis==='col'?at:0);const b=g.clientRect({r1:axis==='row'?at:0,r2:axis==='row'?at:0,c1:axis==='col'?at:0,c2:axis==='col'?at:0}),v=g.viewEl.getBoundingClientRect();return axis==='row'?{x:Math.round(v.left+g.hw*g.z*.5),y:Math.round((b.top+b.bottom)/2)}:{x:Math.round((b.left+b.right)/2),y:Math.round(v.top+g.hh*g.z*.5)};},{axis,at});
      if(at!==start)await p.keyboard.down('Shift');await p.mouse.click(pt.x,pt.y);if(at!==start)await p.keyboard.up('Shift');await raf(p);}
  }else{await p.evaluate(({axis,start,n})=>tabula.selectRange(axis==='row'?{r1:start,c1:0,r2:start+n-1,c2:16383}:{r1:0,c1:start,r2:1048575,c2:start+n-1},axis==='row'?'rows':'cols',axis==='row'?{r:start,c:0}:{r:0,c:start}),{axis,start,n});await raf(p);await p.locator('#cellEditor').focus();}
}
async function key(p,keys,code){const at=await p.evaluate(()=>__structureTrace.length);await p.keyboard.press(keys);await raf(p);return p.evaluate(({at,code})=>__structureTrace.slice(at).findLast(e=>e.type==='keydown'&&e.code===code),{at,code});}
const gesture=(op,numpad=false)=>({keys:'Control+Shift+'+(numpad?(op==='insert'?'NumpadAdd':'NumpadSubtract'):(op==='insert'?'Equal':'Minus')),code:numpad?(op==='insert'?'NumpadAdd':'NumpadSubtract'):(op==='insert'?'Equal':'Minus')});
async function structural(p,op,{numpad=false}={}){const g=gesture(op,numpad),e=await key(p,g.keys,g.code);eq(e?.prevented,true,'The app consumes the explicit Ctrl+Shift structure shortcut');return e;}
async function synthetic(p,data){const result=await p.evaluate(data=>{const e=new KeyboardEvent('keydown',{bubbles:true,cancelable:true,...data});document.activeElement.dispatchEvent(e);return{prevented:e.defaultPrevented,key:e.key,code:e.code,ctrl:e.ctrlKey,meta:e.metaKey,shift:e.shiftKey,composing:e.isComposing,keyCode:e.keyCode};},data);await raf(p);return result;}
async function verifyAxis(p,axis,op,n,{start=6,repeats=1}={}){
  const offset=n*repeats*(op==='insert'?1:-1);
  if(op==='insert'){for(let i=0;i<n*repeats;i++)eq(await raw(p,axis==='row'?start+i:3,axis==='col'?start+i:3),'','Each inserted line is blank');eq(await raw(p,axis==='row'?start+n*repeats:3,axis==='col'?start+n*repeats:3),'R'+(axis==='row'?start:3)+'C'+(axis==='col'?start:3),'Original selected line shifts by the exact count');}
  else eq(await raw(p,axis==='row'?start:3,axis==='col'?start:3),'R'+(axis==='row'?start+n*repeats:3)+'C'+(axis==='col'?start+n*repeats:3),'Delete closes the exact selected line count');
  eq(await raw(p,3,3),'R3C3','Cells before the structure edit remain unchanged');
  eq(await raw(p,axis==='row'?35+offset:35,axis==='col'?19+offset:19),axis==='row'?'=A'+(21+offset):'=A21','Formula location and reference track the structure edit');
}
const lineSelection=s=>({si:s.si,sel:s.sel,kind:s.kind,active:s.active});
async function expectLineSelection(p,selected,label){const actual=await state(p);eq(lineSelection(actual),lineSelection(selected),label);eq(actual.dialogs,0,'Whole-line deletion never opens the cell shift dialog');}
async function history(p,before,after){eq(after.undo,before.undo+1,'One structural key is one Undo');await run(p,'undo');eq((await book(p)).sheets,before.sheets,'Undo restores exact cells, formulas, comments, styles and dimensions');await run(p,'redo');eq((await book(p)).sheets,after.sheets,'Redo restores the exact structured worksheet');}
async function test(name,fn){
  if(only&&!only.split('|').some(v=>name.includes(v)))return;
  const context=await browser.newContext({viewport:{width:1500,height:1100},serviceWorkers:'block'}),p=await context.newPage(),info={},pe=[],cw=[];p.setDefaultTimeout(12000);
  p.on('pageerror',e=>{errors.push({name,error:e.message});pe.push(e.message);});p.on('dialog',d=>d.dismiss());
  await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');localStorage.setItem('wixel.mobile-work.v1','off');window.__clipboard={text:'',html:''};
    Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>__clipboard.text,writeText:async text=>{__clipboard.text=text;}}});
    window.__nativeCopy=()=>{const data={},e=new Event('copy',{bubbles:true,cancelable:true});Object.defineProperty(e,'clipboardData',{value:{getData:()=>'',setData:(k,v)=>data[k]=v,files:[]}});document.activeElement.dispatchEvent(e);__clipboard={text:data['text/plain']||'',html:data['text/html']||''};return{handled:e.defaultPrevented,htmlRows:(data['text/html']?.match(/<tr>/g)||[]).length};};
    window.addEventListener('keydown',e=>{if((e.ctrlKey||e.metaKey)&&e.code==='KeyC')e.preventDefault();});
  });
  await context.route('**/*',route=>{const q=route.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push({name,path:u.pathname,method:q.method()});cw.push(u.pathname);return route.abort();}if(u.origin!==origin||/^\/api(?:\/|$)/.test(u.pathname)){blocked.push({name,path:u.pathname});return route.abort();}if(overrides.has(u.pathname))return route.fulfill({status:200,contentType:'text/javascript; charset=utf-8',body:overrides.get(u.pathname)});return route.continue();});
  const start=checks;
  try{
    await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.gv());
    for(const src of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(src);
    await p.evaluate(()=>{window.__structureTrace=[];
      window.__structureState=()=>{const t=tabula,w=t.wb(),g=t.gv(),s=g.host.state(),node=document.activeElement;return{si:t.si,active:{...t.active},sel:{...t.sel},kind:s.selKind,editing:!!s.editing,object:s.chartSel??null,undo:w.undoStack.length,redo:w.redoStack.length,zoom:g.z,focus:{id:node?.id,tag:node?.tagName,value:node?.value,cmd:node?.dataset?.ribbonCommand??node?.dataset?.cmd},keytip:document.body.dataset.keytipSequence??null,dialogs:document.querySelectorAll('.dialog').length,menus:document.querySelectorAll('.menu').length,viewport:{innerWidth,scale:visualViewport?.scale}};};
      window.__structureBook=()=>({sheets:tabula.wb().serialize().sheets.map(({fileValues,_sid,...s})=>s),undo:tabula.wb().undoStack.length,redo:tabula.wb().redoStack.length});
      for(const type of ['keydown','keyup'])window.addEventListener(type,e=>{const record={type,key:e.key,code:e.code,ctrl:e.ctrlKey,meta:e.metaKey,shift:e.shiftKey,alt:e.altKey,composing:e.isComposing,keyCode:e.keyCode,trusted:e.isTrusted,target:e.target?.id,before:__structureState()};__structureTrace.push(record);setTimeout(()=>{record.prevented=e.defaultPrevented;record.after=__structureState();},0);},true);
    });
    await fn(p,info);await raf(p);eq(pe,[],'No browser exception');eq(cw,[],'No API or remote writes');
    results.push({name,ok:true,checks:checks-start,info,state:await state(p),trace:await p.evaluate(()=>__structureTrace)});console.log('OK '+name);
  }catch(e){results.push({name,ok:false,checks:checks-start,error:e.message,info,state:await state(p).catch(()=>null),trace:await p.evaluate(()=>window.__structureTrace).catch(()=>[])});console.error('NG '+name+': '+e.message.split('\n')[0]);await p.screenshot({path:path.join(out,'failure-'+results.length+'.png')}).catch(()=>{});
  }finally{await context.close();}
}
try{
  for(const axis of ['row','col'])for(const op of ['insert','delete'])for(const n of [1,3])await test(axis+'-'+n+'-CtrlShift-'+op+'-UndoRedo-F4',async(p,info)=>{
    await fixture(p);await selectAxis(p,axis,6,n,{ui:true});info.before=await state(p);eq(info.before.kind,axis==='row'?'rows':'cols','Real header selection identifies full lines');
    const before=await book(p);info.event=await structural(p,op);await verifyAxis(p,axis,op,n);info.after=await state(p);
    await expectLineSelection(p,info.before,'Insertion/deletion preserves the full selected line range and active cell');
    const after=await book(p);await history(p,before,after);await expectLineSelection(p,info.before,'Undo/Redo preserves the full-line selection without reselection');
    info.repeatEvent=await key(p,'F4','F4');eq(info.repeatEvent?.prevented,true,'F4 is handled once');await verifyAxis(p,axis,op,n,{repeats:2});eq((await state(p)).undo,2,'F4 repeats the selected line count without reselection');await expectLineSelection(p,info.before,'F4 preserves the full selected line range and active cell');
    await run(p,'undo');eq((await book(p)).sheets,after.sheets,'Undo repeated structure returns to the first edit');await expectLineSelection(p,info.before,'Undo repeated structure preserves full-line selection');
  });
  for(const axis of ['row','col'])for(const n of [1,3])await test(axis+'-'+n+'-consecutive-CtrlShift-delete-selection-UndoRedo-F4',async(p,info)=>{
    await fixture(p);await selectAxis(p,axis,6,n,{ui:true});const selected=await state(p),initial=await book(p),snapshots=[initial];info.before=selected;info.steps=[];
    eq(selected.kind,axis==='row'?'rows':'cols','Real header selects full lines for consecutive deletion');
    for(let i=1;i<=3;i++){
      const event=await structural(p,'delete'),current=await state(p);info.steps.push({step:i,event,state:current});
      // Capture the second physical shortcut before checking the first selection, so the baseline records its erroneous popup.
      eq(current.dialogs,0,'Consecutive whole-line shortcut never opens a delete popup');
      await verifyAxis(p,axis,'delete',n,{repeats:i});eq(current.undo,initial.undo+i,'Each consecutive key creates exactly one history item');eq(current.redo,0,'A fresh deletion clears the redo stack');
      snapshots.push(await book(p));
      if(i>=2){await expectLineSelection(p,selected,'Consecutive deletion retains line kind, range and active cell');eq(lineSelection(info.steps[0].state),lineSelection(selected),'The first deletion also retained the exact full-line selection');}
    }
    const repeatEvent=await key(p,'F4','F4'),repeated=await state(p);info.steps.push({step:4,event:repeatEvent,state:repeated});eq(repeatEvent?.prevented,true,'F4 repeats deletion without reselection');eq(repeated.undo,initial.undo+4,'Three shortcuts and F4 produce four atomic history items');
    await verifyAxis(p,axis,'delete',n,{repeats:4});await expectLineSelection(p,selected,'F4 retains full-line selection');snapshots.push(await book(p));
    for(let i=3;i>=0;i--){await run(p,'undo');const current=await book(p);eq(current.sheets,snapshots[i].sheets,'Undo restores exact rows/columns, dimensions, styles, comments and formula references');eq(current.undo,initial.undo+i,'Undo removes exactly one deletion history item');eq(current.redo,4-i,'Undo exposes the exact remaining redo history');await expectLineSelection(p,selected,'Undo preserves the selected full-line range and active cell');}
    for(let i=1;i<=4;i++){await run(p,'redo');const current=await book(p);eq(current.sheets,snapshots[i].sheets,'Redo restores exact structural data and formula references');eq(current.undo,initial.undo+i,'Redo restores exactly one deletion history item');eq(current.redo,4-i,'Redo consumes exactly one history item');await expectLineSelection(p,selected,'Redo restores the full selected line range and active cell');}
    info.final=await state(p);
  });
  for(const axis of ['row','col'])await test(axis+'-merged-boundary-consecutive-delete-selection-UndoRedo',async(p,info)=>{
    await fixture(p);const initialMerge=axis==='row'?{r1:9,c1:3,r2:13,c2:4}:{r1:2,c1:9,r2:3,c2:13};
    await p.evaluate(m=>{const w=tabula.wb();w.transact(()=>w.setSheetProp(0,'merges',[m]));w.undoStack=[];w.redoStack=[];tabula.gv().renderAll();},initialMerge);await raf(p);
    await selectAxis(p,axis,6,3,{ui:true});const selected=await state(p),before=await book(p);info.before=selected;info.initialMerge=initialMerge;
    eq(selected.kind,axis==='row'?'rows':'cols','Merge fixture begins with exact whole-line selection');eq(axis==='row'?selected.sel.r2-selected.sel.r1+1:selected.sel.c2-selected.sel.c1+1,3,'Exactly three lines are initially selected');
    await structural(p,'delete');await verifyAxis(p,axis,'delete',3);await expectLineSelection(p,selected,'The first deletion keeps three whole lines beside the shifted merge');
    const first=await book(p),shifted=axis==='row'?{r1:6,c1:3,r2:10,c2:4}:{r1:2,c1:6,r2:3,c2:10};info.first=await state(p);eq(first.sheets[0].merges,[shifted],'The five-line merge shifts into the three-line selection');
    eq(first.undo,1,'The first merged-boundary deletion is a single history item');await run(p,'undo');eq((await book(p)).sheets,before.sheets,'Undo restores the original merge and complete worksheet');await expectLineSelection(p,selected,'Undo keeps the exact three-line selection');
    await run(p,'redo');eq((await book(p)).sheets,first.sheets,'Redo restores the shifted merge and complete worksheet');info.afterFirstRedo=await state(p);await expectLineSelection(p,selected,'Redo must not expand the three-line history selection to the five-line merge');
    const event=axis==='row'?await structural(p,'delete'):await key(p,'F4','F4');info.secondEvent=event;eq(event?.prevented,true,'The second physical delete or F4 is consumed');await verifyAxis(p,axis,'delete',3,{repeats:2});await expectLineSelection(p,selected,'The second deletion preserves three whole lines');
    const second=await book(p),remaining=axis==='row'?{r1:6,c1:3,r2:7,c2:4}:{r1:2,c1:6,r2:3,c2:7};eq(second.sheets[0].merges,[remaining],'The second deletion removes exactly three lines from the intersecting merge');eq(second.undo,2,'The second deletion adds one atomic history item');
    for(const snapshot of [first,before]){await run(p,'undo');eq((await book(p)).sheets,snapshot.sheets,'Undo restores each exact merge/data/formula state');eq((await book(p)).undo,snapshot.undo,'Undo removes one merged-boundary edit');await expectLineSelection(p,selected,'Each Undo keeps the exact whole-line selection');}
    for(const snapshot of [first,second]){await run(p,'redo');eq((await book(p)).sheets,snapshot.sheets,'Redo restores each exact merge/data/formula state');eq((await book(p)).undo,snapshot.undo,'Redo restores one merged-boundary edit');await expectLineSelection(p,selected,'Each Redo keeps the exact whole-line selection');}
    info.final=await state(p);
  });
  for(const axis of ['row','col'])for(const op of ['insert','delete'])await test(axis+'-CtrlShift-Numpad-'+op,async(p,info)=>{
    await fixture(p);await selectAxis(p,axis,6,2);const before=await book(p);info.event=await structural(p,op,{numpad:true});await verifyAxis(p,axis,op,2);await history(p,before,await book(p));
  });
  for(const op of ['insert','delete'])await test('cell-CtrlShift-'+op+'-dialog-cancel-preserves-borders',async(p,info)=>{
    await fixture(p);await p.evaluate(()=>tabula.selectCell(6,6));await p.locator('#cellEditor').focus();const before=await book(p),selected=await state(p);await structural(p,op);
    const dialog=p.getByRole('dialog',{name:op==='insert'?'삽입':'삭제',exact:true});await dialog.waitFor();eq(await dialog.getByRole('radio').count(),4,'Cell dialog offers directional and full-line choices');eq(await book(p),before,'Opening creates no Undo and changes no borders');eq((await state(p)).sel,selected.sel,'Opening keeps cell selection');await p.keyboard.press('Escape');eq(await book(p),before,'Cancel preserves all cell borders and contents');
  });
  for(const op of ['insert','delete'])await test('cell-CtrlShift-'+op+'-directional-shift-and-UndoRedo',async(p,info)=>{
    await fixture(p);await p.evaluate(()=>tabula.selectRange({r1:6,c1:3,r2:7,c2:4},'cells',{r:6,c:3}));await p.locator('#cellEditor').focus();const before=await book(p);await structural(p,op);
    const dialog=p.getByRole('dialog',{name:op==='insert'?'삽입':'삭제',exact:true});await dialog.waitFor();await dialog.getByRole('radio',{name:op==='insert'?/아래로 밀기/:/왼쪽으로 밀기/}).check();await dialog.getByRole('button',{name:/^확인/}).click();await raf(p);
    if(op==='insert'){eq(await raw(p,6,3),'','Selected inserted cell is blank');eq(await raw(p,8,3),'R6C3','Insert down shifts exactly two selected rows');eq(await raw(p,6,2),'R6C2','Insert down preserves other columns');}
    else{eq(await raw(p,6,3),'R6C5','Delete left shifts exactly two selected columns');eq(await raw(p,5,3),'R5C3','Delete left preserves other rows');}
    await history(p,before,await book(p));
  });
  for(const focus of ['#cellEditor','#gridView','#ribbon [data-ribbon-command="copy"]','#quickAccess button:not(:disabled)'])await test('plain-Ctrl-browser-zoom-delegation-focus-'+focus,async(p,info)=>{
    await fixture(p);await selectAxis(p,'row',6,2);await p.locator(focus).first().focus();const before=await book(p),s=await state(p);info.events=[];
    for(const [keys,code]of [['Control+Minus','Minus'],['Control+Equal','Equal'],['Control+NumpadAdd','NumpadAdd'],['Control+NumpadSubtract','NumpadSubtract']]){
      const e=await key(p,keys,code);info.events.push(e);eq(e?.prevented,false,'Plain Ctrl +/- is not consumed by the app');eq(await book(p),before,'Plain Ctrl preserves workbook and history');eq((await state(p)).zoom,s.zoom,'Plain Ctrl leaves application grid zoom unchanged');eq((await state(p)).dialogs,0,'Plain Ctrl opens no structure dialog');eq((await state(p)).sel,s.sel,'Plain Ctrl preserves the selection');
    }
    for(const data of [{key:'+',code:'Equal',ctrlKey:true},{key:'-',code:'Minus',metaKey:true}]){const e=await synthetic(p,data);info.events.push(e);eq(e.prevented,false,'Logical plus/Meta minus remain unhandled without Shift');eq(await book(p),before,'Fallback browser gesture preserves data and history');eq((await state(p)).zoom,s.zoom);}
  });
  for(const focus of ['#ribbon [data-ribbon-command="copy"]','#quickAccess button:not(:disabled)'])for(const op of ['insert','delete'])await test('row-CtrlShift-'+op+'-relay-focus-'+focus,async(p,info)=>{
    await fixture(p);await selectAxis(p,'row',6,2);await p.locator(focus).first().focus();const before=await book(p);info.before=await state(p);info.event=await structural(p,op);await verifyAxis(p,'row',op,2);await history(p,before,await book(p));
  });
  for(const axis of ['row','col'])for(const op of ['insert','delete'])await test(axis+'-protected-CtrlShift-'+op+'-and-F4-atomic-denial',async(p,info)=>{
    await fixture(p);await selectAxis(p,axis,6,1);await structural(p,op);await selectAxis(p,axis,6,2);
    await p.evaluate(()=>{const w=tabula.wb();w.sheets[tabula.si].protect={on:true,allow:{selectLocked:true,insertRows:false,deleteRows:false,insertColumns:false,deleteColumns:false,formatCells:false}};w.undoStack=[];w.redoStack=[];});
    const before=await book(p);info.event=await structural(p,op);eq(await book(p),before,'Protected structure shortcut changes neither model nor Undo');await p.keyboard.press('Escape');await p.locator('#cellEditor').focus();await key(p,'F4','F4');eq(await book(p),before,'F4 rechecks protection rather than repeating stale authority');
  });
  await test('copied-ten-rows-CtrlShift-insert-F4-and-UndoRedo',async(p,info)=>{
    await fixture(p);await selectAxis(p,'row',2,10);await p.keyboard.press('Control+c');info.copy=await p.evaluate(()=>__nativeCopy());eq(info.copy.htmlRows,10,'Copy retains ten complete source rows');
    await selectAxis(p,'row',20,1);const before=await book(p);await structural(p,'insert');eq(await raw(p,20,0),'R2C0','Inserted copy begins with original source row');eq(await raw(p,29,0),'R11C0','All ten copied rows insert');eq(await raw(p,30,0),'R20C0','Original target moves down exactly ten');eq((await state(p)).sel.r2-(await state(p)).sel.r1+1,10,'Inserted copy selects all ten rows');const after=await book(p);await history(p,before,after);
    await p.locator('#cellEditor').focus();await key(p,'F4','F4');eq(await raw(p,20,0),'R2C0','F4 reuses the immutable copied source');eq(await raw(p,40,0),'R20C0','F4 inserts ten more rows');eq(await raw(p,2,0),'R2C0','Original source stays unchanged');eq((await state(p)).undo,2,'Copied insertion and repeat are two atomic edits');await run(p,'undo');eq((await book(p)).sheets,after.sheets);
  });
  for(const target of ['#cellEditor','#formulaInput','#searchBox'])await test('editing-and-native-input-isolation-'+target,async(p,info)=>{
    await fixture(p);await selectAxis(p,'row',6,2);await p.locator(target).focus();if(target==='#cellEditor')await p.keyboard.press('F2');await p.locator(target).fill('temporary input');const before=await book(p);
    for(const op of ['insert','delete']){const g=gesture(op);info.event=await key(p,g.keys,g.code);eq(await book(p),before,'Input focus cannot perform a structural edit');eq(await p.locator(target).inputValue(),'temporary input','Shortcut preserves current input text');eq((await state(p)).dialogs,0,'Input shortcut opens no structure dialog');}
    await p.keyboard.press('Escape');eq(await book(p),before,'Cancelling input leaves the model unchanged');
  });
  await test('dialog-and-menu-isolate-CtrlShift-structure',async(p,info)=>{
    await fixture(p);await selectAxis(p,'row',6,2);const before=await book(p);await p.keyboard.press('Control+g');let dialog=p.getByRole('dialog').last();await dialog.waitFor();await dialog.locator('input').first().fill('A1');
    for(const op of ['insert','delete']){const g=gesture(op);await key(p,g.keys,g.code);eq(await book(p),before,'Dialog focus never relays structure keys');eq(await p.locator('.dialog').count(),1,'Structure key opens no nested dialog');}
    await p.keyboard.press('Escape');await p.locator('#cellEditor').focus();await p.keyboard.press('Shift+F10');await p.locator('.menu').first().waitFor();
    for(const op of ['insert','delete']){const g=gesture(op);await key(p,g.keys,g.code);eq(await book(p),before,'Context menu focus never relays structure keys');eq(await p.locator('.dialog').count(),0,'Menu key opens no structure dialog');}await p.keyboard.press('Escape');
  });
  await test('composition-and-Process229-do-not-edit-worksheet-structure',async(p,info)=>{
    await fixture(p);await selectAxis(p,'row',6,2);const before=await book(p);
    for(const data of [{key:'Process',code:'Equal',ctrlKey:true,shiftKey:true,keyCode:229},{key:'Process',code:'Minus',ctrlKey:true,shiftKey:true,keyCode:229},{key:'+',code:'Equal',ctrlKey:true,shiftKey:true,isComposing:true},{key:'_',code:'Minus',ctrlKey:true,shiftKey:true,isComposing:true}]){info.event=await synthetic(p,data);eq(await book(p),before,'Composition/229 preserves every worksheet line');eq((await state(p)).dialogs,0,'Composition/229 opens no structure dialog');}
    await p.evaluate(()=>{const ed=document.getElementById('cellEditor');ed.dispatchEvent(new CompositionEvent('compositionstart',{bubbles:true,data:''}));ed.value='한글 입력';ed.dispatchEvent(new InputEvent('input',{bubbles:true,inputType:'insertCompositionText',data:'한글 입력',isComposing:true}));});eq((await state(p)).editing,true,'Actual composition event starts normal text editing');
    await synthetic(p,{key:'Process',code:'Minus',ctrlKey:true,shiftKey:true,keyCode:229,isComposing:true});eq(await p.locator('#cellEditor').inputValue(),'한글 입력','Composing text remains intact');eq(await book(p),before,'Composing structure gesture does not commit an edit');await p.keyboard.press('Escape');
  });
  await test('selected-shape-CtrlShift-structure-does-not-change-sheet',async(p,info)=>{
    await fixture(p);await p.evaluate(()=>{const w=tabula.wb();w.transact(()=>w.setSheetProp(0,'shapes',[{id:'structure-shape',kind:'rect',x:380,y:45,w:140,h:70,fill:'#e0e7ff'}]));w.undoStack=[];w.redoStack=[];tabula.gv().renderAll();});await raf(p);await p.locator('.obj[data-id="structure-shape"]').first().click({position:{x:3,y:3}});const before=await book(p);info.selected=await state(p);eq(info.selected.object,'structure-shape','Actual shape selection is active before the shortcut');eq(info.selected.editing,false,'Shape test starts outside text editing');
    for(const op of ['insert','delete']){const g=gesture(op);await key(p,g.keys,g.code);eq(await book(p),before,'Object selection excludes worksheet structure commands');eq((await state(p)).dialogs,0,'Object shortcut opens no cell structure dialog');}
  });
}finally{
  await browser.close();const summary={url,engine,baseline:baseline||null,syntheticOnly:true,simulation:'Playwright keyboard/header input and mock native copy delivery. Browser or OS zoom itself is not asserted.',assets:[...assets],cases:results.length,passed:results.filter(r=>r.ok).length,checks,pageErrors:errors,remoteWrites:writes,blockedRequests:blocked,results};await writeFile(path.join(out,'result.json'),JSON.stringify(summary,null,2));console.log(JSON.stringify({url,engine,cases:summary.cases,passed:summary.passed,checks,pageErrors:errors.length,remoteWrites:writes.length,assets:summary.assets,out}));if(!results.length||results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;
}
