// Synthetic regression only: current records, retained items and hidden selections.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5195/',out=process.env.WIXEL_PIVOT_HISTORY_OUT||'D:/Codex/Temp/wixel-pivot-history-ui';
await mkdir(out,{recursive:true});
const browser=await chromium.launch(),context=await browser.newContext({viewport:{width:1400,height:1000}}),p=await context.newPage();
const errors=[],writes=[],blocked=[],results=[];let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(v,m)=>{checks++;assert.ok(v,m);};
p.setDefaultTimeout(8000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.dismiss());
await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(u.pathname);return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/')){blocked.push(u.pathname);return r.abort();}return r.continue();});
const menu=()=>p.locator('.filter-menu.pivot-filter-menu'),state=()=>p.evaluate(()=>({filters:tabula.wb().sheets[1].pivot.filters,undo:tabula.wb().undoStack.length,cache:JSON.stringify(tabula.wb().pivotCacheItems)}));
const search=()=>menu().getByRole('searchbox'),confirm=()=>menu().getByRole('button',{name:'확인',exact:true}).click(),checkbox=n=>menu().getByRole('checkbox',{name:n,exact:true});
const object=()=>p.locator('.obj.slicer[data-id="history-sl"]:visible').first(),item=n=>object().locator('.sl-item').filter({hasText:new RegExp('^'+n+'$')}).first();
async function reset(history=true){for(let i=0;i<3;i++)await p.keyboard.press('Escape');await p.evaluate(history=>{
 const t=tabula,w=t.wb();if(t.si!==0)t.switchSheet(0);const ids=n=>Array.from({length:n},(_,i)=>'item'+String(i).padStart(3,'0'));
 const rows=[['Item','Value'],...ids(31).map((x,i)=>[x,i+1])],cells={};rows.forEach((a,r)=>a.forEach((v,c)=>cells[r+','+c]={raw:String(v)}));
 w.restore({pivotCacheItems:history?{cache:{fields:[{name:'Ignored',shared:['foreign']},{name:'Item',shared:ids(275)}]}}:{},sheets:[{name:'Source',cells},{name:'Report',cells:{},pivot:{name:'SyntheticPivot',source:'Source',range:{r1:0,c1:0,r2:31,c2:1},rows:['Item'],cols:[],pages:[],values:[{field:'Value',agg:'sum'}],filters:{Item:ids(153)},cacheItemsId:'cache',layout:'tabular',top:1,left:1},slicers:[{id:'history-sl',caption:'Items',source:{kind:'pivot',field:'Item',pivots:[{sheet:'Report',name:'SyntheticPivot'}]},x:550,y:40,w:210,h:260,showDeleted:false,hideNoData:false}]}]});
 t.switchSheet(1,false);t.run('pivotRefresh');w.undoStack=[];w.redoStack=[];t.selectCell(5,2);t.gv().layout();t.gv().renderAll();
 },history);}
async function open(){await p.locator('.pbtn[data-k="rows"][data-f="Item"]:visible').first().click();await menu().waitFor();}
async function test(name,fn){try{await reset();await fn();results.push({name,ok:true});console.log('OK '+name);}catch(e){results.push({name,ok:false,error:e.stack});console.error('NG '+name+' '+e.stack);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});}}
try{
 await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());
 for(const history of[true,false])await test('31 current / 153 selected unchanged confirm; retained history '+history,async()=>{
  await reset(history);const before=await state();eq(before.filters.Item.length,153);await open();eq(await checkbox('item000').isChecked(),true);await confirm();eq(await state(),before,'no filter change, no undo entry, immutable cache');
 });
 await test('search add ghost item preserves selected history; replace is intentional',async()=>{
  const before=await state();await open();await search().fill('item200');eq(await checkbox('item200').isChecked(),true);await checkbox('필터에 현재 선택한 내용 추가').check();await confirm();let after=await state();eq(after.filters.Item.length,154);ok(after.filters.Item.includes('item152'));ok(after.filters.Item.includes('item200'));eq(after.cache,before.cache);await p.evaluate(()=>tabula.run('undo'));eq((await state()).filters,before.filters);
  await open();await search().fill('item200');await confirm();eq((await state()).filters.Item,['item200']);
 });
 await test('Ctrl preserves hidden ghost selections; direct click and clear remain intentional',async()=>{
  const before=await state();eq(await object().locator('.sl-item').count(),31);await item('item000').click({modifiers:['Control']});let after=await state();eq(after.filters.Item.length,152);ok(after.filters.Item.includes('item152'));ok(!after.filters.Item.includes('item000'));eq(after.cache,before.cache);await p.evaluate(()=>tabula.run('undo'));eq((await state()).filters,before.filters);await item('item001').click();eq((await state()).filters.Item,['item001']);await object().locator('.sl-clear').click();eq((await state()).filters.Item,undefined);
 });
 await test('showDeleted/hideNoData change list visibility without editing selected history',async()=>{
  const before=await state();await p.evaluate(()=>{const t=tabula,s=t.wb().sheets[1].slicers[0];s.showDeleted=true;t.gv().renderObjectsAll();});
  eq(await object().locator('.sl-item').count(),275);ok(await object().locator('.sl-item').filter({hasText:/^item200$/}).count());
  await p.evaluate(()=>{const t=tabula;t.wb().sheets[1].slicers[0].hideNoData=true;t.gv().renderObjectsAll();});eq(await object().locator('.sl-item').count(),153,'selected absent items remain visible; unselected no-data items hide');eq(await object().locator('.sl-item').filter({hasText:/^item200$/}).count(),0);await p.evaluate(()=>{const t=tabula;t.wb().sheets[1].slicers[0].showDeleted=false;t.gv().renderObjectsAll();});eq(await object().locator('.sl-item').count(),31);eq(await state(),before);
 });
}finally{await browser.close();const report={url,cases:results.length,passed:results.filter(r=>r.ok).length,checks,pageErrors:errors,remoteWrites:writes,blockedRequests:blocked,scope:'Synthetic Chromium UI only; not a physical iPad/Safari run.',results};await writeFile(out+'/result.json',JSON.stringify(report,null,2));console.log(JSON.stringify(report));if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;}
