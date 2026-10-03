import assert from 'node:assert/strict';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
if(!['localhost','127.0.0.1'].includes(new URL(url).hostname))throw Error('로컬 서버만 사용');
for(const engine of ['chromium','webkit']){
 const browser=await pw[engine].launch(),page=await browser.newPage({viewport:{width:1100,height:780}}),errors=[];
 page.on('pageerror',e=>errors.push(e.message));await page.route('**/api/**',r=>r.abort());
 await page.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 try{
  await page.goto(url);await page.waitForFunction(()=>window.tabula?.wb());
  await page.evaluate(()=>{const w=tabula.wb(),values=Array.from({length:50000},(_,i)=>'Item '+String(i).padStart(5,'0'));w.sheets[0].slicers=[{id:'virtual-test',caption:'Large slicer',x:20,y:20,w:340,h:260,columns:3,buttonHeight:24,gap:3,style:'SlicerStyleLight1',source:{kind:'cache',field:'item',cacheKey:'large-fixture',values,items:values.map((_,index)=>({index,hasData:true}))}}];w.version++;tabula.gv().renderAll();});
  const list=page.locator('[data-id="virtual-test"] .sl-items').first();await list.waitFor();
  assert.ok(await list.locator('.sl-item').count()<100);assert.ok(await list.evaluate(n=>n.scrollHeight>400000));
  await list.focus();await list.press('End');await page.waitForTimeout(50);
  assert.equal(await page.evaluate(()=>document.activeElement?.dataset.slIndex),'49999');assert.ok(await list.locator('.sl-item').count()<100);
  const key=i=>list.locator('[data-k="'+i+'"]');await key(49999).click();
  assert.deepEqual(await page.evaluate(()=>tabula.wb().sheets[0].slicers[0].cacheSelection),['49999']);
  assert.ok(await list.evaluate(n=>n.scrollTop>400000));await key(49998).click({modifiers:['Control']});
  assert.deepEqual((await page.evaluate(()=>tabula.wb().sheets[0].slicers[0].cacheSelection)).sort(),['49998','49999']);
  await page.evaluate(()=>tabula.run('undo'));assert.deepEqual(await page.evaluate(()=>tabula.wb().sheets[0].slicers[0].cacheSelection),['49999']);
  await list.focus();await list.press('Home');await page.waitForTimeout(50);assert.equal(await page.evaluate(()=>document.activeElement?.dataset.slIndex),'0');
  await list.press('PageDown');await page.waitForTimeout(50);assert.ok(Number(await page.evaluate(()=>document.activeElement?.dataset.slIndex))>0);
  await list.evaluate(n=>{n.scrollTop=n.scrollHeight/2;n.dispatchEvent(new Event('scroll'));});await page.waitForTimeout(60);
  assert.ok(await list.locator('.sl-item').count()<100);assert.ok(Number(await list.locator('.sl-item').first().getAttribute('data-sl-index'))>20000);
  assert.equal(await page.evaluate(()=>tabula.wb().sheets[0].slicers[0].source.values.length),50000);assert.deepEqual(errors,[]);
  console.log(JSON.stringify({engine,checks:14,ok:true,pageErrors:errors}));
 }finally{await browser.close();}
}
