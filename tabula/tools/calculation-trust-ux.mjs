import assert from 'node:assert/strict';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const browser=await chromium.launch(),page=await browser.newPage({viewport:{width:1440,height:1000}}),errors=[];
page.on('pageerror',e=>errors.push(e.message));
try{
 await page.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true});
 await page.goto(process.env.WIXEL_URL||'http://127.0.0.1:5180/',{waitUntil:'domcontentloaded',timeout:60000});await page.waitForFunction(()=>window.tabula?.wb());
 await page.evaluate(()=>{const t=window.tabula,w=t.wb();w.restore({sheets:[{name:'검증',fileValues:true,cells:{'0,0':{raw:'2'},'0,1':{raw:'=UNSUPPORTED_FN(A1)',cached:20},'0,2':{raw:'=SUM(B1,5)',cached:25}}}]});t.selectCell(0,1);t.gv().renderAll();});
 assert.match(await page.locator('#calcState').innerText(),/파일 저장값/);
 await page.locator('#calcState').click();const dlg=page.getByRole('dialog',{name:'계산 상태 확인'});await dlg.waitFor();assert.match(await dlg.innerText(),/UNSUPPORTED_FN/);await dlg.getByRole('button',{name:'닫기',exact:true}).last().click();
 await page.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setInput(0,0,0,'3'));t.selectCell(0,1);t.gv().renderAll();});
 assert.match(await page.locator('#calcState').innerText(),/재계산 불가/);assert.deepEqual(await page.evaluate(()=>[window.tabula.wb().getValue(0,0,1).code,window.tabula.wb().getValue(0,0,2).code]),['#NAME?','#NAME?']);
 await page.locator('#calcState').click();await dlg.getByRole('button',{name:'검증!B1',exact:true}).click();assert.equal(await dlg.count(),0);assert.equal(await page.evaluate(()=>window.tabula.active.c),1);
 await page.evaluate(()=>window.tabula.run('undo'));assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,0,1)),20);assert.deepEqual(errors,[]);
 console.log(JSON.stringify({ok:true,checks:8,pageErrors:errors}));
}finally{await browser.close();}
