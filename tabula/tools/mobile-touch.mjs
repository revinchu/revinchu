// Synthetic touch regression; fresh contexts, no user documents, remote writes blocked.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5180/',browser=await chromium.launch(),results=[];
const run=(p,cmd)=>p.evaluate(cmd=>window.tabula.run(cmd),cmd);
async function test(name,fn,{readonly=false}={}){
 const c=await browser.newContext({viewport:{width:390,height:844},deviceScaleFactor:2,isMobile:true,hasTouch:true}),p=await c.newPage(),errors=[],writes=[];
 p.setDefaultTimeout(10000);p.on('pageerror',e=>errors.push(e.message));
 await c.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method());return r.abort();}return u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
 try{
  await c.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});const target=readonly?url.replace(/[?#].*$/,'')+'#view='+gzipSync(JSON.stringify({docName:'읽기 전용 합성',workbook:{sheets:[{name:'읽기 전용',cells:{'0,0':{raw:'보존'}}}]},view:{headers:true,grid:true}})).toString('base64url'):url;await p.goto(target,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());if(readonly)await p.waitForFunction(()=>window.tabula.gv().host.state().readonly);
  await p.evaluate(()=>{const t=window.tabula,w=t.wb(),cells={'0,0':{raw:'보존'}};for(let r=1;r<100;r++)for(let c=0;c<12;c++)cells[r+','+c]={raw:r+':'+c};w.restore({sheets:[{name:'터치 합성',cells}]});t.switchSheet(0);t.selectCell(0,0);t.gv().setScroll(0,0);t.gv().renderAll();w.undoStack=[];w.redoStack=[];document.activeElement?.blur();});
  const session=await c.newCDPSession(p);await fn(p,session);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,0,0)),'보존');results.push({name,ok:true});console.log('OK '+name);
 }catch(e){results.push({name,ok:false,error:e.message,errors,writes});console.error('NG '+name+' '+e.stack);}finally{await c.close();}
}
const state=p=>p.evaluate(()=>{const t=window.tabula,g=t.gv();return {active:{...t.active},sel:{...t.sel},editing:g.host.state().editing,x:g.sx,y:g.sy,zoom:g.z,version:t.wb().version};});
async function point(p,r,c){return p.evaluate(({r,c})=>{const g=window.tabula.gv(),b=g.clientRect({r1:r,c1:c,r2:r,c2:c});return{x:(b.left+b.right)/2,y:(b.top+b.bottom)/2};},{r,c});}
async function touch(session,type,points){await session.send('Input.dispatchTouchEvent',{type,touchPoints:points.map((p,i)=>({x:p.x,y:p.y,id:i+1,radiusX:4,radiusY:4,force:1}))});}
async function swipe(session,a,b,hold=0){await touch(session,'touchStart',[a]);if(hold)await new Promise(r=>setTimeout(r,hold));for(let n=1;n<=6;n++)await touch(session,'touchMove',[{x:a.x+(b.x-a.x)*n/6,y:a.y+(b.y-a.y)*n/6}]);await touch(session,'touchEnd',[]);}
try{
 await test('탭선택·키보드억제·두번탭편집·Undo',async(p,s)=>{
  const a=await point(p,3,1);await p.touchscreen.tap(a.x,a.y);let st=await state(p);assert.deepEqual(st.active,{r:3,c:1});assert.equal(st.editing,false);assert.notEqual(await p.evaluate(()=>document.activeElement.id),'cellEditor');
  await p.waitForTimeout(400);await p.touchscreen.tap(a.x,a.y);await p.touchscreen.tap(a.x,a.y);assert.equal((await state(p)).editing,true);await p.locator('#cellEditor').fill('터치 입력');await p.keyboard.press('Enter');assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,3,1)),'터치 입력');await run(p,'undo');assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,3,1)),'3:1');
 });
 await test('한손세로·가로스크롤은선택값을바꾸지않음',async(p,s)=>{
  const b=await p.locator('#gridView').boundingBox(),before=await state(p),bottom=Math.min(770,b.y+b.height-40);
  await swipe(s,{x:180,y:bottom},{x:180,y:bottom-120});let after=await state(p);assert.ok(after.y>before.y+80);assert.deepEqual(after.sel,before.sel);assert.equal(after.editing,false);
  await swipe(s,{x:300,y:bottom-60},{x:140,y:bottom-60});after=await state(p);assert.ok(after.x>80);assert.deepEqual(after.sel,before.sel);assert.equal(after.version,before.version);
 });
 await test('길게누른뒤끌기범위선택',async(p,s)=>{
  const a=await point(p,3,1),b=await point(p,7,2);await swipe(s,a,b,500);const st=await state(p);assert.deepEqual(st.sel,{r1:3,c1:1,r2:7,c2:2});assert.equal(st.editing,false);assert.equal(await p.getByRole('menu').count(),0);
 });
 await test('핀치확대중점유지·한손가락잔여터치무시',async(p,s)=>{
  const b=await p.locator('#gridView').boundingBox(),mid={x:195,y:Math.min(650,b.y+b.height*.6)};
  await p.evaluate(()=>window.tabula.gv().setScroll(200,300));const before=await state(p),anchor=await p.evaluate(m=>window.tabula.gv().hitTest(m.x,m.y),mid);
  await touch(s,'touchStart',[{x:mid.x-50,y:mid.y},{x:mid.x+50,y:mid.y}]);await touch(s,'touchMove',[{x:mid.x-75,y:mid.y},{x:mid.x+75,y:mid.y}]);
  const after=await state(p),pos=await p.evaluate(m=>window.tabula.gv().hitTest(m.x,m.y),mid);assert.ok(Math.abs(after.zoom-before.zoom*1.5)<.011);assert.ok(Math.abs(pos.sheetX-anchor.sheetX)<2);assert.ok(Math.abs(pos.sheetY-anchor.sheetY)<2);
  await touch(s,'touchEnd',[{x:mid.x-75,y:mid.y}]);await touch(s,'touchMove',[{x:mid.x-70,y:mid.y+10}]);await touch(s,'touchEnd',[]);assert.deepEqual((await state(p)).sel,before.sel);assert.equal((await state(p)).editing,false);
  await p.screenshot({path:process.env.WIXEL_MOBILE_SCREENSHOT||'D:/Codex/Temp/wixel-mobile-touch.png'});
 });
 await test('확인란탭·보호시트편집과토글차단',async(p,s)=>{
  await p.evaluate(()=>{const t=window.tabula;t.wb().setCellData(0,2,1,{raw:'FALSE',style:{checkbox:true}});t.gv().renderAll();});let a=await point(p,2,1);await p.touchscreen.tap(a.x,a.y);assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,2,1)),true);assert.equal((await state(p)).editing,false);
  await p.evaluate(()=>window.tabula.wb().setSheetProp(0,'protect',{on:true,allow:{}}));await p.touchscreen.tap(a.x,a.y);assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,2,1)),true);assert.equal((await state(p)).editing,false);await p.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();
  a=await point(p,4,1);await p.touchscreen.tap(a.x,a.y);await p.touchscreen.tap(a.x,a.y);assert.equal((await state(p)).editing,false);assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,4,1)),'4:1');
 });
 await test('서식복사탭·범위1회적용과보호차단',async(p,s)=>{
  await p.evaluate(()=>{const t=window.tabula;t.wb().setStyle(0,1,1,{fill:'#f12345',bold:true});t.selectCell(1,1);});await run(p,'painter');let a=await point(p,3,1);await p.touchscreen.tap(a.x,a.y);assert.equal(await p.evaluate(()=>window.tabula.wb().styleAt(0,3,1).fill),'#f12345');assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,3,1)),'3:1');
  await p.evaluate(()=>window.tabula.selectCell(1,1));await run(p,'painter');const n=await p.evaluate(()=>window.tabula.wb().undoStack.length);await swipe(s,await point(p,4,1),await point(p,6,2),500);assert.equal(await p.evaluate(()=>window.tabula.wb().styleAt(0,6,2).fill),'#f12345');assert.equal(await p.evaluate(()=>window.tabula.wb().undoStack.length),n+1);
  await p.evaluate(()=>window.tabula.selectCell(1,1));await run(p,'painter');await p.evaluate(()=>window.tabula.wb().setSheetProp(0,'protect',{on:true,allow:{}}));a=await point(p,8,1);await p.touchscreen.tap(a.x,a.y);assert.notEqual(await p.evaluate(()=>window.tabula.wb().styleAt(0,8,1).fill),'#f12345');
 });
 await test('숨김행열·병합·배율후탭좌표',async(p,s)=>{
  await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.setSheetProp(0,'hiddenRows',{2:true});w.setSheetProp(0,'hiddenCols',{1:true});w.setSheetProp(0,'merges',[{r1:4,c1:2,r2:5,c2:3}]);t.gv().setZoom(125);t.gv().renderAll();});const a=await point(p,5,2);await p.touchscreen.tap(a.x,a.y);assert.deepEqual((await state(p)).active,{r:4,c:2});assert.deepEqual((await state(p)).sel,{r1:4,c1:2,r2:5,c2:3});
 });
 await test('터치중문서교체와touchcancel은편집하지않음',async(p,s)=>{
  const a=await point(p,3,1);await touch(s,'touchStart',[a]);await p.evaluate(()=>{const t=window.tabula;t.wb().restore({sheets:[{name:'새 합성',cells:{'0,0':{raw:'보존'}}}]});t.switchSheet(0);t.selectCell(0,0);});await touch(s,'touchEnd',[]);assert.deepEqual((await state(p)).active,{r:0,c:0});
  await touch(s,'touchStart',[a]);await touch(s,'touchCancel',[]);assert.equal((await state(p)).editing,false);assert.deepEqual((await state(p)).active,{r:0,c:0});
 });

 await test('셀탭직후필터단추의native터치를가로채지않음',async(p,s)=>{
  await p.evaluate(()=>{const t=window.tabula;t.wb().setSheetProp(0,'filter',{r1:0,c1:0,r2:99,c2:3,criteria:{},hidden:{}});t.gv().renderAll();});
  const a=await point(p,3,1);await p.touchscreen.tap(a.x,a.y);const b=await p.locator('.fbtn[data-c="1"]').first().boundingBox();assert.ok(b);await p.touchscreen.tap(b.x+b.width/2,b.y+b.height/2);await p.locator('.filter-menu').waitFor();assert.equal(await p.locator('.filter-menu').count(),1);
 });
 await test('읽기전용문서탭선택·편집확인란변경차단',async(p,s)=>{
  await p.evaluate(()=>{const t=window.tabula;t.wb().setCellData(0,2,1,{raw:'FALSE',style:{checkbox:true}});t.gv().renderAll();});const a=await point(p,2,1);await p.touchscreen.tap(a.x,a.y);await p.touchscreen.tap(a.x,a.y);assert.equal(await p.evaluate(()=>window.tabula.wb().getValue(0,2,1)),false);assert.equal((await state(p)).editing,false);
  const b=await point(p,4,1);await p.touchscreen.tap(b.x,b.y);await p.touchscreen.tap(b.x,b.y);assert.equal((await state(p)).editing,false);assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,4,1)),'4:1');
 },{readonly:true});
}finally{await browser.close();}
console.log(JSON.stringify({total:results.length,passed:results.filter(x=>x.ok).length,results},null,2));if(results.some(x=>!x.ok))process.exitCode=1;
