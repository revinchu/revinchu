import assert from 'node:assert/strict';
import { mkdir } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/',browser=await chromium.launch(),results=[];
if(!['localhost','127.0.0.1'].includes(new URL(url).hostname))throw Error('로컬 테스트 서버를 사용하세요.');
const out=process.env.CREATIVE_OUT||'D:/Codex/Temp/wixel-creative';await mkdir(out,{recursive:true});
async function test(name,fn,mobile=false){
 const ctx=await browser.newContext({viewport:mobile?{width:390,height:844}:{width:1500,height:1000}}),p=await ctx.newPage(),errors=[];
 p.setDefaultTimeout(10000);p.on('pageerror',e=>errors.push(e.message));
 await ctx.route('**/*',r=>new URL(r.request().url()).origin===new URL(url).origin&&!new URL(r.request().url()).pathname.startsWith('/api/')&&['GET','HEAD'].includes(r.request().method())?r.continue():r.abort());
 await ctx.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.__downloads=[];window.showSaveFilePicker=async options=>({name:options.suggestedName,createWritable:async()=>({write:async blob=>window.__downloads.push({name:options.suggestedName,type:blob.type,data:Array.from(new Uint8Array(await blob.arrayBuffer()))}),close:async()=>{},abort:async()=>{}})});});
 try{await p.goto(url);await p.waitForFunction(()=>window.tabula?.wb());await p.evaluate(()=>{const t=tabula,w=t.wb();w.restore({sheets:[{name:'합성',cells:{'0,0':{raw:'항목'},'0,1':{raw:'금액'},'1,0':{raw:'한국어'},'1,1':{raw:'12500'},'2,0':{raw:'=1+1'}}}]});t.switchSheet(0);w.undoStack=[];w.redoStack=[];t.selectRange({r1:0,c1:0,r2:5,c2:2});});await fn(p);assert.deepEqual(errors,[]);results.push({name,ok:true});console.log('OK '+name);}catch(e){results.push({name,ok:false,error:e.message});console.error('NG '+name+' '+e.stack);}finally{await ctx.close();}
}
const run=(p,c)=>p.evaluate(c=>tabula.run(c),c),shapes=p=>p.evaluate(()=>structuredClone(tabula.wb().sheets[0].shapes));
async function stroke(p,y=360){await p.mouse.move(280,y);await p.mouse.down();await p.mouse.move(330,y+35,{steps:5});await p.mouse.move(400,y+10,{steps:5});await p.mouse.up();}
try{
 await test('팔레트 펜 두 획·서식·한 획 Undo·Esc 다음 셀 선택 불변',async p=>{
  await run(p,'drawingPalette');await stroke(p);await stroke(p,440);const s=await shapes(p);assert.equal(s.length,2);assert.equal(s[0].stroke,'#185c45');assert.equal(s[0].strokeWidth,3);assert.equal(s[0].lineCap,'rnd');
  await p.screenshot({path:out+'/palette.png'});await p.keyboard.press('Escape');const before=await p.evaluate(()=>JSON.stringify(tabula.wb().sheets[0].cells));await p.mouse.click(610,610);assert.equal(await p.evaluate(()=>JSON.stringify(tabula.wb().sheets[0].cells)),before);
  await run(p,'undo');assert.equal((await shapes(p)).length,1);
 });
 await test('형광펜·지우개는 도형만 변경하며 Undo 복원',async p=>{
  await run(p,'drawingPalette');await p.getByRole('button',{name:'형광펜',exact:true}).click();await stroke(p);const s=await shapes(p);assert.equal(s.length,1);assert.equal(s[0].strokeOpacity,.35);assert.equal(s[0].strokeWidth,12);
  await p.getByRole('button',{name:'획 지우개',exact:true}).click();await p.locator(`.obj[data-id="${s[0].id}"]`).first().click();assert.equal((await shapes(p)).length,0);await run(p,'undo');assert.equal((await shapes(p)).length,1);
 });
 await test('교차색상 적용·제거·Undo와 값 보존',async p=>{
  await run(p,'alternatingColors');const d=p.getByRole('dialog',{name:'교차색상',exact:true});await d.getByRole('button',{name:'오션',exact:true}).click();await d.getByRole('checkbox',{name:'바닥글',exact:true}).check();await p.screenshot({path:out+'/banding.png'});await d.getByRole('button',{name:'적용',exact:true}).click();assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].cond.length),4);assert.equal(await p.evaluate(()=>tabula.wb().getRaw(0,1,1)),'12500');
  await run(p,'alternatingColors');await p.getByRole('button',{name:'교차색상 제거',exact:true}).click();assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].cond.length),0);await run(p,'undo');assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].cond.length),4);
 });
 await test('HTML 다운로드 저장창·한글·계산값·조건부색·실행코드 미포함',async p=>{
  await run(p,'alternatingColors');await p.getByRole('button',{name:'적용',exact:true}).click();await run(p,'exportHtml');await p.waitForFunction(()=>window.__downloads.length===1);const file=await p.evaluate(()=>window.__downloads[0]);assert.ok(file.name.endsWith('.html'));const html=new TextDecoder().decode(new Uint8Array(file.data));assert.ok(html.includes('한국어'));assert.ok(html.includes('>2<'));assert.ok(html.includes('#185c45')||html.includes('rgb(24, 92, 69)'));assert.ok(html.includes('Content-Security-Policy'));assert.ok(!/<script[ >]/i.test(html));
 });
 await test('PDF 다운로드 저장창과 실제 PDF 바이트',async p=>{
  await run(p,'exportPdf');await p.getByRole('dialog',{name:'PDF 다운로드',exact:true}).waitFor();await p.getByRole('button',{name:'PDF 파일 저장',exact:true}).click();await p.waitForFunction(()=>window.__downloads.length===1,null,{timeout:30000});const file=await p.evaluate(()=>window.__downloads[0]);assert.equal(new TextDecoder().decode(new Uint8Array(file.data.slice(0,8))),'%PDF-1.4');assert.ok(file.name.endsWith('.pdf'));
 });
 await test('전체화면 진입·종료',async p=>{
  await p.locator('[data-ribbon-tab="view"]').click();await p.locator('[data-ribbon-command="fullScreen"]').click();await p.locator('#exitFullScreen').waitFor();assert.ok(await p.evaluate(()=>!!document.fullscreenElement||document.body.classList.contains('wixel-fullscreen')));await p.locator('#exitFullScreen').click();await p.waitForFunction(()=>!document.fullscreenElement&&!document.body.classList.contains('wixel-fullscreen'));assert.equal(await p.evaluate(()=>!!document.fullscreenElement||document.body.classList.contains('wixel-fullscreen')),false);
 });
 await test('터치 펜 연속 획과 취소는 셀을 바꾸지 않음',async p=>{
  await run(p,'drawingPalette');const cdp=await p.context().newCDPSession(p),before=await p.evaluate(()=>JSON.stringify(tabula.wb().sheets[0].cells));
  const touch=async(type,points)=>cdp.send('Input.dispatchTouchEvent',{type,touchPoints:points});
  for(let n=0;n<2;n++){await touch('touchStart',[{x:320,y:400+n*100}]);await touch('touchMove',[{x:390,y:440+n*100}]);await touch('touchEnd',[]);}
  assert.equal((await shapes(p)).length,2);await touch('touchStart',[{x:500,y:500}]);await touch('touchMove',[{x:550,y:550}]);await touch('touchCancel',[]);assert.equal((await shapes(p)).length,2);assert.equal(await p.evaluate(()=>JSON.stringify(tabula.wb().sheets[0].cells)),before);
 });
 await test('파일 저장창 취소는 파일·문서 변경 없이 종료',async p=>{
  const before=await p.evaluate(()=>JSON.stringify(tabula.wb().sheets));await p.evaluate(()=>{window.showSaveFilePicker=async()=>{throw new DOMException('Cancelled','AbortError');};});await run(p,'exportHtml');assert.equal(await p.evaluate(()=>window.__downloads.length),0);assert.equal(await p.evaluate(()=>JSON.stringify(tabula.wb().sheets)),before);
 });
 await test('390px 그리기 팔레트·교차색상 버튼 접근',async p=>{
  await run(p,'drawingPalette');const rect=await p.locator('.drawing-palette').boundingBox();assert.ok(rect.x>=0&&rect.x+rect.width<=391);await p.screenshot({path:out+'/mobile-palette.png'});await p.getByRole('button',{name:'그리기 팔레트 닫기',exact:true}).click();await run(p,'alternatingColors');const rect2=await p.getByRole('button',{name:'적용',exact:true}).boundingBox();assert.ok(rect2.x>=0&&rect2.y+rect2.height<=844);
 },true);
}finally{await browser.close();}
console.log(JSON.stringify({total:results.length,passed:results.filter(r=>r.ok).length,failed:results.filter(r=>!r.ok).length,results},null,2));if(results.some(r=>!r.ok))process.exitCode=1;
