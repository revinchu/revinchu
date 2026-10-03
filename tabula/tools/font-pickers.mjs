// 새 격리 Chromium 컨텍스트: 실제 iOS Safari 자동 확대는 기기에서 별도 확인한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
assert.ok(['127.0.0.1','localhost','[::1]'].includes(new URL(url).hostname));
const out=process.env.WIXEL_FONT_PICKER_OUT||'D:/Codex/Temp/wixel-font-pickers/source';
await mkdir(out,{recursive:true});
const browser=await chromium.launch(),results=[];let checks=0;
const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);},ok=(a,m)=>{checks++;assert.ok(a,m);};
async function test(name,fn,{mobile=true}={}){
 const context=await browser.newContext({viewport:mobile?{width:390,height:844}:{width:1440,height:1000},isMobile:mobile,hasTouch:mobile}),p=await context.newPage(),errors=[],writes=[];
 p.setDefaultTimeout(12000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.type()==='beforeunload'?d.accept():d.dismiss());
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');});
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD'].includes(q.method())){writes.push(u.pathname);return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return r.abort();return r.continue();});
 try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());await p.evaluate(()=>{const t=window.tabula;t.wb().restore({sheets:[{name:'합성',cells:{'0,0':{raw:'원본'}}},{name:'준비',cells:{}}]});t.switchSheet(1);t.switchSheet(0);t.selectCell(0,0);});await fn(p);eq(errors,[],'페이지 오류');eq(writes,[],'원격 쓰기');results.push({name,ok:true});console.log('OK '+name);}catch(e){results.push({name,ok:false,error:e.message,errors,writes});console.error('NG '+name+': '+e.stack);await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});}finally{await context.close();}
}
async function mini(p){
 await p.evaluate(()=>{window.tabula.selectCell(0,0);});
 const b=await p.evaluate(()=>{const b=window.tabula.gv().clientRect({r1:0,c1:0,r2:0,c2:0});return{x:b.left+b.width/2,y:b.top+b.height/2};});
 await p.mouse.click(b.x,b.y,{button:'right'});await p.locator('.cell-mini-toolbar').waitFor();
}
const noInputFocus=p=>p.evaluate(()=>!document.activeElement.matches('input,textarea,select'));
const style=p=>p.evaluate(()=>window.tabula.wb().styleAt(0,0,0));
try{
 await test('터치 미니 글꼴 전체 목록·선택·Undo, 입력 초점 없음',async p=>{
  await mini(p);eq(await p.locator('.mini-font').isVisible(),false);eq(await p.locator('.cell-mini-toolbar datalist').count(),0);
  await p.getByRole('button',{name:'글꼴 목록',exact:true}).last().tap();eq(await noInputFocus(p),true);
  for(const f of ['맑은 고딕','굴림','돋움','바탕','궁서','Arial','Calibri','Consolas','Times New Roman','Verdana'])eq(await p.locator('.font-item').filter({hasText:new RegExp('^'+f+'$')}).count()>0,true,f);
  await p.screenshot({path:out+'/mobile-font-list.png'});await p.locator('.font-item').filter({hasText:/^Verdana$/}).last().tap();eq((await style(p)).font,'Verdana');
  await p.evaluate(()=>window.tabula.run('undo'));eq((await style(p)).font||'맑은 고딕','맑은 고딕');
 });
 await test('터치 미니 글꼴 크기 전체16개·72선택·Undo',async p=>{
  await mini(p);eq(await p.locator('.mini-size').isVisible(),false);await p.locator('.mini-size-picker button').tap();eq(await noInputFocus(p),true);
  for(const size of [8,9,10,11,12,14,16,18,20,22,24,26,28,36,48,72])eq(await p.getByRole('menuitem',{name:String(size),exact:true}).count(),1);
  await p.getByRole('menuitem',{name:'72',exact:true}).tap();eq((await style(p)).size,72);await p.evaluate(()=>window.tabula.run('undo'));eq((await style(p)).size||11,11);
 });
 await test('명시 직접 입력은16px·소수크기·취소·유효성',async p=>{
  await mini(p);await p.locator('.mini-size-picker button').tap();await p.getByRole('menuitem',{name:'직접 입력…',exact:true}).tap();
  const dlg=p.getByRole('dialog',{name:'글꼴 크기 직접 입력',exact:true}),input=dlg.getByRole('textbox',{name:'글꼴 크기',exact:true});
  eq(await input.evaluate(n=>getComputedStyle(n).fontSize),'16px');await input.fill('500');await dlg.getByRole('button',{name:'확인',exact:true}).click();eq(await dlg.count(),1);eq((await style(p)).size||11,11);
  await input.fill('12.5');await dlg.getByRole('button',{name:'확인',exact:true}).click();eq((await style(p)).size,12.5);
  await mini(p);await p.locator('.mini-size-picker button').tap();await p.getByRole('menuitem',{name:'직접 입력…',exact:true}).tap();const fd=p.getByRole('dialog',{name:'글꼴 크기 직접 입력',exact:true});await fd.getByRole('textbox').fill('38');await fd.getByRole('button',{name:'취소',exact:true}).click();eq((await style(p)).size,12.5);
 });
 await test('목록 이후 보호상태 변경은 적용 거절',async p=>{
  await mini(p);await p.locator('.mini-size-picker button').tap();await p.evaluate(()=>{const w=window.tabula.wb();w.sheets[0].protect={on:true,allow:{}};});await p.getByRole('menuitem',{name:'24',exact:true}).tap();eq((await style(p)).size||11,11);
 });
 await test('모바일 리본 글꼴·크기 목록은 inputfocus 없이 전체목록',async p=>{
  const fb=p.locator('#ribbon .font-caret');await fb.scrollIntoViewIfNeeded();const box=await fb.boundingBox(),wrap=await p.locator('#ribbon .font-box').boundingBox();ok(Math.abs(box.width-wrap.width)<=1,'글꼴 단추가 상자 전체 너비 사용');ok(box.width>=90,'모바일 글꼴명 표시 너비');await fb.tap();eq(await noInputFocus(p),true);eq(await p.locator('.font-search').evaluate(n=>getComputedStyle(n).fontSize),'16px');ok(await p.locator('.font-item').count()>=10);await p.locator('.font-item').filter({hasText:/^Verdana$/}).last().tap();eq((await style(p)).font,'Verdana');
  await p.locator('#ribbon .font-size-picker button').tap();eq(await noInputFocus(p),true);eq(await p.getByRole('menuitem',{name:'8',exact:true}).count(),1);eq(await p.getByRole('menuitem',{name:'72',exact:true}).count(),1);await p.getByRole('menuitem',{name:'24',exact:true}).tap();eq((await style(p)).size,24);
 });
 await test('미니 전체글꼴은 문서글꼴 포함·16px 검색으로 임의글꼴 입력',async p=>{
  await p.evaluate(()=>{const w=window.tabula.wb();w.sheets[0].cells.setRC(1,0,{v:'합성',raw:'합성',style:{font:'합성 문서 글꼴'}});});
  await mini(p);await p.locator('.mini-font-picker button').tap();eq(await noInputFocus(p),true);eq(await p.locator('.font-item').filter({hasText:/^합성 문서 글꼴$/}).count(),1);
  const input=p.locator('.font-search');eq(await input.evaluate(n=>getComputedStyle(n).fontSize),'16px');await input.fill('합성 새 글꼴');await input.press('Enter');eq((await style(p)).font,'합성 새 글꼴');await p.evaluate(()=>window.tabula.run('undo'));eq((await style(p)).font||'맑은 고딕','맑은 고딕');
 });
 await test('모바일 HFF·HFS 키팁은 숨김입력 대신 목록',async p=>{
  for(const [keys,selector] of [['hff','.font-menu'],['hfs','#menuLayer .menu']]){await p.keyboard.press('Alt');for(const k of keys)await p.keyboard.press(k);await p.locator(selector).waitFor();eq(await noInputFocus(p),true);await p.keyboard.press('Escape');}
 });
 await test('데스크톱 미니 직접입력·리본 소수입력 Enter/blur 중복없음',async p=>{
  await mini(p);const font=p.locator('.mini-font'),size=p.locator('.mini-size');eq(await font.isVisible(),true);await font.fill('Consolas');await font.press('Enter');eq((await style(p)).font,'Consolas');await size.fill('14.5');await size.press('Enter');eq((await style(p)).size,14.5);await p.keyboard.press('Escape');await p.evaluate(()=>window.tabula.run('undo'));eq((await style(p)).size||11,11);eq((await style(p)).font,'Consolas');
  const ri=p.locator('#ribbon input.font-size');await ri.fill('17.5');await ri.press('Enter');eq((await style(p)).size,17.5);await p.evaluate(()=>window.tabula.run('undo'));eq((await style(p)).size||11,11);
 },{mobile:false});
}finally{await browser.close();const result={cases:results.length,passed:results.filter(x=>x.ok).length,checks,results,limit:'Chromium touch simulation; no native iPhone/Safari run'};await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify(result));if(results.some(x=>!x.ok))process.exitCode=1;}
