// 표준 개체 링크 UI 회귀. 새 브라우저와 합성 문서만 사용하고 원격 쓰기는 차단한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/',browser=await chromium.launch(),results=[];
const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
async function test(name,fn,mobile=false){
 const context=await browser.newContext({viewport:mobile?{width:390,height:844}:{width:1280,height:900},hasTouch:mobile,isMobile:mobile}),p=await context.newPage(),errors=[],writes=[];
 p.setDefaultTimeout(10000);p.setDefaultNavigationTimeout(60000);p.on('pageerror',e=>errors.push(e.message));
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method());return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return r.abort();return r.continue();});
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;window.__openedLinks=[];window.open=(...args)=>{__openedLinks.push(args);return null;};});
 try{await p.goto(url,{waitUntil:'domcontentloaded'});await p.waitForFunction(()=>window.tabula?.wb());await p.evaluate(png=>{
  const t=window.tabula,shape=(id,y,hyperlink)=>({id,kind:'roundRect',name:id,text:id,x:20,y,w:150,h:34,fill:'#217346',color:'#ffffff',hyperlink});
  t.wb().restore({names:[{name:'분석위치',sheet:null,ref:"='O''Brien 보고서'!B8:D10"}],sheets:[{name:'길잡이',cells:{'0,0':{raw:'원본'}},shapes:[shape('jump',30,{target:"#'O''Brien 보고서'!$C$40",tooltip:'보고서로 이동'}),shape('name',80,{target:'#분석위치'}),shape('bad',130,{target:'#없는시트!A1'}),shape('web',180,{target:'https://example.com/report?a=1&b=2'}),shape('unsafe',230,{target:'javascript:window.__bad=true'}),{id:'group',kind:'group',name:'group',x:220,y:30,w:150,h:80,groupSize:{w:150,h:80},groupItems:[{...shape('leaf',0,{target:"#'O''Brien 보고서'!E12"}),x:0}]}],images:[{id:'picture',name:'picture',x:220,y:150,w:80,h:50,src:png,hyperlink:{target:"#'O''Brien 보고서'!F15"}}]},{name:"O'Brien 보고서",cells:{'39,2':{raw:'목적지'}},shapes:[shape('top',30,{target:"#'O''Brien 보고서'!A1"})]}]});t.switchSheet(0);t.selectCell(0,0);t.gv().renderAll();
 },png);await fn(p);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);results.push({name,ok:true});console.log('OK '+name);}catch(e){results.push({name,ok:false,error:e.message,errors,writes});console.error('NG '+name+': '+e.stack);}finally{await context.close();}
}
const object=(p,id)=>p.locator(`.obj[data-id="${id}"]`).first();
const current=p=>p.evaluate(()=>({si:tabula.si,sel:{...tabula.sel},name:tabula.wb().sheets[tabula.si].name}));
const snapshot=p=>p.evaluate(()=>JSON.stringify(tabula.wb().serialize()));
const home=p=>p.evaluate(()=>{tabula.switchSheet(0);tabula.selectCell(0,0);tabula.gv().setScroll(0,0);tabula.gv().renderAll();});
try{
 await test('도형 클릭은 공백·작은따옴표 시트의 정확한 셀로 이동하고 문서를 바꾸지 않음',async p=>{const before=await snapshot(p);await object(p,'jump').click();const a=await current(p);assert.equal(a.name,"O'Brien 보고서");assert.equal(a.sel.r1,39);assert.equal(a.sel.c1,2);assert.equal(await snapshot(p),before);});
 await test('그림과 그룹 안 개별 도형의 링크 이동',async p=>{await object(p,'picture').click();assert.equal((await current(p)).sel.c1,5);await home(p);await object(p,'group').locator('[data-object-link]').click();assert.equal((await current(p)).sel.c1,4);assert.equal((await current(p)).sel.r1,11);});
 await test('이름 범위 링크는 전체 범위를 선택',async p=>{await object(p,'name').click();const {sel}=await current(p);assert.deepEqual([sel.r1,sel.c1,sel.r2,sel.c2],[7,1,9,3]);});
 await test('맨위로는 같은 시트 스크롤과 선택을 이동',async p=>{await p.evaluate(()=>{tabula.switchSheet(1);tabula.selectCell(200,0);const w=tabula.wb(),o=w.sheets[1].shapes[0];o.y=tabula.gv().rows.pos(201);tabula.gv().renderAll();});await object(p,'top').click();assert.equal((await current(p)).sel.r1,0);assert.equal(await p.evaluate(()=>tabula.gv().sy),0);});
 await test('Ctrl 클릭 선택·Ctrl Enter 열기·끌기 취소',async p=>{await object(p,'jump').click({modifiers:['Control']});assert.equal((await current(p)).si,0);assert.equal(await object(p,'jump').evaluate(x=>x.classList.contains('sel')),true);await p.keyboard.press('Control+Enter');assert.equal((await current(p)).si,1);await home(p);const b=await object(p,'jump').boundingBox();await p.mouse.move(b.x+30,b.y+12);await p.mouse.down();await p.mouse.move(b.x+65,b.y+12,{steps:5});await p.mouse.up();assert.equal((await current(p)).si,0);assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].shapes[0].x),20);});
 await test('Ctrl K 링크 편집·Undo와 우클릭 제거·Undo',async p=>{await object(p,'jump').click({modifiers:['Control']});await p.keyboard.press('Control+k');const d=p.getByRole('dialog',{name:'하이퍼링크 편집',exact:true});await d.locator('input').nth(0).fill('#길잡이!D6');await d.locator('input').nth(1).fill('수정한 이동');await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].shapes[0].hyperlink.target),'#길잡이!D6');await p.evaluate(()=>tabula.run('undo'));assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].shapes[0].hyperlink.target),"#'O''Brien 보고서'!$C$40");await object(p,'jump').click({button:'right'});await p.getByText('하이퍼링크 제거',{exact:true}).click();assert.equal(await p.evaluate(()=>!!tabula.wb().sheets[0].shapes[0].hyperlink),false);await p.evaluate(()=>tabula.run('undo'));assert.equal(await p.evaluate(()=>!!tabula.wb().sheets[0].shapes[0].hyperlink),true);});
 await test('잘못된 내부 링크는 이동·새 이름 생성 없이 안내',async p=>{const before=await snapshot(p);await object(p,'bad').click();assert.equal((await current(p)).si,0);assert.equal(await snapshot(p),before);assert.match(await p.locator('#toast').innerText(),/시트|참조/);});
 await test('외부 링크는 새 창·격리 옵션, 실행 주소는 차단',async p=>{await object(p,'web').click();assert.deepEqual(await p.evaluate(()=>__openedLinks),[['https://example.com/report?a=1&b=2','_blank','noopener,noreferrer']]);await object(p,'unsafe').click();assert.equal(await p.evaluate(()=>__openedLinks.length),1);assert.equal(await p.evaluate(()=>!!window.__bad),false);});
 await test('보호된 시트도 링크 탐색은 허용하고 개체 링크 수정은 차단',async p=>{await p.evaluate(()=>tabula.wb().setSheetProp(0,'protect',{on:true,allow:{}}));const before=await snapshot(p);await object(p,'jump').click();assert.equal((await current(p)).si,1);await home(p);await object(p,'jump').click({modifiers:['Control']});await p.keyboard.press('Control+k');assert.equal(await p.getByRole('dialog',{name:'하이퍼링크 편집',exact:true}).count(),0);assert.equal(await snapshot(p),before);});
 await test('도형·그림·그룹 자식의 # 없는 내부 관계와 External fragment 이동',async p=>{
  for(const id of ['jump','picture','group'])for(const mode of [null,'Internal','External']){
   await home(p);
   await p.evaluate(({id,mode})=>{const w=tabula.wb(),sh=w.sheets[0],obj=id==='picture'?sh.images[0]:id==='group'?sh.shapes.find(x=>x.id===id).groupItems[0]:sh.shapes.find(x=>x.id===id);obj.hyperlink={target:(mode==='External'?'#':'')+"'O''Brien 보고서'!C12",...(mode?{targetMode:mode}:{})};tabula.gv().renderAll();},{id,mode});
   const before=await snapshot(p);
   const target=id==='group'?object(p,id).locator('[data-object-link]'):object(p,id);
   await target.click();
   const out=await current(p);
   assert.equal(out.si,1,id+' '+mode);assert.deepEqual(out.sel,{r1:11,c1:2,r2:11,c2:2},id+' '+mode);
   assert.equal(await snapshot(p),before,'링크 이동은 문서를 수정하지 않음');
  }
 });
 await test('전체 열·행 링크는 선택 종류를 유지하고 스크롤 영역을 백만 행으로 늘리지 않음',async p=>{
  await p.evaluate(()=>{const g=tabula.gv(),base=g.ensureExtentFor;window.__extentCalls=[];g.ensureExtentFor=function(r,c){__extentCalls.push([r,c]);return base.call(this,r,c);};});
  for(const [target,kind,expected] of [['#길잡이!C:E','cols',[0,2,19999999,4]],['#길잡이!4:6','rows',[3,0,5,16383]]]){
   await home(p);
   await p.evaluate(target=>{tabula.wb().sheets[0].shapes[0].hyperlink={target};tabula.gv().renderAll();window.__extentCalls=[];},target);
   const before=await p.evaluate(()=>({r:tabula.gv().extR,c:tabula.gv().extC,doc:JSON.stringify(tabula.wb().serialize())}));
   await object(p,'jump').click();
   const out=await p.evaluate(()=>({sel:tabula.sel,kind:tabula.gv().host.state().selKind,r:tabula.gv().extR,c:tabula.gv().extC,calls:__extentCalls}));
   assert.deepEqual([out.sel.r1,out.sel.c1,out.sel.r2,out.sel.c2],expected);assert.equal(out.kind,kind);
   assert.ok(out.r<=Math.max(before.r,1000),'백만 행까지 스크롤 범위를 확장하지 않음');assert.ok(out.c<=Math.max(before.c,100),'전체 열까지 스크롤 범위를 확장하지 않음');
   assert.ok(out.calls.every(([r,c])=>r<1000&&c<100),'growTo는 실제 이동할 첫 셀만 사용');
   assert.equal(await snapshot(p),before.doc);
  }
 });
 await test('같은 시트 Ctrl Enter 이동은 개체 선택을 해제하여 Delete가 셀에만 작용',async p=>{
  await p.evaluate(()=>{const w=tabula.wb();w.sheets[0].shapes[0].hyperlink={target:'#길잡이!D6'};w.transact(()=>w.setInput(0,5,3,'셀만 지움'));tabula.gv().renderAll();});
  const objectsBefore=await p.evaluate(()=>JSON.stringify({shapes:tabula.wb().sheets[0].shapes,images:tabula.wb().sheets[0].images}));
  await object(p,'jump').click({modifiers:['Control']});assert.equal(await p.evaluate(()=>tabula.chartSel),'jump');
  await p.keyboard.press('Control+Enter');
  assert.deepEqual((await current(p)).sel,{r1:5,c1:3,r2:5,c2:3});assert.equal(await p.evaluate(()=>tabula.chartSel),null);assert.equal(await p.locator('.obj.sel').count(),0);
  await p.keyboard.press('Delete');
  assert.equal(await p.evaluate(()=>tabula.wb().getRaw(0,5,3)),'');
  assert.equal(await p.evaluate(()=>JSON.stringify({shapes:tabula.wb().sheets[0].shapes,images:tabula.wb().sheets[0].images})),objectsBefore);
  await p.keyboard.press('Control+z');assert.equal(await p.evaluate(()=>tabula.wb().getRaw(0,5,3)),'셀만 지움');
 });
 await test('누른 링크는 Esc로 취소되고 마우스를 놓아도 이동하지 않음',async p=>{
  const before=await snapshot(p),box=await object(p,'jump').boundingBox();
  await p.mouse.move(box.x+40,box.y+16);await p.mouse.down();await p.keyboard.press('Escape');await p.mouse.up();
  assert.equal((await current(p)).si,0);assert.deepEqual((await current(p)).sel,{r1:0,c1:0,r2:0,c2:0});assert.equal(await snapshot(p),before);
  await object(p,'jump').click();assert.equal((await current(p)).si,1,'취소 뒤 다음 일반 클릭은 정상 이동');
 });
 await test('모바일 한 번 탭으로 바로가기 이동',async p=>{await object(p,'jump').tap();assert.equal((await current(p)).si,1);assert.equal((await current(p)).sel.r1,39);},true);
}finally{await browser.close();console.log(JSON.stringify({url,total:results.length,good:results.filter(x=>x.ok).length,bad:results.filter(x=>!x.ok).length,results},null,2));if(results.some(x=>!x.ok))process.exitCode=1;}
