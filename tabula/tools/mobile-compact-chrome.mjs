// 합성 통합 문서와 새 Chromium 컨텍스트로 모바일 크롬·휠·우클릭을 검증합니다.
// 실제 모바일 기기/OS 메뉴 검사가 아니며 사용자 문서·클립보드·원격 API를 사용하지 않습니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_MOBILE_COMPACT_OUT || 'D:/Codex/Temp/wixel-mobile-compact/source';
const only = process.env.WIXEL_MOBILE_COMPACT_FILTER || '';
await mkdir(out, { recursive: true });
const browser = await chromium.launch(), results = [], errors = [], writes = [], blocked = [], measurements = [], assets = new Set();
let checks = 0;
const eq = (a,b,m) => { checks++; assert.deepEqual(a,b,m); }, ok = (a,m) => { checks++; assert.ok(a,m); };
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const run = (p,cmd) => p.evaluate(cmd => window.tabula.run(cmd),cmd);
const book = p => p.evaluate(() => JSON.stringify(window.tabula.wb().serialize()));
const bar = p => p.locator('.context-mini-toolbar');
const menu = p => p.locator('#menuLayer > .menu[data-level="0"]');
const sizes = [[320,640],[360,740],[390,844],[768,1024],[844,390],[1366,900]];
const metric = async p => p.evaluate(() => {
  const sels=['.titlebar','.tb-brand','#docTitle','#mobileModeToggle','#mobileTools','#ribbonTabs','[data-ribbon-tab=file]','.footer','.sheetbar','#sheetTabs','.sheet-tab','.sheet-add','.statusbar','#statusMode','#calcState','#statusStats','.sheet-view-modes','.zoom','#zoomSlider','#zoomLabel','#ribbon','.font-family','.font-size','#quickAccess','#gridWrap'];
  const rects=Object.fromEntries(sels.map(sel=>{const n=document.querySelector(sel),b=n?.getBoundingClientRect(),s=n&&getComputedStyle(n);return[sel,n?{...b.toJSON(),display:s.display,visible:b.width>0&&b.height>0&&s.visibility!=='hidden'&&s.display!=='none',font:s.fontSize,clientWidth:n.clientWidth,scrollWidth:n.scrollWidth}:null]}));
  return {rects,viewport:{width:innerWidth,height:innerHeight},mobile:window.tabula.mobile().active,density:window.tabula.mobile().density,app:{w:document.getElementById('app').scrollWidth,h:document.getElementById('app').scrollHeight}};
});
const inside = (b,w,h,name) => ok(b && b.width>0 && b.height>0 && b.left>=-1 && b.top>=-1 && b.right<=w+1 && b.bottom<=h+1, name+' '+JSON.stringify(b));
const shot = (p,name) => p.screenshot({path:out+'/'+name+'.png'});
async function test(name,size,fn,{desktop=false}={}) {
  if(only&&!name.includes(only))return;
  const [width,height]=size,c=await browser.newContext({viewport:{width,height},hasTouch:!desktop,isMobile:!desktop}),p=await c.newPage(),pageErrors=[],pageWrites=[];
  p.setDefaultTimeout(10000);p.on('pageerror',e=>{pageErrors.push(e.message);errors.push({name,error:e.message});});p.on('dialog',d=>d.type()==='beforeunload'?d.accept():d.dismiss());
  await c.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel:version','3.0.0');Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>'',writeText:async()=>{},read:async()=>[],write:async()=>{}}});});
  await c.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method()+' '+u.pathname);pageWrites.push(u.pathname);return r.abort();}if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/')){blocked.push(u.pathname);return r.abort();}return r.continue();});
  try {
    await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());
    for(const s of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(s);
    await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.restore({sheets:Array.from({length:12},(_,i)=>({name:'합성 분석 '+(i+1),zoom:125,cells:{'0,0':{raw:'보존'},'1,0':{raw:'42'},'2,1':{raw:'=1+1'}}}))});t.switchSheet(1);t.switchSheet(0);t.selectCell(0,0);});
    if(!desktop&&!await p.evaluate(()=>window.tabula.mobile().active))await p.locator('#mobileModeToggle').click();
    await raf(p);await fn(p,width,height);eq(pageErrors,[],'페이지 오류');eq(pageWrites,[],'원격 쓰기');results.push({name,ok:true});console.log('OK '+name);
  }catch(e){results.push({name,ok:false,error:e.message});console.error('NG '+name+': '+e.stack);await shot(p,'failure-'+results.length).catch(()=>{});}
  finally{await c.close();}
}
async function openContext(p) {
  await p.evaluate(()=>{window.__contextEvents=[];if(!window.__recordContext){window.__recordContext=true;document.addEventListener('contextmenu',e=>{const target=e.target;queueMicrotask(()=>window.__contextEvents.push({target:target.id||target.className,prevented:e.defaultPrevented}));});}window.tabula.selectCell(0,0);});
  const point=await p.evaluate(()=>{const r=window.tabula.gv().clientRect({r1:0,c1:0,r2:0,c2:0});return{x:r.left+r.width/2,y:r.top+r.height/2};});
  await p.mouse.click(point.x,point.y,{button:'right'});await bar(p).waitFor();await menu(p).waitFor();await raf(p);
}
try {
  for(const size of sizes){const label=size.join('×');
    await test(label+' 상단·하단 한 줄과 글꼴 입력 압축',size,async(p,w,h)=>{
      const before=await book(p),m=await metric(p),r=m.rects;measurements.push({name:label+' chrome',...m});
      eq(m.mobile,true);eq(m.density,'compact');eq(r['.tb-brand'].visible,false,'브랜드 숨김');eq(r['#docTitle'].visible,false,'문서 제목 숨김');
      const mode=r['#mobileModeToggle'],tools=r['#mobileTools'],file=r['[data-ribbon-tab=file]'];
      for(const [n,b] of [['모드',mode],['도구',tools],['파일',file]])inside(b,w,h,n);
      ok(mode.right<=tools.left+1&&tools.right<=file.left+1,'모드·도구·파일 순서');ok(Math.abs(mode.top+mode.height/2-(file.top+file.height/2))<=3,'동일 상단 행');ok(Math.max(mode.bottom,tools.bottom,file.bottom)<=36,'상단은 한 줄 높이');
      for(const sel of ['#statusMode','#calcState','#statusStats'])eq(r[sel].visible,false,sel+' 숨김');
      const sheets=r['.sheetbar'],status=r['.statusbar'];inside(sheets,w,h,'시트 영역');inside(status,w,h,'보기·배율 영역');ok(sheets.right<=status.left+1,'시트와 배율 영역 비겹침');ok(Math.abs(sheets.top+sheets.height/2-status.top-status.height/2)<=3,'동일 하단 행');ok(r['.footer'].height<=36,'하단 한 줄 높이');
      for(const sel of ['.sheet-add','.sheet-view-modes','.zoom','#zoomLabel'])inside(r[sel],w,h,sel);
      ok(r['#sheetTabs'].width>=45,'최소 시트 이름 공간');ok(r['.font-family'].width<=98&&r['.font-family'].height<=30,'글꼴 입력 크기');ok(r['.font-size'].width<=42&&r['.font-size'].height<=30,'글꼴 크기 입력 크기');
      ok(m.app.w<=w+1&&m.app.h<=h+1,'앱 가로/세로 넘침 없음');ok(r['#gridWrap'].height>=95,'격자 작업 공간');
      await shot(p,label+'-chrome');await p.locator('#mobileTools').click();const d=p.getByRole('dialog',{name:'모바일 작업 도구',exact:true});await d.waitFor();await d.getByRole('button',{name:'닫기',exact:true}).last().click();eq(await book(p),before,'화면 모드·도구는 원본 불변');
    });
    await test(label+' 실제 휠로 리본·탭·빠른 실행·시트 가로 스크롤',size,async(p,w,h)=>{
      const before=await book(p),states=[];eq(await p.locator('.sheet-tab').count(),12,'스크롤 가능한 실제 시트 탭 12개');
      for(const selector of ['#ribbonTabs','#ribbon','#quickAccess','#sheetTabs']){
        const n=p.locator(selector);await n.evaluate(n=>{n.scrollLeft=0;});const s=await n.evaluate(n=>({client:n.clientWidth,total:n.scrollWidth,left:n.scrollLeft}));
        const b=await n.boundingBox();if(selector==='#quickAccess'&&h<=600&&!b){states.push({selector,hiddenInShortViewport:true});continue;}ok(b&&b.height>0,selector+' 표시');
        if(s.total>s.client+2){
          await p.mouse.move(Math.max(2,Math.min(w-2,b.x+b.width/2)),Math.max(2,Math.min(h-2,b.y+b.height-3)));await p.mouse.wheel(0,180);await p.waitForFunction(sel=>document.querySelector(sel).scrollLeft>0,selector);
          const right=await n.evaluate(n=>n.scrollLeft);ok(right>0,selector+' 휠 오른쪽');await p.mouse.wheel(0,-180);await p.waitForFunction(({sel,right})=>document.querySelector(sel).scrollLeft<right,{sel:selector,right});states.push({selector,overflow:true,right});
        }else states.push({selector,overflow:false});
      }
      measurements.push({name:label+' wheel',states});eq(await book(p),before,'가로 스크롤은 문서 불변');await p.locator('.sheet-tab').last().click();eq(await p.evaluate(()=>window.tabula.si),11);await p.locator('.sheet-tab').first().click();eq(await p.evaluate(()=>window.tabula.si),0);
      const slider=p.locator('#zoomSlider');await slider.evaluate(n=>{n.value='110';n.dispatchEvent(new Event('input',{bubbles:true}));n.dispatchEvent(new Event('change',{bubbles:true}));});eq(await p.evaluate(()=>window.tabula.gv().z),1.1);eq(await p.evaluate(()=>window.tabula.wb().sheets[0].zoom),125,'모바일 화면 배율은 문서 배율 보존');
    });
    await test(label+' 우클릭 미니 도구 비겹침·힌트 숨김·네이티브 메뉴 차단',size,async(p,w,h)=>{
      await openContext(p);const bs=await bar(p).boundingBox(),ms=await menu(p).boundingBox();inside({...bs,left:bs.x,top:bs.y,right:bs.x+bs.width,bottom:bs.y+bs.height},w,h,'미니 도구');inside({...ms,left:ms.x,top:ms.y,right:ms.x+ms.width,bottom:ms.y+ms.height},w,h,'셀 메뉴');ok(bs.y+bs.height<=ms.y,'미니 도구와 메뉴 비겹침');ok(bs.height<=72,'미니 도구 최대 두 줄 높이');
      const controls=await bar(p).locator('input,button').evaluateAll(ns=>ns.map(n=>({label:n.getAttribute('aria-label'),...n.getBoundingClientRect().toJSON()})));for(const x of controls){ok(x.width>0&&x.height>0&&x.height<=28,'압축 미니 조작 '+x.label);ok(x.left>=bs.x-1&&x.right<=bs.x+bs.width+1&&x.top>=bs.y-1&&x.bottom<=bs.y+bs.height+1,'미니 내부 조작 '+x.label);}
      eq(await bar(p).locator('.access-key-hint').count(),0,'미니 상시 접근키 힌트 없음');await p.keyboard.press('Alt');eq(await p.locator('.access-key-badge').evaluateAll((nodes,b)=>nodes.filter(n=>{const r=n.getBoundingClientRect(),x=r.x+r.width/2,y=r.y+r.height/2;return x>=b.x&&x<=b.x+b.width&&y>=b.y&&y<=b.y+b.height;}).length,bs),0,'미니 영역 Alt 배지 없음');ok(await p.locator('.access-key-badge').count()>0,'셀 메뉴 접근키는 유지');await p.keyboard.press('Escape');
      for(const n of [bar(p),bar(p).getByLabel('글꼴',{exact:true}),menu(p).getByRole('menuitem').first()])await n.dispatchEvent('contextmenu',{bubbles:true,cancelable:true,button:2});
      eq(await p.evaluate(()=>window.__contextEvents.every(e=>e.prevented)),true,'격자·미니 입력·메뉴의 기본 contextmenu 차단');ok((await p.evaluate(()=>window.__contextEvents.length))>=4,'마우스와 합성 contextmenu 모두 기록');
      await bar(p).locator('[data-mini-command=bold]').click();eq(await p.evaluate(()=>window.tabula.wb().styleAt(0,0,0).bold),true);eq(await menu(p).count(),1,'서식 적용 후 메뉴 유지');await shot(p,label+'-mini');await p.keyboard.press('Escape');eq(await bar(p).count(),0);eq(await menu(p).count(),0);await run(p,'undo');eq(!!await p.evaluate(()=>window.tabula.wb().styleAt(0,0,0).bold),false);eq(await p.evaluate(()=>window.tabula.wb().getRaw(0,0,0)),'보존');
      await run(p,'replace');const d=p.getByRole('dialog',{name:/찾기/});await d.waitFor();const input=d.locator('input[type=text]').first();const prevented=await input.evaluate(n=>{const e=new MouseEvent('contextmenu',{bubbles:true,cancelable:true,button:2});n.dispatchEvent(e);return e.defaultPrevented;});eq(prevented,true,'대화상자 입력에서도 브라우저 메뉴 차단');await p.keyboard.press('Escape');
      measurements.push({name:label+' mini',bar:bs,menu:ms,controls});
    });
  }
  await test('데스크톱 수동 모바일 전환 후 원래 크롬·단축키 복원',[1366,900],async(p,w,h)=>{
    eq(await p.evaluate(()=>window.tabula.mobile().active),false);const before=await book(p),base=await metric(p);eq(base.rects['#docTitle'].visible,true);eq(base.rects['#mobileTools'].visible,false);
    await p.locator('#mobileModeToggle').click();await p.waitForFunction(()=>window.tabula.mobile().active);await raf(p);const on=await metric(p);eq(on.rects['#docTitle'].visible,false);ok(on.rects['.footer'].height<=36);await p.locator('#mobileModeToggle').click();await p.waitForFunction(()=>!window.tabula.mobile().active);await raf(p);const after=await metric(p);
    for(const sel of ['.titlebar','#ribbonTabs','.sheetbar','.statusbar','.footer'])eq([after.rects[sel].width,after.rects[sel].height],[base.rects[sel].width,base.rects[sel].height],sel+' 데스크톱 복원');eq(after.rects['#docTitle'].visible,true);eq(after.rects['#statusMode'].visible,true);eq(after.rects['#calcState'].visible,true);eq(after.rects['#mobileTools'].visible,false);eq(await book(p),before);await p.locator('#cellEditor').focus();await p.keyboard.press('Control+b');eq(await p.evaluate(()=>window.tabula.wb().styleAt(0,0,0).bold),true);await p.keyboard.press('Control+z');eq(!!await p.evaluate(()=>window.tabula.wb().styleAt(0,0,0).bold),false);await shot(p,'desktop-restored');
  },{desktop:true});
  await test('키보드 viewport 축소와 복원은 작은 화면·하단 조작을 유지한다',[390,844],async p=>{
    const before=await book(p);await p.locator('#formulaInput').focus();await p.evaluate(()=>{Object.defineProperty(visualViewport,'height',{configurable:true,get:()=>350});visualViewport.dispatchEvent(new Event('resize'));});await p.waitForFunction(()=>document.body.classList.contains('mobile-keyboard'));const short=await metric(p);ok(short.app.h<=351);ok(short.rects['#gridWrap'].height>=95);ok(short.rects['.footer'].height<=36);await shot(p,'keyboard-viewport');
    await p.evaluate(()=>{document.activeElement.blur();delete visualViewport.height;visualViewport.dispatchEvent(new Event('resize'));});await p.waitForFunction(()=>!document.body.classList.contains('mobile-keyboard'));const after=await metric(p);inside(after.rects['#mobileTools'],390,844,'복원 도구');inside(after.rects['.zoom'],390,844,'복원 배율');eq(await book(p),before,'입력하지 않은 focus·viewport 변경 원본 불변');
  });
}finally{
  await browser.close();const result={url,simulation:'Chromium viewport/touch/wheel/visualViewport; actual mobile device and native OS menu not tested',assets:[...assets],cases:results.length,passed:results.filter(x=>x.ok).length,checks,pageErrors:errors,remoteWrites:writes,blockedRequests:blocked,results};await writeFile(out+'/result.json',JSON.stringify(result,null,2));await writeFile(out+'/measurements.json',JSON.stringify(measurements,null,2));console.log(JSON.stringify(result));if(results.some(x=>!x.ok)||errors.length||writes.length)process.exitCode=1;
}
