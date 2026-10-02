// 모바일 셸/CSS의 모든 리본(상황별 포함) 기하 검증. 메뉴/격자 제스처는 별도 도구 담당.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['127.0.0.1','localhost'].includes(new URL(url).hostname)) throw Error('로컬 격리 검사만 허용합니다.');
const browser = await chromium.launch();
const context = await browser.newContext({ viewport:{width:390,height:844}, isMobile:true, hasTouch:true });
const page = await context.newPage(), errors=[], writes=[], results=[], appResults=[];
page.on('pageerror',e=>errors.push(e.message));
await context.route('**/*',route=>{const r=route.request(),u=new URL(r.url()); if(!['GET','HEAD','OPTIONS'].includes(r.method())){writes.push(r.method());return route.abort();}return u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?route.continue():route.abort();});
await context.addInitScript(()=>{window.WIXEL_SKIP_START=true;window.TABULA_STATIC=true;});
try {
  await page.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForFunction(()=>window.tabula?.wb());
  if(process.env.MOBILE_LAYOUT_DRAFT==='1') await page.addStyleTag({content:readFileSync(new URL('../.local/mobile-layout.css',import.meta.url),'utf8')});
  for(const [width,height] of [[320,568],[390,844],[768,1024],[844,390],[568,320]]) {
    await page.setViewportSize({width,height}); await page.evaluate(()=>{localStorage.removeItem('wixel.mobile-work.v1');localStorage.removeItem('wixel.mobile-density.v1');});
    await page.reload({waitUntil:'domcontentloaded'}); await page.waitForFunction(()=>window.tabula?.wb());
    const automatic=width<=1100;
    assert.equal(await page.locator('body').evaluate(n=>n.classList.contains('mobile-work-mode')),automatic,`자동 감지 ${width}x${height}`);
    if(!automatic) await page.locator('#mobileModeToggle').click();
    await page.waitForFunction(()=>document.body.classList.contains('mobile-work-mode'));
    await page.evaluate(()=>{
      const w=tabula.wb();w.transact(()=>{while(w.sheets.length<8) w.addSheet(`분석 보고서 ${w.sheets.length+1}`);w.setInput(0,0,0,'모바일 실제 앱 합성');w.setInput(0,1,0,'=1+1');});tabula.selectCell(0,0);
      window.__mobileBookBefore=JSON.stringify(w.serialize());
    });
    assert.equal(await page.locator('body').evaluate(n=>n.classList.contains('mobile-compact')),true,'기본 모바일 밀도 촘촘하게');
    const densityMetric=async()=>page.evaluate(()=>Object.fromEntries(['.titlebar','.ribbon-tabs','.ribbon','.quick-access','#formulaRow','.sheetbar','.statusbar','#gridWrap'].map(sel=>{const n=document.querySelector(sel),b=n.getBoundingClientRect();return[sel,{w:b.width,h:b.height,font:getComputedStyle(n).fontSize}]})));
    const compact=await densityMetric();
    await page.locator('#mobileTools').click();await page.getByRole('dialog',{name:'모바일 작업 도구'}).getByRole('button',{name:'여유롭게',exact:true}).click();
    await page.waitForFunction(()=>!document.body.classList.contains('mobile-compact'));const comfortable=await densityMetric();
    assert.ok(compact['#gridWrap'].h>comfortable['#gridWrap'].h,`격자 공간 증가 ${width}x${height}`);
    await page.screenshot({path:`D:/Codex/Temp/wixel-mobile-comfortable-${width}x${height}.png`});
    await page.locator('#mobileTools').click();await page.getByRole('dialog',{name:'모바일 작업 도구'}).getByRole('button',{name:'촘촘하게',exact:true}).click();
    await page.waitForFunction(()=>document.body.classList.contains('mobile-compact'));
    assert.equal(await page.locator('#formulaInput').evaluate(n=>getComputedStyle(n).fontSize),'16px');
    const header=await page.locator('.titlebar').boundingBox();
    for(const id of ['mobileModeToggle','mobileTools']) {const b=await page.locator(`#${id}`).boundingBox();assert.ok(b.x>=0&&b.x+b.width<=width+1&&b.width>=31&&b.height>=31,`${id} ${JSON.stringify(b)}`);}
    assert.ok(header.width<=width+1);
    const tabs=await page.locator('[data-ribbon-tab]').evaluateAll(nodes=>nodes.map(n=>n.dataset.ribbonTab).filter(v=>v!=='file'));
    let minGrid=Infinity;
    for(const id of tabs) {
      await page.locator(`[data-ribbon-tab="${id}"]`).click();
      const metric=await page.evaluate(()=>{const a=document.getElementById('app'),g=document.getElementById('gridWrap'),r=document.getElementById('ribbon');r.scrollLeft=r.scrollWidth;const last=r.querySelector('.rgroup:last-of-type')?.getBoundingClientRect(),box=r.getBoundingClientRect();return{sw:a.scrollWidth,sh:a.scrollHeight,grid:g.getBoundingClientRect().height,last:last?.right,right:box.right};});
      assert.ok(metric.sw<=width+1&&metric.sh<=height+1,`${width}x${height} ${id} 실제 앱 overflow ${JSON.stringify(metric)}`);
      assert.ok(metric.grid>=95);if(metric.last)assert.ok(metric.last<=metric.right+1);minGrid=Math.min(minGrid,metric.grid);
    }
    const lastSheet=page.locator('.sheet-tab').last();await lastSheet.click();assert.equal(await page.evaluate(()=>tabula.si),7);
    await page.locator('.sheet-tab').first().click();assert.equal(await page.evaluate(()=>tabula.si),0);
    await page.locator('#mobileTools').click();const tools=page.getByRole('dialog',{name:'모바일 작업 도구'});await tools.waitFor();
    const box=await tools.boundingBox();assert.ok(box.x>=-1&&box.y>=-1&&box.x+box.width<=width+1&&box.y+box.height<=height+1,JSON.stringify(box));
    assert.equal(await tools.locator('[data-mobile-tab="home"]').count(),1);
    await tools.getByRole('button',{name:'닫기',exact:true}).last().click();assert.equal(await tools.count(),0);
    await page.locator('#mobileModeToggle').click();await page.waitForFunction(()=>!document.body.classList.contains('mobile-work-mode'));
    await page.locator('#mobileModeToggle').click();await page.waitForFunction(()=>document.body.classList.contains('mobile-work-mode'));
    assert.equal(await page.evaluate(()=>JSON.stringify(tabula.wb().serialize())===__mobileBookBefore),true,'모드·메뉴·시트전환은 문서 불변');
    await page.locator('[data-ribbon-tab="home"]').click();await page.evaluate(()=>{document.getElementById('ribbon').scrollLeft=0;document.getElementById('ribbonTabs').scrollLeft=0;});
    await page.screenshot({path:`D:/Codex/Temp/wixel-mobile-integrated-${width}x${height}.png`});
    appResults.push({viewport:`${width}x${height}`,automatic,tabs:tabs.length,minGrid,sheets:8,modeToggle:true,toolsDialog:true,comfortableGrid:comfortable['#gridWrap'].h,compactGrid:compact['#gridWrap'].h,gainedPixels:compact['#gridWrap'].h-comfortable['#gridWrap'].h,chrome:compact});
  }
  await page.setViewportSize({width:390,height:844});
  await page.locator('#formulaInput').focus();
  await page.evaluate(()=>{Object.defineProperty(visualViewport,'height',{configurable:true,value:300});visualViewport.dispatchEvent(new Event('resize'));});
  await page.waitForFunction(()=>document.body.classList.contains('mobile-keyboard'));
  const keyboard=await page.evaluate(()=>{const app=document.getElementById('app'),g=document.getElementById('gridWrap').getBoundingClientRect(),input=document.getElementById('formulaInput').getBoundingClientRect(),view=document.getElementById('gridView').getBoundingClientRect();return{appHeight:app.getBoundingClientRect().height,scroll:app.scrollHeight,grid:g.height,view:view.height,inputBottom:input.bottom,tabs:getComputedStyle(document.getElementById('ribbonTabs')).display};});
  assert.ok(keyboard.appHeight<=301&&keyboard.scroll<=301&&keyboard.grid>=95&&keyboard.inputBottom<=301,JSON.stringify(keyboard));assert.equal(keyboard.tabs,'none');assert.ok(Math.abs(keyboard.grid-keyboard.view)<2,JSON.stringify(keyboard));
  await page.screenshot({path:'D:/Codex/Temp/wixel-mobile-keyboard-viewport.png'});
  await page.evaluate(()=>{delete visualViewport.height;document.activeElement.blur();visualViewport.dispatchEvent(new Event('resize'));});
  await page.waitForFunction(()=>!document.body.classList.contains('mobile-keyboard'));
  appResults.push({keyboardViewport:'390×300 visual viewport (emulated)',...keyboard});
  const desktop=await browser.newContext({viewport:{width:1366,height:900}});await desktop.addInitScript(()=>{window.WIXEL_SKIP_START=true;window.TABULA_STATIC=true;});
  await desktop.route('**/*',route=>{const r=route.request(),u=new URL(r.url());return ['GET','HEAD','OPTIONS'].includes(r.method())&&u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?route.continue():route.abort();});
  const dp=await desktop.newPage();dp.on('pageerror',e=>errors.push(e.message));await dp.goto(url,{waitUntil:'domcontentloaded'});await dp.waitForFunction(()=>window.tabula?.wb());
  const desktopMetrics=await dp.evaluate(()=>({mobile:document.body.classList.contains('mobile-work-mode'),compact:document.body.classList.contains('mobile-compact'),heights:Object.fromEntries(['.titlebar','.ribbon-tabs','.ribbon','.quick-access','#formulaRow','.sheetbar','.statusbar','#gridWrap'].map(s=>[s,document.querySelector(s).getBoundingClientRect().height]))}));
  assert.equal(desktopMetrics.mobile,false);assert.equal(desktopMetrics.compact,false);assert.deepEqual(desktopMetrics.heights,{'.titlebar':38,'.ribbon-tabs':32,'.ribbon':96,'.quick-access':33,'#formulaRow':28,'.sheetbar':30,'.statusbar':24,'#gridWrap':619});
  await dp.screenshot({path:'D:/Codex/Temp/wixel-mobile-density-desktop.png'});await desktop.close();appResults.push({desktop:'1366x900',...desktopMetrics});
  if (process.env.MOBILE_LAYOUT_CATALOG !== '0') {
  await page.evaluate(()=>{
    document.body.classList.add('mobile-work-mode');
    const cells={'0,0':{raw:'모바일 합성 문서'},'1,0':{raw:'=1+1'},'2,0':{raw:'클릭과 글자 크기 유지'}};
    tabula.wb().restore({sheets:Array.from({length:8},(_,i)=>({name:`분석 보고서 ${i+1}`,cells}))});tabula.selectCell(0,0);tabula.gv().layout();tabula.gv().renderAll();
  });
  // 실제 앱의 기본 탭을 먼저 검증한 뒤 같은 리본 renderer에 모든 상황을 제공한다.
  const source=await page.evaluate(async()=>{
    const {buildRibbon,TABS}=await import('/src/ribbon.js');
    const state={context:['table','pivot','spark','slicer','object','chart'],font:'맑은 고딕',size:11,numFmt:'general',printFitW:0,printFitH:0,printScale:100,chartGalleryKey:'fixture'};
    let ribbon;
    const app={run(){},openMenu(){},focusGrid(){},hiddenTabs:()=>[],refreshRibbon(){ribbon?.update(state);},gallery:()=>Array.from({length:4},(_,i)=>{const b=document.createElement('button');b.className='chart-style-chip';b.style.cssText='width:86px;height:52px;flex:none';b.textContent=`스타일 ${i+1}`;return b;})};
    ribbon=buildRibbon(app);ribbon.update(state);window.__mobileRibbon=ribbon;window.__mobileState=state;
    return TABS.filter(t=>!t.file).map(t=>({id:t.id,label:t.label}));
  });
  for (const density of ['compact','comfortable']) {
  await page.locator('#mobileTools').click();await page.getByRole('dialog',{name:'모바일 작업 도구'}).getByRole('button',{name:density==='compact'?'촘촘하게':'여유롭게',exact:true}).click();
  await page.waitForFunction(d=>document.body.classList.contains('mobile-compact')===(d==='compact'),density);
  for(const [width,height] of [[320,568],[360,740],[390,844],[768,1024],[844,390],[568,320]]) {
    await page.setViewportSize({width,height});
    await page.evaluate(({height})=>{document.body.classList.add('mobile-work-mode');document.body.classList.toggle('mobile-short',height<=600);document.documentElement.style.setProperty('--mobile-vh',`${height}px`);document.documentElement.style.setProperty('--mobile-vw',`${innerWidth}px`);tabula.gv().layout();}, {height});
    let minGrid=Infinity,controls=0;
    for(const tab of source) {
      await page.evaluate(id=>{__mobileRibbon.selectTab(id);__mobileRibbon.update(__mobileState);tabula.gv().layout();},tab.id);
      const item=page.locator(`[data-ribbon-tab="${tab.id}"]`);await item.scrollIntoViewIfNeeded();
      const metrics=await page.evaluate(()=>{
        const app=document.getElementById('app'),grid=document.getElementById('gridWrap'),ribbon=document.getElementById('ribbon'),tabs=document.getElementById('ribbonTabs');
        const groupRects=[...ribbon.querySelectorAll('.rgroup')].map(n=>({h:n.getBoundingClientRect().height,w:n.getBoundingClientRect().width}));
        const controls=[...ribbon.querySelectorAll('button,input:not([type=checkbox]),select')].filter(n=>n.getBoundingClientRect().width>0).map(n=>({name:n.title||n.textContent.trim(),w:n.getBoundingClientRect().width,h:n.getBoundingClientRect().height}));
        const last=ribbon.querySelector('.rgroup:last-of-type');ribbon.scrollLeft=ribbon.scrollWidth;
        const rb=ribbon.getBoundingClientRect(),lb=last?.getBoundingClientRect(),ab=app.getBoundingClientRect(),gb=grid.getBoundingClientRect();
        return {window:[innerWidth,innerHeight],app:{x:ab.x,y:ab.y,w:ab.width,h:ab.height,scrollWidth:app.scrollWidth,scrollHeight:app.scrollHeight},bodyWidth:document.documentElement.scrollWidth,grid:{x:gb.x,y:gb.y,w:gb.width,h:gb.height},ribbon:{left:rb.x,right:rb.right,bottom:rb.bottom,height:rb.height,client:ribbon.clientWidth,scroll:ribbon.scrollWidth,lastRight:lb?.right},tabs:{client:tabs.clientWidth,scroll:tabs.scrollWidth},groups:groupRects,controls};
      });
      assert.ok(metrics.app.w<=width+1 && metrics.app.scrollWidth<=width+1 && metrics.bodyWidth<=width+1,`${width} ${tab.id} document overflow ${JSON.stringify(metrics)}`);
      assert.ok(metrics.app.h<=height+1 && metrics.app.scrollHeight<=height+1,`${width}x${height} ${tab.id} vertical overflow ${JSON.stringify(metrics)}`);
      assert.ok(metrics.grid.h>=95,`${width}x${height} ${tab.id} grid ${metrics.grid.h}`);
      assert.ok(metrics.ribbon.lastRight<=metrics.ribbon.right+1,`${tab.id} 마지막 그룹 스크롤 접근`);
      const hit=density==='compact'?29:35; const small=metrics.controls.filter(c=>c.h<hit||c.w<hit);
      assert.deepEqual(small,[],`${width}x${height} ${tab.id} 작은 터치 대상`);
      minGrid=Math.min(minGrid,metrics.grid.h);controls+=metrics.controls.length;
    }
    await page.evaluate(()=>{__mobileRibbon.selectTab('home');__mobileRibbon.update(__mobileState);document.getElementById('ribbon').scrollLeft=0;document.getElementById('ribbonTabs').scrollLeft=0;tabula.gv().layout();});
    await page.screenshot({path:`D:/Codex/Temp/wixel-mobile-${density}-${width}x${height}.png`});
    results.push({viewport:`${width}x${height}`,density,tabs:source.length,controls,minGrid});
  }
  }
  }
  assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);
  console.log(JSON.stringify({appResults,viewports:results.length,tabLayouts:results.reduce((n,r)=>n+r.tabs,0),results,pageErrors:errors,blockedWrites:writes},null,2));
} finally {await context.close();await browser.close();}
