// 격리 합성 슬라이서: 모바일 갤러리의 실제 버튼 좌표·터치·스타일·Undo 검사.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const engines = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_SLICER_GALLERY_OUT || 'D:/Codex/Temp/wixel-mobile-slicer-gallery/source';
const only = process.env.WIXEL_SLICER_GALLERY_FILTER || '';
const names = (process.env.WIXEL_BROWSER || 'chromium,webkit').split(',');
await mkdir(out, { recursive: true });
const results = [], errors = [], writes = []; let checks = 0;
const eq = (a, b, label) => { checks++; assert.deepEqual(a, b, label); };
const ok = (value, label) => { checks++; assert.ok(value, label); };
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const fixture = () => ({ sheets: [{ name: '자료', cells: { '0,0': { raw: '지역' }, '0,1': { raw: '매출' }, '1,0': { raw: '서울' }, '1,1': { raw: '10' }, '2,0': { raw: '부산' }, '2,1': { raw: '20' } }, slicers: [{ id: 'sl1', name: '합성 슬라이서', caption: '지역', source: { kind: 'pivot', field: '지역', pivots: [{ sheet: '보고서', name: '피벗1' }] }, x: 16, y: 25, w: 190, h: 180 }] }, { name: '보고서', cells: {}, pivot: { name: '피벗1', source: '자료', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['지역'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 0, left: 0 } }] });
for (const name of names) {
  assert.ok(['chromium', 'webkit'].includes(name));
  const browser = await engines[name].launch();
  try { for (const settings of [{width:320,density:'compact'}, {width:390,density:'compact'}, {width:320,density:'comfortable'}, {width:390,density:'comfortable'}, {width:1440,density:'desktop'}]) {
    const label = `${name}-${settings.width}-${settings.density}`;
    if (only && !label.includes(only)) continue;
    const mobile = settings.density !== 'desktop';
    const context = await browser.newContext({ viewport: { width: settings.width, height: mobile ? 844 : 1000 }, hasTouch: mobile, isMobile: mobile, serviceWorkers: 'block' });
    const page = await context.newPage(), pageErrors = [], networkWrites = [];
    page.setDefaultTimeout(8000); page.on('pageerror', e => { pageErrors.push(e.message); errors.push({label,message:e.message}); });
    await context.addInitScript(({mobile,density}) => { window.TABULA_STATIC=true; window.WIXEL_SKIP_START=true; localStorage.setItem('wixel.mobile-work.v1',mobile?'on':'off'); localStorage.setItem('wixel.mobile-density.v1',density); Object.defineProperty(navigator,'clipboard',{configurable:true,value:{readText:async()=>'',writeText:async()=>{}}}); }, {mobile,density:settings.density});
    await context.route('**/*', route => { const request=route.request(),u=new URL(request.url()); if(!['GET','HEAD'].includes(request.method())){ networkWrites.push(u.pathname);writes.push({label,path:u.pathname});return route.abort(); } if(u.origin!==new URL(url).origin||u.pathname.startsWith('/api/'))return route.abort(); return route.continue(); });
    let geometry; const start=checks;
    try {
      await page.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForFunction(()=>window.tabula?.wb());
      await page.evaluate(data=>{const t=window.tabula,w=t.wb();w.restore(data);w.setSnapshots(data);t.switchSheet(1);t.switchSheet(0);t.gv().layout();t.gv().renderAll();t.selectCell(0,0);w.undoStack=[];w.redoStack=[];},fixture());
      await page.locator('.obj.slicer[data-id="sl1"]:visible .sl-cap').first().click();
      await page.locator('[data-ribbon-tab="slicerTab"]').click();await raf(page);
      const gallery=page.locator('.rgallery[data-gallery="slicerStyles"]'),more=gallery.getByRole('button',{name:'슬라이서 스타일 더 보기',exact:true}),prev=gallery.getByRole('button',{name:'이전 슬라이서 스타일',exact:true}),next=gallery.getByRole('button',{name:'다음 슬라이서 스타일',exact:true});
      await more.scrollIntoViewIfNeeded();
      geometry=await gallery.evaluate(node=>{const rect=n=>{const b=n.getBoundingClientRect();return{x:b.x,y:b.y,w:b.width,h:b.height,bottom:b.bottom,right:b.right};};return{gallery:rect(node),ribbon:rect(document.getElementById('ribbon')),controls:[...node.querySelectorAll('.rg-pages button')].map(n=>({label:n.getAttribute('aria-label'),disabled:n.disabled,...rect(n)}))};});
      for(const control of geometry.controls){ok(control.y>=geometry.ribbon.y-1&&control.bottom<=geometry.ribbon.bottom+1,'각 버튼 전체가 리본 높이 안에 있음: '+JSON.stringify(control));ok(control.h>=(mobile?20:15),'버튼은 눌러볼 수 있는 높이');}
      if(mobile)eq(new Set(geometry.controls.map(x=>Math.round(x.y))).size,1,'모바일 이전·다음·더 보기 한 줄');
      else ok(geometry.controls[0].y<geometry.controls[1].y&&geometry.controls[1].y<geometry.controls[2].y,'데스크톱 세로 페이지 버튼 보존');
      eq(await prev.isDisabled(),true,'첫 페이지 이전만 비활성');eq(await more.isEnabled(),true,'더 보기는 활성');eq(await next.isEnabled(),true,'다음 페이지 활성');
      const titles=()=>gallery.locator('.slstyle').evaluateAll(ns=>ns.map(n=>n.dataset.slicerStyle));const first=await titles();
      const activate=async locator=>{await locator.scrollIntoViewIfNeeded();if(mobile)await locator.tap();else await locator.click();};
      await activate(next);ok(JSON.stringify(await titles())!==JSON.stringify(first),'다음 견본 페이지 실제 변경');await activate(prev);eq(await titles(),first,'이전 페이지 복원');
      await more.scrollIntoViewIfNeeded();const hit=await more.evaluate(n=>{const b=n.getBoundingClientRect();return n.contains(document.elementFromPoint(b.x+b.width/2,b.y+b.height/2));});ok(hit,'더 보기 중앙 hit test');
      await page.screenshot({path:out+'/'+label+'-ribbon.png'});
      const before=await page.evaluate(()=>window.tabula.wb().serialize());
      await activate(more);const styles=page.locator('#menuLayer .slstyle');await styles.first().waitFor();ok(await styles.count()>first.length,'전체 스타일 목록 실제 열림');
      await page.screenshot({path:out+'/'+label+'-gallery.png'});
      const choice=styles.last(),selected=await choice.getAttribute('data-slicer-style');await activate(choice);
      eq(await page.evaluate(()=>window.tabula.wb().sheets[0].slicers[0].style),selected,'실제 스타일 모델에 적용');eq(await page.evaluate(()=>window.tabula.wb().undoStack.length),1,'스타일 적용 Undo 한 번');
      await page.evaluate(()=>window.tabula.run('undo'));eq(await page.evaluate(()=>window.tabula.wb().serialize()),before,'Undo 원본 문서 복원');
      await more.scrollIntoViewIfNeeded();await more.focus();await page.keyboard.press('Enter');await styles.first().waitFor();await page.keyboard.press('Escape');eq(await styles.count(),0,'키보드로 열기·Escape 닫기');eq(await page.evaluate(()=>window.tabula.wb().undoStack.length),0,'열기·닫기는 Undo 변경 없음');
      eq(pageErrors,[]);eq(networkWrites,[]);results.push({label,ok:true,checks:checks-start,geometry});console.log('OK '+label);
    } catch(error) {results.push({label,ok:false,checks:checks-start,error:error.message,geometry});console.error('NG '+label+': '+error.message);await page.screenshot({path:out+'/'+label+'-failure.png'}).catch(()=>{});}
    finally {await context.close();}
  }} finally {await browser.close();}
}
const result={url,cases:results.length,passed:results.filter(r=>r.ok).length,checks,errors,writes,simulation:'Chromium/WebKit mobile emulation; no physical iPhone verification',results};await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify(result));if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;
