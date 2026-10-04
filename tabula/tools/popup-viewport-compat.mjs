import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = process.env.WIXEL_URL || 'http://127.0.0.1:5195';
const out = process.env.WIXEL_POPUP_VIEWPORT_OUT || 'D:/Codex/Temp/wixel-browser-quality-20261004/popup-viewport-' + engine;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch({ headless: true });
const results = [], pageErrors = [], remoteWrites = [];
let checks = 0;
function check(condition, message) { checks++; assert.ok(condition, message); }
async function settle(p) { await p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r)))); }
async function viewport(p, values) { await p.evaluate(v => { Object.assign(window.__popupViewport, v); window.visualViewport.dispatchEvent(new Event('resize')); window.visualViewport.dispatchEvent(new Event('scroll')); }, values); await settle(p); }
async function bounds(p, s) { return p.locator(s).last().evaluate(e => { const r=e.getBoundingClientRect(); return {x:r.x,y:r.y,right:r.right,bottom:r.bottom,width:r.width,height:r.height,scrollHeight:e.scrollHeight,clientHeight:e.clientHeight}; }); }
async function fits(p, s) {
  const [r,v]=await Promise.all([bounds(p,s),p.evaluate(()=>window.__popupViewport)]);
  check(r.x>=v.offsetLeft+3 && r.y>=v.offsetTop+3 && r.right<=v.offsetLeft+v.width-3 && r.bottom<=v.offsetTop+v.height-3,s+' outside visual viewport: '+JSON.stringify({r,v})); return r;
}
async function run(name, body) {
  const context=await browser.newContext({viewport:{width:900,height:700}});
  await context.addInitScript(()=>{
    window.TABULA_STATIC=true; window.WIXEL_SKIP_START=true; localStorage.setItem('wixel.mobile-work.v1','off');
    window.__popupViewport={width:900,height:700,offsetLeft:0,offsetTop:0};
    const v=new EventTarget(); for(const key of Object.keys(window.__popupViewport))Object.defineProperty(v,key,{get:()=>window.__popupViewport[key]});
    Object.defineProperty(window,'visualViewport',{value:v});
  });
  const p=await context.newPage();p.on('pageerror',e=>pageErrors.push(String(e)));
  await p.route('**/*',route=>{
    const r=route.request(),t=new URL(r.url());
    if(!['GET','HEAD','OPTIONS'].includes(r.method())){remoteWrites.push(r.method()+' '+t.origin+t.pathname);return route.abort();}
    if(t.origin!==new URL(url).origin&&!['data:','blob:'].includes(t.protocol))return route.abort(); return route.continue();
  });
  try{
    await p.goto(url);await p.waitForFunction(()=>window.tabula);
    await p.evaluate(async()=>{window.__popupUI=await import('/src/ui.js');});
    check(!await p.locator('body').evaluate(e=>e.classList.contains('mobile-work-mode')),'desktop density required');
    await body(p);results.push({name,ok:true});console.log('OK',name);
  }catch(e){results.push({name,ok:false,error:String(e.stack||e)});console.error('NG',name,String(e.message||e));}
  finally{await context.close();}
}
try{
  await run('desktop menu respects pinch and keyboard viewport; last action remains reachable',async p=>{
    await viewport(p,{width:350,height:210,offsetLeft:180,offsetTop:100});
    await p.evaluate(()=>{window.__popupActivated=-1;window.__popupUI.openMenu({x:750,y:600},Array.from({length:30},(_,i)=>({label:'합성 메뉴 '+i,action:()=>{window.__popupActivated=i;}})),{minWidth:420});});
    const r=await fits(p,'.menu');check(r.scrollHeight>r.clientHeight,'long menu remains scrollable');
    await p.keyboard.press('End');await p.keyboard.press('Enter');check(await p.evaluate(()=>window.__popupActivated)===29,'last keyboard action reachable');
  });
  await run('submenu and parent track viewport; authored minimum returns after expansion',async p=>{
    await p.evaluate(()=>window.__popupUI.openMenu({x:250,y:40},[{label:'하위 메뉴',submenu:Array.from({length:20},(_,i)=>({label:'항목 '+i,action(){}}))}],{minWidth:380}));
    const original=await bounds(p,'.menu');check(original.width>=380,'desktop authored minimum');
    await p.keyboard.press('ArrowRight');check(await p.locator('.menu').count()===2,'submenu opens');
    await viewport(p,{width:310,height:180,offsetLeft:120,offsetTop:130});await fits(p,'.menu');
    await p.keyboard.press('Escape');check(await p.locator('.menu').count()===1,'escape returns to parent');await fits(p,'.menu');
    await viewport(p,{width:900,height:700,offsetLeft:0,offsetTop:0});
    check(Math.abs((await bounds(p,'.menu')).width-original.width)<=1,'authored width restored');
    await p.keyboard.press('Escape');check(await p.locator('.menu').count()===0,'escape closes parent');
  });
  await run('tracked popup uses visible viewport outside mobile density',async p=>{
    await p.evaluate(()=>{const n=document.createElement('div');n.id='viewportTrackedPopup';Object.assign(n.style,{width:'500px',height:'400px',background:'white'});document.body.append(n);window.__popupUI.trackPopupPosition(n,{x:700,y:550});});
    await viewport(p,{width:330,height:200,offsetLeft:200,offsetTop:120});await fits(p,'#viewportTrackedPopup');
    await p.locator('#viewportTrackedPopup').evaluate(e=>e.remove());await settle(p);check(await p.locator('#viewportTrackedPopup').count()===0,'popup removal cleanup');
  });
  await run('safe-area limits and original desktop width survive density toggles',async p=>{
    await p.evaluate(()=>{document.documentElement.style.setProperty('--safe-left','24px');document.documentElement.style.setProperty('--safe-bottom','32px');window.__popupUI.openMenu({x:0,y:680},Array.from({length:25},(_,i)=>({label:'영역 항목 '+i,action(){}})),{minWidth:380});});
    let r=await fits(p,'.menu');check(r.x>=28 && r.bottom<=664,'device safe-area is respected');
    await p.evaluate(()=>document.body.classList.add('mobile-work-mode'));
    await viewport(p,{width:290,height:230,offsetLeft:70,offsetTop:95});await fits(p,'.menu');
    await p.evaluate(()=>document.body.classList.remove('mobile-work-mode'));
    await viewport(p,{width:900,height:700,offsetLeft:0,offsetTop:0});
    r=await fits(p,'.menu');check(r.width>=380,'desktop minimum restored after density toggle');
  });
  await run('desktop context toolbar and menu fit above keyboard without overlap',async p=>{
    await viewport(p,{width:350,height:260,offsetLeft:90,offsetTop:110});
    await p.locator('#cellEditor').focus();await p.keyboard.press('Shift+F10');
    const menu=await fits(p,'.menu'),toolbar=await fits(p,'.context-mini-toolbar');
    check(toolbar.bottom<=menu.y-4,'toolbar remains above menu');
    await p.keyboard.press('Escape');check(await p.locator('.menu,.context-mini-toolbar').count()===0,'context toolbar closes with menu');
  });
  await run('submenu authored height cap returns when keyboard closes',async p=>{
    await p.evaluate(()=>window.__popupUI.openMenu({x:30,y:40},[{label:'하위 메뉴',submenu:Array.from({length:30},(_,i)=>({label:'항목 '+i,action(){}}))}]));
    await p.keyboard.press('ArrowRight');const original=await bounds(p,'.menu');check(original.height<=421,'authored 60vh height cap');
    await viewport(p,{width:300,height:170,offsetLeft:130,offsetTop:190});await fits(p,'.menu');
    await viewport(p,{width:900,height:700,offsetLeft:0,offsetTop:0});
    check(Math.abs((await bounds(p,'.menu')).height-original.height)<=1,'authored height cap restored');
  });
}finally{await browser.close();}
const result={engine,url,scope:'Real browser DOM/key interaction; visualViewport geometry simulated. Physical iPad keyboard not tested.',cases:results.length,passed:results.filter(x=>x.ok).length,checks,pageErrors,remoteWrites,results};
await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify({engine,cases:result.cases,passed:result.passed,checks,pageErrors:pageErrors.length,remoteWrites:remoteWrites.length}));
if(results.some(x=>!x.ok)||pageErrors.length||remoteWrites.length)process.exitCode=1;
