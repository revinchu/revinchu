// Synthetic geometry UI regression; isolated contexts, API/external writes blocked.
import assert from 'node:assert/strict';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/',browser=await chromium.launch(),results=[];
let checks=0;
const eq=(a,b,message)=>{checks++;assert.deepEqual(a,b,message);};
const ok=(value,message)=>{checks++;assert.ok(value,message);};
const run=(p,cmd)=>p.evaluate(cmd=>window.tabula.run(cmd),cmd);
const model=p=>p.evaluate(()=>structuredClone({shapes:window.tabula.wb().sheets[0].shapes,images:window.tabula.wb().sheets[0].images}));
const history=p=>p.evaluate(()=>window.tabula.wb().undoStack.length);
const area=s=>{let sum=0,start=null,last=null;for(const path of s.path.paths)for(const c of path.commands){if(c[0]==='M'){start=last=c.slice(1);}else if(c[0]==='L'){sum+=last[0]*c[2]-last[1]*c[1];last=c.slice(1);}else if(c[0]==='Z'){sum+=last[0]*start[1]-last[1]*start[0];last=start;}else throw new Error('사각형 연산에 곡선이 반환됨');}return Math.round(sum/2*s.w*s.h);};
async function select(p,ids){for(let i=0;i<ids.length;i++){const id=ids[i],position=id==='a'?{x:20,y:40}:id==='b'?{x:80,y:40}:{x:20,y:20};await p.locator(`.obj[data-id="${id}"]`).first().click({position,modifiers:i?['Control']:[]});}}
async function menu(p,label,context=false,first='a'){
  if(context){await p.locator(`.obj[data-id="${first}"]`).first().click({button:'right',position:first==='b'?{x:80,y:40}:{x:20,y:40}});await p.locator('.menu-item').filter({hasText:/^도형 병합/}).hover();}
  else {await p.locator('[data-ribbon-tab="objFormat"]').click();await p.locator('[data-ribbon-menu="shapeMerge"]').click();}
  await p.locator('.menu-item').filter({hasText:new RegExp('^'+label+'\\(')}).click();
}
async function test(name,fn){
  if(process.env.WIXEL_SHAPE_MERGE_FILTER&&!name.includes(process.env.WIXEL_SHAPE_MERGE_FILTER))return;
  const context=await browser.newContext({viewport:{width:1500,height:1000},acceptDownloads:false}),p=await context.newPage(),errors=[],writes=[];p.setDefaultTimeout(10000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.dismiss());
  await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method());return r.abort();}return u.origin===new URL(url).origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
  try {
    await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());
    await p.evaluate(()=>{const t=window.tabula,w=t.wb(),svg='<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><path fill="#2244aa" d="M0 0H40V100H0Z M60 0H100V100H60Z"/></svg>';
      w.restore({sheets:[{name:'조합 합성',cells:{'0,0':{raw:'원본 보존'},'0,1':{raw:'=1+1'}},shapes:[{id:'a',kind:'rect',x:100,y:100,w:100,h:100,fill:'#ff0000',stroke:null,text:'',z:1},{id:'b',kind:'rect',x:150,y:100,w:100,h:100,fill:'#0000ff',stroke:null,text:'',z:2},{id:'untouched',kind:'rect',x:400,y:100,w:80,h:80,fill:'#00ff00',text:'보존',z:3}],images:[{id:'icon',name:'합성 SVG',x:650,y:100,w:150,h:150,src:'data:image/svg+xml;base64,'+btoa(svg),z:4}]}]});t.switchSheet(0);t.selectCell(0,0);t.gv().layout();t.gv().renderAll();w.undoStack=[];w.redoStack=[];
    });
    await fn(p);eq(await p.evaluate(()=>[window.tabula.wb().getRaw(0,0,0),window.tabula.wb().getRaw(0,0,1)]),['원본 보존','=1+1']);eq(errors,[],'페이지 오류');eq(writes,[],'외부 쓰기');results.push({name,ok:true});console.log('OK '+name);
  }catch(error){results.push({name,ok:false,error:error.message,errors,writes});console.error('NG '+name+' '+error.stack);}
  finally {await context.close();}
}
try {
  for(const [label,expected,count] of [['결합',15000,1],['병합',10000,1],['조각',15000,3],['교차',5000,1],['빼기',5000,1]])await test(label+' 실제 Ctrl선택·메뉴·UndoRedo',async p=>{
    const before=await model(p);await select(p,['a','b']);await menu(p,label,label==='결합');const after=await model(p),parts=after.shapes.filter(s=>s.id!=='untouched');
    eq(parts.length,count);eq(parts.reduce((n,s)=>n+area(s),0),expected);ok(parts.every(s=>s.fill==='#ff0000'),'첫 선택 서식');eq(after.shapes.find(s=>s.id==='untouched'),before.shapes.find(s=>s.id==='untouched'));eq(after.images,before.images);eq(await history(p),1,'한 번 Undo');
    if(label==='빼기')eq([parts[0].x,parts[0].w],[100,50]);
    await run(p,'undo');eq(await model(p),before);await run(p,'redo');eq(await model(p),after);
  });
  await test('빼기는 z순서 대신 최초 선택 도형을 기준으로 함',async p=>{await select(p,['b','a']);await menu(p,'빼기',false,'b');const part=(await model(p)).shapes.find(s=>s.id!=='untouched');eq([part.x,part.w,part.fill],[200,50,'#0000ff']);eq(area(part),5000);});
  await test('SVG 우클릭 변환→그룹해제→한 부분만 색변경·Undo',async p=>{
    const before=await model(p);await select(p,['icon']);await p.locator('.obj[data-id="icon"]').first().click({button:'right',position:{x:20,y:20}});await p.locator('.menu-item').filter({hasText:/^SVG 도형 변환/}).click();
    let m=await model(p),g=m.shapes.find(s=>s.kind==='group');ok(g);eq(g.groupItems.length,2);eq(m.images.length,0);eq(await history(p),1);await run(p,'undo');eq(await model(p),before);await run(p,'redo');await run(p,'objUngroup');m=await model(p);eq(m.shapes.some(s=>s.kind==='group'),false);
    const parts=m.shapes.filter(s=>s.id.startsWith('icon-'));eq(parts.length,2);await select(p,[parts[0].id]);await p.locator('[data-ribbon-tab="objFormat"]').click();await p.locator('[data-ribbon-menu="shapeFill"]').click();
    const swatch=p.locator('#menuLayer .swatch[title="#FF0000"],#menuLayer .swatch[title="#ff0000"]').first();await swatch.click();const colored=await model(p);eq(colored.shapes.find(s=>s.id===parts[0].id).fill.toLowerCase(),'#ff0000');eq(colored.shapes.find(s=>s.id===parts[1].id).fill,parts[1].fill);await run(p,'undo');eq(await model(p),m);
  });
  await test('SVG에서 그룹해제를 직접 실행해도 개별파트와 한 번 Undo를 유지함',async p=>{const before=await model(p);await select(p,['icon']);await run(p,'objUngroup');const after=await model(p);eq(after.images.length,0);eq(after.shapes.filter(s=>s.id.startsWith('icon-')).length,2);eq(after.shapes.some(s=>s.kind==='group'),false);eq(await history(p),1);await run(p,'undo');eq(await model(p),before);});
  await test('지원하지 않는 SVG와 선 조합 실패는 원본·Undo기록을 보존함',async p=>{
    await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.sheets[0].images[0].src='data:image/svg+xml;base64,'+btoa('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y="30">unsupported</text></svg>');t.gv().renderObjectsAll();});
    await select(p,['icon']);const before=await model(p);await run(p,'iconToShapes');eq(await model(p),before);eq(await history(p),0);ok((await p.locator('#toast').textContent()).includes('변환할 수 없습니다'));
    await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.sheets[0].shapes[1].kind='line';w.sheets[0].shapes[1].stroke='#000000';t.gv().renderObjectsAll();});await select(p,['a','b']);const original=await model(p);await run(p,'shapeUnion');eq(await model(p),original);eq(await history(p),0);
  });
  await test('보호된 시트에서 5연산·SVG변환·해제는 저장값을 바꾸지 않음',async p=>{
    await select(p,['a','b']);await p.evaluate(()=>window.tabula.wb().setSheetProp(0,'protect',{on:true,allow:{}}));const before=await model(p);
    for(const cmd of ['shapeUnion','shapeCombine','shapeFragment','shapeIntersect','shapeSubtract']){await run(p,cmd);eq(await model(p),before);await p.keyboard.press('Escape');}
    await select(p,['icon']);await p.keyboard.press('Escape');await run(p,'iconToShapes');eq(await model(p),before);await p.keyboard.press('Escape');await run(p,'objUngroup');eq(await model(p),before);await p.keyboard.press('Escape');eq(await history(p),0);
  });
}finally {await browser.close();}
console.log(JSON.stringify({total:results.length,passed:results.filter(r=>r.ok).length,checks,results},null,2));if(results.some(r=>!r.ok))process.exitCode=1;
