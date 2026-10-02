import assert from 'node:assert/strict';
import { inflateSync } from 'node:zlib';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

// 사용자 자료 없이 실제 브라우저 픽셀·틀 고정 좌표·페이지 경계만 검증한다.
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://localhost:5178/';
const output = process.env.PAGE_BREAK_OUTPUT || 'D:/Codex/Temp/wixel-page-break-preview';
await mkdir(output, { recursive: true });
function png(bytes) {
  let w,h,type; const parts=[];
  for(let p=8;p<bytes.length;){ const n=bytes.readUInt32BE(p),t=bytes.toString('ascii',p+4,p+8),d=bytes.subarray(p+8,p+8+n);if(t==='IHDR'){w=d.readUInt32BE(0);h=d.readUInt32BE(4);type=d[9];assert.equal(d[8],8);}if(t==='IDAT')parts.push(d);p+=n+12; }
  assert.ok(type===2||type===6); const bpp=type===6?4:3,raw=inflateSync(Buffer.concat(parts)),out=Buffer.alloc(w*h*bpp),stride=w*bpp;
  const paeth=(a,b,c)=>{const p=a+b-c,pa=Math.abs(p-a),pb=Math.abs(p-b),pc=Math.abs(p-c);return pa<=pb&&pa<=pc?a:pb<=pc?b:c;};let k=0;
  for(let y=0;y<h;y++){const f=raw[k++];assert.ok(f<=4);for(let x=0;x<stride;x++){const i=y*stride+x,a=x>=bpp?out[i-bpp]:0,b=y?out[i-stride]:0,c=y&&x>=bpp?out[i-stride-bpp]:0;out[i]=(raw[k++]+(f===0?0:f===1?a:f===2?b:f===3?Math.floor((a+b)/2):paeth(a,b,c)))&255;}}
  return {at(x,y){return [...out.subarray((Math.floor(y)*w+Math.floor(x))*bpp,(Math.floor(y)*w+Math.floor(x))*bpp+3)];}};
}
const browser=await chromium.launch(),context=await browser.newContext({viewport:{width:1400,height:950},deviceScaleFactor:1}),page=await context.newPage();
page.setDefaultTimeout(15000); const errors=[],writes=[],results=[];
page.on('pageerror',e=>errors.push(e.message));
await context.route('**/*',route=>{const r=route.request(),u=new URL(r.url());if(!['GET','HEAD','OPTIONS'].includes(r.method())){writes.push(r.method()+' '+u.pathname);return route.abort();}return u.origin===new URL(base).origin&&!u.pathname.startsWith('/api/')?route.continue():route.abort();});
await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
const frame=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
async function reset(options={}) {
  await page.evaluate(o=>{
    const t=tabula,w=t.wb(); if(t.si!==0)t.switchSheet(0);
    const cells={};for(let r=0;r<80;r++)for(let c=0;c<24;c++)cells[r+','+c]={raw:'',style:{fill:'#ffe699'}};
    cells['15,14']={raw:'인쇄 밖 내용',style:{fill:'#ff00ff',color:'#000000'}};
    const s={name:'쪽 나누기 합성',cells,defRowH:22,defColW:64,zoom:o.zoom||100,view:{mode:o.mode||'pageBreakPreview',top:0,left:0,r:0,c:0},
      page:o.page||{area:{r1:1,c1:1,r2:4,c2:3}},freeze:o.freeze||null,hiddenRows:o.hiddenRows||{},hiddenCols:o.hiddenCols||{},
      shapes:[{id:'outside',kind:'rect',x:0,y:0,w:60,h:20,fill:'#ff0000',stroke:'none'}]};
    w.restore({sheets:[{name:'초기화',cells:{}},s]});w.undoStack=[];w.redoStack=[];t.switchSheet(1);
    window.__pageModel=JSON.stringify(w.serialize());
  },options);await page.evaluate(()=>document.fonts.ready);await frame();
}
async function zoom(n){
  if(n<25)await page.evaluate(v=>tabula.gv().setZoom(v),n); // 일반 UI의 현재 최소는25%; 렌더러의10%경계는 별도 검사.
  else await page.locator('#zoomSlider').evaluate((e,v)=>{e.value=String(v);e.dispatchEvent(new Event('input',{bubbles:true}));e.dispatchEvent(new Event('change',{bubbles:true}));},n);
  await frame();assert.equal(await page.evaluate(()=>tabula.gv().z),n/100);
}
async function test(name,fn){try{await fn();results.push({name,ok:true});console.log('PASS '+name);}catch(e){results.push({name,ok:false,error:e.message});console.log('FAIL '+name+' '+e.message);}}
async function geometry(){return page.evaluate(()=>{const g=tabula.gv(),p=g.panes.find(p=>p.id==='br'),b=p.el.getBoundingClientRect();return{zoom:g.z,pane:{x:b.x,y:b.y,w:b.width,h:b.height},area:g.clientRect({r1:1,c1:1,r2:4,c2:3}),plan:g.printLayout(),nodes:document.querySelectorAll('.page-preview svg *').length};});}
try {
  await page.goto(base,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForFunction(()=>window.tabula?.wb());
  for(const z of [10,25,100,200,400])await test(z+'% '+(z<25?'렌더러':'슬라이더')+' 확대에서 내용까지 회색 마스크·파란 경계·쪽 번호',async()=>{
    await reset();await zoom(z);const g=await geometry(),image=png(await page.screenshot({path:join(output,'zoom-'+z+'.png')}));
    assert.deepEqual(image.at(g.pane.x+g.pane.w-8,g.pane.y+g.pane.h-8),[182,182,182]);
    const outsideObject=await page.evaluate(()=>{const a=tabula.gv().clientRect({r1:0,c1:0,r2:0,c2:0});return[a.left+a.width/2,a.top+a.height/2];});
    assert.deepEqual(image.at(...outsideObject),[182,182,182],'인쇄 밖 빨간 개체도 가림');
    const a=g.area;let yellow=0,blue=0;
    for(let y=Math.ceil(a.top);y<Math.min(a.bottom,g.pane.y+g.pane.h);y++)for(let x=Math.ceil(a.left);x<Math.min(a.right,g.pane.x+g.pane.w);x++){const p=image.at(x,y);if(p[0]===255&&p[1]===230&&p[2]===153)yellow++;if(p[2]>150&&p[0]<50&&p[1]<100)blue++;}
    assert.ok(yellow>0,'인쇄영역 채우기 보존');assert.ok(blue>0,'실제 파란 외곽 픽셀');
    assert.equal(await page.locator('.pane-br .page-boundary[data-print-area-edge]').count(),4);assert.equal(await page.locator('.pane-br [data-page-number="1"]').count(),1);
  });
  await test('일반·페이지 레이아웃 보기에는 미리보기 마스크 없음',async()=>{
    for(const mode of ['normal','pageLayout']){await reset({mode});assert.equal(await page.locator('.page-preview svg').count(),0);}
  });
  await test('복수 인쇄영역 중첩은 합집합이며 사이는 회색',async()=>{
    await reset({page:{areas:[{r1:1,c1:1,r2:8,c2:3},{r1:4,c1:2,r2:10,c2:5},{r1:1,c1:8,r2:4,c2:10}]}});
    const sample=await page.evaluate(()=>{const g=tabula.gv();return[[5,2],[2,6],[2,8]].map(([r,c])=>{const a=g.clientRect({r1:r,c1:c,r2:r,c2:c});return[a.left+5,a.top+5];});});
    const image=png(await page.screenshot({path:join(output,'multiple-areas.png')}));
    const hasYellow=([x,y])=>{for(let dx=0;dx<10;dx++)for(let dy=0;dy<10;dy++)if(image.at(x+dx,y+dy).join(',')==='255,230,153')return true;return false;};
    assert.ok(hasYellow(sample[0]),'겹친 인쇄영역도 회색으로 가리지 않음');assert.deepEqual(image.at(...sample[1]),[182,182,182]);assert.ok(hasYellow(sample[2]));
    assert.equal((await geometry()).plan.areas.length,3);
  });
  await test('자동 쪽 경계는 점선·수동 경계는 실선 및 정확한 인덱스',async()=>{
    await reset({page:{area:{r1:0,c1:0,r2:79,c2:23},rowBreaks:[10],colBreaks:[4]}});const plan=(await geometry()).plan;
    assert.ok(plan.breaks.some(x=>x.axis==='row'&&x.index===10&&x.manual));assert.ok(plan.breaks.some(x=>x.axis==='col'&&x.index===4&&x.manual));assert.ok(plan.breaks.some(x=>!x.manual));
    assert.ok(await page.locator('[data-page-break-axis="row"][data-page-break-index="10"].manual').count());assert.ok(await page.locator('.page-boundary.automatic[stroke-dasharray]').count());
    await page.screenshot({path:join(output,'breaks.png')});
  });
  await test('숨김 행열·틀 고정·스크롤에서 경계가 시트 좌표에 붙음',async()=>{
    await reset({freeze:{rows:2,cols:2},hiddenRows:{3:true,4:true},hiddenCols:{3:true},page:{area:{r1:1,c1:1,r2:25,c2:9},rowBreaks:[8],colBreaks:[5]}});
    await page.evaluate(()=>tabula.gv().setScroll(80,60));await frame();
    const checks=await page.evaluate(()=>{const g=tabula.gv(),out=[];for(const p of g.panes){for(const n of p.pagePreview.querySelectorAll('[data-page-break-index]')){const b=n.getBoundingClientRect(),axis=n.dataset.pageBreakAxis,idx=+n.dataset.pageBreakIndex,pr=p.el.getBoundingClientRect(),origin=axis==='row'?(p.scrollY?g.boundaryY+g.sy:g.originY):(p.scrollX?g.boundaryX+g.sx:g.originX),pos=(axis==='row'?g.rows:g.cols).pos(idx),expected=(axis==='row'?pr.top:pr.left)+(pos-origin)*g.z;out.push(Math.abs((axis==='row'?b.top:b.left)-expected));}}return out;});
    assert.ok(checks.length>=2);assert.ok(checks.every(n=>n<1));await page.screenshot({path:join(output,'frozen-hidden-scroll.png')});
  });
  await test('대형 영역은 화면 크기의 SVG만 만들고 계획 상한 오류를 알림',async()=>{
    await reset({page:{area:{r1:0,c1:0,r2:999999,c2:100}}});const g=await geometry();assert.ok(g.plan.error);assert.ok(g.nodes<100);assert.equal(g.plan.pages.length,0);assert.ok(await page.locator('.page-preview-warning').count());
  });
  await test('보기 렌더링과 스크롤은 문서·실행 취소 이력을 바꾸지 않음',async()=>{
    await reset();await page.evaluate(()=>{tabula.gv().setScroll(55,77);tabula.gv().update();});await frame();
    assert.equal(await page.evaluate(()=>JSON.stringify(tabula.wb().serialize())===__pageModel),true);assert.equal(await page.evaluate(()=>tabula.wb().undoStack.length),0);
  });
  await test('인쇄영역 설정·추가·해제·Undo 및 상태 표시줄 모드 전환',async()=>{
    await reset({mode:'normal'});
    await page.evaluate(()=>{tabula.selectRange({r1:2,c1:2,r2:6,c2:5});tabula.run('printArea');});await frame();
    assert.equal(await page.evaluate(()=>tabula.wb().sheets[tabula.si].view.mode),'pageBreakPreview');
    assert.equal(await page.locator('.sheet-view-modes [data-view-mode="pageBreakPreview"]').getAttribute('aria-pressed'),'true');
    await page.evaluate(()=>{tabula.selectRange({r1:10,c1:8,r2:13,c2:10});tabula.run('addPrintArea');});await frame();assert.equal((await geometry()).plan.areas.length,2);
    await page.evaluate(()=>tabula.run('undo'));await frame();assert.equal((await geometry()).plan.areas.length,1);
    await page.evaluate(()=>tabula.run('redo'));await frame();assert.equal((await geometry()).plan.areas.length,2);
    await page.evaluate(()=>tabula.run('clearPrintArea'));await frame();assert.equal(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.area),null);assert.ok(await page.locator('.page-preview svg').count());
    await page.locator('.sheet-view-modes [data-view-mode="normal"]').click();assert.equal(await page.locator('.page-preview svg').count(),0);
    await page.locator('.sheet-view-modes [data-view-mode="pageBreakPreview"]').click();assert.ok(await page.locator('.page-preview svg').count());
  });
  await test('인쇄영역 외곽 실제 마우스 조절·Undo 및 Esc 취소',async()=>{
    await reset();
    const line=page.locator('.pane-br [data-print-area-edge="right"]'),from=await line.boundingBox(),to=await page.evaluate(()=>{const a=tabula.gv().clientRect({r1:2,c1:5,r2:2,c2:5});return{x:a.right,y:(a.top+a.bottom)/2};});
    await page.mouse.move(from.x,from.y+from.height/2);await page.mouse.down();await page.mouse.move(to.x,to.y,{steps:4});await page.mouse.up();await frame();
    assert.equal(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.area.c2),5);assert.equal(await page.evaluate(()=>tabula.wb().undoStack.length),1);
    await page.evaluate(()=>tabula.run('undo'));await frame();assert.equal(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.area.c2),3);
    const restored=await line.boundingBox();await page.mouse.move(restored.x,restored.y+restored.height/2);await page.mouse.down();await page.mouse.move(to.x,to.y,{steps:4});await page.keyboard.press('Escape');await page.mouse.up();await frame();
    assert.equal(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.area.c2),3);assert.equal(await page.locator('.print-drag-hint').count(),0);
  });
  await test('수동 쪽 나누기 이동·자동 경계 수동 전환·Undo',async()=>{
    await reset({page:{area:{r1:0,c1:0,r2:79,c2:23},rowBreaks:[10]}});
    let line=page.locator('.pane-br [data-page-break-axis="row"][data-page-break-index="10"]'),b=await line.boundingBox();
    const target=await page.evaluate(()=>{const a=tabula.gv().clientRect({r1:14,c1:2,r2:14,c2:2});return{x:(a.left+a.right)/2,y:a.top};});
    await page.mouse.move(b.x+b.width/3,b.y);await page.mouse.down();await page.mouse.move(target.x,target.y,{steps:4});await page.mouse.up();await frame();
    assert.deepEqual(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.rowBreaks),[14]);await page.evaluate(()=>tabula.run('undo'));await frame();assert.deepEqual(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.rowBreaks),[10]);
    line=page.locator('.pane-br [data-page-break-axis="col"].automatic').first();b=await line.boundingBox();
    const dest=await page.evaluate(()=>{const a=tabula.gv().clientRect({r1:3,c1:8,r2:3,c2:8});return{x:a.left,y:(a.top+a.bottom)/2};});
    await page.mouse.move(b.x,Math.max(b.y+10,dest.y));await page.mouse.down();await page.mouse.move(dest.x,dest.y,{steps:4});await page.mouse.up();await frame();
    assert.deepEqual(await page.evaluate(()=>tabula.wb().sheets[tabula.si].page.colBreaks),[8]);
  });
  await test('보호된 시트는 인쇄 경계의 드래그 손잡이를 내놓지 않음',async()=>{
    await reset();await page.evaluate(()=>{const t=tabula,w=t.wb();w.transact(()=>w.setSheetProp(t.si,'protect',{on:true}));t.gv().renderAll();});
    assert.equal(await page.locator('.page-boundary.draggable').count(),0);assert.ok(await page.locator('.page-boundary').count());
  });
  await test('경계 드래그 중 Ctrl+PageUp으로 시트를 바꾸면 양쪽 문서 설정 보존',async()=>{
    await reset();const before=await page.evaluate(()=>JSON.stringify(tabula.wb().sheets.map(s=>s.page)));
    await page.locator('#cellEditor').focus();const b=await page.locator('.pane-br [data-print-area-edge="right"]').boundingBox();
    await page.mouse.move(b.x,b.y+b.height/2);await page.mouse.down();await page.mouse.move(b.x+128,b.y+b.height/2,{steps:4});await page.keyboard.press('Control+PageUp');
    assert.equal(await page.evaluate(()=>tabula.si),0);await page.mouse.up();await frame();
    assert.equal(await page.evaluate(()=>JSON.stringify(tabula.wb().sheets.map(s=>s.page))),before);assert.equal(await page.evaluate(()=>tabula.wb().undoStack.length),0);assert.equal(await page.locator('.print-drag-hint').count(),0);
  });
  await test('기본 보기에서 페이지 나누기 삽입 후 한 번 Undo로 경계와 보기 함께 복원',async()=>{
    await reset({mode:'normal'});const before=await page.evaluate(()=>JSON.stringify(tabula.wb().sheets[tabula.si].page));
    await page.evaluate(()=>{tabula.selectCell(2,2);tabula.run('insertPageBreak');});await frame();assert.equal(await page.evaluate(()=>tabula.wb().undoStack.length),1);
    assert.deepEqual(await page.evaluate(()=>{const s=tabula.wb().sheets[tabula.si];return[s.page.rowBreaks,s.page.colBreaks,s.view.mode];}),[[2],[2],'pageBreakPreview']);
    await page.evaluate(()=>tabula.run('undo'));await frame();assert.equal(await page.evaluate(()=>JSON.stringify(tabula.wb().sheets[tabula.si].page)),before);assert.equal(await page.evaluate(()=>tabula.wb().sheets[tabula.si].view.mode),'normal');
  });
  await test('320·390 모바일 상태 표시줄의 보기·확대 버튼 접근',async()=>{
    for(const width of [320,390]){
      await page.setViewportSize({width,height:844});await page.evaluate(()=>document.body.classList.add('mobile-work-mode'));await frame();
      for(const selector of ['.sheet-view-modes [data-view-mode="normal"]','.sheet-view-modes [data-view-mode="pageBreakPreview"]','#zoomLabel']){const b=await page.locator(selector).boundingBox();assert.ok(b&&b.x>=0&&b.x+b.width<=width,selector+JSON.stringify(b));}
      await page.locator('.sheet-view-modes [data-view-mode="normal"]').click();await page.locator('.sheet-view-modes [data-view-mode="pageBreakPreview"]').click();await page.screenshot({path:join(output,'mobile-'+width+'.png')});
    }
  });
  assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);
} finally {await browser.close();await writeFile(join(output,'results.json'),JSON.stringify({passed:results.filter(r=>r.ok).length,total:results.length,results,pageErrors:errors,writes},null,2));}
console.log(JSON.stringify({passed:results.filter(r=>r.ok).length,total:results.length,pageErrors:errors,writes}));if(results.some(r=>!r.ok)||errors.length||writes.length)process.exitCode=1;
