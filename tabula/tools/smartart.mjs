// Isolated synthetic SmartArt browser regression; no documents or network writes.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base=process.env.WIXEL_URL || 'http://127.0.0.1:5180/', browser=await chromium.launch(), results=[];
async function test(name,fn){
  const context=await browser.newContext({viewport:{width:1500,height:1000}}),p=await context.newPage(),errors=[],writes=[];
  p.setDefaultTimeout(10000);p.on('pageerror',e=>errors.push(e.message));
  await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(!['GET','HEAD','OPTIONS'].includes(q.method())){writes.push(q.method());return r.abort();}return u.origin===new URL(base).origin&&!u.pathname.startsWith('/api/')?r.continue():r.abort();});
  try{
    await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});await p.goto(base,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb());
    await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.restore({sheets:[{name:'SmartArt 합성',cells:{'0,0':{raw:'보존'}}}]});t.switchSheet(0);t.selectCell(2,2);w.undoStack=[];w.redoStack=[];});
    await fn(p);assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);assert.equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,0,0)),'보존');results.push({name,ok:true});console.log('OK '+name);
  }catch(e){results.push({name,ok:false,error:e.message,errors,writes});console.error('NG '+name+' '+e.stack);}finally{await context.close();}
}
const dialog=p=>p.getByRole('dialog',{name:'SmartArt 그래픽',exact:true});
const shapes=p=>p.evaluate(()=>structuredClone(window.tabula.wb().sheets[0].shapes));
const run=(p,c)=>p.evaluate(c=>window.tabula.run(c),c);
async function open(p){await run(p,'insertSmartArt');await dialog(p).waitFor();return dialog(p);}
async function create(p){const d=await open(p);await d.getByRole('button',{name:'확인',exact:true}).click();return (await shapes(p))[0];}
async function edit(p){await run(p,'editSmartArt');await dialog(p).waitFor();return dialog(p);}
try{
  await test('8범주 60갤러리와 삽입 취소',async p=>{const d=await open(p),c=d.getByRole('tablist',{name:'SmartArt 범주',exact:true});assert.ok((await d.boundingBox()).width>1000,'큰 화면 SmartArt 창 폭');const opts=await c.getByRole('tab').allTextContents();assert.equal(opts.length,9);let count=0;for(const x of opts.filter(x=>x!=='전체')){await c.getByRole('tab',{name:x,exact:true}).click();count+=await d.locator('.sa-layout').count();}assert.equal(count,60);await c.getByRole('tab',{name:'전체',exact:true}).click();await p.screenshot({path:'D:/Codex/Temp/wixel-smartart-gallery.png'});await d.getByRole('button',{name:'취소',exact:true}).click();assert.equal((await shapes(p)).length,0);});
  await test('삽입SVG·한번UndoRedo',async p=>{const s=await create(p);assert.equal(s.kind,'smartart');assert.equal(await p.locator('.obj [data-smartart-node]').count(),3);assert.equal(await p.evaluate(()=>window.tabula.wb().undoStack.length),1);await run(p,'undo');assert.equal((await shapes(p)).length,0);await run(p,'redo');assert.equal((await shapes(p)).length,1);});
  await test('텍스트·항목추가삭제·수준·순서·취소',async p=>{const before=await create(p),d=await edit(p);await d.getByLabel('항목 1 텍스트',{exact:true}).fill('첫 제목');await d.getByRole('button',{name:'항목 추가',exact:true}).click();assert.equal(await d.locator('.sa-text-row').count(),4);await d.getByRole('button',{name:'수준 내리기',exact:true}).click();assert.match(await d.locator('.sa-text-row.on').getAttribute('style'),/14px/);await d.getByRole('button',{name:'수준 올리기',exact:true}).click();await d.getByRole('button',{name:'아래로 이동',exact:true}).click();await d.getByRole('button',{name:'위로 이동',exact:true}).click();await d.getByRole('button',{name:'항목 삭제',exact:true}).click();assert.equal(await d.locator('.sa-text-row').count(),3);await d.getByRole('button',{name:'취소',exact:true}).click();assert.deepEqual((await shapes(p))[0],before);});
  await test('텍스트·색·스타일확인1Undo·배율크기',async p=>{const before=await create(p),d=await edit(p);await d.getByLabel('항목 1 텍스트',{exact:true}).fill('수정한 제목');await d.locator('.sa-appearance > summary').click();await d.getByRole('button',{name:'색: 초록색',exact:true}).click();await d.getByRole('button',{name:'스타일: 그림자',exact:true}).click();await d.getByRole('button',{name:'확인',exact:true}).click();let s=(await shapes(p))[0];assert.equal(s.smartArt.nodes[0].text,'수정한 제목');assert.equal(s.smartArt.style,'shadow');assert.equal(s.smartArt.palette[0],'#375623');await run(p,'undo');assert.deepEqual((await shapes(p))[0],before);await run(p,'redo');
    const handle=p.locator(`.obj[data-id="${s.id}"] .ch-h.se`).first(),box=await handle.boundingBox();assert.ok(box);await p.mouse.move(box.x+box.width/2,box.y+box.height/2);await p.mouse.down();await p.mouse.move(box.x+box.width/2+80,box.y+box.height/2+40,{steps:8});await p.mouse.up();s=(await shapes(p))[0];assert.ok(s.w>before.w+30);assert.deepEqual(s.smartArt.nodes.map(n=>n.text),['수정한 제목','항목 2','항목 3']);
    for(const zoom of [50,125,100]){await p.evaluate(z=>{const g=window.tabula.gv();g.setZoom(z);g.renderAll();},zoom);const b=await p.locator(`.obj[data-id="${s.id}"] svg[aria-label="SmartArt"]`).first().boundingBox();assert.ok(Math.abs(b.width-s.w*zoom/100)<2,'배율에 맞는 SVG 크기');}
    await p.screenshot({path:process.env.WIXEL_SMARTART_SCREENSHOT||'D:/Codex/Temp/wixel-smartart.png'});
  });
  await test('로컬그림선택·그림배치·저장모델',async p=>{const d=await open(p);const png=await p.evaluate(()=>{const c=document.createElement('canvas');c.width=120;c.height=80;const g=c.getContext('2d');g.fillStyle='#008899';g.fillRect(0,0,120,80);return c.toDataURL();});const [chooser]=await Promise.all([p.waitForEvent('filechooser'),d.getByRole('button',{name:'그림 선택…',exact:true}).click()]);await chooser.setFiles({name:'synthetic.png',mimeType:'image/png',buffer:Buffer.from(png.split(',')[1],'base64')});await p.waitForFunction(()=>document.querySelector('.sa-preview image'));await d.getByRole('button',{name:'확인',exact:true}).click();const s=(await shapes(p))[0];assert.equal(s.smartArt.layout,'pictureCards');assert.match(s.smartArt.nodes[0].picture,/^data:image\/png;base64,/);assert.equal(await p.locator('.obj image').count(),1);});
  await test('그림 읽기 중 확인 차단·닫기 뒤 지연 결과 폐기',async p=>{
    await p.evaluate(()=>{const Native=window.FileReader;window.smartArtPendingReads=[];window.FileReader=class{readAsDataURL(file){const reader=new Native();reader.onload=()=>{this.result=reader.result;window.smartArtPendingReads.push(()=>this.onload?.());};reader.onerror=()=>this.onerror?.();reader.readAsDataURL(file);}};});
    const png=await p.evaluate(()=>{const c=document.createElement('canvas');c.width=2;c.height=2;c.getContext('2d').fillRect(0,0,2,2);return c.toDataURL();});
    const upload=async d=>{await d.getByLabel('선택 항목 그림 파일').setInputFiles({name:'delayed.png',mimeType:'image/png',buffer:Buffer.from(png.split(',')[1],'base64')});await p.waitForFunction(()=>window.smartArtPendingReads.length===1);};
    let d=await open(p);await upload(d);await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal(await d.count(),1);assert.equal((await shapes(p)).length,0);assert.match(await d.locator('.sa-note[role=status]').innerText(),/읽은 뒤/);
    await p.evaluate(()=>window.smartArtPendingReads.shift()());await p.waitForFunction(()=>document.querySelector('.sa-preview image'));await d.getByRole('button',{name:'취소',exact:true}).click();assert.equal((await shapes(p)).length,0);
    d=await open(p);await upload(d);await d.getByRole('button',{name:'취소',exact:true}).click();await p.evaluate(()=>window.smartArtPendingReads.shift()());await p.evaluate(()=>new Promise(requestAnimationFrame));assert.equal(await dialog(p).count(),0);assert.equal((await shapes(p)).length,0);
  });
  await test('일반그룹변환과Undo',async p=>{await create(p);const d=await edit(p);await d.getByRole('button',{name:'일반 도형으로 변환',exact:true}).click();assert.equal((await shapes(p))[0].kind,'group');assert.ok((await shapes(p))[0].groupItems.length>=3);await run(p,'undo');assert.equal((await shapes(p))[0].kind,'smartart');});
  await test('보호시트삽입과열린편집창후보호변경차단',async p=>{await p.evaluate(()=>{const t=window.tabula;t.wb().setSheetProp(0,'protect',{on:true,allow:{}});});await run(p,'insertSmartArt');assert.equal(await dialog(p).count(),0);assert.equal((await shapes(p)).length,0);await p.evaluate(()=>window.tabula.wb().setSheetProp(0,'protect',null));await create(p);const d=await edit(p);await d.getByLabel('항목 1 텍스트',{exact:true}).fill('차단돼야 함');await p.evaluate(()=>window.tabula.wb().setSheetProp(0,'protect',{on:true,allow:{}}));await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal(await dialog(p).count(),1);assert.notEqual((await shapes(p))[0].smartArt.nodes[0].text,'차단돼야 함');await d.getByRole('button',{name:'취소',exact:true}).click();});
  await test('열린창문서전환은다른문서에삽입하지않음',async p=>{const d=await open(p);await p.evaluate(()=>{const t=window.tabula;t.wb().restore({sheets:[{name:'다른 합성',cells:{'0,0':{raw:'보존'}}}]});});await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal((await shapes(p)).length,0);assert.equal(await dialog(p).count(),1);await d.getByRole('button',{name:'취소',exact:true}).click();});
  await test('좌우·상하대칭메뉴가SVG에반영되고Undo된다',async p=>{const s=await create(p);for(const label of ['좌우 대칭','상하 대칭']){await p.evaluate(()=>window.tabula.openNamedMenu('objRotate',{x:300,y:150}));await p.getByRole('menuitem',{name:label,exact:true}).click();}const after=(await shapes(p))[0];assert.equal(after.flip,true);assert.equal(after.flipV,true);const transform=await p.locator(`.obj[data-id="${s.id}"] svg[aria-label="SmartArt"] > g`).first().getAttribute('transform');assert.match(transform,/scale\(-1,-1\)/);await run(p,'undo');assert.equal((await shapes(p))[0].flipV,undefined);assert.equal((await shapes(p))[0].flip,true);});
  await test('검색·설명·갤러리 방향키는 초안을 바꾸고 취소로 복원',async p=>{
    const before=await create(p),d=await edit(p),search=d.getByRole('searchbox',{name:'SmartArt 레이아웃 검색'});
    await search.fill('시간선');assert.equal(await d.locator('.sa-layout').count(),1);await search.press('ArrowDown');assert.equal(await p.evaluate(()=>document.activeElement.getAttribute('aria-label')),'기본 시간 표시 막대');await p.keyboard.press('Enter');
    assert.equal(await d.locator('.sa-selected-name').innerText(),'기본 시간 표시 막대');assert.match(await d.locator('.sa-description').innerText(),/시간선/);assert.equal(await d.count(),1,'갤러리 Enter는 확인 버튼을 누르지 않음');
    await search.fill('없는레이아웃999');assert.equal(await d.locator('.sa-layout').count(),0);assert.match(await d.locator('.sa-no-results').innerText(),/일치하는/);
    await search.fill('');assert.equal(await d.locator('.sa-layout').count(),60);await search.press('ArrowDown');await p.keyboard.press('End');await p.keyboard.press('Enter');assert.equal(await d.locator('.sa-selected-name').innerText(),'그림 위 제목 목록');
    await d.getByRole('button',{name:'취소',exact:true}).click();assert.deepEqual((await shapes(p))[0],before);
  });
  await test('Enter 추가·Shift Enter 줄바꿈·Tab 수준과 앞/하위 추가',async p=>{
    const d=await open(p),first=d.getByLabel('항목 1 텍스트',{exact:true});await first.fill('처음');await first.press('End');await first.press('Shift+Enter');await p.keyboard.insertText('둘째 줄');assert.equal(await first.inputValue(),'처음\n둘째 줄');
    await first.press('Enter');assert.equal(await d.locator('.sa-text-row').count(),4);assert.equal(await p.evaluate(()=>document.activeElement.getAttribute('aria-label')),'항목 2 텍스트');await p.keyboard.press('Tab');assert.match(await d.locator('.sa-text-row.on').getAttribute('style'),/14px/);await p.keyboard.press('Shift+Tab');
    await d.getByRole('button',{name:'하위 항목 추가',exact:true}).click();assert.equal(await d.locator('.sa-text-row').count(),5);await d.getByRole('button',{name:'앞에 추가',exact:true}).click();assert.equal(await d.locator('.sa-text-row').count(),6);
    await d.getByRole('button',{name:'확인',exact:true}).click();const result=(await shapes(p))[0];assert.equal(result.smartArt.nodes.length,6);assert.equal(result.smartArt.nodes[0].text,'처음\n둘째 줄');assert.equal(await p.evaluate(()=>window.tabula.wb().undoStack.length),1);await run(p,'undo');assert.equal((await shapes(p)).length,0);
  });
  await test('색 견본8종·스타일 미리보기4종과 그라데이션 적용',async p=>{
    const d=await open(p);await d.locator('.sa-appearance > summary').click();assert.equal(await d.locator('.sa-palette').count(),8);assert.equal(await d.locator('.sa-style').count(),4);
    await d.getByRole('button',{name:'색: 청록색',exact:true}).click();await d.getByRole('button',{name:'스타일: 그라데이션',exact:true}).click();assert.ok(await d.locator('.sa-preview linearGradient').count()>0);assert.equal(await d.getByRole('button',{name:'스타일: 그라데이션',exact:true}).getAttribute('aria-pressed'),'true');
    await p.screenshot({path:'D:/Codex/Temp/wixel-smartart-appearance.png'});await d.getByRole('button',{name:'확인',exact:true}).click();const s=(await shapes(p))[0];assert.equal(s.smartArt.style,'gradient');assert.equal(s.smartArt.palette[0],'#006b76');
  });
  await test('60항목 경계에서 추가 비활성·삭제 복구와 취소',async p=>{
    const initial=await create(p);await p.evaluate(()=>{const t=window.tabula,w=t.wb(),shape=w.sheets[0].shapes[0];w.transact(()=>w.setSheetProp(0,'shapes',[{...shape,smartArt:{...shape.smartArt,nodes:Array.from({length:60},(_,i)=>({id:'n'+i,text:'합성 '+i,level:0}))}}]));});
    const before=(await shapes(p))[0],d=await edit(p);assert.equal(await d.locator('.sa-text-row').count(),60);for(const name of ['항목 추가','앞에 추가','하위 항목 추가'])assert.equal(await d.getByRole('button',{name,exact:true}).isDisabled(),true);
    await d.getByRole('button',{name:'항목 삭제',exact:true}).click();assert.equal(await d.locator('.sa-text-row').count(),59);assert.equal(await d.getByRole('button',{name:'항목 추가',exact:true}).isEnabled(),true);await d.getByRole('button',{name:'취소',exact:true}).click();assert.deepEqual((await shapes(p))[0],before);
  });
  await test('320·390 화면에서 갤러리·텍스트·확인/취소 접근',async p=>{
    for(const width of [390,320]){await p.setViewportSize({width,height:844});const d=await open(p);await d.getByRole('tab',{name:'전체',exact:true}).click();await d.getByRole('button',{name:'그림 위 제목 목록',exact:true}).click();
      const last=d.getByLabel('항목 3 텍스트',{exact:true});await last.scrollIntoViewIfNeeded();const b=await last.boundingBox();assert.ok(b.x>=0&&b.x+b.width<=width+1);const overflow=await d.evaluate(e=>{const b=e.querySelector('.dialog-body');return b.scrollWidth-b.clientWidth});assert.ok(overflow<=2,'본문 가로 넘침 '+overflow);
      for(const name of ['확인','취소']){const button=d.getByRole('button',{name,exact:true});await button.scrollIntoViewIfNeeded();const r=await button.boundingBox();assert.ok(r.x>=0&&r.x+r.width<=width+1&&r.y>=0&&r.y+r.height<=845);}
      await p.screenshot({path:'D:/Codex/Temp/wixel-smartart-'+width+'.png'});await d.getByRole('button',{name:'취소',exact:true}).click();assert.equal((await shapes(p)).length,0);
    }
  });

}finally{await browser.close();}
console.log(JSON.stringify({total:results.length,passed:results.filter(x=>x.ok).length,results},null,2));if(results.some(x=>!x.ok))process.exitCode=1;
