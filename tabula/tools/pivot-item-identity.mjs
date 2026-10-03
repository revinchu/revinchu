// 합성 데이터만 사용. 원본 파일·클립보드·외부/API 접근 없음.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=new URL(process.env.WIXEL_URL||'http://127.0.0.1:8792/');
if(!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw Error('로컬 서버만 검사합니다.');
const out=process.env.WIXEL_ITEM_IDENTITY_OUT||'D:/Codex/Temp/wixel-final-audit/item-identity-ui';
await mkdir(out,{recursive:true});
const reports=[];
for(const engine of (process.env.WIXEL_BROWSER||'chromium,webkit').split(',')){
 const browser=await pw[engine].launch(), context=await browser.newContext({viewport:{width:1200,height:840}}), page=await context.newPage();
 const errors=[],blocked=[];let checks=0;
 const eq=(actual,expected,label)=>{checks++;assert.deepEqual(actual,expected,label);};
 const settle=()=>page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 await context.route('**/*',route=>{const req=route.request(),u=new URL(req.url());if(u.origin!==url.origin||u.pathname.startsWith('/api/')||!['GET','HEAD'].includes(req.method())){blocked.push({kind:u.pathname.startsWith('/api/')?'api':'external',method:req.method(),path:u.pathname});return route.abort();}return route.continue();});
 page.on('pageerror',e=>errors.push(e.message));page.setDefaultTimeout(12000);
 const filter=()=>page.evaluate(()=>tabula.wb().sheets[1].pivot.filters?.Label??null);
 const undo=async()=>{await page.evaluate(()=>tabula.run('undo'));await settle();};
 const total=()=>page.evaluate(()=>{const w=tabula.wb(),d=w.sheets[1].pivot,a=d.area;for(let r=a.r1;r<=a.r2;r++)if(w.getValue(1,r,a.c1)==='총합계')return w.getValue(1,r,a.c1+1);return null;});
 const slicer=()=>page.locator('.obj.slicer[data-id="identity-slicer"]').first();
 const selected=()=>slicer().locator('.sl-item.on').evaluateAll(nodes=>nodes.map(n=>n.dataset.k));
 const original=()=>page.evaluate(()=>[1,2,3,4].map(r=>tabula.wb().getRaw(0,r,0)));
 async function reset(filters){
  await page.keyboard.press('Escape');
  await page.evaluate(filters=>{const t=tabula,w=t.wb();if(t.si!==0)t.switchSheet(0);const rows=[['Label','Amount'],['Alpha',1],['alpha',2],['ALPHA',3],['Beta',4]],cells={};rows.forEach((row,r)=>row.forEach((v,c)=>cells[r+','+c]={raw:String(v)}));w.restore({sheets:[{name:'Source',cells},{name:'Report',cells:{},pivot:{name:'IdentityPivot',source:'Source',range:{r1:0,c1:0,r2:4,c2:1},rows:['Label'],cols:[],pages:[],values:[{field:'Amount',agg:'sum'}],top:1,left:1,layout:'tabular',...(filters?{filters}: {})},slicers:[{id:'identity-slicer',caption:'Label',x:450,y:30,w:210,h:190,style:'SlicerStyleLight1',source:{kind:'pivot',field:'Label',pivots:[{sheet:'Report',name:'IdentityPivot'}]}}]}]});t.switchSheet(1,false);t.selectCell(1,1);t.run('pivotRefresh');w.undoStack=[];w.redoStack=[];t.gv().renderAll();},filters??null);
  await settle();
 }
 async function openFilter(){await page.locator('.pbtn[data-k="rows"][data-f="Label"]').first().click();const menu=page.locator('.filter-menu.pivot-filter-menu');await menu.waitFor();return menu;}
 try{
  await page.goto(url.href);await page.waitForFunction(()=>window.tabula?.wb());
  const asset=await page.evaluate(()=>[...document.scripts].map(s=>s.getAttribute('src')).find(s=>s?.includes('wixel-'))||'source');
  await reset();
  eq(await slicer().locator('.sl-item').evaluateAll(ns=>ns.map(n=>n.dataset.k)),['Alpha','Beta'],'대소문자 항목 하나로 표시');
  eq(await total(),10,'행 전체 총합계');
  await slicer().locator('.sl-item[data-k="Alpha"]').click();await settle();
  eq(await filter(),['Alpha'],'대표 항목으로 선택 저장');eq(await total(),6,'세 원본 행이 한 선택에 포함');
  eq(await selected(),['Alpha'],'단일 선택 표시');
  await slicer().locator('.sl-item[data-k="Beta"]').click({modifiers:['Control']});await settle();
  eq(await total(),10,'Ctrl 다중 선택은 모든 값을 포함');eq(await selected(),['Alpha','Beta'],'Ctrl 선택 표시');
  await undo();eq(await total(),6,'다중 선택 Undo');await undo();eq(await total(),10,'첫 선택 Undo');
  eq(await original(),['Alpha','alpha','ALPHA','Beta'],'원본 대소문자는 변경하지 않음');
  await reset({Label:['ALPHA']});eq(await total(),6,'가져온 alias 필터 집계');eq(await selected(),['Alpha'],'가져온 alias 선택 표시');
  let menu=await openFilter();eq(await menu.getByRole('checkbox',{name:'Alpha',exact:true}).isChecked(),true,'필터 팝업도 alias 선택 표시');
  await menu.getByRole('button',{name:'취소',exact:true}).click();eq(await filter(),['ALPHA'],'팝업 취소는 원래 필터 값을 보존');
  menu=await openFilter();await menu.getByRole('searchbox').fill('beta');
  eq(await menu.getByRole('checkbox',{name:'Beta',exact:true}).isChecked(),true,'검색 결과 기본 선택');
  await menu.getByRole('checkbox',{name:'필터에 현재 선택한 내용 추가',exact:true}).check();
  await menu.getByRole('button',{name:'확인',exact:true}).click();await settle();eq(await total(),10,'검색 결과와 기존 alias 선택 합침');
  await undo();eq(await filter(),['ALPHA'],'추가 선택 Undo는 원래 alias 복원');eq(await total(),6,'추가 선택 Undo 집계');
  menu=await openFilter();await menu.getByRole('searchbox').fill('ALP');
  eq(await menu.locator('[data-filter-index]').count(),1,'검색 목록에도 중복 항목 없음');
  const all=menu.getByRole('checkbox',{name:'표시된 항목 모두 선택',exact:true});
  await all.uncheck();await all.check();await menu.getByRole('button',{name:'확인',exact:true}).click();await settle();
  eq(await total(),6,'검색 모두 선택 집계');eq(await original(),['Alpha','alpha','ALPHA','Beta'],'선택 작업 후에도 원본 문자열 보존');
  eq(errors,[],'실행 오류 없음');eq(blocked.filter(r=>r.kind!=='api'||r.method!=='GET'),[],'외부 및 쓰기 요청 시도 없음 (배경 API GET은 전송 전 차단)');await page.screenshot({path:out+'/'+engine+'.png'});
  reports.push({engine,asset,checks,ok:true,errors,blocked});
 }catch(error){reports.push({engine,checks,ok:false,error:error.message,errors,blocked});await page.screenshot({path:out+'/'+engine+'-failure.png'}).catch(()=>{});}
 finally{await context.close();await browser.close();}
}
await writeFile(out+'/result.json',JSON.stringify({url:url.href,reports},null,2));console.log(JSON.stringify(reports));if(reports.some(r=>!r.ok))process.exitCode=1;
