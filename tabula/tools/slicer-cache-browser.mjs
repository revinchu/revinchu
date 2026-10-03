// 독립 슬라이서 합성 파일만 사용합니다. API와 외부 전송을 모두 차단합니다.
import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {Workbook} from '../src/workbook.js';
import {writeXlsx} from '../src/xlsx.js';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=new URL(process.env.WIXEL_URL||'http://127.0.0.1:5191/');
if(!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw Error('로컬 서버만 검사합니다.');
const out=process.env.WIXEL_SLICER_CACHE_OUT||'D:/Codex/Temp/wixel-final-audit/cache-browser';
await mkdir(out,{recursive:true});
const wb=new Workbook(),source={kind:'cache',field:'Month',cacheKey:'shared',cacheSource:{name:'Missing',external:'file:///unavailable/source.xlsx'},values:[1,2,3,4,5],items:[0,1,2,3,4].map(index=>({index,hasData:true})),format:{}};
wb.sheets[0].slicers=[0,1].map(i=>({id:'c'+i,caption:'Month'+i,source,cacheSelection:['2','3'],x:20+i*200,y:20,w:180,h:230,style:'SlicerStyleLight1'}));
const file=out+'/synthetic.xlsx';await writeFile(file,writeXlsx(wb));
const reports=[];
for(const engine of (process.env.WIXEL_BROWSER||'chromium,webkit').split(',')){
 const browser=await pw[engine].launch(),context=await browser.newContext({viewport:{width:1200,height:900}}),page=await context.newPage(),errors=[],blocked=[];
 await context.route('**/*',r=>{const u=new URL(r.request().url());if(u.origin!==url.origin||u.pathname.startsWith('/api/')||!['GET','HEAD'].includes(r.request().method())){blocked.push(u.protocol+' blocked');return r.abort();}return r.continue();});
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 page.on('pageerror',e=>errors.push(e.message));
 let checks=0;const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);};const settle=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
 try{
  await page.goto(url.href);await page.waitForFunction(()=>window.tabula?.wb());
  await page.locator('#fileInput').setInputFiles(file);await page.waitForFunction(()=>!document.querySelector('.load-progress')&&tabula.wb().sheets[0].slicers.length===2);await settle();
  const sl=()=>page.locator('.obj.slicer').first(),selected=()=>page.locator('.obj.slicer').evaluateAll(ns=>ns.map(n=>[...n.querySelectorAll('.sl-item.on')].map(x=>x.dataset.k)));
  eq(await page.locator('.obj.slicer .sl-item').count(),10,'독립 캐시 항목 두 슬라이서 표시');eq(await selected(),[['2','3'],['2','3']],'저장된 선택 유지');
  await sl().locator('.sl-item[data-k="0"]').click();await settle();eq(await selected(),[['0'],['0']],'단일 선택과 공유 슬라이서 동기화');
  await sl().locator('.sl-item[data-k="1"]').click({modifiers:['Control']});await settle();eq(await selected(),[['0','1'],['0','1']],'Ctrl 다중 선택');
  await sl().locator('.sl-clear').click();await settle();eq((await selected()).map(x=>x.length),[5,5],'선택 해제');
  for(let i=0;i<3;i++){await page.evaluate(()=>tabula.run('undo'));await settle();}eq(await selected(),[['2','3'],['2','3']],'선택 실행 취소 복원');
  await sl().locator('.sl-item').first().click({button:'right'});eq(await page.getByRole('menuitem',{name:/^새로 고침/}).isDisabled(),true,'연결 원본 없는 새로 고침 비활성');eq(await page.getByRole('menuitem',{name:/^셀 서식/}).count(),0,'슬라이서 전용 메뉴');
  await page.getByRole('menuitem',{name:/^텍스트 내림차순 정렬/}).click();await settle();eq(await sl().locator('.sl-item').evaluateAll(ns=>ns.map(n=>n.textContent.trim())),['5','4','3','2','1'],'내림차순 설정');
  await page.evaluate(()=>tabula.run('undo'));await settle();eq(await sl().locator('.sl-item').evaluateAll(ns=>ns.map(n=>n.textContent.trim())),['1','2','3','4','5'],'정렬 실행 취소');
  eq(errors,[],'페이지 오류 없음');eq(blocked,[],'외부 연결 시도 없음');await page.screenshot({path:out+'/'+engine+'.png'});reports.push({engine,checks,ok:true});
 }catch(e){reports.push({engine,checks,ok:false,error:e.message,errors,blocked});await page.screenshot({path:out+'/'+engine+'-failure.png'}).catch(()=>{});}
 finally{await browser.close();}
}
await writeFile(out+'/result.json',JSON.stringify(reports,null,2));console.log(JSON.stringify(reports));if(reports.some(r=>!r.ok))process.exitCode=1;
