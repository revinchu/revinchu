import assert from 'node:assert/strict';
import {mkdir,writeFile} from 'node:fs/promises';
import {Workbook} from '../src/workbook.js';
import {readXlsx,writeXlsx} from '../src/xlsx.js';

const engines=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const engine=process.env.WIXEL_BROWSER||'chromium',url=process.env.WIXEL_URL||'http://localhost:5178/',origin=new URL(url).origin;
const publicCheck=process.env.WIXEL_IMPORT_PUBLIC==='1'&&origin==='https://wixel-3.wizx.workers.dev';
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname)&&!publicCheck)throw new Error('로컬 또는 명시적으로 허용한 공개 주소에서만 실행하세요.');
const out=process.env.WIXEL_OUTPUT||'D:/Codex/Temp/wixel-import-formula-diagnostics-'+engine;
await mkdir(out,{recursive:true});
const browser=await engines[engine].launch(),results=[],assets=new Set();let checks=0;
const eq=(actual,expected,message)=>{assert.deepEqual(actual,expected,message);checks++;};
const ok=(value,message)=>{assert.ok(value,message);checks++;};
const dialog=(p,name)=>p.getByRole('dialog',{name,exact:true});
const frame=p=>p.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve))));
const run=(p,command)=>p.evaluate(command=>tabula.run(command),command);
const cell=(p,si,r,c)=>p.evaluate(({si,r,c})=>{const w=tabula.wb(),value=w.getValue(si,r,c);return {raw:w.getRaw(si,r,c),value:value?.code?{error:value.code}:value,status:w.getCalculationStatus(si,r,c)};},{si,r,c});
const select=(p,si,r,c)=>p.evaluate(({si,r,c})=>{tabula.switchSheet(si);tabula.selectCell(r,c);tabula.gv().renderAll();},{si,r,c});
const longSheet='합성시트'+'가'.repeat(27),longFormula='=UNKNOWN_'+'LONG'.repeat(24)+'(A1)';
const original=()=>new Workbook({sheets:[{name:'안내',cells:{'0,0':{raw:'수식 진단 합성 자료'}}},{name:'원본 오류',fileValues:true,cells:{'0,0':{raw:'2'},'0,1':{raw:'=UNKNOWN_REPORT_FN(A1)',cached:{error:'#NAME?'}},'0,2':{raw:'=UNKNOWN_LABEL',cached:{error:'#NAME?'}}}},{name:longSheet,fileValues:true,cells:{'0,0':{raw:'1'},'0,1':{raw:longFormula,cached:{error:'#NAME?'}}}}]});
const numeric=()=>new Workbook({sheets:[{name:'미지원 저장값',fileValues:true,cells:{'0,0':{raw:'2'},'0,1':{raw:'=UNKNOWN_REPORT_FN(A1)',cached:27},'0,2':{raw:'=SUM(B1,5)',cached:32}}}]});
const locals=()=>new Workbook({names:[{name:'twice',ref:'=LAMBDA(x,x*2)'}],sheets:[{name:'지역 함수',cells:{'0,0':{raw:'3'},'0,1':{raw:'=LET(f,LAMBDA(x,x+1),f(A1))'},'1,1':{raw:'=LAMBDA(fn,fn(A1))(LAMBDA(x,x+2))'},'2,1':{raw:'=LET(sumfn,SUM,sumfn(A1,2))'},'3,1':{raw:'=LET(f,LAMBDA(x,twice(x)),f(A1))'}}}]});
const mixed=()=>new Workbook({sheets:[{name:'혼합 진단',fileValues:true,cells:{'0,0':{raw:'2'},'0,1':{raw:'=MISSING(A1)',cached:{error:'#NAME?'}},'1,1':{raw:'=MISSING(A1)',cached:42},'2,1':{raw:'=MISSING(A1)',cached:'#NAME?'}}}]});
async function reportLayout(d,mobile){
 const metrics=await d.locator('.calculation-report').evaluate(report=>{
  const bounds=report.getBoundingClientRect(),body=report.closest('.dialog-body');
  const rows=[...report.querySelectorAll('tbody tr')].map(row=>{
   const [link,status,detail]=[...row.children],a=link.getBoundingClientRect(),b=status.getBoundingClientRect(),c=detail.getBoundingClientRect(),range=document.createRange();
   range.selectNodeContents(status);const lines=[...new Set([...range.getClientRects()].map(r=>Math.round(r.top)))];
   return {status:status.textContent,statusLines:lines.length,overflow:[link,status,detail].map(n=>n.scrollWidth-n.clientWidth),headerOverlap:a.right>b.left+1,detailBelow:c.top>=Math.max(a.bottom,b.bottom)-1,inside:[link,status,detail].every(n=>{const r=n.getBoundingClientRect();return r.left>=bounds.left-1&&r.right<=bounds.right+1;})};
  });
  return {reportOverflow:report.scrollWidth-report.clientWidth,bodyOverflow:body.scrollWidth-body.clientWidth,tableHeaderVisible:getComputedStyle(report.querySelector('thead')).display!=='none',rows};
 });
 ok(metrics.reportOverflow<=1&&metrics.bodyOverflow<=1,'계산 보고서에 가로 스크롤 없음');
 ok(metrics.rows.every(r=>r.statusLines===1),'상태 문구가 음절 단위로 줄바꿈되지 않음');
 ok(metrics.rows.every(r=>r.inside&&r.overflow.every(n=>n<=1)),'긴 시트명·수식도 카드 또는 표 셀을 벗어나지 않음');
 if(mobile){ok(!metrics.tableHeaderVisible,'좁은 화면 표 머리글 숨김');ok(metrics.rows.every(r=>!r.headerOverlap&&r.detailBelow),'셀 링크·상태가 겹치지 않고 수식 설명은 다음 행에 배치');}
 return metrics;
}
async function openBytes(p,buffer,name){
 await p.evaluate(()=>{window.__diagnosticPrevious=tabula.wb();});
 await p.locator('#fileInput').setInputFiles({name,mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer});
 await p.waitForFunction(()=>tabula.wb()!==window.__diagnosticPrevious&&!document.querySelector('.load-progress'),null,{timeout:60000});
 await frame(p);
}
async function input(p,si,r,c,value){
 await select(p,si,r,c);await p.keyboard.type(String(value));await p.keyboard.press('Enter');await frame(p);
 eq((await cell(p,si,r,c)).raw,String(value),'실제 키보드 입력 반영');
}
async function closeWarning(p){const d=dialog(p,'가져오기');await d.waitFor();const text=await d.innerText();await d.getByRole('button',{name:'확인',exact:true}).click();return text;}
async function test(name,make,fn){
 if(process.env.WIXEL_IMPORT_FILTER&&!process.env.WIXEL_IMPORT_FILTER.split('|').some(part=>name.includes(part)))return;
 const ctx=await browser.newContext({viewport:{width:1400,height:960},serviceWorkers:'block',acceptDownloads:true}),p=await ctx.newPage(),errors=[],writes=[],blocked=[],start=checks,started=Date.now();
 p.setDefaultTimeout(30000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.type()==='beforeunload'?d.accept():d.dismiss());
 await ctx.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1','off');});
 await ctx.route('**/*',route=>{const req=route.request(),u=new URL(req.url());if(!['GET','HEAD','OPTIONS'].includes(req.method())){writes.push(req.method()+' '+u.pathname);return route.abort();}if(u.origin!==origin||/^\/api(?:\/|$)/.test(u.pathname)){blocked.push(u.pathname);return route.abort();}return route.continue();});
 try{
  await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.wb(),null,{timeout:60000});
  for(const asset of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(asset);
  const buffer=Buffer.from(writeXlsx(make()));await writeFile(out+'/'+name+'-input.xlsx',buffer);
  await openBytes(p,buffer,name+'.xlsx');await p.screenshot({path:out+'/'+name+'-imported.png'});
  const observation=await fn(p);eq(errors,[],'페이지 오류 없음');eq(writes,[],'원격 API·외부 쓰기 없음');
  await p.screenshot({path:out+'/'+name+'-after.png'});results.push({name,ok:true,checks:checks-start,elapsedMs:Date.now()-started,blockedRequests:blocked,observation});console.log('OK '+name+' '+(checks-start));
 }catch(e){
  const state=await p.evaluate(()=>({dialogs:[...document.querySelectorAll('#dialogLayer .dialog')].map(n=>n.innerText),calcState:document.getElementById('calcState')?.textContent,si:window.tabula?.si,active:window.tabula?.active,issues:window.tabula?.wb()?.calculationIssues({limit:20})})).catch(()=>null);
  await p.screenshot({path:out+'/'+name+'-failure.png'}).catch(()=>{});results.push({name,ok:false,checks:checks-start,elapsedMs:Date.now()-started,error:e.stack,errors,writes,blockedRequests:blocked,state});console.error('NG '+name+' '+e.message);
 }finally{await ctx.close();}
}
try{
 await test('original-name-error-no-import-warning',original,async p=>{
  eq(await dialog(p,'가져오기').count(),0,'원본 오류만 있는 파일은 미지원 가져오기 경고 없음');
  const initial=await cell(p,1,0,1);eq(initial.raw,'=UNKNOWN_REPORT_FN(A1)','원본 수식 정확히 보존');eq(initial.value,{error:'#NAME?'},'원본 오류값 보존');eq(initial.status.status,'source-error','원본 오류 상태');eq(initial.status.reason,'source-name-error','미지원과 구분된 원인');
  await p.locator('#calcState').click();const d=dialog(p,'계산 상태 확인');await d.waitFor();const rows=d.locator('tbody tr');eq(await rows.count(),3,'원본 함수/이름 오류와 긴 수식 세 셀 표시');eq(await rows.locator('td:nth-child(2)').allTextContents(),Array(3).fill('원본 오류'),'세 행 모두 원본 오류');
  const text=await d.innerText();ok(text.includes('원본 Excel 파일에도 #NAME? 오류'),'원본 오류 안내');ok(text.includes('파일에 있던 참고값: #NAME?'),'오류 객체가 읽을 수 있는 문자로 표시');ok(!text.includes('[object Object]'),'잘못된 객체 문자열 없음');
  const desktopLayout=await reportLayout(d,false);ok(text.includes(longSheet)&&text.includes(longFormula),'긴 합성 시트명과 수식 원문 유지');await p.screenshot({path:out+'/original-error-report.png'});await p.setViewportSize({width:390,height:844});await frame(p);const box=await d.boundingBox();ok(box&&box.x>=-1&&box.y>=-1&&box.x+box.width<=391&&box.y+box.height<=845,'390px 계산 상태 창이 화면 안에 배치');const mobileText=await d.innerText();ok(mobileText.includes('원본 Excel 파일에도 #NAME? 오류')&&mobileText.includes('=UNKNOWN_REPORT_FN(A1)'),'좁은 창에서도 수식·설명 원문 유지');ok(!mobileText.includes('[object Object]'),'좁은 창에 객체 문자열 없음');const mobileLayout=await reportLayout(d,true);ok(mobileText.includes(longSheet)&&mobileText.includes(longFormula),'좁은 화면의 긴 시트명과 수식 원문 유지');const close=d.getByRole('button',{name:'닫기',exact:true}).last();ok(await close.isVisible()&&await close.isEnabled(),'좁은 창 닫기 단추 사용 가능');await p.screenshot({path:out+'/original-error-report-mobile.png'});await rows.last().scrollIntoViewIfNeeded();await p.screenshot({path:out+'/original-error-report-mobile-long.png'});await close.click();eq(await d.count(),0,'좁은 창 닫기 실제 동작');await p.locator('#calcState').click();await d.getByRole('button',{name:'원본 오류!B1',exact:true}).click();eq(await p.evaluate(()=>({si:tabula.si,active:tabula.active})),{si:1,active:{r:0,c:1}},'진단 링크가 다른 시트의 정확한 셀로 이동');eq(await d.count(),0,'이동 후 창 닫기');ok((await p.locator('#calcState').innerText()).includes('원본 오류'),'상태 표시줄 원본 오류');
  await input(p,1,0,0,'3');eq((await cell(p,1,0,1)).status.status,'stale','원본 오류도 입력 변경 후 stale');eq((await cell(p,1,0,1)).raw,initial.raw,'입력 변경으로 수식 원문 미변경');await run(p,'undo');eq((await cell(p,1,0,0)).raw,'2','Undo 원본 입력 복원');eq((await cell(p,1,0,1)).status.status,'source-error','Undo 원본 오류 상태 복원');
  return {initial,desktopLayout,mobileLayout};
 });
 await test('unsupported-numeric-cache-invalidates',numeric,async p=>{
  const warning=await closeWarning(p);ok(/지원하지 않는 함수가 쓰인 수식 1개/.test(warning),'실제 미지원 숫자 저장값 경고 유지');eq((await cell(p,0,0,1)).value,27,'가져올 때 숫자 저장값 표시');eq((await cell(p,0,0,1)).status.status,'cached','저장값 상태');eq((await cell(p,0,0,2)).value,32,'의존 합계 초기 저장값');
  await select(p,0,0,1);ok((await p.locator('#calcState').innerText()).includes('파일 저장값'),'저장값 상태 표시줄');await input(p,0,0,0,'4');
  for(const c of [1,2])eq((await cell(p,0,0,c)).value,{error:'#NAME?'},'오래된 숫자가 계산에 남지 않음 '+c);eq((await cell(p,0,0,1)).status.status,'stale','수정 후 stale 상태');await select(p,0,0,1);ok((await p.locator('#calcState').innerText()).includes('재계산 불가'),'수정 후 상태 표시줄');await run(p,'undo');eq((await cell(p,0,0,1)).value,27,'Undo 저장값 복원');eq((await cell(p,0,0,2)).value,32,'Undo 의존 결과 복원');eq((await cell(p,0,0,1)).status.status,'cached','Undo cached 상태 복원');await run(p,'redo');eq((await cell(p,0,0,1)).value,{error:'#NAME?'},'Redo stale 차단');await run(p,'undo');return{warning};
 });
 await test('mixed-per-cell-saved-error-types',mixed,async p=>{
  const warning=await closeWarning(p);ok(/지원하지 않는 함수가 쓰인 수식 2개/.test(warning),'원본 오류 객체를 빼고 숫자/문자열 두 셀만 경고');
  const cells=await Promise.all([0,1,2].map(r=>cell(p,0,r,1)));eq(cells.map(c=>c.status.status),['source-error','cached','cached'],'같은 수식의 서로 다른 저장 타입을 독립 판정');eq(cells.map(c=>c.value),[{error:'#NAME?'},42,'#NAME?'],'오류 객체/숫자/문자열 타입 보존');eq(cells.map(c=>c.raw),Array(3).fill('=MISSING(A1)'),'동일 수식 원문 보존');await p.locator('#calcState').click();const d=dialog(p,'계산 상태 확인');eq(await d.locator('tbody tr td:nth-child(2)').allTextContents(),['원본 오류','파일 저장값','파일 저장값'],'계산 상태 목록도 독립 상태 표시');return{warning,cells};
 });
 await test('local-functions-recalculate-undo-save-reopen',locals,async p=>{
  eq(await dialog(p,'가져오기').count(),0,'정상 지역 함수는 가져오기 경고 없음');
  const rows=async()=>Promise.all([0,1,2,3].map(r=>cell(p,0,r,1)));const initial=await rows();eq(initial.map(c=>c.value),[4,5,5,6],'LET/LAMBDA·함수값·정의함수 계산');eq(await p.evaluate(()=>tabula.wb().calculationIssues().total),0,'정상 지역 함수 진단 없음');
  await run(p,'recalc');eq((await rows()).map(c=>c.value),[4,5,5,6],'전체 재계산 값 유지');await input(p,0,0,0,'5');eq((await rows()).map(c=>c.value),[6,7,7,10],'입력 변경의 모든 지역 함수 재계산');await run(p,'undo');eq((await rows()).map(c=>c.value),[4,5,5,6],'Undo 정상 결과 복원');await run(p,'redo');eq((await rows()).map(c=>c.value),[6,7,7,10],'Redo 재계산');
  await p.evaluate(()=>{window.__diagnosticSaved=[];window.showSaveFilePicker=async()=>({name:'지역 함수 저장.xlsx',createWritable:async()=>({write:async value=>{const b=value instanceof Blob?value:new Blob([value]);window.__diagnosticSaved.push(...new Uint8Array(await b.arrayBuffer()));},close:async()=>{},abort:async()=>{}})});});
  await p.evaluate(()=>tabula.exportXlsx('지역 함수 저장','xlsx'));await p.waitForFunction(()=>window.__diagnosticSaved.length>0);const saved=Buffer.from(await p.evaluate(()=>window.__diagnosticSaved));ok(saved.length>500,'브라우저 내보내기가 실제 XLSX 바이트 생성');await writeFile(out+'/local-functions-saved.xlsx',saved);const parsed=readXlsx(saved),back=new Workbook(parsed.data);eq(parsed.warnings.filter(w=>/지원하지 않는 함수/.test(w)),[],'저장본 미지원 경고 없음');eq([0,1,2,3].map(r=>back.getValue(0,r,1)),[6,7,7,10],'내보낸 파일 독립 파싱값');eq([0,1,2,3].map(r=>back.getRaw(0,r,1)),initial.map(c=>c.raw),'내보낸 파일 수식 보존');
  await openBytes(p,saved,'local-functions-reopened.xlsx');eq(await dialog(p,'가져오기').count(),0,'제품 재열기 경고 없음');eq((await rows()).map(c=>c.value),[6,7,7,10],'fileInput 재열기 결과 유지');eq((await rows()).map(c=>c.raw),initial.map(c=>c.raw),'fileInput 재열기 수식 보존');eq(await p.evaluate(()=>tabula.wb().calculationIssues().total),0,'재열기 지역 함수 진단 없음');return{savedBytes:saved.length};
 });
}finally{
 await browser.close();const result={engine,url,publicCheck,assets:[...assets],cases:results.length,passed:results.filter(r=>r.ok).length,checks,scope:'합성 XLSX만 사용. 실제 브라우저 fileInput 가져오기·계산 상태 UI·키보드 편집·Undo/Redo·메모리 파일 핸들 내보내기·재열기. 실제 사용자 문서나 외부 쓰기 없음.',results};await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify({engine,cases:result.cases,passed:result.passed,checks,assets:result.assets,out}));if(result.passed!==result.cases||assets.size!==1)process.exitCode=1;
}
