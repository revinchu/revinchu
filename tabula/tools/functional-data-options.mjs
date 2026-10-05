// Synthetic workbooks only. Local static reads are allowed; remote writes and OS save dialogs are blocked.
// WIXEL_DATA_OPTIONS_BASELINE injects only the old app.js; other modules remain current.
import assert from 'node:assert/strict';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
import {Workbook} from '../src/workbook.js';
import {readXlsx} from '../src/xlsx.js';
import {unzip,zip,textOf} from '../src/zip.js';
import {chartModelData} from '../src/chart.js';
import {parseXml,descendants} from '../src/xml.js';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const engine=process.env.WIXEL_BROWSER||'chromium',url=process.env.WIXEL_URL||'http://127.0.0.1:5195/',origin=new URL(url).origin;
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname))throw new Error('로컬 합성 검증 주소만 허용됩니다.');
const out=process.env.WIXEL_DATA_OPTIONS_OUT||'D:/Codex/Temp/wixel-upgrade-20261004/data-options-'+engine;
await mkdir(out,{recursive:true});
const baseline=process.env.WIXEL_DATA_OPTIONS_BASELINE?await readFile(process.env.WIXEL_DATA_OPTIONS_BASELINE+'/app.js','utf8'):null;
const browser=await pw[engine].launch(),results=[];let checks=0;
const eq=(a,b,m)=>{assert.deepEqual(a,b,m);checks++;},ok=(v,m)=>{assert.ok(v,m);checks++;};
const run=(p,c)=>p.evaluate(c=>tabula.run(c),c),settle=p=>p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
async function test(name,fn){
 if(process.env.WIXEL_DATA_OPTIONS_FILTER&&!name.includes(process.env.WIXEL_DATA_OPTIONS_FILTER))return;
 const context=await browser.newContext({viewport:{width:1366,height:950},serviceWorkers:'block'}),p=await context.newPage(),errors=[],writes=[],start=checks,observed={};
 p.setDefaultTimeout(20000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.dismiss());
 await context.addInitScript(()=>{
  window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1','off');localStorage.setItem('wixel.options',JSON.stringify({saveAsk:false,saveConfirm:false}));
  window.__memorySave={completed:[],aborted:0};
  Object.defineProperty(window,'showSaveFilePicker',{configurable:true,value:async()=>({kind:'file',name:'합성 원본 검수.xlsx',async createWritable(){const parts=[];let position=0,extent=0;return{
   async write(v){const bytes=v instanceof Blob?new Uint8Array(await v.arrayBuffer()):new Uint8Array(v.buffer??v,v.byteOffset??0,v.byteLength).slice();parts.push({position,bytes});position+=bytes.length;extent=Math.max(extent,position);},
   async seek(at){position=at;},async close(){const bytes=new Uint8Array(extent);for(const part of parts)bytes.set(part.bytes,part.position);window.__memorySave.completed.push(Array.from(bytes));},async abort(){window.__memorySave.aborted++;}
  };}})});
 });
 await context.route('**/*',route=>{const req=route.request(),u=new URL(req.url());if(!['GET','HEAD','OPTIONS'].includes(req.method())){writes.push(req.method()+' '+u.pathname);return route.abort();}if(u.origin!==origin||/^\/api(?:\/|$)/.test(u.pathname))return route.abort();if(baseline&&u.pathname==='/src/app.js')return route.fulfill({status:200,contentType:'application/javascript; charset=utf-8',body:baseline});return route.continue();});
 try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.gv(),null,{timeout:60000});await fn(p,observed);eq(errors,[],'페이지 오류 없음');eq(writes,[],'원격 쓰기 없음');await p.screenshot({path:out+'/'+name+'.png'});results.push({name,ok:true,checks:checks-start,observed});console.log('OK '+name);}
 catch(e){await p.screenshot({path:out+'/'+name+'-failure.png'}).catch(()=>{});results.push({name,ok:false,checks:checks-start,error:e.stack,observed,errors,writes});console.error('NG '+name+': '+e.message);}
 finally{await context.close();}
}
async function dataFixture(p,{table=true,slicers=false,cross=false,own=false,condition='custom'}={}){
 await p.evaluate(({table,slicers,cross,own,condition})=>{const cells={},rows=[['항목','금액'],['C','30'],['A','10'],['B','20']];rows.forEach((row,r)=>row.forEach((raw,c)=>cells[r+','+c]={raw}));
 const criterion=condition==='avg'?{type:'avg',above:true}:condition==='fill'?{type:'fill',value:'#FF0000'}:{type:'custom',op1:'gt',v1:'15'};
 if(condition==='fill')for(let r=1;r<=3;r++)cells[r+',1'].style={fill:r===2?'#0000FF':'#FF0000'};
 const filter={criteria:{1:criterion},hidden:condition==='avg'?{2:true,3:true}:{2:true}},rg={r1:0,c1:0,r2:3,c2:1};
 const sheets=[{name:'합성 원본',cells,...(table?{tables:[{id:'numbers',name:'NumbersTable',...rg,header:true,filter}]}:{filter:{...rg,...filter}})},{name:'대시보드',cells:{'0,0':{raw:'보존'},'1,1':{raw:'-999',style:{fill:'#0000FF'}},'2,1':{raw:'999',style:{fill:'#FF0000'}},'3,1':{raw:'-999',style:{fill:'#0000FF'}}}}];
 if(slicers)sheets[cross?1:0].slicers=[{id:'condition',source:{kind:'table',table:'NumbersTable',column:own?'금액':'항목'},caption:own?'금액':'항목',hideNoData:!own,showDeleted:false,x:250,y:30,w:230,h:250,style:'SlicerStyleLight1'}];
 const w=tabula.wb();w.restore({sheets});tabula.switchSheet(cross?0:1);tabula.switchSheet(cross?1:0);tabula.gv().setZoom(100);tabula.gv().layout();tabula.gv().renderAll();tabula.selectCell(1,1);w.undoStack=[];w.redoStack=[];window.__sourceCells=JSON.stringify(w.serialize().sheets.map(s=>s.cells));
 },{table,slicers,cross,own,condition});await settle(p);
}
const sourceState=p=>p.evaluate(()=>{const w=tabula.wb(),sh=w.sheets[0],f=sh.tables?.[0]?.filter??sh.filter;return{rows:[1,2,3].map(r=>[w.getValue(0,r,0),w.getValue(0,r,1)]),visible:[1,2,3].filter(r=>!f.hidden?.[r]).map(r=>[w.getValue(0,r,0),w.getValue(0,r,1)]),hidden:f.hidden,criteria:f.criteria,undo:w.undoStack.length,other:w.getValue(1,0,0)};});
const slicer=p=>p.locator('.pane-br .obj.slicer[data-id="condition"]');
const items=p=>slicer(p).locator('.sl-item').evaluateAll(ns=>ns.map(n=>({text:n.textContent,on:n.classList.contains('on'),hasData:!n.classList.contains('nodata')})));
try{
 for(const table of [false,true])for(const custom of [false,true])await test((custom?'custom':'quick')+'-sort-'+(table?'table':'filter'),async(p,o)=>{
  await dataFixture(p,{table});o.before=await sourceState(p);eq(o.before.visible,[['C',30],['B',20]],'초기 숫자 필터');
  if(custom){await run(p,'sortDialog');const d=p.getByRole('dialog',{name:'정렬',exact:true});await d.locator('.sort-row:not(.sort-head) select').nth(0).selectOption('1');await d.getByRole('button',{name:'확인',exact:true}).click();}else await run(p,'sortAsc');
  await settle(p);o.after=await sourceState(p);eq(o.after.rows,[['A',10],['B',20],['C',30]],'헤더 아래 전체 행 정렬');eq(o.after.visible,[['B',20],['C',30]],'정렬 후 조건에 맞는 행 표시');eq(o.after.criteria,o.before.criteria,'조건 보존');eq(o.after.undo,1,'정렬과 필터는 한 Undo');eq(o.after.other,'보존','다른 시트 보존');
  eq(await p.evaluate(()=>[1,2,3].map(r=>tabula.gv().rows.size(r)===0)),[true,false,false],'화면 숨김 높이도 새 행에 맞춤');
  await run(p,'undo');eq(await sourceState(p),o.before,'Undo 값과 필터 복원');await run(p,'redo');eq(await sourceState(p),o.after,'Redo 값과 필터 복원');await run(p,'undo');
 });
 for(const custom of [false,true])await test((custom?'custom':'quick')+'-sort-appended-filter',async(p,o)=>{
  await dataFixture(p,{table:false});
  await p.evaluate(()=>{const w=tabula.wb(),sh=w.sheets[0];w.setSheetProp(0,'filter',{...sh.filter,r2:2});tabula.selectCell(1,1);w.undoStack=[];w.redoStack=[];});
  o.before=await sourceState(p);eq(await p.evaluate(()=>tabula.sel.r1===tabula.sel.r2&&tabula.sel.c1===tabula.sel.c2),true,'선택은 셀 하나뿐');
  eq(await p.evaluate(()=>tabula.wb().sheets[0].filter.r2),2,'추가 행은 아직 기존 필터 범위 밖');
  if(custom){await run(p,'sortDialog');const d=p.getByRole('dialog',{name:'정렬',exact:true});await d.locator('.sort-row:not(.sort-head) select').nth(0).selectOption('1');await d.getByRole('button',{name:'확인',exact:true}).click();}else await run(p,'sortAsc');
  await settle(p);o.after=await sourceState(p);eq(o.after.rows,[['A',10],['B',20],['C',30]],'추가 행 B20까지 전체 정렬');eq(o.after.visible,[['B',20],['C',30]],'추가 행도 숫자 조건에 맞게 표시');
  const f=await p.evaluate(()=>tabula.wb().sheets[0].filter);eq(f.r2,3,'필터 범위를 추가 행까지 확장');eq(f.sort,{col:1,asc:true},'확장된 전체 정렬의 표준 정렬 설정');eq(o.after.undo,1,'필터 확장과 정렬은 한 Undo');
  eq(await p.evaluate(()=>[1,2,3].map(r=>tabula.gv().rows.size(r)===0)),[true,false,false],'그리드 숨김 행도 다시 계산');
  await run(p,'undo');eq(await sourceState(p),o.before,'Undo 원본 순서와 필터 복원');eq(await p.evaluate(()=>tabula.wb().sheets[0].filter.r2),2,'Undo 기존 필터 범위 복원');await run(p,'redo');eq(await sourceState(p),o.after,'Redo 추가 행을 포함한 결과 복원');
 });
 for(const cross of [false,true])for(const own of [false,true])await test('slicer-condition-'+(own?'own':'other')+'-'+(cross?'cross-sheet':'same-sheet'),async(p,o)=>{
  await dataFixture(p,{slicers:true,cross,own});await slicer(p).waitFor();o.before=await sourceState(p);o.items=await items(p);
  if(own){eq(o.items.map(i=>[i.text,i.on]),[['10',false],['20',true],['30',true]],'자기 열 숫자 조건의 실제 선택 상태');eq(await slicer(p).locator('.sl-clear.off').count(),0,'조건 필터 지우기 활성');}
  else{eq(o.items.map(i=>i.text),['B','C'],'다른 열 숫자 조건에 해당하는 두 항목');await slicer(p).getByRole('button',{name:'B',exact:true}).click();await settle(p);o.after=await sourceState(p);eq(o.after.visible,[['B',20]],'클릭한 항목과 숫자 조건 교집합');eq(o.after.criteria['1'],o.before.criteria['1'],'기존 숫자 조건 보존');eq(await p.evaluate(()=>tabula.si),cross?1:0,'다른 시트 연결도 현재 시트 유지');await run(p,'undo');eq(await sourceState(p),o.before,'슬라이서 Undo 복원');eq((await items(p)).map(i=>i.text),['B','C'],'Undo 항목 목록 복원');}
  eq(await p.evaluate(()=>JSON.stringify(tabula.wb().serialize().sheets.map(s=>s.cells))===window.__sourceCells),true,'슬라이서는 원본 셀을 바꾸지 않음');
 });
 for(const condition of ['avg','fill'])for(const own of [false,true])await test('slicer-'+condition+'-'+(own?'own':'other')+'-cross-sheet',async(p,o)=>{
  await dataFixture(p,{slicers:true,cross:true,own,condition});await slicer(p).waitFor();o.items=await items(p);
  if(own)eq(o.items.map(i=>[i.text,i.on]),[['10',false],['20',condition!=='avg'],['30',true]],'조건은 표시 시트가 아닌 원본 시트 값/색으로 평가');
  else eq(o.items.map(i=>i.text),condition==='avg'?['C']:['B','C'],'타 시트 평균/색 조건의 교차 항목');
  eq(await p.evaluate(()=>JSON.stringify(tabula.wb().serialize().sheets.map(s=>s.cells))===window.__sourceCells),true,'셀 데이터 보존');
 });
 for(const command of ['chartColumn','insertChartAll','chartSheet'])await test('full-chart-'+command,async(p,o)=>{
  await p.evaluate(()=>{const cells={'0,0':{raw:'항목'},'0,1':{raw:'금액'}};for(let r=1;r<=3000;r++){cells[r+',0']={raw:'행 '+r};cells[r+',1']={raw:String(r===3000?99999:1)};}const w=tabula.wb();w.restore({sheets:[{name:'합성 원본',cells},{name:'보존',cells:{'0,0':{raw:'수정 금지'}}}]});tabula.switchSheet(1);tabula.switchSheet(0);tabula.selectRange({r1:0,c1:0,r2:3000,c2:1},'cells',{r:0,c:0});w.undoStack=[];w.redoStack=[];});
  await run(p,command);if(command==='insertChartAll')await p.getByRole('dialog',{name:'차트 삽입',exact:true}).getByRole('button',{name:'확인',exact:true}).click();await settle(p);
  const data=await p.evaluate(()=>tabula.wb().serialize()),wb=new Workbook(data),host=wb.sheets.findIndex(s=>s.charts?.length),ch=wb.sheets[host]?.charts[0];ok(ch,'실제 명령으로 차트 생성');o.range=ch.range;const m=chartModelData(wb,host,ch);o.model={categories:m.categories.length,last:m.series[0]?.values.at(-1)};
  await p.evaluate(()=>tabula.exportXlsx('합성 전체 원본','xlsx'));const saved=await p.evaluate(()=>window.__memorySave);eq(saved.aborted,0,'저장 중단 없음');eq(saved.completed.length,1,'실제 앱 XLSX 저장 완료');const bytes=Uint8Array.from(saved.completed[0]);o.bytes=bytes.length;await writeFile(out+'/'+command+'.xlsx',bytes);
  const files=unzip(bytes),paths=Object.keys(files).filter(k=>/^xl\/charts\/chart\d+\.xml$/.test(k));eq(paths.length,1,'표준 차트 파트');const xml=textOf(files[paths[0]]),tree=parseXml(xml);o.refs=descendants(tree,'f').map(n=>n.text);o.cacheCounts=descendants(tree,'ptCount').map(n=>Number(n.attrs.val));const valueCache=descendants(descendants(tree,'val')[0],'numCache')[0];o.valueCache={count:Number(descendants(valueCache,'ptCount')[0]?.attrs.val),last:descendants(valueCache,'pt').at(-1)?.attrs.idx,lastValue:descendants(descendants(valueCache,'pt').at(-1),'v')[0]?.text};
  for(const path of Object.keys(files).filter(k=>/^xl\/(charts|drawings)\/[^/]+\.xml$/.test(k)))files[path]=textOf(files[path]).replace(/<(?:c|cx|a):extLst>[\s\S]*?<\/(?:c|cx|a):extLst>/g,'');
  const reopened=new Workbook(readXlsx(zip(files)).data),reHost=reopened.sheets.findIndex(s=>s.charts?.length),reChart=reopened.sheets[reHost]?.charts[0];ok(reChart,'위셀 보조 확장 없이 표준 차트 재열기');
  eq(ch.range.r2,3000,'차트 원본 범위는 마지막 3000행 포함');eq(o.valueCache,{count:3000,last:'2999',lastValue:'99999'},'표준 Excel 캐시도 전체 개수와 원래 마지막 인덱스 보존');eq(o.model.last,99999,'그리기 데이터의 마지막 outlier');ok(o.refs.some(f=>f.endsWith('$B$2:$B$3001')),'표준 OOXML 값 참조의 전체 3000행');eq(reopened.getValue(0,3000,1),99999,'마지막 원본 값 저장');eq(chartModelData(reopened,reHost,reChart).series[0].values.at(-1),99999,'표준 XLSX 재열기 마지막 차트 값');eq(reopened.getValue(reopened.sheetIndexByName('보존'),0,0),'수정 금지','다른 시트 보존');
 });
}finally{await browser.close();const report={engine,url,baseline:baseline?'app.js only from '+process.env.WIXEL_DATA_OPTIONS_BASELINE:null,cases:results.length,passed:results.filter(r=>r.ok).length,checks,results};await writeFile(out+'/result.json',JSON.stringify(report,null,2));console.log(JSON.stringify({engine,cases:report.cases,passed:report.passed,checks,out}));if(report.passed!==report.cases)process.exitCode=1;}
