// 새 피벗 원본 대화상자가 이전 저장 캐시를 재사용하지 않는지 실제 UI로 검사한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=new URL(process.env.WIXEL_URL||'http://127.0.0.1:5191/');
if(!['localhost','127.0.0.1','[::1]'].includes(url.hostname))throw Error('로컬 서버만 검사합니다.');
const out=process.env.WIXEL_PIVOT_SOURCE_OUT||'D:/Codex/Temp/wixel-final-audit/pivot-source-browser';await mkdir(out,{recursive:true});
const rows=(header,value,second)=>({'0,0':{raw:header[0]},'0,1':{raw:header[1]},'1,0':{raw:'Item'},'1,1':{raw:String(value)},...(second?{'2,0':{raw:'Other'},'2,1':{raw:String(second)}}:{})});
const fixture={pivotSnapshots:{saved:[['Category','Amount'],['Item',10]]},sheets:[
 {name:'A',cells:rows(['Category','Amount'],999,888)},
 {name:'B',cells:rows(['Category','Amount'],123),tables:[{id:'tb',name:'TableB',r1:0,c1:0,r2:1,c2:1,header:true,totals:false}]},
 {name:'Pivot',cells:rows(['Category','Amount'],10),pivot:{name:'Pivot1',snapshotId:'saved',source:'A',range:{r1:0,c1:0,r2:1,c2:1},rows:['Category'],cols:[],values:[{field:'Amount',agg:'sum'}],top:0,left:0,area:{r1:0,c1:0,r2:2,c2:1}}},
 {name:'C',cells:rows(['NewCategory','NewAmount'],321)}
]};
const reports=[];
for(const engine of (process.env.WIXEL_BROWSER||'chromium,webkit').split(',')){
 const browser=await pw[engine].launch(),context=await browser.newContext({viewport:{width:1200,height:900},serviceWorkers:'block'}),page=await context.newPage(),errors=[],blocked=[];let checks=0;
 const eq=(a,b,m)=>{checks++;assert.deepEqual(a,b,m);};const settle=()=>page.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 await context.route('**/*',r=>{const q=r.request(),u=new URL(q.url());if(u.origin!==url.origin||u.pathname.startsWith('/api/')||!['GET','HEAD'].includes(q.method())){blocked.push(q.method()+' blocked');return r.abort();}return r.continue();});page.on('pageerror',e=>errors.push(e.message));
 try{
  await page.goto(url.href);await page.waitForFunction(()=>window.tabula?.wb());
  for(const scenario of [{input:'A!A1:B3',expected:[999,888],source:'A'},{input:'B!A1:B2',expected:[123],source:'B'},{input:'TableB',expected:[123],source:'B',table:'TableB'},{input:'C!A1:B2',expected:[321],source:'C',emptyFields:true}]){
   await page.evaluate(data=>{tabula.wb().load(data);tabula.switchSheet(2);tabula.selectCell(1,1);tabula.gv().renderAll();},fixture);await settle();
   await page.evaluate(()=>tabula.run('pivotChangeSource'));await page.locator('#dialogLayer input[name="src"]').fill(scenario.input);await page.locator('#dialogLayer .dialog').last().getByRole('button',{name:'확인',exact:true}).click();await settle();
   const actual=await page.evaluate(async()=>{const{pivotSourceData}=await import('/src/pivot.js');const w=tabula.wb(),d=w.sheets[2].pivot,s=pivotSourceData(w,d);return{snapshotId:d.snapshotId??null,source:d.source,table:d.table??null,values:Array.from({length:s.cube.n},(_,i)=>s.cube.row(i)[1]),rows:d.rows,fields:d.values};});
   eq(actual.snapshotId,null,'새 원본에서 이전 캐시 연결 제거');eq(actual.source,scenario.source,'선택 원본');eq(actual.table,scenario.table??null,'표 원본 연결');eq(actual.values,scenario.expected,'새 원본 데이터 반영');
   if(scenario.emptyFields){eq(actual.rows,[],'새 원본에 없는 행 필드 제거');eq(actual.fields,[],'새 원본에 없는 값 필드 제거');}
   await page.evaluate(()=>tabula.run('undo'));await settle();
   eq(await page.evaluate(async()=>{const{pivotSourceData}=await import('/src/pivot.js');const w=tabula.wb();return pivotSourceData(w,w.sheets[2].pivot).cube.row(0)[1];}),10,'실행 취소로 이전 저장 캐시 복원');
  }
  eq(errors,[],'페이지 오류 없음');eq(blocked,[],'원격 전송 없음');await page.screenshot({path:out+'/'+engine+'.png'});reports.push({engine,ok:true,checks});
 }catch(e){reports.push({engine,ok:false,checks,error:e.message,errors,blocked});await page.screenshot({path:out+'/'+engine+'-failure.png'}).catch(()=>{});}finally{await browser.close();}
}
await writeFile(out+'/result.json',JSON.stringify(reports,null,2));console.log(JSON.stringify(reports));if(reports.some(r=>!r.ok))process.exitCode=1;
