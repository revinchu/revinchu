// 선형 예측/추세선: 격리 합성 문서, 실제 SVG 래스터와 축 경계를 검사합니다.
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const engine=process.env.WIXEL_BROWSER||'chromium', url=process.env.WIXEL_URL||'http://127.0.0.1:5195/', origin=new URL(url).origin;
const publicCheck=process.env.WIXEL_FORECAST_PUBLIC==='1'&&origin==='https://wixel-3.wizx.workers.dev';
if(!['127.0.0.1','localhost','[::1]'].includes(new URL(url).hostname)&&!publicCheck) throw new Error('예측선 검사는 로컬 서버 또는 명시적으로 허용한 위셀 공개 주소에서만 실행하세요.');
if(publicCheck&&process.env.WIXEL_FORECAST_BASELINE) throw new Error('공개 서비스 검사에는 기준 렌더러를 주입하지 마세요.');
const out=process.env.WIXEL_FORECAST_OUT||'D:/Codex/Temp/wixel-chart-forecast-20261004/'+engine;
const baseline=process.env.WIXEL_FORECAST_BASELINE?await readFile(process.env.WIXEL_FORECAST_BASELINE,'utf8'):null;
await mkdir(out,{recursive:true});
const browser=await pw[engine].launch(), results=[],assets=new Set(); let checks=0;
const eq=(a,b,message)=>{assert.deepEqual(a,b,message);checks++;};
const ok=(value,message)=>{assert.ok(value,message);checks++;};
const chart=p=>p.locator('.pane-br .obj.chart[data-id="forecast-test"]');
const saved=p=>p.evaluate(()=>JSON.stringify(tabula.wb().serialize()));
async function fixture(p, spec={}) {
  const options={type:'line',values:[10,20,30,40,50,60],forward:4,kind:'linear',width:600,height:360,...spec};
  await p.evaluate(o=>{
    const w=tabula.wb(),cells={},n=o.values.length,names=Array.from({length:n},(_,i)=>String(i+1)+'월');
    cells['0,0']={raw:'월'};cells['0,1']={raw:'실적'};
    const x=o.x||o.values.map((_,i)=>i+1);
    for(let i=0;i<n;i++){cells[(i+1)+',0']={raw:o.type==='scatter'?String(x[i]):names[i]};cells[(i+1)+',1']={raw:String(o.values[i]??'')};}
    const ref=c=>({r1:1,c1:c,r2:n,c2:c});
    let series=[{name:{text:'실적'},cat:ref(0),val:ref(1),...(o.type==='scatter'?{x:ref(0)}:{})}];
    const trend={trend:o.kind,trendForward:o.forward,trendPeriod:3,trendColor:'#e000d0',marker:'circle'};
    let seriesFmt=[trend];
    if(o.secondary){cells['0,2']={raw:'보조 실적'};for(let i=0;i<n;i++)cells[(i+1)+',2']={raw:String(1000+(o.values[i]??0)*100)};series.push({name:{text:'보조 실적'},cat:ref(0),val:ref(2)});seriesFmt=[{type:'column',axis:0},{...trend,type:'line',axis:1}];}
    const ch={id:'forecast-test',type:o.type,x:40,y:20,w:o.width,h:o.height,z:1,title:'월별 실적과 예측',legend:'b',series,seriesFmt,axes:o.axes||{},gridY:true,gridX:false,labels:false,marker:'circle',scatterStyle:'lineMarker'};
    w.restore({sheets:[{name:'합성 예측',cells,charts:[ch]},{name:'빈 시트',cells:{}}]});
    tabula.switchSheet(1);tabula.switchSheet(0);tabula.gv().setZoom(100);w.undoStack=[];w.redoStack=[];
  },options);
  await chart(p).waitFor();await p.evaluate(()=>document.fonts.ready);await p.evaluate(()=>new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r))));
  return options;
}
async function inspect(p, options) {
  return chart(p).locator('svg').first().evaluate(async(svg,o)=>{
    const horizontal=o.type==='bar';
    const grid=[...svg.querySelectorAll('line[stroke="#d9d9d9"]')].filter(l=>horizontal?l.getAttribute('x1')===l.getAttribute('x2'):l.getAttribute('y1')===l.getAttribute('y2'));
    const coordinates=grid.flatMap(l=>[[Number(l.getAttribute('x1')),Number(l.getAttribute('y1'))],[Number(l.getAttribute('x2')),Number(l.getAttribute('y2'))]]);
    const area={x:Math.min(...coordinates.map(p=>p[0])),y:Math.min(...coordinates.map(p=>p[1])),right:Math.max(...coordinates.map(p=>p[0])),bottom:Math.max(...coordinates.map(p=>p[1]))};
    area.w=area.right-area.x;area.h=area.bottom-area.y;
    const paths=[...svg.querySelectorAll('path[stroke="#e000d0"]')];
    const trends=paths.map(path=>{
      const length=path.getTotalLength(),start=path.getPointAtLength(0),end=path.getPointAtLength(length),b=path.getBBox();let node=path,clip=null;
      while(node&&node!==svg){if(node.getAttribute('clip-path')){clip=node.getAttribute('clip-path');break;}node=node.parentElement;}
      const id=clip?.match(/url\(["']?#([^"')]+)/)?.[1],cp=id?[...svg.querySelectorAll('clipPath')].find(n=>n.id===id):null,r=cp?.querySelector('rect');
      return {d:path.getAttribute('d'),start:{x:start.x,y:start.y},end:{x:end.x,y:end.y},length,bounds:{x:b.x,y:b.y,w:b.width,h:b.height},clip,clipUnits:cp?.getAttribute('clipPathUnits'),clipRect:r?{x:Number(r.getAttribute('x')),y:Number(r.getAttribute('y')),w:Number(r.getAttribute('width')),h:Number(r.getAttribute('height'))}:null};
    });
    const si=o.secondary?1:0,last=o.values.length-1,mark=svg.querySelector('[data-s="'+si+'"][data-p="'+last+'"]');
    const mb=mark?.getBBox(),lastData=mb?{x:mb.x+mb.width/2,y:mb.y+mb.height/2}:null;
    const source=new XMLSerializer().serializeToString(svg),img=new Image();
    await new Promise((resolve,reject)=>{img.onload=resolve;img.onerror=()=>reject(new Error('합성 SVG 래스터 실패'));img.src='data:image/svg+xml;charset=utf-8,'+encodeURIComponent(source);});
    const cv=document.createElement('canvas');cv.width=Number(svg.getAttribute('width'));cv.height=Number(svg.getAttribute('height'));const ctx=cv.getContext('2d');ctx.drawImage(img,0,0);const pixels=ctx.getImageData(0,0,cv.width,cv.height).data;
    let colored=0,outside=0;const examples=[];
    for(let y=0;y<cv.height;y++)for(let x=0;x<cv.width;x++){const k=(y*cv.width+x)*4;if(pixels[k]>140&&pixels[k+1]<100&&pixels[k+2]>140&&pixels[k+3]>100){colored++;if(x+.5<area.x-1.5||x+.5>area.right+1.5||y+.5<area.y-1.5||y+.5>area.bottom+1.5){outside++;if(examples.length<8)examples.push([x,y]);}}}
    return {area,trends,lastData,colored,outside,examples,width:cv.width,height:cv.height,png:cv.toDataURL('image/png'),axes:[...svg.querySelectorAll('[data-el^="axis-"] text')].map(n=>n.textContent)};
  },options);
}
async function verify(p, options, label) {
  const state=await inspect(p,options);await writeFile(out+'/'+label+'-raster.png',Buffer.from(state.png.split(',')[1],'base64'));delete state.png;
  eq(state.trends.length,options.type==='scatter'?0:1,'지원하는 추세선만 생성');
  if(options.type==='scatter'){ok(state.area.w>20&&state.area.h>20,'기존 XY 차트 그림 영역 유지');ok(await chart(p).locator('[data-s="0"]').count()>0,'기존 XY 계열 표시 유지');return state;}
  eq(state.outside,0,'예측선 픽셀이 그림 영역 밖에 표시되지 않음: '+JSON.stringify(state.examples));
  ok(state.colored>10,'추세선은 그림 영역 안에 실제로 표시');
  const t=state.trends[0];ok(t.clip&&t.clipRect,'추세선 clip-path가 실제 rect로 연결');eq(t.clipUnits,'userSpaceOnUse','절대 SVG 좌표 기준 clip');
  ok(t.length>5&&!/NaN|Infinity/.test(t.d),'유한한 추세선 경로');
  for(const key of ['x','y','w','h'])ok(Math.abs(t.clipRect[key]-state.area[key])<.25,'추세선 clip '+key+'가 실제 그리드 경계와 같음');
  if(!options.axes?.y?.min&&!options.axes?.y?.max&&!options.axes?.y2?.max&&options.forward>0&&options.kind!=='movingAvg'){
    ok(state.lastData,'실제 마지막 데이터 점 존재');
    const coord=options.type==='bar'?'y':'x',reverse=!!options.axes?.x?.reverse;
    ok((t.end[coord]-state.lastData[coord])*(reverse?-1:1)>4,'앞으로 예측한 구간을 차트 안에 확보');
    ok(t.end.x>=state.area.x-.5&&t.end.x<=state.area.right+.5&&t.end.y>=state.area.y-.5&&t.end.y<=state.area.bottom+.5,'자동 축에서 예측 끝점이 그림 영역 안에 있음');
  }
  return state;
}
async function test(name,spec,run) {
  if(process.env.WIXEL_FORECAST_FILTER&&!name.includes(process.env.WIXEL_FORECAST_FILTER))return;
  const context=await browser.newContext({viewport:{width:1280,height:900},deviceScaleFactor:1,serviceWorkers:'block'}),p=await context.newPage(),errors=[],writes=[],start=checks;let options,observation;
  p.setDefaultTimeout(15000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.dismiss());
  await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1','off');});
  await context.route('**/*',route=>{const req=route.request(),u=new URL(req.url());if(!['GET','HEAD','OPTIONS'].includes(req.method())){writes.push(req.method()+' '+u.pathname);return route.abort();}if(u.origin!==origin||/^\/api(?:\/|$)/.test(u.pathname))return route.abort();if(baseline&&u.pathname==='/src/chart.js')return route.fulfill({status:200,contentType:'application/javascript; charset=utf-8',body:baseline});return route.continue();});
  try{
    await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.gv(),null,{timeout:60000});
    for(const asset of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(asset);
    options=await fixture(p,spec);const before=await saved(p);observation=await(run?run(p,options,name):verify(p,options,name));
    if(!run)eq(await saved(p),before,'렌더링은 원본 셀·축 설정·예측값 모델을 변경하지 않음');
    eq(errors,[],'페이지 오류 없음');eq(writes,[],'외부 쓰기 없음');await chart(p).screenshot({path:out+'/'+name+'.png'});
    results.push({name,ok:true,checks:checks-start,observation});console.log('OK '+name);
  }catch(e){observation=options?await inspect(p,options).catch(()=>null):null;if(observation?.png){await writeFile(out+'/'+name+'-raster.png',Buffer.from(observation.png.split(',')[1],'base64'));delete observation.png;}await p.screenshot({path:out+'/'+name+'-failure.png'}).catch(()=>{});results.push({name,ok:false,error:e.stack,errors,writes,checks:checks-start,observation});console.error('NG '+name+': '+e.message);}finally{await context.close();}
}
try{
  await test('line-positive',{});
  await test('line-negative',{values:[60,45,30,15,0,-15]});
  await test('column-forward',{type:'column'});
  await test('bar-forward',{type:'bar'});
  await test('reverse-category',{axes:{x:{reverse:true}}});
  await test('fixed-axis',{axes:{y:{min:0,max:60,major:10}}});
  await test('fixed-secondary',{type:'combo',secondary:true,axes:{y2:{min:0,max:6000,major:1000}}});
  await test('auto-secondary',{type:'combo',secondary:true});
  await test('small-chart',{width:260,height:180});
  await test('long-forward',{forward:100});
  await test('exponential',{kind:'exp',values:[2,4,8,16,32,64],forward:3});
  await test('moving-average',{kind:'movingAvg',values:[10,40,20,60,30,80],forward:0});
  await test('xy-original',{type:'scatter',x:[2,4,8,16,32,64],values:[10,40,20,60,30,80]});
  await test('menu-forecast-undo',{kind:null,forward:0},async(p,options,name)=>{
    const before=await saved(p);await chart(p).click({position:{x:20,y:10}});await chart(p).locator('.ch-sb[data-a="elements"]').click();
    await p.locator('.ce-sub[title="추세선 옵션"]').click();await p.getByRole('menuitem',{name:'선형 예측',exact:true}).click();
    const trend=await p.evaluate(()=>tabula.wb().sheets[0].charts[0].seriesFmt[0]);eq(trend.trend,'linear','메뉴가 선형 추세선을 설정');eq(trend.trendForward,2,'메뉴가 앞으로 2구간 예측 설정');
    const observation=await verify(p,{...options,kind:'linear',forward:2},name);eq(await p.evaluate(()=>tabula.wb().undoStack.length),1,'메뉴 적용은 한 번의 Undo');
    await p.evaluate(()=>tabula.run('undo'));eq(await saved(p),before,'Undo가 추세선 추가 전 문서 복원');eq(await chart(p).locator('path[stroke="#e000d0"]').count(),0,'Undo는 예측선 제거');
    await p.evaluate(()=>tabula.run('redo'));await verify(p,{...options,kind:'linear',forward:2},name+'-redo');return observation;
  });
}finally{
  await browser.close();const result={engine,url,publicCheck,baseline:!!baseline,assets:[...assets],cases:results.length,passed:results.filter(r=>r.ok).length,checks,results};await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify({engine,cases:result.cases,passed:result.passed,checks,out,assets:result.assets,baseline:!!baseline}));if(result.cases!==result.passed||assets.size!==1)process.exitCode=1;
}
