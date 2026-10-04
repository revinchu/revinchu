// 차트 색 UI 회귀: 합성 문서·원격 쓰기 차단·실제 SVG 채우기/선/표식 검증.
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
const pw=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const engine=process.env.WIXEL_BROWSER||'chromium',url=process.env.WIXEL_URL||'http://127.0.0.1:5195/',origin=new URL(url).origin;
const publicCheck=process.env.WIXEL_COLORS_PUBLIC==='1'&&origin==='https://wixel-3.wizx.workers.dev';
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname)&&!publicCheck)throw new Error('격리 검사는 로컬 또는 명시적으로 허용한 위셀 공개 주소에서 실행하세요.');
if(publicCheck&&process.env.WIXEL_COLORS_BASELINE)throw new Error('공개 검증에 기준 코드를 주입하지 마세요.');
const out=process.env.WIXEL_COLORS_OUT||'D:/Codex/Temp/wixel-chart-colors-20261004/'+engine;
await mkdir(out,{recursive:true});
const baseline=new Map();if(process.env.WIXEL_COLORS_BASELINE)for(const f of ['app.js','chart.js','chart-edit.js','chart-selection-ui.js'])baseline.set('/src/'+f,await readFile(process.env.WIXEL_COLORS_BASELINE+'/'+f,'utf8'));
const browser=await pw[engine].launch(),results=[],assets=new Set();let checks=0;
const eq=(a,b,m)=>{assert.deepEqual(a,b,m);checks++;},ok=(v,m)=>{assert.ok(v,m);checks++;};
const palettes={modern:['#4f46e5','#0ea5e9','#10b981'],vivid:['#2563eb','#dc2626','#16a34a']};
const names={modern:'WIXEL 모던',vivid:'WIXEL 비비드'};
const obj=p=>p.locator('.pane-br .obj.chart[data-id="color-test"]');
const model=p=>p.evaluate(()=>structuredClone(tabula.wb().sheets[0].charts[0]));
const depth=p=>p.evaluate(()=>tabula.wb().undoStack.length);
const run=(p,cmd)=>p.evaluate(cmd=>tabula.run(cmd),cmd);
const color=(loc,value)=>loc.evaluate((el,value)=>{el.value=value;el.dispatchEvent(new Event('change',{bubbles:true}));},value);
async function fixture(p,spec={}){
 const options={type:'column',imported:true,...spec};
 await p.evaluate(o=>{
  const w=tabula.wb(),cells={};[['월','매출','방문'],['1월','10','50'],['2월','20','30'],['3월','30','40']].forEach((row,r)=>row.forEach((raw,c)=>cells[r+','+c]={raw}));
  const ref=c=>({r1:1,c1:c,r2:3,c2:c}),series=[{name:{text:'매출'},cat:ref(0),val:ref(1)},...(o.type==='pie'?[]:[{name:{text:'방문'},cat:ref(0),val:ref(2)}])];
  const fmt=o.imported?[{color:'#aa0000',grad:{ang:90,stops:[[0,'#aa0000'],[1,'#ffaa00']]},colors:['#cc00cc','#00cccc','#cccc00'],pointColors:{1:'#008800'},markerColor:'#336699',lineWidth:2.75,marker:'circle',markerSize:7,numFmt:'0.0',labels:false},{color:'#665544',markerColor:'#993333',marker:'circle',labels:false}]:[];
  if(o.gradientOnly){delete fmt[0].colors;delete fmt[0].pointColors;}
  const ch={id:'color-test',type:o.type,x:60,y:30,w:600,h:360,z:1,title:'월별 실적 색 검증',legend:'b',series,seriesFmt:fmt,axes:{y:{min:0,max:60,major:20}},palette:o.palette||(o.imported?['#994422','#225599','#449922']:undefined),...(o.areas?{chartAreaFormat:{fillMode:'solid',fill:'#ffff00'},plotAreaFormat:{fillMode:'solid',fill:'#ffcccc'}}:{})};
  w.restore({sheets:[{name:'합성 차트 색',cells,charts:[ch]},{name:'다른 시트',cells:{'0,0':{raw:'보존'}}}]});tabula.switchSheet(1);tabula.switchSheet(0);tabula.gv().setZoom(100);w.undoStack=[];w.redoStack=[];
  window.__colorsBeforeCells=JSON.stringify(w.serialize().sheets.map(s=>s.cells));
 },options);
 await obj(p).waitFor();await obj(p).click({position:{x:20,y:10}});return options;
}
async function colors(p){return obj(p).locator('svg').first().evaluate(svg=>({
 points:[...svg.querySelectorAll('[data-s][data-p]:not(text)')].filter(n=>['rect','path','circle','polygon'].includes(n.tagName)).map(n=>({s:Number(n.dataset.s),p:Number(n.dataset.p),fill:(n.getAttribute('fill')||'').toLowerCase(),stroke:(n.getAttribute('stroke')||'').toLowerCase()})),
 lines:[...svg.querySelectorAll('path[data-s]:not([data-p])')].map(n=>({s:Number(n.dataset.s),stroke:(n.getAttribute('stroke')||'').toLowerCase(),fill:n.getAttribute('fill')})),
 legend:[...svg.querySelectorAll('[data-el="legend"] rect,[data-el="legend"] line')].map(n=>(n.getAttribute(n.tagName==='line'?'stroke':'fill')||'').toLowerCase()),
 gradients:[...svg.querySelectorAll('linearGradient stop')].map(n=>n.getAttribute('stop-color')),
 background:(svg.querySelector(':scope > rect')?.getAttribute('fill')||svg.querySelector('[data-el="chart"] rect')?.getAttribute('fill')||'').toLowerCase(),
}));}
async function palette(p,path,key){
 if(path==='ribbon'){await p.locator('[data-ribbon-tab="chartDesign"]').click();await p.locator('[data-ribbon-command="chartColorsBtn"]').click();}
 else if(path==='side'){await obj(p).locator('.ch-sb[data-a="styles"]').click();await p.locator('.cs-tabs').getByRole('button',{name:'색',exact:true}).click();}
 else if(path==='selection'){await run(p,'chartFormat');await p.getByLabel('서식을 지정할 차트 요소',{exact:true}).selectOption('chart');await p.getByLabel('색 구성',{exact:true}).selectOption(key);return;}
 else {await run(p,'chartFormat');await p.getByRole('button',{name:'차트 전체 옵션…',exact:true}).click();const d=p.getByRole('dialog',{name:'차트 서식',exact:true});await d.getByRole('tab',{name:'채우기 및 선',exact:true}).click();await d.getByLabel('색 구성',{exact:true}).selectOption(key);await d.getByRole('button',{name:'닫기',exact:true}).click();return;}
 await p.locator('.pal-row[title="'+names[key]+'"]').click();
}
async function expectPalette(p,type,key){
 const state=await colors(p),expected=palettes[key];ok(state.points.length>=3,'실제 데이터 도형 존재');
 for(const point of state.points)eq(point.fill,type==='pie'?expected[point.p]:expected[point.s],'팔레트가 실제 데이터 색에 반영: '+JSON.stringify(point));
 if(type==='line')for(const line of state.lines)eq(line.stroke,expected[line.s],'계열 선도 팔레트 반영');
 eq(state.legend,type==='pie'?expected:expected.slice(0,2),'범례 색과 데이터 색 일치');return state;
}
async function preserve(p,before){const after=await model(p);eq(after.series,before.series,'원본 계열과 참조 보존');eq(after.axes,before.axes,'값 축 설정 보존');for(const [i,f]of (before.seriesFmt||[]).entries())for(const key of ['lineWidth','marker','markerSize','numFmt','labels','axis','type','explode'])if(f[key]!==undefined)eq(after.seriesFmt[i][key],f[key],'색 변경은 '+key+' 보존');}
async function style(p,path,title){
 await obj(p).click({position:{x:20,y:10}});
 if(path==='ribbon'){await p.locator('[data-ribbon-tab="chartDesign"]').click();await p.getByRole('button',{name:'차트 스타일 더 보기',exact:true}).click();}
 else await obj(p).locator('.ch-sb[data-a="styles"]').click();
 await p.locator('.cs-chip[title="'+title+'"]').click();await p.keyboard.press('Escape');
}
async function test(name,spec,fn){
 if(process.env.WIXEL_COLORS_FILTER&&!name.includes(process.env.WIXEL_COLORS_FILTER))return;
 const context=await browser.newContext({viewport:{width:1366,height:950},deviceScaleFactor:1,serviceWorkers:'block'}),p=await context.newPage(),errors=[],writes=[],start=checks;let observation;
 p.setDefaultTimeout(30000);p.on('pageerror',e=>errors.push(e.message));p.on('dialog',d=>d.dismiss());
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;localStorage.setItem('wixel.mobile-work.v1','off');});
 await context.route('**/*',route=>{const req=route.request(),u=new URL(req.url());if(!['GET','HEAD','OPTIONS'].includes(req.method())){writes.push(req.method()+' '+u.pathname);return route.abort();}if(u.origin!==origin||/^\/api(?:\/|$)/.test(u.pathname))return route.abort();if(baseline.has(u.pathname))return route.fulfill({status:200,contentType:'application/javascript; charset=utf-8',body:baseline.get(u.pathname)});return route.continue();});
 try{await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await p.waitForFunction(()=>window.tabula?.gv(),null,{timeout:60000});for(const a of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src'))))assets.add(a);
  await fixture(p,spec);await obj(p).screenshot({path:out+'/'+name+'-before.png'});observation=await fn(p);eq(await p.evaluate(()=>JSON.stringify(tabula.wb().serialize().sheets.map(s=>s.cells))===window.__colorsBeforeCells),true,'원본 셀/다른 시트 보존');eq(errors,[],'페이지 오류 없음');eq(writes,[],'API·외부 쓰기 없음');await obj(p).screenshot({path:out+'/'+name+'-after.png'});results.push({name,ok:true,checks:checks-start,observation});console.log('OK '+name);
 }catch(e){observation={colors:await colors(p).catch(()=>null),model:await model(p).catch(()=>null)};await p.screenshot({path:out+'/'+name+'-failure.png'}).catch(()=>{});results.push({name,ok:false,error:e.stack,errors,writes,checks:checks-start,observation});console.error('NG '+name+': '+e.message);}finally{await context.close();}
}
try{
 await test('ribbon-imported-column',{},async p=>{const before=await model(p);await palette(p,'ribbon','modern');const state=await expectPalette(p,'column','modern');await preserve(p,before);eq(await depth(p),1,'팔레트 변경은 한 번의 Undo');return state;});
 await test('side-imported-pie',{type:'pie'},async p=>{await palette(p,'side','vivid');return expectPalette(p,'pie','vivid');});
 await test('format-imported-line',{type:'line'},async p=>{await palette(p,'format','modern');return expectPalette(p,'line','modern');});
 await test('repeat-current-palette',{palette:'vivid'},async p=>{await palette(p,'ribbon','vivid');return expectPalette(p,'column','vivid');});
 await test('ribbon-style-imported',{areas:true},async p=>{await style(p,'ribbon','WIXEL 카드');const state=await expectPalette(p,'column','modern'),ch=await model(p);eq(ch.chartAreaFormat,undefined,'스타일이 기존 차트 영역 직접 서식을 해제');eq(ch.plotAreaFormat,undefined,'스타일이 기존 그림 영역 직접 서식을 해제');eq(state.background,'#ffffff','카드의 배경 적용');return state;});
 await test('side-style-imported',{areas:true},async p=>{await style(p,'side','WIXEL 대시보드 다크');const state=await expectPalette(p,'column','vivid'),ch=await model(p);eq(ch.chartAreaFormat,undefined,'옆 스타일도 기존 차트 영역 직접 서식 해제');eq(ch.plotAreaFormat,undefined,'옆 스타일도 기존 그림 영역 직접 서식 해제');eq(state.background,'#0f172a','옆 스타일의 어두운 배경 실제 적용');return state;});
 await test('series-color-gradient',{gradientOnly:true},async p=>{await run(p,'chartFormat');await p.getByRole('button',{name:'차트 전체 옵션…',exact:true}).click();const d=p.getByRole('dialog',{name:'차트 서식',exact:true});await d.getByRole('tab',{name:'차트 옵션',exact:true}).click();await color(d.locator('[data-format-series="0"] input[type="color"]').first(),'#123456');const state=await colors(p);for(const point of state.points.filter(n=>n.s===0))eq(point.fill,'#123456','계열 색 지정이 기존 그라데이션을 대신함');return state;});
 await test('point-color-and-auto',{type:'pie'},async p=>{await run(p,'chartFormat');await p.getByLabel('서식을 지정할 차트 요소',{exact:true}).selectOption('series:0');await p.getByRole('button',{name:'개별 데이터 요소 선택',exact:true}).click();await p.getByLabel('데이터 요소',{exact:true}).selectOption('1');await color(p.getByLabel('선택한 요소 색',{exact:true}),'#123456');let state=await colors(p);eq(state.points.find(n=>n.p===1).fill,'#123456','선택한 원형 조각만 새 색 적용');eq(state.points.find(n=>n.p===0).fill,'#cc00cc','다른 조각의 직접 색 보존');await p.getByRole('button',{name:'요소 색 자동으로',exact:true}).click();state=await colors(p);eq(state.points.find(n=>n.p===1).fill,'#aa0000','자동 색은 현재 계열 색으로 복귀');return state;});
 await test('palette-undo-redo',{},async p=>{const before=await model(p),old=await colors(p);await palette(p,'ribbon','modern');await expectPalette(p,'column','modern');eq(await depth(p),1,'팔레트는 원자적 Undo');await run(p,'undo');eq(await model(p),before,'Undo는 가져온 색·그라데이션·요소 색까지 복원');eq((await colors(p)).points,old.points,'Undo 실제 데이터 색 복원');await run(p,'redo');return expectPalette(p,'column','modern');});
 await test('palette-protected',{},async p=>{await p.evaluate(()=>{const w=tabula.wb();w.setSheetProp(0,'protect',{on:true,allow:{objects:false}});w.undoStack=[];w.redoStack=[];});const before=await model(p),old=await colors(p);await palette(p,'ribbon','modern');eq(await model(p),before,'보호된 개체의 팔레트는 변경되지 않음');eq((await colors(p)).points,old.points,'보호된 개체 실제 색 유지');eq(await depth(p),0,'거부된 색 변경은 Undo 추가 없음');});
 await test('selection-format-palette',{},async p=>{const before=await model(p);await palette(p,'selection','vivid');const state=await expectPalette(p,'column','vivid');await preserve(p,before);return state;});
 await test('style-preview-matches-applied',{areas:true},async p=>{await p.locator('[data-ribbon-tab="chartDesign"]').click();const title='WIXEL 카드',selector='.rg-chip[title="'+title+'"]';const preview=await p.locator(selector).locator('svg').evaluate(svg=>({fills:[...svg.querySelectorAll('rect[data-s][data-p]')].map(n=>(n.getAttribute('fill')||'').toLowerCase()),background:(svg.querySelector(':scope > rect')?.getAttribute('fill')||'').toLowerCase()}));ok(preview.fills.length>=6,'스타일 미리보기에도 데이터 막대가 있음');await p.locator(selector).click();const state=await expectPalette(p,'column','modern');eq(preview.fills,state.points.map(n=>n.fill),'리본 미리보기와 실제 적용 데이터 색 일치');eq(preview.background,state.background,'리본 미리보기와 실제 적용 배경 일치');await obj(p).locator('.ch-sb[data-a="styles"]').click();const chip=p.locator('.cs-chip[title="WIXEL 대시보드 다크"]');const side=await chip.locator('svg').evaluate(svg=>({fills:[...svg.querySelectorAll('rect[data-s][data-p]')].map(n=>(n.getAttribute('fill')||'').toLowerCase()),background:(svg.querySelector(':scope > rect')?.getAttribute('fill')||'').toLowerCase()}));await chip.click();await p.keyboard.press('Escape');const applied=await expectPalette(p,'column','vivid');eq(side.fills,applied.points.map(n=>n.fill),'옆 스타일 미리보기와 실제 적용 색 일치');eq(side.background,applied.background,'옆 스타일 미리보기와 실제 적용 배경 일치');return applied;});
 await test('dark-to-basic-style',{imported:false},async p=>{const titleColor=()=>obj(p).locator('svg').first().locator('[data-el="title"] text').getAttribute('fill');const original=await colors(p),originalTitle=await titleColor();await style(p,'side','WIXEL 대시보드 다크');eq(await titleColor(),'#f8fafc','어두운 스타일 제목 색 적용');await style(p,'ribbon','스타일 1 (기본)');const state=await colors(p),ch=await model(p);eq(state.background,original.background,'기본 스타일 복귀 시 배경 복원');eq(await titleColor(),originalTitle,'기본 스타일 복귀 시 흰 제목 색 잔류 없음');eq(ch.titleColor,undefined,'기본 스타일은 이전 제목 색 해제');eq(ch.plotFill,undefined,'기본 스타일은 이전 그림 영역 배경 해제');return state;});
 await test('fresh-chart-palette-control',{imported:false},async p=>{await palette(p,'ribbon','modern');await expectPalette(p,'column','modern');await palette(p,'side','vivid');return expectPalette(p,'column','vivid');});
}finally{await browser.close();const result={engine,url,publicCheck,baseline:baseline.size>0,assets:[...assets],cases:results.length,passed:results.filter(r=>r.ok).length,checks,results};await writeFile(out+'/result.json',JSON.stringify(result,null,2));console.log(JSON.stringify({engine,cases:result.cases,passed:result.passed,checks,assets:result.assets,out}));if(result.passed!==result.cases||assets.size!==1)process.exitCode=1;}
