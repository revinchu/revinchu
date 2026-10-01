// 현재 PC의 합성 문서 측정. 공인 성능 점수/실제 사용자 문서 검사가 아니다.
// CPU4x, 실제 키 입력 + 합성 ClipboardEvent, 서버 쓰기 금지. 접근성 비교는 DOM 갱신 비용이다.
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost','127.0.0.1','[::1]'].includes(new URL(base).hostname)) throw new Error('성능 검사는 로컬 서버에서만 실행합니다.');
const sizes = (process.env.WIXEL_PERF_CELLS || '5000,50000,200000').split(',').map(Number);
if (sizes.some(n => !Number.isInteger(n) || n < 100 || n > 1000000 || n % 10)) throw new Error('셀 수는 100~1000000 사이의 10의 배수입니다.');
const cpu = Number(process.env.WIXEL_PERF_CPU || 4);
const output = path.resolve(process.env.WIXEL_PERF_OUTPUT || '../.local/performance-regression.json');
const profileOps = new Set((process.env.WIXEL_PERF_PROFILE || '').split(','));
const browser = await chromium.launch();
const report = { generated: new Date().toISOString(), platform: `${os.platform()} ${os.release()}`, cpu: os.cpus()[0]?.model, logicalCpus: os.cpus().length, browser: await browser.version(), url:base, throttle:cpu, autosave:process.env.WIXEL_PERF_AUTOSAVE !== 'off', note:'현재 PC·합성 숫자/수식/서식 자료. wallMs는 Playwright 왕복·두 프레임 대기를 포함. maxLongTaskMs는 브라우저 main thread 50ms 이상 작업이며 CDP 직접 실행은 수집에서 빠질 수 있으므로 0이 비차단을 뜻하지 않는다. calls는 계측한 동기 함수 시간이고 중첩 호출을 합산하면 중복된다. 입력은 12345+Enter, 붙여넣기는 1000셀. AX는 접근성 DOM 갱신 비교이며 실제 화면 읽기 프로그램 비용 아님. heap은 CDP 지원 시 현재 JS heap이며 peak가 아님.', cases:[] };
const round = n => Math.round(n*10)/10;
try { for (const count of sizes) {
  const context = await browser.newContext({ viewport:{width:1400,height:900} }), page = await context.newPage(), errors=[], writes=[];
  page.setDefaultTimeout(60000); page.on('pageerror', e=>errors.push(e.message));
  await context.route('**/*', route=>{if(!['GET','HEAD','OPTIONS'].includes(route.request().method())){writes.push(route.request().url());return route.abort();}return route.continue();});
  const entry = { cells:count, rows:count/10, columns:10, formulaCells:(count/10-1)*2, metrics:[], errors, writes };
  report.cases.push(entry);
  let cdp;
  try {
    await page.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
    await page.goto(base,{waitUntil:'domcontentloaded',timeout:60000}); await page.waitForFunction(()=>window.tabula?.wb());
    if(!report.autosave) await page.locator('#autosaveToggle').click();
    cdp = await context.newCDPSession(page); await cdp.send('Emulation.setCPUThrottlingRate',{rate:cpu}); await cdp.send('Performance.enable');
    await page.evaluate(()=>{
      window.__perf={long:[],calls:{},started:0};
      new PerformanceObserver(list=>{for(const e of list.getEntries())window.__perf.long.push({start:e.startTime,duration:e.duration});}).observe({type:'longtask',buffered:false});
      const instrument=(object,key,label)=>{const original=object[key];if(typeof original!=='function')return;object[key]=function(...args){const start=performance.now();try{return original.apply(this,args);}finally{const t=window.__perf.calls[label]??={count:0,ms:0,max:0};const ms=performance.now()-start;t.count++;t.ms+=ms;t.max=Math.max(t.max,ms);}};};
      const t=window.tabula,w=t.wb(),v=t.gv();
      for(const key of ['restore','setInput','transact','invalidate','invalidateChanged','serialize','recalculate'])instrument(w,key,'wb.'+key);
      for(const key of ['layout','renderPane','renderSelection','update'])instrument(v,key,'view.'+key);
      if(v.a11y)instrument(v.a11y,'update','a11y.update');
      window.__perfDigest=()=>{let hash=2166136261,count=0;w.sheets[0].cells.forEachRC((cell,r,c)=>{const str=r+','+c+':'+cell.raw+':'+JSON.stringify(cell.style??{});for(let i=0;i<str.length;i++){hash^=str.charCodeAt(i);hash=Math.imul(hash,16777619);}count++;});return {count,hash:hash>>>0};};
    });
    const measure = async (name, fn) => {
      await page.evaluate(()=>{window.__perf.started=performance.now();window.__perf.long=[];window.__perf.calls={};});
      const profile=profileOps.has(name);if(profile){await cdp.send('Profiler.enable');await cdp.send('Profiler.start');}
      const value = await fn();
      await page.evaluate(()=>new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(()=>setTimeout(resolve,0)))));
      const metric=await page.evaluate(()=>{const p=window.__perf,l=p.long.filter(x=>x.start>=p.started);return {wallMs:performance.now()-p.started,longTaskCount:l.length,maxLongTaskMs:l.reduce((m,x)=>Math.max(m,x.duration),0),totalLongTaskMs:l.reduce((n,x)=>n+x.duration,0),calls:p.calls};});
      if(profile){const {profile:data}=await cdp.send('Profiler.stop');const file=output.replace(/\.json$/,`-${count}-${name}.cpuprofile`);await fs.mkdir(path.dirname(file),{recursive:true});await fs.writeFile(file,JSON.stringify(data));metric.profile=file;}
      try{const {metrics}=await cdp.send('Performance.getMetrics');const h=metrics.find(x=>x.name==='JSHeapUsedSize');if(h)metric.jsHeapMiB=round(h.value/1048576);}catch{}
      for(const key of ['wallMs','maxLongTaskMs','totalLongTaskMs'])metric[key]=round(metric[key]);
      for(const v of Object.values(metric.calls)){v.ms=round(v.ms);v.max=round(v.max);}
      entry.metrics.push({name,...metric});console.log(`${count} ${name}: ${metric.wallMs}ms; longest task ${metric.maxLongTaskMs}ms; heap ${metric.jsHeapMiB??'n/a'}MiB`);return value;
    };
    await measure('load',()=>page.evaluate(count=>{
      const cells={},rows=count/10;
      for(let r=0;r<rows;r++)for(let c=0;c<10;c++){
        let raw=r===0?'열'+(c+1):c===0?(r%2?'홀수':'짝수'):c===8?`=SUM(B${r+1}:H${r+1})`:c===9?`=I${r+1}*2`:String(r*10+c);
        cells[`${r},${c}`]={raw,...((r+c)%5===0?{style:{numFmt:c?'comma':'general',decimals:0,bold:r%3===0,fill:r%2?'#edf4fa':'#fff2cc'}}:{})};
      }
      const t=window.tabula,w=t.wb();w.restore({sheets:[{name:'성능 합성 검사',cells}]});w.undoStack=[];w.redoStack=[];t.gv().layout();t.gv().renderAll();t.selectCell(1,1);window.__perfRows=rows;
    },count));
    const baseline=await measure('audit-all-values',()=>page.evaluate(()=>{const w=window.tabula.wb();for(let r=1;r<window.__perfRows;r++){for(let c=1;c<=7;c++)if(w.getValue(0,r,c)!==r*10+c)throw new Error(`숫자 불일치 ${r},${c}`);if(w.getValue(0,r,8)!==r*70+28||w.getValue(0,r,9)!==(r*70+28)*2)throw new Error('수식 불일치 '+r);}return window.__perfDigest();}));
    assert.equal(baseline.count,count);
    await page.locator('#cellEditor').focus();
    await measure('input',async()=>{await page.keyboard.type('12345');await page.keyboard.press('Enter');});
    assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,1,1)),12345);
    assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,1,8)),12432);
    await measure('undo-input',()=>page.keyboard.press('Control+z'));assert.deepEqual(await page.evaluate(()=>window.__perfDigest()),baseline);
    const pasteRows=Math.min(100,entry.rows-2);
    await page.evaluate(()=>window.tabula.selectCell(1,0));
    await measure('paste',()=>page.evaluate(n=>{const dt=new DataTransfer(),rows=[];for(let r=0;r<n;r++){const row=[];for(let c=0;c<10;c++)row.push(String(100000+r*10+c));rows.push(row.join('\t'));}dt.setData('text/plain',rows.join('\n'));document.getElementById('cellEditor').dispatchEvent(new ClipboardEvent('paste',{clipboardData:dt,bubbles:true,cancelable:true}));},pasteRows));
    assert.equal(await page.evaluate(r=>window.tabula.wb().getValue(0,r,9),pasteRows),100000+(pasteRows-1)*10+9);
    await measure('undo-paste',()=>page.keyboard.press('Control+z'));assert.deepEqual(await page.evaluate(()=>window.__perfDigest()),baseline);
    await measure('scroll',()=>page.evaluate(async()=>{const v=window.tabula.gv();for(const ratio of [.25,.5,.75,.98,0]){v.scroll.scrollTop=v.rows.pos(Math.floor(window.__perfRows*ratio));await new Promise(resolve=>requestAnimationFrame(()=>requestAnimationFrame(resolve)));}}));
    assert.deepEqual(await page.evaluate(()=>window.__perfDigest()),baseline);
    await page.evaluate(()=>window.tabula.selectCell(0,0));
    await measure('filter-create',()=>page.evaluate(()=>window.tabula.run('toggleFilter')));
    await measure('filter-menu',()=>page.keyboard.press('Alt+ArrowDown'));
    const menu=page.locator('.filter-menu');await menu.getByRole('checkbox',{name:'표시된 항목 모두 선택',exact:true}).uncheck();await menu.getByRole('checkbox',{name:'짝수',exact:true}).check();
    await measure('filter-apply',()=>menu.getByRole('button',{name:'확인',exact:true}).click());
    assert.equal(await page.evaluate(()=>{const h=window.tabula.wb().sheets[0].filter.hidden;return h.__bits?h.count:Object.values(h).filter(Boolean).length;}),Math.ceil((entry.rows-1)/2));
    await measure('undo-filter',async()=>{await page.keyboard.press('Control+z');await page.keyboard.press('Control+z');});assert.deepEqual(await page.evaluate(()=>window.__perfDigest()),baseline);
    await measure('serialize',()=>page.evaluate(()=>{window.__perfSaved=JSON.stringify(window.tabula.wb().serialize());return window.__perfSaved.length;})).then(n=>entry.serializedCharacters=n);
    await measure('reopen',()=>page.evaluate(()=>{const t=window.tabula;t.wb().restore(JSON.parse(window.__perfSaved));t.gv().layout();t.gv().renderAll();t.selectCell(1,1);}));
    assert.deepEqual(await page.evaluate(()=>window.__perfDigest()),baseline);
    assert.equal(await page.evaluate(()=>window.tabula.wb().getValue(0,window.__perfRows-1,9)),((entry.rows-1)*70+28)*2);
    entry.accessibility=await page.evaluate(()=>{
      const t=window.tabula,v=t.gv(),a=v.a11y;if(!a)return {supported:false};const original=a.update,results={on:[],off:[]};
      try{for(const mode of ['off','on','on','off']){a.update=mode==='on'?original:()=>{};for(let i=0;i<12;i++){const start=performance.now();t.selectCell(1+i%7,1+i%5,{scroll:false});if(i>=2)results[mode].push(performance.now()-start);}}}finally{a.update=original;a.update();}
      const summary=xs=>{xs.sort((a,b)=>a-b);return {medianMs:xs[Math.floor(xs.length/2)],p95Ms:xs[Math.floor(xs.length*.95)],samples:xs.length};};return {supported:true,on:summary(results.on),off:summary(results.off),logicalCells:a.nodes.size};
    });
    for(const mode of ['on','off'])if(entry.accessibility[mode])for(const key of ['medianMs','p95Ms'])entry.accessibility[mode][key]=round(entry.accessibility[mode][key]);
    console.log(`${count} AX DOM: ${JSON.stringify(entry.accessibility)}`);
    assert.deepEqual(errors,[]);assert.deepEqual(writes,[]);entry.ok=true;
  }catch(e){entry.ok=false;entry.failure=e.stack;console.error('NG '+count+': '+e.stack);}finally{await context.close();}
  await fs.mkdir(path.dirname(output),{recursive:true});await fs.writeFile(output,JSON.stringify(report,null,2));
} } finally { await browser.close(); }
console.log(JSON.stringify({total:report.cases.length,ok:report.cases.filter(x=>x.ok).length,output}));
if(report.cases.some(x=>!x.ok))process.exitCode=1;
