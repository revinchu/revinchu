// 항목 축 레이블: 합성 문서와 로컬 서버만 사용하며 원격 쓰기를 차단합니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium';
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const origin = new URL(url).origin;
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw new Error('차트 축 검사는 로컬 서버에서만 실행하세요.');
const out = process.env.WIXEL_AXIS_LABELS_OUT || 'D:/Codex/Temp/wixel-axis-labels/' + engine;
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch(), results = [], assets = new Set(); let checks = 0;
const eq = (a, b, message) => { assert.deepEqual(a, b, message); checks++; };
const ok = (v, message) => { assert.ok(v, message); checks++; };
const labels = ['브랜드검색 시드니면세점 PC', '브랜드검색 시드니면세점 모바일', '쇼핑검색 브레인올로지 PC', '쇼핑검색 브레인올로지 모바일', '파워링크 신규 방문 고객', '파워링크 재방문 고객', '일반키워드 어린이 건강식품', '브랜드키워드 성인 건강식품', '일반키워드 주말 할인행사', '브랜드키워드 평일 할인행사', '신규 캠페인 검색 광고', '재방문 캠페인 배너 광고'];
const compact = value => value.replace(/\s+/g, '');
const chart = p => p.locator('.pane-br .obj.chart[data-id="axis-test"]');
async function fixture(p, type, width, count, axis = {}, short = false, extra = {}) {
  const names = short ? ['서울', '부산', '대전', '광주', '제주', '인천'] : labels;
  const categories = count > names.length ? Array.from({length:count},(_,i)=>names[i%names.length]+' '+(i+1)) : names.slice(0,count);
  const data = { categories, series: type === 'boxWhisker' ? categories.map((name, i) => ({ name, values: [10+i, 20+i, 25+i, 32+i, 39+i] })) : [{ name: '광고비', values: categories.map((_, i) => 130 + i * 17) }] };
  if (type === 'combo') data.series.push({ name: '전환율', values: categories.map((_, i) => 2 + i * .2), type: 'line', axis: 1 });
  await p.evaluate(({ type, width, data, axis, extra }) => {
    const w = tabula.wb();
    w.restore({ sheets: [{ name: '합성 축 검증', cells: { '0,0': { raw:'축 레이블' } }, charts: [{ id:'axis-test', type, x:20, y:20, w:width, h:380, z:1, title:'카테고리별 광고 성과', legend:'none', axes: { x:axis }, snapshotData:data, ...extra }] }, { name:'빈 시트', cells:{} }] });
    tabula.switchSheet(1); tabula.switchSheet(0); tabula.gv().setZoom(100); w.undoStack=[]; w.redoStack=[];
  }, { type, width, data, axis, extra });
  await chart(p).waitFor();
  await p.evaluate(() => document.fonts.ready);
  await p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
  return categories;
}
async function inspect(p) {
  return chart(p).locator('svg').first().evaluate(svg => {
    const candidates = [...svg.querySelectorAll('[data-el="axis-x"] text')].filter(n => !n.hasAttribute('data-axis-title'));
    const primary = candidates.filter(n => n.hasAttribute('data-axis-label'));
    const textNodes = primary.length ? primary : candidates;
    const sr = svg.getBoundingClientRect();
    const labels = textNodes.map(node => {
      const clone = node.cloneNode(true); clone.querySelectorAll('title,desc').forEach(t => t.remove());
      const b = node.getBBox(), m = node.getScreenCTM();
      const quad = [[b.x,b.y], [b.x+b.width,b.y], [b.x+b.width,b.y+b.height], [b.x,b.y+b.height]].map(([x,y]) => ({x:m.a*x+m.c*y+m.e-sr.x, y:m.b*x+m.d*y+m.f-sr.y}));
      return { text:clone.textContent, lines:[...clone.querySelectorAll('tspan')].map(n=>n.textContent), full:node.getAttribute('aria-label'), index:node.getAttribute('data-axis-label'), quad, rotation:node.getAttribute('transform'), fontSize:getComputedStyle(node).fontSize };
    });
    // Rotated labels have overlapping axis-aligned boxes even when their actual text rectangles do not overlap.
    const overlap = (a,b) => {
      for(const poly of [a,b]) for(let i=0;i<poly.length;i++) {
        const p=poly[i],q=poly[(i+1)%poly.length],dx=q.x-p.x,dy=q.y-p.y,len=Math.hypot(dx,dy);
        if (!len) continue;
        const ax=-dy/len,ay=dx/len,pa=a.map(p=>p.x*ax+p.y*ay),pb=b.map(p=>p.x*ax+p.y*ay);
        if(Math.min(Math.max(...pa),Math.max(...pb))-Math.max(Math.min(...pa),Math.min(...pb)) <= .75) return false;
      }
      return true;
    };
    const overlaps=[];
    for(let i=0;i<labels.length;i++) for(let j=i+1;j<labels.length;j++) if(overlap(labels[i].quad, labels[j].quad)) overlaps.push([i,j]);
    const extras = [...svg.querySelectorAll('[data-el="legend"] text,[data-el="axis-x"] text:not([data-axis-label])')].filter(n=>!textNodes.includes(n)).map(node=>{const b=node.getBBox(),m=node.getScreenCTM();return {text:node.textContent,quad:[[b.x,b.y],[b.x+b.width,b.y],[b.x+b.width,b.y+b.height],[b.x,b.y+b.height]].map(([x,y])=>({x:m.a*x+m.c*y+m.e-sr.x,y:m.b*x+m.d*y+m.f-sr.y}))};});
    const extraOverlaps=[]; for(const [i,label] of labels.entries()) for(const other of extras) if(overlap(label.quad,other.quad)) extraOverlaps.push({index:i,text:other.text});
    return { width:sr.width,height:sr.height, labels, overlaps, extras, extraOverlaps };
  });
}
async function checkLabels(p, expected, options={}) {
  const state = await inspect(p);
  eq(state.labels.length, expected.length, '범주별 레이블 개수');
  eq(state.labels.map(x => compact(x.text)), expected.map(compact), '화면 글자는 말줄임 없이 원문 보존');
  ok(state.labels.every(n=>n.lines.every(line=>line.trim().length>0)),'공백만 있는 줄로 축 공간을 낭비하지 않음');
  for(const [i,node] of state.labels.entries()) {
    ok(node.quad.every(q => q.x >= -1 && q.y >= -1 && q.x <= state.width+1 && q.y <= state.height+1), '레이블 '+i+' SVG 경계 안에 표시: '+JSON.stringify(node));
  }
  if(!options.allowOverlap) eq(state.overlaps, [], '실제 회전한 텍스트 사각형끼리 겹치지 않음');
  return state;
}
async function test(name, run) {
  if(process.env.WIXEL_AXIS_CASE && !name.includes(process.env.WIXEL_AXIS_CASE)) return;
  const context=await browser.newContext({ viewport:{width:1400,height:1000}, deviceScaleFactor:1, serviceWorkers:'block' });
  const p=await context.newPage(), errors=[], writes=[], start=checks;
  p.setDefaultTimeout(15000); p.on('pageerror',e=>errors.push(e.message)); p.on('dialog',d=>d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC=true; window.WIXEL_SKIP_START=true; localStorage.setItem('wixel.mobile-work.v1','off'); });
  await context.route('**/*',r=> { const req=r.request(),target=new URL(req.url()); if(!['GET','HEAD','OPTIONS'].includes(req.method())) { writes.push(req.method()+' '+target.pathname); return r.abort(); } return target.origin===origin&&!/^\/api(?:\/|$)/.test(target.pathname)?r.continue():r.abort(); });
  let observation;
  try {
    await p.goto(url,{waitUntil:'domcontentloaded',timeout:60000}); await p.waitForFunction(()=>window.tabula?.gv(),null,{timeout:60000});
    for(const src of await p.locator('script[src]').evaluateAll(ns=>ns.map(n=>n.getAttribute('src')))) assets.add(src);
    observation=await run(p); eq(errors,[],'페이지 오류 없음'); eq(writes,[],'원격 쓰기 없음');
    results.push({name,ok:true,checks:checks-start,observation}); console.log('OK '+name);
  } catch(e) {
    observation=await inspect(p).catch(()=>null); await p.screenshot({path:out+'/failure-'+results.length+'.png'}).catch(()=>{});
    results.push({name,ok:false,error:e.stack,checks:checks-start,errors,writes,observation}); console.error('NG '+name+': '+e.message);
  } finally { await context.close(); }
}
try {
  for(const type of ['column','line','combo','bar','waterfall','boxWhisker']) for(const [width,count] of [[420,6],[640,8],[900,12]]) {
    await test(type+' '+width+'px 긴 범주 '+count+'개', async p=> { const cats=await fixture(p,type,width,count); const before=await p.evaluate(()=>JSON.stringify(tabula.wb().serialize())); const observation=await checkLabels(p,cats); eq(await p.evaluate(()=>JSON.stringify(tabula.wb().serialize())),before,'렌더링이 원본 모델을 바꾸지 않음'); await chart(p).screenshot({path:out+'/'+type+'-'+width+'.png'}); return observation; });
  }
  await test('420px 짧은 6범주 모두 수평 표시',async p=> { const cats=await fixture(p,'column',420,6,{},true); const observation=await checkLabels(p,cats); ok(observation.labels.every(x=>!x.rotation||! /rotate\((?!0(?:\s|\)))/.test(x.rotation)),'짧은 이름은 불필요하게 회전하지 않음'); return observation; });
  await test('명시적 간격 2를 일반·특수 차트가 보존',async p=> { const observations=[]; for(const type of ['column','bar','waterfall','boxWhisker']) { const cats=await fixture(p,type,900,12,{labelInterval:2}); observations.push(await checkLabels(p,cats.filter((_,i)=>i%2===0))); eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelInterval),2,'저장 간격 설정 유지'); } return observations; });
  await test('명시적 간격 1은 모든 항목 표시',async p=> { const cats=await fixture(p,'column',640,8,{labelInterval:1}); return checkLabels(p,cats); });
  for(const type of ['column','bar']) await test('3D '+type+'의 긴 6범주 표시',async p=> { const cats=await fixture(p,type,640,6,{},false,{threeD:true}); return checkLabels(p,cats); });
  for(const angle of [-45,45,90]) await test('명시 회전각 '+angle+'° 유지',async p=> { const cats=await fixture(p,'column',640,6,{labelRotation:angle}); const observation=await checkLabels(p,cats); ok(observation.labels.every(x=>x.rotation?.includes('rotate('+angle)),'모든 표시 레이블 지정 회전각 적용'); return observation; });
  await test('다단계 축의 내부 항목과 상위 그룹 함께 표시',async p=> { const cats=await fixture(p,'column',900,12); await p.evaluate(()=>{const w=tabula.wb(),ch=w.sheets[0].charts[0]; w.setSheetProp(0,'charts',[{...ch,snapshotData:{...ch.snapshotData,catLevels:[[{text:'검색 광고',start:0,end:5},{text:'쇼핑 광고',start:6,end:11}]]}}]); tabula.gv().renderObjectsAll();}); const observation=await checkLabels(p,cats); const groups=await chart(p).locator('[data-el="axis-x"] text:not([data-axis-label])').allTextContents(); eq(groups,['검색 광고','쇼핑 광고'],'상위 항목 그룹 이름 유지'); return observation; });
  await test('120범주 자동 간격은 여러 이름을 겹침 없이 표시',async p=> { const cats=await fixture(p,'column',640,120); const observation=await inspect(p); ok(observation.labels.length>1&&observation.labels.length<120,'밀집 범주에서 필요한 자동 간격만 적용'); eq(observation.overlaps,[],'밀집 범주 텍스트 겹침 없음'); for(const label of observation.labels) { eq(label.full,cats[Number(label.index)],'생략한 항목도 접근 가능한 원문 제공'); ok(label.quad.every(q=>q.x>=-1&&q.y>=-1&&q.x<=observation.width+1&&q.y<=observation.height+1),'밀집 범주 표시가 SVG 안에 있음'); } return observation; });
  await test('축 서식 UI 간격·각도·자동복원 및 Undo',async p=> {
    const cats=await fixture(p,'column',900,12); await chart(p).click({position:{x:20,y:10}});
    await p.evaluate(()=>tabula.run('chartFormat'));
    await p.getByLabel('서식을 지정할 차트 요소',{exact:true}).selectOption('axis-x');
    await p.getByLabel('레이블 간격',{exact:true}).selectOption('manual');
    const interval=p.getByLabel('간격 단위',{exact:true}); await interval.fill('2'); await interval.press('Tab');
    await checkLabels(p,cats.filter((_,i)=>i%2===0));
    eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelInterval),2,'UI 지정 간격이 모델과 렌더링에 연결');
    await p.getByLabel('텍스트 방향',{exact:true}).selectOption('manual');
    const angle=p.getByLabel('사용자 지정 각도(°)',{exact:true}); await angle.fill('-45'); await angle.press('Tab');
    eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelRotation),-45,'UI 각도 설정 저장');
    ok((await inspect(p)).labels.every(n=>n.rotation?.includes('rotate(-45')),'UI 각도가 실제 SVG에 적용');
    await p.getByLabel('텍스트 방향',{exact:true}).selectOption('auto'); await p.getByLabel('레이블 간격',{exact:true}).selectOption('auto');
    eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelInterval),undefined,'자동 간격 복원');
    eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelRotation),undefined,'자동 각도 복원');
    await p.keyboard.press('Escape'); await checkLabels(p,cats);
    await p.evaluate(()=>tabula.run('undo')); eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelInterval),2,'Undo가 수동 간격 복원');
    await p.evaluate(()=>tabula.run('redo')); eq(await p.evaluate(()=>tabula.wb().sheets[0].charts[0].axes.x.labelInterval),undefined,'Redo가 자동 간격 복원');
  });
  await test('긴 범주·하단 범례·축 제목이 서로 겹치지 않음',async p=> { const observations=[]; for(const type of ['column','bar']) { const cats=await fixture(p,type,640,8,{title:'매체 및 기기 구분'},false,{legend:'b'}); const observation=await checkLabels(p,cats); eq(observation.extras.map(x=>x.text).sort(),['광고비','매체 및 기기 구분'].sort(),'하단 범례와 축 제목 표시'); eq(observation.extraOverlaps,[],'긴 범주가 범례나 축 제목을 침범하지 않음'); for(const extra of observation.extras) ok(extra.quad.every(q=>q.x>=-1&&q.y>=-1&&q.x<=observation.width+1&&q.y<=observation.height+1),'범례와 제목이 SVG 안에 있음'); await chart(p).screenshot({path:out+'/axis-legend-title-'+type+'.png'}); observations.push(observation); } return observations; });
  await test('축 숨김은 레이블 없음',async p=> { await fixture(p,'column',640,8,{hide:true}); const observation=await inspect(p); eq(observation.labels.length,0,'숨긴 축 레이블 없음'); return observation; });
  await test('역순 축은 텍스트 손실 없이 배치 반전',async p=> { const cats=await fixture(p,'column',900,12,{reverse:true}); const observation=await checkLabels(p,cats); const centers=observation.labels.map(n=>n.quad.reduce((sum,p)=>sum+p.x,0)/4); ok(centers.every((x,i)=>!i||x<centers[i-1]),'역순 시 화면 위치가 감소'); return observation; });
} finally {
  await browser.close();
  const result={engine,url,assets:[...assets],cases:results.length,passed:results.filter(r=>r.ok).length,checks,results};
  await writeFile(out+'/result.json',JSON.stringify(result,null,2)); console.log(JSON.stringify({engine,cases:result.cases,passed:result.passed,checks,out,assets:result.assets}));
  if(result.passed!==result.cases) process.exitCode=1;
}
