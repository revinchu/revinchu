// 이미 로컬 격리 브라우저에 읽어 둔 실파일만 검사한다. 원본 저장/내보내기/원격 전송 없음.
import { readFile, writeFile } from 'node:fs/promises';
import assert from 'node:assert/strict';
import { auditNativeSamples } from './real-file-native-samples.mjs';
const settle = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function state(p) {
  return p.evaluate(() => {
    const w=tabula.wb();
    return {undo:w.undoStack.length, active:tabula.si, sheets:w.sheets.map((s,si)=>({
      name:s.name, pivots:[s.pivot,...(s.pivotsExtra||[])].filter(Boolean).map(d=>({name:d.name,filters:d.filters||{},fieldFilters:d.fieldFilters||{},pages:d.pages||[], area:d.area,
        sample:d.area?Array.from({length:Math.min(5,d.area.r2-d.area.r1+1)},(_,i)=>[w.getRaw(si,d.area.r1+i,d.area.c1),w.getRaw(si,d.area.r1+i,Math.min(d.area.c2,d.area.c1+1))]):[]})),
      tables:(s.tables||[]).map(t=>({name:t.name,criteria:t.filter?.criteria||{}}))
    }))};
  });
}
function filters(s) {return s.sheets.map(x=>({name:x.name,pivots:x.pivots.map(p=>({name:p.name,filters:p.filters,fieldFilters:p.fieldFilters})),tables:x.tables}));}
export async function auditObjects(page,{id,out,metadata,nativeSamples}) {
  const report={id,scope:'실파일의 첫 표시 슬라이서와 첫 표시 피벗 필터를 실제 클릭하고 실행 취소. 셀 값은 결과에 기록하지 않음.',results:[]};
  let native;
  try { native=JSON.parse((await readFile('D:/Codex/Temp/wixel-final-audit/native/'+id+'.json','utf8')).replace(/^\uFEFF/,'')); } catch {}
  if(nativeSamples) Object.assign(report,nativeSamples);
  else if(native?.opened) Object.assign(report,await auditNativeSamples(page,native));
  const original=await state(page), originalTimeout=Math.max(15000,Number(process.env.WIXEL_AUDIT_CLICK_TIMEOUT)||15000);
  const undoTo=async n=>{for(let i=0;i<10;i++){const current=await page.evaluate(()=>tabula.wb().undoStack.length);if(current<=n)break;await page.evaluate(()=>tabula.run('undo'));await settle(page);}};
  async function test(name,fn){const b=await state(page),started=Date.now();try{const details=await fn(b);const ms=Date.now()-started;report.results.push({name,ok:true,ms,...(ms>15000?{slow:true,performanceFailure:true}:{}),...details});}catch(e){report.results.push({name,ok:false,ms:Date.now()-started,error:e.message});await page.screenshot({path:out+'/'+id+'-objects-failure-'+report.results.length+'.png',timeout:10000}).catch(()=>{});}finally{await page.keyboard.press('Escape').catch(()=>{});await undoTo(b.undo);assert.deepEqual(filters(await state(page)),filters(b),'상호작용 실행 취소 후 원본 필터 복원');}}
  try {
    const candidate=metadata.sheets.find(s=>s.slicers&&s.state!=='hidden'&&s.state!=='veryHidden'&&!s.protected);
    if(candidate){
      await page.evaluate(i=>{tabula.switchSheet(i);tabula.gv().setScroll(0,0);tabula.gv().renderAll();},candidate.index);await settle(page);
      const sl=await page.evaluate(()=>{const hit=e=>{if(!e)return false;const r=e.getBoundingClientRect(),x=r.x+r.width/2,y=r.y+r.height/2;return r.width>0&&r.height>0&&x>=0&&y>=0&&x<innerWidth&&y<innerHeight&&e.contains(document.elementFromPoint(x,y));};return [...document.querySelectorAll('.obj.slicer')].map(e=>({id:e.dataset.id,keys:[...e.querySelectorAll('.sl-item:not(.nodata)')].filter(hit).map(b=>b.dataset.k),clear:hit(e.querySelector('.sl-clear'))})).find(e=>e.keys.length>=2&&e.clear);});
      if(sl){
        const slicerSelector='.obj.slicer[data-id='+JSON.stringify(sl.id)+']:visible';
        const base=()=>page.locator(slicerSelector).first();
        const target=async part=>{
          const matches=page.locator(slicerSelector+' '+part);
          const i=await matches.evaluateAll(ns=>ns.findIndex(e=>{const r=e.getBoundingClientRect(),x=r.x+r.width/2,y=r.y+r.height/2;return r.width>0&&r.height>0&&x>=0&&y>=0&&x<innerWidth&&y<innerHeight&&e.contains(document.elementFromPoint(x,y));}));
          assert.ok(i>=0,'고정 창 복제본 중 실제 클릭 가능한 슬라이서 단추');
          return matches.nth(i);
        };
        await test('슬라이서 우클릭에는 셀 서식 메뉴가 없음',async()=>{
          await (await target('.sl-item[data-k='+JSON.stringify(sl.keys[0])+']')).click({button:'right',timeout:originalTimeout});
          await page.getByRole('menuitem',{name:/^슬라이서 설정/}).waitFor({timeout:originalTimeout});
          assert.equal(await page.getByRole('menuitem',{name:/^셀 서식/}).count(),0);
          assert.equal(await page.locator('#menuLayer .context-mini-toolbar').count(),0);
          return {sheet:candidate.index};
        });
        await test('슬라이서 단일 선택·Ctrl 추가 선택·해제·실행 취소',async b=>{
          const options=(await base().locator('.sl-item:not(.nodata)').evaluateAll(ns=>ns.map(n=>({key:n.dataset.k,on:n.classList.contains('on')})))).filter(x=>sl.keys.includes(x.key));
          const selectedCount=options.filter(x=>x.on).length;
          const start=(selectedCount===1?options.find(x=>!x.on):options[0])||options[0];
          const keys=[start.key,options.find(x=>x.key!==start.key).key];
          await (await target('.sl-item[data-k='+JSON.stringify(keys[0])+']')).click({timeout:originalTimeout});await settle(page);
          const one=await state(page);
          assert.notDeepEqual(filters(one),filters(b),'첫 항목 선택이 연결 피벗/표 필터에 적용');
          await (await target('.sl-item[data-k='+JSON.stringify(keys[1])+']')).click({modifiers:['Control'],timeout:originalTimeout});await settle(page);
          const selected=await base().locator('.sl-item.on').evaluateAll(ns=>ns.map(n=>n.dataset.k));
          assert.ok(keys.every(k=>selected.includes(k)),'Ctrl 클릭이 두 항목을 함께 선택');
          await (await target('.sl-clear')).click({timeout:originalTimeout});await settle(page);
          assert.ok(await base().locator('.sl-clear.off').count(),'필터 해제 후 전체 선택 상태');
          return {sheet:candidate.index,linkedFilterChanged:true,multiSelection:true,clear:true};
        });
      }else report.results.push({name:'슬라이서',skipped:'표시 영역에서 데이터 있는 항목 두 개를 가진 슬라이서를 찾지 못함'});
    }else report.results.push({name:'슬라이서',skipped:'편집 가능한 표시 슬라이서 없음'});
    const pivotSheet=metadata.sheets.find(s=>s.pivots&&s.state!=='hidden'&&s.state!=='veryHidden'&&!s.protected);
    if(pivotSheet){
      await page.evaluate(i=>{tabula.switchSheet(i);tabula.gv().setScroll(0,0);tabula.gv().renderAll();},pivotSheet.index);await settle(page);
      const btn=page.locator('.pbtn[data-k="rows"]:visible').first();
      if(await btn.count())await test('피벗 행 필터 검색·현재 선택 추가·적용·실행 취소',async b=>{
        await btn.click({timeout:originalTimeout});const menu=page.locator('.pivot-filter-menu').last();await menu.waitFor({timeout:originalTimeout});
        const items=menu.locator('[data-filter-index]');
        if(!await items.count()){await menu.getByRole('button',{name:'취소',exact:true}).click();return {skipped:'표시 필터 항목 없음'};}
        const text=(await items.first().evaluate(e=>(e.closest('label')?.textContent||e.getAttribute('aria-label')||e.textContent||'').trim())).trim();
        assert.ok(text,'필터 항목 텍스트');
        await menu.getByRole('searchbox').fill(text.replace(/[~*?]/g,m=>'~'+m));
        assert.equal(await menu.getByRole('checkbox',{name:'필터에 현재 선택한 내용 추가',exact:true}).isVisible(),true,'검색 후 현재 선택 추가 표시');
        await menu.getByRole('button',{name:'확인',exact:true}).click();await settle(page);
        assert.notDeepEqual(filters(await state(page)),filters(b),'검색 결과 필터 적용');
        return {sheet:pivotSheet.index,search:true,addCurrentSelectionShown:true,applied:true};
      });else report.results.push({name:'피벗 행 필터',skipped:'현재 화면에 행 필터 버튼이 없음'});
    }else report.results.push({name:'피벗 필터',skipped:'편집 가능한 표시 피벗 없음'});
  }catch(e){report.fatal=e.message;}
  finally{await page.keyboard.press('Escape').catch(()=>{});await undoTo(original.undo).catch(e=>report.restoreError=e.message);await page.evaluate(i=>tabula.switchSheet(i),original.active).catch(()=>{});await writeFile(out+'/'+id+'-objects.json',JSON.stringify(report,null,2));}
  report.ok=!report.fatal&&!report.restoreError&&report.results.every(r=>r.ok||r.skipped);
  return report;
}
