// 실제 업무 파일 없이 대량 스타일 갤러리의 DOM 상한·검색·적용을 검사한다.
import assert from 'node:assert/strict';
const { chromium, webkit } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
if (!['localhost','127.0.0.1'].includes(new URL(url).hostname)) throw Error('로컬 서버만 사용');
for (const [engine, type] of [['chromium',chromium],['webkit',webkit]]) {
 const browser=await type.launch();const page=await browser.newPage({viewport:{width:1100,height:780}});const errors=[];
 page.on('pageerror',e=>errors.push(e.message));
 await page.route('**/api/**',r=>r.abort());
 await page.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 try {
  await page.goto(url);await page.waitForFunction(()=>window.tabula?.wb());
  await page.evaluate(()=>{const w=tabula.wb();w.cellStyles=Array.from({length:50000},(_,i)=>({name:'Audit '+String(i).padStart(5,'0'),style:{fill:'#abcdef',italic:!!(i%2)}}));w.setInput(0,0,0,'42');tabula.selectCell(0,0);tabula.openNamedMenu('cellStyles',{x:100,y:100});});
  assert.equal(await page.locator('[data-cell-style]').count(),144);
  assert.equal(await page.locator('[data-cell-style="Audit 00000"]').count(),1);
  await page.getByRole('button',{name:'다음 셀 스타일 페이지',exact:true}).click();
  assert.equal(await page.locator('[data-cell-style]').count(),144);assert.equal(await page.locator('[data-cell-style="Audit 00144"]').count(),1);
  await page.getByRole('button',{name:'이전 셀 스타일 페이지',exact:true}).click();
  assert.equal(await page.locator('[data-cell-style="Audit 00000"]').count(),1);
  const search=page.getByRole('searchbox',{name:'셀 스타일 검색'});await search.fill('Audit 49999');
  assert.equal(await page.locator('[data-cell-style]').count(),1);await search.press('Enter');
  const value=await page.evaluate(()=>({v:tabula.wb().getValue(0,0,0),s:tabula.wb().styleAt(0,0,0)}));assert.equal(value.v,42);assert.equal(value.s.fill,'#abcdef');assert.equal(value.s.italic,true);
  await page.evaluate(()=>tabula.run('undo'));assert.equal(await page.evaluate(()=>tabula.wb().styleAt(0,0,0).fill),undefined);
  await page.evaluate(()=>tabula.openNamedMenu('cellStyles',{x:100,y:100}));await search.fill('No matches anywhere');assert.equal(await page.locator('[data-cell-style]').count(),0);
  assert.equal(await page.evaluate(()=>tabula.wb().cellStyles.length),50000);assert.deepEqual(errors,[]);
  console.log(JSON.stringify({engine,checks:13,ok:true,pageErrors:errors}));
 } finally {await browser.close();}
}
