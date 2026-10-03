// Local synthetic documents only. Chromium GC diagnoses retained JS objects; it is not an iPad OS test.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5191/');
if (!['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)) throw Error('Local source or compiled server required.');
const out = process.env.WIXEL_PIVOT_RELEASE_OUT || 'D:/Codex/Temp/wixel-pivot-document-release';
await mkdir(out, { recursive:true });
const browser = await chromium.launch({ headless:true }), results = [];
let checks = 0;
const eq = (a,b,message) => { checks++; assert.deepEqual(a,b,message); };
function fixture(name) {
  const cells = {};
  [['Region','Sales'],['Seoul',100],['Busan',200]].forEach((row,r) => row.forEach((value,c) => { cells[r+','+c] = {raw:String(value)}; }));
  return { name:name+'.wixel', mimeType:'application/json', buffer:Buffer.from(JSON.stringify({docName:name,workbook:{sheets:[{name,cells}]}})) };
}
try {
  for (const hidden of [false,true]) {
    const name = hidden ? 'hidden-pane-release' : 'visible-pane-release';
    const context = await browser.newContext({viewport:{width:1440,height:1000},serviceWorkers:'block'}), page = await context.newPage(), errors = [], writes = [];
    page.setDefaultTimeout(15000);
    page.on('pageerror', e => errors.push(e.message));
    page.on('dialog', d => d.dismiss());
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await context.route('**/*', route => {
      const request=route.request(), target=new URL(request.url());
      if (!['GET','HEAD'].includes(request.method())) { writes.push(request.method()+' '+target.pathname); return route.abort(); }
      if (target.origin!==url.origin || target.pathname.startsWith('/api/')) return route.abort();
      return route.continue();
    });
    try {
      await page.goto(url.href,{waitUntil:'domcontentloaded'});
      await page.waitForFunction(() => !!window.tabula?.wb());
      await page.locator('#fileInput').setInputFiles(fixture('Source A'));
      await page.waitForFunction(() => window.tabula?.wb().sheets[0].name==='Source A');
      await page.evaluate(() => { window.tabula.selectRange({r1:0,c1:0,r2:2,c2:1}); window.tabula.run('insertPivot'); });
      await page.getByRole('dialog',{name:'피벗 테이블 만들기',exact:true}).getByRole('button',{name:'확인',exact:true}).click();
      await page.locator('#pivotPane .pp-fields').waitFor();
      await page.locator('#pivotPane [data-pivot-field="Region"]').check();
      await page.locator('#pivotPane [data-pivot-field="Sales"]').check();
      await page.evaluate(() => {
        window.tabula.wb().__releaseAudit = 'previous-pivot-workbook';
        window.__oldBookWeak = new WeakRef(window.tabula.wb());
        window.__oldPaneWeak = new WeakRef(document.getElementById('pivotPane'));
      });
      if (hidden) {
        await page.locator('#pivotPane .pp-head button').click();
        eq(await page.locator('#pivotPane').isVisible(),false,'Pane hidden in the same workbook');
        eq(await page.evaluate(() => window.__oldPaneWeak.deref()===document.getElementById('pivotPane')),true,'Same-document hide preserves pane');
      }
      await page.locator('#fileInput').setInputFiles(fixture('Destination B'));
      await page.waitForFunction(() => window.tabula?.wb().sheets[0].name==='Destination B');
      eq(await page.locator('#dialogLayer .dialog').allTextContents(),[],'Fewer-sheet replacement must complete without an import error');
      eq(await page.locator('#pivotPane .pp-fields').count(),0,'Old field controls removed at document replacement');
      eq(await page.locator('.pivot-field-ghost,.pivot-field-marker').count(),0,'No old drag overlay');
      eq(await page.evaluate(() => document.getElementById('gridWrap').classList.contains('with-pane')),false,'New grid has no stale sidebar width');
      // Document open schedules finite graph preparation at 1200 ms. Allow it to finish before GC.
      await page.waitForTimeout(1600);
      const cdp=await context.newCDPSession(page);
      await cdp.send('HeapProfiler.collectGarbage');
      await page.evaluate(() => new Promise(resolve => setTimeout(resolve,0)));
      await cdp.send('HeapProfiler.collectGarbage');
      const retained=await page.evaluate(() => ({book:!!window.__oldBookWeak.deref(),pane:!!window.__oldPaneWeak.deref()}));
      if (retained.book && process.env.WIXEL_PIVOT_RELEASE_HEAP === '1') {
        const chunks=[];
        cdp.on('HeapProfiler.addHeapSnapshotChunk',event=>chunks.push(event.chunk));
        await cdp.send('HeapProfiler.takeHeapSnapshot',{reportProgress:false});
        await writeFile(out+'/'+name+'.heapsnapshot',chunks.join(''));
      }
      eq(retained,{book:false,pane:false},'Previous workbook and its field pane are collectible');
      await cdp.detach();
      await page.evaluate(() => { const t=window.tabula,w=t.wb(); w.transact(() => w.setInput(0,1,1,'250')); t.selectCell(1,1); });
      eq(await page.evaluate(() => window.tabula.wb().getRaw(0,1,1)),'250','New workbook remains editable');
      eq(errors,[],'No application errors'); eq(writes,[],'No remote writes');
      results.push({name,ok:true,retained});
    } catch (error) { results.push({name,ok:false,error:error.stack}); }
    finally { await context.close(); }
  }
} finally { await browser.close(); }
await writeFile(out+'/result.json',JSON.stringify({checks,results},null,2));
console.log(JSON.stringify({checks,results},null,2));
if (results.some(result=>!result.ok)) process.exitCode=1;
