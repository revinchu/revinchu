import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';

// Screenshot-derived synthetic dates only. Never loads or writes a customer document.
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const target = process.env.WIXEL_URL || 'http://127.0.0.1:5186/';
const output = process.env.DATE_FIT_OUTPUT || 'D:/Codex/Temp/wixel-iphone-date/local';
const engines = (process.env.DATE_FIT_BROWSERS || 'chromium,webkit,firefox').split(',');
const failures = [], runs = [], errors = []; let checks = 0;
function check(ok, label, detail) { checks++; if (!ok) failures.push({ label, detail }); }
await mkdir(output, { recursive: true });
for (const engine of engines) {
  const browser = await pw[engine].launch();
  try {
    for (const mobile of [true, false]) {
      const context = await browser.newContext({ viewport: { width: mobile ? 393 : 1280, height: mobile ? 852 : 900 }, deviceScaleFactor: mobile ? 3 : 1,
        ...(engine !== 'firefox' ? { isMobile: mobile } : {}), hasTouch: mobile, serviceWorkers: 'block' });
      await context.route('**/*', route => {
        const req = route.request(), u = new URL(req.url());
        if (u.origin !== new URL(target).origin || req.method() !== 'GET' || u.pathname.startsWith('/api/')) return route.abort();
        return route.continue();
      });
      await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
      const page = await context.newPage(); page.on('pageerror', e => errors.push({ engine, message: e.message }));
      await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await page.waitForFunction(() => !!window.tabula?.gv(), null, { timeout: 60000 });
      await page.evaluate(() => document.fonts.ready);
      const initial = await page.evaluate(() => {
        const t = tabula, wb = t.wb();
        const base = { font: '맑은 고딕', size: 9, align: 'center', valign: 'middle', fill: '#fff', bt: true, bb: true, bl: true, br: true };
        const cells = {}, put = (r,c,raw,style={}) => { cells[`${r},${c}`] = { raw, style: { ...base, ...style } }; };
        put(1,1,'전주 대비'); put(1,2,'기간'); put(1,3,'노출');
        put(2,1,'2023-03-08~2023-03-14'); put(2,2,'7일'); put(2,3,'6098');
        put(3,1,'2023-03-15~2023-03-21', { shrink:true }); put(3,2,'7일');
        put(4,1,'="2023-03-22~2023-03-28"'); put(4,2,'7일');
        put(5,1,'2023-03-29 ~ 2023-04-04', { font:'Missing Date Font' }); put(5,2,'7일');
        put(6,1,'2023-03-08~2023-03-14', { align:'left', indent:2, shrink:true }); put(6,2,'7일');
        put(7,1,'2023-03-08~2023-03-14', { align:'right', indent:1, shrink:true }); put(7,2,'7일');
        put(8,1,'2023-03-08~2023-03-14', { wrap:true }); put(8,2,'7일');
        put(9,1,'2023-03-01', { numFmt:'text' }); put(9,2,'월');
        put(10,1,'45000', { numFmt:'custom',code:'yyyy-mm-dd' }); put(10,2,'수');
        put(11,1,'일반 텍스트는 빈 옆 칸으로 이어짐', { align:'left', bl:false, br:false });
        put(12,1,'긴 일반 텍스트 축소 맞춤', { shrink:true }); put(12,2,'값');
        put(13,1,'2023-03-08~2023-03-14', { shrink:true });
        wb.load({ defaultFont:{name:'맑은 고딕',size:9}, sheets:[{ name:'날짜 범위 재현', defColW:80, defRowH:25,
          colWidths:{0:18,1:116,2:36,3:90}, rowHeights:{8:45}, merges:[{r1:13,c1:1,r2:13,c2:2}],cells }] });
        t.switchSheet(0); t.selectCell(0,0); t.gv().setScroll(0,0); t.gv().renderAll();
        return { data:JSON.stringify(wb.serialize()), version:wb.version };
      });
      for (const frozen of [false,true]) for (const zoom of [85,100,50,150]) {
        await page.evaluate(({zoom,frozen}) => {
          const t=tabula; t.wb().sheets[0].freeze=frozen?{rows:1,cols:1}:{}; t.gv().setZoom(zoom); t.gv().setScroll(0,0); t.gv().renderAll();
        }, {zoom,frozen});
        await page.evaluate(() => new Promise(r => requestAnimationFrame(()=>requestAnimationFrame(r))));
        const info = await page.evaluate(() => {
          const cells = [];
          for (let r=2;r<=13;r++) {
            const n=[...document.querySelectorAll(`.c[data-r="${r}"][data-c="1"]`)].find(n=>n.closest('.pane').getBoundingClientRect().width>0);
            const span=n.querySelector(':scope>span'), a=n.getBoundingClientRect(), b=span.getBoundingClientRect(), css=getComputedStyle(n), z=tabula.gv().z;
            const range=document.createRange(); range.selectNodeContents(span);
            cells.push({r,text:span.textContent,ovf:n.classList.contains('ovf'),x:(b.x-a.x)/z,w:b.width/z,cw:a.width/z,
              padLeft:parseFloat(css.paddingLeft),padRight:parseFloat(css.paddingRight),font:css.fontSize,
              glyphs:[...range.getClientRects()].map(q=>({left:(q.left-a.left)/z,right:(q.right-a.left)/z})),rect:a.toJSON()});
          }
          return {cells,supportsTextAdjust:CSS.supports('-webkit-text-size-adjust','100%')||CSS.supports('text-size-adjust','100%'),textAdjust:getComputedStyle(document.querySelector('.grid-view')).getPropertyValue('-webkit-text-size-adjust')||getComputedStyle(document.querySelector('.grid-view')).getPropertyValue('text-size-adjust')};
        });
        const label=`${engine}-${mobile?'phone':'desktop'}-${frozen?'frozen':'plain'}-${zoom}`;
        for (const cell of info.cells) {
          if(cell.r===11) { check(cell.ovf,label+' ordinary text still spills',cell); continue; }
          if (cell.r!==9) check(!cell.ovf,label+` row${cell.r} confined`,cell);
          check(cell.glyphs.every(q=>q.left>=cell.padLeft-.7&&q.right<=cell.cw-cell.padRight+.7),label+` row${cell.r} full text fits padding`,cell);
          if([2,3,4,5,12,13].includes(cell.r)) check(Math.abs(cell.x+cell.w/2-cell.cw/2)<.7,label+` row${cell.r} centered`,cell);
        }
        if (info.supportsTextAdjust) check(info.textAdjust==='100%',label+' fixed sheet font scale',info.textAdjust);
        if (!frozen && zoom===85) await page.screenshot({path:join(output,label+'.png')});
        runs.push({label,...info});
      }
      const after=await page.evaluate(()=>{tabula.wb().sheets[0].freeze={};return {data:JSON.stringify(tabula.wb().serialize()),version:tabula.wb().version};});
      // The fixture explicitly toggles frozen panes; normalize that empty property only.
      const a=JSON.parse(initial.data), b=JSON.parse(after.data); delete a.sheets[0].freeze; delete b.sheets[0].freeze;
      check(JSON.stringify(a)===JSON.stringify(b),engine+' values/styles unchanged');
      check(initial.version===after.version,engine+' no edit history from rendering');
      await context.close();
    }
  } finally { await browser.close(); }
}
check(errors.length===0,'no page errors',errors);
const report={target,checks,cases:runs.length,bad:failures.length,failures,errors,runs};
await writeFile(join(output,'result.json'),JSON.stringify(report,null,2));
console.log(JSON.stringify({...report,runs:undefined,failures:failures.slice(0,8)},null,2));
assert.equal(failures.length,0,'date cell fit; see result.json');
