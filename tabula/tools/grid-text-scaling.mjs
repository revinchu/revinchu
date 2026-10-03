import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { dirname, resolve, join } from 'node:path';

// 합성 셀만 사용한다. 실제 파일·원격 보관함·시스템 클립보드에 접근하지 않는다.
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const target = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const output = process.env.GRID_TEXT_OUTPUT || 'D:/Codex/Temp/wixel-grid-text-scaling';
const baselineRef = process.env.GRID_TEXT_BASELINE_REF;
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const oldView = baselineRef ? execFileSync('git', ['-c', `safe.directory=${root}`, 'show', `${baselineRef}:tabula/src/view.js`], { cwd: root, encoding: 'utf8', maxBuffer: 4 * 1024 * 1024 }) : null;
await mkdir(output, { recursive: true });
const errors = [], writes = [], failures = [], runs = [];
let checks = 0;
function eq(a, b, label) { checks++; try { assert.deepEqual(a, b); } catch { if (failures.length < 80) failures.push({ label, actual: a, expected: b }); } }
function near(a, b, label, tolerance = .06) { eq(Math.abs(a - b) <= tolerance, true, `${label} (${a}, ${b})`); }
const browser = await chromium.launch();
try {
  for (const cfg of [{ dpr: 1, mobile: false }, { dpr: 1.5, mobile: false }, { dpr: 2, mobile: false }, { dpr: 2, mobile: true }]) {
    const context = await browser.newContext({ viewport: { width: cfg.mobile ? 390 : 1800, height: 1200 }, deviceScaleFactor: cfg.dpr, isMobile: cfg.mobile, hasTouch: cfg.mobile });
    await context.route('**/*', route => {
      const r = route.request(), u = new URL(r.url());
      if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); }
      if (u.origin !== new URL(target).origin || u.pathname.startsWith('/api/')) return route.abort();
      if (oldView && u.pathname === '/src/view.js') return route.fulfill({ contentType: 'text/javascript', body: oldView });
      return route.continue();
    });
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    const page = await context.newPage(); page.on('pageerror', e => errors.push(e.message));
    await page.goto(target, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await page.waitForFunction(() => !!window.tabula?.gv());
    for (const frozen of [false, true]) {
      const before = await page.evaluate(frozen => {
        const t = window.tabula, w = t.wb(), g = t.gv();
        const dot = { font: '돋움', size: 9, align: 'center', valign: 'middle' };
        const malgun = { font: '맑은 고딕', size: 11, align: 'center', valign: 'middle' };
        w.restore({ defaultFont: { name: '맑은 고딕', size: 11 }, sheets: [{ name: '배율 합성', zoom: 100, defRowH: 24, defColW: 150,
          rowHeights: { 2: 70, 3: 45, 8: 95 }, freeze: frozen ? { rows: 1, cols: 1 } : {}, merges: [{ r1: 3, c1: 2, r2: 3, c2: 3 }],
          cells: {
            '0,0': { raw: '정산내역 ABC123', style: dot }, '1,0': { raw: '정산내역 ABC123', style: { ...dot, wrap: true } },
            '0,1': { raw: '성과 보고서 마케팅부 월별 정산 금액 (원)', style: malgun }, '1,1': { raw: '성과 보고서 마케팅부 월별 정산 금액 (원)', style: { ...malgun, wrap: true } },
            '2,0': { raw: '첫째 줄\n둘째 줄\n셋째 줄', style: { ...malgun, wrap: true, align: 'left' } },
            '3,2': { raw: '여러 칸을 병합한 보고서 제목 ABC123 줄바꿈 검사', style: { ...malgun, wrap: true } },
            '4,0': { raw: '칸 너비에 맞춰 축소할 긴 문자열 1234567890', style: { ...malgun, shrink: true } },
            '5,0': { raw: '1234567890.12', style: { ...malgun, numFmt: 'custom', code: '#,##0.00', shrink: true, align: 'right' } },
            '6,1': { raw: '0.1234', style: { ...malgun, numFmt: 'custom', code: '0.00%', align: 'right' } },
            '7,0': { raw: '빈 셀 쪽으로 이어지는 긴 문자열 ABC123456789', style: { ...dot, align: 'left' } },
            '8,0': { raw: '회전 텍스트', style: { ...malgun, rotate: 45 } },
            '9,1': { raw: '들여쓰기', style: { ...malgun, align: 'left', indent: 2 } },
            '10,0': { raw: '편집 취소 보존', style: dot }
          } }] });
        t.switchSheet(0); g.setZoom(100); g.setScroll(0, 0); g.renderAll();
        // 각 런타임의 실제 글꼴 폭에 맞춘 경계 셀. 100%의 한 줄이 축소 후 두 줄이면 회귀다.
        const width = c => document.querySelector(`.c[data-r="0"][data-c="${c}"] > span`).getBoundingClientRect().width;
        w.setColWidth(0, 0, Math.ceil(width(0)) + 6);
        w.setColWidth(0, 1, Math.ceil(width(1)) + 6);
        g.layout(); g.renderAll(); t.selectCell(10, 0); g.setScroll(0, 0); g.renderAll();
        w.undoStack = []; w.redoStack = [];
        return { data: JSON.stringify(w.serialize()), version: w.version, widths: [w.colWidth(0, 0), w.colWidth(0, 1)] };
      }, frozen);
      let base;
      for (const zoom of [100, 25, 30, 35, 40, 45, 50, 55, 100, 200]) {
        await page.evaluate(zoom => { const g = tabula.gv(); g.setZoom(zoom); g.setScroll(0, 0); }, zoom);
        await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
        const info = await page.evaluate(() => {
          const g = tabula.gv(), z = g.z, rect = n => n.getBoundingClientRect();
          const cells = [[0,0],[1,0],[0,1],[1,1],[2,0],[3,2],[4,0],[5,0],[6,1],[7,0],[8,0],[9,1]].map(([r,c]) => {
            const node = [...document.querySelectorAll(`.c[data-r="${r}"][data-c="${c}"]`)].find(n => rect(n).width);
            const span = node.querySelector(':scope > span'), a = rect(node), b = rect(span), css = getComputedStyle(node);
            return { r,c, x:(b.x-a.x)/z,y:(b.y-a.y)/z,w:b.width/z,h:b.height/z,cw:a.width/z,ch:a.height/z, font:css.fontSize, line:css.lineHeight,pad:css.padding,text:span.textContent,overflow:node.classList.contains('ovf') };
          });
          const p = g.clientRect({ r1:2,c1:0,r2:2,c2:0 });
          return { zoom:g.z*100, cells, hit:g.hitTest(p.left+p.width/2,p.top+p.height/2), geometry:{width:g.viewEl.getBoundingClientRect().width,height:g.viewEl.getBoundingClientRect().height,w:g.scroll.clientWidth,h:g.scroll.clientHeight} };
        });
        const name = `dpr${cfg.dpr}-${cfg.mobile?'mobile':'desktop'}-${frozen?'frozen':'plain'}-${zoom}`;
        near(info.zoom, zoom, name+' zoom');
        near(info.geometry.width, info.geometry.w, name+' viewport width'); near(info.geometry.height, info.geometry.h, name+' viewport height');
        eq({r:info.hit.r,c:info.hit.c}, {r:2,c:0}, name+' hit mapping');
        if (!base) { base = info.cells; for (const i of [1,3]) near(base[i].h,base[i-1].h,name+' baseline one-line '+i); }
        for (let i=0;i<info.cells.length;i++) {
          const a=info.cells[i],b=base[i];
          for(const k of ['x','y','w','h','cw','ch']) near(a[k],b[k],name+` cell${a.r},${a.c} ${k}`);
          for(const k of ['font','line','pad','text','overflow']) eq(a[k],b[k],name+` cell${a.r},${a.c} ${k}`);
        }
        // 선택·편집 취소도 같은 좌표계를 써야 한다. 작은 모바일에서도 A3은 화면 안이다.
        const point = await page.evaluate(()=>{const q=tabula.gv().clientRect({r1:2,c1:0,r2:2,c2:0});return{x:q.left+q.width/2,y:q.top+q.height/2};});
        await page.mouse.click(point.x,point.y);
        eq(await page.evaluate(()=>({...tabula.active})),{r:2,c:0},name+' mouse selection');
        await page.locator('#cellEditor').focus(); await page.keyboard.press('F2');
        const editor = await page.evaluate(()=>{const g=tabula.gv(),e=document.querySelector('#cellEditor').getBoundingClientRect(),a=g.clientRect({r1:2,c1:0,r2:2,c2:0});return{editing:g.host.state().editing,dx:e.x-a.left,dy:e.y-a.top};});
        eq(editor.editing,true,name+' F2');near(editor.dx,0,name+' editor x',.2);near(editor.dy,0,name+' editor y',.2);
        await page.keyboard.press('Escape');
        eq(await page.evaluate(()=>tabula.wb().undoStack.length),0,name+' cancel undo unchanged');
        if(cfg.dpr===1&&!cfg.mobile&&!frozen&&[25,45,100,200].includes(zoom))await page.screenshot({path:join(output,`text-${zoom}.png`)});
        runs.push({ ...cfg, frozen, zoom, wrapHeight:info.cells[1].h, mergeHeight:info.cells[5].h });
      }
      eq(await page.evaluate(()=>JSON.stringify(tabula.wb().serialize())),before.data,'workbook values/styles/size/saved zoom unchanged');
      eq(await page.evaluate(()=>tabula.wb().version),before.version,'workbook version unchanged');
    }
    await context.close();
  }
  eq(errors,[],'page errors');eq(writes,[],'remote writes');
  const report={ok:failures.length===0,cases:runs.length,checks,failures,pageErrors:errors,blockedWrites:writes,baselineRef:baselineRef||null,runs};
  await writeFile(join(output,'result.json'),JSON.stringify(report,null,2));console.log(JSON.stringify(report));
  assert.equal(failures.length,0,'저배율 셀 배치 회귀; result.json 확인');
}finally{await browser.close();}
