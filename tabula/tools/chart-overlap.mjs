// 합성 차트의 회전 텍스트 사각형과 실제 글자 픽셀을 함께 검사한다. 업무 파일/원격 쓰기는 사용하지 않는다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5184/';
if (!['127.0.0.1', 'localhost', 'wixel-3.wizx.workers.dev'].includes(new URL(url).hostname)) throw new Error('검증용 주소를 확인하세요.');
const out = process.env.CHART_OVERLAP_OUT || 'D:/Codex/Temp/wixel-chart-overlap/ui';
await mkdir(out, { recursive: true });
const engines = (process.env.CHART_OVERLAP_ENGINES || 'chromium,firefox,webkit').split(',');
let cases = 0, checks = 0; const errors = [], writes = [];
const fixtures = [
  { name: 'manual', w: 1407, h: 223, title: 'Daily Imps & CTR', titleLayout: { x: .103, y: .022 },
    legend: 'r', legendLayout: { x: .6713, y: .0314, w: .2722, h: .0897 }, legendOverlay: false,
    plotLayout: { x: .0605, y: .157, w: .8702, h: .6457, target: 'inner' },
    series: [{}, {}, {}], snapshotData: { categories: Array(14).fill(''), series: ['Imps', 'Clicks', 'CTR'].map((name, i) => ({ name, type: i ? 'line' : 'column', axis: i ? 1 : 0, values: Array(14).fill(null) })) } },
  { name: 'collision', w: 669, h: 290, legend: 'b', legendLayout: { x: .3912, y: .8923, w: .2212, h: .0757 }, legendOverlay: false,
    axes: { y: { min: 0, max: 600000, major: 100000, numFmt: '#,##0' }, y2: { min: 0, max: .045, major: .005, numFmt: '0.00%' } },
    snapshotData: { categories: ['메인 배너'], series: [{ name: '실제노출량', values: [570169], type: 'column', labels: true, numFmt: '#,##0' }, { name: 'CTR', values: [.0414], type: 'line', axis: 1, labels: true, marker: 'none', lineNone: true, labelPos: 'above', numFmt: '0.00%' }] } },
];
for (const engine of engines) {
  const browser = await pw[engine].launch();
  try { for (const dpr of [1, 1.25, 2]) {
    const p = await browser.newPage({ viewport: { width: 2100, height: 1100 }, deviceScaleFactor: dpr });
    p.on('pageerror', e => errors.push(e.message));
    await p.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.route('**/*', r => {
      if (!['GET', 'HEAD'].includes(r.request().method())) { writes.push(r.request().method()); return r.abort(); }
      const u = new URL(r.request().url()); return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? r.continue() : r.abort();
    });
    await p.goto(url); await p.waitForFunction(() => window.tabula?.gv());
    for (const fixture of fixtures) {
      await p.evaluate(f => {
        const chart = { ...f, id: 'overlap', type: 'combo', x: 20, y: 20, gridY: false };
        tabula.wb().restore({ sheets: [{ name: '첫 배율', cells: {}, zoom: 85, charts: [chart] }, { name: '검사', cells: {}, zoom: 80, freeze: { rows: 20, cols: 8 }, charts: [chart] }] });
        tabula.switchSheet(1); tabula.switchSheet(0);
      }, fixture);
      await p.screenshot(); await p.evaluate(() => tabula.switchSheet(1));
      for (const zoom of [80, 100, 125, 85, 80]) {
        await p.evaluate(z => { tabula.gv().setZoom(z); tabula.gv().setScroll(0, 0); }, zoom);
        await p.evaluate(() => document.fonts.ready); await p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
        const observation = await p.locator('.chart[data-id="overlap"] svg').first().evaluate(svg => {
          const sr = svg.getBoundingClientRect();
          const labels = [...svg.querySelectorAll('text')].filter(n => n.textContent.trim()).map(n => {
            const b = n.getBBox(), m = n.getScreenCTM(), r = n.getBoundingClientRect();
            return { text: n.textContent, kind: n.closest('[data-el]')?.getAttribute('data-el'), rect: { x: r.x, y: r.y, w: r.width, h: r.height },
              quad: [[b.x,b.y],[b.x+b.width,b.y],[b.x+b.width,b.y+b.height],[b.x,b.y+b.height]].map(([x,y]) => ({ x:m.a*x+m.c*y+m.e, y:m.b*x+m.d*y+m.f })) };
          });
          function overlap(a,b) {
            for (const poly of [a,b]) for (let i=0;i<poly.length;i++) {
              const p=poly[i],q=poly[(i+1)%poly.length],len=Math.hypot(q.x-p.x,q.y-p.y); if(!len)return false;
              const ax=-(q.y-p.y)/len,ay=(q.x-p.x)/len,pa=a.map(p=>p.x*ax+p.y*ay),pb=b.map(p=>p.x*ax+p.y*ay);
              if(Math.min(Math.max(...pa),Math.max(...pb))-Math.max(Math.min(...pa),Math.min(...pb))<=.75)return false;
            } return true;
          }
          const overlaps=[];
          for(let i=0;i<labels.length;i++) for(let j=i+1;j<labels.length;j++) if(overlap(labels[i].quad,labels[j].quad))overlaps.push([labels[i].text,labels[j].text]);
          return { labels, overlaps, outside:labels.filter(l=>l.quad.some(p=>p.x<sr.x-1||p.y<sr.y-1||p.x>sr.right+1||p.y>sr.bottom+1)).map(l=>l.text) };
        });
        const name = `${engine}-${fixture.name}-dpr${dpr}-zoom${zoom}-${cases}`, png = await p.screenshot();
        if (dpr === 1 || observation.overlaps.length || observation.outside.length) await writeFile(`${out}/${name}.png`, png);
        assert.deepEqual(observation.overlaps, [], name + ' 텍스트 겹침'); assert.deepEqual(observation.outside, [], name + ' 글자 잘림'); checks += 2;
        assert.equal(observation.labels.filter(n => n.kind === 'legend').length, fixture.name === 'manual' ? 3 : 1); checks++;
        const counts = await p.evaluate(async ({ base64, labels, dpr }) => {
          const img = new Image(); img.src='data:image/png;base64,'+base64;await img.decode();
          const c=document.createElement('canvas');c.width=img.width;c.height=img.height;const ctx=c.getContext('2d');ctx.drawImage(img,0,0);const data=ctx.getImageData(0,0,c.width,c.height).data;
          return labels.map(({rect:r})=>{let ink=0;for(let y=Math.max(0,Math.floor(r.y*dpr));y<Math.min(c.height,Math.ceil((r.y+r.h)*dpr));y++)for(let x=Math.max(0,Math.floor(r.x*dpr));x<Math.min(c.width,Math.ceil((r.x+r.w)*dpr));x++){const k=(y*c.width+x)*4;if(Math.max(data[k],data[k+1],data[k+2])<200)ink++;}return ink;});
        }, { base64: png.toString('base64'), labels: observation.labels, dpr });
        await writeFile(`${out}/${name}.json`,JSON.stringify({observation,counts},null,2));
        counts.forEach((n,i) => { assert.ok(n>=2,name+' '+observation.labels[i].text+' 실제 글자 픽셀 ('+n+')');checks++; }); cases++;
      }
    }
    await p.close();
  }} finally { await browser.close(); }
}
assert.deepEqual(errors, []);assert.deepEqual(writes, []);
const result = { engines, cases, checks, errors, writes }; await writeFile(`${out}/result.json`,JSON.stringify(result,null,2)); console.log(JSON.stringify(result));
