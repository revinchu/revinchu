// 차트 요소 직접 선택·이동·삭제·서식: 합성 문서와 격리 브라우저만 사용한다.
// WIXEL_URL(소스/번들/공개), WIXEL_CHART_DIRECT_FILTER, WIXEL_CHART_DIRECT_SCREENSHOT 선택 가능.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
const out = process.env.CHART_PARITY_OUT || 'D:/Codex/Temp/wixel-chart-parity';
await mkdir(out, { recursive: true });
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const browser = await chromium.launch(), results = [], filter = process.env.CHART_PARITY_FILTER;
async function test(name, fn) {
  if (filter && !name.includes(filter)) return;
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 }, acceptDownloads: false });
  const p = await context.newPage(), errors = [], writes = []; p.setDefaultTimeout(10000);
  p.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', route => {
    const r = route.request(), target = new URL(r.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); }
    return target.origin === new URL(url).origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
    await fixture(p); await fn(p); await cellsUnchanged(p);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors, blockedWrites: writes }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
const current = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets[0].charts[0]));
const state = p => p.evaluate(() => structuredClone(window.tabula.gv().host.state().chartPart));
const depth = p => p.evaluate(() => window.tabula.wb().undoStack.length);
const command = (p, cmd) => p.evaluate(cmd => window.tabula.run(cmd), cmd);
async function fixture(p, patch = {}, zoom = 100) {
  await p.evaluate(({ patch, zoom }) => {
    const t = window.tabula, w = t.wb(), cells = {};
    [['분류', '매출', '비용', '목표'], ['가', '10', '30', '15'], ['나', '20', '20', '25'], ['다', '30', '10', '35']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
    cells['6,0'] = { raw: '=1+1' };
    w.restore({ sheets: [{ name: '차트 합성', cells }, { name: '다른 합성', cells: { '0,0': { raw: '다른 시트 보존' } } }] }); t.switchSheet(0);
    t.gv().setZoom(zoom); t.gv().layout(); t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 3 }); t.run('chartColumn');
    w.transact(() => w.setSheetProp(0, 'charts', [{ ...w.sheets[0].charts[0], x: 200, y: 60, w: 600, h: 360, title: '합성 제목', legend: 'b', ...patch }]));
    t.gv().renderObjectsAll(); w.undoStack = []; w.redoStack = [];
    window.__chartCellSource = w;
    window.__chartCellsBefore = JSON.stringify(w.serialize().sheets.map(s => s.cells));
  }, { patch, zoom });
}
const cellsUnchanged = p => p.evaluate(() => {
  if (JSON.stringify(window.__chartCellSource.serialize().sheets.map(s => s.cells)) !== window.__chartCellsBefore) throw Error('차트 편집이 원본 셀/수식을 변경함');
});
const element = (p, kind) => p.locator(`.obj.chart [data-el="${kind}"]${kind === 'legend' ? ' text' : ''}`).first();
const point = (p, s = 0, n = 1) => p.locator(`.obj.chart svg [data-s="${s}"][data-p="${n}"]:not(text)`).first();
const center = async locator => { const b = await locator.boundingBox(); assert.ok(b, '요소 화면 위치'); return { x: b.x + b.width / 2, y: b.y + b.height / 2 }; };
async function drag(p, target, dx, dy, cancel = false) {
  const a = typeof target.x === 'number' ? target : await center(target);
  await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(a.x + dx, a.y + dy, { steps: 8 });
  if (cancel) await p.keyboard.press('Escape'); await p.mouse.up();
}
const undo = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+z'); };
const redo = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+y'); };
const pane = p => p.getByRole('dialog').filter({ has: p.locator('.cfp, .chart-selection-pane') });
const picker = p => pane(p).getByRole('combobox', { name: '서식을 지정할 차트 요소' });
const color = (loc, value) => loc.evaluate((el, value) => { el.value = value; el.dispatchEvent(new Event('change', { bubbles: true })); }, value);
async function piePosition(p, index = 1) {
  return point(p, 0, index).evaluate(el => {
    const d = el.dataset, angle = Number(d.pieAngle), radius = Number(d.pieR), squash = Number(d.pieSquash || 1);
    const v = new DOMPoint(Number(d.pieCx) + Math.cos(angle) * radius * .4, Number(d.pieCy) + Math.sin(angle) * radius * .4 * squash).matrixTransform(el.ownerSVGElement.getScreenCTM());
    return { x: v.x, y: v.y, dx: Math.cos(angle) * 45, dy: Math.sin(angle) * 45 * squash };
  });
}

const number = async (loc, value) => { await loc.fill(String(value)); await loc.press('Tab'); };
const openPart = async (p, kind='series', series=0) => { await point(p, series, 1).click(); await command(p, 'chartFormat'); await pane(p).waitFor(); if(kind !== 'series') await picker(p).selectOption(kind); };
try {
  await test('삽입 리본에서 선버스트·대조원형·표면·지도에 접근', async p => {
    await p.locator('[data-ribbon-tab="insert"]').click();
    for(const [menu,name] of [['계층 구조 차트 삽입','선버스트'],['원형 또는 도넛형 차트 삽입','대조 원형'],['표면형 또는 등고선형 차트 삽입','등고선형'],['지도 차트 삽입','국가/지역 색칠지도']]) {
      await p.getByRole('button',{name:menu,exact:true}).click();
      assert.equal(await p.locator('.ct-chip').getByText(name).count(),0); // SVG thumbnails have accessible button names, not duplicate text.
      await p.getByRole('button',{name,exact:true}).waitFor();
      if(name==='선버스트'){await p.getByRole('button',{name,exact:true}).click(); assert.equal(await p.evaluate(()=>tabula.wb().sheets[0].charts.at(-1).type),'sunburst');}
      else await p.keyboard.press('Escape');
    }
  });
  await test('검색과 Enter 확정은 보이는 하위 유형에 일치', async p => {
    await command(p,'chartChangeType'); const d=p.getByRole('dialog',{name:'차트 종류 변경',exact:true});
    await d.getByRole('searchbox').fill('피라미드');
    assert.match(await d.locator('.cg-sub.on:visible').getAttribute('title'),/피라미드/);
    await d.getByRole('button',{name:'확인',exact:true}).click(); assert.equal((await current(p)).barShape,'pyramid');
    await command(p,'chartChangeType'); assert.match(await d.locator('.cg-sub.on').getAttribute('title'),/피라미드/);
    await d.getByRole('button',{name:'취소',exact:true}).click();
  });
  await test('모든 차트 버튼은 추천이 아닌 전체 목록을 직접 엶', async p=>{
    await command(p,'insertChartCatalog');const d=p.getByRole('dialog',{name:'차트 삽입',exact:true});assert.equal(await d.locator('.cg-wrap').isVisible(),true);assert.equal(await d.locator('.cg-cat').count(),17);await d.getByRole('button',{name:'취소',exact:true}).click();
  });
  await test('누적 영역 콤보의 계열 종류·배치·축과 재열기',async p=>{
    await command(p,'chartChangeType');const d=p.getByRole('dialog',{name:'차트 종류 변경',exact:true});await d.locator('.cg-cat').getByText('콤보',{exact:true}).click();await d.getByRole('button',{name:'누적 영역형 - 묶은 세로 막대형',exact:true}).click();
    assert.deepEqual(await d.locator('[data-combo-type]').evaluateAll(ns=>ns.map(n=>n.value)),['area','area','column']);
    await d.getByRole('button',{name:'확인',exact:true}).click();const c=await current(p);assert.equal(c.comboLayout,'areaColumn');assert.deepEqual(c.seriesFmt.map(f=>f.grouping),['stacked','stacked','clustered']);
    await command(p,'chartChangeType');assert.match(await d.locator('.cg-sub.on').getAttribute('title'),/누적 영역형/);await d.getByRole('button',{name:'취소',exact:true}).click();
  });
  await test('선택 계열의 종류·보조축·표식·레이블 실제렌더와 Undo',async p=>{
    await openPart(p);await pane(p).getByLabel('계열 차트 종류',{exact:true}).selectOption('line');await pane(p).getByLabel('계열 표시 축',{exact:true}).selectOption('1');
    await pane(p).getByLabel('표식 모양',{exact:true}).selectOption('diamond');await pane(p).getByLabel('값 표시',{exact:true}).check();await pane(p).getByLabel('항목 이름',{exact:true}).check();
    assert.equal((await current(p)).seriesFmt[0].axis,1);assert.match(await p.locator('.obj.chart [data-el="label"][data-s="0"]').first().textContent(),/가/);
    await command(p,'undo');await p.waitForFunction(()=>!document.querySelector('.chart-selection-pane input[aria-label="항목 이름"]')?.checked);assert.equal(await pane(p).getByLabel('항목 이름',{exact:true}).isChecked(),false);
    await p.screenshot({path:out+'/series-format.png'});
  });
  await test('축 직접 선택·범위·Delete·Undo는 다른 축/계열을 보존',async p=>{
    await p.locator('.obj.chart [data-el="axis-y"] text').first().click();assert.equal((await state(p)).kind,'axis-y');await command(p,'chartFormat');
    await number(pane(p).getByLabel('최대값',{exact:true}),80);await number(pane(p).getByLabel('주 단위',{exact:true}),20);assert.equal((await current(p)).axes.y.max,80);
    assert.match(await p.locator('.obj.chart [data-el="axis-y"]').textContent(),/80/);
    const before=await current(p);await pane(p).getByRole('button',{name:'축 삭제',exact:true}).click();assert.equal((await current(p)).axes.y.hide,true);await command(p,'undo');assert.deepEqual(await current(p),before);
  });
  await test('선버스트 부모 선택·하위색 상속·필터·Undo',async p=>{
    await p.evaluate(()=>{const t=tabula,w=t.wb(),cells={};[['지역','제품','매출'],['서울','가','10'],['서울','나','20'],['부산','가','30']].forEach((row,r)=>row.forEach((raw,c)=>cells[`${r},${c}`]={raw}));w.restore({sheets:[{name:'계층 합성',cells,charts:[{id:'hierarchy',type:'sunburst',range:{r1:0,c1:0,r2:3,c2:2},x:170,y:70,w:680,h:480,title:'지역별 제품 매출',legend:'b'}]}]});t.gv().layout();t.gv().renderObjectsAll();window.__chartCellsBefore=JSON.stringify(w.serialize().sheets.map(s=>s.cells));});
    const parent=p.locator('.obj.chart svg path[data-node]').filter({hasNot:p.locator('[data-p]')}).first();
    await parent.dispatchEvent('mousedown',{button:0});await p.mouse.up();assert.equal((await state(p)).kind,'node');await command(p,'chartFormat');
    await color(pane(p).getByLabel('계층 항목 색',{exact:true}),'#ab1256');assert.ok(Object.values((await current(p)).seriesFmt[0].hierarchyColors).includes('#ab1256'));
    assert.ok(await p.locator('.obj.chart svg [fill="#ab1256"]').count());await command(p,'undo');assert.equal((await current(p)).seriesFmt,undefined);
    await p.screenshot({path:out+'/sunburst-selection.png'});
  });
  await test('완료 표시된 문서는 열린 차트 서식에서 변경할 수 없음',async p=>{
    await openPart(p);await p.evaluate(()=>{tabula.wb().props.markedFinal=true;});const before=await current(p);await color(pane(p).getByLabel('계열 채우기 색',{exact:true}),'#001122');assert.deepEqual(await current(p),before);
  });
  await test('390px 모바일의 선택 서식 컨트롤은 화면을 벗어나지 않음',async p=>{
    await openPart(p);await p.setViewportSize({width:390,height:844});const d=pane(p);await d.getByLabel('계열 표시 축',{exact:true}).scrollIntoViewIfNeeded();
    for(const el of [d,d.getByLabel('계열 표시 축',{exact:true}),picker(p)]){const b=await el.boundingBox();assert.ok(b&&b.x>=0&&b.x+b.width<=391);}
    assert.ok((await picker(p).boundingBox()).height < 60, '요소 선택기가 세로로 과도하게 늘어남');
    await p.screenshot({path:out+'/format-mobile.png'});
  });
} finally { await browser.close(); await writeFile(out+'/results.json',JSON.stringify({url,results},null,2)); }
console.log(JSON.stringify({tests:results.length,passed:results.filter(r=>r.ok).length,results}));if(results.some(r=>!r.ok))process.exitCode=1;
