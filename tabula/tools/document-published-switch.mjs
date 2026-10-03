// 합성 게시본·URL 응답의 지연을 제어해 문서 전환 뒤 이전 내용이 덮이는 경쟁을 검사합니다.
// 모든 API 응답은 메모리에서 만들며 원격 문서를 읽거나 쓰지 않습니다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const origin = new URL(url).origin;
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 합성 문서 전용 검사입니다.');
const engines = (process.env.WIXEL_BROWSER || 'chromium,webkit').split(',');
const out = process.env.WIXEL_DOCUMENT_PUBLISHED_OUT || 'D:/Codex/Temp/wixel-document-switch/published';
const only = process.env.WIXEL_DOCUMENT_PUBLISHED_FILTER || '';
await mkdir(out, { recursive: true });
const results = [], assets = new Set(); let checks = 0;
const eq = (a, b, message) => { checks++; assert.deepEqual(a, b, message); };
const snap = (name, value) => ({ docName: name, workbook: { sheets: [{ name, cells: { '0,0': { raw: value } } }] } });
const state = p => p.evaluate(() => ({ title: document.querySelector('#docTitle').textContent, value: tabula.wb().getValue(0,0,0), edit: tabula.wb().getValue(0,1,0), readonly: document.body.classList.contains('view-mode') }));
function gate() { let release, started; return { promise: new Promise(r => { release = r; }), ready: new Promise(r => { started = r; }), release: () => release(), started: () => started() }; }
const settle = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function poll(p) { await p.evaluate(() => Promise.all([...window.__publishedIntervals.values()].map(fn => fn()))); await settle(p); }
async function beginPoll(p) { await p.evaluate(() => { window.__pendingPublishedPoll = Promise.all([...window.__publishedIntervals.values()].map(fn => fn())); }); }
async function openB(p) {
  await p.locator('#fileInput').setInputFiles({ name: 'B.wixel', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(snap('B', 'BBB'))) });
  await p.waitForFunction(() => tabula.wb().getValue(0,0,0) === 'BBB');
  await p.evaluate(() => tabula.wb().transact(() => tabula.wb().setInput(0,1,0,'B edit')));
  await settle(p);
}
async function publishedReady(p) {
  await p.waitForFunction(() => window.__publishedIntervals?.size === 1 && window.tabula?.wb()?.getValue(0,0,0) === 'A1');
  // 최초 1회는 기존 구현의 seen 초기화도 끝낸다.
  await poll(p);
}
const pivotFixture = (pending = false) => {
  const rawName = pending ? 'PendingRaw' : 'Raw';
  const pivot = { name:'Same', source:rawName, range:{r1:0,c1:0,r2:2,c2:1}, top:0, left:0, rows:['Group'], cols:[], pages:[], values:[{field:'Amount',agg:'sum'}], layout:'tabular', needsRender:true };
  return {sheets:[{name:rawName,cells:{'0,0':{raw:'Group'},'0,1':{raw:'Amount'},'1,0':{raw:'A'},'1,1':{raw:'10'},'2,0':{raw:'B'},'2,1':{raw:'20'}}},{name:'Pivot',cells:{},pivot,...(pending?{pivotsExtra:[{...pivot,name:'Second',left:4}]}:{})}]};
};
async function openPivotAndColor(p) {
  await p.locator('#fileInput').setInputFiles({name:'Pivot.wixel',mimeType:'application/json',buffer:Buffer.from(JSON.stringify({docName:'Pivot',workbook:pivotFixture()}))});
  await p.waitForFunction(()=>tabula.wb().sheets[1]?.name==='Pivot');
  await p.evaluate(()=>{tabula.switchSheet(1);tabula.wb().transact(()=>tabula.wb().setStyle(1,1,1,{fill:'#123456'}));}); await settle(p);
}
const tests = [
  { name:'opened-pivot-keeps-direct-format-on-first-refresh', query:'', async run(p,s) {
    await openPivotAndColor(p); await p.evaluate(()=>tabula.run('pivotRefresh')); await settle(p);
    eq(await p.evaluate(()=>tabula.wb().getCell(1,1,1)?.style?.fill),'#123456','준비한 피벗 서식 기준을 승계해 첫 새로 고침도 직접 서식 유지');
  } },
  { name:'cancelled-pivot-preparation-keeps-current-document-format', query:'', async run(p,s) {
    await openPivotAndColor(p);
    await p.evaluate(()=>{
      window.__pivotPreparation={written:false,paused:false};
      const originalSet=window.setTimeout, proto=tabula.wb().constructor.prototype, originalTransact=proto.transact;
      proto.transact=function(...args){const result=originalTransact.apply(this,args);if(this.sheets[0]?.name==='PendingRaw' && this.sheets[1]?.cells?.getRC(1,1))window.__pivotPreparation.written=true;return result;};
      window.setTimeout=(fn,ms,...args)=>{if(ms===0&&window.__pivotPreparation.written&&!window.__pivotPreparation.paused){window.__pivotPreparation.paused=true;window.__releasePivotPreparation=()=>originalSet(fn,0,...args);return 954000;}return originalSet(fn,ms,...args);};
    });
    await p.locator('#fileInput').setInputFiles({name:'Pending.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(writeXlsx(new Workbook(pivotFixture(true))))});
    await p.waitForFunction(()=>window.__pivotPreparation.paused);
    eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'Raw','피벗 준비 중 현재 문서가 화면에 유지');
    await p.locator('#fileInput').setInputFiles({name:'Invalid.wixel',mimeType:'application/json',buffer:Buffer.from('{invalid')});
    await p.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();
    await p.evaluate(()=>window.__releasePivotPreparation()); await p.waitForFunction(()=>!document.querySelector('.load-progress'));
    await p.evaluate(()=>tabula.run('pivotRefresh')); await settle(p);
    eq(await p.evaluate(()=>tabula.wb().sheets[0].name),'Raw','취소된 준비 문서가 현재 문서로 반영되지 않음');
    eq(await p.evaluate(()=>tabula.wb().getCell(1,1,1)?.style?.fill),'#123456','취소된 피벗 준비는 현재 문서 서식 기준을 바꾸지 않음');
  } },
  { name: 'old-published-timer-after-copy-and-local-open', query: '?view=synthetic', async run(p, s) {
    await publishedReady(p);
    await p.getByRole('button', { name: '편집용 사본 만들기', exact: true }).click();
    await openB(p); const before = await state(p);
    await poll(p);
    eq(await state(p), before, '이전 게시본 타이머가 새 로컬 문서·편집을 바꾸지 않음');
    eq(await p.evaluate(() => window.__publishedIntervals.size), 0, '보기 종료 시 게시본 타이머 해제');
  } },
  { name: 'inflight-published-response-after-copy-and-local-open', query: '?view=synthetic', async run(p, s) {
    await publishedReady(p);
    const delay = gate(); s.publishedGate = delay; s.gates.push(delay);
    await beginPoll(p); await delay.ready;
    await p.getByRole('button', { name: '편집용 사본 만들기', exact: true }).click();
    await openB(p); const before = await state(p);
    delay.release(); await p.evaluate(() => window.__pendingPublishedPoll); await settle(p);
    eq(await state(p), before, '이미 출발한 이전 게시본 응답도 새 로컬 문서·편집을 바꾸지 않음');
  } },
  { name: 'current-published-view-still-refreshes', query: '?view=synthetic', async run(p, s) {
    await publishedReady(p); await poll(p);
    eq((await state(p)).value, 'A3', '현재 게시 보기에는 최신 서버 내용 적용');
    eq((await state(p)).readonly, true, '현재 게시 보기의 읽기 전용 유지');
    eq(await p.evaluate(() => window.__publishedIntervals.size), 1, '현재 게시 보기 타이머 유지');
  } },
  { name: 'copy-of-refreshed-published-view-keeps-current-data', query: '?view=synthetic', async run(p, s) {
    await publishedReady(p); await poll(p);
    eq((await state(p)).value, 'A3', '사본 만들기 전 현재 게시본 최신 내용');
    await p.getByRole('button', { name: '편집용 사본 만들기', exact: true }).click(); await settle(p);
    eq((await state(p)).value, 'A3', '편집용 사본은 최초 응답이 아닌 최신 화면 내용');
    eq((await state(p)).readonly, false, '편집용 사본 읽기 전용 해제');
    eq(await p.evaluate(() => window.__publishedIntervals.size), 0, '편집용 사본 타이머 해제');
  } },
  { name:'failed-file-open-keeps-current-published-refresh-and-copy', query:'?view=synthetic', async run(p,s) {
    await publishedReady(p);
    await p.locator('#fileInput').setInputFiles({name:'Invalid.wixel',mimeType:'application/json',buffer:Buffer.from('{invalid')});
    await p.getByRole('dialog').getByRole('button',{name:'확인',exact:true}).click();
    await poll(p);
    eq((await state(p)).value,'A3','파일 열기 실패 뒤 현재 게시본 갱신을 재개');
    eq((await state(p)).readonly,true,'파일 열기 실패 뒤 현재 게시보기 유지');
    await p.getByRole('button',{name:'편집용 사본 만들기',exact:true}).click(); await settle(p);
    eq((await state(p)).value,'A3','파일 열기 실패 뒤 편집용 사본 버튼도 정상 동작');
    eq((await state(p)).readonly,false,'파일 열기 실패 뒤 편집용 사본 읽기 전용 해제');
  } },
  { name: 'late-startup-health-does-not-open-old-url', query: '?doc=A-url', startup: 'health', async run(p, s) {
    await s.startupGate.ready; await p.waitForFunction(() => window.tabula);
    await openB(p); const before = await state(p);
    const response = p.waitForResponse(r => r.url().endsWith('/api/health')); s.startupGate.release(); await (await response).finished();
    await p.waitForTimeout(150); await settle(p);
    eq(await state(p), before, '시작 서버 확인 지연 뒤 URL의 옛 문서를 열지 않음');
  } },
  { name: 'late-startup-published-response-does-not-replace-local-open', query: '?view=synthetic', startup: 'published', async run(p, s) {
    await s.startupGate.ready; await p.waitForFunction(() => window.tabula);
    await openB(p); const before = await state(p);
    const response = p.waitForResponse(r => r.url().includes('/api/published/')); s.startupGate.release(); await (await response).finished();
    await p.waitForTimeout(150); await settle(p);
    eq(await state(p), before, '시작 게시본 응답 지연 뒤 새 로컬 문서·편집 유지');
  } },
];
for (const engine of engines) {
  const browser = await pw[engine].launch();
  try { for (const t of tests) {
    const label = engine + '-' + t.name; if (only && !label.includes(only)) continue;
    const context = await browser.newContext({ viewport: { width: 1024, height: 768 }, hasTouch: true, isMobile: true, serviceWorkers: 'block', userAgent: 'Mozilla/5.0 (iPad; CPU OS 18_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/18.0 Mobile/15E148 Safari/604.1' });
    const p = await context.newPage(), errors = [], writes = [], start = checks;
    const s = { revision: 0, publishedGate: null, startupGate: t.startup ? gate() : null, gates: [] };
    if (s.startupGate) s.gates.push(s.startupGate);
    p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message));
    await context.addInitScript(() => {
      window.WIXEL_SKIP_START = true;
      localStorage.setItem('wixel.connection.v3', JSON.stringify({ kind: 'node' }));
      const originalSet = window.setInterval, originalClear = window.clearInterval;
      window.__publishedIntervals = new Map(); let serial = 950000;
      window.setInterval = (fn, ms, ...args) => { if (ms === 30000) { const id = ++serial; window.__publishedIntervals.set(id, () => fn(...args)); return id; } return originalSet(fn, ms, ...args); };
      window.clearInterval = id => { window.__publishedIntervals.delete(id); originalClear(id); };
    });
    await context.route('**/*', async route => {
      const request = route.request(), u = new URL(request.url()), path = u.pathname;
      if (u.origin !== origin || !['GET', 'HEAD'].includes(request.method())) { writes.push(request.method() + ' ' + request.url()); return route.abort(); }
      const respond = (data, headers = {}) => route.fulfill({ contentType: 'application/json', headers, body: JSON.stringify(data) });
      if (path.endsWith('/api/health')) {
        if (t.startup === 'health') { s.startupGate.started(); await s.startupGate.promise; }
        return respond({ ok: true });
      }
      if (path.includes('/api/published/')) {
        const revision = ++s.revision;
        if (t.startup === 'published' && revision === 1) { s.startupGate.started(); await s.startupGate.promise; }
        const delay = s.publishedGate; s.publishedGate = null;
        if (delay) { delay.started(); await delay.promise; }
        return respond(snap('A-published', 'A' + revision), { 'X-Modified': String(revision) });
      }
      if (path.includes('/api/files/')) return respond(snap('A-url', 'AAA'));
      if (path.endsWith('/api/files')) return respond([]);
      if (path.includes('/api/')) return route.fulfill({ status: 404, contentType: 'application/json', body: '{}' });
      return route.continue();
    });
    try {
      const target = new URL(url); target.search = t.query;
      await p.goto(target.href, { waitUntil: 'domcontentloaded', timeout: 60000 });
      await p.waitForFunction(() => window.tabula);
      assets.add(await p.evaluate(() => [...document.scripts].map(x => x.getAttribute('src')).find(x => x?.includes('wixel-')) || 'source'));
      await t.run(p, s); eq(errors, [], '실행 오류 없음'); eq(writes, [], '원격 요청·쓰기 없음');
      results.push({ label, ok: true, checks: checks - start }); console.log('OK ' + label);
    } catch (error) {
      results.push({ label, ok: false, error: error.message, checks: checks - start, errors, writes, state: await state(p).catch(() => null) });
      console.error('NG ' + label + ': ' + error.message);
      await p.screenshot({ path: out + '/' + label + '-failure.png' }).catch(() => {});
    } finally { for (const g of s.gates) g.release(); await context.close(); }
  } } finally { await browser.close(); }
}
const result = { url, assets: [...assets], cases: results.length, passed: results.filter(x => x.ok).length, checks, results };
await writeFile(out + '/result.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify(result));
if (result.passed !== result.cases) process.exitCode = 1;
