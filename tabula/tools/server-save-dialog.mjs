// 실제 UI + 합성 보관함 API. 모든 /api 요청은 메모리에서 처리하며 원격 파일은 쓰지 않는다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
assert.ok(['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname));
const browser = await chromium.launch(), results = [];
const dialogTitle = '서버에 저장 (다른 기기에서 열기)';
const snap = (name, value) => ({ app: 'wixel', docName: name, si: 0, workbook: { sheets: [{ name: '합성', cells: { '0,0': { raw: String(value) } } }] } });
function gate() { let release, started; return { promise: new Promise(r => { release = r; }), ready: new Promise(r => { started = r; }), get release() { return release; }, get started() { return started; } }; }
async function test(name, action) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage(), errors = [], unexpected = [];
  const state = { docs: new Map([['대상', { revision: 7, data: snap('대상', 99) }]]), writes: [], listGate: null, failList: false };
  page.setDefaultTimeout(15000); page.on('pageerror', e => errors.push(e.message));
  await context.route('**/*', async route => {
    const req = route.request(), path = new URL(req.url()).pathname;
    const respond = (status, data, revision) => route.fulfill({ status, contentType: 'application/json', headers: revision === undefined ? {} : { 'X-Wixel-Revision': String(revision) }, body: JSON.stringify(data) });
    if (path.endsWith('/api/health')) return respond(200, { ok: true, vault: true });
    if (path.endsWith('/api/files')) {
      const rows = [...state.docs].map(([name, doc]) => ({ name, revision: doc.revision, modified: 1, size: 100 }));
      const delay = state.listGate; if (delay) { state.listGate = null; delay.started(); await delay.promise; }
      return state.failList ? respond(503, { error: '합성 목록 실패' }) : respond(200, rows);
    }
    const file = /\/api\/files\/(.+)$/.exec(path);
    if (file) {
      const name = decodeURIComponent(file[1]), old = state.docs.get(name);
      if (req.method() === 'GET') return old ? respond(200, old.data, old.revision) : respond(404, { error: '없음' });
      if (req.method() === 'PUT') {
        const expected = Number((req.headers()['if-match'] ?? '').replaceAll('"', '')), data = req.postDataJSON();
        state.writes.push({ name, expected, data });
        if (expected !== (old?.revision ?? 0)) return respond(412, { error: '다른 기기에서 문서가 변경되었습니다.', code: 'REVISION_CONFLICT', currentRevision: old?.revision ?? 0 });
        const revision = (old?.revision ?? 0) + 1; state.docs.set(name, { revision, data });
        return respond(200, { revision, modified: 1234 }, revision);
      }
    }
    if (path.includes('/api/')) { unexpected.push(req.method() + ' ' + path); return respond(404, { error: '예상하지 않은 합성 API' }); }
    if (!['GET', 'HEAD'].includes(req.method())) { unexpected.push(req.method() + ' ' + path); return route.abort(); }
    return route.continue();
  });
  await page.addInitScript(() => {
    window.WIXEL_SKIP_START = true;
    localStorage.setItem('wixel.connection.v3', JSON.stringify({ kind: 'vault', key: 'A'.repeat(43) }));
    localStorage.setItem('tabula.workbook.v1', JSON.stringify({ docName: '원본문서', docId: 'synthetic-source', autosave: false, remoteDoc: false, workbook: { sheets: [{ name: '합성', cells: { '0,0': { raw: '42' } } }] } }));
  });
  try {
    const health = page.waitForResponse(r => r.url().endsWith('/api/health'));
    await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await (await health).finished();
    await page.waitForFunction(() => window.tabula?.wb()?.getValue(0, 0, 0) === 42);
    await action(page, state);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(unexpected, [], '실제/예상 밖 API 쓰기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (err) { results.push({ name, ok: false, error: err.message, pageErrors: errors, unexpected }); console.error('NG ' + name + ': ' + err.stack); }
  finally { state.listGate?.release(); await context.close(); }
}
async function start(p) { await p.evaluate(() => window.tabula.run('saveLocations')); await p.getByRole('button', { name: /^온라인 개인 보관함/ }).click(); }
async function form(p, name = '대상') { const d = p.getByRole('dialog', { name: dialogTitle, exact: true }); await d.waitFor(); await d.getByLabel('파일 이름', { exact: true }).fill(name); return d; }
async function confirm(p) { await p.getByRole('dialog', { name: '서버에 저장', exact: true }).getByRole('button', { name: '확인', exact: true }).click(); }
async function settled(p) { await p.evaluate(() => new Promise(r => setTimeout(r, 120))); }
async function switchBook(p) {
  await p.locator('#fileInput').setInputFiles({ name: '다른문서.wixel', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(snap('다른문서', 84))) });
  await p.waitForFunction(() => window.tabula.wb().getValue(0, 0, 0) === 84);
}
async function assertOther(p, state) { await settled(p); assert.equal(state.writes.length, 0); assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 84); assert.match(await p.locator('#docTitle').innerText(), /다른문서/); assert.equal(state.docs.get('대상').data.workbook.sheets[0].cells['0,0'].raw, '99'); }
try {
  await test('기존 대상 덮어쓰기 확인은 목록의 revision으로 CAS 저장', async (p, s) => {
    await start(p); const d = await form(p); await d.getByRole('button', { name: '저장', exact: true }).click(); await confirm(p);
    await p.waitForFunction(() => document.querySelector('#saveState').textContent.includes('저장됨'));
    await settled(p); assert.equal(s.writes.length, 1); assert.equal(s.writes[0].expected, 7); assert.equal(s.docs.get('대상').revision, 8);
    assert.equal(s.docs.get('대상').data.workbook.sheets[0].cells['0,0'].raw, '42'); assert.match(await p.locator('#docTitle').innerText(), /대상/);
  });
  await test('목록 이후 대상 변경은 최신 서버 내용 보존·412 안내', async (p, s) => {
    await start(p); const d = await form(p); s.docs.set('대상', { revision: 8, data: snap('대상', 777) });
    await d.getByRole('button', { name: '저장', exact: true }).click(); await confirm(p);
    await p.getByRole('dialog', { name: '저장 충돌 — 사본을 유지했습니다', exact: true }).waitFor();
    assert.equal(s.writes.length, 1); assert.equal(s.writes[0].expected, 7); assert.equal(s.docs.get('대상').data.workbook.sheets[0].cells['0,0'].raw, '777');
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 42);
  });
  await test('새 이름은 revision0으로 만들며 기존 대상 보존', async (p, s) => {
    await start(p); const d = await form(p, '새 사본'); await d.getByRole('button', { name: '저장', exact: true }).click(); await settled(p);
    assert.equal(s.writes.length, 1); assert.equal(s.writes[0].expected, 0); assert.equal(s.docs.get('새 사본').data.workbook.sheets[0].cells['0,0'].raw, '42'); assert.equal(s.docs.get('대상').revision, 7);
  });
  await test('목록 응답 대기 중 문서 전환은 오래된 저장 폼도 열지 않음', async (p, s) => {
    const delay = gate(); s.listGate = delay; await start(p); await delay.ready;
    try { await switchBook(p); } finally { delay.release(); }
    await assertOther(p, s); assert.equal(await p.getByRole('dialog', { name: dialogTitle, exact: true }).count(), 0);
  });
  await test('폼 표시 후 문서 전환은 submit 시 새 문서 이름·내용 보존', async (p, s) => {
    await start(p); const d = await form(p, '새 이름'); await switchBook(p); await d.getByRole('button', { name: '저장', exact: true }).click(); await assertOther(p, s);
  });
  await test('덮어쓰기 확인 중 문서 전환은 확인 후에도 쓰기 없음', async (p, s) => {
    await start(p); const d = await form(p); await d.getByRole('button', { name: '저장', exact: true }).click();
    await p.getByRole('dialog', { name: '서버에 저장', exact: true }).waitFor(); await switchBook(p); await confirm(p); await assertOther(p, s);
  });
  await test('목록 대기 중 보관함 연결 해제는 저장 폼·쓰기 없음', async (p, s) => {
    const delay = gate(); s.listGate = delay; await start(p); await delay.ready;
    try { await p.getByRole('button', { name: '이 기기 연결 해제', exact: true }).click(); } finally { delay.release(); }
    await settled(p); assert.equal(s.writes.length, 0); assert.equal(await p.getByRole('dialog', { name: dialogTitle, exact: true }).count(), 0); assert.match(await p.locator('#docTitle').innerText(), /원본문서/);
  });
  await test('목록 실패는 빈 보관함으로 오인하지 않고 이름·대상 유지', async (p, s) => {
    s.failList = true; await start(p); await settled(p); assert.equal(s.writes.length, 0); assert.equal(await p.getByRole('dialog', { name: dialogTitle, exact: true }).count(), 0); assert.match(await p.locator('#docTitle').innerText(), /원본문서/); assert.equal(s.docs.get('대상').revision, 7);
  });
  await test('덮어쓰기 취소는 현재 문서 이름·브라우저 저장 위치 보존', async (p, s) => {
    await start(p); const d = await form(p); await d.getByRole('button', { name: '저장', exact: true }).click();
    await p.getByRole('dialog', { name: '서버에 저장', exact: true }).getByRole('button', { name: '취소', exact: true }).click(); await settled(p);
    assert.equal(s.writes.length, 0); assert.match(await p.locator('#docTitle').innerText(), /원본문서/);
    assert.equal(await p.evaluate(() => JSON.parse(localStorage.getItem('tabula.workbook.v1')).remoteDoc), false); assert.equal(s.docs.get('대상').revision, 7);
  });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(r => r.ok).length, bad: results.filter(r => !r.ok), note: '실제 UI + 메모리 API. 원격 보관함·사용자 파일 쓰기 없음.' }, null, 2));
if (results.some(r => !r.ok)) process.exitCode = 1;
