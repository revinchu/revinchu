// 실제 키보드/UI + 실제 IndexedDB, 온라인 API는 격리된 메모리 응답만 사용합니다.
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium';
const base = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5191/');
assert.ok(['localhost', '127.0.0.1', '[::1]'].includes(base.hostname), '고장 주입은 로컬 서버만 허용');
const out = process.env.WIXEL_RECOVERY_OUT || 'D:/Codex/Temp/wixel-command-recovery/' + engine;
const baseline = process.env.WIXEL_BASELINE_REF;
const oldApp = baseline ? execFileSync('git', ['-c', 'safe.directory=D:/위셀/wixel-3', 'show', `${baseline}:tabula/src/app.js`], { encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 }) : null;
await mkdir(out, { recursive: true });
const browser = engine === 'webkit' ? null : await pw[engine].launch(), results = [], assets = new Set();
const snap = (value = 42, name = '합성 충돌') => ({ app: 'wixel', docName: name, docId: 'synthetic-source', si: 0, workbook: { sheets: [{ name: '합성', cells: { '0,0': { raw: String(value) }, '1,0': { raw: '보존' } } }] } });
const conflictTitle = '저장 충돌 — 사본을 유지했습니다';
const d = p => p.getByRole('dialog', { name: conflictTitle, exact: true });
const reopen = p => d(p).getByRole('button', { name: '온라인 문서 다시 열기', exact: true }).click();
const value = p => p.evaluate(() => tabula.wb().getValue(0, 0, 0));
const undoCount = p => p.evaluate(() => tabula.wb().undoStack.length);
const state = p => p.evaluate(() => ({ sheets: tabula.wb().sheets.length, undo: tabula.wb().undoStack.length, raw: tabula.wb().getRaw(tabula.si, 0, 0), style: tabula.wb().getCell(tabula.si, 0, 0)?.style ?? {} }));
async function key(p, key) { await p.locator('#cellEditor').focus(); await p.keyboard.press(key); }
async function conflict(p, s) {
  await p.locator('#autosaveToggle').click();
  await d(p).waitFor();
  assert.equal(s.writes, 1); assert.equal(await value(p), 42);
}
async function switchBook(p) {
  await p.locator('#fileInput').setInputFiles({ name: '다른 합성 문서.wixel', mimeType: 'application/json', buffer: Buffer.from(JSON.stringify(snap(84, '다른 합성 문서'))) });
  await p.waitForFunction(() => tabula.wb().getValue(0, 0, 0) === 84);
}
async function records(p) {
  return p.evaluate(async () => {
    const rows = await new Promise((resolve, reject) => {
      const req = indexedDB.open('tabula', 1);
      req.onupgradeneeded = () => req.result.createObjectStore('docs'); req.onerror = () => reject(req.error);
      req.onsuccess = () => { const db = req.result, tx = db.transaction('docs'), st = tx.objectStore('docs'), a = st.getAll(), b = st.getAllKeys(); tx.oncomplete = () => { db.close(); resolve(b.result.map((key, i) => ({ key, record: a.result[i] })).filter(x => String(x.key).startsWith('lib:ver:'))); }; };
    });
    return Promise.all(rows.map(async ({ key, record }) => ({ key, data: JSON.parse(await (record.gz ? new Response(record.blob.stream().pipeThrough(new DecompressionStream('gzip'))).text() : record.blob.text())) })));
  });
}
async function test(name, action) {
  if (process.env.WIXEL_CASE && !name.includes(process.env.WIXEL_CASE)) return;
  // Windows WebKit의 ephemeral IndexedDB는 Blob 저장 자체가 실패하므로 새 전용 프로필을 씁니다.
  const options = { viewport: { width: 1400, height: 950 }, serviceWorkers: 'block' };
  const context = browser ? await browser.newContext(options) : await pw.webkit.launchPersistentContext(await mkdtemp(out + '/profile-'), options);
  const p = await context.newPage(), errors = [], blocked = [], warnings = [];
  const s = { writes: 0, reads: 0, data: snap(777), holdLoad: null };
  p.setDefaultTimeout(15000); p.on('pageerror', e => errors.push(e.message)); p.on('console', m => { if (m.type() === 'warning' || m.type() === 'error') warnings.push(m.text()); });
  await context.route('**/*', async r => {
    const q = r.request(), u = new URL(q.url());
    const json = (status, data, revision) => r.fulfill({ status, contentType: 'application/json', headers: revision ? { 'X-Wixel-Revision': String(revision) } : {}, body: JSON.stringify(data) });
    if (u.origin !== base.origin) { blocked.push(q.method() + ' external'); return r.fulfill({ status: 204, body: '' }); }
    if (u.pathname === '/api/health') return json(200, { ok: true, vault: true });
    if (u.pathname === '/api/files') return json(200, [{ name: '합성 충돌', revision: 2, modified: 1, size: 100 }]);
    if (u.pathname.startsWith('/api/files/')) {
      if (q.method() === 'PUT') { s.writes++; return json(412, { error: '합성 저장 충돌', code: 'REVISION_CONFLICT', currentRevision: 2 }); }
      if (q.method() === 'GET') { s.reads++; if (s.holdLoad) await s.holdLoad; return json(200, s.data, 2); }
    }
    if (u.pathname.startsWith('/api/')) { blocked.push(q.method() + ' ' + u.pathname); return json(404, { error: '허용하지 않은 합성 API' }); }
    if (!['GET', 'HEAD'].includes(q.method())) { blocked.push(q.method()); return r.abort(); }
    if (oldApp && u.pathname === '/src/app.js') return r.fulfill({ contentType: 'application/javascript', body: oldApp });
    return r.continue();
  });
  await context.addInitScript(snapshot => {
    window.WIXEL_SKIP_START = true;
    localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off');
    localStorage.setItem('wixel.options', JSON.stringify({ saveConfirm: false }));
    localStorage.setItem('wixel.connection.v3', JSON.stringify({ kind: 'vault', key: 'A'.repeat(43) }));
    localStorage.setItem('tabula.workbook.v1', JSON.stringify({ ...snapshot, autosave: false, remoteDoc: true }));
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {}, write: async () => {} } });
  }, snap());
  try {
    const health = p.waitForResponse(r => r.url().endsWith('/api/health'));
    await p.goto(base.href, { waitUntil: 'domcontentloaded', timeout: 60000 }); await (await health).finished();
    await p.waitForFunction(() => tabula?.wb()?.getValue(0, 0, 0) === 42);
    for (const asset of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => new URL(n.src).pathname).filter(n => /wixel-[a-f0-9]+\.js$/.test(n)))) assets.add(asset);
    await action(p, s);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(blocked, [], '예상 밖 원격 요청');
    results.push({ name, ok: true, reads: s.reads, mockedWrites: s.writes }); console.log('OK ' + name);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors, blocked, warnings, current: await state(p).catch(() => null), reads: s.reads }); console.error('NG ' + name + ': ' + e.stack); await p.screenshot({ path: out + '/failure-' + results.length + '.png' }).catch(() => {}); }
  finally { await context.close(); }
}
try {
  await test('F4 정상 명령 반복과 실행 취소·다시 실행', async p => {
    await p.evaluate(() => tabula.run('insertRows')); await key(p, 'F4');
    assert.equal(await p.evaluate(() => tabula.wb().getValue(0, 2, 0)), 42); assert.equal(await undoCount(p), 2);
    await key(p, 'Control+z'); assert.equal(await p.evaluate(() => tabula.wb().getValue(0, 1, 0)), 42);
    await key(p, 'Control+y'); assert.equal(await p.evaluate(() => tabula.wb().getValue(0, 2, 0)), 42);
  });
  await test('F4 구조 보호 전후 현재 상태 검사', async p => {
    await p.evaluate(() => { tabula.run('addSheet'); tabula.wb().props.lockStructure = true; });
    const before = await state(p); await key(p, 'F4'); assert.deepEqual(await state(p), before);
    await p.keyboard.press('Escape');
    await p.evaluate(() => { tabula.wb().props.lockStructure = false; }); await key(p, 'F4');
    assert.equal((await state(p)).sheets, 3); assert.equal(await undoCount(p), before.undo + 1);
  });
  await test('서식 반복과 Ctrl+Y의 정상 반복·Undo', async p => {
    await p.evaluate(() => { tabula.run('fontSize', 18); tabula.selectCell(1, 0); }); await key(p, 'F4');
    assert.equal(await p.evaluate(() => tabula.wb().styleAt(0, 1, 0).size), 18);
    await p.evaluate(() => tabula.selectCell(2, 0)); await key(p, 'Control+y');
    assert.equal(await p.evaluate(() => tabula.wb().styleAt(0, 2, 0).size), 18); assert.equal(await undoCount(p), 3);
    await key(p, 'Control+z'); assert.notEqual(await p.evaluate(() => tabula.wb().styleAt(0, 2, 0).size), 18);
  });
  await test('서식 반복은 보호 시트에서 거부하고 서식 허용 시 실행', async p => {
    await p.evaluate(() => { tabula.run('fontSize', 18); tabula.selectCell(1, 0); tabula.wb().sheets[0].protect = { on: true, allow: { formatCells: false } }; });
    const before = await undoCount(p); await key(p, 'F4');
    assert.equal(await undoCount(p), before); assert.notEqual(await p.evaluate(() => tabula.wb().styleAt(0, 1, 0).size), 18);
    await p.keyboard.press('Escape');
    await p.evaluate(() => { tabula.wb().sheets[0].protect.allow.formatCells = true; }); await key(p, 'F4');
    assert.equal(await p.evaluate(() => tabula.wb().styleAt(0, 1, 0).size), 18); assert.equal(await undoCount(p), before + 1);
  });
  await test('서식 반복은 최종본에서 거부', async p => {
    await p.evaluate(() => { tabula.run('fontSize', 18); tabula.selectCell(1, 0); tabula.wb().props.markedFinal = true; });
    const before = await undoCount(p); await key(p, 'F4');
    assert.equal(await undoCount(p), before); assert.notEqual(await p.evaluate(() => tabula.wb().styleAt(0, 1, 0).size), 18);
  });
  await test('명령 반복은 시트 전환 후 현재 시트 보호 검사', async p => {
    await p.evaluate(() => { tabula.run('clearContents'); const w = tabula.wb(); w.transact(() => { w.addSheet('보호 시트'); w.setInput(1, 0, 0, '잠긴 값'); w.setSheetProp(1, 'protect', { on: true, allow: {} }); }); tabula.switchSheet(1); });
    const before = await state(p); await key(p, 'F4'); assert.deepEqual(await state(p), before);
  });
  await test('충돌 다시 열기 성공은 현재 값을 실제 버전 기록에 보존', async (p, s) => {
    await conflict(p, s); await reopen(p); await p.waitForFunction(() => tabula.wb().getValue(0, 0, 0) === 777);
    const saved = await records(p); assert.ok(saved.some(x => x.data.workbook.sheets[0].cells['0,0'].raw === '42'));
    assert.equal(s.reads, 1); assert.equal(s.writes, 1); await d(p).waitFor({ state: 'detached' });
  });
  await test('복구 사본 IDB 실패는 교체 금지·오류 안내·재시도 성공', async (p, s) => {
    await conflict(p, s);
    await p.evaluate(() => { const put = IDBObjectStore.prototype.put; window.__restorePut = () => { IDBObjectStore.prototype.put = put; }; IDBObjectStore.prototype.put = function(v, k) { if (String(k).startsWith('lib:')) throw new DOMException('합성 용량 초과', 'QuotaExceededError'); return put.call(this, v, k); }; });
    await reopen(p); await d(p).getByRole('alert').filter({ hasText: '복구 사본을 보관하지 못해' }).waitFor();
    assert.equal(await value(p), 42); assert.equal(s.reads, 0); assert.equal(s.writes, 1);
    await p.evaluate(() => window.__restorePut()); await reopen(p); await p.waitForFunction(() => tabula.wb().getValue(0, 0, 0) === 777);
    assert.ok((await records(p)).some(x => x.data.workbook.sheets[0].cells['0,0'].raw === '42'));
  });
  await test('복구 사본 직렬화 실패도 현재 문서 유지', async (p, s) => {
    await conflict(p, s);
    await p.evaluate(() => { tabula.wb().serializeBlob = () => { throw new Error('합성 직렬화 실패'); }; });
    await reopen(p); await d(p).getByRole('alert').filter({ hasText: '복구 사본을 보관하지 못해' }).waitFor();
    assert.equal(await value(p), 42); assert.equal(s.reads, 0);
  });
  await test('충돌 창 표시 후 문서 전환은 이전 다시 열기 거부', async (p, s) => {
    await conflict(p, s); await switchBook(p); await reopen(p);
    await d(p).getByRole('alert').filter({ hasText: '현재 문서나 보관함이 바뀌었습니다' }).waitFor();
    assert.equal(await value(p), 84); assert.equal(s.reads, 0);
  });
  await test('복구 사본 저장 대기 중 문서 전환은 이전 요청 중단', async (p, s) => {
    await conflict(p, s);
    await p.evaluate(() => { const blob = Response.prototype.blob; let release; const gate = new Promise(r => { release = r; }); window.__releaseBlob = release; Response.prototype.blob = function() { Response.prototype.blob = blob; window.__blobHeld = true; return gate.then(() => blob.call(this)); }; });
    await reopen(p); await p.waitForFunction(() => window.__blobHeld); await switchBook(p);
    await p.evaluate(() => window.__releaseBlob());
    await d(p).getByRole('alert').filter({ hasText: '현재 문서나 보관함이 바뀌었습니다' }).waitFor();
    assert.equal(await value(p), 84); assert.equal(s.reads, 0);
    assert.ok((await records(p)).some(x => x.data.workbook.sheets[0].cells['0,0'].raw === '42'), '사본은 시작 당시의 값');
  });
  await test('온라인 응답 대기 중 문서 전환은 현재 문서 유지', async (p, s) => {
    await conflict(p, s); let release; s.holdLoad = new Promise(r => { release = r; });
    try {
      const request = p.waitForRequest(q => q.method() === 'GET' && new URL(q.url()).pathname.startsWith('/api/files/'));
      await reopen(p); await request;
      await switchBook(p); release(); await d(p).getByRole('button', { name: '온라인 문서 다시 열기', exact: true }).waitFor({ state: 'visible' });
      await p.waitForFunction(() => !document.querySelector('[data-remote-conflict] [aria-busy="true"]') && !document.querySelector('[data-remote-conflict]')?.hasAttribute('aria-busy'));
      assert.equal(await value(p), 84); assert.equal(s.reads, 1);
    } finally { release(); }
  });
} finally { await browser?.close(); }
const result = { url: base.href, assets: [...assets], engine, baseline: baseline || null, total: results.length, passed: results.filter(x => x.ok).length, failed: results.filter(x => !x.ok), results, note: '합성 문서·격리 IndexedDB. API 전부 메모리 모의, 실제 원격 쓰기 없음. Windows WebKit은 Blob IDB를 지원하는 새 persistent 프로필 사용.' };
await writeFile(out + '/result.json', JSON.stringify(result, null, 2)); console.log(JSON.stringify(result));
if (result.failed.length) process.exitCode = 1;
