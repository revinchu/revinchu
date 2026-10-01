// 로컬 source/compiled: 자동 복원본의 보관함 CAS 충돌을 새 사본으로 저장하는 실제 UI 검사.
// 매 시나리오마다 새 브라우저 컨텍스트·합성 문서만 사용하며 서버 쓰기를 막습니다.
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5180/');
if (!['127.0.0.1', 'localhost', '[::1]'].includes(base.hostname)) throw new Error('로컬 서버에서 실행하세요.');
// 배포 번들은 src를 노출하지 않으므로 fixture 준비/읽기용 두 모듈만 테스트 라우트로 공급합니다.
// 실제 앱은 source에서는 동일 소스를, compiled에서는 서버의 빌드된 JS를 그대로 실행합니다.
const fixtureModules = new Map(await Promise.all(['library', 'storage'].map(async n => [`/src/${n}.js`, await readFile(new URL(`../src/${n}.js`, import.meta.url), 'utf8')])));
const browser = await chromium.launch(), errors = [], writes = []; let checks = 0;
const eq = (a, b) => { assert.deepEqual(a, b); checks++; };
const snapshot = (value) => ({ app: 'wixel', docId: 'synthetic-original', docName: '합성 복원', si: 0, workbook: new Workbook({ sheets: [{ name: '합성', cells: { '0,0': { raw: String(value) } } }] }).serialize() });
const contexts = [];
async function setup() {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }); contexts.push(context);
  await context.route('**/*', (route) => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort(); }
    if (route.request().url().endsWith('/__library_recovery__')) return route.fulfill({ body: '<!doctype html><title>합성 복원 검사</title>', contentType: 'text/html' });
    const module = fixtureModules.get(new URL(route.request().url()).pathname);
    if (module) return route.fulfill({ body: module, contentType: 'application/javascript' });
    return route.continue();
  });
  const page = await context.newPage(); page.setDefaultTimeout(10000); page.on('pageerror', (e) => errors.push(e.message));
  await page.goto(new URL('/__library_recovery__', base).href);
  await page.evaluate(async ({ original, restored }) => {
    const L = await import('/src/library.js');
    await L.libSave(original.docId, original.docName, JSON.stringify(original), { version: { label: '최신 원본' } });
    localStorage.setItem('tabula.workbook.v1', JSON.stringify({ ...restored, autosave: false, remoteDoc: false }));
  }, { original: snapshot(99), restored: snapshot(10) });
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(base.href); await page.waitForFunction(() => !!window.tabula?.wb());
  eq(await page.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 10);
  await page.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 0, 0, '20')); });
  return page;
}
async function records(page) {
  return page.evaluate(async () => {
    const L = await import('/src/library.js'), S = await import('/src/storage.js'), list = await L.libList();
    return Promise.all(list.map(async e => { const d = JSON.parse(await L.unpackText(await S.idbGet(`lib:doc:${e.id}`))); return { id: e.id, value: d.workbook.sheets[0].cells['0,0'].raw, jsonId: d.docId }; }));
  });
}
try {
  // 재시작 후 첫 저장: 원본 99를 보존하고 현재 20을 새 id로 저장합니다.
  const p = await setup();
  await p.evaluate(() => window.tabula.run('versionHistory'));
  await p.getByRole('dialog', { name: '버전 기록', exact: true }).waitFor();
  const first = await records(p), copy = first.find(e => e.id !== 'synthetic-original');
  eq(first.length, 2); eq(first.find(e => e.id === 'synthetic-original').value, '99'); eq(copy.value, '20'); eq(copy.jsonId, copy.id);
  eq(await p.evaluate(() => JSON.parse(localStorage.getItem('tabula.workbook.v1')).docId), copy.id);
  eq(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 20);
  await p.keyboard.press('Escape');
  await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 0, 0, '21')); window.tabula.run('versionHistory'); });
  await p.getByRole('dialog', { name: '버전 기록', exact: true }).waitFor();
  const second = await records(p); eq(second.length, 2); eq(second.find(e => e.id === copy.id).value, '21');

  // 같은 원본 다시 열기: 현재 화면을 사본으로 보존한 뒤 최신 원본으로 전환합니다.
  const q = await setup(); await q.evaluate(() => window.tabula.run('recentFiles'));
  await q.locator('.backstage-list tr.file').filter({ has: q.getByText('합성 복원', { exact: true }) }).locator('td').nth(1).click();
  await q.waitForFunction(() => window.tabula.wb().getValue(0, 0, 0) === 99);
  const reopened = await records(q); eq(reopened.length, 2); eq(reopened.find(e => e.id !== 'synthetic-original').value, '20');
  eq(await q.getByRole('dialog', { name: '보관 문서 열기', exact: true }).count(), 0);

  // 사본 저장 자체가 실패하면 id를 바꾸거나 성공 처리하지 않으며, 재시도는 가능합니다.
  const r = await setup();
  await r.evaluate(() => {
    const put = IDBObjectStore.prototype.put;
    window.undoSyntheticFault = () => { IDBObjectStore.prototype.put = put; };
    IDBObjectStore.prototype.put = function(v, k) { if (String(k).startsWith('lib:doc:') && k !== 'lib:doc:synthetic-original') throw new DOMException('합성 용량 오류', 'QuotaExceededError'); return put.call(this, v, k); };
    window.tabula.run('versionHistory');
  });
  await r.getByRole('dialog', { name: '버전 기록', exact: true }).waitFor();
  eq((await records(r)).length, 1); eq(await r.evaluate(() => JSON.parse(localStorage.getItem('tabula.workbook.v1')).docId), 'synthetic-original');
  eq(await r.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 20);
  await r.keyboard.press('Escape'); await r.evaluate(() => { window.undoSyntheticFault(); window.tabula.run('versionHistory'); });
  await r.getByRole('dialog', { name: '버전 기록', exact: true }).waitFor();
  eq((await records(r)).some(e => e.id !== 'synthetic-original' && e.value === '20'), true);
  // 사본 압축 대기 중 새 편집은 이전 스냅샷의 id로 채택하지 않습니다.
  const s = await setup();
  await s.evaluate(() => {
    const blob = Response.prototype.blob;
    let delay = true;
    Response.prototype.blob = function() {
      const result = blob.call(this);
      if (!delay) return result;
      delay = false;
      return result.then(data => { window.syntheticPackWaiting = true; return new Promise(resolve => { window.releaseSyntheticPack = () => { Response.prototype.blob = blob; resolve(data); }; }); });
    };
    window.tabula.run('versionHistory');
  });
  await s.waitForFunction(() => window.syntheticPackWaiting);
  await s.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setInput(0, 0, 0, '30')); window.releaseSyntheticPack(); });
  await s.getByRole('dialog', { name: '버전 기록', exact: true }).waitFor();
  eq(await s.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 30);
  eq(await s.evaluate(() => JSON.parse(localStorage.getItem('tabula.workbook.v1')).docId), 'synthetic-original');
  const delayed = await records(s); eq(delayed.find(e => e.id === 'synthetic-original').value, '99'); eq(delayed.find(e => e.id !== 'synthetic-original').value, '20');
  eq(errors, []); eq(writes, []);
  console.log(JSON.stringify({ ok: true, cases: 4, checks, pageErrors: errors, blockedWrites: writes }));
} finally { for (const c of contexts) await c.close(); await browser.close(); }
