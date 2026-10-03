// 새 격리 브라우저 프로필의 합성 IndexedDB만 사용합니다. source 서버 전용, 원격 쓰기 없음.
// WIXEL_URL=http://127.0.0.1:5180/, PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH.
import assert from 'node:assert/strict';
import { mkdir, mkdtemp } from 'node:fs/promises';
const engines = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium';
const base = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5180/');
if (!['127.0.0.1', 'localhost', '[::1]'].includes(base.hostname)) throw new Error('이 합성 IDB 검사는 로컬 source 서버에서만 실행하세요.');
// Windows WebKit의 ephemeral context는 앱과 무관하게 IDB Blob put이 실패합니다.
// 실제 Blob 저장을 검사하려고 새 D: 전용 persistent profile을 씁니다(사용자 프로필 아님).
let browser, context;
if (engine === 'webkit' && process.platform === 'win32') {
  const root = process.env.WIXEL_LIBRARY_PROFILE_ROOT || 'D:/Codex/Temp/wixel-improvement';
  await mkdir(root, {recursive:true});
  context = await engines[engine].launchPersistentContext(await mkdtemp(root + '/webkit-library-'));
} else { browser = await engines[engine].launch(); context = await browser.newContext(); }
let checks = 0;
const eq = (a, b, msg) => { assert.deepEqual(a, b, msg); checks++; };
try {
  await context.route('**/__library_test__', (route) => route.fulfill({ body: '<!doctype html><title>합성 보관함 검사</title>', contentType: 'text/html' }));
  const open = async () => {
    const p = await context.newPage();
    await p.goto(new URL('/__library_test__', base).href);
    await p.evaluate(async () => { window.L = await import('/src/library.js'); window.S = await import('/src/storage.js'); window.W = await import('/src/workbook.js'); window.B = await import('/src/snapshot-blob.js'); });
    return p;
  };
  const p = await open(), q = await open();
  const save = (page, id, n, options = {}) => page.evaluate(async ({ id, n, options }) => L.libSave(id, id, JSON.stringify({ value: n }), options), { id, n, options });
  const load = (page, id) => page.evaluate((id) => L.libLoad(id), id);
  const keys = () => p.evaluate(() => S.idbKeys('lib:'));
  const fail = (phase, action) => p.evaluate(async ({ phase, action }) => {
    const put = IDBObjectStore.prototype.put, del = IDBObjectStore.prototype.delete;
    IDBObjectStore.prototype.put = function (v, k) {
      if (phase === 'index' && k === 'lib:index') throw new DOMException('합성 용량 오류', 'QuotaExceededError');
      if (phase === 'version' && String(k).startsWith('lib:ver:')) throw new DOMException('합성 버전 오류', 'QuotaExceededError');
      const req = put.call(this, v, k);
      if (phase === 'abort' && k === 'lib:index') req.onsuccess = () => this.transaction.abort();
      return req;
    };
    IDBObjectStore.prototype.delete = function (k) {
      if (phase === 'delete') throw new DOMException('합성 삭제 오류', 'UnknownError');
      return del.call(this, k);
    };
    try {
      if (action === 'remove') await L.libRemove('atomic');
      else if (action === 'update') await L.libUpdate('atomic', { name: '덮어쓴 이름', pinned: true });
      else if (action === 'label') await L.libNameVersion('atomic', (await L.libList()).find(e => e.id === 'atomic').versions[0].ts, '덮어쓴 버전 이름');
      else await L.libSave('atomic', '덮어쓴 이름', JSON.stringify({ value: 2 }), { version: { label: '새 버전' } });
      return 'unexpected-success';
    } catch (e) { return e.name; }
    finally { IDBObjectStore.prototype.put = put; IDBObjectStore.prototype.delete = del; }
  }, { phase, action });

  const original = await save(p, 'atomic', 1, { version: { label: '원본' } });
  const originalKeys = await keys();
  for (const [phase, action, expected] of [['index', 'save', 'QuotaExceededError'], ['version', 'save', 'QuotaExceededError'], ['abort', 'save', 'AbortError'], ['delete', 'remove', 'UnknownError'], ['index', 'remove', 'QuotaExceededError'], ['index', 'update', 'QuotaExceededError'], ['index', 'label', 'QuotaExceededError']]) {
    eq(await fail(phase, action), expected, `${phase}/${action}`);
    eq(await load(p, 'atomic'), { value: 1 }, '실패 후 본문 보존');
    eq(await p.evaluate(() => L.libList()), [original], '실패 후 목록·버전 보존');
    eq(await keys(), originalKeys, '실패 후 고아 버전/누락 없음');
  }
  // 같은 밀리초의 버전도 각각의 내용을 유지합니다.
  const versions = await p.evaluate(async () => {
    const now = Date.now; Date.now = () => 1700000000000;
    try { await L.libSave('same-ms', '같은 시각', '{"value":1}', { version: {} }); return await L.libSave('same-ms', '같은 시각', '{"value":2}', { version: {} }); }
    finally { Date.now = now; }
  });
  eq(versions.versions.length, 2);
  eq(new Set(versions.versions.map(v => v.ts)).size, 2);
  eq(await p.evaluate(async (vs) => Promise.all(vs.map(v => L.libLoadVersion('same-ms', v.ts))), versions.versions), [{ value: 1 }, { value: 2 }]);
  // 다른 탭의 목록을 읽은 뒤 쓰더라도 단일 readwrite tx가 누락 없이 직렬화합니다.
  await Promise.all([save(p, 'tab-a', 10), save(q, 'tab-b', 20)]);
  eq((await p.evaluate(() => L.libList())).map(e => e.id).sort(), ['atomic', 'same-ms', 'tab-a', 'tab-b']);
  eq(await load(p, 'tab-a'), { value: 10 }); eq(await load(q, 'tab-b'), { value: 20 });
  await Promise.all([p.evaluate(() => L.libUpdate('tab-a', { name: '새 이름' })), q.evaluate(() => L.libUpdate('tab-b', { pinned: true }))]);
  eq(await p.evaluate(() => L.libList().then(xs => [xs.find(x => x.id === 'tab-a').name, xs.find(x => x.id === 'tab-b').pinned])), ['새 이름', true]);
  // 재시작 후처럼 revision 관측이 없는 첫 저장도 기존 본문과 다르면 거부합니다.
  await save(p, 'first-restored', 50);
  eq(await q.evaluate(async () => { try { await L.libSave('first-restored', 'first-restored', '{"value":51}'); return null; } catch(e) { return e.code; } }), 'LIB_CONFLICT');
  eq(await load(p, 'first-restored'), { value: 50 });
  await save(q, 'first-restored', 50); // 동일한 JSON 사본만 현재 revision을 채택
  await save(q, 'first-restored', 51);
  eq(await load(q, 'first-restored'), { value: 51 });
  // 두 탭의 최초 생성이 겹쳐도 첫 commit 이후의 본문을 덮어쓰지 않습니다.
  const first = await Promise.allSettled([save(p, 'same-new-id', 60), save(q, 'same-new-id', 70)]);
  eq(first.filter(x => x.status === 'fulfilled').length, 1);
  eq(first.filter(x => x.status === 'rejected').length, 1);
  eq(await load(p, 'same-new-id'), { value: first[0].status === 'fulfilled' ? 60 : 70 });
  // 옛 revision 없는 레코드와 삭제된 문서도 구분합니다.
  await p.evaluate(async () => {
    const list = await L.libList(); list.push({ id: 'legacy', name: '옛 문서', updated: Date.now(), versions: [] });
    await S.idbSet('lib:doc:legacy', await L.packText('{"value":90}')); await S.idbSet('lib:index', list);
  });
  eq(await load(q, 'legacy'), { value: 90 });
  await p.evaluate(() => L.libRemove('legacy'));
  eq(await q.evaluate(async () => { try { await L.libSave('legacy', 'legacy', '{"value":91}'); return null; } catch(e) { return e.code; } }), 'LIB_CONFLICT');
  // 같은 문서를 읽은 두 탭: 뒤늦은 저장·삭제가 최신 본문을 지우지 않습니다.
  await load(p, 'atomic'); await load(q, 'atomic'); await save(q, 'atomic', 3);
  eq(await p.evaluate(async () => { try { await L.libSave('atomic', 'atomic', '{"value":4}'); return null; } catch(e) { return e.code; } }), 'LIB_CONFLICT');
  eq(await p.evaluate(async () => { try { await L.libRemove('atomic'); return null; } catch(e) { return e.code; } }), 'LIB_CONFLICT');
  eq(await load(q, 'atomic'), { value: 3 });
  await load(p, 'atomic'); await save(p, 'atomic', 4); eq(await load(p, 'atomic'), { value: 4 });
  // 성공한 삭제는 본문과 모든 버전을 함께 제거합니다.
  await p.evaluate(() => L.libRemove('atomic'));
  eq(await load(p, 'atomic'), null);
  eq((await keys()).some(k => k === 'lib:doc:atomic' || k.startsWith('lib:ver:atomic:')), false);
  // 이름 붙인 버전과 고정 문서는 정리에서 유지합니다.
  for (let i = 0; i < 22; i++) {
    const e = await save(p, 'history', i, { version: { label: i === 0 ? '일반 저장 라벨' : '' }, max: 2 });
    if (i === 0) { eq(e.versions[0].named, false); await p.evaluate((ts) => L.libNameVersion('history', ts, '유지'), e.versions[0].ts); }
  }
  const list = await p.evaluate(() => L.libList()), history = list.find(e => e.id === 'history');
  eq(history.versions.length, 20); eq(history.versions[0].label, '유지');
  eq(list.map(e => e.id).sort(), ['history', 'tab-b']);
  eq(await p.evaluate((ts) => L.libLoadVersion('history', ts), history.versions[0].ts), { value: 0 });
  eq((await keys()).filter(k => k.startsWith('lib:ver:history:')).length, 20);
  // 과거본만 처음 연 탭도 이후의 다른 탭 저장을 덮어쓸 수 없습니다.
  eq(await q.evaluate((ts) => L.libLoadVersion('history', ts), history.versions[0].ts), { value: 0 });
  await save(p, 'history', 99, { max: 2 });
  eq(await q.evaluate(async () => { try { await L.libSave('history', 'history', '{"value":0}'); return null; } catch(e) { return e.code; } }), 'LIB_CONFLICT');
  eq(await load(p, 'history'), { value: 99 });
  // 충돌 사본 저장으로 최신 원본이 정리 대상이 되지 않습니다.
  await save(p, 'keep-original', 1, { max: 99 });
  await save(p, 'keep-copy', 2, { max: 1, keepIds: ['keep-original'] });
  eq(await load(p, 'keep-original'), { value: 1 });
  eq(await load(p, 'keep-copy'), { value: 2 });
  // 새 경로의 불변 Blob·옛 JSON 블록·첫 저장 revision·실패 복구를 함께 검사합니다.
  const blobSource = { app:'wixel',docId:'blob-current',docName:'불변 스냅숏',si:0,
    workbook:{date1904:true,sheets:[{name:'원본',cells:{'0,0':{raw:'😀한글'}},blocks:[{r0:2,c0:0,n:3,ver:0,
      cols:[{num:{0:0,1:null,2:5},str:null,dict:[],fmt:null}]}]}]} };
  const savedBlob = await p.evaluate(async data => {
    const w=new W.Workbook(data.workbook), captured=B.librarySnapshot(w,{app:data.app,docId:data.docId,docName:data.docName,si:data.si});
    window.__blobSnapshot=captured;
    w.transact(()=>w.setInput(0,0,0,'스냅숏 뒤 편집'));
    return L.libSave(data.docId,data.docName,captured.blob,{version:{label:'원본'},max:99});
  },blobSource);
  const loaded=await load(p,'blob-current');
  eq(loaded.workbook.sheets[0].cells['0,0'].raw,'😀한글');
  eq(await p.evaluate(async()=>{const d=await L.libLoad('blob-current'),w=new W.Workbook(d.workbook);return [w.getValue(0,2,0),w.getValue(0,3,0),w.getValue(0,4,0),w.date1904];}),[0,null,5,true]);
  eq(await p.evaluate(ts=>L.libLoadVersion('blob-current',ts),savedBlob.versions[0].ts),loaded);
  // 새 탭은 같은 byte의 Blob만 revision 관측 없이 채택한다.
  eq(await q.evaluate(async data=>{try{await L.libSave('blob-current','동일',new Blob([JSON.stringify(data)]),{max:99});return 'ok';}catch(e){return e.code;}},loaded),'ok');
  const fresh=await open();
  eq(await fresh.evaluate(async data=>{data.workbook.sheets[0].cells['0,0'].raw='다른 내용';try{await L.libSave('blob-current','상이',new Blob([JSON.stringify(data)]));return 'unexpected';}catch(e){return e.code;}},loaded),'LIB_CONFLICT');
  eq(await load(q,'blob-current'),loaded);
  await load(p,'blob-current');
  eq(await p.evaluate(async()=>{const put=IDBObjectStore.prototype.put;IDBObjectStore.prototype.put=function(v,k){if(k==='lib:index')throw new DOMException('합성 quota','QuotaExceededError');return put.call(this,v,k);};try{await L.libSave('blob-current','실패',new Blob(['{"bad":1}']),{version:{}});return 'unexpected';}catch(e){return e.name;}finally{IDBObjectStore.prototype.put=put;}}),'QuotaExceededError');
  eq(await load(p,'blob-current'),loaded);
  eq(await p.evaluate(async()=>{await L.libSave('blob-copy','사본',window.__blobSnapshot.withDocId('blob-copy'),{max:99});return (await L.libLoad('blob-copy')).docId;}),'blob-copy');
  eq(await p.evaluate(async()=>{const original=globalThis.CompressionStream;globalThis.CompressionStream=undefined;try{await L.libSave('plain-blob','압축 없음',new Blob(['{"text":"한글😀"}']),{max:99});return [await L.libLoad('plain-blob'),(await S.idbGet('lib:doc:plain-blob')).gz];}finally{globalThis.CompressionStream=original;}}),[{text:'한글😀'},false]);
  await fresh.close();
  console.log(JSON.stringify({ ok: true, engine, checks, contexts: 1, tabs: 3, syntheticOnly: true }));
} finally { await context.close(); await browser?.close(); }
