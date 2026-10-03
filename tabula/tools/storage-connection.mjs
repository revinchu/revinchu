// 합성 문서와 격리된 실제 IndexedDB만 사용합니다. iPad OS 종료 재현은 아닙니다.
import assert from 'node:assert/strict';
import { mkdir, mkdtemp, writeFile } from 'node:fs/promises';
const engines = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = new URL(process.env.WIXEL_URL || 'http://127.0.0.1:5191/');
if (!['localhost', '127.0.0.1', '[::1]'].includes(url.hostname)) throw Error('로컬 소스 서버만 검사할 수 있습니다.');
const out = process.env.WIXEL_CONNECTION_OUT || 'D:/Codex/Temp/wixel-storage-connection';
await mkdir(out, { recursive: true });
const results = [];
for (const engine of (process.env.WIXEL_ENGINES || 'chromium,webkit').split(',')) {
  const profile = await mkdtemp(out + '/' + engine + '-profile-');
  const context = await engines[engine].launchPersistentContext(profile, { headless: true });
  const page = context.pages()[0] || await context.newPage(), pageErrors = [], blocked = [];
  page.on('pageerror', e => pageErrors.push(e.message));
  await context.route('**/*', route => {
    const request = route.request(), target = new URL(request.url());
    if (target.origin !== url.origin || target.pathname.startsWith('/api/') || !['GET', 'HEAD'].includes(request.method())) { blocked.push({ method: request.method(), path: target.pathname }); return route.abort(); }
    if (target.pathname === '/__storage_connection') return route.fulfill({ contentType: 'text/html', body: '<!doctype html><title>저장 연결 복구 검사</title>' });
    return route.continue();
  });
  try {
    await page.goto(new URL('/__storage_connection', url).href, { waitUntil: 'domcontentloaded' });
    const checks = await page.evaluate(async () => {
      const s = await import('/src/storage.js');
      const { saveLargeWorkbook, loadLargeWorkbook } = await import('/src/big-storage.js');
      const { Workbook } = await import('/src/workbook.js');
      const results = [], eq = (actual, expected, name) => { if (JSON.stringify(actual) !== JSON.stringify(expected)) throw Error(name + ': ' + JSON.stringify(actual)); results.push(name); };
      const originalOpen = indexedDB.open.bind(indexedDB), connections = [];
      let opens = 0, failOpen = true;
      indexedDB.open = (...args) => {
        opens++;
        if (failOpen) { failOpen = false; throw new DOMException('합성 일시 연결 실패', 'UnknownError'); }
        const request = originalOpen(...args);
        request.addEventListener('success', () => connections.push(request.result));
        return request;
      };
      let error;
      try { await s.idbSet('old', 1); } catch (e) { error = e.name; }
      eq(error, 'UnknownError', '일시 연결 실패 보고');
      await s.idbSet('old', 1); eq(await s.idbGet('old'), 1, '실패 후 다음 저장 정상'); eq(opens, 2, '실패한 open 캐시 제거');
      const old = connections.at(-1); old.close();
      await Promise.all(Array.from({ length: 12 }, (_, i) => s.idbSet('concurrent:' + i, i)));
      eq(opens, 3, '닫힌 연결 동시 요청 한 번 재연결'); eq(await s.idbGet('concurrent:11'), 11, '동시 요청 최신값 저장');
      old.dispatchEvent(new Event('close')); eq(await s.idbGet('old'), 1, '늦은 이전 close 이후 읽기'); eq(opens, 3, '늦은 이전 close가 새 연결을 버리지 않음');
      const closeEventDb = connections.at(-1); closeEventDb.close(); closeEventDb.dispatchEvent(new Event('close'));
      eq(await s.idbGet('old'), 1, 'close 이벤트 뒤 재연결'); eq(opens, 4, 'close 재연결 횟수');
      connections.at(-1).dispatchEvent(new Event('versionchange'));
      await s.idbSet('old', 2); eq(await s.idbGet('old'), 2, 'versionchange 뒤 저장'); eq(opens, 5, 'versionchange 재연결 횟수');
      const put = IDBObjectStore.prototype.put; let writes = 0;
      IDBObjectStore.prototype.put = function (value, key) { const request = put.call(this, value, key); if (key === 'old') { writes++; this.transaction.abort(); } return request; };
      error = null; try { await s.idbSet('old', 3); } catch (e) { error = e?.name || 'rejected'; } finally { IDBObjectStore.prototype.put = put; }
      eq(!!error, true, '진행 중 실제 트랜잭션 abort 보고'); eq(writes, 1, '중단된 쓰기 재시도 없음'); eq(await s.idbGet('old'), 2, '중단 이전 값 보존'); eq(opens, 5, '진행 중 abort 재연결 없음');
      let callbacks = 0;
      error = null; try { await s.idbUpdate(['old'], () => { callbacks++; throw new DOMException('합성 콜백 오류', 'InvalidStateError'); }); } catch (e) { error = e.name; }
      eq(error, 'InvalidStateError', '트랜잭션 콜백 오류 유지'); eq(callbacks, 1, '콜백 중복 실행 없음'); eq(opens, 5, '콜백 오류에 재연결 없음');
      connections.at(-1).close(); await s.idbCompareAndSet('old', 2, 4); eq(await s.idbGet('old'), 4, 'CAS 재연결 후 쓰기');
      error = null; try { await s.idbCompareAndSet('old', 2, 5); } catch (e) { error = e.code; }
      eq(error, 'IDB_CONFLICT', 'CAS 충돌 보호 유지'); eq(await s.idbGet('old'), 4, '충돌 시 현재 값 유지');
      connections.at(-1).close(); await s.idbDel('concurrent:0'); eq(await s.idbGet('concurrent:0'), null, '삭제 재연결');
      connections.at(-1).close(); const keys = await s.idbKeys('concurrent:'); eq(keys.length, 11, '키 목록 재연결');
      connections.at(-1).close(); await s.idbDeleteMany(keys); eq(await s.idbKeys('concurrent:'), [], '묶음 삭제 재연결');
      connections.at(-1).close(); eq(await s.idbUpdate(['old'], values => ({ set: [['old', values.get('old') + 1]], result: 'done' })), 'done', '원자 갱신 재연결');
      eq(await s.idbGet('old'), 5, '원자 갱신 값');

      const w = new Workbook({ sheets: [{ name: '합성 복구', cells: { '0,0': { raw: '10' } } }] });
      const key = 'synthetic:connection-book', identity = { docId: 'synthetic-doc', sessionId: 'synthetic-session', openId: 'synthetic-open', version: 0 };
      w.transact(() => w.setInput(0, 0, 0, '11'));
      const metadata = { docId: identity.docId, docName: '합성 복구', recovery: identity };
      const first = await saveLargeWorkbook(key, w, metadata);
      eq(first.manifest.recovery.version, w.version, '저장 버전은 현재 문서 버전'); eq(metadata.recovery.version, 0, '호출자 메타 불변');
      connections.at(-1).close();
      const restored = await loadLargeWorkbook(key);
      eq(restored.generation, first.manifest.generation, '복구 세대 읽기'); eq(restored.recovery, first.manifest.recovery, '복구 문서·세션·열기 식별자 읽기');
      eq(new Workbook(restored.workbook).getValue(0, 0, 0), 11, '연결 종료 후 Blob 청크 복구');
      if (!navigator.locks) throw Error('실제 Web Locks 검사가 필요합니다.');
      let release, entered; const enteredPromise = new Promise(resolve => { entered = resolve; });
      const held = navigator.locks.request('wixel-large-save:' + key, () => { entered(); return new Promise(resolve => { release = resolve; }); });
      await enteredPromise;
      const waitingVersion = w.version, queued = saveLargeWorkbook(key, w, { ...metadata, recovery: { ...identity, version: waitingVersion } });
      w.transact(() => w.setInput(0, 0, 0, '12')); release(); await held;
      const afterWait = await queued, reopened = await loadLargeWorkbook(key);
      eq(afterWait.version > waitingVersion, true, '잠금 대기 중 새 편집 발생'); eq(afterWait.manifest.recovery.version, afterWait.version, '잠금 획득 후 실제 저장 버전');
      eq(reopened.recovery.version, afterWait.version, '실제 저장 버전 복구'); eq(new Workbook(reopened.workbook).getValue(0, 0, 0), 12, '잠금 대기 중 최신값 보존');
      // Repeated JSON styles used to become one object per formula on restore.
      const styled = new Workbook(), sharedStyle = { fontFamily:'맑은 고딕', fontSize:11, fill:'#ddeeff', bold:true, borders:{left:{color:'#112233',width:2}} };
      for (let r=0;r<50000;r++) styled.sheets[0].cells.setRC(r,0,{raw:'=1+1',formula:true,cached:2,style:sharedStyle});
      const styledKey='synthetic:styles', styledSaved=await saveLargeWorkbook(styledKey,styled,{docName:'합성 서식'});
      const styledRecord=await s.idbGet(styledSaved.manifest.sheets[0].key);
      eq(styledRecord.chunks.length>1,true,'서식이 여러 JSON 청크에 걸쳐 저장됨');
      const firstText=styledRecord.gz?await new Response(styledRecord.chunks[0].stream().pipeThrough(new DecompressionStream('gzip'))).text():await styledRecord.chunks[0].text();
      const rawEntries=JSON.parse(firstText);eq(rawEntries[0][3].style===rawEntries[1][3].style,false,'원래 JSON 복원은 같은 서식을 별개 객체로 생성');
      const styledData=await loadLargeWorkbook(styledKey), styleSet=new Set();
      for(const [,cell] of styledData.workbook.sheets[0].cells)styleSet.add(cell.style);
      eq(styleSet.size,1,'복구된 5만 수식 셀 서식을 한 객체로 공유');
      const styledBook=new Workbook(styledData.workbook);eq(styledBook.getValue(0,49999,0),2,'끝 수식 저장값 유지');
      eq(styledBook.getCell(0,0,0).style,sharedStyle,'중첩 테두리·글꼴·색 유지');
      styledBook.transact(()=>styledBook.setStyle(0,0,0,{fill:'#ff0000',borders:{left:{color:'#abcdef',width:4}}}));
      eq(styledBook.getCell(0,1,0).style,sharedStyle,'한 셀 서식 편집 후 이웃 서식 보존');
      styledBook.undo();eq(styledBook.getCell(0,0,0).style,sharedStyle,'서식 실행 취소 보존');
      const oldStyleKey='synthetic:oldstyles',oldStyleRecord=oldStyleKey+'#g#old#s0';
      await s.idbSet(oldStyleRecord,{meta:{name:'옛 서식'},chunks:[new Blob([JSON.stringify([['0,0',{raw:'1',style:sharedStyle}],['1,0',{raw:'2',style:sharedStyle}]])])],gz:false,blocks:[]});
      await s.idbSet(oldStyleKey,{v:3,generation:'old',book:{},sheets:[{id:'old',ev:0,key:oldStyleRecord}]});
      const oldStyles=(await loadLargeWorkbook(oldStyleKey)).workbook.sheets[0].cells;
      eq(oldStyles.getRC(0,0).style===oldStyles.getRC(1,0).style,true,'v3 이전 셀 청크도 동일 서식 공유');
      const conflictKey='synthetic:sequential-conflict', originalBook=new Workbook({sheets:[{name:'동시 편집',cells:{'0,0':{raw:'1'}}}]});
      const base=await saveLargeWorkbook(conflictKey,originalBook,{docId:'shared-document'},{expectedGeneration:null});
      const firstTab=new Workbook((await loadLargeWorkbook(conflictKey)).workbook),staleTab=new Workbook((await loadLargeWorkbook(conflictKey)).workbook);
      firstTab.transact(()=>firstTab.setInput(0,0,0,'101'));staleTab.transact(()=>staleTab.setInput(0,0,0,'202'));
      const competing=await saveLargeWorkbook(conflictKey,firstTab,{docId:'shared-document'},{expectedGeneration:base.manifest.generation});
      eq(competing.manifest.generation!==base.manifest.generation,true,'관찰한 세대에서 정상 후속 저장');
      const keysBefore=(await s.idbKeys(conflictKey)).sort(), originalPut=IDBObjectStore.prototype.put, originalDelete=IDBObjectStore.prototype.delete;let attemptedWrites=0;
      IDBObjectStore.prototype.put=function(...args){attemptedWrites++;return originalPut.apply(this,args);};
      IDBObjectStore.prototype.delete=function(...args){attemptedWrites++;return originalDelete.apply(this,args);};
      error=null;try{await saveLargeWorkbook(conflictKey,staleTab,{docId:'shared-document'},{expectedGeneration:base.manifest.generation});}catch(e){error=e.code;}finally{IDBObjectStore.prototype.put=originalPut;IDBObjectStore.prototype.delete=originalDelete;}
      eq(error,'IDB_CONFLICT','다른 탭 완료 후 오래된 작업본 저장 차단');eq(attemptedWrites,0,'세대 충돌은 청크 저장·GC 전에 차단');
      eq((await s.idbKeys(conflictKey)).sort(),keysBefore,'경쟁 탭 저장본·청크 유지');eq(new Workbook((await loadLargeWorkbook(conflictKey)).workbook).getValue(0,0,0),101,'경쟁 탭 최신 저장값 유지');
      eq(staleTab.getValue(0,0,0),202,'충돌한 현재 메모리 편집 내용 유지');
      error=null;try{await saveLargeWorkbook(conflictKey,staleTab,{},{expectedGeneration:null});}catch(e){error=e.code;}eq(error,'IDB_CONFLICT','새 문서 예상 빈 키가 기존 세대를 덮지 않음');
      const forkKey='synthetic:sequential-fork';await saveLargeWorkbook(forkKey,staleTab,{docId:'fork-document'},{expectedGeneration:null});
      eq(new Workbook((await loadLargeWorkbook(forkKey)).workbook).getValue(0,0,0),202,'충돌 작업본 새 키에 안전하게 저장');eq(new Workbook((await loadLargeWorkbook(conflictKey)).workbook).getValue(0,0,0),101,'사본 저장 후 기존 문서 보존');
      const legacy = 'synthetic:legacy'; await s.idbSet(legacy, { v: 1, docName: '옛 형식', workbook: { sheets: [{ name: '옛 시트', cells: {} }] } });
      eq((await loadLargeWorkbook(legacy)).docName, '옛 형식', '이전 저장본 읽기 유지');
      return results;
    });
    assert.deepEqual(pageErrors, []); assert.deepEqual(blocked, []);
    results.push({ engine, ok: true, checks: checks.length, names: checks, pageErrors, blocked });
    console.log(engine + ': ' + checks.length + ' checks');
  } catch (error) { results.push({ engine, ok: false, error: error.stack, pageErrors, blocked }); }
  finally { await context.close(); }
}
const report = { cases: results.length, passed: results.filter(item => item.ok).length, checks: results.reduce((n, item) => n + (item.checks || 0), 0), results };
await writeFile(out + '/result.json', JSON.stringify(report, null, 2)); console.log(JSON.stringify(report));
assert.equal(report.passed, report.cases);
