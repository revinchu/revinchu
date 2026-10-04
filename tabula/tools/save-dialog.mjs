// 실제 OS 파일 대신 메모리 File System Access API를 사용한다. 사용자 파일/보관함 쓰기 없음.
// 소스와 최종 번들 공통. native OS 창 자체의 외형/권한 동작 검증은 포함하지 않는다.
import assert from 'node:assert/strict';
import { readXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';
import { readWixelFile } from '../src/wixel-file.js';
import { parseDelimited } from '../src/csv.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const browser = await chromium.launch(), results = [];
async function test(name, fn, { picker = true } = {}) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
  const p = await context.newPage(), errors = [], remoteWrites = [], downloads = [];
  p.setDefaultTimeout(12000); p.on('pageerror', e => errors.push(e.message));
  p.on('download', d => downloads.push(d));
  await context.route('**/*', route => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { remoteWrites.push(route.request().url()); return route.abort(); }
    return route.continue();
  });
  await p.addInitScript(({ picker }) => {
    window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
    try { localStorage.setItem('wixel.options', JSON.stringify({ saveAsk: false, saveConfirm: false, qatPosition: 'below', qatOrder: ['save', 'undo', 'redo'] })); } catch {}
    const mock = window.__saveMock = { calls: [], events: [], plans: [], completed: [] };
    const makeHandle = (id, name, plan = {}) => ({
      kind: 'file', name, mockId: id,
      async createWritable() {
        mock.events.push({ id, stage: 'create' });
        const parts = []; let writes = 0, position = 0, extent = 0;
        return {
          async write(blob) {
            mock.events.push({ id, stage: 'write' });
            writes++;
            if (plan.pauseWrite && writes === 1) await new Promise(resolve => { mock.releaseWrite = resolve; });
            if (plan.changeAfterWrite === writes) window.tabula.wb().transact(() => window.tabula.wb().setInput(0, 0, 0, 'changed during save'));
            if (plan.writeError) throw new DOMException('합성 쓰기 권한 오류', 'NotAllowedError');
            const bytes = blob instanceof Blob ? new Uint8Array(await blob.arrayBuffer()) : new Uint8Array(blob.buffer ?? blob, blob.byteOffset ?? 0, blob.byteLength).slice();
            parts.push({position,bytes}); position += bytes.length; extent = Math.max(extent,position);
          },
          async seek(at) { if (plan.seekError) throw new DOMException('위치 이동 실패','InvalidStateError'); position=at; mock.events.push({id,stage:'seek'}); },
          async close() { if (plan.pauseClose) await new Promise(resolve => { mock.releaseClose = resolve; }); if (plan.closeError) throw new DOMException('디스크 공간 부족', 'QuotaExceededError'); const bytes=new Uint8Array(extent); for(const part of parts)bytes.set(part.bytes,part.position); mock.completed.push({ id, name, bytes:Array.from(bytes) }); mock.events.push({ id, stage: 'close' }); },
          async abort() { mock.events.push({ id, stage: 'abort' }); },
        };
      },
    });
    mock.unselected = makeHandle('unselected', '선택하지 않은 파일.xlsx');
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: picker ? async options => {
      const call = mock.calls.length + 1, plan = mock.plans.shift() || {};
      mock.calls.push({ suggestedName: options.suggestedName, id: options.id, startIn: options.startIn?.mockId ?? options.startIn, types: options.types, userActive: navigator.userActivation?.isActive });
      if (plan.error) throw new DOMException('합성 파일 선택 오류', plan.error);
      return makeHandle(plan.id || `selected-${call}`, plan.name || `선택한 파일 ${call}.xlsx`, plan);
    } : undefined });
  }, { picker });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    if (await p.locator('#autosaveToggle').getAttribute('aria-checked') === 'true') await p.locator('#autosaveToggle').click();
    await p.evaluate(() => { const t = window.tabula; t.wb().transact(() => t.wb().setInput(0, 0, 0, '42')); t.selectCell(0, 0); });
    await fn(p, downloads);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(remoteWrites, [], '원격 쓰기 요청');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors, remoteWrites }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
const ctrlSave = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+s'); };
const settle = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 50)))));
const closed = (p, count) => p.waitForFunction(count => window.__saveMock.events.filter(e => e.stage === 'close').length === count, count);
const mock = p => p.evaluate(() => window.__saveMock);
const valueOf = bytes => new Workbook(readXlsx(Uint8Array.from(bytes)).data).getValue(0, 0, 0);
try {

  for (const kind of ['xlsx','wixel','csv']) await test(`저장 중 취소는 ${kind} 파일을 폐기하고 다시 저장 가능`, async (p, downloads) => {
    const title = await p.locator('#docTitle').innerText();
    await p.evaluate(kind => { window.__saveMock.plans.push({name:'취소 검사.'+kind,pauseWrite:true}); }, kind);
    if (kind === 'xlsx') await ctrlSave(p);
    else if (kind === 'csv') await p.evaluate(() => window.tabula.run('exportCsv'));
    else {
      await p.evaluate(() => window.tabula.run('saveAs'));
      const d=p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true});
      await d.locator('select').selectOption('wixel'); await d.getByRole('button',{name:'저장',exact:true}).click();
    }
    await p.waitForFunction(() => typeof window.__saveMock.releaseWrite === 'function');
    await p.getByRole('button',{name:'저장 취소',exact:true}).click();
    await p.evaluate(() => window.__saveMock.releaseWrite());
    await p.waitForFunction(() => !document.querySelector('.load-progress'));
    let m=await mock(p); assert.equal(m.completed.length,0); assert.equal(m.events.filter(e=>e.stage==='abort').length,1);
    assert.equal(await p.locator('#docTitle').innerText(),title);
    assert.equal(await p.getByRole('dialog').count(),0); assert.equal(downloads.length,0);
    await ctrlSave(p); await closed(p,1);
    m=await mock(p); assert.equal(valueOf(m.completed[0].bytes),42);
  });
  await test('디스크 확정 단계에는 취소를 비활성화하고 실제 완료 후 성공 표시', async (p, downloads) => {
    await p.evaluate(() => window.__saveMock.plans.push({pauseClose:true}));
    await ctrlSave(p); await p.waitForFunction(() => typeof window.__saveMock.releaseClose === 'function');
    assert.equal(await p.getByRole('button',{name:'저장 취소',exact:true}).isDisabled(),true);
    assert.equal((await mock(p)).completed.length,0);
    await p.evaluate(() => window.__saveMock.releaseClose()); await closed(p,1);
    assert.equal(downloads.length,0);
  });

  await test('자동 저장은 셀 변경을 브라우저에 보관하고 파일 선택기를 호출하지 않음', async (p, downloads) => {
    await p.locator('#autosaveToggle').click();
    assert.equal(await p.locator('#autosaveToggle').getAttribute('aria-checked'), 'true');
    await p.locator('#cellEditor').focus(); await p.keyboard.type('94'); await p.keyboard.press('Enter');
    await p.waitForFunction(() => {
      const saved = JSON.parse(localStorage.getItem('tabula.workbook.v1') || 'null');
      return saved?.workbook?.sheets?.[0]?.cells?.['0,0']?.raw === '94';
    });
    await p.waitForFunction(() => document.getElementById('saveState').textContent.includes('이 브라우저에 저장됨'));
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 94);
    const m = await mock(p); assert.equal(m.calls.length, 0); assert.deepEqual(m.events, []); assert.equal(downloads.length, 0);
  });
  await test('Ctrl+S 반복: 예전 확인 설정 false여도 매번 picker 1회·선택한 두 파일에 최신 값 저장', async (p, downloads) => {
    await ctrlSave(p); await closed(p, 1);
    await p.evaluate(() => window.tabula.wb().transact(() => window.tabula.wb().setInput(0, 0, 0, '73')));
    await ctrlSave(p); await closed(p, 2);
    const m = await mock(p); assert.equal(m.calls.length, 2);
    assert.ok(m.calls.every(call => call.userActive), 'Ctrl+S 사용자 활성화가 유지된 동안 파일 선택기 호출');
    assert.deepEqual(m.completed.map(x => x.id), ['selected-1', 'selected-2']);
    assert.deepEqual(m.completed.map(x => valueOf(x.bytes)), [42, 73]);
    assert.deepEqual(m.events.filter(e=>!['write','seek'].includes(e.stage)).map(e=>[e.id,e.stage]), [['selected-1','create'],['selected-1','close'],['selected-2','create'],['selected-2','close']]);
    for(const id of ['selected-1','selected-2'])assert.ok(m.events.some(e=>e.id===id&&e.stage==='write'));
    assert.equal(downloads.length, 0);
  });
  await test('공개 exportXlsx의 이전 target 인수와 저장 버튼 모두 기존 handle 즉시 쓰기 우회 불가', async (p, downloads) => {
    await p.evaluate(() => window.tabula.exportXlsx('우회 검사', 'xlsx', window.__saveMock.unselected)); await closed(p, 1);
    await p.locator('[data-cmd="save"]:visible').first().click(); await closed(p, 2);
    const m = await mock(p); assert.equal(m.calls.length, 2); assert.ok(m.events.every(e => e.id !== 'unselected'));
    assert.deepEqual(m.completed.map(x => x.id), ['selected-1', 'selected-2']); assert.equal(downloads.length, 0);
  });
  await test('셀 편집 중 Ctrl+S는 편집값을 확정하고 파일 선택기로 저장', async (p, downloads) => {
    await p.locator('#cellEditor').focus(); await p.keyboard.press('F2'); await p.locator('#cellEditor').fill('81');
    await p.keyboard.press('Control+s'); await closed(p, 1);
    const m = await mock(p); assert.equal(m.calls.length, 1); assert.equal(valueOf(m.completed[0].bytes), 81);
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 81); assert.equal(downloads.length, 0);
  });
  await test('수식 입력줄 편집 중 Ctrl+S는 수식·계산값을 확정해 파일에 저장', async (p, downloads) => {
    await p.locator('#formulaInput').fill('=SUM(2,3)'); await p.keyboard.press('Control+s'); await closed(p, 1);
    const m = await mock(p), book = new Workbook(readXlsx(Uint8Array.from(m.completed[0].bytes)).data);
    assert.equal(m.calls.length, 1); assert.equal(book.getRaw(0, 0, 0), '=SUM(2,3)'); assert.equal(book.getValue(0, 0, 0), 5); assert.equal(downloads.length, 0);
  });
  await test('기존 파일 handle이 있어도 다음 picker 취소는 쓰기·다운로드 0', async (p, downloads) => {
    await ctrlSave(p); await closed(p, 1); const before = await mock(p);
    await p.evaluate(() => window.__saveMock.plans.push({ error: 'AbortError' }));
    await ctrlSave(p); await p.waitForFunction(() => window.__saveMock.calls.length === 2); await settle(p);
    const after = await mock(p); assert.deepEqual(after.events, before.events); assert.deepEqual(after.completed, before.completed); assert.equal(downloads.length, 0);
  });
  await test('picker 권한 오류는 자동 다운로드·다른 파일 쓰기로 대체하지 않음', async (p, downloads) => {
    await p.evaluate(() => window.__saveMock.plans.push({ error: 'SecurityError' }));
    await ctrlSave(p); await p.waitForFunction(() => window.__saveMock.calls.length === 1); await settle(p);
    assert.deepEqual((await mock(p)).events, []); assert.equal(downloads.length, 0);
    assert.equal(await p.getByRole('dialog', { name: '파일로 저장', exact: true }).count(), 0);
    assert.match(await p.locator('body').innerText(), /(?:저장.*(?:못|오류|실패)|파일 선택.*(?:못|오류|실패))/);
  });
  await test('선택 파일 쓰기 실패도 다운로드 없이 오류를 표시', async (p, downloads) => {
    await p.evaluate(() => window.__saveMock.plans.push({ writeError: true }));
    await ctrlSave(p); await p.getByRole('dialog').filter({ hasText: '저장하지 못했습니다' }).waitFor();
    const m = await mock(p); assert.equal(m.calls.length, 1); assert.equal(m.completed.length, 0); assert.equal(downloads.length, 0);
  });
  await test('스트리밍 중 문서 변경·디스크 마무리 실패는 abort하고 성공 파일을 남기지 않음', async (p, downloads) => {
    for (const plan of [{changeAfterWrite:1}, {closeError:true}]) {
      await p.evaluate(plan => window.__saveMock.plans.push(plan), plan);
      await ctrlSave(p); const dialog = p.getByRole('dialog').filter({hasText:'저장하지 못했습니다'}); await dialog.waitFor();
      await dialog.getByRole('button', {name:/확인/}).click();
    }
    const m = await mock(p); assert.equal(m.completed.length,0); assert.equal(m.events.filter(e=>e.stage==='abort').length,2);
    assert.equal(downloads.length,0);
  });
  await test('WIXEL은 picker 직접 저장 후 값·서식 재열기 가능', async (p, downloads) => {
    await p.evaluate(() => { window.__saveMock.plans.push({name:'보관.wixel'}); window.tabula.wb().setStyle(0,0,0,{bold:true,fill:'#123456'}); window.tabula.run('saveAs'); });
    const d = p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true});
    await d.locator('select').selectOption('wixel'); await d.getByRole('button',{name:'저장',exact:true}).click(); await closed(p,1);
    const m=await mock(p), result=await readWixelFile(new Blob([Uint8Array.from(m.completed[0].bytes)]));
    const restored=new Workbook(result.workbook);
    assert.equal(restored.getValue(0,0,0),42); assert.equal(restored.styleAt(0,0,0).bold,true); assert.equal(restored.styleAt(0,0,0).fill,'#123456');
    assert.equal(downloads.length,0);
  });
  await test('ODS 비보존 개체가 있으면 저장 범위를 알리고 취소는 파일을 쓰지 않음', async (p, downloads) => {
    await p.evaluate(() => { window.tabula.wb().sheets[0].charts.push({id:'chart-save-test',type:'column',range:{r1:0,c1:0,r2:0,c2:0},x:0,y:0,w:200,h:100}); window.tabula.run('saveAs'); });
    const d=p.getByRole('dialog',{name:'다른 이름으로 저장',exact:true}); await d.locator('select').selectOption('ods'); await d.getByRole('button',{name:'저장',exact:true}).click();
    const warning=p.getByRole('dialog',{name:'OpenDocument 저장 범위',exact:true}); await warning.waitFor(); assert.match(await warning.innerText(),/피벗.*포함하지 않습니다/);
    await warning.getByRole('button',{name:'취소',exact:true}).click(); await settle(p);
    assert.equal((await mock(p)).calls.length,0); assert.equal(downloads.length,0);
  });
  await test('CSV는 여러 번 나눠 디스크에 저장하고 전체 행·따옴표·한글 복원', async (p, downloads) => {
    await p.evaluate(() => {
      window.__saveMock.plans.push({name:'큰 값.csv'});
      const book=window.tabula.wb();book.transact(()=>{for(let r=0;r<4500;r++){book.setInput(0,r,0,'한글😀,'.repeat(70));book.setInput(0,r,1,String(r));}});
      window.tabula.run('exportCsv');
    });
    await closed(p,1); const m=await mock(p), bytes=Uint8Array.from(m.completed[0].bytes), rows=parseDelimited(new TextDecoder().decode(bytes));
    assert.equal(rows.length,4500); assert.equal(rows[0][0],'한글😀,'.repeat(70)); assert.equal(rows[4499][1],'4499');
    assert.ok(m.events.filter(e=>e.stage==='write').length>1); assert.equal(downloads.length,0);
  });
  await test('다른 이름으로 저장의 native picker 취소는 문서 이름·내용을 그대로 유지', async (p, downloads) => {
    const title = await p.locator('#docTitle').innerText();
    await p.evaluate(() => { window.__saveMock.plans.push({ error: 'AbortError' }); window.tabula.run('saveAs'); });
    const d = p.getByRole('dialog', { name: '다른 이름으로 저장', exact: true });
    await d.getByLabel('파일 이름', { exact: true }).fill('취소한 새 문서 이름');
    await d.getByRole('button', { name: '저장', exact: true }).click();
    await p.waitForFunction(() => window.__saveMock.calls.length === 1); await settle(p);
    assert.equal(await p.locator('#docTitle').innerText(), title); assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 42);
    assert.deepEqual((await mock(p)).events, []); assert.equal(downloads.length, 0);
  });
  await test('파일 화면의 Excel 다운로드·공유 내보내기도 picker를 매번 실행', async (p, downloads) => {
    await p.evaluate(() => window.tabula.run('saveLocations'));
    await p.getByRole('button', { name: /Excel 파일 다운로드/ }).click(); await closed(p, 1);
    await p.locator('[data-page="share"]').click();
    await p.getByRole('button', { name: /^Excel 통합 문서/ }).click(); await closed(p, 2);
    assert.equal((await mock(p)).calls.length, 2); assert.equal(downloads.length, 0);
  });
  await test('picker 미지원: 파일 이름·폴더 제한 안내 후 명시 다운로드 버튼으로만 저장', async (p, downloads) => {
    await ctrlSave(p); const d = p.getByRole('dialog', { name: '파일로 저장', exact: true }); await d.waitFor();
    assert.match(await d.innerText(), /폴더/); assert.equal(downloads.length, 0);
    await d.getByLabel('파일 이름', { exact: true }).fill('검사 결과.xlsx');
    const event = p.waitForEvent('download'); await d.getByRole('button', { name: '다운로드', exact: true }).click();
    const download = await event; assert.equal(download.suggestedFilename(), '검사 결과.xlsx');
    const stream = await download.createReadStream(), chunks = []; for await (const chunk of stream) chunks.push(chunk);
    assert.equal(valueOf(Buffer.concat(chunks)), 42); assert.deepEqual((await mock(p)).events, []);
    await d.waitFor({ state: 'hidden' }); assert.equal(downloads.length, 1);
  }, { picker: false });
  await test('picker 미지원 대체 저장 창 취소는 다운로드·문서 이름 변경 0', async (p, downloads) => {
    const title = await p.locator('#docTitle').innerText(); await ctrlSave(p);
    const d = p.getByRole('dialog', { name: '파일로 저장', exact: true }); await d.getByLabel('파일 이름', { exact: true }).fill('취소.xlsx');
    await d.getByRole('button', { name: '취소', exact: true }).click(); await d.waitFor({ state: 'hidden' }); await settle(p);
    assert.equal(downloads.length, 0); assert.equal(await p.locator('#docTitle').innerText(), title);
  }, { picker: false });
} finally { await browser.close(); }
console.log(JSON.stringify({ url, total: results.length, ok: results.filter(x => x.ok).length, bad: results.filter(x => !x.ok), note: 'File System Access API 가상 파일 검사. 실제 OS 선택창 UI·권한은 별도 수동 확인 필요.' }, null, 2));
if (results.some(x => !x.ok)) process.exitCode = 1;
