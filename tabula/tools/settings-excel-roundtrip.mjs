import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { excelHash } from '../src/protect.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
function fixture({ protection = '', manual = true } = {}) {
  const files = unzip(writeXlsx(new Workbook({ sheets: [
    { name: '보기설정', cells: { '0,0': { raw: '3' }, '0,1': { raw: '=A1*2', cached: 6 } } },
    { name: '기본보기', cells: { '0,0': { raw: '다음 시트' } } },
  ] })));
  files['xl/workbook.xml'] = textOf(files['xl/workbook.xml']).replace(/<calcPr[^>]*\/>/, `<calcPr calcMode="${manual ? 'manual' : 'auto'}" calcOnSave="0"/>`).replace('<bookViews>', `${protection}<bookViews>`);
  files['xl/worksheets/sheet1.xml'] = textOf(files['xl/worksheets/sheet1.xml']).replace(/<sheetView\b[^>]*>/, '<sheetView workbookViewId="0" showRowColHeaders="0" showFormulas="1" showGridLines="0" defaultGridColor="0" colorId="10">');
  if (manual === null) files['xl/workbook.xml'] = files['xl/workbook.xml'].replace(/<calcPr[^>]*\/>/, '');
  return Buffer.from(zip(files));
}
const browser = await chromium.launch(), results = [];
let checks = 0;
const equal = (actual, expected, label) => { assert.deepEqual(actual, expected, label); checks++; };
const button = (d, label) => d.getByRole('button', { name: new RegExp(`^${label}(?:\\s*\\([A-Z]\\))?$`) }).last();
async function test(name, fn) {
  const context = await browser.newContext({ viewport: { width: 1400, height: 1000 } });
  const errors = [], blocked = [];
  await context.route('**/*', route => {
    const req = route.request();
    if (req.url().startsWith(new URL(url).origin) && ['GET', 'HEAD'].includes(req.method())) return route.continue();
    if (/^(blob|data):/.test(req.url())) return route.continue();
    blocked.push(`${req.method()} ${req.url()}`); return route.abort();
  });
  const p = await context.newPage(); p.setDefaultTimeout(12000); p.on('pageerror', e => errors.push(e.message));
  try {
    await p.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel.options', JSON.stringify({ calcMode: 'manual', gridColor: '#0000ff' })); window.showSaveFilePicker = async opts => ({ name: opts.suggestedName, createWritable: async () => ({ write: async blob => { window.__savedBytes = [...new Uint8Array(await blob.arrayBuffer())]; }, close: async () => {} }) }); });
    await p.goto(url); await p.waitForFunction(() => window.tabula?.wb());
    await fn(p); equal(errors, [], '페이지 오류'); equal(blocked, [], '외부 요청'); results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false, error: e.stack, errors, blocked }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await context.close(); }
}
async function load(p, bytes = fixture(), name = 'native-settings.xlsx') {
  await p.locator('#fileInput').setInputFiles({ name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes });
  await p.waitForFunction(() => window.tabula.wb().sheets[0].name === '보기설정');
  await p.waitForFunction(() => !document.querySelector('.progress-overlay'));
}
const view = p => p.evaluate(() => { const t = window.tabula, s = t.gv().host.state(); return { headers: s.showHeaders, formulas: s.showFormulas, grid: s.showGrid, color: document.documentElement.style.getPropertyValue('--grid-line'), manual: t.wb().manualCalc }; });
const openOptions = async p => { await p.evaluate(() => window.tabula.run('options')); return p.getByRole('dialog', { name: 'WIXEL 옵션', exact: true }); };
const save = async p => { await p.evaluate(() => window.tabula.exportXlsx('settings-result')); return readXlsx(Uint8Array.from(await p.evaluate(() => window.__savedBytes))).data; };
try {
  await test('표준 Excel 보기 속성→화면, 시트 전환, 문서 전환', async p => {
    await load(p); equal(await view(p), { headers: false, formulas: true, grid: false, color: '#ff0000', manual: true });
    await p.evaluate(() => window.tabula.switchSheet(1)); equal(await view(p), { headers: true, formulas: false, grid: true, color: '', manual: true });
    await p.evaluate(() => window.tabula.switchSheet(0)); equal((await view(p)).headers, false);
    await load(p, fixture({ manual: false }), 'auto.xlsx'); await p.waitForFunction(() => !window.tabula.wb().manualCalc); equal((await view(p)).manual, false, '기기 수동 옵션이 파일 자동 설정을 덮지 않음');
    const d = await openOptions(p); await d.getByRole('tab', { name: '수식', exact: true }).click(); equal(await d.getByRole('radio', { name: '자동', exact: true }).isChecked(), true); await button(d, '취소').click();
    await load(p, fixture({ manual: null }), 'no-calc-property.xlsx'); await p.waitForFunction(() => window.tabula.wb().calculation?.mode === 'auto'); await p.reload(); await p.waitForFunction(() => window.tabula?.wb()); equal((await view(p)).manual,false,'계산 속성 없는 외부 파일도 로컬 재열기 시 자동 유지');
  });
  await test('보기 변경·실행 취소·Excel 재저장·시트별 인쇄 눈금선', async p => {
    await load(p); await p.evaluate(() => window.tabula.run('toggleHeaders')); equal((await view(p)).headers, true);
    await p.evaluate(() => window.tabula.run('undo')); equal((await view(p)).headers, false);
    await p.evaluate(() => window.tabula.run('toggleFormulas')); equal((await view(p)).formulas, false);
    await p.evaluate(() => window.tabula.run('togglePrintGrid')); equal(await p.evaluate(() => window.tabula.wb().sheets[0].page.gridlines), true);
    await p.evaluate(() => window.tabula.switchSheet(1)); await p.evaluate(() => window.tabula.run('togglePrintGrid')); await p.evaluate(() => window.tabula.run('undo')); equal(await p.evaluate(() => !!window.tabula.wb().sheets[1].page?.gridlines), false);
    await p.evaluate(() => window.tabula.switchSheet(0)); const data = await save(p); equal(data.sheets[0].view.headers, false); equal(data.sheets[0].view.showFormulas, false); equal(data.sheets[0].view.gridColor, '#ff0000'); equal(data.sheets[0].page.gridlines, true); equal(!!data.sheets[1].page?.gridlines, false);
  });
  await test('수동 계산 유지·F9·자동 전환·저장 전 계산 옵션', async p => {
    await load(p); equal(await p.evaluate(() => window.tabula.wb().getValue(0,0,1)), 6);
    await p.evaluate(() => { const w=window.tabula.wb(); w.transact(()=>w.setInput(0,0,0,'5')); }); equal(await p.evaluate(() => window.tabula.wb().getValue(0,0,1)), 6);
    let data = await save(p); equal(data.calculation.mode, 'manual'); equal(data.calculation.calcOnSave, false); equal(new Workbook(data).getValue(0,0,1), 6);
    let d = await openOptions(p); await d.getByRole('tab', { name: '수식', exact: true }).click(); await d.getByRole('checkbox', { name: '저장하기 전에 통합 문서 다시 계산', exact: true }).check(); await button(d, '확인').click();
    data = await save(p); equal(data.calculation.calcOnSave, true); equal(new Workbook(data).getValue(0,0,1), 10);
    await p.evaluate(() => { const w=window.tabula.wb(); w.transact(()=>w.setInput(0,0,0,'7')); window.tabula.run('recalc'); }); equal(await p.evaluate(() => window.tabula.wb().getValue(0,0,1)), 14);
    await p.evaluate(() => { const w=window.tabula.wb(); w.transact(()=>w.setInput(0,0,0,'9')); window.tabula.run('calcAuto'); }); equal(await p.evaluate(() => window.tabula.wb().getValue(0,0,1)), 18); equal((await view(p)).manual, false);
    await p.evaluate(() => window.tabula.run('undo')); equal((await view(p)).manual, true);
  });
  await test('통합 문서 보호: Excel 암호 검증·옵션 우회 방지·Undo·재저장', async p => {
    await load(p, fixture({ protection: `<workbookProtection lockStructure="1" workbookPassword="${excelHash('native')}"/>` }));
    await p.evaluate(() => window.tabula.run('protectWorkbook')); let d=p.getByRole('dialog',{name:'통합 문서 보호 해제',exact:true}); await d.locator('input[type=password]').fill('wrong'); await button(d,'확인').click(); await d.getByRole('alert').waitFor(); equal(await p.evaluate(()=>window.tabula.wb().props.lockStructure),true);
    await d.locator('input[type=password]').fill('native'); await button(d,'확인').click(); await d.waitFor({state:'detached'}); equal(await p.evaluate(()=>!!window.tabula.wb().props.lockStructure),false);
    await p.evaluate(()=>window.tabula.run('undo')); equal(await p.evaluate(()=>window.tabula.wb().props.lockStructure),true);
    d=await openOptions(p);await d.getByRole('tab',{name:'고급',exact:true}).click();await d.getByRole('button',{name:'통합 문서 구조 보호 해제...',exact:true}).click();let unlock=p.getByRole('dialog',{name:'통합 문서 보호 해제',exact:true});equal(await unlock.count(),1);await button(unlock,'취소').click();await button(d,'확인').click();equal(await p.evaluate(()=>window.tabula.wb().props.lockStructure),true);
    const data=await save(p);equal(data.props.workbookProtection.workbookPassword,excelHash('native'));equal(data.props.lockStructure,true);
    await p.evaluate(()=>window.tabula.run('protectWorkbook'));d=p.getByRole('dialog',{name:'통합 문서 보호 해제',exact:true});await d.locator('input[type=password]').fill('native');await button(d,'확인').click();await d.waitFor({state:'detached'});
    const unlocked=await save(p);equal(!!unlocked.props?.lockStructure,false);equal(unlocked.props?.workbookProtection?.workbookPassword,undefined);
  });
  await test('보호 해제 중 동일한 속성의 다른 문서를 열면 적용 차단', async p => {
    await load(p, fixture({ protection: `<workbookProtection lockStructure="1" workbookPassword="${excelHash('native')}"/>` }));
    await p.evaluate(()=>window.tabula.run('protectWorkbook'));const d=p.getByRole('dialog',{name:'통합 문서 보호 해제',exact:true});await d.locator('input[name=pw]').fill('native');
    await p.evaluate(()=>{const w=window.tabula.wb(),data=w.serialize();data.sheets[0].name='다른 문서';w.load(data);});await button(d,'확인').click();await d.getByRole('alert').waitFor();assert.match(await d.getByRole('alert').innerText(),/통합 문서 또는 보호 설정이 바뀌었습니다/);checks++;
    equal(await p.evaluate(()=>window.tabula.wb().props.lockStructure),true);equal(await p.evaluate(()=>window.tabula.wb().sheets[0].name),'다른 문서');await button(d,'취소').click();
  });
  await test('신규 보호 암호 확인·Excel 저장·해제', async p => {
    await load(p); await p.evaluate(()=>window.tabula.run('protectWorkbook'));const d=p.getByRole('dialog',{name:'통합 문서 구조 보호',exact:true});await d.locator('input[name=pw]').fill('new');await d.locator('input[name=pw2]').fill('different');await button(d,'확인').click();await d.getByRole('alert').waitFor();equal(await p.evaluate(()=>!!window.tabula.wb().props?.lockStructure),false);
    await d.locator('input[name=pw2]').fill('new');await button(d,'확인').click();await d.waitFor({state:'detached'});const data=await save(p);equal(data.props.lockStructure,true);equal(data.props.workbookProtection.workbookPassword,excelHash('new'));
  });
  await test('현대식 Excel 보호 해시: 완료 전 유지·성공 후 해제', async p => {
    const salt=Buffer.from('wixel-audit-salt'),pw='hash-password';let hash=createHash('sha512').update(Buffer.concat([salt,Buffer.from(pw,'utf16le')])).digest();for(let i=0;i<20;i++){const n=Buffer.alloc(4);n.writeUInt32LE(i);hash=createHash('sha512').update(Buffer.concat([hash,n])).digest();}
    await load(p,fixture({protection:`<workbookProtection lockStructure="1" workbookAlgorithmName="SHA-512" workbookHashValue="${hash.toString('base64')}" workbookSaltValue="${salt.toString('base64')}" workbookSpinCount="20"/>`}));await p.evaluate(()=>window.tabula.run('protectWorkbook'));const d=p.getByRole('dialog',{name:'통합 문서 보호 해제',exact:true});await d.locator('input[type=password]').fill(pw);await button(d,'확인').click();await d.waitFor({state:'detached'});equal(await p.evaluate(()=>!!window.tabula.wb().props.lockStructure),false);
  });
  await test('수동 계산 목표값 찾기: 원본 무변경 미리보기·취소·확인·Undo', async p => {
    await load(p); const before = await p.evaluate(()=>{const t=window.tabula;t.selectCell(0,1);return t.wb().version;});
    await p.evaluate(()=>window.tabula.run('goalSeek'));let d=p.getByRole('dialog',{name:'목표값 찾기',exact:true});await d.locator('input[name=to]').fill('18');await d.locator('input[name=by]').fill('A1');await button(d,'확인').click();let status=p.getByRole('dialog',{name:'목표값 찾기 상태',exact:true});await status.waitFor();assert.match(await status.innerText(),/해를 찾았습니다/);checks++;
    equal(await p.evaluate(()=>window.tabula.wb().version),before,'미리보기는 원본을 수정하지 않음');equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),3);await button(status,'취소').click();equal(await p.evaluate(()=>window.tabula.wb().needsCalc),false);
    await p.evaluate(()=>window.tabula.run('goalSeek'));d=p.getByRole('dialog',{name:'목표값 찾기',exact:true});await d.locator('input[name=to]').fill('18');await d.locator('input[name=by]').fill('A1');await button(d,'확인').click();await button(status,'확인').click();equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),9);equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,1)),18);await p.evaluate(()=>window.tabula.run('undo'));equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),3);
  });
  await test('수동 계산 데이터 표: 입력 원본 보존·서로 다른 결과·Undo', async p => {
    await load(p); await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>{w.setInput(0,0,4,'=B1');w.setInput(0,1,3,'1');w.setInput(0,2,3,'2');w.setInput(0,3,3,'3');});w.calculateNow();t.selectRange({r1:0,c1:3,r2:3,c2:4},'cells',{r:0,c:3});});
    await p.evaluate(()=>window.tabula.run('dataTable'));const d=p.getByRole('dialog',{name:'데이터 표',exact:true});await d.locator('input[name=col]').fill('A1');await button(d,'확인').click();await d.waitFor({state:'detached'});equal(await p.evaluate(()=>{const w=window.tabula.wb();return [w.getValue(0,1,4),w.getValue(0,2,4),w.getValue(0,3,4)];}),[2,4,6]);equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),3);await p.evaluate(()=>window.tabula.run('undo'));equal(await p.evaluate(()=>window.tabula.wb().getRaw(0,1,4)),'');equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),3);
  });
  await test('현재 시트 계산: 다른 시트 대기값과 미지원 저장값 보존', async p => {
    await load(p);await p.evaluate(()=>{const t=window.tabula,w=t.wb(),sheet=name=>({name,fileValues:true,cells:{'0,0':{raw:'3'},'0,1':{raw:'=A1*2',cached:6},'0,2':{raw:'=NOT_SUPPORTED()',cached:42}}});w.restore({calculation:{mode:'manual'},sheets:[sheet('A'),sheet('B')]});w.transact(()=>{w.setInput(0,0,0,'5');w.setInput(1,0,0,'7');});});
    const values=()=>p.evaluate(()=>{const w=window.tabula.wb();return [w.getValue(0,0,1),w.getValue(1,0,1),w.getValue(0,0,2),w.getValue(1,0,2),w.needsCalc];});
    equal(await values(),[6,6,42,42,true]);await p.evaluate(()=>window.tabula.run('calcNowSheet'));equal(await values(),[10,6,42,42,true]);await p.evaluate(()=>window.tabula.run('recalc'));equal(await values(),[10,14,42,42,false]);
  });
  await test('가상 분석은 이전 텍스트 자료형을 숫자 결과에 남기지 않는다', async p => {
    await load(p);await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.restore({calculation:{mode:'manual'},sheets:[{name:'문자',cells:{'0,0':{raw:'2',inputType:'text',comment:'메모'},'0,1':{raw:'=SUM(A1)*2'},'0,4':{raw:'=B1'},'1,3':{raw:'2'},'1,4':{raw:'이전 글자',inputType:'text',style:{numFmt:'text'}}}}]});t.selectCell(0,1);});
    await p.evaluate(()=>window.tabula.run('goalSeek'));let d=p.getByRole('dialog',{name:'목표값 찾기',exact:true});await d.locator('input[name=to]').fill('20');await d.locator('input[name=by]').fill('A1');await button(d,'확인').click();const status=p.getByRole('dialog',{name:'목표값 찾기 상태',exact:true});await status.waitFor();assert.match(await status.innerText(),/해를 찾았습니다/);checks++;await button(status,'확인').click();const solved=await p.evaluate(()=>window.tabula.wb().getValue(0,0,0));equal(typeof solved,'number');assert.ok(Math.abs(solved-10)<1e-9);checks++;equal(await p.evaluate(()=>window.tabula.wb().getCell(0,0,0).comment),'메모');await p.evaluate(()=>window.tabula.run('undo'));equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),'2');
    await p.evaluate(()=>{const t=window.tabula;t.selectRange({r1:0,c1:3,r2:1,c2:4},'cells',{r:0,c:3});t.run('dataTable');});d=p.getByRole('dialog',{name:'데이터 표',exact:true});await d.locator('input[name=col]').fill('A1');await button(d,'확인').click();await d.waitFor({state:'detached'});equal(await p.evaluate(()=>window.tabula.wb().getValue(0,1,4)),4);equal(await p.evaluate(()=>window.tabula.wb().getValue(0,0,0)),'2');
  });
  // Optional fixtures are generated by an isolated Excel instance; never select a user's workbook automatically.
  if (process.env.WIXEL_NATIVE_SETTINGS_DIR) for (const ext of ['xlsx', 'xlsb']) await test(`실제 Excel ${ext}: 보기·인쇄·유효성·현대식 암호`, async p => {
    const bytes=await readFile(`${process.env.WIXEL_NATIVE_SETTINGS_DIR}/native.${ext}`);
    await p.locator('#fileInput').setInputFiles({name:`native.${ext}`,mimeType:'application/octet-stream',buffer:bytes});await p.waitForFunction(()=>window.tabula.wb().calculation?.iterateCount===77);
    const warning=p.getByRole('dialog',{name:'가져오기',exact:true});await warning.waitFor();assert.match(await warning.innerText(),/반복 계산/);checks++;await button(warning,'확인').click();
    equal(await view(p),{headers:false,formulas:true,grid:true,color:'#003366',manual:true});
    const state=await p.evaluate(()=>{const t=window.tabula,w=t.wb(),s=w.sheets[0];return {mode:w.calculation.mode,onSave:w.calculation.calcOnSave,freeze:s.freeze,page:s.page,validations:s.validations.length,hiddenRow:!!s.hiddenRows[4],hiddenCol:!!s.hiddenCols[3],locked:w.props.lockStructure};});
    equal(state.mode,'manual');equal(state.onSave,false);equal(state.freeze,{rows:2,cols:1});equal(state.validations,1);equal(state.hiddenRow,true);equal(state.hiddenCol,true);equal(state.locked,true);equal(state.page.orientation,'landscape');equal(state.page.scale,85);
    await p.evaluate(()=>window.tabula.run('protectWorkbook'));let d=p.getByRole('dialog',{name:'통합 문서 보호 해제',exact:true});await d.locator('input[type=password]').fill('book123');await button(d,'확인').click();await d.waitFor({state:'detached',timeout:120000});equal(await p.evaluate(()=>!!window.tabula.wb().props.lockStructure),false);
    await p.evaluate(()=>window.tabula.run('unprotectSheet'));d=p.getByRole('dialog',{name:'시트 보호 해제',exact:true});await d.locator('input[type=password]').fill('sheet123');await button(d,'확인').click();await d.waitFor({state:'detached',timeout:120000});equal(await p.evaluate(()=>!!window.tabula.wb().sheets[0].protect?.on),false);
  });
  const failed=results.filter(r=>!r.ok).length;console.log(JSON.stringify({tests:results.length,checks,failed,results},null,2));if(failed)process.exitCode=1;
} finally { await browser.close(); }
