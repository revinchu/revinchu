// 선택하여 붙여넣기: 합성 문서·모의 시스템 클립보드·격리 브라우저. 원격 읽기/쓰기 차단.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/', browser = await chromium.launch(), results = [];
const filter = process.env.WIXEL_PASTE_SPECIAL_FILTER;
async function test(name, fn) {
  if (filter && !name.includes(filter)) return;
  const ctx = await browser.newContext({ viewport: { width: 1400, height: 1000 } }), p = await ctx.newPage(), errors = [], writes = [];
  p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message)); p.on('dialog', d => d.dismiss());
  await ctx.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET','HEAD','OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); } return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort(); });
  try {
    await ctx.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.__clipboardText = ''; window.__clipboardDenied = false; Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => { if (window.__clipboardDenied) throw new DOMException('합성 권한 거부', 'NotAllowedError'); return window.__clipboardText; }, writeText: async value => { window.__clipboardText = value; } } }); });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb()); await fixture(p); await fn(p);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기'); results.push({ name, ok: true }); console.log('OK ' + name);
  } catch(e) { results.push({ name, ok: false, error: e.message, pageErrors: errors, writes }); console.error('NG ' + name + ': ' + e.stack); }
  finally { await ctx.close(); }
}
async function fixture(p) { await p.evaluate(() => { const t=window.tabula,w=t.wb();w.restore({sheets:[{name:'붙여넣기 합성',cells:{'0,0':{raw:'2',style:{fill:'#ffee11',numFmt:'percent',decimals:2,bl:true,locked:false},comment:'원본 메모'},'0,1':{raw:'=A1*3'},'1,0':{raw:'3'},'1,1':{raw:''},'4,4':{raw:'10',style:{fill:'#112233',bt:true,locked:true},comment:'대상 메모',link:'https://example.com'},'4,5':{raw:'20'},'5,4':{raw:'30'},'5,5':{raw:'40'}},colWidths:{0:123,4:80},cond:[{r1:0,c1:0,r2:1,c2:1,type:'formula',formula:'=A1>0',style:{fill:'#ff0000'}},{r1:4,c1:4,r2:5,c2:5,type:'formula',formula:'=E5<0',style:{fill:'#00ff00'}}],validations:[{r1:0,c1:0,r2:1,c2:1,type:'custom',f1:'A1>0'},{r1:4,c1:4,r2:5,c2:5,type:'whole',f1:'9'}]},{name:'다른 시트',cells:{}}]});t.switchSheet(0);t.selectRange({r1:0,c1:0,r2:1,c2:1});w.undoStack=[];w.redoStack=[]; }); }
const dlg = p => p.getByRole('dialog', { name: '선택하여 붙여넣기', exact: true });
const run = (p, cmd) => p.evaluate(cmd => window.tabula.run(cmd), cmd);
const serialized = p => p.evaluate(() => JSON.stringify(window.tabula.wb().serialize()));
const state = p => p.evaluate(() => { const w=window.tabula.wb();return { raw:w.getRaw(0,4,4),value:w.getValue(0,4,4),other:w.getValue(0,4,5),style:w.styleAt(0,4,4),cell:w.getCell(0,4,4),cond:w.sheets[0].cond,validations:w.sheets[0].validations,width:w.colWidth(0,4),undo:w.undoStack.length }; });
async function prepare(p, single=false) { if(single)await p.evaluate(()=>window.tabula.selectCell(0,0)); await run(p,'copy');await p.evaluate(()=>window.tabula.selectCell(4,4));await p.locator('#cellEditor').focus(); }
async function open(p) { await p.keyboard.press('Control+Alt+v');await dlg(p).waitFor();await p.waitForFunction(()=>!document.querySelector('.paste-special-dialog [role="status"]')?.textContent.includes('확인하는 중')); return dlg(p); }
async function confirm(p) { await dlg(p).getByRole('button',{name:'확인',exact:true}).click();await dlg(p).waitFor({state:'detached'}); }
const types=[['all','모두'],['formulas','수식'],['values','값'],['formats','서식'],['comments','메모'],['validation','유효성 검사'],['sourceTheme','원본 테마 사용'],['noBorders','테두리만 제외'],['colWidths','열 너비'],['formulasNum','수식 및 숫자 서식'],['valuesNum','값 및 숫자 서식'],['mergeCond','조건부 서식 병합']];
try {
 for(const [what,label] of types) await test('옵션 '+label+' 실제 적용·한 번 실행 취소', async p => {
   await prepare(p);const before=await serialized(p);const d=await open(p);assert.equal(await d.getByRole('radio').count(),17);await d.getByRole('radio',{name:label,exact:true}).check();await confirm(p);const s=await state(p);
   assert.equal(s.undo,1);assert.equal(s.raw,['formats','comments','validation','colWidths'].includes(what)?'10':'2');
   if(['all','sourceTheme','mergeCond','formats','noBorders'].includes(what))assert.equal(s.style.fill,'#ffee11');
   if(what==='values'||what==='valuesNum')assert.equal(s.other,6);
   if(what==='comments')assert.equal(s.cell.comment,'원본 메모');
   if(what==='validation')assert.ok(s.validations.some(x=>x.r1===4&&x.c1===4&&x.f1==='E5>0'));
   if(what==='noBorders'){assert.equal(s.style.bt,true);assert.equal(s.style.bl,false);}
   if(what==='colWidths')assert.equal(s.width,123);
   if(what==='mergeCond')assert.equal(s.cond.filter(x=>x.r1===4&&x.c1===4).length,2);
   if(what==='valuesNum'||what==='formulasNum'){assert.equal(s.style.numFmt,'percent');assert.equal(s.style.decimals,2);assert.equal(s.style.fill,'#112233');}
   await run(p,'undo');assert.equal(await serialized(p),before);
 });
 await test('Alt 접근키 12종·연산·건너뛰기·전치와 Escape 취소',async p=>{
   await prepare(p);const before=await serialized(p),d=await open(p);const letters=['a','f','v','t','c','n','h','x','w','r','u','g'];
   for(let i=0;i<letters.length;i++){await p.keyboard.press('Alt+'+letters[i]);assert.equal(await d.getByRole('radio',{name:types[i][1],exact:true}).isChecked(),true,'Alt+'+letters[i]);}
   await p.keyboard.press('Alt+v');await p.keyboard.press('Alt+d');assert.equal(await d.getByRole('radio',{name:'더하기',exact:true}).isChecked(),true);await p.keyboard.press('Alt+b');await p.keyboard.press('Alt+e');assert.equal(await d.getByRole('checkbox',{name:'내용 있는 셀만 붙여넣기',exact:true}).isChecked(),true);assert.equal(await d.getByRole('checkbox',{name:'행/열 바꿈',exact:true}).isChecked(),true);
   await p.screenshot({path:process.env.WIXEL_PASTE_SPECIAL_SCREENSHOT||'D:/Codex/Temp/wixel-paste-special.png'});await p.keyboard.press('Escape');assert.equal(await d.count(),0);assert.equal(await serialized(p),before);
 });
 await test('문자 V/D/B/E 및 Enter: 값 더하기·빈칸 보존·전치',async p=>{
   await prepare(p);await open(p);for(const key of ['v','d','b','e'])await p.keyboard.press(key);await p.keyboard.press('Enter');await dlg(p).waitFor({state:'detached'});const values=await p.evaluate(()=>{const w=window.tabula.wb();return [[w.getValue(0,4,4),w.getValue(0,4,5)],[w.getValue(0,5,4),w.getValue(0,5,5)]]});assert.deepEqual(values,[[12,23],[36,40]]);assert.equal((await state(p)).undo,1);
 });
 await test('연결 L: 실제 절대 참조·원본 변경 반영·Undo',async p=>{
   await prepare(p);const before=await serialized(p);await open(p);await p.keyboard.press('Alt+l');await dlg(p).waitFor({state:'detached'});assert.equal((await state(p)).raw,'=$A$1');await run(p,'undo');assert.equal(await serialized(p),before);
 });
 await test('외부 TSV만 있어도 Ctrl+Alt+V가 열리고 값/전치를 적용',async p=>{
   await p.evaluate(()=>{window.__clipboardText='8\t9\n10\t11';window.tabula.selectCell(4,4)});await p.locator('#cellEditor').focus();const d=await open(p);assert.match(await d.getByRole('status').textContent(),/외부 텍스트/);assert.equal(await d.getByRole('radio',{name:'서식',exact:true}).isDisabled(),true);await d.getByRole('radio',{name:'값',exact:true}).check();await d.getByRole('checkbox',{name:'행/열 바꿈',exact:true}).check();await confirm(p);assert.equal((await state(p)).value,8);assert.equal((await state(p)).other,10);
 });
 await test('클립보드 권한 거부 뒤 대화상자 Ctrl+V 입력으로 완료',async p=>{
   await p.evaluate(()=>{window.__clipboardDenied=true;window.tabula.selectCell(4,4)});await p.locator('#cellEditor').focus();const d=await open(p);const input=d.getByRole('textbox',{name:'외부 클립보드 데이터',exact:true});await input.evaluate(el=>{const transfer=new DataTransfer();transfer.setData('text/plain','25\t26');el.dispatchEvent(new ClipboardEvent('paste',{clipboardData:transfer,bubbles:true,cancelable:true}));});await confirm(p);assert.equal((await state(p)).value,25);assert.equal((await state(p)).other,26);
 });
 await test('복사 후 외부 클립보드가 바뀌면 오래된 내부 복사보다 새 TSV 사용',async p=>{
   await prepare(p);await p.evaluate(()=>window.__clipboardText='77');const d=await open(p);assert.match(await d.getByRole('status').textContent(),/외부 텍스트/);await confirm(p);assert.equal((await state(p)).value,77);
 });
 await test('한국어 키 값이어도 물리 KeyV Ctrl+Alt 단축키 열림',async p=>{
   await prepare(p);await p.locator('#cellEditor').evaluate(el=>el.dispatchEvent(new KeyboardEvent('keydown',{key:'ㅍ',code:'KeyV',ctrlKey:true,altKey:true,bubbles:true,cancelable:true})));await dlg(p).waitFor();assert.equal(await dlg(p).getByRole('radio',{name:'모두',exact:true}).isChecked(),true);
 });
 await test('보호된 대상은 값 붙여넣기 차단·서식 권한으로 잠금 해제 불가',async p=>{
   await prepare(p,true);await p.evaluate(()=>{window.tabula.wb().sheets[0].protect={on:true,allow:{formatCells:true}}});const before=await serialized(p),d=await open(p);await d.getByRole('radio',{name:'값',exact:true}).check();await d.getByRole('button',{name:'확인',exact:true}).click();assert.equal(await serialized(p),before);assert.equal((await state(p)).undo,0);const blocked=p.getByRole('dialog',{name:'WIXEL',exact:true});if(await blocked.count())await blocked.getByRole('button',{name:'확인',exact:true}).click();await d.getByRole('radio',{name:'서식',exact:true}).check();await confirm(p);const s=await state(p);assert.equal(s.style.locked,true);assert.equal(s.value,10);assert.equal(s.style.fill,'#ffee11');
 });
 await test('창이 열린 뒤 시트 교체는 새 시트에 쓰지 않음',async p=>{
   await prepare(p);const d=await open(p);await p.evaluate(()=>window.tabula.switchSheet(1));const before=await serialized(p);await d.getByRole('button',{name:'확인',exact:true}).click();await d.getByRole('alert').waitFor();assert.match(await d.getByRole('alert').textContent(),/시트가 바뀌/);assert.equal(await serialized(p),before);
 });
 await test('Ctrl+Shift+V도 외부 TSV를 값으로 붙여넣기',async p=>{
   await p.evaluate(()=>{window.__clipboardText='91\t92';window.tabula.selectCell(4,4)});await p.locator('#cellEditor').focus();await p.keyboard.press('Control+Shift+v');await p.waitForFunction(()=>window.tabula.wb().getValue(0,4,4)===91);assert.equal((await state(p)).other,92);assert.equal((await state(p)).undo,1);
 });
 await test('전치 수식은 상대 참조의 행과 열도 바뀜',async p=>{
   await p.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>{w.setInput(0,0,0,'=B1');w.setInput(0,0,1,'5')});t.selectRange({r1:0,c1:0,r2:0,c2:1})});await run(p,'copy');await p.evaluate(()=>{window.tabula.selectCell(4,4);window.tabula.wb().undoStack=[]});await p.locator('#cellEditor').focus();const d=await open(p);await d.getByRole('checkbox',{name:'행/열 바꿈',exact:true}).check();await confirm(p);assert.equal((await state(p)).raw,'=E6');assert.equal((await state(p)).value,5);assert.equal((await state(p)).undo,1);
 });
 await test('값 + 더하기에서도 기존 대상 수식은 보존',async p=>{
   await prepare(p,true);await p.evaluate(()=>{const w=window.tabula.wb();w.transact(()=>w.setInput(0,4,4,'=F5'));w.undoStack=[]});const d=await open(p);await d.getByRole('radio',{name:'값',exact:true}).check();await d.getByRole('radio',{name:'더하기',exact:true}).check();await confirm(p);assert.equal((await state(p)).raw,'=(F5)+2');assert.equal((await state(p)).value,22);
 });
 await test('클립보드 읽기 대기 중 수동 입력은 늦은 읽기 결과보다 우선',async p=>{
   await p.evaluate(()=>{navigator.clipboard.readText=()=>new Promise(r=>window.__releaseClipboard=r);window.tabula.selectCell(4,4)});await p.locator('#cellEditor').focus();await p.keyboard.press('Control+Alt+v');const d=dlg(p);await d.waitFor();await p.waitForFunction(()=>window.__releaseClipboard);await d.getByRole('textbox',{name:'외부 클립보드 데이터',exact:true}).fill('88');await p.evaluate(()=>window.__releaseClipboard('99'));await confirm(p);assert.equal((await state(p)).value,88);
 });
 await test('Ctrl+Shift+V 읽기 대기 중 시트 교체는 쓰기 차단',async p=>{
   await p.evaluate(()=>{navigator.clipboard.readText=()=>new Promise(r=>window.__releaseClipboard=r);window.tabula.selectCell(4,4)});await p.locator('#cellEditor').focus();await p.keyboard.press('Control+Shift+v');await p.waitForFunction(()=>window.__releaseClipboard);await p.evaluate(()=>window.tabula.switchSheet(1));const before=await serialized(p);await p.evaluate(()=>window.__releaseClipboard('99'));await p.waitForTimeout(80);assert.equal(await serialized(p),before);assert.equal((await state(p)).undo,0);
 });
 for (const mode of ['paste','pasteValuesKey']) await test('늦은 붙여넣기: 같은 시작 셀에서 선택 범위 변경 '+mode,async p=>{
   await p.evaluate(mode=>{navigator.clipboard.readText=()=>new Promise(r=>window.__releaseClipboard=r);tabula.selectCell(4,4);tabula.run(mode)},mode);
   await p.waitForFunction(()=>window.__releaseClipboard);
   await p.evaluate(()=>tabula.selectRange({r1:4,c1:4,r2:5,c2:5},'cells',{r:4,c:4}));const before=await serialized(p);
   await p.evaluate(()=>window.__releaseClipboard('99'));await p.waitForTimeout(80);assert.equal(await serialized(p),before);assert.equal((await state(p)).undo,0);
 });
 await test('늦은 붙여넣기: 대기 중 편집한 값 유지',async p=>{
   await p.evaluate(()=>{navigator.clipboard.readText=()=>new Promise(r=>window.__releaseClipboard=r);tabula.selectCell(4,4);tabula.run('paste')});await p.waitForFunction(()=>window.__releaseClipboard);
   await p.evaluate(()=>{const w=tabula.wb();w.transact(()=>w.setInput(0,4,4,'123'))});const before=await serialized(p);
   await p.evaluate(()=>window.__releaseClipboard('99'));await p.waitForTimeout(80);assert.equal(await serialized(p),before);assert.equal((await state(p)).value,123);
 });
 await test('늦은 붙여넣기: 나중 요청의 결과를 이전 응답이 덮어쓰지 않음',async p=>{
   await p.evaluate(()=>{window.__pasteReads=[];navigator.clipboard.readText=()=>new Promise(r=>window.__pasteReads.push(r));tabula.selectCell(4,4);tabula.run('paste');tabula.run('paste')});await p.waitForFunction(()=>window.__pasteReads.length===2);
   await p.evaluate(()=>window.__pasteReads[1]('222'));await p.waitForFunction(()=>tabula.wb().getValue(0,4,4)===222);const before=await serialized(p);
   await p.evaluate(()=>window.__pasteReads[0]('111'));await p.waitForTimeout(80);assert.equal(await serialized(p),before);assert.equal((await state(p)).undo,1);
 });
} finally { await browser.close(); }
const bad=results.filter(x=>!x.ok);console.log(JSON.stringify({total:results.length,good:results.length-bad.length,bad:bad.length,results},null,2));if(bad.length)process.exitCode=1;
