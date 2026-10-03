// 인쇄/테마 합성 문서만 검사한다. 사용자 탭·서버 보관함에는 접근하지 않는다.
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { newSmartArt } from '../src/smartart.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 합성 검사에서만 실행하세요.');
const browser = await chromium.launch(), context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: true });
const page = await context.newPage(), errors = [], writes = [], results = [];
page.setDefaultTimeout(20000); page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); } return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort(); });
await context.addInitScript(() => { window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; window.__printCalls = 0; window.__pdfFiles=[]; window.showSaveFilePicker=async options=>({name:options.suggestedName,createWritable:async()=>{const chunks=[];return{write:async data=>chunks.push(data),close:async()=>{const data=await new Promise((resolve,reject)=>{const reader=new FileReader();reader.onload=()=>resolve(reader.result);reader.onerror=reject;reader.readAsDataURL(new Blob(chunks));});window.__pdfFiles.push({name:options.suggestedName,data});}};}}); window.print = () => { window.__printCalls++; }; });
const savePdf = async path => { const count=await page.evaluate(()=>__pdfFiles.length);await page.getByRole('button',{name:'PDF 파일 저장',exact:true}).click();await page.waitForFunction(n=>__pdfFiles.length===n+1,count,{timeout:60000});const file=await page.evaluate(()=>__pdfFiles.at(-1));assert.ok(file.name.endsWith('.pdf'));const bytes=Buffer.from(file.data.split(',')[1],'base64');writeFileSync(path,bytes);return bytes; };
const current = () => page.evaluate(() => ({ page: tabula.wb().sheets[0].page, theme: tabula.wb().theme, fonts: tabula.wb().themeFonts, effects: tabula.wb().themeEffects, font: tabula.wb().defaultFont, undo: tabula.wb().undoStack.length }));
try {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => window.tabula?.wb());
  await page.evaluate(() => {
    const cells = {}, rowHeights = {}, colWidths = {};
    for (let r = 0; r < 64; r++) { rowHeights[r] = 28; for (let c = 0; c < 8; c++) cells[`${r},${c}`] = { raw: r === 0 ? `보고서 제목 ${c + 1}` : c === 0 ? `한글 항목 ${r}` : String(r * (c + 1)), style: r === 0 ? { fill: '#17365d', color: '#ffffff', bold: true } : { numFmt: 'number', decimals: 0, bb: 'thin', bbc: '#d9e2f3' } }; }
    for (let c = 0; c < 8; c++) colWidths[c] = 100;
    tabula.wb().restore({ sheets: [{ name: '합성 인쇄', cells, rowHeights, colWidths, page: { area: { r1: 0, c1: 0, r2: 63, c2: 7 }, fitW: 1, fitH: 0, titleRows: [0, 0], header: '&C보고서 &P / &N', footer: '&L합성 검증&R&A', gridlines: true } }] });
    tabula.gv().layout(); tabula.gv().renderAll(); tabula.selectCell(0, 0);
    window.__printBefore = JSON.stringify(tabula.wb().serialize()); tabula.run('print');
  });
  await page.getByRole('dialog', { name: '인쇄 미리보기' }).waitFor();
  assert.equal(await page.evaluate(() => __printCalls), 0);
  const preview = page.locator('.print-preview');
  assert.match(await preview.innerText(), /2쪽/); assert.match(await preview.locator('.wixel-print-page').innerText(), /보고서 1 \/ 2/);
  await preview.getByRole('button', { name: '다음 인쇄 페이지' }).click();
  const second = await preview.locator('.wixel-print-page').innerText(); assert.match(second, /보고서 2 \/ 2/); assert.match(second, /보고서 제목 1/); assert.match(second, /한글 항목 63/);
  results.push('명령 실행 시 보이는 미리보기·2쪽 분할·한글·반복 제목·쪽수 머리글');
  const pdfPath = process.env.WIXEL_PRINT_PDF || 'D:/Codex/Temp/wixel-page-layout.pdf';
  await savePdf(pdfPath);
  const bytes = readFileSync(pdfPath), text = bytes.toString('latin1');
  assert.ok(bytes.length > 20000); assert.ok(text.startsWith('%PDF-1.4')); assert.match(text, /\/Count 2/);
  assert.equal((text.match(/\/Filter \/DCTDecode/g) ?? []).length, 2); assert.equal(await page.evaluate(()=>__pdfFiles.length), 1);
  assert.match(await preview.getByRole('status').innerText(), /2쪽 PDF 파일을 저장했습니다/);
  results.push('명시적 PDF 저장·모의 파일 핸들의 실제 PDF 2페이지/JPEG 스트림·쓰기 완료');
  await preview.getByRole('button', { name: '인쇄', exact: true }).click();
  assert.equal(await page.evaluate(() => __printCalls), 1); assert.equal(await page.locator('#printArea > .wixel-print-page').count(), 2);
  assert.match(await preview.getByRole('status').innerText(), /인쇄 창이 열리지 않으면/);
  await page.pdf({ path: 'D:/Codex/Temp/wixel-page-layout-native.pdf', preferCSSPageSize: true, printBackground: true });
  assert.equal(await page.evaluate(() => JSON.stringify(tabula.wb().serialize()) === __printBefore), true);
  results.push('브라우저 인쇄 호출·동일 페이지 DOM·네이티브 PDF 생성·문서 불변');
  await page.screenshot({ path: process.env.WIXEL_PRINT_SCREENSHOT || 'D:/Codex/Temp/wixel-page-layout.png' });
  await page.getByRole('dialog', { name: '인쇄 미리보기' }).getByRole('button', { name: '닫기', exact: true }).last().click();
  await page.getByRole('button', { name: '페이지 레이아웃', exact: true }).click();
  const width = page.getByTitle('인쇄 너비 맞춤', { exact: true }), height = page.getByTitle('인쇄 높이 맞춤', { exact: true }), scale = page.getByTitle('인쇄 배율 (%) · 너비와 높이가 자동일 때 사용', { exact: true });
  assert.equal(await width.inputValue(), '1'); assert.equal(await scale.isDisabled(), true);
  await width.selectOption('0'); assert.equal(await scale.isDisabled(), false);
  await scale.fill('75'); await scale.press('Enter'); assert.equal((await current()).page.scale, 75);
  await height.selectOption('1'); assert.equal((await current()).page.fitH, 1); assert.equal(await scale.isDisabled(), true);
  await page.evaluate(() => tabula.run('undo')); assert.equal((await current()).page.fitH, 0);
  results.push('리본 너비/높이 자동·수동 배율·fit 비활성 상태·한 번 Undo');
  await page.evaluate(() => tabula.openNamedMenu('themeFonts', { x: 200, y: 180 }));
  await page.getByRole('menuitem', { name: 'Arial — Arial / Arial', exact: true }).click();
  assert.equal((await current()).fonts.minor, 'Arial'); assert.equal((await current()).font.name, 'Arial');
  await page.evaluate(() => tabula.run('undo')); assert.equal((await current()).fonts, null);
  await page.evaluate(() => tabula.run('redo')); assert.equal((await current()).font.name, 'Arial');
  await page.evaluate(() => tabula.openNamedMenu('themeEffects', { x: 200, y: 180 }));
  await page.getByRole('menuitem', { name: '보통 그림자', exact: true }).click(); assert.equal((await current()).effects.id, 'moderate');
  results.push('테마 제목/본문 글꼴 실제 기본폰트·Undo/Redo·기본 효과 선택');
  const smartArt = newSmartArt('process', { id: 'print-smartart', x: 100, y: 300, w: 600, h: 240 });
  smartArt.smartArt.nodes.forEach((node, i) => { node.text = ['계획', '실행', '검토'][i]; });
  await page.evaluate(smartArt => {
    const canvas = document.createElement('canvas'); canvas.width = 320; canvas.height = 180;
    const ctx = canvas.getContext('2d'); ctx.fillStyle = '#1463b0'; ctx.fillRect(0, 0, 160, 180); ctx.fillStyle = '#f4bd42'; ctx.fillRect(160, 0, 160, 180); ctx.fillStyle = '#ffffff'; ctx.font = 'bold 28px Malgun Gothic'; ctx.fillText('그림 PDF', 85, 100);
    tabula.wb().restore({ sheets: [{ name: '그림과 SmartArt', page: { fitW: 1, fitH: 0 }, cells: { '0,0': { raw: '그림·도형 인쇄' } }, images: [{ id: 'print-image', src: canvas.toDataURL('image/png'), x: 100, y: 80, w: 320, h: 180 }, { id: 'hidden-image', hidden: true, src: canvas.toDataURL('image/png'), x: 2000, y: 2000, w: 320, h: 180 }, { id: 'no-print-image', noPrint: true, src: canvas.toDataURL('image/png'), x: 500, y: 80, w: 320, h: 180 }], shapes: [smartArt, { id: 'print-text', kind: 'rect', x: 100, y: 650, w: 320, h: 120, fill: '#fce4d6', stroke: '#c65911', text: '도형 텍스트와 한국어\n인쇄 검증', size: 20, align: 'center', valign: 'middle', pad: [10, 10, 10, 10] }] }] });
    tabula.gv().layout(); tabula.gv().renderAll(); tabula.run('print');
  }, smartArt);
  await page.getByRole('dialog', { name: '인쇄 미리보기' }).waitFor();
  assert.match(await preview.innerText(), /1쪽/);
  assert.equal(await preview.locator('[data-print-object="hidden-image"],[data-print-object="no-print-image"]').count(), 0);
  assert.equal(await preview.locator('[data-print-object="print-image"]').count(), 1);
  assert.equal(await preview.locator('[data-print-object="print-smartart"]').count(), 1);
  const pictureBox = await preview.locator('[data-print-object="print-image"]').boundingBox(), smartBox = await preview.locator('[data-print-object="print-smartart"]').boundingBox();
  assert.ok(smartBox.y > pictureBox.y + pictureBox.height); assert.ok(Math.abs(smartBox.x - pictureBox.x) < 1);
  await savePdf('D:/Codex/Temp/wixel-page-layout-objects.pdf');
  const objectPdf = readFileSync('D:/Codex/Temp/wixel-page-layout-objects.pdf'); assert.match(objectPdf.toString('latin1'), /\/Count 1/); assert.ok(objectPdf.length > 30000);
  assert.match(await preview.locator('.wixel-print-page').innerText(), /계획/);
  await page.screenshot({ path: 'D:/Codex/Temp/wixel-print-smartart.png' });
  await page.getByRole('dialog', { name: '인쇄 미리보기' }).getByRole('button', { name: '닫기', exact: true }).last().click();
  results.push('그림/SmartArt/도형 텍스트 원래좌표 동일1쪽PDF·숨김/인쇄제외 보존');
  await page.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true } })); tabula.openNamedMenu('themeFonts', { x: 200, y: 180 }); });
  await page.getByRole('menuitem', { name: 'Arial — Arial / Arial', exact: true }).click();
  assert.equal((await current()).fonts, null);
  await page.getByRole('dialog').last().getByRole('button', { name: '확인', exact: true }).click();
  await page.setViewportSize({ width: 320, height: 480 });
  await page.evaluate(() => tabula.run('print')); await page.getByRole('dialog', { name: '인쇄 미리보기' }).waitFor();
  const bounds = await page.getByRole('dialog', { name: '인쇄 미리보기' }).boundingBox();
  assert.ok(bounds.x >= -1 && bounds.y >= -1 && bounds.x + bounds.width <= 321 && bounds.y + bounds.height <= 481, JSON.stringify(bounds));
  await page.getByRole('dialog', { name: '인쇄 미리보기' }).getByRole('button', { name: '닫기', exact: true }).last().click();
  results.push('보호된 시트 테마 변경 차단·인쇄 허용·320×480 미리보기 화면 안 배치');
  await page.setViewportSize({ width: 1440, height: 1000 });
  await page.evaluate(() => {
    const cells = {}, rowHeights = {}, colWidths = {};
    for (let r = 0; r < 64; r++) { rowHeights[r] = 28; for (let c = 0; c < 8; c++) cells[`${r},${c}`] = { raw: r === 0 ? `반복 제목 ${c + 1}` : `행${r}열${c}`, style: { fill: r === 0 ? '#d9e2f3' : '#ffffff', color: '#222222', bb: 'thin', bbc: '#aaaaaa' } }; }
    cells['36,1'].raw = '페이지 경계 병합'; for (let c = 0; c < 8; c++) colWidths[c] = 100;
    tabula.wb().restore({ sheets: [{ name: '병합 세로 경계', cells, rowHeights, colWidths, merges: [{r1:0,c1:0,r2:0,c2:2},{r1:36,c1:1,r2:43,c2:3}], page: {area:{r1:0,c1:0,r2:63,c2:7},fitW:1,fitH:0,titleRows:[0,0],gridlines:true} }] });
    tabula.gv().layout(); tabula.gv().renderAll(); tabula.run('print');
  });
  await page.getByRole('dialog', { name: '인쇄 미리보기' }).waitFor(); assert.match(await preview.innerText(), /2쪽/);
  const merged1 = preview.locator('td').filter({hasText:'페이지 경계 병합'});
  const rs1 = Number(await merged1.getAttribute('rowspan')); assert.equal(await merged1.getAttribute('colspan'), '3'); assert.ok(rs1 > 0 && rs1 < 8);
  assert.equal(await preview.locator('thead td').first().getAttribute('colspan'), '3');
  await preview.getByRole('button',{name:'다음 인쇄 페이지'}).click();
  const merged2 = preview.locator('td').filter({hasText:'페이지 경계 병합'});
  assert.equal(Number(await merged2.getAttribute('rowspan')) + rs1, 8); assert.equal(await merged2.getAttribute('colspan'), '3');
  const row43 = preview.locator('tr[data-print-row="43"]'); assert.equal(await row43.locator('td').nth(1).innerText(), '행43열4');
  await savePdf('D:/Codex/Temp/wixel-page-layout-merged.pdf');
  await page.getByRole('dialog',{name:'인쇄 미리보기'}).getByRole('button',{name:'닫기',exact:true}).last().click();
  results.push('세로2쪽 병합셀 분할·원래8행 합계·열위치 유지·반복 병합제목·실제PDF');
  await page.evaluate(() => {
    const cells = {}, colWidths = {}; for(let r=0;r<3;r++) for(let c=0;c<5;c++) { cells[`${r},${c}`] = {raw:`행${r}열${c}`}; colWidths[c]=200; }
    cells['1,1'].raw='가로 경계 병합';
    tabula.wb().restore({sheets:[{name:'병합 가로 경계',cells,colWidths,merges:[{r1:1,c1:1,r2:1,c2:4}],page:{area:{r1:0,c1:0,r2:2,c2:4},fitW:0,fitH:0,scale:100,titleCols:[0,0],gridlines:true}}]});
    tabula.gv().layout();tabula.gv().renderAll();tabula.run('print');
  });
  await page.getByRole('dialog',{name:'인쇄 미리보기'}).waitFor(); assert.match(await preview.innerText(),/2쪽/);
  assert.equal(await preview.locator('td').filter({hasText:'가로 경계 병합'}).getAttribute('colspan'),'2');
  await preview.getByRole('button',{name:'다음 인쇄 페이지'}).click();
  assert.equal(await preview.locator('td').filter({hasText:'가로 경계 병합'}).getAttribute('colspan'),'2');
  assert.deepEqual(await preview.locator('tbody tr').first().locator('td').allTextContents(),['행0열0','행0열3','행0열4']);
  await page.getByRole('dialog',{name:'인쇄 미리보기'}).getByRole('button',{name:'닫기',exact:true}).last().click();
  results.push('가로2쪽 병합셀 논리열 매핑·반복열·선택범위 보존');
  await page.evaluate(() => {
    const cells = {}, rowHeights = {}, colWidths = {};
    for(let r=0;r<60;r++) {rowHeights[r]=28; cells[`${r},0`]={raw:`경계 행 ${r}`};} for(let c=0;c<6;c++) colWidths[c]=100;
    tabula.wb().restore({sheets:[{name:'인쇄영역과 교차개체',cells,rowHeights,colWidths,shapes:[
      {id:'cross-page',kind:'rect',x:100,y:850,w:200,h:200,fill:'#00a050',stroke:'#004020'},
      {id:'outside-area',kind:'rect',x:900,y:50,w:100,h:100,fill:'#ff0000'},
      {id:'hidden-shape',kind:'rect',x:100,y:50,w:100,h:100,fill:'#000000',hidden:true}
    ],page:{area:{r1:0,c1:0,r2:59,c2:5},fitW:1,fitH:0,scale:100,gridlines:true}}]});
    tabula.gv().layout();tabula.gv().renderAll();tabula.run('print');
  });
  await page.getByRole('dialog',{name:'인쇄 미리보기'}).waitFor(); assert.match(await preview.innerText(),/2쪽/);
  const firstClip = await preview.locator('[data-print-object="cross-page"]').evaluate(n => { const top=parseFloat(n.style.top), h=parseFloat(n.style.height), clip=parseFloat(n.parentElement.style.height); return Math.min(clip,top+h)-Math.max(0,top); });
  assert.ok(firstClip > 0 && firstClip < 200); assert.equal(await preview.locator('[data-print-object="outside-area"],[data-print-object="hidden-shape"]').count(),0);
  await preview.getByRole('button',{name:'다음 인쇄 페이지'}).click();
  const secondClip = await preview.locator('[data-print-object="cross-page"]').evaluate(n => { const top=parseFloat(n.style.top), h=parseFloat(n.style.height), clip=parseFloat(n.parentElement.style.height); return Math.min(clip,top+h)-Math.max(0,top); });
  assert.ok(Math.abs(firstClip + secondClip - 200)<1, `${firstClip}+${secondClip}`);
  await savePdf('D:/Codex/Temp/wixel-page-layout-cross-object.pdf');
  await page.getByRole('dialog',{name:'인쇄 미리보기'}).getByRole('button',{name:'닫기',exact:true}).last().click();
  results.push('명시 인쇄영역 내 개체 포함·2쪽 교차조각 높이 합200px·영역밖/숨김 제외·실제PDF');
  assert.deepEqual(errors, []); assert.deepEqual(writes, []);
  console.log(JSON.stringify({ cases: results.length, results, pdfPath, pdfBytes: bytes.length, pageErrors: errors, blockedWrites: writes }, null, 2));
} catch (error) { console.error(JSON.stringify({ errors, writes, dialogs: await page.locator('#dialogLayer').innerText() })); throw error; }
finally { await context.close(); await browser.close(); }
