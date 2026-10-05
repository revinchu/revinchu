// 도형 내부 편집 회귀: 격리된 합성 문서만 사용하며 모든 외부/API 쓰기를 차단한다.
import assert from 'node:assert/strict';
import { mkdir, readFile, writeFile } from 'node:fs/promises';
import { isAbsolute } from 'node:path';
import { pathToFileURL } from 'node:url';
const moduleName = process.env.PLAYWRIGHT_MODULE || 'playwright';
const pw = await import(isAbsolute(moduleName) ? pathToFileURL(moduleName).href : moduleName);
const engine = process.env.WIXEL_BROWSER || 'chromium';
const url = process.env.WIXEL_URL || 'http://localhost:5195/';
const origin = new URL(url).origin;
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname) && !(process.env.WIXEL_SHAPE_INLINE_PUBLIC === '1' && origin === 'https://wixel-3.wizx.workers.dev')) throw Error('격리 검사 주소를 지정하세요.');
const out = process.env.WIXEL_SHAPE_INLINE_OUT || `D:/Codex/Temp/wixel-shape-inline/current-${engine}`;
const baseline = process.env.WIXEL_SHAPE_INLINE_BASELINE === '1';
const filter = process.env.WIXEL_SHAPE_INLINE_FILTER || '';
await mkdir(out, { recursive: true });
const browser = await pw[engine].launch();
const results = [];
let checks = 0;
const eq = (a, b, message) => { checks++; assert.deepEqual(a, b, message); };
const ok = (value, message) => { checks++; assert.ok(value, message); };
const editor = p => p.getByRole('textbox', { name: '도형 텍스트 편집', exact: true });
const object = (p, id = 'inline-a') => p.locator(`.obj[data-id="${id}"]`).first();
const model = (p, si = 0) => p.evaluate(si => structuredClone(tabula.wb().sheets[si].shapes[0]), si);
const text = p => editor(p).evaluate(e => e.innerText.replace(/\r/g, ''));
const run = (p, command, arg) => p.evaluate(({ command, arg }) => tabula.run(command, arg), { command, arg });
const frame = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
async function fixture(p, { plain = false, small = false, kind = 'rect' } = {}) {
  await p.evaluate(({ plain, small, kind }) => {
    const t = tabula, w = t.wb();
    const shape = { id: 'inline-a', kind, x: small ? 14 : 100, y: 75, w: small ? 265 : 460, h: 160,
      text: 'Alpha Beta Gamma', fill: '#ffffff', stroke: '#4472c4', strokeWidth: 1,
      font: 'Arial', size: 15, color: '#112233', bold: true, italic: true, underline: true, strike: true,
      align: 'left', valign: 'top', pad: [10, 12, 10, 12] };
    if (!plain) shape.paras = [{ align: 'left', runs: [
      { t: 'Alpha ', b: true, i: true, u: true, s: true, sz: 15, font: 'Arial', color: '#112233' },
      { t: 'Beta Gamma', b: false, i: false, u: false, s: false, sz: 19, font: 'Arial', color: '#cc3322' },
    ] }];
    w.restore({ sheets: [
      { name: '합성 도형', cells: { '0,0': { raw: '셀 보존' }, '0,1': { raw: '83' } }, shapes: [shape] },
      { name: '다른 시트', cells: { '0,0': { raw: '다른 셀 보존' } }, shapes: [{ ...shape, text: 'Other shape', paras: undefined }] },
    ] });
    t.switchSheet(1); t.switchSheet(0); t.selectCell(0, 0);
    t.gv().setZoom(100); t.gv().setScroll(0, 0); t.gv().layout(); t.gv().renderAll();
    w.undoStack = []; w.redoStack = [];
    window.__inlineCellBefore = JSON.stringify([w.getRaw(0, 0, 0), w.getRaw(0, 0, 1), w.styleAt(0, 0, 0), w.styleAt(0, 0, 1)]);
  }, { plain, small, kind });
  await frame(p);
}
async function selectShape(p) { await object(p).click({ position: { x: 4, y: 4 } }); }
async function enter(p, how = 'F2') {
  if (how === 'dblclick') await object(p).dblclick({ position: { x: 35, y: 30 } });
  else if (how === 'context') {
    await object(p).click({ button: 'right', position: { x: 40, y: 30 } });
    await p.getByRole('menuitem', { name: /^텍스트 편집/ }).click();
  } else { await selectShape(p); await p.keyboard.press(how); }
  await editor(p).waitFor();
  eq(await p.locator('.shape-format-pane').count(), 0, '본문 편집은 패널을 열지 않음');
}
async function selectText(p, start, end = start) {
  await editor(p).evaluate((root, { start, end }) => {
    root.focus({ preventScroll: true });
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), nodes = [];
    for (let n; (n = walker.nextNode());) nodes.push(n);
    const at = offset => { for (const n of nodes) { if (offset <= n.length) return [n, offset]; offset -= n.length; } return [nodes.at(-1), nodes.at(-1)?.length || 0]; };
    const [a, ai] = at(start), [b, bi] = at(end), range = document.createRange();
    range.setStart(a, ai); range.setEnd(b, bi);
    const selection = window.getSelection(); selection.removeAllRanges(); selection.addRange(range);
    document.dispatchEvent(new Event('selectionchange'));
  }, { start, end });
}
async function styles(p, si = 0) {
  return p.evaluate(si => {
    const s = tabula.wb().sheets[si].shapes[0], chars = [];
    for (const paragraph of s.paras || [{ runs: [{ t: s.text || '' }] }]) {
      for (const r of paragraph.runs || []) for (const char of r.t || '') chars.push({ char,
        bold: r.b ?? !!s.bold, italic: r.i ?? !!s.italic, underline: r.u ?? !!s.underline, strike: r.s ?? !!s.strike,
        font: r.font ?? s.font, size: r.sz ?? s.size, color: r.color ?? s.color, align: paragraph.align ?? s.align });
    }
    return chars;
  }, si);
}
async function cellsUnchanged(p) {
  eq(await p.evaluate(() => {
    const w = tabula.wb();
    return JSON.stringify([w.getRaw(0, 0, 0), w.getRaw(0, 0, 1), w.styleAt(0, 0, 0), w.styleAt(0, 0, 1)]) === window.__inlineCellBefore;
  }), true, '아래 셀 값과 서식 불변');
}
async function exit(p) { await editor(p).focus(); await p.keyboard.press('Escape'); await editor(p).waitFor({ state: 'detached' }); }
async function test(name, fn, { mobile = false } = {}) {
  if (filter && !filter.split('|').some(part => name.includes(part))) return;
  const context = await browser.newContext({ viewport: mobile ? { width: 390, height: 844 } : { width: 1440, height: 960 }, serviceWorkers: 'block', acceptDownloads: true });
  const p = await context.newPage(), errors = [], writes = [], start = checks;
  p.setDefaultTimeout(10000); p.on('pageerror', error => errors.push(error.message));
  await context.addInitScript(mobile => {
    window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
    localStorage.setItem('wixel.mobile-work.v1', mobile ? 'on' : 'off');
    Object.defineProperty(window, 'showSaveFilePicker', { configurable: true, value: undefined });
  }, mobile);
  await context.route('**/*', route => {
    const request = route.request(), target = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.url()); return route.abort(); }
    return target.origin === origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await p.waitForFunction(() => window.tabula?.gv());
    await fixture(p, { small: mobile });
    await fn(p);
    eq(errors, [], '페이지 오류 없음'); eq(writes, [], '외부 쓰기 없음');
    results.push({ name, ok: true, checks: checks - start }); console.log('OK ' + name);
  } catch (error) {
    results.push({ name, ok: false, checks: checks - start, error: error.stack, errors, writes });
    console.error('NG ' + name + ': ' + error.message);
    await p.screenshot({ path: `${out}/failure-${results.length}.png` }).catch(() => {});
  } finally { await context.close(); }
}
try {
  if (baseline) {
    for (const how of ['F2', 'dblclick']) await test('baseline-' + how, async p => {
      if (how === 'F2') { await selectShape(p); await p.keyboard.press('F2'); }
      else await object(p).dblclick();
      await p.locator('.shape-format-pane').waitFor();
      eq(await editor(p).count(), 0, '수정 전에는 내부 커서 없이 서식 패널만 열림');
      await p.screenshot({ path: `${out}/baseline-${how}.png` });
    });
  } else {
    await test('entry-routes', async p => {
      for (const how of ['F2', 'dblclick', 'context']) {
        await enter(p, how); eq(await editor(p).getAttribute('contenteditable'), 'true', '편집 가능한 본문');
        await exit(p); eq(await p.evaluate(() => tabula.wb().undoStack.length), 0, '무수정 종료는 Undo 없음');
      }
      await enter(p, 'x'); ok((await text(p)).includes('x'), '선택한 도형에서 바로 문자 입력');
      await exit(p); await cellsUnchanged(p);
    });
    await test('mouse-caret-and-selection', async p => {
      await enter(p);
      const point = await editor(p).evaluate(root => {
        const node = document.createTreeWalker(root, NodeFilter.SHOW_TEXT).nextNode(), r = document.createRange();
        r.setStart(node, 3); r.setEnd(node, 4); const b = r.getBoundingClientRect(); return { x: b.left + 1, y: b.top + b.height / 2 };
      });
      await p.mouse.click(point.x, point.y);
      const caret = await editor(p).evaluate(root => { const s = getSelection(), r = document.createRange(); r.selectNodeContents(root); r.setEnd(s.anchorNode, s.anchorOffset); return r.toString().length; });
      ok(caret > 0 && caret < 6, '마우스로 문자열 중간에 커서 배치');
      await p.keyboard.insertText('|');
      const first = 'Alpha Beta Gamma'.slice(0, caret) + '|' + 'Alpha Beta Gamma'.slice(caret);
      eq(await text(p), first, '커서 위치 삽입');
      const bounds = await editor(p).evaluate(root => {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT), nodes = []; for (let n; (n = walker.nextNode());) nodes.push(n);
        const at = offset => { for (const n of nodes) { if (offset <= n.length) return [n, offset]; offset -= n.length; } };
        const offset = root.textContent.indexOf('Beta'), [a, ai] = at(offset), [b, bi] = at(offset + 4), r = document.createRange(); r.setStart(a, ai); r.setEnd(b, bi);
        const box = r.getBoundingClientRect(); return { left: box.left, right: box.right, y: box.top + box.height / 2 };
      });
      await p.mouse.move(bounds.left + 1, bounds.y); await p.mouse.down(); await p.mouse.move(bounds.right - 1, bounds.y, { steps: 8 }); await p.mouse.up();
      const selected = await editor(p).evaluate(root => { const s = getSelection(), r = s.getRangeAt(0), prefix = document.createRange(); prefix.selectNodeContents(root); prefix.setEnd(r.startContainer, r.startOffset); return { start: prefix.toString().length, length: r.toString().length }; });
      ok(selected.length >= 2, '마우스 드래그로 글자 범위 선택');
      await p.keyboard.insertText('Z'); await exit(p);
      eq((await model(p)).text, first.slice(0, selected.start) + 'Z' + first.slice(selected.start + selected.length), '마우스 선택 범위만 치환');
      await enter(p); const frameBox = await p.locator('.shape-text-editor').boundingBox();
      await p.mouse.click(frameBox.x + frameBox.width - 3, frameBox.y + frameBox.height - 3);
      eq(await editor(p).evaluate(e => document.activeElement === e), true, '빈 테두리 여백 클릭 후 본문 초점 유지');
      await p.keyboard.insertText(' blank'); await exit(p);
      ok((await model(p)).text.includes(' blank'), '빈 여백 클릭 뒤에도 도형 본문에 입력'); await cellsUnchanged(p);
    });
    await test('keyboard-partial-format', async p => {
      const before = await styles(p); await enter(p); await selectText(p, 6, 10);
      await p.keyboard.press('Control+b'); await p.keyboard.press('Control+i'); await p.keyboard.press('Control+u');
      for (const [key, code] of [['ㅠ', 'KeyB'], ['ㅑ', 'KeyI'], ['ㅕ', 'KeyU']]) {
        await editor(p).dispatchEvent('keydown', { key, code, ctrlKey: true, bubbles: true, cancelable: true });
      }
      eq(await editor(p).evaluate(e => {
        const walker = document.createTreeWalker(e, NodeFilter.SHOW_TEXT);
        let remaining = 6, node;
        while ((node = walker.nextNode())) { if (remaining < node.length) break; remaining -= node.length; }
        const css = getComputedStyle(node.parentElement);
        return [css.fontWeight, css.fontStyle, css.textDecorationLine];
      }), ['400', 'normal', 'none'], '한글 자판 물리 키 B/I/U는 선택 글자 서식 해제');
      for (const [key, code] of [['ㅠ', 'KeyB'], ['ㅑ', 'KeyI'], ['ㅕ', 'KeyU']]) {
        await editor(p).dispatchEvent('keydown', { key, code, metaKey: true, bubbles: true, cancelable: true });
      }
      await exit(p); const after = await styles(p);
      eq(after.slice(0, 6), before.slice(0, 6), '앞쪽 서식 유지'); eq(after.slice(10), before.slice(10), '뒤쪽 서식 유지');
      for (const c of after.slice(6, 10)) eq([c.bold, c.italic, c.underline], [true, true, true], '선택한 글자만 B/I/U');
      await cellsUnchanged(p);
    });
    await test('ribbon-partial-font-size-color', async p => {
      const before = await styles(p); await enter(p); await selectText(p, 6, 10);
      await p.locator('[data-ribbon-tab="home"]').click();
      await p.locator('.font-family').fill('Courier New'); await p.locator('.font-family').press('Enter');
      await p.locator('.font-size').fill('24'); await p.locator('.font-size').press('Enter');
      await p.getByRole('button', { name: '글꼴 색 선택', exact: true }).click();
      const swatch = p.locator('.color-palette-menu .swatch').last(), color = await swatch.getAttribute('data-color');
      await swatch.click(); await p.screenshot({ path: `${out}/desktop-partial-format.png` }); await exit(p); const after = await styles(p);
      eq(after.slice(0, 6), before.slice(0, 6), '리본 입력 중 앞쪽 서식 유지'); eq(after.slice(10), before.slice(10), '리본 입력 중 뒤쪽 서식 유지');
      for (const c of after.slice(6, 10)) eq([c.font, c.size, c.color], ['Courier New', 24, color], '선택 범위만 글꼴·크기·색');
      await cellsUnchanged(p);
    });
    await test('newline-composition-plain-paste', async p => {
      await enter(p); await selectText(p, 0, 16); await p.keyboard.insertText('First'); await p.keyboard.press('Enter');
      await editor(p).dispatchEvent('compositionstart', { data: '' });
      await p.keyboard.insertText('한글'); await editor(p).dispatchEvent('keydown', { key: 'Enter', code: 'Enter', isComposing: true, keyCode: 229 });
      eq(await editor(p).count(), 1, '조합 중 Enter는 편집 종료 아님');
      await editor(p).dispatchEvent('compositionend', { data: '한글' });
      await editor(p).evaluate(e => { const data = new DataTransfer(); data.setData('text/plain', '<plain>'); data.setData('text/html', '<img src="https://invalid.example/secret">rich'); const event = new Event('paste', { bubbles: true, cancelable: true }); Object.defineProperty(event, 'clipboardData', { value: data }); e.dispatchEvent(event); });
      eq(await editor(p).locator('img,script').count(), 0, '붙여넣기 HTML 삽입 없음');
      await exit(p); eq((await model(p)).text, 'First\n한글<plain>', '줄바꿈·조합·일반 텍스트 저장'); await cellsUnchanged(p);
    });
    await test('local-undo-single-workbook-undo', async p => {
      const before = await model(p); await enter(p); await selectText(p, 0, 16); await p.keyboard.insertText('Changed');
      await p.keyboard.press('Control+z'); eq(await text(p), before.text, '편집 중 로컬 Undo');
      await p.keyboard.press('Control+y'); eq(await text(p), 'Changed', '편집 중 로컬 Redo');
      await editor(p).dispatchEvent('keydown', { key: 'ㅋ', code: 'KeyZ', ctrlKey: true, bubbles: true, cancelable: true });
      eq(await text(p), before.text, '한글 자판 물리 키 Ctrl+Z 로컬 Undo');
      await editor(p).dispatchEvent('keydown', { key: 'ㅛ', code: 'KeyY', metaKey: true, bubbles: true, cancelable: true });
      eq(await text(p), 'Changed', '한글 자판 물리 키 Command+Y 로컬 Redo');
      eq(await p.evaluate(() => tabula.wb().undoStack.length), 0, '입력 도중 통합 문서 Undo 누적 없음');
      await exit(p); eq(await p.evaluate(() => tabula.wb().undoStack.length), 1, '한 편집 세션은 한 번 Undo');
      await run(p, 'undo'); eq(await model(p), before, '통합 문서 Undo는 원래 리치텍스트 복원');
      await cellsUnchanged(p);
    });
    await test('protection-and-line-guard', async p => {
      const before = await model(p);
      await p.evaluate(() => tabula.wb().setSheetProp(0, 'protect', { on: true, allow: { objects: false, selectLocked: true, selectUnlocked: true } }));
      await selectShape(p); await p.keyboard.press('F2'); eq(await editor(p).count(), 0, '보호된 도형 편집 차단');
      if (await p.getByRole('dialog').count()) await p.getByRole('dialog').locator('.dialog-head button').click();
      await p.evaluate(() => tabula.wb().setSheetProp(0, 'protect', undefined));
      await enter(p); await selectText(p, 0, 16); await p.keyboard.insertText('Blocked');
      await p.evaluate(() => { tabula.wb().props.markedFinal = true; }); await p.keyboard.press('Escape');
      eq(await model(p), before, '편집 중 읽기 전용으로 바뀌면 쓰기 차단');
      await p.evaluate(() => { tabula.wb().props.markedFinal = false; });
      await fixture(p, { kind: 'line' }); await selectShape(p); await p.keyboard.press('F2');
      eq(await editor(p).count(), 0, '선은 본문 편집 대상 아님'); await cellsUnchanged(p);
    });
    await test('sheet-document-target-isolation', async p => {
      await enter(p); await selectText(p, 0, 16); await p.keyboard.insertText('Old draft');
      await p.evaluate(() => { window.__staleInline = document.querySelector('.shape-text-editor-content'); tabula.switchSheet(1); window.__postSwitch = JSON.stringify(tabula.wb().sheets.map(s => s.shapes)); });
      await p.evaluate(() => { const e = window.__staleInline; e.textContent = 'Late input'; e.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText', data: 'Late input' })); e.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' })); });
      eq(await p.evaluate(() => JSON.stringify(tabula.wb().sheets.map(s => s.shapes)) === window.__postSwitch), true, '지난 편집기의 이벤트가 시트 변경 후 쓰지 않음');
      eq((await model(p, 1)).text, 'Other shape', '다른 시트의 같은 ID 도형 보존');
      await p.evaluate(() => tabula.switchSheet(0)); await enter(p); await p.keyboard.insertText(' more');
      await p.evaluate(async () => { window.__oldInline = document.querySelector('.shape-text-editor-content'); await tabula.newWorkbook({ name: '새 합성 문서', build: () => ({ sheets: [{ name: '새 시트', cells: { '0,0': { raw: 'NEW CELL' } }, shapes: [{ id: 'inline-a', kind: 'rect', x: 100, y: 80, w: 200, h: 100, text: 'NEW SHAPE' }] }] }) }); });
      if (await p.getByRole('button', { name: '새로 만들기', exact: true }).count()) await p.getByRole('button', { name: '새로 만들기', exact: true }).click();
      await p.waitForFunction(() => tabula.wb().sheets[0].name === '새 시트');
      await p.evaluate(() => { const e = window.__oldInline; e.textContent = 'Late document input'; e.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertText' })); e.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, key: 'Escape' })); });
      eq((await model(p)).text, 'NEW SHAPE', '다른 문서의 같은 ID 도형 보존');
      eq(await p.evaluate(() => tabula.wb().getRaw(0, 0, 0)), 'NEW CELL', '새 문서 셀 보존');
    });
    await test('scroll-zoom-frozen-pane-follow', async p => {
      await p.evaluate(() => { const w = tabula.wb(); w.sheets[0].shapes[0].x = 280; w.sheets[0].shapes[0].y = 160; w.setSheetProp(0, 'freeze', { rows: 2, cols: 2 }); tabula.gv().layout(); tabula.gv().renderAll(); w.undoStack = []; });
      await enter(p);
      const geometry = async () => {
        await frame(p);
        const b = await editor(p).boundingBox(), a = await object(p).boundingBox();
        ok(b && a && b.x < a.x + a.width && b.x + b.width > a.x && b.y < a.y + a.height && b.y + b.height > a.y, '편집기가 실제 도형과 겹침');
        return { b, a };
      };
      const before = await geometry();
      await p.evaluate(() => { const v = tabula.gv(); v.setScroll(45, 30); v.layout(); v.renderAll(); });
      const scrolled = await geometry(); ok(Math.abs(scrolled.b.x - before.b.x) > 5 || Math.abs(scrolled.b.y - before.b.y) > 5, '스크롤에 따라 편집기 이동');
      await p.evaluate(() => { const v = tabula.gv(); v.setZoom(70); v.layout(); v.renderAll(); });
      const zoomed = await geometry(); ok(zoomed.b.width < scrolled.b.width * .9, '시트 배율에 따라 본문 편집 크기 조정');
      await selectText(p, 0, 5); await p.keyboard.insertText('Zoom'); await exit(p);
      ok((await model(p)).text.startsWith('Zoom'), '이동·배율 변경 뒤 편집 가능'); await cellsUnchanged(p);
    });
    await test('panel-preserves-mixed-rich-text', async p => {
      const before = await styles(p); await selectShape(p); await run(p, 'shapeFormat');
      const pane = p.locator('.shape-format-pane'); await pane.getByRole('tab', { name: '텍스트 옵션', exact: true }).click();
      const family = pane.getByLabel('도형 글꼴', { exact: true }); await family.fill('Courier New'); await family.dispatchEvent('change');
      const afterFont = await styles(p); eq(afterFont.map(c => ({ ...c, font: 'Arial' })), before, '패널 글꼴 변경은 다른 리치 서식 보존');
      ok(afterFont.every(c => c.font === 'Courier New'), '글꼴만 균일 변경');
      const input = pane.getByLabel('도형 텍스트', { exact: true }); await input.fill('Alpha Beta Gamma'); await input.dispatchEvent('change');
      eq(await styles(p), afterFont, '같은 본문 입력은 리치텍스트 유지');
      await pane.getByRole('tab', { name: '텍스트 상자', exact: true }).click(); await pane.getByLabel('가로 맞춤', { exact: true }).selectOption('right');
      const aligned = await styles(p); eq(aligned.map(c => ({ ...c, align: 'left' })), afterFont, '맞춤만 변경해도 각 글자 서식 유지');
      const rendered = await object(p).locator('.sh-text span').evaluateAll(nodes => nodes.filter(e => e.textContent.includes('Beta')).map(e => { const s = getComputedStyle(e); return [s.fontWeight, s.fontStyle, s.textDecorationLine]; }));
      ok(rendered.some(x => x[0] === '400' && x[1] === 'normal' && x[2] === 'none'), 'false 리치서식은 도형의 굵기·기울임·밑줄을 상속하지 않음');
      await cellsUnchanged(p);
    });
    await test('xlsx-download-reopen-rich-text', async p => {
      await enter(p); await selectText(p, 6, 10); await p.keyboard.press('Control+b');
      await p.locator('[data-ribbon-tab="home"]').click(); await p.locator('.font-size').fill('27'); await p.locator('.font-size').press('Enter');
      await exit(p); const before = await styles(p), beforeText = (await model(p)).text;
      await p.evaluate(() => { void tabula.exportXlsx('합성-도형-내부편집', 'xlsx'); });
      const save = p.getByRole('dialog', { name: '파일로 저장', exact: true }); await save.waitFor();
      const [download] = await Promise.all([p.waitForEvent('download'), save.getByRole('button', { name: /^다운로드/ }).click()]);
      const path = `${out}/synthetic-inline-rich.xlsx`; await download.saveAs(path); ok((await readFile(path)).length > 1000, '실제 XLSX 파일 다운로드');
      await p.evaluate(() => { window.__beforeReopen = tabula.wb(); }); await p.locator('#fileInput').setInputFiles(path);
      await p.waitForFunction(() => tabula.wb() !== window.__beforeReopen, null, { timeout: 60000 });
      await p.waitForFunction(() => !document.querySelector('.load-progress'));
      eq((await model(p)).text, beforeText, '다운로드·재열기 후 본문 동일'); eq(await styles(p), before, '부분 서식과 false 서식 왕복 유지');
      eq(await p.evaluate(() => tabula.wb().getRaw(0, 0, 0)), '셀 보존', '저장 뒤 아래 셀 보존');
    });
    await test('clipboard-toolbar-and-late-paste', async p => {
      await p.evaluate(() => {
        window.__shapeClipboard = { value: '', reads: 0, writes: 0, denyRead: false, denyWrite: false, deferred: false };
        Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
          writeText: async value => { const c = window.__shapeClipboard; c.writes++; if (c.denyWrite) throw new DOMException('denied', 'NotAllowedError'); c.value = value; },
          readText: async () => { const c = window.__shapeClipboard; c.reads++; if (c.denyRead) throw new DOMException('denied', 'NotAllowedError'); if (c.deferred) return new Promise(resolve => { window.__resolveShapePaste = resolve; }); return c.value; },
        } });
      });
      await enter(p); await selectText(p, 6, 10); await p.locator('[data-ribbon-tab="home"]').click();
      const copy = p.locator('button[title^="복사 (Ctrl+C)"]'), cut = p.locator('button[title^="잘라내기 (Ctrl+X)"]'), paste = p.locator('button[title^="붙여넣기"]').first();
      await copy.click({ position: { x: 8, y: 8 } }); await p.waitForFunction(() => window.__shapeClipboard.value === 'Beta');
      eq(await text(p), 'Alpha Beta Gamma', '리본 복사는 선택 텍스트만 복사');
      await cut.click({ position: { x: 8, y: 8 } }); await p.waitForFunction(() => document.querySelector('.shape-text-editor-content')?.innerText === 'Alpha  Gamma');
      eq(await p.evaluate(() => tabula.wb().sheets[0].shapes.length), 1, '리본 잘라내기가 도형을 지우지 않음');
      await p.evaluate(() => { window.__shapeClipboard.value = 'Inserted'; });
      await paste.click({ position: { x: 8, y: 8 } }); await p.waitForFunction(() => document.querySelector('.shape-text-editor-content')?.innerText.includes('Inserted'));
      eq(await text(p), 'Alpha Inserted Gamma', '리본 붙여넣기는 텍스트 커서에 적용');
      await selectText(p, 0, 5); await p.evaluate(() => { window.__shapeClipboard.denyWrite = true; });
      const writes = await p.evaluate(() => window.__shapeClipboard.writes);
      await cut.click({ position: { x: 8, y: 8 } }); await p.waitForFunction(n => window.__shapeClipboard.writes > n, writes); await frame(p);
      eq(await text(p), 'Alpha Inserted Gamma', '클립보드 쓰기 거절 시 잘라낼 본문 보존');
      await p.evaluate(() => { window.__shapeClipboard.denyRead = true; }); const reads = await p.evaluate(() => window.__shapeClipboard.reads);
      await paste.click({ position: { x: 8, y: 8 } }); await p.waitForFunction(n => window.__shapeClipboard.reads > n, reads); await frame(p);
      eq(await text(p), 'Alpha Inserted Gamma', '클립보드 읽기 거절 시 본문 보존');
      await p.evaluate(() => { window.__shapeClipboard.denyRead = false; window.__shapeClipboard.deferred = true; });
      await paste.click({ position: { x: 8, y: 8 } }); await p.waitForFunction(() => typeof window.__resolveShapePaste === 'function');
      await exit(p); const committed = await model(p);
      await p.evaluate(() => window.__resolveShapePaste('LATE')); await frame(p);
      eq(await model(p), committed, '편집 종료 후 늦게 도착한 붙여넣기 무시');
      await cellsUnchanged(p);
    });
    await test('alt-keytip-font-color-selection', async p => {
      const before = await styles(p); await enter(p); await selectText(p, 6, 10);
      await p.keyboard.press('Alt'); for (const key of ['h', 'f', 'c']) await p.keyboard.press(key);
      await p.locator('.color-palette-menu').waitFor(); await p.keyboard.press('ArrowDown'); await p.keyboard.press('ArrowRight');
      const color = await p.evaluate(() => document.activeElement?.getAttribute('data-color'));
      ok(/^#[0-9a-f]{6}$/i.test(color || ''), 'Alt H F C 뒤 화살표로 색상 선택');
      await p.keyboard.press('Enter'); eq(await editor(p).count(), 1, '색 적용 후 본문 편집 유지');
      await p.keyboard.press('F10'); for (const key of ['h', 'f', 'c']) await p.keyboard.press(key);
      await p.locator('.color-palette-menu').waitFor(); await p.keyboard.press('Escape');
      eq(await editor(p).count(), 1, 'F10 팔레트 취소 뒤 편집 유지'); await exit(p);
      const after = await styles(p); eq(after.slice(0, 6), before.slice(0, 6), '키팁 서식도 앞쪽 보존'); eq(after.slice(10), before.slice(10), '키팁 서식도 뒤쪽 보존');
      for (const c of after.slice(6, 10)) eq(c.color, color, '키팁 색은 선택한 글자에만 적용'); await cellsUnchanged(p);
    });
    await test('composition-frame-race', async p => {
      await enter(p); await selectText(p, 0, 16); await p.keyboard.insertText('First');
      await p.evaluate(() => { window.__inlineNativeRAF = requestAnimationFrame; window.__inlineFrames = []; window.requestAnimationFrame = callback => (window.__inlineFrames.push(callback), window.__inlineFrames.length); });
      try {
        await editor(p).dispatchEvent('compositionstart', { data: '' }); await editor(p).dispatchEvent('compositionend', { data: '' });
        await p.keyboard.press('Enter'); await editor(p).dispatchEvent('compositionstart', { data: '' });
        const queued = await p.evaluate(() => { const callbacks = window.__inlineFrames.splice(0); for (const callback of callbacks) callback(performance.now()); window.requestAnimationFrame = window.__inlineNativeRAF; return callbacks.length; });
        ok(queued > 0, '이전 조합 종료의 예약 프레임을 다음 조합 중 강제로 실행');
        await p.keyboard.insertText('한글'); await editor(p).dispatchEvent('compositionend', { data: '한글' }); await frame(p);
        eq(await text(p), 'First\n한글', '예약 렌더가 다음 조합의 커서를 앞으로 옮기지 않음');
        await exit(p); eq((await model(p)).text, 'First\n한글', '조합 경합 후 문단 순서와 입력 보존'); await cellsUnchanged(p);
      } finally { await p.evaluate(() => { if (window.__inlineNativeRAF) window.requestAnimationFrame = window.__inlineNativeRAF; }); }
    });
    await test('compact-mobile-editor', async p => {
      await enter(p); await selectText(p, 6, 10); await p.keyboard.insertText('Mobile');
      const box = await editor(p).boundingBox(); ok(box.x < 390 && box.x + box.width > 0 && box.y < 844 && box.y + box.height > 0, '작은 화면에서 본문 편집 접근 가능');
      await p.screenshot({ path: `${out}/mobile-editor.png` }); await exit(p);
      eq((await model(p)).text, 'Alpha Mobile Gamma', '모바일 배치에서도 선택 부분만 수정'); await cellsUnchanged(p);
    }, { mobile: true });
  }
} finally {
  await browser.close();
  const summary = { engine, url, baseline, cases: results.length, passed: results.filter(r => r.ok).length, checks, results };
  await writeFile(`${out}/result.json`, JSON.stringify(summary, null, 2));
  console.log(JSON.stringify({ engine, cases: summary.cases, passed: summary.passed, checks, out }));
  if (!results.length || summary.passed !== summary.cases) process.exitCode = 1;
}
