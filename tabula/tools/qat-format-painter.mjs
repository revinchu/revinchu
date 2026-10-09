// 빠른 실행 서식 복사: 실제 Alt 키·마우스 클릭으로 값/서식/Undo를 검증한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const engines = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.BROWSER || 'chromium';
if (!['chromium', 'webkit'].includes(engine)) throw Error('BROWSER는 chromium 또는 webkit이어야 합니다.');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const repo = fileURLToPath(new URL('../../', import.meta.url));
const baseline = process.env.QAT_BASELINE_REF;
const baselineApp = baseline ? execFileSync('git', ['-c', `safe.directory=${repo.replace(/\\/g, '/').replace(/\/$/, '')}`, 'show', `${baseline}:tabula/src/app.js`], { cwd: repo, encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 }) : null;
const out = path.resolve(process.env.QAT_OUT || '.local/qat-format-painter/' + (baseline ? 'baseline-' : '') + engine);
await mkdir(out, { recursive: true });
const browser = await engines[engine].launch(), results = [];
const frame = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const state = p => p.evaluate(() => {
  const t = tabula, w = t.wb(), e = document.getElementById('cellEditor');
  return { cells: w.serialize().sheets.map(s => s.cells), undo: w.undoStack.length, redo: w.redoStack.length, painting: document.getElementById('gridView').classList.contains('painting'), active: { ...t.active }, selection: { ...t.sel }, editing: !e.classList.contains('idle'), editor: e.value, focus: document.activeElement.id || document.activeElement.tagName, keytips: document.body.classList.contains('keytips') };
});
async function fixture(p) {
  await p.evaluate(() => {
    const t = tabula, w = t.wb();
    w.restore({ sheets: [{ name: '서식 복사 합성', cells: {
      '0,0': { raw: '12.34', style: { bold: true, italic: true, fill: '#fff2cc', color: '#123456', numFmt: 'number', decimals: 3, bt: true, bb: true, bl: true, br: true, bts: 'double', bbs: 'double', bls: 'double', brs: 'double', btc: '#112233', bbc: '#112233', blc: '#112233', brc: '#112233' } },
      '2,2': { raw: '98.76', style: { fill: '#ddeeff', numFmt: 'percent' }, comment: '보존할 메모' },
      '4,4': { raw: '=7*11', style: { italic: false, fill: '#ccffaa' }, cached: 77 },
      '6,6': { raw: '유지' },
    } }, { name: '다른 시트', cells: { '0,0': { raw: '둘째' } } }] });
    t.switchSheet(1); t.switchSheet(0); w.undoStack = []; w.redoStack = [];
    t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
    window.__qatEvents = [];
    if (!window.__qatRecording) {
      window.__qatRecording = true;
      for (const type of ['keydown', 'keyup', 'mousedown', 'mouseup', 'click', 'beforeinput', 'input', 'compositionstart', 'compositionend']) document.addEventListener(type, e => {
        const rec = { type, key: e.key, code: e.code, target: e.target.id || e.target.className, alt: e.altKey, repeat: e.repeat, canceled: e.defaultPrevented, painting: document.getElementById('gridView').classList.contains('painting') };
        if (window.__qatEvents.length < 100) window.__qatEvents.push(rec);
      });
    }
  });
  await clickCell(p, 0, 0);
  assert.equal((await state(p)).focus, 'cellEditor');
}
async function cellPoint(p, r, c) {
  return p.evaluate(([r, c]) => { const b = tabula.gv().clientRect({ r1: r, c1: c, r2: r, c2: c }); return { x: (b.left + b.right) / 2, y: (b.top + b.bottom) / 2 }; }, [r, c]);
}
async function clickCell(p, r, c) {
  const a = await cellPoint(p, r, c); await p.mouse.click(a.x, a.y); await frame(p);
}
async function assertApplied(p, before, r = 2, c = 2) {
  const after = await state(p), key = `${r},${c}`;
  assert.deepEqual(after.active, { r, c }, '실제 대상 셀 선택');
  assert.deepEqual(after.cells[0][key].style, before.cells[0]['0,0'].style, '원본 전체 셀 서식 복사');
  for (const k of Object.keys(before.cells[0])) assert.equal(after.cells[0][k].raw, before.cells[0][k].raw, '셀 값·수식 보존');
  assert.equal(after.cells[0]['2,2'].comment, before.cells[0]['2,2'].comment, '메모 보존');
  assert.deepEqual(after.cells[1], before.cells[1], '다른 시트 보존');
  assert.equal(after.undo, before.undo + 1, '한 번의 실행 취소 기록');
  assert.equal(after.editing, false, '셀 편집 모드가 아님');
  assert.equal(after.editor, '', '단축키 문자가 셀 에디터에 남지 않음');
  return after;
}
async function undoRedo(p, before, after) {
  await p.keyboard.press('Control+z'); await frame(p); assert.deepEqual((await state(p)).cells, before.cells, '실행 취소');
  await p.keyboard.press('Control+y'); await frame(p); assert.deepEqual((await state(p)).cells, after.cells, '다시 실행');
}
async function altNumber(p, digit = '1', sequential = false, repeat = false) {
  if (sequential) { await p.keyboard.press('Alt'); await p.keyboard.press('Digit' + digit); }
  else { await p.keyboard.down('Alt'); await p.keyboard.down('Digit' + digit); if (repeat) await p.keyboard.down('Digit' + digit); await p.keyboard.up('Digit' + digit); await p.keyboard.up('Alt'); }
  await frame(p);
}
async function test(name, fn, options = {}) {
  if (process.env.QAT_TEST_FILTER && !name.includes(process.env.QAT_TEST_FILTER)) return;
  const context = await browser.newContext({ viewport: { width: 1450, height: 950 } }), p = await context.newPage(), errors = [], writes = [], diagnostics = [], info = {};
  p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message)); p.on('console', m => {
    if (m.type() !== 'error') return;
    const message = m.text();
    // WebKit의 미지원 viewport 힌트는 실행 오류와 구분하되 결과에 그대로 남긴다.
    if (engine === 'webkit' && message === 'Viewport argument key "interactive-widget" not recognized and ignored.') diagnostics.push(message);
    else errors.push(message);
  }); p.on('dialog', d => d.dismiss());
  await context.route('**/*', route => {
    const req = route.request(), u = new URL(req.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.pathname); return route.abort(); }
    if (u.origin === new URL(url).origin && u.pathname === '/api/health') return route.fulfill({ json: { ok: true, vault: false, auth: false, publish: false } });
    return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  if (baselineApp) await context.route('**/src/app.js', r => r.fulfill({ contentType: 'text/javascript', body: baselineApp }));
  try {
    await context.addInitScript(options => { window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; if (options.qatOrder) localStorage.setItem('wixel.options', JSON.stringify({ qatOrder: options.qatOrder, qatPosition: 'above' })); }, options);
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    await fixture(p); await fn(p, info); assert.deepEqual(errors, [], '브라우저 오류'); assert.deepEqual(writes, [], '외부 쓰기 시도');
    results.push({ name, ok: true, diagnostics, info }); console.log('OK ' + name);
  } catch (e) {
    info.state = await state(p).catch(() => null); info.events = await p.evaluate(() => window.__qatEvents).catch(() => null);
    results.push({ name, ok: false, error: e.message, errors, writes, diagnostics, info }); console.error('NG ' + name + ': ' + e.message);
    await p.screenshot({ path: path.join(out, name + '.png') }).catch(() => {});
  } finally { await context.close(); }
}
try {
  const basicRoutes = ['qat-click', 'ribbon-click', 'alt-held', 'alt-sequential', 'alt-repeat'];
  if (process.env.QAT_NUMPAD === '1') basicRoutes.push('alt-numpad-held', 'alt-numpad-sequential');
  for (const route of basicRoutes) await test(route, async (p, info) => {
    const before = await state(p);
    if (route === 'qat-click') await p.locator('#quickAccess [data-qat-cmd="painter"]').click();
    else if (route === 'ribbon-click') await p.locator('#ribbon [data-ribbon-command="painter"]').click();
    else if (route === 'alt-numpad-held') await p.keyboard.press('Alt+Numpad1');
    else if (route === 'alt-numpad-sequential') { await p.keyboard.press('Alt'); await p.keyboard.press('Numpad1'); }
    else await altNumber(p, '1', route === 'alt-sequential', route === 'alt-repeat');
    info.armed = await state(p); assert.equal(info.armed.painting, true, '서식 복사 모드 시작'); assert.deepEqual(info.armed.cells, before.cells, '시작만으로 셀 변경 없음'); assert.equal(info.armed.undo, before.undo);
    await clickCell(p, 2, 2); const after = await assertApplied(p, before); assert.equal(after.painting, false, '한 번 적용 후 종료');
    await clickCell(p, 4, 4); assert.deepEqual((await state(p)).cells, after.cells, '다음 일반 클릭에는 서식 미적용');
    await undoRedo(p, before, after); info.events = await p.evaluate(() => window.__qatEvents);
  });
  await test('escape-cancel', async p => { const before = await state(p); await altNumber(p); await p.keyboard.press('Escape'); assert.equal((await state(p)).painting, false); await clickCell(p, 2, 2); assert.deepEqual((await state(p)).cells, before.cells); assert.equal((await state(p)).undo, 0); });
  for (const route of ['qat-click', 'alt-held', 'alt-sequential']) for (const editing of ['f2', 'typing', 'formula-bar']) await test('editing-' + editing + '-' + route, async (p, info) => {
    const initial = await state(p);
    if (editing === 'f2') await p.keyboard.press('F2');
    else if (editing === 'typing') await p.keyboard.type('25.5');
    else { await p.locator('#formulaInput').focus(); await p.keyboard.press('Control+a'); await p.keyboard.type('25.5'); }
    info.beforeKey = await state(p); assert.equal(info.beforeKey.editing, true, '실제 편집 시작');
    if (route === 'qat-click') await p.locator('#quickAccess [data-qat-cmd="painter"]').click();
    else await altNumber(p, '1', route === 'alt-sequential');
    info.armed = await state(p); assert.equal(info.armed.painting, true, '편집 확정 후 서식 복사 모드'); assert.equal(info.armed.editing, false, '편집 확정');
    const expected = structuredClone(initial);
    if (editing !== 'f2') { expected.cells[0]['0,0'].raw = '25.5'; expected.undo++; }
    assert.deepEqual(info.armed.cells, expected.cells, '입력값 확정 및 다른 셀 보존');
    await clickCell(p, 2, 2); const after = await assertApplied(p, expected); await undoRedo(p, expected, after);
  });
  for (const route of ['qat-click', 'alt-held']) await test('invalid-formula-' + route, async p => {
    const before = await state(p); await p.keyboard.type('=1+*2');
    if (route === 'qat-click') await p.locator('#quickAccess [data-qat-cmd="painter"]').click(); else await altNumber(p);
    await p.getByRole('dialog').waitFor(); assert.match(await p.getByRole('dialog').innerText(), /입력한 수식에 문제가/);
    const after = await state(p); assert.deepEqual(after.cells, before.cells); assert.equal(after.undo, 0); assert.equal(after.painting, false, '잘못된 수식을 확정하지 않고 모드 시작 차단'); assert.equal(after.editing, true);
    await p.getByRole('dialog').getByRole('button', { name: /^확인/ }).click(); await p.keyboard.press('Escape');
  });
  for (const route of ['qat-click', 'alt-held']) await test('validation-stop-' + route, async p => {
    await p.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setSheetProp(0, 'validations', [{ r1: 0, r2: 0, c1: 0, c2: 0, type: 'whole', op: 'between', f1: '1', f2: '20', showError: true, errorStyle: 'stop', errorTitle: '합성 제한', error: '1~20만 입력' }])); w.undoStack = []; w.redoStack = []; });
    const before = await state(p); await p.keyboard.type('99');
    if (route === 'qat-click') await p.locator('#quickAccess [data-qat-cmd="painter"]').click(); else await altNumber(p);
    const d = p.getByRole('dialog', { name: '합성 제한', exact: true }); await d.waitFor(); const after = await state(p); assert.deepEqual(after.cells, before.cells); assert.equal(after.undo, 0); assert.equal(after.painting, false, '불허 입력 확정과 서식 복사를 함께 차단');
    await d.getByRole('button', { name: '취소', exact: true }).click(); await clickCell(p, 2, 2); assert.deepEqual((await state(p)).cells, before.cells);
  });
  for (const route of ['qat-click', 'alt-held']) await test('protected-sheet-' + route, async p => {
    await p.evaluate(() => { const w = tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true, formatCells: false } })); w.undoStack = []; w.redoStack = []; tabula.switchSheet(1); tabula.switchSheet(0); tabula.selectCell(0, 0); });
    await clickCell(p, 0, 0); const before = await state(p);
    if (route === 'qat-click') { const b = p.locator('#quickAccess [data-qat-cmd="painter"]'); if (await b.isEnabled()) await b.click(); }
    else await altNumber(p);
    assert.equal((await state(p)).painting, false, '시트 보호 우회 금지'); if (await p.getByRole('dialog').count()) await p.keyboard.press('Escape'); await clickCell(p, 2, 2); assert.deepEqual((await state(p)).cells, before.cells); assert.equal((await state(p)).undo, 0);
  });
  await test('custom-qat-slot', async p => { const before = await state(p); assert.equal(await p.locator('#quickAccess [data-qat-cmd="painter"]').getAttribute('data-qat-key'), '3'); await altNumber(p, '3'); assert.equal((await state(p)).painting, true); await clickCell(p, 2, 2); await assertApplied(p, before); }, { qatOrder: ['bold', 'italic', 'painter'] });
  await test('sticky-ribbon', async p => { const before = await state(p); await p.locator('#ribbon [data-ribbon-command="painter"]').dblclick(); assert.equal((await state(p)).painting, true); await clickCell(p, 2, 2); const first = await assertApplied(p, before); assert.equal(first.painting, true); await clickCell(p, 4, 4); const after = await state(p); assert.deepEqual(after.cells[0]['4,4'].style, before.cells[0]['0,0'].style); assert.equal(after.cells[0]['4,4'].raw, '=7*11'); assert.equal(after.undo, 2); assert.equal(after.painting, true); await p.keyboard.press('Escape'); assert.equal((await state(p)).painting, false); });
  await test('ime-delayed-events', async (p, info) => {
    const before = await state(p); await p.keyboard.press('Alt');
    info.prevented = await p.evaluate(() => {
      const ed = document.getElementById('cellEditor');
      const key = new KeyboardEvent('keydown', { key: 'Process', code: 'Digit1', keyCode: 229, isComposing: true, bubbles: true, cancelable: true }); ed.dispatchEvent(key);
      const blocked = [key.defaultPrevented];
      for (const type of ['compositionstart', 'compositionupdate', 'compositionend']) { const e = new CompositionEvent(type, { data: 'ㅁ', bubbles: true, cancelable: true }); ed.dispatchEvent(e); blocked.push(e.defaultPrevented); }
      const e = new InputEvent('beforeinput', { data: 'ㅁ', inputType: 'insertCompositionText', isComposing: true, bubbles: true, cancelable: true }); ed.dispatchEvent(e); blocked.push(e.defaultPrevented);
      ed.value = 'ㅁ'; ed.dispatchEvent(new InputEvent('input', { data: 'ㅁ', inputType: 'insertCompositionText', isComposing: false, bubbles: true }));
      return blocked;
    });
    assert.ok(info.prevented.every(Boolean), 'Alt 숫자 명령과 지연된 IME 입력 처리'); assert.equal((await state(p)).painting, true); await clickCell(p, 2, 2); const after = await assertApplied(p, before); await undoRedo(p, before, after);
    await clickCell(p, 6, 6); await p.keyboard.type('fresh'); await p.keyboard.press('Enter'); assert.equal(await p.evaluate(() => tabula.wb().getRaw(0, 6, 6)), 'fresh', '다음 정상 입력은 보존');
  });
  await test('editing-alt-enter-linebreak', async p => {
    const before = await state(p); await p.keyboard.type('first'); await p.keyboard.press('Alt+Enter'); await p.keyboard.type('second');
    assert.equal((await state(p)).editing, true); assert.equal((await state(p)).painting, false); assert.equal((await state(p)).keytips, false); assert.equal((await state(p)).editor, 'first\nsecond');
    await p.keyboard.press('Enter'); assert.equal(await p.evaluate(() => tabula.wb().getRaw(0, 0, 0)), 'first\nsecond');
    assert.deepEqual((await state(p)).cells[0]['2,2'], before.cells[0]['2,2']);
  });
  await test('editing-normal-text', async p => {
    await p.keyboard.type('plain1'); assert.equal((await state(p)).editing, true); assert.equal((await state(p)).painting, false); await p.keyboard.press('Enter'); assert.equal(await p.evaluate(() => tabula.wb().getRaw(0, 0, 0)), 'plain1');
  });
  await test('unrelated-search-input', async p => {
    const before = await state(p); await p.locator('#searchBox').focus(); await altNumber(p); const after = await state(p);
    assert.equal(after.painting, false); assert.deepEqual(after.cells, before.cells); assert.equal(after.undo, before.undo); assert.equal(after.focus, 'searchBox');
  });
  await test('active-composition-not-committed', async (p, info) => {
    const before = await state(p); await p.keyboard.type('25.5');
    info.prevented = await p.evaluate(() => {
      const ed = document.getElementById('cellEditor'), e = new KeyboardEvent('keydown', { key: 'Process', code: 'Digit1', keyCode: 229, altKey: true, isComposing: true, bubbles: true, cancelable: true });
      ed.dispatchEvent(e); return e.defaultPrevented;
    });
    const after = await state(p); assert.deepEqual(after.cells, before.cells); assert.equal(after.undo, before.undo); assert.equal(after.editing, true); assert.equal(after.painting, false); assert.equal(after.editor, '25.5');
    await p.keyboard.press('Escape');
  });
  for (const variant of ['physical-process', 'legacy-numeric']) await test('synthetic-keyboard-' + variant, async (p, info) => {
    const before = await state(p); await p.keyboard.type('25.5'); await p.keyboard.down('Alt');
    info.eventKind = 'Synthetic KeyboardEvent; OS IME 실기 검증과 구분';
    info.prevented = await p.evaluate(variant => {
      const ed = document.getElementById('cellEditor');
      const init = variant === 'physical-process' ? { key: 'Process', code: 'Digit1', keyCode: 229 } : { key: '1', code: 'Unidentified', keyCode: 49 };
      const event = new KeyboardEvent('keydown', { ...init, altKey: true, isComposing: false, bubbles: true, cancelable: true }); ed.dispatchEvent(event);
      ed.dispatchEvent(new KeyboardEvent('keyup', { ...init, altKey: true, isComposing: false, bubbles: true, cancelable: true })); return event.defaultPrevented;
    }, variant);
    await p.keyboard.up('Alt'); await frame(p); assert.equal(info.prevented, true, '물리 코드 또는 숫자 키 코드로 명령 인식');
    const expected = structuredClone(before); expected.cells[0]['0,0'].raw = '25.5'; expected.undo++;
    assert.equal((await state(p)).painting, true); assert.equal((await state(p)).editing, false); await clickCell(p, 2, 2); const after = await assertApplied(p, expected); await undoRedo(p, expected, after);
  });
} finally { await browser.close(); }
const summary = { url, engine, baseline: baseline || null, total: results.length, passed: results.filter(x => x.ok).length, failed: results.filter(x => !x.ok).length, results };
await writeFile(path.join(out, 'result.json'), JSON.stringify(summary, null, 2));
console.log(JSON.stringify({ url, engine, baseline: baseline || null, total: summary.total, passed: summary.passed, failed: summary.failed, out }));
if (!results.length || summary.failed) process.exitCode = 1;
