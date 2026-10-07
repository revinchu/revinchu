// Synthetic F4 focus and timing probes. No private workbook or document API writes.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const origin = new URL(url).origin;
const local = ['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname);
if (!local && process.env.WIXEL_ALLOWED_TEST_URL !== url) throw new Error('Public synthetic tests require an exact WIXEL_ALLOWED_TEST_URL.');
const out = path.resolve(repo, process.env.WIXEL_ROW_PERF_OUT || '.local/pivot-row-insert/performance-source');
if (!/^[dD]:/.test(out)) throw new Error('Outputs must be on D:.');
const only = process.env.WIXEL_ROW_PERF_FILTER || '';
const baseline = process.env.WIXEL_ROW_PERF_BASELINE || '';
const overrides = new Map();
if (baseline) {
  if (!local) throw new Error('Baseline routing is local only.');
  for (const name of ['app.js', 'workbook.js', 'view.js', 'header-resize.js', 'keyboard-shortcuts.js', 'ui.js']) overrides.set('/src/' + name,
    execFileSync('git', ['-c', 'safe.directory=' + path.resolve(repo, '..').replaceAll(String.fromCharCode(92), '/'), '-C', repo, 'show', baseline + ':tabula/src/' + name], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }));
}
const stressRows = Number(process.env.WIXEL_ROW_PERF_ROWS || 2500);
assert.ok(Number.isInteger(stressRows) && stressRows >= 10 && stressRows <= 20000, 'Bounded synthetic stress row count');
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await pw.chromium.launch();
await mkdir(out, { recursive: true });
const results = [], errors = [], writes = [], blocked = [], assets = new Set(); let checks = 0;
const eq = (a, b, message) => { checks++; assert.deepEqual(a, b, message); };
const ok = (value, message) => { checks++; assert.ok(value, message); };
const raf = p => p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
const state = p => p.evaluate(() => window.__f4State());
const selection = s => ({ si: s.si, active: s.active, sel: s.sel, kind: s.kind });
async function fixture(p, large = false) {
  await p.evaluate(({ large, stressRows }) => {
    const t = tabula, w = t.wb(), cells = {};
    for (let r = 0; r < 50; r++) for (let c = 0; c < 6; c++) cells[r + ',' + c] = { raw: 'row-' + r + '-col-' + c };
    const sheets = [{ name: 'F4 합성', cells, rowHeights: { 4: 18, 5: 20, 6: 22 } }];
    if (large) {
      const largeCells = {};
      for (let r = 0; r < stressRows; r++) for (let c = 0; c < 6; c++) largeCells[r + ',' + c] = { raw: c < 2 ? String(r + c) : '=A' + (r + 1) + '+B' + (r + 1) };
      sheets.push({ name: '대량 합성', cells: largeCells });
    }
    w.restore({ sheets }); t.switchSheet(0); t.gv().setZoom(100); t.gv().renderAll(); t.selectCell(5, 0);
    w.undoStack = []; w.redoStack = []; window.__f4Trace = []; window.__f4Timings = []; window.__f4LongTasks = [];
  }, { large, stressRows });
  await raf(p);
  const pt = await p.evaluate(() => { const g = tabula.gv(), r = g.clientRect({ r1: 5, c1: 0, r2: 5, c2: 0 }), v = g.viewEl.getBoundingClientRect(); return { x: Math.round(v.left + g.hw * g.z / 2), y: Math.round(r.top + r.height / 2) }; });
  await p.mouse.click(pt.x, pt.y);
  await p.locator('[data-ribbon-tab=insert]').click();
  await p.locator('#ribbon [data-ribbon-command=insertRows]').click();
  await raf(p);
  eq((await state(p)).undo, 1, 'Initial actual ribbon click inserts one row');
  eq((await state(p)).kind, 'rows', 'Actual header selection remains whole-row selection');
}
async function assertRows(p, original, expected) {
  await raf(p);
  const after = await state(p);
  eq(after.undo, expected, 'Every delivered ordinary F4 inserts exactly one row');
  eq(selection(after), selection(original), 'Same new row remains selected');
  eq(await p.evaluate(() => { const w = tabula.wb(), rows = []; w.sheets[0].cells.forEachRC((cell, r, c) => { if (!c && cell.raw === 'row-5-col-0') rows.push(r); }); return rows; }), [5 + expected], 'Original selected data is displaced by the exact insertion count');
  eq(after.editing, false, 'Repeat leaves idle grid state');
  return after;
}
async function syntheticKey(p, target, data) {
  return p.evaluate(({ target, data }) => {
    const node = document.querySelector(target); node.focus({ preventScroll: true });
    const event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'F4', code: 'F4', keyCode: 115, which: 115, ...data });
    if (data.fn) { const native = event.getModifierState.bind(event); Object.defineProperty(event, 'getModifierState', { value: name => name === 'Fn' || native(name) }); }
    node.dispatchEvent(event); return { defaultPrevented: event.defaultPrevented, key: event.key, code: event.code, keyCode: event.keyCode, isComposing: event.isComposing, fn: event.getModifierState('Fn') };
  }, { target, data });
}
async function test(name, fn) {
  if (only && !only.split('|').some(s => name.includes(s))) return;
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 }, serviceWorkers: 'block' });
  const p = await context.newPage(), info = {}, pageErrors = [], caseWrites = []; p.setDefaultTimeout(15000);
  p.on('pageerror', e => { pageErrors.push(e.message); errors.push({ name, error: e.message }); }); p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {} } }); });
  await context.route('**/*', route => { const request = route.request(), u = new URL(request.url()); if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push({ name, method: request.method(), path: u.pathname }); caseWrites.push(u.pathname); return route.abort(); } if (u.origin !== origin || /^\/api(?:\/|$)/.test(u.pathname)) { blocked.push(u.pathname); return route.abort(); } if (overrides.has(u.pathname)) return route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: overrides.get(u.pathname) }); return route.continue(); });
  const start = checks;
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    for (const src of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => n.getAttribute('src')))) assets.add(src);
    await p.evaluate(() => {
      const t = tabula, w = t.wb(), g = t.gv();
      window.__f4State = () => ({ si: t.si, active: { ...t.active }, sel: { ...t.sel }, kind: g.host.state().selKind, editing: g.host.state().editing, undo: w.undoStack.length, redo: w.redoStack.length, focus: { id: document.activeElement?.id, cls: document.activeElement?.className, tag: document.activeElement?.tagName }, editor: { value: document.getElementById('cellEditor').value, idle: document.getElementById('cellEditor').classList.contains('idle') }, keytip: document.body.dataset.keytipSequence ?? null });
      window.__f4Trace = []; window.__f4Timings = []; window.__f4LongTasks = [];
      for (const [object, names, owner] of [[w, ['shiftAxis', 'transact', 'serialize'], 'wb'], [g, ['renderAll', 'layout'], 'grid']]) for (const name of names) {
        const original = object[name]; if (typeof original !== 'function') continue;
        object[name] = function (...args) { const start = performance.now(), before = w.undoStack.length; try { return original.apply(this, args); } finally { __f4Timings.push({ owner, name, ms: performance.now() - start, before, after: w.undoStack.length }); } };
      }
      if (typeof PerformanceObserver === 'function') try { new PerformanceObserver(list => __f4LongTasks.push(...list.getEntries().map(e => ({ start: e.startTime, duration: e.duration })))).observe({ type: 'longtask', buffered: true }); } catch {}
      const pending = new WeakMap();
      const finish = (e, phase) => {
        const event = pending.get(e); if (!event || event.finished) return;
        event.finished = true; event.phase = phase; event.after = __f4State(); event.defaultPrevented = e.defaultPrevented; event.handlerMs = performance.now() - event.at;
      };
      // Native dispatch can flush microtasks between listeners, so a capture
      // microtask is not a reliable post-handler measurement.
      for (const type of ['keydown', 'keyup', 'focusin', 'mousedown', 'click', 'beforeinput', 'input', 'compositionstart', 'compositionend']) {
        window.addEventListener(type, e => {
          const event = { type, key: e.key, code: e.code, keyCode: e.keyCode, repeat: e.repeat, composing: e.isComposing, trusted: e.isTrusted, target: e.target?.id, cls: e.target?.className, fn: e.getModifierState?.('Fn'), ctrl: e.ctrlKey, alt: e.altKey, shift: e.shiftKey, at: performance.now(), eventStamp: e.timeStamp, before: __f4State() };
          pending.set(e, event); __f4Trace.push(event); setTimeout(() => finish(e, 'async-fallback'), 0);
        }, true);
        document.addEventListener(type, e => { if (e.defaultPrevented || e.cancelBubble) finish(e, 'document-capture-after-app'); }, true);
        for (const node of document.querySelectorAll('#cellEditor,#gridView,#searchBox,#formulaInput')) node.addEventListener(type, e => finish(e, 'target-or-grid-bubble-after-app'));
        window.addEventListener(type, e => finish(e, 'window-bubble-after-app'));
      }
    });
    await fn(p, info);
    await raf(p);
    eq(pageErrors, [], 'No application exception'); eq(caseWrites, [], 'No API or remote writes');
    const diagnostic = await p.evaluate(() => ({ trace: __f4Trace, timings: __f4Timings, longTasks: __f4LongTasks, final: __f4State() }));
    results.push({ name, ok: true, checks: checks - start, info, ...diagnostic }); console.log('OK ' + name);
  } catch (error) {
    const diagnostic = await p.evaluate(() => ({ trace: __f4Trace, timings: __f4Timings, longTasks: __f4LongTasks, final: __f4State() })).catch(() => ({}));
    results.push({ name, ok: false, checks: checks - start, error: error.message, info, ...diagnostic }); console.error('NG ' + name + ': ' + error.message.split('\n')[0]); await p.screenshot({ path: path.join(out, 'failure-' + results.length + '.png') }).catch(() => {});
  } finally { await context.close(); }
}
try {
  for (const target of ['#cellEditor', '#gridView', '.sheet-tab.active', '[data-ribbon-tab=insert]', '#ribbon [data-ribbon-command=insertRows]']) await test('ordinary-F4-focus-' + target, async (p, info) => {
    await fixture(p); const before = await state(p); info.before = before; await p.locator(target).focus(); info.focused = await state(p);
    eq(info.focused.editing, false, 'This focus target remains outside cell editing');
    for (let i = 0; i < 8; i++) await p.keyboard.press('F4');
    info.after = await assertRows(p, before, 9);
    eq(info.after.editor.value, '', 'Idle editor receives no literal F4 input');
  });
  await test('Fn-modifier-with-F4-key-and-code', async (p, info) => {
    await fixture(p); const before = await state(p); info.event = await syntheticKey(p, '#cellEditor', { fn: true });
    eq(info.event.fn, true, 'Probe reports the simulated Fn modifier'); ok(info.event.defaultPrevented, 'F4 remains handled when key and code are F4'); await assertRows(p, before, 2);
  });
  await test('idle-IME-flags-and-physical-code-diagnostic', async (p, info) => {
    await fixture(p); info.events = [];
    for (const target of ['#cellEditor', '#gridView', '#ribbon [data-ribbon-command=insertRows]']) for (const data of [{ key: 'F4', code: 'F4', keyCode: 115 }, { key: 'F4', code: 'F4', keyCode: 229 }, { key: 'F4', code: 'F4', isComposing: true }, { key: 'Unidentified', code: 'F4' }, { key: 'Process', code: 'F4', keyCode: 229 }]) {
      const before = await state(p), event = await syntheticKey(p, target, data); await raf(p); const after = await state(p); info.events.push({ target, supplied: data, event, before, after, inserts: after.undo - before.undo });
      eq(after.editing, false, 'Synthetic keydown does not enter text editing');
      if (data.keyCode === 115) eq(after.undo, before.undo + 1, 'Ordinary F4 control case repeats');
    }
    info.assertionScope = 'Nonstandard IME flags and key/code mismatch are diagnostics, not proof of physical keyboard behavior.';
  });
  await test('physical-code-F4-Process229-idle-focus-repeat-and-keyup', async (p, info) => {
    await fixture(p); const before = await state(p); info.events = []; let count = 1;
    for (const target of ['#cellEditor', '#gridView', '#ribbon [data-ribbon-command=insertRows]']) for (const data of [{ key: 'Process', code: 'F4', keyCode: 229 }, { key: 'Unidentified', code: 'F4', keyCode: 229 }, { key: 'F4', code: 'F4', keyCode: 229 }]) {
      const event = await syntheticKey(p, target, data);
      await p.evaluate(({ target, data }) => document.querySelector(target).dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, cancelable: true, ...data })), { target, data });
      info.events.push({ target, supplied: data, event, after: await state(p) });
      ok(event.defaultPrevented, 'Idle physical F4 is handled despite a Process/229 key label'); await assertRows(p, before, ++count); eq((await state(p)).editor.value, '', 'F4 keyup leaves the idle editor empty');
    }
  });
  await test('composing-F4-preserves-IME-body-and-does-not-insert', async (p, info) => {
    await fixture(p); const before = await state(p);
    await p.evaluate(() => { const ed = document.getElementById('cellEditor'); ed.focus(); ed.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Process', code: 'KeyG', keyCode: 229 })); ed.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })); ed.value = '한글 본문'; ed.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '한글 본문', isComposing: true })); });
    info.beforeF4 = await state(p); eq(info.beforeF4.editing, true, 'Composition enters normal cell text editing'); eq(info.beforeF4.editor.value, '한글 본문');
    info.event = await syntheticKey(p, '#cellEditor', { key: 'Process', code: 'F4', keyCode: 229, isComposing: true });
    await p.evaluate(() => document.getElementById('cellEditor').dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Process', code: 'F4', keyCode: 229, isComposing: true })));
    info.afterF4 = await state(p); eq(info.afterF4.undo, before.undo, 'Composing F4 never inserts a row'); eq(info.afterF4.editor.value, '한글 본문', 'Composition text is preserved'); eq(info.afterF4.editing, true); await p.keyboard.press('Escape'); await p.keyboard.press('F4'); await assertRows(p, before, 2);
  });
  await test('ordinary-input-physical-F4-isolated-and-new-IME-input-preserved', async (p, info) => {
    await fixture(p); const before = await state(p); await p.locator('#searchBox').fill('검색 본문');
    info.event = await syntheticKey(p, '#searchBox', { key: 'Process', code: 'F4', keyCode: 229 }); await p.evaluate(() => document.getElementById('searchBox').dispatchEvent(new KeyboardEvent('keyup', { bubbles: true, key: 'Process', code: 'F4', keyCode: 229 })));
    eq((await state(p)).undo, before.undo, 'Ordinary input stays outside global F4'); eq(await p.locator('#searchBox').inputValue(), '검색 본문', 'F4 preserves unrelated input text'); await p.keyboard.press('Escape');
    await p.locator('#cellEditor').focus(); await p.keyboard.press('F4'); await assertRows(p, before, 2);
    await p.evaluate(() => { const ed = document.getElementById('cellEditor'); ed.dispatchEvent(new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'Process', code: 'KeyG', keyCode: 229 })); ed.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' })); ed.value = '다음 한글'; ed.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: '다음 한글', isComposing: true })); ed.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '다음 한글' })); });
    info.afterNewInput = await state(p); eq(info.afterNewInput.editor.value, '다음 한글', 'A fresh IME gesture after F4 is kept'); eq(info.afterNewInput.editing, true); eq(info.afterNewInput.undo, 2, 'Composition does not insert an extra row'); await p.keyboard.press('Escape');
  });
  await test('held-F4-autoRepeat-keydown-events', async (p, info) => {
    await fixture(p); const before = await state(p); info.events = [];
    for (let i = 0; i < 8; i++) info.events.push(await syntheticKey(p, '#cellEditor', { repeat: i > 0 }));
    await assertRows(p, before, 9); ok(info.events.every(e => e.defaultPrevented), 'Every held F4 keydown is handled');
  });
  await test('same-task-ribbon-click-then-F4-before-queued-render', async (p, info) => {
    await fixture(p); const before = await state(p);
    info.operations = await p.evaluate(() => { const result = []; for (let i = 0; i < 10; i++) { const before = __f4State(), start = performance.now(); document.querySelector('#ribbon [data-ribbon-command=insertRows]').click(); const editor = document.getElementById('cellEditor'), event = new KeyboardEvent('keydown', { bubbles: true, cancelable: true, key: 'F4', code: 'F4', keyCode: 115 }); editor.dispatchEvent(event); result.push({ ms: performance.now() - start, before, after: __f4State(), prevented: event.defaultPrevented }); } return result; });
    await assertRows(p, before, 21); ok(info.operations.every(e => e.prevented), 'F4 is handled before each queued UI render');
  });
  await test('trusted-CDP-rapid-F4-delivery-during-long-task', async (p, info) => {
    await fixture(p); const before = await state(p); await p.locator('#cellEditor').focus(); const cdp = await p.context().newCDPSession(p);
    await p.evaluate(() => { setTimeout(() => { const until = performance.now() + 180; while (performance.now() < until) {} }, 0); });
    const requests = [], start = Date.now();
    for (let i = 0; i < 12; i++) { requests.push(cdp.send('Input.dispatchKeyEvent', { type: 'keyDown', key: 'F4', code: 'F4', windowsVirtualKeyCode: 115, nativeVirtualKeyCode: 115, autoRepeat: i > 0 })); requests.push(cdp.send('Input.dispatchKeyEvent', { type: 'keyUp', key: 'F4', code: 'F4', windowsVirtualKeyCode: 115, nativeVirtualKeyCode: 115 })); }
    await Promise.all(requests); info.requested = 12; info.wallMs = Date.now() - start; await assertRows(p, before, 13);
    const delivered = await p.evaluate(() => __f4Trace.filter(e => e.type === 'keydown' && e.key === 'F4'));
    eq(delivered.length, 12, 'Browser delivered every requested protocol F4 keydown'); ok(delivered.every(e => e.trusted), 'Browser protocol events use trusted input dispatch'); info.delivered = delivered.length;
    info.limit = 'A deliberately busy browser main thread and protocol input do not establish whether a physical device drops keys.';
  });
  await test('Alt-H-I-R-shortcut-guard-then-repeated-F4', async (p, info) => {
    await fixture(p); const before = await state(p); await p.locator('#cellEditor').focus(); await p.keyboard.press('Alt'); await p.keyboard.press('h'); await p.keyboard.press('i'); await p.keyboard.press('r'); await raf(p);
    info.afterAlt = await state(p); eq(info.afterAlt.undo, 2, 'Alt H I R really inserted a sheet row'); eq(info.afterAlt.editing, false);
    info.repeat = await syntheticKey(p, '#cellEditor', { repeat: true }); await assertRows(p, before, 3); await p.keyboard.press('F4'); await assertRows(p, before, 4);
  });
  await test('Alt-keytips-open-F4-closes-keytips-and-repeats', async (p, info) => {
    await fixture(p); const before = await state(p); await p.locator('#cellEditor').focus(); await p.keyboard.press('Alt'); info.keytip = await state(p); eq(info.keytip.keytip, '', 'Alt opens keytips'); await p.keyboard.press('F4'); info.after = await assertRows(p, before, 2); eq(info.after.keytip, null, 'F4 closes keytips');
  });
  await test('formula-bar-edit-F4-is-reference-toggle-and-Escape-recovers-repeat', async (p, info) => {
    await fixture(p); const before = await state(p); await p.locator('#formulaInput').focus(); info.focused = await state(p); eq(info.focused.editing, true, 'Formula focus enters edit mode through its normal focus handler'); await p.locator('#formulaInput').fill('=A1'); await p.keyboard.press('F4'); info.formula = await p.locator('#formulaInput').inputValue(); eq(info.formula, '=$A$1'); eq((await state(p)).undo, 1, 'Editing F4 does not insert a row'); await p.keyboard.press('Escape'); info.escaped = await state(p); eq(info.escaped.editing, false); await p.keyboard.press('F4'); await assertRows(p, before, 2);
  });
  await test('blank-cell-editor-edit-F4-no-insert-then-Escape-repeat', async (p, info) => {
    await fixture(p); const before = await state(p); await p.locator('#cellEditor').focus(); await p.keyboard.press('F2'); info.editing = await state(p); eq(info.editing.editing, true); await p.keyboard.press('F4'); eq((await state(p)).undo, 1, 'Blank cell edit F4 stays in editing semantics'); await p.keyboard.press('Escape'); await p.keyboard.press('F4'); await assertRows(p, before, 2);
  });
  await test('unrelated-text-input-F4-no-repeat-then-grid-repeat', async (p, info) => {
    await fixture(p); const before = await state(p); const input = p.locator('#searchBox'); await input.focus(); info.focused = await state(p); eq(info.focused.editing, false); await p.keyboard.press('F4'); eq((await state(p)).undo, 1, 'Unrelated input is isolated from global repeat'); await p.keyboard.press('Escape'); await p.locator('#gridView').focus(); await p.keyboard.press('F4'); await assertRows(p, before, 2);
  });
  await test('synthetic-formula-sheet-row-insert-timing-and-rapid-repeat', async (p, info) => {
    await fixture(p, true); const before = await state(p); info.syntheticCells = stressRows * 6; info.syntheticFormulas = stressRows * 4; info.operations = [];
    for (let i = 0; i < 6; i++) { const start = Date.now(); await p.keyboard.press('F4'); await raf(p); info.operations.push({ wallMs: Date.now() - start, state: await state(p) }); }
    await assertRows(p, before, 7); const values = info.operations.map(e => e.wallMs).sort((a, b) => a - b); info.medianWallMs = values[Math.floor(values.length / 2)]; info.maxWallMs = values.at(-1); info.performanceScope = 'Timing is reported without a machine-dependent hard threshold; unrelated synthetic formula sheet remains present.';
  });
} finally {
  await browser.close();
  const summary = { url, baseline: baseline || null, syntheticOnly: true, physicalKeyboardVerified: false, stressRows, assets: [...assets], cases: results.length, passed: results.filter(r => r.ok).length, checks, pageErrors: errors, remoteWrites: writes, blockedRequests: blocked, results };
  await writeFile(path.join(out, 'result.json'), JSON.stringify(summary, null, 2));
  console.log(JSON.stringify({ url, baseline: summary.baseline, cases: summary.cases, passed: summary.passed, checks, out, failures: results.filter(r => !r.ok).map(({ name, error }) => ({ name, error })) }));
  if (summary.passed !== summary.cases || errors.length || writes.length) process.exitCode = 1;
}
