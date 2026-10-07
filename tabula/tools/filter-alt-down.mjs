// Synthetic filter Alt+Down regression. Private files, clipboard and remote writes are not used.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
const repo = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/', origin = new URL(url).origin;
const local = ['127.0.0.1', 'localhost', '[::1]'].includes(new URL(url).hostname);
if (!local && process.env.WIXEL_ALLOWED_TEST_URL !== url) throw new Error('Public synthetic tests need an exact WIXEL_ALLOWED_TEST_URL.');
const out = path.resolve(repo, process.env.WIXEL_FILTER_ALT_OUT || '.local/filter-alt-down/source');
if (!/^[dD]:/.test(out)) throw new Error('Outputs must stay on D:.');
const only = process.env.WIXEL_FILTER_ALT_FILTER || '', baseline = process.env.WIXEL_FILTER_ALT_BASELINE || '', overrides = new Map();
if (baseline) {
  if (!local) throw new Error('Baseline memory routing is local only.');
  for (const name of ['app.js', 'ui.js', 'view.js', 'keyboard-shortcuts.js', 'app-filter-checklist.js']) overrides.set('/src/' + name,
    execFileSync('git', ['-c', 'safe.directory=' + path.resolve(repo, '..').replaceAll(String.fromCharCode(92), '/'), '-C', repo, 'show', baseline + ':tabula/src/' + name], { encoding: 'utf8', maxBuffer: 32 * 1024 * 1024 }));
}
const pw = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium';
if (!['chromium', 'firefox', 'webkit'].includes(engine)) throw new Error('Unsupported browser');
const browser = await pw[engine].launch(); await mkdir(out, { recursive: true });
const results = [], errors = [], writes = [], blocked = [], assets = new Set(); let checks = 0;
const eq = (a, b, message) => { checks++; assert.deepEqual(a, b, message); };
const ok = (value, message) => { checks++; assert.ok(value, message); };
const raf = p => p.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
const state = p => p.evaluate(() => window.__filterState());
const book = p => p.evaluate(() => window.__filterBook());
const run = async (p, command) => { await p.evaluate(command => tabula.run(command), command); await raf(p); };
async function fixture(p, type = 'ordinary', { many = false, distant = false, hiddenHeaders = false, sigmaOnly = false, classic = false } = {}) {
  const info = await p.evaluate(({ type, many, distant, hiddenHeaders, sigmaOnly, classic }) => {
    const t = tabula, w = t.wb(), cells = {}, n = many ? 600 : 12, r0 = distant ? 40 : 4, c0 = distant ? 9 : 2;
    const header = ['Group', 'Channel', 'Week', 'Value', 'Cost'];
    const data = Array.from({ length: n }, (_, i) => [many ? 'Item' + String(i + 1).padStart(4, '0') : ['Alpha', 'Beta', 'Gamma'][i % 3], i % 2 ? 'Offline' : 'Online', Math.floor(i / 3) % 2 ? 'Week B' : 'Week A', i + 1, (i + 1) * 10]);
    if (type === 'pivot') {
      [header, ...data].forEach((row, r) => row.forEach((value, c) => cells[r + ',' + c] = { raw: String(value) }));
      const def = { name: 'Keyboard Pivot', source: 'Source', range: { r1: 0, c1: 0, r2: n, c2: 4 }, rows: sigmaOnly ? [] : ['Group'], cols: sigmaOnly ? [] : ['Channel'], pages: ['Week'],
        values: [{ field: 'Value', agg: 'sum' }, { field: 'Cost', agg: 'sum' }], layout: 'tabular', subtotals: 'none', grandRows: true, grandCols: true, autofit: false, preserveFormat: true, classic,
        top: r0, left: c0, ...(hiddenHeaders ? { showHeaders: false } : {}) };
      w.restore({ sheets: [{ name: 'Source', cells }, { name: 'Report', cells: {}, pivot: def, colWidths: { [c0]: 140, [c0 + 1]: 140 } }] });
      t.switchSheet(1); t.selectCell(r0, c0); t.run('pivotRefresh');
    } else {
      [header, ...data].forEach((row, r) => row.forEach((value, c) => cells[(r0 + r) + ',' + (c0 + c)] = { raw: String(value) }));
      cells['0,0'] = { raw: 'unrelated' };
      const sh = { name: 'Filter sheet', cells };
      const range = { r1: r0, c1: c0, r2: r0 + n, c2: c0 + 4 };
      if (type === 'table') sh.tables = [{ id: 'keyboard-table', name: 'KeyboardTable', ...range, header: true, style: 'TableStyleMedium2', filter: { criteria: {}, hidden: {} } }];
      else sh.filter = { ...range, criteria: {}, hidden: {} };
      w.restore({ sheets: [sh] }); t.switchSheet(0); t.selectCell(r0, c0);
    }
    t.gv().setZoom(100); t.gv().layout(); w.undoStack = []; w.redoStack = [];
    window.__filterFixture = { type, n, r0, c0, data, header, si: t.si };
    window.__filterTrace = [];
    return { type, n, r0, c0, si: t.si, buttons: structuredClone(w.sheets[t.si].pivot?.buttons ?? []) };
  }, { type, many, distant, hiddenHeaders, sigmaOnly, classic });
  await raf(p); await p.locator('#cellEditor').focus(); return info;
}
async function selectButton(p, kind, field = null, { labelCell = false, focus = '#cellEditor' } = {}) {
  const point = await p.evaluate(({ kind, field, labelCell }) => {
    const t = tabula, f = __filterFixture, sh = t.wb().sheets[t.si];
    let point = { r: f.r0, c: f.c0 };
    if (f.type === 'pivot') {
      const b = sh.pivot.buttons.find(b => b.kind === kind && b.kind !== 'toggle' && (field === null || b.field === field) && !b.sigma);
      if (!b) throw Error('Missing fixture pivot button ' + kind + ':' + field);
      point = { r: b.r, c: b.c - (labelCell ? 1 : 0) };
    }
    t.selectCell(point.r, point.c); t.gv().ensureVisible(point.r, point.c); return point;
  }, { kind, field, labelCell }); await raf(p); await p.locator(focus).focus(); return point;
}
async function altDown(p, { releaseFirst = false, held = false } = {}) {
  await p.keyboard.down('Alt');
  if (releaseFirst) await p.keyboard.up('Alt');
  await p.keyboard.press('ArrowDown');
  if (!releaseFirst && !held) await p.keyboard.up('Alt');
  await raf(p);
}
async function open(p, type, options = {}) {
  const before = await state(p); await altDown(p, options);
  const menu = p.locator('.menu:has(.filter-menu)').first();
  await menu.waitFor({ state: 'visible', timeout: 3000 });
  if (baseline) await p.waitForFunction(() => document.activeElement?.matches('input[type=search]') && !!document.activeElement.closest('.menu'), null, { timeout: 3000 });
  else {
    ok(await p.evaluate(() => document.activeElement?.matches('.menu-item:not(:disabled),.pf-act:not(:disabled),.pf-one[tabindex="0"]')), 'Keyboard starts at a menu command or report-filter option');
    if(options.search !== false) { await p.keyboard.press('e');
    await p.waitForFunction(() => document.activeElement?.matches('input[type=search]') && !!document.activeElement.closest('.menu'), null, { timeout: 3000 });
    eq(await menu.getByRole('searchbox').first().inputValue(), '', 'E focuses search without inserting a letter'); }
  }
  const after = await state(p);
  eq(after.menus, 1, 'One filter popup opens');
  eq(after.active, before.active, 'Alt+Down keeps the selected button cell');
  eq(after.sel, before.sel, 'Opening does not change the cell selection');
  eq(after.undo, before.undo, 'Opening makes no Undo entry');
  eq(after.editing, false, 'Opening stays outside cell editing');
  eq(after.keytip, null, 'Alt+Down ends ribbon keytips');
  ok(await menu.locator(type === 'pivot' ? '.pivot-filter-menu' : '.filter-menu').isVisible(), 'Correct filter popup is visible');
  return menu;
}
async function esc(p, before) {
  await p.keyboard.press('Escape'); await raf(p);
  eq((await state(p)).menus, 0, 'Escape closes the popup');
  eq(await book(p), before, 'Escape discards a filter draft and preserves Undo/Redo');
  ok(['cellEditor', 'gridView'].includes((await state(p)).focus.id), 'Popup returns focus to the idle grid');
}
async function test(name, fn) {
  if (only && !only.split('|').some(filter => name.includes(filter))) return;
  const context = await browser.newContext({ viewport: { width: 1440, height: 1020 }, serviceWorkers: 'block' }), p = await context.newPage(), info = {}, pageErrors = [], caseWrites = [];
  p.setDefaultTimeout(12000); p.on('pageerror', e => { errors.push({ name, error: e.message }); pageErrors.push(e.message); }); p.on('dialog', d => d.dismiss());
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel:version', '3.0.0'); localStorage.setItem('wixel.mobile-work.v1', 'off'); Object.defineProperty(navigator, 'clipboard', { configurable: true, value: { readText: async () => '', writeText: async () => {} } }); });
  await context.route('**/*', route => { const request = route.request(), target = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push({ name, path: target.pathname, method: request.method() }); caseWrites.push(target.pathname); return route.abort(); }
    if (target.origin !== origin || /^\/api(?:\/|$)/.test(target.pathname)) { blocked.push({ name, path: target.pathname }); return route.abort(); }
    if (overrides.has(target.pathname)) return route.fulfill({ status: 200, contentType: 'text/javascript; charset=utf-8', body: overrides.get(target.pathname) }); return route.continue(); });
  const start = checks;
  try {
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.gv());
    for (const src of await p.locator('script[src]').evaluateAll(nodes => nodes.map(n => n.getAttribute('src')))) assets.add(src);
    await p.evaluate(() => {
      window.__filterTrace = [];
      window.__filterState = () => { const t = tabula, w = t.wb(), g = t.gv(), node = document.activeElement; return { si: t.si, active: { ...t.active }, sel: { ...t.sel }, editing: !!g.host.state().editing, undo: w.undoStack.length, redo: w.redoStack.length,
        keytip: document.body.dataset.keytipSequence ?? null, menus: document.querySelectorAll('.menu').length, dialogs: document.querySelectorAll('.dialog').length,
        focus: { id: node?.id ?? '', tag: node?.tagName ?? '', type: node?.type ?? '', cls: String(node?.className ?? ''), label: node?.getAttribute('aria-label'), text: node?.textContent?.trim().slice(0, 80), value: node?.value, checked: node?.checked },
        scroll: { x: g.sx, y: g.sy }, zoom: g.z, editor: document.getElementById('cellEditor').value }; };
      window.__filterBook = () => ({ sheets: tabula.wb().sheets.map((sh, si) => { const { _sid, fileValues, ...data } = tabula.wb().serializeSheet(si); return data; }), undo: tabula.wb().undoStack.length, redo: tabula.wb().redoStack.length });
      for (const type of ['keydown', 'keyup', 'focusin']) {
        const records = new WeakMap();
        window.addEventListener(type, e => { const record = { type, key: e.key, code: e.code, alt: e.altKey, ctrl: e.ctrlKey, composing: e.isComposing, trusted: e.isTrusted, target: e.target?.id, before: __filterState() }; __filterTrace.push(record); records.set(e, record); }, true);
        window.addEventListener(type, e => { const record = records.get(e); if (record) { record.after = __filterState(); record.prevented = e.defaultPrevented; } });
      }
    });
    await fn(p, info); await raf(p); eq(pageErrors, [], 'No browser exception'); eq(caseWrites, [], 'No API or remote writes');
    results.push({ name, ok: true, checks: checks - start, info, state: await state(p), trace: await p.evaluate(() => __filterTrace) }); console.log('OK ' + name);
  } catch (error) {
    results.push({ name, ok: false, checks: checks - start, error: error.message, info, state: await state(p).catch(() => null), trace: await p.evaluate(() => window.__filterTrace).catch(() => []), pageErrors, remoteWrites: caseWrites });
    console.error('NG ' + name + ': ' + error.message.split('\n')[0]); await p.screenshot({ path: path.join(out, 'failure-' + results.length + '.png') }).catch(() => {});
  } finally { await context.close(); }
}
try {
  for (const type of ['ordinary', 'table']) for (const focus of ['#cellEditor', '#gridView']) await test(type + '-header-AltDown-focus-' + focus, async (p, info) => {
    info.fixture = await fixture(p, type); await selectButton(p, null, null, { focus }); const before = await book(p); await open(p, type); await esc(p, before);
  });
  for (const type of ['ordinary', 'table', 'pivot']) await test(type + '-keyboard-arrow-Tab-Space-draft-Escape', async (p, info) => {
    info.fixture = await fixture(p, type); await selectButton(p, type === 'pivot' ? 'rows' : null, 'Group'); const before = await book(p); const menu = await open(p, type);
    await p.keyboard.press('ArrowDown'); eq((await state(p)).focus.label, '표시된 항목 모두 선택', 'Search ArrowDown focuses select-all checkbox');
    await p.keyboard.press('ArrowDown'); eq((await state(p)).focus.label, 'Alpha', 'Checklist ArrowDown focuses first data item');
    await p.keyboard.press('ArrowDown'); eq((await state(p)).focus.label, 'Beta', 'Checklist moves to the second data item');
    await p.keyboard.press('Space'); eq(await menu.getByRole('checkbox', { name: 'Beta', exact: true }).isChecked(), false, 'Space changes the focused draft checkbox');
    await p.keyboard.press('Tab'); eq((await state(p)).focus.label, 'Gamma', 'Tab reaches the next checkbox'); await p.keyboard.press('Shift+Tab'); eq((await state(p)).focus.label, 'Beta', 'Shift+Tab returns to the prior checkbox');
    await p.keyboard.press('Space'); eq(await menu.getByRole('checkbox', { name: 'Beta', exact: true }).isChecked(), true, 'Space restores the checkbox'); await esc(p, before);
  });
  for (const type of ['ordinary', 'table', 'pivot']) await test(type + '-search-Enter-applies-one-filter-and-UndoRedo', async (p, info) => {
    info.fixture = await fixture(p, type); await selectButton(p, type === 'pivot' ? 'rows' : null, 'Group'); const before = await book(p); const menu = await open(p, type);
    await menu.getByRole('searchbox').first().fill('Beta'); await p.keyboard.press('Enter'); await raf(p);
    const after = await book(p); eq(after.undo, before.undo + 1, 'Filter confirmation is exactly one Undo'); eq((await state(p)).menus, 0, 'Enter closes the popup');
    const selection = await p.evaluate(type => { const sh = tabula.wb().sheets[tabula.si]; return type === 'pivot' ? sh.pivot.filters.Group : (type === 'table' ? sh.tables[0].filter : sh.filter).criteria[__filterFixture.c0]; }, type);
    eq(selection, ['Beta'], 'Search Enter applies the exact selected item');
    if (type !== 'pivot') eq(await p.evaluate(type => { const sh = tabula.wb().sheets[0], f = type === 'table' ? sh.tables[0].filter : sh.filter; return Object.keys(f.hidden).map(Number).sort((a, b) => a - b); }, type), [5, 7, 8, 10, 11, 13, 14, 16], 'Only nonmatching fixture rows are hidden');
    else eq(await p.evaluate(() => { const t = tabula, d = t.wb().sheets[t.si].pivot; return t.wb().getValue(t.si, d.area.r2, d.area.c2); }), 260, 'Filtered pivot grand total equals independent source total');
    await run(p, 'undo'); eq((await book(p)).sheets, before.sheets, 'Undo restores exact original filter and pivot cells'); await run(p, 'redo'); eq((await book(p)).sheets, after.sheets, 'Redo restores confirmed filter and output');
  });
  for (const [kind, field] of [['rows', 'Group'], ['cols', 'Channel'], ['page', 'Week']]) await test('pivot-' + kind + '-button-cell-AltDown-Escape', async (p, info) => {
    info.fixture = await fixture(p, 'pivot'); info.point = await selectButton(p, kind, field); const before = await book(p); await open(p, 'pivot'); await esc(p, before);
  });
  await test('pivot-report-filter-value-cell-single-choice-keyboard', async (p, info) => {
    info.fixture = await fixture(p, 'pivot'); info.point = await selectButton(p, 'page', 'Week'); const before = await book(p); const menu = await open(p, 'pivot');
    await menu.getByRole('searchbox').fill('Week B'); await p.keyboard.press('ArrowDown'); eq((await state(p)).focus.tag, 'DIV', 'Search ArrowDown focuses a report-filter option'); await p.keyboard.press('Enter'); await raf(p);
    eq(await p.evaluate(() => tabula.wb().sheets[tabula.si].pivot.filters.Week), ['Week B'], 'Report-filter keyboard applies one item');
    eq((await state(p)).undo, before.undo + 1, 'Report-filter confirmation is one Undo'); await run(p, 'undo'); eq((await book(p)).sheets, before.sheets, 'Undo restores report-filter output and definition');
  });
  await test('ordinary-virtual-checklist-End-Space-Tab-Enter', async (p, info) => {
    info.fixture = await fixture(p, 'ordinary', { many: true }); await selectButton(p); const before = await book(p); const menu = await open(p, 'ordinary');
    ok(await menu.locator('.filter-window input').count() < 80, 'Large item list creates only viewport checkbox nodes');
    await p.keyboard.press('ArrowDown'); await p.keyboard.press('End'); eq((await state(p)).focus.label, 'Item0600', 'End focuses the last virtual item');
    await p.keyboard.press('Space'); eq((await state(p)).focus.checked, false, 'Space unchecks the last virtual item'); await p.keyboard.press('Tab'); eq((await state(p)).focus.text, '확인', 'Tab reaches the confirm button after the last item');
    await p.keyboard.press('Enter'); await raf(p); eq((await state(p)).undo, before.undo + 1, 'Virtual selection confirms as one transaction');
    eq(await p.evaluate(() => Object.keys(tabula.wb().sheets[0].filter.hidden).map(Number)), [604], 'Only last source row is hidden'); await run(p, 'undo'); eq((await book(p)).sheets, before.sheets, 'Undo restores the original virtual-list selection');
  });
  for (const type of ['ordinary', 'table', 'pivot']) await test(type + '-zoom55-scroll-freeze-visible-button-anchor', async (p, info) => {
    info.fixture = await fixture(p, type, { distant: true }); await p.evaluate(() => { const t = tabula, w = t.wb(); w.sheets[t.si].freeze = { rows: 2, cols: 1 }; t.gv().setZoom(55); t.gv().layout(); });
    await selectButton(p, type === 'pivot' ? 'rows' : null, 'Group'); const before = await book(p); info.before = await state(p); const menu = await open(p, type);
    info.menuRect = await menu.boundingBox(); ok(info.menuRect.x >= -1 && info.menuRect.y >= -1 && info.menuRect.x + info.menuRect.width <= 1441 && info.menuRect.y + info.menuRect.height <= 1021, 'Popup stays inside the viewport');
    eq((await state(p)).scroll, info.before.scroll, 'Popup opening preserves grid scroll'); await esc(p, before); await p.screenshot({ path: path.join(out, type + '-zoom55-freeze.png') });
  });
  await test('released-Alt-then-ArrowDown-is-navigation-not-filter', async (p, info) => {
    info.fixture = await fixture(p); await selectButton(p); const before = await state(p); await altDown(p, { releaseFirst: true });
    eq(await p.locator('.filter-menu').count(), 0, 'Released Alt does not become Alt+Down'); eq((await state(p)).active, { r: before.active.r + 1, c: before.active.c }, 'ArrowDown after released Alt navigates normally');
    eq((await state(p)).undo, 0, 'Navigation does not edit the workbook');
  });
  await test('ordinary-input-and-formula-focus-do-not-open-filter', async (p, info) => {
    info.fixture = await fixture(p); await selectButton(p);
    for (const target of ['#searchBox', '#formulaInput']) { await p.locator(target).focus(); const before = await book(p); await altDown(p); eq(await p.locator('.filter-menu').count(), 0, 'Unrelated input stays outside grid filter shortcuts'); eq(await book(p), before, 'Input focus shortcut does not mutate the workbook'); await p.keyboard.press('Escape'); }
  });
  await test('open-menu-search-AltDown-isolated-from-grid-filter', async (p, info) => {
    info.fixture = await fixture(p); await selectButton(p); const before = await book(p); const menu = await open(p, 'ordinary'); await menu.getByRole('searchbox').fill('Beta');
    await altDown(p); eq((await state(p)).menus, 1, 'Alt+Down inside menu opens no second grid filter'); eq(await menu.getByRole('searchbox').inputValue(), 'Beta', 'Menu keyboard navigation preserves search text'); await esc(p, before);
  });
  await test('hidden-pivot-headers-ignore-row-column-but-page-stays-available', async (p, info) => {
    info.fixture = await fixture(p, 'pivot', { hiddenHeaders: true }); eq(await p.locator('.pbtn[data-k=rows],.pbtn[data-k=cols]').count(), 0, 'Hidden field captions remove row/column buttons');
    await selectButton(p, 'rows', 'Group'); await altDown(p); eq(await p.locator('.pivot-filter-menu').count(), 0, 'Hidden header cell opens no field filter'); await p.keyboard.press('Escape');
    await selectButton(p, 'page', 'Week'); const before = await book(p); await open(p, 'pivot'); await esc(p, before);
  });
  await test('pivot-sigma-with-column-field-opens-the-mouse-field-menu', async (p, info) => {
    info.fixture = await fixture(p, 'pivot');
    info.point = await p.evaluate(() => { const t = tabula, b = t.wb().sheets[t.si].pivot.buttons.find(b => b.sigma); if (!b) throw Error('Missing sigma button'); t.selectCell(b.r,b.c); return b; });
    await p.locator('#cellEditor').focus(); const before=await book(p), menu=await open(p,'pivot');
    eq(await menu.locator('.filter-window input').evaluateAll(nodes=>nodes.map(n=>n.getAttribute('aria-label'))), ['Offline','Online'], 'Sigma with columns exposes the actual Channel field'); await esc(p,before);
  });
  for (const type of ['ordinary','table','pivot']) await test(type+'-mouse-filter-retains-search-focus', async (p,info)=>{
    info.fixture=await fixture(p,type); await selectButton(p,type==='pivot'?'rows':null,'Group'); const before=await book(p);
    const button=type==='pivot'?p.locator('.pbtn[data-k="rows"][data-f="Group"]').first():p.locator('.fbtn:not(.pbtn)[data-c="2"]').first(); await button.click();
    await p.locator('.menu:has(.filter-menu)').waitFor(); await p.waitForFunction(()=>document.activeElement?.matches('input[type=search]')&&!!document.activeElement.closest('.menu'));
    eq((await state(p)).focus.type,'search','Mouse filter opening keeps the existing search focus'); await esc(p,before);
  });
  for(const focus of ['#ribbon [data-ribbon-command="copy"]','#quickAccess button:not(:disabled)']) await test('ordinary-AltDown-relay-focus-'+focus,async(p,info)=>{
    info.fixture=await fixture(p); await selectButton(p); const before=await book(p); await p.locator(focus).first().focus(); info.focused=await state(p); await open(p,'ordinary'); await esc(p,before);
  });
  await test('ordinary-physical-ArrowDown-Process229-and-composition-guard',async(p,info)=>{
    info.fixture=await fixture(p); await selectButton(p); const before=await book(p);
    info.composing=await p.evaluate(()=>{const e=new KeyboardEvent('keydown',{key:'Process',code:'ArrowDown',keyCode:229,altKey:true,isComposing:true,bubbles:true,cancelable:true});document.activeElement.dispatchEvent(e);return e.defaultPrevented;}); await raf(p);
    eq(await p.locator('.filter-menu').count(),0,'Composing physical arrow never opens a filter');eq(await book(p),before,'Composing arrow preserves data and Undo');
    info.event=await p.evaluate(()=>{const e=new KeyboardEvent('keydown',{key:'Process',code:'ArrowDown',keyCode:229,altKey:true,isComposing:false,bubbles:true,cancelable:true});document.activeElement.dispatchEvent(e);return e.defaultPrevented;}); await raf(p);
    await p.locator('.menu:has(.filter-menu)').waitFor({timeout:3000}); eq(info.event,true,'Physical arrow fallback is handled when composition is inactive'); await esc(p,before);
  });
  await test('ordinary-cell-edit-and-dialog-AltDown-are-isolated',async(p,info)=>{
    info.fixture=await fixture(p);await selectButton(p);const before=await book(p);await p.keyboard.press('F2');await p.locator('#cellEditor').fill('temporary text');await altDown(p);
    eq((await state(p)).editing,true,'Editing shortcut stays in cell editing');eq(await p.locator('.filter-menu').count(),0,'Editing opens no filter');eq(await p.locator('#cellEditor').inputValue(),'temporary text','Editing text is preserved');await p.keyboard.press('Escape');eq(await book(p),before,'Cancel editing preserves the workbook');
    await p.keyboard.press('Control+g');const dialog=p.getByRole('dialog').last();await dialog.waitFor();const input=dialog.locator('input').first();await input.focus();await input.fill('C5');await altDown(p);
    eq(await p.locator('.filter-menu').count(),0,'Dialog input opens no grid filter');eq(await input.inputValue(),'C5','Dialog input text stays intact');eq(await book(p),before,'Dialog shortcut makes no workbook change');await p.keyboard.press('Escape');
  });
  await test('ordinary-body-AltDown-keeps-cell-pick-list',async(p,info)=>{
    info.fixture=await fixture(p);await p.evaluate(()=>{tabula.selectCell(7,2);tabula.gv().ensureVisible(7,2);});await p.locator('#cellEditor').focus();const before=await book(p);await altDown(p);
    await p.locator('.cell-pick-list').waitFor();eq(await p.locator('.filter-menu').count(),0,'Body cell opens its existing data-entry list');eq(await book(p),before,'Pick list opening preserves filters and Undo');await p.keyboard.press('Escape');eq(await book(p),before,'Pick list cancellation preserves original data');
  });
  for(const type of ['ordinary','table']) await test(type+'-open-menu-protection-change-and-version-change-reject-stale-apply',async(p,info)=>{
    info.fixture=await fixture(p,type);await selectButton(p);let menu=await open(p,type);await menu.getByRole('searchbox').fill('Beta');
    await p.evaluate(()=>{tabula.wb().sheets[tabula.si].protect={on:true,allow:{selectLocked:true,autoFilter:false}};});const protectedBook=await book(p);await p.keyboard.press('Enter');await raf(p);eq(await book(p),protectedBook,'Protection changed after opening blocks stale confirmation');await p.keyboard.press('Escape');
    await p.evaluate(()=>{delete tabula.wb().sheets[tabula.si].protect;});await selectButton(p);menu=await open(p,type);await menu.getByRole('searchbox').fill('Beta');
    await p.evaluate(()=>{const w=tabula.wb();w.transact(()=>w.setInput(tabula.si,0,0,'changed independently'));});const changed=await book(p);await menu.getByRole('searchbox').focus();await p.keyboard.press('Enter');await raf(p);
    eq(await book(p),changed,'Version changed after opening blocks stale filter confirmation');await p.keyboard.press('Escape');
  });
  for(const type of ['ordinary','table','pivot']) await test(type+'-keyboard-command-ArrowDown-Enter-sort-and-Undo',async(p,info)=>{
    info.fixture=await fixture(p,type);await selectButton(p,type==='pivot'?'rows':null,'Group');const before=await book(p);await open(p,type,{search:false});
    await p.keyboard.press('ArrowDown');ok(/내림차순/.test((await state(p)).focus.text),'ArrowDown reaches descending sort');await p.keyboard.press('Enter');await raf(p);
    eq((await state(p)).undo,before.undo+1,'Sort command adds exactly one Undo');eq((await state(p)).menus,0,'Sort command closes the popup');
    const row=await p.evaluate(type=>{const t=tabula,d=t.wb().sheets[t.si].pivot;return type==='pivot'?t.wb().getValue(t.si,d.buttons.find(b=>b.kind==='rows').r+1,__filterFixture.c0):t.wb().getValue(t.si,__filterFixture.r0+1,__filterFixture.c0);},type);
    eq(row,'Gamma','Descending sort puts Gamma before Beta and Alpha');if(type==='pivot')eq(await p.evaluate(()=>tabula.wb().sheets[tabula.si].pivot.sort.Group.dir),'desc','Pivot preserves the selected sort field and direction');
    await run(p,'undo');eq((await book(p)).sheets,before.sheets,'Undo restores exact sort metadata and original row order');
  });
  for(const type of ['ordinary','table']) for(const guard of ['protected','markedFinal']) await test(type+'-'+guard+'-filter-menu-actions-start-disabled',async(p,info)=>{
    info.fixture=await fixture(p,type);await selectButton(p);await p.evaluate(guard=>{const t=tabula,w=t.wb();if(guard==='protected')w.sheets[t.si].protect={on:true,allow:{selectLocked:true,autoFilter:false}};else w.props.markedFinal=true;},guard);
    const before=await book(p);await altDown(p);const menu=p.locator('.menu:has(.filter-menu)').first();await menu.waitFor({timeout:3000});
    eq(await menu.locator('.filter-foot .primary').isDisabled(),true,'Confirmation starts disabled without filter permission');eq(await menu.locator('.menu-item:not(:disabled)').count(),0,'Every filter/sort command starts disabled without permission');
    await p.keyboard.press('Enter');await raf(p);eq(await book(p),before,'Enter in protected menu preserves all cells, criteria and Undo');await p.keyboard.press('Escape');eq(await book(p),before,'Cancelling the protected popup preserves the workbook');
  });
  await test('legacy-value-only-sigma-header-is-not-a-data-field-filter', async (p, info) => {
    info.fixture = await fixture(p, 'pivot', { sigmaOnly: true });
    await p.evaluate(() => { const t=tabula,d=t.wb().sheets[t.si].pivot;d.buttons.push({r:d.area.r2,c:d.area.c2,kind:'cols',sigma:true});t.gv().layout(); });
    eq(await p.locator('.pbtn[data-k=cols]').count(),1,'Legacy value-only sigma metadata renders one axis button');
    const point = await p.evaluate(() => { const t = tabula, b = t.wb().sheets[t.si].pivot.buttons.find(b => b.sigma); if (!b) throw Error('Missing sigma fixture button'); t.selectCell(b.r, b.c); return b; });
    info.point = point; await p.locator('#cellEditor').focus(); const before = await book(p); await altDown(p); eq(await p.locator('.pivot-filter-menu').count(), 0, 'Sigma-only values header opens no data-field filter'); eq(await book(p), before, 'Sigma key gesture preserves definitions, cells and Undo');
  });
} finally {
  await browser.close(); const summary = { url, engine, baseline: baseline || null, syntheticOnly: true, assets: [...assets], cases: results.length, passed: results.filter(result => result.ok).length, checks, pageErrors: errors, remoteWrites: writes, blockedRequests: blocked, results };
  await writeFile(path.join(out, 'result.json'), JSON.stringify(summary, null, 2)); console.log(JSON.stringify({ cases: summary.cases, passed: summary.passed, checks, pageErrors: errors.length, remoteWrites: writes.length, assets: summary.assets, out }));
}
if (results.some(result => !result.ok) || errors.length || writes.length) process.exitCode = 1;
