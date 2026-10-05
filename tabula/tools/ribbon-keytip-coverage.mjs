// 실제 리본 조작과 키팁 등록의 교차 검사. 합성 문서·격리 컨텍스트만 사용한다.
// IME는 DOM 이벤트 회귀이며 실제 Windows IME 드라이버 검사는 아니다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch(), results = [], coverage = [];
let registryCounts = null;
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const wanted = process.env.WIXEL_KEYTIP_FILTER;
const screenshot = process.env.WIXEL_KEYTIP_SCREENSHOT;
async function test(name, run) {
  if (wanted && !name.includes(wanted)) return;
  const context = await browser.newContext({ viewport: { width: 1600, height: 1000 }, acceptDownloads: false });
  const p = await context.newPage(), errors = [], writes = []; p.setDefaultTimeout(10000);
  p.on('pageerror', error => errors.push(error.message));
  await context.route('**/*', route => {
    const r = route.request(), target = new URL(r.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); }
    if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) return route.abort();
    return route.continue();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    await p.evaluate(() => {
      const t = window.tabula, w = t.wb(), cells = {};
      [['지역', '수량', '단가'], ['서울', '10', '100'], ['부산', '20', '200'], ['대구', '30', '300']].forEach((row, r) => row.forEach((raw, c) => { cells[`${r},${c}`] = { raw }; }));
      w.restore({ sheets: [{ name: '단축키 합성', cells }] }); w.undoStack = []; w.redoStack = [];
      t.gv().layout(); t.gv().renderAll(); t.selectCell(1, 1);
    });
    await p.locator('#cellEditor').focus(); await run(p);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기 시도');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
const sequence = async (p, keys, start = 'Alt') => { if (start) await p.keyboard.press(start); for (const ch of keys) await p.keyboard.press(ch); };
const snapshot = p => p.evaluate(() => ({ sheet: window.tabula.wb().serialize(), selection: { ...window.tabula.sel }, active: { ...window.tabula.active } }));
const style = p => p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1));
const registry = p => p.evaluate(() => window.tabula.keytipRegistry());
const resetTips = async p => {
  for (let i = 0; i < 8; i++) {
    const open = await p.evaluate(() => document.body.classList.contains('keytips') || !!document.querySelector('.menu'));
    if (!open) break;
    await p.keyboard.press('Escape');
  }
  await p.locator('#cellEditor').focus();
};
const idleEditor = async p => {
  assert.equal(await p.locator('#cellEditor').inputValue(), '', '단축키 입력이 셀에 남음');
  assert.equal(await p.locator('#cellEditor').evaluate(e => e.classList.contains('idle')), true);
};
async function contextual(p, id) {
  if (id === 'tableDesign') {
    await p.evaluate(() => { const t = window.tabula; t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 2 }); t.run('createTable'); });
    await p.getByRole('dialog', { name: '표 만들기', exact: true }).getByRole('button', { name: '확인', exact: true }).click();
    await p.evaluate(() => window.tabula.selectCell(1, 1));
  } else if (['pivotAnalyze', 'pivotDesign', 'slicerTab'].includes(id)) {
    await p.evaluate(() => {
      const t = window.tabula, w = t.wb(); w.transact(() => { w.addSheet('피벗 합성'); w.setSheetProp(1, 'pivot', { name: '검사피벗', source: '단축키 합성', range: { r1: 0, c1: 0, r2: 3, c2: 2 }, rows: ['지역'], cols: [], values: [{ field: '수량', agg: 'sum' }], top: 0, left: 0 }); });
      t.switchSheet(1); t.run('pivotRefresh'); t.selectCell(1, 0);
    });
    if (id === 'slicerTab') {
      await p.evaluate(() => window.tabula.run('insertSlicer'));
      const dialog = p.getByRole('dialog', { name: '슬라이서 삽입', exact: true });
      await dialog.getByRole('checkbox', { name: '지역', exact: true }).check(); await dialog.getByRole('button', { name: '확인', exact: true }).click();
      await p.locator('.obj.slicer').first().click({ position: { x: 30, y: 10 } });
    }
  } else if (id === 'sparkTab') {
    await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'sparklines', [{ id: 'keytip-spark', type: 'line', items: [{ r: 1, c: 3, ref: 'B2:C2' }] }])); t.selectCell(1, 3); });
  } else if (id === 'chartDesign' || id === 'objFormat') {
    await p.evaluate(() => { const t = window.tabula; t.selectRange({ r1: 0, c1: 0, r2: 3, c2: 2 }); t.run('chartColumn'); });
  }
}
try {
  await test('등록표의 전체 경로·탭·필수 HJ HH HFC와 충돌 없음', async p => {
    const r = await registry(p), entries = r.entries; assert.ok(entries.length > 100);
    registryCounts = { entries: entries.length, controls: r.controls.length, tabs: Object.keys(r.tabs).length };
    const commands = new Set(await p.evaluate(() => window.tabula.commands()));
    assert.deepEqual(entries.filter(e => e.kind === 'command' && !commands.has(e.target)), [], '등록된 명령이 실행 표에 없음');
    assert.equal(new Set(entries.map(e => e.path)).size, entries.length, '중복 키 경로');
    const terminal = entries.map(e => e.path);
    assert.deepEqual(terminal.flatMap(a => terminal.filter(b => a !== b && b.startsWith(a)).map(b => `${a}/${b}`)), [], '명령 키가 다른 경로 접두사');
    for (const [path, kind, target] of [['hj', 'menu', 'cellStyles'], ['hh', 'menu', 'fillColor'], ['hfc', 'menu', 'fontColor']]) {
      const e = entries.find(e => e.path === path); assert.ok(e, path); assert.equal(e.kind, kind); assert.equal(e.target, target);
    }
  });
  for (const id of ['home', 'insert', 'layout', 'formulas', 'data', 'review', 'view', 'help', 'tableDesign', 'pivotAnalyze', 'pivotDesign', 'sparkTab', 'slicerTab', 'objFormat', 'chartDesign']) {
    await test(`실제 리본 ${id}: 모든 조작의 등록·작은 배지·입력란 연결`, async p => {
      await contextual(p, id); const r = await registry(p);
      const prefix = Object.entries(r.tabs).find(([, tab]) => tab === id)?.[0]; assert.ok(prefix, `${id} 탭 키 없음`);
      await sequence(p, prefix);
      if (id === 'home') {
        const overlaps = await p.locator('.keytip-badge').evaluateAll(nodes => {
          const boxes = nodes.map(n => ({ path: n.dataset.keytipPath, r: n.getBoundingClientRect() })).filter(x => x.r.width && x.r.height);
          const pairs = [];
          for (let i = 0; i < boxes.length; i++) for (let j = i + 1; j < boxes.length; j++) {
            const a = boxes[i], b = boxes[j];
            if (Math.min(a.r.right,b.r.right) - Math.max(a.r.left,b.r.left) > 0.5 && Math.min(a.r.bottom,b.r.bottom) - Math.max(a.r.top,b.r.top) > 0.5) pairs.push([a.path,b.path]);
          }
          return pairs;
        });
        assert.deepEqual(overlaps, [], '홈 배지가 서로 겹침');
      }
      const controls = r.controls.filter(control => control.tabId === id).map(control => ({ ...control, controlId: control.controlId ?? control.id })); assert.ok(controls.length > 0);
      const actual = await p.locator('#ribbon [data-ribbon-control]').evaluateAll(nodes => nodes.flatMap(n => (n.dataset.ribbonControls || n.dataset.ribbonControl).split(' ')));
      const expected = [...new Set(controls.map(c => c.controlId))];
      assert.deepEqual([...new Set(actual)].sort(), expected.sort(), '등록표와 실제 리본 조작 목록 차이');
      const missing = await p.locator('#ribbon button,#ribbon input,#ribbon select').evaluateAll(nodes => nodes.filter(n => !n.disabled && n.getBoundingClientRect().width > 0
        && !n.closest('[data-ribbon-control]') && !n.closest('[data-ribbon-command],[data-ribbon-menu]')?.querySelector('[data-ribbon-control]'))
        .map(n => ({ tag: n.tagName, title: n.title, text: n.textContent?.trim().slice(0, 80) })));
      assert.deepEqual(missing, [], '활성 리본 조작에 등록 메타데이터가 없음');
      for (const control of controls) {
        const node = p.locator(`#ribbon [data-ribbon-controls~="${control.controlId}"],#ribbon [data-ribbon-control="${control.controlId}"]`).first();
        assert.equal(await node.count(), 1, control.controlId);
        await node.scrollIntoViewIfNeeded();
        const entry = r.entries.find(e => e.path === control.path) || r.entries.find(e => e.controlId === control.controlId);
        assert.ok(entry, `${control.controlId} 키 미등록`);
        let visual = p.locator(`.keytip-badge[data-keytip-path="${entry.path}"]`);
        let inGroup = false;
        if (!await visual.isVisible()) {
          const group = await p.locator('.keytip-badge[data-keytip-path]').evaluateAll((nodes, path) => nodes.map(n => n.dataset.keytipPath).filter(key => key !== path && path.startsWith(key)).sort((a, b) => b.length - a.length)[0], entry.path);
          if (group) {
            await sequence(p, group.slice(prefix.length), null); inGroup = true;
            visual = p.locator(`.keytip-command-menu [data-keytip-path="${entry.path}"],.keytip-badge[data-keytip-path="${entry.path}"]`);
          }
        }
        await visual.waitFor({ state: 'visible' });
        assert.ok(await visual.innerText(), `${entry.path} 빈 배지/하위 명령`);
        if (inGroup) { await resetTips(p); await sequence(p, prefix); }
        if (entry.kind === 'input') {
          await resetTips(p); await sequence(p, entry.path);
          assert.ok(await node.evaluate(el => el === document.activeElement || el.contains(document.activeElement)), `${entry.path} 입력란 포커스`);
          await resetTips(p); await sequence(p, prefix);
        }
        if (entry.kind === 'menu') {
          await resetTips(p); await sequence(p, entry.path);
          assert.ok(await p.locator('.menu:visible,.dialog:visible').count(), `${entry.path} ${entry.target} 메뉴가 열리지 않음`);
          await resetTips(p); await sequence(p, prefix);
        }
      }
      coverage.push({ tab: id, controls: expected.length, menusOpened: controls.filter(c => c.kind === 'menu').length, inputsFocused: controls.filter(c => c.kind === 'input').length });
    });
  }
  await test('HJ 셀 스타일 검색·방향키 적용·Undo·Esc 무변경', async p => {
    const before = await snapshot(p); await sequence(p, 'hj');
    const search = p.getByRole('searchbox', { name: '셀 스타일 검색' }); await search.waitFor();
    if (screenshot) await p.screenshot({ path: screenshot });
    assert.ok(await search.evaluate(e => e === document.activeElement)); await p.keyboard.press('ArrowDown');
    assert.ok(await p.evaluate(() => document.activeElement.matches('.style-chip'))); await p.keyboard.press('ArrowRight');
    const selected = await p.evaluate(() => document.activeElement.dataset.cellStyle); await p.keyboard.press('Enter');
    assert.equal((await style(p)).cellStyleName, selected); await p.keyboard.press('Control+z'); assert.deepEqual(await snapshot(p), before);
    await sequence(p, 'hj'); await search.fill('없는스타일__'); assert.equal(await p.locator('.style-chip:visible').count(), 0);
    await search.fill(''); assert.ok(await p.locator('.style-chip:visible').count() >= 47);
    await p.keyboard.press('Escape'); assert.deepEqual(await snapshot(p), before); await idleEditor(p);
  });
  for (const [keys, prop] of [['hh', 'fill'], ['hfc', 'color']]) {
    await test(`${keys.toUpperCase()} 색 팔레트 방향키·Enter 적용·한 번 Undo·Escape`, async p => {
      const before = await snapshot(p); await sequence(p, keys); await p.locator('.palette').waitFor();
      await p.locator('.palette .swatch').last().focus(); await p.keyboard.press('ArrowDown');
      assert.ok(await p.evaluate(() => document.activeElement.matches('.color-palette-menu > .menu-item')), '마지막 색 행 아래로 하단 명령 이동');
      await p.locator('.palette .swatch').first().focus(); await p.keyboard.press('ArrowUp');
      if (keys === 'hfc') assert.ok(await p.evaluate(() => document.activeElement.matches('.color-palette-menu > .menu-item') && document.activeElement.textContent.includes('자동')), '첫 색 행 위로 자동 글꼴색 이동');
      else assert.ok(await p.locator('.palette .swatch').first().evaluate(el => el === document.activeElement), '상단 명령 없는 첫 색 위에서는 첫 견본 유지');
      await p.keyboard.press('ArrowDown'); await p.keyboard.press('ArrowRight');
      assert.ok(await p.evaluate(() => document.activeElement.matches('.swatch')), '방향키가 실제 색 견본에 도달하지 않음');
      const color = await p.evaluate(() => document.activeElement.getAttribute('title')); await p.keyboard.press('Enter');
      assert.equal(String((await style(p))[prop]).toLowerCase(), color.toLowerCase()); await p.keyboard.press('Control+z');
      assert.deepEqual(await snapshot(p), before); await sequence(p, keys); await p.locator('.palette').waitFor();
      await p.keyboard.press('Escape'); assert.deepEqual(await snapshot(p), before); await idleEditor(p);
    });
  }
  await test('F10·Alt를 누른 채 HH/HFC/HJ 실행하고 Esc로 취소', async p => {
    const before = await snapshot(p);
    for (const path of ['hh', 'hfc', 'hj']) {
      await sequence(p, path, 'F10'); assert.ok(await p.locator('.palette,.style-chip').count()); await resetTips(p);
      await p.keyboard.down('Alt'); for (const key of path) await p.keyboard.press(key); await p.keyboard.up('Alt');
      assert.ok(await p.locator('.palette,.style-chip').count()); await resetTips(p);
    }
    assert.deepEqual(await snapshot(p), before); await idleEditor(p);
  });
  await test('조합 중인 키는 리본 명령이나 셀 입력으로 실행하지 않음', async p => {
    const before = await snapshot(p); await p.keyboard.press('Alt');
    await p.evaluate(() => {
      const ed = document.getElementById('cellEditor');
      for (const [key, code] of [['ㅗ', 'KeyH'], ['ㄹ', 'KeyF'], ['ㅊ', 'KeyC']]) ed.dispatchEvent(new KeyboardEvent('keydown', { key, code, keyCode: 229, isComposing: true, bubbles: true, cancelable: true }));
    });
    assert.equal(await p.locator('.palette').count(), 0); assert.equal((await p.evaluate(() => document.body.dataset.keytipSequence)), '');
    await p.keyboard.press('Escape'); await idleEditor(p); assert.deepEqual(await snapshot(p), before);
  });
  await test('한글 물리 키·지연 composition은 팔레트 취소 뒤 셀/선택을 오염하지 않음', async p => {
    const before = await snapshot(p); await p.keyboard.press('Alt');
    await p.evaluate(() => {
      const ed = document.getElementById('cellEditor');
      for (const [key, code] of [['ㅗ', 'KeyH'], ['ㄹ', 'KeyF'], ['ㅊ', 'KeyC']]) ed.dispatchEvent(new KeyboardEvent('keydown', { key, code, isComposing: false, bubbles: true, cancelable: true }));
    });
    await p.locator('.palette').waitFor(); await p.keyboard.press('Escape');
    await p.evaluate(() => {
      const ed = document.getElementById('cellEditor');
      ed.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true, data: '' }));
      ed.value = 'ㅊㅎ'; ed.dispatchEvent(new InputEvent('input', { bubbles: true, inputType: 'insertCompositionText', data: 'ㅊㅎ', isComposing: true }));
      ed.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: 'ㅊㅎ' }));
      ed.dispatchEvent(new KeyboardEvent('keyup', { key: 'Process', code: 'KeyC', keyCode: 229, bubbles: true }));
    });
    await idleEditor(p); assert.deepEqual(await snapshot(p), before);
    await p.keyboard.type('next'); await p.keyboard.press('Enter'); assert.equal(await p.evaluate(() => window.tabula.wb().getRaw(0, 1, 1)), 'next');
  });
  const failed = results.filter(result => !result.ok).length;
  console.log(JSON.stringify({ url, tests: results.length, failed, registryCounts, coverage, results }, null, 2));
  if (failed) process.exitCode = 1;
} finally { await browser.close(); }
