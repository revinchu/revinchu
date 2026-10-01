// 지우기 메뉴·하이퍼링크·보호·실행 취소 브라우저 회귀. 합성 데이터만 사용한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const results = [];
async function test(name, run) {
  const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(process.env.WIXEL_URL || 'http://localhost:5180/');
    await page.waitForFunction(() => !!window.tabula);
    await run(page); assert.deepEqual(errors, []);
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (e) { results.push({ name, ok: false }); console.error(`NG ${name}: ${e.stack}`); }
  finally { await page.close(); }
}
const setup = async (page) => page.evaluate(() => {
  const t = window.tabula; const wb = t.wb();
  wb.transact(() => {
    wb.setCellData(0, 0, 0, { raw: '링크 글자', link: 'https://example.com/one', comment: '첫 메모', style: { color: '#0563c1', underline: true, fill: '#ffeecc', bold: true, numFmt: 'custom', code: '0.00', cellStyleName: '사용자' } });
    wb.setCellData(0, 1, 0, { raw: '링크 없는 글자', style: { fill: '#ffccff', italic: true } });
    wb.setCellData(0, 2, 0, { raw: '=HYPERLINK("https://example.com","함수 링크")', style: { color: '#0563c1', underline: true } });
    wb.setCellData(0, 3, 0, { raw: '범위 밖', link: 'https://example.com/out', style: { fill: '#abcdef' } });
  });
  t.selectRange({ r1: 0, c1: 0, r2: 2, c2: 0 });
});
const snapshot = (page) => page.evaluate(() => {
  const cells = window.tabula.wb().serialize().sheets[0].cells;
  return [0, 1, 2, 3].map((r) => cells[`${r},0`] ?? null);
});
const menu = (page) => page.evaluate(() => window.tabula.openNamedMenu('clear', { x: 160, y: 180 }));
try {
  await test('홈 지우기: 링크만 삭제·모든 내용과 서식·HYPERLINK 수식 보존·실행 취소', async (p) => {
    await setup(p); const before = await snapshot(p);
    await menu(p);
    assert.equal(await p.getByRole('menuitem').count(), 6);
    await p.getByText('하이퍼링크 지우기', { exact: true }).click();
    const after = await snapshot(p); const expected = structuredClone(before); delete expected[0].link;
    assert.deepEqual(after, expected);
    assert.match(await p.locator('#toast').textContent(), /HYPERLINK 수식은 유지/);
    await p.evaluate(() => window.tabula.run('undo')); assert.deepEqual(await snapshot(p), before);
    await p.evaluate(() => window.tabula.run('redo')); assert.deepEqual(await snapshot(p), expected);
  });
  await test('제거는 실제 링크 셀 서식·조건부 서식만 초기화하고 그림·값·메모를 유지', async (p) => {
    await setup(p);
    await p.evaluate(() => {
      const wb = window.tabula.wb();
      wb.transact(() => {
        wb.setLineStyle(0, 'row', 0, { bold: true, fill: '#998877', gradient: { stops: [[0, '#000000'], [1, '#ffffff']] } });
        wb.setCellData(0, 0, 0, { ...wb.getCell(0, 0, 0), image: { src: 'data:image/png;base64,AA==', alt: '보존할 그림' } });
        wb.setSheetProp(0, 'cond', [{ r1: 0, c1: 0, r2: 3, c2: 0, type: 'formula', formula: '=A1<>""', style: { color: '#ff0000' } }]);
      });
    });
    const before = await snapshot(p);
    const condBefore = await p.evaluate(() => window.tabula.wb().sheets[0].cond);
    await menu(p); await p.getByText('하이퍼링크 제거', { exact: true }).click();
    const after = await snapshot(p);
    assert.equal(after[0].link, undefined); assert.equal(after[0].raw, before[0].raw); assert.equal(after[0].comment, before[0].comment); assert.deepEqual(after[0].image, before[0].image);
    assert.deepEqual(after.slice(1), before.slice(1));
    const state = await p.evaluate(() => ({ style: window.tabula.wb().styleAt(0, 0, 0), rules: window.tabula.wb().sheets[0].cond }));
    assert.equal(state.style.color, '#000000'); assert.equal(state.style.fill, '#ffffff'); assert.equal(state.style.bold, false); assert.equal(state.style.underline, false);
    assert.equal(state.style.gradient, false); assert.equal(state.style.align, 'general'); assert.equal(state.style.numFmt, 'general'); assert.equal(state.style.cellStyleName || '', '');
    assert.equal(state.rules[0].r1, 1); assert.equal(state.rules[0].r2, 3); assert.equal(state.rules[0].formula, '=A2<>""');
    await p.evaluate(() => window.tabula.run('undo')); assert.deepEqual(await snapshot(p), before); assert.deepEqual(await p.evaluate(() => window.tabula.wb().sheets[0].cond), condBefore);
    await p.evaluate(() => window.tabula.run('redo')); assert.equal((await snapshot(p))[0].link, undefined);
  });
  await test('비연속 선택은 가운데 링크를 보존하고 전체 열 선택도 기존 링크만 처리', async (p) => {
    for (const cmd of ['clearHyperlinks', 'removeHyperlink']) {
      await p.evaluate((command) => {
        const t = window.tabula; const wb = t.wb();
        wb.transact(() => { for (let r = 0; r < 3; r++) wb.setCellData(0, r, 0, { raw: String(r), link: `https://example.com/${r}`, style: { fill: '#abcdef' }, ...(r !== 1 ? { comment: '선택' } : {}) }); });
        t.selectCell(0, 0); t.run('selectComments'); t.run(command);
      }, cmd);
      assert.deepEqual(await p.evaluate(() => [0, 1, 2].map((r) => window.tabula.wb().getCell(0, r, 0).link ?? null)), [null, 'https://example.com/1', null]);
      assert.equal(await p.evaluate(() => window.tabula.wb().getCell(0, 1, 0).style.fill), '#abcdef');
      await p.evaluate(() => window.tabula.run('undo'));
    }
    await p.evaluate(() => { const t = window.tabula; t.selectRange({ r1: 0, c1: 0, r2: 1048575, c2: 0 }, 'cols'); t.run('clearHyperlinks'); });
    assert.deepEqual(await p.evaluate(() => [0, 1, 2].map((r) => window.tabula.wb().getCell(0, r, 0).link ?? null)), [null, null, null]);
  });
  await test('잠긴 링크는 차단·잠금해제 링크만 삭제·서식 제거에는 서식 변경 권한 필요', async (p) => {
    await setup(p);
    await p.evaluate(() => { const t = window.tabula; t.selectCell(0, 0); t.wb().transact(() => t.wb().setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true } })); t.run('clearHyperlinks'); });
    assert.equal((await snapshot(p))[0].link, 'https://example.com/one');
    await p.keyboard.press('Escape');
    await p.evaluate(() => { const t = window.tabula; t.wb().transact(() => t.wb().setStyle(0, 0, 0, { locked: false })); t.run('clearHyperlinks'); });
    assert.equal((await snapshot(p))[0].link, undefined);
    await p.evaluate(() => { const t = window.tabula; t.run('undo'); t.run('removeHyperlink'); });
    assert.equal((await snapshot(p))[0].link, 'https://example.com/one');
    await p.keyboard.press('Escape');
    await p.evaluate(() => { const t = window.tabula; t.wb().transact(() => t.wb().setSheetProp(0, 'protect', { on: true, allow: { formatCells: true } })); t.run('removeHyperlink'); });
    assert.equal((await snapshot(p))[0].link, undefined);
  });
  await test('기존 지우기 4종은 값·서식·메모·링크의 대상만 지우고 실행 취소', async (p) => {
    for (const cmd of ['clearAll', 'clearFormats', 'clearContents', 'clearComments']) {
      await setup(p); const before = await snapshot(p);
      await p.evaluate((command) => { window.tabula.selectCell(0, 0); window.tabula.run(command); }, cmd);
      const after = await snapshot(p);
      assert.deepEqual(after.slice(1), before.slice(1));
      if (cmd === 'clearAll') assert.equal(after[0], null);
      if (cmd === 'clearFormats') { const expected = { ...before[0] }; delete expected.style; assert.deepEqual(after[0], expected); }
      if (cmd === 'clearContents') assert.deepEqual(after[0], { raw: '', style: before[0].style, comment: before[0].comment });
      if (cmd === 'clearComments') { const expected = { ...before[0] }; delete expected.comment; assert.deepEqual(after[0], expected); }
      await p.evaluate(() => window.tabula.run('undo')); assert.deepEqual(await snapshot(p), before);
    }
  });
  await test('우클릭 메뉴도 보호 경로를 거쳐 같은 하이퍼링크 명령 실행', async (p) => {
    await setup(p); await p.evaluate(() => window.tabula.selectCell(0, 0));
    const box = await p.locator('.c[data-r="0"][data-c="0"]').boundingBox();
    await p.mouse.click(box.x + 15, box.y + box.height / 2, { button: 'right' });
    await p.getByText('하이퍼링크 지우기', { exact: true }).click();
    const cell = (await snapshot(p))[0]; assert.equal(cell.link, undefined); assert.equal(cell.style.fill, '#ffeecc');
  });
  await test('서식·링크 제거는 앞자리 0과 문자 수식을 유지하고 비연속 선택도 같은 동작', async (p) => {
    const seed = async () => p.evaluate(() => {
      const t = window.tabula, wb = t.wb();
      wb.transact(() => {
        for (const [r, raw] of ['00123', '=1+1', '가운데 보존'].entries()) wb.setCellData(0, r, 0, { raw, link: 'https://example.com', style: { numFmt: 'text', bold: true }, ...(r < 2 ? { comment: '선택' } : {}) });
      });
      t.selectRange({ r1: 0, c1: 0, r2: 1, c2: 0 });
    });
    const read = () => p.evaluate(() => [0, 1, 2].map((r) => window.tabula.wb().getValue(0, r, 0)));
    await seed(); await menu(p); await p.getByText('하이퍼링크 제거', { exact: true }).click();
    assert.deepEqual(await read(), ['00123', '=1+1', '가운데 보존']);
    assert.equal(await p.evaluate(() => !!window.tabula.wb().getCell(0, 1, 0).formula), false);
    await p.evaluate(() => { window.tabula.run('undo'); window.tabula.run('redo'); });
    assert.deepEqual(await read(), ['00123', '=1+1', '가운데 보존']);
    await seed();
    await p.evaluate(() => { const t = window.tabula; t.selectCell(0, 0); t.run('selectComments'); t.run('clearFormats'); });
    assert.deepEqual(await read(), ['00123', '=1+1', '가운데 보존']);
    assert.equal(await p.evaluate(() => window.tabula.wb().getCell(0, 2, 0).style.numFmt), 'text');
  });
  await test('텍스트 서식·서식만 붙여넣기에서도 숫자는 숫자로 유지하고 수식 문자열 복사는 이동하지 않음', async (p) => {
    await p.evaluate(() => {
      const t = window.tabula, wb = t.wb();
      wb.transact(() => { wb.setInput(0, 0, 0, '123'); wb.setInput(0, 0, 1, '서식 원본'); wb.setStyle(0, 0, 1, { numFmt: 'text' }); });
      t.selectCell(0, 0); t.run('numFmt', 'text');
    });
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 123);
    await p.evaluate(() => { const t = window.tabula; t.run('undo'); t.selectCell(0, 1); t.run('copy'); t.selectCell(0, 0); t.run('pasteSpecial'); });
    await p.locator('input[name="psWhat"][value="formats"]').check();
    await p.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 123);
    await p.evaluate(() => {
      const t = window.tabula, wb = t.wb();
      wb.transact(() => wb.setCellData(0, 1, 0, { raw: '=A1+1', style: { numFmt: 'text' } }));
      t.selectCell(1, 0); t.run('clearFormats'); t.run('copy'); t.selectCell(2, 0); t.run('pasteSpecial');
    });
    await p.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 2, 0)), '=A1+1');
  });
  await test('메모만 붙여넣기는 문자 수식·캐시·링크·그림·fx를 보존한다', async (p) => {
    await p.evaluate(() => {
      const t = window.tabula, wb = t.wb();
      wb.transact(() => {
        wb.setCellData(0, 0, 0, { raw: '=1+1', style: { numFmt: 'text' }, link: 'https://example.com/keep', image: { src: 'data:image/png;base64,AA==', alt: '보존' } });
        wb.setStyle(0, 0, 0, { numFmt: 'general' });
        wb.setCellData(0, 1, 0, { raw: '=NO_SUCH_FUNCTION(1)', cached: 42, fx: true, style: { numFmt: 'text' }, link: 'https://example.com/formula' });
        wb.setCellData(0, 0, 1, { raw: '메모 원본', comment: '붙여 넣은 메모' });
      });
      t.selectCell(0, 1); t.run('copy');
    });
    const before = await snapshot(p);
    for (const r of [0, 1]) {
      await p.evaluate((row) => { const t = window.tabula; t.selectCell(row, 0); t.run('pasteSpecial'); }, r);
      await p.locator('input[name="psWhat"][value="comments"]').check();
      await p.getByRole('button', { name: '확인', exact: true }).click();
    }
    const after = await snapshot(p);
    for (const r of [0, 1]) assert.deepEqual(after[r], { ...before[r], comment: '붙여 넣은 메모' });
    assert.deepEqual(await p.evaluate(() => [0, 1].map((r) => window.tabula.wb().getValue(0, r, 0))), ['=1+1', 42]);
    await p.evaluate(() => { const t = window.tabula; t.run('undo'); t.run('undo'); });
    assert.deepEqual(await snapshot(p), before);
    await p.evaluate(() => { const t = window.tabula; t.run('redo'); t.run('redo'); });
    assert.deepEqual(await snapshot(p), after);
  });
  await test('Ctrl+K는 표시 문구가 같으면 문자 수식과 파일 캐시를 보존하고 변경한 문구만 새 입력으로 처리', async (p) => {
    await p.evaluate(() => {
      const t = window.tabula, wb = t.wb();
      wb.transact(() => {
        wb.setCellData(0, 0, 0, { raw: '=1+1', style: { numFmt: 'text' }, comment: '유지할 메모' });
        wb.setStyle(0, 0, 0, { numFmt: 'general' });
        wb.setCellData(0, 1, 0, { raw: '=NO_SUCH_FUNCTION(1)', cached: 42, fx: true, style: { numFmt: 'text' } });
      });
    });
    for (const r of [0, 1]) {
      await p.evaluate((row) => window.tabula.selectCell(row, 0), r);
      await p.keyboard.press('Control+k');
      await p.getByLabel('주소 (웹 주소 또는 #시트!A1)').fill('https://example.com/new');
      await p.getByRole('button', { name: '확인', exact: true }).click();
    }
    assert.deepEqual(await p.evaluate(() => [0, 1].map((r) => window.tabula.wb().getValue(0, r, 0))), ['=1+1', 42]);
    const after = await snapshot(p);
    assert.equal(after[0].inputType, 'text'); assert.equal(after[0].comment, '유지할 메모');
    assert.equal(after[1].cached, 42); assert.equal(after[1].fx, true);
    assert.ok(after.slice(0, 2).every((cell) => cell.link === 'https://example.com/new'));
    await p.evaluate(() => { const t = window.tabula; t.run('undo'); t.run('redo'); t.selectCell(0, 0); });
    assert.deepEqual(await snapshot(p), after);
    await p.keyboard.press('Control+k');
    await p.getByLabel('표시할 텍스트').fill('=3+4');
    await p.getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 0, 0)), 7);
  });
  console.log(JSON.stringify({ tests: results.length, failed: results.filter((x) => !x.ok).length, results }, null, 2));
  if (results.some((x) => !x.ok)) process.exitCode = 1;
} finally { await browser.close(); }
