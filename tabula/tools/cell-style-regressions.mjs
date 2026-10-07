// 합성 문서만 사용하는 셀 스타일 UI 회귀. 실제 사용자 파일은 읽지 않는다.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
import { cellStylePatch } from '../src/cell-style.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
page.setDefaultTimeout(10000);
const errors = []; page.on('pageerror', (e) => errors.push(e.message));
let groups = 0;
const ok = (s) => { groups++; console.log(`OK ${s}`); };
const menu = async () => { await page.evaluate(() => window.tabula.openNamedMenu('cellStyles', { x: 120, y: 110 })); };
const context = async (name, action) => { await menu(); await page.locator(`[data-cell-style="${name}"]`).click({ button: 'right' }); await page.getByRole('menuitem', { name: action, exact: true }).click(); };
const dialog = (name) => page.getByRole('dialog', { name, exact: true });
const state = () => page.evaluate(() => ({ styles: window.tabula.wb().cellStyles, book: window.tabula.wb().serialize() }));
const color = async (outer, value) => {
  await outer.getByRole('button', { name: '서식...', exact: true }).click();
  const fmt = dialog('셀 서식'); await fmt.getByRole('tab', { name: '글꼴', exact: true }).click();
  await fmt.getByLabel('색', { exact: true }).fill(value);
  await fmt.getByRole('button', { name: '확인', exact: true }).click();
};
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(process.env.WIXEL_URL || 'http://localhost:5180/', { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => !!window.tabula, null, { timeout: 60000 });
  await page.evaluate(() => {
    const t = window.tabula, w = t.wb();
    w.transact(() => { w.setInput(0, 0, 0, '원본'); w.setStyle(0, 0, 0, { color: '#112233', bold: false, size: 14 }); w.setInput(0, 1, 0, '대상'); w.setStyle(0, 1, 0, { bold: true, fill: '#ffdddd', align: 'right', numFmt: 'currency' }); });
    t.selectCell(0, 0);
  });
  const beforeCancel = await state();
  await menu(); await page.getByRole('menuitem', { name: '새 셀 스타일...', exact: true }).click();
  await dialog('스타일').getByLabel('스타일 이름', { exact: true }).fill('취소');
  await color(dialog('스타일'), '#aabbcc');
  await dialog('스타일').getByRole('button', { name: '취소', exact: true }).click();
  assert.deepEqual(await state(), beforeCancel); ok('서식 편집 뒤 스타일 취소는 현재 문서와 셀을 바꾸지 않음');

  await menu(); await page.getByRole('menuitem', { name: '새 셀 스타일...', exact: true }).click();
  await dialog('스타일').getByLabel('스타일 이름', { exact: true }).fill('본문');
  for (const label of ['표시 형식', '맞춤', '테두리', '채우기', '보호']) await dialog('스타일').getByRole('checkbox', { name: label, exact: true }).uncheck();
  await dialog('스타일').getByRole('button', { name: '확인', exact: true }).click();
  assert.equal((await state()).styles.find((s) => s.name === '본문').include.fill, false);
  await page.evaluate(() => window.tabula.selectCell(1, 0)); await menu(); await page.locator('[data-cell-style="본문"]').click();
  let st = await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0));
  assert.equal(st.bold, false); assert.equal(st.color, '#112233'); assert.equal(st.fill, '#ffdddd'); assert.equal(st.numFmt, 'currency'); assert.equal(st.align, 'right'); assert.equal(st.cellStyleName, '본문');
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).bold), true);
  await page.evaluate(() => window.tabula.run('redo')); ok('6개 포함 항목 선택·적용·기존 서식 유지·실행 취소');

  const fixturePatch = cellStylePatch((await state()).styles.find((s) => s.name === '본문'));
  await page.evaluate((patch) => {
    const w = window.tabula.wb();
    w.transact(() => {
      w.setLineStyle(0, 'col', 2, patch); w.setLineStyle(0, 'row', 5, patch);
      w.setSheetProp(0, 'blocks', [{ r0: 10, c0: 4, n: 2, cols: [{ num: new Float64Array([11, 12]), str: null, dict: [], fmt: patch }] }]);
    });
    // 실제 가져오기처럼 clean 파일 저장값을 주입합니다. 입력 무효화로 stale 표시한 캐시는 별도 신뢰성 회귀 대상입니다.
    const data = w.serialize(); data.sheets[0].fileValues = true; data.sheets[0].cells['2,0'] = { raw: '=1+1', cached: 42, style: { ...patch, color: '#cc00cc' } }; w.restore(data);
  }, fixturePatch);
  assert.equal(await page.evaluate(() => window.tabula.wb().getValue(0, 2, 0)), 42);
  await context('본문', '수정...'); await dialog('스타일 수정').getByLabel('스타일 이름', { exact: true }).fill('본문 수정');
  await color(dialog('스타일 수정'), '#008844');
  await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
  const modified = await page.evaluate(() => { const w = window.tabula.wb(); return { a: w.styleAt(0, 1, 0), b: w.getCell(0, 2, 0), col: w.sheets[0].colStyles[2], row: w.sheets[0].rowStyles[5], block: w.sheets[0].blocks[0].cols[0].fmt }; });
  for (const style of [modified.a, modified.col, modified.row, modified.block]) { assert.equal(style.color, '#008844'); assert.equal(style.cellStyleName, '본문 수정'); }
  assert.equal(modified.b.style.color, '#cc00cc'); assert.equal(modified.b.cached, 42);
  assert.equal(modified.b.raw, '=1+1');
  assert.equal(await page.evaluate(() => window.tabula.wb().getValue(0, 2, 0)), 42);
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().sheets[0].blocks[0].cols[0].fmt.cellStyleName), '본문');
  await page.evaluate(() => window.tabula.run('redo')); ok('스타일 수정·이름 변경이 셀/행/열/블록에 반영되고 직접 서식·계산 캐시 유지');

  await context('본문 수정', '복제...'); await dialog('스타일 복제').getByLabel('스타일 이름', { exact: true }).fill('본문 복제');
  await dialog('스타일 복제').getByRole('button', { name: '확인', exact: true }).click();
  assert.equal((await state()).styles.some((s) => s.name === '본문 복제'), true);
  await context('본문 수정', '삭제'); await dialog('스타일 삭제').getByRole('button', { name: '삭제', exact: true }).click();
  assert.equal((await state()).styles.some((s) => s.name === '본문 수정'), false);
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).cellStyleName), undefined);
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).cellStyleName), '본문 수정');
  await menu(); await page.locator('[data-cell-style="좋음"]').click({ button: 'right' }); assert.equal(await page.getByRole('menuitem', { name: '삭제', exact: true }).isDisabled(), true); await page.keyboard.press('Escape');
  ok('복제·사용자 스타일 삭제/복원·기본 스타일 삭제 금지');

  const source = { sheets: [{ name: '가져오면 안 됨', cells: { '0,0': { raw: '다른 데이터' } } }], cellStyles: [
    { name: '본문 수정', style: { color: '#aa0000' }, include: { number: false, alignment: false, border: false, fill: false, protection: false } },
    { name: '병합 전용', style: { bold: true, fill: '#ddffaa' } },
  ] };
  const merge = async (replace, buffer, name = '스타일.wixel') => {
    await menu(); await page.getByRole('menuitem', { name: '스타일 병합...', exact: true }).click();
    await dialog('스타일 병합').getByLabel('스타일을 가져올 파일', { exact: true }).setInputFiles({ name, mimeType: 'application/octet-stream', buffer });
    await page.waitForFunction(() => document.querySelector('[role=status]')?.textContent?.includes('스타일 2개'));
    if (replace) await dialog('스타일 병합').getByLabel('같은 이름의 스타일', { exact: true }).selectOption('replace');
    await page.screenshot({ path: process.env.WIXEL_STYLE_MERGE_SCREENSHOT || 'D:/Codex/Temp/wixel3-style-merge.png' });
    await dialog('스타일 병합').getByRole('button', { name: '병합', exact: true }).click();
  };
  const buffer = Buffer.from(JSON.stringify({ workbook: source }));
  await merge(false, buffer);
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).color), '#008844');
  assert.equal((await state()).styles.some((s) => s.name === '병합 전용'), true);
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 0, 0).raw), '원본');
  await page.evaluate(() => window.tabula.run('undo')); assert.equal((await state()).styles.some((s) => s.name === '병합 전용'), false);
  await merge(true, buffer); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).color), '#aa0000');
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).color), '#008844');
  ok('WIXEL 파일 스타일만 병합·이름 충돌 유지/덮어쓰기·현재 데이터 유지·실행 취소');

  const imported = new Workbook(); imported.restore(source);
  await merge(true, Buffer.from(writeXlsx(imported)), '스타일.xlsx');
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).color.toLowerCase()), '#aa0000');
  assert.equal(await page.evaluate(() => window.tabula.wb().sheets.length), 1);
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 0, 0).raw), '원본');
  ok('합성 XLSX의 이름 있는 스타일 병합·시트/셀 미변경');

  await context('표준', '수정...'); await color(dialog('스타일 수정'), '#334455');
  await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
  assert.equal(await page.evaluate(() => window.tabula.wb().baseStyle.color), '#334455');
  await page.evaluate(() => window.tabula.run('undo')); assert.notEqual(await page.evaluate(() => window.tabula.wb().baseStyle?.color), '#334455');
  await menu(); await page.locator('[data-cell-style="표준"]').click({ button: 'right' }); assert.equal(await page.getByRole('menuitem', { name: '삭제', exact: true }).isDisabled(), true); await page.keyboard.press('Escape');
  await context('본문 수정', '수정...'); await page.screenshot({ path: process.env.WIXEL_STYLE_SCREENSHOT || 'D:/Codex/Temp/wixel3-cell-style.png' });
  await dialog('스타일 수정').getByRole('button', { name: '취소', exact: true }).click();
  ok('표준 스타일 수정·실행 취소·삭제 금지');
  await page.evaluate(() => {
    const w = window.tabula.wb(), include = { number: true, alignment: true, font: true, border: true, fill: true, protection: true };
    w.transact(() => w.setCellStyles([...w.cellStyles,
      { name: '검사 기본 5', builtinId: 5, style: { color: '#102030' }, include },
      { name: '검사 기본 6', builtinId: 6, customBuiltin: false, style: { color: '#203040' }, include },
    ]));
  });
  const builtinBeforeOpen = await state();
  await menu(); await page.keyboard.press('Escape');
  assert.deepEqual(await state(), builtinBeforeOpen);
  for (const [name, flag] of [['검사 기본 5', undefined], ['검사 기본 6', false]]) {
    const beforeCancel = await state();
    await context(name, '수정...'); await color(dialog('스타일 수정'), '#654321');
    await dialog('스타일 수정').getByRole('button', { name: '취소', exact: true }).click();
    assert.deepEqual(await state(), beforeCancel);
    await context(name, '수정...');
    await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
    const unchanged = (await state()).styles.find(s => s.name === name);
    assert.equal(unchanged.customBuiltin, flag);
    assert.equal(Object.hasOwn(unchanged, 'customBuiltin'), flag !== undefined);
    await context(name, '수정...'); await color(dialog('스타일 수정'), '#654321');
    await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
    const changed = (await state()).styles.find(s => s.name === name);
    assert.equal(changed.customBuiltin, true); assert.equal(changed.style.color.toLowerCase(), '#654321');
    await page.evaluate(() => window.tabula.run('undo'));
    assert.equal((await state()).styles.find(s => s.name === name).customBuiltin, flag);
    await page.evaluate(() => window.tabula.run('redo'));
    assert.equal((await state()).styles.find(s => s.name === name).customBuiltin, true);
  }
  ok('기본 스타일 메뉴·취소·변경 없는 확인은 customBuiltin 보존, 실제 수정·Undo/Redo는 정확한 표시');
  await page.evaluate(() => {
    const t = window.tabula, w = t.wb();
    w.transact(() => { w.setCellStyles([...w.cellStyles, { name: '잠금 해제 스타일', style: { locked: false, hideFormula: false, color: '#0099aa' } }]); w.setStyle(0, 0, 0, { locked: true, hideFormula: true }); w.setSheetProp(0, 'protect', { on: true, allow: { formatCells: true, selectLocked: true, selectUnlocked: true } }); });
    t.selectCell(0, 0);
  });
  await menu(); await page.locator('[data-cell-style="잠금 해제 스타일"]').click();
  const protectedStyle = await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0));
  assert.equal(protectedStyle.color, '#0099aa'); assert.equal(protectedStyle.locked, true); assert.equal(protectedStyle.hideFormula, true);
  await page.evaluate(() => window.tabula.run('formatCells')); await dialog('셀 서식').getByRole('tab', { name: '보호', exact: true }).click();
  assert.equal(await dialog('셀 서식').getByRole('checkbox', { name: '잠금', exact: true }).isDisabled(), true);
  assert.equal(await dialog('셀 서식').getByRole('checkbox', { name: '숨김', exact: true }).isDisabled(), true);
  await dialog('셀 서식').getByRole('button', { name: '취소', exact: true }).click();
  await page.locator('#cellEditor').focus(); await page.keyboard.type('변경');
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 0, 0).raw), '원본');
  ok('보호된 시트의 서식 허용은 셀 잠금·수식 숨김을 바꾸거나 값 편집 권한을 넓히지 않음');
  const openAlign = async () => {
    await page.evaluate(() => window.tabula.run('formatCells'));
    await dialog('셀 서식').getByRole('tab', { name: '맞춤', exact: true }).click();
    return dialog('셀 서식').getByLabel('세로', { exact: true });
  };
  await page.evaluate(() => {
    const t = window.tabula, w = t.wb(); w.load({ baseStyle: { valign: 'middle' }, sheets: [{ name: '맞춤 검사', cells: {} }] });
    w.transact(() => { w.setInput(0, 0, 0, '12.5'); w.setStyle(0, 0, 0, { valign: 'bottom' }); w.setInput(0, 1, 0, '부모 맞춤'); }); t.selectCell(0, 0);
  });
  const cellBeforeCancel = await state();
  let vertical = await openAlign(); assert.equal(await vertical.inputValue(), '');
  await vertical.selectOption('top'); await dialog('셀 서식').getByRole('button', { name: '취소', exact: true }).click();
  assert.deepEqual(await state(), cellBeforeCancel);
  vertical = await openAlign(); assert.equal(await vertical.inputValue(), '');
  await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click();
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 0, 0).style.valign), 'bottom');
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'bottom');
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'bottom');
  await page.evaluate(() => window.tabula.run('redo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'bottom');
  await page.evaluate(() => window.tabula.selectCell(1, 0));
  vertical = await openAlign(); assert.equal(await vertical.inputValue(), 'middle'); await vertical.selectOption('');
  await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click();
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 1, 0).style.valign), 'bottom');
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).valign), 'bottom');
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).valign), 'middle');
  await page.evaluate(() => window.tabula.run('redo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 1, 0).valign), 'bottom');
  ok('일반 셀의 아래쪽 초기 선택·확인/취소·부모 가운데 재상속 방지 및 Undo/Redo');

  await page.evaluate(() => {
    const t = window.tabula, w = t.wb(), include = { number: true, alignment: true, font: true, border: true, fill: true, protection: true };
    w.load({ baseStyle: { valign: 'middle' }, cellStyles: [{ name: '아래쪽 기본', builtinId: 5, style: { numFmt: 'general', align: 'general', valign: 'bottom', bold: true }, include }], sheets: [{ name: '맞춤 검사', cells: {} }] });
    w.setInput(0, 0, 0, '12.5'); t.selectCell(0, 0);
  });
  await menu(); await page.locator('[data-cell-style="아래쪽 기본"]').click();
  const builtinBottomBeforeCancel = await state();
  await context('아래쪽 기본', '수정...'); await dialog('스타일 수정').getByRole('button', { name: '서식...', exact: true }).click();
  await dialog('셀 서식').getByRole('tab', { name: '맞춤', exact: true }).click();
  assert.equal(await dialog('셀 서식').getByLabel('세로', { exact: true }).inputValue(), '');
  await dialog('셀 서식').getByLabel('세로', { exact: true }).selectOption('top');
  await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click();
  await dialog('스타일 수정').getByRole('button', { name: '취소', exact: true }).click();
  assert.deepEqual(await state(), builtinBottomBeforeCancel);
  await context('아래쪽 기본', '수정...'); await dialog('스타일 수정').getByRole('button', { name: '서식...', exact: true }).click();
  await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click();
  await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
  let bottomDef = (await state()).styles.find(s => s.name === '아래쪽 기본');
  assert.equal(bottomDef.style.valign, 'bottom'); assert.equal(Object.hasOwn(bottomDef, 'customBuiltin'), false);
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'bottom');
  await context('아래쪽 기본', '수정...'); await dialog('스타일 수정').getByRole('button', { name: '서식...', exact: true }).click();
  await dialog('셀 서식').getByRole('tab', { name: '맞춤', exact: true }).click(); await dialog('셀 서식').getByLabel('세로', { exact: true }).selectOption('top');
  await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click();
  await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
  assert.equal((await state()).styles.find(s => s.name === '아래쪽 기본').customBuiltin, true);
  assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'top');
  await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'bottom');
  assert.equal(Object.hasOwn((await state()).styles.find(s => s.name === '아래쪽 기본'), 'customBuiltin'), false);
  await page.evaluate(() => window.tabula.run('redo')); assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), 'top');
  ok('이름 기본 스타일의 하위 서식 확인·상위 취소·명시 아래쪽 유지·실제 변경 표시 및 Undo/Redo');

  await page.evaluate(() => { const t = window.tabula, w = t.wb(); w.load({ baseStyle: { valign: 'bottom' }, sheets: [{ name: '표준 맞춤 검사', cells: {} }] }); w.setInput(0, 0, 0, '표준'); t.selectCell(0, 0); });
  const standardBeforeCancel = await state();
  await context('표준', '수정...'); await dialog('스타일 수정').getByRole('button', { name: '서식...', exact: true }).click();
  await dialog('셀 서식').getByRole('tab', { name: '맞춤', exact: true }).click();
  assert.equal(await dialog('셀 서식').getByLabel('세로', { exact: true }).inputValue(), '');
  await dialog('셀 서식').getByLabel('세로', { exact: true }).selectOption('top'); await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click();
  await dialog('스타일 수정').getByRole('button', { name: '취소', exact: true }).click(); assert.deepEqual(await state(), standardBeforeCancel);
  for (const [target, before] of [['', 'bottom'], ['middle', 'bottom'], ['', 'middle']]) {
    await context('표준', '수정...'); await dialog('스타일 수정').getByRole('button', { name: '서식...', exact: true }).click();
    await dialog('셀 서식').getByRole('tab', { name: '맞춤', exact: true }).click();
    assert.equal(await dialog('셀 서식').getByLabel('세로', { exact: true }).inputValue(), before === 'bottom' ? '' : before);
    await dialog('셀 서식').getByLabel('세로', { exact: true }).selectOption(target);
    await dialog('셀 서식').getByRole('button', { name: '확인', exact: true }).click(); await dialog('스타일 수정').getByRole('button', { name: '확인', exact: true }).click();
    assert.equal(await page.evaluate(() => window.tabula.wb().baseStyle.valign), target || 'bottom');
    assert.equal(await page.evaluate(() => window.tabula.wb().styleAt(0, 0, 0).valign), target || 'bottom');
    await page.evaluate(() => window.tabula.run('undo')); assert.equal(await page.evaluate(() => window.tabula.wb().baseStyle.valign), before);
    await page.evaluate(() => window.tabula.run('redo')); assert.equal(await page.evaluate(() => window.tabula.wb().baseStyle.valign), target || 'bottom');
  }
  ok('표준 스타일 아래쪽 확인/취소·가운데와 아래쪽 실제 변경 및 Undo/Redo');
  assert.deepEqual(errors, []); console.log(JSON.stringify({ groups, bad: 0, pageErrors: errors.length }));
} catch (err) {
  console.error(JSON.stringify({ errors, menus: await page.locator('#menuLayer').innerText(), dialogs: await page.locator('[role=dialog]').allTextContents() }));
  throw err;
} finally { await browser.close(); }
