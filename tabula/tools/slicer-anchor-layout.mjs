// Native XLSX slicer anchoring and shared grid scale. Uses only an isolated browser.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { execFileSync } from 'node:child_process';
import { join, resolve, isAbsolute } from 'node:path';
import { pathToFileURL } from 'node:url';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';

const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
const trustedPublic = process.env.SLICER_ANCHOR_PUBLIC === '1' && new URL(url).origin === 'https://wixel-3.wizx.workers.dev';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname) && !trustedPublic) throw Error('로컬 서버 또는 명시한 위셀 공개 검증 주소에서만 실행하세요.');
const output = resolve(process.env.SLICER_ANCHOR_OUTPUT || (process.env.SLICER_ANCHOR_BASELINE ? '.local/slicer-anchor-layout/baseline' : '.local/slicer-anchor-layout'));
const repository = resolve('..');
const baseline = process.env.SLICER_ANCHOR_BASELINE ? execFileSync('git', ['-c', 'safe.directory=' + repository.replaceAll('\\\\', '/'), '-C', repository, 'show', process.env.SLICER_ANCHOR_BASELINE + ':tabula/src/app.js'], { encoding: 'utf8', maxBuffer: 8 * 1024 * 1024 }) : null;
await mkdir(output, { recursive: true });
const browserModule = process.env.PLAYWRIGHT_MODULE || 'playwright';
const { chromium } = await import(isAbsolute(browserModule) ? pathToFileURL(browserModule).href : browserModule);
const browser = await chromium.launch();
const results = [], metrics = [], errors = [], writes = [];
let checks = 0;
const close = (actual, expected, message, tolerance = 0.08) => {
  assert.ok(Math.abs(actual - expected) <= tolerance, message + ': ' + actual + ' ≈ ' + expected); checks++;
};
const equal = (actual, expected, message) => { assert.deepEqual(actual, expected, message); checks++; };
const ok = (value, message) => { assert.ok(value, message); checks++; };
const sourceRef = { kind: 'table', table: '배치표', column: '캠페인' };
const slicer = (id, caption, x, y, w, h, extra = {}) => ({
  id, name: caption, caption, x, y, w, h, source: sourceRef, columns: 1,
  style: 'SlicerStyleLight2', buttonHeight: 24.125, placement: 'oneCell', ...extra,
});
const cells = {
  '1,1': { raw: '캠페인' }, '1,2': { raw: '금액' },
  '2,1': { raw: '브랜드검색' }, '2,2': { raw: '10' },
  '3,1': { raw: '쇼핑검색' }, '3,2': { raw: '20' },
  '4,1': { raw: '파워링크' }, '4,2': { raw: '30' },
};
const sheet = {
  name: 'Excel 위치 검증', defRowH: 18, defColW: 80,
  colWidths: { 0: 30, 1: 160, 2: 150, 3: 60, 4: 60, 5: 160 },
  hiddenCols: { 8: true }, rowHeights: { 5: 208 }, rowManual: { 5: true },
  cells, tables: [{ id: 't', name: '배치표', r1: 1, c1: 1, r2: 4, c2: 2, header: true, columns: ['캠페인', '금액'] }],
  slicers: [
    slicer('campaign', '캠페인', 354.375, 20.625, 338.125, 145.25, { columns: 2 }),
    slicer('week', '주차', 710.625, 20.625, 430.25, 140.125, { columns: 3 }),
    slicer('month', '년월', 30.125, 104.375, 118.125, 108.375),
    slicer('type', '캠페인유형', 166.625, 104.375, 166.75, 116.125),
  ],
};
for (let r = 14; r < 18; r++) for (let c = 1; c < 8; c++) {
  sheet.cells[r + ',' + c] = { raw: r === 14 ? ['주차', '총비용', '노출수', '클릭수', '전환수', '전환매출', 'ROAS'][c - 1] : String((r - 13) * (c + 1)),
    style: r === 14 ? { fill: '#2f5597', color: '#ffffff', bold: true } : { numFmt: 'number' } };
}
const placements = ['oneCell', 'twoCell', 'absolute'];
const fitSheet = {
  name: '자동 높이 검증', defRowH: 18, defColW: 80,
  cells: { '0,0': { raw: '첫째 줄\n둘째 줄', style: { wrap: true, size: 22 } } },
  slicers: placements.map((placement, i) => slicer('fit-' + placement, placement, 200.375 + i * 230, 70.625, 180.25, 110.5, { placement })),
};
const fixture = join(output, 'slicer-anchors-native.xlsx');
const bytes = writeXlsx(new Workbook({ defaultFont: { name: '맑은 고딕', size: 11 }, sheets: [sheet, fitSheet] }));
await writeFile(fixture, bytes);
const imported = readXlsx(bytes).data;
const expectedMain = imported.sheets[0].slicers.map(o => ({ id: o.id, caption: o.caption, x: o.x, y: o.y, w: o.w, h: o.h }));
const expectedFit = imported.sheets[1].slicers.map(o => ({ id: o.id, placement: o.placement ?? 'oneCell', x: o.x, y: o.y, w: o.w, h: o.h }));
ok(imported.fitRows?.[1]?.includes(0), '자동 행 높이 후보가 표준 XLSX에서 읽힘');
const context = await browser.newContext({ viewport: { width: 1660, height: 1060 }, deviceScaleFactor: 1 });
const page = await context.newPage();
page.setDefaultTimeout(15000);
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => {
  const request = route.request(), target = new URL(request.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method()); return route.abort(); }
  if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) return route.abort();
  return route.continue();
});
if (baseline) await context.route('**/src/app.js', route => route.fulfill({ contentType: 'text/javascript', body: baseline }));
await context.addInitScript(() => {
  window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true;
  Object.defineProperty(navigator, 'clipboard', { value: { writeText: async text => { window.__anchorClipboardText = text; }, readText: async () => window.__anchorClipboardText || '' }, configurable: true });
});
async function settle() {
  await page.evaluate(() => document.fonts.ready);
  await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
}
async function openFixture() {
  await page.evaluate(() => { window.__anchorOldBook = tabula.wb(); });
  await page.locator('#fileInput').setInputFiles(fixture);
  await page.waitForFunction(() => window.tabula?.wb() !== window.__anchorOldBook && window.tabula?.wb().sheets[0]?.name === 'Excel 위치 검증' && !document.querySelector('.load-progress'), null, { timeout: 60000 });
  await settle();
}
async function zoom(percent) {
  await page.locator('#zoomSlider').evaluate((input, value) => { input.value = String(value); input.dispatchEvent(new Event('input', { bubbles: true })); }, percent);
  await settle();
}
const model = () => page.evaluate(() => {
  const t = tabula, w = t.wb(), s = w.sheets[t.si];
  return { slicers: s.slicers.map(o => ({ id: o.id, placement: o.placement ?? 'oneCell', caption: o.caption, x: o.x, y: o.y, w: o.w, h: o.h })),
    row0: w.rowHeight(t.si, 0), row5: w.rowHeight(t.si, 5), col0: w.colWidth(t.si, 0), undo: w.undoStack.length };
});
async function snapshot(label) {
  const result = await page.evaluate(label => {
    const t = tabula, g = t.gv(), s = t.wb().sheets[t.si], panes = [];
    const info = g.sheetRect({ r1: 1, c1: 1, r2: 4, c2: 2 });
    const report = g.sheetRect({ r1: 14, c1: 1, r2: 17, c2: 7 });
    for (const pane of g.panes) {
      if (pane.el.style.display === 'none') continue;
      const origin = pane.content.getBoundingClientRect();
      const normalize = node => {
        const rect = node.getBoundingClientRect();
        return { x: (rect.left - origin.left) / g.z + pane.ox, y: (rect.top - origin.top) / g.z + pane.oy,
          w: rect.width / g.z, h: rect.height / g.z, screen: { left: rect.left, top: rect.top, right: rect.right, bottom: rect.bottom } };
      };
      const objects = [...pane.objects.querySelectorAll('.obj.slicer:not(.slicer-group)')].map(node => {
        const o = s.slicers.find(o => o.id === node.dataset.id);
        return { id: o.id, caption: o.caption, actual: normalize(node), expected: { x: o.x, y: o.y, w: o.w, h: o.h } };
      });
      const tableCells = [...pane.cells.querySelectorAll('.c')].filter(node => +node.dataset.r >= 1 && +node.dataset.r <= 4 && +node.dataset.c >= 1 && +node.dataset.c <= 2)
        .map(node => ({ r: +node.dataset.r, c: +node.dataset.c, actual: normalize(node), expected: g.sheetRect({ r1: +node.dataset.r, c1: +node.dataset.c, r2: +node.dataset.r, c2: +node.dataset.c }) }));
      panes.push({ id: pane.id, objects, tableCells });
    }
    return { label, zoom: g.z * 100, viewport: { w: innerWidth, h: innerHeight }, scroll: { x: g.sx, y: g.sy },
      freeze: s.freeze, hiddenColumnWidth: g.cols.size(8), tableEdges: { info, report }, panes };
  }, label);
  ok(result.panes.some(p => p.objects.length), label + ': 슬라이서 DOM이 실제 표시됨');
  for (const pane of result.panes) {
    for (const item of [...pane.objects, ...pane.tableCells]) for (const key of ['x', 'y', 'w', 'h']) close(item.actual[key], item.expected[key], label + ' / ' + pane.id + ' / ' + (item.caption || item.r + ',' + item.c) + ' / ' + key);
    for (const item of pane.objects) for (const [name, table] of Object.entries(result.tableEdges)) {
      const a = item.actual;
      ok(a.x >= table.x + table.w - 0.08 || a.x + a.w <= table.x + 0.08 || a.y >= table.y + table.h - 0.08 || a.y + a.h <= table.y + 0.08, label + ': ' + item.caption + '가 ' + name + ' 표를 침범하지 않음');
    }
  }
  equal(result.hiddenColumnWidth, 0, '숨긴 열은 표와 슬라이서의 공통 축에서 0px');
  metrics.push(result);
  return result;
}
async function test(name, fn) {
  try { await fn(); results.push({ name, ok: true }); console.log('OK ' + name); }
  catch (error) { results.push({ name, ok: false, error: error.message }); console.error('NG ' + name + ': ' + error.stack); }
}
async function dimension(kind, index, value) {
  await page.evaluate(({ kind, index }) => { tabula.selectCell(kind === 'row' ? index : 0, kind === 'col' ? index : 0); tabula.run(kind === 'row' ? 'rowHeight' : 'colWidth'); }, { kind, index });
  const dialog = page.locator('.dialog').last();
  await dialog.locator('input[type=number]').fill(String(value));
  await dialog.getByRole('button', { name: '확인', exact: true }).click();
  await settle();
}
try {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
  await openFixture();
  await page.evaluate(() => { window.__anchorImportedSnapshot = tabula.wb().serialize(); });
  await test('표준 XLSX 슬라이서 원래 좌표·소수 크기와 큰 빈 행 보존', async () => {
    const current = await model();
    for (const expected of expectedMain) {
      const actual = current.slicers.find(o => o.caption === expected.caption);
      for (const key of ['x', 'y', 'w', 'h']) close(actual[key], expected[key], actual.caption + ' 원래 ' + key, 1 / 9525 + 1e-7);
    }
    equal(current.row5, 208, '수동 빈 행 높이 보존');
  });
  await test('25·55·70·100·150·200·400%에서 표·슬라이서 공통 배율', async () => {
    const before = await model();
    for (const percent of [25, 55, 70, 100, 150, 200, 400]) {
      await zoom(percent);
      await page.evaluate(() => tabula.gv().setScroll(0, 0)); await settle();
      await snapshot('zoom-' + percent);
      if ([55, 100].includes(percent)) await page.screenshot({ path: join(output, 'slicer-layout-' + percent + '.png') });
    }
    equal((await model()).slicers, before.slicers, '확대/축소가 저장 좌표와 크기를 바꾸지 않음');
  });
  await test('뷰포트 변경과 분수 스크롤에도 상대 위치 보존', async () => {
    await zoom(70);
    for (const viewport of [{ width: 1280, height: 900 }, { width: 1960, height: 1180 }]) {
      await page.setViewportSize(viewport); await settle();
      await page.evaluate(() => { tabula.gv().setScroll(45.375, 31.625); tabula.gv().renderAll(); }); await settle();
      await snapshot('viewport-' + viewport.width);
    }
  });
  await test('행·열 틀 고정의 네 패널도 같은 앵커를 표시', async () => {
    await zoom(70);
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'freeze', { rows: 6, cols: 3 })); t.gv().layout(); t.gv().setScroll(41.375, 22.625); t.gv().renderAll(); });
    await settle(); const frozen = await snapshot('freeze');
    ok(frozen.panes.length >= 2, '틀 고정 패널이 둘 이상 생성됨');
    await page.screenshot({ path: join(output, 'slicer-layout-freeze.png') });
  });
  await test('파일 열기 중 자동 행 높이는 셀을 따라 이동하고 절대 배치를 보존', async () => {
    await page.evaluate(() => tabula.switchSheet(1)); await zoom(100);
    const current = await model(), delta = current.row0 - 18;
    ok(delta > 10, '22pt 여러 줄의 자동 행 높이가 실제로 커짐');
    for (const expected of expectedFit) {
      const actual = current.slicers.find(o => o.placement === expected.placement);
      close(actual.y, expected.y + (expected.placement === 'absolute' ? 0 : delta), expected.placement + ' 열기 후 y', 1 / 9525 + 1e-7);
      for (const key of ['x', 'w', 'h']) close(actual[key], expected[key], expected.placement + ' 열기 후 ' + key, 1 / 9525 + 1e-7);
    }
  });
  await test('수동 행 높이: 이동만·이동과 크기·절대 배치 및 Undo', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); t.gv().layout(); t.gv().renderAll(); }); await zoom(100);
    await page.evaluate(() => { tabula.wb().undoStack = []; tabula.wb().redoStack = []; });
    const before = await model();
    await dimension('row', 5, 25.5); // 34px, inside all three slicers.
    const after = await model(), delta = after.row5 - before.row5;
    close(delta, 16, '선택 행이 18px에서 34px로 변경');
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id);
      for (const key of ['x', 'y', 'w']) close(current[key], original[key], original.placement + ' 내부 행 변경 ' + key, 1e-7);
      close(current.h, original.h + (original.placement === 'twoCell' ? delta : 0), original.placement + ' 내부 행 변경 h', 1e-7);
    }
    equal(after.undo, 1, '행 높이와 개체 변화는 Undo 한 번');
    await page.evaluate(() => tabula.run('undo')); equal((await model()).slicers, before.slicers, 'Undo는 소수 좌표까지 복원');
  });
  await test('수동 열 너비의 반복 변경·행 삽입은 소수 좌표를 누적 반올림하지 않음', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); t.gv().layout(); t.gv().renderAll(); }); await zoom(100);
    const before = await model();
    for (let i = 0; i < 4; i++) {
      await dimension('col', 0, 7.25);
      const wider = await model(), delta = wider.col0 - before.col0;
      for (const original of before.slicers) {
        const current = wider.slicers.find(o => o.id === original.id);
        close(current.x, original.x + (original.placement === 'absolute' ? 0 : delta), original.placement + ' 열 너비 이동', 1e-7);
        for (const key of ['y', 'w', 'h']) close(current[key], original[key], original.placement + ' 열 너비 ' + key, 1e-7);
      }
      await page.evaluate(() => tabula.run('undo')); equal((await model()).slicers, before.slicers, '반복 Undo 소수 좌표');
    }
    await page.evaluate(() => { tabula.selectRange({ r1: 2, c1: 0, r2: 2, c2: 0 }, 'rows'); tabula.run('insertRows'); });
    const after = await model();
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id);
      close(current.y, original.y + (original.placement === 'absolute' ? 0 : 18), original.placement + ' 행 삽입 이동', 1e-7);
      for (const key of ['x', 'w', 'h']) close(current[key], original[key], original.placement + ' 행 삽입 ' + key, 1e-7);
    }
    await page.evaluate(() => tabula.run('undo')); equal((await model()).slicers, before.slicers, '행 삽입 Undo 소수 좌표');
  });
  await test('F2 여러 줄 편집의 자동 행 높이도 셀 앵커를 따라 이동하고 Undo로 복원', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; t.selectCell(1, 0); }); await zoom(100);
    const before = await model(), beforeHeight = await page.evaluate(() => tabula.wb().rowHeight(1, 1));
    await page.locator('#cellEditor').focus(); await page.keyboard.press('F2');
    await page.locator('#cellEditor').fill('첫째 줄\n둘째 줄\n셋째 줄'); await page.keyboard.press('Enter'); await settle();
    const after = await model(), delta = await page.evaluate(() => tabula.wb().rowHeight(1, 1)) - beforeHeight;
    ok(delta > 15, '여러 줄 편집으로 자동 행 높이가 실제 증가');
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id);
      close(current.y, original.y + (original.placement === 'absolute' ? 0 : delta), original.placement + ' 자동 높이 편집 y', 1e-7);
      for (const key of ['x', 'w', 'h']) close(current[key], original[key], original.placement + ' 자동 높이 편집 ' + key, 1e-7);
    }
    equal(after.undo, 1, '셀 편집과 행 높이·개체 배치를 Undo 한 번으로 복원');
    await page.evaluate(() => tabula.run('undo')); equal((await model()).slicers, before.slicers, '편집 Undo는 소수 좌표까지 복원');
    equal(await page.evaluate(() => tabula.wb().rowHeight(1, 1)), beforeHeight, '편집 Undo는 자동 행 높이 복원');
  });
  await test('F2 숫자 입력의 자동 열 확장도 배치 속성을 지키고 Undo로 복원', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; t.selectCell(1, 0); }); await zoom(100);
    const before = await model(), beforeHeight = await page.evaluate(() => tabula.wb().rowHeight(1, 1));
    equal(await page.evaluate(() => tabula.wb().sheets[1].colWidths[0]), undefined, '입력 열은 수동 너비 없는 기본 열');
    await page.locator('#cellEditor').focus(); await page.keyboard.press('F2');
    await page.locator('#cellEditor').fill('123456789012.3456'); await page.keyboard.press('Enter'); await settle();
    const after = await model(), dx = after.col0 - before.col0, dy = await page.evaluate(() => tabula.wb().rowHeight(1, 1)) - beforeHeight;
    ok(dx > 15, '큰 숫자 입력으로 기본 열 너비가 실제 증가');
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id), anchored = original.placement !== 'absolute';
      close(current.x, original.x + (anchored ? dx : 0), original.placement + ' 자동 열 확장 x', 1e-7);
      close(current.y, original.y + (anchored ? dy : 0), original.placement + ' 자동 열 확장 후 행 맞춤 y', 1e-7);
      for (const key of ['w', 'h']) close(current[key], original[key], original.placement + ' 자동 열 확장 ' + key, 1e-7);
    }
    equal(after.undo, 1, '숫자 입력과 열 너비·행 높이·개체 배치는 Undo 한 번');
    await page.evaluate(() => tabula.run('undo'));
    equal((await model()).slicers, before.slicers, '숫자 입력 Undo는 소수 좌표까지 복원');
    equal((await model()).col0, before.col0, '숫자 입력 Undo는 원래 열 너비 복원');
    equal(await page.evaluate(() => tabula.wb().sheets[1].colWidths[0]), undefined, '숫자 입력 Undo는 자동 열 상태 복원');
    equal(await page.evaluate(() => tabula.wb().rowHeight(1, 1)), beforeHeight, '숫자 입력 Undo는 행 높이 복원');
    equal(await page.evaluate(() => tabula.wb().getRaw(1, 1, 0)), '', '숫자 입력 Undo는 빈 셀 복원');
  });
  await test('선택하여 붙여넣기 열 너비는 슬라이서 배치와 Undo를 함께 유지', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); w.transact(() => { w.setColWidth(1, 8, 128); w.setInput(1, 1, 8, '복사 원본'); }); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; }); await zoom(100);
    const before = await model();
    await page.evaluate(() => { tabula.selectCell(1, 8); tabula.run('copy'); tabula.selectCell(1, 0); tabula.run('pasteSpecial'); });
    const dialog = page.getByRole('dialog', { name: '선택하여 붙여넣기', exact: true });
    await dialog.getByRole('radio', { name: '열 너비', exact: true }).check();
    await dialog.getByRole('button', { name: '확인', exact: true }).click(); await settle();
    const after = await model(), dx = after.col0 - before.col0;
    close(after.col0, 128, '복사한 원본의 128px 열 너비 적용', 1e-7);
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id);
      close(current.x, original.x + (original.placement === 'absolute' ? 0 : dx), original.placement + ' 열 너비 붙여넣기 x', 1e-7);
      for (const key of ['y', 'w', 'h']) close(current[key], original[key], original.placement + ' 열 너비 붙여넣기 ' + key, 1e-7);
    }
    equal(await page.evaluate(() => tabula.wb().getRaw(1, 1, 0)), '', '열 너비 붙여넣기는 셀 값을 덮지 않음');
    equal(after.undo, 1, '열 너비 붙여넣기와 배치를 Undo 한 번으로 복원');
    await page.evaluate(() => tabula.run('undo'));
    equal((await model()).slicers, before.slicers, '붙여넣기 Undo 소수 좌표 복원');
    equal((await model()).col0, before.col0, '붙여넣기 Undo 원래 열 너비');
  });
  await test('필터 머리글 자동 열 확장은 배치 속성을 지키고 Undo로 복원', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); w.transact(() => { w.setInput(1, 1, 0, '슬라이서 앵커 검증 필터 머리글'); w.setInput(1, 2, 0, 'A'); w.setInput(1, 3, 0, 'B'); }); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; t.selectRange({ r1: 1, c1: 0, r2: 3, c2: 0 }); }); await zoom(100);
    const before = await model();
    await page.evaluate(() => tabula.run('toggleFilter')); await settle();
    const after = await model(), dx = after.col0 - before.col0;
    ok(dx > 15, '긴 머리글과 필터 단추 때문에 열 너비가 실제 증가');
    ok(await page.evaluate(() => !!tabula.wb().sheets[1].filter), '워크시트 필터가 실제 생성됨');
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id);
      close(current.x, original.x + (original.placement === 'absolute' ? 0 : dx), original.placement + ' 필터 자동폭 x', 1e-7);
      for (const key of ['y', 'w', 'h']) close(current[key], original[key], original.placement + ' 필터 자동폭 ' + key, 1e-7);
    }
    equal(after.undo, 1, '필터 생성과 열 너비·배치를 Undo 한 번으로 복원');
    await page.evaluate(() => tabula.run('undo'));
    equal((await model()).slicers, before.slicers, '필터 Undo 소수 좌표 복원');
    equal((await model()).col0, before.col0, '필터 Undo 원래 열 너비');
    equal(await page.evaluate(() => tabula.wb().sheets[1].filter), null, '필터 Undo는 필터 해제');
  });
  await test('서식 메뉴 기본 너비는 기본 열의 앵커와 크기를 바꾸고 Undo로 복원', async () => {
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.restore(window.__anchorImportedSnapshot); t.switchSheet(1); t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = []; t.selectCell(1, 0); }); await zoom(100);
    const before = await model();
    await page.evaluate(() => tabula.openNamedMenu('format', document.querySelector('#gridView')));
    await page.getByRole('menuitem', { name: /^기본 너비/ }).click();
    const dialog = page.getByRole('dialog', { name: '기본 너비', exact: true });
    await dialog.locator('input[type=number]').fill('96');
    await dialog.getByRole('button', { name: '확인', exact: true }).click(); await settle();
    const after = await model();
    close(after.col0, 96, '기본 열 너비 80px에서 96px로 변경', 1e-7);
    for (const original of before.slicers) {
      const current = after.slicers.find(o => o.id === original.id), anchored = original.placement !== 'absolute';
      const start = Math.floor(original.x / 80), end = Math.floor((original.x + original.w) / 80);
      close(current.x, original.x + (anchored ? start * 16 : 0), original.placement + ' 기본폭 x', 1e-7);
      close(current.w, original.w + (original.placement === 'twoCell' ? (end - start) * 16 : 0), original.placement + ' 기본폭 w', 1e-7);
      for (const key of ['y', 'h']) close(current[key], original[key], original.placement + ' 기본폭 ' + key, 1e-7);
    }
    equal(after.undo, 1, '기본 너비와 배치는 Undo 한 번으로 복원');
    await page.evaluate(() => tabula.run('undo'));
    equal((await model()).slicers, before.slicers, '기본폭 Undo 소수 좌표 복원');
    equal((await model()).col0, before.col0, '기본폭 Undo 원래 기본 너비');
    equal(await page.evaluate(() => tabula.wb().sheets[1].colWidths[0]), undefined, '기본폭 Undo는 명시적 너비를 추가하지 않음');
  });
  equal(errors, [], '페이지 오류'); equal(writes, [], '외부 쓰기 없음');
} finally {
  await writeFile(join(output, 'slicer-anchor-metrics.json'), JSON.stringify({ url, checks, results, metrics, pageErrors: errors, blockedWrites: writes }, null, 2));
  await browser.close();
}
console.log(JSON.stringify({ checks, total: results.length, passed: results.filter(r => r.ok).length, failed: results.filter(r => !r.ok).length, output }));
assert.equal(results.every(r => r.ok), true, '슬라이서 앵커 화면 회귀 검사');
