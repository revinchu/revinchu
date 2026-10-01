// 합성 5만 항목: 서버 쓰기 금지, 목록 가상화·검색·키보드·필터 적용/Undo 회귀.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1440, height: 1000 } });
const page = await context.newPage(), errors = [], writes = [], times = [];
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
let checks = 0;
const check = (value, message) => { assert.ok(value, message); checks++; };
const equal = (value, expected, message) => { assert.deepEqual(value, expected, message); checks++; };
page.setDefaultTimeout(15000);
page.on('pageerror', error => errors.push(error.message));
await context.route('**/*', route => {
  if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort(); }
  // Optional local baseline source; never load a user's spreadsheet here.
  if (process.env.FILTER_CHECKLIST_BASELINE && new URL(route.request().url()).pathname === '/src/app-filter-checklist.js') return route.fulfill({ contentType: 'text/javascript', body: readFileSync(process.env.FILTER_CHECKLIST_BASELINE, 'utf8') });
  return route.continue();
});
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => !!window.tabula?.wb());
  await page.evaluate(() => {
    const n = 50000, t = window.tabula, w = t.wb();
    const dict = Array.from({ length: n }, (_, i) => `합성항목${String(i).padStart(5, '0')}`);
    w.restore({ sheets: [{ name: '필터 합성', cells: { '0,0': { raw: '항목' }, '0,1': { raw: '값' } },
      blocks: [{ r0: 1, c0: 0, n, ver: 0, cols: [
        { str: Int32Array.from({ length: n }, (_, i) => i), num: null, dict, fmt: null },
        { str: null, num: Float64Array.from({ length: n }, (_, i) => i), dict: [], fmt: null },
      ] }], filter: { r1: 0, r2: n, c1: 0, c2: 1, criteria: {}, hidden: {} } }] });
    w.undoStack = []; w.redoStack = []; t.selectCell(0, 0); t.gv().layout();
  });
  const open = async () => {
    const start = performance.now();
    await page.locator('.fbtn[data-c="0"]').first().click();
    await page.getByRole('searchbox', { name: '필터 항목 검색' }).waitFor();
    times.push(Math.round(performance.now() - start));
  };
  const all = () => page.getByRole('checkbox', { name: '표시된 항목 모두 선택', exact: true });
  const apply = async () => { await page.locator('.filter-foot').getByRole('button', { name: '확인', exact: true }).click(); await page.waitForFunction(() => !document.querySelector('.filter-menu')); };
  await open();
  const initialNodes = await page.locator('.filter-list input[type=checkbox]').count();
  if (process.env.FILTER_CHECKLIST_BASELINE) {
    console.log(JSON.stringify({ baseline: true, items: 50000, initialCheckboxNodes: initialNodes, openMs: times[0], pageErrors: errors.length, remoteWrites: writes.length }));
  } else {
    check(initialNodes <= 40, '5만 항목의 DOM은 뷰포트 크기로 제한');
    equal(await all().isChecked(), true, '초기 모든 항목 선택');
    await all().focus(); await page.keyboard.press('End');
    equal(await page.evaluate(() => document.activeElement?.getAttribute('aria-label')), '합성항목49999', 'End로 가상 목록 마지막 항목 이동');
    await page.keyboard.press('Space'); await page.keyboard.press('Home');
    equal(await page.evaluate(() => document.activeElement?.getAttribute('aria-label')), '표시된 항목 모두 선택', 'Home으로 모두 선택 복귀');
    equal(await all().evaluate(el => el.indeterminate), true, '화면 밖 마지막 항목 해제 보존');
    await apply();
    equal(await page.evaluate(() => window.tabula.wb().sheets[0].filter.criteria[0].length), 49999, '그려지지 않은 모든 선택값까지 적용');
    equal(await page.evaluate(() => { const h = window.tabula.wb().sheets[0].filter.hidden; return h.__bits ? h.count : Object.keys(h).length; }), 1, '마지막 한 행만 숨김');
    equal(await page.evaluate(() => window.tabula.wb().getValue(0, 50000, 1)), 49999, '숨긴 값 보존');
    await page.evaluate(() => window.tabula.run('undo'));
    equal(await page.evaluate(() => Object.keys(window.tabula.wb().sheets[0].filter.criteria).length), 0, 'Undo로 원래 필터 복원');
    await open(); await all().uncheck();
    const search = page.getByRole('searchbox', { name: '필터 항목 검색' });
    await search.fill('49999');
    equal(await page.getByRole('checkbox', { name: '합성항목49999', exact: true }).isChecked(), true, '검색된 끝 항목 선택');
    await search.fill('');
    equal(await all().isChecked(), false, '검색 지우면 검색 전 선택 복원');
    await search.fill('49999'); await apply();
    equal(await page.evaluate(() => window.tabula.wb().sheets[0].filter.criteria[0]), ['합성항목49999'], '검색 결과 하나만 적용');
    equal(await page.evaluate(() => { const h = window.tabula.wb().sheets[0].filter.hidden; return h.__bits ? h.count : Object.keys(h).length; }), 49999, '5만 행 중 검색 결과 한 행만 표시');
    await page.evaluate(() => window.tabula.run('undo'));
    await open();
    await search.fill('존재하지않는합성문자열');
    equal(await all().isDisabled(), true, '검색 결과 없을 때 모두 선택 비활성');
    check(await page.getByText('검색 결과가 없습니다.', { exact: true }).isVisible(), '빈 검색 결과 안내');
    await page.keyboard.press('Escape');
    await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), cells = { '0,0': { raw: '항목' }, '0,1': { raw: '값' }, '0,2': { raw: '색' } };
      for (let r = 1; r <= 4; r++) {
        cells[`${r},0`] = { raw: ['가', '나', '다', '라'][r - 1] };
        cells[`${r},1`] = { raw: String(r) };
        cells[`${r},2`] = { raw: '색', style: { fill: r === 3 ? '#ff0000' : '#0000ff' } };
      }
      w.restore({ sheets: [{ name: '조건 조합 합성', cells, filter: { r1: 0, r2: 4, c1: 0, c2: 2, hidden: {}, criteria: {
        0: { type: 'custom', op1: 'begins', v1: '라' }, 1: { type: 'custom', op1: 'gt', v1: '2' },
      } } }] });
      t.selectCell(0, 0); t.gv().layout();
    });
    await open();
    equal(await page.locator('.filter-window input').evaluateAll(nodes => nodes.map(node => node.getAttribute('aria-label'))), ['다', '라'], '다른 열 숫자 조건 반영·현재 열 조건 제외');
    await page.keyboard.press('Escape');
    await page.evaluate(() => { const t = window.tabula, w = t.wb(), f = w.sheets[0].filter; w.transact(() => w.setSheetProp(0, 'filter', { ...f, criteria: { ...f.criteria, 2: { type: 'fill', value: '#ff0000' } } })); });
    await open();
    equal(await page.locator('.filter-window input').evaluateAll(nodes => nodes.map(node => node.getAttribute('aria-label'))), ['다'], '다른 열 색상 조건과 숫자 조건 함께 반영');
    await page.keyboard.press('Escape');
    await page.evaluate(() => { const w = window.tabula.wb(), f = w.sheets[0].filter; w.transact(() => w.setSheetProp(0, 'filter', { ...f, criteria: { 1: { type: 'top', n: 1 } } })); });
    await open();
    equal(await page.locator('.filter-window input').evaluateAll(nodes => nodes.map(node => node.getAttribute('aria-label'))), ['라'], '다른 열 상위 조건 반영');
    await page.keyboard.press('Escape');
    equal(errors, [], '페이지 오류 없음'); equal(writes, [], '서버 쓰기 없음');
    console.log(JSON.stringify({ ok: true, checks, items: 50000, initialCheckboxNodes: initialNodes, openMs: times, pageErrors: errors.length, remoteWrites: writes.length }));
  }
} finally { await context.close(); await browser.close(); }
