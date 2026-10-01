// 실제 브라우저의 클래식 표 필드 클릭/HTML drag, 실행 취소, 보호/읽기 전용 회귀.
// 합성 문서만 사용. WIXEL_URL, PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH.
// CLASSIC_TOP=0: 원본을 별도 시트에 둔 A1 피벗. CLASSIC_SCREENSHOT: 화면/드래그 PNG 경로.
import assert from 'node:assert/strict';
import { gzipSync } from 'node:zlib';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';

const top = process.env.CLASSIC_TOP === '0' ? 0 : 6;
const sourceSi = top === 0 ? 1 : 0;
function fixture(classic = false) {
  const wb = new Workbook();
  if (sourceSi) wb.addSheet('원본');
  wb.transact(() => {
    [['지역', '품목', '매출', '비용', '시기', '채널'], ['서울', 'A', 10, 1, '1월', '온라인'], ['서울', 'B', 20, 2, '2월', '오프라인'], ['부산', 'A', 30, 3, '1월', '온라인']].forEach((row, r) => row.forEach((v, c) => wb.setInput(sourceSi, r, c, String(v))));
    wb.setSheetProp(0, 'pivot', {
      name: '클래식격자검증', source: wb.sheets[sourceSi].name, range: { r1: 0, c1: 0, r2: 3, c2: 5 },
      rows: ['지역', '품목'], cols: [], values: [{ field: '매출', name: '총매출', agg: 'sum' }, { field: '비용', name: '평균비용', agg: 'average', numFmt: { numFmt: 'number', decimals: 2 } }],
      layout: classic ? 'tabular' : 'compact', classic, showValuesRow: false, top, left: 0, area: { r1: top, c1: 0, r2: top + 9, c2: 3 },
      calcFields: [{ name: '이익', formula: '매출-비용' }],
    });
  });
  return wb;
}
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const errors = [], writes = [];
let checks = 0, drags = 0;
const eq = (a, b, label) => { assert.deepEqual(a, b, label); checks++; };
const bounded = async (promise, label) => {
  let timer;
  try { return await Promise.race([promise, new Promise((_, reject) => { timer = setTimeout(() => reject(new Error(label + ' (20초 초과)')), 20000); })]); }
  finally { clearTimeout(timer); }
};
const makePage = async () => {
  const context = await browser.newContext({ viewport: { width: 1800, height: 1080 } });
  const page = await context.newPage();
  page.setDefaultTimeout(10000);
  page.on('pageerror', (e) => errors.push(e.message));
  await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await context.route('**/*', (route) => {
    if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort(); }
    return route.continue();
  });
  return { context, page };
};
const { context, page } = await makePage();
const def = () => page.evaluate(() => {
  const d = window.tabula.wb().sheets[0].pivot;
  return { rows: d.rows, cols: d.cols, pages: d.pages ?? [], values: d.values, valuesOnRows: !!d.valuesOnRows, valuesPos: d.valuesPos ?? null, classic: d.classic, showValuesRow: d.showValuesRow, layout: d.layout };
});
const select = () => page.evaluate((top) => window.tabula.selectCell(top + 1, 0), top);
const undo = async (before, label) => { await page.evaluate(() => window.tabula.wb().undo()); eq(await def(), before, label); await select(); };
const drag = async (source, target) => {
  if (process.env.CLASSIC_DEBUG) console.error("drag-start", drags, await source.textContent());
  await source.scrollIntoViewIfNeeded();
  if (process.env.CLASSIC_DEBUG) console.error("drag-scrolled", drags);
  const from = await source.boundingBox(); assert.ok(from, 'drag source');
  const fromGrid = await source.evaluate((node) => node.classList.contains('pv-classic-field'));
  await page.mouse.move(from.x + from.width / 2, from.y + from.height / 2);
  await page.mouse.down();
  if (process.env.CLASSIC_DEBUG) console.error("drag-down", drags, from);
  await bounded(page.mouse.move(from.x + from.width / 2 + 12, from.y + from.height / 2 + 36, { steps: 1 }), '드래그 시작 mouse.move');
  if (process.env.CLASSIC_DEBUG) console.error("drag-initial-move", drags);
  await page.waitForFunction(() => document.body.classList.contains('pivot-grid-dragging'));
  const zonePosition = await page.locator('.pv-classic-zones').first().evaluate((node) => Number.isFinite(parseFloat(node.style.top)));
  eq(zonePosition, true, '드롭 띠의 top이 유효한 숫자');
  if (fromGrid) {
    const zoneBounds = await page.locator('.pv-classic-zones').first().boundingBox();
    eq(from.x < zoneBounds.x + zoneBounds.width && from.x + from.width > zoneBounds.x && from.y < zoneBounds.y + zoneBounds.height && from.y + from.height > zoneBounds.y, false, '드롭 띠가 드래그 시작 버튼을 덮지 않음');
  }
  if (process.env.CLASSIC_SCREENSHOT && drags === 5) await page.screenshot({ path: process.env.CLASSIC_SCREENSHOT.replace(/\.png$/i, '-drag.png') });
  if (process.env.CLASSIC_DEBUG) console.error("drag-body-class", drags);
  const to = await target.boundingBox(); assert.ok(to, 'drag target');
  if (process.env.CLASSIC_DEBUG) console.error("drag-target", drags, to);
  await bounded(page.mouse.move(to.x + to.width / 2, to.y + to.height / 2, { steps: 12 }), '드롭 대상 mouse.move');
  if (process.env.CLASSIC_DEBUG) console.error("drag-before-up", drags);
  await bounded(page.mouse.up(), '드롭 mouse.up');
  if (process.env.CLASSIC_DEBUG) console.error("drag-after-up", drags);
  await page.waitForFunction(() => !document.body.classList.contains('pivot-grid-dragging'));
  drags++;
};
const gridField = (name) => page.locator('.pv-classic-field').filter({ hasText: new RegExp('^' + name + '$') }).first();
const zone = (area) => page.locator(`.pv-classic-zone[data-area="${area}"]`).first();
const paneArea = (area) => page.locator(`#pivotPane [data-pivot-area="${area}"]`);
const paneField = (name) => page.locator('#pivotPane .pp-fields .pp-field').filter({ has: page.locator(`[data-pivot-field="${name}"]`) });
try {
  await page.goto(url);
  await page.waitForFunction(() => !!window.tabula?.wb());
  await page.locator('#fileInput').setInputFiles({ name: 'synthetic-classic-grid.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(writeXlsx(fixture())) });
  await page.waitForFunction(() => window.tabula.wb().sheets[0].pivot?.name === '클래식격자검증');
  await page.waitForFunction(() => !document.querySelector('.progress-overlay') && !window.tabula.wb().noUndo && window.tabula.wb().listeners.size > 0);
  await select();
  eq(await page.locator('.pv-classic-field').count(), 0, '클래식 해제 상태');
  await page.evaluate(() => window.tabula.run('pivotOptions'));
  await page.locator('.dialog').getByRole('button', { name: '표시', exact: true }).click();
  await page.getByLabel('클래식 피벗 테이블 레이아웃 (표 안으로 필드 끌어 놓기)', { exact: true }).check();
  await page.locator('.dialog').getByRole('button', { name: '확인', exact: true }).click();
  await page.waitForSelector('.pv-classic-field');
  eq((await def()).layout, 'tabular', '클래식 켜면 행 필드를 개별 열에 표시');
  eq(await page.locator('.pv-classic-field[data-sigma="false"]').allTextContents(), ['지역', '품목'], '한 행 필드당 버튼 하나');
  if (process.env.CLASSIC_SCREENSHOT) {
    await page.mouse.move(1700, 950);
    await page.screenshot({ path: process.env.CLASSIC_SCREENSHOT });
  }
  const baseline = await def();
  await gridField('품목').click();
  await page.getByRole('menuitem', { name: /열 영역으로 이동/ }).click();
  eq((await def()).rows, ['지역'], '클릭 메뉴가 기존 영역에서 제거');
  eq((await def()).cols, ['품목'], '클릭 메뉴가 열 영역에 추가');
  await undo(baseline, '클릭 메뉴 이동 Undo');
  await drag(gridField('품목'), zone('cols'));
  eq((await def()).cols, ['품목'], '격자 필드 → 격자 열 영역');
  await undo(baseline, '격자 드래그 Undo');
  await drag(paneField('시기'), zone('cols'));
  eq((await def()).cols, ['시기'], '필드 목록 → 격자 열 영역');
  const singleColumn = await def();
  await drag(paneField('채널'), zone('cols'));
  eq((await def()).cols, ['시기', '채널'], '여러 열 필드 배치');
  eq(await page.locator('.pv-classic-field[data-area="cols"][data-sigma="false"]').allTextContents(), ['시기', '채널'], '한 열 필드당 버튼 하나');
  await undo(singleColumn, '두 번째 열 필드 이동 Undo');
  await undo(baseline, '필드 목록 드래그 Undo');
  await drag(gridField('품목'), paneArea('pages'));
  eq((await def()).rows, ['지역'], '격자 → 필드 창에서 이전 영역 제거');
  eq((await def()).pages, ['품목'], '격자 → 필드 창의 필터 영역');
  await undo(baseline, '격자에서 필드 창 이동 Undo');
  await drag(paneArea('cols').locator('.pp-item.sigma'), zone('rows'));
  eq((await def()).valuesOnRows, true, 'Σ 값 → 행 축');
  eq((await def()).values, baseline.values, 'Σ 값 이동은 집계 설정 유지');
  await drag(paneArea('rows').locator('.pp-item.sigma'), zone('cols'));
  eq((await def()).valuesOnRows, false, 'Σ 값 → 열 축');
  await gridField('Σ 값').click();
  await page.getByRole('menuitem', { name: /행 영역으로 이동/ }).click();
  eq((await def()).valuesOnRows, true, '격자 Σ 값 클릭 → 행 축');
  eq(await page.locator('.pv-classic-field[data-sigma="true"][data-area="rows"]').count(), 1, '행 축에도 Σ 값 격자 버튼 유지');
  await drag(gridField('Σ 값'), zone('cols'));
  eq((await def()).valuesOnRows, false, '격자 Σ 값 드래그 → 열 축');
  await drag(gridField('Σ 값'), paneArea('rows'));
  eq((await def()).valuesOnRows, true, '격자 Σ 값 → 필드 창 행 축');
  await drag(gridField('Σ 값'), paneArea('cols'));
  eq((await def()).valuesOnRows, false, '격자 Σ 값 → 필드 창 열 축');
  const afterSigma = await def();
  await drag(paneArea('values').locator('.pp-item').filter({ hasText: '평균비용' }), zone('values'));
  eq((await def()).values, afterSigma.values, '동일 값 영역에 다시 놓아도 평균·숫자 형식 보존');
  const beforeValueMove = await def();
  await drag(paneArea('values').locator('.pp-item').filter({ hasText: '평균비용' }), zone('rows'));
  eq((await def()).values.map((v) => v.field), ['매출'], '값 필드를 행으로 옮길 때 해당 값 하나 제거');
  eq((await def()).rows, ['지역', '품목', '비용'], '숫자 필드의 행 배치');
  await undo(beforeValueMove, '값 필드 이동 Undo는 집계·형식 복원');
  await drag(paneField('이익'), zone('values'));
  eq((await def()).values.at(-1).agg, 'sum', '계산 필드는 값 영역에서 합계');
  await undo(beforeValueMove, '계산 필드 추가 Undo');
  await drag(paneField('이익'), zone('rows'));
  eq(await def(), beforeValueMove, '계산 필드를 행 영역에 놓는 것은 거절');

  // 끌기 도중 다른 작업으로 정의가 바뀌면 옛 드래그가 새 정의를 덮어쓰지 않는다.
  const staleTransfer = await page.evaluateHandle(() => new DataTransfer());
  await gridField('품목').dispatchEvent('dragstart', { dataTransfer: staleTransfer });
  await page.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'pivot', { ...w.sheets[0].pivot, showValuesRow: true })); });
  const staleBefore = await def();
  await zone('cols').dispatchEvent('drop', { dataTransfer: staleTransfer });
  await page.dispatchEvent('body', 'dragend', { dataTransfer: staleTransfer });
  eq(await def(), staleBefore, '정의가 바뀐 후 도착한 드래그 차단');
  await staleTransfer.dispose();
  await undo(beforeValueMove, '동시 변경 검사의 설정 원복');

  await page.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: {} })); });
  const protectedBefore = await def();
  await gridField('품목').click();
  const protectedMenu = page.getByRole('menuitem', { name: /열 영역으로 이동/ });
  if (await protectedMenu.count()) await protectedMenu.click();
  eq(await def(), protectedBefore, '보호된 시트 클릭 이동 차단');
  // 드래그 이벤트를 직접 보내도 보호를 우회하지 못하는지 확인한다.
  const transfer = await page.evaluateHandle(() => new DataTransfer());
  await paneField('채널').dispatchEvent('dragstart', { dataTransfer: transfer });
  await zone('cols').dispatchEvent('drop', { dataTransfer: transfer });
  await page.dispatchEvent('body', 'dragend', { dataTransfer: transfer });
  eq(await def(), protectedBefore, '보호된 시트 필드 창 → 격자 차단');
  await paneField('채널').dispatchEvent('dragstart', { dataTransfer: transfer });
  await paneArea('cols').dispatchEvent('drop', { dataTransfer: transfer });
  await page.dispatchEvent('body', 'dragend', { dataTransfer: transfer });
  eq(await def(), protectedBefore, '보호된 시트 필드 창 드롭 차단');
  await transfer.dispose();
  await page.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', null)); });
  const serialized = await page.evaluate(() => window.tabula.wb().serialize());
  const back = new Workbook(readXlsx(writeXlsx(new Workbook(serialized))).data).sheets[0].pivot;
  eq([back.classic, back.showValuesRow, back.layout], [true, false, 'tabular'], '이동 후 옵션 저장 왕복');
  eq(await page.evaluate((sourceSi) => window.tabula.wb().getValue(sourceSi, 1, 2), sourceSi), 10, '원본 데이터 보존');

  const readonly = await makePage();
  try {
    const hash = gzipSync(JSON.stringify({ docName: '읽기전용합성', workbook: fixture(true).serialize(), view: { headers: true, grid: true } })).toString('base64url');
    await readonly.page.goto(url + '#view=' + hash);
    await readonly.page.waitForSelector('body.view-mode');
    await readonly.page.waitForFunction(() => !!window.tabula?.wb().sheets[0].pivot);
    eq(await readonly.page.locator('.pv-classic-field').count(), 0, '공유 읽기 전용에는 이동 버튼 없음');
    eq(await readonly.page.locator('.pv-classic-zone').count(), 0, '공유 읽기 전용에는 드롭 영역 없음');
    await readonly.page.evaluate((top) => window.tabula.selectCell(top + 1, 0), top);
    const before = await readonly.page.evaluate(() => JSON.stringify(window.tabula.wb().sheets[0].pivot));
    const check = readonly.page.locator('#pivotPane [data-pivot-field="채널"]');
    if (await check.count()) await check.evaluate((el) => { el.checked = true; el.dispatchEvent(new Event('change', { bubbles: true })); });
    eq(await readonly.page.evaluate(() => JSON.stringify(window.tabula.wb().sheets[0].pivot)), before, '읽기 전용 필드 창 변경 이벤트도 차단');
  } finally { await readonly.context.close(); }
  eq(errors, [], '브라우저 오류'); eq(writes, [], '서버 쓰기');
  console.log(JSON.stringify({ ok: true, top, checks, nativeDrags: drags, pageErrors: errors, blockedWrites: writes }));
} catch (error) {
  console.error(JSON.stringify({ checks, nativeDrags: drags, dragClass: await page.evaluate(() => document.body.classList.contains('pivot-grid-dragging')), definition: await def(), pageErrors: errors }));
  throw error;
} finally { await context.close(); await browser.close(); }
