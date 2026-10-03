// 합성 XLSX를 실제로 다시 열어 피벗 기본 스타일과 직접 서식의 우선순위를 검사한다.
import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
import { resolvePivot, computePivot, pivotSourceData } from '../src/pivot.js';
const engines = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const engine = process.env.WIXEL_BROWSER || 'chromium', url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const out = process.env.WIXEL_PIVOT_STYLE_OUT || 'D:/Codex/Temp/wixel-pivot-style-capture/source';
await mkdir(out, { recursive: true });
const elements = [{ type: 'wholeTable', style: { color: '#123456', fill: '#e0f2fe' } }, { type: 'headerRow', style: { fill: '#fed7aa', bold: true } }];
const wb = new Workbook({ sheets: [{ name: '자료', cells: { '0,0': { raw: '지역' }, '0,1': { raw: '매출' }, '1,0': { raw: '서울' }, '1,1': { raw: '123.45' }, '2,0': { raw: '부산' }, '2,1': { raw: '234.56' } } }, { name: '피벗', cells: {}, pivot: { name: '피벗1', source: '자료', range: { r1: 0, c1: 0, r2: 2, c2: 1 }, rows: ['지역'], cols: [], values: [{ field: '매출', agg: 'sum' }], top: 0, left: 0, style: '사용자', styleElements: elements } }], objectStyles: { tables: [{ name: '사용자', table: false, pivot: true, elements }], slicers: [] } });
const res = resolvePivot(pivotSourceData(wb, wb.sheets[1].pivot), wb.sheets[1].pivot), grid = computePivot(res, res.def).grid;
grid.forEach((row, r) => row.forEach((cd, c) => { if (cd) wb.setCellData(1, r, c, { raw: cd.raw, style: { ...cd.style, ...(r === 0 && c === 0 ? { fill: '#aa33bb', align: 'center' } : {}), ...(r > 0 && c === 1 ? { numFmt: 'custom', code: '0.000', align: 'right' } : {}) } }); }));
wb.sheets[1].pivot.area = { r1: 0, c1: 0, r2: grid.length - 1, c2: 1 };
const bytes = Buffer.from(writeXlsx(wb)), results = [], errors = [], writes = [];let checks = 0;
const eq = (a, b, name) => { checks++; assert.deepEqual(a, b, name); };
const browser = await engines[engine].launch(), context = await browser.newContext({ viewport: { width: 1366, height: 900 }, serviceWorkers: 'block' });
await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; localStorage.setItem('wixel.mobile-work.v1', 'off'); });
await context.route('**/*', route => { const r = route.request(), u = new URL(r.url()); if (!['GET', 'HEAD'].includes(r.method())) { writes.push(u.pathname); return route.abort(); } if (u.origin !== new URL(url).origin || u.pathname.startsWith('/api/')) return route.abort(); return route.continue(); });
const p = await context.newPage();p.setDefaultTimeout(15000);p.on('pageerror', e => errors.push(e.message));
const state = () => p.evaluate(() => { const w = tabula.wb(), s = w.sheets[1]; return { head: w.getCell(1, 0, 1).style, direct: w.getCell(1, 0, 0).style, number: w.getCell(1, 1, 1).style, raw: w.getRaw(1, 1, 1), style: s.pivot.style, cellFmt: s.pivot.cellFmt, source: w.serializeSheet(0).cells }; });
const color = c => p.locator('.c[data-r="0"][data-c="' + c + '"]:visible').first().evaluate(n => getComputedStyle(n.previousElementSibling?.classList.contains('cell-fill') ? n.previousElementSibling : n).backgroundColor);
try {
  await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });await p.waitForFunction(() => window.tabula?.wb());
  await p.locator('#fileInput').setInputFiles({ name: 'synthetic-pivot-capture.xlsx', mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: bytes });
  await p.waitForFunction(() => document.querySelector('#toast')?.textContent.includes('synthetic-pivot-capture') && document.querySelector('#toast')?.textContent.includes('열었습니다'));
  await p.evaluate(() => { tabula.switchSheet(1); tabula.selectCell(1, 1); });
  await p.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
  const before = await state();eq(await color(1), 'rgb(254, 215, 170)', '가져온 머리글 기본색');eq(await color(0), 'rgb(170, 51, 187)', '직접 지정한 머리글 색');eq(before.number.decimals, 3);eq(before.number.align, 'right');results.push('합성 XLSX 열기와 직접 서식');
  await p.evaluate(() => tabula.openNamedMenu('pivotStylesDesign', { x: 80, y: 90 }));await p.locator('#menuLayer [data-object-style="사용자"]').click({ button: 'right' });await p.getByRole('menuitem', { name: '수정...', exact: true }).click();
  const d = p.locator('.dialog').filter({ has: p.locator('.object-style-editor') });await d.locator('[data-element="headerRow"]').click();await d.getByRole('button', { name: '서식(F)...', exact: true }).click();const f = p.getByRole('dialog', { name: '셀 서식', exact: true });await f.getByRole('tab', { name: '채우기', exact: true }).click();await f.getByLabel('다른 색', { exact: true }).fill('#c0ffee');await f.getByRole('button', { name: '확인', exact: true }).click();await f.waitFor({ state: 'detached' });await d.getByRole('button', { name: '확인', exact: true }).click();await d.waitFor({ state: 'detached' });
  await p.waitForFunction(() => tabula.wb().getCell(1, 0, 1).style.fill === '#c0ffee');
  const changed = await state();eq(await color(1), 'rgb(192, 255, 238)', '수정한 머리글의 실제 픽셀 배경');eq(await color(0), 'rgb(170, 51, 187)', '사용자 직접 색 보존');eq(changed.number, before.number, '숫자·맞춤·문서 직접 서식 보존');eq(changed.source, before.source, '원본 자료 불변');await p.screenshot({ path: out + '/' + engine + '-modified.png' });results.push('가져온 스타일 편집과 직접 서식 보존');
  await p.evaluate(() => tabula.run('undo'));eq((await state()).head, before.head, 'Undo 이전 머리글');await p.evaluate(() => tabula.run('redo'));eq((await state()).head.fill, '#c0ffee', 'Redo 수정 머리글');
  await p.evaluate(() => tabula.openNamedMenu('pivotStylesDesign', { x: 80, y: 90 }));await p.getByRole('menuitem', { name: '지우기 (스타일 없음)', exact: true }).click();const cleared = await state();eq(cleared.style, 'None');eq(cleared.head.fill, undefined, '스타일 지우기는 옛 채움을 재적용하지 않음');eq(cleared.direct.fill, '#aa33bb', '지우기도 직접 채움 보존');eq(cleared.number.decimals, 3);eq(cleared.number.align, 'right');eq(cleared.source, before.source);results.push('Undo·Redo와 스타일 지우기');
  eq(errors, []);eq(writes, []);
} catch (e) { console.error(e.stack);results.push({ error: e.message });process.exitCode = 1; } finally { await browser.close(); }
const report = { url, engine, cases: 3, passed: results.filter(r => typeof r === 'string').length, checks, pageErrors: errors, remoteWrites: writes, results };await writeFile(out + '/' + engine + '-result.json', JSON.stringify(report, null, 2));console.log(JSON.stringify(report));
