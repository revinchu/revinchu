// 실제 사용자 파일·서버 저장과 분리한 그림 서식 UI 회귀. source/cloud 공통.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = [], writes = [], results = []; let checks = 0;
page.setDefaultTimeout(10000); page.on('pageerror', e => errors.push(e.message));
await page.route('**/*', route => { if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().method()); return route.abort('blockedbyclient'); } const target = new URL(route.request().url()); return target.origin === new URL(url).origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort(); });
const eq = (a, b, message) => { assert.deepEqual(a, b, message); checks++; };
const ev = (fn, arg) => page.evaluate(fn, arg);
const current = () => ev(() => window.tabula.wb().sheets[0].images[0]);
const fixture = async (patch = {}) => {
  for (let i = 0; i < 3; i++) await page.keyboard.press('Escape');
  await ev(patch => {
    const t = window.tabula, w = t.wb(); t.switchSheet(0);
    const canvas = document.createElement('canvas'); canvas.width = 240; canvas.height = 120;
    const c = canvas.getContext('2d'); c.fillStyle = '#ff0000'; c.fillRect(0, 0, 120, 120); c.fillStyle = '#0000ff'; c.fillRect(120, 0, 120, 120);
    w.restore({ sheets: [{ name: '합성 그림', cells: {}, images: [{ id: 'picture-test', name: '빨강 파랑', alt: '색상 비교', src: canvas.toDataURL(), x: 90, y: 70, w: 240, h: 120, ...patch }] }] });
    t.gv().setZoom(100); t.gv().renderAll(); t.selectCell(0, 0); w.undoStack = []; w.redoStack = [];
  }, patch);
};
const open = async () => { await page.locator('.obj.pic[data-id="picture-test"]').first().click(); await page.keyboard.press('F2'); const dlg = page.getByRole('dialog', { name: '그림 서식', exact: true }); await dlg.waitFor(); return dlg; };
const reveal = async target => {
  if (!await target.isVisible()) {
    const id = await target.evaluate(node => node.closest('[role="tabpanel"]')?.id);
    if (id) await page.locator(`[role="tab"][aria-controls="${id}"]`).click();
    const closed = await target.evaluate(node => { const labels = []; for (let n = node.parentElement; n; n = n.parentElement) if (n.tagName === 'DETAILS' && !n.open) labels.push(n.querySelector('summary')?.textContent); return labels.reverse(); });
    for (const label of closed) if (label) await page.locator('.picture-format-dialog summary').filter({ hasText: label }).first().click();
  }
  await target.scrollIntoViewIfNeeded(); return target;
};
const input = label => reveal(page.getByRole('dialog', { name: '그림 서식', exact: true }).getByLabel(label, { exact: true }));
const button = label => reveal(page.getByRole('dialog', { name: '그림 서식', exact: true }).locator('button').filter({ hasText: new RegExp('^' + label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:\\s*\\([A-Z0-9]\\))?$') }));
const test = async (name, fn) => { if (process.env.WIXEL_PICTURE_FILTER && !name.includes(process.env.WIXEL_PICTURE_FILTER)) return; try { await fn(); results.push({ name, ok: true }); console.log('OK ' + name); } catch (e) { results.push({ name, ok: false }); console.error('NG ' + name + ': ' + e.stack); } };
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(url, { timeout: 60000 }); await page.waitForFunction(() => !!window.tabula?.wb());
  await test('미리보기·비율 연동·Esc 취소는 통합 문서와 Undo에 영향 없음', async () => {
    await fixture(); const before = await current(); const dlg = await open();
    await (await input('너비(px)')).fill('360'); eq(await (await input('높이(px)')).inputValue(), '180');
    await (await input('왼쪽 자르기(%)')).fill('25'); await (await input('회전(°)')).fill('30'); await (await input('좌우 대칭')).check();
    eq(await current(), before); eq(await ev(() => window.tabula.wb().undoStack.length), 0);
    const preview = await page.locator('.picture-preview-frame').evaluate(el => ({ transform: el.style.transform, crop: el.querySelector('img').style.width }));
    eq(preview.transform.includes('rotate(30deg)') && /scale\(-1,\s*1\)/.test(preview.transform), true); eq(preview.crop, '133.333%');
    await page.screenshot({ path: process.env.WIXEL_PICTURE_SCREENSHOT || 'D:/Codex/Temp/wixel3-picture-format.png' });
    await page.keyboard.press('Escape'); eq(await dlg.count(), 0); eq(await current(), before);
  });
  await test('자르기·회전·대칭·설명·배치 확인은 1회 Undo 및 XLSX 왕복', async () => {
    await fixture(); const before = await current(); const dlg = await open();
    await (await input('가로 세로 비율 고정')).uncheck(); await (await input('너비(px)')).fill('300'); await (await input('높이(px)')).fill('180');
    await (await input('회전(°)')).fill('22.5'); await (await input('좌우 대칭')).check(); await (await input('상하 대칭')).check();
    await (await input('왼쪽 자르기(%)')).fill('10'); await (await input('오른쪽 자르기(%)')).fill('20');
    await (await input('그림 설명')).fill('빨강 & 파랑 비교'); await (await input('셀 크기가 바뀔 때')).selectOption('absolute');
    await (await button('확인')).click();
    const after = await current(); eq([after.w, after.h, after.lockAspect, after.rot, after.flip, after.flipV, after.alt, after.placement], [300, 180, false, 22.5, true, true, '빨강 & 파랑 비교', 'absolute']);
    eq(after.crop, { l: .1, r: .2 }); eq(await ev(() => window.tabula.wb().undoStack.length), 1);
    const dom = await page.locator('.obj.pic[data-id="picture-test"]').first().evaluate(el => ({ transform: el.style.transform, alt: el.querySelector('img').alt, width: el.querySelector('img').style.width }));
    eq(dom, { transform: 'rotate(22.5deg) scale(-1, -1)', alt: after.alt, width: '142.857%' });
    const serialized = await ev(() => window.tabula.wb().serialize()), back = readXlsx(writeXlsx(new Workbook(serialized))).data.sheets[0].images[0];
    for (const k of ['w', 'h', 'lockAspect', 'rot', 'flip', 'flipV', 'alt', 'placement', 'crop']) eq(back[k], after[k], k);
    await ev(() => window.tabula.run('undo')); eq(await current(), before); await ev(() => window.tabula.run('redo')); eq(await current(), after);
  });
  await test('원본 크기·서식 재설정·자르기 초기화도 확인 전에는 원본 보존', async () => {
    await fixture({ w: 480, h: 240, rot: 45, flip: true, crop: { l: .25 }, border: '#000000', radius: 10, shadow: true });
    const before = await current(); const dlg = await open();
    await (await button('원본 크기')).click(); eq(await (await input('너비(px)')).inputValue(), '240'); eq(await (await input('높이(px)')).inputValue(), '120');
    await (await button('자르기 초기화')).click(); eq(await (await input('왼쪽 자르기(%)')).inputValue(), '0');
    await (await button('그림 서식 원래대로')).click(); eq(await (await input('회전(°)')).inputValue(), '0'); eq(await (await input('좌우 대칭')).isChecked(), false);
    eq(await current(), before); await (await button('확인')).click();
    const after = await current(); eq([after.w, after.h, after.rot, after.crop, after.border, after.radius, after.shadow], [240, 120, undefined, undefined, undefined, undefined, undefined]); eq(after.alt, before.alt);
  });
  await test('자르기 합계100% 방지 및 Enter 확정', async () => {
    await fixture({ crop: { r: .6 } }); await open(); await (await input('왼쪽 자르기(%)')).fill('90'); eq(await (await input('왼쪽 자르기(%)')).inputValue(), '39');
    await (await input('이름')).fill('Enter 확정'); await (await input('이름')).press('Enter'); eq((await current()).name, 'Enter 확정'); eq((await current()).crop.l, .39);
  });
  await test('모서리 끌기 잠금/해제와 Shift 반전', async () => {
    for (const [lockAspect, shift, expected] of [[true, false, 2], [false, false, null], [false, true, 2]]) {
      await fixture({ lockAspect }); await page.locator('.obj.pic[data-id="picture-test"]').first().click();
      const h = await page.locator('.obj.pic.sel .ch-h.se').boundingBox(); await page.mouse.move(h.x + h.width / 2, h.y + h.height / 2); if (shift) await page.keyboard.down('Shift');
      await page.mouse.down(); await page.mouse.move(h.x + h.width / 2 + 80, h.y + h.height / 2 + 10, { steps: 5 }); await page.mouse.up(); if (shift) await page.keyboard.up('Shift');
      const p = await current(); if (expected) eq(Math.abs(p.w / p.h - expected) < .01, true, JSON.stringify({lockAspect,shift,w:p.w,h:p.h})); else eq(Math.abs(p.w / p.h - 2) > .1, true, JSON.stringify({lockAspect,shift,w:p.w,h:p.h}));
    }
  });
  await test('투명도·모서리·그림자 수치 조절과 저장·재열기·Undo', async () => {
    await page.setViewportSize({ width: 1440, height: 1000 }); await fixture(); const before = await current(); const dlg = await open();
    await (await input('그림 투명도(%)')).fill('65'); await (await input('둥근 모서리(px)')).fill('8');
    await (await input('그림자 표시')).check(); await (await input('그림자 가로 거리(px)')).fill('-4'); await (await input('그림자 세로 거리(px)')).fill('5'); await (await input('그림자 흐리게(px)')).fill('6'); await (await input('그림자 투명도(%)')).fill('75');
    await (await input('그림자 표시')).uncheck(); await (await input('그림자 표시')).check();
    eq(await (await input('그림자 가로 거리(px)')).inputValue(), '-4');
    const preview = await page.locator('.picture-preview-frame').evaluate(el => ({ opacity: el.querySelector('.picture-visual').style.opacity, radius: el.querySelector('.picture-pixels > div').style.borderRadius, shadow: [...el.querySelectorAll('feDropShadow')].map(n => [n.getAttribute('dx'), n.getAttribute('dy'), n.getAttribute('stdDeviation')]) }));
    eq(preview.opacity, '0.35'); eq(preview.radius, '8px'); eq(preview.shadow.some(s => s[0] === '-4' && s[1] === '5' && s[2] === '3'), true);
    await (await button('확인')).click(); const after = await current();
    eq(after.opacity, .35); eq(after.shadow.dx, -4); eq(after.radius, 8);
    const serialized = await ev(() => window.tabula.wb().serialize()), back = readXlsx(writeXlsx(new Workbook(serialized))).data.sheets[0].images[0];
    eq(back.opacity, .35); eq(Math.abs(back.radius - 8) < .001, true); eq(back.shadow.dx, -4);
    await ev(() => window.tabula.run('undo')); eq(await current(), before); await ev(() => window.tabula.run('redo')); eq(await current(), after);
  });
  await test('작은 화면에서 대화상자 가로 넘침과 키보드 초점 이탈 없음', async () => {
    await page.setViewportSize({ width: 620, height: 800 }); await fixture();
    await page.locator('.obj.pic[data-id="picture-test"]').first().click({ position: { x: 35, y: 35 } });
    await page.locator('#ribbon [data-ribbon-control="pictureFormat:command:pictureSize"]').click();
    const dlg = page.getByRole('dialog', { name: '그림 서식', exact: true }); await dlg.waitFor();
    const bounds = await dlg.evaluate(el => ({ left: el.getBoundingClientRect().left, right: el.getBoundingClientRect().right, width: innerWidth, over: el.scrollWidth > el.clientWidth + 2 }));
    eq(bounds.left >= 0 && bounds.right <= bounds.width, true); eq(bounds.over, false);
    await (await button('취소')).focus(); await page.keyboard.press('Tab'); eq(await dlg.evaluate(el => el.contains(document.activeElement)), true);
    await page.keyboard.press('Escape'); eq(await dlg.count(), 0);
  });
  eq(errors, [], 'pageErrors'); eq(writes, [], 'server writes');
  console.log(JSON.stringify({ cases: results.length, checks, bad: results.filter(r => !r.ok).length, results, pageErrors: errors, blockedWrites: writes }));
  if (results.some(r => !r.ok)) process.exitCode = 1;
} finally { await browser.close(); }
