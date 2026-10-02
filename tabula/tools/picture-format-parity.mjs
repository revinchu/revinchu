// 그림 서식 parity: 합성 PNG만 사용, 격리 컨텍스트, API·외부 요청·원격 쓰기 차단.
// WIXEL_URL=소스/번들 주소, WIXEL_PICTURE_PARITY_FILTER=검사 이름 일부.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip } from '../src/zip.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
const browser = await chromium.launch(), results = [];
let checks = 0, ribbonCoverage = null;
const eq = (actual, expected, message) => { assert.deepEqual(actual, expected, message); checks++; };
const ok = (value, message) => { assert.ok(value, message); checks++; };
const dialog = p => p.getByRole('dialog', { name: '그림 서식', exact: true });
const current = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets[0].images[0]));
const history = p => p.evaluate(() => window.tabula.wb().undoStack.length);
const run = (p, command, arg) => p.evaluate(([command, arg]) => window.tabula.run(command, arg), [command, arg]);
async function reveal(p, target) {
  if (!await target.isVisible()) {
    const id = await target.evaluate(node => node.closest('[role="tabpanel"]')?.id);
    if (id) await dialog(p).locator(`[role="tab"][aria-controls="${id}"]`).click();
    const closed = await target.evaluate(node => [...function* () { for (let n = node.parentElement; n; n = n.parentElement) if (n.tagName === 'DETAILS' && !n.open) yield n; }()].map(n => n.querySelector('summary')?.textContent));
    for (const label of closed.reverse()) if (label) await dialog(p).locator('summary').filter({ hasText: label }).first().click();
  }
  await target.scrollIntoViewIfNeeded(); return target;
}
const field = async (p, label) => reveal(p, dialog(p).getByLabel(label, { exact: true }));
const button = async (p, label) => reveal(p, dialog(p).locator('button').filter({ hasText: new RegExp('^' + label.replace(/[.*+?^${}()|[\]\\]/g, '\\$&') + '(?:\\s*\\([A-Z0-9]\\))?$') }));
async function fill(p, label, value) { await (await field(p, label)).fill(String(value)); }
async function select(p) { await p.locator('.obj.pic[data-id="parity-picture"]').first().click({ position: { x: 35, y: 35 } }); }
async function open(p, command = 'pictureFormat') { await select(p); await run(p, command); await dialog(p).waitFor(); }
const confirm = async p => { await dialog(p).getByRole('button', { name: '확인', exact: true }).click(); await dialog(p).waitFor({ state: 'detached' }); };
async function savePicker(p) {
  await p.evaluate(() => {
    window.__pictureFiles = [];
    window.showSaveFilePicker = async ({ suggestedName }) => ({ kind: 'file', name: suggestedName, createWritable: async () => ({
      write: async blob => window.__pictureFiles.push({ name: suggestedName, bytes: Array.from(new Uint8Array(await blob.arrayBuffer())) }), close: async () => {}, abort: async () => {},
    }) });
  });
}
async function pixels(p, src, points) {
  return p.evaluate(async ({ src, points }) => {
    const image = new Image(); image.src = src; await image.decode();
    const c = document.createElement('canvas'); c.width = image.naturalWidth; c.height = image.naturalHeight;
    const ctx = c.getContext('2d'); ctx.drawImage(image, 0, 0);
    return { w: c.width, h: c.height, values: points.map(([x, y]) => Array.from(ctx.getImageData(Math.floor(x * c.width), Math.floor(y * c.height), 1, 1).data)) };
  }, { src, points });
}
async function memoryClipboard(p) {
  await p.evaluate(() => {
    window.__pictureClipboard = { writes: [], deny: false };
    Object.defineProperty(navigator, 'clipboard', { configurable: true, value: {
      write: async items => { if (window.__pictureClipboard.deny) throw new DOMException('합성 거절', 'NotAllowedError'); const blob = await items[0].getType('image/png'); window.__pictureClipboard.writes.push({ type: blob.type, bytes: Array.from(new Uint8Array(await blob.arrayBuffer())) }); },
      read: async () => [], readText: async () => '', writeText: async () => {},
    } });
  });
}
async function fixture(p, patch = {}) {
  await p.evaluate(patch => {
    const canvas = document.createElement('canvas'); canvas.width = 640; canvas.height = 320;
    const c = canvas.getContext('2d'); c.fillStyle = '#ffffff'; c.fillRect(0, 0, 640, 320);
    c.fillStyle = '#dc2626'; c.fillRect(100, 60, 180, 200); c.fillStyle = '#2563eb'; c.fillRect(330, 60, 180, 200);
    c.fillStyle = '#111827'; c.font = '32px sans-serif'; c.fillText('SYNTHETIC', 210, 42);
    const t = window.tabula, w = t.wb();
    w.restore({ sheets: [{ name: '그림 합성', cells: { '0,0': { raw: '보존' }, '0,1': { raw: '=1+1', cached: 2 } }, images: [{ id: 'parity-picture', name: '합성 PNG', alt: '흰 배경 위 빨강과 파랑', src: canvas.toDataURL(), x: 90, y: 65, w: 320, h: 160, ...patch }] }, { name: '다른 합성 시트', cells: {} }] });
    t.switchSheet(0); t.gv().setZoom(100); t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
    w.undoStack = []; w.redoStack = [];
  }, patch);
}
async function test(name, fn) {
  if (process.env.WIXEL_PICTURE_PARITY_FILTER && !name.includes(process.env.WIXEL_PICTURE_PARITY_FILTER)) return;
  const context = await browser.newContext({ viewport: { width: 1550, height: 1050 }, acceptDownloads: true });
  const p = await context.newPage(), errors = [], writes = [], blocked = []; p.setDefaultTimeout(15000);
  p.on('pageerror', e => errors.push(e.message)); p.on('dialog', d => d.dismiss());
  await context.route('**/*', route => {
    const request = route.request(), target = new URL(request.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method()); return route.abort(); }
    if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) { blocked.push(target.pathname); return route.abort(); }
    return route.continue();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    await fixture(p); await fn(p);
    eq(await p.evaluate(() => [window.tabula.wb().getRaw(0, 0, 0), window.tabula.wb().getRaw(0, 0, 1)]), ['보존', '=1+1'], '그림 편집이 셀 내용을 변경하지 않음');
    eq(errors, [], '페이지 오류'); eq(writes, [], '원격 쓰기');
    results.push({ name, ok: true, blockedRequests: blocked.length }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors, writes }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
try {
  await test('그림 선택 전용 리본·전체 조작 키팁·명령 등록', async p => {
    await select(p);
    eq(await p.locator('[data-ribbon-tab="pictureFormat"].active').count(), 1, '그림 선택 시 전용 탭 자동 활성화');
    eq(await p.locator('[data-ribbon-tab="objFormat"]').count(), 0, '그림에 도형용 탭을 잘못 노출하지 않음');
    const r = await p.evaluate(() => window.tabula.keytipRegistry()), controls = r.controls.filter(c => c.tabId === 'pictureFormat');
    ok(controls.length >= 12, '그림 서식 리본 조작이 등록됨');
    eq(r.tabs.jq, 'pictureFormat');
    const ids = controls.map(c => c.id), actual = await p.locator('#ribbon [data-ribbon-control]').evaluateAll(nodes => nodes.flatMap(n => (n.dataset.ribbonControls || n.dataset.ribbonControl).split(' ')));
    eq([...new Set(actual)].sort(), [...new Set(ids)].sort(), '등록표와 실제 리본 조작의 일치');
    const commands = new Set(await p.evaluate(() => window.tabula.commands()));
    eq(controls.filter(c => c.kind === 'command' && !commands.has(c.target)), [], '명령 실존');
    eq(controls.filter(c => !r.entries.some(e => e.controlId === c.id)), [], '전체 조작 접근키');
    ribbonCoverage = { controls: controls.length, keytips: r.entries.filter(e => e.tabId === 'pictureFormat').length };
    await p.keyboard.press('Alt'); await p.keyboard.press('j'); await p.keyboard.press('q');
    eq(await p.locator('[data-ribbon-tab="pictureFormat"].active').count(), 1); await p.keyboard.press('Escape'); await p.keyboard.press('Escape');
  });

  await test('우클릭 그림 서식 진입·모달 편집 취소는 원본과 Undo 불변', async p => {
    const before = await current(p); await p.locator('.obj.pic[data-id="parity-picture"]').first().click({ button: 'right', position: { x: 35, y: 35 } });
    await p.locator('.menu-item').filter({ hasText: '그림 서식(O)' }).click(); await dialog(p).waitFor();
    await fill(p, '너비(px)', 480); eq(await (await field(p, '높이(px)')).inputValue(), '240');
    await fill(p, '왼쪽 자르기(%)', 10); await fill(p, '회전(°)', 30);
    eq(await current(p), before); eq(await history(p), 0);
    await dialog(p).getByRole('button', { name: '취소', exact: true }).click(); eq(await current(p), before); eq(await history(p), 0);
  });

  await test('크기·자르기·대칭·설명 한 번 확정과 Undo Redo·XLSX 재열기', async p => {
    const before = await current(p); await open(p, 'pictureCrop');
    await fill(p, '왼쪽 자르기(%)', 12); await fill(p, '오른쪽 자르기(%)', 18);
    await (await field(p, '가로 세로 비율 고정')).uncheck(); await fill(p, '너비(px)', 350); await fill(p, '높이(px)', 200);
    await fill(p, '회전(°)', 20); await (await field(p, '좌우 대칭')).check(); await fill(p, '그림 설명', '새 합성 설명');
    await confirm(p); const after = await current(p);
    eq([after.w, after.h, after.rot, after.flip, after.lockAspect, after.alt], [350, 200, 20, true, false, '새 합성 설명']); eq(after.crop, { l: .12, r: .18 }); eq(after.src, before.src); eq(await history(p), 1);
    const back = readXlsx(writeXlsx(new Workbook(await p.evaluate(() => window.tabula.wb().serialize())))).data.sheets[0].images[0];
    for (const key of ['w', 'h', 'rot', 'flip', 'lockAspect', 'alt', 'crop']) eq(back[key], after[key], key);
    await run(p, 'undo'); eq(await current(p), before); await run(p, 'redo'); eq(await current(p), after);
  });

  await test('리본 cm 크기 입력은 비율을 지키고 다른 개체·셀을 변경하지 않음', async p => {
    const before = await current(p); await select(p);
    const size = p.locator('#ribbon [data-ribbon-command="pictureW"] input,#ribbon input[data-ribbon-command="pictureW"],#ribbon [data-ribbon-control*="pictureW"] input,#ribbon input[data-ribbon-control*="pictureW"]').first();
    await size.fill('10.16'); await size.press('Enter'); const after = await current(p);
    ok(Math.abs(after.w - 384) < .1); ok(Math.abs(after.h - 192) < .1); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before);
  });

  await test('열린 그림 서식에서 보호 또는 최종본으로 전환하면 확정 거부', async p => {
    for (const marked of [false, true]) {
      await fixture(p); const before = await current(p); await open(p); await fill(p, '너비(px)', 400);
      await p.evaluate(marked => { const w = window.tabula.wb(); if (marked) w.props = { ...w.props, markedFinal: true }; else w.sheets[0].protect = { on: true, allow: { selectLocked: true, selectUnlocked: true, objects: false } }; }, marked);
      await dialog(p).getByRole('button', { name: '확인', exact: true }).click(); eq(await current(p), before); eq(await history(p), 0);
      for (let i = 0; i < 3 && await p.locator('#dialogLayer .dialog').count(); i++) await p.keyboard.press('Escape');
      eq(await p.locator('#dialogLayer .dialog').count(), 0, '보호 안내와 미확정 서식 창을 모두 닫음');
    }
  });

  await test('개체 편집 허용된 보호 시트에서 그림 서식 정상 적용', async p => {
    await p.evaluate(() => { window.tabula.wb().sheets[0].protect = { on: true, allow: { selectLocked: true, selectUnlocked: true, objects: true } }; });
    await open(p); await fill(p, '그림 설명', '개체 허용'); await confirm(p); eq((await current(p)).alt, '개체 허용'); eq(await history(p), 1);
  });

  await test('최종본·보호 시트의 그림 키보드 삭제·잘라내기·복제도 변경 거부', async p => {
    await memoryClipboard(p);
    for (const [mode, key] of [['final', 'Delete'], ['final', 'Control+x'], ['final', 'Control+d'], ['protected', 'Control+d']]) {
      await fixture(p); const before = await current(p); await select(p);
      await p.evaluate(mode => { const w = window.tabula.wb(); if (mode === 'final') w.props = { ...w.props, markedFinal: true }; else w.sheets[0].protect = { on: true, allow: { selectLocked: true, selectUnlocked: true, objects: false } }; }, mode);
      await p.locator('#cellEditor').focus(); await p.keyboard.press(key);
      eq(await p.evaluate(() => window.tabula.wb().sheets[0].images.length), 1, mode + ' ' + key); eq(await current(p), before); eq(await history(p), 0);
      for (let i = 0; i < 3 && await p.locator('#dialogLayer .dialog').count(); i++) await p.keyboard.press('Escape');
    }
  });

  await test('리본 그림 스타일 실제 적용·Undo 및 원래대로', async p => {
    const before = await current(p); await select(p);
    await p.locator('#ribbon .picture-style-swatch[aria-label="흰색 단순 프레임"]').click();
    const after = await current(p); eq(after.border, '#ffffff'); eq(after.borderW, 6); ok(after.shadow.blur > 0); eq(after.src, before.src); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before); await run(p, 'redo'); eq(await current(p), after);
    await run(p, 'pictureReset'); eq((await current(p)).border, undefined); eq((await current(p)).shadow, undefined); eq((await current(p)).src, before.src);
  });

  await test('각 리본 편집 명령은 해당 범주를 열고 방향키로 범주 탐색', async p => {
    for (const [command, section] of [['pictureCorrections', 'corrections'], ['pictureColor', 'color'], ['pictureArtistic', 'artistic'], ['pictureTransparency', 'effects'], ['pictureBackground', 'background'], ['pictureCompress', 'compress'], ['pictureBorder', 'effects'], ['pictureEffects', 'effects'], ['pictureCrop', 'crop'], ['pictureAlt', 'alt']]) {
      await open(p, command); eq(await dialog(p).locator('[role="tab"][aria-selected="true"]').getAttribute('data-section'), section, command);
      await dialog(p).getByRole('button', { name: '취소', exact: true }).click();
    }
    await open(p); const selected = dialog(p).locator('[role="tab"][aria-selected="true"]'); await selected.focus(); await p.keyboard.press('ArrowRight');
    eq(await selected.evaluate(n => n === document.activeElement), true); eq(await history(p), 0); await p.keyboard.press('Escape');
  });

  await test('보정·색·꾸밈·테두리·네온·반사 실제 초안과 한 번의 Undo', async p => {
    const before = await current(p); await open(p, 'pictureCorrections');
    await fill(p, '밝기(%)', 15); await fill(p, '대비(%)', -20); await fill(p, '선명하게(%)', 30);
    await fill(p, '채도(%)', 75); await (await field(p, '다시 칠하기')).selectOption('sepia');
    await (await field(p, '꾸밈 효과 종류')).selectOption('posterize'); await fill(p, '꾸밈 효과 강도(%)', 60);
    await (await field(p, '그림 테두리')).check(); await fill(p, '그림 테두리 두께(px)', 4); await (await field(p, '그림 테두리 선 종류')).selectOption('dashDot');
    await (await field(p, '네온 표시')).check(); await fill(p, '네온 크기(px)', 12); await fill(p, '부드러운 가장자리(px)', 3);
    await (await field(p, '반사 표시')).check(); await fill(p, '반사 크기(%)', 25);
    eq(await current(p), before); eq(await history(p), 0);
    const preview = await dialog(p).locator('.picture-preview-frame').evaluate(n => ({ filters: n.querySelectorAll('filter').length, reflected: !!n.querySelector('.picture-reflection'), dash: n.querySelector('.picture-border rect')?.getAttribute('stroke-dasharray') }));
    ok(preview.filters > 0); eq(preview.reflected, true); eq(preview.dash, '16 8 4 8');
    await p.screenshot({ path: process.env.WIXEL_PICTURE_PARITY_SCREENSHOT || 'D:/Codex/Temp/wixel-picture-format-parity.png' });
    await confirm(p); const after = await current(p);
    eq(after.correction, { brightness: .15, contrast: -.2, sharpness: .3 }); eq(after.color.saturation, .75); eq(after.color.recolor, 'sepia'); eq(after.artistic, { type: 'posterize', amount: .6 });
    eq(after.borderDash, 'dashDot'); eq(after.glow.size, 12); eq(after.softEdge, 3); eq(after.reflection.size, .25); eq(after.src, before.src); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before); await run(p, 'redo'); eq(await current(p), after);
  });

  await test('자르기 비율 프리셋은 원본을 유지하고 정사각형·중앙 자르기를 적용', async p => {
    const before = await current(p); await open(p, 'pictureCrop');
    await (await field(p, '자르기 가로 세로 비율')).selectOption('square'); await confirm(p); const after = await current(p);
    eq([after.w, after.h], [320, 320]); eq(after.crop, { l: .25, r: .25, t: 0, b: 0 }); eq(after.src, before.src); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before);
  });

  await test('자르기 8개 조절점 실제 드래그·키보드·취소는 초안에만 적용', async p => {
    const before = await current(p); await open(p, 'pictureCrop');
    eq(await dialog(p).locator('.picture-v2-crop-handle:visible').count(), 8);
    const handle = dialog(p).getByRole('button', { name: '오른쪽 자르기 조절점', exact: true }), h = await handle.boundingBox(), frame = await dialog(p).locator('.picture-preview-frame').boundingBox();
    await p.mouse.move(h.x + h.width / 2, h.y + h.height / 2); await p.mouse.down(); await p.mouse.move(h.x + h.width / 2 - frame.width * .1, h.y + h.height / 2, { steps: 4 }); await p.mouse.up();
    const crop = Number(await (await field(p, '오른쪽 자르기(%)')).inputValue()); ok(Math.abs(crop - 10) < .2, `드래그 비율 ${crop}`);
    await handle.focus(); await p.keyboard.press('ArrowLeft'); await p.keyboard.press('Shift+ArrowLeft');
    const keyboardCrop = Number(await (await field(p, '오른쪽 자르기(%)')).inputValue()); ok(Math.abs(keyboardCrop - 16) < .2, `원본 기준 1%+5%: ${keyboardCrop}`);
    await (await field(p, '오른쪽 자르기(%)')).focus(); const again = await handle.boundingBox();
    await p.mouse.move(again.x + again.width / 2, again.y + again.height / 2); await p.mouse.down(); await p.mouse.move(again.x + again.width / 2 - 18, again.y + again.height / 2, { steps: 3 });
    await p.keyboard.press('Escape'); await p.mouse.up(); eq(await dialog(p).count(), 1, '드래그 Escape는 창을 닫지 않음');
    eq(Number(await (await field(p, '오른쪽 자르기(%)')).inputValue()), keyboardCrop, '드래그 시작 전 자르기 초안 복원');
    eq(await current(p), before); eq(await history(p), 0);
    await dialog(p).getByRole('button', { name: '취소', exact: true }).click(); eq(await current(p), before); eq(await history(p), 0);
    await open(p, 'pictureCrop'); await dialog(p).getByRole('button', { name: '왼쪽 자르기 조절점', exact: true }).focus(); await p.keyboard.press('ArrowRight'); await confirm(p);
    eq((await current(p)).crop.l, .01); eq(await history(p), 1); await run(p, 'undo'); eq(await current(p), before);
  });

  await test('그림 복사는 메모리 클립보드 PNG와 내부 편집 사본을 보존하고 권한 거절도 안전', async p => {
    await fixture(p, { color: { recolor: 'grayscale' }, crop: { l: .1 }, border: '#008800' }); const before = await current(p);
    await memoryClipboard(p);
    const contextMenu = async () => p.locator('.obj.pic[data-id="parity-picture"]').first().click({ button: 'right', position: { x: 35, y: 35 } });
    await contextMenu(); await p.locator('.menu-item').filter({ hasText: /^복사\(C\)/ }).click(); await p.waitForFunction(() => window.__pictureClipboard.writes.length === 1);
    const copy = await p.evaluate(() => window.__pictureClipboard.writes[0]); eq(copy.type, 'image/png'); const image = await pixels(p, 'data:image/png;base64,' + Buffer.from(copy.bytes).toString('base64'), [[.25, .5]]);
    const [r, g, b] = image.values[0]; ok(Math.abs(r - g) <= 1 && Math.abs(g - b) <= 1); eq(await current(p), before); eq(await history(p), 0);
    await contextMenu(); await p.locator('.menu-item').filter({ hasText: /^붙여넣기\(P\)/ }).click();
    const pasted = await p.evaluate(() => structuredClone(window.tabula.wb().sheets[0].images[1]));
    for (const key of ['src', 'crop', 'color', 'border', 'w', 'h']) eq(pasted[key], before[key]); eq([pasted.x, pasted.y], [before.x + 12, before.y + 12]); eq(await history(p), 1);
    await run(p, 'undo'); eq(await p.evaluate(() => window.tabula.wb().sheets[0].images.length), 1);
    await p.evaluate(() => { window.__pictureClipboard.deny = true; }); await contextMenu(); await p.locator('.menu-item').filter({ hasText: /^복사\(C\)/ }).click();
    await p.waitForFunction(() => document.querySelector('#toast')?.textContent.includes('위셀 안에 복사'));
    await contextMenu(); await p.locator('.menu-item').filter({ hasText: /^붙여넣기\(P\)/ }).click(); eq(await p.evaluate(() => window.tabula.wb().sheets[0].images.length), 2); eq((await current(p)).src, before.src);
  });

  await test('그림 바꾸기 파일 선택은 크기·서식을 보존하고 새 원본 복원·Undo와 오래된 요청 차단', async p => {
    await fixture(p, { crop: { l: .2 }, border: '#008800', correction: { brightness: .15 } }); const before = await current(p);
    const src = await p.evaluate(() => { const c = document.createElement('canvas'); c.width = 120; c.height = 240; const ctx = c.getContext('2d'); ctx.fillStyle = '#f97316'; ctx.fillRect(0, 0, 120, 240); return c.toDataURL(); });
    const file = { name: 'synthetic-replacement.png', mimeType: 'image/png', buffer: Buffer.from(src.split(',')[1], 'base64') };
    const choose = async () => {
      await select(p); await p.locator('#ribbon [data-ribbon-control="pictureFormat:menu:pictureChange"]').click();
      const promise = p.waitForEvent('filechooser'); await p.locator('.menu-item').filter({ hasText: '이 장치에서(D)' }).click(); return promise;
    };
    await (await choose()).setFiles(file); await p.waitForFunction(src => window.tabula.wb().sheets[0].images[0].src === src, src);
    const after = await current(p); eq(after.crop, undefined); eq(after.originalSrc, src); eq([after.originalWidth, after.originalHeight], [120, 240]);
    for (const key of ['x', 'y', 'w', 'h', 'border', 'correction']) eq(after[key], before[key]); eq(await history(p), 1);
    await run(p, 'pictureResetSize'); eq([(await current(p)).w, (await current(p)).h, (await current(p)).src], [120, 240, src]);
    await run(p, 'undo'); eq(await current(p), after); await run(p, 'undo'); eq(await current(p), before);
    const pending = await choose(); await p.evaluate(() => window.tabula.switchSheet(1)); await pending.setFiles(file);
    await p.waitForFunction(() => document.querySelector('#toast')?.textContent.includes('문서 또는 그림이 변경'));
    eq(await current(p), before); eq(await history(p), 0); eq(await p.evaluate(() => window.tabula.wb().sheets[1].images.length), 0);
  });

  await test('셀에 배치는 회전·보정을 보존하고 처리 중 시트 전환은 다른 셀을 덮지 않음', async p => {
    await fixture(p, { rot: 90, color: { recolor: 'grayscale' } }); const before = await current(p);
    const rc = await p.evaluate(() => { const t = window.tabula, im = t.wb().sheets[0].images[0]; return [t.gv().rows.indexAt(im.y + 1), t.gv().cols.indexAt(im.x + 1)]; });
    const move = async () => { await p.locator('.obj.pic[data-id="parity-picture"]').first().click({ button: 'right' }); await p.locator('.menu-item').filter({ hasText: /^셀에 배치$/ }).click(); };
    await move(); await p.waitForFunction(() => window.tabula.wb().sheets[0].images.length === 0);
    const cellImage = await p.evaluate(([r, c]) => window.tabula.wb().getCell(0, r, c)?.image, rc), raster = await pixels(p, cellImage.src, [[.5, .25], [.5, .65]]);
    eq([raster.w, raster.h], [320, 640]); for (const [r, g, b] of raster.values) ok(Math.abs(r - g) <= 1 && Math.abs(g - b) <= 1); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before); eq(await p.evaluate(([r, c]) => window.tabula.wb().getCell(0, r, c)?.image, rc), undefined);
    await p.evaluate(() => {
      const NativeImage = window.Image; window.__releasePictureLoad = []; window.__nativePictureImage = NativeImage;
      window.Image = function (w, h) { const im = new NativeImage(w, h); let callback; Object.defineProperty(im, 'onload', { configurable: true, get: () => callback, set: fn => { callback = fn; } }); im.addEventListener('load', e => window.__releasePictureLoad.push(() => callback?.call(im, e))); return im; };
    });
    await move(); await p.waitForFunction(() => window.__releasePictureLoad.length > 0);
    await p.evaluate(() => { window.tabula.switchSheet(1); window.Image = window.__nativePictureImage; for (const fn of window.__releasePictureLoad.splice(0)) fn(); });
    await p.waitForFunction(() => document.querySelector('#toast')?.textContent.includes('문서 또는 그림이 변경'));
    eq(await current(p), before); eq(await history(p), 0); eq(await p.evaluate(([r, c]) => window.tabula.wb().getCell(1, r, c)?.image, rc), undefined);
  });

  await test('로컬 압축 미리보기·적용·취소·확정과 보관한 원본 복원', async p => {
    const before = await current(p);
    for (const commit of [false, true]) {
      await open(p, 'pictureCompress'); await (await field(p, '압축 파일 형식')).selectOption('image/jpeg'); await fill(p, '압축 품질(%)', 60);
      await (await button(p, '압축 미리보기')).click(); await p.waitForFunction(() => !!document.querySelector('.picture-v2-compress-image')?.src);
      await (await button(p, '압축 적용')).click(); eq(await current(p), before); eq(await history(p), 0);
      if (!commit) { await dialog(p).getByRole('button', { name: '취소', exact: true }).click(); eq(await current(p), before); }
      else await confirm(p);
    }
    const after = await current(p); ok(after.src.startsWith('data:image/jpeg;base64,')); eq(after.originalSrc, before.src); eq([after.w, after.h, after.originalWidth, after.originalHeight], [320, 160, 640, 320]); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before); await run(p, 'redo'); eq(await current(p), after);
    await open(p, 'pictureCompress'); await (await button(p, '원본 그림 복원')).click(); await confirm(p); eq((await current(p)).src, before.src); eq((await current(p)).originalSrc, undefined);
  });

  await test('로컬 배경 제거는 흰색 연결 배경만 투명하게 하고 원본·실행 취소를 보존', async p => {
    const before = await current(p); await open(p, 'pictureBackground');
    await (await button(p, '배경 제거 미리보기')).click(); const apply = await button(p, '배경 제거 적용'); await p.waitForFunction(() => ![...document.querySelectorAll('button')].find(n => n.textContent.trim() === '배경 제거 적용')?.disabled);
    await apply.click(); eq(await current(p), before); eq(await history(p), 0); await confirm(p); const after = await current(p);
    eq(after.originalSrc, before.src); const data = await pixels(p, after.src, [[.02, .8], [.25, .5], [.65, .5]]);
    eq(data.values[0][3], 0); eq(data.values[1], [220, 38, 38, 255]); eq(data.values[2], [37, 99, 235, 255]); eq(await history(p), 1);
    await run(p, 'undo'); eq(await current(p), before); await run(p, 'redo'); eq(await current(p), after);
  });

  await test('그림 PNG 저장은 실제 색 보정 픽셀을 포함하고 원본·Undo를 유지', async p => {
    await fixture(p, { color: { recolor: 'grayscale' } }); const before = await current(p); await savePicker(p); await select(p);
    await run(p, 'pictureSave'); await p.waitForFunction(() => window.__pictureFiles?.length === 1);
    const file = await p.evaluate(() => window.__pictureFiles[0]); ok(file.name.endsWith('.png'));
    const image = await pixels(p, 'data:image/png;base64,' + Buffer.from(file.bytes).toString('base64'), [[.25, .5], [.65, .5]]);
    eq([image.w, image.h], [640, 320]);
    for (const [r, g, b, a] of image.values) { ok(Math.abs(r - g) <= 1 && Math.abs(g - b) <= 1, '원본의 빨강·파랑이 회색조 픽셀로 출력'); eq(a, 255); }
    eq(await current(p), before); eq(await history(p), 0);
  });

  await test('회전90도 그림 저장·메모리 클립보드 PNG는 세로 방향과 실제 색 위치 유지', async p => {
    await fixture(p, { rot: 90 }); const before = await current(p); await savePicker(p); await memoryClipboard(p);
    await p.locator('.obj.pic[data-id="parity-picture"]').first().click(); await run(p, 'pictureSave'); await p.waitForFunction(() => window.__pictureFiles?.length === 1);
    await p.locator('.obj.pic[data-id="parity-picture"]').first().click({ button: 'right' }); await p.locator('.menu-item').filter({ hasText: /^복사\(C\)/ }).click();
    await p.waitForFunction(() => window.__pictureClipboard.writes.length === 1);
    const files = await p.evaluate(() => [window.__pictureFiles[0], window.__pictureClipboard.writes[0]]);
    for (const file of files) {
      const data = await pixels(p, 'data:image/png;base64,' + Buffer.from(file.bytes).toString('base64'), [[.5, .25], [.5, .65]]);
      eq([data.w, data.h], [320, 640]); eq(data.values[0], [220, 38, 38, 255]); eq(data.values[1], [37, 99, 235, 255]);
    }
    eq(await current(p), before); eq(await history(p), 0);
  });

  await test('비동기 XLSX 내보내기는 보정 픽셀·원본 편집 서식을 보존하고 현재 문서를 변경하지 않음', async p => {
    await fixture(p, { color: { saturation: 0 }, correction: { brightness: .1 } }); const before = await current(p); await savePicker(p);
    await p.evaluate(() => window.tabula.exportXlsx('합성 그림 저장')); await p.waitForFunction(() => window.__pictureFiles?.length === 1);
    const file = await p.evaluate(() => window.__pictureFiles[0]), bytes = Uint8Array.from(file.bytes), saved = readXlsx(bytes).data.sheets[0].images[0];
    eq(saved.src, before.src); eq(saved.color.saturation, before.color.saturation); eq(saved.correction.brightness, before.correction.brightness);
    const files = unzip(bytes), decoder = new TextDecoder(), encoder = new TextEncoder();
    for (const name of Object.keys(files)) if (/^xl\/drawings\/drawing\d+\.xml$/.test(name)) files[name] = encoder.encode(decoder.decode(files[name]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g, ''));
    const native = readXlsx(zip(files)).data.sheets[0].images[0]; ok(native.src !== before.src, '전용 메타 없이 Excel에 전달되는 보정 PNG');
    const image = await pixels(p, native.src, [[.25, .5], [.65, .5]]);
    for (const [r, g, b, a] of image.values) { ok(Math.abs(r - g) <= 1 && Math.abs(g - b) <= 1); eq(a, 255); }
    eq(await current(p), before); eq(await history(p), 0);
  });

  await test('그림 레이아웃 변환은 취소 시 보존·확인 한 번 Undo·보정 이미지 포함', async p => {
    await fixture(p, { color: { recolor: 'grayscale' } }); const before = await current(p); await select(p); await run(p, 'pictureLayout');
    const smart = p.getByRole('dialog', { name: 'SmartArt 그래픽', exact: true }); await smart.waitFor();
    await smart.getByRole('button', { name: '취소', exact: true }).click(); eq(await current(p), before); eq(await history(p), 0);
    await select(p); await run(p, 'pictureLayout'); await smart.waitFor(); await smart.getByRole('button', { name: '확인', exact: true }).click(); await smart.waitFor({ state: 'detached' });
    const state = await p.evaluate(() => { const s = window.tabula.wb().sheets[0]; return { images: s.images.length, shapes: structuredClone(s.shapes) }; });
    eq(state.images, 0); eq(state.shapes.length, 1); eq(state.shapes[0].smartArt.layout, 'pictureCards'); eq(await history(p), 1);
    const image = await pixels(p, state.shapes[0].smartArt.nodes[0].picture, [[.25, .5]]); const [r, g, b] = image.values[0]; ok(Math.abs(r - g) <= 1 && Math.abs(g - b) <= 1);
    await run(p, 'undo'); eq(await current(p), before); eq(await p.evaluate(() => window.tabula.wb().sheets[0].shapes.length), 0);
  });

  await test('320px 화면에서 서식 탭·본문 스크롤·확인 취소가 화면 안에 유지', async p => {
    await p.setViewportSize({ width: 320, height: 680 }); await open(p);
    const bounds = await dialog(p).evaluate(el => { const r = el.getBoundingClientRect(); return { x: r.x, y: r.y, right: r.right, bottom: r.bottom, over: el.scrollWidth > el.clientWidth + 2, width: innerWidth, height: innerHeight }; });
    ok(bounds.x >= -1 && bounds.y >= -1 && bounds.right <= bounds.width + 1 && bounds.bottom <= bounds.height + 1, JSON.stringify(bounds)); eq(bounds.over, false);
    for (const name of ['확인', '취소']) ok(await dialog(p).getByRole('button', { name, exact: true }).isVisible());
    await p.screenshot({ path: process.env.WIXEL_PICTURE_MOBILE_SCREENSHOT || 'D:/Codex/Temp/wixel-picture-format-mobile.png' });
    await p.keyboard.press('Escape'); eq(await dialog(p).count(), 0);
  });
  console.log(JSON.stringify({ cases: results.length, checks, bad: results.filter(r => !r.ok).length, ribbonCoverage, results }));
  if (results.some(r => !r.ok)) process.exitCode = 1;
} finally { await browser.close(); }
