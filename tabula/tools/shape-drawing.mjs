// 도형 그리기·점 편집·선 서식: 합성 문서, 격리 브라우저, API/원격 쓰기 차단.
// WIXEL_URL(소스/컴파일), WIXEL_SHAPE_FILTER, WIXEL_SHAPE_SCREENSHOT 선택 가능.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const browser = await chromium.launch(), results = [];
const filter = process.env.WIXEL_SHAPE_FILTER, skip = process.env.WIXEL_SHAPE_SKIP;
async function test(name, fn) {
  if ((filter && !name.includes(filter)) || (skip && name.includes(skip))) return;
  const context = await browser.newContext({ viewport: { width: 1500, height: 1000 }, acceptDownloads: false });
  const p = await context.newPage(), errors = [], writes = []; p.setDefaultTimeout(10000);
  p.on('pageerror', e => errors.push(e.message));
  p.on('dialog', d => d.dismiss());
  await context.route('**/*', route => {
    const r = route.request(), target = new URL(r.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(r.method()); return route.abort(); }
    return target.origin === new URL(url).origin && !target.pathname.startsWith('/api/') ? route.continue() : route.abort();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb());
    await p.evaluate(() => {
      const t = window.tabula, w = t.wb();
      w.restore({ sheets: [{ name: '도형 합성', cells: { '0,0': { raw: '보존' }, '0,1': { raw: '=1+1', cached: 42 } } }] });
      w.undoStack = []; w.redoStack = []; t.gv().layout(); t.gv().renderAll(); t.selectCell(0, 0);
    });
    await p.locator('#cellEditor').focus(); await fn(p);
    assert.equal(await p.evaluate(() => window.tabula.wb().getRaw(0, 0, 0)), '보존', '그리기가 셀 값을 바꿈');
    assert.equal(await p.evaluate(() => window.tabula.wb().getRaw(0, 0, 1)), '=1+1', '그리기가 수식을 바꿈');
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (error) { results.push({ name, ok: false, error: error.message, pageErrors: errors, blockedWrites: writes }); console.error('NG ' + name + ': ' + error.stack); }
  finally { await context.close(); }
}
const shapes = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets[0].shapes));
const shape = async p => (await shapes(p))[0];
const undo = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+z'); };
const redo = async p => { await p.locator('#cellEditor').focus(); await p.keyboard.press('Control+y'); };
const shapeLabel = name => ({ 자유형: '자유형: 도형', 자유곡선: '자유형: 자유곡선' })[name] ?? name;
async function pick(p, name) {
  await p.locator('#cellEditor').focus();
  for (const key of ['Alt', 'n', 's', 'h']) await p.keyboard.press(key);
  await p.locator('.shape-gallery').waitFor();
  await p.locator(`.shape-btn[title="${shapeLabel(name)}"]`).click();
}
const origin = p => p.evaluate(() => { const r = document.getElementById('gridView').getBoundingClientRect(); return { x: r.left + 170, y: r.top + 140 }; });
const commands = s => s.path?.paths.flatMap(path => path.commands) ?? [];
async function draw(p, kind, close = false) {
  await pick(p, kind); const a = await origin(p);
  await p.mouse.click(a.x, a.y); await p.mouse.click(a.x + 120, a.y + 80); await p.mouse.click(a.x + 200, a.y - 20);
  if (close) await p.mouse.click(a.x + 1, a.y + 1);
  else await p.mouse.dblclick(a.x + 280, a.y + 90);
  assert.equal((await shapes(p)).length, 1, '한 번 그리기에 도형 하나');
  return a;
}
async function fixture(p, patch = {}) {
  await p.evaluate(patch => {
    const t = window.tabula, w = t.wb(); w.transact(() => w.setSheetProp(0, 'shapes', [{ id: 'shape-fixture', kind: 'line', x: 190, y: 130, w: 260, h: 0, stroke: '#4472c4', strokeWidth: 2, fill: null, text: '', ...patch }]));
    t.gv().renderObjectsAll(); w.undoStack = []; w.redoStack = [];
  }, patch);
}
async function format(p) {
  await p.locator('.obj[data-id="shape-fixture"]').first().dblclick({ position: { x: 60, y: 4 } });
  await p.locator('.shape-format-pane').waitFor();
  return p.getByRole('dialog').filter({ has: p.locator('.shape-format-pane') });
}
const number = async (pane, name, value) => { const input = pane.getByLabel(name, { exact: true }); await input.fill(String(value)); await input.press('Tab'); };
try {
  await test('선 갤러리에 곡선·자유형·자유곡선과 기본 선이 표시됨', async p => {
    for (const key of ['Alt', 'n', 's', 'h']) await p.keyboard.press(key);
    for (const name of ['선', '곡선', '자유형', '자유곡선']) assert.equal(await p.locator(`.shape-btn[title="${shapeLabel(name)}"]`).count(), 1, name);
    await p.keyboard.press('Escape'); assert.equal((await shapes(p)).length, 0);
  });
  await test('곡선: 클릭으로 경유점 추가·더블클릭 열린 곡선 완료·Undo/Redo', async p => {
    await draw(p, '곡선'); const s = await shape(p); assert.equal(s.kind, 'curve'); assert.ok(commands(s).some(c => c[0] === 'C')); assert.ok(!commands(s).some(c => c[0] === 'Z'));
    await undo(p); assert.equal((await shapes(p)).length, 0); await redo(p); assert.deepEqual(await shape(p), s);
  });
  await test('곡선: 시작점 근처 클릭으로 닫힌 경로 완료', async p => {
    await draw(p, '곡선', true); assert.ok(commands(await shape(p)).some(c => c[0] === 'Z')); assert.ok((await shape(p)).fill);
  });
  await test('자유형: 클릭 직선 구간과 시작점 닫기', async p => {
    await draw(p, '자유형', true); const s = await shape(p); assert.equal(s.kind, 'freeform'); assert.ok(commands(s).filter(c => c[0] === 'L').length >= 2); assert.equal(commands(s).at(-1)[0], 'Z');
  });
  await test('자유형: 클릭과 드래그를 섞고 열린 경로 종료', async p => {
    await pick(p, '자유형'); const a = await origin(p); await p.mouse.click(a.x, a.y); await p.mouse.move(a.x + 70, a.y + 70); await p.mouse.down(); await p.mouse.move(a.x + 130, a.y + 10, { steps: 12 }); await p.mouse.up(); await p.mouse.dblclick(a.x + 250, a.y + 80);
    const s = await shape(p); assert.equal((await shapes(p)).length, 1); assert.equal(s.kind, 'freeform'); assert.ok(commands(s).length >= 4); assert.ok(!commands(s).some(c => c[0] === 'Z'));
  });
  await test('자유곡선: 연속 드래그 경로를 만들고 한 번 Undo', async p => {
    await pick(p, '자유곡선'); const a = await origin(p); await p.mouse.move(a.x, a.y); await p.mouse.down();
    for (let i = 1; i <= 28; i++) await p.mouse.move(a.x + i * 7, a.y + Math.sin(i / 4) * 45);
    await p.mouse.up(); const s = await shape(p); assert.equal((await shapes(p)).length, 1); assert.equal(s.kind, 'scribble'); assert.ok(commands(s).length >= 4);
    await undo(p); assert.equal((await shapes(p)).length, 0);
  });
  await test('그리기 시작 전·점 추가 중 Esc는 미완성 도형과 Undo 기록을 남기지 않음', async p => {
    await pick(p, '자유형'); await p.keyboard.press('Escape'); assert.equal((await shapes(p)).length, 0);
    await pick(p, '곡선'); const a = await origin(p); await p.mouse.click(a.x, a.y); await p.mouse.click(a.x + 90, a.y + 60); await p.keyboard.press('Escape');
    assert.equal((await shapes(p)).length, 0); assert.equal(await p.evaluate(() => window.tabula.wb().undoStack.length), 0);
  });
  await test('직선: 드래그 중 Esc 후 mouseup도 취소 상태 유지', async p => {
    await pick(p, '선'); const a = await origin(p); await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(a.x + 220, a.y + 80, { steps: 6 }); await p.keyboard.press('Escape'); await p.mouse.up();
    assert.equal((await shapes(p)).length, 0); assert.equal(await p.evaluate(() => window.tabula.wb().undoStack.length), 0);
  });
  await test('직선: 가로 선의 높이0과 기본 화살표·한 번 Undo', async p => {
    await pick(p, '선 화살표'); const a = await origin(p); await p.mouse.move(a.x, a.y); await p.mouse.down(); await p.mouse.move(a.x + 220, a.y, { steps: 6 }); await p.mouse.up();
    const s = await shape(p); assert.equal(s.h, 0); assert.ok(s.arrow === 'end' || s.tailEnd?.type && s.tailEnd.type !== 'none');
    assert.ok(await p.locator('.obj.shape svg').count()); await undo(p); assert.equal((await shapes(p)).length, 0);
  });
  await test('선 서식: 색·투명도·pt 너비·대시·끝/연결/복합 설정 실제 반영', async p => {
    await fixture(p); const pane = await format(p);
    await pane.getByLabel('선 색', { exact: true }).fill('#123456'); await pane.getByLabel('선 색', { exact: true }).dispatchEvent('change');
    await number(pane, '선 투명도(%)', 25); await number(pane, '너비(pt)', 3);
    await pane.getByLabel('대시 종류', { exact: true }).selectOption('dashDot'); await pane.getByLabel('끝 모양', { exact: true }).selectOption('sq');
    await pane.getByLabel('연결 모양', { exact: true }).selectOption('bevel'); await pane.getByLabel('복합 종류', { exact: true }).selectOption('dbl');
    const s = await shape(p); assert.equal(s.stroke, '#123456'); assert.equal(s.strokeOpacity, .75); assert.equal(s.strokeWidth, 4); assert.equal(s.dash, 'dashDot'); assert.equal(s.lineCap, 'sq'); assert.equal(s.lineJoin, 'bevel'); assert.equal(s.compound, 'dbl');
    assert.doesNotMatch(await p.locator('.obj.shape svg').first().evaluate(svg => svg.outerHTML), /NaN|Infinity/);
    await pane.getByRole('button', { name: '닫기', exact: true }).click(); await undo(p); assert.notEqual((await shape(p)).compound, 'dbl');
  });
  await test('선 서식: 시작·끝 화살표 종류/너비/길이를 독립 적용', async p => {
    await fixture(p); const pane = await format(p);
    for (const [name, type, w, len] of [['시작', 'oval', 'lg', 'sm'], ['끝', 'triangle', 'sm', 'lg']]) {
      await pane.getByLabel(`${name} 화살표 종류`, { exact: true }).selectOption(type);
      await pane.getByLabel(`${name} 화살표 너비`, { exact: true }).selectOption(w);
      await pane.getByLabel(`${name} 화살표 길이`, { exact: true }).selectOption(len);
    }
    const s = await shape(p); assert.deepEqual(s.headEnd, { type: 'oval', w: 'lg', len: 'sm' }); assert.deepEqual(s.tailEnd, { type: 'triangle', w: 'sm', len: 'lg' });
    if (process.env.WIXEL_SHAPE_SCREENSHOT) await p.screenshot({ path: process.env.WIXEL_SHAPE_SCREENSHOT });
    await pane.getByRole('button', { name: '닫기', exact: true }).click(); await undo(p); assert.notEqual((await shape(p)).tailEnd.len, 'lg');
  });
  await test('선 서식: 비율 고정 가로선·세로선 치수 변경이 유한하고 높이0을 보존', async p => {
    await fixture(p); const pane = await format(p); await pane.getByRole('tab', { name: '크기 및 속성', exact: true }).click();
    await pane.getByLabel('가로 세로 비율 고정', { exact: true }).check(); await number(pane, '너비(px)', 400);
    let s = await shape(p); assert.equal(s.w, 400); assert.equal(s.h, 0); assert.equal(s.lockAspect, true);
    await number(pane, '높이(px)', 100); s = await shape(p); assert.ok(Number.isFinite(s.w) && Number.isFinite(s.h));
    await number(pane, '회전(°)', 45); assert.equal((await shape(p)).rot, 45);
  });
  await test('도형 서식: 탭 방향키·텍스트·채우기·즉시적용 후 닫기 유지', async p => {
    await fixture(p, { kind: 'rect', h: 120, fill: '#4472c4', text: '합성 도형' }); const pane = await format(p);
    await pane.getByLabel('채우기 색', { exact: true }).fill('#13579b'); await pane.getByLabel('채우기 색', { exact: true }).dispatchEvent('change');
    await pane.getByRole('radio', { name: '그라데이션 채우기', exact: true }).check(); assert.equal((await shape(p)).grad.stops[0][1], '#13579b', '모드 전환은 최신 채우기색 보존');
    await pane.getByRole('button', { name: '중지점 추가', exact: true }).click(); assert.equal((await shape(p)).grad.stops.length, 3);
    await number(pane, '중지점 투명도(%)', 70); await pane.getByLabel('중지점 색', { exact: true }).fill('#345678'); await pane.getByLabel('중지점 색', { exact: true }).dispatchEvent('change');
    assert.ok(Math.abs((await shape(p)).grad.stops[1][2] - .3) < 1e-9, '중지점 색 변경은 투명도 보존');
    await pane.getByRole('tab', { name: '효과', exact: true }).click(); await pane.getByLabel('그림자', { exact: true }).check();
    await number(pane, '그림자 투명도(%)', 65); await pane.getByLabel('그림자 색', { exact: true }).fill('#123456'); await pane.getByLabel('그림자 색', { exact: true }).dispatchEvent('change');
    assert.ok(Math.abs((await shape(p)).shadow.opacity - .35) < 1e-9, '그림자 색 변경은 투명도 보존');
    const first = pane.getByRole('tab', { name: '채우기 및 선', exact: true }); await first.focus(); await p.keyboard.press('End');
    assert.equal(await p.evaluate(() => document.activeElement.textContent), '크기 및 속성');
    await pane.getByRole('tab', { name: '텍스트 옵션', exact: true }).click();
    await pane.getByLabel('도형 텍스트', { exact: true }).fill('수정된 합성 도형'); await pane.getByLabel('도형 텍스트', { exact: true }).press('Tab');
    await pane.getByRole('button', { name: '닫기', exact: true }).click(); assert.equal((await shape(p)).text, '수정된 합성 도형');
    await undo(p); assert.equal((await shape(p)).text, '합성 도형');
  });
  await test('점 편집: 꼭짓점 이동·한 번 Undo/Redo', async p => {
    await draw(p, '자유형', true); const before = await shape(p);
    await p.evaluate(() => window.tabula.run('shapeEditPoints'));
    const handle = p.locator('.shape-point.anchor').first(); await handle.waitFor(); const box = await handle.boundingBox();
    await p.mouse.move(box.x + box.width / 2, box.y + box.height / 2); await p.mouse.down(); await p.mouse.move(box.x + 40, box.y + 25, { steps: 5 }); await p.mouse.up();
    const after = await shape(p); assert.notDeepEqual(after, before); await undo(p); assert.deepEqual(await shape(p), before); await redo(p); assert.deepEqual(await shape(p), after);
  });
  await test('점 편집: 드래그 Esc는 전체 모양·Undo 깊이를 복원', async p => {
    await draw(p, '곡선'); const before = await shape(p); const depth = await p.evaluate(() => window.tabula.wb().undoStack.length);
    await p.evaluate(() => window.tabula.run('shapeEditPoints')); const handle = p.locator('.shape-point.anchor').first(); await handle.waitFor(); const box = await handle.boundingBox();
    await p.mouse.move(box.x + box.width / 2, box.y + box.height / 2); await p.mouse.down(); await p.mouse.move(box.x + 50, box.y + 35, { steps: 5 }); await p.keyboard.press('Escape'); await p.mouse.up();
    assert.deepEqual(await shape(p), before); assert.equal(await p.evaluate(() => window.tabula.wb().undoStack.length), depth);
  });
  await test('점 편집: Ctrl+윤곽 추가·Ctrl+점 삭제·Undo', async p => {
    const a = await draw(p, '자유형', true); await p.evaluate(() => window.tabula.run('shapeEditPoints'));
    await p.locator('.shape-point.anchor').first().waitFor(); const before = await shape(p), count = await p.locator('.shape-point.anchor').count();
    await p.keyboard.down('Control'); await p.mouse.click(a.x + 60, a.y + 40); await p.keyboard.up('Control');
    assert.equal(await p.locator('.shape-point.anchor').count(), count + 1, '윤곽의 중간점 추가');
    await p.locator('.shape-point.anchor').first().click({ modifiers: ['Control'] }); assert.equal(await p.locator('.shape-point.anchor').count(), count);
    await undo(p); assert.equal(await p.locator('.shape-point.anchor').count(), count + 1); await undo(p); assert.deepEqual(await shape(p), before);
  });
  await test('점 편집: 선택점 Delete는 도형 전체를 지우지 않음', async p => {
    await draw(p, '곡선'); await p.evaluate(() => window.tabula.run('shapeEditPoints'));
    await p.locator('.shape-point.anchor').first().waitFor(); const before = await shape(p); const count = await p.locator('.shape-point.anchor').count();
    await p.locator('.shape-point.anchor').nth(1).click(); await p.keyboard.press('Delete');
    assert.equal((await shapes(p)).length, 1); assert.equal(await p.locator('.shape-point.anchor').count(), count - 1);
    await undo(p); assert.deepEqual(await shape(p), before);
  });
  await test('보호된 시트는 새 경로와 기존 점 편집을 막음', async p => {
    await draw(p, '자유형', true); const before = await shapes(p);
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true, objects: false } })); window.tabula.run('shapeEditPoints'); });
    assert.equal(await p.locator('.shape-point').count(), 0);
    const blocked = p.getByRole('dialog'); if (await blocked.count()) await blocked.getByRole('button', { name: '확인', exact: true }).click();
    await p.locator('#cellEditor').focus(); for (const key of ['Alt', 'n', 's', 'h']) await p.keyboard.press(key);
    const curve = p.locator('.shape-btn[title="곡선"]'); await curve.waitFor();
    if (await curve.isEnabled()) { await curve.click(); const a = await origin(p); await p.mouse.click(a.x + 350, a.y); await p.mouse.dblclick(a.x + 440, a.y + 80); }
    else await p.keyboard.press('Escape');
    assert.deepEqual(await shapes(p), before);
  });
  await test('열린 도형 서식 중 보호로 전환하면 후속 변경 차단', async p => {
    await fixture(p); const pane = await format(p); const before = await shape(p);
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { objects: false } })); });
    const color = pane.getByLabel('선 색', { exact: true }); if (await color.isEnabled()) { await color.fill('#ff00ff'); await color.dispatchEvent('change'); }
    assert.deepEqual(await shape(p), before);
  });
  for (const modifier of ['Shift', 'Control', 'Alt']) await test(`점 편집: ${modifier} 제어 손잡이의 부드러운/직선/모서리 관계`, async p => {
    await fixture(p, { kind: 'curve', w: 260, h: 180, path: { paths: [{ fill: false, stroke: true, commands: [['M', 0, 0], ['C', .1, .05, .3, .7, .5, .5], ['C', .7, .3, .8, 1, 1, .9]] }] } });
    await p.locator('.obj[data-id="shape-fixture"]').first().click({ position: { x: 20, y: 10 } }); await p.evaluate(() => window.tabula.run('shapeEditPoints'));
    const before = await shape(p), handle = p.locator('.shape-point.control[data-point="0:1:3"]').first(); await handle.waitFor(); const box = await handle.boundingBox();
    await p.keyboard.down(modifier); await p.mouse.move(box.x + box.width / 2, box.y + box.height / 2); await p.mouse.down(); await p.mouse.move(box.x + 25, box.y + 35, { steps: 5 }); await p.mouse.up(); await p.keyboard.up(modifier);
    const after = await shape(p), c = after.path.paths[0].commands, old = before.path.paths[0].commands;
    const a = [c[1][3] - c[1][5], c[1][4] - c[1][6]], b = [c[2][1] - c[1][5], c[2][2] - c[1][6]];
    assert.notDeepEqual(c[1].slice(3, 5), old[1].slice(3, 5), '제어점 이동');
    if (modifier === 'Alt') {
      const dx = after.x + c[2][1] * after.w - (before.x + old[2][1] * before.w);
      const dy = after.y + c[2][2] * after.h - (before.y + old[2][2] * before.h);
      assert.ok(Math.hypot(dx, dy) < .01, '모서리는 반대 손잡이의 실제 위치 유지');
    }
    else {
      assert.ok(Math.abs(a[0] * b[1] - a[1] * b[0]) < 1e-6 && a[0] * b[0] + a[1] * b[1] < 0, '두 손잡이는 반대 공선 방향');
      if (modifier === 'Shift') assert.ok(Math.hypot(a[0] + b[0], a[1] + b[1]) < 1e-6, '부드러운 점의 대칭 길이');
      else {
        const oldLength = Math.hypot((old[2][1] - old[1][5]) * before.w, (old[2][2] - old[1][6]) * before.h);
        assert.ok(Math.abs(Math.hypot(b[0] * after.w, b[1] * after.h) - oldLength) < .01, '직선 점은 반대 손잡이의 실제 길이 유지');
      }
    }
    await undo(p); assert.deepEqual(await shape(p), before);
  });
  await test('닫힌 곡선은 채우기·텍스트가 있는 도형 서식으로 편집', async p => {
    await draw(p, '곡선', true); await p.locator('.obj.shape').first().click({ button: 'right', position: { x: 30, y: 20 } });
    await p.locator('.menu-item').filter({ hasText: /도형 서식|선 서식/ }).click();
    const pane = p.getByRole('dialog', { name: '도형 서식', exact: true }); await pane.waitFor();
    await pane.getByRole('radio', { name: '단색 채우기', exact: true }).check(); assert.ok((await shape(p)).fill);
    await pane.getByRole('tab', { name: '텍스트 옵션', exact: true }).click(); await pane.getByLabel('도형 텍스트', { exact: true }).fill('닫힌 곡선'); await pane.getByLabel('도형 텍스트', { exact: true }).press('Tab');
    assert.equal((await shape(p)).text, '닫힌 곡선');
  });
  await test('사용자 경로의 모양 변경은 자유 경로·옛 geometry를 제거하고 사각형 표시·Undo', async p => {
    await fixture(p, { kind: 'freeform', h: 140, fill: '#4472c4', path: { paths: [{ fill: true, stroke: true, commands: [['M', 0, 0], ['L', 1, .2], ['L', .4, 1], ['Z']] }] }, customGeometry: '<a:custGeom/>' });
    const before = await shape(p), object = p.locator('.obj[data-id="shape-fixture"]').first();
    const pathBefore = await object.locator('svg').evaluate(svg => svg.innerHTML);
    await object.click({ position: { x: 70, y: 45 } }); await p.evaluate(() => window.tabula.run('shapeEditPoints'));
    await p.locator('.shape-point.anchor').first().waitFor();
    await object.click({ button: 'right', position: { x: 70, y: 45 } }); await p.getByRole('menuitem', { name: '도형 모양 변경...', exact: true }).click();
    await p.locator('.shape-btn[title="사각형"]').click(); const after = await shape(p);
    assert.equal(after.kind, 'rect'); assert.equal(after.path, undefined); assert.equal(after.customGeometry, undefined);
    assert.notEqual(await object.locator('svg').evaluate(svg => svg.innerHTML), pathBefore, '실제 SVG 모양도 변경');
    assert.equal(await p.locator('.shape-point').count(), 0, '모양 변경은 이전 경로의 점 편집 상태도 종료');
    await undo(p); assert.deepEqual(await shape(p), before);
  });
  await test('인쇄: 제외한 도형·그림·차트를 숨기고 회전된 긴 서식 텍스트의 축소 보존', async p => {
    await p.evaluate(() => {
      const t = window.tabula, w = t.wb(), text = ('매출 보고서와 긴 텍스트 서식 확인 2026\n').repeat(8);
      const shape = { id: 'print-visible', kind: 'rect', x: 100, y: 80, w: 320, h: 190, fill: '#ffffff', stroke: '#4472c4',
        text, size: 26, pad: [13, 19, 23, 29], valign: 'middle', textFit: 'shrink', textRot: 90,
        paras: [{ runs: [{ t: text.slice(0, 100), sz: 36, b: true, color: '#123456' }, { t: text.slice(100), sz: 20, i: true }] }] };
      const canvas = document.createElement('canvas'); canvas.width = 2; canvas.height = 2;
      canvas.getContext('2d').fillRect(0, 0, 2, 2);
      const image = { id: 'image-visible', src: canvas.toDataURL(), x: 10, y: 10, w: 12, h: 12 };
      const chart = { id: 'chart-visible', type: 'column', title: '인쇄 포함 차트', range: { r1: 2, c1: 0, r2: 4, c2: 1 }, x: 20, y: 20, w: 300, h: 180 };
      w.transact(() => {
        [['분류', '값'], ['가', '10'], ['나', '20']].forEach((row, r) => row.forEach((value, c) => w.setInput(0, r + 2, c, value)));
        w.setSheetProp(0, 'shapes', [shape, { ...shape, id: 'print-hidden', noPrint: true }]);
        w.setSheetProp(0, 'images', [image, { ...image, id: 'image-hidden', noPrint: true }]);
        w.setSheetProp(0, 'charts', [chart, { ...chart, id: 'chart-hidden', title: '인쇄 제외 차트', noPrint: true }]);
      });
      t.gv().layout(); t.gv().renderAll(); w.undoStack = []; w.redoStack = [];
      window.__printBefore = JSON.stringify(w.serialize()); window.__printCalls = 0;
      window.print = () => { window.__printCalls++; };
      t.run('print');
    });
    await p.emulateMedia({ media: 'print' });
    const actual = await p.evaluate(() => {
      const area = document.getElementById('printArea'), node = area.querySelector('[data-shape-print="print-visible"]');
      const box = node.querySelector('.sh-text'), content = box.firstElementChild, css = getComputedStyle(box), scale = Number(content.style.zoom || 1);
      const range = document.createRange(); range.selectNodeContents(content);
      const rect = element => { const r = element.getBoundingClientRect(); return { x: r.x, y: r.y, right: r.right, bottom: r.bottom }; };
      return { calls: window.__printCalls, preview: !!document.querySelector('.dialog[aria-label="인쇄 미리보기"]'), shapeIds: [...area.querySelectorAll('[data-shape-print]')].map(el => el.dataset.shapePrint),
        images: area.querySelectorAll('[data-print-object="image-visible"] svg image').length, objects: area.querySelectorAll('.chart-print').length, text: area.textContent,
        fit: box.dataset.fitDone, scale, transform: box.style.transform, rich: box.classList.contains('rich'), bold: [...content.querySelectorAll('span')].some(el => Number(getComputedStyle(el).fontWeight) >= 700),
        available: [box.clientWidth - parseFloat(css.paddingLeft) - parseFloat(css.paddingRight), box.clientHeight - parseFloat(css.paddingTop) - parseFloat(css.paddingBottom)],
        scroll: [content.scrollWidth * scale, content.scrollHeight * scale], box: rect(box), actual: rect(range),
        same: JSON.stringify(window.tabula.wb().serialize()) === window.__printBefore, undoDepth: window.tabula.wb().undoStack.length };
    });
    assert.equal(actual.calls, 0, '미리보기를 열 때 OS 인쇄를 자동 실행하지 않음'); assert.equal(actual.preview, true, '사용자가 확인할 인쇄 미리보기 열림'); assert.deepEqual(actual.shapeIds, ['print-visible']);
    assert.equal(actual.images, 1); assert.equal(actual.objects, 3); assert.match(actual.text, /인쇄 포함 차트/); assert.doesNotMatch(actual.text, /인쇄 제외 차트/);
    assert.equal(actual.fit, 'print'); assert.equal(actual.rich, true); assert.equal(actual.bold, true); assert.match(actual.transform, /rotate\(90deg\)/);
    assert.ok(actual.scale > 0 && actual.scale < 1, '인쇄용 긴 서식 텍스트 축소');
    assert.ok(actual.scroll[0] <= actual.available[0] + 1 && actual.scroll[1] <= actual.available[1] + 1, '인쇄 텍스트가 내부 여백 안에 맞음: ' + JSON.stringify(actual));
    assert.ok(actual.actual.x >= actual.box.x - 1 && actual.actual.y >= actual.box.y - 1 && actual.actual.right <= actual.box.right + 1 && actual.actual.bottom <= actual.box.bottom + 1, '회전 후 인쇄 글자 경계가 도형 안에 있음');
    assert.equal(actual.same, true, '인쇄는 문서 모델을 변경하지 않음'); assert.equal(actual.undoDepth, 0);
  });
  const failed = results.filter(r => !r.ok).length;
  console.log(JSON.stringify({ url, tests: results.length, failed, results }, null, 2)); if (failed) process.exitCode = 1;
} finally { await browser.close(); }
