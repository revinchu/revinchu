// Alt 키팁 등록표·모든 접두 경로·실제 동작 회귀. 합성 문서와 새 브라우저 컨텍스트만 사용한다.
// PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH, WIXEL_URL은 tools/README.md의 다른 브라우저 도구와 같다.
import assert from 'node:assert/strict';
import { EXCEL_KEYTIP_COMPAT } from '../src/excel-keytip-compat.js';
import { readFile } from 'node:fs/promises';

const source = await readFile(new URL('../src/app.js', import.meta.url), 'utf8');
const table = source.match(/const KEYTIPS\s*=\s*\{([\s\S]*?)\n\};/)?.[1];
assert.ok(table, 'app.js KEYTIPS 등록표를 찾을 수 없음');
const clean = table.replace(/\/\/[^\n]*|\/\*[\s\S]*?\*\//g, '');
const entryPattern = /\b([a-z][a-z0-9]*)\s*:\s*\[\s*'([^']+)'\s*,\s*'((?:\\.|[^'])*)'\s*\]/g;
const entries = [...clean.matchAll(entryPattern)].map((m) => ({ key: m[1], command: m[2], label: m[3] }));
for (const compatible of EXCEL_KEYTIP_COMPAT) if (!entries.some(e => e.key === compatible.path)) entries.push({ key: compatible.path, command: compatible.target, kind: compatible.kind, label: compatible.label });
const parsedRemainder = clean.replace(entryPattern, '').replace(/[\s,]/g, '');
const tabsText = source.match(/const KEYTIP_TABS\s*=\s*\{([^}]+)\}/)?.[1] ?? '';
const tabs = [...tabsText.matchAll(/\b([a-z])\s*:\s*'([^']+)'/g)].map((m) => m[1]);
const counts = new Map();
for (const { key } of entries) counts.set(key, (counts.get(key) ?? 0) + 1);
const duplicates = [...counts].filter(([, n]) => n > 1).map(([key, n]) => ({ key, n }));
const keys = [...counts.keys()];
const prefixConflicts = keys.flatMap((a) => keys.filter((b) => a !== b && b.startsWith(a)).map((b) => `${a} → ${b}`));
const prefixes = [...new Set([...tabs, ...keys.flatMap((key) => Array.from({ length: key.length - 1 }, (_, i) => key.slice(0, i + 1)))])].sort();
const coverage = { registeredKeytips: entries.length, uniqueCommands: new Set(entries.map((e) => e.command)).size, registeredTabs: tabs.length, duplicateKeys: duplicates, terminalPrefixConflicts: prefixConflicts, checkedPrefixes: [], verifiedActionKeys: [] };
const results = [];
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const url = process.env.WIXEL_URL || 'http://localhost:5178/';
const actionKeys = new Set();

async function test(name, run) {
  const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: false });
  const page = await context.newPage();
  const errors = [], writes = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await context.route('**/*', (route) => {
    const request = route.request();
    if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) {
      writes.push(`${request.method()} ${request.url()}`); return route.abort('blockedbyclient');
    }
    return route.continue();
  });
  try {
    await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await page.goto(url);
    await page.waitForFunction(() => !!window.tabula?.wb());
    await page.evaluate(() => {
      const t = window.tabula, wb = t.wb();
      wb.transact(() => [['항목', '수량'], ['가', '10'], ['나', '20']].forEach((row, r) => row.forEach((raw, c) => wb.setInput(0, r, c, raw))));
      t.selectCell(1, 1);
    });
    await run(page);
    assert.deepEqual(errors, [], '페이지 스크립트 오류');
    assert.deepEqual(writes, [], '검증 중 서버 쓰기 요청이 발생함 (요청은 차단됨)');
    results.push({ name, ok: true }); console.log(`OK ${name}`);
  } catch (error) {
    results.push({ name, ok: false, error: error.message, pageErrors: errors, blockedWrites: writes });
    console.error(`NG ${name}: ${error.stack}`);
  } finally { await context.close(); }
}
const state = (p) => p.evaluate(() => ({
  keytips: document.body.classList.contains('keytips'), sequence: document.body.dataset.keytipSequence, status: document.getElementById('statusMode').textContent,
  editor: document.getElementById('cellEditor').value, editing: !document.getElementById('cellEditor').classList.contains('idle'),
  raw: window.tabula.wb().getCell(0, 1, 1)?.raw, gridHidden: !!window.tabula.wb().sheets[0].noGrid,
  headerWidth: window.tabula.gv().hw, tab: document.querySelector('.ribbon-tab.active')?.textContent,
}));
const sequence = async (p, key, start = 'Alt') => { if (start) await p.keyboard.press(start); for (const ch of key) await p.keyboard.press(ch); };
const path = async (p, prefix) => {
  const s = await state(p);
  assert.equal(s.keytips, true, `키팁 경로 ${prefix || '(루트)'}가 닫힘: ${JSON.stringify(s)}`);
  assert.equal(s.sequence, prefix, `현재 키팁 경로: ${JSON.stringify(s.sequence)}`);
  assert.equal(await p.locator('.keytip-panel').count(), 0, '제거한 큰 리본 바로 가기 키 안내판이 표시됨');
  assert.equal(s.editing, false, `키팁 중 편집 진입: ${JSON.stringify(s)}`);
  assert.equal(s.editor, '', `키팁 문자가 셀 편집기에 남음: ${s.editor}`);
  return s;
};
const action = (key) => actionKeys.add(key);

try {
  await test('전체 등록표: 키 중복·terminal 접두 충돌·실제 명령 연결', async (p) => {
    assert.equal(parsedRemainder, '', '등록표 형식이 바뀌어 일부 항목을 검사하지 못함');
    assert.ok(entries.length > 0); assert.ok(tabs.length > 0);
    assert.deepEqual(duplicates, []); assert.deepEqual(prefixConflicts, []);
    const commands = new Set(await p.evaluate(() => window.tabula.commands()));
    const menus = new Set(await p.evaluate(() => window.tabula.menus()));
    assert.deepEqual(entries.filter((e) => !(e.kind === 'menu' ? menus : commands).has(e.command)), [], '존재하지 않는 명령 또는 메뉴');
    for (const [key, command] of [['wvg', 'toggleGrid'], ['wvh', 'toggleHeaders'], ['wvf', 'toggleFormulaBar']]) {
      assert.equal(entries.find((e) => e.key === key)?.command, command, `${key.toUpperCase()} 등록`);
    }
  });
  await test('모든 등록 키팁의 중간 접두사와 리본 탭을 실제 순서 입력으로 방문', async (p) => {
    for (const prefix of prefixes) {
      await sequence(p, prefix); await path(p, prefix); coverage.checkedPrefixes.push(prefix);
      for (let i = 0; i <= prefix.length; i++) await p.keyboard.press('Escape');
      assert.equal((await state(p)).keytips, false);
    }
  });
  await test('Alt → W → V 표시 계층 → G: 눈금선 전환·재전환·입력 오염 없음', async (p) => {
    const initialStatus = (await state(p)).status;
    await sequence(p, 'w'); assert.equal((await path(p, 'w')).tab, '보기');
    await sequence(p, 'v', null); await path(p, 'wv');
    assert.equal((await state(p)).status, initialStatus, '키팁이 기존 상태줄을 긴 키 목록으로 바꿈');
    for (const [key, label] of [['wvg', '눈금선'], ['wvh', '머리글'], ['wvf', '수식 입력줄']]) {
      const badge = p.locator(`.keytip-badge[data-keytip-path="${key}"]`);
      assert.equal(await badge.isVisible(), true);
      assert.ok((await badge.getAttribute('aria-label')).includes(label));
    }
    await p.keyboard.press('g');
    let s = await state(p); assert.equal(s.gridHidden, true); assert.equal(s.keytips, false); assert.equal(s.raw, '10'); assert.equal(s.editing, false);
    await sequence(p, 'wvg'); s = await state(p); assert.equal(s.gridHidden, false); assert.equal(s.editor, ''); action('wvg');
  });
  await test('Alt를 누른 채 W 시작 후 V·G 순차 입력', async (p) => {
    await p.keyboard.down('Alt'); await p.keyboard.press('w'); await p.keyboard.up('Alt');
    await path(p, 'w'); await sequence(p, 'vg', null);
    assert.equal((await state(p)).gridHidden, true); action('wvg');
  });
  await test('F10 시작·F10 종료와 F10 → WVG 실행', async (p) => {
    await p.keyboard.press('F10'); await path(p, ''); await p.keyboard.press('F10'); assert.equal((await state(p)).keytips, false);
    await sequence(p, 'wvg', 'F10'); assert.equal((await state(p)).gridHidden, true); action('wvg');
  });
  await test('WVH는 행·열 머리글, WVF는 수식 입력줄 표시를 각각 전환', async (p) => {
    const originalWidth = (await state(p)).headerWidth; assert.ok(originalWidth > 0);
    await sequence(p, 'wvh'); assert.equal((await state(p)).headerWidth, 0);
    await sequence(p, 'wvh'); assert.equal((await state(p)).headerWidth, originalWidth); action('wvh');
    assert.equal(await p.locator('#formulaRow').isVisible(), true);
    await sequence(p, 'wvf'); assert.equal(await p.locator('#formulaRow').isVisible(), false);
    await sequence(p, 'wvf'); assert.equal(await p.locator('#formulaRow').isVisible(), true); action('wvf');
  });
  await test('기존 WG 눈금선 별칭도 유지', async (p) => {
    await sequence(p, 'wg'); assert.equal((await state(p)).gridHidden, true); action('wg');
  });
  await test('Escape는 WV → W → 루트 → 종료로 한 단계씩 복귀', async (p) => {
    await sequence(p, 'wv'); await path(p, 'wv');
    for (const prefix of ['w', '']) { await p.keyboard.press('Escape'); await path(p, prefix); }
    await p.keyboard.press('Escape'); const s = await state(p); assert.equal(s.keytips, false); assert.equal(s.raw, '10'); assert.equal(s.editing, false);
  });
  await test('Backspace는 한 단계 복귀하고 루트에서 셀 내용을 지우지 않음', async (p) => {
    await sequence(p, 'wv');
    for (const prefix of ['w', '', '']) { await p.keyboard.press('Backspace'); await path(p, prefix); }
    assert.equal((await state(p)).raw, '10'); await sequence(p, 'wvg', null); assert.equal((await state(p)).gridHidden, true); action('wvg');
  });
  await test('Alt를 다시 누르면 키팁 종료하고 이후 타이핑에 문자가 섞이지 않음', async (p) => {
    await sequence(p, 'wv'); await p.keyboard.press('Alt'); assert.equal((await state(p)).keytips, false);
    await p.keyboard.type('abc'); await p.keyboard.press('Enter');
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 1, 1)), 'abc');
  });
  await test('잘못된 키는 현재 키팁 경로에 머물며 후속 G만 실행', async (p) => {
    await sequence(p, 'wv'); await p.keyboard.press('x'); await path(p, 'wv');
    await p.keyboard.press('g'); const s = await state(p); assert.equal(s.gridHidden, true); assert.equal(s.raw, '10'); assert.equal(s.editing, false); action('wvg');
  });
  await test('한글 key 값도 물리 KeyW·KeyV·KeyG로 눈금선 전환', async (p) => {
    await p.keyboard.press('Alt');
    const prevented = await p.evaluate(() => {
      const ed = document.getElementById('cellEditor');
      return [['ㅈ', 'KeyW'], ['ㅍ', 'KeyV'], ['ㅎ', 'KeyG']].map(([key, code]) => {
        const event = new KeyboardEvent('keydown', { key, code, bubbles: true, cancelable: true });
        ed.dispatchEvent(event); return event.defaultPrevented;
      });
    });
    assert.deepEqual(prevented, [true, true, true]);
    const s = await state(p); assert.equal(s.gridHidden, true); assert.equal(s.raw, '10'); assert.equal(s.editor, ''); action('wvg');
  });
  await test('AltGr와 Ctrl+Alt 문자는 키팁이 가로채지 않음', async (p) => {
    const prevented = await p.evaluate(() => {
      const ed = document.getElementById('cellEditor');
      return [
        { key: 'AltGraph', code: 'AltRight', altKey: true, ctrlKey: true, modifierAltGraph: true },
        { key: 'ł', code: 'KeyW', altKey: true, ctrlKey: true, modifierAltGraph: true },
        { key: 'ł', code: 'KeyW', altKey: true, modifierAltGraph: true },
      ].map((init) => { const e = new KeyboardEvent('keydown', { ...init, bubbles: true, cancelable: true }); ed.dispatchEvent(e); return e.defaultPrevented; });
    });
    assert.deepEqual(prevented, [false, false, false]);
    const s = await state(p); assert.equal(s.keytips, false); assert.equal(s.raw, '10'); assert.equal(s.editing, false);
  });
  await test('일반 WVG 타이핑과 셀 편집 중 Alt는 키팁으로 오인하지 않음', async (p) => {
    await p.keyboard.type('wvg'); await p.keyboard.press('Enter');
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 1, 1)), 'wvg');
    await p.keyboard.press('F2'); await p.keyboard.press('Alt'); assert.equal((await state(p)).keytips, false);
  });
  await test('홈 키팁 HAC·HW·HEF가 맞춤·줄 바꿈·서식 지우기를 실제 적용', async (p) => {
    await p.evaluate(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setStyle(0, 1, 1, { align: 'left', bold: true })); });
    await sequence(p, 'hac'); assert.equal(await p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1).align), 'center'); action('hac');
    await sequence(p, 'hw'); assert.equal(await p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1).wrap), true); action('hw');
    await sequence(p, 'hef'); assert.equal(await p.evaluate(() => window.tabula.wb().getCell(0, 1, 1).style ?? null), null); action('hef');
  });
  await test('데이터 AT 필터와 보기 WFR·WFC 틀 고정의 실제 상태 확인', async (p) => {
    await sequence(p, 'at'); assert.equal(await p.evaluate(() => !!window.tabula.wb().sheets[0].filter), true); action('at');
    await sequence(p, 'wfr'); assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].freeze.rows), 1); action('wfr');
    await sequence(p, 'wfc'); assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].freeze.cols), 1); action('wfc');
  });
  await test('HBA·HBN 테두리 하위 명령이 셀 테두리를 추가·제거', async (p) => {
    await sequence(p, 'hb'); await path(p, 'hb');
    await p.keyboard.press('a');
    const before = await p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1));
    assert.ok(before.bt && before.bb && before.bl && before.br); action('hba');
    await sequence(p, 'hbn'); const after = await p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1));
    assert.ok(!after.bt && !after.bb && !after.bl && !after.br); action('hbn');
  });
  await test('추가 테두리 HBS·HBO·HBP·HBL·HBR·HBB·HBT의 방향과 선 종류', async (p) => {
    const cases = [['hbs', ['bt', 'bb', 'bl', 'br']], ['hbo', ['bb']], ['hbp', ['bt']], ['hbl', ['bl']], ['hbr', ['br']], ['hbb', ['bb'], 'double'], ['hbt', ['bt', 'bb', 'bl', 'br'], 'medium']];
    for (const [key, edges, line] of cases) {
      await p.evaluate(() => { const wb = window.tabula.wb(); wb.transact(() => wb.setCellData(0, 1, 1, { raw: '10' })); });
      await sequence(p, key);
      const st = await p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1));
      for (const edge of ['bt', 'bb', 'bl', 'br']) {
        assert.equal(!!st[edge], edges.includes(edge), `${key} ${edge}`);
        if (line && edges.includes(edge)) assert.equal(st[`${edge}s`], line, `${key} ${edge} 선 종류`);
      }
      assert.equal((await state(p)).raw, '10'); action(key);
    }
  });
  await test('HEA·HEC·HEM·HEH가 전체·내용·메모·링크만 정확히 지움', async (p) => {
    for (const key of ['hea', 'hec', 'hem', 'heh']) {
      await p.evaluate(() => {
        const wb = window.tabula.wb();
        wb.transact(() => { wb.setCellData(0, 1, 1, { raw: '00123', style: { numFmt: 'text', bold: true }, comment: '메모', link: 'https://example.com' }); wb.setStyle(0, 1, 1, { numFmt: 'general' }); });
      });
      const before = await p.evaluate(() => window.tabula.wb().serialize().sheets[0].cells['1,1']);
      await sequence(p, key);
      const after = await p.evaluate(() => window.tabula.wb().serialize().sheets[0].cells['1,1'] ?? null);
      if (key === 'hea') assert.equal(after, null);
      else if (key === 'hec') assert.deepEqual(after, { raw: '', style: before.style, comment: before.comment });
      else { const expected = { ...before }; delete expected[key === 'hem' ? 'comment' : 'link']; assert.deepEqual(after, expected); }
      await p.keyboard.press('Control+z');
      assert.deepEqual(await p.evaluate(() => window.tabula.wb().serialize().sheets[0].cells['1,1']), before); action(key);
    }
  });
  await test('HMU는 병합을 해제하며 보호된 시트에서는 거절', async (p) => {
    await p.evaluate(() => { const t = window.tabula, wb = t.wb(); wb.transact(() => wb.merge(0, 1, 1, 1, 2)); t.selectRange({ r1: 1, c1: 1, r2: 1, c2: 2 }); });
    await sequence(p, 'hmu'); assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].merges.length), 0); action('hmu');
    await p.evaluate(() => { const wb = window.tabula.wb(); wb.transact(() => { wb.merge(0, 1, 1, 1, 2); wb.setSheetProp(0, 'protect', { on: true, allow: { selectLocked: true, selectUnlocked: true } }); }); });
    await sequence(p, 'hmu');
    assert.equal(await p.evaluate(() => window.tabula.wb().sheets[0].merges.length), 1);
    assert.match(await p.locator('.dialog').textContent(), /보호/);
  });
  await test('WJ는 100%로, WI는 선택 범위가 보이는 배율로 변경', async (p) => {
    await p.evaluate(() => window.tabula.run('zoomOut'));
    assert.notEqual(await p.locator('#zoomLabel').textContent(), '100%');
    await sequence(p, 'wj'); assert.equal(await p.locator('#zoomLabel').textContent(), '100%'); action('wj');
    await p.evaluate(() => window.tabula.selectRange({ r1: 0, c1: 0, r2: 39, c2: 25 }));
    await sequence(p, 'wi');
    const zoom = Number(await p.locator('#zoomSlider').inputValue());
    assert.ok(zoom >= 25 && zoom < 100, `선택 범위 배율 ${zoom}`); action('wi');
  });
  await test('PSP·RS·WM은 페이지 설정·맞춤법·매크로 대화상자를 열기', async (p) => {
    for (const [key, title] of [['psp', '페이지 설정'], ['rs', '맞춤법 검사'], ['wm', '매크로']]) {
      await sequence(p, key);
      assert.ok((await p.locator('.dialog').textContent()).includes(title), `${key}: ${title}`);
      assert.equal((await state(p)).keytips, false); action(key);
      await p.keyboard.press('Escape');
    }
  });
  await test('320×240 화면: 큰 안내판 없이 작은 배지·하위 메뉴·단계 복귀가 작동', async (p) => {
    await p.setViewportSize({ width: 320, height: 240 });
    await p.evaluate(() => window.tabula.selectCell(1, 1));
    const initialStatus = (await state(p)).status;
    const fits = async (locator) => {
      const box = await locator.boundingBox();
      assert.ok(box && box.x >= 0 && box.y >= 0 && box.x + box.width <= 320.5 && box.y + box.height <= 240.5, `화면 밖 키팁: ${JSON.stringify(box)}`);
      return box;
    };
    await p.keyboard.press('Alt'); await path(p, '');
    assert.ok(await p.locator('.ribbon-tab[data-keytip]').count(), '탭의 작은 키 배지가 없음');
    await sequence(p, 'wv', null); await path(p, 'wv');
    await p.locator('#ribbon [data-ribbon-command="toggleGrid"]').scrollIntoViewIfNeeded();
    const badge = p.locator('.keytip-badge[data-keytip-path="wvg"]');
    await badge.waitFor({ state: 'visible' }); await fits(badge);
    assert.equal(await badge.innerText(), 'G');
    assert.equal((await state(p)).status, initialStatus);
    await p.keyboard.press('Alt'); await sequence(p, 'hb'); await path(p, 'hb');
    const menu = p.locator('.keytip-command-menu[data-keytip-menu="hb"]');
    await fits(menu);
    const lastChoice = menu.locator('.menu-item').last();
    await lastChoice.scrollIntoViewIfNeeded(); await fits(lastChoice);
    assert.ok(await lastChoice.locator('.mi-key').innerText(), '하위 메뉴의 작은 키가 없음');
    await p.keyboard.press('Escape'); await path(p, 'h');
    assert.equal(await p.locator('.keytip-command-menu').count(), 0);
  });
  await test('데스크톱 모드 QAT 12개와 사용자 지정 버튼은 좁은 화면에서 줄바꿈하며 넓은 화면의 키팁과 겹치지 않음', async (p) => {
    for (const height of [240, 480]) {
      await p.setViewportSize({ width: 320, height });
      if (await p.evaluate(() => window.tabula.mobile().active)) await p.locator('#mobileModeToggle').click();
      assert.equal(await p.evaluate(() => window.tabula.mobile().active), false, 'QAT는 데스크톱 모드에서 검사');
      const layout = await p.locator('#quickAccess').evaluate((bar) => ({
        below: bar.classList.contains('below'), commands: bar.querySelectorAll('[data-qat-cmd]').length,
        buttons: [...bar.querySelectorAll('button')].map((button) => {
          const rect = button.getBoundingClientRect();
          return { command: button.dataset.qatCmd || 'customize', left: rect.left, right: rect.right, width: rect.width };
        }),
      }));
      assert.equal(layout.below, true); assert.equal(layout.commands, 12); assert.equal(layout.buttons.length, 13);
      assert.ok(layout.buttons.some((b) => b.command === 'customize'));
      for (const button of layout.buttons) assert.ok(button.width > 0 && button.left >= 0 && button.right <= 320.5, `320×${height} 화면 밖 QAT: ${JSON.stringify(button)}`);
    }
    await p.setViewportSize({ width: 1440, height: 1000 });
    await p.evaluate(() => window.tabula.selectCell(1, 1));
    await sequence(p, 'wv');
    await path(p, 'wv');
    const qat = await p.locator('#quickAccess').boundingBox();
    assert.ok(qat);
    const badges = p.locator('.keytip-badge'); assert.ok(await badges.count());
    for (const badge of await badges.all()) {
      const b = await badge.boundingBox();
      assert.ok(b && (b.x + b.width <= qat.x || b.x >= qat.x + qat.width || b.y + b.height <= qat.y || b.y >= qat.y + qat.height), `작은 키팁이 QAT를 가림: ${JSON.stringify({ qat, b })}`);
    }
  });
  await test('홈 실제 버튼 옆 B·M·AC 배지와 가로 스크롤 뒤 E 배지', async (p) => {
    await sequence(p, 'h');
    for (const [key, anchor] of [['hb', 'borderLast'], ['hm', 'mergeCenter'], ['hac', 'alignCenter']]) {
      const badge = p.locator(`.keytip-badge[data-keytip-path="${key}"]`); assert.equal(await badge.isVisible(), true);
      assert.equal(await badge.getAttribute('data-keytip-anchor'), anchor);
      const b = await badge.boundingBox(), a = await p.locator(`#ribbon [data-ribbon-command="${anchor}"]`).boundingBox();
      assert.ok(b && a && Math.abs((b.x + b.width / 2) - (a.x + a.width / 2)) < 15 && Math.abs(b.y - (a.y + a.height)) < 10, `버튼에 붙지 않은 배지 ${key}`);
    }
    await p.screenshot({ path: process.env.WIXEL_KEYTIP_HOME_SCREENSHOT || 'D:/Codex/Temp/wixel3-keytips-home.png' });
    await p.locator('#ribbon [data-ribbon-menu="clear"]').scrollIntoViewIfNeeded();
    await p.locator('.keytip-badge[data-keytip-path="he"]').waitFor({ state: 'visible' });
    assert.equal(await p.locator('.keytip-badge[data-keytip-path="he"]').innerText(), 'E');
  });
  await test('WV 실제 눈금선·머리글·수식 입력줄 옆 G·H·F 배지 클릭', async (p) => {
    await sequence(p, 'wv');
    for (const [key, command] of [['wvg', 'toggleGrid'], ['wvh', 'toggleHeaders'], ['wvf', 'toggleFormulaBar']]) {
      const b = p.locator(`.keytip-badge[data-keytip-path="${key}"]`); assert.equal(await b.isVisible(), true); assert.equal(await b.getAttribute('data-keytip-anchor'), command);
    }
    await p.screenshot({ path: process.env.WIXEL_KEYTIP_VIEW_SCREENSHOT || 'D:/Codex/Temp/wixel3-keytips-view.png' });
    await p.locator('.keytip-badge[data-keytip-path="wvg"]').click(); assert.equal((await state(p)).gridHidden, true); assert.equal((await state(p)).keytips, false);
  });
  await test('HB는 테두리 버튼 아래 메뉴·행별 작은 키·클릭 실행', async (p) => {
    await sequence(p, 'h'); await p.locator('.keytip-badge[data-keytip-path="hb"]').click(); await path(p, 'hb');
    const menu = p.locator('.keytip-command-menu[data-keytip-menu="hb"]'); assert.equal(await menu.isVisible(), true);
    assert.equal(await menu.locator('[data-keytip-path="hba"] .mi-key').innerText(), 'A');
    const m = await menu.boundingBox(), a = await p.locator('#ribbon [data-ribbon-menu="borders"]').boundingBox();
    assert.ok(m && a && Math.abs(m.x - a.x) < 5 && Math.abs(m.y - (a.y + a.height + 2)) < 5);
    await p.screenshot({ path: process.env.WIXEL_KEYTIP_MENU_SCREENSHOT || 'D:/Codex/Temp/wixel3-keytips-borders.png' });
    await menu.locator('[data-keytip-path="hba"]').click(); const s = await p.evaluate(() => window.tabula.wb().styleAt(0, 1, 1)); assert.ok(s.bt && s.bb && s.bl && s.br); assert.equal((await state(p)).keytips, false);
  });
  await test('HE 지우기 메뉴 Esc 단계 복귀·다시 진입 후 F 실행', async (p) => {
    await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setStyle(0, 1, 1, { bold: true })); });
    await sequence(p, 'he'); assert.equal(await p.locator('.keytip-command-menu[data-keytip-menu="he"]').isVisible(), true);
    await p.keyboard.press('Escape'); await path(p, 'h'); assert.equal(await p.locator('.keytip-command-menu').count(), 0);
    await p.keyboard.press('e'); await path(p, 'he'); await p.keyboard.press('f');
    assert.equal(await p.evaluate(() => !!window.tabula.wb().styleAt(0, 1, 1).bold), false); assert.equal((await state(p)).raw, '10');
  });
  await test('키팁 하위 메뉴 방향키·Enter와 QAT 눈금선 SVG 아이콘', async (p) => {
    assert.equal(await p.locator('#quickAccess [data-qat-cmd="toggleGrid"] svg').count(), 1);
    await sequence(p, 'hb'); await p.keyboard.press('ArrowDown'); await p.keyboard.press('Enter');
    assert.equal((await state(p)).keytips, false); assert.equal(await p.locator('.keytip-command-menu').count(), 0);
    assert.equal((await state(p)).raw, '10');
  });
  await test('홈 H1·H2·H3 실제 굵게·기울임꼴·밑줄 배지와 적용·해제', async (p) => {
    for (const [key, prop] of [['h1', 'bold'], ['h2', 'italic'], ['h3', 'underline']]) {
      await sequence(p, 'h'); const badge = p.locator(`.keytip-badge[data-keytip-path="${key}"]`);
      assert.equal(await badge.isVisible(), true); assert.equal(await badge.getAttribute('data-keytip-anchor'), prop);
      await p.keyboard.press(key.slice(1)); assert.equal(await p.evaluate((prop) => window.tabula.wb().styleAt(0, 1, 1)[prop], prop), true);
      await sequence(p, key); assert.equal(await p.evaluate((prop) => !!window.tabula.wb().styleAt(0, 1, 1)[prop], prop), false); action(key);
    }
  });
  await test('홈 HC 복사 배지와 실제 Ctrl+V 붙여넣기·원본 유지', async (p) => {
    await sequence(p, 'h'); assert.equal(await p.locator('.keytip-badge[data-keytip-path="hc"]').getAttribute('data-keytip-anchor'), 'copy');
    await p.keyboard.press('c'); await p.evaluate(() => window.tabula.selectCell(1, 2)); await p.keyboard.press('Control+v');
    await p.waitForFunction(() => window.tabula.wb().getValue(0, 1, 2) === 10);
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 1, 1)), 10); action('hc');
  });
  await test('홈 HX 잘라내기 배지와 실제 붙여넣기·원본 이동', async (p) => {
    await sequence(p, 'h'); assert.equal(await p.locator('.keytip-badge[data-keytip-path="hx"]').getAttribute('data-keytip-anchor'), 'cut');
    await p.keyboard.press('x'); await p.evaluate(() => window.tabula.selectCell(1, 3)); await p.keyboard.press('Control+v');
    await p.waitForFunction(() => window.tabula.wb().getValue(0, 1, 3) === 10);
    assert.equal(await p.evaluate(() => window.tabula.wb().getValue(0, 1, 1)), null); action('hx');
  });
  coverage.verifiedActionKeys = [...actionKeys].sort();
  coverage.checkedPrefixCount = coverage.checkedPrefixes.length;
  coverage.verifiedActionCount = coverage.verifiedActionKeys.length;
  const failed = results.filter((r) => !r.ok).length;
  console.log(JSON.stringify({ url, tests: results.length, failed, coverage, results }, null, 2));
  if (failed) process.exitCode = 1;
} finally { await browser.close(); }
