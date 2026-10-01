// 온라인 그림 검색: 외부 API/이미지는 모두 합성 응답, 사용자 문서/원격 쓰기 접근 없음.
// WIXEL_URL(소스/번들), WIXEL_ONLINE_PICTURES_FILTER, WIXEL_ONLINE_PICTURES_SCREENSHOT 선택 가능.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
const origin = new URL(url).origin, results = [], browser = await chromium.launch();
const filter = process.env.WIXEL_ONLINE_PICTURES_FILTER;
const diagnostics = { pageErrors: [], blockedWrites: [], unexpectedReads: [] };
const hosts = { 'api.openverse.org': 'openverse', 'commons.wikimedia.org': 'wikimedia', 'api.inaturalist.org': 'inaturalist', 'images-api.nasa.gov': 'nasa' };
const imageURL = (id, size = 'full') => `https://fixture.test/${encodeURIComponent(id)}/${size}.png`;
const ccURL = 'https://creativecommons.org/licenses/by/4.0/';
function response(source, q, page, empty = false) {
  const ov = (id, license = 'by') => ({ id, title: id, url: imageURL(id), thumbnail: imageURL(id, 'thumb'), creator: '합성 작가', license, license_url: license === 'by' ? ccURL : 'https://creativecommons.org/publicdomain/mark/1.0/', foreign_landing_url: `https://fixture.test/source/${encodeURIComponent(id)}` });
  const wm = (id, index) => ({ pageid: index, index, title: `File:${id}`, imageinfo: [{ url: imageURL(id), thumburl: imageURL(id, 'thumb'), descriptionurl: `https://fixture.test/source/${encodeURIComponent(id)}`, extmetadata: { Artist: { value: '<b>합성 작가</b>' }, License: { value: 'cc-by' }, LicenseShortName: { value: 'CC BY 4.0' }, LicenseUrl: { value: ccURL } } }] });
  if (source === 'openverse') return { page_count: 2, results: empty ? [] : page === 1 ? [ov(q + ' 공개'), ov(q + ' 공공영역', 'pdm')] : [ov(q + ' 공개'), ov(q + ' 추가')] };
  if (source === 'wikimedia') return empty ? {} : { query: { pages: { 1: wm(q + ' 공용', 1), ...(page > 1 ? { 2: wm(q + ' 공용 추가', 2) } : {}) } }, ...(page === 1 ? { continue: { gsroffset: 30 } } : {}) };
  if (source === 'inaturalist') return { total_results: 1, results: empty ? [] : [{ taxon: { preferred_common_name: q + ' 자연' }, photos: [{ id: 101, url: imageURL(q + ' 자연', 'square'), original_url: imageURL(q + ' 자연'), license_code: 'cc-by', attribution: '합성 자연 작가' }] }] };
  return { collection: { metadata: { total_hits: 1 }, items: empty ? [] : [{ data: [{ nasa_id: q + '-nasa', media_type: 'image', title: q + ' 우주', photographer: '합성 우주 작가', center: 'NASA' }], links: [{ href: imageURL(q + ' 우주', 'thumb'), render: 'image', rel: 'preview' }, { href: imageURL(q + ' 우주'), render: 'image', rel: 'canonical' }] }] } };
}
async function until(fn, message = '비동기 조건') {
  const deadline = Date.now() + 10000;
  while (!fn()) { if (Date.now() > deadline) throw Error(message + ' 시간 초과'); await new Promise(r => setTimeout(r, 10)); }
}
function hold(net, predicate) {
  let release, seen, done;
  const gate = { predicate, pending: new Promise(r => { release = r; }), started: new Promise(r => { seen = r; }), finished: new Promise(r => { done = r; }), release, seen, done };
  net.gates.push(gate); return gate;
}
async function test(name, fn) {
  if (filter && !name.includes(filter)) return;
  const context = await browser.newContext({ viewport: { width: 1400, height: 1000 }, acceptDownloads: false });
  const p = await context.newPage(), errors = [], writes = [], unexpected = [];
  const net = { requests: [], images: [], failed: new Set(), empty: false, gates: [], imageFail: null, png: null };
  p.setDefaultTimeout(10000); p.on('pageerror', e => errors.push(e.message)); p.on('dialog', d => d.dismiss());
  await context.route('**/*', async route => {
    const req = route.request(), u = new URL(req.url());
    if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + u.href); return route.abort(); }
    const source = hosts[u.hostname];
    if (source || u.hostname === 'fixture.test') {
      const entry = source ? { kind: 'api', source, q: u.searchParams.get(source === 'wikimedia' ? 'gsrsearch' : 'q'), page: source === 'wikimedia' ? Number(u.searchParams.get('gsroffset') || 0) / 30 + 1 : Number(u.searchParams.get('page') || 1), url: u.href } : { kind: 'image', url: u.href };
      (source ? net.requests : net.images).push(entry);
      const gate = net.gates.find(g => !g.used && g.predicate(entry)); if (gate) { gate.used = true; gate.seen(); await gate.pending; }
      try {
        if (source) return await route.fulfill({ status: net.failed.has(source) ? 503 : 200, json: net.failed.has(source) ? { error: '합성 검색 실패' } : response(source, entry.q, entry.page, net.empty), headers: { 'access-control-allow-origin': '*' } });
        if (u.pathname.startsWith('/source/')) { unexpected.push('출처 링크 자동 탐색: ' + u.href); return await route.abort(); }
        const broken = net.imageFail?.(entry);
        return await route.fulfill({ status: broken ? 500 : 200, contentType: broken ? 'text/plain' : 'image/png', body: broken ? '합성 이미지 오류' : net.png, headers: { 'access-control-allow-origin': '*' } });
      } catch (e) { if (!/closed|canceled|already handled|aborted/i.test(e.message)) throw e; }
      finally { gate?.done(); }
    }
    if (u.origin === origin && !u.pathname.startsWith('/api/')) return route.continue();
    // 앱 상태 확인도 실제 보관함/문서를 조회하지 않는다.
    if (u.origin === origin && u.pathname === '/api/health') return route.fulfill({ status: 503, json: { ok: false } });
    unexpected.push(u.href); return route.abort();
  });
  try {
    await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
    await p.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await p.waitForFunction(() => window.tabula?.wb(), null, { timeout: 60000 });
    const png = await p.evaluate(() => { const c = document.createElement('canvas'); c.width = 120; c.height = 60; const x = c.getContext('2d'); x.fillStyle = '#1570a0'; x.fillRect(0, 0, 120, 60); x.fillStyle = '#ffcc00'; x.fillRect(50, 10, 30, 40); return c.toDataURL(); }); net.png = Buffer.from(png.split(',')[1], 'base64');
    await fixture(p); await fn(p, net);
    assert.deepEqual(errors, [], '페이지 오류'); assert.deepEqual(writes, [], '원격 쓰기'); assert.deepEqual(unexpected, [], '모의 처리 밖 외부 읽기');
    results.push({ name, ok: true }); console.log('OK ' + name);
  } catch (e) { results.push({ name, ok: false, error: e.message, pageErrors: errors, remoteWrites: writes, unexpectedReads: unexpected }); console.error('NG ' + name + ': ' + e.stack); }
  finally { diagnostics.pageErrors.push(...errors); diagnostics.blockedWrites.push(...writes); diagnostics.unexpectedReads.push(...unexpected); for (const g of net.gates) g.release(); await context.close(); }
}
async function fixture(p) {
  await p.evaluate(() => { const t = window.tabula, w = t.wb(); w.restore({ sheets: [{ name: '그림 합성', cells: { '0,0': { raw: '보존' }, '0,1': { raw: '=1+1', cached: 42 }, '2,2': { raw: '교체 셀', style: { fill: '#eeeeee', bold: true } }, '3,2': { raw: '둘째 셀' } } }, { name: '다른 합성', cells: { '0,0': { raw: '다른 시트 보존' } } }] }); t.switchSheet(0); t.selectCell(2, 2); w.undoStack = []; w.redoStack = []; });
}
const dialog = p => p.getByRole('dialog', { name: '온라인 그림', exact: true });
async function open(p, inCell = false) { await p.evaluate(() => window.tabula.openNamedMenu('picture', { x: 120, y: 120 })); await p.getByRole('menuitem', { name: inCell ? '온라인 그림을 셀에 배치...' : '온라인 그림...', exact: true }).click(); await dialog(p).waitFor(); return dialog(p); }
async function search(p, q = '합성', source) { const d = dialog(p); if (source) await d.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption(source); await d.getByRole('searchbox', { name: '그림 검색어', exact: true }).fill(q); await d.getByRole('button', { name: '검색', exact: true }).click(); await settled(p); }
const settled = async p => { await p.waitForFunction(() => { const grid = document.querySelector('.online-picture-dialog .online-grid'); return grid && !grid.hasAttribute('aria-busy') && !document.querySelector('.online-status')?.textContent.includes('검색 중'); }); };
const titles = p => dialog(p).locator('.online-item').evaluateAll(es => es.map(e => e.getAttribute('aria-label')));
const saved = p => p.evaluate(() => JSON.stringify(window.tabula.wb().serialize()));
const images = p => p.evaluate(() => structuredClone(window.tabula.wb().sheets.map(s => s.images ?? [])));
const undo = p => p.evaluate(() => window.tabula.run('undo'));
const redo = p => p.evaluate(() => window.tabula.run('redo'));
const depth = p => p.evaluate(() => window.tabula.wb().undoStack.length);
const pick = (p, name) => dialog(p).getByRole('button', { name, exact: true }).click();
const insert = p => dialog(p).getByRole('button', { name: '삽입', exact: true }).click();
try {
  await test('기본 CC 해제·다섯 사이트 선택·빈 검색 안내', async (p, net) => {
    const d = await open(p); assert.equal(await d.getByRole('checkbox', { name: 'Creative Commons만', exact: true }).isChecked(), false);
    assert.deepEqual(await d.getByRole('combobox', { name: '검색 사이트', exact: true }).locator('option').evaluateAll(es => es.map(e => e.value)), ['all', 'openverse', 'wikimedia', 'inaturalist', 'nasa']);
    await d.getByRole('button', { name: '검색', exact: true }).click(); assert.match(await d.getByRole('status').textContent(), /검색어를 입력/); assert.equal(net.requests.length, 0); assert.equal(await d.getByRole('button', { name: '더 보기', exact: true }).isDisabled(), true);
  });
  await test('네 출처 통합·라이선스·원문 링크·원본/미리보기 주소 구분', async (p, net) => {
    const d = await open(p); await search(p); assert.equal((await titles(p)).length, 5); assert.deepEqual(net.requests.map(r => r.source).sort(), ['inaturalist', 'nasa', 'openverse', 'wikimedia']);
    assert.equal(await d.locator('.online-license').count(), 5); assert.equal(await d.locator('.online-source').count(), 5);
    const links = await d.locator('.online-card a').evaluateAll(es => es.map(e => ({ href: e.href, target: e.target, rel: e.rel, label: e.textContent })));
    assert.equal(links.length, 5); for (const a of links) { assert.match(a.href, /^https:\/\//); assert.equal(a.target, '_blank'); assert.match(a.rel, /noopener/); assert.match(a.rel, /noreferrer/); assert.equal(a.label, '출처 보기'); }
    assert.equal(net.images.some(r => r.url.endsWith('/full.png')), false, '검색만으로 원본을 내려받지 않음');
    await p.screenshot({ path: process.env.WIXEL_ONLINE_PICTURES_SCREENSHOT || 'D:/Codex/Temp/wixel-online-pictures.png' });
  });
  await test('CC 체크 즉시 재검색·개별 라이선스 확인·NASA 제외', async (p, net) => {
    const d = await open(p); await search(p); const before = net.requests.length; await d.getByRole('checkbox', { name: 'Creative Commons만', exact: true }).check(); await settled(p);
    assert.deepEqual(await titles(p), ['합성 공개', '합성 공용', '합성 자연']); const next = net.requests.slice(before); assert.equal(next.length, 3); assert.equal(next.some(r => r.source === 'nasa'), false);
    assert.match(new URL(next.find(r => r.source === 'openverse').url).searchParams.get('license'), /by-sa/); assert.match(new URL(next.find(r => r.source === 'inaturalist').url).searchParams.get('photo_license'), /cc-by/);
    await d.getByRole('checkbox', { name: 'Creative Commons만', exact: true }).uncheck(); await settled(p); assert.equal((await titles(p)).length, 5);
  });
  await test('사이트 변경 자동 재검색·CC 전용 NASA 빈 결과 안내', async (p, net) => {
    const d = await open(p); await search(p); let before = net.requests.length; await d.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption('wikimedia'); await settled(p); assert.deepEqual(net.requests.slice(before).map(r => r.source), ['wikimedia']); assert.deepEqual(await titles(p), ['합성 공용']);
    await d.getByRole('checkbox', { name: 'Creative Commons만', exact: true }).check(); await settled(p); before = net.requests.length; await d.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption('nasa'); await settled(p); assert.deepEqual(await titles(p), []); assert.equal(net.requests.length, before); assert.match(await d.getByRole('status').textContent(), /NASA.*CC/);
  });
  await test('검색어를 지운 뒤 필터 변경은 기존 결과·선택을 모두 초기화', async (p, net) => {
    const d = await open(p); await search(p); await pick(p, '합성 공공영역'); const before = net.requests.length;
    await d.getByRole('searchbox', { name: '그림 검색어', exact: true }).fill(''); await d.getByRole('checkbox', { name: 'Creative Commons만', exact: true }).check();
    assert.deepEqual(await titles(p), []); assert.equal(await d.getByRole('button', { name: '더 보기', exact: true }).isDisabled(), true); assert.equal(net.requests.length, before);
    await d.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption('wikimedia'); assert.deepEqual(await titles(p), []); assert.equal(net.requests.length, before);
    await insert(p); assert.match(await d.getByRole('status').textContent(), /그림을 선택/); assert.deepEqual(await images(p), [[], []]);
  });
  await test('새 검색 뒤 늦게 도착한 이전 응답 폐기', async (p, net) => {
    const d = await open(p); await d.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption('openverse'); const g = hold(net, r => r.kind === 'api' && r.q === '이전');
    const q = d.getByRole('searchbox', { name: '그림 검색어', exact: true }); await q.fill('이전'); await q.press('Enter'); await g.started; await q.fill('최신'); await q.press('Enter'); await settled(p); const expected = ['최신 공개', '최신 공공영역']; assert.deepEqual(await titles(p), expected); g.release(); await g.finished; await settled(p); assert.deepEqual(await titles(p), expected);
  });
  await test('더 보기 중복 제외·기존 선택 보존·마지막 페이지 비활성화', async (p) => {
    const d = await open(p); await search(p); await pick(p, '합성 공개'); await d.getByRole('button', { name: '더 보기', exact: true }).click(); await settled(p);
    const all = await titles(p); assert.equal(all.length, 7); assert.equal(new Set(all).size, all.length); assert.equal(await d.getByRole('button', { name: '합성 공개', exact: true }).getAttribute('aria-pressed'), 'true'); assert.equal(await d.getByRole('button', { name: '더 보기', exact: true }).isDisabled(), true);
  });
  await test('일부 검색 사이트 실패에도 다른 출처 결과 유지·실패 표시', async (p, net) => {
    net.failed = new Set(['openverse', 'nasa']); const d = await open(p); await search(p); assert.deepEqual(await titles(p), ['합성 공용', '합성 자연']); assert.match(await d.getByRole('status').textContent(), /연결 실패.*Openverse.*NASA.*다른 출처/);
  });
  await test('전체 검색 실패 후 재검색 복구·문서/Undo 무변경', async (p, net) => {
    const before = await saved(p); net.failed = new Set(Object.values(hosts)); const d = await open(p); await search(p); assert.match(await d.getByRole('status').textContent(), /검색할 수 없습니다/); assert.deepEqual(await titles(p), []); net.failed.clear(); await search(p); assert.equal((await titles(p)).length, 5); assert.equal(await saved(p), before); assert.equal(await depth(p), 0);
  });
  await test('결과 없는 검색은 이전 카드·선택 제거', async (p, net) => {
    const d = await open(p); await search(p); await pick(p, '합성 공개'); net.empty = true; await search(p, '없음'); assert.deepEqual(await titles(p), []); assert.match(await d.getByRole('status').textContent(), /결과가 없습니다/); await insert(p); assert.match(await d.getByRole('status').textContent(), /그림을 선택/); assert.deepEqual(await images(p), [[], []]);
  });
  await test('여러 그림 원본 삽입·원래 셀 보존·한 번 Undo/Redo', async (p, net) => {
    const before = await saved(p); await open(p); await search(p); await pick(p, '합성 공개'); await pick(p, '합성 공용'); await insert(p); await dialog(p).waitFor({ state: 'detached' });
    const after = await images(p); assert.equal(after[0].length, 2); assert.equal(after[1].length, 0); for (const im of after[0]) { assert.match(im.src, /^data:image\/png;base64,/); assert.deepEqual([im.w, im.h], [120, 60]); }
    assert.deepEqual(net.images.filter(r => r.url.endsWith('/full.png')).map(r => decodeURIComponent(new URL(r.url).pathname)).sort(), ['/합성 공개/full.png', '/합성 공용/full.png']); assert.equal(await depth(p), 1); const snapshot = await saved(p); await undo(p); assert.equal(await saved(p), before); await redo(p); assert.equal(await saved(p), snapshot);
  });
  await test('셀 그림 다중 삽입은 원래 주소/서식 유지·한 번 Undo', async (p) => {
    const before = await saved(p); await open(p, true); await search(p); await pick(p, '합성 공개'); await pick(p, '합성 공용'); await p.evaluate(() => window.tabula.selectCell(8, 8)); await insert(p); await dialog(p).waitFor({ state: 'detached' });
    const cells = await p.evaluate(() => { const w = window.tabula.wb(); return [w.getCell(0, 2, 2), w.getCell(0, 3, 2), w.getCell(0, 8, 8)].map(c => c && { image: c.image, raw: c.raw, style: c.style }); });
    assert.equal(cells[0].image.alt, '합성 공개'); assert.equal(cells[1].image.alt, '합성 공용'); assert.equal(cells[0].style.bold, true); assert.equal(cells[0].style.fill, '#eeeeee'); assert.equal(cells[2], undefined); assert.deepEqual(await images(p), [[], []]); assert.equal(await depth(p), 1); await undo(p); assert.equal(await saved(p), before);
  });
  await test('검색 중 닫기는 늦은 결과/문서 변경 없이 취소', async (p, net) => {
    const before = await saved(p); const d = await open(p); await d.getByRole('combobox', { name: '검색 사이트', exact: true }).selectOption('openverse'); const g = hold(net, r => r.kind === 'api'); await d.getByRole('searchbox', { name: '그림 검색어', exact: true }).fill('늦음'); await d.getByRole('button', { name: '검색', exact: true }).click(); await g.started; await d.getByRole('button', { name: '취소', exact: true }).click(); g.release(); await g.finished; assert.equal(await dialog(p).count(), 0); assert.equal(await saved(p), before); assert.equal(await depth(p), 0); await open(p); assert.deepEqual(await titles(p), []);
  });
  await test('삽입 준비 중 시트 전환은 어느 시트에도 쓰지 않음', async (p, net) => {
    await open(p); await search(p); await pick(p, '합성 공개'); const g = hold(net, r => r.kind === 'image' && r.url.endsWith('/full.png')); await insert(p); await g.started; await p.evaluate(() => window.tabula.switchSheet(1)); g.release(); await g.finished; await dialog(p).getByRole('alert').waitFor(); assert.match(await dialog(p).getByRole('alert').textContent(), /문서 또는 시트/); assert.deepEqual(await images(p), [[], []]); assert.equal(await depth(p), 0);
  });
  await test('삽입 준비 중 통합 문서 교체는 새 문서에 쓰지 않음', async (p, net) => {
    await open(p); await search(p); await pick(p, '합성 공개'); const g = hold(net, r => r.kind === 'image' && r.url.endsWith('/full.png')); await insert(p); await g.started; await fixture(p); const before = await saved(p); g.release(); await g.finished; await dialog(p).getByRole('alert').waitFor(); assert.equal(await saved(p), before); assert.deepEqual(await images(p), [[], []]); assert.equal(await depth(p), 0);
  });
  await test('삽입 중 창/선택 잠금·Escape와 중복 클릭은 중복 삽입하지 않음', async (p, net) => {
    const d = await open(p); await search(p); await pick(p, '합성 공개'); const g = hold(net, r => r.kind === 'image' && r.url.endsWith('/full.png')); await insert(p); await g.started;
    assert.equal(await d.getAttribute('aria-busy'), 'true'); assert.equal(await d.locator('.online-picture-dialog').evaluate(el => el.inert), true); assert.equal(await d.getByRole('button', { name: '삽입', exact: true }).isDisabled(), true); await p.keyboard.press('Escape'); assert.equal(await d.count(), 1); g.release(); await g.finished; await d.waitFor({ state: 'detached' }); assert.equal((await images(p))[0].length, 1); assert.equal(await depth(p), 1);
  });
  await test('두 번째 그림 로드 실패는 전체 삽입 롤백·오류 후 취소', async (p, net) => {
    const before = await saved(p); const d = await open(p); await search(p); await pick(p, '합성 공개'); await pick(p, '합성 공용'); net.imageFail = r => r.url === imageURL('합성 공용'); await insert(p); await d.getByRole('alert').waitFor(); assert.match(await d.getByRole('alert').textContent(), /그림을 불러올 수 없습니다/); assert.equal(await saved(p), before); assert.equal(await depth(p), 0); await d.getByRole('button', { name: '취소', exact: true }).click(); assert.equal(await d.count(), 0);
  });
  await test('직접 주소의 실행 프로토콜·인증 정보 거부', async (p, net) => {
    const d = await open(p); for (const address of ['javascript:alert(1)', 'https://user:secret@fixture.test/private.png']) { await d.getByRole('textbox', { name: '그림 웹 주소', exact: true }).fill(address); await insert(p); await d.getByRole('alert').waitFor(); assert.match(await d.getByRole('alert').textContent(), /http/); } assert.equal(net.images.length, 0); assert.deepEqual(await images(p), [[], []]);
  });
  await test('삽입 준비 중 시트 보호 전환은 변경 차단', async (p, net) => {
    await open(p); await search(p); await pick(p, '합성 공개'); const g = hold(net, r => r.kind === 'image' && r.url.endsWith('/full.png')); await insert(p); await g.started; await p.evaluate(() => { const w = window.tabula.wb(); w.transact(() => w.setSheetProp(0, 'protect', { on: true, allow: { objects: false } })); }); g.release(); await g.finished; await dialog(p).getByRole('alert').waitFor(); assert.match(await dialog(p).getByRole('alert').textContent(), /보호된 위치/); assert.deepEqual(await images(p), [[], []]); assert.equal(await depth(p), 1, '보호 변경만 기록');
  });
} finally { await browser.close(); }
const bad = results.filter(x => !x.ok); console.log(JSON.stringify({ total: results.length, good: results.length - bad.length, bad: bad.length, pageErrors: diagnostics.pageErrors.length, blockedWrites: diagnostics.blockedWrites.length, unexpectedReads: diagnostics.unexpectedReads.length, results }, null, 2)); if (bad.length) process.exitCode = 1;
