import assert from 'node:assert/strict';
import { mkdir, writeFile } from 'node:fs/promises';
import { join } from 'node:path';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://localhost:5178/';
const output = process.env.IMPORTRANGE_SCREENSHOTS;
if (output) await mkdir(output, { recursive: true });
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1600, height: 1000 } });
const page = await context.newPage();
page.setDefaultTimeout(10000);
const results = [], pageErrors = [], writes = [], unexpectedReads = [], fixtures = new Map(), requests = [];
let sequence = 0;
page.on('pageerror', e => pageErrors.push(e.message));
await context.addInitScript(() => { window.TABULA_STATIC = false; window.WIXEL_SKIP_START = true; });
await context.route('**/*', async route => {
  const req = route.request(), url = new URL(req.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(req.method())) { writes.push(req.method() + ' ' + url.pathname); return route.abort(); }
  if (url.origin !== new URL(base).origin) { unexpectedReads.push(url.hostname); return route.abort(); }
  if (url.pathname === '/api/health') return route.fulfill({ json: { ok: true, auth: false, vault: true, publish: false } });
  if (url.pathname === '/api/fetch') {
    let source;
    try { source = new URL(url.searchParams.get('url')); } catch { return route.fulfill({ status: 400, body: '잘못된 합성 요청' }); }
    const id = /^\/spreadsheets\/d\/(?:e\/)?([^/]+)/.exec(source.pathname)?.[1], fixture = fixtures.get(id);
    if (source.hostname !== 'docs.google.com' || !fixture) { unexpectedReads.push('unknown fixture'); return route.fulfill({ status: 404, body: '등록되지 않은 합성 응답' }); }
    const record = { id, path: source.pathname, query: Object.fromEntries(source.searchParams), hash: source.hash };
    requests.push(record);
    if (source.pathname.endsWith('/htmlview')) return route.fulfill({ contentType: 'text/html; charset=utf-8', body: fixture.tabsHtml || '<html></html>' });
    fixture.calls++;
    const body = typeof fixture.body === 'function' ? fixture.body(fixture.calls, source) : fixture.body;
    return route.fulfill({ status: fixture.status || 200, contentType: fixture.contentType || 'text/csv; charset=utf-8', body });
  }
  if (url.pathname.startsWith('/api/')) { unexpectedReads.push(url.pathname); return route.abort(); }
  return route.continue();
});
const csv = rows => rows.map(row => row.map(value => '"' + String(value ?? '').replaceAll('"', '""') + '"').join(',')).join('\r\n');
function fixture(rows, options = {}) {
  const id = 'WIXEL_SYNTHETIC_IMPORT_' + String(++sequence).padStart(4, '0');
  const f = { id, body: Array.isArray(rows) ? csv(rows) : rows, calls: 0, ...options };
  fixtures.set(id, f); return f;
}
const quote = value => '"' + String(value).replaceAll('"', '""') + '"';
const formula = (source, range) => `=IMPORTRANGE(${quote(source)},${quote(range)})`;
const queryFormula = (source, range, query, headers = 1) => `=QUERY(IMPORTRANGE(${quote(source)},${quote(range)}),${quote(query)},${headers})`;
const frame = () => page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
async function reset() {
  for (let i = 0; i < 3; i++) await page.keyboard.press('Escape');
  await page.evaluate(() => {
    const t = tabula, w = t.wb(); w.restore({ sheets: [{ name: '가져오기 합성', defColW: 180, cells: {} }] });
    w.undoStack = []; w.redoStack = []; t.gv().setScroll(0, 0); t.gv().layout(); t.selectCell(0, 0);
    const toggle = document.querySelector('#autosaveToggle'); if (toggle?.checked) toggle.click();
  });
  await frame();
}
async function put(text, r = 0, c = 0) {
  await page.evaluate(({ text, r, c }) => {
    const t = tabula, w = t.wb(); w.transact(() => w.setInput(t.si, r, c, text)); t.gv().renderAll();
  }, { text, r, c });
  await page.waitForFunction(({ r, c }) => {
    const t = tabula, v = t.wb().getValue(t.si, r, c); return v !== '로딩 중…';
  }, { r, c }, { timeout: 10000 });
  await frame();
}
async function matrix(height, width, r = 0, c = 0) {
  return page.evaluate(({ height, width, r, c }) => {
    const t = tabula, w = t.wb(), rows = [], types = [];
    for (let y = 0; y < height; y++) {
      const values = [], tags = [];
      for (let x = 0; x < width; x++) { const v = w.getValue(t.si, r + y, c + x); values.push(v?.code || (v ?? '')); tags.push(typeof v); }
      rows.push(values); types.push(tags);
    }
    return { rows, types, spill: w.spillRange(t.si, r, c) };
  }, { height, width, r, c });
}
async function display(r, c) {
  await frame();
  return page.evaluate(({ r, c }) => {
    const t = tabula, w = t.wb(), n = [...document.querySelectorAll(`.c[data-r="${r}"][data-c="${c}"]`)].find(n => n.getBoundingClientRect().width);
    const v = w.getValue(t.si, r, c); return { value: v?.code || v, type: typeof v, text: n?.textContent, fontSize: n ? parseFloat(getComputedStyle(n).fontSize) : null, style: w.styleAt(t.si, r, c) };
  }, { r, c });
}
function sourceRequest(f) { const hit = requests.find(r => r.id === f.id && !r.path.endsWith('/htmlview')); assert.ok(hit, '합성 API 요청'); return hit; }
async function test(name, fn) {
  try { await reset(); await fn(); results.push({ name, ok: true }); console.log('OK ' + name); }
  catch (e) { results.push({ name, ok: false, error: e.message }); console.error('NG ' + name + ': ' + e.message); }
}
try {
  const health = page.waitForResponse(r => new URL(r.url()).pathname === '/api/health', { timeout: 60000 });
  await page.goto(base, { waitUntil: 'commit', timeout: 60000 }); await health;
  await page.waitForFunction(() => !!window.tabula?.gv(), null, { timeout: 60000 }); await frame();
  await test('IMPORTRANGE는 같은 열의 숫자와 텍스트를 원본 CSV 경로로 보존', async () => {
    const rows = [['구분', '값', '코드'], ['첫째', 17, '00123'], ['둘째', '확인 필요', 'A-17'], ['셋째', 23, '00234']], f = fixture(rows);
    await put(formula(f.id, 'A1:C4'));
    assert.deepEqual((await matrix(4, 3)).rows, rows);
    const req = sourceRequest(f); assert.equal(req.path, `/spreadsheets/d/${f.id}/export`); assert.equal(req.query.format, 'csv');
  });
  await test('앞·중간의 빈 행과 빈 열 및 마지막 빈 열이 셀 좌표를 유지', async () => {
    const rows = [['', '', '', '', ''], ['', '항목', '', '금액', ''], ['', '첫째', '', 10, ''], ['', '', '', '', ''], ['', '셋째', '', 30, '']], f = fixture(rows);
    f.body = ',,,,\r\n,항목,,금액,\r\n,첫째,,10,\r\n\r\n,셋째,,30,';
    await put(formula(f.id, 'A1:E5'));
    const actual = await matrix(5, 5); assert.deepEqual(actual.rows, rows); assert.deepEqual(actual.spill, { r1: 0, c1: 0, r2: 4, c2: 4 });
    const bounded = fixture(f.body); await put(formula(bounded.id, 'A1:E8'), 0, 7);
    const padded = await matrix(8, 5, 0, 7); assert.deepEqual(padded.rows, [...rows, ['', '', '', '', ''], ['', '', '', '', ''], ['', '', '', '', '']]);
    assert.deepEqual(padded.spill, { r1: 0, c1: 7, r2: 7, c2: 11 });
    const opened = fixture(f.body); await put(formula(opened.id, 'A1:E'), 10, 0);
    const openRange = await matrix(5, 5, 10, 0); assert.deepEqual(openRange.rows, rows); assert.deepEqual(openRange.spill, { r1: 10, c1: 0, r2: 14, c2: 4 });
  });
  await test('쉼표 금액·백분율·소수 자릿수는 숫자 값과 표시를 함께 보존', async () => {
    const f = fixture([['금액', '비율', '소수', '코드'], ['1,234,567', '12.50%', '0.1250', '00123']]);
    await put(formula(f.id, 'A1:D2'));
    const actual = await matrix(2, 4); assert.deepEqual(actual.rows[1], [1234567, 0.125, 0.125, '00123']); assert.deepEqual(actual.types[1], ['number', 'number', 'number', 'string']);
    assert.equal((await display(1, 0)).text, '1,234,567'); assert.equal((await display(1, 1)).text, '12.50%'); assert.equal((await display(1, 2)).text, '0.1250');
    const wide = await display(1, 0);
    await page.evaluate(() => { const t = tabula, w = t.wb(); w.transact(() => w.setColWidth(t.si, 0, 50)); t.gv().layout(); });
    const narrow = await display(1, 0); assert.equal(narrow.value, 1234567); assert.equal(narrow.type, 'number'); assert.equal(narrow.text, '1,234,567'); assert.ok(narrow.fontSize < wide.fontSize, '좁은 열은 숫자를 지우거나 ####로 바꾸지 않고 축소 표시');
    if (output) await page.screenshot({ path: join(output, 'import-formats.png') });
  });
  await test('사용자가 지정한 백분율 형식이 가져온 표시 힌트보다 우선하고 실행 취소 가능', async () => {
    const f = fixture([['비율', '표시'], ['12.50%', '합성']]); await put(formula(f.id, 'A1:B2'));
    await page.evaluate(() => { tabula.selectCell(1, 0); tabula.run('fmtPercent'); });
    const custom = await display(1, 0); assert.equal(custom.value, 0.125); assert.equal(custom.type, 'number'); assert.equal(custom.text, '13%');
    await page.evaluate(() => tabula.run('undo')); assert.equal((await display(1, 0)).text, '12.50%');
  });
  await test('시트명을 생략하면 URL의 gid를 무시하고 첫 탭을 가져오며 자격 증명을 전달하지 않음', async () => {
    const f = fixture((count, source) => csv(source.searchParams.has('gid') ? [['다른 탭', 99]] : [['첫 번째 탭', 9]]));
    await put(formula(`https://docs.google.com/spreadsheets/d/${f.id}/edit?access_token=synthetic-secret#gid=73`, 'A1:B1'));
    assert.deepEqual((await matrix(1, 2)).rows, [['첫 번째 탭', 9]]);
    const req = sourceRequest(f); assert.equal(req.query.gid, undefined); assert.equal(req.query.range, 'A1:B1'); assert.equal(req.query.access_token, undefined); assert.equal(req.hash, '');
  });
  await test('한글·따옴표 시트 이름과 B3:D5 오프셋 범위를 중복 자르지 않음', async () => {
    const rows = [['서울', 12, ''], ['부산', 23, '메모'], ['대전', 34, '']], f = fixture(rows);
    f.tabsHtml = `<script>var items=[];items.push({name:"매출'분석",pageUrl:"https://docs.google.com/spreadsheets/d/${f.id}/htmlview/sheet?headers=true&gid=37",gid:"37",initialSheet:("37" == gid)});</script>`;
    await put(formula(f.id, "'매출''분석'!$B$3:$D$5"), 2, 2);
    const req = sourceRequest(f); assert.equal(req.query.gid, '37'); assert.equal(req.query.range, 'B3:D5'); assert.equal(req.query.sheet, undefined);
    assert.equal(requests.filter(r => r.id === f.id && r.path.endsWith('/htmlview')).length, 1);
    const actual = await matrix(3, 3, 2, 2); assert.deepEqual(actual.rows, rows); assert.deepEqual(actual.spill, { r1: 2, c1: 2, r2: 4, c2: 4 });
  });
  await test('웹 게시 주소는 gid를 유지하고 전체 CSV에서 B2:D4를 정확히 자름', async () => {
    const f = fixture([['제외', '제목1', '제목2', '제목3'], ['제외', '서울', 12, ''], ['제외', '', '', ''], ['제외', '부산', 23, '완료']]);
    await put(formula(`https://docs.google.com/spreadsheets/d/e/${f.id}/pubhtml?gid=91`, 'B2:D4'));
    assert.deepEqual((await matrix(3, 3)).rows, [['서울', 12, ''], ['', '', ''], ['부산', 23, '완료']]);
    const req = sourceRequest(f); assert.equal(req.query.gid, '91'); assert.equal(req.query.output, 'csv');
  });
  await test('QUERY와 결합한 큰 금액 집계는 문자열 변환 없이 숫자로 계산', async () => {
    const f = fixture([['매체', '금액'], ['검색', '1,234,567,890,123'], ['영상', 10], ['검색', 7]]);
    await put(queryFormula(f.id, 'A1:B4', "select Col1,sum(Col2) group by Col1 order by sum(Col2) desc label sum(Col2) '합계'"));
    const actual = await matrix(3, 2); assert.deepEqual(actual.rows, [['매체', '합계'], ['검색', 1234567890130], ['영상', 10]]); assert.equal(actual.types[1][1], 'number');
  });
  await test('QUERY의 소수 자료형 null 처리는 직접 가져오기 자료 손실과 구분', async () => {
    const rows = [['구분', '값'], ['A', 10], ['B', '문자'], ['C', 20]], f = fixture(rows);
    await put(formula(f.id, 'A1:B4'));
    await put(queryFormula(f.id, 'A1:B4', 'select Col1,Col2'), 0, 4);
    assert.deepEqual((await matrix(4, 2)).rows, rows);
    assert.deepEqual((await matrix(4, 2, 0, 4)).rows, [['구분', '값'], ['A', 10], ['B', ''], ['C', 20]]);
    assert.equal(f.calls, 1, '같은 원본은 캐시를 공유');
  });
  await test('머리글 수 0과 1이 IMPORTRANGE의 첫 행을 임의로 버리지 않음', async () => {
    const f = fixture([[101, 10], [102, 20]]); await put(queryFormula(f.id, 'A1:B2', 'select Col1,Col2', 0));
    assert.deepEqual((await matrix(2, 2)).rows, [[101, 10], [102, 20]]);
    const g = fixture([['번호', '금액'], [101, 10], [102, 20]]); await put(queryFormula(g.id, 'A1:B3', 'select Col1,Col2', 1), 0, 4);
    assert.deepEqual((await matrix(3, 2, 0, 4)).rows, [['번호', '금액'], [101, 10], [102, 20]]);
  });
  await test('F9는 캐시를 새로 요청하고 원본 값·비율 표시를 갱신', async () => {
    let updated = false;
    const f = fixture(() => csv([['금액', '비율'], [updated ? '2,000' : '1,000', updated ? '15.00%' : '12.50%']]));
    await put(formula(f.id, 'A1:B2')); assert.equal(f.calls, 1);
    updated = true; await page.locator('#cellEditor').focus(); await page.keyboard.press('F9');
    await page.waitForFunction(() => tabula.wb().getValue(tabula.si, 1, 0) === 2000); await frame();
    assert.equal(f.calls, 2); assert.deepEqual((await matrix(2, 2)).rows[1], [2000, 0.15]); assert.equal((await display(1, 1)).text, '15.00%');
  });
  await test('비공개 로그인 HTML은 데이터로 넣지 않고 오류를 표시', async () => {
    const f = fixture('<!doctype html><html><body>로그인 필요</body></html>', { contentType: 'text/html' }); await put(formula(f.id, 'A1:B2'));
    const actual = await matrix(2, 2); assert.equal(actual.rows[0][0], '#N/A'); assert.equal(actual.spill, null); assert.equal(actual.rows[1][0], '');
  });
  await test('긴 분산 텍스트는 이웃 숫자·false·0을 덮지 않고 실제 빈 칸으로만 넘침', async () => {
    const long = '옆 셀을 덮어서는 안 되는 긴 합성 문구입니다';
    const rows = [[long, 42, ''], [long, false, ''], [long, 0, ''], [long, '', ''], [42, long, ''], ['', long, 0], ['', long, '']];
    const f = fixture(rows); await put(formula(f.id, 'A1:C7'));
    const cases = [
      { r: 0, c: 0, align: 'left', overflow: false }, { r: 1, c: 0, align: 'left', overflow: false },
      { r: 2, c: 0, align: 'left', overflow: false }, { r: 3, c: 0, align: 'left', overflow: true },
      { r: 4, c: 1, align: 'right', overflow: false }, { r: 5, c: 1, align: 'center', overflow: false },
      { r: 6, c: 1, align: 'center', overflow: true },
    ];
    await page.evaluate(cases => {
      const t = tabula, w = t.wb();
      w.transact(() => { for (let c = 0; c < 3; c++) w.setColWidth(t.si, c, 70); for (const x of cases) w.setStyle(t.si, x.r, x.c, { align: x.align }); });
      t.gv().layout();
    }, cases);
    await frame();
    const actual = await page.evaluate(cases => cases.map(x => {
      const n = [...document.querySelectorAll(`.c[data-r="${x.r}"][data-c="${x.c}"]`)].find(n => n.getBoundingClientRect().width);
      return { ...x, actualOverflow: n?.classList.contains('ovf'), text: n?.textContent, cssOverflow: n ? getComputedStyle(n).overflow : null };
    }), cases);
    for (const x of actual) { assert.equal(x.text, long); assert.equal(x.actualOverflow, x.overflow, `행 ${x.r + 1} ${x.align} 넘침`); if (!x.overflow) assert.match(x.cssOverflow, /hidden|clip/); }
    assert.deepEqual((await matrix(7, 3)).rows, rows);
    const childRaw = await page.evaluate(() => [0, 1, 2].map(r => tabula.wb().getCell(tabula.si, r, 1)?.raw ?? ''));
    assert.deepEqual(childRaw, ['', '', ''], '인접 값은 raw 입력이 없는 분산 자식');
    if (output) await page.screenshot({ path: join(output, 'spill-text-overflow.png') });
  });
  const summary = { total: results.length, passed: results.filter(r => r.ok).length, failed: results.filter(r => !r.ok).length, results, pageErrors, blockedWrites: writes, unexpectedReads, requests };
  console.log(JSON.stringify(summary));
  if (output) await writeFile(join(output, 'importrange-fidelity.json'), JSON.stringify(summary, null, 2));
  assert.equal(summary.failed, 0); assert.deepEqual(pageErrors, []); assert.deepEqual(writes, []); assert.deepEqual(unexpectedReads, []);
} catch (error) {
  console.error(JSON.stringify({ fatal: error.message, pageErrors, blockedWrites: writes, unexpectedReads }));
  throw error;
} finally { await browser.close(); }
