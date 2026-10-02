// CSV 파일 입력 회귀: 새 격리 브라우저와 D:의 합성 파일만 사용한다.
import assert from 'node:assert/strict';
import { mkdirSync, writeFileSync, statSync } from 'node:fs';
import { join, resolve } from 'node:path';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5191/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 합성 검사만 허용합니다.');
const dir = resolve(process.env.CSV_FIXTURE_DIR || 'D:/Codex/Temp/wixel-csv-quality');
if (!/^D:[\\/]/i.test(dir)) throw Error('합성 파일은 D:에 저장하세요.');
mkdirSync(dir, { recursive: true });
const korean = '한글가나다라마바똠힣'.repeat(90), ascii = 'A'.repeat(korean.length);
assert.equal(korean.length, 900);
// Python codecs(cp949)로 대조한 합성 문자 바이트. '똠/힣'은 KS X 1001 밖의 CP949 확장이다.
function cp949(text) {
  const needed = new Set([...text].filter(c => c.charCodeAt(0) > 127));
  const known = Buffer.from('c7d1b1dbb0a1b3aab4d9b6f3b8b6b9d98c63c652', 'hex');
  const map = new Map([...'한글가나다라마바똠힣'].map((c, i) => [c, known.subarray(i * 2, i * 2 + 2)]));
  for (const c of needed) assert.ok(map.has(c), `CP949 합성 문자 ${c}`);
  const bytes = Buffer.alloc(text.length * 2); let at = 0;
  for (const c of text) { if (c.charCodeAt(0) < 128) bytes[at++] = c.charCodeAt(0); else { const pair = map.get(c); bytes[at++] = pair[0]; bytes[at++] = pair[1]; } }
  return bytes.subarray(0, at);
}
const fixtures = [];
function add(id, bytes, extra) { const file = join(dir, `${id}.csv`); writeFileSync(file, bytes); fixtures.push({ id, file, bytes: statSync(file).size, ...extra }); }
add('single-column-160000', Buffer.from(Array.from({ length: 160000 }, (_, i) => `${i}\r\n`).join('')), { single: true });
const rows = '항목,번호\r\n' + Array.from({ length: 5000 }, (_, i) => `${korean},${i}\r\n`).join('');
add('large-utf8', Buffer.from(rows), { header: '항목', koreanRow: 1 });
add('large-utf16le', Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from(rows, 'utf16le')]), { header: '항목', koreanRow: 1 });
const be = Buffer.from(rows, 'utf16le'); be.swap16();
add('large-utf16be', Buffer.concat([Buffer.from([0xfe, 0xff]), be]), { header: '항목', koreanRow: 1 });
const delayed = 'name,number\r\n' + Array.from({ length: 5000 }, (_, i) => `${i < 80 ? ascii : korean},${i}\r\n`).join('');
assert.ok(Buffer.byteLength('name,number\r\n' + Array.from({ length: 80 }, (_, i) => `${ascii},${i}\r\n`).join('')) > 65536);
add('large-cp949-after-ascii', cp949(delayed), { header: 'name', koreanRow: 81, asciiFirst: true });
assert.ok(fixtures[0].bytes < 8 * 1024 * 1024);
for (const f of fixtures.slice(1)) assert.ok(f.bytes > 8 * 1024 * 1024, `${f.id}: 스트리밍 경계 초과`);
const browser = await chromium.launch(), results = [];
try {
  for (const f of fixtures) {
    const context = await browser.newContext({ viewport: { width: 1280, height: 900 } }), errors = [], writes = [];
    await context.route('**/*', route => {
      const r = route.request(), u = new URL(r.url());
      if (!['GET', 'HEAD', 'OPTIONS'].includes(r.method())) { writes.push(`${r.method()} ${u.pathname}`); return route.abort(); }
      return u.origin === new URL(url).origin && !u.pathname.startsWith('/api/') ? route.continue() : route.abort();
    });
    await context.addInitScript(() => { window.WIXEL_SKIP_START = true; window.TABULA_STATIC = true; });
    const page = await context.newPage(); page.setDefaultTimeout(30000); page.on('pageerror', e => errors.push(e.message));
    let details = {}, started;
    try {
      await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }); await page.waitForFunction(() => window.tabula?.wb());
      if (await page.locator('#autosaveToggle').getAttribute('aria-checked') === 'true') await page.locator('#autosaveToggle').click();
      await page.evaluate(() => { window.__csvGaps = []; let last = performance.now(); window.__csvClock = setInterval(() => { const now = performance.now(); __csvGaps.push(now - last); last = now; }, 50); });
      started = Date.now(); await page.locator('#fileInput').setInputFiles(f.file);
      await page.waitForFunction(id => document.querySelector('[role="dialog"]') || (tabula.wb().sheets[0].name === id.slice(0, 31) && !document.querySelector('.load-progress')), f.id, { timeout: 90000 });
      details = await page.evaluate(({ single, row }) => {
        clearInterval(__csvClock); const w = tabula.wb(), s = w.sheets[0];
        return { dialog: document.querySelector('[role="dialog"]')?.textContent || '', sheet: s.name, first: w.getValue(0, 0, 0), last: w.getValue(0, single ? 159999 : 5000, single ? 0 : 1), sample: single ? null : w.getValue(0, row, 0), firstData: single ? null : w.getValue(0, 1, 0), lastText: single ? null : w.getValue(0, 5000, 0), blocks: s.blocks.length, selection: { ...tabula.sel }, maxTickGapMs: Math.round(__csvGaps.reduce((a, b) => Math.max(a, b), 0)) };
      }, { single: f.single, row: f.koreanRow });
      assert.equal(details.dialog, '', '오류 대화상자 없음'); assert.equal(details.last, f.single ? 159999 : 4999, '마지막 숫자');
      if (f.single) {
        assert.equal(details.first, 0); assert.equal(details.selection.r1, 0); assert.equal(details.selection.r2, 159999); assert.equal(details.selection.c2, 0);
        await page.evaluate(() => tabula.run('undo')); assert.equal(await page.evaluate(() => tabula.wb().getValue(0, 159999, 0)), null, '단 한 번 실행 취소');
        await page.evaluate(() => tabula.run('redo')); assert.equal(await page.evaluate(() => tabula.wb().getValue(0, 159999, 0)), 159999, '다시 실행'); details.undoRedo = true;
      } else {
        assert.equal(details.first, f.header); assert.equal(details.sample, korean, '원본 한글 전체'); assert.equal(details.lastText, korean, '마지막 한글 전체');
        if (f.asciiFirst) assert.equal(details.firstData, ascii);
        assert.equal(details.blocks, 1, '큰 CSV 열 블록 경로');
        await page.evaluate(() => { tabula.selectCell(5000, 1); tabula.wb().transact(() => tabula.wb().setInput(0, 5000, 1, '12345')); tabula.run('undo'); });
        assert.equal(await page.evaluate(() => tabula.wb().getValue(0, 5000, 1)), 4999, '블록 끝 셀 편집 실행 취소'); details.endEditUndo = true;
      }
      assert.deepEqual(errors, []); assert.deepEqual(writes, []);
      results.push({ id: f.id, ok: true, bytes: f.bytes, elapsedMs: Date.now() - started, ...details, sample: undefined, firstData: undefined, lastText: undefined, errors, blockedWrites: writes });
      console.log(`OK ${f.id}`);
    } catch (error) {
      results.push({ id: f.id, ok: false, bytes: f.bytes, elapsedMs: started ? Date.now() - started : null, error: error.message, ...details, sample: typeof details.sample === 'string' ? details.sample.slice(0, 80) : details.sample, firstData: undefined, lastText: undefined, errors, blockedWrites: writes });
      console.error(`NG ${f.id}: ${error.message}`);
    } finally { await context.close(); }
  }
} finally { await browser.close(); }
const report = { url, fixturesDirectory: dir, results, passed: results.filter(r => r.ok).length, total: results.length };
if (process.env.CSV_RESULT_PATH) writeFileSync(process.env.CSV_RESULT_PATH, JSON.stringify(report, null, 2));
console.log(JSON.stringify(report, null, 2));
if (report.passed !== report.total) process.exitCode = 1;
