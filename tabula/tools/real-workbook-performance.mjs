// 실제 파일의 로컬 성능 진단. 값/수식/시트명/필터 항목/화면/원본 CPU profile은 출력하지 않는다.
// node tools/real-workbook-performance.mjs FILE.xlsx [OUTPUT.json]
import fs from 'node:fs/promises';
import path from 'node:path';
import os from 'node:os';
import { performance } from 'node:perf_hooks';

const input = process.argv[2] || process.env.WIXEL_PERF_FILE;
if (!input) { console.error('사용법: node tools/real-workbook-performance.mjs FILE.xlsx [OUTPUT.json]'); process.exit(2); }
const base = process.env.WIXEL_URL || 'http://127.0.0.1:8787/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(base).hostname)) throw new Error('로컬 서버만 허용됩니다.');
const output = path.resolve(process.argv[3] || process.env.WIXEL_PERF_OUTPUT || '../.local/real-workbook-performance.json');
const timeout = Math.min(55000, Math.max(5000, Number(process.env.WIXEL_PERF_TIMEOUT || 55000)));
const cpu = Math.max(1, Number(process.env.WIXEL_PERF_CPU || 1));
const columns = process.env.WIXEL_PERF_COLUMNS?.split(',').map(Number);
const modes = (process.env.WIXEL_PERF_MODES || 'existing').split(',');
const verifyRestore = process.env.WIXEL_PERF_RESTORE === '1';
if (columns?.some(c => !Number.isInteger(c) || c < 0 || c >= 16384) || modes.some(mode => !['existing', 'cleared'].includes(mode))) throw new Error('열 번호/필터 상태 옵션이 올바르지 않습니다.');
const round = n => Math.round(n * 10) / 10;
const report = { started: new Date().toISOString(), inputBytes: (await fs.stat(input)).size, origin: new URL(base).origin,
  cpuThrottle: cpu, cpu: os.cpus()[0]?.model, logicalCpus: os.cpus().length, watchdogMs: timeout, metrics: [], pageErrorCount: 0,
  blockedWrites: 0, blockedExternalRequests: 0, blockedApiReads: 0, note: '실파일 내용·시트명·수식·화면·원본 CPU profile은 기록하지 않음. heap은 현재 JS heap이며 peak가 아님. wallMs에는 자동 저장·Playwright 왕복·두 프레임 대기가 포함됨. CPU self/inclusive는 표본 추정값이며 inclusive 중첩은 합산하면 중복됨.' };
const persist = async () => { await fs.mkdir(path.dirname(output), { recursive: true }); await fs.writeFile(output, JSON.stringify(report, null, 2)); };
const bounded = async (promise, ms, code) => {
  let timer;
  try { return await Promise.race([promise, new Promise((_, reject) => { timer = setTimeout(() => { const e = new Error(code); e.code = code; reject(e); }, ms); })]); }
  finally { clearTimeout(timer); }
};
function profileSummary(profile) {
  const nodes = new Map(profile.nodes.map(n => [n.id, n])), parents = new Map(), self = new Map(), inclusive = new Map();
  for (const n of profile.nodes) for (const id of n.children || []) parents.set(id, n.id);
  const times = profile.timeDeltas || [], samples = profile.samples || [];
  const add = (map, id, dt) => map.set(id, (map.get(id) || 0) + dt);
  for (let i = 0; i < samples.length; i++) {
    const id = samples[i], dt = times[i] || 0; add(self, id, dt);
    let at = id, depth = 0; while (at && depth++ < 80) { add(inclusive, at, dt); at = parents.get(at); }
  }
  const summarize = map => [...map].sort((a, b) => b[1] - a[1]).slice(0, 20).map(([id, us]) => {
    const frame = nodes.get(id)?.callFrame || {}, safeName = String(frame.functionName || '(anonymous)').replace(/[^\w.$<> ():-]/g, '?').slice(0, 100);
    let file = ''; try { const last = new URL(frame.url).pathname.split('/').at(-1); if (/^[\w.-]+\.m?js$/.test(last)) file = last; } catch { /* no source label */ }
    return { function: safeName, file, line: (frame.lineNumber ?? -1) + 1, sampledMs: round(us / 1000) };
  });
  return { sampledDurationMs: round((profile.endTime - profile.startTime) / 1000), samples: samples.length, topSelf: summarize(self), topInclusive: summarize(inclusive) };
}

const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
// launchServer의 PID는 이 검사가 생성한 브라우저만 종료하기 위한 소유권 표식이다.
const server = await chromium.launchServer();
const browser = await chromium.connect(server.wsEndpoint());
const context = await browser.newContext({ viewport: { width: 1440, height: 1000 }, acceptDownloads: false });
let page = await context.newPage(); page.setDefaultTimeout(timeout); report.browser = browser.version();
page.on('pageerror', () => report.pageErrorCount++);
await context.route('**/*', route => {
  const request = route.request();
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { report.blockedWrites++; return route.abort(); }
  const url = new URL(request.url());
  if (url.origin !== new URL(base).origin) { report.blockedExternalRequests++; return route.abort(); }
  if (url.pathname.startsWith('/api/')) { report.blockedApiReads++; return route.abort(); }
  return route.continue();
});
let cdp, profiling = false;
const initializePage = () => {
  window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.__realPerfLong = []; window.__realPerfStarted = 0;
  new PerformanceObserver(list => { for (const e of list.getEntries()) window.__realPerfLong.push({ start: e.startTime, duration: e.duration }); }).observe({ type: 'longtask', buffered: false });
  window.__realPerfReadManifest = () => new Promise((resolve, reject) => {
    const request = indexedDB.open('tabula', 1);
    request.onupgradeneeded = () => { if (!request.result.objectStoreNames.contains('docs')) request.result.createObjectStore('docs'); };
    request.onerror = () => reject(new Error('IDB_OPEN_FAILED'));
    request.onsuccess = () => {
      const db = request.result, tx = db.transaction('docs', 'readonly'), get = tx.objectStore('docs').get('tabula.workbook.v1');
      let value; get.onsuccess = () => { value = get.result; };
      tx.oncomplete = () => { db.close(); resolve(value); };
      tx.onerror = tx.onabort = () => { db.close(); reject(new Error('IDB_READ_FAILED')); };
    };
  });
};
// Fingerprints never appear in the report. All data stays in the isolated browser;
// only SHA-256 digests pass between its two pages in this local Node process.
async function fingerprintBook() {
  const w = window.tabula.wb(), encoder = new TextEncoder();
  const hash = async value => {
    const bytes = ArrayBuffer.isView(value) ? new Uint8Array(value.buffer, value.byteOffset, value.byteLength) : encoder.encode(JSON.stringify(value));
    const out = await crypto.subtle.digest('SHA-256', bytes);
    return Array.from(new Uint8Array(out), v => v.toString(16).padStart(2, '0')).join('');
  };
  const values = [], styles = [], filters = [], counts = [];
  styles.push(await hash({ baseStyle: w.baseStyle, cellStyles: w.cellStyles, theme: w.theme }));
  for (let si = 0; si < w.sheets.length; si++) {
    const s = w.sheets[si], meta = w.sheetMeta(si), chunks = [...w.cellChunks(si)], cellValues = [], cellStyles = [];
    for (const chunk of chunks) for (const [coordinate, data] of chunk) {
      const { style, ...value } = data; cellValues.push([coordinate, value]); cellStyles.push([coordinate, style ?? null]);
    }
    values.push(await hash(cellValues)); styles.push(await hash(cellStyles));
    styles.push(await hash({ allStyle: meta.allStyle, rowStyles: meta.rowStyles, colStyles: meta.colStyles, cond: meta.cond, merges: meta.merges }));
    for (const block of s.blocks || []) {
      values.push(await hash({ r0: block.r0, c0: block.c0, n: block.n, columns: block.cols.length }));
      values.push(await hash(block.perm ?? null));
      for (const col of block.cols) {
        values.push(await hash(col.num ?? null), await hash(col.str ?? null), await hash(col.dict));
        styles.push(await hash(col.fmt ?? null));
      }
    }
    filters.push(await hash({ filter: meta.filter, tables: meta.tables, hiddenRows: meta.hiddenRows, hiddenCols: meta.hiddenCols }));
    const hidden = s.filter?.hidden || {}, extent = w.usedRange(si);
    counts.push({ index: si, cellRecords: s.cells.size, blockRows: (s.blocks || []).reduce((n, b) => n + b.n, 0), rows: extent.rows, columns: extent.cols,
      criteria: Object.keys(s.filter?.criteria || {}).length, hiddenRows: hidden.__bits ? hidden.count : Object.keys(hidden).length });
    await new Promise(resolve => setTimeout(resolve, 0));
  }
  return { values: await hash(values), styles: await hash(styles), filters: await hash(filters), date1904: w.date1904 === true, counts };
}
async function configureProfile() {
  cdp = await context.newCDPSession(page); await cdp.send('Profiler.enable'); await cdp.send('Profiler.setSamplingInterval', { interval: 1000 });
  await cdp.send('Performance.enable'); await cdp.send('Emulation.setCPUThrottlingRate', { rate: cpu });
}
async function stopProfile(metric) {
  if (!profiling) return;
  profiling = false;
  try { const data = await bounded(cdp.send('Profiler.stop'), 3000, 'PROFILE_STOP_TIMEOUT'); metric.profile = profileSummary(data.profile); }
  catch { metric.profileUnavailable = true; }
}
async function measure(name, action) {
  const metric = { name }; report.metrics.push(metric); await persist();
  await bounded(page.evaluate(() => { window.__realPerfLong = []; window.__realPerfStarted = performance.now(); }), 3000, 'METRIC_START_TIMEOUT');
  await bounded(cdp.send('Profiler.start'), 3000, 'PROFILE_START_TIMEOUT'); profiling = true;
  const started = performance.now();
  try {
    const value = await bounded((async () => {
      const value = await action();
      await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(resolve, 0)))));
      return value;
    })(), timeout, 'OPERATION_WATCHDOG');
    metric.wallMs = round(performance.now() - started); await stopProfile(metric);
    const tasks = await bounded(page.evaluate(() => window.__realPerfLong.filter(t => t.start >= window.__realPerfStarted)), 3000, 'LONGTASK_READ_TIMEOUT');
    metric.longTaskCount = tasks.length; metric.maxLongTaskMs = round(tasks.reduce((max, t) => Math.max(max, t.duration), 0)); metric.totalLongTaskMs = round(tasks.reduce((sum, t) => sum + t.duration, 0));
    const { metrics } = await bounded(cdp.send('Performance.getMetrics'), 3000, 'HEAP_READ_TIMEOUT');
    metric.jsHeapMiB = round((metrics.find(m => m.name === 'JSHeapUsedSize')?.value || 0) / 1048576);
    metric.ok = true; await persist(); console.log(JSON.stringify({ step: name, wallMs: metric.wallMs, maxLongTaskMs: metric.maxLongTaskMs, heapMiB: metric.jsHeapMiB })); return value;
  } catch (error) {
    metric.wallMs = round(performance.now() - started); metric.ok = false; metric.errorKind = error.code || error.name || 'Error';
    await stopProfile(metric); await persist(); console.log(JSON.stringify({ step: name, failed: true, errorKind: metric.errorKind, wallMs: metric.wallMs })); throw error;
  }
}
try {
  await context.addInitScript(initializePage);
  await bounded(page.goto(base, { waitUntil: 'domcontentloaded', timeout }), timeout, 'BOOT_WATCHDOG');
  await bounded(page.waitForFunction(() => window.tabula?.wb()), timeout, 'BOOT_WATCHDOG');
  await configureProfile();
  await page.evaluate(() => { window.__previousBook = window.tabula.wb(); });
  await measure('file-open-first-frame', async () => {
    await page.locator('#fileInput').setInputFiles(path.resolve(input));
    await page.waitForFunction(() => window.tabula.wb() !== window.__previousBook && !document.querySelector('.load-progress'));
  });
  report.workbook = await measure('metadata-counts', () => page.evaluate(() => {
    const t = window.tabula, w = t.wb(); return { initialSheetIndex: t.si, sheetCount: w.sheets.length, sheets: w.sheets.map((s, i) => {
      const extent = w.usedRange(i); return { index: i, cellRecords: s.cells.size, usedRows: extent.rows, usedColumns: extent.cols,
        blocks: s.blocks?.length || 0, blockRows: (s.blocks || []).reduce((n, b) => n + b.n, 0), pivots: [s.pivot, ...(s.pivotsExtra || [])].filter(Boolean).length,
        tables: s.tables?.length || 0, charts: s.charts?.length || 0, hasFilter: !!s.filter, hidden: !!s.state };
    }), visibleGridCells: document.querySelectorAll('.c').length, modalCount: document.querySelectorAll('.dialog').length };
  }));
  // 가져오기 경고 내용은 읽지 않고 닫기만 한다.
  for (let i = 0; i < 4 && await page.locator('.dialog').count(); i++) await page.keyboard.press('Escape');
  const candidates = report.workbook.sheets.filter(s => !s.hidden).sort((a, b) => (b.cellRecords + b.blockRows * b.usedColumns) - (a.cellRecords + a.blockRows * a.usedColumns));
  const targetIndex = Number.isInteger(Number(process.env.WIXEL_PERF_SHEET)) && process.env.WIXEL_PERF_SHEET !== undefined ? Number(process.env.WIXEL_PERF_SHEET) : candidates[0].index;
  report.targetSheetIndex = targetIndex;
  await measure('switch-largest-sheet', () => page.evaluate(index => window.tabula.switchSheet(index), targetIndex));
  const target = await measure('filter-target', () => page.evaluate(() => {
    const t = window.tabula, w = t.wb(), s = w.sheets[t.si], table = (s.tables || []).find(item => item.filter !== false);
    const filter = table ? (table.autoFilter || table) : s.filter;
    const used = w.usedRange(t.si), rg = filter || { r1: 0, c1: 0, r2: Math.max(0, used.rows - 1), c2: Math.max(0, used.cols - 1) };
    t.selectCell(rg.r1, rg.c1); return { needsFilter: !filter, table: !!table, row: rg.r1, column: rg.c1, rows: rg.r2 - rg.r1, columns: rg.c2 - rg.c1 + 1 };
  })); report.filterTarget = target;
  if (verifyRestore) await page.evaluate(() => { window.__realPerfOriginalFilter = structuredClone(window.tabula.wb().sheets[window.tabula.si].filter); });
  if (target.needsFilter) await measure('filter-create', () => page.evaluate(() => window.tabula.run('toggleFilter')));
  report.filterResults = [];
  for (const mode of modes) {
    if (mode === 'cleared') await measure('clear-existing-filters', () => page.evaluate(() => window.tabula.run('clearFilter')));
    for (const column of columns || [target.column]) {
      const prefix = `${mode}-c${column}`;
      await bounded(page.evaluate(({ row, column }) => window.tabula.selectCell(row, column), { row: target.row, column }), timeout, 'SELECT_COLUMN_TIMEOUT');
      await page.locator('#cellEditor').focus();
      await measure(`${prefix}-filter-menu-open`, async () => { await page.keyboard.press('Alt+ArrowDown'); await page.locator('.filter-menu').waitFor(); });
      const itemCount = await page.locator('.filter-menu .filter-list input[type="checkbox"]').count();
      const entry = { mode, column, displayedItems: Math.max(0, itemCount - 1) }; report.filterResults.push(entry);
      await measure(`${prefix}-filter-selection`, async () => {
        const menu = page.locator('.filter-menu'), boxes = menu.locator('.filter-list input[type="checkbox"]');
        await menu.getByRole('checkbox', { name: '표시된 항목 모두 선택', exact: true }).uncheck();
        const first = boxes.nth(1); if (!(await first.count())) throw Object.assign(new Error('NO_FILTER_VALUES'), { code: 'NO_FILTER_VALUES' });
        await first.check();
      });
      await measure(`${prefix}-filter-apply`, async () => { await page.locator('.filter-menu').getByRole('button', { name: '확인', exact: true }).click(); await page.locator('.filter-menu').waitFor({ state: 'hidden' }); });
      entry.after = await bounded(page.evaluate(() => {
        const s = window.tabula.wb().sheets[window.tabula.si], f = s.filter || s.tables?.find(t => t.filter)?.filter, hidden = f?.hidden || {};
        return { hiddenRows: hidden.__bits ? hidden.count : Object.keys(hidden).length, criteriaCount: Object.keys(f?.criteria || {}).length };
      }), 3000, 'FILTER_COUNTS_TIMEOUT');
      await measure(`${prefix}-filter-undo`, () => page.keyboard.press('Control+z'));
    }
    await measure(`${mode}-scroll`, () => page.evaluate(async () => {
      const g = window.tabula.gv(), max = Math.max(0, g.scroll.scrollHeight - g.scroll.clientHeight);
      for (const ratio of [.25, .5, .9, 0]) { g.scroll.scrollTop = max * ratio; await new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))); }
    }));
  }
  if (verifyRestore) {
    if (!await page.evaluate(() => !!window.__realPerfOriginalFilter)) throw Object.assign(new Error('RESTORE_REQUIRES_FILTER'), { code: 'RESTORE_REQUIRES_FILTER' });
    await measure('storage-original-filter-change', async () => {
      if (await page.locator('#autosaveToggle').getAttribute('aria-checked') !== 'true') await page.locator('#autosaveToggle').click();
      await page.evaluate(() => {
        const t = window.tabula;
        window.__realPerfPriorGeneration = JSON.parse(localStorage.getItem('tabula.workbook.v1') || 'null')?.generation;
        t.wb().transact(() => t.wb().setSheetProp(t.si, 'filter', structuredClone(window.__realPerfOriginalFilter)));
      });
    });
    const before = await measure('storage-fingerprint-before', () => page.evaluate(fingerprintBook));
    await measure('storage-autosave-complete', async () => {
      // Playwright's waitForFunction treats an async predicate's Promise as truthy.
      // Poll inside the page and await the actual IDB result instead.
      await page.evaluate(async () => {
        for (;;) {
          const ptr = JSON.parse(localStorage.getItem('tabula.workbook.v1') || 'null'), manifest = await window.__realPerfReadManifest(), sheets = window.tabula.wb().sheets;
          if (manifest?.v === 3 && manifest.generation !== window.__realPerfPriorGeneration && ptr?.generation === manifest.generation
            && manifest.sheets.length === sheets.length && manifest.sheets.every((entry, i) => entry.id === sheets[i]._sid && entry.ev === (sheets[i]._ev ?? 0))
            && document.getElementById('saveState')?.textContent.includes('이 브라우저에 저장됨')) return;
          await new Promise(resolve => setTimeout(resolve, 30));
        }
      });
    });
    await bounded(page.close({ runBeforeUnload: false }), 3000, 'PAGE_CLOSE_TIMEOUT');
    page = await context.newPage(); page.setDefaultTimeout(timeout); page.on('pageerror', () => report.pageErrorCount++);
    await configureProfile();
    await measure('storage-new-page-reopen', async () => {
      await page.goto(base, { waitUntil: 'domcontentloaded', timeout });
      await page.waitForFunction(() => window.tabula?.wb() && !document.querySelector('.load-progress'));
    });
    const after = await measure('storage-fingerprint-after', () => page.evaluate(fingerprintBook));
    report.storageRestore = { valuesMatch: before.values === after.values, stylesMatch: before.styles === after.styles,
      filtersMatch: before.filters === after.filters, dateSystemMatch: before.date1904 === after.date1904,
      countsMatch: JSON.stringify(before.counts) === JSON.stringify(after.counts), counts: after.counts };
    if (Object.entries(report.storageRestore).some(([key, value]) => key.endsWith('Match') && value !== true)) throw Object.assign(new Error('STORAGE_RESTORE_MISMATCH'), { code: 'STORAGE_RESTORE_MISMATCH' });
  }
  report.ok = report.pageErrorCount === 0 && report.blockedWrites === 0;
} catch (error) { report.ok = false; report.failureKind = error.code || error.name || 'Error'; process.exitCode = 1; }
finally {
  if (profiling && cdp) await stopProfile(report.metrics.at(-1) || {});
  report.finished = new Date().toISOString(); await persist();
  try { await bounded(browser.close(), 3000, 'BROWSER_CLOSE_TIMEOUT'); } catch { /* Own server PID only. */ }
  try { await bounded(server.close(), 3000, 'SERVER_CLOSE_TIMEOUT'); } catch { try { server.process().kill('SIGKILL'); } catch { /* Already exited. */ } }
  console.log(JSON.stringify({ ok: report.ok, steps: report.metrics.length, failed: report.failureKind || null, pageErrors: report.pageErrorCount, blockedWrites: report.blockedWrites, output }));
}
if (!report.ok) process.exitCode = 1;
