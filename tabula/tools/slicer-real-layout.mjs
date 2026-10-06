// Read-only layout audit of an explicitly provided local workbook in an isolated browser.
import assert from 'node:assert/strict';
import { readFile, mkdir, writeFile, stat } from 'node:fs/promises';
import { resolve, join, isAbsolute } from 'node:path';
import { pathToFileURL } from 'node:url';
import { readXlsx } from '../src/xlsx.js';
import { DEFAULT_COL_WIDTH, DEFAULT_ROW_HEIGHT } from '../src/workbook.js';
import { Axis } from '../src/axis.js';
import { captureDrawingAnchors } from '../src/drawing-anchor.js';

const file = resolve(process.argv[2] || process.env.REAL_WORKBOOK || '');
if (!process.argv[2] && !process.env.REAL_WORKBOOK) throw Error('사용법: node tools/slicer-real-layout.mjs "로컬 파일.xlsb"');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5193/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw Error('로컬 검증 서버에서만 실행하세요.');
const output = resolve(process.env.SLICER_REAL_OUTPUT || '.local/slicer-real-layout');
await mkdir(output, { recursive: true });
const sourceInfo = await stat(file);
console.log('Reading local workbook ' + sourceInfo.size + ' bytes');
let raw = readXlsx(new Uint8Array(await readFile(file)));
const sourceMetadata = raw.data.sheets.map((s, si) => {
  const cols = new Axis(s.defColW ?? DEFAULT_COL_WIDTH, s.colWidths, [s.hiddenCols], 16384);
  const rows = new Axis(s.defRowH ?? DEFAULT_ROW_HEIGHT, s.rowHeights, [s.hiddenRows, s.filter?.hidden, ...(s.tables ?? []).map(t => t.filter?.hidden)], 1048576);
  const captured = captureDrawingAnchors(s, cols, rows);
  return {
    si, name: s.name, state: s.state, zoom: s.zoom ?? s.view?.zoom, fitRows: raw.data.fitRows?.[si]?.length ?? 0,
    defRowH: s.defRowH ?? DEFAULT_ROW_HEIGHT, defColW: s.defColW ?? DEFAULT_COL_WIDTH, rowHeights: { ...s.rowHeights }, colWidths: { ...s.colWidths },
    slicers: (s.slicers ?? []).map((o, index) => ({ id: o.id, index, caption: o.caption, placement: o.placement ?? 'oneCell', x: o.x, y: o.y, w: o.w, h: o.h, anchor: captured.get('slicers:' + o.id) })),
    pivotAreas: [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean).map(p => ({ name: p.name, ...p.area })),
    tables: (s.tables ?? []).map(t => ({ name: t.name, r1: t.r1, c1: t.c1, r2: t.r2, c2: t.c2 })),
  };
});
const sourceActive = raw.active ?? 0, warnings = raw.warnings;
raw = null;
console.log(JSON.stringify(sourceMetadata.map(s => ({ si: s.si, name: s.name, slicers: s.slicers.length, fitRows: s.fitRows, pivots: s.pivotAreas.length }))));
const browserModule = process.env.PLAYWRIGHT_MODULE || 'playwright';
const { chromium } = await import(isAbsolute(browserModule) ? pathToFileURL(browserModule).href : browserModule);
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 2048, height: 1260 }, deviceScaleFactor: 1 });
const page = await context.newPage(), errors = [], writes = [], snapshots = [], comparison = [];
let checks = 0;
const close = (actual, expected, label, tolerance = .08) => { assert.ok(Math.abs(actual - expected) <= tolerance, label + ': ' + actual + ' ≈ ' + expected); checks++; };
page.setDefaultTimeout(30000);
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => {
  const request = route.request(), target = new URL(request.url());
  if (!['GET', 'HEAD', 'OPTIONS'].includes(request.method())) { writes.push(request.method()); return route.abort(); }
  if (target.origin !== new URL(url).origin || target.pathname.startsWith('/api/')) return route.abort();
  return route.continue();
});
await context.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
const settle = async () => {
  await page.evaluate(() => document.fonts.ready);
  await page.evaluate(() => new Promise(r => requestAnimationFrame(() => requestAnimationFrame(r))));
};
let failure;
try {
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.gv(), null, { timeout: 60000 });
  await page.locator('#fileInput').setInputFiles(file);
  await page.waitForFunction(name => tabula.wb().sheets.some(s => s.name === name) && !document.querySelector('.load-progress'), sourceMetadata[sourceActive].name, { timeout: 300000 });
  await settle();
  for (const source of sourceMetadata.filter(s => s.slicers.length && s.state !== 'hidden' && s.state !== 'veryHidden')) {
    await page.evaluate(si => { tabula.switchSheet(si); tabula.gv().setScroll(0, 0); tabula.gv().renderAll(); }, source.si); await settle();
    const afterOpen = await page.evaluate(source => {
      const t = tabula, g = t.gv(), s = t.wb().sheets[t.si];
      return {
        name: s.name, rowHeights: { ...s.rowHeights }, colWidths: { ...s.colWidths },
        slicers: source.slicers.map(before => {
          const actual = s.slicers[before.index];
          if (!actual) return { caption: before.caption, missing: true };
          const a = before.anchor, placement = before.placement;
          let expected = { x: before.x, y: before.y, w: before.w, h: before.h };
          if (a) {
            const position = (axis, marker) => axis.pos(marker[0]) + Math.min(marker[1], axis.size(marker[0]) || marker[1]);
            expected.x = position(g.cols, a.c1); expected.y = position(g.rows, a.r1);
            if (placement === 'twoCell') {
              expected.w = Math.max(1, position(g.cols, a.c2) - expected.x);
              expected.h = Math.max(1, position(g.rows, a.r2) - expected.y);
            }
          }
          return { caption: actual.caption, placement, before: { x: before.x, y: before.y, w: before.w, h: before.h }, expected,
            actual: { id: actual.id, x: actual.x, y: actual.y, w: actual.w, h: actual.h } };
        }),
      };
    }, source);
    const rowChanges = [];
    for (const r of new Set([...Object.keys(source.rowHeights), ...Object.keys(afterOpen.rowHeights)])) {
      const before = source.rowHeights[r] ?? source.defRowH, after = afterOpen.rowHeights[r] ?? source.defRowH;
      if (before !== after) rowChanges.push({ row: +r, before, after });
    }
    const columnChanges = [];
    for (const c of new Set([...Object.keys(source.colWidths), ...Object.keys(afterOpen.colWidths)])) {
      const before = source.colWidths[c] ?? source.defColW, after = afterOpen.colWidths[c] ?? source.defColW;
      if (before !== after) columnChanges.push({ col: +c, before, after });
    }
    comparison.push({ changedColumnCount: columnChanges.length, changedColumns: columnChanges.slice(0, 100), si: source.si, name: source.name, fitRows: source.fitRows, changedRowCount: rowChanges.length, changedRows: rowChanges.slice(0, 100), slicers: afterOpen.slicers });
    for (const o of afterOpen.slicers) {
      assert.equal(o.missing, undefined, source.name + ': 원본 슬라이서 보존'); checks++;
      for (const key of ['x', 'y', 'w', 'h']) close(o.actual[key], o.expected[key], source.name + '/' + o.caption + ' 파일 열기 ' + key, 1 / 9525 + 1e-7);
    }
    const initialGeometry = afterOpen.slicers.map(o => o.actual);
    for (const percent of [25, 55, 70, 100, 150, 200, 400]) {
      await page.locator('#zoomSlider').evaluate((input, value) => { input.value = String(value); input.dispatchEvent(new Event('input', { bubbles: true })); }, percent);
      await page.evaluate(() => { tabula.gv().setScroll(0, 0); tabula.gv().renderAll(); }); await settle();
      const snapshot = await page.evaluate(({ source, percent }) => {
        const t = tabula, g = t.gv(), s = t.wb().sheets[t.si], panes = [];
        const areas = [...(source.si === 0 ? [{ name: '상단 정보 B2:C5', r1: 1, c1: 1, r2: 4, c2: 2 }] : []), ...source.pivotAreas.filter(a => Number.isInteger(a.r1))];
        const tableEdges = areas.map(a => ({ name: a.name, range: a, rect: g.sheetRect(a) }));
        for (const pane of g.panes) {
          if (pane.el.style.display === 'none') continue;
          const origin = pane.content.getBoundingClientRect();
          const normalized = node => { const r = node.getBoundingClientRect(); return { x: (r.left - origin.left) / g.z + pane.ox, y: (r.top - origin.top) / g.z + pane.oy, w: r.width / g.z, h: r.height / g.z }; };
          const objects = [...pane.objects.querySelectorAll('.obj.slicer:not(.slicer-group)')].map(node => {
            const o = s.slicers.find(o => o.id === node.dataset.id); return { caption: o.caption, actual: normalized(node), expected: { x: o.x, y: o.y, w: o.w, h: o.h } };
          });
          const cells = [...pane.cells.querySelectorAll('.c')].filter(node => areas.some(a => +node.dataset.r === a.r1 && (+node.dataset.c === a.c1 || +node.dataset.c === a.c2)))
            .map(node => ({ r: +node.dataset.r, c: +node.dataset.c, actual: normalized(node), expected: g.sheetRect({ r1: +node.dataset.r, c1: +node.dataset.c, r2: +node.dataset.r, c2: +node.dataset.c }) }));
          panes.push({ id: pane.id, objects, cells });
        }
        return { si: t.si, name: s.name, percent, zoom: g.z * 100, scroll: { x: g.sx, y: g.sy }, tableEdges, panes,
          geometry: s.slicers.map(o => ({ id: o.id, x: o.x, y: o.y, w: o.w, h: o.h })) };
      }, { source, percent });
      for (const pane of snapshot.panes) for (const item of [...pane.objects, ...pane.cells]) for (const key of ['x', 'y', 'w', 'h']) close(item.actual[key], item.expected[key], source.name + '/' + percent + '/' + (item.caption || item.r + ',' + item.c) + '/' + key);
      for (const before of initialGeometry) {
        const after = snapshot.geometry.find(o => o.id === before.id);
        for (const key of ['x', 'y', 'w', 'h']) close(after[key], before[key], source.name + '/' + percent + ' 원본 좌표 고정 ' + key, 1e-7);
      }
      for (const table of snapshot.tableEdges) for (const slicer of initialGeometry) {
        const a = slicer, b = table.rect;
        assert.ok(a.x >= b.x + b.w - .08 || a.x + a.w <= b.x + .08 || a.y >= b.y + b.h - .08 || a.y + a.h <= b.y + .08, source.name + '/' + percent + '/' + table.name + ' 표와 슬라이서 겹침'); checks++;
      }
      snapshots.push(snapshot);
      console.log('OK real sheet ' + source.si + ' ' + percent + '% ' + snapshot.panes.reduce((n, p) => n + p.objects.length, 0) + ' slicer copies');
      if ([55, 70, 100].includes(percent)) await page.screenshot({ path: join(output, 'sheet-' + source.si + '-zoom-' + percent + '.png') });
    }
  }
  assert.equal(comparison.length > 0, true, '검사 가능한 표시 슬라이서 시트 존재'); checks++;
  assert.deepEqual(errors, [], '페이지 오류'); checks++;
  assert.deepEqual(writes, [], '외부 쓰기 없음'); checks++;
  const unchanged = await stat(file);
  assert.equal(unchanged.size, sourceInfo.size, '원본 파일 크기 보존'); checks++;
  assert.equal(unchanged.mtimeMs, sourceInfo.mtimeMs, '원본 파일 수정 시각 보존'); checks++;
} catch (error) { failure = error; console.error(error.stack); }
finally {
  await writeFile(join(output, 'real-slicer-layout.json'), JSON.stringify({ file, url, sourceActive, warnings, checks, ok: !failure, error: failure?.message, sourceMetadata, comparison, snapshots, pageErrors: errors, blockedWrites: writes }, null, 2));
  await browser.close();
}
console.log(JSON.stringify({ ok: !failure, checks, sheets: comparison.length, zoomSnapshots: snapshots.length, output }));
if (failure) throw failure;
