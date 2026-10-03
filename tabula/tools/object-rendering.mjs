// 개체 렌더 성능·회귀: 별도 브라우저 컨텍스트의 합성 문서만 사용한다.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const url = process.env.WIXEL_URL || 'http://127.0.0.1:5180/';
if (!['localhost', '127.0.0.1', '[::1]'].includes(new URL(url).hostname)) throw new Error('로컬 서버에서만 실행하세요.');
const baseline = process.env.WIXEL_OBJECT_BASELINE === '1';
const browser = await chromium.launch();
const context = await browser.newContext({ viewport: { width: 1400, height: 950 } });
const page = await context.newPage(), errors = [], writes = [], results = [];
page.on('pageerror', e => errors.push(e.message));
await context.route('**/*', route => {
  if (!['GET', 'HEAD', 'OPTIONS'].includes(route.request().method())) { writes.push(route.request().url()); return route.abort(); }
  return route.continue();
});
page.setDefaultTimeout(30000);
try {
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await page.waitForFunction(() => window.tabula?.gv());
  if (await page.locator('#autosaveToggle').getAttribute('aria-checked') !== 'false') await page.evaluate(() => { const el = document.querySelector('#autosaveToggle'); if (el?.classList.contains('on') || el?.getAttribute('aria-checked') === 'true') el.click(); });
  await page.evaluate(() => {
    const t = window.tabula, w = t.wb(), v = t.gv();
    w.load({ sheets: [{ name: '합성 개체', cells: { '0,0': { raw: '이름' }, '0,1': { raw: '금액' }, '1,0': { raw: '가' }, '1,1': { raw: '10' }, '2,0': { raw: '나' }, '2,1': { raw: '20' } } }] });
    const chart = (id, x, y, hidden = false) => ({ id, type: 'column', title: id, range: { r1: 0, c1: 0, r2: 2, c2: 1 }, x, y, w: 300, h: 190, hidden });
    const slicer = (id, x, y, hidden = false) => ({ id, x, y, w: 150, h: 190, caption: id, source: { kind: 'table', table: '객체표', column: '이름' }, hidden });
    const pic = (id, x, y, linked = false) => ({ id, x, y, w: 120, h: 90, src: 'data:image/svg+xml;base64,' + btoa('<svg xmlns="http://www.w3.org/2000/svg" width="20" height="20"><rect width="20" height="20" fill="blue"/></svg>'), ...(linked ? { linked: { sheet: '합성 개체', r1: 0, c1: 0, r2: 2, c2: 1 } } : {}) });
    w.transact(() => {
      w.setSheetProp(0, 'tables', [{ id: 'table', name: '객체표', r1: 0, c1: 0, r2: 2, c2: 1, header: true, columns: [{ name: '이름' }, { name: '금액' }], filter: { criteria: {}, hidden: {} } }]);
      w.setSheetProp(0, 'charts', [chart('visible-chart', 100, 80), chart('hidden-chart', 100, 80, true), ...Array.from({ length: 40 }, (_, i) => chart('far-chart-' + i, 3000 + i * 340, 3500))]);
      w.setSheetProp(0, 'slicers', [slicer('visible-slicer', 450, 80), slicer('hidden-slicer', 450, 80, true), ...Array.from({ length: 40 }, (_, i) => slicer('far-slicer-' + i, 3000 + i * 170, 3750))]);
      w.setSheetProp(0, 'images', [pic('visible-pic', 650, 80), pic('visible-linked', 650, 220, true), ...Array.from({ length: 40 }, (_, i) => pic('far-pic-' + i, 3000 + i * 130, 4000, i % 2 === 0))]);
      w.setSheetProp(0, 'shapes', [{ id: 'visible-shape', kind: 'rect', x: 820, y: 80, w: 100, h: 100, fill: '#4472c4' }, ...Array.from({ length: 40 }, (_, i) => ({ id: 'far-shape-' + i, kind: 'rect', x: 3000 + i * 130, y: 4150, w: 100, h: 100 }))]);
    });
    t.selectCell(0, 0); v.setScroll(0, 0); v.layout();
    window.__objectCalls = {};
    for (const name of ['chartSvg', 'slicerHtml', 'linkedHtml']) {
      const fn = v[name]; v[name] = function(o) { const calls = window.__objectCalls; calls[name] ??= []; calls[name].push(o.id); return fn.call(this, o); };
    }
    window.__objectMeasure = () => {
      window.__objectCalls = {}; const start = performance.now(); v.renderObjectsAll();
      return { ms: +(performance.now() - start).toFixed(2), calls: structuredClone(window.__objectCalls), nodes: [...document.querySelectorAll('.obj')].map(n => n.dataset.id) };
    };
  });
  await page.evaluate(() => new Promise(resolve => requestAnimationFrame(() => requestAnimationFrame(resolve))));
  const measurement = await page.evaluate(() => [window.__objectMeasure(), window.__objectMeasure(), window.__objectMeasure()]);
  console.log(JSON.stringify({ baseline, objects: 167, measurement: measurement.map(m => ({ ms: m.ms, calls: Object.fromEntries(Object.entries(m.calls).map(([k, v]) => [k, v.length])), renderedNodes: m.nodes.length })) }));
  if (!baseline) {
    assert.deepEqual(measurement.at(-1).calls, {}, '동일 데이터·개체의 내용 생성 재사용');
    assert.deepEqual([...new Set(measurement[0].nodes)].sort(), ['visible-chart', 'visible-linked', 'visible-pic', 'visible-shape', 'visible-slicer'].sort());
    results.push('화면 밖·숨긴 개체 제외와 동일 개체 재사용');
    const changes = await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), v = t.gv(), ch = w.sheets[0].charts[0];
      const before = document.querySelector('[data-id="visible-chart"] svg').outerHTML;
      w.transact(() => w.setInput(0, 1, 1, '999'));
      const values = window.__objectMeasure();
      const changed = document.querySelector('[data-id="visible-chart"] svg').outerHTML !== before;
      ch.title = '직접 바뀐 제목'; ch.w = 330; ch.seriesFmt = [{ color: '#ff0000' }];
      const style = window.__objectMeasure();
      const node = document.querySelector('[data-id="visible-chart"]');
      return { values, changed, style, title: node.textContent.includes('직접 바뀐 제목'), red: node.innerHTML.includes('#ff0000'), width: node.style.width };
    });
    assert.equal(changes.changed, true); assert.deepEqual(changes.values.calls.chartSvg, ['visible-chart']);
    assert.deepEqual(changes.style.calls.chartSvg, ['visible-chart']); assert.equal(changes.title, true); assert.equal(changes.red, true); assert.equal(changes.width, '330px');
    results.push('값·직접 서식·크기 변경 캐시 무효화');
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(resolve)));
    const scroll = await page.evaluate(() => {
      const v = window.tabula.gv(); v.setScroll(2900, 3400); v.renderObjectsAll();
      const far = [...document.querySelectorAll('.obj')].map(n => n.dataset.id);
      v.setScroll(0, 0); v.renderObjectsAll();
      return { far, home: [...document.querySelectorAll('.obj')].map(n => n.dataset.id) };
    });
    assert.ok(scroll.far.includes('far-chart-0')); assert.ok(scroll.far.includes('far-slicer-0')); assert.ok(scroll.far.includes('far-pic-0'));
    assert.ok(!scroll.far.includes('visible-chart')); assert.ok(scroll.home.includes('visible-chart'));
    results.push('스크롤 진입·이탈·복귀');
    const overscan = await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), v = t.gv(), item = w.sheets[0].charts[1];
      item.hidden = false; item.x = v.viewW - v.hw + 10; item.y = 120;
      v.renderObjectsAll(); const before = v.panes[3].win;
      const prepared = !!document.querySelector('.pane-br [data-id="hidden-chart"]');
      v.setScroll(100, 0);
      const node = document.querySelector('.pane-br [data-id="hidden-chart"]'), box = node?.getBoundingClientRect(), pane = v.panes[3].el.getBoundingClientRect();
      const result = { prepared, sameWindow: before === v.panes[3].win, shown: !!box && box.left < pane.right && box.right > pane.left };
      item.hidden = true; v.setScroll(0, 0); v.renderObjectsAll(); return result;
    });
    assert.deepEqual(overscan, { prepared: true, sameWindow: true, shown: true }); results.push('셀 창 재생성 없는 작은 스크롤에서도 여유 영역 개체 표시');
    const frozen = await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), v = t.gv();
      w.transact(() => w.setSheetProp(0, 'freeze', { rows: 6, cols: 3 }));
      v.layout(); window.__objectCalls = {}; v._objectRenderCache = null; v.renderObjectsAll();
      return { calls: window.__objectCalls, panes: [...document.querySelectorAll('[data-id="visible-chart"]')].map(n => n.closest('.pane').className) };
    });
    assert.ok(frozen.panes.length >= 2, '경계를 가로지르는 개체는 고정 창 양쪽에 존재');
    assert.deepEqual(frozen.calls.chartSvg, ['visible-chart'], '여러 고정 창에서도 차트 내용 생성은 한 번');
    results.push('고정 창 겹침과 공통 내용 캐시');
    const shape = await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), v = t.gv();
      w.transact(() => w.setSheetProp(0, 'freeze', { rows: 0, cols: 0 })); v.layout(); v.setScroll(0, 0); // 틀 해제는 현재 스크롤을 보존하므로 선택 검사의 기준 화면으로 복귀
      const p = v.panes[3], edge = v.cols.pos(p.win.c2 + 1);
      w.transact(() => w.setSheetProp(0, 'shapes', [...w.sheets[0].shapes, { id: 'rotated-edge', kind: 'rect', x: edge + 30, y: 150, w: 20, h: 240, rot: 90, fill: '#f00' }]));
      v.renderObjectsAll(); return { shown: !!document.querySelector('[data-id="rotated-edge"]') };
    });
    assert.equal(shape.shown, true); results.push('회전 개체의 표시 경계 보존');
    await page.evaluate(() => new Promise(resolve => requestAnimationFrame(resolve)));
    await page.locator('.pane-br [data-id="visible-chart"]').click({ position: { x: 12, y: 12 } });
    assert.equal(await page.locator('.pane-br [data-id="visible-chart"].sel .ch-h').count(), 4);
    const selection = await page.evaluate(() => window.__objectMeasure());
    assert.deepEqual(selection.calls, {}, '선택은 내용 캐시를 유지하며 손잡이만 갱신');
    results.push('차트 선택 손잡이와 캐시 공존');
    const undo = await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), v = t.gv();
      const current = () => document.querySelector('.pane-br [data-id="visible-chart"] svg').outerHTML;
      const before = current();
      w.transact(() => w.setInput(0, 1, 1, '777')); v.renderObjectsAll(); const changed = current();
      w.undo(); v.renderObjectsAll(); const restored = current();
      const ch = w.sheets[0].charts[0]; ch.hidden = true; v.renderObjectsAll();
      const hidden = !document.querySelector('.pane-br [data-id="visible-chart"]');
      ch.hidden = false; v.renderObjectsAll();
      return { changed: changed !== before, restored: restored === before, hidden, reshown: !!document.querySelector('.pane-br [data-id="visible-chart"]') };
    });
    assert.deepEqual(undo, { changed: true, restored: true, hidden: true, reshown: true }); results.push('실행 취소 및 숨김·표시 토글');
    const slicer = await page.evaluate(() => {
      const t = window.tabula, w = t.wb(), v = t.gv();
      w.transact(() => w.setInput(0, 1, 0, '새 항목')); v.renderObjectsAll();
      const text = document.querySelector('.pane-br [data-id="visible-slicer"]').textContent;
      const linked = document.querySelector('.pane-br [data-id="visible-linked"]').textContent;
      return { slicer: text.includes('새 항목'), linked: linked.includes('새 항목') };
    });
    assert.deepEqual(slicer, { slicer: true, linked: true }); results.push('표 데이터 변경이 슬라이서·연결 그림에 반영');
    const serialized = await page.evaluate(() => { const w = window.tabula.wb(), saved = w.serialize().sheets[0]; return { charts: saved.charts.length, images: saved.images.length, slicers: saved.slicers.length, shapes: saved.shapes.length }; });
    assert.deepEqual(serialized, { charts: 42, images: 42, slicers: 42, shapes: 42 }); results.push('가상화는 원본 개체·저장 데이터에 영향 없음');
  }
  assert.deepEqual(errors, []); assert.deepEqual(writes, []);
  console.log(JSON.stringify({ passed: results.length, results, pageErrors: errors, blockedWrites: writes }));
} finally { await context.close(); await browser.close(); }
