// DOM XSS 회귀 검사. CSP를 의도적으로 우회하여 HTML 정화 자체의 방어를 검사한다.
// npm start 후 PLAYWRIGHT_MODULE / WIXEL_URL 설정은 tools/README.md 참고.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
const results = [];
try {
  const context = await browser.newContext({ bypassCSP: true, viewport: { width: 1440, height: 1000 } });
  const page = await context.newPage();
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto(process.env.WIXEL_URL || 'http://localhost:5178/');
  await page.waitForFunction(() => !!window.tabula);
  const check = async (name, fn) => {
    try { await fn(); results.push({ name, ok: true }); console.log(`OK ${name}`); }
    catch (error) { results.push({ name, ok: false }); console.error(`NG ${name}: ${error.stack}`); }
  };

  await check('CSP 없이 차트 속성 탈출·이벤트 실행 차단', async () => {
    const result = await page.evaluate(async () => {
      const { renderChartSvg } = await import('/src/chart.js');
      const { el } = await import('/src/ui.js');
      window.__wixelXss = 0;
      const svg = renderChartSvg({ type: 'column', w: 300, h: 200, fill: 'white" onmouseover="window.__wixelXss++' }, { categories: ['x'], series: [{ name: 'A', values: [1] }] });
      const host = el('div', { html: svg }); document.body.append(host);
      for (const node of host.querySelectorAll('*')) node.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
      const out = { hasHandler: !!host.querySelector('[onmouseover]'), executed: window.__wixelXss, svg: !!host.querySelector('svg') };
      host.remove(); return out;
    });
    assert.deepEqual(result, { hasHandler: false, executed: 0, svg: true });
  });

  await check('위험 태그·네임스페이스·mXSS·주소·CSS 공격 조각 차단', async () => {
    const result = await page.evaluate(async () => {
      const { setSafeHtml } = await import('/src/safe-html.js');
      window.__wixelXss = 0;
      const attack = 'window.__wixelXss++';
      const payloads = [
        `<script>${attack}</script><iframe srcdoc="<script>${attack}</script>"></iframe><object data="data:text/html,bad"></object><embed src="data:text/html,bad"><base href="https://evil.invalid"><link rel="stylesheet" href="https://evil.invalid/a"><meta http-equiv="refresh" content="0;url=javascript:${attack}">`,
        `<img src="bad://image" onerror="${attack}"><svg onload="${attack}"><rect onmouseover="${attack}"/></svg>`,
        `<svg><foreignObject><div onmouseover="${attack}">x</div></foreignObject><script>${attack}</script><animate attributeName="href" values="javascript:${attack}"/></svg>`,
        `<math><mtext><table><mglyph><style><!--</style><img title="--><img src=x onerror=${attack}>">`,
        `<svg><desc><![CDATA[</desc><img src=x onerror="${attack}">]]></svg>`,
        `<a href="java&#10;script:${attack}">a</a><a href="data:text/html,bad">b</a><a href="file:///D:/secret">c</a><a href="vbscript:bad">d</a>`,
        `<img src="data:text/html,bad"><svg><use href="https://evil.invalid/a.svg#x"/><image href="javascript:${attack}"/></svg>`,
        `<div style="background:url(javascript:${attack});width:2px">x</div><div style="background:u\\72l(javascript:bad)">y</div><div style="background:image-set('https://evil.invalid/x' 1x)">z</div>`,
        `<form id="tabula"><input name="location" autofocus onfocus="${attack}"></form><input id="toast" name="tabula" srcdoc="bad" formaction="javascript:${attack}">`,
        `<template shadowrootmode="open"><img src=x onerror="${attack}"></template><svg><a xlink:href="javascript:${attack}">x</a></svg>`,
      ];
      const host = document.createElement('div'); document.body.append(host);
      const failures = [];
      for (let i = 0; i < payloads.length; i++) {
        setSafeHtml(host, payloads[i]);
        if (host.querySelector('script,iframe,object,embed,link,meta,base,form,foreignObject,math,style,animate,template')) failures.push(`${i}: tag`);
        for (const node of host.querySelectorAll('*')) {
          for (const attr of node.attributes) {
            if (/^on|srcdoc|formaction/i.test(attr.name)) failures.push(`${i}: ${attr.name}`);
            if (/^(?:href|src|xlink:href)$/i.test(attr.name) && /javascript|vbscript|file:|data:text/i.test(attr.value)) failures.push(`${i}: URL`);
          }
          for (const name of ['load', 'error', 'mouseover', 'focus']) node.dispatchEvent(new Event(name));
        }
      }
      await new Promise((resolve) => requestAnimationFrame(resolve));
      host.remove(); return { failures, executed: window.__wixelXss, count: payloads.length };
    });
    assert.deepEqual(result.failures, []);
    assert.equal(result.executed, 0);
    assert.equal(result.count, 10);
  });

  await check('정상 SVG 효과·아이콘·셀 무늬·표 미리보기 보존', async () => {
    const result = await page.evaluate(async () => {
      const { sanitizeHtml, setSafeHtml } = await import('/src/safe-html.js');
      const { shapeSvg } = await import('/src/shapes.js');
      const { patternCss } = await import('/src/view.js');
      const { hydrateIcons } = await import('/src/ui.js');
      const host = document.createElement('div'); document.body.append(host);
      const shape = shapeSvg({ id: 'safe-shape', kind: 'rect', w: 160, h: 90, fill: '#ff0000', grad: { ang: 40, stops: [[0, '#ff0000', 0.5], [1, '#0000ff']] }, shadow: { dx: 2, dy: 3, blur: 4, color: '#222222', opacity: 0.4 }, glow: { color: '#00ff00', size: 3, opacity: 0.6 }, stroke: '#000000', strokeWidth: 2 });
      const fragment = sanitizeHtml(shape);
      const first = fragment.firstChild; host.append(fragment);
      const out = { sameNode: host.firstChild === first, gradient: !!host.querySelector('linearGradient'), shadow: !!host.querySelector('feDropShadow'), glow: !!host.querySelector('feMorphology'), references: [...host.querySelectorAll('[fill],[filter]')].some((x) => /url\(#/.test(x.getAttribute('fill') || x.getAttribute('filter'))) };
      const css = patternCss('darkGrid', '#112233', '#eeeeee');
      setSafeHtml(host, `<table><tbody><tr><td style="background:${css};text-align:center">123</td></tr></tbody></table><span data-icon="save"></span>`);
      hydrateIcons(host);
      out.table = host.querySelector('td')?.textContent === '123';
      out.pattern = host.querySelector('td')?.style.backgroundImage.startsWith('url(');
      out.icon = !!host.querySelector('[data-icon] svg');
      host.remove(); return out;
    });
    for (const [key, value] of Object.entries(result)) assert.equal(value, true, key);
  });

  await check('실제 격자·개체·차트 미리보기의 악성 JSON과 일반 맞춤', async () => {
    await page.evaluate(() => {
      const t = window.tabula; const wb = t.wb();
      window.__wixelXss = 0;
      wb.transact(() => {
        wb.setInput(0, 0, 0, '123'); wb.setStyle(0, 0, 0, { align: 'general', size: '11" onmouseover="window.__wixelXss++' });
        wb.setInput(0, 1, 0, '일반 문자'); wb.setStyle(0, 1, 0, { align: 'general' });
        wb.setInput(0, 2, 0, '456'); wb.setStyle(0, 2, 0, { align: 'general' });
        wb.setSheetProp(0, 'charts', [{ id: 'security-chart', title: '보안 검증', type: 'column', x: 300, y: 30, w: 400, h: 250, range: { r1: 0, c1: 0, r2: 1, c2: 0 }, fill: 'white" onmouseover="window.__wixelXss++' }]);
        wb.setSheetProp(0, 'shapes', [{ id: 'security-shape', kind: 'rect', x: 750, y: 30, w: 150, h: 80, fill: '#112233', strokeOpacity: '1" onmouseover="window.__wixelXss++' }]);
      });
      t.gv().renderAll();
    });
    const before = await page.evaluate(() => {
      for (const node of document.querySelectorAll('.c,.obj svg,.obj rect')) node.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
      const number = document.querySelector('.c[data-r="2"][data-c="0"]');
      const text = document.querySelector('.c[data-r="1"][data-c="0"]');
      return { executed: window.__wixelXss, handlers: document.querySelectorAll('.c[onmouseover],.obj [onmouseover]').length, number: number?.style.justifyContent, text: text?.style.justifyContent, chart: !!document.querySelector('[data-id="security-chart"] svg'), shape: !!document.querySelector('[data-id="security-shape"] svg') };
    });
    assert.deepEqual(before, { executed: 0, handlers: 0, number: 'flex-end', text: '', chart: true, shape: true });
    await page.locator('.obj[data-id="security-chart"]').click({ position: { x: 20, y: 20 } });
    await page.evaluate(() => window.tabula.run('chartSelectData'));
    const preview = page.locator('.sd-preview-section'); if (!await preview.evaluate(el => el.open)) await preview.locator('summary').click();
    await page.waitForSelector('.sd-preview svg');
    assert.equal(await page.locator('.dialog [onmouseover]').count(), 0);
    await page.keyboard.press('Escape');
  });

  await check('차트 요소 선택 강조와 해제 유지', async () => {
    await page.locator('.obj[data-id="security-chart"] [data-el="title"]').click();
    assert.equal(await page.locator('.obj[data-id="security-chart"] [data-el="title"]').evaluate((node) => node.style.outlineStyle), 'dashed');
    assert.equal(await page.locator('.obj[data-id="security-chart"] style').count(), 0);
    await page.locator('.obj[data-id="security-chart"]').click({ position: { x: 8, y: 240 } });
    assert.equal(await page.locator('.obj[data-id="security-chart"] [data-el="title"]').evaluate((node) => node.style.outlineStyle), '');
  });

  await check('인쇄 표·차트 경로에서도 이벤트 제거 및 일반 맞춤 보존', async () => {
    const result = await page.evaluate(() => {
      const print = window.print; window.print = () => {};
      try {
        window.tabula.run('print');
        const area = document.getElementById('printArea');
        for (const node of area.querySelectorAll('*')) node.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
        return { executed: window.__wixelXss, handlers: area.querySelectorAll('[onmouseover]').length, charts: area.querySelectorAll('svg').length, numericAlign: [...area.querySelectorAll('td')].find((td) => td.textContent === '456')?.style.textAlign };
      } finally { window.print = print; window.dispatchEvent(new Event('afterprint')); }
    });
    assert.equal(result.executed, 0); assert.equal(result.handlers, 0);
    assert.ok(result.charts >= 2); assert.equal(result.numericAlign, 'right');
    assert.equal(await page.locator('#printArea').textContent(), '');
  });

  await check('셀 하이퍼링크 javascript 차단과 HTTPS 링크 유지', async () => {
    await page.evaluate(() => {
      window.__wixelOpened = []; window.__originalOpen = window.open;
      window.open = (...args) => window.__wixelOpened.push(args);
      const t = window.tabula; const wb = t.wb();
      wb.transact(() => {
        wb.setCellData(0, 5, 0, { raw: '차단 링크', link: 'javascript:window.__wixelXss++' });
        wb.setCellData(0, 6, 0, { raw: '안전 링크', link: 'https://example.com/' });
      });
      t.gv().renderAll(); t.selectCell(4, 0);
    });
    try {
      for (const r of [5, 6]) {
        const box = await page.locator(`.c[data-r="${r}"][data-c="0"]`).boundingBox();
        assert.ok(box);
        await page.mouse.click(box.x + 20, box.y + box.height / 2);
        if (r === 5) assert.deepEqual(await page.evaluate(() => window.__wixelOpened), []);
      }
      assert.deepEqual(await page.evaluate(() => window.__wixelOpened), [['https://example.com/', '_blank', 'noopener,noreferrer']]);
      assert.equal(await page.evaluate(() => window.__wixelXss), 0);
    } finally { await page.evaluate(() => { window.open = window.__originalOpen; }); }
  });
  await check('페이지 런타임 오류 없음', async () => assert.deepEqual(errors, []));
  console.log(JSON.stringify({ tests: results.length, failed: results.filter((r) => !r.ok).length, results }, null, 2));
  if (results.some((r) => !r.ok)) process.exitCode = 1;
} finally { await browser.close(); }
