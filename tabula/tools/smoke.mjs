// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch();
try {
const page = await browser.newPage({ viewport: { width: 1500, height: 900 } });
let errs = [];
page.on('pageerror', (e) => errs.push('pageerror: ' + e.message + ' @' + (e.stack ?? '').split('\n')[1]?.trim()));
page.on('console', (m) => { if (m.type() === 'error') errs.push('console: ' + m.text().slice(0, 200)); });
page.on('dialog', (d) => d.dismiss());
// 번들 정적 서버에도 문서 API를 연결하지 않고 같은 기능 가용성으로 검사한다.
await page.route('**/api/health', route => route.fulfill({ json: { ok: true, vault: false, auth: false, publish: false } }));
await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; window.showSaveFilePicker = undefined; window.showOpenFilePicker = undefined; });
await page.goto((process.env.WIXEL_URL || 'http://localhost:5178/')); await page.evaluate(() => localStorage.clear()); await page.reload();
await page.waitForFunction(() => typeof window.tabula?.commands === 'function');
await page.waitForTimeout(800);
const skip = new Set(['open', 'openFile', 'print', 'newWorkbook', 'insertPicture', 'insertPictureInCell', 'importCsv', 'pickFile', 'closeDoc', 'close', 'quit']);
const cmds = (await page.evaluate(() => window.tabula.commands())).filter((c) => !skip.has(c));
console.log('commands', cmds.length);
const setup = async () => page.evaluate(() => {
  const t = window.tabula; const wb = t.wb();
  document.querySelectorAll('.dialog-backdrop, .menu, .pat-pop').forEach((e) => e.remove());
  if (wb.sheets[0].cells.size < 5) {
    const rows = [['날짜', '채널', '비용', '매출'], ['2026-01-01', '네이버', 100, 300], ['2026-01-02', '구글', 200, 500], ['2026-01-03', '메타', 150, 200], ['2026-01-04', '네이버', 120, 330]];
    wb.transact(() => rows.forEach((r, i) => r.forEach((v, c) => wb.setCellData(0, i, c, { raw: String(v) }))));
  }
  t.switchSheet(0);
  t.selectRange({ r1: 0, c1: 0, r2: 4, c2: 3 }, 'cells', { r: 1, c: 1 });
});
const bad = errs.length ? [['시작', errs.join(' | ')]] : [];
for (const c of cmds) {
  errs = [];
  try {
    await setup();
    await page.evaluate((c) => { try { const r = window.tabula.run(c); if (r && r.catch) r.catch((e) => console.error('async ' + e.message)); } catch (e) { console.error('throw ' + e.message); } }, c);
    await page.waitForTimeout(120);
    await page.keyboard.press('Escape');
  } catch (e) { errs.push('pw: ' + e.message.slice(0, 120)); }
  if (errs.length) bad.push([c, errs.slice(0, 2).join(' | ')]);
}
for (const [c, e] of bad) console.log(c, '::', e.slice(0, 300));
console.log('bad', bad.length);
if (bad.length) process.exitCode = 1;
} finally {
  await browser.close();
}
