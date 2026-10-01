// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
import { statSync } from 'node:fs';
import { extname } from 'node:path';
const file = process.argv[2];
if (!file || !['.xlsx', '.xlsm', '.xlsb', '.xls'].includes(extname(file).toLowerCase())) {
  console.error('사용법: node tools/brcheck.mjs 파일.xlsx (xlsx/xlsm/xlsb/xls 지원)');
  process.exit(2);
}
try {
  if (!statSync(file).isFile()) throw new Error('일반 파일이 아닙니다.');
} catch (e) {
  console.error('입력 파일을 확인하세요: ' + e.message);
  process.exit(2);
}
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser = await chromium.launch({ args: ['--js-flags=--max-old-space-size=8192'] });
try {
const page = await browser.newPage({ viewport: { width: 1500, height: 950 } });
const errs = [];
page.on('pageerror', (e) => errs.push('pageerror: ' + e.message));
await page.addInitScript(() => { window.TABULA_STATIC = true; });
await page.goto((process.env.WIXEL_URL || 'http://localhost:5178/')); await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForTimeout(300);
await page.waitForFunction(() => typeof window.tabula?.wb === 'function');
const t0 = Date.now();
await page.setInputFiles('#fileInput', file);
await page.waitForTimeout(800);
await page.waitForFunction(() => !document.querySelector('.load-progress'), null, { timeout: 900000, polling: 200 });
console.log('loaded', Date.now() - t0, 'ms');
const out = await page.evaluate(() => {
  const wb = window.tabula.wb();
  const saved = [];
  wb.sheets.forEach((s, si) => s.cells.forEachRC((c, r, cc) => { if (c.formula && c.cached !== undefined) saved.push([si, r, cc, c.cached]); }));
  wb.invalidate();
  const norm = (v) => (v && typeof v === 'object' ? (v.code ?? v.error ?? String(v)) : v ?? '');
  const same = (a, b) => a === b || (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)));
  const bad = new Map();
  for (const [si, r, c, cached] of saved) {
    let v; try { v = wb.getValue(si, r, c); } catch (e) { v = 'THROW ' + e.message; }
    const a = norm(v), b = norm(cached);
    if (!same(a, b)) { const raw = wb.getCell(si, r, c).raw; const k = wb.sheets[si].name + ' | ' + (raw.match(/[A-Z][A-Z0-9.]+(?=\()/g) ?? ['-']).join(','); const e = bad.get(k) ?? { n: 0, ex: `r${r + 1}c${c + 1} ${raw.slice(0, 110)} → app=${JSON.stringify(a)} xl=${JSON.stringify(b)}` }; e.n++; bad.set(k, e); }
  }
  return { total: [...bad.values()].reduce((x, e) => x + e.n, 0), list: [...bad].sort((x, y) => y[1].n - x[1].n).slice(0, 20).map(([k, e]) => `${e.n} ${k}\n     ${e.ex}`) };
});
console.log('formula mismatches (browser):', out.total); console.log(out.list.join('\n'));
console.log(errs.join('\n') || 'no errors');
if (out.total || errs.length) process.exitCode = 1;
} finally {
  await browser.close();
}
