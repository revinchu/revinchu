// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
// 피벗 영역: 파일에 저장된 엑셀 값 vs 위셀이 다시 그린 값
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
import { readFileSync } from 'node:fs';
import { readXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';
const file = process.argv[2];
const A = new Workbook(readXlsx(new Uint8Array(readFileSync(file))).data);
const orig = {};
A.sheets.forEach((s, si) => { for (const d of [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean)) orig[`${si}|${d.name}`] = true; });
const browser = await chromium.launch({ args: ['--js-flags=--max-old-space-size=8192'] });
const page = await browser.newPage();
await page.goto((process.env.WIXEL_URL || 'http://localhost:5178/')); await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForTimeout(300);
await page.setInputFiles('#fileInput', file);
await page.waitForTimeout(800);
await page.waitForFunction(() => !document.querySelector('.load-progress'), null, { timeout: 900000, polling: 200 });
const areas = await page.evaluate(() => { const wb = window.tabula.wb(); const out = []; wb.sheets.forEach((s, si) => { for (const d of [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean)) { if (!d.area) continue; const a = d.area; const vals = []; for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) { const v = wb.getValue(si, r, c); vals.push([r, c, v && typeof v === 'object' ? (v.code ?? String(v)) : v ?? null]); } out.push({ si, sheet: s.name, name: d.name, area: a, vals }); } }); return out; });
await browser.close();
let bad = 0;
const norm = (v) => (v === '' ? null : v);
for (const p of areas) {
  let n = 0; let ex = '';
  for (const [r, c, v] of p.vals) {
    let o = A.getValue(p.si, r, c); o = o && typeof o === 'object' ? (o.code ?? String(o)) : o ?? null;
    const a = norm(v), b = norm(o);
    const same = a === b || (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b))) || (typeof a === 'number' && typeof b === 'string' && String(a) === b);
    if (!same) { n++; if (!ex) ex = `r${r + 1}c${c + 1} wixel=${JSON.stringify(a)} excel=${JSON.stringify(b)}`; }
  }
  if (n) { bad++; console.log(`${p.sheet} / ${p.name} ${JSON.stringify(p.area)}: ${n} cells differ  e.g. ${ex}`); }
}
console.log('pivots', areas.length, 'differing', bad);
