// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
// 피벗 영역: 파일에 저장된 엑셀 값 vs 위셀이 다시 그린 값
import { readFileSync } from 'node:fs';
import { extname, isAbsolute } from 'node:path';
import { pathToFileURL } from 'node:url';
import { readXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';
const file = process.argv[2];
if (process.argv.length !== 3 || !file) {
  console.error('사용법: node tools/pvcmp.mjs "파일.xlsx" (xlsx/xlsm/xlsb)');
  process.exit(2);
}
if (!['.xlsx', '.xlsm', '.xlsb'].includes(extname(file).toLowerCase())) {
  console.error('지원하는 형식은 xlsx/xlsm/xlsb입니다. 옛 .xls 파일은 tools/brcheck.mjs로 검사하세요.');
  process.exit(2);
}
let browser;
try {
  const A = new Workbook(readXlsx(new Uint8Array(readFileSync(file))).data);
  const orig = new Map();
  const key = (sheet, name) => JSON.stringify([sheet, name]);
  A.sheets.forEach((s, si) => { for (const d of [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean)) orig.set(key(s.name, d.name), { si, sheet: s.name, name: d.name, area: d.area }); });
  const module = process.env.PLAYWRIGHT_MODULE || 'playwright';
  const { chromium } = await import(isAbsolute(module) ? pathToFileURL(module).href : module);
  browser = await chromium.launch({ args: ['--js-flags=--max-old-space-size=8192'] });
  const page = await browser.newPage();
  const errors = [];
  page.on('pageerror', (e) => errors.push(e.message));
  await page.addInitScript(() => { window.TABULA_STATIC = true; window.WIXEL_SKIP_START = true; });
  await page.goto((process.env.WIXEL_URL || 'http://localhost:5178/')); await page.evaluate(() => localStorage.clear()); await page.reload(); await page.waitForTimeout(300);
  await page.setInputFiles('#fileInput', file);
  await page.waitForTimeout(800);
  await page.waitForFunction(() => !document.querySelector('.load-progress'), null, { timeout: 900000, polling: 200 });
  const areas = await page.evaluate(() => { const wb = window.tabula.wb(); const out = []; wb.sheets.forEach((s, si) => { for (const d of [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean)) { if (!d.area) continue; const a = d.area; const vals = []; for (let r = a.r1; r <= a.r2; r++) for (let c = a.c1; c <= a.c2; c++) { const v = wb.getValue(si, r, c); vals.push([r, c, v && typeof v === 'object' ? (v.code ?? String(v)) : v ?? null]); } out.push({ si, sheet: s.name, name: d.name, area: a, vals }); } }); return out; });
  let bad = 0;
  const norm = (v) => (v === '' ? null : v);
  const seen = new Set();
  for (const p of areas) {
    const id = key(p.sheet, p.name), source = orig.get(id);
    seen.add(id);
    if (!source) { bad++; console.log(`${p.sheet} / ${p.name}: 원본에 없는 피벗입니다.`); continue; }
    let n = 0; let ex = '';
    const areaChanged = ['r1', 'c1', 'r2', 'c2'].some((k) => source.area?.[k] !== p.area[k]);
    for (const [r, c, v] of p.vals) {
      let o = A.getValue(source.si, r, c); o = o && typeof o === 'object' ? (o.code ?? String(o)) : o ?? null;
      const a = norm(v), b = norm(o);
      const same = a === b || (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 1e-6 * Math.max(1, Math.abs(b))) || (typeof a === 'number' && typeof b === 'string' && String(a) === b);
      if (!same) { n++; if (!ex) ex = `r${r + 1}c${c + 1} wixel=${JSON.stringify(a)} excel=${JSON.stringify(b)}`; }
    }
    if (n || areaChanged) { bad++; console.log(`${p.sheet} / ${p.name} ${JSON.stringify(p.area)}: ${n} cells differ${n ? `  e.g. ${ex}` : ''}${areaChanged ? ` / 피벗 영역 변경: ${JSON.stringify(source.area)} → ${JSON.stringify(p.area)}` : ''}`); }
  }
  for (const [id, p] of orig) if (!seen.has(id)) { bad++; console.log(`${p.sheet} / ${p.name}: 피벗 또는 출력 영역이 누락되었습니다.`); }
  console.log('pivots', areas.length, 'differing', bad);
  if (errors.length) throw new Error(`브라우저 오류: ${errors.join(' | ')}`);
  process.exitCode = bad ? 1 : 0;
} catch (e) {
  console.error(`검사를 완료하지 못했습니다: ${e.message}`);
  process.exitCode = 2;
} finally {
  if (browser) await browser.close();
}
