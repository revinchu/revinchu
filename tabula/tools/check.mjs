// WIXEL 검증 도구 (tools/README.md 참고). Playwright 가 필요한 스크립트는 `npm start` 로 서버를 먼저 띄우세요.
import { readFileSync } from 'node:fs';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { Workbook } from '../src/workbook.js';
const file = process.argv[2];
let t = Date.now();
const res = readXlsx(new Uint8Array(readFileSync(file)));
const A = new Workbook(res.data);
console.log(`== ${file} read ${Date.now() - t}ms`, res.warnings?.length ? 'warnings: ' + JSON.stringify(res.warnings).slice(0, 300) : '');
const norm = (v) => (v && typeof v === 'object' ? (v.code ?? v.error ?? String(v)) : v ?? '');
const same = (a, b) => a === b || (typeof a === 'number' && typeof b === 'number' && Math.abs(a - b) <= 1e-9 * Math.max(1, Math.abs(b)));
let cells = 0, forms = 0;
const saved = [];
A.sheets.forEach((s, si) => s.cells.forEachRC((c, r, cc) => { cells++; if (c.formula) { forms++; saved.push([si, r, cc, c.cached]); } }));
const summary = A.sheets.map((s) => `${s.name}[${s.cells.size}c` + (s.charts?.length ? ` ch${s.charts.length}` : '') + (s.images?.length ? ` img${s.images.length}` : '') + (s.shapes?.length ? ` shp${s.shapes.length}` : '') + (s.cond?.length ? ` cf${s.cond.length}` : '') + (s.validations?.length ? ` dv${s.validations.length}` : '') + (s.merges?.length ? ` mg${s.merges.length}` : '') + (s.tables?.length ? ` tb${s.tables.length}` : '') + ([s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean).length ? ` pv` : '') + (s.state ? ` ${s.state}` : '') + ']');
console.log('cells', cells, 'formulas', forms, 'names', A.names.length, '|', summary.join(' '));
const unsup = new Map(); A.sheets.forEach((s) => s.cells.forEach((c) => { if (c.formula && c.parseError) unsup.set('PARSE:' + c.raw.slice(0, 60), 1); }));
A.invalidate();
const bad = new Map();
for (const [si, r, c, cached] of saved) {
  if (cached === undefined) continue;
  let v; try { v = A.getValue(si, r, c); } catch (e) { v = 'THROW ' + e.message; }
  const a = norm(v), b = norm(cached);
  if (!same(a, b)) { const raw = A.getCell(si, r, c).raw; const fn = (raw.match(/[A-Z][A-Z0-9.]+(?=\()/g) ?? ['-']).join(','); const k = A.sheets[si].name + ' | ' + fn; const e = bad.get(k) ?? { n: 0, ex: `r${r + 1}c${c + 1} ${raw.slice(0, 100)} → app=${JSON.stringify(a)} xl=${JSON.stringify(b)}` }; e.n++; bad.set(k, e); }
}
console.log('formula mismatches:', [...bad.values()].reduce((x, e) => x + e.n, 0));
for (const [k, e] of [...bad].sort((x, y) => y[1].n - x[1].n).slice(0, 15)) console.log('  ', e.n, k, '\n      ', e.ex);
for (const k of unsup.keys()) console.log('  ', k);
// 왕복
const A2 = new Workbook(res.data === undefined ? null : readXlsx(new Uint8Array(readFileSync(file))).data);
t = Date.now(); const out = writeXlsx(A2); const B = new Workbook(readXlsx(out).data);
console.log(`roundtrip write+read ${Date.now() - t}ms`);
const rbad = new Map(); const add = (k, ex) => { const e = rbad.get(k) ?? { n: 0, ex }; e.n++; rbad.set(k, e); };
const js = (x) => JSON.stringify(x ?? null);
A2.sheets.forEach((s, si) => {
  const bi = B.sheets.findIndex((x) => x.name === s.name); if (bi < 0) { add('missing sheet', s.name); return; }
  const b = B.sheets[bi];
  s.cells.forEachRC((cell, r, c) => {
    const cb = b.cells.getRC(r, c);
    const va = norm(A2.getValue(si, r, c)), vb = norm(B.getValue(bi, r, c));
    if (!same(va, vb)) add(s.name + ' value', `r${r + 1}c${c + 1} ${js(va)} vs ${js(vb)} raw=${cell.raw.slice(0, 50)}`);
    if (cell.formula && cb?.raw !== cell.raw) add(s.name + ' formula', `r${r + 1}c${c + 1} ${cell.raw.slice(0, 60)} vs ${cb?.raw?.slice(0, 60)}`);
    const srt = (o) => JSON.stringify(Object.fromEntries(Object.entries(o ?? {}).filter(([, v]) => v !== undefined && v !== null).sort()));const sa = srt(A2.styleAt(si, r, c)), sb = srt(B.styleAt(bi, r, c));
    if (sa !== sb) add(s.name + ' style', `r${r + 1}c${c + 1} ${sa.slice(0, 160)} vs ${sb.slice(0, 160)}`);
    if ((cell.comment ?? null) && js(cell.comment) !== js(cb?.comment)) add(s.name + ' comment', `r${r + 1}c${c + 1}`);
    if ((cell.link ?? null) && cell.link !== cb?.link) add(s.name + ' link', `r${r + 1}c${c + 1} ${cell.link} vs ${cb?.link}`);
  });
  for (const p of ['merges', 'charts', 'images', 'shapes', 'cond', 'validations', 'tables', 'freeze', 'colWidths', 'rowHeights', 'hiddenRows', 'hiddenCols', 'tabColor', 'noGrid', 'page', 'outline', 'protect', 'filter', 'sparklines']) {
    const x = s[p], y = b[p];
    const len = (v) => (Array.isArray(v) ? v.length : v && typeof v === 'object' ? Object.keys(v).length : v ? 1 : 0);
    if (len(x) !== len(y)) add(s.name + ' prop ' + p, `${len(x)} vs ${len(y)}`);
    else if (['colWidths', 'rowHeights', 'merges', 'freeze', 'tabColor', 'noGrid'].includes(p) && js(x) !== js(y)) add(s.name + ' prop ' + p, `${js(x).slice(0, 120)} vs ${js(y).slice(0, 120)}`);
  }
});
if (A2.names.length !== B.names.length) add('names', `${A2.names.length} vs ${B.names.length}`);
console.log('roundtrip issues:', [...rbad.values()].reduce((x, e) => x + e.n, 0));
for (const [k, e] of [...rbad].sort((x, y) => y[1].n - x[1].n).slice(0, 25)) console.log('  ', e.n, k, '|', e.ex);
