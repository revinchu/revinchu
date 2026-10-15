#!/usr/bin/env node
// WIXEL MCP 서버 (stdio, JSON-RPC 2.0, 의존성 없음) — AI 가 디스크의 엑셀 파일을 읽고 고치고 저장
//   Claude Code:     claude mcp add wixel -- node /경로/tabula/mcp.js
//   Claude Desktop:  { "mcpServers": { "wixel": { "command": "node", "args": ["/경로/tabula/mcp.js"] } } }
// 열기: .xlsx .xlsm .xlsb .xls .ods .csv .tsv .wixel(json) / 저장: .xlsx .xlsm .csv .tsv .ods .wixel
// 계산은 앱과 같은 수식 엔진(workbook.js · formula.js), 파일 읽기 · 쓰기도 같은 코드(xlsx.js · xls.js · ods.js)
import { readFile, writeFile } from 'node:fs/promises';
import { resolve, extname, basename } from 'node:path';
import { createInterface } from 'node:readline';
import { Workbook } from './src/workbook.js';
import { readXlsx, writeXlsx } from './src/xlsx.js';
import { readXls } from './src/xls.js';
import { readOds, writeOds } from './src/ods.js';
import { parseDelimited, toDelimited, guessDelimiter } from './src/csv.js';
import { formatValue, styleForCode } from './src/format.js';

const VERSION = '1.0.0';
const log = (...a) => process.stderr.write(`[wixel-mcp] ${a.join(' ')}\n`);

// ───────────── 열린 통합 문서 ─────────────
const books = new Map(); // 이름 → { wb, path, dirty }
let current = null;
const cur = () => {
  if (!current || !books.has(current)) { books.set('새 통합 문서', { wb: new Workbook(), path: null, dirty: false }); current = '새 통합 문서'; }
  return books.get(current);
};

// ───────────── 주소 ─────────────
const colName = (c) => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
const colNum = (s) => [...s.toUpperCase()].reduce((n, ch) => n * 26 + ch.charCodeAt(0) - 64, 0) - 1;
const addr = (r, c) => `${colName(c)}${r + 1}`;
function parseCell(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(String(a ?? '').trim());
  if (!m) throw new Error(`셀 주소를 읽을 수 없습니다: '${a}' (예: B3)`);
  return { r: Number(m[2]) - 1, c: colNum(m[1]) };
}
/** 'A1:C10' · 'A:C' · '3:5' · 'B2' → { r1, c1, r2, c2 } (없으면 사용 범위) */
function parseRange(wb, si, ref) {
  const used = wb.usedRange(si);
  if (!ref) return { r1: 0, c1: 0, r2: Math.max(0, used.rows - 1), c2: Math.max(0, used.cols - 1) };
  const s = String(ref).replace(/\$/g, '').trim();
  let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(s);
  if (m) { const a = { r: Number(m[2]) - 1, c: colNum(m[1]) }; const b = m[3] ? { r: Number(m[4]) - 1, c: colNum(m[3]) } : a; return { r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c), r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c) }; }
  m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(s);
  if (m) return { r1: 0, c1: Math.min(colNum(m[1]), colNum(m[2])), r2: Math.max(0, used.rows - 1), c2: Math.max(colNum(m[1]), colNum(m[2])) };
  m = /^(\d+):(\d+)$/.exec(s);
  if (m) return { r1: Number(m[1]) - 1, c1: 0, r2: Number(m[2]) - 1, c2: Math.max(0, used.cols - 1) };
  throw new Error(`범위를 읽을 수 없습니다: '${ref}' (예: A1:D20, B:D, 3:10)`);
}
function sheetOf(wb, sheet) {
  if (sheet == null || sheet === '') return 0;
  if (typeof sheet === 'number') { if (!wb.sheets[sheet]) throw new Error(`${sheet}번 시트가 없습니다`); return sheet; }
  const si = wb.sheetIndexByName(String(sheet));
  if (si < 0) throw new Error(`시트 '${sheet}' 이(가) 없습니다. 있는 시트: ${wb.sheets.map((s) => s.name).join(', ')}`);
  return si;
}
/** 셀 값 → JSON 값 */
const jsonVal = (v) => {
  if (v == null) return null;
  if (typeof v === 'number' || typeof v === 'string' || typeof v === 'boolean') return v;
  if (v.code) return v.code; // 오류 값
  if (v.rows) return v.rows; // 배열
  if (v.src || v.alt != null) return `[그림${v.alt ? `: ${v.alt}` : ''}]`;
  return String(v);
};

/** 화면에 보이는 글 (표시 형식 적용) */
const disp = (wb, si, r, c, v = wb.getValue(si, r, c)) => {
  if (v == null) return '';
  if (typeof v === 'object') return String(jsonVal(v));
  const f = formatValue(v, wb.styleAt(si, r, c));
  return typeof f === 'string' ? f : f?.text ?? String(v);
};

// ───────────── 파일 ─────────────
function kindOf(p) { const e = extname(p).slice(1).toLowerCase(); return e === 'tabula' || e === 'json' ? 'wixel' : e; }
async function loadBook(path) {
  const bytes = new Uint8Array(await readFile(path));
  const kind = kindOf(path);
  const ole = bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0;
  if (ole && (kind === 'xls' || kind === 'xlt' || kind === 'xlsx')) { const r = readXls(bytes); return { wb: new Workbook(r.data), warnings: r.warnings ?? [] }; }
  if (bytes[0] === 0x50 && bytes[1] === 0x4b && kind !== 'ods') { const r = readXlsx(bytes); return { wb: new Workbook(r.data), warnings: r.warnings ?? [] }; }
  if (kind === 'ods' || kind === 'fods') { const r = readOds(kind === 'fods' ? new TextDecoder().decode(bytes) : bytes); return { wb: new Workbook(r.data), warnings: r.warnings ?? [] }; }
  let text;
  if (bytes[0] === 0xff && bytes[1] === 0xfe) text = new TextDecoder('utf-16le').decode(bytes.subarray(2));
  else { try { text = new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch { text = new TextDecoder('euc-kr').decode(bytes); } }
  text = text.replace(/^\uFEFF/, '');
  if (kind === 'wixel') { const data = JSON.parse(text); return { wb: new Workbook(data.workbook ?? data), warnings: [] }; }
  const rows = parseDelimited(text, kind === 'tsv' ? '\t' : guessDelimiter(text));
  const wb = new Workbook({ sheets: [{ name: basename(path).replace(/\.[^.]+$/, '').slice(0, 31) || 'Sheet1', cells: {} }] });
  wb.transact(() => rows.forEach((row, r) => row.forEach((v, c) => { if (v !== '') wb.setInput(0, r, c, v); })));
  return { wb, warnings: [] };
}
function sheetRows(wb, si, { formulas = false } = {}) {
  const { rows, cols } = wb.usedRange(si);
  const out = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      const cell = wb.getCell(si, r, c);
      if (formulas && cell?.raw?.startsWith?.('=')) { row.push(cell.raw); continue; }
      const v = wb.getValue(si, r, c);
      row.push(disp(wb, si, r, c, v));
    }
    out.push(row);
  }
  return out;
}
async function saveBook(b, path) {
  const kind = kindOf(path);
  let out;
  if (kind === 'xlsx' || kind === 'xlsm') out = writeXlsx(b.wb);
  else if (kind === 'csv' || kind === 'tsv') out = `${String.fromCharCode(0xfeff)}${toDelimited(sheetRows(b.wb, 0), kind === 'tsv' ? '\t' : ',')}`;
  else if (kind === 'wixel') out = JSON.stringify({ workbook: b.wb.serialize() });
  else if (kind === 'ods') {
    const wb = b.wb;
    out = writeOds(wb.sheets.map((s, i) => ({ si: i, name: s.name, merges: s.merges ?? [], hidden: s.state === 'hidden' || s.state === 'veryHidden' })), {
      raw: (s, r, c) => wb.getRaw(s, r, c), value: (s, r, c) => wb.getValue(s, r, c), style: (s, r, c) => wb.styleAt(s, r, c),
      used: (s) => wb.usedRange(s), colWidth: (s, c) => wb.colWidth(s, c), rowHeight: (s, r) => wb.sheets[s].rowHeights?.[r] ?? null,
      display: (s, r, c) => disp(wb, s, r, c),
    });
  }
  else throw new Error(`${kind} 로는 저장할 수 없습니다 (.xlsx .xlsm .csv .tsv .ods .wixel)`);
  await writeFile(path, typeof out === 'string' ? out : Buffer.from(out));
  return kind;
}
function summary(b) {
  const { wb } = b;
  return {
    name: current, path: b.path,
    sheets: wb.sheets.slice(0, wb.ownSheetCount()).map((s, si) => {
      const u = wb.usedRange(si);
      return { name: s.name, used: u.rows && u.cols ? `A1:${addr(u.rows - 1, u.cols - 1)}` : '(빈 시트)', rows: u.rows, cols: u.cols, ...(s.hidden ? { hidden: s.hidden } : {}), ...(s.tables?.length ? { tables: s.tables.map((t) => `${t.name} ${addr(t.r1, t.c1)}:${addr(t.r2, t.c2)}`) } : {}), ...(s.pivot || s.pivotsExtra?.length ? { pivots: [s.pivot, ...(s.pivotsExtra ?? [])].filter(Boolean).map((p) => p.name) } : {}), ...(s.charts?.length ? { charts: s.charts.length } : {}) };
    }),
    names: wb.names.map((n) => `${n.name} = ${n.ref}`).slice(0, 100),
  };
}

// ───────────── 도구 ─────────────
const SHEET = { sheet: { type: 'string', description: '시트 이름 (생략하면 첫 시트)' } };
const TOOLS = [
  { name: 'open_workbook', description: '디스크의 엑셀 파일을 엽니다 (.xlsx .xlsm .xlsb .xls .ods .csv .tsv .wixel). 시트 · 사용 범위 · 표 · 피벗 · 이름을 돌려줍니다.', input: { path: { type: 'string' } }, required: ['path'],
    run: async (a) => {
      const path = resolve(a.path);
      const { wb, warnings } = await loadBook(path);
      const name = basename(path);
      books.set(name, { wb, path, dirty: false }); current = name;
      return { ...summary(books.get(name)), warnings };
    } },
  { name: 'new_workbook', description: '빈 통합 문서를 새로 만듭니다 (메모리, save_workbook 으로 저장).', input: { name: { type: 'string' } },
    run: async (a) => { const name = a.name || `통합 문서${books.size + 1}`; books.set(name, { wb: new Workbook(), path: null, dirty: true }); current = name; return summary(books.get(name)); } },
  { name: 'list_workbooks', description: '열린 통합 문서 목록과 지금 다루는 문서.', input: {}, run: async () => ({ current, books: [...books.entries()].map(([n, b]) => ({ name: n, path: b.path, dirty: b.dirty, sheets: b.wb.sheets.length })) }) },
  { name: 'switch_workbook', description: '다룰 통합 문서를 바꿉니다 (list_workbooks 의 이름).', input: { name: { type: 'string' } }, required: ['name'], run: async (a) => { if (!books.has(a.name)) throw new Error(`'${a.name}' 은(는) 열려 있지 않습니다`); current = a.name; return summary(books.get(a.name)); } },
  { name: 'workbook_info', description: '시트 목록 · 사용 범위 · 표 · 피벗 · 차트 수 · 이름 정의.', input: {}, run: async () => summary(cur()) },
  { name: 'read_range', description: '범위의 값을 읽습니다. 기본은 화면에 보이는 글(서식 적용), values=true 면 실제 값, formulas=true 면 수식 글. range 생략 = 사용 범위 (최대 2000칸씩, offset 으로 이어 읽기).', input: { ...SHEET, range: { type: 'string', description: '예: A1:D20, B:B, 3:10' }, values: { type: 'boolean' }, formulas: { type: 'boolean' }, offset_rows: { type: 'number' } },
    run: async (a) => {
      const { wb } = cur(); const si = sheetOf(wb, a.sheet);
      const g = parseRange(wb, si, a.range);
      const width = g.c2 - g.c1 + 1;
      const maxRows = Math.max(1, Math.floor(2000 / Math.max(1, width)));
      const r0 = g.r1 + (a.offset_rows ?? 0); const r9 = Math.min(g.r2, r0 + maxRows - 1);
      const rows = [];
      for (let r = r0; r <= r9; r++) {
        const row = [];
        for (let c = g.c1; c <= g.c2; c++) {
          const cell = wb.getCell(si, r, c);
          if (a.formulas && typeof cell?.raw === 'string' && cell.raw.startsWith('=')) { row.push(cell.raw); continue; }
          const v = wb.getValue(si, r, c);
          row.push(a.values ? jsonVal(v) : disp(wb, si, r, c, v));
        }
        rows.push(row);
      }
      return { sheet: wb.sheets[si].name, range: `${addr(r0, g.c1)}:${addr(r9, g.c2)}`, rows, ...(r9 < g.r2 ? { more: true, next_offset_rows: r9 - g.r1 + 1 } : {}) };
    } },
  { name: 'write_range', description: '셀에 값 · 수식을 씁니다. values 는 2차원 배열 (행마다 배열), start 칸부터 채움. "=" 로 시작하면 수식, 숫자 · 날짜 · 백분율 글은 엑셀처럼 해석. 빈 글은 그 칸을 지움.', input: { ...SHEET, start: { type: 'string', description: '왼쪽 위 칸 (예: A1)' }, values: { type: 'array', items: { type: 'array' } } }, required: ['start', 'values'],
    run: async (a) => {
      const b = cur(); const { wb } = b; const si = sheetOf(wb, a.sheet);
      const { r, c } = parseCell(a.start);
      let n = 0;
      wb.transact(() => {
        a.values.forEach((row, i) => (Array.isArray(row) ? row : [row]).forEach((v, j) => {
          const raw = v == null ? '' : typeof v === 'boolean' ? (v ? 'TRUE' : 'FALSE') : String(v);
          if (raw === '') wb.setCellData(si, r + i, c + j, null); else wb.setInput(si, r + i, c + j, raw);
          n++;
        }));
      });
      b.dirty = true;
      const h = a.values.length; const w = Math.max(1, ...a.values.map((x) => (Array.isArray(x) ? x.length : 1)));
      return { written: n, range: `${addr(r, c)}:${addr(r + h - 1, c + w - 1)}`, preview: sheetPreview(wb, si, r, c, r + Math.min(h, 10) - 1, c + Math.min(w, 10) - 1) };
    } },
  { name: 'clear_range', description: '범위를 지웁니다. what: contents(값 · 수식, 기본) · formats · all.', input: { ...SHEET, range: { type: 'string' }, what: { type: 'string', enum: ['contents', 'formats', 'all'] } }, required: ['range'],
    run: async (a) => { const b = cur(); const si = sheetOf(b.wb, a.sheet); const g = parseRange(b.wb, si, a.range); b.wb.transact(() => b.wb.clearRange(si, g.r1, g.c1, g.r2, g.c2, a.what ?? 'contents')); b.dirty = true; return `지웠습니다: ${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)}`; } },
  { name: 'set_format', description: '범위의 서식: bold · italic · color(글자색 #RRGGBB) · fill(채우기) · number_format(엑셀 표시 형식 코드, 예 "#,##0" "0.0%" "yyyy-mm-dd") · align(left|center|right) · wrap.', input: { ...SHEET, range: { type: 'string' }, bold: { type: 'boolean' }, italic: { type: 'boolean' }, color: { type: 'string' }, fill: { type: 'string' }, number_format: { type: 'string' }, align: { type: 'string' }, wrap: { type: 'boolean' } }, required: ['range'],
    run: async (a) => {
      const b = cur(); const si = sheetOf(b.wb, a.sheet); const g = parseRange(b.wb, si, a.range);
      const patch = {};
      for (const k of ['bold', 'italic', 'color', 'fill', 'wrap']) if (a[k] != null) patch[k] = a[k];
      if (a.align) patch.align = a.align;
      if (a.number_format) Object.assign(patch, styleForCode(a.number_format) ?? { numFmt: 'custom', code: a.number_format });
      if ((g.r2 - g.r1 + 1) * (g.c2 - g.c1 + 1) > 200000) throw new Error('한 번에 20만 칸까지 서식을 바꿀 수 있습니다');
      b.wb.transact(() => { for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) b.wb.setStyle(si, r, c, patch); });
      b.dirty = true;
      return `서식을 바꿨습니다: ${addr(g.r1, g.c1)}:${addr(g.r2, g.c2)} ${JSON.stringify(patch)}`;
    } },
  { name: 'evaluate', description: '수식을 계산만 하고 결과를 돌려줍니다 (칸에 쓰지 않음). 예: "=SUMIFS(C:C, A:A, \\"서울\\")", "=AVERAGE(Sheet2!B2:B100)".', input: { ...SHEET, formula: { type: 'string' } }, required: ['formula'],
    run: async (a) => {
      const { wb } = cur(); const si = sheetOf(wb, a.sheet);
      const f = String(a.formula).startsWith('=') ? a.formula : `=${a.formula}`;
      const ext = wb.extent(si); const r = Math.max(ext.rows + 5, 1048000); const c = 0;
      const saveUndo = wb.noUndo; wb.noUndo = true;
      try {
        wb.transact(() => wb.setInput(si, r, c, f));
        const v = wb.getValue(si, r, c);
        return { formula: f, value: jsonVal(v), text: disp(wb, si, r, c, v) };
      } finally { wb.transact(() => wb.setCellData(si, r, c, null)); wb.noUndo = saveUndo; }
    } },
  { name: 'find', description: '값 · 수식에서 글을 찾습니다 (대소문자 무시). 결과는 칸 주소와 값 (최대 200개).', input: { text: { type: 'string' }, ...SHEET, in_formulas: { type: 'boolean' } }, required: ['text'],
    run: async (a) => {
      const { wb } = cur(); const q = String(a.text).toLowerCase(); const hits = [];
      const sheets = a.sheet ? [sheetOf(wb, a.sheet)] : wb.sheets.slice(0, wb.ownSheetCount()).map((_, i) => i);
      for (const si of sheets) {
        const u = wb.usedRange(si);
        for (let r = 0; r < u.rows && hits.length < 200; r++) for (let c = 0; c < u.cols && hits.length < 200; c++) {
          const cell = wb.getCell(si, r, c);
          const v = wb.getValue(si, r, c);
          const txt = disp(wb, si, r, c, v);
          if (txt.toLowerCase().includes(q) || (a.in_formulas && typeof cell?.raw === 'string' && cell.raw.toLowerCase().includes(q))) hits.push({ sheet: wb.sheets[si].name, cell: addr(r, c), value: txt, ...(cell?.raw?.startsWith?.('=') ? { formula: cell.raw } : {}) });
        }
      }
      return { count: hits.length, hits };
    } },
  { name: 'sort_range', description: '범위를 한 열 기준으로 정렬합니다 (머리글 행이 있으면 header=true).', input: { ...SHEET, range: { type: 'string' }, by: { type: 'string', description: '기준 열 (예: C)' }, descending: { type: 'boolean' }, header: { type: 'boolean' } }, required: ['range', 'by'],
    run: async (a) => {
      const b = cur(); const si = sheetOf(b.wb, a.sheet); const g = parseRange(b.wb, si, a.range);
      const r1 = g.r1 + (a.header ? 1 : 0);
      b.wb.transact(() => b.wb.sortRange(si, r1, g.c1, g.r2, g.c2, colNum(a.by), !a.descending));
      b.dirty = true;
      return { sorted: `${addr(r1, g.c1)}:${addr(g.r2, g.c2)}`, preview: sheetPreview(b.wb, si, g.r1, g.c1, Math.min(g.r2, g.r1 + 9), g.c2) };
    } },
  { name: 'add_sheet', description: '시트를 더합니다.', input: { name: { type: 'string' } }, run: async (a) => { const b = cur(); let at; b.wb.transact(() => { at = b.wb.addSheet(a.name ?? ''); }); b.dirty = true; return `시트를 더했습니다: ${b.wb.sheets[at].name}`; } },
  { name: 'rename_sheet', description: '시트 이름을 바꿉니다 (수식의 참조도 함께).', input: { sheet: { type: 'string' }, name: { type: 'string' } }, required: ['sheet', 'name'], run: async (a) => { const b = cur(); const si = sheetOf(b.wb, a.sheet); let ok; b.wb.transact(() => { ok = b.wb.renameSheet(si, a.name); }); if (!ok) throw new Error('쓸 수 없는 이름이거나 이미 있는 이름입니다 (31자 이내, \\ / ? * [ ] : 제외)'); b.dirty = true; return `이름을 바꿨습니다: ${a.sheet} → ${a.name}`; } },
  { name: 'delete_sheet', description: '시트를 지웁니다 (마지막 시트는 지울 수 없음).', input: { sheet: { type: 'string' } }, required: ['sheet'], run: async (a) => { const b = cur(); const si = sheetOf(b.wb, a.sheet); let ok; b.wb.transact(() => { ok = b.wb.deleteSheet(si); }); if (!ok) throw new Error('마지막 시트는 지울 수 없습니다'); b.dirty = true; return `지웠습니다: ${a.sheet}`; } },
  { name: 'save_workbook', description: '저장합니다. path 를 주면 그 이름으로 (확장자로 형식: .xlsx .xlsm .csv .tsv .ods .wixel — csv/tsv 는 첫 시트만). 없으면 연 파일에 덮어씀 (xlsx · xlsm · csv · ods 일 때).', input: { path: { type: 'string' } },
    run: async (a) => {
      const b = cur();
      const path = a.path ? resolve(a.path) : b.path;
      if (!path) throw new Error('새 통합 문서는 path 를 주세요 (예: 결과.xlsx)');
      if (!a.path && ['xls', 'xlsb'].includes(kindOf(path))) throw new Error('.xls · .xlsb 에는 덮어쓸 수 없습니다. path 로 .xlsx 이름을 주세요');
      const kind = await saveBook(b, path);
      b.dirty = false; if (a.path) b.path = path;
      return `저장했습니다: ${path} (${kind})`;
    } },
  { name: 'convert', description: '형식 바꾸기: input(.xlsx .xlsm .xlsb .xls .ods .csv .tsv) → output(.xlsx .csv .tsv .ods .wixel). 예: 옛 엑셀(.xls) → .xlsx, 엑셀 → CSV.', input: { input: { type: 'string' }, output: { type: 'string' } }, required: ['input', 'output'],
    run: async (a) => { const { wb, warnings } = await loadBook(resolve(a.input)); const kind = await saveBook({ wb }, resolve(a.output)); return { output: resolve(a.output), to: kind, sheets: wb.sheets.length, warnings }; } },
];
function sheetPreview(wb, si, r1, c1, r2, c2) {
  const rows = [];
  for (let r = r1; r <= r2; r++) { const row = []; for (let c = c1; c <= c2; c++) { const v = wb.getValue(si, r, c); row.push(disp(wb, si, r, c, v)); } rows.push(row); }
  return rows;
}
const BY_NAME = Object.fromEntries(TOOLS.map((t) => [t.name, t]));
const toolSchema = (t) => ({ name: t.name, description: t.description, inputSchema: { type: 'object', properties: t.input ?? {}, ...(t.required ? { required: t.required } : {}) } });

// ───────────── JSON-RPC (stdio, 줄 단위) ─────────────
const send = (msg) => process.stdout.write(`${JSON.stringify(msg)}\n`);
const ok = (id, result) => send({ jsonrpc: '2.0', id, result });
const fail = (id, code, message) => send({ jsonrpc: '2.0', id, error: { code, message } });
const text = (v) => (typeof v === 'string' ? v : JSON.stringify(v, null, 2));
async function handle(msg) {
  const { id, method, params } = msg;
  switch (method) {
    case 'initialize':
      return ok(id, {
        protocolVersion: params?.protocolVersion ?? '2024-11-05',
        capabilities: { tools: { listChanged: false } },
        serverInfo: { name: 'wixel', version: VERSION },
        instructions: 'WIXEL 스프레드시트 (엑셀 호환). open_workbook → workbook_info · read_range 로 읽기 → write_range(값 · "=" 수식) · set_format · sort_range · evaluate → save_workbook. 수식은 엑셀 함수 400여 개(동적 배열 · XLOOKUP · LET · LAMBDA 포함)를 계산합니다.',
      });
    case 'notifications/initialized': case 'notifications/cancelled': return null;
    case 'ping': return ok(id, {});
    case 'tools/list': return ok(id, { tools: TOOLS.map(toolSchema) });
    case 'tools/call': {
      const t = BY_NAME[params?.name];
      if (!t) return fail(id, -32602, `알 수 없는 도구: ${params?.name}`);
      try { return ok(id, { content: [{ type: 'text', text: text(await t.run(params.arguments ?? {})) }] }); } catch (e) { return ok(id, { content: [{ type: 'text', text: `오류: ${e.message}` }], isError: true }); }
    }
    case 'resources/list': return ok(id, { resources: [] });
    case 'prompts/list': return ok(id, { prompts: [] });
    default: if (id !== undefined) return fail(id, -32601, `지원하지 않는 메서드: ${method}`); return null;
  }
}
const pending = new Set();
const rl = createInterface({ input: process.stdin, crlfDelay: Infinity });
rl.on('line', (line) => {
  if (!line.trim()) return;
  let msg;
  try { msg = JSON.parse(line); } catch { send({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'JSON 파싱 오류' } }); return; }
  const p = Promise.resolve(handle(msg)).catch((e) => { log('오류', e.stack ?? e); if (msg.id !== undefined) fail(msg.id, -32603, e.message); });
  pending.add(p); p.finally(() => pending.delete(p));
});
rl.on('close', () => { Promise.allSettled([...pending]).then(() => process.exit(0)); });
log('시작');
