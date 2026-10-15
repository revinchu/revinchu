#!/usr/bin/env node
// WIPOINT MCP 서버 (stdio, JSON-RPC 2.0, 의존성 없음) — AI 가 디스크의 발표 자료를 읽고 고치고 만들고 저장
//   Claude Code:     claude mcp add wipoint -- node /경로/wipoint/mcp.js
//   Claude Desktop:  { "mcpServers": { "wipoint": { "command": "node", "args": ["/경로/wipoint/mcp.js"] } } }
// 열기: .pptx .ppsx .potx .pptm .ppt .pdf .wpt(json) / 저장: .pptx .wpt / 개요(마크다운)로 새 발표 자료 만들기
import { readFile, writeFile } from 'node:fs/promises';
import { resolve, extname, basename } from 'node:path';
import { createInterface } from 'node:readline';
import { readPptx, writePptx } from './src/pptx.js';
import { readPpt } from './src/ppt.js';
import { readPdf } from './src/pdf.js';
import { newPresentation, newSlide, duplicateSlide, plainText, setPlainText, slideTitle, replaceAll, findAll, LAYOUTS, para, run, newTable, validatePresentation } from './src/model.js';
import { THEMES } from './src/themes.js';

const VERSION = '1.0.0';
const log = (...a) => process.stderr.write(`[wipoint-mcp] ${a.join(' ')}\n`);

const decks = new Map(); // 이름 → { pres, path, dirty }
let current = null;
const cur = () => {
  if (!current || !decks.has(current)) { decks.set('새 프레젠테이션', { pres: newPresentation(), path: null, dirty: true }); current = '새 프레젠테이션'; }
  return decks.get(current);
};
const kindOf = (p) => extname(p).slice(1).toLowerCase();
const TITLE = (o) => o.ph === 'title' || o.ph === 'ctrTitle';
const BODY = (o) => o.text && o.ph && !TITLE(o) && !['dt', 'ftr', 'sldNum'].includes(o.ph);

async function loadDeck(path) {
  const bytes = new Uint8Array(await readFile(path));
  const k = kindOf(path);
  if (bytes[0] === 0x25 && bytes[1] === 0x50 && bytes[2] === 0x44 && bytes[3] === 0x46) { const r = readPdf(bytes, {}); return { pres: r.pres, warnings: ['PDF 를 슬라이드로 바꾸어 열었습니다 (그림은 일부만).', ...(r.warnings ?? [])] }; }
  if (bytes[0] === 0xd0 && bytes[1] === 0xcf) { const r = readPpt(bytes); return { pres: r.pres, warnings: r.warnings ?? [] }; }
  if (bytes[0] === 0x50 && bytes[1] === 0x4b) { const r = readPptx(bytes); return { pres: r.pres, warnings: r.warnings ?? [] }; }
  if (k === 'wpt' || k === 'json') return { pres: validatePresentation(JSON.parse(new TextDecoder().decode(bytes))), warnings: [] };
  throw new Error('발표 자료 파일이 아닙니다 (.pptx .ppt .pdf .wpt)');
}
/** 슬라이드 번호 (1부터) → 슬라이드 */
function slideAt(pres, n) {
  const i = Number(n) - 1;
  if (!Number.isInteger(i) || !pres.slides[i]) throw new Error(`${n}번 슬라이드가 없습니다 (전체 ${pres.slides.length}장)`);
  return pres.slides[i];
}
/** 개체 고르기: 'title' · 'body' · 'body2' · 개체 id · 번호(1부터) */
function objectAt(slide, which = 'body') {
  if (which === 'title') return slide.objects.find(TITLE) ?? null;
  const m = /^body(\d*)$/.exec(String(which));
  if (m) return slide.objects.filter(BODY)[Math.max(0, Number(m[1] || 1) - 1)] ?? null;
  if (/^\d+$/.test(String(which))) return slide.objects[Number(which) - 1] ?? null;
  return slide.objects.find((o) => o.id === which || o.name === which) ?? null;
}
/** 줄 → 단락 (앞의 탭 · 공백 2칸 · '- ' 단계 = 수준) — 첫 단락 서식을 이어받음 */
function setLines(body, lines) {
  const tpl = body.paras[0] ?? para('');
  const r0 = tpl.runs[0] ?? {};
  const { t: _t, ...rp } = r0;
  body.paras = lines.map((line) => {
    let s = String(line); let lvl = 0;
    const m = /^((?:\t| {2})*)(?:[-*•]\s+)?/.exec(s);
    lvl = (m[1].match(/\t| {2}/g) ?? []).length;
    s = s.slice(m[0].length);
    return { ...JSON.parse(JSON.stringify({ ...tpl, runs: [] })), lvl: Math.min(8, lvl), runs: s ? [run(s, JSON.parse(JSON.stringify(rp)))] : [] };
  });
  if (!body.paras.length) body.paras = [para('')];
}
function slideInfo(s, i, full = false) {
  const objs = s.objects.map((o, k) => {
    const base = { n: k + 1, id: o.id, type: o.type, ...(o.ph ? { placeholder: o.ph } : {}), ...(o.name ? { name: o.name } : {}) };
    if (o.text) base.text = full ? o.text.paras.map((p) => `${'  '.repeat(p.lvl ?? 0)}${p.runs.map((r) => r.t).join('')}`).join('\n') : plainText(o.text).slice(0, 200);
    if (o.type === 'table') base.table = o.rows.map((r) => r.cells.map((c) => plainText(c.text)));
    if (o.type === 'chart') base.chart = { kind: o.chart?.type, series: (o.chart?.series ?? []).map((x) => x.name) };
    return base;
  });
  return { slide: i + 1, layout: s.layout, title: slideTitle(s), ...(s.hidden ? { hidden: true } : {}), objects: objs, ...(s.notes ? { notes: full ? s.notes : s.notes.slice(0, 200) } : {}) };
}
function summary(d) {
  const p = d.pres;
  return { name: current, path: d.path, size: p.size, theme: p.theme?.name, slides: p.slides.map((s, i) => ({ slide: i + 1, layout: s.layout, title: slideTitle(s) })) };
}
/** 마크다운 개요 → 슬라이드: '# 제목'(표지) · '## 제목'(슬라이드) · '- 글머리' (들여쓰기 = 수준) · '> 메모' · '| 표 |' */
function deckFromOutline(md, { theme = 'office' } = {}) {
  const pres = newPresentation({ theme, firstLayout: null });
  let s = null; let body = []; let notes = []; let table = [];
  const flush = () => {
    if (!s) return;
    const b = s.objects.find(BODY);
    if (table.length && b) {
      // 표: 본문 자리에 표 개체
      const rows = table.filter((r) => !/^\s*\|?\s*:?-{2,}/.test(r)).map((r) => r.replace(/^\s*\||\|\s*$/g, '').split('|').map((x) => x.trim()));
      const t = newTable(rows.length, Math.max(...rows.map((r) => r.length)), { x: b.x, y: b.y, w: b.w, h: Math.min(b.h, rows.length * 48) });
      rows.forEach((r, ri) => r.forEach((v, ci) => { const cell = t.rows[ri]?.cells[ci]; if (cell) setPlainText(cell.text, v); }));
      s.objects = s.objects.filter((o) => o !== b).concat(t);
    } else if (b) { if (body.length) setLines(b.text, body); else if (!plainText(b.text).trim()) s.objects = s.objects.filter((o) => o !== b); }
    if (notes.length) s.notes = notes.join('\n');
    pres.slides.push(s);
  };
  for (const raw of String(md).replace(/\r/g, '').split('\n')) {
    const line = raw.replace(/\s+$/, '');
    const h = /^(#{1,3})\s+(.*)$/.exec(line);
    if (h) {
      flush();
      const layout = h[1] === '#' ? (pres.slides.length ? 'section' : 'title') : 'titleContent';
      s = newSlide(pres, layout); body = []; notes = []; table = [];
      const t = s.objects.find(TITLE);
      if (t) setPlainText(t.text, h[2]);
      continue;
    }
    if (!s) { if (!line.trim()) continue; s = newSlide(pres, 'titleContent'); body = []; notes = []; table = []; }
    if (/^\s*>/.test(line)) { notes.push(line.replace(/^\s*>\s?/, '')); continue; }
    if (/^\s*\|/.test(line)) { table.push(line); continue; }
    if (!line.trim()) continue;
    if (s.layout === 'title' || s.layout === 'section') {
      // 표지의 부제
      const sub = s.objects.find(BODY);
      if (sub) { setLines(sub.text, [...plainText(sub.text).split('\n').filter(Boolean), line.replace(/^\s*[-*]\s+/, '')]); continue; }
    }
    const m = /^(\s*)[-*+]\s+(.*)$/.exec(line) ?? /^(\s*)\d+[.)]\s+(.*)$/.exec(line);
    if (m) body.push(`${'\t'.repeat(Math.floor(m[1].replace(/\t/g, '  ').length / 2))}${m[2]}`);
    else body.push(line.trim());
  }
  flush();
  if (!pres.slides.length) pres.slides.push(newSlide(pres, 'title'));
  return pres;
}

const S1 = { slide: { type: 'number', description: '슬라이드 번호 (1부터)' } };
const TOOLS = [
  { name: 'open_presentation', description: '디스크의 발표 자료를 엽니다 (.pptx .ppsx .potx .pptm .ppt .pdf .wpt).', input: { path: { type: 'string' } }, required: ['path'],
    run: async (a) => { const path = resolve(a.path); const { pres, warnings } = await loadDeck(path); const name = basename(path); decks.set(name, { pres, path, dirty: false }); current = name; return { ...summary(decks.get(name)), warnings }; } },
  { name: 'new_presentation', description: `새 발표 자료. theme: ${THEMES.map((t) => t.name).join(' · ')}. size: wide(16:9) · standard(4:3). outline 에 마크다운 개요를 주면 바로 슬라이드로 ('# 표지 제목', '## 슬라이드 제목', '- 글머리'(들여쓰기 = 수준), '> 발표자 메모', '| 표 |').`, input: { name: { type: 'string' }, theme: { type: 'string' }, size: { type: 'string', enum: ['wide', 'standard'] }, outline: { type: 'string' } },
    run: async (a) => {
      const pres = a.outline ? deckFromOutline(a.outline, { theme: a.theme ?? 'office' }) : newPresentation({ theme: a.theme ?? 'office', size: a.size ?? 'wide' });
      const name = a.name || `프레젠테이션${decks.size + 1}`; decks.set(name, { pres, path: null, dirty: true }); current = name; return summary(decks.get(name));
    } },
  { name: 'list_presentations', description: '열린 발표 자료 목록.', input: {}, run: async () => ({ current, decks: [...decks.entries()].map(([n, d]) => ({ name: n, path: d.path, slides: d.pres.slides.length, dirty: d.dirty })) }) },
  { name: 'switch_presentation', description: '다룰 발표 자료를 바꿉니다.', input: { name: { type: 'string' } }, required: ['name'], run: async (a) => { if (!decks.has(a.name)) throw new Error(`'${a.name}' 은(는) 열려 있지 않습니다`); current = a.name; return summary(decks.get(a.name)); } },
  { name: 'presentation_info', description: '슬라이드 목록(번호 · 레이아웃 · 제목), 크기, 테마.', input: {}, run: async () => summary(cur()) },
  { name: 'read_slides', description: '슬라이드의 개체 · 글 · 표 · 메모를 읽습니다 (from~to, 생략하면 전부, 최대 30장씩).', input: { from: { type: 'number' }, to: { type: 'number' } },
    run: async (a) => { const p = cur().pres; const f = Math.max(1, a.from ?? 1); const t = Math.min(p.slides.length, a.to ?? f + 29); return p.slides.slice(f - 1, t).map((s, i) => slideInfo(s, f - 1 + i, true)); } },
  { name: 'set_text', description: '개체의 글을 바꿉니다. object: title · body · body2 · 개체 번호 · id. text 는 줄마다 단락 (앞의 탭 · 공백 2칸 = 수준, "- " 는 지움). 첫 단락의 서식을 이어받음.', input: { ...S1, object: { type: 'string' }, text: { type: 'string' } }, required: ['slide', 'text'],
    run: async (a) => {
      const d = cur(); const s = slideAt(d.pres, a.slide); const o = objectAt(s, a.object ?? 'body');
      if (!o?.text) throw new Error(`글을 넣을 개체를 찾지 못했습니다 (object: ${a.object ?? 'body'}). read_slides 로 개체 번호를 확인하세요`);
      setLines(o.text, String(a.text).split('\n')); d.dirty = true;
      return slideInfo(s, d.pres.slides.indexOf(s), true);
    } },
  { name: 'add_slide', description: `슬라이드를 더합니다. layout: ${LAYOUTS.map(([k, l]) => `${k}(${l})`).join(' · ')}. title · bullets(줄 배열, 탭 = 수준) · notes. at = 넣을 위치(번호, 생략하면 끝).`, input: { layout: { type: 'string' }, title: { type: 'string' }, bullets: { type: 'array', items: { type: 'string' } }, notes: { type: 'string' }, at: { type: 'number' } },
    run: async (a) => {
      const d = cur(); const p = d.pres;
      const s = newSlide(p, a.layout ?? 'titleContent');
      const t = s.objects.find(TITLE); if (t && a.title != null) setPlainText(t.text, a.title);
      const b = s.objects.find(BODY); if (b && a.bullets) setLines(b.text, a.bullets);
      if (a.notes) s.notes = a.notes;
      const at = a.at ? Math.max(0, Math.min(p.slides.length, a.at - 1)) : p.slides.length;
      p.slides.splice(at, 0, s); d.dirty = true;
      return slideInfo(s, at, true);
    } },
  { name: 'add_table', description: '슬라이드에 표를 넣습니다 (rows = 2차원 배열, 첫 행은 머리글). 본문 자리가 비어 있으면 그 자리에.', input: { ...S1, rows: { type: 'array', items: { type: 'array' } } }, required: ['slide', 'rows'],
    run: async (a) => {
      const d = cur(); const p = d.pres; const s = slideAt(p, a.slide);
      const b = s.objects.find((o) => BODY(o) && !plainText(o.text).trim());
      const W = p.size.w; const H = p.size.h;
      const rect = b ? { x: b.x, y: b.y, w: b.w, h: Math.min(b.h, a.rows.length * 44) } : { x: W * 0.08, y: H * 0.25, w: W * 0.84, h: Math.min(H * 0.65, a.rows.length * 44) };
      const cols = Math.max(...a.rows.map((r) => r.length));
      const t = newTable(a.rows.length, cols, rect);
      a.rows.forEach((r, ri) => r.forEach((v, ci) => { const c = t.rows[ri]?.cells[ci]; if (c) setPlainText(c.text, v == null ? '' : String(v)); }));
      if (b) s.objects = s.objects.filter((o) => o !== b);
      s.objects.push(t); d.dirty = true;
      return slideInfo(s, p.slides.indexOf(s));
    } },
  { name: 'delete_slide', description: '슬라이드를 지웁니다.', input: { ...S1 }, required: ['slide'], run: async (a) => { const d = cur(); const s = slideAt(d.pres, a.slide); if (d.pres.slides.length === 1) throw new Error('마지막 슬라이드는 지울 수 없습니다'); d.pres.slides.splice(d.pres.slides.indexOf(s), 1); d.dirty = true; return summary(d); } },
  { name: 'move_slide', description: '슬라이드를 to 위치로 옮깁니다.', input: { ...S1, to: { type: 'number' } }, required: ['slide', 'to'], run: async (a) => { const d = cur(); const s = slideAt(d.pres, a.slide); d.pres.slides.splice(d.pres.slides.indexOf(s), 1); d.pres.slides.splice(Math.max(0, Math.min(d.pres.slides.length, a.to - 1)), 0, s); d.dirty = true; return summary(d); } },
  { name: 'duplicate_slide', description: '슬라이드를 복제해 바로 뒤에 넣습니다.', input: { ...S1 }, required: ['slide'], run: async (a) => { const d = cur(); const s = slideAt(d.pres, a.slide); d.pres.slides.splice(d.pres.slides.indexOf(s) + 1, 0, duplicateSlide(s)); d.dirty = true; return summary(d); } },
  { name: 'set_notes', description: '발표자 메모를 바꿉니다.', input: { ...S1, notes: { type: 'string' } }, required: ['slide', 'notes'], run: async (a) => { const d = cur(); slideAt(d.pres, a.slide).notes = a.notes; d.dirty = true; return '메모를 바꿨습니다'; } },
  { name: 'set_hidden', description: '슬라이드 숨기기 / 보이기 (슬라이드 쇼에서 건너뜀).', input: { ...S1, hidden: { type: 'boolean' } }, required: ['slide', 'hidden'], run: async (a) => { const d = cur(); slideAt(d.pres, a.slide).hidden = !!a.hidden; d.dirty = true; return a.hidden ? '숨겼습니다' : '보이게 했습니다'; } },
  { name: 'find', description: '모든 슬라이드 · 표 · 메모에서 글을 찾습니다.', input: { text: { type: 'string' }, match_case: { type: 'boolean' } }, required: ['text'],
    run: async (a) => { const p = cur().pres; const hits = findAll(p, a.text, { matchCase: !!a.match_case }); return { count: hits.length, hits: hits.slice(0, 200).map((h) => ({ slide: h.slide + 1, ...(h.notes ? { notes: true } : { object: h.obj }), at: h.start })) }; } },
  { name: 'replace', description: '모든 슬라이드 · 표 · 메모의 글을 바꿉니다 (서식 유지).', input: { find: { type: 'string' }, replace: { type: 'string' }, match_case: { type: 'boolean' } }, required: ['find', 'replace'],
    run: async (a) => { const d = cur(); const n = replaceAll(d.pres, a.find, a.replace, { matchCase: !!a.match_case }); if (n) d.dirty = true; return `${n}곳을 바꿨습니다`; } },
  { name: 'get_outline', description: '발표 자료 전체를 마크다운 개요로 ("## 제목" + 글머리 + "> 메모").', input: {},
    run: async () => cur().pres.slides.map((s, i) => {
      const lines = [`## ${slideTitle(s) || `(제목 없음 ${i + 1})`}`];
      for (const o of s.objects.filter(BODY)) for (const p of o.text.paras) { const t = p.runs.map((r) => r.t).join(''); if (t) lines.push(`${'  '.repeat(p.lvl ?? 0)}- ${t}`); }
      for (const o of s.objects.filter((x) => x.type === 'table')) { const rows = o.rows.map((r) => `| ${r.cells.map((c) => plainText(c.text).replace(/\|/g, '\\|')).join(' | ')} |`); lines.push(rows[0], `|${' --- |'.repeat(o.rows[0].cells.length)}`, ...rows.slice(1)); }
      if (s.notes) lines.push(...s.notes.split('\n').map((l) => `> ${l}`));
      return lines.join('\n');
    }).join('\n\n') },
  { name: 'save_presentation', description: '저장합니다 (.pptx · .wpt). path 가 없으면 연 파일에 덮어씀 (.pptx 일 때).', input: { path: { type: 'string' } },
    run: async (a) => {
      const d = cur();
      const path = a.path ? resolve(a.path) : d.path;
      if (!path) throw new Error('새 발표 자료는 path 를 주세요 (예: 발표.pptx)');
      const k = kindOf(path);
      if (!a.path && k !== 'pptx') throw new Error(`.${k} 에는 덮어쓸 수 없습니다. path 로 .pptx 이름을 주세요`);
      if (k === 'pptx' || k === 'potx' || k === 'ppsx') await writeFile(path, Buffer.from(writePptx(d.pres)));
      else if (k === 'wpt' || k === 'json') await writeFile(path, JSON.stringify(d.pres));
      else throw new Error(`${k} 로는 저장할 수 없습니다 (.pptx .wpt)`);
      d.dirty = false; if (a.path) d.path = path;
      return `저장했습니다: ${path} (슬라이드 ${d.pres.slides.length}장)`;
    } },
];
const BY_NAME = Object.fromEntries(TOOLS.map((t) => [t.name, t]));
const toolSchema = (t) => ({ name: t.name, description: t.description, inputSchema: { type: 'object', properties: t.input ?? {}, ...(t.required ? { required: t.required } : {}) } });

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
        serverInfo: { name: 'wipoint', version: VERSION },
        instructions: 'WIPOINT 발표 자료 (PowerPoint 호환). open_presentation → read_slides · get_outline 으로 읽기 → set_text · add_slide · add_table · replace · set_notes → save_presentation(.pptx). 새로 만들 때는 new_presentation 에 마크다운 개요(outline)를 주면 한 번에 슬라이드가 됩니다.',
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
