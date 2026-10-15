#!/usr/bin/env node
// WIWORD MCP 서버 (stdio, JSON-RPC 2.0, 의존성 없음)
//   Claude Code:     claude mcp add wiword -- node /경로/wiword/mcp.js
//   Claude Desktop:  { "mcpServers": { "wiword": { "command": "node", "args": ["/경로/wiword/mcp.js"] } } }
// 두 가지 대상:
//   file — 디스크의 .docx · .hwp · .hwpx · .doc · .pdf · .md · .html · .rtf · .txt 를 열어 고치고 .docx · .hwpx · .md · .html · .txt 로 저장
//   app  — 브라우저에 열린 WIWORD 창 (npm start 로 띄운 server.js 를 거쳐 실시간으로 읽고 고침, 화면에 바로 보임)
import { readFile, writeFile } from 'node:fs/promises';
import { resolve, extname, basename } from 'node:path';
import { createInterface } from 'node:readline';
import { readBytes, docBytes } from './src/formats.js';
import { TOOLS, runTool, toolSchema } from './src/ops.js';
import { newDoc, docText, serialize } from './src/model.js';
import { sampleDoc, TEMPLATES } from './src/sample.js';

const VERSION = '1.0.0';
const APP_URL = (process.env.WIWORD_URL ?? 'http://127.0.0.1:5180').replace(/\/$/, '');
const log = (...a) => process.stderr.write(`[wiword-mcp] ${a.join(' ')}\n`);

// ───────────── 열린 문서 ─────────────
const docs = new Map(); // 이름 → { doc, path, kind, dirty }
let current = null;
const cur = () => (current ? docs.get(current) : null);
const SAVE_KINDS = ['docx', 'hwpx', 'md', 'html', 'txt', 'wiword', 'dotx'];
const SAVE_FORMATS = [...SAVE_KINDS, 'pdf'];
function kindOfPath(p) { const e = extname(p).slice(1).toLowerCase(); return e === 'markdown' ? 'md' : e === 'htm' ? 'html' : e; }

async function appCall(tool, args = {}) {
  let res;
  try { res = await fetch(`${APP_URL}/api/bridge/call`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ tool, args }) }); } catch { throw new Error(`WIWORD 앱(${APP_URL})에 연결할 수 없습니다. wiword 폴더에서 npm start 를 실행하고 브라우저에서 연 뒤 다시 시도하세요. (파일만 다루려면 open_document 를 쓰세요)`); }
  const j = await res.json().catch(() => ({}));
  if (!res.ok || j.error) throw new Error(j.error ?? `앱 오류 (${res.status})`);
  return j.result;
}
async function appAvailable() {
  try { const r = await fetch(`${APP_URL}/api/bridge/status`, { signal: AbortSignal.timeout(800) }); if (!r.ok) return false; const j = await r.json(); return j.apps > 0; } catch { return false; }
}
async function pickTarget(t) {
  if (t === 'file' || t === 'app') return t;
  if (cur()) return 'file';
  if (await appAvailable()) return 'app';
  // 아무것도 없으면 빈 새 문서
  docs.set('새 문서', { doc: newDoc(), path: null, kind: null, dirty: false });
  current = '새 문서';
  return 'file';
}

// ───────────── 도구 ─────────────
const TARGET = { target: { type: 'string', enum: ['file', 'app'], description: 'file = open_document 로 연 문서, app = 브라우저의 WIWORD 창 (생략하면 연 문서가 있으면 file, 없고 앱이 켜져 있으면 app)' } };
const META_TOOLS = [
  {
    name: 'open_document', description: '디스크의 문서를 엽니다 (.docx .doc .hwp .hwpx .pdf .md .html .rtf .odt .txt .csv). to_app=true 면 브라우저의 WIWORD 창에서 엽니다.',
    input: { path: { type: 'string', description: '파일 경로' }, password: { type: 'string', description: '암호로 보호된 한글(.hwp · .hwpx) 문서의 열기 암호' }, to_app: { type: 'boolean' }, pdf_mode: { type: 'string', enum: ['flow', 'layout'], description: 'PDF: flow(문단으로 다시 잇기, 기본) | layout(원래 모양)' } }, required: ['path'],
    run: async (a) => {
      const p = resolve(a.path);
      const bytes = new Uint8Array(await readFile(p));
      if (a.to_app) return appCall('open_document', { name: basename(p), data: Buffer.from(bytes).toString('base64'), password: a.password });
      const r = readBytes(bytes, p, { mode: a.pdf_mode, password: a.password });
      const name = basename(p);
      docs.set(name, { doc: r.doc, path: p, kind: r.kind, dirty: false });
      current = name;
      return { opened: name, kind: r.kind, blocks: r.doc.body.length, chars: docText(r.doc).length, warnings: r.warnings ?? [] };
    },
  },
  {
    name: 'new_document', description: '새 문서를 만듭니다 (메모리). template: blank(기본) · report · gongmun(공문서) · letter · resume · minutes · press · essay · hwp(한글 기본) · tour. markdown 을 주면 그 내용으로.',
    input: { name: { type: 'string' }, template: { type: 'string' }, markdown: { type: 'string' } },
    run: async (a) => {
      const t = TEMPLATES.find((x) => x.id === (a.template ?? 'blank'));
      let doc = t ? t.build() : newDoc();
      if (a.markdown) doc = runTool(doc, 'set_markdown', { markdown: a.markdown }).doc;
      const name = a.name ?? `문서${docs.size + 1}`;
      docs.set(name, { doc, path: null, kind: null, dirty: true });
      current = name;
      return { created: name, template: t?.id ?? 'blank' };
    },
  },
  {
    name: 'list_documents', description: '열려 있는 문서 목록 (file) 과 WIWORD 앱 연결 상태.',
    input: {},
    run: async () => ({ current, documents: [...docs].map(([n, d]) => ({ name: n, path: d.path, kind: d.kind, dirty: d.dirty, blocks: d.doc.body.length })), app: (await appAvailable()) ? await appCall('app_status').catch((e) => e.message) : `연결 안 됨 (${APP_URL})` }),
  },
  {
    name: 'switch_document', description: '현재 문서를 바꿉니다 (list_documents 의 이름).',
    input: { name: { type: 'string' } }, required: ['name'],
    run: async (a) => { if (!docs.has(a.name)) throw new Error(`'${a.name}' 이(가) 열려 있지 않습니다`); current = a.name; return `현재 문서: ${a.name}`; },
  },
  {
    name: 'save_document', description: '현재 문서를 저장합니다. path 를 주면 그 이름으로 (확장자로 형식 결정: .docx .hwpx .md .html .txt .wiword .dotx .pdf — PDF 는 WIWORD 창이 열려 있어야 함). 없으면 연 파일에 덮어씀 (docx · hwpx · md · html · txt 일 때). target=app 이면 앱의 문서를 path 로 저장.',
    input: { path: { type: 'string' }, format: { type: 'string', enum: SAVE_FORMATS }, ...TARGET },
    run: async (a) => {
      const target = await pickTarget(a.target);
      if (target === 'app') {
        if (!a.path) throw new Error('앱의 문서를 저장하려면 path 가 필요합니다');
        const kind = a.format ?? kindOfPath(a.path);
        const r = await appCall('export_document', { format: kind });
        await writeFile(resolve(a.path), Buffer.from(r.data, 'base64'));
        return `저장했습니다: ${resolve(a.path)} (${kind})`;
      }
      const d = cur();
      if (!d) throw new Error('열린 문서가 없습니다');
      const path = a.path ? resolve(a.path) : d.path;
      if (!path) throw new Error('새 문서는 path 가 필요합니다 (예: 보고서.docx)');
      const kind = a.format ?? kindOfPath(path);
      if (kind === 'pdf') {
        // PDF 는 글자 배치(쪽 나누기)가 필요해 열려 있는 WIWORD 창이 그림
        const r = await appCall('render_pdf', { doc: serialize(d.doc), name: basename(path) }).catch((e) => { throw new Error(`PDF 저장은 WIWORD 창이 열려 있어야 합니다 (wiword 폴더에서 npm start 후 브라우저로 열기). ${e.message}`); });
        await writeFile(path, Buffer.from(r.data, 'base64'));
        return `PDF 로 저장했습니다: ${path}`;
      }
      if (!SAVE_KINDS.includes(kind)) throw new Error(`${kind} 형식으로는 저장할 수 없습니다. .docx · .hwpx · .md · .html · .txt · .pdf 중에서 고르세요.`);
      const out = docBytes(d.doc, kind, { name: basename(path) });
      await writeFile(path, typeof out === 'string' ? (kind === 'txt' ? `\uFEFF${out}` : out) : Buffer.from(out));
      d.dirty = false;
      return `저장했습니다: ${path} (${kind})`;
    },
  },
  {
    name: 'convert', description: '파일 형식 바꾸기: input(.docx .doc .hwp .hwpx .pdf .md .html .rtf .odt .txt) → output(.docx .hwpx .md .html .txt). 예: 한글(hwp) → Word(docx), PDF → 마크다운.',
    input: { input: { type: 'string' }, output: { type: 'string' }, password: { type: 'string', description: '암호 문서의 열기 암호' } }, required: ['input', 'output'],
    run: async (a) => {
      const r = readBytes(new Uint8Array(await readFile(resolve(a.input))), a.input, { password: a.password });
      const kind = kindOfPath(a.output);
      if (!SAVE_KINDS.includes(kind)) throw new Error(`${kind} 로는 바꿀 수 없습니다`);
      const out = docBytes(r.doc, kind, { name: basename(a.output) });
      await writeFile(resolve(a.output), typeof out === 'string' ? out : Buffer.from(out));
      return { from: r.kind, to: kind, output: resolve(a.output), warnings: r.warnings ?? [] };
    },
  },
  {
    name: 'execute', description: '[앱] WIWORD 명령을 실행합니다 (리본 · 단축키와 같은 명령: bold, applyStyle, insertTable, pageBreak, insertToc, zoomPageWidth …). list_commands 로 목록을 보세요. args 는 명령 인자 배열.',
    input: { command: { type: 'string' }, args: { type: 'array' } }, required: ['command'],
    run: async (a) => appCall('execute', { command: a.command, args: a.args ?? [] }),
  },
  { name: 'list_commands', description: '[앱] 실행할 수 있는 WIWORD 명령 목록 (이름 · 설명 · 인자).', input: {}, run: async () => appCall('list_commands') },
  { name: 'get_selection', description: '[앱] 사용자가 WIWORD 에서 선택한 글.', input: {}, run: async () => appCall('get_selection') },
];

const DOC_TOOLS = TOOLS.map((t) => ({
  name: t.name, description: t.description, input: { ...(t.input ?? {}), ...TARGET }, required: t.required,
  run: async (a) => {
    const { target: tg, ...args } = a;
    const target = await pickTarget(tg);
    if (target === 'app') return appCall(t.name, args);
    const d = cur();
    const r = runTool(d.doc, t.name, args);
    if (r.doc && r.doc !== d.doc) { d.doc = r.doc; d.dirty = true; }
    return r.result;
  },
}));
const ALL = [...META_TOOLS, ...DOC_TOOLS];
const BY_NAME = Object.fromEntries(ALL.map((t) => [t.name, t]));

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
        serverInfo: { name: 'wiword', version: VERSION },
        instructions: 'WIWORD 문서 편집기. 파일: open_document → (get_markdown · inspect · list_paragraphs 로 읽기) → replace · edit_paragraph · insert_markdown · apply_style … 로 고치기 → save_document. 열린 WIWORD 창을 다루려면 target="app" (또는 문서를 열지 않은 상태). 한글(.hwp/.hwpx) · Word(.docx/.doc) · PDF · 마크다운을 읽고 .docx · .hwpx · .md 로 저장합니다.',
      });
    case 'notifications/initialized': case 'notifications/cancelled': return null;
    case 'ping': return ok(id, {});
    case 'tools/list': return ok(id, { tools: ALL.map(toolSchema) });
    case 'tools/call': {
      const t = BY_NAME[params?.name];
      if (!t) return ok(id, { content: [{ type: 'text', text: `알 수 없는 도구: ${params?.name}` }], isError: true });
      try {
        const r = await t.run(params.arguments ?? {});
        return ok(id, { content: [{ type: 'text', text: text(r ?? 'OK') }] });
      } catch (e) {
        return ok(id, { content: [{ type: 'text', text: `오류: ${e.message ?? e}` }], isError: true });
      }
    }
    case 'resources/list': return ok(id, { resources: [] });
    case 'prompts/list': return ok(id, { prompts: [] });
    default:
      if (id !== undefined) return fail(id, -32601, `지원하지 않는 메서드: ${method}`);
      return null;
  }
}

const rl = createInterface({ input: process.stdin, crlfDelay: Infinity });
rl.on('line', (line) => {
  if (!line.trim()) return;
  let msg;
  try { msg = JSON.parse(line); } catch { send({ jsonrpc: '2.0', id: null, error: { code: -32700, message: 'JSON 파싱 오류' } }); return; }
  Promise.resolve(handle(msg)).catch((e) => { log('오류', e.stack ?? e); if (msg.id !== undefined) fail(msg.id, -32603, e.message); });
});
rl.on('close', () => process.exit(0));
log(`시작 (앱: ${APP_URL})`);
void sampleDoc;
