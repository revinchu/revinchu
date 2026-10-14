// AI 연결 (MCP) 도구 · MCP 서버 (node --test)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtempSync, existsSync, readFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { runTool, TOOLS, toolSchema } from '../src/ops.js';
import { newDoc, para, docText, paraPlain } from '../src/model.js';
import { readDocx } from '../src/docx.js';

const base = () => { const d = newDoc(); d.body = [para('첫 문단 사과'), para('둘째 문단 사과 배')]; return d; };

test('도구 스키마', () => {
  for (const t of TOOLS) { const s = toolSchema(t); assert.equal(s.inputSchema.type, 'object'); assert.ok(s.description.length > 10, t.name); }
});

test('replace · 변경 내용 추적 · 적용', () => {
  let d = base();
  let r = runTool(d, 'replace', { find: '사과', replace: '포도' });
  assert.equal(docText(r.doc), '첫 문단 포도\n둘째 문단 포도 배');
  r = runTool(d, 'replace', { find: '사과', replace: '귤', track: true, author: 'AI' });
  d = r.doc;
  const runs = d.body[0].runs;
  assert.ok(runs.some((x) => x.del && x.text === '사과'));
  assert.ok(runs.some((x) => x.ins && x.text === '귤'));
  assert.equal(runTool(d, 'track_changes', { action: 'list' }).result.length, 4);
  const acc = runTool(d, 'track_changes', { action: 'accept_all' }).doc;
  assert.equal(docText(acc), '첫 문단 귤\n둘째 문단 귤 배');
  const rej = runTool(d, 'track_changes', { action: 'reject_all' }).doc;
  assert.equal(docText(rej), '첫 문단 사과\n둘째 문단 사과 배');
});

test('정규식 바꾸기 · 찾기', () => {
  const r = runTool(base(), 'replace', { find: '(\\S+) 문단', replace: '[$1]', regex: true });
  assert.equal(docText(r.doc), '[첫] 사과\n[둘째] 사과 배');
  const f = runTool(base(), 'find', { query: '사과' }).result;
  assert.equal(f.count, 2);
  assert.ok(f.hits[0].context.includes('【사과】'));
});

test('마크다운 넣기 · 문단 고치기 · 스타일 · 서식 · 표 · 메모 · 목차', () => {
  let d = base();
  d = runTool(d, 'insert_markdown', { markdown: '# 새 제목\n\n본문', position: 'start' }).doc;
  assert.equal(d.body[0].style, 'Heading1');
  d = runTool(d, 'edit_paragraph', { index: 1, text: '바뀐 본문', style: '제목 2' }).doc;
  assert.equal(paraPlain(d.body[1]), '바뀐 본문');
  assert.equal(d.body[1].style, 'Heading2');
  d = runTool(d, 'apply_style', { match: '둘째', style: 'Quote' }).doc;
  assert.equal(d.body.find((b) => paraPlain(b).startsWith('둘째')).style, 'Quote');
  d = runTool(d, 'format_text', { match: '사과', bold: true, color: '#FF0000' }).doc;
  assert.ok(d.body.find((b) => paraPlain(b).startsWith('첫')).runs.some((r) => r.b && r.color === '#FF0000' && r.text === '사과'));
  d = runTool(d, 'insert_table', { rows: [['이름', '값'], ['a', '1']] }).doc;
  assert.ok(d.body.some((b) => b.type === 'tbl' && b.rows[0].header));
  d = runTool(d, 'add_comment', { match: '바뀐', text: '확인' }).doc;
  assert.equal(d.comments.length, 1);
  assert.equal(runTool(d, 'comments', { action: 'reply', id: d.comments[0].id, text: '네' }).doc.comments[0].replies.length, 1);
  d = runTool(d, 'insert_toc', {}).doc;
  assert.equal(d.body[0].style, 'TOCHeading');
  const info = runTool(d, 'inspect', {}).result;
  assert.ok(info.headings.length >= 2);
  assert.equal(info.tables, 1);
});

test('쪽 설정 · 머리글 · 속성', () => {
  let d = base();
  d = runTool(d, 'page_setup', { paper: 'A4', orientation: 'landscape', margins: 'narrow', columns: 2 }).doc;
  assert.equal(d.section.page.orient, 'landscape');
  assert.ok(d.section.page.w > d.section.page.h);
  assert.equal(d.section.cols.n, 2);
  d = runTool(d, 'header_footer', { kind: 'footer', text: '- {PAGE} -' }).doc;
  assert.ok(d.section.footers.default[0].runs.some((r) => r.obj === 'field' && r.code === 'PAGE'));
  d = runTool(d, 'set_properties', { title: '보고서' }).doc;
  assert.equal(d.props.title, '보고서');
});

// MCP 서버 (stdio)
const here = dirname(fileURLToPath(import.meta.url));
function mcp() {
  const p = spawn(process.execPath, [join(here, '..', 'mcp.js')], { stdio: ['pipe', 'pipe', 'ignore'], env: { ...process.env, WIWORD_URL: 'http://127.0.0.1:9' } });
  let buf = ''; const waits = new Map(); let id = 0;
  p.stdout.on('data', (d) => { buf += d; let i; while ((i = buf.indexOf('\n')) >= 0) { const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1); waits.get(m.id)?.(m); } });
  const call = (method, params) => new Promise((r) => { const k = ++id; waits.set(k, r); p.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: k, method, params })}\n`); });
  const tool = async (name, args) => { const r = await call('tools/call', { name, arguments: args }); return { text: r.result.content[0].text, err: !!r.result.isError }; };
  return { call, tool, close: () => p.stdin.end() };
}
test('MCP 서버: 새 문서 → 고치기 → docx 저장', async () => {
  const s = mcp();
  const init = await s.call('initialize', { protocolVersion: '2025-06-18' });
  assert.equal(init.result.serverInfo.name, 'wiword');
  const tools = (await s.call('tools/list', {})).result.tools.map((t) => t.name);
  for (const n of ['open_document', 'save_document', 'convert', 'get_markdown', 'replace', 'execute']) assert.ok(tools.includes(n), n);
  assert.equal((await s.tool('new_document', { markdown: '# 제목\n\n안녕하세요' })).err, false);
  await s.tool('replace', { find: '안녕하세요', replace: '반갑습니다' });
  const md = await s.tool('get_markdown', {});
  assert.ok(md.text.includes('반갑습니다'));
  const dir = mkdtempSync(join(tmpdir(), 'wiword-'));
  const out = join(dir, 'a.docx');
  const sv = await s.tool('save_document', { path: out });
  assert.equal(sv.err, false, sv.text);
  assert.ok(existsSync(out));
  assert.ok(docText(readDocx(new Uint8Array(readFileSync(out))).doc).includes('반갑습니다'));
  const conv = await s.tool('convert', { input: out, output: join(dir, 'a.hwpx') });
  assert.equal(conv.err, false, conv.text);
  const app = await s.tool('execute', { command: 'bold' });
  assert.equal(app.err, true); // 앱 없음 → 친절한 오류
  s.close();
});
