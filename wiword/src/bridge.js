// AI 연결 (MCP) — 열린 WIWORD 창과 MCP 서버(mcp.js) 를 server.js 중계로 잇기
// 창: GET /api/bridge/events (SSE) 로 요청 { id, tool, args } 를 받고 POST /api/bridge/reply 로 { id, result | error } 를 돌려줌
import { S, run, register, emit, change, setDoc, curSel, COMMAND_INFO } from './state.js';
import { runTool, TOOLS } from './ops.js';
import { docText, posEq, deserialize } from './model.js';
import { readBytes, docBytes } from './formats.js';
import { exportPdfBytes } from './pdfexport.js';
import { selectedText } from './commands.js';
import { toast } from './ui.js';

let es = null;
const log = (m) => { S.bridge = S.bridge ?? {}; (S.bridge.log ??= []).push(`${new Date().toLocaleTimeString('ko-KR')} ${m}`); if (S.bridge.log.length > 500) S.bridge.log.splice(0, 100); emit('pane'); };
const b64 = (bytes) => { let s = ''; for (let i = 0; i < bytes.length; i += 0x8000) s += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000)); return btoa(s); };
const unb64 = (s) => { const bin = atob(s); const out = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i); return out; };
const clientId = `w${Math.random().toString(36).slice(2, 10)}`;

/** 요청 처리 (AI 가 이 창의 문서를 읽고 고침) */
async function handle({ tool, args = {} }) {
  if (tool === 'list_commands') return Object.entries(COMMAND_INFO).map(([name, i]) => ({ name, label: i.label, args: i.args ?? [] }));
  if (tool === 'execute') {
    if (S.opts.bridgeAsk && !confirm(`AI 가 명령을 실행하려 합니다: ${args.command}`)) throw new Error('사용자가 거부했습니다');
    const r = await run(args.command, ...(args.args ?? []));
    return r === undefined ? 'OK' : r;
  }
  if (tool === 'get_selection') { const s = curSel(); return { text: posEq(s.a, s.f) ? '' : selectedText(), anchor: s.a, focus: s.f }; }
  if (tool === 'app_status') return { name: S.docName, dirty: S.dirty, pages: S.layout?.pages.length ?? 0, chars: docText(S.doc).length, view: S.view };
  if (tool === 'open_document') {
    const bytes = unb64(args.data);
    const r = readBytes(bytes, args.name ?? 'document.docx', { password: args.password });
    setDoc(r.doc, (args.name ?? '문서').replace(/\.[^.]+$/, ''));
    return { opened: args.name, warnings: r.warnings ?? [] };
  }
  if (tool === 'new_document') { run('newDocument'); return 'OK'; }
  if (tool === 'render_pdf') {
    const doc = deserialize(args.doc);
    const out = await exportPdfBytes({ doc, name: args.name });
    return { format: 'pdf', data: b64(out) };
  }
  if (tool === 'export_document' && args.format === 'pdf') {
    const out = await exportPdfBytes({});
    return { format: 'pdf', name: S.docName, data: b64(out) };
  }
  if (tool === 'export_document') {
    const out = docBytes(S.doc, args.format ?? 'docx', { name: S.docName });
    return { format: args.format ?? 'docx', name: S.docName, data: typeof out === 'string' ? b64(new TextEncoder().encode(out)) : b64(out) };
  }
  // 문서 도구 (ops.js)
  if (S.opts.bridgeAsk && !['get_text', 'get_markdown', 'inspect', 'list_paragraphs', 'find'].includes(tool) && !confirm(`AI 가 문서를 고치려 합니다: ${tool}`)) throw new Error('사용자가 거부했습니다');
  const r = runTool(S.doc, tool, args);
  if (r.doc && r.doc !== S.doc) {
    const nd = r.doc;
    change((d) => { for (const k of Object.keys(nd)) d.doc[k] = nd[k]; return S.sel && clampSel(nd) ? S.sel : { p: [0], o: 0 }; });
  }
  return r.result;
}
function clampSel(doc) { const s = S.sel; return s && s.f.p.length === 1 && s.f.p[0] < doc.body.length && s.a.p[0] < doc.body.length; }

export function bridgeConnect() {
  if (es) es.close();
  if (!location.protocol.startsWith('http')) { S.bridge = { ...(S.bridge ?? {}), connected: false, available: false }; emit('pane'); return; }
  S.bridge = { ...(S.bridge ?? {}), connected: false, available: true, calls: S.bridge?.calls ?? 0 };
  try { es = new EventSource(`/api/bridge/events?c=${clientId}`); } catch { S.bridge.available = false; emit('pane'); return; }
  es.addEventListener('open', () => { S.bridge.connected = true; log('AI 연결 대기 중 (MCP 서버가 이 창을 씁니다)'); emit('ui'); });
  es.addEventListener('error', () => { if (S.bridge.connected) log('연결이 끊겼습니다 — 다시 연결합니다'); S.bridge.connected = false; emit('ui'); if (es?.readyState === 2) { S.bridge.available = false; es = null; } });
  es.addEventListener('message', async (ev) => {
    let msg;
    try { msg = JSON.parse(ev.data); } catch { return; }
    if (!msg?.id || !msg.tool) return;
    S.bridge.calls = (S.bridge.calls ?? 0) + 1;
    log(`← ${msg.tool} ${JSON.stringify(msg.args ?? {}).slice(0, 120)}`);
    let body;
    try { body = { id: msg.id, result: await handle(msg) }; log(`→ ${msg.tool} 완료`); } catch (e) { body = { id: msg.id, error: e.message ?? String(e) }; log(`→ ${msg.tool} 오류: ${body.error}`); }
    fetch('/api/bridge/reply', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) }).catch(() => {});
    if (msg.tool !== 'get_text' && msg.tool !== 'inspect') document.getElementById('mcpDot')?.removeAttribute('hidden');
  });
}
register({
  bridgeConnect: () => { bridgeConnect(); toast('AI 연결 (MCP) 을 기다립니다'); },
  bridgeDisconnect: () => { es?.close(); es = null; S.bridge = { ...(S.bridge ?? {}), connected: false }; emit('ui'); emit('pane'); },
}, { bridgeConnect: { label: 'AI 연결 (MCP) 켜기' } });

export const BRIDGE_TOOLS = TOOLS.map((t) => t.name);
