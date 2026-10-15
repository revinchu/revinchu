// MCP 서버 (mcp.js): JSON-RPC 로 열기 · 쓰기 · 계산 · 저장 · 다시 열기
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join } from 'node:path';

function server(path) {
  const p = spawn(process.execPath, [path], { stdio: ['pipe', 'pipe', 'ignore'] });
  let buf = ''; const waits = new Map(); let id = 0;
  p.stdout.on('data', (d) => { buf += d; let i; while ((i = buf.indexOf('\n')) >= 0) { const m = JSON.parse(buf.slice(0, i)); buf = buf.slice(i + 1); waits.get(m.id)?.(m); } });
  const rpc = (method, params) => new Promise((res) => { const k = ++id; waits.set(k, res); p.stdin.write(`${JSON.stringify({ jsonrpc: '2.0', id: k, method, params })}\n`); });
  const call = async (name, args) => { const r = await rpc('tools/call', { name, arguments: args }); const t = r.result.content[0].text; if (r.result.isError) throw new Error(t); try { return JSON.parse(t); } catch { return t; } };
  return { rpc, call, close: () => p.stdin.end() };
}

test('WIPOINT MCP: 개요로 만들기 → 고치기 → pptx 저장 → 다시 열기', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'wipoint-mcp-'));
  const s = server(new URL('../mcp.js', import.meta.url).pathname);
  try {
    const init = await s.rpc('initialize', {});
    assert.equal(init.result.serverInfo.name, 'wipoint');
    const d = await s.call('new_presentation', { theme: 'navy', outline: '# 분기 보고\n2026년 3분기\n\n## 요약\n- 매출 12% 증가\n  - 온라인 30%\n> 매출 강조\n\n## 지역별\n| 지역 | 매출 |\n|---|---|\n| 서울 | 100 |' });
    assert.deepEqual(d.slides.map((x) => x.title), ['분기 보고', '요약', '지역별']);
    const [cover] = await s.call('read_slides', { from: 1, to: 1 });
    assert.equal(cover.objects.find((o) => o.placeholder === 'subTitle').text, '2026년 3분기');
    const [sum] = await s.call('read_slides', { from: 2, to: 2 });
    assert.equal(sum.objects.find((o) => o.placeholder === 'body').text, '매출 12% 증가\n  온라인 30%');
    assert.equal(sum.notes, '매출 강조');
    await s.call('add_slide', { title: '다음 단계', bullets: ['채용', '\t신입 3명'], at: 3 });
    assert.match(await s.call('replace', { find: '매출', replace: '수익' }), /3곳/);
    await s.call('set_text', { slide: 2, object: 'title', text: '핵심 요약' });
    await s.call('save_presentation', { path: join(dir, 'out.pptx') });
    const back = await s.call('open_presentation', { path: join(dir, 'out.pptx') });
    assert.deepEqual(back.slides.map((x) => x.title), ['분기 보고', '핵심 요약', '다음 단계', '지역별']);
    const [tbl] = await s.call('read_slides', { from: 4, to: 4 });
    assert.deepEqual(tbl.objects.find((o) => o.type === 'table').table, [['지역', '수익'], ['서울', '100']]);
    assert.match(await s.call('get_outline', {}), /## 핵심 요약\n- 수익 12% 증가\n  - 온라인 30%\n> 수익 강조/);
    await assert.rejects(s.call('set_text', { slide: 9, text: 'x' }), /없습니다/);
  } finally { s.close(); rmSync(dir, { recursive: true, force: true }); }
});

