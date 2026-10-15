// MCP 서버 (mcp.js): JSON-RPC 로 열기 · 쓰기 · 계산 · 저장 · 다시 열기
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { spawn } from 'node:child_process';
import { mkdtempSync, writeFileSync, existsSync, rmSync } from 'node:fs';
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

test('WIXEL MCP: CSV 열기 → 수식 · 서식 · 정렬 · 시트 → xlsx 저장 → 다시 열기', async () => {
  const dir = mkdtempSync(join(tmpdir(), 'wixel-mcp-'));
  const s = server(new URL('../mcp.js', import.meta.url).pathname);
  try {
    const init = await s.rpc('initialize', {});
    assert.equal(init.result.serverInfo.name, 'wixel');
    const tools = (await s.rpc('tools/list', {})).result.tools.map((t) => t.name);
    for (const n of ['open_workbook', 'read_range', 'write_range', 'evaluate', 'save_workbook', 'convert']) assert.ok(tools.includes(n), n);
    writeFileSync(join(dir, 'in.csv'), 'name,city,sales\nKim,서울,100\nLee,부산,250\nPark,서울,75\n');
    const info = await s.call('open_workbook', { path: join(dir, 'in.csv') });
    assert.equal(info.sheets[0].used, 'A1:C4');
    await s.call('write_range', { start: 'A6', values: [['합계', '', '=SUM(C2:C4)', '=SUMIFS(C2:C4,B2:B4,"서울")']] });
    await s.call('set_format', { range: 'C2:D6', number_format: '#,##0원' });
    const r = await s.call('read_range', { range: 'C6:D6' });
    assert.deepEqual(r.rows, [['425원', '175원']]);
    assert.deepEqual((await s.call('read_range', { range: 'C6:D6', values: true })).rows, [[425, 175]]);
    assert.equal((await s.call('evaluate', { formula: 'XLOOKUP("Lee",A2:A4,C2:C4)*2' })).value, 500);
    assert.equal((await s.call('find', { text: '서울' })).count, 2);
    await s.call('sort_range', { range: 'A1:C4', by: 'C', descending: true, header: true });
    assert.deepEqual((await s.call('read_range', { range: 'A2:A4' })).rows.flat(), ['Lee', 'Kim', 'Park']);
    await s.call('add_sheet', { name: '요약' });
    await s.call('write_range', { sheet: '요약', start: 'A1', values: [['=in!C6*2']] });
    await s.call('rename_sheet', { sheet: 'in', name: '원본' });
    assert.deepEqual((await s.call('read_range', { sheet: '요약', formulas: true })).rows, [['=원본!C6*2']]);
    await s.call('save_workbook', { path: join(dir, 'out.xlsx') });
    assert.ok(existsSync(join(dir, 'out.xlsx')));
    const back = await s.call('open_workbook', { path: join(dir, 'out.xlsx') });
    assert.deepEqual(back.sheets.map((x) => x.name), ['원본', '요약']);
    assert.deepEqual((await s.call('read_range', { sheet: '요약', values: true })).rows, [[850]]);
    await assert.rejects(s.call('read_range', { sheet: '없음' }), /없습니다/);
  } finally { s.close(); rmSync(dir, { recursive: true, force: true }); }
});
