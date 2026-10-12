import assert from 'node:assert/strict';
import { randomBytes, createHash } from 'node:crypto';
import { mkdir } from 'node:fs/promises';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_URL || 'http://127.0.0.1:8792/';
const root = new URL(base).origin, key = randomBytes(32).toString('base64url'), name = 'AI-화면검증-' + Date.now();
const dir = process.env.WIXEL_TEST_OUTPUT || 'D:/Codex/Temp/wixel-ai-mcp'; await mkdir(dir, { recursive: true });
const browser = await chromium.launch(); let context, page; const errors = [], requests = [];
try {
  context = await browser.newContext({ viewport: { width: 1440, height: 1000 } }); page = await context.newPage();
  page.on('pageerror', e => errors.push(e.message));
  await page.addInitScript(key => { localStorage.setItem('wixel.connection.v3', JSON.stringify({ kind: 'vault', key })); window.WIXEL_SKIP_START = true; }, key);
  await page.route('**/api/ai/chat', async route => {
    const data = route.request().postDataJSON(); requests.push(data);
    assert.equal(route.request().headers()['x-wixel-ai-key'], 'fixture-api-key');
    await route.fulfill({ json: { text: JSON.stringify({ answer: '<script>alert(1)</script> 분석 결과입니다.', changes: [{ cell: 'C1', input: '=SUM(A1:B1)' }] }), citations: [], model: data.model } });
  });
  const healthReady = page.waitForResponse(r => r.url().endsWith('/api/health'));
  await page.goto(base); await healthReady; await page.waitForFunction(() => window.tabula?.wb());
  await page.waitForTimeout(500);
  await page.evaluate(() => { const t = window.tabula, w = t.wb(); w.transact(() => { w.setInput(0, 0, 0, '100'); w.setInput(0, 0, 1, '300'); }); t.selectRange({ r1: 0, c1: 0, r2: 0, c2: 1 }, 'cells', { r: 0, c: 0 }); });
  await page.evaluate(() => window.tabula.run('aiAssistant'));
  await page.getByRole('dialog', { name: 'AI 도우미', exact: true }).waitFor();
  await page.getByLabel('AI API 키', { exact: true }).fill('fixture-api-key');
  await page.getByLabel('AI 요청', { exact: true }).fill('합계 수식을 만들어줘');
  await page.getByRole('button', { name: 'AI에게 요청', exact: true }).click();
  await page.getByRole('button', { name: '변경안 적용', exact: true }).waitFor();
  await page.waitForFunction(() => ![...document.querySelectorAll('button')].find(x => x.textContent === '변경안 적용')?.disabled);
  assert.equal(requests.at(-1).context, null);
  assert.equal(await page.locator('.ai-answer script').count(), 0);
  await page.screenshot({ path: dir + '/ai-assistant.png' });
  await page.getByRole('button', { name: '변경안 적용', exact: true }).click();
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 0, 2).raw), '=SUM(A1:B1)');
  await page.evaluate(() => window.tabula.wb().undo());
  assert.equal(await page.evaluate(() => window.tabula.wb().getCell(0, 0, 2)?.raw ?? ''), '');
  await page.getByText('선택한 셀을 함께 전송', { exact: false }).check();
  for (const provider of ['anthropic', 'perplexity', 'xai', 'deepseek']) {
    await page.getByLabel('AI 서비스', { exact: true }).selectOption(provider);
    await page.getByLabel('AI API 키', { exact: true }).fill('fixture-api-key');
    await page.getByRole('button', { name: 'AI에게 요청', exact: true }).click();
    await page.waitForFunction(() => ![...document.querySelectorAll('button')].find(x => x.textContent === 'AI에게 요청')?.disabled);
  }
  assert.equal(requests.length, 5); assert.equal(requests.at(-1).context.cells.length, 2);
  await page.evaluate(() => window.tabula.wb().transact(() => window.tabula.wb().setInput(0, 3, 0, 'changed')));
  await page.getByRole('button', { name: '변경안 적용', exact: true }).click();
  assert.match(await page.locator('.ai-panel [role=status]').textContent(), /문서가 바뀌/);
  assert.equal(await page.evaluate(() => JSON.stringify(localStorage).includes('fixture-api-key')), false);
  await page.getByRole('button', { name: '닫기', exact: true }).last().click();
  const workbook = await page.evaluate(() => window.tabula.wb().serialize());
  const saved = await fetch(root + '/api/files/' + encodeURIComponent(name), { method: 'PUT', headers: { 'X-Wixel-Vault': key, 'If-Match': '0', 'Content-Type': 'application/json' }, body: JSON.stringify({ workbook }) }); assert.equal(saved.status, 200);
  await page.evaluate(() => window.tabula.run('aiConnections'));
  await page.getByRole('dialog', { name: '외부 AI · MCP 연결', exact: true }).waitFor();
  await page.getByText('토큰을 직접 입력하는 MCP 클라이언트', { exact: true }).click();
  await page.waitForFunction(() => document.querySelector('[aria-label="연결할 온라인 문서"]')?.options.length > 0);
  await page.getByRole('button', { name: '수동 연결 토큰 발급', exact: true }).click();
  await page.getByLabel('MCP 연결 토큰', { exact: true }).waitFor();
  assert.match(await page.getByLabel('MCP 연결 토큰').inputValue(), /^[\w-]{43}\.[\w-]{43}$/);
  await page.screenshot({ path: dir + '/mcp-connections.png' });
  await page.getByRole('button', { name: '연결 해제', exact: true }).click();
  await page.waitForFunction(() => !document.querySelector('.ai-connection'));
  await page.getByRole('button', { name: '닫기', exact: true }).last().click();
  // Browser-based OAuth consent checks its real cookie binding and displays the chosen document.
  const redirect = 'https://client.example/callback', verifier = randomBytes(32).toString('base64url');
  const reg = await (await fetch(root + '/api/oauth/register', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ client_name: '브라우저 검증 앱', redirect_uris: [redirect] }) })).json();
  const query = new URLSearchParams({ client_id: reg.client_id, redirect_uri: redirect, resource: root + '/api/mcp', response_type: 'code', code_challenge_method: 'S256', code_challenge: createHash('sha256').update(verifier).digest('base64url'), scope: 'document:read document:write', state: 'test' });
  await page.goto(root + '/api/oauth/authorize?' + query);
  await page.waitForFunction(() => document.getElementById('documents')?.options[0]?.value);
  assert.match(await page.locator('#client').textContent(), /브라우저 검증 앱/);
  await page.locator('#write').check(); await page.screenshot({ path: dir + '/mcp-consent.png' });
  await page.route('https://client.example/callback**', route => route.fulfill({ contentType: 'text/plain', body: 'OAuth callback received' }));
  await page.locator('#allow').click(); await page.waitForURL('https://client.example/callback**');
  assert.ok(new URL(page.url()).searchParams.get('code'));
  assert.deepEqual(errors, []);
  console.log(JSON.stringify({ ok: true, providers: requests.map(x => x.provider), previewApplyUndo: true, staleBlocked: true, keysNotPersisted: true, oauthBrowserConsent: true, mockedPaidAI: true }));
} catch (error) {
  if (page) { await page.screenshot({ path: dir + '/ai-ui-failure.png' }); console.log((await page.locator('body').innerText()).slice(-3500)); }
  throw error;
} finally {
  const headers = { 'X-Wixel-Vault': key };
  const list = await fetch(root + '/api/ai/connections', { headers }); if (list.ok) for (const grant of await list.json()) await fetch(root + '/api/ai/connections/' + grant.id, { method: 'DELETE', headers });
  const path = root + '/api/files/' + encodeURIComponent(name), response = await fetch(path, { headers });
  if (response.ok) { const revision = response.headers.get('X-Wixel-Revision'); await response.body.cancel(); await fetch(path, { method: 'DELETE', headers: { ...headers, 'If-Match': revision } }); }
  await browser.close();
}
