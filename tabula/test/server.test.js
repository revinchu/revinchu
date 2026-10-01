import { test } from 'node:test';
import assert from 'node:assert/strict';
import { request } from 'node:http';
import { mkdtemp, mkdir, readFile, writeFile, readdir, rm, symlink } from 'node:fs/promises';
import { join, resolve, sep } from 'node:path';
import { tmpdir } from 'node:os';
import { fileURLToPath } from 'node:url';
import { execFile, spawn } from 'node:child_process';
import { promisify } from 'node:util';
import { createWixelServer, serverConfig } from '../server/app.js';
import { fetchPublicResource } from '../server/network.js';

const exec = promisify(execFile);
const serverFile = fileURLToPath(new URL('../server.js', import.meta.url));
const tempBase = process.platform === 'win32' ? resolve('D:/Codex/Temp/wixel-quality') : join(tmpdir(), 'wixel-quality');

async function fixture(t, options = {}) {
  await mkdir(tempBase, { recursive: true });
  const dir = await mkdtemp(join(tempBase, 'server-test-'));
  const root = join(dir, 'public'), data = join(dir, 'private');
  await mkdir(join(root, 'assets'), { recursive: true });
  await mkdir(join(root, 'src'), { recursive: true });
  await writeFile(join(root, 'index.html'), '<h1>위셀</h1>');
  await writeFile(join(root, 'src', 'app.js'), 'export const ready = true;');
  await writeFile(join(root, '.env'), 'TEST_ONLY=private');
  const server = createWixelServer({ root, data, ...options });
  await new Promise((resolve, reject) => { server.once('error', reject); server.listen(0, '127.0.0.1', resolve); });
  t.after(async () => {
    await new Promise((resolve) => { server.close(resolve); server.closeAllConnections?.(); });
    assert.ok(dir.startsWith(tempBase + sep));
    await rm(dir, { recursive: true, force: true });
  });
  return { server, root, data };
}

function call(server, path, { method = 'GET', headers = {}, body } = {}) {
  return new Promise((resolve, reject) => {
    const req = request({ hostname: '127.0.0.1', port: server.address().port, path, method, headers, agent: false }, (res) => {
      const chunks = [];
      res.on('data', (chunk) => chunks.push(chunk));
      res.on('error', reject);
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, text: Buffer.concat(chunks).toString() }));
    });
    req.setTimeout(3000, () => req.destroy(new Error('테스트 요청 시간 초과')));
    req.once('error', reject); req.end(body);
  });
}

test('서버: 기본은 루프백이며 외부 수신 주소의 무인증 시작을 거부한다', async () => {
  assert.equal(serverConfig({}).host, '127.0.0.1');
  assert.equal(serverConfig({ HOST: 'localhost' }).host, '127.0.0.1');
  assert.equal(serverConfig({ HOST: '::1' }).host, '::1');
  for (const host of ['0.0.0.0', '::', '192.168.0.2', 'public.example']) {
    assert.throws(() => serverConfig({ HOST: host }), /TABULA_TOKEN/);
    assert.throws(() => serverConfig({ HOST: host, TABULA_TOKEN: '  ' }), /TABULA_TOKEN/);
  }
  assert.equal(serverConfig({ HOST: '0.0.0.0', TABULA_TOKEN: 'test-only' }).token, 'test-only');
  assert.throws(() => serverConfig({ PORT: 'wrong' }), /PORT/);
  assert.throws(() => createWixelServer({ host: '0.0.0.0' }), /TABULA_TOKEN/);
  await assert.rejects(exec(process.execPath, [serverFile], {
    env: { ...process.env, HOST: '0.0.0.0', TABULA_TOKEN: '', PORT: '0' }, timeout: 3000,
  }), (error) => error.code === 1 && /TABULA_TOKEN/.test(error.stderr));
  const imported = await exec(process.execPath, ['--input-type=module', '-e', `await import(${JSON.stringify(new URL('../server.js', import.meta.url).href)})`], { timeout: 3000 });
  assert.equal(imported.stdout, '');
});

test('서버: junction/symlink 경로로 실행해도 실제 서버가 시작된다', async (t) => {
  await mkdir(tempBase, { recursive: true });
  const dir = await mkdtemp(join(tempBase, 'server-entry-'));
  const alias = join(dir, 'linked');
  await symlink(fileURLToPath(new URL('../', import.meta.url)), alias, process.platform === 'win32' ? 'junction' : 'dir');
  const child = spawn(process.execPath, [join(alias, 'server.js')], {
    env: { ...process.env, HOST: '127.0.0.1', TABULA_TOKEN: '', PORT: '0', TABULA_DATA: join(dir, 'data') },
    windowsHide: true, stdio: ['ignore', 'pipe', 'pipe'],
  });
  t.after(async () => {
    if (child.exitCode === null && child.signalCode === null) await new Promise((resolve) => { child.once('exit', resolve); child.kill(); });
    assert.ok(dir.startsWith(tempBase + sep)); await rm(dir, { recursive: true, force: true });
  });
  const port = await new Promise((resolve, reject) => {
    let output = '', errors = '';
    const timer = setTimeout(() => reject(new Error('서버 시작 시간 초과: ' + errors)), 3000);
    child.stdout.on('data', (chunk) => {
      output += chunk; const match = /WIXEL: http:\/\/localhost:(\d+)/.exec(output);
      if (match) { clearTimeout(timer); resolve(Number(match[1])); }
    });
    child.stderr.on('data', (chunk) => { errors += chunk; });
    child.once('error', (error) => { clearTimeout(timer); reject(error); });
    child.once('exit', (code) => { clearTimeout(timer); reject(new Error(`서버가 먼저 종료됨: ${code} ${errors}`)); });
  });
  assert.equal((await call({ address: () => ({ port }) }, '/api/health')).status, 200);
});

test('서버: 로컬 무암호 문서 저장·읽기·목록·삭제 흐름과 기본 보안 헤더', async (t) => {
  const { server } = await fixture(t);
  const name = encodeURIComponent('업무 문서');
  const health = await call(server, '/api/health');
  assert.deepEqual(JSON.parse(health.text), { ok: true, auth: false, publish: true });
  assert.equal(health.headers['x-content-type-options'], 'nosniff');
  assert.match(health.headers['content-security-policy'], /frame-ancestors 'self'/);
  assert.equal((await call(server, '/api/files/' + name, { method: 'PUT', body: '{"value":123}' })).status, 200);
  assert.equal((await call(server, '/api/files/' + name)).text, '{"value":123}');
  assert.equal(JSON.parse((await call(server, '/api/files')).text)[0].name, '업무 문서');
  assert.equal((await call(server, '/api/files/' + name, { method: 'DELETE' })).status, 200);
  assert.equal((await call(server, '/api/files/' + name)).status, 404);
  assert.equal((await call(server, '/api/files/%00', { method: 'PUT', body: '{}' })).status, 400);
  assert.equal((await call(server, '/api/files/%broken')).status, 400);
  assert.equal((await call(server, '/api/files/CON', { method: 'PUT', body: '{}' })).status, 400);
});

test('서버: 로컬 무암호 모드도 외부 Origin과 DNS rebinding Host로부터 보호한다', async (t) => {
  const { server } = await fixture(t);
  assert.equal((await call(server, '/api/files', { headers: { Host: 'evil.example' } })).status, 401);
  assert.equal((await call(server, '/api/files/doc', { method: 'PUT', headers: { Origin: 'https://evil.example' }, body: '{}' })).status, 403);
  assert.equal((await call(server, '/api/files/doc', { method: 'PUT', headers: { Origin: 'null' }, body: '{}' })).status, 403);
  assert.equal((await call(server, '/api/files')).text, '[]');
});

test('서버: 모든 보호 API와 네이버 중계가 인증 후 실행되며 키는 재전송하지 않는다', async (t) => {
  let proxyCalls = 0;
  const { server } = await fixture(t, {
    token: 'test-token', fetchResource: async (url, options) => {
      proxyCalls++;
      assert.equal(new URL(url).hostname, 'api.searchad.naver.com');
      assert.equal(options.maxRedirects, 0);
      assert.equal(options.headers['X-API-KEY'], 'test-key');
      assert.ok(options.headers['X-Signature']); assert.equal(options.headers['X-Secret'], undefined);
      return { status: 200, body: Buffer.from('{"keywordList":[]}'), headers: {} };
    },
  });
  for (const [method, path] of [
    ['GET', '/api/files'], ['GET', '/api/files/doc'], ['PUT', '/api/files/doc'], ['DELETE', '/api/files/doc'],
    ['POST', '/api/publish'], ['PUT', '/api/published/12345678'], ['DELETE', '/api/published/12345678'],
    ['GET', '/api/fetch?url=https%3A%2F%2Fpublic.example'], ['GET', '/api/naver/keywordstool?hintKeywords=test'],
  ]) {
    assert.equal((await call(server, path, { method, body: '{}' })).status, 401, path);
    assert.equal((await call(server, path, { method, headers: { 'X-Tabula-Token': 'wrong' }, body: '{}' })).status, 401, path);
  }
  assert.equal(proxyCalls, 0);
  const headers = { 'X-Tabula-Token': 'test-token', 'X-Customer': 'test-customer', 'X-API-KEY': 'test-key', 'X-Secret': 'test-secret' };
  assert.equal((await call(server, '/api/naver/keywordstool?hintKeywords=test', { headers })).status, 200);
  assert.equal(proxyCalls, 1);
  assert.equal(JSON.parse((await call(server, '/api/health')).text).auth, true);
  assert.equal((await call(server, '/api/files', { headers })).status, 200);
});

test('서버: 게시 읽기는 공개지만 게시·갱신·삭제에는 토큰이 필요하다', async (t) => {
  const { server } = await fixture(t, { token: 'test-token' });
  const headers = { 'X-Tabula-Token': 'test-token' };
  const created = await call(server, '/api/publish', { method: 'POST', headers, body: '{"shared":1}' });
  assert.equal(created.status, 200);
  const path = '/api/published/' + JSON.parse(created.text).id;
  assert.equal((await call(server, path)).text, '{"shared":1}');
  assert.equal((await call(server, path, { method: 'PUT', headers, body: '{"shared":2}' })).status, 200);
  assert.equal((await call(server, path)).text, '{"shared":2}');
  assert.equal((await call(server, path, { method: 'DELETE', headers })).status, 200);
  assert.equal((await call(server, path)).status, 404);
});

test('서버: 실제 API 요청에서 내부 DNS 주소를 차단하고 중계 속도 제한을 적용한다', async (t) => {
  let outbound = 0;
  const { server } = await fixture(t, {
    proxyRateLimit: 2,
    fetchResource: (url) => fetchPublicResource(url, {
      lookup: async () => [{ address: '93.184.216.34', family: 4 }, { address: '127.0.0.1', family: 4 }],
      request: async () => { outbound++; throw new Error('내부 주소로 요청하면 안 됨'); },
    }),
  });
  const path = '/api/fetch?url=' + encodeURIComponent('http://public.example/');
  assert.equal((await call(server, path)).status, 400);
  assert.equal((await call(server, path)).status, 400);
  const limited = await call(server, path);
  assert.equal(limited.status, 429); assert.ok(Number(limited.headers['retry-after']) > 0);
  assert.equal(outbound, 0);
});

test('서버: 정적 프런트엔드만 제공하고 숨김 파일·서버 소스·데이터·심볼릭 링크 탈출을 차단한다', async (t) => {
  const { server, root, data } = await fixture(t);
  await mkdir(data, { recursive: true }); await writeFile(join(data, 'secret.json'), '{"private":true}');
  await symlink(data, join(root, 'assets', 'private'), process.platform === 'win32' ? 'junction' : 'dir');
  await mkdir(join(root, 'server')); await writeFile(join(root, 'server', 'config.json'), '{"private":true}');
  await symlink(join(root, 'server'), join(root, 'assets', 'server-alias'), process.platform === 'win32' ? 'junction' : 'dir');
  assert.equal((await call(server, '/')).status, 200);
  assert.equal((await call(server, '/src/app.js')).status, 200);
  assert.equal((await call(server, '/src/app.js', { method: 'HEAD' })).text, '');
  for (const path of ['/.env', '/server.js', '/server/app.js', '/test/server.test.js', '/data/secret.json',
    '/assets/private/secret.json', '/assets/server-alias/config.json', '/assets/%2e%2e%5c.env', '/%2e%2e%5cprivate%5csecret.json']) {
    assert.equal((await call(server, path)).status, 403, path);
  }
});

test('서버: 같은 문서의 동시 저장에서도 임시 파일 충돌이나 JSON 손실이 없다', async (t) => {
  const { server, data } = await fixture(t);
  const results = await Promise.all(Array.from({ length: 12 }, (_, n) => call(server, '/api/files/shared', {
    method: 'PUT', body: JSON.stringify({ n, text: '가'.repeat(1000) }),
  })));
  assert.ok(results.every((result) => result.status === 200), results.map((result) => result.status).join(','));
  const value = JSON.parse(await readFile(join(data, 'shared.tabula.json'), 'utf8'));
  assert.ok(value.n >= 0 && value.n < 12); assert.equal(value.text.length, 1000);
  assert.deepEqual(await readdir(data), ['shared.tabula.json']);
});
