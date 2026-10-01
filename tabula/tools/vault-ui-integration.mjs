// 로컬 Cloudflare Worker(8787)에서 새 임시 보관함으로 실제 UI/API 왕복 검증. 종료 전에 만든 자료 삭제.
import assert from 'node:assert/strict';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const base = process.env.WIXEL_CLOUD_URL || 'http://localhost:8787/';
if (!['localhost', '127.0.0.1'].includes(new URL(base).hostname)) throw new Error('이 검사는 로컬 Worker에서만 실행합니다.');
const browser = await chromium.launch();
const first = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const second = await browser.newPage({ viewport: { width: 1440, height: 1000 } });
const errors = []; let publishedUrl = null, key = null, name = null;
for (const p of [first, second]) { p.setDefaultTimeout(15000); p.on('pageerror', (e) => errors.push(e.message)); }
try {
  let writes = 0;
  first.on('request', (req) => { if (/\/api\/files\//.test(req.url()) && req.method() === 'PUT') writes++; });
  await first.goto(base); await first.waitForFunction(() => !!window.tabula);
  await first.locator('[data-page=storage]').click();
  await first.getByRole('button', { name: '다른 복구키로 연결', exact: true }).click();
  const recoveryDownload = first.waitForEvent('download');
  await first.getByRole('button', { name: '새 보관함 만들기', exact: true }).click();
  assert.match((await recoveryDownload).suggestedFilename(), /복구키\.json$/);
  assert.equal(writes, 0);
  key = await first.evaluate(() => JSON.parse(localStorage.getItem('wixel.connection.v3')).key);
  await first.keyboard.press('Escape');
  await first.evaluate(() => window.tabula.wb().transact(() => window.tabula.wb().setInput(0, 0, 0, '123')));
  await first.evaluate(() => window.tabula.run('saveLocations'));
  await first.getByRole('button', { name: /온라인 개인 보관함/ }).click();
  await first.waitForFunction(() => document.querySelector('#saveState').textContent.includes('저장됨'));
  assert.equal(writes, 1);
  const files = await first.evaluate(async (key) => (await fetch('api/files', { headers: { 'X-Wixel-Vault': key } })).json(), key);
  assert.equal(files.length, 1); name = files[0].name;
  assert.equal(await first.evaluate(async () => (await fetch('api/files')).status), 401);
  console.log('OK 새 보관함 복구키 다운로드 · 연결만으로 업로드 0 · 명시 저장 · 익명 접근 거부');
  await first.screenshot({ path: 'D:/Codex/Temp/wixel3-vault.png' });

  await second.goto(base); await second.waitForFunction(() => !!window.tabula);
  await second.locator('[data-page=storage]').click(); await second.getByRole('button', { name: '다른 복구키로 연결', exact: true }).click();
  await second.getByRole('textbox', { name: '개인 보관함 복구키' }).fill(key);
  await second.getByRole('button', { name: '기존 복구키로 연결', exact: true }).click();
  await second.locator('[data-page=open]').click(); await second.locator('.hub-file-name').filter({ hasText: name }).click();
  await second.waitForFunction(() => window.tabula.wb().getValue(0, 0, 0) === 123);
  console.log('OK 별도 기기 컨텍스트에서 복구키 연결 · 저장 문서 열기');

  await first.locator('[data-page=share]').click(); await first.getByRole('button', { name: /읽기 전용 링크 공유/ }).click();
  await first.getByRole('button', { name: '온라인에 게시 (짧은 링크)', exact: true }).click();
  await first.locator('.pub-link input').waitFor(); publishedUrl = await first.locator('.pub-link input').inputValue();
  assert.equal(publishedUrl.includes(key), false);
  const id = new URL(publishedUrl).searchParams.get('view');
  const publicStatus = await first.evaluate(async (id) => (await fetch(`api/published/${id}`)).status, id); assert.equal(publicStatus, 200);
  await second.evaluate(() => window.tabula.run('backstage')); await second.locator('[data-page=share]').click();
  await second.locator('.hub-publications input').waitFor(); assert.equal(await second.locator('.hub-publications input').inputValue(), publishedUrl);
  await second.locator('.hub-publications').getByRole('button', { name: '게시 중지', exact: true }).click();
  await second.getByRole('dialog', { name: '게시 중지', exact: true }).getByRole('button', { name: '게시 중지', exact: true }).click();
  await second.getByText('게시한 링크가 없습니다', { exact: true }).waitFor();
  assert.equal(await first.evaluate(async (id) => (await fetch(`api/published/${id}`)).status, id), 404);
  publishedUrl = null;
  console.log('OK 공유 링크에 복구키 미포함 · 익명 읽기 · 다른 기기 게시 목록/중지');
  assert.deepEqual(errors, []);
  console.log(JSON.stringify({ groups: 3, bad: 0, pageErrors: 0 }));
} finally {
  if (key && name) await first.evaluate(async ({ key, name, publishedUrl }) => {
    if (publishedUrl) { const id = new URL(publishedUrl).searchParams.get('view'); const p = await fetch(`api/published/${id}`); if (p.ok) await fetch(`api/published/${id}`, { method: 'DELETE', headers: { 'X-Wixel-Vault': key, 'If-Match': p.headers.get('ETag') } }); }
    const path = `api/files/${encodeURIComponent(name)}`, file = await fetch(path, { headers: { 'X-Wixel-Vault': key } });
    if (file.ok) await fetch(path, { method: 'DELETE', headers: { 'X-Wixel-Vault': key, 'If-Match': file.headers.get('ETag') } });
  }, { key, name, publishedUrl }).catch(() => {});
  await browser.close();
}
