import { test } from 'node:test';
import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { readFileSync, readdirSync, mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');

test('빌드: 한 파일 · Cloudflare 용 모두 성공 (import 형식 · 순환 검사 포함)', () => {
  const dir = mkdtempSync(join(tmpdir(), 'wipoint-build-'));
  try {
    execFileSync(process.execPath, [join(root, 'build.mjs'), join(dir, 'a')]);
    const html = readFileSync(join(dir, 'a', 'index.html'), 'utf8');
    assert.match(html, /<style>/);
    assert.doesNotMatch(html, /src="src\/app\.js"/);
    execFileSync(process.execPath, [join(root, 'build.mjs'), join(dir, 'b'), '--cloud']);
    const files = readdirSync(join(dir, 'b'));
    assert.ok(files.some((f) => /^wipoint-[a-f0-9]{16}\.js$/.test(f)));
    assert.ok(files.includes('_headers'));
    assert.match(readFileSync(join(dir, 'b', '_headers'), 'utf8'), /script-src 'self'/);
  } finally { rmSync(dir, { recursive: true, force: true }); }
});

test('소스 규칙: 동적 import · export { } · 다른 앱 폴더 import 금지', () => {
  for (const f of readdirSync(join(root, 'src'))) {
    const s = readFileSync(join(root, 'src', f), 'utf8');
    assert.doesNotMatch(s, /\bimport\(/, f);
    assert.doesNotMatch(s, /^export\s*\{/m, f);
    assert.doesNotMatch(s, /from '\.\.\//, f);
  }
});
