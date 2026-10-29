// kv_reference 의 페이지와 데이터를 public/ 으로 복사한다 (ingest/, worker/ 는 제외).
import { cpSync, mkdirSync, rmSync, existsSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const site = join(here, '..');
const out = join(here, 'public');
rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });
for (const f of ['index.html', 'data.js', 'app.js', 'hosted.js', 'live.js', 'bulk.js', 'ai.js']) cpSync(join(site, f), join(out, f));
if (existsSync(join(site, 'bulk'))) cpSync(join(site, 'bulk'), join(out, 'bulk'), { recursive: true });
console.log('public/ 준비 완료');
