// 한 파일짜리 배포본 만들기 (의존성 없음) — WIXEL(tabula/build.mjs)과 같은 방식
//   node build.mjs            → dist/index.html (CSS·JS 를 모두 넣은 단일 HTML) + dist/assets/
//   node build.mjs --cloud    → dist-cloudflare/ (index.html + 해시 이름 JS + _headers: Cloudflare Workers 정적 자산용)
import { readFileSync, writeFileSync, mkdirSync, copyFileSync, existsSync, readdirSync, unlinkSync } from 'node:fs';
import { dirname, join, resolve, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { gunzipSync } from 'node:zlib';

const root = dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
const cloud = args.includes('--cloud');
const positional = args.filter((a) => a !== '--cloud');
const outDir = resolve(root, positional[0] ?? (cloud ? 'dist-cloudflare' : 'dist'));

/** ES 모듈 하나를 함수 범위로 감싸고 import/export 를 모듈 표 참조로 바꿈 */
function transform(file, source) {
  const deps = [];
  const exportsList = [];
  let code = source.replace(/^import\s*\{([^}]*)\}\s*from\s*'([^']+)';?/gm, (_, names, spec) => {
    const dep = posix.join(posix.dirname(file), spec);
    deps.push(dep);
    const binds = names.split(',').map((n) => n.trim()).filter(Boolean).map((n) => n.replace(/\s+as\s+/, ': '));
    return `const { ${binds.join(', ')} } = __mods[${JSON.stringify(dep)}];`;
  });
  if (/^\s*import\s/m.test(code) || /\bimport\(/.test(code)) throw new Error(`${file}: 지원하지 않는 import 형식 (정적 import { … } from './x.js' 만 가능)`);
  code = code.replace(/^export\s+(async\s+function|function\*?|class|const|let|var)\s+([A-Za-z_$][\w$]*)/gm, (_, kw, name) => {
    exportsList.push(name);
    return `${kw} ${name}`;
  });
  if (/^export\s/m.test(code)) throw new Error(`${file}: 지원하지 않는 export 형식`);
  const exp = exportsList.map((n) => `get ${n}() { return ${n}; }`).join(', ');
  return { deps, code: `__mods[${JSON.stringify(file)}] = await (async () => {\n${code}\nreturn { ${exp} };\n})();` };
}

function bundle(entry) {
  const order = [];
  const state = new Map();
  const visit = (file) => {
    if (state.get(file) === 'done') return;
    if (state.get(file) === 'visiting') throw new Error(`순환 import: ${file}`);
    state.set(file, 'visiting');
    const t = transform(file, readFileSync(join(root, file), 'utf8'));
    for (const d of t.deps) visit(d);
    state.set(file, 'done');
    order.push(t.code);
  };
  visit(entry);
  return `const __mods = {};\n${order.join('\n')}`;
}

let html = readFileSync(join(root, 'index.html'), 'utf8');
if (!html.includes('<script type="module" src="src/app.js"></script>') || !html.includes('<link rel="stylesheet" href="styles.css">')) {
  throw new Error('index.html 구조가 바뀌었습니다 (styles.css / src/app.js 참조를 찾지 못함)');
}
const css = readFileSync(join(root, 'styles.css'), 'utf8');
const js = bundle('src/app.js').replace(/<\/script/gi, '<\\/script');
const jsName = `wipoint-${createHash('sha256').update(js).digest('hex').slice(0, 16)}.js`;
html = html
  .replace('<link rel="stylesheet" href="styles.css">', () => `<style>\n${css}\n</style>`)
  .replace('<script type="module" src="src/app.js"></script>', () => (cloud ? `<script type="module" src="${jsName}"></script>` : `<script type="module">\n${js}\n</script>`));
mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, 'index.html'), html);
writeFileSync(join(outDir, '.nojekyll'), '');
if (cloud) {
  for (const old of readdirSync(outDir)) if (/^wipoint-[a-f0-9]{16}\.js$/.test(old) && old !== jsName) unlinkSync(join(outDir, old));
  writeFileSync(join(outDir, jsName), js);
  writeFileSync(join(outDir, 'version.json'), `${JSON.stringify({ asset: jsName })}\n`);
  // 인라인 스크립트가 없으므로 script-src 'self' 로 잠금 (스타일은 인라인 허용: 슬라이드가 style 속성으로 그려짐)
  writeFileSync(join(outDir, '_headers'), `/*
  Content-Security-Policy: script-src 'self'; script-src-attr 'none'; object-src 'none'; base-uri 'self'; frame-ancestors 'self'; form-action 'self'
  X-Content-Type-Options: nosniff
  Referrer-Policy: same-origin
  X-Frame-Options: SAMEORIGIN
/${jsName}
  Cache-Control: public, max-age=31536000, immutable
/index.html
  Cache-Control: no-cache
/
  Cache-Control: no-cache
`);
}
// 아이콘 모음: WIXEL 과 같은 파일을 씀 (저장소에 두 벌 두지 않도록 ../tabula/assets 에서 복사)
mkdirSync(join(outDir, 'assets'), { recursive: true });
const iconSrc = [join(root, 'assets', 'iconlib.json.gz'), join(root, '..', 'tabula', 'assets', 'iconlib.json.gz')].find((p) => existsSync(p));
if (iconSrc) {
  copyFileSync(iconSrc, join(outDir, 'assets', 'iconlib.json.gz'));
  writeFileSync(join(outDir, 'assets', 'iconlib.json'), gunzipSync(readFileSync(iconSrc)));
} else console.warn('경고: 아이콘 모음(iconlib.json.gz)을 찾지 못했습니다 — 삽입 › 아이콘이 동작하지 않습니다');
console.log(`${join(outDir, 'index.html')} (${(html.length / 1024).toFixed(0)} KB)${cloud ? ` + ${jsName} (${(js.length / 1024).toFixed(0)} KB)` : ''}`);
