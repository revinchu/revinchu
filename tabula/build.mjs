// 한 파일짜리 배포본 만들기 (의존성 없음)
//   node build.mjs            → dist/index.html (CSS·JS 를 모두 넣은 단일 HTML) + dist/.nojekyll
// 정적 호스팅(GitHub Pages 등)에 올리거나 파일을 바로 열어도 동작합니다.
// 서버 저장소(/api/files)가 없으면 앱이 자동으로 브라우저 저장(localStorage)을 사용합니다.
import { readFileSync, writeFileSync, mkdirSync, copyFileSync, existsSync, readdirSync, unlinkSync } from 'node:fs';
import { dirname, join, resolve, posix } from 'node:path';
import { fileURLToPath } from 'node:url';
import { gunzipSync } from 'node:zlib';
import { createHash } from 'node:crypto';

const root = dirname(fileURLToPath(import.meta.url));
const args = process.argv.slice(2);
const cloud = args.includes('--cloud');
const positional = args.filter((arg) => arg !== '--cloud');
if (positional.length > 1 || positional.some((arg) => arg.startsWith('--'))) throw new Error('사용법: node build.mjs [출력 폴더] [--cloud]');
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
  if (/^\s*import\s/m.test(code)) throw new Error(`${file}: 지원하지 않는 import 형식`);
  code = code.replace(/^export\s+(async\s+function|function\*?|class|const|let|var)\s+([A-Za-z_$][\w$]*)/gm, (_, kw, name) => {
    exportsList.push(name);
    return `${kw} ${name}`;
  });
  if (/^export\s/m.test(code)) throw new Error(`${file}: 지원하지 않는 export 형식`);
  // 내보낸 이름은 getter 로 연결해 재할당(let)에도 최신 값을 보게 함
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
const cloudJs = `globalThis.TABULA_STATIC = false;\n${js}`;
const cloudName = `wixel-${createHash('sha256').update(cloudJs).digest('hex').slice(0, 16)}.js`;
html = html
  .replace('<link rel="stylesheet" href="styles.css">', () => `<style>\n${css}\n</style>`)
  .replace('<script type="module" src="src/app.js"></script>', () => cloud ? `<script type="module" src="${cloudName}"></script>` : `<script>globalThis.TABULA_STATIC = true;</script>\n<script type="module">\n${js}\n</script>`);
mkdirSync(outDir, { recursive: true });
writeFileSync(join(outDir, 'index.html'), html);
writeFileSync(join(outDir, '.nojekyll'), '');
if (cloud) {
  for (const old of readdirSync(outDir)) {
    if (/^wixel-[a-f0-9]{16}\.js$/.test(old) && old !== cloudName) unlinkSync(join(outDir, old));
  }
  writeFileSync(join(outDir, cloudName), cloudJs);
  writeFileSync(join(outDir, 'version.json'), `${JSON.stringify({ asset: cloudName })}\n`);
  writeFileSync(join(outDir, '_headers'), `/*
  Content-Security-Policy: script-src 'self'; script-src-attr 'none'; object-src 'none'; base-uri 'self'; frame-ancestors 'self'; form-action 'self'
  X-Content-Type-Options: nosniff
  Referrer-Policy: same-origin
  X-Frame-Options: SAMEORIGIN
/${cloudName}
  Cache-Control: public, max-age=31536000, immutable
/index.html
  Cache-Control: no-cache
/version.json
  Cache-Control: no-cache
/
  Cache-Control: no-cache
`);
}

// 따로 불러오는 큰 자료 (아이콘 모음 등): dist/assets 로 복사
mkdirSync(join(outDir, 'assets'), { recursive: true });
for (const f of ['iconlib.json.gz', '네이버 연관검색어 키워드 검색.xlsm']) {
  if (!existsSync(join(root, 'assets', f))) continue;
  copyFileSync(join(root, 'assets', f), join(outDir, 'assets', f));
  // 압축 · 엑셀 파일을 못 올리는 곳(일부 정적 호스팅)을 위한 대체본: 풀어 둔 JSON · base64 텍스트
  const bytes = readFileSync(join(root, 'assets', f));
  if (f.endsWith('.gz')) writeFileSync(join(outDir, 'assets', f.slice(0, -3)), gunzipSync(bytes));
  else writeFileSync(join(outDir, 'assets', `${f}.b64.txt`), bytes.toString('base64'));
}
console.log(`${join(outDir, 'index.html')} (${(html.length / 1024).toFixed(0)} KB)`);
