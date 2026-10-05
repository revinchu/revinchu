import test from 'node:test';
import assert from 'node:assert/strict';
import { createFontUsage, addFontUsage, collectMarkupFontUsage, embedFontCss, embedSvgFonts, fontCssForSvg } from '../src/font-export.js';
import { createSheetHtmlBlob, writeSheetHtmlToSink } from '../src/html-export.js';
import { rangeImageSvg } from '../src/range-image-export.js';
import { getWebFont } from '../src/fonts.js';
import { Workbook } from '../src/workbook.js';

const family = 'Noto Sans KR';
const urls = { latin: 'https://fonts.gstatic.com/s/fixture/latin.woff2', korean: 'https://fonts.gstatic.com/s/fixture/korean.woff2', unused: 'https://fonts.gstatic.com/s/fixture/cyrillic.woff2', bold: 'https://fonts.gstatic.com/s/fixture/bold.woff2' };
const face = (url, unicodeRange, weight = 400) => `@font-face {font-family:'${family}';font-style:normal;font-weight:${weight};src:url('${url}') format('woff2');unicode-range:${unicodeRange};}`;
const fontCss = face(urls.latin, 'U+0000-00FF') + face(urls.korean, 'U+AC00-D7A3') + face(urls.unused, 'U+0400-04FF') + face(urls.bold, 'U+0000-00FF', 700);
const license = 'Copyright 2026 Font Fixture Authors\nSIL OPEN FONT LICENSE Version 1.1\nPERMISSION & CONDITIONS\nPermission is hereby granted, free of charge.';
const bytes = new Uint8Array([119, 79, 70, 50, 0, 1, 2, 3]);
function network(css = fontCss) {
  const requests = [];
  const fetchImpl = async (url, options) => {
    requests.push({ url, options });
    return new Response(String(url).startsWith('https://raw.githubusercontent.com/') ? license : String(url).startsWith('https://fonts.googleapis.com/') ? css : bytes, { status: 200 });
  };
  return { requests, fetchImpl };
}

test('SVG 상속·부분 서식의 실제 문자에 필요한 폰트 조각만 파일에 포함한다', async () => {
  const source = `<svg xmlns="http://www.w3.org/2000/svg"><g font-family="${family}"><text>가A</text><text font-family="Arial">Ж</text></g></svg>`;
  const usage = collectMarkupFontUsage(source), mock = network(), result = await embedFontCss(usage, mock);
  assert.equal(usage.size, 1); assert.deepEqual([...usage.get(family).points].sort((a,b) => a-b), [65, 0xac00]);
  assert.deepEqual(mock.requests.filter(r => r.url.startsWith('https://fonts.gstatic.com/')).map(r => r.url).sort(), [urls.korean, urls.latin].sort());
  assert.match(result.css, /data:font\/woff2;base64,d09GMgABAgM=/); assert.equal(result.warnings.length, 0);
  assert.doesNotMatch(result.css, /src:url\(['"]https?:/); assert.match(result.css, /SIL OPEN FONT LICENSE/); assert.match(result.css, /Copyright 2026 Font Fixture Authors/);
  for (const request of mock.requests) { assert.equal(request.options.credentials, 'omit'); assert.equal(request.options.referrerPolicy, 'no-referrer'); assert.ok(!request.url.includes('text=')); assert.ok(!request.url.includes('가A')); }
  const embedded = await embedSvgFonts(source, mock);
  assert.match(embedded, /data-wixel-embedded-fonts="true"/); assert.match(embedded, /<text>가A<\/text>/);
  assert.equal(await embedSvgFonts(embedded, { fetchImpl: () => { throw Error('중복 요청'); } }), embedded);
});

test('글꼴 목록·스크립트·사용하지 않은 style 선언은 다운로드 대상이 아니다', async () => {
  const source = `<svg><style>text{font-family:'${family}'}</style><script>secret</script><text font-family="Arial">로컬 글꼴</text></svg>`;
  const mock = network(); assert.equal(await embedSvgFonts(source, mock), source); assert.equal(mock.requests.length, 0);
});

test('굵기 선택·CSS entity·unicode wildcard를 적용하며 외부 URL은 포함하지 않는다', async () => {
  const css = face(urls.korean, 'U+AC??', 700) + face(urls.latin, 'U+AC??', 400);
  const mock = network(css), source = `<svg><text style="font-family:&quot;${family}&quot;;font-weight:700">&#xAC00;</text></svg>`;
  const result = await embedSvgFonts(source, mock);
  assert.deepEqual(mock.requests.filter(r => r.url.startsWith('https://fonts.gstatic.com/')).map(r => r.url), [urls.korean]); assert.match(result, /font-weight:700/);
  assert.doesNotMatch(result, /src:url\(['"]https?:/);
});

test('허용되지 않은 폰트 URL·잘못된 바이너리는 출력 유지와 명시 경고로 처리한다', async () => {
  const source = `<svg><text font-family="${family}">A</text></svg>`, warnings = [];
  const mock = network(face('https://example.invalid/unsafe.woff2', 'U+0000-00FF'));
  assert.equal(await embedSvgFonts(source, { ...mock, onWarning: w => warnings.push(w) }), source);
  assert.equal(mock.requests.length, 2); assert.equal(warnings.length, 1); assert.match(warnings[0], /대체 글꼴/);
  const usage = createFontUsage(); addFontUsage(usage, family, 'A');
  const bad = await embedFontCss(usage, { fetchImpl: async url => new Response(String(url).includes('raw.githubusercontent') ? license : String(url).includes('googleapis') ? fontCss : '<html>not a font</html>') });
  assert.doesNotMatch(bad.css, /@font-face/); assert.equal(bad.warnings.length, 1);
});

test('폰트 로딩 취소는 경고로 삼켜 계속 저장하지 않는다', async () => {
  const controller = new AbortController(); controller.abort();
  const usage = createFontUsage(); addFontUsage(usage, family, 'A');
  await assert.rejects(embedFontCss(usage, { signal: controller.signal, fetchImpl: () => { throw Error('불필요한 요청'); } }), /취소/);
});

test('HTML Blob과 스트림 저장은 실제 셀·개체 웹폰트와 경고를 보존하고 문서를 수정하지 않는다', async () => {
  const wb = new Workbook({ defaultFont: { name: family, size: 11 }, sheets: [{ name: '합성', cells: { '0,0': { raw: '가A' } } }] });
  const before = JSON.stringify(wb.serialize()), version = wb.version, mock = network();
  const blob = await createSheetHtmlBlob(wb, 0, { fetchFont: mock.fetchImpl }), html = await blob.text();
  assert.match(html, /data-wixel-embedded-fonts/); assert.match(html, /font-src data:/); assert.match(html, /가A/);
  assert.doesNotMatch(html, /<link/); assert.equal(wb.version, version); assert.equal(JSON.stringify(wb.serialize()), before);
  const chunks = [], warnings = [];
  await writeSheetHtmlToSink(wb, 0, { write: bytes => chunks.push(bytes) }, { fetchFont: async () => { throw Error('offline'); }, onWarning: w => warnings.push(w) });
  const fallback = await new Blob(chunks).text(); assert.match(fallback, /가A/); assert.match(fallback, /role="alert"/); assert.match(fallback, /대체 글꼴/); assert.equal(warnings.length, 1);
  const svg = rangeImageSvg(wb, 0, { r1: 0, c1: 0, r2: 0, c2: 0 }).svg;
  assert.match(await embedSvgFonts(svg, network()), /data-wixel-embedded-fonts/);
  assert.equal(JSON.stringify(wb.serialize()), before);
});


test('라이선스 원문을 받지 못하면 폰트 바이너리를 재배포하지 않는다', async () => {
  const usage = createFontUsage(); addFontUsage(usage, family, '가A');
  const requests = [];
  const result = await embedFontCss(usage, { fetchImpl: async url => { requests.push(url); return new Response('not a license'); } });
  assert.equal(result.css, ''); assert.equal(result.warnings.length, 1);
  assert.equal(requests.length, 1); assert.match(requests[0], /^https:\/\/raw\.githubusercontent\.com\//);
});


test('SVG와 PDF XML의 라이선스 특수문자는 안전하게 보존한다', async () => {
  const source = `<svg xmlns="http://www.w3.org/2000/svg"><text font-family="${family}">A</text></svg>`;
  const result = await embedSvgFonts(source, network());
  assert.match(result, /PERMISSION &amp; CONDITIONS/);
  assert.doesNotMatch(result, /PERMISSION & CONDITIONS/);
  assert.equal(fontCssForSvg('a & b < c > d'), 'a &amp; b &lt; c &gt; d');
});

test('공식 카탈로그의 별도 라이선스 원문과 CDN 상대 주소도 내장한다', async () => {
  const entry = getWebFont('Jalnan 2'); assert.ok(entry.licenseText);
  const usage = createFontUsage(); addFontUsage(usage, entry.family, '가');
  const requests = [];
  const result = await embedFontCss(usage, { fetchImpl: async url => { requests.push(url); return new Response(bytes); } });
  assert.equal(result.warnings.length, 0); assert.match(result.css, /여기어때컴퍼니/);
  assert.match(result.css, /data:font/); assert.deepEqual(requests, [entry.faces[0].url]);
  const relative = network(face('//fonts.gstatic.com/s/fixture/relative.woff2', 'U+0000-00FF'));
  await embedSvgFonts(`<svg><text font-family="${family}">A</text></svg>`, relative);
  assert.ok(relative.requests.some(r => r.url === 'https://fonts.gstatic.com/s/fixture/relative.woff2'));
});


test('Ubuntu Font Licence 원문도 사용 폰트와 함께 동봉한다', async () => {
  const usage = createFontUsage(); addFontUsage(usage, 'Ubuntu', 'Abc123');
  const raw = 'UBUNTU FONT LICENCE Version 1.0\nCopyright Canonical Ltd.\nPERMISSION & CONDITIONS';
  const result = await embedFontCss(usage, { fetchImpl: async url => new Response(String(url).startsWith('https://raw.githubusercontent.com/') ? raw : String(url).startsWith('https://fonts.googleapis.com/') ? face(urls.latin,'U+0000-00FF') : bytes, {status:200}) });
  assert.deepEqual(result.warnings, []);
  assert.match(result.css, /UBUNTU FONT LICENCE/);
  assert.match(result.css, /data:font\/woff2/);
});
