import test from 'node:test';
import assert from 'node:assert/strict';
import { fontDesktopStyle } from '../src/font-identity.js';
import { WEB_FONT_CATALOG } from '../src/web-font-catalog.js';
import { getWebFont, fontList, fontFamilyCandidates, fontLabel, fontMatches, fontInSource, fontIsKorean, webFontCssUrl, isAllowedFontUrl, requestWebFont, webFontStatus, onWebFontChange } from '../src/fonts.js';

test('무료 글꼴 목록은 기기 권한 없이 1,900종 이상이며 모든 항목을 검색한다', () => {
  const all = fontList();
  assert.ok(WEB_FONT_CATALOG.length >= 1900);
  assert.equal(new Set(all).size, all.length);
  for (const f of WEB_FONT_CATALOG) {
    assert.ok(all.includes(fontDesktopStyle(f.family).font), f.family);
    assert.equal(getWebFont(f.family).licenseUrl, f.licenseUrl);
    assert.ok(f.licenseUrl.startsWith('https://'), f.family);
    assert.ok(f.rawLicenseUrl?.startsWith('https://') || f.licenseText?.length > 100, f.family);
    assert.ok(f.source === 'google' || f.source === 'cdn');
  }
  assert.ok(all.includes('맑은 고딕'));
  assert.ok(all.includes('Verdana'));
});

test('한글·기존 영어 별칭은 실제 CSS 글꼴명으로 연결한다', () => {
  assert.equal(getWebFont('나눔고딕').family, 'NanumGothic');
  assert.equal(getWebFont('NanumGothic').family, 'NanumGothic');
  assert.ok(fontFamilyCandidates('나눔고딕').includes('NanumGothic'));
  assert.equal(fontMatches('Nanum Gothic', '나눔'), true);
  assert.equal(fontMatches('Nanum Gothic', 'NANUMGOTHIC'), true);
  assert.equal(fontIsKorean('Noto Sans KR'), true);
  assert.equal(fontInSource('Bungee','google'), true);
  assert.equal(fontInSource('Bungee','korean'), false);
  assert.equal(fontInSource('맑은 고딕','local'), true);
  assert.equal(fontLabel('Verdana'), 'Verdana');
});

test('실제 제공하는 굵기·기울임 조합만 CSS2에 요청하며 사용자 텍스트를 보내지 않는다', () => {
  const f = getWebFont('Cardo'), url = webFontCssUrl(f);
  assert.match(url, /^https:\/\/fonts.googleapis.com\/css2\?family=Cardo/);
  assert.ok(!url.includes('text='));
  assert.ok(!decodeURIComponent(url).includes('1,700'), 'Cardo has no bold italic');
  assert.match(decodeURIComponent(url), /1,400/);
  assert.equal(webFontCssUrl('document-unknown-font'), null);
  for (const f of WEB_FONT_CATALOG.filter(f=>f.source==='google')) {
    assert.ok(isAllowedFontUrl(webFontCssUrl(f)), f.family);
    const tuples = [...new URL(webFontCssUrl(f)).searchParams.get('family').matchAll(/(?:@|;)(\d),(\d+)/g)];
    for (const t of tuples) if (f.variants?.length) assert.ok(f.variants.includes(t[2] + (t[1]==='1'?'i':'')), f.family + t[0]);
  }
});

test('무료 글꼴 배포 주소는 정확한 HTTPS 출처만 허용한다', () => {
  for (const value of ['https://fonts.googleapis.com/css2?family=Inter','https://cdn.jsdelivr.net/gh/font/file.woff2']) assert.equal(isAllowedFontUrl(value), true);
  for (const value of ['javascript:alert(1)','http://fonts.googleapis.com/css','https://fonts.googleapis.com.evil.invalid/css','https://user:pass@cdn.jsdelivr.net/file','data:font/woff2;base64,AA','/font.woff2']) assert.equal(isAllowedFontUrl(value),false);
});

test('글꼴은 요청 시 한 번 로드하며 실패 후 명시적 재시도를 허용한다', async () => {
  const old = globalThis.document, nodes = [], loads = [], events = [];
  globalThis.document = {
    fonts: { addEventListener() {}, async load(spec,sample) { loads.push({spec,sample}); return [{status:'loaded'}]; } },
    createElement(tag) { assert.ok(['link','style'].includes(tag)); return {tag,dataset:{},remove(){this.removed=true;}}; },
    head: { append(node) { nodes.push(node); } },
  };
  const unsub = onWebFontChange(e=>events.push(e));
  try {
    const first = requestWebFont('Bungee');
    assert.equal(webFontStatus('Bungee'),'loading');
    assert.equal(requestWebFont('Bungee'),first);
    assert.equal(nodes.length,1);
    assert.equal(nodes[0].referrerPolicy,'no-referrer');
    assert.equal(nodes[0].crossOrigin,'anonymous');
    await nodes[0].onload();
    assert.equal((await first).status,'loaded');
    assert.equal(webFontStatus('Bungee'),'loaded');
    assert.ok(loads.every(x=>x.sample==='Aa0123'));
    assert.equal(requestWebFont('Bungee'), first);
    const bad = requestWebFont('Black Han Sans');
    nodes.at(-1).onerror();
    assert.equal((await bad).status,'error');
    assert.equal(requestWebFont('Black Han Sans'),bad);
    assert.ok(nodes.at(-1).removed);
    const retry = requestWebFont('Black Han Sans',{retry:true});
    assert.notEqual(retry,bad);
    await nodes.at(-1).onload();
    assert.equal((await retry).status,'loaded');
    assert.ok(loads.some(x=>x.sample==='가나다Aa0123'));
    assert.equal((await requestWebFont('Custom local')).status,'local');
    assert.deepEqual(events.map(e=>e.status),['loaded','error','loaded']);
  } finally { unsub(); if(old===undefined)delete globalThis.document;else globalThis.document=old; }
});


test('공식 가변 글꼴의 1~1000 굵기도 임의로 잘라내지 않는다', () => {
  const full = WEB_FONT_CATALOG.filter(f=>f.source==='google' && f.variants?.some(v=>v==='1'||v==='1000'));
  assert.ok(full.length>0);
  for (const f of full) {
    const family = new URL(webFontCssUrl(f)).searchParams.get('family');
    for (const v of f.variants) {
      const italic = family.includes(':ital,wght@');
      const entry = italic ? (v.endsWith('i')?'1,':'0,')+v.replace('i','') : v;
      assert.ok(family.split('@')[1].split(';').includes(entry), f.family + ': ' + v);
    }
  }
});
