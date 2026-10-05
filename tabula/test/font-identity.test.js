import test from 'node:test';
import assert from 'node:assert/strict';
import { fontDesktopStyle, fontIdentityPatch } from '../src/font-identity.js';
import { fontList, getWebFont, fontMatches, requestWebFont, webFontCssUrl } from '../src/fonts.js';
import { parseFontFaces, fontIdentityFaces } from '../src/font-faces.js';
import { createFontUsage, addFontUsage, embedFontCss } from '../src/font-export.js';

test('G마켓 별칭은 Medium, 굵게는 실제 Bold 가족으로 저장한다',()=>{
  for(const font of ['지마켓 산스','G마켓 산스','Gmarket Sans','GmarketSans']) {
    assert.deepEqual(fontDesktopStyle(font),{font:'Gmarket Sans TTF Medium'});
    assert.deepEqual(fontDesktopStyle(font,{bold:true}),{font:'Gmarket Sans TTF Bold',bold:false});
  }
  for(const font of ['Gmarket Sans TTF Light','Gmarket Sans TTF Medium','Gmarket Sans TTF Bold','Gmarket Sans Bold','G마켓 산스 TTF Light']) assert.deepEqual(fontDesktopStyle(font,{bold:true}),{font});
});
test('다른 실제 설치명·알 수 없는 글꼴은 바꾸지 않고 CSS/표시용 별칭만 교정한다',()=>{
  for(const [input,output] of [['나눔스퀘어 네오','NanumSquare Neo Regular'],['cafe24Ssurround','Cafe24 Ssurround Bold'],['Cafe24ProSlimFit','Cafe24 PRO Slim Fit'],['cafe24Proup','카페24 프로업'],['Nanum Gothic','NanumGothic'],['Nanum Pen Script','Nanum Pen'],['프리텐다드','Pretendard']]) assert.equal(fontDesktopStyle(input).font,output,input);
  for(const font of ['나눔고딕','나눔명조','맑은 고딕','Malgun Gothic','Private Brand Font 2026','Gmarket Sans TTF Medium']) assert.deepEqual(fontDesktopStyle(font,{bold:false}),{font});
  assert.deepEqual(fontDesktopStyle('Pretendard',{bold:true}),{font:'Pretendard'});
  assert.equal(fontDesktopStyle('NanumSquare Neo TTF Light').font,'NanumSquare Neo Light');
});
test('새 글꼴만 현재 굵기를 고려해 패치하고 굵게 버튼 단독·빈 기본값은 보존한다',()=>{
  assert.deepEqual(fontIdentityPatch({font:'지마켓 산스',italic:true},{bold:true}),{font:'Gmarket Sans TTF Bold',italic:true,bold:false});
  const toggle={bold:true}; assert.equal(fontIdentityPatch(toggle,{font:'Gmarket Sans TTF Medium'}),toggle);
  assert.deepEqual(fontIdentityPatch({font:undefined},{bold:true}),{font:undefined});
});
test('글꼴 목록은 설치 가능한 세 가지 G마켓·다섯 가지 나눔 네오 굵기를 제공한다',()=>{
  const list=fontList();
  for(const suffix of ['Light','Medium','Bold']) {const f='Gmarket Sans TTF '+suffix;assert.ok(list.includes(f));assert.ok(fontMatches(f,'지마켓'));}
  for(const suffix of ['Light','Regular','Bold','ExtraBold','Heavy']) assert.ok(list.includes('NanumSquare Neo '+suffix));
  assert.ok(!list.includes('Gmarket Sans'));
  assert.ok(!list.includes('지마켓 산스'));
});
test('설치용 Bold와 Light 가족은 서로 다른 실제 웹 파일을 CSS Regular로 사용한다',()=>{
  for(const suffix of ['Light','Medium','Bold']) {
    const f=getWebFont('Gmarket Sans TTF '+suffix);
    assert.equal(f.faces.length,1);assert.equal(f.faces[0].weight,400);assert.ok(f.faces[0].url.endsWith('GmarketSans'+suffix+'.woff2'));
  }
  assert.ok(getWebFont('NanumSquare Neo ExtraBold').faces[0].url.includes('-dEb.woff'));
  assert.equal(getWebFont('Wanted Sans ExtraBlack').fixedWeight,1000);
});
test('CSS 원본 이름과 설치용 이름을 분리해 Google 요청을 유지한다',()=>{
  assert.ok(decodeURIComponent(webFontCssUrl('NanumGothic')).includes('Nanum+Gothic'));
  assert.ok(decodeURIComponent(webFontCssUrl('NanumGothicExtraBold')).includes('wght@800'));
  const f=getWebFont('Pretendard Medium');assert.equal(f.fixedWeight,500);assert.equal(f.sourceFamily,'Pretendard');
  const parsed=parseFontFaces('@font-face{font-family:Test;src:url(regular.woff2);font-weight:400;}@font-face{font-family:Test;src:url(medium.woff2);font-weight:500;unicode-range:U+AC00-D7A3;}', 'https://cdn.jsdelivr.net/a/css');
  const faces=fontIdentityFaces(f,parsed);assert.equal(faces.length,1);assert.equal(faces[0].weight,400);assert.equal(faces[0].url,'https://cdn.jsdelivr.net/a/medium.woff2');assert.equal(faces[0].unicodeRange,'U+AC00-D7A3');
});
test('HTML/SVG 내장 글꼴도 Bold 가족의 실제 파일을 저장하며 라이선스를 포함한다',async()=>{
  const usage=createFontUsage();addFontUsage(usage,'Gmarket Sans TTF Bold','가나다');const fetched=[];
  const result=await embedFontCss(usage,{fetchImpl:async url=>{fetched.push(url);return {ok:true,url,arrayBuffer:async()=>new TextEncoder().encode('wOF2fixture').buffer,text:async()=> 'SIL OPEN FONT LICENSE Version 1.1'};}});
  assert.equal(result.warnings.length,0);assert.ok(fetched.some(u=>u.endsWith('GmarketSansBold.woff2')));assert.ok(!fetched.some(u=>u.endsWith('GmarketSansMedium.woff2')));assert.match(result.css,/font-weight:400/);assert.match(result.css,/Gmarket Sans TTF Bold/);
});
test('고정 굵기 CSS의 비동기 로딩도 실제 Medium만 새 가족으로 등록한다',async()=>{
  const oldDocument=globalThis.document, oldFetch=globalThis.fetch,nodes=[],specs=[];
  globalThis.document={createElement:()=>({dataset:{},remove(){}}),head:{append(n){nodes.push(n);}},fonts:{async load(spec){specs.push(spec);return [{}];}}};
  globalThis.fetch=async url=>({ok:true,url,text:async()=> '@font-face{font-family:SUIT;src:url(https://cdn.jsdelivr.net/medium.woff2);font-weight:500;}@font-face{font-family:SUIT;src:url(https://cdn.jsdelivr.net/bold.woff2);font-weight:700;}'});
  try {const result=await requestWebFont('SUIT Medium');assert.equal(result.status,'loaded');assert.equal(nodes.length,1);assert.match(nodes[0].textContent,/font-family:"SUIT Medium"/);assert.match(nodes[0].textContent,/medium.woff2/);assert.ok(!nodes[0].textContent.includes('bold.woff2'));assert.ok(specs.every(s=>s.startsWith('400 ')));}
  finally{if(oldDocument===undefined)delete globalThis.document;else globalThis.document=oldDocument;globalThis.fetch=oldFetch;}
});
