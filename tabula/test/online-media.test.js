import test from 'node:test';
import assert from 'node:assert/strict';
import { normalizeStockMedia, safeMediaUrl, mediaSiteSearch } from '../src/stock-media.js';
import { searchOnlineImages, resolveOnlineMedia, trackOnlineMediaInsert } from '../src/online-images.js';
import { mediaSearch, trackMediaDownload } from '../server/media-search.js';
import { createWixelServer } from '../server/app.js';
const json = x => new Response(JSON.stringify(x), { headers: { 'content-type': 'application/json' } });
const params = (source, extra = {}) => new URLSearchParams({ source, q: '합성 & query', page: '1', type: 'image', ...extra });
const unsplash = { total: 31, results: [{ id: 'abc-12', urls: { regular: 'https://images.unsplash.com/full?ixid=keep', small: 'https://images.unsplash.com/thumb?ixid=keep' }, user: { name: '사진가', links: { html: 'https://unsplash.com/@artist' } }, links: { html: 'https://unsplash.com/photos/abc-12' }, alt_description: '풍경', width: 1600, height: 900 }] };
const pexels = { total_results: 1, photos: [{ id: 2, src: { original: 'https://images.pexels.com/full.jpg', medium: 'https://images.pexels.com/thumb.jpg' }, photographer: '작가', url: 'https://pexels.com/photo/2', alt: '그림' }] };
const pixabay = { totalHits: 1, hits: [{ id: 3, largeImageURL: 'https://pixabay.com/get/full.jpg', previewURL: 'https://cdn.pixabay.com/thumb.jpg', tags: '태그', user: '작가', pageURL: 'https://pixabay.com/photo/3' }] };
const wm = (mime, extra = {}) => ({ query: { pages: { 1: { pageid: 1, index: 1, title: 'File:합성', imageinfo: [{ mime, url: 'https://fixture.test/full.gif', thumburl: 'https://fixture.test/thumb.png', extmetadata: { LicenseUrl: { value: 'https://creativecommons.org/licenses/by/4.0/' } }, ...extra }] } } } });

test('새 공급자 그림 원본·출처·저작권과 링크/내장 정책', () => {
 const u = normalizeStockMedia('unsplash', unsplash)[0], p = normalizeStockMedia('pexels', pexels)[0], x = normalizeStockMedia('pixabay', pixabay)[0];
 assert.equal(u.full, unsplash.results[0].urls.regular); assert.equal(new URL(u.page).searchParams.get('utm_source'), 'wixel'); assert.equal(u.linkOnly, true); assert.equal(u.trackId, 'abc-12');
 assert.equal(p.full, pexels.photos[0].src.original); assert.equal(x.embedRequired, true); assert.ok([u,p,x].every(x => !x.cc && x.kind === 'image')); assert.notEqual(x.full, x.thumb);
});
test('Pexels/Pixabay 실제 영상rendition·poster·크기·길이', () => {
 const p = normalizeStockMedia('pexels', { videos: [{ id: 1, image: 'https://test.example/poster.jpg', duration: 7, user: { name: '제작자' }, video_files: [{ file_type: 'text/html', link: 'https://test.example/wrong' }, { file_type: 'video/mp4', link: 'https://test.example/4k.mp4', width: 3840 }, { file_type: 'video/mp4', link: 'https://test.example/hd.mp4', width: 1280, height: 720 }] }] }, 'video')[0];
 assert.equal(p.full, 'https://test.example/hd.mp4'); assert.equal(p.poster, 'https://test.example/poster.jpg'); assert.equal(p.duration, 7); assert.equal(p.width, 1280);
 const x = normalizeStockMedia('pixabay', { hits: [{ id: 1, videos: { large: { url: '' }, medium: { url: 'https://test.example/v.mp4', thumbnail: 'https://test.example/p.jpg', width: 1920, height: 1080 } } }] }, 'video')[0];
 assert.equal(x.full, 'https://test.example/v.mp4'); assert.equal(x.poster, 'https://test.example/p.jpg'); assert.equal(x.height, 1080);
});
test('미디어 URL/사이트 fallback는 실행 스킴·자격증명 차단, 검색어 인코딩', () => {
 for (const value of ['javascript:alert(1)', 'data:text/html,x', 'https://a:b@host.test/file', 'https://host.test/\nx']) assert.equal(safeMediaUrl(value), '');
 assert.ok(mediaSiteSearch('pexels', '고양이/꽃', 'video').includes('/search/videos/')); assert.ok(mediaSiteSearch('pixabay', 'a/b').includes('a%2Fb'));
 assert.deepEqual(normalizeStockMedia('unsplash', { results: [{ ...unsplash.results[0], urls: { regular: 'javascript:1', small: 'https://test.example/i' } }] }), []);
});
test('키 없음은 무요청+미설정 상태, 지원하지 않는 GIF는 별도 구분', async () => {
 for (const source of ['unsplash','pexels','pixabay']) { const r = await mediaSearch(params(source), {}, () => assert.fail('keyless request')); assert.equal(r.status, 'unconfigured'); assert.equal(r.items.length, 0); assert.ok(r.searchUrl.startsWith('https://')); }
 const r = await mediaSearch(params('pexels', { type: 'gif' }), {}, () => assert.fail('unsupported')); assert.equal(r.status, 'unsupported');
});
test('공급자 고정 URL·인증은 서버에만, 응답에는 키가 없음', async () => {
 for (const [source, keyName, body] of [['unsplash','UNSPLASH_ACCESS_KEY',unsplash],['pexels','PEXELS_API_KEY',pexels],['pixabay','PIXABAY_API_KEY',pixabay]]) {
   let call; const key = 'synthetic-secret-' + source;
   const result = await mediaSearch(params(source), { [keyName]: key }, async (url, opts) => { call = { u: new URL(url), opts }; return json(body); });
   assert.ok(['api.unsplash.com','api.pexels.com','pixabay.com'].includes(call.u.hostname)); assert.equal(call.opts.redirect, 'error'); assert.equal(call.opts.credentials, 'omit'); assert.equal(call.u.searchParams.get(source === 'pixabay' ? 'q' : 'query'), '합성 & query');
   assert.equal(source === 'pixabay' ? call.u.searchParams.get('key') : call.opts.headers.Authorization, source === 'unsplash' ? 'Client-ID ' + key : key); assert.ok(!JSON.stringify(result).includes(key)); assert.equal(result.items.length, 1);
 }
});
test('Pixabay 캐시 24시간 및 반환객체 변형 격리', async () => {
 let calls = 0; const fetcher = async () => { calls++; return json(pixabay); }, p = params('pixabay', { q: 'cache-only-synthetic' }), env = { PIXABAY_API_KEY: 'mock-cache' };
 const a = await mediaSearch(p, env, fetcher); a.items[0].title = 'changed'; const b = await mediaSearch(p, env, fetcher); assert.equal(calls, 1); assert.notEqual(b.items[0].title, 'changed');
 await mediaSearch(p, { PIXABAY_API_KEY: 'new-key' }, fetcher); assert.equal(calls, 2);
});
test('서버 검색 입력/오류/리디렉션/큰 응답 제한', async () => {
 for (const change of [{ source:'http://127.0.0.1' },{page:'51'},{page:'1.5'},{q:'x'.repeat(201)},{type:'html'}]) await assert.rejects(mediaSearch(params('unsplash', change), {}, () => assert.fail()), e => e.status === 400);
 await assert.rejects(mediaSearch(params('pixabay', { q:'x'.repeat(101) }), { PIXABAY_API_KEY:'test' }, () => assert.fail()), e => e.status === 400);
 for (const response of [new Response('',{status:302,headers:{location:'https://evil.test'}}),new Response('key leak',{status:401}),new Response('x',{headers:{'content-length':'3000000'}})]) await assert.rejects(mediaSearch(params('unsplash',{q:crypto.randomUUID()}), { UNSPLASH_ACCESS_KEY:'secret' }, async()=>response), e=>e.status===502 && !e.message.includes('key leak'));
});
test('서버 10초 제한은 취소를 무시하는 upstream에도 적용', async t => {
 t.mock.timers.enable({apis:['setTimeout']}); let signal;
 const pending = mediaSearch(params('pexels',{q:'timeout'}), {PEXELS_API_KEY:'test'}, async (u,o)=>{signal=o.signal;return new Promise(()=>{});});
 const check=assert.rejects(pending,e=>e.status===504); await Promise.resolve();t.mock.timers.tick(10000);await check;assert.equal(signal.aborted,true);
});
test('Unsplash 삽입 추적은 검증 ID 공식경로만 사용·키 미노출', async () => {
 let call;const r=await trackMediaDownload(new URLSearchParams({source:'unsplash',id:'abc-12'}),{UNSPLASH_ACCESS_KEY:'secret'},async(u,o)=>{call={u,o};return json({url:'https://image.test/x'});});
 assert.deepEqual(r,{ok:true});assert.equal(call.u,'https://api.unsplash.com/photos/abc-12/download');
 for(const id of ['../x','https://evil.test','a?x=2']) await assert.rejects(trackMediaDownload(new URLSearchParams({source:'unsplash',id}),{},()=>assert.fail()),e=>e.status===400);
});
test('GIF는 filemime 필터+응답 MIME 재검사로 PNG 제외·원본 보존', async () => {
 let u; const result=await searchOnlineImages('cat',{source:'wikimedia',kind:'gif',fetcher:async url=>{u=new URL(url);return json(wm('image/gif'));}});
 assert.match(u.searchParams.get('gsrsearch'),/filemime:image\/gif/);assert.equal(result.items[0].kind,'gif');assert.equal(result.items[0].full,'https://fixture.test/full.gif');
 const empty=await searchOnlineImages('cat',{source:'wikimedia',kind:'gif',fetcher:async()=>json(wm('image/png'))});assert.equal(empty.items.length,0);
});
test('Wikimedia 영상은 재생 가능한 derivative와 이미지poster 사용', async()=>{
 const response=wm('video/ogg',{url:'https://fixture.test/v.ogv',derivatives:[{type:'video/webm; codecs="vp9"',src:'https://fixture.test/v.webm',width:720,height:480}]});
 response.query.pages[1].videoinfo=response.query.pages[1].imageinfo;delete response.query.pages[1].imageinfo;
 let u; const r=await searchOnlineImages('ocean',{source:'wikimedia',kind:'video',ccOnly:true,fetcher:async url=>{u=new URL(url);return json(response);}});
 assert.equal(u.searchParams.get('prop'),'videoinfo');assert.equal(r.items[0].full,'https://fixture.test/v.webm');assert.equal(r.items[0].poster,'https://fixture.test/thumb.png');assert.equal(r.items[0].kind,'video');
});
test('NASA 영상 주소는 선택 시 지연해결하고 MP4만 반환', async()=>{
 const calls=[];const f=async u=>{calls.push(u);return json(u.includes('/asset/')?{collection:{items:[{href:'http://images-assets.nasa.gov/video/a~orig.mov'},{href:'http://images-assets.nasa.gov/video/a~medium.mp4'}]}}:{collection:{items:[{data:[{media_type:'video',nasa_id:'test id',title:'영상'}],links:[{rel:'preview',href:'https://fixture.test/thumb.jpg'}]}],metadata:{total_hits:1}}});};
 const r=await searchOnlineImages('moon',{source:'nasa',kind:'video',fetcher:f});assert.equal(calls.length,1);assert.equal(r.items[0].resolve,'nasa');const ready=await resolveOnlineMedia(r.items[0],{fetcher:f});assert.equal(calls.length,2);assert.match(calls[1],/test%20id$/);assert.equal(ready.full,'https://images-assets.nasa.gov/video/a~medium.mp4');assert.equal(ready.mime,'video/mp4');
});
test('CC 모드는 새 라이선스 공급자에 요청하지 않음, GIF all은Wikimedia만', async()=>{
 for(const source of ['unsplash','pexels','pixabay']){const r=await searchOnlineImages('x',{source,ccOnly:true,fetcher:()=>assert.fail('CC excluded')});assert.equal(r.items.length,0);}
 const calls=[];await searchOnlineImages('cat',{kind:'gif',fetcher:async u=>{calls.push(u);return json(wm('image/gif'));}});assert.equal(calls.length,1);assert.match(calls[0],/commons.wikimedia.org/);
});
test('프론트 중계 미설정은 failures가 아닌 unavailable·잘못된 URL 무시', async()=>{
 const r=await searchOnlineImages('cat',{source:'unsplash',fetcher:async()=>json({status:'unconfigured',message:'설정 필요'})});assert.deepEqual(r.failures,[]);assert.equal(r.unavailable[0].source,'unsplash');
 let request;const good=normalizeStockMedia('pexels',pexels)[0];const x=await searchOnlineImages('cat',{source:'pexels',fetcher:async(u,o)=>{request={u,o};return json({status:'ok',items:[good,{...good,full:'javascript:1'}],hasMore:false});}});assert.equal(x.items.length,1);assert.match(request.u,/^\/api\/media\/search/);assert.equal(request.o.credentials,'omit');
});
test('프론트 Unsplash 추적은 삽입 시만 별도호출',async()=>{
 let calls=0;await trackOnlineMediaInsert({source:'pexels'},{fetcher:()=>assert.fail()});await trackOnlineMediaInsert({source:'unsplash',trackId:'photo-1'},{fetcher:async u=>{calls++;assert.match(u,/\/api\/media\/track/);return json({ok:true});}});assert.equal(calls,1);
});
test('Node endpoint 인증·동일출처·rate limit·미설정 상태',async()=>{
 const app=createWixelServer({token:'server-token',mediaEnv:{},proxyRateLimit:1});await new Promise(resolve=>app.listen(0,'127.0.0.1',resolve));const base=`http://127.0.0.1:${app.address().port}`,path='/api/media/search?source=unsplash&q=cat';
 try{assert.equal((await fetch(base+path)).status,401);assert.equal((await fetch(base+path,{headers:{'X-Tabula-Token':'server-token',Origin:'https://evil.test'}})).status,403);const r=await fetch(base+path,{headers:{'X-Tabula-Token':'server-token'}});assert.equal(r.status,200);assert.equal((await r.json()).status,'unconfigured');assert.equal((await fetch(base+path,{headers:{'X-Tabula-Token':'server-token'}})).status,429);}finally{app.closeAllConnections();await new Promise(resolve=>app.close(resolve));}
});
