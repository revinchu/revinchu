import test from 'node:test';
import assert from 'node:assert/strict';
import { ONLINE_IMAGE_SOURCES, searchOnlineImages } from '../src/online-images.js';

const photo = (id, extra = {}) => ({ id, url: `https://example.test/original-${id}.jpg`, thumbnail: `https://example.test/thumb-${id}.jpg`, title: `사진 ${id}`, creator: '사진가', license: 'by', license_url: 'https://creativecommons.org/licenses/by/4.0/', foreign_landing_url: `https://example.test/page/${id}`, ...extra });
const response = json => ({ ok: true, json: async () => json });
const search = (json, options = {}) => searchOnlineImages('고양이', { source: 'openverse', fetcher: async () => response(json), ...options });

test('출처 계약과 Openverse 원본/썸네일 분리, 검색 URL 인코딩 및 요청 개인정보 옵션', async () => {
  assert.deepEqual(ONLINE_IMAGE_SOURCES.map(s => s.id), ['all', 'openverse', 'wikimedia', 'inaturalist', 'nasa', 'unsplash', 'pexels', 'pixabay']);
  let request;
  const result = await searchOnlineImages('고양이 & 꽃', { source: 'openverse', page: 2, ccOnly: true, fetcher: async (url, options) => { request = { url: new URL(url), options }; return response({ results: [photo(7, { source: 'flickr', provider: 'flickr' })], page_count: 3 }); } });
  assert.equal(request.url.searchParams.get('q'), '고양이 & 꽃'); assert.equal(request.url.searchParams.get('page'), '2');
  assert.equal(request.url.searchParams.get('page_size'), '20'); assert.match(request.url.searchParams.get('license'), /cc0/);
  assert.equal(request.options.credentials, 'omit'); assert.equal(request.options.method, 'GET'); assert.equal(request.options.mode, 'cors');
  assert.ok(request.options.signal instanceof AbortSignal); assert.equal(result.hasMore, true);
  assert.deepEqual(result.items[0], { id: 'openverse:7', thumb: 'https://example.test/thumb-7.jpg', full: 'https://example.test/original-7.jpg', title: '사진 7', credit: '사진가 · CC BY', page: 'https://example.test/page/7', license: 'CC BY', licenseUrl: 'https://creativecommons.org/licenses/by/4.0/', cc: true, source: 'openverse', provider: 'flickr' });
});

test('엄격한 CC 판정: BY 계열/CC0만, PD/PDM·유사 호스트·알 수 없는 URL 제외', async () => {
  const cases = [
    ['by', '', true], ['by-sa', '', true], ['by-nc', '', true], ['by-nd', '', true], ['by-nc-sa', '', true], ['by-nc-nd', '', true], ['cc0', '', true],
    ['pd', '', false], ['pdm', 'https://creativecommons.org/publicdomain/mark/1.0/', false], ['cc0', 'https://creativecommons.org/publicdomain/zero/1.0/', true],
    ['by', 'https://creativecommons.org/licenses/by/4.0/', true], ['by-sa', 'https://creativecommons.org/licenses/by-sa/3.0/de/', true],
    ['by', 'https://creativecommons.org.evil.test/licenses/by/4.0/', false], ['by', 'javascript:alert(1)', false], ['by', 'https://creativecommons.org/licenses/by/99.0/', false], ['unknown', '', false],
  ];
  const result = await search({ results: cases.map(([license, license_url], i) => photo(i, { license, license_url })), page_count: 1 });
  assert.deepEqual(result.items.map(x => x.cc), cases.map(x => x[2]));
  const filtered = await search({ results: cases.map(([license, license_url], i) => photo(i, { license, license_url })), page_count: 1 }, { ccOnly: true });
  assert.equal(filtered.items.length, cases.filter(x => x[2]).length); assert.ok(filtered.items.every(x => x.cc));
});

test('외부 URL은 http(s)만, 원본 주소가 잘못되면 썸네일로 속여 삽입하지 않는다', async () => {
  const result = await search({ results: [photo(1, { url: 'javascript:alert(1)' }), photo(2, { url: 'data:image/png;base64,a' }), photo(3, { url: 'file:///tmp/a.png' }), photo(4, { url: 'https://name:secret@example.test/a.jpg' }), photo(5, { thumbnail: 'javascript:alert(1)', foreign_landing_url: 'data:text/html,a' }), photo(6, { url: 'https://example.test/\na.jpg' })] });
  assert.equal(result.items.length, 1); assert.equal(result.items[0].id, 'openverse:5');
  assert.equal(result.items[0].thumb, result.items[0].full); assert.equal(result.items[0].page, '');
});

test('Wikimedia 검색 순서·CC URL 우선·원본 URL·HTML 저자 제거·페이지 계속', async () => {
  let url;
  const result = await searchOnlineImages('꽃', { page: 2, source: 'wikimedia', ccOnly: true, fetcher: async u => { url = new URL(u); return response({ continue: { gsroffset: 60 }, query: { pages: {
    9: { pageid: 9, index: 2, title: 'File:나무.jpg', imageinfo: [{ url: 'https://upload.wikimedia.org/tree.jpg', thumburl: 'https://upload.wikimedia.org/thumb/tree.jpg', descriptionurl: 'https://commons.wikimedia.org/wiki/File:Tree.jpg', extmetadata: { License: { value: 'cc-by-sa-4.0' }, LicenseShortName: { value: 'CC BY-SA 4.0' }, LicenseUrl: { value: 'https://creativecommons.org/licenses/by-sa/4.0/' }, Artist: { value: '<a href="/artist">A &amp; B</a><script>x()</script>' } } }] },
    8: { pageid: 8, index: 1, title: 'File:공공.jpg', imageinfo: [{ url: 'https://upload.wikimedia.org/public.jpg', extmetadata: { License: { value: 'cc-by' }, LicenseUrl: { value: 'https://creativecommons.org/publicdomain/mark/1.0/' } } }] },
  } } }); } });
  assert.equal(url.searchParams.get('gsroffset'), '30'); assert.equal(url.searchParams.get('gsrlimit'), '30');
  assert.equal(result.items.length, 1); assert.equal(result.items[0].full, 'https://upload.wikimedia.org/tree.jpg');
  assert.equal(result.items[0].credit, 'A & B · CC BY-SA 4.0'); assert.equal(result.hasMore, true);
  const empty = await searchOnlineImages('없음', { source: 'wikimedia', fetcher: async () => response({ batchcomplete: '' }) });
  assert.deepEqual(empty, { items: [], hasMore: false, failures: [] });
});

test('iNaturalist는 관찰 라이선스 대신 사진별 CC 검증, 숨긴 사진 제외·공식 사진 경로 원본', async () => {
  let url;
  const json = { total_results: 41, results: [{ id: 1, license_code: 'cc0', taxon: { preferred_common_name: '고양이' }, photos: [
    { id: 1, license_code: null, url: 'https://static.inaturalist.org/photos/1/square.jpg' },
    { id: 2, license_code: 'cc-by-sa', url: 'https://inaturalist-open-data.s3.amazonaws.com/photos/2/square.jpg', attribution: '사진가 (CC BY-SA)' },
    { id: 3, license_code: 'cc0', hidden: true, url: 'https://static.inaturalist.org/photos/3/square.jpg' },
    { id: 4, license_code: 'cc0', url: 'https://external.test/photos/4/square.jpg' },
  ] }] };
  const options = { source: 'inaturalist', page: 2, ccOnly: true, fetcher: async u => { url = new URL(u); return response(json); } };
  const result = await searchOnlineImages('고양이', options);
  assert.equal(url.searchParams.get('per_page'), '20'); assert.match(url.searchParams.get('photo_license'), /cc-by-sa/);
  assert.deepEqual(result.items.map(x => x.id), ['inaturalist:2', 'inaturalist:4']);
  assert.match(result.items[0].thumb, /\/medium.jpg$/); assert.match(result.items[0].full, /\/original.jpg$/);
  assert.equal(result.items[1].full, 'https://external.test/photos/4/square.jpg'); assert.equal(result.hasMore, true);
  const all = await searchOnlineImages('고양이', { ...options, ccOnly: false });
  assert.equal(all.items[0].cc, false); assert.equal(all.items[0].licenseUrl, '');
});

test('Wikimedia는 이미지 MIME만 사용하며 PDF·동영상 썸네일을 원본 그림으로 삽입하지 않음', async () => {
  let request;
  const result = await searchOnlineImages('문서', { source: 'wikimedia', fetcher: async url => { request = new URL(url); return response({ query: { pages: Object.fromEntries(['application/pdf', 'video/webm', 'image/svg+xml', 'image/jpeg'].map((mime, i) => [i, { pageid: i, imageinfo: [{ mime, url: `https://example.test/full-${i}`, thumburl: `https://example.test/thumb-${i}.jpg` }] }])) } }); } });
  assert.match(request.searchParams.get('iiprop'), /mime/);
  assert.deepEqual(result.items.map(x => x.id), ['wikimedia:2', 'wikimedia:3']);
});

test('iNaturalist 한 페이지는 관찰 20개와 사진 40개 이내', async () => {
  const result = await searchOnlineImages('꽃', { source: 'inaturalist', fetcher: async () => response({ total_results: 50, results: Array.from({ length: 30 }, (_, o) => ({ photos: Array.from({ length: 5 }, (_, p) => ({ id: o * 5 + p, url: `https://static.inaturalist.org/photos/${o * 5 + p}/square.jpg`, license_code: 'cc0' })) })) }) });
  assert.equal(result.items.length, 40); assert.equal(result.hasMore, true);
});

test('NASA canonical 원본·큰 alternate 우선, CC 전용 검색에서는 요청하지 않음', async () => {
  const items = [
    { data: [{ title: '달', nasa_id: 'Moon 1', media_type: 'image', center: 'JPL' }], links: [{ rel: 'preview', render: 'image', href: 'https://images-assets.nasa.gov/thumb.jpg' }, { rel: 'canonical', render: 'image', href: 'https://images-assets.nasa.gov/original.jpg' }] },
    { data: [{ title: '별', nasa_id: 'Star', media_type: 'image' }], links: [{ rel: 'alternate', render: 'image', width: 500, height: 500, href: 'https://images-assets.nasa.gov/small.jpg' }, { rel: 'alternate', render: 'image', width: 2000, height: 1500, href: 'https://images-assets.nasa.gov/large.jpg' }] },
  ];
  const result = await searchOnlineImages('moon', { source: 'nasa', fetcher: async () => response({ collection: { items, metadata: { total_hits: 31 } } }) });
  assert.equal(result.items[0].full, 'https://images-assets.nasa.gov/original.jpg'); assert.equal(result.items[0].thumb, 'https://images-assets.nasa.gov/thumb.jpg');
  assert.equal(result.items[0].page, 'https://images.nasa.gov/details/Moon%201'); assert.equal(result.items[0].cc, false);
  assert.equal(result.items[1].full, 'https://images-assets.nasa.gov/large.jpg'); assert.equal(result.hasMore, true);
  const filtered = await searchOnlineImages('moon', { source: 'nasa', ccOnly: true, fetcher: () => assert.fail('NASA must not be fetched') });
  assert.deepEqual(filtered, { items: [], hasMore: false, failures: [] });
});

test('연합 검색은 병렬 실행·교차 배치·중복 원본 URL 제거하고 일부 실패를 알림', async () => {
  const started = [], completions = [];
  const pending = searchOnlineImages('꽃', { fetcher: url => String(url).startsWith('/api/media/') ? Promise.resolve(response({ status: 'unconfigured', items: [], hasMore: false })) : new Promise((resolve, reject) => { started.push(new URL(url).hostname); completions.push({ resolve, reject }); }) });
  await Promise.resolve(); await Promise.resolve();
  assert.equal(started.length, 4);
  completions[0].resolve(response({ results: [photo(1), photo(2)], page_count: 2 }));
  completions[1].reject(Error('offline'));
  completions[2].resolve(response({ results: [{ photos: [{ id: 9, license_code: 'cc0', url: 'https://example.test/original-1.jpg#different' }, { id: 8, license_code: 'cc0', url: 'https://example.test/inat.jpg' }] }], total_results: 1 }));
  completions[3].resolve(response({ collection: { items: [{ data: [{ nasa_id: 'n', media_type: 'image' }], links: [{ rel: 'canonical', render: 'image', href: 'https://example.test/nasa.jpg' }] }], metadata: { total_hits: 1 } } }));
  const result = await pending;
  assert.deepEqual(result.items.map(x => x.id), ['openverse:1', 'nasa:n', 'openverse:2', 'inaturalist:8']);
  assert.deepEqual(result.failures, ['Wikimedia Commons']); assert.equal(result.hasMore, true);
});

test('CC 연합은 NASA 제외하고 응답을 다시 필터, 마지막 50페이지는 더 보기 없음', async () => {
  const hosts = [];
  const result = await searchOnlineImages('꽃', { ccOnly: true, page: 50, fetcher: async url => { hosts.push(new URL(url).hostname); if (url.includes('openverse')) return response({ results: [photo(1), photo(2, { license: 'pd', license_url: '' })], page_count: 100 }); if (url.includes('wikimedia')) return response({ batchcomplete: '' }); return response({ results: [], total_results: 0 }); } });
  assert.equal(hosts.length, 3); assert.ok(hosts.every(h => !h.includes('nasa')));
  assert.equal(result.items.length, 1); assert.equal(result.hasMore, false);
});

test('전 사이트 오류는 한국어로 실패하며 HTTP/오류 JSON/잘못된 JSON 형식을 삼키지 않음', async () => {
  await assert.rejects(searchOnlineImages('꽃', { fetcher: async () => ({ ok: false, status: 503 }) }), /그림 검색 사이트에 연결하지 못했습니다/);
  await assert.rejects(search({ error: 'rate limited' }), /연결하지 못했습니다/);
  await assert.rejects(search({ results: null }), /연결하지 못했습니다/);
  await assert.rejects(searchOnlineImages('꽃', { source: 'nasa', fetcher: async () => ({ ok: true, json: async () => { throw Error('invalid JSON'); } }) }), /연결하지 못했습니다/);
});

test('취소는 부분 성공보다 우선, 요청 신호 중단·늦게 온 결과 미사용', async () => {
  const controller = new AbortController(), calls = [];
  const pending = searchOnlineImages('꽃', { signal: controller.signal, fetcher: (url, options) => { calls.push(options); return new Promise(() => {}); } });
  await Promise.resolve(); await Promise.resolve(); assert.equal(calls.length, 7);
  controller.abort(); await assert.rejects(pending, e => e.name === 'AbortError' && /취소/.test(e.message));
  assert.ok(calls.every(c => c.signal.aborted));
  await assert.rejects(searchOnlineImages('꽃', { signal: controller.signal, fetcher: () => assert.fail('pre-aborted') }), { name: 'AbortError' });
});

test('10초 시간 제한은 응답 JSON을 기다리는 동안도 적용되고 요청을 중단함', async t => {
  t.mock.timers.enable({ apis: ['setTimeout'] });
  let requestSignal;
  const pending = searchOnlineImages('꽃', { source: 'openverse', fetcher: async (url, options) => { requestSignal = options.signal; return { ok: true, json: () => new Promise(() => {}) }; } });
  await Promise.resolve(); await Promise.resolve();
  const check = assert.rejects(pending, /연결하지 못했습니다/);
  t.mock.timers.tick(10000); await check; assert.equal(requestSignal.aborted, true);
});

test('검색 입력 제약, 빈 검색 무요청 및 결과 건수 제한', async () => {
  const fetcher = () => assert.fail('잘못된 입력은 요청하면 안 됨');
  for (const page of [0, 51, 1.5, '1', NaN]) await assert.rejects(searchOnlineImages('꽃', { page, fetcher }), /1부터 50/);
  await assert.rejects(searchOnlineImages('꽃'.repeat(201), { fetcher }), /200자/);
  await assert.rejects(searchOnlineImages('꽃', { source: 'https://bad.test', fetcher }), /지원하지 않는/);
  assert.deepEqual(await searchOnlineImages('   ', { fetcher }), { items: [], hasMore: false, failures: [] });
  const result = await search({ results: Array.from({ length: 50 }, (_, i) => photo(i)), page_count: 2 }); assert.equal(result.items.length, 20);
});
