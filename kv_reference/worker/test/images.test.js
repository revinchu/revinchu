import test from 'node:test';
import assert from 'node:assert/strict';
import { searchImages } from '../src/images.js';

const fixtures = {
  'api.openverse.org': { results: [{ title: 'Poster', thumbnail: 'https://api.openverse.org/t/1.jpg', foreign_landing_url: 'https://flickr.com/p/1', creator: 'A', width: 10, height: 20 }] },
  'commons.wikimedia.org': { query: { pages: { 1: { title: 'File:Old ad.jpg', imageinfo: [{ thumburl: 'https://upload.wikimedia.org/t.jpg', descriptionurl: 'https://commons.wikimedia.org/wiki/File:Old_ad.jpg', thumbwidth: 400, thumbheight: 300 }] } } } },
  'api.artic.edu': { config: { iiif_url: 'https://www.artic.edu/iiif/2' }, data: [{ id: 7, title: 'Poster', image_id: 'abc', artist_display: 'X' }, { id: 8, title: 'no image' }] },
  'collectionapi.metmuseum.org': null,
};
test('출처를 섞고 실패한 출처는 상태로만 보고', async () => {
  globalThis.fetch = async (url) => {
    const host = new URL(url).hostname;
    if (host === 'collectionapi.metmuseum.org') return new Response('down', { status: 503 });
    return new Response(JSON.stringify(fixtures[host]), { status: 200 });
  };
  const r = await searchImages('poster', [], 1, {});
  assert.equal(r.results.length, 3);
  assert.deepEqual(r.results.map((x) => x.source), ['Openverse', 'Wikimedia', 'Art Institute of Chicago']);
  assert.equal(r.results[2].thumb, 'https://www.artic.edu/iiif/2/abc/full/400,/0/default.jpg');
  assert.equal(r.sources.met.ok, false);
  assert.equal(r.sources.brave, undefined, '키 없는 출처는 꺼짐');
});
test('키가 있으면 Brave 가 켜진다', async () => {
  globalThis.fetch = async (url, init) => {
    assert.equal(init.headers['x-subscription-token'], 'k');
    return new Response(JSON.stringify({ results: [{ title: 'KV', url: 'https://behance.net/g/1', source: 'behance.net', thumbnail: { src: 'https://imgs.search.brave.com/x' }, properties: { url: 'https://mir.behance.net/x.jpg', width: 1, height: 1 } }] }));
  };
  const r = await searchImages('kv', ['brave'], 1, { BRAVE_API_KEY: 'k' });
  assert.equal(r.results[0].thumb, 'https://imgs.search.brave.com/x');
  assert.equal(r.results[0].source, 'behance.net');
});
