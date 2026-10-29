import test from 'node:test';
import assert from 'node:assert/strict';
import { extractPreview, normalizeTarget } from '../src/preview.js';

test('og:image 우선, 상대 경로는 절대 경로로', () => {
  const html = '<html><head><meta property="og:title" content="Nike &amp; Co"><meta property="og:image" content="/img/kv.jpg"></head></html>';
  assert.deepEqual(extractPreview(html, 'https://www.example.com/a/b'), { image: 'https://www.example.com/img/kv.jpg', title: 'Nike & Co', site: 'example.com' });
});
test('twitter:image, 따옴표 없는 속성', () => {
  const r = extractPreview('<meta name=twitter:image content=https://cdn.x.com/1.png>', 'https://x.com/p');
  assert.equal(r.image, 'https://cdn.x.com/1.png');
});
test('메타가 없으면 로고가 아닌 큰 첫 이미지', () => {
  const html = '<img src="/logo.png"><img src="/icon.svg"><img src="/t.gif" width="1"><img src="/kv/main.jpg" width="800">';
  assert.equal(extractPreview(html, 'https://s.com/').image, 'https://s.com/kv/main.jpg');
});
test('유튜브는 썸네일 주소를 바로 만든다', () => {
  assert.equal(extractPreview('', 'https://www.youtube.com/watch?v=dQw4w9WgXcQ').image, 'https://i.ytimg.com/vi/dQw4w9WgXcQ/hqdefault.jpg');
});
test('사설 주소와 비표준 포트는 거부', () => {
  for (const u of ['http://localhost/', 'http://127.0.0.1/', 'http://192.168.0.1/', 'http://10.1.2.3/', 'file:///etc/passwd', 'https://a.com:8443/', 'http://[::1]/']) assert.equal(normalizeTarget(u), null, u);
  assert.equal(normalizeTarget('https://www.behance.net/gallery/1/x'), 'https://www.behance.net/gallery/1/x');
});
