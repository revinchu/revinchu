/* 링크 미리보기: 페이지 HTML에서 대표 이미지 URL을 찾는다 (og:image → twitter:image → image_src → 큰 첫 이미지). */

const PRIVATE_HOST = /^(localhost|.*\.local|.*\.internal|0\.0\.0\.0|127\.|10\.|192\.168\.|172\.(1[6-9]|2\d|3[01])\.|169\.254\.|\[?::1\]?|\[?f[cd][0-9a-f]{2}:)/i;

export function normalizeTarget(raw) {
  let u;
  try { u = new URL(raw); } catch (e) { return null; }
  if (!/^https?:$/.test(u.protocol) || PRIVATE_HOST.test(u.hostname) || u.port && !['80', '443'].includes(u.port)) return null;
  return u.toString();
}

function attrs(tag) {
  const out = {};
  tag.replace(/([a-zA-Z_:-]+)\s*=\s*("([^"]*)"|'([^']*)'|([^\s>]+))/g, (_, k, __, a, b, c) => { out[k.toLowerCase()] = decode(a ?? b ?? c ?? ''); return ''; });
  return out;
}
const decode = (s) => s.replace(/&amp;/g, '&').replace(/&quot;/g, '"').replace(/&#39;|&#x27;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>');

function youtube(url) {
  const m = url.match(/(?:youtube\.com\/(?:watch\?(?:.*&)?v=|shorts\/|embed\/)|youtu\.be\/)([\w-]{11})/);
  return m ? 'https://i.ytimg.com/vi/' + m[1] + '/hqdefault.jpg' : null;
}

export function extractPreview(html, pageUrl) {
  const abs = (u) => { try { return new URL(u, pageUrl).toString(); } catch (e) { return null; } };
  const site = (() => { try { return new URL(pageUrl).hostname.replace(/^www\./, ''); } catch (e) { return ''; } })();
  const head = html.slice(0, 400000);
  const metas = (head.match(/<meta\b[^>]*>/gi) || []).map(attrs);
  const meta = (...names) => {
    for (const n of names) {
      const m = metas.find((a) => (a.property || a.name || a.itemprop || '').toLowerCase() === n && a.content);
      if (m) return m.content;
    }
    return '';
  };
  const title = meta('og:title', 'twitter:title') || ((head.match(/<title[^>]*>([^<]*)<\/title>/i) || [])[1] || '').trim();
  let image = youtube(pageUrl) || meta('og:image:secure_url', 'og:image', 'og:image:url', 'twitter:image', 'twitter:image:src', 'image', 'thumbnailurl');
  if (!image) {
    const link = (head.match(/<link\b[^>]*rel\s*=\s*["']?image_src[^>]*>/i) || [])[0];
    if (link) image = attrs(link).href || '';
  }
  if (!image) {
    // 로고·아이콘·추적 픽셀을 피해서 크기가 적혀 있거나 이름이 그럴듯한 첫 이미지
    const imgs = (head.match(/<img\b[^>]*>/gi) || []).map(attrs);
    const ok = imgs.find((a) => {
      const src = a.src || a['data-src'] || '';
      if (!src || /^data:|\.svg(\?|$)|logo|icon|sprite|pixel|spacer|avatar|badge/i.test(src)) return false;
      const w = parseInt(a.width, 10);
      return !w || w >= 200;
    });
    if (ok) image = ok.src || ok['data-src'];
  }
  image = image ? abs(decode(image)) : null;
  return { image: image && /^https?:/.test(image) ? image : null, title: decode(title).slice(0, 200), site };
}
