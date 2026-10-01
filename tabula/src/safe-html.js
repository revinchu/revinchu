// 통합 문서의 값/서식에서 만든 HTML은 신뢰하지 않는다. 비활성 template에서 검사한
// 노드만 그대로 이동하며, 정화 결과를 문자열로 직렬화하여 다시 파싱하지 않는다.
const HTML_NS = 'http://www.w3.org/1999/xhtml';
const SVG_NS = 'http://www.w3.org/2000/svg';
const HTML_TAGS = new Set('a abbr b bdi bdo blockquote br button canvas caption code col colgroup dd del details div dl dt em figcaption figure h1 h2 h3 h4 h5 h6 hr i img input kbd label li mark ol optgroup option p pre progress s samp section select small span strong sub summary sup table tbody td textarea tfoot th thead time tr u ul var'.split(' '));
const SVG_TAGS = new Set('svg g path rect circle ellipse line polyline polygon text tspan defs lineargradient radialgradient stop filter clippath mask marker pattern use image title desc fedropshadow fegaussianblur fecolormatrix feoffset feflood fecomposite femerge femergenode femorphology fecomponenttransfer fefuncr fefuncg fefuncb fefunca feblend fetile feturbulence fedisplacementmap feconvolvematrix fediffuselighting fespecularlighting fedistantlight fepointlight fespotlight'.split(' '));
const URL_ATTRS = new Set('href src xlink:href poster background cite action formaction ping srcset imagesrcset longdesc manifest'.split(' '));
const PAINT_ATTRS = new Set('fill stroke filter clip-path mask marker marker-start marker-mid marker-end cursor'.split(' '));

/** 목적에 맞는 주소만 허용. 문자열의 제어문자로 프로토콜 검사를 우회하지 못하게 한다. */
export function safeUrl(value, purpose = 'link', base = 'https://wixel.invalid/') {
  const raw = String(value ?? '').trim();
  if (!raw || /[\u0000-\u0020\u007f]/.test(raw) || raw.includes('\\')) return null;
  if (raw.startsWith('#')) return raw;
  if (purpose === 'reference') return null;
  if (purpose === 'image' && /^data:image\/(?:png|jpeg|gif|webp|avif|bmp|x-icon|svg\+xml)(?:;[a-z0-9=+._-]+)*,/i.test(raw)) return raw;
  let url;
  try { url = new URL(raw, base); } catch { return null; }
  if (url.username || url.password) return null;
  if (url.protocol === 'http:' || url.protocol === 'https:') return raw;
  if (purpose === 'image' && url.protocol === 'blob:') return raw;
  if (purpose === 'link' && (url.protocol === 'mailto:' || url.protocol === 'tel:')) return raw;
  return null;
}

// CSSOM으로 유효한 선언만 남긴다. CSS 이스케이프/주석을 이용한 url·식 위장도 거부한다.
function safeCss(value, base) {
  if (/[\\\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(value) || /\/\*|expression\s*\(|@import|(?:^|[;\s])(?:behavior|-moz-binding)\s*:/i.test(value)) return false;
  let remainder = value;
  const urlPattern = /url\(\s*(?:"([^"]*)"|'([^']*)'|([^\s)'" ]+))\s*\)/gi;
  let match;
  while ((match = urlPattern.exec(value))) {
    // 같은 조각의 SVG 정의와 셀 무늬용 이미지 데이터만 참조한다.
    const url = match[1] ?? match[2] ?? match[3];
    if (!safeUrl(url, 'reference', base) && !(/^data:image\//i.test(url) && safeUrl(url, 'image', base))) return false;
    remainder = remainder.replace(match[0], '');
  }
  return !/url\s*\(|image-set\s*\(|-webkit-image-set\s*\(/i.test(remainder);
}

/** 안전한 DocumentFragment. 여기에서만 문자열을 파싱하며 반환값은 직접 append한다. */
export function sanitizeHtml(html, doc = document) {
  const template = doc.createElement('template');
  template.innerHTML = String(html ?? '');
  const base = doc.baseURI;
  const visit = (parent) => {
    for (const node of [...parent.childNodes]) {
      if (node.nodeType === 3) continue;
      if (node.nodeType !== 1) { node.remove(); continue; }
      const tag = node.localName.toLowerCase();
      const svg = node.namespaceURI === SVG_NS;
      if (!(svg ? SVG_TAGS.has(tag) : node.namespaceURI === HTML_NS && HTML_TAGS.has(tag))) { node.remove(); continue; }
      for (const attr of [...node.attributes]) {
        const name = attr.name.toLowerCase();
        const local = attr.localName.toLowerCase();
        const value = attr.value;
        let remove = name.startsWith('on') || local.startsWith('on') || ['srcdoc', 'is', 'autofocus', 'form', 'formaction', 'slot', 'xmlns', 'xmlns:xlink'].includes(name);
        // HTML id/name은 전역 변수 및 폼 속성 가리기를 막는다. SVG의 도형/필터 id는 유지.
        if (!svg && (name === 'id' || name === 'name')) remove = true;
        if (attr.namespaceURI && attr.namespaceURI !== 'http://www.w3.org/1999/xlink') remove = true;
        if (URL_ATTRS.has(name) || local === 'href') {
          let purpose = null;
          if ((tag === 'img' || tag === 'image') && (local === 'src' || local === 'href')) purpose = 'image';
          else if ((tag === 'a') && local === 'href') purpose = 'link';
          else if (svg && tag === 'use' && local === 'href') purpose = 'reference';
          if (!purpose || !safeUrl(value, purpose, base)) remove = true;
        }
        if (name === 'style') {
          if (!safeCss(value, base)) remove = true;
          else node.setAttribute('style', node.style.cssText);
        }
        if (svg && PAINT_ATTRS.has(name) && !safeCss(value, base)) remove = true;
        if (remove) node.removeAttributeNode(attr);
      }
      if (tag === 'a') node.setAttribute('rel', 'noopener noreferrer');
      visit(node);
    }
  };
  visit(template.content);
  return template.content;
}

export function setSafeHtml(node, html) {
  node.replaceChildren(sanitizeHtml(html, node.ownerDocument));
}
