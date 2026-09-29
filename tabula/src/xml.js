// 아주 작은 XML 파서/작성 도우미 (xlsx 용, DOM 없이 동작)

const ENTITIES = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };
/** 엑셀 문자열의 _xHHHH_ 이스케이프 (줄 바꿈 _x000D_ · _x000A_ 등, 글자 그대로의 '_x' 는 _x005F_x) */
export function unx(s) {
  return typeof s === 'string' && s.includes('_x') ? s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, h) => String.fromCharCode(parseInt(h, 16))) : s;
}

export function decodeEntities(s) {
  // XML 줄 끝 정규화: 글자 그대로 있는 CR LF · CR 은 LF (엑셀과 같음). &#13; 으로 적은 CR 은 아래에서 그대로 남음
  if (s.includes('\r')) s = s.replace(/\r\n?/g, '\n');
  if (!s.includes('&')) return s;
  return s.replace(/&(#x[0-9a-f]+|#\d+|\w+);/gi, (m, e) => {
    if (e[0] === '#') {
      const code = e[1].toLowerCase() === 'x' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10);
      return Number.isFinite(code) ? String.fromCodePoint(code) : m;
    }
    return ENTITIES[e] ?? m;
  });
}

export function esc(s) {
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
    // XML 1.0에서 허용되지 않는 제어 문자 제거
    .replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f]/g, '');
}

const localName = (n) => { const i = n.indexOf(':'); return i < 0 ? n : n.slice(i + 1); };

/**
 * XML 문자열 → 요소 트리 { name, attrs, children, text }
 * name 은 네임스페이스 접두사를 뗀 이름, attrs 는 원래 이름 그대로
 */
export function parseXml(str) {
  const root = { name: '#root', attrs: {}, children: [], text: '' };
  const stack = [root];
  const re = /<!\[CDATA\[([\s\S]*?)\]\]>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!(?:[^>[]|\[[^\]]*\])*>|<(\/?)([^\s/>]+)((?:\s+[^\s=/>]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m;
  while ((m = re.exec(str))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) { top.text += m[1]; continue; }
    if (m[6] !== undefined) { top.text += unx(decodeEntities(m[6])); continue; }
    if (m[3] === undefined) continue; // 주석/선언
    if (m[2] === '/') {
      if (stack.length > 1) stack.pop();
      continue;
    }
    const attrs = {};
    const ar = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let a;
    while ((a = ar.exec(m[4]))) attrs[a[1]] = decodeEntities(a[2] ?? a[3]);
    const el = { name: localName(m[3]), attrs, children: [], text: '' };
    top.children.push(el);
    if (!m[5]) stack.push(el);
  }
  return root.children[0] ?? root;
}

export const child = (el, name) => el?.children.find((c) => c.name === name) ?? null;
export const kids = (el, name) => (el ? el.children.filter((c) => c.name === name) : []);

/** 하위 모든 요소 중 이름이 같은 것 */
export function descendants(el, name, out = []) {
  if (!el) return out;
  for (const c of el.children) {
    if (c.name === name) out.push(c);
    descendants(c, name, out);
  }
  return out;
}

/** 요소 안의 모든 <t> 텍스트 (서식 있는 텍스트 포함) */
export function allText(el) {
  if (!el) return '';
  if (el.name === 't') return el.text;
  let s = '';
  for (const c of el.children) {
    if (c.name === 'rPh' || c.name === 'phoneticPr') continue; // 윗주 제외
    s += allText(c);
  }
  return s;
}
