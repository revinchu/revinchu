// 선택한 무료 웹폰트만 출력 문서에 담는다. 문서 문자열은 서버로 전송하지 않는다.
import { getWebFont, webFontCssUrl, isAllowedFontUrl, fontFamilyCandidates } from './fonts.js';
import { parseFontFaces, fontIdentityFaces } from './font-faces.js';

const binaryCache = new Map(), licenseCache = new Map();
let cachedBytes = 0;
const CACHE_BYTES = 16 * 1024 * 1024, FONT_BYTES = 24 * 1024 * 1024;
const cssString = value => String(value ?? '').replace(/[\\'"\r\n\f<>]/g, '');
const decode = value => String(value ?? '').replace(/&(#x[\da-f]+|#\d+|amp|lt|gt|quot|apos|nbsp);/gi, (all, entity) => {
  const named = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' }, key = entity.toLowerCase();
  if (key[0] !== '#') return named[key] ?? all;
  const point = key[1] === 'x' ? parseInt(key.slice(2), 16) : Number(key.slice(1));
  return point > 0 && point <= 0x10ffff ? String.fromCodePoint(point) : '\ufffd';
});
const firstFamily = value => {
  const text = decode(value).trim(), quoted = /^(['"])(.*?)\1/.exec(text);
  return (quoted ? quoted[2] : text.split(',')[0]).trim();
};
const abort = signal => { if (signal?.aborted) throw new Error('글꼴을 포함한 파일 저장을 취소했습니다.'); };

export function createFontUsage() { return new Map(); }
export function addFontUsage(usage, name, text, style = {}) {
  const requested = firstFamily(name), entry = getWebFont(requested);
  if (!entry || text === null || text === undefined || text === '') return usage;
  const family = entry.family ?? entry.name ?? requested;
  let record = usage.get(family);
  if (!record) { record = { entry, family, names: new Set([family]), points: new Set(), variants: new Map() }; usage.set(family, record); }
  record.names.add(requested);
  for (const alias of fontFamilyCandidates(requested)) if (getWebFont(alias) === entry) record.names.add(alias);
  for (const character of String(text)) record.points.add(character.codePointAt(0));
  const weight = Number(style.weight ?? (style.bold ? 700 : 400)) || (style.weight === 'bold' ? 700 : 400);
  const italic = style.italic || style.style === 'italic' || style.style === 'oblique';
  record.variants.set(`${weight}:${italic ? 'italic' : 'normal'}`, { weight, style: italic ? 'italic' : 'normal' });
  return usage;
}

/** SVG/출력 HTML의 글꼴 상속을 따라 실제 텍스트에 쓰인 family·문자를 수집한다. */
export function collectMarkupFontUsage(markup, usage = createFontUsage()) {
  const stack = [{ tag: '', font: '', weight: 400, style: 'normal', skip: false }];
  const tokens = String(markup).matchAll(/<!--[\s\S]*?-->|<![^>]*>|<\/?[A-Za-z][^>]*>|[^<]+/g);
  for (const [token] of tokens) {
    if (token[0] !== '<') { const current = stack.at(-1); if (!current.skip) addFontUsage(usage, current.font, decode(token), current); continue; }
    if (/^<!/.test(token)) continue;
    const tag = /^<\/?([\w:-]+)/.exec(token)?.[1].toLowerCase(); if (!tag) continue;
    if (/^<\//.test(token)) { for (let i = stack.length - 1; i > 0; i--) if (stack[i].tag === tag) { stack.length = i; break; } continue; }
    const attrs = {};
    for (const m of token.matchAll(/([\w:-]+)\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))/g)) attrs[m[1].toLowerCase()] = decode(m[2] ?? m[3] ?? m[4]);
    const current = { ...stack.at(-1), tag }, css = attrs.style ?? '';
    const property = name => new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([^;]+)`, 'i').exec(css)?.[1]?.trim();
    current.font = property('font-family') ?? attrs['font-family'] ?? current.font;
    const shorthand = property('font');
    if (shorthand) { const family = /\b\d[\d.]*(?:px|pt|em|rem|%)\s*(?:\/[^\s]+\s*)?(.+)$/i.exec(shorthand)?.[1]; if (family) current.font = family; }
    current.weight = property('font-weight') ?? attrs['font-weight'] ?? (/^(b|strong)$/.test(tag) ? 700 : current.weight);
    current.style = property('font-style') ?? attrs['font-style'] ?? (/^(i|em)$/.test(tag) ? 'italic' : current.style);
    current.skip ||= /^(script|style|title|desc)$/.test(tag);
    if (!/\/\s*>$/.test(token) && !/^(br|img|input|meta|link|hr|col|source|wbr)$/.test(tag)) stack.push(current);
  }
  return usage;
}

function rangeContains(range, points) {
  if (!range) return true;
  for (const item of String(range).split(',')) {
    const match = /^\s*U\+([\da-f?]+)(?:-([\da-f]+))?\s*$/i.exec(item); if (!match) continue;
    const first = parseInt(match[1].replace(/\?/g, '0'), 16), last = parseInt(match[2] ?? match[1].replace(/\?/g, 'f'), 16);
    for (const point of points) if (point >= first && point <= last) return true;
  }
  return false;
}
function weightDistance(face, weight) {
  const values = String(face.weight ?? '400').trim().split(/\s+/).map(n => n === 'normal' ? 400 : n === 'bold' ? 700 : Number(n));
  const start = Number.isFinite(values[0]) ? values[0] : 400, end = Number.isFinite(values[1]) ? values[1] : start;
  return weight < start ? start - weight : weight > end ? weight - end : 0;
}
function usedFaces(faces, record) {
  const selected = new Set();
  for (const variant of record.variants.values()) {
    const matching = faces.filter(f => String(f.style ?? 'normal').startsWith(variant.style));
    const pool = matching.length ? matching : faces;
    let distance = Infinity; for (const face of pool) distance = Math.min(distance, weightDistance(face, variant.weight));
    for (const face of pool) if (weightDistance(face, variant.weight) === distance && rangeContains(face.unicodeRange ?? face.unicode, record.points)) selected.add(face);
  }
  return [...selected];
}
async function fetchAsset(url, { signal, fetchImpl }, css = false, licenseEntry = null) {
  abort(signal);
  const allowed = value => {
    if (!licenseEntry) return isAllowedFontUrl(value);
    try { const u = new URL(value); return value === licenseEntry.rawLicenseUrl && u.protocol === 'https:' && !u.username && !u.password; } catch { return false; }
  };
  if (!allowed(url)) throw new Error('허용되지 않은 글꼴 주소');
  const controller = new AbortController(), cancel = () => controller.abort();
  signal?.addEventListener('abort', cancel, { once: true });
  const timer = setTimeout(cancel, 15000);
  try {
    const response = await fetchImpl(url, { credentials: 'omit', referrerPolicy: 'no-referrer', signal: controller.signal });
    if (!response.ok || (response.url && !allowed(response.url))) throw new Error('글꼴 다운로드 실패');
    if (css) { const text = await response.text(); if (text.length > 2 * 1024 * 1024) throw new Error('글꼴 스타일시트가 너무 큽니다.'); return text; }
    const data = new Uint8Array(await response.arrayBuffer());
    if (!data.length || data.length > FONT_BYTES) throw new Error('글꼴 파일 크기가 올바르지 않습니다.');
    const signature = String.fromCharCode(...data.subarray(0, 4));
    if (!['wOF2', 'wOFF', 'OTTO', 'ttcf', '\u0000\u0001\u0000\u0000'].includes(signature)) throw new Error('글꼴 파일 형식이 올바르지 않습니다.');
    return data;
  } finally { clearTimeout(timer); signal?.removeEventListener('abort', cancel); }
}
function base64(bytes) {
  let text = ''; for (let at = 0; at < bytes.length; at += 8192) text += String.fromCharCode(...bytes.subarray(at, at + 8192));
  return btoa(text);
}
async function fontData(url, options) {
  const cached = !options.fetchImplOverride && binaryCache.get(url); if (cached) return cached;
  const bytes = await fetchAsset(url, options), sig = String.fromCharCode(...bytes.subarray(0, 4));
  const result = { data: `data:${sig === 'wOF2' ? 'font/woff2' : sig === 'wOFF' ? 'font/woff' : 'font/ttf'};base64,${base64(bytes)}`, bytes: bytes.length };
  if (!options.fetchImplOverride && bytes.length <= CACHE_BYTES) {
    while (cachedBytes + bytes.length > CACHE_BYTES && binaryCache.size) { const key = binaryCache.keys().next().value; cachedBytes -= binaryCache.get(key).bytes; binaryCache.delete(key); }
    binaryCache.set(url, result); cachedBytes += bytes.length;
  }
  return result;
}

async function fontLicense(entry, options) {
  // URL은 검증된 카탈로그에 기록된 그 주소만 사용한다. 문서의 문자열은 URL이 될 수 없다.
  if (getWebFont(entry.family) !== entry || (!entry.rawLicenseUrl && !entry.licenseText)) throw new Error('재배포 라이선스가 없습니다.');
  const key = entry.rawLicenseUrl || entry.family;
  const cached = !options.fetchImplOverride && licenseCache.get(key); if (cached) return cached;
  const text = entry.licenseText || await fetchAsset(entry.rawLicenseUrl, options, true, entry);
  if (!entry.licenseText && (!/SIL OPEN FONT LICENSE|Apache License|UBUNTU FONT LICEN[CS]E/i.test(text) || /^\s*</.test(text))) throw new Error('라이선스 원문을 확인하지 못했습니다.');
  const comment = `/* WIXEL embedded font: ${cssString(entry.family)}\nLicense source: ${entry.rawLicenseUrl || entry.licenseUrl || entry.homepage || entry.family}\n${text.replace(/\*\//g, '* /').replace(/</g, '\\3c ').replace(/>/g, '\\3e ')}\n*/`;
  if (!options.fetchImplOverride) { if (licenseCache.size >= 64) licenseCache.delete(licenseCache.keys().next().value); licenseCache.set(key, comment); }
  return comment;
}

export async function embedFontCss(usage, { onWarning, signal, fetchImpl = globalThis.fetch } = {}) {
  const css = [], warnings = [], assets = new Map(), options = { signal, fetchImpl, fetchImplOverride: fetchImpl !== globalThis.fetch };
  const warn = family => {
    const message = `'${family}' 웹폰트를 파일에 포함하지 못했습니다. 이 파일에서는 대체 글꼴이 보일 수 있습니다.`;
    if (!warnings.includes(message)) { warnings.push(message); onWarning?.(message); }
  };
  for (const record of usage.values()) {
    abort(signal);
    try {
      const entry = record.entry, url = webFontCssUrl(entry);
      const license = await fontLicense(entry, options);
      const faces = fontIdentityFaces(entry, entry.faces?.length ? entry.faces : url ? parseFontFaces(await fetchAsset(url, options, true), url) : []);
      const selected = usedFaces(faces, record);
      if (!selected.length) { warn(record.family); continue; }
      const familyCss = [];
      // 한글의 unicode-range 조각은 사용 문자에 해당하는 것만 포함한다. 동시에 네 개까지 받는다.
      for (let at = 0; at < selected.length; at += 4) {
        const blocks = await Promise.all(selected.slice(at, at + 4).map(async face => {
        try {
          if (!assets.has(face.url)) assets.set(face.url, fontData(face.url, options));
          const data = await assets.get(face.url), weight = /^[\d ]+$/.test(String(face.weight ?? '400')) ? String(face.weight ?? '400') : face.weight === 'bold' ? '700' : '400';
          const style = /^(normal|italic|oblique)$/.test(face.style ?? '') ? face.style : 'normal';
          const range = face.unicodeRange ?? face.unicode;
          const unicode = range && /^U\+[\da-f? -]+(?:\s*,\s*U\+[\da-f? -]+)*$/i.test(range) ? `unicode-range:${range};` : '';
          return [...record.names].map(name => `@font-face{font-family:'${cssString(name)}';src:url('${data.data}');font-weight:${weight};font-style:${style};font-display:swap;${unicode}}`).join('\n');
        } catch { abort(signal); warn(record.family); return ''; }
        }));
        for (const block of blocks) if (block) familyCss.push(block);
      }
      if (familyCss.length) { css.push(license); for (const block of familyCss) css.push(block); }
    } catch { abort(signal); warn(record.family); }
  }
  return { css: css.join('\n'), warnings };
}

// SVG <style>은 HTML raw-text가 아니라 XML 텍스트이므로 라이선스의 &도 이스케이프한다.
export function fontCssForSvg(css) { return String(css).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

export async function embedSvgFonts(svg, options = {}) {
  if (/data-wixel-embedded-fonts\s*=/.test(svg)) return svg;
  const result = await embedFontCss(collectMarkupFontUsage(svg), options);
  if (!result.css) return svg;
  return String(svg).replace(/<svg\b[^>]*>/i, root => `${root}<style data-wixel-embedded-fonts="true">${fontCssForSvg(result.css)}</style>`);
}
