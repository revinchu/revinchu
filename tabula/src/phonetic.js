import { child, kids, allText, esc } from './xml.js';

const TYPES = new Set(['halfwidthKatakana', 'fullwidthKatakana', 'Hiragana', 'noConversion']);
const ALIGNMENTS = new Set(['noControl', 'left', 'center', 'distributed']);
export function normalizePhonetic(value, text, strict = false) {
  if (!value || typeof value !== 'object') return undefined;
  text = String(text ?? '');
  const fail = message => { if (strict) throw new Error(message); };
  const runs = [];
  let end = 0;
  for (const item of [...(Array.isArray(value.runs) ? value.runs : [])].sort((a, b) => Number(a.sb) - Number(b.sb))) {
    const sb = Number(item.sb), eb = Number(item.eb), hint = String(item.text ?? '');
    if (!hint) continue;
    const splitsPair = at => at > 0 && at < text.length && /[\uD800-\uDBFF]/.test(text[at - 1]) && /[\uDC00-\uDFFF]/.test(text[at]);
    if (!Number.isInteger(sb) || !Number.isInteger(eb) || sb < end || sb < 0 || eb <= sb || eb > text.length || splitsPair(sb) || splitsPair(eb)) { fail('윗주 구간은 원문 안에서 서로 겹치지 않게 지정하세요.'); continue; }
    if (hint.length > 32767 || runs.length >= 1024) { fail('윗주가 너무 깁니다. 구간은 1,024개 이내로 지정하세요.'); continue; }
    runs.push({ sb, eb, text: hint }); end = eb;
  }
  const font = {}, f = value.font ?? {};
  if (typeof f.font === 'string' && f.font) font.font = f.font.slice(0, 255);
  if (Number.isFinite(Number(f.size)) && Number(f.size) > 0) font.size = Math.min(409, Number(f.size));
  for (const key of ['bold', 'italic', 'underline', 'strike']) if (f[key]) font[key] = true;
  if (/^#[a-f\d]{6}$/i.test(f.color ?? '')) font.color = f.color;
  return { runs, visible: value.visible === true, type: TYPES.has(value.type) ? value.type : 'noConversion', alignment: ALIGNMENTS.has(value.alignment) ? value.alignment : 'left', ...(Object.keys(font).length ? { font } : {}) };
}
export function phoneticText(text, value) {
  text = String(text ?? '');
  const p = normalizePhonetic(value, text);
  if (!p?.runs.length) return text;
  let out = '', at = 0;
  for (const run of p.runs) { out += text.slice(at, run.sb) + run.text; at = run.eb; }
  return out + text.slice(at);
}
export function readPhonetic(node, text, fonts = []) {
  const runs = kids(node, 'rPh').map(r => ({ sb: Number(r.attrs.sb), eb: Number(r.attrs.eb), text: allText(r) }));
  const pr = child(node, 'phoneticPr');
  if (!runs.length && !pr) return undefined;
  return normalizePhonetic({ runs, type: pr?.attrs.type, alignment: pr?.attrs.alignment, font: fonts[Number(pr?.attrs.fontId)] }, text);
}
export function phoneticXml(text, value, fontId, encode = esc) {
  const p = normalizePhonetic(value, text);
  if (!p) return '';
  return p.runs.map(r => `<rPh sb="${r.sb}" eb="${r.eb}"><t xml:space="preserve">${encode(r.text)}</t></rPh>`).join('')
    + `<phoneticPr fontId="${fontId}" type="${p.type}" alignment="${p.alignment}"/>`;
}
export function phoneticHtml(text, value) {
  const p = normalizePhonetic(value, text);
  if (!p?.visible || !p.runs.length) return null;
  const f = p.font ?? {}, css = [];
  if (f.size) css.push(`font-size:${f.size}pt`);
  if (f.font) css.push(`font-family:'${f.font.replace(/[^\p{L}\p{N} _-]/gu, '')}'`);
  if (f.color) css.push(`color:${f.color}`);
  if (f.bold) css.push('font-weight:700'); if (f.italic) css.push('font-style:italic');
  if (f.underline || f.strike) css.push(`text-decoration:${f.underline ? 'underline ' : ''}${f.strike ? 'line-through' : ''}`);
  const segment = (base, hint = '') => `<span class="phonetic-segment"><span class="phonetic-guide" style="${esc(css.join(';'))}" aria-hidden="true">${hint ? esc(hint) : '&#160;'}</span><span class="phonetic-base">${esc(base)}</span></span>`;
  let at = 0, html = '';
  for (const r of p.runs) { if (r.sb > at) html += segment(text.slice(at, r.sb)); html += segment(text.slice(r.sb, r.eb), r.text); at = r.eb; }
  if (at < text.length) html += segment(text.slice(at));
  return `<span class="cell-phonetic phonetic-${p.alignment}" aria-label="${esc(text)}">${html}</span>`;
}
