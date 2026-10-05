// Display aliases, CSS family names and Windows/Excel family names are distinct.
import { WEB_FONT_CATALOG } from './web-font-catalog.js';
import { FONT_DESKTOP_IDENTITIES, FONT_DESKTOP_COMPAT_ALIASES } from './font-identity-data.js';
const key = value => String(value ?? '').trim().toLowerCase().replace(/\s+/g, '');
const nativeKey = value => String(value ?? '').trim().toLowerCase();
const catalog = new Map(), identities = new Map(), nativeNames = new Set();
for (const entry of WEB_FONT_CATALOG) for (const name of [entry.family, entry.label, ...(entry.aliases || [])]) if (name) catalog.set(key(name), entry);
for (const item of FONT_DESKTOP_IDENTITIES) {
  identities.set(item.family, item);
  for (const face of item.faces) for (const name of [face.font, ...face.aliases]) nativeNames.add(nativeKey(name));
}
// These are valid installed/localized Windows families, not WIXEL display labels.
for (const name of ['맑은 고딕','Malgun Gothic','나눔고딕','NanumGothic','나눔명조','NanumMyeongjo','나눔손글씨 펜','나눔고딕코딩','NanumGothicCoding','본고딕','Source Han Sans KR','본명조','Source Han Serif KR']) nativeNames.add(nativeKey(name));
const compatibility = new Map(FONT_DESKTOP_COMPAT_ALIASES.map(([a,b]) => [key(a),b]));
/** A patch only: callers merge it into their style. Never rewrite the source workbook. */
export function fontDesktopStyle(font, style = {}) {
  if (typeof font !== 'string' || !font.trim() || nativeNames.has(nativeKey(font))) return { font };
  const old = compatibility.get(key(font));
  if (old) return { font: old };
  const entry = catalog.get(key(font));
  if (!entry) return { font };
  const identity = identities.get(entry.family);
  if (!identity) return { font: entry.family };
  const chosen = style.bold ? identity.bold : identity.normal;
  return { font: chosen, ...(style.bold && identity.bold !== identity.normal ? { bold: false } : {}) };
}
export function fontIdentityForFamily(family) { return identities.get(family) ?? null; }
export function fontIdentityCompatibleAliases() { return FONT_DESKTOP_COMPAT_ALIASES; }
/** Normalize only explicit font edits; toggling bold alone must retain native-family semantics. */
export function fontIdentityPatch(patch, current = {}) {
  return patch && Object.hasOwn(patch, 'font') ? { ...patch, ...fontDesktopStyle(patch.font, { ...current, ...patch }) } : patch;
}
