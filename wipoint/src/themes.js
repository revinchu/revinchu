// 테마(색 · 글꼴)와 색 계산 — DOM 없음
// 색 값: '#RRGGBB' 직접 색, 또는 '@이름' 테마 색 (예 '@accent1', '@tx1', '@bg1'),
// 테마 색 뒤에 밝기 변형을 붙일 수 있음: '@accent1:lm75' (lumMod 75%), ':lo25' (lumOff), ':t40' (tint), ':s50' (shade), ':a50' (투명도 50%)

export const SLOT_NAMES = ['dk1', 'lt1', 'dk2', 'lt2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];

const T = (name, label, colors, major, minor, extra = {}) => ({
  name, label,
  colors: Object.fromEntries(SLOT_NAMES.map((k, i) => [k, colors[i]])),
  fonts: { major, minor },
  ...extra,
});

/** 기본 제공 테마 (원본 색 구성, PowerPoint 기본 'Office 테마'와 같은 색 순서) */
export const THEMES = [
  T('office', 'Office 테마', ['#000000', '#FFFFFF', '#0E2841', '#E8E8E8', '#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E', '#467886', '#96607D'], '맑은 고딕', '맑은 고딕'),
  T('office2013', 'Office 2013', ['#000000', '#FFFFFF', '#44546A', '#E7E6E6', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72'], '맑은 고딕', '맑은 고딕'),
  T('navy', '네이비 비즈니스', ['#1B1B1B', '#FFFFFF', '#14213D', '#F2F4F8', '#1F4E9E', '#F5A623', '#2BA59A', '#E4572E', '#7A5CC4', '#8AA1B1', '#1F4E9E', '#7A5CC4'], '맑은 고딕', '맑은 고딕', { bg: '@bg1', titleBar: true }),
  T('mint', '민트 프레시', ['#22313F', '#FFFFFF', '#0B5D5E', '#E9F7F4', '#18A999', '#F2B134', '#2F80ED', '#EB5757', '#6FCF97', '#9B51E0', '#18A999', '#9B51E0'], '맑은 고딕', '맑은 고딕', { accentBand: true }),
  T('sunset', '선셋', ['#2B2118', '#FFFBF5', '#5C2E0E', '#FFF1E0', '#E85D04', '#F48C06', '#D00000', '#FFBA08', '#6A040F', '#9D0208', '#E85D04', '#6A040F'], '맑은 고딕', '맑은 고딕', { accentBand: true }),
  T('slate', '슬레이트 다크', ['#FFFFFF', '#1E2329', '#E6E9ED', '#2B323B', '#4FC3F7', '#FFB74D', '#81C784', '#E57373', '#BA68C8', '#90A4AE', '#4FC3F7', '#BA68C8'], '맑은 고딕', '맑은 고딕', { dark: true }),
  T('forest', '포레스트', ['#1E2A1E', '#FFFFFF', '#254D32', '#EEF3EA', '#3A7D44', '#9DC08B', '#C9A227', '#A44A3F', '#5B8E7D', '#69995D', '#3A7D44', '#A44A3F'], '맑은 고딕', '맑은 고딕', { titleBar: true }),
  T('berry', '베리', ['#2A1E2B', '#FFFFFF', '#4A154B', '#F7EEF7', '#A4266F', '#E8A33D', '#3D8EB9', '#5FB49C', '#7B3FA0', '#DE6E4B', '#A4266F', '#7B3FA0'], '맑은 고딕', '맑은 고딕', { accentBand: true }),
  T('mono', '모노크롬', ['#111111', '#FFFFFF', '#333333', '#F2F2F2', '#404040', '#7F7F7F', '#A6A6A6', '#262626', '#595959', '#BFBFBF', '#0563C1', '#954F72'], '맑은 고딕', '맑은 고딕'),
  T('naver', '그린 마케팅', ['#1A1A1A', '#FFFFFF', '#03361F', '#F1FAF4', '#03C75A', '#1EC800', '#00A495', '#FF6F3C', '#4B6BFB', '#8E8E8E', '#03C75A', '#4B6BFB'], '맑은 고딕', '맑은 고딕', { titleBar: true }),
];

export const themeByName = (name) => THEMES.find((t) => t.name === name) ?? null;

/** 테마 복사본 (문서가 자기 테마를 따로 가짐) */
export function cloneTheme(t) {
  return { name: t.name, label: t.label, colors: { ...t.colors }, fonts: { ...t.fonts }, dark: !!t.dark, accentBand: !!t.accentBand, titleBar: !!t.titleBar };
}

// ───────────── 색 계산 ─────────────
const ALIAS = { tx1: 'dk1', bg1: 'lt1', tx2: 'dk2', bg2: 'lt2', dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2' };

export function hexToRgb(hex) {
  let h = String(hex).replace('#', '');
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  const n = parseInt(h.slice(0, 6), 16);
  if (!Number.isFinite(n)) return [0, 0, 0];
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
export const rgbToHex = (r, g, b) => `#${[r, g, b].map((v) => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase()}`;

function rgbToHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const mx = Math.max(r, g, b);
  const mn = Math.min(r, g, b);
  const l = (mx + mn) / 2;
  if (mx === mn) return [0, 0, l];
  const d = mx - mn;
  const s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
  let h;
  if (mx === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (mx === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}
function hslToRgb(h, s, l) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

const toLinear = (c) => { const x = c / 255; return x <= 0.04045 ? x / 12.92 : ((x + 0.055) / 1.055) ** 2.4; };
const fromLinear = (l) => 255 * (l <= 0.0031308 ? l * 12.92 : 1.055 * l ** (1 / 2.4) - 0.055);

/** OOXML 색 변형 적용: mods = [[종류, 값(0~1)]] (lumMod, lumOff, tint, shade, alpha) → { hex, alpha } */
export function applyMods(hex, mods = []) {
  let [r, g, b] = hexToRgb(hex);
  let alpha = 1;
  for (const [k, v] of mods) {
    if (k === 'lumMod' || k === 'lumOff') {
      const [h, s, l] = rgbToHsl(r, g, b);
      const nl = k === 'lumMod' ? l * v : l + v;
      [r, g, b] = hslToRgb(h, s, Math.max(0, Math.min(1, nl)));
    } else if (k === 'tint') {
      // PowerPoint 는 tint/shade 를 선형 RGB 에서 계산 (검정 tint 75% = #898989)
      [r, g, b] = [r, g, b].map((c) => fromLinear(toLinear(c) * v + (1 - v)));
    } else if (k === 'shade') {
      [r, g, b] = [r, g, b].map((c) => fromLinear(toLinear(c) * v));
    } else if (k === 'alpha') alpha = v;
  }
  return { hex: rgbToHex(r, g, b), alpha };
}

const MOD_CODES = { lm: 'lumMod', lo: 'lumOff', t: 'tint', s: 'shade', a: 'alpha' };
const MOD_SHORT = Object.fromEntries(Object.entries(MOD_CODES).map(([a, b]) => [b, a]));

/** '@accent1:lm75:lo25' → { slot, mods } */
export function parseColorRef(c) {
  if (typeof c !== 'string') return null;
  const [base, ...rest] = c.split(':');
  const mods = rest.map((m) => { const x = /^([a-z]+)(-?[\d.]+)$/.exec(m); return x && MOD_CODES[x[1]] ? [MOD_CODES[x[1]], Number(x[2]) / 100] : null; }).filter(Boolean);
  if (base.startsWith('@')) return { slot: ALIAS[base.slice(1)] ?? base.slice(1), mods };
  return { hex: base, mods };
}
export function colorRef(slotOrHex, mods = []) {
  const head = slotOrHex.startsWith('#') ? slotOrHex.toUpperCase() : `@${slotOrHex}`;
  return [head, ...mods.map(([k, v]) => `${MOD_SHORT[k]}${Math.round(v * 1000) / 10}`)].join(':');
}

/** 색 값 → CSS 색 (테마 색 해석) */
export function resolveColor(theme, c, fallback = null) {
  if (c === null || c === undefined || c === '' || c === 'none') return fallback;
  const p = parseColorRef(c);
  if (!p) return fallback;
  const hex = p.slot ? theme?.colors?.[p.slot] ?? '#000000' : p.hex;
  if (!p.mods.length) return hex;
  const { hex: out, alpha } = applyMods(hex, p.mods);
  if (alpha >= 1) return out;
  const [r, g, b] = hexToRgb(out);
  return `rgba(${r},${g},${b},${Math.round(alpha * 1000) / 1000})`;
}

/** 밝은 배경 위인지 (글자 색 자동 결정) */
export function isDark(cssColor) {
  const m = /rgba?\((\d+),(\d+),(\d+)/.exec(cssColor ?? '');
  const [r, g, b] = m ? [Number(m[1]), Number(m[2]), Number(m[3])] : hexToRgb(cssColor ?? '#FFFFFF');
  return (r * 299 + g * 587 + b * 114) / 1000 < 140;
}

/** 색 고르기 표: 테마 색 10개 × (기본 + 밝게 80/60/40 + 어둡게 25/50) — PowerPoint 색 상자와 같은 배열 */
export function themePaletteRows() {
  const cols = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'];
  const lighter = (slot) => (slot === 'lt1' ? [[['lumMod', 0.95]], [['lumMod', 0.85]], [['lumMod', 0.75]], [['lumMod', 0.65]], [['lumMod', 0.5]]]
    : slot === 'dk1' ? [[['lumMod', 0.5], ['lumOff', 0.5]], [['lumMod', 0.65], ['lumOff', 0.35]], [['lumMod', 0.75], ['lumOff', 0.25]], [['lumMod', 0.85], ['lumOff', 0.15]], [['lumMod', 0.95], ['lumOff', 0.05]]]
      : [[['lumMod', 0.2], ['lumOff', 0.8]], [['lumMod', 0.4], ['lumOff', 0.6]], [['lumMod', 0.6], ['lumOff', 0.4]], [['lumMod', 0.75]], [['lumMod', 0.5]]]);
  const rows = [cols.map((s) => colorRef(s))];
  for (let i = 0; i < 5; i++) rows.push(cols.map((s) => colorRef(s, lighter(s)[i])));
  return rows;
}
export const STANDARD_COLORS = ['#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050', '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0'];

/** 테마 색 이름 (도움말 표시) */
export const SLOT_LABEL = { lt1: '배경 1', dk1: '텍스트 1', lt2: '배경 2', dk2: '텍스트 2', accent1: '강조 1', accent2: '강조 2', accent3: '강조 3', accent4: '강조 4', accent5: '강조 5', accent6: '강조 6', hlink: '하이퍼링크', folHlink: '열어 본 하이퍼링크' };
