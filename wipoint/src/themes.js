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
    else if (k === 'alphaMod') alpha *= v;
    else if (k === 'alphaOff') alpha = Math.max(0, Math.min(1, alpha + v));
    else if (k === 'satMod' || k === 'satOff' || k === 'hueMod' || k === 'hueOff' || k === 'comp') {
      // PowerPoint 와 같이 HSL 에서 (hueOff 는 도)
      let [h, s2, l] = rgbToHsl(r, g, b);
      if (k === 'satMod') s2 *= v;
      else if (k === 'satOff') s2 += v;
      else if (k === 'hueMod') h = (h * v) % 1;
      else if (k === 'hueOff') h = (((h + v / 360) % 1) + 1) % 1;
      else h = (h + 0.5) % 1;
      [r, g, b] = hslToRgb(h, Math.max(0, Math.min(1, s2)), l);
    } else if (k === 'inv') [r, g, b] = [255 - r, 255 - g, 255 - b];
    else if (k === 'gray') { const y = Math.round(0.3 * r + 0.59 * g + 0.11 * b); r = g = b = y; }
    else if (k === 'red' || k === 'green' || k === 'blue' || k === 'redMod' || k === 'greenMod' || k === 'blueMod' || k === 'redOff' || k === 'greenOff' || k === 'blueOff') {
      const ch = k.startsWith('red') ? 0 : k.startsWith('green') ? 1 : 2;
      const arr = [r, g, b];
      const lin = toLinear(arr[ch]);
      const nv = k.endsWith('Mod') ? lin * v : k.endsWith('Off') ? lin + v : v;
      arr[ch] = fromLinear(Math.max(0, Math.min(1, nv)));
      [r, g, b] = arr;
    }
  }
  return { hex: rgbToHex(r, g, b), alpha };
}

const MOD_CODES = { lm: 'lumMod', lo: 'lumOff', t: 'tint', s: 'shade', a: 'alpha', sm: 'satMod', so: 'satOff', hm: 'hueMod', ho: 'hueOff', am: 'alphaMod', ao: 'alphaOff', cp: 'comp', iv: 'inv', gr: 'gray', rd: 'red', gn: 'green', bl: 'blue', rm: 'redMod', gm: 'greenMod', bm: 'blueMod', ro: 'redOff', go: 'greenOff', bo: 'blueOff' };
const MOD_SHORT = Object.fromEntries(Object.entries(MOD_CODES).map(([a, b]) => [b, a]));

/** '@accent1:lm75:lo25' → { slot, mods } */
export function parseColorRef(c) {
  if (typeof c !== 'string') return null;
  const [base, ...rest] = c.split(':');
  // ho (hueOff) 는 도 단위 그대로, 나머지는 %
  const mods = rest.map((m) => { const x = /^([a-z]+)(-?[\d.]+)$/.exec(m); return x && MOD_CODES[x[1]] ? [MOD_CODES[x[1]], x[1] === 'ho' ? Number(x[2]) : Number(x[2]) / 100] : null; }).filter(Boolean);
  if (base.startsWith('@')) return { slot: ALIAS[base.slice(1)] ?? base.slice(1), mods };
  return { hex: base, mods };
}
export function colorRef(slotOrHex, mods = []) {
  const head = slotOrHex.startsWith('#') ? slotOrHex.toUpperCase() : `@${slotOrHex}`;
  return [head, ...mods.map(([k, v]) => `${MOD_SHORT[k]}${k === 'hueOff' ? Math.round(v * 100) / 100 : Math.round(v * 1000) / 10}`)].join(':');
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

/** 색 구성 (디자인 › 적용 › 색): [이름, SLOT_NAMES 순서 12색] — Office 색 구성과 같은 계열 + WIPOINT 추가 */
export const COLOR_SCHEMES = [
  ['Office', ['#000000', '#FFFFFF', '#0E2841', '#E8E8E8', '#156082', '#E97132', '#196B24', '#0F9ED5', '#A02B93', '#4EA72E', '#467886', '#96607D']],
  ['Office 2013 - 2022', ['#000000', '#FFFFFF', '#44546A', '#E7E6E6', '#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#0563C1', '#954F72']],
  ['Office 2007 - 2010', ['#000000', '#FFFFFF', '#1F497D', '#EEECE1', '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', '#F79646', '#0000FF', '#800080']],
  ['회색조', ['#000000', '#FFFFFF', '#000000', '#F8F8F8', '#DDDDDD', '#B2B2B2', '#969696', '#808080', '#5F5F5F', '#4D4D4D', '#5F5F5F', '#919191']],
  ['파랑 따뜻하게', ['#000000', '#FFFFFF', '#242852', '#ACCBF9', '#4A66AC', '#629DD1', '#297FD5', '#7F8FA9', '#5AA2AE', '#9D90A0', '#9454C3', '#3EBBF0']],
  ['파랑', ['#000000', '#FFFFFF', '#17406D', '#DBEFF9', '#0F6FC6', '#009DD9', '#0BD0D9', '#10CF9B', '#7CCA62', '#A5C249', '#F49100', '#85DFD0']],
  ['파랑 II', ['#000000', '#FFFFFF', '#335B74', '#DFE3E5', '#1CADE4', '#2683C6', '#27CED7', '#42BA97', '#3E8853', '#62A39F', '#6B9F25', '#B26B02']],
  ['파랑 녹색', ['#000000', '#FFFFFF', '#373545', '#CEDBE6', '#3494BA', '#58B6C0', '#75BDA7', '#7A8C8E', '#84ACB6', '#2683C6', '#6B9F25', '#9F6715']],
  ['녹색', ['#000000', '#FFFFFF', '#455F51', '#E3DED1', '#549E39', '#8AB833', '#C0CF3A', '#029676', '#4AB5C4', '#0989B1', '#6B9F25', '#BA6906']],
  ['녹색 노랑', ['#000000', '#FFFFFF', '#455F51', '#E3DED1', '#99CB38', '#63A537', '#37A76F', '#44C1A3', '#4EB3CF', '#51C3F9', '#EE7B08', '#977B2D']],
  ['노랑', ['#000000', '#FFFFFF', '#39302A', '#E5DEDB', '#FFCA08', '#F8931D', '#CE8D3E', '#EC7016', '#E64823', '#9C6A6A', '#2998E3', '#7F723D']],
  ['노랑 주황', ['#000000', '#FFFFFF', '#4E3B30', '#FBEEC9', '#F0A22E', '#A5644E', '#B58B80', '#C3986D', '#A19574', '#C17529', '#AD1F1F', '#FFC42F']],
  ['주황', ['#000000', '#FFFFFF', '#637052', '#CCDDEA', '#E48312', '#BD582C', '#865640', '#9B8357', '#C2BC80', '#94A088', '#2998E3', '#8C8C8C']],
  ['주황 빨강', ['#000000', '#FFFFFF', '#696464', '#E9E5DC', '#D34817', '#9B2D1F', '#A28E6A', '#956251', '#918485', '#855D5D', '#CC9900', '#96A9A9']],
  ['빨강 주황', ['#000000', '#FFFFFF', '#505046', '#EEECE1', '#E84C22', '#FFBD47', '#B64926', '#FF8427', '#CC9900', '#B22600', '#CC9900', '#666699']],
  ['빨강', ['#000000', '#FFFFFF', '#323232', '#E5C243', '#A5300F', '#D55816', '#E19825', '#B19C7D', '#7F5F52', '#B27D49', '#6B9F25', '#B26B02']],
  ['빨강 보라', ['#000000', '#FFFFFF', '#454551', '#D8D9DC', '#E32D91', '#C830CC', '#4EA6DC', '#4775E7', '#8971E1', '#D54773', '#6B9F25', '#8C8C8C']],
  ['보라', ['#000000', '#FFFFFF', '#373545', '#DCD8DC', '#AD84C6', '#8784C7', '#5D739A', '#6997AF', '#84ACB6', '#6F8183', '#69A020', '#8C8C8C']],
  ['보라 II', ['#000000', '#FFFFFF', '#632E62', '#EAE5EB', '#92278F', '#9B57D3', '#755DD9', '#665EB8', '#45A5ED', '#5982DB', '#0066FF', '#666699']],
  ['중앙값', ['#000000', '#FFFFFF', '#775F55', '#EBDDC3', '#94B6D2', '#DD8047', '#A5AB81', '#D8B25C', '#7BA79D', '#968C8C', '#F7B615', '#704404']],
  ['종이', ['#000000', '#FFFFFF', '#444D26', '#FEFAC9', '#A5B592', '#F3A447', '#E7BC29', '#D092A7', '#9C85C0', '#809EC2', '#8E58B6', '#7F6F6F']],
  ['마키', ['#000000', '#FFFFFF', '#5E5E5E', '#DDDDDD', '#418AB3', '#A6B727', '#F69200', '#838383', '#FEC306', '#DF5327', '#F59E00', '#B2B2B2']],
  ['슬립스트림', ['#000000', '#FFFFFF', '#212745', '#B4DCFA', '#4E67C8', '#5ECCF3', '#A7EA52', '#5DCEAF', '#FF8021', '#F14124', '#56C7AA', '#59A8D1']],
  ['기류', ['#000000', '#FFFFFF', '#323232', '#E3DED1', '#F07F09', '#9F2936', '#1B587C', '#4E8542', '#604878', '#C19859', '#6B9F25', '#B26B02']],
  ['패싯', ['#000000', '#FFFFFF', '#2C3C43', '#EBEBEB', '#90C226', '#54A021', '#E6B91E', '#E76618', '#C42F1A', '#918655', '#99CA3C', '#B9D181']],
  ['이온', ['#000000', '#FFFFFF', '#1E5155', '#EBEBEB', '#B01513', '#EA6312', '#E6B729', '#6AAC90', '#5F9C9D', '#9E5E9B', '#58C1BA', '#9DFFCB']],
  ['슬라이스', ['#000000', '#FFFFFF', '#146194', '#DEF5FA', '#052F61', '#A50E82', '#14967C', '#6A9E1F', '#E87D37', '#C62324', '#0D2E46', '#356A95']],
  ['네이버', ['#000000', '#FFFFFF', '#03C75A', '#EAFBF1', '#03C75A', '#1EC800', '#00B2FF', '#FF6B00', '#7C4DFF', '#FFC300', '#03C75A', '#7C4DFF']],
  ['파스텔', ['#3B3B3B', '#FFFFFF', '#5B6C8F', '#F4F6FB', '#8EC5FC', '#E0C3FC', '#F9C6C9', '#FDE2A7', '#B9F3D0', '#A0E7E5', '#5B6C8F', '#9B7EDE']],
  ['비비드', ['#111111', '#FFFFFF', '#1A1A2E', '#F5F5F5', '#FF3366', '#FFB400', '#00C2A8', '#3D5AFE', '#AA00FF', '#FF6D00', '#3D5AFE', '#AA00FF']],
  ['어스 톤', ['#2B2118', '#FFFFFF', '#5C4033', '#F3EEE7', '#A0522D', '#CD853F', '#8B7D3A', '#6B8E23', '#556B2F', '#B8860B', '#8B4513', '#6B8E23']],
  ['모노 블루', ['#0B1F33', '#FFFFFF', '#0B3C5D', '#EEF4FA', '#0B3C5D', '#328CC1', '#6FB1E8', '#A9D3F5', '#1D5C8A', '#4F7CAC', '#328CC1', '#6FB1E8']],
];
export const schemeColors = (list) => Object.fromEntries(SLOT_NAMES.map((k, i) => [k, list[i]]));
