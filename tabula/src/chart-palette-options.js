// 차트 색 구성의 분류와 사용자 입력 검증. DOM과 통합 문서 상태에 의존하지 않는다.
export const CHART_PALETTE_GROUPS = [
  { id: 'theme', label: '테마' }, { id: 'report', label: '보고서' },
  { id: 'vivid', label: '선명' }, { id: 'pastel', label: '파스텔' },
  { id: 'dark', label: '짙은' }, { id: 'mono', label: '단색' },
  { id: 'contrast', label: '대비' },
];

export const EXTRA_CHART_PALETTES = {
  reportBlue: { label: '보고서 · 푸른 균형', group: 'report', keywords: '청색 남색 청록 실적 매출', colors: ['#285A8E', '#3D8C95', '#7BA3A8', '#D6AA5B', '#A56E78', '#6D718A'] },
  reportEarth: { label: '보고서 · 차분한 대지', group: 'report', keywords: '갈색 녹색 황토 자연', colors: ['#6B705C', '#A58D6F', '#CB997E', '#8B6F78', '#718C8B', '#C4B17A'] },
  reportCool: { label: '보고서 · 맑은 도시', group: 'report', keywords: '파랑 회색 보라 중립', colors: ['#365C79', '#587EA1', '#6E9EB1', '#8E94B2', '#B2A4B8', '#A5B9B0'] },
  vividFestival: { label: '선명 · 다채로운 축제', group: 'vivid', keywords: '빨강 파랑 노랑 보라 다색', colors: ['#2563EB', '#E11D48', '#F59E0B', '#16A34A', '#7C3AED', '#0891B2'] },
  vividCitrus: { label: '선명 · 상큼한 과일', group: 'vivid', keywords: '주황 노랑 초록 라임 따뜻한', colors: ['#F97316', '#EAB308', '#65A30D', '#059669', '#0284C7', '#C026D3'] },
  vividOcean: { label: '선명 · 바다와 산호', group: 'vivid', keywords: '파랑 청록 주황 분홍 해양', colors: ['#0077B6', '#00A6A6', '#F07167', '#FFB703', '#5A4FCF', '#D1495B'] },
  pastelGarden: { label: '파스텔 · 봄 정원', group: 'pastel', keywords: '분홍 연두 연보라 봄', colors: ['#E8A6B1', '#A8CFA1', '#B6AFE0', '#F1CA92', '#9FCED1', '#C9B6A1'] },
  pastelSunset: { label: '파스텔 · 부드러운 노을', group: 'pastel', keywords: '살구 분홍 노랑 보라 따뜻한', colors: ['#E6A57E', '#EAB6BB', '#D5B3D6', '#B5B8DA', '#E3CE99', '#B5CEC3'] },
  pastelSky: { label: '파스텔 · 하늘빛', group: 'pastel', keywords: '연파랑 민트 청록 차가운', colors: ['#90B7D9', '#9DCAD4', '#AAD9C8', '#B8B1D7', '#D3B8CB', '#C9D2A3'] },
  darkJewel: { label: '짙은 · 보석빛', group: 'dark', keywords: '남색 보라 자주 녹색', colors: ['#243B6B', '#5B2C6F', '#8E244D', '#17645B', '#84611D', '#3D5363'] },
  darkForest: { label: '짙은 · 깊은 숲', group: 'dark', keywords: '초록 청록 갈색 자연', colors: ['#164A41', '#315D43', '#50684C', '#49636E', '#745344', '#645B7A'] },
  darkUrban: { label: '짙은 · 도시의 밤', group: 'dark', keywords: '남색 회색 자주 청록', colors: ['#25344A', '#4F4C70', '#76465A', '#38626A', '#71634D', '#5E6874'] },
  monoOcean: { label: '단색 · 푸른 바다', group: 'mono', mono: true, keywords: '파랑 청색 순차', colors: ['#153B66', '#22598B', '#3478AD', '#5598C4', '#82B7D7', '#B6D7E8'] },
  monoForest: { label: '단색 · 초록 잎', group: 'mono', mono: true, keywords: '녹색 초록 순차', colors: ['#164B3B', '#256A4F', '#3B8B66', '#64A782', '#96C1A8', '#C8DECE'] },
  monoPlum: { label: '단색 · 자두빛', group: 'mono', mono: true, keywords: '보라 자주 순차', colors: ['#4D285F', '#6B3A7C', '#895998', '#AA80B4', '#C6A7CD', '#E0D0E3'] },
  contrastBlueOrange: { label: '대비 · 파랑과 주황', group: 'contrast', keywords: '청색 주황 반대 비교', colors: ['#174A7E', '#D96016', '#4C87B9', '#F2A15F', '#7AADC9', '#8E3B10'] },
  contrastPurpleGreen: { label: '대비 · 보라와 초록', group: 'contrast', keywords: '보라 녹색 반대 비교', colors: ['#5B2C83', '#238443', '#9870B5', '#71B978', '#C4A7D3', '#155D33'] },
  contrastInkGold: { label: '대비 · 먹색과 금빛', group: 'contrast', keywords: '검정 노랑 회색 대비', colors: ['#252A34', '#CA8A04', '#566174', '#F2BE45', '#8C99AD', '#815B12'] },
};
for (const palette of Object.values(EXTRA_CHART_PALETTES)) palette.wixel = true;

/** #RGB 또는 #RRGGBB를 표준 HEX로 바꾼다. CSS 함수와 이름 색은 받지 않는다. */
export function normalizeChartPaletteColor(value) {
  if (typeof value !== 'string') return null;
  const color = value.trim();
  if (/^#[0-9a-f]{6}$/i.test(color)) return color.toUpperCase();
  if (/^#[0-9a-f]{3}$/i.test(color)) return '#' + [...color.slice(1)].map(c => c + c).join('').toUpperCase();
  return null;
}

/** 사용자 입력은 전체를 검증한다. 일부를 버리거나 순서·중복색을 바꾸지 않는다. */
export function normalizeChartPalette(input, { maxColors = 32 } = {}) {
  const values = Array.isArray(input) ? input : typeof input === 'string' ? input.trim().split(/[\s,;]+/).filter(Boolean) : [];
  const max = maxColors === Infinity || Number.isInteger(maxColors) && maxColors > 0 ? maxColors : 32;
  if (!values.length || values.length > max) return { colors: null, error: max === Infinity ? '색을 1개 이상 입력하세요.' : `색을 1~${max}개 입력하세요.` };
  const colors = [];
  for (let i = 0; i < values.length; i++) {
    const color = normalizeChartPaletteColor(values[i]);
    if (!color) return { colors: null, error: `${i + 1}번째 색을 #RGB 또는 #RRGGBB 형식으로 입력하세요.` };
    colors.push(color);
  }
  return { colors, error: null };
}

/** 호출할 때 테마 getter를 평가하므로 테마를 바꾼 뒤에도 미리보기 색이 갱신된다. */
export function chartPaletteOptions(palettes, query = '', group = 'all') {
  const words = String(query ?? '').trim().toLowerCase().split(/\s+/).filter(Boolean), result = [];
  for (const [id, palette] of Object.entries(palettes ?? {})) {
    const category = palette.group ?? (palette.mono ? 'mono' : palette.wixel ? 'report' : 'theme');
    if (group !== 'all' && group !== category) continue;
    const colors = [...palette.colors], groupLabel = CHART_PALETTE_GROUPS.find(g => g.id === category)?.label ?? category;
    const text = `${id} ${palette.label} ${groupLabel} ${palette.keywords ?? ''} ${colors.join(' ')}`.toLowerCase();
    if (!words.every(word => text.includes(word))) continue;
    result.push({ id, label: palette.label, group: category, colors, mono: !!palette.mono, wixel: !!palette.wixel });
  }
  return result;
}
