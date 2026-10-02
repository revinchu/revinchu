// Conditional rules preserve banding in Excel without changing cell values or direct formats.
export const BANDING_PALETTES = [
  ['포레스트', '#185c45', '#ffffff', '#e8f3ed', '#c9e2d3'], ['오션', '#2456a6', '#ffffff', '#edf3fc', '#cddcf4'],
  ['라벤더', '#65448d', '#ffffff', '#f2edf8', '#dfd1ef'], ['코랄', '#994530', '#ffffff', '#fcf0e9', '#f1d0bf'],
  ['슬레이트', '#354657', '#ffffff', '#edf0f4', '#d4dce5'], ['샌드', '#785f2a', '#ffffff', '#faf5e7', '#ece0be'],
  ['민트', '#246c69', '#ffffff', '#e9f6f4', '#c5e5df'], ['로즈', '#923f63', '#ffffff', '#f9edf2', '#eccddb'],
];
export function isBandingRule(rule) { return rule?.type === 'formula' && rule.formula?.includes('N("WIXEL_BANDING")'); }
export function alternatingRules(area, { colors = BANDING_PALETTES[0].slice(1), header = true, footer = false } = {}) {
  if (!['r1','r2','c1','c2'].every(k => Number.isInteger(area[k]) && area[k] >= 0) || area.r2 < area.r1 || area.c2 < area.c1) throw new Error('교차색상 범위를 확인하세요.');
  if (colors.length !== 4 || colors.some(c => !/^#[0-9a-f]{6}$/i.test(c))) throw new Error('교차색상에 올바른 색을 지정하세요.');
  const rules = [], rule = (test, fill, extra = {}) => ({ ...area, type: 'formula', formula: `=AND(N("WIXEL_BANDING")=0,${test})`, style: { fill, ...extra }, stopIfTrue: false });
  if (header) rules.push(rule(`ROW()=${area.r1 + 1}`, colors[0], { color: '#ffffff', bold: true }));
  if (footer && area.r2 > area.r1) rules.push(rule(`ROW()=${area.r2 + 1}`, colors[3], { bold: true }));
  const start = area.r1 + (header ? 1 : 0), end = area.r2 - (footer && area.r2 > area.r1 ? 1 : 0);
  if (end >= start) for (let parity = 0; parity < 2; parity++) rules.push(rule(`ROW()>=${start + 1},ROW()<=${end + 1},MOD(ROW()-${start + 1},2)=${parity}`, colors[parity + 1]));
  return rules;
}
