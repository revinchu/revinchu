// 이름 있는 셀 스타일의 구성요소. 셀 내용/메모/하이퍼링크는 이 목록에 포함하지 않는다.
export const CELL_STYLE_PARTS = [
  ['number', '표시 형식', ['numFmt', 'code', 'decimals', 'queryFormat']],
  ['alignment', '맞춤', ['align', 'valign', 'indent', 'wrap', 'shrink', 'rotate', 'readingOrder']],
  ['font', '글꼴', ['font', 'size', 'bold', 'italic', 'underline', 'strike', 'color', 'vertAlign']],
  ['border', '테두리', ['bt', 'bb', 'bl', 'br', 'btc', 'bbc', 'blc', 'brc', 'bts', 'bbs', 'bls', 'brs', 'du', 'duc', 'dus', 'dd', 'ddc', 'dds']],
  ['fill', '채우기', ['fill', 'pattern', 'patternColor', 'gradient']],
  ['protection', '보호', ['locked', 'hideFormula']],
];

export function cellStyleKey(name) { return String(name ?? '').trim().toLocaleLowerCase(); }
export function validCellStyleName(name) { return typeof name === 'string' && name.trim().length > 0 && name.trim().length <= 255 && !/[\u0000-\u001f]/.test(name); }
export function cellStyleIncludes(style) { return Object.fromEntries(CELL_STYLE_PARTS.map(([key]) => [key, style?.include?.[key] !== false])); }

export function cellStyleDefaults(base = {}) {
  return {
    numFmt: 'general', code: undefined, decimals: undefined, queryFormat: undefined,
    align: 'general', valign: 'bottom', indent: 0, wrap: false, shrink: false, rotate: 0, readingOrder: 0,
    font: '맑은 고딕', size: 11, bold: false, italic: false, underline: false, strike: false, color: '#000000', vertAlign: undefined,
    bt: false, bb: false, bl: false, br: false, btc: '#000000', bbc: '#000000', blc: '#000000', brc: '#000000',
    bts: 'thin', bbs: 'thin', bls: 'thin', brs: 'thin', du: false, duc: '#000000', dus: 'thin', dd: false, ddc: '#000000', dds: 'thin',
    fill: '#ffffff', pattern: '', patternColor: '#000000', gradient: false, locked: true, hideFormula: false,
    ...base,
  };
}

/** 포함한 요소는 빠진 속성도 초기화하고, 제외한 요소는 기존 셀 서식을 보존한다. */
export function cellStylePatch(def, base = {}) {
  const defaults = cellStyleDefaults(base);
  const include = cellStyleIncludes(def);
  const patch = {};
  for (const [part, , keys] of CELL_STYLE_PARTS) if (include[part]) {
    for (const key of keys) patch[key] = key === 'queryFormat' ? undefined : def.style?.[key] ?? defaults[key];
  }
  patch.cellStyleName = /^(표준|normal)$/i.test(def.name ?? '') ? undefined : def.name;
  return patch;
}

/** 스타일을 수정해도 그 뒤 사용자가 직접 고친 서식은 유지한다. */
export function cellStyleUpdatePatch(current, before, after, base = {}) {
  const old = cellStylePatch(before, base);
  const next = cellStylePatch(after ?? { name: '표준', style: base }, base);
  const defaults = cellStyleDefaults(base);
  const parts = after ? cellStyleIncludes(after) : cellStyleIncludes(before);
  const patch = { cellStyleName: next.cellStyleName };
  const same = (a, b) => JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
  for (const [part, , keys] of CELL_STYLE_PARTS) if (parts[part]) {
    for (const key of keys) if (!(key in old) || same(current[key] ?? defaults[key], old[key])) patch[key] = next[key];
  }
  return patch;
}

/** 외부 파일에서는 스타일 목록만 읽는다. 통합 문서나 셀을 교체하지 않는다. */
export function importCellStyleList(input) {
  if (!Array.isArray(input)) return [];
  const out = []; const seen = new Set();
  for (const item of input) {
    if (!item || typeof item !== 'object' || typeof item.name !== 'string') continue;
    const name = item.name.trim(); const key = cellStyleKey(name);
    if (!validCellStyleName(name) || /^(표준|normal)$/.test(key) || item.builtinId === 0 || seen.has(key)) continue;
    const style = {};
    for (const [, , keys] of CELL_STYLE_PARTS) for (const prop of keys) if (prop !== 'queryFormat' && item.style?.[prop] !== undefined && item.style[prop] !== null) style[prop] = structuredClone(item.style[prop]);
    out.push({ name, style, include: cellStyleIncludes(item), ...(Number.isInteger(item.builtinId) ? { builtinId: item.builtinId } : {}), ...(typeof item.customBuiltin === 'boolean' ? { customBuiltin: item.customBuiltin } : {}), ...(item.hidden ? { hidden: true } : {}) });
    seen.add(key);
  }
  return out;
}
