// 파일 셀의 직접 서식과 피벗 스타일에서 생성된 시각 서식을 구분합니다.
// 표시 형식·맞춤·보호 등은 보존하고, 같은 시각 속성만 상속으로 돌립니다.
const VISUAL_KEYS = new Set(['font', 'size', 'color', 'bold', 'italic', 'underline', 'strike', 'fill', 'pattern', 'patternColor', 'gradient']);
const EDGES = ['bt', 'bb', 'bl', 'br', 'bh', 'bv', 'du', 'dd'];
const same = (a, b) => JSON.stringify(a) === JSON.stringify(b);
const edge = (st, key) => st[key] ? [true, st[key + 's'] ?? 'thin', st[key + 'c'] ?? '#000000'] : [false];

export function capturePivotCellFormat(own, generated = {}, inherited = {}) {
  const out = { ...(generated.numFmt ? { numFmt: 'general' } : {}), ...own };
  const base = { ...inherited, ...generated };
  for (const key of Object.keys(out)) if (VISUAL_KEYS.has(key) && Object.hasOwn(base, key) && same(out[key], base[key])) delete out[key];
  // 선의 색/굵기 중 하나라도 직접 바꿨으면 선 전체를 유지합니다.
  for (const key of EDGES) if (Object.hasOwn(out, key) && Object.hasOwn(base, key) && same(edge(out, key), edge(base, key))) {
    delete out[key]; delete out[key + 's']; delete out[key + 'c'];
  }
  return out;
}
