// 외부 또는 누락된 원본의 가져온 캐시만 원래 연결과 함께 사용합니다.
const binding = def => ({
  source: def.source == null ? null : String(def.source).toLowerCase(),
  table: def.table == null ? null : String(def.table).toLowerCase(),
  range: def.range ? { r1: def.range.r1, c1: def.range.c1, r2: def.range.r2, c2: def.range.c2 } : null,
});
export function makeImportedPivotSourceReference(source, def) {
  return { kind: 'imported-pivot-cache', ref: source.ref ?? null, sheet: source.sheet ?? null, name: source.name ?? null,
    ...(source.external ? { external: source.external } : {}), expectedCacheItemsId: def.cacheItemsId ?? null, binding: binding(def) };
}
export function importedPivotSourceReference(def) {
  const reference = def?.sourceReference;
  return reference?.kind === 'imported-pivot-cache' && reference.binding ? reference : null;
}
export function pivotSourceReferenceCurrent(def) {
  const reference = importedPivotSourceReference(def);
  if (!reference || !reference.expectedCacheItemsId || reference.expectedCacheItemsId !== def.cacheItemsId
    || def.snapshotId != null && def.snapshotId !== reference.expectedCacheItemsId) return false;
  return pivotSourceBindingCurrent(def);
}
export function pivotSourceBindingCurrent(def) {
  const reference = importedPivotSourceReference(def);
  if (!reference) return false;
  const saved = reference.binding, current = binding(def);
  return saved.source === current.source && saved.table === current.table
    && (saved.range === null ? current.range === null : current.range !== null
      && ['r1', 'c1', 'r2', 'c2'].every(key => saved.range?.[key] === current.range[key]));
}

const layoutKeys = ['rows', 'cols', 'values', 'pages'];
const queryKeys = ['filters', 'pageMulti', 'pageOrder', 'pageWrap', 'calcFields', 'calcItems', 'sort', 'order', 'groups', 'fieldFilters', 'collapsed', 'layout', 'subtotals', 'grandRows', 'grandCols', 'valuesOnRows', 'valuesPos', 'showValuesRow', 'valuesHeadRow', 'showHeaders', 'customListSort', 'mergeLabels', 'classic', 'blankRows', 'subtotalTop', 'repeatLabels', 'missingItems'];
// 항목 배열을 복제하지 않고 records 없는 배치의 변경 여부를 기록합니다.
export function pivotCacheLayoutBinding(def) {
  let a = 2166136261, b = 2654435769;
  const feed = text => { for (let i = 0; i < text.length; i++) { const c = text.charCodeAt(i); a = Math.imul(a ^ c, 16777619); b = Math.imul(b ^ c, 2246822507); } };
  const walk = value => {
    if (value === null || value === undefined) { feed(value === null ? 'n;' : 'u;'); return; }
    if (Array.isArray(value)) { feed('a' + value.length + ':'); for (const item of value) walk(item); return; }
    if (typeof value === 'object') { const keys = Object.keys(value).filter(key => value[key] !== undefined).sort(); feed('o' + keys.length + ':'); for (const key of keys) { walk(key); walk(value[key]); } return; }
    const text = String(value); feed(typeof value + text.length + ':' + text + ';');
  };
  for (const key of queryKeys) { walk(key); walk(def[key]); }
  return { fields: JSON.stringify(Object.fromEntries(layoutKeys.map(key => [key, def[key] ?? []]))), query: (a >>> 0).toString(16).padStart(8, '0') + (b >>> 0).toString(16).padStart(8, '0') };
}
export function pivotCacheLayoutCurrent(def, layout) {
  const current = pivotCacheLayoutBinding(def);
  return layout?.binding?.fields === current.fields && layout.binding.query === current.query;
}
