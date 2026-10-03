// 캐시의 정렬·데이터 없는 항목 옵션은 연결된 모든 슬라이서가 공유합니다.
const CACHE_OPTIONS = ['sort', 'customList', 'hideNoData', 'markNoData', 'noDataLast', 'showDeleted'];
const fold = value => String(value ?? '').toLowerCase();

/** resolveTargets는 앱이 사용하는 피벗 참조 해석기입니다. 셀이나 원본 데이터를 계산하지 않습니다. */
export function slicerSettingsKey(source, hostSi, resolveTargets) {
  if (!source) return null;
  if (source.kind === 'cache') return source.cacheKey && source.field
    ? JSON.stringify(['cache', source.cacheKey, fold(source.field)]) : null;
  if (source.kind === 'table') return source.table && source.column
    ? JSON.stringify(['table', fold(source.table), fold(source.column)]) : null;
  if ((source.kind && source.kind !== 'pivot') || !source.field || !resolveTargets) return null;
  const links = new Map();
  for (const entry of resolveTargets(source, hostSi)) {
    const name = entry.def.name ?? `피벗 테이블${entry.index + 2}`;
    const link = [entry.si, fold(name)];
    links.set(JSON.stringify(link), link);
  }
  if (!links.size) return null;
  const sorted = [...links.values()].sort((a, b) => a[0] - b[0] || a[1].localeCompare(b[1]));
  return JSON.stringify(['pivot', fold(source.field), sorted]);
}

/** 선택 개체의 모양은 독립적으로 바꾸고 캐시 옵션만 공유합니다. 전체 변경은 한 번에 실행 취소합니다. */
export function applySlicerSettings(wb, hostSi, id, patch, resolveTargets, metadata) {
  const selected = wb.sheets[hostSi]?.slicers?.find(sl => sl.id === id);
  if (!selected) return false;
  const options = Object.fromEntries(CACHE_OPTIONS.filter(key => Object.hasOwn(patch, key)).map(key => [key, patch[key]]));
  const sharedKey = Object.keys(options).length ? slicerSettingsKey(selected.source, hostSi, resolveTargets) : null;
  let changed = false;
  wb.transact(() => {
    wb.sheets.forEach((sheet, si) => {
      let sheetChanged = false;
      const slicers = (sheet.slicers ?? []).map(sl => {
        const own = si === hostSi && sl === selected;
        const next = own ? patch : sharedKey && slicerSettingsKey(sl.source, si, resolveTargets) === sharedKey ? options : null;
        if (!next || !Object.keys(next).some(key => sl[key] !== next[key])) return sl;
        sheetChanged = true;
        return { ...sl, ...next };
      });
      if (sheetChanged) { wb.setSheetProp(si, 'slicers', slicers); changed = true; }
    });
  }, metadata);
  return changed;
}
