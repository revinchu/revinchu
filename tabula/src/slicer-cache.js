// 연결된 피벗이 없는 Excel 슬라이서도 저장된 항목과 선택을 유지한다.
import { formatValue } from './format.js';
export function cachedSlicerItems(slicer, date1904 = false) {
  const source = slicer.source ?? {}, values = source.values ?? [];
  const selected = Array.isArray(slicer.cacheSelection) ? new Set(slicer.cacheSelection) : null;
  return (source.items ?? values.map((_, index) => ({ index, hasData: true })))
    .filter(it => Number.isInteger(it.index) && it.index >= 0 && it.index < values.length)
    .map(it => {
      const v = values[it.index], key = String(it.index);
      const text = v == null ? '(비어 있음)' : v?.error ?? (typeof v === 'boolean' ? v ? 'TRUE' : 'FALSE' : typeof v === 'number' ? formatValue(v, source.format ?? {}, date1904).text : String(v));
      return { key, v, text, hasData: it.hasData !== false, selected: !selected || selected.has(key) };
    });
}
export function applyCachedSlicerSelection(wb, source, values, metadata) {
  const valid = new Set((source.items ?? (source.values ?? []).map((_, index) => ({ index }))).map(it => String(it.index)));
  const selection = values === null ? null : [...new Set(values)].filter(key => valid.has(key));
  const cacheSelection = selection?.length === valid.size ? null : selection;
  wb.transact(() => {
    wb.sheets.forEach((sheet, si) => {
      let changed = false;
      const slicers = (sheet.slicers ?? []).map(sl => {
        if (sl.source?.kind !== 'cache' || sl.source.cacheKey !== source.cacheKey || sl.source.field !== source.field) return sl;
        changed = true;
        return { ...sl, cacheSelection };
      });
      if (changed) wb.setSheetProp(si, 'slicers', slicers);
    });
  }, metadata);
}
