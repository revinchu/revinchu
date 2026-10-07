import { filterButtonVisible } from './filter-display.js';
import { tableAt } from './tables.js';

// Match model cells rather than virtualized buttons. Permission and menu focus
// remain with the caller, as they do for the existing mouse menu commands.
export function cellDropdownTarget(sheet, r, c) {
  if (!sheet || !Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) return null;
  const table = tableAt(sheet, r, c);
  const tableHeader = table?.filter && table.header && r === table.r1;
  if (tableHeader) {
    if (filterButtonVisible(table.filter, c)) return { kind: 'filter', column: c, key: table.id };
  } else {
    const filter = sheet.filter;
    if (filter && r === filter.r1 && c >= filter.c1 && c <= filter.c2 && filterButtonVisible(filter, c)) {
      return { kind: 'filter', column: c, key: '' };
    }
  }
  // This index is the same compact primary/extra order used by the rendered
  // buttons and app.pivotDefs(), including workbooks without a primary pivot.
  const pivots = [sheet.pivot, ...(sheet.pivotsExtra ?? [])].filter(Boolean);
  for (let pivotIndex = 0; pivotIndex < pivots.length; pivotIndex++) {
    const def = pivots[pivotIndex];
    for (const button of def.buttons ?? []) {
      if (button?.r !== r || button.c !== c || !['rows', 'cols', 'page'].includes(button.kind)) continue;
      if (button.kind !== 'page' && (def.showHeaders === false || def.fieldCaptions === false)) continue;
      const field = button.sigma ? undefined : button.field || undefined;
      const fields = button.kind === 'rows' ? def.rows ?? [] : button.kind === 'cols' ? def.cols ?? [] : [];
      // A sigma button opens the real axis-field choices, just like its mouse
      // menu. A value-only axis and an unbound page button have no choices.
      if (!field && !fields.length) continue;
      return { kind: 'pivot', pivotIndex, buttonKind: button.kind, field };
    }
  }
  return null;
}

// An idle IME can report a physical ArrowDown as Process/229. Composition and
// AltGr stay with text entry; without a physical code, use only an ordinary key.
export function altArrowDownKey(event) {
  if (!event || !event.altKey || event.ctrlKey || event.metaKey || event.isComposing || event.getModifierState?.('AltGraph')) return false;
  if (event.code && event.code !== 'Unidentified') return event.code === 'ArrowDown';
  return event.key === 'ArrowDown' && event.keyCode !== 229;
}
