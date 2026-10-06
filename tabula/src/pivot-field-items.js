import { keyOf, itemText, itemIdentity, itemProperty, sortKeys, groupKey, EMPTY } from './cube.js';
import { pivotSourceData } from './pivot.js';
import { formatValue, isDateCode } from './format.js';

const fold = value => String(value ?? '').toLowerCase();
const fieldEntry = (object, field) => Object.entries(object ?? {}).find(([name]) => fold(name) === fold(field));
const sameSet = (a, b) => a.size === b.size && [...a].every(value => b.has(value));

/** UI identity stays typed; persisted filters keep the existing itemText semantics. */
export function pivotFieldSourceIndex(source, def, field) {
  const header = source?.cube?.header ?? [], direct = header.findIndex(name => fold(name) === fold(field));
  if (direct >= 0) return direct;
  const spec = fieldEntry(def.groups, field)?.[1];
  return spec?.base ? header.findIndex(name => fold(name) === fold(spec.base)) : -1;
}

/** Current records, retained cache items and explicit selection form one UI universe.
 * Missing-data visibility never edits the retained cache or the filter definition. */
export function pivotFieldItemModel(wb, def, field, options = {}) {
  const source = options.source ?? pivotSourceData(wb, def, { preserveSnapshot: !!options.preserveSnapshot }), index = pivotFieldSourceIndex(source, def, field);
  const spec0 = fieldEntry(def.groups, field)?.[1], spec = spec0 ? { ...spec0, date1904: !!wb.date1904 } : null;
  const cache = wb.pivotCacheItems?.[def.cacheItemsId], exact = cache?.fields?.find(f => fold(f.name) === fold(field));
  const fallback = !exact && spec?.base ? cache?.fields?.find(f => fold(f.name) === fold(spec.base)) : null;
  const retained = exact ?? fallback;
  const own = fieldEntry(def.filters, field), original = own?.[1] ?? null;
  const allowed = original ? new Set(original.map(itemIdentity)) : null;
  const dateFormat = retained?.format;
  const sourceStyle = source?.ref && index >= 0 ? wb.styleAt(source.si, Math.min(source.ref.r1 + 1, source.ref.r2), source.ref.c1 + index) : null;
  const items = new Map();
  const project = value => spec && value !== EMPTY ? groupKey(value, spec) : value;
  const add = (value, { current = false, hasData = false, date = false, grouped = false, unknown = false, filterValue } = {}) => {
    const v = grouped ? value : project(value), id = itemIdentity(v), key = filterValue ?? itemText(v);
    let entry = items.get(id);
    if (!entry) {
      entry = { id, key, v, hasData: !!hasData, deleted: !current, unknown, date: !!date };
      items.set(id, entry);
    } else { entry.hasData ||= !!hasData; entry.deleted &&= !current; entry.unknown &&= unknown; entry.date ||= !!date; }
    return entry;
  };
  const current = options.currentItems ?? (index < 0 ? [] : source.cube.col(index).dim().keys.map(v => ({ v, hasData: true })));
  const dateOnly = retained?.dateOnly || retained?.sharedTypes?.includes('d') && !/[nsbe]/.test(retained.sharedTypes);
  const dates = new Set();
  if (!dateOnly && retained?.sharedTypes) for (let i = 0; i < retained.sharedTypes.length; i++) {
    if (retained.sharedTypes[i] === 'd') dates.add(retained.shared[i]);
  }
  for (const entry of current) add(entry.v, { current: true, hasData: entry.hasData, date: typeof entry.v === 'number' && (dateOnly || dates.has(entry.v)) });
  // A loaded none policy keeps its saved cache until refresh; selected values are
  // retained below even if an explicit policy change removes unselected history.
  const snapshot = def.snapshotId && wb.pivotSnapshots?.get(def.snapshotId);
  const checking = snapshot && options.preserveSnapshot ? { ...snapshot } : snapshot;
  const keepHistory = def.missingItems !== 'none' || cache?.missingItemsLimit === 0 && snapshot && (!wb.pivotSnapshotCurrent || wb.pivotSnapshotCurrent(checking, def));
  if (keepHistory && retained) for (let i = 0; i < (retained.shared?.length ?? 0); i++) {
    // A derived cache field already contains group labels. Never add its base
    // serials to that field, or apply a date group to an already grouped year.
    const grouped = !!spec?.base && !!exact && fold(exact.name) !== fold(spec.base);
    add(keyOf(retained.shared[i]), { date: retained.sharedTypes?.[i] === 'd', grouped });
  }
  const knownSelections = new Set([...items.values()].map(entry => itemIdentity(entry.key)));
  for (const value of original ?? []) if (!knownSelections.has(itemIdentity(value))) {
    add(keyOf(value), { grouped: true, unknown: true, filterValue: value });
    knownSelections.add(itemIdentity(value));
  }
  const ordered = sortKeys([...items.values()].map(entry => entry.v)).map(value => items.get(itemIdentity(value)));
  const captions = fieldEntry(def.itemCaptions, field)?.[1];
  for (const entry of ordered) {
    const sourceDate = ['date', 'shortDate', 'longDate', 'time'].includes(sourceStyle?.numFmt) || sourceStyle?.numFmt === 'custom' && isDateCode(sourceStyle.code ?? '');
    const st = spec ? null : entry.date ? dateFormat?.numFmt && dateFormat.numFmt !== 'general' ? dateFormat : sourceDate ? sourceStyle : { numFmt: 'date' } : sourceStyle;
    entry.text = entry.key === '' ? '(비어 있음)' : typeof entry.v === 'number' && st?.numFmt && st.numFmt !== 'general'
      ? formatValue(entry.v, st, wb.date1904).text : String(entry.key);
    entry.text = itemProperty(captions, entry.key) ?? entry.text;
    entry.selected = !allowed || allowed.has(itemIdentity(entry.key));
  }
  return { items: ordered, index, field: own?.[0] ?? field, original, filtered: !!original,
    selected: new Set(ordered.filter(entry => entry.selected).map(entry => entry.id)) };
}

/** Opening a menu and confirming preserves the exact original selection array. */
export function pivotItemSelection(model, chosen) {
  if (chosen === null) return { unchanged: model.original === null, values: null };
  const selected = new Set(chosen);
  if (sameSet(selected, model.selected)) return { unchanged: true, values: model.original };
  const all = model.items.every(entry => selected.has(entry.id));
  if (all && !model.items.some(entry => entry.unknown)) return { unchanged: model.original === null, values: null };
  const values = [], seen = new Set();
  for (const entry of model.items) if (selected.has(entry.id)) {
    const id = itemIdentity(entry.key); if (!seen.has(id)) { seen.add(id); values.push(entry.key); }
  }
  return { unchanged: false, values };
}

/** Ctrl/additive selection includes hidden entries; a direct click replaces it. */
export function slicerPickValues(model, key, additive) {
  const universe = model.allItems ?? model.items, all = universe.map(item => item.key);
  let chosen;
  if (additive) {
    chosen = new Set(universe.filter(item => item.selected).map(item => item.key));
    if (chosen.has(key)) chosen.delete(key); else chosen.add(key);
    if (!chosen.size) chosen = new Set(all);
  } else {
    chosen = new Set([key]);
    if (model.filtered && universe.filter(item => item.selected).length === 1 && universe.find(item => item.key === key)?.selected) chosen = new Set(all);
  }
  // Pivot model performs its own full-universe/no-change decision. Other slicers
  // keep their existing null-means-all contract.
  return chosen.size === all.length && !model.preserveSelection ? null : all.filter(value => chosen.has(value));
}
