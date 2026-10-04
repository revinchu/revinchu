import { EMPTY, groupKey, itemIdentity, itemText } from './cube.js';
import { pivotFieldSourceIndex } from './pivot-field-items.js';

const fold = value => String(value ?? '').toLowerCase();
const otherFilters = (def, field) => Object.entries(def.filters ?? {}).filter(([name]) => fold(name) !== fold(field));
const dateSystem = (source, def) => !!(source?.date1904 ?? def.date1904);

/** Availability excludes only this slicer field. Group labels are expanded to
 * the original column values so itemStats keeps its column/rollup fast path. */
export function slicerPivotFilters(source, def, field) {
  const cube = source?.cube;
  if (!cube) return [];
  const groups = new Map(Object.entries(def.groups ?? {}).map(([name, spec]) => [fold(name), spec]));
  const result = [];
  for (const [name, allowed] of otherFilters(def, field)) {
    const index = pivotFieldSourceIndex(source, def, name);
    if (index < 0) continue;
    const spec = groups.get(fold(name));
    if (!spec) { result.push([index, new Set(allowed)]); continue; }
    const selected = new Set(allowed.map(itemIdentity));
    const group = { ...spec, date1904: dateSystem(source, def) }, raw = new Set();
    // Distinct dimension keys only: never materialize source rows or a new cube.
    for (const value of cube.col(index).dim().keys) {
      const label = itemText(value === EMPTY ? EMPTY : groupKey(value, group));
      if (selected.has(itemIdentity(label))) raw.add(itemText(value));
    }
    result.push([index, raw]);
  }
  return result;
}

/** A change to another group definition/date system changes availability even
 * when the saved selection strings stay the same. Own selections do not. */
export function slicerPivotFilterKey(source, def, field) {
  return JSON.stringify([fold(field), otherFilters(def, field), def.groups ?? {}, dateSystem(source, def)]);
}
