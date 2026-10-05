const lower = value => String(value).toLowerCase();
const isAutomatic = sort => sort?.dir === 'asc' || sort?.dir === 'desc';

/** User-driven pivot sort changes. Initial file rendering continues to respect
 * the saved item order; only explicit sort requests invalidate that field's order.
 */
export function pivotSortPatch(def, { field, sort, customListSort } = {}) {
  const patch = {}, order = { ...(def.order ?? {}) };
  let changedOrder = false;
  const clearOrder = name => {
    for (const key of Object.keys(order)) if (lower(key) === lower(name)) { delete order[key];changedOrder = true; }
  };
  if (field !== undefined && field !== null) {
    const sorting = { ...(def.sort ?? {}) };
    const name = [...(def.rows ?? []), ...(def.cols ?? []), ...(def.pages ?? []), ...Object.keys(sorting), ...Object.keys(order)]
      .find(name => lower(name) === lower(field)) ?? field;
    for (const key of Object.keys(sorting)) if (lower(key) === lower(field)) delete sorting[key];
    if (isAutomatic(sort)) { sorting[name] = { ...sort };clearOrder(field); }
    patch.sort = sorting;
  }
  if (typeof customListSort === 'boolean' && customListSort !== (def.customListSort !== false)) {
    patch.customListSort = customListSort ? undefined : false;
    for (const [name, existing] of Object.entries(patch.sort ?? def.sort ?? {})) {
      if (isAutomatic(existing) && (existing.by === undefined || existing.by === null)) clearOrder(name);
    }
  }
  if (changedOrder) patch.order = order;
  return patch;
}
