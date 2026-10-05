import { dataBottom, dataTop } from './tables.js';

const rectangle = r => ({ r1: r.r1, c1: r.c1, r2: r.r2, c2: r.c2 });
const contains = (r, cell) => cell && cell.r >= r.r1 && cell.r <= r.r2 && cell.c >= r.c1 && cell.c <= r.c2;

/** Resolve the same editable data rows for quick and custom sorting. */
export function sortScope({ range, activeOnly = false, active = null, table = null, filter = null, header = false }) {
  if (table) {
    const wholeTable = range.r1 === table.r1 && range.c1 === table.c1 && range.c2 === table.c2 &&
      (range.r2 === table.r2 || range.r2 === dataBottom(table));
    if (activeOnly || wholeTable) return {
      range: { r1: table.header ? table.r1 : dataTop(table), c1: table.c1, r2: dataBottom(table), c2: table.c2 },
      header: !!table.header,
    };
  }
  if (filter && activeOnly && contains(filter, active)) return { range: rectangle(filter), header: true };
  return { range: rectangle(range), header: !!header };
}

/** Re-evaluate every overlapping filter after cell movement, in the caller's transaction.
 * The supplied evaluator is shared with normal filter application, including condition,
 * color, top-N and average rules. Do not infer row visibility from its previous bitmap.
 */
export function refreshSortedFilters({ range, filters, keys = [], byCols = false, options = {}, recompute }) {
  if (range.r1 > range.r2 || range.c1 > range.c2) return [];
  const changed = [];
  for (const [key, filter] of filters) {
    if (!filter || range.r2 < filter.r1 + 1 || range.r1 > filter.r2 || range.c2 < filter.c1 || range.c1 > filter.c2) continue;
    const fullData = range.r1 === filter.r1 + 1 && range.r2 === filter.r2 && range.c1 === filter.c1 && range.c2 === filter.c2;
    const primary = keys[0];
    // Existing XLSX state describes one value column only. Clear stale arrows for
    // partial, horizontal, color, custom-list or multi-key sorts instead of claiming
    // the saved single-column state can reproduce those different operations.
    const simple = !byCols && !options.byCols && !options.caseSensitive && !options.natural && fullData && keys.length === 1 && primary &&
      (!primary.on || primary.on === 'value') && !primary.list && primary.at >= filter.c1 && primary.at <= filter.c2;
    const next = { ...filter, sort: simple ? { col: primary.at, asc: primary.asc !== false } : undefined };
    changed.push([key, recompute(next, key)]);
  }
  return changed;
}
