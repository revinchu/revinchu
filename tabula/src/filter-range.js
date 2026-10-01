/** Keep an explicit filter range; only inspect newly appended contiguous rows.
 * Opening a menu must not rediscover all existing rows from its header.
 */
export function expandedFilterEnd(wb, si, filter) {
  const used = wb.usedRange(si);
  let end = filter.r2;
  const last = used.rows - 1, right = Math.min(filter.c2, used.cols - 1);
  if (end >= last || right < filter.c1) return end;
  for (let r = end + 1; r <= last; r++) {
    let filled = false;
    for (let c = filter.c1; c <= right; c++) {
      if (wb.getCell(si, r, c)?.raw) { filled = true; break; }
    }
    if (!filled) break;
    end = r;
  }
  return end;
}
