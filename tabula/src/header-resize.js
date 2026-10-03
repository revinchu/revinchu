/**
 * Find a visible row/column header boundary near a pointer. Geometry is in
 * worksheet units; tolerance stays in CSS pixels when the worksheet is zoomed.
 * Frozen and scrolling panes are checked separately so a clipped, offscreen
 * boundary cannot resize a row/column hidden behind the frozen pane.
 */
export function headerResizeEdge(axis, coordinate, {
  zoom = 1, header = 0, viewport = Infinity, origin = 0, scroll = 0,
  frozenStart = 0, frozenEnd = 0, tolerance = 5,
} = {}) {
  if (!axis || !Number.isFinite(coordinate) || coordinate < header || coordinate >= viewport || !(viewport > header)) return null;
  const scale = Number.isFinite(zoom) && zoom > 0 ? zoom : 1;
  const slop = Math.max(0, tolerance) / scale;
  const frozenSize = Math.max(0, axis.pos(frozenEnd) - origin);
  const seam = Math.min(viewport, header + frozenSize);
  let found = null, distance = Infinity;
  const inspect = (start, end, sheetStart, first, last) => {
    if (!(end > start) || first > last || coordinate < start - slop || coordinate > end + slop) return;
    const point = Math.max(start, Math.min(coordinate, end));
    const index = axis.indexAt(sheetStart + point - start);
    // The containing item's trailing edge and the previous visible item's
    // trailing edge are the only possible nearest boundaries, even with hides.
    const candidates = [axis.nextVisible(index, -1), axis.nextVisible(index - 1, -1)];
    for (const item of candidates) {
      if (item < first || item > last || item < 0 || item >= axis.max || !(axis.size(item) > 0)) continue;
      const edge = start + axis.pos(item + 1) - sheetStart;
      // The pane's leading border belongs to an offscreen item, not a handle.
      if (edge <= start + 1e-7 || edge > end + 1e-7) continue;
      const delta = Math.abs(edge - coordinate);
      if (delta <= slop + 1e-7 && (delta < distance || (delta === distance && item < found))) { found = item; distance = delta; }
    }
  };
  if (frozenSize > 0) inspect(header, seam, origin, frozenStart, frozenEnd - 1);
  inspect(seam, viewport, origin + scroll + seam - header, frozenEnd, axis.max - 1);
  return found;
}
