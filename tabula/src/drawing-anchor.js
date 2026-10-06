import { LINE_SHAPES } from './shapes.js';
// Excel drawing coordinates are sheet pixels at 96 dpi, independent of view zoom.
const PROPS = ['charts', 'images', 'shapes', 'slicers'];
const MARKERS = ['c1', 'r1', 'c2', 'r2'];
const near = (a, b) => Math.abs(a - b) < 1e-7;
const sameRect = (a, b) => !!a && !!b && ['x', 'y', 'w', 'h'].every(key => a[key] === b[key]);
const at = (axis, value) => {
  const index = axis.indexAt(Math.max(0, value));
  return [index, Math.max(0, value - axis.pos(index))];
};
const hiddenMarker = (axis, marker, value) => Array.isArray(marker)
  && marker.length === 2 && Number.isInteger(marker[0]) && marker[0] >= 0
  && marker[0] < (axis.max ?? Infinity) && Number.isFinite(marker[1]) && marker[1] >= 0
  && axis.size(marker[0]) === 0 && near(axis.pos(marker[0]) + marker[1], value);
const sameHiddenAnchor = (a, b) => (!a && !b) || (sameRect(a?.rect, b?.rect) && MARKERS.every(key => {
  const before = a?.markers?.[key], after = b?.markers?.[key];
  return !before && !after || !!before && !!after && before[0] === after[0] && before[1] === after[1];
}));

export function captureDrawingAnchors(sheet, cols, rows) {
  const anchors = new Map();
  for (const prop of PROPS) for (const object of sheet[prop] ?? []) {
    const placement = object.placement ?? (prop === 'slicers' ? 'oneCell' : 'twoCell');
    if (placement === 'absolute') continue;
    // Collapsed cells share a visible position. Keep their original identity through
    // hide/unhide and structuredClone (Undo/Redo), only while geometry still matches.
    const saved = sameRect(object._hiddenCellAnchor?.rect, object) ? object._hiddenCellAnchor.markers : null;
    const corner = (key, axis, value) => hiddenMarker(axis, saved?.[key], value) ? [...saved[key]] : at(axis, value);
    anchors.set(`${prop}:${object.id}`, { placement, c1: corner('c1', cols, object.x), r1: corner('r1', rows, object.y),
      c2: corner('c2', cols, object.x + object.w), r2: corner('r2', rows, object.y + object.h) });
  }
  return anchors;
}

export function reflowDrawingAnchors(sheet, anchors, cols, rows, { shift = null, moveOnly = false } = {}) {
  const move = (axis, [index, offset], kind) => {
    if (shift?.axis === kind) {
      if (shift.count > 0 && index >= shift.index) index += shift.count;
      else if (shift.count < 0) {
        const end = shift.index - shift.count;
        if (index >= end) index += shift.count;
        else if (index >= shift.index) { index = shift.index; offset = 0; }
      }
    }
    return { marker: [index, offset], value: axis.pos(index) + Math.min(offset, axis.size(index) || offset) };
  };
  const changes = {};
  for (const prop of PROPS) {
    let changed = false;
    const next = (sheet[prop] ?? []).map(object => {
      const anchor = anchors.get(`${prop}:${object.id}`);
      if (!anchor) return object;
      const c1 = move(cols, anchor.c1, 'col'), r1 = move(rows, anchor.r1, 'row');
      const c2 = move(cols, anchor.c2, 'col'), r2 = move(rows, anchor.r2, 'row');
      const resize = anchor.placement === 'twoCell' && !moveOnly;
      const candidate = { x: c1.value, y: r1.value,
        w: resize ? Math.max(1, c2.value - c1.value) : object.w,
        h: resize ? Math.max(LINE_SHAPES.has(object.kind) ? 0 : 1, r2.value - r1.value) : object.h };
      // An unchanged axis must never round fractional drawing coordinates.
      const rect = {};
      for (const key of ['x', 'y', 'w', 'h']) rect[key] = near(candidate[key], object[key]) ? object[key] : candidate[key];
      const markers = {};
      for (const [key, axis, moved, value] of [
        ['c1', cols, c1, rect.x], ['r1', rows, r1, rect.y],
        ['c2', cols, c2, rect.x + rect.w], ['r2', rows, r2, rect.y + rect.h],
      ]) if (hiddenMarker(axis, moved.marker, value)) markers[key] = moved.marker;
      const hidden = Object.keys(markers).length ? { rect, markers } : undefined;
      if (sameRect(rect, object) && sameHiddenAnchor(object._hiddenCellAnchor, hidden)) return object;
      changed = true;
      const updated = { ...object, ...rect };
      if (hidden) updated._hiddenCellAnchor = hidden; else delete updated._hiddenCellAnchor;
      return updated;
    });
    if (changed) changes[prop] = next;
  }
  return changes;
}
