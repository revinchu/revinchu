import { CellMap } from './cellmap.js';

const cachePrefix = si => String(si) + ':';
const layoutName = (key, si) => key.slice(cachePrefix(si).length);
const writtenKey = (si, name, layout) => cachePrefix(si) + name + ':' + layout.top + ',' + layout.left;

function pivotSlots(sheet) {
  const slots = [];
  if (sheet?.pivot) slots.push({ prop:'pivot', index:-1, def:sheet.pivot });
  (sheet?.pivotsExtra ?? []).forEach((def, index) => { if (def) slots.push({ prop:'pivotsExtra', index, def }); });
  return slots;
}

function findSlot(sheet, definition, name) {
  const slots = pivotSlots(sheet);
  const match = slots.find(slot => slot.def === definition);
  if (match) return { prop:match.prop, index:match.index, name:match.def.name ?? '' };
  const named = slots.filter(slot => (slot.def.name ?? '') === name);
  return named.length === 1 ? { prop:named[0].prop, index:named[0].index, name } : null;
}

function restoredDefinition(sheet, slot, name) {
  if (slot) {
    const def = slot.prop === 'pivot' ? sheet?.pivot : sheet?.pivotsExtra?.[slot.index];
    if (def && (def.name ?? '') === slot.name) return def;
    // 앞 피벗의 전체 삭제로 추가 피벗 배열이 압축될 수 있다.
    if (slot.prop === 'pivotsExtra' && slot.name) {
      const named = (sheet?.pivotsExtra ?? []).filter(item => item && (item.name ?? '') === slot.name);
      if (named.length === 1) return named[0];
    }
    return null;
  }
  const named = pivotSlots(sheet).filter(item => (item.def.name ?? '') === name);
  return named.length === 1 ? named[0].def : null;
}

function copyWritten(map, move = (r, c) => ({ r, c })) {
  const next = new CellMap();
  map?.forEachRC((style, r, c) => {
    const point = move(r, c);
    if (point) next.setRC(point.r, point.c, style);
  });
  return next;
}

function clearSheetCache(map, si) {
  const prefix = cachePrefix(si);
  for (const key of map.keys()) if (typeof key === 'string' && key.startsWith(prefix)) map.delete(key);
}

/** 서식 객체와 불변 역할 배열을 공유하며 해당 시트의 렌더 캐시만 보관한다. */
export function capturePivotCaches(layouts, written, si, sheet = null) {
  const prefix = cachePrefix(si), snapshot = { layouts:new Map(), written:new Map(), slots:new Map() }, owned = new Set();
  for (const [key, layout] of layouts) if (typeof key === 'string' && key.startsWith(prefix)) {
    const { sheet:cachedSheet, definition, ...saved } = layout;
    const name = layoutName(key, si), owner = sheet ?? cachedSheet, slot = findSlot(owner, definition, name);
    const current = restoredDefinition(owner, slot, name);
    // 일반 피벗 Undo 뒤 남은 이전 역할을 새 정의의 유효 캐시로 만들지 않는다.
    if (!slot || current !== definition || (sheet && (layout.top !== (current.top ?? 0) || layout.left !== (current.left ?? 0)))) continue;
    // Undo 내역이 거대한 이전 시트 객체를 붙잡지 않게 슬롯 정보만 보관한다.
    snapshot.layouts.set(key, saved);
    snapshot.slots.set(key, slot);
    owned.add(writtenKey(si, name, layout));
  }
  for (const [key, styles] of written) if (owned.has(key)) snapshot.written.set(key, copyWritten(styles));
  return snapshot;
}

/** Undo/Redo의 복제된 시트·정의에 다시 연결한다. 스냅샷 자체는 변경하지 않는다. */
export function restorePivotCaches(layouts, written, si, snapshot, afterSheet) {
  clearSheetCache(layouts, si); clearSheetCache(written, si);
  const restored = new Set();
  for (const [key, layout] of snapshot?.layouts ?? []) {
    const name = layoutName(key, si), definition = restoredDefinition(afterSheet, snapshot.slots?.get(key), name);
    if (!definition) continue;
    layouts.set(key, { ...layout, sheet:afterSheet, definition });
    restored.add(writtenKey(si, name, layout));
  }
  for (const [key, styles] of snapshot?.written ?? []) if (restored.has(key)) written.set(key, copyWritten(styles));
}

function moveCoordinate(point, index, count) {
  if (count > 0) return point >= index ? point + count : point;
  const end = index - count;
  if (point >= index && point < end) return null;
  return point >= end ? point + count : point;
}

function moveStart(point, index, count) {
  if (count >= 0) return point >= index ? point + count : point;
  return point >= index - count ? point + count : point >= index ? index : point;
}

function shiftVector(values, start, index, count, blank) {
  const at = index - start;
  if (count > 0) {
    if (at <= 0 || at >= values.length) return values;
    const added = Array.from({ length:count }, () => typeof blank === 'function' ? blank() : blank);
    return values.slice(0, at).concat(added, values.slice(at));
  }
  const from = Math.max(0, at), until = Math.min(values.length, at - count);
  return until > from ? values.slice(0, from).concat(values.slice(until)) : values;
}

function shiftedLayout(layout, axis, index, count) {
  let roles = layout.roles, rowDepth = layout.rowDepth;
  const top = moveStart(layout.top, index, axis === 'row' ? count : 0);
  const left = moveStart(layout.left, index, axis === 'col' ? count : 0);
  if (axis === 'row') {
    roles = shiftVector(roles, layout.top, index, count, () => []);
    if (rowDepth) rowDepth = shiftVector(rowDepth, layout.top, index, count, -1);
  } else {
    const next = roles.map(row => shiftVector(row, layout.left, index, count, ''));
    if (next.some((row, at) => row !== roles[at])) roles = next;
  }
  if (!roles.length || !roles.some(row => row.length)) return null;
  return { ...layout, top, left, roles, rowDepth };
}

/** 전체 행·열 이동에 맞춰 역할과 마지막 생성 서식의 절대 좌표를 함께 옮긴다. */
export function shiftPivotCaches(layouts, written, { axis, index, count }, si, beforeSheet, afterSheet) {
  if (!['row','col'].includes(axis) || !Number.isInteger(index) || index < 0 || !Number.isInteger(count)) return;
  const snapshot = capturePivotCaches(layouts, written, si, beforeSheet), moved = { layouts:new Map(), written:new Map(), slots:new Map() };
  for (const [key, layout] of snapshot.layouts) {
    const name = layoutName(key, si);
    const slot = snapshot.slots.get(key) ?? findSlot(beforeSheet, null, name);
    const definition = restoredDefinition(afterSheet, slot, name);
    if (!definition) continue;
    const next = shiftedLayout(layout, axis, index, count);
    if (!next) continue;
    next.sheet = afterSheet; next.definition = definition;
    moved.layouts.set(key, next); moved.slots.set(key, slot);
    const styles = snapshot.written.get(writtenKey(si, name, layout));
    if (styles) moved.written.set(writtenKey(si, name, next), copyWritten(styles, (r, c) => {
      const point = moveCoordinate(axis === 'row' ? r : c, index, count);
      return point === null ? null : axis === 'row' ? { r:point, c } : { r, c:point };
    }));
  }
  restorePivotCaches(layouts, written, si, moved, afterSheet);
}
