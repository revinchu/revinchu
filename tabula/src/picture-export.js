import { bakePictureEffects } from './picture-raster.js';
import { pictureVisual } from './picture-filter.js';

// Export-only object copies share immutable cell data; never write raster caches into
// the live book or its undo history. Image source and editable effects stay separate.
export async function preparePictureExport(book, progress = () => {}) {
  const version = book.version;
  let completed = 0;
  const prepare = async (object, picture = false) => {
    let next = object;
    if (picture) {
      const effectPng = await bakePictureEffects(object);
      if (effectPng || object.effectPng) next = { ...object, effectPng: effectPng || undefined };
      if (++completed % 4 === 0) { progress(completed); await new Promise(resolve => setTimeout(resolve, 0)); }
    }
    if (object.groupItems?.length) {
      const items = [];
      for (const child of object.groupItems) items.push(await prepare(child, child.kind === 'picture'));
      if (items.some((child, i) => child !== object.groupItems[i])) next = { ...next, groupItems: items };
    }
    return next;
  };
  const sheets = [];
  for (const sheet of book.sheets) {
    const images = [], shapes = [];
    for (const image of sheet.images || []) images.push(await prepare(image, true));
    for (const shape of sheet.shapes || []) shapes.push(await prepare(shape));
    sheets.push({ ...sheet, images, shapes });
  }
  if (book.version !== version) throw new Error('저장 중 문서가 변경되었습니다. 다시 저장하세요.');
  const snapshot = Object.create(book); snapshot.sheets = sheets;
  // Sheet identities changed for the export copy. Bind independent cache metadata
  // to these sheets while sharing the immutable rows; exporting must not evict live caches.
  snapshot.setSnapshots({ pivotSnapshots: book.snapshotData() });
  return snapshot;
}

/** Visible object bounds, including rotation, flipped reflection and filter bleed. */
export function pictureExportBounds(picture) {
  const w = Math.max(4, Number(picture.w) || 4), h = Math.max(4, Number(picture.h) || 4);
  const e = pictureVisual(picture).effects, s = e.shadow;
  const blur = e.artistic.type === 'blur' ? (.2 + e.artistic.amount * 5) * 3 : e.artistic.type === 'paint' ? (1 + e.artistic.amount * 3) * 3 : 0;
  const pad = Math.max(blur, (e.glow?.size || 0) * 2, s ? Math.max(Math.abs(s.dx), Math.abs(s.dy)) + s.blur * 2 : 0);
  const reflection = e.reflection ? h * e.reflection.size + e.reflection.gap : 0;
  const left = -pad, right = w + pad, top = -pad - (picture.flipV ? reflection : 0), bottom = h + pad + (picture.flipV ? 0 : reflection);
  const angle = (Number(picture.rot) || 0) * Math.PI / 180, co = Math.cos(angle), sn = Math.sin(angle);
  let x1 = Infinity, y1 = Infinity, x2 = -Infinity, y2 = -Infinity;
  for (const [x, y] of [[left,top],[right,top],[right,bottom],[left,bottom]]) {
    const dx = x - w / 2, dy = y - h / 2, xx = w / 2 + dx * co - dy * sn, yy = h / 2 + dx * sn + dy * co;
    x1 = Math.min(x1, xx); y1 = Math.min(y1, yy); x2 = Math.max(x2, xx); y2 = Math.max(y2, yy);
  }
  return { x: x1, y: y1, w: x2 - x1, h: y2 - y1 };
}
