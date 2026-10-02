// Add only rectangular neighbours; disjoint areas remain separate printed sections.
export function addPrintAreas(existing, selected) {
  const areas = existing.map(a => ({ ...a }));
  let next = { ...selected };
  for (let i = 0; i < areas.length;) {
    const a = areas[i];
    const contains = (x, y) => x.r1 <= y.r1 && x.r2 >= y.r2 && x.c1 <= y.c1 && x.c2 >= y.c2;
    const sameRows = a.r1 === next.r1 && a.r2 === next.r2 && a.c1 <= next.c2 + 1 && next.c1 <= a.c2 + 1;
    const sameCols = a.c1 === next.c1 && a.c2 === next.c2 && a.r1 <= next.r2 + 1 && next.r1 <= a.r2 + 1;
    if (contains(a, next)) return areas;
    if (contains(next, a) || sameRows || sameCols) {
      next = { r1: Math.min(a.r1, next.r1), r2: Math.max(a.r2, next.r2), c1: Math.min(a.c1, next.c1), c2: Math.max(a.c2, next.c2) };
      areas.splice(i, 1); i = 0;
    } else i++;
  }
  return [...areas, next];
}
export function resizePrintArea(area, edge, boundary) {
  const next = { ...area }, n = Math.max(0, Math.floor(boundary));
  if (edge === 'top') next.r1 = Math.min(n, next.r2);
  if (edge === 'bottom') next.r2 = Math.max(next.r1, n - 1);
  if (edge === 'left') next.c1 = Math.min(n, next.c2);
  if (edge === 'right') next.c2 = Math.max(next.c1, n - 1);
  return next;
}
