/**
 * Enumerate occupied rows inside one inclusive area without expanding unrelated
 * rows of a dense/ compressed column. Consumers must not depend on key order:
 * short ranges use row order; long sparse ranges retain the column's live keys.
 */
export function* columnKeysInRange(column, r1, r2) {
  if (!column || !Number.isSafeInteger(r1) || !Number.isSafeInteger(r2) || r2 < r1) return;
  const height = r2 - r1 + 1;
  if (height <= column.size) {
    for (let r = r1; r <= r2; r++) if (column.has(r)) yield r;
  } else {
    for (const r of column.keys()) if (r >= r1 && r <= r2) yield r;
  }
}
