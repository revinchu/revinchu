// Share only scalar formulas whose immutable AST and relative origin agree.
// Bounded lookahead avoids a workbook-sized index and never changes the cells.
export function createXlsxSharedFormulas(cells, rowLimit, colLimit) {
  const active = new Map(); let nextId = 0;
  // Excel does not reliably preserve structured references in shared formulas.
  // Keep table/external references explicit even when the raw text is equal.
  const eligible = cell => cell?.formula && cell.ast && !cell.maybeArray && !cell.parseError && !cell.raw.includes('[');
  return (r, c, cell) => {
    if (c >= colLimit || r >= rowLimit || !eligible(cell)) return null;
    const previous = active.get(c);
    if (previous && r > previous.r1 && r <= previous.r2) return previous;
    const originR = r - (cell.dr ?? 0), originC = c - (cell.dc ?? 0);
    let end = r;
    for (const limit = Math.min(rowLimit, r + 4096); end + 1 < limit;) {
      const other = cells.getRC(end + 1, c);
      if (!eligible(other) || other.ast !== cell.ast) break;
      if (end + 1 - (other.dr ?? 0) !== originR || c - (other.dc ?? 0) !== originC) break;
      end++;
    }
    if (end === r) { active.delete(c); return null; }
    const run = { r1: r, r2: end, c, si: nextId++ };
    active.set(c, run); return run;
  };
}
