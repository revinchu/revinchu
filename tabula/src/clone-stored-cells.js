import { CellMap, shareBlankCell, getSharedBlankCell } from './cellmap.js';
import { storedCellEntries } from './cell-storage.js';

/** 셀의 논리 범위를 펼치지 않고 새 저장소로 복제합니다.
 * 일반 셀 mapper는 (cell,row,col), 빈 셀 mapper는 위치에 의존하지 않습니다.
 * 같은 빈 열 패턴은 한 번만 변환하며 다른 wrapper의 이후 변경과 분리합니다. */
export function cloneStoredCells(cells, pointMapper = cell => ({ ...cell }), blankMapper = cell => cell) {
  const out = new CellMap(), patterns = new WeakMap(), blanks = new WeakMap();
  const mappedBlank = cell => {
    if (blanks.has(cell)) return blanks.get(cell);
    const result = blankMapper(cell), next = result == null ? null : shareBlankCell(result);
    if (next && (next.raw !== '' || !next.style || next !== getSharedBlankCell(next.style, next.v === null))) throw new Error('공유 빈 셀 복제 결과가 불변 빈 셀이 아닙니다.');
    blanks.set(cell, next); return next;
  };
  if (cells instanceof CellMap) {
    for (const [c, column] of cells.cols) {
      if (!column.size) continue;
      if (column.blankOnly) {
        let copy = patterns.get(column.dataKey);
        if (!copy) {
          copy = column.shareData();
          for (const count of copy.mapValues(value => pointMapper(value, undefined, c), mappedBlank)) { /* 논리 범위만 처리 */ }
          patterns.set(column.dataKey, copy);
        }
        if (copy.size) out.replaceImportedColumn(c, copy.shareData());
      } else for (const [r, cell, count] of column.storageEntries()) {
        const next = column.isShared(cell) ? mappedBlank(cell) : pointMapper(cell, r, c);
        if (next) out.setRunRC(r, c, count, next);
      }
    }
  } else for (const [r, c, cell, count] of storedCellEntries(cells)) {
    const next = count > 1 ? mappedBlank(cell) : pointMapper(cell, r, c);
    if (next) out.setRunRC(r, c, count, next);
  }
  return out;
}