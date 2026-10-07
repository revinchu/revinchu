import { cellData, DEFAULT_ROW_HEIGHT } from './workbook.js';
import { tableCellDisplayStyle } from './table-format.js';
import { cellStyleDefaults, CELL_STYLE_PARTS } from './cell-style.js';
import { MAX_COLS } from './formula.js';
import { hid } from './axis.js';
const clone = value => value == null ? value : structuredClone(value);
const borderKeys = CELL_STYLE_PARTS.find(part => part[0] === 'border')[2];

/** 전체 행의 값 배열은 사용 열만 두고, 행 속성과 먼 서식 열은 별도로 보존한다. */
export function snapshotCopiedRows(wb, source) {
  if (source.kind !== 'rows') return {};
  const sh = wb.sheets[source.si], rows = source.rows;
  const base = cellStyleDefaults({ ...wb.baseStyle, ...sh.allStyle });
  const rowMeta = rows.map(r => ({ height: wb.rowHeight(source.si, r), manual: !!sh.rowManual[r], hidden: hid(sh.hiddenRows, r),
    style: sh.rowStyles[r] ? cellStyleDefaults({ ...base, ...clone(sh.rowStyles[r]) }) : base }));
  const first = source.c1 + (source.data[0]?.length ?? 0), columns = new Set();
  for (const c of sh.cells.cols.keys()) if (c >= first) columns.add(c);
  for (const key of Object.keys(sh.colStyles ?? {})) if (+key >= first && +key < MAX_COLS) columns.add(+key);
  for (const block of sh.blocks) if (rows.some(r => r >= block.r0 && r < block.r0 + block.n)) {
    for (let c = Math.max(first, block.c0); c < block.c0 + block.cols.length; c++) columns.add(c);
  }
  if (rows.length * ((source.data[0]?.length ?? 0) + columns.size) > 2000000) throw Error('한 번에 200만 셀까지 복사할 수 있습니다.');
  const styles = new Map(), tailCells = rows.map(() => []);
  rows.forEach((r, i) => {
    for (const c of columns) {
      const cell = wb.getCell(source.si, r, c);
      if (!cell && !sh.colStyles[c] && !wb.blockAt(source.si, r, c)) continue;
      const effective = tableCellDisplayStyle(wb, source.si, r, c), key = JSON.stringify(effective);
      if (!styles.has(key)) styles.set(key, cellStyleDefaults(clone(effective)));
      const data = cellData(cell, styles.get(key)) ?? { raw: '', style: styles.get(key) };
      tailCells[i].push({ c, data, value: clone(wb.getValue(source.si, r, c)) });
    }
  });
  return { rowMeta, tailCells };
}

/** 전체 행 덮어쓰기에서 source 희소 폭 밖의 실제 대상 열만 방문한다. */
export function copiedRowTailColumns(wb, si, source, area, { insertRows = false } = {}) {
  const first = source.data[0]?.length ?? 0, columns = new Set();
  for (const row of source.tailCells ?? []) for (const cell of row) if (cell.c >= first) columns.add(cell.c);
  // 삽입될 행은 비어 있다. 덮어쓰기는 대상 범위와 만나는 저장 구간만 확인한다.
  if (!insertRows) for (const [c, col] of wb.sheets[si].cells.cols) if (c >= first) {
    for (const [r, , count] of col.storageEntries()) if (r <= area.r2 && r + count > area.r1) { columns.add(c); break; }
  }
  if (!insertRows) for (const block of wb.sheets[si].blocks) if (block.r0 <= area.r2 && block.r0 + block.n > area.r1) {
    for (let c = Math.max(first, block.c0); c < block.c0 + block.cols.length; c++) columns.add(c);
  }
  if ((area.r2 - area.r1 + 1) * (first + columns.size) > 2000000) throw Error('한 번에 200만 셀까지 붙여넣을 수 있습니다.');
  return columns;
}

/** 복사한 행 높이·수동 높이·숨김·행 서식을 단일 트랜잭션 안에서 적용한다. */
export function applyCopiedRowMetadata(wb, si, source, area, what) {
  if (!source.rowMeta?.length) return;
  const sh = wb.sheets[si], heights = { ...sh.rowHeights }, manual = { ...sh.rowManual }, styles = { ...sh.rowStyles };
  let hidden = clone(sh.hiddenRows ?? {});
  if (hidden.__bits) {
    const start = Math.min(hidden.start, area.r1), end = Math.max(hidden.start + hidden.__bits.length, area.r2 + 1), bits = new Uint8Array(end - start);
    bits.set(hidden.__bits, hidden.start - start); hidden = { ...hidden, start, __bits: bits };
  }
  for (let r = area.r1; r <= area.r2; r++) {
    const incoming = source.rowMeta[(r - area.r1) % area.ph];
    if (incoming.height === (sh.defRowH ?? DEFAULT_ROW_HEIGHT) && !incoming.manual) delete heights[r]; else heights[r] = incoming.height;
    if (incoming.manual) manual[r] = true; else delete manual[r];
    if (hidden.__bits) hidden.__bits[r - hidden.start] = incoming.hidden ? 1 : 0;
    else if (incoming.hidden) hidden[r] = true; else delete hidden[r];
    const style = clone(incoming.style);
    if (what === 'noBorders') for (const key of borderKeys) { if (styles[r]?.[key] !== undefined) style[key] = styles[r][key]; else delete style[key]; }
    if (sh.protect?.on) for (const key of ['locked', 'hideFormula']) { if (styles[r]?.[key] !== undefined) style[key] = styles[r][key]; else delete style[key]; }
    styles[r] = style;
  }
  if (hidden.__bits) { hidden.count = 0; for (const bit of hidden.__bits) hidden.count += bit; }
  wb.setSheetProp(si, 'rowHeights', heights); wb.setSheetProp(si, 'rowManual', manual);
  wb.setSheetProp(si, 'rowStyles', styles); wb.setSheetProp(si, 'hiddenRows', hidden);
}
