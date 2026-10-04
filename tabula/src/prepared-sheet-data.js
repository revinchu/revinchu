// An IDB read owns its freshly decoded cells and metadata. This capability is
// kept out of serialized data: only the opt-in loader can mark an input, and
// Workbook consumes it once. A copied sheet/map must not re-normalize cells
// already adopted by another Workbook.
const preparedBooks = new WeakMap(), preparedSheets = new WeakMap(), preparedCells = new WeakMap();
const invalidPrepared = () => Object.assign(new Error('준비된 복원 데이터를 다시 사용하거나 변경할 수 없습니다. 저장본을 다시 불러오세요.'), { code:'PREPARED_RESTORE_INVALID' });
export function markPreparedWorkbook(data, cachedArrays) {
  const record = { date1904:data.date1904 === true, consumed:false, sheets:data.sheets.map((sheet,i) => ({ sheet, cells:sheet.cells, blocks:sheet.blocks, cachedArrays:!!cachedArrays[i] })) };
  preparedBooks.set(data,record);
  for(const item of record.sheets) { preparedSheets.set(item.sheet,record); preparedCells.set(item.cells,record); }
  return data;
}
/** Validate every sheet before consuming any; errors leave the current book alone. */
export function takePreparedWorkbook(data) {
  const record = preparedBooks.get(data);
  if(!record) {
    for(const sheet of data.sheets) if(preparedSheets.has(sheet) || preparedCells.has(sheet.cells)) throw invalidPrepared();
    return null;
  }
  if(record.consumed || record.date1904 !== (data.date1904 === true) || record.sheets.length !== data.sheets.length) throw invalidPrepared();
  for(let i=0;i<record.sheets.length;i++) {
    const item=record.sheets[i],sheet=data.sheets[i];
    if(sheet!==item.sheet || sheet.cells!==item.cells || sheet.blocks!==item.blocks) throw invalidPrepared();
  }
  record.consumed=true;
  const result=record.sheets;
  record.sheets=null; // Do not retain adopted cell maps through a consumed input.
  return result;
}
