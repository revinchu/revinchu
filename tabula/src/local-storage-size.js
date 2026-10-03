import { createJsonSizer, jsonStringSize } from './json-size.js';
// localStorage 직렬화 전에 크기를 제한한다. 서버 연결/업로드 권한 판단에는 사용하지 않는다.
const CELL_LIMIT = 50000;
const PAYLOAD_LIMIT = 2e6;
const snapshotSizes = new WeakMap();
const stringSize = value => typeof value === 'string' ? jsonStringSize(value) : 0;
const scalarSize = value => value == null ? 4 : typeof value === 'string' ? stringSize(value)
  : typeof value === 'number' ? 24 : typeof value === 'boolean' ? 5 : stringSize(value.error ?? value.code) + 16;

/** 저장된 피벗 행은 불변이다. 같은 rows의 재계산은 하지 않고, 임계값까지만 센다.
 * 형식화 배열은 JSON 숫자 키·구두점으로 더 커지므로 원시 byteLength에 보수적인 여유를 둔다.
 * 정확한 파일 크기를 계산하는 함수가 아니며 작은 문서를 IDB로 보내는 것은 안전하다. */
function snapshotSize(rows) {
  if (!rows || typeof rows !== 'object') return 0;
  const memo = snapshotSizes.get(rows);
  if (memo !== undefined) return memo;
  let size = 0;
  const add = amount => { size = Math.min(PAYLOAD_LIMIT + 1, size + amount); return size > PAYLOAD_LIMIT; };
  if (rows.kind === 'pivot-cache') {
    for (const header of rows.header || []) if (add(stringSize(header))) break;
    if (size <= PAYLOAD_LIMIT) for (const col of rows.columns || []) {
      // num/str은 load 시 정규화된 typed array다. 원소를 JSON 객체로 복사하지 않는다.
      if (add((col.num?.byteLength || 0) * 4 + (col.str?.byteLength || 0) * 4)) break;
      for (const value of col.dict || []) if (add(scalarSize(value) + 1)) break;
      if (size > PAYLOAD_LIMIT) break;
    }
  } else if (Array.isArray(rows)) {
    add(rows.length * 3); // 행 배열의 괄호와 구분자도 보관 공간이다.
    if (size <= PAYLOAD_LIMIT) for (const row of rows) {
      if (Array.isArray(row)) { for (const value of row) if (add(scalarSize(value) + 1)) break; }
      else add(scalarSize(row) + 1);
      if (size > PAYLOAD_LIMIT) break;
    }
  }
  snapshotSizes.set(rows, size);
  return size;
}

// 직렬화 대상 속성만 읽습니다. Workbook의 AST/그래프/캐시나 전체 복사본은 만들지 않습니다.
const BOOK_FIELDS = ['date1904','calculation','vba','externals','defaultFont','baseStyle','cellStyles','objectStyles',
  'theme','themeXml','themeName','themeFonts','themeEffects','props'];
const SHEET_FIELDS = ['name','colWidths','rowHeights','merges','cond','colStyles','rowStyles','allStyle',
  'hiddenRows','hiddenCols','rowManual','freeze','filter','charts','pivot','validations','images','shapes','tables',
  'slicers','pivotsExtra','state','noGrid','noZeros','outline','protect','sparklines','page','defRowH','defColW',
  'zoom','view','tabColor','scenarios','external','protectedRanges','noteVisibility','fileValues'];
const CELL_FIELDS = ['raw','style','comment','link','image','phonetic','fx','inputType','cachedArray'];

/** 셀 내용과 모든 저장 메타데이터도 임계값까지만 세어 IndexedDB 저장을 선택합니다. */
export function isLargeLocalWorkbook(wb) {
  let cells=0,payload=0;
  // 큰 영역은 물리 셀/서식 메타데이터조차 순회할 필요가 없습니다.
  for(const sheet of wb.sheets||[]) {
    cells+=sheet.cells?.size||0;if(cells>CELL_LIMIT)return true;
    for(const block of sheet.blocks||[]){cells+=block.n*block.cols.length;if(cells>CELL_LIMIT)return true;}
  }
  const {size,fields}=createJsonSizer(PAYLOAD_LIMIT);
  const add=amount=>{payload+=amount;return payload>PAYLOAD_LIMIT;};
  if(add(fields(wb,BOOK_FIELDS)))return true;
  // 이름의 파싱 캐시(_ast/_text)는 Workbook.bookMeta가 저장하지 않습니다.
  for(const name of wb.names||[]) {
    for(const key in name)if(Object.hasOwn(name,key)&&key!=='_ast'&&key!=='_text') {
      if(add(key.length+4+size(name[key])))return true;
    }
  }
  for(const sheet of wb.sheets||[]) {
    if(add(fields(sheet,SHEET_FIELDS)))return true;
    const cellSize=cell=>fields(cell,CELL_FIELDS)+(cell?.formula&&cell.cached!==undefined?size(cell.cached)+16:0)+32;
    if(sheet.cells?.storageEntries) {
      // 연속 빈 셀은 저장 형식과 같이 범위 하나만 셉니다. 좌표를 펼치지 않습니다.
      for(const [,,cell] of sheet.cells.storageEntries())if(add(cellSize(cell)))return true;
    } else if(sheet.cells?.values) {
      for(const cell of sheet.cells.values())if(add(cellSize(cell)))return true;
    }
    for(const block of sheet.blocks||[]) {
      if(block.perm&&add(size(block.perm)))return true;
      for(const col of block.cols)if(add(fields(col,['num','str','dict','fmt'])))return true;
    }
  }
  for(const snapshot of wb.pivotSnapshots?.values()||[])if(add(snapshotSize(snapshot.rows)))return true;
  return false;
}
