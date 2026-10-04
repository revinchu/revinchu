import { createJsonSizer, jsonStringSize } from './json-size.js';
// 빈 셀 범위는 값/수식/메모가 없는 동일 서식만 묶는다. 좌표와 셀 수는 보존된다.
export function* storedCellEntries(cells) {
  if (cells.storageEntries) { yield* cells.storageEntries(); return; }
  for (const [key, cell] of cells) { const at=key.indexOf(','); yield [+key.slice(0,at),+key.slice(at+1),cell,1]; }
}
export function serializeCells(cells, encode) {
  const out={cells:{}},runs=[];
  for (const [r,c,cell,count] of storedCellEntries(cells)) {
    if(count>1) runs.push([r,c,count,encode(cell)]);
    else out.cells[`${r},${c}`]=encode(cell);
  }
  if(runs.length)out.cellRuns=runs;
  return out;
}
export function* cellJsonParts(cells, encode) {
  yield ',"cells":{';
  let first=true,chunk=[];
  for(const [r,c,cell,count] of storedCellEntries(cells)) {
    if(count>1)continue;
    chunk.push(JSON.stringify(`${r},${c}`)+':'+JSON.stringify(encode(cell)));
    if(chunk.length===512){yield (first?'':',')+chunk.join(',');first=false;chunk=[];}
  }
  if(chunk.length)yield (first?'':',')+chunk.join(',');
  yield '}';first=true;chunk=[];let emitted=false;
  for(const [r,c,cell,count] of storedCellEntries(cells)) {
    if(count<=1)continue;
    if(first){yield ',"cellRuns":[';first=false;}
    chunk.push(JSON.stringify([r,c,count,encode(cell)]));
    if(chunk.length===512){yield (emitted?',':'')+chunk.join(',');emitted=true;chunk=[];}
  }
  if(!first){if(chunk.length)yield (emitted?',':'')+chunk.join(',');yield ']';}
}

// 1 Mi UTF-16 characters: about 2 MiB of JS text, at most 3 MiB of UTF-8 JSON.
export const CELL_CHUNK_CHAR_LIMIT = 1 << 20;
/** Keep one oversized cell intact; every other chunk is bounded by both count and text. */
export function* storedCellChunks(cells, encode, size = 20000) {
  const measure=createJsonSizer(CELL_CHUNK_CHAR_LIMIT,{conservativeStrings:true}).size;
  let chunk=[],chars=2;
  for(const [r,c,cell,count] of storedCellEntries(cells)) {
    const data=encode(cell);
    let entryChars=80; // Coordinates, counts, brackets, property separators, and comma.
    for(const key in data)if(Object.hasOwn(data,key)&&data[key]!==undefined) {
      // encode copies the style; measure its identical source once per shared style.
      entryChars+=jsonStringSize(key,CELL_CHUNK_CHAR_LIMIT)+2+measure(key==='style'?cell.style:data[key]);
      if(entryChars>CELL_CHUNK_CHAR_LIMIT)break;
    }
    if(chunk.length&&chars+entryChars>CELL_CHUNK_CHAR_LIMIT){yield chunk;chunk=[];chars=2;}
    chunk.push([r,c,count,data]);chars+=entryChars;
    if(chunk.length>=size||chars>CELL_CHUNK_CHAR_LIMIT){yield chunk;chunk=[];chars=2;}
  }
  if(chunk.length)yield chunk;
}

/** JSON-decoded styles are immutable model values, just like imported shared XFs.
 * Reuse only byte-for-byte JSON-equivalent styles. Bound both the style count and
 * retained key text so documents with many unique styles cannot grow this memo.
 */
export function createStoredStyleMemo({ limit = 4096, charLimit = 1 << 20 } = {}) {
  const styles = new Map();
  let chars = 0;
  return style => {
    if (!style || typeof style !== 'object' || Array.isArray(style)) return style;
    const key = JSON.stringify(style), found = styles.get(key);
    if (found) return found;
    if (styles.size < limit && chars + key.length <= charLimit) {
      styles.set(key, style); chars += key.length;
    }
    return style;
  };
}
