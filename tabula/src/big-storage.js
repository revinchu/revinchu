import { idbGet, idbSet, idbCompareAndSet, idbKeys, idbDeleteMany } from './storage.js';
import { CellMap } from './cellmap.js';
import { makeCellRC } from './workbook.js';
import { createScalarCellMemo } from './scalar-cell-memo.js';
import { markPreparedWorkbook } from './prepared-sheet-data.js';
import { createStoredFormulaMemo } from './stored-formula-memo.js';
import { createStoredStyleMemo } from './cell-storage.js';
import { storeArrayParts, restoreArrayParts } from './large-array-storage.js';
import { storedJsonEqual, currentSheetMetadata } from './storage-json-equal.js';

const savedSnapshots = new WeakMap(), savedCacheItems = new WeakMap();
const sameSnapshots = (a, b) => {
  const keys = Object.keys(a);
  return keys.length === Object.keys(b).length && keys.every(k => a[k] === b[k]);
};
const bigYield = () => new Promise(resolve => setTimeout(resolve, 0));
const bigAbort = () => Object.assign(new Error('저장 중 문서가 바뀌어 이전 저장본을 유지했습니다.'), { code: 'BIG_SAVE_ABORT' });
const bigGeneration = () => crypto.randomUUID();
const sheetRecordKey = (key, entry) => entry.key ?? `${key}#${entry.id}`;
// state만 달라진 시트는 기존 불변 셀 청크·블록을 새 메타 레코드에서 재사용합니다.
const sameSheetStateFree = (a,b) => storedJsonEqual(a,b,{omitRoot:['state']});
const bigLock = (key, action, signal) => globalThis.navigator?.locks?.request
  ? navigator.locks.request(`wixel-large-save:${key}`, { ...(signal ? { signal } : {}) }, () => action(true))
  : action(false);
async function bigPack(text, gz) {
  if (!gz) return new Blob([text], { type:'application/json' });
  // Creating an uncompressed input Blob per chunk delays native backing-store
  // release until browser Blob transport/GC catches up. Feed owned UTF-8 bytes
  // directly and retain only the compressed output Blob in IndexedDB.
  const stream = new CompressionStream('gzip'), writer = stream.writable.getWriter();
  const output = new Response(stream.readable).blob();
  // Start consuming before write: compression can apply readable backpressure.
  // Observe read failures immediately and unblock a pending transform write.
  output.catch(error => { void writer.abort(error).catch(() => {}); });
  try {
    await writer.write(new TextEncoder().encode(text));
    await writer.close();
    return await output;
  } catch (error) {
    await writer.abort(error).catch(() => {});
    await output.catch(() => {});
    throw error;
  } finally { writer.releaseLock(); }
}
async function bigUnpack(value, gz) {
  if (typeof value === 'string') return value;
  return gz ? new Response(value.stream().pipeThrough(new DecompressionStream('gzip'))).text() : value.text();
}
async function removeBigKeys(keys) {
  for (let i=0;i<keys.length;i+=128) { await idbDeleteMany(keys.slice(i,i+128)); await bigYield(); }
}
/** Called only while the same origin lock excludes v3 readers and writers.
 * Legacy tabs do not hold this lock, so never collect their v2 records/parts.
 */
async function collectBigGarbage(key, manifest) {
  const keep = new Set();
  for (const key of [manifest?.snapshotKey, manifest?.cacheItemsKey]) if (key) {
    keep.add(key); const record = await idbGet(key);
    if (!record) throw new Error('저장된 피벗 캐시가 없습니다.');
    if (['pivot-cache-parts-v1', 'cache-items-parts-v1'].includes(record.format)) for (const part of record.partKeys) keep.add(part);
  }
  for (const entry of manifest?.sheets ?? []) {
    const recordKey = sheetRecordKey(key,entry); keep.add(recordKey);
    const record = await idbGet(recordKey);
    if (!record) throw new Error('저장된 시트가 없습니다.');
    for (const part of record.partKeys ?? []) keep.add(part);
  }
  const keys = await idbKeys(key+'#g#');
  await removeBigKeys(keys.filter(item => !keep.has(item)));
}

/** v5 stores cell chunks and cache arrays as independently committed parts.
 * v2/v3/v4 remain readable, including inline cell chunks and pivot snapshots.
 * Immutable sheet/part generations + one CAS manifest commit.
 * Failure/cancellation never overwrites records referenced by the last manifest.
 * Locks also protect a reader holding an old manifest while garbage is collected.
 * Without Web Locks we retain old generations rather than race an unknown reader.
 */
export async function saveLargeWorkbook(key, book, metadata, options = {}) {
  return bigLock(key, async (locked) => {
    const sheets = [...book.sheets], version = book.version, ids = new Set();
    for (const sheet of sheets) {
      if (!sheet._sid || ids.has(sheet._sid)) sheet._sid = 's'+bigGeneration();
      ids.add(sheet._sid);
    }
    const edits = sheets.map(sheet => sheet._ev ?? 0), meta = structuredClone(metadata);
    // A queued Web Lock can start after more edits; tag the actual captured
    // checkpoint, not the version seen before waiting for the lock.
    if (meta.recovery) meta.recovery = { ...meta.recovery, version };
    const sheetMetadata = sheets.map((_,i) => book.sheetMeta(i));
    const bookMeta = structuredClone(book.bookMeta(false,false,{shareData:true})), snapshots = book.snapshotData?.() ?? {}, cacheItems = book.pivotCacheItems ?? null;
    const valid = (checkMeta = false) => {
      if (options.signal?.aborted || options.isCurrent?.() === false || book.version !== version || book.sheets.length !== sheets.length || sheets.some((sheet,i) => book.sheets[i] !== sheet || (sheet._ev ?? 0) !== edits[i])) throw bigAbort();
      // Some document-level UI properties predate version tracking.
      if (checkMeta && (book.pivotCacheItems ?? null) !== cacheItems) throw bigAbort();
      if (checkMeta && (!storedJsonEqual(book.bookMeta(false,false,{shareData:true}),bookMeta) || !sameSnapshots(snapshots, book.snapshotData?.() ?? {}))) throw bigAbort();
      if (checkMeta && sheetMetadata.some((meta,i) => !storedJsonEqual(currentSheetMetadata(sheets[i],meta),meta))) throw bigAbort();
    };
    valid();
    const previous = await idbGet(key);
    // The caller may bind a working copy to the complete generation it opened.
    // Checking before GC/staging prevents a later stale tab from overwriting a
    // different tab's sequential (already committed) edits, not just a CAS race.
    if (Object.hasOwn(options, 'expectedGeneration')) {
      const actualGeneration = [3,4,5].includes(previous?.v) ? previous.generation ?? null : null;
      if (actualGeneration !== options.expectedGeneration) throw Object.assign(new Error('다른 탭에서 이 문서의 저장본이 바뀌었습니다. 현재 편집 내용을 별도 사본으로 저장하세요.'), { code:'IDB_CONFLICT' });
    }
    const saved = new Map(([2,3,4,5].includes(previous?.v) ? previous.sheets : []).map(entry => [entry.id,entry]));
    const generation = bigGeneration(), staged = [], list = [], gz = typeof CompressionStream === 'function';
    let committed = false;
    const checkpoint = async () => { valid(); await bigYield(); await options.waitForIdle?.(); valid(); };
    const arrayParts = async (array, prefix, partKeys) => storeArrayParts(array, n => `${prefix}.${n}`, async (partKey, part) => {
      valid(); staged.push(partKey); await idbSet(partKey,part); partKeys.push(partKey); valid();
    }, checkpoint);
    const splitColumn = async (col, prefix, partKeys) => {
      const out = {...col};
      for (const name of ['num','str','dict']) if (col[name]) { out[name+'Parts'] = await arrayParts(col[name], `${prefix}.${name}`, partKeys); out[name] = null; }
      return out;
    };
    try {
      // Remove abandoned generations from terminated tabs without touching current data.
      if (locked) await collectBigGarbage(key, previous).catch(() => {});
      valid();
      for (let i=0;i<sheets.length;i++) {
        const sheet=sheets[i], prior=saved.get(sheet._sid), ev=edits[i];
        const recordKey=`${key}#g#${generation}#s${i}`;
        if ([3,4,5].includes(previous?.v) && prior?.ev === ev) {
          const previousRecord = await idbGet(sheetRecordKey(key,prior));
          if (previousRecord && storedJsonEqual(previousRecord.meta,sheetMetadata[i])) { list.push({...prior,key:sheetRecordKey(key,prior)}); valid(); continue; }
          if (previousRecord && sameSheetStateFree(previousRecord.meta,sheetMetadata[i])) {
            // 원본 청크/partKeys는 staged에 넣지 않습니다. 실패 시 새 메타만 지우고,
            // 성공 뒤 GC는 새 레코드가 참조하는 기존 분할 청크를 계속 보존합니다.
            valid();staged.push(recordKey);
            await idbSet(recordKey,{...previousRecord,meta:sheetMetadata[i]});
            valid();list.push({id:sheet._sid,ev,key:recordKey});continue;
          }
        }
        const cellPartKeys=[], blocks=[], partKeys=[];
        // JSON captures a chunk synchronously before the first await. Model styles
        // are immutable; edits replace them and valid() rejects a changed version.
        // Avoid allocating one shallow style copy per stored cell during autosave.
        for (const chunk of book.cellRunChunks(i, 20000, { shareStyle: true, bounded: true })) {
          valid(); const partKey=`${recordKey}#cells.${cellPartKeys.length}`; staged.push(partKey);
          const packed = await bigPack(JSON.stringify(chunk),gz); valid();
          await idbSet(partKey,packed); cellPartKeys.push(partKey); partKeys.push(partKey);
          await checkpoint();
        }
        for (let bi=0;bi<(sheet.blocks??[]).length;bi++) {
          const block=sheet.blocks[bi], cols=[];
          for(let ci=0;ci<block.cols.length;ci++) cols.push(await splitColumn(block.cols[ci],`${recordKey}#p${bi}.${ci}`,partKeys));
          const permParts=block.perm ? await arrayParts(block.perm,`${recordKey}#p${bi}.perm`,partKeys) : null;
          blocks.push({...block,cols,perm:undefined,...(permParts?{permParts}:{})});
        }
        valid();staged.push(recordKey);
        await idbSet(recordKey,{meta:sheetMetadata[i],cellPartKeys,cellEncoding:'runs-v1',gz,blocks,partKeys});
        valid();list.push({id:sheet._sid,ev,key:recordKey});
      }
      valid(true);
      let snapshotKey = null;
      if (Object.keys(snapshots).length) {
        const saved = savedSnapshots.get(book);
        if (saved && saved.key === previous?.snapshotKey && sameSnapshots(saved.data, snapshots)) snapshotKey = saved.key;
        else {
          snapshotKey = `${key}#g#${generation}#pivot-cache`;
          staged.push(snapshotKey); valid(true);
          const data={}, partKeys=[];
          for(const [id, snapshot] of Object.entries(snapshots)) {
            const prefix=`${snapshotKey}#${Object.keys(data).length}`;
            if(snapshot?.kind==='pivot-cache') {
              const columns=[];
              for(let ci=0;ci<snapshot.columns.length;ci++) columns.push(await splitColumn(snapshot.columns[ci],`${prefix}.${ci}`,partKeys));
              Object.defineProperty(data,id,{value:{...snapshot,columns},enumerable:true});
            } else Object.defineProperty(data,id,{value:{rowsParts:await arrayParts(snapshot,prefix,partKeys)},enumerable:true});
          }
          await idbSet(snapshotKey, {format:'pivot-cache-parts-v1',data,partKeys});
          await bigYield(); valid(true);
        }
      }
      let cacheItemsKey = null;
      if (cacheItems && Object.keys(cacheItems).length) {
        const saved = savedCacheItems.get(book);
        if (saved && saved.key === previous?.cacheItemsKey && saved.data === cacheItems) cacheItemsKey = saved.key;
        else {
          cacheItemsKey=`${key}#g#${generation}#cache-items`;staged.push(cacheItemsKey);
          const data={},partKeys=[];let cacheIndex=0;
          for(const [id,cache] of Object.entries(cacheItems)) {
            const fields=[];
            for(let fi=0;fi<cache.fields.length;fi++) {
              const field=cache.fields[fi], sharedParts=await arrayParts(field.shared??[],`${cacheItemsKey}#${cacheIndex}.${fi}`,partKeys);
              fields.push({...field,shared:undefined,sharedParts});
            }
            Object.defineProperty(data,id,{value:{...cache,fields},enumerable:true});cacheIndex++;
          }
          await idbSet(cacheItemsKey,{format:'cache-items-parts-v1',data,partKeys});valid(true);
        }
      }
      const manifest={...meta,v:5,generation,book:bookMeta,sheets:list,...(snapshotKey?{snapshotKey}:{}),...(cacheItemsKey?{cacheItemsKey}:{})};
      await idbCompareAndSet(key,previous,manifest,()=>valid(true));committed=true;
      savedSnapshots.set(book, { key: snapshotKey, data: snapshots });
      savedCacheItems.set(book, { key: cacheItemsKey, data: cacheItems });
      // GC failure cannot turn an already committed save into a reported failure.
      if(locked) await collectBigGarbage(key,manifest).catch(()=>{});
      return {manifest,version,book,isCurrent:()=>{try{valid(true);return true;}catch{return false;}}};
    } finally {
      if(!committed)await removeBigKeys(staged).catch(()=>{});
    }
  },options.signal);
}

/** Reads legacy v1/v2 and immutable v3 without discarding date-system book metadata. */
export async function loadLargeWorkbook(key,onProgress,{prepareCells=false}={}) {
  return bigLock(key,async()=>{
    const idx=await idbGet(key);if(!idx)return null;
    if(![2,3,4,5].includes(idx.v))return idx.workbook?idx:null;
    const sheets=[], cachedArrays=[], date1904=idx.book?.date1904===true;
    for(let i=0;i<idx.sheets.length;i++) {
      const entry=idx.sheets[i],record=await idbGet(sheetRecordKey(key,entry));
      if(!record)throw new Error('저장된 시트가 없습니다. 이전 백업을 확인하세요.');
      const cells=new CellMap(), shareStyle=createStoredStyleMemo();
      const shareScalar=prepareCells?createScalarCellMemo():null, shareFormula=createStoredFormulaMemo();
      let hasCachedArrays=false;
      // Decode a bounded chunk directly into the final cell representation. A
      // later Workbook adoption must not recreate all decoded cells a second time.
      const prepare=(data,r,c)=>{
        if(data?.style)data.style=shareStyle(data.style);
        if(typeof data?.raw==='string')data.raw=shareFormula.share(data.raw);
        if(!prepareCells)return data;
        const cell=shareScalar(makeCellRC(data,r,c,date1904));
        if(cell?.cachedArray)hasCachedArrays=true;
        return cell;
      };
      if (!Array.isArray(record.cellPartKeys ?? record.chunks)) throw new Error('저장된 셀 조각 목록이 없습니다.');
      try {
        for(const part of record.cellPartKeys ?? record.chunks) {
          const chunk=record.cellPartKeys ? await idbGet(part) : part;
          if (chunk == null) throw new Error('저장된 셀 조각이 없습니다. 이전 저장본을 확인하세요.');
          const entries=JSON.parse(await bigUnpack(chunk,record.gz));
          if(record.cellEncoding==='runs-v1')for(const [r,c,count,data] of entries){
            const cell=prepare(data,r,c);if(prepareCells&&cell==null)continue;
            if(count===1)cells.setRC(r,c,cell);else cells.setRunRC(r,c,count,cell);
          } else for(const [key,data] of entries){
            const comma=key.indexOf(','),r=+key.slice(0,comma),c=+key.slice(comma+1),cell=prepare(data,r,c);
            if(!prepareCells||cell!=null)cells.setRC(r,c,cell);
          }
          onProgress?.((i+.5)/idx.sheets.length);await bigYield();
        }
      } finally {shareFormula?.clear();}
      const join=info=>restoreArrayParts(info,idbGet,bigYield);
      const blocks=[];
      for(const block of record.blocks??[]) {
        const cols=[];
        for(const col of block.cols){const {numParts,strParts,dictParts,...rest}=col;cols.push({...rest,num:numParts?await join(numParts):col.num,str:strParts?await join(strParts):col.str,dict:dictParts?await join(dictParts):col.dict});}
        const {permParts,...rest}=block;blocks.push({...rest,cols,perm:permParts?await join(permParts):block.perm??undefined});
        onProgress?.((i+.9)/idx.sheets.length);
      }
      sheets.push({...record.meta,cells,blocks,_sid:entry.id,_ev:entry.ev});cachedArrays.push(hasCachedArrays);
    }
    let pivotSnapshots = idx.snapshotKey ? await idbGet(idx.snapshotKey) : idx.book?.pivotSnapshots;
    if (idx.snapshotKey && !pivotSnapshots) throw new Error('저장된 피벗 캐시가 없습니다. 이전 백업을 확인하세요.');
    if (pivotSnapshots?.format === 'pivot-cache-parts-v1') {
      const data={};
      for(const [id,snapshot] of Object.entries(pivotSnapshots.data)) {
        let value;
        if(snapshot.rowsParts) value=await restoreArrayParts(snapshot.rowsParts,idbGet,bigYield);
        else {
          const columns=[];
          for(const col of snapshot.columns) {
            const {numParts,strParts,dictParts,...rest}=col;
            columns.push({...rest,num:numParts?await restoreArrayParts(numParts,idbGet,bigYield):col.num,str:strParts?await restoreArrayParts(strParts,idbGet,bigYield):col.str,dict:dictParts?await restoreArrayParts(dictParts,idbGet,bigYield):col.dict});
          }
          value={...snapshot,columns};
        }
        Object.defineProperty(data,id,{value,enumerable:true});
      }
      pivotSnapshots=data;
    }
    let pivotCacheItems=idx.book?.pivotCacheItems;
    if(idx.cacheItemsKey) {
      const record=await idbGet(idx.cacheItemsKey);
      if(record?.format!=='cache-items-parts-v1')throw new Error('저장된 피벗 항목 목록이 없습니다.');
      pivotCacheItems={};
      for(const [id,cache] of Object.entries(record.data)) {
        const fields=[];
        for(const field of cache.fields) { const {sharedParts,...rest}=field;fields.push({...rest,shared:await restoreArrayParts(sharedParts,idbGet,bigYield)}); }
        Object.defineProperty(pivotCacheItems,id,{value:{...cache,fields},enumerable:true});
      }
    }
    const workbook={...idx.book,...(pivotCacheItems?{pivotCacheItems}:{}),...(pivotSnapshots?{pivotSnapshots}:{}),...(idx.v===2?{names:idx.names,vba:idx.vba}:{}),sheets};
    if(prepareCells)markPreparedWorkbook(workbook,cachedArrays);
    return {rev:idx.rev,docName:idx.docName,docId:idx.docId,remoteDoc:idx.remoteDoc,si:idx.si,autosave:idx.autosave,generation:idx.generation,recovery:idx.recovery,storageFormat:idx.v,workbook};
  });
}
