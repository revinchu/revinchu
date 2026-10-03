import { idbGet, idbSet, idbCompareAndSet, idbKeys, idbDeleteMany } from './storage.js';

const BIG_PART_BYTES = 16 * 1024 * 1024;
const bigYield = () => new Promise(resolve => setTimeout(resolve, 0));
const bigAbort = () => Object.assign(new Error('저장 중 문서가 바뀌어 이전 저장본을 유지했습니다.'), { code: 'BIG_SAVE_ABORT' });
const bigGeneration = () => crypto.randomUUID();
const sheetRecordKey = (key, entry) => entry.key ?? `${key}#${entry.id}`;
// state만 달라진 시트는 기존 불변 셀 청크·블록을 새 메타 레코드에서 재사용합니다.
const stateFreeTag = meta => JSON.stringify({ ...meta, state: undefined });
const bigLock = (key, action, signal) => globalThis.navigator?.locks?.request
  ? navigator.locks.request(`wixel-large-save:${key}`, { ...(signal ? { signal } : {}) }, () => action(true))
  : action(false);
async function bigPack(text, gz) {
  const blob = new Blob([text], { type:'application/json' });
  return gz ? new Response(blob.stream().pipeThrough(new CompressionStream('gzip'))).blob() : blob;
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
  for (const entry of manifest?.sheets ?? []) {
    const recordKey = sheetRecordKey(key,entry); keep.add(recordKey);
    const record = await idbGet(recordKey);
    if (!record) throw new Error('저장된 시트가 없습니다.');
    for (const part of record.partKeys ?? []) keep.add(part);
  }
  const keys = await idbKeys(key+'#g#');
  await removeBigKeys(keys.filter(item => !keep.has(item)));
}

/** v3 = immutable sheet/part generations + one CAS manifest commit.
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
    const sheetMetadata = sheets.map((_,i) => book.sheetMeta(i)), sheetTags = sheetMetadata.map(item => JSON.stringify(item));
    const bookMeta = structuredClone(book.bookMeta());
    const valid = (checkMeta = false) => {
      if (options.signal?.aborted || options.isCurrent?.() === false || book.version !== version || book.sheets.length !== sheets.length || sheets.some((sheet,i) => book.sheets[i] !== sheet || (sheet._ev ?? 0) !== edits[i])) throw bigAbort();
      // Some document-level UI properties predate version tracking.
      if (checkMeta && JSON.stringify(book.bookMeta()) !== JSON.stringify(bookMeta)) throw bigAbort();
      if (checkMeta && sheetTags.some((tag,i) => JSON.stringify(book.sheetMeta(i)) !== tag)) throw bigAbort();
    };
    valid();
    const previous = await idbGet(key);
    const saved = new Map(([2,3].includes(previous?.v) ? previous.sheets : []).map(entry => [entry.id,entry]));
    const generation = bigGeneration(), staged = [], list = [], gz = typeof CompressionStream === 'function';
    let committed = false;
    try {
      // Remove abandoned generations from terminated tabs without touching current data.
      if (locked) await collectBigGarbage(key, previous).catch(() => {});
      valid();
      for (let i=0;i<sheets.length;i++) {
        const sheet=sheets[i], prior=saved.get(sheet._sid), ev=edits[i];
        const recordKey=`${key}#g#${generation}#s${i}`;
        if (previous?.v === 3 && prior?.ev === ev) {
          const previousRecord = await idbGet(sheetRecordKey(key,prior));
          if (previousRecord && JSON.stringify(previousRecord.meta) === sheetTags[i]) { list.push({...prior,key:sheetRecordKey(key,prior)}); valid(); continue; }
          if (previousRecord && stateFreeTag(previousRecord.meta) === stateFreeTag(sheetMetadata[i])) {
            // 원본 청크/partKeys는 staged에 넣지 않습니다. 실패 시 새 메타만 지우고,
            // 성공 뒤 GC는 새 레코드가 참조하는 기존 분할 청크를 계속 보존합니다.
            valid();staged.push(recordKey);
            await idbSet(recordKey,{...previousRecord,meta:sheetMetadata[i]});
            valid();list.push({id:sheet._sid,ev,key:recordKey});continue;
          }
        }
        const chunks=[], blocks=[], partKeys=[];
        for (const chunk of book.cellChunks(i)) {
          valid(); chunks.push(await bigPack(JSON.stringify(chunk),gz));
          await bigYield(); await options.waitForIdle?.(); valid();
        }
        for (let bi=0;bi<(sheet.blocks??[]).length;bi++) {
          const block=sheet.blocks[bi];
          const split = async (array,tag) => {
            if (!array || array.byteLength<=BIG_PART_BYTES) return {inline:array};
            const per=Math.floor(BIG_PART_BYTES/array.BYTES_PER_ELEMENT),parts=[];
            for(let offset=0,part=0;offset<array.length;offset+=per,part++) {
              valid();const partKey=`${recordKey}#p${bi}.${tag}.${part}`;staged.push(partKey);
              await idbSet(partKey,array.slice(offset,Math.min(array.length,offset+per)));parts.push(partKey);partKeys.push(partKey);
              await bigYield();await options.waitForIdle?.();valid();
            }
            return {parts,len:array.length,kind:array.constructor.name};
          };
          const cols=[];
          for(let ci=0;ci<block.cols.length;ci++) {
            const col=block.cols[ci],num=await split(col.num,`${ci}n`),str=await split(col.str,`${ci}s`);
            cols.push({...col,num:num.inline??null,str:str.inline??null,...(num.parts?{numParts:num}:{}),...(str.parts?{strParts:str}:{})});
          }
          const perm=await split(block.perm??null,'perm');
          blocks.push({...block,cols,perm:perm.inline??undefined,...(perm.parts?{permParts:perm}:{})});
        }
        valid();staged.push(recordKey);
        await idbSet(recordKey,{meta:sheetMetadata[i],chunks,gz,blocks,partKeys});
        valid();list.push({id:sheet._sid,ev,key:recordKey});
      }
      valid(true);
      const manifest={...meta,v:3,generation,book:bookMeta,sheets:list};
      await idbCompareAndSet(key,previous,manifest,()=>valid(true));committed=true;
      // GC failure cannot turn an already committed save into a reported failure.
      if(locked) await collectBigGarbage(key,manifest).catch(()=>{});
      return {manifest,version,book,isCurrent:()=>{try{valid(true);return true;}catch{return false;}}};
    } finally {
      if(!committed)await removeBigKeys(staged).catch(()=>{});
    }
  },options.signal);
}

/** Reads legacy v1/v2 and immutable v3 without discarding date-system book metadata. */
export async function loadLargeWorkbook(key,onProgress) {
  return bigLock(key,async()=>{
    const idx=await idbGet(key);if(!idx)return null;
    if(![2,3].includes(idx.v))return idx.workbook?idx:null;
    const sheets=[];
    for(let i=0;i<idx.sheets.length;i++) {
      const entry=idx.sheets[i],record=await idbGet(sheetRecordKey(key,entry));
      if(!record)throw new Error('저장된 시트가 없습니다. 이전 백업을 확인하세요.');
      const cells=new Map();
      for(const chunk of record.chunks) {
        for(const [cell,data] of JSON.parse(await bigUnpack(chunk,record.gz)))cells.set(cell,data);
        onProgress?.((i+.5)/idx.sheets.length);await bigYield();
      }
      const join=async info=>{
        const Ctor={Float64Array,Int32Array,Uint32Array,Float32Array,Uint8Array,Int16Array,Uint16Array}[info.kind];
        if(!Ctor||!Number.isSafeInteger(info.len)||info.len<0)throw new Error('저장된 열 블록 형식이 올바르지 않습니다.');
        const out=new Ctor(info.len);let at=0;
        for(const partKey of info.parts) {
          const part=await idbGet(partKey);if(!(part instanceof Ctor)||at+part.length>out.length)throw new Error('저장된 열 블록 조각이 없습니다.');
          out.set(part,at);at+=part.length;await bigYield();
        }
        if(at!==out.length)throw new Error('저장된 열 블록이 완전하지 않습니다.');
        return out;
      };
      const blocks=[];
      for(const block of record.blocks??[]) {
        const cols=[];
        for(const col of block.cols){const {numParts,strParts,...rest}=col;cols.push({...rest,num:numParts?await join(numParts):col.num,str:strParts?await join(strParts):col.str});}
        const {permParts,...rest}=block;blocks.push({...rest,cols,perm:permParts?await join(permParts):block.perm??undefined});
        onProgress?.((i+.9)/idx.sheets.length);
      }
      sheets.push({...record.meta,cells,blocks,_sid:entry.id,_ev:entry.ev});
    }
    return {rev:idx.rev,docName:idx.docName,docId:idx.docId,remoteDoc:idx.remoteDoc,si:idx.si,autosave:idx.autosave,storageFormat:idx.v,workbook:{...idx.book,...(idx.v===2?{names:idx.names,vba:idx.vba}:{}),sheets}};
  });
}
