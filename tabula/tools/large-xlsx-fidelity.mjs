// 비공개 실제 파일 검수용. 모델 데이터/논리 좌표 검사이며 Excel/Safari 화면 렌더 합격을 뜻하지 않습니다.
import { createHash } from 'node:crypto';
import { createReadStream, createWriteStream } from 'node:fs';
import { mkdir, readFile, writeFile, stat, realpath, readdir, open, rename } from 'node:fs/promises';
import { resolve, dirname, join, extname } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { createInterface } from 'node:readline';
import { once } from 'node:events';
import { spawn } from 'node:child_process';

const TOOL = fileURLToPath(import.meta.url), DEFAULT_SOURCE = resolve(dirname(TOOL), '../src'), NATIVE_SAMPLER=resolve(dirname(TOOL),'native-layout-compare.mjs');
const CATEGORIES = ['contents', 'styles', 'metadata', 'drawings', 'book', 'caches'];
const META = ['colWidths','rowHeights','merges','colStyles','rowStyles','allStyle','hiddenRows','hiddenCols','rowManual','cond','validations','tables','pivot','pivotsExtra','filter','freeze','state','noGrid','noZeros','outline','protect','sparklines','page','defRowH','defColW','zoom','view','tabColor','scenarios','external','protectedRanges','noteVisibility'];
const DRAW = ['charts','images','shapes','slicers'];
const ID_KEYS = new Set(['id','objectId','groupId','targetId','ownerId','imageId','shapeId','chartId','tableId','slicerId','snapshotId','cacheItemsId','expectedCacheItemsId']);

function tokens(value, put, ids, key = '', seen = new Set()) {
  if (typeof value === 'string') {
    if (ID_KEYS.has(key) && ids?.has(value)) value = ids.get(value);
    put('s'+value.length+':');
    for (let i=0;i<value.length;) {
      let end=Math.min(value.length,i+32768);
      if(end<value.length && value.charCodeAt(end-1)>=0xd800 && value.charCodeAt(end-1)<=0xdbff) end--;
      put(value.slice(i,end));i=end;
    }
    put(';');return;
  }
  if (value === null) { put('n;');return; }
  if (value === undefined) { put('u;');return; }
  if (typeof value === 'number') { put('d'+(Number.isNaN(value)?'NaN':(Object.is(value,-0)?'-0':String(value)))+';');return; }
  if (typeof value === 'boolean') { put(value?'t;':'f;');return; }
  if (typeof value === 'bigint') { put('i'+value+';');return; }
  if (typeof value !== 'object') throw new TypeError('검수할 수 없는 값: '+typeof value);
  if (seen.has(value)) throw new TypeError('검수 metadata에 순환 참조가 있습니다.');
  seen.add(value);
  try {
    if(value instanceof Date) { tokens(value.toISOString(),put,ids);return; }
    if(value instanceof ArrayBuffer || value instanceof Uint8Array) {
      const bytes=value instanceof ArrayBuffer?new Uint8Array(value):value;
      put('bytes:'+bytes.byteLength+':'+createHash('sha256').update(bytes).digest('hex')+';');return;
    }
    if(Array.isArray(value)||ArrayBuffer.isView(value)) {
      put('a[');for(const item of value)tokens(item,put,ids,'',seen);put(']');return;
    }
    if(value instanceof Map) {
      const entries=Array.from(value,([k,v])=>[canonicalHash(k),k,v]).sort((a,b)=>a[0].localeCompare(b[0]));
      put('m{');for(const [,k,v] of entries){tokens(k,put,ids,'',seen);tokens(v,put,ids,'',seen);}put('}');return;
    }
    if(value instanceof Set) { tokens([...value].sort((a,b)=>canonicalHash(a).localeCompare(canonicalHash(b))),put,ids,'',seen);return; }
    put('o{');
    for(const k of Object.keys(value).filter(k=>value[k]!==undefined && typeof value[k]!=='function').sort()) {
      tokens(k,put,ids);tokens(value[k],put,ids,k,seen);
    }
    put('}');
  } finally {seen.delete(value);}
}

export function canonicalHash(value, { ids } = {}) {
  const hash=createHash('sha256');tokens(value,part=>hash.update(part),ids);return hash.digest('hex');
}

function objectIds(sheet) {
  const ids=new Map();
  const visit=(value,path)=>{
    if(!value || typeof value!=='object')return;
    if(typeof value.id==='string')ids.set(value.id,'@'+path);
    if(Array.isArray(value))for(let i=0;i<value.length;i++)visit(value[i],path+'/'+i);
    else for(const k of ['children','items','image','groupItems'])if(value[k] && typeof value[k]==='object')visit(value[k],path+'/'+k);
  };
  for(const k of [...DRAW,'tables','sparklines'])visit(sheet[k],k);
  return ids;
}

// 한 컬럼의 실제 저장 구간만 정렬합니다. 압축 빈 셀의 논리 행 수만큼 배열을 만들지 않습니다.
function* cellIntervals(column) {
  if(!column)return;
  if(column.plain instanceof Map || column instanceof Map) {
    const points=column.plain??column;
    for(const row of [...points.keys()].sort((a,b)=>a-b))yield {r1:row,r2:row,cell:points.get(row)};
    return;
  }
  if(column.points instanceof Map && Array.isArray(column.runs)) {
    const points=[...column.points.keys()].sort((a,b)=>a-b);let pi=0,ri=0;
    while(pi<points.length || ri<column.runs.length) {
      if(ri>=column.runs.length || pi<points.length && points[pi]<column.runs[ri].start) {
        const row=points[pi++];yield {r1:row,r2:row,cell:column.points.get(row).value};
      } else {const run=column.runs[ri++];yield {r1:run.start,r2:run.end,cell:run.value};}
    }
    return;
  }
  const entries=[...column.storageEntries()].sort((a,b)=>a[0]-b[0]);
  for(const [row,cell,count]of entries)yield {r1:row,r2:row+count-1,cell};
}

function blockValueAt(block,column,i) {
  const physical=block.perm?block.perm[i]:i;
  if(column.str && column.str[physical]>=0)return column.dict[column.str[physical]];
  if(column.num && column.num[physical]===column.num[physical])return column.num[physical];
  return null;
}
function* blockIntervals(block,col) {
  const column=block.cols[col-block.c0];
  if(!block.n)return;
  if(!column.num && !column.str) {yield {r1:block.r0,r2:block.r0+block.n-1,value:null,fmt:column.fmt};return;}
  let start=0,value=blockValueAt(block,column,0);
  for(let i=1;i<=block.n;i++) {
    const next=i<block.n?blockValueAt(block,column,i):undefined;
    if(i===block.n || !Object.is(value,next)) {
      yield {r1:block.r0+start,r2:block.r0+i-1,value,fmt:column.fmt};start=i;value=next;
    }
  }
}

// 일반 셀이 블록보다 우선, 겹친 블록은 Workbook과 같이 먼저 등록된 블록이 우선입니다.
function* columnIntervals(sheet,col) {
  const readers=[cellIntervals(sheet.cells.cols.get(col)),...(sheet.blocks??[]).filter(b=>col>=b.c0 && col<b.c0+b.cols.length).map(b=>blockIntervals(b,col))];
  const current=readers.map(it=>it.next());
  for(;;) {
    let start=Infinity;
    for(const item of current)if(!item.done && item.value.r1<start)start=item.value.r1;
    if(start===Infinity)return;
    let selected=-1,end=Infinity;
    for(let i=0;i<current.length;i++) {
      const item=current[i];if(item.done)continue;
      if(item.value.r1<=start) {if(selected<0)selected=i;end=Math.min(end,item.value.r2);}
      else end=Math.min(end,item.value.r1-1);
    }
    yield {...current[selected].value,r1:start,r2:end};
    for(let i=0;i<current.length;i++)if(!current[i].done && current[i].value.r1<=end) {
      if(current[i].value.r2<=end)current[i]=readers[i].next();
      else current[i]={done:false,value:{...current[i].value,r1:end+1}};
    }
  }
}

function valueOf(value) {
  if(value && typeof value==='object') {
    const error=value.code??value.error;
    if(typeof error==='string')return {error};
    if(value.constructor?.name==='CellImage')return {image:true};
  }
  return value??null;
}
function contentOf(item) {
  const cell=item.cell;
  if(!cell)return {value:valueOf(item.value),formula:null};
  const result={value:valueOf(cell.formula?cell.cached:cell.v),formula:cell.formula?cell.raw:null};
  for(const k of ['comment','link','image','phonetic','cachedArray'])if(cell[k]!=null)result[k]=cell[k];
  return result;
}


// 캐시의 저장 path는 다시 저장할 때 바뀔 수 있습니다. 정의의 연결 순서로만 이름을 정규화합니다.
function pivotCacheIds(book) {
  const ids=new Map(),add=id=>{if(typeof id==='string' && !ids.has(id))ids.set(id,'@pivot-cache/'+ids.size);};
  for(const sheet of book.sheets)for(const def of [sheet.pivot,...(sheet.pivotsExtra??[])])if(def){add(def.snapshotId);add(def.cacheItemsId);}
  for(const id of [...new Set([...(book.pivotSnapshots?.keys()??[]),...Object.keys(book.pivotCacheItems??{})])].sort())add(id);
  return ids;
}
function cacheShape(rows) {
  if(Array.isArray(rows))return {header:rows[0]??[],n:Math.max(0,rows.length-1)};
  if(rows?.kind!=='pivot-cache' || rows.version!==1 || !Array.isArray(rows.header) || !Number.isSafeInteger(rows.n) || rows.n<0 || rows.columns?.length!==rows.header.length)throw new Error('검수할 피벗 캐시 표현이 올바르지 않습니다.');
  return {header:rows.header,n:rows.n};
}
function cacheSame(a,b) {
  if(Object.is(a,b))return true;
  const ae=a && typeof a==='object'?(a.error??a.code):undefined,be=b && typeof b==='object'?(b.error??b.code):undefined;
  return typeof ae==='string' && ae===be;
}
function* cacheIntervals(rows,col,n) {
  if(!n)return;
  const column=Array.isArray(rows)?null:rows.columns[col];
  const tail=column?.tail;
  if(tail && (!Number.isSafeInteger(tail.start)||tail.start<0||tail.start>=n||!Object.hasOwn(tail,'value')))throw new Error('검수할 피벗 캐시 꼬리가 올바르지 않습니다.');
  const prefix=Array.isArray(rows)?n:Math.min(n,tail?.start??Math.max(column.num?.length??0,column.str?.length??0));
  const valueAt=i=>{
    if(Array.isArray(rows))return rows[i+1]?.[col]??null;
    if(column.str && column.str[i]>=0)return column.dict[column.str[i]]??null;
    const num=column.num?.[i];return typeof num==='number' && !Number.isNaN(num)?num:null;
  };
  let start=0,previous,initialized=false;
  for(let i=0;i<prefix;i++) {
    const value=valueAt(i);
    if(!initialized){previous=value;initialized=true;}
    else if(!cacheSame(previous,value)){yield {r1:start,r2:i-1,value:previous};start=i;previous=value;}
  }
  const remainder=tail?tail.value:null;
  if(prefix<n) {
    if(!initialized){previous=remainder;initialized=true;start=prefix;}
    else if(!cacheSame(previous,remainder)){yield {r1:start,r2:prefix-1,value:previous};start=prefix;previous=remainder;}
  }
  if(initialized)yield {r1:start,r2:n-1,value:previous};
}
async function digestPivotCaches(book,ids,publish,onProgress) {
  const result={snapshots:{count:0,rows:0,logicalValues:0,runs:0,hash:''},items:{count:0,hash:''}},hashes={snapshots:createHash('sha256'),items:createHash('sha256')};
  const record=async(type,entry)=>{hashes[type].update(canonicalHash(entry)+'\n');await publish({sheet:-1,category:'caches',type,...entry});};
  for(const [id,snapshot]of [...(book.pivotSnapshots??[])].sort((a,b)=>ids.get(a[0]).localeCompare(ids.get(b[0])))) {
    const rows=snapshot.rows,{header,n}=cacheShape(rows),key=ids.get(id);result.snapshots.count++;result.snapshots.rows+=n;
    await record('snapshots',{key,hash:canonicalHash({header,n})});
    for(let col=0;col<header.length;col++) {
      let pending=null,work=0;
      for(const run of cacheIntervals(rows,col,n)) {
        const hash=canonicalHash(valueOf(run.value));
        if(pending && pending.r2+1===run.r1 && pending.hash===hash)pending.r2=run.r2;
        else {if(pending){result.snapshots.runs++;await record('snapshots',{key,col,...pending});}pending={r1:run.r1,r2:run.r2,hash};}
        result.snapshots.logicalValues+=run.r2-run.r1+1;
        if(++work%65536===0){onProgress?.({stage:'cache-digest',cache:key,column:col,rows:n});await new Promise(done=>setImmediate(done));}
      }
      if(pending){result.snapshots.runs++;await record('snapshots',{key,col,...pending});}
    }
    onProgress?.({stage:'cache-digest',cache:key,rows:n});await new Promise(done=>setImmediate(done));
  }
  for(const id of Object.keys(book.pivotCacheItems??{}).sort((a,b)=>ids.get(a).localeCompare(ids.get(b)))) {
    result.items.count++;await record('items',{key:ids.get(id),hash:canonicalHash(book.pivotCacheItems[id])});
  }
  for(const type of Object.keys(hashes))result[type].hash=hashes[type].digest('hex');
  result.hash=canonicalHash({snapshots:result.snapshots.hash,items:result.items.hash});return result;
}

export async function digestWorkbook(book, {onProgress,onRecord,chunkRows=65536,maxSamples=8} = {}) {
  if(!Number.isSafeInteger(chunkRows)||chunkRows<1)throw new Error('chunkRows는 양의 정수여야 합니다.');
  const result={schemaVersion:1,scope:{data:'전체 sparse cells 및 모든 blocks의 논리 값·수식·직접 서식; 저장 피벗 캐시 전체 논리 값과 과거 항목',geometry:'모델의 전체 시트 크기·병합·drawing anchor/선/색/내용',formulaEvaluation:false,nativeExcel:false,visualRender:false,conditionalFormatting:'정의 보존 검사; 실제 표시 결과 아님',excludedVolatileFields:['props.modified'],idNormalization:'generated drawing/table IDs including nested groupItems and sourceReference.expectedCacheItemsId using pivot-cache paths by definition reference order; names/captions/source/position retained'},book:{},sheets:[]};
  const started=performance.now();let processed=0,nextProgress=chunkRows;
  const publish=async(record)=>{if(onRecord)await onRecord(record);};
  const cacheIds=pivotCacheIds(book);
  const meta=book.bookMeta(false,false,{shareData:true});
  for(const key of Object.keys(meta).sort()) {
    let value=meta[key];if(key==='props'){const {modified,...stable}=value;value=stable;result.volatileModifiedHash=canonicalHash(modified);}const hash=canonicalHash(value);result.book[key]=hash;
    await publish({sheet:-1,category:'book',key,hash});
  }
  result.caches=await digestPivotCaches(book,cacheIds,publish,onProgress);
  for(let si=0;si<book.sheets.length;si++) {
    const sheet=book.sheets[si],ids=new Map([...cacheIds,...objectIds(sheet)]),styleMemo=new WeakMap();
    if(!sheet.cells?.cols)throw new Error('CellMap.cols 인터페이스가 필요합니다.');
    const hashes=Object.fromEntries(CATEGORIES.filter(k=>!['book','caches'].includes(k)).map(k=>[k,createHash('sha256')]));
    const row={index:si,name:sheet.name,hashes:{},properties:{},counts:{logicalCells:0,formulas:0,blankCells:0,values:0,cellMapLogical:sheet.cells.size,blockStoredSlots:0,blockValues:0,canonicalRuns:0},samples:[]};
    const record=async(category,entry)=>{
      const rec={sheet:si,category,...entry};hashes[category].update(canonicalHash(entry)+'\n');
      if(row.samples.length<maxSamples)row.samples.push(rec);
      await publish(rec);
    };
    await record('metadata',{key:'name',hash:canonicalHash(sheet.name)});
    for(const key of META) {
      const value=sheet[key]??null,hash=canonicalHash(value,{ids});
      row.properties[key]=hash;
      await record('metadata',{key,hash,count:Array.isArray(value)?value.length:value&&typeof value==='object'?Object.keys(value).length:value==null?0:1});
    }
    for(const key of DRAW) {
      const list=sheet[key]??[];row.counts[key]=list.length;
      for(let i=0;i<list.length;i++)await record('drawings',{key,index:i,hash:canonicalHash(list[i],{ids})});
    }
    const columns=new Set(sheet.cells.cols.keys());
    for(const block of sheet.blocks??[])for(let c=0;c<block.cols.length;c++) {columns.add(block.c0+c);row.counts.blockStoredSlots+=block.n;}
    const styledRows=Object.keys(sheet.rowStyles??{}).map(Number).sort((a,b)=>a-b);
    for(const col of [...columns].sort((a,b)=>a-b)) {
      let contentRun=null,styleRun=null,styleRowIndex=0;
      const flush=async(category,run)=>{if(run){row.counts.canonicalRuns++;await record(category,{col,r1:run.r1,r2:run.r2,hash:run.hash});}};
      for(const item of columnIntervals(sheet,col)) {
        // 할당된 블록의 빈 슬롯은 다른 블록을 가리지만, 직접 서식도 없으면 저장 셀은 아닙니다.
        if(!item.cell && item.value===null && !item.fmt)continue;
        let r=item.r1;
        while(r<=item.r2) {
          while(styleRowIndex<styledRows.length && styledRows[styleRowIndex]<r)styleRowIndex++;
          let end=item.r2;
          if(styleRowIndex<styledRows.length && styledRows[styleRowIndex]<=end)end=styledRows[styleRowIndex]===r?r:styledRows[styleRowIndex]-1;
          const content=contentOf(item),contentHash=canonicalHash(content,{ids}),style=book.styleAt(si,r,col);
          let styleHash=style && typeof style==='object'?styleMemo.get(style):undefined;
          if(!styleHash) {styleHash=canonicalHash(style??{});if(style && typeof style==='object')styleMemo.set(style,styleHash);}
          const count=end-r+1;row.counts.logicalCells+=count;processed+=count;
          if(item.cell?.formula)row.counts.formulas+=count;
          if(content.value===null && !item.cell?.formula)row.counts.blankCells+=count;else row.counts.values+=count;
          if(!item.cell && item.value!==null)row.counts.blockValues+=count;
          if(contentRun && contentRun.r2+1===r && contentRun.hash===contentHash)contentRun.r2=end;
          else {await flush('contents',contentRun);contentRun={r1:r,r2:end,hash:contentHash};}
          if(styleRun && styleRun.r2+1===r && styleRun.hash===styleHash)styleRun.r2=end;
          else {await flush('styles',styleRun);styleRun={r1:r,r2:end,hash:styleHash};}
          r=end+1;
          if(processed>=nextProgress) {
            onProgress?.({stage:'digest',sheet:si,logicalCells:processed,elapsedMs:performance.now()-started});
            nextProgress=processed+chunkRows;await new Promise(done=>setImmediate(done));
          }
        }
      }
      await flush('contents',contentRun);await flush('styles',styleRun);
    }
    for(const key of Object.keys(hashes))row.hashes[key]=hashes[key].digest('hex');
    row.hash=canonicalHash({name:row.name,hashes:row.hashes,logicalCells:row.counts.logicalCells,formulas:row.counts.formulas});
    result.sheets.push(row);
  }
  result.hash=canonicalHash({book:result.book,caches:result.caches.hash,sheets:result.sheets.map(s=>({name:s.name,hash:s.hash}))});
  result.elapsedMs=performance.now()-started;
  return result;
}

export function compareDigests(before,after) {
  const differences=[];
  for(const type of ['snapshots','items'])if(before.caches?.[type]?.hash!==after.caches?.[type]?.hash)differences.push({sheet:-1,category:'caches.'+type});
  for(const key of new Set([...Object.keys(before.book),...Object.keys(after.book)]))if(before.book[key]!==after.book[key])differences.push({sheet:-1,category:'book.'+key});
  if(before.sheets.length!==after.sheets.length)differences.push({sheet:-1,category:'sheet-count',before:before.sheets.length,after:after.sheets.length});
  for(let si=0;si<Math.max(before.sheets.length,after.sheets.length);si++) {
    const a=before.sheets[si],b=after.sheets[si];
    if(!a || !b) {differences.push({sheet:si,category:'sheet-presence'});continue;}
    if(a.name!==b.name)differences.push({sheet:si,category:'sheet-name',before:a.name,after:b.name});
    for(const category of ['contents','styles','metadata','drawings'])if(a.hashes[category]!==b.hashes[category])differences.push({sheet:si,name:a.name,category});
    for(const key of ['logicalCells','formulas'])if(a.counts[key]!==b.counts[key])differences.push({sheet:si,name:a.name,category:'count.'+key,before:a.counts[key],after:b.counts[key]});
  }
  return {equal:differences.length===0,differences};
}

async function hashFile(path) {
  const hash=createHash('sha256');for await(const chunk of createReadStream(path))hash.update(chunk);return hash.digest('hex');
}
async function fingerprint(path) {
  const info=await stat(path);return {size:info.size,mtimeMs:info.mtimeMs,sha256:await hashFile(path)};
}
async function sourceFingerprint(source) {
  const files={};
  const visit=async(root,prefix='')=>{for(const entry of (await readdir(root,{withFileTypes:true})).sort((a,b)=>a.name.localeCompare(b.name))) {
    if(entry.isDirectory())await visit(join(root,entry.name),prefix+entry.name+'/');
    else if(entry.name.endsWith('.js'))files[prefix+entry.name]=await hashFile(join(root,entry.name));
  }};
  await visit(source);return {hash:canonicalHash(files),files};
}
export async function importWorkbook(path,{source=DEFAULT_SOURCE,asyncRead=true,onProgress,onDiagnostics}={}) {
  const xlsx=await import(pathToFileURL(join(resolve(source),'xlsx.js')).href);
  const {Workbook}=await import(pathToFileURL(join(resolve(source),'workbook.js')).href);
  let bytes=await readFile(path),parsed=asyncRead?await xlsx.readXlsxAsync(bytes,onProgress,{onDiagnostics}):xlsx.readXlsx(bytes);
  bytes=null;
  const book=new Workbook();
  if(asyncRead)await book.loadAsync(parsed.data,p=>onProgress?.({stage:'load',p}));else book.load(parsed.data);
  const warnings=parsed.warnings??[];parsed.data=null;parsed=null;
  return {book,warnings,xlsx};
}

// 구간별 hash는 전부 계산하되 disk 진단은 4096행 band로 묶어 고객 값 크기만큼 log를 만들지 않습니다.
function diagnosticWriter(streams,bandRows=4096) {
  const pending=new Map();
  const writeLine=async record=>{const stream=streams.get(record.category);if(stream.errored)throw stream.errored;if(!stream.write(JSON.stringify(record)+'\n'))await once(stream,'drain');};
  const flush=async category=>{const item=pending.get(category);if(!item)return;pending.delete(category);const {hash,...record}=item;await writeLine({...record,hash:hash.digest('hex')});};
  return {
    async write(record) {
      if(record.r1===undefined){await flush(record.category);await writeLine(record);return;}
      for(let row=record.r1;row<=record.r2;) {
        const band=Math.floor(row/bandRows),end=Math.min(record.r2,(band+1)*bandRows-1);
        let item=pending.get(record.category);
        if(item && (item.sheet!==record.sheet || item.col!==record.col || item.band!==band || item.key!==record.key || item.type!==record.type)){await flush(record.category);item=null;}
        if(!item){item={sheet:record.sheet,category:record.category,...(record.key!==undefined?{key:record.key,type:record.type}:{}),col:record.col,band,r1:row,r2:end,records:0,hash:createHash('sha256')};pending.set(record.category,item);}
        item.r2=end;item.records++;item.hash.update(row+','+end+','+record.hash+';');row=end+1;
      }
    },
    async finish(){for(const category of pending.keys())await flush(category);}
  };
}
// callback 시점의 관측 최고값이며, callback 사이의 순간 peak나 Safari 실기기 메모리 한도가 아닙니다.
export function createMemoryMetrics({readMemory=()=>process.memoryUsage(),now=()=>performance.now()}={}) {
  const started=now(),highWater={};let samples=0;
  return {
    sample(stage) {
      const usage=readMemory(),snapshot={stage,elapsedMs:now()-started};samples++;
      for(const key of ['rss','heapUsed','arrayBuffers','external','heapTotal']) {
        const bytes=usage[key];if(!Number.isFinite(bytes))continue;snapshot[key]=bytes;
        if(!highWater[key]||bytes>highWater[key].bytes)highWater[key]={bytes,stage,elapsedMs:snapshot.elapsedMs};
      }
      return snapshot;
    },
    report(){return {samples,highWater:Object.fromEntries(Object.entries(highWater).map(([key,value])=>[key,{...value}])),scope:'Node callback 시점의 관측 최고값; renderer·물리 iPad Safari 미검사'};}
  };
}
async function phase(configPath,phaseName) {
  const config=JSON.parse(await readFile(configPath,'utf8')),folder=join(config.out,phaseName);
  await mkdir(folder,{recursive:true});
  const input=phaseName==='original'?config.input:config.roundtrip;
  const report={phase:phaseName,input,started:new Date().toISOString(),sourceBefore:await sourceFingerprint(config.source),inputBefore:await fingerprint(input),auditToolHash:await hashFile(TOOL),diagnostics:[],scope:'Node 전체 모델 데이터·직접 서식·논리 geometry; Native Excel/화면/Safari 실기기 미검사'};
  const lastDiagnostics=new Map(),streams=new Map(CATEGORIES.map(k=>[k,createWriteStream(join(folder,k+'.ndjson'))]));
  for(const stream of streams.values())stream.on('error',()=>{});const diagnostics=diagnosticWriter(streams);report.diagnosticBandRows=4096;
  let lastProgress=0;
  const importMemory=createMemoryMetrics();let importing=false;
  report.memory={scope:'Node process 메모리와 callback 관측값; renderer·물리 iPad Safari 미검사'};
  const progress=data=>{if(importing)importMemory.sample(data.stage??'import');if(Date.now()-lastProgress>2000){console.log(JSON.stringify({phase:phaseName,stage:data.stage??'import',p:data.p??null,logicalCells:data.logicalCells??null}));lastProgress=Date.now();}};
  try {
    const importStarted=performance.now();importing=true;report.memory.beforeImport=importMemory.sample('import-start');
    const loaded=await importWorkbook(input,{source:config.source,asyncRead:config.asyncRead,onProgress:progress,onDiagnostics:item=>{importMemory.sample(item.stage??('import-diagnostic:'+String(item.part??'unknown')));lastDiagnostics.set(item.part,item);}});
    report.importMs=performance.now()-importStarted;importing=false;
    report.memory.afterImport=importMemory.sample('import-complete');report.memory.import=importMemory.report();
    const gcStarted=performance.now(),gcAvailable=typeof globalThis.gc==='function';if(gcAvailable)globalThis.gc();
    report.modelGc={performed:gcAvailable,elapsedMs:performance.now()-gcStarted};report.memory.modelAfterGc=importMemory.sample('model-after-gc');
    report.warnings=loaded.warnings;
    if(config.nativeTargets) {
      const targets=JSON.parse((await readFile(config.nativeTargets,'utf8')).replace(/^\ufeff/,'')),{sampleNativeTargets}=await import(pathToFileURL(NATIVE_SAMPLER).href);
      report.nativeSamplerHash=await hashFile(NATIVE_SAMPLER);report.nativeTargetsFingerprint=await fingerprint(config.nativeTargets);
      const samples=await sampleNativeTargets(loaded.book,targets,{source:config.source});
      await writeFile(join(folder,'native-model-samples.json'),JSON.stringify(samples,null,2));
      report.nativeModelSamples={file:'native-model-samples.json',sheets:samples.sheets.length,cells:samples.sheets.reduce((n,s)=>n+s.cells.length,0),scope:samples.scope};
    }
    const digestStarted=performance.now();
    report.digest=await digestWorkbook(loaded.book,{onProgress:progress,maxSamples:config.maxDiffs,onRecord:record=>diagnostics.write(record)});
    report.digestMs=performance.now()-digestStarted;report.memory.afterDigest=importMemory.sample('digest-complete');
    await diagnostics.finish();report.diagnostics=[...lastDiagnostics.values()];
    await writeFile(join(folder,'digest.json'),JSON.stringify(report.digest,null,2));
    // 기준 digest를 디스크에 기록한 뒤 같은 모델 하나만 사용해 순차 ZIP sink로 저장합니다.
    if(phaseName==='original' && config.roundtrip) {
      const exportStarted=performance.now(),temp=config.roundtrip+'.partial-'+process.pid,handle=await open(temp,'wx');
      try {
        await loaded.xlsx.writeXlsxToSink(loaded.book,{}, {write:async bytes=>{
          let offset=0;while(offset<bytes.byteLength){const step=await handle.write(bytes,offset,bytes.byteLength-offset);if(!step.bytesWritten)throw new Error('파일 쓰기가 진행되지 않습니다.');offset+=step.bytesWritten;}
        }},progress);
      } finally {await handle.close();}
      // 목적 파일은 실행 전 존재하지 않아야 하며, 원본은 절대로 덮어쓰지 않습니다.
      await stat(config.roundtrip).then(()=>{throw new Error('저장 대상이 검수 중 생성되어 덮어쓰기를 거절합니다.');},error=>{if(error.code!=='ENOENT')throw error;});
      await rename(temp,config.roundtrip);report.exportMs=performance.now()-exportStarted;report.memory.afterExport=importMemory.sample('export-complete');report.export=await fingerprint(config.roundtrip);
    }
    report.sourceAfter=await sourceFingerprint(config.source);report.inputAfter=await fingerprint(input);
    report.sourceUnchanged=report.sourceBefore.hash===report.sourceAfter.hash;
    report.inputUnchanged=canonicalHash(report.inputBefore)===canonicalHash(report.inputAfter);
    report.completed=true;
  } catch(error) {if(importing){report.memory.failedImport=importMemory.sample('import-failed');report.memory.import=importMemory.report();}report.completed=false;report.error={name:error.name,message:error.message};process.exitCode=2;}
  finally {
    await Promise.all([...streams.values()].map(stream=>new Promise((done,reject)=>{if(stream.destroyed){stream.errored?reject(stream.errored):done();return;}stream.once('error',reject);stream.end(done);}))).catch(error=>{report.completed=false;report.error={name:error.name,message:error.message};process.exitCode=2;});
    report.finished=new Date().toISOString();await writeFile(join(folder,'phase.json'),JSON.stringify(report,null,2));
  }
  if(!report.sourceUnchanged || !report.inputUnchanged)process.exitCode=2;
}

async function runChild(configPath,name) {
  const args=process.execArgv.filter(a=>!a.startsWith('--inspect'));if(!args.includes('--expose-gc'))args.push('--expose-gc');
  const child=spawn(process.execPath,[...args,TOOL,'--phase',name,'--config',configPath],{stdio:'inherit'});
  const [code,signal]=await once(child,'exit');if(code!==0)throw new Error(name+' 단계 실패: '+(signal??code));
}
async function* records(path) {const input=createReadStream(path),lines=createInterface({input,crlfDelay:Infinity});try{for await(const line of lines)if(line)yield JSON.parse(line);}finally{lines.close();input.destroy();}}
async function targetedDiff(out,limit) {
  const differences=[],counts=new Map();if(limit===0)return {limitPerSheet:0,differences};
  for(const category of CATEGORIES) {
    const a=records(join(out,'original',category+'.ndjson'))[Symbol.asyncIterator](),b=records(join(out,'reimport',category+'.ndjson'))[Symbol.asyncIterator]();
    try {
      for(;;) {
        const [aa,bb]=await Promise.all([a.next(),b.next()]);if(aa.done&&bb.done)break;
        if(canonicalHash(aa.value)===canonicalHash(bb.value))continue;
        const sheet=aa.value?.sheet??bb.value?.sheet??-1,count=counts.get(sheet)??0;
        if(count<limit) {differences.push({sheet,category,before:aa.value??null,after:bb.value??null});counts.set(sheet,count+1);}
      }
    } finally {await a.return?.();await b.return?.();}
  }
  return {limitPerSheet:limit,scope:'전체 category hash 실패의 제한된 4096행 band 및 metadata record 예시; 값·수식 원문은 출력하지 않음',differences};
}

function parseCli(args) {
  const config={asyncRead:true,maxDiffs:8,source:DEFAULT_SOURCE};
  for(let i=0;i<args.length;i++) {
    const arg=args[i];
    if(arg==='--async')config.asyncRead=true;
    else if(arg==='--sync')config.asyncRead=false;
    else if(['--out','--roundtrip','--source','--max-diffs','--native-targets'].includes(arg)) {
      if(!args[i+1])throw new Error(arg+' 값이 없습니다.');
      const value=args[++i];if(arg==='--max-diffs')config.maxDiffs=Number(value);else config[{'--out':'out','--roundtrip':'roundtrip','--source':'source','--native-targets':'nativeTargets'}[arg]]=resolve(value);
    } else if(!arg.startsWith('--')&&!config.input)config.input=resolve(arg);
    else throw new Error('알 수 없는 인수: '+arg);
  }
  if(!config.input||!config.out)throw new Error('사용법: node tools/large-xlsx-fidelity.mjs INPUT --out REPORT_DIR [--roundtrip OUTPUT.xlsx] [--source SRC_DIR] [--native-targets TARGETS.json] [--async|--sync] [--max-diffs 8]');
  if(!['.xlsx','.xlsm','.xlsb'].includes(extname(config.input).toLowerCase()))throw new Error('xlsx/xlsm/xlsb만 검사할 수 있습니다.');
  if(!Number.isSafeInteger(config.maxDiffs)||config.maxDiffs<0)throw new Error('--max-diffs는 0 이상의 정수여야 합니다.');
  return config;
}
async function main(args) {
  if(args[0]==='--phase') {await phase(args[3],args[1]);return;}
  const config=parseCli(args);config.source=await realpath(config.source);config.input=await realpath(config.input);if(config.nativeTargets)config.nativeTargets=await realpath(config.nativeTargets);
  if(config.roundtrip) {
    if(!['.xlsx','.xlsm'].includes(extname(config.roundtrip).toLowerCase()))throw new Error('저장 대상은 새 xlsx/xlsm 경로여야 합니다.');
    if(config.roundtrip===config.input)throw new Error('원본 덮어쓰기는 허용하지 않습니다.');
    await stat(config.roundtrip).then(()=>{throw new Error('저장 대상이 이미 존재합니다. 새 파일 경로를 사용하세요.');},error=>{if(error.code!=='ENOENT')throw error;});
    await mkdir(dirname(config.roundtrip),{recursive:true});
  }
  await mkdir(config.out,{recursive:true});const configPath=join(config.out,'config.json');await writeFile(configPath,JSON.stringify(config,null,2),{flag:'wx'});
  try {
    await runChild(configPath,'original');
  // 자식 종료가 원본 모델의 확실한 해제 경계입니다. GC 성공 여부에 기대지 않습니다.
  if(config.roundtrip)await runChild(configPath,'reimport');
  const original=JSON.parse(await readFile(join(config.out,'original','phase.json'),'utf8'));
  const report={schemaVersion:1,mode:config.roundtrip?'roundtrip':'digest-only',original,scope:original.scope};
  if(config.roundtrip) {
    report.reimport=JSON.parse(await readFile(join(config.out,'reimport','phase.json'),'utf8'));
    report.comparison=compareDigests(original.digest,report.reimport.digest);
    report.targetedDiff=report.comparison.equal?{limitPerSheet:config.maxDiffs,differences:[]}:await targetedDiff(config.out,config.maxDiffs);
    report.sameSourceBetweenPhases=original.sourceBefore.hash===report.reimport.sourceBefore.hash;
    report.sameAuditTool=original.auditToolHash===report.reimport.auditToolHash && original.nativeSamplerHash===report.reimport.nativeSamplerHash;
    if(config.nativeTargets)report.sameNativeTargets=canonicalHash(original.nativeTargetsFingerprint)===canonicalHash(report.reimport.nativeTargetsFingerprint);
    report.pass=report.comparison.equal && report.sameSourceBetweenPhases && report.sameAuditTool && report.sameNativeTargets!==false && original.sourceUnchanged && original.inputUnchanged && report.reimport.sourceUnchanged && report.reimport.inputUnchanged;
  } else report.pass=original.completed && original.sourceUnchanged && original.inputUnchanged;
  await writeFile(join(config.out,'report.json'),JSON.stringify(report,null,2));
  console.log(JSON.stringify({mode:report.mode,pass:report.pass,sheets:original.digest.sheets.length,hash:original.digest.hash,differenceCategories:report.comparison?.differences.length??0}));
  if(!report.pass)process.exitCode=1;
  } catch(error) {
    const report={schemaVersion:1,mode:config.roundtrip?'roundtrip':'digest-only',pass:false,error:{name:error.name,message:error.message},scope:'검수 미완료; Excel/렌더 합격 아님'};
    for(const name of ['original','reimport'])try{report[name]=JSON.parse(await readFile(join(config.out,name,'phase.json'),'utf8'));}catch{}
    await writeFile(join(config.out,'report.json'),JSON.stringify(report,null,2));throw error;
  }
}

if(process.argv[1] && resolve(process.argv[1])===TOOL)main(process.argv.slice(2)).catch(error=>{console.error(error.message);process.exitCode=2;});