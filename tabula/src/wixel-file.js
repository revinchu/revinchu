import { CellMap } from './cellmap.js';
import { createStoredStyleMemo } from './cell-storage.js';
import { createJsonSizer } from './json-size.js';

// Framed WIXEL v2: each JSON record is bounded independently. The complete
// workbook never becomes one string or a second object/cell dictionary.
export const WIXEL_FILE_MAGIC = '{"format":"wixel-stream","version":2}';
const TEXT_PART = 32768, BYTE_PART = 65536;
const STYLE_LIMIT = 4096, STYLE_CHAR_LIMIT = 1 << 20, STYLE_ENTRY_LIMIT = 65536;
const TYPES = { Float64Array, Float32Array, Int32Array, Uint32Array, Int16Array, Uint16Array, Int8Array, Uint8Array, Uint8ClampedArray };
const fail = () => new Error('WIXEL 파일이 완전하지 않거나 저장 형식이 올바르지 않습니다.');
const check = options => {
  if (options.signal?.aborted || options.isCurrent?.() === false) throw Object.assign(new Error('저장 또는 열기 중 문서가 바뀌어 작업을 중단했습니다.'), { code: 'WIXEL_FILE_ABORT' });
};
function b64(bytes) {
  let out = ''; for (let i=0;i<bytes.length;i+=8192) out += String.fromCharCode(...bytes.subarray(i,i+8192));
  return btoa(out);
}
function unb64(text) { const raw=atob(text), out=new Uint8Array(raw.length); for(let i=0;i<raw.length;i++)out[i]=raw.charCodeAt(i); return out; }
function* valueRecords(key, value) {
  if (value === undefined || typeof value === 'function') return;
  if(value&&typeof value==='object'&&typeof value.toJSON==='function'){yield* valueRecords(key,value.toJSON(key));return;}
  if (typeof value === 'string' && value.length > TEXT_PART) {
    yield ['s',key,value.length];
    for(let i=0;i<value.length;i+=TEXT_PART)yield ['p',value.slice(i,i+TEXT_PART)];
    yield ['e']; return;
  }
  if (ArrayBuffer.isView(value) && TYPES[value.constructor.name]) {
    yield ['t',key,value.constructor.name,value.length];
    const bytes=new Uint8Array(value.buffer,value.byteOffset,value.byteLength);
    for(let i=0;i<bytes.length;i+=BYTE_PART)yield ['b',b64(bytes.subarray(i,i+BYTE_PART))];
    yield ['e'];return;
  }
  if (value === null || typeof value !== 'object') { yield ['v',key,value]; return; }
  if (Array.isArray(value)) {
    yield ['a',key]; let small=[],size=0;
    for(const item of value) {
      if(item===null || ['number','boolean','undefined'].includes(typeof item) || typeof item==='string'&&item.length<=4096) {
        small.push(item??null);size+=typeof item==='string'?item.length*6+8:32;
        if(small.length<1024&&size<TEXT_PART)continue;
      } else {
        if(small.length){yield ['l',small];small=[];size=0;}
        yield* valueRecords(null,item);continue;
      }
      yield ['l',small];small=[];size=0;
    }
    if(small.length)yield ['l',small];
    yield ['e'];return;
  }
  yield ['o',key];
  for(const prop of Object.keys(value))yield* valueRecords(prop,value[prop]);
  yield ['e'];
}
// A file-local bounded dictionary removes repeated XF style JSON from millions
// of cells. Both source-identity and equivalent-style caches are bounded; novel
// styles beyond the budget remain inline and preserve their complete contents.
function styleDictionary() {
  const identities=new Map(),equivalents=new Map();let chars=0;
  return style=>{
    if(!style||typeof style!=='object'||Array.isArray(style))return null;
    const hit=identities.get(style);if(hit!==undefined)return hit;
    if(equivalents.size>=STYLE_LIMIT||chars>=STYLE_CHAR_LIMIT)return null;
    if(createJsonSizer(STYLE_ENTRY_LIMIT).size(style)>STYLE_ENTRY_LIMIT)return null;
    const text=JSON.stringify(style);if(!text||text.length>STYLE_ENTRY_LIMIT)return null;
    let id=equivalents.get(text);
    if(id===undefined){if(chars+text.length>STYLE_CHAR_LIMIT)return null;id=equivalents.size;equivalents.set(text,id);chars+=text.length;}
    if(identities.size<STYLE_LIMIT*4)identities.set(style,id);
    return id;
  };
}
function* fileRecords(book, metadata) {
  const styleRef=styleDictionary();let definedStyles=0;
  yield ['o',null];
  for(const key of Object.keys(metadata))if(key!=='workbook')yield* valueRecords(key,metadata[key]);
  yield ['o','workbook'];yield ['v','version',1];
  for(const [key,value] of Object.entries(book.bookMeta(false)))yield* valueRecords(key,value);
  const snapshots=book.snapshotData();
  if(Object.keys(snapshots).length)yield* valueRecords('pivotSnapshots',snapshots);
  yield ['a','sheets'];
  for(let si=0;si<book.sheets.length;si++) {
    const sheet=book.sheets[si];yield ['o',null];
    for(const [key,value] of Object.entries(book.sheetMeta(si)))yield* valueRecords(key,value);
    yield* valueRecords('blocks',sheet.blocks??[]);
    yield ['c','cells'];
    for(const chunk of book.cellRunChunks(si,2048,{shareStyle:true})) {
      for(const entry of chunk) {
        const data=entry[3],style=data?.style,ref=styleRef(style);
        if(ref===null)continue;
        if(ref===definedStyles){yield ['d',ref,style];definedStyles++;}
        // Only this fresh storage record is changed, never the live cell/style.
        data.style=undefined;entry.push(ref);
      }
      // One exceptional cell (large picture/formula result) is framed too.
      if(chunk.length===1) { const [r,c,count,data,id]=chunk[0];yield* valueRecords(id===undefined?[r,c,count]:[r,c,count,id],data); }
      else yield ['r',chunk];
    }
    yield ['e'];yield ['e'];yield ['progress',(si+1)/book.sheets.length];
  }
  yield ['e'];yield ['e'];yield ['e'];
}

function snapshotCheck(book,options) {
  const version=book.version,sheets=[...book.sheets],edits=sheets.map(s=>s._ev??0);
  return ()=>{check(options);if(book.version!==version||book.sheets.length!==sheets.length||sheets.some((s,i)=>book.sheets[i]!==s||(s._ev??0)!==edits[i]))throw Object.assign(new Error('저장 중 문서가 편집되었습니다. 편집이 끝난 뒤 다시 저장하세요.'),{code:'WIXEL_FILE_ABORT'});};
}
/** Lossless .wixel byte stream: direct file writes keep only bounded buffers. */
export function createWixelFileStream(book, metadata = {}, options = {}) {
  const current=snapshotCheck(book,options);
  current();const records=fileRecords(book,structuredClone(metadata)), encoder=new TextEncoder();let count=0,started=false,done=false,last=performance.now(),failure=null;
  const stream=new ReadableStream({
    async pull(controller) {
      try {
        current();
        // Sheet metadata can contain hundreds of thousands of tiny records.
        // Batch the same newline-framed bytes to avoid one native stream buffer
        // and promise chain for every individual property/brace.
        const parts=[];let chars=0;const began=performance.now();
        if(!started){started=true;parts.push(WIXEL_FILE_MAGIC+'\n');chars+=WIXEL_FILE_MAGIC.length+1;}
        do {
          const step=records.next();
          if(step.done){done=true;parts.push(JSON.stringify(['finish',count])+'\n');break;}
          if(step.value[0]==='progress')options.onProgress?.(step.value[1]);
          const text=JSON.stringify(step.value)+'\n';parts.push(text);chars+=text.length;count++;
        } while(chars<65536&&parts.length<2048&&performance.now()-began<4);
        controller.enqueue(encoder.encode(parts.join('')));
        if(done){controller.close();return;}
        if(performance.now()-last>=4){await new Promise(resolve=>setTimeout(resolve,0));last=performance.now();current();}
      }catch(error){failure=error;records.return?.();controller.error(error);}
    },cancel(){records.return?.();}
  });
  const compressed=options.gzip!==false&&typeof CompressionStream==='function';
  const reader=(compressed?stream.pipeThrough(new CompressionStream('gzip')):stream).getReader();
  let released=false;
  const release=()=>{if(!released){reader.releaseLock();released=true;}};
  return new ReadableStream({
    async pull(controller) {
      try {
        current();const next=await reader.read();current();
        if(next.done){if(!done)throw fail();release();controller.close();}
        else controller.enqueue(next.value);
      } catch(error) {
        const reason=failure??error;
        try {await reader.cancel(reason);} catch { /* Preserve the source failure. */ }
        release();controller.error(reason);
      }
    },
    async cancel(reason) {try{await reader.cancel(reason);}finally{release();}}
  },{highWaterMark:0});
}

/** Blob compatibility for browser downloads and existing online storage. */
export async function writeWixelFile(book, metadata = {}, options = {}) {
  const current=snapshotCheck(book,options);
  const blob=await new Response(createWixelFileStream(book,metadata,options)).blob();
  current();return new Blob([blob],{type:'application/x-wixel'});
}

// Blob.stream() chunk sizes are browser-specific. In particular, one large
// compressed chunk can make a native decompressor allocate the entire expanded
// workbook before downstream backpressure applies. Bound its input explicitly.
function blobByteStream(blob) {
  if(!Number.isSafeInteger(blob.size)||blob.size<0)throw fail();
  let offset=0;
  return new ReadableStream({
    async pull(controller) {
      if(offset>=blob.size){controller.close();return;}
      const end=Math.min(blob.size,offset+BYTE_PART);
      try {const bytes=await blob.slice(offset,end).arrayBuffer();offset=end;controller.enqueue(new Uint8Array(bytes));}
      catch(error){controller.error(error);}
    }
  });
}
/** Bounded UTF-8 decoder, independent of the browser's stream chunk boundaries. */
async function* linesOf(stream,options) {
  const reader=stream.getReader(),decoder=new TextDecoder('utf-8',{fatal:true});let pending='',last=performance.now();
  try {
    for(;;) {
      check(options);const next=await reader.read(),bytes=next.value??new Uint8Array();
      for(let offset=0;offset<bytes.length||next.done&&offset===0;offset+=BYTE_PART) {
        check(options);
        const text=pending+decoder.decode(bytes.subarray(offset,offset+BYTE_PART),{stream:!next.done});let from=0,at;
        while((at=text.indexOf('\n',from))>=0){yield text.slice(from,at);from=at+1;}
        pending=text.slice(from);
        if(pending.length>16*1024*1024)throw new Error('WIXEL 파일의 저장 조각이 손상되었거나 너무 큽니다.');
        if(performance.now()-last>25){await new Promise(resolve=>setTimeout(resolve,0));last=performance.now();}
      }
      if(next.done){if(pending)throw fail();break;}
    }
  }finally{await reader.cancel().catch(()=>{});reader.releaseLock();}
}
/** Read v2 incrementally; previous JSON .wixel/.json files remain readable. */
export async function readWixelFile(blob,options={}) {
  check(options);const signature=new Uint8Array(await blob.slice(0,2).arrayBuffer());
  const gzip=signature[0]===31&&signature[1]===139;
  if(gzip&&typeof DecompressionStream!=='function')throw new Error('압축된 WIXEL 파일을 열려면 최신 브라우저를 사용하세요.');
  if(!gzip) {
    const prefix=await blob.slice(0,256).text();
    if(!prefix.startsWith(WIXEL_FILE_MAGIC+'\n')) {const data=JSON.parse(await blob.text());check(options);return data;}
  }
  let stream=blobByteStream(blob);if(gzip)stream=stream.pipeThrough(new DecompressionStream('gzip'));
  if(gzip) {
    // Existing #view links contain gzip JSON rather than v2 records. Peek only
    // the signature, then replay the same bytes; never stringify a v2 payload.
    const reader=stream.getReader(),prefix=[];let size=0,ended=false;
    try {
      while(size<WIXEL_FILE_MAGIC.length+1) {
        check(options);const next=await reader.read();if(next.done){ended=true;break;}
        prefix.push(next.value);size+=next.value.length;
      }
    } catch(error) { await reader.cancel(error).catch(()=>{});reader.releaseLock();throw error; }
    const head=new Uint8Array(Math.min(size,WIXEL_FILE_MAGIC.length+1));let at=0;
    for(const bytes of prefix){const part=bytes.subarray(0,head.length-at);head.set(part,at);at+=part.length;if(at===head.length)break;}
    stream=new ReadableStream({start(controller){for(const bytes of prefix)controller.enqueue(bytes);},async pull(controller){try{if(ended){controller.close();reader.releaseLock();return;}const next=await reader.read();if(next.done){ended=true;controller.close();reader.releaseLock();}else controller.enqueue(next.value);}catch(error){controller.error(error);}},cancel(reason){return reader.cancel(reason);}});
    if(new TextDecoder().decode(head)!==WIXEL_FILE_MAGIC+'\n') {
      const data=JSON.parse(await new Response(stream).text());check(options);return data;
    }
  }
  const stack=[],styles=[];let root,count=0,header=false,finished=false,styleChars=0;
  const attach=(key,value)=>{
    const parent=stack.at(-1);
    if(!parent){if(root!==undefined)throw fail();root=value;return;}
    if(parent.kind==='c') {
      if(!Array.isArray(key)||![3,4].includes(key.length))throw fail();
      const [r,c,n,id]=key;if(!Number.isSafeInteger(r)||r<0||!Number.isSafeInteger(c)||c<0||!Number.isSafeInteger(n)||n<1)throw fail();
      if(key.length===4){if(!Number.isSafeInteger(id)||id<0||id>=styles.length||!value||typeof value!=='object'||value.style!==undefined)throw fail();value.style=styles[id];}
      else if(value?.style)value.style=parent.shareStyle(value.style);
      if(n===1)parent.value.setRC(r,c,value);else parent.value.setRunRC(r,c,n,value);return;
    }
    if(parent.kind==='a'){parent.value.push(value);return;}
    if(parent.kind!=='o'||typeof key!=='string'||Object.hasOwn(parent.value,key))throw fail();
    Object.defineProperty(parent.value,key,{value,writable:true,configurable:true,enumerable:true});
  };
  for await(const line of linesOf(stream,options)) {
    if(!header){if(line!==WIXEL_FILE_MAGIC)throw fail();header=true;continue;}
    if(finished)throw fail();
    let record;try{record=JSON.parse(line);}catch{throw fail();}
    if(!Array.isArray(record))throw fail();const [op,key,arg,len]=record,parent=stack.at(-1);
    if(op==='finish'){if(stack.length||key!==count||!root?.workbook?.sheets?.length)throw fail();finished=true;continue;}
    count++;
    if(op==='progress'){options.onProgress?.(key);continue;}
    if(op==='d'){
      if(parent?.kind!=='c'||key!==styles.length||styles.length>=STYLE_LIMIT||!arg||typeof arg!=='object'||Array.isArray(arg)||line.length>STYLE_ENTRY_LIMIT+32)throw fail();
      styleChars+=line.length;if(styleChars>STYLE_CHAR_LIMIT+STYLE_LIMIT*32)throw fail();
      styles.push(arg);continue;
    }
    if(op==='o'||op==='a'||op==='c') {
      const value=op==='a'?[]:op==='c'?new CellMap():{};
      // Attach complete values on close so cell styles can be shared then.
      stack.push({kind:op,key,value,...(op==='c'?{shareStyle:createStoredStyleMemo()}: {})});
    } else if(op==='s'){if(!Number.isSafeInteger(arg)||arg<0)throw fail();stack.push({kind:op,key,value:[],length:arg,at:0});}
    else if(op==='t'){
      const Type=TYPES[arg];if(!Type||!Number.isSafeInteger(len)||len<0||len*Type.BYTES_PER_ELEMENT>2147483647)throw fail();
      const value=new Type(len);stack.push({kind:op,key,value,bytes:new Uint8Array(value.buffer),at:0});
    } else if(op==='e'){
      if(!parent)throw fail();stack.pop();
      if(parent.kind==='s'){if(parent.at!==parent.length)throw fail();parent.value=parent.value.join('');}
      if(parent.kind==='t'&&parent.at!==parent.bytes.length)throw fail();
      attach(parent.key,parent.value);
    } else if(op==='v')attach(key,arg);
    else if(op==='l'){if(parent?.kind!=='a'||!Array.isArray(key))throw fail();for(const value of key)parent.value.push(value);}
    else if(op==='p'){if(parent?.kind!=='s'||typeof key!=='string'||parent.at+key.length>parent.length)throw fail();parent.value.push(key);parent.at+=key.length;}
    else if(op==='b'){if(parent?.kind!=='t'||typeof key!=='string')throw fail();const bytes=unb64(key);if(parent.at+bytes.length>parent.bytes.length)throw fail();parent.bytes.set(bytes,parent.at);parent.at+=bytes.length;}
    else if(op==='r'){if(parent?.kind!=='c'||!Array.isArray(key))throw fail();for(const entry of key){if(!Array.isArray(entry)||![4,5].includes(entry.length))throw fail();const [r,c,n,data,id]=entry;attach(entry.length===5?[r,c,n,id]:[r,c,n],data);}}
    else throw fail();
    if(stack.length>128)throw fail();
  }
  check(options);if(!finished)throw fail();return root;
}
