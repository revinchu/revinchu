// JSON.stringify-compatible fragments, including the existing typed-array
// replacer's __ta/base64 representation, without one giant encoded array.
function* base64Parts(bytes) {
  const chunk=24576; // divisible by 3, so only the final part has padding
  for(let i=0;i<bytes.length;i+=chunk) {
    const part=bytes.subarray(i,i+chunk);let text='';
    for(let j=0;j<part.length;j+=8192)text+=String.fromCharCode(...part.subarray(j,j+8192));
    yield btoa(text);
  }
}
export function* jsonValueParts(value, typed=false) {
  if(value&&typeof value==='object'&&typeof value.toJSON==='function'){yield* jsonValueParts(value.toJSON(),typed);return;}
  if(value===null||typeof value!=='object'){yield JSON.stringify(value)??'null';return;}
  if(typed&&ArrayBuffer.isView(value)&&!(value instanceof DataView)) {
    yield '{"__ta":'+JSON.stringify(value.constructor.name)+',"b64":"';
    yield* base64Parts(new Uint8Array(value.buffer,value.byteOffset,value.byteLength));yield '"}';return;
  }
  if(Array.isArray(value)) {
    yield '[';for(let i=0;i<value.length;i++){if(i)yield ',';yield* jsonValueParts(value[i],typed);}yield ']';return;
  }
  yield '{';let first=true;
  for(const key in value)if(Object.hasOwn(value,key)&&value[key]!==undefined&&typeof value[key]!=='function') {
    if(!first)yield ',';first=false;yield JSON.stringify(key)+':';yield* jsonValueParts(value[key],typed);
  }
  yield '}';
}
