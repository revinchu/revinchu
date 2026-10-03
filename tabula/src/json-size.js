// JSON 문자열을 만들지 않고 저장 할당의 상한을 추정합니다. 호출마다 공유 객체 메모를 분리합니다.
/** JSON에 필요한 문자열 길이. 큰 문자열은 살펴보기 전에 종료하고 이스케이프도 셉니다. */
export function jsonStringSize(value, limit = 2e6) {
  const over = limit + 1;
  let size = value.length + 2;
  if (size > limit || !/["\\\u0000-\u001f\ud800-\udfff]/.test(value)) return Math.min(over, size);
  for (let i=0;i<value.length;i++) {
    const c=value.charCodeAt(i);
    if (c===34 || c===92) size++;
    else if (c<32) size += [8,9,10,12,13].includes(c) ? 1 : 5;
    else if (c>=0xd800 && c<=0xdbff && value.charCodeAt(i+1)>=0xdc00 && value.charCodeAt(i+1)<=0xdfff) i++;
    else if (c>=0xd800 && c<=0xdfff) size+=5;
    if (size>limit) return over;
  }
  return size;
}

export function createJsonSizer(limit = 2e6) {
  // 공유 서식도 저장되는 위치마다 크기를 더하되 같은 객체를 반복해서 훑지는 않습니다.
  const over=limit+1, memo=new WeakMap(),parents=new Set();
  const size=value=>{
    if(value==null)return 4;
    if(typeof value==='string')return jsonStringSize(value,limit);
    if(typeof value==='number')return 24;
    if(typeof value==='boolean')return 5;
    if(typeof value!=='object')return 0;
    if(memo.has(value))return memo.get(value);
    if(parents.has(value)||parents.size>=64)return over;
    if(ArrayBuffer.isView(value))return Math.min(over,value.length*40+2);
    parents.add(value);let total=2;
    try {
      if(Array.isArray(value)) {
        for(const item of value){total+=(size(item)||4)+1;if(total>limit)break;}
      } else {
        for(const key in value)if(Object.hasOwn(value,key)&&value[key]!==undefined&&typeof value[key]!=='function') {
          total+=jsonStringSize(key,limit)+1+size(value[key])+1;if(total>limit)break;
        }
      }
    } finally {parents.delete(value);}
    total=Math.min(over,total);memo.set(value,total);return total;
  };
  const fields=(object,keys)=>{
    let total=2;
    for(const key of keys)if(object?.[key]!==undefined){total+=key.length+4+size(object[key]);if(total>limit)return over;}
    return total;
  };
  return {size,fields};
}
