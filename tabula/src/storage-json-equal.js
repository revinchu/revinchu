// Compare the JSON data already used by workbook metadata without retaining
// a second, potentially huge JSON string. Object key order follows stringify;
// undefined object properties are omitted, array holes and non-finite numbers
// become null. Workbook metadata is structured-cloneable data, not custom
// serializers or objects with side-effecting accessors.
const omitted = value => value === undefined || typeof value === 'function' || typeof value === 'symbol';
function primitive(value, array) {
  if (value instanceof Date) return value.toJSON();
  if (value instanceof Number || value instanceof String || value instanceof Boolean) value = value.valueOf();
  if (typeof value === 'number' && !Number.isFinite(value)) return null;
  if(typeof value==='bigint'||(value&&typeof value==='object'&&Object.getPrototypeOf(value)===BigInt.prototype))throw new TypeError('BigInt workbook metadata');
  return omitted(value) ? (array?null:undefined) : value;
}
function* keys(value, omit) {
  for (const key in value) if (Object.hasOwn(value,key) && !omit?.has(key) && !omitted(value[key])) yield key;
}
export function storedJsonEqual(left, right, { omitRoot = [] } = {}) {
  const leftPath = new Set(), rightPath = new Set(), omit = omitRoot.length ? new Set(omitRoot) : null;
  const equal = (a,b,array=false,root=false) => {
    a=primitive(a,array);b=primitive(b,array);
    if(a===null||b===null||typeof a!=='object'||typeof b!=='object')return a===b;
    if(Array.isArray(a)!==Array.isArray(b))return false;
    if(leftPath.has(a)||rightPath.has(b))throw new TypeError('Circular workbook metadata');
    leftPath.add(a);rightPath.add(b);
    try {
      if(Array.isArray(a)) {
        if(a.length!==b.length)return false;
        for(let i=0;i<a.length;i++)if(!equal(a[i],b[i],true))return false;
        return true;
      }
      const ak=keys(a,root?omit:null),bk=keys(b,root?omit:null);
      while(true) {
        const x=ak.next(),y=bk.next();
        if(x.done||y.done)return x.done===y.done;
        if(x.value!==y.value||!equal(a[x.value],b[y.value]))return false;
      }
    } finally {leftPath.delete(a);rightPath.delete(b);}
  };
  return equal(left,right,false,true);
}
/** Read only the captured sheet schema; never clone its images/styles again. */
export function currentSheetMetadata(sheet, captured) {
  const view={};
  for(const key in captured)if(Object.hasOwn(captured,key)&&key!=='fileValues')view[key]=sheet[key];
  if(sheet.fileValues)view.fileValues=true;
  return view;
}
