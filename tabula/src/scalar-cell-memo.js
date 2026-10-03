// 가져온 리터럴 셀의 완전히 같은 상태만 공유합니다. 메모는 복원 작업이 끝나면 버립니다.
const NUMERIC_RAW = /^-?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;
export function createScalarCellMemo({limit=10000,perStyleLimit=512}={}) {
 const styles=new Map();let size=0;
 return cell=>{
  if(!cell||typeof cell.raw!=='string'||cell.raw==='')return cell;
  const type=typeof cell.v;
  if(type!=='string'&&type!=='boolean'&&!(type==='number'&&Number.isFinite(cell.v)&&NUMERIC_RAW.test(cell.raw)&&Object.is(Number(cell.raw),cell.v)))return cell;
  // 날짜 문자열은 날짜 기준 변경에서 raw가 직접 바뀝니다. 수식/오류/메모 등도 공유하지 않습니다.
  for(const key in cell)if(key!=='raw'&&key!=='style'&&key!=='v'&&key!=='inputType')return cell;
  const style=cell.style, inputType=cell.inputType;
  let group=styles.get(style),values=group?.types.get(inputType);
  const hit=values?.get(cell.raw);
  if(hit)return Object.is(hit.v,cell.v)?hit:cell;
  if(size>=limit||(group?.size??0)>=perStyleLimit)return cell;
  if(!group){group={size:0,types:new Map()};styles.set(style,group);}
  if(!values){values=new Map();group.types.set(inputType,values);}
  values.set(cell.raw,Object.freeze(cell));group.size++;size++;return cell;
 };
}
