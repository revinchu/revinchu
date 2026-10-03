// XML/BIFF에서 모델 복원 전 보유하는 반복 리터럴도 공유합니다.
// 수식과 위치별 속성은 제외하며, 후속 주석/링크 적용은 셀을 복제해 교체합니다.
export function createImportedLiteralMemo({limit=10000,perStyleLimit=512}={}) {
  const styles=new Map();let size=0;
  return cell=>{
    if(typeof cell?.raw!=='string'||!cell.raw||cell.raw.startsWith('='))return cell;
    for(const key in cell)if(key!=='raw'&&key!=='style'&&key!=='inputType')return cell;
    let group=styles.get(cell.style),values=group?.types.get(cell.inputType);
    const found=values?.get(cell.raw);if(found)return found;
    if(size>=limit||(group?.size??0)>=perStyleLimit)return cell;
    if(!group){group={types:new Map(),size:0};styles.set(cell.style,group);}
    if(!values){values=new Map();group.types.set(cell.inputType,values);}
    values.set(cell.raw,Object.freeze(cell));group.size++;size++;return cell;
  };
}
