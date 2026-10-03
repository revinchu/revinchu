// Choose one small, literal numeric source cell and edit only the in-memory audit copy.
export async function editPivotSourceForExport(page,{id}) {
 const result=await page.evaluate(()=>{
  const w=tabula.wb(),valid=w.snapshotData(),defs=w.sheets.flatMap(s=>[s.pivot,...(s.pivotsExtra??[])]).filter(Boolean);
  let chosen=null;
  for(const def of defs){
   if(!def.snapshotId||!valid[def.snapshotId])continue;
   const source=w.pivotSnapshotSource(def);if(!source)continue;
   const [r1,c1,r2,c2]=JSON.parse(source.key);
   if(![r1,c1,r2,c2].every(Number.isInteger))continue;
   for(let r=r1+1;r<=Math.min(r2,r1+64)&&!chosen;r++)for(let c=c1;c<=Math.min(c2,c1+63);c++){
    const cell=w.getCell(source.si,r,c);
    if(!cell||cell.formula||cell.image||w.mergeAt(source.si,r,c))continue;
    const value=w.getValue(source.si,r,c),style=w.styleAt(source.si,r,c),fmt=style.numFmt;
    // 날짜 일련번호/General의 의미 불명 숫자는 피하고 명시된 숫자 서식의 정수 지표만 고릅니다.
    const code=String(style.code||'').replace(/"[^"]*"/g,'').replace(/\\./g,'').replace(/\[\$[^\]]*\]/g,'');
    const numericFormat=['number','comma','currency','accounting','scientific'].includes(fmt)||fmt==='custom'&&/[0#?]/.test(code)&&!/[ymdhs]/i.test(code);
    if(!numericFormat||!/^[-+]?\d+$/.test(cell.raw)||!Number.isSafeInteger(value)||value<0||!Number.isSafeInteger(value+1))continue;
    chosen={si:source.si,r,c,value,snapshotId:def.snapshotId};break;
   }
   if(chosen)break;
  }
  if(!chosen)throw Error('유효 피벗 캐시의 제한된 원본 범위에서 숫자 리터럴 셀을 찾지 못했습니다.');
  const {si,r,c,value,snapshotId}=chosen,beforeVersion=w.sourceVersion(si),beforeUndo=w.undoStack.length;
  w.transact(()=>w.setInput(si,r,c,String(value+1)));
  const after=w.snapshotData();
  const result={sheet:si,row:r,column:c,delta:1,literalNumeric:true,explicitNumericFormat:true,integerMeasurement:true,incrementApplied:w.getValue(si,r,c)===value+1,sourceVersionBefore:beforeVersion,sourceVersionAfter:w.sourceVersion(si),validSnapshotsBefore:Object.keys(valid).length,validSnapshotsAfter:Object.keys(after).length,editedSourceCacheInvalidated:!after[snapshotId],undoEntryAdded:w.undoStack.length===beforeUndo+1};
  return {...result,ok:result.incrementApplied&&result.editedSourceCacheInvalidated&&result.sourceVersionAfter>beforeVersion&&result.undoEntryAdded};
 });
 if(!result.ok){const error=Error('피벗 원본 셀 변경 또는 저장 캐시 무효화 검사가 실패했습니다.');error.audit=result;throw error;}
 // 앱의 지연 피벗 갱신/화면 반영을 마친 다음 저장 스냅샷을 시작합니다.
 await page.waitForTimeout(750);
 return {id,...result};
}
