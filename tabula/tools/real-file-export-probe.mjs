// 실파일 검수에서 브라우저 다운로드 경로로 새 사본만 저장한다. 원본 핸들을 사용하지 않는다.
import { stat } from 'node:fs/promises';
export async function exportAuditXlsx(page,{id,out,timeout=600000,fileStem=id+'-roundtrip'}) {
 if(!/^[A-Za-z0-9_-]+$/.test(fileStem))throw Error('검수 사본 이름에 허용되지 않은 문자가 있습니다.');
 const name=fileStem,file=out+'/'+name+'.xlsx',started=Date.now();
 const result={path:file,route:'browser-download',compatibilityConfirmation:false};
 let cdp=null;
 if(process.env.WIXEL_AUDIT_EXPORT_TRACE==='1'){
  cdp=await page.context().newCDPSession(page);result.exceptionTrace=[];
  cdp.on('Debugger.paused',event=>{const description=String(event.data?.description||'');if(result.exceptionTrace.length<30||description.includes('Invalid string length'))result.exceptionTrace.push({reason:event.reason,className:event.data?.className,invalidStringLength:description.includes('Invalid string length'),frames:event.callFrames.slice(0,16).map(f=>({function:f.functionName,url:f.url,line:f.location.lineNumber+1,column:f.location.columnNumber+1}))});void cdp.send('Debugger.resume').catch(()=>{});});
  await cdp.send('Debugger.enable');await cdp.send('Debugger.setPauseOnExceptions',{state:'all'});
 }
 let pendingDownload=null,downloadError=null;
 const onDownload=download=>{pendingDownload=download.saveAs(file).catch(error=>{downloadError=error;});};page.once('download',onDownload);
 try{
  result.liveCacheBefore=await page.evaluate(()=>{
   const book=tabula.wb(),map=book.pivotSnapshots,valid=Object.keys(book.snapshotData?.()??{}).sort();
   const entries=[...(map??[])].map(([id,entry])=>({id,entry,rows:entry.rows,sourceSheet:entry.sourceSheet,sourceKey:entry.sourceKey,version:entry.ver}));
   const versions=book.sheets.map((sheet,i)=>book.sourceVersion(i));
   const arrays=[];book.sheets.forEach((sheet,si)=>sheet.cells.forEachStoredRC((cell,r,c)=>{if(cell.cachedArray)arrays.push({si,r,c,cell,data:cell.cachedArray,dirty:!!cell.dirty,count:cell.cachedArray.values.length/3});}));
   window.__auditExportCacheBefore={book,map,entries,valid,versions,arrays,version:book.version};
   return {arrayCaches:{anchors:arrays.length,dirty:arrays.filter(x=>x.dirty).length,storedValues:arrays.reduce((n,x)=>n+x.count,0)},snapshots:map?.size??0,validSnapshots:valid.length,sourceVersionCount:versions.length,workbookVersion:book.version};
  });
  await page.evaluate(name=>{
   Object.defineProperty(window,'showSaveFilePicker',{value:undefined,configurable:true});
   window.__auditExportState={pending:true};
   tabula.exportXlsx(name,'xlsx').then(value=>{window.__auditExportState={pending:false,accepted:!!value};},error=>{window.__auditExportState={pending:false,error:error.message};});
  },name);
  while(Date.now()-started<timeout){
   if(pendingDownload){
    await pendingDownload;if(downloadError)throw downloadError;
    await page.waitForFunction(()=>!window.__auditExportState?.pending,null,{timeout:Math.max(1000,timeout-(Date.now()-started))});
    const completed=await page.evaluate(()=>window.__auditExportState);
    if(!completed.accepted)throw Error(completed.error||'다운로드 후 앱 저장 명령이 완료되지 않았습니다.');
    result.liveCachePreservation=await page.evaluate(()=>{
     const before=window.__auditExportCacheBefore,book=tabula.wb(),map=book.pivotSnapshots,valid=Object.keys(book.snapshotData?.()??{}).sort();
     const sameArrayCaches=before.arrays.every(a=>{const c=book.sheets[a.si]?.cells.getRC(a.r,a.c);return c===a.cell&&c.cachedArray===a.data&&!!c.dirty===a.dirty&&c.cachedArray.values.length/3===a.count;});
     const checks={sameArrayCaches,sameBook:book===before.book,sameMap:map===before.map,sameCount:(map?.size??0)===before.entries.length,sameEntries:before.entries.every(e=>map?.get(e.id)===e.entry),sameRows:before.entries.every(e=>map?.get(e.id)?.rows===e.rows),sameMetadata:before.entries.every(e=>{const after=map?.get(e.id);return after?.sourceSheet===e.sourceSheet&&after?.sourceKey===e.sourceKey&&after?.ver===e.version;}),sameValidSnapshots:valid.length===before.valid.length&&valid.every((id,i)=>id===before.valid[i]),sameSourceVersions:book.sheets.length===before.versions.length&&before.versions.every((v,i)=>book.sourceVersion(i)===v)};
     return {...checks,ok:Object.values(checks).every(Boolean),snapshots:map?.size??0,validSnapshots:valid.length,workbookVersionBefore:before.version,workbookVersionAfter:book.version};
    });
    if(!result.liveCachePreservation.ok)throw Error('저장 과정에서 원본 피벗 캐시의 정체성·유효성·버전이 바뀌었습니다.');
    const info=await stat(file);return{...result,bytes:info.size,ms:Date.now()-started,ok:true};
   }
   const dialog=page.locator('#dialogLayer .dialog').last();
   if(await dialog.count()){
    const title=(await dialog.locator('.dialog-title').textContent())?.trim();
    if(title==='Excel 호환성 확인'){result.compatibilityConfirmation=true;await dialog.getByRole('button',{name:'확인',exact:true}).click({timeout});}
    else if(title==='파일로 저장')await dialog.getByRole('button',{name:'다운로드',exact:true}).click({timeout});
    else throw Error('내보내기 안내: '+(await dialog.locator('.dialog-body').textContent()).slice(0,500));
   }
   const state=await page.evaluate(()=>window.__auditExportState);
   if(!state.pending&&!state.accepted&&!pendingDownload){const message=await page.locator('#dialogLayer .dialog-body').last().textContent({timeout:1000}).catch(()=>'');throw Error(state.error||message||'다운로드 생성 전 저장이 종료되었습니다.');}
   await page.waitForTimeout(100);
  }
  throw Error('표준 XLSX 다운로드 제한 시간을 초과했습니다.');
 }catch(error){result.ok=false;result.ms=Date.now()-started;result.error=error.message;error.audit=result;throw error;}finally{await page.evaluate(()=>{delete window.__auditExportCacheBefore;}).catch(()=>{});page.off('download',onDownload);if(cdp){await cdp.send('Debugger.setPauseOnExceptions',{state:'none'}).catch(()=>{});await cdp.detach().catch(()=>{});}}
}
