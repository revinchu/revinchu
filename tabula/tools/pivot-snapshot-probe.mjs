// Read-only instrumentation for imported pivot snapshot identity/version checks.
export async function installPivotSnapshotProbe(page, events) {
  page.on('console', message => {
    const text = message.text();
    if (!text.startsWith('WIXEL_PIVOT_SNAPSHOT_PROBE ')) return;
    try { const event = JSON.parse(text.slice(27)); events.push(event); console.log(JSON.stringify({stage:'pivot-snapshot-probe',...event})); } catch {}
  });
  await page.evaluate(() => {
    const proto = tabula.wb().constructor.prototype;
    const emit = event => console.log('WIXEL_PIVOT_SNAPSHOT_PROBE '+JSON.stringify(event));
    const summary = book => {
      let definitions=0,definitionsWithSnapshot=0;
      for(const sheet of book.sheets??[])for(const def of [sheet.pivot,...(sheet.pivotsExtra??[])])if(def){definitions++;if(def.snapshotId)definitionsWithSnapshot++;}
      return {snapshots:book.pivotSnapshots?.size??0,definitions,definitionsWithSnapshot};
    };
    let arrayDirtyCalls=0;
    const originalDirty=proto.markFormulaDirty;
    proto.markFormulaDirty=function(si,r,c,cell){
      if(cell?.cachedArray&&!cell.dirty){arrayDirtyCalls++;if(arrayDirtyCalls<=50)emit({method:'markArrayDirty',si,r,c,storedValues:cell.cachedArray.values.length/3,at:Math.round(performance.now()),stack:new Error().stack?.split('\n').slice(1,9).join('\n')});}
      return originalDirty.apply(this,arguments);
    };
    const originalSet = proto.setSnapshots;
    proto.setSnapshots = function(data) {
      const inputSnapshots=Object.keys(data?.pivotSnapshots??{}).length;
      const before=summary(this);
      const result=originalSet.apply(this,arguments);
      emit({method:'setSnapshots',inputSnapshots,before,after:summary(this),at:Math.round(performance.now())});
      return result;
    };
    const originalCurrent = proto.pivotSnapshotCurrent;
    let calls=0, failures=0;
    proto.pivotSnapshotCurrent = function(snap,def) {
      calls++;
      const result=originalCurrent.apply(this,arguments);
      if(!result&&failures++<100){
        const source=this.pivotSnapshotSource(def);
        emit({method:'pivotSnapshotCurrent',ok:false,sourceFound:!!source,sourceSheetSame:!!source&&snap.sourceSheet===source.sheet,sourceKeySame:!!source&&snap.sourceKey===source.key,snapshotVersion:snap.ver??null,sourceVersion:source?this.sourceVersion(source.si):null,at:Math.round(performance.now()),stack:new Error().stack?.split('\n').slice(1,9).join('\n')});
      }
      return result;
    };
    window.__pivotSnapshotProbe=()=>({calls,failures,arrayDirtyCalls,...summary(tabula.wb())});
  });
}
