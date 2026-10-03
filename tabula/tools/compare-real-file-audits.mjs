// 로컬 감사 JSON만 비교. 원본 열기·셀 값/수식 출력·네트워크 요청 없음.
import { readFile, writeFile, mkdir } from 'node:fs/promises';
import path from 'node:path';
import { normPage } from '../src/page.js';
const root=process.env.WIXEL_AUDIT_ROOT||'D:/Codex/Temp/wixel-final-audit';
const out=process.env.WIXEL_AUDIT_COMPARE_OUT||path.join(root,'comparison.json');
const nativeDir=process.env.WIXEL_NATIVE_AUDIT_DIR||path.join(root,'native');
const dirs=(process.env.WIXEL_BROWSER_AUDIT_DIRS||'browser-F20-compacted,browser-all-chromium').split(',').filter(Boolean).map(p=>path.isAbsolute(p)?p:path.join(root,p));
async function json(p){try{return JSON.parse((await readFile(p,'utf8')).replace(/^\uFEFF/,''));}catch(e){if(e.code==='ENOENT')return null;throw e;}}
const manifest=await json(process.env.WIXEL_REAL_AUDIT_MANIFEST||path.join(root,'manifest.json'));
if(!Array.isArray(manifest))throw Error('검수 매니페스트가 없습니다.');
const sourceOracle=await json(path.join(root,'saved-sheet-metadata-oracle.json'))||[];
const optionReviewOracle=await json(path.join(root,'option-review-oracle.json'))||[];
const valueEvidenceSelection=await json(path.join(root,'value-evidence-selection.json'))||{};
const editedExportSelection=await json(path.join(root,'native-edited-export-selection.json'))||{};
const editedReopenSelection=await json(path.join(root,'browser-edited-reopen-selection.json'))||{};
const editedReopenDirs=new Set(Object.values(editedReopenSelection));
const nativeExportSelection=await json(path.join(root,'native-export-selection.json'))||{};
const runAnnotations=await json(path.join(root,'audit-run-annotations.json'))||[];
const runs=(await Promise.all(dirs.map(async dir=>({dir,data:await json(path.join(dir,'result.json'))})))).filter(x=>x.data);
const same=(a,b)=>JSON.stringify(a)===JSON.stringify(b);
const nativeCounts=book=>({sheets:book.sheets.length,pivots:book.sheets.reduce((v,s)=>v+(s.pivots?.length||0),0),slicers:book.slicers?.length||0,tables:book.sheets.reduce((v,s)=>v+(typeof s.tables==='number'?s.tables:s.tables?.length||0),0),charts:book.sheets.reduce((v,s)=>v+(s.shapes||[]).filter(o=>o.type===3).length,0),shapes:book.sheets.reduce((v,s)=>v+(s.shapes?.length||0),0)});
function colName(c){let s='';for(c++;c;c=Math.floor((c-1)/26))s=String.fromCharCode(65+(c-1)%26)+s;return s;}
const area=r=>`${colName(r.c1)}${r.r1+1}:${colName(r.c2)}${r.r2+1}`;
const cleanRef=value=>String(value??'').replace(/\$/g,'').split(',').map(s=>s.slice(s.lastIndexOf('!')+1)).sort();
const nativeState=v=>v===-1?'visible':v===2?'veryHidden':'hidden';
function nativeStatus(n){if(!n)return'not-run';if(!n.opened)return'open-failed';if(n.errors?.length||n.unchanged===false)return'partial-or-error';return'completed';}
function browserStatus(r,run){if(!r)return'not-run';if(r.resourceInconclusive||(!r.opened&&!r.ok&&run?.aborted))return'resource-inconclusive';if(!r.opened)return'open-failed';if(r.error)return'partial-or-error';if(r.outcome==='unsupported-olap')return'unsupported-olap';if(r.scope==='import-only'||r.outcome==='import-only')return'import-only';if(r.scope==='objects-only'&&r.ok)return'objects-checked';return r.outcome==='unsupported-olap'?'unsupported-olap':r.ok?'completed':'partial-or-error';}
function compare(n,b,fileId){
 const differences=[],reviews=[],checked=[],skippedChecks=['전체 셀 대신 native 고정 8좌표/시트 표본만 비교','조건부 서식 규칙/계산 결과·차트/도형 세부 속성 전체 비교는 미수집'];
 // 이전 도구는 뒤에 붙는 외부 참조 가상 시트까지 포함했다. 원본 시트 이름에 해당하는 모델만 비교한다.
 const virtualExcluded=b.sheets.filter(s=>!n.sheets.some(x=>x.name===s.name)).length;
 b={...b,sheets:b.sheets.filter(s=>n.sheets.some(x=>x.name===s.name))};
 const check=(field,expected,actual)=>{checked.push(field);if(!same(expected,actual))differences.push({field,native:expected,wixel:actual});};
 check('sheetNames',n.sheets.map(s=>s.name),b.sheets.map(s=>s.name));
 const sheets=n.sheets.flatMap(ns=>{
  const bs=b.sheets.find(s=>s.name===ns.name);if(!bs)return[];
  const before=differences.length,reviewed=reviews.length,prefix=`sheet[${ns.index}].`;
  const eq=(field,expected,actual)=>check(prefix+field,expected,actual);
  eq('state',nativeState(ns.visible),bs.state||'visible');eq('protected',!!ns.protected,!!bs.protected);
  eq('tables',ns.tables?.length||0,bs.tables||0);eq('pivots',ns.pivots?.length||0,bs.pivots||0);
  // Excel 그룹은 최상위 1개, 가져오기 모델은 과거 자식별 개수였으므로 전체 도형 수를 단순 비교하지 않는다.
  eq('slicers',(n.slicers||[]).filter(s=>s.sheet===ns.name).length,bs.slicers||0);
  eq('charts',(ns.shapes||[]).filter(s=>s.type===3).length,bs.charts||0);
  if(ns.cfCount!=null&&bs.condCount!=null)eq('conditionalFormatCount',ns.cfCount,bs.condCount);
  if(ns.view){
   if(ns.view.zoom!=null)eq('zoom',ns.view.zoom,bs.zoom??100);
   if(ns.view.grid!=null)eq('gridlines',!!ns.view.grid,!bs.noGrid);
   if(ns.view.headers!=null)eq('headers',!!ns.view.headers,bs.view?.headers!==false);
   if(ns.view.view!=null)eq('viewMode',({1:'normal',2:'pageBreakPreview',3:'pageLayout'})[ns.view.view],bs.view?.mode||'normal');
   if(ns.view.freeze){eq('freeze.rows',ns.view.splitRow,(bs.freeze?.top||0)+(bs.freeze?.rows||0));eq('freeze.cols',ns.view.splitCol,(bs.freeze?.left||0)+(bs.freeze?.cols||0));}
   else eq('freeze.enabled',false,!!(bs.freeze?.rows||bs.freeze?.cols));
  }
  if(ns.page){
   const p=normPage(bs.page),np=ns.page;
   if(np.orientation!=null)eq('page.orientation',np.orientation===2?'landscape':'portrait',p.orientation);
   if(np.paperSize!=null)eq('page.paper',np.paperSize,p.paper);
   eq('page.printArea',cleanRef(np.printArea),cleanRef((p.areas?.length?p.areas:p.area?[p.area]:[]).map(area).join(',')));
   eq('page.titleRows',cleanRef(np.repeatRows),cleanRef(p.titleRows?p.titleRows.map(r=>r+1).join(':'):''));
   eq('page.titleCols',cleanRef(np.repeatCols),cleanRef(p.titleCols?p.titleCols.map(colName).join(':'):''));
   // Excel 비활성 자동맞춤 기본값1은 비교하지 않는다.
   if(np.zoom===false){eq('page.fitW',Number(np.fitWide)||0,p.fitW);eq('page.fitH',Number(np.fitTall)||0,p.fitH);}
   else if(typeof np.zoom==='number')eq('page.scale',np.zoom,p.scale);
  }
  if(ns.filterRange!=null)eq('autoFilter.range',cleanRef(ns.filterRange),cleanRef(bs.filter?.range?area({r1:bs.filter.range[0],c1:bs.filter.range[1],r2:bs.filter.range[2],c2:bs.filter.range[3]}):''));
  for(const[axis,rows,sizes,hidden]of[['row',ns.rowHeights,bs.rowSizes,bs.effectiveRowHidden],['col',ns.colWidths,bs.colSizes,bs.colHidden]])for(const r of rows||[]){
   if(hidden?.[r.index]!==undefined)eq(`${axis}[${r.index}].hidden`,!!r.hidden,hidden[r.index]);else skippedChecks.push(prefix+`${axis}[${r.index}].hidden: effective visibility not recorded`);
   if(!r.hidden&&typeof r.points==='number'&&sizes?.[r.index]!=null){
    const nativePx=Math.round(r.points*4/3*1000)/1000,actual=sizes[r.index];checked.push(prefix+`${axis}[${r.index}].size`);
    if(Math.abs(nativePx-actual)>1.01)reviews.push({field:prefix+`${axis}[${r.index}].size`,nativePx,wixelPx:actual,reason:axis==='row'?'원본 수동 높이와 Excel 자동 맞춤·글꼴 차이를 추가 확인':'원본 열 너비 환산을 추가 확인'});
   }
  }
  const nativeUsed=ns.used?{rows:ns.used.r2+1,cols:ns.used.c2+1}:null;
  if(nativeUsed&&!same(nativeUsed,bs.used))reviews.push({field:prefix+'usedRange',native:nativeUsed,wixel:bs.used,reason:'Excel은 서식만 있는 셀도 UsedRange에 포함하므로 데이터 누락 판정으로 사용하지 않음'});
  return[{index:ns.index,name:ns.name,differences:differences.length-before,reviews:reviews.length-reviewed}];
 });
 for(let i=differences.length-1;i>=0;i--){const d=differences[i],m=/^sheet\[(\d+)\]\.(.+)$/.exec(d.field);if(!m)continue;
  const saved=sourceOracle.find(x=>x.id===fileId&&x.index===+m[1]);if(!saved)continue;let source,reason;
  if(m[2]==='autoFilter.range'){
   const tag=saved.filter?.[0];source=typeof tag==='string'?cleanRef(/ref="([^"]+)"/.exec(tag)?.[1]||''):[];
   if(!saved.filter?.length)source=[''];
   reason='원본 저장 필터 범위는 WIXEL과 일치. Excel 열기 후 확장 또는 표의 필터 조회와 구분';
  }else if(m[2]==='freeze.rows'&&saved.pane?.[0]?.y!=null){source=saved.pane[0].y+(saved.view?.[0]?.top||0);reason='원본 저장 pane 값은 WIXEL과 일치. 숨긴 행이 있는 native SplitRow 반환과 구분';}
  else if(m[2]==='conditionalFormatCount'&&saved.cf!=null){source=saved.cf;reason='원본에 규칙 레코드 존재. native UsedRange.FormatConditions 개수와 조회범위가 다름';}
  if(source!==undefined&&same(source,d.wixel)){reviews.push({...d,source,reason});differences.splice(i,1);}
 }
 for(const row of sheets){const prefix=`sheet[${row.index}].`;row.differences=differences.filter(x=>x.field.startsWith(prefix)).length;row.reviews=reviews.filter(x=>x.field.startsWith(prefix)).length;}
 return{checks:checked.length,differences,reviews,skippedChecks,sheets,virtualExcluded};
}
const results=[];
for(const file of manifest){
 const n=await json(path.join(nativeDir,file.id+'.json'));
 const valueDigestName=valueEvidenceSelection[file.id]?.digest||file.id+'-all-values-digest.json';
 if(path.basename(valueDigestName)!==valueDigestName)throw Error('Value digest evidence must be a local filename.');
 const valueDigestPath=path.join(root,valueDigestName),valueDigest=await json(valueDigestPath);
 const nativeExportName=nativeExportSelection[file.id]||file.id+'.json';
 if(path.basename(nativeExportName)!==nativeExportName)throw Error('Native export evidence must be a local result filename.');
 const nativeExportPath=path.join(root,'native-export',nativeExportName);
 const nativeExport=await json(nativeExportPath);
 const arrayName=valueEvidenceSelection[file.id]?.array||file.id+'-array-native-comparison-final.json';
 if(path.basename(arrayName)!==arrayName)throw Error('Array evidence must be a local filename.');
 const arrayEvidence=path.join(root,arrayName);
 const editedName=editedExportSelection[file.id];
 if(editedName&&path.basename(editedName)!==editedName)throw Error('Edited export evidence must be a local filename.');
 const editedEvidence=editedName?path.join(root,'native-export',editedName):null,editedExport=editedEvidence?await json(editedEvidence):null;
 const editedReopenName=editedReopenSelection[file.id];
 if(editedReopenName&&path.basename(editedReopenName)!==editedReopenName)throw Error('Edited reopen evidence must be a local directory.');
 const editedReopenEvidence=editedReopenName?path.join(root,editedReopenName,'result.json'):null;
 const editedReopenData=editedReopenEvidence?await json(editedReopenEvidence):null;
 const editedReopenRecord=editedReopenData?.results?.find(r=>r.id===file.id);
 const editedComparisons={};
 for(const [key,configKey] of [['arrayComparison','editedArray'],['sourceEditComparison','sourceEdit']]){
  const name=valueEvidenceSelection[file.id]?.[configKey];if(!name)continue;
  if(path.basename(name)!==name)throw Error('Edited value evidence must be a local filename.');
  const evidence=path.join(root,name),candidate=await json(evidence);
  if(candidate?.nativeEvidence&&editedEvidence&&path.resolve(candidate.nativeEvidence)===path.resolve(editedEvidence))editedComparisons[key]={...candidate,evidence};
 }
 const arrayCandidate=await json(arrayEvidence);
 const nativeArrayComparison=arrayCandidate&&path.resolve(arrayCandidate.nativeEvidence)===path.resolve(nativeExportPath)?{...arrayCandidate,evidence:arrayEvidence}:null;
 const exportPreservation=nativeExport?.opened&&n?.opened?{expected:nativeCounts(n),actual:nativeCounts(nativeExport),sampleDifferences:nativeExport.sampleDifferences?.length||0}:null;
 if(exportPreservation)exportPreservation.countsMatch=same(exportPreservation.expected,exportPreservation.actual);
 const repairDiagnostic=await json(path.join(root,'native-export',file.id+'-repair-diagnostic.json'));
 const displayOracle=await json(path.join(root,'native-display-format',file.id+'.json'));
 const displayCells=displayOracle?.opened&&displayOracle.readOnly&&displayOracle.originalUnchanged&&!displayOracle.errors?.length?displayOracle.cells||[]:[];
 const candidates=runs.flatMap(run=>(run.data.results||[]).filter(r=>r.id===file.id&&!r.sourceEdit&&!editedReopenDirs.has(path.basename(run.dir))).map(record=>({record,run}))).sort((a,b)=>String(a.record.started||a.run.data.started).localeCompare(String(b.record.started||b.run.data.started)));
 const browsers=candidates.map(({record:r,run})=>{
  const annotation=runAnnotations.find(a=>a.id===file.id&&a.run===path.basename(run.dir));
  const status=browserStatus(r,run.data),metadata=r.metadata;
  const samples=r.nativeSamples||r.objectProbe||{},modelComparisonComplete=nativeStatus(n)==='completed'&&metadata?.sheets&&n.sheets.every(ns=>metadata.sheets.some(bs=>bs.name===ns.name));
  const displayFormatMatches=(samples.nativeStyleSamples?.differences||[]).flatMap(d=>{
   const cell=displayCells.find(c=>c.sheetIndex===d.sheet&&c.address===d.address);if(!cell?.display)return[];
   let value;if(d.field==='fill'){const color=cell.display.fill?.color;if(typeof color==='number')value='#'+[color&255,(color>>8)&255,(color>>16)&255].map(n=>n.toString(16).padStart(2,'0')).join('');}
   else if(d.field.startsWith('font.'))value=cell.display.font?.[d.field.slice(5)];
   return same(value,d.wixel)?[{sheet:d.sheet,address:d.address,field:d.field,display:value,reason:'Excel DisplayFormat matches WIXEL; direct Range style was not the displayed style'}]:[];
  });
  const optionReviews=(samples.nativeOptionSamples?.differences||[]).flatMap(d=>optionReviewOracle.filter(o=>o.id===file.id&&o.sheet===d.sheet&&o.field===d.field&&same(o.native,d.native)&&same(o.wixel,d.wixel)));
  return{evidence:path.join(run.dir,'result.json'),engine:r.engine,status,opened:!!r.opened,scope:r.scope||'workflow',contended:!!(annotation?.contended||r.contended||run.data.contended),...(annotation?.note?{auditNote:annotation.note}:{}),started:r.started||run.data.started,
   heapPolicy:run.data.browserLaunch?.heapPolicy??'legacy-launch-not-recorded',memoryObserver:run.data.browserLaunch?.memoryObserver??null,cdpSamples:(r.phases||[]).filter(p=>p.name==='cdp'),sourceHashes:r.scriptHashes||run.data.sourceHashesAtStart||run.data.sourceHashes,openMs:r.timing?.openMs,totalMs:r.timing?.totalMs,finalHeapBytes:r.finalMetrics?.heapUsed,
   systemMemory:r.systemMemory,warnings:r.warnings?.length||0,pageErrors:r.errors?.length||0,blockedRequests:r.blockedRequests?.length||0,edit:r.edit,
   outcome:r.outcome,error:r.error,structureProbe:r.structureProbe,exportProbe:r.exportProbe,export:r.export,workflowOk:r.workflowOk,nativeObjects:r.nativeObjects,memoryTelemetryNote:run.data.memoryTelemetryNote,objectChecks:r.objectProbe?.results,richImagePlacement:r.richImagePlacement,nativeStyleSamples:samples.nativeStyleSamples,displayFormatMatches,nativeContentSamples:samples.nativeContentSamples,nativeOptionSamples:samples.nativeOptionSamples,optionReviews,sampleCoverage:{nativeCells:samples.nativeContentSamples?.count??0,formulaCells:samples.nativeContentSamples?.formulas??0,pivotCells:samples.nativeContentSamples?.pivotCells??0,contentDifferences:samples.nativeContentSamples?.differences?.length??null,formulaDifferences:samples.nativeContentSamples?.formulaDifferences?.length??null,equivalentFormulaSyntax:samples.nativeContentSamples?.equivalentFormulaSyntax?.length??0,richValueReviews:samples.nativeContentSamples?.richValueReviews?.length??0,styleDifferences:samples.nativeStyleSamples?.differences? samples.nativeStyleSamples.differences.length-displayFormatMatches.length:null,styleRawDifferences:samples.nativeStyleSamples?.differences?.length??null,styleDisplayMatches:displayFormatMatches.length,styleComparisonScope:'native Range.Font/Interior와 WIXEL styleAt 비교. 별도 DisplayFormat oracle 일치 항목은 재분류하며 나머지 실제 표시 차이는 미확정',optionDifferences:samples.nativeOptionSamples?.differences? samples.nativeOptionSamples.differences.length-optionReviews.length:null,optionRawDifferences:samples.nativeOptionSamples?.differences?.length??null,optionReviewed:optionReviews.length},modelComparisonComplete:!!modelComparisonComplete,comparison:n?.opened&&metadata?.sheets?compare(n,metadata,file.id):null,
   completedWorkflowWithModelComparison:status==='completed'&&!!modelComparisonComplete,objectCoverage:{performed:(r.objectProbe?.results||[]).filter(x=>!x.skipped&&x.ok===true).length,skipped:(r.objectProbe?.results||[]).filter(x=>x.skipped).map(x=>({name:x.name,reason:x.skipped})),failed:(r.objectProbe?.results||[]).filter(x=>!x.skipped&&x.ok===false).length}};
 });
 results.push({id:file.id,bytes:file.bytes,editedReopenValidation:editedReopenRecord?{evidence:editedReopenEvidence,status:browserStatus(editedReopenRecord,editedReopenData),checks:editedReopenRecord.checks,sourceHashes:editedReopenRecord.scriptHashes,nativeSamples:editedReopenRecord.nativeSamples,objectProbe:editedReopenRecord.objectProbe,record:editedReopenRecord}:null,editedExportValidation:editedExport?{evidence:editedEvidence,status:editedExport.opened&&editedExport.readOnly&&editedExport.repairMode===false&&!editedExport.errors?.length?'normal-open-completed':'normal-open-not-passed',native:editedExport,structuralAudit:editedExport.structuralAudit??null,...editedComparisons}:null,storedValueAudit:valueDigest?{evidence:valueDigestPath,scope:'전체 저장된 셀 값 비교. 독립 수식 재계산 검증 아님; 검사 사본 버전과 진단 기록을 함께 확인.',ok:valueDigest.ok,sourcesUnchanged:valueDigest.sourcesUnchanged,differences:valueDigest.differences,formulaMetadataDifferences:valueDigest.formulaMetadataDifferences,left:valueDigest.left,right:valueDigest.right}:null,exportValidation:nativeExport?{status:nativeExport.opened&&nativeExport.readOnly&&nativeExport.repairMode===false&&!nativeExport.errors?.length?'normal-open-completed':'normal-open-not-passed',nativeEvidence:nativeExportPath,native:nativeExport,repairDiagnostic,preservation:exportPreservation,nativeArrayComparison}:null,native:{status:nativeStatus(n),evidence:path.join(nativeDir,file.id+'.json'),sheetCount:n?.sheets?.length,slicerCount:n?.slicers?.length,errors:n?.errors?.length||0,originalUnchanged:n?.unchanged,openMs:n?.openMs},browsers});
}
const summary={files:results.length,editedReopenCompleted:results.filter(r=>r.editedReopenValidation?.status==='completed').length,nativeCompleted:results.filter(r=>r.native.status==='completed').length,browserCompletedFiles:results.filter(r=>r.browsers.at(-1)?.status==='completed').length,nativeExportCompleted:results.filter(r=>r.exportValidation?.status==='normal-open-completed').length,nativeExportNotPassed:results.filter(r=>r.exportValidation?.status==='normal-open-not-passed').length,editedExportCompleted:results.filter(r=>r.editedExportValidation?.status==='normal-open-completed').length,editedExportNotPassed:results.filter(r=>r.editedExportValidation?.status==='normal-open-not-passed').length,editedExportValueFailures:results.filter(r=>r.editedExportValidation?.arrayComparison?.ok===false||r.editedExportValidation?.sourceEditComparison?.ok===false).length,nativeArrayComparisonFailures:results.filter(r=>r.exportValidation?.nativeArrayComparison?.ok===false).length,nativeExportPreservationDifferences:results.filter(r=>r.exportValidation?.preservation&&(r.exportValidation.preservation.countsMatch===false||r.exportValidation.preservation.sampleDifferences>0)).length,focusedObjectFiles:results.filter(r=>r.browsers.at(-1)?.status==='objects-checked').length,completedWorkflowsWithModelComparison:results.filter(r=>r.browsers.at(-1)?.completedWorkflowWithModelComparison).length,filesWithSkippedObjectChecks:results.filter(r=>r.browsers.at(-1)?.objectCoverage.skipped.length).length,modelComparedFiles:results.filter(r=>r.browsers.at(-1)?.modelComparisonComplete).length,filesWithComparisonDifferences:results.filter(r=>{const b=r.browsers.at(-1);return b?.modelComparisonComplete&&(b.comparison.differences.length||b.sampleCoverage.contentDifferences||b.sampleCoverage.formulaDifferences||b.sampleCoverage.styleDifferences||b.sampleCoverage.optionDifferences);}).length};
await mkdir(path.dirname(out),{recursive:true});await writeFile(out,JSON.stringify({generated:new Date().toISOString(),summarySelection:'파일별 가장 최근 실행 기준. 이전 실행은 근거 배열에 보존. import-only/objects-only는 전체 작업 완료로 승격하지 않으며 미지원 판정과 실행 범위를 별도 기록.',scope:'검사한 모델·설정·표본 서식만 비교. 전체 기능·셀 값·수식·모든 개체의 100% 동등성을 의미하지 않음.',summary,results},null,2));
console.log(JSON.stringify({...summary,out}));
