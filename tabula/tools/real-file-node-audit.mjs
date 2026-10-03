// 업무 파일은 읽기만 합니다. 익명 ID 요약만 stdout에, 상세 진단은 지정한 로컬 경로에 기록합니다.
// 대형 파일은 real-file-node-audit.py로 한 개씩 실행해 메모리/시간을 제한하세요.
import { readFileSync, writeFileSync, existsSync } from 'node:fs';
import { pathToFileURL, fileURLToPath } from 'node:url';
const root = fileURLToPath(new URL('..', import.meta.url)).replace(/[\\/]$/, '');
const { readXlsx, readXlsxAsync } = await import(pathToFileURL(root + '/src/xlsx.js'));
const { Workbook } = await import(pathToFileURL(root + '/src/workbook.js'));
const { pivotSourceData, resolvePivot, computePivot } = await import(pathToFileURL(root + '/src/pivot.js'));
const [manifestPath = process.env.WIXEL_AUDIT_MANIFEST, id, mode = 'audit', out, reader = process.env.WIXEL_AUDIT_READER ?? 'async'] = process.argv.slice(2);
if (!manifestPath || !/^[A-Za-z0-9_-]{1,64}$/.test(id ?? '') || !out || !['inventory','pivots','formulas','audit'].includes(mode)) throw new Error('사용법: node tools/real-file-node-audit.mjs manifest.json 익명ID audit 결과.json');
// 같은 출력 폴더에 .audit-pause를 두면 다음 파일의 읽기만 잠시 멈춥니다.
process.on('uncaughtException', error => {
 let partial={id,mode,reader};try{partial={...partial,...JSON.parse(readFileSync(out,'utf8'))};}catch{}
 try{writeFileSync(out,JSON.stringify({...partial,phase:'failed',status:'failed',error:{name:error.name,message:error.message}},null,2));}catch{}
 console.error(JSON.stringify({id,mode,status:'failed',error:error.name}));process.exit(1);
});
const pauseFile = new URL('./.audit-pause', pathToFileURL(out));
while (existsSync(pauseFile)) { writeFileSync(out, JSON.stringify({id,mode,reader,phase:'waiting-slot'})); await new Promise(resolve=>setTimeout(resolve,500)); }
const manifest=JSON.parse(readFileSync(manifestPath,'utf8').replace(/^\uFEFF/,''));
const entry=(Array.isArray(manifest)?manifest:manifest.files).find(x=>String(x.id)===id);
if(!entry) throw new Error('파일 ID가 없습니다.');
const started=Date.now();
writeFileSync(out,JSON.stringify({id,mode,phase:'reading'}));
console.log(JSON.stringify({id,mode,phase:'reading'}));
const bytes=readFileSync(entry.path);
let progressAt=0, diagnostics; const sheetDiagnostics=new Map();
const progress=event=>{if(Date.now()-progressAt<1000)return;progressAt=Date.now();writeFileSync(out,JSON.stringify({id,mode,reader,phase:'reading',elapsedMs:Date.now()-started,progress:event,diagnostics,memory:process.memoryUsage()},null,2));};
const res=reader==='sync'?readXlsx(bytes):await readXlsxAsync(bytes,progress,{onDiagnostics:event=>{diagnostics=event;sheetDiagnostics.set(event.part,event);}});
const wb=new Workbook(res.data);
const parsedMs=Date.now()-started;
const report={id,mode,reader,bytes:bytes.length,parsedMs,importDiagnostics:[...sheetDiagnostics.values()],warnings:res.warnings??[],sheetCount:wb.sheets.length,ownSheetCount:wb.ownSheetCount(),cells:0,formulas:0,parseErrors:0,pivots:0,charts:0,images:0,shapes:0,slicers:0,tables:0,blocks:0,blockCells:0,sheets:[]};
function checkpoint(phase) {report.phase=phase;report.elapsedMs=Date.now()-started;report.memory=process.memoryUsage();writeFileSync(out,JSON.stringify(report,null,2));}
wb.sheets.forEach((s,si)=>{
 let formulas=0,parseErrors=0; s.cells.forEachStoredRC(cell=>{if(cell.formula)formulas++;if(cell.parseError)parseErrors++;});
 const pivots=[s.pivot,...(s.pivotsExtra??[])].filter(Boolean);
 const blockCells=(s.blocks??[]).reduce((n,b)=>n+b.n*b.cols.length,0);
 report.cells+=s.cells.size; report.formulas+=formulas; report.parseErrors+=parseErrors; report.pivots+=pivots.length; report.blockCells+=blockCells;
 for(const prop of ['charts','images','shapes','slicers','tables','blocks']) report[prop]+=(s[prop]??[]).length;
 report.sheets.push({si,cells:s.cells.size,formulas,parseErrors,pivots:pivots.length,blockCells,hidden:s.state??null});
});
checkpoint('parsed');
if(mode==='pivots'||mode==='audit'){
 report.pivotResults=[];
 const norm=v=>v&&typeof v==='object'?(v.code??String(v)):v??null;
 const eq=(a,b)=>a===b||((a===null||a==='')&&(b===null||b===''))||(typeof a==='number'&&typeof b==='number'&&Math.abs(a-b)<=1e-6*Math.max(1,Math.abs(b)))||(typeof a==='number'&&typeof b==='string'&&String(a)===b)||(typeof b==='number'&&typeof a==='string'&&String(b)===a);
 wb.sheets.forEach((s,si)=>{ for(const [pi,d] of [s.pivot,...(s.pivotsExtra??[])].filter(Boolean).entries()){
  const t=Date.now(), p={si,pi,compared:0,differences:0,examples:[]};
  try{
   const src=pivotSourceData(wb,d);if(!src)throw new Error('피벗 원본을 찾을 수 없습니다.');
   const r=resolvePivot(src,d);const hdr=src.cube.header.map(h=>String(h??'').toLowerCase());
   r.fieldStyle=f=>{const i=hdr.indexOf(String(f).toLowerCase());return i>=0&&src.ref&&src.si!==undefined?wb.styleAt(src.si,Math.min(src.ref.r1+1,src.ref.r2),src.ref.c1+i):null;};
   const output=computePivot(r,r.def);let width=0;
   output.grid.forEach((row,ri)=>{width=Math.max(width,row.length);row.forEach((cd,ci)=>{
    if(!cd||cd.role==='empty')return;
    const rr=(d.top??0)+ri,cc=(d.left??0)+ci;
    let actual=String(cd.raw??'');if(actual.startsWith("'"))actual=actual.slice(1);else if(actual!==''&&Number.isFinite(Number(actual)))actual=Number(actual);
    const expected=norm(wb.getValue(si,rr,cc));p.compared++;
    if(!eq(actual,expected)){p.differences++;if(p.examples.length<8)p.examples.push({r:rr,c:cc,role:cd.role,actual,expected});}
   });});
   p.savedArea=d.area;p.computedArea={r1:d.top??0,c1:d.left??0,r2:(d.top??0)+output.grid.length-1,c2:(d.left??0)+width-1};
   p.areaChanged=d.area?['r1','c1','r2','c2'].some(k=>d.area[k]!==p.computedArea[k]):null;
  }catch(e){p.error=e.message;}
  p.ms=Date.now()-t;report.pivotResults.push(p);checkpoint('pivots');
 }});
}
if(mode==='formulas'||mode==='audit') {
 const saved=[];
 wb.sheets.forEach((s,si)=>s.cells.forEachStoredRC((cell,r,c)=>{if(cell.formula)saved.push([si,r,c,cell.cached]);}));
 const norm=v=>v&&typeof v==='object'?(v.code??v.error??String(v)):v??'';
 const equal=(a,b)=>a===b||(typeof a==='number'&&typeof b==='number'&&Math.abs(a-b)<=1e-9*Math.max(1,Math.abs(b)));
 const t=Date.now();wb.invalidate();report.formulaResults={total:saved.length,compared:0,noSavedValue:0,differences:0,throws:0,groups:{},examples:[]};
 for(const [si,r,c,cached] of saved) {
  const f=report.formulaResults;if(cached===undefined){f.noSavedValue++;continue;}
  let actual;try{actual=norm(wb.getValue(si,r,c));}catch(e){actual='THROW: '+e.message;f.throws++;}
  const expected=norm(cached);f.compared++;if(f.compared%2500===0)checkpoint('formulas');
  if(!equal(actual,expected)){
   f.differences++;const raw=wb.getCell(si,r,c)?.raw??'';const group=(raw.match(/[A-Z][A-Z0-9.]+(?=\()/g)??['-']).join(',');f.groups[group]=(f.groups[group]??0)+1;
   if(f.examples.length<30)f.examples.push({si,r,c,raw,actual,expected});
  }
 }
 report.formulaResults.ms=Date.now()-t;
}
report.totalMs=Date.now()-started;report.memory=process.memoryUsage();report.status='completed';report.phase='completed';
writeFileSync(out,JSON.stringify(report,null,2));
console.log(JSON.stringify({...report,sheets:undefined,warnings:report.warnings.length,formulaResults:report.formulaResults?{...report.formulaResults,examples:undefined}:undefined,pivotResults:report.pivotResults?.map(p=>({...p,examples:undefined}))}));
