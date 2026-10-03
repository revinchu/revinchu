// Synthetic functional-suite inventory/orchestrator. Does not modify product files or user documents.
import { readdir,readFile,writeFile,mkdir,access } from 'node:fs/promises';
import { createWriteStream } from 'node:fs';
import { spawn,execFileSync } from 'node:child_process';
import { dirname,resolve,join } from 'node:path';
import { fileURLToPath,pathToFileURL } from 'node:url';
import { createHash } from 'node:crypto';
const root=resolve(dirname(fileURLToPath(import.meta.url)),'..'),tools=join(root,'tools');
const args=process.argv.slice(2),get=(flag,fallback)=>{const i=args.indexOf(flag);return i<0?fallback:args[i+1];};
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/',out=resolve(get('--out',process.env.WIXEL_AUDIT_OUT||'D:/Codex/Temp/wixel-functional-audit'));
if(!['localhost','127.0.0.1','[::1]'].includes(new URL(url).hostname)||!/^D:[\\/]/i.test(out))throw Error('로컬 서버와 D: 출력만 허용합니다.');
const concurrency=Math.max(1,Math.min(2,Number(get('--concurrency','2')))),timeoutMs=Number(get('--timeout','600'))*1000;
const only=get('--only','').split(',').filter(Boolean),listOnly=args.includes('--list'),resume=args.includes('--resume');
await mkdir(out,{recursive:true});
const delegated=new Set(['pivot-report-layout','pivot-display-options','pivot-filter-menu']);
const fileCli=new Set(['brcheck','check','pvcmp','real-workbook-performance','excel-fixtures','excel-pivot-fixtures','excel-chart-gallery','excel-roundtrip','chart-ex-fidelity']);
const operational=new Set(['recovery-ux','vault-ui-integration']);
const performance=new Set(['performance-regression','hidden-grid-performance','filter-performance','axis-window-benchmark']);
const infrastructure=new Set(['functional-audit','audit-playwright','audit-network-preload']);
const priority=['settings-ux','ui-regressions','interaction-ux','keys','ribbon-keytip-coverage','validation-parity','text-to-columns-ux','format-regressions','paste-special','review-ux','cell-style-regressions','clear-ui','range-query','print-area-dialog','page-break-preview','print-pagination','page-layout','save-dialog','server-save-dialog','server-open-guard','storage-atomic','library-atomic','library-recovery','online-pictures','online-media','image-export','picture-format-parity','picture-ux','shape-drawing','shape-format-parity','shape-merge-ui','shape-text','drawing-links','drawing-mode-isolation','object-group','smartart'];
function category(id){if(/^(pivot|slicer)-/.test(id))return '피벗·슬라이서';if(/^(chart|axis-overlap)/.test(id))return '차트';if(/key|access|palette/.test(id))return '키보드·접근키';if(/^(grid|freeze|merged|mobile|home)/.test(id))return '격자·모바일·접근성';if(/^(page|print)/.test(id))return '인쇄·페이지';if(/image|picture|shape|drawing|object|smartart|creative|svg-|online/.test(id))return '그림·도형·미디어';if(/save|storage|library|recovery|vault|server|release|wixel3/.test(id))return '저장·복구·연결';if(/csv|importrange|date|calculation/.test(id))return '가져오기·계산';if(/dialog|popup|gallery/.test(id))return '대화상자';return '셀 편집·설정·선택';}
const manifest=[];
for(const file of(await readdir(tools)).filter(f=>f.endsWith('.mjs')).sort()){
 const id=file.slice(0,-4);if(infrastructure.has(id))continue;const source=await readFile(join(tools,file),'utf8'),env=[...new Set([...source.matchAll(/process\.env\.([A-Z_0-9]+)/g)].map(m=>m[1]))];
 let disposition='run',reason='합성 격리 브라우저 기능 검사';
 if(delegated.has(id)){disposition='delegated';reason='handoff_review 별도 실행 예정. 결과 수신 전 통과로 세지 않음';}
 else if(fileCli.has(id)){disposition='fixture-required';reason='명시적 파일 CLI 또는 Excel 합성/왕복 도구: script_check 담당, 사용자 파일 자동 접근 금지';}
 else if(operational.has(id)){disposition='separate-environment';reason='실제 보관함 API 쓰기 검사. 별도 임시 Worker/격리 저장소 준비 없이는 실행하지 않음';}
 else if(performance.has(id)){disposition='isolated-performance';reason='동시 부하에서 성능 수치 왜곡. 기능 배치 이후 단독 실행 가능';}
 else if(id==='update-map-data'){disposition='not-a-test';reason='외부 데이터 다운로드 및 제품 자산 갱신 도구';}
 else if(['smoke','settings-excel-roundtrip','chart-setting-bindings'].includes(id)){disposition='root-owned';reason='root 담당 별도 검사. 이 배치 결과에 합산하지 않으며 명령 열기 smoke는 기능 정확성을 대체하지 않음';}
 else if(id==='command-connection-audit'){disposition='static-connectivity';reason='명령·메뉴 연결 전용 감사, 기능 배치와 별도 실행';}
 else if(!/chromium|playwright/.test(source)){disposition='node';reason='순수 Node 합성 검사';}
 const sourceOnly=/import\(['"]\/src\//.test(source),fileArg=/process\.argv/.test(source);
 manifest.push({id,file,category:category(id),disposition,reason,sourceOnly,fileArg,env,existingNetworkGuard:/route\(/.test(source),fixtures:[...new Set([...source.matchAll(/(?:readFile|readFileSync)\([^\n]{0,150}/g)].map(m=>m[0]))],assertionSignals:{assertions:(source.match(/\bassert\.|\beq\(|\bok\(|\bcheck\(/g)||[]).length,undo:/\bundo|Undo/.test(source),cancel:/취소|cancel|Escape/.test(source)}});
}
await writeFile(join(out,'manifest.json'),JSON.stringify({created:new Date().toISOString(),url,scope:'기능별 기존 회귀이며 Excel 동등성/실기 전체 보장은 아님',tools:manifest},null,2));
console.log(JSON.stringify({inventory:manifest.length,dispositions:manifest.reduce((a,t)=>(a[t.disposition]=(a[t.disposition]||0)+1,a),{}),manifest:join(out,'manifest.json')}));
if(listOnly)process.exit(0);
const selected=manifest.filter(t=>only.length?only.includes(t.id):['run','node'].includes(t.disposition)).sort((a,b)=>(priority.indexOf(a.id)<0?999:priority.indexOf(a.id))-(priority.indexOf(b.id)<0?999:priority.indexOf(b.id))||a.id.localeCompare(b.id));
if(selected.some(t=>!['run','node','isolated-performance','static-connectivity'].includes(t.disposition)))throw Error('--only에도 실제 API/사용자 파일 도구는 허용하지 않습니다.');
const real=process.env.WIXEL_AUDIT_PLAYWRIGHT_REAL||process.env.PLAYWRIGHT_MODULE||'playwright';
const results=[];let cursor=0,persistQueue=Promise.resolve();
// Re-enumerate each time: additions/deletions under src must invalidate resume too.
const SOURCE_FINGERPRINT_VERSION=2;
async function fingerprint(){
 const names=['index.html','styles.css'];
 async function collect(dir){for(const entry of await readdir(join(root,dir),{withFileTypes:true})){const name=dir+'/'+entry.name;if(entry.isDirectory())await collect(name);else if(entry.isFile()&&entry.name.endsWith('.js'))names.push(name);}}
 await collect('src');names.sort();
 const values=await Promise.all(names.map(async name=>[name,createHash('sha256').update(await readFile(join(root,name))).digest('hex')]));
 return Object.fromEntries(values);
}
function summary(log){const lines=log.split(/\r?\n/);for(let i=lines.length-1;i>=0;i--){const text=lines[i].trim();if(text[0]!=='{'&&text[0]!=='[')continue;try{const value=JSON.parse(lines.slice(i).join(String.fromCharCode(10)).trim());if(Array.isArray(value))return{cases:value.length,passed:value.filter(x=>x.ok===true||x.pass===true).length,raw:value};return value;}catch{}}
 return{okLines:lines.filter(s=>/^OK\b|^PASS\b/.test(s)).length,badLines:lines.filter(s=>/^NG\b|^FAIL\b/.test(s)).length,tail:lines.slice(-6)};}
async function persist(){const done=results.filter(r=>r.status!=='running');const data=JSON.stringify({url,updated:new Date().toISOString(),selected:selected.length,completed:done.length,passed:done.filter(r=>r.status==='pass').length,failed:done.filter(r=>r.status==='fail').length,results},null,2);persistQueue=persistQueue.then(()=>writeFile(join(out,'results.json'),data));await persistQueue;}
async function execute(tool){const dir=join(out,tool.id);await mkdir(dir,{recursive:true});const resultPath=join(dir,'run.json');if(resume){try{const old=JSON.parse(await readFile(resultPath,'utf8'));if(old.status==='pass'&&old.sourceFingerprintVersion===SOURCE_FINGERPRINT_VERSION&&JSON.stringify(old.sourceAfter)===JSON.stringify(await fingerprint())){results.push({...old,resumed:true});return;}}catch{}}
 const env={...process.env};for(const key of Object.keys(env))if(/FILTER|BASELINE|_SKIP$/.test(key)&&/WIXEL|GRID|DRAWING|SLICER|GALLERY|MEDIA|CLASSIC/.test(key))delete env[key];
 Object.assign(env,{WIXEL_URL:url,WIXEL_AUDIT_PLAYWRIGHT_REAL:real,PLAYWRIGHT_MODULE:pathToFileURL(join(tools,'audit-playwright.mjs')).href,WIXEL_AUDIT_GUARD_LOG:join(dir,'network.jsonl'),TEMP:process.env.TEMP||'D:/Codex/Temp',TMP:process.env.TMP||'D:/Codex/Temp'});
 for(const key of tool.env){if(key==='WIXEL_URL'||key==='PLAYWRIGHT_MODULE')continue;if(/SCREENSHOTS$|_OUT$|_OUTPUT$|_DIR$/.test(key))env[key]=join(dir,key.toLowerCase());else if(/SCREENSHOT$/.test(key))env[key]=join(dir,key.toLowerCase()+'.png');else if(/RESULT_PATH$/.test(key))env[key]=join(dir,key.toLowerCase()+'.json');else if(/_PDF$/.test(key))env[key]=join(dir,key.toLowerCase()+'.pdf');if(env[key]&&/SCREENSHOTS$|_OUT$|_OUTPUT$|_DIR$/.test(key))await mkdir(env[key],{recursive:true});}
 const before=await fingerprint(),entry={toolSha256:createHash('sha256').update(await readFile(join(tools,tool.file))).digest('hex'),sourceFingerprintVersion:SOURCE_FINGERPRINT_VERSION,sourceScope:Object.keys(before),id:tool.id,category:tool.category,status:'running',started:new Date().toISOString(),log:join(dir,'output.log'),sourceBefore:before};results.push(entry);await persist();console.log('START '+tool.id);
 await writeFile(env.WIXEL_AUDIT_GUARD_LOG,'');const output=createWriteStream(entry.log);let timedOut=false;const start=Date.now();const child=spawn(process.execPath,['--import',pathToFileURL(join(tools,'audit-network-preload.mjs')).href,join(tools,tool.file)],{cwd:root,env,stdio:['ignore','pipe','pipe'],windowsHide:true});child.stdout.pipe(output,{end:false});child.stderr.pipe(output,{end:false});
 const timer=setTimeout(()=>{timedOut=true;if(process.platform==='win32')spawn('taskkill.exe',['/PID',String(child.pid),'/T','/F'],{windowsHide:true,stdio:'ignore'});else child.kill('SIGKILL');},timeoutMs);
 const exit=await new Promise(resolve=>{child.on('error',e=>resolve({code:null,error:e.message}));child.on('close',(code,signal)=>resolve({code,signal}));});clearTimeout(timer);await new Promise(resolve=>output.end(resolve));
 const text=await readFile(entry.log,'utf8'),after=await fingerprint();let network=[];try{network=(await readFile(env.WIXEL_AUDIT_GUARD_LOG,'utf8')).trim().split('\n').filter(Boolean).map(JSON.parse);}catch{}
 const parsed=summary(text),summaryFailed=parsed.ok===false||Number(parsed.failed)>0||Number(parsed.bad)>0||(Array.isArray(parsed.bad)&&parsed.bad.length>0)||(Array.isArray(parsed.raw)&&parsed.raw.some(x=>x.ok===false||x.pass===false));
 Object.assign(entry,{status:exit.code===0&&!timedOut&&!summaryFailed?'pass':'fail',durationMs:Date.now()-start,exitCode:exit.code,signal:exit.signal,timedOut,summary:parsed,sourceAfter:after,sourceScopeAfter:Object.keys(after),sourceChanged:JSON.stringify(before)!==JSON.stringify(after),blockedNetwork:network.filter(x=>x.kind!=='pageerror'),observedPageErrors:network.filter(x=>x.kind==='pageerror'),classification:summaryFailed?'failing-result-summary':exit.code===0?'passed-existing-assertions':timedOut?'timeout-needs-triage':/ERR_CONNECTION_REFUSED|Executable doesn't exist|ERR_MODULE_NOT_FOUND/.test(text)?'environment':'untriaged-test-failure'});
 await writeFile(resultPath,JSON.stringify(entry,null,2));await persist();console.log((entry.status==='pass'?'PASS ':'FAIL ')+tool.id+' '+(entry.durationMs/1000).toFixed(1)+'s code='+entry.exitCode);
}
await Promise.all(Array.from({length:concurrency},async()=>{while(cursor<selected.length){const tool=selected[cursor++];await execute(tool);}}));await persist();console.log(JSON.stringify({completed:results.length,passed:results.filter(r=>r.status==='pass').length,failed:results.filter(r=>r.status==='fail').length,out}));if(results.some(r=>r.status==='fail'))process.exitCode=1;
