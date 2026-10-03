// 별도 Node 프로세스의 합성 30만 셀. 관측 heap은 peak RSS/실제 iPhone 측정이 아닙니다.
import { spawnSync } from 'node:child_process';
import { writeFileSync, mkdirSync } from 'node:fs';
import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { createHash } from 'node:crypto';
import { Workbook } from '../src/workbook.js';
import { librarySnapshot } from '../src/snapshot-blob.js';
const mode=process.argv[2], count=300000;
if(mode==='old'||mode==='blob'){
  const w=new Workbook();for(let i=0;i<count;i++)w.setInput(0,Math.floor(i/10),i%10,String(i));
  for(const [,cell] of w.sheets[0].cells)cell.style={font:'맑은 고딕',size:11,bold:true,numFmt:'comma',fill:'#ffffff'};
  globalThis.gc?.();const before=process.memoryUsage();let sampledPeak=before.heapUsed;
  const cells=w.sheets[0].cells, iterator=cells[Symbol.iterator].bind(cells);
  cells[Symbol.iterator]=function*(){let n=0;for(const v of iterator()){if(++n%4096===0)sampledPeak=Math.max(sampledPeak,process.memoryUsage().heapUsed);yield v;}};
  const metadata={app:'wixel',docName:'합성 30만 셀',docId:'memory-fixture',si:0},start=performance.now();
  const output=mode==='old'?JSON.stringify({...metadata,workbook:w.serialize()}):librarySnapshot(w,metadata).blob;
  const ms=performance.now()-start,after=process.memoryUsage();sampledPeak=Math.max(sampledPeak,after.heapUsed);
  const hash=createHash('sha256');let bytes=0;
  if(typeof output==='string'){bytes=Buffer.byteLength(output);hash.update(output);}else{for await(const part of output.stream()){bytes+=part.length;hash.update(part);}}
  console.log(JSON.stringify({mode,ms,heapBefore:before.heapUsed,heapAfter:after.heapUsed,sampledPeakHeap:sampledPeak,externalAfter:after.external,arrayBuffersAfter:after.arrayBuffers,rssAfter:after.rss,bytes,sha256:hash.digest('hex')}));
}else{
  const runs=[];for(let i=0;i<3;i++)for(const mode of ['old','blob']){
    const result=spawnSync(process.execPath,['--expose-gc',fileURLToPath(import.meta.url),mode],{encoding:'utf8',maxBuffer:1048576});
    if(result.status!==0)throw new Error(result.stderr||result.stdout);const run=JSON.parse(result.stdout);runs.push(run);console.log(JSON.stringify(run));
  }
  if(new Set(runs.map(r=>r.sha256)).size!==1)throw new Error('직렬화 내용 불일치');
  const median=xs=>xs.sort((a,b)=>a-b)[Math.floor(xs.length/2)];
  const summary=Object.fromEntries(['old','blob'].map(mode=>{const rows=runs.filter(r=>r.mode===mode);return[mode,{ms:median(rows.map(r=>r.ms)),addedSampledHeapMiB:median(rows.map(r=>(r.sampledPeakHeap-r.heapBefore)/1048576)),bytes:rows[0].bytes}];}));
  const report={date:new Date().toISOString(),node:process.version,count,runs,summary,note:'같은 PC의 별도 Node 프로세스 3회 중앙값. heap은 4096셀마다 및 종료시 관측, 완전한 최대 메모리나 iPhone 측정이 아님. Blob native 메모리·RSS는 별도 기록. 성능 보증이 아님.'};
  const path=process.env.WIXEL_SNAPSHOT_BENCHMARK||'D:/Codex/Temp/wixel-improvement/snapshot-memory.json';mkdirSync(dirname(path),{recursive:true});writeFileSync(path,JSON.stringify(report,null,2));console.log(JSON.stringify(summary,null,2));
}
