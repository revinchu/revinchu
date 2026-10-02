import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readFile, mkdir, mkdtemp, writeFile, rm } from 'node:fs/promises';
import { resolve, join, dirname, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { execFile } from 'node:child_process';
import { promisify } from 'node:util';
import { createWixelServer } from '../server/app.js';
const exec = promisify(execFile),root=fileURLToPath(new URL('../',import.meta.url));
const paths=['manifest.webmanifest','apple-touch-icon.png','favicon.ico','favicon-32x32.png','icons/wixel-192.png','icons/wixel-512.png','icons/wixel-maskable-512.png'];
const pngSizes={'apple-touch-icon.png':180,'favicon-32x32.png':32,'icons/wixel-192.png':192,'icons/wixel-512.png':512,'icons/wixel-maskable-512.png':512};
const tempBase=process.platform==='win32'?'D:/Codex/Temp/wixel-icons':join(tmpdir(),'wixel-icons');
async function temporary(t){await mkdir(tempBase,{recursive:true});const dir=await mkdtemp(join(tempBase,'test-'));t.after(async()=>{assert.ok(resolve(dir).startsWith(resolve(tempBase)+sep));await rm(dir,{recursive:true,force:true});});return dir;}
const typeFor=p=>p.endsWith('.webmanifest')?'application/manifest+json':p.endsWith('.ico')?'image/vnd.microsoft.icon':'image/png';

test('홈 아이콘 manifest와 링크는 하위 경로·문서 query/hash에서도 앱 루트를 유지한다',async()=>{
 const m=JSON.parse(await readFile(join(root,'manifest.webmanifest'),'utf8')),html=await readFile(join(root,'index.html'),'utf8');
 assert.equal(m.short_name,'WIXEL');assert.equal(m.display,'standalone');assert.equal(m.icons.length,3);assert.equal(m.icons.filter(x=>x.purpose==='maskable').length,1);assert.equal(m.icons.find(x=>x.purpose==='maskable').sizes,'512x512');
 for(const path of ['manifest.webmanifest','apple-touch-icon.png','favicon.ico','favicon-32x32.png'])assert.ok(html.includes('href="'+path+'"'),path);
 for(const folder of ['https://example.test/','https://example.test/wixel/','https://example.test/repo/tabula/dist/']){
  const page=new URL('index.html?view=sample#private-fragment',folder),manifestURL=new URL('manifest.webmanifest',page),start=new URL(m.start_url,manifestURL);
  assert.equal(start.href,folder);assert.equal(new URL(m.scope,manifestURL).href,folder);assert.equal(new URL(m.id,start).href,folder);
  for(const icon of m.icons){assert.equal(new URL(icon.src,manifestURL).href,folder+icon.src);assert.equal(icon.type,'image/png');}
 }
});

test('Node 서버는 공개 아이콘과 manifest를 정확한 MIME으로 제공하고 다른 파일은 차단한다',async(t)=>{
 const dir=await temporary(t),publicDir=join(dir,'public'),data=join(dir,'private');await mkdir(publicDir,{recursive:true});
 for(const prefix of ['', 'dist/'])for(const path of paths){await mkdir(dirname(join(publicDir,prefix+path)),{recursive:true});await writeFile(join(publicDir,prefix+path),path.endsWith('.webmanifest')?'{}':Buffer.from([1,2,3]));}
 await mkdir(join(publicDir,'icons'),{recursive:true});await writeFile(join(publicDir,'icons/private.json'),'synthetic private marker');
 const server=createWixelServer({root:publicDir,data});await new Promise((yes,no)=>{server.once('error',no);server.listen(0,'127.0.0.1',yes);});
 t.after(async()=>{server.closeAllConnections?.();await new Promise(yes=>server.close(yes));});const base='http://127.0.0.1:'+server.address().port+'/';
 for(const prefix of ['', 'dist/'])for(const path of paths){const res=await fetch(base+prefix+path,{headers:{connection:'close'}});assert.equal(res.status,200,prefix+path);assert.ok(res.headers.get('content-type').startsWith(typeFor(path)));assert.equal(res.headers.get('x-content-type-options'),'nosniff');await res.arrayBuffer();}
 const head=await fetch(base+'manifest.webmanifest',{method:'HEAD',headers:{connection:'close'}});assert.equal(head.status,200);assert.equal((await head.arrayBuffer()).byteLength,0);
 for(const path of ['icons/private.json','icons/unrelated.png','server.js','dist/server.js']){const res=await fetch(base+path,{headers:{connection:'close'}});assert.equal(res.status,403,path);await res.arrayBuffer();}
});

test('일반·Cloudflare 빌드는 규격에 맞는 모든 아이콘을 포함하고 단일 HTML favicon을 유지한다',async(t)=>{
 const dir=await temporary(t),plain=join(dir,'portable'),cloud=join(dir,'cloud');
 for(const [path,size]of Object.entries(pngSizes)){const bytes=await readFile(join(root,path));assert.equal(bytes.subarray(0,8).toString('hex'),'89504e470d0a1a0a',path);assert.equal(bytes.readUInt32BE(16),size,path);assert.equal(bytes.readUInt32BE(20),size,path);}
 const ico=await readFile(join(root,'favicon.ico'));assert.equal(ico.readUInt16LE(0),0);assert.equal(ico.readUInt16LE(2),1);assert.equal(ico.readUInt16LE(4),3);assert.deepEqual(Array.from({length:3},(_,i)=>ico[6+i*16]),[16,32,48]);
 for(const [out,args]of [[plain,[]],[cloud,['--cloud']]]){
  await exec(process.execPath,[join(root,'build.mjs'),out,...args],{cwd:root,windowsHide:true,maxBuffer:1024*1024,timeout:60000});
  for(const path of paths)assert.deepEqual(await readFile(join(out,path)),await readFile(join(root,path)),path);
 }
 const portable=await readFile(join(plain,'index.html'),'utf8'),deployed=await readFile(join(cloud,'index.html'),'utf8'),headers=await readFile(join(cloud,'_headers'),'utf8');
 assert.match(portable,/<link rel="icon" href="data:image\/png;base64,/);assert.ok(!portable.includes('src="src/app.js"'));assert.ok(!portable.includes('href="styles.css"'));
 assert.match(deployed,/<link rel="icon" href="favicon-32x32.png"/);assert.match(deployed,/<script type="module" src="wixel-[a-f0-9]{16}\.js"><\/script>/);assert.match(headers,/\/manifest.webmanifest\s+Content-Type: application\/manifest\+json; charset=utf-8\s+Cache-Control: no-cache/);
});
