// Read-only command/menu/ribbon connection inventory. Existence is not functional parity.
import assert from 'node:assert/strict';
import {mkdir,writeFile,readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
const {chromium}=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const url=process.env.WIXEL_URL||'http://127.0.0.1:5191/';
const out=process.env.WIXEL_CONNECTION_OUT||'D:/Codex/Temp/wixel-functional-audit/connections';
if(!['localhost','127.0.0.1'].includes(new URL(url).hostname)||!/^D:[\\/]/i.test(out))throw Error('로컬 서버와 D: 출력만 사용하세요.');
await mkdir(out,{recursive:true});
const app=await readFile(new URL('../src/app.js',import.meta.url),'utf8');
const browser=await chromium.launch(),context=await browser.newContext({viewport:{width:1440,height:1000}}),page=await context.newPage(),errors=[],blocked=[];
try{
 await context.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});
 await context.route('**/*',route=>{const r=route.request(),u=new URL(r.url());if(!['GET','HEAD','OPTIONS'].includes(r.method())||u.origin!==new URL(url).origin||/^\/api(?:\/|$)/.test(u.pathname)){blocked.push({method:r.method(),url:r.url()});return route.abort();}return route.continue();});
 page.on('pageerror',e=>errors.push(e.message));
 await page.goto(url,{waitUntil:'domcontentloaded',timeout:60000});await page.waitForFunction(()=>window.tabula?.commands());
 const snapshot=await page.evaluate(async()=>{const t=tabula,{TABS}=await import('/src/ribbon.js'),{collectRibbonControls,auditRibbonKeytips}=await import('/src/ribbon-keytips.js');
  const controls=collectRibbonControls(TABS).map(({item,...control})=>control),registry=t.keytipRegistry();
  const leaves=[];function walk(items,path){for(const [i,item]of(items||[]).entries()){const next=[...path,i];if(item.items)walk(item.items,next);else leaves.push({path:next,type:item.type,cmd:item.cmd,menu:item.menu,title:item.title,label:item.label});}}
  for(const tab of TABS)for(const [i,g]of(tab.groups||[]).entries())walk(g.items,[tab.id,i]);
  return{commands:t.commands(),menus:t.menus(),controls,leaves,entries:registry.entries,keytipAudit:auditRibbonKeytips(controls,registry.entries,registry.tabs),tabs:TABS.map(({id,label,context,file})=>({id,label,context,file}))};});
 const commands=new Set(snapshot.commands),menus=new Set(snapshot.menus),issues=[];
 const dynamicMenuRoutes=snapshot.controls.filter(c=>c.kind==='menu'&&c.target.startsWith('fn:')).map(c=>({target:c.target,router:"openNamedMenu: name.startsWith('fn:') → fnMenu(category)"}));
 assert.ok(app.includes("if (name.startsWith('fn:'))"),'함수 메뉴 동적 라우터 존재');
 const linked=(kind,target)=>kind==='menu'?(menus.has(target)||dynamicMenuRoutes.some(m=>m.target===target)):commands.has(target);
 for(const c of snapshot.controls)if(!linked(c.kind,c.target))issues.push({kind:'missing-control-target',id:c.id,target:c.target});
 for(const e of snapshot.entries)if(!linked(e.kind,e.target))issues.push({kind:'missing-keytip-target',path:e.path,target:e.target});
 const start=app.indexOf('const COMMANDS = {'),end=app.indexOf('const NO_COMMIT',start),body=app.slice(start,end);
 const commandDefinitions=snapshot.commands.map(id=>{const escaped=id.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),match=new RegExp('(?:^|[,\\n])\\s*'+escaped+'\\s*:').exec(body);return{id,line:match?app.slice(0,start+match.index).split('\n').length:null,referencedBy:snapshot.controls.filter(c=>c.kind!=='menu'&&c.target===id).map(c=>c.id)};});
 const result={url,created:new Date().toISOString(),sourceSha256:createHash('sha256').update(app).digest('hex'),scope:'리본·키팁 참조가 등록 COMMANDS/MENUS에 연결되는지 확인. 메뉴 실행/기능 정확성/데이터 보존은 별도 기능 도구 범위.',counts:{commands:commands.size,menus:menus.size,tabs:snapshot.tabs.length,leafItems:snapshot.leaves.length,controls:snapshot.controls.length,keytips:snapshot.entries.length,issues:issues.length},issues,dynamicMenuRoutes,pageErrors:errors,blockedRequests:blocked,commandDefinitions,...snapshot};
 await writeFile(out+'/connections.json',JSON.stringify(result,null,2));console.log(JSON.stringify({counts:result.counts,issues,pageErrors:errors,out}));assert.deepEqual(issues,[]);assert.deepEqual(errors,[]);assert.deepEqual(snapshot.keytipAudit.missingControlIds,[]);
}finally{await context.close();await browser.close();}
