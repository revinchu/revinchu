// 합성 문서만 사용: 두 날짜 체계 XLSX 열기, 날짜 입력/Undo, 화면·차트 표시와 저장값 검사.
// WIXEL_URL, PLAYWRIGHT_MODULE, PLAYWRIGHT_BROWSERS_PATH 사용. 서버 쓰기를 차단합니다.
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
const { chromium } = await import(process.env.PLAYWRIGHT_MODULE || 'playwright');
const browser=await chromium.launch(); const errors=[],writes=[];let checks=0;
const eq=(a,b)=>{assert.deepEqual(a,b);checks++;};
try{
 for(const mode of [false,true]){
  const wb=new Workbook({date1904:mode,sheets:[{name:'날짜검증',colWidths:{0:230,1:180,2:140},cells:{'0,0':{raw:'0',style:{numFmt:'date'}},'1,0':{raw:'0.5',style:{numFmt:'custom',code:'yyyy-mm-dd hh:mm:ss'}},'2,0':{raw:'=-0.5',style:{numFmt:'custom',code:'[h]:mm'}},'3,0':{raw:'45322'},'0,1':{raw:'=YEAR(A1)'},'1,1':{raw:'=DATE(2024,2,29)'},'2,1':{raw:'=TEXT(A1,"yyyy-mm-dd")'}}}]});
  const ctx=await browser.newContext({viewport:{width:1440,height:1000}}),page=await ctx.newPage();page.on('pageerror',e=>errors.push(e.message));
  await ctx.route('**/*',route=>{if(!['GET','HEAD','OPTIONS'].includes(route.request().method())){writes.push(route.request().method());return route.abort();}return route.continue();});
  try{
   await page.addInitScript(()=>{window.TABULA_STATIC=true;window.WIXEL_SKIP_START=true;});await page.goto(process.env.WIXEL_URL||'http://127.0.0.1:5180/');await page.waitForFunction(()=>!!window.tabula?.wb());
   await page.locator('#fileInput').setInputFiles({name:'synthetic-date-system.xlsx',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',buffer:Buffer.from(writeXlsx(wb))});await page.waitForFunction(()=>window.tabula.wb().sheets[0].name==='날짜검증');
   const cell=(r,c)=>page.locator(`.c[data-r="${r}"][data-c="${c}"]`).first().innerText();
   eq(await cell(0,0),mode?'1904-01-01':'1900-01-00');eq(await cell(1,0),mode?'1904-01-01 12:00:00':'1900-01-00 12:00:00');eq(await cell(2,0),mode?'-12:00':'########');eq(await cell(3,0),'45322');
   eq(await page.evaluate(()=>window.tabula.wb().date1904),mode);
   await page.evaluate(()=>{const t=window.tabula,w=t.wb();w.transact(()=>w.setInput(0,4,0,'2024-02-29'));t.gv().renderAll();});eq(await page.evaluate(()=>window.tabula.wb().getValue(0,4,0)),mode?43889:45351);eq(await cell(4,0),'2024-02-29');
   const saved=await page.evaluate(()=>window.tabula.wb().serialize());const reopened=new Workbook(readXlsx(writeXlsx(new Workbook(saved))).data);eq(reopened.date1904,mode);eq(reopened.getValue(0,4,0),mode?43889:45351);
   await page.evaluate(()=>{window.tabula.wb().undo();window.tabula.gv().renderAll();});eq(await page.evaluate(()=>window.tabula.wb().getValue(0,4,0)),null);
  }finally{await ctx.close();}
 }
 eq(errors,[]);eq(writes,[]);console.log(JSON.stringify({ok:true,cases:2,checks,pageErrors:errors,blockedWrites:writes}));
}finally{await browser.close();}
