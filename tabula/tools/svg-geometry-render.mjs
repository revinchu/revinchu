// Browser raster oracle: original SVG vs editable WIXEL geometry (no app/user files).
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { gunzipSync } from 'node:zlib';
import { svgToEditableGroup } from '../src/svg-to-shapes.js';
import { shapeSvg } from '../src/shapes.js';
const { chromium }=await import(process.env.PLAYWRIGHT_MODULE||'playwright');
const library=JSON.parse(gunzipSync(readFileSync(new URL('../assets/iconlib.json.gz',import.meta.url))));
const fixtures=[
  '<svg viewBox="0 0 100 100"><path fill-rule="evenodd" d="M0 0H100V100H0Z M20 20H80V80H20Z" fill="red"/><circle cx="50" cy="50" r="10" fill="blue"/></svg>',
  '<svg viewBox="0 0 100 100"><g transform="translate(50 50) rotate(30) scale(-1 1) translate(-50 -50)"><rect x="15" y="20" width="30" height="40" rx="4" fill="#00aaff"/><path d="M50 30q20-10 30 20t-20 30" stroke="green" stroke-width="2" fill="none"/></g></svg>',
  '<svg viewBox="0 0 100 100"><path d="M5 5H45V45H5Z M55 5H95V45H55Z M15 15V35H35V15Z" fill="#123456"/><rect x="25" y="65" width="50" height="15" fill="none" stroke="red" stroke-width="2"/></svg>',
  '<svg viewBox="0 0 100 100"><rect x="10" y="10" width="30" height="80" fill="rgba(255,0,0,.5)" fill-opacity=".5" opacity=".5"/><rect x="60" y="10" width="30" height="80" fill="none" opacity=".5" stroke="rgba(0,0,255,.5)" stroke-opacity=".8" stroke-width="2"/></svg>',
];
for(const [,icons] of library.slice(0,24)){const [,body,vb='0 0 96 96']=icons.find(([,b])=>b.includes('fill-rule'))??icons[0];fixtures.push(`<svg viewBox="${vb}" fill="#000000">${body.replace(/\sfill="(?!none)[^"]*"/g,'').replace(/fill:\s*#[0-9a-fA-F]{3,8};?/g,'')}</svg>`);}
const browser=await chromium.launch();let tested=0,worst=0;
try {
  const page=await browser.newPage();await page.route('**/*',route=>route.abort());
  for(const [index,raw] of fixtures.entries()){
    const original=raw.replace('<svg ','<svg xmlns="http://www.w3.org/2000/svg" '),converted=shapeSvg(svgToEditableGroup(original,{w:100,h:100}));
    const result=await page.evaluate(async ({original,converted})=>{
      const read=async text=>{const img=new Image();img.src='data:image/svg+xml;base64,'+btoa(unescape(encodeURIComponent(text)));await img.decode();const c=document.createElement('canvas');c.width=c.height=300;const x=c.getContext('2d');x.fillStyle='white';x.fillRect(0,0,300,300);x.drawImage(img,0,0,300,300);return x.getImageData(0,0,300,300).data;};
      const a=await read(original),b=await read(converted);let mismatch=0,solid=0;
      const same=(v,i,j)=>Math.abs(v[i]-v[j])+Math.abs(v[i+1]-v[j+1])+Math.abs(v[i+2]-v[j+2])<12;
      for(let y=2;y<298;y++)for(let x=2;x<298;x++){const i=(y*300+x)*4,neighbors=[i-8,i+8,i-2400,i+2400];if(!neighbors.every(j=>same(a,i,j)&&same(b,i,j)))continue;solid++;if(Math.abs(a[i]-b[i])+Math.abs(a[i+1]-b[i+1])+Math.abs(a[i+2]-b[i+2])>30)mismatch++;}
      return {mismatch,solid};
    },{original,converted});
    assert.ok(result.mismatch<=8,`fixture ${index}: interior pixels differ ${JSON.stringify(result)}`);worst=Math.max(worst,result.mismatch);tested++;
  }
  console.log(JSON.stringify({fixtures:tested,passed:tested,worstInteriorPixelDifference:worst,externalRequests:'blocked'},null,2));
} finally {await browser.close();}
