// Local bundled icons only. Report supported/refused conversions; no network.
import { readFileSync } from 'node:fs';
import { gunzipSync } from 'node:zlib';
import { svgToEditableGroup } from '../src/svg-to-shapes.js';
const library=JSON.parse(gunzipSync(readFileSync(new URL('../assets/iconlib.json.gz',import.meta.url))));
let total=0,converted=0,maxParts=0,totalMs=0,maxMs=0;const refused=new Map();
for(const [,icons] of library)for(const [,body,vb='0 0 96 96'] of icons){
  total++;const start=performance.now();
  try {
    const source=`<svg xmlns="http://www.w3.org/2000/svg" viewBox="${vb}" fill="#000000">${body.replace(/\sfill="(?!none)[^"]*"/g,'').replace(/fill:\s*#[0-9a-fA-F]{3,8};?/g,'')}</svg>`;
    const g=svgToEditableGroup(source,{w:96,h:96});converted++;maxParts=Math.max(maxParts,g.groupItems.length);
  }catch(error){refused.set(error.message,(refused.get(error.message)??0)+1);}
  const ms=performance.now()-start;totalMs+=ms;maxMs=Math.max(maxMs,ms);
}
console.log(JSON.stringify({total,converted,maxParts,totalMs:Math.round(totalMs),maxMs:Math.round(maxMs),refused:[...refused].map(([reason,count])=>({reason,count}))},null,2));
