// Excel이 재저장한 합성 XLSX를 다시 WIXEL에서 검사합니다.
import assert from 'node:assert/strict';
import {readFileSync}from'node:fs';import{join}from'node:path';
import {readXlsx}from'../src/xlsx.js';import{Workbook}from'../src/workbook.js';import{parseRangeName}from'../src/formula.js';
const root=process.argv[2];if(!root)throw new Error('합성 fixture 폴더를 지정하세요.');
const manifest=JSON.parse(readFileSync(join(root,'synthetic-fixtures.json'),'utf8'));let checks=0;
for(const c of manifest.fixtures){
 const wb=new Workbook(readXlsx(readFileSync(join(root,'excel-saved',c.file))).data);
 if(c.date1904!==undefined){assert.equal(wb.date1904,c.date1904);checks++;}
 for(const cell of c.cells??[]){const a=parseRangeName(cell.address);assert.equal(wb.getValue(0,a.r1,a.c1),cell.value,c.file+'!'+cell.address);checks++;}
 if(c.picture){const p=wb.sheets[0].images[0];for(const k of['opacity','radius','borderW']){assert.ok(Math.abs(p[k]-c.picture[k])<.001,c.file+': '+k);checks++;}for(const k of['dx','dy','blur','opacity']){assert.ok(Math.abs(p.shadow[k]-c.picture.shadow[k])<.001,c.file+': shadow.'+k);checks++;}assert.equal(p.shadow.color,c.picture.shadow.color);checks++;}
 if(c.chartTypes){assert.equal(wb.sheets[0].charts.length,c.chartTypes.length);checks++;}
}
console.log(JSON.stringify({ok:true,files:manifest.fixtures.length,checks}));
