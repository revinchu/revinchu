// 합성 XLSX 안전성 fixture만 생성한다. 사용자 파일은 읽거나 수정하지 않는다.
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx, xlsxExportWarnings } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
const dir = resolve(process.argv[2] || 'D:/Codex/Temp/wixel-xlsx-safety');
mkdirSync(dir, {recursive:true});
const manifest = {kind:'wixel-xlsx-safety',synthetic:true,fixtures:[]};
for (const type of ['column','bar','line','scatter','bubble']) {
  const wb = new Workbook();
  const axes = {y:{logBase:10,min:1,max:1000}};
  if (['scatter','bubble'].includes(type)) axes.x={logBase:2,min:2,max:256};
  else axes.y2={logBase:2,min:1,max:256};
  wb.sheets[0].charts = [{id:'c',type,x:20,y:30,w:520,h:340,axes,series:[{name:{text:'Primary'},cache:[1,10,100],xCache:[2,20,200],sizeCache:[5,10,15]},...(!['scatter','bubble'].includes(type)?[{name:{text:'Secondary'},cache:[2,20,200]}]:[])],...(!['scatter','bubble'].includes(type)?{seriesFmt:[{}, {axis:'secondary'}]}:{})}];
  const name=`log-${type}.xlsx`, files=unzip(writeXlsx(wb));
  files['xl/charts/chart1.xml']=textOf(files['xl/charts/chart1.xml']).replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
  writeFileSync(resolve(dir,name),zip(files));
  manifest.fixtures.push({file:name,axes});
}
const table=new Workbook({sheets:[{name:'Synthetic',cells:{'0,0':{raw:'2'},'0,1':{raw:'=A1*2'},'1,0':{raw:'2'},'1,1':{raw:'4'},'2,0':{raw:'3'},'2,1':{raw:'6'}}}]});
const data=unzip(writeXlsx(table));data['xl/worksheets/sheet1.xml']=textOf(data['xl/worksheets/sheet1.xml']).replace(/(<c\b[^>]*r="B2"[^>]*>)/,'$1<f t="dataTable" ref="B2:B3" r1="A1"/>');
writeFileSync(resolve(dir,'data-table-input.xlsx'),zip(data));
const imported=new Workbook(readXlsx(zip(data)).data);writeFileSync(resolve(dir,'data-table-values.xlsx'),writeXlsx(imported));
manifest.fixtures.push({file:'data-table-values.xlsx',valuesOnly:true,warnings:xlsxExportWarnings(imported)});
writeFileSync(resolve(dir,'synthetic-fixtures.json'),JSON.stringify(manifest,null,2));
console.log(JSON.stringify({directory:dir,fixtures:manifest.fixtures.length,synthetic:true}));
