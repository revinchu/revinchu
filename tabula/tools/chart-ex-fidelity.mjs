// 합성 현대/3D/콤보 차트: XLSX 표준 옵션과 실제 Excel 열기 게이트용 생성기.
// node tools/chart-ex-fidelity.mjs [D:/Codex/Temp/wixel-chartex-fidelity] [--verify]
import assert from 'node:assert/strict';
import { mkdirSync, readFileSync, writeFileSync, existsSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, readXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { chartModelData } from '../src/chart.js';
const root=resolve(process.argv.find((x,i)=>i>1&&!x.startsWith('--'))??'D:/Codex/Temp/wixel-chartex-fidelity');
const verify=process.argv.includes('--verify');
const rows=[['상위','중간','항목','금액'],['가','공통','첫째',10],['가','공통','둘째',20],['나','공통','첫째',30],['나','','',7]];
const normal=[['분류','첫째','둘째','셋째'],['가',10,2,100],['나',20,3,200],['다',30,4,300]];
const cases=[...['banner','overlapping','none'].map(treemapLabelLayout=>({name:`treemap-${treemapLabelLayout}`,type:'treemap',treemapLabelLayout,chartType:117})),{name:'sunburst',type:'sunburst',chartType:120},...['box','cylinder','cone','pyramid'].map(barShape=>({name:`column-${barShape}`,type:'column',threeD:true,grouping:'standard',barShape})),...['columnLine','stackedColumnLine','areaColumn'].map(comboLayout=>({name:`combo-${comboLayout}`,type:'combo',comboLayout,comboAxis:'primary'})),{name:'label-column',type:'column',labelCase:true,chartType:51},{name:'label-pie',type:'pie',labelCase:true,chartType:5}];
mkdirSync(root,{recursive:true});
const manifest={format:'wixel-excel-interop',version:1,fixtures:[]};const results=[];
for(const item of cases){
  const modern=['treemap','sunburst'].includes(item.type), source=modern?rows:item.labelCase?normal.map(r=>r.slice(0,2)):normal;
  if(verify){
    const wb=new Workbook(readXlsx(readFileSync(join(root,'excel-saved',item.name+'.xlsx'))).data), ch=wb.sheets[0].charts[0], data=chartModelData(wb,0,ch);
    assert.equal(ch.type,item.type); assert.deepEqual(data.series.map(s=>s.values),modern?[[10,20,30,7]]:item.labelCase?[[10,20,30]]:[[10,20,30],[2,3,4],[100,200,300]]);
    if(modern){assert.deepEqual(data.categories,['첫째','둘째','첫째','']);assert.equal(ch.seriesFmt[0].catName,true);assert.equal(ch.seriesFmt[0].labels,true);assert.equal(ch.seriesFmt[0].labelSize,14);assert.equal(ch.seriesFmt[0].labelColor,'#123456');assert.equal(ch.seriesFmt[0].labelBold,true);assert.equal(ch.seriesFmt[0].labelPos,'insideEnd');assert.equal(ch.seriesFmt[0].pointColors[1],'#abcdef');assert.equal(ch.legend,'r');assert.equal(ch.palette[0],'#336699');if(item.type==='treemap')assert.equal(ch.treemapLabelLayout,item.treemapLabelLayout);}
    if(item.labelCase){for(const key of ['catName','serName','labelBold'])assert.equal(ch.seriesFmt[0][key],true);assert.equal(ch.seriesFmt[0].labels,false);assert.equal(ch.seriesFmt[0].labelColor,'#234567');assert.equal(ch.seriesFmt[0].labelSize,13);}
    if(item.barShape){assert.equal(ch.barShape,item.barShape);assert.equal(ch.grouping,'standard');}
    results.push({file:item.name+'.xlsx',type:ch.type,series:data.series.length,values:true});continue;
  }
  const wb=new Workbook();wb.transact(()=>source.forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v)))));
  wb.sheets[0].charts=[{id:'fidelity',x:0,y:120,w:640,h:400,title:'합성 차트',...item,range:{r1:0,c1:0,r2:source.length-1,c2:source[0].length-1},legend:'r',...(modern?{palette:['#336699','#cc6600','#339966'],seriesFmt:[{labels:true,catName:true,serName:false,labelSize:14,labelColor:'#123456',labelBold:true,labelPos:'insideEnd',pointColors:{1:'#abcdef'}}]}:item.labelCase?{seriesFmt:[{labels:false,catName:true,serName:true,pct:item.type==='pie',labelSize:13,labelBold:true,labelColor:'#234567'}]}:{})}];
  const files=unzip(writeXlsx(wb));for(const key of Object.keys(files))if(/^xl\/(drawings|charts)\/[^/]+\.xml$/.test(key))files[key]=textOf(files[key]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'').replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
  writeFileSync(join(root,item.name+'.xlsx'),zip(files));writeFileSync(join(root,item.name+'.xml'),files['xl/charts/chart1.xml']);
  manifest.fixtures.push({file:item.name+'.xlsx',...(item.chartType?{chartTypes:[item.chartType]}:{}),seriesCounts:[modern||item.labelCase?1:3],cells:[{address:item.labelCase?'B2':'D2',value:source[1][item.labelCase?1:3]}]});
}
if(!verify)writeFileSync(join(root,'synthetic-fixtures.json'),JSON.stringify(manifest,null,2));
// Native Excel-generated synthetic hierarchy files have no ChartEx caches: resolve named ranges.
for(const type of ['treemap','sunburst']){const file=join(root,`native-${type}.xlsx`);if(!existsSync(file))continue;const wb=new Workbook(readXlsx(readFileSync(file)).data),ch=wb.sheets[0].charts[0],d=chartModelData(wb,0,ch);assert.equal(ch.type,type);assert.deepEqual(d.series[0].values,[10,20,5,30,7,8]);assert.deepEqual(d.categories,['항목1','항목2','항목3','항목1','','']);assert.equal(d.catLevels.length,2);assert.equal(ch.seriesFmt[0].pointColors[1],'#abcdef');assert.equal(ch.seriesFmt[0].labels,true);assert.equal(ch.seriesFmt[0].catName,true);if(type==='treemap')assert.equal(ch.treemapLabelLayout,'overlapping');results.push({native:type,values:true,hierarchyLevels:3});}
writeFileSync(join(root,verify?'verify-results.json':'generation-results.json'),JSON.stringify(results,null,2));
console.log(JSON.stringify({mode:verify?'verify':'generate',fixtures:cases.length,native:results.filter(x=>x.native).length,passed:verify?results.length:undefined,root}));
