// 실제 Excel 일반 열기·재저장용 합성 자료만 생성합니다. 업무 파일은 사용하지 않습니다.
import { mkdirSync, writeFileSync } from 'node:fs';
import { resolve, join } from 'node:path';
import { Workbook } from '../src/workbook.js';
import { writeXlsx } from '../src/xlsx.js';
const out=process.argv[2]; if(!out) throw new Error('사용법: node tools/excel-fixtures.mjs D:/Codex/Temp/wixel-excel-fixtures');
mkdirSync(out,{recursive:true}); const fixtures=[];
function save(file,wb,expected={}) {writeFileSync(join(out,file+'.xlsx'),writeXlsx(wb));fixtures.push({file:file+'.xlsx',...expected});}
function chart(name,options,rows,excelType){
 const wb=new Workbook();wb.transact(()=>rows.forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v)))));
 wb.setSheetProp(0,'charts',[{id:'interop',title:'합성 검증',x:40,y:60,w:600,h:360,range:{r1:0,c1:0,r2:rows.length-1,c2:rows[0].length-1},...options}]);
 save(name,wb,{chartTypes:[excelType],shape:{width:450,height:270}});
}
const rows=[['분류','금액','이익'],['한국',10,2],['일본',5,4],['미국',7,6]];
for(const [type,excelType]of[['waterfall',119],['funnel',123],['histogram',118],['pareto',122],['treemap',117],['sunburst',120],['boxWhisker',121],['map',140]]){
 const r=['treemap','sunburst'].includes(type)?[['지역','국가','금액'],['아시아','한국',10],['아시아','일본',5],['미주','미국',7]]:type==='boxWhisker'?rows:rows.map(r=>r.slice(0,2));
 chart(type,{type},r,excelType);
}
for(const [surfaceStyle,excelType]of[['surface',83],['wireframe',84],['contour',85],['wireframeContour',86]])chart(surfaceStyle,{type:'surface',surfaceStyle,threeD:!surfaceStyle.toLowerCase().includes('contour')},[['항목','A','B','C'],['행1',5,20,9],['행2',10,4,30]],excelType);
for(const [type,excelType]of[['pieOfPie',68],['barOfPie',71]])chart(type,{type,splitType:'position',splitPos:2},[['항목','값'],['A',30],['B',25],['C',5],['D',10]],excelType);
for(const ohlc of[false,true])chart(ohlc?'volumeOHLC':'volumeHLC',{type:'stock',volume:true,ohlc},ohlc?[['일','거래량','시가','고가','저가','종가'],['월',1000,20,30,10,25],['화',2000,25,35,15,30]]:[['일','거래량','고가','저가','종가'],['월',1000,30,10,25],['화',2000,35,15,30]],ohlc?91:90);
const src='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/a9sAAAAASUVORK5CYII=';
const image={id:'im',src,name:'그림효과',alt:'서식 왕복',x:40,y:60,w:240,h:120,opacity:.35,radius:8,shadow:{dx:-4,dy:5,blur:6,opacity:.25,color:'#336699'},border:'#112233',borderW:.75};
save('picture-effects',new Workbook({sheets:[{name:'그림검증',cells:{},images:[image]}]}),{shape:{width:180,height:90},picture:{opacity:.35,radius:8,borderW:.75,shadow:image.shadow}});
for(const date1904 of[false,true]){
 const wb=new Workbook({date1904,sheets:[{name:'날짜검증',cells:{'0,0':{raw:'=DATE(2024,2,29)'},'1,0':{raw:'=TEXT(0,"yyyy-mm-dd")'},'2,0':{raw:'=YEAR(0)'},'3,0':{raw:'0'},'4,0':{raw:'45322.5',style:{numFmt:'yyyy-mm-dd hh:mm:ss'}},'5,0':{raw:'=SUM(1,2,3)'}}}]});
 save(date1904?'date1904':'date1900',wb,{date1904,cells:[{address:'A1',value:date1904?43889:45351},{address:'A2',value:date1904?'1904-01-01':'1900-01-00'},{address:'A3',value:date1904?1904:1900},{address:'A4',value:0},{address:'A5',value:45322.5},{address:'A6',value:6}]});
}
const wb=new Workbook();wb.transact(()=>[['구분','금액'],['A',10],['B',20],['C',30]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v)))));wb.setSheetProp(0,'tables',[{id:'table',name:'합성표',r1:0,c1:0,r2:3,c2:1,header:true,totals:false,style:'TableStyleMedium2',banded:true}]);
save('table',wb,{tables:1,cells:[{address:'B2',value:10},{address:'B4',value:30}]});
writeFileSync(join(out,'synthetic-fixtures.json'),JSON.stringify({format:'wixel-excel-interop',version:1,fixtures},null,2));
console.log(JSON.stringify({path:resolve(out),fixtures:fixtures.length}));
