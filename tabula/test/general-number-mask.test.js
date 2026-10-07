import test from 'node:test';
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { zip,unzip,textOf } from '../src/zip.js';
import { readXlsx,readXlsxAsync,writeXlsx } from '../src/xlsx.js';
import { Workbook,cellData } from '../src/workbook.js';
import { snapshotPasteSource,applyPasteSpecial } from '../src/paste-special.js';
import { formatValue,formatGeneral,fmtCode } from '../src/format.js';
import { parseXml,child,kids } from '../src/xml.js';
import { NET,importRangeSource } from '../src/fx-web.js';
import { rangeImageSvg } from '../src/range-image-export.js';
const NS='xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
const formats=[{id:1,name:'0'},{id:2,name:'0.00'},{id:164,name:'code'},{id:12,name:'fraction'},{id:9,name:'percent'}];
function fixture(parent,scope) {
 const xfs=[0,parent,0].map((numFmtId,i)=>`<xf numFmtId="${numFmtId}" fontId="0" fillId="0" borderId="0" xfId="0"${i?' applyNumberFormat="1"':''}/>`).join('');
 const styles=`<styleSheet ${NS}><numFmts count="1"><numFmt numFmtId="164" formatCode="0.000&amp;quot;kg&amp;quot;"/></numFmts><fonts count="1"><font><name val="Calibri"/><sz val="11"/></font></fonts><fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3">${xfs}</cellXfs></styleSheet>`.replaceAll('&amp;quot;','&quot;');
 const rows=[2.5,1234.567].map((n,i)=>`<row r="${i+1}"${scope==='row'?' s="1" customFormat="1"':''}><c r="A${i+1}" s="0"><v>${n}</v></c><c r="B${i+1}" s="1"><v>${n}</v></c><c r="C${i+1}" s="2"><v>${n}</v></c></row>`).join('');
 return zip({'[Content_Types].xml':'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>','_rels/.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>','xl/workbook.xml':`<workbook ${NS}><sheets><sheet name="General" sheetId="1" r:id="r1"/></sheets></workbook>`,'xl/_rels/workbook.xml.rels':'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="r1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="r2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>','xl/styles.xml':styles,'xl/worksheets/sheet1.xml':`<worksheet ${NS}><dimension ref="A1:C2"/>${scope==='column'?'<cols><col min="1" max="3" style="1"/></cols>':''}<sheetData>${rows}</sheetData></worksheet>`});
}
const readers=[['DOM',bytes=>readXlsx(bytes)],['async DOM',bytes=>readXlsxAsync(bytes,null,{streamThreshold:1<<30})],['stream bytes',bytes=>readXlsxAsync(bytes,null,{streamThreshold:0})],['stream Blob',bytes=>readXlsxAsync(new Blob([bytes]),null,{streamThreshold:0})]];
function verify(book,label) {
 for(const [r,value]of [2.5,1234.567].entries())for(const c of [0,2]) {
  const st=book.styleAt(0,r,c);assert.equal(book.getValue(0,r,c),value,label);
  assert.equal(formatValue(value,st).text,String(value),label);assert.equal(fmtCode(st),null,label+' export General');
 }
 for(const [r,value]of [2.5,1234.567].entries())assert.notEqual(formatValue(value,book.styleAt(0,r,1)).text,String(value),label+' explicit parent format kept');
}
for(const [mode,read]of readers)for(const scope of ['column','row'])test(`${mode}: General 셀은 ${scope}의 0·0.00·code·fraction·percent를 차단하고 왕복한다`,async()=>{
 for(const parent of formats) {
  const label=mode+' '+scope+' '+parent.name,book=new Workbook((await read(fixture(parent.id,scope))).data);verify(book,label);
  const bytes=writeXlsx(book),files=unzip(bytes),root=parseXml(textOf(files['xl/styles.xml'])),xfs=kids(child(root,'cellXfs'),'xf'),sheet=parseXml(textOf(files['xl/worksheets/sheet1.xml']));
  for(const cell of kids(child(sheet,'sheetData'),'row').flatMap(row=>kids(row,'c')).filter(c=>/^[AC]/.test(c.attrs.r)))assert.equal(xfs[Number(cell.attrs.s??0)].attrs.numFmtId,'0',label+' OOXML General');
  verify(new Workbook((await read(bytes)).data),label+' roundtrip');
 }
});
test('null decimals는 일반 표시와 좁은 열의 자릿수 축소/지수 표시를 유지하며 명시0과 구별한다',()=>{
 assert.equal(formatValue(2.5,{numFmt:'',decimals:null}).text,'2.5');assert.equal(formatValue(2.5,{decimals:0}).text,'3');
 assert.equal(fmtCode({numFmt:'',decimals:null}),null);assert.equal(fmtCode({decimals:0}),'0');assert.equal(fmtCode({decimals:2}),'0.00');
 const source=readFileSync(new URL('../src/view.js',import.meta.url),'utf8'),body=/function fitNumber\([^]*?^}/m.exec(source)?.[0];assert.ok(body);
 const fit=vm.runInNewContext('('+body+')',{formatGeneral,measureText:text=>text.length*7});
 for(const [value,width]of [[1234.567,35],[1e12,35],[1234.567,2],[1e12,0]])assert.equal(fit(value,width,{numFmt:'',decimals:null}),fit(value,width,{}));
 assert.equal(fit(1e12,35,{numFmt:'',decimals:null}),'1E+12');assert.equal(fit(1e12,35,{decimals:0}),'#####');
});


test('General null mask는 JSON/복사 붙여넣기/서식지우기 UndoRedo에서 정수0으로 변하지 않는다',()=>{
 const book=new Workbook(readXlsx(fixture(1,'column')).data),loaded=new Workbook(JSON.parse(JSON.stringify(book.serialize())));verify(loaded,'JSON');
 const src=snapshotPasteSource(loaded,{si:0,r1:0,c1:0,r2:1,c2:0,data:[[cellData(loaded.getCell(0,0,0))],[cellData(loaded.getCell(0,1,0))]],values:[[2.5],[1234.567]]});
 assert.equal(src.data[0][0].style.decimals,null);loaded.transact(()=>{loaded.setLineStyle(0,'col',3,{decimals:0});applyPasteSpecial(loaded,0,src,{r1:0,c1:3,r2:1,c2:3});});
 const show=()=>formatValue(loaded.getValue(0,0,3),loaded.styleAt(0,0,3)).text;assert.equal(show(),'2.5');assert.equal(loaded.styleAt(0,0,3).decimals,null);
 loaded.undo();assert.equal(loaded.getValue(0,0,3),null);loaded.redo();assert.equal(show(),'2.5');assert.equal(loaded.styleAt(0,0,3).decimals,null);
 loaded.transact(()=>loaded.clearRange(0,0,3,1,3,'formats'));assert.equal(loaded.getCell(0,0,3).style,undefined);loaded.undo();assert.equal(show(),'2.5');assert.equal(loaded.styleAt(0,0,3).decimals,null);loaded.redo();loaded.undo();assert.equal(show(),'2.5');
 verify(new Workbook(readXlsx(writeXlsx(loaded)).data),'clipboard save');
});

test('General decimals null만 보존하며 다른 null/undefined 서식 제거 계약을 유지한다',()=>{
 const book=new Workbook();book.transact(()=>book.setCellData(0,0,0,{raw:'2.5',style:{decimals:null,bold:null,color:null,fill:undefined}}));
 assert.deepEqual(book.getCell(0,0,0).style,{decimals:null});assert.equal(formatValue(book.getValue(0,0,0),book.styleAt(0,0,0)).text,'2.5');
 book.transact(()=>book.setStyle(0,0,0,{bold:true}));assert.equal(book.styleAt(0,0,0).decimals,null);book.undo();assert.deepEqual(book.getCell(0,0,0).style,{decimals:null});book.redo();assert.equal(book.styleAt(0,0,0).bold,true);assert.equal(book.styleAt(0,0,0).decimals,null);
});


test('IMPORTRANGE 숫자 표시 힌트는 명시 General null mask·0을 덮지 않고 undefined에는 적용한다',()=>{
 const old={...NET},id='synthetic_general_mask_fixture',range='A1:A2';NET.cache=new Map();NET.authorize=null;NET.fetcher=()=>{throw Error('네트워크 금지');};NET.onDone=null;
 NET.cache.set(importRangeSource(id,range).url,{state:'ok',data:'값\n2.500',sheets:new Set(),t:Date.now()});
 try {
  for(const [decimals,expected]of [[undefined,'2.500'],[null,'2.5'],[0,'3']]) {
   const book=new Workbook();book.transact(()=>{if(decimals!==undefined)book.setCellData(0,1,0,{raw:'',style:{decimals}});book.setInput(0,0,0,'=IMPORTRANGE("'+id+'","'+range+'")');});
   assert.equal(book.getValue(0,1,0),2.5);assert.equal(formatValue(book.getValue(0,1,0),book.styleAt(0,1,0)).text,expected);
  }
 } finally {Object.assign(NET,old);}
});


test('범위 그림 내보내기도 null General의 좁은 열 소수 축소를 유지한다',()=>{
 const render=decimals=>{const book=new Workbook();book.sheets[0].defColW=50;book.setCellData(0,0,0,{raw:'1234.567',style:decimals===undefined?undefined:{decimals}});return rangeImageSvg(book,0,{r1:0,c1:0,r2:0,c2:0},{measureText:text=>String(text).length*7}).svg;};
 assert.equal(render(null),render(undefined));assert.match(render(null),/>1234[.]6</);assert.doesNotMatch(render(0),/>1234[.]6</);
});
