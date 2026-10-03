import test from 'node:test';
import assert from 'node:assert/strict';
import {Workbook} from '../src/workbook.js';
import {readXlsx,readXlsxAsync,writeXlsx,xlsxExportWarnings,mergeXlsxImportWarnings} from '../src/xlsx.js';
import {unzip,zip,textOf} from '../src/zip.js';
const CODE='olapDataModelValuesOnly';
function fixture(){return unzip(writeXlsx(new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'Key'},'0,1':{raw:'Value'},'1,0':{raw:'A'},'1,1':{raw:'10'}}},{name:'Report',cells:{'0,0':{raw:'Saved'},'0,1':{raw:'42'}},pivot:{name:'P',source:'Source',range:{r1:0,c1:0,r2:1,c2:1},rows:['Key'],values:[{field:'Value',agg:'sum'}],top:0,left:0}}]})));}
function assertWarning(result){
 assert.ok(result.data.props.xlsxImportWarnings.includes(CODE));
 assert.ok(result.warnings.some(w=>w.includes('OLAP')&&w.includes('DAX')&&w.includes('보존되지')));
 const book=new Workbook(result.data);assert.equal(book.getValue(1,0,1),42);
 assert.ok(xlsxExportWarnings(book).some(w=>w.includes('OLAP')));return book;
}
test('embedded data model reports explicit loss without decoding the binary model',async()=>{
 const files=fixture();files['xl/model/item.data']=new Uint8Array([255,0,255]);
 const bytes=zip(files);assertWarning(readXlsx(bytes));assertWarning(await readXlsxAsync(bytes));
});
test('external OLAP cache warns even without an embedded model and preserves saved cell results',()=>{
 const files=fixture(),path='xl/pivotCache/pivotCacheDefinition1.xml';
 files[path]=textOf(files[path]).replace(/<cacheSource[^>]*>[\s\S]*?<\/cacheSource>/,'<cacheSource type="external" connectionId="1"/>').replace('</pivotCacheDefinition>','<cacheHierarchies count="0"/></pivotCacheDefinition>');
 const result=readXlsx(zip(files)),book=assertWarning(result);assert.equal(book.sheets[1].pivot,null);
});
test('OLAP slicer metadata alone records the unsupported model warning',()=>{
 const files=fixture(),rels='xl/_rels/workbook.xml.rels';
 files[rels]=textOf(files[rels]).replace('</Relationships>','<Relationship Id="rOlapCache" Type="http://schemas.microsoft.com/office/2007/relationships/slicerCache" Target="slicerCaches/olap.xml"/></Relationships>');
 files['xl/slicerCaches/olap.xml']='<slicerCacheDefinition xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" name="Olap" sourceName="[Model].[Key]"><data><olap pivotCacheId="7"/></data></slicerCacheDefinition>';
 assertWarning(readXlsx(zip(files)));
});
test('OLAP conversion warning survives local serialization and XLSX export history',()=>{
 const files=fixture();files['xl/model/item.data']=new Uint8Array([1]);const loaded=assertWarning(readXlsx(zip(files)));
 const restored=new Workbook(JSON.parse(JSON.stringify(loaded.serialize())));assert.ok(xlsxExportWarnings(restored).some(w=>w.includes('OLAP')));
 const exported=writeXlsx(restored);assert.equal(Object.hasOwn(unzip(exported),'xl/model/item.data'),false);assertWarning(readXlsx(exported));
});
test('warning merge preserves all known codes and rejects arbitrary document strings',()=>{
 assert.deepEqual(mergeXlsxImportWarnings(undefined,{xlsxImportWarnings:[CODE,'untrusted <img>',CODE]},{xlsxImportWarnings:['dataTableValuesOnly']}),[CODE,'dataTableValuesOnly']);
 assert.deepEqual(mergeXlsxImportWarnings({xlsxImportWarnings:'bad'}),[]);
 assert.ok(!readXlsx(zip(fixture())).warnings.some(w=>w.includes('OLAP')));
});
