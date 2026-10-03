import test from 'node:test';
import assert from 'node:assert/strict';
import {createScalarCellMemo} from '../src/scalar-cell-memo.js';
import {CellMap} from '../src/cellmap.js';
import {Workbook,cellData} from '../src/workbook.js';
import {readXlsx,writeXlsx} from '../src/xlsx.js';
const style={fill:'#ffeeaa',align:'left'};
function imported(list){const cells=new CellMap();list.forEach((value,r)=>cells.setRC(r,0,value));return new Workbook({sheets:[{name:'Data',cells}]});}
test('scalar memo shares only exact raw, style identity, input type and primitive value',()=>{
 const share=createScalarCellMemo(),a=share({raw:'same',style,v:'same'});
 assert.equal(share({raw:'same',style,v:'same'}),a);assert.ok(Object.isFrozen(a));
 for(const d of [{raw:'Same',style,v:'Same'},{raw:'same',style:{...style},v:'same'},{raw:'same',style,inputType:'text',v:'same'},{raw:'same',style,v:true}])assert.notEqual(share(d),a);
 const number=share({raw:'1',style,v:1});assert.notEqual(share({raw:'1',style,v:'1'}),number);
});
test('scalar memo leaves formulas, cached results, annotations, errors, blanks and date input mutable',()=>{
 const share=createScalarCellMemo();
 for(const d of [{raw:'=1',v:1,formula:true},{raw:'v',v:'v',cached:'v'},{raw:'v',v:'v',comment:'note'},{raw:'v',v:'v',link:'#A1'},{raw:'v',v:'v',phonetic:{}},{raw:'v',v:'v',image:{}},{raw:'#N/A',v:{error:'#N/A'}},{raw:'2024-01-01',v:45292},{raw:'',v:null,style}]){assert.equal(share(d),d);assert.ok(!Object.isFrozen(d));}
});
test('memo limits remain bounded but existing entries still share after saturation',()=>{
 const share=createScalarCellMemo({limit:3,perStyleLimit:2}),a=share({raw:'a',v:'a',style}),b=share({raw:'b',v:'b',style});
 const over=share({raw:'c',v:'c',style});assert.ok(!Object.isFrozen(over));
 const other={bold:true};assert.ok(Object.isFrozen(share({raw:'d',v:'d',style:other})));
 assert.ok(!Object.isFrozen(share({raw:'e',v:'e',style:other})));assert.equal(share({raw:'a',v:'a',style}),a);assert.equal(share({raw:'b',v:'b',style}),b);
 assert.notEqual(createScalarCellMemo()({raw:'a',v:'a',style}),a);
});
test('imported shared cells keep neighbors independent across edits, formatting, notes, links and undo',()=>{
 const wb=imported(Array.from({length:10},()=>({raw:'same',style})));assert.equal(wb.getCell(0,0,0),wb.getCell(0,9,0));assert.ok(Object.isFrozen(wb.getCell(0,0,0)));
 const actions=[()=>wb.setInput(0,3,0,'changed'),()=>wb.setStyle(0,3,0,{bold:true}),()=>wb.setComment(0,3,0,'note'),()=>wb.setCellData(0,3,0,{...cellData(wb.getCell(0,3,0)),link:'#A1'}),()=>wb.clearRange(0,3,0,3,0)];
 for(const change of actions){wb.transact(change);assert.equal(wb.getValue(0,4,0),'same');assert.deepEqual(cellData(wb.getCell(0,4,0)),{raw:'same',style});wb.undo();for(let r=0;r<10;r++)assert.equal(wb.getValue(0,r,0),'same');}
 wb.transact(()=>wb.insertRows(0,3,2));assert.equal(wb.getValue(0,5,0),'same');wb.undo();assert.equal(wb.sheets[0].cells.size,10);
});
test('date-string inputs remain unshared and change date-system safely while literal neighbors share',()=>{
 const wb=imported([{raw:'2024-01-01'},{raw:'2024-01-01'},{raw:'42',style},{raw:'42',style},{raw:'same',style},{raw:'same',style}]);
 assert.notEqual(wb.getCell(0,0,0),wb.getCell(0,1,0));assert.ok(!Object.isFrozen(wb.getCell(0,0,0)));assert.equal(wb.getCell(0,2,0),wb.getCell(0,3,0));
 const serial=wb.getValue(0,0,0);wb.transact(()=>wb.setDate1904(true));assert.equal(wb.getValue(0,0,0),serial);assert.equal(wb.getValue(0,1,0),serial);assert.equal(wb.getValue(0,2,0),42);wb.undo();assert.equal(wb.getCell(0,0,0).raw,'2024-01-01');
});
test('type and style differences survive sharing, local restore, and XLSX roundtrip',async()=>{
 const textStyle={numFmt:'text',align:'left'};
 const wb=imported([{raw:'42',style:textStyle},{raw:'42',style:textStyle,inputType:'value'},{raw:'42',style:textStyle},{raw:'FALSE',style},{raw:'FALSE',style}]);
 assert.equal(wb.getCell(0,0,0),wb.getCell(0,2,0));assert.notEqual(wb.getCell(0,0,0),wb.getCell(0,1,0));
 const expected=['42',42,'42',false,false];
 for(const next of [wb,new Workbook(wb.serialize()),new Workbook(JSON.parse(await wb.serializeBlob().text())),new Workbook(readXlsx(writeXlsx(wb)).data)])for(let r=0;r<expected.length;r++){assert.equal(next.getValue(0,r,0),expected[r]);assert.equal(next.styleAt(0,r,0).align,'left');}
});
