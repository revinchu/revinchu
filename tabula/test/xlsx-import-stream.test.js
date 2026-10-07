import test from 'node:test';
import assert from 'node:assert/strict';
import {sheetXmlStream} from '../src/xlsx-sheet-stream.js';
import {scanXmlChildren} from '../src/xml-stream.js';
import {readSharedStrings} from '../src/shared-strings.js';
import {readPivotSnapshotXml} from '../src/pivot-cache-data.js';
import {readXlsx,readXlsxAsync,writeXlsx} from '../src/xlsx.js';
import {Workbook} from '../src/workbook.js';
import {unzip,zip,textOf} from '../src/zip.js';
const enc=new TextEncoder();
function* chunks(text,size=1){const b=enc.encode(text);for(let p=0;p<b.length;p+=size)yield b.subarray(p,p+size);}
function drain(g){for(;;){const s=g.next();if(s.done)return s.value;}}
test('worksheet stream preserves UTF-8, namespaced rows and before/after metadata at every byte boundary',()=>{
  const head='<?xml version="1.0"?><x:worksheet xmlns:x="main"><!-- <sheetData>fake</sheetData> --><x:cols><x:col min="1" max="1" width="12"/></x:cols>';
  const row='<x:row r="1"><x:c r="A1" t="inlineStr"><x:is><x:t>한글😀</x:t></x:is></x:c></x:row>';
  const tail='<x:mergeCells><x:mergeCell ref="A1:B1"/></x:mergeCells><x:drawing r:id="r1"/></x:worksheet>';
  for(const size of [1,2,7,31,65536]){const s=sheetXmlStream(chunks(head+'<x:sheetData>'+row+'<x:row r="2"/></x:sheetData>'+tail,size));assert.ok(s.head.includes('<x:cols>'));assert.deepEqual([...s.rows()],[row,'<x:row r="2"/>']);assert.equal(s.rest,head+'<x:sheetData/>'+tail);}
});
test('worksheet stream consumes EOF for CRC and closes generator on early cancellation or truncated row',()=>{
  let closed=false;function* source(){try{yield* chunks('<worksheet><sheetData><row r="1"/><row r="2"/></sheetData></worksheet>');throw Error('CRC sentinel');}finally{closed=true;}}
  assert.throws(()=>[...sheetXmlStream(source()).rows()],/CRC sentinel/);assert.equal(closed,true);
  closed=false;const it=sheetXmlStream(source()).rows();assert.equal(it.next().value,'<row r="1"/>');it.return();assert.equal(closed,true);
  assert.throws(()=>[...sheetXmlStream(chunks('<worksheet><sheetData><row><c>')).rows()],/끝까지/);
});
test('self-closing sheet data and CDATA/comment row terminators preserve metadata',()=>{
  const s=sheetXmlStream(chunks('<worksheet><sheetData/><mergeCells/></worksheet>'));assert.deepEqual([...s.rows()],[]);assert.equal(s.rest,'<worksheet><sheetData/><mergeCells/></worksheet>');
  const row='<row><c><![CDATA[</row>한글]]><!-- </row> --></c></row>';
  assert.deepEqual([...sheetXmlStream(chunks('<worksheet><sheetData>'+row+'</sheetData></worksheet>')).rows()],[row]);
});
test('XML item scanner accepts iterables and preserves every shared string and pivot record',()=>{
  const xml='<sst>'+Array.from({length:2050},(_,i)=>'<si><t>한글😀'+i+'</t></si>').join('')+'</sst>';
  const s=drain(readSharedStrings(chunks(xml,7)));assert.equal(s.strings.length,2050);assert.equal(s.strings[2049],'한글😀2049');
  const records='<pivotCacheRecords>'+Array.from({length:2100},(_,i)=>'<r><n v="'+i+'"/><s v="한글"/></r>').join('')+'</pivotCacheRecords>';
  const snapshot=drain(readPivotSnapshotXml(chunks(records,31),[{name:'n',db:true,shared:[]},{name:'s',db:true,shared:[]}],false,2100));assert.equal(snapshot.n,2100);assert.equal(snapshot.columns[0].num[2099],2099);
  assert.throws(()=>[...scanXmlChildren(chunks('<sst><si><t>incomplete'),'si')],/끝까지/);
});
test('bounded async byte/Blob import preserves borders, colors, sizes, merges, view, links and drawing geometry',async()=>{
  const wb=new Workbook({sheets:[{name:'서식',cells:{'0,0':{raw:'한글😀',style:{bold:true,size:18,color:'#123456',fill:'#abcdef',borderBottom:{style:'dashDot',color:'#f01020'},wrap:true},link:'https://example.test/'},'3,1':{raw:'7'},'3,2':{raw:'=B4*2',cached:14}},merges:[{r1:0,c1:0,r2:0,c2:1}],colWidths:{0:32,1:180},rowHeights:{0:64},zoom:75,images:[],shapes:[{id:'sh1',type:'rect',x:100,y:90,w:40,h:30,fill:'#ff0000',line:'#00ff00'}]}]});
  let bytes=writeXlsx(wb);const files=unzip(bytes);const path='xl/worksheets/sheet1.xml';let xml=textOf(files[path]);xml=xml.replace(/<(\/?)(worksheet|sheetData|row|c|v|f|is|t)(?=[\s/>])/g,'<$1x:$2').replace('<x:worksheet ','<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ');files[path]=enc.encode(xml);bytes=zip(files);
  const expected=new Workbook(readXlsx(bytes).data);
  for(const input of [bytes,new Blob([bytes])]){const res=await readXlsxAsync(input,null,{streamThreshold:0});const got=new Workbook(res.data);assert.equal(got.sheets[0].cells.size,expected.sheets[0].cells.size);for(const [r,c] of [[0,0],[3,1],[3,2]]){assert.equal(got.getValue(0,r,c),expected.getValue(0,r,c));assert.deepEqual(got.styleAt(0,r,c),expected.styleAt(0,r,c));}for(const key of ['colWidths','rowHeights','merges','zoom'])assert.deepEqual(got.sheets[0][key],expected.sheets[0][key]);assert.deepEqual(got.sheets[0].shapes.map(({id,...shape})=>shape),expected.sheets[0].shapes.map(({id,...shape})=>shape));const rt=new Workbook(readXlsx(writeXlsx(got)).data);assert.equal(rt.getValue(0,3,2),14);assert.deepEqual(rt.styleAt(0,0,0),got.styleAt(0,0,0));}
});
