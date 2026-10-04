import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsxBlobAsync, writeXlsxToSink, readXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { pivotSourceData } from '../src/pivot.js';

function fixture() {
  const wb = new Workbook({ sheets: [
    { name: 'Source', cells: { '0,0': {raw:'Group'}, '0,1': {raw:'Value'}, '1,0': {raw:'가😀'}, '1,1': {raw:'10'}, '2,0': {raw:'나'}, '2,1': {raw:'20'} }, cellRuns:[[5,4,32000,{raw:'',style:{fill:'#abcdef',bold:true}}]] },
    { name: 'Report', cells: { '0,3': {raw:'=Source!B2+Source!B3',style:{italic:true}}, '1,3': {raw:'별도 문자열',comment:'메모',link:'https://example.invalid'} },
      pivot: { name:'Pivot',source:'Source',range:{r1:0,c1:0,r2:2,c2:1},rows:['Group'],values:[{field:'Value',agg:'sum'}],top:3,left:0,style:'PivotStyleDark2' },
      slicers: [{ id:'Slicer',caption:'한글 필터',source:{kind:'pivot',field:'Group',pivots:[{sheet:'Report',name:'Pivot'}]},x:20,y:20,w:160,h:200,style:'SlicerStyleDark2',showHeader:false }],
    },
  ] });
  wb.sheets[0].rowHeights[1] = 42; wb.sheets[0].rowManual[1] = true;
  wb.sheets[1].page = { area:{r1:0,c1:0,r2:9,c2:6},orientation:'landscape' };
  return wb;
}

for (const disk of [false,true]) for (const compress of [true,false]) test(`XLSX roundtrip keeps cells, edits, formats, pivot cache and slicer design: disk=${disk}, native=${compress}`, async () => {
  const wb = fixture(), before = JSON.stringify(wb.serialize()), original = globalThis.CompressionStream;
  let bytes;
  try {
    if (!compress) globalThis.CompressionStream = undefined;
    if (disk) {
      const chunks = []; let total = 0;
      const result = await writeXlsxToSink(wb, {}, { async write(chunk) { total += chunk.length; chunks.push(chunk.slice()); } });
      assert.equal(result.bytesWritten, total);
      bytes = new Uint8Array(await new Blob(chunks).arrayBuffer());
    } else {
      const blob = await writeXlsxBlobAsync(wb);
      assert.ok(blob instanceof Blob); assert.ok(blob.size > 0);
      bytes = new Uint8Array(await blob.arrayBuffer());
    }
  } finally { globalThis.CompressionStream = original; }
  const files = unzip(bytes);
  assert.ok(files['[Content_Types].xml']);
  if (!disk) assert.equal(Object.keys(files)[0], '[Content_Types].xml');
  const xml = textOf(files['xl/worksheets/sheet1.xml']);
  assert.ok(xml.length > 1 << 20);
  assert.match(xml, /<c r="E32005" s="\d+"\/>/);
  const back = new Workbook(readXlsx(bytes).data);
  assert.equal(back.sheets.length, 2);
  assert.equal(back.getRaw(0,1,0), '가😀');
  assert.equal(back.getRaw(1,0,3), '=Source!B2+Source!B3'); assert.equal(back.getValue(1,0,3), 30);
  assert.equal(back.styleAt(0,32004,4).fill, '#abcdef'); assert.equal(back.styleAt(1,0,3).italic, true);
  assert.equal(back.getCell(1,1,3).comment, '메모'); assert.equal(back.getCell(1,1,3).link, 'https://example.invalid');
  assert.equal(back.sheets[0].rowHeights[1], 42); assert.equal(back.sheets[1].page.area.r2, 9);
  assert.equal(back.sheets[1].slicers[0].style, 'SlicerStyleDark2'); assert.equal(back.sheets[1].slicers[0].showHeader, false);
  const cube = pivotSourceData(back, back.sheets[1].pivot).cube;
  assert.equal(cube.n, 2); assert.equal(cube.col(1).get(1), 20);
  assert.equal(JSON.stringify(wb.serialize()), before, 'export never changes workbook data');
});

test('Blob XLSX compresses the previous sheet before constructing the next sheet', async () => {
  const wb = new Workbook({sheets:[{name:'A',cells:{'0,0':{raw:'first'}}},{name:'B',cells:{'0,0':{raw:'second'}}}]});
  const original = globalThis.CompressionStream, getValue = wb.getValue.bind(wb); let compressors = 0;
  try {
    globalThis.CompressionStream = class { constructor(format) { compressors++; return new original(format); } };
    wb.getValue = (si,r,c) => { if (si === 1) assert.ok(compressors > 0, 'first sheet was already handed to ZIP'); return getValue(si,r,c); };
    const bytes = new Uint8Array(await (await writeXlsxBlobAsync(wb)).arrayBuffer());
    assert.equal(new Workbook(readXlsx(bytes).data).getRaw(1,0,0), 'second');
  } finally { globalThis.CompressionStream = original; }
});


test('wide formula rows stream individual cells without joining one oversized row string', async () => {
  const value='가😀'.repeat(2000), raw='="'+value+'"', cells={};
  for(let c=0;c<96;c++)cells[`0,${c}`]={raw};
  const wb=new Workbook({sheets:[{name:'Wide',cells}]}), join=Array.prototype.join;let bytes;
  try {
    Array.prototype.join=function(separator){
      if(separator==='') {let size=0;for(const part of this){if(typeof part!=='string')break;size+=part.length;if(size>(1<<20))throw Error('oversized row join');}}
      return join.call(this,separator);
    };
    bytes=new Uint8Array(await (await writeXlsxBlobAsync(wb)).arrayBuffer());
  } finally {Array.prototype.join=join;}
  const back=new Workbook(readXlsx(bytes).data);
  for(let c=0;c<96;c++){assert.equal(back.getRaw(0,0,c),raw);assert.equal(back.getValue(0,0,c),value);}
});


test('disk XLSX writes a sheet before producing the next and stops on disk error', async () => {
  const wb = new Workbook({sheets:[{name:'A',cells:{'0,0':{raw:'first'}}},{name:'B',cells:{'0,0':{raw:'second'}}}]}), getValue = wb.getValue.bind(wb);
  let writes = 0, reachedSecond = false;
  wb.getValue = (si,r,c) => { if (si === 1) { reachedSecond = true; assert.ok(writes > 0); } return getValue(si,r,c); };
  await writeXlsxToSink(wb, {}, {async write() { writes++; }});
  assert.equal(reachedSecond, true);
  reachedSecond = false;
  await assert.rejects(writeXlsxToSink(wb, {}, {async write() { throw Error('disk full'); }}), /disk full/);
  assert.equal(reachedSecond, false, 'disk failure must stop constructing later sheets');
});
