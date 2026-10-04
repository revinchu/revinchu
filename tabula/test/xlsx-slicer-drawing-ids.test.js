import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { writeXlsx, writeXlsxBlobAsync, readXlsx } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, child, descendants } from '../src/xml.js';

function fixture() {
  const wb = new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'Group'},'0,1':{raw:'Value'},'1,0':{raw:'A'},'1,1':{raw:'1'},'2,0':{raw:'B'},'2,1':{raw:'2'}}}]});
  wb.sheets[0].tables=[{id:'t1',name:'Data',r1:0,c1:0,r2:2,c2:1,header:true}];
  wb.sheets[0].pivot={name:'Pivot',source:'Source',range:{r1:0,c1:0,r2:2,c2:1},rows:['Group'],values:[{field:'Value',agg:'sum'}],top:5,left:0};
  wb.sheets[0].slicers=[
    {id:'t',caption:'Table',source:{kind:'table',table:'Data',column:'Group'},x:200,y:0,w:150,h:200},
    {id:'p',caption:'Pivot',source:{kind:'pivot',field:'Group',pivots:[{sheet:'Source',name:'Pivot'}]},x:370,y:0,w:150,h:200},
  ];
  return wb;
}

for (const blob of [false,true]) test(`slicer drawing identities stay unique in modern and fallback Excel branches: Blob=${blob}`,async()=>{
  const wb=fixture(),bytes=blob?new Uint8Array(await(await writeXlsxBlobAsync(wb)).arrayBuffer()):writeXlsx(wb);
  const root=parseXml(textOf(unzip(bytes)['xl/drawings/drawing1.xml']));
  const alternatives=descendants(root,'AlternateContent');assert.equal(alternatives.length,2);
  for(const branch of ['Choice','Fallback']) {
    const objects=alternatives.map(ac=>descendants(child(ac,branch),'cNvPr')[0]);
    assert.equal(new Set(objects.map(node=>node.attrs.id)).size,2,'IDs cannot collide when Excel selects either compatibility branch');
    for(const node of objects){assert.ok(Number(node.attrs.id)>0);assert.ok(node.attrs.name);}
  }
  for(const ac of alternatives) assert.deepEqual(descendants(child(ac,'Fallback'),'cNvPr')[0].attrs,descendants(child(ac,'Choice'),'cNvPr')[0].attrs);
  assert.equal(readXlsx(bytes).data.sheets[0].slicers.length,2);
});
