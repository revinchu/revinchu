import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { shapeSvg } from '../src/shapes.js';
import { ungroupObjects } from '../src/object-group.js';
const PNG='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
const path='xl/drawings/drawing1.xml';
const native=bytes=>{const f=unzip(bytes);f[path]=textOf(f[path]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'');return zip(f)};
const fixture=()=>new Workbook({sheets:[{name:'Sheet1',cells:{},shapes:[{
 id:'outer',name:'outer',kind:'group',x:50,y:60,w:600,h:180,rot:17,flip:true,flipV:true,alt:'native group',groupSize:{w:300,h:180},groupItems:[
  {id:'text',name:'text',kind:'roundRect',adjustments:{adj:1746},x:10,y:20,w:90,h:40,rot:30,text:'Native text',font:'Arial',size:18,stroke:'#112233',strokeWidth:2},
  {id:'pic',name:'pic',kind:'picture',src:PNG,x:120,y:15,w:60,h:75,crop:{l:.2},opacity:.7},
  {id:'nested',name:'nested',kind:'group',x:180,y:95,w:90,h:70,rot:12,flip:true,groupSize:{w:45,h:35},groupItems:[{id:'inner',name:'inner',kind:'ellipse',x:5,y:6,w:25,h:18,text:'Nested'}]},
 ]},{id:'separate',name:'separate',kind:'rect',x:900,y:30,w:50,h:50,text:'Separate'}]}]});
const geometry=o=>({kind:o.kind,x:o.x,y:o.y,w:o.w,h:o.h,rot:o.rot??0,flip:!!o.flip,flipV:!!o.flipV,...(o.groupItems?{groupSize:o.groupSize,items:o.groupItems.map(geometry)}:{})});
test('metadata-free Excel groups preserve nesting, local coordinates and nonuniform scale',()=>{
 const wb=fixture(),before=structuredClone(wb.sheets[0].shapes),out=readXlsx(native(writeXlsx(wb))).data.sheets[0];
 assert.equal(out.shapes.length,2);assert.equal(out.images.length,0);assert.deepEqual(geometry(out.shapes[0]),geometry(before[0]));assert.equal(out.shapes[1].name,'separate');
 const group=out.shapes[0],text=group.groupItems[0],pic=group.groupItems[1];
 assert.equal(text.size,18);assert.equal(text.font,'Arial');assert.equal(text.strokeWidth,2);assert.deepEqual(text.adjustments,{adj:1746});assert.equal(pic.crop.l,.2);assert.equal(pic.opacity,.7);
 assert.deepEqual(ungroupObjects(group).map(geometry),ungroupObjects(before[0]).map(geometry));
 assert.deepEqual(wb.sheets[0].shapes,before);assert.match(shapeSvg(group),/viewBox="0 0 300 180"/);assert.equal(text.text,'Native text');assert.match(shapeSvg(group),/<text /);
});
test('native group coordinate origins normalize once and survive a standard XML roundtrip',()=>{
 const wb=fixture(),f=unzip(native(writeXlsx(wb)));let xml=textOf(f[path]);
 // Shift the outer coordinate origin and direct child offsets equally, leaving nested local coordinates alone.
 const original='<a:chOff x="0" y="0"/>';
 xml=xml.replace(original,'<a:chOff x="952500" y="1905000"/>');
 for(const [x,y] of [[10,20],[120,15],[180,95]])xml=xml.replace(`<a:off x="${x*9525}" y="${y*9525}"/>`,`<a:off x="${(x+100)*9525}" y="${(y+200)*9525}"/>`);
 f[path]=xml;const first=readXlsx(zip(f)).data;
 assert.deepEqual(geometry(first.sheets[0].shapes[0]),geometry(wb.sheets[0].shapes[0]));
 const second=readXlsx(native(writeXlsx(new Workbook(first)))).data;
 assert.deepEqual(geometry(second.sheets[0].shapes[0]),geometry(first.sheets[0].shapes[0]));
 assert.equal(second.sheets[0].shapes[0].groupItems[0].size,18);
});
test('native group membership and independent top-level objects persist through Undo and Redo',()=>{
 const wb=new Workbook(readXlsx(native(writeXlsx(fixture()))).data),group=structuredClone(wb.sheets[0].shapes[0]);
 wb.transact(()=>wb.setSheetProp(0,'shapes',[{...group,x:150},wb.sheets[0].shapes[1]]));
 wb.undo();assert.equal(wb.sheets[0].shapes[0].x,50);assert.deepEqual(wb.sheets[0].shapes[0].groupItems,group.groupItems);
 wb.redo();assert.equal(wb.sheets[0].shapes[0].x,150);assert.equal(wb.sheets[0].shapes[1].x,900);
});

test('collapsed native Excel groups retain zero extents and editable children across saves',()=>{
 for (const dim of ['w','h']) {
  const wb=fixture(),g=wb.sheets[0].shapes[0];g[dim]=0;
  const bytes=native(writeXlsx(wb)),out=readXlsx(bytes).data;
  assert.equal(out.sheets[0].shapes.length,2);assert.equal(out.sheets[0].images.length,0);
  const group=out.sheets[0].shapes[0];assert.equal(group[dim],0);
  assert.deepEqual(geometry(group),geometry(g));
  assert.ok(!/NaN|Infinity/.test(shapeSvg(group)));
  const second=readXlsx(writeXlsx(new Workbook(out))).data.sheets[0].shapes[0];
  assert.deepEqual(geometry(second),geometry(group));
 }
});
