import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants, esc } from '../src/xml.js';
const PNG='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
const path='xl/drawings/drawing1.xml',rels='xl/drawings/_rels/drawing1.xml.rels';
function fixture(src=PNG) {
 const pic={id:'p',name:'picture',kind:'picture',src,originalSrc:src,x:10,y:20,w:40,h:50,crop:{l:.2},opacity:.6};
 const nested={id:'n',kind:'group',x:70,y:10,w:100,h:80,groupSize:{w:100,h:80},groupItems:[{...pic,id:'p2'}]};
 const group={id:'g',kind:'group',x:20,y:30,w:300,h:120,groupSize:{w:300,h:120},groupItems:[pic,nested]};
 return new Workbook({sheets:[{name:'One',cells:{},shapes:[group],images:[{...pic,id:'independent'}]},{name:'Two',cells:{},shapes:[{...group,id:'g2'}]}]});
}
test('nested groups store large shared originals once in package media, never in metadata JSON',()=>{
 const source='data:image/png;base64,'+Buffer.alloc(256*1024,17).toString('base64');
 const wb=fixture(source),bytes=writeXlsx(wb),files=unzip(bytes),xml=textOf(files[path]);
 assert.equal(Object.keys(files).filter(p=>p.startsWith('xl/media/')).length,1);
 assert.ok(xml.length<20000,'group XML stays small independently of image bytes');
 assert.ok(!xml.includes(source)); assert.match(xml,/<wx:media /);
 const back=readXlsx(bytes).data;
 assert.deepEqual(back.sheets[0].shapes[0].groupItems,wb.sheets[0].shapes[0].groupItems);
 assert.equal(back.sheets[1].shapes[0].groupItems[0].originalSrc,source);
 assert.equal(back.sheets[0].images[0].src,source);
});
test('renumbered native and metadata relationship IDs resolve the same group originals',()=>{
 const wb=fixture(),files=unzip(writeXlsx(wb));
 files[path]=textOf(files[path]).replaceAll('rId1','renumbered1');
 files[rels]=textOf(files[rels]).replaceAll('rId1','renumbered1');
 const out=readXlsx(zip(files)).data.sheets[0].shapes[0];
 assert.deepEqual(out.groupItems,wb.sheets[0].shapes[0].groupItems);
});
test('missing metadata image references fall back to standard native group pictures',()=>{
 const files=unzip(writeXlsx(fixture()));
 files[path]=textOf(files[path]).replace(/(<wx:media[^>]* r:embed=")[^"]+("\/>)/g,'$1missing$2');
 const result=readXlsx(zip(files)),out=result.data.sheets[0].shapes[0];
 assert.equal(out.kind,'group');assert.equal(out.groupItems[0].src,PNG);
 assert.equal(out.groupItems[1].groupItems[0].src,PNG);
 assert.match(result.warnings.join(' '),/외부에서 수정/);
});
test('legacy inline group metadata remains readable',()=>{
 const wb=fixture(),files=unzip(writeXlsx(wb));
 let i=0;const originalGroups=[wb.sheets[0].shapes[0],wb.sheets[0].shapes[0].groupItems[1]];
 files[path]=textOf(files[path]).replace(/<wx:group\b[^>]*>[\s\S]*?<\/wx:group>/g,tag=>{
   const node=parseXml(tag),g=originalGroups[i++];
   return '<wx:group xmlns:wx="https://wixel.app/drawing/group/1" signature="'+node.attrs.signature+'" json="'+esc(JSON.stringify({kind:'group',groupItems:g.groupItems,groupSize:g.groupSize}))+'"/>';
 });
 assert.equal(i,2);
 const out=readXlsx(zip(files)).data.sheets[0].shapes[0];
 assert.deepEqual(out.groupItems,wb.sheets[0].shapes[0].groupItems);
});
