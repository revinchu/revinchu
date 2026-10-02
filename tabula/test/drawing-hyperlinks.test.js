import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants, child, esc } from '../src/xml.js';
const PNG='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jRZkAAAAASUVORK5CYII=';
const DRAW='xl/drawings/drawing1.xml', RELS='xl/drawings/_rels/drawing1.xml.rels';
const link={target:"#'월간 보고서'!A4",tooltip:'바로가기 <설명> & "내용"'};
const shape=(id,kind='rect',hyperlink=link)=>({id,name:id,kind,x:10,y:20,w:120,h:60,text:'바로가기',fill:'#4472c4',hyperlink});
const save=files=>readXlsx(zip(files));
const xmlSet=(files,path,xml)=>{files[path]=new TextEncoder().encode(xml);};
function addNativeLink(files,name,model,rid='rIdOriginal99') {
 const xml=textOf(files[DRAW]);const re=new RegExp('(<xdr:cNvPr\\b[^>]*name="'+name+'"[^>/]*)(?:/>|></xdr:cNvPr>)');
 assert.match(xml,re);xmlSet(files,DRAW,xml.replace(re,'$1><a:hlinkClick r:id="'+rid+'" tooltip="'+esc(model.tooltip??'')+'"/></xdr:cNvPr>'));
 const rel='<Relationship Id="'+rid+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="'+esc(model.target)+'"'+(model.targetMode?' TargetMode="'+model.targetMode+'"':'')+'/>';
 xmlSet(files,RELS,textOf(files[RELS]).replace('</Relationships>',rel+'</Relationships>'));
}

test('native drawing links without TargetMode retain raw sheet fragments on shapes/connectors/pictures',()=>{
 const wb=new Workbook();wb.sheets[0].shapes=[shape('button'),shape('connector','line')];wb.sheets[0].images=[{id:'picture',name:'picture',x:10,y:100,w:50,h:50,src:PNG}];
 for(const o of wb.sheets[0].shapes)delete o.hyperlink;
 const files=unzip(writeXlsx(wb));let i=0;for(const name of ['button','connector','picture'])addNativeLink(files,name,link,'rIdNative'+(++i));
 const out=save(files).data.sheets[0];for(const o of [...out.shapes,...out.images])assert.deepEqual(o.hyperlink,link);
});

test('shape, connector and picture hyperlinks use standard relationships and survive save/reopen',()=>{
 const wb=new Workbook();const external={target:'https://example.invalid/report?a=1&b="값"',targetMode:'External',tooltip:'웹 보고서'};
 wb.sheets[0].shapes=[shape('button'),shape('connector','line',external)];wb.sheets[0].images=[{id:'p',name:'p',x:1,y:2,w:50,h:50,src:PNG,hyperlink:link,linked:{sheet:'Sheet1',r1:0,c1:0,r2:1,c2:1}}];
 const before=structuredClone(wb.sheets[0].shapes),bytes=writeXlsx(wb),files=unzip(bytes),drawing=parseXml(textOf(files[DRAW])),rels=parseXml(textOf(files[RELS]));
 const targets=descendants(rels,'Relationship').filter(x=>x.attrs.Type.endsWith('/hyperlink'));assert.equal(targets.length,3);
 assert.equal(descendants(drawing,'hlinkClick').length,3);assert.ok(textOf(files[RELS]).includes('&amp;b=&quot;'));
 assert.equal(new Set(descendants(rels,'Relationship').map(x=>x.attrs.Id)).size,descendants(rels,'Relationship').length);
 const out=readXlsx(bytes).data.sheets[0];assert.deepEqual(out.shapes.map(s=>s.hyperlink),[link,external]);assert.deepEqual(out.images[0].hyperlink,link);assert.ok(out.images[0].linked);assert.deepEqual(wb.sheets[0].shapes,before);
 const wb2=new Workbook(readXlsx(bytes).data);assert.deepEqual(readXlsx(writeXlsx(wb2)).data.sheets[0].shapes.map(s=>s.hyperlink),[link,external]);
});

function groupBook(){const wb=new Workbook();wb.sheets[0].shapes=[{id:'group',name:'group',kind:'group',x:50,y:70,w:300,h:200,hyperlink:link,groupSize:{w:300,h:200},groupItems:[shape('child'),{id:'image',name:'image',kind:'picture',src:PNG,x:150,y:20,w:50,h:50,hyperlink:{target:'#Sheet1!B7'}},{id:'nested',name:'nested',kind:'group',x:40,y:100,w:100,h:80,hyperlink:{target:'#Sheet1!C8'},groupSize:{w:100,h:80},groupItems:[shape('nested-child','ellipse',{target:'#Sheet1!D9'})]}]}];return wb;}

test('nested group parent and child links use native DrawingML as well as editable metadata',()=>{
 const wb=groupBook(),bytes=writeXlsx(wb),files=unzip(bytes),out=readXlsx(bytes).data.sheets[0].shapes[0];
 assert.deepEqual(out.hyperlink,link);assert.deepEqual(out.groupItems,wb.sheets[0].shapes[0].groupItems);
 assert.equal(descendants(parseXml(textOf(files[DRAW])),'hlinkClick').length,5);
 xmlSet(files,DRAW,textOf(files[DRAW]).replace(/<a:extLst>.*?<\/a:extLst>/gs,''));const native=save(files).data.sheets[0];
 assert.equal(native.shapes.find(s=>s.name==='child').hyperlink.target,link.target);assert.equal(native.shapes.find(s=>s.name==='nested-child').hyperlink.target,'#Sheet1!D9');assert.equal(native.images[0].hyperlink.target,'#Sheet1!B7');
});

test('external hyperlink target edits override stale group metadata even when XML r:id is unchanged',()=>{
 const files=unzip(writeXlsx(groupBook())),drawing=parseXml(textOf(files[DRAW]));
 const nv=descendants(drawing,'cNvPr').find(n=>n.attrs.name==='child'),id=child(nv,'hlinkClick').attrs['r:id'];
 const rels=parseXml(textOf(files[RELS])),rel=descendants(rels,'Relationship').find(r=>r.attrs.Id===id);
 xmlSet(files,RELS,textOf(files[RELS]).replace('Id="'+id+'" Type="'+rel.attrs.Type+'" Target="'+esc(rel.attrs.Target)+'"','Id="'+id+'" Type="'+rel.attrs.Type+'" Target="#Sheet1!Z99"'));
 const out=save(files);assert.equal(out.data.sheets[0].shapes[0].groupItems.find(s=>s.name==='child').hyperlink.target,'#Sheet1!Z99');assert.match(out.warnings.join(' '),/외부에서 수정/);
});

test('renumbered relationship IDs keep native group signatures valid',()=>{
 const files=unzip(writeXlsx(groupBook()));for(const path of [DRAW,RELS])xmlSet(files,path,textOf(files[path]).replace(/rId(\d+)/g,(_,n)=>'rId'+(Number(n)+100)));
 const out=save(files);assert.ok(!out.warnings.some(w=>/외부에서 수정/.test(w)));assert.deepEqual(out.data.sheets[0].shapes[0].groupItems,groupBook().sheets[0].shapes[0].groupItems);
});

test('unsafe URLs/actions remain inert model data; drawing import does not fetch resources',()=>{
 const wb=new Workbook();wb.sheets[0].shapes=[shape('unsafe','rect',{target:'javascript:alert(1)',targetMode:'External',action:'macro://never-run'})];
 let called=0;const fetch=globalThis.fetch;globalThis.fetch=()=>{called++;throw Error('unexpected fetch')};
 try{const out=readXlsx(writeXlsx(wb)).data.sheets[0].shapes[0];assert.deepEqual(out.hyperlink,wb.sheets[0].shapes[0].hyperlink);assert.equal(called,0);assert.equal(out.macro,undefined)}finally{globalThis.fetch=fetch}
});

test('native parent links inherit only when a flattened group child has no own link',()=>{
 const wb=groupBook();delete wb.sheets[0].shapes[0].groupItems[0].hyperlink;
 const files=unzip(writeXlsx(wb));xmlSet(files,DRAW,textOf(files[DRAW]).replace(/<a:extLst>.*?<\/a:extLst>/gs,''));
 const out=save(files).data.sheets[0];assert.deepEqual(out.shapes.find(s=>s.name==='child').hyperlink,link);assert.equal(out.images[0].hyperlink.target,'#Sheet1!B7');
 assert.equal(out.shapes.find(s=>s.name==='nested-child').hyperlink.target,'#Sheet1!D9');
});

test('native removal of a child link cannot be restored by old editable metadata',()=>{
 const files=unzip(writeXlsx(groupBook())),drawing=parseXml(textOf(files[DRAW]));
 const nv=descendants(drawing,'cNvPr').find(n=>n.attrs.name==='child'),id=child(nv,'hlinkClick').attrs['r:id'];
 const re=new RegExp('<a:hlinkClick r:id="'+id+'"[^>]*/>');assert.match(textOf(files[DRAW]),re);xmlSet(files,DRAW,textOf(files[DRAW]).replace(re,''));
 const out=save(files);assert.equal(out.data.sheets[0].shapes[0].groupItems.find(s=>s.name==='child').hyperlink,undefined);assert.match(out.warnings.join(' '),/외부에서 수정/);
});

test('native parent link edit takes precedence without invalidating unchanged child geometry',()=>{
 const files=unzip(writeXlsx(groupBook())),drawing=parseXml(textOf(files[DRAW]));
 const nv=descendants(drawing,'cNvPr').find(n=>n.attrs.name==='group'),id=child(nv,'hlinkClick').attrs['r:id'];
 const rels=parseXml(textOf(files[RELS])),rel=descendants(rels,'Relationship').find(r=>r.attrs.Id===id);
 xmlSet(files,RELS,textOf(files[RELS]).replace('Id="'+id+'" Type="'+rel.attrs.Type+'" Target="'+esc(rel.attrs.Target)+'"','Id="'+id+'" Type="'+rel.attrs.Type+'" Target="#Sheet1!Z99"'));
 const out=save(files);assert.equal(out.data.sheets[0].shapes[0].hyperlink.target,'#Sheet1!Z99');assert.deepEqual(out.data.sheets[0].shapes[0].groupItems,groupBook().sheets[0].shapes[0].groupItems);assert.ok(!out.warnings.some(w=>/외부에서 수정/.test(w)));
});
