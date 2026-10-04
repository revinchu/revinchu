import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { createSheetHtmlBlob, htmlSheetRange, exportSlicerHtml } from '../src/html-export.js';

const book = cells => new Workbook({sheets:[{name:'보고서',cells}]});
const pages = html => [...html.matchAll(/<script type="application\/json" data-wx-page[^>]*>([\s\S]*?)<\/script>/g)].map(m=>JSON.parse(m[1]));

test('HTML defaults to actual data, independent of gigantic print areas and formatted tails',async()=>{
  const wb=book({'0,0':{raw:'이름'},'1,0':{raw:'마지막 값'},'999999,5':{raw:'',style:{fill:'#ff0000'}}});
  wb.sheets[0].page={area:{r1:0,c1:0,r2:999999,c2:100},scale:400};
  assert.deepEqual(htmlSheetRange(wb,0),{r1:0,c1:0,r2:1,c2:0});
  const blob=await createSheetHtmlBlob(wb,0,{name:'HTML'}),html=await blob.text();
  assert.equal(blob.type,'text/html;charset=utf-8');assert.ok(blob.size>0&&blob.size<15000);assert.match(html,/마지막 값/);assert.equal((html.match(/<tr /g)||[]).length,2);assert.doesNotMatch(html,/<script/);
});

test('explicit print-area mode preserves multiple requested areas without pagination limits',async()=>{
  const wb=book({'0,0':{raw:'첫째'},'3,2':{raw:'둘째'}});wb.sheets[0].page={areas:[{r1:0,c1:0,r2:0,c2:0},{r1:3,c1:2,r2:3,c2:2}]};
  const html=await(await createSheetHtmlBlob(wb,0,{respectPrintArea:true})).text();
  assert.equal((html.match(/<table /g)||[]).length,2);assert.match(html,/첫째/);assert.match(html,/둘째/);
});

test('large HTML contains every row and column in bounded inert chunks and a self-contained pager',async()=>{
  const cells={};for(let r=0;r<601;r++)for(let c=0;c<70;c++)cells[`${r},${c}`]={raw:`값-${r}-${c}`};
  const wb=book(cells),html=await(await createSheetHtmlBlob(wb,0,{staticCellLimit:100,pageRows:100})).text(),chunks=pages(html);
  assert.equal(chunks.length,14);let count=0;const seen=new Set();
  for(const p of chunks){const body=p.html.join('');assert.ok((body.match(/<tr /g)||[]).length<=100);for(const m of body.matchAll(/data-row="(\d+)" data-col="(\d+)"/g)){count++;seen.add(m[1]+','+m[2]);}}
  assert.equal(count,601*70);assert.equal(seen.size,count);assert.ok(seen.has('600,69'));assert.match(html,/nonce-wixel-html-export/);assert.match(html,/모든 데이터가 이 파일/);assert.match(html,/id="wx-jump"/);
});

test('hidden/filter rows and columns plus merged anchors retain style and displayed values',async()=>{
  const wb=book({'0,0':{raw:'병합',style:{bold:true,fill:'#ffc000',color:'#001122'}},'2,0':{raw:'3.5',style:{numFmt:'percent',decimals:1}},'1,0':{raw:'숨김'}});
  Object.assign(wb.sheets[0],{merges:[{r1:0,c1:0,r2:1,c2:2}],hiddenRows:{1:true},hiddenCols:{1:true},rowHeights:{0:31},colWidths:{0:91}});
  const html=await(await createSheetHtmlBlob(wb,0)).text();assert.match(html,/rowspan="1" colspan="2"/);assert.match(html,/font-weight:700/);assert.match(html,/background:#ffc000/);assert.match(html,/width:91px/);assert.match(html,/height:31px/);assert.match(html,/350.0%/);assert.doesNotMatch(html,/>숨김</);
});

test('merges crossing output chunks remain bounded and use the original anchor content',async()=>{
  const wb=book({'0,0':{raw:'결합'},'4,2':{raw:'끝'}});wb.sheets[0].merges=[{r1:0,c1:0,r2:4,c2:1}];
  const html=await(await createSheetHtmlBlob(wb,0,{staticCellLimit:1,pageRows:2})).text(),p=pages(html);assert.equal(p.length,3);
  assert.match(p[0].html.join(''),/rowspan="2" colspan="2"/);assert.match(p[1].html.join(''),/결합/);assert.match(p[2].html.join(''),/끝/);
});

test('cell text, titles, links and JSON cannot inject active markup',async()=>{
  const attack='</script><script>alert(1)</script><img src=x onerror=alert(1)>';
  const wb=book({'0,0':{raw:attack,link:'javascript:alert(1)',style:{font:'a; background:url(https://bad)',fill:'url(https://bad)'}}});wb.sheets[0].name=attack;
  for(const staticCellLimit of [0,50000]){const html=await(await createSheetHtmlBlob(wb,0,{name:attack,staticCellLimit})).text();assert.doesNotMatch(html,/<script>alert/);assert.doesNotMatch(html,/href="javascript/);assert.doesNotMatch(html,/background:url/);assert.match(html,/default-src 'none'/);if(staticCellLimit===0){const p=pages(html);assert.equal(p.length,1);assert.match(p[0].html.join(''),/&lt;script&gt;alert/);}}
});

test('charts, pictures and shapes use one sanitized object at a time, once each',async()=>{
  const wb=book({'0,0':{raw:'본문'}});Object.assign(wb.sheets[0],{charts:[{id:'chart',w:100,h:80}],images:[{id:'image',w:90,h:60}],shapes:[{id:'shape',w:70,h:40},{id:'hidden',hidden:true}]});
  const calls=[],html=await(await createSheetHtmlBlob(wb,0,{renderObject:async(kind,o)=>{calls.push([kind,o.id]);return '<svg><text>안전한 개체</text></svg>';}})).text();
  assert.deepEqual(calls,[['chart','chart'],['image','image'],['shape','shape']]);assert.equal((html.match(/안전한 개체/g)||[]).length,3);await assert.rejects(createSheetHtmlBlob(wb,0),/개체|차트/);
});

test('dynamic array followers extend the exported data extent and remain visible',async()=>{
  const wb=book({'0,0':{raw:'=SEQUENCE(4,2)'}});const html=await(await createSheetHtmlBlob(wb,0)).text();assert.equal((html.match(/<td /g)||[]).length,8);assert.match(html,/>8<\/td>/);
});

test('cancellation or document replacement rejects instead of returning a partial/zero-byte file',async()=>{
  const wb=book({'0,0':{raw:'값'}}),controller=new AbortController();controller.abort();await assert.rejects(createSheetHtmlBlob(wb,0,{signal:controller.signal}),/취소/);
  let checks=0;await assert.rejects(createSheetHtmlBlob(wb,0,{assertCurrent:()=>{if(++checks>1)throw Error('문서 변경');}}),/문서 변경/);
});


test('static and paged object canvases preserve coordinates, z order, noPrint and slicers',async()=>{
  const wb=book({'0,0':{raw:'표'}});Object.assign(wb.sheets[0],{charts:[{id:'chart',x:91,y:37,w:200,h:100,z:2,noPrint:true}],slicers:[{id:'slicer',x:12,y:8,w:60,h:70,z:1}]});
  for(const staticCellLimit of [0,50000]){const calls=[],html=await(await createSheetHtmlBlob(wb,0,{staticCellLimit,renderObject:(kind,o)=>{calls.push(kind);return '<svg></svg>';}})).text();assert.deepEqual(calls,['slicer','chart']);assert.match(html,/left:91px;top:37px;width:200px;height:100px/);assert.match(html,/left:12px;top:8px/);assert.ok(html.indexOf('data-object="slicer"')<html.indexOf('data-object="chart"'));if(!staticCellLimit)assert.match(html,/원래 좌표/);}
});

test('stored formula values and workbook source remain unchanged by HTML export',async()=>{
  const wb=book({'0,0':{raw:'=UNSUPPORTED()',cached:99},'1,0':{raw:'본문',style:{fill:'#ffffff'}}});wb.sheets[0].fileValues=true;
  const before=JSON.stringify(wb.serialize()),html=await(await createSheetHtmlBlob(wb,0)).text();assert.match(html,/>99<\/td>/);assert.equal(JSON.stringify(wb.serialize()),before);
});

test('all-hidden rows produce a nonempty valid empty document',async()=>{
  const wb=book({'0,0':{raw:'숨김'}});wb.sheets[0].hiddenRows={0:true};const html=await(await createSheetHtmlBlob(wb,0)).text();assert.match(html,/표시할 셀이 없습니다/);assert.equal((html.match(/<table/g)||[]).length,0);assert.equal((html.match(/<\/table>/g)||[]).length,0);
});


test('slicer snapshot includes visible item states without expanding a 50,000 item DOM',()=>{
  const items=Array.from({length:50000},(_,i)=>({text:i===0?'<img src=x onerror=alert(1)>':`항목 ${i}`,selected:i%2===0,hasData:i%3!==0}));
  const html=exportSlicerHtml({caption:'선택',columns:2,h:140,buttonHeight:24,style:'SlicerStyleDark2'},{items});
  assert.ok((html.match(/data-slicer-selected=/g)||[]).length<20);assert.match(html,/50,000개/);assert.match(html,/25,000개/);assert.match(html,/번째 항목 표시/);assert.match(html,/&lt;img/);assert.doesNotMatch(html,/<img/);assert.match(html,/data-slicer-selected="1"/);assert.match(html,/data-slicer-selected="0"/);
});

test('HTML border widths match thin, medium dashed and double UI styles, including no-grid sheets',async()=>{
  const wb=book({'0,0':{raw:'선',style:{bb:true,bbs:'double',bbc:'#123456',bt:true,bts:'mediumDashed',btc:'#abcdef'}}});wb.sheets[0].noGrid=true;
  const html=await(await createSheetHtmlBlob(wb,0)).text();assert.match(html,/border-bottom:3px double #123456/);assert.match(html,/border-top:2px dashed #abcdef/);assert.match(html,/<table class=""/);
});
