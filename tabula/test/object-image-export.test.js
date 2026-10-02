import { test } from 'node:test';
import assert from 'node:assert/strict';
import { objectImageBounds, objectImageSvg, objectImageBlob, objectRasterSize, safeObjectSvg, prepareImageSvg, svgImageBlob } from '../src/object-image-export.js';
import { parseXml, descendants } from '../src/xml.js';
import { newSmartArt } from '../src/smartart.js';

const png='data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+aGRkAAAAASUVORK5CYII=';
const rect={id:'test',kind:'rect',x:500,y:700,w:100,h:50,fill:'#ff0000',stroke:'#000000',text:''};
const svg=body=>`<svg xmlns="http://www.w3.org/2000/svg" width="100" height="50" viewBox="0 0 100 50">${body}</svg>`;
const close=(a,b)=>assert.ok(Math.abs(a-b)<1e-8,`${a} ≈ ${b}`);

test('object image rotates local bounds independently of sheet position and retains vector geometry',async()=>{
  const b=objectImageBounds({...rect,rot:90});close(b.x,25);close(b.y,-25);close(b.w,50);close(b.h,100);
  assert.deepEqual(objectRasterSize(objectImageBounds({...rect,rot:180}).w,objectImageBounds({...rect,rot:180}).h,1),{width:100,height:50});
  const result=await objectImageSvg({...rect,rot:90,flip:true},{title:'도형 <제목>'});
  close(result.width,50);close(result.height,100);assert.match(result.svg,/rotate\(90,50,25\)/);assert.match(result.svg,/scale\(-1,1\)/);
  assert.match(result.svg,/<path /);assert.doesNotMatch(result.svg,/<image /);assert.match(result.svg,/도형 &lt;제목&gt;/);assert.doesNotMatch(result.svg,/translate\(500,700\)/);
});

test('picture crop, correction, border, flip and reflected overflow share the display renderer',async()=>{
  const pic={w:100,h:50,src:png,rot:90,flip:true,flipV:true,crop:{l:.25,r:.25,t:0,b:0},reflection:{size:.5,gap:5,opacity:.5},correction:{brightness:.2},border:'#123456',borderDash:'dash'};
  const before=structuredClone(pic),out=await objectImageSvg(pic,{kind:'picture'});
  assert.deepEqual(pic,before);close(out.width,80);close(out.height,100);
  assert.match(out.svg,/x="-50"/);assert.match(out.svg,/width="200"/);assert.match(out.svg,/feComponentTransfer/);assert.match(out.svg,/stroke-dasharray=/);assert.match(out.svg,/scale\(-1,-1\)/);assert.match(out.svg,/<mask /);
});

test('custom paths outside the object box, line arrows and combined effects reserve export bleed',()=>{
  const p={...rect,w:100,h:100,path:{paths:[{fill:false,stroke:true,commands:[['M',-1,0],['L',2,1]]}]},strokeWidth:4,arrow:'both',shadow:{dx:10,dy:-8,blur:5},glow:{size:3},soft:2};
  const b=objectImageBounds(p);assert.ok(b.x<-130);assert.ok(b.x+b.w>230);assert.ok(b.y<0&&b.y+b.h>100);
});

test('group bounds include transformed children, nested resize and flips; hidden resources are not loaded',async()=>{
  const group={id:'parent',kind:'group',x:1000,y:1000,w:200,h:100,groupSize:{w:100,h:50},flip:true,groupItems:[{...rect,id:'same',x:80,y:0,w:40,h:20,text:'그룹 글자'}, {kind:'picture',id:'same',x:0,y:0,w:10,h:10,src:'https://example.invalid/no',hidden:true}]};
  const before=structuredClone(group),out=await objectImageSvg(group);
  close(out.bounds.x,-40);close(out.width,240);assert.deepEqual(group,before);assert.equal(descendants(parseXml(out.svg),'tspan').map(s=>s.text).join(''),'그룹 글자');assert.doesNotMatch(out.svg,/example.invalid/);assert.match(out.svg,/viewBox="0 0 100 50"/);
});

test('rich text, Korean and mixed tspan ordering survive a standalone XML save',async()=>{
  const out=await objectImageSvg({...rect,w:300,paras:[{align:'left',runs:[{t:'한글 <&> ',b:true},{t:'다음',color:'#123456',font:'맑은 고딕'}]}]});
  const root=parseXml(out.svg),spans=descendants(root,'tspan');assert.equal(spans.map(s=>s.text).join(''),'한글 <&> 다음');assert.equal(spans[0].attrs['font-weight'],'bold');
  assert.equal(safeObjectSvg(svg('<text>A<tspan>B</tspan>C</text>')),svg('<text>A<tspan>B</tspan>C</text>'));
});

test('SmartArt vector parts and labels remain vectors',async()=>{
  const o=newSmartArt('process',{w:600,h:200});const out=await objectImageSvg(o);assert.match(out.svg,/<path /);assert.match(out.svg,/항목 1/);assert.doesNotMatch(out.svg,/<image /);
});

test('SVG image data expands into editable paths with safe filters, never stale PNG caches',async()=>{
  const inner=svg('<defs><linearGradient id="paint"><stop offset="0" stop-color="#ff0000"/></linearGradient></defs><path d="M0 0H100V50H0Z" fill="url(#paint)"/>');
  const pic={w:100,h:50,src:'data:image/svg+xml;charset=utf-8,'+encodeURIComponent(inner),png:'stale',effectPng:'stale',icon:{body:'<script/>',vb:'unsafe'}};
  const out=await objectImageSvg(pic,{kind:'picture'});assert.doesNotMatch(out.svg,/<image |data:image|stale|script/);assert.match(out.svg,/linearGradient/);assert.match(out.svg,/<path /);
  const twice=await objectImageSvg({kind:'group',w:200,h:50,groupItems:[{...pic,kind:'picture',id:'first',x:0,y:0},{...pic,kind:'picture',id:'second',x:100,y:0}]});
  const ids=[...twice.svg.matchAll(/\sid="([^"]*)"/g)].map(m=>m[1]);assert.equal(new Set(ids).size,ids.length);
  for(const ref of twice.svg.matchAll(/url\(#([^)]*)\)/g))assert.ok(ids.includes(ref[1]),ref[1]);
  const clip=await prepareImageSvg(svg(`<defs><clipPath id="outerclip"><rect width="10" height="10"/></clipPath></defs><image width="100" height="50" href="${pic.src}" clip-path="url(#outerclip)"/>`));assert.match(clip,/clip-path="url\(#outerclip\)"/);
});

test('standalone SVG forbids active content, external resources, CSS escape and malformed XML',()=>{
  for(const body of ['<script>alert(1)</script>','<foreignObject/>','<path onload="a()"/>','<image href="https://example.com/x.png"/>','<use href="other.svg#x"/>','<path style="fill:url(https://example.com/x)"/>','<path style="fill:u\\72l(#x)"/>','<path fill="url(&quot;https://example.com/x&quot;)"/>','<g><path/></svg>','<path/><svg xmlns="bad"/>','<path xml:base="https://example.com"/>'])assert.throws(()=>safeObjectSvg(svg(body)),undefined,body);
  assert.throws(()=>safeObjectSvg('<!DOCTYPE svg SYSTEM "https://example.com"><svg/>'));
  assert.throws(()=>safeObjectSvg('<svg/><svg/>'));
  assert.throws(()=>safeObjectSvg(svg('<image href="data:image/svg+xml,'+encodeURIComponent(svg('<script/>'))+'"/>')));
  assert.throws(()=>safeObjectSvg('<svg width="1" width="2"/>'));
});

test('explicit resolver is deduplicated and never changes original picture sources',async()=>{
  const source='https://example.invalid/p.png';let calls=0;
  const g={kind:'group',w:100,h:100,groupItems:[{kind:'picture',id:'p1',w:50,h:50,x:0,y:0,src:source},{kind:'picture',id:'p2',w:50,h:50,x:50,y:50,src:source}]};
  await assert.rejects(()=>objectImageSvg(g),/먼저 파일/);
  const out=await objectImageSvg(g,{resolveImage:async u=>{assert.equal(u,source);calls++;return png;}});
  assert.equal(calls,1);assert.doesNotMatch(out.svg,/https:\/\/example/);assert.equal(g.groupItems[0].src,source);
  let rawCalls=0;const prepared=await prepareImageSvg(svg(`<image href="${source}"/><image href="${source}"/>`),{resolveImage:async()=>{rawCalls++;return png;}});assert.equal(rawCalls,1);assert.match(prepared,/data:image\/png/);
});

test('resource resolver may not return active or linked data; empty and live-linked images fail clearly',async()=>{
  await assert.rejects(()=>objectImageSvg({w:10,h:10,src:'https://example.invalid/x'},{kind:'picture',resolveImage:async()=> 'javascript:alert(1)'}));
  await assert.rejects(()=>objectImageSvg({w:10,h:10,src:png,linked:{sheet:'A'}},{kind:'picture'}),/현재 모습/);
  await assert.rejects(()=>objectImageSvg({w:10,h:10},{kind:'picture'}),/그림 데이터/);
  await assert.rejects(()=>objectImageSvg({w:10,h:10,src:png,media:{type:'video'}},{kind:'picture'}),/동영상/);
  let calls=0;await assert.rejects(()=>prepareImageSvg(svg('<image href="https://example.invalid/x"/><script/>'),{resolveImage:async()=>{calls++;return png;}}));assert.equal(calls,0);
  await prepareImageSvg(svg('<!-- <image href="https://example.invalid/x"/> --><path d="M0 0"/>'),{resolveImage:async()=>{calls++;return png;}});assert.equal(calls,0);
});

test('raster dimensions retain requested scale and refuse oversized output without silent shrink',()=>{
  assert.deepEqual(objectRasterSize(100.2,50.2,3),{width:301,height:151});
  assert.deepEqual(objectRasterSize(2048,100,2),{width:4096,height:200});
  for(const a of [[2049,100,2],[4096,4096,1],[1,1,4],[0,1,1],[NaN,1,1]])assert.throws(()=>objectRasterSize(...a));
  assert.throws(()=>objectImageBounds({...rect,w:Infinity}));
});

test('SVG file supports large vector dimensions and white background independently of raster limits',async()=>{
  const text='<svg xmlns="http://www.w3.org/2000/svg" width="10000" height="100" viewBox="-5 -8 10000 100"><path d="M0 0L5 5"/></svg>';
  const blob=await svgImageBlob(text,{width:10000,height:100,format:'svg',background:'white'});
  assert.match(blob.type,/image\/svg\+xml/);assert.match(await blob.text(),/<rect x="-5" y="-8" width="10000" height="100" fill="#ffffff"\/>/);
  const source=await objectImageBlob({...rect,text:'벡터 저장'},{format:'svg'});assert.match(await source.text(),/벡터 저장/);
  assert.match(await (await svgImageBlob('<svg/>',{width:2,height:2,format:'svg',background:'white'})).text(),/<rect.*<\/svg>$/);
  await assert.rejects(()=>svgImageBlob(text,{width:65537,height:100,format:'svg'}),/65,536/);
});
