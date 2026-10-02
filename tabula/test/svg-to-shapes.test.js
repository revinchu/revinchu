import { test } from 'node:test';
import assert from 'node:assert/strict';
import { svgToEditableGroup } from '../src/svg-to-shapes.js';
import { parseSvgTransform, transformCommands, flattenCommands } from '../src/geometry-path.js';
import { ringsContain } from '../src/shape-boolean.js';
import { shapeSvg } from '../src/shapes.js';
import { ungroupObjects } from '../src/object-group.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx,writeXlsx } from '../src/xlsx.js';
import { unzip,zip,textOf } from '../src/zip.js';
import { readFileSync } from 'node:fs';
import { gunzipSync } from 'node:zlib';
const svg=body=>`<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">${body}</svg>`;
const at=(sh,x,y)=>ringsContain(sh.path.paths.flatMap(p=>flattenCommands(p.commands).map(r=>r.map(([a,b])=>[sh.x+a*sh.w,sh.y+b*sh.h]))),[x,y]);

test('SVG의 독립 색 요소는 해제·개별 서식 편집 가능한 그룹이 된다',()=>{
  const pic={id:'icon1',x:50,y:70,w:200,h:200,rot:30,flip:true,alt:'합성 아이콘',hyperlink:{target:'#Sheet1!A1'}},before=structuredClone(pic);
  const g=svgToEditableGroup(svg('<rect x="5" y="5" width="35" height="90" fill="#ff0000"/><circle cx="70" cy="50" r="20" fill="#0000ff"/>'),pic);
  assert.deepEqual(pic,before);assert.equal(g.kind,'group');assert.equal(g.groupItems.length,2);assert.deepEqual(g.groupItems.map(x=>x.fill),['#ff0000','#0000ff']);
  assert.equal(g.rot,30);assert.equal(g.flip,true);assert.equal(g.alt,pic.alt);assert.deepEqual(g.hyperlink,pic.hyperlink);
  const parts=ungroupObjects(g);assert.equal(parts.length,2);parts[0].fill='#00ff00';assert.equal(parts[1].fill,'#0000ff');assert.equal(g.groupItems[0].fill,'#ff0000');
  assert.match(shapeSvg(g),/fill="#ff0000"/);assert.match(shapeSvg(g),/fill="#0000ff"/);
});
test('단일 path 여러 구성요소와 구멍은 nonzero·evenodd 규칙대로 남는다',()=>{
  const d='M0 0H100V100H0Z M20 20H80V80H20Z';
  const even=svgToEditableGroup(svg(`<path d="${d}" fill-rule="evenodd" fill="red"/>`)).groupItems;
  assert.equal(even.some(s=>at(s,50,50)),false);assert.equal(even.some(s=>at(s,10,50)),true);
  const nonzero=svgToEditableGroup(svg(`<path d="${d}" fill="red"/>`)).groupItems;
  assert.equal(nonzero.some(s=>at(s,50,50)),true);
  const reverse=svgToEditableGroup(svg('<path d="M0 0H100V100H0Z M20 20V80H80V20Z"/>')).groupItems;
  assert.equal(reverse.some(s=>at(s,50,50)),false);
  const separated=svgToEditableGroup(svg('<path d="M0 0H40V40H0Z M60 0H100V40H60Z M10 10V30H30V10Z"/>'));
  assert.equal(separated.groupItems.length,2);assert.equal(separated.groupItems.some(s=>at(s,20,20)),false);assert.equal(separated.groupItems.some(s=>at(s,80,20)),true);
});
test('변환·상속 색·원호·상대좌표·기본 도형을 실제 경로로 변환한다',()=>{
  const g=svgToEditableGroup(svg('<g transform="translate(10 20) scale(2)" fill="#123" color="blue"><path d="m0 0h10v10h-10z"/><path d="M15 0q5 10 10 0t10 0" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round"/></g><ellipse cx="50" cy="75" rx="20" ry="10" fill="green"/><polygon points="5,80 10,95 1,95"/><line x1="80" y1="80" x2="95" y2="95" stroke="red"/>'));
  assert.equal(g.groupItems.length,5);assert.deepEqual([g.groupItems[0].x,g.groupItems[0].y,g.groupItems[0].w,g.groupItems[0].h],[10,20,20,20]);assert.equal(g.groupItems[0].fill,'#112233');
  assert.equal(g.groupItems[1].stroke,'#0000ff');assert.equal(g.groupItems[1].strokeWidth,4);assert.equal(g.groupItems[1].lineCap,'rnd');assert.ok(g.groupItems[2].path.paths[0].commands.some(c=>c[0]==='C'));
  assert.deepEqual(transformCommands([['M',1,2]],parseSvgTransform('translate(10 20) scale(2)')),[['M',12,24]]);
});
test('viewBox 원점·기본 meet·none 배치 및 부모 회전을 보존한다',()=>{
  const text='<svg viewBox="10 20 20 40"><rect x="10" y="20" width="20" height="40"/></svg>';
  const g=svgToEditableGroup(text,{w:100,h:100});assert.deepEqual([g.groupItems[0].x,g.groupItems[0].y,g.groupItems[0].w,g.groupItems[0].h],[25,0,50,100]);
  const stretch=svgToEditableGroup(text.replace('<svg','<svg preserveAspectRatio="none"'),{w:100,h:100});assert.deepEqual([stretch.groupItems[0].x,stretch.groupItems[0].y,stretch.groupItems[0].w,stretch.groupItems[0].h],[0,0,100,100]);
});
test('지원하지 않는 효과·외부요소·CSS·잘린 경로는 일부변환 없이 거절한다',()=>{
  for(const body of ['<image href="https://example.invalid/a.png"/>','<use href="#x"/>','<text x="2" y="20">글</text>','<script>bad()</script>','<rect width="30" height="30" fill="url(#paint)"/>','<g opacity=".5"><rect width="10" height="10"/></g>','<rect width="10" height="10" style="filter: blur(2px)"/>','<g transform="scale(2 1)"><path d="M0 0L20 20" stroke="red"/></g>','<rect x="95" width="10" height="10"/>'])assert.throws(()=>svgToEditableGroup(svg(body)),/변환할 수 없습니다/);
  for(const text of ['<svg viewBox="0 0 10 10"><rect width="5" height="5"></svg>','<svg viewBox="0 0 10 10"><rect width="5" height="5" nope=bad/></svg>','<svg viewBox="0 0 10 10"><path d="M0 0L5 5" d="M1 1"/></svg>'])assert.throws(()=>svgToEditableGroup(text),/변환할 수 없습니다/);
  assert.throws(()=>svgToEditableGroup(svg('<rect width="10" height="10"/>'),{crop:{l:.1}}),/자르기/);
});
test('currentColor 자기 참조와 숨긴 요소·소수 픽셀 도형도 안전하게 처리한다',()=>{
  const g=svgToEditableGroup(svg('<g color="currentColor"><rect width=".5" height=".5" fill="currentColor"/><rect display="none" width="10" height="10"/></g>'));
  assert.equal(g.groupItems.length,1);assert.equal(g.groupItems[0].fill,'#000000');assert.equal(g.groupItems[0].w,1);assert.equal(at(g.groupItems[0],.25,.25),true);assert.equal(at(g.groupItems[0],.75,.75),false);
  assert.equal(svgToEditableGroup(svg('<g visibility="hidden"><rect visibility="visible" width="10" height="10"/></g>')).groupItems.length,1);
  assert.throws(()=>svgToEditableGroup(svg('<rect width="100" height="100" fill="none" stroke="red"/>')),/잘리는 경로/);
});
test('채우기·선 알파와 요소 투명도는 실제 SVG 및 표준 XLSX에 곱해서 보존한다',()=>{
  const g=svgToEditableGroup(svg('<rect x="10" y="10" width="30" height="80" fill="rgba(255,0,0,.5)" fill-opacity=".5" opacity=".5"/><rect x="60" y="10" width="30" height="80" fill="none" opacity=".5" stroke="rgba(0,0,255,.5)" stroke-opacity=".8" stroke-width="2"/>'));
  assert.equal(g.groupItems[0].fillOpacity,.125);assert.equal(g.groupItems[1].strokeOpacity,.2);assert.match(shapeSvg(g),/fill-opacity="0.125"/);assert.match(shapeSvg(g),/stroke-opacity="0.2"/);
  const wb=new Workbook();wb.setSheetProp(0,'shapes',[g]);const files=unzip(writeXlsx(wb));for(const k of Object.keys(files))if(k.endsWith('.xml'))files[k]=textOf(files[k]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'');
  const back=readXlsx(zip(files)).data.sheets[0].shapes.flatMap(s=>s.kind==='group'?ungroupObjects(s):[s]);assert.equal(back[0].fillOpacity,.125);assert.equal(back[1].strokeOpacity,.2);
  assert.throws(()=>svgToEditableGroup(svg('<rect x="10" y="10" width="80" height="80" fill="red" stroke="blue" opacity=".5"/>')),/함께 합성/);
});
test('SVG 변환 그룹은 표준 grpSp+custGeom으로 메타데이터 제거 후에도 구멍·색·곡선을 저장한다',()=>{
  const g=svgToEditableGroup(svg('<path fill-rule="evenodd" fill="red" d="M0 0H100V100H0Z M20 20H80V80H20Z"/><circle cx="50" cy="50" r="10" fill="blue"/>'),{x:10,y:20,w:100,h:100});
  const wb=new Workbook();wb.setSheetProp(0,'shapes',[g]);const files=unzip(writeXlsx(wb));
  const xml=textOf(files['xl/drawings/drawing1.xml']);assert.match(xml,/<xdr:grpSp>/);assert.match(xml,/<a:custGeom>/);
  for(const k of Object.keys(files))if(k.endsWith('.xml'))files[k]=textOf(files[k]).replace(/<a:extLst>[\s\S]*?<\/a:extLst>/g,'');
  const back=readXlsx(zip(files)).data.sheets[0].shapes;
  // Metadata-free importer flattens standard group children into editable shapes.
  const parts=back.flatMap(s=>s.kind==='group'?ungroupObjects(s):[s]);assert.equal(parts.length,2);assert.deepEqual(parts.map(s=>s.fill),['#ff0000','#0000ff']);
  assert.equal(parts[0].path.paths[0].commands.filter(c=>c[0]==='M').length,2);assert.ok(parts[1].path.paths[0].commands.some(c=>c[0]==='C'));
});
test('실제 내장 아이콘 모음의 각 범주 대표와 빈 Office CSS 선언을 변환한다',()=>{
  const library=JSON.parse(gunzipSync(readFileSync(new URL('../assets/iconlib.json.gz',import.meta.url))));
  let tested=0,parts=0,emptyStyle=false;
  for(const [,icons] of library){const samples=[icons[0],icons.find(([,body])=>body.includes('fill-rule')),icons.find(([,body])=>body.includes('<style'))].filter(Boolean);
    for(const [,body,vb='0 0 96 96'] of samples){
      const text=`<svg viewBox="${vb}" fill="#000000">${body.replace(/\sfill="(?!none)[^"]*"/g,'').replace(/fill:\s*#[0-9a-fA-F]{3,8};?/g,'')}</svg>`;
      const group=svgToEditableGroup(text,{w:96,h:96});tested++;parts+=group.groupItems.length;if(body.includes('<style'))emptyStyle=true;
    }
  }assert.ok(tested>=20);assert.ok(parts>tested);assert.equal(emptyStyle,true);
});
