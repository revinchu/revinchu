import test from 'node:test';
import assert from 'node:assert/strict';
import { Axis } from '../src/axis.js';
import { headerResizeEdge } from '../src/header-resize.js';
import { GridView } from '../src/view.js';

const columns = () => new Axis(80, {}, [], 100);
const rows = () => new Axis(20, {}, [], 100);

test('column boundary has the same five CSS pixel hit region at every zoom', () => {
  for (const zoom of [.25, .55, .7, 1, 2, 4]) {
    const options = { zoom, header: 34, viewport: 700 };
    for (const delta of [-5, -4.9, 0, 4.9, 5]) assert.equal(headerResizeEdge(columns(), 114 + delta / zoom, options), 0, `zoom=${zoom} delta=${delta}`);
    for (const delta of [-5.2, 5.2]) assert.equal(headerResizeEdge(columns(), 114 + delta / zoom, options), null);
  }
});

test('row edge uses the nearest boundary and a symmetric CSS pixel tolerance', () => {
  assert.equal(headerResizeEdge(rows(), 37, { header:20 }), 0);
  assert.equal(headerResizeEdge(rows(), 43, { header:20 }), 0);
  assert.equal(headerResizeEdge(rows(), 47, { header:20 }), null);
  assert.equal(headerResizeEdge(rows(), 57, { header:20 }), 1);
  assert.equal(headerResizeEdge(rows(), 35, { header:20, zoom:.25 }), 0);
  assert.equal(headerResizeEdge(rows(), 49, { header:20, zoom:.25 }), 0);
  assert.equal(headerResizeEdge(rows(), 51, { header:20, zoom:.25 }), 1);
});

test('hidden and zero-sized items never become resize targets', () => {
  const axis = new Axis(80, { 1:0 }, [{ 2:true, 4:true }], 8);
  assert.equal(headerResizeEdge(axis, 114, { header:34 }), 0);
  assert.equal(headerResizeEdge(axis, 116, { header:34 }), 0);
  assert.equal(headerResizeEdge(axis, 194, { header:34 }), 3);
  assert.equal(headerResizeEdge(axis, 196, { header:34 }), 3);
  assert.equal(headerResizeEdge(new Axis(0, {}, [], 10), 40, { header:34 }), null);
});

test('a clipped scroll edge never targets the offscreen previous column', () => {
  const options = { header:34, viewport:400, scroll:81 };
  assert.equal(headerResizeEdge(columns(), 35, options), null);
  assert.equal(headerResizeEdge(columns(), 113, options), 1);
  assert.equal(headerResizeEdge(columns(), 116, options), 1);
  assert.equal(headerResizeEdge(columns(), 397, { ...options, viewport:398, scroll:35 }), null);
});

test('the frozen seam resizes the frozen column, not a hidden scroll predecessor', () => {
  const options = { header:34, viewport:600, frozenStart:0, frozenEnd:2, origin:0, scroll:200 };
  for (const point of [190, 194, 198]) assert.equal(headerResizeEdge(columns(), point, options), 1);
  assert.equal(headerResizeEdge(columns(), 234, options), 4);
  assert.equal(headerResizeEdge(columns(), 114, options), 0);
});

test('saved frozen origins and partially scrolled rows retain their actual row index', () => {
  const options = { header:20, viewport:300, frozenStart:5, frozenEnd:7, origin:100, scroll:43 };
  assert.equal(headerResizeEdge(rows(), 40, options), 5);
  assert.equal(headerResizeEdge(rows(), 60, options), 6);
  assert.equal(headerResizeEdge(rows(), 63, options), 6);
  assert.equal(headerResizeEdge(rows(), 77, options), 9);
});

test('grid borders, outside coordinates and a fully clipped pane are not handles', () => {
  const options = { header:34, viewport:300 };
  for (const point of [-5, 0, 33, 34, 35, 299, 300, 305, NaN]) assert.equal(headerResizeEdge(columns(), point, options), null);
  assert.equal(headerResizeEdge(columns(), 35, { header:34, viewport:34 }), null);
  assert.equal(headerResizeEdge(columns(), 114, { header:34, zoom:0 }), 0);
});

test('variable widths select the closest of two adjacent resize boundaries', () => {
  const axis = new Axis(80, { 0:4, 1:6 }, [], 10);
  assert.equal(headerResizeEdge(axis, 39, { header:34 }), 0);
  assert.equal(headerResizeEdge(axis, 43, { header:34 }), 1);
  assert.equal(headerResizeEdge(axis, 44, { header:34 }), 1);
});

function linearEdge(axis, point, options) {
  const { zoom, header, viewport, frozenStart, frozenEnd, origin, scroll } = options;
  if (point < header || point >= viewport) return null;
  const seam = Math.min(viewport, header + Math.max(0, axis.pos(frozenEnd) - origin));
  const visible = [];
  for (let item = frozenStart; item < axis.max; item++) {
    if (!axis.size(item)) continue;
    const frozen = item < frozenEnd;
    const start = frozen ? header : seam, end = frozen ? seam : viewport;
    const edge = header + axis.pos(item + 1) - origin - (frozen ? 0 : scroll);
    if (end > start && edge > start + 1e-7 && edge <= end + 1e-7) visible.push({ item, delta:Math.abs(edge-point) });
  }
  visible.sort((a,b)=>a.delta-b.delta);
  return visible[0]?.delta <= 5 / zoom + 1e-7 ? visible[0].item : null;
}

test('randomized visible boundaries agree with a linear geometric reference', () => {
  let seed = 91377;
  const random = () => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);
  for (let round=0; round<40; round++) {
    const sizes = {}, hidden = {};
    for (let i=0;i<60;i++) { if(random()<.25) sizes[i]=Math.round(random()*60); if(random()<.15) hidden[i]=true; }
    const axis = new Axis(23, sizes, [hidden], 60), frozenStart=Math.floor(random()*4), frozenEnd=frozenStart+Math.floor(random()*4);
    const options = { zoom:[.25,.55,1,4][round%4], header:34, viewport:400, frozenStart, frozenEnd, origin:axis.pos(frozenStart), scroll:Math.round(random()*350) };
    for (let point=34.31;point<400;point+=2.73) assert.equal(headerResizeEdge(axis,point,options), linearEdge(axis,point,options), `round=${round} point=${point}`);
  }
});

test('long filtered row runs use bounded lookups without enumerating every row', () => {
  const flags=new Uint8Array(1_000_000); flags.fill(1); flags[0]=flags[999999]=0;
  const axis=new Axis(20,{},[{start:0,__bits:flags,count:999998}],flags.length);
  let reads=0; const size=axis.size.bind(axis); axis.size=i=>{reads++;return size(i);};
  assert.equal(headerResizeEdge(axis,41,{header:20}),0);
  assert.equal(headerResizeEdge(axis,60,{header:20}),999999);
  assert.equal(headerResizeEdge(axis,30,{header:20,zoom:.5,tolerance:3,maxFraction:.2}),null);
  assert.equal(headerResizeEdge(axis,60,{header:20,zoom:.5,tolerance:3,maxFraction:.2}),999999);
  assert.ok(reads<40, `size reads: ${reads}`);
});

function grid(overrides={}) {
  return { viewEl:{getBoundingClientRect:()=>({left:100,top:200})}, z:.5,
    hw:34,hh:20,viewW:900,viewH:500,cols:columns(),rows:rows(),
    originX:0,originY:0,sx:0,sy:0,fc:0,fr:0,frozenLeft:0,frozenTop:0,frozenW:0,frozenH:0, ...overrides };
}

test('GridView keeps the column tolerance while row headers retain a selection center', () => {
  const view=grid();
  const col=GridView.prototype.hitTest.call(view,100+114*.5+4,205);
  assert.equal(col.zone,'colHeader'); assert.equal(col.edgeCol,0);
  const row=GridView.prototype.hitTest.call(view,105,200+40*.5+4);
  assert.equal(row.zone,'rowHeader'); assert.equal(row.r,1); assert.equal(row.edgeRow,null);
  assert.equal(GridView.prototype.hitTest.call(view,105,200+40*.5+1).edgeRow,0);
});

test('outline bands and cell bodies are excluded from resize hit testing', () => {
  const view=grid({olw:14,olh:14,hw:48,hh:34});
  const outline=GridView.prototype.hitTest.call(view,100+128*.5,202);
  assert.equal(outline.zone,'outline'); assert.equal(outline.edgeCol,null);
  const body=GridView.prototype.hitTest.call(view,100+128*.5,230);
  assert.equal(body.zone,'cell'); assert.equal(body.edgeCol,null); assert.equal(body.edgeRow,null);
  const corner=GridView.prototype.hitTest.call(view,105,205);
  assert.equal(corner.zone,'corner'); assert.equal(corner.edgeCol,null);
});

test('18~22px 행의 가운데는 낮은 배율에서도 선택하고 정확 경계는 크기를 조절한다', () => {
  for (const height of [18, 20, 22]) for (const zoom of [.25, .55, .7, 1, 1.5, 2, 4]) {
    const view=grid({z:zoom, rows:new Axis(height,{},[],100)}), r=3;
    const hit=y=>GridView.prototype.hitTest.call(view,105,200+y*zoom);
    for(const fraction of [.25,.5,.75]) {
      const row=hit(20+(r+fraction)*height);
      assert.equal(row.zone,'rowHeader'); assert.equal(row.r,r);
      assert.equal(row.edgeRow,null,'height='+height+' zoom='+zoom+' fraction='+fraction);
    }
    const edge=20+(r+1)*height, offset=Math.min(1/zoom,height*.1);
    for(const delta of [-offset,0,offset]) assert.equal(hit(edge+delta).edgeRow,r,'정확한 행 경계 유지');
  }
});

test('높은 행도 3 CSS px 밖에서는 선택하고 이전 5px 영역을 크기 조절로 잡지 않는다', () => {
  for(const zoom of [.55,1,2]) {
    const view=grid({z:zoom,rows:new Axis(40,{},[],100)}), edge=60;
    for(const delta of [-4,4]) assert.equal(GridView.prototype.hitTest.call(view,105,200+edge*zoom+delta).edgeRow,null);
    for(const delta of [-2.9,0,2.9]) assert.equal(GridView.prototype.hitTest.call(view,105,200+edge*zoom+delta).edgeRow,0);
  }
});

test('선택 중심 상한은 이전 행 대신 포인터가 가리키는 실제 행 높이를 쓴다', () => {
  const axis=new Axis(40,{1:8},[],100), options={header:20,zoom:.5,maxFraction:.2,tolerance:3};
  assert.equal(headerResizeEdge(axis,59,options),0);
  assert.equal(headerResizeEdge(axis,61,options),0);
  assert.equal(headerResizeEdge(axis,64,options),null,'짧은 다음 행의 중앙은 선택 영역');
  assert.equal(headerResizeEdge(axis,67,options),1);
  assert.equal(headerResizeEdge(axis,68,options),1);
  assert.equal(headerResizeEdge(axis,71,options),1,'긴 다음 행 쪽의 가까운 경계는 유지');
});

test('숨김과 틀 고정·스크롤 뒤에도 짧은 행의 선택 중심과 경계 인덱스는 유지한다', () => {
  const axis=new Axis(20,{6:8},[{7:true,8:true}],100);
  const options={header:20,viewport:300,zoom:.5,frozenStart:0,frozenEnd:2,origin:0,scroll:80,tolerance:3,maxFraction:.2};
  assert.equal(headerResizeEdge(axis,60,options),1,'틀 고정 끝 경계');
  assert.equal(headerResizeEdge(axis,64,options),null,'스크롤 영역의 짧은 행 중앙');
  assert.equal(headerResizeEdge(axis,68,options),6,'숨김 행 앞의 실제 경계');
  assert.equal(headerResizeEdge(axis,78,options),null,'숨김 행 뒤 보이는 행 중앙');
  assert.equal(headerResizeEdge(axis,88,options),9,'숨김 행은 크기 조절 대상에서 제외');
});
