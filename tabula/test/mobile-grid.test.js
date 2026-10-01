import { test } from 'node:test';
import assert from 'node:assert/strict';
import { TouchGridGesture, gridTouchIgnored } from '../src/mobile-grid.js';
const p = (x,y,id=1) => ({x,y,id});
function fixture() {
  const calls=[];let context={},zoom=100;
  const host={context:()=>context,hit:(x,y)=>({zone:x<0?'rowHeader':y<0?'colHeader':'cell',r:Math.floor(y/20),c:Math.floor(x/80)}),
    begin:()=>{calls.push(['begin']);},zoom:()=>zoom,
    tap:h=>calls.push(['tap',h]),edit:h=>calls.push(['edit',h]),range:(a,b)=>calls.push(['range',a,b]),
    scroll:(x,y)=>calls.push(['scroll',x,y]),pinch:(z,a,b)=>{calls.push(['pinch',z,a,b]);zoom=z;},end:()=>calls.push(['end'])};
  return {g:new TouchGridGesture(host),calls,host,switchContext:()=>{context={};}};
}
test('탭은 선택하고 같은 셀의 빠른 두 번째 탭만 편집한다',()=>{
  const {g,calls}=fixture();g.start([p(10,30)],0);g.end([],100);assert.equal(calls.filter(x=>x[0]==='tap').length,1);assert.equal(calls.some(x=>x[0]==='edit'),false);
  g.start([p(12,32)],220);g.end([],270);assert.equal(calls.filter(x=>x[0]==='edit').length,1);
  g.start([p(92,32)],330);g.end([],360);assert.equal(calls.filter(x=>x[0]==='edit').length,1);
});
test('8px를 넘는 한 손가락 이동은 선택·편집 대신 스크롤한다',()=>{
  const {g,calls}=fixture();g.start([p(110,230)],0);g.move([p(106,228)],30);assert.equal(calls.some(x=>x[0]==='scroll'),false);
  g.move([p(105,190)],50);g.move([p(80,160)],70);g.end([],90);
  assert.deepEqual(calls.filter(x=>x[0]==='scroll'),[['scroll',1,38],['scroll',25,30]]);assert.equal(calls.some(x=>['tap','edit','range'].includes(x[0])),false);
});
test('길게 누른 뒤 끌기는 같은 시작 셀을 유지하는 범위를 만들고 탭하지 않는다',()=>{
  const {g,calls}=fixture();g.start([p(10,30)],0);assert.equal(g.hold(449),false);assert.equal(g.hold(450),true);g.move([p(200,170)],500);g.end([],600);
  const ranges=calls.filter(x=>x[0]==='range');assert.equal(ranges.length,2);assert.deepEqual(ranges[1].slice(1),[{zone:'cell',r:1,c:0},{zone:'cell',r:8,c:2}]);assert.equal(calls.some(x=>['tap','edit','scroll'].includes(x[0])),false);
});
test('행·열 머리글 길게 선택과 corner 탭을 구분한다',()=>{
  const {g,calls,host}=fixture();g.start([p(-1,30)],0);g.hold(500);g.move([p(-1,90)],520);g.end([],540);assert.equal(calls.find(x=>x[0]==='range')[1].zone,'rowHeader');
  host.hit=()=>({zone:'corner',r:0,c:0});g.start([p(0,0)],700);assert.equal(g.hold(1300),false);g.end([],1400);assert.equal(calls.at(-2)[0],'tap');
});
test('핀치는 초기 거리 비율·25~400% 제한을 적용하고 손가락을 떼도 탭하지 않는다',()=>{
  const {g,calls}=fixture();g.start([p(100,200),p(200,200,2)],0);g.move([p(50,200),p(250,200,2)],20);assert.equal(calls.at(-1)[1],200);
  g.move([p(0,200),p(800,200,2)],40);assert.equal(calls.at(-1)[1],400);g.move([p(100,200),p(101,200,2)],60);assert.equal(calls.at(-1)[1],25);
  g.end([p(100,200)],80);g.move([p(200,200)],90);g.end([],100);assert.equal(calls.some(x=>['tap','edit','scroll'].includes(x[0])),false);
});
test('탭 대기·스크롤 중 두 번째 손가락은 핀치로 전환하며 길게 선택하지 않는다',()=>{
  const {g,calls}=fixture();g.start([p(100,200)],0);g.move([p(100,100)],100);g.start([p(100,100),p(200,100,2)],200);assert.equal(g.hold(700),false);g.move([p(75,100),p(225,100,2)],750);g.end([],800);assert.equal(calls.filter(x=>x[0]==='pinch').at(-1)[1],150);assert.equal(calls.some(x=>x[0]==='range'),false);
});
test('문서 교체·시트 교체·touchcancel은 다음 문서에 선택이나 편집을 남기지 않는다',()=>{
  const {g,calls,switchContext}=fixture();g.start([p(10,30)],0);switchContext();g.hold(500);g.end([],600);assert.equal(calls.some(x=>['tap','edit','range'].includes(x[0])),false);
  g.start([p(10,30)],700);g.end([],800,true);assert.equal(calls.some(x=>x[0]==='tap'),false);g.start([p(10,30)],850);g.end([],900);assert.equal(calls.some(x=>x[0]==='edit'),false);
});
test('편집 완료 거절과 비격자 영역은 제스처를 시작하지 않는다',()=>{
  const {g,host}=fixture();host.begin=()=>false;assert.equal(g.start([p(10,30)],0),false);assert.equal(g.phase,'idle');host.hit=()=>({zone:'outline'});assert.equal(g.start([p(10,30)],0),false);
});
test('입력 중 셀 편집기·개체·필터는 제외하고 idle 셀 편집기는 그리드로 처리한다',()=>{
  assert.equal(gridTouchIgnored({id:'cellEditor',classList:{contains:x=>x==='idle'},closest:()=>true}),false);
  assert.equal(gridTouchIgnored({id:'cellEditor',classList:{contains:()=>false},closest:()=>true}),true);
  assert.equal(gridTouchIgnored({closest:()=>false}),false);assert.equal(gridTouchIgnored({closest:()=>true}),true);
});
test('확인란·서식복사 탭 false 결과는 두 번 탭 편집으로 이어지지 않는다',()=>{
 const {g,host,calls}=fixture();host.tap=h=>{calls.push(['tap',h]);return false;};g.start([p(10,30)],0);g.end([],50);g.start([p(10,30)],100);g.end([],150);assert.equal(calls.filter(x=>x[0]==='tap').length,2);assert.equal(calls.some(x=>x[0]==='edit'),false);
});
test('범위 완료 콜백은 실제 끝남과 취소·문서 교체를 구별한다',()=>{
 const {g,host,switchContext}=fixture(),ends=[];host.end=(phase,cancelled)=>ends.push([phase,cancelled]);
 g.start([p(10,30)],0);g.hold(500);g.end([],600);assert.deepEqual(ends.at(-1),['range',false]);
 g.start([p(10,30)],700);g.hold(1200);g.end([],1300,true);assert.deepEqual(ends.at(-1),['range',true]);
 g.start([p(10,30)],1400);g.hold(1900);switchContext();g.end([],2000);assert.deepEqual(ends.at(-1),['range',true]);
});
