import test from 'node:test';
import assert from 'node:assert/strict';
import { categoryTrendPoints } from '../src/chart-trend.js';
import { renderChartSvg, chartModelData } from '../src/chart.js';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, descendants } from '../src/xml.js';

const near = (a,b,eps=.11) => assert.ok(Math.abs(a-b)<=eps, `${a} ≈ ${b}`);
const data = (patch={}) => ({categories:['A','B','C','D'], series:[{name:'매출',values:[10,20,30,40],trend:'linear',trendForward:2,...patch}]});
const draw = (chart={}, input=data()) => {
  const svg=renderChartSvg({type:'column',w:500,h:320,legend:'none',...chart},input);
  assert.doesNotMatch(svg,/NaN|Infinity|undefined/);
  return parseXml(svg);
};
const trends = root => descendants(root,'path').filter(n=>n.attrs['data-trend']);
const points = node => [...node.attrs.d.matchAll(/[ML](-?[\d.e+]+),(-?[\d.e+]+)/g)].map(m=>m.slice(1).map(Number));
const ticks = (root, axis='y') => {
  const group=descendants(root,'g').find(n=>n.attrs['data-el']===`axis-${axis}`);
  return descendants(group,'text').map(n=>n.text).filter(s=>/^-?[\d,.]+%?$/.test(s));
};
const numericTicks = (root,axis='y') => ticks(root,axis).map(s=>Number(s.replaceAll(',','')));
const clip = (root, trend=trends(root)[0]) => {
  assert.ok(trend,'추세선이 렌더링됨');
  const id=trend.attrs['clip-path']?.match(/^url\(#(.+)\)$/)?.[1]; assert.ok(id);
  const node=descendants(root,'clipPath').find(n=>n.attrs.id===id); assert.ok(node);
  assert.equal(node.attrs.clipPathUnits,'userSpaceOnUse');
  const rect=descendants(node,'rect')[0]; assert.ok(rect);
  const {x,y,width:w,height:h}=rect.attrs;
  return {x:Number(x),y:Number(y),w:Number(w),h:Number(h)};
};
const labels = root => descendants(root,'text').filter(n=>n.attrs['data-axis-label']!==undefined);
const labelPosition = n => n.attrs.transform.match(/translate\(([-\d.]+) ([-\d.]+)\)/).slice(1).map(Number);
const inside = (p,a) => p[0]>=a.x-.11&&p[0]<=a.x+a.w+.11&&p[1]>=a.y-.11&&p[1]<=a.y+a.h+.11;

test('선형 예측은 원래 항목 좌표로 회귀하고 예측 끝값이 정확하다',()=>{
  const s={values:[10,null,30,40],trend:'linear',trendForward:2}; const before=structuredClone(s);
  const actual=categoryTrendPoints(s);
  assert.deepEqual(actual.map(p=>p[0]),[0,5]); near(actual[0][1],10,1e-10); near(actual[1][1],60,1e-10);
  assert.deepEqual(s,before);
  assert.deepEqual(categoryTrendPoints({values:[null,20,null,40],trend:'linear',trendForward:1.5}),[[1,20],[4.5,55]]);
});

test('음수·무효 예측구간과 유효점 부족이 무한 범위를 만들지 않는다',()=>{
  for(const trendForward of [-3,NaN,Infinity,-Infinity,'2',null,undefined]) assert.deepEqual(categoryTrendPoints({values:[10,20,30],trend:'linear',trendForward}),[[0,10],[2,30]]);
  for(const values of [[],[1],[null,2],[NaN,Infinity]]) assert.deepEqual(categoryTrendPoints({values,trend:'linear',trendForward:5}),[]);
  assert.deepEqual(categoryTrendPoints({values:[1,2,3],trend:'unknown'}),[]);
});

test('이동 평균의 결측 원래 좌표와 정수 구간을 보존하고 앞으로 옵션을 무시한다',()=>{
  assert.deepEqual(categoryTrendPoints({values:[10,null,20,30],trend:'movingAvg',trendPeriod:2.9,trendForward:100}),[[2,15],[3,25]]);
  assert.deepEqual(categoryTrendPoints({values:[10,20,30],trend:'movingAvg',trendPeriod:NaN}),[[2,20]]);
});

test('지수 예측의 정상 끝값과 큰 예측의 유한 샘플만 생성한다',()=>{
  const regular=categoryTrendPoints({values:[1,2,4],trend:'exp',trendForward:2});
  near(regular[0][1],1,1e-9); near(regular.at(-1)[0],4,1e-9); near(regular.at(-1)[1],16,1e-8);
  const large=categoryTrendPoints({values:[1,1e100,1e200],trend:'exp',trendForward:100});
  assert.ok(large.every(p=>p.every(Number.isFinite))); assert.ok(large.length<=65);
  assert.deepEqual(categoryTrendPoints({values:[-1,2,4],trend:'exp',trendForward:2}),[]);
  draw({},data({values:[1,1e100,1e200],trend:'exp',trendForward:100}));
});

for(const type of ['column','line','bar']) test(`${type}: 예측 공간과 자동값축을 확보하고 최종 그림 영역으로 자른다`,()=>{
  for(const reverse of [false,true]){
    const input=data(), before=structuredClone(input), root=draw({type,axes:{x:{reverse}}},input);
    const a=clip(root), ps=points(trends(root)[0]);
    assert.equal(ps.length,2); assert.ok(ps.every(p=>inside(p,a)));
    assert.ok(Math.max(...numericTicks(root))>=60);
    const firstSlot=reverse?5.5:.5, lastSlot=reverse?.5:5.5;
    near(type==='bar'?ps[0][1]:ps[0][0],(type==='bar'?a.y:a.x)+(type==='bar'?a.h:a.w)*firstSlot/6);
    near(type==='bar'?ps[1][1]:ps[1][0],(type==='bar'?a.y:a.x)+(type==='bar'?a.h:a.w)*lastSlot/6);
    assert.deepEqual(input,before);
  }
});

test('수동 값 축 최소·최대는 예측값 때문에 바뀌지 않고 클립으로 잘린다',()=>{
  for(const type of ['column','bar']) for(const reverse of [false,true]){
    const root=draw({type,axes:{y:{min:0,max:25,major:5,reverse}}}), a=clip(root);
    assert.deepEqual(numericTicks(root),[0,5,10,15,20,25]);
    assert.ok(points(trends(root)[0]).some(p=>!inside(p,a)),'범위 밖 원래 기하를 clipPath가 제한');
  }
});

test('뒤쪽 빈 항목 공간이 이미 있으면 필요한 만큼만 예측 범위를 쓴다',()=>{
  const input={categories:['A','B','C','D','E'],series:[{name:'매출',values:[10,20,30,null,null],trend:'linear',trendForward:2}]};
  const root=draw({},input),a=clip(root),ns=labels(root),ps=points(trends(root)[0]);
  assert.equal(ns.length,5);
  near(labelPosition(ns[0])[0],a.x+a.w*.5/5);
  near(labelPosition(ns.at(-1))[0],a.x+a.w*4.5/5);
  near(ps.at(-1)[0],a.x+a.w*4.5/5);
  const fractional=draw({},data({trendForward:1.5})),af=clip(fractional);
  near(points(trends(fractional)[0]).at(-1)[0],af.x+af.w*5/5.5);
});

test('보조축 예측은 보조값축만 확대하고 기본값축은 유지한다',()=>{
  const input={categories:['A','B','C','D'],series:[{name:'매출',values:[100,200,300,400],type:'column',axis:0},{name:'수익률',values:[2,4,6,8],type:'line',axis:1}]};
  const baseline=draw({type:'combo'},input), forecast=structuredClone(input);
  Object.assign(forecast.series[1],{trend:'linear',trendForward:3});
  const root=draw({type:'combo'},forecast);
  assert.deepEqual(ticks(root),ticks(baseline));
  assert.ok(Math.max(...numericTicks(root,'y2'))>=14);
  assert.ok(Math.max(...numericTicks(root,'y2'))>Math.max(...numericTicks(baseline,'y2')));
  const a=clip(root); assert.ok(points(trends(root)[0]).every(p=>inside(p,a)));
});

test('보조축만 있는 차트도 해당 자동 축과 예측 끝값을 사용한다',()=>{
  const root=draw({},data({axis:1})),a=clip(root);
  assert.equal(ticks(root).length,0);
  assert.ok(Math.max(...numericTicks(root,'y2'))>=60);
  assert.ok(points(trends(root)[0]).every(p=>inside(p,a)));
});

test('예측 중 데이터 표 항목·값과 막대 중심이 정방향/역방향 모두 일치한다',()=>{
  for(const reverse of [false,true]){
    const root=draw({dataTable:true,axes:{x:{reverse}}}),a=clip(root);
    assert.equal(labels(root).length,0);
    const rects=descendants(root,'rect').filter(n=>n.attrs['data-s']==='0'&&n.attrs['data-p']!==undefined);
    for(let i=0;i<4;i++){
      const header=descendants(root,'text').find(n=>n.text==='ABCD'[i]);assert.ok(header);
      const expected=a.x+a.w*((reverse?5-i:i)+.5)/6;
      near(Number(header.attrs.x),expected);
      const rect=rects.find(n=>n.attrs['data-p']===String(i));assert.ok(rect);
      near(Number(rect.attrs.x)+Number(rect.attrs.width)/2,expected,.61); // 막대의 1px 간격은 기존 렌더 정책
      const value=descendants(root,'text').find(n=>n.text===String((i+1)*10)&&Number(n.attrs.y)>a.y+a.h);
      assert.ok(value);near(Number(value.attrs.x),expected);
    }
  }
});

test('원래 범주 레이블이 막대 중심과 같고 예측의 빈 항목을 생성하지 않는다',()=>{
  for(const reverse of [false,true]){
    const root=draw({axes:{x:{reverse,labelInterval:1}}}),a=clip(root),ns=labels(root);
    assert.deepEqual(ns.map(n=>n.attrs['aria-label']),['A','B','C','D']);
    for(let i=0;i<4;i++)near(labelPosition(ns[i])[0],a.x+a.w*((reverse?5-i:i)+.5)/6);
  }
});

test('100% 누적 원래 축과 범주 간격은 raw 예측값에 오염되지 않는다',()=>{
  const input={categories:['A','B','C','D'],series:[{name:'S1',values:[100,200,300,400]},{name:'S2',values:[200,300,400,500]}]};
  const c={grouping:'percentStacked'}, baseline=draw(c,input), forecast=structuredClone(input);
  Object.assign(forecast.series[0],{trend:'linear',trendForward:20});
  const root=draw(c,forecast);
  assert.deepEqual(ticks(root),ticks(baseline));
  assert.ok(ticks(root).includes('100%'));
  assert.deepEqual(labels(root).map(labelPosition),labels(baseline).map(labelPosition));
  clip(root);
});

test('없는 추세선·이동 평균·분산형은 불필요한 예측 영역을 만들지 않는다',()=>{
  const empty=draw({},data({trend:undefined})); assert.equal(trends(empty).length,0);
  const avg0=draw({},data({trend:'movingAvg',trendPeriod:2,trendForward:0}));
  const avg100=draw({},data({trend:'movingAvg',trendPeriod:2,trendForward:100}));
  assert.deepEqual(labels(avg0).map(labelPosition),labels(avg100).map(labelPosition));
  assert.deepEqual(ticks(avg0),ticks(avg100));
  const xy=data({x:[1,10,50,100]});
  const scatter=draw({type:'scatter'},xy); assert.equal(trends(scatter).length,0,'기존 XY 구현 범위를 임의 확장하지 않음');
});

test('표준 XLSX 예측구간을 불러와도 같은 자동 범위와 클립이 적용된다',()=>{
  const wb=new Workbook();
  [['항목','매출'],['A',10],['B',20],['C',30],['D',40]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));
  const ch={id:'forecast-native',type:'column',x:0,y:0,w:500,h:320,legend:'none',range:{r1:0,c1:0,r2:4,c2:1},seriesFmt:[{trend:'linear',trendForward:2}]};
  wb.sheets[0].charts=[ch];
  const before=draw(ch,chartModelData(wb,0,ch)),files=unzip(writeXlsx(wb)),path='xl/charts/chart1.xml';
  const xml=typeof files[path]==='string'?files[path]:textOf(files[path]);
  assert.match(xml,/<c:forward val="2"\/>/);
  files[path]=xml.replace(/<c:extLst>[\s\S]*?<\/c:extLst>/g,'');
  const restored=new Workbook(readXlsx(zip(files)).data),back=restored.sheets[0].charts[0];
  assert.equal(back.seriesFmt[0].trend,'linear');assert.equal(back.seriesFmt[0].trendForward,2);
  const after=draw(back,chartModelData(restored,0,back));
  assert.deepEqual(ticks(after),ticks(before));assert.deepEqual(clip(after),clip(before));
  assert.deepEqual(points(trends(after)[0]),points(trends(before)[0]));
});


test('정확한 눈금 경계의 예측값이 부동소수 오차로 축을 한 눈금 더 늘리지 않는다',()=>{
  const root=draw({},data({values:[10,null,30,40]}));
  assert.equal(Math.max(...numericTicks(root)),60);
});
