import test from 'node:test';
import assert from 'node:assert/strict';
import { CellMap } from '../src/cellmap.js';
import { capturePivotCaches, restorePivotCaches, shiftPivotCaches } from '../src/pivot-structure-cache.js';

function fixture({ si=0, name='소재 피벗', top=10, left=4, extra=false }={}) {
  const def={name,top,left,area:{r1:top,c1:left,r2:top+2,c2:left+1}},sheet=extra?{pivot:null,pivotsExtra:[def]}:{pivot:def,pivotsExtra:[]};
  const roles=[['rowHead:0','colHead'],['rowItem:0','data:0'],['grandLabel','grandData:0']], rowDepth=[-1,0,-1];
  const layout={top,left,roles,rowDepth,rows:['소재'],cols:[],values:['합계'],sheet,definition:def};
  const written=new CellMap(), styles=[{bold:true},{numFmt:'number',decimals:0},{color:'#112233'}];
  written.setRC(top,left,styles[0]);written.setRC(top+1,left+1,styles[1]);written.setRC(top+2,left+1,styles[2]);
  const key=si+':'+name,wkey=key+':'+top+','+left;
  return {si,def,sheet,roles,rowDepth,layout,styles,key,wkey,layouts:new Map([[key,layout]]),written:new Map([[wkey,written]])};
}
const nextSheet=(f,patch={})=>({pivot:f.sheet.pivot?{...f.def,...patch}:null,pivotsExtra:f.sheet.pivotsExtra.map(def=>({...def,...patch}))});
const coords=map=>{const out=[];map.forEachRC((value,r,c)=>out.push([r,c,value]));return out.sort((a,b)=>a[0]-b[0]||a[1]-b[1]);};

for(const axis of ['row','col'])test('피벗 앞 '+axis+' 삽입은 역할 배열과 서식 identity를 유지하며 절대 좌표를 옮긴다',()=>{
  const f=fixture(),after=nextSheet(f,{[axis==='row'?'top':'left']:(axis==='row'?f.def.top:f.def.left)+3});
  shiftPivotCaches(f.layouts,f.written,{axis,index:2,count:3},f.si,f.sheet,after);
  const layout=f.layouts.get(f.key),styles=f.written.get(f.key+':'+layout.top+','+layout.left);
  assert.equal(layout.roles,f.roles);assert.equal(layout.rowDepth,f.rowDepth);
  assert.equal(layout.top,axis==='row'?13:10);assert.equal(layout.left,axis==='col'?7:4);
  assert.equal(layout.sheet,after);assert.equal(layout.definition,after.pivot);
  assert.equal(styles.getRC(axis==='row'?14:11,axis==='col'?8:5),f.styles[1]);
  assert.equal(f.written.has(f.wkey),false);assert.equal(f.layout.top,10);assert.equal(f.layout.left,4);
});

test('피벗보다 앞의 삭제는 새 시작 셀로 이동하며 마지막 생성 서식을 수동 서식과 구분할 수 있다',()=>{
  const f=fixture(),after=nextSheet(f,{top:8}),manual={...f.styles[1],color:'#ff0000'};
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:2,count:-2},f.si,f.sheet,after);
  const styles=f.written.get(f.key+':8,4');assert.equal(styles.getRC(9,5),f.styles[1]);
  assert.notDeepEqual(manual,styles.getRC(9,5));assert.equal(manual.color,'#ff0000');
  assert.equal(f.layouts.get(f.key).roles,f.roles);
});

test('피벗 내부 행 삽입은 빈 역할과 깊이만 끼우고 이후 셀 서식을 옮긴다',()=>{
  const f=fixture(),after=nextSheet(f);
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:11,count:2},0,f.sheet,after);
  const layout=f.layouts.get(f.key),styles=f.written.get(f.wkey);
  assert.deepEqual(layout.roles,[f.roles[0],[],[],f.roles[1],f.roles[2]]);
  assert.deepEqual(layout.rowDepth,[-1,-1,-1,0,-1]);
  assert.equal(layout.roles[0],f.roles[0]);assert.equal(layout.roles[3],f.roles[1]);
  assert.equal(styles.getRC(11,5),undefined);assert.equal(styles.getRC(13,5),f.styles[1]);
  assert.equal(styles.getRC(14,5),f.styles[2]);assert.equal(f.roles.length,3);
});

test('피벗 내부 행 삭제는 삭제된 역할·서식을 제거하고 나머지를 압축한다',()=>{
  const f=fixture(),after=nextSheet(f);
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:11,count:-1},0,f.sheet,after);
  assert.deepEqual(f.layouts.get(f.key).roles,[f.roles[0],f.roles[2]]);
  assert.deepEqual(f.layouts.get(f.key).rowDepth,[-1,-1]);
  assert.equal(f.written.get(f.wkey).getRC(11,5),f.styles[2]);
  assert.equal(f.written.get(f.wkey).getRC(12,5),undefined);
});

test('피벗 시작 행 일부를 삭제하면 남은 역할의 시작 좌표를 삭제 위치로 옮긴다',()=>{
  const f=fixture(),after=nextSheet(f,{top:8});
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:8,count:-3},0,f.sheet,after);
  assert.equal(f.layouts.get(f.key).top,8);assert.deepEqual(f.layouts.get(f.key).roles,[f.roles[1],f.roles[2]]);
  assert.equal(f.written.get(f.key+':8,4').getRC(8,5),f.styles[1]);
});

test('내부 열 삽입·삭제는 행마다 실제 역할 폭에 적용하고 깊이를 유지한다',()=>{
  const f=fixture(),after=nextSheet(f);f.layout.roles=[['pageField'],...f.roles];
  shiftPivotCaches(f.layouts,f.written,{axis:'col',index:5,count:1},0,f.sheet,after);
  const inserted=f.layouts.get(f.key);
  assert.deepEqual(inserted.roles,[['pageField'],['rowHead:0','','colHead'],['rowItem:0','','data:0'],['grandLabel','','grandData:0']]);
  assert.equal(inserted.roles[0],f.layout.roles[0]);assert.equal(inserted.rowDepth,f.rowDepth);
  assert.equal(f.written.get(f.wkey).getRC(11,6),f.styles[1]);
  shiftPivotCaches(f.layouts,f.written,{axis:'col',index:5,count:-1},0,after,after);
  assert.deepEqual(f.layouts.get(f.key).roles,f.layout.roles);assert.equal(f.written.get(f.wkey).getRC(11,5),f.styles[1]);
});

test('피벗 전체가 삭제되거나 해당 슬롯이 사라지면 다른 시트 캐시를 남기고 제거한다',()=>{
  for(const removeSlot of [false,true]) {
    const f=fixture(),other=fixture({si:1});f.layouts.set(other.key,other.layout);f.written.set(other.wkey,other.written.get(other.wkey));
    shiftPivotCaches(f.layouts,f.written,{axis:'row',index:10,count:removeSlot?1:-3},0,f.sheet,removeSlot?{pivot:null,pivotsExtra:[]}:nextSheet(f));
    assert.equal(f.layouts.has(f.key),false);assert.equal(f.written.has(f.wkey),false);
    assert.equal(f.layouts.get(other.key),other.layout);assert.equal(f.written.get(other.wkey),other.written.get(other.wkey));
  }
});

test('삽입·Undo·Redo의 스냅샷은 복제된 시트 정의에 재연결하고 수동 비교 기준을 보존한다',()=>{
  const f=fixture(),before=capturePivotCaches(f.layouts,f.written,0),afterSheet=nextSheet(f,{top:13});
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:7,count:3},0,f.sheet,afterSheet);
  const after=capturePivotCaches(f.layouts,f.written,0);
  const undoSheet=structuredClone(f.sheet);restorePivotCaches(f.layouts,f.written,0,before,undoSheet);
  assert.equal(f.layouts.get(f.key).sheet,undoSheet);assert.equal(f.layouts.get(f.key).definition,undoSheet.pivot);
  assert.equal(f.layouts.get(f.key).roles,f.roles);assert.equal(f.written.get(f.wkey).getRC(11,5),f.styles[1]);
  const redoSheet=structuredClone(afterSheet);restorePivotCaches(f.layouts,f.written,0,after,redoSheet);
  assert.equal(f.layouts.get(f.key).definition,redoSheet.pivot);assert.equal(f.layouts.get(f.key).top,13);
  assert.equal(f.written.get(f.key+':13,4').getRC(14,5),f.styles[1]);assert.equal(f.written.has(f.wkey),false);
  assert.equal(before.layouts.get(f.key).top,10);assert.equal(before.written.get(f.wkey).getRC(11,5),f.styles[1]);
});

test('복원은 스냅샷에 없는 캐시를 제거하고 빈 캐시를 임의 서식으로 채우지 않는다',()=>{
  const f=fixture(),snapshot=capturePivotCaches(f.layouts,new Map(),0),other=fixture({si:10});
  f.layouts.set('0:새 피벗',{});f.written.set('0:새 피벗:0,0',new CellMap());
  f.layouts.set(other.key,other.layout);f.written.set(other.wkey,other.written.get(other.wkey));
  restorePivotCaches(f.layouts,f.written,0,snapshot,structuredClone(f.sheet));
  assert.equal(f.layouts.has('0:새 피벗'),false);assert.equal(f.written.has(f.wkey),false);
  assert.equal(f.written.has('0:새 피벗:0,0'),false);assert.equal(f.layouts.get(other.key),other.layout);
  assert.equal(f.written.get(other.wkey),other.written.get(other.wkey));
});

test('이름 없는 추가 피벗도 슬롯을 기억해 Undo의 정확한 정의에 재연결한다',()=>{
  const f=fixture({name:'',extra:true});f.sheet.pivot={name:'',top:1,left:0};
  const snapshot=capturePivotCaches(f.layouts,f.written,0),restored=structuredClone(f.sheet);
  restorePivotCaches(f.layouts,f.written,0,snapshot,restored);
  assert.equal(f.layouts.get(f.key).definition,restored.pivotsExtra[0]);
  assert.notEqual(f.layouts.get(f.key).definition,restored.pivot);
  restored.pivotsExtra[0].name='다른 피벗';restorePivotCaches(f.layouts,f.written,0,snapshot,restored);
  assert.equal(f.layouts.size,0);assert.equal(f.written.size,0);
});

test('콜론을 포함한 피벗 이름과 이동 대상 밖의 피벗은 캐시 키 충돌 없이 보존한다',()=>{
  const f=fixture({name:'소재:비교'}),after=nextSheet(f);
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:50,count:3},0,f.sheet,after);
  assert.equal(f.layouts.get(f.key).top,10);assert.equal(f.layouts.get(f.key).roles,f.roles);
  assert.equal(f.written.get(f.wkey).getRC(11,5),f.styles[1]);
});

test('동결한 역할 배열과 서식 객체는 캐시 이동·복원에서 변경하지 않는다',()=>{
  const f=fixture();for(const row of f.roles)Object.freeze(row);Object.freeze(f.roles);Object.freeze(f.rowDepth);for(const style of f.styles)Object.freeze(style);
  const before=capturePivotCaches(f.layouts,f.written,0);
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:11,count:1},0,f.sheet,nextSheet(f));
  restorePivotCaches(f.layouts,f.written,0,before,structuredClone(f.sheet));
  assert.equal(f.written.get(f.wkey).getRC(11,5),f.styles[1]);assert.equal(f.layouts.get(f.key).roles,f.roles);
});

test('잘못된 축과 인덱스는 캐시를 변경하지 않는다',()=>{
  const f=fixture(),styles=f.written.get(f.wkey);
  for(const shift of [{axis:'other',index:2,count:1},{axis:'row',index:-1,count:1},{axis:'row',index:1.5,count:1},{axis:'row',index:1,count:.5}])shiftPivotCaches(f.layouts,f.written,shift,0,f.sheet,f.sheet);
  assert.equal(f.layouts.get(f.key),f.layout);assert.equal(f.written.get(f.wkey),styles);
});

test('호출자가 준 현재 시트의 슬롯을 저장해 오래된 캐시 sheet 참조에 의존하지 않는다',()=>{
  const f=fixture({extra:true});f.layout.sheet={pivot:{name:f.def.name},pivotsExtra:[]};
  const snapshot=capturePivotCaches(f.layouts,f.written,0,f.sheet),restored=structuredClone(f.sheet);
  restorePivotCaches(f.layouts,f.written,0,snapshot,restored);
  assert.equal(f.layouts.get(f.key).definition,restored.pivotsExtra[0]);
});

test('앞의 추가 피벗 전체 삭제로 슬롯이 압축되어도 뒤 피벗의 역할과 서식은 이동한다',()=>{
  const first=fixture({name:'삭제 대상',top:5,extra:true}),second=fixture({name:'생존 피벗',top:12,extra:true});
  const sheet={pivot:null,pivotsExtra:[first.def,second.def]};first.layout.sheet=sheet;second.layout.sheet=sheet;
  const layouts=new Map([[first.key,first.layout],[second.key,second.layout]]),written=new Map([...first.written,...second.written]);
  const before=capturePivotCaches(layouts,written,0,sheet),afterSheet={pivot:null,pivotsExtra:[{...second.def,top:9}]};
  shiftPivotCaches(layouts,written,{axis:'row',index:5,count:-3},0,sheet,afterSheet);
  assert.equal(layouts.has(first.key),false);assert.equal(layouts.get(second.key).definition,afterSheet.pivotsExtra[0]);
  assert.equal(layouts.get(second.key).roles,second.roles);assert.equal(written.get(second.key+':9,4').getRC(10,5),second.styles[1]);
  const after=capturePivotCaches(layouts,written,0,afterSheet),undoSheet=structuredClone(sheet);
  restorePivotCaches(layouts,written,0,before,undoSheet);assert.equal(layouts.get(second.key).definition,undoSheet.pivotsExtra[1]);
  const redoSheet=structuredClone(afterSheet);restorePivotCaches(layouts,written,0,after,redoSheet);
  assert.equal(layouts.get(second.key).definition,redoSheet.pivotsExtra[0]);assert.equal(written.get(second.key+':9,4').getRC(10,5),second.styles[1]);
});

test('Undo 스냅샷은 이전 시트·정의 객체 대신 슬롯 정보와 서식만 보관한다',()=>{
  const f=fixture(),snapshot=capturePivotCaches(f.layouts,f.written,0,f.sheet),saved=snapshot.layouts.get(f.key);
  assert.equal(Object.hasOwn(saved,'sheet'),false);assert.equal(Object.hasOwn(saved,'definition'),false);
  assert.equal(saved.roles,f.roles);assert.equal(saved.rowDepth,f.rowDepth);
  assert.deepEqual(snapshot.slots.get(f.key),{prop:'pivot',index:-1,name:f.def.name});
  const restored=structuredClone(f.sheet);restorePivotCaches(f.layouts,f.written,0,snapshot,restored);
  assert.equal(f.layouts.get(f.key).sheet,restored);assert.equal(f.layouts.get(f.key).definition,restored.pivot);
  assert.equal(f.written.get(f.wkey).getRC(11,5),f.styles[1]);
  assert.equal(Object.hasOwn(saved,'sheet'),false);assert.equal(Object.hasOwn(saved,'definition'),false);
});

test('일반 피벗 Undo가 복원한 정의와 다른 오래된 역할은 행 이동으로 유효하게 만들지 않는다',()=>{
  const f=fixture(),restored={pivot:structuredClone(f.def),pivotsExtra:[]};
  const snapshot=capturePivotCaches(f.layouts,f.written,0,restored);
  assert.equal(snapshot.layouts.size,0);assert.equal(snapshot.written.size,0);
  const after={pivot:{...restored.pivot,top:13},pivotsExtra:[]};
  shiftPivotCaches(f.layouts,f.written,{axis:'row',index:2,count:3},0,restored,after);
  assert.equal(f.layouts.size,0);assert.equal(f.written.size,0);
});

test('명시한 현재 정의의 시작 좌표와 다른 역할 및 소유 레이아웃 없는 서식은 캡처하지 않는다',()=>{
  const f=fixture();f.written.set(f.key+':9,4',new CellMap().setRC(9,4,{color:'#ff0000'}));
  const valid=capturePivotCaches(f.layouts,f.written,0,f.sheet);
  assert.deepEqual([...valid.written.keys()],[f.wkey]);
  for(const axis of ['top','left']) {
    f.layout[axis]++;const snapshot=capturePivotCaches(f.layouts,f.written,0,f.sheet);
    assert.equal(snapshot.layouts.size,0);assert.equal(snapshot.written.size,0);f.layout[axis]--;
  }
  f.layout.top++;assert.equal(capturePivotCaches(f.layouts,f.written,0).layouts.size,1,'명시하지 않은 시트의 좌표는 재검사하지 않음');
});
