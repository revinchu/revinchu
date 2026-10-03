import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { applySlicerSettings, slicerSettingsKey } from '../src/slicer-settings.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
const refs=ids=>ids.map(i=>({sheet:'Report'+i,name:'Pivot'+i}));
const slicer=(id,ids,extra={})=>({id,caption:id,source:{kind:'pivot',field:'Group',pivots:refs(ids)},x:20,y:20,w:160,h:200,style:'SlicerStyleLight1',...extra});
function fixture() {
  const wb=new Workbook({sheets:[{name:'Source',cells:{'0,0':{raw:'Group'},'0,1':{raw:'Value'},'1,0':{raw:'A'},'1,1':{raw:'10'},'2,0':{raw:'B'},'2,1':{raw:'20'}}},...[1,2,3].map(i=>({name:'Report'+i,cells:{},pivot:{name:'Pivot'+i,source:'Source',range:{r1:0,c1:0,r2:2,c2:1},rows:['Group'],values:[{field:'Value'}],top:0,left:0},slicers:[]}))]});
  // 앱과 같은 참조 해석 계약; 원본 셀 계산은 금지합니다.
  const resolve=(source,host)=> {
    const links=source.pivots?.length?source.pivots:[{sheet:source.self?null:source.sheet}];
    return links.map(link=> { const si=link.sheet?wb.sheetIndexByName(link.sheet):host, def=wb.sheets[si]?.pivot; return def && (!link.name || link.name.toLowerCase()===def.name.toLowerCase()) ? {si,def,index:-1}:null; }).filter(Boolean);
  };
  return {wb,resolve};
}
const options={sort:'desc',customList:false,hideNoData:true,markNoData:false,noDataLast:false,showDeleted:true};
test('같은 피벗 캐시의 옵션만 모든 시트로 전파하고 한 번의 Undo/Redo로 복원', () => {
  const {wb,resolve}=fixture();
  wb.sheets[1].slicers=[slicer('first',[1,2]),slicer('different-field',[1,2],{source:{kind:'pivot',field:'Value',pivots:refs([1,2])}})];
  wb.sheets[2].slicers=[slicer('second',[2,1,1],{caption:'독립 캡션',x:230,style:'SlicerStyleDark2',showHeader:false,locked:true})];
  wb.sheets[3].slicers=[slicer('different-links',[3])];
  const old=wb.sheets.map(s=>structuredClone(s.slicers)), undo=wb.undoStack.length;
  for(const s of wb.sheets)for(const sl of s.slicers??[])Object.freeze(sl);
  assert.equal(applySlicerSettings(wb,1,'first',{...options,caption:'새 캡션',showHeader:false,noMove:true},resolve),true);
  const first=wb.sheets[1].slicers[0], second=wb.sheets[2].slicers[0];
  for(const [key,value] of Object.entries(options)){assert.equal(first[key],value);assert.equal(second[key],value);}
  assert.equal(first.caption,'새 캡션'); assert.equal(first.noMove,true);
  for(const key of ['caption','x','style','showHeader','locked','noMove'])assert.equal(second[key],old[2][0][key]);
  assert.deepEqual(wb.sheets[1].slicers[1],old[1][1]);assert.deepEqual(wb.sheets[3].slicers,old[3]);assert.equal(wb.undoStack.length,undo+1);
  const after=wb.sheets.map(s=>structuredClone(s.slicers)); wb.undo(); assert.deepEqual(wb.sheets.map(s=>s.slicers),old);wb.redo();assert.deepEqual(wb.sheets.map(s=>s.slicers),after);
  const saved=readXlsx(writeXlsx(wb)).data.sheets;
  for(const si of [1,2])for(const [key,value]of Object.entries(options))if(key!=='noDataLast')assert.equal(saved[si].slicers[0][key],value,key);
  // crossFilter=none이면 마지막 배치 옵션은 비활성 상태여서 OOXML에 따로 저장되지 않습니다.
  applySlicerSettings(wb,1,'first',{markNoData:undefined,noDataLast:false},resolve);
  const withMarking=readXlsx(writeXlsx(wb)).data.sheets;
  for(const si of [1,2])assert.equal(withMarking[si].slicers[0].noDataLast,false);
});
test('우클릭 오름차순의 undefined도 공유되며 모양만 수정할 때는 캐시를 변경하지 않음', () => {
  const {wb,resolve}=fixture();wb.sheets[1].slicers=[slicer('first',[1],{sort:'desc'}),slicer('peer',[1],{sort:'desc'})];
  applySlicerSettings(wb,1,'first',{sort:undefined},resolve);assert.equal(wb.sheets[1].slicers[1].sort,undefined);
  const peer=wb.sheets[1].slicers[1];applySlicerSettings(wb,1,'first',{caption:'다름',style:'SlicerStyleDark2'},resolve);assert.equal(wb.sheets[1].slicers[1],peer);
  const undo=wb.undoStack.length;assert.equal(applySlicerSettings(wb,1,'missing',options,resolve),false);assert.equal(wb.undoStack.length,undo);
});
test('독립 캐시와 표 열은 별도 종류로 구분하고 다른 필드·연결 없는 개체는 합치지 않음', () => {
  const {wb,resolve}=fixture(), sources=[{kind:'cache',cacheKey:'saved',field:'Group'},{kind:'cache',cacheKey:'other',field:'Group'},{kind:'cache',cacheKey:'saved',field:'Value'},{kind:'table',table:'Data',column:'Group'},{kind:'table',table:'Data',column:'Value'},{kind:'pivot',field:'Group',pivots:[{sheet:'Missing',name:'Missing'}]}];
  wb.sheets[1].slicers=sources.map((source,i)=>slicer('s'+i,[],{source}));wb.sheets[2].slicers=[slicer('peer',[],{source:structuredClone(sources[0])}),slicer('table-peer',[],{source:{kind:'table',table:'data',column:'group'}})];
  applySlicerSettings(wb,1,'s0',{sort:'desc'},resolve);assert.equal(wb.sheets[2].slicers[0].sort,'desc');for(const sl of wb.sheets[1].slicers.slice(1))assert.equal(sl.sort,undefined);
  applySlicerSettings(wb,1,'s3',{hideNoData:true},resolve);assert.equal(wb.sheets[2].slicers[1].hideNoData,true);assert.equal(wb.sheets[1].slicers[4].hideNoData,undefined);
  assert.equal(slicerSettingsKey(sources[5],1,resolve),null);assert.equal(slicerSettingsKey({kind:'cache',field:'Group'},1,resolve),null);
});
test('대소문자·연결 순서·옛 self 형식은 같은 실제 피벗 참조로 정규화', () => {
  const {resolve}=fixture();assert.equal(slicerSettingsKey({kind:'pivot',field:'group',pivots:[{sheet:'report1',name:'pivot1'}]},1,resolve),slicerSettingsKey({field:'Group',self:true},1,resolve));
  assert.notEqual(slicerSettingsKey({field:'Group',self:true},1,resolve),slicerSettingsKey({field:'Group',self:true},2,resolve));
});
