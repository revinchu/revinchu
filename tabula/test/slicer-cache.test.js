import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { unzip, zip } from '../src/zip.js';
import { cachedSlicerItems, applyCachedSlicerSelection } from '../src/slicer-cache.js';
const source = () => ({kind:'cache',field:'월',cacheKey:'original-cache-9',cacheSource:{name:'원본표',external:'file:///unavailable/source.xlsx'},format:{numFmt:'custom',code:'0"월"'},values:[1,2,3,null,true,{error:'#N/A'},'문자'],items:[{index:0,hasData:true},{index:1,hasData:true},{index:2,hasData:true},{index:3,hasData:false},{index:4,hasData:true},{index:5,hasData:false},{index:6,hasData:true}]});
const slicer = (id='s1') => ({id,caption:'월',source:source(),cacheSelection:['0','2'],x:20,y:20,w:180,h:230,style:'SlicerStyleLight1',columns:1});
test('연결 없는 슬라이서는 캐시 항목·선택·형식을 표시하고 외부 원본을 읽지 않음', () => {
  const items = cachedSlicerItems(slicer());
  assert.deepEqual(items.map(x=>x.text), ['1월','2월','3월','(비어 있음)','TRUE','#N/A','문자']);
  assert.deepEqual(items.filter(x=>x.selected).map(x=>x.key), ['0','2']);
  assert.equal(items[3].hasData,false);
});
test('독립 슬라이서 선택은 같은 캐시를 쓰는 슬라이서만 갱신하고 실행 취소 가능', () => {
  const wb=new Workbook(); wb.sheets[0].slicers=[slicer()];
  const other=wb.addSheet('다른 시트');wb.sheets[other].slicers=[slicer('s2'),{...slicer('s3'),source:{...source(),cacheKey:'unrelated'}}];
  wb.transact(()=>wb.setInput(0,0,0,'original'));const before=wb.undoStack.length;
  applyCachedSlicerSelection(wb,source(),['1','2']);
  assert.deepEqual(wb.sheets[0].slicers[0].cacheSelection,['1','2']);
  assert.deepEqual(wb.sheets[other].slicers[0].cacheSelection,['1','2']);
  assert.deepEqual(wb.sheets[other].slicers[1].cacheSelection,['0','2']);
  assert.equal(wb.getRaw(0,0,0),'original');assert.equal(wb.undoStack.length,before+1);
  applyCachedSlicerSelection(wb,source(),null);assert.equal(wb.sheets[0].slicers[0].cacheSelection,null);
  wb.undo();assert.deepEqual(wb.sheets[0].slicers[0].cacheSelection,['1','2']);
  wb.undo();assert.deepEqual(wb.sheets[other].slicers[0].cacheSelection,['0','2']);
});
test('독립 슬라이서 표준 XLSX 왕복: 공유 캐시·선택·형식·위치·외부 원본 메타 보존', () => {
  const wb=new Workbook();wb.sheets[0].slicers=[slicer(),{...slicer('s2'),x:220}];
  const bytes=writeXlsx(wb), parts=unzip(bytes), text=p=>new TextDecoder().decode(parts[p]);
  assert.equal(Object.keys(parts).filter(p=>/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(p)).length,1);
  assert.equal(Object.keys(parts).filter(p=>/^xl\/pivotTables\//.test(p)).length,0);
  assert.equal(Object.keys(parts).filter(p=>/^xl\/slicerCaches\/slicerCache\d+\.xml$/.test(p)).length,1);
  assert.match(text('xl/pivotCache/pivotCacheDefinition1.xml'),/refreshOnLoad="0"/);
  assert.match(text('xl/pivotCache/pivotCacheDefinition1.xml'),/containsInteger="1" minValue="1" maxValue="3"/);
  assert.match(text('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels'),/TargetMode="External"/);
  const result=readXlsx(bytes);assert.deepEqual(result.warnings,[]);
  const list=result.data.sheets[0].slicers;assert.equal(list.length,2);
  assert.equal(list[0].source.cacheKey,list[1].source.cacheKey);
  assert.deepEqual(list[0].cacheSelection,['0','2']);
  assert.deepEqual(list[0].source.values,source().values);
  assert.equal(list[0].source.cacheSource.external,source().cacheSource.external);
  assert.equal(list[1].x,220);assert.equal(cachedSlicerItems(list[0])[0].text,'1월');
  list[0].cacheSelection=null;list[1].cacheSelection=null;
  const second=readXlsx(writeXlsx(new Workbook(result.data)));
  assert.equal(second.data.sheets[0].slicers[0].cacheSelection,null);
});
test('슬라이서 캐시 ID는 workbook ID와 달라도 x14 별도 ID로 읽음', () => {
  const wb=new Workbook();wb.sheets[0].slicers=[slicer()];const parts=unzip(writeXlsx(wb));
  for(const path of ['xl/pivotCache/pivotCacheDefinition1.xml','xl/slicerCaches/slicerCache1.xml'])parts[path]=new TextEncoder().encode(new TextDecoder().decode(parts[path]).replaceAll('pivotCacheId="1"','pivotCacheId="878493998"'));
  const result=readXlsx(zip(parts));assert.deepEqual(result.warnings,[]);assert.equal(result.data.sheets[0].slicers.length,1);assert.deepEqual(result.data.sheets[0].slicers[0].cacheSelection,['0','2']);
});

test('showMissing은 생략·true·1을 켬, false·0을 끔으로 읽고 재저장', () => {
  const wb=new Workbook();wb.sheets[0].slicers=[{...slicer(),showDeleted:true}];
  const original=unzip(writeXlsx(wb)), path='xl/slicerCaches/slicerCache1.xml';
  const originalXml=new TextDecoder().decode(original[path]);
  for(const setting of [undefined,'true','1','false','0']) {
    const parts={...original}, xml=originalXml.replace(/ showMissing="[^"]*"/g,'').replace('<tabular ',`<tabular ${setting===undefined?'':`showMissing="${setting}" `}`);
    parts[path]=new TextEncoder().encode(xml);
    const after=new Workbook(readXlsx(zip(parts)).data), expected=setting!=='false'&&setting!=='0';
    assert.equal(!!after.sheets[0].slicers[0].showDeleted,expected,String(setting));
    const reread=readXlsx(writeXlsx(after)).data.sheets[0].slicers[0];
    assert.equal(!!reread.showDeleted,expected,String(setting));
  }
});
