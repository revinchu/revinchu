import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { readWixelFile, writeWixelFile } from '../src/wixel-file.js';
import { unzip, zip, textOf } from '../src/zip.js';
import { parseXml, child, kids, descendants } from '../src/xml.js';
import { pivotSourceData, resolvePivot, computePivot } from '../src/pivot.js';

const cachePath='xl/pivotCache/pivotCacheDefinition1.xml';
const tablePath='xl/pivotTables/pivotTable1.xml';
const history=['과거','현재','숨김',null,'',7,true,{error:'#N/A'}];
const historyXml='<s v="과거"/><s v="현재"/><s v="숨김"/><m/><s v=""/><n v="7"/><b v="1"/><e v="#N/A"/>';
function externalFixture(options={}) {
  const wb=new Workbook();wb.sheets[0].name='원본';
  wb.transact(()=>{
    [['항목','매출','미사용'],['현재',10,'현재 값'],['숨김',20,'현재 값']].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));
    wb.addSheet('보고서');
    wb.setSheetProp(1,'pivot',{name:'보고서',source:'원본',range:{r1:0,c1:0,r2:2,c2:2},rows:['항목'],values:[{field:'매출',agg:'sum'}],filters:{항목:['현재']},top:0,left:0,saveData:options.saveData,refreshOnOpen:options.refreshOnOpen});
    wb.setSheetProp(1,'slicers',[{id:'s1',caption:'항목',source:{kind:'pivot',field:'항목',self:true},x:10,y:100,w:140,h:200,style:'SlicerStyleLight1',showDeleted:true}]);
  });
  const parts=unzip(writeXlsx(wb));
  let cf=0;
  parts[cachePath]=textOf(parts[cachePath]).replace(/<cacheField\b[^>]*>[\s\S]*?<\/cacheField>/g,xml=>{
    const f=cf++;
    if(f!==0&&f!==2)return xml;
    const items=f===0?historyXml:'<s v="미사용 과거"/><s v="현재 값"/>';
    return xml.replace(/<sharedItems\b[^>]*\/>|<sharedItems\b[^>]*>[\s\S]*?<\/sharedItems>/,`<sharedItems containsBlank="1" containsNumber="1" containsMixedTypes="1" count="${f===0?history.length:2}">${items}</sharedItems>`);
  });
  parts[tablePath]=textOf(parts[tablePath]).replace(/<pivotField\b[^>]*>[\s\S]*?<\/pivotField>/,xml=>xml.replace(/<items\b[^>]*>[\s\S]*?<\/items>/,`<items count="8">${history.map((_,i)=>`<item x="${i}"${i>1?' h="1"':''}${i===0?' n="옛 항목 이름"':''}/>`).join('')}</items>`));
  return zip(parts);
}
const fields=parts=>kids(child(parseXml(textOf(parts[cachePath])),'cacheFields'),'cacheField');
const rawItems=field=>child(field,'sharedItems').children.map(x=>x.name==='n'?Number(x.attrs.v):x.name==='b'?x.attrs.v==='1':x.name==='m'?null:x.name==='e'?{error:x.attrs.v}:x.attrs.v??'');
const resultSum=wb=>{const def=wb.sheets[1].pivot,src=pivotSourceData(wb,def),resolved=resolvePivot(src,def);return {rows:src.cube.n,grid:computePivot(resolved,resolved.def).grid.map(row=>row.map(c=>c?.raw))};};

for(const options of [{},{saveData:false},{refreshOnOpen:true}])test(`과거 캐시 항목·선택·번호·레이블은 XLSX 저장 왕복에서 보존 ${JSON.stringify(options)}`,()=>{
  const wb=new Workbook(readXlsx(externalFixture(options)).data),def=wb.sheets[1].pivot;
  assert.deepEqual(def.filters.항목,['과거','현재']);
  assert.deepEqual(wb.pivotCacheItems[def.cacheItemsId].fields[0].shared,history);
  const before=resultSum(wb),metadata=JSON.stringify(wb.pivotCacheItems);
  wb.transact(()=>wb.setSheetProp(1,'pivot',{...def,style:'PivotStyleMedium9'}));
  const bytes=writeXlsx(wb),parts=unzip(bytes),back=new Workbook(readXlsx(bytes).data);
  assert.deepEqual(rawItems(fields(parts)[0]),history,'레코드에 없는 값과 원래 캐시 번호');
  assert.deepEqual(rawItems(fields(parts)[2]),['미사용 과거','현재 값'],'현재 피벗에서 사용하지 않는 필드도 보존');
  assert.deepEqual(back.sheets[1].pivot.filters.항목,['과거','현재']);
  assert.equal(back.sheets[1].pivot.itemCaptions.항목.과거,'옛 항목 이름');
  const shared=child(fields(parts)[0],'sharedItems');
  assert.equal(shared.attrs.containsBlank,'1');assert.equal(shared.attrs.containsNumber,'1');assert.equal(shared.attrs.minValue,'7');
  const slicer=descendants(parseXml(textOf(parts['xl/slicerCaches/slicerCache1.xml'])),'i');
  assert.equal(slicer.length,8);assert.deepEqual(slicer.filter(i=>i.attrs.s==='1').map(i=>i.attrs.x),['0','1']);
  assert.equal(slicer[0].attrs.nd,'1','과거 항목은 선택되었어도 데이터 없음으로 기록');assert.equal(slicer[1].attrs.nd,undefined);
  assert.deepEqual(resultSum(back),before,'과거 항목은 실제 레코드나 집계에 삽입하지 않음');
  assert.equal(JSON.stringify(wb.pivotCacheItems),metadata,'내보내기는 공유 원본 캐시를 수정하지 않음');
});

test('원본 수정으로 저장 레코드가 무효화되어도 과거 캐시 번호는 유지하고 새 항목만 추가',()=>{
  const wb=new Workbook(readXlsx(externalFixture()).data);
  wb.transact(()=>{wb.setInput(0,2,0,'신규');wb.setInput(0,2,1,'35');});
  const parts=unzip(writeXlsx(wb)),actual=rawItems(fields(parts)[0]);
  assert.deepEqual(actual,[...history,'신규']);
  const back=new Workbook(readXlsx(zip(parts)).data),src=pivotSourceData(back,back.sheets[1].pivot);
  assert.equal(src.cube.n,2);assert.equal(src.cube.col(0).get(1),'신규');assert.equal(src.cube.col(1).get(1),35);
  assert.deepEqual(back.sheets[1].pivot.filters.항목,['과거','현재']);
});

test('공유 캐시 메타데이터는 피벗마다 복제하지 않고 JSON·Blob·WIXEL에 한 번 저장',async()=>{
  const wb=new Workbook(readXlsx(externalFixture()).data),def=wb.sheets[1].pivot;
  wb.sheets[1].pivotsExtra=[{...structuredClone(def),name:'추가 보고서',top:20}];
  assert.equal(Object.keys(wb.pivotCacheItems).length,1);
  assert.equal(wb.bookMeta(false,false).pivotCacheItems,undefined);
  assert.equal(wb.bookMeta(false).pivotCacheItems,wb.pivotCacheItems);
  const variants=[new Workbook(wb.serialize()),new Workbook(JSON.parse(await wb.serializeBlob().text())),new Workbook((await readWixelFile(await writeWixelFile(wb))).workbook)];
  for(const restored of variants){
    assert.deepEqual(restored.pivotCacheItems,wb.pivotCacheItems);
    const parts=unzip(writeXlsx(restored));
    assert.equal(Object.keys(parts).filter(p=>/^xl\/pivotCache\/pivotCacheDefinition\d+\.xml$/.test(p)).length,1);
    assert.deepEqual(rawItems(fields(parts)[0]),history);
    const back=readXlsx(zip(parts)).data;
    assert.equal(back.sheets[1].pivot.cacheItemsId,back.sheets[1].pivotsExtra[0].cacheItemsId);
  }
});

test('같은 원본 범위를 가리키는 독립 캐시는 각자의 과거 항목과 번호를 보존',()=>{
  const wb=new Workbook(readXlsx(externalFixture()).data),def=wb.sheets[1].pivot;
  wb.pivotCacheItems={...wb.pivotCacheItems,independent:{fields:[{name:'항목',shared:['다른 과거','숨김','현재']}]}};
  wb.sheets[1].pivotsExtra=[{...structuredClone(def),name:'독립 보고서',top:20,cacheItemsId:'independent',filters:{항목:['다른 과거','현재']}}];
  const parts=unzip(writeXlsx(wb)),other={...parts,[cachePath]:parts['xl/pivotCache/pivotCacheDefinition2.xml']};
  assert.deepEqual(rawItems(fields(parts)[0]),history);
  assert.deepEqual(rawItems(fields(other)[0]),['다른 과거','숨김','현재']);
  const back=readXlsx(zip(parts)).data;
  assert.notEqual(back.sheets[1].pivot.cacheItemsId,back.sheets[1].pivotsExtra[0].cacheItemsId);
  assert.deepEqual(back.sheets[1].pivotsExtra[0].filters.항목,['현재','다른 과거']);
});
