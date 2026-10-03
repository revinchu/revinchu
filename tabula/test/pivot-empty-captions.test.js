import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxAsync } from '../src/xlsx.js';
import { computePivot, pivotSourceData, resolvePivot } from '../src/pivot.js';
import { unzip, textOf, zip } from '../src/zip.js';
import { parseXml, descendants, child } from '../src/xml.js';
const path = 'xl/pivotTables/pivotTable1.xml';
const attrs = { rowCaption:'rowHeaderCaption', colCaption:'colHeaderCaption', grandCaption:'grandTotalCaption', dataCaption:'dataCaption' };
function fixture(captions = {}) {
  const wb = new Workbook();
  wb.transact(() => {
    [['지역','상품','매출','비용'],['서울','A',10,1],['부산','B',20,2]].forEach((row,r) => row.forEach((v,c) => wb.setInput(0,r,c,String(v))));
    wb.addSheet('보고서');
    wb.setSheetProp(1,'pivot',{name:'캡션',source:'Sheet1',range:{r1:0,c1:0,r2:2,c2:3},rows:['지역'],cols:['상품'],values:[{field:'매출'}],top:0,left:0,...captions});
  });
  return wb;
}
function render(wb) { const d=wb.sheets[1].pivot, r=resolvePivot(pivotSourceData(wb,d),d); return computePivot(r,r.def).grid; }
for (const async of [false,true]) test(`명시적 빈 피벗 캡션은 읽기·쓰기·다시 그리기에서도 기본 문구가 되지 않음 async=${async}`, async () => {
  const empty=Object.fromEntries(Object.keys(attrs).map(key=>[key,''])), wb=fixture(empty), before=render(wb);
  const bytes=async ? await writeXlsxAsync(wb) : writeXlsx(wb), xml=parseXml(textOf(unzip(bytes)[path]));
  for(const attr of Object.values(attrs)) assert.equal(xml.attrs[attr],'');
  const after=new Workbook(readXlsx(bytes).data);
  for(const key of Object.keys(attrs)) assert.equal(after.sheets[1].pivot[key],'');
  assert.deepEqual(render(after).map(row=>row.map(c=>c.raw)),before.map(row=>row.map(c=>c.raw)));
  for(const role of ['rowHead:0','colHead','grandLabel']) {
    const cells=render(after).flat().filter(c=>c.role===role); assert.ok(cells.length); assert.equal(cells[0].raw,"'",role); assert.ok(cells.every(c=>c.raw==="'" || c.raw===''),role);
  }
  const colHeads=render(after).flat().filter(c=>c.role==='colHead');
  assert.equal(colHeads.filter(c=>c.raw==="'").length,1,'명시적 열 캡션 한 칸만 문자열이고 나머지 머리글 여백은 빈 셀');
  assert.ok(render(after).flat().filter(c=>c.role==='corner').every(c=>c.raw===''),'모서리 여백은 빈 셀');
  const again=readXlsx(writeXlsx(after)).data.sheets[1].pivot;
  for(const key of Object.keys(attrs)) assert.equal(again[key],'');
});
test('캡션이 없을 때의 기본값은 유지하며 외부 XML의 빈 캡션을 읽음', () => {
  const bytes=writeXlsx(fixture()), files=unzip(bytes), xml=parseXml(textOf(files[path]));
  for(const attr of ['rowHeaderCaption','colHeaderCaption','grandTotalCaption']) assert.equal(xml.attrs[attr],undefined);
  const original=new Workbook(readXlsx(bytes).data), cells=render(original).flat();
  assert.equal(cells.find(c=>c.role==='rowHead:0').raw,'행 레이블');
  assert.equal(cells.find(c=>c.role==='colHead').raw,'열 레이블');
  assert.equal(cells.find(c=>c.role==='grandLabel').raw,'총합계');
  files[path]=textOf(files[path]).replace('<pivotTableDefinition ','<pivotTableDefinition rowHeaderCaption="" colHeaderCaption="" grandTotalCaption="" ');
  const imported=readXlsx(zip(files)).data.sheets[1].pivot;
  for(const key of ['rowCaption','colCaption','grandCaption']) assert.equal(imported[key],'');
});
test('빈 값 캡션은 여러 지표의 값 머리글에도 보존됨', () => {
  const wb=fixture({cols:[],values:[{field:'매출'},{field:'비용'}],showValuesRow:true,dataCaption:''});
  const before=render(wb); assert.equal(before[0][1].raw,"'");
  const after=new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(after.sheets[1].pivot.dataCaption,''); assert.equal(render(after)[0][1].raw,"'");
});

test('명시적 빈 행 캡션은 셀값 빈 문자열과 XLSX 빈 문자열 저장값으로 남음', () => {
  const wb=fixture({rowCaption:'',cols:[]}), grid=render(wb);
  const header=grid[0][0];assert.equal(header.role,'rowHead:0');assert.equal(header.raw,"'");
  wb.transact(()=>grid.forEach((row,r)=>row.forEach((cell,c)=>wb.setCellData(1,r,c,cell))));
  assert.equal(wb.getValue(1,0,0),'');assert.notEqual(wb.getValue(1,0,0),null);
  const parts=unzip(writeXlsx(wb)), root=parseXml(textOf(parts['xl/worksheets/sheet2.xml']));
  const cell=descendants(root,'c').find(c=>c.attrs.r==='A1');assert.ok(cell);
  if(cell.attrs.t==='s') {
    const index=Number(child(cell,'v')?.text), strings=descendants(parseXml(textOf(parts['xl/sharedStrings.xml'])),'si');
    assert.equal(descendants(strings[index],'t').map(t=>t.text??'').join(''),'');
  } else {
    assert.equal(cell.attrs.t,'str');assert.ok(child(cell,'v'));assert.equal(child(cell,'v').text,'');
  }
  const restored=new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(restored.getValue(1,0,0),'');assert.equal(restored.sheets[1].pivot.rowCaption,'');
  const hidden=render(fixture({rowCaption:'',showHeaders:false,cols:[]}));
  assert.equal(hidden[0][0].raw,'','머리글 숨김은 여전히 빈 셀');
});
