import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { computePivot, pivotSourceData, resolvePivot } from '../src/pivot.js';

function fixture(layout,extra={}) {
  const wb=new Workbook();wb.sheets[0].name='원본';
  wb.transact(()=>{
    [['키워드','지역','매출','비용'],['검색어','서울',10,3],['검색어2','부산',20,4]].forEach((row,r)=>row.forEach((v,c)=>wb.setInput(0,r,c,String(v))));
    wb.addSheet('보고서');wb.setSheetProp(1,'pivot',{name:'행 캡션',source:'원본',range:{r1:0,c1:0,r2:2,c2:3},top:0,left:0,rows:['키워드'],values:[{field:'매출',agg:'sum'}],fieldCaptions:{키워드:'필드 이름'},layout,...extra});
  });return wb;
}
function headers(wb) {const def=wb.sheets[1].pivot,res=resolvePivot(pivotSourceData(wb,def),def);return computePivot(res,res.def).grid.flat().filter(c=>c?.role?.startsWith('rowHead:'));}
for(const layout of ['tabular','outline'])for(const caption of [' 키워드','사용자 머리글',''])test(`단일 행 필드 ${layout}의 명시적 행 캡션을 공백까지 유지 ${JSON.stringify(caption)}`,()=>{
  const wb=fixture(layout,{rowCaption:caption}),expected=caption===''?"'":caption;
  assert.equal(headers(wb)[0].raw,expected);
  const back=new Workbook(readXlsx(writeXlsx(wb)).data);
  assert.equal(back.sheets[1].pivot.rowCaption,caption);assert.equal(headers(back)[0].raw,expected);
});
for(const layout of ['tabular','outline'])test(`${layout} 기본·여러 행 필드·머리글 숨김·값 머리글에는 영향 없음`,()=>{
  assert.equal(headers(fixture(layout))[0].raw,'필드 이름');
  assert.deepEqual(headers(fixture(layout,{rowCaption:'다른 캡션',rows:['키워드','지역']})).map(c=>c.raw),['필드 이름','지역']);
  assert.ok(headers(fixture(layout,{rowCaption:'',showHeaders:false})).every(c=>c.raw===''));
  const values=fixture(layout,{rowCaption:' 키워드',values:[{field:'매출'},{field:'비용'}],valuesOnRows:true,dataCaption:'지표'});
  assert.deepEqual(headers(values).map(c=>c.raw),[' 키워드','지표']);
});
