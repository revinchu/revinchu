import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { readXlsx, writeXlsx, writeXlsxAsync } from '../src/xlsx.js';
import { unzip, textOf } from '../src/zip.js';
import { parseXml, descendants, child } from '../src/xml.js';
const escaped='  한글 😊 & <tag> _x000A_\n끝  ';
function fixture(cached) {
  return new Workbook({sheets:[{name:'결과',cells:{
    '0,0':{raw:'=TRANSPOSE(D1:F1)',...(cached?{cached:'첫 값',cachedArray:{h:5,w:1,values:[0,0,'첫 값',2,0,escaped,4,0,'']}}:{})},
    [cached?'2,0':'1,0']:{raw:'',style:{bold:true}},
    '0,3':{raw:'첫 값'},'0,4':{raw:escaped,inputType:'text'},'0,5':{raw:"'"},'0,7':{raw:'일반 문자열'},
  }}]});
}
for(const cached of [false,true])for(const async of [false,true])test(`배열 후속 문자열은 수식 결과 t=str로 저장 (cached=${cached}, async=${async})`, async()=>{
  const w=fixture(cached), bytes=async?await writeXlsxAsync(w):writeXlsx(w), files=unzip(bytes);
  const root=parseXml(textOf(files['xl/worksheets/sheet1.xml'])), cells=new Map(descendants(root,'c').map(c=>[c.attrs.r,c]));
  assert.equal(cells.get('A1').attrs.t,'str');assert.ok(child(cells.get('A1'),'f'),'앵커 수식은 유지');
  const at=cached?'A3':'A2', empty=cached?'A5':'A3';
  for(const ref of [at,empty]){assert.equal(cells.get(ref).attrs.t,'str');assert.equal(child(cells.get(ref),'f'),null,'f는 앵커에만 기록');}
  assert.equal(child(cells.get(at),'v').text,escaped);assert.equal(child(cells.get(empty),'v').text,'');
  assert.ok(cells.get(at).attrs.s,'후속 셀의 직접 서식도 유지');
  assert.equal(cells.get('D1').attrs.t,'s');assert.equal(cells.get('H1').attrs.t,'s','일반 문자열은 기존 SST 사용');
  if(cached){assert.equal(cells.has('A2'),false);assert.equal(cells.has('A4'),false);}
  const back=new Workbook(readXlsx(bytes).data);
  assert.equal(back.getValue(0,cached?2:1,0),escaped);assert.equal(back.getValue(0,cached?4:2,0),'');
  assert.equal(back.getValue(0,0,7),'일반 문자열');assert.equal(back.styleAt(0,cached?2:1,0).bold,true);
});
