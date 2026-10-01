import { test } from 'node:test';
import assert from 'node:assert/strict';
import { MAX_ROWS } from '../src/formula.js';
import { Workbook, cellData } from '../src/workbook.js';
import { snapshotPasteSource, pasteSourceFromText, pasteSpecialRange, applyPasteSpecial, PASTE_SPECIAL_TYPES, transposePasteFormula } from '../src/paste-special.js';
const box = (r1, c1, r2 = r1, c2 = c1) => ({ r1, c1, r2, c2 });
const copy = (w, area, si = 0) => { const rows = [], data = [], values = []; for (let r = area.r1; r <= area.r2; r++) { rows.push(r); const d = [], v = []; for (let c = area.c1; c <= area.c2; c++) { d.push(cellData(w.getCell(si, r, c))); v.push(w.getValue(si, r, c)); } data.push(d); values.push(v); } return snapshotPasteSource(w, { ...area, si, rows, data, values }); };
const paste = (w, src, target, opts = {}, si = 0) => w.transact(() => applyPasteSpecial(w, si, src, target, opts));
const book = cells => new Workbook({ sheets: [{ name: '원본 문서', cells }, { name: '대상', cells: {} }] });

test('선택하여 붙여넣기 12개 옵션과 표시 키가 중복되지 않는다', () => { assert.equal(PASTE_SPECIAL_TYPES.length, 12); assert.equal(new Set(PASTE_SPECIAL_TYPES.map(x => x[2])).size, 12); });
test('복사 시점 상속 서식/열너비 snapshot 및 단일 Undo/Redo', () => {
  const w=book({'0,0':{raw:'2'},'0,1':{raw:'=A1*3'},'4,4':{raw:'보존'}});w.sheets[0].colStyles[0]={fill:'#abc123'};w.sheets[0].colWidths[0]=123;
  const s=copy(w,box(0,0,0,1));w.sheets[0].colStyles[0].fill='#fff';w.sheets[0].colWidths[0]=234;
  paste(w,s,box(4,4));assert.equal(w.getValue(0,4,4),2);assert.equal(w.getRaw(0,4,5),'=E5*3');assert.equal(w.styleAt(0,4,4).fill,'#abc123');assert.equal(w.undoStack.length,1);w.undo();assert.equal(w.getRaw(0,4,4),'보존');w.redo();assert.equal(w.getValue(0,4,5),6);
  paste(w,s,box(6,6),{what:'colWidths'});assert.equal(w.colWidth(0,6),123);
});
test('값만은 숫자 같은 텍스트/수식 같은 문자열/참 값을 보존하고 기존 링크·메모·서식을 유지한다',()=>{
  const w=book({'0,0':{raw:'00123',style:{numFmt:'text'}},'0,1':{raw:'=1+1',inputType:'text'},'0,2':{raw:'TRUE'},'2,0':{raw:'9',style:{fill:'#112233'},comment:'원래 메모',link:'https://example.com'}});paste(w,copy(w,box(0,0,0,2)),box(2,0),{what:'values'});
  assert.equal(w.getValue(0,2,0),'00123');assert.equal(w.getValue(0,2,1),'=1+1');assert.equal(w.getValue(0,2,2),true);assert.equal(w.getCell(0,2,0).comment,'원래 메모');assert.equal(w.getCell(0,2,0).link,'https://example.com');assert.equal(w.styleAt(0,2,0).fill,'#112233');
});
test('수식만은 상대참조 이동·대상 텍스트 표시형식 위에서도 수식 유지',()=>{
 const w=book({'0,0':{raw:'=B1+$C$1'},'3,2':{raw:'기존',style:{numFmt:'text'}}});paste(w,copy(w,box(0,0)),box(3,2),{what:'formulas'});assert.equal(w.getRaw(0,3,2),'=D4+$C$1');assert.equal(w.getCell(0,3,2).formula,true);assert.equal(w.styleAt(0,3,2).numFmt,'text');
});
test('서식만은 기존 텍스트/수식/그림/캐시를 재해석하지 않는다',()=>{
 const w=book({'0,0':{raw:'4',style:{numFmt:'general'}},'2,0':{raw:'001',style:{numFmt:'text'}},'2,1':{raw:'=1+1',cached:42},'2,2':{raw:'',image:{src:'data:image/png;base64,AA=='}}});const s=copy(w,box(0,0));paste(w,s,box(2,0,2,2),{what:'formats'});assert.equal(w.getValue(0,2,0),'001');assert.equal(w.getCell(0,2,1).cached,42);assert.equal(w.getCell(0,2,2).image.src,'data:image/png;base64,AA==');
});
test('테두리 제외는 대상 대각선 색/굵기와 사방 테두리를 유지한다',()=>{
 const w=book({'0,0':{raw:'7',style:{fill:'#112233',bt:true,du:true,duc:'#ff0000',dus:'thick'}},'2,2':{raw:'8',style:{br:true,brc:'#00ff00',brs:'double',du:true,duc:'#0000ff',dus:'dashed'}}});paste(w,copy(w,box(0,0)),box(2,2),{what:'noBorders'});const st=w.styleAt(0,2,2);assert.equal(st.fill,'#112233');assert.equal(st.bt,false);assert.equal(st.br,true);assert.equal(st.brc,'#00ff00');assert.equal(st.duc,'#0000ff');assert.equal(st.dus,'dashed');
});
test('숫자서식은 기존 custom/queryFormat을 지우고 나머지 서식은 유지한다',()=>{
 for(const what of ['formulasNum','valuesNum']){const w=book({'0,0':{raw:'=1+2',style:{numFmt:'percent',decimals:2}},'3,0':{raw:'9',style:{numFmt:'custom',code:'0.00"원"',queryFormat:'0.000',bold:true}}});paste(w,copy(w,box(0,0)),box(3,0),{what});const st=w.styleAt(0,3,0);assert.equal(st.numFmt,'percent');assert.equal(st.decimals,2);assert.equal(st.code,undefined);assert.equal(st.queryFormat,undefined);assert.equal(st.bold,true);assert.equal(!!w.getCell(0,3,0).formula,what==='formulasNum');}
});
test('메모만은 메모를 복사/제거하고 나머지 셀을 유지한다',()=>{
 const w=book({'0,0':{raw:'2',comment:'새 메모'},'0,1':{raw:'3'},'2,0':{raw:'=1+1',comment:'옛 메모'},'2,1':{raw:'9',comment:'지움'}});paste(w,copy(w,box(0,0,0,1)),box(2,0),{what:'comments'});assert.equal(w.getRaw(0,2,0),'=1+1');assert.equal(w.getCell(0,2,0).comment,'새 메모');assert.equal(w.getCell(0,2,1).comment,undefined);
});
test('유효성만은 기존 규칙 대상 부분을 교체하고 상대수식·전치·반복 영역을 반영한다',()=>{
 const w=book({'0,0':{raw:'1'},'1,0':{raw:'2'},'4,4':{raw:'보존'}});w.sheets[0].validations=[{...box(0,0,1,0),type:'custom',f1:'A1>0'},{...box(4,4,6,6),type:'whole',f1:'9'}];const s=copy(w,box(0,0,1,0));paste(w,s,box(4,4,5,5),{what:'validation',transpose:true});assert.equal(w.getRaw(0,4,4),'보존');const v=w.sheets[0].validations;assert.ok(v.some(x=>x.r1===4&&x.c1===4&&x.c2===5&&x.f1==='E5>0'));assert.ok(v.some(x=>x.r1===5&&x.c1===4&&x.f1==='E6>0'));assert.ok(v.some(x=>x.c1===6||x.r1===6));
});
test('모두와 조건부 서식 병합은 기존 규칙을 교체/추가로 구분한다',()=>{
 for(const what of ['all','mergeCond']) {const w=book({'0,0':{raw:'1'},'2,0':{raw:'2'}});w.sheets[0].cond=[{...box(0,0),type:'formula',formula:'=A1>0',style:{fill:'#ff0000'}},{...box(2,0),type:'formula',formula:'=A3<0',style:{fill:'#00ff00'}}];paste(w,copy(w,box(0,0)),box(2,0),{what});const rules=w.sheets[0].cond.filter(x=>x.r1===2);assert.equal(rules.length,what==='all'?1:2);assert.ok(rules.some(x=>x.formula==='=A3>0'));}
});
test('일반 붙여넣기는 원본에 없는 대상 유효성/조건부서식도 제거한다',()=>{
 const w=book({'0,0':{raw:'1'},'2,0':{raw:'2'}});w.sheets[0].cond=[{...box(2,0,4,0),type:'formula',formula:'=A3>0'}];w.sheets[0].validations=[{...box(2,0,4,0),type:'whole',f1:'1'}];paste(w,copy(w,box(0,0)),box(2,0));assert.equal(w.sheets[0].cond[0].r1,3);assert.equal(w.sheets[0].cond[0].formula,'=A4>0');assert.equal(w.sheets[0].validations[0].r1,3);
});
test('빈 셀 건너뛰기는 식 결과 빈문자와 셀 그림을 빈 셀로 취급하지 않는다',()=>{
 const w=book({'0,0':{raw:''},'0,1':{raw:'=""'},'0,2':{raw:'',image:{src:'data:image/png;base64,AA=='}},'2,0':{raw:'첫째'},'2,1':{raw:'둘째'},'2,2':{raw:'셋째'}});paste(w,copy(w,box(0,0,0,2)),box(2,0),{skipBlanks:true});assert.equal(w.getRaw(0,2,0),'첫째');assert.equal(w.getRaw(0,2,1),'=""');assert.equal(w.getCell(0,2,2).image.src,'data:image/png;base64,AA==');
});
test('전치와 선택 영역 반복은 모든 값과 서식을 정확한 위치에 놓는다',()=>{
 const w=book({'0,0':{raw:'1'},'0,1':{raw:'2'},'1,0':{raw:'3'},'1,1':{raw:'4'}});paste(w,copy(w,box(0,0,1,1)),box(4,4,7,7),{transpose:true});assert.deepEqual([w.getValue(0,4,4),w.getValue(0,4,5),w.getValue(0,5,4),w.getValue(0,7,7)],[1,3,2,4]);
});
test('덧셈/뺄셈/곱셈/나눗셈과 0 나누기',()=>{
 for(const [op,expected] of [['add',12],['sub',8],['mul',20],['div',5]]){const w=book({'0,0':{raw:'2'},'2,0':{raw:'10'}});paste(w,copy(w,box(0,0)),box(2,0),{op});assert.equal(w.getValue(0,2,0),expected);assert.equal(w.undoStack.length,1);w.undo();assert.equal(w.getValue(0,2,0),10);}
 const w=book({'0,0':{raw:'0'},'2,0':{raw:'10'}});paste(w,copy(w,box(0,0)),box(2,0),{op:'div'});assert.equal(w.getValue(0,2,0).code,'#DIV/0!');
});
test('연산은 대상/원본 수식 상대 참조를 유지하고 값 연산은 식을 제거한다',()=>{
 const w=book({'0,0':{raw:'=B1*2'},'0,1':{raw:'3'},'2,0':{raw:'=B3'},'2,1':{raw:'4'}});paste(w,copy(w,box(0,0)),box(2,0),{op:'add'});assert.equal(w.getRaw(0,2,0),'=(B3)+(B3*2)');assert.equal(w.getValue(0,2,0),12);
});
test('연결 붙여넣기는 숨김 행 제외된 원본 주소·시트명·절대 참조를 보존',()=>{
 const w=book({'0,0':{raw:'1'},'2,0':{raw:'3'}});const s=snapshotPasteSource(w,{si:0,...box(0,0,1,0),rows:[0,2],data:[[cellData(w.getCell(0,0,0))],[cellData(w.getCell(0,2,0))]],values:[[1],[3]]});paste(w,s,box(1,1),{what:'link'},1);assert.equal(w.getRaw(1,1,1),"='원본 문서'!$A$1");assert.equal(w.getRaw(1,2,1),"='원본 문서'!$A$3");assert.equal(w.getValue(1,2,1),3);
});
test('외부 TSV의 인용/줄바꿈/전치 및 값 자료형 보존',()=>{
 const w=book({});const s=pasteSourceFromText('12\t"두\n줄"\r\nTRUE\t=1+1');paste(w,s,box(0,0),{what:'values',transpose:true});assert.equal(w.getValue(0,0,0),12);assert.equal(w.getValue(0,0,1),true);assert.equal(w.getValue(0,1,0),'두\n줄');assert.equal(w.getValue(0,1,1),'=1+1');assert.throws(()=>paste(w,s,box(5,5),{what:'link'}),/원본 시트/);
});
test('원본 테마는 복사 시점 실제 색·글꼴을 적용하며 전체 통합 문서 테마를 교체하지 않는다',()=>{
 const w=book({'0,0':{raw:'1',style:{font:'Arial',color:'#123456',fill:'#abcdef'}}});w.theme=['#fff','#000'];const theme=[...w.theme];paste(w,copy(w,box(0,0)),box(2,0),{what:'sourceTheme'});assert.equal(w.styleAt(0,2,0).font,'Arial');assert.equal(w.styleAt(0,2,0).color,'#123456');assert.deepEqual(w.theme,theme);
});
test('보호 상태 서식 적용은 기존 셀 잠금/수식 숨김 정의를 변경하지 않는다',()=>{
 const w=book({'0,0':{raw:'1',style:{locked:false}},'2,0':{raw:'2',style:{locked:true,hideFormula:true}}});w.sheets[0].protect={on:true,allow:{formatCells:true}};paste(w,copy(w,box(0,0)),box(2,0),{what:'formats'});assert.equal(w.styleAt(0,2,0).locked,true);assert.equal(w.styleAt(0,2,0).hideFormula,true);
});
test('시트 범위/최대 면적 검사에서 부분 변경 없이 거부한다',()=>{
 const w=book({'0,0':{raw:'1'},'1,0':{raw:'2'}}),s=copy(w,box(0,0,1,0)),before=JSON.stringify(w.serialize());assert.throws(()=>paste(w,s,box(MAX_ROWS-1,0)),/마지막 행/);assert.equal(JSON.stringify(w.serialize()),before);assert.throws(()=>pasteSpecialRange(s,box(0,0,1999999,1)),/200만/);
});

test('실제 Excel 대조: 값 연산에서도 대상 수식 유지, 비숫자 원본은 값 미변경',()=>{
 const w=book({'0,0':{raw:'=B1*2'},'0,1':{raw:'3'},'2,0':{raw:'=B3'},'2,1':{raw:'4'}});paste(w,copy(w,box(0,0)),box(2,0),{what:'values',op:'add'});assert.equal(w.getRaw(0,2,0),'=(B3)+6');assert.equal(w.getValue(0,2,0),10);
 const x=book({'0,0':{raw:'abc'},'2,0':{raw:'10'}});paste(x,copy(x,box(0,0)),box(2,0),{op:'add'});assert.equal(x.getValue(0,2,0),10);
});
test('실제 Excel 전치 참조 21종: 상대/절대/혼합/범위/전체 행·열',()=>{
 const src={rows:[0,1,2],c1:0,c2:2};
 const cases=[['=B1','=D5'],['=A2','=E4'],['=$B1','=D$5'],['=B$1','=$D5'],['=$B$1','=$D$5'],['=B1:C2','=D5:E6'],['=SUM(B:B)','=SUM($D5:$XFD5)'],['=SUM(2:2)','=SUM(E$4:E$1048576)'],['=B9','=L5'],['=$B$9','=$B$9'],['=D1','=D7'],['=B1:D2','=D5:E7'],['=Sheet1!B1','=Sheet1!D5'],['=SUM(A1:XFD1)','=SUM(D4:D1048576)'],['=$B9','=$B9'],['=B$9','=B$9'],['=$D1','=$D1'],['=D$1','=D$1'],['=$D$1','=$D$1'],['=B$1:D$2','=$D5:$E7'],['=$B1:$D2','=D$5:E$7']];
 for(const [formula,expected] of cases) assert.equal(transposePasteFormula(formula,src,0,0,3,3),expected,formula);
});

test('오류 값 복사는 문자열이나 일반 객체로 바뀌지 않는다',()=>{const w=book({'0,0':{raw:'=1/0'},'2,0':{raw:'7'}}),src=copy(w,box(0,0));paste(w,src,box(2,0),{what:'values'});assert.equal(w.getValue(0,2,0).code,'#DIV/0!');w.undo();paste(w,src,box(2,0),{what:'values',op:'add'});assert.equal(w.getValue(0,2,0).code,'#DIV/0!');});
