import { test } from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { checkValidation, listItems, validationList, validationAt, invalidCells, subtractRange, relocateValidation, validationSignature, validateValidationRule, VALIDATION_IME_MODES } from '../src/validation.js';
import { readXlsx, writeXlsx } from '../src/xlsx.js';
import { zip, unzip, textOf } from '../src/zip.js';
import { ColBuilder } from '../src/block.js';
const rule=(type,extra={})=>({r1:0,c1:0,r2:9,c2:0,type,allowBlank:false,...extra});
const round=async wb=>{const copy=new Workbook();copy.restore(readXlsx(writeXlsx(wb)).data);return copy;};

test('8 numeric operators reject invalid formula results and text numbers',()=>{
 const w=new Workbook(),cases={between:[5,true],notBetween:[15,true],equal:[5,true],notEqual:[6,true],greaterThan:[6,true],lessThan:[4,true],greaterThanOrEqual:[5,true],lessThanOrEqual:[5,true]};
 for(const [op,[v,ok]] of Object.entries(cases))assert.equal(checkValidation(w,0,rule('whole',{op,f1:'5',f2:'10'}),0,0,'='+v),ok);
 const r=rule('whole',{f1:'1',f2:'10'});
 for(const input of ['=20','=1/0',"'5",'5.1','abc'])assert.equal(checkValidation(w,0,r,0,0,input),false,input);
 assert.equal(checkValidation(w,0,r,0,0,'=5'),true);
 assert.equal(checkValidation(w,0,{...r,allowBlank:true},0,0,''),true);
 assert.equal(checkValidation(w,0,r,0,0,''),false);
});
test('custom candidate reaches dependent formulas and ROW without changing cells/cache/undo',()=>{
 const w=new Workbook();w.setInput(0,2,0,'2');w.setInput(0,2,1,'=A3*2');w.setInput(0,0,4,'=UNSUPPORTED(2)');
 const file=w.getCell(0,0,4);file.cached=41;file.dirty=false;assert.equal(w.getValue(0,2,1),4);
 const before=JSON.stringify(w.serialize()),version=w.version,cache=w.caches[0],r=rule('custom',{r1:2,r2:2,f1:'=AND(B3>10,ROW()=3)'});
 assert.equal(checkValidation(w,0,r,2,0,'7'),true);assert.equal(checkValidation(w,0,r,2,0,'3'),false);
 assert.equal(JSON.stringify(w.serialize()),before);assert.equal(w.version,version);assert.equal(w.caches[0],cache);assert.equal(cache.getRC(2,1),4);assert.equal(file.dirty,false);assert.equal(file.cached,41);
});
test('custom formulas validate candidate FORMULATEXT and Unicode text lengths use UTF16',()=>{
 const w=new Workbook();assert.equal(checkValidation(w,0,rule('custom',{f1:'=ISFORMULA(A1)'}),0,0,'=5'),true);assert.equal(checkValidation(w,0,rule('custom',{f1:'=ISFORMULA(A1)'}),0,0,'5'),false);
 for(const input of ['가나','😀','="가나"'])assert.equal(checkValidation(w,0,rule('textLength',{op:'equal',f1:'2'}),0,0,input),true);
 assert.equal(checkValidation(w,0,rule('textLength',{op:'equal',f1:'2'}),0,0,'가나다'),false);
});
test('later errors and missing custom functions never silently validate a candidate',()=>{
 const w=new Workbook();w.setInput(0,0,6,'=1/0');w.setInput(0,1,6,'=MissingFunction(1)');
 for(const r of [rule('whole',{op:'greaterThan',f1:'=$G$1'}),rule('custom',{f1:'=$G$2>0'}),rule('custom',{f1:'='})])for(const text of ['1','-1','abc','=20','=1/0'])assert.equal(checkValidation(w,0,r,0,0,text),false);
});
test('blank numeric reference follows ignoreBlank without permitting nonnumbers',()=>{
 const w=new Workbook(),r=rule('whole',{op:'greaterThan',f1:'=$G$5',allowBlank:true});
 assert.equal(checkValidation(w,0,r,0,0,'-1'),true);assert.equal(checkValidation(w,0,r,0,0,'abc'),false);
 assert.equal(checkValidation(w,0,{...r,allowBlank:false},0,0,'1'),true);assert.equal(checkValidation(w,0,{...r,allowBlank:false},0,0,'-1'),false);
});
test('names, quoted sheet ranges, INDIRECT and relative list sources resolve actual values',()=>{
 const w=new Workbook();w.sheets[0].name="선택 원본";w.setInput(0,0,6,'서울');w.setInput(0,1,6,'부산');w.setInput(0,2,6,'대전');w.names=[{name:'지역',ref:"'선택 원본'!$G$1:$G$2",sheet:null}];
 for(const f1 of ['=지역',"='선택 원본'!$G$1:$G$2",'=INDIRECT("G1:G2")','=OFFSET($G$1,0,0,2,1)'])assert.deepEqual(listItems(w,0,rule('list',{f1})),['서울','부산'],f1);
 const r=rule('list',{f1:'=G1:G2'});assert.deepEqual(listItems(w,0,r,1,0),['부산','대전']);assert.equal(checkValidation(w,0,r,1,0,'대전'),true);assert.equal(checkValidation(w,0,r,1,0,'서울'),false);
});
test('typed numeric list values and formula inputs retain actual equality',()=>{
 const w=new Workbook();w.setInput(0,0,6,'0.25');w.setInput(0,1,6,'12');w.getCell(0,0,6).style={numFmt:'percent'};const r=rule('list',{f1:'=$G$1:$G$2'});
 assert.deepEqual(listItems(w,0,r),['25%','12']);for(const text of ['25%','0.25','=1/4','12'])assert.equal(checkValidation(w,0,r,0,0,text),true);assert.equal(checkValidation(w,0,r,0,0,'="X"'),false);
});
test('blank in a named list matches native IgnoreBlank while direct list stays restricted',()=>{
 const w=new Workbook();w.setInput(0,0,6,'A');w.names=[{name:'WithBlank',ref:'Sheet1!$G$1:$G$2',sheet:null}];w.sheets[0].name='Sheet1';
 assert.equal(checkValidation(w,0,rule('list',{f1:'=WithBlank',allowBlank:true}),0,0,'X'),true);
 assert.equal(checkValidation(w,0,rule('list',{f1:'=WithBlank',allowBlank:false}),0,0,'X'),false);
 assert.equal(checkValidation(w,0,rule('list',{f1:'=$G$1:$G$2',allowBlank:true}),0,0,'X'),false);
});
test('list errors, two-dimensional ranges and large finite sources are explicit',()=>{
 const w=new Workbook();for(const f1 of ['=A1:B2','=MissingList','=A1:A100001'])assert.ok(validationList(w,0,rule('list',{f1})).error);
 assert.equal(validationList(w,0,rule('list',{f1:'=A1:A100001'})).limited,true);
 assert.deepEqual(listItems(w,0,rule('list',{f1:'="Alpha,Beta"'})),[]);
 assert.deepEqual(listItems(w,0,rule('list',{f1:'"Alpha,Beta"'})),['Alpha','Beta']);
});
test('subtractRange rebases every surviving relative condition without changing source',()=>{
 const original=rule('custom',{r1:1,c1:1,r2:9,c2:3,f1:'=A1+$Z$1>0'}),before=structuredClone(original);
 const parts=subtractRange(original,{r1:3,c1:2,r2:5,c2:2});assert.equal(parts.length,4);assert.deepEqual(original,before);
 for(const part of parts)assert.equal(validationSignature(part),validationSignature(original));
 assert.equal(parts.find(p=>p.r1===6).f1,'=A6+$Z$1>0');
 assert.equal(relocateValidation(original,{r1:3,c1:2,r2:4,c2:2}).f1,'=B3+$Z$1>0');
});
test('signatures preserve signed reference differences and metadata defaults',()=>{
 const a=rule('custom',{r1:4,f1:'=A1>0'}),b={...a,f1:'=A2>0'};assert.notEqual(validationSignature(a),validationSignature(b));
 assert.equal(validationSignature(a),validationSignature({...a,r1:5,f1:'=A2>0'}));
 assert.notEqual(validationSignature(a),validationSignature({...a,imeMode:'fullHangul'}));
 assert.equal(validationSignature(rule('any')),validationSignature(rule('any',{op:'between',showError:true,showPrompt:true,showDropdown:true,imeMode:'noControl'})));
});
test('editor settings reject malformed/error/reversed/oversized inputs but accept false custom result',()=>{
 const w=new Workbook();for(const r of [rule('whole',{f1:'1',f2:'0'}),rule('time',{op:'equal',f1:'24:00'}),rule('time',{op:'equal',f1:'1'}),rule('custom',{f1:'=1/0'}),rule('custom',{f1:'=MissingFunction(1)'}),rule('list',{f1:'"'+'x'.repeat(256)+'"'}),rule('textLength',{op:'equal',f1:'1.5'}),rule('any',{promptTitle:'x'.repeat(33)})])assert.equal(typeof validateValidationRule(w,0,r),'string',JSON.stringify(r));
 for(const r of [rule('any',{error:'keep',imeMode:'fullHangul'}),rule('time',{op:'between',f1:'09:00',f2:'18:00'}),rule('custom',{f1:'=FALSE'}),rule('list',{f1:'"a,b"'})])assert.equal(validateValidationRule(w,0,r),null);
});
test('invalidCells includes imported blocks, respects sparse overrides and effective overlapping rules',()=>{
 const w=new Workbook(),col=new ColBuilder();[2,-2,3,-3,4,-4].forEach((x,i)=>col.set(i,x));w.sheets[0].blocks=[{r0:0,c0:0,n:6,cols:[col.finish(6)]}];
 w.sheets[0].validations=[rule('whole',{r2:5,op:'greaterThan',f1:'0'}),rule('any',{r1:3,r2:3})];w.setInput(0,1,0,'8');
 assert.deepEqual(invalidCells(w,0),[{r:5,c:0}]);assert.deepEqual(invalidCells(w,0,0),[]);
});
test('all IME modes and all-value rules with only metadata survive standard XLSX',async()=>{
 const w=new Workbook();w.sheets[0].validations=VALIDATION_IME_MODES.map((imeMode,i)=>rule('any',{r1:i,r2:i,imeMode,promptTitle:'제목',showPrompt:false,errorTitle:'오류',error:'보존',showError:false,allowBlank:false}));
 const copy=await round(w);assert.equal(copy.sheets[0].validations.length,VALIDATION_IME_MODES.length);
 copy.sheets[0].validations.forEach((r,i)=>{assert.equal(r.imeMode,VALIDATION_IME_MODES[i]);assert.equal(r.promptTitle,'제목');assert.equal(r.error,'보존');assert.equal(r.showError,false);assert.equal(r.showPrompt,false);assert.equal(r.allowBlank,false);});
});
test('list named/formula/literal sources stay distinct after XLSX roundtrip',async()=>{
 const w=new Workbook();w.sheets[0].name='Sheet1';w.setInput(0,0,6,'A');w.setInput(0,1,6,'B');w.names=[{name:'Choices',ref:'Sheet1!$G$1:$G$2',sheet:null}];
 const sources=['=Choices','=INDIRECT("G1:G2")','=$G$1:$G$2','"A,B"'];w.sheets[0].validations=sources.map((f1,i)=>rule('list',{r1:i,r2:i,f1}));
 const copy=await round(w);for(const r of copy.sheets[0].validations)assert.deepEqual(listItems(copy,0,r),['A','B']);assert.equal(copy.sheets[0].validations[0].f1,'=Choices');assert.equal(copy.sheets[0].validations[3].f1,'"A,B"');
});
test('multi-sqref relative formulas are rebased before splitting and remain equivalent after export',async()=>{
 const w=new Workbook();w.sheets[0].validations=[rule('custom',{r1:1,c1:1,r2:1,c2:1,f1:'=A1>0'})];const parts=unzip(writeXlsx(w));parts['xl/worksheets/sheet1.xml']=textOf(parts['xl/worksheets/sheet1.xml']).replace('sqref="B2"','sqref="B2 D5:D6"');
 const imported=new Workbook();imported.restore(readXlsx(zip(parts)).data);const rows=imported.sheets[0].validations;assert.equal(rows.length,2);assert.equal(rows[1].f1,'C4>0');assert.equal(validationSignature(rows[0]),validationSignature(rows[1]));const copy=await round(imported);assert.equal(copy.sheets[0].validations[1].f1,'C4>0');
});

test('custom rule decides independently of candidate errors when it does not reference the candidate',()=>{
 const w=new Workbook();w.setInput(0,1,6,'5');assert.equal(checkValidation(w,0,rule('custom',{f1:'=$G$2>0'}),0,0,'=1/0'),true);assert.equal(checkValidation(w,0,rule('custom',{f1:'=ISERROR(A1)'}),0,0,'=1/0'),true);assert.equal(checkValidation(w,0,rule('custom',{f1:'=A1>0'}),0,0,'=1/0'),false);
});

test('IgnoreBlank false evaluates blanks as zero/empty text or a custom condition instead of requiring nonempty input',()=>{
 const w=new Workbook();for(const r of [rule('whole',{f1:'-1',f2:'1'}),rule('decimal',{op:'equal',f1:'0'}),rule('textLength',{op:'equal',f1:'0'}),rule('custom',{f1:'=TRUE()'})])for(const v of ['',null,undefined])assert.equal(checkValidation(w,0,r,0,0,v),true);assert.equal(checkValidation(w,0,rule('custom',{f1:'=FALSE()'}),0,0,''),false);assert.equal(checkValidation(w,0,rule('list',{f1:'"A,B"'}),0,0,''),false);
});
test('invalid list scan reuses absolute source per scan rather than re-reading it for every target',()=>{
 const w=new Workbook();for(let r=0;r<100;r++){w.setInput(0,r,0,'A');w.setInput(0,r,6,r?'B':'A');}w.sheets[0].validations=[rule('list',{r2:99,f1:'=$G$1:$G$100'})];let sourceReads=0;const get=w.getCell.bind(w);w.getCell=(si,r,c)=>{if(c===6)sourceReads++;return get(si,r,c);};assert.deepEqual(invalidCells(w,0),[]);assert.ok(sourceReads<=110,sourceReads+' source reads');
});

test('empty referenced bounds are valid editor settings without treating missing text as zero',()=>{
 const w=new Workbook();assert.equal(validateValidationRule(w,0,rule('whole',{op:'greaterThan',f1:'=$G$5'})),null);assert.equal(validateValidationRule(w,0,rule('custom',{f1:'=$G$5'})),null);assert.equal(typeof validateValidationRule(w,0,rule('whole',{op:'greaterThan',f1:''})),'string');
});

test('dropdown itemValues retain formatted numeric precision and literal formula-looking text',()=>{
 const w=new Workbook();w.setInput(0,0,6,'0.25123');w.getCell(0,0,6).style={numFmt:'percent',decimals:0};w.setInput(0,1,6,'0.25456');w.getCell(0,1,6).style={numFmt:'percent',decimals:0};w.setInput(0,2,6,'TRUE');
 const list=validationList(w,0,rule('list',{f1:'=$G$1:$G$3'}));assert.deepEqual(list.items,['25%','TRUE']);assert.deepEqual(list.itemValues,[0.25123,true]);assert.equal(list.literal,false);assert.deepEqual(list.values,[0.25123,0.25456,true]);assert.equal(checkValidation(w,0,rule('list',{f1:'=$G$1:$G$3'}),0,0,'0.25456'),true);
 const literal=validationList(w,0,rule('list',{f1:'"=1+1,값"'}));assert.deepEqual(literal.itemValues,['=1+1','값']);assert.equal(literal.literal,true);assert.deepEqual(validationList(w,0,rule('list',{f1:'=MissingList'})).itemValues,[]);
});
