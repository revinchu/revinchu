import test from 'node:test';
import assert from 'node:assert/strict';
import { Workbook } from '../src/workbook.js';
import { rangesCover, cellInEditRange, protectedRangeKey, noteVisible, setNoteVisibility, shiftNoteVisibility } from '../src/review-state.js';
test('편집 허용은 인증한 범위와 현재 암호 정의에만 적용된다',()=>{
 const a={name:'입력',ranges:[{r1:1,c1:1,r2:2,c2:3}]}, b={name:'암호',hash:'ABCD',ranges:[{r1:5,c1:0,r2:8,c2:2}]},grants=new Set();
 assert.equal(cellInEditRange([a,b],grants,1,2),true);assert.equal(cellInEditRange([a,b],grants,5,1),false);
 grants.add(protectedRangeKey(b));assert.equal(cellInEditRange([a,b],grants,5,1),true);assert.equal(cellInEditRange([a,{...b,hash:'ABCE'}],grants,5,1),false);assert.equal(cellInEditRange([a,b],grants,0,0),false);
});
test('보호 범위 합집합은 전체 행·열도 셀 순회 없이 빈틈을 검사한다',()=>{
 const whole={r1:0,c1:0,r2:1048575,c2:16383};assert.equal(rangesCover(whole,[{...whole,c2:10},{...whole,c1:11}]),true);assert.equal(rangesCover(whole,[{...whole,c2:10},{...whole,c1:12}]),false);
 assert.equal(rangesCover({r1:1,c1:1,r2:2,c2:2},[{r1:0,c1:0,r2:10,c2:10}]),true);
});
test('메모 모두 표시의 개별 숨김과 행/열 이동·삭제를 보존한다',()=>{
 let s=setNoteVisibility({all:true},2,3,false);assert.equal(noteVisible(s,2,3),false);assert.equal(noteVisible(s,3,3),true);
 s=shiftNoteVisibility(s,'row',1,2);assert.equal(noteVisible(s,4,3),false);s=shiftNoteVisibility(s,'col',3,-1);assert.deepEqual(s.states,{});
});
test('보호 범위와 메모 상태는 저장·구조 이동·Undo에서 유지된다',()=>{
 const w=new Workbook({sheets:[{name:'검사',cells:{'2,2':{raw:'입력',comment:'메모'}},protectedRanges:[{name:'입력',ranges:[{r1:2,c1:2,r2:3,c2:3}]}],noteVisibility:{states:{'2,2':true}}}]});
 w.transact(()=>w.insertRows(0,1,2));assert.equal(w.sheets[0].protectedRanges[0].ranges[0].r1,4);assert.equal(noteVisible(w.sheets[0].noteVisibility,4,2),true);
 const reread=new Workbook(w.serialize());assert.equal(reread.sheets[0].protectedRanges[0].name,'입력');w.undo();assert.equal(w.sheets[0].protectedRanges[0].ranges[0].r1,2);
});
