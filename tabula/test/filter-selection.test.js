import {test} from 'node:test';
import assert from 'node:assert/strict';
import {filterSearchPredicate,filterSelection} from '../src/filter-selection.js';

test('필터 적용 가능 상태는 유효한 표시 항목 체크가 있을 때만 참이다',()=>{
  const s=filterSelection(['서울','부산'],['서울']);assert.equal(s.canApply(),true);
  s.search('없는값');assert.equal(s.canApply(),false);assert.deepEqual(s.result(true),['서울']);
  s.search('부산');assert.equal(s.canApply(),true);s.selectVisible(false);assert.equal(s.canApply(),false);assert.deepEqual(s.result(true),['서울']);
  s.search('');assert.equal(s.canApply(),true);assert.deepEqual(s.result(),['서울']);s.selectVisible(false);assert.equal(s.canApply(),false);
});
test('더 이상 없는 이전 필터 항목과 미등록 값은 선택에 다시 나타나지 않는다',()=>{
  const old=['사라진 값','부산'],s=filterSelection(['서울','부산','서울'],old);
  assert.deepEqual(s.result(),['부산']);assert.deepEqual(old,['사라진 값','부산']);
  s.toggle('사라진 값',true);assert.deepEqual(s.result(),['부산']);
  const empty=filterSelection(['서울'],['사라진 값']);assert.equal(empty.canApply(),false);assert.deepEqual(empty.result(),[]);
});
test('검색 중 화면 밖 값의 늦은 change는 임시 선택을 오염하지 않는다',()=>{
  const s=filterSelection(['서울','부산','대구'],['서울']);s.search('부');s.toggle('대구',true);s.toggle('서울',false);
  assert.deepEqual(s.result(),['부산']);assert.deepEqual(s.result(true),['서울','부산']);s.search('');assert.deepEqual(s.result(),['서울']);
});
test('현재 선택 추가는 검색 직전 초안과 검색 결과를 중복 없이 합친다',()=>{
  const s=filterSelection(['서울','서울2','부산','대구'],['서울']);s.toggle('대구',true);s.search('부산');assert.deepEqual(s.result(true),['서울','대구','부산']);
  s.search('서울');s.toggle('서울2',false);assert.deepEqual(s.result(true),['서울','대구']);s.search('');assert.deepEqual(s.result(),['서울','대구']);
});
test('검색을 좁힐 때 해제한 체크는 유지하고 지우면 원래 초안을 복원한다',()=>{
  const s=filterSelection(['a1','a2','b1'],['b1']);s.search('a');s.toggle('a2',false);s.search('a2');assert.deepEqual(s.visible(),['a2']);assert.deepEqual(s.result(),[]);assert.equal(s.canApply(),false);
  s.search('');assert.deepEqual(s.result(),['b1']);s.search('a2');assert.deepEqual(s.result(),['a2']);
});
test('Excel 와일드카드 *와 ?는 기존 부분검색에 문자 수 의미를 더한다',()=>{
  const s=filterSelection(['Accounts Manager','Sales Manager','Manager','aXb','ab','aXXb']);
  s.search('manager');assert.deepEqual(s.visible(),['Accounts Manager','Sales Manager','Manager']);
  s.search('a?b');assert.deepEqual(s.visible(),['aXb']);s.search('a*b');assert.deepEqual(s.visible(),['aXb','ab','aXXb']);
  assert.equal(filterSearchPredicate('?')(''),false);assert.equal(filterSearchPredicate('*')(''),true);
});
test('물결표는 별표·물음표·물결표만 문자로 이스케이프한다',()=>{
  const s=filterSelection(['A*B','A?B','A~B','A~xB','AXB','[x].+$']);
  s.search('~*');assert.deepEqual(s.visible(),['A*B']);s.search('~?');assert.deepEqual(s.visible(),['A?B']);
  s.search('~~');assert.deepEqual(s.visible(),['A~B','A~xB']);s.search('~x');assert.deepEqual(s.visible(),['A~xB']);
  s.search('[x].+$');assert.deepEqual(s.visible(),['[x].+$']);
});
test('원본 숫자값과 표시 레이블은 각각 검색하며 둘 사이를 횡단하지 않는다',()=>{
  const s=filterSelection(['44927','x'],null,v=>v==='44927'?'2023년 1월 1일':'표시 이름');s.search('2023*1월');assert.deepEqual(s.visible(),['44927']);s.search('449');assert.deepEqual(s.visible(),['44927']);
  assert.equal(filterSearchPredicate('a?b')('a','b'),false);assert.equal(filterSearchPredicate('a*b')('a','b'),false);
  assert.equal(filterSearchPredicate('표시')('x','표시 이름'),true);assert.equal(filterSearchPredicate('MANAGER')('Account Manager'),true);
});
test('공백·중복·빈항목과 결과 배열의 외부 변형은 모델을 바꾸지 않는다',()=>{
  const items=['','서울','서울'],s=filterSelection(items);assert.deepEqual(s.result(),['','서울']);s.search('   ');assert.equal(s.searching(),false);const result=s.result();result.length=0;assert.equal(s.canApply(),true);assert.deepEqual(items,['','서울','서울']);
  assert.equal(filterSelection([]).canApply(),false);
});
test('5만개 항목의 와일드카드 검색은 모든 값을 유지하고 결과만 선택한다',()=>{
  const values=Array.from({length:50000},(_,i)=>`항목-${String(i).padStart(5,'0')}`),s=filterSelection(values,['항목-00001']);
  s.search('항목-49???');assert.equal(s.visible().length,1000);s.selectVisible(false);assert.equal(s.canApply(),false);s.toggle('항목-49999',true);assert.deepEqual(s.result(true),['항목-00001','항목-49999']);s.search('');assert.deepEqual(s.result(),['항목-00001']);
  // Many wildcard/literal alternations terminate without regex exponential backtracking.
  assert.equal(filterSearchPredicate('*a'.repeat(60)+'b')('a'.repeat(300)),false);
});
