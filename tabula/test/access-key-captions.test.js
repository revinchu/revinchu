import { test } from 'node:test';
import assert from 'node:assert/strict';
import { accessKeyCaption, accessKeyDisplayLabel, allocateAccessKeys } from '../src/access-keys.js';
test('상시 접근키: 실제 키 표시와 기존 표기 중복 방지', () => {
  assert.equal(accessKeyCaption('다음', 'n'), '(N)');
  assert.equal(accessKeyCaption('다음(N)', 'n'), '');
  assert.equal(accessKeyCaption('다음(&N)', 'N'), '');
  assert.equal(accessKeyCaption('&Next', 'n'), '');
  assert.equal(accessKeyCaption('항목', '3'), '(3)');
  assert.equal(accessKeyCaption('항목', 'none'), '');
});
test('상시 접근키: 새 옵션이 추가되어도 기존 자동 배정 유지', () => {
  const before = allocateAccessKeys([{ label: '합성 하나', previous: 'p' }, { label: '합성 둘', previous: 'f' }]);
  const after = allocateAccessKeys([{ label: '합성 하나', previous: before[0].key }, { label: '합성 둘', previous: before[1].key }, { label: '암호' }, { label: '글꼴' }]);
  assert.deepEqual(after.slice(0, 2), before);
  assert.ok(!['p', 'f'].includes(after[2].key));
  assert.ok(!['p', 'f'].includes(after[3].key));
});
test('상시 접근키: 새 명시키는 기존 자동 키보다 우선하고 출처를 유지', () => {
  const result = allocateAccessKeys([{ label: '합성', previous: 'n' }, { label: '다음', explicit: 'n' }, { label: '암호', previous: 'p' }]);
  assert.notEqual(result[0].key, 'n');
  assert.deepEqual(result[1], { key: 'n', automatic: false, aliases: [] });
  assert.deepEqual(result[2], { key: 'p', automatic: false, aliases: [] });
});

test('접근키 충돌 표기: 괄호 키는 실제 키로 바꾸고 영문 단어는 보존한다',()=>{
 assert.equal(accessKeyDisplayLabel('내용 지우기(C)','a'),'내용 지우기(A)');
 assert.equal(accessKeyDisplayLabel('다음(&N)','n'),'다음(&N)');
 assert.equal(accessKeyDisplayLabel('&Close','a'),'Close');
 assert.equal(accessKeyDisplayLabel('추가(C)',''),'추가');
 assert.equal(accessKeyDisplayLabel('A && B','b'),'A && B');
});
