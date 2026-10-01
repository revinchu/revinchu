import { test } from 'node:test';
import assert from 'node:assert/strict';
import { accessKeyFromLabel, accessKeyFromEvent, accessKeyHint, accessKeyAliases, dialogButtonAccessKey, allocateAccessKeys } from '../src/access-keys.js';

test('접근키: 명시적인 괄호·앰퍼샌드 레이블 해석', () => {
  for (const [label, key] of [['모두 바꾸기(&A)', 'a'], ['닫기(D)', 'd'], ['&Open', 'o'], ['항목(1):', '1'], ['&취소', ''], ['A && B', ''], ['한글', '']]) assert.equal(accessKeyFromLabel(label), key);
});
test('접근키: 한글 물리 키와 Alt+IME229, AltGr·Ctrl 분리', () => {
  assert.equal(accessKeyFromEvent({ code: 'KeyA', key: 'ㅁ', altKey: true, isComposing: true, keyCode: 229 }), 'a');
  assert.equal(accessKeyFromEvent({ code: 'Digit3', key: '#' }), '3');
  assert.equal(accessKeyFromEvent({ key: 'D' }), 'd');
  assert.equal(accessKeyFromEvent({ key: 'a', isComposing: true }), '');
  assert.equal(accessKeyFromEvent({ code: 'KeyA', ctrlKey: true, altKey: true }), '');
  assert.equal(accessKeyFromEvent({ code: 'KeyA', getModifierState: (name) => name === 'AltGraph' }), '');
});
test('접근키: 명시 키 우선·중복 유지·자동 배정은 남은 키 사용', () => {
  const assigned = allocateAccessKeys([{ label: '기본' }, { label: '닫기(D)', explicit: 'A' }, { label: '다른 항목(A)' }, { label: '입력', previous: 'z' }]);
  assert.deepEqual(assigned, [{ key: 'b', automatic: true }, { key: 'a', automatic: false }, { key: 'a', automatic: false }, { key: 'z', automatic: true }]);
});
test('접근키: 36개 초과 컨트롤도 빈 키 없이 순환 대상으로 배정', () => {
  const assigned = allocateAccessKeys(Array.from({ length: 80 }, () => ({ label: '합성 컨트롤' })));
  assert.equal(new Set(assigned.slice(0, 36).map((x) => x.key)).size, 36);
  assert.ok(assigned.every((x) => /^[a-z0-9]$/.test(x.key) && x.automatic));
});
test('접근키: 한국어 Excel 단일 레이블 추천과 명시 키 충돌 분리', () => {
  assert.equal(accessKeyHint('글꼴...'), 'f'); assert.equal(accessKeyHint('암호:'), 'p');
  assert.equal(accessKeyHint('이름'), ''); assert.equal(accessKeyHint('닫기'), '');
  const keys = allocateAccessKeys([{ label: '글꼴' }, { explicit: 'f', label: '별도 명령' }, { label: '암호' }]);
  assert.equal(keys[0].automatic, true); assert.notEqual(keys[0].key, 'f');
  assert.deepEqual(keys[1], { key: 'f', automatic: false }); assert.deepEqual(keys[2], { key: 'p', automatic: false });
});
test('접근키: 별칭은 자동 배정과 충돌하지 않음', () => {
  assert.deepEqual(accessKeyAliases('C, d|D bad 3'), ['c', 'd', '3']);
  const keys = allocateAccessKeys([{ label: '합성 컨트롤', aliases: 'a b' }, { explicit: 'd', label: '닫기', aliases: 'c' }]);
  assert.deepEqual(keys[0], { key: 'e', automatic: true });
});
test('접근키: 공통 확인 O·취소 C·닫기 D 기본 키', () => {
  assert.equal(dialogButtonAccessKey('확인'), 'o'); assert.equal(dialogButtonAccessKey('취소'), 'c');
  assert.equal(dialogButtonAccessKey('닫기(&D)'), 'd'); assert.equal(dialogButtonAccessKey('다음'), '');
});
test('접근키: 자동 키가 36개를 넘어도 명시·별칭 예약 키를 재사용하지 않음', () => {
  const assigned = allocateAccessKeys([{ explicit: 'd', aliases: 'c' }, ...Array.from({ length: 80 }, () => ({ label: '합성 컨트롤' }))]);
  assert.ok(assigned.slice(1).every((x) => x.key !== 'd' && x.key !== 'c'));
  const allReserved = allocateAccessKeys([... 'abcdefghijklmnopqrstuvwxyz1234567890'].map((explicit) => ({ explicit })).concat({ label: '추가' }));
  assert.equal(allReserved.at(-1).key, '', '명시 키만으로 꽉 찬 경우 잘못된 명령과 겹치기보다 Tab 사용');
});
