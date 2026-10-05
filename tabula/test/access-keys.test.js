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
test('접근키: 명시 키 우선·중복 재배정·자동 배정은 남은 키 사용', () => {
  const assigned = allocateAccessKeys([{ label: '기본' }, { label: '닫기(D)', explicit: 'A' }, { label: '다른 항목(A)' }, { label: '입력', previous: 'z' }]);
  assert.deepEqual(assigned, [{ key: 'b', automatic: true, aliases: [] }, { key: 'a', automatic: false, aliases: [] }, { key: 'c', automatic: true, aliases: [] }, { key: 'z', automatic: true, aliases: [] }]);
});
test('접근키: 36개 초과 컨트롤은 모호한 재사용 대신 Tab 탐색', () => {
  const assigned = allocateAccessKeys(Array.from({ length: 80 }, () => ({ label: '합성 컨트롤' })));
  assert.equal(new Set(assigned.slice(0, 36).map((x) => x.key)).size, 36);
  assert.ok(assigned.slice(0,36).every((x) => /^[a-z0-9]$/.test(x.key) && x.automatic));
  assert.ok(assigned.slice(36).every((x) => x.key === ''));
});
test('접근키: 한국어 Excel 단일 레이블 추천과 명시 키 충돌 분리', () => {
  assert.equal(accessKeyHint('글꼴...'), 'f'); assert.equal(accessKeyHint('암호:'), 'p');
  assert.equal(accessKeyHint('이름'), ''); assert.equal(accessKeyHint('닫기'), '');
  const keys = allocateAccessKeys([{ label: '글꼴' }, { explicit: 'f', label: '별도 명령' }, { label: '암호' }]);
  assert.equal(keys[0].automatic, true); assert.notEqual(keys[0].key, 'f');
  assert.deepEqual(keys[1], { key: 'f', automatic: false, aliases: [] }); assert.deepEqual(keys[2], { key: 'p', automatic: false, aliases: [] });
});
test('접근키: 별칭은 자동 배정과 충돌하지 않음', () => {
  assert.deepEqual(accessKeyAliases('C, d|D bad 3'), ['c', 'd', '3']);
  const keys = allocateAccessKeys([{ label: '합성 컨트롤', aliases: 'a b' }, { explicit: 'd', label: '닫기', aliases: 'c' }]);
  assert.deepEqual(keys[0], { key: 'e', automatic: true, aliases: ['a','b'] });
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


test('접근키: 명시 키가 다른 항목 별칭보다 우선하고 별칭도 단일 소유', () => {
  const keys=allocateAccessKeys([{explicit:'d',aliases:'c x',label:'닫기'},{explicit:'c',aliases:'x y',label:'복사'},{explicit:'c',aliases:'y z',label:'내용 지우기'}]);
  assert.deepEqual(keys,[{key:'d',automatic:false,aliases:['x']},{key:'c',automatic:false,aliases:['y']},{key:'a',automatic:true,aliases:['z']}]);
  const all=keys.flatMap(k=>[k.key,...k.aliases]).filter(Boolean);assert.equal(new Set(all).size,all.length);
});
test('접근키: 공통 취소 기본값은 자동 배정을 막고 창별 명시 메모 C는 우선', () => {
  const normal=allocateAccessKeys([{label:'합성 선택'},{label:'취소',explicit:'c',default:true}]);assert.equal(normal[1].key,'c');assert.notEqual(normal[0].key,'c');
  const custom=allocateAccessKeys([{label:'취소',explicit:'c',default:true},{label:'메모(C)'}]);assert.equal(custom[1].key,'c');assert.notEqual(custom[0].key,'c');
});
test('접근키: 숨김 예약을 넘기지 않고 활성/비활성 중복 명시도 한 소유자', () => {
  const result=allocateAccessKeys([{label:'활성',explicit:'a'},{label:'비활성',explicit:'a'},{label:'기타'}]);assert.equal(result[0].key,'a');assert.equal(new Set(result.map(x=>x.key)).size,3);
});
