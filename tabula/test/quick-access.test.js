import { test } from 'node:test';
import assert from 'node:assert/strict';
import { DEFAULT_QAT_ORDER, DEFAULT_QAT_POSITION, normalizeQatOptions } from '../src/quick-access.js';

const expected = {
  qatOrder: ['painter', 'mergeCenter', 'alignCenter', 'incDecimal', 'autosum', 'calcField', 'toggleGrid', 'condColorScale', 'condDataBar', 'refreshAll', 'textToColumns', 'replace'],
  qatPosition: 'below',
};

test('빠른 실행 기본: 요청한 12개 명령의 순서와 리본 아래 위치', () => {
  assert.deepEqual(DEFAULT_QAT_ORDER, expected.qatOrder);
  assert.equal(DEFAULT_QAT_POSITION, 'below');
  for (const input of [undefined, null, {}, { qat: [], qatOrder: null, qatPosition: 'above' }]) {
    assert.deepEqual(normalizeQatOptions(input), expected);
  }
});

test('구 기본 3개는 이전하되 순서나 위치를 바꾼 사용자 설정은 보존', () => {
  for (const qatPosition of [undefined, 'above']) {
    assert.deepEqual(normalizeQatOptions({ qatOrder: ['save', 'undo', 'redo'], qatPosition }), expected);
  }
  assert.deepEqual(normalizeQatOptions({ qatOrder: ['save', 'undo', 'redo'], qatPosition: 'below' }), { qatOrder: ['save', 'undo', 'redo'], qatPosition: 'below' });
  assert.deepEqual(normalizeQatOptions({ qatOrder: ['undo', 'save', 'redo'], qatPosition: 'above' }), { qatOrder: ['undo', 'save', 'redo'], qatPosition: 'above' });
});

test('사용자 순서와 명시적인 빈 도구 모음 및 새 기본의 위쪽 배치 보존', () => {
  for (const qatOrder of [[], ['save', 'futureCommand', 'painter'], [...DEFAULT_QAT_ORDER]]) {
    for (const qatPosition of ['above', 'below']) {
      assert.deepEqual(normalizeQatOptions({ qatOrder, qatPosition, qat: ['ignoredLegacy'] }), { qatOrder, qatPosition });
    }
    assert.deepEqual(normalizeQatOptions({ qatOrder }), { qatOrder, qatPosition: 'above' }, '구 사용자 설정의 미지정 위치는 위쪽');
  }
});

test('구 qat 추가 명령은 기본 3개 뒤에 보존하고 원래 순서대로 중복 제거', () => {
  const legacy = ['painter', 'save', 'replace', 'painter'];
  const order = ['save', 'undo', 'redo', 'painter', 'replace'];
  for (const qatPosition of ['above', 'below']) {
    assert.deepEqual(normalizeQatOptions({ qat: legacy, qatOrder: null, qatPosition }), { qatOrder: order, qatPosition });
  }
  assert.deepEqual(normalizeQatOptions({ qat: legacy }), { qatOrder: order, qatPosition: 'above' });
});

test('잘못된 옵션과 명령 배열은 안전한 새 기본으로 복구', () => {
  for (const input of [true, 7, 'text', [], { qatOrder: 'save' }, { qatOrder: {} }, { qatOrder: [null] }, { qatOrder: Array(3) }, { qatOrder: ['save', 1] }, { qatOrder: [''] }, { qatOrder: [' save'] }, { qatOrder: ['save '] }, { qat: ['save', false] }, { qat: {} }]) {
    assert.deepEqual(normalizeQatOptions(input), expected);
  }
  assert.deepEqual(normalizeQatOptions({ qatOrder: ['painter'], qatPosition: 'sideways' }), { qatOrder: ['painter'], qatPosition: 'above' });
});

test('정규화는 입력을 수정하지 않고 배열 참조를 공유하지 않음', () => {
  const input = Object.freeze({ qatOrder: Object.freeze(['painter', 'save']), qatPosition: 'below' });
  const normalized = normalizeQatOptions(input);
  normalized.qatOrder.push('replace');
  assert.deepEqual(input.qatOrder, ['painter', 'save']);
  const legacy = Object.freeze({ qat: Object.freeze(['painter']) });
  normalizeQatOptions(legacy).qatOrder.pop();
  assert.deepEqual(legacy.qat, ['painter']);
  const first = normalizeQatOptions();
  first.qatOrder[0] = 'changed';
  assert.deepEqual(normalizeQatOptions(), expected);
  assert.deepEqual(DEFAULT_QAT_ORDER, expected.qatOrder);
  assert.ok(Object.isFrozen(DEFAULT_QAT_ORDER));
});

test('저장 후 다시 읽어도 기본 이전과 사용자 설정의 결과가 바뀌지 않음', () => {
  for (const input of [{}, { qatOrder: [] }, { qat: ['painter'] }, { qat: ['save', 'undo', 'save'] }, { qatOrder: [...DEFAULT_QAT_ORDER], qatPosition: 'above' }]) {
    const once = normalizeQatOptions(input);
    assert.deepEqual(normalizeQatOptions(JSON.parse(JSON.stringify(once))), once);
  }
});
