import test from 'node:test';
import assert from 'node:assert/strict';
import { Axis } from '../src/axis.js';
import { cellTextOverflow } from '../src/cell-text-overflow.js';

// 네이티브 Excel의 합성 PDF에서 확인한 차단 경계: 숨긴 셀의 값은 건너뛰고
// 다음 표시 셀의 값에서 멈춘다. 가운데 정렬은 양쪽 경계를 독립적으로 쓴다.
function probe(axis, column, align, textWidth, occupied = []) {
  const values = new Set(occupied), calls = [];
  const result = cellTextOverflow(axis, column, align, textWidth, c => {
    calls.push(c);
    assert.ok(c >= 0 && c < axis.max, '시트 밖의 셀을 읽지 않는다');
    assert.ok(axis.size(c) > 0, '숨긴 셀의 내용은 조회하지 않는다');
    return values.has(c);
  });
  return { result, calls };
}

test('숨긴 빈 열 뒤의 첫 표시 값이 왼쪽 텍스트를 차단한다', () => {
  const axis = new Axis(80, {}, [{ 2:true }], 8);
  assert.deepEqual(probe(axis, 1, 'left', 500, [3]), { result:{ left:0, right:0 }, calls:[3] });
});

test('숨긴 열의 값은 차단하지 않고 다음 표시 빈 열만 허용한다', () => {
  const axis = new Axis(80, {}, [{ 2:true }], 8);
  assert.deepEqual(probe(axis, 1, 'left', 500, [2, 4]), { result:{ left:0, right:80 }, calls:[3, 4] });
});

test('첫 표시 차단자 뒤의 빈 칸은 넘침 영역에 포함하지 않는다', () => {
  const axis = new Axis(70, { 2:23, 3:105, 4:210 }, [], 8);
  assert.deepEqual(probe(axis, 1, 'left', 400, [3]), { result:{ left:0, right:23 }, calls:[2, 3] });
});

test('오른쪽 정렬도 숨긴 열을 건너뛰고 왼쪽 표시 값에서 멈춘다', () => {
  const axis = new Axis(80, { 1:35 }, [{ 3:true }], 8);
  assert.deepEqual(probe(axis, 4, 'right', 300, [3, 0]), { result:{ left:115, right:0 }, calls:[2, 1, 0] });
});

test('가운데 정렬은 오른쪽이 막혀도 왼쪽의 빈 폭을 유지한다', () => {
  const axis = new Axis(80, { 1:50, 2:30 }, [], 8);
  assert.deepEqual(probe(axis, 3, 'center', 180, [4]), { result:{ left:80, right:0 }, calls:[2, 1, 4] });
});

test('가운데 정렬은 왼쪽이 막혀도 오른쪽의 빈 폭을 유지한다', () => {
  const axis = new Axis(80, { 4:17, 5:39, 6:100 }, [], 8);
  assert.deepEqual(probe(axis, 3, 'center', 180, [2]), { result:{ left:0, right:56 }, calls:[2, 4, 5] });
});

test('가운데 양쪽의 숨김과 서로 다른 열 너비를 각각 적용한다', () => {
  const axis = new Axis(80, { 0:20, 1:21, 5:35, 7:49 }, [{ 2:true, 4:true, 6:true }], 9);
  assert.deepEqual(probe(axis, 3, 'center', 230, [0, 8]), { result:{ left:21, right:84 }, calls:[1, 0, 5, 7] });
});

test('긴 텍스트의 실제 도달 폭까지만 조회하고 행 끝까지 탐색하지 않는다', () => {
  const axis = new Axis(80, { 11:27, 12:51, 13:13 }, [], 1_000_000);
  const { result, calls } = probe(axis, 10, 'left', 150);
  assert.deepEqual(result, { left:0, right:78 });
  assert.deepEqual(calls, [11, 12]);
});

test('짧은 텍스트의 기존 채우기 층을 유지하면서 방향당 한 이웃만 조회한다', () => {
  const axis = new Axis(80, {}, [], 1_000_000);
  for (const align of ['left', 'right', 'center']) {
    for (const width of [0, 5, 79, 80]) {
      const { result, calls } = probe(axis, 10, align, width);
      assert.deepEqual(calls, align === 'left' ? [11] : align === 'right' ? [9] : [9, 11]);
      assert.ok(Number.isFinite(result.left) && Number.isFinite(result.right));
    }
  }
});

test('시트 양끝과 빈 축에서는 경계 밖을 읽지 않는다', () => {
  const axis = new Axis(80, {}, [], 4);
  assert.deepEqual(probe(axis, 0, 'right', 900), { result:{ left:0, right:0 }, calls:[] });
  assert.deepEqual(probe(axis, 3, 'left', 900), { result:{ left:0, right:0 }, calls:[] });
  assert.deepEqual(probe(axis, 0, 'center', 900), { result:{ left:0, right:240 }, calls:[1, 2, 3] });
  assert.deepEqual(probe(new Axis(0, {}, [], 0), 0, 'center', 900), { result:{ left:0, right:0 }, calls:[] });
});

test('남은 모든 열이 숨겨졌거나 기본 너비가 0이면 내용을 조회하지 않는다', () => {
  assert.deepEqual(probe(new Axis(80, {}, [{ 0:true, 1:true, 3:true, 4:true }], 5), 2, 'center', 900), { result:{ left:0, right:0 }, calls:[] });
  assert.deepEqual(probe(new Axis(0, { 2:80 }, [], 1_000_000), 2, 'center', 900), { result:{ left:0, right:0 }, calls:[] });
});

test('기본 너비 0의 희소 표시 열은 인덱스 거리에 관계없이 폭으로 합산한다', () => {
  const axis = new Axis(0, { 5:23, 100:80, 900_000:37, 999_999:110 }, [], 1_000_000);
  assert.deepEqual(probe(axis, 100, 'center', 180), { result:{ left:23, right:147 }, calls:[5, 900_000, 999_999] });
});

test('백만 개의 연속 숨긴 열을 건너뛰어 표시 셀만 두 번 조회한다', () => {
  const count = 1_000_000, bitmap = new Uint8Array(count).fill(1);
  const hidden = { __bits:bitmap, start:1, count }, sizes = Object.freeze({ [count + 1]:31 });
  const axis = new Axis(80, sizes, [hidden], count + 4);
  let sizeCalls = 0;
  const originalSize = axis.size.bind(axis);
  axis.size = c => { sizeCalls++; return originalSize(c); };
  const { result, calls } = probe(axis, 0, 'left', 400, [count + 2]);
  assert.deepEqual(result, { left:0, right:31 });
  assert.deepEqual(calls, [count + 1, count + 2]);
  assert.ok(sizeCalls < 40, `숨긴 인덱스를 선형 조회하지 않는다: ${sizeCalls}`);
  assert.equal(hidden.start, 1);
  assert.equal(hidden.count, count);
  assert.equal(bitmap[0], 1);
  assert.equal(bitmap[count - 1], 1);
  assert.equal(sizes[count + 1], 31);
});

test('축, 숨김 원본과 이웃의 값을 수정하지 않는다', () => {
  const sizes = Object.freeze({ 0:15, 3:41, 6:0 }), hidden = Object.freeze({ 1:true, 5:true });
  const axis = new Axis(80, sizes, [hidden], 10);
  const cells = new Map([[1, 'hidden value'], [7, 'barrier']]);
  const before = { keys:axis.keys.slice(), cum:[...axis.cum], hidden:[...axis.hidden], cells:[...cells] };
  const result = cellTextOverflow(axis, 2, 'center', 600, c => cells.has(c));
  assert.deepEqual(result, { left:15, right:121 });
  assert.deepEqual(axis.keys, before.keys);
  assert.deepEqual([...axis.cum], before.cum);
  assert.deepEqual([...axis.hidden], before.hidden);
  assert.deepEqual([...cells], before.cells);
  assert.equal(axis.sizes, sizes);
});

// 작은 축의 선형 기하 기준으로 숨김 조합과 분수 너비를 교차 검증한다.
function geometricReference(axis, column, align, textWidth, values) {
  const reach = (textWidth - axis.size(column)) / (align === 'center' ? 2 : 1);
  const sum = dir => {
    let amount = 0;
    for (let c = column + dir; c >= 0 && c < axis.max; c += dir) {
      const width = axis.size(c);
      if (!width) continue;
      if (values.has(c)) break;
      amount += width;
      if (amount >= reach) break;
    }
    return amount;
  };
  return { left:align === 'right' || align === 'center' ? sum(-1) : 0, right:align === 'left' || align === 'center' ? sum(1) : 0 };
}

test('다양한 숨김·분수 너비·차단 조합이 표시 열의 선형 기하 기준과 같다', () => {
  let seed = 41791;
  const random = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) / 2 ** 32);
  for (let round = 0; round < 60; round++) {
    const sizes = {}, hidden = {}, values = new Set();
    for (let c = 0; c < 30; c++) {
      if (random() < .35) sizes[c] = Math.round(random() * 120) / 2;
      if (random() < .3) hidden[c] = true;
      if (random() < .3) values.add(c);
    }
    const column = 1 + Math.floor(random() * 28); sizes[column] = 80; delete hidden[column];
    const axis = new Axis(55.5, sizes, [hidden], 30), textWidth = 81 + random() * 600;
    for (const align of ['left', 'right', 'center']) {
      assert.deepEqual(probe(axis, column, align, textWidth, values).result, geometricReference(axis, column, align, textWidth, values), `round=${round} align=${align}`);
    }
  }
});
