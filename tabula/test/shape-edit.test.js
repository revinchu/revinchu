import { test } from 'node:test';
import assert from 'node:assert/strict';
import { shapeFromPoints } from '../src/shape-path.js';
import { editableShapePath, shapePathHandles, shapeLocalPoint, moveShapePoint, deleteShapePoint, insertShapePoint, normalizeEditedShape } from '../src/shape-edit.js';
const close = (a, b) => assert.ok(Math.abs(a - b) < 1e-7, `${a} != ${b}`);
const world = (s, p) => { const a = (s.rot ?? 0) * Math.PI / 180, x = ((s.flip ? 1 - p[0] : p[0]) - 0.5) * Math.max(1, s.w), y = ((s.flipV ? 1 - p[1] : p[1]) - 0.5) * Math.max(1, s.h); return [s.x + Math.max(1, s.w) / 2 + x * Math.cos(a) - y * Math.sin(a), s.y + Math.max(1, s.h) / 2 + x * Math.sin(a) + y * Math.cos(a)]; };
const cubic = { paths: [{ fill: false, stroke: true, commands: [['M', 0, 0], ['C', 0.2, 0, 0.3, 1, 0.5, 0.5], ['C', 0.7, 0, 0.8, 1, 1, 1]] }] };
test('회전·가로세로 대칭 도형의 화면 좌표를 편집 좌표로 되돌린다', () => {
  for (const rot of [0, 32, 90, 270]) for (const flip of [false, true]) for (const flipV of [false, true]) {
    const shape = { x: 31, y: 85, w: 370, h: 65, rot, flip, flipV };
    for (const point of [[0, 0], [0.2, 0.9], [1.4, -0.2]]) { const back = shapeLocalPoint(shape, world(shape, point)); close(back[0], point[0]); close(back[1], point[1]); }
  }
});
test('Bezier 점 이동은 양쪽 연결 제어점을 함께 옮기고 원본을 보존한다', () => {
  const before = JSON.stringify(cubic), h = shapePathHandles({ path: cubic }).find(h => h.c === 1 && h.anchor);
  const changed = moveShapePoint(cubic, h, [0.6, 0.7]);
  assert.equal(JSON.stringify(cubic), before);
  close(changed.paths[0].commands[1][3], 0.4); close(changed.paths[0].commands[1][4], 1.2);
  close(changed.paths[0].commands[2][1], 0.8); close(changed.paths[0].commands[2][2], 0.2);
});
test('Shift 대칭·Ctrl 공선 길이 보존·Alt 독립 제어점은 가로세로 비율에도 정확하다', () => {
  const handle = shapePathHandles({ path: cubic }).find(h => h.c === 2 && h.slot === 1), point = [0.9, 0.8], anchor = [0.5, 0.5];
  const symmetric = moveShapePoint(cubic, handle, point, 'symmetric', 400, 50).paths[0].commands[1];
  close(symmetric[3], 0.1); close(symmetric[4], 0.2);
  const smooth = moveShapePoint(cubic, handle, point, 'smooth', 400, 50).paths[0].commands[1];
  close(Math.hypot((smooth[3] - anchor[0]) * 400, (smooth[4] - anchor[1]) * 50), Math.hypot(0.2 * 400, 0.5 * 50));
  const corner = moveShapePoint(cubic, handle, point, false).paths[0].commands[1]; assert.deepEqual(corner, cubic.paths[0].commands[1]);
});
test('곡선 점 추가는 de Casteljau 분할로 곡선을 보존하고 닫힘 구간도 나눈다', () => {
  const original = JSON.stringify(cubic), result = insertShapePoint(cubic, [0.27, 0.43], 300, 90);
  assert.equal(result.path.paths[0].commands.length, 4); assert.equal(JSON.stringify(cubic), original);
  const left = result.path.paths[0].commands[result.handle.c], right = result.path.paths[0].commands[result.handle.c + 1];
  assert.equal(left[0], 'C'); assert.equal(right[0], 'C');
  const a = [left[5] - left[3], left[6] - left[4]], b = [right[1] - left[5], right[2] - left[6]];
  close(a[0] * b[1] - a[1] * b[0], 0);
  const box = editableShapePath({ kind: 'rect' }), added = insertShapePoint(box, [0, 0.5]);
  assert.deepEqual(added.path.paths[0].commands.at(-1), ['Z']); assert.equal(added.path.paths[0].commands.length, 6);
});
test('최소 점 수·시작점 삭제는 유효 경로를 유지한다', () => {
  const path = editableShapePath({ kind: 'rect' }), handle = shapePathHandles({ path }).find(h => h.c === 0);
  const triangle = deleteShapePoint(path, handle); assert.equal(triangle.paths[0].commands[0][0], 'M');
  assert.equal(deleteShapePoint(triangle, shapePathHandles({ path: triangle })[0]), null);
  assert.equal(deleteShapePoint(editableShapePath({ kind: 'line' }), { p: 0, c: 0, slot: 1, anchor: true }), null);
});
test('범위 밖으로 옮긴 점은 경계 상자 재조정 뒤에도 회전·대칭 위치가 유지된다', () => {
  const { path } = shapeFromPoints('curve', [[0, 0], [50, 80], [100, 0]]);
  for (const rot of [0, 36, 90]) for (const flip of [false, true]) for (const flipV of [false, true]) {
    const shape = { x: 150, y: 120, w: 300, h: 60, path, rot, flip, flipV };
    const h = shapePathHandles(shape).find(h => h.c === 1 && h.anchor), nextPath = moveShapePoint(path, h, [1.4, -0.5]);
    const patch = normalizeEditedShape(shape, nextPath), normalized = { ...shape, ...patch };
    const before = shapePathHandles({ path: nextPath }), after = shapePathHandles(normalized);
    assert.ok(patch.w > 0 && patch.h > 0);
    for (let i = 0; i < before.length; i++) { const a = world(shape, [before[i].x, before[i].y]), b = world(normalized, [after[i].x, after[i].y]); close(a[0], b[0]); close(a[1], b[1]); }
  }
});

const closedCubic = { paths: [{ fill: true, stroke: true, commands: [
  ['M', 0, 0], ['C', .3, -.2, .7, -.2, 1, 0], ['C', 1.2, .3, 1.2, .7, 1, 1],
  ['C', .7, 1.2, .3, 1.2, 0, 1], ['C', -.2, .7, -.2, .3, 0, 0], ['Z'],
] }] };

test('닫힌 cubic 시작·끝은 한 꼭짓점이며 양쪽 제어점과 함께 이동한다', () => {
  const original = JSON.stringify(closedCubic), handles = shapePathHandles({ path: closedCubic });
  assert.equal(handles.filter(h => h.anchor).length, 4);
  assert.equal(handles.filter(h => h.c === 4 && !h.anchor).length, 2);
  for (const handle of [handles.find(h => h.c === 0), { p: 0, c: 4, slot: 5, anchor: true }]) {
    const p = moveShapePoint(closedCubic, handle, [.2, .1]).paths[0].commands;
    assert.deepEqual(p[0], ['M', .2, .1]); assert.deepEqual(p[4].slice(-2), [.2, .1]);
    close(p[1][1], .5); close(p[1][2], -.1); close(p[4][3], 0); close(p[4][4], .4);
    assert.deepEqual(p[2], closedCubic.paths[0].commands[2]);
  }
  assert.equal(JSON.stringify(closedCubic), original);
});

test('닫힌 cubic seam의 Shift/Ctrl 접선은 양방향·비정방형 좌표에서 연결된다', () => {
  const handles = shapePathHandles({ path: closedCubic }), width = 400, height = 70;
  for (const [c, slot, otherC, otherSlot, point] of [[1, 1, 4, 3, [.4, .3]], [4, 3, 1, 1, [-.3, .4]]]) {
    const handle = handles.find(h => h.c === c && h.slot === slot);
    const symmetric = moveShapePoint(closedCubic, handle, point, 'symmetric', width, height).paths[0].commands;
    close(symmetric[otherC][otherSlot], -point[0]); close(symmetric[otherC][otherSlot + 1], -point[1]);
    const aligned = moveShapePoint(closedCubic, handle, point, 'smooth', width, height).paths[0].commands;
    const old = closedCubic.paths[0].commands[otherC];
    close(Math.hypot(aligned[otherC][otherSlot] * width, aligned[otherC][otherSlot + 1] * height), Math.hypot(old[otherSlot] * width, old[otherSlot + 1] * height));
    close(aligned[otherC][otherSlot] * point[1] - aligned[otherC][otherSlot + 1] * point[0], 0);
    assert.ok(aligned[otherC][otherSlot] * point[0] + aligned[otherC][otherSlot + 1] * point[1] < 0);
    const corner = moveShapePoint(closedCubic, handle, point, false).paths[0].commands;
    assert.deepEqual(corner[otherC], closedCubic.paths[0].commands[otherC]);
  }
});

test('닫힌 quadratic seam도 시작점·단일 제어점·반대편 cubic 접선을 유지한다', () => {
  const path = structuredClone(closedCubic); path.paths[0].commands[4] = ['Q', -.3, .5, 0, 0];
  const handles = shapePathHandles({ path }), endControl = handles.find(h => h.c === 4 && !h.anchor);
  assert.deepEqual(endControl.origin, [0, 0]); assert.equal(handles.filter(h => h.anchor).length, 4);
  const moved = moveShapePoint(path, handles.find(h => h.c === 0), [.2, -.1]).paths[0].commands;
  assert.deepEqual(moved[0], ['M', .2, -.1]); assert.deepEqual(moved[4].slice(-2), [.2, -.1]);
  close(moved[4][1], -.1); close(moved[4][2], .4); close(moved[1][1], .5); close(moved[1][2], -.3);
  const smooth = moveShapePoint(path, endControl, [-.4, .2], 'symmetric', 300, 50).paths[0].commands;
  close(smooth[1][1], .4); close(smooth[1][2], -.2);
  const other = moveShapePoint(path, handles.find(h => h.c === 1 && h.slot === 1), [.4, -.2], 'symmetric').paths[0].commands;
  close(other[4][1], -.4); close(other[4][2], .2);
});

test('닫힌 seam 삭제는 중복 끝점을 세지 않고 새 시작점에 재연결한다', () => {
  for (const quadratic of [false, true]) {
    const path = structuredClone(closedCubic); if (quadratic) path.paths[0].commands[4] = ['Q', -.3, .5, 0, 0];
    const before = JSON.stringify(path), first = shapePathHandles({ path }).find(h => h.c === 0);
    const triangle = deleteShapePoint(path, first), commands = triangle.paths[0].commands;
    assert.deepEqual(commands[0], ['M', 1, 0]); assert.deepEqual(commands.at(-2).slice(-2), [1, 0]); assert.deepEqual(commands.at(-1), ['Z']);
    assert.equal(shapePathHandles({ path: triangle }).filter(h => h.anchor).length, 3);
    for (const anchor of shapePathHandles({ path: triangle }).filter(h => h.anchor)) assert.equal(deleteShapePoint(triangle, anchor), null);
    const legacyLast = { p: 0, c: 4, slot: quadratic ? 3 : 5, anchor: true };
    assert.deepEqual(deleteShapePoint(path, legacyLast), triangle); assert.equal(JSON.stringify(path), before);
  }
  const path = { paths: [{ commands: [...closedCubic.paths[0].commands, ['M', 3, 3], ['L', 4, 4]] }] };
  assert.equal(deleteShapePoint(path, { p: 0, c: 6, slot: 1, anchor: true }), null, '별도 열린 하위 경로는 자체 최소2점을 지킨다');
});

test('이미 닫힌 곡선의 점 추가는 길이0인 Z 대신 실제 마지막 곡선을 분할한다', () => {
  const result = insertShapePoint(closedCubic, [-.01, .01], 300, 100);
  assert.equal(result.path.paths[0].commands[result.handle.c][0], 'C');
  const commands = result.path.paths[0].commands;
  assert.deepEqual(commands.at(-1), ['Z']); assert.deepEqual(commands.at(-2).slice(-2), [0, 0]);
  assert.equal(shapePathHandles({ path: result.path }).filter(h => h.anchor).length, 5);
});
