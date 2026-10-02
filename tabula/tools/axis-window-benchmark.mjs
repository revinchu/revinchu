import assert from 'node:assert/strict';
import { performance } from 'node:perf_hooks';
import { Axis } from '../src/axis.js';
import { visibleAxisIndices, nextVisibleAxisIndex } from '../src/axis-window.js';

// 합성 입력만 사용. Axis 생성 시간을 제외하고 기존 전체 범위 순회와 같은 결과인지 비교한다.
// 절대 시간은 장비마다 달라서 성공/실패 기준은 동일 결과와 크기 조회 수이다.
function scan(axis, first, last) { const out = []; for (let i = first; i <= last; i++) if (axis.size(i)) out.push(i); return out; }
function oldNext(axis, i, dir) { let j = i; while (j >= 0 && j < axis.max && axis.size(j) === 0) j += dir; return j < 0 || j >= axis.max ? i : j; }
function sample(fn, n = 15) { for (let i = 0; i < 4; i++) fn(); const times = []; for (let i = 0; i < n; i++) { const t = performance.now(); fn(); times.push(performance.now() - t); } times.sort((a, b) => a - b); return +times[Math.floor(n / 2)].toFixed(4); }
function fixture(n, every) { const data = new Uint8Array(n); data.fill(1); let visible = 0; for (let i = 0; i < n; i += every) { data[i] = 0; visible++; } return new Axis(23, {}, [{ __bits: data, start: 0, count: n - visible }], n + 50); }
const rows = [];
for (const [name, axis, last] of [['dense-visible-60', new Axis(23, {}, [], 1000), 59], ['filtered-58690', fixture(58690, 500), 58689], ['filtered-1million', fixture(1000000, 25000), 999999]]) {
  const expected = scan(axis, 0, last); assert.deepEqual(visibleAxisIndices(axis, 0, last), expected);
  const before = sample(() => scan(axis, 0, last)), after = sample(() => visibleAxisIndices(axis, 0, last));
  rows.push({ name, visible: expected.length, beforeMs: before, afterMs: after, speedup: +(before / after).toFixed(1) });
}
const axis = fixture(1000000, 1000000);
assert.equal(nextVisibleAxisIndex(axis, 1, 1), oldNext(axis, 1, 1));
rows.push({ name: 'next-visible-million-hidden', beforeMs: sample(() => oldNext(axis, 1, 1)), afterMs: sample(() => nextVisibleAxisIndex(axis, 1, 1)) });
console.log(JSON.stringify({ kind: 'synthetic Axis enumeration only; constructor excluded', rows }, null, 2));
