import { test } from 'node:test';
import assert from 'node:assert/strict';
import { NET, netText, LOADING } from '../src/fx-web.js';

const context = { here: { si: 0 } };
const flush = () => new Promise((resolve) => setImmediate(resolve));
function setup(t) {
  const saved = { cache: NET.cache, fetcher: NET.fetcher, authorize: NET.authorize, onDone: NET.onDone, waiting: NET.waiting };
  NET.cache = new Map(); NET.fetcher = null; NET.authorize = null; NET.onDone = null; NET.waiting = 0;
  t.after(() => Object.assign(NET, saved));
}

test('네트워크 권한은 성공 캐시보다 먼저 검사하고 명시적 허용 후 다시 가져온다', async (t) => {
  setup(t);
  const url = 'wixel-doc:개인문서\u0001Sheet1\u0001A1';
  NET.cache.set(url, { state: 'ok', data: '개인 데이터', t: Date.now(), sheets: new Set([0]) });
  let calls = 0;
  NET.fetcher = async () => { calls++; return '허용한 데이터'; };
  NET.authorize = () => '다른 개인 문서의 참조는 허용하지 않았습니다.';
  assert.equal(netText(url, context).code, '#N/A');
  assert.equal(calls, 0);
  assert.equal(NET.cache.get(url).data, undefined);
  assert.match(NET.cache.get(url).message, /허용하지/);
  NET.authorize = () => null;
  assert.equal(netText(url, context), LOADING);
  await flush();
  assert.equal(netText(url, context), '허용한 데이터');
  assert.equal(calls, 1);
});

test('진행 중이던 개인 요청은 권한 변경 후 거부 캐시를 덮어쓰지 않는다', async (t) => {
  setup(t);
  const url = 'wixel-doc:개인문서\u0001Sheet1\u0001A1';
  let complete;
  NET.fetcher = () => new Promise((resolve) => { complete = resolve; });
  assert.equal(netText(url, context), LOADING);
  await flush();
  NET.authorize = () => '현재 문서는 개인 데이터에 접근할 수 없습니다.';
  assert.equal(netText(url, context).code, '#N/A');
  complete('늦게 도착한 비공개 데이터');
  await flush();
  assert.equal(NET.cache.get(url).state, 'error');
  assert.equal(NET.cache.get(url).data, undefined);
  assert.equal(netText(url, context).code, '#N/A');
  assert.equal(NET.waiting, 0);
  NET.authorize = () => { throw new Error('권한 검사 실패'); };
  assert.equal(netText(url, context).code, '#N/A');
  assert.match(NET.cache.get(url).message, /권한/);
});
